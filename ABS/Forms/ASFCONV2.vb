Public Class ASFCONV2
    Public result As String = ""
    Public CONV_NO As String
    Public passedInLog As String = ""
    Public rowTATCONV1 As DataRow
    Public rowTATCONV1_PREV As DataRow
    Dim context_TABLE_NAME As String
    Dim context_TABLE_KEY As String

    Public frmASFBASE0 As ASFBASE0

    ' optional behaviors depending upon TABLE_NAME
    Public subject_is_mandatory As Boolean = True
    Public subject_is_hidden As Boolean = False
    Public followup_is_defaulted_to_true As Boolean = False
    Public followup_is_mandatory As Boolean = False

    Public Sub New(ByVal FF As ASFBASE0, ByVal TABLE_NAME As String, ByVal TABLE_KEY As String, Optional ByVal logText As String = "")
        frmASFBASE0 = FF
        context_TABLE_NAME = TABLE_NAME
        context_TABLE_KEY = TABLE_KEY
        passedInLog = logText
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "TATCONV1", "*")
            .Tables("TATCONV1").Columns.Add("CONV_ATTACHMENTS", GetType(System.Int64))
        End With

        If EntryMode = "N" Or EntryMode = "F" Then
            rowTATCONV1 = dst.Tables("TATCONV1").NewRow
            CONV_NO = ASCMAIN1.Next_Control_No("TATCONV1.CONV_NO")
            rowTATCONV1.Item("CONV_NO") = CONV_NO
            rowTATCONV1.Item("CONV_DATE") = DATETIME_STAMP.Date
            rowTATCONV1.Item("CONV_NOTES") = passedInLog
            rowTATCONV1("INIT_OPER") = ASCMAIN1.USER_ID
            rowTATCONV1("INIT_DATE") = DATETIME_STAMP ' Now + ASCMAIN1.NowTSD
            rowTATCONV1("TABLE_NAME") = context_TABLE_NAME
            rowTATCONV1("TABLE_KEY") = context_TABLE_KEY
            rowTATCONV1("CONV_STATUS") = "0"
            dst.Tables("TATCONV1").Rows.Add(rowTATCONV1)

            If EntryMode = "F" Then
                rowTATCONV1.Item("CONV_SUBJECT") = rowTATCONV1_PREV.Item("CONV_SUBJECT")
                rowTATCONV1.Item("CONV_NO_PREV") = rowTATCONV1_PREV.Item("CONV_NO")
                Me.Text &= " - (Follow-Up)"
            End If

            Show_Log_Recorded_By()
            chkFollowUp.Checked = False
            Set_Follow_Up_Controls()

        ElseIf EntryMode = "V" Or EntryMode = "E" Then
            chkFollowUp.Appearance.ForeColorDisabled = Color.Black
            Dim row As DataRow = dst.Tables("TATCONV1").NewRow
            For i As Integer = 0 To dst.Tables("TATCONV1").Columns.Count - 1
                row.Item(dst.Tables("TATCONV1").Columns(i).ColumnName) = rowTATCONV1.Item(dst.Tables("TATCONV1").Columns(i).ColumnName)
            Next
            dst.Tables("TATCONV1").Rows.Add(row)
            'dst.Tables("TATCONV1").Rows.Add(rowTATCONV1.ItemArray)

            If EntryMode = "V" Then
                Set_Read_Only(grpTATCONV1, True)
            Else
                Set_Read_Only(chkFollowUp, rowTATCONV1.Item("CONV_STATUS") & "" = "2")
                Set_Read_Only(txtCONV_FOLLOWUP_BY, rowTATCONV1.Item("CONV_STATUS") & "" = "2")
                Set_Read_Only(Absx1.dteFor("CONV_FOLLOWUP_DATE"), rowTATCONV1.Item("CONV_STATUS") & "" = "2")
            End If
            Show_Log_Recorded_By()

            If rowTATCONV1.Item("CONV_FOLLOWUP_BY") & "" <> "" Then
                chkFollowUp.Checked = True
            Else
            End If

            Set_Follow_Up_Controls()
            If EntryMode = "V" Then
                cmdUpdate.Visible = False
                cmdCancel.Text = "Done"
            End If
        End If

        splLog.Panel2Collapsed = Not (chkFollowUp.Checked And EntryMode = "V") And Not (EntryMode = "F")

        Me.Text &= " - " & rowTATCONV1.Item("CONV_NO")
        Set_Defaults()

        If EntryMode = "F" Or EntryMode = "N" Then
            If followup_is_defaulted_to_true Then
                chkFollowUp.Checked = True
            End If
        End If

        Absx1.txtFor("CONV_SUBJECT").Visible = Not subject_is_hidden
        lblCONV_SUBJECT.Visible = Not subject_is_hidden

        Set_Read_Only(Absx1.dteFor("CONV_DATE"), True)


    End Sub

    Sub Show_Log_Recorded_By(Optional ByVal show_user As Boolean = False)
        grpLog.Text = "Notes by " & rowTATCONV1("INIT_OPER") & " " & Format(rowTATCONV1.Item("INIT_DATE"), "MM/dd/yyyy HH:mm") & IIf(EntryMode = "N" Or EntryMode = "F", " (Now)", "")

        If EntryMode = "F" Then
            grpFollowup.Text = "This entry is a Follow-Up to a Log Entered by " & rowTATCONV1_PREV("INIT_OPER") & " " & Format(rowTATCONV1_PREV.Item("INIT_DATE"), "MM/dd/yyyy HH:mm")
            Absx1.txtFor("CONV_FOLLOWUP_NOTES").Text = rowTATCONV1_PREV.Item("CONV_NOTES") & ""
        ElseIf EntryMode = "V" Or EntryMode = "E" Then
            If rowTATCONV1("CONV_FOLLOWUP_BY") & "" <> "" Then
                If rowTATCONV1("CONV_FOLLOWUP_CONV_NO") & "" <> "" Then
                    grpFollowup.Text = "Latest Follow-Up by " & rowTATCONV1("LAST_OPER") & " " & Format(rowTATCONV1.Item("LAST_DATE"), "MM/dd/yyyy HH:mm")
                    Dim row As DataRow = frmASFBASE0.dst.Tables("TATCONV1").Rows.Find(rowTATCONV1("CONV_FOLLOWUP_CONV_NO"))
                    If row Is Nothing Then
                        row = LookUp("TATCONV1", rowTATCONV1("CONV_FOLLOWUP_CONV_NO"))
                    End If
                    Absx1.txtFor("CONV_FOLLOWUP_NOTES").Text = row.Item("CONV_NOTES")
                Else
                    If rowTATCONV1("LAST_OPER") & "" <> "" Then
                        grpFollowup.Text = "Latest Follow-Up by " & rowTATCONV1("LAST_OPER") & " " & Format(rowTATCONV1.Item("LAST_DATE"), "MM/dd/yyyy HH:mm")
                    Else
                        grpFollowup.Text = "No Follow-Up (yet)"
                    End If
                End If
            End If
        End If
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CONV_FOLLOWUP_BY"
        End Select
    End Sub

    Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)

        Select Case Absx1.GetABSColumnName(sender)
            Case ""
        End Select

    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        If Update_Log(True) = "" Then
            Me.Close()
        End If

    End Sub

    Function Update_Log(ByVal display_errors As Boolean) As String

        Dim EMsg As String = ""

        Select Case EntryMode
            Case "N", "E", "F"
                If subject_is_mandatory Then
                    If Absx1.txtFor("CONV_SUBJECT").Text = "" Then
                        EMsg &= vbCr & "Subject/Contact is Required"
                    End If
                End If

                If Absx1.txtFor("CONV_NOTES").Text = "" Then
                    EMsg &= vbCr & "Notes may not be blank"
                End If

                If chkFollowUp.Checked Then
                    Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", Absx1.txtFor("CONV_FOLLOWUP_BY").Text)
                    If rowASTUSER1 Is Nothing Then
                        EMsg &= vbCr & "Invalid User ID to Follow Up with"
                    End If

                    If Absx1.dteFor("CONV_FOLLOWUP_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Invalid Follow-Up Date"
                    End If
                Else
                    If followup_is_mandatory Then
                        If EntryMode = "E" And _
                        Not chkFollowUp.Checked Then '  rowTATCONV1.Item("CONV_STATUS") & "" <> "1" Then
                        Else
                            EMsg &= vbCr & "Follow-Up is Mandatory"
                        End If
                    End If
                End If

        End Select

        If Not display_errors Then
            EMsg = ""
        End If

        If EMsg <> "" Then
            If display_errors Then
                MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update")
            End If
            Return EMsg
            Exit Function
        End If

        Me.Cursor = Cursors.WaitCursor

        result = "U"

        If Not chkFollowUp.Checked Then
            Absx1.txtFor("CONV_FOLLOWUP_BY").Text = ""
            Absx1.dteFor("CONV_FOLLOWUP_DATE").Value = DBNull.Value
        End If

        If Me.BindingContext.Contains(dst.Tables(TABLE_NAME)) Then
            ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
            Dim X As CurrencyManager = Me.BindingContext(dst.Tables(TABLE_NAME))
            X.EndCurrentEdit()
        End If


        If EntryMode = "F" Then
            rowTATCONV1_PREV("LAST_OPER") = rowTATCONV1.Item("INIT_OPER")
            rowTATCONV1_PREV("LAST_DATE") = rowTATCONV1.Item("INIT_DATE")
            rowTATCONV1_PREV("CONV_STATUS") = "2"
            rowTATCONV1_PREV("CONV_FOLLOWUP_CONV_NO") = rowTATCONV1.Item("CONV_NO")
        ElseIf EntryMode = "E" Then
            rowTATCONV1.ItemArray = dst.Tables("TATCONV1").Rows(0).ItemArray
            'rowTATCONV1.ItemArray = frmASFBASE0.dst.Tables("TATCONV1").Rows(0).ItemArray
        End If
        If chkFollowUp.Checked Then
            rowTATCONV1("CONV_STATUS") = "1"
        Else
            rowTATCONV1("CONV_STATUS") = "0"
        End If

        Me.Cursor = Cursors.Default
        Return EMsg

    End Function

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        result = "C"
        Me.Close()
    End Sub

    Private Sub chkFollowUp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFollowUp.CheckedChanged
        Set_Follow_Up_Controls()
    End Sub

    Sub Set_Follow_Up_Controls()
        lblCONV_FOLLOWUP_DATE.Visible = chkFollowUp.Checked
        Absx1.txtFor("CONV_FOLLOWUP_BY").Visible = chkFollowUp.Checked
        Absx1.dteFor("CONV_FOLLOWUP_DATE").Visible = chkFollowUp.Checked

        If chkFollowUp.Checked = False Then
            Absx1.txtFor("CONV_FOLLOWUP_BY").Text = ""
            Absx1.dteFor("CONV_FOLLOWUP_DATE").Value = DBNull.Value
            chkFollowUp.Text = "Needs Follow-Up"
        Else
            Absx1.txtFor("CONV_FOLLOWUP_BY").Text = ASCMAIN1.USER_ID
            chkFollowUp.Text = "Needs Follow-Up by"
        End If

        'Absx1.numFor("CONV_PROMISE_AMT").Value = DBNull.Value
        'Absx1.dteFor("CONV_PROMISE_BY").Value = DBNull.Value
    End Sub

    Private Sub cmdAttach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAttach.Click
        Show_Attachments_TATCONV1()
    End Sub

    Sub Show_Attachments_TATCONV1()
        Dim E As Dropped_On_Entity = Dropped_On_Context()

        If E.TABLE_NAME <> "" Then
            Dim F As New ASFATTA1
            F.ENTITY = E
            F.ShowDialog()
            F.Dispose()
        End If
    End Sub

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        'If ScreenMode Then
        E.TABLE_NAME = "TATCONV1"
        E.COLUMN_NAME = "CONV_NO"
        E.CODE_VALUE = rowTATCONV1.Item("CONV_NO")
        E.DESC_VALUE = grpLog.Text & " (" & Absx1.txtFor("CONV_SUBJECT").Text & ")"
        E.ATTACHMENT_NOTES = ""
        'End If

        Return E
    End Function

    Public Function Add_Log( _
    ByVal CONV_SUBJECT As String, _
    ByVal CONV_NOTES As String, _
    Optional ByVal CONV_FOLLOWUP_BY As String = "", _
    Optional ByVal CONV_FOLLOWUP_DATE As Date = Nothing) As String

        Me.Width = 100
        Me.Height = 100

        Me.Show()

        rowTATCONV1.Item("CONV_SUBJECT") = CONV_SUBJECT
        rowTATCONV1.Item("CONV_NOTES") = CONV_NOTES

        If CONV_FOLLOWUP_BY <> "" Then
            rowTATCONV1.Item("CONV_FOLLOWUP_BY") = CONV_FOLLOWUP_BY
            rowTATCONV1.Item("CONV_FOLLOWUP_DATE") = CONV_FOLLOWUP_DATE
            chkFollowUp.Checked = False
            chkFollowUp.Checked = True
            Set_Follow_Up_Controls()
        End If

        Dim X As CurrencyManager = Me.BindingContext(dst.Tables("TATCONV1"))
        X.EndCurrentEdit()

        Return Update_Log(False)

        MsgBox("", MsgBoxStyle.OkOnly, "")
    End Function

    Private Sub txtCONV_FOLLOWUP_BY_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCONV_FOLLOWUP_BY.DoubleClick
        txtCONV_FOLLOWUP_BY.Text = ASCMAIN1.USER_ID
    End Sub

    Sub Set_Defaults()
        ' need to place these into log_entity
        ' actually, these should go into a parameters table
        Select Case context_TABLE_NAME
            Case "PMTPROP1"
                subject_is_mandatory = False
                subject_is_hidden = True
                followup_is_defaulted_to_true = True
        End Select
    End Sub
End Class