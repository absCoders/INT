Public Class APTVEND1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select APTVEND9.*, GLTACCT1.ACCT_DESC from APTVEND9,GLTACCT1 where GLTACCT1.ACCT_CODE = APTVEND9.ACCT_CODE and APTVEND9.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTVEND9", "**", 0, True, "V", 5)
            ASCMAIN1.sql = "Select APTVENDB.* from APTVENDB where APTVENDB.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "APTVENDB", "**", 0, True, "V", 3)

            If ASCMAIN1.CLIENT = "XXX" Then
                ASCMAIN1.sql = "Select APTVENDA.*, APTATTR1.VEND_ATTR_DESC from APTVENDA,APTATTR1 where APTATTR1.VEND_ATTR_CODE = APTVENDA.VEND_ATTR_CODE and APTVENDA.VEND_CODE = :PARM1"
                Create_TDA(.Tables.Add, "APTVENDA", "**", 0, True, "V", 2)
            End If

        End With

        grdAPTVEND9.DataSource = dst.Tables("APTVEND9")
        grdAPTVENDB.DataSource = dst.Tables("APTVENDB")

        Get_PARM("GLTPARM1")
        Get_PARM("APTPARM1")

        Set_SEGS(grdAPTVEND9, "APTVEND9")
        Create_Summary(grdAPTVEND9, "DIST_AMT")

        If ASCMAIN1.DBS_SERVER = "EXP" Or ASCMAIN1.DBS_COMPANY = "EXP" Then
            UltraTabControl1.Tabs("Purchasing Information").Visible = False
        End If

        If ASCMAIN1.CLIENT = "AHA" Then
            grdAPTVENDB.Visible = True
            ASCMAIN1.Add_Value_List(grdAPTVENDB, "ACH_ACCT_STATUS", Nothing, New String() {":", "A:Active", "I:Inactive"})
        Else
            grdAPTVENDB.Visible = False
        End If

        If ASCMAIN1.CLIENT = "XXX" Then
            grdAPTVENDA.Visible = True
            grdAPTVENDA.DataSource = dst.Tables("APTVENDA")
        End If

        grpBankingInfo.Visible = (ASCMAIN1.USER_SECURITY_CODEs.Split(",").Contains("P3"))

    End Sub
    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "APTVEND1"
            E.COLUMN_NAME = "VEND_CODE"
            E.CODE_VALUE = Absx1.txtFor("VEND_CODE").Text
            E.DESC_VALUE = "Vendor"
            E.ATTACHMENT_NOTES = ""
            'E.RESTRICTIONS = "D"
            'E.READ_ONLY = True
        End If

        Return E
    End Function

    'Public Overrides Function Audit_Context() As Audit_Entity

    '    Dim E As New Audit_Entity
    '    If ScreenMode Then
    '        E.TABLE_NAME = "APTVEND1"
    '        E.KEY_VALUE = Absx1.txtFor("VEND_CODE").Text
    '        E.KEY_DESC = "Vendor"
    '    End If
    '    Return E
    'End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "APTVEND1"
        E.TABLE_KEY_CAPTION = "Vendor"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("VEND_CODE").Text
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text & " " & Absx1.txtFor("VEND_NAME").Text
            E.TABLE_KEY_locked = ScreenMode
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()

        Dim sql As String = "Delete from APTVEND9 where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
        Update_Record_TDA("APTVEND9", sql)

        If ASCMAIN1.CLIENT = "AHA" Then
            Update_Record_TDA("APTVENDB")
        End If

        If ASCMAIN1.CLIENT = "XXX" Then
            sql = "Delete from APTVENDA where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
            Update_Record_TDA("APTVENDA", sql)
        End If

    End Sub

    Overrides Sub Show_Record_Special()
        Dim txtctl As UltraWinEditors.UltraTextEditor = Absx1.txtFor("VEND_CODE")
        Clear_Record_Special()
        Load_Report_Form(txtctl.Text)

        
        If EntryMode = "New" Then
            rowASFBASE1.Item("TERM_CODE") = ROWs("APTPARM1").Item("AP_PARM_TERM_CODE")
            'rowASFBASE1.Item("POST_CODE") = ROWs("APTPARM1").Item("AP_PARM_POST_CODE")
            rowASFBASE1.Item("VEND_STATUS") = "A"
            rowASFBASE1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        End If

    End Sub

    Sub Load_Report_Form(ByVal VEND_CODE As String)

        Fill_Records("APTVENDB", VEND_CODE)

        Fill_Records("APTVEND9", VEND_CODE)
        For Each r As DataRow In dst.Tables("APTVEND9").Rows
            'If r.Item("COLUMN_CAPTION") & "" = "" Then
            '    Dim rowASTDSQLK As DataRow = dst.Tables("ASTDSQLK").Rows.Find(r.Item("COLUMN_NAME"))
            '    If Not rowASTDSQLK Is Nothing Then
            '        r.Item("COLUMN_CAPTION") = rowASTDSQLK.Item("COLUMN_CAPTION")
            '    End If
            'End If
        Next

        If ASCMAIN1.CLIENT = "XXX" Then
            Fill_Records("APTVENDA", VEND_CODE)
        End If


        dst.EnforceConstraints = False

        dst.EnforceConstraints = True

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("APTVEND9").Rows.Clear()
            dst.Tables("APTVENDB").Rows.Clear()
            If ASCMAIN1.CLIENT = "XXX" Then
                dst.Tables("APTVENDA").Rows.Clear()
            End If
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdAPTVEND9.Enabled = tf
        grdAPTVENDB.Enabled = tf

        If ASCMAIN1.CLIENT = "XXX" Then
            grdAPTVENDA.Enabled = tf
        End If

        If tf And Not ASCMAIN1.USER_SECURITY_CODEs.Contains("AP") Then
            Set_Read_Only(UltraTabControl1.Tabs("Name && Address").TabPage, True)
            Set_Read_Only(UltraTabControl1.Tabs("Codes && Other Info").TabPage, True)
            Set_Read_Only(UltraTabControl1.Tabs("Payment Information").TabPage, True)
            Set_Read_Only_for_ctl(Absx1.optFor("VEND_STATUS"), True)
            Set_Read_Only_for_ctl(Absx1.txtFor("VEND_NAME"), True)
            UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Purchasing Information")
        End If

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"
                If Absx1.chkFor("PO_EMAIL_CONFIRM").Checked And Absx1.optFor("PO_XMIT_VIA").Value = "M" Then
                    EMsg &= EMsg & "email confirmation not necessary when Transmitting PO's via email"
                End If
                'Stop
                'If grdAPTVEND9.ActiveRow.IsAddRow Then
                '    EMsg &= vbCr & "Data Remaining in Addrow of GL Distribution Template"
                'End If

                ' Added by edz on 01/23/2008 as per Maria
                Dim VEND_TAX_ID As String = Absx1.txtFor("VEND_TAX_ID").Text.Trim
                Dim VEND_1099_BOX As Int32 = Val(Absx1.numFor("VEND_1099_BOX").Value & "")
                Dim VEND_TAX_ID_TYPE As String = Absx1.optFor("VEND_TAX_ID_TYPE").Value

                Select Case VEND_TAX_ID.Length
                    Case 0
                        ' NOTHING
                        If VEND_1099_BOX > 0 Then
                            EMsg &= EMsg & "Tax ID in 1099 Reporting section must be 9 numeric values when providing a value in the 1099 Box."
                        End If

                    Case 9
                        Dim temp As Long = 0
                        Long.TryParse(VEND_TAX_ID, temp)
                        If temp = 0 Then
                            EMsg &= EMsg & "Tax ID in 1099 Reporting section must be 9 numeric values."
                        Else
                            If (VEND_1099_BOX < 1 Or VEND_1099_BOX > 14 Or VEND_1099_BOX = 11 Or VEND_1099_BOX = 12) Then
                                EMsg &= EMsg & "Box values in the 1099 Reporting section must be 1 - 10, 13, 14."
                            ElseIf VEND_TAX_ID_TYPE Is Nothing OrElse VEND_TAX_ID_TYPE.Length = 0 Then
                                EMsg &= EMsg & "When providing 1099 Reporting information you must select the type."
                            End If
                        End If

                    Case Else
                        EMsg &= EMsg & "Tax ID in 1099 Reporting section must be 9 numeric values."

                End Select

                If ASCMAIN1.CLIENT = "AHA" Then
                    Dim rows() As DataRow = dst.Tables("APTVENDB").Select("ACH_PREFERRED = '1'")
                    If ROWs.Length > 1 Then
                        EMsg &= EMsg & "You may have only 1 Preferred ACH Account"
                    ElseIf rows.Length = 1 Then
                        If rows(0).Item("ACH_ACCT_STATUS") & "" <> "A" Then
                            EMsg &= EMsg & "Preferred ACH Account must be Active"
                        End If
                    Else
                        If Absx1.txtFor("VEND_PYMT_METHOD").Text = "ACH" Then
                            EMsg &= EMsg & "Vendor's default Payment Methid is ACH and there is no preferred ACH Account defined"
                        End If
                    End If
                End If

                Dim VEND_BANK_ROUTING_NO As String = Absx1.txtFor("VEND_BANK_ROUTING_NO").Text
                Dim VEND_BANK_SWIFT_NO As String = Absx1.txtFor("VEND_BANK_SWIFT_NO").Text
                Dim VEND_BANK_ACCT_ID As String = Absx1.txtFor("VEND_BANK_ACCT_ID").Text
                Dim VEND_BANK_COUNTRY As String = Absx1.txtFor("VEND_BANK_COUNTRY").Text
                Dim VEND_BANK_ACCT_CLASS As String = Absx1.optFor("VEND_BANK_ACCT_CLASS").Value & ""
                Dim VEND_BANK_ACCT_TYPE As String = Absx1.optFor("VEND_BANK_ACCT_TYPE").Value & ""

                If VEND_BANK_ROUTING_NO <> "" Or VEND_BANK_ACCT_ID <> "" Or VEND_BANK_SWIFT_NO <> "" Then
                    If VEND_BANK_ACCT_ID = "" Then EMsg &= vbCr & "Bank Account No is Mandatory"
                    If VEND_BANK_COUNTRY = "" Then
                        EMsg &= vbCr & "Bank Country is Mandatory"
                    Else
                        Dim rowTATCNTRY As DataRow = LookUp("TATCNTRY", VEND_BANK_COUNTRY)
                        If rowTATCNTRY Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Bank Country"
                        Else
                            If VEND_BANK_COUNTRY = "USA" Then
                                If VEND_BANK_ROUTING_NO = "" Then
                                    EMsg &= vbCr & "You must have a Routing No for a US Bank"
                                Else
                                    If Format(Val(VEND_BANK_ROUTING_NO), "000000000") <> VEND_BANK_ROUTING_NO Then
                                        EMsg &= vbCr & "Bank Routing No should be 9 digits all numeric"
                                    End If
                                End If
                                If VEND_BANK_SWIFT_NO <> "" Then EMsg &= vbCr & "You cannot have a Swift No for a US Bank"
                            Else
                                If VEND_BANK_ROUTING_NO <> "" Then EMsg &= vbCr & "You cannot have a Routing No for a non-US Bank"
                                If VEND_BANK_SWIFT_NO = "" Then
                                    EMsg &= vbCr & "You must have a Swift No for a non-US Bank"
                                Else
                                    'https://stackoverflow.com/questions/3028150/what-is-proper-regex-expression-for-swift-codes
                                    ' Dim rx As String = "[A-Z]{6,6}[A-Z2-9][A-NP-Z0-9]([A-Z0-9]{3,3}){0,1}" from JPMC spec
                                    Dim rx As String = "^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$"
                                    Dim r As New System.Text.RegularExpressions.Regex(rx)
                                    ' RBOSGB2L
                                    If r.IsMatch(VEND_BANK_SWIFT_NO) Then
                                    Else
                                        EMsg &= vbCr & $"{VEND_BANK_SWIFT_NO} has Special Characters which are not allowed"
                                        EMsg &= vbCr & "A swift code should be 8 or 11 letters or digits where the first six must be letters."
                                    End If
                                End If

                            End If
                        End If
                    End If

                    If VEND_BANK_ACCT_CLASS = "" Then EMsg &= vbCr & "Bank Account Class is Mandatory"
                    If VEND_BANK_ACCT_TYPE = "" Then EMsg &= vbCr & "Bank Account Type is Mandatory"
                End If

                'If Not String.IsNullOrWhiteSpace(Absx1.txtFor("COST_CLASS_CODE").Text) Then
                '    Dim rowICTCCLS1 = LookUp("ICTCCLS1", Absx1.txtFor("COST_CLASS_CODE").Text)
                '    If rowICTCCLS1 Is Nothing Then
                '        EMsg &= vbCr & "Invalid Value specified for Cost Class Code"
                '    End If
                'End If
                'If Not String.IsNullOrWhiteSpace(Absx1.txtFor("COST_LIST_CODE").Text) Then
                '    Dim rowICTCLST1 = LookUp("ICTCLST1", Absx1.txtFor("COST_LIST_CODE").Text)
                '    If rowICTCLST1 Is Nothing Then
                '        EMsg &= vbCr & "Invalid Value specified for Cost List Code"
                '    End If
                'End If
        End Select

    End Sub
#End Region

#Region "grdAPTVEND9"

    Private Sub grdAPTVEND9_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVEND9.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""

                grdCodeDesc(grdAPTVEND9, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
        End Select
    End Sub

    Private Sub grdAPTVEND9_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVEND9.AfterExitEditMode
        With grdAPTVEND9
            Select Case .ActiveCell.Column.Key
                Case "ACCT_CODE"
                    Dim ACCT_CODE As String = .ActiveCell.Text
                    If ACCT_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With

    End Sub

    Private Sub grdAPTVEND9_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVEND9.AfterRowActivate
        With grdAPTVEND9
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdAPTVEND9.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdAPTVEND9_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTVEND9.BeforeRowUpdate
        With grdAPTVEND9
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                e.Cancel = True
            Else
                Call LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        Call LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("VEND_CODE").Text = "" Then
                    .ActiveRow.Cells("VEND_CODE").Value = Absx1.CtlFor("VEND_CODE").Text
                End If
            End If
        End With

    End Sub

    Private Sub grdAPTVEND9_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVEND9.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdAPTVEND9, sql_where, sql_where <> "")
    End Sub

#End Region


#Region "grdAPTVENDA"

    Private Sub grdAPTVENDA_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVENDA.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "VEND_ATTR_CODE"
                Dim VEND_ATTR_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdAPTVENDA, "APTATTR1", "VEND_ATTR_CODE", "VEND_ATTR_DESC")
        End Select
    End Sub

    Private Sub grdAPTVENDA_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVENDA.AfterExitEditMode
        With grdAPTVENDA
            Select Case .ActiveCell.Column.Key
                Case "VEND_ATTR_CODE"
                    Dim VEND_ATTR_CODE As String = .ActiveCell.Text
                    If VEND_ATTR_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(VEND_ATTR_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With

    End Sub

    Private Sub grdAPTVENDA_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVENDA.AfterRowActivate
        With grdAPTVENDA
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("VEND_ATTR_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdAPTVENDA.ActiveRow.Cells("VEND_ATTR_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                '.DisplayLayout.Bands(0).Columns("VEND_ATTR_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                ' why cant we edit the acct code?
            End If
        End With
    End Sub

    Private Sub grdAPTVENDA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTVENDA.BeforeRowUpdate
        With grdAPTVENDA
            If e.Row.Cells("VEND_ATTR_CODE").Text = "" Then
                e.Cancel = True
            Else
                Call LookUp("APTATTR1", e.Row.Cells("VEND_ATTR_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("VEND_ATTR_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.Cells("VEND_CODE").Text = "" Then
                    .ActiveRow.Cells("VEND_CODE").Value = Absx1.CtlFor("VEND_CODE").Text
                End If
            End If
        End With

    End Sub

    Private Sub grdAPTVENDA_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVENDA.ClickCellButton
        Dim sql_where As String = ""
        Call grdClickCellButton(grdAPTVENDA, sql_where, sql_where <> "")
    End Sub

#End Region

#Region "grdAPTVENDB"

    Private Sub grdAPTVENDB_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVENDB.AfterCellUpdate

        Select Case e.Cell.Column.Key
            'Case "ACCT_CODE"
            '    Dim ACCT_CODE As String = e.Cell.Value & ""

            '    grdCodeDesc(grdAPTVENDB, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
            '    For i As Integer = 2 To 4
            '        If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
            '            e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
            '        End If
            '    Next
        End Select
    End Sub

    Private Sub grdAPTVENDB_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVENDB.AfterExitEditMode
        With grdAPTVENDB
            'Select Case .ActiveCell.Column.Key
            '    Case "ACCT_CODE"
            '        Dim ACCT_CODE As String = .ActiveCell.Text
            '        If ACCT_CODE <> "" Then
            '            .ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, .ActiveCell.Column.Key)
            '        End If
            'End Select
        End With

    End Sub

    Private Sub grdAPTVENDB_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVENDB.AfterRowActivate
        With grdAPTVENDB
            'If .ActiveRow.IsAddRow Then
            '    .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            '    .ActiveCell = grdAPTVENDB.ActiveRow.Cells("ACCT_CODE")
            '    .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            'Else
            '    '.DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            '    ' why cant we edit the acct code?
            'End If
        End With
    End Sub

    Private Sub grdAPTVENDB_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdAPTVENDB.BeforeRowUpdate
        With grdAPTVENDB
            'If e.Row.Cells("ACCT_CODE").Text = "" Then
            '    e.Cancel = True
            'Else
            '    Call LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
            '    If cdr Is Nothing Then
            '        MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '        e.Cancel = True
            '    End If
            'End If

            'Dim COLUMN_NAME As String
            'For i As Integer = 2 To 4
            '    COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
            '    If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
            '        If e.Row.Cells(COLUMN_NAME).Text = "" Then
            '            e.Cancel = True
            '        Else
            '            Call LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
            '            If cdr Is Nothing Then
            '                MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '                e.Cancel = True
            '            End If
            '        End If
            '    End If
            'Next

            'If Not e.Cancel Then
            '    If e.Row.Cells("VEND_CODE").Text = "" Then
            '        .ActiveRow.Cells("VEND_CODE").Value = Absx1.CtlFor("VEND_CODE").Text
            '    End If
            'End If
        End With

    End Sub

    Private Sub grdAPTVENDB_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVENDB.ClickCellButton
        '    Dim sql_where As String = ""
        '    grdClickCellButton(grdAPTVENDB, sql_where, sql_where <> "")
    End Sub

    Private Sub UltraGroupBox5_Click(sender As Object, e As EventArgs) Handles grpBankingInfo.Click

    End Sub

#End Region

End Class