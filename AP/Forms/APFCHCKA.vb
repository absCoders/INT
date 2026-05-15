Imports System.Drawing

Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.Serialization
Imports Infragistics.Win.UltraWinGrid

Public Class APFCHCKA

    Dim rowAPTCHCKA As DataRow
    Dim XMIT_FILE_PATHANDNAME As String
    Dim APTCHCK1 As String = ""
    Dim FILENAME_SIGNED As String
    Dim FILENAME As String
    Dim rowGLTBANK1 As DataRow
    Dim SSH_APP_CODE As String

    Dim clsLastError As String = ""

    Dim BANK_CODE As String
    Dim PYMT_METHOD_CODE As String
    Dim BATCH_NO_PYMT As String

    Dim voided_checks As Integer = 0
    Dim issued_checks As Integer = 0

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("APTPARM1")

        ASCMAIN1.sql = "Select * from APTCHCK1 where ROWNUM < 1"
        APTCHCK1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & APTCHCK1 & " Add Primary Key (BANK_CODE, CHECK_NUM)")


        With dst
            ASCMAIN1.sql = "Select APTCHCKA.*" & vbCrLf _
                & " from APTCHCKA, (Select BATCH_NO_ACH" & vbCrLf _
                & ", COUNT (*) CHECKS" & vbCrLf _
                & ", SUM (CHECK_AMT) TOTAL_CHECK_AMT" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT = 0 THEN 1 ELSE 0 END) ZERO" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT < 0 THEN 1 ELSE 0 END) NEGC" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT < 0 THEN CHECK_AMT ELSE 0 END) NEGA" & vbCrLf _
                & " from APTCHCK1" & vbCrLf _
                & " group by BATCH_NO_ACH) X" & vbCrLf _
                & " where APTCHCKA.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and X.BATCH_NO_ACH (+) = APTCHCKA.BATCH_NO_ACH"

            Create_TDA(.Tables.Add, "APTCHCKX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "APTCHCKA", "*")

            Create_TDA(.Tables.Add, "TATCNTRY", "*", 0, False)
            Fill_Records("TATCNTRY")

            ASCMAIN1.sql = "Select APTCHCK1.*" _
                & " from APTCHCK1 where APTCHCK1.BATCH_NO_ACH = :PARM1"
            Create_TDA(.Tables.Add, "APTCHCK1", "**", 0, False, "V")

            ASCMAIN1.sql = "Select APTCHCK5.*" _
    & " from APTCHCK1,APTCHCK5 where APTCHCK1.BATCH_NO_ACH = :PARM1" & vbCrLf _
    & " and APTCHCK5.BANK_CODE = APTCHCK1.BANK_CODE and APTCHCK5.CHECK_NUM = APTCHCK1.CHECK_NUM"
            Create_TDA(.Tables.Add, "APTCHCK5", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select BATCH_NO_PYMT, BANK_CODE, PYMT_METHOD, CHECK_DATE" & vbCrLf _
                & ", MIN (VEND_CODE) MINVEND, MAX (VEND_CODE) MAXVEND" & vbCrLf _
                & ", COUNT (*) PYMTS, SUM (CHECK_AMT) TOTAL" & vbCrLf _
                & " from APTCHCK1 WHERE ACH_PAY_STATUS_IND = 'P' " & vbCrLf _
                & " group by BATCH_NO_PYMT, BANK_CODE, PYMT_METHOD, CHECK_DATE"
            Create_TDA(.Tables.Add, "APTCHCK0", "**", 0, False, "", 1)

            .Tables("APTCHCK0").Columns("PYMTS").DataType = GetType(System.Int32)
        End With

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        grdAPTCHCKX.DataSource = dst.Tables("APTCHCKX")
        grdAPTCHCK0.DataSource = dst.Tables("APTCHCK0")

        grdAPTCHCK1.DataSource = dst.Tables("APTCHCK1")

        Create_Summary(grdAPTCHCKX, "BATCH_NO_ACH", "Count")


        Create_Summary(grdAPTCHCK1, "CHECK_NUM", "Count")
        Create_Summary(grdAPTCHCK1, "CHECK_AMT")

        Create_Summary(grdAPTCHCK0, "BATCH_NO_PYMT", "Count")
        Create_Summary(grdAPTCHCK0, New String() {"TOTAL", "PYMTS"})

        With grdAPTCHCKX.DisplayLayout.Bands("APTCHCKX")
            .Columns("BATCH_NO_ACH").Header.Fixed = True
        End With

        grpHeader.Visible = False

        ASCMAIN1.Add_Value_List(grdAPTCHCKX, "BATCH_ACH_STATUS", Nothing, New String() {":", "P:Pending", "S:Sent", "R:Reset"})


        ASCMAIN1.Add_Value_List(grdAPTCHCK1, "CHECK_STATUS", Nothing, New String() {":", "I:Issued", "V:Voided"})
        ASCMAIN1.Add_Value_List(grdAPTCHCK1, "ACH_PAY_STATUS_IND", Nothing, New String() {":", "P:Pending", "S:Sent"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("BANK_CODE")

                Absx1.dteFor("XMIT_DATE").Value = DATETIME_STAMP.Date

                Dim DT As Date = Absx1.dteFor("XMIT_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                End If

                BANK_CODE = Absx1.txtFor("BANK_CODE").Text
                If BANK_CODE.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Bank"
                Else
                    rowGLTBANK1 = LookUp("GLTBANK1", BANK_CODE)
                    If IsNothing(rowGLTBANK1) Then
                        EMsg &= vbCr & "Bank Entered Is Not Valid"
                    Else
                        If rowGLTBANK1.Item("SSH_APP_CODE") & "" = "" Then
                            EMsg &= vbCr & $"Bank {BANK_CODE} is not set up for sftp Transmission"
                        End If
                    End If
                End If

                If Not ASCMAIN1.Logical_Lock("APTCHCKA", BANK_CODE) Then
                    Exit Sub
                End If

                PYMT_METHOD_CODE = Absx1.optFor("PYMT_METHOD_CODE").Value & ""
                If PYMT_METHOD_CODE = "" Then
                    EMsg &= vbCr & "You must supply a Valid Payment Method (ACH or WIRE)"
                End If

                BATCH_NO_PYMT = Absx1.txtFor("BATCH_NO_PYMT").Text
                If BATCH_NO_PYMT = "" Then
                    EMsg &= vbCr & "You must supply a Valid Payment Batch (double click a pending batch)"
                End If

                If EMsg = "" Then
                    ASCMAIN1.sql = "Select * from APTCHCK1" & vbCrLf _
                   & $" where BANK_CODE = '{BANK_CODE}'" & vbCrLf _
                   & $"   and PYMT_METHOD = '{PYMT_METHOD_CODE}'" & vbCrLf _
                   & $"   and BATCH_NO_PYMT = '{BATCH_NO_PYMT}'" & vbCrLf _
                   & "   and NVL(ACH_PAY_STATUS_IND,'0') = 'P'"
                    If ASCDATA1.GetDataTable.Rows.Count = 0 Then
                        EMsg &= vbCr & $"No {PYMT_METHOD_CODE} payments pending transmission for {BANK_CODE}"
                    Else
                        SSH_APP_CODE = rowGLTBANK1.Item("SSH_APP_CODE")
                    End If
                End If


            Case "View"
                If Absx1.txtFor("BATCH_NO_ACH").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowAPTCHCKA = LookUp("APTCHCKA", Absx1.txtFor("BATCH_NO_ACH").Text)
                    If rowAPTCHCKA Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("BATCH_NO_ACH").Text & " on File"
                    Else
                        SSH_APP_CODE = ""
                    End If
                End If

                BANK_CODE = rowAPTCHCKA.Item("BANK_CODE")
                If Not ASCMAIN1.Logical_Lock("APTCHCKA", BANK_CODE) Then
                    Exit Sub
                End If


            Case "Update"
                If grdAPTCHCK1.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Checks to Transmit"
                End If

            Case "Reset"
                If MsgBox("Are you sure that you want to Reset this batch?" _
                          & vbCrLf & "Note: After Reset, these payments will appear as pending Transmission", vbYesNo, "Verification") <> vbYes Then
                    Exit Sub
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Reset"
                Reset_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Transmit"
                ' uses Update

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode = "V" Then
                        .Items("Reset").Visible = (rowAPTCHCKA.Item("BATCH_ACH_STATUS") & "" <> "R")
                    Else
                        .Items("Reset").Visible = False
                    End If

                    If Not ScreenMode Then
                        .Items("Update").Visible = True
                    End If
                End With

                .Groups("Show if Transmitted in").Visible = Not ScreenMode
                .Groups("PGP Key Ring").Visible = False ' not complete - see C:\Users\Walter\Desktop\Interparfums\JPMC\PGP\KeyGen\Executable\openpgp.exe

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        spl0.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        'grdAPTCHCKX.Visible = Not ScreenMode
        SplitContainer1.Visible = Not ScreenMode
        spl0.Visible = ScreenMode

        If ScreenMode Then
            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"APTCHCKA", "APTCHCK1", "APTCHCK0"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
        Absx1.txtFor("BANK_CODE").Text = ""
        optPYMT_METHOD_CODE.Value = DBNull.Value
        Absx1.dteFor("XMIT_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("BATCH_NO_ACH").Text = ""

        lblBANK_DESC.Text = "Name"

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowAPTCHCKA = dst.Tables("APTCHCKA").NewRow
            With rowAPTCHCKA
                .Item("BATCH_NO_ACH") = ASCMAIN1.Next_Control_No("APTCHCKA.BATCH_NO_ACH")
                .Item("BANK_CODE") = BANK_CODE
                .Item("XMIT_DATE") = HFs("XMIT_DATE")
                .Item("BATCH_ACH_STATUS") = "S"
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("XMIT_FILE_PATHANDNAME") = XMIT_FILE_PATHANDNAME
                .Item("PYMT_METHOD_CODE") = PYMT_METHOD_CODE
                .Item("BATCH_NO_PYMT") = BATCH_NO_PYMT
            End With
            dst.Tables("APTCHCKA").Rows.Add(rowAPTCHCKA)


            Dim rowTATSSHK1 As DataRow = LookUp("TATSSHK1", SSH_APP_CODE)
            Dim SSH_APP_PARTNER_URI As String = rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD") & ""
            lblBANK_DESC.Text = "Name, URL: " & SSH_APP_PARTNER_URI


        Else
            rowAPTCHCKA = Fill_Record("APTCHCKA", Absx1.txtFor("BATCH_NO_ACH").Text)
            dst.AcceptChanges()


        End If

        If rowAPTCHCKA.Item("BATCH_ACH_STATUS") & "" = "R" Then
            Dim RESET_BY As String = rowAPTCHCKA.Item("LAST_OPER") & ""
            Dim RESET_DT As String = Format(rowAPTCHCKA.Item("LAST_DATE"), "MM/dd/yyyy HH:mm:ss")
            lblReset.Text = $"Trnasmission Batch was Reset by {RESET_BY} {RESET_DT}"
            lblReset.Visible = True
        Else
            lblReset.Visible = False
        End If

        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowAPTCHCKA.Item("BANK_CODE"))

        If EntryMode = "N" Then
            ASCDATA1.ExecuteSQL("Delete from " & APTCHCK1)
            ASCMAIN1.sql = $"Select * from APTCHCK1 where BANK_CODE = '{BANK_CODE}' and PYMT_METHOD = '{PYMT_METHOD_CODE}' and BATCH_NO_PYMT = '{BATCH_NO_PYMT}' and NVL(ACH_PAY_STATUS_IND,'0') = 'P'"
            ASCDATA1.ExecuteSQL("Insert into " & APTCHCK1 & " " & ASCMAIN1.sql)

            Fill_Records("APTCHCK1", "", True, "Select * from " & APTCHCK1)

            Dim sqlAPTCHCK5 As String = $"Select APTCHCK5.* from {APTCHCK1} APTCHCK1, APTCHCK5 
                where APTCHCK5.BANK_CODE = APTCHCK1.BANK_CODE and APTCHCK5.CHECK_NUM = APTCHCK1.CHECK_NUM"
            Fill_Records("APTCHCK5", "", True, sqlAPTCHCK5)

            Write_XML(Absx1.txtFor("BATCH_NO_ACH").Text)
        Else
            Fill_Records("APTCHCK1", Absx1.txtFor("BATCH_NO_ACH").Text)
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        ' TO RESET A BATCH
        ' update aptchck1 set ACH_PAY_STATUS_IND = 'P', BATCH_NO_ACH = NULL where BATCH_NO_ACH = '0000000249'
        ' PROB SHOULD DELETE APTCHCKA WHERE  where BATCH_NO_ACH = '0000000249' - BUT NOT SURE  SO LEAVING FOR NOW.

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Transmitting File")

        If issued_checks > 0 Then

            Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME_SIGNED)
            Dim FILENAME_SIGNED_ARCHIVE As String = ASCMAIN1.Folders("Archive") & "ACH\" & FI.Name
            If System.IO.File.Exists(FILENAME_SIGNED_ARCHIVE) Then
                System.IO.File.Move(FILENAME_SIGNED_ARCHIVE, FILENAME_SIGNED_ARCHIVE & Format(Now, "yyyyMMddhhmmss"))
            End If
            My.Computer.FileSystem.CopyFile(FILENAME_SIGNED, ASCMAIN1.Folders("Archive") & "ACH\" & FI.Name)


            FI = My.Computer.FileSystem.GetFileInfo(FILENAME & ".XML")
            Dim FILENAME_ARCHIVE As String = ASCMAIN1.Folders("Archive") & "ACH\" & FI.Name
            If System.IO.File.Exists(FILENAME_ARCHIVE) Then
                System.IO.File.Move(FILENAME_ARCHIVE, FILENAME_ARCHIVE & Format(Now, "yyyyMMddhhmmss"))
            End If
            My.Computer.FileSystem.CopyFile(FILENAME & ".XML", ASCMAIN1.Folders("Archive") & "ACH\" & FI.Name)
        End If

        Dim BATCH_NO_ACH As String = Absx1.txtFor("BATCH_NO_ACH").Text

        If ASCMAIN1.Running_in_VS Then
            Stop
        End If

        Dim in_production As Boolean = True

        If issued_checks > 0 Then
            'TAC.APCMAIN1.Send_Positive_Pay(Me, SSH_APP_CODE, in_production, FILENAME_SIGNED, BATCH_NO_ACH & "S")
            TAC.TACSCOM1.sftp_put(Me, SSH_APP_CODE, in_production, FILENAME_SIGNED, BATCH_NO_ACH & ".XML" & "S")
        End If

        BeginTrans()

        Update_Record_TDA("APTCHCKA")

        ASCMAIN1.sql = "Update APTCHCK1 Set ACH_PAY_STATUS_IND = 'S', BATCH_NO_ACH = :PARM1" _
            & " where (BANK_CODE, CHECK_NUM) in (Select BANK_CODE, CHECK_NUM from " & APTCHCK1 & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {BATCH_NO_ACH})

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    Sub Reset_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Resetting Batch")

        Dim BATCH_NO_ACH As String = Absx1.txtFor("BATCH_NO_ACH").Text

        BeginTrans()

        rowAPTCHCKA.Item("BATCH_ACH_STATUS") = "R"
        rowAPTCHCKA.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowAPTCHCKA.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("APTCHCKA")

        ' TO RESET A BATCH
        ' update aptchck1 set ACH_PAY_STATUS_IND = 'P', BATCH_NO_ACH = NULL where BATCH_NO_ACH = '0000000249'
        ' PROB SHOULD DELETE APTCHCKA WHERE  where BATCH_NO_ACH = '0000000249' - BUT NOT SURE  SO LEAVING FOR NOW.

        ASCMAIN1.sql = "Update APTCHCK1 Set ACH_PAY_STATUS_IND = 'P', BATCH_NO_ACH = NULL" _
            & " where BATCH_NO_ACH = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {BATCH_NO_ACH})

        CommitTrans("Reset Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdAPTCHCKX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdAPTCHCK1, "B", "Check Inquiry")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            Case "Check Inquiry"
                Dim BANK_CODE As String = grd.ActiveRow.Cells("BANK_CODE").Text
                Dim CHECK_NUM As String = grd.ActiveRow.Cells("CHECK_NUM").Text
                Dim rowAPTCHCK1 As DataRow = LookUp("APTCHCK1", New String() {BANK_CODE, CHECK_NUM})
                If rowAPTCHCK1 IsNot Nothing Then
                    Context_Launch("View", CHECK_NUM, e.Tool.Key, "APFCHCKI")
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                'If e.KeyCode = Windows.Forms.Keys.Enter Then
                '    If Not InquiryMode Then
                '        Click_Command("New", e)
                '    End If
                'End If
            Case "BATCH_NO_ACH"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BANK_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "BATCH_NO_ACH"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "BANK_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

    Private Sub grdAPTCHCKX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTCHCKX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("BATCH_NO_ACH").Text = e.Row.Cells("BATCH_NO_ACH").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim YP As String = cbeYP.Value
        Fill_Records("APTCHCKX", YP)
        Sort_grdColumns(grdAPTCHCKX, "BATCH_NO_ACH".ToLower)
        grdAPTCHCKX.Text = "Transmitted in " & cbeYP.Text

        Fill_Records("APTCHCK0")
        Sort_grdColumns(grdAPTCHCK0, "BATCH_NO_PYMT")
        'grdAPTCHCK0.Text = "Transmitted in " & cbeYP.Text
    End Sub

    Sub Write_XML(BATCH_NO_ACH As String)
        'If ASCMAIN1.Running_in_VS Then Stop
        FILENAME = ASCMAIN1.Folders("Temp") & BATCH_NO_ACH

        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", BANK_CODE)

        Dim tbl As DataTable = dst.Tables("APTCHCK1")

        voided_checks = tbl.Select("CHECK_STATUS <> 'I'").Length
        issued_checks = tbl.Select("CHECK_STATUS = 'I'").Length

        If issued_checks > 0 Then

            Dim chase As New Document

            Dim DT As Date = Now + ASCMAIN1.NowTSD

            ' Dim G As New chase.CstmrCdtTrfInitn.GrpHd
            chase.CstmrCdtTrfInitn = New CustomerCreditTransferInitiationV03
            chase.CstmrCdtTrfInitn.GrpHdr = New GroupHeader32

            With chase.CstmrCdtTrfInitn.GrpHdr
                .MsgId = BATCH_NO_ACH

                .CreDtTm = $"{Format(DATETIME_STAMP, "yyyy-MM-dd")}T{Format(DATETIME_STAMP, "hh:mm:ss")}"
                .NbOfTxs = issued_checks
                .CtrlSum = Format(Val(tbl.Compute("SUM(CHECK_AMT)", "CHECK_STATUS = 'I'")), "###0.00")
                .CtrlSumSpecified = True
                chase.CstmrCdtTrfInitn.GrpHdr.InitgPty = New PartyIdentification32
                .InitgPty.Nm = rowGLTBANK1.Item("BANK_NAME")
            End With

            ReDim chase.CstmrCdtTrfInitn.PmtInf(0)

            chase.CstmrCdtTrfInitn.PmtInf(0) = New PaymentInstructionInformation3

            Dim CHECK_AMT_TOTAL As Decimal = Val(tbl.Compute("SUM(CHECK_AMT)", "CHECK_STATUS = 'I'"))
            Dim CHECK_AMT_MIN As Decimal = Val(tbl.Compute("MIN(CHECK_AMT)", "CHECK_STATUS = 'I'"))

            With chase.CstmrCdtTrfInitn.PmtInf(0)
                .PmtInfId = BATCH_NO_ACH
                .PmtMtd = New PaymentMethod3Code

                .PmtMtd = PaymentMethod3Code.TRF ' "TRF"
                .NbOfTxs = issued_checks
                .CtrlSum = Format(CHECK_AMT_TOTAL, "###0.00")
                .CtrlSumSpecified = True

                .PmtTpInf = New PaymentTypeInformation19

                .PmtTpInf.SvcLvl = New ServiceLevel8Choice
                '.PmtTpInf.SvcLvl.Item("Cd") = "NURG"
                .PmtTpInf.SvcLvl.ItemElementName = ItemChoiceType4.Cd ' "Cd"

                Dim XMLVALUE As String = ""


                ' May be ""URGP" Or ""BKTR".
                ' "URGP" Is tested for Book Transfer And automatically converted if it qualifies.
                ' Required If Not provided at transaction level."


                ' Stop ' ANTHONY SAYS THAT IF THE ROUTING NO IS THE SAME AS 021000021 THEN I MIGHT NEED TO SET THIS UP USING BKTR
                ' 11/17 - AM assures use that JPMC will automatically convert FEDWIRE to BANKTRANSER
                'Stop ' BUT THAT WOULD MEAN THAT I CANNOT SEND WIRES TO CHASE VENDORS IN THE SAME BATCH AS WIRES TO NON-CHASE VENDORS

                If PYMT_METHOD_CODE = "ACH" Then
                    XMLVALUE = "NURG"
                ElseIf PYMT_METHOD_CODE = "WIRE" Then
                    XMLVALUE = "URGP"
                    'If XMLVALUE = "SAME BANK As JPMC" Then
                    '    XMLVALUE = "BKTR"
                    'End If
                End If
                .PmtTpInf.SvcLvl.Item = XMLVALUE



                If PYMT_METHOD_CODE = "ACH" Then
                    .PmtTpInf.LclInstrm = New LocalInstrument2Choice
                    .PmtTpInf.LclInstrm.ItemElementName = ItemChoiceType5.Cd ' "Cd"
                    .PmtTpInf.LclInstrm.Item = "CCD"

                    If CHECK_AMT_MIN < 1 Then
                        .PmtTpInf.CtgyPurp = New CategoryPurpose1Choice
                        .PmtTpInf.CtgyPurp.ItemElementName = ItemChoiceType6.Prtry ' "Prtry"
                        .PmtTpInf.CtgyPurp.Item = "ACCTVERIFY"
                    End If

                End If

                .ReqdExctnDt = $"{Format(DATETIME_STAMP, "yyyy-MM-dd")}"


                .Dbtr = New PartyIdentification32
                .Dbtr.Nm = rowGLTBANK1.Item("ACCT_NAME")
                .Dbtr.PstlAdr = New PostalAddress6
                .Dbtr.PstlAdr.Ctry = "US"

                If PYMT_METHOD_CODE = "ACH" Then
                    ' AM 11/29 Please remove <Id> tag from <Dbtr> group. Guessing it’s just a remnant from the OrgId element that wasn’t supposed to be there since that’s only for ACH.
                    .Dbtr.Id = New Party6Choice

                    Dim item As OrganisationIdentification4 = New OrganisationIdentification4
                    Dim itemOther As GenericOrganisationIdentification1 = New GenericOrganisationIdentification1
                    itemOther.Id = rowGLTBANK1.Item("BANK_COMP_ID") ' "9272278928" ' "IPLB_COMP_ID"

                    ' RELOCATED ABOVE
                    'Dim item As OrganisationIdentification4 = New OrganisationIdentification4
                    'Dim itemOther As GenericOrganisationIdentification1 = New GenericOrganisationIdentification1
                    'itemOther.Id = rowGLTBANK1.Item("BANK_COMP_ID") ' "9272278928" ' "IPLB_COMP_ID"

                    Dim itemOtherScheNm As OrganisationIdentificationSchemeName1Choice = New OrganisationIdentificationSchemeName1Choice
                    itemOtherScheNm.ItemElementName = ItemChoiceType.Prtry  ' "Prtry"
                    itemOtherScheNm.Item = "JPMCOID"
                    itemOther.SchmeNm = itemOtherScheNm

                    ' RELOCATED BELOW
                    ReDim item.Othr(0)
                    item.Othr(0) = itemOther
                    .Dbtr.Id.Item = item
                End If

                'ReDim item.Othr(0)
                'item.Othr(0) = itemOther
                '.Dbtr.Id.Item = item

                .DbtrAcct = New CashAccount16

                Dim cashacct As CashAccount16 = New CashAccount16
                Dim DbtrAcctId As AccountIdentification4Choice = New AccountIdentification4Choice
                Dim DbtrAcctIdOthr As GenericAccountIdentification1 = New GenericAccountIdentification1
                DbtrAcctIdOthr.Id = rowGLTBANK1.Item("BANK_ACCT_ID")
                DbtrAcctId.Item = DbtrAcctIdOthr
                cashacct.Id = DbtrAcctId
                cashacct.Ccy = "USD"
                .DbtrAcct = cashacct

                'DbtrAcctId.Othr(0) = DbtrAcctIdOthr

                'DbtrAcctId.item("Othr") = DbtrAcctIdOthr

                '.DbtrAcct.Id = DbtrAcctId
                '.DbtrAcct.Ccy = "USD"

                .DbtrAgt = New BranchAndFinancialInstitutionIdentification4
                Dim itemFinInstnId As New FinancialInstitutionIdentification7
                itemFinInstnId.BIC = rowGLTBANK1.Item("BANK_BIC") '  "CHASUS33"
                ' Stop ' ANTHONY SAYS THAT i MIGHT NEED TO PUT THE CHASU33 CODE IN THE CrtrAgt IF I AM DOING A BOOK XFR
                ' 11/17 AM SAYS OK TO HARD CODE CHASU33 FOR ALL ACH AND WIRE TYPES

                itemFinInstnId.ClrSysMmbId = New ClearingSystemMemberIdentification2
                itemFinInstnId.ClrSysMmbId.MmbId = rowGLTBANK1.Item("ROUTING_NO")

                itemFinInstnId.PstlAdr = New PostalAddress6
                itemFinInstnId.PstlAdr.Ctry = "US"
                chase.CstmrCdtTrfInitn.PmtInf(0).DbtrAgt.FinInstnId = itemFinInstnId

                ReDim .CdtTrfTxInf(tbl.Rows.Count - 1)

                Dim pmtIndex As Integer = -1
                For Each row As DataRow In tbl.Select("CHECK_STATUS = 'I'", "CHECK_NUM")
                    pmtIndex += 1

                    Dim CHECK_NUM As String = row.Item("CHECK_NUM")

                    .CdtTrfTxInf(pmtIndex) = New CreditTransferTransactionInformation10

                    With .CdtTrfTxInf(pmtIndex)

                        .PmtId = New PaymentIdentification1
                        .PmtId.InstrId = row.Item("CHECK_NUM")
                        .PmtId.EndToEndId = row.Item("CHECK_NUM")
                        .Amt = New AmountType3Choice

                        Dim amtxx As ActiveOrHistoricCurrencyAndAmount = New ActiveOrHistoricCurrencyAndAmount
                        amtxx.Ccy = "USD"
                        amtxx.Value = Format(Val(row.Item("CHECK_AMT") & ""), "###0.00")


                        .Amt.Item = amtxx

                        Dim VEND_CODE As String = row.Item("VEND_CODE")
                        Dim VEND_CODE_AP As String = row.Item("VEND_CODE_AP") & ""
                        Dim VEND_ALT_CODE As String = row.Item("VEND_ALT_CODE") & ""

                        Dim colpfx As String = "VEND_"
                        Dim rowV As DataRow = Nothing
                        If VEND_CODE_AP <> VEND_CODE And VEND_CODE_AP <> "" Then
                            rowV = LookUp("APTVEND1", VEND_CODE_AP)
                        ElseIf VEND_ALT_CODE <> "" And VEND_ALT_CODE <> "VENDOR" Then
                            rowV = LookUp("APTVEND2", New String() {VEND_CODE, VEND_ALT_CODE})
                            colpfx = "VEND_ALT_"
                        Else
                            rowV = LookUp("APTVEND1", VEND_CODE)
                        End If

                        Dim rowAPTCHCK5 As DataRow = dst.Tables("APTCHCK5").Rows.Find(New String() {BANK_CODE, CHECK_NUM})

                        Dim VEND_BANK_ACCT_ID As String = rowAPTCHCK5.Item("VEND_BANK_ACCT_ID")
                        VEND_BANK_ACCT_ID = ASCMAIN1.DecryptAES(VEND_BANK_ACCT_ID)
                        Dim VEND_BANK_ROUTING_NO As String = rowAPTCHCK5.Item("VEND_BANK_ROUTING_NO") & ""
                        Dim VEND_BANK_SWIFT_NO As String = rowAPTCHCK5.Item("VEND_BANK_SWIFT_NO") & ""
                        Dim VEND_BANK_ACCT_CLASS As String = rowAPTCHCK5.Item("VEND_BANK_ACCT_CLASS")
                        Dim VEND_BANK_ACCT_TYPE As String = rowAPTCHCK5.Item("VEND_BANK_ACCT_TYPE")
                        Dim VEND_BANK_NAME As String = rowAPTCHCK5.Item("VEND_BANK_NAME") & ""
                        Dim VEND_BANK_ADDR1 As String = rowAPTCHCK5.Item("VEND_BANK_ADDR1") & ""
                        Dim VEND_BANK_ADDR2 As String = rowAPTCHCK5.Item("VEND_BANK_ADDR2") & ""
                        Dim VEND_BANK_ADDR3 As String = rowAPTCHCK5.Item("VEND_BANK_ADDR3") & ""
                        Dim VEND_BANK_CITY As String = rowAPTCHCK5.Item("VEND_BANK_CITY") & ""
                        Dim VEND_BANK_STATE As String = rowAPTCHCK5.Item("VEND_BANK_STATE") & ""
                        Dim VEND_BANK_COUNTRY As String = rowAPTCHCK5.Item("VEND_BANK_COUNTRY")

                        Dim rowTATCNTRY As DataRow = dst.Tables("TATCNTRY").Rows.Find(VEND_BANK_COUNTRY)
                        Dim VEND_BANK_COUNTRY2 As String = rowTATCNTRY.Item("COUNTRY_CODE2") & ""

                        ' Stop ' ANTHONY SAYS THAT i MIGHT NEED TO PUT THE CHASU33 CODE IN THE CrtrAgt IF I AM DOING A BOOK XFR

                        .CdtrAgt = New BranchAndFinancialInstitutionIdentification4
                        .CdtrAgt.FinInstnId = New FinancialInstitutionIdentification7

                        ' AM 11/29 From  <FinInstnId> group, please remove <ClrSysMmbId> tag as IPLB won’t be providing CHIPS UID.

                        If (PYMT_METHOD_CODE = "WIRE" And VEND_BANK_COUNTRY <> "USA") Then
                        Else
                            .CdtrAgt.FinInstnId.ClrSysMmbId = New ClearingSystemMemberIdentification2
                        End If

                        If (PYMT_METHOD_CODE = "WIRE" And VEND_BANK_COUNTRY <> "USA") Then
                            .CdtrAgt.FinInstnId.BIC = VEND_BANK_SWIFT_NO
                            ' .CdtrAgt.FinInstnId.ClrSysMmbId.MmbId = VEND_BANK_SWIFT_NO
                            ' AM SAYS DO NOT PROVIDE ANYTHING IN MmbId for CHIPS - and if you do it should be the CHIPS UID
                            ' https://www.theclearinghouse.org/uid-lookup
                        Else
                            .CdtrAgt.FinInstnId.ClrSysMmbId.MmbId = VEND_BANK_ROUTING_NO
                            '.CdtrAgt.FinInstnId.BIC = VEND_BANK_ROUTING_NO
                            '11/17 AM SAYS IT IS OK TO USE BIC FOR ALL

                        End If

                        If VEND_BANK_NAME <> "" Then .CdtrAgt.FinInstnId.Nm = VEND_BANK_NAME
                        .CdtrAgt.FinInstnId.PstlAdr = New PostalAddress6

                        If VEND_BANK_CITY <> "" Then .CdtrAgt.FinInstnId.PstlAdr.TwnNm = VEND_BANK_CITY
                        If VEND_BANK_STATE <> "" Then .CdtrAgt.FinInstnId.PstlAdr.CtrySubDvsn = VEND_BANK_STATE
                        If VEND_BANK_COUNTRY2 <> "" Then .CdtrAgt.FinInstnId.PstlAdr.Ctry = VEND_BANK_COUNTRY2

                        .Cdtr = New PartyIdentification32
                        .Cdtr.Nm = rowV.Item($"{colpfx}NAME")

                        .Cdtr.PstlAdr = New PostalAddress6
                        'If rowV.Item($"{colpfx}COUNTRY") & "" <> "" Then .Cdtr.PstlAdr.Ctry = rowV.Item($"{colpfx}COUNTRY") & ""
                        If VEND_BANK_COUNTRY2 <> "" Then .Cdtr.PstlAdr.Ctry = VEND_BANK_COUNTRY2

                        'ReDim .Cdtr.PstlAdr.AdrLine(2)
                        Dim AD As Integer = -1
                        Dim ADDR1 As String = rowV.Item($"{colpfx}ADDR1") & ""
                        If ADDR1 <> "" Then
                            AD += 1 : ReDim Preserve .Cdtr.PstlAdr.AdrLine(AD)
                            .Cdtr.PstlAdr.AdrLine(AD) = ADDR1
                        End If
                        Dim ADDR2 As String = rowV.Item($"{colpfx}ADDR2") & ""
                        If ADDR2 <> "" Then
                            AD += 1 : ReDim Preserve .Cdtr.PstlAdr.AdrLine(AD)
                            .Cdtr.PstlAdr.AdrLine(AD) = ADDR2
                        End If
                        Dim VEND_CSZ As String = rowV.Item($"{colpfx}CITY") & "," & rowV.Item($"{colpfx}STATE") & " " & rowV.Item($"{colpfx}ZIP_CODE")
                        If rowV.Item($"{colpfx}CITY") & "" <> "" AndAlso rowV.Item($"{colpfx}STATE") & "" <> "" AndAlso rowV.Item($"{colpfx}ZIP_CODE") & "" <> "" Then
                            AD += 1 : ReDim Preserve .Cdtr.PstlAdr.AdrLine(AD)
                            .Cdtr.PstlAdr.AdrLine(AD) = VEND_CSZ
                        End If

                        .CdtrAcct = New CashAccount16
                        .CdtrAcct.Id = New AccountIdentification4Choice
                        Dim cdtrAcctOthr As GenericAccountIdentification1 = New GenericAccountIdentification1
                        cdtrAcctOthr.Id = VEND_BANK_ACCT_ID
                        '.CdtrAcct.Id.Item("Oth") = cdtrAcctOthr
                        .CdtrAcct.Id.Item = cdtrAcctOthr

                        If PYMT_METHOD_CODE = "ACH" Then
                            .CdtrAcct.Tp = New CashAccountType2
                            Dim VEND_BANK_ACCT_TYPE_code As String = IIf(VEND_BANK_ACCT_TYPE = "S", "SVGS", "CHKG")
                            '.CdtrAcct.Tp.Item("Cd") = VEND_BANK_ACCT_TYPE_code
                            Dim ct As CashAccountType4Code
                            If VEND_BANK_ACCT_TYPE = "S" Then ct = CashAccountType4Code.SVGS
                            If VEND_BANK_ACCT_TYPE = "C" Then ct = CashAccountType4Code.CASH

                            '.CdtrAcct.Tp.Item("Cd") = VEND_BANK_ACCT_TYPE_code
                            .CdtrAcct.Tp.Item = ct
                        End If

                        ' presence of a remittance value changes to CCD+

                        .RmtInf = New RemittanceInformation5
                        ReDim .RmtInf.Ustrd(0)

                        Dim INV_NUMs As String = ""
                        ASCMAIN1.sql = $"Select INV_NUM from APTCHCK2 where BANK_CODE = '{BANK_CODE}' And CHECK_NUM = '{CHECK_NUM}'"
                        For Each row2 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select
                            Dim INV_NUM As String = row2.Item("INV_NUM")
                            If INV_NUMs.Length + INV_NUM.Length + 1 > 70 Then
                                INV_NUMs &= "+"
                                Exit For
                            End If
                            INV_NUMs &= "," & INV_NUM
                        Next
                        INV_NUMs = Mid(INV_NUMs, 2)
                        If INV_NUMs.Length > 71 Then
                            INV_NUMs = Mid(INV_NUMs, 1, 71)
                        End If
                        .RmtInf.Ustrd(0) = INV_NUMs
                    End With

                Next
            End With


            'Dim XMLH As XMLHelper
            'XMLH.SerializeXml(chase, FILENAME & ".XML")
            SerializeXml(chase, FILENAME & ".XML")

            Dim vv As New XmlValidationErrorBuilder
            Dim XSD As String = ASCMAIN1.Folders("Archive") & SSH_APP_CODE & "\Chase.xsd"
            If ASCMAIN1.Running_in_VS Then
                XSD = "C:\Users\wjz\Desktop\Interparfums\" & SSH_APP_CODE & "\Chase.xsd"
            End If
            'LoadValidatedXmlDocument(FILENAME & ".xml", "C:\VS\AHA\Work\Chase.xsd")
            LoadValidatedXmlDocument(FILENAME & ".xml", XSD)
            If clsLastError <> "" Then
                Using frm As New ASFMSGBF
                    frm.Show_Formatted_txt("XML Document Errors", clsLastError, Me)
                End Using

                UltraExplorerBar1.Groups("Screen Control").Items("Update").Visible = False

            Else

                FILENAME_SIGNED = FILENAME & ".XML" & "S"
                'Sign_File(Me, SSH_APP_CODE, FILENAME)
                Sign_File_nSoftware(Me, SSH_APP_CODE, FILENAME & ".XML")
            End If


        End If

        If voided_checks > 0 Then
            MsgBox("There are " & CStr(voided_checks) & " Voided Checks in this batch", MsgBoxStyle.OkOnly, "Verfication")
        End If

    End Sub

    Public Sub Sign_File_nSoftware(frmASFBASE0 As ASFBASE0, SSH_APP_CODE As String, FILENAME As String, Optional FILENAME_SIGNED As String = "")

        Dim rowTATSSHK1 As DataRow = ASCDATA1.GetDataRow("Select * from TATSSHK1 where SSH_APP_CODE = '" & SSH_APP_CODE & "'")
        Dim SSH_APP_PGP_PVTKEY_PWD As String = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY_PWD") & ""
        'Dim SSH_APP_PGP_PVTKEY As String = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY") & ""

        Dim openpgp1 As New nsoftware.IPWorksEncrypt.Openpgp
        openpgp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareEncryptionkey")

        Dim asciiArmor As Boolean = True

        If FILENAME_SIGNED = "" Then
            FILENAME_SIGNED = FILENAME & "S"
        End If

        Try
            openpgp1.Reset()

            openpgp1.Overwrite = False

            openpgp1.InputFile = FILENAME
            openpgp1.OutputFile = FILENAME_SIGNED
            openpgp1.ASCIIArmor = asciiArmor

            Dim KEY_FOLDER As String = ASCMAIN1.Folders("Archive") & SSH_APP_CODE
            If ASCMAIN1.Running_in_VS Then
                KEY_FOLDER = "C:\Users\wjz\Desktop\Interparfums\" & SSH_APP_CODE
            End If
            openpgp1.Keys.Add(New nsoftware.IPWorksEncrypt.Key(KEY_FOLDER, SSH_APP_CODE))
            openpgp1.Keys(0).Passphrase = SSH_APP_PGP_PVTKEY_PWD
            openpgp1.RecipientKeys.Add(New nsoftware.IPWorksEncrypt.Key(KEY_FOLDER, SSH_APP_CODE))

            If System.IO.File.Exists(FILENAME_SIGNED) Then
                If MsgBox("Signed File Exists - Do you want to overwrite?", MsgBoxStyle.YesNo, "Please Contact ABS") = MsgBoxResult.Yes Then
                    'System.IO.File.Delete(FILENAME_SIGNED)
                    System.IO.File.Move(FILENAME_SIGNED, FILENAME_SIGNED & Format(Now, "yyyyMMddhhmmss"))

                End If
            End If
            openpgp1.Sign()

        Catch ex As nsoftware.IPWorksEncrypt.IPWorksEncryptException
            MessageBox.Show("Error: " + ex.Message)
        End Try

    End Sub

    Private Sub grdAPTCHCK1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTCHCK1.InitializeRow
        If e.Row.Cells("CHECK_STATUS").Value = "V" Then
            e.Row.Appearance.ForeColor = Color.Red
        End If
    End Sub

    Private Sub cmdRegenKeyRing_Click(sender As Object, e As EventArgs) Handles cmdRegenKeyRing.Click

        'If ASCMAIN1.Running_in_VS Then
        Dim SSH_APP_CODE As String = "JPMC"
        Dim rowTATSSHK1 As DataRow = LookUp("TATSSHK1", SSH_APP_CODE)
        If rowTATSSHK1 Is Nothing Then
            MsgBox("Cannot Find Encryption Parameter Record for " & SSH_APP_CODE, MsgBoxStyle.OkOnly, "Cannot Continue")
            Exit Sub
        End If

        Dim SSH_APP_PGP_PUBKEY As String = rowTATSSHK1.Item("SSH_APP_PGP_PUBKEY") & ""
        Dim SSH_APP_PGP_PVTKEY As String = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY") & ""
        Dim SSH_APP_PGP_PVTKEY_PWD As String = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY_PWD") & ""

        Stop
        ' CODE BELOW IS TO CREATE A KEYRING BY IMPORTING PUBLIC AND PRIVATE KEYS
        ' THIS WILL CREATE A NEW KEYRING USING THE PUBLIC AND PRIVATE KEYS GENERATED BY WHATEVER APP WAS USED TO GENERATE THEM
        ' NSOFTWAARE CREATES (AND EXPECTS) FILES NAMED secring.gpp AND pubring.gpg
        ' https://www.nsoftware.com/kb/articles/openpgp.rst
        ' search for keymgr1 and other text that talks about signing files
        Dim keymgr1 As New nsoftware.IPWorksEncrypt.Keymgr
        keymgr1.CreateKey(SSH_APP_CODE, SSH_APP_PGP_PVTKEY_PWD)
        keymgr1.ImportKey("C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_IPLB_pvt.asc", "")
        keymgr1.ImportKey("C:\Users\wjz\Desktop\Interparfums\JPMC\JPMC_IPLB_pub.asc", "")
        keymgr1.SaveKeyring("C:\Users\wjz\Desktop\Interparfums\JPMC")

        'End If
    End Sub



    Sub SerializeXml(
    ByVal toSerialize As Object,
    ByVal targetFileName As String)

        Dim sw As New StreamWriter(targetFileName)
        Dim ser As New XmlSerializer(toSerialize.GetType())
        ser.Serialize(sw, toSerialize)
        sw.Close()

    End Sub


    Private Function LoadValidatedXmlDocument(ByVal xmlFilePath As String, ByVal xsdFilePath As String) As XmlDocument
        clsLastError = String.Empty

        Try
            Dim doc As New XmlDocument()
            doc.Load(xmlFilePath)
            doc.Schemas.Add(Nothing, xsdFilePath)

            Dim errorBuilder As New XmlValidationErrorBuilder()
            doc.Validate(New ValidationEventHandler(AddressOf errorBuilder.ValidationEventHandler))
            Dim errorsText As String = errorBuilder.GetErrors()

            If errorsText IsNot Nothing Then
                Throw New Exception(errorsText)
            End If
            Return doc

        Catch ex As Exception
            clsLastError = ex.Message
            Return Nothing
        End Try

    End Function

    Private Class XmlValidationErrorBuilder
        Private _errors As New List(Of ValidationEventArgs)()

        Public Sub ValidationEventHandler(ByVal sender As Object, ByVal args As ValidationEventArgs)
            If args.Severity = XmlSeverityType.Error Then
                _errors.Add(args)
            End If
        End Sub

        Public Function GetErrors() As String
            If _errors.Count <> 0 Then
                Dim builder As New StringBuilder()
                builder.Append("The following ")
                builder.Append(_errors.Count.ToString())
                builder.AppendLine(" error(s) were found while validating the XML document against the XSD:")
                For Each i As ValidationEventArgs In _errors
                    builder.Append("* ")
                    builder.AppendLine(i.Message)
                Next
                Return builder.ToString()
            Else
                Return Nothing
            End If
        End Function
    End Class

    Private Sub grdAPTCHCK0_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdAPTCHCK0.DoubleClickRow
        If e.Row.IsDataRow Then
            Dim BANK_CODE As String = e.Row.Cells("BANK_CODE").Value
            Dim PYMT_METHOD As String = e.Row.Cells("PYMT_METHOD").Value
            Dim BATCH_NO_PYMT As String = e.Row.Cells("BATCH_NO_PYMT").Value

            Absx1.txtFor("BANK_CODE").Text = BANK_CODE
            Absx1.optFor("PYMT_METHOD_CODE").Value = PYMT_METHOD
            Absx1.txtFor("BATCH_NO_PYMT").Text = BATCH_NO_PYMT

            Click_Command("New")
        End If
    End Sub
End Class



Public Class XMLHelper

    Public Shared Sub SerializeXml(
    ByVal toSerialize As Object,
    ByVal targetFileName As String)

        Dim sw As New StreamWriter(targetFileName)
        Dim ser As New XmlSerializer(toSerialize.GetType())
        ser.Serialize(sw, toSerialize)
        sw.Close()

    End Sub

    Public Shared Function DeSerializeXml(
    ByVal toDeSerializeType As Type,
    ByVal sourceFileName As String) As Object

        Dim sr As New StreamReader(sourceFileName)
        Dim ser As New XmlSerializer(toDeSerializeType)
        Dim o As Object = ser.Deserialize(sr)
        sr.Close()

        Return o
    End Function

End Class

