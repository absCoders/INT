Imports System.Drawing
Imports System.Math

Public Class APFCHCKN

    Dim rowAPTCHCKN As DataRow
    Dim XMIT_FILE_PATHANDNAME As String
    Dim APTCHCK1 As String = ""
    Dim FILENAME_SIGNED As String
    Dim FILENAME As String
    Dim rowGLTBANK1 As DataRow
    Dim SSH_APP_CODE As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("APTPARM1")

        ASCMAIN1.sql = "Select * from APTCHCK1 where ROWNUM < 1"
        APTCHCK1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & APTCHCK1 & " Add Primary Key (BANK_CODE, CHECK_NUM)")


        With dst
            ASCMAIN1.sql = "Select APTCHCKN.*" & vbCrLf _
                & " from APTCHCKN, (Select BATCH_NO_NA" & vbCrLf _
                & ", COUNT (*) CHECKS" & vbCrLf _
                & ", SUM (CHECK_AMT) TOTAL_CHECK_AMT" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT = 0 THEN 1 ELSE 0 END) ZERO" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT < 0 THEN 1 ELSE 0 END) NEGC" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT < 0 THEN CHECK_AMT ELSE 0 END) NEGA" & vbCrLf _
                & " from APTCHCK1" & vbCrLf _
                & " group by BATCH_NO_NA) X" & vbCrLf _
                & " where APTCHCKN.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and X.BATCH_NO_NA = APTCHCKN.BATCH_NO_NA"

            Create_TDA(.Tables.Add, "APTCHCKX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "APTCHCKN", "*")

            ASCMAIN1.sql = "Select APTCHCK1.*" _
                & " from APTCHCK1 where APTCHCK1.BATCH_NO_NA = :PARM1"
            Create_TDA(.Tables.Add, "APTCHCK1", "**", 0, False, "V")
        End With

        grdAPTCHCKX.DataSource = dst.Tables("APTCHCKX")

        grdAPTCHCK1.DataSource = dst.Tables("APTCHCK1")

        Create_Summary(grdAPTCHCKX, "BATCH_NO_NA", "Count")


        Create_Summary(grdAPTCHCK1, "CHECK_NUM", "Count")
        Create_Summary(grdAPTCHCK1, "CHECK_AMT")

        With grdAPTCHCKX.DisplayLayout.Bands("APTCHCKX")
            .Columns("BATCH_NO_NA").Header.Fixed = True
        End With

        grpHeader.Visible = False

        ASCMAIN1.Add_Value_List(grdAPTCHCKX, "BATCH_NA_STATUS", Nothing, New String() {":", "P:Pending", "S:Sent"})

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        ASCMAIN1.Add_Value_List(grdAPTCHCK1, "CHECK_STATUS", Nothing, New String() {":", "I:Issued", "V:Voided"})
        ASCMAIN1.Add_Value_List(grdAPTCHCK1, "ACH_PAY_STATUS_IND", Nothing, New String() {":", "P:Pending", "S:Sent"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("BANK_CODE")

                Dim DT As Date = Absx1.dteFor("XMIT_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                End If

                If Absx1.txtFor("BANK_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Bank"
                Else
                    rowGLTBANK1 = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                    If IsNothing(rowGLTBANK1) Then
                        EMsg &= vbCr & "Bank Entered Is Not Valid"
                    Else
                        If rowGLTBANK1.Item("SSH_APP_CODE") & "" = "" Then
                            EMsg &= vbCr & "Bank is not set up for Secure Transmission"
                        Else

                            ASCMAIN1.sql = "Select * from APTCHCK1" & vbCrLf _
                           & " where BANK_CODE = '" & Absx1.txtFor("BANK_CODE").Text & "'" & vbCrLf _
                           & "   and NVL(ACH_PAY_STATUS_IND,'0') = 'P'"
                            If ASCDATA1.GetDataTable.Rows.Count = 0 Then
                                EMsg &= vbCr & "No payments pending transmission for " & Absx1.txtFor("BANK_CODE").Text
                            Else
                                SSH_APP_CODE = rowGLTBANK1.Item("SSH_APP_CODE")
                            End If
                        End If
                    End If
                End If

            Case "View"
                If Absx1.txtFor("BATCH_NO_NA").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowAPTCHCKN = LookUp("APTCHCKN", Absx1.txtFor("BATCH_NO_NA").Text)
                    If rowAPTCHCKN Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("BATCH_NO_NA").Text & " on File"
                    End If
                End If

            Case "Update"
                If grdAPTCHCK1.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Checks to Transmit"
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
                End With

                .Groups("Show if Transmitted in").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        spl0.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        grdAPTCHCKX.Visible = Not ScreenMode
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
        For Each TABLE_NAME As String In New String() {"APTCHCKN", "APTCHCK1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
        Absx1.txtFor("BANK_CODE").Text = ""
        Absx1.dteFor("XMIT_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("BATCH_NO_NA").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowAPTCHCKN = dst.Tables("APTCHCKN").NewRow
            rowAPTCHCKN.Item("BATCH_NO_NA") = ASCMAIN1.Next_Control_No("APTCHCKN.BATCH_NO_NA")
            rowAPTCHCKN.Item("BANK_CODE") = HFs("BANK_CODE")
            rowAPTCHCKN.Item("XMIT_DATE") = HFs("XMIT_DATE")
            rowAPTCHCKN.Item("BATCH_NA_STATUS") = "S"
            rowAPTCHCKN.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowAPTCHCKN.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowAPTCHCKN.Item("INIT_DATE") = DATETIME_STAMP
            rowAPTCHCKN.Item("XMIT_FILE_PATHANDNAME") = XMIT_FILE_PATHANDNAME
            dst.Tables("APTCHCKN").Rows.Add(rowAPTCHCKN)
        Else
            Fill_Record("APTCHCKN", Absx1.txtFor("BATCH_NO_NA").Text)
            dst.AcceptChanges()
        End If

        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowAPTCHCKN.Item("BANK_CODE"))

        If EntryMode = "N" Then
            ASCDATA1.ExecuteSQL("Delete from " & APTCHCK1)
            ASCMAIN1.sql = "Select * from APTCHCK1 where BANK_CODE = '" & HFs("BANK_CODE") & "' and NVL(ACH_PAY_STATUS_IND,'0') = 'P'"
            ASCDATA1.ExecuteSQL("Insert into " & APTCHCK1 & " " & ASCMAIN1.sql)

            Fill_Records("APTCHCK1", "", True, "Select * from " & APTCHCK1)

            Write_File(Absx1.txtFor("BATCH_NO_NA").Text)
        Else
            Fill_Records("APTCHCK1", Absx1.txtFor("BATCH_NO_NA").Text)
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Transmitting File")

        Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME_SIGNED)
        My.Computer.FileSystem.CopyFile(FILENAME_SIGNED, ASCMAIN1.Folders("Archive") & "ACH\" & FI.Name)
        FI = My.Computer.FileSystem.GetFileInfo(FILENAME)
        My.Computer.FileSystem.CopyFile(FILENAME, ASCMAIN1.Folders("Archive") & "ACH\" & FI.Name)

        Dim BATCH_NO_NA As String = Absx1.txtFor("BATCH_NO_NA").Text

        If ASCMAIN1.Running_in_VS Then
            Stop
        Else
            Dim in_production As Boolean = True
            TAC.TACSCOM1.sftp_put(Me, SSH_APP_CODE, in_production, FILENAME_SIGNED, BATCH_NO_NA & "S")
        End If

        BeginTrans()

        Update_Record_TDA("APTCHCKN")

        ASCMAIN1.sql = "Update APTCHCK1 Set ACH_PAY_STATUS_IND = 'S', BATCH_NO_NA = :PARM1" _
            & " where (BANK_CODE, CHECK_NUM) in (Select BANK_CODE, CHECK_NUM from " & APTCHCK1 & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("BATCH_NO_NA").Text})

        CommitTrans("Update Complete")
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
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "BATCH_NO_NA"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BANK_CODE"
                If Not InquiryMode Then
                    Click_Command("New")
                End If
            Case "BATCH_NO_NA"
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

    Private Sub grdICTIADJX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTCHCKX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("BATCH_NO_NA").Text = e.Row.Cells("BATCH_NO_NA").Text
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
        Sort_grdColumns(grdAPTCHCKX, "BATCH_NO_NA".ToLower)
        grdAPTCHCKX.Text = "Transmitted in " & cbeYP.Text

    End Sub

    Sub Write_File(BATCH_NO_NA As String)

        ' LINE TERMINATION VBCRLF OR VBLF
        ' LASALLE VS STD FEDERAL
        ' AUTOMATED FILE TRANSFER
        ' RETURN FILE
        ' HANDLING STOP PAYMENTS
        ' TESTING
        ' RETURNED PAYMENTS (NOT DISBURSABLE)
        ' ACH CREDITS
        ' BATCH (6 OR 10) -> 7

        ' PREFERENCE FOR 'PAYMENT' IN NACHA FILE
        ' VENDOR CODES IN NACHA FILE
        ' PAYMENT METHODS - ACH DEBIT, ACH CREDIT
        ' HANDLING VOIDED CHECKS
        ' FED ID IN NACHA FILE
        ' VENDOR ACH INFO IN GRID
        ' CLEAN UP VENDOR PYMT METHODS
        ' DISCUSS USE OF PYMT METHOD IN PYMT SELECTION - AND WARNINGS
        ' DO WE RUN CHECK PRINTING FOR ACH PAYMENTS?
        ' SET UP A NEW BANK CODE FOR ACH PAYMENTS? IF NOT, THEN WE NEED TO DO SOMETHING SPECIAL IN CHECK PRINTING


        If ASCMAIN1.Running_in_VS Then Stop
        FILENAME = ASCMAIN1.Folders("Temp") & BATCH_NO_NA & ".txt"

        Dim voided_checks As Integer = 0
        Dim LINES As Integer = 0

        Dim FID As String = "07100050" ' ?
        Dim FID_NAME As String = "LaSalle" ' ?
        Dim FID_TR As String = " 071000505" ' ?

        Using sw As New System.IO.StreamWriter(FILENAME)
            Dim BANK_CODE As String = Absx1.txtFor("BANK_CODE").Text
            Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", BANK_CODE)

            Dim T As String

            ' File Header Record

            Dim YYYYMMDD As String = Format(DATETIME_STAMP, "yyMMdd")
            Dim HHMM As String = Format(DATETIME_STAMP, "HHmm")
            Dim ID As String = ROWs("APTPARM1").Item("AP_PARM_1099_TAX_ID")

            T = "".PadLeft(94)
            Mid(T, 1, 1) = "1"              ' Record Type Code
            Mid(T, 2, 2) = "01"             ' Priority Code
            Mid(T, 4, 10) = FID_TR          ' Immediate Desination (LaSalle Bank or Standard Federal Bank) ?
            Mid(T, 14, 10) = ID             ' Immediate Origin - need Ahava Fed Tax ID No ?
            Mid(T, 24, 6) = YYYYMMDD        ' File Creation Date
            Mid(T, 30, 4) = HHMM            ' File Creation Time
            Mid(T, 34, 1) = "A"             ' File ID Modifier
            Mid(T, 35, 3) = "094"           ' Record Size
            Mid(T, 38, 2) = "10"            ' Blocking Factor
            Mid(T, 40, 1) = "1"             ' Format Code
            Mid(T, 41, 23) = FID_NAME       ' Immediate Destination Name
            Mid(T, 64, 23) = "Ahava"        ' Immediate Origin Name
            Mid(T, 87, 8) = ASCMAIN1.USER_ID ' Reference Code

            sw.Write(T & vbLf) ' for unix style
            LINES += 1


            ' Batch Header Record

            Dim ACCT_NAME As String = rowGLTBANK1.Item("ACCT_NAME")

            T = "".PadLeft(94)
            Mid(T, 1, 1) = "5"              ' Record Type Code
            Mid(T, 2, 3) = "200"            ' Service Class Code
            Mid(T, 5, 16) = ACCT_NAME       ' Company Name
            Mid(T, 21, 20) = ""             ' Discretionary Data
            Mid(T, 41, 10) = ID             ' Company Identification
            Mid(T, 51, 3) = "PPD"           ' Standard Entry Class ?
            Mid(T, 54, 10) = "PAYMENT"      ' Company Entry Description ?
            Mid(T, 64, 6) = ""              ' Company Descriptive Date ?
            Mid(T, 70, 6) = YYYYMMDD        ' Effective Entry Date
            Mid(T, 76, 3) = ""              ' Settlement Date (Reserved)
            Mid(T, 79, 1) = "1"             ' Originator Status Code
            Mid(T, 80, 8) = FID             ' Originating DFI Identification
            Mid(T, 88, 7) = Mid(BATCH_NO_NA, 4) ' Batch Number

            sw.Write(T & vbLf) ' for unix style
            LINES += 1

            Dim RECORDS As Integer = 0
            Dim DR_AMTS As Decimal = 0
            Dim CR_AMTS As Decimal = 0

            ASCMAIN1.sql = "Select * from " & APTCHCK1

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("CHECK_AMT > 0", "CHECK_NUM")

                Dim CHECK_NUM As String = row.Item("CHECK_NUM")
                Dim VEND_CODE As String = row.Item("VEND_CODE")
                Dim VEND_NAME As String = row.Item("VEND_CODE")

                Dim CHECK_AMT As Decimal = Val(row.Item("CHECK_AMT") & "")

                If CHECK_AMT >= 0 Then
                    DR_AMTS += CHECK_AMT
                Else
                    CR_AMTS += CHECK_AMT
                End If

                If CHECK_AMT <= 0 Then
                    Throw New Exception("Negative Amount: " & Format(CHECK_AMT, "#,##0.00") & " in Check " & row.Item("CHECK_NUM"))
                    Stop ' only positive dollar amounts permitted
                End If

                'Dim T As String = "".PadLeft(94)
                T = "".PadLeft(94)
                Mid(T, 1, 1) = "6"      ' Record Type Code
                Mid(T, 2, 2) = "22"     ' Transaction Code
                Mid(T, 4, 8) = ""  ' Receiving DFI Identification
                Mid(T, 12, 1) = ""      ' Check Digit
                Mid(T, 13, 17) = ""          ' DFI Account Number 
                Mid(T, 30, 10) = Format(100 * CHECK_AMT, "0000000000")      ' Amount
                Mid(T, 40, 15) = VEND_CODE  ' Individual Identification Number
                Mid(T, 55, 22) = VEND_NAME  ' Individual Name
                Mid(T, 77, 2) = ""      ' Discretionary Data
                Mid(T, 79, 1) = "0"     ' Addenda Record Indicator
                Mid(T, 80, 15) = ""      ' Trace Number

                sw.Write(T & vbLf) ' for unix style
                'sw.WriteLine(T) ' for windows style
                LINES += 1

                RECORDS += 1
            Next



            ' Batch Control Record

            ' Dim ACCT_NAME As String = rowGLTBANK1.Item("ACCT_NAME")

            T = "".PadLeft(94)
            Mid(T, 1, 1) = "8"              ' Record Type Code
            Mid(T, 2, 3) = "200"            ' Service Class Code
            Mid(T, 5, 6) = Format(RECORDS, "000000")       ' Entry / Addenda Count
            Mid(T, 11, 10) = ""             ' Entry Hash - no 6 records
            Mid(T, 21, 12) = Format(100 * DR_AMTS, "000000000000")    ' Total Debit Entry Dollar Amount
            Mid(T, 33, 12) = Format(100 * CR_AMTS, "000000000000")    ' Total Credit Entry Dollar Amount
            Mid(T, 45, 10) = ID             ' Company Identification
            Mid(T, 55, 19) = ""              ' Message Authentication Code
            Mid(T, 74, 6) = ""              ' Reserved (Federal Reserve Use)
            Mid(T, 80, 8) = FID             ' Originating Financial Institution ID

            Mid(T, 88, 7) = Mid(BATCH_NO_NA, 4) ' Batch Number

            sw.Write(T & vbLf) ' for unix style
            LINES += 1

            If LINES < 10 Then
                For I As Integer = 1 To 10 - LINES
                    T = "".PadLeft(94, "9")
                    sw.Write(T & vbLf) ' for unix style
                Next
            End If
        End Using


        If voided_checks > 0 Then
            MsgBox("There are " & CStr(voided_checks) & " Voided Checks in this batch", MsgBoxStyle.OkOnly, "Verfication")
        End If

        ' FILENAME_SIGNED = FILENAME & "S"

        ' Sign_File_nSoftware(Me, SSH_APP_CODE, FILENAME)
    End Sub

    Public Sub Sign_File_nSoftware(frmASFBASE0 As ASFBASE0, SSH_APP_CODE As String, FILENAME As String, Optional FILENAME_SIGNED As String = "")

        Dim rowTATSSHK1 As DataRow = ASCDATA1.GetDataRow("Select * from TATSSHK1 where SSH_APP_CODE = '" & SSH_APP_CODE & "'")
        ' NOTE: SSH_APP_PASSWORD IS THE PASSWORD FOR sftp, NOT the password for Signing files, 
        ' although it is the same in this case
        Dim SSH_APP_PGP_PVTKEY_PWD As String = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY_PWD") & ""
        Dim SSH_APP_PGP_PVTKEY As String = rowTATSSHK1.Item("SSH_APP_PGP_PVTKEY") & ""

        'Dim pgp As New DidiSoft.Pgp.PGPLib()
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

            'Stop ' fix these lines

            Dim KEY_FOLDER As String = ASCMAIN1.Folders("Archive") & SSH_APP_CODE
            If ASCMAIN1.Running_in_VS Then
                KEY_FOLDER = "C:\Users\wjz\Desktop\Interparfums\" & SSH_APP_CODE
            End If
            openpgp1.Keys.Add(New nsoftware.IPWorksEncrypt.Key(KEY_FOLDER, SSH_APP_CODE))
            'openpgp1.Keys.Add(New nsoftware.IPWorksEncrypt.Key(SSH_APP_PGP_PVTKEY))
            openpgp1.Keys(0).Passphrase = SSH_APP_PGP_PVTKEY_PWD
            'openpgp1.RecipientKeys.Add(New Key(txtKeyringDir.Text, cboRecipientKeys.Text))
            openpgp1.RecipientKeys.Add(New nsoftware.IPWorksEncrypt.Key(KEY_FOLDER, SSH_APP_CODE))

            'openpgp1.SignAndEncrypt()
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
End Class