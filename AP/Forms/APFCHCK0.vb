Imports System.Drawing
Imports System.Math

Public Class APFCHCK0

    Dim rowAPTCHCK0 As DataRow
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
            ASCMAIN1.sql = "Select APTCHCK0.*" & vbCrLf _
                & " from APTCHCK0, (Select BATCH_NO_PP" & vbCrLf _
                & ", COUNT (*) CHECKS" & vbCrLf _
                & ", SUM (CHECK_AMT) TOTAL_CHECK_AMT" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT = 0 THEN 1 ELSE 0 END) ZERO" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT < 0 THEN 1 ELSE 0 END) NEGC" & vbCrLf _
                & ", SUM (CASE WHEN CHECK_AMT < 0 THEN CHECK_AMT ELSE 0 END) NEGA" & vbCrLf _
                & " from APTCHCK1" & vbCrLf _
                & " group by BATCH_NO_PP) X" & vbCrLf _
                & " where APTCHCK0.OPS_YYYYPP = :PARM1" & vbCrLf _
                & "   and X.BATCH_NO_PP = APTCHCK0.BATCH_NO_PP"

            Create_TDA(.Tables.Add, "APTCHCKX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "APTCHCK0", "*")

            ASCMAIN1.sql = "Select APTCHCK1.*" _
                & " from APTCHCK1 where APTCHCK1.BATCH_NO_PP = :PARM1"
            Create_TDA(.Tables.Add, "APTCHCK1", "**", 0, False, "V")
        End With
 
        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        grdAPTCHCKX.DataSource = dst.Tables("APTCHCKX")

        grdAPTCHCK1.DataSource = dst.Tables("APTCHCK1")

        Create_Summary(grdAPTCHCKX, "BATCH_NO_PP", "Count")


        Create_Summary(grdAPTCHCK1, "CHECK_NUM", "Count")
        Create_Summary(grdAPTCHCK1, "CHECK_AMT")

        With grdAPTCHCKX.DisplayLayout.Bands("APTCHCKX")
            .Columns("BATCH_NO_PP").Header.Fixed = True
        End With

        grpHeader.Visible = False

        ASCMAIN1.Add_Value_List(grdAPTCHCKX, "BATCH_PP_STATUS", Nothing, New String() {":", "P:Pending", "S:Sent"})


        ASCMAIN1.Add_Value_List(grdAPTCHCK1, "CHECK_STATUS", Nothing, New String() {":", "I:Issued", "V:Voided"})
        ASCMAIN1.Add_Value_List(grdAPTCHCK1, "POS_PAY_STATUS_IND", Nothing, New String() {":", "P:Pending", "S:Sent"})
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
                            EMsg &= vbCr & "Bank is not set up for Postive Pay"
                        Else

                            ASCMAIN1.sql = "Select * from APTCHCK1" & vbCrLf _
                           & " where BANK_CODE = '" & Absx1.txtFor("BANK_CODE").Text & "'" & vbCrLf _
                           & "   and NVL(POS_PAY_STATUS_IND,'0') = 'P'"
                            If ASCDATA1.GetDataTable.Rows.Count = 0 Then
                                EMsg &= vbCr & "No Checks pending transmission for " & Absx1.txtFor("BANK_CODE").Text
                            Else
                                SSH_APP_CODE = rowGLTBANK1.Item("SSH_APP_CODE")
                            End If
                        End If
                    End If
                End If

            Case "View"
                If Absx1.txtFor("BATCH_NO_PP").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowAPTCHCK0 = LookUp("APTCHCK0", Absx1.txtFor("BATCH_NO_PP").Text)
                    If rowAPTCHCK0 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("BATCH_NO_PP").Text & " on File"
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
                .Groups("PGP Key Ring").Visible = False ' not complete - see C:\Users\Walter\Desktop\Interparfums\JPMC\PGP\KeyGen\Executable\openpgp.exe

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
        For Each TABLE_NAME As String In New String() {"APTCHCK0", "APTCHCK1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()
        Absx1.txtFor("BANK_CODE").Text = ""
        Absx1.dteFor("XMIT_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("BATCH_NO_PP").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowAPTCHCK0 = dst.Tables("APTCHCK0").NewRow
            rowAPTCHCK0.Item("BATCH_NO_PP") = ASCMAIN1.Next_Control_No("APTCHCK0.BATCH_NO_PP")
            rowAPTCHCK0.Item("BANK_CODE") = HFs("BANK_CODE")
            rowAPTCHCK0.Item("XMIT_DATE") = HFs("XMIT_DATE")
            rowAPTCHCK0.Item("BATCH_PP_STATUS") = "S"
            rowAPTCHCK0.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowAPTCHCK0.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowAPTCHCK0.Item("INIT_DATE") = DATETIME_STAMP
            rowAPTCHCK0.Item("XMIT_FILE_PATHANDNAME") = XMIT_FILE_PATHANDNAME
            dst.Tables("APTCHCK0").Rows.Add(rowAPTCHCK0)
        Else
            Fill_Record("APTCHCK0", Absx1.txtFor("BATCH_NO_PP").Text)
            dst.AcceptChanges()
        End If

        Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowAPTCHCK0.Item("BANK_CODE"))

        If EntryMode = "N" Then
            ASCDATA1.ExecuteSQL("Delete from " & APTCHCK1)
            ASCMAIN1.sql = "Select * from APTCHCK1 where BANK_CODE = '" & HFs("BANK_CODE") & "' and NVL(POS_PAY_STATUS_IND,'0') = 'P'"
            ASCDATA1.ExecuteSQL("Insert into " & APTCHCK1 & " " & ASCMAIN1.sql)

            Fill_Records("APTCHCK1", "", True, "Select * from " & APTCHCK1)

            Write_Positive_Pay(Absx1.txtFor("BATCH_NO_PP").Text)
        Else
            Fill_Records("APTCHCK1", Absx1.txtFor("BATCH_NO_PP").Text)
        End If

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        ' TO RESET A BATCH
        ' update aptchck1 set POS_PAY_STATUS_IND = 'P', BATCH_NO_PP = NULL where batch_no_pp = '0000000249'
        ' PROB SHOULD DELETE APTCHCK0 WHERE  where batch_no_pp = '0000000249' - BUT NOT SURE  SO LEAVING FOR NOW.

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Transmitting File")

        Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME_SIGNED)
        My.Computer.FileSystem.CopyFile(FILENAME_SIGNED, ASCMAIN1.Folders("Archive") & "PP\" & FI.Name)
        FI = My.Computer.FileSystem.GetFileInfo(FILENAME)
        My.Computer.FileSystem.CopyFile(FILENAME, ASCMAIN1.Folders("Archive") & "PP\" & FI.Name)

        Dim BATCH_NO_PP As String = Absx1.txtFor("BATCH_NO_PP").Text

        If ASCMAIN1.Running_in_VS Then
            Stop
        Else
            Dim in_production As Boolean = True
            'TAC.APCMAIN1.Send_Positive_Pay(Me, SSH_APP_CODE, in_production, FILENAME_SIGNED, BATCH_NO_PP & "S")
            TAC.TACSCOM1.sftp_put(Me, SSH_APP_CODE, in_production, FILENAME_SIGNED, BATCH_NO_PP & "S")
        End If
 
        BeginTrans()

        Update_Record_TDA("APTCHCK0")

        ASCMAIN1.sql = "Update APTCHCK1 Set POS_PAY_STATUS_IND = 'S', BATCH_NO_PP = :PARM1" _
            & " where (BANK_CODE, CHECK_NUM) in (Select BANK_CODE, CHECK_NUM from " & APTCHCK1 & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("BATCH_NO_PP").Text})

        CommitTrans("Update Complete")

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
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode Then
                        Click_Command("New", e)
                    End If
                End If
            Case "BATCH_NO_PP"
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
            Case "BATCH_NO_PP"
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
            Absx1.txtFor("BATCH_NO_PP").Text = e.Row.Cells("BATCH_NO_PP").Text
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
        Sort_grdColumns(grdAPTCHCKX, "BATCH_NO_PP".ToLower)
        grdAPTCHCKX.Text = "Transmitted in " & cbeYP.Text
 
    End Sub

    Sub Write_Positive_Pay(BATCH_NO_PP As String)
        'If ASCMAIN1.Running_in_VS Then Stop
        FILENAME = ASCMAIN1.Folders("Temp") & BATCH_NO_PP

        Dim voided_checks As Integer = 0

        Using sw As New System.IO.StreamWriter(FILENAME)
            Dim BANK_CODE As String = Absx1.txtFor("BANK_CODE").Text
            Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", BANK_CODE)

            '  If ASCMAIN1.DBS_COMPANY <> "EXP" Then Stop

            ASCMAIN1.sql = "Select * from " & APTCHCK1

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("CHECK_AMT > 0", "CHECK_NUM")

                If Val(row.Item("CHECK_AMT") & "") <= 0 Then
                    Throw New Exception("Negative Amount: " & Format(Val(row.Item("CHECK_AMT") & ""), "#,##0.00") & " in Check " & row.Item("CHECK_NUM"))
                    Stop ' only positive dollar amounts permitted
                End If

                Dim T As String = "".PadLeft(200)
                If row.Item("CHECK_STATUS") = "I" Then
                    Mid(T, 1, 1) = "I"      ' I=Issued, V=Voided, S=Stop
                Else
                    Mid(T, 1, 1) = "V"      ' I=Issued, V=Voided, S=Stop
                    voided_checks += 1
                End If
                Mid(T, 2, 1) = ""       ' Space
                Mid(T, 3, 20) = CStr(rowGLTBANK1.Item("BANK_ACCT_ID") & "").PadLeft(20, "0")    ' Bank Account No
                Mid(T, 23, 1) = ""      ' Space
                Mid(T, 24, 18) = CStr(row.Item("CHECK_NUM") & "").PadLeft(18, "0")              ' Check No
                Mid(T, 42, 1) = ""      ' Space
                Mid(T, 43, 18) = Format(100 * Val(row.Item("CHECK_AMT") & ""), "000000000000000000")  ' positive amounts only, no decimal point
                Mid(T, 61, 1) = ""      ' Space
                Mid(T, 62, 8) = Format(row.Item("CHECK_DATE"), "yyyyMMdd")                      ' Check Date - no slashes
                Mid(T, 70, 1) = ""      ' Space
                Mid(T, 71, 8) = ""      ' Paid Date
                Mid(T, 79, 1) = ""      ' Space
                Mid(T, 80, 15) = ""     ' Additional Information pertaining to Check
                Mid(T, 95, 50) = row.Item("VEND_NAME") & "" ' Expanded Additional Information - Payee Name if desired
                Mid(T, 145, 50) = ""    ' 2nd Payee Name (only if we subscribe to Payee Name Verification Service)
                Mid(T, 195, 6) = ""       ' Spaces

                sw.Write(T & vbLf) ' for unix style
                'sw.WriteLine(T) ' for windows style
            Next
        End Using

        If voided_checks > 0 Then
            MsgBox("There are " & CStr(voided_checks) & " Voided Checks in this batch", MsgBoxStyle.OkOnly, "Verfication")
        End If

        FILENAME_SIGNED = FILENAME & "S"
        'Sign_File(Me, SSH_APP_CODE, FILENAME)
        Sign_File_nSoftware(Me, SSH_APP_CODE, FILENAME)
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
            openpgp1.Sign()

        Catch ex As nsoftware.IPWorksEncrypt.IPWorksEncryptException
            MessageBox.Show("Error: " + ex.Message)
        End Try

    End Sub

    Private Sub grdAPTCHCK1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTCHCK1.InitializeLayout

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
End Class