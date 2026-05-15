Imports System.Threading.Tasks
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports Newtonsoft.Json
Imports System.Reflection
Imports Infragistics.Win.UltraWinGrid

Public Class GLFBREC1

    Dim rowGLTBANK1 As DataRow
    Dim rowGLTBREC1 As DataRow
    Dim rowGLTPARM2 As DataRow

    Dim BATCH_NO_CLEARED As String
    Dim BANK_CODE As String
    Dim OPS_YYYYPP As String
    Dim BANK_STMT_BALANCE_previous As Decimal
    Dim IS_API_BANK As Boolean = False

    Dim GLTBREC2 As String
    Dim ACCESS_TOKEN As String
    Dim XNO As String
    Dim XDATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))
    Dim XNO_LNO As Integer = 0
    Dim XNO_LNO_0 As Integer = 0
    Dim XNOs As New List(Of String)
    Dim unmatchedRowFound As Boolean = False
    Dim BANK_CODE_FILTER As String

    Dim BANK_AMT_OS_APCD_LM As Decimal
    Dim BANK_BOOK_BALANCE As Decimal

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        For i As Integer = 1 To tabs.Tabs.Count - 1
            tabs.Tabs(i).Visible = False
        Next
        tabs.SelectedTab = tabs.Tabs(0)
        MakeTransparent(chkShowAll)
        MakeTransparent(chk_showReconciled)
        SplitContainer1.Panel2Collapsed = True
        UltraExplorerBar1.Groups("Reconcile").Items("Save Reconcile").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Reconcile").Items("Cancel Reconcile").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Reconcile").Items("Clear Reconciled").Settings.Enabled = DefaultableBoolean.False
        With dst
            '  Create_TDA(.Tables.Add, "GLTBREC1", "*")

            ASCMAIN1.sql = "Select GLTBREC1.*, GLTBANK1.BANK_DESC, GLTPARM2.LEGEND" _
                & " from GLTBREC1,GLTBANK1,GLTPARM2" _
                & " where GLTBANK1.BANK_CODE = GLTBREC1.BANK_CODE" _
                & "   and GLTPARM2.OPS_YYYYPP = GLTBREC1.OPS_YYYYPP" _
                & "   and GLTBREC1.BANK_CODE = NVL(:PARM1,GLTBREC1.BANK_CODE)"
            Create_TDA(.Tables.Add, "GLTBRECX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select GLTBREC1.*, GLTBANK1.BANK_DESC, GLTPARM2.LEGEND" _
                & " from GLTBREC1,GLTBANK1,GLTPARM2" _
                & " where GLTBANK1.BANK_CODE = GLTBREC1.BANK_CODE" _
                & "   and GLTPARM2.OPS_YYYYPP = GLTBREC1.OPS_YYYYPP" _
                & "   and GLTBREC1.BATCH_NO_CLEARED = :PARM1"
            Create_TDA(.Tables.Add, "GLTBREC1", "**", 0, True, "V")

            GLTBREC2 = ASCMAIN1.Temp_Table("Select GLTBREC2.* from GLTBREC2 where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & GLTBREC2 & " Add Primary Key (BATCH_NO_CLEARED,JOURNAL_TYPE,TRAN_YP,TRAN_KEY,TRAN_KEY_LNO)")
            ASCMAIN1.sql = "Select * from " & GLTBREC2
            Create_TDA(.Tables.Add("GLTBREC2"), GLTBREC2, "**", 0)
            .Tables("GLTBREC2").Columns("TRAN_SEL").DefaultValue = "0"
            With .Tables("GLTBREC2").Columns
                .Add("TRAN_AMT_ARCR", GetType(System.Decimal), "IIF(JOURNAL_TYPE='ARCR',IIF(TRAN_SEL='1',TRAN_AMT,0),NULL)")
                .Add("TRAN_AMT_APCD", GetType(System.Decimal), "IIF(JOURNAL_TYPE='APCD',IIF(TRAN_SEL='1',TRAN_AMT,0),NULL)")
                .Add("TRAN_AMT_GLJE", GetType(System.Decimal), "IIF(JOURNAL_TYPE='ARCR' OR JOURNAL_TYPE='APCD',NULL,IIF(TRAN_SEL='1',TRAN_AMT,0))")
            End With

            With .Tables.Add("GLTBRECT")
                .Columns.Add("T_LNO", GetType(System.Int16))
                .Columns.Add("T_DESC")
                .Columns.Add("T_AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("T_LNO")}
            End With

            ASCMAIN1.sql = "Select * from GLTTYPE1"
            Create_TDA(.Tables.Add, "GLTTYPE1", "**", 0, False)

            ASCMAIN1.sql = "Select * from GLTCHBL1"
            Create_TDA(.Tables.Add, "GLTCHBL1", "**", 0, True)

            ASCMAIN1.sql = "Select * from GLTCHBL2"
            Create_TDA(.Tables.Add, "GLTCHBL2", "**", 0, True)

            ASCMAIN1.sql = "Select * from GLTCHTR0"
            Create_TDA(.Tables.Add, "GLTCHTR0", "**", 0, True)

            ASCMAIN1.sql = "Select * from GLTCHTR1"
            Create_TDA(.Tables.Add, "GLTCHTR1", "**", 0, True, ,, "MATCHED, MATCHED_AMT, RECONCILING_ITEM_IND, RECONCILING_ITEM_REASON")
            .Tables("GLTCHTR1").Columns("RECONCILE").DefaultValue = "0"

            ASCMAIN1.sql = "Select * from GLTBANK1"
            Create_TDA(.Tables.Add, "GLTBANK1", "**", 0, True)

            ASCMAIN1.sql = "Select * from GLTBANK3"
            Create_TDA(.Tables.Add, "GLTBANK3", "**", 0, True)

            ASCMAIN1.sql = "SELECT JOURNAL_TYPE, TRAN_YP, TRAN_KEY, CUST_VEND, TRAN_DATE, TRAN_DESC, TRAN_AMT FROM GLTBREC2 WHERE TRAN_SEL <> '1' AND JOURNAL_TYPE = 'APCD'"
            Create_TDA(.Tables.Add, "GLTOSREC", "**", 0, False)

            'ASCMAIN1.sql = "Select * from ARTPYMT2 WHERE ROWNUM < 1"
            'Create_TDA(.Tables.Add, "ARTPYMT2", "**", 0, True)
            ASCMAIN1.sql = "Select * from GLTBRECR"
            Create_TDA(.Tables.Add, "GLTBRECR", "**", 0, True)
        End With

        Fill_Records("GLTTYPE1")
        Fill_Records("GLTBANK1")
        Fill_Records("GLTBANK3")

        grdGLTBREC2.DataSource = dst.Tables("GLTBREC2")
        Dim dvwGLTMATCH As New DataView(dst.Tables("GLTBREC2"))
        grdGLTMATCH.DataSource = dvwGLTMATCH
        grdGLTBRECX.DataSource = dst.Tables("GLTBRECX")
        grdGLTBRECT.DataSource = dst.Tables("GLTBRECT")
        grdGLTCHTR1.DataSource = dst.Tables("GLTCHTR1")
        'grdGLTMATCH.DataSource = dst.Tables("GLTBREC2")
        grdGLTOSREC.DataSource = dst.Tables("GLTOSREC")

        With grdGLTBREC2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With

        With grdGLTCHTR1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "MATCHED" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdGLTMATCH.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "RECONCILE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        Create_Summary(grdGLTBREC2, "TRAN_KEY", "Count")
        Create_Summary(grdGLTBREC2, New String() {"TRAN_AMT", "TRAN_SEL", "TRAN_AMT_ARCR", "TRAN_AMT_APCD", "TRAN_AMT_GLJE"})

        Create_Summary(grdGLTBRECX, "BATCH_NO_CLEARED", "Count")

        Create_Summary(grdGLTCHTR1, "ACCOUNT_ACCOUNTID", "Count")
        Create_Summary(grdGLTCHTR1, "MATCHED")
        Create_Summary(grdGLTCHTR1, "RECONCILING_ITEM_IND")
        Create_Summary(grdGLTCHTR1, "MATCHED_AMT")
        'Create_Summary(grdGLTCHTR1, "AMOUNT")
        Create_Summary(grdGLTCHTR1, "RECONCILE")
        Create_Summary(grdGLTCHTR1, "RECONCILE_AMT")

        Create_Summary(grdGLTMATCH, "TRAN_KEY", "Count")
        'Create_Summary(grdGLTMATCH, New String() {"TRAN_AMT", "RECONCILE", "RECONCILE_AMT"})
        Create_Summary(grdGLTMATCH, "TRAN_AMT", "Sum")
        Create_Summary(grdGLTMATCH, "RECONCILE", "Sum")
        Create_Summary(grdGLTMATCH, "RECONCILE_AMT", "Sum")
        Create_Summary(grdGLTOSREC, "JOURNAL_TYPE", "Count")
        Create_Summary(grdGLTOSREC, "TRAN_AMT", "Sum")


        With dst.Tables("GLTBRECT").Rows
            .Add(New Object() {1, "Prev Balance", 0})
            .Add(New Object() {2, "Deposits", 0})
            .Add(New Object() {3, "Disbursements", 0})
            .Add(New Object() {4, "Adjustments", 0})
            .Add(New Object() {5, "LM Outstanding", 0})
            .Add(New Object() {6, "Roll Forward", 0})
            .Add(New Object() {7, "Stmt Balance", 0})
            .Add(New Object() {8, "Reconciling Items", 0})
            .Add(New Object() {9, "Difference", 0})
            .Add(New Object() {10, "Book Balance", 0})
            .Add(New Object() {11, "TM Outstanding", 0})
        End With
        Sort_grdColumns(grdGLTBRECT, "T_LNO", True)

        Bind_Controls(grpStatement, "GLTBREC1")
        'Set_Read_Only_for_ctl(Absx1.numFor("INV_TOTAL_AMOUNT"), True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("BANK_CODE")

                If EMsg = "" Then
                    BANK_CODE = Absx1.txtFor("BANK_CODE").Text
                    'Don't let them start a new bank rec for a subordinate bank
                    ASCMAIN1.sql = "SELECT COUNT(*) FROM GLTBANK3 WHERE BANK_CODE_SUB = '" & BANK_CODE & "'"
                    Dim count As Integer = Val(ASCDATA1.GetDataValue)

                    If count > 0 Then
                        ASCMAIN1.sql = "SELECT BANK_CODE FROM GLTBANK3 WHERE BANK_CODE_SUB = '" & BANK_CODE & "'"
                        Dim correctBankCode As String = ASCDATA1.GetDataValue
                        EMsg &= "The entered Bank Code is listed as a subordinate bank." & vbCr &
                            "Please use the main bank code '" & correctBankCode & "' instead."
                    Else
                        ASCMAIN1.sql = "SELECT BANK_CODE_SUB FROM GLTBANK3 WHERE BANK_CODE = '" & BANK_CODE & "'"
                        Dim dtBankCodes As DataTable = ASCDATA1.GetDataTable()
                        If dtBankCodes.Rows.Count > 0 Then
                            Dim BANK_CODES As New List(Of String)
                            BANK_CODES.Add("'" & BANK_CODE & "'")
                            For Each rowBankCode As DataRow In dtBankCodes.Rows
                                BANK_CODES.Add("'" & rowBankCode("BANK_CODE_SUB").ToString() & "'")
                            Next
                            BANK_CODE_FILTER = "IN (" & String.Join(", ", BANK_CODES) & ")"
                        Else
                            BANK_CODE_FILTER = "= '" & BANK_CODE & "'"
                        End If
                    End If

                    Dim row As DataRow = dst.Tables("GLTBANK1").Rows.Find(BANK_CODE)
                    If row IsNot Nothing AndAlso Not IsDBNull(row("BANK_API_IND")) Then
                        IS_API_BANK = True
                    End If

                    If Not ASCMAIN1.Logical_Lock("GLTBREC1", BANK_CODE) Then Exit Sub

                    ASCMAIN1.sql = "Select Max (OPS_YYYYPP) from GLTBREC1 where BANK_CODE = '" & BANK_CODE & "'"
                    OPS_YYYYPP = ASCDATA1.GetDataValue
                    Dim LYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)

                    If OPS_YYYYPP = "" Then
                        BANK_STMT_BALANCE_previous = 0
                        OPS_YYYYPP = LYP
                    Else
                        If OPS_YYYYPP >= LYP Then
                            EMsg &= vbCr & "The last Bank Reconciliation for Bank " & BANK_CODE & " was for Period " & OPS_YYYYPP _
                                & vbCr & "The next Bank Reconciliation will be for the Current Period: " & ASCMAIN1.CYP _
                                & vbCr & "Cannot start a new Bank Reconcilation until the current Period is Closed"
                        Else
                            ASCMAIN1.sql = "Select BANK_STMT_BALANCE from GLTBREC1" _
                                 & " where BANK_CODE = '" & BANK_CODE & "' and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                            BANK_STMT_BALANCE_previous = Val(ASCDATA1.GetDataValue)
                            OPS_YYYYPP = ASCMAIN1.Period_Calc(OPS_YYYYPP, 1)
                        End If
                    End If
                End If

                If EMsg <> "" Then ASCMAIN1.MultiTask_Release()

            Case "View", "Edit"

                If Absx1.txtFor("BATCH_NO_CLEARED").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Batch No"
                Else
                    BATCH_NO_CLEARED = Absx1.txtFor("BATCH_NO_CLEARED").Text
                    ASCMAIN1.sql = "Select * from GLTBREC1 where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
                    Dim rowGLTBREC1 As DataRow = ASCDATA1.GetDataRow
                    If rowGLTBREC1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Batch No"
                    Else
                        BANK_CODE = rowGLTBREC1.Item("BANK_CODE")

                        If eItemKey = "Edit" Then

                            If rowGLTBREC1.Item("FINALIZED") = 1 Then
                                EMsg &= vbCr & "You cannot edit a finalized bank reconciliation."
                            End If

                            If Not ASCMAIN1.Logical_Lock("GLTBREC1", BANK_CODE) Then Exit Sub
                            ASCMAIN1.sql = "Select Max (OPS_YYYYPP) from GLTBREC1 where BANK_CODE = '" & BANK_CODE & "'"
                            OPS_YYYYPP = ASCDATA1.GetDataValue
                            If OPS_YYYYPP <> rowGLTBREC1.Item("OPS_YYYYPP") Then
                                EMsg &= vbCr & "Cannot Edit this Bank Reconciliation" _
                                    & vbCr & "- Last Bank Reconciliation for Bank " & BANK_CODE & " was for Period: " & OPS_YYYYPP
                            End If
                        End If
                    End If
                End If

                Dim row As DataRow = dst.Tables("GLTBANK1").Rows.Find(BANK_CODE)
                If row IsNot Nothing AndAlso Not IsDBNull(row("BANK_API_IND")) Then
                    IS_API_BANK = True
                End If

                If EMsg <> "" And eItemKey = "Edit" Then ASCMAIN1.MultiTask_Release()

            Case "Update"

                Dim DT As Date = Absx1.dteFor("BANK_STMT_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Statement Date is Mandatory"
                End If

                'Dim ALL_MATCHED As Boolean = Check_if_Matches_Are_in_Balance("Update")
                'If Not ALL_MATCHED Then
                '    'EMsg &= "You still have unmatched transactions on the API Transactions tab. This Bank Rec cannot be finalized until everything is matched." & vbCrLf
                '    'MsgBox("This Bank Rec will not be finalized as there are still unmatched API Transactions", MsgBoxStyle.OkOnly, "Warning")
                'End If

                Dim totalGLTCHTR1 As Decimal = Val(dst.Tables("GLTCHTR1").Compute("SUM(MATCHED_AMT)", "MATCHED = '1'") & "")
                Dim totalGLTBREC2 As Decimal = Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT)", "TRAN_SEL = '1'") & "")

                If totalGLTCHTR1 <> totalGLTBREC2 Then
                    Dim formattedTotalGLTCHTR1 As String = totalGLTCHTR1.ToString("C2")
                    Dim formattedTotalGLTBREC2 As String = totalGLTBREC2.ToString("C2")
                    EMsg &= $"The total TRAN_AMT for the bottom grid ({formattedTotalGLTBREC2}) does not match the total MATCHED_AMT for the top grid ({formattedTotalGLTCHTR1})." & vbCrLf
                End If

            Case "Delete"

                If EMsg = "" Then
                    If MsgBox("Do you really want to Delete this Bank Reconciliation",
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If


            Case "Save Reconcile"
                Dim ALL_MATCHED As Boolean = Check_if_Matches_Are_in_Balance("Save Reconcile")
                If Not ALL_MATCHED Then
                    EMsg &= $"The total matched for the top grid does not equal the total matched for the bottom grid." & vbCrLf
                End If

            Case "Cancel Reconcile"

            Case "Clear Reconciled"
                Dim ALL_MATCHED As Boolean = Check_if_Matches_Are_in_Balance("Clear Reconciled")
                If Not ALL_MATCHED Then
                    EMsg &= $"The total RECONCILE_AMT for the top grid does not match the total for the bottom grid." & vbCrLf
                End If

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Async Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                XNOs.Clear()
                BATCH_NO_CLEARED = ASCMAIN1.Next_Control_No("GLTBREC1.BATCH_NO_CLEARED")
                If IS_API_BANK Then
                    'IF THERE ARE RECORDS WITH THIS OPS_YYYYPP, WE DONT NEED TO CALL THE API
                    Dim existingXNOs As DataTable = ASCDATA1.GetDataTable("SELECT DISTINCT XNO FROM GLTCHBL1 WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "'")

                    If existingXNOs.Rows.Count > 0 Then
                        For Each row As DataRow In existingXNOs.Rows
                            XNOs.Add(row("XNO").ToString())
                        Next
                    Else
                        Dim accountIDs As List(Of String) = Get_Account_ID(BANK_CODE)

                        ASCMAIN1.Progress("Retrieving authentication tokens...")
                        Threading.Tasks.Task.Run(Function() Get_Tokens()).Wait()

                        For Each accountID As String In accountIDs
                            ASCMAIN1.Progress($"Fetching account balance for Account {accountID}...")
                            Threading.Tasks.Task.Run(Function() Get_Bal(accountID)).Wait()

                            ASCMAIN1.Progress($"Fetching transactions for Account {accountID}...")
                            Threading.Tasks.Task.Run(Function() Get_Trans(accountID)).Wait()

                        Next
                    End If
                    tabs.Tabs(0).Visible = False
                    tabs.Tabs(1).Visible = True
                    tabs.Tabs(2).Visible = True
                    tabs.SelectedTab = tabs.Tabs(1)

                    Dim TRAN_SEL As UltraWinGrid.UltraGridColumn = grdGLTBREC2.DisplayLayout.Bands(0).Columns("TRAN_SEL")
                    TRAN_SEL.CellActivation = Activation.NoEdit
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                tabs.Tabs(0).Visible = False
                tabs.Tabs(1).Visible = True
                tabs.Tabs(2).Visible = True
                tabs.SelectedTab = tabs.Tabs(1)
                Load_Record()
                Mode_Settings(True)


            Case "View"
                EntryMode = "V"
                tabs.Tabs(0).Visible = False
                tabs.Tabs(1).Visible = True
                tabs.Tabs(2).Visible = True
                tabs.SelectedTab = tabs.Tabs(1)
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Print"
                Export_OS_To_Excel()
                Print_Record()

            Case "Match Transactions"
                Match_Transactions()


            Case "Reconcile"
                Reconcile_Mode()

            Case "Save Reconcile"
                Save_Reconcile()

            Case "Clear Reconciled"
                Clear_Reconcile()

            Case "Cancel Reconcile"
                Cancel_Reconcile()

                ASCMAIN1.Progress("")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)
        Dim isFinalized As Boolean = False
        Dim isApproved As Boolean = False
        Dim finalizedByUser As String = ""

        If rowGLTBREC1 IsNot Nothing Then
            If Not IsDBNull(rowGLTBREC1("FINALIZED")) Then
                isFinalized = (rowGLTBREC1("FINALIZED").ToString() = "1")
            End If
            If Not IsDBNull(rowGLTBREC1("APPROVED")) Then
                isApproved = (rowGLTBREC1("APPROVED").ToString() = "1")
            End If
            finalizedByUser = rowGLTBREC1("FINALIZED_USER") & ""
        End If
        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" And EntryMode <> "E" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = DefaultableBoolean.False
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                        .Items("Edit").Settings.Enabled = iScreenMode
                    End If

                    .Items("Delete").Visible = ScreenMode And EntryMode = "E"
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Match Transactions").Visible = ScreenMode And IS_API_BANK And EntryMode = "N"
                    .Items("Match Transactions").Settings.Enabled = Not ScreenMode
                End With

                .Groups("Statement").Visible = ScreenMode
                .Groups("Reconciliation").Visible = ScreenMode
                .Groups("Reconcile").Visible = ScreenMode And IS_API_BANK And Not isFinalized And (EntryMode = "N" Or EntryMode = "E")

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If EntryMode = "N" Or EntryMode = "E" Then
            grdGLTBREC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Set_Read_Only(grpStatement, False)
        Else
            grdGLTBREC2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Set_Read_Only(grpStatement, True)
            Set_Read_Only_for_ctl(chkShowFuture, False)
        End If

        tabs.Tabs(3).Visible = isFinalized
        Dim currentUser As String = ASCMAIN1.USER_ID
        btnApprove.Visible = isFinalized And ScreenMode And Not isApproved AndAlso Not (finalizedByUser = currentUser)

        grpDetails.Visible = tf
        grdGLTBREC2.Visible = tf
        grdGLTBRECX.Visible = Not tf

        Set_Read_Only(grpDetails, (EntryMode = "V"))

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLTBREC1", "GLTBREC2", "GLTBRECX", "GLTOSREC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdGLTBREC2.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Refresh_Documents()
        tabs.Tabs(0).Visible = True
        tabs.Tabs(1).Visible = False
        tabs.Tabs(2).Visible = False
        tabs.Tabs(3).Visible = False

        tabs.SelectedTab = tabs.Tabs(0)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowGLTBREC1 = dst.Tables("GLTBREC1").NewRow

            rowGLTBREC1.Item("BATCH_NO_CLEARED") = BATCH_NO_CLEARED
            rowGLTBREC1.Item("BANK_CODE") = BANK_CODE
            rowGLTBREC1.Item("OPS_YYYYPP") = OPS_YYYYPP
            rowGLTBREC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowGLTBREC1.Item("INIT_DATE") = DATETIME_STAMP

            rowGLTPARM2 = LookUp("GLTPARM2", OPS_YYYYPP)
            Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")
            rowGLTBREC1.Item("BANK_STMT_DATE") = PRD_END_DATE

            dst.Tables("GLTBREC1").Rows.Add(rowGLTBREC1)
            Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", rowGLTBREC1.Item("BANK_CODE") & "")
        Else
            rowGLTBREC1 = Fill_Record("GLTBREC1", New String() {Absx1.txtFor("BATCH_NO_CLEARED").Text})
            OPS_YYYYPP = rowGLTBREC1.Item("OPS_YYYYPP")
            rowGLTPARM2 = LookUp("GLTPARM2", OPS_YYYYPP)
        End If

        ASCMAIN1.sql = "Select BANK_STMT_BALANCE from GLTBREC1" _
                              & " where BANK_CODE = '" & BANK_CODE & "' and OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(OPS_YYYYPP, -1) & "'"
        BANK_STMT_BALANCE_previous = Val(ASCDATA1.GetDataValue)


        ASCMAIN1.sql = "Select NVL(BANK_AMT_OS_APCD,0) + NVL(BANK_AMT_OS_ARCR,0) + NVL(BANK_AMT_OS_GLJE,0) BANK_AMT_OS_APCD_LM from GLTBREC1" _
                              & " where BANK_CODE = '" & BANK_CODE & "' and OPS_YYYYPP = '" & ASCMAIN1.Period_Calc(OPS_YYYYPP, -1) & "'"
        BANK_AMT_OS_APCD_LM = Val(ASCDATA1.GetDataValue)


        rowGLTBANK1 = LookUp("GLTBANK1", BANK_CODE)
        UltraExplorerBar1.Groups("Statement").Text = "Statement " & rowGLTPARM2.Item("LEGEND")

        Dim BANK_BOOK_BALANCE_LY = 0
        Dim sqlXX As String = ""
        If Mid(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 1, 4) < Mid(OPS_YYYYPP, 1, 4) Then
            ' note - the compensating code below assumes that the current GL period is no more that 1 year prior to the bank rec YP
            sqlXX = GetPeriods(12)
            ASCMAIN1.sql = $"Select NVL(ACCT_BEG_BAL,0) {sqlXX} BANK_BOOK_BALANCE" & vbCrLf _
                & $" from GLTACCT3" & vbCrLf _
                & $" where ACCT_CODE = '{rowGLTBANK1.Item("ACCT_CODE")}'" & vbCrLf _
                & $"   and ACCT_YEAR = '{Format(Val(Mid(OPS_YYYYPP, 1, 4)) - 1, "0000")}'"
            ASCMAIN1.sql = $"Select Sum (BANK_BOOK_BALANCE) BANK_BOOK_BALANCE from ({ASCMAIN1.sql})"
            BANK_BOOK_BALANCE_LY = Val(ASCDATA1.GetDataValue() & "")
        End If
        sqlXX = GetPeriods(Val(Mid(OPS_YYYYPP, 5, 2)))
        ASCMAIN1.sql = $"Select NVL(ACCT_BEG_BAL,0) {sqlXX} BANK_BOOK_BALANCE" & vbCrLf _
                & $" from GLTACCT3" & vbCrLf _
                & $" where ACCT_CODE = '{rowGLTBANK1.Item("ACCT_CODE")}'" & vbCrLf _
                & $"   and ACCT_YEAR = '{Mid(OPS_YYYYPP, 1, 4)}'"
        ASCMAIN1.sql = $"Select Sum (BANK_BOOK_BALANCE) BANK_BOOK_BALANCE from ({ASCMAIN1.sql})"
        ASCDATA1.ExecuteSQL()
        BANK_BOOK_BALANCE = Val(ASCDATA1.GetDataValue() & "") + BANK_BOOK_BALANCE_LY

        ASCDATA1.ExecuteSQL("Truncate Table " & GLTBREC2)
        ASCMAIN1.sql = "Insert into " & GLTBREC2 _
            & " Select GLTBREC2.* from GLTBREC2" _
            & " where GLTBREC2.BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
        ASCDATA1.ExecuteSQL()

        If EntryMode = "N" Then

            'Dim bankCodeFilter As String = If(BANK_CODE = "CHASE", "in ('CHASE', 'WIRE')", "= '" & BANK_CODE & "'")
            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'APCD' JOURNAL_TYPE" & vbCrLf _
                       & ", OPS_YYYYPP TRAN_YP, CHECK_NUM TRAN_KEY, 0 TRAN_KEY_LNO" & vbCrLf _
                       & ", CHECK_DATE TRAN_DATE, -1 * CHECK_AMT TRAN_AMT, VEND_NAME TRAN_DESC, '0' TRAN_SEL," & vbCrLf _
                       & " VEND_CODE CUST_VEND, NULL CUST_REF_NO, BANK_CODE, '0' RECONCILE, NULL MATCHED_AMT, NULL RECONCILE_AMT" & vbCrLf _
                       & " from APTCHCK1 where BANK_CODE " & BANK_CODE_FILTER & " and BATCH_NO_CLEARED is Null" & vbCrLf _
                       & " and OPS_YYYYPP <= '" & OPS_YYYYPP & "'"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'APCD' JOURNAL_TYPE" & vbCrLf _
                       & ", OPS_YYYYPP_F TRAN_YP, CHECK_NUM TRAN_KEY, 1 TRAN_KEY_LNO" & vbCrLf _
                       & ", CHECK_DATE TRAN_DATE, CHECK_AMT TRAN_AMT, VEND_NAME TRAN_DESC, '0' TRAN_SEL," & vbCrLf _
                       & " VEND_CODE CUST_VEND, NULL CUST_REF_NO, BANK_CODE, '0' RECONCILE, NULL MATCHED_AMT, NULL RECONCILE_AMT" & vbCrLf _
                       & " from APTCHCK1 where BANK_CODE " & BANK_CODE_FILTER & " and BATCH_NO_CLEARED_F is Null" & vbCrLf _
                       & " and CHECK_STATUS = 'V'" & vbCrLf _
                       & " and OPS_YYYYPP_F <= '" & OPS_YYYYPP & "'"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)

            ' ARCR CHECKS
            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'ARCR' JOURNAL_TYPE" & vbCrLf _
                       & ", ARTPYMT1.OPS_YYYYPP TRAN_YP, ARTPYMT1.PYMT_BATCH_NO TRAN_KEY, ARTPYMT2.PYMT_BATCH_LNO TRAN_KEY_LNO" & vbCrLf _
                       & ", ARTPYMT1.PYMT_BATCH_DATE TRAN_DATE, ARTPYMT2.CUST_PYMT_AMT TRAN_AMT," & vbCrLf _
                       & " 'Total:' || TO_CHAR(SUM(ARTPYMT2.CUST_PYMT_AMT) OVER (PARTITION BY ARTPYMT1.PYMT_BATCH_NO), '99,999,999.99') TRAN_DESC, '0' TRAN_SEL," & vbCrLf _
                       & " ARTPYMT2.CUST_CODE CUST_VEND, ARTPYMT2.CUST_PYMT_REF_NO CUST_REF_NO, ARTPYMT1.BANK_CODE, '0' RECONCILE, NULL MATCHED_AMT, NULL RECONCILE_AMT " & vbCrLf _
                       & " from ARTPYMT1, ARTPYMT2 where ARTPYMT1.BANK_CODE " & BANK_CODE_FILTER & " and ARTPYMT1.BATCH_NO_CLEARED is Null" & vbCrLf _
                       & " and ARTPYMT1.OPS_YYYYPP <= '" & OPS_YYYYPP & "'" & vbCrLf _
                       & " and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                       & " and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" & vbCrLf _
                       & " group by ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO, ARTPYMT1.PYMT_BATCH_DATE, ARTPYMT2.CUST_PYMT_REF_NO, ARTPYMT2.CUST_PYMT_AMT, ARTPYMT2.CUST_CODE, ARTPYMT2.CUST_NAME, ARTPYMT1.BANK_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)


        End If

        If EntryMode = "N" Or EntryMode = "E" Then

            ASCMAIN1.sql = "Select OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO, ACCT_CODE, DETL_POSTING_AMT, DETL_CTL_DATE" & vbCrLf _
                & " from GLTDETL1 where (OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO) in (" & vbCrLf _
                & "Select GLTDETL1.OPS_YYYYPP, GLTDETL1.JOURNAL_NO, GLTDETL1.JOURNAL_LNO" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1" & vbCrLf _
                & " where GLTDETL1.ACCT_CODE = '" & rowGLTBANK1.Item("ACCT_CODE") & "'" & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO " & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_TYPE = 'GLJE'" & vbCrLf _
                & " minus " & vbCrLf _
                & "Select OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO from GLTBREC0" & vbCrLf _
                & " where ACCT_CODE = '" & rowGLTBANK1.Item("ACCT_CODE") & "'" & vbCrLf _
                & ")"
            ASCDATA1.ExecuteSQL("Insert into GLTBREC0 (OPS_YYYYPP, JOURNAL_NO, JOURNAL_LNO, ACCT_CODE, DETL_POSTING_AMT, DETL_CTL_DATE) " & ASCMAIN1.sql)

            ASCMAIN1.sql = "Delete from " & GLTBREC2 & " where JOURNAL_TYPE = 'GLJE' and NVL(TRAN_SEL,'0') <> '1'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select '" & BATCH_NO_CLEARED & "' BATCH_NO_CLEARED, 'GLJE' JOURNAL_TYPE" & vbCrLf _
                 & ", GLTBREC0.OPS_YYYYPP TRAN_YP, GLTBREC0.JOURNAL_NO TRAN_KEY, GLTBREC0.JOURNAL_LNO TRAN_KEY_LNO" & vbCrLf _
                 & ", GLTBREC0.DETL_CTL_DATE TRAN_DATE, GLTBREC0.DETL_POSTING_AMT TRAN_AMT, GLTJRNL1.JOURNAL_DESC TRAN_DESC, '0' TRAN_SEL," & vbCrLf _
                 & " NULL CUST_VEND, NULL CUST_REF_NO, NULL BANK_CODE, '0' RECONCILE, NULL MATCHED_AMT, NULL RECONCILE_AMT" & vbCrLf _
                 & " from GLTBREC0,GLTJRNL1 where GLTBREC0.ACCT_CODE = '" & rowGLTBANK1.Item("ACCT_CODE") & "' and GLTBREC0.BATCH_NO_CLEARED is Null" & vbCrLf _
                 & " and GLTBREC0.OPS_YYYYPP <= '" & OPS_YYYYPP & "'" & vbCrLf _
                 & " and GLTJRNL1.JOURNAL_NO = GLTBREC0.JOURNAL_NO"
            ASCDATA1.ExecuteSQL("Insert into " & GLTBREC2 & " " & ASCMAIN1.sql)



        End If

        Fill_Records("GLTBREC2")
        Set_RowFilter()
        Sort_grdColumns(grdGLTBREC2, "TRAN_DATE,JOURNAL_TYPE,TRAN_KEY,TRAN_KEY_LNO")

        If EntryMode = "N" Then
            Dim xnoQuery As String = String.Join("', '", XNOs)
            ASCMAIN1.sql = "Select * from GLTCHTR1 WHERE XNO IN ('" & xnoQuery & "') OR MATCHED IS NULL OR MATCHED <> '1'" ' = '0000000082 WIRE OCT 24, 0000000083 CHASE OCT 24
            Fill_Records("GLTCHTR1",,, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from GLTCHBL2 WHERE XNO IN ('" & xnoQuery & "')" '= '0000000082 WIRE OCT 24, 0000000083 CHASE OCT 24 'CHASE, CHASE IS 0 BALANCE SO DO I GET BAL FROM 77?
            Fill_Records("GLTCHBL2",,, ASCMAIN1.sql)

        Else
            ASCMAIN1.sql = "Select * from GLTCHTR1 WHERE BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
            Fill_Records("GLTCHTR1",,, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from GLTCHBL2 WHERE BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
            Fill_Records("GLTCHBL2",,, ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT JOURNAL_TYPE, TRAN_YP, TRAN_KEY, CUST_VEND, TRAN_DATE, TRAN_DESC, TRAN_AMT " &
                           "FROM GLTBREC2 WHERE TRAN_SEL <> '1' AND JOURNAL_TYPE = 'APCD' AND BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
            Fill_Records("GLTOSREC",,, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from GLTBRECR WHERE BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
            Fill_Records("GLTBRECR",,, ASCMAIN1.sql)
        End If

        Display_Totals()
        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Dim totalAPCD_OS As Decimal = Math.Abs(Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT)", "TRAN_SEL <> '1' AND JOURNAL_TYPE = 'APCD'") & ""))
        Dim totalARCR_OS As Decimal = Math.Abs(Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT)", "TRAN_SEL <> '1' AND JOURNAL_TYPE = 'ARCR'") & ""))
        Dim totalGLJE_OS As Decimal = Math.Abs(Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT)", "TRAN_SEL <> '1' AND JOURNAL_TYPE = 'GLJE'") & ""))

        ASCDATA1.ExecuteSQL("Truncate Table " & GLTBREC2)

        Dim API_OK As Boolean = Can_Finalize()
        BeginTrans()
        'If Check_if_Matches_Are_in_Balance("Update") Then
        '    rowGLTBREC1("FINALIZED") = 1
        '    rowGLTBREC1("FINALIZED_DATE") = DATETIME_STAMP
        '    rowGLTBREC1("FINALIZED_USER") = ASCMAIN1.USER_ID
        'Else
        '    Dim result As DialogResult = MessageBox.Show("Not all API Transactions have been matched. Would you like to finalize?", "Finalize Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)

        '    If result = DialogResult.Yes Then
        '        If API_OK Then
        '            rowGLTBREC1("FINALIZED") = 1
        '            rowGLTBREC1("FINALIZED_DATE") = DATETIME_STAMP
        '            rowGLTBREC1("FINALIZED_USER") = ASCMAIN1.USER_ID
        '        Else
        '            MsgBox("There are still unmatched API transactions that are NOT marked as reconciling items. You cannot finalize until this is resolved.", MsgBoxStyle.Exclamation, "Cannot Finalize")
        '            rowGLTBREC1("FINALIZED") = 0
        '        End If
        '    Else
        '        rowGLTBREC1("FINALIZED") = 0
        '    End If
        'End If
        Dim canFinalizeByMatch As Boolean = Check_if_Matches_Are_in_Balance("Update")
        Dim wantsFinalize As Boolean = False

        If canFinalizeByMatch Then
            wantsFinalize = True
        Else
            Dim result As DialogResult = MessageBox.Show("Not all API Transactions have been matched. Would you like to finalize?", "Finalize Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            wantsFinalize = (result = DialogResult.Yes)
        End If

        rowGLTBREC1("FINALIZED") = 0

        If wantsFinalize Then
            If Not canFinalizeByMatch AndAlso Not API_OK Then
                MsgBox("There are still unmatched API transactions that are NOT marked as reconciling items. You cannot finalize until this is resolved.", MsgBoxStyle.Exclamation, "Cannot Finalize")
            Else
                Dim proofAmt As Decimal = GetProofAmountFromData()
                If Math.Round(proofAmt, 2) <> 0D Then
                    MsgBox($"Proof is nonzero ({proofAmt.ToString("C2")}). You cannot finalize until Proof equals 0.", MsgBoxStyle.Exclamation, "Cannot Finalize")
                Else
                    ' All good — finalize
                    rowGLTBREC1("FINALIZED") = 1
                    rowGLTBREC1("FINALIZED_DATE") = DATETIME_STAMP
                    rowGLTBREC1("FINALIZED_USER") = ASCMAIN1.USER_ID
                End If
            End If
        End If

        rowGLTBREC1("BANK_AMT_OS_APCD") = totalAPCD_OS
        rowGLTBREC1("BANK_AMT_OS_ARCR") = totalARCR_OS
        rowGLTBREC1("BANK_AMT_OS_GLJE") = totalGLJE_OS

        If EntryMode = "E" Then
            Dependent_Updates(-1)
        End If

        Update_Record_TDA("GLTBREC1")
        Update_Record_TDA("GLTBREC2", "1=1")

        For Each rowGLTCHTR1 As DataRow In dst.Tables("GLTCHTR1").Rows
            Dim cur = If(rowGLTCHTR1("BATCH_NO_CLEARED") Is DBNull.Value, "", CStr(rowGLTCHTR1("BATCH_NO_CLEARED")))
            If cur <> BATCH_NO_CLEARED Then
                rowGLTCHTR1("BATCH_NO_CLEARED") = BATCH_NO_CLEARED
            End If
        Next

        Update_Record_TDA("GLTCHTR1")
        Update_Record_TDA("GLTBRECR")

        ASCDATA1.ExecuteSQL("Delete from GLTBREC2 where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'")
        ASCDATA1.ExecuteSQL("Insert into GLTBREC2 Select * from " & GLTBREC2)
        'ASCDATA1.ExecuteSQL("UPDATE GLTCHTR1 SET BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' WHERE BATCH_NO_CLEARED IS NULL")

        Dependent_Updates(1)

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Dependent_Updates(-1)
        Delete_Records("GLTBREC1")
        Delete_Records("GLTBREC2")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'")
    End Sub

    Sub Print_Record()

        Print_Report_Begin()
        ' CR_params.Add("SUBT", "")

        Dim RPT As String = "GLRBREC1"
        Generate_Report(RPT, Me.Text, "Cleared Items", "{GLTBREC2.TRAN_SEL}='1'", , , False)
        'Generate_Report(RPT, Me.Text, "Open Items", "{GLTBREC2.TRAN_SEL}='0' and {GLTBREC2.TRAN_DATE} <= {GLTBREC1.BANK_STMT_DATE}", , , False)

        'RPT = "GLRBREC2"
        Generate_Report(RPT, Me.Text, "Outstanding Checks & Reconciling Items", "{GLTBREC2.TRAN_SEL}='0' and {GLTBREC2.TRAN_DATE} <= {GLTBREC1.BANK_STMT_DATE}", , , False)

        RPT = "GLRBREC3"
        Dim YEAR As String = OPS_YYYYPP.Substring(0, 4) ' Extracts "2025"
        Dim PERIOD As String = OPS_YYYYPP.Substring(4, 2) ' Extracts "01"
        Dim MONTH As String = MonthName(CInt(PERIOD))
        Dim SUBT As String = $"For the Month of {MONTH} {YEAR}"

        Generate_Report(RPT, Me.Text, SUBT, "", , , False)

        Print_Report_End()

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    ' Click_Command("New", e)
                    Load_GLTBRECX()
                End If
            Case "BATCH_NO_CLEARED"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BANK_CODE"
                '   Click_Command("New")
                Load_GLTBRECX()
            Case "BATCH_NO_CLEARED"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "BANK_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BANK_CODE").Text <> "" Then
                        LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                        If cdr IsNot Nothing Then
                            Load_GLTBRECX()
                        End If
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_STMT_DATE"
                Set_RowFilter()
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_STMT_BALANCE"
                Display_Totals()
        End Select
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTBRECX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdGLTBREC2, "SSSBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select Selected", "De-Select Selected", "Find Possible Matches", "Clear Offsetting Amounts")
        Load_Popup_Menu(grdGLTCHTR1, "SSSBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select Selected", "De-Select Selected", "Select All", "De-Select All", "Add Reconciling Item", "Remove Reconciling Item")
        Load_Popup_Menu(grdGLTMATCH, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Journal Inquiry", "Select Selected", "De-Select Selected", "Select All", "De-Select All")
        Load_Popup_Menu(grdGLTOSREC, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If EntryMode = "V" Or Not ScreenMode Then
            e.Cancel = True
            Exit Sub
        End If

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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdGLTBREC2"
                tlb_btn = DirectCast(tlb_pop.Tools("Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

                Dim offsetting_Rows As Boolean = Selected_Rows_Offset(grd)
                tlb_btn = DirectCast(tlb_pop.Tools("Clear Offsetting Amounts"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = offsetting_Rows

            Case "grdGLTMATCH"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    Dim JOURNAL_TYPE As String = grd.ActiveRow.Cells("JOURNAL_TYPE").Text
                    tlb_btn = DirectCast(tlb_pop.Tools("Journal Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (JOURNAL_TYPE = "GLJE")
                End If

            Case "grdGLTCHTR1"
                Dim singleRowSelected As Boolean = (grd.Selected.Rows.Count = 1 AndAlso grd.Selected.Rows(0).IsDataRow)

                tlb_btn = DirectCast(tlb_pop.Tools("Add Reconciling Item"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = singleRowSelected

                tlb_btn = DirectCast(tlb_pop.Tools("Remove Reconciling Item"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = singleRowSelected

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select Selected", "De-Select Selected"

                Dim COL As String = "TRAN_SEL"
                If grd.Name = "grdGLTCHTR1" Or grd.Name = "grdGLTMATCH" Then
                    COL = "RECONCILE"
                End If

                For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                    grow.Cells(COL).Value = IIf(e.Tool.Key = "De-Select Selected", "0", "1")
                    grow.Update()
                Next

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        grow.Cells("RECONCILE").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Clear Offsetting Amounts"
                If Selected_Rows_Offset(grd) Then
                    For Each row As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        row.Cells("TRAN_SEL").Value = 1
                        row.Update()
                    Next
                End If

            Case "Find Possible Matches"
                Find_Possible_Matches()

            Case "Journal Inquiry"
                Dim JOURNAL_TYPE As String = grd.ActiveRow.Cells("JOURNAL_TYPE").Value

                If JOURNAL_TYPE = "GLJE" Then
                    Dim JOURNAL_NO As String = grd.ActiveRow.Cells("TRAN_KEY").Value
                    Dim rowGLTJRNL1 As DataRow = LookUp("GLTJRNL1", JOURNAL_NO)
                    If rowGLTJRNL1 IsNot Nothing Then
                        Context_Launch("View", JOURNAL_NO, "Journal Inquiry", "GLFJRNL1")
                    End If
                End If
                'Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                'If rowSOTORDR1 IsNot Nothing Then
                '    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                'End If
            Case "Add Reconciling Item"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    Dim XNO As String = grd.ActiveRow.Cells("XNO").Text
                    Dim XNO_LNO As Integer = Val(grd.ActiveRow.Cells("XNO_LNO").Value & "")
                    Dim XNO_DTL_LNO As Integer = Val(grd.ActiveRow.Cells("XNO_DTL_LNO").Value & "")
                    Dim AMOUNT As Decimal = Val(grd.ActiveRow.Cells("AMOUNT").Value & "")
                    Dim REASON As String = InputBox("Enter a reason for marking this row as a reconciling item:", "Reconciling Item Reason")

                    If String.IsNullOrWhiteSpace(REASON) Then
                        MsgBox("A reason is required to mark this as a reconciling item.", MsgBoxStyle.Exclamation)
                        Exit Sub
                    End If

                    Dim rowGLTCHTR1 As DataRow = dst.Tables("GLTCHTR1").Rows.Find(New Object() {XNO, XNO_LNO, XNO_DTL_LNO})
                    rowGLTCHTR1("RECONCILING_ITEM_IND") = "1"
                    rowGLTCHTR1("RECONCILING_ITEM_REASON") = REASON

                    Dim rowGLTBRECR As DataRow = dst.Tables("GLTBRECR").Rows.Find(New Object() {BATCH_NO_CLEARED, XNO_DTL_LNO})
                    If rowGLTBRECR IsNot Nothing Then
                        rowGLTBRECR("RECONCILING_ITEM_IND") = "1"
                        rowGLTBRECR("RECONCILING_ITEM_REASON") = REASON
                        rowGLTBRECR("AMOUNT") = AMOUNT
                    Else
                        rowGLTBRECR = dst.Tables("GLTBRECR").NewRow()
                        rowGLTBRECR("BATCH_NO_CLEARED") = BATCH_NO_CLEARED
                        rowGLTBRECR("XNO_DTL_LNO") = XNO_DTL_LNO
                        rowGLTBRECR("RECONCILING_ITEM_IND") = "1"
                        rowGLTBRECR("RECONCILING_ITEM_REASON") = REASON
                        rowGLTBRECR("AMOUNT") = AMOUNT
                        dst.Tables("GLTBRECR").Rows.Add(rowGLTBRECR)
                    End If
                    dst.Tables("GLTCHTR1").AcceptChanges()
                    rowGLTCHTR1.SetModified()
                    grdGLTCHTR1.Refresh()

                    MsgBox("Marked as reconciling item.", MsgBoxStyle.Information)
                    Display_Totals()
                End If


            Case "Remove Reconciling Item"
                If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                    Dim XNO As String = grd.ActiveRow.Cells("XNO").Text
                    Dim XNO_LNO As Integer = Val(grd.ActiveRow.Cells("XNO_LNO").Value & "")
                    Dim XNO_DTL_LNO As Integer = Val(grd.ActiveRow.Cells("XNO_DTL_LNO").Value & "")

                    Dim rowGLTCHTR1 As DataRow = dst.Tables("GLTCHTR1").Rows.Find(New Object() {XNO, XNO_LNO, XNO_DTL_LNO})
                    rowGLTCHTR1("RECONCILING_ITEM_IND") = DBNull.Value
                    rowGLTCHTR1("RECONCILING_ITEM_REASON") = DBNull.Value
                    dst.Tables("GLTCHTR1").AcceptChanges()
                    rowGLTCHTR1.SetModified()
                    grdGLTCHTR1.Refresh()

                    Dim rowGLTBRECR As DataRow = dst.Tables("GLTBRECR").Rows.Find(New Object() {BATCH_NO_CLEARED, XNO_DTL_LNO})
                    If rowGLTBRECR IsNot Nothing Then rowGLTBRECR.Delete()

                    MsgBox("Reconciling item cleared.", MsgBoxStyle.Information)
                    Display_Totals()
                End If



        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            '    If rowICTITEM1 IsNot Nothing Then
            '        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

        End Select
    End Sub

#End Region

    Sub Dependent_Updates(S As Integer)
        If S = -1 Then
            Dim sql0 As String = "Set BATCH_NO_CLEARED = NULL where BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "'"
            ASCDATA1.ExecuteSQL("Update APTCHCK1 " & sql0)
            ASCDATA1.ExecuteSQL("Update ARTPYMT1 " & sql0)
            ASCDATA1.ExecuteSQL("Update GLTBREC0 " & sql0)
            ASCDATA1.ExecuteSQL("Update APTCHCK1 " & Replace(sql0, "BATCH_NO_CLEARED", "BATCH_NO_CLEARED_F"))
        Else
            'ASCMAIN1.sql = "SELECT BANK_CODE, BANK_CODE_SUB FROM GLTBANK3 WHERE BANK_CODE = '" & BANK_CODE & "'"
            'Dim DT As DataTable = ASCDATA1.GetDataTable()

            'If DT.Rows.Count > 0 Then
            '    Dim BANK_CODE_MAIN As String = DT.Rows(0)("BANK_CODE") & ""
            '    Dim BANK_CODE_SUB As String = DT.Rows(0)("BANK_CODE_SUB") & ""

            '    BANK_CODE_FILTER = If(BANK_CODE_SUB <> "",
            '              "IN ('" & BANK_CODE_MAIN & "', '" & BANK_CODE_SUB & "')",
            '              "= '" & BANK_CODE_MAIN & "'")
            'End If
            ''Dim sql1 As String = "Set BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' where BANK_CODE = '" & BANK_CODE & "'" 'THIS WOULD ALWAYS SET IT FOR CHASE, EVEN IF WE WANT IT FOR WIRE 
            'Dim sql1 As String = "Set BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' where BANK_CODE " & BANK_CODE_FILTER & " AND BATCH_NO_CLEARED IS NULL" 'dont want to overwrite if it already exists
            'Dim sql2 As String = "in (Select TRAN_KEY from " & GLTBREC2 & " where TRAN_SEL = '1' and JOURNAL_TYPE = "
            'ASCDATA1.ExecuteSQL("Update APTCHCK1 " & sql1 & " and CHECK_NUM " & sql2 & "'APCD' and TRAN_KEY_LNO = 0)")
            'ASCDATA1.ExecuteSQL("Update APTCHCK1 " & Replace(sql1, "BATCH_NO_CLEARED", "BATCH_NO_CLEARED_F") & " and CHECK_NUM " & sql2 & "'APCD' and TRAN_KEY_LNO = 1)")
            'ASCDATA1.ExecuteSQL("Update ARTPYMT1 " & sql1 & " and PYMT_BATCH_NO " & sql2 & "'ARCR')")
            'ASCDATA1.ExecuteSQL("Update GLTBREC0 Set BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' where (OPS_YYYYPP,JOURNAL_NO,JOURNAL_LNO) in (Select TRAN_YP,TRAN_KEY,TRAN_KEY_LNO from " & GLTBREC2 & " where TRAN_SEL = '1' and JOURNAL_TYPE = 'GLJE')")

            ASCDATA1.ExecuteSQL(
            "UPDATE APTCHCK1 a " &
            "   SET BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' " &
            " WHERE a.BATCH_NO_CLEARED IS NULL " &
            "   AND EXISTS ( " &
            "         SELECT 1 FROM " & GLTBREC2 & " g " &
            "          WHERE g.TRAN_SEL = '1' " &
            "            AND g.JOURNAL_TYPE = 'APCD' " &
            "            AND g.TRAN_KEY_LNO = 0 " &
            "            AND g.TRAN_KEY = a.CHECK_NUM " &
            "            AND g.BANK_CODE = a.BANK_CODE " &
            "       )"
        )

            ASCDATA1.ExecuteSQL(
                "UPDATE APTCHCK1 a " &
                "   SET BATCH_NO_CLEARED_F = '" & BATCH_NO_CLEARED & "' " &
                " WHERE a.BATCH_NO_CLEARED_F IS NULL " &
                "   AND EXISTS ( " &
                "         SELECT 1 FROM " & GLTBREC2 & " g " &
                "          WHERE g.TRAN_SEL = '1' " &
                "            AND g.JOURNAL_TYPE = 'APCD' " &
                "            AND g.TRAN_KEY_LNO = 1 " &
                "            AND g.TRAN_KEY = a.CHECK_NUM " &
                "            AND g.BANK_CODE = a.BANK_CODE " &
                "       )"
            )

            ASCDATA1.ExecuteSQL(
                "UPDATE ARTPYMT1 r " &
                "   SET BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' " &
                " WHERE r.BATCH_NO_CLEARED IS NULL " &
                "   AND EXISTS ( " &
                "         SELECT 1 FROM " & GLTBREC2 & " g " &
                "          WHERE g.TRAN_SEL = '1' " &
                "            AND g.JOURNAL_TYPE = 'ARCR' " &
                "            AND g.TRAN_KEY = r.PYMT_BATCH_NO " &
                "            AND g.BANK_CODE = r.BANK_CODE " &
                "       )"
            )

            ASCDATA1.ExecuteSQL(
                "UPDATE GLTBREC0 j " &
                "   SET BATCH_NO_CLEARED = '" & BATCH_NO_CLEARED & "' " &
                " WHERE (j.OPS_YYYYPP, j.JOURNAL_NO, j.JOURNAL_LNO) IN ( " &
                "       SELECT g.TRAN_YP, g.TRAN_KEY, g.TRAN_KEY_LNO " &
                "         FROM " & GLTBREC2 & " g " &
                "        WHERE g.TRAN_SEL = '1' AND g.JOURNAL_TYPE = 'GLJE' " &
                " )"
            )
        End If
    End Sub

    Private Sub grdARTOPENA_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTBREC2.AfterCellUpdate
        'If e.Cell.Column.Key = "TRAN_SEL" Then
        '    If e.Cell.Value = "1" Then
        '        e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = e.Cell.Row.Cells("INV_BALANCE").Value
        '    Else
        '        e.Cell.Row.Cells("INV_BALANCE_APPLIED").Value = 0
        '    End If
        'End If
    End Sub

    Private Sub grdGLTBREC2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGLTBREC2.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdARTOPENA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTBREC2.BeforeRowUpdate
        'If e.Row.Cells("TRAN_SEL").Value = "1" Then
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = e.Row.Cells("INV_BALANCE").Value
        'Else
        '    e.Row.Cells("INV_BALANCE_APPLIED").Value = 0
        'End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor

        'If Absx1.txtFor("BANK_CODE").Text <> "" Then
        Load_GLTBRECX()
        'End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub grdGLTBRECX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTBRECX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("BATCH_NO_CLEARED").Text = e.Row.Cells("BATCH_NO_CLEARED").Text
            Click_Command("View")
        End If
    End Sub

    Sub Load_GLTBRECX()
        Me.Cursor = Cursors.WaitCursor
        Dim BANK_CODE As String = Absx1.txtFor("BANK_CODE").Text
        Fill_Records("GLTBRECX", New String() {BANK_CODE})
        Sort_grdColumns(grdGLTBRECX, "BATCH_NO_CLEARED".ToLower)
        grdGLTBRECX.Text = "Bank Reconciliations" & IIf(BANK_CODE = "", "", " for " & BANK_CODE)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Display_Totals()

        Dim BANK_STMT_BALANCE As Decimal = Val(Absx1.numFor("BANK_STMT_BALANCE").Value & "")
        Dim TRAN_AMT_ARCR As Decimal = Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT_ARCR)", "") & "")
        Dim TRAN_AMT_APCD As Decimal = Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT_APCD)", "") & "")
        Dim TRAN_AMT_GLJE As Decimal = Val(dst.Tables("GLTBREC2").Compute("SUM(TRAN_AMT_GLJE)", "") & "")
        Dim ROLL_FORWARD As Decimal = BANK_STMT_BALANCE_previous + TRAN_AMT_ARCR + TRAN_AMT_APCD + TRAN_AMT_GLJE '+ BANK_AMT_OS_APCD_LM

        Dim RECONCILING_TOTAL As Decimal = 0
        For Each row As DataRow In dst.Tables("GLTCHTR1").Select("RECONCILING_ITEM_IND = '1'")
            RECONCILING_TOTAL += Val(row("AMOUNT") & "")
        Next

        With dst.Tables("GLTBRECT").Rows
            .Find(1).Item("T_AMT") = BANK_STMT_BALANCE_previous
            .Find(2).Item("T_AMT") = TRAN_AMT_ARCR
            .Find(3).Item("T_AMT") = TRAN_AMT_APCD
            .Find(4).Item("T_AMT") = TRAN_AMT_GLJE
            .Find(5).Item("T_AMT") = BANK_AMT_OS_APCD_LM
            .Find(6).Item("T_AMT") = ROLL_FORWARD
            .Find(7).Item("T_AMT") = BANK_STMT_BALANCE
            .Find(8).Item("T_AMT") = RECONCILING_TOTAL
            .Find(9).Item("T_AMT") = BANK_STMT_BALANCE - ROLL_FORWARD + RECONCILING_TOTAL
            .Find(10).Item("T_AMT") = BANK_BOOK_BALANCE
            .Find(11).Item("T_AMT") = BANK_BOOK_BALANCE - BANK_STMT_BALANCE - RECONCILING_TOTAL

        End With
    End Sub

    Private Sub grdGLTBREC2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTBREC2.InitializeRow
        Dim JOURNAL_TYPE As String = e.Row.Cells("JOURNAL_TYPE").Value & ""
        If JOURNAL_TYPE = "ARCR" Then
            e.Row.Cells("JOURNAL_TYPE").Appearance.BackColor = Drawing.Color.LightGreen
        ElseIf JOURNAL_TYPE = "GLJE" Then
            e.Row.Cells("JOURNAL_TYPE").Appearance.BackColor = Drawing.Color.Orange
        ElseIf JOURNAL_TYPE = "APCD" Then
            If Val(e.Row.Cells("TRAN_KEY_LNO").Value & "") = 1 Then ' Void Check
                e.Row.Appearance.ForeColor = Drawing.Color.Red
                e.Row.ToolTipText = "Voided Check"
            End If
        End If
    End Sub

    Private Sub grdGLTBRECT_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTBRECT.InitializeRow
        Dim T_LNO As Integer = Val(e.Row.Cells("T_LNO").Value & "")
        If T_LNO = 1 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGray
        ElseIf T_LNO = 6 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightBlue
        ElseIf T_LNO = 7 Then
            e.Row.Appearance.BackColor = Drawing.Color.LightGreen
        ElseIf T_LNO = 9 Then
            Dim T_AMT As Decimal = Val(e.Row.Cells("T_AMT").Value & "")
            If T_AMT <> 0 Then
                e.Row.Cells("T_AMT").Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Cells("T_AMT").Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If
    End Sub

    Private Sub chkShowFuture_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowFuture.CheckedChanged
        Set_RowFilter()
    End Sub

    Sub Set_RowFilter()
        Dim dvw As DataView = DirectCast(grdGLTBREC2.DataSource, DataTable).DefaultView
        If chkShowFuture.Checked Then
            dvw.RowFilter = ""
        Else
            Dim BANK_STMT_DATE As Date = Absx1.dteFor("BANK_STMT_DATE").Value
            dvw.RowFilter = "TRAN_SEL = '1' or TRAN_DATE <= '" & Format(BANK_STMT_DATE, "MM/dd/yyyy") & "'"
        End If
    End Sub

    Async Function Get_Tokens() As Task
        ' Get JWT
        Dim client As HttpClient = New HttpClient()
        Dim url As String = "https://intapi.interparfums.com/"
        client.BaseAddress = New Uri(url)
        client.DefaultRequestHeaders.Accept.Clear()
        client.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

        Dim env As String = "PROD"

        Dim response As String = Await client.GetStringAsync($"GetJPMC_JWT")
        Dim JWT As String = Mid(response, 2, response.Length - 2)

        'Get OAuth

        Dim clientJPMC As HttpClient = New HttpClient()
        Dim urlJPMC As String = "https://idauatg2.jpmorganchase.com/adfs/oauth2/" ' UAT
        urlJPMC = "https://idag2.jpmorganchase.com/adfs/oauth2/token/" ' PROD
        urlJPMC = "https://idag2.jpmorganchase.com/adfs/oauth2/" ' PROD
        clientJPMC.BaseAddress = New Uri(urlJPMC)
        clientJPMC.DefaultRequestHeaders.Accept.Clear()
        clientJPMC.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("*/*"))

        Dim CLIENT_ID As String = "CC-104221-K050475-289878-UAT" ' UAT
        CLIENT_ID = "CC-104221-A056581-309559-PROD" ' PROD

        Dim reqkv As New List(Of KeyValuePair(Of String, String))
        reqkv.Add(New KeyValuePair(Of String, String)("grant_type", "client_credentials"))
        reqkv.Add(New KeyValuePair(Of String, String)("client_id", CLIENT_ID))
        reqkv.Add(New KeyValuePair(Of String, String)("resource", "https://apigeeproductProd.jpmchase.net"))
        reqkv.Add(New KeyValuePair(Of String, String)("client_assertion_type", "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"))
        reqkv.Add(New KeyValuePair(Of String, String)("client_assertion", JWT))

        Dim q As HttpContent = New FormUrlEncodedContent(reqkv)

        Dim res As HttpResponseMessage = Await clientJPMC.PostAsync("token", q)
        Dim body As String = Await res.Content.ReadAsStringAsync()
        Dim c As OAuthResponse = JsonConvert.DeserializeObject(Of OAuthResponse)(body)
        ACCESS_TOKEN = c.access_token
    End Function

    Async Function Get_Bal(accountID As String) As Task
        'Get Balances
        Dim clientJPMC As HttpClient = New HttpClient()
        Dim urlJPMC_Balances As String = "https://openbanking.jpmorgan.com/accessapi/"
        clientJPMC.BaseAddress = New Uri(urlJPMC_Balances)
        clientJPMC.DefaultRequestHeaders.Accept.Clear()
        clientJPMC.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("*/*"))
        clientJPMC.DefaultRequestHeaders.Add("Authorization", "Bearer " & ACCESS_TOKEN)
        XNO = ASCMAIN1.Next_Control_No("GLTCHBL1.XNO")
        XNOs.Add(XNO)
        'Dim accountId As String = "000000899558985" '000000899558928
        'accountId = "000000899558985" ' CHASE
        'accountId = "000000899558928,000000899558985" ' CHASE

        Dim req2 As New JPMC_Balances_Request2 With {.accountId = accountID}
        Dim req3 As New List(Of JPMC_Balances_Request2)
        req3.Add(req2)

        Dim year As Integer = Integer.Parse(OPS_YYYYPP.Substring(0, 4))
        Dim month As Integer = Integer.Parse(OPS_YYYYPP.Substring(4, 2))

        Dim firstDayOfLastMonth As Date = New Date(year, month, 1)
        Dim lastDayOfLastMonth As Date = firstDayOfLastMonth.AddMonths(1).AddDays(-1)

        Dim startDate As String = firstDayOfLastMonth.ToString("yyyy-MM-dd") 'Format(Now.AddDays(-31), "yyyy-MM-dd") xno 70 = sept 2024, XNO 73 = AUG 2024, xno 75 = july (only has data as of 3 months ago (7/22), no june data, 77 is october up until today
        Dim endDate As String = lastDayOfLastMonth.ToString("yyyy-MM-dd") 'Format(Now.AddDays(-1), "yyyy-MM-dd")
        Dim req1 As New JPMC_Balances_Request With {
        .startDate = startDate,
        .endDate = endDate,
        .accountList = req3
    }

        Dim res = Await clientJPMC.PostAsJsonAsync("balance", req1) 'need nuget Microsoft.AspNet.WebApi.Client

        Dim bodyBalances As String = Await res.Content.ReadAsStringAsync()
        Dim cBalances As JPMC_Balances_Response = JsonConvert.DeserializeObject(Of JPMC_Balances_Response)(bodyBalances)
        Dim cc As New List(Of JPMC_Balances_Response)
        cc.Add(cBalances)
        Dim tbl As New DataTable
        Dim rowGLTCHBL1 As DataRow = dst.Tables("GLTCHBL1").NewRow
        Dim rowGLTCHBL2 As DataRow

        Dim cols As PropertyInfo() = cc(0).accountList(0).GetType().GetProperties
        Dim cols2 As PropertyInfo() = cc(0).accountList(0).currency.GetType().GetProperties
        Dim cols3 As PropertyInfo() = Nothing
        If cc(0).accountList(0).balanceList IsNot Nothing AndAlso cc(0).accountList(0).balanceList.Any() Then
            cols3 = cc(0).accountList(0).balanceList(0).GetType().GetProperties
        End If
        For Each col As PropertyInfo In cols
            If dst.Tables("GLTCHBL1").Columns.Contains(col.Name) Then
                Dim val As Object = col.GetValue(cc(0).accountList(0))
                rowGLTCHBL1.Item(col.Name) = val.ToString
            End If
            If col.Name = "currency" Then
                For Each col2 As PropertyInfo In cols2
                    If dst.Tables("GLTCHBL1").Columns.Contains(col.Name & "_" & col2.Name) Then
                        Dim val As Object = col2.GetValue(cc(0).accountList(0).currency)
                        rowGLTCHBL1.Item(col.Name & "_" & col2.Name) = val
                    End If
                Next
            End If
            If col.Name = "balanceList" AndAlso cols3 IsNot Nothing Then
                XNO_LNO = 0
                For Each balanceList As BalanceList In cc(0).accountList(0).balanceList
                    XNO_LNO += 1
                    rowGLTCHBL2 = dst.Tables("GLTCHBL2").NewRow
                    rowGLTCHBL2("XNO") = XNO
                    rowGLTCHBL2("XNO_LNO") = XNO_LNO
                    For Each col3 As PropertyInfo In cols3
                        If dst.Tables("GLTCHBL2").Columns.Contains(col.Name & "_" & col3.Name) Then
                            Dim val As Object = col3.GetValue(balanceList)
                            rowGLTCHBL2.Item(col.Name & "_" & col3.Name) = val
                        End If
                    Next
                    rowGLTCHBL2("BATCH_NO_CLEARED") = BATCH_NO_CLEARED
                    dst.Tables("GLTCHBL2").Rows.Add(rowGLTCHBL2)
                Next
            End If
        Next
        rowGLTCHBL1("XNO") = XNO
        rowGLTCHBL1("XDATE") = XDATE
        rowGLTCHBL1("BATCH_NO_CLEARED") = BATCH_NO_CLEARED
        rowGLTCHBL1("OPS_YYYYPP") = OPS_YYYYPP
        dst.Tables("GLTCHBL1").Rows.Add(rowGLTCHBL1)

        'BeginTrans()
        Update_Record_TDA("GLTCHBL1")
        Update_Record_TDA("GLTCHBL2")

    End Function
    Async Function Get_Trans(accountID As String) As Task
        'Get Transactions
        Dim clientJPMC As HttpClient = New HttpClient()
        Dim urlJPMC_Transactions As String = "https://openbanking.jpmorgan.com/tsapi/v3/"
        clientJPMC.BaseAddress = New Uri(urlJPMC_Transactions)
        clientJPMC.DefaultRequestHeaders.Accept.Clear()
        clientJPMC.DefaultRequestHeaders.Accept.Add(New MediaTypeWithQualityHeaderValue("*/*"))
        clientJPMC.DefaultRequestHeaders.Add("Authorization", "Bearer " & ACCESS_TOKEN)

        Dim d As New Dictionary(Of String, String)
        d.Add("pageNumber", "0")

        'Dim accountId As String = "000000899558985" '"000000010013324"
        'accountId = "000000899558928"
        '    accountId = "000000899558985"
        'accountId = "000000899558928,000000899558985"
        d.Add("accountIds", accountID)

        Dim startDate As Date = Date.MinValue
        Dim endDate As Date = Date.MinValue
        Dim morePages As Boolean = False
        d.Add("startDate", startDate)
        d.Add("endDate", endDate)

        Dim DATES As Date() = ASCMAIN1.Get_Dates(OPS_YYYYPP)
        'Dim DATES As Date() = ASCMAIN1.Get_Dates(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
        ' Dim DATES As Date() = {#9/30/2024#, #10/1/2024#, #10/2/2024#, #10/3/2024#, #10/4/2024#, #10/5/2024#, #10/6/2024#, #10/7/2024#, #10/8/2024#, #10/9/2024#, #10/10/2024#}
        Dim count As Integer = 0
        XNO_LNO = 0
        XNO_LNO_0 = 0
        For Each day As Date In DATES
            If day = DATES(0) Then
                Continue For
            End If
            count += 1
            If startDate = Date.MinValue Then
                startDate = day
            End If
            endDate = day
            If count Mod 5 = 0 Or count = DATES.Length - 1 Then
                d("startDate") = Format(startDate, "yyyy-MM-dd")
                d("endDate") = Format(endDate, "yyyy-MM-dd")
                d("pageNumber") = "1" '(Val(d("pageNumber")) + 1).ToString
                morePages = True

                Do While morePages = True
                    Dim queryString As New StringBuilder()
                    Dim delim As String = "?"
                    For Each k As String In d.Keys
                        queryString.Append($"{delim}{k}=").Append(Uri.EscapeDataString(d(k)))
                        delim = "&"
                    Next
                    Dim reqURL = "transactions" & queryString.ToString()
                    Dim res2 As HttpResponseMessage = Await clientJPMC.GetAsync(reqURL)

                    Dim bodyTransactions As String = Await res2.Content.ReadAsStringAsync()
                    Dim cTransactions As JPMC_Transactions_Response = JsonConvert.DeserializeObject(Of JPMC_Transactions_Response)(bodyTransactions)

                    'grdTransactions.DataSource = New List(Of JPMC_Transactions_Response) From {cTransactions}
                    Dim cc As New List(Of JPMC_Transactions_Response)
                    cc.Add(cTransactions)
                    If cc(0).pagination.pageNumber = cc(0).pagination.totalPages Then
                        morePages = False
                    ElseIf cc(0).pagination.totalPages = 0 Then
                        morePages = False
                    Else
                        d("pageNumber") = (Val(d("pageNumber")) + 1).ToString
                    End If
                    ProcessTransactions(cc, d)
                Loop
                startDate = Date.MinValue
            End If
        Next
        Update_Record_TDA("GLTCHTR0")
        Update_Record_TDA("GLTCHTR1")

    End Function
    Private Sub MapProperties(dataObj As Object, row As DataRow)
        ' Get properties of the top-level object
        Dim properties As PropertyInfo() = dataObj.GetType().GetProperties()

        ' Iterate over the properties
        For Each prop As PropertyInfo In properties
            ' Check if the top-level column exists and assign the value
            If dst.Tables("GLTCHTR0").Columns.Contains(prop.Name) Then
                Dim val As Object = prop.GetValue(dataObj)
                row.Item(prop.Name) = If(val IsNot Nothing, val.ToString(), DBNull.Value)
            End If

            'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" AndAlso prop.Name = "narrativeText" Then Stop


            If dst.Tables("GLTCHTR1").Columns.Contains(prop.Name) Then
                Dim val As Object = prop.GetValue(dataObj)
                row.Item(prop.Name) = If(val IsNot Nothing, val.ToString(), DBNull.Value)
            End If

            ' Handle nested objects by calling MapNestedProperties
            Select Case prop.Name
                Case "account"
                    MapNestedProperties(prop.GetValue(dataObj), "account", row)
                Case "baiType"
                    MapNestedProperties(prop.GetValue(dataObj), "baiType", row)
                Case "bankReferenceSearchable"
                    MapNestedProperties(prop.GetValue(dataObj), "bankReferenceSearchable", row)
                Case "currency"
                    MapNestedProperties(prop.GetValue(dataObj), "currency", row)
                Case "customerReferenceSearchable"
                    MapNestedProperties(prop.GetValue(dataObj), "customerReferenceSearchable", row)
                Case "lockbox"
                    MapNestedProperties(prop.GetValue(dataObj), "lockbox", row)
                Case "narrativeText"
                    MapNestedProperties(prop.GetValue(dataObj), "narrativeText", row)
            End Select
        Next
    End Sub

    Private Sub MapNestedProperties(nestedObj As Object, prefix As String, row As DataRow)
        ' Ensure the nested object is not null
        If nestedObj Is Nothing Then Return

        Dim nestedProps As PropertyInfo() = nestedObj.GetType().GetProperties()

        ' Iterate over the properties
        For Each nestedProp As PropertyInfo In nestedProps
            Dim colName As String = prefix & "_" & nestedProp.Name
            If dst.Tables("GLTCHTR1").Columns.Contains(colName) Then
                Dim val As Object = nestedProp.GetValue(nestedObj)
                row.Item(colName) = If(val IsNot Nothing, val, DBNull.Value)
            End If
        Next
    End Sub

    Private Sub ProcessTransactions(transactions As List(Of JPMC_Transactions_Response), D As Dictionary(Of String, String))
        For Each transaction As JPMC_Transactions_Response In transactions
            Dim rowGLTCHTR0 As DataRow = dst.Tables("GLTCHTR0").NewRow
            XNO_LNO_0 += 1
            MapProperties(transaction.pagination, rowGLTCHTR0)

            rowGLTCHTR0("XNO") = XNO
            rowGLTCHTR0("XDATE") = XDATE
            rowGLTCHTR0("XNO_LNO") = XNO_LNO_0
            rowGLTCHTR0("XSTARTDATE") = D("startDate")
            rowGLTCHTR0("XENDDATE") = D("endDate")

            dst.Tables("GLTCHTR0").Rows.Add(rowGLTCHTR0)
            If transaction.data IsNot Nothing AndAlso transaction.data.Count > 0 Then
                For Each dataItem As Object In transaction.data

                    Dim rowGLTCHTR1 As DataRow = dst.Tables("GLTCHTR1").NewRow

                    XNO_LNO += 1

                    MapProperties(dataItem, rowGLTCHTR1)

                    rowGLTCHTR1("XNO") = XNO
                    rowGLTCHTR1("XDATE") = XDATE
                    rowGLTCHTR1("XNO_LNO") = XNO_LNO_0
                    rowGLTCHTR1("XNO_DTL_LNO") = XNO_LNO
                    rowGLTCHTR1("BATCH_NO_CLEARED") = BATCH_NO_CLEARED

                    ' Add the processed DataRow
                    dst.Tables("GLTCHTR1").Rows.Add(rowGLTCHTR1)
                Next
            End If
        Next
    End Sub

    Private Sub Match_Transactions()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Matching Transactions")

        For Each rowGLTBREC2 As DataRow In dst.Tables("GLTBREC2").Select("TRAN_SEL= '1'")
            rowGLTBREC2.Item("TRAN_SEL") = "0"
        Next
        For Each rowGLTCHTR1 As DataRow In dst.Tables("GLTCHTR1").Select("MATCHED = '1'")
            rowGLTCHTR1.Item("MATCHED") = "0"
            rowGLTCHTR1.Item("MATCHED_AMT") = 0
        Next

        'Check_if_Matches_Are_in_Balance()
        Dim MAX_DATE As String = dst.Tables("GLTCHBL2").Compute("MAX(BALANCELIST_ASOFDATE)", "BALANCELIST_ENDINGAVAILABLEAMOUNT > 0").ToString()
        Dim balanceRow As DataRow() = dst.Tables("GLTCHBL2").Select($"BALANCELIST_ASOFDATE = '{MAX_DATE}' AND BALANCELIST_ENDINGAVAILABLEAMOUNT > 0")
        If balanceRow.Length > 0 Then
            txtBALANCE.Text = Convert.ToDecimal(balanceRow(0)("BALANCELIST_ENDINGAVAILABLEAMOUNT")).ToString("F2")
        End If

        ASCMAIN1.sql = "SELECT CUST_PYMT_REF_NO, CUST_PYMT_AMT, PYMT_BATCH_NO FROM ARTPYMT2 WHERE NVL(PYMT_DELETED, '0') <> '1'"
        Dim DT As DataTable = ASCDATA1.GetDataTable


        ' Match Funding Transfers

        Dim rowsGLTCHTR1 As DataTable = dst.Tables("GLTCHTR1")
        For Each row1 As DataRow In rowsGLTCHTR1.Select("NARRATIVETEXT_REMARK like 'FUNDING XFER FROM*' AND ISNULL(MATCHED,'0') <> '1'")
            For Each row2 As DataRow In rowsGLTCHTR1.Select("NARRATIVETEXT_REMARK like 'FUNDING XFER TO*' AND ISNULL(MATCHED,'0') <> '1'")

                Dim amount1 As Decimal = -1 * Convert.ToDecimal(row1("AMOUNT"))
                Dim amount2 As Decimal = Convert.ToDecimal(row2("AMOUNT"))
                Dim date1 As Date = Convert.ToDateTime(row1("ASOFDATE"))
                Dim date2 As Date = Convert.ToDateTime(row2("ASOFDATE"))

                ' Proceed only if both records have offsetting amounts
                If amount1 = -1 * amount2 AndAlso Format(date1, "yyyyMMdd") = Format(date2, "yyyyMMdd") Then
                    ' Set both as matched
                    row1("MATCHED") = "1"
                    row1("MATCHED_AMT") = amount1
                    row2("MATCHED") = "1"
                    row2("MATCHED_AMT") = amount2
                End If
            Next
        Next

        'Check_if_Matches_Are_in_Balance()

        ' Match ACH Payments recorded by Bank to the individual checks recorded in ABS ACH Batch

        For Each rowGLTCHTR1 As DataRow In dst.Tables("GLTCHTR1").Select("NARRATIVETEXT_REMARK Like 'ACH*' AND ISNULL(MATCHED,'0') <> '1'") ' Look for ACH
            Dim ASOFDATE As Date = Convert.ToDateTime(rowGLTCHTR1("ASOFDATE"))
            Dim AMOUNT As Decimal = Val(rowGLTCHTR1.Item("AMOUNT") & "")

            'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" And AMOUNT = 779907.68 Then Stop

            'Dim bankCodeFilter As String = If(BANK_CODE = "CHASE", "in ('CHASE', 'WIRE')", "= '" & BANK_CODE & "'") ' this line is also at 549
            ' 1 it should be form global
            ' 2 it should not refer to CHASE - insted it should refer to multiple banks that are tied together

            Dim BATCH_NO_ACH As String = ""
            Dim BATCH_TOTAL As Decimal = 0

            Dim sqlACH As String = $"SELECT APTCHCK1.BATCH_NO_ACH, APTCHCK1.BANK_CODE, APTCHCK1.CHECK_NUM
            , APTCHCK1.CHECK_AMT, APTCHCK1.CHECK_DATE, APTCHCKA.XMIT_DATE
            FROM APTCHCKA,APTCHCK1
            WHERE APTCHCK1.BATCH_NO_ACH = APTCHCKA.BATCH_NO_ACH
            AND APTCHCKA.BATCH_ACH_STATUS = 'S'
            AND APTCHCKA.BANK_CODE {BANK_CODE_FILTER}"

            ASCMAIN1.sql = $"SELECT BATCH_NO_ACH, XMIT_DATE
            , SUM (CHECK_AMT) BATCH_TOTAL, COUNT (*) CHECKS FROM (
            {sqlACH}
            AND APTCHCKA.XMIT_DATE <= '{Format(ASOFDATE, "dd-MMM-yyyy")}'
            AND APTCHCKA.XMIT_DATE >= '{Format(ASOFDATE.AddDays(-5), "dd-MMM-yyyy")}'
            ) GROUP BY BATCH_NO_ACH, XMIT_DATE
            ORDER BY XMIT_DATE DESC"
            For Each row As DataRow In ASCDATA1.GetDataTable().Select("", "XMIT_DATE DESC")
                BATCH_TOTAL = Val(row.Item("BATCH_TOTAL") & "")
                If BATCH_TOTAL = AMOUNT Then
                    BATCH_NO_ACH = row.Item("BATCH_NO_ACH")
                    Exit For
                End If
            Next

            If BATCH_NO_ACH <> "" Then
                'ASCMAIN1.sql =
                '"SELECT COUNT(*) AS VOID_COUNT " &
                '"FROM APTCHCK1 " & vbCrLf &
                '"WHERE BATCH_NO_ACH = '" & BATCH_NO_ACH & "' " & vbCrLf &
                '" AND CHECK_STATUS = 'V'"

                'Dim dtVoidCheck As DataTable = ASCDATA1.GetDataTable()
                'Dim voidCount As Integer = 0
                'If dtVoidCheck.Rows.Count > 0 Then
                '    voidCount = Convert.ToInt32(dtVoidCheck.Rows(0)("VOID_COUNT"))
                'End If

                'If voidCount > 0 Then
                '    Continue For
                'End If
                ASCMAIN1.sql = $"{sqlACH} AND APTCHCKA.BATCH_NO_ACH = '{BATCH_NO_ACH}'"
                For Each row As DataRow In ASCDATA1.GetDataTable().Select("")
                    Dim BANK_CODE As String = row.Item("BANK_CODE")
                    Dim CHECK_NUM As String = row.Item("CHECK_NUM")
                    Dim rowGLTBREC2s() As DataRow = dst.Tables("GLTBREC2").Select($"BANK_CODE = '{BANK_CODE}' and TRAN_KEY = '{CHECK_NUM}'")
                    If rowGLTBREC2s.Length = 1 Then
                        Dim rowGLTBREC2 As DataRow = rowGLTBREC2s(0)
                        rowGLTBREC2("TRAN_SEL") = "1"
                    End If
                Next

                rowGLTCHTR1.Item("MATCHED") = "1"
                rowGLTCHTR1("MATCHED_AMT") = -1 * BATCH_TOTAL
            End If
        Next

        'Check_if_Matches_Are_in_Balance()

        For Each rowGLTBREC2 As DataRow In dst.Tables("GLTBREC2").Select("ISNULL(TRAN_SEL,'0') <> '1'")
            Dim TRAN_KEY As String = rowGLTBREC2.Item("TRAN_KEY") & ""
            Dim TRAN_AMT As Decimal = Convert.ToDecimal(rowGLTBREC2("TRAN_AMT"))
            'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" AndAlso TRAN_AMT = 154707.5 Then Stop
            Dim TRAN_DATE As String = Convert.ToDateTime(rowGLTBREC2("TRAN_DATE")).ToString("yyyy-MM-dd")
            Dim TRAN_DATE_1 As String = Convert.ToDateTime(rowGLTBREC2("TRAN_DATE")).AddDays(-1).ToString("yyyy-MM-dd")
            Dim TRAN_DATE_2 As String = Convert.ToDateTime(rowGLTBREC2("TRAN_DATE")).AddDays(-2).ToString("yyyy-MM-dd")
            Dim TRAN_DATE_3 As String = Convert.ToDateTime(rowGLTBREC2("TRAN_DATE")).AddDays(-3).ToString("yyyy-MM-dd")

            Dim sqlDates As String = $" and (ASOFDATE = '{TRAN_DATE}' or ASOFDATE = '{TRAN_DATE_1}' or ASOFDATE = '{TRAN_DATE_2}' or ASOFDATE = '{TRAN_DATE_3}')"

            Dim TRAN_DESC As String = rowGLTBREC2.Item("TRAN_DESC") & ""
            Dim JOURNAL_TYPE As String = rowGLTBREC2.Item("JOURNAL_TYPE") & ""

            If JOURNAL_TYPE = "ARCR" AndAlso TRAN_AMT = 0 Then
                rowGLTBREC2.Item("TRAN_SEL") = "1"
                Continue For
            End If

            Dim totalAmountString As String = String.Empty
            If TRAN_DESC.Contains("Total") Then
                Dim parts As String() = TRAN_DESC.Split(New String() {"Total: "}, StringSplitOptions.None)
                If parts.Length > 1 Then
                    Dim amountPart As String = parts(1).Trim()
                    If amountPart.Length > 0 Then
                        totalAmountString = amountPart
                    End If
                End If
            End If

            'Check_if_Matches_Are_in_Balance()

            Dim matched As Boolean = False
            Dim matchingRows As DataRow()

            If totalAmountString <> "" Then

                ' Convert extracted amount to decimal and see if theres a match
                Dim TRAN_AMT_FROM_DESC As Decimal = 0
                Decimal.TryParse(totalAmountString, TRAN_AMT_FROM_DESC)

                matchingRows = dst.Tables("GLTCHTR1").Select($"AMOUNT = {TRAN_AMT_FROM_DESC} {sqlDates} AND ISNULL(MATCHED,'0') <> '1'")
                If matchingRows.Length = 1 Then

                    For Each rowGLTCHTR1 As DataRow In matchingRows
                        Dim amount As Decimal = Val(rowGLTCHTR1("AMOUNT") & "")
                        'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" And (TRAN_AMT_FROM_DESC = 32787.8 Or amount = 32787.8) Then Stop
                        If TRAN_AMT_FROM_DESC = amount Then
                            rowGLTCHTR1("MATCHED") = "1"
                            rowGLTCHTR1("MATCHED_AMT") = amount
                            'Else
                            '    rowGLTBREC2.Item("TRAN_SEL") = "0" ' Checkbox unchecked
                            '    rowGLTCHTR1.Item("MATCHED") = "0"
                        End If
                    Next

                    Dim sqlTotal As String = $"TRAN_KEY = '{TRAN_KEY}' and TRAN_DESC = '{TRAN_DESC}' AND ISNULL(TRAN_SEL,'0') <> '1'"
                    For Each row2 As DataRow In dst.Tables("GLTBREC2").Select(sqlTotal)
                        row2.Item("TRAN_SEL") = "1" ' Checkbox checked
                    Next
                    matched = True

                    'Check_if_Matches_Are_in_Balance()
                End If
            End If

            'Check_if_Matches_Are_in_Balance()

            If matched Then Continue For

            ' 1. Check if there is a match based on CHECKNUMBER = TRAN_KEY
            Dim sqlwC As String = $"CHECKNUMBER = '{TRAN_KEY}' AND ISNULL(MATCHED,'0') <> '1'"
            matchingRows = dst.Tables("GLTCHTR1").Select(sqlwC)

            For Each rowGLTCHTR1 As DataRow In matchingRows
                Dim amount As Decimal = Val(rowGLTCHTR1.Item("AMOUNT") & "")

                ' NEED TO PUT THE AMOUNT MATCHING  IN THE SQLWHERE AND LIMIT TO DEBITS IF BREC JOURNAL_TYPE = APCD
                ' Standard match check: TRAN_AMT + amount = 0
                If (TRAN_AMT + amount) = 0 Then
                    rowGLTBREC2("TRAN_SEL") = "1" ' Checkbox checked
                    rowGLTCHTR1("MATCHED") = "1"
                    rowGLTCHTR1("MATCHED_AMT") = TRAN_AMT
                    matched = True
                    Exit For
                ElseIf TRAN_AMT <> amount Then
                    rowGLTBREC2("TRAN_SEL") = "0" ' Checkbox unchecked
                    rowGLTCHTR1("MATCHED") = "2"
                    rowGLTCHTR1("MATCHED_AMT") = amount.ToString("F2")
                End If
            Next

            'Check_if_Matches_Are_in_Balance()

            If matched Then Continue For


            ' Step 2: Special handling for CHECKNUMBER = '0000000000'

            Dim sqlAmount As String = $" and ((DEBITCREDITCODE = 'DEBIT' and AMOUNT = {-1 * TRAN_AMT}) OR (DEBITCREDITCODE = 'CREDIT' and AMOUNT = {TRAN_AMT}))"
            matchingRows = dst.Tables("GLTCHTR1").Select($"CHECKNUMBER = '0000000000' {sqlDates} {sqlAmount} AND ISNULL(MATCHED,'0') <> '1'")
            If matchingRows.Length = 1 Then
                rowGLTBREC2("TRAN_SEL") = "1" ' Checkbox checked
                matchingRows(0).Item("MATCHED") = "1"
                matchingRows(0).Item("MATCHED_AMT") = TRAN_AMT
                matched = True
            End If
            If matched Then Continue For

            ' DELETE AFTER REVIEW
            'For Each rowGLTCHTR1 As DataRow In matchingRows
            '    Dim amount As Decimal = Convert.ToDecimal(rowGLTCHTR1("AMOUNT"))
            '    Dim DEBITCREDITCODE As String = rowGLTCHTR1("DEBITCREDITCODE") & ""
            '    If DEBITCREDITCODE = "DEBIT" Then amount = -1 * amount
            '    ' Match based on amount + TRAN_AMT = 0 and matching date
            '    If ((TRAN_AMT + amount) = 0 Or TRAN_AMT = amount) AndAlso (Convert.ToDateTime(rowGLTCHTR1("ASOFDATE")).ToString("yyyy-MM-dd") = TRAN_DATE) Then
            '        rowGLTBREC2("TRAN_SEL") = "1" ' Checkbox checked
            '        rowGLTCHTR1("MATCHED") = "1"
            '        rowGLTCHTR1("MATCHED_AMT") = amount.ToString("F2")
            '        matched = True
            '        Exit For
            '    End If
            'Next

            ' 3. Match based on CUST_PYMT_REF_NO and AMT
            Dim artPymtRows As DataRow() = DT.Select("PYMT_BATCH_NO = '" & TRAN_KEY & "'")

            For Each rowARTPYMT2 As DataRow In artPymtRows
                Dim CUST_PYMT_REF_NO As String = rowARTPYMT2("CUST_PYMT_REF_NO").ToString()
                Dim CUST_PYMT_AMT As Decimal = Math.Round(Convert.ToDecimal(rowARTPYMT2("CUST_PYMT_AMT")), 2)

                Dim refMatchingRows As DataRow() = dst.Tables("GLTCHTR1").Select("CUSTOMERREFERENCESEARCHABLE_STANDARDVALUE = '" & CUST_PYMT_REF_NO & "' AND ISNULL(MATCHED,'0') <> '1'")
                For Each rowGLTCHTR1 As DataRow In refMatchingRows
                    Dim amount As Decimal = Val(rowGLTCHTR1("AMOUNT") & "")

                    ' Check if the amounts match to 2 decimal places
                    If CUST_PYMT_AMT = amount Then
                        rowGLTBREC2("TRAN_SEL") = "1" ' Checkbox checked
                        rowGLTCHTR1("MATCHED") = "1"
                        rowGLTCHTR1("MATCHED_AMT") = amount
                        matched = True
                    Else
                        rowGLTBREC2("TRAN_SEL") = "0" ' Checkbox unchecked
                        rowGLTCHTR1("MATCHED") = "0"
                    End If
                Next
            Next

        Next ' done with rowGLTBREC2

        'Check_if_Matches_Are_in_Balance()

        ' Match MACYS Payments to combined payments applied from MACYS EDI 820

        Dim sqlw As String = " and JOURNAL_TYPE = 'ARCR' AND ISNULL(TRAN_SEL,'0') <> '1'"
        For Each row As DataRow In dst.Tables("GLTBREC2").Select("CUST_VEND = 'MACYS'" & sqlw)
            Dim CUST_REF_NO As String = row.Item("CUST_REF_NO") & ""
            Dim TRAN_DATE As Date = row.Item("TRAN_DATE")

            If CUST_REF_NO <> "" Then
                Dim TOTAL As Decimal = 0
                Dim rows() As DataRow = dst.Tables("GLTBREC2").Select($"CUST_REF_NO = '{CUST_REF_NO}' and TRAN_DATE = #{Format(TRAN_DATE, "MM/dd/yyyy")}#" & sqlw)
                For Each row2 As DataRow In rows
                    Dim TRAN_AMT As Decimal = Val(row2.Item("TRAN_AMT") & "")
                    TOTAL += TRAN_AMT
                Next

                Dim sqlw2 As String = $"AMOUNT = {CStr(TOTAL)} AND DEBITCREDITCODE = 'CREDIT' AND ASOFDATE = '{Format(TRAN_DATE, "yyyy-MM-dd")}'"
                sqlw2 = $"AMOUNT = {CStr(TOTAL)} AND DEBITCREDITCODE = 'CREDIT' and NARRATIVETEXT_REMARK LIKE 'TRN*' and NARRATIVETEXT_REMARK like '*{CUST_REF_NO}*'"

                Dim rowGLTCHTR1s() As DataRow = dst.Tables("GLTCHTR1").Select(sqlw2)
                If rowGLTCHTR1s.Length = 1 Then
                    rowGLTCHTR1s(0).Item("MATCHED") = "1"
                    rowGLTCHTR1s(0).Item("MATCHED_AMT") = TOTAL
                    For Each row2 As DataRow In rows
                        row2("TRAN_SEL") = "1"
                    Next
                    'Else
                    '    Stop
                End If
            End If
            'Check_if_Matches_Are_in_Balance()
            ' match positives and negatives - not easy because they don't always have the same batch no

        Next

        'Check_if_Matches_Are_in_Balance()

        dst.Tables("GLTCHTR1").AcceptChanges()
        dst.Tables("GLTBREC2").AcceptChanges()
        Display_Totals()

        ' UltraExplorerBar1.Groups("Screen Control").Items("Match Transactions").Settings.Enabled = DefaultableBoolean.False

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


    End Sub
    Private Sub Find_Possible_Matches()
        If grdGLTBREC2.ActiveRow Is Nothing Then
            MessageBox.Show("Please select a row to find possible matches.", "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Dim selectedRow As UltraWinGrid.UltraGridRow = grdGLTBREC2.ActiveRow
        Dim selectedAmount As Decimal = Convert.ToDecimal(selectedRow.Cells("TRAN_AMT").Value)
        Dim TRAN_DESC As String = selectedRow.Cells("TRAN_DESC").Value.ToString()

        Dim descAmount As Decimal = 0
        If TRAN_DESC.Contains("Total") Then
            Dim parts As String() = TRAN_DESC.Split(New String() {"Total: "}, StringSplitOptions.None)
            If parts.Length > 1 Then
                Dim amountPart As String = parts(1).Trim()
                Decimal.TryParse(amountPart, descAmount)
            End If
        End If

        Dim tblPossibleMatches As New DataTable
        tblPossibleMatches.Columns.Add("CHECKNUMBER", GetType(String))
        tblPossibleMatches.Columns.Add("ASOFDATE", GetType(Date))
        tblPossibleMatches.Columns.Add("AMOUNT", GetType(Decimal))

        For Each rowGLTCHTR1 As DataRow In dst.Tables("GLTCHTR1").Rows
            Dim matched As String = If(IsDBNull(rowGLTCHTR1("MATCHED")), String.Empty, rowGLTCHTR1("MATCHED").ToString())
            Dim amount As Decimal = Convert.ToDecimal(rowGLTCHTR1("AMOUNT"))

            If matched <> "1" AndAlso (amount = selectedAmount OrElse amount = descAmount) Then
                Dim newRow As DataRow = tblPossibleMatches.NewRow()
                newRow("CHECKNUMBER") = rowGLTCHTR1("CHECKNUMBER").ToString()
                newRow("ASOFDATE") = Convert.ToDateTime(rowGLTCHTR1("ASOFDATE"))
                newRow("AMOUNT") = amount
                tblPossibleMatches.Rows.Add(newRow)
            End If
        Next

        If tblPossibleMatches.Rows.Count = 0 Then
            MessageBox.Show("No possible matches found for the selected transaction.", "No Matches", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Using F As New ASFMSGBF
            F.Show_grd(tblPossibleMatches, Me, "Possible Matches")
        End Using
    End Sub

    Private Sub grdGLTCHTR1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdGLTCHTR1.AfterRowActivate
        If chkShowAll.Checked Then
            grdGLTMATCH.Rows.Refresh(RefreshRow.FireInitializeRow)
            Exit Sub
        End If
        Dim activeRow As UltraWinGrid.UltraGridRow = grdGLTCHTR1.ActiveRow
        If activeRow Is Nothing Or activeRow.IsFilterRow Then Exit Sub

        Dim dvwGLTBREC2 As New DataView(dst.Tables("GLTBREC2"))
        If activeRow.Cells("MATCHED").Value.ToString() <> "1" Then
            Dim amount As Decimal = Convert.ToDecimal(activeRow.Cells("AMOUNT").Value)
            dvwGLTBREC2.RowFilter = $"(TRAN_SEL IS NULL OR TRAN_SEL <> '1') AND (TRAN_AMT = {amount} OR TRAN_AMT + {amount} = 0)"
            grdGLTMATCH.DataSource = dvwGLTBREC2
        End If
    End Sub


    Private Sub grdGLTMATCH_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdGLTMATCH.AfterCellUpdate
        If e.Cell.Column.Key = "RECONCILE" Then
            Dim isChecked As String = e.Cell.Value
            Dim amount As Decimal = e.Cell.Row.Cells("TRAN_AMT").Value
            Dim dateValue As Date = Convert.ToDateTime(e.Cell.Row.Cells("TRAN_DATE").Value)
            If isChecked Then
                e.Cell.Row.Cells("RECONCILE_AMT").Value = amount
            Else
                e.Cell.Row.Cells("RECONCILE_AMT").Value = DBNull.Value
            End If
            'Dim XNO As String = e.Cell.Row.Cells("XNO").Value
            'Dim XNO_LNO As Int32 = Val(e.Cell.Row.Cells("XNO_LNO").Value & "")
            'Dim XNO_DTL_LNO As Int32 = Val(e.Cell.Row.Cells("XNO_DTL_LNO").Value & "")
            'Dim rowGLTCHTR1 As DataRow = dst.Tables("GLTCHTR1").Rows.Find(New Object() {XNO, XNO_LNO, XNO_DTL_LNO})
            'rowGLTCHTR1.Item("MATCHED") = isChecked
            ' Update the matched status in grdGLTCHTR1
            'For Each row As UltraGridRow In grdGLTCHTR1.Rows
            '    If Convert.ToDecimal(row.Cells("AMOUNT").Value) = amount Then
            '        row.Cells("MATCHED").Value = isChecked
            '        row.Update()
            '    End If
            'Next

            ' Update the corresponding DataRow in GLTBREC2
            'For Each row As DataRow In dst.Tables("GLTBREC2").Rows
            '    If Convert.ToDecimal(row("TRAN_AMT")) = amount AndAlso Convert.ToDateTime(row("TRAN_DATE")) = dateValue Then
            '        row("TRAN_SEL") = isChecked
            '        Exit For
            '    End If
            'Next
        End If
    End Sub
    Private Sub grdGLTCHTR1_AfterCellUpdate(sender As Object, e As CellEventArgs) Handles grdGLTCHTR1.AfterCellUpdate
        If e.Cell.Column.Key = "RECONCILE" Then
            'if lower grid has one row
            Dim RECONCILE_AMT As Decimal = Val(e.Cell.Row.Cells("AMOUNT").Value & "")
            Dim DEBITCREDITCODE As String = e.Cell.Row.Cells("DEBITCREDITCODE").Value.ToString()
            If DEBITCREDITCODE = "DEBIT" Then
                RECONCILE_AMT = -Math.Abs(RECONCILE_AMT)
            End If
            If grdGLTMATCH.Rows.Count = 1 Then

                If Val(grdGLTMATCH.Rows(0).Cells("TRAN_AMT").Value & "") = RECONCILE_AMT Then
                    grdGLTMATCH.Rows(0).Cells("RECONCILE").Value = "1"
                    grdGLTMATCH.Rows(0).Cells("RECONCILE_AMT").Value = RECONCILE_AMT
                    grdGLTMATCH.Rows(0).Update()
                End If
            End If
            Dim isChecked As String = e.Cell.Value

            If isChecked Then
                e.Cell.Row.Cells("RECONCILE_AMT").Value = RECONCILE_AMT
            Else
                e.Cell.Row.Cells("RECONCILE_AMT").Value = DBNull.Value
            End If

            'Dim dvwGLTBREC2 As DataView = CType(grdGLTMATCH.DataSource, DataView)
            'If dvwGLTBREC2 IsNot Nothing Then
            '    For Each row As DataRowView In dvwGLTBREC2
            '        If Convert.ToDecimal(row("TRAN_AMT")) = amount Then
            '            row("RECONCILE") = If(isChecked, "1", "0")
            '            row("RECONCILE_AMT") = If(isChecked, amount, DBNull.Value)
            '        End If
            '    Next
            'End If
        End If
    End Sub

    Private Sub chkShowAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAll.CheckedChanged
        Dim dvw As DataView = CType(grdGLTMATCH.DataSource, DataView)

        If chkShowAll.Checked Then
            dvw.RowFilter = "TRAN_SEL <> '1' OR TRAN_SEL IS NULL"
        Else
            Dim activeRow As UltraGridRow = grdGLTCHTR1.ActiveRow
            If activeRow IsNot Nothing AndAlso Not activeRow.IsFilterRow Then
                Dim amount As Decimal = Convert.ToDecimal(activeRow.Cells("AMOUNT").Value)
                dvw.RowFilter = $"(TRAN_SEL IS NULL OR TRAN_SEL <> '1') AND (TRAN_AMT = {amount} OR TRAN_AMT + {amount} = 0)"
            End If
        End If

        grdGLTMATCH.Refresh()
    End Sub
    Private Sub Reconcile_Mode()
        dst.Tables("GLTCHTR1").BeginLoadData()
        dst.Tables("GLTBREC2").BeginLoadData()
        For Each row As DataRow In dst.Tables("GLTBREC2").Rows
            If row("RECONCILE").ToString() = "1" Then
                row("RECONCILE") = "0"
                row("RECONCILE_AMT") = DBNull.Value
            End If
        Next

        Dim dvw As DataView = CType(grdGLTCHTR1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "(MATCHED <> '1' OR MATCHED IS NULL) AND (RECONCILING_ITEM_IND <> '0' OR RECONCILING_ITEM_IND IS NULL)"
        SplitContainer1.Panel2Collapsed = False
        tabs.Tabs(1).Visible = False
        tabs.SelectedTab = tabs.Tabs(2)
        UltraExplorerBar1.Groups("Screen Control").Visible = False
        UltraExplorerBar1.Groups("Statement").Visible = False
        UltraExplorerBar1.Groups("Reconcile").Items("Reconcile").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Reconcile").Items("Save Reconcile").Settings.Enabled = DefaultableBoolean.True
        UltraExplorerBar1.Groups("Reconcile").Items("Cancel Reconcile").Settings.Enabled = DefaultableBoolean.True
        UltraExplorerBar1.Groups("Reconcile").Items("Clear Reconciled").Settings.Enabled = DefaultableBoolean.True

        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED").Hidden = True
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED_AMT").Hidden = True
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE").Hidden = False
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE_AMT").Hidden = False

        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED").CellActivation = Activation.AllowEdit
        grdGLTMATCH.DisplayLayout.Bands(0).Columns("RECONCILE").CellActivation = Activation.AllowEdit
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE").CellActivation = Activation.AllowEdit
        grdGLTMATCH.Visible = True
        chkShowAll.Visible = True
        chkShowAll.Checked = False
        grdGLTMATCH.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        'chk_showReconciled.Enabled = True
        chk_showReconciled.Checked = False
        chk_showReconciled.Visible = True
    End Sub
    Private Sub Save_Reconcile()
        For Each row As DataRow In dst.Tables("GLTCHTR1").Rows
            If row("RECONCILE").ToString() = "1" Then
                row("MATCHED") = "1"
                row("MATCHED_AMT") = row("RECONCILE_AMT")
            End If
        Next
        For Each row As DataRow In dst.Tables("GLTBREC2").Rows
            If row("RECONCILE") = "1" Then
                row("TRAN_SEL") = "1"
            End If
        Next
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED").Hidden = False
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED_AMT").Hidden = False
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE").Hidden = True
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE_AMT").Hidden = True

        Dim dvw As DataView = CType(grdGLTCHTR1.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Empty
        SplitContainer1.Panel2Collapsed = True
        tabs.Tabs(1).Visible = True
        tabs.SelectedTab = tabs.Tabs(2)
        UltraExplorerBar1.Groups("Screen Control").Visible = True
        UltraExplorerBar1.Groups("Statement").Visible = True
        UltraExplorerBar1.Groups("Reconcile").Items("Reconcile").Settings.Enabled = DefaultableBoolean.True
        UltraExplorerBar1.Groups("Reconcile").Items("Save Reconcile").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Reconcile").Items("Cancel Reconcile").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Reconcile").Items("Clear Reconciled").Settings.Enabled = DefaultableBoolean.False
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED").CellActivation = Activation.NoEdit
        grdGLTCHTR1.DataSource = dst.Tables("GLTCHTR1")
        grdGLTCHTR1.Refresh()
        Display_Totals()
        grdGLTMATCH.Visible = False
        chkShowAll.Visible = False
        chk_showReconciled.Checked = False
        'chk_showReconciled.Enabled = False
        chk_showReconciled.Visible = False
    End Sub

    Private Sub Cancel_Reconcile()
        dst.Tables("GLTCHTR1").RejectChanges()
        dst.Tables("GLTBREC2").RejectChanges()
        Dim checkedRowsGLTCHTR1 = dst.Tables("GLTCHTR1").Select("RECONCILE = '1'")
        For Each row As DataRow In checkedRowsGLTCHTR1
            row("RECONCILE") = "0"
            row("RECONCILE_AMT") = DBNull.Value
        Next

        Dim checkedRowsGLTMATCH = dst.Tables("GLTBREC2").Select("RECONCILE = '1'")
        For Each row As DataRow In checkedRowsGLTMATCH
            row("RECONCILE") = "0"
            row("RECONCILE_AMT") = DBNull.Value
        Next
        Dim dvw As DataView = CType(grdGLTCHTR1.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Empty
        SplitContainer1.Panel2Collapsed = True
        tabs.Tabs(1).Visible = True
        tabs.SelectedTab = tabs.Tabs(2)
        UltraExplorerBar1.Groups("Screen Control").Visible = True
        UltraExplorerBar1.Groups("Statement").Visible = True
        UltraExplorerBar1.Groups("Reconcile").Items("Reconcile").Settings.Enabled = DefaultableBoolean.True
        UltraExplorerBar1.Groups("Reconcile").Items("Save Reconcile").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Reconcile").Items("Cancel Reconcile").Settings.Enabled = DefaultableBoolean.False
        UltraExplorerBar1.Groups("Reconcile").Items("Clear Reconciled").Settings.Enabled = DefaultableBoolean.False

        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED").Hidden = False
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED_AMT").Hidden = False
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE").Hidden = True
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE_AMT").Hidden = True

        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("MATCHED").CellActivation = Activation.NoEdit
        grdGLTCHTR1.DisplayLayout.Bands(0).Columns("RECONCILE").CellActivation = Activation.NoEdit
        grdGLTCHTR1.Refresh()
        grdGLTMATCH.Visible = False
        chkShowAll.Visible = False
        chk_showReconciled.Checked = False
        'chk_showReconciled.Enabled = False
        chk_showReconciled.Visible = False ' True
    End Sub
    Private Sub Clear_Reconcile()
        For Each row As DataRow In dst.Tables("GLTCHTR1").Rows
            If row("RECONCILE").ToString() = "1" Then
                row("MATCHED") = "1"
                row("MATCHED_AMT") = row("RECONCILE_AMT")
            End If
        Next

        For Each row As DataRow In dst.Tables("GLTBREC2").Rows
            If row("RECONCILE").ToString() = "1" Then
                row("TRAN_SEL") = "1"
            End If
        Next

        Display_Totals()
    End Sub
    Private Function Get_Account_ID(BANK_CODE As String) As List(Of String)
        Dim accountIDs As New List(Of String)

        Dim primaryAccountRow As DataRow = dst.Tables("GLTBANK1").Select($"BANK_CODE = '{BANK_CODE}'")(0)
        If primaryAccountRow IsNot Nothing Then
            Dim primaryAccountID As String = primaryAccountRow("BANK_ACCT_ID").ToString().PadLeft(15, "0"c)
            accountIDs.Add(primaryAccountID)
        End If

        Dim subordinateRows As DataRow() = dst.Tables("GLTBANK3").Select($"BANK_CODE = '{BANK_CODE}'")
        For Each row As DataRow In subordinateRows
            Dim subordinateCode As String = row("BANK_CODE_SUB").ToString()
            Dim subordinateAccountRow As DataRow = dst.Tables("GLTBANK1").Select($"BANK_CODE = '{subordinateCode}'")(0)
            If subordinateAccountRow IsNot Nothing Then
                Dim subordinateAccountID As String = subordinateAccountRow("BANK_ACCT_ID").ToString().PadLeft(15, "0"c)
                accountIDs.Add(subordinateAccountID)
            End If
        Next

        Return accountIDs
    End Function

    Private Sub chk_showReconciled_CheckedChanged(sender As Object, e As EventArgs) Handles chk_showReconciled.CheckedChanged
        Dim dvwGLTCHTR1 As DataView = CType(grdGLTCHTR1.DataSource, DataTable).DefaultView
        Dim dvwGLTMATCH As DataView = CType(grdGLTMATCH.DataSource, DataView)

        If chk_showReconciled.Checked Then
            dvwGLTCHTR1.RowFilter = "RECONCILE IS NOT NULL OR (MATCHED <> '1' OR MATCHED IS NULL)"
            dvwGLTMATCH.RowFilter = "RECONCILE <> '0' OR (TRAN_SEL <> '1' OR TRAN_SEL IS NULL)"
        Else
            dvwGLTCHTR1.RowFilter = "MATCHED <> '1' OR MATCHED IS NULL"
            dvwGLTMATCH.RowFilter = "TRAN_SEL <> '1' OR TRAN_SEL IS NULL"
        End If

        grdGLTCHTR1.Refresh()
        grdGLTMATCH.Refresh()
    End Sub

    Private Sub grdGLTMATCH_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdGLTMATCH.InitializeRow
        If e.Row.IsDataRow And Not e.Row.IsFilterRow And Not e.Row.IsGroupByRow Then
            Dim C As System.Drawing.Color = System.Drawing.Color.Empty
            If grdGLTCHTR1.ActiveRow IsNot Nothing Then
                Dim ASOFDATETIME As Date

                With grdGLTCHTR1.ActiveRow
                    Dim NOTES As String = .Cells("NARRATIVETEXT_REMARK").Value & ""
                    If NOTES.StartsWith("ACH ") Then
                        ASOFDATETIME = .Cells("ASOFDATETIME").Value
                        Dim TRAN_DATE As Date = e.Row.Cells("TRAN_DATE").Value
                        Dim JOURNAL_TYPE As String = e.Row.Cells("JOURNAL_TYPE").Value & ""
                        If JOURNAL_TYPE = "APCD" And Format(TRAN_DATE, "yyyyMMdd") = Format(ASOFDATETIME.AddDays(-1), "yyyyMMdd") Then
                            C = System.Drawing.Color.Pink
                        End If
                    End If

                End With
            End If
            e.Row.Appearance.BackColor = C
        End If
    End Sub
    Function Check_if_Matches_Are_in_Balance(Optional caseType As String = "") As Boolean
        Dim inBalance As Boolean = True

        Select Case caseType.ToUpper()
            Case "SAVE RECONCILE", "CLEAR RECONCILED"
                Dim totalGLTCHTR1 As Decimal = Val(dst.Tables("GLTCHTR1").Compute(
                "SUM(RECONCILE_AMT)", "RECONCILE = '1' AND (MATCHED IS NULL OR MATCHED <> '1')") & "")
                Dim totalGLTBREC2 As Decimal = Val(dst.Tables("GLTBREC2").Compute(
                "SUM(RECONCILE_AMT)", "RECONCILE = '1' AND (TRAN_SEL IS NULL OR TRAN_SEL <> '1')") & "")

                If totalGLTCHTR1 <> totalGLTBREC2 Then inBalance = False

            Case "UPDATE"
                Dim totalGLTCHTR1 As Decimal = Val(dst.Tables("GLTCHTR1").Compute(
                "SUM(MATCHED_AMT)", "MATCHED = '1'") & "")
                Dim totalGLTBREC2 As Decimal = Val(dst.Tables("GLTBREC2").Compute(
                "SUM(TRAN_AMT)", "TRAN_SEL = '1'") & "")

                If totalGLTCHTR1 <> totalGLTBREC2 Then inBalance = False

                Dim hasUnmatched As Boolean = dst.Tables("GLTCHTR1").Select("MATCHED IS NULL OR MATCHED <> '1'").Length > 0
                If hasUnmatched Then inBalance = False
        End Select

        Return inBalance

    End Function

    Function GetPeriods(p As Integer) As String
        Dim sqlXX As String = ""

        For i As Integer = 1 To p
            sqlXX &= $" + NVL(ACCT_ACT_P{Format(i, "00")},0)"
        Next

        Return sqlXX
    End Function
    Private Function Selected_Rows_Offset(ByVal grd As UltraWinGrid.UltraGrid) As Boolean
        If grd.Selected.Rows.Count < 2 Then Return False ' Must select at least 2 rows

        Dim total As Decimal = 0
        For Each row As UltraWinGrid.UltraGridRow In grd.Selected.Rows
            Dim amt As Object = row.Cells("TRAN_AMT").Value
            If IsNumeric(amt) Then
                total += Convert.ToDecimal(amt)
            End If
        Next

        Return total = 0
    End Function
    Sub Export_OS_To_Excel()
        tabs.Tabs(3).Visible = True
        Dim DT As DataTable = dst.Tables("GLTBREC2").Copy()
        Dim DT_OS_Only As DataTable = DT.Select("TRAN_SEL <> '1' AND JOURNAL_TYPE = 'APCD'").CopyToDataTable()
        DT_OS_Only = DT_OS_Only.DefaultView.ToTable(False, "JOURNAL_TYPE", "TRAN_YP", "TRAN_KEY", "CUST_VEND", "TRAN_DATE", "TRAN_DESC", "TRAN_AMT")
        grdGLTOSREC.DataSource = DT_OS_Only
        If dst.Tables("GLTOSREC").Rows.Count > 0 Then
            grdGLTOSREC.DataSource = dst.Tables("GLTOSREC")
        End If

        Dim OPS_YYYYPP As String = dst.Tables("GLTBREC1").Rows(0).Item("OPS_YYYYPP").ToString()
        Dim YEAR As Integer = Integer.Parse(OPS_YYYYPP.Substring(0, 4))
        Dim MONTH As Integer = Integer.Parse(OPS_YYYYPP.Substring(4, 2))
        Dim CUR_DATE As New DateTime(YEAR, MONTH, 1)
        Dim NEXT_MONTH_DATE As DateTime = CUR_DATE.AddMonths(1)
        Dim PREV_DATE As DateTime = CUR_DATE.AddMonths(-1)
        Dim CUR_MONTH_LABEL As String = CUR_DATE.ToString("MMMM yyyy")
        Dim PREV_MONTH_LABEL As String = PREV_DATE.ToString("MMMM yyyy")
        Dim BANK_CODE As String = "CHASE"

        ' 1. Last Month Outstanding
        Dim rowLM_OS() As DataRow = dst.Tables("GLTBRECT").Select("T_LNO = 5")
        Dim LM_OS As Decimal = If(rowLM_OS.Length > 0, Math.Abs(Convert.ToDecimal(rowLM_OS(0)("T_AMT"))), 0D)

        ' 2. This Month Issued
        Dim issuedRows() As DataRow = dst.Tables("GLTBREC2").Select($"TRAN_KEY_LNO = '0' AND JOURNAL_TYPE = 'APCD' AND BANK_CODE = '{BANK_CODE}' AND TRAN_DATE >= #{CUR_DATE:MM/dd/yyyy}#")
        Dim TM_ISSUED As Decimal = issuedRows.Sum(Function(row) Math.Abs(Convert.ToDecimal(row("TRAN_AMT"))))

        ' 3. Voided Checks
        ' 3. Voided Checks (exclude ACH/wire reversals that appear on bank/API)
        Dim rowVOIDED() As DataRow = GetCountableVoidedChecks(CUR_DATE, NEXT_MONTH_DATE, BANK_CODE)

        Dim VOIDED_TOTAL As Decimal =
    If(rowVOIDED IsNot Nothing AndAlso rowVOIDED.Length > 0,
       -1D * rowVOIDED.Sum(Function(r) Convert.ToDecimal(r("TRAN_AMT"))),
       0D)

        Dim VOIDED_CHECKS As String =
    If(rowVOIDED IsNot Nothing AndAlso rowVOIDED.Length > 0,
       String.Join(", ", rowVOIDED.Select(Function(r) (r("TRAN_KEY") & "").ToString())),
       "")


        ' 4. Checks Cashed
        Dim CASHED_FILTER As String = $"BANK_CODE = '{BANK_CODE}' AND JOURNAL_TYPE = 'APCD' AND TRAN_SEL = '1'"
        Dim rowCASHED() As DataRow = dst.Tables("GLTBREC2").Select(CASHED_FILTER)
        Dim TM_CASHED As Decimal = rowCASHED.Sum(Function(r) Convert.ToDecimal(r("TRAN_AMT")))

        ' 5. Total Outstanding
        Dim TM_OS As Decimal = LM_OS + TM_ISSUED + VOIDED_TOTAL + TM_CASHED

        ' Export to Excel
        Dim workbook As New Infragistics.Documents.Excel.Workbook
        Dim worksheet As Infragistics.Documents.Excel.Worksheet = workbook.Worksheets.Add("Outstanding Checks")

        Dim headers() As String = {"Journal Type", "YP Issued", "Check No", "Vendor", "Date", "Remit", "Other"}
        For col As Integer = 0 To headers.Length - 1
            worksheet.Rows(0).Cells(col).Value = headers(col)
            worksheet.Rows(0).Cells(col).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
            worksheet.Columns(col).Width = 5000
        Next

        For row As Integer = 0 To DT_OS_Only.Rows.Count - 1
            For col As Integer = 0 To DT_OS_Only.Columns.Count - 1
                worksheet.Rows(row + 1).Cells(col).Value = DT_OS_Only.Rows(row)(col)
            Next
        Next

        Dim DT_OS_TOTAL As Decimal = DT_OS_Only.AsEnumerable().Sum(Function(row) Convert.ToDecimal(row("TRAN_AMT")))
        Dim totalRow As Integer = DT_OS_Only.Rows.Count + 1

        worksheet.Rows(totalRow).Cells(6).Value = DT_OS_TOTAL * -1
        worksheet.Rows(totalRow).Cells(6).CellFormat.FormatString = "$#,##0.00"
        worksheet.Rows(totalRow).Cells(6).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
        worksheet.Rows(totalRow).Cells(5).Value = "Total Checks Outstanding – " & CUR_MONTH_LABEL

        Dim START_ROW As Integer = DT_OS_Only.Rows.Count + 3
        Dim i As Integer = START_ROW

        worksheet.Rows(i).Cells(5).Value = PREV_MONTH_LABEL & " Outstanding Checks"
        worksheet.Rows(i).Cells(6).Value = LM_OS
        worksheet.Rows(i).Cells(6).CellFormat.FormatString = "$#,##0.00"

        i += 1 : worksheet.Rows(i).Cells(5).Value = CUR_MONTH_LABEL & " Issued Checks"
        worksheet.Rows(i).Cells(6).Value = TM_ISSUED
        worksheet.Rows(i).Cells(6).CellFormat.FormatString = "$#,##0.00"

        i += 1 : worksheet.Rows(i).Cells(5).Value = "Voided Checks"
        worksheet.Rows(i).Cells(6).Value = VOIDED_TOTAL
        worksheet.Rows(i).Cells(6).CellFormat.FormatString = "$#,##0.00_);($#,##0.00)"
        worksheet.Rows(i).Cells(7).Value = "ck " & VOIDED_CHECKS


        i += 1 : worksheet.Rows(i).Cells(5).Value = "Checks cashed in " & CUR_MONTH_LABEL
        worksheet.Rows(i).Cells(6).Value = TM_CASHED
        worksheet.Rows(i).Cells(6).CellFormat.FormatString = "$#,##0.00_);($#,##0.00)"

        i += 2
        worksheet.Rows(i).Cells(5).Value = "Proof"
        Dim proofFormula As String = $"=SUM(G{START_ROW + 1}:G{i - 1}) - G{totalRow + 1}"
        worksheet.Rows(i).Cells(6).ApplyFormula(proofFormula)
        worksheet.Rows(i).Cells(6).CellFormat.Font.Bold = Infragistics.Documents.Excel.ExcelDefaultableBoolean.True
        worksheet.Rows(i).Cells(6).CellFormat.Fill = Infragistics.Documents.Excel.CellFill.CreateSolidFill(System.Drawing.Color.LawnGreen)
        worksheet.Rows(i).Cells(6).CellFormat.FormatString = "$#,##0.00"

        Export_to_Excel_Show(workbook, "Outstanding Checks")
    End Sub

    Private Sub btnApprove_Click(sender As Object, e As EventArgs) Handles btnApprove.Click
        Dim confirmResult As DialogResult = MessageBox.Show("You are about to approve this bank reconciliation. Do you want to proceed?",
                                                         "Approve Bank Rec",
                                                         MessageBoxButtons.YesNo,
                                                         MessageBoxIcon.Question)

        If confirmResult = DialogResult.Yes Then
            If rowGLTBREC1 IsNot Nothing Then
                rowGLTBREC1("APPROVED") = "1"
                rowGLTBREC1("APPROVED_USER") = ASCMAIN1.USER_ID
                rowGLTBREC1("APPROVED_DATE") = DATETIME_STAMP

                Update_Record_TDA("GLTBREC1")

                MessageBox.Show("Bank reconciliation has been approved successfully.", "Approval Confirmed", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Error: Unable to locate the bank reconciliation record.", "Approval Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End If
    End Sub

    Private Function Can_Finalize() As Boolean

        Dim badRows = dst.Tables("GLTCHTR1").Select(
                      "(MATCHED IS NULL OR MATCHED <> '1') AND " &
                      "(RECONCILING_ITEM_IND IS NULL OR RECONCILING_ITEM_IND <> '1')")

        Return badRows.Length = 0
    End Function

    Private Function GetProofAmountFromData() As Decimal

        Dim rowLM_OS() As DataRow = dst.Tables("GLTBRECT").Select("T_LNO = 5")
        Dim LM_OS As Decimal = If(rowLM_OS.Length > 0, Math.Abs(Convert.ToDecimal(rowLM_OS(0)("T_AMT"))), 0D)

        Dim OPS_YYYYPP As String = dst.Tables("GLTBREC1").Rows(0).Item("OPS_YYYYPP").ToString()
        Dim yyyy As Integer = Integer.Parse(OPS_YYYYPP.Substring(0, 4))
        Dim mm As Integer = Integer.Parse(OPS_YYYYPP.Substring(4, 2))
        Dim curFirst As New DateTime(yyyy, mm, 1)

        Dim bank As String = "CHASE"

        Dim issuedRows() As DataRow = dst.Tables("GLTBREC2").Select(
        $"TRAN_KEY_LNO = '0' AND JOURNAL_TYPE = 'APCD' AND BANK_CODE = '{bank}' AND TRAN_DATE >= #{curFirst:MM/dd/yyyy}#")
        Dim TM_ISSUED As Decimal = issuedRows.Sum(Function(r) Math.Abs(Convert.ToDecimal(r("TRAN_AMT"))))

        Dim nextMonth As DateTime = curFirst.AddMonths(1)

        Dim voidRows() As DataRow = GetCountableVoidedChecks(curFirst, nextMonth, bank)
        Dim VOIDED_TOTAL As Decimal =
    If(voidRows IsNot Nothing AndAlso voidRows.Length > 0,
       -1D * voidRows.Sum(Function(r) Convert.ToDecimal(r("TRAN_AMT"))),
       0D)

        Dim cashedRows() As DataRow = dst.Tables("GLTBREC2").Select(
        $"BANK_CODE = '{bank}' AND JOURNAL_TYPE = 'APCD' AND TRAN_SEL = '1'")
        Dim TM_CASHED As Decimal = cashedRows.Sum(Function(r) Convert.ToDecimal(r("TRAN_AMT")))

        Dim osRows() As DataRow = dst.Tables("GLTBREC2").Select("TRAN_SEL <> '1' AND JOURNAL_TYPE = 'APCD'")
        Dim DT_OS_TOTAL As Decimal = osRows.Sum(Function(r) Convert.ToDecimal(r("TRAN_AMT")))
        Dim DT_OS_TOTAL_POS As Decimal = DT_OS_TOTAL * -1D

        Dim proof As Decimal = (LM_OS + TM_ISSUED + VOIDED_TOTAL + TM_CASHED) - DT_OS_TOTAL_POS
        Return proof
    End Function
    Private Function GetCountableVoidedChecks(curFirst As DateTime, nextMonth As DateTime, bankCode As String) As DataRow()

        Dim voidRows As DataRow() =
        dst.Tables("GLTBREC2").Select($"JOURNAL_TYPE = 'APCD' AND TRAN_KEY_LNO = 1 AND BANK_CODE = '{bankCode}'")

        ' If we don't have API transactions loaded, fall back to old behavior
        If Not dst.Tables.Contains("GLTCHTR1") OrElse dst.Tables("GLTCHTR1").Rows.Count = 0 Then
            Return voidRows
        End If

        ' Build a set of check numbers that actually appear on the bank/API side for the month
        Dim bankCheckNos As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each r As DataRow In dst.Tables("GLTCHTR1").Rows
            Dim ck As String = (r("CHECKNUMBER") & "").Trim()
            If ck = "" Then Continue For

            Dim asOfObj As Object = r("ASOFDATE")
            If asOfObj Is Nothing OrElse asOfObj Is DBNull.Value Then Continue For

            Dim asOf As DateTime = Convert.ToDateTime(asOfObj)
            If asOf >= curFirst AndAlso asOf < nextMonth Then
                bankCheckNos.Add(ck)
            End If
        Next

        ' Only count voided checks that do NOT exist on the bank/API side
        Return voidRows.
        Where(Function(v) Not bankCheckNos.Contains((v("TRAN_KEY") & "").Trim())).
        ToArray()

    End Function


End Class


Public Class JPMC_Balances_Request
    Public Property relativeDateType As String
    Public Property startDate As String
    Public Property endDate As String
    Public Property accountList As List(Of JPMC_Balances_Request2)
End Class

Public Class JPMC_Balances_Request2
    Public Property accountId As String
End Class

Public Class AccountList
    Public Property accountId As String
    Public Property accountName As String
    Public Property branchId As String
    Public Property bankId As String
    Public Property bankName As String
    Public Property currency As Currency
    Public Property balanceList As List(Of BalanceList)
End Class

Public Class BalanceList
    Public Property asOfDate As String
    Public Property recordTimestamp As DateTime
    Public Property currentDay As Boolean
    Public Property openingAvailableAmount As Double
    Public Property openingLedgerAmount As Double
    Public Property endingAvailableAmount As Double
    Public Property endingLedgerAmount As Double
End Class

Public Class Currency
    Public Property code As String
    Public Property currencySequence As Integer
    Public Property decimalLocation As Integer
    Public Property description As String
End Class

'Public Class Currency
'    Public Property code As String
'    Public Property description As String
'End Class

Public Class JPMC_Balances_Response
    Public Property accountList As List(Of AccountList)
End Class



Public Class Account
    Public Property accountId As String
    Public Property accountName As String
    Public Property bankId As String
    Public Property branchId As String
    Public Property bankName As String
    Public Property aba As String
    Public Property swift As Object
    Public Property currency As Currency
End Class

Public Class BaiType
    Public Property typeCode As String
    Public Property description As String
    Public Property btrsTypeCode As String
End Class

Public Class BankReferenceSearchable
    Public Property standardValue As String
End Class


Public Class CustomerReferenceSearchable
    Public Property standardValue As String
End Class

Public Class Datum
    Public Property account As Account
    Public Property asOfDateTime As DateTime
    Public Property valueDateTime As DateTime
    Public Property asOfDate As String
    Public Property valueDate As String
    Public Property receivedTimestamp As DateTime
    Public Property debitCreditCode As String
    Public Property baiType As BaiType
    Public Property fundsTypeCode As String
    Public Property currency As Currency
    Public Property amount As Double
    Public Property immediateAvailable As Double
    Public Property day1Available As Double
    Public Property day2Available As Double
    Public Property day2PlusAvailable As Object
    Public Property day3PlusAvailable As Double
    Public Property bankReferenceSearchable As BankReferenceSearchable
    Public Property customerReferenceSearchable As CustomerReferenceSearchable
    Public Property repairCode As String
    Public Property reversal As Boolean
    Public Property checkNumber As Integer
    Public Property wireType As String
    Public Property shortDescription As String
    Public Property postCode As String
    Public Property lockbox As Lockbox
    Public Property narrativeText As NarrativeText
    Public Property addenda As List(Of Object)
    Public Property sepaDetailsXml As Object
    Public Property supplementalTextSet As SupplementalTextSet
    Public Property supplementalTextRecordList As Object
    Public Property supplementalText As Object
    Public Property achBatchItems As Object
    Public Property transactionId As String
End Class

Public Class Lockbox
    Public Property lockboxSequenceCode As String
    Public Property lockboxItems As Double
    Public Property lockboxNumber As String
    Public Property lockboxDepositDate As Object
    Public Property lockboxDepositTime As Object
End Class

Public Class NarrativeText
    <JsonProperty("YOUR REF    ")>
    Public Property YOURREF As String
    <JsonProperty("REC FROM    ")>
    Public Property RECFROM As String
    <JsonProperty("REMARK      ")>
    Public Property REMARK As String
    <JsonProperty("REC GFP     ")>
    Public Property RECGFP As String
    <JsonProperty("B/O CUSTOMER")>
    Public Property BOCUSTOMER As String
    <JsonProperty("B/O BANK    ")>
    Public Property BOBANK As String
    <JsonProperty("CHIP SEQ    ")>
    Public Property CHIPSEQ As String
    <JsonProperty("CHIP REF    ")>
    Public Property CHIPREF As String
    <JsonProperty("ACCT PARTY  ")>
    Public Property ACCTPARTY As String
    <JsonProperty("ULTI BENE   ")>
    Public Property ULTIBENE As String
    <JsonProperty("PAID TO     ")>
    Public Property PAIDTO As String
End Class

Public Class Pagination
    Public Property pageSize As Integer
    Public Property totalPages As Integer
    Public Property pageNumber As Integer
    Public Property totalRecords As Integer
End Class

Public Class JPMC_Transactions_Response
    Public Property pagination As Pagination
    Public Property data As List(Of Datum)
End Class

Public Class SupplementalTextSet
End Class


Public Class OAuthResponse
    Public access_token As String
    Public token_type As String
    Public expires_in As Int32
End Class
