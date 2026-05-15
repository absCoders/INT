Public Class RSFCUST1
    Dim rowEDT852T1 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim rowARTCUST2 As DataRow
    Dim sqlEDT852T1 As String
    Dim EDI_DOC_SEQ_NO As String
    Dim CUST_CODE As String
    Dim CUST_STORE_NO As String

    Private importCommissionsOnly As Boolean = False
    Private createCommissionRecords As Boolean = True

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "RSFCUSTI" Then
            InquiryMode = True
        End If

        Get_PARM("EDTPARM1")

        With dst
            sqlEDT852T1 = "Select EDT852T1.*, ARTCUST1.CUST_NAME" _
            & " from EDT852T1, ARTCUST1" _
            & " where ARTCUST1.CUST_CODE = EDT852T1.CUST_CODE"
            ASCMAIN1.sql = sqlEDT852T1 _
            & " and EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "EDT852T1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = sqlEDT852T1
            Create_TDA(.Tables.Add, "EDT852TX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select RSTRETL1.*" _
                       & " from RSTRETL1 where EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "RSTRETL1", "**", 0, True, "V", 4)

            Create_TDA(.Tables.Add, "RSTRETL2", "*", 0, True)
            Create_TDA(.Tables.Add, "RSTRETLC", "*", 0, True)

            ASCMAIN1.sql = "Select X.*" _
                & ",ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_CITY, ARTCUST2.CUST_STORE_STATE" _
                & ", ARTCUST2.SREP_CODE, ARTCUST2.SELL_CODE" _
                & " from (" _
                & "Select RSTRETL1.EDI_DOC_SEQ_NO, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" _
                & ", SUM (RSTRETL1.AMT_SOLD) RETAIL_SALES" _
                & " from RSTRETL1" _
                & " where RSTRETL1.EDI_DOC_SEQ_NO = :PARM1" _
                & " group by RSTRETL1.EDI_DOC_SEQ_NO, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" _
                & ") X, ARTCUST2" _
                & " where ARTCUST2.CUST_CODE (+) = X.CUST_CODE" _
                & "   and ARTCUST2.CUST_STORE_NO (+) = X.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "RSTRETLY", "**", 0, False, "V", 3)
            .Tables("RSTRETLY").Columns.Add("QTY_SOLD", GetType(System.Int16))
            .Tables("RSTRETLY").Columns("QTY_SOLD").DefaultValue = 1

        End With

        grdEDT852TX.DataSource = dst.Tables("EDT852TX")
        grdRSTRETLY.DataSource = dst.Tables("RSTRETLY")

        Create_Summary(grdEDT852TX, "EDI_DOC_SEQ_NO", "Count")

        Create_Summary(grdRSTRETLY, "CUST_STORE_NO", "Count")
        Create_Summary(grdRSTRETLY, New String() {"RETAIL_SALES"})

        With grdRSTRETLY.DisplayLayout.Bands("RSTRETLY")
            .Columns("CUST_STORE_NO").Header.Fixed = True
            For Each COLUMN_NAME As String In New String() _
                {"CUST_STORE_NAME", "CUST_STORE_CITY", "CUST_STORE_STATE", "SREP_CODE", "SELL_CODE"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "New"
                Validate_Customer()
                Validate_Code("CUST_CODE")

                If Absx1.txtFor("OPS_YYYYPP").Text & "" = "" Then
                    EMsg &= vbCr & "Selection of a Period is Required"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("EDT852T1_C", CUST_CODE) Then Exit Sub
                End If

                If EMsg.Length = 0 Then
                    If Not EvaluateCustomerSettings(eItemKey) Then
                        Exit Sub
                    End If
                End If

            Case "Edit", "View"

                If Absx1.txtFor("EDI_DOC_SEQ_NO").Text = "" Then
                    EMsg &= vbCr & "No EDI Doc Seq No Specified"
                Else
                    EDI_DOC_SEQ_NO = Absx1.txtFor("EDI_DOC_SEQ_NO").Text
                    rowEDT852T1 = LookUp("EDT852T1", EDI_DOC_SEQ_NO)
                    If rowEDT852T1 Is Nothing Then
                        EMsg &= vbCr & "Invalid EDI Doc Seq No (" & EDI_DOC_SEQ_NO & ")"
                    Else
                        CUST_CODE = rowEDT852T1.Item("CUST_CODE")
                        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                    End If
                End If

                If EMsg = "" Then
                    EDI_DOC_SEQ_NO = Absx1.txtFor("EDI_DOC_SEQ_NO").Text
                    If eItemKey = "Edit" Then
                        If Not ASCMAIN1.Logical_Lock("EDT852T1", EDI_DOC_SEQ_NO) Then Exit Sub
                        If Not ASCMAIN1.Logical_Lock("EDT852T1_C", CUST_CODE) Then Exit Sub
                    End If
                End If

                If EMsg.Length = 0 AndAlso eItemKey = "Edit" Then
                    If Not EvaluateCustomerSettings(eItemKey) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

            Case "Cancel"
                If MsgBox("Do you want to Cancel work done on this Entry", _
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                          "Confirmation") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"
                If MsgBox("Do you want to Delete this Entry", _
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                          "Confirmation") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Import Data"
                Dim zMsg As String = "The Import file should be an Excel Workbook. The data should be placed on the first worksheet" & _
                    " and should contain the following three columns:" & Environment.NewLine & "Store No, Units, Amount$$." & _
                    Environment.NewLine & Environment.NewLine & "There should not be any header rows." & _
                    " The first row will be imported. The import will stop when it reaches a Store No cell that is blank."

                If MessageBox.Show(zMsg, "Import Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Import Data"
                ImportData()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Import Data").Settings.Enabled = iScreenMode

                    .Items("New").Visible = (Not InquiryMode)
                    .Items("Edit").Visible = (Not InquiryMode)
                    .Items("Done").Visible = Not (EntryMode = "N" Or EntryMode = "E")
                    .Items("Print").Visible = (InquiryMode Or EntryMode = "V")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Delete").Visible = (Not InquiryMode And EntryMode <> "V" And EntryMode <> "N")
                    .Items("Import Data").Visible = (Not InquiryMode And EntryMode = "N")

                    If importCommissionsOnly AndAlso Not createCommissionRecords Then
                        .Items("Update").Visible = False
                    End If

                End With
            End With
        End If

        If ScreenMode Then

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdRSTRETLY}
                If InquiryMode Or (EntryMode = "V") Then
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                Else
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                End If
            Next
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdEDT852TX.Visible = Not ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"RSTRETLY", "RSTRETL1", "RSTRETL2", "EDT852T1", "RSTRETLC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""
        Load_EDT852TX()

        importCommissionsOnly = False
        createCommissionRecords = True

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowEDT852T1 = dst.Tables("EDT852T1").NewRow

            If importCommissionsOnly = True Then
                EDI_DOC_SEQ_NO = "0000000000"
            Else
                EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDTJRNL3.EDI_DOC_SEQ_NO")
            End If
            rowEDT852T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
            rowEDT852T1.Item("CUST_CODE") = HFs("CUST_CODE")
            rowEDT852T1.Item("OPS_YYYYPP") = HFs("OPS_YYYYPP")
            rowEDT852T1.Item("COMPANY_CODE") = ASCMAIN1.SOLUTION

            Dim rowGLTPARM2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM2 WHERE OPS_YYYYPP = :PARM1", "V", New Object() {HFs("OPS_YYYYPP")})
            rowEDT852T1.Item("EDI_TO_DATE") = rowGLTPARM2.Item("PRD_END_DATE")
            rowEDT852T1.Item("EDI_FROM_DATE") = DateAdd(DateInterval.Day, -7, rowGLTPARM2.Item("PRD_END_DATE"))

            dst.Tables("EDT852T1").Rows.Add(rowEDT852T1)
        Else
            rowEDT852T1 = Fill_Record("EDT852T1", EDI_DOC_SEQ_NO)
        End If

        Fill_Records("RSTRETLY", EDI_DOC_SEQ_NO)
        Sort_grdColumns(grdRSTRETLY, "CUST_STORE_NO")

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        If importCommissionsOnly AndAlso Not createCommissionRecords Then
            Exit Sub
        End If

        Try
            BeginTrans()

            INIT_LAST("EDT852T1", True, "", True)

            If importCommissionsOnly AndAlso createCommissionRecords Then
                ASCDATA1.ExecuteSQL("Delete From RSTRETLC WHERE CUST_CODE = :PARM1 AND OPS_YYYYPP = :PARM2", "VV", New Object() {CUST_CODE, MyBase.Absx1.txtFor("OPS_YYYYPP").Text})
            End If

            If Not importCommissionsOnly Then
                ASCMAIN1.sql = "begin declare cursor c1 is select * from RSTRETL1 where CUST_CODE = '" & CUST_CODE & "' and OPS_YYYYPP = '" & MyBase.Absx1.txtFor("OPS_YYYYPP").Text & "';"
                ASCMAIN1.sql &= " begin for r1 in c1 loop"
                ASCMAIN1.sql &= "   Delete from  RSTRETL2 where CUST_CODE = r1.cust_code and OPS_YYYYPP = r1.OPS_YYYYPP;"
                ASCMAIN1.sql &= "   Delete from  RSTRETL4 where CUST_CODE = r1.cust_code and OPS_YYYYPP = r1.OPS_YYYYPP;"
                ASCMAIN1.sql &= "   Delete from  RSTRETL1 where edi_doc_seq_no = r1.edi_doc_seq_no;"
                ASCMAIN1.sql &= " End Loop; End; End;"
                ASCDATA1.ExecuteSQL()
            End If

            ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where YYYYPP = '" & Absx1.txtFor("OPS_YYYYPP").Text & "'"
            Dim OPS_YYYYWW As String = ASCDATA1.GetDataValue

            dst.Tables("RSTRETL1").Rows.Clear()
            For Each rowRSTRETLY As DataRow In dst.Tables("RSTRETLY").Select("")
                Dim rowRSTRETL1 As DataRow = dst.Tables("RSTRETL1").NewRow
                With rowRSTRETL1
                    .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    .Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
                    .Item("CUST_STORE_NO") = rowRSTRETLY.Item("CUST_STORE_NO")
                    .Item("ITEM_CODE") = "84015060"
                    .Item("QTY_SOLD") = Val(rowRSTRETLY.Item("QTY_SOLD") & String.Empty)
                    If .Item("QTY_SOLD") = 0 Then .Item("QTY_SOLD") = 1
                    .Item("AMT_SOLD") = Val(rowRSTRETLY.Item("RETAIL_SALES") & String.Empty)
                    .Item("OPS_YYYYPP") = Absx1.txtFor("OPS_YYYYPP").Text
                    .Item("OPS_YYYYWW") = OPS_YYYYWW
                    .Item("QTY_EOW") = 0
                    dst.Tables("RSTRETL1").Rows.Add(rowRSTRETL1)

                    ' Need to create RSTRETLC records if in a previous period.
                    If EntryMode = "N" AndAlso Absx1.txtFor("OPS_YYYYPP").Text < ASCMAIN1.CYP AndAlso createCommissionRecords Then
                        Dim rowRSTRETLC As DataRow = dst.Tables("RSTRETLC").NewRow
                        rowRSTRETLC.Item("CUST_CODE") = .Item("CUST_CODE")
                        rowRSTRETLC.Item("CUST_STORE_NO") = .Item("CUST_STORE_NO")
                        rowRSTRETLC.Item("OPS_YYYYPP") = .Item("OPS_YYYYPP")
                        rowRSTRETLC.Item("QTY_SOLD") = .Item("QTY_SOLD")
                        rowRSTRETLC.Item("AMT_SOLD") = .Item("AMT_SOLD")
                        rowRSTRETLC.Item("SELL_CODE") = rowRSTRETLY.Item("SELL_CODE")

                        If rowRSTRETLC.Item("SELL_CODE") & String.Empty = String.Empty Then
                            rowRSTRETLC.Item("SELL_CODE") = "98"
                        End If

                        Dim rowSOTSREP1 As DataRow = LookUp("SOTSREP1", rowRSTRETLC.Item("SELL_CODE") & String.Empty)
                        If rowSOTSREP1 IsNot Nothing Then
                            rowRSTRETLC.Item("SELL_COMM_PCT") = Val(rowSOTSREP1.Item("SELL_COMM_PCT") & String.Empty)
                        Else
                            rowRSTRETLC.Item("SELL_COMM_PCT") = 0
                        End If

                        dst.Tables("RSTRETLC").Rows.Add(rowRSTRETLC)
                    End If
                End With
            Next

            If Not importCommissionsOnly Then
                Update_RSTRETLx(EDI_DOC_SEQ_NO, "-")

                Update_Record_TDA("EDT852T1")
                Update_Record_TDA("RSTRETL1", "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")

                Update_RSTRETLx(EDI_DOC_SEQ_NO, "+")
            End If

            If createCommissionRecords Then
                Update_Record_TDA("RSTRETLC")
            End If

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

    Sub Delete_Record()
        If EntryMode = "N" Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Deleting Record")
        BeginTrans()

        Update_RSTRETLx(EDI_DOC_SEQ_NO, "-")
        ASCDATA1.ExecuteSQL("Delete from EDT852T1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
        ASCDATA1.ExecuteSQL("Delete from RSTRETL1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")

        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                sql_where = "OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' and OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"

        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("EDI_DOC_SEQ_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT852TX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdRSTRETLY, "S", "Allow New/Edit to Stores")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Allow New/Edit to Stores"
                tlb_sbt = DirectCast(tlb.Tools("Allow New/Edit to Stores"), UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    grdRSTRETLY.Tag = "Y"
                Else
                    grdRSTRETLY.Tag = ""
                End If
                With grdRSTRETLY.DisplayLayout.Bands(0)
                    For Each COLUMN_NAME As String In New String() _
                        {"CUST_STORE_NAME", "CUST_STORE_CITY", "CUST_STORE_STATE", "SREP_CODE", "SELL_CODE"}
                        If tlb_sbt.Checked Then
                            .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                        Else
                            .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                        End If
                    Next
                End With
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Job Order Inquiry"
            '    Dim JOB_NO As String = grd.ActiveRow.Cells("JOB_NO").Text
            '    Context_Launch("Load", JOB_NO, e.Tool.Key, "DEFJOBMI")

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Click_Command("Load", e)
                    Load_EDT852TX()
                End If

            Case "EDI_DOC_SEQ_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If InquiryMode Then
                        Click_Command("Load", e)
                    Else
                        Click_Command("Edit", e)
                    End If
                End If

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                If InquiryMode Then
                    'Click_Command("Load")
                Else
                    'Click_Command("New")
                End If
                Load_EDT852TX()

            Case "EDI_DOC_SEQ_NO"
                If InquiryMode Then
                    Click_Command("Load")
                Else
                    Click_Command("Edit")
                End If

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdRSTRETLY"

    Private Sub grdRSTRETLY_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETLY.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_STORE_NO"
                Dim CUST_STORE_NO As String = e.Cell.Value & ""
                Dim rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                If rowARTCUST2 IsNot Nothing Then
                    e.Cell.Row.Cells("CUST_STORE_NAME").Value = rowARTCUST2.Item("CUST_STORE_NAME") & ""
                    e.Cell.Row.Cells("CUST_STORE_CITY").Value = rowARTCUST2.Item("CUST_STORE_CITY") & ""
                    e.Cell.Row.Cells("CUST_STORE_STATE").Value = rowARTCUST2.Item("CUST_STORE_STATE") & ""
                    e.Cell.Row.Cells("SREP_CODE").Value = rowARTCUST2.Item("SREP_CODE") & ""
                    e.Cell.Row.Cells("SELL_CODE").Value = rowARTCUST2.Item("SELL_CODE") & ""
                End If
        End Select
    End Sub

    Private Sub grdRSTRETLY_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTRETLY.AfterExitEditMode

        With grdRSTRETLY
            If .ActiveCell Is Nothing Then Exit Sub
            Select Case .ActiveCell.Column.Key

            End Select
        End With
    End Sub

    Private Sub grdRSTRETLY_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdRSTRETLY.AfterRowActivate
        With grdRSTRETLY.DisplayLayout.Bands(0)

        End With

        With grdRSTRETLY.ActiveRow
            If .Cells("CUST_STORE_NO").Value & "" = "" Then
                grdRSTRETLY.ActiveCell = .Cells("CUST_STORE_NO")
            End If
        End With

    End Sub

    Private Sub grdRSTRETLY_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdRSTRETLY.AfterRowsDeleted

    End Sub

    Private Sub grdRSTRETLY_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdRSTRETLY.AfterRowUpdate

    End Sub

    Private Sub grdRSTRETLY_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdRSTRETLY.BeforeExitEditMode

        If e.CancellingEditOperation Then Exit Sub
        With grdRSTRETLY.ActiveCell
            Select Case .Column.Key
                Case "CUST_STORE_NO"
                    If .EditorResolved.IsValid Then
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                        Dim CUST_STORE_NO As String = .EditorResolved.Value & ""
                        Dim rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                        If rowARTCUST2 Is Nothing Then
                            If grdRSTRETLY.Tag = "Y" Then
                            Else
                                e.Cancel = True
                            End If
                        End If
                    End If                
                Case Else
                    ' e.Cancel = Validate_Columns_5(.Column.Key)
            End Select
        End With
    End Sub

    Private Sub grdRSTRETLY_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdRSTRETLY.BeforeRowsDeleted

    End Sub

    Private Sub grdRSTRETLY_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTRETLY.BeforeRowUpdate

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("EDI_DOC_SEQ_NO").Value = EDI_DOC_SEQ_NO
            e.Row.Cells("CUST_CODE").Value = CUST_CODE
        End If
    End Sub

    Private Sub grdRSTRETLY_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETLY.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdRSTRETLY.ActiveCell.Column.Key
            Case "CUST_STORE_NO"
                sql_where = "CUST_CODE = '" & CUST_CODE & "'"
                grdClickCellButton(grdRSTRETLY, sql_where, False)
            Case Else
                grdClickCellButton(grdRSTRETLY, sql_where, False)
        End Select
    End Sub
#End Region

    Sub Load_EDT852TX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

        ASCMAIN1.sql = sqlEDT852T1
        If CUST_CODE = "" Then
            grdEDT852TX.Text = "All Retail Sales"
        Else
            ASCMAIN1.sql &= " and EDT852T1.CUST_CODE = '" & CUST_CODE & "'"
            grdEDT852TX.Text = "Retail Sales for Customer " & CUST_CODE
        End If

        Fill_Records("EDT852TX", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdEDT852TX, "EDI_DOC_SEQ_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        ' tabPO.SelectedTab = tabPO.Tabs("Open POs")
    End Sub

    Function Validate_Customer() As Boolean
        CUST_CODE = String.Empty
        rowARTCUST1 = Nothing

        If Absx1.txtFor("CUST_CODE").Text = "" Then
            Return False
        End If

        rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)

        If rowARTCUST1 Is Nothing Then
            EMsg &= vbCr & "Customer is Not on File" & vbCrLf
        Else
            If rowARTCUST1.Item("CUST_STATUS") & "" <> "A" Then
                EMsg &= vbCr & "Customer Status is not Active" & vbCrLf
            End If
        End If

        If EMsg = "" Then
            CUST_CODE = rowARTCUST1.Item("CUST_CODE")
        End If

        Return (CUST_CODE <> "")
    End Function

    Private Sub grdEDT852TX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT852TX.DoubleClickRow
        Absx1.txtFor("EDI_DOC_SEQ_NO").Text = e.Row.Cells("EDI_DOC_SEQ_NO").Value
        Click_Command("View")
    End Sub

    Private Sub grdRSTRETLY_Error(sender As Object, e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdRSTRETLY.Error
        e.Cancel = True
    End Sub

    Sub Update_RSTRETLx(ByVal EDI_DOC_SEQ_NO As String, Optional plus_or_minus As String = "+")

        Dim sql As String = ""

        sql = "BEGIN DECLARE CURSOR C1 IS"
        sql = sql & " SELECT RSTRETL1.OPS_YYYYPP, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO,"
        sql = sql & "  ICTITEM1.COLLECTION_CODE, SUM (RSTRETL1.AMT_SOLD) AS RETAIL_SALES"
        sql = sql & "  From RSTRETL1, ICTITEM1"
        sql = sql & "  Where RSTRETL1.ITEM_CODE = ICTITEM1.ITEM_CODE (+) "
        sql = sql & "    and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        sql = sql & "  GROUP BY RSTRETL1.OPS_YYYYPP, RSTRETL1.CUST_CODE, "
        sql = sql & "           RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE"
        sql = sql & "  HAVING SUM (RSTRETL1.AMT_SOLD) <> 0;"
        sql = sql & " BEGIN FOR R1 IN C1 LOOP"
        sql = sql & " Update RSTRETL2"
        sql = sql & " Set RETAIL_SALES = NVL(RETAIL_SALES, 0) " & plus_or_minus & " NVL(R1.RETAIL_SALES,0)"
        sql = sql & "  Where OPS_YYYYPP = R1.OPS_YYYYPP And CUST_CODE = R1.CUST_CODE      "
        sql = sql & "    AND CUST_STORE_NO = R1.CUST_STORE_NO AND COLLECTION_CODE = R1.COLLECTION_CODE;"
        sql = sql & " IF SQL%NOTFOUND THEN"
        sql = sql & "  INSERT INTO RSTRETL2 VALUES"
        sql = sql & "   (R1.OPS_YYYYPP, R1.CUST_CODE, R1.CUST_STORE_NO, R1.COLLECTION_CODE, " & plus_or_minus & "1 * " & "NVL(R1.RETAIL_SALES,0));"
        sql = sql & " END IF;"
        sql = sql & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

        sql = "BEGIN DECLARE CURSOR C1 IS" _
        & "  SELECT RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE, " _
        & "   ICTCOLL1.BRAND_CODE," _
        & "   NULL SREP_CODE, NULL SELL_CODE," _
        & "   RSTRETL1.OPS_YYYYWW, RSTRETL1.OPS_YYYYPP" _
        & ",  SUM(RSTRETL1.AMT_SOLD) AS AMT_SOLD" _
        & ",  SUM(RSTRETL1.QTY_SOLD) AS QTY_SOLD" _
        & ",  SUM(RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE) QTY_EOW" _
        & "   From RSTRETL1, ICTITEM1, ICTCOLL1" _
        & "   Where RSTRETL1.ITEM_CODE = ICTITEM1.ITEM_CODE (+) " _
        & "   AND ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE (+)" _
        & "    and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & "   GROUP BY RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, ICTITEM1.COLLECTION_CODE, " _
        & "   ICTCOLL1.BRAND_CODE, " _
        & "   NULL, NULL, RSTRETL1.OPS_YYYYWW, RSTRETL1.OPS_YYYYPP" _
        & "   HAVING SUM (RSTRETL1.AMT_SOLD) <> 0 OR SUM (RSTRETL1.QTY_SOLD) <> 0 OR SUM(QTY_EOW) <> 0;" _
        & "  BEGIN FOR R1 IN C1 LOOP" _
        & "  Update RSTRETL4" _
        & " Set AMT_SOLD = NVL(AMT_SOLD, 0) " & plus_or_minus & " R1.AMT_SOLD," _
        & "    QTY_SOLD = NVL(QTY_SOLD, 0) " & plus_or_minus & " R1.QTY_SOLD," _
        & "    QTY_EOW = NVL(QTY_EOW, 0) " & plus_or_minus & " R1.QTY_EOW" _
        & "   Where  CUST_CODE = R1.CUST_CODE AND CUST_STORE_NO = R1.CUST_STORE_NO AND COLLECTION_CODE = R1.COLLECTION_CODE" _
        & "     And OPS_YYYYWW = R1.OPS_YYYYWW And OPS_YYYYPP = R1.OPS_YYYYPP;" _
        & "  IF SQL%NOTFOUND THEN" _
        & "   INSERT INTO RSTRETL4 VALUES" _
        & "    (R1.CUST_CODE, R1.CUST_STORE_NO, R1.COLLECTION_CODE, R1.OPS_YYYYWW, R1.OPS_YYYYPP, " _
        & plus_or_minus & "1 * R1.QTY_SOLD, " & plus_or_minus & "1 * R1.AMT_SOLD, R1.BRAND_CODE, R1.SELL_CODE, R1.SREP_CODE, R1.QTY_EOW);" _
        & "  END IF;" _
        & "  END LOOP; END; END;"
        ASCDATA1.ExecuteSQL(sql)

    End Sub

    Private Sub ImportData()

        Try

            Dim fileName As String = String.Empty
            Dim excelSheet As String = String.Empty

            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                openFileDialog1.Filter = "xls files (*.xls)|*.xls"
                openFileDialog1.RestoreDirectory = True

                If openFileDialog1.ShowDialog() = DialogResult.OK Then
                    fileName = openFileDialog1.FileName
                End If
            End Using

            If fileName.Length = 0 Then
                Exit Sub
            End If

            If Not My.Computer.FileSystem.FileExists(fileName) Then
                MessageBox.Show("Invalid File selected", "Import Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Try
                Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" & _
                    "data source=" & fileName & ";" & _
                    "Extended Properties=""Excel 8.0;HDR=No;IMEX=1"""
                Dim SHEETS As Int32 = 0
                Dim dbSchema As DataTable

                Using objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                    objConnection.Open()

                    dbSchema = objConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                    If dbSchema.Rows.Count > 0 Then
                        excelSheet = dbSchema.Rows(0).Item("TABLE_NAME")
                    End If
                    SHEETS = dbSchema.Rows.Count

                    If SHEETS = 0 Then
                        MessageBox.Show("No data to Import", "Import Data", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim strSQL As String = "SELECT * FROM [" & excelSheet & "]"
                    'Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
                    Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
                    Dim dt As New DataTable
                    objAdapter.Fill(dt)
                    objConnection.Close()

                    Dim STORE As String = ""

                    For R As Int32 = 0 To dt.Rows.Count - 1
                        Dim row As DataRow = dt.Rows(R)
                        STORE = row.Item(0) & String.Empty
                        If STORE = String.Empty Then
                            Continue For
                        End If
                        STORE = ASCMAIN1.Format_Field(STORE, "CUST_STORE_NO")

                        ' If this is an edit then edit existing retail sales.
                        If dst.Tables("RSTRETLY").Select("CUST_STORE_NO = '" & CUST_STORE_NO & "'").Length > 0 Then
                            dst.Tables("RSTRETLY").Select("CUST_STORE_NO = '" & CUST_STORE_NO & "'")(0).Item("QTY_SOLD") = Convert.ToInt16(row.Item(1))
                            dst.Tables("RSTRETLY").Select("CUST_STORE_NO = '" & CUST_STORE_NO & "'")(0).Item("RETAIL_SALES") = row.Item(2)
                            Continue For
                        End If

                        grdRSTRETLY.DisplayLayout.Bands(0).AddNew()
                        With grdRSTRETLY.ActiveRow
                            .Cells("CUST_STORE_NO").Value = STORE
                            .Cells("QTY_SOLD").Value = Convert.ToInt16(row.Item(1))
                            .Cells("RETAIL_SALES").Value = row.Item(2)
                            .Update()
                        End With

                        If R Mod 10 = 0 Then
                            ASCMAIN1.Progress("-", CStr(R))
                        End If
                    Next

                End Using

            Catch ex As Exception
                MsgBox("Exception Occurred:" & vbCrLf & ex.Message, MsgBoxStyle.OkOnly, "Error Opening Excel Workbook")
            Finally

            End Try


        Catch ex As Exception
            MessageBox.Show("Error Importing Data: " & ex.Message)
        End Try
    End Sub

    Private Function EvaluateCustomerSettings(ByVal eItemKey As String) As Boolean

        Try
            Dim sql As String = String.Empty
            Dim OPS_YYYYPP As String = Absx1.txtFor("OPS_YYYYPP").Text

            If ASCMAIN1.Period_Diff(ASCMAIN1.CYP, OPS_YYYYPP) > 0 Then
                MessageBox.Show("The selected Period is in the future. You cannot proceed.", "Enter Retail Sales", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
             End If

            ' Initialize variables
            importCommissionsOnly = False
            createCommissionRecords = True

            If rowARTCUST1.Item("CUST_EDI_COMM_SEP") & String.Empty = "1" Then
                importCommissionsOnly = True
            Else
                ' Do not allow editting EDI Retail Sales Data. Item level data will be lost.
                sql = "select * from EDT852T1 where CUST_CODE = :PARM1 AND OPS_YYYYPP = :PARM2"
                If ASCDATA1.GetDataTable(sql, "", "VV", New Object() {CUST_CODE, OPS_YYYYPP}).Rows.Count > 0 Then
                    MessageBox.Show("The selected Customer (" & CUST_CODE & ") is and EDI customer. There are EDI Retail Sales imported " _
                        & " for the selected period (" & OPS_YYYYPP & "); therefore you cannot proceed.", "Enter Retail Sales", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return False
                End If
            End If

            ' If editting an entry then only update retail sales.
            If eItemKey = "Edit" AndAlso importCommissionsOnly Then
                createCommissionRecords = False
                importCommissionsOnly = False
            End If


            If createCommissionRecords Then
                sql = "Select distinct SELL_COMM_XNO from RSTRETLC Where SELL_COMM_XNO is not null and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                sql &= " Union"
                sql &= " select distinct SREP_COMM_XNO from SOTSCOMO Where SREP_COMM_XNO is not null and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                sql &= " Union"
                sql &= " select distinct SREP_COMM_XNO from SOTSCOM1 Where VOUCHER_NO is not null and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                If ASCDATA1.GetDataTable(sql).Rows.Count > 0 Then
                    If MessageBox.Show("The selected period has already been finalized. Commission entries will not be made." _
                                        & Environment.NewLine & Environment.NewLine & "Do you want to continue?", "Enter Retail Sales", _
                                         MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Return False
                    End If

                    createCommissionRecords = False
                ElseIf ASCDATA1.GetDataTable("Select * from RSTRETLC WHERE OPS_YYYYPP = '" & OPS_YYYYPP & "' AND CUST_CODE = '" & CUST_CODE & "'").Rows.Count > 0 Then
                    If MessageBox.Show("There are Retails Sales Commissions calculated for Customer (" & CUST_CODE & ") and Period (" & Absx1.txtFor("OPS_YYYYPP").Text & "). These will be replaced when you Update your Import." _
                                        & Environment.NewLine & Environment.NewLine _
                                        & "Do you want to continue?", "Enter Retail Sales", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Return False
                    End If
                End If
            Else
                If MessageBox.Show("Commissions will not be recalculated for Customer (" & CUST_CODE & ") and Period (" & Absx1.txtFor("OPS_YYYYPP").Text & ")." _
                                    & Environment.NewLine & Environment.NewLine _
                                    & "Do you want to continue?", "Enter Retail Sales", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Return False
                End If
            End If

            If importCommissionsOnly AndAlso Not createCommissionRecords Then
                EMsg &= vbCr & "The selected Customer is setup to create only Commission Records from the Retail sales data you will provide." _
                    & " However, the selected period has already been finalized and commission entries cannot be made."
                Return True
            End If

            If importCommissionsOnly Then
                If MessageBox.Show("The selected Customer is setup to create only Commission Records from the Retail Sales data you will provide." _
                                   & " No Retail Sales data will be saved." _
                                   & Environment.NewLine & Environment.NewLine _
                                   & "Do you want to continue?", "Enter Retail Sales", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Return False
                End If
            End If

            Return True

        Catch ex As Exception
            EMsg &= vbCr & ex.Message
            Return True
        End Try
    End Function

End Class