Imports ABSolution

Public Class SOFALLO1

    ' UPDATING SOTORDR2 = AND GETTING ROWS CODED TO ALLO_CTL_NO REGARDLESS OF DATE
    'SOTORDR1.ORDR_DATE_REL S/B DATE SHIPPED
    ' START END END DATES WIDER

    Const maxAlloations As Integer = 99
    Dim iColumn As Integer
    Dim ALLO_CTL_NOs As String
    Dim ALLO_CTL_NOs_to_Delete As New List(Of String)
    Dim ALLO_CTL_NOi() As String
    Dim COLLECTION_CODE As String
    Dim ALLO_CTL_NO_to_copy As String
    Dim CUST_CODE_to_copy As String
    Dim sql_ICTITEM1 As String
    Dim col_ICTITEM1 As New List(Of String)
    Dim rowARTCUST1 As DataRow
    Dim ALLO_CTL_NO_new As New List(Of String)
    Dim ITEM_CODE_new As New List(Of String)
    Dim iCol As New Dictionary(Of String, Int64)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        sql_ICTITEM1 = ", ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_STATUS" & vbCrLf _
                & ", ICTITEM1.PROD_CODE, ICTITEM1.ITEM_NOT_ALLOCATED, ICTITEM1.COLLECTION_CODE" & vbCrLf
        For Each COLUMN_NAME In Split(Replace(Replace(sql_ICTITEM1, " ICTITEM1.", ""), vbCrLf, ""), ",")
            col_ICTITEM1.Add(COLUMN_NAME)
        Next

        With dst

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & sql_ICTITEM1 _
                & ", SOTALLO1.ALLO_CTL_NO, SOTALLO1.DATE_START, SOTALLO1.DATE_END" & vbCrLf _
                & ", SOTALLO1.INIT_OPER, SOTALLO1.INIT_DATE, SOTALLO1.LAST_OPER, SOTALLO1.LAST_DATE" & vbCrLf _
                & " from SOTALLO1, ICTITEM1" & vbCrLf _
                & " where SOTALLO1.ITEM_CODE = ICTITEM1.ITEM_CODE"
            MyBase.Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, String.Empty, 1)
            .Tables("ICTITEM1").Columns.Add("SELECTED")
            .Tables("ICTITEM1").Columns("SELECTED").DefaultValue = "0"
            .Tables("ICTITEM1").Columns("ALLO_CTL_NO").AllowDBNull = True

            ASCMAIN1.sql = "Select SOTALLO1.*" & sql_ICTITEM1 _
                & " from SOTALLO1, ICTITEM1" & vbCrLf _
                & " where SOTALLO1.ITEM_CODE = ICTITEM1.ITEM_CODE"
            MyBase.Create_TDA(.Tables.Add, "SOTALLOX", "**", 0, False, String.Empty, 1)
            .Tables("SOTALLOX").Columns.Add("SELECTED")
            .Tables("SOTALLOX").Columns("SELECTED").DefaultValue = "0"
            .Tables("SOTALLOX").Columns("ALLO_CTL_NO").AllowDBNull = True

            ASCMAIN1.sql = "Select ICTCOLL1.*, ICTBRAN1.BRAND_NAME, ICTCOLL0.HC_NAME" & vbCrLf _
                & " from ICTCOLL1, ICTBRAN1, ICTCOLL0" & vbCrLf _
                & " where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and ICTCOLL0.HC_CODE = ICTCOLL1.HC_CODE"
            MyBase.Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTALLO1.*" & sql_ICTITEM1 _
                & " from SOTALLO1,ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTALLO1", "**", 0)

            ASCMAIN1.sql = "Select SOTALLO2.*, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE" & vbCrLf _
                & " from SOTALLO2, ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTALLO2", "**", 0)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, SREP_CODE, TRADE_CLASS_CODE from ARTCUST1"
            MyBase.Create_TDA(.Tables.Add, "SOTALLOC", "**", 0, False, String.Empty, 1)
            For iCtr As Integer = 1 To maxAlloations
                .Tables("SOTALLOC").Columns.Add("ALLO_" & Format(iCtr, "00"), GetType(System.Int64))
            Next

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                & ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DATE_REL" & vbCrLf _
                & " from SOTORDR1,SOTORDR2 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "  and SOTORDR2.ITEM_CODE = :PARM1 and SOTORDR1.ORDR_SHIP_DATE >= :PARM2 and SOTORDR1.ORDR_SHIP_DATE <= :PARM3"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "VDD", 2)
            .Tables("SOTORDR2").Columns.Add("SELECTED")
            .Tables("SOTORDR2").Columns("SELECTED").DefaultValue = "0"



            ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_OPEN" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_PICK" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_SHIP" & vbCrLf _
                & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
                & "   where SOTORDR2.ALLO_CTL_NO = :PARM1" & vbCrLf _
                & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"
            Create_TDA(.Tables.Add, "SOTALLOZ", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select ICTSTAT2.ITEM_CODE" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_ONPO) WHSE_QTY_ONPO" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_PLAN) WHSE_QTY_PLAN" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_OPEN) WHSE_QTY_OPEN" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
                 & ", SUM (ICTSTAT2.WHSE_QTY_COMM) WHSE_QTY_COMM" & vbCrLf _
                 & "  from ICTSTAT2" & vbCrLf _
                 & " where ICTSTAT2.ITEM_CODE = :PARM1" & vbCrLf _
                 & " group by ICTSTAT2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTSTAT2", "**", 0, False, "V", 1)


            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1" _
                & " where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "V", 1)


        End With

        grdSOTALLOX.DataSource = dst.Tables("SOTALLOX")
        grdICTITEM1.DataSource = dst.Tables("ICTITEM1")
        grdICTCOLL1.DataSource = dst.Tables("ICTCOLL1")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        Create_Summary(grdSOTALLOX, "ITEM_CODE", "Count")
        Create_Summary(grdSOTALLOX, "SELECTED", "Sum")

        Create_Summary(grdICTITEM1, "ITEM_CODE", "Count")
        Create_Summary(grdICTITEM1, "SELECTED", "Sum")

        grdSOTALLO1.DataSource = dst.Tables("SOTALLO1")
        grdSOTALLOC.DataSource = dst.Tables("SOTALLOC")

        grdSOTALLOX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdSOTALLOX.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
        grdSOTALLOX.DisplayLayout.Bands(0).Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        grdSOTALLOX.DisplayLayout.Bands(0).Columns("SELECTED").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdSOTALLOX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key = "ITEM_CODE" Or col_ICTITEM1.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"DATE_START", "DATE_END", "ALLO_CTL_NO", "ALLOW_OVER"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        grdICTITEM1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdICTITEM1.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
        grdICTITEM1.DisplayLayout.Bands(0).Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
        grdICTITEM1.DisplayLayout.Bands(0).Columns("SELECTED").CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        With grdICTITEM1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SELECTED" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key = "ITEM_CODE" Or col_ICTITEM1.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"SELECTED"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf New String() {"DATE_START", "DATE_END", "ALLO_CTL_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With


        With grdSOTALLO1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "DATE_START" Or GCOL.Key = "DATE_END" Or GCOL.Key = "ALLOW_OVER" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                    GCOL.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                GCOL.Header.Appearance.BackColor2 = Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If GCOL.Key = "ITEM_CODE" Or col_ICTITEM1.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf New String() {"DATE_START", "DATE_END", "ALLO_CTL_NO", "ALLOW_OVER"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                Else
                    GCOL.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdSOTALLOC.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_CODE", "SREP_CODE", "TRADE_CLASS_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
            For iCtr As Integer = 1 To maxAlloations
                Dim COLUMN_NAME As String = "ALLO_" & Format(iCtr, "00")
                With .Columns(COLUMN_NAME)
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                End With
                Create_Summary(grdSOTALLOC, COLUMN_NAME)
            Next
        End With

        Fill_Records("ICTCOLL1")
        grdICTCOLL1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect
        grdICTCOLL1.DisplayLayout.AutoFitStyle = UltraWinGrid.AutoFitStyle.ExtendLastColumn
        With grdICTCOLL1.DisplayLayout.Bands(0)
            .ColHeadersVisible = False
            .SortedColumns.Add("BRAND_NAME", False, True)
            .Columns("BRAND_NAME").HiddenWhenGroupBy = DefaultableBoolean.True
            .Columns("HC_NAME").Header.Caption = "High Collection"
            ' .SortedColumns(1).Header.Caption = ""
            .SortedColumns.Add("HC_NAME", False, True)
            .Columns("HC_NAME").HiddenWhenGroupBy = DefaultableBoolean.True
            .SortedColumns.Add("COLLECTION_NAME", False)
        End With
        grdICTCOLL1.Rows.ExpandAll(False)

        ASCMAIN1.Add_Value_List(grdSOTALLOX, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdSOTALLOX, "ITEM_STATUS")
        ASCMAIN1.Add_Value_List(grdICTITEM1, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdICTITEM1, "ITEM_STATUS")

        ASCMAIN1.Add_Value_List(grdSOTALLO1, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdSOTALLO1, "ITEM_STATUS")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If dst.Tables("SOTALLOX").Select("SELECTED = '1'").Length = 0 _
                    And dst.Tables("ICTITEM1").Select("SELECTED = '1'").Length = 0 Then
                    EMsg = "There are no Items or Allocations Selected"
                End If

                If dst.Tables("SOTALLOX").Select("SELECTED = '1'").Length + dst.Tables("ICTITEM1").Select("SELECTED = '1'").Length > maxAlloations Then
                    EMsg = "You may load a maximum of " & CStr(maxAlloations) & " Items/Allocations"
                End If

                If EMsg = "" Then
                    For Each row As DataRow In dst.Tables("SOTALLOX").Select("SELECTED = '1'")
                        Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")
                        If Not ASCMAIN1.Logical_Lock("SOTALLO1", ALLO_CTL_NO) Then
                            Exit Sub
                        End If
                        Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                        If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE) Then
                            Exit Sub
                        End If
                    Next
                    For Each row As DataRow In dst.Tables("ICTITEM1").Select("SELECTED = '1'")
                        Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                        If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE) Then
                            Exit Sub
                        End If
                    Next
                End If

            Case "Cancel"
                If MsgBox("Are you sure you want to Cancel your changes?", _
                            MsgBoxStyle.YesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"

                grdSOTALLO1.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdSOTALLOC.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

                Dim dateError As Boolean = False
                For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("", "ITEM_CODE,ALLO_CTL_NO")
                    Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")
                    Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")

                    dateError = False
                    If Not IsDate(rowSOTALLO1.Item("DATE_START") & String.Empty) Then
                        EMsg &= Environment.NewLine & "Missing or Invalid Start Date for Allocation: " & rowSOTALLO1.Item("ALLO_CTL_NO")
                        dateError = True
                    End If

                    If Not IsDate(rowSOTALLO1.Item("DATE_END") & String.Empty) Then
                        EMsg &= Environment.NewLine & "Missing or Invalid End Date for Allocation: " & rowSOTALLO1.Item("ALLO_CTL_NO")
                        dateError = True
                    End If

                    If Not dateError Then
                        If DateDiff(DateInterval.Day, rowSOTALLO1.Item("DATE_START"), rowSOTALLO1.Item("DATE_END")) < 0 Then
                            EMsg &= Environment.NewLine & "End Date is prior to Start Date for Allocation: " & rowSOTALLO1.Item("ALLO_CTL_NO")
                        Else
                            ASCMAIN1.sql = "Select * from SOTALLO1" & vbCrLf _
                                & " where ITEM_CODE = :PARM1" & vbCrLf _
                                & " and ALLO_CTL_NO <> :PARM2" & vbCrLf _
                                & " and (:PARM3 between DATE_START and DATE_END or :PARM4 between DATE_START and DATE_END)"
                            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVDD", _
                                                                     New Object() {ITEM_CODE, _
                                                                                   ALLO_CTL_NO, _
                                                                                   rowSOTALLO1.Item("DATE_START"), _
                                                                                   rowSOTALLO1.Item("DATE_END")})
                            If row IsNot Nothing Then
                                EMsg &= vbCr & "Allocation " & ALLO_CTL_NO & " for Item " & ITEM_CODE & vbCrLf _
                                    & " (Date Range " & Format(rowSOTALLO1.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(rowSOTALLO1.Item("DATE_END"), "MM/dd/yyyy") & ")" & vbCrLf _
                                    & " is within Date Range of Allocation " & row.Item("ALLO_CTL_NO") & vbCrLf _
                                    & " (Date Range " & Format(row.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(row.Item("DATE_END"), "MM/dd/yyyy") & ")"
                            Else
                                For Each row2 As DataRow In dst.Tables("SOTALLO1").Select("ITEM_CODE = '" & ITEM_CODE & "' and ALLO_CTL_NO <> '" & ALLO_CTL_NO & "' and DATE_START is Not Null and DATE_END is Not Null")
                                    Dim DATE_START_ymd As String = Format(rowSOTALLO1.Item("DATE_START"), "yyyyMMdd")
                                    Dim DATE_END_ymd As String = Format(rowSOTALLO1.Item("DATE_END"), "yyyyMMdd")
                                    If DATE_START_ymd >= Format(row2.Item("DATE_START"), "yyyyMMdd") And DATE_START_ymd <= Format(row2.Item("DATE_END"), "yyyyMMdd") _
                                    Or DATE_END_ymd >= Format(row2.Item("DATE_START"), "yyyyMMdd") And DATE_END_ymd <= Format(row2.Item("DATE_END"), "yyyyMMdd") Then
                                        EMsg &= vbCr & "Allocation " & ALLO_CTL_NO & " for Item " & ITEM_CODE & vbCrLf _
                                              & " (Date Range " & Format(rowSOTALLO1.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(rowSOTALLO1.Item("DATE_END"), "MM/dd/yyyy") & ")" & vbCrLf _
                                              & " is within Date Range of Allocation " & row2.Item("ALLO_CTL_NO") & vbCrLf _
                                              & " (Date Range " & Format(row2.Item("DATE_START"), "MM/dd/yyyy") & " thru " & Format(row2.Item("DATE_END"), "MM/dd/yyyy") & ")"

                                    End If


                                Next
                            End If
                        End If
                    End If
                Next

                For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select
                    If LookUp("ARTCUST1", rowSOTALLOC.Item("CUST_CODE")) Is Nothing Then
                        EMsg &= Environment.NewLine & "Invalid Customer Code: " & rowSOTALLOC.Item("CUST_CODE")
                    End If
                Next
        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                MyBase.EntryMode = "E"
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Cancel"
                Me.Mode_Settings(False)

            Case "Print"
                Me.Print_Record()

            Case "Update"
                Me.Update_Record()
                Me.Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Save Orig").Settings.Enabled = iScreenMode
                End With
                '.Groups("Items").Visible = Not ScreenMode
                lblITEM_CODE.Visible = Not ScreenMode
                txtITEM_CODE.Visible = Not ScreenMode
                .Groups("Customers").Visible = ScreenMode
                .Groups("Items").Visible = Not ScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        splSOTALLOX.Visible = Not ScreenMode
        splSOTALLO1.Visible = ScreenMode

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTALLO1, grdSOTALLOC}
                With grd.DisplayLayout.Override
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        If grd.Name = "grdSOTALLOC" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        End If
                        '  .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.True
                        .AllowDelete = DefaultableBoolean.False
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowUpdate = DefaultableBoolean.False
                        .AllowDelete = DefaultableBoolean.False
                    End If
                End With
            Next
        End If

        spl.Panel1Collapsed = True ' Not ScreenMode

        If ScreenMode Then
            grdSOTALLOX.Parent = tabSOTALLOC.Tabs("Other Allocations").TabPage
            grdICTITEM1.Parent = tabSOTALLOC.Tabs("Other Items").TabPage
            grdSOTALLOX.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True
            grdICTITEM1.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True

            'grdSOTALLO1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Else
            Me.Clear_Record()
            grdSOTALLOX.Parent = tabSOTALLOX.Tabs("Allocations").TabPage
            grdICTITEM1.Parent = tabSOTALLOX.Tabs("Items").TabPage
            grdSOTALLOX.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False
            grdICTITEM1.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = False

            'grdSOTALLO1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.RowSelect
        End If

    End Sub

    Private Sub Clear_Record()

        MyBase.EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTALLO1", "SOTALLO2", "SOTALLOC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Load_SOTALLOX()
        MyBase.EnforceConstraints(True)

        For i As Integer = 1 To maxAlloations
            With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(i, "00"))
                .Header.Caption = ""
                .Hidden = True
                .Tag = ""
            End With
        Next

        txtCOLLECTION_CODE.Clear()

    End Sub

    Private Sub Load_Record()

        MyBase.EnforceConstraints(False)

        Absx1.txtFor("COLLECTION_CODE").Text = COLLECTION_CODE

        ALLO_CTL_NOs_to_Delete.Clear()

        ALLO_CTL_NO_to_copy = ""
        CUST_CODE_to_copy = ""

        ALLO_CTL_NO_new.Clear()
        ITEM_CODE_new.Clear()
        iCol.Clear()
        ReDim ALLO_CTL_NOi(maxAlloations)

        ALLO_CTL_NOs = ""
        For Each row As DataRow In dst.Tables("SOTALLOX").Select("SELECTED = '1'")
            ALLO_CTL_NOs &= ", '" & row.Item("ALLO_CTL_NO") & "'"
        Next

        If ALLO_CTL_NOs <> "" Then
            Add_Allocations(ALLO_CTL_NOs, True)
        End If

        For Each row As DataRow In dst.Tables("ICTITEM1").Select("SELECTED = '1'")
            Add_Item(row.Item("ITEM_CODE"))
        Next

        iColumn = 0
        For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select
            Add_Allocation_to_Grid(rowSOTALLO1)
        Next

        MyBase.EnforceConstraints(True)

        ASCMAIN1.Progress(String.Empty, String.Empty)
    End Sub

    Sub Print_Record()
        Create_Report()
    End Sub
    Function Create_Report() As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Dim REPORT_NAME As String = "SORALLO1"
        Dim RPT As String = REPORT_NAME

        If Not REPORTS.ContainsKey(REPORT_NAME) Then
            REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
            REPORTS(REPORT_NAME).Prepare_dst(False, "")
        End If


        'REPORTS(REPORT_NAME).Fill_Records_RPT("")

        dst.Tables("SOTALLOZ").Rows.Clear()
        dst.Tables("ICTSTAT2").Rows.Clear()
        dst.Tables("ARTCUST1").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ARTCUST1").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO2").Rows.Clear()
        REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO1").Rows.Clear()

        Dim ITEM_CODEs As New List(Of String)
        Dim CUST_CODEs As New List(Of String)

        For Each row As DataRow In dst.Tables("SOTALLO1").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO1")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() _
                    {"ALLO_CTL_NO", "ITEM_CODE", "DATE_START", "DATE_END", "INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE", "ALLOW_OVER", _
                     "ITEM_DESC", "COLLECTION_CODE", "BRAND_CODE", "ITEM_BASIC_PROMO", "ITEM_SNU_CODE"}
                    If COLUMN_NAME = "BRAND_CODE" Then
                    Else
                        rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                    End If
                Next
                .Rows.Add(rowR)
            End With
            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")
            Fill_Records("SOTALLOZ", ALLO_CTL_NO, False)
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            If Not ITEM_CODEs.Contains(ITEM_CODE) Then
                Fill_Records("ICTSTAT2", ITEM_CODE, False)
            End If
        Next

        For Each row As DataRow In dst.Tables("SOTALLO2").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLO2")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() {"ALLO_CTL_NO", "CUST_CODE", "QTY_ALLO"}
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            If Not CUST_CODEs.Contains(CUST_CODE) Then
                Fill_Records("ARTCUST1", CUST_CODE, False)
            End If
        Next

        For Each row As DataRow In dst.Tables("ARTCUST1").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ARTCUST1")
                Dim rowR As DataRow = .NewRow
                For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME"}
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
        Next

        For Each row As DataRow In dst.Tables("SOTALLOZ").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("SOTALLOZ")
                Dim rowR As DataRow = .NewRow
                For Each DC As DataColumn In dst.Tables("SOTALLOZ").Columns
                    Dim COLUMN_NAME As String = DC.ColumnName
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
        Next

        For Each row As DataRow In dst.Tables("ICTSTAT2").Select("")
            With REPORTS(REPORT_NAME).clsASCBASE1.dst.Tables("ICTSTAT2")
                Dim rowR As DataRow = .NewRow
                For Each DC As DataColumn In dst.Tables("ICTSTAT2").Columns
                    Dim COLUMN_NAME As String = DC.ColumnName
                    rowR.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                Next
                .Rows.Add(rowR)
            End With
        Next


        With REPORTS(REPORT_NAME).clsASCBASE1
            .Print_Report_Begin()
            '.CR_params.Add("SUBT", "")
            '.CR_params.Add("CONS_INV", "")
            'Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", , False)

            Dim SUBT As String = "Allocations by Item/Customer (Screen Report)"
            .CR_params.Add("SUBT", SUBT) ' "")
            .CR_params.Add("PAGE_EJECT", "0")
            .CR_params.Add("EXC_ONLY", "0")
            .Generate_Report(RPT, Me.Text, SUBT)
            .Print_Report_End()
            '.Print_Report_End(,  True)




            'CR_params.Add("SUBT", "Customer/Item")
            'CR_params.Add("PAGE_EJECT", IIf(Absx1.chkFor("CHKPAGE_EJECT").Checked, "1", "0"))
            'CR_params.Add("EXC_ONLY", IIf(Absx1.chkFor("CHKEXC_ONLY").Checked, "1", "0"))
            'RPT = "SORALLO2"
            'Generate_Report(RPT, , SUBT)

        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return ""
    End Function

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()

            Update_Record_TDA("SOTALLO1", "ALLO_CTL_NO in (" & Mid(ALLO_CTL_NOs, 2) & ")")

            dst.Tables("SOTALLO2").Rows.Clear()
            For Each row As DataRow In dst.Tables("SOTALLOC").Select
                For ictr As Integer = 1 To iColumn
                    If ALLO_CTL_NOi(ictr) <> "" Then
                        Dim QTY_ALLO As Int64 = Val(row.Item("ALLO_" & Format(ictr, "00")) & "")
                        If QTY_ALLO <> 0 Then
                            Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").NewRow
                            rowSOTALLO2.Item("ALLO_CTL_NO") = ALLO_CTL_NOi(ictr)
                            rowSOTALLO2.Item("CUST_CODE") = row.Item("CUST_CODE")
                            rowSOTALLO2.Item("QTY_ALLO") = QTY_ALLO
                            dst.Tables("SOTALLO2").Rows.Add(rowSOTALLO2)
                        End If
                    End If
                Next
            Next

            Update_Record_TDA("SOTALLO2", "ALLO_CTL_NO in (" & Mid(ALLO_CTL_NOs, 2) & ")")

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is" & vbCrLf _
                & "  Select Distinct ORDR_GROUP_NO from SOTORDR1 where ORDR_NO IN (" & vbCrLf _
                & "  Select Distinct ORDR_NO from SOTORDR2 where ITEM_CODE IN (" & vbCrLf _
                & "  Select ITEM_CODE from SOTALLO1 where ALLO_CTL_NO IN (" & Mid(ALLO_CTL_NOs, 2) & "))" & vbCrLf _
                & "   and ORDR_STATUS IN ('O','P') and ALLO_CTL_NO IS NULL);" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into SOTALLH1 Select '" & XNO & "' XNO, SOTALLO1.* from SOTALLO1 where ALLO_CTL_NO in (" & Mid(ALLO_CTL_NOs, 2) & ")"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into SOTALLH2 Select '" & XNO & "' XNO, SOTALLO2.* from SOTALLO2 where ALLO_CTL_NO in (" & Mid(ALLO_CTL_NOs, 2) & ")"
            ASCDATA1.ExecuteSQL()

            If ALLO_CTL_NOs_to_Delete.Count <> 0 Then
                Dim sqlw As String = " where ALLO_CTL_NO in ('" & Join(ALLO_CTL_NOs_to_Delete.ToArray, "','") & "')"
                ASCMAIN1.sql = "Delete from SOTALLO1" & sqlw
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Delete from SOTALLO2" & sqlw
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Update SOTORDR2 Set ALLO_CTL_NO = Null" & sqlw
                ASCDATA1.ExecuteSQL()
            End If

            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try

    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTALLO1, "BBBBBBB", _
                        "Item Status Inquiry", _
                        "Copy Allocation Dates & Qtys", _
                        "Paste Allocation Dates to Selected Rows", _
                        "Paste Allocation Qtys to Selected Rows", _
                        "Remove from List (No Update)", _
                        "Clear Allocation Qtys", _
                        "Delete Allocation")
        Load_Popup_Menu(grdSOTALLOX, "BBBB", "Select All", "De-Select All", "Select All in Group", "Item Status Inquiry")
        Load_Popup_Menu(grdICTITEM1, "BBB", "Select All", "De-Select All", "Item Status Inquiry")

        Load_Popup_Menu(grdSOTALLOC, "SSBB", "Show Filter", "Show Pins", "Add Customers", "Copy Customer Qtys", "Paste Customer Qtys to Selected Rows")

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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name
            Case "grdICTITEM1"
                tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode

            Case "grdSOTALLOX"
                tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode
                tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode
                tlb_btn = DirectCast(tlb_pop.Tools("Select All in Group"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not ScreenMode

            Case "grdSOTALLO1"

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Allocation Qtys to Selected Rows"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ALLO_CTL_NO_to_copy <> "") And grdSOTALLO1.Selected.Rows.Count > 0
                tlb_btn.SharedProps.Caption = "Paste Allocation " & ALLO_CTL_NO_to_copy & " Qtys to Selected Rows"

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Allocation Dates to Selected Rows"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ALLO_CTL_NO_to_copy <> "") And grdSOTALLO1.Selected.Rows.Count > 0
                tlb_btn.SharedProps.Caption = "Paste Allocation " & ALLO_CTL_NO_to_copy & " Dates to Selected Rows"

            Case "grdSOTALLOC"
                'tlb_btn = DirectCast(tlb_pop.Tools("Clear Allocation Qtys"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = ScreenMode
                'tlb_btn = DirectCast(tlb_pop.Tools("Delete Allocation"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = ScreenMode

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Customer Qtys to Selected Rows"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (CUST_CODE_to_copy <> "") And grdSOTALLOC.Selected.Rows.Count > 0
                tlb_btn.SharedProps.Caption = "Paste Customer " & CUST_CODE_to_copy & " Qtys to Selected Rows"


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '   e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                'Case "grdSOTORDR0"
                '    tlb_pop = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
                '    tlb_btn = DirectCast(tlb_pop.Tools("Store Configuration Report"), UltraWinToolbars.ButtonTool)
                '    Dim ORDR_TYPE As String = ""
                '    If grd.ActiveRow IsNot Nothing And grd.ActiveRow.IsDataRow Then
                '        ORDR_TYPE = grd.ActiveRow.Cells("ORDR_TYPE").Value & ""
                '    End If
                '    tlb_btn.SharedProps.Visible = (ORDR_TYPE = "O")
                '    tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Summary"), UltraWinToolbars.ButtonTool)
                '    tlb_btn.SharedProps.Visible = ScreenMode


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Add Customers"
                Add_Codes(grdSOTALLOC, "ARTCUST1", "CUST_CODE", "Customers")

            Case "Paste Allocation Dates to Selected Rows"
                If ALLO_CTL_NO_to_copy = "" Or grdSOTALLO1.Selected.Rows.Count = 0 Then
                Else
                    Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO_to_copy)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLO1.Selected.Rows
                        grow.Cells("DATE_START").Value = rowSOTALLO1.Item("DATE_START")
                        grow.Cells("DATE_END").Value = rowSOTALLO1.Item("DATE_END")
                        grow.Update()
                    Next
                End If

            Case "Paste Allocation Qtys to Selected Rows"
                If ALLO_CTL_NO_to_copy = "" Or grdSOTALLO1.Selected.Rows.Count = 0 Then
                Else
                    Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO_to_copy)
                    Dim ALLO_CTL_NOs_to_copy_to As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLO1.Selected.Rows
                        If grow.Cells("ALLO_CTL_NO").Value & "" <> ALLO_CTL_NO_to_copy Then
                            ALLO_CTL_NOs_to_copy_to.Add(grow.Cells("ALLO_CTL_NO").Value)
                        End If
                    Next

                    If ALLO_CTL_NOs_to_copy_to.Count <> 0 Then
                        Dim i_to_copy As Integer = iCol(ALLO_CTL_NO_to_copy)
                        For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                            For Each ALLO_CTL_NO As String In ALLO_CTL_NOs_to_copy_to
                                Dim i As Integer = iCol(ALLO_CTL_NO)
                                rowSOTALLOC.Item("ALLO_" & Format(i, "00")) = rowSOTALLOC.Item("ALLO_" & Format(i_to_copy, "00"))
                            Next
                        Next
                    End If
                End If

            Case "Paste Customer Qtys to Selected Rows"
                If CUST_CODE_to_copy = "" Or grdSOTALLOC.Selected.Rows.Count = 0 Then
                Else
                    Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE_to_copy)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTALLOC.Selected.Rows
                        For I As Integer = 1 To iColumn
                            grow.Cells("ALLO_" & Format(I, "00")).Value = rowSOTALLOC.Item("ALLO_" & Format(I, "00"))
                        Next
                        grow.Update()
                    Next
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Select All in Group"
                Dim ALLO_GROUP_CODE As String = grd.ActiveRow.Cells("ALLO_GROUP_CODE").Value & ""
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.Cells("ALLO_GROUP_CODE").Value & "" = ALLO_GROUP_CODE Then
                        grow.Cells("SELECTED").Value = "1"
                        grow.Update()
                    End If
                Next

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Remove from List (No Update)"
                Dim ALLO_CTL_NO = grd.ActiveRow.Cells("ALLO_CTL_NO").Value
                Dim I As Integer = iCol(ALLO_CTL_NO)
                grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(I, "00")).Hidden = True
                iCol.Remove(ALLO_CTL_NO)
                ALLO_CTL_NOi(I) = ""
                ALLO_CTL_NOs = Replace(ALLO_CTL_NOs, ", '" & ALLO_CTL_NO & "'", "")
                Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                rowSOTALLO1.Delete()
           
            Case "Clear Allocation Qtys"
                Dim ALLO_CTL_NO As String = grd.ActiveRow.Cells("ALLO_CTL_NO").Value
                Dim i As Integer = iCol(ALLO_CTL_NO)
                For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                    rowSOTALLOC.Item("ALLO_" & Format(i, "00")) = DBNull.Value
                Next

            Case "Delete Allocation"
                Dim ALLO_CTL_NO As String = grd.ActiveRow.Cells("ALLO_CTL_NO").Value

                ASCMAIN1.sql = "Select Count (*) from SOTORDR2 where ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                Dim USES As Integer = Val(ASCDATA1.GetDataValue)

                If USES <> 0 Then
                    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1.Item("ITEM_NOT_ALLOCATED") & "" <> "1" Then
                        MsgBox("You Cannot Delete an Allocation if it has been used unless the Item is flagged as Not-Allocated", MsgBoxStyle.OkOnly, "Cannot Peform Update Requested")
                        Exit Sub
                    End If
                End If

                Dim i As Integer = iCol(ALLO_CTL_NO)
                For Each rowSOTALLOC As DataRow In dst.Tables("SOTALLOC").Select("")
                    rowSOTALLOC.Item("ALLO_" & Format(i, "00")) = DBNull.Value
                Next
                ALLO_CTL_NOs_to_Delete.Add(ALLO_CTL_NO)
                grd.ActiveRow.Delete(False)

            Case "Copy Allocation Dates & Qtys"
                ALLO_CTL_NO_to_copy = grd.ActiveRow.Cells("ALLO_CTL_NO").Value
                grdSOTALLO1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

            Case "Copy Customer Qtys"
                CUST_CODE_to_copy = grd.ActiveRow.Cells("CUST_CODE").Value
                grdSOTALLOC.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Find_ITEM_CODE()
                    'Click_Command("Load", e)
                End If

            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 Is Nothing Then
                            MsgBox("Invalid Value Specified for Customer Code (" & CUST_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                        Else
                            Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
                            If rowSOTALLOC IsNot Nothing Then
                                MsgBox("Customer " & CUST_CODE & " is already in Allocation List", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                            Else
                                rowSOTALLOC = dst.Tables("SOTALLOC").NewRow
                                rowSOTALLOC.Item("CUST_CODE") = CUST_CODE
                                rowSOTALLOC.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                                rowSOTALLOC.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
                                rowSOTALLOC.Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE")
                                dst.Tables("SOTALLOC").Rows.Add(rowSOTALLOC)
                            End If
                            Absx1.txtFor("CUST_CODE").Text = ""
                        End If
                        Application.DoEvents()
                        Absx1.txtFor("CUST_CODE").Focus()
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ITEM_CODE"
                Find_ITEM_CODE()
                'Click_Command("Load")
            Case "CUST_CODE"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                If CUST_CODE <> "" Then
                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                    If rowARTCUST1 Is Nothing Then
                        MsgBox("Invalid Value Specified for Customer Code (" & CUST_CODE & ")", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                    Else
                        Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
                        If rowSOTALLOC IsNot Nothing Then
                            MsgBox("Customer " & CUST_CODE & " is already in Allocation List", MsgBoxStyle.OkOnly, "Cannot Add Customer")
                        Else
                            rowSOTALLOC = dst.Tables("SOTALLOC").NewRow
                            rowSOTALLOC.Item("CUST_CODE") = CUST_CODE
                            rowSOTALLOC.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                            rowSOTALLOC.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
                            rowSOTALLOC.Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE")
                            dst.Tables("SOTALLOC").Rows.Add(rowSOTALLOC)
                        End If
                        Absx1.txtFor("CUST_CODE").Text = ""
                    End If
                    Application.DoEvents()
                    Absx1.txtFor("CUST_CODE").Focus()
                End If
        End Select
    End Sub
#End Region

#Region "grdSOTALLOC"
    Private Sub grdSOTALLOC_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTALLOC.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                Dim CUST_CODE As String = Validate_Customer(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)
                If CUST_CODE <> "" Then
                    e.Cell.Row.Cells("CUST_NAME").Value = rowARTCUST1.Item("CUST_NAME") & ""
                    e.Cell.Row.Cells("SREP_CODE").Value = rowARTCUST1.Item("SREP_CODE") & ""
                    e.Cell.Row.Cells("TRADE_CLASS_CODE").Value = rowARTCUST1.Item("TRADE_CLASS_CODE")

                End If

                'Case "ORDR_QTY"
                '    grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value

                'Case "ORDR_QTY_OPEN"
                '    grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") _
                '        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") _
                '        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value & "") _
                '        - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "")
                '    If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Text) < 0 Then
                '        grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = 0
                '    End If

        End Select
    End Sub

    Private Sub grdSOTALLOC_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTALLOC.BeforeCellUpdate

        Select Case e.Cell.Column.Key

        End Select
    End Sub

    Private Sub grdSOTALLOC_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTALLOC.BeforeExitEditMode
        If grdSOTALLOC.ActiveCell IsNot Nothing Then
            With grdSOTALLOC.ActiveCell
                Select Case .Column.Key
                    Case "CUST_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTALLOC_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTALLOC.BeforeRowUpdate

        ' Validate_Columns("CUST_CODE", e.Cancel)
        'If Not e.Cancel Then
        '    Validate_Columns("ORDR_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        'End If

        If e.Cancel = True Then
            Exit Sub
        End If

        ' ITEM_CODE_last_entry = e.Row.Cells("ITEM_CODE").Value & ""

        'If e.Row.IsAddRow Then
        '    e.Row.Cells("ORDR_NO").Value = ORDR_NO

        'End If
    End Sub

    Private Sub grdSOTALLOC_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTALLOC.ClickCellButton

        Dim COLUMN_NAME As String = e.Cell.Column.Key
        'Dim SQL As String = ""

        'If e.Cell.Column.CellActivation = UltraWinGrid.Activation.NoEdit Then
        '    Exit Sub
        'End If


        Select Case COLUMN_NAME
            Case "CUST_CODE"
                '    SQL = "SELECT CUST_CODE, CUST_NAME, SREP_CODE FROM ARTCUST1"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTALLOC, sql_where)
                'Case Else
                '    Exit Sub
        End Select
    End Sub

#End Region

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTALLOC.ActiveRow
            Select Case COLUMN_NAME

                Case "CUST_CODE"
                    If .Cells("CUST_CODE").Text <> "" Then
                        Dim CUST_CODE As String = Validate_Customer(.Cells("CUST_CODE").Value & "")
                        Cancel = (CUST_CODE = "")
                    End If

                    'Case "ORDR_QTY"
                    '    If Trim(.Cells("ITEM_CODE").Value & "") = "" Then
                    '        Cancel = True
                    '        Exit Sub
                    '    End If
                    '    If Trim(.Cells("ORDR_QTY").Value & "") = "" Then
                    '        MsgBox("Order Qty Not Specified", vbOKOnly, "Cannot Update Record")
                    '        Cancel = True
                    '        grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY")
                    '        Exit Sub
                    '    End If
                    '    If Val(.Cells("ORDR_QTY").Value & "") < 0 Then
                    '        MsgBox("Order Qty May Not be Negative", vbOKOnly, "Invalid Order Quantity")
                    '        Cancel = True
                    '    End If
            End Select
        End With
    End Sub

    Function Validate_Customer(CUST_CODE_z As String) As String
        Dim E As String = ""

        Dim CUST_CODE As String = ""
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE_z)

        If rowARTCUST1 Is Nothing Then
            E = "Customer is Not on File" & vbCrLf
        Else
            If rowARTCUST1.Item("CUST_STATUS") & "" <> "A" Then
                E = "Customer Status is not Active" & vbCrLf
            End If
            If rowARTCUST1.Item("SREP_CODE") & "" = "" Then
                E = "Customer does not have a valid Sales Rep" & vbCrLf
            End If
            If rowARTCUST1.Item("TRADE_CLASS_CODE") & "" = "" Then
                E = "Customer does not have a valid Trade Class" & vbCrLf
            End If
            'If rowICTITEM1.Item("SALES_DIVISION_CODE") & "" = "" Then
            '    E = "Item does not have a valid Division Code" & vbCrLf
            'End If
        End If

        If E <> "" And grdSOTALLOC.ActiveRow IsNot Nothing AndAlso grdSOTALLOC.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Customer Code Entered is Invalid because ...")
        Else
            If E = "" Then
                CUST_CODE = rowARTCUST1.Item(0)
            End If
        End If
        Return CUST_CODE
    End Function

    Sub Load_SOTALLOX()

        dst.Tables("SOTALLOX").Rows.Clear()
        dst.Tables("ICTITEM1").Rows.Clear()

        If grdICTCOLL1.ActiveRow Is Nothing OrElse Not grdICTCOLL1.ActiveRow.IsDataRow Then
            grdSOTALLOX.Text = "You must select a Collection"
            grdICTITEM1.Text = "You must select a Collection"
        Else

            COLLECTION_CODE = grdICTCOLL1.ActiveRow.Cells("COLLECTION_CODE").Value
            grdSOTALLOX.Text = "Allocations for Items in Collection " & COLLECTION_CODE
            grdICTITEM1.Text = "Items in Collection " & COLLECTION_CODE

            ASCMAIN1.sql = "Select SOTALLO1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
                    & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                    & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_STATUS" & vbCrLf _
                    & " from ICTITEM1,SOTALLO1" & vbCrLf _
                    & " where ICTITEM1.COLLECTION_CODE = '" & COLLECTION_CODE & "'" & vbCrLf _
                    & "   and SOTALLO1.ITEM_CODE = ICTITEM1.ITEM_CODE"
            Fill_Records("SOTALLOX", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                    & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                    & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_STATUS" & vbCrLf _
                    & " from ICTITEM1 where COLLECTION_CODE = '" & COLLECTION_CODE & "'"
            Fill_Records("ICTITEM1", "", True, ASCMAIN1.sql)

            Set_SOTALLOX()

            ASCMAIN1.sql = "Select * from SOTALLO1 where ITEM_CODE in (Select ITEM_CODE from (" & ASCMAIN1.sql & "))"

            For Each row As DataRow In dst.Tables("SOTALLOX").Select("", "ITEM_CODE,ALLO_CTL_NO DESC")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                If rowICTITEM1.Item("ALLO_CTL_NO") & "" = "" Then
                    rowICTITEM1.Item("ALLO_CTL_NO") = row.Item("ALLO_CTL_NO")
                    rowICTITEM1.Item("DATE_START") = row.Item("DATE_START")
                    rowICTITEM1.Item("DATE_END") = row.Item("DATE_END")
                End If
            Next
        End If
    End Sub

    Sub Set_SOTALLOX()
        Dim sqlw As String = ""
        If optSN.Value <> "*" Then
            sqlw &= " and ITEM_SNU_CODE = '" & optSN.Value & "'"
            grdSOTALLOX.Text &= ", " & optSN.Text
            grdICTITEM1.Text &= ", " & optSN.Text
        End If
        If optBP.Value <> "*" Then
            sqlw &= " and ITEM_BASIC_PROMO = '" & optBP.Value & "'"
            grdSOTALLOX.Text &= ", " & optBP.Text
            grdICTITEM1.Text &= ", " & optBP.Text
        End If
        If chkActiveOnly.Checked Then
            sqlw &= " and ITEM_STATUS = 'A'"
            grdSOTALLOX.Text &= ", " & chkActiveOnly.Text
            grdICTITEM1.Text &= ", " & chkActiveOnly.Text
        End If

        Dim dvw As DataView
        dvw = DirectCast(grdSOTALLOX.DataSource, DataTable).DefaultView
        dvw.RowFilter = Mid(sqlw, 5)
        dvw = DirectCast(grdICTITEM1.DataSource, DataTable).DefaultView
        dvw.RowFilter = Mid(sqlw, 5)

        Sort_grdColumns(grdSOTALLOX, "ITEM_CODE")
        Sort_grdColumns(grdICTITEM1, "ITEM_CODE")
    End Sub

    Private Sub optSN_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSN.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_SOTALLOX()
    End Sub

    Private Sub optBP_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optBP.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_SOTALLOX()
    End Sub

    Private Sub chkActiveOnly_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkActiveOnly.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Set_SOTALLOX()
    End Sub

    Private Sub grdICTCOLL1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTCOLL1.AfterRowActivate
        Load_SOTALLOX()
    End Sub

    Sub Find_ITEM_CODE()
        Dim ITEM_CODE As String = Absx1.txtFor("ITEM_CODE").Text
        If ITEM_CODE <> "" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1 Is Nothing Then
                MsgBox("Invalid Item Code Specified (" & ITEM_CODE & ")")
                Exit Sub
            Else
                Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE")
                If COLLECTION_CODE = "" Then
                    MsgBox("No Collection Code Specified for Item " & ITEM_CODE)
                    Exit Sub
                Else
                    For Each growBrand As UltraWinGrid.UltraGridRow In grdICTCOLL1.Rows
                        For Each growHC As UltraWinGrid.UltraGridRow In growBrand.ChildBands(0).Rows
                            For Each grow As UltraWinGrid.UltraGridRow In growHC.ChildBands(0).Rows
                                If grow.IsDataRow Then
                                    If grow.Cells("COLLECTION_CODE").Value & "" = COLLECTION_CODE Then
                                        grdICTCOLL1.ActiveRow = grow
                                        rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                                        ' rowICTITEM1.Item("SELECTED") = "1"

                                        For Each growITEM As UltraWinGrid.UltraGridRow In grdICTITEM1.Rows
                                            If growITEM.Cells("ITEM_CODE").Value = ITEM_CODE Then
                                                growITEM.Selected = True
                                                Exit For
                                            End If

                                        Next


                                        tabSOTALLOX.SelectedTab = tabSOTALLOX.Tabs("Items")
                                        If rowICTITEM1.Item("ITEM_BASIC_PROMO") & "" <> optBP.Value And optBP.Value <> "*" Then
                                            optBP.Value = "*"
                                        End If
                                        If rowICTITEM1.Item("ITEM_SNU_CODE") & "" <> optSN.Value And optSN.Value <> "*" Then
                                            optSN.Value = "*"
                                        End If
                                        If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
                                            chkActiveOnly.Checked = False
                                        End If
                                        Absx1.txtFor("ITEM_CODE").Text = ""
                                        Exit Sub
                                    End If
                                End If
                            Next
                        Next
                      
                    Next
                    MsgBox("Cannot Find Item " & ITEM_CODE)
                End If
            End If
        End If
    End Sub

#Region "grdSOTALLOX"

    Private Sub grdSOTALLOX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTALLOX.DoubleClickRow
        If ScreenMode Then
            Dim ALLO_CTL_NO As String = e.Row.Cells("ALLO_CTL_NO").Value
            Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
            If rowSOTALLO1 IsNot Nothing Then
                If MsgBox("Allocation " & ALLO_CTL_NO & " is already in the Allocation List.", MsgBoxStyle.OkOnly, "Cannot Add an Allocation Twice") Then
                    Exit Sub
                End If
            End If

            If Add_Allocations(",'" & ALLO_CTL_NO & "'", False, True) Then
                rowSOTALLO1 = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                Add_Allocation_to_Grid(rowSOTALLO1)
                MsgBox("Allocation " & ALLO_CTL_NO & " for " & rowSOTALLO1.Item("ITEM_CODE") & " has been Added", _
                       MsgBoxStyle.OkOnly, "Verification")
            End If
        End If
    End Sub

    Private Sub grdSOTALLOX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOX.InitializeRow
        With e.Row.Cells("SELECTED")
            If .Value & "" = "1" Then
                .Appearance.BackColor = Drawing.Color.Red
            Else
                .Appearance.BackColor = Drawing.Color.Empty
            End If
        End With
    End Sub

#End Region

#Region "grdICTITEM1"

    Private Sub grdICTITEM1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTITEM1.DoubleClickRow
        If ScreenMode Then
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
            If ITEM_CODE_new.Contains(ITEM_CODE) Then
                If MsgBox("Item " & ITEM_CODE & " is already in the Allocation List with a new allocation." _
                          & vbCrLf & vbCrLf & "Do you really want to add it again?", _
                          MsgBoxStyle.YesNo, "Item has already been selected to create a New Allocation") = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If
            Dim rowSOTALLO1 As DataRow = Add_Item(ITEM_CODE, True)
            If rowSOTALLO1 IsNot Nothing Then
                Add_Allocation_to_Grid(rowSOTALLO1)
                MsgBox("A New Allocation record has been added for Item " & ITEM_CODE, MsgBoxStyle.OkOnly, "Verification")
            End If
        End If
    End Sub

    Private Sub grdICTITEM1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEM1.InitializeRow
        With e.Row.Cells("SELECTED")
            If .Value & "" = "1" Then
                .Appearance.BackColor = Drawing.Color.Red
            Else
                .Appearance.BackColor = Drawing.Color.Empty
            End If
        End With
    End Sub
#End Region

    Function Add_Allocations(sqlALLO_CTL_NOs As String, clear_table As Boolean, Optional multi_task As Boolean = False) As Boolean

        If multi_task Then
            Dim ALLO_CTL_NO As String = Replace(Mid(sqlALLO_CTL_NOs, 2), "'", "")
            If Not ASCMAIN1.Logical_Lock("SOTALLO1", ALLO_CTL_NO, False, True, False) Then
                Return False
            End If
            Dim row As DataRow = LookUp("SOTALLO1", ALLO_CTL_NO)
            If row IsNot Nothing Then
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE) Then
                    Exit Function
                End If
            End If
        End If

        ASCMAIN1.sql = "Select SOTALLO1.*" & sql_ICTITEM1 _
                      & " from SOTALLO1,ICTITEM1" & vbCrLf _
                      & " where ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
                      & "   and SOTALLO1.ALLO_CTL_NO IN (" & Mid(sqlALLO_CTL_NOs, 2) & ")"
        Fill_Records("SOTALLO1", "", clear_table, ASCMAIN1.sql)

        ASCMAIN1.sql = "Select SOTALLO2.*, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " from SOTALLO2, ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
            & " and SOTALLO2.ALLO_CTL_NO IN (" & Mid(sqlALLO_CTL_NOs, 2) & ")"
        Fill_Records("SOTALLO2", "", clear_table, ASCMAIN1.sql)
        Return True
    End Function

    Function Add_Allocation_to_Grid(rowSOTALLO1 As DataRow) As Int64

        iColumn += 1

        Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")

        With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(iColumn, "00"))
            .Hidden = False
            .Width = 100
            .Tag = ALLO_CTL_NO
        End With

        dst.Tables("SOTALLOC").Columns("ALLO_" & Format(iColumn, "00")).ExtendedProperties("ALLO_CTL_NO") = ALLO_CTL_NO

        For Each rowSOTALLO2 As DataRow In dst.Tables("SOTALLO2").Select("ALLO_CTL_NO = '" & ALLO_CTL_NO & "'")
            Dim CUST_CODE As String = rowSOTALLO2.Item("CUST_CODE")
            Dim rowSOTALLOC As DataRow = dst.Tables("SOTALLOC").Rows.Find(CUST_CODE)
            If rowSOTALLOC Is Nothing Then
                rowSOTALLOC = dst.Tables("SOTALLOC").NewRow
                rowSOTALLOC.Item("CUST_CODE") = rowSOTALLO2.Item("CUST_CODE")
                rowSOTALLOC.Item("CUST_NAME") = rowSOTALLO2.Item("CUST_NAME")
                rowSOTALLOC.Item("SREP_CODE") = rowSOTALLO2.Item("SREP_CODE")
                rowSOTALLOC.Item("TRADE_CLASS_CODE") = rowSOTALLO2.Item("TRADE_CLASS_CODE")
                dst.Tables("SOTALLOC").Rows.Add(rowSOTALLOC)
            End If
            rowSOTALLOC.Item("ALLO_" & Format(iColumn, "00")) = Val(rowSOTALLO2.Item("QTY_ALLO") & "")
        Next

        iCol.Add(ALLO_CTL_NO, iColumn)
        ALLO_CTL_NOi(iColumn) = ALLO_CTL_NO
        Set_Header(ALLO_CTL_NO)
        Return iColumn
    End Function

    Sub Set_Header(ALLO_CTL_NO As String)
        Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
        Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")

        Dim DATE_START As String = ""
        If rowSOTALLO1.Item("DATE_START") & "" <> "" Then DATE_START = Format(rowSOTALLO1.Item("DATE_START"), "MM/dd/yy")

        Dim DATE_END As String = ""
        If rowSOTALLO1.Item("DATE_END") & "" <> "" Then DATE_END = Format(rowSOTALLO1.Item("DATE_END"), "MM/dd/yy")

        Dim header As String = "" _
            & ITEM_CODE & vbCrLf _
            & DATE_START & vbCrLf _
            & DATE_END & vbCrLf

        Dim I As Integer = iCol(ALLO_CTL_NO)
        With grdSOTALLOC.DisplayLayout.Bands(0).Columns("ALLO_" & Format(I, "00"))
            .Header.Caption = header
        End With

    End Sub

    Function Add_Item(ITEM_CODE As String, Optional multi_task As Boolean = False) As DataRow

        If multi_task Then
            If Not ASCMAIN1.Logical_Lock("SOTALLO1", "ITEM:" & ITEM_CODE, False, True, False) Then
                Return Nothing
            End If
        End If

        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
        Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").NewRow

        Dim ALLO_CTL_NO As String = ASCMAIN1.Next_Control_No("SOTALLO1.ALLO_CTL_NO")
        ALLO_CTL_NO_new.Add(ALLO_CTL_NO)
        If Not ITEM_CODE_new.Contains(ITEM_CODE) Then ITEM_CODE_new.Add(ITEM_CODE)

        rowSOTALLO1.Item("ALLO_CTL_NO") = ALLO_CTL_NO
        rowSOTALLO1.Item("ITEM_CODE") = ITEM_CODE
        rowSOTALLO1.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
        rowSOTALLO1.Item("ITEM_BASIC_PROMO") = rowICTITEM1.Item("ITEM_BASIC_PROMO")
        rowSOTALLO1.Item("ITEM_SNU_CODE") = rowICTITEM1.Item("ITEM_SNU_CODE")
        rowSOTALLO1.Item("ITEM_CATGY_CODE") = rowICTITEM1.Item("ITEM_CATGY_CODE")
        rowSOTALLO1.Item("ITEM_STATUS") = rowICTITEM1.Item("ITEM_STATUS")
        rowSOTALLO1.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
        rowSOTALLO1.Item("ITEM_NOT_ALLOCATED") = rowICTITEM1.Item("ITEM_NOT_ALLOCATED")

        'rowSOTALLO1.Item("DATE_START") = String.Empty
        'rowSOTALLO1.Item("DATE_END") = String.Empty

        rowSOTALLO1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSOTALLO1.Item("INIT_DATE") = DATETIME_STAMP
        rowSOTALLO1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowSOTALLO1.Item("LAST_DATE") = DATETIME_STAMP

        rowSOTALLO1.Item("ALLOW_OVER") = "0"
        dst.Tables("SOTALLO1").Rows.Add(rowSOTALLO1)

        ALLO_CTL_NOs &= ",'" & ALLO_CTL_NO & "'"
        Return rowSOTALLO1
    End Function

    Private Sub tabSOTALLOC_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSOTALLOC.SelectedTabChanged
        If tabSOTALLOC.SelectedTab.Key = "Sales Details" Then
            Get_Sales()
        End If
    End Sub

    Sub Get_Sales()
        Dim COL As String = ""
        If grdSOTALLOC.ActiveCell IsNot Nothing Then
            COL = grdSOTALLOC.ActiveCell.Column.Key
        End If
        If Not COL.StartsWith("ALLO_") Then
            grdSOTORDR2.Visible = False
        Else
            Dim ALLO_CTL_NO As String = grdSOTALLOC.ActiveCell.Column.Tag
            Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
            If rowSOTALLO1.Item("DATE_START") & "" = "" _
            Or rowSOTALLO1.Item("DATE_END") & "" = "" Then
                dst.Tables("SOTORDR2").Rows.Clear()
            Else
                Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")
                Dim DATE_START As Date = rowSOTALLO1.Item("DATE_START")
                Dim DATE_END As Date = rowSOTALLO1.Item("DATE_END")
                Fill_Records("SOTORDR2", New Object() {ITEM_CODE, DATE_START, DATE_END})
            End If

            grdSOTORDR2.Visible = True
        End If
    End Sub

#Region "grdSOTALLO1"

    Private Sub grdSOTALLO1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTALLO1.AfterRowActivate

    End Sub

    Private Sub grdSOTALLO1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTALLO1.AfterRowUpdate

        Dim ALLO_CTL_NO As String = e.Row.Cells("ALLO_CTL_NO").Value & ""
        Set_Header(ALLO_CTL_NO)
    End Sub

    Private Sub grdSOTALLO1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLO1.InitializeRow
        Dim DATE_START As String = ""
        If e.Row.Cells("DATE_START").Value & "" <> "" Then DATE_START = Format(e.Row.Cells("DATE_START").Value, "yyyyMMdd")
        Dim DATE_END As String = ""
        If e.Row.Cells("DATE_END").Value & "" <> "" Then DATE_END = Format(e.Row.Cells("DATE_END").Value, "yyyyMMdd")

        If DATE_START = "" Then
            e.Row.Cells("DATE_START").Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("DATE_START").Appearance.BackColor = Drawing.Color.Empty
        End If
        If DATE_END = "" Then
            e.Row.Cells("DATE_END").Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("DATE_END").Appearance.BackColor = Drawing.Color.Empty
        End If
        If DATE_START <> "" And DATE_END <> "" Then
            If DATE_START > DATE_END Then
                e.Row.Cells("DATE_END").Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Cells("DATE_END").Appearance.ForeColor = Drawing.Color.Empty
            End If
        End If

        If e.Row.Cells("ALLO_CTL_NO").Value & "" = ALLO_CTL_NO_to_copy Then
            e.Row.Cells("ALLO_CTL_NO").Appearance.ForeColor = Drawing.Color.Green
        Else
            e.Row.Cells("ALLO_CTL_NO").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub
#End Region

    Private Sub grdSOTALLOC_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOC.InitializeRow
        If e.Row.Cells("CUST_CODE").Value & "" = CUST_CODE_to_copy Then
            e.Row.Cells("CUST_CODE").Appearance.ForeColor = Drawing.Color.Green
        Else
            e.Row.Cells("CUST_CODE").Appearance.ForeColor = Drawing.Color.Empty
        End If
    End Sub
End Class