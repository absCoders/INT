Public Class GLFGPAN1

    Dim ACCT_END_BAL_TOTAL As Double = 0
    Dim STMT_CODE As String = "I001" ' this should be parameterized
    Dim STMT_LINE_NO_max As Integer = 18 ' GP LINE
    Dim GLTDETL2 As String
    Dim GLTFINRD As String
    Dim COLS As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' period list should include distinct list of OPS_YYYYPP from GLTFINR2 plus the current PI period from the control file?

        With dst
            ASCMAIN1.sql = "Select GLTFINR1.*" _
            & " from EMP.GLTFINR1" _
            & " where GLTFINR1.STMT_CODE = :PARM1"
            Create_TDA(.Tables.Add, "GLTFINR1", "**", 0, False, "V")

            ASCMAIN1.sql = "Select GLTFINR2.*" _
            & " from EMP.GLTFINR2" _
            & " where GLTFINR2.STMT_CODE = :PARM1"
            Create_TDA(.Tables.Add, "GLTFINR2", "**", 0, False, "V")
            .Tables("GLTFINR2").Columns.Add("AMT", GetType(System.Double))
            '.Tables("GLTFINR2").Columns.Add("AMT", GetType(System.Double), "QTY_PHY_TOTAL - QTY_PHY_FOUND")
            '.Tables("GLTFINR2").Columns.Add("SHORT_LAYERS")


            ASCMAIN1.sql = "Select GLTDETL1.ACCT_CODE" _
            & ", GLTJRNL1.JOURNAL_TYPE" _
            & ", SUM (GLTDETL1.DETL_POSTING_AMT) AMT" _
            & " from EMP.GLTDETL1, EMP.GLTJRNL1" _
            & " where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO" _
            & " and GLTDETL1.OPS_YYYYPP = :PARM1" _
            & " group by GLTDETL1.ACCT_CODE, GLTJRNL1.JOURNAL_TYPE"
            Create_TDA(.Tables.Add, "GLTSTMTY", "**", 0, False, "V")

            ' EMP HARDCODE TO AVOID DIV T

            ASCMAIN1.sql = "Select GLTFINR3.*" _
            & " from EMP.GLTFINR3" _
            & " where GLTFINR3.STMT_CODE = :PARM1"
            Create_TDA(.Tables.Add, "GLTFINR3", "**", 0, False, "V")

            ASCMAIN1.sql = "Select GLTFINR4.*" _
            & " from EMP.GLTFINR4" _
            & " where GLTFINR4.STMT_CODE = :PARM1"
            Create_TDA(.Tables.Add, "GLTFINR4", "**", 0, False, "V")

            ASCMAIN1.sql = "Select GLTFINR3.STMT_CODE, GLTFINR3.STMT_LINE_NO" _
            & ", GLTACCT3.ACCT_CODE, GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE" _
            & ", GLTACCT1.ACCT_DESC" _
            & ", GLTACCT3.ACCT_BEG_BAL" _
            & ", GLTACCT3.ACCT_ACT_P01, GLTACCT3.ACCT_ACT_P02, GLTACCT3.ACCT_ACT_P03" _
            & ", GLTACCT3.ACCT_ACT_P04, GLTACCT3.ACCT_ACT_P05, GLTACCT3.ACCT_ACT_P06" _
            & ", GLTACCT3.ACCT_ACT_P07, GLTACCT3.ACCT_ACT_P08, GLTACCT3.ACCT_ACT_P09" _
            & ", GLTACCT3.ACCT_ACT_P10, GLTACCT3.ACCT_ACT_P11, GLTACCT3.ACCT_ACT_P12" _
            & " from EMP.GLTACCT3, EMP.GLTFINR3, EMP.GLTACCT1" _
            & " where GLTACCT3.ACCT_YEAR = :PARM1" _
            & " and GLTFINR3.STMT_CODE = :PARM2" _
            & " and GLTACCT3.ACCT_CODE = GLTFINR3.ACCT_CODE" _
            & " and GLTACCT3.SEG2_CODE = GLTFINR3.SEG2_CODE" _
            & " and GLTACCT3.SEG3_CODE = GLTFINR3.SEG3_CODE" _
            & " and GLTACCT3.SEG4_CODE = GLTFINR3.SEG4_CODE" _
            & " and GLTACCT1.ACCT_CODE = GLTFINR3.ACCT_CODE" _
            & " and GLTACCT3.SEG2_CODE <> 'T'" _
            & " and GLTFINR3.STMT_LINE_NO <= " & CStr(STMT_LINE_NO_max)
            Create_TDA(.Tables.Add, "GLTSTMTX", "**", 0, False, "VV", 7)
            .Tables("GLTSTMTX").Columns.Add("AMT", GetType(System.Double), "ACCT_BEG_BAL")

            .Relations.Add("GLTSTMTX", _
            New DataColumn() {.Tables("GLTFINR2").Columns("STMT_CODE"), .Tables("GLTFINR2").Columns("STMT_LINE_NO")}, _
            New DataColumn() {.Tables("GLTSTMTX").Columns("STMT_CODE"), .Tables("GLTSTMTX").Columns("STMT_LINE_NO")})

            .Tables("GLTFINR2").Columns("AMT").Expression = "SUM(CHILD.AMT)"

            ' EMP HARDCODE TO AVOID DIV T

            ASCMAIN1.sql = "Select SOTINVH0.CUST_CODE, SOTINVH0.PROD_CODE" _
            & ", Sum (SOTINVH0.QTY_UNITS) QTY_UNITS" _
            & ", Sum (SOTINVH0.QTY_UNITS * SOTINVH0.ORDR_PRICE_GRS) GRS" _
            & ", Sum (SOTINVH0.QTY_UNITS * SOTINVH0.ORDR_PRICE_NET) NET" _
            & " from EMP.SOTINVH0" _
            & " where SOTINVH0.OPS_YYYYPP = :PARM1" _
            & " and SOTINVH0.ORDR_DIV_CODE <> 'T'" _
            & " group by SOTINVH0.CUST_CODE, SOTINVH0.PROD_CODE"
            Create_TDA(.Tables.Add, "GLTGPAN1", "**", 0, False, "V", 2)
            '.Tables("GLTGPAN1").Columns.Add("EXT_COST", GetType(System.Double), "COUNT * UNIT_COST")


            GLTDETL2 = ASCMAIN1.Temp_Table("Select * from EMP.GLTDETL2 where ROWNUM <1")
            ASCMAIN1.sql = "Select * from " & GLTDETL2
            Create_TDA(.Tables.Add, "GLTDETL2", "**", 0, False)

            .Tables.Add("GLTGPANC")
            .Tables("GLTGPANC").Columns.Add("COLUMN_NAME")
            .Tables("GLTGPANC").Columns.Add("COLUMN_CAPTION")
            .Tables("GLTGPANC").Columns.Add("SELECTED")

            .Tables.Add("GLTGPANR")
        End With

        dst.Tables("GLTGPANC").Rows.Add(New Object() {"JOURNAL_TYPE", "Journal", "0"})
        dst.Tables("GLTGPANC").Rows.Add(New Object() {"SEG2_CODE", "Division", "0"})
        dst.Tables("GLTGPANC").Rows.Add(New Object() {"CUST_CODE", "Sold-To", "0"})
        dst.Tables("GLTGPANC").Rows.Add(New Object() {"VEND_CODE", "Supplier", "0"})
        dst.Tables("GLTGPANC").Rows.Add(New Object() {"PROD_CODE", "Product", "0"})
        dst.Tables("GLTGPANC").Rows.Add(New Object() {"SIZE_CODE", "Size", "0"})
        dst.Tables("GLTGPANC").Rows.Add(New Object() {"ORIG_CODE", "Origin", "0"})
        dst.Tables("GLTGPANC").Rows.Add(New Object() {"BRAND_CODE", "Brand", "0"})

        grdGLTFINR2.DataSource = dst.Tables("GLTFINR2")

        Call Get_PARM("GLTPARM1")

        grdGLTFINR2.DataSource = dst.Tables("GLTFINR2")

        'grdGLTFINR2.DataSource = dst
        'grdGLTFINR2.DataMember = "GLTFINR2"

        grdGLTGPAN1.DataSource = dst.Tables("GLTGPAN1")
        grdGLTGPANC.DataSource = dst.Tables("GLTGPANC")


        'Call Create_Summary(grdGLTFINR2, "PRICE_CATGY_CODE", "Count")
        'Call Create_Summary(grdGLTFINR2, "QTY_PHY_TOTAL")

        'Call Create_Summary(grdGLTFINR2, "RECEIPT_NO", "Count", "GLTSTMTX")
        'Call Create_Summary(grdGLTFINR2, "QTY_REC", "Sum", "GLTSTMTX")

        'Call Create_Summary(grdGLTGPAN1, "COUNT")
        'Call Create_Summary(grdGLTGPAN1, "EXT_COST")

        'With grdGLTFINR2.DisplayLayout
        '    .Bands("GLTFINR2").SortedColumns.Clear()
        '    .Bands("GLTFINR2").SortedColumns.Add("PRICE_CATGY_CODE", False)
        '    .Bands("GLTSTMTX").SortedColumns.Clear()
        '    .Bands("GLTSTMTX").SortedColumns.Add("PRICE_CATGY_CODE", False)
        '    .Bands("GLTSTMTX").SortedColumns.Add("RECEIPT_DATE", True)
        'End With

        Absx1.txtFor("STMT_CODE").Text = "I001"
        Absx1.optFor("STMT_TYPE").Value = "I"

        Call Set_SEGS(grdGLTFINR2, "GLTSTMTX")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Call Validate_Code("OPS_YYYYPP")
                If Absx1.txtFor("OPS_YYYYPP").Text = "" Then
                    EMsg &= vbCr & "You must specifiy the Period which ended just prior to taking the Count"
                End If

            Case "Rebuild Tiers"
                If HFs("OPS_YYYYPP") <> ROWs("ICTPARM1").Item("IC_PARM_PHY_OPS_YYYYPP") Then
                    EMsg &= vbCr & "Rebuild Tiers option only applicable to a P/I which is in Process"
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                'grdGLTGPANR.DataMember = "GLTGPANR"
                'grdGLTGPANR.DataSource = dst
                'grdGLTGPANJ.DataMember = "GLTGPANR.GLTDETL2"
                'grdGLTGPANJ.DataSource = dst


                Call Mode_Settings(False)

            Case "Excel"
                If UltraTabControl1.ActiveTab.Key = "Financial Statement" Then
                    Call Export_to_Excel(grdGLTFINR2)
                Else
                    Call Export_to_Excel(grdGLTGPAN1)
                End If

            Case "Print Report"
                'Call Print_Report_Begin()
                'CR_params.Add("GYPLEGEND", "???")
                'Generate_Report("GLRTBAL1")
                'Call Print_Report_End()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Excel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Print Report").Settings.Enabled = iScreenMode
                .Groups("Display Options").Visible = tf
            End With

        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf

        If ScreenMode Then
            Call Apply_Breakdown()
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("GLTFINR1").Rows.Clear()
        dst.Tables("GLTFINR2").Rows.Clear()
        dst.Tables("GLTFINR3").Rows.Clear()
        dst.Tables("GLTFINR4").Rows.Clear()
        dst.Tables("GLTSTMTX").Rows.Clear()
        dst.Tables("GLTDETL2").Rows.Clear()
        dst.Tables("GLTGPANR").Rows.Clear()
        dst.EnforceConstraints = True

        Absx1.txtFor("OPS_YYYYPP").Text = ""
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Financial Statement Information")
        Call Save_Header_Fields(UltraGroupBox1)

        Dim RYP As String = HFs("OPS_YYYYPP")

        Call Fill_Records("GLTFINR1", STMT_CODE)
        Call Fill_Records("GLTFINR2", STMT_CODE)
        Call Fill_Records("GLTFINR3", STMT_CODE)
        Call Fill_Records("GLTFINR4", STMT_CODE)

        Dim TTA As String = ""
        Dim TTB As String = ""
        Dim SQLP() As String

        Dim SQLF() As String

        Dim RY As String = Mid(RYP, 1, 4)
        Dim P As Integer = Val(Mid(RYP, 5, 2))
        Dim sqlA_select As String = ""
        Dim sqlA_group_by As String = ""
        Dim sqlA_where As String = ""

        ReDim SQLP(4)
        ReDim SQLF(4)

        Dim TTC As String = GLCMAIN1.Prepare_Work_File( _
        Me, TTA, TTB, GLTFINRD, RY, P, SQLP, SQLF, _
        sqlA_select, sqlA_group_by, sqlA_where, _
        BY_SEG2:=True, BY_SEG3:=True, BY_SEG4:=True)

        Dim selectcmd As String = Get_SelectCommand("GLTSTMTX")
        Set_SelectCommand("GLTSTMTX", Replace(selectcmd, "EMP.GLTFINR3", GLTFINRD & " GLTFINR3"))

        Call Load_GLTFINRx()

        Call Prepare_JE_Details()

        For Each row As DataRow In dst.Tables("GLTFINR2").Select("STMT_LINE_NO > " & CStr(STMT_LINE_NO_max))
            row.Delete()
        Next

        dst.Tables("GLTSTMTX").Columns("AMT").Expression = "ACCT_ACT_P" & Format(P, "00")

        If grdGLTFINR2.Rows.Count <> 0 Then grdGLTFINR2.ActiveRow = grdGLTFINR2.Rows(0)

        'Call Build_GLTGPANR()



        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub
#End Region

    Sub Rebuild_GLTFINRx()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Creating Cost-Tiers Work Tables")

        Dim sql As String = ""
        'sql = "Insert into GLTSTMTX Select * from " & GLTSTMTX
        'ASCDATA1.ExecuteSQL(sql)

        Call ASCMAIN1.Progress("")

        Call Load_GLTFINRx()
    End Sub

    Sub Load_GLTFINRx()
        Call ASCMAIN1.Progress("Now Loading Financial Statement Data")

        dst.EnforceConstraints = False

        Call Fill_Records("GLTSTMTX", New Object() {Mid(HFs("OPS_YYYYPP"), 1, 4), STMT_CODE})
        Call Fill_Records("GLTGPAN1", HFs("OPS_YYYYPP"))
        'dst.EnforceConstraints = True

        'For Each rowGLTFINR2 As DataRow In dst.Tables("GLTFINR2").Select("QTY_SHORT <> 0")
        '    If Val(rowGLTFINR2.Item("QTY_SHORT") & "") = Val(rowGLTFINR2.Item("QTY_PHY_TOTAL") & "") Then
        '        rowGLTFINR2.Item("SHORT_LAYERS") = "N"
        '    Else
        '        rowGLTFINR2.Item("SHORT_LAYERS") = "S"
        '    End If
        'Next

        Call ASCMAIN1.Progress("")

    End Sub

    Private Sub chkShort_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShort.CheckedChanged
        With grdGLTFINR2.DisplayLayout.Bands("GLTFINR2")
            'If chkShort.Checked Then
            '    .ColumnFilters("QTY_SHORT").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.GreaterThan, 0)
            'Else
            '    .ColumnFilters.ClearAllFilters()
            'End If

        End With
    End Sub

    Sub Prepare_JE_Details()
        ASCMAIN1.sql = "Truncate Table " & GLTDETL2
        ASCDATA1.ExecuteSQL()

        Call Prepare_JE_OPSJ_Revenue("S", HFs("OPS_YYYYPP"))
        Call Prepare_JE_OPSJ_Revenue("C", HFs("OPS_YYYYPP"))
        ' NEED W/D CHG, AR, IC
        ' NEED REBILLABLES
        ' NEED CGS
        Call Fill_Records("GLTDETL2")

    End Sub

    Sub Prepare_JE_OPSJ_Revenue(ByVal S_or_C As String, ByVal YP As String)

        ' SOTINVHD = DETAILS: SOTINVH3,SOTINVH7
        ' SOTINVHP = PRICES:  SOTINVH2,SOTINVH7
        ' SOTINVHM = MIDDLE:  SOTINVH2,SOTINVH6

        Dim ACCT_CODEs(7) As String
        Dim AMTs(7) As String
        Dim SQLW(7) As String
        Dim SQLJ(1, 7) As String

        For i As Integer = 0 To 7
            SQLW(i) = " and SOTINVH1.MARKET_TYPE = 'E'"

            Select Case i
                Case 0 ' CGS
                    ACCT_CODEs(i) = "" _
                    & "DECODE (ORDR_TYPE_CODE,'P',SOTPARM1.SO_PARM_ACCT_RETURNS_CGS" _
                    & ", ICTCREG1.ACCT_CGS)"

                    AMTs(i) = "NVL(SOTINVHD.STD_COST_EXT,0)"
                    SQLJ(0, i) = ",EMP.ICTCREG1"
                    SQLJ(1, i) = " and ICTCREG1.CON_REG_IND = SOTINVHD.CON_REG_IND"

                Case 1 ' Sales
                    ACCT_CODEs(i) = "" _
                    & "DECODE (ORDR_TYPE_CODE,'P',SOTPARM1.SO_PARM_ACCT_RETURNS_SLS" _
                    & " ,'S', DECODE (SOTINVHD.CON_REG_IND,'R'," _
                    & "       SOTPARM1.SO_PARM_ACCT_SALES_REG,SOTPARM1.SO_PARM_ACCT_SALES_CON)" _
                    & " ,'D', DECODE (SOTINVHD.CON_REG_IND,'R'," _
                    & "       SOTPARM1.SO_PARM_ACCT_SALES_REG,SOTPARM1.SO_PARM_ACCT_SALES_CON)" _
                    & " ,'C', DECODE (SOTINVHD.CON_REG_IND,'R'," _
                    & "       SOTPARM1.SO_PARM_ACCT_RETURNS_REG,SOTPARM1.SO_PARM_ACCT_RETURNS_CON)" _
                    & " ,NULL)"

                    AMTs(i) = "" _
                    & " DECODE (SOTINVH1.MARKET_TYPE, 'E'," _
                    & "    NVL(SOTINVHP.ORDR_PRICE_GRS,0) - NVL(SOTINVHP.BRKR_RATE,0)" _
                    & "  - NVL(SOTINVHP.REBATE,0) - NVL(SOTINVHP.FUND_RATE,0)" _
                    & "  - (CASE WHEN NVL(SOTINVHP.ORDR_PRICE_GRS,0) <> 0 THEN NVL(SOTINVH1.FRT_RATE,0) ELSE 0 END) " _
                    & "  - NVL(SOTINVHP.SVC_CHG_RATE,0)" _
                    & " , SOTINVHP.ORDR_PRICE_GRS)"

                    SQLW(i) = ""

                Case 2 ' Brokerage
                    ACCT_CODEs(i) = "SOTPARM1.SO_PARM_ACCT_BRKR"
                    AMTs(i) = "NVL(SOTINVHP.BRKR_RATE,0)"

                Case 3 ' Rebates
                    ACCT_CODEs(i) = "SOTPARM1.SO_PARM_ACCT_REBATE"
                    AMTs(i) = "NVL(SOTINVHP.REBATE,0)"

                Case 4 ' Funds
                    ACCT_CODEs(i) = "SOTPARM1.SO_PARM_ACCT_BRKR"
                    AMTs(i) = "NVL(SOTINVHP.BRKR_RATE,0)"

                Case 5 ' Freight
                    ACCT_CODEs(i) = "SOTPARM1.SO_PARM_ACCT_FRT"
                    AMTs(i) = "(CASE WHEN NVL(SOTINVHP.ORDR_PRICE_GRS,0) <> 0 THEN NVL(SOTINVH1.FRT_RATE,0) ELSE 0 END)"

                Case 6 ' Service Charges
                    ACCT_CODEs(i) = "SOTSVCG1.SVC_CHG_ACCT"
                    AMTs(i) = "NVL(SOTINVHP.SVC_CHG_RATE,0)"
                    SQLJ(0, i) = ",EMP.SOTSVCG1,EMP.POTCATG1"
                    SQLJ(1, i) = " and SOTSVCG1.SVC_CHG_CODE = SOTINVHP.SVC_CHG_CODE" _
                               & " and POTCATG1.COST_CATGY_CODE = SOTSVCG1.COST_CATGY_CODE"

                Case 7 ' Allowances
                    'Stop ' NEED TO GO TO CONTRACT

                    ACCT_CODEs(i) = "SOTPARM1.SO_PARM_ACCT_BRKR"
                    AMTs(i) = "NVL(SOTINVHP.BRKR_RATE,0)"
                    SQLJ(0, i) = ",EMP.SOTALOW1"
                    SQLJ(1, i) = " and SOTALOW1.SVC_CHG_CODE = SOTINVHP.ALLOW_CODE"

            End Select

            If i <> 0 Then
                AMTs(i) = " -1 * NVL(SOTINVHD.SO_LOT_UNITS,0) * " & AMTs(i)
            End If

        Next

        For I As Integer = 0 To UBound(ACCT_CODEs)
            If I = 4 Or I = 7 Then Continue For

            Dim SQL As String = ""
            SQL = "INSERT INTO " & GLTDETL2 _
            & " SELECT SOTINVH1.OPS_YYYYPP, 'OPSJ'" _
            & " , " & ACCT_CODEs(I) & " ACCT_CODE " _
            & " , SOTINVH1.ORDR_DIV_CODE SEG2_CODE, '000' SEG3_CODE, '000' SEG4_CODE" _
            & " , 'S' DETL_CTL_TYPE, SOTINVHD.SO_ORDER_NO DETL_CTL_NO" _
            & " , SOTINVHD.SO_ORDER_LNO DETL_CTL_LNO, SOTINVHD.SO_LOT_LNO DETL_CTL_SUB_LNO" _
            & " , SOTINVH1.ORDR_INV_REG_XNO DETL_EXE_NO" _
            & " , " & AMTs(I) & " DETL_POSTING_AMT " _
            & " , SOTINVH1.CUST_CODE, ICTLOTD1.VEND_CODE, ICTLOTD1.PROD_CODE, ICTLOTD1.SIZE_CODE" _
            & " , ICTLOTD1.ORIG_CODE, ICTLOTD1.BRAND_CODE, ICTLOTD1.WHSE_CODE, ICTLOTD1.LOT_NO, ICTLOTD1.LOT_SEQ_NO, NULL FORCE_BALANCE" _
            & " FROM EMP.SOTINVH1," & IIf(S_or_C = "S", "EMP.SOTINVH2,EMP.SOTINVH3", "EMP.SOTINVH6,EMP.SOTINVH7") & ",EMP.SOTPARM1,EMP.ICTLOTD1" _
            & SQLJ(0, I) _
            & "  WHERE SOTINVH1.OPS_YYYYPP = '" & YP & "'" _
            & " AND SOTINVHM.SO_ORDER_NO = SOTINVH1.SO_ORDER_NO" _
            & " AND SOTINVHD.SO_ORDER_NO = SOTINVHM.SO_ORDER_NO" _
            & " AND SOTINVHD.SO_ORDER_LNO = SOTINVHM.SO_ORDER_LNO" _
            & " AND ICTLOTD1.WHSE_CODE = SOTINVHD.WHSE_CODE" _
            & " AND ICTLOTD1.LOT_NO = SOTINVHD.LOT_NO" _
            & " AND ICTLOTD1.LOT_SEQ_NO = SOTINVHD.LOT_SEQ_NO" _
            & " AND SOTINVH1.ORDR_TYPE_CODE IN ('S','P','C','D')" _
            & " AND SOTINVHD.SO_LOT_UNITS <> 0" _
            & " AND " & AMTs(I) & " <> 0" _
            & SQLW(I) & SQLJ(1, I) _
            & IIf(S_or_C = "S", "", " AND (SOTINVH1.ORDR_TYPE_CODE <> 'C' OR NVL(SOTINVH7.CLAIM_IND,'0') <> '1')") _
            & " AND SOTPARM1.SO_PARM_KEY = 'Z'"

            If S_or_C = "S" Then
                SQL = Replace(SQL, "SOTINVHD", "SOTINVH3")
                SQL = Replace(SQL, "SOTINVHP", "SOTINVH2")
                SQL = Replace(SQL, "SOTINVHM", "SOTINVH2")
            Else
                SQL = Replace(SQL, "SOTINVHD", "SOTINVH7")
                SQL = Replace(SQL, "SOTINVHP", "SOTINVH7")
                SQL = Replace(SQL, "SOTINVHM", "SOTINVH6")
                SQL = Replace(SQL, "SO_LOT_LNO", "SO_SUB_LNO")
                SQL = Replace(SQL, "SO_LOT_UNITS", "CHG_UNITS")
            End If

            If I = 0 Then
                SQL = Replace(SQL, "'OPSJ'", "'OPCG'")
            End If

            ASCDATA1.ExecuteSQL(SQL)
            Console.WriteLine(SQL)

            Select Case I
                Case 2 ' Brokerage
                    SQL = Replace(SQL, "SO_PARM_ACCT_BRKR", "SO_PARM_ACCT_BRKR_EXP")

                Case 3 ' Rebates
                    SQL = Replace(SQL, "SO_PARM_ACCT_REBATE", "SO_PARM_ACCT_REBATE_EXP")

                Case 4 ' Funds
                    SQL = Replace(SQL, "ACCT_CODE_REC", "ACCT_CODE_EXP")

                Case 5 ' Freight
                    SQL = Replace(SQL, "SO_PARM_ACCT_FRT", "SO_PARM_ACCT_FRT_EXP")

                Case 6 ' Service Charges
                    SQL = Replace(SQL, "SOTSVCG1.SVC_CHG_ACCT", "POTCATG1.ACCT_CODE_VAR_R")

                Case Else
                    SQL = ""

            End Select

            If SQL <> "" Then
                SQL = Replace(SQL, "'OPSJ'", "'OPXA'")
                SQL = Replace(SQL, " -1 * NVL(SOTINVH", "NVL(SOTINVH")
                ASCDATA1.ExecuteSQL(SQL)
                Console.WriteLine(SQL)
            End If
        Next

    End Sub

    Private Sub optNRC_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optNRC.ValueChanged
        If dst.Tables.Count = 0 Then Exit Sub

        SplitContainer1.Panel2Collapsed = (optNRC.Value <> "R")
        grdGLTGPANC.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = (optNRC.Value <> "R")
    End Sub

    Sub Build_GLTGPANR(ByVal COLUMN_NAMEs As List(Of String))
        Try
            dst.Relations.Remove("GLTGPANR")
            dst.Relations.Remove("GLTDETL2")
        Catch ex As Exception
        End Try

        grdGLTGPANR.DataSource = Nothing
        dst.Tables("GLTDETL2").Constraints.Clear()
        dst.Tables("GLTGPANR").Constraints.Clear()
        'For i As Integer = 0 To dst.Tables("GLTDETL2").Constraints.Count - 1 to 0 step -1
        '    Dim c As Constraint = dst.Tables("GLTDETL2").Constraints(i)
        '    dst.Tables("GLTDETL2").Constraints.Remove(c)
        'Next

        dst.EnforceConstraints = False

        Dim DC1 As DataColumn()
        Dim DC2 As DataColumn()
        ReDim DC1(COLUMN_NAMEs.Count - 1)
        ReDim DC2(COLUMN_NAMEs.Count - 1)

        Dim C() As String
        ReDim C(COLUMN_NAMEs.Count - 1)

        With dst.Tables("GLTGPANR")
            .Rows.Clear()
            .AcceptChanges()
            .Columns.Clear()
            Dim i As Integer = 0
            For Each COLUMN_NAME As String In COLUMN_NAMEs
                .Columns.Add(COLUMN_NAME)
                C(i) = COLUMN_NAME
                DC1(i) = dst.Tables("GLTGPANR").Columns(COLUMN_NAME)
                DC2(i) = dst.Tables("GLTDETL2").Columns(COLUMN_NAME)
                i = i + 1
            Next
            '.Columns.Add("CUST_CODE")
            '.Columns.Add("PROD_CODE")
            .Columns.Add("AMT")
        End With

        dst.Relations.Add("GLTGPANR", _
        dst.Tables("GLTSTMTX").Columns("ACCT_CODE"), _
        dst.Tables("GLTGPANR").Columns("ACCT_CODE"))

        dst.Relations.Add("GLTDETL2", DC1, DC2)

        dst.Tables("GLTGPANR").Columns("AMT").Expression = "SUM(CHILD.DETL_POSTING_AMT)"

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("GLTDETL2"), C).Rows
            dst.Tables("GLTGPANR").Rows.Add(row.ItemArray)
        Next

        'grdGLTGPANR.DataMember = "GLTFINR2.GLTSTMTX.GLTGPANR"
        ''grdGLTGPANR.DataMember = "GLTGPANR"
        ''grdGLTGPANR.DataSource = dst
        'grdGLTGPANR.DataSource = dst.Tables("GLTGPANR")

        Dim view1 As New DataView(dst.Tables("GLTGPANR"))
        Dim source1 As New BindingSource()
        source1.DataSource = view1

        grdGLTGPANR.DataSource = source1
        Call ASCMAIN1.grdInitializeLayout(grdGLTGPANR)

        'source1.Filter = "artist = 'Dave Matthews'"







        Call Create_Summary(grdGLTGPANR, "ACCT_CODE", "Count")
        Call Create_Summary(grdGLTGPANR, "AMT")

        'grdGLTGPANJ.DataMember = "GLTFINR2.GLTSTMTX.GLTGPANR.GLTDETL2"
        grdGLTGPANJ.DataMember = "GLTGPANR.GLTDETL2"
        grdGLTGPANJ.DataSource = dst

        Call Create_Summary(grdGLTGPANJ, "DETL_POSTING_AMT")
        'dst.Tables("GLTGPANC").Rows.Find(cc.Key)
        Dim COLUMN_CAPTION As String = ""
        For Each cc As UltraWinGrid.UltraGridColumn In grdGLTGPANR.DisplayLayout.Bands(0).Columns
            If cc.Key <> "GLTDETL2" Then

                If cc.Key = "ACCT_CODE" Then
                    COLUMN_CAPTION = "Acct"
                ElseIf cc.Key = "AMT" Then
                    COLUMN_CAPTION = "Amount"
                Else
                    COLUMN_CAPTION = dst.Tables("GLTGPANC").Select("COLUMN_NAME = '" & cc.Key & "'", "")(0).Item("COLUMN_CAPTION") & ""
                End If
                cc.Header.Caption = COLUMN_CAPTION
            End If

        Next
    End Sub

    Private Sub grdGLTFINR2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTFINR2.AfterRowActivate
        If grdGLTGPANR.DataSource Is Nothing Then
            Exit Sub
        End If

        If grdGLTFINR2.ActiveRow.IsDataRow Then
            If grdGLTFINR2.ActiveRow.Band.Key = "GLTSTMTX" Then
                ''Dim x As BindingSource = grdGLTGPANR.DataSource
                'Dim Xc As CurrencyManager = Me.BindingContext(dst.Tables("GLTSTMTX"))
                'Stop
                ''Dim xx As BindingSource = grdGLTGPANR.DataMember.
                'Stop
                Dim source1 As BindingSource = grdGLTGPANR.DataSource
                source1.Filter = "ACCT_CODE = '" & grdGLTFINR2.ActiveRow.Cells("ACCT_CODE").Text & "'"
                'grdGLTGPANR.DataSource = New DataView(DirectCast(grdGLTGPANR.DataSource, DataTable), "ACCT_CODE = '" & grdGLTFINR2.ActiveRow.Cells("ACCT_CODE").Text & "'", "", DataViewRowState.CurrentRows)
            End If
        End If
    End Sub

    Private Sub grdGLTFINR2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTFINR2.InitializeLayout

    End Sub

    Private Sub cmdApply_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdApply.Click
        Apply_Breakdown()
    End Sub

    Sub Apply_Breakdown()

        If COLS.Count <> 0 Then
            For Each COLUMN_NAME As String In COLS.Keys
                dst.Tables("GLTFINR2").Columns.Remove(COLUMN_NAME)
            Next
            COLS.Clear()
        End If

        SplitContainer1.Panel2Collapsed = (optNRC.Value <> "R")

        If optNRC.Value <> "N" Then
            Dim C As New List(Of String)
            C.Add("ACCT_CODE")
            For Each ROW As DataRow In dst.Tables("GLTGPANC").Select("SELECTED = '1'")
                C.Add(ROW.Item("COLUMN_NAME"))
            Next

            Call Build_GLTGPANR(C)
            If optNRC.Value = "C" Then
                For Each row As DataRow In ASCDATA1.SelectDistinct("GLTGPANR", C(1)).ROWS

                Next
            End If
        End If
    End Sub

End Class