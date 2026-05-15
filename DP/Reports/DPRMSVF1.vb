Public Class DPRMSVF1

    Dim YPP(4, 1) As String
    Dim YPF(4, 1) As String
    Dim YPFD(4, 1) As Date

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()
        ' taking MARKET_CODE out of sort/filter. Maybe should be used only as a filter. To restore, need to resolve how to get each of the following by market: PO, On Hand; sales and forecasts are easy.
        ' adding MARKET_CODE back in as a filter

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ASCMAIN1.sql = "Select ITEM_CODE" & vbCrLf _
            & ", 0 QTY_SHP_P3, 0 QTY_SHP_P2, 0 QTY_SHP_P1, 0 QTY_SHP_P0, 0 QTY_SHP_PTOT" & vbCrLf _
            & ", 0 QTY_BEG, 0 QTY_ONH, 0 QTY_OPEN, 0 QTY_PICK" & vbCrLf _
            & ", 0 QTY_OPEN_F0, 0 QTY_OPEN_F1, 0 QTY_OPEN_F2, 0 QTY_OPEN_F3, 0 QTY_OPEN_F4, 0 QTY_OPEN_FTOT" & vbCrLf _
            & ", 0 QTY_FC_P3, 0 QTY_FC_P2, 0 QTY_FC_P1, 0 QTY_FC_P0, 0 QTY_FC_PTOT" & vbCrLf _
            & ", 0 QTY_FC_F0, 0 QTY_FC_F1, 0 QTY_FC_F2, 0 QTY_FC_F3, 0 QTY_FC_F4, 0 QTY_FC_FTOT" & vbCrLf _
            & ", 0 QTY_OPO_F0, 0 QTY_OPO_F1, 0 QTY_OPO_F2, 0 QTY_OPO_F3, 0 QTY_OPO_F4, 0 QTY_OPO_FTOT" & vbCrLf _
            & ", 0 QTY_PLN_F0, 0 QTY_PLN_F1, 0 QTY_PLN_F2, 0 QTY_PLN_F3, 0 QTY_PLN_F4, 0 QTY_PLN_FTOT" & vbCrLf _
            & ", 0 QTY_EOH_F0, 0 QTY_EOH_F1, 0 QTY_EOH_F2, 0 QTY_EOH_F3, 0 QTY_EOH_F4" & vbCrLf _
            & ", 0 QTY_POS_F0, 0 QTY_POS_F1, 0 QTY_POS_F2, 0 QTY_POS_F3, 0 QTY_POS_F4" & vbCrLf _
            & ", 0 QTY_FC_VAR, 0.01 QTY_FC_VAR_PCT" & vbCrLf _
            & " from ICTITEM1 where ROWNUM < 1"

        Dim DPTMSVF1 As String = ASCMAIN1.Temp_Table

 
        For i As Int16 = 0 To 4
            YPP(i, 0) = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i)
            YPP(i, 1) = ASCMAIN1.Get_Legend(YPP(i, 0), False, True)
            YPF(i, 0) = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, i)
            YPF(i, 1) = ASCMAIN1.Get_Legend(YPF(i, 0), False, True)

            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YPF(i, 0))
            YPFD(i, 1) = rowGLTPARM2.Item("PRD_END_DATE")
            If i > 0 Then
                YPFD(i, 0) = YPFD(i - 1, 1).AddDays(1)
            End If
        Next

        ' Inventory Status

        ASCMAIN1.sql = "Insert into " & DPTMSVF1 & " DPTMSVF1" & vbCrLf _
            & " (ITEM_CODE,QTY_OPEN,QTY_PICK,QTY_ONH) " & vbCrLf _
            & " Select ICTSTAT2.ITEM_CODE" & vbCrLf _
            & ", SUM (ICTSTAT2.WHSE_QTY_OPEN) QTY_OPEN" & vbCrLf _
            & ", SUM (ICTSTAT2.WHSE_QTY_PICK) QTY_PICK" & vbCrLf _
            & ", SUM (ICTSTAT2.WHSE_QTY_ON_HAND) QTY_ONH" & vbCrLf _
            & " from ICTSTAT2,ICTWHSE1 " & vbCrLf _
            & " where ICTWHSE1.WHSE_CODE = ICTSTAT2.WHSE_CODE " & vbCrLf _
            & "   and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" _
            & " group by ICTSTAT2.ITEM_CODE"
        ASCDATA1.ExecuteSQL()

        ' Shipments

        ASCMAIN1.sql = "Insert into " & DPTMSVF1 & " DPTMSVF1" & vbCrLf _
            & " (ITEM_CODE,QTY_SHP_P3,QTY_SHP_P2,QTY_SHP_P1,QTY_SHP_P0) " & vbCrLf _
            & " Select ITEM_CODE" & vbCrLf _
            & ", SUM (DECODE(ORDR_YYYYPP_UPDATED,'" & YPP(3, 0) & "',ORDR_QTY_SHIP,0)) QTY_SHP_P3" & vbCrLf _
            & ", SUM (DECODE(ORDR_YYYYPP_UPDATED,'" & YPP(2, 0) & "',ORDR_QTY_SHIP,0)) QTY_SHP_P2" & vbCrLf _
            & ", SUM (DECODE(ORDR_YYYYPP_UPDATED,'" & YPP(1, 0) & "',ORDR_QTY_SHIP,0)) QTY_SHP_P1" & vbCrLf _
            & ", SUM (DECODE(ORDR_YYYYPP_UPDATED,'" & YPP(0, 0) & "',ORDR_QTY_SHIP,0)) QTY_SHP_P0" & vbCrLf _
            & " from SOTINVH2 " & vbCrLf _
            & " where ORDR_YYYYPP_UPDATED >= '" & YPP(3, 0) & "' and ORDR_YYYYPP_UPDATED <= '" & YPP(0, 0) & "'" & vbCrLf _
            & IIf(Absx1.optFor("GN").Value = "G", " and INV_TYPE = 'I'" & vbCrLf, "") _
            & " group by ITEM_CODE"
        ASCDATA1.ExecuteSQL()


        ' Open Sales Orders

        ASCMAIN1.sql = "Insert into " & DPTMSVF1 & " DPTMSVF1" & vbCrLf _
            & " (ITEM_CODE,QTY_OPEN_F0,QTY_OPEN_F1,QTY_OPEN_F2,QTY_OPEN_F3,QTY_OPEN_F4) " & vbCrLf _
            & " Select SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN                                                                            SOTORDR1.ORDR_SHIP_DATE <= '" & Format(YPFD(0, 1), "dd-MMM-yyyy") & "' THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) QTY_OPEN_F0" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE >= '" & Format(YPFD(1, 0), "dd-MMM-yyyy") & "' AND SOTORDR1.ORDR_SHIP_DATE <= '" & Format(YPFD(1, 1), "dd-MMM-yyyy") & "' THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) QTY_OPEN_F1" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE >= '" & Format(YPFD(2, 0), "dd-MMM-yyyy") & "' AND SOTORDR1.ORDR_SHIP_DATE <= '" & Format(YPFD(2, 1), "dd-MMM-yyyy") & "' THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) QTY_OPEN_F2" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE >= '" & Format(YPFD(3, 0), "dd-MMM-yyyy") & "' AND SOTORDR1.ORDR_SHIP_DATE <= '" & Format(YPFD(3, 1), "dd-MMM-yyyy") & "' THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) QTY_OPEN_F3" & vbCrLf _
            & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE >= '" & Format(YPFD(4, 0), "dd-MMM-yyyy") & "'                                                                            THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) QTY_OPEN_F4" & vbCrLf _
            & " from SOTORDR1,SOTORDR2 " & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " group by SOTORDR2.ITEM_CODE"
        ASCDATA1.ExecuteSQL()


        ' Open Purchase Orders

        ASCMAIN1.sql = "Insert into " & DPTMSVF1 & " DPTMSVF1" & vbCrLf _
            & " (ITEM_CODE,QTY_OPO_F0,QTY_OPO_F1,QTY_OPO_F2,QTY_OPO_F3,QTY_OPO_F4) " & vbCrLf _
            & " Select POTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN                                                                                  POTORDR2.PO_DATE_REQUIRED <= '" & Format(YPFD(0, 1), "dd-MMM-yyyy") & "' THEN NVL(POTORDR2.PO_QTY_OPN,0) ELSE 0 END) QTY_OPO_F0" & vbCrLf _
            & ", SUM (CASE WHEN POTORDR2.PO_DATE_REQUIRED >= '" & Format(YPFD(1, 0), "dd-MMM-yyyy") & "' AND POTORDR2.PO_DATE_REQUIRED <= '" & Format(YPFD(1, 1), "dd-MMM-yyyy") & "' THEN NVL(POTORDR2.PO_QTY_OPN,0) ELSE 0 END) QTY_OPO_F1" & vbCrLf _
            & ", SUM (CASE WHEN POTORDR2.PO_DATE_REQUIRED >= '" & Format(YPFD(2, 0), "dd-MMM-yyyy") & "' AND POTORDR2.PO_DATE_REQUIRED <= '" & Format(YPFD(2, 1), "dd-MMM-yyyy") & "' THEN NVL(POTORDR2.PO_QTY_OPN,0) ELSE 0 END) QTY_OPO_F2" & vbCrLf _
            & ", SUM (CASE WHEN POTORDR2.PO_DATE_REQUIRED >= '" & Format(YPFD(3, 0), "dd-MMM-yyyy") & "' AND POTORDR2.PO_DATE_REQUIRED <= '" & Format(YPFD(3, 1), "dd-MMM-yyyy") & "' THEN NVL(POTORDR2.PO_QTY_OPN,0) ELSE 0 END) QTY_OPO_F3" & vbCrLf _
            & ", SUM (CASE WHEN POTORDR2.PO_DATE_REQUIRED >= '" & Format(YPFD(4, 0), "dd-MMM-yyyy") & "'                                                                                  THEN NVL(POTORDR2.PO_QTY_OPN,0) ELSE 0 END) QTY_OPO_F4" & vbCrLf _
            & " from POTORDR2 " & vbCrLf _
            & " where NVL(POTORDR2.PO_QTY_OPN,0) <> 0" & vbCrLf _
            & " group by POTORDR2.ITEM_CODE"
        ASCDATA1.ExecuteSQL()


        ' Forecasts

        ASCMAIN1.sql = "Insert into " & DPTMSVF1 & " DPTMSVF1" & vbCrLf _
            & " (ITEM_CODE,QTY_FC_P3,QTY_FC_P2,QTY_FC_P1,QTY_FC_P0,QTY_FC_F0,QTY_FC_F1,QTY_FC_F2,QTY_FC_F3,QTY_FC_F4) " & vbCrLf _
            & " Select ITEM_CODE" & vbCrLf
        For I As Int16 = 3 To 1 Step -1
            ASCMAIN1.sql &= ", SUM (CASE WHEN OPS_YYYYPP = '" & YPP(I, 0) & "' AND OPS_YYYYPP = OPS_YYYYPP_FC THEN FORECAST ELSE 0 END) QTY_FC_P" & Format(I, "0") & vbCrLf
        Next
        ASCMAIN1.sql &= ", SUM (CASE WHEN OPS_YYYYPP = '" & YPP(0, 0) & "' AND (OPS_YYYYPP_FC = OPS_YYYYPP OR OPS_YYYYPP_FC = '000000') THEN FORECAST ELSE 0 END) QTY_FC_P0" & vbCrLf
        ASCMAIN1.sql &= ", SUM (CASE WHEN OPS_YYYYPP = '" & YPF(0, 0) & "' AND (OPS_YYYYPP_FC = OPS_YYYYPP OR OPS_YYYYPP_FC = '000000') THEN FORECAST ELSE 0 END) QTY_FC_F0" & vbCrLf
        For I As Int16 = 1 To 4
            ASCMAIN1.sql &= ", SUM (CASE WHEN OPS_YYYYPP = '" & YPF(0, 0) & "' AND OPS_YYYYPP_FC = '" & YPF(I, 0) & "' THEN FORECAST ELSE 0 END) QTY_FC_F" & Format(I, "0") & vbCrLf
        Next
        ASCMAIN1.sql &= " from DPTITMF1 " & vbCrLf _
            & " where OPS_YYYYPP >= '" & YPP(3, 0) & "' and OPS_YYYYPP <= '" & YPP(0, 0) & "'" & vbCrLf _
            & " group by ITEM_CODE"
        ASCDATA1.ExecuteSQL()



        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM" _
            & ", ICTCOLL1.BRAND_CODE, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_SNU_CODE" _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_COST_MAKE_BUY, ICTITEM1.ITEM_COST_STD" _
            & " from ICTITEM1,ICTCOLL1 " _
            & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & DPTMSVF1 & ")" _
            & " and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)
        Fill_Records("ICTITEM1")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*", DPTMSVF1)

        Dim SOURCE_TABLE_NAME As String = DPTMSVF1
        Dim sql_Data As String = ""
        Dim DATA_COLUMN_NAMEs As String = ""

        For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
            sql_Data &= ", SUM (" & COLUMN_NAME & ") " & COLUMN_NAME
            DATA_COLUMN_NAMEs &= "," & COLUMN_NAME
        Next

        sql = "Select " & sql_SELECT_cols & vbCrLf & Replace(COLUMN_NAMEs_appended, ",", ",DPTMSVF1" & ".") & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " DPTMSVF1 " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols & Replace(COLUMN_NAMEs_appended, ",", ",DPTMSVF1" & ".")

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended & DATA_COLUMN_NAMEs & ")" & vbCrLf _
        & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()



        ' Post Preparation Updates

        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set " _
        & "  QTY_OPEN_FTOT = (NVL(QTY_OPEN_F0,0) + NVL(QTY_OPEN_F1,0) + NVL(QTY_OPEN_F2,0) + NVL(QTY_OPEN_F3,0) + NVL(QTY_OPEN_F4,0))" _
        & ", QTY_OPO_FTOT = (NVL(QTY_OPO_F0,0) + NVL(QTY_OPO_F1,0) + NVL(QTY_OPO_F2,0) + NVL(QTY_OPO_F3,0) + NVL(QTY_OPO_F4,0))" _
        & ", QTY_PLN_FTOT = (NVL(QTY_PLN_F0,0) + NVL(QTY_PLN_F1,0) + NVL(QTY_PLN_F2,0) + NVL(QTY_PLN_F3,0) + NVL(QTY_PLN_F4,0))" _
        & ", QTY_FC_FTOT = (NVL(QTY_FC_F0,0) + NVL(QTY_FC_F1,0) + NVL(QTY_FC_F2,0) + NVL(QTY_FC_F3,0) + NVL(QTY_FC_F4,0))" _
        & ", QTY_FC_PTOT = (NVL(QTY_FC_P1,0) + NVL(QTY_FC_P2,0) + NVL(QTY_FC_P3,0))" _
        & ", QTY_SHP_PTOT = (NVL(QTY_SHP_P1,0) + NVL(QTY_SHP_P2,0) + NVL(QTY_SHP_P3,0))"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set QTY_FC_VAR = QTY_SHP_PTOT - QTY_FC_PTOT"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from " & ASTSRPT1 _
            & " where NVL(QTY_OPEN_FTOT,0) = 0" _
            & " and NVL(QTY_SHP_PTOT,0) = 0 and NVL(QTY_SHP_P0,0) = 0" _
            & " and NVL(QTY_FC_PTOT,0) = 0 and NVL(QTY_FC_P0,0) = 0" _
            & " and NVL(QTY_FC_FTOT,0) = 0 " _
            & " and NVL(QTY_OPO_FTOT,0) = 0 and NVL(QTY_PLN_FTOT,0) = 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set " _
            & "  QTY_FC_VAR_PCT = 100 * QTY_FC_VAR / QTY_FC_PTOT" _
            & " where QTY_FC_PTOT <> 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set QTY_EOH_F0 = NVL(QTY_ONH,0) + NVL(QTY_SHP_P0,0) - GREATEST(NVL(QTY_FC_F0,0),NVL(QTY_SHP_P0,0)+NVL(QTY_OPEN_F0,0)) + NVL(QTY_OPO_F0,0)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set QTY_EOH_F1 = QTY_EOH_F0 - GREATEST(NVL(QTY_FC_F1,0),NVL(QTY_OPEN_F1,0)) + NVL(QTY_OPO_F1,0)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set QTY_EOH_F2 = QTY_EOH_F1 - GREATEST(NVL(QTY_FC_F2,0),NVL(QTY_OPEN_F2,0)) + NVL(QTY_OPO_F2,0)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set QTY_EOH_F3 = QTY_EOH_F2 - GREATEST(NVL(QTY_FC_F3,0),NVL(QTY_OPEN_F3,0)) + NVL(QTY_OPO_F3,0)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set QTY_EOH_F4 = QTY_EOH_F3 - GREATEST(NVL(QTY_FC_F4,0),NVL(QTY_OPEN_F4,0)) + NVL(QTY_OPO_F4,0)"
        ASCDATA1.ExecuteSQL()

        If Absx1.chkFor("OPEN").Checked Then
            ASCMAIN1.sql = "Delete from " & ASTSRPT1 & " where NVL(QTY_OPEN_FTOT,0) <= 0"
            ASCDATA1.ExecuteSQL()
        End If
        If Absx1.chkFor("OPO").Checked Then
            ASCMAIN1.sql = "Delete from " & ASTSRPT1 & " where NVL(QTY_OPO_FTOT,0) <= 0"
            ASCDATA1.ExecuteSQL()
        End If
        If Absx1.chkFor("OPEN_EXCEEDS_FC").Checked Then
            ASCMAIN1.sql = "Delete from " & ASTSRPT1 & " where NVL(QTY_OPEN_FTOT,0) <= NVL(QTY_FC_FTOT,0)"
            ASCDATA1.ExecuteSQL()
        End If
        If Val(Absx1.numFor("VARPCT").Value & "") <> 0 Then
            ASCMAIN1.sql = "Delete from " & ASTSRPT1 & " where ABS(NVL(QTY_FC_VAR_PCT,0)) < " & CStr(Absx1.numFor("VARPCT").Value)
            ASCDATA1.ExecuteSQL()
        End If
        If Absx1.chkFor("NEG_EOH").Checked Then
            Dim sqlw As String = ""
            For I = 0 To Val(Absx1.numFor("NEG_EOHX").Value & "")
                sqlw &= " and QTY_EOH_F" & Format(I, "0") & " >= 0"
            Next
            ASCMAIN1.sql = "Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlw)
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""

        CR_params.Add("MDP3", YPP(3, 1))
        CR_params.Add("MDP2", YPP(2, 1))
        CR_params.Add("MDP1", YPP(1, 1))
        CR_params.Add("MDP0", YPP(0, 1))

        CR_params.Add("MDF0", YPF(0, 1))
        CR_params.Add("MDF1", YPF(1, 1))
        CR_params.Add("MDF2", YPF(2, 1))
        CR_params.Add("MDF3", YPF(3, 1))
        CR_params.Add("MDF4", YPF(4, 1))

        CR_params.Add("VARPCT", 0)

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        Call ASCMAIN1.Progress("Report Calculations")

        'dst.Tables("ASTSRPT1").Columns("QTY_OPEN_TOT").Expression = "ISNULL(QTY_OPEN_F0,0)+ISNULL(QTY_OPEN_F1,0)+ISNULL(QTY_OPEN_F2,0)+ISNULL(QTY_OPEN_F3,0)+ISNULL(QTY_OPEN_F4,0)"

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If SEQs = 0 Then
                EMsg &= vbCr & "You must Select at least 1 field to Sort By"
            End If
        End If
    End Sub
End Class