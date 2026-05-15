Public Class SORLOST1

    Dim SOTLOST1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")
      
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        Set_cmbYW("RYW0", ASCMAIN1.CYW, -60, 0, 0)
        Set_cmbYW_Child("RYW1", 60, "RYW0", 0)

        grpWEEK_RANGE.Left = grpPERIOD_RANGE.Left
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        Dim DTE1 As Date = Nothing
        Dim DTE2 As Date = Nothing
        If optRange.Value = "W" Then
            Dim rowGLTPARM3 As DataRow
            rowGLTPARM3 = LookUp("GLTPARM3", RYW0)
            DTE1 = CDate(rowGLTPARM3.Item("WEEK_END_DATE")).AddDays(-6)
            rowGLTPARM3 = LookUp("GLTPARM3", RYW1)
            DTE2 = CDate(rowGLTPARM3.Item("WEEK_END_DATE"))
        End If
        ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
            & ", SOTORDR1.CUST_CODE, ARTCUST2.CUST_STORE_NAME, SOTORDR1.CUST_STORE_NO, SOTSREP1.SREP_NAME" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, SOTORDR2.ITEM_DESC, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" & vbCrLf _
            & ", SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ORDR_AMT_CANC" & vbCrLf _
            & ", SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & ", SOTORDR1.SREP_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.REASON_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_SOURCE, SOTORDR2.ORDR_STATUS, SOTORDR2.ORDR_RELEASE" & vbCrLf _
            & " from SOTORDR1,SOTORDR2,ARTCUST2,SOTSREP1" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE (+) = SOTORDR1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO (+) = SOTORDR1.CUST_STORE_NO" & vbCrLf _
            & "   and SOTSREP1.SREP_CODE = SOTORDR1.SREP_CODE" & vbCrLf _
            & "   and SOTORDR2.ORDR_QTY_CANC <> 0" & vbCrLf _
            & IIf(ASCMAIN1.CLIENT = "INT", "", "   and NVL(SOTORDR2.ORDR_RELEASE,'?') <> 'D'" & vbCrLf) _
            & "   and SOTORDR1.ORDR_STATUS IN ('F','C')" & vbCrLf _
            & IIf(optRange.Value = "P", _
                  "   and SOTORDR1.ORDR_YYYYPP_CLOSED between '" & RYP0 & "' and '" & RYP1 & "'", _
                  "   and SOTORDR1.ORDR_DATE_CLOSED between '" & Format(DTE1, "dd-MMM-yyyy") & "' and '" & Format(DTE2, "dd-MMM-yyyy") & "'")
        ' 11/15/16 - LM WANTS LOST SALES TO INCLUDE THE D CODE
        ' Dim SOTLOST1 As String
        SOTLOST1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTLOST1 & " Add Primary Key (ORDR_NO,ORDR_LNO)")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""
        'Dim sqlX As String = ""
        'Dim sqlYP As String = ""
       
        ASCMAIN1.Progress("Lost Sales")
        MyBase.Get_SQL("*", SOTLOST1)

        'sqlX = "NVL(SOTORDR2.ORDR_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        'sqlYP = "SOTORDR1.ORDR_YYYYPP_BOOKED"

        sql_Data = "" _
            & ", SUM (ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM (ORDR_AMT) ORDR_AMT" & vbCrLf _
            & ", SUM (ORDR_AMT_SHIP) ORDR_AMT_SHIP" & vbCrLf _
            & ", SUM (ORDR_AMT_CANC) ORDR_AMT_CANC" & vbCrLf

        sql_Cols = "" _
            & ",ORDR_QTY,ORDR_QTY_SHIP,ORDR_QTY_CANC,ORDR_AMT,ORDR_AMT_SHIP,ORDR_AMT_CANC"

        sql_filter = "" _
            & ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", SOTLOST1.ORDR_NO, SOTLOST1.ORDR_LNO" & vbCrLf _
            & sql_Data _
            & " from " & SOTLOST1 & " SOTLOST1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols & vbCrLf _
            & IIf(sql_GROUP_BY_cols = "", "", ",") & "SOTLOST1.ORDR_NO, SOTLOST1.ORDR_LNO"

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Select * from " & SOTLOST1 & " SOTLOST1"

        ASCMAIN1.sql = "Select SOTLOST1.*" & vbCrLf _
            & " from " & SOTLOST1 & " SOTLOST1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter)

        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTLOST1", 2))

        ' Eliminate 0s

        'Dim sqlz As String = ""
        'For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
        '    sqlz &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        'Next
        'ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlz))

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = "Lost Sales from Bookings from " & RYPLEGEND0 & " thru " & RYPLEGEND1
        CR_params.Add("DTL", "Y")
        Generate_Report(RPT, , SUBT)
        CR_params.Add("DTL", "N")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        'If eItemKey = "Proceed" Then
        '    If Absx1.cmbFor("RYP").Value & "" = "" Then
        '        EMsg &= vbCr & "You must Specify a Reporting Period"
        '    End If
        'End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'dst.Tables("ASTSRPT1").Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
        'dst.Tables("ASTSRPT1").Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"

        Load_grdSOTLOST1()
    End Sub

    Private Sub optRange_ValueChanged(sender As Object, e As EventArgs) Handles optRange.ValueChanged
        grpPERIOD_RANGE.Visible = (optRange.Value = "P")
        grpWEEK_RANGE.Visible = (optRange.Value = "W")
    End Sub

    Sub Load_grdSOTLOST1()

        'ASCMAIN1.sql = "Select SOTLOST1.* from " & SOTLOST1 & " SOTLOST1"
        'dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTLOSTX", 0))


        grdSOTLOST1.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTLOST1.DataSource = dst.Tables("SOTLOST1")
        'With grdSOTLOST1.DisplayLayout.Bands(0)
        '    .Columns("CUST_CODE").Header.Caption = "Customer Code"
        'End With   

        grdSOTLOST1.DisplayLayout.Bands(0).Summaries.Clear()
        Create_Summary(grdSOTLOST1, New String() {"ORDR_QTY", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_AMT", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})

        Sort_grdColumns(grdSOTLOST1, "CUST_CODE, ORDR_CUST_PO")
        Show_Filter(grdSOTLOST1, True)
        grdSOTLOST1.DisplayLayout.GroupByBox.Hidden = False
        'grdSOTORDRC.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)
        'grdSOTORDRC.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_CUST_PO", False, True)
        'grdSOTLOST1.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Hidden = True
        grdSOTLOST1.Visible = True
        ' btnExcelExport.Visible = True
    End Sub
End Class