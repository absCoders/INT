Public Class SORSORD1
    ' hardcoding below for walmart, kmart, sears

    Dim REPORT_DATE0 As Date
    Dim REPORT_DATE1 As Date

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
        REPORT_DATE0 = Absx1.dteFor("DTE0").Value
        REPORT_DATE1 = Absx1.dteFor("DTE1").Value

        Dim CONDITION As String = " SOTORDR1." _
            & IIf(optSR.Value = "R", "ORDR_DATE_RECD", "ORDR_SHIP_DATE") _
            & IIf(Format(REPORT_DATE0, "dd-MMM-yyyy") = Format(REPORT_DATE1, "dd-MMM-yyyy"), _
                  " = '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "'", _
                  " between '" & Format(REPORT_DATE0, "dd-MMM-yyyy") & "' and '" & Format(REPORT_DATE1, "dd-MMM-yyyy") & "'")

        Dim sqlw As String = " where " & CONDITION _
             & " and (SOTORDR1.ORDR_STATUS = 'O' OR SOTORDR1.ORDR_STATUS = 'P' OR SOTORDR1.ORDR_STATUS = 'F')"

        If Absx1.chkFor("CHKEDI_ONLY").Checked Then
            sqlw &= " and SOTORDR1.ORDR_SOURCE = 'E' "
        End If

        'sql &= SQL_in("CUST_CODE", "SOTORDR1.CUST_CODE")
        MyBase.Get_SQL("*")
        sqlw &= sql_WHERE & sql_JOIN




        ASCMAIN1.Progress("Order Summary", "")

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", NVL(SOTORDR1.CUST_DC_NO,'XXXXXX') CUST_DC_NO" & vbCrLf _
            & ", Min (SOTORDR1.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", Min (SOTORDR1.ORDR_DATE) ORDR_DATE" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_SHIP_DATE) ORDR_SHIP_DATE" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_CANCEL_DATE) ORDR_CANCEL_DATE" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_CUST_PO) ORDR_CUST_PO" & vbCrLf _
            & ", Count (*) ORDR_CNT" & vbCrLf _
            & ", MIN (SOTORDR1.SALES_DIVISION_CODE) SALES_DIVISION_CODE" & vbCrLf _
            & ", MIN (EDI_APPOINTMENT) EDI_APPOINTMENT" & vbCrLf _
            & ", MIN (ORDR_DEPT) ORDR_DEPT" & vbCrLf _
            & " from SOTORDR1" & vbCrLf _
            & sqlw & vbCrLf _
            & " GROUP BY SOTORDR1.ORDR_GROUP_NO, NVL(SOTORDR1.CUST_DC_NO,'XXXXXX')"
        Dim SOTORDR0 As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO,CUST_DC_NO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_QTY NUMBER (6,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_AMT NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add GROUP_SEQ NUMBER (6,0)")

        If Absx1.optFor("OPTCONS_GROUPS").Value = "C" Then
            ASCDATA1.ExecuteSQL("Update " & SOTORDR0 & " Set GROUP_SEQ = 0")
        ElseIf Absx1.optFor("OPTCONS_GROUPS").Value = "D" Then
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is Select X.*, ROWNUM GROUP_SEQ from " & vbCrLf _
                & "  (Select Distinct CUST_CODE, NVL(ORDR_DEPT,'X') ORDR_DEPT from " & SOTORDR0 & ") X;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set GROUP_SEQ = R1.GROUP_SEQ" & vbCrLf _
                & "    where CUST_CODE = R1.CUST_CODE and NVL(ORDR_DEPT,'X') = R1.ORDR_DEPT;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        ElseIf Absx1.optFor("OPTCONS_GROUPS").Value = "DC" Then
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is Select X.*, ROWNUM GROUP_SEQ from " & vbCrLf _
                & "  (Select Distinct CUST_CODE, NVL(CUST_DC_NO,'X') CUST_DC_NO from " & SOTORDR0 & ") X;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTORDR0 & " Set GROUP_SEQ = R1.GROUP_SEQ" & vbCrLf _
                & "    where CUST_CODE = R1.CUST_CODE and NVL(CUST_DC_NO,'X') = R1.CUST_DC_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        ElseIf Absx1.optFor("OPTCONS_GROUPS").Value = "N" Then
            ASCDATA1.ExecuteSQL("Update " & SOTORDR0 & " Set GROUP_SEQ = ROWNUM")
        End If

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO" _
            & ", NVL(SOTORDR1.CUST_DC_NO,'XXXXXX') CUST_DC_NO" _
            & ", SUM (ORDR_QTY) ORDR_QTY, SUM (ORDR_QTY * ORDR_UNIT_PRICE) ORDR_AMT" _
            & " from SOTORDR1,SOTORDR2" _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" _
            & "   and SOTORDR1.ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR0 & ")" _
            & " GROUP BY SOTORDR1.ORDR_GROUP_NO, NVL(SOTORDR1.CUST_DC_NO,'XXXXXX');"

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is " & ASCMAIN1.sql _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update " & SOTORDR0 & " Set " _
            & "     ORDR_QTY = R1.ORDR_QTY" _
            & "    ,ORDR_AMT = R1.ORDR_AMT" _
            & "    where ORDR_GROUP_NO = R1.ORDR_GROUP_NO" _
            & "      and CUST_DC_NO = R1.CUST_DC_NO;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & SOTORDR0, "SOTORDR0", 1))

        dst.Tables.Add(ASCDATA1.GetDataTable("Select Distinct CUST_CODE, GROUP_SEQ from " & SOTORDR0, "SOTSORD0", 2))


        ASCMAIN1.Progress("Order Information", "")
        ASCMAIN1.sql = "Select SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY ELSE 0 END) ORDR_QTY_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_OPEN ELSE 0 END) ORDR_QTY_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_PICK ELSE 0 END) ORDR_QTY_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_SHIP ELSE 0 END) ORDR_QTY_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_CANC ELSE 0 END) ORDR_QTY_CANC_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_OPEN_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_PICK_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_SHIP_X" & vbCrLf _
            & ", SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) ORDR_AMT_CANC_X" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from SOTORDR1, SOTORDR2, " & SOTORDR0 & " SOTORDR0" & vbCrLf _
            & sql_TABLE_NAMEs & vbCrLf _
            & sqlw & vbCrLf _
            & " and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " and SOTORDR0.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & " group by SOTORDR0.CUST_CODE, SOTORDR0.GROUP_SEQ, SOTORDR2.ITEM_CODE" & vbCrLf _
            & " having SUM (CASE WHEN " & CONDITION & " THEN SOTORDR2.ORDR_QTY ELSE 0 END) <> 0"
        Dim SOTSORD1 As String = ASCMAIN1.Temp_Table

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & SOTSORD1, "SOTSORD1", 0))

        ASCMAIN1.Progress("Master Files", "")
        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from SOTSDIV1", "SOTSDIV1", 1))

        ASCMAIN1.sql = "Select * from ARTCUST1" _
            & " where CUST_CODE in (Select DISTINCT CUST_CODE from " & SOTSORD1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ARTCUST1", 1))

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC " & vbCrLf _
            & " from ICTITEM1 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE in " & vbCrLf _
            & " (Select Distinct ITEM_CODE from " & SOTSORD1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTITEM1", 1))

        ASCMAIN1.Progress("Inventory Status", "")

        ASCMAIN1.sql = "Select ICTSTAT2.ITEM_CODE" & vbCrLf _
            & ", SUM(ICTSTAT2.WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND, SUM(ICTSTAT2.WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
            & ", SUM(ICTSTAT2.WHSE_QTY_ONPO) WHSE_QTY_ONPO" & vbCrLf _
            & ", SYSDATE SOON_SHIP_DATE" & vbCrLf _
            & " from ICTSTAT2 " & vbCrLf _
            & " where ICTSTAT2.STYLE_CODE in " & vbCrLf _
            & " (Select Distinct STYLE_CODE from " & SOTSORD1 & ")" & vbCrLf _
            & " GROUP BY ICTSTAT2.STYLE_CODE, ICTSTAT2.COLOR_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSORD2", 2))
        With dst.Tables("SOTSORD2")
            .Columns("WHSE_QTY_PICK").ReadOnly = False
            .Columns.Add("SOON_SHIP_QTY", GetType(System.Int64))
            .Columns.Add("REC_LW", GetType(System.Int64))
        End With

        ASCMAIN1.sql = "Select ICTSTAT2.ITEM_CODE" & vbCrLf _
            & ", SUM (ICTSTAT2.WHSE_QTY_OPEN) WHSE_QTY_OPEN from ICTSTAT2 " & vbCrLf _
            & " where ICTSTAT2.ITEM_CODE in " & vbCrLf _
            & " (Select Distinct ITEM_CODE from " & SOTSORD1 & ")" & vbCrLf _
            & " group by ICTSTAT2.ITEM_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTSTATX", 2))
        With dst.Tables("ICTSTATX")
            .Columns("WHSE_QTY_OPEN").ReadOnly = False
        End With

        ' Put This Order's Pick Qty & Amt into Open

        Create_Relation("SOTSORD2", "SOTSORD1", "ITEM_CODE")
        dst.Tables("SOTSORD2").Columns.Add("ORDR_QTY_PICK_X", GetType(System.Int64), "SUM(CHILD(SOTSORD2_SOTSORD1).ORDR_QTY_PICK_X)")

        For Each rowSOTSORD2 As DataRow In dst.Tables("SOTSORD2").Select("ORDR_QTY_PICK_X <> 0")
            rowSOTSORD2.Item("WHSE_QTY_PICK") = Val(rowSOTSORD2.Item("WHSE_QTY_PICK") & "") _
                                              - Val(rowSOTSORD2.Item("ORDR_QTY_PICK_X") & "")
        Next

        For Each rowSOTSORD1 As DataRow In dst.Tables("SOTSORD1").Select("ORDR_QTY_PICK_X <> 0")
            With rowSOTSORD1
                .Item("ORDR_QTY_OPEN_X") = Val(.Item("ORDR_QTY_OPEN_X") & "") _
                                         + Val(.Item("ORDR_QTY_PICK_X") & "")
                .Item("ORDR_AMT_OPEN_X") = Val(.Item("ORDR_AMT_OPEN_X") & "") _
                                         + Val(.Item("ORDR_AMT_PICK_X") & "")
                .Item("ORDR_QTY_OPEN") = Val(.Item("ORDR_QTY_OPEN") & "") _
                                         + Val(.Item("ORDR_QTY_PICK_X") & "")
            End With
        Next

        Create_Relation("ICTSTATX", "SOTSORD1", "ITEM_CODE")
        dst.Tables("ICTSTATX").Columns.Add("ORDR_QTY_PICK_X", GetType(System.Int64), "SUM(CHILD(ICTSTATX_SOTSORD1).ORDR_QTY_PICK_X)")

        For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select("ORDR_QTY_PICK_X <> 0")
            rowICTSTATX.Item("WHSE_QTY_OPEN") = Val(rowICTSTATX.Item("WHSE_QTY_OPEN") & "") _
                                              - Val(rowICTSTATX.Item("ORDR_QTY_PICK_X") & "")
        Next

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""

        Dim z As String = ""
        If Format(REPORT_DATE0, "yyyyMMdd") = Format(REPORT_DATE1, "yyyyMMdd") Then
            z = " on " & Format(REPORT_DATE0, "MM/dd/yyyy")
        Else
            z = " between " & Format(REPORT_DATE0, "MM/dd/yyyy") & " and " & Format(REPORT_DATE1, "MM/dd/yyyy")
        End If
        If optSR.Value = "R" Then
            SUBT &= "Orders Received" & z
        Else
            SUBT &= "Orders to Ship" & z
        End If

        If Absx1.chkFor("CHKEDI_ONLY").Checked Then
            SUBT &= " (EDI Orders Only)"
        End If

        CR_params.Add("NO_PRICING", IIf(Absx1.chkFor("CHKNO_PRICING").Checked, "1", "0"))
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub
End Class