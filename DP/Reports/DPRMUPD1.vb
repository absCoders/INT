Public Class DPRMUPD1

    Dim rowICTITEM1 As DataRow
    Dim GDATE As Date
    Dim DP_PARM_DEF_PLAN_WHSE As String
    Dim DATE_REQUIRED_min As Date
    Dim XDT_1 As String
    Dim XDT_2 As String
    Dim p24() As String   ' Table of 24 periods YYYYPP, 1 based, 0 = current period
    Dim m24() As String   ' Table of 24 months YYYYMM, 1 based, 0 = current month
    Dim d24() As Date   ' Table of 24 period-ending dates, 1 based, 0 = current period
    Dim f24() As String   ' Table of 24 period-ending dates, formatted as yyyyMM99, 1 based, 0 = current period
    Dim e24() As String   ' Table of 24 period-ending dates, formatted as yyyyMMdd, 1 based, 0 = current period

    Dim ICTPINVX As String
    Dim POTORDRX As String
    Dim sqlICTPINVX As String
    Dim use_binding As Boolean = True

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        '        ASCMAIN1.Progress("Building Work File")

        ' Note: some of the following sections update Oracle before pulling
        '       data into an MDB work table.  This is non-standard coding,
        '       since it is not encapsulated by Begin/Commit.  It has been
        '       coded this way to relieve the hang time experience by others
        '       when running MRP Update mid-day.

        Fix_MIN_MAX_MDS_for_NC()

        Load_Memory()        ' Periods, Codes, Parameterized Dynasets
        Load_MRP_Parms()     ' Load into memory & set GDATE in MRTPARM1
        Load_Plan_Waste()    ' Load ICTPLAN1 waste into ICTITEM1 in Oracle
        Load_DataTables()           ' Load data into dst
        Load_Max_Level()     ' Presently supports single level and Multi-level - might need to support # of levels, some day
        Load_Status()        ' Loads OH Status & Shipments by WHSE_CODE
        Load_Forecast()      ' Loads Forecasted Demand, by Item & Market, into Supply Demand table DPTMUPD0
        Load_PO()            ' Loads Open Purchase Orders
        Load_Plans()         ' For Later, when we support Planner Entered Plans
        Load_Prod_Comm()     ' should load DTL from POTORDR9 for Planner Entered Plans only - NEEDS WORK

        'Dim DV As DataView = New DataView(dst.Tables("DPTMUPD0"), "ITEM_CODE = 'CC001A03'", "", DataViewRowState.CurrentRows)
        'Dim DT As DataTable = DV.ToTable

        BeginTrans()
        Regen()              ' Performs MRP Calcs, and Updates Oracle: ICTSTAT2,POTORDR2,POTORDR3
        ' the A array now contains PCOMM from Load_Prod_Comm in order to resolve the Excess/Obsolete Demand issue
        ' in the future, we may have a PCOMM generated within Regen_1 that is more accurate than the PCOMM loaded into the A array by Load_Prod_Comm
        ' at that time, we should look at when the A array is loaded into DPTMRPG1 - and delay that loading until after all Regen is done
        ' and at that time, we will once again remove the update to DPTMUPD0/POTORDR9 from Load_Prod_Comm and restore those updates in Regen_1

        'BeginTrans()
        Update_Server()
        CommitTrans()

        ' no planner plans yet:
        '   retract plans would not have wiped out MRTPLAN1 in its entirety
        '   recalc compsdue would have worked on MRTPLAN1 also
    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""

        For i As Integer = 0 To 2
            RPT_TITLE = New String() {"Critical Data Missing Report",
                                      "Planning Exceptions Report",
                                      "Planning Action Report"}(i)

            RPT = "DPRMUPD" & Format$(i + 1, "0")
            SUBT = ""
            If i = 1 Then
                SUBT = "By Item"
                CR_params.Add("SORTBY", "I")
                For Each rowDPTMUPD2 As DataRow In dst.Tables("DPTMUPD2").Select("")
                    rowDPTMUPD2.Item("SORT1") = rowDPTMUPD2.Item("ITEM_CODE")
                    rowDPTMUPD2.Item("SORT2") = rowDPTMUPD2.Item("EXC_MSG_CODE")
                Next
            End If

            Generate_Report(RPT, RPT_TITLE, SUBT)

            If i = 1 Then
                SUBT = "By Exception"
                CR_params.Add("SORTBY", "E")
                For Each rowDPTMUPD2 As DataRow In dst.Tables("DPTMUPD2").Select("")
                    rowDPTMUPD2.Item("SORT1") = rowDPTMUPD2.Item("EXC_MSG_CODE")
                    rowDPTMUPD2.Item("SORT2") = rowDPTMUPD2.Item("ITEM_CODE")
                Next
                Generate_Report(RPT, RPT_TITLE, SUBT)
            End If
        Next i
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        ASCMAIN1.Progress("Report Calculations")
        'dst.Tables("ASTSRPT1").Columns("QTY_OPEN_TOT").Expression = "ISNULL(QTY_OPEN_F0,0)+ISNULL(QTY_OPEN_F1,0)+ISNULL(QTY_OPEN_F2,0)+ISNULL(QTY_OPEN_F3,0)+ISNULL(QTY_OPEN_F4,0)"
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
        End If
    End Sub

    Sub Load_Memory()
        ReDim p24(25) ' Periods
        ReDim m24(25) ' Months
        ReDim d24(25) ' Period Ending Dates
        ReDim e24(25) ' same as d24
        ReDim f24(25) ' same as d24 but with 99 as the date

        For i As Integer = 0 To 24
            Dim j As Integer = i + 1
            p24(j) = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, i)
            m24(j) = ASCMAIN1.Get_YYYYMM(p24(j))
            Dim rowGLTPAR2 As DataRow = LookUp("GLTPARM2", p24(j))
            d24(j) = rowGLTPAR2.Item("PRD_END_DATE")
            f24(j) = Format(d24(j), "yyyyMMdd")
            Mid(f24(j), 7, 2) = "99"
            e24(j) = Format(d24(j), "yyyyMMdd")
        Next
    End Sub

    Sub Update_Server()

        ASCMAIN1.Progress("Now Updating Server", "")

        'ASCMAIN1.Progress("-", "Retract Plans")
        'ASCMAIN1.sql = "Delete from MRTPLAN1"
        'ASCDATA1.ExecuteSQL()

        'ASCMAIN1.Progress("-", "Retract Prod Comms")
        'ASCMAIN1.sql = "Delete from POTORDR9"
        'ASCDATA1.ExecuteSQL()

        Update_Record_TDA("POTORDR2")

        ASCMAIN1.Progress("-", "Clear Status")
        ASCMAIN1.sql = "Update ICTSTAT2 set WHSE_QTY_COMM = 0, WHSE_QTY_PLAN = 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from ICTSTAT2" & vbCrLf _
            & " where NVL(WHSE_QTY_ON_HAND,0) = 0" & vbCrLf _
            & "   and NVL(WHSE_QTY_ONPO,0) = 0" & vbCrLf _
            & "   and NVL(WHSE_QTY_PLAN,0) = 0" & vbCrLf _
            & "   and NVL(WHSE_QTY_OPEN,0) = 0" & vbCrLf _
            & "   and NVL(WHSE_QTY_PICK,0) = 0" & vbCrLf _
            & "   and NVL(WHSE_QTY_COMM,0) = 0" & vbCrLf _
            & "   and NVL(WHSE_QTY_HOLD,0) = 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Update Plans")
        Update_Record_TDA("DPTPLAN1")

        'ASCMAIN1.Progress("-", "Update Prod Comms")
        'Update_Record_TDA("POTORDR9")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare" & vbCrLf _
            & "  Cursor C1 is " & vbCrLf _
            & "   Select 'O', POTORDR9.ITEM_CODE, " & vbCrLf _
            & "          POTORDR1.VEND_WHSE_CODE WHSE_CODE," & vbCrLf _
            & "          SUM (POTORDR9.PO_QTY_COM) QTY " & vbCrLf _
            & "    from POTORDR1,POTORDR9" & vbCrLf _
            & "    where POTORDR1.PO_ORDER_NO = POTORDR9.PO_ORDER_NO " & vbCrLf _
            & "      and POTORDR9.PO_ORDER_LNO <> 0 " & vbCrLf _
            & "    group by POTORDR9.ITEM_CODE, " & vbCrLf _
            & "             POTORDR1.VEND_WHSE_CODE" & vbCrLf _
            & "    Union" & vbCrLf _
            & "   Select 'P', POTORDR9.ITEM_CODE, " & vbCrLf _
            & "          DPTPLAN1.AT_WHSE WHSE_CODE," & vbCrLf _
            & "          SUM (POTORDR9.PO_QTY_COM) QTY " & vbCrLf _
            & "    from DPTPLAN1,POTORDR9" & vbCrLf _
            & "    where DPTPLAN1.PLAN_NO = POTORDR9.PO_ORDER_NO " & vbCrLf _
            & "      and POTORDR9.PO_ORDER_LNO = 0 " & vbCrLf _
            & "    group by POTORDR9.ITEM_CODE, " & vbCrLf _
            & "             DPTPLAN1.AT_WHSE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT2 Set WHSE_QTY_COMM = NVL(WHSE_QTY_COMM,0) + NVL(R1.QTY,0)" & vbCrLf _
            & "    where ITEM_CODE = R1.ITEM_CODE AND WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then" & vbCrLf _
            & "    Insert into ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_COMM)" & vbCrLf _
            & "     Values (R1.ITEM_CODE, R1.WHSE_CODE, R1.QTY); " & vbCrLf _
            & "   End If;" & vbCrLf _
            & "  End Loop; " & vbCrLf _
            & " End; " & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Exception Queues")

        ASCDATA1.DeleteRows(dst.Tables("DPTMUPD1"), "CRDM_SUPPRESS = '1'")
        ASCDATA1.DeleteRows(dst.Tables("DPTMUPD2"), "EXC_MSG_SUPPRESS = '1'")

        For Each rowDPTEXCS1 As DataRow In dst.Tables("DPTEXCS1").Select("EXC_SUPP_UNTIL >= '" & Format(GDATE, "MM/ddyyyy") & "'")
            Dim EXC_MSG_CODE As String = rowDPTEXCS1.Item("EXC_MSG_CODE")
            Dim ITEM_CODE As String = rowDPTEXCS1.Item("ITEM_CODE")
            Dim rowDPTMUPD2 As DataRow = dst.Tables("DPTMUPD2").Rows.Find(New String() {ITEM_CODE, EXC_MSG_CODE})
            rowDPTMUPD2.Delete()
            rowDPTEXCS1.Item("EXC_SUPP_USED") = GDATE
        Next
        Update_Record_TDA("DPTEXCS1")

        ASCMAIN1.Progress("-", "MRP Reports")
        For i As Integer = 0 To 3
            Update_Record_TDA("DPTMUPD" & Format(i, "0"), "1=1")
        Next i

        ASCMAIN1.Progress("-", "MRP Grid")
        Dim DT As DateTime = Now

        If use_binding Then

            ASCMAIN1.sql = "Delete from DPTMRPG1"
            ASCDATA1.ExecuteSQL()

            'Dim r As Int32 = 0
            'Dim tbl As DataTable = dst.Tables("DPTMRPG1").Copy
            'dst.Tables("DPTMRPG1").Rows.Clear()
            'For Each rowDPTMRPG1 As DataRow In tbl.Select("")
            '    dst.Tables("DPTMRPG1").Rows.Add(rowDPTMRPG1.ItemArray)
            '    r += 1
            '    If r Mod 1000 = 0 Then
            '        Update_BAs("DPTMRPG1")
            '        dst.Tables("DPTMRPG1").Rows.Clear()
            '    End If
            'Next
            'If r Mod 1000 <> 0 Then ' THERE WERE ROWS LEFT
            '    Update_BAs("DPTMRPG1")
            'End If
            'tbl = Nothing

            Update_BAs("DPTMRPG1") ' ALL ROWS In 1 GO
        Else
            Update_Record_TDA("DPTMRPG1", "1=1")
        End If

        'Update_Record_TDA("DPTMRPG1", "1=1")
        Debug.Print("MRP Grid Update took " & CStr(Now.Subtract(DT).TotalSeconds) & " seconds")
    End Sub

#Region "VB6"

    Sub Regen()
        ASCMAIN1.Progress("Performing MRP Update", "")

        ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
            & ", POTORDR1.VEND_WHSE_CODE, POTORDR1.VEND_CODE, POTORDR1.PO_ORDER_TYPE " & vbCrLf _
            & " from POTORDR2, POTORDR1" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_QTY_OPN <> 0" & vbCrLf _
            & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
            & "   and POTORDR2.BM_ISSUE_SEL is Not Null" & vbCrLf _
            & "   and POTORDR2.ITEM_CODE = :PARM1"

        ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
            & ", POTORDR1.VEND_WHSE_CODE, POTORDR1.VEND_CODE, POTORDR1.PO_ORDER_TYPE " & vbCrLf _
            & " from POTORDR2, POTORDR1" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_QTY_OPN <> 0" & vbCrLf _
            & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
            & "   and (POTORDR2.BM_ISSUE_SEL is Not Null OR POTORDR2.BM_ISSUE_NO is Not Null)"
        Dim sqlPOTORDR2 As String = ASCMAIN1.sql
        Create_TDA(dst.Tables.Add, "POTORDR2", "**", 0, True, "PO_DATE_COMPSDUE", 2, "")
        Fill_Records("POTORDR2")
        ' THE DST ABOVE IS ONLY FOR MAKE POS - THERE ARE JUST A HANDFUL, YET WE FILL_RECORDS ON THIS TABLE 8K TIMES
        ' AND WHEN WE UPDATE POTORDR2, WE ARE UPDATING ALL COLS WHETHER THEY NEED IT OR NOT

        Create_TDA(dst.Tables.Add, "POTORDR1", "*")
        Create_TDA(dst.Tables.Add, "POTORDR9", "*")

        With dst.Tables.Add("DPTCOMMX")
            .Columns.Add("PO_ORDER_NO", GetType(System.String))
            .Columns.Add("PO_ORDER_LNO", GetType(System.Int32))
            .Columns.Add("ITEM_CODE", GetType(System.String))
            .Columns.Add("P_INDEX", GetType(System.Int32))
            .Columns.Add("VEND_WHSE_CODE", GetType(System.String))
            .Columns.Add("DATE_COMPSDUE", GetType(System.DateTime))
            '  .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE"), .Columns("DATE_REQ"), .Columns("WHSE_CODE")}
        End With

        ' THERE ARE OVER 8K ITEMS IN DPTMRPD0 - WE NEED TO BE ABLE TO ELIMINATE ITEMS WITH NO SUPPLY OR DEMAND
        Dim rows() As DataRow = dst.Tables("DPTMUPD0").Select("", "ITEM_CODE,DATE_REQ,WHSE_CODE")
        Dim tblDPTMUPD0d As DataTable = ASCDATA1.SelectDistinct(rows, New String() {"ITEM_CODE", "DATE_REQ", "WHSE_CODE"})

        With dst.Tables.Add("DPTMUPDS")
            .Columns.Add("ITEM_CODE", GetType(System.String))
            .Columns.Add("DATE_REQ", GetType(System.DateTime))
            .Columns.Add("WHSE_CODE", GetType(System.String))
            .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE"), .Columns("DATE_REQ"), .Columns("WHSE_CODE")}
        End With

        dst.Tables("DPTMUPDS").Merge(tblDPTMUPD0d)
        Create_Relation("DPTMUPDS", "DPTMUPD0", "ITEM_CODE,DATE_REQ,WHSE_CODE")

        With dst.Tables("DPTMUPDS")
            .Columns.Add("QTY_OH", GetType(System.Int64), "SUM(CHILD.QTY_OH)")
            .Columns.Add("QTY_ORD", GetType(System.Int64), "SUM(CHILD.QTY_ORD)")
            .Columns.Add("QTY_PLN", GetType(System.Int64), "SUM(CHILD.QTY_PLN)")
            .Columns.Add("QTY_REQ", GetType(System.Int64), "SUM(CHILD.QTY_REQ)")
            .Columns.Add("QTY_NET", GetType(System.Int64), "ISNULL(QTY_OH,0) + ISNULL(QTY_ORD,0) + ISNULL(QTY_PLN,0) - ISNULL(QTY_REQ,0)")
        End With

        ASCMAIN1.Progress("-", "Retract Plans")
        ASCMAIN1.sql = "Delete from DPTPLAN1"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("-", "Retract Prod Comms")
        ASCMAIN1.sql = "Delete from POTORDR9"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ICTSTAT2 set WHSE_QTY_COMM = 0, WHSE_QTY_PLAN = 0"
        ASCDATA1.ExecuteSQL()

        For lvl As Integer = 1 To 0 Step -1
            Dim sqlw As String = "EXCLUDE_FROM_MRP <> '1' and "
            If lvl = 1 Then
                sqlw &= "MAX_LEVEL > 0"
            Else
                sqlw &= "MAX_LEVEL = 0"
            End If
            For Each Me.rowICTITEM1 In dst.Tables("ICTITEM1").Select(sqlw, "MAX_LEVEL DESC, ITEM_CODE")
                Dim ITEM_CODE As String = rowICTITEM1.Item("ITEM_CODE")
                If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" And Format(Now.Date, "yyyyMMdd") = "20250618" Then
                    If Not New String() {"KS001A03"}.Contains(ITEM_CODE) Then
                        Continue For
                    End If
                End If
                Regen_1()
            Next
            ' IF LVL = 1 THEN   Call Rebuild_Commits
        Next
    End Sub

    Sub Regen_1()

        Dim minpbase As Integer = Val(Mid(ASCMAIN1.CYM, 1, 4)) * 12 + Val(Mid(ASCMAIN1.CYM, 5, 2))

        Dim ITEM_CODE As String = rowICTITEM1.Item("ITEM_CODE")
        Dim MINPOS As Decimal = Val(rowICTITEM1.Item("ITEM_POS_MIN") & "")
        Dim graceused As String = "N"
        ASCMAIN1.Progress("-", ITEM_CODE)


        Dim ITEM_BUFFER_QTY As Int32 = Val(rowICTITEM1.Item("ITEM_BUFFER_QTY") & "")
        Dim ITEM_BUFFER_PCT As Decimal = Val(rowICTITEM1.Item("ITEM_BUFFER_PCT") & "")

        'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "KS001A03" Then Stop
        'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CC001A01" Then Stop

        ' THIS WAS A GENERIC ACTION DATE THAT COULD BE SPECIFIED IN THE ITEM MASTER
        '    If vtos(tblICWITEM1.Fields("ITEM_DATE_MRP_ACTION").Value) <> "" Then
        '        If Format$(tblICWITEM1.Fields("ITEM_DATE_MRP_ACTION").Value,"yyyymmdd") <= Format$(GDATE,"yyyymmdd") Then
        '            msg = "00"
        '            suf = ""
        '            If Format$(tblICWITEM1.Fields("ITEM_DATE_MRP_ACTION").Value,"yyyymmdd") < Format$(GDATE,"yyyymmdd") Then
        '                suf = "*"
        '            End If
        '            Call Act_Msg
        '        End If
        '    End If
        Dim ITEM_ABC_CODE As String = rowICTITEM1.Item("ITEM_ABC_CODE") & ""
        Dim ITEM_PLAN_MAKE_BUY As String = rowICTITEM1.Item("ITEM_PLAN_MAKE_BUY") & ""
        Dim VEND_WHSE_CODE As String = ""
        Dim VEND_CODE As String = rowICTITEM1.Item("VEND_CODE") & ""
        Dim rowAPTVEND1 As DataRow = dst.Tables("APTVEND1").Rows.Find(VEND_CODE)
        If rowAPTVEND1 Is Nothing Then
            rowAPTVEND1 = dst.Tables("APTVEND1").Rows.Add(New String() {VEND_CODE})
            If VEND_CODE <> "" Then
                Dim row As DataRow = LookUp("APTVEND1", VEND_CODE)
                rowAPTVEND1.ItemArray = row.ItemArray
            End If
        End If
        VEND_WHSE_CODE = rowAPTVEND1.Item("VEND_WHSE_CODE") & ""
        Dim rowDPTVNDI1 As DataRow = dst.Tables("DPTVNDI1").Rows.Find(New String() {VEND_CODE, ITEM_CODE})
        If rowDPTVNDI1 Is Nothing Then
            rowDPTVNDI1 = dst.Tables("DPTVNDI1").Rows.Add(New String() {VEND_CODE, ITEM_CODE})
        End If
        Dim ITEM_MRP_PLANR_CODE As String = rowICTITEM1.Item("ITEM_MRP_PLANR_CODE") & ""
        Dim rowDPTABCP1 As DataRow = dst.Tables("DPTABCP1").Rows.Find(ITEM_ABC_CODE)

        Dim ITEM_DAYS_SUPPLY_MIN As Int64 = 0
        If rowDPTABCP1 IsNot Nothing Then
            ITEM_DAYS_SUPPLY_MIN = Val(rowDPTABCP1.Item("ABC_MIN_DAYS_SUPPLY") & "")
        End If
        Dim ITEM_POS_MAX As Decimal = Val(rowICTITEM1.Item("ITEM_POS_MAX") & "")
        Dim ITEM_POS_MIN As Decimal = Val(rowICTITEM1.Item("ITEM_POS_MIN") & "")
        Dim ITEM_DAYS_MASCERATE As Integer = 0 ' Val(rowICTITEM1.Item("ITEM_DAYS_MASCERATE") & "")

        Dim ITEM_FC_DAY_REQ As Integer = Val(rowICTITEM1.Item("ITEM_FC_DAY_REQ") & "")
        If ITEM_FC_DAY_REQ = 0 Then
            ITEM_FC_DAY_REQ = Val(ROWs("DPTPARM1").Item("DP_PARM_FC_REQ_DAY") & "")
        End If

        Dim PO_MULTIPLE As Int64 = 0
        If Val(rowDPTVNDI1.Item("PO_MULTIPLE") & "") <> 0 Then
            PO_MULTIPLE = Val(rowDPTVNDI1.Item("PO_MULTIPLE") & "")
        Else
            PO_MULTIPLE = Val(rowICTITEM1.Item("ITEM_PO_QTY_MULT") & "")
        End If

        If PO_MULTIPLE = 0 Then
            PO_MULTIPLE = Val(rowICTITEM1.Item("CARTON_PACK_QTY") & "")
        End If

        Dim PO_MIN_QTY As Int64
        If Val(rowDPTVNDI1.Item("PO_MIN_QTY") & "") <> 0 Then
            PO_MIN_QTY = Val(rowDPTVNDI1.Item("PO_MIN_QTY") & "")
        Else
            PO_MIN_QTY = Val(rowICTITEM1.Item("ITEM_PO_QTY_MIN") & "")
        End If

        Dim PO_LEAD_TIME As Int64
        If Val(rowDPTVNDI1.Item("PO_LEAD_TIME") & "") <> 0 Then
            PO_LEAD_TIME = Val(rowDPTVNDI1.Item("PO_LEAD_TIME") & "")
        Else
            If Val(rowICTITEM1.Item("ITEM_LEAD_TIME_DAYS") & "") <> 0 Then
                ' INTRODUCED ICTITEM1.ITEM_LEAD_TIME_DAYS AND REVERSED ORDER ITEM VS VENDOR
                PO_LEAD_TIME = Val(rowICTITEM1.Item("ITEM_LEAD_TIME_DAYS") & "")
            Else
                PO_LEAD_TIME = Val(rowAPTVEND1.Item("PO_LEAD_TIME") & "")
            End If
            'If Val(rowAPTVEND1.Item("PO_LEAD_TIME") & "") <> 0 Then
            '    PO_LEAD_TIME = Val(rowAPTVEND1.Item("PO_LEAD_TIME") & "")
            'Else
            '    'PO_LEAD_TIME = 0 ' Val(rowICTITEM1.Item("ITEM_DAYS_PO_LT") & "")
            'End If
        End If

        Dim PO_DRP As Int64 = Val(rowDPTVNDI1.Item("PO_DRP") & "")
        Dim PO_SCH_DAYS As Int64 = Get_PO_SCH_DAYS(ITEM_PLAN_MAKE_BUY, rowDPTVNDI1, rowAPTVEND1)

        Dim PO_XIT_DAYS As Int64
        If Val(rowDPTVNDI1.Item("PO_XIT_DAYS") & "") <> 0 Then
            PO_XIT_DAYS = Val(rowDPTVNDI1.Item("PO_XIT_DAYS") & "")
        Else
            If Val(rowAPTVEND1.Item("PO_XIT_DAYS") & "") <> 0 Then
                PO_XIT_DAYS = Val(rowAPTVEND1.Item("PO_XIT_DAYS") & "")
            Else
                PO_XIT_DAYS = Val(ROWs("DPTPARM1").Item("DP_PARM_DEF_PO_XIT_DAYS") & "")
            End If
        End If
        'Dim mfg As Long

        Dim CRITICAL_LEAD_TIME_days As Integer = PO_LEAD_TIME + PO_SCH_DAYS + PO_XIT_DAYS ' + mfg

        Dim CRITICAL_LEAD_TIME_Pindex As Integer = 0
        '    CRITICAL_LEAD_TIME_date = Format$(DateAdd("d", CRITICAL_LEAD_TIME_days, GDATE), "yyyymmdd")
        Dim CRITICAL_LEAD_TIME_date As String = Format(GDATE.AddDays(CRITICAL_LEAD_TIME_days), "yyyyMMdd")

        '  Stop ' want 25 if there is nothing in m24 greater than critical_lead_time_date
        CRITICAL_LEAD_TIME_Pindex = First_Date_Index(m24, CRITICAL_LEAD_TIME_date)

        If CRITICAL_LEAD_TIME_Pindex <> 0 Then
            CRITICAL_LEAD_TIME_Pindex = 1 + (CRITICAL_LEAD_TIME_Pindex - 1) / 8
        End If

        Dim bmissue As String = Format(Val(rowICTITEM1.Item("BM_ISSUE_COUNTER") & ""), "00")

        Dim ITEM_SNU_CODE As String = rowICTITEM1.Item("ITEM_SNU_CODE") & ""
        Dim ITEM_STATUS As String = rowICTITEM1.Item("ITEM_STATUS") & ""

        ' Check Critical Data Missing

        Dim CRDM_CODEs As New List(Of String)

        If VEND_CODE = "" Then CRDM_CODEs.Add("01")

        If InStr("MB", ITEM_PLAN_MAKE_BUY) = 0 Then
            CRDM_CODEs.Add("08")
        Else
            If ITEM_PLAN_MAKE_BUY = "M" Then
                If bmissue = "" Then CRDM_CODEs.Add("02")
                If PO_SCH_DAYS = 0 Then CRDM_CODEs.Add("12")
                If PO_DRP = 0 Then CRDM_CODEs.Add("04")
                If VEND_WHSE_CODE = "" Then CRDM_CODEs.Add("10")
            End If
        End If

        ' If ASCMAIN1.Running_in_VS And ITEM_CODE = "VA010P85" Then Stop

        If PO_LEAD_TIME = 0 Then CRDM_CODEs.Add("13")
        If PO_XIT_DAYS = 0 Then CRDM_CODEs.Add("11")
        If ITEM_SNU_CODE <> "U" Then
            If ITEM_POS_MAX = 0 Or ITEM_POS_MIN = 0 Then
                CRDM_CODEs.Add("06")
            End If
        End If
        If PO_MIN_QTY = 0 Then CRDM_CODEs.Add("03")

        If CRDM_CODEs.Count <> 0 Then
            For Each CRDM_CODE In CRDM_CODEs
                dst.Tables("DPTMUPD1").Rows.Add(New Object() {ITEM_CODE, CRDM_CODE, ITEM_STATUS})
            Next
            Exc_Msg(ITEM_CODE, "00", "")
        End If




        ' Cycle thru all Supply & Demand Records

        Dim dti As Integer = 0
        Dim dtimax As Integer = 0
        Dim dt() As String = Nothing
        Dim minp(25, 1) As Int64
        Dim dtbal() As Int64 = Nothing
        Dim dtbalw(,) As Int64 = Nothing
        ReDim dtbalw(20, 0)
        Dim wi As New Dictionary(Of String, Integer)

        Dim ym0i As Integer = 1

        'Dim dV As DataView = New DataView(dst.Tables("DPTMUPDS"), "ITEM_CODE = '" & ITEM_CODE & "'", "", DataViewRowState.CurrentRows)
        'Dim DTX As DataTable = dV.ToTable

        For Each rowDPTMUPDS As DataRow In dst.Tables("DPTMUPDS").Select("ITEM_CODE = '" & ITEM_CODE & "'", "DATE_REQ, WHSE_CODE")
            Dim WHSE_CODE As String = rowDPTMUPDS.Item("WHSE_CODE")

            If Not wi.ContainsKey(WHSE_CODE) Then wi.Add(WHSE_CODE, wi.Count + 1)

            Dim DATE_REQ As String = Format(rowDPTMUPDS.Item("DATE_REQ"), "yyyyMMdd")

            Do While DATE_REQ > f24(ym0i) And ym0i <= 24
                If ym0i >= f24.Length - 1 Then
                    Exit Do
                Else
                    Regen_Append_Event_Date(f24(ym0i), dti, dt, dtbal, dtbalw, dtimax)
                    ym0i = ym0i + 1
                End If
            Loop

            Dim QTY_REQ As Int64 = Val(rowDPTMUPDS.Item("QTY_REQ") & "")
            Dim minpi As Integer = 0
            If QTY_REQ <> 0 Then
                minpi = Val(Mid$(DATE_REQ, 1, 4)) * 12 + Val(Mid$(DATE_REQ, 5, 2)) - minpbase + 1
                If minpi < 0 Then
                    minpi = 0
                End If
                If minpi > 25 Then
                    minpi = 25
                End If
                minp(minpi, 0) = minp(minpi, 0) + QTY_REQ
            End If
            Dim QTY_NET As Int64 = Val(rowDPTMUPDS.Item("QTY_NET") & "") ' NET SUPPLY - DEMAND FOR THE PERIOD
            If DATE_REQ <> f24(ym0i) Then ' AND THIS WOULD NEVER BE, SINCE f24 ends in 99
                Regen_Append_Event_Date(DATE_REQ, dti, dt, dtbal, dtbalw, dtimax)
            End If
            dtbal(dti) += QTY_NET
            dtbalw(wi(WHSE_CODE), dti) += QTY_NET
        Next

        '     If ASCMAIN1.Running_in_VS And ITEM_CODE = "CH017A02" Then Stop

        Dim iMINPOS As Decimal = CInt(MINPOS + 0.99)
        Dim maxii = 24 - System.Math.Floor(MINPOS) + 1
        If maxii > 24 Then
            maxii = 24
        End If

        'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "KS001A03" Then Stop

        For ii = 0 To maxii ' 24 - System.Math.Floor(MINPOS)   ' iminpos

            'If ITEM_BUFFER_QTY > 0 Then
            '    If QTY_REQ > 0 Then
            '        QTY_REQ += ITEM_BUFFER_QTY
            '    End If
            'ElseIf ITEM_BUFFER_PCT > 0 Then
            '    QTY_REQ = System.Math.Round(QTY_REQ * (100 + ITEM_BUFFER_PCT) / 100, 0)
            'End If

            '    If ii = 19 Then Stop
            minp(ii, 1) = 0 ' minp(ii, 0)
            If iMINPOS > 0 Then
                For j = 1 To iMINPOS
                    If ii + j <= 25 Then
                        Dim QR As Int64 = minp(ii + j, 0)



                        If ITEM_BUFFER_QTY <> 0 Or ITEM_BUFFER_PCT <> 0 Then
                            Dim addl_demand_from_buffer As Integer = 0
                            If ITEM_BUFFER_QTY > 0 Then
                                addl_demand_from_buffer = ITEM_BUFFER_QTY
                            ElseIf ITEM_BUFFER_PCT > 0 Then
                                addl_demand_from_buffer = System.Math.Round(QR * (ITEM_BUFFER_PCT) / 100, 0)
                            End If
                            QR += addl_demand_from_buffer
                        End If




                        If j = iMINPOS Then
                            QR = QR * (1 - (iMINPOS - MINPOS))
                        End If
                        minp(ii, 1) = minp(ii, 1) + QR
                    End If
                Next j
            End If
        Next ii

        '  If ASCMAIN1.Running_in_VS And ITEM_CODE = "CH010C07USA" Then Stop

        If ITEM_SNU_CODE = "U" Or ROWs("DPTPARM1").Item("DP_PARM_MR_PLAN_FG") & "" = "1" Then
            dti = 0
            Dim j As Integer = First_Date_Index(dt, Format(DATE_REQUIRED_min, "yyyyMMdd"))

            ' Stop ' ubound (dt) may be an empty date.


            'Dim dV As DataView = New DataView(dst.Tables("DPTMUPDS"), "ITEM_CODE = '" & ITEM_CODE & "'", "", DataViewRowState.CurrentRows)
            'Dim DTX As DataTable = dV.ToTable


            Do While dt IsNot Nothing AndAlso dti < UBound(dt)
                dti = dti + 1
                '            If dtbal(dti) < 0 Then
                Dim minpz As String = dt(dti)
                Dim minpi As Integer = Val(Mid(minpz, 1, 4)) * 12 + Val(Mid(minpz, 5, 2)) - minpbase + 1
                If minpi < 0 Then
                    minpi = 0
                Else
                    If minpi > 24 Then
                        minpi = 25
                    End If
                End If

                'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CC004P80" Then Stop

                Dim dtz As String = dt(dti)
                If Mid(dtz, 7, 2) = "99" Then
                    Mid(dtz, 7, 2) = "01"
                    Dim dtz_date As Date = DateValue(Mid(dtz, 5, 2) & "/" & Mid(dtz, 7, 2) & "/" & Mid(dtz, 1, 4))
                    dtz = Format(dtz_date.AddMonths(1).AddDays(-1), "yyyyMMdd")
                End If

                If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "KS001A03 NOT" Then
                    Debug.Print(dtz & " EOM=" & dtbal(dti) & ":" & "Future Demand = " & minp(minpi, 1))
                    Stop
                End If

                ' If ASCMAIN1.Running_in_VS And ITEM_CODE = "KS001A03" Then Stop
                If (Mid(minpz, 7, 2) <> "99" And dtbal(dti) < 0) _
                Or (Mid(minpz, 7, 2) = "99" And dtbal(dti) < minp(minpi, 1)) Then
                    'If ASCMAIN1.Running_in_VS And ITEM_CODE = "KS001A03" Then Stop
                    Dim z As String = "N"
                    If dti < dtimax Then
                        ' this next line looks only at the next event date, and if within the grace days, instead of looking at the final balance after grace days, in order to determine if the grace kicks in
                        If ((Mid(minpz, 7, 2) <> "99" And dtbal(dti + 1) >= 0) _
                        Or (Mid(minpz, 7, 2) = "99" And dtbal(dti + 1) >= minp(minpi, 1))) _
                            And CDate(FND(dtz, 4)).AddDays(Val(ROWs("DPTPARM1").Item("DP_PARM_PLAN_GRACE_DAYS") & "")) >= DateValue(FND(dtz, 4)) Then
                            ' I think that the CDate line above needs to compare to a date in the next month, but I am not sure
                            'If dtbal(i + 1) > 0 _
                            'And DateDiffw(DateValue(FND(dtz, 4)), MR_PARM_PLAN_GRACE_DAYS) >= DateValue(FND(dtz, 4)) Then
                            z = "Y"
                            graceused = "Y"
                        End If
                    End If
                    If z = "N" Then
                        Dim DTREQ As Date = DateValue(FND(dtz, 4))

                        '                    q = -1 * dtbal(dti)
                        Dim QTY_PLANNED As Int64 = minp(minpi, 1) - dtbal(dti)

                        If ITEM_DAYS_SUPPLY_MIN <> 0 Then
                            z = Format(DTREQ.AddDays(ITEM_DAYS_SUPPLY_MIN), "yyyyMMdd")
                            j = First_Date_Index(dt, z)
                            If -1 * dtbal(j) > QTY_PLANNED Then
                                QTY_PLANNED = -1 * dtbal(j)
                            End If
                        End If

                        'If ITEM_BUFFER_QTY <> 0 Or ITEM_BUFFER_PCT <> 0 Then
                        '    Dim FC_future As Int64 = 0
                        '    If minpi <= 24 Then
                        '        Dim pimax As Integer = minpi + MINPOS
                        '        If pimax > 25 Then pimax = 25
                        '        For i As Integer = minpi + 1 To pimax
                        '            FC_future += minp(i, 0)
                        '        Next
                        '    End If

                        '    If FC_future > 0 Then
                        '        Dim addl_demand_from_buffer As Integer = 0
                        '        If ITEM_BUFFER_QTY > 0 Then
                        '            addl_demand_from_buffer = ITEM_BUFFER_QTY
                        '        ElseIf ITEM_BUFFER_PCT > 0 Then
                        '            addl_demand_from_buffer = System.Math.Round(FC_future * (ITEM_BUFFER_PCT) / 100, 0)
                        '        End If
                        '        QTY_PLANNED += addl_demand_from_buffer
                        '    End If
                        'End If


                        If QTY_PLANNED < PO_MIN_QTY Then
                            QTY_PLANNED = PO_MIN_QTY
                        End If
                        If PO_MULTIPLE <> 0 Then
                            If QTY_PLANNED Mod PO_MULTIPLE <> 0 Then
                                QTY_PLANNED = QTY_PLANNED + (PO_MULTIPLE - (QTY_PLANNED Mod PO_MULTIPLE))
                            End If
                        End If

                        Dim bm As String = ""
                        If ITEM_PLAN_MAKE_BUY = "M" Then
                            bm = bmissue
                        End If
                        '                    If minpos <> 0 Then
                        '                        dtreq = DateAdd("m", -1 * minpos, dtreq)
                        '                    End If

                        If Format(DTREQ, "yyyyMMdd") < Format(DATE_REQUIRED_min, "yyyyMMdd") Then
                            DTREQ = DATE_REQUIRED_min
                        End If

                        '   If rowICTITEM1.Item("ITEM_CODE") = "32915465" Then Stop
                        ' If ASCMAIN1.Running_in_VS And ITEM_CODE = "853218006001" Then Stop

                        If rowICTITEM1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") & "" = "4" Then
                        ElseIf rowICTITEM1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") & "" = "1" AndAlso Format(DTREQ, "yyyyMMdd") < Format((Now + ASCMAIN1.NowTSD).AddDays(CRITICAL_LEAD_TIME_days), "yyyyMMdd") Then
                        ElseIf rowICTITEM1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") & "" = "2" AndAlso Format(DTREQ, "yyyyMMdd") < Format(rowICTITEM1.Item("ITEM_PLAN_QUIET_ZONE_DATE"), "yyyyMMdd") Then
                        ElseIf rowICTITEM1.Item("ITEM_PLAN_QUIET_ZONE_TYPE") & "" = "3" AndAlso Format(DTREQ, "yyyyMMdd") < Format((Now + ASCMAIN1.NowTSD).AddDays(Val(rowICTITEM1.Item("ITEM_PLAN_QUIET_ZONE_DAYS") & "")), "yyyyMMdd") Then
                        Else
                            Dim MFG As Int64 = 0 ' THIS SHOULD BE CALCULATED
                            Create_Plan(ITEM_CODE, QTY_PLANNED, DTREQ, VEND_CODE, VEND_WHSE_CODE, PO_LEAD_TIME, PO_SCH_DAYS, MFG, PO_XIT_DAYS, bm)
                            If Not wi.ContainsKey(DP_PARM_DEF_PLAN_WHSE) Then wi.Add(DP_PARM_DEF_PLAN_WHSE, wi.Count + 1)
                            For j = dti To dtimax
                                dtbal(j) += QTY_PLANNED
                                dtbalw(wi(DP_PARM_DEF_PLAN_WHSE), j) += QTY_PLANNED
                            Next j
                        End If
                    End If
                End If
            Loop
        End If


        'Dim dV As DataView = New DataView(dst.Tables("DPTMUPDS"), "ITEM_CODE = '" & ITEM_CODE & "'", "", DataViewRowState.CurrentRows)
        'Dim DTX As DataTable = dV.ToTable
        Dim dV As DataView = New DataView(dst.Tables("DPTMUPD0"), "ITEM_CODE = '" & ITEM_CODE & "'", "", DataViewRowState.CurrentRows)
        Dim DTX As DataTable = dV.ToTable

        dti = 0
        Do While dti < dtimax
            dti = dti + 1
            If dtbal(dti) < 0 Then
                If dt(dti) <= CRITICAL_LEAD_TIME_date Then
                    Exc_Msg(ITEM_CODE, "02", "") ' Neg Avail I/S CLT
                Else
                    Exc_Msg(ITEM_CODE, "07", "") ' Neg Avail O/S CLT
                End If
            End If
        Loop

        If dtimax <> 0 Then
            For dti = 1 To dtimax
                If dt(dti) > XDT_2 Then
                    Exit For
                Else
                    Dim z As String = ""
                    For Each WHSE_CODE In wi.Keys
                        If InStr(z, "+") = 0 Then
                            If dtbalw(wi(WHSE_CODE), dti) > 0 Then
                                z = z & "+"
                            End If
                        End If
                        If InStr(z, "-") = 0 Then
                            If dtbalw(wi(WHSE_CODE), dti) < 0 Then
                                z = z & "-"
                            End If
                        End If
                    Next
                    If z = "+-" Or z = "-+" Then
                        If dt(dti) > XDT_1 Then
                            Act_Msg(ITEM_CODE, "02", "") ' Transfer Order Required
                        Else
                            Exc_Msg(ITEM_CODE, "07", "") ' Transfer Order Past Due
                        End If
                    End If
                End If
            Next dti
        End If

        Dim HFCOP As String = "     "
        'ReDim a(6, 25)
        Dim A(6, 25) As Decimal

        'Dim minpos As Integer
        Dim fcmax As Integer = 0

        For Each rowDPTMUPD0 As DataRow In dst.Tables("DPTMUPD0").Select("ITEM_CODE = '" & ITEM_CODE & "'")
            Dim P_INDEX As Integer = Val(rowDPTMUPD0.Item("P_INDEX") & "")
            Dim j As Integer = InStr("HFCOP", rowDPTMUPD0.Item("SD")) - 1
            Mid(HFCOP, j + 1, 1) = Mid("HFCOP", j + 1, 1)
            Dim QTY As Int64 = Val(rowDPTMUPD0.Item("QTY") & "")
            If j = 1 Or j = 2 Then
                A(j, P_INDEX) = A(j, P_INDEX) - QTY
                A(j, 25) = A(j, 25) - QTY
            Else
                If P_INDEX > fcmax Then
                    fcmax = P_INDEX
                End If
                A(j, P_INDEX) = A(j, P_INDEX) + QTY
                A(j, 25) = A(j, 25) + QTY
            End If
        Next

        Dim ap(,) As Integer = Nothing ' i/s (0) & o/s (1) CRITICAL_LEAD_TIME_PINDEX, min (0) & max (1) pos
        Me.Calculate_Supply_and_Demand(fcmax, MINPOS, A, ap, CRITICAL_LEAD_TIME_Pindex)

        If ITEM_SNU_CODE <> "U" And ITEM_POS_MAX > 0 And ITEM_POS_MIN > 0 Then
            If ap(0, 0) < ITEM_POS_MIN Then Exc_Msg(ITEM_CODE, "08", "")
            If ap(0, 1) > ITEM_POS_MAX Then Exc_Msg(ITEM_CODE, "09", "")
            If ap(1, 0) < ITEM_POS_MIN Then Exc_Msg(ITEM_CODE, "10", "")
            If ap(1, 1) > ITEM_POS_MAX Then Exc_Msg(ITEM_CODE, "11", "")
        End If
        ' If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CH012A02" Then Stop

        Dim P As Decimal = 0
        Dim OH As Integer = A(5, 1) - A(3, 0) - A(3, 1) + A(1, 0) + A(1, 1)  ' 5,1 - 3,0 - 3,1 + 1,0 + 1,1
        Dim DEMCUM As Integer = A(1, 0)
        Dim REMFC As Integer = 0
        For J As Integer = 1 To 25
            REMFC += A(1, J)
        Next
        If OH > DEMCUM Then
            For I As Integer = 1 To 25
                If REMFC = 0 Then Exit For
                If OH >= DEMCUM + A(1, I) Or A(1, I) <= 0 Then
                    DEMCUM += A(1, I)
                    REMFC -= A(1, I)
                    P += 1
                Else
                    P += (OH - DEMCUM) / A(1, I)
                    Exit For
                End If
            Next
        End If
        A(6, 0) = P

        For i As Integer = 0 To 6
            Dim rowDPTMRPG1 As DataRow = dst.Tables("DPTMRPG1").NewRow
            rowDPTMRPG1.Item("ITEM_CODE") = ITEM_CODE
            rowDPTMRPG1.Item("MRP_TYPE") = Format(i, "0")
            For j = 0 To 25
                rowDPTMRPG1.Item("QTY_" & Format(j, "00")) = A(i, j)
            Next j
            dst.Tables("DPTMRPG1").Rows.Add(rowDPTMRPG1)
        Next i

        ' Record Production Commitments

        VEND_CODE = ""

        ' Fill_Records("POTORDR2", ITEM_CODE)

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select($"ITEM_CODE = '{ITEM_CODE}'")

            Dim PO_ORDER_NO As String = rowPOTORDR2.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(rowPOTORDR2.Item("PO_ORDER_LNO") & "")
            VEND_WHSE_CODE = rowPOTORDR2.Item("VEND_WHSE_CODE")

            If Len(HFCOP) = 5 Then
                HFCOP = HFCOP & "X"
            End If

            If VEND_CODE <> rowPOTORDR2.Item("VEND_CODE") Then
                VEND_CODE = rowPOTORDR2.Item("VEND_CODE")
                rowAPTVEND1 = dst.Tables("APTVEND1").Rows.Find(VEND_CODE)
                rowDPTVNDI1 = dst.Tables("DPTVNDI1").Rows.Find(New String() {VEND_CODE, ITEM_CODE})
                If rowDPTVNDI1 Is Nothing Then
                    rowDPTVNDI1 = dst.Tables("DPTVNDI1").Rows.Add(New String() {VEND_CODE, ITEM_CODE})
                End If
            End If

            Dim rowPOTORDR1 As DataRow = Fill_Record("POTORDR1", PO_ORDER_NO)
            'Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find _
            '                             (New Object() {rowPOTORDRX.Item("PO_ORDER_NO"), _
            '                                            rowPOTORDRX.Item("PO_ORDER_LNO")})
            Dim PO_DATE_REQUIRED As Date = rowPOTORDR2.Item("PO_DATE_REQUIRED")

            PO_DRP = Val(rowDPTVNDI1.Item("PO_DRP") & "")
            PO_SCH_DAYS = Get_PO_SCH_DAYS(ITEM_PLAN_MAKE_BUY, rowDPTVNDI1, rowAPTVEND1)

            Dim MFG As Int64 = 0
            If PO_DRP <> 0 Then
                MFG = Val(rowPOTORDR2.Item("PO_QTY_ORD") & "") / PO_DRP
            End If
            Dim PO_DATE_COMPSDUE As Date = PO_DATE_REQUIRED.AddDays(-1 * (PO_SCH_DAYS + MFG))
            rowPOTORDR2.Item("PO_DATE_COMPSDUE") = PO_DATE_COMPSDUE

            If rowPOTORDR2.Item("PO_ORDER_TYPE") <> "W" Then
                Dim P_Index As Integer = Calc_P_Index(PO_DATE_COMPSDUE)
                'TAC.POCMAIN1.Update_POTORDR9(Me, PO_ORDER_NO, VEND_WHSE_CODE, True, False, PO_ORDER_LNO, P_Index)
                Dim rowDPTCOMMX As DataRow = dst.Tables("DPTCOMMX").NewRow
                With rowDPTCOMMX
                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("DATE_COMPSDUE") = PO_DATE_COMPSDUE
                    .Item("P_INDEX") = P_Index
                    .Item("VEND_WHSE_CODE") = VEND_WHSE_CODE
                End With
                dst.Tables("DPTCOMMX").Rows.Add(rowDPTCOMMX)
            End If
        Next
        '  Update_Record_TDA("POTORDR2")

        For Each rowDPTCOMMX As DataRow In dst.Tables("DPTCOMMX").Select($"ITEM_CODE = '{ITEM_CODE}'", "DATE_COMPSDUE")
            With rowDPTCOMMX
                Dim P_INDEX As Integer = Val(.Item("P_INDEX") & "")
                Dim PO_ORDER_NO As String = .Item("PO_ORDER_NO")
                Dim PO_ORDER_LNO As Integer = Val(.Item("PO_ORDER_LNO") & "")
                VEND_WHSE_CODE = .Item("VEND_WHSE_CODE")
                If PO_ORDER_LNO = 0 Then
                    Dim rowDPTPLAN1 As DataRow = dst.Tables("DPTPLAN1").Rows.Find(PO_ORDER_NO)
                    Plan_Production_Commit(rowDPTPLAN1)
                Else
                    ' NEXT LINE CHANGED WRITE_PLAN TO FALSE TO AVOID DUPLICATING C RECORD IN DPTMUPD0
                    ' this change may need to be revisited -
                    '  production commitments were re-located to this process for a reason many years ago,
                    '  for reasons like making sure we were using the latest BM, or latest lead times data
                    '  that is why it was disconnected from the earlier stage in this process (Load_PComm)
                    '  but the earlier process contributed to the A array, and this later stage process does not
                    '  and the A array is what is sent to DPTMRPG1 which is used in reporting
                    '  and the absence of production commitments from the A array (and DPTMRPG1) is what causedd
                    '  the Excess/Obsolete report to be blind to production commitment demand (see WJZ email to SMZ 08/04/2025 9:57AM)
                    'TAC.POCMAIN1.Update_POTORDR9(Me, PO_ORDER_NO, VEND_WHSE_CODE, True, False, PO_ORDER_LNO, P_INDEX)
                    TAC.POCMAIN1.Update_POTORDR9(Me, PO_ORDER_NO, VEND_WHSE_CODE, False, False, PO_ORDER_LNO, P_INDEX)
                End If
            End With

        Next

        If graceused = "Y" Then Exc_Msg(ITEM_CODE, "13", "")

        If Trim(HFCOP) = "H" Or Trim(HFCOP) = "" Then
            ASCDATA1.DeleteRows(dst.Tables("DPTMUPD1"), "ITEM_CODE = '" & ITEM_CODE & "'")
            ASCDATA1.DeleteRows(dst.Tables("DPTMUPD2"), "ITEM_CODE = '" & ITEM_CODE & "' and EXC_MSG_CODE = '00'")
        End If
    End Sub

    Function FND(z As String, Optional YL24 As Integer = 2) As String
        If YL24 <> 2 And YL24 <> 4 Then
            YL24 = 2
        End If
        FND = Mid(z, 5, 2) & "/" & Mid(z, 7, 2) & "/" & Mid(z, 5 - YL24, YL24)
    End Function

    Sub Regen_Append_Event_Date(
                               yyyyMMdd As String,
                               ByRef dti As Integer,
                               ByRef dt() As String,
                               ByRef dtbal() As Int64,
                               ByRef dtbalw(,) As Int64,
                               ByRef dtimax As Integer)
        dti += 1
        dtimax = dti
        If dt Is Nothing OrElse dti > UBound(dt) Then
            ReDim Preserve dt(dti + 10)
            ReDim Preserve dtbal(dti + 10)
            ReDim Preserve dtbalw(UBound(dtbalw, 1), dti + 10)
        End If
        dt(dti) = yyyyMMdd
        dtbal(dti) = dtbal(dti - 1)
        For w As Integer = 1 To UBound(dtbalw, 1)
            dtbalw(w, dti) = dtbalw(w, dti - 1)
        Next
    End Sub

    Sub Create_Plan(
                   ITEM_CODE As String,
                   QTY_PLANNED As Int64,
                   DATE_REQUIRED As Date,
                   VEND_CODE As String,
                   VEND_WHSE_CODE As String,
                   PO_LEAD_TIME As Int64,
                   PO_SCH_DAYS As Int64,
                   MFG As Int64,
                   PO_XIT_DAYS As Int64,
                   BM_ISSUE_NO As String)

        Dim PLAN_NO As String = ASCMAIN1.Next_Control_No("DPTPLAN1.PLAN_NO")
        PLAN_NO = "P" & Mid(PLAN_NO, 2)

        If ITEM_CODE = "853218006001" Then
            If ASCMAIN1.Running_in_VS Then
                ' Stop
            End If
        End If

        Dim rowDPTPLAN1 As DataRow = dst.Tables("DPTPLAN1").NewRow
        With rowDPTPLAN1
            .Item("ITEM_CODE") = ITEM_CODE
            .Item("PLAN_NO") = PLAN_NO
            .Item("DATE_ENTERED") = GDATE
            .Item("VEND_CODE") = VEND_CODE
            If BM_ISSUE_NO <> "" And VEND_WHSE_CODE <> "" Then
                .Item("PLAN_MB") = "M"
                .Item("AT_WHSE") = VEND_WHSE_CODE
                .Item("DATE_COMPSDUE") = DATE_REQUIRED.AddDays(-1 * (PO_SCH_DAYS + MFG + PO_XIT_DAYS))
                If VEND_WHSE_CODE = "" Then
                    .Item("TO_WHSE") = DP_PARM_DEF_PLAN_WHSE
                Else
                    .Item("TO_WHSE") = VEND_WHSE_CODE
                End If
                .Item("BM_ISSUE_NO") = BM_ISSUE_NO
            Else
                .Item("PLAN_MB") = "B"
            End If
            .Item("TO_WHSE") = DP_PARM_DEF_PLAN_WHSE

            If Format(DATE_REQUIRED, "MM/dd/yyyy") = "01/01/1900" Then
                .Item("DATE_REQUIRED") = GDATE
            Else
                .Item("DATE_REQUIRED") = DATE_REQUIRED
            End If
            .Item("QTY_PLANNED") = QTY_PLANNED
            .Item("DATE_DELETE") = Null
            .Item("DATE_PLAN_ACTION") = CDate(.Item("DATE_REQUIRED")).AddDays(-1 * (Val(ROWs("DPTPARM1").Item("DP_PARM_PLNR_ANAL_DAYS") & "") + PO_LEAD_TIME + PO_SCH_DAYS + MFG + PO_XIT_DAYS))
            .Item("DATE_PO_ISSUE") = CDate(.Item("DATE_REQUIRED")).AddDays(-1 * (PO_LEAD_TIME + PO_SCH_DAYS + MFG + PO_XIT_DAYS))
            .Item("PLAN_TYPE") = "S"
            .Item("ACT_MSG_FLAG") = ""
            .Item("ACT_MSG_DATE") = Null
            .Item("DATE_VEND_SHIP") = CDate(.Item("DATE_REQUIRED")).AddDays(-1 * (PO_XIT_DAYS))
            If Format(.Item("DATE_PO_ISSUE"), "yyyyMMdd") < Format(GDATE, "yyyyMMdd") Then
                Exc_Msg(ITEM_CODE, "03", "")
            End If
            If Format(.Item("DATE_PLAN_ACTION"), "yyyyMMdd") <= Format$(GDATE, "yyyyMMdd") Then
                Act_Msg(ITEM_CODE, "01", "")
            End If
        End With
        dst.Tables("DPTPLAN1").Rows.Add(rowDPTPLAN1)

        If rowDPTPLAN1.Item("PLAN_MB") = "M" Then
            'TAC.POCMAIN1.Update_POTORDR9(Me, PLAN_NO, VEND_WHSE_CODE, True, False, , P_Index)
            'PCOM_Commit("", bmissue, q, DATE_REQUIRED.AddDays(-1 * (PO_SCH_DAYS + MFG + PO_XIT_DAYS)))

            'Plan_Production_Commit(rowDPTPLAN1)

            Dim rowDPTCOMMX As DataRow = dst.Tables("DPTCOMMX").NewRow
            With rowDPTCOMMX
                .Item("PO_ORDER_NO") = PLAN_NO
                .Item("PO_ORDER_LNO") = 0
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("DATE_COMPSDUE") = rowDPTPLAN1.Item("DATE_COMPSDUE")
                .Item("P_INDEX") = 0 ' P_INDEX
                .Item("VEND_WHSE_CODE") = VEND_WHSE_CODE
            End With

            dst.Tables("DPTCOMMX").Rows.Add(rowDPTCOMMX)

        End If
        If dst.Tables("DPTMUPDS").Rows.Find(New Object() {ITEM_CODE, DATE_REQUIRED, DP_PARM_DEF_PLAN_WHSE}) Is Nothing Then
            dst.Tables("DPTMUPDS").Rows.Add(New Object() {ITEM_CODE, DATE_REQUIRED, DP_PARM_DEF_PLAN_WHSE})
        End If

        Dim rowDPTMUPD0 As DataRow = dst.Tables("DPTMUPD0").NewRow
        With rowDPTMUPD0
            .Item("ITEM_CODE") = ITEM_CODE
            .Item("SD") = "P"
            .Item("BM_ISSUE_NO") = BM_ISSUE_NO
            .Item("DATE_REQ") = DATE_REQUIRED
            .Item("QTY_PLN") = QTY_PLANNED
            .Item("PO_ORDER_NO") = PLAN_NO
            .Item("VEND_CODE") = VEND_CODE
            .Item("WHSE_CODE") = DP_PARM_DEF_PLAN_WHSE
            .Item("NOTES") = "System Plan"
            .Item("P_INDEX") = Calc_P_Index(DATE_REQUIRED)
        End With
        dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)

    End Sub

    Sub Plan_Production_Commit(rowDPTPLAN1 As DataRow)

        Dim PLAN_NO As String = rowDPTPLAN1.Item("PLAN_NO")
        Dim BM_PROD_ITEM As String = rowDPTPLAN1.Item("ITEM_CODE")
        Dim BM_ISSUE_NO As String = rowDPTPLAN1.Item("BM_ISSUE_NO")
        Dim QTY_PLANNED As Int64 = Val(rowDPTPLAN1.Item("QTY_PLANNED") & "")
        Dim VEND_WHSE_CODE As String = rowDPTPLAN1.Item("AT_WHSE")
        Dim DATE_COMPSDUE As Date = rowDPTPLAN1.Item("DATE_COMPSDUE")
        Dim VEND_CODE As String = rowDPTPLAN1.Item("VEND_CODE")

        Dim P_INDEX As Integer = Calc_P_Index(DATE_COMPSDUE)

        Dim TBL As DataTable = TAC.POCMAIN1.Get_BM(Me, BM_PROD_ITEM, "C", BM_ISSUE_NO,
                    False, True, "C", QTY_PLANNED, VEND_WHSE_CODE, "HOPCA")

        dst.Tables("POTORDR9").Rows.Clear()

        For Each rowBMTMAIN3 As DataRow In TBL.Select("")

            Dim rowPOTORDR9 As DataRow = dst.Tables("POTORDR9").NewRow
            rowPOTORDR9.Item("PO_ORDER_NO") = PLAN_NO
            rowPOTORDR9.Item("PO_ORDER_LNO") = 0
            rowPOTORDR9.Item("ITEM_CODE") = rowBMTMAIN3.Item("BM_COMP_ITEM")
            rowPOTORDR9.Item("PO_QTY_COM") = rowBMTMAIN3.Item("QTY_COM")
            dst.Tables("POTORDR9").Rows.Add(rowPOTORDR9)

            If dst.Tables("DPTMUPDS").Rows.Find(New Object() {rowBMTMAIN3.Item("BM_COMP_ITEM"), DATE_COMPSDUE, VEND_WHSE_CODE}) Is Nothing Then
                dst.Tables("DPTMUPDS").Rows.Add(New Object() {rowBMTMAIN3.Item("BM_COMP_ITEM"), DATE_COMPSDUE, VEND_WHSE_CODE})
            End If

            Dim rowDPTMUPD0 As DataRow = dst.Tables("DPTMUPD0").NewRow
            With rowDPTMUPD0
                .Item("ITEM_CODE") = rowBMTMAIN3.Item("BM_COMP_ITEM")
                .Item("SD") = "C"
                .Item("NOTES") = "Plan C for" & .Item("ITEM_CODE")
                .Item("BM_ISSUE_NO") = BM_ISSUE_NO
                .Item("DATE_REQ") = DATE_COMPSDUE
                .Item("QTY_REQ") = rowBMTMAIN3.Item("QTY_COM")
                .Item("PO_ORDER_NO") = PLAN_NO
                .Item("VEND_CODE") = VEND_CODE
                .Item("WHSE_CODE") = VEND_WHSE_CODE
                .Item("P_INDEX") = P_INDEX
            End With
            dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)
        Next

        Update_Record_TDA("POTORDR9")

        TAC.POCMAIN1.Production_Commit(1, PLAN_NO, , VEND_WHSE_CODE)

        'Dim rowDPTCOMMX As DataRow = dst.Tables("DPTCOMMX").NewRow
        'With rowDPTCOMMX
        '    .Item("PO_ORDER_NO") = PLAN_NO
        '    .Item("PO_ORDER_LNO") = 0
        '    .Item("ITEM_CODE") = BM_PROD_ITEM
        '    .Item("DATE_COMPSDUE") = DATE_COMPSDUE
        '    .Item("P_INDEX") = P_INDEX
        '    .Item("VEND_WHSE_CODE") = VEND_WHSE_CODE
        'End With

        'dst.Tables("DPTCOMMX").Rows.Add(rowDPTCOMMX)

    End Sub

    Sub Calculate_Supply_and_Demand(
                                   fcmax As Decimal,
                                   minpos As Decimal,
                                   ByRef A(,) As Decimal,
                                   ByRef AP(,) As Integer,
                                   ByRef CRITICAL_LEAD_TIME_Pindex As Integer)
        '    ASCMAIN1.Progress("Supply & Demand Grid", "")
        '        Dim i As Integer
        Dim j As Integer = 0
        'Dim t As Long

        Dim dmax As Integer = 0
        For i As Integer = 24 To 0 Step -1
            If A(1, i) > 0 Or A(2, i) > 0 Then
                dmax = i
                Exit For
            End If
        Next i

        ReDim AP(1, 1)
        AP(0, 0) = 999
        AP(1, 0) = 999

        For i As Integer = 0 To 24
            A(5, i) = A(0, i) - A(1, i) - A(2, i) + A(3, i) + A(4, i)
            '   EOM = BOM     - FC      -PCOM     + PO      + PLAN
            A(0, i + 1) = A(5, i)
            ' BOM next  = EOM this
            Dim t As Int64 = 0
            If i < 24 Then
                j = i + 1
                Do While A(5, i) > t + A(1, j) And A(1, j) > 0 And j < 24
                    t = t + A(1, j)
                    j = j + 1
                    A(6, i) += 1
                Loop
                If A(1, j) > 0 And j < 24 Then
                    A(6, i) += (A(5, i) - t) / A(1, j)
                End If

            End If
            If i < CRITICAL_LEAD_TIME_Pindex Then
                j = 0
            Else
                j = 1
            End If
            If A(6, i) < AP(j, 0) And i <= fcmax - minpos Then
                AP(j, 0) = A(6, i)
            End If
            If A(6, i) > AP(j, 1) Then
                AP(j, 1) = A(6, i)
            End If
        Next i
        A(5, 25) = A(5, 24)

    End Sub

    Sub Load_PO()
        ASCMAIN1.Progress("Purchase Orders", "")

        ASCMAIN1.sql = "Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO
            , POTORDR2.ITEM_CODE, POTORDR2.PO_DATE_REQUIRED, POTORDR2.PO_QTY_OPN
            , POTORDR1.VEND_CODE, POTORDR2.WHSE_CODE, POTORDR2.BM_ISSUE_NO, PO_DATE_ORDERED
             from POTORDR2, POTORDR1, ICTWHSE1
             where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
               and POTORDR2.PO_QTY_OPN <> 0
               and POTORDR2.PO_STATUS = 'O'
               and ICTWHSE1.WHSE_CODE = POTORDR2.WHSE_CODE
               and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'"

        POTORDRX = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL($"Alter Table {POTORDRX} Add Primary Key (PO_ORDER_NO, PO_ORDER_LNO)")

        sqlICTPINVX = "Select ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO, ICTPINV2.PINV_NO, ICTPINV2.PINV_LNO
                , ICTPINV2.ITEM_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV1.WHSE_CODE, ICTPINV1.VEND_CODE
                , ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE
                , ICTPINV1.ETA_DATE, ICTPINV2.PINV_QTY, ICTPINV1.ETA_DATE ETA_DATE_DC, ICTITEM1.PORT_CODE
                 from ICTPINV1,ICTPINV2, ICTITEM1
                 where (ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO) In (
                Select POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO
                             from POTORDR2,POTORDR1,ICTWHSE1
                             where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
                               And POTORDR2.PO_QTY_OPN <> 0
                               and POTORDR2.PO_STATUS = 'O'
                               and ICTWHSE1.WHSE_CODE = POTORDR2.WHSE_CODE
                               and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'
                ) And ICTPINV1.PINV_STATUS = 'O'
And ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE
AND ICTPINV1.PINV_NO = ICTPINV2.PINV_NO"
        ICTPINVX = ASCMAIN1.Temp_Table(sqlICTPINVX)
        ASCDATA1.ExecuteSQL($"Alter Table {ICTPINVX} Add BM_ISSUE_NO VARCHAR2(2)")

        ASCMAIN1.sql = $"Select PO_ORDER_NO, PO_ORDER_LNO, Sum (PINV_QTY) PINV_QTY from {ICTPINVX} group by PO_ORDER_NO, PO_ORDER_LNO"
        ASCMAIN1.sql = $"
            Begin
              Declare Cursor C1 is {ASCMAIN1.sql};
              Begin
                For R1 in C1 Loop
                    Update {POTORDRX} 
                      Set PO_QTY_OPN = NVL(PO_QTY_OPN,0) - NVL(R1.PINV_QTY,0)
                      where PO_ORDER_NO = R1.PO_ORDER_NO
                        and PO_ORDER_LNO = R1.PO_ORDER_LNO;
                End Loop;
              End;
            End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Insert into {ICTPINVX}
            Select X.PO_ORDER_NO, X.PO_ORDER_LNO, '000000' PINV_NO, 1 PINV_LNO
                , X.ITEM_CODE, NULL INV_NUM, X.PO_DATE_ORDERED INV_DATE, X.WHSE_CODE, X.VEND_CODE
                , NULL VESSEL_NAME, NULL BOL_NO, 'Open PO' CONTAINER_NO, NULL SHIP_DATE
                , X.PO_DATE_REQUIRED ETA_DATE, X.PO_QTY_OPN PINV_QTY, X.PO_DATE_REQUIRED ETA_DATE_DC, ICTITEM1.PORT_CODE, X.BM_ISSUE_NO
                from {POTORDRX} X, ICTITEM1 where ICTITEM1.ITEM_CODE = X.ITEM_CODE"
        ASCDATA1.ExecuteSQL()

        '    & " O EDT864T3%ROWTYPE;" & vbCrLf _
        '    & " BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
        '    & " SELECT EDT864T3.* INTO O FROM EDT864T3 WHERE EDI_DOC_SEQ_NO = R1.EDI_DOC_SEQ_NO" & vbCrLf _
        '    & " AND EDI_DTL_SEQ = R1.EDI_DTL_SEQ AND EDI_MSG_SEQ = R1.EDI_MSG_SEQ + 2;" & vbCrLf _
        '    & " IF NOT SQL%NOTFOUND THEN" & vbCrLf _

        '& "       If SQL%NOTFOUND Then" & vbCrLf _
        '& "         xPINV_LT := xPOTPARM1.PO_PARM_PINV_LT;" & vbCrLf _
        '& "       Else" & vbCrLf _
        '& "         xPINV_LT := NVL(xICTPORT2.ETD_TO_ETA,xPOTPARM1.PO_PARM_PINV_LT);" & vbCrLf _
        '& "       End If;" & vbCrLf _

        'ASCMAIN1.sql = $"Update {ICTPINVX} Set ETA_DATE = INV_DATE + 49 WHERE CONTAINER_NO IS NULL"
        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & "  Declare" & vbCrLf _
            & "    xPOTPARM1 POTPARM1%ROWTYPE;" & vbCrLf _
            & "    xICTPORT2 ICTPORT2%ROWTYPE;" & vbCrLf _
            & "    xPORT_CODE VARCHAR2(6);" & vbCrLf _
            & "    xWHSE_CODE VARCHAR2(6);" & vbCrLf _
            & "    xPINV_LT NUMBER(3,0);" & vbCrLf _
            & "    Cursor C1 is " & vbCrLf _
            & $"      Select * from {ICTPINVX}" & vbCrLf _
            & "        where CONTAINER_NO is Null for Update;" & vbCrLf _
            & "  Begin" & vbCrLf _
            & "    Select POTPARM1.* into xPOTPARM1 from POTPARM1 where PO_PARM_KEY = 'Z';" & vbCrLf _
            & "    For R1 in C1 Loop" & vbCrLf _
            & "       xWHSE_CODE := R1.WHSE_CODE;" & vbCrLf _
            & "       xPORT_CODE := R1.PORT_CODE;" & vbCrLf _
            & "       If NVL(xPORT_CODE,'') = '' Then" & vbCrLf _
            & "         xPORT_CODE := xPOTPARM1.PO_PARM_PINV_PORT;" & vbCrLf _
            & "       End If;" & vbCrLf _
            & "       Begin" & vbCrLf _
            & "         Select ICTPORT2.* into xICTPORT2 from ICTPORT2" & vbCrLf _
            & "           where PORT_CODE = xPORT_CODE and WHSE_CODE = xWHSE_CODE;" & vbCrLf _
            & "         xPINV_LT := NVL(xICTPORT2.ETD_TO_ETA,xPOTPARM1.PO_PARM_PINV_LT);" & vbCrLf _
            & "       Exception" & vbCrLf _
            & "         When NO_DATA_FOUND Then xPINV_LT := xPOTPARM1.PO_PARM_PINV_LT;" & vbCrLf _
            & "       End;" & vbCrLf _
            & $"      Update {ICTPINVX} Set" & vbCrLf _
            & "         CONTAINER_NO = 'Inv Date ' || TO_CHAR(xPINV_LT)" & vbCrLf _
            & "       , ETA_DATE = INV_DATE + xPINV_LT" & vbCrLf _
            & "       where PO_ORDER_NO = R1.PO_ORDER_NO and PO_ORDER_LNO = R1.PO_ORDER_LNO;" & vbCrLf _
            & "    End Loop;" & vbCrLf _
            & "  End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Update {ICTPINVX} Set ETA_DATE_DC = WEEKDAYS_FROM (ETA_DATE, 5)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"
            Select ITEM_CODE, PO_ORDER_NO, PO_ORDER_LNO, VEND_CODE, WHSE_CODE
                , BM_ISSUE_NO, ETA_DATE_DC PO_DATE_REQUIRED, PINV_QTY PO_QTY_OPN
                from {ICTPINVX}"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "ITEM_CODE")
            Dim rowDPTMUPD0 As DataRow = dst.Tables("DPTMUPD0").NewRow
            With rowDPTMUPD0
                .Item("ITEM_CODE") = row.Item("ITEM_CODE")
                .Item("SD") = "O"
                .Item("BM_ISSUE_NO") = row.Item("BM_ISSUE_NO")
                .Item("DATE_REQ") = row.Item("PO_DATE_REQUIRED")
                .Item("QTY_ORD") = row.Item("PO_QTY_OPN")
                .Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                .Item("VEND_CODE") = row.Item("VEND_CODE")
                .Item("WHSE_CODE") = row.Item("WHSE_CODE")
                .Item("P_INDEX") = Calc_P_Index(row.Item("PO_DATE_REQUIRED"))
            End With
            dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)

            If Format(row.Item("PO_DATE_REQUIRED"), "yyyyMMdd") < Format(GDATE, "yyyyMMdd") Then
                Exc_Msg(row.Item("ITEM_CODE"), "05", "")
            End If
        Next
    End Sub

    Function Calc_P_Index(d As Date) As Integer
        Dim z As String = Format(d, "yyyyMM")
        Dim i As Integer = 0
        If z >= m24(1) Then
            For i = 1 To 24
                If z = Mid(m24(i), 1, 6) Then
                    Exit For
                End If
            Next
        End If
        Return i
    End Function

    Sub Load_Plans()
        ' LATER, WHEN WE SUPPORT PLANNER ENTERED PLANS
    End Sub

    Sub Load_Prod_Comm()

        'ASCMAIN1.sql = "Select POTORDR9.*, POTORDR2.PO_DATE_COMPSDUE" & vbCrLf _
        '    & ", POTORDR2.WHSE_CODE, POTORDR2.PO_DATE_REQUIRED" & vbCrLf _
        '    & ", POTORDR1.VEND_CODE" & vbCrLf _
        '    & " from POTORDR9, POTORDR2, POTORDR1" & vbCrLf _
        '    & " where POTORDR1.PO_ORDER_NO = POTORDR9.PO_ORDER_NO" & vbCrLf _
        '    & "   And POTORDR2.PO_ORDER_NO = POTORDR9.PO_ORDER_NO" & vbCrLf _
        '    & "   And POTORDR2.PO_ORDER_LNO = POTORDR9.PO_ORDER_LNO" & vbCrLf _
        '    & "   And POTORDR1.PO_ORDER_TYPE = 'W'"

        ASCMAIN1.sql = "Select POTORDR9.*, POTORDR2.PO_DATE_COMPSDUE" & vbCrLf _
            & ", POTORDR2.WHSE_CODE, POTORDR2.PO_DATE_REQUIRED, POTORDR2.BM_ISSUE_NO" & vbCrLf _
            & ", POTORDR1.VEND_CODE" & vbCrLf _
            & " from POTORDR9, POTORDR2, POTORDR1" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR9.PO_ORDER_NO" & vbCrLf _
            & "   And POTORDR2.PO_ORDER_NO = POTORDR9.PO_ORDER_NO" & vbCrLf _
            & "   And POTORDR2.PO_ORDER_LNO = POTORDR9.PO_ORDER_LNO" '  & vbCrLf _
        ' & "   And POTORDR1.PO_ORDER_TYPE = 'R-not'"

        Dim df As String = "PO_DATE_COMPSDUE"

        For Each rowPOTORDR9 As DataRow In ASCDATA1.GetDataTable.Select("", "ITEM_CODE")
            Dim PO_QTY_COM As Int64 = rowPOTORDR9.Item("PO_QTY_COM")
            Dim ITEM_CODE As String = rowPOTORDR9.Item("ITEM_CODE")
            Dim WHSE_CODE As String = rowPOTORDR9.Item("WHSE_CODE")
            Dim BM_ISSUE_NO As String = rowPOTORDR9.Item("BM_ISSUE_NO") & ""
            Dim rowDPTMUPD0 As DataRow = dst.Tables("DPTMUPD0").NewRow
            With rowDPTMUPD0
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("SD") = "C"
                .Item("BM_ISSUE_NO") = BM_ISSUE_NO
                ' The Compsdue date shouldn't be null, but this is a failsafe for when it is, we will default to the date required
                If rowPOTORDR9.Item(df) & "" = "" Then
                    df = "PO_DATE_REQUIRED"
                End If
                .Item("NOTES") = "MRP Update"

                .Item("DATE_REQ") = rowPOTORDR9.Item(df)
                .Item("QTY_REQ") = PO_QTY_COM
                .Item("PO_ORDER_NO") = rowPOTORDR9.Item("PO_ORDER_NO")
                .Item("VEND_CODE") = rowPOTORDR9.Item("VEND_CODE")
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("P_INDEX") = Calc_P_Index(rowPOTORDR9.Item(df))
            End With
            dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)

            Dim rowICTSTAT2 As DataRow = dst.Tables("ICTSTAT2").Rows.Find(New String() {ITEM_CODE, WHSE_CODE})
            If rowICTSTAT2 Is Nothing Then
                rowICTSTAT2 = dst.Tables("ICTSTAT2").NewRow
                rowICTSTAT2.Item("ITEM_CODE") = ITEM_CODE
                rowICTSTAT2.Item("WHSE_CODE") = WHSE_CODE
                dst.Tables("ICTSTAT2").Rows.Add(rowICTSTAT2)
            End If
            rowICTSTAT2.Item("WHSE_QTY_COMM") = Val(rowICTSTAT2.Item("WHSE_QTY_COMM") & "") + PO_QTY_COM
        Next

    End Sub

    Sub Load_Status()
        ASCMAIN1.Progress("Item Status && Activity", "")
        ASCMAIN1.sql = "Select ITEM_CODE, WHSE_CODE, SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND, SUM (WHSE_QTY_SHP) WHSE_QTY_SHP from (" & vbCrLf _
            & "Select ICTSTAT1.ITEM_CODE, ICTSTAT1.WHSE_CODE, 0 WHSE_QTY_ON_HAND, ICTSTAT1.WHSE_QTY_SHP" & vbCrLf _
            & " from ICTSTAT1,ICTWHSE1" & vbCrLf _
            & " where ICTWHSE1.WHSE_CODE = ICTSTAT1.WHSE_CODE" & vbCrLf _
            & "   and ICTSTAT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and (NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1')" & vbCrLf _
            & " union " & vbCrLf _
            & "Select ICTSTAT2.ITEM_CODE, ICTSTAT2.WHSE_CODE, ICTSTAT2.WHSE_QTY_ON_HAND, 0 WHSE_QTY_SHP" & vbCrLf _
            & " from ICTSTAT2,ICTWHSE1" & vbCrLf _
            & " where ICTWHSE1.WHSE_CODE = ICTSTAT2.WHSE_CODE" & vbCrLf _
            & "   and (NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1')" & vbCrLf _
            & ") group by ITEM_CODE, WHSE_CODE" & vbCrLf _
            & " having SUM (WHSE_QTY_ON_HAND) <> 0 OR SUM (WHSE_QTY_SHP) <> 0"

        For Each rowICTSTATX As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim WHSE_QTY_ON_HAND As Int64 = Val(rowICTSTATX.Item("WHSE_QTY_ON_HAND") & "")
            Dim WHSE_QTY_SHP As Int64 = Val(rowICTSTATX.Item("WHSE_QTY_SHP") & "")
            Dim ITEM_CODE As String = rowICTSTATX.Item("ITEM_CODE")
            If WHSE_QTY_ON_HAND <> 0 Or WHSE_QTY_SHP <> 0 Then
                Dim rowDPTMUPD0 As DataRow = dst.Tables("DPTMUPD0").NewRow
                With rowDPTMUPD0
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("SD") = "H"
                    .Item("DATE_REQ") = "01/01/1900"
                    .Item("QTY_OH") = WHSE_QTY_ON_HAND + WHSE_QTY_SHP
                    .Item("NOTES") = "OH + MTD Ship:" & CStr(WHSE_QTY_ON_HAND + WHSE_QTY_SHP)
                    .Item("WHSE_CODE") = rowICTSTATX.Item("WHSE_CODE")
                    .Item("P_INDEX") = 0
                End With
                dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)

                If WHSE_QTY_ON_HAND < 0 Then Exc_Msg(ITEM_CODE, "01", "")
            End If
        Next

    End Sub

    Sub Load_Forecast()
        ASCMAIN1.Progress("Forecasts", "")
        Dim MARKET_CODE As String = ""
        Dim WHSE_CODE As String = ""
        Dim rowSOTMKTC1 As DataRow = Nothing

        For Each rowDPTITMFX As DataRow In dst.Tables("DPTITMFX").Select("", "ITEM_CODE")
            Dim ITEM_CODE As String = rowDPTITMFX.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)

            'Dim ITEM_BUFFER_QTY As Int32 = Val(rowICTITEM1.Item("ITEM_BUFFER_QTY") & "")
            'Dim ITEM_BUFFER_PCT As Decimal = Val(rowICTITEM1.Item("ITEM_BUFFER_PCT") & "")

            'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CC004P80" Then Stop

            If MARKET_CODE <> rowDPTITMFX.Item("MARKET_CODE") Then
                MARKET_CODE = rowDPTITMFX.Item("MARKET_CODE")
                rowSOTMKTC1 = dst.Tables("SOTMKTC1").Rows.Find(MARKET_CODE)
            End If
            WHSE_CODE = DP_PARM_DEF_PLAN_WHSE
            If rowSOTMKTC1 IsNot Nothing Then
                If rowSOTMKTC1.Item("WHSE_CODE") & "" = "" Then
                    WHSE_CODE = DP_PARM_DEF_PLAN_WHSE
                Else
                    WHSE_CODE = rowSOTMKTC1.Item("WHSE_CODE") & ""
                End If
            End If
            Dim ITEM_FC_DAY_REQ As Integer = Val(rowICTITEM1.Item("ITEM_FC_DAY_REQ") & "")
            If ITEM_FC_DAY_REQ = 0 Then ITEM_FC_DAY_REQ = Val(ROWs("DPTPARM1").Item("DP_PARM_FC_REQ_DAY") & "")
            If ITEM_FC_DAY_REQ = 0 Then ITEM_FC_DAY_REQ = 1
            For i As Integer = 0 To 25
                Dim QTY_REQ As Int64 = Val(rowDPTITMFX.Item(i + 2) & "")

                'If ITEM_BUFFER_QTY > 0 Then
                '    If QTY_REQ > 0 Then
                '        QTY_REQ += ITEM_BUFFER_QTY
                '    End If
                'ElseIf ITEM_BUFFER_PCT > 0 Then
                '    QTY_REQ = System.Math.Round(QTY_REQ * (100 + ITEM_BUFFER_PCT) / 100, 0)
                'End If

                If QTY_REQ <> 0 Then
                    Dim rowDPTMUPD0 As DataRow = dst.Tables("DPTMUPD0").NewRow
                    With rowDPTMUPD0
                        .Item("ITEM_CODE") = ITEM_CODE
                        .Item("SD") = "F"
                        Dim Z As String = ""
                        If i = 0 Then
                            Z = m24(1)
                        Else
                            Z = m24(i)
                        End If
                        .Item("DATE_REQ") = Mid$(Z, 5, 2) & "/" & CStr(ITEM_FC_DAY_REQ) & "/" & Mid$(Z, 1, 4)

                        .Item("QTY_REQ") = QTY_REQ
                        .Item("NOTES") = "Forecast " & MARKET_CODE & IIf(i = 0, " (PD)", "")
                        .Item("WHSE_CODE") = WHSE_CODE
                        .Item("P_INDEX") = i
                    End With
                    dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)
                End If
            Next i
        Next

        MARKET_CODE = ""
        For Each rowSOTORDRM As DataRow In dst.Tables("SOTORDRM").Select("", "MARKET_CODE")

            Dim ITEM_CODE As String = rowSOTORDRM.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            Dim ITEM_FC_DAY_REQ As Integer = Val(rowICTITEM1.Item("ITEM_FC_DAY_REQ") & "")
            If ITEM_FC_DAY_REQ = 0 Then ITEM_FC_DAY_REQ = Val(ROWs("DPTPARM1").Item("DP_PARM_FC_REQ_DAY") & "")
            If ITEM_FC_DAY_REQ = 0 Then ITEM_FC_DAY_REQ = 1

            'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CC010A01" Then Stop

            Dim SNU As String = rowICTITEM1.Item("ITEM_SNU_CODE") & ""
            Dim BP As String = rowICTITEM1.Item("ITEM_BASIC_PROMO") & ""

            If SNU = "" Then SNU = "S"
            If BP = "" Then BP = "B"

            If MARKET_CODE <> rowSOTORDRM.Item("MARKET_CODE") & "" Then
                MARKET_CODE = rowSOTORDRM.Item("MARKET_CODE") & ""
                rowSOTMKTC1 = dst.Tables("SOTMKTC1").Rows.Find(MARKET_CODE)
                WHSE_CODE = DP_PARM_DEF_PLAN_WHSE
                If rowSOTMKTC1 IsNot Nothing Then
                    If rowSOTMKTC1.Item("WHSE_CODE") & "" = "" Then
                        WHSE_CODE = DP_PARM_DEF_PLAN_WHSE
                    Else
                        WHSE_CODE = rowSOTMKTC1.Item("WHSE_CODE") & ""
                    End If
                End If
            End If

            Dim PAST_DUE_FC As String = "" ' THIS SECTION MUST BE CONSISTENT IN MRP UPDATE AND DP INQUIRY
            If rowSOTMKTC1 IsNot Nothing Then
                PAST_DUE_FC = rowSOTMKTC1.Item("PAST_DUE_FC_" & SNU & BP) & ""
                If PAST_DUE_FC = "" Then
                    PAST_DUE_FC = "0"
                End If
            Else
                'PAST_DUE_FC = "A"
                PAST_DUE_FC = "0" ' PER NF 06/27/25
            End If


            'Dim dv As DataView = New DataView(dst.Tables("SOTORDRM"), "ITEM_CODE = 'MB008A01'", "", DataViewRowState.CurrentRows)
            'Dim tb As DataTable = dv.ToTable


            'If ASCMAIN1.Running_in_VS And ITEM_CODE = "KS001A03" Then Stop

            If "0P".Contains(PAST_DUE_FC) Then

                Dim rowDPTITMFX As DataRow = dst.Tables("DPTITMFX").Rows.Find(New String() {ITEM_CODE, MARKET_CODE})

                For i As Integer = 0 To 23

                    Dim FORECAST As Int64 = 0 ' NOT DPTITMFX IS 1 FOR CURR PRD, AND SOTORDRM IS 0 FOR CURR PRD
                    If rowDPTITMFX IsNot Nothing Then FORECAST = Val(rowDPTITMFX.Item("FC_" & Format(i + 1, "00")) & "")

                    Dim QTY_SHP_COM_REL As Int64 = Val(rowSOTORDRM.Item("SO" & Format(i, "00")) & "")
                    If i = 0 Then
                        If rowDPTITMFX IsNot Nothing Then FORECAST += Val(rowDPTITMFX.Item("FC_PD") & "")
                        QTY_SHP_COM_REL += Val(rowSOTORDRM.Item("SOPD") & "")
                        Dim rowSOTINVHX As DataRow = dst.Tables("SOTINVHX").Rows.Find(New String() {MARKET_CODE, ITEM_CODE})
                        If rowSOTINVHX IsNot Nothing Then
                            QTY_SHP_COM_REL += Val(rowSOTINVHX.Item("QTYI") & "")
                        End If
                    End If

                    If QTY_SHP_COM_REL > FORECAST Then

                        Dim QTY As Int64 = QTY_SHP_COM_REL - FORECAST
                        Dim rowDPTMUPD0 As DataRow = dst.Tables("DPTMUPD0").NewRow
                        With rowDPTMUPD0
                            .Item("ITEM_CODE") = ITEM_CODE
                            .Item("SD") = "F"
                            '  If ASCMAIN1.Running_in_VS And ITEM_CODE = "CH005A01" Then Stop
                            Dim Z As String = m24(i + 1)
                            .Item("DATE_REQ") = Mid$(Z, 5, 2) & "/" & CStr(ITEM_FC_DAY_REQ) & "/" & Mid$(Z, 1, 4)
                            .Item("QTY_REQ") = QTY
                            .Item("NOTES") = "Overship " & MARKET_CODE & ":" & CStr(QTY)
                            .Item("WHSE_CODE") = DP_PARM_DEF_PLAN_WHSE
                            .Item("P_INDEX") = i ' 0
                        End With
                        dst.Tables("DPTMUPD0").Rows.Add(rowDPTMUPD0)

                    End If
                Next
            End If
        Next
    End Sub

    Sub Exc_Msg(ITEM_CODE As String, EXC_MSG_CODE As String, SUFFIX As String)
        If dst.Tables("DPTMUPD2").Rows.Find(New Object() {ITEM_CODE, EXC_MSG_CODE}) Is Nothing Then
            dst.Tables("DPTMUPD2").Rows.Add(New Object() {ITEM_CODE, EXC_MSG_CODE, SUFFIX})
        End If
    End Sub

    Sub Act_Msg(ITEM_CODE As String, ACT_MSG_CODE As String, SUFFIX As String)
        If dst.Tables("DPTMUPD3").Rows.Find(New Object() {ITEM_CODE, ACT_MSG_CODE}) Is Nothing Then
            dst.Tables("DPTMUPD3").Rows.Add(New Object() {ITEM_CODE, ACT_MSG_CODE, SUFFIX})
        End If
    End Sub

    Sub Load_DataTables()

        ASCMAIN1.Progress("Load Work Files", "")

        ASCMAIN1.Progress("-", "Item Master")
        'NVL(SOTSDIV1.EXCLUDE_FROM_MRP,'0') EXCLUDE_FROM_MRP
        ASCMAIN1.sql = "Select ICTITEM1.*" & vbCrLf _
            & ", 0 MAX_LEVEL, BMTMAIN1.BM_ISSUE_COUNTER, '0' EXCLUDE_FROM_MRP, ICTCOLL1.BRAND_CODE " & vbCrLf _
            & " from ICTITEM1, BMTMAIN1, ICTCOLL1, SOTSDIV1" & vbCrLf _
            & " where BMTMAIN1.BM_PROD_ITEM (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = ICTITEM1.SALES_DIVISION_CODE"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)
        Fill_Records("ICTITEM1")

        ASCMAIN1.Progress("-", "Item Status")
        ASCMAIN1.sql = "Select * from ICTSTAT2"
        Dim ICTSTAT2 As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTSTAT2 & " Add Primary Key (ITEM_CODE, WHSE_CODE)")
        ASCDATA1.ExecuteSQL("Update " & ICTSTAT2 & " set WHSE_QTY_COMM = 0, WHSE_QTY_PLAN = 0")
        Create_TDA(dst.Tables.Add("ICTSTAT2"), ICTSTAT2, "**", 0, True, "", 2)
        Fill_Records("ICTSTAT2")

        ASCMAIN1.sql = "Select NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))) MARKET_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.INV_TYPE,'I',SOTINVH2.ORDR_QTY_SHIP)) QTYI" & vbCrLf _
            & ", SUM (DECODE(SOTINVH2.INV_TYPE,'C',SOTINVH2.ORDR_QTY_SHIP)) QTYC" & vbCrLf _
            & " from SOTINVH2,ARTCUST1,SOTTCLS1,SOTMKTC1,ICTWHSE1,SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
            & " where SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ICTWHSE1.WHSE_CODE = SOTINVH2.WHSE_CODE" & vbCrLf _
            & "   and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'" & vbCrLf _
            & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
            & " group by NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTINVH2.ITEM_CODE"
        ASCMAIN1.sql = "Select X.*, DPTITMF1.FORECAST" & vbCrLf _
            & " from DPTITMF1,(" & vbCrLf & ASCMAIN1.sql & ") X" & vbCrLf _
            & " where DPTITMF1.OPS_YYYYPP (+) = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and DPTITMF1.OPS_YYYYPP_FC (+) = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and DPTITMF1.MARKET_CODE (+) = X.MARKET_CODE" & vbCrLf _
            & "   and DPTITMF1.ITEM_CODE (+) = X.ITEM_CODE"

        Create_TDA(dst.Tables.Add, "SOTINVHX", "**", 0, False, "", 2)
        Fill_Records("SOTINVHX")

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
        Dim LM As Date = rowGLTPARM2.Item("PRD_END_DATE") ' d24(0)
        ASCMAIN1.sql = "Select NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))) MARKET_CODE, SOTORDR2.ITEM_CODE" & vbCrLf _
        & ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE <= '" & Format(LM, "dd-MMM-yyyy") & "' THEN " _
        & " NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) SOPD" & vbCrLf
        For P = 0 To 23
            Dim TM As Date = d24(P + 1)
            ASCMAIN1.sql &= ", SUM (CASE WHEN SOTORDR1.ORDR_SHIP_DATE " _
            & " > '" & Format(LM, "dd-MMM-yyyy") & "'" _
            & " AND SOTORDR1.ORDR_SHIP_DATE <= '" & Format(TM, "dd-MMM-yyyy") & "'" _
            & " THEN NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) ELSE 0 END) SO" & Format(P, "00") & vbCrLf
            LM = TM
        Next
        ASCMAIN1.sql &= " from SOTORDR2,SOTORDR1,ARTCUST1,SOTTCLS1,SOTMKTC1,SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
        & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO " & vbCrLf _
        & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE " & vbCrLf _
        & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE " & vbCrLf _
        & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTORDR2.CUST_CODE " & vbCrLf _
        & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
        & "   and SOTORDR2.ORDR_STATUS >= 'O' AND SOTORDR2.ORDR_STATUS <= 'P' " & vbCrLf _
        & "   and SOTORDR2.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where NVL(WHSE_MRP_EXC_IND,'0') <> '1')" & vbCrLf _
        & " group by NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTORDR2.ITEM_CODE"

        Create_TDA(dst.Tables.Add, "SOTORDRM", "**", 0, False, "", 2)
        Fill_Records("SOTORDRM")

        ' put a placeholder record in SOTORDRM for every row in SOTINVHX that is not already represented in SOTORDRM
        For Each row As DataRow In dst.Tables("SOTINVHX").Select("")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim MARKET_CODE As String = row.Item("MARKET_CODE")
            If dst.Tables("SOTORDRM").Rows.Find(New String() {MARKET_CODE, ITEM_CODE}) Is Nothing Then
                dst.Tables("SOTORDRM").Rows.Add(New Object() {MARKET_CODE, ITEM_CODE, 0})
            End If
        Next


        ASCMAIN1.Progress("-", "Vendor Master")
        ASCMAIN1.sql = "Select * from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from ICTITEM1 union Select Distinct VEND_CODE from POTORDR1)"
        Create_TDA(dst.Tables.Add, "APTVEND1", "**", 0, False, "", 1)
        Fill_Records("APTVEND1")

        ASCMAIN1.Progress("-", "Vendor/Item")
        Create_TDA(dst.Tables.Add, "DPTVNDI1", "*", 0, False, "", 2)
        Fill_Records("DPTVNDI1")

        ASCMAIN1.Progress("-", "Forecasts")
        ASCMAIN1.sql = "Select ITEM_CODE, MARKET_CODE, Sum (Decode (OPS_YYYYPP_FC,'000000',FORECAST,0)) FC_PD"
        For i As Integer = 1 To 25
            ASCMAIN1.sql &= ", Sum (Decode (OPS_YYYYPP_FC,'" & p24(i) & "',FORECAST,0)) FC_" & Format$(i, "00")
        Next i
        ASCMAIN1.sql &= " from DPTITMF1 "
        ASCMAIN1.sql &= " where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        ASCMAIN1.sql &= " group by ITEM_CODE, MARKET_CODE"
        Create_TDA(dst.Tables.Add, "DPTITMFX", "**", 0, False, "", 2)
        Fill_Records("DPTITMFX")

        ASCMAIN1.Progress("-", "Misc Code Tables")
        For Each TABLE_NAME As String In New String() {
            "SOTMKTC1", "DPTCRDM1", "DPTEXCM1", "DPTEXCS1", "DPTACTM1", "DPTABCP1"}
            Create_TDA(dst.Tables.Add, TABLE_NAME, "*", 0, False)
            Fill_Records(TABLE_NAME)
        Next

        ASCMAIN1.Progress("-", "Report Work Tables")

        Create_TDA(dst.Tables.Add, "DPTMUPD0", "*")
        With dst.Tables("DPTMUPD0")
            .Columns.Add("QTY", GetType(System.Int64), "ISNULL(QTY_OH,0)+ISNULL(QTY_ORD,0)+ISNULL(QTY_PLN,0)-ISNULL(QTY_REQ,0)")
        End With

        Create_TDA(dst.Tables.Add, "DPTMUPD1", "*")
        Create_TDA(dst.Tables.Add, "DPTMUPD2", "*")
        Create_TDA(dst.Tables.Add, "DPTMUPD3", "*")

        With dst.Tables("DPTMUPD2")
            .Columns.Add("SORT1")
            .Columns.Add("SORT2")
        End With

        Create_Relation("DPTCRDM1", "DPTMUPD1", "CRDM_CODE")
        dst.Tables("DPTMUPD1").Columns.Add("CRDM_SUPPRESS", GetType(System.String), "PARENT.CRDM_SUPPRESS")

        Create_Relation("DPTEXCM1", "DPTMUPD2", "EXC_MSG_CODE")
        dst.Tables("DPTMUPD2").Columns.Add("EXC_MSG_SUPPRESS", GetType(System.String), "PARENT.EXC_MSG_SUPPRESS")

        ASCMAIN1.Progress("-", "Plans")
        Create_TDA(dst.Tables.Add, "DPTPLAN1", "*", 0)

        ASCMAIN1.Progress("-", "Grid")
        Create_TDA(dst.Tables.Add, "DPTMRPG1", "*", 0)
        If use_binding Then Create_BAs("DPTMRPG1")
    End Sub

    Sub Load_Max_Level()
        ASCMAIN1.sql = "Select BM_PROD_ITEM from BMTMAIN1" '  where BM_ISSUE_COUNTER > 0" - PRELIM ISSUES ARE USED TO CREATE COMPONENT COMMITTMENTS FROM SYSTEM GENERATED PLANS IF ITEMS THAT DO NOT YET HAVE STD COSTS
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ITEM_CODE As String = row.Item("BM_PROD_ITEM")
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            If rowICTITEM1 IsNot Nothing Then
                rowICTITEM1.Item("MAX_LEVEL") = 1
            End If
        Next
        ASCMAIN1.sql = "Select Distinct BMTMAIN3.BM_PROD_ITEM from BMTMAIN3,BMTMAIN1 " _
            & " where BMTMAIN1.BM_PROD_ITEM = BMTMAIN3.BM_COMP_ITEM" '  AND BMTMAIN1.BM_ISSUE_COUNTER > 0" - PRELIM ISSUES ARE USED TO CREATE COMPONENT COMMITTMENTS FROM SYSTEM GENERATED PLANS IF ITEMS THAT DO NOT YET HAVE STD COSTS
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ITEM_CODE As String = row.Item("BM_PROD_ITEM")
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            If rowICTITEM1 IsNot Nothing Then
                rowICTITEM1.Item("MAX_LEVEL") = 2
            End If
        Next
    End Sub

    Sub Load_Plan_Waste()
        ASCMAIN1.Progress("Load Plan Waste by Type", "")
        ASCMAIN1.sql = "Update ICTITEM1 set ITEM_PLAN_WASTE_PCT = " _
            & " (Select ITEM_PLAN_WASTE_PCT from ICTTYPE1 where ICTTYPE1.ITEM_TYPE_CODE = ICTITEM1.ITEM_TYPE_CODE) " _
            & " where ITEM_TYPE_CODE is not Null"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Load_MRP_Parms()
        ASCMAIN1.Progress("Load MRP Parameters", "")
        Get_PARM("DPTPARM1")
        DP_PARM_DEF_PLAN_WHSE = ROWs("DPTPARM1").Item("DP_PARM_DEF_PLAN_WHSE")
        GDATE = (Now + ASCMAIN1.NowTSD).Date

        'Dim DP_PARM_PLNR_ANAL_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_PLNR_ANAL_DAYS") & "")
        Dim DP_PARM_XFR_PICK_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_XFR_PICK_DAYS") & "")
        Dim DP_PARM_XFR_XIT_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_XFR_XIT_DAYS") & "")
        'Dim DP_PARM_DEF_VNDR_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_DEF_VNDR_DAYS") & "")
        'Dim DP_PARM_DEF_PO_XIT_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_DEF_PO_XIT_DAYS") & "")
        'Dim DP_PARM_DEF_PO_SCH_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_DEF_PO_SCH_DAYS") & "")
        'Dim DP_PARM_FC_REQ_DAY As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_FC_REQ_DAY") & "")
        'If DP_PARM_FC_REQ_DAY = 0 Then
        '    DP_PARM_FC_REQ_DAY = 1
        'End If
        'Dim DP_PARM_PLAN_GRACE_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_PLAN_GRACE_DAYS") & "")
        Dim DP_PARM_MIN_PLAN_DAYS As Integer = Val(ROWs("DPTPARM1").Item("DP_PARM_MIN_PLAN_DAYS") & "")
        'DP_PARM_DEF_PLAN_WHSE = ROWs("DPTPARM1").Item("DP_PARM_DEF_PLAN_WHSE") & ""
        'Dim DP_PARM_MR_PLAN_FG As String = ROWs("DPTPARM1").Item("DP_PARM_MR_PLAN_FG") & ""
        XDT_1 = Format(GDATE.AddDays(DP_PARM_XFR_XIT_DAYS), "yyyyMMdd")
        XDT_2 = Format(GDATE.AddDays(DP_PARM_XFR_PICK_DAYS + DP_PARM_XFR_XIT_DAYS), "yyyyMMdd")
        DATE_REQUIRED_min = GDATE.AddDays(DP_PARM_MIN_PLAN_DAYS)

        ASCMAIN1.sql = "Update DPTPARM1 Set DP_PARM_MRP_GDATE = :PARM1 where DP_PARM_KEY = 'Z'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "D", New Object() {GDATE})
    End Sub
#End Region

    Function Get_PO_SCH_DAYS(ITEM_PLAN_MAKE_BUY As String, rowDPTVNDI1 As DataRow, rowAPTVEND1 As DataRow) As Int64
        Dim PO_SCH_DAYS As Int64 = 0

        If ITEM_PLAN_MAKE_BUY = "M" Then
            If Val(rowDPTVNDI1.Item("PO_SCH_DAYS") & "") <> 0 Then
                PO_SCH_DAYS = Val(rowDPTVNDI1.Item("PO_SCH_DAYS") & "")
            Else
                If Val(rowAPTVEND1.Item("PO_SCH_DAYS") & "") <> 0 Then
                    PO_SCH_DAYS = Val(rowAPTVEND1.Item("PO_SCH_DAYS") & "")
                Else
                    PO_SCH_DAYS = Val(ROWs("DPTPARM1").Item("DP_PARM_DEF_PO_SCH_DAYS") & "")
                End If
            End If
        End If

        Return PO_SCH_DAYS
    End Function

    Function First_Date_Index(dt() As String, dtz As String) As Integer
        Dim i As Integer = 0
        If dt IsNot Nothing Then
            For i = 1 To UBound(dt)
                If dtz <= dt(i) Or i = UBound(dt) Then
                    Exit For
                End If
            Next
        End If

        Return i
    End Function

    Sub Fix_MIN_MAX_MDS_for_NC()

        ASCMAIN1.Progress("Now Re-Setting No-Charge ABC Parameters")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & "  Declare Cursor C1 is " & vbCrLf _
            & "   Select * from ICTPROD1" & vbCrLf _
            & "    where COST_CATGY_CODE = 'N'" & vbCrLf _
            & "      and NVL(PROD_MIN_DAYS_SUPPLY,0) >= 0" & vbCrLf _
            & "      and NVL(PROD_MAX_POS,0) > 0 AND NVL(PROD_MIN_POS,0) > 0" & vbCrLf _
            & "      and NVL(PROD_MAX_POS,0) > NVL(PROD_MIN_POS,0);" & vbCrLf _
            & "  Begin" & vbCrLf _
            & "    For R1 in C1 Loop" & vbCrLf _
            & "      Update ICTITEM1" & vbCrLf _
            & "       Set ITEM_POS_MAX = R1.PROD_MAX_POS" & vbCrLf _
            & "         , ITEM_POS_MIN = R1.PROD_MIN_POS" & vbCrLf _
            & "         , ITEM_MIN_DAYS_SUPPLY = R1.PROD_MIN_DAYS_SUPPLY" & vbCrLf _
            & "       where PROD_CODE = R1.PROD_CODE and NVL(ITEM_ABC_PARMS_LOCKED,'0') <> '1';" & vbCrLf _
            & "    End Loop;" & vbCrLf _
            & "  End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("")
    End Sub
End Class