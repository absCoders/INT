Public Class GLRDTLA1

#Region "General Declarations"
    Dim GLTDTLA1 As String  ' Transaction Details
    Dim sqlGLTDTLA1 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Dim P As Integer = Val(Mid(ASCMAIN1.CYP, 5, 2))
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 60, -1 * (P - 1))
        'Set_cmbYP_Child("RYP1", (P - 1 + 1), "RYP0")
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        'Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 60, -11)
        'Set_cmbYP_Child("RYP1", 12, "RYP0")

        Absx1.chkFor("CHKSEG2").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & ""
        Absx1.chkFor("CHKSEG3").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & ""
        Absx1.chkFor("CHKSEG4").Text = ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & ""
    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Work File")

        Dim sqlw As String = ""
        'sqlw &= "   and SPTCOOP1.OPS_YYYYPP <= '" & RYP0 & "'"

        Prepare_dst(True, New Object() {""})

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Check_if_Empty("GLTDTLA1")
    End Sub

    Public Overrides Sub Print_Report()
        Create_Pivot()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Dim YP0 As String = Absx1.cmbFor("RYP0").Value & ""
            Dim YP1 As String = Absx1.cmbFor("RYP1").Value & ""

            If YP0 = "" Or YP1 = "" Then
                EMsg &= vbCrLf & "Starting and Ending Periods are Required"
            Else
                If Mid(YP0, 1, 4) <> Mid(YP1, 1, 4) Then
                    EMsg &= vbCrLf & "Starting and Ending Periods must be in Same Year"
                End If
            End If

            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("ACCT_CODE")
            If rowASTDSQLA("CODE_VALUES") & "" = "" Then
                EMsg &= vbCr & "You Must Specify at least 1 Account Code"
            End If

        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SPTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables()

        With dst
            ASCMAIN1.sql = "Select GLTDTLA1.*" & vbCrLf _
                & ", GLTSEGM2.ACCT_SEG_CLASS SEG2_CLASS_CODE" & vbCrLf _
                & ", GLTSEGM3.ACCT_SEG_CLASS SEG3_CLASS_CODE" & vbCrLf _
                & ", GLTSEGM4.ACCT_SEG_CLASS SEG4_CLASS_CODE" & vbCrLf _
                & " from " & GLTDTLA1 & " GLTDTLA1, GLTSEGM1 GLTSEGM2, GLTSEGM1 GLTSEGM3, GLTSEGM1 GLTSEGM4" & vbCrLf _
                & " where GLTSEGM2.ACCT_SEG_ID (+) = '2'" & vbCrLf _
                & "   and GLTSEGM2.ACCT_SEG_CODE (+) = GLTDTLA1.SEG2_CODE" & vbCrLf _
                & "   and GLTSEGM3.ACCT_SEG_ID (+) = '2'" & vbCrLf _
                & "   and GLTSEGM3.ACCT_SEG_CODE (+) = GLTDTLA1.SEG3_CODE" & vbCrLf _
                & "   and GLTSEGM4.ACCT_SEG_ID (+) = '2'" & vbCrLf _
                & "   and GLTSEGM4.ACCT_SEG_CODE (+) = GLTDTLA1.SEG4_CODE" & vbCrLf

            Create_TDA(.Tables.Add("GLTDTLA1"), GLTDTLA1, "**", 0, True, "", 0)

            'ASCMAIN1.sql = "Select * from " & ARTCUST1 & " ARTCUST1"
            'Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)
        End With

        If perform_fill Then
            Fill_Records_RPT(sqlw)
        End If

        Return clsASCBASE1
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        If parms.Length > 0 Then
            sqlw = parms(0)
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        ASCDATA1.ExecuteSQL("Truncate Table " & GLTDTLA1)

        Dim sqlACCT_CODEs As String = SQL_in("ACCT_CODE", "GLTDETL1.ACCT_CODE")
        Dim sqlACCT_CODEs_in As String = Replace(Replace(sqlACCT_CODEs, " And GLTDETL1.ACCT_CODE", ""), " and GLTDETL1.ACCT_CODE", "")

        ASCMAIN1.sql = "Select Distinct GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & " from GLTJRNL1,GLTDETL1" & vbCrLf _
            & " where GLTDETL1.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
            & "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
            & sqlACCT_CODEs

        Dim sqlX As String = ""
        Dim sqlWhereClause As String = "{sql where}" & vbCrLf
        Dim sql_SEG As String = ""

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim JOURNAL_TYPE As String = row.Item("JOURNAL_TYPE")
            ' If ASCMAIN1.Running_in_VS And JOURNAL_TYPE = "OPSJ" Then Stop
            sqlX = GLCMAIN1.Load_Details("9999", JOURNAL_TYPE, "", sql_SEG, sqlWhereClause, sqlACCT_CODEs_in)
            GLCMAIN1.Load_Sub_Details(JOURNAL_TYPE, "9999", "", sqlX, sqlWhereClause, "", "", "")

            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "between '999901' AND '999912'", "between '" & RYP0 & "' AND '" & RYP1 & "'")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "between '999812' AND '999911'", "between '" & ASCMAIN1.Period_Calc(RYP0, -1) & "' AND '" & ASCMAIN1.Period_Calc(RYP1, -1) & "'")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "NULL CODE", "'X' XCODE")

            ASCMAIN1.sql = Replace(ASCMAIN1.sql, " 'GL' CODE2_VALUE,", " 'GL' REASON_CODE,")

            ASCMAIN1.sql = Replace(ASCMAIN1.sql, " CODE1_VALUE,", ",")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, " CODE2_VALUE,", ",")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, " CODE3_VALUE,", ",")

            If ASCMAIN1.sql <> "" Then
                Dim T As String = ASCMAIN1.Temp_Table
                ASCMAIN1.sql = "Select * from " & T & " where ROWNUM < 1"
                Dim tbl As DataTable = ASCDATA1.GetDataTable

                ASCMAIN1.sql = ""
                For Each DC As DataColumn In dst.Tables("GLTDTLA1").Columns
                    Dim C As String = DC.ColumnName

                    '  If JOURNAL_TYPE = "SPCA" And C = "DATE_START" Then Stop

                    If C = "JOURNAL_TYPE" Then
                        ASCMAIN1.sql &= ", '" & JOURNAL_TYPE & "'" & C
                    ElseIf C = "SEG2_CLASS_CODE" Or C = "SEG3_CLASS_CODE" Or C = "SEG4_CLASS_CODE" Then

                    Else
                        If tbl.Columns.Contains(C) Then
                            ASCMAIN1.sql &= ", " & C
                        Else
                            ASCMAIN1.sql &= ", NULL " & C
                        End If
                    End If
                Next

                ASCMAIN1.sql = "Insert into " & GLTDTLA1 & " Select " & Mid(ASCMAIN1.sql, 2) & " from " & T
                ASCDATA1.ExecuteSQL()
            End If
        Next

        EnforceConstraints(False)

        Fill_Records("GLTDTLA1")
        ' Fill_Records("ARTCUST1")

        EnforceConstraints(True)
    End Sub

    Sub Create_Work_Tables()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        'DETL_CVX_NO, DETL_CVX_REF_DATE, DETL_CVX_REF_NO, DETL_CVX_TYPE
        'DETL_CTL_DATE, DETL_CTL_NO, DETL_CTL_TYPE

        sqlGLTDTLA1 = "Select GLTDETL1.ACCT_CODE, GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE" & vbCrLf _
            & ", GLTDETL1.OPS_YYYYPP, GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & ", GLTDETL1.DETL_CVX_REF_NO DETL_TAG" & vbCrLf _
            & ", SOTINVH1.CUST_CODE, APTINVH1.VEND_CODE, ICTCOLL1.COLLECTION_CODE" & vbCrLf _
            & ", SOTINVH1.REASON_CODE, SOTINVH1.MISC_CHG_CODE, SOTINVH1.INV_COMMENT DESC_VALUE" & vbCrLf _
            & ", SOTINVH1.INV_NO, SOTINVH1.ORDR_DEPT INV_NUM, SOTINVH1.INV_DATE, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
            & ", SOTINVH1.EVENT_CODE, SOTINVH1.ORDR_TYPE_CODE, SPTCOOP1.AUTH_NO" & vbCrLf _
            & ", SPTCOOP1.DATE_START, SPTCOOP1.DATE_END, SPTCOOP1.DATE_ACCRUE, SPTCOOP1.OPS_YYYYPP_ACCRUE" & vbCrLf _
            & ", SPTCOOP1.INIT_DATE, SPTCOOP1.INIT_OPER, SPTCOOP1.LAST_DATE, SPTCOOP1.LAST_OPER" & vbCrLf _
            & ", APTINVH1.VOUCHER_NO, APTINVH1.INV_AMT AMT" & vbCrLf _
            & " from SOTINVH1,APTINVH1,ARTCUST1,GLTDETL1,ICTCOLL1,GLTJRNL1,SPTCOOP1 where ROWNUM < 1"

        ASCMAIN1.sql = "Select * from (" & sqlGLTDTLA1 & ") where ROWNUM < 1"
        GLTDTLA1 = ASCMAIN1.Temp_Table()
        '  ASCDATA1.ExecuteSQL("Alter Table " & GLTDTLA1 & " Add Primary Key (CUST_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & GLTDTLA1 & " Modify CUST_CODE NULL")
        ASCDATA1.ExecuteSQL("Alter Table " & GLTDTLA1 & " Modify INV_NO NULL")
        ' ASCDATA1.ExecuteSQL("Alter Table " & GLTDTLA1 & " Modify ACCT_CODE NULL")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Create_Pivot()

        Dim SQLW As String = ""
        SQLW &= SQLA_filter("ACCT_CODE", "GLTDETL1")
        'SQLW &= SQLA_filter("SREP_CODE", "SOTORDR1")

        ASCMAIN1.Progress("Now Creating Workbook")

        Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
        Dim DataTable As DataTable
        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)

        Dim rTotal As Int64
        Dim r0 As Integer = 2
        Dim c0 As Integer = 1

        ASCMAIN1.Progress("-", "Codes")
        ws = wb.Worksheets("Codes")

        ASCMAIN1.sql = "Select ACCT_CODE, ACCT_DESC from GLTACCT1"
        Load_Codes("Accounts", r0, c0, wb)

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from ARTCUST1"
        Load_Codes("Customers", r0, c0, wb)

        ASCMAIN1.sql = "Select VEND_CODE, VEND_NAME from APTVEND1"
        Load_Codes("Vendors", r0, c0, wb)

        ASCMAIN1.sql = "Select REASON_CODE, REASON_DESC, ACCT_CODE from ARTREAS1"
        Load_Codes("Reason Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select MISC_CHG_CODE, MISC_CHG_DESC, ACCT_CODE from SOTMISC1"
        Load_Codes("Misc Charge Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC, MARKET_CODE, SEG3_CODE, CHANNEL_CODE from SOTTCLS1"
        Load_Codes("Trade Class Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME, BRAND_CODE, SEG4_CODE from ICTCOLL1"
        Load_Codes("Collection Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select EXPENSE_TYPE_CODE, EXPENSE_TYPE_DESC, ACCT_CODE from SPTTYPE1"
        Load_Codes("Expense Type Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select ASP_CODE, ASP_DESC, ASP_ACCT_CODE_EXP from SPTACOM0"
        Load_Codes("Commission Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select ORDR_TYPE_CODE, ORDR_TYPE_DESC from SOTTYPE1"
        Load_Codes("Order Type Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select EVENT_CODE, EVENT_DESC, ACCT_CODE, SEG2_CODE from SOTEVNT1"
        Load_Codes("Event Codes", r0, c0, wb)

        ASCMAIN1.sql = "Select * from (" & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, ICTCOLL1.BRAND_CODE, SOTINVHT.OPS_YYYYPP, SOTINVHT.ACCT_CODE_SLS" & vbCrLf _
            & ", SUM (SOTINVHT.SLS) SALES" & vbCrLf _
            & " from SOTINVHT,SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where SOTINVHT.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
            & "   and SOTINVHT.INV_TYPE= 'I'" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = SOTINVHT.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = SOTINVHT.INV_NO" & vbCrLf _
            & "   and SOTINVH2.INV_LNO = SOTINVHT.INV_LNO" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTINVHT.SLS <> 0" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, ICTCOLL1.BRAND_CODE, SOTINVHT.OPS_YYYYPP, SOTINVHT.ACCT_CODE_SLS" & vbCrLf _
            & ") order by CUST_CODE, BRAND_CODE, OPS_YYYYPP, ACCT_CODE_SLS"

        c0 = 1
        Load_Codes("Sales", r0, c0, wb, "Sales")

        Dim sqlACCT_CODEs As String = SQL_in("ACCT_CODE", "GLTDETL1.ACCT_CODE")

        ASCMAIN1.sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, OPS_YYYYPP, JOURNAL_TYPE, SUM (AMT) SUB_GL, SUM (DETL_POSTING_AMT) GL, SUM (AMT - DETL_POSTING_AMT) DIFF from (" & vbCrLf _
            & "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, OPS_YYYYPP, JOURNAL_TYPE, SUM (AMT) AMT, 0 DETL_POSTING_AMT from " & GLTDTLA1 & vbCrLf _
            & " group by ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, OPS_YYYYPP, JOURNAL_TYPE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select GLTDETL1.ACCT_CODE, GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE, GLTDETL1.OPS_YYYYPP, GLTJRNL1.JOURNAL_TYPE, 0 AMT, SUM (GLTDETL1.DETL_POSTING_AMT) DIST_AMT " & vbCrLf _
            & " from GLTDETL1,GLTJRNL1" & vbCrLf _
            & " where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
            & "   and GLTDETL1.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf _
            & sqlACCT_CODEs & vbCrLf _
            & " group by GLTDETL1.ACCT_CODE, GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE, GLTDETL1.OPS_YYYYPP, GLTJRNL1.JOURNAL_TYPE" & vbCrLf _
            & ") group by ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, OPS_YYYYPP, JOURNAL_TYPE"

        c0 = 1
        Dim TBL As DataTable = Load_Codes("Balancing", r0, c0, wb, "Balancing")
        For Each row As DataRow In TBL.Select("DIFF<>0")
            Dim rowGLTDTLA1 As DataRow = dst.Tables("GLTDTLA1").NewRow
            With rowGLTDTLA1
                .Item("ACCT_CODE") = row.Item("ACCT_CODE")
                .Item("SEG2_CODE") = row.Item("SEG2_CODE")
                .Item("SEG3_CODE") = row.Item("SEG3_CODE")
                .Item("SEG4_CODE") = row.Item("SEG4_CODE")
                .Item("OPS_YYYYPP") = row.Item("OPS_YYYYPP")
                .Item("JOURNAL_TYPE") = row.Item("JOURNAL_TYPE")
                .Item("DESC_VALUE") = "Balancing"
                .Item("AMT") = -1 * Val(row.Item("DIFF") & "")
            End With
            dst.Tables("GLTDTLA1").Rows.Add(rowGLTDTLA1)
        Next
        Update_Record_TDA("GLTDTLA1")

        ASCMAIN1.Progress("-", "Data")

        ws = wb.Worksheets("Data")

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  GLTDTLA1.*" & vbCrLf _
            & ", GLTSEGM2.ACCT_SEG_CLASS SEG2_CLASS_CODE" & vbCrLf _
            & ", GLTSEGM3.ACCT_SEG_CLASS SEG3_CLASS_CODE" & vbCrLf _
            & ", GLTSEGM4.ACCT_SEG_CLASS SEG4_CLASS_CODE" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE, SOTTCLS1.CHANNEL_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " from " & GLTDTLA1 & " GLTDTLA1, GLTSEGM1 GLTSEGM2, GLTSEGM1 GLTSEGM3, GLTSEGM1 GLTSEGM4" & vbCrLf _
            & ", ARTCUST1, ICTCOLL1, SOTTCLS1" & vbCrLf _
            & " where GLTSEGM2.ACCT_SEG_ID (+) = '2'" & vbCrLf _
            & "   and GLTSEGM2.ACCT_SEG_CODE (+) = GLTDTLA1.SEG2_CODE" & vbCrLf _
            & "   and GLTSEGM3.ACCT_SEG_ID (+) = '3'" & vbCrLf _
            & "   and GLTSEGM3.ACCT_SEG_CODE (+) = GLTDTLA1.SEG3_CODE" & vbCrLf _
            & "   and GLTSEGM4.ACCT_SEG_ID (+) = '4'" & vbCrLf _
            & "   and GLTSEGM4.ACCT_SEG_CODE (+) = GLTDTLA1.SEG4_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = GLTDTLA1.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = GLTDTLA1.COLLECTION_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE"

        ASCMAIN1.sql = "Select X.*, GLTACCT1.ACCT_DESC, ARTCUST1.CUST_NAME, APTVEND1.VEND_NAME" & vbCrLf _
            & ", SOTTCLS1.TRADE_CLASS_DESC, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", ARTREAS1.REASON_DESC, SOTMISC1.MISC_CHG_DESC" & vbCrLf _
            & " from (" & ASCMAIN1.sql & ") X,GLTACCT1,ARTCUST1,APTVEND1,SOTTCLS1,ICTCOLL1,ARTREAS1,SOTMISC1" & vbCrLf _
            & " where GLTACCT1.ACCT_CODE (+) = X.ACCT_CODE" & vbCrLf _
            & "   And ARTCUST1.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
            & "   And APTVEND1.VEND_CODE (+) = X.VEND_CODE" & vbCrLf _
            & "   And SOTTCLS1.TRADE_CLASS_CODE (+) = X.TRADE_CLASS_CODE" & vbCrLf _
            & "   And ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE" & vbCrLf _
            & "   And ARTREAS1.REASON_CODE (+) = X.REASON_CODE" & vbCrLf _
            & "   And SOTMISC1.MISC_CHG_CODE (+) = X.MISC_CHG_CODE"

        DataTable = ASCDATA1.GetDataTable

        r0 = 2
        c0 = 5

        Dim c1 As String = Excel_Cell(0, c0)
        Dim c2 As String = Excel_Cell(0, c0 + DataTable.Columns.Count - 1)

        Headings(DataTable, r0, c0, ws)

        rTotal = DataTable.Select("").Length
        r = 0
        For Each row As DataRow In DataTable.Select("", "ACCT_CODE,OPS_YYYYPP,CUST_CODE,VEND_CODE")
            r += 1
            ' ws.Range("B" & CStr(r0 + r)).Value2 = row.ItemArray
            ws.Range(c1 & CStr(r0 + r) & ":" & c2 & CStr(r0 + r)).Value2 = row.ItemArray

            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        wb.Names.Add("PivotBase", "=Data!$A$" & CStr(r0) & ":$" & c2 & "$" & CStr(r0 + DataTable.Rows.Count))
        ws.Columns(c1 & ":" & c2).AutoFit()

        'wb.Names.Add("KEY", "=MAINDATA!$A$" & CStr(r0) & ":$A$" & CStr(r0 + DataTable.Rows.Count))

        ' ws.Cells(r0 - 2, 15).Value = "X"

        ASCMAIN1.Progress("-", "Formulas")

        xlSourceRange = ws.Range("A" & CStr(r0 + 1) & ":D" & CStr(r0 + 1))
        xlDestRange = ws.Range("A" & CStr(r0 + 1) & ":D" & CStr(r0 + DataTable.Rows.Count))
        xlSourceRange.Copy(xlDestRange)


        With ws.Range("A" & CStr(r0) & ":" & c2 & CStr(r0))
            .Interior.Color = System.Drawing.Color.LightBlue
            .Font.Color = System.Drawing.Color.Black
            .Font.Bold = True
        End With


        ws.Cells(1, 1).Value = Now

        ASCMAIN1.Progress("-", "Pivots")
        '   excel.Run("ResetData")

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "GL_Detail_Analysis"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"

                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(ASCMAIN1.Folders("Work") & XLS_FILENAME _
                          , Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                wb.Close(False, objOpt, objOpt)

                success = True

            Catch ex As Exception
                ' Stop
            End Try
        Loop

        excel.Quit()
        ws = Nothing
        wb = Nothing
        excel = Nothing
        xlSourceRange = Nothing
        xlDestRange = Nothing

        ReleaseCOMObject(xlDestRange)
        ReleaseCOMObject(xlSourceRange)
        ReleaseCOMObject(ws)
        ReleaseCOMObject(wb)
        ReleaseCOMObject(excel)

        Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

        ASCMAIN1.Progress("")

    End Sub

    Function Load_Codes(DataTableName As String, r0 As Integer, ByRef c0 As Integer,
                   wb As Microsoft.Office.Interop.Excel.Workbook, Optional Sheet_Name As String = "Codes") As DataTable

        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets(Sheet_Name)

        Dim DataTable As DataTable = ASCDATA1.GetDataTable


        Dim c1 As String = Excel_Cell(0, c0)
        Dim c2 As String = Excel_Cell(0, c0 + DataTable.Columns.Count - 1)

        Headings(DataTable, r0, c0, ws)

        Dim rTotal As Int64 = DataTable.Select("").Length
        Dim r As Int64 = 0

        For Each row As DataRow In DataTable.Select("", DataTable.Columns(0).ColumnName)
            r += 1
            ws.Range(c1 & CStr(r0 + r) & ":" & c2 & CStr(r0 + r)).Value2 = row.ItemArray

            If r Mod 100 = 0 Then
                ASCMAIN1.Progress("-", CStr(r) & "/" & CStr(rTotal))
            End If
        Next
        'ws.Cells(c1 & CStr(r0 - 1)).Value = DataTableName
        ws.Cells(r0 - 1, c0).Value = DataTableName
        ws.Cells(r0 - 1, c0).Font.Color = System.Drawing.Color.Blue
        ' ws.Range(c1 & ":" & c2).AutoFit()
        ws.Columns(c1 & ":" & c2).AutoFit()
        If c0 <> 1 Then
            'ws.Columns(c0).width = 2
            Dim C As String = Excel_Cell(0, c0 - 1)
            ws.Range(C & ":" & C).EntireColumn.ColumnWidth = 2
        End If

        wb.Names.Add(Replace(DataTableName, " ", "_"), "=" & Sheet_Name & "!$" & c1 & "$" & CStr(r0) & ":$" & c2 & "$" & CStr(r0 + DataTable.Rows.Count))

        c0 += DataTable.Columns.Count + 1

        Return DataTable
    End Function

    Sub Headings(DataTable As DataTable, r0 As Integer, c0 As Integer, ws As Microsoft.Office.Interop.Excel.Worksheet)

        Dim COLUMN_NAMEs As New List(Of String)
        For C As Integer = 0 To DataTable.Columns.Count - 1
            COLUMN_NAMEs.Add(DataTable.Columns(C).ColumnName)
            Dim F As String = "@"
            Dim T As String = DataTable.Columns(C).DataType.ToString
            T = Replace(T, "System.", "")
            Dim COL As String = Excel_Cell(0, c0 + C)
            With ws.Range(COL & ":" & COL)
                If T.StartsWith("Date") Then
                    F = "MM/dd/yyyy"
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                ElseIf T.StartsWith("Int") Then
                    F = "#,##0"
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                ElseIf T.StartsWith("String") Then
                Else
                    F = "#,##0.00"
                    .HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                End If
                .EntireColumn.NumberFormat = F '
            End With
        Next

        Dim c1 As String = Excel_Cell(0, c0)
        Dim c2 As String = Excel_Cell(0, c0 + DataTable.Columns.Count - 1)

        With ws.Range(c1 & CStr(r0 + 0) & ":" & c2 & CStr(r0 + 0))
            .Value2 = COLUMN_NAMEs.ToArray
            .Interior.Color = System.Drawing.Color.CornflowerBlue
            .Font.Color = System.Drawing.Color.White
            .Font.Bold = True
            .Font.Size = 9
            .Font.Name = "Calibri"
        End With

    End Sub
End Class