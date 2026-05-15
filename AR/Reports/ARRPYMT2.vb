Imports System.Math

Public Class ARRPYMT2

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date

    Dim ARTPYMT1 As String
    Dim ARTPYMT2 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("ARTPARM1")
        Call Get_PARM("GLTPARM1")

        Absx1.optFor("RANGE").CheckedIndex = 2

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left

        ' optARONLY.Visible = (ASCMAIN1.CLIENT = "VAN")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        dst.EnforceConstraints = False

        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "N" Then
            sqlw = " and ARTPYMT1.STATUS = '1'"
        ElseIf Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Payment Receipts Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Payment Receipts Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = " and ARTPYMT1.PYMT_BATCH_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            'sqlw = " and ARTPYMT1.REGISTER_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Payment Receipts Posted in " & xRYP0_legend
            Else
                SUBT = "Payment Receipts Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw = " and ARTPYMT1.OPS_YYYYPP between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "F" Then
            SUBT = ""
            RWU = "N"
            sqlw = ""
        End If

        sqlw &= Get_Filter("BANK_CODE", "ARTPYMT1.BANK_CODE")
        sqlw &= Get_Filter("PYMT_BATCH_NO", "ARTPYMT1.PYMT_BATCH_NO")

        Dim sqlPYMT_BATCH_LNO As String = ""
        Dim sqlCUST_CODE As String = Get_Filter("CUST_CODE", "ARTPYMT2.CUST_CODE")
        If sqlCUST_CODE <> "" Or sqlPYMT_BATCH_LNO <> "" Then
            sqlw &= " and ARTPYMT1.PYMT_BATCH_NO in (Select Distinct PYMT_BATCH_NO from ARTPYMT2 " & ASCMAIN1.SQL_Add_WHERE(sqlCUST_CODE & sqlPYMT_BATCH_LNO) & ")"
        End If

        ASCMAIN1.sql = "Select ARTPYMT1.*, 0 PYMTS, 0 PYMTS_APPLIED" _
        & " from ARTPYMT1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
        ARTPYMT1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMT1 & " Add Primary Key (PYMT_BATCH_NO)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMT1 & " Add PYMT_TOTAL NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMT1 & " Add PYMT_TOTAL_CURR NUMBER (13,2)")

        ASCMAIN1.sql = "Select ARTPYMT2.*" _
        & " from " & ARTPYMT1 & " ARTPYMT1,ARTPYMT2 " _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & IIf(Absx1.optFor("RANGE").Value = "N", "", "   and ARTPYMT2.PYMT_STATUS = '2'") _
        & sqlCUST_CODE & sqlPYMT_BATCH_LNO

        If Absx1.optFor("RANGE").Value = "N" Then
        Else
            If optARONLY.Value = "AR" Then
                ASCMAIN1.sql &= " and ARTPYMT2.CUST_CODE is NOT NULL"
            ElseIf optARONLY.Value = "NON" Then
                ASCMAIN1.sql &= " and ARTPYMT2.CUST_CODE IS NULL"
            End If
        End If

        ARTPYMT2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMT2 & " Add Primary Key (PYMT_BATCH_NO,PYMT_BATCH_LNO)")

        ASCMAIN1.sql = "" _
        & " Begin " _
        & "  Declare Cursor C1 is " _
        & "   Select ARTPYMT2.PYMT_BATCH_NO" _
        & "    , Sum (ARTPYMT2.CUST_PYMT_AMT) PYMT_TOTAL" _
        & "    , Sum (ARTPYMT2.CUST_PYMT_AMT_CURR) PYMT_TOTAL_CURR" _
        & "    , Count (*) PYMTS" _
        & "    , SUM (DECODE(ARTPYMT2.PYMT_STATUS,'2',1,0)) PYMTS_APPLIED" _
        & "    from " & ARTPYMT2 & " ARTPYMT2," & ARTPYMT1 & " ARTPYMT1" _
        & "    where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "      and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & "    group by ARTPYMT2.PYMT_BATCH_NO;" _
        & "  Begin " _
        & "   For R1 in C1 Loop" _
        & "    Update " & ARTPYMT1 _
        & "     Set PYMT_TOTAL = R1.PYMT_TOTAL, PYMT_TOTAL_CURR = R1.PYMT_TOTAL_CURR, " _
        & "         PYMTS = R1.PYMTS, PYMTS_APPLIED = R1.PYMTS_APPLIED" _
        & "     where PYMT_BATCH_NO = R1.PYMT_BATCH_NO;" _
        & "   End Loop; " _
        & "  End; " _
        & " End;"
        ASCDATA1.ExecuteSQL()



        If Absx1.optFor("RANGE").Value = "N" Then
            ASCMAIN1.sql = "Delete from " & ARTPYMT2 & " where PYMT_BATCH_NO in (Select PYMT_BATCH_NO from " & ARTPYMT1 & " where PYMTS <> PYMTS_APPLIED)"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Delete from " & ARTPYMT1 & " where PYMTS <> PYMTS_APPLIED"
            ASCDATA1.ExecuteSQL()
        End If


        ASCMAIN1.sql = "Select ARTPYMT1.*,GLTBANK1.BANK_DESC " _
        & " from " & ARTPYMT1 & " ARTPYMT1,GLTBANK1 " _
        & " where GLTBANK1.BANK_CODE (+) = ARTPYMT1.BANK_CODE " _
        & sqlw
        Create_TDA(dst.Tables.Add, "ARTPYMT1", "**", 0)
        Fill_Records("ARTPYMT1")

        ASCMAIN1.sql = "Select ARTPYMT2.*" _
        & " from ARTPYMT2," & ARTPYMT1 & " ARTPYMT1" _
        & " where ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & sqlCUST_CODE & sqlPYMT_BATCH_LNO
        Create_TDA(dst.Tables.Add, "ARTPYMT2", "**", 0)
        Fill_Records("ARTPYMT2")

        ASCMAIN1.sql = "Select ARTPYMT3.*, 'A' RECORD_TYPE, 'X' INV_COMMENT, 'X' INV_COMMENT2, 'X' CHARGEBACK_IND, 'X' LINE_DESC" _
        & " from ARTPYMT3," & ARTPYMT2 & " ARTPYMT2," & ARTPYMT1 & " ARTPYMT1" _
        & " where ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO"
        Dim ARTPYMTX As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMTX & " Modify INV_COMMENT VARCHAR2(1000)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMTX & " Modify INV_COMMENT2 VARCHAR2(1000)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTPYMTX & " Modify LINE_DESC VARCHAR2(100)")

        ASCMAIN1.sql = "Insert into " & ARTPYMTX _
        & " (PYMT_BATCH_NO,PYMT_BATCH_LNO,PYMT_BATCH_ILNO,RECORD_TYPE)" _
        & " Select ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO, 0, 'X' RECORD_TYPE" _
        & " from " & ARTPYMT2 & " ARTPYMT2," & ARTPYMT1 & " ARTPYMT1" _
        & " where ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & "   and ARTPYMT2.PYMT_DELETED = '1'"
        ASCDATA1.ExecuteSQL() ' DELETED PAYMENTS

        ASCMAIN1.sql = "Insert into " & ARTPYMTX _
        & " (PYMT_BATCH_NO,PYMT_BATCH_LNO,PYMT_BATCH_ILNO," _
        & "  POST_CODE,INV_CUST_PO,INV_PMT,INV_COMMENT,CHARGEBACK_IND," _
        & "  SEG2_CODE,SEG3_CODE,SEG4_CODE,INV_PMT_CURR, RECORD_TYPE, LINE_DESC)" _
        & " Select ARTPYMT4.*, 'G' RECORD_TYPE, GLTACCT1.ACCT_DESC" _
        & " from ARTPYMT4," & ARTPYMT2 & " ARTPYMT2," & ARTPYMT1 & " ARTPYMT1,GLTACCT1" _
        & " where ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and GLTACCT1.ACCT_CODE (+) = ARTPYMT4.ACCT_CODE"
        ASCDATA1.ExecuteSQL() ' GL DISTRIBUTIONS

        ASCMAIN1.sql = "Insert into " & ARTPYMTX _
        & " (PYMT_BATCH_NO,PYMT_BATCH_LNO,PYMT_BATCH_ILNO," _
        & "  REASON_CODE,POST_CODE,INV_PMT,INV_COMMENT," _
        & "  CHARGEBACK_IND,INV_NUM,INV_CUST_PO,CUST_CODE_SO," _
        & "  SEG2_CODE,SEG3_CODE,SEG4_CODE," _
        & "  INV_TYPE,INV_COMMENT2,INV_PMT_CURR,RECORD_TYPE, LINE_DESC)" _
        & " Select ARTPYMT5.*, 'D' RECORD_TYPE, ARTREAS1.REASON_DESC" _
        & " from ARTPYMT5," & ARTPYMT2 & " ARTPYMT2," & ARTPYMT1 & " ARTPYMT1, ARTREAS1" _
        & " where ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" _
        & "   and ARTREAS1.REASON_CODE (+) = ARTPYMT5.REASON_CODE"
        ASCDATA1.ExecuteSQL() ' DEDUCTIONS ALLOWED

        ASCMAIN1.sql = "Insert into " & ARTPYMTX _
        & " (PYMT_BATCH_NO,PYMT_BATCH_LNO,PYMT_BATCH_ILNO," _
        & "  REASON_CODE,POST_CODE,INV_PMT,INV_COMMENT," _
        & "  CHARGEBACK_IND,INV_NUM,INV_CUST_PO,CUST_CODE_SO," _
        & "  SEG2_CODE,SEG3_CODE,SEG4_CODE," _
        & "  INV_TYPE,INV_COMMENT2,INV_PMT_CURR,RECORD_TYPE)" _
        & " Select ARTPYMT5.*, 'B' RECORD_TYPE" _
        & " from ARTPYMT5," & ARTPYMT2 & " ARTPYMT2," & ARTPYMT1 & " ARTPYMT1" _
        & " where ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '1'" _
        & "   and INV_TYPE_CB = 'B'"
        ASCDATA1.ExecuteSQL() ' DEDUCTIONS CHARGED BACK

        ASCMAIN1.sql = "Insert into " & ARTPYMTX _
        & " (PYMT_BATCH_NO,PYMT_BATCH_LNO,PYMT_BATCH_ILNO," _
        & "  REASON_CODE,POST_CODE,INV_PMT,INV_COMMENT," _
        & "  CHARGEBACK_IND,INV_NUM,INV_CUST_PO,CUST_CODE_SO," _
        & "  SEG2_CODE,SEG3_CODE,SEG4_CODE," _
        & "  INV_TYPE,INV_COMMENT2,INV_PMT_CURR,RECORD_TYPE)" _
        & " Select ARTPYMT5.*, 'C' RECORD_TYPE" _
        & " from ARTPYMT5," & ARTPYMT2 & " ARTPYMT2," & ARTPYMT1 & " ARTPYMT1" _
        & " where ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '1'" _
        & "   and INV_TYPE_CB = 'O'"
        ASCDATA1.ExecuteSQL() ' OVER PAYMENTS KEPT ON ACCOUNT

        ASCMAIN1.sql = "Update " & ARTPYMTX & " Set INV_PMT = -1 * INV_PMT where RECORD_TYPE in ('C','B')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ARTPYMTX & " Set INV_PMT = -1 * INV_PMT where RECORD_TYPE in ('D')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update " & ARTPYMTX & " Set INV_PMT = -1 * INV_PMT where RECORD_TYPE in ('G')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & ARTPYMTX & " ARTPYMTX Set POST_CODE = (SELECT ACCT_CODE FROM ARTREAS1 WHERE REASON_CODE = ARTPYMTX.REASON_CODE) where RECORD_TYPE = 'D'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select ARTPYMTX.* from " & ARTPYMTX & " ARTPYMTX"
        Create_TDA(dst.Tables.Add, "ARTPYMTX", "**", 0, False)
        With dst.Tables("ARTPYMTX")
            .PrimaryKey = New DataColumn() _
                {.Columns("PYMT_BATCH_NO"), .Columns("PYMT_BATCH_LNO"), .Columns("RECORD_TYPE"), .Columns("PYMT_BATCH_ILNO")}
        End With
        Fill_Records("ARTPYMTX")

        dst.EnforceConstraints = True

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        Prepare_GL_Interface("ARCR")

        Check_if_Empty("ARTPYMT1")
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("GL_PARM_CURR_CODE", ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE"))
        CR_params.Add("SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
        CR_params.Add("SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
        CR_params.Add("SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")
        CR_params.Add("DISC_HDG", ROWs("ARTPARM1").Item("AR_PARM_HDG_DISC") & "")
        CR_params.Add("WOFF_HDG", ROWs("ARTPARM1").Item("AR_PARM_HDG_WOFF") & "")
        Generate_Report(RPT, , SUBT)
        Call Print_GL()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        optARONLY.Visible = Not (optRANGE.Value = "N") And (ASCMAIN1.CLIENT = "VAN")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()

        Dim sql As String = "Update ARTPYMT1 " _
        & " Set STATUS = '2', REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2, REGISTER_DATE = :PARM3" _
        & " where PYMT_BATCH_NO in (Select PYMT_BATCH_NO from " & ARTPYMT1 & " )"
        ASCDATA1.ExecuteSQL(sql, "VVD", New Object() {"1", MyBase.XNO, DATETIME_STAMP.Date})

        Call GL_Update()
    End Sub

    Sub Prepare_GL_Interface(ByVal JOURNAL_TYPE As String)

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Double
        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))
        Dim DETL_CTL_NO As String = ""

        Dim BY_PYMT_BATCH_NO As String = ", ARTPYMT1.PYMT_BATCH_NO"
        'If ASCMAIN1.CLIENT = "INT" Then
        '    BY_PYMT_BATCH_NO = ", ARTPYMT1.PYMT_BATCH_NO"
        'Else
        '    BY_PYMT_BATCH_NO = ", NULL PYMT_BATCH_NO"
        '    ' WHAT IS GOOD FOR INT MAY BE GOOD FOR AHA, BUT NOT SURE UNTIL WE VERIFY
        'End If

        For Each DIST_TYPE As String In New String() {"CASH", "AR", "DISC", "WOFF", "DED", "CB", "OA", "GL", "FX"}

            Select Case DIST_TYPE
                Case "CASH"
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", GLTBANK1.ACCT_CODE, GLTBANK1.SEG2_CODE, GLTBANK1.SEG3_CODE, GLTBANK1.SEG4_CODE"
                    ASCMAIN1.sql = "Select " & sql1 _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, SUM (ARTPYMT2.CUST_PYMT_AMT) DIST_AMT " _
                    & " FROM " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, GLTBANK1" _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT1.BANK_CODE = GLTBANK1.BANK_CODE" _
                    & "   and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
                    & " GROUP BY " & sql1 _
                    & " ORDER BY " & sql1

                Case "AR"
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", ARTPOST1.ACCT_CODE, ARTPOST1.SEG2_CODE, ARTPOST1.SEG3_CODE, ARTPOST1.SEG4_CODE"
                    ASCMAIN1.sql = "Select " & sql1 _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, -1 * SUM (NVL(ARTPYMT3.INV_PMT,0)+NVL(ARTPYMT3.INV_DISC_TAKEN,0)+NVL(ARTPYMT3.INV_WRITE_OFF,0)) DIST_AMT " _
                    & " FROM " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, ARTPYMT3, ARTPOST1" _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
                    & "   and ARTPOST1.POST_CODE (+) = ARTPYMT3.POST_CODE" _
                    & "   and NVL(ARTPYMT3.INV_PMT,0)+NVL(ARTPYMT3.INV_DISC_TAKEN,0)+NVL(ARTPYMT3.INV_WRITE_OFF,0) <> 0" _
                    & " GROUP BY " & sql1 _
                    & " ORDER BY " & sql1

                Case "FX"
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", GLTPARM1.GL_PARM_CURR_GAIN_LOSS ACCT_CODE" _
                                         & ", NVL(SOTCHAN1.SEG2_CODE,GLTPARM1.GL_PARM_DEF_SEG2) SEG2_CODE" _
                                         & ", NVL(NVL(SOTTCLS1.SEG3_CODE,ARTCUST1.TRADE_CLASS_CODE),GLTPARM1.GL_PARM_DEF_SEG3) SEG3_CODE" _
                                         & ", GLTPARM1.GL_PARM_DEF_SEG4 SEG4_CODE"
                    Dim SQL2 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", GLTPARM1.GL_PARM_CURR_GAIN_LOSS" _
                                         & ", NVL(SOTCHAN1.SEG2_CODE,GLTPARM1.GL_PARM_DEF_SEG2)" _
                                         & ", NVL(NVL(SOTTCLS1.SEG3_CODE,ARTCUST1.TRADE_CLASS_CODE),GLTPARM1.GL_PARM_DEF_SEG3)" _
                                         & ", GLTPARM1.GL_PARM_DEF_SEG4"

                    'If GL_by_PYMT_BATCH Then sql1 &= ", ARTPYMT1.PYMT_BATCH_NO"
                    'If GL_by_PYMT_BATCH Then SQL2 &= ", ARTPYMT1.PYMT_BATCH_NO"

                    ASCMAIN1.sql = "Select " & sql1 & vbCrLf _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, -1 * SUM (NVL(ARTPYMT3.CURR_GAIN_LOSS,0)) DIST_AMT " & vbCrLf _
                    & " from " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, ARTPYMT3, GLTPARM1, ARTCUST1, SOTTCLS1, SOTCHAN1" & vbCrLf _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                    & "   and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
                    & "   and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
                    & "   and GLTPARM1.GL_PARM_KEY ='Z'" & vbCrLf _
                    & "   and NVL(ARTPYMT3.CURR_GAIN_LOSS,0) <> 0" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = ARTPYMT2.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTCHAN1.CHANNEL_CODE (+) = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                    & " group by " & SQL2 & vbCrLf _
                    & " order by " & SQL2

                Case "DISC"
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", ARTREAS1.ACCT_CODE, SOTCHAN1.SEG2_CODE, ARTREAS1.SEG3_CODE, ARTREAS1.SEG4_CODE"
                    ASCMAIN1.sql = "Select " & sql1 _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, SUM (ARTPYMT3.INV_DISC_TAKEN) DIST_AMT " _
                    & " FROM " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, ARTPYMT3, ARTREAS1, ARTCUST1, SOTTCLS1, SOTCHAN1" _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
                    & "   and ARTCUST1.CUST_CODE(+) = ARTPYMT2.CUST_CODE" _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE(+) = ARTCUST1.TRADE_CLASS_CODE" _
                    & "   and SOTCHAN1.CHANNEL_CODE(+) = SOTTCLS1.CHANNEL_CODE" _
                    & "   and ARTREAS1.REASON_CODE = '" & ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_DISC") & "'" _
                    & "   and NVL(ARTPYMT3.INV_DISC_TAKEN,0) <> 0" _
                    & " GROUP BY " & sql1 _
                    & " ORDER BY " & sql1

                Case "WOFF"
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", ARTREAS1.ACCT_CODE, SOTCHAN1.SEG2_CODE, ARTREAS1.SEG3_CODE, ARTREAS1.SEG4_CODE"
                    ASCMAIN1.sql = "Select " & sql1 _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, SUM (ARTPYMT3.INV_WRITE_OFF) DIST_AMT " _
                    & " FROM " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, ARTPYMT3, ARTREAS1, ARTCUST1, SOTTCLS1, SOTCHAN1" _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
                    & "   and ARTCUST1.CUST_CODE(+) = ARTPYMT2.CUST_CODE" _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE(+) = ARTCUST1.TRADE_CLASS_CODE" _
                    & "   and SOTCHAN1.CHANNEL_CODE(+) = SOTTCLS1.CHANNEL_CODE" _
                    & "   and ARTREAS1.REASON_CODE = '" & ROWs("ARTPARM1").Item("AR_PARM_REASON_CODE_WOFF") & "'" _
                    & "   and NVL(ARTPYMT3.INV_WRITE_OFF,0) <> 0" _
                    & " GROUP BY " & sql1 _
                    & " ORDER BY " & sql1

                Case "DED"
                    'Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO _
                    '                     & ", ARTREAS1.ACCT_CODE, ARTREAS1.SEG2_CODE, ARTREAS1.SEG3_CODE, ARTREAS1.SEG4_CODE"
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & vbCrLf _
                            & ", ARTREAS1.ACCT_CODE" & vbCrLf _
                            & ", ARTPYMT5.SEG2_CODE, ARTPYMT5.SEG3_CODE, ARTPYMT5.SEG4_CODE"

                    ASCMAIN1.sql = "Select " & sql1 _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, SUM (ARTPYMT5.GL_DIST_AMT) DIST_AMT " _
                    & " FROM " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, ARTPYMT5, ARTREAS1" _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
                    & "   and ARTREAS1.REASON_CODE = ARTPYMT5.REASON_CODE" _
                    & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" _
                    & " GROUP BY " & sql1 _
                    & " ORDER BY " & sql1

                Case "CB"
                    Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_CB"))
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", ARTPOST1.ACCT_CODE, ARTPOST1.SEG2_CODE, ARTPOST1.SEG3_CODE, ARTPOST1.SEG4_CODE"
                    ASCMAIN1.sql = "Select " & sql1 _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, SUM (ARTPYMT5.GL_DIST_AMT) DIST_AMT " _
                    & " FROM " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, ARTPYMT5, ARTPOST1" _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
                    & "   and ARTPOST1.POST_CODE = '" & rowSOTTYPE1.Item("POST_CODE") & "'" _
                    & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '1'" _
                    & "   and NVL(ARTPYMT5.INV_TYPE_CB,'?') = 'B'" _
                    & " GROUP BY " & sql1 _
                    & " ORDER BY " & sql1

                Case "OA"
                    Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ROWs("ARTPARM1").Item("AR_PARM_ORDR_TYPE_OA"))
                    Dim sql1 As String = "ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE" & BY_PYMT_BATCH_NO & ", ARTPOST1.ACCT_CODE, ARTPOST1.SEG2_CODE, ARTPOST1.SEG3_CODE, ARTPOST1.SEG4_CODE"
                    ASCMAIN1.sql = "Select " & sql1 _
                    & ", '" & DIST_TYPE & "' DIST_TYPE, SUM (ARTPYMT5.GL_DIST_AMT) DIST_AMT " _
                    & " FROM " & ARTPYMT1 & " ARTPYMT1, " & ARTPYMT2 & " ARTPYMT2, ARTPYMT5, ARTPOST1" _
                    & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
                    & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
                    & "   and ARTPOST1.POST_CODE = '" & rowSOTTYPE1.Item("POST_CODE") & "'" _
                    & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '1'" _
                    & "   and NVL(ARTPYMT5.INV_TYPE_CB,'?') = 'O'" _
                    & " GROUP BY " & sql1 _
                    & " ORDER BY " & sql1

                Case "GL"
                    Dim sql1 As String = $"ARTPYMT1.OPS_YYYYPP, ARTPYMT1.PYMT_BATCH_DATE {BY_PYMT_BATCH_NO}
                        , ARTPYMT4.ACCT_CODE, ARTPYMT4.SEG2_CODE, ARTPYMT4.SEG3_CODE, ARTPYMT4.SEG4_CODE"
                    Dim sql1X As String = "
                        , CASE WHEN ARTPYMT2.CUST_CODE IS NULL THEN ARTPYMT2.CUST_NAME ELSE NULL END CUST_NAME
                        , CASE WHEN ARTPYMT2.CUST_CODE IS NULL THEN ARTPYMT2.PYMT_NOTE ELSE NULL END PYMT_NOTE"
                    Dim sql1XG As String = "
                        , CASE WHEN ARTPYMT2.CUST_CODE IS NULL THEN ARTPYMT2.CUST_NAME ELSE NULL END
                        , CASE WHEN ARTPYMT2.CUST_CODE IS NULL THEN ARTPYMT2.PYMT_NOTE ELSE NULL END"
                    ASCMAIN1.sql = $"Select {sql1} {sql1X}
                    , '{DIST_TYPE}' DIST_TYPE, SUM (ARTPYMT4.GL_DIST_AMT) DIST_AMT
                     FROM {ARTPYMT1} ARTPYMT1, {ARTPYMT2} ARTPYMT2, ARTPYMT4
                     where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO
                       and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO
                       and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO
                     GROUP BY {sql1} {sql1XG}
                     ORDER BY {sql1}"

            End Select

            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim SEG_CODEs(4) As String
                For i As Int32 = 2 To 4
                    SEG_CODEs(i) = row("SEG" & CStr(i) & "_CODE") & ""
                    If row("SEG" & CStr(i) & "_CODE") & "" = "" Then
                        SEG_CODEs(i) = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
                DETL_CTL_DATE = row.Item("PYMT_BATCH_DATE")
                DETL_CTL_NO = row.Item("PYMT_BATCH_NO") & ""
                DETL_POSTING_AMT = Val(row.Item("DIST_AMT") & "")
                Dim rowGLTINTF1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").NewRow
                With rowGLTINTF1
                    .Item("OPS_YYYYPP") = row("OPS_YYYYPP")
                    .Item("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    .Item("JOURNAL_LNO") = JOURNAL_LNO
                    .Item("ACCT_CODE") = row("ACCT_CODE")
                    .Item("SEG2_CODE") = SEG_CODEs(2)
                    .Item("SEG3_CODE") = SEG_CODEs(3)
                    .Item("SEG4_CODE") = SEG_CODEs(4)
                    .Item("DETL_CTL_DATE") = DETL_CTL_DATE
                    .Item("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                    .Item("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
                    .Item("DETL_CTL_NO") = DETL_CTL_NO
                    .Item("DETL_CTL_LNO") = DBNull.Value
                    .Item("DETL_CVX_NO") = DBNull.Value
                    .Item("DETL_CVX_REF_DATE") = DBNull.Value
                    .Item("DETL_CVX_REF_NO") = DBNull.Value
                    .Item("DETL_DESC") = DBNull.Value
                    .Item("DETL_CVX_TYPE") = DBNull.Value
                    .Item("JOURNAL_TYPE") = JOURNAL_TYPE
                    If DIST_TYPE = "GL" Then
                        .Item("DETL_DESC") = row.Item("CUST_NAME")
                        .Item("DETL_CVX_REF_NO") = row.Item("PYMT_NOTE")
                    End If
                End With
                ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
            Next
        Next
    End Sub


    Public Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String

        Select Case COLUMN_NAME
            Case "PYMT_BATCH_NO"
                If chkShowUnpostedBatchesOnly.Checked Then
                    Return "ARTPYMT1.STATUS = '1'"
                Else
                    Return String.Empty
                End If

            Case Else
                Return MyBase.Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME)
        End Select
    End Function
End Class