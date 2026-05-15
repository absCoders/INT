Public Class SPRACOMA

#Region "General Declarations"
    Dim SPTACOMA As String ' Drives Report
    Dim SPTACOMB As String ' Detailed to Store
    Dim sqlSPTACOMA As String
    Dim sqlSPTACOMB As String
    Dim sqlSPTACOMC As String

    Dim reprint As Boolean

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "R"
        Dim sqlw As String = ""
        Prepare_dst(True, sqlw)
        Check_if_Empty("SPTACOMA")
    End Sub

    Public Overrides Sub Print_Report()

        SUBT = ""
        If reprint Then SUBT = "Re-Print"

        CR_params.Add("SUBT", SUBT)
        CR_params.Add("SUMMARY", "0")
        Generate_Report(RPT, , SUBT)

        CR_params.Add("SUBT", SUBT)
        CR_params.Add("SUMMARY", "1")
        Generate_Report(RPT, , SUBT)

        RPT = "SPRACOMC"
        SUBT = "Accrual Items"
        Generate_Report(RPT, , SUBT)

        Print_GL()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        ASCMAIN1.sql = "Insert into SPTACOMB Select * from " & SPTACOMB
        ASCDATA1.ExecuteSQL()

        For Each row As DataRow In dst.Tables("SPTACOMC").Select
            row.Item("ACC_CTL_NO") = ASCMAIN1.Next_Control_No("SPTACOMC.ACC_CTL_NO")
            row.Item("JOURNAL_XNO") = XNO
            row.Item("JOURNAL_IND") = "1"
            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
            row.Item("INIT_DATE") = DATETIME_STAMP
            row.Item("ACC_CTL_NO_ACCRUAL") = row.Item("ACC_CTL_NO")
            row.AcceptChanges()
            row.SetAdded()
        Next
        Update_Record_TDA("SPTACOMC")

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is " & vbCrLf _
            & " Select * from SPTACOMC" & vbCrLf _
            & "  where ACC_CTL_NO_ORIG is Null and JOURNAL_XNO = '" & XNO & "';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SPTACOMB Set ACC_CTL_NO_ACCRUAL = R1.ACC_CTL_NO_ACCRUAL" & vbCrLf _
            & "    where OPS_YYYYPP = R1.OPS_YYYYPP and ASP_CODE = R1.ASP_CODE " & vbCrLf _
            & "      and CUST_CODE = R1.CUST_CODE and INV_NO = R1.INV_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ' Because this accrual is produced after period-end, we need to insert some rows to GLTCREC3
        ' This block was taken from TARPEND1 and modified to work from the accrual work tables

        Dim cols As String = " (OPS_YYYYPP,CREC_TYPE_CODE,CREC_CLASS_CODE,ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,DETL_CTL_TYPE,DETL_CTL_NO,DETL_CVX_TYPE,DETL_CVX_NO,CREC_AMT) "

        ASCMAIN1.sql = "Select '" & RYP & "' OPS_YYYYPP, 'AC' CREC_TYPE_CODE, SPTACOMC.ASP_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", NVL(SPTACOM0.ASP_ACCT_CODE_ACC,SPTPARM1.SP_PARM_ASP_ACCT_CODE_ACC) ACCT_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            & ", 'A' DETL_CTL_TYPE, SPTACOMC.ACC_CTL_NO DETL_CTL_NO" & vbCrLf _
            & ", 'C' DETL_CVX_TYPE, SPTACOMC.CUST_CODE DETL_CVX_NO" & vbCrLf _
            & ", NVL(SPTACOMC.AMT_COMM,0)-NVL(SPTACOMC.AMT_COMM_OFFSET,0) CREC_AMT" & vbCrLf _
            & " from SPTACOMC,SPTACOM0,SPTPARM1" & vbCrLf _
            & " where SPTACOMC.JOURNAL_XNO = '" & XNO & "'" & vbCrLf _
            & "   and SPTACOM0.ASP_CODE = SPTACOMC.ASP_CODE" & vbCrLf _
            & "   and SPTPARM1.SP_PARM_KEY = 'Z'"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        GL_Update()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables()

        With dst

            ASCMAIN1.sql = "Select SPTACOMA.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & " from " & SPTACOMA & " SPTACOMA, ARTCUST1 where ARTCUST1.CUST_CODE = SPTACOMA.CUST_CODE"
            Create_TDA(.Tables.Add, "SPTACOMA", "**", 0, False)

            ASCMAIN1.sql = "Select SPTACOMB.* from " & SPTACOMB & " SPTACOMB"
            Create_TDA(.Tables.Add, "SPTACOMB", "**", 0, False)

            Create_TDA(.Tables.Add, "SPTACOMC", "*")

            ASCMAIN1.sql = "SELECT SPTACOMB.CUST_CODE, SPTACOMB.BRAND_CODE, SPTACOMB.ASP_CODE, SPTACOM2.CUST_STORE_NO" & vbCrLf _
                & ", NVL(SPTACOM2.ASP_COMM_PCT,SPTACOM1.ASP_COMM_PCT) ASP_COMM_PCT" & vbCrLf _
                & ", SUM (SPTACOMB.QTY_SOLD) QTY_SOLD, SUM (SPTACOMB.AMT_SOLD) AMT_SOLD, SUM (SPTACOMB.QTY_EOW) QTY_EOW, SUM (SPTACOMB.AMT_EOW) AMT_EOW" & vbCrLf _
                & ", MIN (SPTACOMB.OPS_YYYYWW_MIN) OPS_YYYYWW_MIN, MIN (SPTACOMB.OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
                & ", SUM (SPTACOMB.AMT_COMM) AMT_COMM" & vbCrLf _
                & " FROM " & SPTACOMB & " SPTACOMB, SPTACOM2, SPTACOM1" & vbCrLf _
                & "where SPTACOM2.ASP_CODE (+) = SPTACOMB.ASP_CODE AND SPTACOM2.CUST_CODE (+) = SPTACOMB.CUST_CODE" & vbCrLf _
                & "  and SPTACOM2.CUST_STORE_NO (+) = SPTACOMB.CUST_STORE_NO" & vbCrLf _
                & "  and SPTACOM1.ASP_CODE (+) = SPTACOMB.ASP_CODE AND SPTACOM1.CUST_CODE (+) = SPTACOMB.CUST_CODE" & vbCrLf _
                & "group by SPTACOMB.CUST_CODE, SPTACOMB.BRAND_CODE, SPTACOMB.ASP_CODE, SPTACOM2.CUST_STORE_NO, NVL(SPTACOM2.ASP_COMM_PCT,SPTACOM1.ASP_COMM_PCT)"
            Create_TDA(.Tables.Add, "SPTACOMD", "**", 0, False, , 0)

            Create_TDA(.Tables.Add, "SPTACOM1", "*", 0)

            ASCMAIN1.sql = "Select GLTPARM3.* from GLTPARM3 where YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select GLTPARM2.* from GLTPARM2 where OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)

            For Each TABLE_NAME As String In New String() _
                {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1", "SPTACOM0", "SOTCHAN1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                If TABLE_NAME = "ARTCUST1" Then ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 1)
            Next

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1", "SPTACOM0", "SOTCHAN1"}
            Fill_Records(TABLE_NAME)
        Next

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

        Fill_Records("GLTPARM3", RYP)
        Fill_Records("GLTPARM2", RYP)
        RYW = dst.Tables("GLTPARM3").Compute("MAX(YYYYWW)", "")

        ASCMAIN1.sql = "Select Count (*) from SPTACOMC where OPS_YYYYPP = '" & RYP & "'"
        reprint = (Val(ASCDATA1.GetDataValue) > 0)

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTACOMB)
        If reprint Then
            RWU = "N"
            ASCDATA1.ExecuteSQL("Insert into " & SPTACOMB & " Select * from SPTACOMB where OPS_YYYYPP = '" & RYP & "'")
        Else
            ASCDATA1.ExecuteSQL("Insert into " & SPTACOMB & " " & Replace(Replace(sqlSPTACOMB, "'YYYYWW'", "'" & RYW & "'"), "'YYYYPP'", "'" & RYP & "'"))
        End If
        If RYP >= ASCMAIN1.CYP Then
            ' never permit update to commission journal prior to month end of RYP
            ' month-end wipes out GLTCREC3 for that month before re-establishing records
            ' month-end establishes GLTCREC3 for AC records for all that occurred on or before period in closing
            ' commission journal writes out AC records for current month
            RWU = "N"
        End If

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor c1 is" & vbCrLf _
            & " Select SPTACOM2.CUST_CODE, SPTACOM2.ASP_CODE, SPTACOM2.ASP_COMM_PCT from SPTACOM1,SPTACOM2" & vbCrLf _
            & "  where SPTACOM1.ASP_CODE = SPTACOM2.ASP_CODE and SPTACOM1.CUST_CODE = SPTACOM2.CUST_CODE" & vbCrLf _
            & " and SPTACOM1.ASP_COMM_BASIS = '2' and ASP_COMM_STATUS = 'A';" & vbCrLf _
            & " Begin for R1 in C1 Loop" & vbCrLf _
            & "   Delete from " & SPTACOMB & " where CUST_CODE = R1.CUST_CODE and ASP_CODE = R1.ASP_CODE and NVL(AMT_COMM,0) = 0;" & vbCrLf _
            & "   Begin Declare Cursor C2 is" & vbCrLf _
            & "    Select CUST_STORE_NO, SUM (AMT_COMM) AMT_COMM from " & SPTACOMB & vbCrLf _
            & "     where CUST_CODE = R1.CUST_CODE and ASP_CODE = R1.ASP_CODE group by CUST_STORE_NO;" & vbCrLf _
            & "   Begin For R2 in C2 Loop" & vbCrLf _
            & "    Update " & SPTACOMB & " set AMT_COMM = Round(AMT_COMM * R1.ASP_COMM_PCT / R2.AMT_COMM,2)" & vbCrLf _
            & "     where CUST_CODE = R1.CUST_CODE and ASP_CODE = R1.ASP_CODE and CUST_STORE_NO = R2.CUST_STORE_NO;" & vbCrLf _
            & "    Update " & SPTACOMB & " set AMT_COMM = AMT_COMM + R1.ASP_COMM_PCT - (Select Sum (AMT_COMM) from " & SPTACOMB & vbCrLf _
            & "     where CUST_CODE = R1.CUST_CODE and ASP_CODE = R1.ASP_CODE and CUST_STORE_NO = R2.CUST_STORE_NO)" & vbCrLf _
            & "      where CUST_CODE = R1.CUST_CODE and ASP_CODE = R1.ASP_CODE and CUST_STORE_NO = R2.CUST_STORE_NO and ROWNUM <2;" & vbCrLf _
            & "  End Loop; End; End;" & vbCrLf _
            & "End Loop; End; End;"
        ASCDATA1.ExecuteSQL()


        ' note hard coded collection AVVAL 
        ' this is used only if there are no retail sales and we need to spread the fixed demo by collection
        ' should probably ask for a default collection in the arrangements file

        Dim COLLECTION_CODE_default As String = "AVVAL"

        ASCMAIN1.sql = "Select '" & RYP & "' OPS_YYYYPP, X.ASP_CODE, X.CUST_CODE, X.CUST_STORE_NO" & vbCrLf _
            & ", X.COLLECTION_CODE, '0000000000' INV_NO, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", 0 QTY_SOLD, 0 AMT_SOLD, 0 QTY_EOW, 0 AMT_EOW" & vbCrLf _
            & ", P.YWMIN, P.YWMAX, SPTACOM2.ASP_COMM_PCT AMT_COMM" & vbCrLf _
            & ", 0 ASP_COMM_PCT, 0 AMT_COMM_CLAIMED, 0 AMT_COMM_PAID" & vbCrLf _
            & ", NULL ACC_CTL_NO_ACCRUAL, NULL INIT_OPER, NULL INIT_DATE, NULL LAST_OPER, NULL LAST_DATE" & vbCrLf _
            & "from (" & vbCrLf _
            & "Select SPTACOM2.CUST_CODE, SPTACOM2.CUST_STORE_NO," & vbCrLf _
            & "SPTACOM2.ASP_CODE, '" & COLLECTION_CODE_default & "' COLLECTION_CODE" & vbCrLf _
            & " from SPTACOM1,SPTACOM2" & vbCrLf _
            & " where SPTACOM1.ASP_CODE = SPTACOM2.ASP_CODE" & vbCrLf _
            & "and SPTACOM1.CUST_CODE = SPTACOM2.CUST_CODE" & vbCrLf _
            & "and SPTACOM1.ASP_COMM_BASIS = '2' and ASP_COMM_STATUS = 'A'" & vbCrLf _
            & " minus " _
            & "Select Distinct CUST_CODE, CUST_STORE_NO, ASP_CODE, '" & COLLECTION_CODE_default & "' COLLECTION_CODE from " & SPTACOMB & ") X, ICTCOLL1, SPTACOM2" & vbCrLf _
            & ", (Select YYYYPP, Max(YYYYWW) YWMAX, Min(YYYYWW) YWMIN from GLTPARM3 where YYYYPP = '" & RYP & "' group by YYYYPP) P" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE and P.YYYYPP = '" & RYP & "'" & vbCrLf _
            & "   and SPTACOM2.CUST_CODE = X.CUST_CODE " & vbCrLf _
            & "   and SPTACOM2.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and SPTACOM2.ASP_CODE = X.ASP_CODE"

        ASCMAIN1.sql = "Insert into " & SPTACOMB & " " & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()


        ASCDATA1.ExecuteSQL("Truncate Table " & SPTACOMA)
        ASCMAIN1.sql = "Insert into " & SPTACOMA & " " & sqlSPTACOMA
        ASCDATA1.ExecuteSQL()

        EnforceConstraints(False)
        Fill_Records("SPTACOM1")
        Fill_Records("SPTACOMA")
        Fill_Records("SPTACOMB")
        Fill_Records("SPTACOMC", "", True, sqlSPTACOMC)
        Fill_Records("SPTACOMD")
        EnforceConstraints(True)

        For Each row As DataRow In dst.Tables("SPTACOMD").Select("ISNULL(CUST_STORE_NO,'')=''")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim ASP_CODE As String = row.Item("ASP_CODE")
            Dim BRAND_CODE As String = row.Item("BRAND_CODE")

            If dst.Tables("SPTACOMD").Select("CUST_CODE = '" & CUST_CODE & "' and ASP_CODE = '" & ASP_CODE & "' and BRAND_CODE = '" & BRAND_CODE & "'").Length = 1 Then
                row.Delete()
            Else
                row.Item("CUST_STORE_NO") = ""
            End If
        Next
        dst.Tables("SPTACOMD").AcceptChanges()


        Prepare_GL_Interface("SPAC")

    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP)
        Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        ' Expense based on Trade Class and Brand

        Dim TOTAL_ASP_COMM As Decimal = 0
        Dim ACCRUALS As New Dictionary(Of String, Decimal)

        For Each row As DataRow In _
            ASCDATA1.SelectDistinct(dst.Tables("SPTACOMA"), New String() {"ASP_CODE", "TRADE_CLASS_CODE", "COLLECTION_CODE"}).Select

            Dim ASP_CODE As String = row.Item("ASP_CODE")
            Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE")
            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")

            Dim rowSPTACOM0 As DataRow = dst.Tables("SPTACOM0").Rows.Find(ASP_CODE)
            Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)

            Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE") & ""
            Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)
 
            Dim DETL_POSTING_AMT As Decimal = Val(dst.Tables("SPTACOMA").Compute("SUM(AMT_COMM)", "ASP_CODE = '" & ASP_CODE & "' and TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "' and COLLECTION_CODE = '" & COLLECTION_CODE & "'") & "")

            If DETL_POSTING_AMT <> 0 Then
                TOTAL_ASP_COMM += DETL_POSTING_AMT
                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("OPS_YYYYPP") = RYP
                rowGLTINTF1.Item("ACCT_CODE") = rowSPTACOM0.Item("ASP_ACCT_CODE_EXP")

                Dim SEG2_CODE As String = rowGLTINTF1("SEG2_CODE")
                If rowSOTCHAN1.Item("SEG2_CODE") & "" <> "" Then SEG2_CODE = rowSOTCHAN1.Item("SEG2_CODE")
                rowGLTINTF1.Item("SEG2_CODE") = SEG2_CODE

                Dim SEG3_CODE As String = TRADE_CLASS_CODE
                If rowSOTTCLS1.Item("SEG3_CODE") & "" <> "" Then SEG3_CODE = rowSOTTCLS1.Item("SEG3_CODE")
                rowGLTINTF1.Item("SEG3_CODE") = SEG3_CODE

                Dim SEG4_CODE As String = COLLECTION_CODE
                If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
                rowGLTINTF1.Item("SEG4_CODE") = SEG4_CODE

                rowGLTINTF1.Item("DETL_CVX_NO") = COLLECTION_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"


                Dim ACCT_CODE As String = rowSPTACOM0.Item("ASP_ACCT_CODE_ACC") & ""
                If ACCT_CODE = "" Then ACCT_CODE = ROWs("SPTPARM1").Item("SP_PARM_ASP_ACCT_CODE_ACC")
                If Not ACCRUALS.ContainsKey(ACCT_CODE) Then
                    ACCRUALS.Add(ACCT_CODE, 0)
                End If
                ACCRUALS(ACCT_CODE) += -1 * DETL_POSTING_AMT
                'rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, -1 * DETL_POSTING_AMT)
                'rowGLTINTF1.Item("OPS_YYYYPP") = RYP
                'rowGLTINTF1.Item("ACCT_CODE") = rowSPTACOM0.Item("ASP_ACCT_CODE_ACC")
            End If
        Next

        ' Accrual

        For Each ACCT_CODE As String In ACCRUALS.Keys
            If ACCRUALS(ACCT_CODE) <> 0 Then
                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, ACCRUALS(ACCT_CODE))
                rowGLTINTF1.Item("OPS_YYYYPP") = RYP
                rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
            End If
        Next

        Return JOURNAL_NO
    End Function

    Function Write_GLTINTF1(JOURNAL_TYPE As String, JOURNAL_NO As String, ByRef JOURNAL_LNO As Integer, DETL_CTL_DATE As Date, DETL_POSTING_AMT As Decimal)
        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = RYP0
        rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
        JOURNAL_LNO += 1
        rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
        rowGLTINTF1("ACCT_CODE") = ""
        rowGLTINTF1("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
        rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
        rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
        rowGLTINTF1("DETL_DESC") = DBNull.Value
        rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
        ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Return rowGLTINTF1
    End Function

    Sub Create_Work_Tables()

        ' NOTE - 
        '  I AM AWARE THE THE SQL STATEMENT FOR SUB-QUERY S USES A CARTESIAN JOIN
        '   WHERE THERE MAY BE SEVERAL ASP_CODES THAT OPERATE OVER THE SAME SUMMARY RESULTS FOR RETAIL AND GROSS SALES.
        '  - THIS WAS THE EASIEST WAY TO ACCOMODATE GROSS SALES AGREEMENTS THAT ARE RECORDED BY INVOICE

        sqlSPTACOMB = "" _
            & "Select 'YYYYPP' OPS_YYYYPP, X.*," & vbCrLf _
            & "DECODE(SPTACOM1.ASP_COMM_BASIS,'2',SPTACOM2.ASP_COMM_PCT,X.AMT_SOLD * " & vbCrLf _
            & "NVL(SPTACOM4.ASP_COMM_PCT," & vbCrLf _
            & "NVL(SPTACOM2.ASP_COMM_PCT," & vbCrLf _
            & "NVL(SPTACOM3.ASP_COMM_PCT,SPTACOM1.ASP_COMM_PCT)))/100) AMT_COMM," & vbCrLf _
            & "DECODE(SPTACOM1.ASP_COMM_BASIS,'2',0," & vbCrLf _
            & "NVL(SPTACOM4.ASP_COMM_PCT," & vbCrLf _
            & "NVL(SPTACOM2.ASP_COMM_PCT," & vbCrLf _
            & "NVL(SPTACOM3.ASP_COMM_PCT,SPTACOM1.ASP_COMM_PCT)))) ASP_COMM_PCT," & vbCrLf _
            & "0 AMT_COMM_CLAIMED," & vbCrLf _
            & "0 AMT_COMM_PAID," & vbCrLf _
            & "NULL ACC_CTL_NO_ACCRUAL," & vbCrLf _
            & "NULL INIT_OPER," & vbCrLf _
            & "NULL INIT_DATE," & vbCrLf _
            & "NULL LAST_OPER," & vbCrLf _
            & "NULL LAST_DATE" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select SPTACOM1.ASP_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, '0000000000' INV_NO, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", SUM (RSTRETL1.QTY_SOLD) QTY_SOLD" & vbCrLf _
            & ", SUM (RSTRETL1.AMT_SOLD) AMT_SOLD" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,'YYYYWW',RSTRETL1.QTY_EOW,0)) QTY_EOW" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,'YYYYWW',RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE,0)) AMT_EOW" & vbCrLf _
            & ", MIN (RSTRETL1.OPS_YYYYWW) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (RSTRETL1.OPS_YYYYWW) OPS_YYYYWW_MAX" & vbCrLf _
            & " from RSTRETL1,ICTITEM1,ICTCOLL1,SPTACOM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYPP = 'YYYYPP'" & vbCrLf _
            & "   and RSTRETL1.CUST_CODE = SPTACOM1.CUST_CODE" & vbCrLf _
            & "   and SPTACOM1.ASP_COMM_BASIS = '0' and SPTACOM1.ASP_COMM_STATUS = 'A'" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by SPTACOM1.ASP_CODE, RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SPTACOM1.ASP_CODE, SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, DECODE (SPTACOM1.ASP_COMM_BY_INVOICE,'1',SOTINVH2.INV_NO,'0000000000') INV_NO" & vbCrLf _
            & ", ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY_SOLD" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) AMT_SOLD" & vbCrLf _
            & ", NULL QTY_EOW" & vbCrLf _
            & ", NULL AMT_EOW" & vbCrLf _
            & ", MIN (SOTINVH2.OPS_YYYYWW) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (SOTINVH2.OPS_YYYYWW) OPS_YYYYWW_MAX" & vbCrLf _
            & " from SOTINVH2,ICTITEM1,ICTCOLL1,SPTACOM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED = 'YYYYPP'" & vbCrLf _
            & "   and SOTINVH2.CUST_CODE = SPTACOM1.CUST_CODE" & vbCrLf _
            & "   and SPTACOM1.ASP_COMM_BASIS = '1' and SPTACOM1.ASP_COMM_STATUS = 'A'" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SPTACOM1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & " group by SPTACOM1.ASP_CODE, SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, DECODE (SPTACOM1.ASP_COMM_BY_INVOICE,'1',SOTINVH2.INV_NO,'0000000000')" & vbCrLf _
            & ", ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ") X, SPTACOM1,SPTACOM2,SPTACOM3,SPTACOM4" & vbCrLf _
            & " where SPTACOM1.ASP_CODE = X.ASP_CODE " & vbCrLf _
            & "   and SPTACOM1.CUST_CODE = X.CUST_CODE " & vbCrLf _
            & "   and SPTACOM1.ASP_COMM_STATUS = 'A'" & vbCrLf _
            & "   and SPTACOM2.ASP_CODE (+) = X.ASP_CODE " & vbCrLf _
            & "   and SPTACOM2.CUST_CODE (+) = X.CUST_CODE " & vbCrLf _
            & "   and SPTACOM2.CUST_STORE_NO (+) = X.CUST_STORE_NO" & vbCrLf _
            & "   and SPTACOM3.ASP_CODE (+) = X.ASP_CODE " & vbCrLf _
            & "   and SPTACOM3.CUST_CODE (+) = X.CUST_CODE " & vbCrLf _
            & "   and SPTACOM3.HC_CODE (+) = X.HC_CODE" & vbCrLf _
            & "   and SPTACOM4.ASP_CODE (+) = X.ASP_CODE " & vbCrLf _
            & "   and SPTACOM4.CUST_CODE (+) = X.CUST_CODE " & vbCrLf _
            & "   and SPTACOM4.CUST_STORE_NO (+) = X.CUST_STORE_NO" & vbCrLf _
            & "   and SPTACOM4.HC_CODE (+) = X.HC_CODE"

        ASCMAIN1.sql = "Select * from SPTACOMB where ROWNUM < 1"
        SPTACOMB = ASCMAIN1.Temp_Table

        sqlSPTACOMC = "Select TRIM(TO_CHAR(ROWNUM,'000000')) ACC_CTL_NO, X.* from (" _
            & "Select OPS_YYYYPP,ASP_CODE,CUST_CODE,NULL HC_CODE,BRAND_CODE" & vbCrLf _
            & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
            & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
            & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
            & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
            & ", 0 DAYS_OF_RETAIL" & vbCrLf _
            & ", 0 AMT_COMM_ADJ, 0 AMT_COMM_CLAIMED, 0 AMT_SOLD_CLAIMED, 0 AMT_COMM_PAID" & vbCrLf _
            & ", NULL OPS_YYYYPP_PAID, NULL PYMT_NO" & vbCrLf _
            & ", NULL ACCT_CTL_NO_ORIG" & vbCrLf _
            & ", NULL PYMT_NO_ORIG" & vbCrLf _
            & ", NULL JOURNAL_XNO" & vbCrLf _
            & ", NULL JOURNAL_IND" & vbCrLf _
            & ", INV_NO" & vbCrLf _
            & " from " & SPTACOMB & " SPTACOMB" & vbCrLf _
            & " group by OPS_YYYYPP,ASP_CODE,CUST_CODE,BRAND_CODE,INV_NO) X"

        ASCMAIN1.sql = "" _
            & "Select OPS_YYYYPP,ASP_CODE,CUST_CODE,HC_CODE,BRAND_CODE,COLLECTION_CODE" & vbCrLf _
            & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
            & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
            & ", SUM (QTY_EOW) QTY_EOW" & vbCrLf _
            & ", SUM (AMT_EOW) AMT_EOW" & vbCrLf _
            & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
            & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
            & ", 0 DAYS_OF_RETAIL" & vbCrLf _
            & ", 0 AMT_COMM_ADJ" & vbCrLf _
            & " from " & SPTACOMB & " SPTACOMB" & vbCrLf _
            & " group by OPS_YYYYPP,ASP_CODE,CUST_CODE,HC_CODE,BRAND_CODE,COLLECTION_CODE"

        sqlSPTACOMA = ASCMAIN1.sql
        SPTACOMA = ASCMAIN1.Temp_Table

    End Sub
End Class