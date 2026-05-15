Public Class SPRACCR1

#Region "General Declarations"
    Dim SPTACCR1 As String
    Dim sqlSPTACCR1 As String
    Dim GLTTSPCA As String
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)

        If ASCMAIN1.EOM = "1" Then
            Set_Read_Only(grpPERIOD_RANGE, True)
        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        sqlw &= "   and NVL(SPTCOOP1.OPS_YYYYPP_ACCRUE,SPTCOOP1.OPS_YYYYPP) <= '" & RYP0 & "'"
        Prepare_dst(True, sqlw)
        Check_if_Empty("SPTACCRA")
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = "Accruals for " & RYPLEGEND0
        CR_params.Add("SUBT", SUBT)
        CR_params.Add("CHKOPEN_ONLY", "0")
        Generate_Report(RPT, , SUBT)
        Print_GL()

        If ASCMAIN1.Running_in_VS Then
            Prepare_Data_Extracts()
        End If

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            RWU = "N"
            If ASCMAIN1.EOM = "1" Then
                Dim RYP As String = Absx1.cmbFor("RYP0").Value
                RYP = Mid(RYP, 1, 4) & Mid(RYP, 6, 2)
                If RYP <> ASCMAIN1.CYP Then
                    EMsg &= vbCr & "For Period End you must use Current Period (" & ASCMAIN1.CYP & ")"
                Else
                    RWU = "R"
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()
        'ASCMAIN1.sql = "Update SPTCOOP1 set JOURNAL_IND = '1', JOURNAL_XNO = '" & XNO & "'" & vbCrLf _
        ' & " where AUTH_NO in (Select AUTH_NO from " & SPTPYMT1 & ")"
        'ASCDATA1.ExecuteSQL(sql)

        'ASCMAIN1.sql = "Insrt into SPTCOOPN" & vbCrLf _
        '    & " Select " & vbCrLf _
        '    & " from " & SPTACCR1 & " SPTACCT1"
        'ASCDATA1.ExecuteSQL()


        ASCDATA1.DeleteRows("SPTCOOPA", "OPS_YYYYPP <> '" & RYP0 & "'")
        dst.Tables("SPTCOOPA").AcceptChanges()
        Update_Record_TDA("SPTCOOPA", "OPS_YYYYPP = '" & RYP0 & "'")

        ASCDATA1.ExecuteSQL("Insert into GLTTSPCA Select * from " & GLTTSPCA)

        Dim NYP As String = ASCMAIN1.Period_Calc(RYP0, 1)
        ASCMAIN1.sql = "Update " & GLTTSPCA & " Set DIST_AMT = -1 * DIST_AMT, OPS_YYYYPP = '" & NYP & "'"
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Insert into GLTTSPCA Select * from " & GLTTSPCA)

        GL_Update()
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables()

        With dst

            ASCMAIN1.sql = "Select SPTACCR1.* from " & SPTACCR1 & " SPTACCR1"
            Create_TDA(.Tables.Add, "SPTACCR1", "**", 0, False, "", 0)

            ASCMAIN1.sql = "Select SPTACCR1.COLLECTION_CODE, SPTACCR1.BRAND_CODE" & vbCrLf _
                & ", SPTACCR1.OPS_YYYYPP, SPTACCR1.EXPENSE_TYPE_CODE, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & ", Sum (SPTACCR1.DIST_AMT_BRAND) as DIST_AMT_BRAND" & vbCrLf _
                & " from " & SPTACCR1 & " SPTACCR1, ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SPTACCR1.CUST_CODE" & vbCrLf _
                & " group by SPTACCR1.COLLECTION_CODE, SPTACCR1.BRAND_CODE" & vbCrLf _
                & ", SPTACCR1.OPS_YYYYPP, SPTACCR1.EXPENSE_TYPE_CODE, ARTCUST1.TRADE_CLASS_CODE"
            Create_TDA(.Tables.Add, "SPTACCRA", "**", 0, False, "")

            Create_TDA(.Tables.Add, "GLTTSPCA", "*", 0)

            ASCMAIN1.sql = "Select * from GLTPARM2" & vbCrLf _
                & " where OPS_YYYYPP >= (Select Min (OPS_YYYYPP) from " & SPTACCR1 & ")" & vbCrLf _
                & "   and OPS_YYYYPP <= (Select Max (OPS_YYYYPP) from " & SPTACCR1 & ")" & vbCrLf _
                & " union Select * from GLTPARM2 where OPS_YYYYPP >= :PARM1 and OPS_YYYYPP <= :PARM2"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "VV", 1)

            ASCMAIN1.sql = "Select * from SPTCOOPA where OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOPA", "**", 0, True, "V", 5)
            With .Tables("SPTCOOPA")
                .Columns.Add("DIST_AMT_LP", GetType(System.Decimal), "IIF(OPS_YYYYPP='000000',DIST_AMT,0)")
                .Columns.Add("DIST_AMT_TP", GetType(System.Decimal), "IIF(OPS_YYYYPP='000000',DIST_AMT,0)")
            End With

            For Each TABLE_NAME As String In New String() _
                {"SPTAVEH1", "SPTTYPE1", "ICTBRAN1", "ICTCOLL1", "ARTCUST1", "SOTTCLS1", "SOTCHAN1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                If TABLE_NAME = "ARTCUST1" Then ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 1)
            Next

            ASCMAIN1.sql = "Select Distinct SPTACCR1.OPS_YYYYPP OPS_YYYYPP_H, SPTACCR1.COLLECTION_CODE" & vbCrLf _
                & " from " & SPTACCR1 & " SPTACCR1" & vbCrLf _
                & " where NVL(SPTACCR1.DIST_AMT_BRAND,0) <> 0" & vbCrLf _
                & " group by SPTACCR1.OPS_YYYYPP, SPTACCR1.COLLECTION_CODE"
            Create_TDA(.Tables.Add, "SPTCOOPM", "**", 0, False, "", 2)

            Create_Relation("SPTCOOPM", "SPTCOOPA", "OPS_YYYYPP_H,COLLECTION_CODE")
            With .Tables("SPTCOOPM")
                .Columns.Add("DIST_AMT_LP", GetType(System.Decimal), "SUM(CHILD.DIST_AMT_LP)")
                .Columns.Add("DIST_AMT_TP", GetType(System.Decimal), "SUM(CHILD.DIST_AMT_TP)")
            End With

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"SPTAVEH1", "SPTTYPE1", "ICTBRAN1", "ICTCOLL1", "ARTCUST1", "SOTTCLS1", "SOTCHAN1"}
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

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTACCR1)
        ASCMAIN1.sql = sqlSPTACCR1 & " " & sqlw
        ASCDATA1.ExecuteSQL("Insert into " & SPTACCR1 & " Select X.*, 0, 0, 0 from (" & ASCMAIN1.sql & ") X")

        ' ASCDATA1.ExecuteSQL("Update " & SPTACCR1 & " A Set TOTAL = (Select Sum (NVL(OTHER_COST,0) + (NVL(QTY,0) * NVL(VEHICLE_CPM,0) / 1000)) from " & SPTACCR1 & " where AUTH_NO = A.AUTH_NO)")
        ASCDATA1.ExecuteSQL("Update " & SPTACCR1 & " A Set TOTAL = NVL(OTHER_COST,0) + (NVL(QTY,0) * NVL(VEHICLE_CPM,0) / 1000)")
        ASCDATA1.ExecuteSQL("Update " & SPTACCR1 & " Set DIST_PCT = CASE WHEN TOTAL = 0 THEN 0 ELSE NVL(DIST_AMT,0)/TOTAL END")
        ASCDATA1.ExecuteSQL("Update " & SPTACCR1 & " Set DIST_AMT_BRAND = DIST_PCT * OPEN_AMT")

        EnforceConstraints(False)

        Fill_Records("GLTPARM2", New String() {ASCMAIN1.Period_Calc(RYP0, -1), ASCMAIN1.Period_Calc(RYP0, 1)})
        Fill_Records("SPTACCR1")
        Fill_Records("SPTACCRA")
        Fill_Records("SPTCOOPA", ASCMAIN1.Period_Calc(RYP0, -1))

        ASCMAIN1.sql = "Select '" & RYP0 & "' OPS_YYYYPP, SPTACCR1.OPS_YYYYPP OPS_YYYYPP_H" & vbCrLf _
            & ", SPTACCR1.COLLECTION_CODE, SPTACCR1.VEHICLE_CODE, SPTACCR1.CUST_CODE" & vbCrLf _
            & ", SUM (SPTACCR1.DIST_AMT_BRAND) DIST_AMT" & vbCrLf _
            & " from " & SPTACCR1 & " SPTACCR1" & vbCrLf _
            & " where NVL(SPTACCR1.DIST_AMT_BRAND,0) <> 0" & vbCrLf _
            & " group by SPTACCR1.OPS_YYYYPP, SPTACCR1.COLLECTION_CODE, SPTACCR1.VEHICLE_CODE, SPTACCR1.CUST_CODE"
        Fill_Records("SPTCOOPA", "", False, ASCMAIN1.sql)

        Fill_Records("SPTCOOPM")

        With dst.Tables("SPTCOOPA")
            .Columns("DIST_AMT_LP").Expression = "IIF(OPS_YYYYPP='" & ASCMAIN1.Period_Calc(RYP0, -1) & "',DIST_AMT,0)"
            .Columns("DIST_AMT_TP").Expression = "IIF(OPS_YYYYPP='" & RYP0 & "',DIST_AMT,0)"
        End With

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SPTCOOPA"), New String() {"OPS_YYYYPP_H", "COLLECTION_CODE"}).Select("")
            Dim OPS_YYYYPP_H As String = row.Item("OPS_YYYYPP_H")
            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")
            If dst.Tables("SPTCOOPM").Rows.Find(New String() {OPS_YYYYPP_H, COLLECTION_CODE}) Is Nothing Then
                dst.Tables("SPTCOOPM").Rows.Add(New Object() {OPS_YYYYPP_H, COLLECTION_CODE, 0, 0})
            End If
        Next

        'For Each row As DataRow In dst.Tables("SPTACCR1").Select("")
        '    Dim rowGLTTSPCA As DataRow = dst.Tables("GLTTSPCA").NewRow
        'Next

        EnforceConstraints(True)

        Prepare_GL_Interface("SPCA")
    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        Dim NYP As String = ASCMAIN1.Period_Calc(RYP0, 1)
        Dim accrual As Decimal = 0

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP0)
        Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        ' Expense based on Vehicle, Trade Class and Collection

        Dim EXPENSE_TYPE_CODE_accruals As New Dictionary(Of String, Decimal)

        For Each rowSPTACCRA As DataRow In dst.Tables("SPTACCRA").Select("")
            Dim DETL_POSTING_AMT As Decimal = Val(rowSPTACCRA.Item("DIST_AMT_BRAND") & "")
            If DETL_POSTING_AMT <> 0 Then

                Dim EXPENSE_TYPE_CODE As String = rowSPTACCRA.Item("EXPENSE_TYPE_CODE")
                Dim BRAND_CODE As String = rowSPTACCRA.Item("BRAND_CODE")
                Dim COLLECTION_CODE As String = rowSPTACCRA.Item("COLLECTION_CODE")
                Dim TRADE_CLASS_CODE As String = rowSPTACCRA.Item("TRADE_CLASS_CODE")

                Dim rowSPTTYPE1 As DataRow = dst.Tables("SPTTYPE1").Rows.Find(EXPENSE_TYPE_CODE)
                Dim rowICTBRAN1 As DataRow = dst.Tables("ICTBRAN1").Rows.Find(BRAND_CODE)
                Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)

                Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE") & ""
                Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)

                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)

                Dim ACCT_CODE As String = rowSPTTYPE1.Item("ACCT_CODE") & ""
                Dim ACCT_CODE_ACC As String = rowSPTTYPE1.Item("ACCT_CODE_ACC") & ""
                rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE

                Dim SEG2_CODE As String = rowGLTINTF1("SEG2_CODE")
                If rowSOTCHAN1.Item("SEG2_CODE") & "" <> "" Then SEG2_CODE = rowSOTCHAN1.Item("SEG2_CODE")
                rowGLTINTF1.Item("SEG2_CODE") = SEG2_CODE

                Dim SEG3_CODE As String = TRADE_CLASS_CODE
                If rowSOTTCLS1.Item("SEG3_CODE") & "" <> "" Then SEG3_CODE = rowSOTTCLS1.Item("SEG3_CODE")
                rowGLTINTF1.Item("SEG3_CODE") = SEG3_CODE

                Dim SEG4_CODE As String = COLLECTION_CODE
                If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
                rowGLTINTF1.Item("SEG4_CODE") = SEG4_CODE

                rowGLTINTF1.Item("DETL_CVX_REF_NO") = TRADE_CLASS_CODE & ":" & COLLECTION_CODE
                rowGLTINTF1.Item("DETL_CVX_NO") = EXPENSE_TYPE_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "C" ' S/B EXPENSE_TYPE_CODE

                Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)

                If ACCT_CODE_ACC <> "" Then
                    If Not EXPENSE_TYPE_CODE_accruals.ContainsKey(EXPENSE_TYPE_CODE) Then
                        EXPENSE_TYPE_CODE_accruals.Add(EXPENSE_TYPE_CODE, 0)
                    End If
                    EXPENSE_TYPE_CODE_accruals(EXPENSE_TYPE_CODE) -= DETL_POSTING_AMT
                Else
                    accrual = accrual - DETL_POSTING_AMT
                End If

            End If
        Next

        If EXPENSE_TYPE_CODE_accruals.Count <> 0 Then
            For Each EXPENSE_TYPE_CODE As String In EXPENSE_TYPE_CODE_accruals.Keys
                Dim rowSPTTYPE1 As DataRow = dst.Tables("SPTTYPE1").Rows.Find(EXPENSE_TYPE_CODE)
                Dim ACCT_CODE_ACC As String = rowSPTTYPE1.Item("ACCT_CODE_ACC") & ""

                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, EXPENSE_TYPE_CODE_accruals(EXPENSE_TYPE_CODE))
                rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE_ACC
                Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)
            Next
        End If

        If accrual <> 0 Then
            Dim ACCT_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_PROMO_ACCT_CODE_ACC") & ""
            Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, accrual)
            rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
            Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)
        End If




        ASCMAIN1.sql = "Insert into " & GLTTSPCA & " " & vbCrLf _
                & "Select '" & RYP0 & "' OPS_YYYYPP," & vbCrLf _
                & "SPTTYPE1.ACCT_CODE, " & vbCrLf _
                & "NVL(SOTCHAN1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') SEG2_CODE, " & vbCrLf _
                & "NVL(SOTTCLS1.SEG3_CODE,ARTCUST1.TRADE_CLASS_CODE) SEG3_CODE, " & vbCrLf _
                & "NVL(ICTCOLL1.SEG4_CODE,X.COLLECTION_CODE) SEG4_CODE," & vbCrLf _
                & "'" & JOURNAL_NO & "' JOURNAL_NO," & vbCrLf _
                & "'" & XNO & "' JOURNAL_XNO," & vbCrLf _
                & "X.DIST_AMT_BRAND DIST_AMT," & vbCrLf _
                & "X.AUTH_NO DIST_CTL_NO," & vbCrLf _
                & "0 DIST_CTL_LNO," & vbCrLf _
                & "SPTCOOP1.CUST_REF_NUM DIST_REF," & vbCrLf _
                & "SPTCOOP1.BOOKING_NAME DIST_DESC," & vbCrLf _
                & "X.CUST_CODE," & vbCrLf _
                & "X.EXPENSE_TYPE_CODE," & vbCrLf _
                & "X.COLLECTION_CODE," & vbCrLf _
                & "X.VEHICLE_CODE," & vbCrLf _
                & "SPTCOOP1.SEASON_CODE," & vbCrLf _
                & "X.DATE_START," & vbCrLf _
                & "X.DATE_END," & vbCrLf _
                & "X.DATE_ACCRUE," & vbCrLf _
                & "X.OPS_YYYYPP OPS_YYYYPP_ACCRUE," & vbCrLf _
                & "X.INIT_OPER," & vbCrLf _
                & "X.INIT_DATE," & vbCrLf _
                & "X.LAST_OPER," & vbCrLf _
                & "X.LAST_DATE" & vbCrLf _
                & " from " & SPTACCR1 & " X, SPTTYPE1, ARTCUST1, SOTTCLS1, ICTCOLL1, SPTCOOP1, SOTCHAN1" & vbCrLf _
                & " where SPTCOOP1.AUTH_NO = X.AUTH_NO" & vbCrLf _
                & "   and SPTTYPE1.EXPENSE_TYPE_CODE (+) = X.EXPENSE_TYPE_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE" & vbCrLf _
                & "   and SOTCHAN1.CHANNEL_CODE (+) = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE"
        ASCDATA1.ExecuteSQL()

        Return JOURNAL_NO
    End Function

    Sub Write_GLTINTF1_Reversal(row As DataRow, YP As String)
        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1.ItemArray = row.ItemArray
        rowGLTINTF1.Item("OPS_YYYYPP") = YP
        rowGLTINTF1.Item("DETL_POSTING_AMT") = -1 * Val(rowGLTINTF1.Item("DETL_POSTING_AMT") & "")
        dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
    End Sub

    Function Write_GLTINTF1(
                           JOURNAL_TYPE As String,
                           JOURNAL_NO As String,
                           ByRef JOURNAL_LNO As Integer,
                           DETL_CTL_DATE As Date,
                           DETL_POSTING_AMT As Decimal) As DataRow

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
        dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Return rowGLTINTF1
    End Function

    Sub Create_Work_Tables()
        '            & ", Decode ((SPTCOOP1.OTHER_COST + (SPTCOOP1.QTY * SPTCOOP1.VEHICLE_CPM / 1000)), 0, 0, SPTCOOP3.DIST_AMT / (SPTCOOP1.OTHER_COST + (SPTCOOP1.QTY * SPTCOOP1.VEHICLE_CPM / 1000))) DIST_PCT" & vbCrLf _

        sqlSPTACCR1 = "Select SPTCOOP3.*, ICTCOLL1.BRAND_CODE, SPTCOOP1.EXPENSE_TYPE_CODE" & vbCrLf _
            & ", SPTCOOP1.VEHICLE_CODE, SPTCOOP1.VEHICLE_CPM" & vbCrLf _
            & ", SPTCOOP1.DATE_START, SPTCOOP1.DATE_END, SPTCOOP1.DATE_ACCRUE" & vbCrLf _
            & ", SPTCOOP1.QTY, SPTCOOP1.OTHER_COST, SPTCOOP1.NOTES" & vbCrLf _
            & ", SPTCOOP1.OPEN_AMT, SPTCOOP1.PAID_AMT, SPTCOOP1.PYMTS, NVL(SPTCOOP1.OPS_YYYYPP_ACCRUE,SPTCOOP1.OPS_YYYYPP) OPS_YYYYPP" & vbCrLf _
            & ", SPTCOOP1.CUST_CODE, SPTCOOP1.AUTH_DATE, SPTCOOP1.AUTH_REQ_BY, SPTCOOP1.SREP_CODE" & vbCrLf _
            & ", SPTCOOP1.INIT_DATE, SPTCOOP1.INIT_OPER, SPTCOOP1.LAST_DATE, SPTCOOP1.LAST_OPER" & vbCrLf _
            & " from SPTCOOP1, SPTCOOP3, ICTCOLL1" & vbCrLf _
            & " where SPTCOOP1.AUTH_NO = SPTCOOP3.AUTH_NO" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
            & "   and SPTCOOP1.STATUS_CODE = 'O'"

        ASCMAIN1.sql = sqlSPTACCR1 & " and ROWNUM < 1"
        SPTACCR1 = ASCMAIN1.Temp_Table
        ' ASCDATA1.ExecuteSQL("Alter Table " & SPTACCR1 & " Add Primary Key (AUTH_NO)")

        ASCDATA1.ExecuteSQL("Alter Table " & SPTACCR1 & " Add TOTAL NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & SPTACCR1 & " Add DIST_PCT NUMBER (10,6)")
        ASCDATA1.ExecuteSQL("Alter Table " & SPTACCR1 & " Add DIST_AMT_BRAND NUMBER (13,2)")

        ASCMAIN1.sql = "Select * from GLTTSPCA where ROWNUM < 1"
        GLTTSPCA = ASCMAIN1.Temp_Table

    End Sub


    Sub Prepare_Data_Extracts()


        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = dst.Tables("SPTACCR1")

        grdASTEXPT1.Text = "Promo Expense Accrual - " & Mid(RYPLEGEND, 10, 6)
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        'Dim Cs As New List(Of String)
        'Dim G As Integer = 0
        'For Each COLUMN_NAME As String In COLUMN_NAMEs
        '    Cs.Add(COLUMN_NAME)
        '    G += 1
        '    Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , Color.Gold)
        'Next
        Set_DX_Column(grdASTEXPT1, "EXPENSE_TYPE_CODE", "Expense Type", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP", "YP", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "AUTH_NO", "Auth No", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "CUST_CODE", "Customer", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "BRAND_CODE", "Brand", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collection", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "FEATURE_DESC", "Feature", 200, , , System.Drawing.Color.Gold)

        Set_DX_Column(grdASTEXPT1, "VEHICLE_CODE", "Vehicle", 100, , , System.Drawing.Color.Gold)
        Set_DX_Column(grdASTEXPT1, "DATE_START", "Start", 100, "MM/dd/yyyy", , System.Drawing.Color.Gold)
        'If Not Cs.Contains("COLLECTION_CODE") Then Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collctn", 70)
        'If Not Cs.Contains("ITEM_COST_MAKE_BUY") Then Set_DX_Column(grdASTEXPT1, "ITEM_COST_MAKE_BUY", "MB", 30)
        'If Not Cs.Contains("COST_CATGY_CODE") Then Set_DX_Column(grdASTEXPT1, "COST_CATGY_CODE", "Cost Catgy", 70)
        'If Not Cs.Contains("PROD_CODE") Then Set_DX_Column(grdASTEXPT1, "PROD_CODE", "Prod", 70)
        'If Not Cs.Contains("VEND_CODE") Then Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 100)
        'Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Description", 130)
        'Set_DX_Column(grdASTEXPT1, "STD_COST_LM", "Std Prev", 90, "#.0000", , Color.Orange)
        'Set_DX_Column(grdASTEXPT1, "STD_COST_TM", "Std Curr", 90, "#.0000", , Color.Orange)

        Set_DX_Column(grdASTEXPT1, "DIST_AMT_BRAND", "Dist Amt", 90, "#,##0.00", , System.Drawing.Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "QTY_SHP", "#Shp", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "QTY_RTN", "#Rtn", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "QTY_REC", "#Rec", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "QTY_ADJ", "#Adj", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "QTY_CON", "#Con", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "QTY_EOM", "#End", 90, "#,##0", , Color.LightBlue)
        'Set_DX_Column(grdASTEXPT1, "AMT_BOM_PREV", "$Beg@Prev", 90, "#,##0", , Color.Gold)
        'Set_DX_Column(grdASTEXPT1, "AMT_REVAL", "$ReVal", 90, "#,##0", , Color.Gold)
        'Set_DX_Column(grdASTEXPT1, "AMT_BOM", "$Beg@Curr", 90, "#,##0", , Color.Gold)
        'Set_DX_Column(grdASTEXPT1, "AMT_SHP", "$Shp", 90, "#,##0", , Color.LightGreen)
        'Set_DX_Column(grdASTEXPT1, "AMT_RTN", "$Rtn", 90, "#,##0", , Color.LightGreen)
        'Set_DX_Column(grdASTEXPT1, "AMT_REC", "$Rec", 90, "#,##0", , Color.LightGreen)
        'Set_DX_Column(grdASTEXPT1, "AMT_ADJ", "$Adj", 90, "#,##0", , Color.LightGreen)
        'Set_DX_Column(grdASTEXPT1, "AMT_CON", "$Con", 90, "#,##0", , Color.LightGreen)
        'Set_DX_Column(grdASTEXPT1, "AMT_EOM", "$End", 90, "#,##0", , Color.LightGreen)
        'Set_DX_Column(grdASTEXPT1, "PV_DEF", "$PPV", 90, "#,##0", , Color.LightGreen)
        'Set_DX_Column(grdASTEXPT1, "MV_DEF", "$MUV", 90, "#,##0", , Color.LightGreen)

        Create_Summary(grdASTEXPT1, "DIST_AMT_BRAND")
        'Create_Summary(grdASTEXPT1, "AMT_REVAL")



        'grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        'Sort_grdColumns(grdASTEXPT1, "ITEM_CODE")

    End Sub


End Class