Public Class SPRPYMT1

    ' MT

#Region "General Declarations"
    Dim SPTPYMT1 As String
    Dim CHKMTD As String = ""
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""

        sqlw &= "   and SPTPYMT1.OPS_YYYYPP >= '" & RYP0 & "' and SPTPYMT1.OPS_YYYYPP <= '" & RYP1 & "'"
        If CHKMTD = "1" Then
            sqlw &= "   and NVL(SPTPYMT1.JOURNAL_IND,'0') = '0'" '  is Null"
        End If

        sqlw &= SQL_in("CUST_CODE", "SPTCOOP1.CUST_CODE") & vbCrLf
        sqlw &= SQL_in("AUTH_NO", "SPTCOOP1.AUTH_NO") & vbCrLf

        Prepare_dst(True, sqlw)

        Check_if_Empty("SPTPYMT1")
    End Sub

    Public Overrides Sub Print_Report()

        If CHKMTD = "1" Then
            If RYP0 = RYP1 Then
                SUBT = "Payments made during " & RYPLEGEND0
            Else
                SUBT = "Payments made between " & RYPLEGEND0 & " and " & RYPLEGEND1
            End If
        Else
            SUBT = "Payments Register"
        End If
        CR_params.Add("SUBT", SUBT)
        Generate_Report(RPT, , SUBT)

        If CHKMTD = "1" Then
            Print_GL()
        End If

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Dim YP0 As String = Absx1.cmbFor("RYP0").Value
            Dim YP1 As String = Absx1.cmbFor("RYP1").Value

            YP0 = Mid(YP0, 1, 4) & Mid(YP0, 6, 2)
            YP1 = Mid(YP1, 1, 4) & Mid(YP1, 6, 2)

            RWU = "N"
            CHKMTD = IIf(Absx1.chkFor("CHKMTD").Checked, "1", "0")
            If CHKMTD = "1" Then
                If YP0 = YP1 And YP0 = ASCMAIN1.CYP Then
                    RWU = "R"
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()
        If CHKMTD = "1" Then
            ASCMAIN1.sql = "Update SPTPYMT1 set JOURNAL_IND = '1', JOURNAL_XNO = '" & XNO & "'" & vbCrLf _
                & " where PYMT_NO in (Select PYMT_NO from " & SPTPYMT1 & ")"
            ASCDATA1.ExecuteSQL(sql)

            GL_Update()
        Else
            ASCMAIN1.sql = "Update SPTPYMT1 set REGISTER_IND = '1', REGISTER_XNO = '" & XNO & "'" & vbCrLf _
                & " where PYMT_NO in (Select PYMT_NO from " & SPTPYMT1 & ")"
            ASCDATA1.ExecuteSQL(sql)
        End If
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

            ASCMAIN1.sql = "Select SPTPYMT1.* from " & SPTPYMT1 & " SPTPYMT1"
            Create_TDA(.Tables.Add, "SPTPYMT1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SPTPYMT2.* from SPTPYMT2, " & SPTPYMT1 & " SPTPYMT1 where SPTPYMT2.PYMT_NO = SPTPYMT1.PYMT_NO"
            Create_TDA(.Tables.Add, "SPTPYMT2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SPTPYMT3.* from SPTPYMT3, " & SPTPYMT1 & " SPTPYMT1 where SPTPYMT3.PYMT_NO = SPTPYMT1.PYMT_NO"
            Create_TDA(.Tables.Add, "SPTPYMT3", "**", 0, False, "", 5)

            ASCMAIN1.sql = "Select Distinct SPTCOOP1.*" & vbCrLf _
                & " from SPTCOOP1, SPTPYMT2" & vbCrLf _
                & ", " & SPTPYMT1 & " SPTPYMT1" & vbCrLf _
                & " where SPTCOOP1.AUTH_NO = SPTPYMT2.AUTH_NO" & vbCrLf _
                & "   and SPTPYMT2.PYMT_NO = SPTPYMT1.PYMT_NO"
            ' VS 
            ASCMAIN1.sql = "Select SPTCOOP1.* from SPTCOOP1 where AUTH_NO in " & vbCrLf _
                & " (Select Distinct AUTH_NO from SPTPYMT2, " & SPTPYMT1 & " SPTPYMT1" & vbCrLf _
                & " where SPTPYMT2.PYMT_NO = SPTPYMT1.PYMT_NO)"
            Create_TDA(.Tables.Add, "SPTCOOP1", "**", 0, False, "", 1)
            .Tables("SPTCOOP1").Columns.Add("TOTAL", GetType(System.Decimal), "ISNULL(OTHER_COST,0) + (ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000)")
 
            For Each TABLE_NAME As String In New String() _
                {"SPTAVEH1", "SPTTYPE1", "ICTBRAN1", "ICTCOLL1", "ARTCUST1", "SOTTCLS1", "ARTPOST1", "SOTCHAN1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                If TABLE_NAME = "ARTCUST1" Then ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 1)
            Next

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"SPTAVEH1", "SPTTYPE1", "ICTBRAN1", "ICTCOLL1", "ARTCUST1", "SOTTCLS1", "ARTPOST1", "SOTCHAN1"}
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

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTPYMT1)
        ASCMAIN1.sql = "Select * from SPTPYMT1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
        ASCDATA1.ExecuteSQL("Insert into " & SPTPYMT1 & " " & ASCMAIN1.sql)

        EnforceConstraints(False)
        Fill_Records("SPTCOOP1")
        Fill_Records("SPTPYMT3")
        Fill_Records("SPTPYMT2")
        Fill_Records("SPTPYMT1")
        EnforceConstraints(True)

        If CHKMTD = "1" Then
            Prepare_GL_Interface("SPCP")
        End If
    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim SP_PARM_PROMO_ACCT_CODE_PPD As String = ROWs("SPTPARM1").Item("SP_PARM_PROMO_ACCT_CODE_PPD") & ""

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP0)
        Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        ' Expense based on Expense Type, Trade Class and Collection

        EnforceConstraints(False)

        For Each rowSPTPYMT3 As DataRow In dst.Tables("SPTPYMT3").Select("")
            Dim DETL_POSTING_AMT As Decimal = Val(rowSPTPYMT3.Item("DIST_AMT_PYMT") & "")

            If DETL_POSTING_AMT <> 0 Then
                Dim AUTH_NO As String = rowSPTPYMT3.Item("AUTH_NO")
                Dim AUTH_LNO As Integer = rowSPTPYMT3.Item("AUTH_LNO")
                Dim rowSPTCOOP1 As DataRow = dst.Tables("SPTCOOP1").Rows.Find(New Object() {AUTH_NO})
                Dim OPS_YYYYPP_promo As String = rowSPTCOOP1.Item("OPS_YYYYPP")
                Dim VEHICLE_CODE As String = rowSPTCOOP1.Item("VEHICLE_CODE")
                Dim EXPENSE_TYPE_CODE As String = rowSPTCOOP1.Item("EXPENSE_TYPE_CODE")
                Dim rowSPTAVEH1 As DataRow = dst.Tables("SPTAVEH1").Rows.Find(VEHICLE_CODE)
                Dim rowSPTTYPE1 As DataRow = dst.Tables("SPTTYPE1").Rows.Find(EXPENSE_TYPE_CODE)

                Dim PYMT_NO As String = rowSPTPYMT3.Item("PYMT_NO")
                Dim rowSPTPYMT1 As DataRow = dst.Tables("SPTPYMT1").Rows.Find(New Object() {PYMT_NO})
                Dim OPS_YYYYPP_pymt As String = rowSPTPYMT1.Item("OPS_YYYYPP")

                Dim BRAND_CODE As String = rowSPTPYMT3.Item("BRAND_CODE")
                Dim COLLECTION_CODE As String = rowSPTPYMT3.Item("COLLECTION_CODE")
                Dim TRADE_CLASS_CODE As String = rowSPTPYMT3.Item("TRADE_CLASS_CODE")

                Dim rowICTBRAN1 As DataRow = dst.Tables("ICTBRAN1").Rows.Find(BRAND_CODE)
                Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)

                Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE")
                Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)

                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("ACCT_CODE") = rowSPTTYPE1.Item("ACCT_CODE")

                Dim SEG2_CODE As String = rowSOTCHAN1.Item("SEG2_CODE") & ""
                If SEG2_CODE = "" Then SEG2_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                rowGLTINTF1.Item("SEG2_CODE") = SEG2_CODE

                Dim SEG3_CODE As String = TRADE_CLASS_CODE
                If rowSOTTCLS1.Item("SEG3_CODE") & "" <> "" Then SEG3_CODE = rowSOTTCLS1.Item("SEG3_CODE")
                rowGLTINTF1.Item("SEG3_CODE") = SEG3_CODE

                Dim SEG4_CODE As String = COLLECTION_CODE
                If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then SEG4_CODE = rowICTCOLL1.Item("SEG4_CODE")
                rowGLTINTF1.Item("SEG4_CODE") = SEG4_CODE

                rowGLTINTF1.Item("DETL_CVX_NO") = VEHICLE_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "C" ' S/B VEHICLE_CODE

                If OPS_YYYYPP_pymt >= OPS_YYYYPP_promo Then
                    rowGLTINTF1.Item("OPS_YYYYPP") = OPS_YYYYPP_pymt
                Else
                    rowGLTINTF1.Item("OPS_YYYYPP") = OPS_YYYYPP_promo

                    Dim rowGLTINTF1_PP As DataRow

                    rowGLTINTF1_PP = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, -1 * DETL_POSTING_AMT)

                    rowGLTINTF1_PP.ItemArray = rowGLTINTF1.ItemArray
                    rowGLTINTF1_PP.Item("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1_PP.Item("ACCT_CODE") = SP_PARM_PROMO_ACCT_CODE_PPD
                    rowGLTINTF1_PP.Item("OPS_YYYYPP") = OPS_YYYYPP_promo
                    rowGLTINTF1_PP.Item("DETL_POSTING_AMT") = -1 * DETL_POSTING_AMT

                    rowGLTINTF1_PP = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                    rowGLTINTF1_PP.ItemArray = rowGLTINTF1.ItemArray
                    rowGLTINTF1_PP.Item("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1_PP.Item("ACCT_CODE") = SP_PARM_PROMO_ACCT_CODE_PPD
                    rowGLTINTF1_PP.Item("OPS_YYYYPP") = OPS_YYYYPP_pymt
                End If

            End If
        Next

        EnforceConstraints(True)

        ' AR/AP Offset

        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ROWs("SPTPARM1").Item("SP_PARM_PROMO_TYPE_CODE"))
        Dim rowARTPOST1 As DataRow = dst.Tables("ARTPOST1").Rows.Find(rowSOTTYPE1.Item("POST_CODE"))

        For Each rowSPTPYMT1 As DataRow In dst.Tables("SPTPYMT1").Select("")
            Dim DETL_POSTING_AMT As Decimal = Val(rowSPTPYMT1.Item("PYMT_REF_AMT") & "")
            If DETL_POSTING_AMT <> 0 Then
                Dim PYMT_TYPE As String = rowSPTPYMT1.Item("PYMT_TYPE")
                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, -1 * DETL_POSTING_AMT)
                rowGLTINTF1.Item("OPS_YYYYPP") = rowSPTPYMT1.Item("OPS_YYYYPP")
                If PYMT_TYPE = "R" Then
                    rowGLTINTF1.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
                Else
                    rowGLTINTF1.Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_PROMO_ACCT_CODE_APX")
                End If
                rowGLTINTF1.Item("DETL_CVX_NO") = rowSPTPYMT1.Item("PYMT_NO")
            End If
        Next

        Return JOURNAL_NO
    End Function

    Function Write_GLTINTF1(JOURNAL_TYPE As String, JOURNAL_NO As String, ByRef JOURNAL_LNO As Integer,
                            DETL_CTL_DATE As Date, DETL_POSTING_AMT As Decimal,
                            Optional DETL_DESC As String = "", Optional CUST_CODE As String = "")
        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = RYP0
        rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
        JOURNAL_LNO += 1
        With rowGLTINTF1
            .Item("JOURNAL_LNO") = JOURNAL_LNO
            .Item("ACCT_CODE") = ""
            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            .Item("DETL_CTL_DATE") = DETL_CTL_DATE
            .Item("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
            .Item("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
            .Item("DETL_DESC") = DETL_DESC
            If CUST_CODE <> "" Then
                .Item("DETL_CVX_TYPE") = "C"
                .Item("DETL_CVX_NO") = CUST_CODE
            End If
            .Item("JOURNAL_TYPE") = JOURNAL_TYPE
        End With
        dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Return rowGLTINTF1
    End Function

    Sub Create_Work_Tables()
        ASCMAIN1.sql = "Select * from SPTPYMT1 where ROWNUM < 1"
        SPTPYMT1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SPTPYMT1 & " Add Primary Key (PYMT_NO)")
    End Sub
End Class