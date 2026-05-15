Public Class SPRACOMP

#Region "General Declarations"
    Dim SPTAPMT1 As String
    Dim CHKMTD As String = ""
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        If ASCMAIN1.EOM <> "1" Then
            Absx1.chkFor("CHKMTD").Checked = False
            Absx1.chkFor("CHKMTD").Visible = False
        End If

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
    End Sub

    Protected Overrides Sub Build_Workfile()
        ' RWU = "R"
        Dim sqlw As String = ""
        sqlw &= "   and SPTAPMT1.OPS_YYYYPP >= '" & RYP0 & "' and SPTAPMT1.OPS_YYYYPP <= '" & RYP1 & "'"
        If CHKMTD = "1" Then
            sqlw &= "   and NVL(SPTAPMT1.JOURNAL_IND,'0') = '0'" '  is Null"
        End If
        sqlw &= SQL_in("CUST_CODE", "SPTAPMT1.CUST_CODE") & vbCrLf
        sqlw &= SQL_in("PYMT_NO", "SPTAPMT1.PYMT_NO") & vbCrLf
        Prepare_dst(True, sqlw)
        Check_if_Empty("SPTAPMT1")
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = "Payments made during " & RYPLEGEND

        If RYP0 = RYP1 Then
            SUBT = "Payments made during " & RYPLEGEND0
        Else
            SUBT = "Payments made between " & RYPLEGEND0 & " and " & RYPLEGEND1
        End If

        CR_params.Add("SUBT", SUBT)
        Generate_Report(RPT, , SUBT)

        'If CHKMTD = "1" Then
        Print_GL()
        'End If
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

        'If CHKMTD = "1" Then

        ASCMAIN1.sql = "Update SPTAPMT1 set JOURNAL_IND = '1', JOURNAL_XNO = '" & XNO & "'" & vbCrLf _
            & " where PYMT_NO in (Select PYMT_NO from " & SPTAPMT1 & ")"
        ASCDATA1.ExecuteSQL(sql)

        GL_Update()

        'Else
        '    ASCMAIN1.sql = "Update SPTAPMT1 set REGISTER_IND = '1', REGISTER_XNO = '" & XNO & "'" & vbCrLf _
        '        & " where PYMT_NO in (Select PYMT_NO from " & SPTAPMT1 & ")"
        '    ASCDATA1.ExecuteSQL(sql)
        'End If

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

            ASCMAIN1.sql = "Select SPTAPMT1.* from " & SPTAPMT1 & " SPTAPMT1"
            Create_TDA(.Tables.Add, "SPTAPMT1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SPTAPMT2.*,ARTCUST1.TRADE_CLASS_CODE,SPTAPMT1.OPS_YYYYPP" & vbCrLf _
                & " from SPTAPMT2, " & SPTAPMT1 & " SPTAPMT1, ARTCUST1" & vbCrLf _
                & " where SPTAPMT2.PYMT_NO = SPTAPMT1.PYMT_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SPTAPMT1.CUST_CODE"
            Create_TDA(.Tables.Add, "SPTAPMT2", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select SPTACOMC.* from SPTACOMC, " & SPTAPMT1 & " SPTAPMT1" & vbCrLf _
                & " where SPTACOMC.PYMT_NO = SPTAPMT1.PYMT_NO"
            Create_TDA(.Tables.Add, "SPTACOMC", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select GLTPARM3.* from GLTPARM3 where YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select GLTPARM2.* from GLTPARM2 where OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)

            For Each TABLE_NAME As String In New String() _
                {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1", "ARTPOST1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                If TABLE_NAME = "ARTCUST1" Then ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 1)
            Next

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1", "ARTPOST1"}
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

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTAPMT1)
        ASCMAIN1.sql = "Select * from SPTAPMT1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
        ASCDATA1.ExecuteSQL("Insert into " & SPTAPMT1 & " " & ASCMAIN1.sql)

        EnforceConstraints(False)
        Fill_Records("SPTAPMT1")
        Fill_Records("SPTAPMT2")
        Fill_Records("SPTACOMC")
        EnforceConstraints(True)

        'If CHKMTD = "1" Then
        Prepare_GL_Interface("SPAP")
        'End If
    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        ' Expense based on Trade Class and Brand

        For Each rowYP As DataRow In _
            ASCDATA1.SelectDistinct(dst.Tables("SPTAPMT2"), New String() {"OPS_YYYYPP"}).Select

            Dim OPS_YYYYPP As String = rowYP.Item("OPS_YYYYPP")
            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", OPS_YYYYPP)
            Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

            For Each row As DataRow In _
                    ASCDATA1.SelectDistinct(dst.Tables("SPTAPMT2").Select("OPS_YYYYPP = '" & OPS_YYYYPP & "'"), _
                                            New String() {"TRADE_CLASS_CODE", "COLLECTION_CODE"}).Select

                Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE")
                Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")

                Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)
                Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)

                Dim BRAND_CODE As String = rowICTCOLL1.Item("BRAND_CODE")
                Dim rowICTBRAN1 As DataRow = dst.Tables("ICTBRAN1").Rows.Find(BRAND_CODE)

                Dim sqlw As String = "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "' and COLLECTION_CODE = '" & COLLECTION_CODE & "' and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
                Dim DETL_POSTING_AMT As Decimal = Val(dst.Tables("SPTAPMT2").Compute("SUM(AMT_COMM_ADJ)", sqlw) & "")

                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                    rowGLTINTF1.Item("OPS_YYYYPP") = OPS_YYYYPP

                    If System.Math.Abs(DETL_POSTING_AMT) < 1.0 Then
                        ' THIS IS REALLY JUST ROUNDING ERRORS

                        Dim ACCT_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_ACCT_ROUNDING") & ""
                        rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE

                    Else
                        Dim ACCT_CODE As String = ROWs("SPTPARM1").Item("SP_PARM_ASP_ACCT_CODE_EXP") & ""
                        ACCT_CODE = "4243"
                        '    If ASCMAIN1.Running_in_VS Then Stop ' REALLY SHOULD BE USING THE ASP_CODE TO GET THE EXPENSE, BUT WE ARE REALLY SPLITTING HAIRS HERE.
                        rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
                        If rowSOTTCLS1 IsNot Nothing Then
                            If rowSOTTCLS1.Item("SEG3_CODE") & "" <> "" Then
                                rowGLTINTF1.Item("SEG3_CODE") = rowSOTTCLS1.Item("SEG3_CODE")
                            Else
                                rowGLTINTF1.Item("SEG3_CODE") = TRADE_CLASS_CODE
                            End If
                        Else
                            rowGLTINTF1.Item("SEG3_CODE") = "?"
                        End If
                        If rowICTCOLL1 IsNot Nothing Then
                            If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then
                                rowGLTINTF1.Item("SEG4_CODE") = rowICTCOLL1.Item("SEG4_CODE")
                            Else
                                rowGLTINTF1.Item("SEG4_CODE") = COLLECTION_CODE
                            End If
                        Else
                            rowGLTINTF1.Item("SEG4_CODE") = "?"
                        End If
                        If rowGLTINTF1.Item("SEG2_CODE") & "" = "" Then rowGLTINTF1.Item("SEG2_CODE") = "?"
                        If rowGLTINTF1.Item("SEG3_CODE") & "" = "" Then rowGLTINTF1.Item("SEG3_CODE") = "?"
                        If rowGLTINTF1.Item("SEG4_CODE") & "" = "" Then rowGLTINTF1.Item("SEG4_CODE") = "?"
                        rowGLTINTF1.Item("DETL_CVX_NO") = COLLECTION_CODE
                        rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"
                    End If
                    
                End If
            Next

            ' Accrual
            ' Dim ASP_COMM_OFFSET As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(ASP_COMM_OFFSET)", "") & "")
            Dim AMT_COMM_OFFSET As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(AMT_COMM_OFFSET)", "OPS_YYYYPP_PAID = '" & OPS_YYYYPP & "'") & "")
            If AMT_COMM_OFFSET <> 0 Then
                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, AMT_COMM_OFFSET)
                rowGLTINTF1.Item("OPS_YYYYPP") = OPS_YYYYPP
                rowGLTINTF1.Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_ASP_ACCT_CODE_ACC")
            End If

            ' AR

            Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ROWs("SPTPARM1").Item("SP_PARM_ASP_TYPE_CODE"))
            Dim rowARTPOST1 As DataRow = dst.Tables("ARTPOST1").Rows.Find(rowSOTTYPE1.Item("POST_CODE"))
            Dim AMT_COMM_PAID As Decimal = Val(dst.Tables("SPTACOMC").Compute("SUM(AMT_COMM_PAID)", "OPS_YYYYPP_PAID = '" & OPS_YYYYPP & "'") & "")
            If AMT_COMM_PAID <> 0 Then
                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, -1 * AMT_COMM_PAID)
                rowGLTINTF1.Item("OPS_YYYYPP") = OPS_YYYYPP
                rowGLTINTF1.Item("ACCT_CODE") = rowARTPOST1.Item("ACCT_CODE")
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

        ASCMAIN1.sql = "Select * from SPTAPMT1 where ROWNUM < 1"
        SPTAPMT1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SPTAPMT1 & " Add Primary Key (PYMT_NO)")

    End Sub
End Class