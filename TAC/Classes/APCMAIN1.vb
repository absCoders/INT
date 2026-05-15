Public Class APCMAIN1
    Public Shared Function Prepare_Check_Register(
    ByRef f As ASFSRPTM,
    ByRef dst As DataSet,
    ByVal check_register As Boolean,
    Optional ByVal RYP1 As String = "", Optional ByVal RYP2 As String = "",
    Optional ByVal DATE1 As Date = Nothing, Optional ByVal DATE2 As Date = Nothing) As String

        Dim sql As String
        Dim APTCHKR1 As String

        With dst
            sql = "Select 'I' RECORD_TYPE, APTCHCK1.* from APTCHCK1 "
            If check_register Then
                sql = sql & " where APTCHCK1.REGISTER_IND = '0'"
            Else
                If RYP1 = "" Then
                    sql = sql & " where APTCHCK1.CHECK_DATE >= '" & Format$(DATE1, "dd-MMM-yyyy") & "'"
                    sql = sql & "   and APTCHCK1.CHECK_DATE <= '" & Format$(DATE2, "dd-MMM-yyyy") & "'"
                Else
                    If RYP1 = RYP2 Or RYP2 = "" Then
                        sql = sql & " where APTCHCK1.OPS_YYYYPP = '" & RYP1 & "'"
                    Else
                        sql = sql & " where APTCHCK1.OPS_YYYYPP >= '" & RYP1 & "'"
                        sql = sql & "   and APTCHCK1.OPS_YYYYPP <= '" & RYP2 & "'"
                    End If
                End If
                sql = sql & f.SQL_in("BANK_CODE", "APTCHCK1.BANK_CODE")
                sql = sql & f.SQL_in("VEND_CODE_AP", "APTCHCK1.VEND_CODE_AP")
            End If

            APTCHKR1 = ASCMAIN1.Temp_Table(sql)

            sql = "Select 'V' RECORD_TYPE, APTCHCK1.* from APTCHCK1 "
            If check_register Then
                sql = sql & "where APTCHCK1.REGISTER_IND_F = '0'"
            Else

                If RYP1 = "" Then
                    sql = sql & " where APTCHCK1.CHECK_DATE >= '" & Format$(DATE1, "dd-MMM-yyyy") & "'"
                    sql = sql & "   and APTCHCK1.CHECK_DATE <= '" & Format$(DATE2, "dd-MMM-yyyy") & "'"
                Else
                    If RYP1 = RYP2 Or RYP2 = "" Then
                        sql = sql & " where APTCHCK1.OPS_YYYYPP_F = '" & RYP1 & "'"
                    Else
                        sql = sql & " where APTCHCK1.OPS_YYYYPP_F >= '" & RYP1 & "'"
                        sql = sql & "   and APTCHCK1.OPS_YYYYPP_F <= '" & RYP2 & "'"
                    End If
                End If
                sql = sql & " and (APTCHCK1.CHECK_STATUS = 'V' OR APTCHCK1.CHECK_STATUS = 'R')"
                sql = sql & f.SQL_in("BANK_CODE", "APTCHCK1.BANK_CODE")
                sql = sql & f.SQL_in("VEND_CODE_AP", "APTCHCK1.VEND_CODE_AP")
            End If

            ASCDATA1.ExecuteSQL("Insert into " & APTCHKR1 & " " & sql)
            ASCDATA1.ExecuteSQL("Alter table " & APTCHKR1 & " Add Primary Key (RECORD_TYPE, BANK_CODE, CHECK_NUM)")

            sql = "Select * from " & APTCHKR1
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTCHKR1", 3))

            sql = "Select APTCHKR1.RECORD_TYPE, APTCHCK2.*, APTINVH1.POST_CODE " _
            & ", APTCHKR1.OPS_YYYYPP, APTCHKR1.OPS_YYYYPP_F " _
            & " from APTCHCK2,APTINVH1," & APTCHKR1 & " APTCHKR1" _
            & " where APTCHKR1.BANK_CODE = APTCHCK2.BANK_CODE " _
            & "   and APTCHKR1.CHECK_NUM = APTCHCK2.CHECK_NUM " _
            & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTCHKR2", 4))


            sql = "Select APTINVH1.* from APTINVH1 " _
            & " where VOUCHER_NO in (Select APTINVH1.VOUCHER_NO" _
            & " from APTCHCK2,APTINVH1," & APTCHKR1 & " APTCHKR1" _
            & " where APTCHKR1.BANK_CODE = APTCHCK2.BANK_CODE " _
            & "   and APTCHKR1.CHECK_NUM = APTCHCK2.CHECK_NUM " _
            & "   and APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO)"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTINVH1", 1))

            .Tables.Add(ASCDATA1.GetDataTable("*", "GLTBANK1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "APTPOST1"))

            Call f.Create_TDA(.Tables.Add, "GLTINTF1", "*")
        End With

        For Each rowAPTCHKR1 As DataRow In dst.Tables("APTCHKR1").Select("RECORD_TYPE = 'V'", "")
            rowAPTCHKR1.Item("CHECK_AMT") = -1 * Val(rowAPTCHKR1.Item("CHECK_AMT") & "")
        Next
        For Each rowAPTCHKR2 As DataRow In dst.Tables("APTCHKR2").Select("RECORD_TYPE = 'V'", "")
            rowAPTCHKR2.Item("INV_AMT_APPLIED") = -1 * Val(rowAPTCHKR2.Item("INV_AMT_APPLIED") & "")
            rowAPTCHKR2.Item("INV_DISC_TAKEN") = -1 * Val(rowAPTCHKR2.Item("INV_DISC_TAKEN") & "")
        Next

        If check_register Then

            ' Prepare GL Interface File

            Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
            Dim JOURNAL_TYPE As String = "APCD"
            Dim JOURNAL_LNO As Integer = 0

            Dim DETL_POSTING_AMT As Double
            Dim BANK_CODE As String
            Dim POST_CODE As String
            Dim OPS_YYYYPP As String

            Call f.Summary_Table("APTCHKRB", "APTCHKR1",
            "BANK_CODE,RECORD_TYPE,OPS_YYYYPP,OPS_YYYYPP_F",
            "CHECK_AMT")

            For Each rowAPTCHKRB As DataRow In dst.Tables("APTCHKRB").Rows
                BANK_CODE = rowAPTCHKRB("BANK_CODE")
                Dim rowGLTBANK1 As DataRow = dst.Tables("GLTBANK1").Rows.Find(BANK_CODE)
                If rowAPTCHKRB("RECORD_TYPE") = "I" Then
                    OPS_YYYYPP = rowAPTCHKRB("OPS_YYYYPP")
                Else
                    OPS_YYYYPP = rowAPTCHKRB("OPS_YYYYPP_F")
                End If
                DETL_POSTING_AMT = -1 * Val(rowAPTCHKRB("CHECK_AMT") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = rowGLTBANK1("ACCT_CODE")
                    rowGLTINTF1("SEG2_CODE") = IIf(rowGLTBANK1("SEG2_CODE") & "" = "", f.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2"), rowGLTBANK1("SEG2_CODE"))
                    rowGLTINTF1("SEG3_CODE") = IIf(rowGLTBANK1("SEG3_CODE") & "" = "", f.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3"), rowGLTBANK1("SEG3_CODE"))
                    rowGLTINTF1("SEG4_CODE") = IIf(rowGLTBANK1("SEG4_CODE") & "" = "", f.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4"), rowGLTBANK1("SEG4_CODE"))
                    rowGLTINTF1("DETL_CTL_DATE") = Format(f.DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Math.Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = f.XNO
                    rowGLTINTF1("DETL_CVX_NO") = BANK_CODE
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If
            Next

            Call f.Summary_Table("APTCHKRP", "APTCHKR2",
            "POST_CODE,RECORD_TYPE,OPS_YYYYPP,OPS_YYYYPP_F",
            "INV_AMT_APPLIED,INV_DISC_TAKEN")

            For Each rowAPTCHKRP As DataRow In dst.Tables("APTCHKRP").Rows
                POST_CODE = rowAPTCHKRP("POST_CODE")
                Dim rowAPTPOST1 As DataRow = dst.Tables("APTPOST1").Rows.Find(POST_CODE)
                If rowAPTCHKRP("RECORD_TYPE") = "I" Then
                    OPS_YYYYPP = rowAPTCHKRP("OPS_YYYYPP")
                Else
                    OPS_YYYYPP = rowAPTCHKRP("OPS_YYYYPP_F")
                End If
                DETL_POSTING_AMT = Val(rowAPTCHKRP("INV_AMT_APPLIED") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = rowAPTPOST1("ACCT_CODE")
                    rowGLTINTF1("SEG2_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG2")
                    rowGLTINTF1("SEG3_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG3")
                    rowGLTINTF1("SEG4_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG4")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(f.DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Math.Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = f.XNO
                    rowGLTINTF1("DETL_CVX_NO") = POST_CODE
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If

                DETL_POSTING_AMT = -1 * Val(rowAPTCHKRP("INV_DISC_TAKEN") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = f.ROWs("APTPARM1")("AP_PARM_ACCT_CODE_DISC")
                    rowGLTINTF1("SEG2_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG2")
                    rowGLTINTF1("SEG3_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG3")
                    rowGLTINTF1("SEG4_CODE") = f.ROWs("GLTPARM1")("GL_PARM_DEF_SEG4")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(f.DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Math.Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = f.XNO
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If
            Next
        End If

        Return APTCHKR1
    End Function

    Public Shared Function check_Bank_Payment_Method(frmASFBASE0 As ASFBASE0, BANK_CODE As String, PYMT_METHOD As String) As Boolean

        If frmASFBASE0.ROWs("APTPARM1").Item("AP_PARM_BANK_METHOD") & "" <> "1" Then
            Return True
        End If

        Dim row As DataRow = frmASFBASE0.LookUp("GLTBANK2", New String() {BANK_CODE, PYMT_METHOD})
        If row IsNot Nothing Then
            Return True
        Else
            Return False
        End If

        'If BANK_CODE = "WIRE" Then
        '    If PYMT_METHOD = "ACH" Or PYMT_METHOD = "WIRE" Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'Else
        '    If PYMT_METHOD = "CHECK" Then
        '        Return True
        '    Else
        '        Return False
        '    End If
        'End If

    End Function
End Class