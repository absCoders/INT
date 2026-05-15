Imports System.Math

Public Class APRINVR1

    Dim APTINVR1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        Dim sqlw As String = " from APTINVH1,APTVEND1 where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE and APTINVH1.REGISTER_IND = '0' and APTINVH1.INV_STATUS <> 'R'"

        If Absx1.chkFor("MY_RECORDS_ONLY").Checked Then
            sqlw = sqlw & " and (APTINVH1.INIT_OPER = '" & ASCMAIN1.USER_ID & "' or APTINVH1.LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
        End If
        sqlw &= SQL_in("VOUCHER_NO", "APTINVH1.VOUCHER_NO")
        sqlw &= SQL_in("VEND_CODE", "APTINVH1.VEND_CODE")
        sqlw &= SQL_in("INIT_OPER", "APTVEND1.PROCESSOR_CODE")

        sql = "Select APTINVH1.VOUCHER_NO, APTINVH1.BANK_CODE, APTINVH1.CHECK_NUM, APTINVH1.INV_STATUS " & sqlw
        APTINVR1 = ASCMAIN1.Temp_Table(sql)
        sql = "Alter Table " & APTINVR1 & " Add Primary Key (VOUCHER_NO)"
        ASCDATA1.ExecuteSQL(sql)
        sql = "Create Index I_" & APTINVR1 & "_1 ON " & APTINVR1 & " (BANK_CODE, CHECK_NUM)"
        ASCDATA1.ExecuteSQL(sql)
        Call ASCMAIN1.AnalyzeTable(APTINVR1)

        sqlw = " X, " & APTINVR1 & " APTINVR1 where APTINVR1.VOUCHER_NO = X.VOUCHER_NO"

        With dst
            sql = "Select X.* from APTINVH1" & sqlw
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTINVR1", 1))
            .Tables("APTINVR1").Columns.Add("CHECK_AMT", GetType(System.Double))
            .Tables("APTINVR1").Columns.Add("CHECK_AMT_OTHERS", GetType(System.Double))
            .Tables("APTINVR1").Columns.Add("INV_AMT_GL", GetType(System.Double))

            sql = "Select X.*, DECODE(X.INV_LTYP,NULL,X.INV_LINE_AMT,0) INV_LINE_AMT_GL, GLTACCT1.ACCT_DESC from GLTACCT1, APTINVH2" & sqlw & " and GLTACCT1.ACCT_CODE = X.ACCT_CODE"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTINVR2", 2))
            .Tables("APTINVR2").Columns.Add("OPS_YYYYPP", GetType(System.String))

            .Relations.Add("APTINVR2", _
            .Tables("APTINVR1").Columns("VOUCHER_NO"), _
            .Tables("APTINVR2").Columns("VOUCHER_NO"))

            .Tables("APTINVR1").Columns("INV_AMT_GL").Expression = "SUM(CHILD(APTINVR2).INV_LINE_AMT_GL)"
            .Tables("APTINVR2").Columns("OPS_YYYYPP").Expression = "PARENT(APTINVR2).OPS_YYYYPP"

            sql = "Select X.* from APTINVH5" & sqlw
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTINVR5", 2))

            .Relations.Add("APTINVR5", _
            .Tables("APTINVR1").Columns("VOUCHER_NO"), _
            .Tables("APTINVR5").Columns("VOUCHER_NO"))


            sqlw = " X,(SELECT DISTINCT BANK_CODE, CHECK_NUM " _
            & " from " & APTINVR1 & ") APTINVR1 " _
            & " where APTINVR1.BANK_CODE = X.BANK_CODE " _
            & "   and APTINVR1.CHECK_NUM = X.CHECK_NUM"

            sql = "Select X.* from APTCHCK1" & sqlw
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTCHCK1", 2))

            sql = "Select X.* from APTCHCK2" & sqlw
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTCHCK2", 3))

            sql = "Select APTVEND1.* from APTVEND1 where VEND_CODE in " _
            & "(Select Distinct VEND_CODE from " & APTINVR1 & ")"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTVEND1", 1))

            .Tables.Add(ASCDATA1.GetDataTable("*", "GLTBANK1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "GLTACCT1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "APTPOST1"))

            Create_TDA(.Tables.Add, "GLTINTF1", "*")
        End With

        For Each rowAPTINVR1 As DataRow In dst.Tables("APTINVR1").Rows
            rowAPTINVR1("CHECK_AMT") = 0
            rowAPTINVR1("CHECK_AMT_OTHERS") = 0
        Next
        'sql = "Update APWINVH1,APWCHCK1 set APWINVH1.CHECK_AMT = APWCHCK1.CHECK_AMT"
        'sql = sql & " where APWCHCK1.BANK_CODE = APWINVH1.BANK_CODE"
        'sql = sql & "   and APWCHCK1.CHECK_NUM = APWINVH1.CHECK_NUM"
        'sql = sql & "   and APWINVH1.INV_PAID_UPON_ENTRY = '1'"
        'AccD.Execute(sql)


        ' Check to see if Accrual Feature was turned off, or if a GL Period has been closed since entry

        Dim GL_PARM_CURRENT_YYYYPP As String = ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & ""
        Dim AP_PARM_ALLOW_ACCRUAL As String = ROWs("APTPARM1").Item("AP_PARM_ALLOW_ACCRUAL") & ""
        Dim YP As String = GL_PARM_CURRENT_YYYYPP
        If AP_PARM_ALLOW_ACCRUAL = "2" Then
            YP = ASCMAIN1.Period_Calc(GL_PARM_CURRENT_YYYYPP, -1)
        ElseIf AP_PARM_ALLOW_ACCRUAL = "1" Then
            YP = GL_PARM_CURRENT_YYYYPP
        Else
            YP = ""
        End If
        Dim VOUCHER_ACCRUAL_ERRORS As New List(Of String)
        For Each rowAPTINVR1 As DataRow In dst.Tables("APTINVR1").Select("OPS_YYYYPP_ACCRUE is Not Null", "")
            If YP = "" Then
                rowAPTINVR1("OPS_YYYYPP_ACCRUE") = Null
                VOUCHER_ACCRUAL_ERRORS.Add(rowAPTINVR1("VOUCHER_NO"))
            End If
            If rowAPTINVR1("OPS_YYYYPP_ACCRUE") & "" < YP Then
                If rowAPTINVR1("OPS_YYYYPP") = YP Then
                    rowAPTINVR1("OPS_YYYYPP_ACCRUE") = Null
                Else
                    rowAPTINVR1("OPS_YYYYPP_ACCRUE") = YP
                End If

                VOUCHER_ACCRUAL_ERRORS.Add(rowAPTINVR1("VOUCHER_NO"))
            End If
        Next
        If VOUCHER_ACCRUAL_ERRORS.Count > 0 Then
            If MsgBox("The Accrual feature has encountered a restriction" _
            & vbCr & " causing the Accrual Period to be either changed" _
            & vbCr & " or cleared for at least 1 Voucher (" & VOUCHER_ACCRUAL_ERRORS(0) & ")" _
            & vbCr & vbCr _
            & "Continue Anyway?", MsgBoxStyle.YesNo, "Voucher Accrual Error") = MsgBoxResult.No Then
                RWU = "0"
                xErrMsg = "This execution has been Cancelled"
                Exit Sub
            End If
        End If

        Call APIN_GL()

        Check_if_Empty("APTINVR1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT)
        Call Print_GL()
    End Sub

    Overrides Sub Update_Record()

        Dim sql As String

        sql = "Update APTINVH1 " _
        & " Set REGISTER_IND = '1', REGISTER_XNO = '" & XNO & "'" _
        & " where VOUCHER_NO in (Select VOUCHER_NO from " & APTINVR1 & " )"
        ASCDATA1.ExecuteSQL(sql)

        GL_Update()

        If ASCMAIN1.DBS_SERVER = "EXP" Or ASCMAIN1.DBS_COMPANY = "EXP" Then
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("GLTINTF1"), New String() {"JOURNAL_NO"}).Select("")
                Dim JOURNAL_NO As String = row.Item(0)
                ASCMAIN1.sql = "Insert into GLTDETL1_OBX Select GLTDETL1.*, NULL DATETIME_STAMP, 'APIN' JOURNAL_TYPE from GLTDETL1 where JOURNAL_NO = '" & JOURNAL_NO & "'"
                ASCDATA1.ExecuteSQL()
            Next
        End If

    End Sub

    Sub APIN_GL() '(JOURNAL_NO As String, Optional JOURNAL_TYPE As String)


        '' GL Parameters

        'Dim GL_PARM_CURRENT_YYYYPP As String
        'Dim GL_PARM_DEF_SEG2 As String
        'Dim GL_PARM_DEF_SEG3 As String
        'Dim GL_PARM_DEF_SEG4 As String
        'GL_PARM_CURRENT_YYYYPP = dynGLTPARM1.Fields("GL_PARM_CURRENT_YYYYPP").Value
        'GL_PARM_DEF_SEG2 = dynGLTPARM1.Fields("GL_PARM_DEF_SEG2").Value
        'GL_PARM_DEF_SEG3 = dynGLTPARM1.Fields("GL_PARM_DEF_SEG3").Value
        'GL_PARM_DEF_SEG4 = dynGLTPARM1.Fields("GL_PARM_DEF_SEG4").Value


        '' Load Work Files

        'sql = "Select * from GLTDIST1"
        'Call Ora_to_Acc(Nothing, "GLWDIST1", 1, "", sql)
        'Dim tblGLWDIST1 As Recordset
        'tblGLWDIST1 = AccD.OpenRecordset("GLWDIST1", dbOpenTable)
        'tblGLWDIST1.Index = "PrimaryKey"
        'sql = "Select * from GLTDIST2"
        'Call Ora_to_Acc(Nothing, "GLWDIST2", 3, "", sql)

        'Dim dynGLWDIST2 As Recordset
        'dynGLWDIST2 = AccD.OpenRecordset("GLWDIST2", dbOpenDynaset)
        'sql = "SELECT GLTDIST2.DIST_CODE"
        'sql = sql & " , SUM (DECODE (GLTDIST2.ACCT_SEG_ID,'2',1,0)) SEG2_COUNT"
        'sql = sql & " , SUM (DECODE (GLTDIST2.ACCT_SEG_ID,'3',1,0)) SEG3_COUNT"
        'sql = sql & " , SUM (DECODE (GLTDIST2.ACCT_SEG_ID,'4',1,0)) SEG4_COUNT"
        'sql = sql & " From GLTDIST2, GLTDIST1"
        'sql = sql & " WHERE GLTDIST1.DIST_CODE = GLTDIST2.DIST_CODE AND GLTDIST1.DIST_METHOD = 'M'"
        'sql = sql & " GROUP BY GLTDIST2.DIST_CODE"
        'Dim dynDIST As OraDynaset
        'dynDIST = OraD.CreateDynaset(sql, 8&)
        'Do While Not dynDIST.EOF
        '    For i = 2 To 4
        '        If dynDIST.Fields("SEG" & Format$(i, "0") & "_COUNT").Value = 0 Then
        '            dynGLWDIST2.AddNew()
        '            dynGLWDIST2.Fields("DIST_CODE").Value = dynDIST.Fields("DIST_CODE").Value
        '            dynGLWDIST2.Fields("ACCT_SEG_ID").Value = Format$(i, "0")
        '            dynGLWDIST2.Fields("ACCT_SEG_CODE").Value = dynGLTPARM1.Fields("GL_PARM_DEF_SEG" & Format$(i, "0")).Value
        '            dynGLWDIST2.Fields("DIST_PCT").Value = 100
        '            dynGLWDIST2.Update()
        '        End If
        '    Next i
        '    dynDIST.MoveNext()
        'Loop
        'dynDIST.Close()
        'dynGLWDIST2.Close()

        'For i = 2 To 4
        '    sql = ""
        '    sql = sql & " SELECT DIST_CODE, ACCT_SEG_CODE AS SEG" & Format$(i, "0") & "_CODE, DIST_PCT AS DIST_PCT" & Format$(i, "0")
        '    sql = sql & " INTO GLWDIST2_" & Format$(i, "0") & " FROM GLWDIST2 WHERE ACCT_SEG_ID = '" & Format$(i, "0") & "'"
        '    AccD.Execute(sql)
        '    Call Create_Index("GLWDIST2_" & Format$(i, "0"), "PrimaryKey", "DIST_CODE,SEG" & Format$(i, "0") & "_CODE")
        'Next i

        'sql = "Select * from GLTDIST4"
        'Call Ora_to_Acc(Nothing, "GLWDIST4", 2, "", sql)
        'Dim dynGLWDIST4 As Recordset







        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_TYPE As String = "APIN"
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Double
        Dim DETL_CTL_DATE As Date
        DETL_CTL_DATE = DateValue(Format(DATETIME_STAMP, "MM/dd/yyyy"))

        ' Distributions

        'Dim DIST_PCT As Double
        'Dim DIST_CODE As String
        'Dim SEG2_CODE As String
        'Dim SEG3_CODE As String
        'Dim SEG4_CODE As String
        'sql = "Select * from GLTDIST4 where DIST_CODE = :CODE"
        'Dim dynGLTDIST4 As OraDynaset

        'i = 0
        'sql = "Select IIf (APWINVH2.OPS_YYYYPP is Null, APWINVH1.OPS_YYYYPP,APWINVH2.OPS_YYYYPP), "
        'sql = sql & " '" & JOURNAL_NO & "', 0, APWINVH2.ACCT_CODE, "
        'sql = sql & " APWINVH2.SEG2_CODE, APWINVH2.SEG3_CODE, APWINVH2.SEG4_CODE, "
        'sql = sql & " APWINVH1.INV_POSTING_DATE, APWINVH2.VOUCHER_NO, APWINVH2.VOUCHER_LNO,"
        'sql = sql & "'" & xXNO & "', APWINVH2.INV_LINE_AMT, '','APIN',"
        'sql = sql & " APWINVH1.VEND_CODE, APWINVH1.INV_DATE, APWINVH1.INV_NUM, 0, APWINVH2.DIST_CODE"
        'sql = sql & " from APWINVH1, APWINVH2"
        'sql = sql & " where APWINVH1.VOUCHER_NO = APWINVH2.VOUCHER_NO"
        'sql = sql & "   and APWINVH2.INV_LINE_AMT <> 0"
        'sql = sql & "   and APWINVH2.INV_LTYP is Null"
        ''sql = sql & "   and APWINVH1.REGISTER_IND is Null"
        'sql = sql & " order by APWINVH2.ACCT_CODE, "
        'sql = sql & " APWINVH2.SEG2_CODE, APWINVH2.SEG3_CODE, APWINVH2.SEG4_CODE "
        'dynwk = AccD.OpenRecordset(sql, dbOpenForwardOnly)
        'Do While Not dynwk.EOF
        '    DIST_CODE = dynwk.Fields("DIST_CODE").Value & ""
        '    If DIST_CODE = "" Then
        '    GoSub Write_INTF1
        '    Else
        '        tblGLWDIST1.Seek("=", DIST_CODE)
        '        If tblGLWDIST1.Fields("DIST_METHOD").Value = "S" Then
        '            sql = "Select SEG2_CODE, SEG3_CODE, SEG4_CODE, DIST_PCT"
        '            sql = sql & " from GLWDIST4 where DIST_CODE = '" & DIST_CODE & "'"
        '        Else
        '            sql = "SELECT SEG2_CODE, SEG3_CODE, SEG4_CODE, "
        '            sql = sql & " DIST_PCT2 * DIST_PCT3 * DIST_PCT4 / (100 * 100) AS DIST_PCT"
        '            sql = sql & " From GLWDIST2_2, GLWDIST2_3, GLWDIST2_4"
        '            sql = sql & " WHERE GLWDIST2_2.DIST_CODE = '" & DIST_CODE & "'"
        '            sql = sql & "   AND GLWDIST2_3.DIST_CODE = '" & DIST_CODE & "'"
        '            sql = sql & "   AND GLWDIST2_4.DIST_CODE = '" & DIST_CODE & "'"
        '        End If

        '        dynDISTwk = AccD.OpenRecordset(sql, dbOpenForwardOnly)
        '        Do While Not dynDISTwk.EOF
        '            SEG2_CODE = dynDISTwk.Fields("SEG2_CODE").Value
        '            SEG3_CODE = dynDISTwk.Fields("SEG3_CODE").Value
        '            SEG4_CODE = dynDISTwk.Fields("SEG4_CODE").Value
        '            DIST_PCT = Val(dynDISTwk.Fields("DIST_PCT").Value & "")
        '        GoSub Write_INTF1
        '            dynDISTwk.MoveNext()
        '        Loop
        '        dynDISTwk.Close()
        '    End If
        '    dynwk.MoveNext()
        'Loop
        'dynwk.Close()

        'DIST_CODE = ""



        Dim VOUCHER_NO As String = ""
        Dim YP As String = ""
        Dim rowAPTINVR1 As DataRow = Nothing


        ' GL Distributions - INV_LTYP is Null

        For Each rowAPTINVR2 As DataRow In dst.Tables("APTINVR2").Select("INV_LTYP is Null", "VOUCHER_NO")
            DETL_POSTING_AMT = Val(rowAPTINVR2("INV_LINE_AMT") & "")
            If DETL_POSTING_AMT <> 0 Then

                If VOUCHER_NO <> rowAPTINVR2("VOUCHER_NO") Then
                    VOUCHER_NO = rowAPTINVR2("VOUCHER_NO")
                    rowAPTINVR1 = dst.Tables("APTINVR1").Rows.Find(VOUCHER_NO)
                    If rowAPTINVR1("OPS_YYYYPP_ACCRUE") & "" <> "" Then
                        YP = rowAPTINVR1("OPS_YYYYPP_ACCRUE") & ""
                    Else
                        YP = rowAPTINVR1("OPS_YYYYPP")
                    End If
                End If

                Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                rowGLTINTF1("OPS_YYYYPP") = YP
                rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                JOURNAL_LNO += 1
                rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                rowGLTINTF1("ACCT_CODE") = rowAPTINVR2("ACCT_CODE")
                rowGLTINTF1("SEG2_CODE") = rowAPTINVR2("SEG2_CODE")
                rowGLTINTF1("SEG3_CODE") = rowAPTINVR2("SEG3_CODE")
                rowGLTINTF1("SEG4_CODE") = rowAPTINVR2("SEG4_CODE")
                rowGLTINTF1("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                rowGLTINTF1("DETL_EXE_NO") = XNO
                rowGLTINTF1("DETL_CTL_NO") = VOUCHER_NO
                rowGLTINTF1("DETL_CTL_LNO") = rowAPTINVR2("VOUCHER_LNO")
                rowGLTINTF1("DETL_CVX_NO") = rowAPTINVR1("VEND_CODE")
                rowGLTINTF1("DETL_CVX_REF_DATE") = rowAPTINVR1("INV_DATE")
                rowGLTINTF1("DETL_CVX_REF_NO") = rowAPTINVR1.Item("INV_NUM")
                If rowAPTINVR2("INV_COMMENT_DTL") & "" = "" Then
                    rowGLTINTF1("DETL_DESC") = rowAPTINVR1.Item("INV_REF")
                Else
                    rowGLTINTF1("DETL_DESC") = rowAPTINVR2.Item("INV_COMMENT_DTL")
                End If
                rowGLTINTF1("DETL_CVX_TYPE") = "V"
                rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                rowGLTINTF1("COST_CTR_CODE") = rowAPTINVR2("COST_CTR_CODE")

                Dim DIST_CODE As String = rowAPTINVR2.Item("DIST_CODE") & ""
                If DIST_CODE = "" Then
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                Else
                    If Not dst.Tables.Contains("GLTDIST1") Then
                        Create_TDA(dst.Tables.Add, "GLTDIST1", "*", 0)
                        Fill_Records("GLTDIST1")
                    End If
                    Dim rowGLTDIST1 As DataRow = dst.Tables("GLTDIST1").Rows.Find(DIST_CODE)
                    For Each rowGLTDIST2 As DataRow In dst.Tables("GLTDIST2").Select("DIST_CODE = '" & DIST_CODE & "'")

                    Next
                End If

            End If
        Next


        ' GL Accrual Offsets

        Call Summary_Table("APTINVRO", "APTINVR2", _
        "INV_LTYP,ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE", _
        "INV_LINE_AMT")

        For Each rowAPTINVRO As DataRow In dst.Tables("APTINVRO").Rows
            Dim INV_LTYP = rowAPTINVRO("INV_LTYP") & ""
            If INV_LTYP <> "" Then
                DETL_POSTING_AMT = Val(rowAPTINVRO("INV_LINE_AMT") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = ASCMAIN1.CYP ' rowAPTINVRO("OPS_YYYYPP") - trouble using this column in a relation when setting up the summary table
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                    rowGLTINTF1("ACCT_CODE") = rowAPTINVRO("ACCT_CODE")
                    rowGLTINTF1("SEG2_CODE") = rowAPTINVRO("SEG2_CODE")
                    rowGLTINTF1("SEG3_CODE") = rowAPTINVRO("SEG3_CODE")
                    rowGLTINTF1("SEG4_CODE") = rowAPTINVRO("SEG4_CODE")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = XNO
                    rowGLTINTF1("DETL_CVX_NO") = rowAPTINVRO("INV_LTYP")
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If
            End If
        Next


        ' AP Posting Control Account

        Call Summary_Table("APTINVRP", "APTINVR1", _
        "OPS_YYYYPP,POST_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE", _
        "INV_AMT")

        For Each rowAPTINVRP As DataRow In dst.Tables("APTINVRP").Rows
            Dim POST_CODE As String = rowAPTINVRP("POST_CODE")
            Dim rowAPTPOST1 As DataRow = dst.Tables("APTPOST1").Rows.Find(POST_CODE)
            DETL_POSTING_AMT = Val(rowAPTINVRP("INV_AMT") & "")
            If DETL_POSTING_AMT <> 0 Then
                Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
                rowGLTINTF1("OPS_YYYYPP") = rowAPTINVRP("OPS_YYYYPP")
                rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                JOURNAL_LNO += 1
                rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
                rowGLTINTF1("ACCT_CODE") = rowAPTPOST1("ACCT_CODE")
                rowGLTINTF1("SEG2_CODE") = rowAPTINVRP("SEG2_CODE")
                rowGLTINTF1("SEG3_CODE") = rowAPTINVRP("SEG3_CODE")
                rowGLTINTF1("SEG4_CODE") = rowAPTINVRP("SEG4_CODE")
                rowGLTINTF1("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                rowGLTINTF1("DETL_POSTING_AMT") = Round(-1 * DETL_POSTING_AMT, 2)
                rowGLTINTF1("DETL_EXE_NO") = XNO
                rowGLTINTF1("DETL_CVX_NO") = POST_CODE
                rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
            End If
        Next

        ' Accrued AP

        Call Summary_Table("APTINVRA", "APTINVR1", _
        "OPS_YYYYPP,OPS_YYYYPP_ACCRUE,POST_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE", _
        "INV_AMT_GL")

        For Each rowAPTINVRA As DataRow In dst.Tables("APTINVRA").Rows
            If rowAPTINVRA("OPS_YYYYPP_ACCRUE") & "" <> "" Then
                DETL_POSTING_AMT = Val(rowAPTINVRA("INV_AMT_GL") & "")
                If DETL_POSTING_AMT <> 0 Then
                    Dim ACCT_CODE As String = ROWs("APTPARM1").Item("AP_PARM_ACCT_CODE_ACCRUED_AP")
                    Dim POST_CODE As String = rowAPTINVRA("POST_CODE") & ""
                    Dim rowAPTPOST1 As DataRow = LookUp("APTPOST1", POST_CODE)
                    If rowAPTPOST1 IsNot Nothing AndAlso rowAPTPOST1.Item("ACCT_CODE_ACCRUAL") & "" <> "" Then
                        ACCT_CODE = rowAPTPOST1.Item("ACCT_CODE_ACCRUAL")
                    End If

                    Dim rowGLTINTF1 As DataRow

                    rowGLTINTF1 = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = rowAPTINVRA("OPS_YYYYPP_ACCRUE")
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO

                    rowGLTINTF1("ACCT_CODE") = ACCT_CODE
                    rowGLTINTF1("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Round(-1 * DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = XNO
                    rowGLTINTF1("DETL_CVX_NO") = ""
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)

                    rowGLTINTF1 = dst.Tables("GLTINTF1").NewRow
                    rowGLTINTF1("OPS_YYYYPP") = rowAPTINVRA("OPS_YYYYPP")
                    rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
                    JOURNAL_LNO += 1
                    rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO

                    rowGLTINTF1("ACCT_CODE") = ACCT_CODE
                    rowGLTINTF1("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                    rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                    rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    rowGLTINTF1("DETL_CTL_DATE") = Format(DATETIME_STAMP, "MM/dd/yyyy")
                    rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
                    rowGLTINTF1("DETL_EXE_NO") = XNO
                    rowGLTINTF1("DETL_CVX_NO") = ""
                    rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
                    dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
                End If
            End If
        Next

    End Sub
End Class