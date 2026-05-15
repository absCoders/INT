Imports System.Math

Public Class GLCMAIN1

    Public Shared Function Prepare_Work_File( _
    ByRef F As ASFBASE1, _
    ByRef TTA As String, _
    ByRef TTB As String, _
    ByRef GLTFINRD As String, _
    ByRef RY As String, _
    ByRef P As Integer, _
    ByRef SQLP() As String, _
    ByRef SQLF() As String, _
    ByRef sqlA_select As String, _
    ByRef sqlA_group_by As String, _
    ByRef sqlA_where As String, _
    ByRef BY_SEG2 As Boolean, _
    ByRef BY_SEG3 As Boolean, _
    ByRef BY_SEG4 As Boolean) As String

        Dim ACCT_TYPEs As New Dictionary(Of String, String)

        ACCT_TYPEs.Add("B", "('A','L','E')")
        ACCT_TYPEs.Add("I", "('I','X')")



        TTA = GL_Prep(F, Format(Val(RY) - 5, "0000"), Format(Val(RY) + 1, "0000"))
        TTB = GL_Prep(F, Format(Val(RY) - 5, "0000"), Format(Val(RY) + 1, "0000"), True)

        ' Clean out data which represents actuals beyond the Report Period Selected

        ASCDATA1.ExecuteSQL("Delete from " & TTA & " where ACCT_YEAR > '" & RY & "'")
        If P <> 12 Then
            Dim sql_clear As String = ""
            For i As Integer = P + 1 To 12
                sql_clear &= ", ACCT_ACT_P" & Format(i, "00") & " = 0"
            Next
            ASCDATA1.ExecuteSQL("Update " & TTA & " Set " & Mid(sql_clear, 2) & " where ACCT_YEAR = '" & RY & "'")
        End If

        Call Setup_SQL(F, SQLP, SQLF, TTA, TTB, sqlA_select, sqlA_group_by, sqlA_where, BY_SEG2, BY_SEG3, BY_SEG4)

        Dim sql As String = ""

        Dim A234 As String = "ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE"

        Dim A234T As String = A234 & ", ACCT_TYPE"
        Dim sqlC As String = _
            "Select Distinct " & A234T & " from " & TTA & " union " & _
            "Select Distinct " & A234T & " from " & TTB

        sql = "Select Distinct " & A234T & " from (" & sqlC & ")"
        Dim TTC As String = ASCMAIN1.Temp_Table(sql)

        ASCMAIN1.sql = "Select * from GLTFINR3 where ROWNUM < 1"
        GLTFINRD = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        Dim STMT_LINE_NO_ALL_ELSE As Integer = 0
        Dim STMT_LINE_NO As Integer = 0
        For Each row As DataRow In F.dst.Tables("GLTFINR2").Select("STMT_LINE_TYPE = 'D'", "STMT_LINE_NO", DataViewRowState.CurrentRows)
            STMT_LINE_NO = Val(row.Item("STMT_LINE_NO"))
            sql = ""
            Dim sqlx As String = "Select Distinct '" & F.HFs("STMT_CODE") & "' STMT_CODE, " & CStr(STMT_LINE_NO) & " STMT_LINE_NO, " & A234
            Dim sqlfrom As String = " from " & TTC

            Select Case row.Item("STMT_LINE_ACCTS") & ""
                Case "S"
                    Dim dvwGLTFINR3 As New DataView(F.dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
                    Dim z As String = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        z &= ",'" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                    Next
                    If z <> "" Then
                        sql = sqlx & sqlfrom & " where ACCT_CODE in (" & Mid(z, 2) & ")"
                        For s As Integer = 2 To 4
                            If row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "S" _
                            Or row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "X" Then
                                Dim dvwGLTFINR4 As New DataView(F.dst.Tables("GLTFINR4"))
                                dvwGLTFINR4.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO) & " and ACCT_SEG_ID = '" & CStr(s) & "'"
                                z = ""
                                For i As Integer = 0 To dvwGLTFINR4.Count - 1
                                    z &= ",'" & dvwGLTFINR4(i).Item("ACCT_SEG_CODE") & "'"
                                Next
                                If z <> "" Then
                                    sql &= " and SEG" & CStr(s) & "_CODE" _
                                        & IIf(row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "X", " NOT", "") _
                                        & " in (" & Mid(z, 2) & ")"
                                End If
                            End If
                        Next
                    End If

                Case "R"
                    sql = sqlx & sqlfrom & " where ACCT_CODE >= '" & row.Item("STMT_LINE_ACCT_RANGE1") & "' and ACCT_CODE <= '" & row.Item("STMT_LINE_ACCT_RANGE2") & "'"

                Case "I"
                    If F.HFs("STMT_TYPE") = "B" Then
                        sql = sqlx & sqlfrom & " where ACCT_TYPE in " & ACCT_TYPEs("I")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "B"
                    If F.HFs("STMT_TYPE") = "I" Then
                        sql = sqlx & sqlfrom & " where ACCT_TYPE in " & ACCT_TYPEs("B")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "X"
                    Dim dvwGLTFINR3 As New DataView(F.dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
                    Dim z As String = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        z &= ",('" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG2_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG3_CODE") & "'"
                        z &= ",'" & dvwGLTFINR3(i).Item("SEG4_CODE") & "')"
                    Next
                    If z <> "" Then
                        sql = sqlx & sqlfrom & " where (" & A234 & ") in (" & Mid(z, 2) & ")"
                    End If
            End Select

            If sql <> "" Then
                sql = "Insert into " & GLTFINRD & " " & sql
                ASCDATA1.ExecuteSQL(sql)
            End If
        Next

        If STMT_LINE_NO_ALL_ELSE <> 0 Then
            sql = "Select '" & F.HFs("STMT_CODE") & "' STMT_CODE, " & CStr(STMT_LINE_NO_ALL_ELSE) & " STMT_LINE_NO, " & A234 & " " _
                    & " from (" _
                    & "Select DISTINCT TT." & Replace(A234, ", ", ", TT.") & " from GLTACCT1," & TTC & " TT where TT.ACCT_CODE = GLTACCT1.ACCT_CODE and GLTACCT1.ACCT_TYPE in " & ACCT_TYPEs(F.HFs("STMT_TYPE")) _
                    & " MINUS " _
                    & "Select DISTINCT " & A234 & " from " & GLTFINRD _
                    & ")"
            sql = "Insert into " & GLTFINRD & " " & sql
            ASCDATA1.ExecuteSQL(sql)
        End If

        Call ASCMAIN1.AnalyzeTable(GLTFINRD)

        Return TTC

    End Function

    Public Shared Sub Setup_SQL( _
    ByRef F As ASFBASE1, _
    ByRef SQLP() As String, _
    ByRef SQLF() As String, _
    ByRef TTA As String, _
    ByRef TTB As String, _
    ByRef sqlA_select As String, _
    ByRef sqlA_group_by As String, _
    ByRef sqlA_where As String, _
    ByRef BY_SEG2 As Boolean, _
    ByRef BY_SEG3 As Boolean, _
    ByRef BY_SEG4 As Boolean)

        ReDim SQLP(4)
        ReDim SQLF(4)
        SQLP(0) = ", Sum (ACCT_BEG_BAL)"
        SQLP(1) = ", Sum (ACCT_BEG_BAL)"
        SQLP(4) = ", Sum (ACCT_BEG_BAL)"
        For i As Integer = 1 To 13
            SQLP(0) = SQLP(0) & ", Sum (ACCT_ACT_P" & Format$(i, "00") & ")"
            SQLP(1) = SQLP(1) & ", Sum (ACCT_BUD_P" & Format$(i, "00") & ")"
            SQLP(4) = SQLP(4) & ", Sum (ACCT_BUD_P" & Format$(i, "00") & ")"
        Next i
        SQLF(0) = " from " & TTA & " X"
        SQLF(1) = " from " & TTB & " X"
        SQLF(4) = " from " & TTB & " X"
        ' NEED AN ORIGINAL BUDGET TEMP TABLE

        sqlA_select = ""
        Dim z As String
        For i As Integer = 2 To 4
            Dim SEGX_CODE As String = "SEG" & CStr(i) & "_CODE"
            If Not New Boolean() {BY_SEG2, BY_SEG3, BY_SEG4}(i - 2) Then
                z = "'" & F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "'"
            Else
                z = "X." & SEGX_CODE
                sqlA_group_by = sqlA_group_by & ", " & z
            End If
            sqlA_select = sqlA_select & ", " & z & " " & SEGX_CODE
        Next

        Dim sql As String = ""
        z = SQLA("ACCT_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.ACCT_CODE " & IIf(SQLA("ACCT_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("SEG2_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.SEG2_CODE " & IIf(SQLA("SEG2_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("SEG3_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.SEG3_CODE " & IIf(SQLA("SEG3_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("SEG4_CODE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.SEG4_CODE " & IIf(SQLA("SEG4_CODE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        z = SQLA("ACCT_TYPE", "CODE_VALUES", True)
        If z <> "" Then
            sql &= " AND X.ACCT_TYPE " & IIf(SQLA("ACCT_TYPE", "EXCLUDE") = "1", "NOT ", "") & "IN (" & z & ")" & vbCr
        End If
        sqlA_where = sql
    End Sub

    Public Shared Function GL_Prep( _
    ByRef F As ASFBASE1, _
    ByRef YYYY_beg As String, _
    ByRef YYYY_end As String, _
    Optional ByRef budget As Boolean = False, _
    Optional ByRef OFFSET As Integer = 0, _
    Optional ByRef OFFSET_Y As Integer = 0, _
    Optional ByRef budTblsfx24 As String = "", _
    Optional ByRef TABLE_NAME As String = "") As String

        ' determine YYYY_gyp as lesser of YYYY_beg and GYP
        ' get all years into work table from YYYY_gyp thru endyear
        ' change nulls to zeroes
        ' get balance forward set for years > YYYY_gyp thru YYYY_end
        ' close net profit into RTE for all years from GYP thru YYYY_end

        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim z As String
        Dim sqlbs As String
        Dim sqlis As String
        Dim sql As String

        Dim GYP As String
        Dim RTE As String
        Dim RTEsql As String
        Dim RTEsql_group_by As String

        GYP = F.ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP")
        RTE = F.ROWs("GLTPARM1").Item("GL_PARM_RET_EARN_ACCT")

        RTEsql = ""
        RTEsql_group_by = ""
        Dim seg(4) As String
        seg(2) = F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        seg(3) = F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        seg(4) = F.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

        Dim YYYY_gyp As String
        If YYYY_beg < Mid$(GYP, 1, 4) Then
            YYYY_gyp = YYYY_beg
        Else
            YYYY_gyp = Mid$(GYP, 1, 4)
        End If

        Dim YRS As String
        YRS = ""
        For i = Val(YYYY_gyp) To Val(YYYY_end)
            YRS = YRS & ",'" & Format$(i, "0000") & "'"
        Next i
        YRS = Mid$(YRS, 2)

        '' Make sure that all Segment Codes are accounted for in Segment Master File

        'For i = 2 To 4
        '    z = Format$(i, "0")
        '    sql = "INSERT INTO GLTSEGM1 (ACCT_SEG_ID, ACCT_SEG_CODE, ACCT_SEG_DESC)"
        '    sql = sql & " SELECT '" & z & "', SEG" & z & "_CODE, 'Code ' || SEG" & z & "_CODE "
        '    sql = sql & " FROM"
        '    sql = sql & " (SELECT DISTINCT SEG" & z & "_CODE FROM GLTACCT3 "
        '    sql = sql & " MINUS "
        '    sql = sql & "  SELECT ACCT_SEG_CODE FROM GLTSEGM1 WHERE ACCT_SEG_ID = '" & z & "')"
        '    OraD.ExecuteSQL(sql)
        'Next i

        sql = "Select GLTACCT3.*, GLTACCT1.ACCT_TYPE "
        If budget Then
            If budTblsfx24 = "" Then
                budTblsfx24 = "2"
            End If
            sql = sql & " from EMP.GLTACCT" & budTblsfx24 & " GLTACCT3,EMP.GLTACCT1"
        Else
            sql = sql & " from EMP.GLTACCT3,EMP.GLTACCT1"
        End If
        sql = sql & " where GLTACCT1.ACCT_CODE (+) = GLTACCT3.ACCT_CODE"
        sql = sql & "   and GLTACCT3.ACCT_YEAR in (" & YRS & ")"
        Dim TT As String = ""
        If TABLE_NAME <> "" Then
            TT = TABLE_NAME
            ASCDATA1.ExecuteSQL("Delete from " & TT)
            ASCDATA1.ExecuteSQL("Insert into " & TT & " " & sql)
        Else
            TT = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & TT & " add Primary Key (ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,ACCT_YEAR)")
        End If

        For i = 2 To 4
            z = "SEG" & Format$(i, "0")
            If F.ROWs("GLTPARM1").Item("GL_PARM_" & z & "_RTE") & "" = "1" Then
                RTEsql = RTEsql & z & "_CODE,"
                RTEsql_group_by = RTEsql_group_by & z & "_CODE,"
            Else
                RTEsql = RTEsql & "'" & seg(i) & "' " & z & "_CODE,"
                RTEsql_group_by = RTEsql_group_by & "'" & seg(i) & "',"
            End If
        Next i

        ASCDATA1.ExecuteSQL("Update " & TT & " set ACCT_BEG_BAL = 0 where ACCT_BEG_BAL is Null")

        sqlbs = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE, NVL(ACCT_BEG_BAL,0) "
        sqlis = "Select " & RTEsql & " Sum (NVL(ACCT_BEG_BAL,0) "
        For i = 1 To 13
            If budget Then
                z = "ACCT_BUD_P" & Format$(i, "00")
            Else
                z = "ACCT_ACT_P" & Format$(i, "00")
            End If
            ASCDATA1.ExecuteSQL("Update " & TT & " Set " & z & " = 0 where " & z & " is Null")
            sqlbs = sqlbs & " + NVL (" & z & ",0)"
            sqlis = sqlis & " + NVL (" & z & ",0)"
        Next i
        sqlis = sqlis & ")"
        sqlbs = sqlbs & " ACCT_BEG_BAL"

        If TABLE_NAME = "" Then
            Call F.Create_TDA(F.dst.Tables.Add, TT, "*")
            For j = 0 To 13
                F.dst.Tables(TT).Columns(5 + j).DefaultValue = 0
            Next
        Else
            '            Fill_Records(TT)
        End If

        Dim RTE_imax As Integer

        If Val(Mid$(GYP, 1, 4)) <= YYYY_end - 1 Then
            Dim yz As String
            For i = Val(Mid$(GYP, 1, 4)) To YYYY_end - 1
                yz = Format$(i + 1, "0000")
                sql = sqlbs & " from " & TT & " where ACCT_TYPE in ('A','L','E') and ACCT_YEAR = '" & Format$(i, "0000") & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable(sql, "GLTACCTX").Rows
                    Dim ACCT_BEG_BAL As Double = Val(row.Item("ACCT_BEG_BAL") & "")
                    If ACCT_BEG_BAL <> 0 Then
                        Dim rowTT As DataRow = F.Fill_Record(TT, New String() {row.Item("ACCT_CODE"), _
                        row.Item("SEG2_CODE"), row.Item("SEG3_CODE"), row.Item("SEG4_CODE"), yz}, True)
                        rowTT.Item("ACCT_TYPE") = row.Item("ACCT_TYPE")
                        rowTT.Item("ACCT_BEG_BAL") = Val(rowTT.Item("ACCT_BEG_BAL") & "") + ACCT_BEG_BAL
                        F.Update_Record_TDA(TT) ' F.Update_Record_TDA_Rows(TT)
                    End If
                Next
                RTE_imax = i
                Call RTE_Calc(F, i, YYYY_gyp, TT, RTEsql_group_by, sqlis, RTE_imax, RTE)
            Next i
        End If

        If OFFSET <> 0 Then
            Stop ' WHEN WE HAVE A FRESH MIND
            '    Dim jmax As Integer
            '    j = 0
            '    If budget Then
            '        z = "BUD"
            '    Else
            '        z = "ACT"
            '    End If
            '    sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE" & vbCr
            '    sql = sql & ", SUM (DECODE(ACCT_YEAR,'" & Format$(YYYY_gyp, "0000") & "',NVL(ACCT_BEG_BAL,0))) P000" & vbCr
            '    For i = Val(YYYY_gyp) To Val(YYYY_end)
            '        For k = 1 To 12
            '            j = j + 1
            '            sql = sql & ", SUM (DECODE(ACCT_YEAR,'" & Format$(i, "0000") & "', NVL(ACCT_" & z & "_P" & Format$(k, "00") & ",0))) P" & Format$(j, "000") & vbCr
            '        Next k
            '    Next i
            '    sql = sql & " from " & TT & " group by "
            '    sql = sql & "ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE" & vbCr
            '    jmax = j
            '    Dim dyn As OraDynaset
            '    dyn = OraD.CreateDynaset(sql, 8&)
            '    ASCDATA1.ExecuteSQL("Delete from " & TT)
            '    sql = "Select * from " & TT & " where ROWNUM < 1"
            '    Dim dyntt As OraDynaset
            '    dyntt = OraD.CreateDynaset(sql, 0&)
            '    Dim a As Double
            '    Dim AMT() As Double
            '    Do While Not dyn.EOF
            '        ReDim AMT(12)
            '        k = 0
            '        i = OFFSET_Y
            '        For j = 0 To jmax
            '            a = Val(dyn.Fields("P" & Format$(j, "000")).Value & "")
            '            If j <= OFFSET Then
            '                AMT(0) = AMT(0) + a
            '            Else
            '                k = k + 1
            '                AMT(k) = a
            '                If k = 12 Or j = jmax Then
            '                    If InStr("ALE", dyn.Fields("ACCT_TYPE").Value & "") = 0 Then
            '                        If i = OFFSET_Y Then
            '                            dyntt.AddNew()
            '                            dyntt.Fields("ACCT_CODE").Value = dyn.Fields("ACCT_CODE").Value
            '                            dyntt.Fields("SEG2_CODE").Value = dyn.Fields("SEG2_CODE").Value
            '                            dyntt.Fields("SEG3_CODE").Value = dyn.Fields("SEG3_CODE").Value
            '                            dyntt.Fields("SEG4_CODE").Value = dyn.Fields("SEG4_CODE").Value
            '                            dyntt.Fields("ACCT_YEAR").Value = "0000" ' Val(y0) ' + i - 1
            '                            dyntt.Fields("ACCT_TYPE").Value = dyn.Fields("ACCT_TYPE").Value & ""
            '                            dyntt.Fields("ACCT_BEG_BAL").Value = AMT(0)
            '                            dyntt.Update()
            '                        End If
            '                        AMT(0) = 0
            '                    End If
            '                    If Val(YYYY_gyp) + i <= Val(YYYY_end) Then ' And (amt(0) <> 0 Or amt(1) <> 0 Or amt(2) <> 0 Or amt(3) <> 0 Or amt(4) <> 0 Or amt(5) <> 0 Or amt(6) <> 0 Or amt(7) <> 0 Or amt(8) <> 0 Or amt(9) <> 0 Or amt(10) <> 0 Or amt(11) <> 0 Or amt(12) <> 0) Then
            '                        dyntt.AddNew()
            '                        dyntt.Fields("ACCT_CODE").Value = dyn.Fields("ACCT_CODE").Value
            '                        dyntt.Fields("SEG2_CODE").Value = dyn.Fields("SEG2_CODE").Value
            '                        dyntt.Fields("SEG3_CODE").Value = dyn.Fields("SEG3_CODE").Value
            '                        dyntt.Fields("SEG4_CODE").Value = dyn.Fields("SEG4_CODE").Value
            '                        dyntt.Fields("ACCT_YEAR").Value = Val(y0) + i
            '                        dyntt.Fields("ACCT_TYPE").Value = dyn.Fields("ACCT_TYPE").Value & ""
            '                        dyntt.Fields("ACCT_BEG_BAL").Value = AMT(0)
            '                        For k = 1 To 12
            '                            dyntt.Fields("ACCT_" & z & "_P" & Format$(k, "00")).Value = AMT(k)
            '                            If InStr("ALE", dyn.Fields("ACCT_TYPE").Value & "") <> 0 Then
            '                                AMT(0) = AMT(0) + AMT(k)
            '                            End If
            '                            AMT(k) = 0
            '                        Next k
            '                        dyntt.Update()
            '                    End If
            '                    i = i + 1
            '                    k = 0
            '                End If
            '            End If
            '        Next j
            '        dyn.MoveNext()
            '    Loop

            '    i = 0
            '    RTE_imax = Val(YYYY_end)
            'GoSub Calc_RTE

            '    For i = Val(YYYY_gyp) To Val(YYYY_end)
            '        RTE_imax = Val(YYYY_end)
            '    GoSub Calc_RTE
            '    Next i

            '    sql = "Update " & TT & " SET ACCT_BEG_BAL = 0 "
            '    sql = sql & " where ACCT_TYPE in ('I','X') "
            '    sql = sql & " and ACCT_YEAR = '0000'"
            '    ASCDATA1.ExecuteSQL(sql) ' Clear out Accum R/E from periods prior to start of re-calendarized year which was stuffed into Op Accts

        End If

        ASCDATA1.ExecuteSQL("Delete from " & TT & " where ACCT_YEAR < '" & Format$(Val(YYYY_beg) + OFFSET_Y * Sign(Abs(OFFSET)), "0000") & "'")
        ASCDATA1.ExecuteSQL("Delete from " & TT & " where ACCT_YEAR > '" & Format$(Val(YYYY_end), "0000") & "'")

        If budget Then
            z = "BUD"
        Else
            z = "ACT"
        End If
        sql = "Delete from " & TT
        sql = sql & " where NVL(ACCT_BEG_BAL,0) = 0" & vbCr
        For k = 1 To 12
            sql = sql & " and NVL(ACCT_" & z & "_P" & Format$(k, "00") & ",0) = 0" & vbCr
        Next k
        'OraD.ExecuteSQL sql ' this throws off the TBAL where an account may have had activity which nets to 0

        If TABLE_NAME = "" Then
            ASCDATA1.ExecuteSQL("Create Index I_" & TT & "_1 on " & TT & " (ACCT_YEAR,ACCT_TYPE)")
        End If

        Call ASCMAIN1.AnalyzeTable(TT)

        Return TT

    End Function

    Public Shared Function SQLA( _
    ByRef PB_COLUMN_NAME As String, _
    Optional ByRef COLUMN_NAME As String = "CODE_VALUES", _
    Optional ByRef SQL_List As Boolean = False) As String
        Dim rowASTDSQLA As DataRow ' = tblASTDSQLA.Rows.Find(PB_COLUMN_NAME)
        rowASTDSQLA = Nothing ' THIS IS THE ONLY THING THAT WE NEED TO MOVE THESE ROUTINES FROM GLRSTMT1/ASFBASE1 TO THIS MODULE
        If rowASTDSQLA Is Nothing Then
            SQLA = ""
        Else
            SQLA = rowASTDSQLA.Item(COLUMN_NAME) & ""
            If SQL_List And SQLA <> "" Then
                SQLA = "'" & Replace(SQLA, ",", "','") & "'"
            End If
        End If
        Return SQLA
    End Function


    Public Shared Sub RTE_Calc(
    ByRef F As ASFBASE1,
    ByRef YYYY As Integer,
    ByRef YYYY_gyp As String,
    ByRef TT As String,
    ByRef RTEsql_group_by As String,
    ByRef sqlis As String,
    ByRef RTE_imax As Integer,
    ByRef RTE As String)

        Dim RTE_imin As Integer
        If YYYY = 0 Then
            RTE_imin = YYYY_gyp
        Else
            RTE_imin = YYYY
        End If
        Dim sql As String
        sql = sqlis & " from " & TT & " where ACCT_TYPE in ('I','X') "
        sql = sql & " and ACCT_YEAR = '" & Format$(YYYY, "0000") & "'"
        sql = sql & " group by " & Mid$(RTEsql_group_by, 1, Len(RTEsql_group_by) - 1)
        For Each rowRTE As DataRow In ASCDATA1.GetDataTable(sql, "").Rows
            Dim ACCT_BEG_BAL As Double = Val(rowRTE.Item(3) & "")
            If ACCT_BEG_BAL <> 0 Then
                For RTE_i As Integer = RTE_imin To RTE_imax
                    Dim row As DataRow = F.Fill_Record(TT, New String() {RTE,
                    rowRTE.Item("SEG2_CODE"), rowRTE.Item("SEG3_CODE"), rowRTE.Item("SEG4_CODE"),
                    Format$(RTE_i + 1, "0000")}, True)
                    row.Item("ACCT_TYPE") = "E"
                    row.Item("ACCT_BEG_BAL") = Val(row.Item("ACCT_BEG_BAL") & "") + ACCT_BEG_BAL
                    F.Update_Record_TDA(TT) ' F.Update_Record_TDA_Rows(TT)
                Next RTE_i
            End If
        Next
    End Sub

    Public Shared Function Load_Details(ACCT_YEAR As String, JOURNAL_TYPE As String, ACCT_CODE As String, sql_SEG As String, sqlWhereClause As String, Optional sqlACCT_CODEs As String = "") As String

        ASCMAIN1.sql = ""
        Dim sqlx As String = ""

        Select Case JOURNAL_TYPE

            Case "APCD"

                sqlx = "" _
                    & "Select APTCHCK1.OPS_YYYYPP, APTCHCK1.VEND_CODE CODE1_VALUE, APTCHCK1.BANK_CODE CODE2_VALUE, NULL CODE3_VALUE, Sum (APTCHCK1.CHECK_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTBANK1.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTBANK1.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTBANK1.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTBANK1.SEG4_CODE" & vbCrLf) _
                    & " from APTCHCK1,GLTBANK1" & vbCrLf _
                    & " where GLTBANK1.BANK_CODE = APTCHCK1.BANK_CODE" & vbCrLf _
                    & "   and APTCHCK1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and GLTBANK1.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(GLTBANK1.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by APTCHCK1.OPS_YYYYPP, APTCHCK1.VEND_CODE, APTCHCK1.BANK_CODE" & vbCrLf _
                    & IIf(True, "", ", GLTBANK1.ACCT_CODE" & vbCrLf)


                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, GLTBANK1.BANK_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,GLTBANK1" & vbCrLf _
                    & "  where GLTBANK1.BANK_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, GLTBANK1.BANK_DESC"

            Case "APIN"

                sqlx = "" _
                    & "Select NVL(APTINVH1.OPS_YYYYPP_ACCRUE,APTINVH1.OPS_YYYYPP) OPS_YYYYPP, APTINVH1.VEND_CODE CODE1_VALUE, APTINVH1.POST_CODE CODE2_VALUE, NULL CODE3_VALUE, Sum (APTINVH2.INV_LINE_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", APTINVH2.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", APTINVH2.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", APTINVH2.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", APTINVH2.SEG4_CODE" & vbCrLf) _
                    & " from APTINVH1,APTINVH2" & vbCrLf _
                    & " where APTINVH2.VOUCHER_NO = APTINVH1.VOUCHER_NO" & vbCrLf _
                    & "   and APTINVH1.REGISTER_XNO is Not Null" & vbCrLf _
                    & "   and APTINVH1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and APTINVH2.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(APTINVH2.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by NVL(APTINVH1.OPS_YYYYPP_ACCRUE,APTINVH1.OPS_YYYYPP), APTINVH1.VEND_CODE, APTINVH1.POST_CODE" & vbCrLf _
                    & IIf(True, "", ", APTINVH2.ACCT_CODE" & vbCrLf)


                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, APTVEND1.VEND_NAME DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,APTVEND1" & vbCrLf _
                    & "  where APTVEND1.VEND_CODE (+) = X.CODE1_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, APTVEND1.VEND_NAME"

            Case "ARCR"

                sqlx = "" _
                    & "Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE CODE1_VALUE, 'GL' CODE2_VALUE, NULL CODE3_VALUE, Sum (ARTPYMT4.GL_DIST_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTPYMT4.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTPYMT4.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTPYMT4.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTPYMT4.SEG4_CODE" & vbCrLf) _
                    & " from ARTPYMT1,ARTPYMT2,ARTPYMT4" & vbCrLf _
                    & " where ARTPYMT4.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                    & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT4.PYMT_BATCH_NO" & vbCrLf _
                    & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT4.PYMT_BATCH_LNO" & vbCrLf _
                    & "   and ARTPYMT2.PYMT_STATUS = '2'" & vbCrLf _
                    & "   and ARTPYMT1.STATUS = '2'" & vbCrLf _
                    & "   and ARTPYMT1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and ARTPYMT4.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(ARTPYMT4.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE" & vbCrLf _
                    & IIf(True, "", ", ARTPYMT4.ACCT_CODE" & vbCrLf) _
                    & " union " & vbCrLf _
                    & "Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE CODE1_VALUE, ARTPYMT5.REASON_CODE CODE2_VALUE, NULL CODE3_VALUE, Sum (ARTPYMT5.GL_DIST_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTREAS1.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTPYMT5.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTPYMT5.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTPYMT5.SEG4_CODE" & vbCrLf) _
                    & " from ARTPYMT1,ARTPYMT2,ARTPYMT5,ARTREAS1" & vbCrLf _
                    & " where ARTPYMT5.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
                    & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT5.PYMT_BATCH_NO" & vbCrLf _
                    & "   and ARTPYMT2.PYMT_BATCH_LNO = ARTPYMT5.PYMT_BATCH_LNO" & vbCrLf _
                    & "   and ARTPYMT2.PYMT_STATUS = '2'" & vbCrLf _
                    & "   and ARTPYMT1.STATUS = '2'" & vbCrLf _
                    & "   and ARTREAS1.REASON_CODE (+) = ARTPYMT5.REASON_CODE" & vbCrLf _
                    & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0'" & vbCrLf _
                    & "   and ARTPYMT1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and ARTREAS1.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(ARTPYMT5.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE" & vbCrLf _
                    & IIf(True, "", ", ARTREAS1.ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ARTREAS1.REASON_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,ARTREAS1" & vbCrLf _
                    & "  where ARTREAS1.REASON_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ARTREAS1.REASON_DESC"

            Case "GLJE"

                sqlx = "" _
                    & "Select GLTDETL1.OPS_YYYYPP, GLTDETL1.JOURNAL_NO CODE1_VALUE, GLTDETL1.DETL_DESC CODE2_VALUE, GLTDETL1.DETL_CTL_NO CODE3_VALUE, GLTJRNL1.JOURNAL_DESC DESC_VALUE, Sum (GLTDETL1.DETL_POSTING_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTDETL1.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTDETL1.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTDETL1.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTDETL1.SEG4_CODE" & vbCrLf) _
                    & " from GLTDETL1,GLTJRNL1" & vbCrLf _
                    & " where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
                    & "   and GLTDETL1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and GLTJRNL1.JOURNAL_TYPE = 'GLJE'" _
                    & "   and GLTDETL1.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(GLTDETL1.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by GLTDETL1.OPS_YYYYPP, GLTDETL1.JOURNAL_NO, GLTDETL1.DETL_DESC, GLTDETL1.DETL_CTL_NO, GLTJRNL1.JOURNAL_DESC" & vbCrLf _
                    & IIf(True, "", ", GLTDETL1.ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, GLTJRNL1.JOURNAL_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,GLTJRNL1" & vbCrLf _
                    & " where GLTJRNL1.JOURNAL_NO (+) = X.CODE1_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, GLTJRNL1.JOURNAL_DESC"

                'ARFX
                ', ICIR, ICRV
                'OPRA, 


            Case "ICIA"

                sqlx = "" _
                    & "Select ICTIADJ1.OPS_YYYYPP, ICTIADJ1.REASON_CODE CODE1_VALUE, ICTIADJ2.PROD_CODE CODE2_VALUE, ICTIADJ2.ITEM_CODE CODE3_VALUE, Sum (ICTIADJ3.DIST_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIADJ3.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIADJ3.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIADJ3.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIADJ3.SEG4_CODE" & vbCrLf) _
                    & " from ICTIADJ1,ICTIADJ2,ICTIADJ3,ICTITEM1,ICTCOLL1" & vbCrLf _
                    & " where ICTITEM1.ITEM_CODE = ICTIADJ2.ITEM_CODE" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & "   and ICTIADJ3.ADJ_NO = ICTIADJ2.ADJ_NO" & vbCrLf _
                    & "   and ICTIADJ3.ADJ_LNO = ICTIADJ2.ADJ_LNO" & vbCrLf _
                    & "   and ICTIADJ2.ADJ_NO = ICTIADJ1.ADJ_NO" & vbCrLf _
                    & "   and ICTIADJ1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and ICTIADJ3.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & sqlWhereClause _
                    & " group by ICTIADJ1.OPS_YYYYPP, ICTIADJ1.REASON_CODE, ICTIADJ2.PROD_CODE, ICTIADJ2.ITEM_CODE" & vbCrLf _
                    & IIf(True, "", ", ICTIADJ3.ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ICTPROD1.PROD_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,ICTPROD1" & vbCrLf _
                    & "  where ICTPROD1.PROD_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ICTPROD1.PROD_DESC"


            Case "ICIR"

                sqlx = "" _
                    & "Select ICTIREC1.OPS_YYYYPP, ICTIREC1.VEND_CODE CODE1_VALUE, ICTIREC2.PROD_CODE CODE2_VALUE, ICTIREC2.ITEM_CODE CODE3_VALUE, Sum (ICTIREC3.DIST_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIREC3.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIREC3.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIREC3.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ICTIREC3.SEG4_CODE" & vbCrLf) _
                    & " from ICTIREC1,ICTIREC2,ICTIREC3,ICTITEM1,ICTCOLL1" & vbCrLf _
                    & " where ICTITEM1.ITEM_CODE = ICTIREC2.ITEM_CODE" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & "   and ICTIREC3.RECEIPT_NO = ICTIREC2.RECEIPT_NO" & vbCrLf _
                    & "   and ICTIREC3.RECEIPT_LNO = ICTIREC2.RECEIPT_LNO" & vbCrLf _
                    & "   and ICTIREC2.RECEIPT_NO = ICTIREC1.RECEIPT_NO" & vbCrLf _
                    & "   and ICTIREC1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and ICTIREC3.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & sqlWhereClause _
                    & " group by ICTIREC1.OPS_YYYYPP, ICTIREC1.VEND_CODE, ICTIREC2.PROD_CODE, ICTIREC2.ITEM_CODE" & vbCrLf _
                    & IIf(True, "", ", ICTIREC3.ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ICTPROD1.PROD_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,ICTPROD1" & vbCrLf _
                    & "  where ICTPROD1.PROD_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ICTPROD1.PROD_DESC"

            Case "OPCJ"

                sqlx = "" _
                    & "Select SOTINVHT.OPS_YYYYPP, SOTINVH1.CUST_CODE CODE1_VALUE, SOTINVHT.PROD_CODE CODE2_VALUE, NULL CODE3_VALUE, Sum (SOTINVHT.CGS) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.ACCT_CODE_CGS ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.SEG4_CODE" & vbCrLf) _
                    & " from SOTINVH1,SOTINVHT" & vbCrLf _
                    & " where SOTINVHT.CGS <> 0" & vbCrLf _
                    & "   and SOTINVH1.INV_TYPE = SOTINVHT.INV_TYPE" & vbCrLf _
                    & "   and SOTINVH1.INV_NO = SOTINVHT.INV_NO" & vbCrLf _
                    & "   and SOTINVHT.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SOTINVHT.ACCT_CODE_CGS " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(SOTINVHT.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by SOTINVHT.OPS_YYYYPP, SOTINVH1.CUST_CODE, SOTINVHT.PROD_CODE" & vbCrLf _
                    & IIf(True, "", ", SOTINVHT.ACCT_CODE_CGS" & vbCrLf)


                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ICTPROD1.PROD_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,ICTPROD1" & vbCrLf _
                    & "  where ICTPROD1.PROD_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ICTPROD1.PROD_DESC"

            Case "OPSJ"

                'Dim sqlSEG2_CODE As String = "NVL(SOTTYPE1.SEG2_CODE,'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')"

                If Not ASCMAIN1.ActiveForm.ROWs.ContainsKey("SOTPARM1") Then
                    ASCMAIN1.ActiveForm.Get_PARM("SOTPARM1")
                End If

                Dim sqlSEG3_CODE As String = ""
                If ASCMAIN1.ActiveForm.ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG3") & "" = "1" Then
                    sqlSEG3_CODE = "NVL(SOTTCLS1.SEG3_CODE,ARTCUST1.TRADE_CLASS_CODE)"
                Else
                    sqlSEG3_CODE = "'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'"
                End If

                Dim sqlSEG4_CODE As String = ""
                If ASCMAIN1.ActiveForm.ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG4") & "" = "1" Then
                    sqlSEG4_CODE = "NVL(ICTCOLL1.SEG4_CODE,NVL(ICTITEM1.COLLECTION_CODE,'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'))"
                Else
                    sqlSEG4_CODE = "'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'"
                End If

                Dim sql_SEG2 As String = "NVL(NVL(NVL(SOTCHAN1.SEG2_CODE,SOTTYPE1.SEG2_CODE),NVL(SOTMISC1.SEG2_CODE,SOTSDIV1.SEG2_CODE)),'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')"
                Dim sql_SEG3 As String = "CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVH1.EVENT_CODE IS NULL THEN " & sqlSEG3_CODE & " ELSE '" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END"
                Dim sql_SEG4 As String = "CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVH1.EVENT_CODE IS NULL THEN " & Replace(sqlSEG4_CODE, "ICTITEM1", "SOTINVH1") & " ELSE '" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END"

                sqlx = "" _
                    & "Select SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.CUST_CODE CODE1_VALUE, NULL CODE2_VALUE, SOTINVH1.MISC_CHG_CODE CODE3_VALUE, Sum (-1 * SOTINVH1.INV_MISC_CHG) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTMISC1.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG2 & " SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG3 & " SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG4 & " SEG4_CODE" & vbCrLf) _
                    & " from SOTINVH1,SOTMISC1,ARTCUST1,SOTTCLS1,SOTCHAN1,ICTCOLL1,SOTEVNT1,SOTSDIV1,SOTTYPE1" & vbCrLf _
                    & " where SOTINVH1.MISC_CHG_CODE IS NOT NULL" & vbCrLf _
                    & "   and SOTMISC1.MISC_CHG_CODE = SOTINVH1.MISC_CHG_CODE" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTCHAN1.CHANNEL_CODE (+) = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE (+) = SOTINVH1.COLLECTION_CODE" & vbCrLf _
                    & "   and SOTEVNT1.EVENT_CODE (+) = SOTINVH1.EVENT_CODE" & vbCrLf _
                    & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
                    & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                    & "   and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SOTMISC1.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(SOTMISC1.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.CUST_CODE, SOTINVH1.MISC_CHG_CODE" & vbCrLf _
                    & IIf(True, "", ", SOTMISC1.ACCT_CODE" & vbCrLf) _
                    & " union " & vbCrLf _
                    & "Select SOTINVHT.OPS_YYYYPP, SOTINVH1.CUST_CODE CODE1_VALUE, SOTINVHT.PROD_CODE CODE2_VALUE, NULL CODE3_VALUE, Sum (-1 * SOTINVHT.SLS) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.ACCT_CODE_SLS" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTINVHT.SEG4_CODE" & vbCrLf) _
                    & " from SOTINVH1,SOTINVHT" & vbCrLf _
                    & " where SOTINVHT.SLS <> 0" & vbCrLf _
                    & "   and SOTINVH1.INV_TYPE = SOTINVHT.INV_TYPE" & vbCrLf _
                    & "   and SOTINVH1.INV_NO = SOTINVHT.INV_NO" & vbCrLf _
                    & "   and SOTINVHT.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SOTINVHT.ACCT_CODE_SLS " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & Replace(sql_SEG, "NVL(X.", "NVL(SOTINVHT.") & vbCrLf _
                    & sqlWhereClause _
                    & " group by SOTINVHT.OPS_YYYYPP, SOTINVH1.CUST_CODE, SOTINVHT.PROD_CODE" & vbCrLf _
                    & IIf(True, "", ", SOTINVHT.ACCT_CODE_SLS" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, NVL(ICTPROD1.PROD_DESC,SOTMISC1.MISC_CHG_DESC) DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,SOTMISC1,ICTPROD1" & vbCrLf _
                    & "  where SOTMISC1.MISC_CHG_CODE (+) = X.CODE3_VALUE" & vbCrLf _
                    & "    and ICTPROD1.PROD_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, NVL(ICTPROD1.PROD_DESC,SOTMISC1.MISC_CHG_DESC)"




            Case "OPRA"

                If Not ASCMAIN1.ActiveForm.ROWs.ContainsKey("SOTPARM1") Then
                    ASCMAIN1.ActiveForm.Get_PARM("SOTPARM1")
                End If

                Dim sqlSEG3_CODE As String = ""
                If ASCMAIN1.ActiveForm.ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG3") & "" = "1" Then
                    sqlSEG3_CODE = "NVL(SOTTCLS1.SEG3_CODE,ARTCUST1.TRADE_CLASS_CODE)"
                Else
                    sqlSEG3_CODE = "'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'"
                End If

                Dim sqlSEG4_CODE As String = ""
                If ASCMAIN1.ActiveForm.ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG4") & "" = "1" Then
                    sqlSEG4_CODE = "NVL(ICTCOLL1.SEG4_CODE,NVL(ICTITEM1.COLLECTION_CODE,'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'))"
                Else
                    sqlSEG4_CODE = "'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'"
                End If

                Dim sql_SEG2 As String = "NVL(SOTCHAN1.SEG2_CODE,'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')"
                Dim sql_SEG3 As String = sqlSEG3_CODE ' "CASE WHEN 1=1 THEN " & sqlSEG3_CODE & " ELSE '" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END"
                Dim sql_SEG4 As String = sqlSEG4_CODE ' "CASE WHEN 1=1 THEN " & sqlSEG4_CODE & " ELSE '" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END"


                Dim sqlSOTRMAF3 As String = "SOTRMAF3"
                sqlSOTRMAF3 = "" _
                    & "(" _
                    & " Select SOTRMAF3.OPS_YYYYPP, SOTRMAF3.RA_NO, SOTRMAF3.ITEM_CODE, SOTRMAF3.RA_NET_PRICE, SOTRMAF3.RA_QTY_OPEN from SOTRMAF3" & vbCrLf _
                    & "  where SOTRMAF3.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & " union " & vbCrLf _
                    & " Select Period_Calc(SOTRMAF3.OPS_YYYYPP,1) OPS_YYYYPP, SOTRMAF3.RA_NO, SOTRMAF3.ITEM_CODE, SOTRMAF3.RA_NET_PRICE, -1 * SOTRMAF3.RA_QTY_OPEN RA_QTY_OPEN from SOTRMAF3" & vbCrLf _
                    & "  where SOTRMAF3.OPS_YYYYPP between '" & Format(Val(ACCT_YEAR) - 1, "0000") & "12' AND '" & ACCT_YEAR & "11'" & vbCrLf _
                    & ")"
                sqlx = "" _
                    & "Select SOTRMAF3.OPS_YYYYPP, SOTRMAF1.CUST_CODE CODE1_VALUE, SOTRMAF1.RA_REASON_CODE CODE2_VALUE, NULL CODE3_VALUE, SOTRMAF1.RA_NOTES DESC_VALUE, Sum (1 * SOTRMAF3.RA_NET_PRICE * SOTRMAF3.RA_QTY_OPEN) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTPARM1.SO_PARM_RA_REV_ACCT_CODE ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG2 & " SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG3 & " SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG4 & " SEG4_CODE" & vbCrLf) _
                    & " from " & sqlSOTRMAF3 & " SOTRMAF3,SOTRMAF1,ARTREASR,SOTPARM1,ARTCUST1,SOTTCLS1,SOTCHAN1,ICTITEM1,ICTCOLL1" & vbCrLf _
                    & " where ARTREASR.RA_REASON_CODE = SOTRMAF1.RA_REASON_CODE" & vbCrLf _
                    & "   and SOTPARM1.SO_PARM_KEY = 'Z'" & vbCrLf _
                    & "   and SOTRMAF3.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SOTRMAF1.RA_NO = SOTRMAF3.RA_NO" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE (+) = SOTRMAF1.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTCHAN1.CHANNEL_CODE (+) = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                    & "   and ICTITEM1.ITEM_CODE (+) = SOTRMAF3.ITEM_CODE" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & "   and SOTPARM1.SO_PARM_RA_REV_ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & sqlWhereClause _
                    & " group by SOTRMAF3.OPS_YYYYPP, SOTRMAF1.CUST_CODE, SOTRMAF1.RA_REASON_CODE, SOTRMAF1.RA_NOTES" & vbCrLf _
                    & IIf(True, "", ", SOTPARM1.SO_PARM_RA_REV_ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ARTREASR.RA_REASON_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,ARTREASR" & vbCrLf _
                    & "  where ARTREASR.RA_REASON_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, ARTREASR.RA_REASON_DESC"



            Case "SPAC"

                Dim sqlSEG3_CODE As String = "NVL(SOTTCLS1.SEG3_CODE,NVL(ARTCUST1.TRADE_CLASS_CODE,'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'))"
                Dim sqlSEG4_CODE As String = "NVL(ICTCOLL1.SEG4_CODE,NVL(SPTACOMC.COLLECTION_CODE,'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'))"
 
                Dim sql_SEG2 As String = "NVL(SOTCHAN1.SEG2_CODE,'" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')"
                Dim sql_SEG3 As String = sqlSEG3_CODE ' "CASE WHEN 1=1 THEN " & sqlSEG3_CODE & " ELSE '" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END"
                Dim sql_SEG4 As String = sqlSEG4_CODE ' "CASE WHEN 1=1 THEN " & sqlSEG4_CODE & " ELSE '" & ASCMAIN1.ActiveForm.ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END"

                Dim sqlDESC_VALUE As String = "TO_CHAR(SPTACOM1.ASP_COMM_PCT) || '% of ' || DECODE (NVL(SPTACOM1.ASP_COMM_BASIS,'0'),'0','EDI Retail Sales','1','Grs Ship @Net Pr','2','Fixed Demo','?')"
                sqlx = "" _
                    & "Select SPTACOMC.OPS_YYYYPP, SPTACOMC.CUST_CODE CODE1_VALUE, SPTACOMC.ASP_CODE CODE2_VALUE, NULL CODE3_VALUE, " & sqlDESC_VALUE & " DESC_VALUE, Sum (SPTACOMC.AMT_COMM) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", SPTACOM0.ASP_ACCT_CODE_EXP ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG2 & " SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG3 & " SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", " & sql_SEG4 & " SEG4_CODE" & vbCrLf) _
                    & " from SPTACOMB SPTACOMC,SPTACOM0,SPTACOM1,ARTCUST1,SOTTCLS1,SOTCHAN1,ICTCOLL1" & vbCrLf _
                    & " where SPTACOM0.ASP_CODE = SPTACOMC.ASP_CODE" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE (+) = SPTACOMC.COLLECTION_CODE" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SPTACOMC.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTCHAN1.CHANNEL_CODE (+) = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                    & "   and SPTACOM1.ASP_CODE = SPTACOMC.ASP_CODE" & vbCrLf _
                    & "   and SPTACOM1.CUST_CODE = SPTACOMC.CUST_CODE" & vbCrLf _
                    & "   and SPTACOMC.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SPTACOM0.ASP_ACCT_CODE_EXP " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & sqlWhereClause _
                    & " group by SPTACOMC.OPS_YYYYPP, SPTACOMC.CUST_CODE, SPTACOMC.ASP_CODE, " & sqlDESC_VALUE & vbCrLf _
                    & IIf(True, "", ", SPTACOM0.ASP_ACCT_CODE_EXP" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTACOM0.ASP_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,SPTACOM0" & vbCrLf _
                    & "  where SPTACOM0.ASP_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTACOM0.ASP_DESC"


            Case "SPCA"

                sqlx = "" _
                    & "Select GLTTSPCA.OPS_YYYYPP, GLTTSPCA.CUST_CODE CODE1_VALUE, GLTTSPCA.EXPENSE_TYPE_CODE CODE2_VALUE, GLTTSPCA.COLLECTION_CODE CODE3_VALUE, Sum (GLTTSPCA.DIST_AMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTTSPCA.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTTSPCA.SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTTSPCA.SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", GLTTSPCA.SEG4_CODE" & vbCrLf) _
                    & " from GLTTSPCA,SPTCOOP1" & vbCrLf _
                    & " where GLTTSPCA.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SPTCOOP1.AUTH_NO = GLTTSPCA.DIST_CTL_NO" & vbCrLf _
                    & "   and GLTTSPCA.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & sqlWhereClause _
                    & " group by GLTTSPCA.OPS_YYYYPP, GLTTSPCA.CUST_CODE, GLTTSPCA.EXPENSE_TYPE_CODE, GLTTSPCA.COLLECTION_CODE" & vbCrLf _
                    & IIf(True, "", ", GLTTSPCA.ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTTYPE1.EXPENSE_TYPE_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,SPTTYPE1" & vbCrLf _
                    & "  where SPTTYPE1.EXPENSE_TYPE_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTTYPE1.EXPENSE_TYPE_DESC"


            Case "SPCP"

                '  & Replace(sql_SEG, "NVL(X.", "NVL(ARTPYMT4.") & vbCrLf _

                sqlx = "" _
                    & "Select SPTPYMT1.OPS_YYYYPP, SPTPYMT1.CUST_CODE CODE1_VALUE, SPTCOOP1.EXPENSE_TYPE_CODE CODE2_VALUE, SPTPYMT3.COLLECTION_CODE CODE3_VALUE, SPTCOOP1.BOOKING_NAME DESC_VALUE, Sum (SPTPYMT3.DIST_AMT_PYMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", SPTTYPE1.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTCHAN1.SEG2_CODE SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTCUST1.TRADE_CLASS_CODE SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SPTPYMT3.COLLECTION_CODE SEG4_CODE" & vbCrLf) _
                    & " from SPTPYMT1,SPTPYMT3,SPTCOOP1,SPTTYPE1,ARTCUST1,SOTTCLS1,SOTCHAN1" & vbCrLf _
                    & " where SPTPYMT1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SPTCOOP1.AUTH_NO = SPTPYMT3.AUTH_NO" & vbCrLf _
                    & "   and SPTPYMT3.PYMT_NO = SPTPYMT1.PYMT_NO" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SPTCOOP1.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTCHAN1.CHANNEL_CODE = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                    & "   and SPTTYPE1.EXPENSE_TYPE_CODE = SPTCOOP1.EXPENSE_TYPE_CODE" & vbCrLf _
                    & "   and SPTTYPE1.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & sqlWhereClause _
                    & " group by SPTPYMT1.OPS_YYYYPP, SPTPYMT1.CUST_CODE, SPTCOOP1.EXPENSE_TYPE_CODE, SPTPYMT3.COLLECTION_CODE, SPTCOOP1.BOOKING_NAME" & vbCrLf _
                    & IIf(True, "", ", SPTTYPE1.ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTTYPE1.EXPENSE_TYPE_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,SPTTYPE1" & vbCrLf _
                    & "  where SPTTYPE1.EXPENSE_TYPE_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTTYPE1.EXPENSE_TYPE_DESC"

            Case "SPDP"

                sqlx = "" _
                    & "Select SPTDPMT1.OPS_YYYYPP, SPTDPMT1.CUST_CODE CODE1_VALUE, SPTDPMT1.PYMT_CTL_NO CODE2_VALUE, SPTDPMT2.COLLECTION_CODE CODE3_VALUE, Sum (SPTPYMT3.DIST_AMT_PYMT) AMT" & vbCrLf _
                    & IIf(sqlACCT_CODEs = "", "", ", SPTTYPE1.ACCT_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SOTCHAN1.SEG2_CODE SEG2_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", ARTCUST1.TRADE_CLASS_CODE SEG3_CODE" & vbCrLf) _
                    & IIf(sqlACCT_CODEs = "", "", ", SPTPYMT3.COLLECTION_CODE SEG4_CODE" & vbCrLf) _
                    & " from SPTPYMT1,SPTPYMT3,SPTCOOP1,SPTTYPE1,ARTCUST1,SOTTCLS1,SOTCHAN1" & vbCrLf _
                    & " where SPTPYMT1.OPS_YYYYPP between '" & ACCT_YEAR & "01' AND '" & ACCT_YEAR & "12'" & vbCrLf _
                    & "   and SPTCOOP1.AUTH_NO = SPTPYMT3.AUTH_NO" & vbCrLf _
                    & "   and SPTPYMT3.PYMT_NO = SPTPYMT1.PYMT_NO" & vbCrLf _
                    & "   and ARTCUST1.CUST_CODE = SPTCOOP1.CUST_CODE" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   and SOTCHAN1.CHANNEL_CODE = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                    & "   and SPTTYPE1.EXPENSE_TYPE_CODE = SPTCOOP1.EXPENSE_TYPE_CODE" & vbCrLf _
                    & "   and SPTTYPE1.ACCT_CODE " & IIf(sqlACCT_CODEs = "", " = '" & ACCT_CODE & "'", sqlACCT_CODEs) & vbCrLf _
                    & sqlWhereClause _
                    & " group by SPTPYMT1.OPS_YYYYPP, SPTPYMT1.CUST_CODE, SPTCOOP1.EXPENSE_TYPE_CODE, SPTPYMT3.COLLECTION_CODE" & vbCrLf _
                    & IIf(True, "", ", SPTTYPE1.ACCT_CODE" & vbCrLf)

                Dim T As String = ""
                For i As Integer = 1 To 12
                    T &= ", Sum (DECODE(X.OPS_YYYYPP,'" & ACCT_YEAR & Format(i, "00") & "',AMT,0)) AMT_" & Format(i, "00") & vbCrLf
                Next
                ASCMAIN1.sql = "Select X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTTYPE1.EXPENSE_TYPE_DESC DESC_VALUE" & T & ", Sum (X.AMT) AMT" & vbCrLf _
                    & " from (" & vbCrLf _
                    & Replace(sqlx, sqlWhereClause, "") & vbCrLf _
                    & ") X,SPTTYPE1" & vbCrLf _
                    & "  where SPTTYPE1.EXPENSE_TYPE_CODE (+) = X.CODE2_VALUE" & vbCrLf _
                    & " group by X.CODE1_VALUE, X.CODE2_VALUE, X.CODE3_VALUE, SPTTYPE1.EXPENSE_TYPE_DESC"




        End Select

        Return sqlx
    End Function

    Public Shared Sub Load_Sub_Details(JOURNAL_TYPE As String, ACCT_YEAR As String, OPS_YYYYPP As String,
                                sqlX As String, sqlWhereClause As String,
                                CODE1_VALUE As String, CODE2_VALUE As String, CODE3_VALUE As String)

        ASCMAIN1.sql = ""

        Select Case JOURNAL_TYPE

            Case "APCD"
                Dim sqlw As String = "  and APTCHCK1.VEND_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and APTCHCK1.BANK_CODE = '" & CODE2_VALUE & "'" & vbCrLf & "  and APTCHCK1.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw,, 1)

                Dim SQL1 As String = "APTCHCK1.CHECK_AMT AMT," _
                    & "APTCHCK1.PYMT_METHOD, APTCHCK1.CHECK_STATUS, APTCHCK1.CHECK_NUM, APTCHCK1.CHECK_DATE"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (APTCHCK1.CHECK_AMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by APTCHCK1.OPS_YYYYPP, APTCHCK1.VEND_CODE, APTCHCK1.BANK_CODE", "",, 1)

            Case "APIN"
                Dim sqlw As String = "  and APTINVH1.VEND_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and APTINVH1.POST_CODE = '" & CODE2_VALUE & "'" & vbCrLf & "  and NVL(APTINVH1.OPS_YYYYPP_ACCRUE,APTINVH1.OPS_YYYYPP) = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw,, 1)

                Dim SQL1 As String = "APTINVH2.INV_LINE_AMT AMT," _
                    & "APTINVH1.VOUCHER_NO, APTINVH1.INV_TYPE, APTINVH1.INV_NUM, APTINVH1.INV_DATE," _
                    & "APTINVH1.INV_REF, APTINVH1.CHECK_NUM, APTINVH1.CHECK_DATE," _
                    & "APTINVH1.INIT_DATE, APTINVH1.INIT_OPER, APTINVH1.LAST_DATE, APTINVH1.LAST_OPER"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (APTINVH2.INV_LINE_AMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by NVL(APTINVH1.OPS_YYYYPP_ACCRUE,APTINVH1.OPS_YYYYPP), APTINVH1.VEND_CODE, APTINVH1.POST_CODE", "", , 1)

            Case "ARCR"
                Dim sqlw As String = "  and ARTPYMT2.CUST_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and ARTPYMT1.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX

                ' THE NEXT 2 LINES ARE CAUSING THE GLRDTLA1 REPORT TO CREATE BALANCING ENTRIES WHERE IT DOES NOT NEED TO
                'ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & IIf(CODE2_VALUE = "GL", "", " and ROWNUM < 1") & vbCrLf, , 1)
                'ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & IIf(CODE2_VALUE = "GL", " and ROWNUM < 1", "") & vbCrLf, , 1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & IIf(CODE2_VALUE = "GL", "", "") & vbCrLf, , 1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & IIf(CODE2_VALUE = "GL", "", "") & vbCrLf, , 1)

                Dim SQL1 As String = "ARTPYMT4.GL_DIST_AMT AMT," _
                    & "ARTPYMT1.BANK_CODE, ARTPYMT1.INIT_DATE, ARTPYMT1.INIT_OPER, ARTPYMT1.PYMT_SOURCE," _
                    & "ARTPYMT2.CUST_PYMT_REF_NO, ARTPYMT2.CUST_PYMT_REF_DATE," _
                    & "ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO," _
                    & "ARTPYMT4.GL_DIST_REF REF1, ARTPYMT4.GL_DIST_COMMENT REF2"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (ARTPYMT4.GL_DIST_AMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE", "",, 1)


                Dim SQL2 As String = "ARTPYMT5.GL_DIST_AMT AMT," _
                    & "ARTPYMT1.BANK_CODE, ARTPYMT1.INIT_DATE, ARTPYMT1.INIT_OPER, ARTPYMT1.PYMT_SOURCE," _
                    & "ARTPYMT2.CUST_PYMT_REF_NO, ARTPYMT2.CUST_PYMT_REF_DATE," _
                    & "ARTPYMT2.PYMT_BATCH_NO, ARTPYMT2.PYMT_BATCH_LNO," _
                    & "ARTPYMT5.CUST_REFERENCE REF1, ARTPYMT5.OUR_REFERENCE REF2"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (ARTPYMT5.GL_DIST_AMT) AMT", SQL2)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, ARTPYMT5.REASON_CODE", "")


            Case "GLJE"
                Dim sqlw As String = " And GLTDETL1.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & vbCrLf,, 1)

                '& "GLTJRNL1.JOURNAL_DESC, GLTDETL1.DETL_DESC, GLTDETL1.DETL_CTL_NO, GLTDETL1.DETL_CVX_NO," _
                'Dim SQL1 As String = "GLTDETL1.DETL_POSTING_AMT AMT," _
                '    & "GLTDETL1.SEG2_CODE, GLTDETL1.SEG3_CODE, GLTDETL1.SEG4_CODE, GLTDETL1.DETL_CVX_NO," _
                '    & "GLTJRNL1.INIT_DATE, GLTJRNL1.INIT_OPER"

                Dim SQL1 As String = "GLTDETL1.DETL_POSTING_AMT AMT," _
                    & "GLTDETL1.DETL_CVX_NO," _
                    & "GLTJRNL1.INIT_DATE, GLTJRNL1.INIT_OPER"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (GLTDETL1.DETL_POSTING_AMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by GLTDETL1.OPS_YYYYPP, GLTDETL1.JOURNAL_NO, GLTDETL1.DETL_DESC, GLTDETL1.DETL_CTL_NO, GLTJRNL1.JOURNAL_DESC", "",, 1)

            Case "ICIA"
                Dim sqlw As String = "  And ICTIADJ1.REASON_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and ICTIADJ2.PROD_CODE = '" & CODE3_VALUE & "'" & vbCrLf & "  and ICTIADJ2.ITEM_CODE = '" & CODE3_VALUE & "'" & vbCrLf & "  and ICTIADJ1.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw,, 1)

                Dim SQL1 As String = "ICTIADJ3.DIST_AMT AMT," _
                    & "ICTIADJ2.COST_CATGY_CODE, ICTIADJ1.WHSE_CODE," _
                    & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (ICTIADJ3.DIST_AMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by ICTIADJ1.OPS_YYYYPP, ICTIADJ1.REASON_CODE, ICTIADJ2.PROD_CODE, ICTIADJ2.ITEM_CODE", "",, 1)

            Case "ICIR"
                Dim sqlw As String = "  and ICTIREC1.VEND_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and ICTIREC2.PROD_CODE = '" & CODE3_VALUE & "'" & vbCrLf & "  and ICTIREC2.ITEM_CODE = '" & CODE3_VALUE & "'" & vbCrLf & "  and ICTIREC1.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw,, 1)

                Dim SQL1 As String = "ICTIREC3.DIST_AMT AMT," _
                    & "ICTIREC2.COST_CATGY_CODE, ICTIREC1.WHSE_CODE," _
                    & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (ICTIREC3.DIST_AMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by ICTIREC1.OPS_YYYYPP, ICTIREC1.VEND_CODE, ICTIREC2.PROD_CODE, ICTIREC2.ITEM_CODE", "",, 1)

            Case "OPSJ", "OPCJ"
                Dim sqlw As String = "  and SOTINVH1.CUST_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and SOTINVH1.ORDR_YYYYPP_UPDATED = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & vbCrLf,, 1)

                sqlw = Replace(sqlw, "SOTINVH1.ORDR_YYYYPP_UPDATED", "SOTINVHT.OPS_YYYYPP")
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & vbCrLf,, 1)

                Dim SQL1 As String = "-1 * SOTINVH1.INV_MISC_CHG AMT," _
                    & "SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.ORDR_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.EVENT_CODE, SOTINVH1.ORDR_TYPE_CODE," _
                    & "SOTINVH1.INIT_DATE, SOTINVH1.INIT_OPER, SOTINVH1.CUST_STORE_NO, ROWNUM ROW_NO"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (-1 * SOTINVH1.INV_MISC_CHG) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.CUST_CODE, SOTINVH1.MISC_CHG_CODE", "",, 1)

                Dim SQL2 As String = "-1 * SOTINVHT.SLS AMT," _
                    & "SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.ORDR_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.EVENT_CODE, SOTINVH1.ORDR_TYPE_CODE," _
                    & "SOTINVH1.INIT_DATE, SOTINVH1.INIT_OPER, SOTINVH1.CUST_STORE_NO, ROWNUM ROW_NO"
                If JOURNAL_TYPE = "OPSJ" Then
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (-1 * SOTINVHT.SLS) AMT", SQL2)
                ElseIf JOURNAL_TYPE = "OPCJ" Then
                    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (SOTINVHT.CGS) AMT", SQL2)
                End If
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by SOTINVHT.OPS_YYYYPP, SOTINVH1.CUST_CODE, SOTINVHT.PROD_CODE", "")



            Case "OPRA"
                Dim sqlw As String = "  and SOTRMAF1.CUST_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and SOTRMAF3.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw & vbCrLf, , 1)

                Dim SQL1 As String = "1 * SOTRMAF3.RA_NET_PRICE * SOTRMAF3.RA_QTY_OPEN AMT," _
                    & "SOTRMAF1.RA_NO, SOTRMAF1.CUST_CLAIM_NO, SOTRMAF1.RA_DATE," _
                    & "SOTRMAF1.INIT_DATE, SOTRMAF1.INIT_OPER, SOTRMAF1.CUST_STORE_NO, ROWNUM ROW_NO"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (1 * SOTRMAF3.RA_NET_PRICE * SOTRMAF3.RA_QTY_OPEN) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by SOTRMAF3.OPS_YYYYPP, SOTRMAF1.CUST_CODE, SOTRMAF1.RA_REASON_CODE, SOTRMAF1.RA_NOTES", "", , 1)

                'Dim SQL2 As String = "SOTINVHT.SLS AMT," _
                '    & "SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.ORDR_NO, SOTINVH1.ORDR_CUST_PO, SOTINVH1.EVENT_CODE, SOTINVH1.ORDR_TYPE_CODE," _
                '    & "SOTINVH1.INIT_DATE, SOTINVH1.INIT_OPER, SOTINVH1.CUST_STORE_NO, ROWNUM ROW_NO"
                'If JOURNAL_TYPE = "OPSJ" Then
                '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (SOTINVHT.SLS) AMT", SQL2)
                'ElseIf JOURNAL_TYPE = "OPCJ" Then
                '    ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (SOTINVHT.CGS) AMT", SQL2)
                'End If
                'ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by SOTINVHT.OPS_YYYYPP, SOTINVH1.CUST_CODE, SOTINVHT.PROD_CODE", "")
                 

            Case "SPAC"
                Dim sqlw As String = "  and SPTACOMC.CUST_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and SPTACOMC.ASP_CODE = '" & CODE2_VALUE & "'" & vbCrLf & "  and SPTACOMC.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw,, 1)

                Dim SQL1 As String = "SPTACOMC.AMT_COMM AMT," _
                    & "SPTACOMC.HC_CODE, SPTACOMC.BRAND_CODE," _
                    & "SPTACOMC.OPS_YYYYWW_MIN, SPTACOMC.OPS_YYYYWW_MAX," _
                    & "SPTACOMC.INIT_DATE, SPTACOMC.INIT_OPER, SPTACOMC.LAST_DATE, SPTACOMC.LAST_OPER"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (SPTACOMC.AMT_COMM) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by SPTACOMC.OPS_YYYYPP, SPTACOMC.CUST_CODE, SPTACOMC.ASP_CODE, TO_CHAR(SPTACOM1.ASP_COMM_PCT) || '% of ' || DECODE (NVL(SPTACOM1.ASP_COMM_BASIS,'0'),'0','EDI Retail Sales','1','Grs Ship @Net Pr','2','Fixed Demo','?')", "", , 1)

            Case "SPCA"
                Dim sqlw As String = "  and GLTTSPCA.CUST_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and GLTTSPCA.EXPENSE_TYPE_CODE = '" & CODE2_VALUE & "'" & vbCrLf & "  and GLTTSPCA.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw,, 1)

                Dim SQL1 As String = "GLTTSPCA.DIST_AMT AMT," _
                    & "GLTTSPCA.DIST_CTL_NO AUTH_NO,GLTTSPCA.INIT_OPER,GLTTSPCA.INIT_DATE,GLTTSPCA.LAST_OPER,GLTTSPCA.LAST_DATE,GLTTSPCA.DIST_DESC DESC_VALUE," _
                    & "GLTTSPCA.SEASON_CODE,GLTTSPCA.VEHICLE_CODE," _
                    & "GLTTSPCA.DATE_START,GLTTSPCA.DATE_END,GLTTSPCA.DATE_ACCRUE,GLTTSPCA.OPS_YYYYPP_ACCRUE"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (GLTTSPCA.DIST_AMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by GLTTSPCA.OPS_YYYYPP, GLTTSPCA.CUST_CODE, GLTTSPCA.EXPENSE_TYPE_CODE, GLTTSPCA.COLLECTION_CODE", "",, 1)

            Case "SPCP"
                Dim sqlw As String = "  and SPTPYMT1.CUST_CODE = '" & CODE1_VALUE & "'" & vbCrLf & "  and SPTCOOP1.EXPENSE_TYPE_CODE = '" & CODE2_VALUE & "'" & vbCrLf & "  and SPTPYMT1.OPS_YYYYPP = '" & OPS_YYYYPP & "'" & vbCrLf
                If OPS_YYYYPP = "" Then sqlw = ""
                ASCMAIN1.sql = sqlX
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, sqlWhereClause, sqlw, , 1)
                'REMOVING COLLECTION_CODE AND EXPENSE_TYPE_CODE BECAUSE ACCT DTL ANALYSIS SAYS DUP COLS
                Dim SQL1 As String = "SPTPYMT3.DIST_AMT_PYMT AMT," _
                    & "SPTPYMT1.PYMT_NO,SPTPYMT1.INIT_OPER,SPTPYMT1.INIT_DATE," _
                    & "SPTCOOP1.AUTH_NO,SPTCOOP1.SEASON_CODE,SPTCOOP1.VEHICLE_CODE," _
                    & "SPTCOOP1.DATE_START,SPTCOOP1.DATE_END,SPTCOOP1.DATE_ACCRUE,SPTCOOP1.OPS_YYYYPP_ACCRUE," _
                    & "SPTPYMT1.PYMT_CTL_NO, SPTPYMT1.PYMT_REF_NO, SPTPYMT1.PYMT_REF_DATE, SPTPYMT1.PYMT_TYPE, SPTPYMT1.VEND_CODE," _
                    & "SPTPYMT3.DIST_AMT,DIST_AMT_PYMT,SPTPYMT3.BRAND_CODE,SPTPYMT3.TRADE_CLASS_CODE"
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "Sum (SPTPYMT3.DIST_AMT_PYMT) AMT", SQL1)
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "group by SPTPYMT1.OPS_YYYYPP, SPTPYMT1.CUST_CODE, SPTCOOP1.EXPENSE_TYPE_CODE, SPTPYMT3.COLLECTION_CODE, SPTCOOP1.BOOKING_NAME", "", , 1)

        End Select
    End Sub

End Class
