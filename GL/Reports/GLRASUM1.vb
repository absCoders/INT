Public Class GLRASUM1

    Dim NP As Integer
    Dim OFFSET As Integer
    Dim OFFSET_Y As Integer
    Dim AYP(,) As String
    Dim OPTYE As Integer

    Dim BY_SEG2 As Boolean
    Dim BY_SEG3 As Boolean
    Dim BY_SEG4 As Boolean


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        Call Breakout_By()

        Set_cmbYP("RYP0", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -60, 12, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0")

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables

        Dim valYEAR As New Dictionary(Of String, Integer)
        valYEAR.Add("NY", 1)
        valYEAR.Add("TY", 0)
        valYEAR.Add("LY", -1)
        valYEAR.Add("2LY", -2)
        valYEAR.Add("3LY", -3)
        valYEAR.Add("4LY", -4)
        valYEAR.Add("5LY", -5)

        Dim TY As String = Mid$(RYP0, 1, 4)

        Dim YEAR_BEG As Integer = 0
        Dim YEAR_END As Integer = 0
        Dim ROWYEAR As Integer = 0
        For Each row As DataRow In tblASTRECAP.Rows
            ROWYEAR = valYEAR(row.Item("YEAR"))
            If ROWYEAR < YEAR_BEG Then
                YEAR_BEG = ROWYEAR
            End If
            If ROWYEAR > YEAR_END Then
                YEAR_END = ROWYEAR
            End If
        Next

        Dim LY As String = Mid$(ASCMAIN1.Period_Calc(RYP1, -12), 1, 4)

        Dim TT As String = GL_Prep(Format(TY + YEAR_BEG, "0000"), Format(TY + YEAR_END, "0000"))
        Dim TTB As String = GL_Prep(Format(TY + YEAR_BEG, "0000"), Format(TY + YEAR_END, "0000"), True)

        Dim RYM As String = ASCMAIN1.Get_YYYYMM(RYP1, 0)

        'OPTYE = Val(Absx1.optFor("OPTYE").Value & "")

        'OFFSET = ASCMAIN1.PCO
        'If OFFSET <= 0 Then
        '    OFFSET = OFFSET + 12
        'End If
        'Dim z As String = "*01*02*03*04*05*06*07*08*09*10*11*12"
        'z = z & Mid(z, 1, 3 * OFFSET)
        'z = Mid(z, 1, Len(z) - 3)
        'Dim l As Integer = ((12 - Val(Mid$(RYM, 5, 2)) + OFFSET) Mod 12) * 3
        'z = Mid(z, Len(z) - l + 1, l)
        'If InStr(z, "*" & Format$(OPTYE, "00")) <> 0 Then
        '    OFFSET_Y = 1
        'Else
        '    OFFSET_Y = 0
        'End If
        'OFFSET = OPTYE - OFFSET
        'If OFFSET < 0 Then
        '    OFFSET = OFFSET + 12
        'End If

        Dim AYP0 As String
        AYP0 = ASCMAIN1.Period_Calc(RYP0, -1 * OFFSET)
        Dim AYP1 As String
        AYP1 = ASCMAIN1.Period_Calc(RYP1, -1 * OFFSET)
        NP = ASCMAIN1.Period_Diff(AYP0, AYP1) + 1
        ReDim AYP(12, 2)
        For i As Integer = 1 To NP
            AYP(i, 1) = ASCMAIN1.Period_Calc(AYP0, i - 1)
            AYP(i, 2) = ASCMAIN1.Get_Legend(AYP(i, 1))
        Next i

        BY_SEG2 = Absx1.chkFor("SEG2_CODE").Checked Or SQLA("SEG2_CODE", "SEQUENCE") <> ""
        BY_SEG3 = Absx1.chkFor("SEG3_CODE").Checked Or SQLA("SEG3_CODE", "SEQUENCE") <> ""
        BY_SEG4 = Absx1.chkFor("SEG4_CODE").Checked Or SQLA("SEG4_CODE", "SEQUENCE") <> ""

        Dim sqlx As String = ""
        Dim sqlx_group_by As String = ""

        For i As Integer = 2 To 4
            Dim z As String = "SEG" & CStr(i) & "_CODE"
            If Not New Boolean() {BY_SEG2, BY_SEG3, BY_SEG4}(i - 2) Then
                z = "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "' " & z
            Else
                sqlx_group_by = sqlx_group_by & ", X." & z
            End If
            sqlx = sqlx & ", " & z
        Next

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        Dim BAL_ROWS As String = ""

        For Each row As DataRow In tblASTRECAP.Select("", "ASTSRPT1_RECAP_ROW_NO")

            For L As Integer = 1 To 2
                If L = 1 Or row.Item("DATA_TYPE") = "VARAB" Or row.Item("DATA_TYPE") = "VARPY" Then
                    Dim YEAR_ADJ As Integer = 0
                    If L = 2 And row.Item("DATA_TYPE") = "VARPY" Then
                        YEAR_ADJ = -1
                    End If
                    sql_filter = " and GLTACCT3.ACCT_YEAR = '" & Format(Val(TY) + valYEAR(row.Item("YEAR")) + YEAR_ADJ, "0000") & "'"

                    Dim F As String = ""
                    If L = 2 Then
                        F = "-1 * "
                    End If
                    sql = "Select " & sql_SELECT_cols _
                        & ", GLTACCT3.ACCT_CODE" & sqlx & ", " & row.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO"
                    sql &= ", " & F & "NVL(ACCT_BEG_BAL,0) ACCT_AMT_P00"
                    Dim ACT_BUD As String = "ACT"
                    If (L = 1 And row.Item("DATA_TYPE") = "BUD") _
                    Or (L = 2 And row.Item("DATA_TYPE") = "VARAB") Then
                        ACT_BUD = "BUD"
                    End If

                    For i As Integer = 1 To 13
                        sql &= ", " & F & "NVL(ACCT_" & ACT_BUD & "_P" & Format(i, "00") & ",0)"
                    Next
                    sql &= ", 0 ACCT_AMT_TOTAL"
                    Dim TT_SOURCE As String = TT
                    If (L = 1 And row.Item("DATA_TYPE") = "BUD") _
                    Or (L = 2 And row.Item("DATA_TYPE") = "VARAB") Then
                        TT_SOURCE = TTB
                    End If
                    sql = sql & " from " & TT_SOURCE & " GLTACCT3 " & sql_TABLE_NAMEs
                    sql = sql & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter)
                    'sql = sql & " group by " & sql_GROUP_BY_cols & y
                    ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")
                End If
            Next
            If row.Item("BAL_ACT") = "BAL" Then
                BAL_ROWS &= ", " & CStr(row.Item("ASTSRPT1_RECAP_ROW_NO"))
            End If

        Next

        If BAL_ROWS <> "" Then
            For i As Integer = 1 To 12
                sql = "Update " & ASTSRPT1 _
                    & " Set ACCT_AMT_P" & Format(i, "00") _
                    & " = ACCT_AMT_P" & Format(i, "00") _
                    & " + ACCT_AMT_P" & Format(i - 1, "00")
                sql = sql & " WHERE ASTSRPT1_RECAP_ROW_NO IN (" & Mid(BAL_ROWS, 2) & ")"
                ASCDATA1.ExecuteSQL(sql)
            Next
        End If

        Call Special_Routines_for_ACCT_TYPE()

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = "For " & RYPLEGEND0 & " thru " & RYPLEGEND1 & " with Year Ending " & Format$(Format$(OPTYE, "00") & "/01/01", "Mmm")

        Dim YP As String
        Dim YPlegend As String
        For i As Integer = 1 To 12
            If i <= NP Then
                YP = ASCMAIN1.Period_Calc(AYP(i, 1), OFFSET)
                YPlegend = ASCMAIN1.Get_Legend(YP)
                CR_params.Add("P" & Format$(i, "00"), Mid$(YPlegend, 9, 8))
            Else
                CR_params.Add("P" & Format$(i, "00"), "")
            End If
        Next i

        CR_params.Add("SEG2_DESC", IIf(BY_SEG2, ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "", ""))
        CR_params.Add("SEG3_DESC", IIf(BY_SEG3, ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "", ""))
        CR_params.Add("SEG4_DESC", IIf(BY_SEG4, ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "", ""))

        Generate_Report(RPT, , SUBT)

    End Sub
End Class