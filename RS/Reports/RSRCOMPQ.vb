Public Class RSRCOMPQ

    Dim FYP As String
    Dim QTR As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 12, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim DATA_TYPE As String = optUS.Value
        Dim CALENDAR As String = optCALENDAR.Value
        Dim YP As String

        If CALENDAR = "O" Then
            FYP = Mid(RYP, 1, 4) & "01"
        Else
            Dim RYM As String = ASCMAIN1.Get_YYYYMM(RYP, 0)
            Dim FYM As String = Mid(RYM, 1, 4) & "02"
            If Mid(RYM, 5, 2) = "01" Then
                Mid(FYM, 1, 4) = Format(Val(Mid(FYM, 1, 4)) - 1, "0000")
            End If
            FYP = ASCMAIN1.Period_Calc(FYM, -1 * ASCMAIN1.PCO)
        End If

        'set Qtr 
        Dim MO As String
        If CALENDAR = "O" Then
            MO = Mid(RYP, 5, 2)
        Else
            MO = Mid(RYP - 1, 5, 2)
        End If

        If MO = "01" Or MO = "02" Or MO = "03" Then
            QTR = "1"
        ElseIf MO = "04" Or MO = "05" Or MO = "06" Then
            QTR = "2"
        ElseIf MO = "07" Or MO = "08" Or MO = "09" Then
            QTR = "3"
        ElseIf MO = "10" Or MO = "11" Or MO = "12" Then
            QTR = "4"
        End If

        Dim TDP As Integer = ASCMAIN1.Period_Diff(FYP, RYP)

        Dim FACTOR As Integer = 1
        If Absx1.chkFor("THOUSANDS").Checked Then
            FACTOR = 1000
        End If

        Dim RSTCOMPQ As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        'Stop

        Dim SOURCE_TABLE_NAME As String = ""
        Dim by_Item As Boolean = False
        If COLUMN_NAMEs.Contains("ITEM_CODE") Then ' THIS NEEDS TO BE EXPANDED UPON TO INCLUDE EVEN THOSE FIELDS THAT ARE DERIVED FROM ICTITEM1
            by_Item = True
        End If

        by_Item = True

        If by_Item Then
            SOURCE_TABLE_NAME = "RSTRETL1"
        Else
            SOURCE_TABLE_NAME = "RSTRETL4"
        End If
        SOURCE_TABLE_NAME = "RSTRETL1"

        Dim COLUMN_NAME As String = ""
        If DATA_TYPE = "UNITS" Then
            COLUMN_NAME = "NVL(RSTCOMPQ.QTY_SOLD,0) / " & CStr(FACTOR)
        ElseIf DATA_TYPE = "SALES" Then
            COLUMN_NAME = "NVL(RSTCOMPQ.AMT_SOLD,0) / " & CStr(FACTOR)
        End If

        sql_filter = " and OPS_YYYYPP BETWEEN '" & ASCMAIN1.Period_Calc(FYP, -12) & "' AND '" & RYP & "'" _
        & " and (NVL(RSTCOMPQ.AMT_SOLD,0) <> 0 or NVL(RSTCOMPQ.QTY_SOLD,0) <>  0) "

        Dim sql_Data As String = ""

        Dim COL_NAMEs As String = ""
        For Y As Integer = 1 To 0 Step -1
            YP = ASCMAIN1.Period_Calc(FYP, -1 * Y * 12)
            For P As Integer = 1 To 12
                Dim COL_NAME As String = IIf(Y = 0, "TY", "LY") & "_ACT_" & Format(P, "00")
                COL_NAMEs &= "," & COL_NAME
                sql_Data &= "" _
                & ", Sum (CASE WHEN OPS_YYYYPP = '" & YP & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & COL_NAME & vbCrLf
                YP = ASCMAIN1.Period_Calc(YP, 1)
            Next
        Next

        sql = "Select " & sql_SELECT_cols & vbCrLf & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " RSTCOMPQ " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended _
        & COL_NAMEs & ")" & vbCrLf _
        & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        ' Budgets

        Call MyBase.Get_SQL("B")
        SOURCE_TABLE_NAME = "RSTBUDR1"

        If DATA_TYPE = "UNITS" Then
            COLUMN_NAME = "0 / " & CStr(FACTOR)
        ElseIf DATA_TYPE = "SALES" Then
            COLUMN_NAME = "NVL(RSTBUDR1.BUDGET,0) / " & CStr(FACTOR)
        End If

        sql_filter = " and OPS_YYYYPP BETWEEN '" & ASCMAIN1.Period_Calc(FYP, -12) & "' AND '" & RYP & "'"

        sql_Data = ""

        COL_NAMEs = ""
        YP = FYP
        For P As Integer = 1 To 12
            ' Dim YPr As String = ASCMAIN1.Period_Calc(YP, ASCMAIN1.PCO - 1)

            'Dim COL_NAME As String = "TY_BUD_" & Mid(YPr, 5, 2)
            Dim COL_NAME As String = "TY_BUD_" & Format(P, "00")
            COL_NAMEs &= "," & COL_NAME
            sql_Data &= "" _
            & ", Sum (DECODE(RSTBUDR1.OPS_YYYYPP,'" & YP & "',BUDGET,0)) " & COL_NAME & vbCrLf
            YP = ASCMAIN1.Period_Calc(YP, 1)
        Next

        sql = "Select " & sql_SELECT_cols & vbCrLf & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " RSTBUDR1 " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended _
        & COL_NAMEs & ")" & vbCrLf _
        & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()



        ' Calculate QTD, STD, and YTD

        For Q As Integer = 1 To 4
            If TDP > (Q - 1) * 3 Then
                Dim SQL As String = ""
                For RP As Integer = 1 To 3
                    Dim P As Integer = (Q - 1) * 3 + RP
                    SQL &= "+ NVL(TY_ACT_" & Format(P, "00") & ",0)"
                Next
                SQL = "Update " & ASTSRPT1 & " Set TY_ACT_Q" & Format(Q, "0") & " = " & Mid(SQL, 2)
                ASCDATA1.ExecuteSQL(SQL)
                ASCDATA1.ExecuteSQL(Replace(SQL, "TY_ACT", "LY_ACT"))
                ASCDATA1.ExecuteSQL(Replace(SQL, "TY_ACT", "TY_BUD"))
            End If
        Next

        For H As Integer = 1 To 2
            sql = "Update " & ASTSRPT1 & " Set TY_ACT_H" & Format(H, "0") _
            & " = NVL(TY_ACT_Q" & Format((H - 1) * 2 + 1, "0") & ",0)" _
            & " + NVL(TY_ACT_Q" & Format((H - 1) * 2 + 2, "0") & ",0)"
            ASCDATA1.ExecuteSQL(sql)
            ASCDATA1.ExecuteSQL(Replace(sql, "TY_ACT", "LY_ACT"))
            ASCDATA1.ExecuteSQL(Replace(sql, "TY_ACT", "TY_BUD"))
        Next

        sql = "Update " & ASTSRPT1 & " Set TY_ACT_Y = NVL(TY_ACT_H1,0) + NVL(TY_ACT_H2,0)"
        ASCDATA1.ExecuteSQL(sql)
        ASCDATA1.ExecuteSQL(Replace(sql, "TY_ACT", "LY_ACT"))
        ASCDATA1.ExecuteSQL(Replace(sql, "TY_ACT", "TY_BUD"))


        For P As Integer = 1 To 12

        Next
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        CR_params.Add("OPTUS", Absx1.optFor("OPTUS").Value)
        CR_params.Add("OPTCALENDAR", Absx1.optFor("OPTCALENDAR").Value)
        CR_params.Add("YP_LEGEND", ASCMAIN1.Get_Legend(RYP))
        CR_params.Add("QTR", QTR)
        Dim HEADING As String = ""
        For M As Integer = 1 To 12
            HEADING = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(FYP, M - 1), False, True)
            CR_params.Add("M" & Format(M, "00"), HEADING)
        Next
        Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Period"
            End If
        End If
    End Sub

    Overrides Sub Format_grdASTDSQL1_pre()

        For Each row As DataRow In tblASTDSQLS.Rows
            Dim COLUMN_NAME = row.Item("COLUMN_NAME")
            Select Case Mid(COLUMN_NAME, 1, 6)
                Case "TY_ACT"
                    row.Item("COLUMN_CAPTION") = "TY Act"
                Case "TY_BUD"
                    row.Item("COLUMN_CAPTION") = "Plan"
                Case "LY_ACT"
                    row.Item("COLUMN_CAPTION") = "LY Act"
            End Select
        Next
    End Sub

    Overrides Sub Format_grdASTDSQL1_post(ByVal grdASTDSQL1 As UltraWinGrid.UltraGrid)

        With grdASTDSQL1.DisplayLayout.Bands(0)
            .ColHeadersVisible = True
            .Columns("DESC_VALUE").Group = .Groups("CODE_VALUE")
            .Groups("CODE_VALUE").Width = .Groups("CODE_VALUE").Width + .Groups("DESC_VALUE").Width
            .Groups("DESC_VALUE").Hidden = True

            Dim SFX As String = ""
            For P As Integer = 1 To 12
                Dim LEGEND As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(FYP, P - 1), False, True)
                Set_Group(grdASTDSQL1, Format(P, "00"), LEGEND, Color.Empty)
                If P Mod 3 = 0 Then
                    SFX = "Q" & Format(P / 3, "0")
                    Set_Group(grdASTDSQL1, SFX, SFX, Color.LightBlue)
                End If
                If P Mod 6 = 0 Then
                    SFX = "H" & Format(P / 6, "0")
                    Set_Group(grdASTDSQL1, SFX, SFX, Color.LightCoral)
                End If
                If P = 12 Then
                    SFX = "Y"
                    Set_Group(grdASTDSQL1, SFX, "Year", Color.Yellow)
                End If
            Next
        End With
    End Sub

    Sub Set_Group( _
    ByVal grdASTDSQL1 As UltraWinGrid.UltraGrid, _
    ByVal SFX As String, _
    ByVal CAPTION As String, _
    ByVal C As System.Drawing.Color)

        With grdASTDSQL1.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add(SFX, CAPTION)
            G.Header.Appearance.TextHAlign = HAlign.Center
            For Each PFX As String In New String() {"TY_ACT", "TY_BUD", "LY_ACT"}
                Dim COLUMN_NAME As String = PFX & "_" & SFX
                .Columns(COLUMN_NAME).Group = G
                .Groups(COLUMN_NAME).Hidden = True
                .Columns(COLUMN_NAME).CellAppearance.BackColor = C
            Next
            G.Header.Appearance.BackColor = G.Columns(0).Header.Appearance.BackColor
            G.Header.Appearance.BackColor2 = G.Columns(0).Header.Appearance.BackColor2
            G.Header.Appearance.BackGradientStyle = G.Columns(0).Header.Appearance.BackGradientStyle
        End With
    End Sub
End Class