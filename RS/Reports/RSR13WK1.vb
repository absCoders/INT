Public Class RSR13WK1

    Dim WW(13) As String ' Weeks 1 thru 13

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, 0) ' -12)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim DATA_TYPE As String = optUS.Value

        Dim FACTOR As Integer = 1
        If Absx1.chkFor("THOUSANDS").Checked Then
            FACTOR = 1000
        End If

        Dim RST13WK1 As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

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

        Dim COLUMN_NAME As String = ""
        If DATA_TYPE = "UNITS" Then
            COLUMN_NAME = "NVL(RST13WK1.QTY_SOLD,0) / " & CStr(FACTOR)
        ElseIf DATA_TYPE = "SALES" Then
            COLUMN_NAME = "NVL(RST13WK1.AMT_SOLD,0) / " & CStr(FACTOR)
        End If

        Dim sql_Data As String = ""

        Dim DATA_COLUMN_NAMEs As String = ""
        Dim OPS_WEEKS As New List(Of String)
        'TY
        For W As Integer = 1 To 13
            WW(W) = ASCMAIN1.Week_Calc(RYW, -13 + W)
            OPS_WEEKS.Add("'" & WW(W) & "'")
            Dim DATA_COLUMN_NAME = "TY_AMT_" & Format(W, "00")
            DATA_COLUMN_NAMEs &= "," & DATA_COLUMN_NAME
            If chkCUM.Checked Then
                sql_Data &= "" _
                & ", Sum (CASE WHEN RST13WK1.OPS_YYYYWW BETWEEN '" & WW(1) & "' AND '" & WW(W) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & DATA_COLUMN_NAME & vbCrLf
            Else
                sql_Data &= "" _
                & ", Sum (DECODE(RST13WK1.OPS_YYYYWW,'" & WW(W) & "'," & COLUMN_NAME & ",0)) " & DATA_COLUMN_NAME & vbCrLf
            End If
        Next
        Dim COLUMN_NAME_ONH As String = "NVL(RST13WK1.QTY_EOW,0) / " & CStr(FACTOR)
        For W As Integer = 1 To 13
            WW(W) = ASCMAIN1.Week_Calc(RYW, -13 + W)
            Dim DATA_COLUMN_NAME = "TY_ONH_" & Format(W, "00")
            DATA_COLUMN_NAMEs &= "," & DATA_COLUMN_NAME
            'If chkCUM.Checked Then
            '    sql_Data &= "" _
            '    & ", Sum (CASE WHEN RST13WK1.OPS_YYYYWW BETWEEN '" & WW(1) & "' AND '" & WW(W) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & DATA_COLUMN_NAME & vbCrLf
            'Else
            sql_Data &= "" _
            & ", Sum (DECODE(RST13WK1.OPS_YYYYWW,'" & WW(W) & "'," & COLUMN_NAME_ONH & ",0)) " & DATA_COLUMN_NAME & vbCrLf
            'End If
        Next

        'LY
        If chkTYLY.Checked Then
            Dim LYRYW As String = CStr(CInt(RYW.Substring(0, 4)) - 1) & RYW.Substring(4)
            For W As Integer = 1 To 13
                Dim LY_WW As String = ASCMAIN1.Week_Calc(LYRYW, -13 + W)
                OPS_WEEKS.Add("'" & LY_WW & "'")
                Dim DATA_COLUMN_NAME = "LY_AMT_" & Format(W, "00")
                DATA_COLUMN_NAMEs &= "," & DATA_COLUMN_NAME
                If chkCUM.Checked Then
                    sql_Data &= ", Sum(CASE WHEN RST13WK1.OPS_YYYYWW BETWEEN '" & ASCMAIN1.Week_Calc(LYRYW, -13 + 1) & "' AND '" & LY_WW & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & DATA_COLUMN_NAME & vbCrLf
                Else
                    sql_Data &= ", Sum(DECODE(RST13WK1.OPS_YYYYWW,'" & LY_WW & "'," & COLUMN_NAME & ",0)) " & DATA_COLUMN_NAME & vbCrLf
                End If
            Next

            For W As Integer = 1 To 13
                Dim LY_WW As String = ASCMAIN1.Week_Calc(LYRYW, -13 + W)
                Dim DATA_COLUMN_NAME = "LY_ONH_" & Format(W, "00")
                DATA_COLUMN_NAMEs &= "," & DATA_COLUMN_NAME
                sql_Data &= ", Sum(DECODE(RST13WK1.OPS_YYYYWW,'" & LY_WW & "'," & COLUMN_NAME_ONH & ",0)) " & DATA_COLUMN_NAME & vbCrLf
            Next
        End If
        'sql_filter = " and RST13WK1.OPS_YYYYWW BETWEEN '" & WW(1) & "' AND '" & WW(13) & "'"
        sql_filter = " AND RST13WK1.OPS_YYYYWW IN (" & String.Join(",", OPS_WEEKS) & ")"

        sql = "Select " & sql_SELECT_cols & vbCrLf & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " RST13WK1 " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended & DATA_COLUMN_NAMEs & ")" & vbCrLf _
        & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        CR_params.Add("OPTUS", Absx1.optFor("OPTUS").Value)
        CR_params.Add("CHKCUM", chkCUM.ABSChecked)
        CR_params.Add("CHKONH", chkONH.ABSChecked)

        For w As Integer = 1 To 13
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -13 + w)
            ASCMAIN1.sql = "Select WEEK_END_DATE from GLTPARM3 where YYYYWW = '" & YW & "'"
            Dim DT As Date = ASCDATA1.GetDataValue
            CR_params.Add("D" & Format(w, "00"), ASCMAIN1.Get_Legend_Wk(YW, True) & vbCrLf & Format(DT, "MM/dd/yyyy"))
        Next
        CR_params.Add("YW_LEGEND", ASCMAIN1.Get_Legend_Wk(RYW))

        Generate_Report(RPT)
        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        'With dst.Tables("ASTSRPT1").Columns
        '    .Add("ITEM_DESC")
        '    .Add("ITEM_SAFETY_STOCK", GetType(System.Int32))
        '    .Add("ITEM_RETAIL_PRICE", GetType(System.Decimal))
        'End With

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("")
            For I As Integer = 1 To COLUMN_NAMEs.Count
                Dim CODE_VALUE As String = row.Item("G" & CStr(I))
                row.Item("G" & CStr(I)) = Split(CODE_VALUE, ":")(1)
            Next
            'Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            'Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            'row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            'row.Item("ITEM_SAFETY_STOCK") = rowICTITEM1.Item("ITEM_SAFETY_STOCK")
            'row.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        Next

        'With dst.Tables("ASTSRPT1").Columns
        '    .Add("AMT_BOM_PREV", GetType(System.Decimal), "QTY_BOM * STD_COST_LM")
        '    .Add("AMT_REVAL", GetType(System.Decimal), "QTY_BOM * (STD_COST_TM - STD_COST_LM)")
        '    For Each C As String In New String() {"BOM", "SHP", "RTN", "REC", "ADJ", "CON", "EOM"}
        '        .Add("AMT_" & C, GetType(System.Decimal), "QTY_" & C & " * STD_COST_TM")
        '    Next
        'End With

        grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")


        grdASTEXPT1.Text = RPT_TITLE & " - " & SUBT
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , Color.Gold)
        Next

        For w As Integer = 1 To 13
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -13 + w)
            ASCMAIN1.sql = "Select WEEK_END_DATE from GLTPARM3 where YYYYWW = '" & YW & "'"
            Dim DT As Date = ASCDATA1.GetDataValue
            Set_DX_Column(grdASTEXPT1, "TY_AMT_" & Format(w, "00"), "TY " & IIf(optUS.Value = "UNITS", "QTY ", "SLS ") & Format(DT, "MM/dd/yyyy"), 90, "#,##0", , Color.LightBlue)
            Create_Summary(grdASTEXPT1, "TY_AMT_" & Format(w, "00"))
        Next

        For w As Integer = 1 To 13
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -13 + w)
            ASCMAIN1.sql = "Select WEEK_END_DATE from GLTPARM3 where YYYYWW = '" & YW & "'"
            Dim DT As Date = ASCDATA1.GetDataValue
            Set_DX_Column(grdASTEXPT1, "TY_ONH_" & Format(w, "00"), "TY EOW " & Format(DT, "MM/dd/yyyy"), 90, "#,##0", , Color.LightGreen)
            Create_Summary(grdASTEXPT1, "TY_ONH_" & Format(w, "00"))
        Next

        If chkTYLY.Checked Then
            Dim LYRYW As String = CStr(CInt(RYW.Substring(0, 4)) - 1) & RYW.Substring(4)

            For w As Integer = 1 To 13
                Dim LY_YW As String = ASCMAIN1.Week_Calc(LYRYW, -13 + w)
                ASCMAIN1.sql = "Select WEEK_END_DATE from GLTPARM3 where YYYYWW = '" & LY_YW & "'"
                Dim DT_LY As Date = ASCDATA1.GetDataValue
                Set_DX_Column(grdASTEXPT1, "LY_AMT_" & Format(w, "00"), "LY " & IIf(optUS.Value = "UNITS", "QTY ", "SLS ") & Format(DT_LY, "MM/dd/yyyy"), 90, "#,##0", , Color.LightYellow)
                Create_Summary(grdASTEXPT1, "LY_AMT_" & Format(w, "00"))
            Next

            For w As Integer = 1 To 13
                Dim LY_YW As String = ASCMAIN1.Week_Calc(LYRYW, -13 + w)
                ASCMAIN1.sql = "Select WEEK_END_DATE from GLTPARM3 where YYYYWW = '" & LY_YW & "'"
                Dim DT_LY As Date = ASCDATA1.GetDataValue
                Set_DX_Column(grdASTEXPT1, "LY_ONH_" & Format(w, "00"), "LY EOW " & Format(DT_LY, "MM/dd/yyyy"), 90, "#,##0", , Color.Honeydew)
                Create_Summary(grdASTEXPT1, "LY_ONH_" & Format(w, "00"))
            Next
        End If
        grdASTEXPT1.DisplayLayout.Bands(0).Columns("G1").Header.Fixed = True

        If Not chkTYLY.Checked Then
            For w As Integer = 1 To 13
                Dim colLYAmt = "LY_AMT_" & Format(w, "00")
                Dim colLYOnh = "LY_ONH_" & Format(w, "00")

                If grdASTEXPT1.DisplayLayout.Bands(0).Columns.Exists(colLYAmt) Then
                    grdASTEXPT1.DisplayLayout.Bands(0).Columns(colLYAmt).Hidden = True
                End If

                If grdASTEXPT1.DisplayLayout.Bands(0).Columns.Exists(colLYOnh) Then
                    grdASTEXPT1.DisplayLayout.Bands(0).Columns(colLYOnh).Hidden = True
                End If
            Next
        End If


        Sort_grdColumns(grdASTEXPT1, "G1")

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYW").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Week"
            End If
        End If
    End Sub

End Class