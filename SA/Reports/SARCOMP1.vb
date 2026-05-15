Public Class SARCOMP1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -12)
        Dim LYP_NEXT As String = ASCMAIN1.Period_Calc(RYP, -12 + 1)
        Dim RYP_NEXT As String = ASCMAIN1.Period_Calc(RYP, +1)

        Dim RYP_01 As String = Mid(RYP, 1, 4) & "01"
        Dim LYP_01 As String = Mid(LYP, 1, 4) & "01"

        Dim SATCOMP1 As String = ""

        'sql = "Select GLTACCT1.* from GLTACCT1,(SELECT DISTINCT ACCT_CODE FROM " & TTT & ") TTT where GLTACCT1.ACCT_CODE = TTT.ACCT_CODE"
        'dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1", 1))

        Dim FX As String = ""
        Dim FACTOR As Decimal = 1
        If Absx1.chkFor("THOUSANDS").Checked Then
            FACTOR = 1000
        End If
        If FACTOR <> 1 Then
            FX = "/" & CStr(FACTOR)
        End If


        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        'Stop

        Dim SOURCE_TABLE_NAME As String = ""
        Dim by_Item As Boolean = False
        If COLUMN_NAMEs.Contains("ITEM_CODE") Or InStr(sql_TABLE_NAMEs, "ICTITEM1") <> 0 Then ' THIS NEEDS TO BE EXPANDED UPON
            by_Item = True
        End If
        Dim by_Store As Boolean = False
        If COLUMN_NAMEs.Contains("CUST_STORE_NO") Or InStr(sql_TABLE_NAMEs, "ARTCUST2") <> 0 Then ' THIS NEEDS TO BE EXPANDED UPON
            by_Store = True
        End If

        'by_Item = True
        'by_Store = True

        If by_Item And by_Store Then
            SOURCE_TABLE_NAME = "SOTINVH2"
        Else
            If Not by_Item And Not by_Store Then
                SOURCE_TABLE_NAME = "SATSSUM0"
            Else
                If by_Item Then
                    SOURCE_TABLE_NAME = "SATSSUMI"
                Else
                    SOURCE_TABLE_NAME = "SATSSUMS"
                End If
            End If
        End If

        SOURCE_TABLE_NAME = "SOTINVH2"

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows
            Dim DATA_TYPE As String = rowASTRECAP.Item("DATA_TYPE")
            Dim YP As String = ""
            Dim COLUMN_NAME As String = ""
            Select Case SOURCE_TABLE_NAME
                Case "SOTINVH2"
                    YP = "SATCOMP1.ORDR_YYYYPP_UPDATED"
                    If DATA_TYPE = "UNITS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0)"
                    ElseIf DATA_TYPE = "SALES" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0) * NVL(SATCOMP1.ORDR_UNIT_PRICE,0)" & FX
                    ElseIf DATA_TYPE = "COSTS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0) * DECODE(SATCOMP1.WHSE_CODE,NULL,0,NVL(SATCOMP1.ITEM_UNIT_COST,0))" & FX
                    End If
                Case Else
                    YP = "SATCOMP1.OPS_YYYYPP"
                    If DATA_TYPE = "UNITS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0)"
                    ElseIf DATA_TYPE = "SALES" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_AMT_SHIP,0)" & FX
                    ElseIf DATA_TYPE = "COSTS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_CGS_SHIP,0)" & FX
                    End If
            End Select

            sql_filter = " and " & YP & " BETWEEN '" & LYP_01 & "' AND '" & RYP & "'" & vbCrLf _
            & " and " & COLUMN_NAME & " <> 0"

            Dim sql_Data As String = "" _
            & ", Sum (CASE WHEN " & YP & " = '" & RYP & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) TYMTDGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " = '" & RYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) TYMTDNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " = '" & LYP & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYMONGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " = '" & LYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) LYMONNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & RYP_01 & "' AND '" & RYP & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) TYYTDGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & RYP_01 & "' AND '" & RYP & "' AND SATCOMP1.INV_TYPE = 'C' THEN " & COLUMN_NAME & " ELSE 0 END) TYYTDRTN" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & RYP_01 & "' AND '" & RYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) TYYTDNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & LYP_01 & "' AND '" & LYP & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYYTDGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & LYP_01 & "' AND '" & LYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) LYYTDNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & LYP_01 & "' AND '" & Mid(LYP, 1, 4) & "12" & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYTOTGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & LYP_NEXT & "' AND '" & Mid(LYP, 1, 4) & "12" & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYTOTTOGO" & vbCrLf

            sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
            & " from " & SOURCE_TABLE_NAME & " SATCOMP1 " & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

            ASCMAIN1.sql = "Insert into " & ASTSRPT1 _
            & "(G1,G2,G3,G4,G5,G6,G7,G8,G9" _
            & ",ASTSRPT1_RECAP_ROW_NO" _
            & COLUMN_NAMEs_appended _
            & ",TYMTDGRS,TYMTDNET,LYMONGRS,LYMONNET" _
            & ",TYYTDGRS,TYYTDRTN,TYYTDNET" _
            & ",LYYTDGRS,LYYTDNET,LYTOTGRS,LYTOTTOGO)" & vbCrLf _
            & " (" & sql & ")"

            ASCDATA1.ExecuteSQL()

        Next


        MyBase.Get_SQL("B")
        sql_filter = " and SATBUDW1.OPS_YYYY = '" & Mid(RYP, 1, 4) & "'"

        ' Stop
        Dim SATBUDW1 As String = TAC.SOCMAIN1.Setup_Budgets_by_Customer

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows
            Dim DATA_TYPE As String = rowASTRECAP.Item("DATA_TYPE")
            'If DATA_TYPE = "SALES-LATER" Then
            If DATA_TYPE = "SALES" Then

                Dim budYTD As String = ""
                Dim budTOT As String = ""
                Dim budTOGO As String = ""
                For i As Integer = 1 To 12
                    Dim B As String = "+NVL(SATBUDW1.WB_P" & Format(i, "00") & ",0)"
                    budTOT &= B
                    If i > Val(Mid(RYP, 5, 2)) Then
                        budTOGO &= B
                    Else
                        budYTD &= B
                    End If
                Next
                If budTOGO = "" Then budTOGO = " 0"
                Dim sql_Data As String = "" _
                & ", Sum (NVL(SATBUDW1.WB_P" & Mid(RYP, 5, 2) & ",0))" & FX & " TYMTDBUD" _
                & ", Sum (" & Mid(budYTD, 2) & ")" & FX & " TYYTDBUD" _
                & ", Sum (" & Mid(budTOT, 2) & ")" & FX & " TYTOTBUD" _
                & ", Sum (" & Mid(budTOGO, 2) & ")" & FX & " TYTOTBUDTOGO"

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & sql_Data _
                & " from " & SATBUDW1 & " SATBUDW1" & sql_TABLE_NAMEs _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" & COLUMN_NAMEs_appended _
                & ",TYMTDBUD,TYYTDBUD,TYTOTBUD,TYTOTBUDTOGO)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()
            End If
        Next

    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "G1,G2,G3,G4,G5,G6,G7,G8,G9")
            Dim TYYTDGRS As Decimal = Val(rowASTSRPT1.Item("TYYTDGRS") & "")
            Dim TYYTDBUD As Decimal = Val(rowASTSRPT1.Item("TYYTDBUD") & "")
            Dim LYYTDGRS As Decimal = Val(rowASTSRPT1.Item("LYYTDGRS") & "")
            Dim TYYTDNET As Decimal = Val(rowASTSRPT1.Item("TYYTDNET") & "")
            Dim LYYTDNET As Decimal = Val(rowASTSRPT1.Item("LYYTDNET") & "")

            rowASTSRPT1.Item("YTDVARGRSBUDAMT") = TYYTDGRS - TYYTDBUD
            rowASTSRPT1.Item("YTDVARGRSLYAMT") = TYYTDGRS - LYYTDGRS
            If TYYTDBUD <> 0 Then rowASTSRPT1.Item("YTDVARGRSBUDPCT") = IIf(TYYTDBUD = 0, 0, 100 * TYYTDGRS / TYYTDBUD)
            If LYYTDGRS <> 0 Then rowASTSRPT1.Item("YTDVARGRSLYPCT") = IIf(LYYTDGRS = 0, 0, 100 * TYYTDGRS / LYYTDGRS)
            rowASTSRPT1.Item("YTDVARNETBUDAMT") = TYYTDNET - TYYTDBUD
            rowASTSRPT1.Item("YTDVARNETLYAMT") = TYYTDNET - LYYTDNET
            If TYYTDBUD <> 0 Then rowASTSRPT1.Item("YTDVARNETBUDPCT") = IIf(TYYTDBUD = 0, 0, 100 * TYYTDNET / TYYTDBUD)
            If LYYTDNET <> 0 Then rowASTSRPT1.Item("YTDVARNETLYPCT") = IIf(LYYTDNET = 0, 0, 100 * TYYTDNET / LYYTDNET)
        Next
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        If Not tf Then
            SplitContainer5.Panel2Collapsed = True
        End If
    End Sub

    Public Overrides Sub CheckedChanged_Special(ByVal COLUMN_NAME As String, ByVal chk As UltraWinEditors.UltraCheckEditor)
        'Set_Recaps()
    End Sub

    Sub Set_Recaps()
        tblASTRECAP.Rows.Clear()
        For i As Integer = 0 To 2
            Dim DATA_TYPE As String = New String() _
            {"UNITS", "SALES", "COSTS"}(i)
            If Absx1.chkFor("INCL_" & DATA_TYPE).Checked Then
                Dim rowASTRECAP As DataRow = tblASTRECAP.NewRow
                rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") = i + 1
                rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CAPTION") = Absx1.chkFor("INCL_" & DATA_TYPE).Text
                rowASTRECAP.Item("DATA_TYPE") = DATA_TYPE
                tblASTRECAP.Rows.Add(rowASTRECAP)
            End If
        Next
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "Y", "N"))
        CR_params.Add("SHOW_TOTALS", "1")
        CR_params.Add("YP_LEGEND", ASCMAIN1.Get_Legend(RYP))
        CR_params.Add("OPTVAR", Absx1.optFor("OPTVAR").Value)
        Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Report Period"
            End If
        End If
    End Sub

    Overrides Sub Verify_Special_Pre(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Set_Recaps()
        End If
    End Sub
End Class