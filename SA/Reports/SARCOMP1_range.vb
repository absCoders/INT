Public Class SARCOMP1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Call Range_Events(grpINV_DATE_RANGE)
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, -12)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 0, -1)
        'Set_cmbYP("RYP", ASCMAIN1.CYP, -12, 0, -1)

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim LYP0 As String = ASCMAIN1.Period_Calc(RYP0, -12)
        Dim LYP1 As String = ASCMAIN1.Period_Calc(RYP1, -12)
        Dim LYP1_NEXT As String = ASCMAIN1.Period_Calc(RYP1, -12 + 1)
        Dim RYP1_NEXT As String = ASCMAIN1.Period_Calc(RYP1, +1)

        Dim SATCOMP1 As String = ""

        Dim DATA_TYPEs As New List(Of String)
        If Absx1.chkFor("INCL_UNITS").Checked Then
            DATA_TYPEs.Add("UNITS")
        End If
        If Absx1.chkFor("INCL_SALES").Checked Then
            DATA_TYPEs.Add("SALES")
        End If
        If Absx1.chkFor("INCL_COSTS").Checked Then
            DATA_TYPEs.Add("COSTS")
        End If

        'sql = "Select GLTACCT1.* from GLTACCT1,(SELECT DISTINCT ACCT_CODE FROM " & TTT & ") TTT where GLTACCT1.ACCT_CODE = TTT.ACCT_CODE"
        'dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1", 1))


        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        'Stop

        Dim SOURCE_TABLE_NAME As String = ""
        Dim by_Item As Boolean = False
        If COLUMN_NAMEs.Contains("ITEM_CODE") Then ' THIS NEEDS TO BE EXPANDED UPON
            by_Item = True
        End If
        Dim by_Store As Boolean = False
        If COLUMN_NAMEs.Contains("CUST_STORE_NO") Then ' THIS NEEDS TO BE EXPANDED UPON
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

        For Each DATA_TYPE As String In DATA_TYPEs
            Dim YP As String = ""
            Dim COLUMN_NAME As String = ""
            Select Case SOURCE_TABLE_NAME
                Case "SOTINVH2"
                    YP = "SATCOMP1.ORDR_YYYYPP_UPDATED"
                    If DATA_TYPE = "UNITS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0)"
                    ElseIf DATA_TYPE = "SALES" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0) * NVL(SATCOMP1.ORDR_UNIT_PRICE,0)"
                    ElseIf DATA_TYPE = "COSTS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0) * NVL(SATCOMP1.ITEM_UNIT_COST,0)"
                    End If
                Case Else
                    YP = "SATCOMP1.OPS_YYYYPP"
                    If DATA_TYPE = "UNITS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_QTY_SHIP,0)"
                    ElseIf DATA_TYPE = "SALES" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_AMT_SHIP,0)"
                    ElseIf DATA_TYPE = "COSTS" Then
                        COLUMN_NAME = "NVL(SATCOMP1.ORDR_CGS_SHIP,0)"
                    End If
            End Select

            sql_filter = " and " & YP & " BETWEEN '" & RYP0 & "' AND '" & RYP1 & "'"

            Dim sql_Data As String = "" _
            & ", Sum (CASE WHEN " & YP & " = '" & RYP1 & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) TYMTDGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " = '" & RYP1 & "' THEN " & COLUMN_NAME & " ELSE 0 END) TYMTDNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " = '" & LYP1 & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYMONGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " = '" & LYP1 & "' THEN " & COLUMN_NAME & " ELSE 0 END) LYMONNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & RYP0 & "' AND '" & RYP1 & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) TYYTDGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & RYP0 & "' AND '" & RYP1 & "' AND SATCOMP1.INV_TYPE = 'C' THEN " & COLUMN_NAME & " ELSE 0 END) TYYTDRTN" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & RYP0 & "' AND '" & RYP1 & "' THEN " & COLUMN_NAME & " ELSE 0 END) TYYTDNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & LYP0 & "' AND '" & LYP1 & "' AND SATCOMP1.INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYYTDGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & LYP0 & "' AND '" & LYP1 & "' THEN " & COLUMN_NAME & " ELSE 0 END) LYYTDNET" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & Mid(LYP0, 1, 4) & "01" & "' AND '" & Mid(LYP0, 1, 4) & "12" & "' AND INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYTOTGRS" & vbCrLf _
            & ", Sum (CASE WHEN " & YP & " BETWEEN '" & LYP1_NEXT & "' AND '" & Mid(LYP1, 1, 4) & "12" & "' AND INV_TYPE = 'I' THEN " & COLUMN_NAME & " ELSE 0 END) LYTOTTOGO" & vbCrLf

            sql = "Select " & sql_SELECT_cols & ",'" & DATA_TYPE & "' DATA_TYPE " & vbCrLf & sql_Data _
            & " from " & SOURCE_TABLE_NAME & " SATCOMP1 " & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

            ASCMAIN1.sql = "Insert into " & ASTSRPT1 _
            & "(G1,G2,G3,G4,G5,G6,G7,G8,G9" _
            & COLUMN_NAMEs_appended _
            & ",TYMTDGRS,TYMTDNET,LYMONGRS,LYMONNET" _
            & ",TYYTDGRS,TYYTDRTN,TYYTDNET" _
            & ",LYYTDGRS,LYLYDNET,LYTOTGRS,LYTOTTOGO)" _
            & " (" & sql & ")"

            ASCDATA1.ExecuteSQL()
        Next


        'Call MyBase.Get_SQL("B")
        sql_filter = " and SOTBUDD1.OPS_YYYY = '" & Mid(RYP1, 1, 4) & "'"

        ' Stop

        For Each DATA_TYPE As String In DATA_TYPEs
            If DATA_TYPE = "SALES-LATER" Then

                Dim budYTD As String = ""
                Dim budTOT As String = ""
                Dim budTOGO As String = ""
                For i As Integer = 1 To 12
                    Dim B As String = "+NVL(SOTUDD1.BUDGET_P" & Format(i, "00") & ",0)"
                    budTOT &= B
                    If i > Val(Mid(RYP1, 5, 2)) Then
                        budTOGO &= B
                    Else
                        budYTD &= B
                    End If
                Next

                Dim sql_Data As String = "" _
                & ", Sum (NVL(SOTBUDD1.BUDGET_P" & Mid(RYP, 5, 2) & ",0)) TYMTDBUD" _
                & ", Sum (" & Mid(budYTD, 2) & ") TYYTDBUD" _
                & ", Sum (" & Mid(budTOT, 2) & ") TYTOTBUD" _
                & ", Sum (" & Mid(budTOGO, 2) & ") TYTOTBUDTOGO"

                sql = "Select " & sql_SELECT_cols & "," & DATA_TYPE & sql_Data _
                & " from SOTBUDD1" & sql_TABLE_NAMEs _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) _
                & " group by " & sql_GROUP_BY_cols

                ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 _
                & "(G1,G2,G3,G4,G5,G6,G7,G8,G9" _
                & COLUMN_NAMEs_appended _
                & ",TYMTDBUD,TYYTDBUD,TYTOTBUD,TYTOTBUDTOGO)" _
                & " (" & sql & ")")
                ' Stop

            End If
        Next
        'Call Get_WKCodes("ARTCUSL1", "SREP_CODE", "SOTSREP1")

        'Check_if_Empty("ARTCUSL1")



    End Sub

    Public Overrides Sub Print_Report()

        'RPT = "ARRSTMT1"
        'RPT_TITLE = "AR Statements"
        CR_params.Add("YP1", RYP0)
        CR_params.Add("YP2", RYP1)
        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        CR_params.Add("SHOW_TOTALS", "1") ' Absx1.chkFor("THOUSANDS").Checked)
        Generate_Report(RPT) ' , "Statement Control Sheet")
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP0").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Starting Period"
            End If
            If Absx1.cmbFor("RYP1").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify an Ending Period"
            End If
        End If
    End Sub
End Class