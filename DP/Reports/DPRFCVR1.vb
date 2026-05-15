Public Class DPRFCVR1

    Dim MD(12) As String
    Dim AYP(12) As String
    Dim MOS As Integer
    Dim PCT As Decimal
    Dim LYP As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 12), -60, 0, -12)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        MOS = Val(Absx1.numFor("FPDMOS").Value & "")
        PCT = Val(Absx1.numFor("FPDPCT").Value & "")

        LYP = ASCMAIN1.Period_Calc(RYP, -1 * (MOS - 1))

        For i As Integer = 1 To 12
            Dim ZYP As String = ASCMAIN1.Period_Calc(RYP, i - 12)
            AYP(i) = ZYP
            Dim YYYYMM As String = ASCMAIN1.Get_YYYYMM(ZYP)
            MD(i) = Format(CDate(Mid(YYYYMM, 5, 2) & "/01/" & Mid(YYYYMM, 3, 2)), "MMMyy")
            MD(i) = Mid(MD(i), 1, 3) & "'" & Mid(MD(i), 4)
        Next i

        ' Page0

        Page0.Add("Base Period: " & RYPLEGEND)
        Page0.Add("Sorted By: " & Absx1.optFor("OPTSORTBY").Text)
        Page0.Add("Showing: " & Absx1.optFor("OPTSHOW").Text)
        If Absx1.chkFor("CHKFCONLY").Checked Then Page0.Add(Absx1.chkFor("CHKFCONLY").Text)
        If Absx1.chkFor("CHKSUP").Checked Then Page0.Add(Absx1.chkFor("CHKSUP").Text)
        Page0.Add("Variance Percentage Threshold: " & CStr(PCT))
        Page0.Add("Number of Periods for Variance Pct: " & CStr(MOS))

        ' Prepare Work Tables

        ASCMAIN1.sql = "Select SOTMKTC1.MARKET_CODE, SOTPRIC2.ITEM_CODE, SOTPRIC2.ITEM_PRICE" & vbCrLf _
            & " from SOTPRIC2,ARTCUST1,SOTMKTC1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTMKTC1.CUST_CODE" & vbCrLf _
            & "   and ARTCUST1.PRICE_LIST_CODE = SOTPRIC2.PRICE_LIST_CODE"
        Dim SOTPRIC2 As String = ASCMAIN1.Temp_Table
        'ASCDATA1.ExecuteSQL("Create Index I_" & SOTPRIC2 & "_1 on " & SOTPRIC2 & " (MARKET_CODE, ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTPRIC2 & " add Primary Key (MARKET_CODE, ITEM_CODE)")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""
        Dim sql_filter2 As String = ""
        Dim sql_Having As String = ""

        Dim VZ As String = ""

        Select Case Absx1.optFor("OPTSHOW").Value
            Case "U"
                VZ = "NVL(DPTITMF1.FORECAST,0)"
            Case "S"
                If ASCMAIN1.CLIENT = "AHA" Then
                    VZ = "NVL(DPTITMF1.FORECAST,0) * NVL(ICTITEM1.ITEM_RETAIL_PRICE,0) * .5"
                Else
                    VZ = "NVL(DPTITMF1.FORECAST,0) * NVL(SOTPRIC2.ITEM_PRICE,NVL(ICTITEM1.ITEM_RETAIL_PRICE,0) * (100 - NVL(SOTPCLS1.PRICE_BASE_DPCT,40))/100)"
                End If
            Case "C"
                VZ = "NVL(DPTITMF1.FORECAST,0) * NVL(ICTCOSTA.ITEM_COST_TOTAL,0)"
        End Select

        sql_Cols = ""
        sql_Data = ""
        For i As Integer = 1 To 12
            Dim COLUMN_NAME = "F_" & Format(i, "00")
            sql_Cols &= "," & COLUMN_NAME
            sql_Data &= ", Sum (Decode (DPTITMF1.OPS_YYYYPP,'" & AYP(i) & "'," & VZ & ",0)) " & COLUMN_NAME & vbCrLf
        Next i

        sql_filter2 = "" _
            & "   and DPTITMF1.OPS_YYYYPP <= '" & RYP & "'" _
            & "   and DPTITMF1.OPS_YYYYPP >= '" & LYP & "'" _
            & "   and DPTITMF1.OPS_YYYYPP_FC = DPTITMF1.OPS_YYYYPP"

        ASCMAIN1.Progress("Forecasts")
        MyBase.Get_SQL("F")

        sql_TABLE_NAMEs &= ", " & SOTPRIC2 & " SOTPRIC2, SOTPCLS1, ARTCUST1, SOTMKTC1"
        sql_JOIN &= " and SOTPRIC2.MARKET_CODE (+) = DPTITMF1.MARKET_CODE and SOTPRIC2.ITEM_CODE (+) = DPTITMF1.ITEM_CODE"
        sql_JOIN &= " and SOTMKTC1.MARKET_CODE = DPTITMF1.MARKET_CODE and ARTCUST1.CUST_CODE (+) = SOTMKTC1.CUST_CODE"
        sql_JOIN &= " and SOTPCLS1.PRICE_CLASS_CODE (+) = ARTCUST1.PRICE_CLASS_CODE and SOTPCLS1.PRICE_BASIS (+) = 'R'"

        Get_Data("DPTITMF1", _
                    sql_Sum:=sql_Data, _
                    sql_Sum_Cols:=sql_Cols, _
                    sql_filter:=sql_filter, _
                    sql_filter2:=sql_filter2, _
                    sql_Having:="", _
                    sql_Appended_Cols:=",DPTITMF1.ITEM_CODE")


        Select Case Absx1.optFor("OPTSHOW").Value
            Case "U"
                VZ = "SOTINVH2.ORDR_QTY_SHIP"
            Case "S"
                VZ = "NVL(SOTINVH2.ORDR_QTY_SHIP, 0) * NVL(SOTINVH2.ORDR_UNIT_PRICE, 0)"
            Case "C"
                VZ = "NVL(SOTINVH2.ORDR_QTY_SHIP, 0) * NVL (SOTINVH2.ITEM_UNIT_COST, 0)"
        End Select

        sql_Cols = ""
        sql_Data = ""
        For i As Integer = 1 To 12
            Dim COLUMN_NAME = "S_" & Format(i, "00")
            sql_Cols &= "," & COLUMN_NAME
            sql_Data &= ", Sum (Decode (SOTINVH2.ORDR_YYYYPP_UPDATED,'" & AYP(i) & "'," & VZ & ",0)) " & COLUMN_NAME & vbCrLf
        Next i

        sql_filter2 = "" _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'" _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & LYP & "'" _
            & "   and SOTINVH2.INV_TYPE = 'I'"

        If Absx1.chkFor("CHKFCONLY").Checked Then
            sql_filter2 &= " and SOTINVH2.ITEM_CODE in " _
                & " (Select Distinct ITEM_CODE from DPTITMF1 where OPS_YYYYPP between '" & LYP & "' and '" & RYP & "')"

        End If
        '& "   and DPTITMF1.OPS_YYYYPP_FC = SOTINVH2.ORDR_YYYYPP_UPDATED" _
        '& "   and DPTITMF1.ITEM_CODE = SOTINVH2.ITEM_CODE"

        ASCMAIN1.Progress("Shipments")
        MyBase.Get_SQL("S")
        'If Not sql_TABLE_NAMEs.Contains("DPTITMF1") Then
        '    sql_TABLE_NAMEs &= ", DPTITMF1"
        'End If
        Get_Data("SOTINVH2", _
                    sql_Sum:=sql_Data, _
                    sql_Sum_Cols:=sql_Cols, _
                    sql_filter:=sql_filter, _
                    sql_filter2:=sql_filter2, _
                    sql_Having:="", _
                    sql_Appended_Cols:=",SOTINVH2.ITEM_CODE")

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE in (Select Distinct ITEM_CODE from " & ASTSRPT1 & ")"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)
        Fill_Records("ICTITEM1")
    End Sub

    Public Overrides Sub Print_Report()

        SUBT = "Showing " & Absx1.optFor("OPTSHOW").Text _
            & IIf(Absx1.chkFor("CHKSUP").Checked, "", ", Sorted by " & Absx1.optFor("OPTSORTBY").Text)

        CR_params.Add("FCONLY", IIf(Absx1.chkFor("CHKFCONLY").Checked, "1", "0"))

        For i As Integer = 1 To 12
            CR_params.Add("MD" & Format(i, "00"), MD(i))
        Next i

        CR_params.Add("SUP", IIf(Absx1.chkFor("CHKSUP").Checked, "1", "0"))

        CR_params.Add("MD1", Mid(ASCMAIN1.Get_Legend(LYP), 10, 6))
        CR_params.Add("MD2", Mid(ASCMAIN1.Get_Legend(RYP), 10, 6))
        CR_params.Add("MOS", Absx1.numFor("FPDMOS").Value & "")

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        'If eItemKey = "Proceed" Then
        '    If Absx1.cmbFor("RYP").Value & "" = "" Then
        '        EMsg &= vbCr & "You must Specify a Reporting Period"
        '    End If
        'End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        With dst.Tables("ASTSRPT1")

            Dim exp As String = ""

            exp = "0"
            If Absx1.optFor("OPTSORTBY").Value = "V" Then
                exp = "V_00"
            ElseIf Absx1.optFor("OPTSORTBY").Value = "P" Then
                exp = "P_00"
            Else
                exp = "0"
            End If
            .Columns.Add("SORTBY", GetType(System.Decimal), exp)

            exp = ""
            For i As Integer = 12 To 0 Step -1
                Dim strTemp As String = "_" & Format$(i, "00")
                If i >= 12 - MOS + 1 Then
                    exp &= "+ISNULL(F" & strTemp & ",0)"
                End If
                .Columns("V" & strTemp).Expression = Replace("ISNULL(S_00,0)-ISNULL(F_00,0)", "00", Format(i, "00"))
                .Columns("P" & strTemp).Expression = Replace("IIF(ISNULL(F_00,0)=0,0,V_00 / F_00)", "00", Format(i, "00")) & " * 100"
            Next
            .Columns("F_00").Expression = Mid(exp, 2)
            .Columns("S_00").Expression = Replace(Mid(exp, 2), "F", "S")
        End With

        Dim sqld As String = ""
        If Absx1.chkFor("CHKFCONLY").Checked Then
            sqld &= " or F_00 = 0"
        End If
        If PCT > 0 Then
            'sqld &= " or ABS(P_00) < " & CStr(PCT)
            sqld &= " or (P_00 > -" & CStr(PCT) & " and P_00 < " & CStr(PCT) & ")"
        End If
        If sqld <> "" Then
            ASCDATA1.DeleteRows(dst.Tables("ASTSRPT1"), Mid(sqld, 5))
            ' dst.Tables("ASTSRPT1").AcceptChanges()
        End If
        '  Dim s() As DataRow = dst.Tables("ASTSRPT1").Select(Mid(sqld, 5))

    End Sub
End Class