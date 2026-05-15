Public Class SARMIXB2
    Dim RYPs(24) As String
    Dim LEGENDs(24) As String

    Dim TTL_UNITS(12) As Double
    Dim TTL_SALES(12) As Double

    Private TOTAL_BUSINESS_UNITS As Int32 = 0
    Private TOTAL_BUSINESS_SALES As Int32 = 0

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim RYP01 As String = Mid(RYP, 1, 4) & "01"

        ' Prepare Work Tables

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        ASCMAIN1.Progress("Shipments")

        MyBase.Get_SQL("*")

        sql_Data = ""
        sql_Cols = ""
        For M As Integer = 1 To 12
            Dim YP As String = Mid(RYP, 1, 4) & Format(M, "00")
            RYPs(M) = YP
            LEGENDs(M) = ASCMAIN1.Get_Legend(YP)

            If M <= Val(Mid(RYP, 5, 2)) Then
                sql_Data &= "" _
               & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) UNITS_" & Format(M, "00") & vbCrLf _
               & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SALES_" & Format(M, "00") & vbCrLf
            Else
                sql_Data &= "" _
                & ", 0 UNITS_" & Format(M, "00") & vbCrLf _
                & ", 0 SALES_" & Format(M, "00") & vbCrLf

            End If
            sql_Cols &= "" _
                & ",UNITS_" & Format(M, "00") & ",SALES_" & Format(M, "00")
        Next M

        ' Previous Years
        For M As Integer = 1 To 12
            Dim YP As String = Mid(ASCMAIN1.Period_Calc(RYP, -12), 1, 4) & Format(M, "00")
            RYPs(M + 12) = YP
            LEGENDs(M + 12) = ASCMAIN1.Get_Legend(YP)

            If M <= Val(Mid(RYP, 5, 2)) Then
                sql_Data &= "" _
                   & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) UNITS_" & Format(M + 12, "00") & vbCrLf _
                   & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SALES_" & Format(M + 12, "00") & vbCrLf
            Else
                sql_Data &= "" _
                & ", 0 UNITS_" & Format(M + 12, "00") & vbCrLf _
                & ", 0 SALES_" & Format(M + 12, "00") & vbCrLf

            End If

 
            sql_Cols &= "" _
                & ",UNITS_" & Format(M + 12, "00") & ",SALES_" & Format(M + 12, "00")

            ' Previous Years
        Next M

        sql_filter = " and SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN '" & ASCMAIN1.Period_Calc(RYP01, -12) & "' and '" & RYP & "'" _
            & " and ICTITEM1.ITEM_SNU_CODE = 'S'"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTINVH2" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        Dim sqlwHC_CODE As String = String.Empty
        If MyBase.Absx1.chkFor("EXCL_HIGH").Checked Then
            sqlwHC_CODE = SQL_in("HC_CODE", "ICTCOLL1.HC_CODE") & vbCrLf
            If Not sqlwHC_CODE.ToUpper.Contains(" NOT IN ") Then
                sqlwHC_CODE = ""
            End If
        End If

        Dim sqlwTRADE_CLASS_CODE As String = String.Empty
        If MyBase.Absx1.chkFor("EXCL_TRADE").Checked Then
            sqlwTRADE_CLASS_CODE = SQL_in("TRADE_CLASS_CODE", "ARTCUST1.TRADE_CLASS_CODE") & vbCrLf
            If Not sqlwTRADE_CLASS_CODE.ToUpper.Contains(" NOT IN ") Then
                sqlwTRADE_CLASS_CODE = ""
            End If
        End If

        'Dim sqlw As String = String.Empty
        'sqlw &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE") & vbCrLf
        'sqlw &= SQL_in("COLLECTION_CODE", "ICTCOLL1.COLLECTION_CODE") & vbCrLf
        'sqlw &= SQL_in("CUST_CODE", "ARTCUST1.CUST_CODE") & vbCrLf
        'sqlw &= SQL_in("ITEM_BASIC_PROMO", "ICTIEM1.ITEM_BASIC_PROMO") & vbCrLf
        'sqlw &= SQL_in("ITEM_CODE", "SOTINVH1.ITEM_CODE") & vbCrLf

        sql = "Select " & sql_SELECT_cols & vbCrLf _
              & "" & vbCrLf & sql_Data _
              & " from SOTINVH2" & sql_TABLE_NAMEs & vbCrLf _
              & ASCMAIN1.SQL_Add_WHERE((sqlwHC_CODE & sqlwTRADE_CLASS_CODE & sql_JOIN & sql_filter).Trim) & vbCrLf _
              & " group by " & sql_GROUP_BY_cols

        Dim TT As String = ASCMAIN1.Temp_Table(sql)

        Dim rowUnits As DataRow = ASCDATA1.GetDataRow("SELECT SUM(NVL(UNITS_01,0)) , SUM(NVL(UNITS_02,0)), SUM(NVL(UNITS_04,0)), SUM(NVL(UNITS_04,0))" _
                                                     & " , SUM(NVL(UNITS_05,0)), SUM(NVL(UNITS_06,0)), SUM(NVL(UNITS_07,0)), SUM(NVL(UNITS_08,0))" _
                                                     & " , SUM(NVL(UNITS_09,0)), SUM(NVL(UNITS_10,0)), SUM(NVL(UNITS_11,0)), SUM(NVL(UNITS_12,0)))" _
                                                     & " FROM " & TT)

        For i As Integer = 1 To 12
            TTL_UNITS(i) = rowUnits.Item(i - 1)
        Next

        Dim rowSales As DataRow = ASCDATA1.GetDataRow("SELECT SUM(NVL(SALES_01,0)) , SUM(NVL(SALES_02,0)), SUM(NVL(SALES_04,0)), SUM(NVL(SALES_04,0))" _
                                             & " , SUM(NVL(SALES_05,0)), SUM(NVL(SALES_06,0)), SUM(NVL(SALES_07,0)), SUM(NVL(SALES_08,0))" _
                                             & " , SUM(NVL(SALES_09,0)), SUM(NVL(SALES_10,0)), SUM(NVL(SALES_11,0)), SUM(NVL(SALES_12,0)))" _
                                             & " FROM " & TT)

        For i As Integer = 1 To 12
            TTL_SALES(i) = rowSales.Item(i - 1)
        Next

        'TOTAL_BUSINESS_UNITS = ASCDATA1.GetDataValue("SELECT SUM(NVL(UNITS_01,0) + NVL(UNITS_02,0) + NVL(UNITS_03,0) + NVL(UNITS_04,0)" _
        '                                             & " + NVL(UNITS_05,0) + NVL(UNITS_06,0) + NVL(UNITS_07,0) + NVL(UNITS_08,0)" _
        '                                             & " + NVL(UNITS_09,0) + NVL(UNITS_10,0) + NVL(UNITS_11,0) + NVL(UNITS_12,0))" _
        '                                             & " FROM " & TT)

        'TOTAL_BUSINESS_SALES = ASCDATA1.GetDataValue("SELECT SUM(NVL(SALES_01,0) + NVL(SALES_02,0) + NVL(SALES_03,0) + NVL(SALES_04,0)" _
        '                                             & " + NVL(SALES_05,0) + NVL(SALES_06,0) + NVL(SALES_07,0) + NVL(SALES_08,0)" _
        '                                             & " + NVL(SALES_09,0) + NVL(SALES_10,0) + NVL(SALES_11,0) + NVL(SALES_12,0))" _
        '                                             & " FROM " & TT)

    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("INCL_UNITS", IIf(Absx1.chkFor("INCL_UNITS").Checked, "1", "0"))
        CR_params.Add("INCL_SALES", IIf(Absx1.chkFor("INCL_SALES").Checked, "1", "0"))
        CR_params.Add("SHOW_PEN", IIf(Absx1.chkFor("SHOW_PEN").Checked, "1", "0"))
        CR_params.Add("SHOW_TTL", IIf(Absx1.chkFor("SHOW_TTL").Checked, "1", "0"))
        CR_params.Add("MM", Mid(RYP, 5, 2))

        CR_params.Add("TOTAL_BUSINESS_UNITS", TOTAL_BUSINESS_UNITS)
        CR_params.Add("TOTAL_BUSINESS_SALES", TOTAL_BUSINESS_SALES)

        CR_params.Add("RPT_YEAR", RYP.Substring(0, 4))
        CR_params.Add("RPT_YEAR_PY", (Val(RYP.Substring(0, 4)) - 1).ToString)

        For i As Integer = 1 To 12
            CR_params.Add("TTL_UNITS_" & i.ToString.Trim.PadLeft(2, "0"), TTL_UNITS(i))
        Next

        For i As Integer = 1 To 12
            CR_params.Add("TTL_SALES_" & i.ToString.Trim.PadLeft(2, "0"), TTL_SALES(i))
        Next

        For MM As Integer = 1 To 12
            'If MM > Val(Mid(RYP, 5, 2)) Then
            '    CR_params.Add("M" & Format(MM, "00"), "")
            'Else
            CR_params.Add("M" & Format(MM, "00"), Mid(LEGENDs(MM), 10, 3))
            'End If
        Next

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Period"
            End If
            If Not Absx1.chkFor("INCL_UNITS").Checked And Not Absx1.chkFor("INCL_SALES").Checked Then
                EMsg &= vbCr & "You must Specify to include Units, Sales or Both"
            End If
        End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'dst.Tables("ASTSRPT1").Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
        'dst.Tables("ASTSRPT1").Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"
    End Sub
End Class