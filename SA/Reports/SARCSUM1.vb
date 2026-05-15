Public Class SARCSUM1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""
        'Dim sqlX As String = ""
        'Dim sqlYP As String = ""

        ASCMAIN1.Progress("Customers with Sales History")
        MyBase.Get_SQL("*", "ARTCUST1")

        'sqlX = "NVL(SOTORDR2.ORDR_QTY,0) * NVL(SOTORDR2.ORDR_UNIT_PRICE,0)" & FX
        'sqlYP = "SOTORDR1.ORDR_YYYYPP_BOOKED"

        sql_Data = ""

        sql_Cols = ""

        sql_filter = "" _
            & " and ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from SOTINVH1" _
            & " where ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" _
            & "   and ORDR_YYYYPP_UPDATED <= '" & RYP1 & "')"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ", ARTCUST1.CUST_CODE, 0 LINE" & vbCrLf _
            & sql_Data _
            & " from ARTCUST1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols & vbCrLf _
            & IIf(sql_GROUP_BY_cols = "", "", ",") & "ARTCUST1.CUST_CODE"

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        Build_Customer_Summary()
        Update_Customer_Summary()

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = "Customer Summary from " & RYPLEGEND0 & " thru " & RYPLEGEND1
        ' CR_params.Add("DTL", "Y")
        For I As Integer = 1 To 12
            Dim YP As String = ASCMAIN1.Period_Calc(RYP0, I - 1)
            If YP <= RYP1 Then
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YP, False, True)
                CR_params.Add("M" & Format(I, "00"), LEGEND)
            Else
                CR_params.Add("M" & Format(I, "00"), "")
            End If
        Next
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
        'dst.Tables("ASTSRPT1").Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
        'dst.Tables("ASTSRPT1").Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"

        ' Build_Customer_Summary()
    End Sub

    Sub Build_Customer_Summary()

        Dim ARTCSUMC As String = TAC.SOCMAIN1.Create_ARTCSUMC(Me)

        Dim clear_ARTCSUMB As Boolean = True
        ASCMAIN1.sql = "Select CUST_CODE from " & ASTSRPT1
        For Each rowASTSRPT1 As DataRow In ASCDATA1.GetDataTable.Select() ' dst.Tables("ASTSRPT1").Select("")
            Dim CUST_CODE As String = rowASTSRPT1.Item("CUST_CODE")
            TAC.SOCMAIN1.Create_ARTCSUMA(Me, CUST_CODE)
            TAC.SOCMAIN1.Build_ARTCSUMC(Me, CUST_CODE, RYP0, RYP1, ARTCSUMC, clear_ARTCSUMB)
            clear_ARTCSUMB = False
        Next

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, SREP_CODE, TRADE_CLASS_CODE from ARTCUST1 where CUST_CODE in (Select Distinct CUST_CODE from " & ARTCSUMC & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1", 1))
        dst.Tables("ARTCUST1").Columns.Add("NPP", GetType(System.Decimal))

        For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("")
            Dim CUST_CODE As String = rowARTCUST1.Item("CUST_CODE") & ""
            Dim row() As DataRow
            row = dst.Tables("ARTCSUMA").Select("CUST_CODE = '" & CUST_CODE & "' AND LINE_ABBR = 'GRS'")
            Dim GRS As Decimal = Val(row(0).Item("AMT00") & "")
            row = dst.Tables("ARTCSUMA").Select("CUST_CODE = '" & CUST_CODE & "' AND LINE_ABBR = 'NET'")
            Dim NET As Decimal = Val(row(0).Item("AMT00") & "")
            row = dst.Tables("ARTCSUMA").Select("CUST_CODE = '" & CUST_CODE & "' AND LINE_ABBR = 'NP'")
            Dim NP As Decimal = Val(row(0).Item("AMT00") & "")
            Dim NPP As Decimal = 0
            If NET <> 0 Then NPP = 100 * System.Math.Abs(NP / NET) * System.Math.Sign(NP) ' NP / NET

            '    row(0).Item("LINE_DESC") = "Net Profit " & Format(NPP, "##.0%")
            rowARTCUST1.Item("NPP") = NPP
        Next

    End Sub

    Sub Update_Customer_Summary()
        ASCMAIN1.sql = "Select * from " & ASTSRPT1
        Create_TDA(dst.Tables.Add, ASTSRPT1, "*")

        ASCMAIN1.sql = "Select * from " & ASTSRPT1
        For Each rowASTSRPT1 As DataRow In ASCDATA1.GetDataTable.Select() ' dst.Tables("ASTSRPT1").Select("")
            Dim CUST_CODE As String = rowASTSRPT1.Item("CUST_CODE")
            For Each rowARTCSUMA As DataRow In dst.Tables("ARTCSUMA").Select("CUST_CODE = '" & CUST_CODE & "'")
                Dim row As DataRow = dst.Tables(ASTSRPT1).NewRow
                row.ItemArray = rowASTSRPT1.ItemArray
                For I As Integer = 0 To 13
                    row.Item("AMT" & Format(I, "00")) = rowARTCSUMA.Item("AMT" & Format(I, "00"))
                Next
                row.Item("LINE") = rowARTCSUMA.Item("LINE")
                dst.Tables(ASTSRPT1).Rows.Add(row)
            Next
        Next

        ASCMAIN1.sql = "Delete from " & ASTSRPT1
        ASCDATA1.ExecuteSql

        Update_Record_TDA(ASTSRPT1)

        'ASCDATA1.DeleteRows(dst.Tables(ASTSRPT1), "LINE = 0")

        ASCMAIN1.sql = "Select 0 LINE, 'X' LINE_DESC, 'X' LINE_ABBR from DUAL"
        Dim ARTCSUMD As String = ASCMAIN1.Temp_Table

        ASCDATA1.ExecuteSQL("Delete from " & ARTCSUMD)
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCSUMD & " Modify LINE NUMBER (3,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCSUMD & " Modify LINE_DESC VARCHAR2(100)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCSUMD & " Modify LINE_ABBR VARCHAR2(6)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCSUMD & " Add Primary Key (LINE)")

        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (10,'Gross Shipments','GRS')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (20,'Returns','RTN')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (30,'Net Sales','NET')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (40,'Cost of Goods','CGS')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (50,'$GP on Net Sales','GP')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (60,'Testers/Samples/Misc','ST')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (61,'Gift w/Purchase','GWP')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (64,'Displays','DSP')")
        'ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (65,'Pre-Paid Promo','PP')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (70,'Misc Credits','CR')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (71,'Misc Charges','DR')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (80,'Deductions','DED')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (81,'GL Write-Offs','GL')")
        ASCDATA1.ExecuteSQL("Insert into " & ARTCSUMD & " values (99,'Net Profit','NP')")

        ASCMAIN1.sql = "Select * from " & ARTCSUMD
        Create_TDA(dst.Tables.Add, "ARTCSUMD", "**", 0, False)
        Fill_Records("ARTCSUMD")

    End Sub

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        ASCMAIN1.sql = "Delete from " & TT & " where G1 = '" & aRC & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

End Class