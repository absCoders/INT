Imports System.Math

Public Class APR1099E
    Dim Report_Subt As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("APTPARM1")

        ' Range_Events(grpCHK_DATE_RANGE)
        optSHOW.Value = "W"
        Absx1.txtFor("TIN").Text = ROWs("APTPARM1").Item("AP_PARM_1099_TAX_ID") & ""
        Absx1.txtFor("TCC").Text = "" ' "19K37"
        Absx1.numFor("CUTOFF").Value = Val(ROWs("APTPARM1").Item("AP_PARM_1099_LIMIT") & "")

        Dim YYYY = Now.Date.AddMonths(-6).Year
        Absx1.dteFor("CHK_DATE_F").Value = "01/01/" & YYYY
        Absx1.dteFor("CHK_DATE_L").Value = "12/31/" & YYYY
    End Sub

    Protected Overrides Sub Build_Workfile()
        With dst
            Dim SQLX As String = ""
            SQLX = " FROM APTCHCK1, APTCHCK2, APTINVH1, APTVEND1" _
            & "  WHERE APTCHCK1.CHECK_STATUS = 'I'" _
            & " AND APTCHCK1.CHECK_DATE >= '" & Format(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") & "'" _
            & " AND APTCHCK1.CHECK_DATE <= '" & Format(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") & "'" _
            & IIf(optSHOW.Value = "O", " AND APTVEND1.VEND_TAX_ID IS NULL", "") _
            & IIf(optSHOW.Value = "W", " AND APTVEND1.VEND_TAX_ID IS NOT NULL", "") _
            & IIf(chkDTL.Checked = False, "    AND APTINVH1.INV_1099_AMT <> 0", "") _
            & "    AND APTCHCK1.BANK_CODE = APTCHCK2.BANK_CODE" _
            & "    AND APTCHCK1.CHECK_NUM = APTCHCK2.CHECK_NUM" _
            & "    AND APTINVH1.VOUCHER_NO = APTCHCK2.VOUCHER_NO" _
            & "    AND APTVEND1.VEND_CODE = APTCHCK1.VEND_CODE_AP"


            ASCMAIN1.Progress("Compiling Check data", "")

            ASCMAIN1.sql = "SELECT DISTINCT APTCHCK1.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTCHCK1", 2))

            ASCMAIN1.sql = "SELECT DISTINCT APTCHCK2.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTCHCK2", 3))

            ASCMAIN1.Progress("Evaluating Invoice data", "")
            ASCMAIN1.sql = "SELECT APTINVH1.*, DECODE(APTINVH1.INV_AMT,0,0, " _
            & " APTINVH1.INV_1099_AMT * APTCHCK2.INV_AMT_APPLIED / APTINVH1.INV_AMT) PMT_1099 " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTINVH1", 1))

            ASCMAIN1.Progress("Evaluating Vendor data", "")
            ASCMAIN1.sql = "SELECT DISTINCT APTVEND1.* " & SQLX
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTVEND1", 1))


            ASCMAIN1.Progress("Merging Check, Invoice and Vendor data", "")
            ASCMAIN1.sql = " SELECT APTCHCK1.VEND_CODE_AP, " _
            & " SUM(DECODE(APTINVH1.INV_AMT,0,0,APTINVH1.INV_1099_AMT * APTCHCK2.INV_AMT_APPLIED / APTINVH1.INV_AMT)) AS PMT_1099, " _
            & " '1' AS PRINT_IND" & SQLX _
            & " GROUP BY APTCHCK1.VEND_CODE_AP"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APT1099V", 1))
            .Tables("APT1099V").Columns("PRINT_IND").ReadOnly = False

            Dim CUTOFF As Decimal = Val(Absx1.numFor("CUTOFF").Value & "")
            For Each rowAPT1900V As DataRow In dst.Tables("APT1099V").Select("PMT_1099 < " & CStr(CUTOFF))
                rowAPT1900V.Item("PRINT_IND") = "0"
            Next

            ASCMAIN1.sql = "Select * from APTPARM1 Where AP_PARM_KEY = 'Z'"
            .Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "APTPARM1", 1))
        End With

        Check_if_Empty("APTCHCK1")

    End Sub

    Public Overrides Sub Print_Report()
        Report_Subt = "1099 Details for Payments Made from " & Format$(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") _
        & " to " & Format$(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") _
        & IIf(Absx1.numFor("CUTOFF").Value > 0, " over " & Format$(Absx1.numFor("CUTOFF").Value, "$##,###.00"), "")
        Generate_Report(RPT, , Report_Subt)

        '' 1099 Form
        Report_Subt = ""
        RPT = "APR1099F"
        If ASCMAIN1.CLIENT = "AHA" Then
            RPT = "APR1099A" ' 2 ON A FORM
        End If
        Generate_Report(RPT, "1099 Form", Report_Subt)

        '' Payment Review
        Report_Subt = "Summary of Payments Made from " & Format$(Absx1.dteFor("CHK_DATE_F").Value, "dd-MMM-yyyy") _
        & " to " & Format$(Absx1.dteFor("CHK_DATE_L").Value, "dd-MMM-yyyy") _
        & IIf(Absx1.numFor("CUTOFF").Value > 0, " over " & Format$(Absx1.numFor("CUTOFF").Value, "$##,###.00"), "")
        RPT = "APR1099G"
        Generate_Report(RPT, "Payment Review", Report_Subt)
        Prepare_Data_Extracts()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If Absx1.cmbFor("RYP").Text = "" Then
            '    EMsg &= "You Must Select a Period"
            'End If
        End If

    End Sub
    Sub Prepare_Data_Extracts()
        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        Dim DT As New DataTable("VendorExtract")

        With DT.Columns
            .Add("Recipient Type", GetType(String))
            .Add("Recipient TIN Type", GetType(String))
            .Add("Recipient TIN", GetType(String))
            .Add("R Business Name or Last Name", GetType(String))
            .Add("R First Name", GetType(String))
            .Add("R Middle Name", GetType(String))
            .Add("R Suffix", GetType(String))
            .Add("R Address 1", GetType(String))
            .Add("R Address 2", GetType(String))
            .Add("R City", GetType(String))
            .Add("R State", GetType(String))
            .Add("R Zip or Foreign Postal Code", GetType(String))
            .Add("R Country", GetType(String))
            .Add("R Phone Number", GetType(String))
            .Add("R Email Address", GetType(String))
            .Add("Acct No", GetType(String))
            .Add("Total", GetType(Decimal))
        End With

        Dim redHeaders As String() = {
            "Recipient Type",
            "Recipient TIN Type",
            "Recipient TIN",
            "R Business Name or Last Name",
            "R Address 1",
            "R City",
            "R State",
            "R Zip or Foreign Postal Code",
            "R Country"
        }

        For Each col As UltraWinGrid.UltraGridColumn In grdASTEXPT1.DisplayLayout.Bands(0).Columns
            If redHeaders.Contains(col.Key) Then
                col.Header.Appearance.BackColor = Drawing.Color.LightCoral
            End If
        Next

        For Each rowVendor As DataRow In dst.Tables("APTVEND1").AsEnumerable

            Dim VEND_CODE As String = rowVendor.Item("VEND_CODE")
            Dim INV_1099_AMT As Decimal = dst.Tables("APTINVH1").AsEnumerable() _
            .Where(Function(r) r.Field(Of String)("VEND_CODE") = VEND_CODE) _
            .Sum(Function(r) If(r.IsNull("INV_1099_AMT"), 0D, Convert.ToDecimal(r("INV_1099_AMT"))))

            Dim newRow As DataRow = DT.NewRow()
            newRow("Acct No") = VEND_CODE
            newRow("Total") = INV_1099_AMT

            Dim VEND_TAX_ID As String = rowVendor.Item("VEND_TAX_ID")
            Dim VEND_TAX_ID_TYPE As String = rowVendor.Item("VEND_TAX_ID_TYPE")

            Select Case VEND_TAX_ID_TYPE
                Case "E"
                    newRow("Recipient Type") = "BUSINESS"
                    newRow("Recipient TIN Type") = "EIN"
                    newRow("Recipient TIN") = VEND_TAX_ID.Insert(2, "-")
                Case "S"
                    newRow("Recipient Type") = "INDIVIDUAL"
                    newRow("Recipient TIN Type") = "SSN"
                    newRow("Recipient TIN") = VEND_TAX_ID.Insert(3, "-").Insert(6, "-")
            End Select

            Dim suffixes As String() = {"Jr.", "Sr.", "II", "III", "IV", "V"}
            Dim fullName As String = rowVendor.Item("VEND_NAME") & ""
            Dim nameParts As List(Of String) = fullName.Split(" ").ToList()

            Dim possibleSuffix As String = If(nameParts.Count > 1 AndAlso suffixes.Contains(nameParts.Last()), nameParts.Last(), "")
            If possibleSuffix <> "" Then nameParts.RemoveAt(nameParts.Count - 1)

            If VEND_TAX_ID_TYPE = "S" Then ' Individual
                Select Case nameParts.Count
                    Case 2
                        ' First Last
                        newRow("R First Name") = nameParts(0)
                        newRow("R Business Name or Last Name") = nameParts(1)
                    Case 3
                        If suffixes.Contains(nameParts(2)) Then
                            ' First Last Suffix
                            newRow("R First Name") = nameParts(0)
                            newRow("R Business Name or Last Name") = nameParts(1)
                            newRow("R Suffix") = nameParts(2)
                        Else
                            ' First Middle Last
                            newRow("R First Name") = nameParts(0)
                            newRow("R Middle Name") = nameParts(1)
                            newRow("R Business Name or Last Name") = nameParts(2)
                        End If
                    Case Is > 3
                        If suffixes.Contains(nameParts.Last()) Then
                            ' - First Middle Middle Last Suffix
                            ' - First Middle Last Suffix
                            newRow("R Suffix") = nameParts.Last()
                            nameParts.RemoveAt(nameParts.Count - 1) ' Remove suffix from nameParts
                        End If

                        If nameParts.Count = 3 Then
                            ' First Middle Last (after removing suffix)
                            newRow("R First Name") = nameParts(0)
                            newRow("R Middle Name") = nameParts(1)
                            newRow("R Business Name or Last Name") = nameParts(2)
                        Else
                            ' First Middle Middle Last
                            newRow("R First Name") = nameParts(0)
                            newRow("R Middle Name") = String.Join(" ", nameParts.Skip(1).Take(nameParts.Count - 2))
                            newRow("R Business Name or Last Name") = nameParts.Last()
                        End If
                End Select

                If possibleSuffix <> "" Then newRow("R Suffix") = possibleSuffix
            Else ' Business
                newRow("R Business Name or Last Name") = fullName
            End If

            newRow("R Address 1") = rowVendor.Item("VEND_ADDR1") & ""
            newRow("R Address 2") = rowVendor.Item("VEND_ADDR2") & ""
            newRow("R City") = rowVendor.Item("VEND_CITY") & ""
            newRow("R State") = rowVendor.Item("VEND_STATE") & ""
            newRow("R Zip or Foreign Postal Code") = rowVendor.Item("VEND_ZIP_CODE") & ""
            newRow("R Country") = rowVendor.Item("VEND_COUNTRY") & ""
            newRow("R Phone Number") = rowVendor.Item("VEND_PHONE") & ""
            newRow("R Email Address") = rowVendor.Item("VEND_EMAIL") & ""

            DT.Rows.Add(newRow)
        Next

        grdASTEXPT1.DataSource = DT
        grdASTEXPT1.Text = "Vendor Extract Data"
        UltraTabControl1.Tabs("Data Exports").Visible = True
    End Sub


End Class