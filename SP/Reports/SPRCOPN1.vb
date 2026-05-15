Public Class SPRCOPN1
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        Absx1.dteFor("DTE0").Value = CDate("01/01/" & Now.Date.Year)
        Absx1.dteFor("DTE1").Value = Now.Date

        Absx1.cmbFor("RYP").Visible = (optInclude.Value & "" = "P")

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

        MyBase.Get_SQL("*")

        sql_Data = "" _
            & ", COUNT (*) RECORDS" & vbCrLf

        sql_Cols = "" _
            & ",RECORDS"

        sql_filter = ""
        If chkOpenOnly.Checked Then
            sql_filter = "   and SPTCOOP1.STATUS_CODE = 'O'"
        End If

        If optInclude.Value = "P" Then
                sql_filter &= "" _
                & "   and NVL(SPTCOOP1.OPS_YYYYPP_ACCRUE,SPTCOOP1.OPS_YYYYPP) <= '" & RYP & "'"
        End If

        If Not chkAllDates.Checked Then
            sql_filter &= "" _
                & " and SPTCOOP1.DATE_START >= '" & Format(Absx1.dteFor("DTE0").Value, "dd-MMM-yyyy") & "'" _
                & " and SPTCOOP1.DATE_START <= '" & Format(Absx1.dteFor("DTE1").Value, "dd-MMM-yyyy") & "'"
        End If

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & Replace(COLUMN_NAMEs_appended, ",", ",SPTCOOP3.") & vbCrLf _
            & sql_Data _
            & " from SPTCOOP3" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols & vbCrLf _
            & IIf(sql_GROUP_BY_cols = "", "", ",") & Mid(Replace(COLUMN_NAMEs_appended, ",", ",SPTCOOP3."), 2)

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        Create_TDA(dst.Tables.Add, "SPTCOOP1", "*")
        With dst.Tables("SPTCOOP1").Columns
            .Add("TOTAL_AMT", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
            .Add("CANCEL_AMT", GetType(System.Decimal), "IIF(ISNULL(TOTAL_AMT,0) - ISNULL(PAID_AMT,0) - ISNULL(OPEN_AMT,0)<=0,0,ISNULL(TOTAL_AMT,0) - ISNULL(PAID_AMT,0) - ISNULL(OPEN_AMT,0))")
        End With

        Create_TDA(dst.Tables.Add, "SPTCOOP3", "*")
        Create_Relation("SPTCOOP1", "SPTCOOP3", "AUTH_NO")
        With dst.Tables("SPTCOOP3").Columns
            .Add("TOTAL_AMT", GetType(System.Decimal), "PARENT.TOTAL_AMT")
            .Add("OPEN_AMT", GetType(System.Decimal), "PARENT.OPEN_AMT")
            .Add("PAID_AMT", GetType(System.Decimal), "PARENT.PAID_AMT")
            .Add("DIST_AMT_OPEN", GetType(System.Decimal), "IIF(TOTAL_AMT=0,0,OPEN_AMT * ISNULL(DIST_AMT,0) / TOTAL_AMT)")
            .Add("DIST_AMT_PAID", GetType(System.Decimal), "IIF(TOTAL_AMT=0,0,PAID_AMT * ISNULL(DIST_AMT,0) / TOTAL_AMT)")
        End With

        Create_TDA(dst.Tables.Add, "ARTCUST1", "Select CUST_CODE, CUST_NAME from ARTCUST1", 0)

        EnforceConstraints(False)
        sql = "Select * from SPTCOOP1 where AUTH_NO in (Select Distinct AUTH_NO from " & ASTSRPT1 & ")"
        Fill_Records("SPTCOOP1", "", True, sql)
        sql = "Select * from SPTCOOP3 where (AUTH_NO,AUTH_LNO) in (Select AUTH_NO,AUTH_LNO from " & ASTSRPT1 & ")"
        Fill_Records("SPTCOOP3", "", True, sql)
        Fill_Records("ARTCUST1")
        EnforceConstraints(True)
    End Sub

    Public Overrides Sub Print_Report()
        If optInclude.Value = "P" Then
            SUBT = "All Contracts Accrued thru " & cmbRYP.Value
        Else
            SUBT = "All Contracts"
        End If

        If Not chkAllDates.Checked Then
            SUBT &= "" _
                & "; Start Dates " & Format(Absx1.dteFor("DTE0").Value, "MM/dd/yyyy") _
                & " thru " & Format(Absx1.dteFor("DTE1").Value, "MM/dd/yyyy")
        End If

        Generate_Report(RPT, , SUBT)

        Prepare_Data_Extracts()

        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        '    Create_TDA(dst.Tables.Add, "ABS_COOP", "*")
        '    dst.Tables("ABS_COOP").Merge(dst.Tables("SPTCOOP3"))
        '    Update_Record_TDA("ABS_COOP", "1=1")
        'End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
        End If
    End Sub

    Private Sub chkAllDates_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllDates.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Absx1.dteFor("DTE0").Enabled = Not chkAllDates.Checked
        Absx1.dteFor("DTE1").Enabled = Not chkAllDates.Checked
    End Sub

    Private Sub optInclude_ValueChanged(sender As Object, e As EventArgs) Handles optInclude.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Absx1.cmbFor("RYP").Visible = (optInclude.Value & "" = "P")
    End Sub


    Sub Prepare_Data_Extracts()

        Dim TBL As DataTable = dst.Tables("SPTCOOP1").Clone
        TBL.PrimaryKey = Nothing
        With TBL
            .Columns.Add("COLLECTION_CODE")
            .Columns.Add("FEATURE_DESC")
            .Columns.Add("DIST_AMT", GetType(System.Decimal))
            .Columns.Add("DIST_AMT_OPEN", GetType(System.Decimal))
            .Columns.Add("DIST_AMT_PAID", GetType(System.Decimal))
            .Columns.Add("DIST_AMT_TOTAL", GetType(System.Decimal), "ISNULL(DIST_AMT_OPEN,0)+ISNULL(DIST_AMT_PAID,0)")

            '.Columns("DIST_AMT_OPEN").Expression = ""
            '.Columns("DIST_AMT_PAID").Expression = ""
        End With

        For Each rowSPTCOOP1 As DataRow In dst.Tables("SPTCOOP1").Select("")
            Dim AUTH_NO As String = rowSPTCOOP1.Item("AUTH_NO")
            For Each rowSPTCOOP3 As DataRow In dst.Tables("SPTCOOP3").Select($"AUTH_NO = '{AUTH_NO}'")

                Dim row As DataRow = TBL.NewRow
                For Each DCOL As DataColumn In TBL.Columns
                    Dim C As String = DCOL.ColumnName
                    If C <> "COLLECTION_CODE" And C <> "FEATURE_DESC" And C <> "DIST_AMT" And C <> "DIST_AMT_OPEN" And C <> "DIST_AMT_PAID" And C <> "DIST_AMT_TOTAL" Then
                        row.Item(C) = rowSPTCOOP1.Item(C)
                    End If
                Next
                For Each C As String In New String() {"COLLECTION_CODE", "FEATURE_DESC", "DIST_AMT", "DIST_AMT_OPEN", "DIST_AMT_PAID"}
                    row.Item(C) = rowSPTCOOP3.Item(C)
                Next
                TBL.Rows.Add(row)

            Next
        Next

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        grdASTEXPT1.DataSource = TBL
        grdASTEXPT1.Text = RPT_TITLE & "-" & SUBT
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        Set_DX_Column(grdASTEXPT1, "AUTH_NO", "Auth No", 100)
        Set_DX_Column(grdASTEXPT1, "CUST_CODE", "Customer", 100)
        Set_DX_Column(grdASTEXPT1, "DATE_START", "Start", 90, "MM/dd/yy",, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT1, "DATE_END", "End", 90, "MM/dd/yy",, System.Drawing.Color.Orange)
        Set_DX_Column(grdASTEXPT1, "STATUS_CODE", "Sta", 50,,, System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collection", 80,,, System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "BOOKING_NAME", "Brand", 80,,, System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "FEATURE_DESC", "Brand", 80,,, System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "EVENT_TYPE_CODE", "Brand", 80,,, System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VEHICLE_CODE", "Brand", 80,,, System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "VERIFIED_AS_OPEN_DATE", "Verified", 90, "MM/dd/yy")

        Set_DX_Column(grdASTEXPT1, "DIST_AMT", "$Dist", 100, "#,##0.00", , System.Drawing.Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "DIST_AMT_OPEN", "$Dist", 100, "#,##0.00", , System.Drawing.Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "DIST_AMT_PAID", "$Dist", 100, "#,##0.00", , System.Drawing.Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "DIST_AMT_TOTAL", "$Dist", 100, "#,##0.00", , System.Drawing.Color.LightGreen)

        Create_Summary(grdASTEXPT1, New String() {"AUTH_NO"}, "Count")
        Create_Summary(grdASTEXPT1, New String() {"DIST_AMT", "DIST_AMT_OPEN", "DIST_AMT_PAID", "DIST_AMT_TOTAL"})

        'grdASTEXPT1.DisplayLayout.Bands(0).Columns("BRAND_CODE").Header.Fixed = True
        Sort_grdColumns(grdASTEXPT1, "CUST_CODE,AUTH_NO")
    End Sub

End Class