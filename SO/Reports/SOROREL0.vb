Public Class SOROREL0
    Dim SOTOREL1 As String = ""
    Dim SOTOREL2 As String = ""
    Dim SOTORDR1 As String
    Dim SOTORDR2 As String
    Dim SHIP_BY_DATE As Date    ' Release all orders Scheduled to ship by
    'SO_PARM_RELEASE_DAYS_AHEAD
    Dim sqlSOTORDR1 As String = ""
    Dim sqlSOTORDR2 As String = ""
    Dim sqlSOTOREL1 As String = ""
    Dim sqlSOTOREL2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")
        dteSHIP_DATE.CalendarInfo.MaxSelectedDays = 1
        Dim SO_PARM_RELEASE_DAYS_AHEAD As Int64 = Val(ROWs.Item("SOTPARM1").Item("SO_PARM_RELEASE_DAYS_AHEAD") & "")
        dteSHIP_DATE.CalendarInfo.ActivateDay(Now.Date.AddDays(SO_PARM_RELEASE_DAYS_AHEAD))
        dteSHIP_DATE.CalendarInfo.SelectedDateRanges.Add(dteSHIP_DATE.CalendarInfo.ActiveDay.Date)
        Set_Date()
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        SHIP_BY_DATE = dteSHIP_DATE.CalendarInfo.ActiveDay.Date

        Dim sqlw1 As String = ""
        sqlw1 &= SQL_in("CUST_CODE", "SOTORDR1.CUST_CODE")
        sqlw1 &= SQL_in("WHSE_CODE", "SOTORDR1.WHSE_CODE")

        Dim sqlw2 As String = ""
        sqlw2 &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")
        sqlw2 &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sqlw2 &= SQL_in("SALES_DIVISION_CODE", "ICTITEM1.SALES_DIVISION_CODE")

        Prepare_dst(True, New Object() {SHIP_BY_DATE, sqlw1, sqlw2})
        Check_if_Empty("SOTOREL1")
    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        'Dim sqlw As String = CStr(parms(0))
        'If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst
            ASCMAIN1.sql = "Select ICTBRAN1.* from ICTBRAN1"
            Create_TDA(dst.Tables.Add, "ICTBRAN1", "**", 0, False, , 1)

            sqlSOTORDR1 = "" _
                & "Select SOTORDR1.*, ARTCUST1.CUST_PRIORITY_CODE" & vbCrLf _
                & ", 0 ORDR_AMT_OPEN, 0 ORDR_QTY_OPEN, 0 ORDR_AMT_PICK, 0 ORDR_QTY_PICK" & vbCrLf _
                & " from SOTORDR1, ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & " and SOTORDR1.ORDR_STATUS = 'O'"
            ASCMAIN1.sql = sqlSOTORDR1 & vbCrLf _
                & " and ROWNUM < 1"

            SOTORDR1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Modify ORDR_AMT_OPEN NUMBER (13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Modify ORDR_QTY_OPEN NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Modify ORDR_AMT_PICK NUMBER (13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Modify ORDR_QTY_PICK NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")

            ASCMAIN1.sql = "Create Index I_" & SOTORDR1 & "_1 on " & SOTORDR1 & " (ORDR_GROUP_NO)"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.AnalyzeTable(SOTORDR1)


            ' Set up Order Details

            sqlSOTORDR2 = "" _
                & "Select SOTORDR2.*, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
                & ", NVL(ARTCUST1.CUST_CODE_ALLO,ARTCUST1.CUST_CODE) CUST_CODE_ALLO" _
                & ", 0 PICK_QTY_CANC_REL, 0 PICK_QTY_BACK_REL" & vbCrLf _
                & " from SOTORDR2, " & SOTORDR1 & " SOTORDR1, ARTCUST1, ICTITEM1, ICTCOLL1" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE"
            ASCMAIN1.sql = sqlSOTORDR2 & vbCrLf _
                 & " and ROWNUM < 1"

            SOTORDR2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add Primary Key (ORDR_NO, ORDR_LNO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Modify PICK_QTY_CANC_REL NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Modify PICK_QTY_BACK_REL NUMBER (8,0)")

            ' Item Parameters

            sqlSOTOREL1 = "" _
                 & "Select ICTITEM1.ITEM_CODE" & vbCrLf _
                 & ", ICTITEM1.ITEM_ORDR_REL_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                 & ", ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                 & ", ICTITEM1.ITEM_NOT_ALLOCATED, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_WEIGHT" & vbCrLf _
                 & ", ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                 & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                 & " from ICTITEM1,ICTCOLL1 where ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                 & " and ICTITEM1.ITEM_CODE in (Select Distinct ITEM_CODE from " & SOTORDR2 & ")"
            ASCMAIN1.sql = sqlSOTOREL1 & vbCrLf _
                & " and ROWNUM < 1"

            SOTOREL1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL1 & " Add Primary Key (ITEM_CODE)")

            ASCMAIN1.sql = "Select * from " & SOTOREL1
            Create_TDA(dst.Tables.Add("SOTOREL1"), SOTOREL1, "**", 0)

            ' Item Status by Warehouse

            sqlSOTOREL2 = "" _
                & "Select ITEM_CODE, WHSE_CODE" & vbCrLf _
                & ", WHSE_QTY_ON_HAND QTY_ON_HAND, WHSE_QTY_OPEN QTY_OPEN, WHSE_QTY_PICK QTY_PICK" & vbCrLf _
                & ", 0 QTY_WO, 0 QTY_ATS, 0 QTY_TO_PICK, 0 QTY_TO_DE_COMMIT, 0 QTY_UNRELEASED" & vbCrLf _
                & " from ICTSTAT2 where (ITEM_CODE,WHSE_CODE)" & vbCrLf _
                & " in (Select Distinct ITEM_CODE,WHSE_CODE from " & SOTORDR2 & ")"
            ASCMAIN1.sql = sqlSOTOREL2 & vbCrLf _
                & " and ROWNUM < 1"

            SOTOREL2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add Primary Key (ITEM_CODE, WHSE_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Modify QTY_WO NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Modify QTY_ATS NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Modify QTY_TO_PICK NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Modify QTY_TO_DE_COMMIT NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Modify QTY_UNRELEASED NUMBER (8,0)")

            ASCMAIN1.sql = "Select * from " & SOTOREL2
            Create_TDA(dst.Tables.Add("SOTOREL2"), SOTOREL2, "**", 0)
        End With

        Fill_Records("ICTBRAN1")

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        SHIP_BY_DATE = CDate(parms(0))
        Dim sqlw1 As String = parms(1)
        Dim sqlw2 As String = parms(2)

        ASCDATA1.ExecuteSQL("Delete from " & SOTORDR1)
        ASCDATA1.ExecuteSQL("Delete from " & SOTORDR2)
        ASCDATA1.ExecuteSQL("Delete from " & SOTOREL1)
        ASCDATA1.ExecuteSQL("Delete from " & SOTOREL2)

        ASCMAIN1.sql = sqlSOTORDR1 & vbCrLf _
            & " and (SOTORDR1.ORDR_SHIP_DATE <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "' or SOTORDR1.ORDR_PRIORITY = '1')" & vbCrLf _
            & sqlw1
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR1 & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = sqlSOTORDR2 & vbCrLf _
            & sqlw2
        ASCDATA1.ExecuteSQL("Insert into " & SOTORDR2 & " " & ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Insert into " & SOTOREL1 & " " & sqlSOTOREL1)
        ASCDATA1.ExecuteSQL("Insert into " & SOTOREL2 & " " & sqlSOTOREL2)

        ASCMAIN1.sql = "Update " & SOTOREL2 & " SOTOREL2" & vbCrLf _
            & " Set QTY_ATS = NVL(QTY_ON_HAND,0) - NVL(QTY_PICK,0) - NVL(QTY_WO,0)"
        ASCDATA1.ExecuteSQL()

        EnforceConstraints(False)
        Fill_Records("SOTOREL1")
        Fill_Records("SOTOREL2")
        EnforceConstraints(True)
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = "Immediate Position based on Release Date of " & Format$(SHIP_BY_DATE, "MM/dd/yy")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        'If eItemKey = "Proceed" Then
        '    EMsg &= vbCr & "NO GO"
        'End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            'Case "WHSE_CODE"
            '    sqlw = " WHSE_CODE LIKE '1%'"
        End Select
        Return sqlw
    End Function

    Private Sub dteSHIP_DATE_AfterPerformAction(sender As Object, e As Infragistics.Win.UltraWinSchedule.AfterMonthViewMultiPerformActionEventArgs) Handles dteSHIP_DATE.AfterPerformAction
        Set_Date()
    End Sub

    Sub Set_Date()
        lblSHIP_DATE.Text = Format(dteSHIP_DATE.CalendarInfo.SelectedDateRanges(0).StartDate, "MM/dd/yy")
    End Sub

    Private Sub dteSHIP_DATE_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dteSHIP_DATE.MouseUp
        Set_Date()
    End Sub

End Class