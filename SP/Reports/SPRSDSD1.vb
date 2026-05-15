Public Class SPRSDSD1

#Region "General Declarations"
    Dim SEASON_CODE As String

    Dim SPTSDSDC As String  ' Customers
    Dim sqlSPTSDSDC As String
    Dim SPTSDSDL As String  ' Collections
    Dim sqlSPTSDSDL As String
    Dim SPTSDSDX As String  ' Expense Types
    Dim sqlSPTSDSDX As String

    Dim rowICTSEAS1 As DataRow
    Dim SEASON_YEAR As String
    Dim SEASON_YEAR_LY As String
    Dim SEASON_TYPE As String
    Dim SEASON_DESC As String
    Dim SEASON_DESC_LY As String

    Dim YPs(,) As String
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SPTPARM1")
        Absx1.txtFor("SEASON_CODE").Text = Mid(ASCMAIN1.CYM, 1, 4) & IIf(Val(Mid(ASCMAIN1.CYM, 5, 2)) < 8, "S", "F")
    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building Work File")

        Dim sqlw As String = ""
        'sqlw &= "   and SPTCOOP1.OPS_YYYYPP <= '" & RYP0 & "'"

        Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
        Prepare_dst(True, New Object() {SEASON_CODE})
        Check_if_Empty("SPTSDSD0")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Overrides Sub Print_Report()
        Create_XLS()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
            Dim rowICTSEAS1 As DataRow = LookUp("ICTSEAS1", SEASON_CODE)
            If rowICTSEAS1 Is Nothing Then
                EMsg &= vbCrLf & "Invalid Season"
            Else
                Dim SEASON_YEAR As String = rowICTSEAS1.Item("SEASON_YEAR") & ""
                If SEASON_YEAR <> Mid(ASCMAIN1.CYM, 1, 4) And _
                   SEASON_YEAR <> Format(Val(Mid(ASCMAIN1.CYM, 1, 4)) - 1, "0000") And _
                   SEASON_YEAR <> Format(Val(Mid(ASCMAIN1.CYM, 1, 4)) + 1, "0000") Then
                    EMsg &= vbCrLf & "Invalid Season Year"
                End If
                Dim SEASON_TYPE As String = rowICTSEAS1.Item("SEASON_TYPE") & ""
                If SEASON_TYPE <> "S" And SEASON_TYPE <> "F" Then
                    EMsg &= vbCrLf & "Invalid Season Type"
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SPTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables()

        With dst

            ASCMAIN1.sql = "Select * from " & SPTSDSDC & " SPTSDSDC"
            Create_TDA(.Tables.Add, "SPTSDSDC", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from " & SPTSDSDX & " SPTSDSDX"
            Create_TDA(.Tables.Add, "SPTSDSDX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM between :PARM1 and :PARM2"
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False, "VV", 1)
            With .Tables("GLTPARM3").Columns
                .Add("YYYYWW_LY")
                .Add("YYYYWW_2LY")
            End With

            Dim fl As String = ASCMAIN1.Flattened_List("AUTH_NO", "COLLECTION_CODE", "SPTCOOP3",,, "AUTH_LNO")

            ASCMAIN1.sql = "Select DISTINCT SPTCOOP1.*, X.DIST_AMT, X.FEATURE_DESC, X.COLLECTION_CODE, FL.COLLECTION_CODES" & vbCrLf _
                & " from SPTCOOP1" & vbCrLf _
                & ", (" & vbCrLf _
                & "Select SPTCOOP3.AUTH_NO, SUM (SPTCOOP3.DIST_AMT) DIST_AMT, MIN (SPTCOOP3.FEATURE_DESC) FEATURE_DESC, MIN (NVL(SPTCOOP3.COLLECTION_CODE,'_')) COLLECTION_CODE from SPTCOOP3," & SPTSDSDL & " SPTSDSDL" & vbCrLf _
                & " where NVL(SPTCOOP3.COLLECTION_CODE,'_') = SPTSDSDL.COLLECTION_CODE group by SPTCOOP3.AUTH_NO" & vbCrLf _
                & ") X, " & vbCrLf _
                & "(" & Replace(fl, "AUTH_NO, COLLECTION_CODE", "AUTH_NO, NVL(SPTCOOP3.COLLECTION_CODE,'_') COLLECTION_CODE") & ") FL" & vbCrLf _
                & " where SPTCOOP1.OPS_YYYYWW between :PARM1 and :PARM2" & vbCrLf _
                & "   and SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
                & "   and SPTCOOP1.EXPENSE_TYPE_CODE in (Select EXPENSE_TYPE_CODE from " & SPTSDSDX & ")" & vbCrLf _
                & "   and X.AUTH_NO (+) = SPTCOOP1.AUTH_NO" & vbCrLf _
                & "   and FL.AUTH_NO (+) = SPTCOOP1.AUTH_NO" & vbCrLf _
                & "   and X.COLLECTION_CODE IS NOT NULL" & vbCrLf _
                & "   and SPTCOOP1.CUST_CODE in (Select CUST_CODE from " & SPTSDSDC & ")"
            Create_TDA(.Tables.Add, "SPTCOOP1", "**", 0, False, "VV", 1)
            With .Tables("SPTCOOP1").Columns
                .Add("TOTAL", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
                .Add("AMT", GetType(System.Decimal), "DIST_AMT")
                .Add("XLS_ROW", GetType(System.Int32)) ' 0 based Row Coordinate for this event
                .Add("XLS_COL", GetType(System.Int32)) ' 0 based Col Coordinate for this event
                .Add("YYYYWW_END") ' YW that this event ends
            End With

            With dst.Tables.Add("SPTSDSD0")
                With .Columns
                    .Add("CUST_CODE")
                    .Add("OPS_YYYYWW")
                    .Add("CUST_STORE_NO")
                    .Add("RTL_TY", GetType(System.Decimal))
                    .Add("RTL_LY", GetType(System.Decimal))
                    .Add("RTL_2Y", GetType(System.Decimal))
                    .Add("EOW_TY", GetType(System.Decimal))
                    .Add("EOW_LY", GetType(System.Decimal))
                    .Add("EOW_2Y", GetType(System.Decimal))
                    .Add("COMP", GetType(System.String), "IIF((ISNULL(RTL_TY,0) <> 0 OR ISNULL(EOW_TY,0) <> 0) AND (ISNULL(RTL_LY,0) <> 0 OR ISNULL(EOW_LY,0) <> 0),'1','0')")
                End With
                .PrimaryKey = New DataColumn() {.Columns("CUST_CODE"), .Columns("OPS_YYYYWW"), .Columns("CUST_STORE_NO")}
            End With

            Dim sqlB As String = ""
            For i As Integer = 1 To 12
                sqlB &= ", Sum (Decode(OPS_YYYYPP,:PARM" & CStr(i) & ",RSTBUDR1.BUDGET,0)) P" & Format(i, "00") & vbCrLf
            Next
            ASCMAIN1.sql = "Select RSTBUDR1.CUST_CODE" & vbCrLf _
                & sqlB _
                & " from RSTBUDR1," & SPTSDSDC & " SPTSDSDC," & SPTSDSDL & " SPTSDSDL" & vbCrLf _
                & " where RSTBUDR1.CUST_CODE = SPTSDSDC.CUST_CODE" & vbCrLf _
                & "   and SPTSDSDL.COLLECTION_CODE = RSTBUDR1.COLLECTION_CODE" & vbCrLf _
                & "   and RSTBUDR1.OPS_YYYYPP between :PARM13 and :PARM14" & vbCrLf _
                & " group by RSTBUDR1.CUST_CODE"
            Create_TDA(.Tables.Add, "RSTBUDRX", "**", 0, False, "VVVVVVVVVVVVVV", 1)


            ASCMAIN1.sql = "Select RSTRETL1.CUST_CODE, RSTRETL1.OPS_YYYYWW" _
                & ", SUM (RSTRETL1.AMT_SOLD) RTL, SUM (RSTRETL1.QTY_EOW * ITEM_RETAIL_PRICE) EOW" _
                & " from RSTRETL1,ICTITEM1," & SPTSDSDC & " SPTSDSDC," & SPTSDSDL & " SPTSDSDL" & vbCrLf _
                & " where RSTRETL1.CUST_CODE = SPTSDSDC.CUST_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
                & "   and SPTSDSDL.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and RSTRETL1.OPS_YYYYWW between :PARM1 and :PARM2" & vbCrLf _
                & " group by RSTRETL1.CUST_CODE, RSTRETL1.OPS_YYYYWW"
            Create_TDA(.Tables.Add, "SPTSDSD1", "**", 0, False, "VV", 2)

            ASCMAIN1.sql = "Select RSTRETL1.CUST_CODE, RSTRETL1.OPS_YYYYWW, RSTRETL1.CUST_STORE_NO" _
                & ", SUM (RSTRETL1.AMT_SOLD) RTL, SUM (RSTRETL1.QTY_EOW * ITEM_RETAIL_PRICE) EOW" _
                & " from RSTRETL1,ICTITEM1," & SPTSDSDC & " SPTSDSDC," & SPTSDSDL & " SPTSDSDL" & vbCrLf _
                & " where RSTRETL1.CUST_CODE = SPTSDSDC.CUST_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
                & "   and SPTSDSDL.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and RSTRETL1.OPS_YYYYWW between :PARM1 and :PARM2" & vbCrLf _
                & " group by RSTRETL1.CUST_CODE, RSTRETL1.OPS_YYYYWW, RSTRETL1.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SPTSDSD2", "**", 0, False, "VV", 3)

            Create_TDA(.Tables.Add, "GLTPARM2", "*", 0, False)
            Fill_Records("GLTPARM2")
        End With

        If perform_fill Then
            Fill_Records_RPT(sqlw)
        End If

        Return clsASCBASE1
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        If parms.Length > 0 Then
            sqlw = parms(0)
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        SEASON_CODE = parms(0)

        rowICTSEAS1 = LookUp("ICTSEAS1", SEASON_CODE)
        SEASON_YEAR = rowICTSEAS1.Item("SEASON_YEAR")
        SEASON_YEAR_LY = Format(Val(SEASON_YEAR) - 1, "0000")
        SEASON_TYPE = rowICTSEAS1.Item("SEASON_TYPE")
        SEASON_DESC = rowICTSEAS1.Item("SEASON_DESC")
        SEASON_DESC_LY = Replace(SEASON_DESC, SEASON_YEAR, SEASON_YEAR_LY)

        Dim YP As String = ""
        If SEASON_TYPE = "F" Then
            YP = SEASON_YEAR & "08"
        ElseIf SEASON_TYPE = "S" Then
            YP = SEASON_YEAR & "02"
        End If

        ' Get YP information into an array for the 6 months of the chosen season, TY & LY 

        ReDim YPs(6, 1) ' 0 = TY, 1 = LY
        For I As Integer = 1 To 6
            YPs(I, 0) = ASCMAIN1.Period_Calc(YP, I - 1)
            YPs(I, 1) = ASCMAIN1.Period_Calc(YP, I - 1 - 12)
        Next

        ' Pull all of the weeks for the current season into a datatable, and then determine the week which is 52 weeks prior

        Fill_Records("GLTPARM3", New String() {YPs(1, 0), YPs(6, 0)})

        Dim Weeks_TY As New List(Of String)
        Dim Weeks_LY As New List(Of String)
        Dim Weeks_2LY As New List(Of String)
        For Each rowGLTPARM3 As DataRow In dst.Tables("GLTPARM3").Select("", "YYYYWW")
            Dim YYYYWW As String = rowGLTPARM3.Item("YYYYWW")
            Weeks_TY.Add(YYYYWW)
            Dim YYYYWW_LY As String = ASCMAIN1.Week_Calc(YYYYWW, -52)
            If ASCMAIN1.CLIENT = "INT" Then YYYYWW_LY = Format(Val(Mid(YYYYWW, 1, 4) - 1), "0000") & Mid(YYYYWW, 5, 2) ' 08/06/18 SP EMAIL SBS NOT DOING 53 WEEK LY CORRECTLY
            rowGLTPARM3.Item("YYYYWW_LY") = YYYYWW_LY
            Weeks_LY.Add(YYYYWW_LY)
            Dim YYYYWW_2LY As String = ASCMAIN1.Week_Calc(YYYYWW, -52 * 2)
            If ASCMAIN1.CLIENT = "INT" Then YYYYWW_2LY = Format(Val(Mid(YYYYWW, 1, 4) - 2), "0000") & Mid(YYYYWW, 5, 2) ' 08/06/18 SP EMAIL SBS NOT DOING 53 WEEK LY CORRECTLY
            rowGLTPARM3.Item("YYYYWW_2LY") = YYYYWW_2LY
            Weeks_2LY.Add(YYYYWW_2LY)
        Next

        ' Fill Work Table of Customers

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTSDSDC)
        ASCMAIN1.sql = sqlSPTSDSDC
        ASCMAIN1.sql = "Select * from (" & sqlSPTSDSDC & ") " & ASCMAIN1.SQL_Add_WHERE(SQL_in("CUST_CODE"))
        ASCDATA1.ExecuteSQL("Insert into " & SPTSDSDC & " " & ASCMAIN1.sql)

        ' Fill Work Table of Collections

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTSDSDL)
        Dim sqlBRAND_CODEs As String = SQL_in("BRAND_CODE")
        If sqlBRAND_CODEs = "" Then
            ASCMAIN1.sql = "Select COLLECTION_CODE from ICTCOLL1"
        Else
            ASCMAIN1.sql = sqlSPTSDSDL & sqlBRAND_CODEs
        End If
        ASCDATA1.ExecuteSQL("Insert into " & SPTSDSDL & " " & ASCMAIN1.sql)

        If sqlBRAND_CODEs = "" Then
            ASCDATA1.ExecuteSQL("Insert into " & SPTSDSDL & " Values ('_')")
        End If

        ' Fill Work Table of Expense Types

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTSDSDX)
        ASCMAIN1.sql = sqlSPTSDSDX
        ASCMAIN1.sql &= ASCMAIN1.SQL_Add_WHERE(SQL_in("EXPENSE_TYPE_CODE"))
        ASCDATA1.ExecuteSQL("Insert into " & SPTSDSDX & " " & ASCMAIN1.sql)


        EnforceConstraints(False)

        Fill_Records("SPTSDSDC")
        Fill_Records("SPTSDSDX")

        Dim SYPs(13) As String
        For i As Integer = 1 To 12
            Dim SYP As String = SEASON_YEAR & Format(i, "00") ' Format(i + IIf(SEASON_TYPE = "F", 6, 0), "00")
            SYPs(i - 1) = ASCMAIN1.Period_Calc(SYP, 1)
        Next

        SYPs(13 - 1) = SYPs(1 - 1)
        SYPs(14 - 1) = SYPs(12 - 1)

        Fill_Records("RSTBUDRX", SYPs)

        ' Pull all of the Coop Information into the workbook for events which start in any of the weeks in scope for the selected season, TY and LY
        ' Problem - what about an even which starts prior to the 1st week of TY, but ends sometime in the time period reflected by the SxS report

        Fill_Records("SPTCOOP1", New String() {Weeks_TY(0), Weeks_TY(25)}, True)
        Fill_Records("SPTCOOP1", New String() {Weeks_LY(0), Weeks_LY(25)}, False)
        Fill_Records("SPTCOOP1", New String() {Weeks_2LY(0), Weeks_2LY(25)}, False)

        ' ASCDATA1.DeleteRows("SPTCOOP1", "DIST_AMT = 0")

        For i As Integer = dst.Tables("SPTSDSDC").Rows.Count - 1 To 0 Step -1
            ' Each rowSPTSDSDC As DataRow In dst.Tables("SPTSDSDC").Select("")
            Dim rowSPTSDSDC As DataRow = dst.Tables("SPTSDSDC").Rows(i)
            Dim CUST_CODE As String = rowSPTSDSDC.Item("CUST_CODE")
            If dst.Tables("SPTCOOP1").Select("CUST_CODE = '" & CUST_CODE & "'").Length = 0 Then
                ASCDATA1.ExecuteSQL("Delete from " & SPTSDSDC & " where CUST_CODE = '" & CUST_CODE & "'")
                rowSPTSDSDC.Delete()
            End If
        Next

        ' Get Retail by Week

        Fill_Records("SPTSDSD1", New String() {Weeks_TY(0), Weeks_TY(25)}, True)
        Fill_Records("SPTSDSD1", New String() {Weeks_LY(0), Weeks_LY(25)}, False)
        Fill_Records("SPTSDSD1", New String() {Weeks_2LY(0), Weeks_2LY(25)}, False)

        Fill_Records("SPTSDSD2", New String() {Weeks_TY(0), Weeks_TY(25)}, True)
        Fill_Records("SPTSDSD2", New String() {Weeks_LY(0), Weeks_LY(25)}, False)
        Fill_Records("SPTSDSD2", New String() {Weeks_2LY(0), Weeks_2LY(25)}, False)

        dst.Tables("SPTSDSD0").Rows.Clear()
        For Each rowSPTSDSD2 As DataRow In dst.Tables("SPTSDSD2").Select("")
            Dim CUST_CODE As String = rowSPTSDSD2.Item("CUST_CODE")
            Dim OPS_YYYYWW As String = rowSPTSDSD2.Item("OPS_YYYYWW")
            Dim CUST_STORE_NO As String = rowSPTSDSD2.Item("CUST_STORE_NO")
            Dim rowSPTSDSD0 As DataRow = dst.Tables("SPTSDSD0").Rows.Find(New String() {CUST_CODE, OPS_YYYYWW, CUST_STORE_NO})
            If rowSPTSDSD0 Is Nothing Then
                rowSPTSDSD0 = dst.Tables("SPTSDSD0").Rows.Add(New String() {CUST_CODE, OPS_YYYYWW, CUST_STORE_NO})
            End If
            Dim sfx As String = ""
            If Weeks_TY.Contains(OPS_YYYYWW) Then sfx = "TY"
            If Weeks_LY.Contains(OPS_YYYYWW) Then sfx = "LY"
            'If Weeks_2LY.Contains(OPS_YYYYWW) Then sfx = "2LY"
            If sfx <> "" Then
                rowSPTSDSD0.Item("RTL_" & sfx) = Val(rowSPTSDSD0.Item("RTL_" & sfx) & "") + Val(rowSPTSDSD2.Item("RTL") & "")
                rowSPTSDSD0.Item("EOW_" & sfx) = Val(rowSPTSDSD0.Item("EOW_" & sfx) & "") + Val(rowSPTSDSD2.Item("EOW") & "")
            End If
        Next

        EnforceConstraints(True)

    End Sub

    Sub Create_Work_Tables()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        sqlSPTSDSDC = "Select Distinct CUST_CODE from SPTCOOP1 where SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')"
        ASCMAIN1.sql = "Select * from (" & sqlSPTSDSDC & ") where ROWNUM < 1"
        SPTSDSDC = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("Alter Table " & SPTSDSDC & " Add Primary Key (CUST_CODE)")

        sqlSPTSDSDL = "Select COLLECTION_CODE from ICTCOLL1 where COLLECTION_CODE in (Select Distinct COLLECTION_CODE from SPTCOOP3)"
        ASCMAIN1.sql = "Select * from (" & sqlSPTSDSDL & ") where ROWNUM < 1"
        SPTSDSDL = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("Alter Table " & SPTSDSDL & " Add Primary Key (COLLECTION_CODE)")

        sqlSPTSDSDX = "Select Distinct EXPENSE_TYPE_CODE from SPTTYPE1"
        ASCMAIN1.sql = "Select * from (" & sqlSPTSDSDX & ") where ROWNUM < 1"
        SPTSDSDX = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("Alter Table " & SPTSDSDX & " Add Primary Key (EXPENSE_TYPE_CODE)")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub
    Sub Create_XLS()
        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim XLS_FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No(Me.Name & "_SXS") & "SXS.xlsX"

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()


        ' Get List of All Event Types and Associate a background color with each

        Dim EXPENSE_TYPE_CODEs As New List(Of String)
        For Each row As DataRow In dst.Tables("SPTSDSDX").Select("", "EXPENSE_TYPE_CODE")
            Dim EXPENSE_TYPE_CODE As String = row.Item(0)
            EXPENSE_TYPE_CODEs.Add(EXPENSE_TYPE_CODE)
        Next

        Dim ET_COLORs As New List(Of SpreadsheetGear.Color)
        ET_COLORs.Add(SpreadsheetGear.Colors.Yellow)
        ET_COLORs.Add(SpreadsheetGear.Colors.Cyan)
        ET_COLORs.Add(SpreadsheetGear.Colors.LightGray)
        ET_COLORs.Add(SpreadsheetGear.Colors.Orange)
        ET_COLORs.Add(SpreadsheetGear.Colors.Pink)
        ET_COLORs.Add(SpreadsheetGear.Colors.LightBlue)
        ET_COLORs.Add(SpreadsheetGear.Colors.Fuchsia)
        ET_COLORs.Add(SpreadsheetGear.Colors.Gold)
        ET_COLORs.Add(SpreadsheetGear.Colors.Lime)


        ' Create a worksheet for each Customer

        Dim iCUST_CODE As Integer = 0

        For Each rowSPTSDSDC As DataRow In dst.Tables("SPTSDSDC").Select("", "CUST_CODE")
            Dim CUST_CODE As String = rowSPTSDSDC.Item("CUST_CODE")
            iCUST_CODE += 1
            If iCUST_CODE = 1 Then
                worksheet = workbook.Worksheets(0)
            Else
                worksheet = workbook.Worksheets.Add
            End If
            worksheet.Name = CUST_CODE

            Dim rowRSTBUDRC As DataRow = dst.Tables("RSTBUDRX").Rows.Find(CUST_CODE)

            ' Set up Column Counts for Zones

            Dim WK_cols As Integer = 2 ' number of columns reserved to indicate Week Identfier, such as Week Number, and Dates
            Dim CE_cols As Integer = 20 ' number of Columns reserved for Events
            Dim TE_cols As Integer = 8 ' number of columns reserved for Cost and Comparative Analysis figures
            Dim LYR_cols As Integer = 6 ' number of columns reseved for Retail Sales and Comparative % for LY
            Dim LYTE_cols As Integer = 1 ' number of columns reserved for Cost information for LY Events
            Dim Total_Cols As Integer = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols + CE_cols + LYTE_cols

            Dim C_Offset As Integer = 0

            C_Offset = 1 + WK_cols + CE_cols
            worksheet.Cells(0, C_Offset + 0).EntireColumn.NumberFormat = "#,##0" ' Cost
            worksheet.Cells(0, C_Offset + 1).EntireColumn.NumberFormat = "#,##0" ' Plan
            worksheet.Cells(0, C_Offset + 2).EntireColumn.NumberFormat = "#,##0" ' TY Rtl
            worksheet.Cells(0, C_Offset + 3).EntireColumn.NumberFormat = "0.0%" ' %Chg/LY
            worksheet.Cells(0, C_Offset + 4).EntireColumn.NumberFormat = "#,##0" ' Comp
            worksheet.Cells(0, C_Offset + 5).EntireColumn.NumberFormat = "0.0%" ' %Chg/LY
            worksheet.Cells(0, C_Offset + 6).EntireColumn.NumberFormat = "#,##0" ' $OH
            worksheet.Cells(0, C_Offset + 7).EntireColumn.NumberFormat = "0.0%" ' OH%Rtl

            C_Offset = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols
            worksheet.Cells(0, C_Offset + 0).EntireColumn.NumberFormat = "#,##0" ' LY Rtl
            worksheet.Cells(0, C_Offset + 1).EntireColumn.NumberFormat = "0.0%" ' %Chg/2LY
            worksheet.Cells(0, C_Offset + 2).EntireColumn.NumberFormat = "#,##0" ' Comp
            worksheet.Cells(0, C_Offset + 3).EntireColumn.NumberFormat = "#,##0" ' Strs
            worksheet.Cells(0, C_Offset + 4).EntireColumn.NumberFormat = "#,##0" ' $OH
            worksheet.Cells(0, C_Offset + 5).EntireColumn.NumberFormat = "0.0%" ' OH%Rtl

            C_Offset = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols + CE_cols
            worksheet.Cells(0, C_Offset + 0).EntireColumn.NumberFormat = "#,##0" ' Cost


            Dim Event_Columns(,) As DataRow
            ReDim Event_Columns(CE_cols, 1)

            Dim rangeCopyFrom As SpreadsheetGear.IRange
            Dim rangePaste_To As SpreadsheetGear.IRange

            Dim Head_Lines As New List(Of Integer)
            Dim Total_Lines As New List(Of Integer)
            Dim RX As Integer = 0 ' Starting row 0-based for 1st week of 1st month

            For iYP As Integer = 1 To 6

                If iYP = 1 Or iYP = 4 Then
                    If iYP = 4 Then
                        worksheet.Cells(RX, 0).PageBreak = SpreadsheetGear.PageBreak.Manual
                    End If

                    If iYP = 1 Then
                        Head_Lines.Add(RX)

                        ' Paint Event Type Legends 

                        Dim ET_ctr As Integer = 0
                        For Each EXPENSE_TYPE_CODE As String In EXPENSE_TYPE_CODEs
                            ET_ctr += 1
                            For LYA As Integer = 0 To 1
                                Dim Col_to_Start As Integer = 3 + (CE_cols + TE_cols + 1 + WK_cols + LYR_cols) * LYA

                                If ET_ctr = 1 Then
                                    With worksheet.Cells(RX + 2, Col_to_Start, RX + 2, Col_to_Start + CE_cols - 1)
                                        .Font.Size = 8
                                        .Font.Bold = True
                                        .Orientation = 90
                                        .EntireColumn.ColumnWidth = 2
                                    End With
                                End If

                                With worksheet.Cells(RX + 2, Col_to_Start + ET_ctr - 1)
                                    .Value = EXPENSE_TYPE_CODE
                                    .Interior.Color = ET_COLORs((ET_ctr - 1) Mod ET_COLORs.Count)
                                End With
                            Next
                        Next

                        RX += 3

                        If iYP = 1 Then
                            worksheet.Range(RX, 0).Select()
                            worksheet.WindowInfo.FreezePanes = True
                        End If
                    End If
                End If

                Dim TYP As String = YPs(iYP, 0)

                Dim RX_M As Integer = RX ' remember starting row for Month

                Dim PLAN_M As Decimal = 0
                If rowRSTBUDRC IsNot Nothing Then PLAN_M = Val(rowRSTBUDRC.Item("P" & Format(iYP + IIf(SEASON_TYPE = "F", 6, 0), "00")) & "")
                Dim WK_PCT(6) As Decimal
                Dim WKS As Integer = dst.Tables("GLTPARM3").Select("YYYYMM = '" & TYP & "'").Length
                For i As Integer = 1 To WKS
                    WK_PCT(i) = 1 / WKS
                Next

                ' Get all weeks in the month

                Dim iWK As Integer = 0
                For Each rowYW As DataRow In dst.Tables("GLTPARM3").Select("YYYYMM = '" & TYP & "'", "YYYYWW")
                    Dim TYW As String = rowYW.Item("YYYYWW")
                    Dim LYW As String = rowYW.Item("YYYYWW_LY")
                    iWK += 1

                    Dim RX_W As Integer = RX

                    Dim ATY As Integer = 1 + dst.Tables("SPTCOOP1").Select("CUST_CODE = '" & CUST_CODE & "' and OPS_YYYYWW = '" & TYW & "'").Length ' How many Events TY in the current week
                    Dim ALY As Integer = 1 + dst.Tables("SPTCOOP1").Select("CUST_CODE = '" & CUST_CODE & "' and OPS_YYYYWW = '" & LYW & "'").Length ' How many Events LY in the current week

                    Dim RMAX As Integer = 3 ' no less than 3 rows dedicated to a week; note 1st row is always blank in a week
                    If ATY > RMAX Then RMAX = ATY
                    If ALY > RMAX Then RMAX = ALY

                    For LYA As Integer = 0 To 1
                        Dim CxLY As Integer = (WK_cols + CE_cols + TE_cols + 1) * LYA ' Column offset for LY

                        worksheet.Cells(RX_W, CxLY + 1).Value = "'" & Mid(IIf(LYA = 0, TYW, LYW), 5, 2)
                        With worksheet.Cells(RX_W, CxLY + 1, RX_W + RMAX - 1, CxLY + 1)
                            .Merge()
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                            .VerticalAlignment = SpreadsheetGear.VAlign.Center
                        End With

                        With worksheet.Cells(RX_W, CxLY + 2, RX_W + RMAX - 1, CxLY + 2)
                            .Merge()
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                            .VerticalAlignment = SpreadsheetGear.VAlign.Center
                            Dim DT As Date = rowYW.Item("WEEK_END_DATE")
                            If LYA = 1 Then DT = DT.AddDays(-52 * 7)

                            Dim LY_DATES As String = Format(DT.AddDays(-6), "MM/dd") & vbCrLf & Format(DT, "MM/dd")

                            If ASCMAIN1.CLIENT = "INT" And LYA = 1 Then
                                Dim rowYW_LY As DataRow = LookUp("GLTPARM3", LYW)
                                If rowYW_LY IsNot Nothing Then
                                    DT = rowYW_LY.Item("WEEK_END_DATE")
                                    LY_DATES = Format(DT.AddDays(-6), "MM/dd") & vbCrLf & Format(DT, "MM/dd")
                                Else
                                    LY_DATES = "N/A"
                                End If
                            End If

                            .Value = LY_DATES
                        End With
                    Next

                    ' Show Co-Op Agreements

                    For LYA As Integer = 0 To 1
                        Dim CxLY As Integer = (CE_cols + TE_cols + 1 + WK_cols + LYR_cols) * LYA ' Column offset for LY

                        Dim Ri As Integer = 0
                        For Each row As DataRow In dst.Tables("SPTCOOP1").Select("CUST_CODE = '" & CUST_CODE & "' and OPS_YYYYWW = '" & IIf(LYA = 0, TYW, LYW) & "'", "DATE_START")
                            Ri += 1

                            Dim CX_E As Integer = 0
                            For I As Integer = 0 To CE_cols - 1
                                If Event_Columns(I, LYA) Is Nothing Then
                                    Event_Columns(I, LYA) = row
                                    CX_E = I
                                    row.Item("XLS_ROW") = RX_W + Ri
                                    row.Item("XLS_COL") = CX_E
                                    Exit For
                                End If
                            Next

                            'COMBINE COLLECTION, BOOKNAME, FEATURE, CIRC
                            With worksheet.Cells(RX_W + Ri, 1 + WK_cols + CxLY + CX_E)
                                Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE") & ""
                                Dim COLLECTION_CODES As String = row.Item("COLLECTION_CODES") & ""
                                Dim BOOKING_NAME As String = row.Item("BOOKING_NAME") & ""
                                Dim FEATURE_DESC As String = row.Item("FEATURE_DESC") & ""
                                Dim QTY As Int64 = Val(row.Item("QTY") & "")
                                Dim SXS_DESC As String = COLLECTION_CODES & "-" & BOOKING_NAME
                                If ASCMAIN1.CLIENT = "AHA" Then
                                    SXS_DESC = BOOKING_NAME
                                End If

                                .Value = SXS_DESC _
                                    & IIf(FEATURE_DESC = "", "", "-" & FEATURE_DESC) _
                                    & IIf(QTY = 0, "", " (" & Format(QTY, "#,##0") & ")")
                                .Font.Size = 9
                                .VerticalAlignment = SpreadsheetGear.VAlign.Center
                            End With

                            Dim AMT As Decimal = Val(row.Item("AMT") & "")

                            worksheet.Cells(RX_W + Ri, 1 + WK_cols + CxLY + CE_cols).Value = AMT

                            Dim DATE_END As Date = row.Item("DATE_START")
                            If row.Item("DATE_END") & "" <> "" Then
                                DATE_END = row.Item("DATE_END")
                            End If
                            ASCMAIN1.sql = "Select MIN(YYYYWW) YYYYWW from GLTPARM3 where WEEK_END_DATE >= '" & Format(DATE_END, "dd-MMM-yyyy") & "'"
                            Dim YYYYWW As String = ASCDATA1.GetDataValue
                            row.Item("YYYYWW_END") = YYYYWW

                        Next
                    Next


                    ' Center all Retail Sales and Calculation cells

                    C_Offset = 1 + WK_cols + CE_cols

                    For i As Integer = 1 To TE_cols - 1
                        With worksheet.Cells(RX, C_Offset + i, RX + RMAX - 1, C_Offset + i)
                            .Merge()
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                            .VerticalAlignment = SpreadsheetGear.VAlign.Center
                        End With
                    Next

                    ' Retail Sales Budget

                    Dim PLAN_W As Decimal = PLAN_M * WK_PCT(iWK)
                    worksheet.Cells(RX, C_Offset + 1).Value = PLAN_W

                    ' TY Retail Sales and Calculations

                    Dim TYW_RTL As Decimal = 0
                    Dim TYW_EOW As Decimal = 0
                    Dim rowTYW As DataRow = dst.Tables("SPTSDSD1").Rows.Find(New String() {CUST_CODE, TYW})
                    If rowTYW IsNot Nothing Then
                        TYW_RTL = Val(rowTYW.Item("RTL") & "")
                        worksheet.Cells(RX, C_Offset + 2).Value = TYW_RTL
                        TYW_EOW = Val(rowTYW.Item("EOW") & "")
                        worksheet.Cells(RX, C_Offset + 6).Value = TYW_EOW
                    End If
                    worksheet.Cells(RX, C_Offset + 3).Formula = String.Format("=IF({2}=0,0,IF({0}=0,({1}/{2})-1,({0}/{2})-1))", _
                                                          Excel_Cell0(RX, C_Offset + 2), _
                                                          Excel_Cell0(RX, C_Offset + 1), _
                                                          Excel_Cell0(RX, C_Offset + TE_cols + 1 + WK_cols + 0))

                    ' TY Comp

                    Dim TYW_RTL_COMP As Decimal = Val(dst.Tables("SPTSDSD0").Compute("SUM(RTL_TY)", "CUST_CODE = '" & CUST_CODE & "' and COMP='1'") & "")
                    If TYW_RTL_COMP <> 0 Then
                        worksheet.Cells(RX, C_Offset + 4).Value = TYW_RTL_COMP
                        'worksheet.Cells(RX, C_Offset + 5).Formula = String.Format("=IF({1}=0,0,,({0}/{1})-1)", _
                        '                                                          Excel_Cell0(RX, C_Offset + 2), _
                        '                                                          Excel_Cell0(RX, C_Offset + TE_cols + 1 + WK_cols + 0))
                    End If

                    ' LY Retail Sales and Calculations

                    C_Offset += TE_cols + 1 + WK_cols

                    Dim LYW_RTL As Decimal = 0
                    Dim LYW_EOW As Decimal = 0
                    Dim rowLYW As DataRow = dst.Tables("SPTSDSD1").Rows.Find(New String() {CUST_CODE, LYW})
                    If rowLYW IsNot Nothing Then
                        LYW_RTL = Val(rowLYW.Item("RTL") & "")
                        worksheet.Cells(RX, C_Offset + 0).Value = LYW_RTL
                        LYW_EOW = Val(rowLYW.Item("EOW") & "")
                        worksheet.Cells(RX, C_Offset + 4).Value = LYW_EOW
                    End If
                    'worksheet.Cells(RX, C_Offset + 1).Formula = String.Format("=IF({1}=0,0,({0}/{1})-1)", _
                    '                                                          Excel_Cell0(RX, C_Offset + 0), _
                    '                                                          Excel_Cell0(RX, C_Offset + TE_cols + 1 + WK_cols + 0))

                    Dim LYW_RTL_COMP As Decimal = Val(dst.Tables("SPTSDSD0").Compute("SUM(RTL_LY)", "CUST_CODE = '" & CUST_CODE & "' and COMP='1'") & "")
                    Dim COMP_STORES As Decimal = Val(dst.Tables("SPTSDSD0").Compute("COUNT(CUST_STORE_NO)", "CUST_CODE = '" & CUST_CODE & "' and COMP='1'") & "")
                    If LYW_RTL_COMP <> 0 Then
                        worksheet.Cells(RX, C_Offset + 2).Value = LYW_RTL_COMP
                        worksheet.Cells(RX, C_Offset + 3).Value = COMP_STORES
                    End If


                    C_Offset = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols
                    For i As Integer = 0 To LYR_cols - 1
                        With worksheet.Cells(RX, C_Offset + i, RX + RMAX - 1, C_Offset + i)
                            .Merge()
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                            .VerticalAlignment = SpreadsheetGear.VAlign.Center
                        End With
                    Next

                    C_Offset = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols + CE_cols
                    For i As Integer = 1 To LYTE_cols - 1
                        With worksheet.Cells(RX, C_Offset + i, RX + RMAX - 1, C_Offset + i)
                            .Merge()
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                            .VerticalAlignment = SpreadsheetGear.VAlign.Center
                        End With
                    Next


                    RX += RMAX

                    ' Terminate Co-Op Agreements in vertical timeline

                    For LYA As Integer = 0 To 1
                        Dim CxLY As Integer = (CE_cols + TE_cols + 1 + WK_cols + LYR_cols) * LYA ' Column offset for LY

                        For I As Integer = 0 To CE_cols - 1
                            If Event_Columns(I, LYA) IsNot Nothing Then
                                Dim row As DataRow = Event_Columns(I, LYA)
                                'If ASCMAIN1.Running_in_VS And row.Item(12) = "Vday Towers" Then
                                '    Stop
                                'End If

                                Dim XLS_ROW As Integer = Val(row.Item("XLS_ROW"))
                                Dim XLS_COL As Integer = Val(row.Item("XLS_COL"))
                                With worksheet.Cells(XLS_ROW, _
                                                     1 + WK_cols + CxLY + XLS_COL, _
                                                     RX - 1, _
                                                     1 + WK_cols + CxLY + XLS_COL)
                                    Dim EXPENSE_TYPE_CODE As String = row.Item("EXPENSE_TYPE_CODE")
                                    Dim ETI As Integer = 0
                                    For iEXPENSE_TYPE_CODE As Integer = 0 To EXPENSE_TYPE_CODEs.Count - 1
                                        If EXPENSE_TYPE_CODEs(iEXPENSE_TYPE_CODE) = EXPENSE_TYPE_CODE Then
                                            ETI = iEXPENSE_TYPE_CODE
                                            Exit For
                                        End If
                                    Next
                                    .Interior.Color = ET_COLORs(ETI Mod 9)
                                End With
                                If row.Item("YYYYWW_END") <= IIf(LYA = 0, TYW, LYW) Then
                                    Event_Columns(I, LYA) = Nothing
                                End If
                            End If
                        Next
                    Next

                    With worksheet.Cells(RX_W, 1, RX - 1, Total_Cols - 1)
                        .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    End With
                Next

                Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", TYP)
                Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")
                Dim MMM As String = Mid(LEGEND, 10, 6)
                worksheet.Cells(RX_M, 0).Value = "'" & MMM
                With worksheet.Cells(RX_M, 0, RX - 1, 0)
                    .Merge()
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    .VerticalAlignment = SpreadsheetGear.VAlign.Center
                    .Orientation = 90
                End With

                ' Month Totals

                Total_Lines.Add(RX)

                C_Offset = 1 + WK_cols + CE_cols

                worksheet.Cells(RX, C_Offset + 0).Formula = String.Format("=Sum({0}:{1})", _
                                                          Excel_Cell0(RX_M, C_Offset + 0), _
                                                          Excel_Cell0(RX - 1, C_Offset + 0))

                worksheet.Cells(RX, C_Offset + 1).Formula = String.Format("=Sum({0}:{1})", _
                                                                          Excel_Cell0(RX_M, C_Offset + 1), _
                                                                          Excel_Cell0(RX - 1, C_Offset + 1))

                worksheet.Cells(RX, C_Offset + 2).Formula = String.Format("=Sum({0}:{1})", _
                                                                          Excel_Cell0(RX_M, C_Offset + 2), _
                                                                          Excel_Cell0(RX - 1, C_Offset + 2))

                rangeCopyFrom = worksheet.Range(RX, C_Offset + 1, RX, C_Offset + 1)
                rangePaste_To = worksheet.Range(RX, C_Offset + 2, RX, C_Offset + 2)
                rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

                rangeCopyFrom = worksheet.Range(RX_M, C_Offset + 3, RX_M, C_Offset + 3)
                rangePaste_To = worksheet.Range(RX, C_Offset + 3, RX, C_Offset + 3)
                'rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)


                C_Offset += TE_cols + 1 + WK_cols
                worksheet.Cells(RX, C_Offset + 0).Formula = String.Format("=Sum({0}:{1})", _
                                                                          Excel_Cell0(RX_M, C_Offset + 0), _
                                                                          Excel_Cell0(RX - 1, C_Offset + 0))



                C_Offset += +LYR_cols + CE_cols
                worksheet.Cells(RX, C_Offset + 0).Formula = String.Format("=Sum({0}:{1})",
                                                                          Excel_Cell0(RX_M, C_Offset + 0),
                                                                          Excel_Cell0(RX - 1, C_Offset + 0))


                With worksheet.Cells(RX, 0, RX, Total_Cols - 1)
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                    .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                    .Font.Bold = True
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With
                worksheet.Cells(RX, 0).Value = "'" & MMM
                worksheet.Cells(RX, 1).Value = "Totals"

                RX += 1
            Next



            ' Season Totals

            RX += 1
            With worksheet.Cells(RX, 0, RX, Total_Cols - 1)
                .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                .Font.Bold = True
                .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End With
            worksheet.Cells(RX, 0).Value = SEASON_DESC
            worksheet.Cells(RX, 1).Value = "Totals"


            C_Offset = 1 + WK_cols + CE_cols

            worksheet.Cells(RX, C_Offset + 0).Formula = String.Format("={0}+{1}+{2}+{3}+{4}+{5}", _
                                                      Excel_Cell0(Total_Lines(0), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(1), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(2), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(3), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(4), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(5), C_Offset + 0))

            worksheet.Cells(RX, C_Offset + 1).Formula = String.Format("={0}+{1}+{2}+{3}+{4}+{5}", _
                                                      Excel_Cell0(Total_Lines(0), C_Offset + 1), _
                                                      Excel_Cell0(Total_Lines(1), C_Offset + 1), _
                                                      Excel_Cell0(Total_Lines(2), C_Offset + 1), _
                                                      Excel_Cell0(Total_Lines(3), C_Offset + 1), _
                                                      Excel_Cell0(Total_Lines(4), C_Offset + 1), _
                                                      Excel_Cell0(Total_Lines(5), C_Offset + 1))


            worksheet.Cells(RX, C_Offset + 2).Formula = String.Format("={0}+{1}+{2}+{3}+{4}+{5}", _
                                                      Excel_Cell0(Total_Lines(0), C_Offset + 2), _
                                                      Excel_Cell0(Total_Lines(1), C_Offset + 2), _
                                                      Excel_Cell0(Total_Lines(2), C_Offset + 2), _
                                                      Excel_Cell0(Total_Lines(3), C_Offset + 2), _
                                                      Excel_Cell0(Total_Lines(4), C_Offset + 2), _
                                                      Excel_Cell0(Total_Lines(5), C_Offset + 2))

         
            'rangeCopyFrom = worksheet.Range(RX, C_Offset + 1, RX, C_Offset + 1)
            'rangePaste_To = worksheet.Range(RX, C_Offset + 2, RX, C_Offset + 2)
            'rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

            'rangeCopyFrom = worksheet.Range(RX_M, C_Offset + 3, RX_M, C_Offset + 3)
            'rangePaste_To = worksheet.Range(RX, C_Offset + 3, RX, C_Offset + 3)
            ''rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)


            C_Offset += TE_cols + 1 + WK_cols
            worksheet.Cells(RX, C_Offset + 0).Formula = String.Format("={0}+{1}+{2}+{3}+{4}+{5}", _
                                                      Excel_Cell0(Total_Lines(0), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(1), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(2), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(3), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(4), C_Offset + 0), _
                                                      Excel_Cell0(Total_Lines(5), C_Offset + 0))

            C_Offset += +LYR_cols + CE_cols
            worksheet.Cells(RX, C_Offset + 0).Formula = String.Format("={0}+{1}+{2}+{3}+{4}+{5}",
                                                      Excel_Cell0(Total_Lines(0), C_Offset + 0),
                                                      Excel_Cell0(Total_Lines(1), C_Offset + 0),
                                                      Excel_Cell0(Total_Lines(2), C_Offset + 0),
                                                      Excel_Cell0(Total_Lines(3), C_Offset + 0),
                                                      Excel_Cell0(Total_Lines(4), C_Offset + 0),
                                                      Excel_Cell0(Total_Lines(5), C_Offset + 0))


            ' Top of Page

            worksheet.Cells(0, 0).Value = "'" & Format(Now, "MM/dd/yy")
            worksheet.Cells(0, 1).Value = ASCMAIN1.USER_ID


            ' Headings

            For Each RxH As Integer In Head_Lines
                With worksheet.Cells(RxH, 1 + WK_cols)
                    .Value = "'" & SEASON_DESC
                    .Font.Size = 16
                    '.Font.Color = SpreadsheetGear.Colors.Purple
                    '.Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                End With

                With worksheet.Cells(RxH, 1 + WK_cols, RxH, 1 + WK_cols + CE_cols - 1)
                    .Merge()
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With


                With worksheet.Cells(RxH, 1 + WK_cols + CE_cols)
                    .Value = "'" & CUST_CODE
                    .Font.Size = 20
                    .Font.Bold = True
                    .Font.Color = SpreadsheetGear.Colors.DarkBlue
                    .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                End With

                With worksheet.Cells(RxH, 1 + WK_cols + CE_cols, RxH, 1 + WK_cols + CE_cols + TE_cols - 1)
                    .Merge()
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With




                Dim BRANDS As String = "All Brands"
                Dim BRAND_CODEs As String = SQLA("BRAND_CODE")
                If BRAND_CODEs <> "" Then BRANDS = "Brands: " & BRAND_CODEs
                With worksheet.Cells(RxH, 1 + WK_cols + CE_cols + TE_cols)
                    .Value = "'" & BRANDS
                    .Font.Size = 16
                    '.Font.Color = SpreadsheetGear.Colors.Purple
                    '.Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                End With

                With worksheet.Cells(RxH, 1 + WK_cols + CE_cols + TE_cols, RxH, 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols - 1)
                    .Merge()
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With



                With worksheet.Cells(RxH, 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols)
                    .Value = "'" & SEASON_DESC_LY
                    .Font.Size = 16
                    '.Font.Color = SpreadsheetGear.Colors.Purple
                    '.Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                End With

                With worksheet.Cells(RxH, 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols, RxH, 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols + CE_cols - 1)
                    .Merge()
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With

                With worksheet.Cells(RxH, 0, RxH, Total_Cols - 1)
                    .Interior.Color = SpreadsheetGear.Colors.LightBlue
                End With

                worksheet.Cells(RxH + 2, 0).Value = "Month"
                worksheet.Cells(RxH + 2, 1).Value = "Week"
                worksheet.Cells(RxH + 2, 2).Value = "Dates"

                'Dim Cx As Integer = 0

                C_Offset = 1 + WK_cols + CE_cols
                worksheet.Cells(RxH + 2, C_Offset + 0).Value = "Cost"
                worksheet.Cells(RxH + 2, C_Offset + 0).EntireColumn.ColumnWidth = 10
                worksheet.Cells(RxH + 2, C_Offset + 1).Value = "Plan"
                worksheet.Cells(RxH + 2, C_Offset + 1).EntireColumn.ColumnWidth = 10
                worksheet.Cells(RxH + 2, C_Offset + 2).Value = "TY Rtl"
                worksheet.Cells(RxH + 2, C_Offset + 2).EntireColumn.ColumnWidth = 10
                worksheet.Cells(RxH + 2, C_Offset + 3).Value = "%TY/LY"
                worksheet.Cells(RxH + 2, C_Offset + 4).Value = "Comp"
                worksheet.Cells(RxH + 2, C_Offset + 5).Value = "%TY/LY"
                worksheet.Cells(RxH + 2, C_Offset + 6).Value = "$OH"
                worksheet.Cells(RxH + 2, C_Offset + 7).Value = "OH%Rtl"

                ' worksheet.Cells(2, Cx).EntireColumn.ColumnWidth = 2
                C_Offset = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols
                worksheet.Cells(RxH + 2, C_Offset - 2).Value = "LY Wk#"
                worksheet.Cells(RxH + 2, C_Offset - 1).Value = "Dates"
                worksheet.Cells(RxH + 2, C_Offset + 0).Value = "LY Rtl"
                worksheet.Cells(RxH + 2, C_Offset + 0).EntireColumn.ColumnWidth = 10
                worksheet.Cells(RxH + 2, C_Offset + 1).Value = "%Chg/2LY"
                worksheet.Cells(RxH + 2, C_Offset + 2).Value = "Comp"
                worksheet.Cells(RxH + 2, C_Offset + 3).Value = "Strs"
                worksheet.Cells(RxH + 2, C_Offset + 4).Value = "$OH"
                worksheet.Cells(RxH + 2, C_Offset + 5).Value = "OH%Rtl"

                C_Offset = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols + LYR_cols + CE_cols
                worksheet.Cells(RxH + 2, C_Offset + 0).Value = "Cost"

                With worksheet.Cells(RxH + 2, 0, RxH + 2, Total_Cols - 1)
                    .Font.Color = SpreadsheetGear.Colors.Blue
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With
                With worksheet.Cells(RxH + 2, 0, RxH + 2, Total_Cols - 1)
                    .Borders.LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders.Color = SpreadsheetGear.Colors.LightSlateGray
                    '.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With
            Next

            With worksheet.PageSetup
                .Orientation = SpreadsheetGear.PageOrientation.Landscape
                .PrintArea = Excel_Cell0(0, 0) & ":" + Excel_Cell0(RX, Total_Cols - 1)
                '.PrintHeadings = True - this prints A B C across the top, and line numbers down the page
                .PrintTitleRows = Excel_Cell0(0, 0) & ":" & Excel_Cell0(0 + 2, Total_Cols - 1 + 1)
                .FitToPagesWide = 1
                .FitToPagesTall = 0
            End With



            'Delete Comp, %TY/LY, $OH snf OH%Rtl - as per SP - not actionable data and too much info

            C_Offset = 1 + WK_cols + CE_cols + TE_cols + 1 + WK_cols
            worksheet.Range(0, C_Offset + 1, 0, C_Offset + 5).EntireColumn.Hidden = True

            C_Offset = 1 + WK_cols + CE_cols
            worksheet.Range(0, C_Offset + 4, 0, C_Offset + 7).EntireColumn.Hidden = True

        Next

        workbook.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(XLS_FILENAME)
    End Sub

 End Class