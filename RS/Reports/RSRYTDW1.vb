Public Class RSRYTDW1

    Dim SEASON_CODE As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR As String
    Dim SEASON_YEAR_LY As String

    Dim RSTSSTW1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW0", ASCMAIN1.CYW, -5 * 52, 1 * 52, -13) '  -13 - 1 * 52)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -5 * 52, 1 * 52, 0) ' 0 - 1 * 52)
    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        Create_Pivot()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            Dim C As Integer = 1 + ASCMAIN1.Week_Diff(Absx1.cmbFor("RYW0").Value, Absx1.cmbFor("RYW1").Value)
            If C < 1 Or C > 52 Then
                EMsg &= vbCr & "Week range must be between 1 and 52 weeks"
            End If
        End If
    End Sub

    Sub Create_Pivot()

        Dim SQLW As String = ""
        SQLW &= SQLA_filter("BRAND_CODE", "ICTCOLL1")
        SQLW &= SQLA_filter("COLLECTION_CODE", "ICTITEM1")
        SQLW &= SQLA_filter("COST_CATGY_CODE", "ICTITEM1")
        SQLW &= SQLA_filter("CUST_CODE", "RSTRETL1")
        SQLW &= SQLA_filter("HC_CODE", "ICTCOLL1")
        SQLW &= SQLA_filter("ITEM_BASIC_PROMO", "ICTITEM1")
        SQLW &= SQLA_filter("ITEM_CODE", "RSTRETL1")
        SQLW &= SQLA_filter("PROD_CODE", "ICTITEM1")
        SQLW &= SQLA_filter("TRADE_CLASS_CODE", "ARTCUST1")

        ASCMAIN1.Progress("Now Creating Workbook")

        'Dim FILENAME As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsm"
        Dim DataTable As DataTable
        Dim r As Integer = 0
 
        ASCMAIN1.sql = "Select * from GLTPARM3" & vbCrLf _
            & " where YYYYWW >= '" & RYW0 & "'" & vbCrLf _
            & "   and YYYYWW <= '" & RYW1 & "'"

        Dim sql_SI As String = ""
        Dim sql_ST As String = ""
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
            Dim YW As String = row.Item("YYYYWW")
            Dim WEEK_END_DATE As Date = row.Item("WEEK_END_DATE")
            Dim YMD As String = Format(WEEK_END_DATE, "yyMMdd")
            Dim C As String = "YYMMDDSIQTY DATAFLD" '140105STQTY DATAFLD '140105STUSD DATAFLD
            sql_SI &= ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & YW & "',SOTINVH2.ORDR_QTY_SHIP,0)) SIQTY_" & YMD
            sql_SI &= ", 0 STQTY_" & YMD
            sql_SI &= ", 0 STUSD_" & YMD
            sql_ST &= ", 0 SIQTY_" & YMD
            sql_ST &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & YW & "',RSTRETL1.QTY_SOLD,0)) STQTY_" & YMD
            sql_ST &= ", Sum (Decode(RSTRETL1.OPS_YYYYWW,'" & YW & "',RSTRETL1.AMT_SOLD,0)) STUSD_" & YMD
        Next

        ASCMAIN1.sql = "" _
            & "Select " & vbCrLf _
            & "RSTRETL1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", RSTRETL1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE, ICTCLAS1.ITEM_CLASS_DESC" & vbCrLf _
            & ", 0 QTY_SHIP" & vbCrLf _
            & ", SUM (RSTRETL1.QTY_SOLD) QTY_SOLD" & vbCrLf _
            & ", SUM (RSTRETL1.AMT_SOLD) AMT_SOLD" & vbCrLf _
            & sql_ST _
            & " from RSTRETL1,ARTCUST1,ICTITEM1,ICTCOLL1,ICTCLAS1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = RSTRETL1.CUST_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTCLAS1.ITEM_CLASS_CODE = ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYWW BETWEEN '" & RYW0 & "' and '" & RYW1 & "'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "RSTRETL1") _
            & " group by " & vbCrLf _
            & "RSTRETL1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", RSTRETL1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE, ICTCLAS1.ITEM_CLASS_DESC" & vbCrLf _
            & " union " & vbCrLf _
            & "Select " & vbCrLf _
            & "SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE, ICTCLAS1.ITEM_CLASS_DESC" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY_SHIP" & vbCrLf _
            & ", 0 QTY_SOLD" & vbCrLf _
            & ", 0 AMT_SOLD" & vbCrLf _
            & sql_SI _
            & " from SOTINVH2,ARTCUST1,ICTITEM1,ICTCOLL1,ICTCLAS1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTCLAS1.ITEM_CLASS_CODE = ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
            & "   and SOTINVH2.OPS_YYYYWW BETWEEN '" & RYW0 & "' and '" & RYW1 & "'" & vbCrLf _
            & Replace(SQLW, "RSTRETL1", "SOTINVH2") _
            & " group by " & vbCrLf _
            & "SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE, ICTCLAS1.ITEM_CLASS_DESC" & vbCrLf

        ASCMAIN1.Progress("-", "Data")
 
        DataTable = ASCDATA1.GetDataTable



        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        'Dim rangeCopyFrom As SpreadsheetGear.IRange
        'Dim rangePaste_To As SpreadsheetGear.IRange

        Dim XL_ROWS As Integer
        Dim XL_COLS As Integer

        ' Set up parameterized row and col settings

        XL_ROWS = DataTable.Rows.Count ' # of Rows in Store List
        XL_COLS = 0   ' # of numeric columns in Layout Selected

        ' Create Workbook

        ASCMAIN1.Progress("Now Initializing Excel Objects", "")

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim FILENAME As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FILENAME = xls_path & "\" & xls_name & ".XLSx"

                If Not My.Computer.FileSystem.FileExists(FILENAME) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook()

        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i

        Dim TR As Integer = 1
        Dim TC As Integer = 1
        oSheet = oWB.Worksheets(0)
        Load_DataTable_into_SGXLS(TR, TC, DataTable, oSheet, Nothing, Nothing, "ITEM_CODE", "")

        oWB.Worksheets(0).Select()

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

        '  Add_Document_to_ASTSPRF1(ASCMAIN1.Folders("Work") & XLS_FILENAME)

    End Sub

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    '    Click_Command("View", e)
                End If
        End Select


    End Sub

    Public Overrides Sub cmb_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.cmb_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If Absx1.cmbFor("RYW0").Value & "" <> "" And Absx1.cmbFor("RYW1").Value & "" <> "" Then
            Dim C As Integer = 1 + ASCMAIN1.Week_Diff(Absx1.cmbFor("RYW0").Value, Absx1.cmbFor("RYW1").Value)
            lblWeeks.Text = CStr(C) & " Wks"
        Else
            lblWeeks.Text = ""
        End If
    End Sub
#End Region
End Class