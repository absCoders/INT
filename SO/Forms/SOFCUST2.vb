Imports ABSolution

Public Class SOFCUST2

    Private startDate As Date = Nothing
    Private endDate As Date = Nothing

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        With dst

            ASCMAIN1.sql = "Select CUST_CODE from ARTCUST1"
            Create_TDA(.Tables.Add, "ARTCUST1", ASCMAIN1.sql, 0, False)
            .Tables("ARTCUST1").Columns.Add("SEL")
            .Tables("ARTCUST1").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select * from ARTCUST2"
            Create_TDA(.Tables.Add, "ARTCUST2", ASCMAIN1.sql, 0, False)

            Create_Relation("ARTCUST1", "ARTCUST2", "CUST_CODE")
            .Tables("ARTCUST2").Columns.Add("SEL", GetType(System.String), "PARENT.SEL")
            .Tables("ARTCUST1").Columns.Add("STORES", GetType(System.Int64), "COUNT(CHILD(ARTCUST1_ARTCUST2).CUST_STORE_NO)")

            ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_STATUS from ICTCOLL1"
            Create_TDA(.Tables.Add, "ICTCOLLX", ASCMAIN1.sql, 0, False)
            .Tables("ICTCOLLX").Columns.Add("SEL")
            .Tables("ICTCOLLX").Columns("SEL").DefaultValue = "0"

        End With

        grdICTCOLLX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdICTCOLLX.DataSource = dst.Tables("ICTCOLLX")
        Create_Summary(grdICTCOLLX, "COLLECTION_CODE", "Count")
        Create_Summary(grdICTCOLLX, "SEL")


        grdARTCUST1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdARTCUST1.DataSource = dst.Tables("ARTCUST1")
        Create_Summary(grdARTCUST1, "CUST_CODE", "Count")
        Create_Summary(grdARTCUST1, "SEL")
        Create_Summary(grdARTCUST1, "STORES")

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")
        Create_Summary(grdARTCUST2, "CUST_STORE_NO", "Count")

        With grdARTCUST1.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("STORES").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        With grdICTCOLLX.DisplayLayout.Bands(0)
            .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("COLLECTION_STATUS").CellActivation = UltraWinGrid.Activation.NoEdit
            .Columns("SEL").CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
 

            Case "Done"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Ekta"
                Export_Store_Collection()

            Case "Done"
                Mode_Settings(False)

            Case "Refresh"
                Load_Record()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode

                .Groups("Screen Control").Visible = False
            End With


        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        '    grdARTCUST1.Visible = ScreenMode


    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        Fill_Records("ARTCUST1")
        Sort_grdColumns(grdARTCUST1, "CUST_CODE")
        Fill_Records("ARTCUST2")
        Sort_grdColumns(grdARTCUST2, "CUST_CODE,CUST_STORE_NO")
        Dim dvw As DataView = DirectCast(grdARTCUST2.DataSource, DataTable).DefaultView
        dvw.RowFilter = "SEL = '1'"

        Fill_Records("ICTCOLLX")
        For Each ROW As DataRow In dst.Tables("ICTCOLLX").Select("COLLECTION_STATUS = 'A'")
            ROW.Item("SEL") = "1"
        Next
        Sort_grdColumns(grdICTCOLLX, "COLLECTION_CODE")

        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        EnforceConstraints(False)



        EnforceConstraints(True)

        grdARTCUST1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        Sort_grdColumns(grdARTCUST1, "job_no")

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCUST1, "SSBBB", "Show Filter", "Show GroupBox", "Select All", "De-Select All", "Ekta")
        Load_Popup_Menu(grdARTCUST2, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each row As DataRow In dst.Tables("ARTCUST1").Select("") '
                    row.Item("SEL") = IIf(e.Tool.Key = "De-Select All", "0", "1")
                Next

            Case "Ekta"
                Export_Store_Collection()
        End Select
    End Sub

#End Region

    Sub Export_Store_Collection()
        Dim CUST_CODEs As String = ""
        For Each ROW As DataRow In dst.Tables("ARTCUST1").Select("SEL = '1'")
            CUST_CODEs &= "','" & ROW.Item("CUST_CODE")
        Next

        If CUST_CODEs = "" Then
            MsgBox("No Customers Selected", MsgBoxStyle.OkOnly, "Cannot Produce Ekta XLS")
            Exit Sub
        End If
        CUST_CODEs = "'" & CUST_CODEs & "'"

        Dim COLLECTION_CODEs As String = ""
        For Each ROW As DataRow In dst.Tables("ICTCOLLX").Select("SEL = '1'")
            COLLECTION_CODEs &= "','" & ROW.Item("COLLECTION_CODE")
        Next

        If COLLECTION_CODEs = "" Then
            MsgBox("No Collections Selected", MsgBoxStyle.OkOnly, "Cannot Produce Ekta XLS")
            Exit Sub
        End If
        COLLECTION_CODEs = "'" & COLLECTION_CODEs & "'"


        ASCMAIN1.Progress("Now Preparing Ekta")


        ASCMAIN1.sql = "Select ARTCUST1.CUST_NAME `Chain`, ARTCUST2.CUST_STORE_NO `Door Number`, ARTCUST2.CUST_STORE_NAME `Door Name`, " & vbCrLf _
            & "ARTCUST2.CUST_STORE_ADDR1 `Address 1`, ARTCUST2.CUST_STORE_ADDR2 `Address 2`," & vbCrLf _
            & "ARTCUST2.CUST_STORE_CITY `City`, ARTCUST2.CUST_STORE_STATE `ST`, ARTCUST2.CUST_STORE_ZIP_CODE `ZIP CD`, ARTCUST2.SELL_CODE `AE ID`, " & vbCrLf _
            & "SOTSELL1.SELL_NAME `AE Name`, SOTSELL1.SELL_EMAIL `AE Email`, SOTSELL1.REGION_CODE `ASD ID`, " & vbCrLf _
            & "SOTSELL1_ASD.SELL_NAME `ASD Name`, SOTSELL1_ASD.SELL_EMAIL `ASD Email`, " & vbCrLf _
            & "ICTCOLL1.COLLECTION_CODE `Collection Code`, ICTCOLL1.COLLECTION_NAME `Collection`, " & vbCrLf _
            & "ICTBRAN1.BRAND_NAME `Brand Name`, ICTCOLL1.BRAND_CODE `Brand Code`," & vbCrLf _
            & "SOTSELL1_ASD.SELL_PHONE `ASD Cell`, SOTSELL1.SELL_PHONE `AE Cell`" & vbCrLf _
            & "FROM ARTCUST1,ARTCUST2,ICTCOLL1,SATAUTH1,SOTSELL1,SOTSREG1,ICTBRAN1,SOTSELL1 SOTSELL1_ASD" & vbCrLf _
            & "WHERE ARTCUST2.CUST_CODE IN (" & CUST_CODEs & ")" & vbCrLf _
            & "  and ICTCOLL1.COLLECTION_CODE IN (" & COLLECTION_CODEs & ")" & vbCrLf _
            & "AND ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "AND SOTSELL1.SELL_CODE = ARTCUST2.SELL_CODE" & vbCrLf _
            & "AND SOTSREG1.REGION_CODE = SOTSELL1.REGION_CODE" & vbCrLf _
            & "AND SATAUTH1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
            & "AND SATAUTH1.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & "AND SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL AND SATAUTH1.OPS_YYYYPP_CLOSED IS NULL" & vbCrLf _
            & "AND ICTCOLL1.HC_CODE = SATAUTH1.HC_CODE" & vbCrLf _
            & "AND ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
            & "AND SOTSELL1_ASD.SELL_CODE = SOTSELL1.REGION_CODE" & vbCrLf _
            & " order by ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ICTCOLL1.COLLECTION_CODE"

        '            & "AND ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _

        ASCMAIN1.sql = Replace(ASCMAIN1.sql, "`", Chr(34))

        Dim tbl As DataTable = ASCDATA1.GetDataTable



        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim rangeCopyFrom As SpreadsheetGear.IRange
        Dim rangePaste_To As SpreadsheetGear.IRange

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

        oSheet = oWB.Worksheets(0)


        rangePaste_To = oSheet.Cells(0, 0, 2, 0)
        rangePaste_To.CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.InsertCells)

        oSheet.UsedRange.Columns.AutoFit()

        oWB.Worksheets(0).Select()

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        oWB = Nothing

        ASCMAIN1.Progress("")

    End Sub
End Class