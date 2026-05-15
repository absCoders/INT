Public Class SARALLC1

    Dim sqlSOTORDR0 As String
    Dim SOTORDR0 As String

    Dim tblSOTORDR0 As DataTable
    Dim sqlUserFilter As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)

        Dim SHIP_DATE As Date = Now.Date.AddMonths(-1)
        SHIP_DATE = CDate(Format(SHIP_DATE, "MM/01/yyyy"))
        dteStartingShipDate.Value = SHIP_DATE
        grdSOTORDR0.Visible = False
        lblPOs.Visible = False
        txtPOs.Visible = False

        'NVL(SOTORDR0.ORDR_QTY, 0) AS ORDR_QTY
        sqlSOTORDR0 = "Select SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, TRUNC(EDT850T1.EDI_RECEIVED_DATE) EDI_RECEIVED_DATE
, SOTORDR0.CUST_DC_NO, SOTORDR0.ORDR_DEPT, SOTORDR0.SALES_DIVISION_CODE, SOTORDR0.ORDR_DATE
, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTORDR0.ORDR_ORIG_SHIP_DATE, SOTORDR0.ORDR_ORIG_CANCEL_DATE
, SOTORDR0.WHSE_CODE, SOTORDR0.ORDR_SOURCE
, SOTORDR0.ORDR_AMT, SOTORDR0.ORDR_AMT_OPEN, SOTORDR0.ORDR_AMT_PICK, SOTORDR0.ORDR_AMT_SHIP, SOTORDR0.ORDR_AMT_CANC
, SOTORDR0.ORDR_QTY, SOTORDR0.ORDR_QTY_OPEN, SOTORDR0.ORDR_QTY_PICK, SOTORDR0.ORDR_QTY_SHIP, SOTORDR0.ORDR_QTY_CANC
, SOTORDR0.ORDR_CNT, SOTORDR0.ORDR_CNT_OPEN, SOTORDR0.ORDR_CNT_PICK, SOTORDR0.ORDR_NO_MIN, SOTORDR0.ORDR_ARRIVAL_DATE, SOTORDR0.ORDR_ALLO_DATE
, DECODE(NVL(E1.EDI_STATUS, NULL), NULL, '0', '1') CUST_855, E8.SENT_855, SOTORDR0.EDI_DOC_SEQ_NO
, SOTORDR0.ORDR_INTERNAL_NOTES
, SOTORDR1.ORDR_HOLD, SOTORDR1.REVERSE_PO
    from SOTORDR0, SOTORDR1, EDT850T1,
    (SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '855') E1,
    (SELECT ORDR_GROUP_NO, MAX(INIT_DATE) SENT_855 FROM EDT855O1 GROUP BY ORDR_GROUP_NO) E8
    where SOTORDR1.ORDR_NO (+) = SOTORDR0.ORDR_NO_MIN 
    and SOTORDR0.CUST_CODE = E1.CUST_CODE(+)
    and SOTORDR0.ORDR_GROUP_NO = E8.ORDR_GROUP_NO (+)
    and SOTORDR1.EDI_DOC_SEQ_NO = EDT850T1.EDI_DOC_SEQ_NO (+)"

        SOTORDR0 = ASCMAIN1.Temp_Table(sqlSOTORDR0 & " and ROWNUM < 1")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")

    End Sub

    Sub Format_grdSOTORDR0()
        With dst

            ASCMAIN1.sql = "Select X.*, ARTCUST1.CUST_NAME, SOTORDR1.CUST_STORE_NO, NVL(SOTORDR1.CUST_STORE_LOCATION,SOTORDR5.CUST_NAME) CUST_STORE_LOCATION" & vbCrLf _
                & " from " & SOTORDR0 & " X, SOTORDR1, ARTCUST1, SOTORDR5" & vbCrLf _
                & " where SOTORDR1.ORDR_NO = X.ORDR_NO_MIN and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
                & "   and SOTORDR5.ORDR_NO (+) = X.ORDR_NO_MIN" & vbCrLf _
                & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'"

            tblSOTORDR0 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR0", 1, True)
            tblSOTORDR0.Columns.Add("SEL")
            tblSOTORDR0.Columns("SEL").DefaultValue = "0"

            ' DO NOT USE DST FOR FORMS THAT NEED TO SHOW GRIDS BOUND TO DATATABLES - UNFORTUNATELY, WE NEED TO USE FORM LEVEL DATATABLE OBJECTS
            'Create_TDA(.Tables.Add, "SOTORDR0", "**", 0, False, "", 1)
            '.Tables("SOTORDR0").Columns.Add("SEL")
            '.Tables("SOTORDR0").Columns("SEL").DefaultValue = "0"
            ' Fill_Records("SOTORDR0")
        End With

        grdSOTORDR0.DataSource = tblSOTORDR0 ' dst.Tables("SOTORDR0")
        grdSOTORDR0.DisplayLayout.GroupByBox.Hidden = False

        grdSOTORDR0.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSOTORDR0.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdSOTORDR0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        With grdSOTORDR0.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SEL" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each C As String In New String() {"SEL", "CUST_CODE", "CUST_NAME", "ORDR_CUST_PO", "ORDR_GROUP_NO"}
                .Columns(C).Header.Fixed = True
            Next
        End With

        grdSOTORDR0.DisplayLayout.Bands(0).Summaries.Clear()
        Create_Summary(grdSOTORDR0, "SEL")
        Create_Summary(grdSOTORDR0, "ORDR_GROUP_NO", "Count")

        Show_Filter(grdSOTORDR0, True)
        ASCMAIN1.Add_Value_List(grdSOTORDR0, "ORDR_SOURCE", Nothing, New String() {":", "K:Keyed", "E:EDI", "W:Web", "S:SRep"})

    End Sub


#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDR0, "BBB", "Select Selected", "Select All", "De-Select All")
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

        Select Case e.SourceControl.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdARTPYMT3"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select Selected"
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Selected.Rows
                    grow.Cells("SEL").Value = "1"
                    grow.Update()
                Next

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR0.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

                'Dim tbl As DataTable = DirectCast(grdSOTORDR0.DataSource, DataTable)
                'For Each rowSOTORDR0 As DataRow In tbl.Select("")
                '    rowSOTORDR0.Item("SEL") = IIf(e.Tool.Key = "Select All", "1", "0")
                'Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
        '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
        '    grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
        'End If

        If grd.Name = "grdARTCCTR2" Then

            Select Case e.Tool.Key

            End Select
        End If
    End Sub

#End Region


    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "G1,G2,G3,G4,G5,G6,G7,G8,G9")
        '    'Dim TYYTDGRS As Decimal = Val(rowASTSRPT1.Item("TYYTDGRS") & "")
        '    'If TYYTDBUD <> 0 Then rowASTSRPT1.Item("YTDVARNETBUDPCT") = IIf(TYYTDBUD = 0, 0, 100 * TYYTDNET / TYYTDBUD)
        'Next
    End Sub

    Public Overrides Sub Print_Report()
        ' CR_params.Add("SHOW_TOTALS", "1")
        ' Generate_Report(RPT)

        For Each rowCUST_CODE As DataRow In ASCDATA1.SelectDistinct(tblSOTORDR0.Select("SEL='1'"), New String() {"CUST_CODE"}).Select("", "CUST_CODE")
            Dim CUST_CODE As String = rowCUST_CODE.Item("CUST_CODE")
            Debug.Print(CUST_CODE)
            Load_Excel(CUST_CODE)
        Next

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblSOTORDR0 Is Nothing Then
                EMsg &= vbCr & "You must Load Customer POs and then Select at least 1"
            Else
                'Dim tbl As DataTable = DirectCast(grdSOTORDR0.DataSource, DataTable)
                'If tbl.Select("SEL='1'").Length = 0 Then
                '    EMsg &= vbCr & "You must Select at least 1 Customer PO"
                'End If

                If tblSOTORDR0.Select("SEL='1'").Length = 0 Then
                    EMsg &= vbCr & "You must Select at least 1 Customer PO"
                End If
            End If

        End If
    End Sub

    Sub Load_Excel(CUST_CODE As String)

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim ORDR_GROUP_NOs As New List(Of String)
        For Each row As DataRow In tblSOTORDR0.Select($"SEL='1' and CUST_CODE = '{CUST_CODE}'")
            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
            ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
        Next
        Dim ORDR_GROUP_NOs_sql As String = "'" & Join(ORDR_GROUP_NOs.ToArray, "','") & "'"

        ASCMAIN1.sql = $"Select SOTORDR2.ITEM_CODE, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_CUST_PO
, SOTORDR1.ORDR_DATE_SHIPPED, SOTORDR1.ORDR_STATUS
, SOTORDR1.CUST_CODE,  SOTORDR1.CUST_STORE_NO, ICTCOLL1.HC_CODE
, NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) + NVL(SOTORDR2.ORDR_QTY_SHIP,0) ORDR_QTY
 from SOTORDR2, SOTORDR1, ICTITEM1, ICTCOLL1
 where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO
and SOTORDR1.ORDR_STATUS IN ('O','P','F')
{sqlUserFilter}
and SOTORDR1.ORDR_GROUP_NO IN ({ORDR_GROUP_NOs_sql})
AND NVL(SOTORDR2.ORDR_QTY_OPEN,0) + NVL(SOTORDR2.ORDR_QTY_PICK,0) + NVL(SOTORDR2.ORDR_QTY_SHIP,0) <> 0
and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"

        Dim SATALLC1 As String = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = $"Select * from {SATALLC1}"
        Dim tbl As DataTable = ASCDATA1.GetDataTable

        ' Create Workbook

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Initializing Excel Objects", "")

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim FILENAME As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = CUST_CODE ' ASCMAIN1.DBS_SESSION_ID
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

        ' Add a worksheet for each HC in the dataset

        Dim HCi As Integer = 0

        Dim HC_CODEs As New List(Of String)
        For Each rowHC As DataRow In ASCDATA1.SelectDistinct(tbl, New String() {"HC_CODE"}).Select("", "HC_CODE")
            Dim HC_CODE As String = rowHC.Item("HC_CODE")
            HC_CODEs.Add(HC_CODE)
        Next

        For Each HC_CODE As String In HC_CODEs

            HCi += 1
            ASCMAIN1.Progress("-", HC_CODE)

            If HCi = 1 Then
                oSheet = oWB.Worksheets(0)
            Else
                oSheet = oWB.Worksheets.Add
            End If
            oSheet.Name = HC_CODE

            ' Load the DataTable into the Sheet

            Dim TR As Integer = 4
            Dim TC As Integer = 0

            Dim sqlIP As String = ""
            Dim IP As Integer = 0
            Dim sqlIPX As String = ""
            'Dim sqlIPDX As String = ""
            Dim Date_Shipped As String = "Case When ORDR_DATE_SHIPPED is Null Then Decode(ORDR_STATUS,'P','In Pick','O','','?') Else To_Char(ORDR_DATE_SHIPPED,'MM/DD') End"

            ASCMAIN1.sql = $"Select Distinct ITEM_CODE, ORDR_GROUP_NO, ORDR_CUST_PO from {SATALLC1} SATALLC1 where HC_CODE = '{HC_CODE}'"
            ASCMAIN1.sql = $"Select X.*, ICTITEM1.ITEM_DESC from ({ASCMAIN1.sql}) X, ICTITEM1 where ICTITEM1.ITEM_CODE = X.ITEM_CODE"
            Dim tblSATALLC2 As DataTable = ASCDATA1.GetDataTable(,, 2)
            For Each row As DataRow In tblSATALLC2.Select("", "ITEM_CODE, ORDR_CUST_PO")
                IP += 1
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                Dim ORDR_CUST_PO As String = row.Item("ORDR_CUST_PO")
                sqlIP &= $", Sum(Case  When ITEM_CODE = '{ITEM_CODE}' and ORDR_GROUP_NO = '{ORDR_GROUP_NO}' then  ORDR_QTY else 0 END) IP_{Format(IP, "000")}" & vbCrLf
                sqlIP &= $", Max (Case  When ITEM_CODE = '{ITEM_CODE}' and ORDR_GROUP_NO = '{ORDR_GROUP_NO}' then {Date_Shipped} else NULL END) IPD_{Format(IP, "000")}" & vbCrLf
                sqlIPX &= $", NVL(X.IP_{Format(IP, "000")},0) IP_{Format(IP, "000")}" & vbCrLf
                sqlIPX &= $", X.IPD_{Format(IP, "000")}" & vbCrLf
            Next

            ASCMAIN1.sql = $"SELECT ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME, SOTSREG1.REGION_DESC
, SOTSELL1.SELL_NAME, SOTSELL1_AC.SELL_NAME SELL_NAME_AC
{sqlIPX}
FROM ARTCUST2, SOTSELL1, SOTSREG1, SOTSELL1 SOTSELL1_AC,(
Select SATALLC1.CUST_STORE_NO
{sqlIP}
FROM 
{SATALLC1} SATALLC1
WHERE SATALLC1.HC_CODE = '{HC_CODE}'
GROUP BY SATALLC1.CUST_STORE_NO
) X
WHERE ARTCUST2.CUST_CODE  = '{CUST_CODE}'
AND ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO (+)
AND SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE
AND SOTSELL1_AC.SELL_CODE (+) = ARTCUST2.SELL_CODE_AC
AND SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE
AND ((NVL(ARTCUST2.CUST_STORE_STATUS,'?') = 'A' AND NVL(ARTCUST2.CUST_DC_IND,'0') <> '1') OR X.CUST_STORE_NO IS NOT NULL)
ORDER BY ARTCUST2.CUST_STORE_NO
"

            Dim tbl2 As DataTable = ASCDATA1.GetDataTable


            oSheet.Range(TR, 0, TR, 4).EntireColumn.NumberFormat = "@"
            For IPi As Integer = 1 To IP
                oSheet.Range(TR, 4 + (IPi - 1) * 2 + 1).EntireColumn.NumberFormat = "#,##0"
                oSheet.Range(TR, 4 + (IPi - 1) * 2 + 2).EntireColumn.NumberFormat = "@"
                oSheet.Range(TR, 4 + (IPi - 1) * 2 + 2).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
            Next
            oSheet.Range(TR, 0).EntireRow.NumberFormat = "@"

            Load_DataTable_into_SGXLS(TR + 1, TC + 1, tbl2, oSheet, Nothing, Nothing, "CUST_STORE_NO", "")

            Dim Cx As Integer = -1
            Cx += 1 : oSheet.Cells(TR, Cx).Value = "Store No" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 10
            Cx += 1 : oSheet.Cells(TR, Cx).Value = "Store Name" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 30
            Cx += 1 : oSheet.Cells(TR, Cx).Value = "ASD" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20
            Cx += 1 : oSheet.Cells(TR, Cx).Value = "AE" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20
            Cx += 1 : oSheet.Cells(TR, Cx).Value = "AC" : oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 20

            oSheet.Range(0, 0).EntireRow.NumberFormat = "@"
            oSheet.Range(0, 0).EntireRow.Font.Color = SpreadsheetGear.Colors.Red

            oSheet.Range(1, 0).EntireRow.NumberFormat = "@"
            oSheet.Range(2, 0).EntireRow.NumberFormat = "@"
            oSheet.Range(3, 0).EntireRow.NumberFormat = "@"

            Dim CIP As Integer = Cx

            oSheet.Cells(0 + 1, CIP).Value = "Customer PO"
            oSheet.Cells(0 + 2, CIP).Value = "Description"
            oSheet.Cells(0 + 3, CIP).Value = "Item Code"


            For I As Integer = 1 To IP
                Cx = CIP + (I - 1) * 2 + 1
                oSheet.Cells(TR, Cx).Value = "QTY"
                'oSheet.Cells(TR, Cx).Value = Format(I, "000")
                oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 12
                oSheet.Cells(TR, Cx).HorizontalAlignment = SpreadsheetGear.HAlign.Center

                Cx += 1
                oSheet.Cells(TR, Cx).Value = "Status"
                oSheet.Cells(TR, Cx).EntireColumn.ColumnWidth = 12
                oSheet.Cells(TR, Cx).HorizontalAlignment = SpreadsheetGear.HAlign.Center
            Next

            Dim IPx As Integer = 0
            For Each rowSATALLC2 As DataRow In tblSATALLC2.Select("", "ITEM_CODE, ORDR_CUST_PO")
                IPx += 1
                Dim ITEM_CODE As String = rowSATALLC2.Item("ITEM_CODE")
                Dim ITEM_DESC As String = rowSATALLC2.Item("ITEM_DESC") & ""
                Dim ORDR_CUST_PO As String = rowSATALLC2.Item("ORDR_CUST_PO") & ""
                ' oSheet.Cells(0, CIP + IPx).Value = "Wave X"
                oSheet.Cells(1, CIP + IPx).Value = ORDR_CUST_PO
                oSheet.Cells(2, CIP + IPx).Value = ITEM_DESC
                oSheet.Cells(2, CIP + IPx).WrapText = True
                oSheet.Cells(3, CIP + IPx).Value = ITEM_CODE

                oSheet.Cells(TR + tbl2.Rows.Count + 1, CIP + IPx).Formula = $"=SUBTOTAL(9,{Excel_Cell0(TR + 1, CIP + IPx)}:{Excel_Cell0(TR + tbl2.Rows.Count, CIP + IPx)})"

                IPx += 1
            Next
            oSheet.Cells(TR + tbl2.Rows.Count + 1, 0).Value = "Totals"

            oSheet.Range(TR, 0).EntireRow.AutoFilter()


            ' Border around Entry Area

            If IP > 0 Then

                For IPi As Integer = 1 To IP
                    With oSheet.Range(0, CIP + (IPi - 1) * 2 + 1, 3, CIP + (IPi - 1) * 2 + 2)
                        .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    End With
                    For IPr As Integer = 0 To 3
                        With oSheet.Range(IPr, CIP + (IPi - 1) * 2 + 1, IPr, CIP + (IPi - 1) * 2 + 2)
                            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                            .Merge()
                        End With
                    Next

                    'With oSheet.Range(0, CIP + (IPi - 1) * 2 + 1, 3, CIP + (IPi - 1) * 2 + 1)
                    '    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    'End With
                    'With oSheet.Range(0, CIP + (IPi - 1) * 2 + 2, 3, CIP + (IPi - 1) * 2 + 2)
                    '    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    '    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    'End With
                Next
                'With oSheet.Range(0, CIP + 1, 3, CIP + IP * 2)
                '    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                '    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                '    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                '    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                '    .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
                '    .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
                'End With
            End If

            For i As Integer = TR + 1 To TR + tbl2.Rows.Count Step 2
                oSheet.Range(i, 0, i, CIP + IP * 2).Interior.Color = SpreadsheetGear.Colors.LightBlue
            Next

            ' Freeze Panes

            oSheet.Range(TR + 1, 0).Select()
            oSheet.WindowInfo.FreezePanes = True
            oSheet.Range("A1:A1").Select()
        Next

        oWB.Worksheets(0).Select()

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub btnLoadPOs_Click(sender As Object, e As EventArgs) Handles btnLoadPOs.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Building List of Customer POs using filters provided")

        ' get Filters

        ASCMAIN1.sql = $"Select * from ASTDSQLC where FORM_NAME = '{MENU_ITEM_OBJECT}'"
        Dim tblC As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ASTDSQLC", 3)

        sqlUserFilter = ""
        Dim sqlw As String = ""
        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and CODE_VALUES <> ''")
            Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")
            Dim TABLE_NAME As String = "SOTORDR2"
            Dim rowASTDSALC As DataRow = tblC.Rows.Find(New String() {MENU_ITEM_OBJECT, "*", COLUMN_NAME})
            If rowASTDSALC IsNot Nothing Then
                TABLE_NAME = rowASTDSALC.Item("TABLE_NAME")
            End If

            sqlw &= SQL_in(COLUMN_NAME, TABLE_NAME & "." & COLUMN_NAME) & vbCrLf
            'If COLUMN_NAME <> "CUST_CODE" Then
            sqlUserFilter &= SQL_in(COLUMN_NAME, TABLE_NAME & "." & COLUMN_NAME) & vbCrLf
            'End If
        Next

        ASCMAIN1.sql = $"
SELECT DISTINCT SOTORDR1.ORDR_GROUP_NO
FROM SOTORDR1,SOTORDR2,ICTITEM1,ICTCOLL1
WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO
AND ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE
AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE
{sqlw}
AND SOTORDR1.ORDR_SHIP_DATE > '{Format(dteStartingShipDate.Value, "dd-MMM-yyyy")}'
AND SOTORDR1.ORDR_STATUS IN ('O','P','F')
"
        'Stop



        ASCDATA1.ExecuteSQL($"Truncate Table {SOTORDR0}")
        ASCMAIN1.sql = sqlSOTORDR0 & $" and SOTORDR0.ORDR_GROUP_NO in ({ASCMAIN1.sql})"
        ASCDATA1.ExecuteSQL($"Insert into {SOTORDR0} " & ASCMAIN1.sql)

        Format_grdSOTORDR0()

        Dim ORDR_GROUP_NOs_Not_Found As New List(Of String)
        Dim rowASTDSQLA_ORDR_GROUP_NO As DataRow = tblASTDSQLA.Rows.Find("ORDR_GROUP_NO")
        Dim CODE_VALUES As String = rowASTDSQLA_ORDR_GROUP_NO.Item("CODE_VALUES") & ""
        If CODE_VALUES <> "" Then
            Dim ORDR_GROUP_NOs_Saved() As String = Split(CODE_VALUES, ",")
            For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs_Saved
                Dim rowSelect() As DataRow = tblSOTORDR0.Select($"ORDR_GROUP_NO = '{ORDR_GROUP_NO}'")
                If rowSelect.Length = 0 Then
                    ORDR_GROUP_NOs_Not_Found.Add(ORDR_GROUP_NO)
                Else
                    rowSelect(0).Item("SEL") = "1"
                End If
            Next
        End If

        If ORDR_GROUP_NOs_Not_Found.Count > 0 Then
            MsgBox(Join(ORDR_GROUP_NOs_Not_Found.ToArray, ","), MsgBoxStyle.OkOnly, "The Following Order Group Nos could not be found in the List")
        End If

        grdSOTORDR0.Visible = True
        lblPOs.Visible = True
        txtPOs.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub txtPOs_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPOs.KeyPress

    End Sub

    Private Sub txtPOs_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPOs.KeyUp
        If e.KeyCode = Keys.V And e.Control Then
            Stop
            '    If grdSetup.ActiveCell IsNot Nothing AndAlso grdSetup.ActiveCell.Column.Key = "CODE_VALUES" Then
            '        Dim CODE_VALUES As String = grdSetup.ActiveCell.Value & ""
            '        If InStr(CODE_VALUES, vbCrLf) <> 0 Then
            '            CODE_VALUES = Replace(CODE_VALUES, vbCrLf, ",")
            '            If CODE_VALUES.EndsWith(",") Then
            '                CODE_VALUES = Mid(CODE_VALUES, 1, Len(CODE_VALUES) - 1)
            '            End If
            '            grdSetup.ActiveCell.Value = CODE_VALUES
            '            grdSetup.ActiveRow.Update()
            '        End If
            '    End If

        End If
    End Sub

    Private Sub txtPOs_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPOs.KeyDown
        If e.Modifiers = Keys.Control AndAlso e.KeyCode = Keys.V Then

            'Using box As New RichTextBox
            '    box.SelectAll()
            '    box.SelectedRtf = Clipboard.GetText(TextDataFormat.Rtf)
            '    box.SelectAll()
            '    box.SelectionBackColor = Color.White
            '    box.SelectionColor = Color.Black
            '    RichTextBox1.SelectedRtf = box.SelectedRtf
            'End Using
            Dim txt As String = Clipboard.GetText(TextDataFormat.Text)

            Dim EMSG As String = ""
            Dim m As Integer = 0
            Dim m2 As Integer = 0
            For Each ORDR_CUST_PO As String In Split(txt, vbCrLf)

                If ORDR_CUST_PO.Contains("'") Then
                    ORDR_CUST_PO = Replace(ORDR_CUST_PO, "'", "_")
                End If
                If ORDR_CUST_PO <> "" Then
                    Dim row() As DataRow = tblSOTORDR0.Select($"ORDR_CUST_PO = '{ORDR_CUST_PO}'")
                    If row.Length = 1 Then
                        If row(0).Item("SEL") & "" = "1" Then
                            m2 += 1
                        Else
                            row(0).Item("SEL") = "1"
                            m += 1
                        End If

                    ElseIf row.Length = 0 Then
                        EMSG &= vbCr & $"Could not find PO {ORDR_CUST_PO}"
                    Else
                        EMSG &= vbCr & $"Multiple Matches for PO {ORDR_CUST_PO}"
                    End If
                End If
            Next

            MsgBox($"{CStr(m)} Matches, {CStr(m2)} already Matched" & vbCrLf & EMSG, MsgBoxStyle.OkOnly, "Match Results")

            txtPOs.Clear()

            'e.Handled = True
        End If
    End Sub
End Class