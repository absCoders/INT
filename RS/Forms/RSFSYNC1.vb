Imports System.ComponentModel.Component
Imports System.IO
Public Class RSFSYNC1

    Dim OPS_YYYY As String
    Dim SEASON As String
    Dim RYW As String
    Dim RYP As String
    Dim RSTSYNC1 As String
    Dim RSTHIST1 As String
    Dim JHXTEMP1 As String
    Dim HUXX As String = ""
    Dim ICTITEM1 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        CalcSeasonAndYear()

        With dst

            'create temp work table with data
            ASCMAIN1.sql = "SELECT STYLE_NO_ FROM JHX.JHI_ITEM" & vbCrLf _
            & "WHERE " & vbCrLf _
            & "( CONDITION_ID = '13' OR " & vbCrLf _
            & "CONDITION_ID = '19' OR " & vbCrLf _
            & "CONDITION_ID = '27' OR " & vbCrLf _
            & "CONDITION_ID = '27' OR " & vbCrLf _
            & "CONDITION_ID = '30' OR " & vbCrLf _
            & "CONDITION_ID = '5' OR " & vbCrLf _
            & "CONDITION_ID = '6')"

            JHXTEMP1 = ASCMAIN1.Temp_Table

            Dim columnsToSee As String = "DISTINCT ICTITEM1.ITEM_CODE, ITEM_DESC, ICTITEM1.CUST_CODE, ICTITEM1.ITEM_CATGY_CODE," _
            & " COLLECTION_CODE , PROD_CODE, ICTITEM1.METAL_CLASS_CODE, ICTITEM1.MATL_CATGY_CODE, ITEM_CLASS_CODE, " _
            & "DEPT_CODE, ITEM_RETAIL_PRICE, ITEM_PRICE, " _
            & "ITEM_COST_STD, STYLE_CODE"

            'RETRIEVE ITEMS FROM JHX.JHI_ITEM
            ASCMAIN1.sql = "SELECT " & columnsToSee & " FROM ICTITEM1, " & JHXTEMP1 & vbCrLf _
            & " WHERE STYLE_CODE = STYLE_NO_ AND ITEM_CATGY_CODE = 'N'"
            RSTSYNC1 = ASCMAIN1.Temp_Table

            ASCDATA1.ExecuteSQL("Alter Table " & RSTSYNC1 & " Add Primary Key (ITEM_CODE)")

            'CREATE HISTORY TABLE
            ASCMAIN1.sql = "SELECT " & columnsToSee & ", OPS_YYYY, SEASON " & vbCrLf _
            & " FROM ICTITEM1, DPTHIST1" & vbCrLf _
            & " WHERE ICTITEM1.ITEM_CODE = DPTHIST1.ITEM_CODE AND " & vbCrLf _
            & "OPS_YYYY = '" & OPS_YYYY & "' AND SEASON = '" & Mid(cmbSeason.Value, 1, 1) & "" & "'"
            RSTHIST1 = ASCMAIN1.Temp_Table

            ASCDATA1.ExecuteSQL("Alter Table " & RSTHIST1 & " Add Primary Key (ITEM_CODE)")

            'import data from temp work table to ADO.NET table
            ASCMAIN1.sql = "SELECT * FROM " & RSTSYNC1
            Create_TDA(.Tables.Add("RSTSYNC1"), RSTSYNC1, "**", 0)

            ASCMAIN1.sql = "Select * from " & RSTHIST1
            Create_TDA(.Tables.Add("RSTHIST1"), RSTHIST1, "**", 0)

            With .Tables("RSTSYNC1").Columns
                .Add("SEL", GetType(Integer))
                .Add("ITEM_PICTURE", GetType(System.Byte()))
            End With

            With .Tables("RSTHIST1").Columns
                .Add("SEL", GetType(Integer))
                .Add("ITEM_PICTURE", GetType(System.Byte()))
            End With

            Create_TDA(.Tables.Add, "DPTHIST1", "*", , True)

            Dim YEARS As DataTable = ASCDATA1.GetDataTable("select distinct YYYY FROM (SELECT SUBSTR(YYYYPP,1,4) YYYY FROM GLTPARM3) ORDER BY YYYY DESC")
            cmbYYYY.DataSource = YEARS
            cmbYYYY.Value = Today.Year

        End With

        Dim seasons() As String = {"Spring", "Fall"}
        cmbSeason.DataSource = seasons
        cmbSeason.Value = If(SEASON = "S", "Spring", "Fall")

        '    ASCMAIN1.Add_Value_List(grdRSTSYNC1, "ITEM_CATGY_CODE")
        '    ASCMAIN1.Add_Value_List(grdRSTSYNC1, "PRICE_POINT_CODE")
        '    ASCMAIN1.Add_Value_List(grdRSTSYNC1, "MATL_CATGY_CODE", , New String() {":", "Z:Any"}, , "SELECT MATL_CATGY_CODE, MATL_CATGY_DESC FROM ICTMATLA")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If EMsg = "" Then

                End If

            Case "Export"

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)
                grdRSTSYNC1.DataSource = Nothing
                grdRSTSYNC1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
            Case ("Export")
                Stop
                Export_Cust_Codes()
                Sync_ICTITEM1()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Export").Visible = ScreenMode
                .Groups("Import Criteria").Enabled = Not ScreenMode
                cmbSeason.Enabled = Not ScreenMode
                cmbYYYY.Enabled = Not ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdRSTSLOH1.Visible = tf
        tabSales.Visible = tf
        Setup_tabSales()

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("RSTSYNC1").Rows.Clear()

        dst.EnforceConstraints = True
        grdRSTSYNC1.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()


    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Loading Data")
        Application.DoEvents()

        Call Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Dim table As String
        If cmbYYYY.Value = OPS_YYYY And cmbSeason.Value.ToString.Contains(SEASON) Then
            Fill_Records("RSTSYNC1")
            grdRSTSYNC1.DataSource = dst.Tables("RSTSYNC1")
            table = "RSTSYNC1"
        Else
            Fill_Record("RSTHIST1")
            grdRSTSYNC1.DataSource = dst.Tables("RSTHIST1")
            table = "RSTHIST1"
        End If
        EnforceConstraints(True)

        Setup_Pictures()

        ASCMAIN1.Progress("Now Setting Up Screen")

        grdRSTSYNC1.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdRSTSYNC1, "ITEM_CODE")
        grdRSTSYNC1.DisplayLayout.Bands(0).SortedColumns.Add("DEPT_CODE", True, True)
        grdRSTSYNC1.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)
        'grdRSTSLOH1.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CATGY_CODE", False, True)
        'grdRSTSLOH1.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
        grdRSTSYNC1.Rows.ExpandAll(True)


        Setup_grdRSTSYNC1(table)


        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTSYNC1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub


    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Rank Columns") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Rank Columns"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("RANK_WITHIN_COLLECTION").Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Rank Columns"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns("RANK_WITHIN_COLLECTION").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Columns("RANK").Hidden = Not tlb_sbt.Checked

            Case "Load List"
                Dim LIST_CODE As String = View_Lookup(Nothing, "LIST_CODE", "", "", "COLUMN_NAME = 'CUST_CODE'")
                If LIST_CODE <> "" Then
                    ASCMAIN1.sql = "Select CODE_VALUE from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'"
                    Dim CUST_CODEs As New List(Of String)
                    For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                        CUST_CODEs.Add(row.Item("CODE_VALUE"))
                    Next
                End If

            Case "Maintain Lists"
                ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE")
                ASCMAIN1.CodeSelector.MultipleSelections = True
                Using frmASFCODE1 As New ASFCODE1
                    frmASFCODE1.ShowDialog()
                End Using

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select


    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region


    Sub Setup_grdRSTSYNC1(ByVal tbl As String)
        grdRSTSYNC1.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()

        With grdRSTSYNC1.DisplayLayout.Bands(0)
            .Columns("SEL").Header.VisiblePosition = 0
            .Columns("ITEM_PICTURE").Header.VisiblePosition = 1
        End With

        Create_Summary(grdRSTSYNC1, "SEL", "Sum")
        Create_Summary(grdRSTSYNC1, "ITEM_CODE", "Count")

        For Each row As DataRow In dst.Tables(tbl).Rows
            row.Item("SEL") = 1
        Next

        For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdRSTSYNC1.DisplayLayout.Bands(0).Columns
            If col.Key = "SEL" Then
                col.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                col.CellActivation = UltraWinGrid.Activation.NoEdit
            End If

        Next

        'Sort_grdColumns(grdRSTSYNC1, "RANK_WITHIN_COLLECTION")
        'grdRSTSYNC1.DisplayLayout.Bands(0).SortedColumns.Add("DEPT_CODE", False, True)
        'grdRSTSYNC1.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False, False)
        'grdRSTSYNC1.Rows.ExpandAll(True)

        'grdRSTSYNC1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        'grdRSTSYNC1.DisplayLayout.Bands(0).ColumnFilters("ITEM_CATGY_CODE").FilterConditions.Add _
        '           (UltraWinGrid.FilterComparisionOperator.StartsWith, "N")
    End Sub

    Sub Setup_Pictures()
        Dim IMAGE_FOLDER As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then IMAGE_FOLDER = "C:\Documents and Settings\wjz\Desktop\Clients\JHI\Images\"
    End Sub

    Sub Setup_tabSales()
        'UltraExplorerBar1.Groups("Import Criteria").Visible = ScreenMode And (tabSales.SelectedTab.Key = "Import Exclusives")
    End Sub

    Public Overrides Function Excel_Export(ByVal grd As Infragistics.Win.UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile
        Try
            Return MyBase.Excel_Export(grd)
        Catch ex As Exception
            MsgBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Function XC( _
ByVal C As Int16, _
Optional ByVal R As Int16 = 0, _
Optional ByVal absolute As Boolean = False) As String

        Dim COL As String = ""
        If C >= 1 Then
            Dim B As Int16 = (C - 1) Mod 26 + 1
            Dim A As Int16 = (C - B) / 26
            COL = Chr(Asc("A") + B - 1)
            If A > 0 Then
                COL = Chr(Asc("A") + A - 1) & COL
            End If
            If absolute Then
                COL = "$" & COL
            End If

            If R = 0 Then
                COL = COL & ":" & COL
            ElseIf R > 0 Then
                COL = COL & IIf(absolute, "$", "") & CStr(R)
            End If
        End If

        Return COL
    End Function

    Sub Export_Cust_Codes()

        Dim insertStatement As String


        For Each row As DataRow In dst.Tables("RSTSYNC1").Rows
            With dst.Tables("DPTHIST1")
                Dim histRow As DataRow = Nothing
                histRow.Item("OPS_YYYY") = OPS_YYYY
                histRow.Item("SEASON") = SEASON
                histRow.Item("ITEM_CODE") = row.Item("ITEM_CODE")
                histRow.Item("CUST_CODE") = row.Item("CUST_CODE")
                .Rows.Add(histRow)
            End With

            'TDAs("DPTHIST1").Update(dst.Tables("DPTHIST1"))
            'STOPPING HERE...

            If row.Item("SEL") = 1 Then
                cmbYYYY.Value = OPS_YYYY
                cmbSeason.Value = SEASON

                Dim CUST_CODE As String = row.Item("ITEM_CATGY_CODE")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")

                insertStatement = "('" & OPS_YYYY & "'," & " '" & SEASON & "', '" _
                    & ITEM_CODE & "', '" _
                    & CUST_CODE & "')"

                ASCMAIN1.sql = "INSERT INTO DPTHIST1 VALUES " & vbCrLf _
                    & insertStatement
                ASCDATA1.ExecuteSQL()


            Else
                Continue For
            End If
        Next
    End Sub

    Sub Sync_ICTITEM1()
        Dim CUST_CODE As String
        Dim ITEM_CODE As String



        'SET ALL CUST CODES = NULL
        Stop
        ASCMAIN1.sql = "UPDATE ICTITEM1 SET CUST_CODE = '' WHERE CUST_CODE IS NOT NULL"
        ASCDATA1.ExecuteSQL()

        '
        For Each row As DataRow In dst.Tables("RSTSYNC1").Rows
            If row.Item("SEL") = 1 Then
                CUST_CODE = row.Item("ITEM_CATGY_CODE")
                ITEM_CODE = row.Item("ITEM_CODE")
                ASCMAIN1.sql = "UPDATE ICTITEM1 SET CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "WHERE ITEM_CODE = '" & ITEM_CODE & "'"
                ASCDATA1.ExecuteSQL()
            Else
                Continue For
            End If
        Next

    End Sub

    Sub CalcSeasonAndYear()
        OPS_YYYY = Today.Year
        SEASON = IIf(Today.Month - 1 > 6, "F", "S")

    End Sub

End Class