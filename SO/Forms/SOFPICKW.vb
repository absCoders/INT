Public Class SOFPICKW

    Private wkTable As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "SELECT SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_STATUS, SOTPICK1.*, SOTPICK1.INV_NO STATUS, SOTPICK1.INV_NO WHSE_CODE" _
                & " from SOTORDR1, SOTPICK1" _
                & " WHERE ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICKO", ASCMAIN1.sql)
            wkTable = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            grdSOTPICKO.DataSource = dst.Tables("SOTPICKO")
            grdSOTPICKO_M.DataSource = dst.Tables("SOTPICKO")
            grdSOTPICKO_R.DataSource = dst.Tables("SOTPICKO")
            grdSOTPICKO_E.DataSource = dst.Tables("SOTPICKO")
            grdSOTPICKO_A.DataSource = dst.Tables("SOTPICKO")

        End With

        Create_Summary(grdSOTPICKO, "CUST_CODE", "Count")
        Create_Summary(grdSOTPICKO_M, "CUST_CODE", "Count")
        Create_Summary(grdSOTPICKO_R, "CUST_CODE", "Count")
        Create_Summary(grdSOTPICKO_E, "CUST_CODE", "Count")
        Create_Summary(grdSOTPICKO_A, "CUST_CODE", "Count")

        ASCMAIN1.Add_Value_List(grdSOTPICKO, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_M, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_R, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_E, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_A, "ORDR_STATUS")

        ASCMAIN1.Add_Value_List(grdSOTPICKO, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_M, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_R, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_E, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTPICKO_A, "PICK_STATUS")

        Show_Filter(grdSOTPICKO, True)
        Show_Filter(grdSOTPICKO_M, True)
        Show_Filter(grdSOTPICKO_R, True)
        Show_Filter(grdSOTPICKO_E, True)
        Show_Filter(grdSOTPICKO_A, True)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTPICKO, grdSOTPICKO_M, grdSOTPICKO_R, grdSOTPICKO_E, grdSOTPICKO_A}
            grd.DisplayLayout.Bands(0).Columns("ORDR_NO").CellAppearance.BackColor = Drawing.Color.LightBlue
            grd.DisplayLayout.Bands(0).Columns("ORDR_STATUS").CellAppearance.BackColor = Drawing.Color.LightBlue
            grd.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").CellAppearance.BackColor = Drawing.Color.LightBlue

            grd.DisplayLayout.Bands(0).Columns("PICK_NO").CellAppearance.BackColor = Drawing.Color.LightGreen
            grd.DisplayLayout.Bands(0).Columns("PICK_STATUS").CellAppearance.BackColor = Drawing.Color.LightGreen
            grd.DisplayLayout.Bands(0).Columns("SHIP_BOL_NO").CellAppearance.BackColor = Drawing.Color.LightGreen
        Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"


            Case "Done"

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
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabOrders.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTPICKO"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Clear_All_Filters(grdSOTPICKO)
        Clear_All_Filters(grdSOTPICKO_M)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & wkTable)

        ASCMAIN1.sql = "SELECT SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_STATUS, SOTPICK1.*, SOTPICK1.INV_NO STATUS, SOTSHIP1.WHSE_CODE" _
            & "  from SOTORDR1, SOTPICK1, SOTSHIP1" _
            & "  WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" _
            & "  AND SOTPICK1.PICK_STATUS = 'P' 
                    AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO (+)"
        ASCDATA1.ExecuteSQL("INSERT INTO " & wkTable & " " & ASCMAIN1.sql)

        ' Update missing Error Reasons.
        ASCMAIN1.Progress("-", "Error Reason")
        ASCMAIN1.sql = "UPDATE " & wkTable & " SET ERROR_REASON = (SELECT MIN(NVL(ERROR_REASON, '')) FROM SOTPICKC WHERE PICK_NO = " & wkTable & ".PICK_NO)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.Progress("-", "Order No")
        ASCMAIN1.sql = "Update " & wkTable & " SET STATUS = ORDR_NO_3PL"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.Progress("-", "SOTPICKO")
        ASCMAIN1.sql = "UPDATE " & wkTable & " SET STATUS = (SELECT ORDR_NO_3PL FROM SOTPICKO WHERE PICK_NO = " & wkTable & ".PICK_NO) WHERE STATUS IS NULL"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.Progress("-", "Printed")
        ASCMAIN1.sql = "UPDATE " & wkTable & " SET STATUS = 'R' WHERE STATUS IS NULL AND PICK_PRINTED IS NULL"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.Progress("-", "Not Printed")
        ASCMAIN1.sql = "UPDATE " & wkTable & " SET STATUS = 'E' WHERE STATUS IS NULL AND PICK_PRINTED IS NOT NULL AND ERROR_REASON IS NOT NULL"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCMAIN1.sql = "SELECT SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_STATUS" _
        '    & ", NVL(SOTPICK1.ORDR_NO_3PL, SOTPICKO.ORDR_NO_3PL) ORDR_NO_3PL, SOTPICKO.PICK_NO, SOTPICKO.ORDR_NO, SOTPICKO.PICK_STATUS, SOTPICK1.ERROR_REASON" _
        '    & "  from SOTORDR1, SOTPICKO, SOTPICK1" _
        '    & "  where SOTPICKO.ORDR_NO = SOTORDR1.ORDR_NO (+)" _
        '    & "  and SOTPICKO.PICK_NO = SOTPICK1.PICK_NO (+)" _
        '    & " union " _
        '    & " SELECT SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_STATUS" _
        '    & ", 'R' || ROWNUM ORDR_NO_3PL, SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_STATUS, SOTPICK1.ERROR_REASON " _
        '    & " from SOTORDR1, SOTPICK1" _
        '    & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO (+) AND SOTPICK1.PICK_STATUS = 'P' AND SOTPICK1.PICK_PRINTED IS NULL" _
        '    & " union " _
        '    & " SELECT SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_STATUS" _
        '    & ", 'E' || ROWNUM ORDR_NO_3PL, SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_STATUS, SOTPICK1.ERROR_REASON " _
        '    & " from SOTORDR1, SOTPICK1" _
        '    & " where SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO (+) AND SOTPICK1.PICK_STATUS = 'P' AND SOTPICK1.PICK_PRINTED IS NOT NULL" _
        '    & " and SOTPICK1.ORDR_NO_3PL IS NULL AND SOTPICK1.ERROR_REASON IS NOT NULL "

        ASCMAIN1.sql = "SELECT * FROM " & wkTable
        Fill_Records("SOTPICKO", String.Empty, True, ASCMAIN1.sql)

        Dim view As New DataView(dst.Tables("SOTPICKO"))
        view.RowFilter = "STATUS NOT LIKE 'C%' and STATUS NOT LIKE 'R%' AND ORDR_STATUS <> 'F'"
        grdSOTPICKO.DataSource = view

        Dim viewM As New DataView(dst.Tables("SOTPICKO"))
        viewM.RowFilter = "STATUS LIKE 'C%' or STATUS IS NULL"
        grdSOTPICKO_M.DataSource = viewM

        Dim viewR As New DataView(dst.Tables("SOTPICKO"))
        viewR.RowFilter = "STATUS LIKE 'R%'"
        grdSOTPICKO_R.DataSource = viewR

        Dim viewE As New DataView(dst.Tables("SOTPICKO"))
        viewE.RowFilter = "STATUS LIKE 'E%'"
        grdSOTPICKO_E.DataSource = viewE

        grdSOTPICKO_A.DataSource = dst.Tables("SOTPICKO")

        EnforceConstraints(True)

        Sort_grdColumns(grdSOTPICKO, "ORDR_GROUP_NO,ORDR_NO")
        Sort_grdColumns(grdSOTPICKO_M, "ORDR_GROUP_NO,ORDR_NO")
        Sort_grdColumns(grdSOTPICKO_R, "ORDR_GROUP_NO,ORDR_NO")
        Sort_grdColumns(grdSOTPICKO_E, "ORDR_GROUP_NO,ORDR_NO")
        Sort_grdColumns(grdSOTPICKO_A, "ORDR_GROUP_NO,ORDR_NO")

        grdSOTPICKO.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTPICKO_M.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTPICKO_R.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTPICKO_E.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTPICKO_A.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

    End Sub

    Sub Delete_Record()

        Try
            BeginTrans()

            CommitTrans("Delete Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICKO, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICKO_A, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICKO_E, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICKO_M, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICKO_R, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

        Load_Popup_Menu(tabOrders, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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


        End Select
    End Sub

#End Region

End Class