Public Class SOF3PLER

    Private wkShipHeaderTable As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            GetClarinsShipmentsWithErrors()

            Dim sql As String = String.Empty

            sql = "SELECT SHIPHDR.PROCESS_IND, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_GROUP_NO"
            sql &= ", SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTSHIP1.SHIP_BOL_NO"
            sql &= " from sotship1, sotordr0," & wkShipHeaderTable & " shiphdr"
            sql &= " where sotship1.ordr_group_no = sotordr0.ordr_group_no"
            sql &= " and sotship1.ship_bol_no = shiphdr.ship_bol_no"

            Create_TDA(.Tables.Add, "SOTSHIPX", sql, 0, False, String.Empty, 0)

            sql = "Select SOTPICK1.SHIP_BOL_NO, SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO, SOTPICK1.PICK_RELEASED, ship_hdr.OHORDN, SHIP_HDR.OHINVN"
            sql &= " FROM SOTPICK1,  conv.cfg_shiphdr SHIP_HDR"
            sql &= " WHERE SOTPICK1.PICK_NO = SHIP_HDR.ABSPICKNBR (+)"
            sql &= " and sotpick1.ship_bol_no in (select ship_bol_no from " & wkShipHeaderTable & ")"
            Create_TDA(.Tables.Add, "SOTPICKX", sql, 0, False, String.Empty, 0)

            Dim wkSHIPDTL As String = "SELECT SAINVN INV_NO, SAITEM ITEM_CODE, SUM(SAQTYS) UNITS" _
                                      & " FROM CONV.CFG_SHIPDTL" _
                                      & " WHERE SAINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR where SHIP_BOL_NO IN (select ship_bol_no from " & wkShipHeaderTable & "))" _
                                      & " GROUP BY SAINVN, SAITEM HAVING SUM(SAQTYS) <> 0 "

            Dim wkCARTON As String = "SELECT CHINVN INV_NO, CDITEM ITEM_CODE, SUM(CDQTYS * CDSKMQ) UNITS" _
                                     & " FROM CONV.CFG_CARTON" _
                                     & "  WHERE CHINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR where SHIP_BOL_NO IN (select ship_bol_no from " & wkShipHeaderTable & "))" _
                                     & " GROUP BY CHINVN, CDITEM"

            sql = " Select A.INV_NO, A.ITEM_CODE, A.UNITS DTL_UNITS, B.UNITS CTN_UNITS"
            sql &= " FROM"
            sql &= " ("
            sql &= wkSHIPDTL
            sql &= " ) A,"
            sql &= " ("
            sql &= wkCARTON
            sql &= " ) B"
            sql &= " WHERE A.INV_NO = B.INV_NO AND A.ITEM_CODE = B.ITEM_CODE"

            Create_TDA(.Tables.Add, "SHIPDTLX", sql, 0, False, String.Empty, 0)

            .Relations.Add("SOTSHIPX_SOTPICKX", dst.Tables("SOTSHIPX").Columns("SHIP_BOL_NO"), dst.Tables("SOTPICKX").Columns("SHIP_BOL_NO"))
            .Relations.Add("SOTPICKX_SHIPDTLX", dst.Tables("SOTPICKX").Columns("OHINVN"), dst.Tables("SHIPDTLX").Columns("INV_NO"))

            With dst.Tables("SOTPICKX")
                .Columns.Add("TOT_DTL_UNITS", GetType(System.Int32), "SUM(CHILD.DTL_UNITS)")
                .Columns.Add("TOT_CTN_UNITS", GetType(System.Int32), "SUM(CHILD.CTN_UNITS)")
            End With

            grdSOTSHIP1.DataSource = dst.Tables("SOTSHIPX")

            ASCMAIN1.Add_Value_List(grdSOTSHIP1, "PROCESS_IND", , New String() {":", "M:Missing Pick Tickets", "P:Mismatch Quantities"})

        End With

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

        grdSOTSHIP1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIPX", "SOTPICKX", "SHIPDTLX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Clear_All_Filters(grdSOTSHIP1)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        GetClarinsShipmentsWithErrors()

        Fill_Records("SOTSHIPX")
        Fill_Records("SOTPICKX")
        Fill_Records("SHIPDTLX")

        EnforceConstraints(True)

        Sort_grdColumns(grdSOTSHIP1, "CUST_CODE,ORDR_DATE")

        grdSOTSHIP1.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIP1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

#Region "ABSColumn Controls"

#End Region

    Private Sub GetClarinsShipmentsWithErrors()

        Dim sql As String = " Select distinct ship_bol_no, process_ind from conv.cfg_shiphdr where process_ind IN ('M', 'P')"

        If wkShipHeaderTable.Length = 0 Then
            wkShipHeaderTable = ASCMAIN1.Temp_Table(sql)
            Exit Sub
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & wkShipHeaderTable)

        ASCDATA1.ExecuteSQL("Insert Into " & wkShipHeaderTable & " " & sql)

    End Sub


    Private Sub grdSOTSHIP1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIP1.InitializeRow

        If e.Row.Band.Key = grdSOTSHIP1.DisplayLayout.Bands(2).Key Then
            If e.Row.Cells("DTL_UNITS").Value <> e.Row.Cells("CTN_UNITS").Value Then
                e.Row.Appearance.BackColor = Drawing.Color.Red
            End If
        ElseIf e.Row.Band.Key = grdSOTSHIP1.DisplayLayout.Bands(1).Key Then
            If e.Row.Cells("TOT_DTL_UNITS").Value <> e.Row.Cells("TOT_CTN_UNITS").Value Then
                e.Row.Appearance.BackColor = Drawing.Color.Red
            End If
        End If
    End Sub
End Class