Imports Infragistics.Win.UltraWinGrid

Public Class ICTWHSE1
    Private sqlSOTFORM2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SOTSVIA2.*, TATSTATE.STATE_NAME" _
                & " from SOTSVIA2,TATSTATE" _
                & " where SOTSVIA2.WHSE_CODE = :PARM1" _
                & "   and TATSTATE.STATE_CODE (+) = SOTSVIA2.STATE_CODE"
            Create_TDA(.Tables.Add, "SOTSVIA2", "**", 0, True, "V", 2)
            Create_TDA(.Tables.Add, "SOTSVIA3", "*", 1)

            dst.Tables("SOTSVIA3").Columns("TRANSIT_BUS_DAYS").DataType = GetType(System.Int16)

            Create_TDA(.Tables.Add, "ICTWHSE2", "*", 1)
            With dst.Tables("ICTWHSE2")
                .Columns.Add("ITEM_DESC", GetType(String))
                .Columns.Add("PROD_CODE", GetType(String))
                .Columns.Add("ITEM_CLASS_CODE", GetType(String))
                .Columns.Add("COLLECTION_CODE", GetType(String))
                .Columns.Add("ITEM_UPC_CODE", GetType(String))
            End With
        End With

        grdSOTSVIA2.DataSource = dst.Tables("SOTSVIA2")
        grdSOTSVIA3.DataSource = dst.Tables("SOTSVIA3")

        grdSOTSVIA3.DisplayLayout.Bands(0).Columns("PICKUP_DAY").Hidden = False
        grdSOTSVIA3.DisplayLayout.Bands(0).Columns("PICKUP_DAY").Header.Caption = "Pickup"
        grdSOTSVIA3.DisplayLayout.Bands(0).Columns("PICKUP_DAY").Width = 65
        tabTransitDays.Width = 365

        grdICTWHSE2.DataSource = dst.Tables("ICTWHSE2")
        Create_Summary(grdICTWHSE2, "ITEM_CODE", "Count")
        ASCMAIN1.Add_Value_List(grdICTWHSE2, "PROD_CODE")
        ASCMAIN1.Add_Value_List(grdICTWHSE2, "ITEM_CLASS_CODE")
        ASCMAIN1.Add_Value_List(grdICTWHSE2, "COLLECTION_CODE")

        With grdSOTSVIA2.DisplayLayout.Bands(0)
            .Columns("STATE_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        End With

        ASCMAIN1.Add_Value_List(grdSOTSVIA3, "PICKUP_DAY", , {":", "0:Sun", "1:Mon", "2:Tue", "3:Wed", "4:Thu", "5:Fri", "6:Sat"})
    End Sub

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWHSE2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Auto Fit Columns")
        Load_Popup_Menu(grdSOTSVIA2, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes", "Auto Fit Columns")
        Load_Popup_Menu(grdSOTSVIA3, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Auto Fit Columns")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case grd.Name

            Case "SOTSVIA2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            Case "Add Codes"
                If grd.Name = "SOTSVIA2" Then
                    Add_Codes(grdSOTSVIA2, "TATSTATE", "STATE_CODE", "States")
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Auto Fit Columns"
                grd.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
        End Select
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                ' This is a secondary check in case the user set the LP CODE after adding items.
                If Absx1.txtFor("LP_CODE").TextLength > 0 Then
                    Dim LP_CODE As String = Absx1.txtFor("LP_CODE").Text
                    dst.Tables("ICTWHSE2").AcceptChanges()
                    If dst.Tables("ICTWHSE2").Rows.Count > 0 Then
                        Dim lstItems As New List(Of String)
                        For Each drICTWHSE2 As DataRow In dst.Tables("ICTWHSE2").Select("")
                            lstItems.Add(drICTWHSE2.Item("ITEM_CODE") & String.Empty)
                        Next
                        ASCMAIN1.sql = "SELECT * FROM ICTWHSE2
                                    WHERE ITEM_CODE IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))
                                    AND WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE LP_CODE = :PARM2)
                                    AND WHSE_CODE <> :PARM3"
                        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTWHSE2", "VVV", {String.Join(",", lstItems.ToArray), LP_CODE, Absx1.txtFor("WHSE_CODE").Text})
                        If tbl.Rows.Count > 0 Then
                            Dim uMsg As String = "The following items are already assigned to another warehouse." & Environment.NewLine & Environment.NewLine
                            For Each dr As DataRow In tbl.Select("", "ITEM_CODE")
                                uMsg &= dr.Item("ITEM_CODE") & " - " & dr.Item("WHSE_CODE") & Environment.NewLine
                            Next
                            EMsg &= vbCr & uMsg
                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        Dim sqlDelete = "WHSE_CODE = '" & WHSE_CODE & "'"
        Update_Record_TDA("SOTSVIA2", sqlDelete)
        Update_Record_TDA("SOTSVIA3", sqlDelete)
        Update_Record_TDA("ICTWHSE2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        ' grpLOCATIONs.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked)
        EnforceConstraints(False)

        grpReturnsLocations.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked)
        grpVirtual.Visible = (Absx1.chkFor("WHSE_LOCATOR").Checked)

        EnforceConstraints(False)
        Fill_Records("SOTSVIA2", New String() {Absx1.txtFor("WHSE_CODE").Text})
        Sort_grdColumns(grdSOTSVIA2, "STATE_CODE")

        Fill_Records("SOTSVIA3", Absx1.txtFor("WHSE_CODE").Text)

        ASCMAIN1.sql = $"SELECT ICTWHSE2.*, 
                            ICTITEM1.ITEM_DESC, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.COLLECTION_CODE, 
                            NVL(ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE) ITEM_UPC_CODE
                            FROM ICTWHSE2, ICTITEM1 
                            WHERE ICTWHSE2.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                            AND ICTWHSE2.WHSE_CODE = '{Absx1.txtFor("WHSE_CODE").Text}'"
        Fill_Records("ICTWHSE2", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdICTWHSE2, "ITEM_CODE")

        EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SOTSVIA2", "SOTSVIA3", "ICTWHSE2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        Set_Read_Only_for_ctl(Absx1.chkFor("WHSE_LOCATOR"), Not (EntryMode = "New"))
        grdSOTSVIA2.Enabled = tf
        grdSOTSVIA3.Enabled = tf
        grdICTWHSE2.Enabled = tf

        tabTransitDays.Visible = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTSVIA2, grdSOTSVIA3, grdICTWHSE2}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

#End Region

#Region "grdSOTSVIA2"

    Private Sub grdSOTSVIA2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSVIA2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "STATE_CODE"
                grdCodeDesc(grdSOTSVIA2, "TATSTATE", "STATE_CODE", "STATE_NAME")
        End Select
    End Sub

    Private Sub grdSOTSVIA2_AfterExitEditMode(sender As Object, e As EventArgs) Handles grdSOTSVIA2.AfterExitEditMode
        With grdSOTSVIA2
            Select Case .ActiveCell.Column.Key
                Case "STATE_CODE"
                    Dim STATE_CODE As String = .ActiveCell.Text
                    If STATE_CODE <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(STATE_CODE, .ActiveCell.Column.Key)
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSOTSVIA2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSVIA2.AfterRowActivate

        With grdSOTSVIA2.DisplayLayout.Bands(0).Columns("STATE_CODE")
            If grdSOTSVIA2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSOTSVIA2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTSVIA2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowSOTSVIA2 As DataRow = dst.Tables("SOTSVIA2").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowSOTSVIA2.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdSOTSVIA2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSVIA2.BeforeRowUpdate

        Dim row As DataRow = LookUp("TATSTATE", e.Row.Cells("STATE_CODE").Text)
        If row Is Nothing Then
            e.Cancel = True
            MessageBox.Show("Valid State Code is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
        End If

    End Sub

    Private Sub grdSOTSVIA2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSVIA2.ClickCellButton
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        If grd.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key
            Case "STATE_CODE"
                sql_where = Get_List_of_Codes("TATSTATE.STATE_CODE not in", "SOTSVIA2", "STATE_CODE")
                grdClickCellButton(grd, sql_where, True)
        End Select
    End Sub

#End Region

#Region "grdSOTSVIA3"

    Private Sub grdSOTSVIA3_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSVIA3.BeforeRowUpdate

        Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Text
        Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Text

        Dim row As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
        If row Is Nothing Then
            e.Cancel = True
            MessageBox.Show("Invalid Customer / Store.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim TRANSIT_BUS_DAYS As Int32 = Val(e.Row.Cells("TRANSIT_BUS_DAYS").Value & String.Empty)
        If TRANSIT_BUS_DAYS < 1 Then
            e.Cancel = True
            MessageBox.Show("Transit Days must be greater 0.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Select Case CUST_CODE
            Case "ULTA"

            Case Else
                e.Row.Cells("PICKUP_DAY").Value = DBNull.Value
                MsgBox("Pickup Day only in use for ULTA", MsgBoxStyle.OkOnly, "Cannot set Pickup Day")
        End Select

        e.Row.Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
    End Sub

    Private Sub grdSOTSVIA3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTSVIA3.ClickCellButton
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        If grd.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key

            Case "CUST_CODE"
                grdClickCellButton(grd, sql_where, False)

            Case "CUST_STORE_NO"
                sql_where = "CUST_CODE = '" & e.Cell.Row.Cells("CUST_CODE").Value & "'"
                grdClickCellButton(grd, sql_where, False)
        End Select
    End Sub

#End Region

#Region "grdICTWHSE2"

    Private Sub grdICTWHSE2_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTWHSE2.BeforeRowUpdate

        Try
            e.Row.Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Text

            Dim drICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If drICTITEM1 Is Nothing Then
                MessageBox.Show($"Update Error: Invalid Item Code", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If

            Dim LP_CODE As String = Absx1.txtFor("LP_CODE").Text
            If LP_CODE.Length > 0 Then
                ASCMAIN1.sql = "SELECT * FROM ICTWHSE2
                                    WHERE ITEM_CODE = :PARM1
                                    AND WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE LP_CODE = :PARM2)
                                    AND WHSE_CODE <> :PARM3"
                Dim drICTWHSE2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVV", {ITEM_CODE, LP_CODE, Absx1.txtFor("WHSE_CODE").Text})
                If drICTWHSE2 IsNot Nothing Then
                    MessageBox.Show($"Update Error: Invalid Item Code / Warehouse Code combination. The item exists in warehouse {drICTWHSE2.Item("WHSE_CODE")}", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                    Exit Sub
                End If
            End If

            e.Row.Cells("ITEM_DESC").Value = drICTITEM1.Item("ITEM_DESC") & String.Empty
            e.Row.Cells("PROD_CODE").Value = drICTITEM1.Item("PROD_CODE") & String.Empty
            e.Row.Cells("ITEM_CLASS_CODE").Value = drICTITEM1.Item("ITEM_CLASS_CODE") & String.Empty
            e.Row.Cells("COLLECTION_CODE").Value = drICTITEM1.Item("COLLECTION_CODE") & String.Empty

            If drICTITEM1.Item("ITEM_UPC_CODE") & String.Empty <> String.Empty Then
                e.Row.Cells("ITEM_UPC_CODE").Value = drICTITEM1.Item("ITEM_UPC_CODE") & String.Empty
            Else
                e.Row.Cells("ITEM_UPC_CODE").Value = drICTITEM1.Item("ITEM_EAN_CODE") & String.Empty
            End If

            If e.Row.Cells("INIT_OPER").Value & String.Empty = String.Empty Then
                e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                e.Row.Cells("INIT_DATE").Value = DateTime.Now
            End If

        Catch ex As Exception
            e.Cancel = True
            MessageBox.Show($"Update Error: {ex.Message}", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub grdICTWHSE2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTWHSE2.ClickCellButton
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        If grd.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                sql_where = $"NVL(ITEM_STATUS, 'A') = 'A' 
                                AND ITEM_CODE NOT IN 
                                (
                                    SELECT ITEM_CODE FROM ICTWHSE2 WHERE WHSE_CODE IN (SELECT WHSE_CODE FROM ICTWHSE1 WHERE LP_CODE = '{Absx1.txtFor("LP_CODE").Text}') 
                                )"
                grdClickCellButton(grd, sql_where, True)
        End Select
    End Sub

#End Region

End Class