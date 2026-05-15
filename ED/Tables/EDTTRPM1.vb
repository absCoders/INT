Public Class EDTTRPM1

    Dim sqlEDTTRPM2 As String = ""
    Dim sqlEDTTRPM3 As String = ""
    Dim sqlEDTUPCX1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            sqlEDTTRPM2 = "Select EDTTRPM2.*, ARTCUST1.CUST_NAME" _
                & " from EDTTRPM2,EDTTRPM1,ARTCUST1" _
                & " where ARTCUST1.CUST_CODE (+) = EDTTRPM2.CUST_CODE" _
                & "   and EDTTRPM2.EDI_TP_QUAL = EDTTRPM1.EDI_TP_QUAL" _
                & "   and EDTTRPM2.EDI_TP_ID = EDTTRPM1.EDI_TP_ID" _
                & "   and EDTTRPM2.EDI_DOC_NO = EDTTRPM1.EDI_DOC_NO"
            ASCMAIN1.sql = sqlEDTTRPM2 _
                & "   and EDTTRPM1.EDI_TP_QUAL = :PARM1" _
                & "   and EDTTRPM1.EDI_TP_ID = :PARM2" _
                & "   and EDTTRPM1.EDI_DOC_NO = :PARM3"
            Create_TDA(.Tables.Add, "EDTTRPM2", "**", 0, True, "VVV", 4)

            sqlEDTTRPM3 = "Select EDTTRPM3.*, ARTCUST1.CUST_NAME" _
                & " from EDTTRPM3,EDTTRPM1,ARTCUST1" _
                & " where ARTCUST1.CUST_CODE (+) = EDTTRPM3.CUST_CODE" _
                & "   and EDTTRPM3.EDI_TP_QUAL = EDTTRPM1.EDI_TP_QUAL" _
                & "   and EDTTRPM3.EDI_TP_ID = EDTTRPM1.EDI_TP_ID" _
                & "   and EDTTRPM3.EDI_DOC_NO = EDTTRPM1.EDI_DOC_NO"
            ASCMAIN1.sql = sqlEDTTRPM3 _
                & "   and EDTTRPM1.EDI_TP_QUAL = :PARM1" _
                & "   and EDTTRPM1.EDI_TP_ID = :PARM2" _
                & "   and EDTTRPM1.EDI_DOC_NO = :PARM3"
            Create_TDA(.Tables.Add, "EDTTRPM3", "**", 0, True, "VVV", 4)

            sqlEDTUPCX1 = "Select EDTUPCX1.*, ARTCUST1.CUST_NAME FROM EDTUPCX1, ARTCUST1 WHERE ARTCUST1.CUST_CODE (+) = EDTUPCX1.CUST_CODE"
            ASCMAIN1.sql = sqlEDTUPCX1 & " AND EDTUPCX1.CUST_CODE = :PARM1"
            Create_TDA(dst.Tables.Add, "EDTUPCX1", "**", 0, True, "V", 2)

            'Create_Relation("EDTTRPM1", "EDTUPCX1", "CUST_CODE", "CUST_CODE")
            'dst.Relations("EDTTRPM1_EDTUPCX1").ChildKeyConstraint.UpdateRule = Rule.Cascade


        End With

        grdEDTTRPM2.DataSource = dst.Tables("EDTTRPM2")
        grdEDTTRPM3.DataSource = dst.Tables("EDTTRPM3")
        grdEDTUPCX1.DataSource = dst.Tables("EDTUPCX1")

        'With grdEDTTRPM2.DisplayLayout.Bands(0)
        '    .Columns("CUST_CODE").Header.Fixed = True
        '    .Columns("CUST_NAME").Header.Fixed = True
        'End With

        Create_Summary(grdEDTTRPM2, "CUST_CODE", "Count")
        Create_Summary(grdEDTTRPM3, "CUST_CODE", "Count")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdEDTTRPM2.DisplayLayout.Bands(0).Columns
            If gcol.Key <> "CUST_CODE" And gcol.Key <> "EDI_DEPT_NO" Then
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        For Each gcol As UltraWinGrid.UltraGridColumn In grdEDTTRPM3.DisplayLayout.Bands(0).Columns
            If gcol.Key <> "CUST_CODE" And gcol.Key <> "EDI_STORE" Then
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        With grdEDTUPCX1.DisplayLayout.Bands("EDTUPCX1")
            If .Columns.Exists("EDI_ITEM_CODE") Then
                .Columns("EDI_ITEM_CODE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                .Columns("EDI_ITEM_CODE").CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            End If
        End With

        With grdEDTUPCX1.DisplayLayout
            .Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        End With


    End Sub
 
#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                Dim newCustCode As String = Absx1.txtFor("CUST_CODE").Text
                If String.IsNullOrWhiteSpace(newCustCode) Then
                    EMsg += "CUST_CODE required" & vbCrLf
                    MessageBox.Show(
                        "You cannot remove the Cust Code from this screen.",
                        "Validation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    )
                    Exit Sub
                End If
                Dim bad As New List(Of String)

                Dim t As DataTable = dst.Tables("EDTUPCX1")
                If t IsNot Nothing Then
                    For Each r As DataRow In t.Rows
                        If r.RowState = DataRowState.Deleted Then Continue For

                        Dim ignore As Boolean = False
                        If t.Columns.Contains("IGNORE") AndAlso Not IsDBNull(r("IGNORE")) Then
                            ignore = CBool(r("IGNORE"))
                        End If

                        Dim itemCode As String = (If(r("ITEM_CODE"), "") & "").Trim()

                        If ignore AndAlso itemCode <> "" Then
                            Dim cust = (If(r("CUST_CODE"), "") & "").Trim()
                            Dim upc = If(t.Columns.Contains("EDI_ITEM_CODE"), (If(r("EDI_ITEM_CODE"), "") & "").Trim(), "")
                            bad.Add($"Cust {cust} | UPC Sent {upc} | Item {itemCode}")
                        End If
                    Next
                End If

                If bad.Count > 0 Then
                    Dim detail As String = String.Join(vbCrLf, bad.Take(10))
                    If bad.Count > 10 Then detail &= vbCrLf & $"(+{bad.Count - 10} more...)"

                    MessageBox.Show(
                        "You cannot check 'Ignore' for a UPC if a replacement Item Code is provided." & vbCrLf &
                        "Please uncheck Ignore or clear ITEM_CODE on these rows:" & vbCrLf & vbCrLf & detail,
                        "Validation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    )
                    Exit Sub
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = "EDI_TP_QUAL = '" & Absx1.txtFor("EDI_TP_QUAL").Text & "' and EDI_TP_ID = '" & Absx1.txtFor("EDI_TP_ID").Text & "' and EDI_DOC_NO = '" & Absx1.txtFor("EDI_DOC_NO").Text & "'"

        Update_Record_TDA("EDTTRPM2", sqlDelete)
        Update_Record_TDA("EDTTRPM3", sqlDelete)

        '' ISSUE-7371 - this screen allows the user to change the CUST_CODE so we have to delete the existing rows
        ' And there are too many different scenarios to have a reliable way to update EDTUPCX1. So we are going to leave it as-is
        ' and not try to guess what the user wants to do



        ''Dim t As DataTable = dst.Tables("EDTUPCX1")
        ''Dim old_cust_code As String = Absx1.txtFor("CUST_CODE").Text
        ''If t IsNot Nothing Then
        ''    For Each r As DataRow In t.Rows
        ''        If r.RowState = DataRowState.Deleted Then Continue For
        ''        If r("CUST_CODE") <> Absx1.txtFor("CUST_CODE").Text Then
        ''            old_cust_code = r("CUST_CODE")
        ''            r("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
        ''        End If
        ''    Next
        ''End If
        'Dim custCodeOrig As String = Absx1.txtFor("CUST_CODE").Text
        'If EntryMode = "Edit" Then
        '    custCodeOrig = dst.Tables("EDTTRPM1").Rows(0)?.Item("CUST_CODE", DataRowVersion.Original) & ""
        '    Dim custCodeCurrent As String = Absx1.txtFor("CUST_CODE").Text
        '    If (custCodeOrig <> custCodeCurrent) Then
        '        dst.Tables("EDTUPCX1").AcceptChanges()
        '        For Each r As DataRow In dst.Tables("EDTUPCX1").Rows
        '            r.SetAdded()
        '            r.Item("CUST_CODE") = custCodeCurrent
        '        Next
        '    End If
        '    Dim sqlDelete2 = $"CUST_CODE = '{custCodeOrig}'"
        '    Update_Record_TDA("EDTUPCX1", sqlDelete2)
        'End If

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        'If EntryMode = "New" Then
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        'Else
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("EDTTRPM2", New String() {Absx1.txtFor("EDI_TP_QUAL").Text, Absx1.txtFor("EDI_TP_ID").Text, Absx1.txtFor("EDI_DOC_NO").Text})
        Fill_Records("EDTTRPM3", New String() {Absx1.txtFor("EDI_TP_QUAL").Text, Absx1.txtFor("EDI_TP_ID").Text, Absx1.txtFor("EDI_DOC_NO").Text})
        Fill_Records("EDTUPCX1", New String() {Absx1.txtFor("CUST_CODE").Text})
        Sort_grdColumns(grdEDTTRPM2, "CUST_CODE")
        Sort_grdColumns(grdEDTTRPM3, "CUST_CODE")
        Sort_grdColumns(grdEDTUPCX1, "ITEM_CODE")
        EnforceConstraints(True)

        'grdEDTTRPM2.Text = "Override Customer by Dept"
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("EDTTRPM2").Rows.Clear()
            dst.Tables("EDTTRPM3").Rows.Clear()
            dst.Tables("EDTUPCX1").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdEDTTRPM2.Enabled = tf
        grdEDTTRPM3.Enabled = tf
        grdEDTUPCX1.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdEDTTRPM2, grdEDTTRPM3, grdEDTUPCX1}
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

        With grdEDTUPCX1.DisplayLayout.Override
            .AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        End With

        With grdEDTUPCX1.DisplayLayout.Bands("EDTUPCX1")
            If .Columns.Exists("EDI_ITEM_CODE") Then
                .Columns("EDI_ITEM_CODE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            End If
        End With

    End Sub

#End Region

#Region "grdEDTTRPM2"

    Private Sub grdEDTTRPM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDTTRPM2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                grdCodeDesc(grdEDTTRPM2, "ARTCUST1", "CUST_CODE", "CUST_NAME")

        End Select
    End Sub

    Private Sub grdEDTTRPM2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDTTRPM2.AfterRowActivate
        With grdEDTTRPM2.DisplayLayout.Bands("EDTTRPM2")
            If grdEDTTRPM2.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdEDTTRPM2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdEDTTRPM2.BeforeRowsDeleted

    End Sub

    Private Sub grdEDTTRPM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdEDTTRPM2.BeforeRowUpdate

        Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdEDTTRPM2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDTTRPM2.ClickCellButton
        Dim sql_where As String = "" ' Get_List_of_Customers("ARTCUST1.CUST_CODE not in")
        grdClickCellButton(grdEDTTRPM2, sql_where, True)
    End Sub

    Private Sub grdEDTTRPM2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdEDTTRPM2.Error
        grdEDTTRPM2.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

#End Region


#Region "grdEDTTRPM3"

    Private Sub grdEDTTRPM3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDTTRPM3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                grdCodeDesc(grdEDTTRPM3, "ARTCUST1", "CUST_CODE", "CUST_NAME")

        End Select
    End Sub

    Private Sub grdEDTTRPM3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDTTRPM3.AfterRowActivate
        With grdEDTTRPM3.DisplayLayout.Bands("EDTTRPM3")
            If grdEDTTRPM3.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdEDTTRPM3_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdEDTTRPM3.BeforeRowsDeleted

    End Sub

    Private Sub grdEDTTRPM3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdEDTTRPM3.BeforeRowUpdate

        Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdEDTTRPM3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDTTRPM3.ClickCellButton
        Dim sql_where As String = "" ' Get_List_of_Customers("ARTCUST1.CUST_CODE not in")
        grdClickCellButton(grdEDTTRPM3, sql_where, True)
    End Sub

    Private Sub grdEDTTRPM3_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdEDTTRPM3.Error
        grdEDTTRPM3.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

#End Region
#Region "grdEDTUPCX1"
    Private Sub grdEDTUPCX1_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdEDTUPCX1.BeforeCellUpdate

        Dim row = e.Cell.Row
        Dim hasIgnoreCol As Boolean = row.Cells.Exists("IGNORE")
        Dim hasItemCol As Boolean = row.Cells.Exists("ITEM_CODE")

        Dim currentIgnore As Boolean = False
        If hasIgnoreCol AndAlso Not IsDBNull(row.Cells("IGNORE").Value) Then
            currentIgnore = CBool(row.Cells("IGNORE").Value)
        End If

        Select Case e.Cell.Column.Key

            Case "IGNORE"
                Dim proposedIgnore As Boolean = False
                If e.NewValue IsNot Nothing AndAlso e.NewValue IsNot DBNull.Value Then
                    proposedIgnore = CBool(e.NewValue)
                End If

                Dim itemCode As String = (If(row.Cells("ITEM_CODE").Value, "") & "").Trim()
                If proposedIgnore AndAlso itemCode <> "" Then
                    MessageBox.Show("You can’t ignore a UPC that already has a replacement Item Code. Clear ITEM_CODE or leave Ignore unchecked.",
                                "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    e.Cancel = True
                    Try
                        If e.Cell.IsInEditMode Then
                            grdEDTUPCX1.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.UndoCell)
                        End If
                    Catch
                    End Try
                    Exit Sub
                End If

            Case "ITEM_CODE"
                Dim proposed As String = If(e.NewValue, "").ToString().Trim().ToUpper()

                If currentIgnore AndAlso proposed <> "" Then
                    MessageBox.Show("You must uncheck the Ignore checkbox before entering a replacement Item Code.",
                                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    e.Cancel = True
                    Exit Sub
                End If

                If proposed <> "" AndAlso Not IsValidItem(proposed) Then
                    MessageBox.Show("Invalid Item Code. Please enter a valid Item Code.",
                                    "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    e.Cancel = True
                    Exit Sub
                End If

        End Select

    End Sub

    Private Sub grdEDTUPCX1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDTUPCX1.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                'grdCodeDesc(grdEDTUPCX1, "ARTCUST1", "CUST_CODE", "CUST_NAME")
        End Select
    End Sub

    Private Sub grdEDTUPCX1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDTUPCX1.AfterRowActivate
        With grdEDTUPCX1.DisplayLayout.Bands("EDTUPCX1")
            If grdEDTUPCX1.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdEDTUPCX1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdEDTUPCX1.BeforeRowsDeleted

    End Sub

    Private Sub grdEDTUPCX1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdEDTUPCX1.BeforeRowUpdate

        Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        Dim item As String = If(e.Row.Cells("ITEM_CODE").Value, "").ToString().Trim().ToUpper()
        If Not IsValidItem(item) Then
            MessageBox.Show("Invalid Item Code. Please enter a valid ITEM_CODE.",
                            "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            e.Cancel = True : Exit Sub
        End If

        e.Row.Cells("ITEM_CODE").Value = item

    End Sub

    Private Sub grdEDTUPCX1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDTUPCX1.ClickCellButton
        Dim sql_where As String = "" ' Get_List_of_Customers("ARTCUST1.CUST_CODE not in")
        grdClickCellButton(grdEDTUPCX1, sql_where, True)
    End Sub

    Private Sub grdEDTUPCX1_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdEDTUPCX1.Error
        grdEDTUPCX1.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub
    Private Function IsValidItem(itemCode As String) As Boolean
        If String.IsNullOrWhiteSpace(itemCode) Then Return True
        ' Nonblank must exist in ICTITEM1
        Dim r As DataRow = LookUp("ICTITEM1", itemCode.Trim().ToUpper())
        Return (r IsNot Nothing)
    End Function



#End Region


End Class