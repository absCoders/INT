Public Class ARTCUST8

    Dim sqlARTCUSTX As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            sqlARTCUSTX = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_STATUS, ARTCUST1.CUST_NAME" _
            & ", ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_GROUP_CODE" _
            & ", ARTCUST1.CUST_CREDIT_GROUP_CUST, ARTCUST1.CUST_BILL_TO_CUST" _
            & ", ARTCUST1.SREP_CODE, SOTSREP1.SREP_NAME" _
            & " from ARTCUST1,SOTSREP1 where SOTSREP1.SREP_CODE (+) = ARTCUST1.SREP_CODE"
            ASCMAIN1.sql = sqlARTCUSTX _
            & "  and ARTCUST1.CUST_GROUP_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ARTCUSTX", "**", 0, False, "V", 1)
        End With

        grdARTCUSTX.DataSource = dst.Tables("ARTCUSTX")

        With grdARTCUSTX.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
        End With

        Create_Summary(grdARTCUSTX, "CUST_CODE", "Count")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdARTCUSTX.DisplayLayout.Bands(0).Columns
            If gcol.Key <> "CUST_CODE" Then
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdARTCUSTX, "SS", "Show Filter", "Show GroupBox")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
            Case "Edit"

            Case "Update"

                'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

                'Dim rowSOTSREP1 = LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
                'If rowSOTSREP1 Is Nothing Then
                '    EMsg &= vbCr & "Invalid Value entered for Sales Rep Code"
                'End If


        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim sqlDelete = ""
        'Update_Record_TDA("ARTCUST1")

        ASCMAIN1.sql = "Update ARTCUST1 Set CUST_GROUP_CODE = NULL " _
        & " where CUST_GROUP_CODE = '" & Absx1.txtFor("CUST_GROUP_CODE").Text & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update ARTCUST1 Set CUST_GROUP_CODE = '" & Absx1.txtFor("CUST_GROUP_CODE").Text & "'" _
        & " where " & Get_List_of_Customers("ARTCUST1.CUST_CODE in")
        ASCDATA1.ExecuteSQL()

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
        Fill_Records("ARTCUSTX", New String() {Absx1.txtFor("CUST_GROUP_CODE").Text})
        Sort_grdColumns(grdARTCUSTX, "CUST_CODE")
        EnforceConstraints(True)

        grdARTCUSTX.Text = "Members of Customer Group " & Absx1.txtFor("CUST_GROUP_CODE").Text
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ARTCUSTX").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdARTCUSTX.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        With grdARTCUSTX.DisplayLayout.Override
            If EntryMode = "New" Or EntryMode = "Edit" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
                cmdCustomers.Visible = True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False

                cmdCustomers.Visible = False
            End If
        End With
    End Sub

#End Region

#Region "grdARTCUSTX"

    Private Sub grdARTCUSTX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTX.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                ASCMAIN1.sql = sqlARTCUSTX & " AND ARTCUST1.CUST_CODE = '" & e.Cell.Value & "'"
                Dim rowARTCUSTX As DataRow = ASCDATA1.GetDataRow
                If rowARTCUSTX IsNot Nothing Then
                    For Each dcol As DataColumn In rowARTCUSTX.Table.Columns
                        Dim COLUMN_NAME As String = dcol.ColumnName
                        If COLUMN_NAME <> "CUST_CODE" Then
                            grdARTCUSTX.ActiveRow.Cells(COLUMN_NAME).Value = rowARTCUSTX.Item(COLUMN_NAME)
                        End If
                    Next
                End If
        End Select
    End Sub

    Private Sub grdARTCUSTX_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTCUSTX.AfterRowActivate
        With grdARTCUSTX.DisplayLayout.Bands("ARTCUSTX")
            If grdARTCUSTX.ActiveRow.IsAddRow Then
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdARTCUSTX_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdARTCUSTX.BeforeRowsDeleted
     
    End Sub

    Private Sub grdARTCUSTX_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUSTX.BeforeRowUpdate

        Dim row As DataRow = LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
        
        If row Is Nothing Then
            e.Cancel = True
        End If
    End Sub

    Private Sub grdARTCUSTX_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTCUSTX.ClickCellButton
        Dim sql_where As String = Get_List_of_Customers("ARTCUST1.CUST_CODE not in")
        Call grdClickCellButton(grdARTCUSTX, sql_where, True)
    End Sub

    Private Sub grdARTCUSTX_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTCUSTX.Error
        grdARTCUSTX.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdARTCUSTX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdARTCUSTX.InitializeRow
        If e.Row.Cells("CUST_GROUP_CODE").Text <> "" And e.Row.Cells("CUST_GROUP_CODE").Text <> Absx1.txtFor("CUST_GROUP_CODE").Text Then
            e.Row.Cells("CUST_GROUP_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ARTCUSTX"), e.Row)
    End Sub

#End Region

    Private Sub cmdCustomers_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustomers.Click

        Dim sql_where As String = Get_List_of_Customers("ARTCUST1.CUST_CODE not in")
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE", , sql_where)


        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Customers")

                grdARTCUSTX.Visible = False
                For Each CUST_CODE As String In ASCMAIN1.CodeSelector.SelectedCodes
                    grdARTCUSTX.ActiveRow = grdARTCUSTX.DisplayLayout.Bands(0).AddNew
                    grdARTCUSTX.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
                    grdARTCUSTX.ActiveRow.Update()
                Next
                grdARTCUSTX.Visible = True
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            End If
        End If

    End Sub

    Function Get_List_of_Customers(ByVal sql_where_clause As String) As String
        Dim sql_where As String = ""
        Dim CUST_CODEs As String = ""
        For Each rowARTCUSTX As DataRow In dst.Tables("ARTCUSTX").Select
            CUST_CODEs &= ",'" & rowARTCUSTX.Item("CUST_CODE") & "'"
        Next
        If CUST_CODEs <> "" Then
            sql_where = sql_where_clause & " (" & Mid(CUST_CODEs, 2) & ")"
        End If
        Return sql_where
    End Function
     
End Class