Public Class ARTMALL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim ARTCUST2_Maintainable As String = "CUST_STORE_CITY,CUST_STORE_STATE,CUST_STORE_ZIP_CODE,MALL_CODE,LAST_DATE,LAST_OPER"
        With dst
            ASCMAIN1.sql = "Select ARTCUST2.*" _
                & " from ARTCUST2" _
                & " where ARTCUST2.MALL_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "V", 2, ARTCUST2_Maintainable)
        End With

        grdARTCUST2.DataSource = dst.Tables("ARTCUST2")
        Create_Summary(grdARTCUST2, "CUST_CODE", "Count")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdARTCUST2.DisplayLayout.Bands(0).Columns
            If Split(ARTCUST2_Maintainable, ",").Contains(gcol.Key) Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
            End If
        Next
        With grdARTCUST2.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Header.Fixed = True
            .Columns("CUST_STORE_NO").Header.Fixed = True
        End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCUST2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins")
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
 
            'Case "grdARTCUST2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

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

        End Select

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

                For Each ROW As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.ModifiedCurrent)
                    Dim CUST_CODE As String = ROW.Item("CUST_CODE")
                    Dim CUST_STORE_NO As String = ROW.Item("CUST_STORE_NO")
                    Dim CUST_STORE_STATE As String = ROW.Item("CUST_STORE_STATE") & ""
                    Dim MALL_CODE As String = ROW.Item("MALL_CODE") & ""
                    If LookUp("TATSTATE", CUST_STORE_STATE) Is Nothing Then
                        EMsg &= vbCr & "Invalid State for " & CUST_CODE & "-" & CUST_STORE_NO
                    End If
                    If LookUp("ARTMALL1", MALL_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Mall for " & CUST_CODE & "-" & CUST_STORE_NO
                    End If
                Next
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        For Each row As DataRow In dst.Tables("ARTCUST2").Select("", "", DataViewRowState.ModifiedCurrent)
            row.Item("LAST_OPER") = ASCMAIN1.USER_ID
            row.Item("LAST_DATE") = DATETIME_STAMP
        Next
        Update_Record_TDA("ARTCUST2")
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("ARTCUST2", New String() {Absx1.txtFor("MALL_CODE").Text})
        Sort_grdColumns(grdARTCUST2, "CUST_CODE")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ARTCUST2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdARTCUST2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCUST2}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False

                    '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
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

#Region "grdARTCUST2"

#End Region

    Private Sub grdARTCUST2_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdARTCUST2.BeforeRowUpdate

    End Sub

    Private Sub grdARTCUST2_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTCUST2.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdARTCUST2, sql_where)
    End Sub

    Private Sub grdARTCUST2_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdARTCUST2.InitializeLayout

    End Sub
End Class