Public Class ASFCONV1
    Private _E As Log_Entity
    Public specific_tab As String = ""
    Public single_row_grid As Boolean = False
    Public tblTATCONV1 As DataTable
    Public sqlTATCONV1_where As String = ""

    Public Sub New(ByVal FF As ASFBASE1, ByVal E As Log_Entity)
        frmASFBASE1 = FF
        _E = E
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select TATCONV1.* from TATCONV1 " & vbCrLf _
            & " where NVL(CONV_STATUS,'O') <> 'D'" & vbCrLf _
            & "   and TABLE_NAME = :PARM1 and TABLE_KEY = :PARM2" & vbCrLf _
            & sqlTATCONV1_where
            Create_TDA(.Tables.Add, "TATCONV1", "**", 0, , "VV", 1)
            .Tables("TATCONV1").Columns.Add("CONV_ATTACHMENTS", GetType(System.Int64))
        End With

        tblTATCONV1 = dst.Tables("TATCONV1")
        grdTATCONV1.DataSource = dst.Tables("TATCONV1")

        With grdTATCONV1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CONV_FOLLOWUP_BY", "CONV_FOLLOWUP_DATE", "LAST_OPER", "LAST_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.ForeColor = Color.DodgerBlue
            Next
            For Each COLUMN_NAME As String In New String() {"CONV_PROMISE_AMT", "CONV_PROMISE_BY"}
                .Columns(COLUMN_NAME).Header.Appearance.ForeColor = Color.Red
            Next
            .Columns("CONV_ATTACHMENTS").Header.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "ATTACH")
            .Columns("CONV_ATTACHMENTS").Header.Appearance.ImageHAlign = HAlign.Center
            .Columns("CONV_ATTACHMENTS").Header.Appearance.ImageVAlign = VAlign.Middle
            .Columns("CONV_ATTACHMENTS").Header.Caption = ""
        End With

        Fill_Records("TATCONV1", New String() {_E.TABLE_NAME, _E.TABLE_KEY})
        Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)

        Me.Text = _E.TABLE_KEY_CAPTION & " " & _E.TABLE_KEY & ":" & _E.TABLE_KEY_DESC

        If specific_tab <> "" Then
            If specific_tab = "G" Then grdTATCONV1.Parent = tabTATCONV1.Parent
            If specific_tab = "T" Then tvwTATCONV1.Parent = tabTATCONV1.Parent
            tabTATCONV1.Visible = False
        End If

        If single_row_grid Then
            grdTATCONV1.DisplayLayout.Bands(0).Override.RowSizing = UltraWinGrid.RowSizing.Fixed
            grdTATCONV1.DisplayLayout.Bands(0).Override.RowSizing = UltraWinGrid.RowSizing.Free
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
    End Sub


#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdTATCONV1, "BBBB", "Add to Log", "Show Log", "Edit Log", "Follow-Up")
        Load_Popup_Menu(tvwTATCONV1, "BBBBBB", "Add to Log", "Show Log", "Edit Log", "Follow-Up", "Expand All", "Attachments")
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

        If tlb_pop.Tools.Exists("Add to Log") Then
            tlb_btn = DirectCast(tlb_pop.Tools("Add to Log"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Enabled = Not _E.read_only
            tlb_btn = DirectCast(tlb_pop.Tools("Show Log"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Enabled = False
            tlb_btn = DirectCast(tlb_pop.Tools("Follow-Up"), UltraWinToolbars.ButtonTool)
            tlb_btn.SharedProps.Enabled = False
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdTATCONV1"
                    Dim btnEnabled_Show As Boolean = False
                    Dim btnEnabled_Follow_Up As Boolean = False
                    If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then
                        Dim CONV_NO As String = grd.ActiveRow.Cells("CONV_NO").Text
                        Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                        If rowTATCONV1 IsNot Nothing Then
                            btnEnabled_Show = True
                            btnEnabled_Follow_Up = Not _E.read_only
                        End If
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Show Log"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = btnEnabled_Show
                    tlb_btn = DirectCast(tlb_pop.Tools("Follow-Up"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = btnEnabled_Follow_Up

                Case "tvwTATCONV1"
                    Dim btnEnabled_Show As Boolean = False
                    Dim btnEnabled_Follow_Up As Boolean = False
                    If tvwTATCONV1.ActiveNode IsNot Nothing Then
                        Dim CONV_NO As String = tvwTATCONV1.ActiveNode.Cells("CONV_NO").Text
                        Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                        If rowTATCONV1 IsNot Nothing Then
                            btnEnabled_Show = True
                            btnEnabled_Follow_Up = Not _E.read_only
                        End If
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Show Log"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = btnEnabled_Show
                    tlb_btn = DirectCast(tlb_pop.Tools("Follow-Up"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = btnEnabled_Follow_Up

            End Select
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Add to Log"

                If Not _E.TABLE_KEY_locked Then
                    If Not ASCMAIN1.Logical_Lock(_E.TABLE_NAME, _E.TABLE_KEY, , , , 1) Then Exit Sub
                End If
                Dim F As New ASFCONV2(Me, _E.TABLE_NAME, _E.TABLE_KEY)
                F.EntryMode = "N"
                F.ShowDialog()
                If F.result = "U" Then
                    dst.Tables("TATCONV1").Rows.Add(F.rowTATCONV1.ItemArray)
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)
                    If tabTATCONV1.SelectedTab.Key = "Tree View" Then
                        Setup_tvwTATCONV1()
                    End If
                    Update_Record_TDA("TATCONV1")
                End If
                F.Dispose()
                ASCMAIN1.MultiTask_Release(, , 1)

            Case "Expand All"
                tvwTATCONV1.ExpandAll(UltraWinTree.ExpandAllType.Always)

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge"
            '    Call Log_Ack_FollowUp()

            Case "Show Log"
                Dim CONV_NO As String
                If e.Tool.OwningMenu.Key = "tvwTATCONV1" Then
                    CONV_NO = tvwTATCONV1.ActiveNode.Cells("CONV_NO").Text
                Else
                    CONV_NO = grd.ActiveRow.Cells("CONV_NO").Text
                End If

                Dim F As New ASFCONV2(Me, _E.TABLE_NAME, _E.TABLE_KEY)
                F.EntryMode = "V"
                F.rowTATCONV1 = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.ShowDialog()
                F.Dispose()

            Case "Edit Log"
                Dim CONV_NO As String
                If e.Tool.OwningMenu.Key = "tvwTATCONV1" Then
                    CONV_NO = tvwTATCONV1.ActiveNode.Cells("CONV_NO").Text
                Else
                    CONV_NO = grd.ActiveRow.Cells("CONV_NO").Text
                End If

                If Not _E.TABLE_KEY_locked Then
                    If Not ASCMAIN1.Logical_Lock(_E.TABLE_NAME, _E.TABLE_KEY, , , , 1) Then Exit Sub
                End If
                If Not ASCMAIN1.Logical_Lock("TATCONV1", CONV_NO, , , , 1) Then Exit Sub

                Dim F As New ASFCONV2(Me, _E.TABLE_NAME, _E.TABLE_KEY)
                F.EntryMode = "E"
                F.followup_is_mandatory = Not (EntryMode = "N" Or EntryMode = "E")
                F.rowTATCONV1 = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.ShowDialog()
                If F.result = "U" Then
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)

                    If tabTATCONV1.SelectedTab.Key = "Tree View" Then
                        Setup_tvwTATCONV1()
                    End If
                    Update_Record_TDA("TATCONV1")
                End If
                F.Dispose()
                ASCMAIN1.MultiTask_Release(, , 1)

            Case "Follow-Up"
                If Not _E.TABLE_KEY_locked Then
                    If Not ASCMAIN1.Logical_Lock(_E.TABLE_NAME, _E.TABLE_KEY, , , , 1) Then Exit Sub
                End If

                Dim CONV_NO As String
                If e.Tool.OwningMenu.Key = "tvwTATCONV1" Then
                    CONV_NO = tvwTATCONV1.ActiveNode.Cells("CONV_NO").Text
                Else
                    CONV_NO = grd.ActiveRow.Cells("CONV_NO").Text
                End If

                If Not ASCMAIN1.Logical_Lock("TATCONV1", CONV_NO, False, True, False, 1) Then
                    Exit Sub
                End If

                Dim F As New ASFCONV2(Me, _E.TABLE_NAME, _E.TABLE_KEY)
                F.EntryMode = "F" ' "E"
                Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find(CONV_NO)
                F.rowTATCONV1_PREV = rowTATCONV1
                F.ShowDialog()
                If F.result = "U" Then
                    dst.Tables("TATCONV1").Rows.Add(F.rowTATCONV1.ItemArray)
                    Sort_grdColumns(grdTATCONV1, "INIT_DATE".ToLower)
                    If tabTATCONV1.SelectedTab.Key = "Tree View" Then
                        Setup_tvwTATCONV1()
                    End If
                    Update_Record_TDA("TATCONV1")
                End If
                F.Dispose()
                ASCMAIN1.MultiTask_Release(, , 1)

            Case "Attachments"

                If tvwTATCONV1.ActiveNode IsNot Nothing Then
                    Dim ENTITY As New Dropped_On_Entity
                    ENTITY.TABLE_NAME = "TATCONV1"
                    ENTITY.COLUMN_NAME = "CONV_NO"
                    ENTITY.CODE_VALUE = tvwTATCONV1.ActiveNode.Cells("CONV_NO").Value
                    Dim DESC_VALUE = "Log by " & tvwTATCONV1.ActiveNode.Cells("INIT_OPER").Value _
                        & " " & Format(tvwTATCONV1.ActiveNode.Cells.Item("INIT_DATE").Value, "MM/dd/yyyy HH:mm") _
                        & " (" & tvwTATCONV1.ActiveNode.Cells("CONV_SUBJECT").Value & ")"
                    ENTITY.DESC_VALUE = DESC_VALUE
                    ENTITY.ATTACHMENT_NOTES = ""

                    Dim F As New ASFATTA1
                    F.ENTITY = ENTITY
                    F.ShowDialog()
                    F.Dispose()

                    tvwTATCONV1.ActiveNode.Cells("CONV_ATTACHMENTS").Value = Get_CONV_ATTACHMENTS(tvwTATCONV1.ActiveNode.Cells("CONV_NO").Value)
                End If

        End Select
    End Sub

#End Region


#Region "tvwTATCONV1"
    Sub Setup_tvwTATCONV1()
        tvwTATCONV1.Nodes.Clear()
        Load_tvwTATCONV1("", Nothing)
    End Sub

    Sub Load_tvwTATCONV1(ByVal CONV_NO_PREV As String, ByVal anode As UltraWinTree.UltraTreeNode)
        Dim tnode As UltraWinTree.UltraTreeNode
        For Each rowTATCONV1 As DataRow In dst.Tables("TATCONV1") _
            .Select("ISNULL(CONV_NO_PREV,'') = '" & CONV_NO_PREV & "'", "CONV_NO")
            Dim NODE_TEXT As String = rowTATCONV1.Item("CONV_SUBJECT") & ""
            If anode Is Nothing Then
                tnode = tvwTATCONV1.Nodes.Add(rowTATCONV1.Item("CONV_NO"), NODE_TEXT)
            Else
                tnode = anode.Nodes.Add(rowTATCONV1.Item("CONV_NO"), NODE_TEXT)
            End If
            For I As Integer = 0 To tnode.Cells.Count - 1
                Dim COLUMN_NAME As String = dst.Tables("TATCONV1").Columns(I).ColumnName
                tnode.Cells(COLUMN_NAME).Value = rowTATCONV1.Item(COLUMN_NAME)
            Next
            Load_tvwTATCONV1(rowTATCONV1.Item("CONV_NO"), tnode)
        Next
    End Sub

    Private Sub tvwTATCONV1_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvwTATCONV1.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then
            tvwTATCONV1.PerformAction(UltraWinTree.UltraTreeAction.SelectActiveNode, False, False)
        End If
    End Sub

    Private Sub tvwTATCONV1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvwTATCONV1.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim tnode As UltraWinTree.UltraTreeNode = tvwTATCONV1.GetNodeFromPoint(e.X, e.Y)
            If tnode IsNot Nothing Then
                tvwTATCONV1.ActiveNode = tnode
                tnode.Selected = True
            End If
        End If
    End Sub
#End Region

#Region "grdTATCONV1"

    Function Get_CONV_ATTACHMENTS(ByVal CONV_NO As String) As String
        ASCMAIN1.sql = "Select Count (*) from ASTATTA2 " _
        & " where TABLE_NAME = 'TATCONV1' and COLUMN_NAME = 'CONV_NO' AND NVL(ATTACHMENT_STATUS,'O') <> 'D' " _
        & " and CODE_VALUE = '" & CONV_NO & "'"
        Dim CONV_ATTACHMENTS As Int64 = Val(ASCDATA1.GetDataValue)
        If CONV_ATTACHMENTS = 0 Then
            Return ""
        Else
            Return CStr(CONV_ATTACHMENTS)
        End If
    End Function

    Private Sub grdTATCONV1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdTATCONV1.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("INIT_OPER").Value & "" <> ASCMAIN1.USER_ID Then
                e.Cancel = True
                MsgBox("Cannot Delete Logs Created by Others", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Exit For
            End If

            Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find _
                (New Object() {grow.Cells("CONV_NO").Value})

            If ENTITY.RESTRICTIONS IsNot Nothing Then
                If ENTITY.RESTRICTIONS.Contains("D") And rowTATCONV1.RowState <> DataRowState.Added Then
                    e.Cancel = True
                    MsgBox("Deletion of Conversations is Not Permitted at this time", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit For
                End If
            End If

            e.Cancel = True
            If MsgBox("Delete Conversation: " & grow.Cells("CONV_SUBJECT").Value & "?", MsgBoxStyle.YesNo, "Confirm Deletion") = MsgBoxResult.No Then
            Else
                Delete_Conversation(grow.Cells("CONV_NO").Value)
            End If

        Next

        e.DisplayPromptMsg = False

    End Sub
    Sub Delete_Conversation(ByVal cNo As String)
        Dim rowTATCONV1 As DataRow = dst.Tables("TATCONV1").Rows.Find(New Object() {cNo})

        rowTATCONV1.Item("CONV_STATUS") = "D"

        Update_Record_TDA("TATCONV1")

        Fill_Records("TATCONV1", New String() {_E.TABLE_NAME, _E.TABLE_KEY}, True)

    End Sub
    Private Sub grdTATCONV1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdTATCONV1.ClickCellButton
        If e.Cell.Column.Key = "CONV_ATTACHMENTS" Then
            Dim ENTITY As New Dropped_On_Entity
            ENTITY.TABLE_NAME = "TATCONV1"
            ENTITY.COLUMN_NAME = "CONV_NO"
            ENTITY.CODE_VALUE = e.Cell.Row.Cells("CONV_NO").Value
            Dim DESC_VALUE = "Log by " & e.Cell.Row.Cells("INIT_OPER").Value _
                & " " & Format(e.Cell.Row.Cells.Item("INIT_DATE").Value, "MM/dd/yyyy HH:mm") _
                & " (" & e.Cell.Row.Cells("CONV_SUBJECT").Value & ")"
            ENTITY.DESC_VALUE = DESC_VALUE
            ENTITY.ATTACHMENT_NOTES = ""

            Dim F As New ASFATTA1
            F.ENTITY = ENTITY
            F.ShowDialog()
            F.Dispose()

            grdTATCONV1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        End If
    End Sub

    Private Sub grdTATCONV1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdTATCONV1.DoubleClickRow

    End Sub

    Private Sub grdTATCONV1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdTATCONV1.InitializeRow
        Dim CONV_STATUS As String = e.Row.Cells("CONV_STATUS").Value & ""
        If CONV_STATUS = "1" Then
            e.Row.Cells("CONV_FOLLOWUP_BY").Appearance.BackColor = Drawing.Color.LightGreen
            e.Row.Cells("CONV_FOLLOWUP_DATE").Appearance.BackColor = Drawing.Color.LightGreen
        End If

        Dim CONV_NO As String = ""
        Dim CONV_ATTACHMENTS As String = Get_CONV_ATTACHMENTS(e.Row.Cells("CONV_NO").Value)
        If CONV_ATTACHMENTS <> e.Row.Cells("CONV_ATTACHMENTS").Value & "" Then
            e.Row.Cells("CONV_ATTACHMENTS").Value = CONV_ATTACHMENTS
            grdTATCONV1.UpdateData()
        End If

    End Sub
#End Region

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabTATCONV1.SelectedTabChanged
        If tabTATCONV1.SelectedTab.Key = "Tree View" Then
            Setup_tvwTATCONV1()
        End If
    End Sub
End Class