Public Class ASFMENU1
    Dim MENU_ITEM_ctr As Long
    Dim strMENU_ID As String
    Dim strMENU_ITEM_TYPE As String
    Dim strMENU_ITEM_OBJECT As String
    Dim KEY_PREFIX As String
    Dim rowASTMENU1 As DataRow
    Dim row_Node As UltraWinTree.UltraTreeNode
    Dim anode As UltraWinTree.UltraTreeNode

    Dim tblASTMENUZ As New DataTable
    Dim tblASTMENUY As New DataTable

    Private WithEvents UltraTree_DropHightLight_DrawFilter As New UltraTree_DropHightLight_DrawFilter_Class()
    Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Tree\"
    Dim IMAGES As New Dictionary(Of String, Byte())
    Dim SQLs As New Dictionary(Of String, List(Of String))

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst

            Create_TDA(.Tables.Add, "TATALRT1", "*")


            Create_TDA(.Tables.Add, "ASTMENU1", "*", 0)

            ASCMAIN1.sql = "Select ASTMENU1.MENU_ID, ASTMENU1.MENU_ITEM_TYPE, ASTMENU1.MENU_ITEM_OBJECT" & vbCrLf _
                & ", ASTMENU1.MENU_ITEM_SECURITY, ASTMENU1.MENU_ITEM_SECURITY MENU_ITEM_SECURITY_ORIG from ASTMENU1"
            Create_TDA(.Tables.Add, "ASTMENUX", "**", 0, False,, 3)
            .Tables("ASTMENUX").Columns("MENU_ITEM_SECURITY_ORIG").MaxLength = -1
            .Tables("ASTMENUX").Columns("MENU_ITEM_SECURITY").MaxLength = -1

            ASCMAIN1.sql = "Select ASTSECM1.*, '0' SEL from ASTSECM1"
            Create_TDA(.Tables.Add, "ASTSECM1", "**", 0)
            .Tables("ASTSECM1").Columns("SEL").ReadOnly = False

            Create_TDA(.Tables.Add, "ASTMTKC1", "*", 0)

            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                With dst.Tables.Add("ASTUSERS")
                    .Columns.Add("USER_ID")
                    .Columns.Add("USER_STATUS")
                    .Columns.Add("SECURITY_CODES")
                    .PrimaryKey = New DataColumn() { .Columns("USER_ID")}
                End With

                ASCMAIN1.sql = "Select * from ASTUSER2"
                For Each rowASTUSER2 As DataRow In ASCDATA1.GetDataTable.Rows
                    Dim USER_ID As String = rowASTUSER2.Item("USER_ID")
                    Dim SECURITY_CODE As String = rowASTUSER2.Item("SECURITY_CODE")
                    Dim rowASTUSERS As DataRow = dst.Tables("ASTUSERS").Rows.Find(USER_ID)
                    If rowASTUSERS Is Nothing Then
                        rowASTUSERS = dst.Tables("ASTUSERS").NewRow
                        rowASTUSERS.Item("USER_ID") = USER_ID
                        rowASTUSERS.Item("SECURITY_CODES") = SECURITY_CODE
                        dst.Tables("ASTUSERS").Rows.Add(rowASTUSERS)
                    Else
                        rowASTUSERS.Item("SECURITY_CODES") &= "," & SECURITY_CODE
                    End If
                Next
            Else
                ASCMAIN1.sql = "Select USER_ID," & vbCrLf _
                & "ltrim(sys_connect_by_path(SECURITY_CODE,','),',') SECURITY_CODES" & vbCrLf _
                & "  from" & vbCrLf _
                & " (select USER_ID, SECURITY_CODE," & vbCrLf _
                & "      row_number() over(partition by USER_ID order by SECURITY_CODE) rn," & vbCrLf _
                & "    row_number() over(partition by USER_ID order by SECURITY_CODE desc)" & vbCrLf _
                & "        rn_desc " & vbCrLf _
                & " FROM ASTUSER2)" & vbCrLf _
                & "         Where rn_desc = 1" & vbCrLf _
                & "  start with rn = 1" & vbCrLf _
                & " connect by prior USER_ID = USER_ID" & vbCrLf _
                & " and prior rn = rn-1"
                ASCMAIN1.sql = "Select ASTUSER1.USER_ID, ASTUSER1.USER_STATUS, X.SECURITY_CODES " & vbCrLf _
                & " from ASTUSER1,(" & ASCMAIN1.sql & ") X where ASTUSER1.USER_ID = X.USER_ID (+)"
                Create_TDA(.Tables.Add, "ASTUSERS", "**", 0, False, "", 1)
                Fill_Records("ASTUSERS")
            End If

            ASCMAIN1.sql = "Select USER_ID, COUNT (*) USES" _
                & ", MIN (INIT_DATE) DATE1, MAX (INIT_DATE) DATE2" _
                & ", MAX(MENU_ID) MENU1, MIN (MENU_ID) MENU2" _
                & " from ASTOPST1 where MENU_ITEM_OBJECT = :PARM1 " _
                & " group by USER_ID"
            Create_TDA(.Tables.Add, "ASTOPSTX", "**", 0, False, "V", 1)
            .Tables("ASTOPSTX").Columns("USES").DataType = GetType(System.Int64)
        End With

        grdASTOPSTX.DataSource = dst.Tables("ASTOPSTX")

        grdASTMTKC1.DataSource = dst.Tables("ASTMTKC1")
        grdASTMTKC1.Visible = False

        grdASTSECM1.DataSource = dst.Tables("ASTSECM1")
        grdASTSECM1.Visible = False

        'Fill_Records("ASTUSERS")
        Fill_Records("ASTMENU1")
        tblASTMENUY = dst.Tables("ASTMENU1").Copy

        Fill_Records("ASTSECM1")
        Sort_grdColumns(grdASTSECM1, "SECURITY_CODE")
        Fill_Records("ASTMTKC1")

        Call Set_tvw_Columns()
        Call Build_Menu()

        grdASTMENUU.Parent = tvwASTMENU1.Parent
        grdASTMENUU.Dock = DockStyle.Fill
        grdASTMENUU.Visible = False

        splMenu.Panel2Collapsed = True

        dst.Tables("ASTMENU1").Rows.Clear()
        Write_ASTMENU1_from_Tree(tvwASTMENU1.Nodes) ' INITIALIZE SECURITY
        For Each ROW As DataRow In dst.Tables("ASTMENUX").Select("")
            ROW.Item("MENU_ITEM_SECURITY_ORIG") = ROW.Item("MENU_ITEM_SECURITY")
        Next

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        EMsg = ""
        Select Case eItemKey
            Case "Save"
                tvwASTMENU1.Tag = ""

                dst.Tables("ASTMENU1").Rows.Clear()

                For Each ROW As DataRow In dst.Tables("ASTMENUX").Select("")
                    ROW.Item("MENU_ITEM_SECURITY") = ""
                Next

                Write_ASTMENU1_from_Tree(tvwASTMENU1.Nodes)

                If tvwASTMENU1.Tag <> "" Then
                    EMsg = tvwASTMENU1.Tag
                Else
                    Dim DVW As DataView = New DataView(dst.Tables("ASTMENUX"))
                    DVW.RowFilter = "ISNULL(MENU_ITEM_SECURITY,'?') <> ISNULL(MENU_ITEM_SECURITY_ORIG,'?')"

                    Dim TBL As DataTable = DVW.ToTable
                    With New ASFMSGBF
                        .Show_grd(TBL, Me, "Menu Items with Security Code Changes")
                    End With
                    Dim rowASTPARMP As DataRow = LookUp("ASTPARMP", "Z")
                    Dim AS_PARM_SEC_ALERT_EMAIL As String = rowASTPARMP.Item("AS_PARM_SEC_ALERT_EMAIL") & ""
                    If AS_PARM_SEC_ALERT_EMAIL <> "" Then
                        Dim ALERT_MESSAGE As String = "Menu Security Assignments have Changed"
                        For Each ROW As DataRow In TBL.Select("")
                            ALERT_MESSAGE &= vbCrLf & ROW.Item("MENU_ID") & ":" & ROW.Item("MENU_ITEM_TYPE") & ":" & ROW.Item("MENU_ITEM_OBJECT") & " from " & ROW.Item("MENU_ITEM_SECURITY_ORIG") & " to " & ROW.Item("MENU_ITEM_SECURITY")
                        Next

                        If ASCMAIN1.CLIENT = "INT" And ASCMAIN1.DBS_SERVER = "INT" Then
                            TAC.TACMAIN1.Send_alert_email(Me, AS_PARM_SEC_ALERT_EMAIL, "", "MENU", ALERT_MESSAGE)
                        End If
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Save"


                BeginTrans()
                Me.Cursor = Cursors.WaitCursor

                ' HERE IS THE METHOD WE USED TO RECORD AUDIT TRAIL FOR ASTUSER2 - WE NEED TO MODIFY THIS FOR ASTMENU1
                ' HAVE A COPY CALLED ASTUSERY - THIS WOULD BE THE ORIGINAL VALUES IN ASTMENU1
                ' SAVE ASTMENU1 TO ASTMENUZ
                ' PUT ASTUSERY INTO ASTMENU1
                ' RIP THRU ASTUSERZ AND MAKE CHANGES NECESSARY TO ASTMENU1
                ' WRITE AUDIT TRAIL FOR ASTMENU1
                ' PUT ASTUSERZ BACK INTO ASTMENU1
                ' PUT NEW ASTMENU1 INTO ASTUSERY

                tblASTMENUZ = dst.Tables("ASTMENU1").Copy
                dst.Tables("ASTMENU1").Rows.Clear()
                dst.Tables("ASTMENU1").Merge(tblASTMENUY)
                dst.Tables("ASTMENU1").AcceptChanges()

                ' NEED TO KEEP TRACK OF DELETES
                For Each row As DataRow In tblASTMENUZ.Select("")
                    Dim MENU_ID As String = row.Item("MENU_ID")
                    Dim MENU_ITEM_TYPE As String = row.Item("MENU_ITEM_TYPE")
                    Dim MENU_ITEM_OBJECT As String = row.Item("MENU_ITEM_OBJECT")
                    Dim rowASTMENU1 As DataRow = dst.Tables("ASTMENU1").Rows.Find(New String() {MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT})
                    If rowASTMENU1 Is Nothing Then
                        rowASTMENU1 = dst.Tables("ASTMENU1").NewRow
                        rowASTMENU1.Item("MENU_ID") = MENU_ID
                        rowASTMENU1.Item("MENU_ITEM_TYPE") = MENU_ITEM_TYPE
                        rowASTMENU1.Item("MENU_ITEM_OBJECT") = MENU_ITEM_OBJECT
                        dst.Tables("ASTMENU1").Rows.Add(rowASTMENU1)
                    End If
                    For Each dcol As DataColumn In tblASTMENUZ.Columns
                        If row.Item(dcol.ColumnName) & "" <> rowASTMENU1.Item(dcol.ColumnName) & "" Then
                            rowASTMENU1.Item(dcol.ColumnName) = row.Item(dcol.ColumnName)
                        End If
                    Next
                Next

                For Each rowASTMENU1 As DataRow In dst.Tables("ASTMENU1").Select("", "", DataViewRowState.Unchanged)
                    Dim MENU_ID As String = rowASTMENU1.Item("MENU_ID")
                    Dim MENU_ITEM_TYPE As String = rowASTMENU1.Item("MENU_ITEM_TYPE")
                    Dim MENU_ITEM_OBJECT As String = rowASTMENU1.Item("MENU_ITEM_OBJECT")
                    Dim row As DataRow = tblASTMENUZ.Rows.Find(New String() {MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT})
                    If row Is Nothing Then
                        rowASTMENU1.Delete()
                    End If
                Next

                WriteAuditTrail("ASTMENU1")

                dst.Tables("ASTMENU1").Rows.Clear()
                dst.Tables("ASTMENU1").Merge(tblASTMENUZ)
                tblASTMENUY = dst.Tables("ASTMENU1").Copy
                dst.Tables("ASTMENU1").AcceptChanges()



                Update_Record_TDA("ASTMENU1", "Delete from ASTMENU1")
                Update_Record_TDA("ASTMTKC1", "Delete from ASTMTKC1")

                CommitTrans()

                For Each ROW As DataRow In dst.Tables("ASTMENUX").Select("")
                    ROW.Item("MENU_ITEM_SECURITY_ORIG") = ROW.Item("MENU_ITEM_SECURITY")
                Next



                MsgBox("Menu has been Saved", MsgBoxStyle.OkOnly, "Verification")
                Me.Cursor = Cursors.Default

                If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                    'TODO: the following block of code 1) needs to be coded for sql server, and 2) does not do anything because the 2 results sets being MINUSed are exactly the same
                    ' changes for sql server include using except instead of MINUS and using + to concatenate the fields with a field delimiter instead of using a multi-column IN clause
                    'ASCMAIN1.sql = "SELECT DISTINCT MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT FROM ASTMENU2 MINUS" _
                    '& " SELECT MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT FROM ASTMENU2 "
                    'Dim sql As String = "Select * from ASTMENU2 where (MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT) in (" & ASCMAIN1.sql & ")"
                    'Dim tbl As DataTable = ASCDATA1.GetDataTable(Sql)
                    'If tbl.Rows.Count <> 0 Then
                    '    Dim frm As New ASFMSGBF
                    '    frm.Show_grd(tbl, Me, "The following Menu Favorites no longer point to a valid Menu Node")
                    '    If MsgBox("OK to Delete these Orphan Favorites?", MsgBoxStyle.YesNo, "There are Menu Favorites which no longer point to a valid Menu Node") = vbYes Then
                    '        ASCMAIN1.sql = "Delete from ASTMENU2 where (MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT) in (" & ASCMAIN1.sql & ")"
                    '        ASCDATA1.ExecuteSQL()
                    '    End If
                    'End If
                Else
                    ASCMAIN1.sql = "SELECT DISTINCT MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT FROM ASTMENU2 MINUS" _
                        & " SELECT MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT FROM ASTMENU2 "
                    Dim tbl As DataTable = ASCDATA1.GetDataTable("Select * from ASTMENU2 where (MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT) in (" & ASCMAIN1.sql & ")")
                    If tbl.Rows.Count <> 0 Then
                        Dim frm As New ASFMSGBF
                        frm.Show_grd(tbl, Me, "The following Menu Favorites no longer point to a valid Menu Node")
                        If MsgBox("OK to Delete these Orphan Favorites?", MsgBoxStyle.YesNo, "There are Menu Favorites which no longer point to a valid Menu Node") = vbYes Then
                            ASCMAIN1.sql = "Delete from ASTMENU2 where (MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT) in (" & ASCMAIN1.sql & ")"
                            ASCDATA1.ExecuteSQL()
                        End If
                    End If
                End If

            Case "Users"
                Setup_grdASTMENUU()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            '.Groups("Screen Control").Items("Load Reports").Settings.Enabled = not_iScreenMode
            '.Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        'Call ASCMAIN1.Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
        End If

    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(tvwASTMENU1, "BBBBBBS", "Insert Menu", "Insert Form", "Insert Table", "Insert Report", "Delete Item", "Export to Excel", "Show Utilization")
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)


        '  Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing


        Select Case e.Tool.Key

            Case "Delete Item"
                tvwASTMENU1.ActiveNode.Remove()

            Case "Export to Excel"
                Setup_grdASTMENUU()

            Case "Show Utilization"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splMenu.Panel2Collapsed = Not tlb_sbt.Checked
                LOAD_ASTOPSTX()

            Case Else
                Dim KEY_PREFIX As String

                Dim MENU_ITEM_TYPE As String = Mid(e.Tool.Key, 8, 1)
                Dim MENU_ITEM_OBJECT As String

                If tvwASTMENU1.Nodes.Count = 0 Then
                    KEY_PREFIX = "MAIN" & Chr(0)
                    anode = tvwASTMENU1.Nodes.Add()
                ElseIf tvwASTMENU1.ActiveNode.IsRootLevelNode Then
                    KEY_PREFIX = "MAIN" & Chr(0)
                    anode = tvwASTMENU1.Nodes.Insert(tvwASTMENU1.ActiveNode.Index + 1)
                Else
                    KEY_PREFIX = tvwASTMENU1.ActiveNode.Parent.Key & Chr(0)
                    If Me.IsParentNode(tvwASTMENU1.ActiveNode) Then
                        anode = tvwASTMENU1.ActiveNode.Nodes.Insert(tvwASTMENU1.ActiveNode.Nodes.Count)
                    Else
                        anode = tvwASTMENU1.ActiveNode.Parent.Nodes.Insert(tvwASTMENU1.ActiveNode.Index + 1)
                    End If
                End If

                MENU_ITEM_ctr = MENU_ITEM_ctr + 1
                MENU_ITEM_OBJECT = ""
                anode.Key = Format(MENU_ITEM_ctr, "00000000")
                anode.Tag = "" & Chr(1) & MENU_ITEM_TYPE & Chr(1) & ""
                anode.Text = "{New Menu Item}"
                anode.Expanded = False
                If MENU_ITEM_TYPE = "M" Then
                    anode.Override.NodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "M")
                    anode.Override.ExpandedNodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, "M_OPEN")
                Else
                    anode.LeftImages.Add(ASCMAIN1.Get_Image(IMAGE_FOLDER, MENU_ITEM_TYPE))
                End If

                anode.Cells("MENU_ITEM_DESC").Value = "{New Menu Item}" ' anode.Text
                anode.Cells("MENU_ITEM_OBJECT").Value = MENU_ITEM_OBJECT
                anode.Cells("MENU_ITEM_TYPE").Value = MENU_ITEM_TYPE
                anode.Cells("MENU_ITEM_DESC").Appearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, MENU_ITEM_TYPE)
                ' UltraOptionSet1.Value = "G"

                Dim KEYS() As String = Split(KEY_PREFIX, Chr(0))

                tvwASTMENU1.SelectedNodes.Clear()
                tvwASTMENU1.ActiveNode = anode
                anode.BeginCellEdit(anode.Cells("MENU_ITEM_DESC").Column)


        End Select
    End Sub

#End Region

    Sub Set_tvw_Columns()
        With tvwASTMENU1

            .Appearances.Add("DropHighLightAppearance")
            With .Appearances("DropHighLightAppearance")
                .BackColor = System.Drawing.Color.Cyan
            End With

            .DrawFilter = UltraTree_DropHightLight_DrawFilter
            .Override.SelectionType = UltraWinTree.SelectType.ExtendedAutoDrag


            Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet

            Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add("MENU_ITEM_DESC")
            column.CanShowExpansionIndicator = DefaultableBoolean.True
            'column.CellWrapText = DefaultableBoolean.False
            'column.DataType = DataColumn.DataType
            'column.Editor = New EditorWithText()
            'column.Format = "00000"
            column.LayoutInfo.SpanX = 4
            column.Text = "Description"
            column.LayoutInfo.PreferredCellSize = New System.Drawing.Size(350, 19)
            column.LayoutInfo.PreferredLabelSize = New System.Drawing.Size(350, 0)

            column = rootColumnSet.Columns.Add("MENU_ITEM_OBJECT")
            column.Text = "Object"
            column.LayoutInfo.PreferredCellSize = New System.Drawing.Size(100, 19)
            column.LayoutInfo.PreferredLabelSize = New System.Drawing.Size(100, 0)
            column.MaxLength = dst.Tables("ASTMENU1").Columns("MENU_ITEM_OBJECT").MaxLength

            column = rootColumnSet.Columns.Add("MENU_ITEM_SECURITY")
            column.Text = "Security"
            column.LayoutInfo.PreferredCellSize = New System.Drawing.Size(150, 19)
            column.LayoutInfo.PreferredLabelSize = New System.Drawing.Size(150, 0)
            column.AllowCellEdit = UltraWinTree.AllowCellEdit.ReadOnly
            column.MaxLength = dst.Tables("ASTMENU1").Columns("MENU_ITEM_SECURITY").MaxLength

            column = rootColumnSet.Columns.Add("MENU_ITEM_PP")
            column.Text = "PP"
            column.MaxLength = dst.Tables("ASTMENU1").Columns("MENU_ITEM_PP").MaxLength

            column = rootColumnSet.Columns.Add("MENU_ITEM_PASSWORD")
            column.Text = "Password"
            column.MaxLength = dst.Tables("ASTMENU1").Columns("MENU_ITEM_PASSWORD").MaxLength

            column = rootColumnSet.Columns.Add("MENU_ITEM_HIDDEN")
            column.Text = "Hidden"
            column.DataType = GetType(Boolean)
            column.CellAppearance.TextHAlign = HAlign.Center
            column.HeaderAppearance.TextHAlign = HAlign.Center

            column = rootColumnSet.Columns.Add("MENU_ITEM_STANDALONE")
            column.Text = "Standalone"
            column.DataType = GetType(Boolean)
            column.CellAppearance.TextHAlign = HAlign.Center
            column.HeaderAppearance.TextHAlign = HAlign.Center

            column = rootColumnSet.Columns.Add("MENU_ITEM_STATUS")
            column.Text = "Status"
            column.LayoutInfo.PreferredCellSize = New System.Drawing.Size(100, 8)
            column.LayoutInfo.PreferredLabelSize = New System.Drawing.Size(100, 0)
            column.MaxLength = dst.Tables("ASTMENU1").Columns("MENU_ITEM_STATUS").MaxLength

            column = rootColumnSet.Columns.Add("MENU_ITEM_FORM")
            column.Text = "Form"
            column.LayoutInfo.PreferredCellSize = New System.Drawing.Size(100, 19)
            column.LayoutInfo.PreferredLabelSize = New System.Drawing.Size(100, 0)
            column.MaxLength = dst.Tables("ASTMENU1").Columns("MENU_ITEM_FORM").MaxLength

            column = rootColumnSet.Columns.Add("MENU_ITEM_EOM_CHECK")
            column.Text = "EOM"
            column.DataType = GetType(Boolean)
            column.CellAppearance.TextHAlign = HAlign.Center
            column.HeaderAppearance.TextHAlign = HAlign.Center

            column = rootColumnSet.Columns.Add("MENU_ITEM_TYPE")
            column.Text = "Type"
            column.Visible = False

            .ColumnSettings.RootColumnSet.AllowCellEdit = UltraWinTree.AllowCellEdit.Full
            .ColumnSettings.RootColumnSet.ShowSortIndicators = DefaultableBoolean.False
            .ColumnSettings.RootColumnSet.AllowSorting = DefaultableBoolean.False

            Call Setup_Menu_ViewStyle()
            .ColumnSettings.TabNavigation = UltraWinTree.TabNavigation.NextCell

            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_DESC").LayoutInfo.SpanX = 10
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_OBJECT").LayoutInfo.LabelSpan = 10
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_SECURITY").LayoutInfo.PreferredCellSize = New System.Drawing.Size(500, 0)
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_HIDDEN").EditorComponent = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
            .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_STANDALONE").EditorComponent = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
            ' .ColumnSettings.RootColumnSet.Columns("MENU_ITEM_STATUS").EditorControl = New Infragistics.Win.UltraWinEditors.UltraOptionSet

        End With
    End Sub

    Sub Build_Menu()
        Call ASCMAIN1.Add_Menu_to_Tree("MAIN", "M" & Chr(1) & "MAIN" & Chr(0), tvwASTMENU1, 0, dst.Tables("ASTMENU1"))
    End Sub

    Private Function IsParentNode(ByVal Node As UltraWinTree.UltraTreeNode) As Boolean
        Dim Tag As String
        Tag = Node.Tag
        Return Split(Tag, Chr(1))(1) = "M"
    End Function

    Private Function IsParentNodeSelected(ByVal Tree As UltraWinTree.UltraTree) As Boolean
        For Each SelectedNode As UltraWinTree.UltraTreeNode In Tree.SelectedNodes
            If Me.IsParentNode(SelectedNode) Then Return True
        Next
        Return False
    End Function

    Private Function IsAnyParentSelected(ByVal Node As UltraWinTree.UltraTreeNode) As Boolean
        Dim ParentNode As UltraWinTree.UltraTreeNode

        ParentNode = Node.Parent
        Do Until ParentNode Is Nothing
            If ParentNode.Selected Then Return True
            ParentNode = ParentNode.Parent
        Loop
        Return False
    End Function

    Private Sub UltraTree_DropHightLight_DrawFilter_Invalidate(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraTree_DropHightLight_DrawFilter.Invalidate
        tvwASTMENU1.Invalidate()
    End Sub

    Private Sub UltraTree_DropHightLight_DrawFilter_QueryStateAllowedForNode(ByVal sender As Object, ByVal e As UltraTree_DropHightLight_DrawFilter_Class.QueryStateAllowedForNodeEventArgs) Handles UltraTree_DropHightLight_DrawFilter.QueryStateAllowedForNode
        If Not IsParentNode(e.Node) Then
            e.StatesAllowed = DropLinePositionEnum.AboveNode Or DropLinePositionEnum.BelowNode
            UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 2
        Else
            If e.Node.Selected Then
                e.StatesAllowed = DropLinePositionEnum.AboveNode Or DropLinePositionEnum.BelowNode
                UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 2
            Else
                UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 3
            End If
        End If
    End Sub

    Private Sub UltraOptionSet1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraOptionSet1.ValueChanged
        Call Setup_Menu_ViewStyle()
    End Sub

    Sub Setup_Menu_ViewStyle()
        With tvwASTMENU1
            If UltraOptionSet1.Value = "T" Then
                .Override.CellClickAction = UltraWinTree.CellClickAction.Default
                .ViewStyle = UltraWinTree.ViewStyle.Standard
                .AllowDrop = True
                .Override.AllowCut = DefaultableBoolean.True
                .Override.AllowCopy = DefaultableBoolean.True
                .Override.AllowPaste = DefaultableBoolean.True
                .Override.AllowDelete = DefaultableBoolean.True
            Else
                .Override.CellClickAction = UltraWinTree.CellClickAction.EditCellSelectText '  .EditCell
                .ViewStyle = UltraWinTree.ViewStyle.OutlookExpress
                .AllowDrop = False
                .Override.AllowCut = DefaultableBoolean.False
                .Override.AllowCopy = DefaultableBoolean.False
                .Override.AllowPaste = DefaultableBoolean.False
                .Override.AllowDelete = DefaultableBoolean.False
            End If
        End With
    End Sub

    Sub Write_ASTMENU1_from_Tree(
                                ByVal tnodes As Infragistics.Win.UltraWinTree.TreeNodesCollection,
                                Optional MENU_ITEM_SECURITY_PARENT_NODE As String = "")

        Dim MENU_ITEM_SEQ As Integer = 0
        Dim MENU_ITEM As String = ""

        Try

            For Each n As Infragistics.Win.UltraWinTree.UltraTreeNode In tnodes
                MENU_ITEM = "?"
                Dim row As DataRow = dst.Tables("ASTMENU1").NewRow
                If tnodes.ParentNode Is Nothing Then
                    row.Item("MENU_ID") = "MAIN"
                Else
                    row.Item("MENU_ID") = tnodes.ParentNode.Cells("MENU_ITEM_OBJECT").Value
                End If
                row.Item("MENU_ITEM_TYPE") = n.Cells("MENU_ITEM_TYPE").Value
                row.Item("MENU_ITEM_OBJECT") = n.Cells("MENU_ITEM_OBJECT").Value
                MENU_ITEM_SEQ = MENU_ITEM_SEQ + 1
                row.Item("MENU_ITEM_SEQ") = MENU_ITEM_SEQ
                row.Item("MENU_ITEM_PP") = n.Cells("MENU_ITEM_PP").Value
                row.Item("MENU_ITEM_DESC") = n.Text
                row.Item("MENU_ITEM_PASSWORD") = n.Cells("MENU_ITEM_PASSWORD").Value
                row.Item("MENU_ITEM_SECURITY") = n.Cells("MENU_ITEM_SECURITY").Value
                row.Item("MENU_ITEM_HIDDEN") = IIf(n.Cells("MENU_ITEM_HIDDEN").Value, "1", "0")
                row.Item("MENU_ITEM_STANDALONE") = IIf(n.Cells("MENU_ITEM_STANDALONE").Value, "1", "0")
                row.Item("MENU_ITEM_STATUS") = n.Cells("MENU_ITEM_STATUS").Value
                row.Item("MENU_ITEM_FORM") = n.Cells("MENU_ITEM_FORM").Value
                row.Item("MENU_ITEM_EOM_CHECK") = IIf(n.Cells("MENU_ITEM_EOM_CHECK").Value, "1", "0")
                MENU_ITEM = IIf(row.Item("MENU_ID") = "", "{No Menu ID}", row.Item("MENU_ID")) & ":" _
                          & IIf(row.Item("MENU_ITEM_TYPE") = "", "{No Type}", row.Item("MENU_ITEM_TYPE")) & ":" _
                          & IIf(row.Item("MENU_ITEM_OBJECT") = "", "{No Type}", row.Item("MENU_ITEM_OBJECT")) & ":" _
                          & IIf(row.Item("MENU_ITEM_DESC") = "", "{No Type}", row.Item("MENU_ITEM_DESC"))
                dst.Tables("ASTMENU1").Rows.Add(row)

                If row.Item("MENU_ID") = "" _
                Or row.Item("MENU_ITEM_TYPE") = "" _
                Or row.Item("MENU_ITEM_OBJECT") = "" _
                Or row.Item("MENU_ITEM_DESC") = "" Then
                    tvwASTMENU1.Tag &= vbCr & "Error with " & MENU_ITEM
                Else
                    If Mid$(row.Item("MENU_ITEM_OBJECT"), 3, 1) = "F" And row.Item("MENU_ITEM_TYPE") <> "F" _
                    Or Mid$(row.Item("MENU_ITEM_OBJECT"), 3, 1) <> "F" And row.Item("MENU_ITEM_TYPE") = "F" Then
                        tvwASTMENU1.Tag &= vbCr & "Type Mis-Match with " & MENU_ITEM
                    End If
                    If Mid$(row.Item("MENU_ITEM_OBJECT"), 3, 1) = "T" And row.Item("MENU_ITEM_TYPE") <> "T" _
                    Or Mid$(row.Item("MENU_ITEM_OBJECT"), 3, 1) <> "T" And row.Item("MENU_ITEM_TYPE") = "T" Then
                        tvwASTMENU1.Tag &= vbCr & "Type Mis-Match with " & MENU_ITEM
                    End If
                    If Mid$(row.Item("MENU_ITEM_OBJECT"), 3, 1) = "X" And row.Item("MENU_ITEM_TYPE") <> "R" _
                    Or Mid$(row.Item("MENU_ITEM_OBJECT"), 3, 1) <> "X" And row.Item("MENU_ITEM_TYPE") = "R" Then
                        ' tvwASTMENU1.Tag &= vbCr & "Type Mis-Match with " & MENU_ITEM
                    End If
                End If

                Dim rowASTMENUX As DataRow = dst.Tables("ASTMENUX").Rows.Find(New Object() {
                                                                              row.Item("MENU_ID"),
                                                                              row.Item("MENU_ITEM_TYPE"),
                                                                              row.Item("MENU_ITEM_OBJECT")})
                If rowASTMENUX Is Nothing Then
                    rowASTMENUX = dst.Tables("ASTMENUX").NewRow
                    rowASTMENUX.Item("MENU_ID") = row.Item("MENU_ID")
                    rowASTMENUX.Item("MENU_ITEM_TYPE") = row.Item("MENU_ITEM_TYPE")
                    rowASTMENUX.Item("MENU_ITEM_OBJECT") = row.Item("MENU_ITEM_OBJECT")
                    dst.Tables("ASTMENUX").Rows.Add(rowASTMENUX)
                End If

                Dim MENU_ITEM_SECURITY As String = row.Item("MENU_ITEM_SECURITY") & ""
                rowASTMENUX.Item("MENU_ITEM_SECURITY") = MENU_ITEM_SECURITY_PARENT_NODE & MENU_ITEM_SECURITY

                If n.Nodes.Count > 0 Then
                    Write_ASTMENU1_from_Tree(n.Nodes, rowASTMENUX.Item("MENU_ITEM_SECURITY") & "")
                End If
            Next
        Catch ex As Exception
            tvwASTMENU1.Tag &= vbCr & "Error with " & MENU_ITEM & ": " & ex.Message
        End Try

    End Sub

    Private Sub grdASTSECM1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTSECM1.AfterRowUpdate
        Dim MENU_ITEM_SECURITY As String = ""
        For Each row As DataRow In dst.Tables("ASTSECM1").Select("SEL = '1'")
            MENU_ITEM_SECURITY = MENU_ITEM_SECURITY & row.Item("SECURITY_CODE")
        Next

        tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_SECURITY").Value = MENU_ITEM_SECURITY
    End Sub

    Private Sub grdASTSECM1_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTSECM1.CellChange
        grdASTSECM1.UpdateData()
    End Sub

    Private Sub grdASTSECM1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTSECM1.InitializeLayout
        grdASTSECM1.DisplayLayout.Bands(0).Columns("SEL").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        grdASTSECM1.DisplayLayout.Bands(0).Columns("SEL").Editor.DataFilter = New CheckEditorDataFilter
    End Sub

    Private Sub tvwASTMENU1_AfterActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.NodeEventArgs) Handles tvwASTMENU1.AfterActivate

        If dst.Tables("ASTSECM1") Is Nothing Then
            Exit Sub
        End If

        Dim sqlMENU_ITEM_SECURITY As String = ASCMAIN1.Select_List(tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_SECURITY").Text, 2)

        For Each row As DataRow In dst.Tables("ASTSECM1").Rows
            If InStr(sqlMENU_ITEM_SECURITY, row.Item("SECURITY_CODE")) <> 0 Then
                row.Item("SEL") = "1"
            Else
                row.Item("SEL") = "0"
            End If
        Next

        grdASTSECM1.Visible = True

        Dim dvw As DataView = DirectCast(grdASTMTKC1.DataSource, DataTable).DefaultView
        dvw.RowFilter = String.Format("MENU_ITEM_TYPE = '{0}' AND MENU_ITEM_OBJECT = '{1}'", _
                                      tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_TYPE").Text, _
                                      tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_OBJECT").Text)
        grdASTMTKC1.Visible = True

        Absx1.txtFor("MENU_ITEM_TYPE").Text = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_TYPE").Text
        Absx1.txtFor("MENU_ITEM_OBJECT").Text = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_OBJECT").Text
        Absx1.txtFor("MENU_ITEM_PP").Text = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_PP").Text
        Absx1.txtFor("MENU_ITEM_PASSWORD").Text = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_PASSWORD").Text
        Absx1.chkFor("MENU_ITEM_HIDDEN").Checked = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_HIDDEN").Value
        Absx1.chkFor("MENU_ITEM_STANDALONE").Checked = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_STANDALONE").Value
        Absx1.optFor("MENU_ITEM_STATUS").Value = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_STATUS").Value
        Absx1.txtFor("MENU_ITEM_FORM").Text = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_FORM").Text
        Absx1.chkFor("MENU_ITEM_EOM_CHECK").Checked = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_EOM_CHECK").Value

        LOAD_ASTOPSTX()
    End Sub

    Sub Load_ASTOPSTX()
        If Not splMenu.Panel2Collapsed Then
            Fill_Records("ASTOPSTX", tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_OBJECT").Text)
            grdASTOPSTX.Text = "Ultilization for " & tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_OBJECT").Text & " - " & tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_DESC").Text
            Sort_grdColumns(grdASTOPSTX, "DATE2".ToLower)
        End If
    End Sub

    Private Sub tvwASTMENU1_AfterCellExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.AfterCellExitEditModeEventArgs) Handles tvwASTMENU1.AfterCellExitEditMode
        If e.Cell.Column.Key = "MENU_ITEM_DESC" Then
            e.Node.Text = e.Cell.Text
        End If
    End Sub

    Private Sub tvwASTMENU1_AfterPaste(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.AfterPasteEventArgs) Handles tvwASTMENU1.AfterPaste
        Dim i As Integer
        With e.Nodes(0)
            For i = 0 To .Cells.Count - 1
                .Cells(i).Value = row_Node.Cells(i).Value
            Next
            strMENU_ITEM_TYPE = Split(.Tag, Chr(1))(1)
            .Cells("MENU_ITEM_DESC").Appearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, strMENU_ITEM_TYPE)
        End With
    End Sub

    Private Sub tvwASTMENU1_AfterSelect(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.SelectEventArgs) Handles tvwASTMENU1.AfterSelect
        If e.NewSelections.Count = 1 Then
            anode = e.NewSelections(0)
        End If
    End Sub

    Private Sub tvwASTMENU1_BeforeCopy(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.BeforeCopyEventArgs) Handles tvwASTMENU1.BeforeCopy
        row_Node = e.Nodes(0)
    End Sub

    Private Sub tvwASTMENU1_BeforePaste(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.BeforePasteEventArgs) Handles tvwASTMENU1.BeforePaste
        If tvwASTMENU1.SelectedNodes.Count = 1 Then
            If Split(tvwASTMENU1.SelectedNodes(0).Tag, Chr(1))(1) <> "M" Then
                MsgBox("Cannot Paste a Menu Item into another non-Folder Menu Item" & vbCr & "Try selecting a Folder first, and then Paste", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
            Else
                MENU_ITEM_ctr = MENU_ITEM_ctr + 1
                e.Nodes(0).Key = Format(MENU_ITEM_ctr, "00000000")
            End If
        End If
    End Sub

    Private Sub tvwASTMENU1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwASTMENU1.DragDrop
        Dim Node As UltraWinTree.UltraTreeNode
        Dim SelectedNodes As UltraWinTree.SelectedNodesCollection
        Dim DropNode As UltraWinTree.UltraTreeNode
        Dim i As Integer

        DropNode = UltraTree_DropHightLight_DrawFilter.DropHightLightNode

        SelectedNodes = e.Data.GetData(GetType(UltraWinTree.SelectedNodesCollection))
        SelectedNodes = SelectedNodes.Clone()

        SelectedNodes.SortByPosition()

        Select Case UltraTree_DropHightLight_DrawFilter.DropLinePosition
            Case DropLinePositionEnum.OnNode
                For i = 0 To SelectedNodes.Count - 1
                    Node = SelectedNodes(i)
                    Node.Reposition(DropNode.Nodes)
                Next
            Case DropLinePositionEnum.BelowNode
                For i = 0 To SelectedNodes.Count - 1
                    Node = SelectedNodes(i)
                    Node.Reposition(DropNode, UltraWinTree.NodePosition.Next)
                    DropNode = Node
                Next
            Case DropLinePositionEnum.AboveNode
                For i = 0 To SelectedNodes.Count - 1
                    Node = SelectedNodes(i)
                    Node.Reposition(DropNode, UltraWinTree.NodePosition.Previous)
                Next
        End Select

        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
    End Sub

    Private Sub tvwASTMENU1_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwASTMENU1.DragLeave
        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
    End Sub

    Private Sub tvwASTMENU1_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwASTMENU1.DragOver
        Dim Node As UltraWinTree.UltraTreeNode
        Dim PointInTree As System.Drawing.Point

        With tvwASTMENU1
            PointInTree = .PointToClient(New System.Drawing.Point(e.X, e.Y))

            Node = .GetNodeFromPoint(PointInTree)

            If Node Is Nothing Then
                e.Effect = DragDropEffects.None
                UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
                Return
            End If

            ' THE NEXT FEW LINES PREVENT MOVING A NEW FOLDER INTO AN ALSO NEW (MAIN) FOLDER, SO WE DISBALED IT.  WHAT USEFUL PURPOSE DID THIS CODE EVER SERVE?
            'If Me.IsParentNode(Node) And Me.IsParentNodeSelected(Me.tvwASTMENU1) Then
            '    If PointInTree.Y > (Node.Bounds.Top + 2) AndAlso PointInTree.Y < (Node.Bounds.Bottom - 2) Then
            '        e.Effect = DragDropEffects.None
            '        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
            '        Return
            '    End If
            'End If

            'If IsAnyParentSelected(Node) Then
            '    e.Effect = DragDropEffects.None
            '    UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
            '    Return
            'End If

            UltraTree_DropHightLight_DrawFilter.SetDropHighlightNode(Node, PointInTree)
            e.Effect = DragDropEffects.Move
        End With
    End Sub

    Private Sub tvwASTMENU1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvwASTMENU1.MouseUp
        tvwASTMENU1.SelectedNodes.Clear()
        Dim anode As Infragistics.Win.UltraWinTree.UltraTreeNode = tvwASTMENU1.GetNodeFromPoint(e.Location)
        If anode IsNot Nothing Then
            anode.Selected = True
            tvwASTMENU1.ActiveNode = anode
        End If
    End Sub

    Private Sub tvwASTMENU1_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles tvwASTMENU1.QueryContinueDrag
        If e.EscapePressed Then
            e.Action = DragAction.Cancel
            UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
        End If
    End Sub

    Private Sub tvwASTMENU1_SelectionDragStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwASTMENU1.SelectionDragStart
        tvwASTMENU1.DoDragDrop(tvwASTMENU1.SelectedNodes, DragDropEffects.Move)
    End Sub

    Private Sub grdASTMTKC1_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTMTKC1.AfterRowInsert
        e.Row.Cells("MENU_ITEM_TYPE").Value = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_TYPE").Text
        e.Row.Cells("MENU_ITEM_OBJECT").Value = tvwASTMENU1.ActiveNode.Cells("MENU_ITEM_OBJECT").Text
    End Sub

    Overrides Sub Leaving_txt_Special_Before(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        tvwASTMENU1.ActiveNode.Cells(COLUMN_NAME).Value = ctl.Text
    End Sub

    Overrides Sub CheckedChanged_Special(ByVal COLUMN_NAME As String, ByVal chk As UltraWinEditors.UltraCheckEditor)
        tvwASTMENU1.ActiveNode.Cells(COLUMN_NAME).Value = chk.Checked
    End Sub

    Public Overrides Sub opt_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If COLUMN_NAME <> "" Then
            Dim opt As UltraWinEditors.UltraOptionSet = DirectCast(sender, UltraWinEditors.UltraOptionSet)
            tvwASTMENU1.ActiveNode.Cells(COLUMN_NAME).Value = opt.Value
        End If
    End Sub
    Sub Setup_grdASTMENUU()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Exporting Menu / User Matrix to Excel")

        If dst.Tables.Contains("ASTMENUU") Then
            dst.Tables("ASTMENUU").Rows.Clear()
        Else
            With dst.Tables.Add("ASTMENUU")
                .Columns.Add("RECORD_NO", GetType(System.Int32))
                .Columns.Add("LEVEL", GetType(System.Int32))
                '.Columns.Add("MENU_ITEM_IMAGE", GetType(System.Byte()))
                .Columns.Add("MENU_ID")
                .Columns.Add("MENU_ITEM_TYPE")
                .Columns.Add("MENU_ITEM_OBJECT")
                .Columns.Add("MENU_ITEM_DESC")
                .Columns.Add("MENU_ITEM_SECURITY")
                .Columns.Add("MENU_ITEM_STATUS")
                .Columns.Add("USER_COUNT", GetType(System.Int32))
                .Columns.Add("USERS")
                .Columns.Add("SEC0")
                .Columns.Add("SEC1")
                .Columns.Add("SEC2")
                .Columns.Add("SEC3")
                .Columns.Add("SEC4")
                .Columns.Add("SEC5")

                For Each rowASTUSER1 As DataRow In dst.Tables("ASTUSERS").Select("USER_STATUS = 'A'", "USER_ID")
                    .Columns.Add(rowASTUSER1.Item("USER_ID"))
                Next

                grdASTMENUU.DataSource = dst.Tables("ASTMENUU")

            End With
        End If

        SQLs.Clear()

        For Each MENU_ITEM_TYPE As String In New String() {"M", "F", "R", "T"}
            Dim aBitMap As System.Drawing.Bitmap = ASCMAIN1.Get_Image(IMAGE_FOLDER, MENU_ITEM_TYPE)
            ' IMAGES.Add(MENU_ITEM_TYPE, aBitMap)
        Next

        Dim SECs() As String
        ReDim SECs(5)
        Dim RECORD_NO As Int32 = 0
        For Each n As UltraWinTree.UltraTreeNode In tvwASTMENU1.Nodes
            Load_Node(n, 0, SECs, RECORD_NO)
        Next


        grdASTMENUU.Text = "Menu / User Matrix (Active Users only)"
        grdASTMENUU.Visible = True

        Gembox_Excel_Export_grd(grdASTMENUU)
        'Export_to_Excel(grdASTMENUU)

        grdASTMENUU.Visible = False

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Shared Function ConvertToByteArray(ByVal value As System.Drawing.Bitmap) As Byte()

        Dim bitmapBytes As Byte()
        Using stream As New System.IO.MemoryStream
            value.Save(stream, value.RawFormat)
            bitmapBytes = stream.ToArray
        End Using
        Return bitmapBytes
    End Function

    Sub Load_Node(ByVal n As UltraWinTree.UltraTreeNode, _
                  ByVal LEVEL As Integer, _
                  ByVal SECs() As String, ByRef RECORD_NO As Int32)

        If n.Cells("MENU_ITEM_HIDDEN").Value Then
            'Stop
        Else

            SECs(LEVEL) = n.Cells("MENU_ITEM_SECURITY").Text

            Dim SPACES As String = ""
            SPACES = SPACES.PadRight(3 * LEVEL)

            Dim rowASTMENUU As DataRow = dst.Tables("ASTMENUU").NewRow
            RECORD_NO += 1
            rowASTMENUU.Item("RECORD_NO") = RECORD_NO
            rowASTMENUU.Item("LEVEL") = LEVEL
            If n.Parent Is Nothing Then
                rowASTMENUU.Item("MENU_ID") = "MAIN"
            Else
                rowASTMENUU.Item("MENU_ID") = n.Parent.Cells("MENU_ITEM_OBJECT").Value
            End If

            'rowASTMENUU.Item("MENU_ITEM_IMAGE") = ASCMAIN1.Get_Image(IMAGE_FOLDER, n.Cells("MENU_ITEM_TYPE").Text) ' IMAGES(n.Cells("MENU_ITEM_TYPE").Text)
            rowASTMENUU.Item("MENU_ITEM_TYPE") = n.Cells("MENU_ITEM_TYPE").Text
            rowASTMENUU.Item("MENU_ITEM_OBJECT") = n.Cells("MENU_ITEM_OBJECT").Text
            rowASTMENUU.Item("MENU_ITEM_DESC") = SPACES & n.Cells("MENU_ITEM_DESC").Text
            rowASTMENUU.Item("MENU_ITEM_SECURITY") = n.Cells("MENU_ITEM_SECURITY").Text
            rowASTMENUU.Item("MENU_ITEM_STATUS") = n.Cells("MENU_ITEM_STATUS").Text
            'If n.Cells("MENU_ITEM_DESC").Text = "Advertising & Sales Promotion" Then Stop
            Dim sql As String = ""
            For s As Integer = 0 To LEVEL
                rowASTMENUU.Item("SEC" & Format(s, "0")) = SECs(s)
                If SECs(s) <> "" Then
                    Dim SECURITY_CODEs As String = ""
                    For ss As Integer = 1 To Len(SECs(s)) / 2
                        Dim SECURITY_CODE As String = Mid(SECs(s), (ss - 1) * 2 + 1, 2)
                        SECURITY_CODEs &= ",'" & SECURITY_CODE & "'"
                    Next
                    sql &= " Intersect Select Distinct USER_ID from ASTUSER2 where SECURITY_CODE in (" & Mid(SECURITY_CODEs, 2) & ")"
                End If
            Next

            If sql = "" Then
                sql = "Select USER_ID from ASTUSER1 where USER_STATUS = 'A'"
            Else
                sql = "Select USER_ID from ASTUSER1 where USER_STATUS = 'A' and USER_ID in (" & Mid(sql, 12) & ")"
            End If

            Dim USER_IDs As New List(Of String)
            If SQLs.ContainsKey(sql) Then
                USER_IDs = SQLs(sql)
            Else
                ASCMAIN1.sql = sql
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "USER_ID")
                    USER_IDs.Add(row.Item("USER_ID"))
                Next
                SQLs.Add(sql, USER_IDs)
            End If

            Dim USERS As String = ""
            'USER_IDs.Clear()
            For Each USER_ID As String In USER_IDs
                Try
                    rowASTMENUU.Item(USER_ID) = "x"
                    USERS &= "," & USER_ID
                Catch ex As Exception

                End Try
            Next

            rowASTMENUU.Item("USER_COUNT") = USER_IDs.Count
            rowASTMENUU.Item("USERS") = Mid(USERS, 2)

            dst.Tables("ASTMENUU").Rows.Add(rowASTMENUU)

            If n.HasNodes Then
                For Each CN As UltraWinTree.UltraTreeNode In n.Nodes
                    Load_Node(CN, LEVEL + 1, SECs, RECORD_NO)
                Next
            End If
        End If

    End Sub

    Private Sub grdASTMENUU_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTMENUU.InitializeRow
    End Sub
End Class