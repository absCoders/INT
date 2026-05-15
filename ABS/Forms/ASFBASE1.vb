Public Class ASFBASE1

    Public Group_Items_Prohibited As New List(Of String)

    Private Sub ASFBASE1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If ASCMAIN1.USER_ID = "" Then
            Exit Sub
        End If

        Call ASCMAIN1.Activate_Form(Me)
        Me.UltraDockManager1.PaneFromControl(UltraExplorerBar1).Activate()
        Me.UltraDockManager1.Visible = True

        ASFMAIN1.UltraStatusBar1.Panels("MSG1").Text = ""
        ASFMAIN1.UltraStatusBar1.Panels("MSG2").Text = ""

        Call Set_MODE_in_StatusBar1()

        'Dim seconds As Int32 = DateDiff(DateInterval.Second, ASCMAIN1.Timer, Now)
        'ASFMAIN1.UltraStatusBar1.Panels(7).Text = CStr(seconds)

    End Sub

    Protected Overrides Sub OnDeactivate(ByVal e As System.EventArgs)
        MyBase.OnDeactivate(e)

        If (Me.UltraDockManager1.ControlPanes.Count = 2) Then
            If Me.UltraDockManager1.PaneFromControl(UltraExplorerBar1).DockedState = UltraWinDock.DockedState.Floating Then
                Me.UltraDockManager1.PaneFromControl(UltraExplorerBar1).Dock(True)
                'Me.UltraDockManager1.PaneFromControl(UltraExplorerBar1).ParentDocked.ChildPaneStyle = UltraWinDock.ChildPaneStyle.TabGroup
                'Me.UltraDockManager1.DockControls(New Control() {UltraExplorerBar1}, UltraWinDock.DockedLocation.DockedRight, UltraWinDock.ChildPaneStyle.TabGroup)
            End If

        End If

    End Sub

    Private Sub ASFBASE1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool
        tlb_pop = DirectCast(ASFMAIN1.UltraToolbarsManager1.Tools("Help"), UltraWinToolbars.PopupMenuTool)

        'If tlb_pop.Tools.Contains(CStr(SELECTION_NO)) AndAlso tlb_pop.Tools.Contains(tlb_pop.Tools(CStr(SELECTION_NO))) Then

        If tlb_pop.Tools.Contains(CStr(SELECTION_NO)) AndAlso tlb_pop.Tools.Contains(tlb_pop.Tools(CStr(SELECTION_NO))) Then
            Try
                tlb_pop.Tools.Remove(tlb_pop.Tools(CStr(SELECTION_NO)))
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub ASFBASE1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        If e.CloseReason = CloseReason.MdiFormClosing Then
            For Each frm As System.Windows.Forms.Form In ASCMAIN1.ABS_FORMS
                If TypeOf (frm) Is ASFBASE1 Then
                    Dim ABSF As ASFBASE1 = DirectCast(frm, ASFBASE1)
                    If ABSF IsNot Nothing AndAlso ABSF.Name <> "ASFMAIN1" Then
                        Dim form_is_active As Boolean = ABSF.ScreenMode
                        If ABSF.ScreenMode Then
                            If ABSF.UltraExplorerBar1 IsNot Nothing AndAlso ABSF.UltraExplorerBar1.Groups.Exists("Screen Control") AndAlso ABSF.UltraExplorerBar1.Groups("Screen Control").Items.Exists("Done") Then
                                If ABSF.UltraExplorerBar1.Groups("Screen Control").Items("Done").Visible And _
                                   ABSF.UltraExplorerBar1.Groups("Screen Control").Items("Done").SettingsResolved.Enabled Then
                                    form_is_active = False
                                End If
                            End If
                        End If
                        If form_is_active Then
                            MsgBox("Form Still Active: " & ABSF.MENU_ITEM_DESC, MsgBoxStyle.OkOnly, "Cannot Close All Forms")
                            e.Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            Next
        End If

        ASFMAIN1.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Closing")
        'Application.DoEvents()

        If ScreenMode Then
            Dim GroupKey As String = "Screen Control"
            If MENU_ITEM_TYPE = "R" Then
                GroupKey = "Update Controls"
            End If

            If UltraExplorerBar1.Groups.Exists(GroupKey) Then
                If UltraExplorerBar1.Groups(GroupKey).Items.Exists("Done") Then
                    If UltraExplorerBar1.Groups(GroupKey).Items("Done").Settings.Enabled = DefaultableBoolean.True And _
                           UltraExplorerBar1.Groups(GroupKey).Items("Done").Visible = True _
                    Then
                        'Proceed_PreReq("Done")
                        IsClosing = True
                        Click_Command("Done")
                        IsClosing = False
                    End If
                End If
            End If

            ' note that this does not capture the DATETIME_STAMP (see ClickCommand)
            '  - but now it does since I changed it to use Click_Command instead of Proceed_PreReq
        End If


        e.Cancel = ASCMAIN1.Reset(Me)
        If Not e.Cancel Then

            Me.Visible = False
            If rowASTOPST1 IsNot Nothing Then
                Try
                    ' Record that the Process has Ended
                    rowASTOPST1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                    tdaASTOPST1.Update(tblASTOPST1)
                Catch ex As Exception

                End Try
            End If

            If dst IsNot Nothing Then
                If dst.Tables.Contains("ASTSQLX1") Then
                    Try

                        For Each rowASTDSQLX1 As DataRow In dst.Tables("ASTSQLX1").Select("", "", DataViewRowState.ModifiedCurrent)
                            rowASTDSQLX1.AcceptChanges()
                            rowASTDSQLX1.SetAdded()
                        Next

                        Update_Record_TDA("ASTSQLX1")
                    Catch ex As Exception

                    End Try

                End If

            End If

            If MENU_ITEM_TYPE <> "R" Then
                Call Write_DataSet()
                Call Write_DataSet(True)
            End If

            If ASCMAIN1.Running_in_VS And "" = "NOT NECESSARY ANY MORE" Then
                If MENU_ITEM_TYPE = "R" Or MENU_ITEM_TYPE = "T" Then
                    'UltraExplorerBar1.SaveAsXml(Me.Name & ".exb.xml")
                Else
                    UltraExplorerBar1.Visible = False
                    Application.DoEvents()

                    For Each grp As UltraWinExplorerBar.UltraExplorerBarGroup In UltraExplorerBar1.Groups
                        grp.Settings.AllowEdit = DefaultableBoolean.Default
                        grp.Settings.AllowDrag = DefaultableBoolean.Default
                        grp.Settings.AllowItemDrop = DefaultableBoolean.Default
                        grp.Settings.AllowItemUncheck = DefaultableBoolean.Default
                        grp.Visible = True
                        For Each itm As UltraWinExplorerBar.UltraExplorerBarItem In grp.Items
                            'itm.Settings.AppearancesLarge.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "32\", itm.Key)
                            'itm.Settings.AppearancesSmall.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", itm.Key)
                            itm.Settings.AppearancesLarge.Appearance.Image = Nothing
                            itm.Settings.AppearancesSmall.Appearance.Image = Nothing
                            itm.Settings.Enabled = DefaultableBoolean.Default
                            itm.Settings.AllowEdit = DefaultableBoolean.Default
                            itm.Settings.AllowDragMove = DefaultableBoolean.Default
                            itm.Settings.AllowDragCopy = DefaultableBoolean.Default
                            itm.Settings.UseMnemonics = DefaultableBoolean.Default
                            itm.Visible = True
                        Next
                    Next
                    'UltraExplorerBar1.SaveAsXml(Me.Name & ".exb.xml")
                End If
            End If

            Try
                Dim tlb_btn As New UltraWinToolbars.ButtonTool(CStr(SELECTION_NO))
                Dim tlb_pop As UltraWinToolbars.PopupMenuTool
                tlb_pop = DirectCast(ASFMAIN1.UltraToolbarsManager1.Tools("Help"), UltraWinToolbars.PopupMenuTool)
                If tlb_pop.Tools.Contains(tlb_btn) Then
                    tlb_pop.Tools.Remove(tlb_btn)
                End If
                If tlb.Tools.Contains(tlb_btn) Then
                    tlb.Tools.Remove(tlb_btn)
                End If

            Catch ex As Exception

            End Try


            RemoveContextMenuUltra(Me)

            If ASCMAIN1.ActiveForm Is Me Then
                ASCMAIN1.ActiveForm = Nothing
            End If

            For i As Integer = 0 To ASCMAIN1.ABS_FORMS.Length - 1
                If ASCMAIN1.ABS_FORMS(i) IsNot Nothing AndAlso ASCMAIN1.ABS_FORMS(i) Is Me Then
                    ASCMAIN1.ABS_FORMS(i) = Nothing
                End If
            Next

            If clsASCBASE1 IsNot Nothing Then
                clsASCBASE1.Dispose()
            Else
                ASCMAIN1.Progress("Disposing")
            End If

        End If

        ASFMAIN1.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub RemoveContextMenuUltra(ByVal cntrl As Control)
        ASFMAIN1.UltraToolbarsManager1.SetContextMenuUltra(cntrl, "")
        If cntrl.Controls.Count > 0 Then
            For Each ctl As Control In cntrl.Controls
                RemoveContextMenuUltra(ctl)
            Next
        End If
    End Sub

    Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.OnKeyDown(e)

        If e.KeyCode = System.Windows.Forms.Keys.F8 Then
            Try
                With UltraExplorerBar1.Groups("Screen Control").Items("Update")
                    If .Visible And .Settings.Enabled = DefaultableBoolean.True Then
                        Me.Validate()
                        UltraExplorerBar1.Focus()
                        Click_Command("Update")
                        e.Handled = True
                        Exit Sub
                    End If
                End With

            Catch ex As Exception

            End Try
        End If

        'If e.KeyCode = System.Windows.Forms.Keys.Escape Then
        '    If Not ScreenMode Then
        '        Me.Close()
        '    Else
        '    End If
        'End If


    End Sub

    Private Sub ASFBASE1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If ASCMAIN1.USER_ID = "" Then
            Exit Sub
        End If

        UltraExplorerBar1.ShowDefaultContextMenu = False

        If Not ASCMAIN1.Running_in_VS Then
            UltraExplorerBar1.GroupSettings.AllowEdit = DefaultableBoolean.False
            UltraExplorerBar1.ItemSettings.AllowEdit = DefaultableBoolean.False
        End If

        Call ASCMAIN1.Load_ExplorerBar(Me, Me.UltraExplorerBar1)

        If Me.UltraExplorerBar1.Groups.Contains("Screen Control") Then
            If Me.UltraExplorerBar1.Groups("Screen Control").Items.Contains("Update") Then
                Try
                    Me.UltraExplorerBar1.Groups("Screen Control").Items("Update").Text = "Update (F8)"
                Catch ex As Exception

                End Try
            End If
        End If

        If ASFBASE1_Fill_Panel.Controls.ContainsKey("spl") Then
            Dim spl As SplitContainer = DirectCast(ASFBASE1_Fill_Panel.Controls("spl"), SplitContainer)
            If spl.Panel2.Controls.ContainsKey("tab") Then
                ReParent_Tabs(spl.Panel2.Controls("tab"))
            End If
        End If

        Implement_Security()

        Dim tlb_btn As New UltraWinToolbars.ButtonTool(CStr(SELECTION_NO))
        tlb_btn.Tag = Me.Name
        tlb_btn.SharedProps.Tag = Me.Name
        tlb_btn.SharedProps.Caption = Me.Text
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool
        tlb_pop = DirectCast(ASFMAIN1.UltraToolbarsManager1.Tools("Help"), UltraWinToolbars.PopupMenuTool)
        ASFMAIN1.UltraToolbarsManager1.Tools.Add(tlb_btn)
        tlb_pop.Tools.AddTool(CStr(SELECTION_NO))

        Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(ASFMAIN1.UltraToolbarsManager1.Tools("Small Font"), UltraWinToolbars.StateButtonTool)
        sbt.Tag = "X"
        sbt.Checked = False
        sbt.Tag = ""
    End Sub

    Sub Implement_Security()

        Dim GI As String = ""
        ASCMAIN1.sql = "Select * from ASTCSEC1 where FORM_NAME = '" & Me.Name & "'"
        For Each rowASTCSEC1 As DataRow In ASCDATA1.GetDataTable.Select("", "GROUP_KEY,ITEM_KEY")
            Dim GROUP_KEY As String = rowASTCSEC1.Item("GROUP_KEY")
            Dim ITEM_KEY As String = rowASTCSEC1.Item("ITEM_KEY")
            If UltraExplorerBar1.Groups.Exists(GROUP_KEY) AndAlso _
               UltraExplorerBar1.Groups(GROUP_KEY).Items.Exists(ITEM_KEY) Then
                Dim GIX As String = GROUP_KEY & "." & ITEM_KEY
                If GI <> GIX Then
                    UltraExplorerBar1.Groups(GROUP_KEY).Items(ITEM_KEY).Visible = False
                    Group_Items_Prohibited.Add(ITEM_KEY)
                    GI = GIX
                End If
                Dim SECURITY_CODE As String = rowASTCSEC1.Item("SECURITY_CODE")
                If ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
                    UltraExplorerBar1.Groups(GROUP_KEY).Items(ITEM_KEY).Visible = True
                    Group_Items_Prohibited.Remove(ITEM_KEY)
                End If
            End If
        Next
    End Sub

    Private Sub UltraExplorerBar1_GroupAdded(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.GroupEventArgs) Handles UltraExplorerBar1.GroupAdded
        If ASCMAIN1.Running_in_VS Then
            e.Group.Settings.AllowEdit = IIf(ASCMAIN1.Running_in_VS, 1, 2)
            UltraExplorerBar1.Tag = "*"
        End If
    End Sub

    Private Sub UltraExplorerBar1_GroupClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.GroupEventArgs) Handles UltraExplorerBar1.GroupClick
        If ASCMAIN1.Running_in_VS Then ASCMAIN1.MainForm_pgd.SelectedObject = e.Group
    End Sub

    Private Sub UltraExplorerBar1_GroupEnteringEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.CancelableGroupEventArgs) Handles UltraExplorerBar1.GroupEnteringEditMode
        exbOLDKEY = e.Group.Text
    End Sub

    Private Sub UltraExplorerBar1_GroupExitedEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.GroupEventArgs) Handles UltraExplorerBar1.GroupExitedEditMode
        If ASCMAIN1.Running_in_VS Then
            If e.Group.Key = "" Or e.Group.Key = exbOLDKEY Then
                e.Group.Key = e.Group.Text
            End If
            UltraExplorerBar1.Tag = "*"
            ASCMAIN1.MainForm_pgd.SelectedObject = e.Group
        End If
    End Sub

    Private Sub UltraExplorerBar1_ItemAdded(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemAdded
        If ASCMAIN1.Running_in_VS Then
            e.Item.Settings.AllowEdit = IIf(ASCMAIN1.Running_in_VS, 1, 2)
            UltraExplorerBar1.Tag = "*"
        End If
    End Sub

    Private Sub UltraExplorerBar1_ItemCheckStateChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemCheckStateChanged
        If ASCMAIN1.Running_in_VS Then
            'If UltraExplorerBar1.CheckedItem Is Not null Then
            ASCMAIN1.MainForm_pgd.SelectedObject = UltraExplorerBar1.CheckedItem
            'End If
        End If
    End Sub

    Private Sub UltraExplorerBar1_ItemClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemClick
        If ASCMAIN1.Running_in_VS Then
            ASCMAIN1.MainForm_pgd.SelectedObject = e.Item
        End If

        If Not e.Item.SettingsResolved.Enabled Then
            Exit Sub
        End If

        If UltraExplorerBar1.Tag <> "CLICK" Then
            UltraExplorerBar1.Tag = "CLICK"
            Call Click_Command(e.Item.Key)
            UltraExplorerBar1.Tag = ""
        End If
    End Sub

    Sub Click_Command(ByVal eItemKey As String, Optional ByVal e As System.Windows.Forms.KeyEventArgs = Nothing)
        If e IsNot Nothing Then
            e.Handled = True
            Me.ProcessTabKey(Not e.Shift)
        End If

        Application.DoEvents()

        '********************** start of lots of flashing on the screen

        If dst IsNot Nothing Then
            For Each obj_ds As Object In Bound_DataSources
                If obj_ds IsNot Nothing AndAlso TypeOf obj_ds Is DataTable Then
                    Dim tbl As DataTable = DirectCast(obj_ds, DataTable)
                    Dim TABLE_NAME As String = tbl.TableName
                    If dst.Tables.Contains(TABLE_NAME) Then
                        If Me.BindingContext.Contains(dst.Tables(TABLE_NAME)) Then

                            Try
                                Dim X As CurrencyManager = Me.BindingContext(obj_ds)
                                X.EndCurrentEdit()
                            Catch ex As Exception

                            End Try

                        End If
                    End If
                End If
            Next
        End If

        '******************** end of flashing on the screen

        If EntryMode = "" Then
            Dim T() As String
            ReDim T(ROWs.Count - 1)
            Dim i As Integer = 0
            Dim TABLE_NAME As String = ""
            For Each TABLE_NAME In ROWs.Keys
                T(i) = TABLE_NAME
                i = i + 1
            Next
            For i = 0 To ROWs.Count - 1 ' Each TABLE_NAME As String In New String() {ROWs.Keys.ToString}
                TABLE_NAME = T(i)
                If Mid(TABLE_NAME, 3, 5) = "TPARM" Then
                    If ROWs(TABLE_NAME) IsNot Nothing Then
                        If ROWs(TABLE_NAME).Table.Columns(0).ColumnName = Mid(TABLE_NAME, 1, 2) & "_PARM_KEY" Then
                            ROWs(TABLE_NAME) = LookUp(TABLE_NAME, "Z")
                        End If
                    End If
                End If
            Next
        End If

        DATETIME_STAMP = Now + ASCMAIN1.NowTSD
        ' RECORD STATISTIC HERE

        Dim rowASTOPST2 As DataRow = Nothing
        Dim HFS_VALUES As String = ""

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            If eItemKey <> "Done" And eItemKey <> "Proceed" Then
            If Not dst.Tables.Contains("ASTOPST2") Then
                Create_TDA(dst.Tables.Add, "ASTOPST2", "*")
            End If

            rowASTOPST2 = dst.Tables("ASTOPST2").NewRow
            With rowASTOPST2
                .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                .Item("SELECTION_NO") = SELECTION_NO
                .Item("RE_XNO") = RE_XNO
                RE_XNO_STAT += 1
                .Item("RE_XNO_STAT") = RE_XNO_STAT
                .Item("USER_ID") = ASCMAIN1.USER_ID
                .Item("MENU_ID") = MENU_ID
                .Item("MENU_ITEM_TYPE") = MENU_ITEM_TYPE
                .Item("MENU_ITEM_OBJECT") = MENU_ITEM_OBJECT
                .Item("INIT_DATE") = DATETIME_STAMP
                '   .Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                .Item("YYYYPP") = ASCMAIN1.CYP
                .Item("XNO") = XNO
                .Item("PRD_CLOSE_IND") = ASCMAIN1.EOM
                .Item("FORM_INSTANCE_NO") = FORM_INSTANCE_NO
                HFS_VALUES = ""
                For Each k As String In HFs.Keys
                    HFS_VALUES &= ";" & k & ":" & HFs(k)
                Next
                HFS_VALUES = Mid(Mid(HFS_VALUES, 2), 1, 500)
                .Item("HFS_VALUES") = Mid(Mid(HFS_VALUES, 2), 1, 500)
                .Item("COMMAND") = eItemKey
            End With
            dst.Tables("ASTOPST2").Rows.Add(rowASTOPST2)
                If dst.Tables.Contains("ASTOPST2") Then
                    Update_Record_TDA("ASTOPST2")
                End If
            End If

        End If


        IsLoading = True
        If Not Group_Items_Prohibited.Contains(eItemKey) Then
            Proceed_PreReq(eItemKey)
            'Write_Audit(MENU_ITEM_OBJECT, eItemKey, HFs, EMsg)
            ' - NEED TO WRITE TO ASTOPST2 HERE BUT IN BETWEEN PREREQ AND PROCEED
            If Not ScreenMode Then
                ASCMAIN1.MultiTask_Release()
            End If
        Else
            MsgBox("Access not Permitted", MsgBoxStyle.OkOnly, "Cannot Perform the Requested Action")
        End If
        IsLoading = False
        IsDone = False

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            If MENU_ITEM_TYPE = "R" Then
                ASCMAIN1.sql = "Update ASTOPST2 Set LAST_DATE = SYSDATE" & vbCrLf _
                    & " where SESSION_NO = '" & ASCMAIN1.SESSION_NO & "'" & vbCrLf _
                    & "   and SELECTION_NO = " & CStr(SELECTION_NO) & vbCrLf _
                    & "   and RE_XNO = " & CStr(RE_XNO) & vbCrLf _
                    & "   and RE_XNO_STAT = " & CStr(RE_XNO_STAT)
                ASCDATA1.ExecuteSQL()

            Else
                If rowASTOPST2 IsNot Nothing Then
                    If rowASTOPST2.Item("HFS_VALUES") & "" = "" Then
                        HFS_VALUES = ""
                        For Each k As String In HFs.Keys
                            HFS_VALUES &= ";" & k & ":" & HFs(k)
                        Next
                        HFS_VALUES = Mid(Mid(HFS_VALUES, 2), 1, 500)
                        rowASTOPST2.Item("HFS_VALUES") = HFS_VALUES
                    End If
                    rowASTOPST2.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                    Update_Record_TDA("ASTOPST2")
                End If
            End If
        End If

    End Sub

    Public Overridable Sub Proceed_PreReq(ByVal eItemKey As String)

    End Sub

    Private Sub UltraExplorerBar1_ItemEnteringEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.CancelableItemEventArgs) Handles UltraExplorerBar1.ItemEnteringEditMode
        exbOLDKEY = e.Item.Text
    End Sub

    Private Sub UltraExplorerBar1_ItemExitedEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinExplorerBar.ItemEventArgs) Handles UltraExplorerBar1.ItemExitedEditMode
        If ASCMAIN1.Running_in_VS Then
            If e.Item.Key = "" Or e.Item.Key = exbOLDKEY Then
                e.Item.Key = e.Item.Text
            End If
            e.Item.Settings.AppearancesLarge.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "32\", e.Item.Key)
            e.Item.Settings.AppearancesSmall.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", e.Item.Key)
            UltraExplorerBar1.Tag = "*"
            ASCMAIN1.MainForm_pgd.SelectedObject = e.Item
        End If
    End Sub

    Sub Contain_Control(ByVal KEY As String, ByVal ctl As Control)

        Try
            With ctl
                UltraExplorerBar1.Groups(KEY).Settings.Style = UltraWinExplorerBar.GroupStyle.ControlContainer
                .Parent = UltraExplorerBar1.Groups(KEY).Container
                .Location = New System.Drawing.Point(0, 0)
                .Visible = True
                UltraExplorerBar1.Groups(KEY).Container.Height = .Height
                .Dock = System.Windows.Forms.DockStyle.Fill
            End With

        Catch ex As Exception

        End Try
    End Sub

    Sub Context_Launch( _
    ByVal command As String, _
    ByVal key As String, _
    ByVal MENU_ITEM_DESC As String, _
    ByVal MENU_ITEM_OBJECT As String, _
    Optional ByVal MENU_ITEM_TYPE As String = "", _
    Optional ByVal MENU_ID As String = "", _
    Optional ByVal CreateNewForm As Boolean = False)

        If Not ASCMAIN1.MENU_ITEM_OBJECTs.Contains(MENU_ITEM_OBJECT) Then
            MsgBox("Insufficient Security", MsgBoxStyle.OkOnly, "Cannot Launch " & MENU_ITEM_DESC)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading " & MENU_ITEM_DESC)

        If Not frmASFBASE1s.ContainsKey(MENU_ITEM_OBJECT) Then
            Dim f As ASFBASE1 = Nothing
            frmASFBASE1s.Add(MENU_ITEM_OBJECT, f)
        End If

        If MENU_ITEM_TYPE = "" Or MENU_ID = "" Then
            Dim SQL As String = "MENU_ITEM_OBJECT = '" & MENU_ITEM_OBJECT & "'"
            If MENU_ITEM_TYPE <> "" Then
                SQL &= " and MENU_ITEM_TYPE = '" & MENU_ITEM_TYPE & "'"
            End If
            If MENU_ID <> "" Then
                SQL &= " and MENU_ID = '" & MENU_ID & "'"
            End If
            Dim rows() As DataRow = ASCMAIN1.tblASTMENU1.Select(SQL)
            If rows.Length > 0 Then
                MENU_ID = rows(0).Item("MENU_ID")
                MENU_ITEM_TYPE = rows(0).Item("MENU_ITEM_TYPE")
            End If
        End If

        If frmASFBASE1s(MENU_ITEM_OBJECT) Is Nothing OrElse frmASFBASE1s(MENU_ITEM_OBJECT).IsDisposed Then
            frmASFBASE1s(MENU_ITEM_OBJECT) = ASCMAIN1.Launch_Form(MENU_ITEM_OBJECT, MENU_ITEM_TYPE, MENU_ID)
            CreateNewForm = False
        End If

        Dim origForm As ABSolution.ASFBASE1 = Nothing

        If frmASFBASE1s(MENU_ITEM_OBJECT) IsNot Nothing Then
            If frmASFBASE1s(MENU_ITEM_OBJECT).ScreenMode Then
                If CreateNewForm Then
                    origForm = frmASFBASE1s(MENU_ITEM_OBJECT)
                    frmASFBASE1s(MENU_ITEM_OBJECT) = ASCMAIN1.Launch_Form(MENU_ITEM_OBJECT, MENU_ITEM_TYPE, MENU_ID)
                Else
                    frmASFBASE1s(MENU_ITEM_OBJECT).Remote_Control("Done")
                End If
            End If

            If Not frmASFBASE1s(MENU_ITEM_OBJECT).ScreenMode Then
                frmASFBASE1s(MENU_ITEM_OBJECT).Activate()
                frmASFBASE1s(MENU_ITEM_OBJECT).Remote_Control(command, key)
                frmASFBASE1s(MENU_ITEM_OBJECT).Focus()
            End If
        End If

        If origForm IsNot Nothing Then
            frmASFBASE1s(MENU_ITEM_OBJECT) = origForm
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Context_Launch( _
    ByVal command As String, _
    ByVal keys As Dictionary(Of String, Object), _
    ByVal MENU_ITEM_DESC As String, _
    ByVal MENU_ITEM_OBJECT As String, _
    Optional ByVal MENU_ITEM_TYPE As String = "", _
    Optional ByVal MENU_ID As String = "", _
    Optional ByVal CreateNewForm As Boolean = False)

        If Not ASCMAIN1.MENU_ITEM_OBJECTs.Contains(MENU_ITEM_OBJECT) Then
            MsgBox("Insufficient Security", MsgBoxStyle.OkOnly, "Cannot Launch " & MENU_ITEM_DESC)
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading " & MENU_ITEM_DESC)

        If Not frmASFBASE1s.ContainsKey(MENU_ITEM_OBJECT) Then
            Dim f As ASFBASE1 = Nothing
            frmASFBASE1s.Add(MENU_ITEM_OBJECT, f)
        End If

        If frmASFBASE1s(MENU_ITEM_OBJECT) Is Nothing OrElse frmASFBASE1s(MENU_ITEM_OBJECT).IsDisposed Then
            frmASFBASE1s(MENU_ITEM_OBJECT) = ASCMAIN1.Launch_Form(MENU_ITEM_OBJECT, MENU_ITEM_TYPE, MENU_ID)
            CreateNewForm = False
        End If

        Dim origForm As ABSolution.ASFBASE1 = Nothing

        If frmASFBASE1s(MENU_ITEM_OBJECT) IsNot Nothing Then
            If frmASFBASE1s(MENU_ITEM_OBJECT).ScreenMode Then
                If Not CreateNewForm Then
                    If frmASFBASE1s(MENU_ITEM_OBJECT).UltraExplorerBar1.Groups("Screen Control").Items("Done").Visible And _
                        frmASFBASE1s(MENU_ITEM_OBJECT).UltraExplorerBar1.Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.True Then

                    Else
                        CreateNewForm = True
                    End If
                End If
                If CreateNewForm Then
                    origForm = frmASFBASE1s(MENU_ITEM_OBJECT)
                    frmASFBASE1s(MENU_ITEM_OBJECT) = ASCMAIN1.Launch_Form(MENU_ITEM_OBJECT, MENU_ITEM_TYPE, MENU_ID)
                Else
                    frmASFBASE1s(MENU_ITEM_OBJECT).Remote_Control("Done")
                End If
            End If

            If Not frmASFBASE1s(MENU_ITEM_OBJECT).ScreenMode Then
                frmASFBASE1s(MENU_ITEM_OBJECT).Activate()
                frmASFBASE1s(MENU_ITEM_OBJECT).Remote_Control(command, keys)
                frmASFBASE1s(MENU_ITEM_OBJECT).Focus()
            End If
        End If

        If origForm IsNot Nothing Then
            frmASFBASE1s(MENU_ITEM_OBJECT) = origForm
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Public Overridable Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        'Dim return_key As Object = Nothing
        'Return return_key

        Dim x As Dictionary(Of String, Object) = Nothing
        Return Remote_Control(command, x)

    End Function

    Public Overridable Function Remote_Control( _
    ByVal command As String, _
    ByVal keys As Dictionary(Of String, Object)) As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        ' can't see command in Items collection when command is a string, like "View"
        ' command = "Done" added to facillitate FM
        If command <> "" AndAlso (command = "View" Or command = "Edit" Or command = "Done" Or UltraExplorerBar1.Groups("Screen Control").Items.Contains(command)) Then
        If UltraExplorerBar1.Groups("Screen Control").Items(command).Visible And _
           UltraExplorerBar1.Groups("Screen Control").Items(command).Settings.Enabled = DefaultableBoolean.True Then

            If keys IsNot Nothing Then
                For Each COLUMN_NAME As String In keys.Keys
                    Absx1.txtFor(COLUMN_NAME).Text = keys(COLUMN_NAME)
                Next
            End If

            Click_Command(command)
            End If
        End If

        Return return_key
    End Function

    Public Function Column_Values(ByVal ParamArray Column_Value() As Object) As Dictionary(Of String, Object)

        Dim D As New Dictionary(Of String, Object)

        Dim key As String = ""

        For i As Int16 = 1 To Column_Value.Length
            If i Mod 2 = 1 Then
                key = Column_Value(i - 1)
            Else
                D.Add(key, Column_Value(i - 1))
            End If
        Next

        Return D
    End Function
End Class