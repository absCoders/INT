Imports Infragistics.Win
Imports System.Configuration
Imports System.Threading

Public Class ASFMAIN1

    ' Used for Serial and Comm Port Control
    '************************************************************************************
    Public Delegate Sub ScaleDelegate(ByVal ScannedString As String)
    'Public WithEvents scaleSerialPort As New System.IO.Ports.SerialPort
    'Public scaleWeightDelegate As ScaleDelegate = Nothing

    Public laserPrinterIP As String = String.Empty
    Public laserPrinterName As String = String.Empty
    Public WithEvents labelPrinterSerialPort As New System.IO.Ports.SerialPort
    '************************************************************************************

    Dim Form_Loaded As Boolean
    Dim WithEvents tvw As Infragistics.Win.UltraWinTree.UltraTree
    Dim ActiveNode As Infragistics.Win.UltraWinTree.UltraTreeNode
    Dim dNode As Infragistics.Win.UltraWinTree.UltraTreeNode
    Dim tblASTMENU1 As DataTable

    Public tblImages As New DataTable
    Public MENU_ID As String
    Public MENU_ITEM_TYPE As String
    Public MENU_ITEM_OBJECT As String
    Public MENU_ITEM_DESC As String
    Public MENU_ITEM_SECURITY As String
    Public MENU_ITEM_PP As String
    Public MENU_ITEM_EOM_CHECK As String
    Public MENU_ITEM_FORM As String
    Public MENU_ITEM_PASSWORD As String
    Public MENU_ITEM_STANDALONE As String
    Public MODULE_ID As String
    Public loadSplash As ASFFLOAD
    Dim MODULE_IDs As String = ""

    Public FormToShow As ASFBASE1

    Private Sub StartSplash()
        loadSplash = New ASFFLOAD(MENU_ITEM_DESC, Me.Bounds, True)
        Application.Run(loadSplash)
    End Sub

    Private Sub CloseSplash()
        If loadSplash Is Nothing OrElse loadSplash.IsHandleCreated = False Then
            Return
        End If
        loadSplash.Invoke(New EventHandler(AddressOf loadSplash.EndForm))
        loadSplash.Dispose()
        loadSplash = Nothing
    End Sub


    Private Sub ASFMAIN1_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated

    End Sub

    Private Sub ASFMAIN1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not e.Cancel Then
            Call ASCMAIN1.Temp_Table_Cleanup(False)
            Dim Sql As String
            If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                Sql = "Update ASTLOGS1 set SESSION_STATUS = 'X', DATE_LOGGED_OFF = GETDATE() where SESSION_NO = '" & ASCMAIN1.SESSION_NO & "'"
            Else
                Sql = "Update ASTLOGS1 set SESSION_STATUS = 'X', DATE_LOGGED_OFF = SYSDATE where SESSION_NO = '" & ASCMAIN1.SESSION_NO & "'"
            End If
            ASCDATA1.ExecuteSQL(Sql)
            Sql = "Update ASTOPST1 set LAST_DATE = (SELECT DATE_LOGGED_OFF from ASTLOGS1 where SESSION_NO = '" & ASCMAIN1.SESSION_NO & "') where SESSION_NO = '" & ASCMAIN1.SESSION_NO & "' and SELECTION_NO = 0"
            ASCDATA1.ExecuteSQL(Sql)
        End If
    End Sub

    Private Sub ASFMAIN1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Setup_Automated_Login()

    End Sub

    Public Sub Setup_Automated_Login(Optional fromService As Boolean = False)

        ASCMAIN1.ABSWEB = fromService

        If Not ASCMAIN1.ABSWEB Then
            Dim ISLfile As String = Application.StartupPath & "\ABS1.isl"

            If Not IO.File.Exists(ISLfile) Then
                Dim i As Integer = InStr(ISLfile, "\bin\x86")
                IO.File.Copy(Mid(ISLfile, 1, i) & "\ABS1.isl", ISLfile)
            End If

            Infragistics.Win.AppStyling.StyleManager.Load(ISLfile)
        End If

        Dim folder_prefix As String

        If Debugger.IsAttached Then ' UCase(My.Application.Info.DirectoryPath) Like "C:\USERS\*\VS\*" Then ' If UCase(My.Application.Info.DirectoryPath) Like "C:\VS\*" Then
            ASCMAIN1.Running_in_VS = True
            folder_prefix = "\..\..\..\..\"
            Dim i As Integer = InStr(My.Application.Info.DirectoryPath, "\ABS\bin")
            ASCMAIN1.SOLUTION = UCase(Mid(My.Application.Info.DirectoryPath, i - 3, 3))
            ' ASCMAIN1.SOLUTION = UCase(Mid(My.Application.Info.DirectoryPath, 7, 3))
        Else
            ASCMAIN1.Running_in_VS = False
            folder_prefix = "\..\"
            ASCMAIN1.SOLUTION = UCase(Split(My.Application.Info.DirectoryPath, "\")(3))
        End If

        Dim root_folder As String = ""
        If ASCMAIN1.ABSWEB Then
            root_folder = "C:\VS\ANE_SVC\folders\"
        Else
            root_folder = My.Application.Info.DirectoryPath & folder_prefix
        End If

        ASCMAIN1.Folders.Add("Images", ASCMAIN1.GetPath(root_folder & "Images\"))
        ASCMAIN1.Folders.Add("Reports", ASCMAIN1.GetPath(root_folder & "Reports\"))
        ASCMAIN1.Folders.Add("DataSets", ASCMAIN1.GetPath(root_folder & "DataSets\"))
        ASCMAIN1.Folders.Add("Temp", ASCMAIN1.GetPath(root_folder & "Temp\"))
        ASCMAIN1.Folders.Add("Work", ASCMAIN1.GetPath(root_folder & "Work\"))
        ASCMAIN1.Folders.Add("bin", ASCMAIN1.GetPath(root_folder & "bin\"))
        ASCMAIN1.Folders.Add("Help", ASCMAIN1.GetPath(root_folder & "Help\"))
        ASCMAIN1.Folders.Add("Archive", ASCMAIN1.GetPath(root_folder & "Archive\"))
        ASCMAIN1.Folders.Add("Attach", ASCMAIN1.GetPath(root_folder & "Attach\"))
        ASCMAIN1.Folders.Add("root", ASCMAIN1.GetPath(root_folder))
        ASCMAIN1.Folders.Add("SharedRoot", "R:\" & ASCMAIN1.SOLUTION & "\")



        If My.Computer.Name = "WJZ64B" Then
            ASCMAIN1.Folders.Add("Oracle", "C:\oracle\product\11.2.0\dbhome_1\")
        Else
            ASCMAIN1.Folders.Add("Oracle", "C:\oracle\product\11.2.0\Client_1\")
        End If

        Dim image_filename As String = ASCMAIN1.SOLUTION & ".bmp"
        If Not My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Images") & "ABS\" & image_filename) Then
            image_filename = "abs.bmp"
        End If
        If Not ASCMAIN1.ABSWEB Then UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & image_filename)

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            image_filename = "SLP.BMP"
            UltraPictureBox1.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "ABS\" & image_filename)
        End If


        If Not ASCMAIN1.ABSWEB Then Me.UltraSpellChecker1.UserDictionary = ASCMAIN1.Folders("Archive") & ASCMAIN1.CLIENT & ".DIC"

        Try
            ASCMAIN1.Register_Form(Me)

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.InnerException.ToString)
            MsgBox(ex.InnerException.Message)
            MsgBox(ex.Data)
            MsgBox(ex.Source)
        End Try
        'MsgBox("3x")
        Call ASCMAIN1.Set_DBS_Dependent_Strings()
        'Me.Text = "ABSolute-1; v" & My.Application.Info.Version.ToString
        'Me.Text = "ABSolution; v" & My.Application.Info.Version.ToString
        Dim Title As String = My.Application.Info.Title
        Try
            Title = My.Settings("Title")
        Catch ex As Exception

        End Try
        If Title = "" Then Title = "ABSolution"
        Me.Text = Title & "; v" & My.Application.Info.Version.ToString

        Me.Width = Screen.PrimaryScreen.WorkingArea.Width * 0.95
        Me.Height = Screen.PrimaryScreen.WorkingArea.Height * 0.95
        Me.WindowState = FormWindowState.Maximized
        Call ASCMAIN1.Center(Me)

        ASCMAIN1.MainForm = Me
        ASCMAIN1.MainForm_pgd = Me.pgdExplorerBar

        UltraDockManager1.ControlPanes(0).Activate()

        Dim popupcontrol As UltraWinToolbars.PopupControlContainerTool = UltraToolbarsManager1.Tools("Resources")
        popupcontrol.Control = grpSupport
        'Dim txtResources As New UltraWinEditors.UltraTextEditor
        'popupcontrol.Control = txtResources
        'txtResources.Visible = True
        'txtResources.ReadOnly = True
        'txtResources.Text = "www.infragistics.com"

        'UltraToolbarsManager1.Toolbars("tlbtooltip").ParentCollection.Add.Parent = grpToolTip

        With tblImages
            .Columns.Add("ITEM_KEY")
            .Columns.Add("ITEM_TEXT")
            .Columns.Add("ITEM_TOOLTIPTEXT")
            .PrimaryKey = New DataColumn() { .Columns("ITEM_KEY")}
        End With

        'NotifyIcon1.Icon = Me.Icon
    End Sub

    Private Sub ASFMAIN1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown

        Dim load_DLLs As System.Threading.Thread = New Thread(New ThreadStart(AddressOf StartLoadingDLLs))
        load_DLLs.Start()

        If Not Form_Loaded And Me.Visible Then
            Form_Loaded = True
            Dim F As New ASFLOGON
            F.ShowDialog()
            F.Dispose()

            If ASCMAIN1.USER_ID = "" Then
                End
            Else
                Try
                    Site_Specific_Settings()
                    ASCMAIN1.absTimer.Enabled = True
                    Load_User_Icon()
                    Dim sbt As UltraWinToolbars.StateButtonTool
                    sbt = DirectCast(UltraToolbarsManager1.Tools("Hide Menu after Selection"), UltraWinToolbars.StateButtonTool)
                    sbt.Checked = ("True" = GetSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFMAIN1.HIDE_MENU_AFTER_SELECTION"))
                    sbt = DirectCast(UltraToolbarsManager1.Tools("Enable Change to Favorite Descriptions"), UltraWinToolbars.StateButtonTool)
                    sbt.Checked = ("True" = GetSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFMAIN1.ENABLE_CHANGE_TO_FAVORITE_DESCRIPTIONS"))
                    UltraStatusBar1.Panels("USER_ID").Text = ASCMAIN1.USER_ID
                    If ASCMAIN1.USER_ADMIN = "1" Then
                        UltraStatusBar1.Panels("USER_ID").Appearance.ForeColor = Color.Red
                        UltraStatusBar1.Panels("USER_ID").ToolTipText &= " (admin)"
                    End If
                    UltraStatusBar1.Panels("DBS_COMPANY").Text = ASCMAIN1.DBS_COMPANY
                    If ASCMAIN1.DBS_SERVER = "" Then
                        UltraStatusBar1.Panels("DBS_SERVER").Text = "(local)"
                    Else
                        UltraStatusBar1.Panels("DBS_SERVER").Text = ASCMAIN1.DBS_SERVER
                    End If

                    If ASCMAIN1.DBS_SERVER <> ASCMAIN1.DBS_COMPANY Then
                        UltraStatusBar1.Panels("DBS_SERVER").Appearance.ForeColor = Color.Red
                        Me.Text &= " - YOU ARE NOT IN PRODUCTION"
                    End If

                    UltraStatusBar1.Panels("MENU_ITEM_OBJECT").Text = Me.Name
                    Call Build_Menu()
                    'exbASTMENU1.NavigationCurrentGroupAreaMinHeight = 300
                    exbASTMENU1.NavigationMaxGroupHeaders = 8
                    UltraDockManager1.Visible = True

                    exbASTMENU1.Groups(0).Selected = True
                    If Not exbASTMENU1.Groups(0).Active Then
                        Try
                            exbASTMENU1.Groups(0).Active = True
                        Catch ex As Exception

                        End Try
                    End If

                    Call ASCMAIN1.Load_Views()
                    Call ASCMAIN1.Load_MRUs()

                    If InStr(ASCMAIN1.USER_SECURITY_CODEs, "SY") <> 0 Or ASCMAIN1.Running_in_VS Then
                        UltraToolbarsManager1.Tools("Developer").SharedProps.Visible = True
                        UltraToolbarsManager1.Tools("Show Developer").SharedProps.Visible = False
                    Else
                        UltraDockManager1.PaneFromControl(Me.exbABS).Close()
                    End If

                    If Environment.GetCommandLineArgs.Count >= 6 Then
                        If Environment.GetCommandLineArgs.ElementAt(4) = "JS" Then
                            Dim JOB_STREAM_CODE As String = Environment.GetCommandLineArgs.ElementAt(5)
                            Dim rowASTJOBM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ASTJOBM1 WHERE JOB_STREAM_CODE = :PARM1", "V", JOB_STREAM_CODE)
                            If rowASTJOBM1 IsNot Nothing Then
                                ASCMAIN1.JOB_STREAM_CODE = JOB_STREAM_CODE
                                ASCMAIN1.USER_MENU_ITEM_OBJECT = "ASFJOBM1"
                            Else
                                End
                            End If
                        End If
                    End If

                    If ASCMAIN1.USER_MENU_ITEM_OBJECT <> "" Then
                        UltraToolbarsManager1.Tools("Window").SharedProps.Visible = False
                        UltraToolbarsManager1.Tools("Menu").SharedProps.Visible = False
                        Call Launch_Single_Form()
                    End If
                Catch ex As Exception

                    MsgBox(ex.Message)

                End Try

            End If
        End If

        load_DLLs.Abort()
        load_DLLs = Nothing


        If ASCMAIN1.Running_in_VS Then
            Dim x As New secrets
            x.Set_SecretS()
        End If

        ' Check for Single Execution
        Dim AppEXEName As String = My.Application.Info.AssemblyName
        If (Len(AppEXEName) = 3 Or (Len(AppEXEName) = 8 And AppEXEName <> "ASEMENU1")) And ASCMAIN1.SEXE_NO <> "" Then

            Dim rowASTMENU1 As DataRow = ASCDATA1.GetDataRow(
               "Select * from ASTMENU1 where MENU_ID = :PARM1 and MENU_ITEM_TYPE = :PARM2 and MENU_ITEM_OBJECT = :PARM3",
               "VVV", New String() {ASCMAIN1.rowASTSEXE1.Item("MENU_ID"), ASCMAIN1.rowASTSEXE1.Item("MENU_ITEM_TYPE"), ASCMAIN1.rowASTSEXE1.Item("MENU_ITEM_OBJECT")})
            If rowASTMENU1 IsNot Nothing Then
                ASCMAIN1.JOB_STREAM_CODE = ASCMAIN1.rowASTSEXE1.Item("JOB_STREAM_CODE") & ""
                Launch_Form(rowASTMENU1)
            End If
        End If
    End Sub

    Sub Add_Group(ByVal KEY As String, ByVal MENU_ITEM_DESC As String, ByVal IMAGE_FOLDER As String, ByVal IMAGE_FILE As String, ByRef tvw As Infragistics.Win.UltraWinTree.UltraTree)
        exbASTMENU1.Groups.Add(KEY, MENU_ITEM_DESC)
        exbASTMENU1.Groups(KEY).Settings.AppearancesLarge.HeaderAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, IMAGE_FILE)
        exbASTMENU1.Groups(KEY).Settings.AppearancesSmall.HeaderAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, IMAGE_FILE)
        exbASTMENU1.Groups(KEY).Settings.Style = Infragistics.Win.UltraWinExplorerBar.GroupStyle.ControlContainer

        tvw = New Infragistics.Win.UltraWinTree.UltraTree
        tvw.Parent = exbASTMENU1.Groups(KEY).Container
        tvw.Location = New Point(0, 0)
        tvw.Dock = System.Windows.Forms.DockStyle.Fill
        tvw.Override.HotTracking = Infragistics.Win.DefaultableBoolean.True
        tvw.HideSelection = False
        tvw.UseOsThemes = False
        tvw.Visible = True
        AddHandler tvw.MouseUp, AddressOf tvw_MouseUp
        AddHandler tvw.AfterLabelEdit, AddressOf tvw_AfterLabelEdit
        AddHandler tvw.BeforeDelete, AddressOf tvw_BeforeDelete
        AddHandler tvw.AfterDelete, AddressOf tvw_AfterDelete
        AddHandler tvw.DoubleClick, AddressOf tvw_DoubleClick

        Me.UltraToolbarsManager1.SetContextMenuUltra(tvw, "popASTMENU1")
    End Sub

    Private Sub UltraToolbarsManager1_AfterToolCloseup(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolDropdownEventArgs) Handles UltraToolbarsManager1.AfterToolCloseup
        If e.Tool.Key = "txtMenu" Then

        End If
    End Sub

    Private Sub UltraToolbarsManager1_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs) Handles UltraToolbarsManager1.BeforeToolDropdown
        If e.Tool.OwningToolbar IsNot Nothing Then
            If e.Tool.OwningToolbar.Key = "Main" Then
                Exit Sub
            End If
        End If

        If e.Tool.Key = "txtMenu" Then
            Dim txtctl As UltraWinEditors.UltraTextEditor = Nothing
            If TypeOf (e.SourceControl) Is UltraWinEditors.UltraTextEditor Then
                txtctl = DirectCast(e.SourceControl, UltraWinEditors.UltraTextEditor)
            ElseIf TypeOf (e.SourceControl) Is EmbeddableTextBoxWithUIPermissions Then
                txtctl = DirectCast(e.SourceControl.Parent, UltraWinEditors.UltraTextEditor)
            End If
            If txtctl Is Nothing OrElse txtctl.ReadOnly Or Not txtctl.Enabled Then
                e.Cancel = True
            Else
                Dim show_popup As Boolean = False

                Dim F As ASFBASE0 = Nothing
                Dim c As Control = txtctl
                Do While c.Parent IsNot Nothing
                    If TypeOf (c.Parent) Is ASFBASE0 Then
                        F = DirectCast(c.Parent, ASFBASE0)
                        Exit Do
                    End If
                    c = c.Parent
                Loop

                Dim COLUMN_NAME As String = F.Absx1.GetABSColumnName(txtctl) ' ASCMAIN1.MRU_COLUMN_NAME

                'If COLUMN_NAME = "" Then
                '    e.Cancel = True
                '    Exit Sub
                'End If

                If Not ASCMAIN1.MRUs.ContainsKey(COLUMN_NAME) Then
                    UltraToolbarsManager1.Tools("COLUMN_NAME").SharedProps.Visible = False
                    UltraToolbarsManager1.Tools("History").SharedProps.Visible = False
                    UltraToolbarsManager1.Tools("Most Recently Used").SharedProps.Visible = False
                Else
                    show_popup = True
                    UltraToolbarsManager1.Tools("COLUMN_NAME").SharedProps.Visible = True
                    UltraToolbarsManager1.Tools("History").SharedProps.Visible = True
                    UltraToolbarsManager1.Tools("Most Recently Used").SharedProps.Visible = True
                    UltraToolbarsManager1.Tools("COLUMN_NAME").SharedProps.Caption = ASCMAIN1.Make_Caption(COLUMN_NAME)
                    Dim LL As List(Of String) = ASCMAIN1.MRUs(COLUMN_NAME)
                    If LL.Count = 0 Then
                        e.Cancel = True
                        Exit Sub
                    End If
                    Dim comboboxtool As UltraWinToolbars.ComboBoxTool = UltraToolbarsManager1.Tools("History")
                    Dim valuelist1 As New ValueList()
                    For i As Integer = 0 To LL.Count - 1
                        valuelist1.ValueListItems.Add(New ValueListItem(LL(i)))
                    Next
                    comboboxtool.ValueList = valuelist1
                    comboboxtool.SelectedIndex = 0
                    'comboboxtool.DropDownStyle = DropDownStyle.DropDown
                    comboboxtool.DropDownStyle = DropDownStyle.DropDownList
                    comboboxtool.Value = "<Select Value>"
                    comboboxtool.SharedProps.ToolTipText = "Code Values Entered sorted from First to Last"
                    comboboxtool.AutoComplete = True


                    Dim listtool As UltraWinToolbars.ListTool = UltraToolbarsManager1.Tools("Most Recently Used")
                    listtool.ListToolItems.Clear()
                    If LL.Count > 0 Then
                        For i As Integer = LL.Count - 1 To 0 Step -1
                            listtool.ListToolItems.Add(LL(i).ToString, LL(i).ToString)
                            If LL.Count - i > 3 Then
                                Exit For
                            End If
                        Next
                    End If
                End If



                Dim btntool As UltraWinToolbars.ButtonTool = UltraToolbarsManager1.Tools("Attachments")
                btntool.SharedProps.Enabled = ASCMAIN1.ActiveForm.ScreenMode
                btntool.SharedProps.Visible = False ' no need to offer attachements in the MRU menu - attachments should be form-context sensitive


                btntool = UltraToolbarsManager1.Tools("Spell Check")
                If txtctl.SpellChecker Is Nothing Then
                    btntool.SharedProps.Visible = False
                Else
                    btntool.SharedProps.Visible = True
                    show_popup = True
                End If

                If Not show_popup Then
                    e.Cancel = True
                End If
            End If

            Exit Sub
        End If

        If e.Tool.Key <> "popASTMENU1" Then
            Exit Sub
        End If

        Dim t As Infragistics.Win.UltraWinTree.UltraTree = DirectCast(e.SourceControl, Infragistics.Win.UltraWinTree.UltraTree)
        If t.HotTrackingNode IsNot Nothing Then
            t.HotTrackingNode.Selected = True
        End If
        If t.SelectedNodes.Count = 0 Then
            e.Cancel = True
            Exit Sub
        End If
        If Mid(Split(t.SelectedNodes(0).Tag, Chr(1))(1), 1, 1) = "M" Then
            e.Cancel = True
            Exit Sub
        End If

        ActiveNode = t.SelectedNodes(0) ' .Clone

        If exbASTMENU1.ActiveGroup Is Nothing Then Exit Sub

        If exbASTMENU1.ActiveGroup.Key = "*" Then
            UltraToolbarsManager1.Tools("Add to Favorites").SharedProps.Enabled = False
            UltraToolbarsManager1.Tools("Sort Favorites").SharedProps.Enabled = True
            UltraToolbarsManager1.Tools("Rename").SharedProps.Enabled = True
            UltraToolbarsManager1.Tools("Delete").SharedProps.Enabled = True
        Else
            UltraToolbarsManager1.Tools("Add to Favorites").SharedProps.Enabled = True
            UltraToolbarsManager1.Tools("Sort Favorites").SharedProps.Enabled = False
            UltraToolbarsManager1.Tools("Rename").SharedProps.Enabled = False
            UltraToolbarsManager1.Tools("Delete").SharedProps.Enabled = False
            'UltraToolbarsManager1.Tools("Launch").SharedProps.Enabled = True ' .Caption = "nonfav" '.Visible = Infragistics.Win.DefaultableBoolean.True
        End If
    End Sub

    Private Sub UltraToolbarsManager1_ToolClick _
    (ByVal sender As System.Object,
     ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs) Handles UltraToolbarsManager1.ToolClick


        If e.Tool.OwnerIsMenu AndAlso e.Tool.OwningMenu.Key = "Help" Then

            Dim txt As String = ""
            ASCMAIN1.sql = "Select ASTHLPV1.* from ASTHLPV1,ASTHLPV2 " _
            & " where ASTHLPV1.VIDEO_NO = ASTHLPV2.VIDEO_NO " _
            & "   and ASTHLPV2.FORM_NAME = '" & e.Tool.SharedProps.Tag & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim FILE As String = row.Item("VIDEO_FILENAME")
                If InStr(FILE, ".") <> 0 Then
                    Dim PATH As String = "R:\" & ASCMAIN1.SOLUTION & "\Videos"
                    If ASCMAIN1.Running_in_VS Then
                        PATH = "C:\Documents and Settings\wjz\My Documents\Camtasia Studio"
                    End If

                    Dim VIDEO As String = PATH & "\" & Split(FILE, ".")(0) & "\" & FILE
                    If txt = "" Then txt &= "Videos: <br>"
                    txt &= "<a href='" & VIDEO & "' title='" & row.Item("VIDEO_NOTES") & "' target='_blank'>" & row.Item("VIDEO_DESC") & "</a><br>"
                End If
            Next

            Dim frm As ASFBASE1 = DirectCast(ASCMAIN1.ABS_FORMS(Val(e.Tool.Key)), ASFBASE1)
            If frm.dst.Tables("ASTTTIP1") Is Nothing Then Exit Sub
            With frm.dst.Tables("ASTTTIP1")
                Dim rowASTTTIP1 As DataRow = .Rows.Find _
                (New String() {e.Tool.SharedProps.Tag, "*", "*"})
                If rowASTTTIP1 IsNot Nothing Then
                    txt = "<b>" & "<font size=+1>" & rowASTTTIP1.Item("TOOLTIP_TITLE") & "</font>" & "</b><br/><br/>" & rowASTTTIP1.Item("TOOLTIP_TEXT")
                End If

                txt &= "<hr NoShade=true style=color:blue size=1px/>"

                For Each rowASTTTIP1 In .Select("TABLE_NAME <> '*'", "TABLE_NAME,COLUMN_NAME")
                    'If rowASTTTIP1.Item("TABLE_NAME") <> "*" Then
                    '    txt &= vbCrLf & rowASTTTIP1.Item("TOOLTIP_TITLE") _
                    '    & vbCrLf & vbCrLf & rowASTTTIP1.Item("TOOLTIP_TEXT")
                    'End If
                    txt &= "<br/><b>" & rowASTTTIP1.Item("TOOLTIP_TITLE") & "</b>" _
                    & "<br/><br/>" & rowASTTTIP1.Item("TOOLTIP_TEXT") & "<br/>"

                Next

            End With

            Dim frmASFMSGBF As New ASFMSGBF
            frmASFMSGBF.Show_Formatted_txt _
            ("Help for " & e.Tool.SharedProps.Caption, txt, Me)
            Exit Sub
        End If


        Select Case e.Tool.Key
            Case "History"    ' ComboBoxTool
                ASCMAIN1.MRU_used = True
                Dim comboboxtool As UltraWinToolbars.ComboBoxTool = UltraToolbarsManager1.Tools("History")
                ASCMAIN1.MRU_txtctl.Text = comboboxtool.Text
                SendKeys.Send(Chr(13))


            Case "Show Total of Selected Grid Cells"
                Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If Not sbt.Checked Then
                    ASCMAIN1.Progress("")
                End If

            Case "Spell Check"
                Me.UltraSpellChecker1.ShowSpellCheckDialog(ASCMAIN1.MRU_txtctl) ' (grdPMTTIME1.ActiveCell.EditorResolved)

            Case "Change Password"

                Dim f As New ASFPWDC1
                f.ShowDialog()
                f.Dispose()

            Case "Change Icon"
                Dim FILENAME As String = ""

                Using openFileDialog1 As New OpenFileDialog
                    'openFileDialog1.InitialDirectory = "C:\ABS\icons\iconexperience\48x48\plain\"
                    openFileDialog1.Title = "Select a .BMP file"
                    openFileDialog1.Filter = "png files (*.png)|*.png|bmp files (*.bmp)|*.bmp|ico files (*.ico)|*.ico|All files (*.*)|*.*"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True

                    If openFileDialog1.ShowDialog() = DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                    End If
                End Using

                If FILENAME = "OK" Then
                    Dim IMAGE_FILENAME As String = My.Computer.FileSystem.GetName(FILENAME)
                    Dim IMAGE_FOLDER As String = My.Computer.FileSystem.GetParentPath(FILENAME) & "\"

                    Dim ICO_FILENAME As String = ASCMAIN1.USER_ID & ".ICO"
                    Dim BMP_FILENAME As String = ASCMAIN1.USER_ID & ".BMP"
                    Dim objBmp As New Bitmap(FILENAME)

                    Me.UltraPictureBox1.Image = Nothing

                    My.Computer.FileSystem.CreateDirectory(ASCMAIN1.Folders("Images") & "Users\")
                    objBmp.Save(ASCMAIN1.Folders("Images") & "Users\" & ICO_FILENAME, System.Drawing.Imaging.ImageFormat.Icon)
                    Try
                        My.Computer.FileSystem.DeleteFile(ASCMAIN1.Folders("Images") & "Users\" & BMP_FILENAME)
                    Catch ex As Exception

                    End Try

                    Try
                        objBmp.Save(ASCMAIN1.Folders("Images") & "Users\" & BMP_FILENAME, System.Drawing.Imaging.ImageFormat.Bmp)
                    Catch ex As Exception

                    End Try

                    'objBmp.Dispose()
                    'objBmp = Nothing

                    Load_User_Icon()

                End If

            Case "Show Folders"

                Dim folders As String = "" _
                & "Images: " & ASCMAIN1.Folders("Images") & vbCr _
                & "Reports: " & ASCMAIN1.Folders("Reports") & vbCr _
                & "DataSets: " & ASCMAIN1.Folders("DataSets") & vbCr _
                & "Temp: " & ASCMAIN1.Folders("Temp") & vbCr _
                & "Archive: " & ASCMAIN1.Folders("Archive") & vbCr _
                & "Work: " & ASCMAIN1.Folders("Work") & vbCr _
                & "Help: " & ASCMAIN1.Folders("Help") & vbCr _
                & "bin: " & ASCMAIN1.Folders("bin") & vbCr _
                & "Attach: " & ASCMAIN1.Folders("Attach") & vbCr _
                & "Oracle: " & ASCMAIN1.Folders("Oracle") & vbCr _
                & "SharedRoot: " & ASCMAIN1.Folders("SharedRoot") & vbCr _
                & "sftp: " & ASCMAIN1.Folders("sftp") & vbCr

                MsgBox(folders, MsgBoxStyle.OkOnly, "Folder Paths")


            Case "Show Printers"

                Dim PRINTER_NAMEs As String = ""
                PRINTER_NAMEs &= vbCrLf & "Session ID from API: " & CStr(ASCMAIN1.WTS_SESSION_ID)

                For Each PRINTER_NAME As String In
                System.Drawing.Printing.PrinterSettings.InstalledPrinters
                    PRINTER_NAMEs &= vbCrLf & PRINTER_NAME
                Next
                MsgBox(PRINTER_NAMEs, MsgBoxStyle.OkOnly, "Printers on the System")

            Case "txtMenu"

            Case "Hide Menu after Selection"
                Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(UltraToolbarsManager1.Tools("Hide Menu after Selection"), UltraWinToolbars.StateButtonTool)
                SaveSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFMAIN1.HIDE_MENU_AFTER_SELECTION", sbt.Checked)

            Case "Enable Change to Favorite Descriptions"
                Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(UltraToolbarsManager1.Tools("Enable Change to Favorite Descriptions"), UltraWinToolbars.StateButtonTool)
                SaveSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFMAIN1.ENABLE_CHANGE_TO_FAVORITE_DESCRIPTIONS", sbt.Checked)

                If exbASTMENU1.Groups.Contains("*") Then
                    Try
                        Dim tvw As UltraWinTree.UltraTree = DirectCast(exbASTMENU1.Groups("*").Container.Controls(0), UltraWinTree.UltraTree)

                        If sbt.Checked Then
                            tvw.Override.LabelEdit = Infragistics.Win.DefaultableBoolean.True
                        Else
                            tvw.Override.LabelEdit = Infragistics.Win.DefaultableBoolean.False
                        End If

                    Catch ex As Exception

                    End Try
                End If

            Case "Small Font"
                Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(UltraToolbarsManager1.Tools("Small Font"), UltraWinToolbars.StateButtonTool)
                If sbt.Tag <> "X" Then
                    Dim F As System.Windows.Forms.Form = ASCMAIN1.ActiveForm
                    If F Is Nothing Then
                        F = Me
                    End If

                    If sbt.Checked Then
                        F.Font = New Font(Me.Font.FontFamily, 8, Me.Font.Style)
                    Else
                        F.Font = New Font(Me.Font.FontFamily, 9.75, Me.Font.Style)
                    End If
                End If

            Case "Most Recently Used"
                ASCMAIN1.MRU_used = True
                ASCMAIN1.MRU_txtctl.Text = e.ListToolItem.Key
                SendKeys.Send(Chr(13))

            Case "Menu"
                If Not UltraDockManager1.PaneFromControl(Me.exbASTMENU1).IsVisible Then
                    'Stop
                    UltraDockManager1.PaneFromControl(Me.exbASTMENU1).Show()
                Else
                    Me.Cursor = Cursors.WaitCursor

                    UltraDockManager1.PaneFromControl(Me.exbASTMENU1).Close()
                    exbASTMENU1.Groups.Clear()
                    exbASTMENU1.Controls.Clear()
                    Call Build_Menu()
                    UltraDockManager1.PaneFromControl(Me.exbASTMENU1).Show()
                    UltraDockManager1.PaneFromControl(Me.exbASTMENU1).Activate()

                    Me.Cursor = Cursors.Default
                End If

            Case "Write XML"

                ASCMAIN1.ActiveForm.Write_DataSet(True)
                MsgBox(".xml file has been prepared")

            Case "Export Meta-Data"
                Dim dst2 As New DataSet
                dst2.Tables.Add(ASCMAIN1.ActiveForm.dst.Tables("ASTTTIP1").Copy)
                'ASCMAIN1.ActiveForm.ExportMetaData(True)
                Dim FILE_NAME As String = ASCMAIN1.ActiveForm.Name

                Dim FOLDER_NAME As String = ASCMAIN1.Folders("Work") & "..\" & Mid(FILE_NAME, 1, 2) & "\" & IIf(Mid(FILE_NAME, 3, 1) = "F", "Forms", IIf(Mid(FILE_NAME, 3, 1) = "R", "Reports", "Tables")) & "\"
                Dim FORM_FILE_NAME As String = FOLDER_NAME & FILE_NAME & ".vb"
                If My.Computer.FileSystem.FileExists(FORM_FILE_NAME) Then
                    FOLDER_NAME = My.Computer.FileSystem.GetParentPath(FORM_FILE_NAME)
                    dst2.WriteXml(FOLDER_NAME & FILE_NAME & ".xml", XmlWriteMode.WriteSchema)
                    MsgBox("Meta-Data file " & FOLDER_NAME & FILE_NAME & " has been Exported")
                Else
                    MsgBox("Cannot Locate Form for " & FORM_FILE_NAME, MsgBoxStyle.OkOnly, "Export Not Performed")
                End If



            Case "Import Meta-Data"

                'ASCMAIN1.ActiveForm.ImportMetaData(True)
                'MsgBox(".xml file has been prepared")

            Case "Attachments"

                Dim F As New ASFATTA1
                F.TABLE_NAME = "GLTJRNL1"
                F.COLUMN_NAME = "JOURNAL_NO"
                F.CODE_VALUE = "000000"
                'F.DESC_VALUE = "000000"
                'F.ATTACHMENT_NOTES = "These are the notes"

                F.ShowDialog()
                F.Dispose()



            Case "Printer Setup"

                Dim result As DialogResult = PrintDialog1.ShowDialog()

            Case "Print Screen"
                'PrintForm1.Form = Me
                'PrintForm1.Print()

                Try

                    img = CaptureForm1()
                    pd = New System.Drawing.Printing.PrintDocument

                    'pd.Print()

                    Dim ppDialog As PrintPreviewDialog = New PrintPreviewDialog()
                    ppDialog.ClientSize = New Size(400, 500)
                    ppDialog.Document = pd
                    ppDialog.ShowDialog()

                    pd = Nothing
                    img = Nothing

                Catch ex As Exception
                    MsgBox("Please verify that the Default Printer is available", MsgBoxStyle.OkOnly, "Cannot Select Default Printer")
                End Try

            Case "Print Screen to PNG", "Capture Screen"

                Try
                    If e.Tool.Key = "Print Screen to PNG" Then
                        img = Clipboard.GetData(System.Windows.Forms.DataFormats.Bitmap)
                    Else
                        img = CaptureForm1()
                    End If

                    pd = New System.Drawing.Printing.PrintDocument
                    Dim p As Process = Nothing

                    ' this code comes from Show_Document - why not put it into ASCMAIN1 so ABS can see it?

                    Dim PRINT_SCREEN As String = ""
                    If e.Tool.Key = "Print Screen to PNG" Then
                        Dim FORM_NAME_current As String = ASCMAIN1.ActiveForm.Name
                        If FORM_NAME_current = "" Then FORM_NAME_current = "ASFMAIN1"
                        PRINT_SCREEN = ASCMAIN1.Folders("Help") & "png\" & FORM_NAME_current & ".png"
                    Else
                        PRINT_SCREEN = ASCMAIN1.Folders("Temp") & ASCMAIN1.Next_Control_No("PRINT_SCREEN") & ".bmp"
                    End If

                    img.Save(PRINT_SCREEN)

                    If My.Computer.FileSystem.FileExists(PRINT_SCREEN) Then
                        p = Process.Start(PRINT_SCREEN)
                        p.Dispose()
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Trying to Capture Screen")
                Finally

                End Try

                pd = Nothing
                img = Nothing

            Case "Multi-Task Conflict Control"

                Dim f As New ASFMTSK1
                f.ShowDialog()
                f.Dispose()

            Case "Launch"
                Call Launch_Menu_Item(ActiveNode)

            Case "Rename"
                ActiveNode.BeginEdit()
                'Stop

            Case "DataSet Viewer"
                Dim frmASFDSET1 As New ASFDSET1(ASCMAIN1.ActiveForm)
                frmASFDSET1.Show()


            Case "Deploy"
                ASFDEPL1.Show()

            Case "Security"
                ' ASFCSEC1.Show
                Dim frmASFCSEC1 As New ASFCSEC1
                frmASFCSEC1.Show()



            Case "Export to SQL Insert"
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    'ASCMAIN1.ActiveForm.Export_rows_to_SQL_custom()
                End If

            Case "Test"
                'Dim f As New ASFTEST1
                'f.Show()

                'MENU_ID = "TA"
                'MENU_ITEM_TYPE = "F"
                'MENU_ITEM_OBJECT = "TAFTEST1"
                'MENU_ITEM_DESC = e.Tool.CaptionResolved
                'MENU_ITEM_SECURITY = ""
                'MENU_ITEM_PP = ""
                'MENU_ITEM_EOM_CHECK = ""
                'MENU_ITEM_STANDALONE = ""
                'MENU_ITEM_FORM = ""
                'MODULE_ID = "TA"
                'MENU_ITEM_PASSWORD = ""
                'Call Launch_Form()
                'Exit Sub

                'MENU_ID = "AR"
                'MENU_ITEM_TYPE = "F"
                'MENU_ITEM_OBJECT = "ARFMEMO1"
                'MENU_ITEM_DESC = "Memo Entry"
                'MENU_ITEM_SECURITY = ""
                'MENU_ITEM_PP = ""
                'MENU_ITEM_EOM_CHECK = ""
                'MENU_ITEM_STANDALONE = ""
                'MENU_ITEM_FORM = ""
                'MODULE_ID = "AR"
                'MENU_ITEM_PASSWORD = ""
                'Call Launch_Form()

                'ASFTEST1.Show()


            Case "Sort Favorites"
                ActiveNode.ParentNodesCollection.Override.Sort = Infragistics.Win.UltraWinTree.SortType.Ascending
                ActiveNode.ParentNodesCollection.Override.Sort = Infragistics.Win.UltraWinTree.SortType.None

            Case "Update DLL"
                If ASCMAIN1.Running_in_VS Then
                    Stop ' you must be testing
                End If

                Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode
                If ActiveNode IsNot Nothing Then
                    Dim MENU_ITEM_TYPE As String = Split(ActiveNode.Tag, Chr(1))(1)
                    Dim MENU_ITEM_OBJECT As String = Split(ActiveNode.Tag, Chr(1))(2)
                    If MENU_ITEM_TYPE = "F" _
                    Or MENU_ITEM_TYPE = "T" _
                    Or MENU_ITEM_TYPE = "R" Then
                        Dim MODULE_ID As String = Mid(MENU_ITEM_OBJECT, 1, 2)
                        For Each open_form As System.Windows.Forms.Form In ASCMAIN1.ABS_FORMS
                            If open_form.Name.StartsWith(MODULE_ID) Then
                                MsgBox("You Must First Log Out of All Forms belonging to the " & MODULE_ID & " Module",
                                       MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                                Exit Sub
                            End If
                        Next

                        Dim FILENAME As String = MODULE_ID & ".DLL"
                        Dim PATH As String = ASCMAIN1.Folders("bin") ' Application.ExecutablePath
                        Dim SOURCE_FOLDER As String = ASCMAIN1.Folders("SharedRoot") & "bin\"
                        Try
                            FileCopy(SOURCE_FOLDER & FILENAME, PATH & FILENAME)
                            MsgBox("Update Successful (" & MODULE_ID & " Module)", MsgBoxStyle.OkOnly, "Verification")

                        Catch ex As Exception
                            MsgBox("Error Occurred attempting to Update " & FILENAME, MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                        End Try
                    End If
                End If


            Case "Add to Favorites"

                Dim t As Infragistics.Win.UltraWinTree.UltraTree = DirectCast(exbASTMENU1.Groups("*").Container.Controls(0), Infragistics.Win.UltraWinTree.UltraTree)
                If (t.GetNodeByKey(ActiveNode.Tag)) IsNot Nothing Then
                    MsgBox("Cannot Add '" & ActiveNode.Text & "'" & " to the Favorites Group" & vbCr & " because it already exists in the Favorites Group", MsgBoxStyle.OkOnly, "Cannot Proceed")
                Else
                    Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode
                    aNode.Text = ActiveNode.Text
                    aNode.Key = ActiveNode.Tag
                    aNode.Tag = ActiveNode.Tag
                    aNode.LeftImages.Add(ActiveNode.LeftImages.Item(0))
                    t.Nodes.Add(aNode)

                    Dim tblASTMENU2 As New DataTable
                    With ASCDATA1.GetDataAdapter(tblASTMENU2, "ASTMENU2", "*", True, -1, False)
                        Dim row As DataRow = tblASTMENU2.NewRow
                        row.Item("USER_ID") = ASCMAIN1.USER_ID
                        row.Item("MENU_ID") = Split(aNode.Tag, Chr(1))(0)
                        row.Item("MENU_ITEM_TYPE") = Split(aNode.Tag, Chr(1))(1)
                        row.Item("MENU_ITEM_OBJECT") = Split(aNode.Tag, Chr(1))(2)
                        row.Item("MENU_ITEM_DESC") = aNode.Text
                        tblASTMENU2.Rows.Add(row)
                        Try
                            .Update(tblASTMENU2)
                        Catch ex As Exception
                        End Try
                        .Dispose()
                    End With


                End If

            Case "Delete"
                Call Delete_Favorite(ActiveNode)
                ActiveNode.Remove()

            Case "Reports Archive"
                MENU_ID = "AS"
                MENU_ITEM_TYPE = "F"
                MENU_ITEM_OBJECT = "ASFSPRF1"
                MENU_ITEM_DESC = e.Tool.CaptionResolved
                MENU_ITEM_SECURITY = ""
                MENU_ITEM_PP = ""
                MENU_ITEM_EOM_CHECK = ""
                MENU_ITEM_STANDALONE = ""
                MENU_ITEM_FORM = ""
                MODULE_ID = "AS"
                MENU_ITEM_PASSWORD = ""
                Launch_Form()

            Case "Publish Documents to Portal"
                Using f As New ASFSRPTV
                    f.publish_documents = True
                    f.ShowDialog()
                    f.Dispose()
                End Using

            Case "App Stylist"
                Me.AppStylistRuntime1.ShowRuntimeApplicationStylingEditor(Me, "ABS1.isl")

            Case "Generate Help TOC"
                Call Generate_Help_TOC()

            Case "Show Tool-Tips"
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    Try
                        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                        ASCMAIN1.ActiveForm.tip.Enabled = tlb_sbt.Checked
                    Catch ex As Exception

                    End Try
                End If
            Case "Developer Mode"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    ASCMAIN1.developerMode = True
                Else
                    ASCMAIN1.developerMode = False
                End If
            Case "Show Developer"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    Dim FAILED_ATTEMPTS As Int32 = Val(tlb_sbt.Tag)
                    Dim frmASFMSGBF As New ASFMSGBF
                    Dim T As Int32 = Val(Now.Second)
                    Dim M As Int32 = Val(Now.Month)
                    Dim D As Int32 = Val(Now.Day)

                    If frmASFMSGBF.Get_txt_from_User("Enter Secret Password (T*M-D)", "Developer Options (" & CStr(T) & ")", True) = CStr(T * M - D) Then
                        If FAILED_ATTEMPTS > 2 Then
                            tlb_sbt.Checked = False
                        End If
                    Else
                        tlb_sbt.Checked = False
                        If frmASFMSGBF.user_option <> -1 Then
                            Dim SQL As String = "Insert into ASTEVNT1 " _
                            & " (USER_ID, FORM_NAME, INIT_OPER, INIT_DATE, EVENT_DESC) " _
                            & " VALUES (:PARM1, :PARM2, :PARM3, SYSDATE, :PARM4)"
                            ASCDATA1.ExecuteSQL(SQL, "VVVV", New Object() _
                                                {ASCMAIN1.USER_ID, Me.Name, ASCMAIN1.USER_ID, "Developer Options - Failed Attempt"})
                            tlb_sbt.Tag = CStr(FAILED_ATTEMPTS + 1)
                        End If
                    End If
                    If tlb_sbt.Checked Or ASCMAIN1.USER_ID = "wjz" Then
                        Dim SQL As String = "Insert into ASTEVNT1 " _
                        & " (USER_ID, FORM_NAME, INIT_OPER, INIT_DATE, EVENT_DESC) " _
                        & " VALUES (:PARM1, :PARM2, :PARM3, SYSDATE, :PARM4)"

                        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
                            SQL = "Insert into ASTEVNT1 " _
                            & " (USER_ID, FORM_NAME, INIT_OPER, INIT_DATE, EVENT_DESC) " _
                            & " VALUES (@PARM1, @PARM2, @PARM3, GETDATE(), @PARM4)"
                        End If


                        ASCDATA1.ExecuteSQL(SQL, "VVVV", New Object() _
                                            {ASCMAIN1.USER_ID, Me.Name, ASCMAIN1.USER_ID, "Developer Options Enabled"})

                        UltraToolbarsManager1.Tools("Developer").SharedProps.Visible = True
                    End If
                Else
                    UltraToolbarsManager1.Tools("Developer").SharedProps.Visible = False
                End If

            Case "Exit"

                Me.Close()

        End Select
    End Sub

    Sub Launch_Single_Form()

        UltraDockManager1.DockAreas(0).Close()

        Dim rowASTMENU1 As DataRow

        Dim sql As String
        Dim rows() As DataRow

        sql = "MENU_ITEM_TYPE in ('F','R','T')" _
        & " and MENU_ITEM_OBJECT = '" & ASCMAIN1.USER_MENU_ITEM_OBJECT & "'" _
        & "  and MENU_ID like '" & Mid(ASCMAIN1.USER_MENU_ITEM_OBJECT, 1, 2) & "%'"
        rows = tblASTMENU1.Select(sql, "")
        If rows.Length = 0 Then

            sql = "MENU_ITEM_TYPE in ('F','R','T')" _
            & " and MENU_ITEM_OBJECT = '" & ASCMAIN1.USER_MENU_ITEM_OBJECT & "'"
            rows = tblASTMENU1.Select(sql, "")

            If rows.Length = 0 Then
                MsgBox("Cannot Find Menu Item Object " & ASCMAIN1.USER_MENU_ITEM_OBJECT, MsgBoxStyle.OkOnly, "Closing Down Application")
                End
            End If
        End If


        rowASTMENU1 = rows(0)

        MENU_ID = rowASTMENU1.Item("MENU_ID") & ""
        MENU_ITEM_TYPE = rowASTMENU1.Item("MENU_ITEM_TYPE") & ""
        MENU_ITEM_OBJECT = ASCMAIN1.USER_MENU_ITEM_OBJECT
        MENU_ITEM_DESC = rowASTMENU1.Item("MENU_ITEM_DESC") & ""
        MENU_ITEM_SECURITY = rowASTMENU1.Item("MENU_ITEM_SECURITY") & ""
        MENU_ITEM_PP = rowASTMENU1.Item("MENU_ITEM_PP") & ""
        MENU_ITEM_EOM_CHECK = rowASTMENU1.Item("MENU_ITEM_EOM_CHECK") & ""
        MENU_ITEM_STANDALONE = rowASTMENU1.Item("MENU_ITEM_STANDALONE") & ""
        MENU_ITEM_FORM = rowASTMENU1.Item("MENU_ITEM_FORM") & ""
        MODULE_ID = Mid(ASCMAIN1.USER_MENU_ITEM_OBJECT, 1, 2)
        MENU_ITEM_PASSWORD = rowASTMENU1.Item("MENU_ITEM_PASSWORD") & ""
        Call Launch_Form()

    End Sub

    Sub Build_Menu()
        ASCMAIN1.MENU_ITEM_OBJECTs = New List(Of String)
        ASCMAIN1.tblASTMENU1.Rows.Clear()

        tblASTMENU1 = New DataTable

        If ASCMAIN1.DBS_TYPE = ASCMAIN1.DBS_TYPE_types.SQLServer Then
            ASCMAIN1.sql = "Select ASTMENU1.*" _
                     & ", ASTMENU2.MENU_ITEM_DESC MENU_ITEM_DESC_FAVORITE " _
                     & " from ASTMENU1 LEFT OUTER JOIN ASTMENU2 ON " _
                     & "  ASTMENU1.MENU_ID = ASTMENU2.MENU_ID " _
                     & "   and ASTMENU1.MENU_ITEM_TYPE = ASTMENU2.MENU_ITEM_TYPE " _
                     & "   and ASTMENU1.MENU_ITEM_OBJECT = ASTMENU2.MENU_ITEM_OBJECT " _
                     & "   and ASTMENU2.USER_ID  = '" & ASCMAIN1.USER_ID & "'" _
                     & "   where ISNULL(ASTMENU1.MENU_ITEM_HIDDEN,'0') = '0'"

        Else
            ASCMAIN1.sql = "Select ASTMENU1.*" _
                     & ", ASTMENU2.MENU_ITEM_DESC MENU_ITEM_DESC_FAVORITE " _
                     & " from ASTMENU1,ASTMENU2" _
                     & " where ASTMENU1.MENU_ID = ASTMENU2.MENU_ID (+)" _
                     & "   and ASTMENU1.MENU_ITEM_TYPE = ASTMENU2.MENU_ITEM_TYPE (+)" _
                     & "   and ASTMENU1.MENU_ITEM_OBJECT = ASTMENU2.MENU_ITEM_OBJECT (+)" _
                     & "   and ASTMENU2.USER_ID (+) = '" & ASCMAIN1.USER_ID & "'" _
                     & "   and NVL(ASTMENU1.MENU_ITEM_HIDDEN,'0') = '0'"
        End If

        'If Not ASCMAIN1.Running_in_VS Then
        ' P = PRODUCTION (ALSO NULL) {BOTH}
        ' T = TEST (ONLY IF RUNNING ON TEST MACHINE) {V2 ONLY} - V1 MUST IGNORE T
        ' D = DEVELOPMENT (ONLY IF RUNNING IN VS) {V1 ONLY}
        ' NEEDS TO BE EXPANDED UPON TO MATCH TEST ENVIRONMENT TNSNAMES OR COMPANY
        ASCMAIN1.sql &= "and (NVL(ASTMENU1.MENU_ITEM_STATUS,'P') = 'P' OR NVL(ASTMENU1.MENU_ITEM_STATUS,'P') = 'T')"
        'End If

        tblASTMENU1 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ASTMAIN1", 3)
        For Each r As DataRow In tblASTMENU1.Select("MENU_ITEM_SECURITY is Not Null")
            Dim z As String = ""
            Dim access_denied As Boolean = True
            Dim MENU_ITEM_SECURITY As String = r.Item("MENU_ITEM_SECURITY")
            For I As Integer = 1 To Len(MENU_ITEM_SECURITY) / 2
                z = Mid$(MENU_ITEM_SECURITY, (I - 1) * 2 + 1, 2)
                If InStr(ASCMAIN1.USER_SECURITY_CODEs, z) <> 0 Then
                    access_denied = False
                    Exit For
                End If
            Next
            If access_denied Then
                r.Delete()
            End If
        Next

        tblASTMENU1.AcceptChanges()
        ASCMAIN1.tblASTMENU1 = tblASTMENU1.Clone

        Add_Menu_to_ExplorerBar("MAIN", "M" & Chr(1) & "MAIN" & Chr(0), tblASTMENU1)
    End Sub

    Sub Add_Menu_to_ExplorerBar(
        ByVal MENU_ID As String,
        ByVal KEY_PREFIX As String,
        ByRef tblASTMENU1 As DataTable)

        Dim KEY As String
        Dim MENU_ITEM_DESC As String
        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Groups\"
        Dim IMAGE_FILE As String
        Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode

        KEY = "*"
        MENU_ITEM_DESC = "Favorites"
        IMAGE_FILE = "FAVORITES.PNG"

        Add_Group(KEY, MENU_ITEM_DESC, IMAGE_FOLDER, IMAGE_FILE, tvw)
        Dim tvwFavorites As Infragistics.Win.UltraWinTree.UltraTree = tvw
        Dim KP As String = KEY_PREFIX & KEY & Chr(0)
        ' ASCMAIN1.Add_Menu_to_Tree("*", KEY_PREFIX & KEY & Chr(0), tvw, 0, tblASTMENU1)
        tvw.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True

        Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(UltraToolbarsManager1.Tools("Enable Change to Favorite Descriptions"), UltraWinToolbars.StateButtonTool)
        If sbt.Checked Then
            tvw.Override.LabelEdit = Infragistics.Win.DefaultableBoolean.True
        Else
            tvw.Override.LabelEdit = Infragistics.Win.DefaultableBoolean.False
        End If

        For Each row As DataRow In tblASTMENU1.Select("MENU_ID = '" & MENU_ID & "' and MENU_ITEM_TYPE = 'M'", "MENU_ITEM_SEQ")
            KEY = row.Item("MENU_ITEM_TYPE") & Chr(1) & row.Item("MENU_ITEM_OBJECT")
            MENU_ITEM_DESC = row.Item("MENU_ITEM_DESC")
            IMAGE_FILE = row.Item("MENU_ITEM_OBJECT") & ".PNG"
            Add_Group(KEY, MENU_ITEM_DESC, IMAGE_FOLDER, IMAGE_FILE, tvw)
            ASCMAIN1.Add_Menu_to_Tree(row.Item("MENU_ITEM_OBJECT"), KEY_PREFIX & KEY & Chr(0), tvw, 0, tblASTMENU1)
        Next


        ' now add the favorites
        ASCMAIN1.Add_Menu_to_Tree("*", KP, tvwFavorites, 0, tblASTMENU1)


    End Sub

    Sub tvw_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        tvw.SelectedNodes.Clear()
        Dim anode As Infragistics.Win.UltraWinTree.UltraTreeNode = tvw.GetNodeFromPoint(e.Location)
        If anode IsNot Nothing Then
            anode.Selected = True
            tvw.ActiveNode = anode
        Else

        End If
    End Sub

    Sub tvw_AfterLabelEdit(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.NodeEventArgs)

        Dim tblASTMENU2 As New DataTable
        ASCMAIN1.sql = "Select * from ASTMENU2 " _
                     & " where USER_ID = '" & ASCMAIN1.USER_ID & "'" _
                     & " and MENU_ID = '" & Split(e.TreeNode.Tag, Chr(1))(0) & "'" _
                     & " and MENU_ITEM_TYPE = '" & Split(e.TreeNode.Tag, Chr(1))(1) & "'" _
                     & " and MENU_ITEM_OBJECT = '" & Split(e.TreeNode.Tag, Chr(1))(2) & "'"

        With ASCDATA1.GetDataAdapter(tblASTMENU2, "ASTMENU2", ASCMAIN1.sql, True, -1)
            If tblASTMENU2.Rows.Count = 0 Then
                Dim row As DataRow = tblASTMENU2.Rows(0)
                row.Item("MENU_ITEM_DESC") = e.TreeNode.Text

                .Update(tblASTMENU2)
                .Dispose()
            End If
        End With
    End Sub

    Sub tvw_BeforeDelete(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.BeforeNodesDeletedEventArgs)
        dNode = e.Nodes(0).Clone
    End Sub

    Sub Delete_Favorite(ByVal aNode As Infragistics.Win.UltraWinTree.UltraTreeNode)

        Dim tblASTMENU2 As New DataTable

        ASCMAIN1.sql = "Select * from ASTMENU2 " _
                     & " where USER_ID = '" & ASCMAIN1.USER_ID & "'" _
                     & " and MENU_ID = '" & Split(aNode.Tag, Chr(1))(0) & "'" _
                     & " and MENU_ITEM_TYPE = '" & Split(aNode.Tag, Chr(1))(1) & "'" _
                     & " and MENU_ITEM_OBJECT = '" & Split(aNode.Tag, Chr(1))(2) & "'"

        With ASCDATA1.GetDataAdapter(tblASTMENU2, "ASTMENU2", ASCMAIN1.sql, True, -1)
            Dim row As DataRow = tblASTMENU2.Rows(0)
            row.Delete()
            .Update(tblASTMENU2)
            .Dispose()
        End With
    End Sub

    Sub tvw_AfterDelete(ByVal sender As Object, ByVal e As System.EventArgs)
        Call Delete_Favorite(dNode)
    End Sub

    Sub Launch_Menu_Item(ByVal aNode As Infragistics.Win.UltraWinTree.UltraTreeNode)
        MENU_ID = Split(aNode.Tag, Chr(1))(0)
        MENU_ITEM_TYPE = Split(aNode.Tag, Chr(1))(1)
        MENU_ITEM_OBJECT = Split(aNode.Tag, Chr(1))(2)

        ASCMAIN1.MENU_ID = MENU_ID
        ASCMAIN1.MENU_ITEM_TYPE = MENU_ITEM_TYPE
        ASCMAIN1.MENU_ITEM_OBJECT = MENU_ITEM_OBJECT

        Dim rowASTMENU1 As DataRow = tblASTMENU1.Rows.Find(New String() {MENU_ID, MENU_ITEM_TYPE, MENU_ITEM_OBJECT})
        ' THIS IS NEC TO BLOCK USERS WHO ARE ALREADY LOGGED IN
        ASCMAIN1.sql = "Select * from ASTMENU1 where MENU_ID = '" & MENU_ID & "' and MENU_ITEM_TYPE = '" & MENU_ITEM_TYPE & "' and MENU_ITEM_OBJECT = '" & MENU_ITEM_OBJECT & "'"
        rowASTMENU1 = ASCDATA1.GetDataRow

        If MENU_ITEM_TYPE = "M" Then
            Exit Sub
        End If

        MENU_ITEM_DESC = aNode.Text
        MENU_ITEM_SECURITY = rowASTMENU1.Item("MENU_ITEM_SECURITY") & ""
        MODULE_ID = Mid(MENU_ITEM_OBJECT, 1, 2)

        MENU_ITEM_PP = rowASTMENU1.Item("MENU_ITEM_PP") & ""
        MENU_ITEM_EOM_CHECK = rowASTMENU1.Item("MENU_ITEM_EOM_CHECK") & ""
        MENU_ITEM_STANDALONE = rowASTMENU1.Item("MENU_ITEM_STANDALONE") & ""
        MENU_ITEM_FORM = rowASTMENU1.Item("MENU_ITEM_FORM") & ""
        MENU_ITEM_PASSWORD = rowASTMENU1.Item("MENU_ITEM_PASSWORD") & ""

        Call Launch_Form()

        Dim sbt As UltraWinToolbars.StateButtonTool = DirectCast(UltraToolbarsManager1.Tools("Hide Menu after Selection"), UltraWinToolbars.StateButtonTool)
        If sbt.Checked Then
            UltraDockManager1.UnpinAll()
            'UltraDockManager1.PaneFromControl(Me.exbASTMENU1).ToggleDockState()
            UltraDockManager1.HideAll()

            Application.DoEvents()
            'UltraDockManager1.Visible = True
            'UltraDockManager1.PaneFromControl(Me.exbASTMENU1).Show()
            'UltraDockManager1.FlyoutPane.Close()
        End If
    End Sub

    Function Launch_Form(Optional ByVal rowASTMENU1 As DataRow = Nothing) As ASFBASE1

        Launch_Form = Nothing
        ASCMAIN1.Timer = Now

        Try
            If rowASTMENU1 IsNot Nothing Then
                MENU_ID = rowASTMENU1.Item("MENU_ID") & ""
                MENU_ITEM_TYPE = rowASTMENU1.Item("MENU_ITEM_TYPE") & ""
                MENU_ITEM_OBJECT = rowASTMENU1.Item("MENU_ITEM_OBJECT") & ""
                MENU_ITEM_DESC = rowASTMENU1.Item("MENU_ITEM_DESC") & ""
                MENU_ITEM_SECURITY = rowASTMENU1.Item("MENU_ITEM_SECURITY") & ""
                MENU_ITEM_PP = rowASTMENU1.Item("MENU_ITEM_PP") & ""
                MENU_ITEM_EOM_CHECK = rowASTMENU1.Item("MENU_ITEM_EOM_CHECK") & ""
                MENU_ITEM_STANDALONE = rowASTMENU1.Item("MENU_ITEM_STANDALONE") & ""
                MENU_ITEM_FORM = rowASTMENU1.Item("MENU_ITEM_FORM") & ""
                MODULE_ID = Mid(MENU_ITEM_OBJECT, 1, 2)
                MENU_ITEM_PASSWORD = rowASTMENU1.Item("MENU_ITEM_PASSWORD") & ""
            End If

            ASCMAIN1.tblASTSQLX1 = Nothing
            Call ASCMAIN1.Get_Current_YP()


            'UltraPictureBox1.Visible = False
            'Application.DoEvents()
            'Dim splashThread As Thread = New Thread(New ThreadStart(AddressOf StartSplash))
            'splashThread.Start()

            If MENU_ITEM_PASSWORD <> "" Then
                Dim frmASFMSGBF As New ASFMSGBF
                If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
                Else
                    If MENU_ITEM_PASSWORD <> frmASFMSGBF.Get_txt_from_User _
                    ("Enter Password", "Password Required for " & MENU_ITEM_DESC, True) Then

                        MsgBox("Incorrect Password Entered", MsgBoxStyle.OkOnly, "Cannot Perform Requested Function")
                        Exit Function
                    End If
                End If
            End If

            If MENU_ITEM_EOM_CHECK = "1" And ASCMAIN1.EOM = "1" Then
                MsgBox("Period-End is in Progress", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Exit Function
            End If

            UltraPictureBox1.Visible = False
            Application.DoEvents()
            Dim splashThread As Thread = New Thread(New ThreadStart(AddressOf StartSplash))
            splashThread.Start()

            ' Create MT Definitions for Period_End Inititalization

            If MENU_ITEM_OBJECT = "TARPEND0" Then
                Dim sql As String
                sql = "Delete from ASTMTKC1 " _
                & " where MENU_ITEM_TYPE = 'R' and MENU_ITEM_OBJECT = 'TARPEND0'"
                ASCDATA1.ExecuteSQL(sql)
                sql = "Insert into ASTMTKC1 " _
                & " Select Distinct 'R' MENU_ITEM_TYPE, 'TARPEND0' MENU_ITEM_OBJECT, " _
                & " MENU_ITEM_TYPE ENTITY_TYPE, MENU_ITEM_OBJECT ENTITY, " _
                & " 'L' MT_ACTION" _
                & " from ASTMENU1 where MENU_ITEM_EOM_CHECK = '1' or MENU_ITEM_OBJECT in ('TARPEND0','TARPEND1')"
                ASCDATA1.ExecuteSQL(sql)
                sql = "Insert into ASTMTKC1 " _
                & " Select Distinct 'R' MENU_ITEM_TYPE, 'TARPEND0' MENU_ITEM_OBJECT, " _
                & " 'SERVICE' ENTITY_TYPE, SERVICE_OBJECT ENTITY, " _
                & " 'L' MT_ACTION" _
                & " from ASTSVCC1 where SERVICE_EOM_CHECK = '1'"
                ASCDATA1.ExecuteSQL(sql)

            End If

            Call ASCMAIN1.Multi_Task_Cleanup()
            If Not ASCMAIN1.Multi_Task_Menu_Item(MENU_ITEM_TYPE, MENU_ITEM_OBJECT, 1, True, MENU_ITEM_STANDALONE) Then

                If MENU_ITEM_OBJECT = "TARPEND0" Then
                    If ASCMAIN1.Running_in_VS Then ' probably don't want people having this kind of power - ODG does this with Ralph at 2AM, so there is an exception
                        If MsgBox("Users are still logged into applications" _
                              & vbCr & " which conflict with running Period End Initialization" & vbCr _
                              & vbCr & "Would you like to Clear Multi-Task Conflicts?",
                              MsgBoxStyle.YesNo,
                              "There are Multi-Tasking Conflicts") = MsgBoxResult.Yes Then
                            Dim f As New ASFMTSK1
                            f.ShowDialog()
                            f.Dispose()
                        End If
                    End If
                End If

                Exit Function
            End If


            If ASCMAIN1.EOM = "1" Then
                Dim SQL As String = ""
                SQL = "SELECT ASTOPST1.MENU_ITEM_OBJECT, ASTOPST1.INIT_DATE "
                SQL = SQL & ", ASTOPST1.USER_ID, ASTMENU1.MENU_ITEM_DESC"
                SQL = SQL & " FROM ASTOPST1,ASTMENU1 "
                SQL = SQL & " WHERE ASTOPST1.YYYYPP = '" & ASCMAIN1.CYP & "'"
                SQL = SQL & " AND ASTOPST1.PRD_CLOSE_IND = '1' AND ASTOPST1.UPDATED = '1'"
                SQL = SQL & " AND ASTMENU1.MENU_ID = ASTOPST1.MENU_ID"
                SQL = SQL & " AND ASTMENU1.MENU_ITEM_TYPE = ASTOPST1.MENU_ITEM_TYPE"
                SQL = SQL & " AND ASTMENU1.MENU_ITEM_OBJECT = ASTOPST1.MENU_ITEM_OBJECT"
                SQL = SQL & " AND ASTOPST1.MENU_ITEM_TYPE = '" & MENU_ITEM_TYPE & "'"
                SQL = SQL & " AND ASTOPST1.MENU_ITEM_OBJECT = '" & MENU_ITEM_OBJECT & "'"
                Dim z As String = ""
                For Each row As DataRow In ASCDATA1.GetDataTable(SQL).Rows
                    z = z & vbCr & row.Item("MENU_ITEM_OBJECT") _
                          & vbTab & row.Item("INIT_DATE") _
                          & vbTab & row.Item("USER_ID") _
                          & vbTab & row.Item("MENU_ITEM_DESC")
                Next
                If z <> "" Then
                    If vbNo = MsgBox(z & vbCr & vbCr & "Continue Anyway?", vbQuestion + vbYesNo, "This Report/Function has already been Updated since Period-End has been Initialized") Then
                        ASCMAIN1.Multi_Task_Menu_Item(MENU_ITEM_TYPE, MENU_ITEM_OBJECT, -1, True)
                        Exit Function
                    End If
                End If
            End If

            Try

                Me.Cursor = Cursors.WaitCursor

                Dim sLocation As String
                Dim buildType As String

#If DEBUG Then
                buildType = "x64\Debug"
#Else
                buildType = "x86\Release"
#End If
                Dim TICKS As Long = Now.Ticks
                Dim TIMES As String = ""

                Dim t1 As Date = Now

                Dim FormToShow As ASFBASE1 = Nothing
                RaiseEvent ShowForm(Me, FormToShow)

                If FormToShow Is Nothing Then
                    If ASCMAIN1.Running_in_VS Then
                        sLocation = ASCMAIN1.Folders("root") & MODULE_ID & "\bin\" & buildType & "\" & MODULE_ID & ".dll"
                    Else
                        sLocation = ASCMAIN1.Folders("bin") & MODULE_ID & ".dll"
                    End If

                    Dim sType As String = MODULE_ID & "." & MENU_ITEM_OBJECT
                    If MENU_ITEM_FORM <> "" Then
                        sType = MODULE_ID & "." & MENU_ITEM_FORM
                    End If

                    If Not ASCMAIN1.ABS_Assemblies.ContainsKey(MODULE_ID) Then
                        If InStr(MODULE_IDs, MODULE_ID) = 0 Then
                            MODULE_IDs &= "," & MODULE_ID
                            If Mid(MODULE_IDs, 1, 1) = "," Then
                                MODULE_IDs = Mid(MODULE_IDs, 2)
                            End If
                            SaveSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFMAIN1.MODULE_IDs", MODULE_IDs)
                        End If
                        ASCMAIN1.ABS_Assemblies.Add(MODULE_ID, System.Reflection.Assembly.LoadFrom(sLocation))
                    End If
                    Dim formAsm As System.Reflection.Assembly = ASCMAIN1.ABS_Assemblies(MODULE_ID)
                    Dim ClassType As Type = formAsm.GetType(sType)
                    TIMES &= vbCrLf & Format(Now.Ticks - TICKS, "###,###,##0")
                    'Dim FormToShow As ASFBASE1
                    FormToShow = DirectCast(Activator.CreateInstance(ClassType), ASFBASE1)

                End If

                TIMES &= vbCrLf & Format(Now.Ticks - TICKS, "###,###,##0")

                FormToShow.Text = MENU_ITEM_DESC

                FormToShow.MENU_ID = MENU_ID
                FormToShow.MENU_ITEM_TYPE = MENU_ITEM_TYPE
                FormToShow.MENU_ITEM_OBJECT = MENU_ITEM_OBJECT
                FormToShow.MENU_ITEM_DESC = MENU_ITEM_DESC
                FormToShow.MENU_ITEM_SECURITY = MENU_ITEM_SECURITY
                FormToShow.MENU_ITEM_PP = MENU_ITEM_PP
                FormToShow.MENU_ITEM_FORM = MENU_ITEM_FORM
                FormToShow.MODULE_ID = MODULE_ID
                FormToShow.Name = MENU_ITEM_OBJECT

                FormToShow.MdiParent = Me
                FormToShow.WindowState = FormWindowState.Normal

                Dim t2 As Int32 = DateDiff("s", t1, Now)
                'RND WAS REMARKED
                'ASCMAIN1.ActiveForm = FormToShow
                FormToShow.Show()
                If FormToShow.error_has_occured IsNot Nothing Then
                    FormToShow.Close()
                    MsgBox(FormToShow.error_has_occured.Message, MsgBoxStyle.OkOnly, "Error Launching Form " & FormToShow.Name)
                    Exit Function
                    Stop
                End If
                UltraTabbedMdiManager1.ActiveTab.Text = MENU_ITEM_DESC

                TIMES &= vbCrLf & Format(Now.Ticks - TICKS, "###,###,##0")

                Launch_Form = FormToShow

                UltraPictureBox1.Visible = False

                Dim t3 As Int32 = DateDiff("s", t1, Now)
                UltraStatusBar1.Panels(7).Text = CStr(t2) & "/" & CStr(t3)

                'If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then MsgBox(TIMES)

            Catch ex As Exception
                ASCMAIN1.Multi_Task_Menu_Item(MENU_ITEM_TYPE, MENU_ITEM_OBJECT, -1, True)

                If ASCMAIN1.Running_in_VS Then
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Launching Form from the Menu")
                    MsgBox(Mid(ex.StackTrace, 1, 500))
                    MsgBox(Mid(ex.GetBaseException.ToString, 1, 500))
                    'MsgBox(Mid(ex.InnerException.Message, 1, 500))
                    'MsgBox(Mid(ex.Data, 1, 500))
                Else
                    MsgBox("Menu Item (" & MENU_ID & ":" & MENU_ITEM_TYPE & ":" & MENU_ITEM_OBJECT & ") is not set up for use.  Please contact ABS")
                    'If ASCMAIN1.USER_ID = "wjz" Then
                    '    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error Launching Form from the Menu")
                    '    MsgBox(Mid(ex.StackTrace, 1, 500))
                    '    MsgBox(Mid(ex.GetBaseException.ToString, 1, 500))
                    '    'MsgBox(Mid(ex.InnerException.Message, 1, 500))
                    '    'MsgBox(Mid(ex.Data, 1, 500))
                    'End If
                End If

                'Finally
                '    Me.Cursor = Cursors.Default
                '    CloseSplash()

            End Try

        Catch ex As Exception

        Finally
            Me.Cursor = Cursors.Default
            CloseSplash()

        End Try

    End Function

    Sub tvw_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs)
        If exbASTMENU1.Tag <> "" Then Exit Sub
        exbASTMENU1.Tag = "*"

        Try
            Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
            Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
            Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

            If tnode IsNot Nothing Then
                Dim tree As Infragistics.Win.UltraWinTree.UltraTree = DirectCast(sender, Infragistics.Win.UltraWinTree.UltraTree)
                If tnode.Equals(tree.ActiveNode) Then
                    Call Launch_Menu_Item(DirectCast(sender, Infragistics.Win.UltraWinTree.UltraTree).ActiveNode)
                End If
            End If

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then
                Stop ' THIS USUALLY MEANS THAT AN UNHANDLED EXCEPTION OCCURRED
                '  AS THE FORM LOADED OR CYCLED THROUGH CLEAR RECORD OR ELSE IN MODES
                ' PUT A F9 IN EACH OF THOSE PLACES (MOST LIKELY Mode_Settings)
            Else
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    ASCMAIN1.ActiveForm.Close()
                End If
            End If
        End Try

        exbASTMENU1.Tag = ""
    End Sub

    Private Sub UltraStatusBar1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles UltraStatusBar1.DragDrop
        If ASCMAIN1.ActiveForm IsNot Nothing Then
            ASCMAIN1.ActiveForm.Dropped_On(e)
        End If
    End Sub

    Private Sub UltraStatusBar1_PanelClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinStatusBar.PanelClickEventArgs) Handles UltraStatusBar1.PanelClick

        Select Case e.Panel.Key
            Case "MENU_ITEM_OBJECT"
                If e.Panel.Text <> "ASFMAIN1" Then
                    Dim f As New ASFXHST1
                    f.Width = Me.Width * 0.8
                    f.Height = Me.Height * 0.8
                    f.ShowDialog()
                    f.Dispose()
                End If

            Case "AUDIT"
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    ASCMAIN1.ActiveForm.Show_Audit()
                End If

            Case "ATTACH"
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    ASCMAIN1.ActiveForm.Show_Attachments()
                End If

            Case "LOG"
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    ASCMAIN1.ActiveForm.Show_Log()
                End If

            Case "DATA"
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    ASCMAIN1.ActiveForm.Show_Data()
                End If

            Case "EVENTS"
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    ASCMAIN1.ActiveForm.Show_Events()
                End If
        End Select
    End Sub

    Private Sub btlFindIcon_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btlFindIcon.Click
        Dim FILENAME As String = ""

        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.InitialDirectory = "C:\ABS\icons\iconexperience\48x48\plain\"
            openFileDialog1.Title = "Select an icon to Associate with " & grdImages.ActiveRow.Cells("ITEM_KEY").Text
            openFileDialog1.Filter = "png files (*.png)|*.png|All files (*.*)|*.*"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            If InStr(FILENAME, "\48x48\plain\") = 0 Then
                MsgBox("You Must Select an Icon from the 48x48\plain folder", MsgBoxStyle.OkOnly, "Cannot Resolve Icon")
                Exit Sub
            End If
            Try
                My.Computer.FileSystem.CopyFile(Replace(FILENAME, "\48x48\plain\", "\32x32\plain\"), ASCMAIN1.Folders("Images") & "32\" & grdImages.ActiveRow.Cells("ITEM_KEY").Text & ".png", True)
                My.Computer.FileSystem.CopyFile(Replace(FILENAME, "\48x48\plain\", "\16x16\plain\"), ASCMAIN1.Folders("Images") & "16\" & grdImages.ActiveRow.Cells("ITEM_KEY").Text & ".png", True)


                For Each G As UltraWinExplorerBar.UltraExplorerBarGroup In ASCMAIN1.ActiveForm.UltraExplorerBar1.Groups
                    For Each i As UltraWinExplorerBar.UltraExplorerBarItem In G.Items
                        If i.Key = grdImages.ActiveRow.Cells("ITEM_KEY").Text Then
                            i.Settings.AppearancesLarge.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "32\", i.Key)
                            i.Settings.AppearancesSmall.Appearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", i.Key)
                        End If
                    Next
                Next

            Catch ex As Exception
                MsgBox("Exception Occurred", MsgBoxStyle.OkOnly, "Maybe illegal characters in File Name")
            End Try
        End If

    End Sub

    Sub Generate_Help_TOC()

        Dim TOC As String
        TOC = "Table of Contents"

        Dim HelpFolder As String = My.Computer.FileSystem.GetDirectoryInfo(ASCMAIN1.Folders("Help")).FullName
        Dim TOCnew As String = HelpFolder & TOC & ".hhc"

        Using chmWriter As New System.IO.StreamWriter(TOCnew)
            Dim TOCold As String = HelpFolder & TOC & ".orig"
            Using tocReader As New System.IO.StreamReader(TOCold)
                While tocReader.Peek <> -1
                    Dim z As String = tocReader.ReadLine()
                    If z = "</UL>" Then
                        chmWriter.WriteLine(vbTab & "<UL>")
                        Call WriteNodes("MAIN", "Main Menu", chmWriter, 1)
                        chmWriter.WriteLine(vbTab & "</UL>")
                    End If
                    chmWriter.WriteLine(z)
                End While

            End Using
        End Using

        System.Diagnostics.Process.Start("C:\Program Files (x86)\HTML Help Workshop\hhc.exe", HelpFolder & ASCMAIN1.SOLUTION & ".hhp")
        MsgBox("Click OK when Compile is Done", vbOKOnly, "Waiting ...")
        System.Diagnostics.Process.Start("hh", HelpFolder & ASCMAIN1.SOLUTION & ".chm")

    End Sub

    Sub WriteNodes(ByVal MENU_ID As String,
    ByVal MENU_DESC As String,
    ByVal chmWriter As System.IO.StreamWriter,
    ByVal lvl As Integer)

        Dim vbQuo As String = Chr(34)
        Dim tabs As String = "".PadLeft(lvl + 1, vbTab)
        chmWriter.WriteLine(tabs & "<LI>" & vbTab & "<OBJECT type=" & vbQuo & "text/sitemap" & vbQuo & ">")
        chmWriter.WriteLine(tabs & vbTab & "<param name=" & vbQuo & "Name" & vbQuo & " value=" & vbQuo & MENU_DESC & vbQuo & ">")
        chmWriter.WriteLine(tabs & vbTab & "</OBJECT>")
        chmWriter.WriteLine(tabs & "<UL>")

        ASCMAIN1.sql = "Select * from ASTMENU1 where MENU_ID = '" & MENU_ID & "' and NVL(MENU_ITEM_HIDDEN,'0') <> '1'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "MENU_ITEM_SEQ")
            Dim MENU_ITEM_TYPE As String = row.Item("MENU_ITEM_TYPE")
            Dim MENU_ITEM_OBJECT As String = row.Item("MENU_ITEM_OBJECT")
            Dim MENU_ITEM_DESC As String = row.Item("MENU_ITEM_DESC")

            If MENU_ITEM_TYPE = "M" Then
                Call WriteNodes(MENU_ITEM_OBJECT, MENU_ITEM_DESC, chmWriter, lvl + 1)
            Else

                Dim image_value As String = ""
                If MENU_ITEM_TYPE = "M" Then
                    image_value = "6"
                ElseIf MENU_ITEM_TYPE = "F" Then
                    image_value = "27"
                ElseIf MENU_ITEM_TYPE = "R" Then
                    image_value = "11"
                ElseIf MENU_ITEM_TYPE = "T" Then
                    image_value = "39"
                End If

                Dim j As Integer = 1
                Do While InStr(j, MENU_ITEM_DESC, "&") <> 0
                    Dim i As Integer = InStr(MENU_ITEM_DESC, "&")
                    j = i + 1
                    MENU_ITEM_DESC = Mid$(MENU_ITEM_DESC, 1, i - 1) & "&amp;" & Mid$(MENU_ITEM_DESC, i + 1)
                Loop

                chmWriter.WriteLine(tabs & vbTab & "<LI>" & vbTab & "<OBJECT type=" & vbQuo & "text/sitemap" & vbQuo & ">")
                chmWriter.WriteLine(tabs & vbTab & vbTab & "<param name=" & vbQuo & "Name" & vbQuo & " value=" & vbQuo & MENU_ITEM_DESC & vbQuo & ">")
                chmWriter.WriteLine(tabs & vbTab & vbTab & "<param name=" & vbQuo & "Local" & vbQuo & " value=" & vbQuo & ".\html\" & MENU_ITEM_OBJECT & ".htm" & vbQuo & ">")
                chmWriter.WriteLine(tabs & vbTab & vbTab & "<param name=" & vbQuo & "ImageNumber" & vbQuo & " value=" & vbQuo & image_value & vbQuo & ">")
                chmWriter.WriteLine(tabs & vbTab & vbTab & "</OBJECT>")
            End If
        Next

        chmWriter.WriteLine(tabs & vbTab & "</UL>")

    End Sub


    Sub Write_Nodes(ByVal anode As Infragistics.Win.UltraWinTree.UltraTreeNode,
    ByVal T As Infragistics.Win.UltraWinTree.UltraTree,
    ByVal chmWriter As System.IO.StreamWriter,
    ByVal lvl As Integer)


        Dim i As Integer
        Dim j As Integer
        Dim z As String
        Dim i1 As String
        Dim i2 As String
        Dim i0 As String

        Dim NAME As String
        Dim vbQuo As String = Chr(34)
        Dim tabs As String = "".PadLeft(lvl, vbTab)

        lvl = 1

        z = anode.Key
        i1 = Split(z, Chr(1))(0 + 0)
        i2 = Split(z, Chr(1))(0 + 1)
        If i1 = "M" Then
            i0 = "MENU_" & i2
        Else
            i0 = Mid$(i2, 1, 2) & i1 & Mid$(i2, 4)
        End If

        NAME = anode.Text

        Dim image_value As String = ""
        If i1 = "M" Then
            image_value = "6"
        ElseIf i1 = "F" Then
            image_value = "27"
        ElseIf i1 = "R" Then
            image_value = "11"
        ElseIf i1 = "T" Then
            image_value = "39"
        End If

        j = 1
        Do While InStr(j, NAME, "&") <> 0
            i = InStr(NAME, "&")
            j = i + 1
            NAME = Mid$(NAME, 1, i - 1) & "&amp;" & Mid$(NAME, i + 1)
        Loop

        'chmWriter.WriteLine(tabs & "<UL>")
        chmWriter.WriteLine(tabs & vbTab & "<LI> <OBJECT type=" & vbQuo & "text/sitemap" & vbQuo & ">")
        chmWriter.WriteLine(tabs & vbTab & vbTab & "<param name=" & vbQuo & "Name" & vbQuo & " value=" & vbQuo & NAME & vbQuo & ">")
        chmWriter.WriteLine(tabs & vbTab & vbTab & "<param name=" & vbQuo & "Local" & vbQuo & " value=" & vbQuo & ".\html\" & i0 & ".htm" & vbQuo & ">")
        chmWriter.WriteLine(tabs & vbTab & vbTab & "<param name=" & vbQuo & "ImageNumber" & vbQuo & " value=" & vbQuo & image_value & vbQuo & ">")
        chmWriter.WriteLine(tabs & vbTab & vbTab & "</OBJECT>")

        If anode.HasNodes Then
            chmWriter.WriteLine(tabs & "<UL>")
            For Each cNode As Infragistics.Win.UltraWinTree.UltraTreeNode In anode.Nodes
                Call Write_Nodes(cNode, T, chmWriter, lvl + 1)
            Next
            chmWriter.WriteLine(tabs & "</UL>")
        End If

    End Sub

    Sub Load_User_Icon()
        Try
            'Dim objBmp As New Bitmap(ASCMAIN1.Folders("Images") & ASCMAIN1.USER_ID & ".bmp")
            'Me.UltraPictureBox1.Image = objBmp
            'objBmp = Nothing

            UltraPictureBox1.Visible = True
            UltraPictureBox1.Appearance.BackColor = Color.Transparent

            Dim RCX As Int32 = UltraPictureBox1.Left + UltraPictureBox1.Width
            Dim RCY As Int32 = UltraPictureBox1.Top + UltraPictureBox1.Height


            If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Images") & "Users\" & ASCMAIN1.USER_ID & ".bmp") Then
                Me.UltraPictureBox1.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "Users\", ASCMAIN1.USER_ID & ".bmp")
            Else
                Me.UltraPictureBox1.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "ABS\", ASCMAIN1.DBS_COMPANY & ".bmp")
            End If

            UltraPictureBox1.AutoSize = True

            Dim MAX_Height As Integer = 400
            If UltraPictureBox1.Height > MAX_Height Then
                Dim H As Integer = UltraPictureBox1.Height
                Dim W As Integer = UltraPictureBox1.Width
                UltraPictureBox1.AutoSize = False
                UltraPictureBox1.Height = H * MAX_Height / H
                UltraPictureBox1.Width = W * MAX_Height / H
            End If

            Dim MAX_Width As Integer = 600
            If UltraPictureBox1.Width > MAX_Width Then
                Dim H As Integer = UltraPictureBox1.Height
                Dim W As Integer = UltraPictureBox1.Width
                UltraPictureBox1.AutoSize = False
                UltraPictureBox1.Height = H * MAX_Width / W
                UltraPictureBox1.Width = W * MAX_Width / W
            End If

            UltraPictureBox1.Left = RCX - UltraPictureBox1.Width
            UltraPictureBox1.Top = RCY - UltraPictureBox1.Height

            'UltraPictureBox1.ScaleImage = ScaleImage.Always
            'UltraToolbarsManager1.Tools("Change Icon").SharedProps.AppearancesSmall.Appearance.Image = objBmp
        Catch ex As Exception
            UltraPictureBox1.Visible = False
        End Try
    End Sub

    Sub ChildFormDisposed(ByVal sender As Object, ByVal e As EventArgs)

    End Sub

    'Imports System.Runtime.InteropServices
    ' Global Variables 
    Dim img As Bitmap
    Dim WithEvents pd As System.Drawing.Printing.PrintDocument

    'Returns the Form as a bitmap
    Public Function CaptureForm1() As Bitmap

        Dim g1 As Graphics = Me.CreateGraphics()
        Dim MyImage = New Bitmap(Me.ClientRectangle.Width, Me.ClientRectangle.Height, g1)

        Dim g2 As Graphics = Graphics.FromImage(MyImage)
        Dim dc1 As IntPtr = g1.GetHdc()
        Dim dc2 As IntPtr = g2.GetHdc()
        BitBlt(dc2, 0, 0, Me.ClientRectangle.Width, (Me.ClientRectangle.Height), dc1, 0, 0, 13369376)
        g1.ReleaseHdc(dc1)
        g2.ReleaseHdc(dc2)
        'saves image to c drive just, u can comment it also
        'MyImage.Save("c:\abc.bmp")
        Return MyImage
    End Function

    <System.Runtime.InteropServices.DllImport("gdi32.DLL", EntryPoint:="BitBlt",
    SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Unicode,
    ExactSpelling:=True,
    CallingConvention:=System.Runtime.InteropServices.CallingConvention.StdCall)>
    Private Shared Function BitBlt(ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As System.Int32) As Boolean

        ' Leave function empty - DLLImport attribute forwards calls to MoveFile to
        ' MoveFileW in KERNEL32.DLL.
    End Function

    Private Sub pd_QueryPageSettings(ByVal sender _
    As Object, ByVal e As _
    System.Drawing.Printing.QueryPageSettingsEventArgs) _
    Handles pd.QueryPageSettings
        e.PageSettings.Landscape = True
    End Sub

    'this method will be called each time when pd.printpage event occurs
    Sub pd_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles pd.PrintPage

        Dim x As Integer = e.MarginBounds.X '/ 2
        Dim y As Integer = e.MarginBounds.Y '/ 2
        'e.Graphics.DrawImage(img, x, y)

        'e.Graphics.DrawImage(img, 0, 0)

        'e.HasMorePages = False



        'local scope
        Dim mySource As Rectangle
        'Dim myDestination As Rectangle

        'define a rectangle as the size of the original image (source)
        mySource = New Rectangle(x:=x, y:=y, Width:=e.MarginBounds.Width, Height:=e.MarginBounds.Height)

        'draw the original bitmap to the source rectangle
        e.Graphics.DrawImage(image:=img, rect:=mySource)

        Dim ABS_logo As Bitmap = New System.Drawing.Bitmap(fileName:=ASCMAIN1.Folders("Images") & "abs\abs_logo.jpg")

        e.Graphics.DrawImage(image:=ABS_logo, rect:=New Rectangle(x:=0, y:=0, Width:=80, Height:=40))
        Dim p As New Pen(Color.Blue, 1)
        e.Graphics.DrawLine(p, 0, 45, e.PageBounds.Width, 45)


        Dim printFont As Font = Me.Font
        Dim myBrush As New SolidBrush(Color.Black)

        e.Graphics.DrawString(Me.Text & " " & Format(Now, "MM/dd/yyyy hh:mm tt") & " " & ASCMAIN1.USER_ID, printFont, myBrush, 0, e.PageBounds.Height - e.MarginBounds.Y + 20, New StringFormat())

    End Sub

    Public Event ShowForm(ByVal sender As Object, ByRef e As ASFBASE1)

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub StartLoadingDLLs()

        Dim buildType As String = ""

#If DEBUG Then
        buildType = "x64\Debug"
#Else
        buildType = "x64\Release"
#End If
        'RND WAS REMARKED
        MODULE_IDs = GetSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFMAIN1.MODULE_IDs")


        If 1 <> 1 Then

            Dim MODULE_IDs_Successful As String = ""
            If MODULE_IDs <> "" Then

                For Each MODULE_ID As String In Split(MODULE_IDs, ",")
                    Dim sLocation As String = ""

                    If ASCMAIN1.Running_in_VS Then
                        sLocation = ASCMAIN1.Folders("root") & MODULE_ID & "\bin\" & buildType & "\" & MODULE_ID & ".dll"
                    Else
                        sLocation = ASCMAIN1.Folders("bin") & MODULE_ID & ".dll"
                    End If

                    If Not ASCMAIN1.ABS_Assemblies.ContainsKey(MODULE_ID) Then
                        Try
                            ASCMAIN1.ABS_Assemblies.Add(MODULE_ID, System.Reflection.Assembly.LoadFrom(sLocation))
                            MODULE_IDs_Successful &= "," & MODULE_ID

                            If MODULE_ID = "AR" Then
                                Dim formAsm As System.Reflection.Assembly = ASCMAIN1.ABS_Assemblies(MODULE_ID)
                                Dim sType As String = MODULE_ID & "." & "ARFCINQ1"
                                Dim ClassType As Type = formAsm.GetType(sType)

                                FormToShow = DirectCast(Activator.CreateInstance(ClassType), ASFBASE1)
                            End If
                        Catch ex As Exception

                        End Try
                    End If
                Next
            End If

            If MODULE_IDs <> Mid(MODULE_IDs_Successful, 2) Then
                MODULE_IDs = Mid(MODULE_IDs_Successful, 2)
                SaveSetting(My.Application.Info.AssemblyName, ASCMAIN1.SOLUTION, "ASFMAIN1.MODULE_IDs", MODULE_IDs)
            End If

        End If

    End Sub

    Private Sub CloseStartLoadingDLLs()
        'If loadSplash Is Nothing Then
        '    Return
        'End If
        'loadSplash.Invoke(New EventHandler(AddressOf loadSplash.EndForm))
        'loadSplash.Dispose()
        'loadSplash = Nothing
    End Sub

    Private Sub UltraStatusBar1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles UltraStatusBar1.DragEnter
        e.Effect = DragDropEffects.All
    End Sub

    Sub Site_Specific_Settings()
        Dim buildType As String

#If DEBUG Then
        buildType = "x64\Debug"
#Else
        buildType = "x64\Release"
#End If

        Dim slocation As String = ""
        Dim MODULE_ID As String = "TAC"
        Dim CLASS_NAME As String = "TACMAIN1"
        If ASCMAIN1.Running_in_VS Then
            slocation = ASCMAIN1.Folders("root") & MODULE_ID & "\bin\" & buildType & "\" & MODULE_ID & ".dll"
        Else
            slocation = ASCMAIN1.Folders("bin") & MODULE_ID & ".dll"
        End If
        Dim formAsm As System.Reflection.Assembly = System.Reflection.Assembly.LoadFrom(slocation)

        Dim sType As String = MODULE_ID & "." & CLASS_NAME
        Dim ClassType As Type = formAsm.GetType(sType)

        ASCMAIN1.TACMAIN1 = DirectCast(Activator.CreateInstance(ClassType), TACMAIN1)
        ASCMAIN1.TACMAIN1.Site_Specific_Settings()
        ASCMAIN1.TACMAIN1.Application_Initialization()

    End Sub

#Region "Serial Ports"

    'Private Sub scaleSerialPort_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles scaleSerialPort.DataReceived
    '    Try
    '        Dim scaleData As String = String.Empty

    '        scaleData = scaleSerialPort.ReadExisting
    '        Me.Invoke(scaleWeightDelegate, New Object() {scaleData})
    '    Catch ex As Exception

    '    End Try

    'End Sub

    'Private Sub scaleSerialPort_ErrorReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialErrorReceivedEventArgs) Handles scaleSerialPort.ErrorReceived
    '    Exit Sub
    'End Sub

    Private Sub labelPrinterSerialPort_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles labelPrinterSerialPort.DataReceived
        Try

        Catch ex As Exception

        End Try

    End Sub

    Private Sub labelPrinterSerialPort_ErrorReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialErrorReceivedEventArgs) Handles labelPrinterSerialPort.ErrorReceived
        Exit Sub
    End Sub

#End Region


End Class