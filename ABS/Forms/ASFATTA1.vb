Imports Infragistics.Win
Imports Microsoft.Office.Interop

Public Class ASFATTA1
    Dim LAST_DATE As Date = Now + ASCMAIN1.NowTSD
    Dim EXT_allowed As List(Of String)
    'Dim tblASTATTA2 As DataTable
    Public tblASTATTA2 As DataTable

#Region "ABS Standard Routines"
    Private Sub ASFATTA1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        With dst

            If ENTITY.CUSTOM_SQL = "" Then
                ASCMAIN1.sql = "Select * from ASTATTA2 " & vbCrLf _
                & " where TABLE_NAME = '" & ENTITY.TABLE_NAME & "'" & vbCrLf _
                & " and COLUMN_NAME = '" & ENTITY.COLUMN_NAME & "'" & vbCrLf _
                & " and CODE_VALUE = '" & ENTITY.CODE_VALUE & "'" & vbCrLf _
                & " and NVL(ATTACHMENT_STATUS,'O') <> 'D'"
            Else
                ASCMAIN1.sql = ENTITY.CUSTOM_SQL
            End If

            If ENTITY.OTHER_ENTITIES IsNot Nothing Then
                For Each E_other As Dropped_On_Entity_Other In ENTITY.OTHER_ENTITIES
                    ASCMAIN1.sql &= " union " & vbCrLf _
                        & "Select ASTATTA2.* from ASTATTA2," & E_other.TABLE_NAME & vbCrLf _
                        & " where ASTATTA2.TABLE_NAME = '" & E_other.TABLE_NAME & "'" & vbCrLf _
                        & "   and ASTATTA2.COLUMN_NAME = '" & E_other.COLUMN_NAME & "'" & vbCrLf _
                        & "   and " & E_other.TABLE_NAME & "." & E_other.COLUMN_NAME & " = ASTATTA2.CODE_VALUE" & vbCrLf _
                        & "   and " & E_other.TABLE_NAME & "." & E_other.COLUMN_NAME_linked & " = '" & ENTITY.CODE_VALUE & "'" & vbCrLf _
                        & "   and NVL(ATTACHMENT_STATUS,'O') <> 'D'"
                Next
            End If
            Create_TDA(.Tables.Add, "ASTATTA2", "**", 0)
        End With

        grdASTATTA2.DataSource = dst.Tables("ASTATTA2")
        For Each GC As UltraWinGrid.UltraGridColumn In grdASTATTA2.DisplayLayout.Bands(0).Columns
            GC.Header.Appearance.BackColor2 = Color.LightBlue
            If GC.Key <> "ATTACHMENT_DESC" And GC.Key <> "ATTACHMENT_TYPE" Then
                GC.CellAppearance.BackColor = Color.LightYellow ' .Drawing.Color.FromArgb(255, 255, 255, 192)
            End If
            If GC.Key = "ATTACHMENT_ORIGINATOR" Or GC.Key = "ATTACHMENT_DATETIME" Then
                GC.Header.Appearance.BackColor2 = Color.LightBlue
            End If
            If GC.Key = "ATTACHMENT_FILENAME" Or GC.Key = "ATTACHMENT_EXT" Then
                GC.Header.Appearance.BackColor2 = Color.Gold
            End If
            If GC.Key = "INIT_OPER" Or GC.Key = "INIT_DATE" Then
                GC.Header.Appearance.BackColor2 = Color.LightBlue
            End If
            'GC.Header.Appearance.BackGradientStyle = GradientStyle.GlassTop20
            GC.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            GC.Header.Appearance.BackColor = Drawing.Color.White

        Next

        If tblASTATTA2 IsNot Nothing Then
            dst.Tables("ASTATTA2").Merge(tblASTATTA2, True)
        Else
            Fill_Records("ASTATTA2")
        End If

        ASCMAIN1.sql = "Select * from ASTATTA1"
        EXT_allowed = New List(Of String)
        For Each rowASTATTA1 As DataRow In ASCDATA1.GetDataTable.Rows
            EXT_allowed.Add(rowASTATTA1.Item("ATTACHMENT_EXT"))
        Next

        ASCMAIN1.sql = "Select COLUMN_NAME_DESC from ASTATTA0 " _
        & " where TABLE_NAME = '" & ENTITY.TABLE_NAME & "'" _
        & " and COLUMN_NAME = '" & ENTITY.COLUMN_NAME & "'"
        Dim COLUMN_NAME_DESC As String = ASCDATA1.GetDataValue

        If ENTITY.DESC_VALUE <> "" Then
            Me.Text = "Attachments for " & IIf(COLUMN_NAME_DESC <> "", COLUMN_NAME_DESC & " ", "") & ENTITY.CODE_VALUE & " - " & ENTITY.DESC_VALUE
        End If

        ASCMAIN1.sql = "Select ATTACHMENT_TYPE,ATTACHMENT_TYPE_DESC from ASTATTA3 " _
        & " where TABLE_NAME = '" & ENTITY.TABLE_NAME & "'" _
        & " and COLUMN_NAME = '" & ENTITY.COLUMN_NAME & "'"
        ASCMAIN1.Add_Value_List(grdASTATTA2, "ATTACHMENT_TYPE", , , , ASCMAIN1.sql)

        ASCMAIN1.Center(Me)

        If ENTITY.READ_ONLY Then
            grdASTATTA2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdASTATTA2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdASTATTA2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            cmdAdd.Visible = False
            cmdSave.Visible = False
            grdASTATTA2.AllowDrop = False

            cmdExit.Text = "Done"

        Else
            'If ENTITY.RESTRICTIONS.Contains("D") Then
            '    grdASTATTA2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            'End If
        End If

        If eDND IsNot Nothing Then
            Process_DragDrop()
        End If
        cmdSave.Visible = False
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdASTATTA2, "SS", "Show Filter", "Show GroupBox")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'Dim tlb_btn As UltraWinToolbars.ButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case ""
                    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                        e.Cancel = True
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case ""

        End Select
    End Sub

#End Region

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click

        If Not ENTITY.READ_ONLY Then
            If dst.Tables("ASTATTA2").Select("", "", DataViewRowState.ModifiedCurrent).Count <> 0 _
            Or dst.Tables("ASTATTA2").Select("", "", DataViewRowState.Added).Count <> 0 _
            Or dst.Tables("ASTATTA2").Select("", "", DataViewRowState.Deleted).Count <> 0 Then
                If MsgBox("Are you sure you want to Exit?", MsgBoxStyle.YesNo, "Changes have not been Saved") = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If
        End If

        Me.Close()
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.InitialDirectory = "c:\"
            openFileDialog1.Title = "Select a File to Attach to this Record"
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim Msg As String = Attach_File(openFileDialog1.FileName, , , , tblASTATTA2 Is Nothing)
                If Msg <> "" Then
                    MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                End If
            End If
        End Using
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        For Each row As DataRow In dst.Tables("ASTATTA2").Select("ATTACHMENT_DESC is Null or ATTACHMENT_DESC = ''")
            MsgBox("All Attachments must have a Description")
            Exit Sub
        Next
        'For Each row As DataRow In dst.Tables("ASTATTA2").Select("ATTACHMENT_TYPE is Null or ATTACHMENT_TYPE = ''")
        '    MsgBox("All Attachments must have a Type")
        '    Exit Sub
        'Next
        If tblASTATTA2 IsNot Nothing Then
            tblASTATTA2.Rows.Clear()
            tblASTATTA2.Merge(dst.Tables("ASTATTA2"), True)
        Else
            Update_Record_TDA("ASTATTA2")
            MsgBox("Attachments have been Saved", MsgBoxStyle.OkOnly, "Verification")
        End If

        Me.Close()
    End Sub

    Private Sub grdASTATTA2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTATTA2.AfterCellUpdate
        cmdSave.Visible = True
        grdASTATTA2.ActiveRow.Cells("LAST_DATE").Value = Now
        grdASTATTA2.ActiveRow.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID

    End Sub

    Private Sub grdASTATTA2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTATTA2.AfterRowActivate

        If grdASTATTA2.ActiveRow.IsDataRow Then
            With grdASTATTA2.DisplayLayout.Bands(0)
                If grdASTATTA2.ActiveRow.Cells("INIT_OPER").Value & "" = ASCMAIN1.USER_ID Then
                    .Columns("ATTACHMENT_DESC").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("ATTACHMENT_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
                'making this field editable by everyone as per Gary 8/11/2011
                .Columns("ATTACHMENT_TYPE").CellActivation = UltraWinGrid.Activation.AllowEdit
            End With
        End If
    End Sub

    Private Sub grdASTATTA2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTATTA2.ClickCellButton
        Dim ATTACHMENT_NO As String = grdASTATTA2.ActiveRow.Cells("ATTACHMENT_NO").Text
        Dim ATTACHMENT_EXT As String = grdASTATTA2.ActiveRow.Cells("ATTACHMENT_EXT").Text.ToUpper
        Call ASCMAIN1.Launch_Attachment(ATTACHMENT_NO, ATTACHMENT_EXT)
    End Sub

    Private Sub grdASTATTA2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTATTA2.InitializeRow

        Dim ATTACHMENT_EXT As String = e.Row.Cells("ATTACHMENT_EXT").Text.ToUpper
        Dim FILENAME As String = ASCMAIN1.Get_Filename(ATTACHMENT_EXT)

        If FILENAME <> "" Then
            e.Row.Cells("ATTACHMENT_EXT").ButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", FILENAME)
        End If

        Dim i As Int32 = e.Row.ListIndex
        Dim grd As UltraWinGrid.UltraGrid = DirectCast(sender, UltraWinGrid.UltraGrid)
        Select Case dst.Tables("ASTATTA2").Rows(i).RowState
            Case DataRowState.Added
                'e.Row.Appearance.BackColor = Color.LightGreen
                e.Row.RowSelectorAppearance.BackColor = Color.LightGreen
                e.Row.RowSelectorAppearance.BackColor2 = Color.Green
            Case DataRowState.Modified
                'e.Row.Appearance.BackColor = Color.LightSkyBlue
                e.Row.RowSelectorAppearance.BackColor = Color.LightSkyBlue
                e.Row.RowSelectorAppearance.BackColor2 = Color.Blue
        End Select
    End Sub

    Private Sub grdASTATTA2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles grdASTATTA2.KeyDown
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            If e.KeyValue = Keys.Enter Then
                .UpdateData()
            End If
        End With
    End Sub

    Overloads Sub Process_DragDrop()

        If ENTITY.READ_ONLY Then
            Exit Sub
        End If

        cmdAdd.Visible = False
        lblNowProcessing.Visible = True
        Application.DoEvents()

        Dim files() As String = eDND.Data.GetData(DataFormats.FileDrop)

        If files IsNot Nothing Then
            For Each FILENAME As String In files
                Dim Msg As String = Attach_File(FILENAME, , , , tblASTATTA2 Is Nothing)
                If Msg <> "" Then
                    MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                End If
            Next
        Else
            Try
                Dim outlook As Outlook.Application = CType(Microsoft.VisualBasic.Interaction.GetObject("", "Outlook.Application"), Outlook.Application)
                Dim explorer As Outlook.Explorer = outlook.ActiveExplorer

                For i As Int32 = 0 To explorer.Selection.Count - 1
                    Dim mail As Outlook.MailItem = CType(explorer.Selection.Item(i + 1), Outlook.MailItem)
                    mail.SaveAs(ASCMAIN1.Folders("Temp") & "mailitem.msg")

                    Dim FILENAME As String = ASCMAIN1.Folders("Temp") & "mailitem.msg"
                    Dim Msg As String = Attach_File(FILENAME, mail.Subject, mail.SenderName, mail.SentOn, tblASTATTA2 Is Nothing)
                    If Msg <> "" Then
                        MsgBox(Msg, MsgBoxStyle.OkOnly, "Error Attempting to Attach File ")
                    End If
                    mail = Nothing
                Next

                outlook = Nothing
                explorer = Nothing

            Catch ex As System.Exception

                MsgBox(ex, "Error - Outlook request not found")

            End Try

        End If

        lblNowProcessing.Visible = False
        cmdAdd.Visible = True
        Application.DoEvents()

        Me.Activate()
        'ASCMAIN1.ActiveForm.Activate()

    End Sub

    Private Sub grdASTATTA2_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdASTATTA2.DragEnter
        If grdASTATTA2.AllowDrop Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub grdASTATTA2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdASTATTA2.DragDrop
        lblNowProcessing.Visible = True
        cmdAdd.Visible = False
        eDND = e
        Process_DragDrop()
        lblNowProcessing.Visible = False
        cmdAdd.Visible = True
    End Sub

    Private Sub grdASTATTA2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTATTA2.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("INIT_OPER").Value & "" <> ASCMAIN1.USER_ID Then
                e.Cancel = True
                MsgBox("Cannot Delete Documents Attached by Others", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Exit For
            End If

            If grow.Cells("TABLE_NAME").Value & "" <> ENTITY.TABLE_NAME _
            Or grow.Cells("COLUMN_NAME").Value & "" <> ENTITY.COLUMN_NAME _
            Or grow.Cells("CODE_VALUE").Value & "" <> ENTITY.CODE_VALUE Then
                e.Cancel = True
                MsgBox("Cannot Delete Documents Attached from Other Entities (" _
                       & grow.Cells("TABLE_NAME").Value & ":" _
                       & grow.Cells("COLUMN_NAME").Value & ":" _
                       & grow.Cells("CODE_VALUE").Value & ")", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                Exit For
            End If

            Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").Rows.Find _
                (New Object() {grow.Cells("TABLE_NAME").Value, grow.Cells("COLUMN_NAME").Value, grow.Cells("CODE_VALUE").Value, grow.Cells("ATTACHMENT_NO").Value})

            If ENTITY.RESTRICTIONS IsNot Nothing Then
                If ENTITY.RESTRICTIONS.Contains("D") And rowASTATTA2.RowState <> DataRowState.Added Then
                    e.Cancel = True
                    MsgBox("Deletion of Attached Documents is Not Permitted at this time", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit For
                End If
            End If

            e.Cancel = True
            If MsgBox("Delete Attachment: " & grow.Cells("ATTACHMENT_DESC").Value & "?", MsgBoxStyle.YesNo, "Confirm Deletion") = MsgBoxResult.No Then
            Else
                Delete_Attachment(grow.Cells("TABLE_NAME").Value, grow.Cells("COLUMN_NAME").Value, grow.Cells("CODE_VALUE").Value, grow.Cells("ATTACHMENT_NO").Value)
            End If

        Next

        e.DisplayPromptMsg = False
    End Sub

    Private Sub grdASTATTA2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdASTATTA2.AfterRowUpdate
        cmdSave.Visible = True
    End Sub

    Private Sub grdASTATTA2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTATTA2.AfterRowsDeleted
        cmdSave.Visible = True
    End Sub

    Private Sub grdASTATTA2_AfterEnterEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTATTA2.AfterEnterEditMode
        cmdSave.Visible = True
    End Sub

    Sub Delete_Attachment(ByVal TABLE_NAME As String, ByVal COLUMN_NAME As String, ByVal CODE_VALUE As String, ByVal ATTACHMENT_NO As String)

        Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").Rows.Find(New Object() {TABLE_NAME, COLUMN_NAME, CODE_VALUE, ATTACHMENT_NO})

        rowASTATTA2.Item("ATTACHMENT_STATUS") = "D"
        rowASTATTA2.Item("LAST_DATE") = Now
        rowASTATTA2.Item("LAST_OPER") = ASCMAIN1.USER_ID

        Update_Record_TDA("ASTATTA2")

        If ENTITY.CUSTOM_SQL = "" Then
            ASCMAIN1.sql = "Select * from ASTATTA2 " _
            & " where TABLE_NAME = '" & ENTITY.TABLE_NAME & "'" _
            & " and COLUMN_NAME = '" & ENTITY.COLUMN_NAME & "'" _
            & " and CODE_VALUE = '" & ENTITY.CODE_VALUE & "'" _
            & " and NVL(ATTACHMENT_STATUS,'O') <> 'D'"
        Else
            ASCMAIN1.sql = ENTITY.CUSTOM_SQL
        End If

        Fill_Records("ASTATTA2", , True, ASCMAIN1.sql)

    End Sub

    Private Sub grdASTATTA2_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdASTATTA2.InitializeLayout

    End Sub
End Class