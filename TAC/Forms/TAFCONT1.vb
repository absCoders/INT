Public Class TAFCONT1

    Public result As Boolean = False
    Public allow_editing As Boolean = True
    Public CONTACT_ENTITY_TABLE As String
    Public CONTACT_ENTITY_KEY As String
    Public CONTACT_ENTITY_NAME As String
    Dim RowState_b4_checked As DataRowState

    Public Sub New( _
    ByVal FF As ASFBASE1)
        frmASFBASE1 = FF
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from TATCONT1 " _
            & " where CONTACT_ENTITY_TABLE = :PARM1 and CONTACT_ENTITY_KEY = :PARM2"
            Create_TDA(.Tables.Add, "TATCONT1", "**", 0, True, "VV", 1)
            .Tables("TATCONT1").Columns.Add("SEL")
            .Tables("TATCONT1").Columns("SEL").DefaultValue = "0"
        End With

        grdTATCONT1.DataSource = dst.Tables("TATCONT1")

        With grdTATCONT1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
            .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        End With
        Allow_Edit(False)
        With grdTATCONT1.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("CONTACT_NAME").Header.Fixed = True
        End With

        If CONTACT_ENTITY_TABLE = "" And CONTACT_ENTITY_KEY = "" Then
            ASCMAIN1.sql = "Select * from TATCONT1 where CONTACT_ENTITY_TABLE is Null and CONTACT_ENTITY_KEY is Null and INIT_OPER = '" & ASCMAIN1.USER_ID & "'"
            Fill_Records("TATCONT1", "", True, ASCMAIN1.sql)
        Else
            Fill_Records("TATCONT1", New String() {CONTACT_ENTITY_TABLE, CONTACT_ENTITY_KEY})
        End If
        TABLE_NAME = ""

        Me.Text = "Contacts" & IIf(CONTACT_ENTITY_KEY = "", "", " for " & CONTACT_ENTITY_NAME & " (" & CONTACT_ENTITY_KEY & ")")

        AUDIT.Add("TATCONT1", "*")

    End Sub

    Sub Allow_Edit(ByVal tf As Boolean)
        With grdTATCONT1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Or tf Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BackColor = Color.Empty
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If
            Next

            If tf Then
                grdTATCONT1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdTATCONT1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                grdTATCONT1.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
                grdTATCONT1.Tag = "X"
            Else
                grdTATCONT1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdTATCONT1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdTATCONT1.DisplayLayout.Bands(0).Columns("SEL").Hidden = False
                grdTATCONT1.Tag = ""
            End If
        End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdTATCONT1, "SSSS", "Show Filter", "Show GroupBox", "Show Pins", "Allow Editing")
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

        If tlb_pop.Tools.Exists("Allow Editing") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Allow Editing"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = allow_editing
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                'Case "grdSOTPICK2"
                '    If grd.DisplayLayout.Override.AllowUpdate <> DefaultableBoolean.True Then
                '        e.Cancel = True
                '    End If
            End Select

        End If
    End Sub


    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        'Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Allow Editing"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    If Not ASCMAIN1.Logical_Lock("TATCONT1", CONTACT_ENTITY_TABLE & ":" & CONTACT_ENTITY_KEY, , , , 99) Then
                        tlb_sbt.Checked = False
                        Exit Sub
                    End If
                End If
                Allow_Edit(tlb_sbt.Checked)

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "email Invoice"
            '    Dim INV_NO As String = grd.ActiveRow.Cells("INV_NO").Value & ""
            '    Dim FILENAME As String = Create_Invoice(INV_NO)
            '    email_Invoice(INV_NO, FILENAME)
        End Select
    End Sub
#End Region

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Exit_Form()
    End Sub

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        If chkSaveChanges.Checked Then
            INIT_LAST("TATCONT1", True)
            Update_Record_TDA("TATCONT1")
        End If
        result = True
        Exit_Form()
    End Sub

    Sub Exit_Form()
        ASCMAIN1.MultiTask_Release(, , 99)
        frmASFBASE1 = Nothing
        Me.Close()
    End Sub

    Private Sub grdTATCONT1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdTATCONT1.AfterRowsDeleted
        Setup_to_Save_Changes()
    End Sub

    Private Sub grdTATCONT1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdTATCONT1.AfterRowUpdate
        If grdTATCONT1.Tag = "X" Then
            Setup_to_Save_Changes()
        Else
            Dim rowTATCONT1 As DataRow = dst.Tables("TATCONT1").Rows.Find _
                (New Object() {e.Row.Cells("CONTACT_NO").Value})
            If rowTATCONT1.RowState = DataRowState.Modified And RowState_b4_checked = DataRowState.Unchanged Then
                rowTATCONT1.AcceptChanges()
            End If

            Dim SEND_TOs As String = Trim(txtSEND_TOs.Text) & ""
            If SEND_TOs.StartsWith(";") Then
                SEND_TOs = Mid(SEND_TOs, 2)
            End If
            If SEND_TOs.EndsWith(";") Then
                SEND_TOs = Mid(SEND_TOs, 1, Len(SEND_TOs) - 1)
            End If

            If e.Row.Cells("SEL").Value & "" = "1" Then
                If SEND_TOs = "" Then
                    SEND_TOs = e.Row.Cells("CONTACT_EMAIL").Value & ""
                Else
                    SEND_TOs &= ";" & e.Row.Cells("CONTACT_EMAIL").Value
                End If
            Else
                If e.Row.Cells("CONTACT_EMAIL").Value & "" <> "" Then
                    ' WIERD - SEND_TOs is acting like an object with .StartsWith and .EndsWith
                    ' if SEND_TOs = "", you get thrown an Object reference not set error
                    ' need to protect with If SEND_TOs <> ""
                    If SEND_TOs <> "" Then
                        SEND_TOs = Replace(SEND_TOs, e.Row.Cells("CONTACT_EMAIL").Value & "", "", 1, 1)
                        SEND_TOs = Replace(SEND_TOs, ";;", ";")
                        If SEND_TOs <> "" AndAlso SEND_TOs.StartsWith(";") Then
                            SEND_TOs = Mid(SEND_TOs, 2)
                        End If
                        If SEND_TOs <> "" AndAlso SEND_TOs.EndsWith(";") Then
                            SEND_TOs = Mid(SEND_TOs, 1, Len(SEND_TOs) - 1)
                        End If
                    End If
                End If
            End If
            txtSEND_TOs.Text = SEND_TOs
        End If
    End Sub

    Private Sub grdTATCONT1_BeforeRowsDeleted(sender As Object, e As UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdTATCONT1.BeforeRowsDeleted
        If MsgBox("Delete Contacts Selected" & "?", MsgBoxStyle.YesNo, "Confirm Deletion") = MsgBoxResult.No Then
            e.Cancel = True
        End If
        e.DisplayPromptMsg = False
    End Sub

    Private Sub grdTATCONT1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdTATCONT1.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("CONTACT_NO").Value = ASCMAIN1.Next_Control_No("TATCONT1.CONTACT_NO")
            e.Row.Cells("CONTACT_ENTITY_TABLE").Value = CONTACT_ENTITY_TABLE
            e.Row.Cells("CONTACT_ENTITY_KEY").Value = CONTACT_ENTITY_KEY
            e.Row.Cells("CONTACT_STATUS").Value = "A"
            e.Row.Cells("SEL").Value = "1"
        Else
            Dim rowTATCONT1 As DataRow = dst.Tables("TATCONT1").Rows.Find _
                (New Object() {e.Row.Cells("CONTACT_NO").Value})
            RowState_b4_checked = rowTATCONT1.RowState
        End If
    End Sub

    Private Sub cmdImportFromOutlook_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImportFromOutlook.Click

        Me.Cursor = Cursors.WaitCursor

        Try

            Dim outlook As Microsoft.Office.Interop.Outlook.Application = CType(Microsoft.VisualBasic.Interaction.GetObject("", "Outlook.Application"), Microsoft.Office.Interop.Outlook.Application)
            'Dim explorer As Microsoft.Office.Interop.Outlook.Explorer = outlook.ActiveExplorer
            Dim mail As Microsoft.Office.Interop.Outlook.MailItem ' = CType(explorer.Selection.Item(i + 1), Microsoft.Office.Interop.Outlook.MailItem)

            mail = outlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem)
            Dim olexpl As Microsoft.Office.Interop.Outlook.Explorer = outlook.Explorers(1)

            Dim tblTATCONT1 As DataTable = dst.Tables("TATCONT1").Clone

            Dim contacts As String = ""
            For Each fldr As Microsoft.Office.Interop.Outlook.Folder In outlook.GetNamespace("mapi").Folders
                If fldr.Name = "Personal Folders" Then
                    For Each fldr2 As Microsoft.Office.Interop.Outlook.Folder In fldr.Folders
                        If fldr2.Name = "Contacts" Then
                            Dim CONTACT_NO_ctr As Integer = 0
                            For Each x As Object In fldr2.Items
                                If TypeOf (x) Is Microsoft.Office.Interop.Outlook.ContactItem Then
                                    Dim xc As Microsoft.Office.Interop.Outlook.ContactItem = DirectCast(x, Microsoft.Office.Interop.Outlook.ContactItem)
                                    If x.Email1Address <> "" And x.Email1DisplayName <> "" Then
                                        Dim row As DataRow = tblTATCONT1.NewRow
                                        row.Item("CONTACT_EMAIL") = Mid(xc.Email1Address, 1, 60)
                                        row.Item("CONTACT_NAME") = Replace(Replace(Replace(xc.FirstName & " " & xc.MiddleName & " " & xc.LastName & " " & xc.Suffix, "  ", " "), "  ", " "), "  ", " ")
                                        row.Item("CONTACT_TITLE") = xc.JobTitle
                                        row.Item("CONTACT_PHONE") = Mid(Replace(Replace(Replace(Replace(xc.BusinessTelephoneNumber, "(", ""), ")", ""), "-", ""), " ", ""), 1, 15)
                                        row.Item("CONTACT_FAX") = Mid(Replace(Replace(Replace(Replace(xc.BusinessFaxNumber, "(", ""), ")", ""), "-", ""), " ", ""), 1, 10)
                                        row.Item("CONTACT_ENTITY_TABLE") = CONTACT_ENTITY_TABLE
                                        row.Item("CONTACT_ENTITY_KEY") = CONTACT_ENTITY_KEY
                                        CONTACT_NO_ctr += 1
                                        row.Item("CONTACT_NO") = Format(CONTACT_NO_ctr, "0000000000")
                                        tblTATCONT1.Rows.Add(row)
                                    End If
                                    'contacts &= x.Email1Address & ":" & x.Email1DisplayName
                                End If
                            Next
                        End If
                    Next
                End If
            Next

            Me.Cursor = Cursors.Default

            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CONTACT_NO")
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = CODE_VALUE
                tblTATCONT1.DefaultView.Sort = "CONTACT_NAME"
                ASCMAIN1.CodeSelector.UseDataFromTable = tblTATCONT1
                Dim F As New ASFCODE1
                F.ShowDialog()
                F.Dispose()
                If ASCMAIN1.CodeSelector.Selections <> 0 Then
                    For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                        Dim rowTATCONT1 As DataRow = dst.Tables("TATCONT1").NewRow
                        For Each DC As DataColumn In row.Table.Columns
                            If rowTATCONT1.Table.Columns.Contains(DC.ColumnName) Then
                                rowTATCONT1.Item(DC.ColumnName) = row.Item(DC.ColumnName)
                            End If
                        Next
                        rowTATCONT1.Item("CONTACT_NO") = ASCMAIN1.Next_Control_No("TATCONT1.CONTACT_NO")
                        rowTATCONT1.Item("CONTACT_STATUS") = "A"
                        rowTATCONT1.Item("CONTACT_ENTITY_TABLE") = CONTACT_ENTITY_TABLE
                        rowTATCONT1.Item("CONTACT_ENTITY_KEY") = CONTACT_ENTITY_KEY
                        rowTATCONT1.Item("SEL") = "1"
                        dst.Tables("TATCONT1").Rows.Add(rowTATCONT1)
                        If allow_editing Then
                            Setup_to_Save_Changes()
                        End If
                    Next
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Import Contacts from Outlook")

        End Try

        Me.Cursor = Cursors.Default

    End Sub

    Sub Setup_to_Save_Changes()
        If Not chkSaveChanges.Visible And Not chkSaveChanges.Checked Then chkSaveChanges.Checked = True
        chkSaveChanges.Visible = True
    End Sub
End Class