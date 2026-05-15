Public Class ASFEVNT1

    Public result As Boolean = False
    Private _E As Events_Entity

    Public Sub New(ByVal FF As ASFBASE1, ByVal E As Events_Entity)
        frmASFBASE1 = FF
        _E = E
        InitializeComponent()
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from TATEVNT1 " _
            & " where TABLE_NAME = :PARM1 and TABLE_KEY = :PARM2"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "VV", 0)
        End With

        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")

        With grdTATEVNT1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False
            .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        End With

        Fill_Records("TATEVNT1", New String() {_E.TABLE_NAME, _E.TABLE_KEY})
        grdTATEVNT1.Text = "Events for " & _E.TABLE_KEY_CAPTION & " " & _E.TABLE_KEY
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        ' Load_Popup_Menu(grdTATCONT1, "SSSS", "Show Filter", "Show GroupBox", "Show Pins")
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
        result = True
        Exit_Form()
    End Sub

    Sub Exit_Form()
        frmASFBASE1 = Nothing
        Me.Close()
    End Sub

End Class