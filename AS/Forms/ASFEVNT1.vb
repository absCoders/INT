Public Class ASFEVNT1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from TATEVNT1 WHERE TABLE_NAME IN ('ASTUSER1','ASTUSER2','ASTMENU1','ASTLOGS1')"
            Create_TDA(.Tables.Add, "TATEVNTX", "**", 0, False)
        End With

        grdTATEVNTX.DataSource = dst.Tables("TATEVNTX")
        Create_Summary(grdTATEVNTX, "TABLE_NAME", "Count")

        dte0.Value = CDate("01/01/" & Now.Year)
        dte1.Value = CDate("01/01/" & Now.Year).AddDays(-1).AddYears(1)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Load_Record()
                Mode_Settings(True)
            Case "Audit Report"
                Generate_Audit_Report()
            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Audit Report").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdTATEVNTX.Visible = ScreenMode

        Set_Read_Only(grpDateRange, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Errors from Log")
        Me.Cursor = Cursors.WaitCursor

        Fill_TATEVNTX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()
        dst.Tables("TATEVNTX").Rows.Clear()
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdTATEVNTX, "SS", "Show Filter", "Show GroupBox", "Refresh")
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

            '    Case "grdATEVNTX"
            '        tlb_btn = DirectCast(tlb_pop.Tools("Refresh"), UltraWinToolbars.ButtonTool)
            '        tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

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

        Select Case e.Tool.Key
            Case "Refresh"
                If grd.Name = "grdTATEVNTX" Then
                    Fill_TATEVNTX()
                End If
        End Select
        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub

#End Region

    Sub Fill_TATEVNTX()
        Fill_Records("TATEVNTX")
        Sort_grdColumns(grdTATEVNTX, "INIT_DATE".ToLower)
    End Sub

    Sub Generate_Audit_Report()
        ASCMAIN1.Progress("Building Audit Workbook")

        Dim FILENAME As String = TAC.TACMAIN1.Generate_Audit_Workbook(Me, dte0.Value, dte1.Value)
        Show_Document(FILENAME)
        ASCMAIN1.Progress("")

    End Sub
End Class