Imports Infragistics.Win.UltraWinGrid

Public Class SPTENOT1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SPTENOT2.*" & vbCrLf _
                & ", GLTPARM3.LEGEND, GLTPARM3.WEEK_END_DATE" & vbCrLf _
                & ", GLTPARM3.YYYYMM, GLTPARM3.REL_WEEK, GLTPARM3.MAX_WEEK" & vbCrLf _
                & " from SPTENOT2,GLTPARM3" & vbCrLf _
                & " where SPTENOT2.SEASON_CODE = :PARM1" & vbCrLf _
                & "   and GLTPARM3.YYYYWW = SPTENOT2.OPS_YYYYWW"
            Create_TDA(.Tables.Add, "SPTENOT2", "**", 0, True, "V", 2)
            .Tables("SPTENOT2").Columns.Add("LEGEND_ABBR", GetType(System.String), "SUBSTRING(LEGEND,10,7)")
        End With

        grdSPTENOT2.DataSource = dst.Tables("SPTENOT2")

        With grdSPTENOT2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightBlue
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If gcol.Key = "WEEKLY_NOTE_SAMP" Or gcol.Key = "WEEKLY_NOTE_ADDL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellMultiLine = DefaultableBoolean.True
                    .Override.RowSizing = RowSizing.AutoFree
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                End If
            Next
        End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSPTENOT2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

            Case "grdSPTENOT2"
                'tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                'tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")


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
            Case "Add Codes"

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "ABSColumn Controls"
    Public Overrides Sub num_ValueChanged(sender As Object, e As EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "DEMO_COMM_PCT"
                'If Not Me.IsDone Then
                '    Rebuild_StoreByHC()
                'End If
        End Select
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"

                Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
                If LookUp("ICTSEAS1", SEASON_CODE) Is Nothing Then
                    EMsg &= vbCrLf & "Invalid Value for Season"
                End If

            Case "Edit"

            Case "Update"

                'If optDEMO_COMM_BASIS.Value & "" = "" Then
                '    EMsg &= vbCr & "Please Choose a Commission Basis"
                'End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

        WriteAuditTrail("SPTENOT2")

        Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
        Dim sqlDelete = "SEASON_CODE = '" & SEASON_CODE & "'"
        Update_Record_TDA("SPTENOT2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        If EntryMode = "New" Then
            Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
            Dim rowICTSEAS1 As DataRow = LookUp("ICTSEAS1", SEASON_CODE)
            Dim SEASON_YEAR As String = rowICTSEAS1.Item("SEASON_YEAR")
            Dim SEASON_TYPE As String = rowICTSEAS1.Item("SEASON_TYPE")
            Dim M1 As String = "01"
            Dim M2 As String = "06"
            If SEASON_TYPE = "F" Then
                M1 = "07"
                M2 = "12"
            End If
            ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM Between :PARM1 and :PARM2"
            Dim TBL As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", New String() {SEASON_YEAR & M1, SEASON_YEAR & M2})
            For Each row As DataRow In TBL.Rows
                dst.Tables("SPTENOT2").Rows.Add(New Object() {SEASON_CODE, row.Item("YYYYWW"), "", "", row.Item("LEGEND"), row.Item("WEEK_END_DATE"),
                                                row.Item("YYYYMM"), row.Item("REL_WEEK"), row.Item("MAX_WEEK")})
            Next
        Else
            Fill_Records("SPTENOT2", New String() {Absx1.txtFor("SEASON_CODE").Text})
        End If
        Sort_grdColumns(grdSPTENOT2, "OPS_YYYYWW")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SPTENOT2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSPTENOT2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTENOT2}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.False
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

    Private Sub grdSPTENOT2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSPTENOT2.InitializeRow
        Dim REL_WEEK As Integer = Val(e.Row.Cells("REL_WEEK").Value & "")
        Dim MAX_WEEK As Integer = Val(e.Row.Cells("MAX_WEEK").Value & "")
        If REL_WEEK = MAX_WEEK Then
            e.Row.RowSpacingAfter = 5
        End If
    End Sub

#End Region
End Class