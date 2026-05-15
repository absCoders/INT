Imports Infragistics.Win.UltraWinToolbars

Public Class GLTCLAY1

    Dim sqlGLTDSTR2 As String = ""
    Dim VV As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select * from GLTCALC1"
            Create_TDA(.Tables.Add, "GLTCALC1", "**", 0, False)
            With .Tables("GLTCALC1")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                .Columns.Add("C", GetType(System.Int32))
                .Columns.Add("P", GetType(System.Int32))
            End With

            Create_TDA(.Tables.Add, "GLTCLAY2", "*", 1)

            With .Tables.Add("GLTCLAYX")
                .Columns.Add("KEY")
                For I As Integer = 1 To 99
                    .Columns.Add("C" & Format(I, "00"), GetType(System.Int32))
                Next
            End With
        End With

        Fill_Records("GLTCALC1")
        Sort_grdColumns(grdGLTCALC1, "STMT_CALC_CODE")

        dst.Tables("GLTCLAYX").Rows.Add("Min Prd ->")

        grdGLTCALC1.DataSource = dst.Tables("GLTCALC1")
        With grdGLTCALC1.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With
        With grdGLTCALC1.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("STMT_CALC_CODE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightBlue
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                End If
                If New String() {"STMT_CALC_YEAR", "STMT_CALC_NO", "STMT_CALC_TYPE", "STMT_CALC_DATA_TYPE", "STMT_CALC_PERIOD"}.Contains(gcol.Key) Then
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                End If
            Next
        End With

        grdGLTCLAYX.DataSource = dst.Tables("GLTCLAYX")
        With grdGLTCLAYX.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.True
            .AllowDelete = DefaultableBoolean.False
        End With
        With grdGLTCLAYX.DisplayLayout.Bands(0)
            With .Columns("KEY")
                .Header.Fixed = True
                .Header.Caption = ""
                .Width = 130
            End With
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightGreen
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If gcol.Key = "KEY" Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        End With

        ASCMAIN1.Add_Value_List(grdGLTCALC1, "STMT_CALC_TYPE")
        ASCMAIN1.Add_Value_List(grdGLTCALC1, "STMT_CALC_NO")
        ASCMAIN1.Add_Value_List(grdGLTCALC1, "STMT_CALC_YEAR")
        ASCMAIN1.Add_Value_List(grdGLTCALC1, "STMT_CALC_DATA_TYPE")
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTCLAYX, "SBC", "B1", "B2", "Back-Color")

        If ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS Then

            ' Create a toolbars and add it to the UltraToolbarManager's toolbars collection.
            Me.UltraToolbarsManager1.Toolbars.AddToolbar("FormattingOptions")

            ' Create a PopupColorPickerTool
            Dim popupColorPickerTool As New PopupColorPickerTool("TextForeColor")

            ' Always add new tools to the UltraToolbarManager's root tools collection
            ' before adding them to menus or toolbars.
            Me.UltraToolbarsManager1.Tools.AddRange(New ToolBase() {popupColorPickerTool})

            ' Add the tools to the toolbar.
            Me.UltraToolbarsManager1.Toolbars("FormattingOptions").Tools.AddTool("TextForeColor")

            ' Set some properties on the PopupColorPickerTool.
            popupColorPickerTool.SelectedColor = System.Drawing.Color.Blue
            Me.UltraToolbarsManager1.Toolbars("FormattingOptions").Tools("TextForeColor").InstanceProps.IsFirstInGroup = True
        End If
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

            'Case "grdGLTDSTR2"
            '    tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
            '    tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.Tool.Key

                'Case "grdGLTCLAYX"
                '    If grdGLTCLAYX.ActiveCell IsNot Nothing Then
                '        Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                '        tlb_cpt.ShowPopup()

                '    End If

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Select Case e.Tool.Key
            'Case "Add Codes"
            '    If grd.Name = "grdGLTDSTR2" Then
            '        Add_Codes(grdGLTDSTR2, "GLTACCT1", "ACCT_CODE", "Accounts")
            '    End If

            Case "Back-Color"
                If grdGLTCLAYX.ActiveCell IsNot Nothing Then
                    Dim gcol As UltraWinGrid.UltraGridColumn = grdGLTCLAYX.ActiveCell.Column
                    Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                    tlb_cpt.ReplaceableColor = gcol.CellAppearance.BackColor

                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub


    Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Back-Color"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                'grd.ActiveRow.Cells("POS_RBG_BACKCOLOR").Value = tlb_cpt.SelectedColor.ToArgb
                'grd.UpdateData()
                'Application.DoEvents()
                'grdGLTCLAYX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
                'Update_Record_TDA("IMTPOSS1")

                'Case "ForeColor"
                '    Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                '    = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                '    grd.ActiveRow.Cells("POS_RBG_FORECOLOR").Value = tlb_cpt.SelectedColor.ToArgb
                '    grd.UpdateData()
                '    'Application.DoEvents()
                '    grdIMTSTATW.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
                '    Update_Record_TDA("IMTPOSS1")

                'Case "Best"
                '    Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                '    = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                '    Me.UltraChart1.ColorModel.ColorEnd = tlb_cpt.SelectedColor
                '    UltraChart1.DataBind()
                '    'grdSATCSLSS.DataBind()
                '    Application.DoEvents()
                '    grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

                'Case "Worst"
                '    Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                '    = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                '    Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
                '    UltraChart1.DataBind()
                '    'grdSATCSLSS.DataBind()
                '    Application.DoEvents()
                '    grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select

    End Sub


#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"
                'Dim PMAX As Integer = Val(dst.Tables("GLTCALC1").Compute("MAX(P)", "") & "")
                'If PMAX > 13 Then
                '    EMsg &= vbCr & "Only 13 columns supported at present"
                'End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim STMT_LAYOUT_CODE As String = Absx1.txtFor("STMT_LAYOUT_CODE").Text

        Dim rowGLTCLAYX As DataRow = dst.Tables("GLTCLAYX").Rows(0)

        dst.Tables("GLTCLAY2").Rows.Clear()
        For Each row As DataRow In dst.Tables("GLTCALC1").Select("SEL='1' and P <> 0")
            Dim rowGLTCLAY2 As DataRow = dst.Tables("GLTCLAY2").NewRow
            rowGLTCLAY2.Item("STMT_LAYOUT_CODE") = STMT_LAYOUT_CODE
            rowGLTCLAY2.Item("STMT_CALC_CODE") = row.Item("STMT_CALC_CODE")
            Dim STMT_COL_POS As Integer = Val(row.Item("P") & "")
            'If STMT_COL_POS <= 13 Then
            '    rowASFBASE1.Item("STMT_CALC_CODE_" & Format(STMT_COL_POS, "00")) = row.Item("STMT_CALC_CODE")
            '    rowASFBASE1.Item("STMT_MIN_PRD_" & Format(STMT_COL_POS, "00")) = rowGLTCLAYX.Item("C" & Format(STMT_COL_POS, "00"))
            'End If
            rowGLTCLAY2.Item("STMT_COL_POS") = STMT_COL_POS
            rowGLTCLAY2.Item("STMT_MIN_PRD") = rowGLTCLAYX.Item("C" & Format(STMT_COL_POS, "00"))
            dst.Tables("GLTCLAY2").Rows.Add(rowGLTCLAY2)
        Next

        Dim sqlDelete = "STMT_LAYOUT_CODE = '" & STMT_LAYOUT_CODE & "'"
        Update_Record_TDA("GLTCLAY2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("GLTCLAY2", New String() {Absx1.txtFor("STMT_LAYOUT_CODE").Text})
        'Sort_grdColumns(grdGLTCLAY2, "STMT_LAYOUT_CODE")
        Sort_grdColumns(grdGLTCALC1, "STMT_CALC_CODE")

        For Each row As DataRow In dst.Tables("GLTCALC1").Select("")
            row.Item("SEL") = "0"
            row.Item("C") = 0
            row.Item("P") = 0
        Next
        Dim P As Integer = 0
        Dim rowGLTCLAYX As DataRow = dst.Tables("GLTCLAYX").Rows(0)
        For Each row As DataRow In dst.Tables("GLTCLAY2").Select("", "STMT_COL_POS")
            Dim STMT_CALC_CODE As String = row.Item("STMT_CALC_CODE")
            'Dim STMT_COL_POS As Integer = Val(row.Item("STMT_COL_POS") & "")
            Dim STMT_MIN_PRD As String = Val(row.Item("STMT_MIN_PRD") & "")
            Dim rowGLTCALC1 As DataRow = dst.Tables("GLTCALC1").Rows.Find(STMT_CALC_CODE)
            rowGLTCALC1.Item("SEL") = "1"
            P += 1
            rowGLTCALC1.Item("C") = P
            rowGLTCALC1.Item("P") = P
            rowGLTCLAYX.Item("C" & Format(P, "00")) = STMT_MIN_PRD
        Next

        EnforceConstraints(True)

        Set_Layout()

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"GLTCLAY2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)

            For Each row As DataRow In dst.Tables("GLTCALC1").Select("")
                row.Item("SEL") = "0"
            Next
            With grdGLTCLAYX.DisplayLayout.Bands(0)
                For I As Integer = 1 To 99
                    If Not .Columns("C" & Format(I, "00")).Hidden Then
                        .Columns("C" & Format(I, "00")).Hidden = True
                    End If
                Next
            End With

        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        'grdGLTDSTR2.Enabled = tf
        grdGLTCALC1.Visible = tf
        grdGLTCLAYX.Visible = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdGLTCALC1, grdGLTCLAYX}
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

#End Region

    Private Sub grdGLTCALC1_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdGLTCALC1.AfterRowUpdate
        Set_Layout()
    End Sub

    Private Sub grdGLTCALC1_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdGLTCALC1.BeforeRowUpdate

        Dim C As Integer = Val(dst.Tables("GLTCALC1").Compute("MAX(C)", "") & "") + 1
        If e.Row.Cells("SEL").Text & "" = "1" Or e.Row.Cells("SEL").Text & "" = "Checked" Then
            e.Row.Cells("C").Value = C
            e.Row.Cells("P").Value = C
        Else
            e.Row.Cells("C").Value = 0
            e.Row.Cells("P").Value = 0
        End If

    End Sub

    Private Sub grdGLTCALC1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTCALC1.DoubleClickRow
        If e.Row.IsDataRow AndAlso (EntryMode = "N" Or EntryMode = "E") Then

        End If
    End Sub

    Sub Set_Layout()

        Dim V As String = "".PadLeft(100, "0")

        With grdGLTCLAYX.DisplayLayout.Bands(0)
            For Each row As DataRow In dst.Tables("GLTCALC1").Select("SEL='1'")
                Dim C As Integer = Val(row.Item("C") & "")
                If C = 0 Then
                    ASCMAIN1.Progress("Problem with Column " & row.Item("STMT_CALC_CODE") & " - Selected but not in grid")
                Else
                    Dim P As Integer = Val(row.Item("P") & "")
                    With .Columns("C" & Format(C, "00"))
                        .Hidden = False
                        .Header.Caption = row.Item("STMT_CALC_CODE") & vbCrLf & Replace(row.Item("STMT_CALC_DESC") & "", ",", vbCrLf)
                        .Width = 80
                        .Header.VisiblePosition = P
                        Mid(V, C, 1) = "1"
                        If VV.Length >= C AndAlso Mid(VV, C, 1) = "1" Then
                            Mid(VV, C, 1) = "0"
                        End If
                    End With
                End If
            Next

            Do While InStr(VV, "1")
                Dim C As Integer = InStr(VV, "1")
                Mid(VV, C, 1) = "0"
                With .Columns("C" & Format(C, "00"))
                    .Hidden = True
                End With
            Loop

            VV = V
        End With
    End Sub

    Private Sub grdGLTCLAYX_AfterColPosChanged(sender As Object, e As UltraWinGrid.AfterColPosChangedEventArgs) Handles grdGLTCLAYX.AfterColPosChanged
        For Each row As DataRow In dst.Tables("GLTCALC1").Select("SEL='1'")
            Dim C As Integer = Val(row.Item("C") & "")
            If C <> 0 Then
                row.Item("P") = grdGLTCLAYX.DisplayLayout.Bands(0).Columns("C" & Format(C, "00")).Header.VisiblePosition
            End If
        Next
    End Sub
End Class