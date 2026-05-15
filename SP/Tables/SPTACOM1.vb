Public Class SPTACOM1

    Dim HC_CODEs() As String
    Dim HC_CODE_2_B As New Dictionary(Of String, Integer)
    Dim ASP_COMM_PCTs() As Decimal
    Dim cellAppearance As New Infragistics.Win.Appearance
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SPTACOM2.*" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_LOCATION, ARTCUST2.MALL_CODE" & vbCrLf _
                & " from SPTACOM2,ARTCUST2" & vbCrLf _
                & " where SPTACOM2.ASP_CODE = :PARM1" & vbCrLf _
                & "   and SPTACOM2.CUST_CODE = :PARM2" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SPTACOM2.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SPTACOM2.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SPTACOM2", "**", 0, True, "VV", 3)

            ASCMAIN1.sql = "Select SPTACOM3.*" & vbCrLf _
                & ", ICTCOLL0.HC_NAME" & vbCrLf _
                & " from SPTACOM3,ICTCOLL0" & vbCrLf _
                & " where SPTACOM3.ASP_CODE = :PARM1" & vbCrLf _
                & "   and SPTACOM3.CUST_CODE = :PARM2" & vbCrLf _
                & "   and ICTCOLL0.HC_CODE = SPTACOM3.HC_CODE"
            Create_TDA(.Tables.Add, "SPTACOM3", "**", 0, True, "VV", 3)

            Create_TDA(.Tables.Add, "SPTACOM4", "*", 2)

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, False)
            Create_TDA(.Tables.Add, "ICTCOLL0", "*", 0, False)

            With .Tables.Add("SPTACOMX")
                .Columns.Add("CUST_STORE_NO")
                For I As Integer = 0 To 99
                    .Columns.Add("ASP_COMM_PCT" & Format(I, "00"), GetType(System.Decimal))
                Next
                .PrimaryKey = New DataColumn() {.Columns("CUST_STORE_NO")}
            End With
        End With

        Fill_Records("ICTCOLL0")

        grdSPTACOM2.DataSource = dst.Tables("SPTACOM2")
        grdSPTACOM3.DataSource = dst.Tables("SPTACOM3")
        grdSPTACOMX.DataSource = dst.Tables("SPTACOMX")

        With grdSPTACOM2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightBlue
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If gcol.Key = "ASP_COMM_PCT" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdSPTACOM3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightGreen
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If gcol.Key = "ASP_COMM_PCT" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdSPTACOMX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = Drawing.Color.Orange
                    If gcol.Key = "CUST_STORE_NO" Or gcol.Key = "ASP_COMM_PCT00" Then
                        gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    End If
                End With
                If gcol.Key.StartsWith("ASP_COMM_PCT") And gcol.Key <> "ASP_COMM_PCT00" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            .Columns("CUST_STORE_NO").Header.Fixed = True
            ReDim HC_CODEs(dst.Tables("ICTCOLL0").Rows.Count)
            HC_CODE_2_B.Clear()

            Dim B As Integer = 0
            With .Columns("ASP_COMM_PCT" & Format(B, "00"))
                .Header.Caption = "Comm%"
                .Hidden = False
                .Width = 70
                .Format = "#.00"
                .Header.Fixed = True
            End With

            For Each row As DataRow In dst.Tables("ICTCOLL0").Select("", "HC_CODE")
                B += 1
                HC_CODEs(B) = row.Item("HC_CODE")
                HC_CODE_2_B.Add(row.Item("HC_CODE"), B)
                With .Columns("ASP_COMM_PCT" & Format(B, "00"))
                    .Header.Caption = row.Item("HC_CODE")
                    '   .Hidden = False
                    .Width = 70
                    .Format = "#.00"
                End With
            Next
        End With

        Dim uddCUST_STORE_NO As New UltraWinGrid.UltraDropDown
        Dim DVWCUST_STORE_NO As DataView = New DataView(dst.Tables("ARTCUST2"), "ISNULL(CUST_DC_IND,'0') <> '1'", "CUST_STORE_NO", DataViewRowState.CurrentRows)
        uddCUST_STORE_NO.DataSource = DVWCUST_STORE_NO
        For Each GC As UltraWinGrid.UltraGridColumn In uddCUST_STORE_NO.DisplayLayout.Bands(0).Columns
            If GC.Key <> "CUST_STORE_NO" And GC.Key <> "CUST_STORE_NAME" And GC.Key <> "CUST_STORE_LOCATION" Then
                GC.Hidden = True
            Else
                If GC.Key = "CUST_STORE_NO" Then GC.Header.Caption = "Store No"
                If GC.Key = "CUST_STORE_NAME" Then GC.Header.Caption = "Store Name"
                If GC.Key = "CUST_STORE_LOCATION" Then GC.Header.Caption = "Location"
            End If
        Next
        uddCUST_STORE_NO.ValueMember = "CUST_STORE_NO"
        With grdSPTACOM2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO")
            .ValueList = uddCUST_STORE_NO
            '.ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
        End With

        Dim uddHC_CODE As New UltraWinGrid.UltraDropDown
        Dim DVWHC_CODE As DataView = New DataView(dst.Tables("ICTCOLL0"), "", "HC_CODE", DataViewRowState.CurrentRows)
        uddHC_CODE.DataSource = DVWHC_CODE
        For Each GC As UltraWinGrid.UltraGridColumn In uddHC_CODE.DisplayLayout.Bands(0).Columns
            If GC.Key <> "HC_CODE" And GC.Key <> "HC_NAME" Then
                GC.Hidden = True
            Else
                If GC.Key = "HC_CODE" Then GC.Header.Caption = "HC Code"
                If GC.Key = "HC_NAME" Then GC.Header.Caption = "HC Name"
            End If
        Next
        uddHC_CODE.ValueMember = "HC_CODE"
        With grdSPTACOM3.DisplayLayout.Bands(0).Columns("HC_CODE")
            .ValueList = uddHC_CODE
            '.ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
        End With

        cellAppearance.ForeColor = Drawing.Color.Red
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSPTACOM2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

            Case "grdSPTACOM2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdARTCUST1"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

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
            Case "ASP_COMM_PCT"
                If Not Me.IsDone Then
                    Rebuild_StoreByHC()
                End If
        End Select
    End Sub

#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"

            Case "Update"

                If optASP_COMM_BASIS.Value & "" = "" Then
                    EMsg &= vbCr & "Please Choose a Commission Basis"
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

        WriteAuditTrail("SPTACOM2")
        WriteAuditTrail("SPTACOM3")
        WriteAuditTrail("SPTACOM4")

        Dim ASP_CODE As String = Absx1.txtFor("ASP_CODE").Text
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim sqlDelete = "ASP_CODE = '" & ASP_CODE & "' and CUST_CODE = '" & CUST_CODE & "'"
        Update_Record_TDA("SPTACOM2", sqlDelete)
        Update_Record_TDA("SPTACOM3", sqlDelete)
        Update_Record_TDA("SPTACOM4", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)

        If EntryMode = "New" Then
            rowASFBASE1.Item("ASP_COMM_STATUS") = "A"
        End If

        Fill_Records("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text})

        Fill_Records("SPTACOM2", New String() {Absx1.txtFor("ASP_CODE").Text, Absx1.txtFor("CUST_CODE").Text})
        Sort_grdColumns(grdSPTACOM2, "CUST_STORE_NO")
        Fill_Records("SPTACOM3", New String() {Absx1.txtFor("ASP_CODE").Text, Absx1.txtFor("CUST_CODE").Text})
        Sort_grdColumns(grdSPTACOM3, "HC_CODE")

        Fill_Records("SPTACOM4", New String() {Absx1.txtFor("ASP_CODE").Text, Absx1.txtFor("CUST_CODE").Text})

        Toggle_HC_Column_Display()
        Rebuild_StoreByHC()
        Sort_grdColumns(grdSPTACOMX, "CUST_STORE_NO")

        grdSPTACOMX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        Setup_BASIS()

        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SPTACOM2", "SPTACOM3", "SPTACOM4", "SPTACOMX", "ARTCUST2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSPTACOM2.Enabled = tf
        grdSPTACOM3.Enabled = tf
        grdSPTACOMX.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTACOM2, grdSPTACOM3, grdSPTACOMX}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    If grd.Name = "grdSPTACOMX" Then
                        .AllowUpdate = DefaultableBoolean.True
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowUpdate = DefaultableBoolean.True
                        .AllowDelete = DefaultableBoolean.True
                    End If

                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next
    End Sub

#End Region

#Region "grdSPTACOM2"

    Private Sub grdSPTACOM2_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTACOM2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_STORE_NO"
                Dim CUST_STORE_NO As String = e.Cell.Value & ""
                Dim row As DataRow = dst.Tables("ARTCUST2").Rows.Find(New String() {Absx1.txtFor("CUST_CODE").Text, CUST_STORE_NO})
                If row IsNot Nothing Then
                    e.Cell.Row.Cells("CUST_STORE_NAME").Value = row.Item("CUST_STORE_NAME")
                    e.Cell.Row.Cells("CUST_STORE_LOCATION").Value = row.Item("CUST_STORE_LOCATION")
                    e.Cell.Row.Cells("MALL_CODE").Value = row.Item("MALL_CODE")
                End If
        End Select
    End Sub

    Private Sub grdSPTACOM2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTACOM2.AfterRowActivate

    End Sub

    Private Sub grdSPTACOM2_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSPTACOM2.AfterRowsDeleted
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTACOM2_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTACOM2.AfterRowUpdate
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTACOM2_BeforeCellUpdate(sender As Object, e As UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSPTACOM2.BeforeCellUpdate

    End Sub

    Private Sub grdSPTACOM2_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTACOM2.BeforeExitEditMode
        'If grdSPTACOM2.ActiveRow.IsDataRow And Not grdSPTACOM2.ActiveRow.IsAddRow Then

        'End If
        If grdSPTACOM2.ActiveCell IsNot Nothing Then
            With grdSPTACOM2.ActiveCell
                Select Case .Column.Key
                    Case "CUST_STORE_NO"
                        If .EditorResolved.IsValid AndAlso .EditorResolved.Value & "" <> "" Then
                            Dim CUST_STORE_NO As String = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                            If .EditorResolved.Value <> CUST_STORE_NO Then
                                .EditorResolved.Value = CUST_STORE_NO
                            End If
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSPTACOM2_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTACOM2.BeforeRowActivate
        With grdSPTACOM2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO")
            If e.Row.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSPTACOM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTACOM2.BeforeRowUpdate

        If Val(e.Row.Cells("ASP_COMM_PCT").Value & "") < 0 Then
            e.Cancel = True
        End If

        Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Value & ""
        If LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, CUST_STORE_NO}) Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
        End If

    End Sub
#End Region

#Region "grdSPTACOM3"

    Private Sub grdSPTACOM3_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTACOM3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "HC_CODE"
                Dim HC_CODE As String = e.Cell.Value & ""
                Dim row As DataRow = dst.Tables("ICTCOLL0").Rows.Find(HC_CODE)
                If row IsNot Nothing Then
                    e.Cell.Row.Cells("HC_NAME").Value = row.Item("HC_NAME")
                End If

        End Select
    End Sub

    Private Sub grdSPTACOM3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTACOM3.AfterRowActivate

    End Sub

    Private Sub grdSPTACOM3_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSPTACOM3.AfterRowsDeleted
        Toggle_HC_Column_Display()
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTACOM3_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTACOM3.AfterRowUpdate
        Toggle_HC_Column_Display()
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTACOM3_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTACOM3.BeforeRowActivate
        With grdSPTACOM3.DisplayLayout.Bands(0).Columns("HC_CODE")
            If e.Row.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSPTACOM3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTACOM3.BeforeRowUpdate
        If Val(e.Row.Cells("ASP_COMM_PCT").Value & "") < 0 Then
            e.Cancel = True
        End If

        Dim HC_CODE As String = e.Row.Cells("HC_CODE").Value
        If LookUp("ICTCOLL0", New String() {HC_CODE}) Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
        End If
    End Sub
#End Region

    Sub Toggle_HC_Column_Display()

        Dim HC_CODEs_to_display As New List(Of String)
        For Each row As DataRow In dst.Tables("SPTACOM3").Select("")
            Dim HC_CODE As String = row.Item("HC_CODE")
            HC_CODEs_to_display.Add(HC_CODE)
        Next

        For b As Integer = 1 To HC_CODEs.Length - 1
            With grdSPTACOMX.DisplayLayout.Bands(0).Columns("ASP_COMM_PCT" & Format(b, "00"))
                Dim HC_CODE As String = HC_CODEs(b)
                .Hidden = Not HC_CODEs_to_display.Contains(HC_CODE)
            End With
        Next

        Dim ASP_COMM_PCT As Decimal = Val(Absx1.numFor("ASP_COMM_PCT").Value & "")
        Get_Pct_by_HC(ASP_COMM_PCT)
    End Sub
    Sub Rebuild_StoreByHC()
        dst.Tables("SPTACOMX").Rows.Clear()

        Dim ASP_COMM_PCT As Decimal = Val(Absx1.numFor("ASP_COMM_PCT").Value & "")
        Dim ASP_CODE As String = Absx1.txtFor("ASP_CODE").Text

        Dim i As Integer = 0
        For Each row As DataRow In dst.Tables("ARTCUST2").Select("")
            i += 1
            Dim rowSPTACOMX As DataRow = dst.Tables("SPTACOMX").NewRow
            rowSPTACOMX.Item("CUST_STORE_NO") = row.Item("CUST_STORE_NO")
            Dim ASP_COMM_PCT00 As Decimal = -1
            Dim rowSPTACOM2 As DataRow = dst.Tables("SPTACOM2").Rows.Find(New String() {ASP_CODE, row.Item("CUST_CODE"), row.Item("CUST_STORE_NO")})
            If rowSPTACOM2 IsNot Nothing Then
                ASP_COMM_PCT00 = rowSPTACOM2.Item("ASP_COMM_PCT")
            End If
            rowSPTACOMX.Item("ASP_COMM_PCT00") = IIf(ASP_COMM_PCT00 = -1, ASP_COMM_PCT, ASP_COMM_PCT00)
            For B As Integer = 1 To HC_CODEs.Length - 1
                rowSPTACOMX.Item("ASP_COMM_PCT" & Format(B, "00")) = IIf(ASP_COMM_PCT00 = -1, _
                                                                          IIf(ASP_COMM_PCTs(B) = -1, ASP_COMM_PCT00, ASP_COMM_PCTs(B)), _
                                                                          ASP_COMM_PCT00)
            Next
            dst.Tables("SPTACOMX").Rows.Add(rowSPTACOMX)
        Next

        For Each rowSPTACOM4 As DataRow In dst.Tables("SPTACOM4").Select("")
            '   Dim ASP_CODE As String = rowSPTACOM4.Item("ASP_CODE")
            Dim CUST_CODE As String = rowSPTACOM4.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSPTACOM4.Item("CUST_STORE_NO")
            Dim HC_CODE As String = rowSPTACOM4.Item("HC_CODE")
            Dim rowSPTACOMX As DataRow = dst.Tables("SPTACOMX").Rows.Find(New String() {CUST_STORE_NO})
            Dim B As Integer = HC_CODE_2_B(HC_CODE)
            rowSPTACOMX.Item("ASP_COMM_PCT" & Format(B, "00")) = rowSPTACOM4.Item("ASP_COMM_PCT")
        Next

        Sort_grdColumns(grdSPTACOMX, "CUST_STORE_NO")
    End Sub

    Private Sub grdSPTACOMX_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTACOMX.AfterRowUpdate
        Dim ASP_CODE As String = Absx1.txtFor("ASP_CODE").Text
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Value
        Dim ASP_COMM_PCT00 As Decimal = Val(e.Row.Cells("ASP_COMM_PCT00").Value & "")
        Dim ASP_COMM_PCT As Decimal = Val(Absx1.numFor("ASP_COMM_PCT").Value & "")
        For B As Integer = 1 To HC_CODEs.Length - 1
            Dim HC_CODE As String = HC_CODEs(B)
            Dim rowSPTACOM3 As DataRow = dst.Tables("SPTACOM3").Rows.Find(New String() {ASP_CODE, CUST_CODE, HC_CODE})
            If rowSPTACOM3 IsNot Nothing Then
                Dim ASP_COMM_PCTHC As Decimal = Val(rowSPTACOM3.Item("ASP_COMM_PCT") & "")
                Dim ASP_COMM_PCTXX As Decimal = Val(e.Row.Cells("ASP_COMM_PCT" & Format(B, "00")).Value & "")

                Dim ASP_COMM_PCT_CALC As Decimal = IIf(ASP_COMM_PCT00 <> ASP_COMM_PCT, ASP_COMM_PCT00, IIf(ASP_COMM_PCTHC <> ASP_COMM_PCT, ASP_COMM_PCTHC, ASP_COMM_PCT))

                Dim rowSPTACOM4 As DataRow = dst.Tables("SPTACOM4").Rows.Find(New String() {ASP_CODE, CUST_CODE, CUST_STORE_NO, HC_CODE})
                If rowSPTACOM4 Is Nothing Then
                    If ASP_COMM_PCTXX <> ASP_COMM_PCT_CALC Then
                        rowSPTACOM4 = dst.Tables("SPTACOM4").NewRow
                        rowSPTACOM4.Item("ASP_CODE") = ASP_CODE
                        rowSPTACOM4.Item("CUST_CODE") = CUST_CODE
                        rowSPTACOM4.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSPTACOM4.Item("HC_CODE") = HC_CODE
                        rowSPTACOM4.Item("ASP_COMM_PCT") = ASP_COMM_PCTXX
                        dst.Tables("SPTACOM4").Rows.Add(rowSPTACOM4)
                    End If
                Else
                    If ASP_COMM_PCTXX <> ASP_COMM_PCT_CALC Then
                        rowSPTACOM4.Item("ASP_COMM_PCT") = ASP_COMM_PCTXX
                    Else
                        rowSPTACOM4.Delete()
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub grdSPTACOMX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSPTACOMX.InitializeRow
        '  If Me.IsDone Then Exit Sub
        Dim ASP_COMM_PCT As Decimal = Val(Absx1.numFor("ASP_COMM_PCT").Value & "")

        Dim C As String = "ASP_COMM_PCT00"
        Dim ASP_COMM_PCT00 As Decimal = Val(e.Row.Cells(C).Value & "")
        If ASP_COMM_PCT00 <> ASP_COMM_PCT Then
            e.Row.Cells(C).Appearance = cellAppearance
            e.Row.Cells("CUST_STORE_NO").Appearance = cellAppearance
        End If

        For B As Integer = 1 To HC_CODEs.Length - 1
            C = "ASP_COMM_PCT" & Format(B, "00")
            Dim ASP_COMM_PCTXX As Decimal = Val(e.Row.Cells(C).Value & "")
            If ASP_COMM_PCTXX <> IIf(ASP_COMM_PCTs(B) = -1, ASP_COMM_PCTXX, ASP_COMM_PCTs(B)) Then
                e.Row.Cells(C).Appearance = cellAppearance
            End If
        Next
    End Sub

    Sub Get_Pct_by_HC(ASP_COMM_PCT As Decimal)
        ReDim ASP_COMM_PCTs(HC_CODEs.Length - 1)
        For B As Integer = 1 To HC_CODEs.Length - 1
            Dim rowSPTACOM3 As DataRow = dst.Tables("SPTACOM3").Rows.Find(New String() {Absx1.txtFor("ASP_CODE").Text, Absx1.txtFor("CUST_CODE").Text, HC_CODEs(B)})
            If rowSPTACOM3 IsNot Nothing Then
                ASP_COMM_PCTs(B) = Val(rowSPTACOM3.Item("ASP_COMM_PCT") & "")
            Else
                ASP_COMM_PCTs(B) = -1 ' ASP_COMM_PCT
            End If
        Next
    End Sub

    Private Sub optASP_COMM_BASIS_ValueChanged(sender As Object, e As EventArgs) Handles optASP_COMM_BASIS.ValueChanged
        Setup_BASIS()

    End Sub

    Sub Setup_BASIS()
        If Me.SELECTION_NO = 0 Then Exit Sub

        SplitContainer1.Panel2Collapsed = (optASP_COMM_BASIS.Value & "" = "2")
        SplitContainer2.Panel2Collapsed = (optASP_COMM_BASIS.Value & "" = "2")
        lblACOM_COMM_BASED_ON.Visible = Not (optASP_COMM_BASIS.Value & "" = "2")
        lblACOM_COMM_PCT.Visible = Not (optASP_COMM_BASIS.Value & "" = "2")
        numACOM_COMM_PCT.Visible = Not (optASP_COMM_BASIS.Value & "" = "2")

        With grdSPTACOM2.DisplayLayout.Bands(0)
            If (optASP_COMM_BASIS.Value & "" = "2") Then
                .Columns("ASP_COMM_PCT").Header.Caption = "Amount"
                grdSPTACOM2.Text = "Fixed Expense by Store"
            Else
                .Columns("ASP_COMM_PCT").Header.Caption = "Comm%"
                grdSPTACOM2.Text = "Exception by Store"
            End If
        End With
        chkASP_COMM_BY_INVOICE.Visible = Not (optASP_COMM_BASIS.Value & "" <> "1")

    End Sub

    Private Sub grdSPTACOM2_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSPTACOM2.InitializeLayout

    End Sub

    Private Sub grdSPTACOMX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdSPTACOMX.InitializeLayout

    End Sub

    Private Sub grdSPTACOMX_Click(sender As Object, e As EventArgs) Handles grdSPTACOMX.Click
        'grdSPTACOMX.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
        'grdSPTACOMX.DisplayLayout.Override.SelectTypeCol = UltraWinGrid.SelectType.None

 
    End Sub

    Private Sub grdSPTACOMX_MouseUp(sender As Object, e As MouseEventArgs) Handles grdSPTACOMX.MouseUp
        'Dim pt As System.Drawing.Point = New System.Drawing.Point(e.X, e.Y)

        'Dim elem As Infragistics.Win.UIElement = grdSPTACOMX.DisplayLayout.UIElement.ElementFromPoint(pt)
        'If elem Is Nothing Then
        '    Exit Sub
        'End If


        'If elem.Parent.GetType.Equals(GetType(Infragistics.Win.UltraWinGrid.HeaderUIElement)) Then
        '    Dim mouseupColumn As UltraWinGrid.UltraGridColumn = DirectCast(elem.GetContext(GetType(UltraWinGrid.UltraGridColumn), True), UltraWinGrid.UltraGridColumn)

        '    Dim COLUMN_NAME As String = "?"
        '    If mouseupColumn IsNot Nothing Then
        '        COLUMN_NAME = mouseupColumn.key

        '    End If
        '    MsgBox(elem.Parent.ToString & " for " & COLUMN_NAME)
        'Else
        '    MsgBox(elem.Parent.ToString)
        'End If

    End Sub
End Class