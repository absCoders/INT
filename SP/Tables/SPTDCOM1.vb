Imports Infragistics.Win.UltraWinGrid

Public Class SPTDCOM1

    Dim HC_CODEs() As String
    Dim HC_CODE_2_B As New Dictionary(Of String, Integer)
    Dim DEMO_COMM_PCTs() As Decimal
    Dim cellAppearance As New Infragistics.Win.Appearance
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SPTDCOM2.*" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_LOCATION, ARTCUST2.MALL_CODE" & vbCrLf _
                & " from SPTDCOM2,ARTCUST2" & vbCrLf _
                & " where SPTDCOM2.CUST_CODE = :PARM1" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SPTDCOM2.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SPTDCOM2.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SPTDCOM2", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select SPTDCOM3.*" & vbCrLf _
                & ", ICTCOLL0.HC_NAME" & vbCrLf _
                & " from SPTDCOM3,ICTCOLL0" & vbCrLf _
                & " where SPTDCOM3.CUST_CODE = :PARM1" & vbCrLf _
                & "   and ICTCOLL0.HC_CODE = SPTDCOM3.HC_CODE"
            Create_TDA(.Tables.Add, "SPTDCOM3", "**", 0, True, "V", 2)

            Create_TDA(.Tables.Add, "SPTDCOM4", "*", 1)

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, False)
            Create_TDA(.Tables.Add, "ICTCOLL0", "*", 0, False)

            With .Tables.Add("SPTDCOMX")
                .Columns.Add("CUST_STORE_NO")
                For I As Integer = 0 To 199
                    .Columns.Add("DEMO_COMM_PCT" & Format(I, "000"), GetType(System.Decimal))
                Next
                .PrimaryKey = New DataColumn() { .Columns("CUST_STORE_NO")}
            End With

            For Each T As String In New String() {"SPTDCOM1", "SPTDCOM2", "SPTDCOM3", "SPTDCOM4"}
                ASCMAIN1.sql = "Select * from " & T
                Create_TDA(.Tables.Add, T & "Z", "**", 0, False)
            Next

            Create_Relation("SPTDCOM1Z", "SPTDCOM2Z", "CUST_CODE")
            Create_Relation("SPTDCOM1Z", "SPTDCOM3Z", "CUST_CODE")
            Create_Relation("SPTDCOM1Z", "SPTDCOM4Z", "CUST_CODE")

        End With

        Fill_Records("ICTCOLL0")

        grdSPTDCOM2.DataSource = dst.Tables("SPTDCOM2")
        grdSPTDCOM3.DataSource = dst.Tables("SPTDCOM3")
        grdSPTDCOMX.DataSource = dst.Tables("SPTDCOMX")
        grdSPTDCOMZ.DataSource = dst.Tables("SPTDCOM1Z")

        With grdSPTDCOM2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightBlue
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If gcol.Key = "DEMO_COMM_PCT" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdSPTDCOM3.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackColor2 = Drawing.Color.LightGreen
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If gcol.Key = "DEMO_COMM_PCT" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdSPTDCOMX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = Drawing.Color.Orange
                    If gcol.Key = "CUST_STORE_NO" Or gcol.Key = "DEMO_COMM_PCT000" Then
                        gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    End If
                End With
                If gcol.Key.StartsWith("DEMO_COMM_PCT") And gcol.Key <> "DEMO_COMM_PCT000" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            .Columns("CUST_STORE_NO").Header.Fixed = True
            ReDim HC_CODEs(dst.Tables("ICTCOLL0").Rows.Count)
            HC_CODE_2_B.Clear()

            Dim B As Integer = 0
            With .Columns("DEMO_COMM_PCT" & Format(B, "000"))
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
                With .Columns("DEMO_COMM_PCT" & Format(B, "000"))
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
        With grdSPTDCOM2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO")
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
        With grdSPTDCOM3.DisplayLayout.Bands(0).Columns("HC_CODE")
            .ValueList = uddHC_CODE
            '.ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
        End With

        cellAppearance.ForeColor = Drawing.Color.Red

        ASCMAIN1.Add_Value_List(grdSPTDCOMZ, "DEMO_COMM_BASIS", Nothing, New String() {":", "0:EDI Retail Sales", "1:Gross Shipments @ Net Price", "2:Gross Shipments @ Retail"})

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSPTDCOM2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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

            Case "grdSPTDCOM2"
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
            Case "DEMO_COMM_PCT"
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

                If optDEMO_COMM_BASIS.Value & "" = "" Then
                    EMsg &= vbCr & "Please Choose a Commission Basis"
                End If

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

        WriteAuditTrail("SPTDCOM2")
        WriteAuditTrail("SPTDCOM3")
        WriteAuditTrail("SPTDCOM4")

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim sqlDelete = "CUST_CODE = '" & CUST_CODE & "'"
        Update_Record_TDA("SPTDCOM2", sqlDelete)
        Update_Record_TDA("SPTDCOM3", sqlDelete)
        Update_Record_TDA("SPTDCOM4", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)

        Fill_Records("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text})

        Fill_Records("SPTDCOM2", New String() {Absx1.txtFor("CUST_CODE").Text})
        Sort_grdColumns(grdSPTDCOM2, "CUST_STORE_NO")
        Fill_Records("SPTDCOM3", New String() {Absx1.txtFor("CUST_CODE").Text})
        Sort_grdColumns(grdSPTDCOM3, "HC_CODE")

        Fill_Records("SPTDCOM4", New String() {Absx1.txtFor("CUST_CODE").Text})

        Toggle_HC_Column_Display()
        Rebuild_StoreByHC()
        Sort_grdColumns(grdSPTDCOMX, "CUST_STORE_NO")

        grdSPTDCOMX.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)

        EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SPTDCOM2", "SPTDCOM3", "SPTDCOM4", "SPTDCOMX", "ARTCUST2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSPTDCOM2.Enabled = tf
        grdSPTDCOM3.Enabled = tf
        grdSPTDCOMX.Enabled = tf

        If ASCMAIN1.Running_in_VS Then
            'If UltraExplorerBar1.Groups("Screen Mode").CheckedItem.Key = "Audit Trail" Then
            '    Dim sql As String = "Select * from ASTAUDT1 where TABLE_NAME in ('SPTDCOM1','SPTDCOM2','SPTDCOM3')"
            '    Dim TBL As DataTable = ASCDATA1.GetDataTable(sql)
            '    dst.Tables("ASTAUDT1").Rows.Clear()
            '    For Each ROW As DataRow In TBL.Rows
            '        dst.Tables("ASTAUDT1").Rows.Add(ROW.ItemArray)
            '    Next
            '    '  dst.Tables("ASTAUDT1").Merge(TBL)
            'End If
        End If

        SplitContainer2.Visible = ScreenMode
        grdSPTDCOMZ.Visible = Not ScreenMode

        If Not ScreenMode Then
            EnforceConstraints(False)
            Fill_Records("SPTDCOM1Z")
            Fill_Records("SPTDCOM2Z")
            Fill_Records("SPTDCOM3Z")
            Fill_Records("SPTDCOM4Z")
            EnforceConstraints(True)

            grdSPTDCOMZ.Rows.ExpandAll(True)
        End If



    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTDCOM2, grdSPTDCOM3, grdSPTDCOMX}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    If grd.Name = "grdSPTDCOMX" Then
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

#Region "grdSPTDCOM2"

    Private Sub grdSPTDCOM2_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTDCOM2.AfterCellUpdate
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

    Private Sub grdSPTDCOM2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTDCOM2.AfterRowActivate

    End Sub

    Private Sub grdSPTDCOM2_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSPTDCOM2.AfterRowsDeleted
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTDCOM2_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTDCOM2.AfterRowUpdate
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTDCOM2_BeforeCellUpdate(sender As Object, e As UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSPTDCOM2.BeforeCellUpdate

    End Sub

    Private Sub grdSPTDCOM2_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTDCOM2.BeforeExitEditMode
        'If grdSPTDCOM2.ActiveRow.IsDataRow And Not grdSPTDCOM2.ActiveRow.IsAddRow Then

        'End If
        If grdSPTDCOM2.ActiveCell IsNot Nothing Then
            With grdSPTDCOM2.ActiveCell
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

    Private Sub grdSPTDCOM2_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTDCOM2.BeforeRowActivate
        With grdSPTDCOM2.DisplayLayout.Bands(0).Columns("CUST_STORE_NO")
            If e.Row.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSPTDCOM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTDCOM2.BeforeRowUpdate

        If Val(e.Row.Cells("DEMO_COMM_PCT").Value & "") < 0 Then
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

#Region "grdSPTDCOM3"

    Private Sub grdSPTDCOM3_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSPTDCOM3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "HC_CODE"
                Dim HC_CODE As String = e.Cell.Value & ""
                Dim row As DataRow = dst.Tables("ICTCOLL0").Rows.Find(HC_CODE)
                If row IsNot Nothing Then
                    e.Cell.Row.Cells("HC_NAME").Value = row.Item("HC_NAME")
                End If

        End Select
    End Sub

    Private Sub grdSPTDCOM3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSPTDCOM3.AfterRowActivate

    End Sub

    Private Sub grdSPTDCOM3_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdSPTDCOM3.AfterRowsDeleted
        Toggle_HC_Column_Display()
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTDCOM3_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTDCOM3.AfterRowUpdate
        Toggle_HC_Column_Display()
        Rebuild_StoreByHC()
    End Sub

    Private Sub grdSPTDCOM3_BeforeRowActivate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTDCOM3.BeforeRowActivate
        With grdSPTDCOM3.DisplayLayout.Bands(0).Columns("HC_CODE")
            If e.Row.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSPTDCOM3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTDCOM3.BeforeRowUpdate
        If Val(e.Row.Cells("DEMO_COMM_PCT").Value & "") < 0 Then
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
        For Each row As DataRow In dst.Tables("SPTDCOM3").Select("")
            Dim HC_CODE As String = row.Item("HC_CODE")
            HC_CODEs_to_display.Add(HC_CODE)
        Next
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SPTDCOM4"), "HC_CODE").Select()
            Dim HC_CODE As String = row.Item("HC_CODE")
            If Not HC_CODEs_to_display.Contains(HC_CODE) Then
                'dst.Tables("SPTDCOM3").Rows.Add(New String() {Absx1.txtFor("CUST_CODE").Text, HC_CODE})
                HC_CODEs_to_display.Add(HC_CODE)
            End If


        Next

        For b As Integer = 1 To HC_CODEs.Length - 1
            With grdSPTDCOMX.DisplayLayout.Bands(0).Columns("DEMO_COMM_PCT" & Format(b, "000"))
                Dim HC_CODE As String = HC_CODEs(b)
                .Hidden = Not HC_CODEs_to_display.Contains(HC_CODE)
            End With
        Next

        Dim DEMO_COMM_PCT As Decimal = Val(Absx1.numFor("DEMO_COMM_PCT").Value & "")
        Get_Pct_by_HC(DEMO_COMM_PCT)
    End Sub
    Sub Rebuild_StoreByHC()
        dst.Tables("SPTDCOMX").Rows.Clear()

        Dim DEMO_COMM_PCT As Decimal = Val(Absx1.numFor("DEMO_COMM_PCT").Value & "")

        Dim i As Integer = 0
        For Each row As DataRow In dst.Tables("ARTCUST2").Select("")
            i += 1
            Dim rowSPTDCOMX As DataRow = dst.Tables("SPTDCOMX").NewRow
            rowSPTDCOMX.Item("CUST_STORE_NO") = row.Item("CUST_STORE_NO")
            Dim DEMO_COMM_PCT000 As Decimal = -1
            Dim rowSPTDCOM2 As DataRow = dst.Tables("SPTDCOM2").Rows.Find(New String() {row.Item("CUST_CODE"), row.Item("CUST_STORE_NO")})
            If rowSPTDCOM2 IsNot Nothing Then
                DEMO_COMM_PCT000 = rowSPTDCOM2.Item("DEMO_COMM_PCT")
            End If
            rowSPTDCOMX.Item("DEMO_COMM_PCT000") = IIf(DEMO_COMM_PCT000 = -1, DEMO_COMM_PCT, DEMO_COMM_PCT000)
            For B As Integer = 1 To HC_CODEs.Length - 1
                rowSPTDCOMX.Item("DEMO_COMM_PCT" & Format(B, "000")) = IIf(DEMO_COMM_PCT000 = -1,
                                                                          IIf(DEMO_COMM_PCTs(B) = -1, DEMO_COMM_PCT000, DEMO_COMM_PCTs(B)),
                                                                          DEMO_COMM_PCT000)
            Next
            dst.Tables("SPTDCOMX").Rows.Add(rowSPTDCOMX)
        Next

        For Each rowSPTDCOM4 As DataRow In dst.Tables("SPTDCOM4").Select("")
            Dim CUST_CODE As String = rowSPTDCOM4.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSPTDCOM4.Item("CUST_STORE_NO")
            Dim HC_CODE As String = rowSPTDCOM4.Item("HC_CODE")
            Dim rowSPTDCOMX As DataRow = dst.Tables("SPTDCOMX").Rows.Find(New String() {CUST_STORE_NO})
            Dim B As Integer = HC_CODE_2_B(HC_CODE)
            rowSPTDCOMX.Item("DEMO_COMM_PCT" & Format(B, "000")) = rowSPTDCOM4.Item("DEMO_COMM_PCT")
        Next

        Sort_grdColumns(grdSPTDCOMX, "CUST_STORE_NO")
    End Sub

    Private Sub grdSPTDCOMX_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdSPTDCOMX.AfterRowUpdate
        ' update SPTDCOM4
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim CUST_STORE_NO As String = e.Row.Cells("CUST_STORE_NO").Value
        Dim DEMO_COMM_PCT000 As Decimal = Val(e.Row.Cells("DEMO_COMM_PCT000").Value & "")
        Dim DEMO_COMM_PCT As Decimal = Val(Absx1.numFor("DEMO_COMM_PCT").Value & "")
        For B As Integer = 1 To HC_CODEs.Length - 1
            Dim HC_CODE As String = HC_CODEs(B)
            Dim rowSPTDCOM3 As DataRow = dst.Tables("SPTDCOM3").Rows.Find(New String() {CUST_CODE, HC_CODE})
            If rowSPTDCOM3 IsNot Nothing Then
                Dim DEMO_COMM_PCTHC As Decimal = Val(rowSPTDCOM3.Item("DEMO_COMM_PCT") & "")
                Dim DEMO_COMM_PCTXX As Decimal = Val(e.Row.Cells("DEMO_COMM_PCT" & Format(B, "000")).Value & "")

                Dim DEMO_COMM_PCT_CALC As Decimal = IIf(DEMO_COMM_PCT000 <> DEMO_COMM_PCT, DEMO_COMM_PCT000, IIf(DEMO_COMM_PCTHC <> DEMO_COMM_PCT, DEMO_COMM_PCTHC, DEMO_COMM_PCT))

                Dim rowSPTDCOM4 As DataRow = dst.Tables("SPTDCOM4").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, HC_CODE})
                If rowSPTDCOM4 Is Nothing Then
                    If DEMO_COMM_PCTXX <> DEMO_COMM_PCT_CALC Then
                        rowSPTDCOM4 = dst.Tables("SPTDCOM4").NewRow
                        rowSPTDCOM4.Item("CUST_CODE") = CUST_CODE
                        rowSPTDCOM4.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSPTDCOM4.Item("HC_CODE") = HC_CODE
                        rowSPTDCOM4.Item("DEMO_COMM_PCT") = DEMO_COMM_PCTXX
                        dst.Tables("SPTDCOM4").Rows.Add(rowSPTDCOM4)
                    End If
                Else
                    If DEMO_COMM_PCTXX <> DEMO_COMM_PCT_CALC Then
                        rowSPTDCOM4.Item("DEMO_COMM_PCT") = DEMO_COMM_PCTXX
                    Else
                        rowSPTDCOM4.Delete()
                    End If
                End If
            Else
                Dim DEMO_COMM_PCTHC As Decimal = -2
                Dim DEMO_COMM_PCTXX As Decimal = Val(e.Row.Cells("DEMO_COMM_PCT" & Format(B, "000")).Value & "")

                Dim DEMO_COMM_PCT_CALC As Decimal = IIf(DEMO_COMM_PCT000 <> DEMO_COMM_PCT, DEMO_COMM_PCT000, IIf(DEMO_COMM_PCTHC <> DEMO_COMM_PCT, DEMO_COMM_PCTHC, DEMO_COMM_PCT))

                Dim rowSPTDCOM4 As DataRow = dst.Tables("SPTDCOM4").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO, HC_CODE})
                If rowSPTDCOM4 Is Nothing Then
                    If DEMO_COMM_PCTXX <> -1 Then
                        rowSPTDCOM4 = dst.Tables("SPTDCOM4").NewRow
                        rowSPTDCOM4.Item("CUST_CODE") = CUST_CODE
                        rowSPTDCOM4.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSPTDCOM4.Item("HC_CODE") = HC_CODE
                        rowSPTDCOM4.Item("DEMO_COMM_PCT") = DEMO_COMM_PCTXX
                        dst.Tables("SPTDCOM4").Rows.Add(rowSPTDCOM4)
                    End If
                Else
                    If DEMO_COMM_PCTXX <> DEMO_COMM_PCT_CALC And DEMO_COMM_PCTXX <> -1 Then
                        rowSPTDCOM4.Item("DEMO_COMM_PCT") = DEMO_COMM_PCTXX
                    Else
                        rowSPTDCOM4.Delete()
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub grdSPTDCOMX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSPTDCOMX.InitializeRow
        '  If Me.IsDone Then Exit Sub
        Dim DEMO_COMM_PCT As Decimal = Val(Absx1.numFor("DEMO_COMM_PCT").Value & "")

        Dim C As String = "DEMO_COMM_PCT000"
        Dim DEMO_COMM_PCT000 As Decimal = Val(e.Row.Cells(C).Value & "")
        If DEMO_COMM_PCT000 <> DEMO_COMM_PCT Then
            e.Row.Cells(C).Appearance = cellAppearance
            e.Row.Cells("CUST_STORE_NO").Appearance = cellAppearance
        End If

        For B As Integer = 1 To HC_CODEs.Length - 1
            C = "DEMO_COMM_PCT" & Format(B, "000")
            Dim DEMO_COMM_PCTXX As Decimal = Val(e.Row.Cells(C).Value & "")
            If DEMO_COMM_PCTXX <> IIf(DEMO_COMM_PCTs(B) = -1, DEMO_COMM_PCT000, DEMO_COMM_PCTs(B)) Then
                If DEMO_COMM_PCTXX = -1 Then
                    e.Row.Cells(C).Appearance.ForeColor = System.Drawing.Color.Transparent
                Else
                    e.Row.Cells(C).Appearance = cellAppearance
                End If

            End If
        Next
    End Sub

    Sub Get_Pct_by_HC(DEMO_COMM_PCT As Decimal)
        ReDim DEMO_COMM_PCTs(HC_CODEs.Length - 1)
        For B As Integer = 1 To HC_CODEs.Length - 1
            Dim rowSPTDCOM3 As DataRow = dst.Tables("SPTDCOM3").Rows.Find(New String() {Absx1.txtFor("CUST_CODE").Text, HC_CODEs(B)})
            If rowSPTDCOM3 IsNot Nothing Then
                DEMO_COMM_PCTs(B) = Val(rowSPTDCOM3.Item("DEMO_COMM_PCT") & "")
            Else
                DEMO_COMM_PCTs(B) = -1 ' DEMO_COMM_PCT
            End If
        Next
    End Sub

    Private Sub optDEMO_COMM_BASIS_ValueChanged(sender As Object, e As EventArgs) Handles optDEMO_COMM_BASIS.ValueChanged

    End Sub

    Private Sub grdSPTDCOMX_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdSPTDCOMX.DoubleClickCell
        'If e.Cell.Hidden Then e.Cell.Hidden = False
    End Sub

    Private Sub grdSPTDCOMX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSPTDCOMX.DoubleClickRow
        'Stop
        'If grdSPTDCOMX.ActiveCell.Value = -1 Then
        '    Stop
        'End If
    End Sub

    Private Sub grdSPTDCOMZ_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSPTDCOMZ.InitializeLayout

    End Sub
End Class