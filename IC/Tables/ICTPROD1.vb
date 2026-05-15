Public Class ICTPROD1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ITEM_POS_MAX, ITEM_POS_MIN, ITEM_MIN_DAYS_SUPPLY, ITEM_ABC_PARMS_LOCKED, ITEM_BUFFER_QTY, ITEM_BUFFER_PCT" & vbCrLf _
                & " from ICTITEM1" & vbCrLf _
                & " where ICTITEM1.PROD_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTITEMI", "**", 0, False, "V", 1)
        End With

        grdICTITEMI.DataSource = dst.Tables("ICTITEMI")

        With grdICTITEMI.DisplayLayout.Bands(0)
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.False
        End With
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTITEMI, "SS", "Show Filter", "Show GroupBox")
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

            Case "grdICTITEMI"
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
            'Case "Add Codes"
            '    If grd.Name = "grdICTITEMI" Then
            '        Add_Codes(grdICTITEMI, "ICTITEM1", "ITEM_CODE", "Items")
            '    End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim ORDR_FORM_CODE As String = Absx1.txtFor("ORDR_FORM_CODE").Text
        'Dim sqlDelete = "ORDR_FORM_CODE = '" & ORDR_FORM_CODE & "'"
        'Update_Record_TDA("SOTFORM2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("ICTITEMI", New String() {Absx1.txtFor("PROD_CODE").Text})
        Sort_grdColumns(grdICTITEMI, "ITEM_CODE")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"ICTITEMI"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdICTITEMI.Enabled = tf
        btnApply.Visible = tf And (EntryMode = "View")
        cmdRefreshItems.Visible = tf And (EntryMode = "View")
        btnApplyBuffer.Visible = tf And (EntryMode = "View")
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        'For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTITEMI}
        '    With grd.DisplayLayout.Override
        '        If EntryMode = "New" Or EntryMode = "Edit" Then
        '            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        '            .AllowUpdate = DefaultableBoolean.True
        '            .AllowDelete = DefaultableBoolean.True
        '        Else
        '            .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '            .AllowUpdate = DefaultableBoolean.False
        '            .AllowDelete = DefaultableBoolean.False
        '        End If
        '    End With
        'Next
        grpABC.Visible = tf AndAlso Absx1.txtFor("COST_CATGY_CODE").Text = "N"

    End Sub

#End Region

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click

        If Not ASCMAIN1.Logical_Lock("T", "ICTITEM1", False, True, True, 1) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("F", "DPFPLAN1", False, True, True, 1) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("R", "DPRMUPD1", False, True, True, 1) Then Exit Sub

        Dim PROD_CODE As String = Absx1.txtFor("PROD_CODE").Text
        Dim PROD_MAX_POS As Decimal = Val(Absx1.numFor("PROD_MAX_POS").Value & "")
        Dim PROD_MIN_POS As Decimal = Val(Absx1.numFor("PROD_MIN_POS").Value & "")
        Dim PROD_MIN_DAYS_SUPPLY As Decimal = Val(Absx1.numFor("PROD_MIN_DAYS_SUPPLY").Value & "")

        Dim EMsg As String = ""
        If PROD_MAX_POS <= 0 Then
            EMsg &= vbCr & "Max Pos must be greater than 0"
        End If
        If PROD_MIN_POS <= 0 Then
            EMsg &= vbCr & "Min Pos must be greater than 0"
        End If
        If PROD_MIN_POS > PROD_MAX_POS Then
            EMsg &= vbCr & "Min Pos must not be greater than Max Pos"
        End If
        If PROD_MIN_DAYS_SUPPLY < 0 Then
            EMsg &= vbCr & "Min Days of Supply cannot be Negative"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If

        If MsgBox($"Do you want to apply the ABC Parameters shown to all Items coded to Product Code {PROD_CODE}?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
            ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_POS_MAX = :PARM1, ITEM_POS_MIN = :PARM2, ITEM_MIN_DAYS_SUPPLY = :PARM3 where PROD_CODE = :PARM4 and NVL(ITEM_ABC_PARMS_LOCKED,'0') <> '1'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NNNV", New Object() {PROD_MAX_POS, PROD_MIN_POS, PROD_MIN_DAYS_SUPPLY, PROD_CODE})
        Else
            Exit Sub
        End If

        Fill_Records("ICTITEMI", PROD_CODE)
        Sort_grdColumns(grdICTITEMI, "ITEM_CODE")

        TAC.TACMAIN1.Record_Event("ICTPROD1", PROD_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "MASSABC", $"Mass Update ABC Max Pos {PROD_MAX_POS} Max Pos {PROD_MIN_POS} MDS {PROD_MIN_DAYS_SUPPLY}", "", Me.Name)
        MsgBox("Mass Update Complete", MsgBoxStyle.OkOnly, "Verification")
        ASCMAIN1.MultiTask_Release(,, 1)

    End Sub

    Private Sub cmdRefreshItems_Click(sender As Object, e As EventArgs) Handles cmdRefreshItems.Click
        Dim PROD_CODE As String = Absx1.txtFor("PROD_CODE").Text
        Fill_Records("ICTITEMI", PROD_CODE)
        Sort_grdColumns(grdICTITEMI, "ITEM_CODE")
    End Sub

    Private Sub btnApplyBuffer_Click(sender As Object, e As EventArgs) Handles btnApplyBuffer.Click

        If Not ASCMAIN1.Logical_Lock("T", "ICTITEM1", False, True, True, 1) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("F", "DPFPLAN1", False, True, True, 1) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("R", "DPRMUPD1", False, True, True, 1) Then Exit Sub

        Dim PROD_CODE As String = Absx1.txtFor("PROD_CODE").Text
        Dim PROD_BUFFER_QTY As Integer = Val(Absx1.numFor("PROD_BUFFER_QTY").Value & "")
        Dim PROD_BUFFER_PCT As Integer = Val(Absx1.numFor("PROD_BUFFER_PCT").Value & "")

        Dim EMsg As String = ""
        If PROD_BUFFER_QTY < 0 Then
            EMsg &= vbCr & "Buffer Qty must be greater than 0"
        End If
        If PROD_BUFFER_PCT < 0 Then
            EMsg &= vbCr & "Buffer Pct must be greater than 0"
        End If
        If PROD_BUFFER_QTY > 0 And PROD_BUFFER_PCT > 0 Then
            EMsg &= vbCr & "You must use either Qty or Pct, but not both"
        End If

        If EMsg <> "" Then
            MsgBox(Mid(EMsg, 2), MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If

        If MsgBox($"Do you want to apply the Buffer Parameters shown to all Items coded to Product Code {PROD_CODE}?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
            ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_BUFFER_QTY = :PARM1, ITEM_BUFFER_PCT = :PARM2 where PROD_CODE = :PARM3"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NNV", New Object() {PROD_BUFFER_QTY, PROD_BUFFER_PCT, PROD_CODE})
        Else
            Exit Sub
        End If

        Fill_Records("ICTITEMI", PROD_CODE)
        Sort_grdColumns(grdICTITEMI, "ITEM_CODE")

        TAC.TACMAIN1.Record_Event("ICTPROD1", PROD_CODE, Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID, "MASSBUF", $"Mass Update Buffer Qty {PROD_BUFFER_QTY} Pct {PROD_BUFFER_QTY}", "", Me.Name)
        MsgBox("Mass Update Complete", MsgBoxStyle.OkOnly, "Verification")
        ASCMAIN1.MultiTask_Release(,, 1)

    End Sub
End Class