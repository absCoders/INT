Public Class ASTEMLS1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ASTEMLS2", "*", 1)

            ASCMAIN1.sql = "Select ASTJOBM2.JOB_STREAM_CODE, ASTJOBM1.JOB_STREAM_DESC, ASTJOBM2.REPORT_ID, ASTJOBM2.SET_ID from ASTJOBM1,ASTJOBM2 where ASTJOBM2.JOB_STREAM_CODE = ASTJOBM1.JOB_STREAM_CODE and EMAIL_LIST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ASTJOBMX", "**", 0, False, "V", 0)
        End With

        grdASTEMLS2.DataSource = dst.Tables("ASTEMLS2")
        grdASTJOBMX.DataSource = dst.Tables("ASTJOBMX")

        Create_Summary(grdASTEMLS2, "EMAIL_ADDRESS", "Count")
        Create_Summary(grdASTJOBMX, "JOB_STREAM_CODE", "Count")
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdASTEMLS2, "SS", "Show Filter", "Show GroupBox")
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

                'Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                'If Absx1.optFor("CUST_STMT_IND").Value & "" = "" Then
                '    EMsg &= vbCr & "You Must Select a Value for Statement Processing"
                'End If

                'Dim rowSOTSREP1 = LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
                'If rowSOTSREP1 Is Nothing Then
                '    EMsg &= vbCr & "Invalid Value entered for Sales Rep Code"
                'End If


        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        'Dim sqlDelete = ""
        For Each rowASTEMLS2 As DataRow In dst.Tables("ASTEMLS2").Rows
            If rowASTEMLS2.RowState = DataRowState.Added Then
                Write_Audit_Trail(rowASTEMLS2, "N")
            ElseIf rowASTEMLS2.RowState = DataRowState.Modified Then
                Write_Audit_Trail(rowASTEMLS2, "E")
            ElseIf rowASTEMLS2.RowState = DataRowState.Deleted Then
                Write_Audit_Trail(rowASTEMLS2, "D")
            End If
        Next

        Update_Record_TDA("ASTEMLS2", $"EMAIL_LIST_CODE = '{Absx1.txtFor("EMAIL_LIST_CODE").Text}'")

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        'If EntryMode = "New" Then
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        'Else
        '    ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")
        'End If
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("ASTEMLS2", New String() {Absx1.txtFor("EMAIL_LIST_CODE").Text})
        Sort_grdColumns(grdASTEMLS2, "EMAIL_ADDRESS")
        grdASTEMLS2.Text = "email Addresses for List " & Absx1.txtFor("EMAIL_LIST_CODE").Text

        Fill_Records("ASTJOBMX", New String() {Absx1.txtFor("EMAIL_LIST_CODE").Text})
        Sort_grdColumns(grdASTJOBMX, "JOB_STREAM_CODE")

        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("ASTEMLS2").Rows.Clear()
            dst.Tables("ASTJOBMX").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdASTEMLS2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        With grdASTEMLS2.DisplayLayout.Override
            If EntryMode = "New" Or EntryMode = "Edit" Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True

            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False

            End If
        End With
    End Sub

#End Region

#Region "grdASTEMLS2"

    Private Sub grdASTEMLS2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTEMLS2.AfterCellUpdate

    End Sub

    Private Sub grdASTEMLS2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTEMLS2.AfterRowActivate

    End Sub

    Private Sub grdASTEMLS2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdASTEMLS2.BeforeRowsDeleted

    End Sub

    Private Sub grdASTEMLS2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTEMLS2.BeforeRowUpdate

        ' VALIDATE EMAIL
        Dim EMAIL_ADDRESS_is_bad As Boolean = False
        If EMAIL_ADDRESS_is_bad Then
            e.Cancel = True
        End If
    End Sub
    Private Sub grdASTEMLS2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdASTEMLS2.Error
        grdASTEMLS2.PerformAction(UltraWinGrid.UltraGridAction.UndoRow)
    End Sub

    Private Sub grdASTEMLS2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTEMLS2.InitializeRow
        If e.Row.Cells("EMAIL_LIST_CODE").Text <> "" And e.Row.Cells("EMAIL_LIST_CODE").Text <> Absx1.txtFor("EMAIL_LIST_CODE").Text Then
            e.Row.Cells("EMAIL_LIST_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ASTEMLS2"), e.Row)
    End Sub

#End Region


End Class