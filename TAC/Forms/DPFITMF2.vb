Imports Infragistics.Win.UltraWinGrid

Public Class DPFITMF2
    Public ITEM_CODE As String
    Public MARKET_CODE As String
    Public OPS_YYYYPP_FC As String
    Public FORECAST As Int64
    Public allow_new_notes As Boolean = False
    Public update_was_clicked As Boolean = False
    Public sqlDPTITMF2 As String = ""
    'Public EntryMode As String = ""
    Dim PREVIOUS_NOTE As String = ""


    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        ' Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select DPTITMF2.*" & vbCrLf _
                & " from DPTITMF2 where DPTITMF2.ITEM_CODE = :PARM1" & vbCrLf _
                & " and DPTITMF2.MARKET_CODE = :PARM2" & vbCrLf _
                & " and DPTITMF2.OPS_YYYYPP_FC = :PARM3"
            Create_TDA(.Tables.Add, "DPTITMF2", "**", 0, True, "VVV")
            sqlDPTITMF2 = ASCMAIN1.sql
        End With

        grdDPTITMF2.DataSource = dst.Tables("DPTITMF2")

        Create_Summary(grdDPTITMF2, "FORECAST_NOTE", "Count")


        With grdDPTITMF2.DisplayLayout.Override
            If allow_new_notes Then
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With
        With grdDPTITMF2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
       {"OPS_YYYYPP", "ITEM_CODE", "MARKET_CODE", "OPS_YYYYPP_FC", "STATUS",
        "INIT_OPER", "INIT_DATE", "LAST_OPER", "LAST_DATE", "FORECAST", "FORECAST_NOTE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Color.LightSkyBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                With grdDPTITMF2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    If COLUMN_NAME = "FORECAST_NOTE" And allow_new_notes Then
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .CellAppearance.BackColor = Color.LightYellow

                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        '.CellAppearance.BackColor = Color.Beige
                    End If
                    If COLUMN_NAME = "STATUS" Then
                        .CellAppearance.TextHAlign = HAlign.Center
                    End If

                End With
            Next
        End With


        Absx1.txtFor("ITEM_CODE").Text = ITEM_CODE
        Absx1.txtFor("MARKET_CODE").Text = MARKET_CODE
        Absx1.txtFor("OPS_YYYYPP_FC").Text = OPS_YYYYPP_FC
        Prepare_DPTITMF2()


        If allow_new_notes Then
        Else
            cmdUpdate.Visible = False
            cmdADD.Visible = False
            cmdCancel.Text = "Done"

            lblFORECAST_NOTE.Visible = False
            txtFORECAST_NOTE.Visible = False
        End If
        optCH.Value = "H"
        optCH.Value = "C"
        EntryMode = "A"
    End Sub

    'Overrides Sub Prepare_for_View_Lookup_Special _
    '(ByVal ctl As Windows.Forms.Control,
    ' ByVal COLUMN_NAME As String,
    ' Optional ByRef sql_where As String = "",
    ' Optional ByRef Cancel As Boolean = False)

    '    Select Case COLUMN_NAME
    '        Case "ORDR_FORM_CODE"
    '            sql_where = "ORDR_FORM_STATUS = 'A'"

    '    End Select
    'End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "ORDR_FORM_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Prepare_SOTFORM2()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "ORDR_FORM_CODE"
            '    Prepare_DPTITMF2()
        End Select
    End Sub
#End Region

    Private Sub cmdUpdate_Click(sender As System.Object, e As System.EventArgs) Handles cmdUpdate.Click


        'If txtFORECAST_NOTE.Text = "" Then
        '    MsgBox("No Note Entered", MsgBoxStyle.OkOnly, "Cannot Update")
        '    Exit Sub
        'End If
        '' grdDPTITMF2.Update()
        'Dim rowDPTITMF2 As DataRow = dst.Tables("DPTITMF2").NewRow
        'With rowDPTITMF2
        '    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
        '    .Item("ITEM_CODE") = ITEM_CODE
        '    .Item("MARKET_CODE") = MARKET_CODE
        '    .Item("OPS_YYYYPP_FC") = OPS_YYYYPP_FC
        '    .Item("FORECAST") = FORECAST
        '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
        '    .Item("INIT_DATE") = DATETIME_STAMP
        '    .Item("FORECAST_NOTE") = txtFORECAST_NOTE.Text
        'End With
        'dst.Tables("DPTITMF2").Rows.Add(rowDPTITMF2)

        Update_Record_TDA("DPTITMF2")

        update_was_clicked = True
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles cmdCancel.Click
        If allow_new_notes = True Then
            If MsgBox("Are you sure that you want to Cancel changes, Changes Have not been written to Database", MsgBoxStyle.YesNo, "Verification") = vbYes Then
            Else
                Exit Sub
            End If
        End If

        Me.Close()
    End Sub

    Sub Prepare_DPTITMF2()
        Fill_Records("DPTITMF2", New String() {ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC})

        grdDPTITMF2.Text = "Previously entered Forecast Notes for Item " & ITEM_CODE
    End Sub
    Private Sub grdDPTITMF2_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdDPTITMF2.BeforeRowsDeleted
        ' WALT ONLY IF NOT NOT ADDNEW OTHERWISE JUST DELETE
        Dim RECORD_NO As String = grdDPTITMF2.ActiveRow.Cells("RECORD_NO").Text
        Dim REVISION As Integer = grdDPTITMF2.ActiveRow.Cells("REVISION").Text
        Dim ROW1 As DataRow = dst.Tables("DPTITMF2").Rows.Find(New Object() {RECORD_NO, REVISION})
        If ROW1.RowState = DataRowState.Added Then

        Else
            For Each grow As UltraWinGrid.UltraGridRow In e.Rows
                RECORD_NO = grow.Cells("RECORD_NO").Value & ""
                REVISION = grow.Cells("REVISION").Value
                If grow.Cells("INIT_OPER").Value & "" <> ASCMAIN1.USER_ID Then
                    e.Cancel = True
                    MsgBox("Cannot Delete Forecast Note Created by Others", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit For
                End If
                If grow.Cells("STATUS").Value & "" = "D" Then
                    e.Cancel = True
                    MsgBox("Cannot Delete a Forecast Note Already Deleted", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit For
                End If
                ' WALT AND NOT ADDNEW
                If e.Cancel <> True Then
                    Dim row As DataRow = dst.Tables("DPTITMF2").Rows.Find(New String() {RECORD_NO, REVISION})
                    If Not row Is Nothing Then
                        row.Item("STATUS") = "D"
                        row.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        row.Item("LAST_DATE") = DATETIME_STAMP
                    End If
                    e.Cancel = True

                End If
            Next

        End If

    End Sub

    Private Sub grdDPTITMF2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdDPTITMF2.InitializeRow
        If e.Row.Band.Key = "DPTITMF2" Then
            If e.Row.Cells("STATUS").Value & "" = "D" Then
                'e.Row.Cells("OPS_YYYYPP").Appearance.BackColor = Drawing.Color.IndianRed
                'e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.IndianRed
                'e.Row.Cells("MARKET_CODE").Appearance.BackColor = Drawing.Color.IndianRed
                'e.Row.Cells("OPS_YYYYPP_FC").Appearance.BackColor = Drawing.Color.IndianRed
                'e.Row.Cells("FORECAST").Appearance.BackColor = Drawing.Color.RosyBrown
                'e.Row.Cells("INIT_OPER").Appearance.BackColor = Drawing.Color.RosyBrown
                'e.Row.Cells("INIT_DATE").Appearance.BackColor = Drawing.Color.SandyBrown
                'e.Row.Cells("FORECAST_NOTE").Appearance.BackColor = Drawing.Color.RosyBrown
                'e.Row.Cells("LAST_OPER").Appearance.BackColor = Drawing.Color.RosyBrown
                'e.Row.Cells("LAST_DATE").Appearance.BackColor = Drawing.Color.RosyBrown
                e.Row.Cells("STATUS").Appearance.ForeColor = Drawing.Color.Red
            ElseIf e.Row.Cells("STATUS").Value & "" = "C" Then
                e.Row.Cells("STATUS").Appearance.BackColor = Drawing.Color.Yellow
            End If
            'If optCH.Value = "C" And e.Row.Cells("STATUS").Value & "" = "" Then
            '    e.Row.Hidden = True
            'End If
        Else

        End If
    End Sub

    Private Sub optCH_ValueChanged(sender As Object, e As EventArgs) Handles optCH.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        SET_NOTES()

    End Sub
    Sub SET_NOTES()
        ' Dim sqlDPTITMF2 As String = ""
        Me.Cursor = Cursors.WaitCursor
        'EnforceConstraints(False)
        Dim dvw As DataView = DirectCast(grdDPTITMF2.DataSource, DataTable).DefaultView

        If optCH.Value = "C" Then
            dvw.RowFilter = "ISNULL(STATUS,'') = ''"

            'ASCMAIN1.sql = sqlDPTITMF2 & " and STATUS IS NULL"
            'Fill_Records("DPTITMF2", New String() {ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC}, True, ASCMAIN1.sql)
            grdDPTITMF2.Text = "Current entered Forecast Notes for Item "
            With grdDPTITMF2.DisplayLayout.Override
                If allow_new_notes Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
            cmdADD.Enabled = True
        ElseIf optCH.Value = "H" Then
            dvw.RowFilter = ""
            'ASCMAIN1.sql = sqlDPTITMF2
            'Fill_Records("DPTITMF2", New String() {ITEM_CODE, MARKET_CODE, OPS_YYYYPP_FC}, True, ASCMAIN1.sql)
            grdDPTITMF2.Text = "History of entered Forecast Notes for Item"

            With grdDPTITMF2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            cmdADD.Enabled = False
        End If
        '' With grdDPTITMF2.DisplayLayout.Bands(0)
        ''     For Each COLUMN_NAME As String In New String() _
        ''{"FORECAST_NOTE"}
        ''         With grdDPTITMF2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
        ''             If COLUMN_NAME = "FORECAST_NOTE" Then
        ''                 .CellAppearance.BackColor = Color.LightYellow
        ''             End If
        ''         End With
        ''     Next
        '' End With



        'EnforceConstraints(True)
        Sort_grdColumns(grdDPTITMF2, "RECORD_NO, REVISION".ToLower)

    End Sub

    Private Sub grdDPTITMF2_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdDPTITMF2.AfterRowsDeleted
        Stop
        ' Update_Record_TDA("DPTITMF2")
        optCH.Value = "C"
    End Sub

    Private Sub grdDPTITMF2_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdDPTITMF2.DoubleClickRow
        'If e.Row.Band.Key = "DPTITMF2" Then
        '    Dim FORECAST_NOTE As String = e.Row.Cells("FORECAST_NOTE").Value
        '    EntryMode = "M"
        '    txtFORECAST_NOTE.Text = FORECAST_NOTE
        'End If

    End Sub

    Private Sub cmdADD_Click(sender As Object, e As EventArgs) Handles cmdADD.Click
        If txtFORECAST_NOTE.Text = "" Then
            MsgBox("No Note Entered", MsgBoxStyle.OkOnly, "Cannot Update")
            Exit Sub
        End If
        ' grdDPTITMF2.Update()

        Dim RECORD_NO As String = ASCMAIN1.Next_Control_No("DPTITMF2.RECORD_NO")
        Dim REVISION = 1

        Dim rowDPTITMF2 As DataRow = dst.Tables("DPTITMF2").NewRow
        With rowDPTITMF2
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("ITEM_CODE") = ITEM_CODE
            .Item("MARKET_CODE") = MARKET_CODE
            .Item("OPS_YYYYPP_FC") = OPS_YYYYPP_FC
            .Item("FORECAST") = FORECAST
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("FORECAST_NOTE") = txtFORECAST_NOTE.Text
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("STATUS") = Null
            .Item("RECORD_NO") = RECORD_NO
            .Item("REVISION") = REVISION
        End With
        dst.Tables("DPTITMF2").Rows.Add(rowDPTITMF2)

        Sort_grdColumns(grdDPTITMF2, "RECORD_NO, REVISION".ToLower)

    End Sub


    Private Sub grdDPTITMF2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdDPTITMF2.AfterRowActivate
        PREVIOUS_NOTE = ""
        PREVIOUS_NOTE = grdDPTITMF2.ActiveRow.Cells("FORECAST_NOTE").Text
    End Sub

    Private Sub grdDPTITMF2_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdDPTITMF2.BeforeRowUpdate
        Dim RECORD_NO As String = grdDPTITMF2.ActiveRow.Cells("RECORD_NO").Text
        Dim REVISION As Integer = grdDPTITMF2.ActiveRow.Cells("REVISION").Text
        Dim ROW1 As DataRow = dst.Tables("DPTITMF2").Rows.Find(New Object() {RECORD_NO, REVISION})
        If ROW1.RowState = DataRowState.Added Then
        Else
            If PREVIOUS_NOTE <> grdDPTITMF2.ActiveRow.Cells("FORECAST_NOTE").Text Then
                ' ADD NEW RECORD WITH NEW REVISION NO AND CHANGE THE OLD ONE TO 'C' STATUS  AND THEN CONTINUE

                '        Dim REVISION = 1

                ' Dim RECORD_NO As String = grdDPTITMF2.ActiveRow.Cells("RECORD_NO").Text


                Dim REVISION_max As Integer = Val(dst.Tables("DPTITMF2").Compute("MAX(REVISION)", "RECORD_NO = '" & RECORD_NO & "'") & "")
                '   rowSOTORDR2.Item("REVISION_max") = REVISION_max + 1


                Dim rowDPTITMF2 As DataRow = dst.Tables("DPTITMF2").NewRow
                With rowDPTITMF2
                    .Item("OPS_YYYYPP") = grdDPTITMF2.ActiveRow.Cells("OPS_YYYYPP").Value
                    .Item("ITEM_CODE") = grdDPTITMF2.ActiveRow.Cells("ITEM_CODE").Value
                    .Item("MARKET_CODE") = grdDPTITMF2.ActiveRow.Cells("MARKET_CODE").Value
                    .Item("OPS_YYYYPP_FC") = grdDPTITMF2.ActiveRow.Cells("OPS_YYYYPP_FC").Value
                    .Item("FORECAST") = grdDPTITMF2.ActiveRow.Cells("FORECAST").Value
                    .Item("INIT_OPER") = grdDPTITMF2.ActiveRow.Cells("INIT_OPER").Value
                    .Item("INIT_DATE") = grdDPTITMF2.ActiveRow.Cells("INIT_DATE").Value
                    .Item("FORECAST_NOTE") = PREVIOUS_NOTE
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("STATUS") = "C"
                    .Item("RECORD_NO") = grdDPTITMF2.ActiveRow.Cells("RECORD_NO").Text
                    .Item("REVISION") = REVISION_max + 1
                End With
                dst.Tables("DPTITMF2").Rows.Add(rowDPTITMF2)

                grdDPTITMF2.ActiveRow.Cells("STATUS").Value = ""
                '     grdDPTITMF2.ActiveRow.Cells("FORECAST_NOTE").Value = PREVIOUS_NOTE
                grdDPTITMF2.ActiveRow.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
                grdDPTITMF2.ActiveRow.Cells("LAST_DATE").Value = DATETIME_STAMP

            End If


        End If

    End Sub
End Class