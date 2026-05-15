Public Class SOFWORK1
    Dim WO_REF_TYPE As String
    Dim WO_REF_NO As String
    Dim frmASFBASE0 As ASFBASE0
    Dim read_only As Boolean = False
    Dim adding_a_workorder As Boolean = False
    Dim Reference1 As String
    Dim Reference2 As String
    Dim Reference3 As Date
    Dim Reference4 As Date

    Public Sub New( _
                  ByVal frmASFBASE0_in As ASFBASE0, _
                  ByVal WO_REF_TYPE_in As String, _
                  ByVal WO_REF_NO_in As String, _
                  ByVal read_only_in As Boolean, _
                  ByVal Reference1_in As String, _
                  ByVal Reference2_in As String, _
                  ByVal Reference3_in As Date, _
                  ByVal Reference4_in As Date, _
                  ByVal Caption_in As String)

        frmASFBASE0 = frmASFBASE0_in
        WO_REF_TYPE = WO_REF_TYPE_in
        WO_REF_NO = WO_REF_NO_in
        read_only = read_only_in

        Reference1 = Reference1_in
        Reference2 = Reference2_in
        Reference3 = Reference3_in
        Reference4 = Reference4_in

        Me.Text = Caption_in

        InitializeComponent()
    End Sub

    Private Sub Form_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        With dst
            'ASCMAIN1.sql = "Select * from ICTCOLR1"
            'Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            'ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
            '    & " from ICTCOLR1,ICTSTYC1" _
            '    & " where ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE" _
            '    & "   and ICTSTYC1.STYLE_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "ICTCOLRM", "**", 0, False, "V", 1)
            '.Tables("ICTCOLRM").Columns.Add("QTY", GetType(System.Int32))
            '.Tables("ICTCOLRM").Columns.Add("SEL", GetType(System.String))
        End With

        grdSOTWORK1.DataSource = frmASFBASE0.dst.Tables("SOTWORK1")
        grdSOTWORK2.DataSource = frmASFBASE0.dst.Tables("SOTWORK2")

        Create_Summary(grdSOTWORK1, "WO_NO", "Count")
        Create_Summary(grdSOTWORK2, "WO_LOG_NO", "Count")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTWORK1, grdSOTWORK2}
            With grd.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
        Next

        If read_only Then
            tabSOTWORK2.Tabs("Add to Log").Visible = False
            cmdAddWO.Visible = False
        End If

        With grdSOTWORK1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "WO_NO" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf New String() {"INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
            .Columns("WO_NO").Header.Fixed = True
            .Columns("WO_DESC").Header.Fixed = True
            .Columns("WO_DUE").Header.Fixed = True
        End With

        With grdSOTWORK2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellAppearance.BackColor = Drawing.Color.Beige
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "WO_LOG_NO" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                ElseIf New String() {"INIT_DATE", "INIT_OPER"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    gcol.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
        End With

        With grdSOTWORK2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "WO_LOG_NOTE" Or gcol.Key = "WO_LOG_QTY" Or gcol.Key = "WO_LOG_AMT" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With
        Sort_grdColumns(grdSOTWORK1, "WO_NO")
        Setup_SOTWORK1()

        ASCMAIN1.Add_Value_List(grdSOTWORK1, "WO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})
        ASCMAIN1.Add_Value_List(grdSOTWORK2, "WO_STATUS", Nothing, New String() {":", "O:Open", "C:Closed"})

        Dim VL As New ValueList
        ASCMAIN1.sql = "Select * from SOTWORKT"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "WO_TYPE_DESC")
            Dim VLI As New ValueListItem(row.Item("WO_TYPE"), row.Item("WO_TYPE_DESC"))
            VL.ValueListItems.Add(VLI)
        Next
        cbeWO_TYPE.ValueList = VL

        grdSOTWORK2.DisplayLayout.Bands(0).Columns("WO_LOG_NO").Width = grdSOTWORK1.DisplayLayout.Bands(0).Columns("WO_NO").Width
        grdSOTWORK2.DisplayLayout.Bands(0).Columns("WO_LOG_NOTE").Width = grdSOTWORK1.DisplayLayout.Bands(0).Columns("WO_DESC").Width
        grdSOTWORK2.DisplayLayout.Bands(0).Columns("WO_ASSIGNED_TO").Width = grdSOTWORK1.DisplayLayout.Bands(0).Columns("WO_ASSIGNED_TO").Width
        grdSOTWORK2.DisplayLayout.Bands(0).Columns("WO_DUE").Width = grdSOTWORK1.DisplayLayout.Bands(0).Columns("WO_DUE").Width
        grdSOTWORK2.DisplayLayout.Bands(0).Columns("WO_STATUS").Width = grdSOTWORK1.DisplayLayout.Bands(0).Columns("WO_STATUS").Width


        txtReference1.Text = Reference1
        txtReference2.Text = Reference2
        dteReference3.Value = Reference3
        dteReference4.Value = Reference4

        Select Case WO_REF_TYPE
            Case "P"
                lblWO_REF_TYPE.Text = "PO No"
                lblReference1.Text = "Supplier"
                lblReference2.Text = "PO Reference"
                lblReference3.Text = "Ship By"
                lblReference4.Text = "ETA"

            Case "S"
                lblWO_REF_TYPE.Text = "Order No"
                lblReference1.Text = "Customer"
                lblReference2.Text = "Customer PO"
                lblReference3.Text = "Ship By"
                lblReference4.Text = "Cancel"

            Case "R"
                lblWO_REF_TYPE.Text = "Reservation No"
                lblReference1.Text = "Customer"
                lblReference2.Text = "Customer PO"
                lblReference3.Text = "Ship By"
                lblReference4.Text = "Cancel"

            Case "W"
                lblWO_REF_TYPE.Text = "Work Order"
                lblReference1.Text = "Ref Type"
                lblReference2.Text = "Reference"
                lblReference3.Text = "Entered"
                lblReference4.Text = "Due"
        End Select

        txtWO_REF_NO.Text = WO_REF_NO


    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "STYLE_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Prepare_ICTCOLRM()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "STYLE_CODE"
            '    Prepare_ICTCOLRM()
        End Select
    End Sub
#End Region

    Private Sub grdSOTWORK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTWORK1.AfterRowActivate
        Setup_SOTWORK1()
        If grdSOTWORK1.ActiveRow.IsAddRow Then
            grdSOTWORK1.ActiveCell = grdSOTWORK1.ActiveRow.Cells("WO_DESC")
            grdSOTWORK1.DisplayLayout.Bands(0).Columns("WO_ASSIGNED_TO").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSOTWORK1.DisplayLayout.Bands(0).Columns("WO_ASSIGNED_TO").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Sub Setup_SOTWORK1()
        If grdSOTWORK1.ActiveRow Is Nothing OrElse (grdSOTWORK1.ActiveRow.IsAddRow Or Not grdSOTWORK1.ActiveRow.IsDataRow) Then
            tabSOTWORK2.Visible = False
        Else
            tabSOTWORK2.Visible = True
            Dim WO_NO As String = grdSOTWORK1.ActiveRow.Cells("WO_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTWORK2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "WO_NO = '" & WO_NO & "'"
            Sort_grdColumns(grdSOTWORK2, "WO_LOG_NO")
            grdSOTWORK2.Text = "Work Order Log Entries for Work Order " & WO_NO
        End If
    End Sub

    Private Sub grdSOTWORK1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTWORK1.BeforeRowUpdate
        'If e.Row.IsAddRow Then
        '    Dim WO_NO As String = ASCMAIN1.Next_Control_No("SOTWORK1.WO_NO")
        '    e.Row.Cells("WO_NO").Value = WO_NO
        '    e.Row.Cells("WO_DATE").Value = DATETIME_STAMP.Date
        '    e.Row.Cells("WO_STATUS").Value = "O"
        '    e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
        '    e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
        'End If
    End Sub

    Private Sub grdSOTWORK1_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTWORK1.AfterCellUpdate

        'Select Case e.Cell.Column.Key
        '    Case "WO_DESC"
        '        e.Cell.Row.Cells("WO_DUE").Value = dteReference3.DateTime.AddDays(-0)

        'End Select
    End Sub

    Private Sub cmdDone_Click(sender As System.Object, e As System.EventArgs) Handles cmdDone.Click
        Me.Close()
    End Sub

    Private Sub tabSOTORDR2_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSOTWORK2.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        If tabSOTWORK2.SelectedTab.Key = "Add to Log" Then

            txtWO_LOG_NOTE.Text = ""
            numWO_LOG_QTY.Value = 0
            numWO_LOG_AMT.Value = 0

            grdSOTWORK1.Enabled = False
            tabSOTWORK2.Tabs("Log").Enabled = False
            cmdDone.Enabled = False
            cmdAddWO.Enabled = False

            If adding_a_workorder Then
                tabSOTWORK2.Tabs("Add to Log").Text = "Add to Work Order"
                grpAdd.Text = "Add a New Work Order"
                lblWO_TYPE.Visible = True
                cbeWO_TYPE.Visible = True
                Set_Read_Only_for_ctl(optWO_STATUS, True)

                tabSOTWORK2.Visible = True
                chkCorrectWO_DESC.Visible = False

                cbeWO_TYPE.Value = ""
                dteWO_DUE.Value = ""
                txtWO_ASSIGNED_TO.Text = ""
                If dteReference3.Value & "" = "" Then
                    dteWO_DUE.Value = Now.Date
                Else
                    dteWO_DUE.Value = dteReference3.Value
                End If
                optWO_STATUS.Value = "O"
            Else
                tabSOTWORK2.Tabs("Add to Log").Text = "Add to Log"
                grpAdd.Text = "Work Order Log Entry for Work Order " & grdSOTWORK1.ActiveRow.Cells("WO_NO").Value
                lblWO_TYPE.Visible = False
                cbeWO_TYPE.Visible = False
                Set_Read_Only_for_ctl(optWO_STATUS, False)

                chkCorrectWO_DESC.Visible = True
                chkCorrectWO_DESC.Checked = False

                dteWO_DUE.Value = grdSOTWORK1.ActiveRow.Cells("WO_DUE").Value
                txtWO_ASSIGNED_TO.Text = grdSOTWORK1.ActiveRow.Cells("WO_ASSIGNED_TO").Value
                optWO_STATUS.Value = grdSOTWORK1.ActiveRow.Cells("WO_STATUS").Value
            End If
            txtWO_LOG_NOTE.Focus()
        Else
            grdSOTWORK1.Enabled = True
            tabSOTWORK2.Tabs("Log").Enabled = True
            cmdDone.Enabled = True
            cmdAddWO.Enabled = True
        End If
    End Sub

    Private Sub cmdAdd_Click(sender As System.Object, e As System.EventArgs) Handles cmdAdd.Click

        If optWO_STATUS.Value = "O" Then
            Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", txtWO_ASSIGNED_TO.Text)
            If rowASTUSER1 Is Nothing Then
                MsgBox("Invalid User ID specified in Assigned-To", MsgBoxStyle.OkOnly, "Cannot Add")
                Exit Sub
            End If

            If txtWO_LOG_NOTE.Text = "" Then
                MsgBox("A Description or Note Entry is Required", MsgBoxStyle.OkOnly, "Cannot Add")
                Exit Sub
            End If
        End If

        If adding_a_workorder Then
            If cbeWO_TYPE.Value & "" = "" Then
                MsgBox("Work Order Type is Mandatory", MsgBoxStyle.OkOnly, "Cannot Add")
                Exit Sub
            End If
        End If

        If optWO_STATUS.Value = "C" Then
            txtWO_ASSIGNED_TO.Text = ""
        End If

        Dim WO_NO As String = ""

        If adding_a_workorder Then
            WO_NO = ASCMAIN1.Next_Control_No("SOTWORK1.WO_NO")
            Dim rowSOTWORK1 As DataRow = frmASFBASE0.dst.Tables("SOTWORK1").NewRow
            With rowSOTWORK1
                .Item("WO_NO") = WO_NO
                .Item("WO_DATE") = DATETIME_STAMP.Date
                .Item("WO_DESC") = txtWO_LOG_NOTE.Text
                .Item("WO_DUE") = dteWO_DUE.Value
                .Item("WO_STATUS") = "O"
                .Item("WO_TYPE") = cbeWO_TYPE.Value
                .Item("WO_REF_TYPE") = WO_REF_TYPE
                .Item("WO_REF_NO") = WO_REF_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("WO_ASSIGNED_TO") = txtWO_ASSIGNED_TO.Text
                .Item("WO_QTY") = numWO_LOG_QTY.Value
                .Item("WO_AMT") = numWO_LOG_AMT.Value
            End With
            frmASFBASE0.dst.Tables("SOTWORK1").Rows.Add(rowSOTWORK1)

            tabSOTWORK2.Tabs("Add to Log").Text = "Add to Log"
            Sort_grdColumns(grdSOTWORK1, "WO_NO")
            Setup_SOTWORK1()
        Else
            WO_NO = grdSOTWORK1.ActiveRow.Cells("WO_NO").Value

            If chkCorrectWO_DESC.Checked Then
                grdSOTWORK1.ActiveRow.Cells("WO_DESC").Value = txtWO_LOG_NOTE.Text
            End If
            If optWO_STATUS.Value = "C" Then
                grdSOTWORK1.ActiveRow.Cells("WO_COMPLETED").Value = DATETIME_STAMP.Date
            End If
            grdSOTWORK1.ActiveRow.Cells("WO_STATUS").Value = optWO_STATUS.Value
            grdSOTWORK1.ActiveRow.Cells("WO_ASSIGNED_TO").Value = txtWO_ASSIGNED_TO.Text
            grdSOTWORK1.ActiveRow.Cells("WO_DUE").Value = dteWO_DUE.Value
            grdSOTWORK1.ActiveRow.Cells("LAST_DATE").Value = DATETIME_STAMP
            grdSOTWORK1.ActiveRow.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
            grdSOTWORK1.ActiveRow.Update()
        End If

        Dim rowSOTWORK2 As DataRow = frmASFBASE0.dst.Tables("SOTWORK2").NewRow
        With rowSOTWORK2
            .Item("WO_LOG_NO") = ASCMAIN1.Next_Control_No("SOTWORK2.WO_LOG_NO")
            .Item("WO_NO") = WO_NO
            .Item("WO_LOG_NOTE") = txtWO_LOG_NOTE.Text
            .Item("WO_LOG_QTY") = numWO_LOG_QTY.Value
            .Item("WO_LOG_AMT") = numWO_LOG_AMT.Value
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("WO_ASSIGNED_TO") = txtWO_ASSIGNED_TO.Text
            .Item("WO_STATUS") = optWO_STATUS.Value
            .Item("WO_DUE") = dteWO_DUE.Value
        End With
        frmASFBASE0.dst.Tables("SOTWORK2").Rows.Add(rowSOTWORK2)

        tabSOTWORK2.Tabs("Log").Enabled = True
        tabSOTWORK2.SelectedTab = tabSOTWORK2.Tabs("Log")
        adding_a_workorder = False
    End Sub

    Private Sub cmdDiscard_Click(sender As System.Object, e As System.EventArgs) Handles cmdDiscard.Click
        If adding_a_workorder Then
            tabSOTWORK2.Tabs("Add to Log").Text = "Add to Log"
            Sort_grdColumns(grdSOTWORK1, "WO_NO")
            Setup_SOTWORK1()
        End If

        tabSOTWORK2.Tabs("Log").Enabled = True
        tabSOTWORK2.SelectedTab = tabSOTWORK2.Tabs("Log")
        adding_a_workorder = False
    End Sub

    Private Sub cmdAddWO_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddWO.Click
        adding_a_workorder = True
        tabSOTWORK2.SelectedTab = tabSOTWORK2.Tabs("Add to Log")
        txtWO_LOG_NOTE.Focus()
    End Sub
End Class