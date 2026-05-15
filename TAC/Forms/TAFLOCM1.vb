Public Class TAFLOCM1

    'Private tblItremsToMove As DataTable = Nothing
    Private WHSE_CODE As String = String.Empty
    Private LOCATION_CODE_TO As String = String.Empty
    Public WHSE_TRAN_NO As String = String.Empty
    Private WHSE_TRAN_LNO As Int32 = 0

    Private Const InvalidLocation As String = "L"
    Private Const InvalidItem As String = "I"
    Private Const InvalidStyleColor As String = "S"

    Private frmDst As DataSet

#Region "Form Events"

    Public Sub New()
        frmASFBASE1 = New ABSolution.ASFBASE1
        InitializeComponent()

        dst = frmASFBASE1.clsASCBASE1.dst

        With dst
            frmASFBASE1.Create_TDA(.Tables.Add, "WHTMOVE1", "*")
            frmASFBASE1.Create_TDA(.Tables.Add, "WHTMOVE2", "*")
            frmDst = dst
        End With

        With dst.Tables("WHTMOVE2")
            .Columns.Add("ITEM_DESC", GetType(System.String))
            .Columns.Add("ERROR_CODES", GetType(System.String))
            .Columns.Add("WHSE_TRAN_QTY_ORIG", GetType(System.Int32))
        End With

    End Sub

    Private Sub TAFLOCM1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        ASCMAIN1.grdInitializeLayout(grdWHTMOVE2)
        grdWHTMOVE2.DataSource = frmDst.Tables("WHTMOVE2")

        Create_Summary(grdWHTMOVE2, "WHSE_TRAN_LNO", "Count")
        Create_Summary(grdWHTMOVE2, "WHSE_TRAN_QTY")
    End Sub

    Private Sub TAFLOCM1_MaximizedBoundsChanged(sender As Object, e As System.EventArgs) Handles Me.MaximizedBoundsChanged

    End Sub

    Private Sub TAFLOCM1_Shown(sender As Object, e As System.EventArgs) Handles Me.Shown

        ' Set the warehouse code
        If frmDst.Tables("WHTMOVE1").Rows.Count > 0 Then
            WHSE_CODE = frmDst.Tables("WHTMOVE1").Rows(0).Item("WHSE_CODE") & String.Empty
        End If

        Sort_grdColumns(grdWHTMOVE2, "WHSE_TRAN_LNO")
    End Sub

#End Region

#Region "Form Prodecures"

    Public Sub ClearItemsToMove()
        frmDst.Tables("WHTMOVE1").Rows.Clear()
        frmDst.Tables("WHTMOVE2").Rows.Clear()
        WHSE_CODE = String.Empty
        WHSE_TRAN_NO = String.Empty
        WHSE_TRAN_LNO = 0
    End Sub

    Public Sub AddItemToMove(ByVal WHSE_CODE As String, _
                               ByVal LOCATION_CODE As String, _
                               ByVal ITEM_CODE As String, _
                               ByVal QTY As Int32, _
                               Optional LOCATION_CODE_TO As String = "")

        Try
            WHSE_CODE = WHSE_CODE.Trim
            LOCATION_CODE = LOCATION_CODE.Trim
            ITEM_CODE = ITEM_CODE.Trim

            If WHSE_TRAN_NO.Length = 0 Then
                WHSE_TRAN_NO = ASCMAIN1.Next_Control_No("WHTMOVE1.WHSE_TRAN_NO")

                Dim rowWHTMOVE1 As DataRow = frmDst.Tables("WHTMOVE1").NewRow
                rowWHTMOVE1.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
                txtWHSE_TRAN_NO.Text = WHSE_TRAN_NO
                rowWHTMOVE1.Item("WHSE_TRAN_TYPE") = "M"
                rowWHTMOVE1.Item("SESSION_NO") = ASCMAIN1.SESSION_NO
                rowWHTMOVE1.Item("WHSE_CODE") = WHSE_CODE
                rowWHTMOVE1.Item("STATUS") = "U"
                frmDst.Tables("WHTMOVE1").Rows.Add(rowWHTMOVE1)
            End If

            Dim rowWHTMOVE2 As DataRow = frmDst.Tables("WHTMOVE2").NewRow
            rowWHTMOVE2.Item("WHSE_TRAN_NO") = WHSE_TRAN_NO
            WHSE_TRAN_LNO += 1
            rowWHTMOVE2.Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
            rowWHTMOVE2.Item("LOCATION_CODE_FROM") = LOCATION_CODE
            rowWHTMOVE2.Item("LOCATION_CODE_TO") = LOCATION_CODE_TO
            rowWHTMOVE2.Item("BAR_CODE") = "0000000000"
            rowWHTMOVE2.Item("WHSE_TRAN_QTY") = QTY
            rowWHTMOVE2.Item("ITEM_CODE") = ITEM_CODE
            rowWHTMOVE2.Item("STATUS") = "U"
            rowWHTMOVE2.Item("WHSE_TRAN_QTY_ORIG") = QTY
            rowWHTMOVE2.Item("ERROR_CODES") = String.Empty
            frmDst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)

            LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE})
            If cdr Is Nothing Then
                rowWHTMOVE2.Item("ERROR_CODES") &= InvalidLocation
            End If

            LookUp("ICTITEM1", ITEM_CODE)
            rowWHTMOVE2.Item("ITEM_DESC") = cdr.Item("ITEM_DESC") & String.Empty

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Add Item To Move", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As System.Windows.Forms.Control, COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & WHSE_CODE & "'"
        End Select
    End Sub

    Private Sub OverMove()

        frmDst.Tables("WHTMOVE2").AcceptChanges()

        For Each row As DataRow In ASCDATA1.SelectDistinct _
                (frmDst.Tables("WHTMOVE2"), New String() {"ITEM_CODE", "LOCATION_CODE_FROM"}).Select()
            Dim ITEM_CODE As String = row.Item("ITEM_CODE") & String.Empty
            Dim LOCATION_CODE_FROM As String = row.Item("LOCATION_CODE_FROM") & String.Empty

            Dim sqlWhere As String = "ITEM_CODE = '" & ITEM_CODE & "' and LOCATION_CODE_FROM = '" & LOCATION_CODE_FROM & "' AND ERROR_CODES = ''"

            If frmDst.Tables("WHTMOVE2").Select(sqlWhere).Length = 0 Then
                Continue For
            End If

            Dim WHSE_TRAN_QTY_ORIG As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY_ORIG)", sqlWhere) & String.Empty)
            Dim WHSE_TRAN_QTY As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY)", sqlWhere) & String.Empty)

            Dim foreColor As System.Drawing.Color = Color.Black
            If WHSE_TRAN_QTY > WHSE_TRAN_QTY_ORIG Then
                foreColor = Color.Red
            End If

            For Each gridRow As UltraWinGrid.UltraGridRow In grdWHTMOVE2.Rows
                If gridRow.Cells("ITEM_CODE").Value = ITEM_CODE AndAlso _
                    gridRow.Cells("LOCATION_CODE_FROM").Value = LOCATION_CODE_FROM Then
                    gridRow.Cells("WHSE_TRAN_QTY").Appearance.ForeColor = foreColor

                    If foreColor = Color.Black Then
                        gridRow.Cells("WHSE_TRAN_QTY").ToolTipText = ""
                    Else
                        gridRow.Cells("WHSE_TRAN_QTY").ToolTipText = "Original Qty was " & WHSE_TRAN_QTY_ORIG & ", over by " & WHSE_TRAN_QTY - WHSE_TRAN_QTY_ORIG
                    End If
                End If
            Next
        Next
    End Sub

#End Region

#Region "Buttons"

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        WHSE_TRAN_NO = ""
        Me.Close()
    End Sub

    Private Sub btnMove_Click(sender As System.Object, e As System.EventArgs) Handles btnMove.Click

        EMsg = String.Empty
        DATETIME_STAMP = DateTime.Now + ASCMAIN1.NowTSD

        Dim warnings As String = String.Empty
        Dim processed As New List(Of String)

        frmDst.Tables("WHTMOVE2").AcceptChanges()
        Dim rowWHTMOVE1 As DataRow = frmDst.Tables("WHTMOVE1").Rows(0)
        rowWHTMOVE1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTMOVE1.Item("INIT_DATE") = DATETIME_STAMP
        rowWHTMOVE1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowWHTMOVE1.Item("LAST_DATE") = DATETIME_STAMP

        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("", "ITEM_CODE, LOCATION_CODE_FROM, LOCATION_CODE_TO")
            Dim ITEM_CODE As String = rowWHTMOVE2.Item("ITEM_CODE") & String.Empty
            Dim LOCATION_CODE_FROM As String = rowWHTMOVE2.Item("LOCATION_CODE_FROM") & String.Empty
            Dim LOCATION_CODE_TO As String = (rowWHTMOVE2.Item("LOCATION_CODE_TO") & String.Empty).ToString.Trim

            rowWHTMOVE2.Item("LOCATION_CODE_TO") = LOCATION_CODE_TO

            Dim ERROR_CODES As String = (rowWHTMOVE2.Item("ERROR_CODES") & String.Empty).ToString.Trim
            If ERROR_CODES.Length > 0 Then
                warnings &= vbCr & "Style-Color " & ITEM_CODE & " will be skipped since it is invalid."
                Continue For
            End If

            If LOCATION_CODE_TO.Length = 0 Then
                warnings &= vbCr & "Style-Color " & ITEM_CODE & " will be skipped since the 'Location To' is empty."
                Continue For
            End If

            Dim sqlWhere As String = "ITEM_CODE = '" & ITEM_CODE & "' and LOCATION_CODE_FROM = '" & LOCATION_CODE_FROM & "'  and ERROR_CODES = ''"

            If processed.Contains(ITEM_CODE & "_" & LOCATION_CODE_FROM & "_" & LOCATION_CODE_TO) Then
                Continue For
            Else
                processed.Add(ITEM_CODE & "_" & LOCATION_CODE_FROM & "_" & LOCATION_CODE_TO)
            End If

            LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE_TO})
            If cdr Is Nothing Then
                MessageBox.Show("Invalid Move-To Location (" & LOCATION_CODE_TO & ") for Item: " & ITEM_CODE, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim WHSE_TRAN_QTY_ORIG As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY_ORIG)", sqlWhere) & String.Empty)
            Dim WHSE_TRAN_QTY As Int32 = Val(frmDst.Tables("WHTMOVE2").Compute("SUM(WHSE_TRAN_QTY)", sqlWhere) & String.Empty)

            If WHSE_TRAN_QTY > WHSE_TRAN_QTY_ORIG Then
                warnings &= vbCr & "Style-Color " & ITEM_CODE & " Original Qty: " & WHSE_TRAN_QTY_ORIG & ", Move Quantity: " & WHSE_TRAN_QTY
            End If

            If rowWHTMOVE2.Item("LOCATION_CODE_TO") & String.Empty = rowWHTMOVE2.Item("LOCATION_CODE_FROM") & String.Empty Then
                warnings &= vbCr & "Style-Color " & ITEM_CODE & " will be skipped since the To-Location is the same as the From-Location"
            End If

            rowWHTMOVE2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowWHTMOVE2.Item("INIT_DATE") = DATETIME_STAMP
            rowWHTMOVE2.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowWHTMOVE2.Item("LAST_DATE") = DATETIME_STAMP
        Next

        If warnings.Length > 0 Then
            Dim msg As String = "Please review the following warnings before continuing."
            msg &= vbCrLf
            msg &= warnings
            msg &= vbCrLf & vbCrLf
            msg &= "Do you want to continue?"

            If MessageBox.Show(msg, "Validation", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        Else
            If MessageBox.Show("OK to Proceed with this Move?", "Verification", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        ' Remove lines with errors
        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("ERROR_CODES <> ''")
            rowWHTMOVE2.Delete()
        Next

        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("LOCATION_CODE_TO = LOCATION_CODE_FROM")
            rowWHTMOVE2.Delete()
        Next

        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("LOCATION_CODE_TO = ''")
            rowWHTMOVE2.Delete()
        Next

        frmDst.Tables("WHTMOVE2").AcceptChanges()
        If frmDst.Tables("WHTMOVE2").Rows.Count = 0 Then
            MessageBox.Show("There are no Valid Location Movement to process.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        For Each rowWHTMOVE2 As DataRow In frmDst.Tables("WHTMOVE2").Select("")
            rowWHTMOVE2.SetAdded()
        Next

        With frmASFBASE1

            Try
                .BeginTrans()

                .clsASCBASE1.Update_Record_TDA("WHTMOVE1")
                .clsASCBASE1.Update_Record_TDA("WHTMOVE2")

                ASCDATA1.ExecuteSP("WHPMOVE1", "VNN", New Object() {WHSE_TRAN_NO, 0, 1}, New String() {"WHSE_TRAN_NO_IN", "WHSE_TRAN_LNO_IN", "S"})

                .CommitTrans("Move successful")
            Catch ex As Exception
                .Rollback(ex.Message)
            End Try
        End With
        Me.Close()
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWHTMOVE2, "B", "Split")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdARTPYMT3"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Split"
                If grdWHTMOVE2.ActiveRow Is Nothing Then Exit Sub
                Dim rowWHTMOVE2 As DataRow = frmDst.Tables("WHTMOVE2").NewRow

                rowWHTMOVE2.Item("WHSE_TRAN_NO") = grdWHTMOVE2.ActiveRow.Cells("WHSE_TRAN_NO").Value
                WHSE_TRAN_LNO += 1
                rowWHTMOVE2.Item("WHSE_TRAN_LNO") = WHSE_TRAN_LNO
                rowWHTMOVE2.Item("LOCATION_CODE_FROM") = grdWHTMOVE2.ActiveRow.Cells("LOCATION_CODE_FROM").Value
                rowWHTMOVE2.Item("LOCATION_CODE_TO") = grdWHTMOVE2.ActiveRow.Cells("LOCATION_CODE_TO").Value
                rowWHTMOVE2.Item("BAR_CODE") = grdWHTMOVE2.ActiveRow.Cells("BAR_CODE").Value
                rowWHTMOVE2.Item("WHSE_TRAN_QTY") = 0 'grdWHTMOVE2.ActiveRow.Cells("WHSE_TRAN_NO").Value
                rowWHTMOVE2.Item("ITEM_CODE") = grdWHTMOVE2.ActiveRow.Cells("ITEM_CODE").Value
                rowWHTMOVE2.Item("INIT_OPER") = grdWHTMOVE2.ActiveRow.Cells("INIT_OPER").Value
                rowWHTMOVE2.Item("INIT_DATE") = grdWHTMOVE2.ActiveRow.Cells("INIT_DATE").Value
                rowWHTMOVE2.Item("LAST_OPER") = grdWHTMOVE2.ActiveRow.Cells("LAST_OPER").Value
                rowWHTMOVE2.Item("LAST_DATE") = grdWHTMOVE2.ActiveRow.Cells("LAST_DATE").Value
                rowWHTMOVE2.Item("STATUS") = grdWHTMOVE2.ActiveRow.Cells("STATUS").Value
                rowWHTMOVE2.Item("ITEM_DESC") = grdWHTMOVE2.ActiveRow.Cells("ITEM_DESC").Value
                rowWHTMOVE2.Item("ERROR_CODES") = grdWHTMOVE2.ActiveRow.Cells("ERROR_CODES").Value
                rowWHTMOVE2.Item("WHSE_TRAN_QTY_ORIG") = 0 'grdWHTMOVE2.ActiveRow.Cells("WHSE_TRAN_NO").Value

                frmDst.Tables("WHTMOVE2").Rows.Add(rowWHTMOVE2)
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        'If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
        '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
        '    grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
        'End If

        If grd.Name = "grdARTCCPA1" Then

            Select Case e.Tool.Key
                Case "Location Inquiry"

            End Select
        End If
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdWHTPULL2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdWHTMOVE2.AfterRowsDeleted
        OverMove()
    End Sub

    Private Sub grdWHTPULL2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdWHTMOVE2.AfterRowUpdate
        OverMove()
    End Sub

    Private Sub grdWHTPULL2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdWHTMOVE2.BeforeRowUpdate
        Dim LOCATION_CODE_TO As String = e.Row.Cells("LOCATION_CODE_TO").Value & String.Empty

        LOCATION_CODE_TO = LOCATION_CODE_TO.Trim

        e.Row.Cells("LOCATION_CODE_TO").Value = LOCATION_CODE_TO

        If LOCATION_CODE_TO.Length > 0 Then
            LookUp("WHTLOCM1", New String() {WHSE_CODE, LOCATION_CODE_TO})
            If cdr Is Nothing Then
                MessageBox.Show("Invalid 'Location To' for the Warehouse.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub

    Private Sub grdWHTPULL2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdWHTMOVE2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "LOCATION_CODE_TO"
                grdClickCellButton(grdWHTMOVE2, "WHSE_CODE = '" & WHSE_CODE & "'", True, "", "LOCATION_CODE")
        End Select
    End Sub

    Private Sub grdWHTPULL2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdWHTMOVE2.InitializeRow
        ' When the data is loaded error messages are placed in the ERROR_MSG column.
        ' These become the tooltip for the column with the error
        Dim ERROR_CODES As String = (e.Row.Cells("ERROR_CODES").Value & String.Empty).ToString.Trim

        If ERROR_CODES.Contains(InvalidLocation) Then
            e.Row.Cells("LOCATION_CODE_TO").Appearance.ForeColor = Color.Red
            e.Row.Cells("LOCATION_CODE_TO").ToolTipText = "Invalid Location"
        ElseIf e.Row.Cells("LOCATION_CODE_FROM").Value.ToString.Trim = e.Row.Cells("LOCATION_CODE_TO").Value.ToString.Trim Then
            e.Row.Cells("LOCATION_CODE_TO").Appearance.ForeColor = Color.Red
            e.Row.Cells("LOCATION_CODE_TO").ToolTipText = "From and To locations are the same"
        Else
            e.Row.Cells("LOCATION_CODE_TO").Appearance.ForeColor = Color.Black
            e.Row.Cells("LOCATION_CODE_TO").ToolTipText = String.Empty
        End If

    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "LOCATION_CODE"
                ValidateSelectedNewLocation()

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "LOCATION_CODE"
                ValidateSelectedNewLocation()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "LOCATION_CODE"
                ValidateSelectedNewLocation()
        End Select
    End Sub

    Private Sub ValidateSelectedNewLocation()
        LookUp("WHTLOCM1", New String() {WHSE_CODE, MyBase.Absx1.txtFor("LOCATION_CODE").Text})

        If cdr IsNot Nothing And MyBase.Absx1.txtFor("LOCATION_CODE").Text <> "" Then
            For Each row As DataRow In frmDst.Tables("WHTMOVE2").Select("ISNULL(LOCATION_CODE_TO, '*') = '*' OR LOCATION_CODE_TO = ''")
                row.Item("LOCATION_CODE_TO") = MyBase.Absx1.txtFor("LOCATION_CODE").Text
            Next

            If LOCATION_CODE_TO.Length > 0 AndAlso LOCATION_CODE_TO <> MyBase.Absx1.txtFor("LOCATION_CODE").Text And MyBase.Absx1.txtFor("LOCATION_CODE").Text <> "" Then
                ' See if the user wants to change the location
                Dim numdiff As Int32 = frmDst.Tables("WHTMOVE2").Select("ISNULL(LOCATION_CODE_TO, '*') = '" & LOCATION_CODE_TO & "'").Length
                If numdiff > 0 Then
                    Dim msg As String = "There are " & numdiff & " Styles stamped with the previously selected Location (" & LOCATION_CODE_TO & ")."
                    msg &= vbCrLf & vbCrLf
                    msg &= " Do you want to update these with the new location?"
                    If MessageBox.Show(msg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.Yes Then
                        For Each row As DataRow In frmDst.Tables("WHTMOVE2").Select("ISNULL(LOCATION_CODE_TO, '*') = '" & LOCATION_CODE_TO & "'")
                            row.Item("LOCATION_CODE_TO") = MyBase.Absx1.txtFor("LOCATION_CODE").Text
                        Next
                    End If
                End If
            End If

            LOCATION_CODE_TO = MyBase.Absx1.txtFor("LOCATION_CODE").Text
        End If
    End Sub

#End Region

End Class