Imports System.Drawing
Imports Infragistics.Win
Imports Infragistics.Win.UltraWinGrid

Public Class SOFPPRC1
    Private PriceListCode_Selected As String = ""
    Private ShowDetail As Boolean = False
    Private OriginalRowsByItem As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
#Region "ABS Standard Routines"
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("PA") Then
            MsgBox("You are not authorized to access Price List Approvals. Please contact ABS if you believe this is a mistake.", MsgBoxStyle.OkOnly, "Access Denied")
            Exit Sub
        End If

        With dst
            Create_TDA(.Tables.Add, "SOTPPRC1", "*")
            Create_TDA(.Tables.Add, "SOTPPRC2", "*")
            Create_TDA(.Tables.Add, "SOTPRIC1", "*")
            Create_TDA(.Tables.Add, "SOTPRIC2", "*")

        End With
        grdSOTPPRC1.DataSource = dst.Tables("SOTPPRC1")
        grdSOTPPRC2.DataSource = dst.Tables("SOTPPRC2")

        Create_Summary(grdSOTPPRC1, "PRICE_LIST_CODE", "Count")
        Create_Summary(grdSOTPPRC2, "ITEM_CODE", "Count")

        Load_Record()

        ShowDetail = False
        Mode_Settings(True)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Approve", "Reject"


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Approve"
                Approve_Selected()
                Load_Record()
                ShowDetail = False
                Mode_Settings(False)

            Case "Reject"
                Reject_Selected()
                Load_Record()
                ShowDetail = False
                Mode_Settings(False)

            Case "Done"
                If ShowDetail Then
                    ShowDetail = False
                    PriceListCode_Selected = ""
                    dst.Tables("SOTPPRC2").Rows.Clear()
                    Mode_Settings(False)
                End If

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        Dim canAction As Boolean = (tf AndAlso ShowDetail AndAlso PriceListCode_Selected <> "")

        With UltraExplorerBar1.Groups("Screen Control").Items
            .Item("Approve").Visible = ShowDetail
            .Item("Approve").Settings.Enabled = Not tf

            .Item("Reject").Visible = ShowDetail
            .Item("Reject").Settings.Enabled = Not tf

            .Item("Done").Visible = True
            .Item("Done").Settings.Enabled = Not tf
        End With

        grdSOTPPRC1.Visible = Not ShowDetail
        grdSOTPPRC2.Visible = ShowDetail

        Try
            spl.Panel1Collapsed = ShowDetail
            If Not ShowDetail Then spl.Panel1Collapsed = False
        Catch
        End Try
    End Sub
    Sub Clear_Record()

        PriceListCode_Selected = ""

        dst.Tables("SOTPPRC1").Rows.Clear()
        dst.Tables("SOTPPRC2").Rows.Clear()

    End Sub


    Sub Load_Record()

        ASCMAIN1.Progress("Loading pending approvals...")
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        dst.Tables("SOTPPRC1").Rows.Clear()
        dst.Tables("SOTPPRC2").Rows.Clear()

        Dim approver As String = (ASCMAIN1.USER_ID & "").Replace("'", "''")

        Dim sqlHdr As String =
        "SELECT * FROM SOTPPRC1 " &
        " WHERE STATUS = 'P' " &
        "   AND NVL(SEND_FOR_APPROVAL,'0') = '1' " &
        "   AND APPROVER_OPER = '" & approver & "'" &
        " ORDER BY SUBMIT_DATE DESC"

        Fill_Records("SOTPPRC1",,, sqlHdr)

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default

    End Sub


#End Region

#Region "Grid events"
    Private Sub grdSOTPPRC1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPPRC1.AfterRowActivate
        If grdSOTPPRC1.ActiveRow Is Nothing OrElse grdSOTPPRC1.ActiveRow.IsAddRow Then
            PriceListCode_Selected = ""
            dst.Tables("SOTPPRC2").Rows.Clear()
            Exit Sub
        End If

        PriceListCode_Selected = (grdSOTPPRC1.ActiveRow.Cells("PRICE_LIST_CODE").Text & "").Trim()
    End Sub
    Private Sub grdSOTPPRC1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTPPRC1.DoubleClickRow
        If e.Row Is Nothing OrElse e.Row.IsAddRow Then Exit Sub

        PriceListCode_Selected = (e.Row.Cells("PRICE_LIST_CODE").Text & "").Trim()
        If PriceListCode_Selected = "" Then Exit Sub

        Load_Detail(PriceListCode_Selected)

        ShowDetail = True
        Mode_Settings(True)
    End Sub

    Private Sub Load_Detail(priceListCode As String)

        If (priceListCode & "").Trim() = "" Then Exit Sub

        dst.Tables("SOTPPRC2").Rows.Clear()
        OriginalRowsByItem.Clear()

        If Not dst.Tables("SOTPPRC2").Columns.Contains("_ROW_STATE") Then
            dst.Tables("SOTPPRC2").Columns.Add("_ROW_STATE", GetType(String))
        End If

        Dim safePLC As String = priceListCode.Replace("'", "''")

        Dim sqlDtl As String =
        "SELECT * FROM SOTPPRC2 " &
        " WHERE PRICE_LIST_CODE = '" & safePLC & "'" &
        " ORDER BY ITEM_CODE"

        Fill_Records("SOTPPRC2",,, sqlDtl)

        Dim dtOrig As DataTable = ASCDATA1.GetDataTable(
        "SELECT * FROM SOTPRIC2 WHERE PRICE_LIST_CODE = '" & safePLC & "'"
    )

        If dtOrig IsNot Nothing Then
            For Each r As DataRow In dtOrig.Rows
                Dim item As String = (r("ITEM_CODE") & "").Trim()
                If item <> "" AndAlso Not OriginalRowsByItem.ContainsKey(item) Then
                    OriginalRowsByItem.Add(item, r)
                End If
            Next
        End If

        Dim pendingItems As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each pr As DataRow In dst.Tables("SOTPPRC2").Rows
            Dim item As String = (pr("ITEM_CODE") & "").Trim()
            If item = "" Then Continue For

            pendingItems.Add(item)

            If Not OriginalRowsByItem.ContainsKey(item) Then
                pr("_ROW_STATE") = "A"
            Else
                pr("_ROW_STATE") = "M"
            End If
        Next

        For Each kvp As KeyValuePair(Of String, DataRow) In OriginalRowsByItem
            Dim item As String = kvp.Key
            If pendingItems.Contains(item) Then Continue For

            Dim orig As DataRow = kvp.Value

            Dim drDel As DataRow = dst.Tables("SOTPPRC2").NewRow()
            drDel("PRICE_LIST_CODE") = priceListCode
            drDel("ITEM_CODE") = item

            If dst.Tables("SOTPPRC2").Columns.Contains("ITEM_PRICE") AndAlso orig.Table.Columns.Contains("ITEM_PRICE") Then
                drDel("ITEM_PRICE") = orig("ITEM_PRICE")
            End If
            If dst.Tables("SOTPPRC2").Columns.Contains("ITEM_SRP") AndAlso orig.Table.Columns.Contains("ITEM_SRP") Then
                drDel("ITEM_SRP") = orig("ITEM_SRP")
            End If

            drDel("_ROW_STATE") = "D"
            dst.Tables("SOTPPRC2").Rows.Add(drDel)
        Next

        grdSOTPPRC2.DataBind()
        grdSOTPPRC2.Refresh()

    End Sub

    Private Sub grdSOTPPRC2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTPPRC2.InitializeRow
        e.Row.Appearance.ResetBackColor()

        For Each c As UltraGridCell In e.Row.Cells
            c.Appearance.ResetBackColor()
        Next

        Dim rowState As String = e.Row.Cells("_ROW_STATE").Text & ""

        If rowState = "D" Then
            e.Row.Appearance.BackColor = System.Drawing.Color.MistyRose
            Return
        End If

        If rowState = "A" Then
            e.Row.Appearance.BackColor = System.Drawing.Color.Honeydew
            Return
        End If

        Dim item As String = e.Row.Cells("ITEM_CODE").Text & ""
        If item = "" Then Return
        If Not OriginalRowsByItem.ContainsKey(item) Then Return

        Dim orig As DataRow = OriginalRowsByItem(item)

        HighlightIfDifferent(e.Row, orig, "ITEM_PRICE", "ITEM_PRICE")
        HighlightIfDifferent(e.Row, orig, "ITEM_NEW_PRICE", "ITEM_PRICE")
        HighlightIfDifferent(e.Row, orig, "ITEM_SRP", "ITEM_SRP")
        HighlightIfDifferent(e.Row, orig, "ITEM_NEW_SRP", "ITEM_SRP")

    End Sub

    Private Sub HighlightIfDifferent(gr As UltraGridRow, orig As DataRow, pendingCol As String, origCol As String)

        If gr.Band.Columns.Exists(pendingCol) = False Then Exit Sub
        If orig.Table.Columns.Contains(origCol) = False Then Exit Sub

        Dim vPending As Decimal = Val(gr.Cells(pendingCol).Value & "")
        Dim vOrig As Decimal = Val(orig(origCol) & "")

        If vPending <> vOrig Then
            gr.Cells(pendingCol).Appearance.BackColor = System.Drawing.Color.LemonChiffon
        End If

    End Sub
    Private Sub grdSOTPPRC2_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSOTPPRC2.InitializeLayout
        With e.Layout.Bands(0)

            If .Columns.Exists("ITEM_PRICE") Then .Columns("ITEM_PRICE").Format = "$#,##0.00"
            If .Columns.Exists("ITEM_NEW_PRICE") Then .Columns("ITEM_NEW_PRICE").Format = "$#,##0.00"
            If .Columns.Exists("ITEM_SRP") Then .Columns("ITEM_SRP").Format = "$#,##0.00"
            If .Columns.Exists("ITEM_NEW_SRP") Then .Columns("ITEM_NEW_SRP").Format = "$#,##0.00"
            If .Columns.Exists("_ROW_STATE") Then .Columns("_ROW_STATE").Hidden = True


            For Each c As UltraGridColumn In .Columns
                c.CellActivation = Activation.NoEdit
            Next

            If .Columns.Exists("ITEM_PRICE") Then .Columns("ITEM_PRICE").Header.Caption = "NET Price"
            If .Columns.Exists("ITEM_NEW_PRICE") Then .Columns("ITEM_NEW_PRICE").Header.Caption = "New NET Price"
            If .Columns.Exists("ITEM_NEW_PRICE_DATE") Then .Columns("ITEM_NEW_PRICE_DATE").Header.Caption = "Date"
            If .Columns.Exists("ITEM_SRP") Then .Columns("ITEM_SRP").Header.Caption = "VENDOR SRP"
            If .Columns.Exists("ITEM_NEW_SRP") Then .Columns("ITEM_NEW_SRP").Header.Caption = "New VENDOR SRP"
            If .Columns.Exists("ITEM_NEW_SRP_DATE") Then .Columns("ITEM_NEW_SRP_DATE").Header.Caption = "Date New VENDOR SRP"

        End With
    End Sub
    Private Sub grdSOTPPRC1_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSOTPPRC1.InitializeLayout

        With e.Layout.Override
            .AllowUpdate = DefaultableBoolean.False
            .AllowAddNew = AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False

            .CellClickAction = CellClickAction.RowSelect
            .SelectTypeRow = SelectType.Single
            .SelectTypeCell = SelectType.None
            .RowSelectors = DefaultableBoolean.True
        End With

        With e.Layout.Bands(0)
            For Each c As UltraGridColumn In .Columns
                c.CellActivation = Activation.NoEdit
                c.TabStop = False
            Next
        End With

    End Sub


#End Region

#Region "Approve/Reject"

    Private Sub Approve_Selected()

        If (PriceListCode_Selected & "").Trim() = "" Then Exit Sub

        If MsgBox("Approve Price List " & PriceListCode_Selected & " ?", MsgBoxStyle.YesNo, "Approve") = MsgBoxResult.No Then
            Exit Sub
        End If

        BeginTrans()

        Try
            Dim safePLC As String = PriceListCode_Selected.Replace("'", "''")

            Dim dtHdr As DataTable =
            ASCDATA1.GetDataTable("SELECT * FROM SOTPPRC1 WHERE PRICE_LIST_CODE = '" & safePLC & "'")

            If dtHdr Is Nothing OrElse dtHdr.Rows.Count = 0 Then
                Rollback("Could not locate pending record.")
                Exit Sub
            End If

            Dim hdr As DataRow = dtHdr.Rows(0)

            Dim status As String = (hdr("STATUS") & "").Trim().ToUpper()
            If status <> "P" Then
                Rollback("This record is no longer Pending.")
                Exit Sub
            End If

            If hdr.Table.Columns.Contains("SEND_FOR_APPROVAL") Then
                Dim sfa As String = (hdr("SEND_FOR_APPROVAL") & "").Trim()
                If sfa <> "1" Then
                    Rollback("This record is not marked Send For Approval.")
                    Exit Sub
                End If
            End If

            Dim approverOper As String = (hdr("APPROVER_OPER") & "").Trim()
            If approverOper <> "" AndAlso approverOper.ToUpper() <> (ASCMAIN1.USER_ID & "").Trim().ToUpper() Then
                Rollback("You are not the approver for this price list.")
                Exit Sub
            End If

            If dst.Tables("SOTPPRC2").Rows.Count = 0 OrElse Not dst.Tables("SOTPPRC2").Columns.Contains("_ROW_STATE") Then
                Load_Detail(PriceListCode_Selected)
            End If
            Try
                If dst.Tables("SOTPRIC1").PrimaryKey Is Nothing OrElse dst.Tables("SOTPRIC1").PrimaryKey.Length = 0 Then
                    dst.Tables("SOTPRIC1").PrimaryKey = New DataColumn() {
                    dst.Tables("SOTPRIC1").Columns("PRICE_LIST_CODE")
                }
                End If
            Catch
            End Try

            Try
                If dst.Tables("SOTPRIC2").PrimaryKey Is Nothing OrElse dst.Tables("SOTPRIC2").PrimaryKey.Length = 0 Then
                    dst.Tables("SOTPRIC2").PrimaryKey = New DataColumn() {
                    dst.Tables("SOTPRIC2").Columns("PRICE_LIST_CODE"),
                    dst.Tables("SOTPRIC2").Columns("ITEM_CODE")
                }
                End If
            Catch
            End Try

            dst.Tables("SOTPRIC1").Rows.Clear()
            Fill_Records("SOTPRIC1",,, "SELECT * FROM SOTPRIC1 WHERE PRICE_LIST_CODE = '" & safePLC & "'")

            Dim plDesc As String = If(hdr.Table.Columns.Contains("PRICE_LIST_DESC"), (hdr("PRICE_LIST_DESC") & "").Trim(), "")
            Dim currCode As String = If(hdr.Table.Columns.Contains("CURR_CODE"), (hdr("CURR_CODE") & "").Trim(), "")

            If dst.Tables("SOTPRIC1").Rows.Count = 0 Then
                Dim rowLiveHdr As DataRow = dst.Tables("SOTPRIC1").NewRow()
                rowLiveHdr("PRICE_LIST_CODE") = PriceListCode_Selected

                If dst.Tables("SOTPRIC1").Columns.Contains("PRICE_LIST_DESC") Then rowLiveHdr("PRICE_LIST_DESC") = plDesc
                If dst.Tables("SOTPRIC1").Columns.Contains("CURR_CODE") Then rowLiveHdr("CURR_CODE") = currCode

                If dst.Tables("SOTPRIC1").Columns.Contains("INIT_OPER") Then rowLiveHdr("INIT_OPER") = ASCMAIN1.USER_ID
                If dst.Tables("SOTPRIC1").Columns.Contains("LAST_OPER") Then rowLiveHdr("LAST_OPER") = ASCMAIN1.USER_ID
                If dst.Tables("SOTPRIC1").Columns.Contains("INIT_DATE") Then rowLiveHdr("INIT_DATE") = DATETIME_STAMP
                If dst.Tables("SOTPRIC1").Columns.Contains("LAST_DATE") Then rowLiveHdr("LAST_DATE") = DATETIME_STAMP

                dst.Tables("SOTPRIC1").Rows.Add(rowLiveHdr)

                Write_Audit_Trail(rowLiveHdr, "N")
            Else
                Dim rowLiveHdr As DataRow = dst.Tables("SOTPRIC1").Rows(0)

                rowLiveHdr.BeginEdit()
                If plDesc <> "" AndAlso dst.Tables("SOTPRIC1").Columns.Contains("PRICE_LIST_DESC") Then rowLiveHdr("PRICE_LIST_DESC") = plDesc
                If currCode <> "" AndAlso dst.Tables("SOTPRIC1").Columns.Contains("CURR_CODE") Then rowLiveHdr("CURR_CODE") = currCode

                If dst.Tables("SOTPRIC1").Columns.Contains("LAST_OPER") Then rowLiveHdr("LAST_OPER") = ASCMAIN1.USER_ID
                If dst.Tables("SOTPRIC1").Columns.Contains("LAST_DATE") Then rowLiveHdr("LAST_DATE") = DATETIME_STAMP
                rowLiveHdr.EndEdit()

                Write_Audit_Trail(rowLiveHdr, "E")
            End If

            Update_Record_TDA("SOTPRIC1")

            dst.Tables("SOTPRIC2").Rows.Clear()
            Fill_Records("SOTPRIC2",,, "SELECT * FROM SOTPRIC2 WHERE PRICE_LIST_CODE = '" & safePLC & "'")

            Dim liveByItem As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
            For Each r As DataRow In dst.Tables("SOTPRIC2").Rows
                Dim it As String = (r("ITEM_CODE") & "").Trim()
                If it <> "" AndAlso Not liveByItem.ContainsKey(it) Then
                    liveByItem.Add(it, r)
                End If
            Next

            Dim dtPend As DataTable = dst.Tables("SOTPPRC2")

            For Each pr As DataRow In dtPend.Rows

                Dim item As String = (pr("ITEM_CODE") & "").Trim()
                If item = "" Then Continue For

                Dim rowState As String = ""
                If pr.Table.Columns.Contains("_ROW_STATE") AndAlso Not IsDBNull(pr("_ROW_STATE")) Then
                    rowState = (pr("_ROW_STATE") & "").Trim().ToUpper()
                End If

                Dim exists As Boolean = liveByItem.ContainsKey(item)

                If rowState = "D" Then
                    If exists Then
                        Dim rowLive As DataRow = liveByItem(item)
                        rowLive.Delete()
                        Write_Audit_Trail(rowLive, "E")
                    End If
                    Continue For
                End If

                Dim item_price As Decimal = Val(pr("ITEM_PRICE") & "")
                Dim item_new_price As Decimal = Val(pr("ITEM_NEW_PRICE") & "")
                Dim item_srp As Decimal = Val(pr("ITEM_SRP") & "")
                Dim item_new_srp As Decimal = Val(pr("ITEM_NEW_SRP") & "")

                Dim newPriceDateObj As Object = DBNull.Value
                If pr.Table.Columns.Contains("ITEM_NEW_PRICE_DATE") AndAlso (pr("ITEM_NEW_PRICE_DATE") & "") <> "" Then
                    newPriceDateObj = pr("ITEM_NEW_PRICE_DATE")
                End If

                Dim newSrpDateObj As Object = DBNull.Value
                If pr.Table.Columns.Contains("ITEM_NEW_SRP_DATE") AndAlso (pr("ITEM_NEW_SRP_DATE") & "") <> "" Then
                    newSrpDateObj = pr("ITEM_NEW_SRP_DATE")
                End If

                If Not exists Then
                    Dim rowLive As DataRow = dst.Tables("SOTPRIC2").NewRow()
                    rowLive("PRICE_LIST_CODE") = PriceListCode_Selected
                    rowLive("ITEM_CODE") = item

                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_PRICE") Then rowLive("ITEM_PRICE") = item_price
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_PRICE") Then rowLive("ITEM_NEW_PRICE") = item_new_price
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_PRICE_DATE") Then rowLive("ITEM_NEW_PRICE_DATE") = newPriceDateObj
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_SRP") Then rowLive("ITEM_SRP") = item_srp
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_SRP") Then rowLive("ITEM_NEW_SRP") = item_new_srp
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_SRP_DATE") Then rowLive("ITEM_NEW_SRP_DATE") = newSrpDateObj

                    If dst.Tables("SOTPRIC2").Columns.Contains("INIT_OPER") Then rowLive("INIT_OPER") = ASCMAIN1.USER_ID
                    If dst.Tables("SOTPRIC2").Columns.Contains("LAST_OPER") Then rowLive("LAST_OPER") = ASCMAIN1.USER_ID
                    If dst.Tables("SOTPRIC2").Columns.Contains("INIT_DATE") Then rowLive("INIT_DATE") = DATETIME_STAMP
                    If dst.Tables("SOTPRIC2").Columns.Contains("LAST_DATE") Then rowLive("LAST_DATE") = DATETIME_STAMP

                    dst.Tables("SOTPRIC2").Rows.Add(rowLive)
                    liveByItem(item) = rowLive

                    Write_Audit_Trail(rowLive, "N")
                Else
                    Dim rowLive As DataRow = liveByItem(item)

                    rowLive.BeginEdit()
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_PRICE") Then rowLive("ITEM_PRICE") = item_price
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_PRICE") Then rowLive("ITEM_NEW_PRICE") = item_new_price
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_PRICE_DATE") Then rowLive("ITEM_NEW_PRICE_DATE") = newPriceDateObj
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_SRP") Then rowLive("ITEM_SRP") = item_srp
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_SRP") Then rowLive("ITEM_NEW_SRP") = item_new_srp
                    If dst.Tables("SOTPRIC2").Columns.Contains("ITEM_NEW_SRP_DATE") Then rowLive("ITEM_NEW_SRP_DATE") = newSrpDateObj

                    If dst.Tables("SOTPRIC2").Columns.Contains("LAST_OPER") Then rowLive("LAST_OPER") = ASCMAIN1.USER_ID
                    If dst.Tables("SOTPRIC2").Columns.Contains("LAST_DATE") Then rowLive("LAST_DATE") = DATETIME_STAMP
                    rowLive.EndEdit()

                    Write_Audit_Trail(rowLive, "E")
                End If

            Next

            Update_Record_TDA("SOTPRIC2", "PRICE_LIST_CODE = '" & safePLC & "'")

            ASCDATA1.ExecuteSQL("DELETE FROM SOTPPRC2 WHERE PRICE_LIST_CODE = '" & safePLC & "'")
            ASCDATA1.ExecuteSQL("DELETE FROM SOTPPRC1 WHERE PRICE_LIST_CODE = '" & safePLC & "'")

            CommitTrans("Approved")

        Catch ex As Exception
            Rollback("Approve failed: " & ex.Message)
        End Try

    End Sub

    Private Sub Reject_Selected()

        If (PriceListCode_Selected & "").Trim() = "" Then Exit Sub

        If MsgBox("Reject Price List " & PriceListCode_Selected & " ? (This will remove it from pending approvals)", MsgBoxStyle.YesNo, "Reject") = MsgBoxResult.No Then
            Exit Sub
        End If

        BeginTrans()

        Try
            Dim safePLC As String = PriceListCode_Selected.Replace("'", "''")

            Dim dtHdr As DataTable =
                ASCDATA1.GetDataTable("SELECT * FROM SOTPPRC1 WHERE PRICE_LIST_CODE = '" & safePLC & "'")

            If dtHdr Is Nothing OrElse dtHdr.Rows.Count = 0 Then
                Rollback("Could not locate pending record.")
                Exit Sub
            End If

            Dim hdr As DataRow = dtHdr.Rows(0)

            Dim status As String = (hdr("STATUS") & "").Trim().ToUpper()
            If status <> "P" Then
                Rollback("This record is no longer Pending.")
                Exit Sub
            End If

            Dim approverOper As String = (hdr("APPROVER_OPER") & "").Trim()
            If approverOper <> "" AndAlso approverOper.ToUpper() <> (ASCMAIN1.USER_ID & "").Trim().ToUpper() Then
                Rollback("You are not the approver for this price list.")
                Exit Sub
            End If

            '(reject = discard changes)
            ASCDATA1.ExecuteSQL("DELETE FROM SOTPPRC2 WHERE PRICE_LIST_CODE = '" & safePLC & "'")
            ASCDATA1.ExecuteSQL("DELETE FROM SOTPPRC1 WHERE PRICE_LIST_CODE = '" & safePLC & "'")

            CommitTrans("Rejected (removed from pending approvals)")

        Catch ex As Exception
            Rollback("Reject failed: " & ex.Message)
        End Try

    End Sub



#End Region

End Class
