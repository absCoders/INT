Imports System.Data
Imports System.Windows.Forms
Public Class SOTPPRC1

    Dim sqlSOTPRIC2 As String = ""
    Dim isCurrRetailPriceList As Boolean = False
    Dim decimalsOK As New Infragistics.Win.Appearance
    Dim decimalsConcealed As New Infragistics.Win.Appearance
    Private IsSubmittingForApproval As Boolean = False
    Private CurrentPriceListCodeSubmitted As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Get_PARM("SOTPARM1")

        decimalsOK.ForeColor = Drawing.Color.Empty
        decimalsConcealed.ForeColor = Drawing.Color.Red

        With dst
            Create_TDA(.Tables.Add, "TATALRT1", "*")

            ASCMAIN1.sql = "Select SOTPRIC2.*, ICTITEM1.ITEM_DESC" _
                & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_LIST_PRICE, ICTITEM1.ITEM_COST_STD " _
            & " from SOTPRIC2,ICTITEM1" _
            & " where SOTPRIC2.PRICE_LIST_CODE = :PARM1 " _
            & "   and ICTITEM1.ITEM_CODE = SOTPRIC2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTPRIC2", "**", 0, True, "V", 2)
            With .Tables("SOTPRIC2").Columns
                .Add("ITEM_GP_AMT", GetType(System.Decimal), "ISNULL(ITEM_PRICE,0)-ISNULL(ITEM_COST_STD,0)")
                .Add("ITEM_GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ITEM_PRICE,0)=0,0,100/100*(ISNULL(ITEM_PRICE,0)-ISNULL(ITEM_COST_STD,0))/ISNULL(ITEM_PRICE,0))")

                Dim DISC As String = "IIF(ISNULL(ITEM_RETAIL_PRICE,0)=0,0,100/100*(ISNULL(ITEM_RETAIL_PRICE,0)-ISNULL(ITEM_PRICE,0))/ISNULL(ITEM_RETAIL_PRICE,0))"
                .Add("DISC_PCT_RETAIL", GetType(System.Decimal), DISC)
                '.Add("DISC_PCT_LIST", GetType(System.Decimal), Replace(DISC, "ITEM_RETAIL_PRICE", "ITEM_LIST_PRICE"))
                .Add("DISC_PCT_LIST", GetType(System.Decimal), Replace(DISC, "ITEM_RETAIL_PRICE", "ITEM_SRP"))

                .Add("ITEM_PRICE_1", GetType(System.Decimal))
                .Add("ITEM_PRICE_1_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_2", GetType(System.Decimal))
                .Add("ITEM_PRICE_2_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_3", GetType(System.Decimal))
                .Add("ITEM_PRICE_3_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_4", GetType(System.Decimal))
                .Add("ITEM_PRICE_4_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_5", GetType(System.Decimal))
                .Add("ITEM_PRICE_5_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_6", GetType(System.Decimal))
                .Add("ITEM_PRICE_6_DATE", GetType(System.DateTime))
            End With
            ASCMAIN1.sql = "Select SOTPPRC2.*, ICTITEM1.ITEM_DESC" &
               ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_LIST_PRICE, ICTITEM1.ITEM_COST_STD " &
               " from SOTPPRC2, ICTITEM1" &
               " where SOTPPRC2.PRICE_LIST_CODE = :PARM1" &
               "   and ICTITEM1.ITEM_CODE = SOTPPRC2.ITEM_CODE"

            Create_TDA(dst.Tables.Add, "SOTPPRC2", "**", 0, True, "V", 2)
            With dst.Tables("SOTPPRC2").Columns
                .Add("ITEM_GP_AMT", GetType(System.Decimal), "ISNULL(ITEM_PRICE,0)-ISNULL(ITEM_COST_STD,0)")
                .Add("ITEM_GP_PCT", GetType(System.Decimal), "IIF(ISNULL(ITEM_PRICE,0)=0,0,100/100*(ISNULL(ITEM_PRICE,0)-ISNULL(ITEM_COST_STD,0))/ISNULL(ITEM_PRICE,0))")

                Dim DISC As String = "IIF(ISNULL(ITEM_RETAIL_PRICE,0)=0,0,100/100*(ISNULL(ITEM_RETAIL_PRICE,0)-ISNULL(ITEM_PRICE,0))/ISNULL(ITEM_RETAIL_PRICE,0))"
                .Add("DISC_PCT_RETAIL", GetType(System.Decimal), DISC)
                .Add("DISC_PCT_LIST", GetType(System.Decimal), Replace(DISC, "ITEM_RETAIL_PRICE", "ITEM_SRP"))
                .Add("ITEM_PRICE_1", GetType(System.Decimal))
                .Add("ITEM_PRICE_1_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_2", GetType(System.Decimal))
                .Add("ITEM_PRICE_2_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_3", GetType(System.Decimal))
                .Add("ITEM_PRICE_3_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_4", GetType(System.Decimal))
                .Add("ITEM_PRICE_4_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_5", GetType(System.Decimal))
                .Add("ITEM_PRICE_5_DATE", GetType(System.DateTime))
                .Add("ITEM_PRICE_6", GetType(System.Decimal))
                .Add("ITEM_PRICE_6_DATE", GetType(System.DateTime))
            End With


        End With

        grdSOTPPRC2.DataSource = dst.Tables("SOTPPRC2")

        grdSOTPPRC2.DisplayLayout.UseFixedHeaders = True
        With grdSOTPPRC2.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTPPRC2.DisplayLayout.Bands(0).Columns
            If New String() {"ITEM_CODE", "ITEM_PRICE", "ITEM_NEW_PRICE", "ITEM_NEW_PRICE_DATE", "ITEM_SRP", "ITEM_NEW_SRP", "ITEM_NEW_SRP_DATE"}.Contains(gcol.Key) Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellAppearance.BackColor = Drawing.Color.LightGray
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        Create_Summary(grdSOTPPRC2, "ITEM_CODE", "Count")

        With grdSOTPPRC2.DisplayLayout.Bands(0)
            .Columns("ITEM_SRP").Hidden = (ASCMAIN1.CLIENT = "AHA")
            .Columns("ITEM_NEW_SRP").Hidden = (ASCMAIN1.CLIENT = "AHA")
            .Columns("ITEM_NEW_SRP_DATE").Hidden = (ASCMAIN1.CLIENT = "AHA")
            .Columns("ITEM_COST_STD").Hidden = (ASCMAIN1.CLIENT = "AHA")
            .Columns("ITEM_GP_AMT").Hidden = (ASCMAIN1.CLIENT = "AHA")
            .Columns("ITEM_GP_PCT").Hidden = (ASCMAIN1.CLIENT = "AHA")


            .Columns("ITEM_LIST_PRICE").Hidden = Not (ASCMAIN1.CLIENT = "AHA")
            ' .Columns("DISC_PCT_LIST").Hidden = NOT (ASCMAIN1.CLIENT = "AHA")

            If Not (ASCMAIN1.CLIENT = "AHA") Then
                .ColHeaderLines = 2

                .Columns("ITEM_PRICE").Header.Caption = "NET Price"
                .Columns("ITEM_NEW_PRICE").Header.Caption = "New NET Price"
                .Columns("ITEM_SRP").Header.Caption = "VENDOR SRP"
                .Columns("ITEM_NEW_SRP").Header.Caption = "New VENDOR SRP"
                .Columns("ITEM_NEW_SRP_DATE").Header.Caption = "Date New VENDOR SRP"
                .Columns("DISC_PCT_RETAIL").Header.Caption = "Disc% to MSRP"
                .Columns("DISC_PCT_LIST").Header.Caption = "Disc% to Vendor SRP"
            End If

            For I As Integer = 1 To 6
                Dim II As String = Format(I, "0")
                .Columns("ITEM_PRICE_" & II).CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                .Columns("ITEM_PRICE_" & II & "_DATE").CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                .Columns("ITEM_PRICE_" & II).Hidden = (ASCMAIN1.CLIENT = "AHA")
                .Columns("ITEM_PRICE_" & II & "_DATE").Hidden = (ASCMAIN1.CLIENT = "AHA")
                .Columns("ITEM_PRICE_" & II).Header.Caption = "Price " & II & " ago"
                .Columns("ITEM_PRICE_" & II & "_DATE").Header.Caption = "Date " & II & " ago"
            Next
        End With

        Dim tbl As DataTable = ASCDATA1.GetDataTable("Select * from TATCURR1")
        If tbl.Rows.Count < 2 Or (ASCMAIN1.CLIENT = "INT") Then
            lblCURR_CODE.Visible = False
            txtCURR_CODE.Visible = False
        Else
            lblCURR_CODE.Visible = True
            txtCURR_CODE.Visible = True
        End If

        MakeTransparent(chkShow4Decimals)
        Load_Approver_List()
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPPRC2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")

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

            Case "grdSOTPPRC2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Add Codes"
                If grd.Name = "grdSOTPPRC2" Then
                    Add_Codes(grdSOTPPRC2, "ICTITEM1", "ITEM_CODE", "Items")
                End If
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
                Dim PRICE_LIST_CODE As String = (Absx1.txtFor("PRICE_LIST_CODE").Text & "").Trim()
                If PRICE_LIST_CODE <> "" Then

                    Dim pending As DataRow = Get_Open_Pending_Info(PRICE_LIST_CODE)
                    If pending IsNot Nothing Then

                        Dim submitOper As String = (pending("SUBMIT_OPER") & "").Trim()

                        EMsg &= vbCrLf &
                            "A pending price list change already exists for this Price List Code." & vbCrLf &
                            "Submitted By: " & submitOper & vbCrLf & vbCrLf &
                            "Please use EDIT instead."
                    End If
                End If

            Case "Edit"

            Case "Update"
                If Absx1.txtFor("PRICE_LIST_DESC").Text = "" Then
                    EMsg &= vbCr & "Description is Mandatory"
                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    If Absx1.txtFor("CURR_CODE").Text <> "USD" Then
                        EMsg &= vbCr & "Non-USD prices are not supported"
                    End If
                Else
                    Dim CURR_CODE As String = Absx1.txtFor("CURR_CODE").Text
                    If CURR_CODE = "" Then
                        EMsg &= vbCr & "Currency Code is Mandatory"
                    Else
                        Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", CURR_CODE)
                        If rowTATCURR1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Currency Code (" & CURR_CODE & ")"
                        End If
                    End If

                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    ' LAUREN SAID NOT TO PUT THIS IN. 11/10
                Else
                    For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPPRC2").Select("")
                        Dim ITEM_CODE As String = rowSOTPRIC2.Item("ITEM_CODE") & ""
                        Dim ITEM_NEW_PRICE As Decimal = Val(rowSOTPRIC2.Item("ITEM_NEW_PRICE") & "")
                        If Val(ITEM_NEW_PRICE) < 0 Then
                            EMsg &= vbCr & ITEM_CODE & ":" & "New Item Price must be > 0"
                        End If
                        If rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") & "" <> "" Then
                            If Val(ITEM_NEW_PRICE) <= 0 Then
                                EMsg &= vbCr & ITEM_CODE & ":" & "New Item Price Date provided without a New Price"
                            End If
                            Dim DTE As Date = rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE")
                            If Format(DTE, "dd") <> "01" Or Format(DTE, "yyyyMM") <= ASCMAIN1.CYM Then
                                EMsg &= vbCr & ITEM_CODE & ":" & "New Item Price Date must be the 1st of a Future Month"
                            End If
                        Else
                            If Val(ITEM_NEW_PRICE) > 0 Then
                                EMsg &= vbCr & ITEM_CODE & ":" & "New Item Price Date is Mandatory if specifying a New Price"
                            End If
                        End If
                    Next
                End If

                Dim pending As DataRow = Get_Open_Pending_Info(Absx1.txtFor("PRICE_LIST_CODE").Text)
                If pending IsNot Nothing Then
                    Dim submitOper As String = (pending("SUBMIT_OPER") & "").Trim()
                    If submitOper <> "" AndAlso submitOper.ToUpper <> (ASCMAIN1.USER_ID & "").ToUpper Then
                        EMsg &= vbCrLf & "This price list is already pending approval (submitted by " & submitOper & ")."
                    End If
                End If

                If chkSendForApproval.Checked Then
                    If cmbApprover.SelectedIndex < 0 Then
                        EMsg &= vbCrLf & "Approver is mandatory when Send for Approval is checked."
                    End If

                    If (cmbApprover.SelectedValue & "").ToUpper = (ASCMAIN1.USER_ID & "").ToUpper Then
                        EMsg &= vbCrLf & "Approver cannot be yourself."
                    End If

                    If dst.Tables("SOTPPRC2").Rows.Count = 0 Then
                        EMsg &= vbCrLf & "You must have at least one item in the price list before sending for approval."
                    End If
                End If

                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each r As DataRow In dst.Tables("SOTPPRC2").Rows
                    If r.RowState = DataRowState.Deleted Then Continue For
                    Dim it As String = (r("ITEM_CODE") & "").Trim()
                    If it = "" Then
                        EMsg &= vbCrLf & "Item Code cannot be blank."
                        Continue For
                    End If
                    If Not seen.Add(it) Then
                        EMsg &= vbCrLf & "Duplicate Item Code in list: " & it
                    End If

                    If LookUp("ICTITEM1", it) Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Item Code: " & it
                    End If
                Next




        End Select
    End Sub

    Private Function Save_PendingPriceList(ByVal sendForApproval As Boolean) As Boolean

        If IsSubmittingForApproval Then Return False

        Dim PRICE_LIST_CODE As String = (Absx1.txtFor("PRICE_LIST_CODE").Text & "").Trim()
        Dim PRICE_LIST_DESC As String = (Absx1.txtFor("PRICE_LIST_DESC").Text & "").Trim()
        Dim CURR_CODE As String = (Absx1.txtFor("CURR_CODE").Text & "").Trim()

        If PRICE_LIST_CODE = "" Then
            MsgBox("Price List Code is required before saving/updating.",
               MsgBoxStyle.OkOnly, "Missing Price List Code")
            Return False
        End If

        Dim approverId As String = (cmbApprover.SelectedValue & "").Trim()
        Dim approverName As String = (cmbApprover.Text & "").Trim()

        'Only enforce approver if sending
        If sendForApproval AndAlso cmbApprover.SelectedIndex < 0 Then
            MsgBox("Please select an approver before sending for approval.",
               MsgBoxStyle.OkOnly, "Approver Required")
            cmbApprover.Focus()
            Return False
        End If

        Dim safePLC As String = SqlSafe(PRICE_LIST_CODE)
        Dim safeDesc As String = SqlSafe(PRICE_LIST_DESC)
        Dim safeCurr As String = SqlSafe(CURR_CODE)
        Dim safeUser As String = SqlSafe(ASCMAIN1.USER_ID)
        Dim safeAppr As String = SqlSafe(approverId)
        Dim dtStamp As String = SqlDateTimeStamp()

        IsSubmittingForApproval = True

        Try
            Dim pendingRow As DataRow = Get_Open_Pending_Info(PRICE_LIST_CODE)
            If pendingRow IsNot Nothing Then
                Dim existingSubmitOper As String = (pendingRow("SUBMIT_OPER") & "").Trim()
                If existingSubmitOper <> "" AndAlso existingSubmitOper.ToUpper() <> (ASCMAIN1.USER_ID & "").ToUpper() Then
                    MsgBox("This price list already has a pending submission by " & existingSubmitOper & ".",
                       MsgBoxStyle.OkOnly, "Already Pending")
                    Return False
                End If
            End If

            Dim prevSendFlag As String = "0"
            Dim prevStatus As String = ""
            Dim prevApprover As String = ""

            Dim dtPrev As DataTable = ASCDATA1.GetDataTable(
            "select STATUS, nvl(SEND_FOR_APPROVAL,'0') as SEND_FOR_APPROVAL, nvl(APPROVER_OPER,'') as APPROVER_OPER " &
            "  from SOTPPRC1 where PRICE_LIST_CODE = '" & safePLC & "'"
        )
            If dtPrev IsNot Nothing AndAlso dtPrev.Rows.Count > 0 Then
                prevStatus = (dtPrev.Rows(0)("STATUS") & "").Trim()
                prevSendFlag = (dtPrev.Rows(0)("SEND_FOR_APPROVAL") & "").Trim()
                prevApprover = (dtPrev.Rows(0)("APPROVER_OPER") & "").Trim()
            End If

            Dim needToEmail As Boolean =
            sendForApproval AndAlso prevApprover <> approverId

            Dim statusToSet As String = If(sendForApproval, "P", "D")
            Dim sendFlagToSet As String = If(sendForApproval, "1", "0")

            Dim exists As Boolean = (dtPrev IsNot Nothing AndAlso dtPrev.Rows.Count > 0)

            If Not exists Then
                Dim sqlIns As String =
                "insert into SOTPPRC1 " &
                "(PRICE_LIST_CODE, PRICE_LIST_DESC, CURR_CODE, " &
                " INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE, " &
                " STATUS, SEND_FOR_APPROVAL, " &
                " SUBMIT_OPER, SUBMIT_DATE, APPROVER_OPER, APPROVED_DATE) " &
                "values (" &
                "'" & safePLC & "'," &
                "'" & safeDesc & "'," &
                "'" & safeCurr & "'," &
                "'" & safeUser & "'," &
                "'" & safeUser & "'," &
                "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                "'" & statusToSet & "'," &
                "'" & sendFlagToSet & "'," &
                If(sendForApproval, "'" & safeUser & "'", "NULL") & "," &
                If(sendForApproval, "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')", "NULL") & "," &
                If(sendForApproval, "'" & safeAppr & "'", "NULL") & "," &
                "NULL)"

                ASCDATA1.ExecuteSQL(sqlIns)

            Else
                Dim sqlUpd As String =
                "update SOTPPRC1 set " &
                " PRICE_LIST_DESC = '" & safeDesc & "'," &
                " CURR_CODE = '" & safeCurr & "'," &
                " LAST_OPER = '" & safeUser & "'," &
                " LAST_DATE = TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                " STATUS = '" & statusToSet & "'," &
                " SEND_FOR_APPROVAL = '" & sendFlagToSet & "'," &
                " SUBMIT_OPER = " & If(sendForApproval, "'" & safeUser & "'", "NULL") & "," &
                " SUBMIT_DATE = " & If(sendForApproval, "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')", "NULL") & "," &
                " APPROVER_OPER = " & If(sendForApproval, "'" & safeAppr & "'", "NULL") & "," &
                " APPROVED_DATE = NULL " &
                "where PRICE_LIST_CODE = '" & safePLC & "'"

                ASCDATA1.ExecuteSQL(sqlUpd)
            End If

            ASCDATA1.ExecuteSQL("delete from SOTPPRC2 where PRICE_LIST_CODE = '" & safePLC & "'")

            Dim dtDetail As DataTable = dst.Tables("SOTPPRC2")
            If dtDetail IsNot Nothing Then
                For Each row As DataRow In dtDetail.Rows
                    If row.RowState = DataRowState.Deleted Then Continue For

                    Dim itemCode As String = (row("ITEM_CODE") & "").Trim()
                    If itemCode = "" Then Continue For

                    Dim safeItem As String = SqlSafe(itemCode)

                    Dim ITEM_PRICE As Decimal = Val(row("ITEM_PRICE") & "")
                    Dim ITEM_NEW_PRICE As Decimal = Val(row("ITEM_NEW_PRICE") & "")
                    Dim ITEM_SRP As Decimal = Val(row("ITEM_SRP") & "")
                    Dim ITEM_NEW_SRP As Decimal = Val(row("ITEM_NEW_SRP") & "")

                    Dim newPriceDateSql As String = "NULL"
                    If Not IsDBNull(row("ITEM_NEW_PRICE_DATE")) AndAlso (row("ITEM_NEW_PRICE_DATE") & "") <> "" Then
                        newPriceDateSql = SqlDateOnly(CDate(row("ITEM_NEW_PRICE_DATE")))
                    End If

                    Dim newSrpDateSql As String = "NULL"
                    If Not IsDBNull(row("ITEM_NEW_SRP_DATE")) AndAlso (row("ITEM_NEW_SRP_DATE") & "") <> "" Then
                        newSrpDateSql = SqlDateOnly(CDate(row("ITEM_NEW_SRP_DATE")))
                    End If

                    Dim sqlDtl As String =
                    "insert into SOTPPRC2 " &
                    "(PRICE_LIST_CODE, ITEM_CODE, " &
                    " ITEM_PRICE, ITEM_NEW_PRICE, ITEM_NEW_PRICE_DATE, " &
                    " ITEM_SRP, ITEM_NEW_SRP, ITEM_NEW_SRP_DATE, " &
                    " INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE) " &
                    "values (" &
                    "'" & safePLC & "'," &
                    "'" & safeItem & "'," &
                    ITEM_PRICE.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    ITEM_NEW_PRICE.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    newPriceDateSql & "," &
                    ITEM_SRP.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    ITEM_NEW_SRP.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    newSrpDateSql & "," &
                    "'" & safeUser & "'," &
                    "'" & safeUser & "'," &
                    "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                    "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS'))"

                    ASCDATA1.ExecuteSQL(sqlDtl)
                Next
            End If

            If needToEmail Then
                Dim approverEmail As String = Get_Approver_Email(approverId, approverName)
                If approverEmail = "" Then Return False

                Send_Approval_Email(approverEmail, approverName, PRICE_LIST_CODE, PRICE_LIST_DESC)
            End If

            CurrentPriceListCodeSubmitted = PRICE_LIST_CODE

            If sendForApproval Then
                MsgBox("Price list " & PRICE_LIST_CODE & " sent for approval." & vbCrLf &
                   "Approver: " & approverName,
                   MsgBoxStyle.Information, "Price List Approval")
            Else
                MsgBox("Draft saved for price list " & PRICE_LIST_CODE & ".",
                   MsgBoxStyle.Information, "Price List Draft Saved")
            End If

            Return True

        Catch ex As Exception
            MsgBox("Error while saving:" & vbCrLf & ex.Message,
               MsgBoxStyle.Critical, "Save Pending Price List")
            Return False

        Finally
            IsSubmittingForApproval = False
        End Try

    End Function
    Private Sub Send_Approval_Email(ByVal approverEmail As String,
                                ByVal approverName As String,
                                ByVal priceListCode As String,
                                ByVal priceListDesc As String)

        If (approverEmail & "").Trim() = "" Then
            MsgBox("Approver email is blank. Cannot send approval email.",
               MsgBoxStyle.OkOnly, "Missing Approver Email")
            Exit Sub
        End If

        'If ASCMAIN1.Running_in_VS Then Exit Sub

        Dim submittedBy As String = (ASCMAIN1.USER_ID & "").Trim()
        Dim subj As String = "Price List Approval Needed: " & priceListCode & " - " & priceListDesc

        Dim totalItems As Integer = 0
        Dim changedItems As Integer = 0
        Dim earliestEffDate As Date = Date.MaxValue
        Dim hasEffDate As Boolean = False

        Dim dtDetail As DataTable = dst.Tables("SOTPPRC2")
        If dtDetail IsNot Nothing Then
            totalItems = dtDetail.Select("ITEM_CODE is not null").Length

            For Each r As DataRow In dtDetail.Rows
                If r.RowState = DataRowState.Deleted Then Continue For

                Dim newNet As Decimal = Val(r("ITEM_NEW_PRICE") & "")
                Dim newSrp As Decimal = 0D
                If dtDetail.Columns.Contains("ITEM_NEW_SRP") Then newSrp = Val(r("ITEM_NEW_SRP") & "")

                If newNet > 0 OrElse newSrp > 0 Then changedItems += 1

                If dtDetail.Columns.Contains("ITEM_NEW_PRICE_DATE") AndAlso Not IsDBNull(r("ITEM_NEW_PRICE_DATE")) Then
                    Dim d As Date = CDate(r("ITEM_NEW_PRICE_DATE"))
                    If d < earliestEffDate Then earliestEffDate = d
                    hasEffDate = True
                End If

                If dtDetail.Columns.Contains("ITEM_NEW_SRP_DATE") AndAlso Not IsDBNull(r("ITEM_NEW_SRP_DATE")) Then
                    Dim d As Date = CDate(r("ITEM_NEW_SRP_DATE"))
                    If d < earliestEffDate Then earliestEffDate = d
                    hasEffDate = True
                End If
            Next
        End If

        Dim msg As String =
        "A price list is ready for your review/approval." & vbCrLf & vbCrLf &
        "Price List Code: " & priceListCode & vbCrLf &
        "Description: " & priceListDesc & vbCrLf &
        "Submitted By: " & submittedBy & vbCrLf &
        "Total Items: " & totalItems & vbCrLf &
        "Please open Price List Approvals to Approve/Reject."

        Dim toDict As New Dictionary(Of String, String)
        toDict.Add(approverEmail, approverName)

        ASCMAIN1.TACMAIN1.Send_email(
        ASCMAIN1.ActiveForm,
        toDict,
        Nothing,
        subj,
        "PC_PRCAPPR",
        True,
        False,
        priceListCode,
        priceListCode,
        "Price List Code",
        msg
    )

    End Sub



    Overrides Sub Proceed_Update_Special_Post()
        Dim sendIt As Boolean = chkSendForApproval.Checked

        If sendIt Then
            If Not Save_PendingPriceList(True) Then
                Throw New ApplicationException("Update cancelled - submit for approval did not complete.")
            End If
        Else
            If Not Save_PendingPriceList(False) Then
                Throw New ApplicationException("Update cancelled - draft save did not complete.")
            End If
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        Dim PRICE_LIST_CODE As String = (Absx1.txtFor("PRICE_LIST_CODE").Text & "").Trim()

        If EntryMode = "New" Then

            If PRICE_LIST_CODE = "" Then
                StartBlankNew("")
            Else
                Dim pending As DataRow = Get_Open_Pending_Info(PRICE_LIST_CODE)
                If pending IsNot Nothing Then
                    MsgBox("A pending price list already exists for " & PRICE_LIST_CODE & "." & vbCrLf &
                       "Please use EDIT instead.",
                       MsgBoxStyle.OkOnly, "Pending Price List Exists")
                    StartBlankNew(PRICE_LIST_CODE)
                    Exit Sub
                End If

                If LiveHeaderExists(PRICE_LIST_CODE) Then
                    PreloadFromLive(PRICE_LIST_CODE)
                Else
                    StartBlankNew(PRICE_LIST_CODE)
                End If
            End If

        Else
            EnforceConstraints(False)
            dst.Tables("SOTPPRC2").Rows.Clear()
            Fill_Records("SOTPPRC2", New String() {PRICE_LIST_CODE})
            EnforceConstraints(True)
        End If

        Sort_grdColumns(grdSOTPPRC2, "ITEM_CODE")
        grdSOTPPRC2.Text = "Price List Details for " & PRICE_LIST_CODE

        Dim need4places As Boolean = False
        For Each row As DataRow In dst.Tables("SOTPPRC2").Select
            For Each C As String In New String() {"ITEM_PRICE", "ITEM_NEW_PRICE"}
                Dim V As Decimal = Val(row.Item(C) & "")
                If Val(Format(V, "#.00")) = V Then
                Else
                    need4places = True
                    Exit For
                End If
            Next
            If need4places Then
                Exit For
            End If
        Next

        chkShow4Decimals.Checked = need4places

        Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", PRICE_LIST_CODE)
        If rowTATCURR1 IsNot Nothing Then
            isCurrRetailPriceList = True
        Else
            isCurrRetailPriceList = False
        End If

        If EntryMode = "New" Then
            If Not isCurrRetailPriceList Then
                rowASFBASE1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            End If
        Else
            If Absx1.txtFor("CURR_CODE").Text <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPPRC2").Select("")
                    Dim ITEM_CODE = rowSOTPRIC2.Item("ITEM_CODE")
                    rowSOTPRIC2.Item("ITEM_RETAIL_PRICE") = Get_ITEM_RETAIL_PRICE_CURR(ITEM_CODE)
                Next
            End If
        End If

        For Each row As DataRow In dst.Tables("SOTPPRC2").Select()
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim I As Integer = 0
            ASCMAIN1.sql = "Select INIT_DATE, OLD_VALUE from ASTAUDT1 " _
                & " where TABLE_NAME = 'SOTPRIC2' " _
                & "   and KEY_VALUE = :PARM1 || ':' || :PARM2" _
                & "   and COLUMN_NAME = 'ITEM_PRICE'" _
                & "   and OLD_VALUE is not null" _
                & " order by INIT_DATE DESC"
            For Each rowASTAUDT1 As DataRow In
                ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV",
                    New Object() {PRICE_LIST_CODE, ITEM_CODE}).Select("", "INIT_DATE DESC")
                I += 1
                If I <= 6 Then
                    Dim II As String = Format(I, "0")
                    row.Item("ITEM_PRICE_" & II) = rowASTAUDT1.Item("OLD_VALUE")
                    row.Item("ITEM_PRICE_" & II & "_DATE") = rowASTAUDT1.Item("INIT_DATE")
                End If
            Next
        Next

        With grdSOTPPRC2.DisplayLayout.Bands(0)
            .Columns("ITEM_RETAIL_PRICE").Hidden = isCurrRetailPriceList
            .Columns("DISC_PCT_RETAIL").Hidden = isCurrRetailPriceList
            .Columns("DISC_PCT_LIST").Hidden = isCurrRetailPriceList
        End With


        EnforceConstraints(True)
    End Sub

    Function Get_ITEM_RETAIL_PRICE_CURR(ITEM_CODE As String) As Decimal
        Dim CURR_CODE As String = Absx1.txtFor("CURR_CODE").Text
        Dim rowSOTPRIC2_RETAIL As DataRow = LookUp("SOTPRIC2", New String() {CURR_CODE, ITEM_CODE})
        If rowSOTPRIC2_RETAIL Is Nothing Then
            Return 0
        Else
            Return Val(rowSOTPRIC2_RETAIL.Item("ITEM_PRICE") & "")
        End If
    End Function

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTPPRC2").Rows.Clear()
            EnforceConstraints(True)
        End If

        grdSOTPPRC2.Text = "Price List Details"

        chkShow4Decimals.Checked = False
        Toggle_Decimal_Places()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        '  grdSOTPRIC2.Enabled = tf

        Absx1.txtFor("CURR_CODE").Enabled = (EntryMode = "New")

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        Dim canApprove As Boolean = ASCMAIN1.USER_SECURITY_CODEs.Contains("PA")

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTPPRC2}
            With grd.DisplayLayout.Override
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
        Next

        chkShow4Decimals.Enabled = tf
        cmbApprover.Enabled = tf
        cmbApprover.Visible = tf
        lblApprover.Visible = tf
        btnSelectPriceList.Enabled = Not tf
        btnSelectPriceList.Visible = True

        Dim isLead As Boolean = (EntryMode & "") = ""
        Dim isEditOrNew As Boolean = ((EntryMode & "") = "New" OrElse (EntryMode & "") = "Edit")

    End Sub

#End Region

#Region "grdSOTPRIC2"
    Private Sub grdSOTPRIC2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPPRC2.AfterCellUpdate
        If e.Cell Is Nothing OrElse e.Cell.Value Is Nothing Then Exit Sub
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", e.Cell.Value)
                If rowICTITEM1 IsNot Nothing Then
                    e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1("ITEM_DESC")
                    If Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                        e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1("ITEM_RETAIL_PRICE")
                    Else

                        e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = Get_ITEM_RETAIL_PRICE_CURR(e.Cell.Value & "")
                    End If
                    e.Cell.Row.Cells("ITEM_LIST_PRICE").Value = rowICTITEM1("ITEM_LIST_PRICE")
                End If

                If txtCURR_CODE.Text <> "USD" Then
                    Dim rowSOTPRIC2 As DataRow = Nothing
                    rowSOTPRIC2 = LookUp("SOTPRIC2", New String() {txtCURR_CODE.Text, e.Cell.Row.Cells("ITEM_CODE").Value})
                    If rowSOTPRIC2 Is Nothing Then
                        '  e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = 0
                    Else
                        e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = rowSOTPRIC2.Item("ITEM_PRICE")
                    End If
                End If

            Case "ITEM_RETAIL_PRICE"


                ' grdCodeDesc(grdSOTPRIC2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
        End Select
    End Sub

    Private Sub grdSOTPRIC2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPPRC2.AfterRowActivate

        With grdSOTPPRC2.DisplayLayout.Bands(0).Columns("ITEM_CODE")
            If grdSOTPPRC2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

    End Sub

    Private Sub grdSOTPRIC2_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTPPRC2.BeforeExitEditMode
        If grdSOTPPRC2.ActiveCell IsNot Nothing Then
            With grdSOTPPRC2.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        If .EditorResolved.IsValid Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTPRIC2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTPPRC2.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTPRIC2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPPRC2.BeforeRowUpdate
        Dim row As DataRow = LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("PRICE_LIST_CODE").Value = Absx1.txtFor("PRICE_LIST_CODE").Text
        End If

        If Val(e.Row.Cells("ITEM_NEW_PRICE").Value & "") < 0 Or Val(e.Row.Cells("ITEM_NEW_SRP").Value & "") < 0 _
        Or Val(e.Row.Cells("ITEM_PRICE").Value & "") < 0 Or Val(e.Row.Cells("ITEM_SRP").Value & "") < 0 Then
            MsgBox("Price cannot be negative", MsgBoxStyle.OkOnly, "Cannot Change Price")
            e.Cancel = True
        End If

        If Not (ASCMAIN1.CLIENT = "AHA") Then
            If Not e.Row.IsAddRow And Val(e.Row.Cells("ITEM_NEW_PRICE").Text & "") <> 0 Then
                If e.Row.Cells("ITEM_NEW_PRICE_DATE").Text & "" = "" Then
                    MsgBox("New Price Date must be entered when changing price", MsgBoxStyle.OkOnly, "Cannot Change Price Record")
                    e.Cancel = True
                Else
                    If e.Row.Cells("ITEM_NEW_PRICE_DATE").DataChanged Then
                        Dim DTE As Date = CDate(e.Row.Cells("ITEM_NEW_PRICE_DATE").Text & "")
                        If Format(DTE, "dd") <> "01" Or Format(DTE, "yyyyMM") <= ASCMAIN1.CYM Then
                            MsgBox("New Price Date must be the 1st of a Future Month", MsgBoxStyle.OkOnly, "Cannot Change Price Record")
                            e.Cancel = True
                        End If
                    End If
                End If
            End If

            If Not e.Row.IsAddRow And Val(e.Row.Cells("ITEM_NEW_SRP").Text & "") <> 0 Then
                If e.Row.Cells("ITEM_NEW_SRP_DATE").Text & "" = "" Then
                    MsgBox("New SRP Price Date must be entered when changing price", MsgBoxStyle.OkOnly, "Cannot Change Price Record")
                    e.Cancel = True
                Else
                    If e.Row.Cells("ITEM_NEW_SRP_DATE").DataChanged Then
                        Dim DTE As Date = CDate(e.Row.Cells("ITEM_NEW_SRP_DATE").Text & "")
                        If Format(DTE, "dd") <> "01" Or Format(DTE, "yyyyMM") <= ASCMAIN1.CYM Then
                            MsgBox("New SRP Price Date must be the 1st of a Future Month", MsgBoxStyle.OkOnly, "Cannot Change Price Record")
                            e.Cancel = True
                        End If
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub grdSOTPRIC2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPPRC2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTITEM1.ITEM_CODE not in", "SOTPPRC2", "ITEM_CODE")
                grdClickCellButton(grdSOTPPRC2, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTPRIC2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPPRC2.InitializeRow
        grd_RowColor(dst.Tables("SOTPPRC2"), e.Row)

        For Each C As String In New String() {"ITEM_PRICE", "ITEM_NEW_PRICE", "ITEM_PRICE_1", "ITEM_PRICE_2", "ITEM_PRICE_3", "ITEM_PRICE_4", "ITEM_PRICE_5", "ITEM_PRICE_6"}
            Dim V As Decimal = Val(e.Row.Cells(C).Value & "")
            If Val(Format(V, "#.00")) = V Or chkShow4Decimals.Checked Then
                e.Row.Cells(C).Appearance = decimalsOK
                e.Row.Cells(C).ToolTipText = ""
            Else
                e.Row.Cells(C).Appearance = decimalsConcealed
                e.Row.Cells(C).ToolTipText = "More than 2 Decimal Places: " & Format(V, "#.0000")
            End If
        Next

    End Sub


    Public Overrides Sub Excel_Export_Post_Process(FILENAME As String)
        Exit Sub

        If ASCMAIN1.Running_in_VS Then
        Else
            Exit Sub

        End If



        Stop

        Dim r As Integer = 0

        Dim excel As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim xlSourceRange As Microsoft.Office.Interop.Excel.Range = Nothing
        Dim xlDestRange As Microsoft.Office.Interop.Excel.Range = Nothing

        excel = New Microsoft.Office.Interop.Excel.Application
        wb = excel.Workbooks.Open(FILENAME)
        ws = wb.Worksheets(0)
        xlSourceRange = ws.Range(6, 1).EntireRow
        xlSourceRange.WrapText = True
        xlSourceRange.AutoFit()

        Stop
    End Sub

#End Region

    Private Sub cmdExport_Click(sender As Object, e As EventArgs) Handles cmdExport.Click

        Sort_grdColumns(grdSOTPPRC2, "ITEM_CODE")

        Set_DX_Column(grdSOTPPRC2, "")
        For Each GCOL As UltraWinGrid.UltraGridColumn In grdSOTPPRC2.DisplayLayout.Bands(0).Columns
            Set_DX_Column(grdSOTPPRC2, GCOL.Key, GCOL.Header.Caption, GCOL.Width)
        Next

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Load_DataTable_into_SGXLS(1, 1, dst.Tables("SOTPPRC2"), workbook.ActiveWorksheet, grdSOTPPRC2, Nothing, "ITEM_CODE", "")
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("SOTPRIC2.XLSX_NO") & ".XLSX"

        workbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Show_Document(FILENAME)
    End Sub

    Private Sub chkShow4Decimals_CheckedChanged(sender As Object, e As EventArgs) Handles chkShow4Decimals.CheckedChanged
        'Dim ITEM_NEW_PRICE As Decimal = VAlign(GRD)
        Toggle_Decimal_Places()
    End Sub

    Sub Toggle_Decimal_Places()
        For Each C As String In New String() _
    {"ITEM_PRICE", "ITEM_NEW_PRICE", "ITEM_PRICE_1", "ITEM_PRICE_2", "ITEM_PRICE_3", "ITEM_PRICE_4", "ITEM_PRICE_5", "ITEM_PRICE_6"}
            With grdSOTPPRC2.DisplayLayout.Bands(0).Columns(C)
                If chkShow4Decimals.Checked Then
                    .Format = "$#.0000"
                    .MaskInput = "nnnnnn.nnnn"
                Else
                    .Format = "$#.00"
                    .MaskInput = "nnnnnn.nn"
                End If
            End With
        Next

        grdSOTPPRC2.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Private Sub Load_Approver_List()
        'ASCMAIN1.sql = "select USER_ID, USER_NAME " _
        '              & "  from ASTUSER1 " _
        '              & " where USER_SECURITY_CODE like '%PLAPPR%' " _
        '              & " order by USER_NAME"

        'Dim dt As DataTable = ASCDATA1.GetDataTable()

        'If dt Is Nothing Then Exit Sub

        Dim dt As New DataTable
        dt.Columns.Add("USER_ID", GetType(String))
        dt.Columns.Add("USER_NAME", GetType(String))

        dt.Rows.Add("nl", "Nick")
        dt.Rows.Add("smz", "Stephanie")
        dt.Rows.Add("PIERRE", "Pierre")

        cmbApprover.DataSource = dt
        cmbApprover.DisplayMember = "USER_NAME"
        cmbApprover.ValueMember = "USER_ID"

        'nothing selected by default
        cmbApprover.SelectedIndex = -1
    End Sub
    Private Function Get_Open_Pending_Info(PRICE_LIST_CODE As String) As DataRow
        Dim plc As String = (PRICE_LIST_CODE & "").Replace("'", "''")

        Dim sql As String =
        "select PRICE_LIST_CODE, SUBMIT_OPER, APPROVER_OPER, SUBMIT_DATE " &
        "  from SOTPPRC1 " &
        " where PRICE_LIST_CODE = '" & plc & "' " &
        "   and (STATUS = 'P' OR STATUS = 'D')"

        Dim dt As DataTable = ASCDATA1.GetDataTable(sql)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return Nothing
        Return dt.Rows(0)
    End Function

    Private Function Submit_For_Approval_From_Update() As Boolean

        If IsSubmittingForApproval Then Return False

        If cmbApprover.SelectedIndex < 0 Then
            MsgBox("Please select an approver before saving/updating.",
               MsgBoxStyle.OkOnly, "Approver Required")
            cmbApprover.Focus()
            Return False
        End If

        Dim PRICE_LIST_CODE As String = (Absx1.txtFor("PRICE_LIST_CODE").Text & "").Trim()
        Dim PRICE_LIST_DESC As String = (Absx1.txtFor("PRICE_LIST_DESC").Text & "").Trim()
        Dim CURR_CODE As String = (Absx1.txtFor("CURR_CODE").Text & "").Trim()

        If PRICE_LIST_CODE = "" Then
            MsgBox("Price List Code is required before saving/updating.",
               MsgBoxStyle.OkOnly, "Missing Price List Code")
            Return False
        End If

        Dim approverId As String = (cmbApprover.SelectedValue & "").Trim()
        Dim approverName As String = (cmbApprover.Text & "").Trim()
        Dim approverEmail As String = Get_Approver_Email(approverId, approverName)
        If approverEmail = "" Then Return False

        IsSubmittingForApproval = True

        Dim safePLC As String = SqlSafe(PRICE_LIST_CODE)
        Dim safeDesc As String = SqlSafe(PRICE_LIST_DESC)
        Dim safeCurr As String = SqlSafe(CURR_CODE)
        Dim safeUser As String = SqlSafe(ASCMAIN1.USER_ID)
        Dim safeAppr As String = SqlSafe(approverId)
        Dim dtStamp As String = SqlDateTimeStamp()

        Try
            '=== If someone else already has it pending, block ===
            Dim pendingRow As DataRow = Get_Open_Pending_Info(PRICE_LIST_CODE)
            If pendingRow IsNot Nothing Then
                Dim existingSubmitOper As String = (pendingRow("SUBMIT_OPER") & "").Trim()
                If existingSubmitOper <> "" AndAlso existingSubmitOper.ToUpper() <> (ASCMAIN1.USER_ID & "").ToUpper() Then
                    MsgBox("This price list already has a pending submission by " & existingSubmitOper & ".",
                       MsgBoxStyle.OkOnly, "Already Pending")
                    Return False
                End If
            End If

            '=== UPSERT HEADER (SOTPPRC1) by PRICE_LIST_CODE ===
            Dim dtH As DataTable = ASCDATA1.GetDataTable(
            "select 1 from SOTPPRC1 where PRICE_LIST_CODE = '" & safePLC & "'"
        )
            Dim exists As Boolean = (dtH IsNot Nothing AndAlso dtH.Rows.Count > 0)

            If Not exists Then
                Dim sqlIns As String =
                "insert into SOTPPRC1 " &
                "(PRICE_LIST_CODE, PRICE_LIST_DESC, CURR_CODE, " &
                " INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE, " &
                " STATUS, SUBMIT_OPER, SUBMIT_DATE, APPROVER_OPER, APPROVED_DATE) " &
                "values (" &
                "'" & safePLC & "'," &
                "'" & safeDesc & "'," &
                "'" & safeCurr & "'," &
                "'" & safeUser & "'," &
                "'" & safeUser & "'," &
                "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                "'P'," &
                "'" & safeUser & "'," &
                "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                "'" & safeAppr & "'," &
                "NULL)"

                ASCDATA1.ExecuteSQL(sqlIns)
            Else
                Dim sqlUpd As String =
                "update SOTPPRC1 set " &
                " PRICE_LIST_DESC = '" & safeDesc & "'," &
                " CURR_CODE = '" & safeCurr & "'," &
                " LAST_OPER = '" & safeUser & "'," &
                " LAST_DATE = TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                " STATUS = 'P'," &
                " SUBMIT_OPER = '" & safeUser & "'," &
                " SUBMIT_DATE = TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                " APPROVER_OPER = '" & safeAppr & "'," &
                " APPROVED_DATE = NULL " &
                "where PRICE_LIST_CODE = '" & safePLC & "'"

                ASCDATA1.ExecuteSQL(sqlUpd)
            End If

            '=== OVERWRITE DETAILS (SOTPPRC2) by PRICE_LIST_CODE ===
            ASCDATA1.ExecuteSQL("delete from SOTPPRC2 where PRICE_LIST_CODE = '" & safePLC & "'")

            Dim dtDetail As DataTable = dst.Tables("SOTPPRC2")
            If dtDetail IsNot Nothing Then
                For Each row As DataRow In dtDetail.Rows
                    If row.RowState = DataRowState.Deleted Then Continue For

                    Dim itemCode As String = (row("ITEM_CODE") & "").Trim()
                    If itemCode = "" Then Continue For

                    Dim safeItem As String = SqlSafe(itemCode)

                    Dim ITEM_PRICE As Decimal = Val(row("ITEM_PRICE") & "")
                    Dim ITEM_NEW_PRICE As Decimal = Val(row("ITEM_NEW_PRICE") & "")
                    Dim ITEM_SRP As Decimal = Val(row("ITEM_SRP") & "")
                    Dim ITEM_NEW_SRP As Decimal = Val(row("ITEM_NEW_SRP") & "")

                    Dim newPriceDateSql As String = "NULL"
                    If dtDetail.Columns.Contains("ITEM_NEW_PRICE_DATE") AndAlso Not IsDBNull(row("ITEM_NEW_PRICE_DATE")) AndAlso (row("ITEM_NEW_PRICE_DATE") & "") <> "" Then
                        newPriceDateSql = SqlDateOnly(CDate(row("ITEM_NEW_PRICE_DATE")))
                    End If

                    Dim newSrpDateSql As String = "NULL"
                    If dtDetail.Columns.Contains("ITEM_NEW_SRP_DATE") AndAlso Not IsDBNull(row("ITEM_NEW_SRP_DATE")) AndAlso (row("ITEM_NEW_SRP_DATE") & "") <> "" Then
                        newSrpDateSql = SqlDateOnly(CDate(row("ITEM_NEW_SRP_DATE")))
                    End If

                    Dim sqlDtl As String =
                    "insert into SOTPPRC2 " &
                    "(PRICE_LIST_CODE, ITEM_CODE, " &
                    " ITEM_PRICE, ITEM_NEW_PRICE, ITEM_NEW_PRICE_DATE, " &
                    " ITEM_SRP, ITEM_NEW_SRP, ITEM_NEW_SRP_DATE, " &
                    " INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE) " &
                    "values (" &
                    "'" & safePLC & "'," &
                    "'" & safeItem & "'," &
                    ITEM_PRICE.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    ITEM_NEW_PRICE.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    newPriceDateSql & "," &
                    ITEM_SRP.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    ITEM_NEW_SRP.ToString(Globalization.CultureInfo.InvariantCulture) & "," &
                    newSrpDateSql & "," &
                    "'" & safeUser & "'," &
                    "'" & safeUser & "'," &
                    "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS')," &
                    "TO_DATE('" & dtStamp & "','YYYYMMDDHH24MISS'))"

                    ASCDATA1.ExecuteSQL(sqlDtl)
                Next
            End If

            CurrentPriceListCodeSubmitted = PRICE_LIST_CODE

            MsgBox("Price list " & PRICE_LIST_CODE & " submitted for approval." & vbCrLf &
               "Approver: " & approverName,
               MsgBoxStyle.Information, "Price List Approval")

            Return True

        Catch ex As Exception
            MsgBox("Error while submitting for approval:" & vbCrLf & ex.Message,
               MsgBoxStyle.Critical, "Submit For Approval")
            Return False

        Finally
            IsSubmittingForApproval = False
        End Try

    End Function
    Private Function Get_Approver_Email(ByVal approverId As String, ByVal approverName As String) As String
        Dim approverEmail As String = ""

        'TODO: replace with ASTUSER1 lookup
        Select Case (approverId & "").Trim().ToUpper()
            Case "NL"
                approverEmail = "nicholas@absolution.com"
            Case "SMZ"
                approverEmail = "stephanie@absolution.com"
            Case "PIERRE"
                approverEmail = ""
        End Select

        If approverEmail = "" Then
            MsgBox("Could not determine an email address for " & approverName & ". " &
               "Please update ASTUSER1 or the fallback mapping.",
               MsgBoxStyle.OkOnly, "Missing Approver Email")
        End If

        Return approverEmail
    End Function
    Private Function SqlDateTimeStamp() As String
        Return Format(DATETIME_STAMP, "yyyyMMddHHmmss")
    End Function

    Private Function SqlDateOnly(d As Date) As String
        Return "TO_DATE('" & Format(d, "yyyyMMdd") & "','YYYYMMDD')"
    End Function
    Private Function SqlSafe(s As String) As String
        Return (s & "").Replace("'", "''").Trim()
    End Function

    Private Function LiveHeaderExists(plc As String) As Boolean
        Dim safe As String = SqlSafe(plc)
        Dim dt As DataTable = ASCDATA1.GetDataTable(
        "select 1 from SOTPRIC1 where PRICE_LIST_CODE = '" & safe & "' and rownum = 1"
    )
        Return (dt IsNot Nothing AndAlso dt.Rows.Count > 0)
    End Function
    Private Sub PreloadFromLive(plc As String)
        Dim safe As String = SqlSafe(plc)

        Dim dtH As DataTable = ASCDATA1.GetDataTable(
        "select PRICE_LIST_DESC, CURR_CODE from SOTPRIC1 where PRICE_LIST_CODE = '" & safe & "'"
    )
        If dtH IsNot Nothing AndAlso dtH.Rows.Count > 0 Then
            rowASFBASE1("PRICE_LIST_DESC") = (dtH.Rows(0)("PRICE_LIST_DESC") & "").Trim()
            rowASFBASE1("CURR_CODE") = (dtH.Rows(0)("CURR_CODE") & "").Trim()
        End If

        EnforceConstraints(False)

        dst.Tables("SOTPRIC2").Rows.Clear()
        dst.Tables("SOTPPRC2").Rows.Clear()

        Fill_Records("SOTPRIC2", New String() {plc})

        For Each src As DataRow In dst.Tables("SOTPRIC2").Rows

            If src.RowState = DataRowState.Deleted Then Continue For

            Dim dest As DataRow = dst.Tables("SOTPPRC2").NewRow()

            For Each col As DataColumn In dst.Tables("SOTPPRC2").Columns

                If (col.Expression & "") <> "" OrElse col.ReadOnly Then Continue For

                If dst.Tables("SOTPRIC2").Columns.Contains(col.ColumnName) Then
                    dest(col.ColumnName) = src(col.ColumnName)
                End If
            Next

            dest("PRICE_LIST_CODE") = plc

            dst.Tables("SOTPPRC2").Rows.Add(dest)
        Next

        dst.Tables("SOTPRIC2").Rows.Clear()

        EnforceConstraints(True)
    End Sub


    Private Sub StartBlankNew(plc As String)
        Absx1.txtFor("CURR_CODE").Text = (ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "")

        EnforceConstraints(False)
        dst.Tables("SOTPPRC2").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub btnSelectPriceList_Click(sender As Object, e As EventArgs) Handles btnSelectPriceList.Click
        '    If EntryMode = "New" OrElse EntryMode = "Edit" Then Exit Sub

        '    Dim dt As DataTable = ASCDATA1.GetDataTable(
        '    "select PRICE_LIST_CODE, PRICE_LIST_DESC, CURR_CODE " &
        '    "  from SOTPRIC1 " &
        '    " order by PRICE_LIST_CODE"
        ')

        '    Dim pick = frmPriceListPicker.Pick(Me, dt)
        '    If pick Is Nothing Then Exit Sub

        '    Dim plc As String = (pick.Code & "").Trim()
        '    If plc = "" Then Exit Sub

        '    Absx1.txtFor("PRICE_LIST_CODE").Text = plc

        '    Dim pending As DataRow = Get_Open_Pending_Info(plc)
        '    If pending IsNot Nothing Then
        '        Call Click_Command("Edit")
        '    Else
        '        Call Click_Command("New")
        '    End If
        If EntryMode = "New" OrElse EntryMode = "Edit" Then Exit Sub

        Dim plc As String = Select_PriceList_From_CodeSelector() & ""
        If plc = "" Then Exit Sub

        Absx1.txtFor("PRICE_LIST_CODE").Text = plc

        Dim pending As DataRow = Get_Open_Pending_Info(plc)
        If pending IsNot Nothing Then
            Call Click_Command("Edit")
        Else
            Call Click_Command("New")
        End If
    End Sub
    Private Function Select_PriceList_From_CodeSelector() As String
        Dim PRICE_LIST_CODE As String = ""

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PRICE_LIST_CODE", "SOTPRIC1")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            PRICE_LIST_CODE = ASCMAIN1.CodeSelector.SelectedCode
        End If

        Return PRICE_LIST_CODE
    End Function

End Class
