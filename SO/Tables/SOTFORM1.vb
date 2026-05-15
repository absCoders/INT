Public Class SOTFORM1

    Dim sqlSOTFORM2 As String = ""
    Dim usersAddedSinceLastEmail As Integer
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")
        Dim defWhse As String = "ADSBLA"
        If Not ROWs Is Nothing AndAlso ROWs.ContainsKey("SOTPARM1") Then
            defWhse = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")

        End If
        With dst
            ASCMAIN1.sql = "Select SOTFORM2.*, ICTITEM1.ITEM_DESC" _
                & " from SOTFORM2,ICTITEM1" _
                & " where SOTFORM2.ORDR_FORM_CODE = :PARM1" _
                & "   and ICTITEM1.ITEM_CODE = SOTFORM2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTFORM2", "**", 0, True, "V", 2)

            ASCMAIN1.sql = $"SELECT X.ITEM_CODE, X.ITEM_DESC, X.ITEM_RETAIL_PRICE, ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_CODE, Q.QTY, FC_TTL,
DPTMRPG1.QTY_00 AS POS
FROM ICTCOLL1, 
(SELECT ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, COLLECTION_CODE
 FROM ICTITEM1
 WHERE ITEM_SNU_CODE = 'S' AND ITEM_BASIC_PROMO = 'B'
 AND ITEM_STATUS = 'A'
) X,
(SELECT ITEM_CODE,
SUM(NVL(FORECAST, 0)) AS FC_TTL
 FROM DPTITMF1 
 WHERE OPS_YYYYPP IN (TO_CHAR(ADD_MONTHS(SYSDATE, 0), 'YYYYMM'), TO_CHAR(ADD_MONTHS(SYSDATE, -1), 'YYYYMM'), TO_CHAR(ADD_MONTHS(SYSDATE, -2), 'YYYYMM'))
 GROUP BY ITEM_CODE
) F,
(SELECT ITEM_CODE, SUM(WHSE_QTY_ON_HAND - WHSE_QTY_OPEN - WHSE_QTY_PICK) AS QTY 
 FROM ICTSTAT2 
 WHERE WHSE_CODE = '{defWhse}' 
 GROUP BY ITEM_CODE
) Q,
DPTMRPG1
WHERE X.ITEM_CODE = F.ITEM_CODE(+)
AND X.ITEM_CODE = Q.ITEM_CODE
AND X.ITEM_CODE = DPTMRPG1.ITEM_CODE
AND DPTMRPG1.MRP_TYPE(+) = '6'
AND ICTCOLL1.COLLECTION_CODE = X.COLLECTION_CODE
AND NVL(Q.QTY,0) > 0
AND X.ITEM_CODE = F.ITEM_CODE(+)"

            Create_TDA(.Tables.Add, "SOTFORMI", "**", 0, False, "", 1)

            Create_TDA(dst.Tables.Add, "SOTRECEM", "*", 0)
            Create_TDA(dst.Tables.Add, "SOTPARMG", "*", 0)
            Create_TDA(dst.Tables.Add, "SOTFORMU", "*", 0)

        End With
        grdSOTFORM2.DataSource = dst.Tables("SOTFORM2")
        grdSOTFORMI.DataSource = dst.Tables("SOTFORMI")
        grdSOTRECEM.DataSource = dst.Tables("SOTRECEM")
        Fill_Records("SOTFORMI")
        Fill_Records("SOTFORMU")
        Fill_Records("SOTPARMG")

        'counts for both grids
        Create_Summary(grdSOTFORM2, "ORDR_FORM_LNO", "Count")
        Create_Summary(grdSOTFORMI, "ITEM_CODE", "Count")
        Sort_grdColumns(grdSOTFORMI, "ITEM_CODE")
        Show_Filter(grdSOTFORMI)

        With grdSOTFORM2.DisplayLayout.Bands(0)
            .Columns("ITEM_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
        End With
        Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_FORM_EMAIL_SENT_DATE"), True, "SOTFORM1")
        Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_FORM_EMAIL_SENT_BY"), True, "SOTFORM1")

        UpdateUsersAddedLabel()
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTFORM2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
        Load_Popup_Menu(grdSOTFORMI, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

            Case "grdSOTFORM2"
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
                If grd.Name = "grdSOTFORM2" Then
                    Add_Codes(grdSOTFORM2, "ICTITEM1", "ITEM_CODE", "Items")
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If
    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        EMsg = ""
        Select Case eItemKey
            Case "New"
            Case "Edit"
            Case "Update"


                If optFormType.Value = "" Then
                    EMsg &= vbCr & "Form Type must be selected"
                End If

                If optFormStatus.Value = "" Then
                    EMsg &= vbCr & "Form Status must be selected"
                End If

                If txtDesc.Text & "" = "" Then
                    EMsg &= vbCr & "Description is Mandatory"
                End If
                'if form type is gratis
                If optFormType.Value = "G" Then
                    'start date can't be null
                    Dim DATE_START As Date = Absx1.dteFor("ORDR_FORM_VALID_FROM").Value
                    If DATE_START & "" = "12:00:00 AM" Then
                        EMsg &= vbCr & "Start Date Is Mandatory"
                    End If

                    'end date can't be null
                    Dim DATE_END As Date = Absx1.dteFor("ORDR_FORM_VALID_UNTIL").Value
                    If DATE_END & "" = "12:00:00 AM" Then
                        EMsg &= vbCr & "End Date Is Mandatory"
                    End If

                    'start date cant be later than end date
                    If DATE_START & "" <> "" And DATE_END & "" <> "" Then
                        If Format(DATE_START, "yyyyMMdd") > Format(DATE_END, "yyyyMMdd") Then
                            EMsg &= vbCr & "Start Date may Not be later than End Date"
                        End If
                    End If

                    'max retail must be greater than zero
                    Dim MAX_RTL As Decimal = System.Math.Round(Val(Absx1.numFor("ORDR_FORM_MAX_RETAIL").Value & ""), 2)
                    If MAX_RTL & "" <= 0 Then
                        EMsg &= vbCr & "Max Retail must be greater than zero"
                    End If

                    'season code can't be empty
                    If Absx1.txtFor("SEASON_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must Specify a Season"
                    Else
                        'season code can't be invalid
                        Dim row As DataRow = LookUp("ICTSEAS1", Absx1.txtFor("SEASON_CODE").Text)
                        If row Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Season"
                        Else
                            Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
                            If Mid(SEASON_CODE, 1, 4) <> Format(DATE_START, "yyyy") Then
                                EMsg &= vbCr & "Season Not congruous with Start Date"
                            Else
                                If Mid(SEASON_CODE, 5, 1) = "S" And Format(DATE_START, "MM") >= "07" Then
                                    EMsg &= vbCr & "Season Not congruous with Start Date"
                                End If
                                If Mid(SEASON_CODE, 5, 1) = "F" And Format(DATE_START, "MM") < "07" Then
                                    EMsg &= vbCr & "Season Not congruous with Start Date"
                                End If
                            End If

                        End If
                    End If
                    If EntryMode = "New" Then
                        Dim SEASON_CODE As String = txt_SEASON_CODE.Text
                        If SeasonCodeExists(SEASON_CODE) Then
                            EMsg &= vbCr & "A Gratis Form with this Season Code already exists"
                        End If
                        Dim ORDR_FORM_VALID_FROM As Date = Absx1.dteFor("ORDR_FORM_VALID_FROM").Value
                        Dim ORDR_FORM_VALID_UNTIL As Date = Absx1.dteFor("ORDR_FORM_VALID_UNTIL").Value
                        If CheckDateOverlap(ORDR_FORM_VALID_FROM, ORDR_FORM_VALID_UNTIL) Then
                            EMsg &= vbCr & "The date range overlaps with another active Gratis form."
                        End If
                    End If
                End If
        End Select

        'If EMsg <> "" Then
        '    MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
        '    Exit Sub
        'End If
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim ORDR_FORM_CODE As String = Absx1.txtFor("ORDR_FORM_CODE").Text
        Dim sqlDelete = "ORDR_FORM_CODE = '" & ORDR_FORM_CODE & "'"
        Update_Record_TDA("SOTFORM2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("SOTFORM2", New String() {Absx1.txtFor("ORDR_FORM_CODE").Text})
        Sort_grdColumns(grdSOTFORM2, "ORDR_FORM_LNO")

        Fill_Records("SOTFORMI")
        Dim ORDR_FORM_CODE As String = txt_ORDR_FORM_CODE.Text
        ASCMAIN1.sql = $"Select * from SOTRECEM WHERE ORDR_FORM_CODE = '{ORDR_FORM_CODE}'"
        Fill_Records("SOTRECEM",,, ASCMAIN1.sql)
        Fill_Records("SOTFORMU")
        UpdateUsersAddedLabel()
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SOTFORM2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSOTFORM2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        If EntryMode <> "New" Then
            Set_Read_Only_for_ctl(optFormType, True, "SOTFORM1")
        End If
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTFORM2}
            With grd.DisplayLayout.Override
                Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_FORM_EMAIL_SENT_DATE"), True, "SOTFORM1")
                Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_FORM_EMAIL_SENT_BY"), True, "SOTFORM1")
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
        'send emails button shouldn't work unless inquiry mode and sentby is null
        Dim isSent As Boolean = String.IsNullOrEmpty(UltraTextEditor3.Text)
        btnSendEmails.Enabled = (EntryMode = "Edit")
        btnResetUsers.Enabled = EntryMode = "Edit"

        grpGratisParm.Visible = ScreenMode And optFormType.Value = "G"
        grdSOTFORM2.Visible = ScreenMode
        lblFormType.Visible = ScreenMode
        optFormType.Visible = ScreenMode
        optBasicPromo.Enabled = Not ScreenMode
        optSaleable.Enabled = Not ScreenMode
        grpItemSelector.Visible = Not ScreenMode


    End Sub

#End Region

#Region "grdSOTFORM2"

    Private Sub grdSOTFORM2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTFORM2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                grdCodeDesc(grdSOTFORM2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
        End Select
    End Sub

    Private Sub grdSOTFORM2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTFORM2.AfterRowActivate

        With grdSOTFORM2.DisplayLayout.Bands(0).Columns("ITEM_CODE")
            If grdSOTFORM2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSOTFORM2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTFORM2.AfterRowsDeleted

    End Sub

    Private Sub grdSOTFORM2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTFORM2.AfterRowUpdate

    End Sub

    Private Sub grdSOTFORM2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTFORM2.BeforeRowsDeleted
        'For Each grow As UltraWinGrid.UltraGridRow In e.Rows
        '    Dim SREP_CODE As String = grow.Cells("SREP_CODE").Value
        '    Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
        '    Dim rowSOTFORM2 As DataRow = dst.Tables("SOTFORM2").Rows.Find(New String() {SREP_CODE, CUST_CODE})
        '    If Not rowSOTFORM2.RowState = DataRowState.Added Then
        '        MsgBox("Cannot Delete Previously Updated Colors", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
        '        e.Cancel = True
        '    End If
        'Next
    End Sub

    Private Sub grdSOTFORM2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTFORM2.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_FORM_CODE").Value = Absx1.txtFor("ORDR_FORM_CODE").Text
            e.Row.Cells("ORDR_FORM_LNO").Value = Val(dst.Tables("SOTFORM2").Compute("MAX(ORDR_FORM_LNO)", "") & "") + 10
        End If

    End Sub

    Private Sub grdSOTFORM2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTFORM2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTITEM1.ITEM_CODE not in", "SOTFORM2", "ITEM_CODE")
                grdClickCellButton(grdSOTFORM2, sql_where, True)
        End Select
    End Sub

#End Region

    Sub Send_email()
        Dim DT As DataTable = GetActiveUsersToEmail()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing email")

        Dim bodyTemplate As String = dst.Tables("SOTPARMG").Rows(0)("SO_PARM_EMAIL_BODY").ToString

        Dim ORDR_FORM_CODE As String = txt_ORDR_FORM_CODE.Text
        ASCMAIN1.sql = $"SELECT ORDR_FORM_VALID_FROM, ORDR_FORM_VALID_UNTIL FROM SOTFORM1 WHERE ORDR_FORM_CODE = '{ORDR_FORM_CODE}'"
        Dim dateTable As DataTable = ASCDATA1.GetDataTable

        Dim ORDR_FORM_VALID_FROM As Date = If(IsDBNull(dateTable.Rows(0)("ORDR_FORM_VALID_FROM")), Date.MinValue, Convert.ToDateTime(dateTable.Rows(0)("ORDR_FORM_VALID_FROM")))
        Dim ORDR_FORM_VALID_UNTIL As Date = If(IsDBNull(dateTable.Rows(0)("ORDR_FORM_VALID_UNTIL")), Date.MinValue, Convert.ToDateTime(dateTable.Rows(0)("ORDR_FORM_VALID_UNTIL")))

        If DT.Rows.Count > 0 Then
            For Each row As DataRow In DT.Rows
                Dim ORDR_FORM_USER_EMAIL As String = row("ORDR_FORM_USER_EMAIL").ToString()
                Dim ORDR_FORM_USER_NAME As String = row("ORDR_FORM_USER_NAME").ToString()
                Dim ORDR_FORM_USER_NO As String = row("ORDR_FORM_USER_NO").ToString()
                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                EMAIL_ADDRESSs.Add(ORDR_FORM_USER_EMAIL, ORDR_FORM_USER_NAME)
                Dim ATTACHMENTs As New Dictionary(Of String, String)

                Dim EMAIL_SUBJECT As String = "Gratis Order Form"
                Dim link As String = "https://intportal.interparfums.com/gratisOrder/" & ORDR_FORM_USER_NO & "/" & ORDR_FORM_CODE

                Dim EMAIL_BODY As String = bodyTemplate.Replace("{ORDR_FORM_USER_NAME}", ORDR_FORM_USER_NAME)
                EMAIL_BODY = EMAIL_BODY.Replace("This product allowance is good from {START_DATE} to {END_DATE}", $"This product allowance is good from {ORDR_FORM_VALID_FROM.ToString("MMMM dd, yyyy")} to {ORDR_FORM_VALID_UNTIL.ToString("MMMM dd, yyyy")}")
                EMAIL_BODY = EMAIL_BODY.Replace("click this link", "<a href='" & link & "' style='text-decoration: underline;'>Click this link</a>")

                ' Send the email
                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    EMAIL_SUBJECT, "GRATIS", True, False, "99999", "ORDR_FORM_CODE", "", EMAIL_BODY)

                LogEmailRecipient(ORDR_FORM_USER_NO, ORDR_FORM_USER_NAME, ORDR_FORM_USER_EMAIL, "A")
                Write_Event_Log("SOTFORM1", "GRATIS", "Email sent!")
            Next
        End If
        ASCMAIN1.sql = $"SELECT * FROM SOTRECEM WHERE ORDR_FORM_CODE = '{ORDR_FORM_CODE}'"
        Fill_Records("SOTRECEM",,, ASCMAIN1.sql)
        grdSOTRECEM.Refresh()
        UpdateUsersAddedLabel()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        UltraDateTimeEditor1.Value = DateTime.Now
        UltraTextEditor3.Text = ASCMAIN1.USER_ID
        MsgBox("email Sent", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Private Sub optFormType_ValueChanged(sender As Object, e As EventArgs) Handles optFormType.ValueChanged
        If Not optFormType.Value = "G" Then
            ClearGratisFields() ' Call the method to clear fields
        End If
        grpGratisParm.Visible = (optFormType.Value = "G")
    End Sub
    Private Sub grdSOTFORMI_ClickCell(sender As Object, e As UltraWinGrid.ClickCellEventArgs)
        If e.Cell.Column.Key = "SEL" AndAlso e.Cell.IsDataCell Then
            ' Toggle the value of the "SEL" cell
            e.Cell.Value = Not CBool(e.Cell.Value)
            grdSOTFORMI.Refresh()

            e.Cell.Row.Selected = CBool(e.Cell.Value)
        End If
    End Sub

    Private Sub btnSendEmails_Click(sender As Object, e As EventArgs) Handles btnSendEmails.Click
        'button should mt lock SOTFORM1.ORDR_FORM_CODE
        ASCMAIN1.Logical_Lock("SOTFORM1", "ORDR_FORM_CODE")

        'FILL_SOTRECEM()
        Send_email()

        'button becomes disabled after sending the email
        btnSendEmails.Enabled = False

        ASCMAIN1.MultiTask_Release()
    End Sub
    Private Sub btnFindItems_Click(sender As Object, e As EventArgs) Handles btnFindItems.Click
        Load_Grid_Data()
    End Sub
    Private Sub Load_Grid_Data()
        Dim basicPromoFilter As String = String.Empty
        Dim saleableFilter As String = String.Empty
        Dim title As String = ""
        Dim promoType As String = ""
        Dim OPS_YYYYPP As String = ASCMAIN1.CYP
        Dim WHSE_CODE As String = txtWhse.Text
        Dim MARKET_CODE As String = txtMarket.Text

        ' Determine the filter for Basic Promo
        Select Case optBasicPromo.Value
            Case "B"
                basicPromoFilter = "AND ITEM_BASIC_PROMO = 'B'"
                promoType = "Basic"
            Case "P"
                basicPromoFilter = "AND ITEM_BASIC_PROMO = 'P'"
                promoType = "Promo"
            Case Else
                basicPromoFilter = ""
        End Select

        ' Determine the filter for Saleable
        Select Case optSaleable.Value
            Case "S"
                saleableFilter = "AND ITEM_SNU_CODE = 'S'"
                title = "Saleable"
            Case "N"
                saleableFilter = "AND ITEM_SNU_CODE = 'N'"
                title = "No-Charge"
            Case Else
                saleableFilter = ""
        End Select

        ASCMAIN1.sql = "SELECT X.ITEM_CODE, X.ITEM_DESC, X.ITEM_RETAIL_PRICE, ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_CODE, Q.QTY, F.FC_TTL, " &
          "DPTMRPG1.QTY_00 AS POS " &
          "FROM ICTCOLL1, " &
          "(SELECT ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, COLLECTION_CODE FROM ICTITEM1 WHERE ITEM_STATUS = 'A' " & saleableFilter & " " & basicPromoFilter & ") X, " &
          "(SELECT ITEM_CODE, SUM(NVL(FORECAST, 0)) AS FC_TTL FROM DPTITMF1 WHERE OPS_YYYYPP IN (TO_CHAR(ADD_MONTHS(SYSDATE, 0), 'YYYYMM'), " &
          "TO_CHAR(ADD_MONTHS(SYSDATE, -1), 'YYYYMM'), TO_CHAR(ADD_MONTHS(SYSDATE, -2), 'YYYYMM')) GROUP BY ITEM_CODE) F, " &
          "(SELECT ITEM_CODE, SUM(WHSE_QTY_ON_HAND - WHSE_QTY_OPEN - WHSE_QTY_PICK) AS QTY FROM ICTSTAT2 WHERE WHSE_CODE = '" & WHSE_CODE & "' GROUP BY ITEM_CODE) Q, " &
          "DPTMRPG1 " &
          "WHERE X.ITEM_CODE = F.ITEM_CODE(+) AND X.ITEM_CODE = Q.ITEM_CODE AND X.ITEM_CODE = DPTMRPG1.ITEM_CODE AND DPTMRPG1.MRP_TYPE(+) = '6' AND ICTCOLL1.COLLECTION_CODE = X.COLLECTION_CODE AND NVL(Q.QTY,0) > 0 AND X.ITEM_CODE = F.ITEM_CODE(+)"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        If DT.Rows.Count > 0 Then
            grdSOTFORMI.DataSource = DT
            grdSOTFORMI.Refresh()
        Else
            MessageBox.Show("No data found.")
        End If
        'Fill_Records("SOTFORMI", "", True, ASCMAIN1.sql)
        'grdSOTFORMI.DataSource = dst.Tables("SOTFORMI")
        grdSOTFORMI.Text = title & " " & promoType & " Pick List"
        Sort_grdColumns(grdSOTFORMI, "ITEM_CODE")
        grdSOTFORMI.Refresh()
    End Sub
    Private Sub grpItemSelector_VisibleChanged(sender As Object, e As EventArgs) Handles grpItemSelector.VisibleChanged
        If grpItemSelector.Visible Then
            optBasicPromo.Value = "B"
            optSaleable.Value = "S"
            txtMarket.Text = "DPT"
            If Not ROWs Is Nothing AndAlso ROWs.ContainsKey("SOTPARM1") Then
                txtWhse.Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
            End If

            Set_Read_Only_for_ctl(txtWhse, True)
            Set_Read_Only_for_ctl(txtMarket, True)
        End If
    End Sub
    Private Sub ClearGratisFields()
        txt_SEASON_CODE.Text = String.Empty ' Clear Season Code
        dteORDR_FORM_VALID_FROM.Value = Nothing ' Clear Valid From Date
        dteORDR_FORM_VALID_UNTIL.Value = Nothing ' Clear Valid Until Date
        UltraNumericEditor1.Value = Nothing ' Clear Max Retail
    End Sub
    Private Function SeasonCodeExists(seasonCode As String) As Boolean
        ASCMAIN1.sql = $"SELECT COUNT(*) FROM SOTFORM1 WHERE SEASON_CODE = '{seasonCode}' and ORDR_FORM_STATUS = 'A' "
        Dim count As Integer = Convert.ToInt32(ASCDATA1.GetDataValue)
        Return count > 0
    End Function
    Private Function CheckDateOverlap(startDate As Date, endDate As Date) As Boolean
        Dim formattedStartDate As String = startDate.ToString("dd-MMM-yyyy")
        Dim formattedEndDate As String = endDate.ToString("dd-MMM-yyyy")
        ASCMAIN1.sql = $"SELECT COUNT(*) FROM SOTFORM1 WHERE ORDR_FORM_STATUS = 'A' AND " &
                   $"((ORDR_FORM_VALID_FROM <= TO_DATE('{formattedEndDate}', 'DD-MON-YYYY') AND ORDR_FORM_VALID_UNTIL >= TO_DATE('{formattedStartDate}', 'DD-MON-YYYY')))"
        Dim count As Integer = Convert.ToInt32(ASCDATA1.GetDataValue())
        Return count > 0
    End Function
    Private Function GetActiveUsersToEmail() As DataTable
        Dim ORDR_FORM_CODE As String = txt_ORDR_FORM_CODE.Text

        ASCMAIN1.sql = $"
        SELECT u.ORDR_FORM_USER_NO, u.ORDR_FORM_USER_NAME, u.ORDR_FORM_USER_EMAIL
        FROM SOTFORMU u
        LEFT JOIN SOTRECEM r ON u.ORDR_FORM_USER_NO = r.ORDR_FORM_USER_NO
            AND r.ORDR_FORM_CODE = '{ORDR_FORM_CODE}'
        WHERE u.ORDR_FORM_USER_STATUS = 'A'
            AND (r.ORDR_FORM_USER_NO IS NULL OR u.ORDR_FORM_USER_EMAIL <> r.ORDR_FORM_USER_EMAIL)"

        Return ASCDATA1.GetDataTable()
    End Function

    Private Sub LogEmailRecipient(ORDR_FORM_USER_NO As String, ORDR_FORM_USER_NAME As String, ORDR_FORM_USER_EMAIL As String, ORDR_FORM_USER_STATUS As String)
        Dim ORDR_FORM_CODE As String = Absx1.txtFor("ORDR_FORM_CODE").Text

        ' Check if the user already exists in SOTRECEM
        Dim existingRow As DataRow = dst.Tables("SOTRECEM").Select($"ORDR_FORM_USER_NO = '{ORDR_FORM_USER_NO}'").FirstOrDefault()

        If existingRow IsNot Nothing Then
            ' User exists, so update the existing row
            existingRow("ORDR_FORM_USER_NAME") = ORDR_FORM_USER_NAME
            existingRow("ORDR_FORM_USER_EMAIL") = ORDR_FORM_USER_EMAIL
            existingRow("ORDR_FORM_USER_STATUS") = ORDR_FORM_USER_STATUS
            existingRow("ORDR_FORM_CODE") = ORDR_FORM_CODE
            existingRow("INIT_DATE") = DateTime.Now
            existingRow("INIT_OPER") = ASCMAIN1.USER_ID

            ASCMAIN1.sql = $"
            UPDATE SOTRECEM 
            SET ORDR_FORM_USER_NAME = '{ORDR_FORM_USER_NAME}', 
            ORDR_FORM_USER_EMAIL = '{ORDR_FORM_USER_EMAIL}', 
            ORDR_FORM_USER_STATUS = '{ORDR_FORM_USER_STATUS}', 
            ORDR_FORM_CODE = '{ORDR_FORM_CODE}',
            INIT_DATE = SYSDATE, 
            INIT_OPER = '{ASCMAIN1.USER_ID}'
            WHERE ORDR_FORM_USER_NO = '{ORDR_FORM_USER_NO}'
            AND ORDR_FORM_CODE = '{ORDR_FORM_CODE}'"
            ASCDATA1.ExecuteSQL()
        Else
            ' User does not exist, so insert a new row
            Dim newRow As DataRow = dst.Tables("SOTRECEM").NewRow()
            newRow("ORDR_FORM_USER_NO") = ORDR_FORM_USER_NO
            newRow("ORDR_FORM_USER_NAME") = ORDR_FORM_USER_NAME
            newRow("ORDR_FORM_USER_EMAIL") = ORDR_FORM_USER_EMAIL
            newRow("ORDR_FORM_USER_STATUS") = ORDR_FORM_USER_STATUS
            newRow("ORDR_FORM_CODE") = ORDR_FORM_CODE
            newRow("INIT_DATE") = DateTime.Now
            newRow("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables("SOTRECEM").Rows.Add(newRow)

            ASCMAIN1.sql = $"
        INSERT INTO SOTRECEM (ORDR_FORM_USER_NO, ORDR_FORM_USER_NAME, ORDR_FORM_USER_EMAIL, ORDR_FORM_USER_STATUS, ORDR_FORM_CODE, INIT_DATE, INIT_OPER) 
        VALUES ('{ORDR_FORM_USER_NO}', '{ORDR_FORM_USER_NAME}', '{ORDR_FORM_USER_EMAIL}', '{ORDR_FORM_USER_STATUS}', '{ORDR_FORM_CODE}', SYSDATE, '{ASCMAIN1.USER_ID}')"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub
    Private Sub UpdateUsersAddedLabel()
        Dim activeUserCount As Integer = 0
        Dim ORDR_FORM_CODE As String = txt_ORDR_FORM_CODE.Text
        For Each userRow As DataRow In dst.Tables("SOTFORMU").Rows
            Dim ORDR_FORM_USER_NO As String = userRow("ORDR_FORM_USER_NO").ToString()
            Dim ORDR_FORM_USER_EMAIL As String = userRow("ORDR_FORM_USER_EMAIL").ToString()
            If userRow("ORDR_FORM_USER_STATUS").ToString() = "A" Then
                ' Check if the user already exists in SOTRECEM for the specific form code
                Dim existingRows = dst.Tables("SOTRECEM").Select($"ORDR_FORM_USER_NO = '{ORDR_FORM_USER_NO}' AND ORDR_FORM_CODE = '{ORDR_FORM_CODE}'")

                If existingRows.Length = 0 Then
                    ' User has not received an email for this form code
                    activeUserCount += 1
                Else
                    ' Check if the email has changed, if so, they need another email
                    Dim existingEmail As String = existingRows(0)("ORDR_FORM_USER_EMAIL").ToString()
                    If ORDR_FORM_USER_EMAIL <> existingEmail Then
                        activeUserCount += 1
                    End If
                End If
            End If
        Next
        lblNewUsers.Text = "Active users added since last email: " & activeUserCount

        grdSOTRECEM.Text = "Users Who Received Email for Form Code " & ORDR_FORM_CODE
    End Sub

    Private Sub btnResetUsers_Click(sender As Object, e As EventArgs) Handles btnResetUsers.Click
        Dim confirmResult As DialogResult = MessageBox.Show(
        "You are about to permanently remove all users from the list of Authorized Users. " &
        "This action cannot be undone. Are you sure you want to proceed?",
        "Confirm Reset Users",
        MessageBoxButtons.YesNo,
        MessageBoxIcon.Warning
    )

        If confirmResult <> DialogResult.Yes Then Exit Sub

        Try
            ASCMAIN1.sql = "TRUNCATE TABLE SOTFORMU"
            ASCDATA1.ExecuteSQL()

            dst.Tables("SOTFORMU").Rows.Clear()

            lblNewUsers.Text = "Active users added since last email: 0"

            MessageBox.Show("All authorized users have been permanently removed.",
                            "Reset Complete",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("An error occurred while resetting users: " & ex.Message,
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error)
        End Try
    End Sub
End Class