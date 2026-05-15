Public Class SOTPRIC1

    Dim sqlSOTPRIC2 As String = ""
    Dim isCurrRetailPriceList As Boolean = False
    Dim decimalsOK As New Infragistics.Win.Appearance
    Dim decimalsConcealed As New Infragistics.Win.Appearance

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' ISSUE-7242  Price List - View only
        InquiryMode = MENU_ITEM_OBJECT = "SOTPRICI"

        Get_PARM("GLTPARM1")
        Get_PARM("SOTPARM1")

        'Dim test As String
        'Dim x As New tester

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
        End With

        grdSOTPRIC2.DataSource = dst.Tables("SOTPRIC2")

        grdSOTPRIC2.DisplayLayout.UseFixedHeaders = True
        With grdSOTPRIC2.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSOTPRIC2.DisplayLayout.Bands(0).Columns
            If New String() {"ITEM_CODE", "ITEM_PRICE", "ITEM_NEW_PRICE", "ITEM_NEW_PRICE_DATE", "ITEM_SRP", "ITEM_NEW_SRP", "ITEM_NEW_SRP_DATE"}.Contains(gcol.Key) Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellAppearance.BackColor = Drawing.Color.LightGray
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        Create_Summary(grdSOTPRIC2, "ITEM_CODE", "Count")

        With grdSOTPRIC2.DisplayLayout.Bands(0)
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
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPRIC2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")

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
        'if not new or edit - hide add codes

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

            Case "grdSOTPRIC2"
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
                If grd.Name = "grdSOTPRIC2" Then
                    Add_Codes(grdSOTPRIC2, "ICTITEM1", "ITEM_CODE", "Items")
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
                    For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Select("")
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

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()

        Dim ALERT_MESSAGE As String = ""
        Dim PRICE_LIST_CODE As String = Absx1.txtFor("PRICE_LIST_CODE").Text
        Dim PRICE_LIST_DESC As String = Absx1.txtFor("PRICE_LIST_DESC").Text
        Dim CHANGE_LOGS As New List(Of String)
        Dim DELETED_ITEMS As New List(Of String)

        For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Rows

            If rowSOTPRIC2.RowState = DataRowState.Deleted And Not (ASCMAIN1.CLIENT = "AHA") Then
                ' DELETED
                Dim ITEM_CODE As String = rowSOTPRIC2.Item("ITEM_CODE", DataRowVersion.Original)
                Dim ITEM_PRICE As Decimal = Val(rowSOTPRIC2.Item("ITEM_PRICE", DataRowVersion.Original) & "")
                Dim rowICTITEM1_STD As DataRow = LookUp("ICTITEM1", ITEM_CODE)

                ALERT_MESSAGE &= vbCrLf _
                    & vbCrLf & " Item Price Change Deleted Record" _
                    & " Item " & ITEM_CODE & ":" & rowICTITEM1_STD.Item("ITEM_DESC") _
                    & " - Item Price: " & Format(ITEM_PRICE, "$#.00")

                Dim DELETED_LOG As String = Build_Deleted_Log(rowSOTPRIC2)
                If DELETED_LOG & "" <> "" Then DELETED_ITEMS.Add(DELETED_LOG)
                Write_Audit_Trail(rowSOTPRIC2, "E")

            Else

                Dim ITEM_CODE As String = rowSOTPRIC2.Item("ITEM_CODE")
                Dim ITEM_PRICE As Decimal = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                Dim ITEM_NEW_PRICE As Decimal = Val(rowSOTPRIC2.Item("ITEM_NEW_PRICE") & "")
                Dim rowICTITEM1_STD As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                Dim ITEM_NEW_PRICE_DATE = rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") & ""

                If rowSOTPRIC2.RowState = DataRowState.Modified And Not (ASCMAIN1.CLIENT = "AHA") Then
                    ' CHANGED
                    Dim ORIG_ITEM_PRICE As Decimal = Val(rowSOTPRIC2.Item("ITEM_PRICE", DataRowVersion.Original) & "")

                    If ITEM_PRICE <> ORIG_ITEM_PRICE Then

                        ALERT_MESSAGE &= vbCrLf _
                        & vbCrLf & "Item Price Change Current " _
                        & " Item: " & ITEM_CODE & ":" & rowICTITEM1_STD.Item("ITEM_DESC") _
                        & " - Item Price: " & Format(ORIG_ITEM_PRICE, "$#.00") _
                        & " Item New Price @" & Format(ITEM_PRICE, "$#.00")
                        Write_Audit_Trail(rowSOTPRIC2, "E")
                    End If


                    If ITEM_PRICE <> ITEM_NEW_PRICE And ITEM_NEW_PRICE <> 0 Then

                        ALERT_MESSAGE &= vbCrLf _
                        & vbCrLf & "Item Price Change Future " _
                        & " Item: " & ITEM_CODE & ":" & rowICTITEM1_STD.Item("ITEM_DESC") _
                        & " - Item Price: " & Format(ITEM_PRICE, "$#.00") _
                        & " Item New Price @" & Format(ITEM_NEW_PRICE, "$#.00") _
                        & " To change on " & ITEM_NEW_PRICE_DATE
                        Write_Audit_Trail(rowSOTPRIC2, "E")
                    End If
                End If
                Dim log As String = Build_Wholesale_Change_Log(rowSOTPRIC2)
                If log <> "" Then CHANGE_LOGS.Add(log)
            End If
        Next

        email_Wholesale_Price_Grouped(PRICE_LIST_DESC, PRICE_LIST_CODE, CHANGE_LOGS, DELETED_ITEMS)

        If ALERT_MESSAGE <> "" And Not (ASCMAIN1.CLIENT = "AHA") Then

            Dim ALERT_SUBJECT As String = ""
            Dim ALERT_EMAIL = ROWs("SOTPARM1").Item("SO_PARM_EML_PRICE") & ""
            Dim rowTATALRT1 As DataRow = dst.Tables("TATALRT1").NewRow
            With rowTATALRT1
                Dim ALERT_NO As String = ASCMAIN1.Next_Control_No("TATALRT1.ALERT_NO")
                .Item("ALERT_NO") = ALERT_NO
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("FORM_NAME") = "SOTPRIC1"
                .Item("FORM_KEY") = ALERT_NO
                .Item("ALERT_EMAIL") = ALERT_EMAIL
                .Item("ALERT_EML") = "1"

                .Item("ALERT_EML_DATE") = DATETIME_STAMP
                ALERT_SUBJECT = "Price Exception " & "" & " PC: " & PRICE_LIST_CODE
                .Item("ALERT_SUBJECT") = Mid(ALERT_SUBJECT, 1, 200)
                ALERT_MESSAGE = "Control No: " & ALERT_NO & vbCrLf & "Price Change: " & PRICE_LIST_CODE & ALERT_MESSAGE
                .Item("ALERT_MESSAGE") = Mid(ALERT_MESSAGE, 1, 2000)
            End With
            dst.Tables("TATALRT1").Rows.Add(rowTATALRT1)

            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            EMAIL_ADDRESSs.Add(ALERT_EMAIL, "Price Auditor")

            Dim SEND_NO As String = ""
            If ASCMAIN1.Running_in_VS Then
                SEND_NO = "TESTING"
                '   Stop
            Else
                SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing,
                    ALERT_SUBJECT, "PC_PRCEXC", True, False, PRICE_LIST_CODE, PRICE_LIST_CODE, "Price List Code", ALERT_MESSAGE)
            End If

            rowTATALRT1.Item("SEND_NO") = SEND_NO
            Update_Record_TDA("TATALRT1")

            TAC.TACMAIN1.Record_Event("SOTPRIC1", PRICE_LIST_CODE, DATETIME_STAMP, ASCMAIN1.USER_ID, "PRCEXC", "Price Exception Alert emailed", SEND_NO, "SOTPRIC1")
        End If



        Dim sqlDelete = "PRICE_LIST_CODE = '" & Absx1.txtFor("PRICE_LIST_CODE").Text & "'"
        Update_Record_TDA("SOTPRIC2", sqlDelete)
    End Sub

    Overrides Sub Show_Record_Special()

        Dim PRICE_LIST_CODE As String = Absx1.txtFor("PRICE_LIST_CODE").Text

        EnforceConstraints(False)
        Fill_Records("SOTPRIC2", New String() {PRICE_LIST_CODE})
        Sort_grdColumns(grdSOTPRIC2, "ITEM_CODE")
        grdSOTPRIC2.Text = "Price List Details for " & PRICE_LIST_CODE

        Dim need4places As Boolean = False
        For Each row As DataRow In dst.Tables("SOTPRIC2").Select
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
                For Each rowSOTPRIC2 As DataRow In dst.Tables("SOTPRIC2").Select("")
                    Dim ITEM_CODE = rowSOTPRIC2.Item("ITEM_CODE")
                    rowSOTPRIC2.Item("ITEM_RETAIL_PRICE") = Get_ITEM_RETAIL_PRICE_CURR(ITEM_CODE)
                Next
            End If
        End If

        For Each row As DataRow In dst.Tables("SOTPRIC2").Select()
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim I As Integer = 0
            ASCMAIN1.sql = "Select INIT_DATE, OLD_VALUE from ASTAUDT1 " _
                & " where TABLE_NAME = 'SOTPRIC2' " _
                & "   and KEY_VALUE = :PARM1 || ':' || :PARM2" _
                & "   and COLUMN_NAME = 'ITEM_PRICE'" _
                & "   and OLD_VALUE is not null" _
                & " order by INIT_DATE DESC"
            For Each rowASTAUDT1 As DataRow In _
                ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", _
                    New Object() {PRICE_LIST_CODE, ITEM_CODE}).Select("", "INIT_DATE DESC")
                I += 1
                If I <= 6 Then
                    Dim II As String = Format(I, "0")
                    row.Item("ITEM_PRICE_" & II) = rowASTAUDT1.Item("OLD_VALUE")
                    row.Item("ITEM_PRICE_" & II & "_DATE") = rowASTAUDT1.Item("INIT_DATE")
                End If
            Next
        Next

        With grdSOTPRIC2.DisplayLayout.Bands(0)
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
            dst.Tables("SOTPRIC2").Rows.Clear()
            EnforceConstraints(True)
        End If

        grdSOTPRIC2.Text = "Price List Details"

        chkShow4Decimals.Checked = False
        Toggle_Decimal_Places()

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        '  grdSOTPRIC2.Enabled = tf

        Absx1.txtFor("CURR_CODE").Enabled = (EntryMode = "New")

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTPRIC2}
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

    End Sub

#End Region

#Region "grdSOTPRIC2"
    Private Sub grdSOTPRIC2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPRIC2.AfterCellUpdate
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

    Private Sub grdSOTPRIC2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPRIC2.AfterRowActivate

        With grdSOTPRIC2.DisplayLayout.Bands(0).Columns("ITEM_CODE")
            If grdSOTPRIC2.ActiveRow.IsAddRow Then
                .CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

    End Sub

    Private Sub grdSOTPRIC2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTPRIC2.AfterRowsDeleted

    End Sub

    Private Sub grdSOTPRIC2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPRIC2.AfterRowUpdate

    End Sub

    Private Sub grdSOTPRIC2_BeforeExitEditMode(sender As Object, e As UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTPRIC2.BeforeExitEditMode
        If grdSOTPRIC2.ActiveCell IsNot Nothing Then
            With grdSOTPRIC2.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        If .EditorResolved.IsValid Then
                            .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTPRIC2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTPRIC2.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTPRIC2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPRIC2.BeforeRowUpdate
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

    Private Sub grdSOTPRIC2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPRIC2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTITEM1.ITEM_CODE not in", "SOTPRIC2", "ITEM_CODE")
                grdClickCellButton(grdSOTPRIC2, sql_where, True)
        End Select
    End Sub

    Private Sub grdSOTPRIC2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPRIC2.InitializeRow
        grd_RowColor(dst.Tables("SOTPRIC2"), e.Row)

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

        Sort_grdColumns(grdSOTPRIC2, "ITEM_CODE")

        Set_DX_Column(grdSOTPRIC2, "")
        For Each GCOL As UltraWinGrid.UltraGridColumn In grdSOTPRIC2.DisplayLayout.Bands(0).Columns
            Set_DX_Column(grdSOTPRIC2, GCOL.Key, GCOL.Header.Caption, GCOL.Width)
        Next

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook()
        Load_DataTable_into_SGXLS(1, 1, dst.Tables("SOTPRIC2"), workbook.ActiveWorksheet, grdSOTPRIC2, Nothing, "ITEM_CODE", "")
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
            With grdSOTPRIC2.DisplayLayout.Bands(0).Columns(C)
                If chkShow4Decimals.Checked Then
                    .Format = "$#.0000"
                    .MaskInput = "nnnnnn.nnnn"
                Else
                    .Format = "$#.00"
                    .MaskInput = "nnnnnn.nn"
                End If
            End With
        Next

        grdSOTPRIC2.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
    End Sub

    Sub email_Wholesale_Price_Grouped(PRICE_LIST_DESC As String, PRICE_LIST_CODE As String, CHANGE_LOGS As List(Of String), DELETED_ITEMS As List(Of String))
        If CHANGE_LOGS.Count = 0 AndAlso DELETED_ITEMS.Count = 0 AndAlso EntryMode <> "New" Then Exit Sub

        Dim userRow As DataRow = LookUp("ASTUSER1", ASCMAIN1.USER_ID)
        Dim USER_NAME As String = If(userRow IsNot Nothing, userRow("USER_NAME") & "", ASCMAIN1.USER_ID)

        Dim ALERT_SUBJECT As String = "Wholesale Price Change - " & PRICE_LIST_CODE & " - " & PRICE_LIST_DESC
        Dim rowTATALRT1 As DataRow = dst.Tables("TATALRT1").NewRow
        Dim ALERT_NO As String = ASCMAIN1.Next_Control_No("TATALRT1.ALERT_NO")

        ASCMAIN1.sql = $"SELECT CUST_CODE FROM ARTCUST1 WHERE PRICE_LIST_CODE = '{PRICE_LIST_CODE}' ORDER BY CUST_CODE"
        Dim dtCUST As DataTable = ASCDATA1.GetDataTable
        Dim CUST_LIST As String = ""
        Dim CUST_CODES As New List(Of String)
        If dtCUST.Rows.Count > 0 Then
            For Each r As DataRow In dtCUST.Rows
                CUST_CODES.Add(r("CUST_CODE") & "")
            Next
            CUST_LIST = "Used by customers: " & String.Join(", ", CUST_CODES) & vbCrLf & vbCrLf
        End If

        Dim ALERT_MESSAGE As String = ""
        If EntryMode = "New" Then
            ALERT_MESSAGE &= "New price list created by " & USER_NAME & ": " & PRICE_LIST_CODE & vbCrLf & vbCrLf
        Else
            ALERT_MESSAGE &= "Price list updated by " & USER_NAME & ": " & PRICE_LIST_CODE & vbCrLf & vbCrLf
        End If

        If CHANGE_LOGS.Count > 0 Then
            ALERT_MESSAGE &= String.Join(vbCrLf & String.Join("", Enumerable.Repeat("-", 40)) & vbCrLf, CHANGE_LOGS)
            ALERT_MESSAGE &= vbCrLf
        End If

        If DELETED_ITEMS.Count > 0 Then
            ALERT_MESSAGE &= "Deleted Items:" & vbCrLf
            ALERT_MESSAGE &= String.Join(vbCrLf, DELETED_ITEMS)
            ALERT_MESSAGE &= vbCrLf
        End If

        Dim htmlMessage As String = $"<h2>Wholesale Price Change - Price List: {PRICE_LIST_CODE} - {PRICE_LIST_DESC}</h2>"
        htmlMessage &= $"<div><b>Updated by:</b> {USER_NAME}</div><br/>"
        If CUST_CODES.Count > 0 Then
            htmlMessage &= $"<div><b>Used by customers:</b> {String.Join(", ", CUST_CODES)}</div><br/>"
        End If


        For Each log As String In CHANGE_LOGS
            htmlMessage &= $"<pre>{log}</pre><hr style='margin:12px 0'/>"
        Next

        If DELETED_ITEMS.Count > 0 Then
            htmlMessage &= "<h3>Deleted Items:</h3>"
            For Each item As String In DELETED_ITEMS
                htmlMessage &= $"<pre>{item}</pre>"
            Next
            htmlMessage &= "<hr style='margin:12px 0'/>"
        End If

        'htmlMessage &= $"<br/><div>Regards,<br/>{TAC.POCMAIN1.GetUserSignature(userRow)}</div>"

        With rowTATALRT1
            .Item("ALERT_NO") = ALERT_NO
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("FORM_NAME") = "SOTPRIC1"
            .Item("FORM_KEY") = ALERT_NO
            .Item("ALERT_EML") = "1"
            .Item("ALERT_EML_DATE") = DATETIME_STAMP
            .Item("ALERT_SUBJECT") = Mid(ALERT_SUBJECT, 1, 200)
            .Item("ALERT_MESSAGE") = Mid("Control No: " & ALERT_NO & vbCrLf & ALERT_MESSAGE, 1, 2000)
        End With
        dst.Tables("TATALRT1").Rows.Add(rowTATALRT1)

        Dim em As New ASCNOTEE(ASCMAIN1.Folders, "PLALERT", Nothing)
        Dim restrictedEmails As New HashSet(Of String) From
            {
                "pdesaulles@interparfums.com",
                "nferon@interparfums.com",
                "pcarraro@interparfums.com",
                "elepore@interparfums.com"
            }
        If Not restrictedEmails.Contains(ASCMAIN1.USER_EMAIL) Then
            em.AddEmailAddress(ASCMAIN1.USER_EMAIL, ASCNOTEE.emailAddresses.emailCC)
        End If

        em.CreateComponents()
        em.SetEmailSubject(ALERT_SUBJECT)
        em.SetDocumentBody(htmlMessage)
        'em.AddEmailAddress("latchananick@gmail.com", ASCNOTEE.emailAddresses.emailCC)
        'em.SetEmailTo("nicholas@absolution.com")

        em.EmailDocument()

        'Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        'EMAIL_ADDRESSs.Add("nicholas@absolution.com", "CEO")
        'EMAIL_ADDRESSs.Add("stephanie@absolution.com", "Sr Accountant")

        'Dim SEND_NO As String = ""
        'SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
        '(ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing,
        'ALERT_SUBJECT, "WHSALR", True, False, PRICE_LIST_CODE, PRICE_LIST_CODE, "Price List", ALERT_MESSAGE)


        rowTATALRT1.Item("SEND_NO") = ALERT_NO
        Update_Record_TDA("TATALRT1")
        TAC.TACMAIN1.Record_Event("SOTPRIC1", PRICE_LIST_CODE, DATETIME_STAMP, ASCMAIN1.USER_ID, "WHSALR", "Wholesale Price Change Alert emailed", ALERT_NO, "SOTPRIC1")
    End Sub


    Function Build_Wholesale_Change_Log(row As DataRow) As String
        Dim IS_NEW_ITEM As Boolean = (row.RowState = DataRowState.Added)

        Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
        Dim PRICE_LIST_CODE As String = row.Item("PRICE_LIST_CODE") & ""
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        If rowICTITEM1 Is Nothing Then Return ""

        Dim ITEM_DESC As String = rowICTITEM1.Item("ITEM_DESC") & ""
        Dim FIELDS As New List(Of String) From {
            "ITEM_PRICE", "ITEM_NEW_PRICE", "ITEM_NEW_PRICE_DATE"
        }

        Dim CHANGE_LOG As String = ""
        For Each FIELD As String In FIELDS
            Dim valCurrent As String = "", valOriginal As String = ""

            If row.RowState = DataRowState.Added Then
                If row.Table.Columns(FIELD).DataType Is GetType(Date) Then
                    If Not IsDBNull(row(FIELD)) Then valCurrent = Format(row(FIELD), "MM/dd/yyyy")
                Else
                    valCurrent = Format(Val(row(FIELD) & ""), "$#.00")
                End If
            Else
                If row.Table.Columns(FIELD).DataType Is GetType(Date) Then
                    If Not IsDBNull(row(FIELD)) Then valCurrent = Format(row(FIELD), "MM/dd/yyyy")
                    If Not IsDBNull(row(FIELD, DataRowVersion.Original)) Then valOriginal = Format(row(FIELD, DataRowVersion.Original), "MM/dd/yyyy")
                Else
                    valCurrent = Format(Val(row(FIELD) & ""), "$#.00")
                    valOriginal = Format(Val(row(FIELD, DataRowVersion.Original) & ""), "$#.00")
                End If
            End If

            If valCurrent <> valOriginal Then
                If String.IsNullOrWhiteSpace(valOriginal) Then
                    CHANGE_LOG &= "- " & Replace(FIELD, "_", " ") & ": Set to [" & valCurrent & "] (was blank)" & vbCrLf
                Else
                    CHANGE_LOG &= "- " & Replace(FIELD, "_", " ") & ": Changed from [" & valOriginal & "] to [" & valCurrent & "]" & vbCrLf
                End If
            End If
        Next

        If CHANGE_LOG = "" Then Return ""

        If IS_NEW_ITEM Then
            Return "Item: " & ITEM_CODE & " - " & ITEM_DESC & vbCrLf & "(New item added to price list)" & vbCrLf & CHANGE_LOG
        Else
            Return "Item: " & ITEM_CODE & " - " & ITEM_DESC & vbCrLf & CHANGE_LOG
        End If

    End Function
    Function Build_Deleted_Log(row As DataRow) As String
        Dim ITEM_CODE As String = row("ITEM_CODE", DataRowVersion.Original) & ""
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        Dim ITEM_DESC As String = If(rowICTITEM1 IsNot Nothing, rowICTITEM1("ITEM_DESC") & "", "")
        Return "Item: " & ITEM_CODE & " - " & ITEM_DESC & " was removed from price list."
    End Function




End Class