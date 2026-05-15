Public Class ICTCATA1

    Dim sqlICTCATA2 As String = ""

    Dim SO_PARM_UPC_VENDOR_ID As String = ""
    Dim rowSOTPARM1 As DataRow

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")
        rowSOTPARM1 = LookUp("SOTPARM1", "Z")
        SO_PARM_UPC_VENDOR_ID = rowSOTPARM1.Item("SO_PARM_UPC_VENDOR_ID") & ""

        AUDIT.Add("ICTCATA1", "E")

        With dst
            sqlICTCATA2 = "Select ICTCATA2.*" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTCOLL1.BRAND_CODE, ICTITEM1.COLLECTION_CODE, ICTITEM1.VEND_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.CARTON_PACK_QTY" _
                & " from ICTITEM1,ICTCOLL1,ICTCATA2" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = ICTCATA2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
            ASCMAIN1.sql = sqlICTCATA2 _
            & "  and ICTCATA2.CATALOG_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTCATA2", "**", 0, True, "V")

            With .Tables("ICTCATA2")
                .Columns.Add("QTY_ON_HAND", GetType(System.Int64))
                .Columns.Add("QTY_OPEN", GetType(System.Int64))
                .Columns.Add("QTY_ONPO", GetType(System.Int64))
                .Columns.Add("QTY_NET", GetType(System.Int64), "ISNULL(QTY_ON_HAND,0)-ISNULL(QTY_OPEN,0)+ISNULL(QTY_ONPO,0)")
            End With

            Create_TDA(.Tables.Add, "POTORDR1", "*")
            Create_TDA(.Tables.Add, "POTORDR2", "*")
        End With

        grdICTCATA2.DataSource = dst.Tables("ICTCATA2")

        With grdICTCATA2.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
            '.Columns("ITEM_COST_STD").MaskInput = "nnnn.nnnnnn"
            '.Columns("ITEM_COST_STD").Format = "##0.000000"
        End With

        Create_Summary(grdICTCATA2, "ITEM_CODE", "Count")


        ASCMAIN1.Add_Value_List(Absx1.cbeFor("CATALOG_TYPE"), "ICTCATA1.CATALOG_TYPE")

        ASCMAIN1.Add_Value_List(grdICTCATA2, "ITEM_BASIC_PROMO")
        ASCMAIN1.Add_Value_List(grdICTCATA2, "ITEM_CATGY_CODE")
        'ASCMAIN1.Add_Value_List(grdICTCATA2, "ITEM_STATUS")
    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTCATA2, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes", "Generate POs")
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
            Case "grdICTCATA2"
                tlb_btn = DirectCast(tlb_pop.Tools("Add Codes"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

                tlb_btn = DirectCast(tlb_pop.Tools("Generate POs"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow.IsDataRow AndAlso Not grd.ActiveRow.IsAddRow
                If grd.ActiveRow.IsDataRow AndAlso Not grd.ActiveRow.IsAddRow Then
                    Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Value
                    If VEND_CODE = "" Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Caption = "Generate PO with " & VEND_CODE
                    End If
                End If


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
                If grd.Name = "grdICTCATA2" Then
                    Add_Codes(grdICTCATA2, "ICTITEM1", "ITEM_CODE", "Items")
                End If

            Case "Generate UPCs"
                Generate_UPCs()
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Allow Edit to this UPC"
            '    With grdICTSTYC1.DisplayLayout.Bands(0).Columns("UPC_CODE")
            '        .CellActivation = UltraWinGrid.Activation.AllowEdit
            '    End With

            Case "Generate POs"
                Dim VEND_CODE As String = grd.ActiveRow.Cells("VEND_CODE").Value
                If MsgBox("OK to Generate PO with " & VEND_CODE & " to cover all Items with Negative Net Position?", _
                          MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Generate_PO(VEND_CODE)
                End If

        End Select
    End Sub
#End Region

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "New"
                If ASCMAIN1.CLIENT = "NYA" Then
                    'If txtCATALOG_CODE_PLM_SOURCE.Tag & "" = "" Then
                    '    EMsg &= vbCr & "You Must Use the Create Style from PLM function to add a New Style"
                    'End If
                End If
            Case "Edit"

            Case "View"

            Case "Update"

                If Absx1.txtFor("CATALOG_NAME").Text = "" Then
                    EMsg &= vbCr & "You Must Enter a Value for Catalog Name"
                End If
        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = "CATALOG_CODE = '" & Absx1.txtFor("CATALOG_CODE").Text & "'"
        'INIT_LAST("ICTSTYC1", True)

        'For Each rowICTCATA2 As DataRow In dst.Tables("ICTCATA2").Select("FUTURE_COST IS NOT NULL", "FUTURE_COST")
        '    rowICTCATA2.Item("PO_COST") = Val(rowICTCATA2.Item("NEW_PO_COST") & "")
        '    rowICTCATA2.Item("PO_COST_DATE") = rowICTCATA2.Item("NEW_PO_COST_DATE")

        '    rowICTCATA2.Item("NEW_PO_COST") = Val(rowICTCATA2.Item("FUTURE_COST") & "")
        '    rowICTCATA2.Item("NEW_PO_COST_DATE") = rowICTCATA2.Item("FUTURE_COST_DATE")
        'Next

        Update_Record_TDA("ICTCATA2", sqlDelete)

    End Sub

    Overrides Sub Proceed_Update_Special_Post()
        If EntryMode = "New" Then
            'ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "New Account Added", "M")
        ElseIf EntryMode = "Edit" Then
            'ARCMAIN1.Record_Customer_Event(Absx1.txtFor("CUST_CODE").Text, "Masterfile Updated", "M")

            ' we should do something about ICTUPCH1 -note that if a UPC is re-assigned, that this table is out of synch with ICTSTYC1, although we do use ICTSTYC1 for UPCs, and we do have the audit trail
        End If
    End Sub

    Overrides Sub Show_Record_Special()

        Dim CATALOG_CODE As String = Absx1.txtFor("CATALOG_CODE").Text

        EnforceConstraints(False)

        Fill_Records("ICTCATA2", New String() {CATALOG_CODE})
        Sort_grdColumns(grdICTCATA2, "ITEM_CODE")

        Refresh_Stats()

        EnforceConstraints(True)

        Dim FOLDER_NAME As String = ASCMAIN1.Folders("Images") & "\COLUMN_NAME\CATALOG_CODE\"
        Dim IMAGE_NAME As String = CATALOG_CODE & ".png"
        If My.Computer.FileSystem.FileExists(FOLDER_NAME & IMAGE_NAME) Then
            img.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , ) ' imgba)
            img.Visible = True
        End If

    End Sub


    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() { _
                "ICTCATA2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

        grdICTCATA2.Enabled = tf

    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        If Not tf Then
            img.Visible = False
        End If


        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTCATA2}
            With grd.DisplayLayout.Override
                If EntryMode = "New" Or EntryMode = "Edit" Then
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowDelete = DefaultableBoolean.True
                    .AllowUpdate = DefaultableBoolean.True
                Else
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End If
            End With
        Next

    End Sub

#End Region

    Sub Generate_PO(VEND_CODE)
        Dim ITEM_CODEs As New List(Of String)
        Dim PO_ORDER_NO As String = ""
        dst.Tables("POTORDR1").Rows.Clear()
        dst.Tables("POTORDR2").Rows.Clear()

        Dim CATALOG_DATE_DELIVERY As Date = Absx1.dteFor("CATALOG_DATE_DELIVERY").DateTime
        Dim PO_DATE_REQUIRED As Date = CATALOG_DATE_DELIVERY.AddDays(-15)
        Dim QTY As Int64 = 0
        Dim AMT As Decimal = 0

        For Each row As DataRow In dst.Tables("ICTCATA2").Select("VEND_CODE = '" & VEND_CODE & "' and QTY_NET < 0")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim QTY_NET As Int64 = Val(row.Item("QTY_NET") & "")
            Dim ITEM_COST As Decimal = Val(row.Item("ITEM_COST") & "")

            If PO_ORDER_NO = "" Then
                PO_ORDER_NO = Generate_PO_Header(VEND_CODE, PO_DATE_REQUIRED)
            End If
            ITEM_CODEs.Add(ITEM_CODE)
            Dim rowPOTORDR2 As DataRow = Generate_PO_Detail(PO_ORDER_NO, ITEM_CODE, QTY_NET, ITEM_COST, PO_DATE_REQUIRED)

            QTY += Val(rowPOTORDR2.Item("PO_QTY_ORD"))
            AMT += Val(rowPOTORDR2.Item("PO_QTY_ORD")) * Val(rowPOTORDR2.Item("PO_COST"))

        Next

        If ITEM_CODEs.Count = 0 Then
            MsgBox("No Items found with Negative Net Position", MsgBoxStyle.OkOnly, "Nothing Done")
        Else
            BeginTrans()
            Update_Record_TDA("POTORDR1")
            Update_Record_TDA("POTORDR2")
            TAC.POCMAIN1.ICTSTAT2_PO(1, PO_ORDER_NO)

            CommitTrans("PO " & PO_ORDER_NO & " has been Generated with " & CStr(ITEM_CODEs.Count) & " Items" & vbCrLf & "Total Qty = " & CStr(QTY) & vbCrLf & "Total Purchase = " & Format(AMT, "$#,##0.00"))
        End If

        Print_PO(PO_ORDER_NO)
        Refresh_Stats()
    End Sub
    Function Print_PO(PO_ORDER_NO As String, Optional make_pdf As Boolean = False, Optional FILENAME_body As String = "") As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing POs")

        Dim REPORTFILE As String = "POROPRT1"
        Dim RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
        If RPT = "" Then RPT = REPORTFILE

        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and PO_ORDER_NO = '" & PO_ORDER_NO & "'"})

        Dim REPORT_NO As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("FORM_TYPE", "P")
            If make_pdf Then
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , "PDF", FILENAME_body, False)
            Else
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , , , False)
            End If
            .Print_Report_End(make_pdf, make_pdf)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return REPORT_NO
    End Function
 
    Function Generate_PO_Header(VEND_CODE As String, PO_DATE_REQUIRED As Date) As String
        Dim PO_ORDER_NO As String = ASCMAIN1.Next_Control_No("POTORDR1.PO_ORDER_NO", 1)
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
        Dim WHSE_CODE As String = "ADC"
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").NewRow
        With rowPOTORDR1
            .Item("PO_ORDER_NO") = PO_ORDER_NO
            .Item("VEND_CODE") = VEND_CODE

            For Each C As String In New String() {"VEND_NAME", "VEND_ADDR1", "VEND_ADDR2", "VEND_ADDR3", "VEND_CITY", "VEND_STATE", "VEND_ZIP_CODE", _
                                                  "VEND_COUNTRY", "VEND_PHONE", "VEND_EXT", "VEND_FAX", "VEND_EMAIL", "VEND_CONTACT", _
                                                  "TERM_CODE", "VEND_WHSE_CODE", "VEND_BUYER_CODE"}
                rowPOTORDR1.Item(C) = rowAPTVEND1.Item(C)
            Next
            For Each C As String In New String() {"SHIP_TO_NAME", "SHIP_TO_ADDR1", "SHIP_TO_ADDR2", "SHIP_TO_ADDR3", "SHIP_TO_CITY", "SHIP_TO_STATE", "SHIP_TO_ZIP_CODE", _
                                                 "SHIP_TO_COUNTRY", "SHIP_TO_PHONE", "SHIP_TO_EXT", "SHIP_TO_FAX", "SHIP_TO_EMAIL", "SHIP_TO_CONTACT"}
                Dim C_WHSE As String = Replace(C, "SHIP_TO_", "WHSE_")
                If C = "SHIP_TO_NAME" Then C_WHSE = "WHSE_DESC"
                rowPOTORDR1.Item(C) = rowICTWHSE1.Item(C_WHSE)
            Next

            .Item("PO_DATE_ORDERED") = DATETIME_STAMP.Date
            .Item("PO_DATE_REQUIRED") = PO_DATE_REQUIRED
            .Item("PO_DATE_CANCEL") = PO_DATE_REQUIRED.AddDays(10)

            .Item("PO_TYPE") = "B"
            If .Item("PO_TYPE") = "B" Then
                .Item("VEND_WHSE_CODE") = ""
            End If

            .Item("PO_CONTACT") = rowAPTVEND1.Item("VEND_PURCH_CONTACT")
            .Item("PO_FOB_DESC") = rowAPTVEND1.Item("VEND_PURCH_FOB_DESC")
            .Item("PO_SHIP_VIA") = rowAPTVEND1.Item("VEND_PURCH_SHIP_VIA")
            .Item("PO_ORDR_NOTES_INTERNAL") = rowAPTVEND1.Item("VEND_PURCH_COMMENT")

            If .Item("PO_FOB_DESC") & "" = "" Then rowPOTORDR1.Item("PO_FOB_DESC") = ROWs("POTPARM1").Item("PO_PARM_FOB")
            If .Item("PO_SHIP_VIA") & "" = "" Then rowPOTORDR1.Item("PO_SHIP_VIA") = ROWs("POTPARM1").Item("PO_PARM_SHIP_VIA")

            .Item("PO_REFERENCE") = Absx1.txtFor("CATALOG_CODE").Text
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("MARKET_CODE") = "DPT"
            .Item("PO_ORDR_NOTES_INTERNAL") = ""
            .Item("PO_ORDR_NOTES_EXTERNAL") = Absx1.txtFor("CATALOG_NAME").Text
            .Item("PO_SHIP_TO_REL") = ""
            .Item("WHSE_CODE") = "ADC"

            .Item("PO_STATUS") = "O"
            .Item("PO_ORDER_TYPE") = "P"
            .Item("MARKET_CODE") = "DPT"
        End With

        dst.Tables("POTORDR1").Rows.Add(rowPOTORDR1)

        Return PO_ORDER_NO
    End Function

    Function Generate_PO_Detail(PO_ORDER_NO As String, ITEM_CODE As String, QTY_NET As Int64, ITEM_COST As Decimal, PO_DATE_REQUIRED As Date) As DataRow
        Dim PO_ORDER_LNO As Integer = dst.Tables("POTORDR2").Rows.Count + 1
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        Dim ITEM_PO_QTY_MIN As Int64 = Val(rowICTITEM1.Item("ITEM_PO_QTY_MIN") & "")
        Dim ITEM_PO_QTY_MULT As Int64 = Val(rowICTITEM1.Item("ITEM_PO_QTY_MULT") & "")
        Dim PO_QTY As Int64 = -1 * QTY_NET

        If PO_QTY < ITEM_PO_QTY_MIN Then
            PO_QTY = ITEM_PO_QTY_MIN
        End If

        If ITEM_PO_QTY_MULT > 0 AndAlso PO_QTY Mod ITEM_PO_QTY_MULT <> 0 Then
            PO_QTY += ITEM_PO_QTY_MULT - (PO_QTY Mod ITEM_PO_QTY_MULT)
        End If

        Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").NewRow
        With rowPOTORDR2
            .Item("PO_ORDER_NO") = PO_ORDER_NO
            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
            .Item("ITEM_CODE") = ITEM_CODE
            .Item("PO_QTY_ORD") = -1 * QTY_NET

            .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            .Item("ITEM_UOM") = rowICTITEM1.Item("ITEM_UOM")
            .Item("ITEM_PCT_ALLOW_OVER") = rowICTITEM1.Item("ITEM_PCT_ALLOW_OVER")
            .Item("ITEM_PCT_ALLOW_UNDER") = rowICTITEM1.Item("ITEM_PCT_ALLOW_UNDER")

            .Item("PO_COST") = ITEM_COST

            .Item("PO_QTY_ORD") = PO_QTY
            .Item("PO_QTY_OPN") = PO_QTY

            .Item("PO_ITEM_NOTE") = Absx1.txtFor("CATALOG_NAME").Text
            .Item("PO_DATE_REQUIRED") = PO_DATE_REQUIRED
            .Item("WHSE_CODE") = "ADC"
            .Item("PO_STATUS") = "O"

        End With
        dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)
        Return rowPOTORDR2
    End Function

#Region "grdICTCATA2"

    Private Sub grdICTCATA2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCATA2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                grdCodeDesc(grdICTCATA2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
        End Select
    End Sub

    Private Sub grdICTCATA2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTCATA2.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CATALOG_CODE").Value = Absx1.txtFor("CATALOG_CODE").Text
        End If

    End Sub

    Private Sub grdICTCATA2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCATA2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTITEM1.ITEM_CODE not in", "ICTCATA2", "ITEM_CODE")
                grdClickCellButton(grdICTCATA2, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTCATA2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCATA2.InitializeRow
        If e.Row.Cells("CATALOG_CODE").Text <> "" And e.Row.Cells("CATALOG_CODE").Text <> Absx1.txtFor("CATALOG_CODE").Text Then
            e.Row.Cells("CATALOG_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ICTCATA2"), e.Row)

        If e.Row.IsDataRow AndAlso Not e.Row.IsAddRow Then
            Dim QTY_NET As Int64 = Val(e.Row.Cells("QTY_NET").Value & "")
            If QTY_NET < 0 Then
                e.Row.Cells("QTY_NET").Appearance.ForeColor = Color.Red
            Else
                e.Row.Cells("QTY_NET").Appearance.ForeColor = Color.Empty
            End If
        End If
    End Sub

#End Region

    Function Get_UPC_Code(CATALOG_CODE As String, COLOR_CODE As String) As String

        Dim UPC_CODE As String = ""
        Do
            Dim UPC_CODE_CTL_NO As String = ASCMAIN1.Next_Control_No("ICTUPCH1.UPC_CODE")

            UPC_CODE = TAC.SOCMAIN1.UPC(Me, UPC_CODE_CTL_NO, SO_PARM_UPC_VENDOR_ID, True)
            If LookUp("ICTUPCH1", UPC_CODE) Is Nothing Then Exit Do
        Loop

        ASCMAIN1.sql = "Insert into ICTUPCH1 (UPC_CODE,CATALOG_CODE,COLOR_CODE,INIT_DATE,INIT_OPER) " & vbCrLf _
            & " values (:PARM1,:PARM2,:PARM3,SYSDATE,:PARM4)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {UPC_CODE, CATALOG_CODE, COLOR_CODE, ASCMAIN1.USER_ID})

        Return UPC_CODE
    End Function

    Sub Generate_UPCs()
        'For Each rowICTSTYC1 As DataRow In dst.Tables("ICTSTYC1").Select("UPC_CODE IS NULL")
        '    rowICTSTYC1.Item("UPC_CODE") = Get_UPC_Code(rowICTSTYC1.Item("CATALOG_CODE"), rowICTSTYC1.Item("COLOR_CODE"))
        'Next
    End Sub

    Sub Get_Stats(ITEM_CODE As String)
        Dim CATALOG_CODE As String = Absx1.txtFor("CATALOG_CODE").Text
        Dim rowICTCATA2 As DataRow = dst.Tables("ICTCATA2").Rows.Find(New Object() {CATALOG_CODE, ITEM_CODE})
        ASCMAIN1.sql = "Select * from ICTSTAT2 where ITEM_CODE = '" & ITEM_CODE & "'"
        Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow()
        rowICTCATA2.Item("QTY_ON_HAND") = rowICTSTAT2.Item("WHSE_QTY_ON_HAND")
        rowICTCATA2.Item("QTY_OPEN") = rowICTSTAT2.Item("WHSE_QTY_OPEN")
        rowICTCATA2.Item("QTY_ONPO") = rowICTSTAT2.Item("WHSE_QTY_ONPO")
    End Sub

    Sub Refresh_Stats()
        For Each ROW As DataRow In dst.Tables("ICTCATA2").Select("")
            Dim ITEM_CODE As String = ROW.Item("ITEM_CODE")
            Get_Stats(ITEM_CODE)
        Next
    End Sub
End Class