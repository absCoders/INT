Public Class ICTARTI1

    Dim sqlICTCATA2 As String = ""
    Dim sqlPOTORDR2 As String = ""

    Dim SO_PARM_UPC_VENDOR_ID As String = ""
    Dim rowSOTPARM1 As DataRow

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")
        rowSOTPARM1 = LookUp("SOTPARM1", "Z")
        SO_PARM_UPC_VENDOR_ID = rowSOTPARM1.Item("SO_PARM_UPC_VENDOR_ID") & ""

        AUDIT.Add("ICTCATA1", "E")

        With dst
            sqlICTCATA2 = "Select ICTCATA2.*" & vbCrLf _
                & ", ICTARTI1.ARTICLE_DESC, ICTARTI1.BRAND_CODE, ICTARTI1.COLLECTION_CODE" & vbCrLf _
                & ", ICTARTI1.CATEGORY_CODE, ICTARTI1.ARTICLE_BASIC_PROMO, ICTARTI1.ARTICLE_CASE_PACK_QTY" _
                & " from ICTARTI1,ICTCATA2" & vbCrLf _
                & " where ICTARTI1.ARTICLE_CODE (+) = ICTCATA2.ARTICLE_CODE"
            ASCMAIN1.sql = sqlICTCATA2 _
            & "  and ICTCATA2.CATALOG_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "ICTCATA2", "**", 0, True, "V")

            'With .Tables("ICTCATA2")
            '    .Columns.Add("FUTURE_COST", GetType(System.Decimal))
            '    .Columns.Add("FUTURE_COST_DATE", GetType(System.DateTime))
            'End With


            sqlPOTORDR2 = "Select POTORDR2.*, POTORDR1.VEND_CODE" & vbCrLf _
                & " from POTORDR2,POTORDR1" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO"
            ASCMAIN1.sql = sqlPOTORDR2 _
            & "  and POTORDR2.ITEM_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, True, "V")


            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, True, "V")
        End With

        grdICTCATA2.DataSource = dst.Tables("ICTCATA2")

        With grdICTCATA2.DisplayLayout.Bands(0)
            .Columns("ARTICLE_CODE").Header.Fixed = True
            .Columns("ARTICLE_DESC").Header.Fixed = True
            .Columns("ARTICLE_COST").MaskInput = "nnnn.nnnnnn"
            .Columns("ARTICLE_COST").Format = "##0.000000"
        End With

        grdPOTORDR2.DataSource = dst.Tables("POTORDR2")


        Create_Summary(grdICTCATA2, "ARTICLE_CODE", "Count")
        Bind_Controls(grp3PL, "ICTITEM1")
        Bind_Controls(grpDimensions, "ICTITEM1")
        Bind_Controls(grpSales, "ICTITEM1")

    End Sub

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTCATA2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Add Codes")
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
                    Add_Codes(grdICTCATA2, "ICTARTI1", "ARTICLE_CODE", "Vendors")
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
        Sort_grdColumns(grdICTCATA2, "ARTICLE_CODE")

        EnforceConstraints(True)

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

#Region "grdICTCATA2"

    Private Sub grdICTCATA2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCATA2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ARTICLE_CODE"
                grdCodeDesc(grdICTCATA2, "ICTARTI1", "ARTICLE_CODE", "ARTICLE_DESC")
        End Select
    End Sub

    Private Sub grdICTCATA2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTCATA2.BeforeRowUpdate

        Dim row As DataRow = LookUp("ICTARTI1", e.Row.Cells("ARTICLE_CODE").Text)

        If row Is Nothing Then
            e.Cancel = True
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("CATALOG_CODE").Value = Absx1.txtFor("CATALOG_CODE").Text
        End If

    End Sub

    Private Sub grdICTCATA2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTCATA2.ClickCellButton
        Select Case e.Cell.Column.Key
            Case "ARTICLE_CODE"
                Dim sql_where As String = Get_List_of_Codes("ICTARTI1.ARTICLE_CODE not in", "ICTCATA2", "ARTICLE_CODE")
                grdClickCellButton(grdICTCATA2, sql_where, True)
        End Select

    End Sub

    Private Sub grdICTCATA2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTCATA2.InitializeRow
        If e.Row.Cells("CATALOG_CODE").Text <> "" And e.Row.Cells("CATALOG_CODE").Text <> Absx1.txtFor("CATALOG_CODE").Text Then
            e.Row.Cells("CATALOG_CODE").Appearance.ForeColor = Drawing.Color.Red
        End If
        grd_RowColor(dst.Tables("ICTCATA2"), e.Row)
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
End Class