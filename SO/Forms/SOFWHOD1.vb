Public Class SOFWHOD1
    Dim COLLECTION_CODE As String
    Dim rowICTCOLL1 As DataRow
    Dim RYP0 As String
    Dim RYP1 As String

    Dim DTE0() As Date
    Dim DTE1() As Date

    Dim SOTWHODI As String
    Dim SOTWHODC As String
    Dim SOTWHOD1 As String
    Dim SOTWHOD2 As String
    Dim sqlSOTORDR2 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Property item As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Create_Temp_Tables()

            ASCMAIN1.sql = "Select SOTWHOD1.*,ARTCUST1.CUST_NAME" & vbCrLf _
                & " from " & SOTWHOD1 & " SOTWHOD1,ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SOTWHOD1.CUST_CODE"
            Create_TDA(.Tables.Add, "SOTWHOD1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTWHOD1.ITEM_CODE, SOTWHOD1.CUST_CODE, ARTCUST1.CUST_NAME, SOTWHOD1.ALLO_UNITS" & vbCrLf _
                & " from " & SOTWHOD1 & " SOTWHOD1,ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SOTWHOD1.CUST_CODE" & vbCrLf _
                & "  and NVL(ORDR_COUNT,0) = 0"
            Create_TDA(.Tables.Add, "SOTWHOD1_NOT", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTWHOD2.*,ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & " from " & SOTWHOD2 & " SOTWHOD2,ARTCUST2" & vbCrLf _
                & " where ARTCUST2.CUST_CODE = SOTWHOD2.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SOTWHOD2.CUST_STORE_NO"
            Create_TDA(.Tables.Add("SOTWHOD2"), SOTWHOD2, "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select SOTWHODI.*, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.PROD_CODE, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
            & " from ICTITEM1," & SOTWHODI & " SOTWHODI" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTWHODI.ITEM_CODE"
            Create_TDA(.Tables.Add("SOTWHODI"), SOTWHODI, "**", 0, True, "", 1)
            .Tables("SOTWHODI").Columns.Add("DID", GetType(System.Int64))
            .Tables("SOTWHODI").Columns.Add("DIDNOT", GetType(System.Int64))

            ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" _
                & " from SOTORDR2,SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)

            'ASCMAIN1.sql = "Select * from ARTREAS1"
            'Create_TDA(.Tables.Add, "ARTREAS1", "**", 0, False)
        End With

        'Fill_Records("ARTREAS1")

        grdSOTWHODI.DataSource = dst.Tables("SOTWHODI")
        grdSOTWHOD1.DataSource = dst.Tables("SOTWHOD1")
        grdSOTWHOD1_NOT.DataSource = dst.Tables("SOTWHOD1_NOT")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")

        Create_Summary(grdSOTWHODI, "ITEM_CODE", "Count")
        Create_Summary(grdSOTWHODI, "SEL")

        Create_Summary(grdSOTWHOD1, "CUST_CODE", "Count")
        'Create_Summary(grdSOTWHOD1, New String() {"ORDR_COUNT", "SHIP_COUNT", "PICK_COUNT", "OPEN_COUNT", _
        '                                           "ORDR_UNITS", "SHIP_UNITS", "PICK_UNITS", "OPEN_UNITS", _
        '                                           "ORDR_SALES", "SHIP_SALES", "PICK_SALES", "OPEN_SALES"})

        With grdSOTWHODI.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            'For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
            '    Dim COLUMN_NAME As String = gcol.Key
            '    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            '    If New String() {"EXT_NET"}.Contains(gcol.Key) Then
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
            '        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            '    Else
            '        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
            '        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            '    End If
            'Next
        End With

        With grdSOTWHOD1.DisplayLayout.Bands(0)
            For Each DS As String In New String() {"ORDR", "SHIP", "PICK", "OPEN", "ALLO"}
                For Each DT As String In New String() {"COUNT", "UNITS", "SALES"}
                    If DS = "ALLO" And DT <> "UNITS" Then
                    Else
                        Dim C As String = DS & "_" & DT
                        .Columns(C).Format = "#,##0"
                        .Columns(C).Width = 70
                        Create_Summary(grdSOTWHOD1, C)
                        If DT = "COUNT" Then
                            .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.Orange
                        ElseIf DT = "UNITS" Then
                            If DS = "ALLO" Then
                                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.Pink
                            Else
                                .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                            End If
                        Else
                            .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        End If
                    End If
                Next
            Next
        End With

        With grdSOTWHOD1_NOT.DisplayLayout.Bands(0)
            Dim C As String = "ALLO_UNITS"
            .Columns(C).Format = "#,##0"
            .Columns(C).Width = 70
            Create_Summary(grdSOTWHOD1_NOT, C)
            .Columns(C).Header.Appearance.BackColor2 = Drawing.Color.Pink
        End With

        ASCMAIN1.sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2" _
            & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        cbeYP0.DataSource = New DataView(tbl, "", "OPS_YYYYPP", DataViewRowState.CurrentRows)
        cbeYP1.DataSource = New DataView(tbl, "", "OPS_YYYYPP", DataViewRowState.CurrentRows)

        ASCMAIN1.Add_Value_List(grdSOTORDR2, "ORDR_STATUS", New String() {":", "O:Open", "P:Pick", "F:Shipped", "C:Cancel"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If Absx1.txtFor("COLLECTION_CODE").Text <> "" Then
                    Validate_Code("COLLECTION_CODE")
                    If EMsg = "" Then
                        COLLECTION_CODE = Absx1.txtFor("COLLECTION_CODE").Text
                        rowICTCOLL1 = LookUp("ICTCOLL1", COLLECTION_CODE)
                        If rowICTCOLL1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Collection Code " & COLLECTION_CODE
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If dst.Tables("SOTWHODI").Select("SEL='1'").Length = 0 Then
                        EMsg &= vbCr & "No Items Selected"
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                    .Items("Print").Visible = (InquiryMode Or EntryMode = "V")
                End With
                .Groups("Item Selection").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpOptions, ScreenMode)


        splSOTWHODI.Visible = ScreenMode
        grdSOTWHODI.DisplayLayout.Bands(0).Columns("SEL").Hidden = ScreenMode
        grdSOTWHODI.DisplayLayout.Bands(0).Columns("DID").Hidden = Not ScreenMode
        grdSOTWHODI.DisplayLayout.Bands(0).Columns("DIDNOT").Hidden = Not ScreenMode

        Dim dvw As DataView = DirectCast(grdSOTWHODI.DataSource, DataTable).DefaultView

        dvw.RowFilter = IIf(ScreenMode, "SEL = '1'", "")
        Setup_grdSOTWHOD1()

        If ScreenMode Then
            grdSOTWHODI.Parent = splSOTWHODI.Panel1
        Else
            Clear_Record()
            grdSOTWHODI.Parent = spl.Panel2

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTWHODI", "SOTWHOD1", "SOTWHOD2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If cbeYP0.Value & "" = "" Then
            cbeYP0.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12)
            cbeYP1.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        End If

        EnforceConstraints(True)

        Clear_Items()

        'Absx1.txtFor("COLLECTION_CODE").Text = ""
        'Absx1.txtFor("ITEM_CODE").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        ASCDATA1.DeleteRows(dst.Tables("SOTWHODI"), "ISNULL(SEL,'0')<>'1'")
        Update_Record_TDA("SOTWHODI")
        Load_SOTWHODx()
        Fill_Records("SOTWHOD1")
        For Each rowSOTWHODI As DataRow In dst.Tables("SOTWHODI").Select("SEL='1'")
            Dim ITEM_CODE As String = rowSOTWHODI.Item("ITEM_CODE")
            rowSOTWHODI.Item("DID") = dst.Tables("SOTWHOD1").Select("ITEM_CODE = '" & ITEM_CODE & "' and ISNULL(ORDR_COUNT,0)<>0").Length
            rowSOTWHODI.Item("DIDNOT") = dst.Tables("SOTWHOD1").Select("ITEM_CODE = '" & ITEM_CODE & "' and ISNULL(ORDR_COUNT,0)=0").Length
        Next
        Fill_Records("SOTWHOD1_NOT")

        Sort_grdColumns(grdSOTWHODI, "ITEM_CODE")
        Setup_grdSOTWHOD1()
        Setup_grdSOTORDR2()

        EnforceConstraints(True)

        'Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTWHODI, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Item Status Inquiry")
        Load_Popup_Menu(grdSOTWHOD1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry", "Customer Order Inquiry")
        Load_Popup_Menu(grdSOTWHOD1_NOT, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Order Inquiry")
        Load_Popup_Menu(grdSOTORDR2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Item Status Inquiry", "Customer Order Inquiry", "Sales Order Inquiry")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSOTWHODI"
                    tlb_pop.Tools("Select All").SharedProps.Visible = Not ScreenMode
                    tlb_pop.Tools("De-Select All").SharedProps.Visible = Not ScreenMode
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each row As DataRow In dst.Tables("SOTWHODI").Select("")
                    row.Item("SEL") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next
            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "COLLECTION_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Load_ICTITEM1()
                End If
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Add_Item(Absx1.txtFor("ITEM_CODE").Text)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "COLLECTION_CODE"
                Load_ICTITEM1()

            Case "ITEM_CODE"
                Add_Item(Absx1.txtFor("ITEM_CODE").Text)

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "COLLECTION_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("COLLECTION_CODE").Text <> "" Then
                        LookUp("ICTCOLL1", Absx1.txtFor("COLLECTION_CODE").Text)
                        If cdr Is Nothing Then
                            Absx1.txtFor("COLLECTION_CODE").Text = ""
                        End If
                    End If
                End If
        End Select
    End Sub

#End Region

#Region "grdSOTWHOD1"

    Private Sub grdSOTWHOD1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTWHOD1.AfterRowActivate
        Setup_grdSOTORDR2()
    End Sub

#End Region

    Sub Setup_grdSOTORDR2()

        If grdSOTWHOD1.ActiveRow Is Nothing OrElse Not grdSOTWHOD1.ActiveRow.IsDataRow Then
            splSOTORDR2.Panel2Collapsed = True
        Else
            Dim ITEM_CODE As String = grdSOTWHOD1.ActiveRow.Cells("ITEM_CODE").Value
            Dim CUST_CODE As String = grdSOTWHOD1.ActiveRow.Cells("CUST_CODE").Value

            ASCMAIN1.sql = "" _
                & "Select SOTORDR2.*,SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & " from SOTORDR2,SOTORDR1" & vbCrLf _
                & " where SOTORDR2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and SOTORDR2.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
                & "   and SOTORDR2.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" & vbCrLf _
                & "   and SOTORDR2.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & " union " & vbCrLf _
                & "Select SOTORDR2.*,SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
                & " from SOTORDR2,SOTORDR1" & vbCrLf _
                & " where SOTORDR2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and SOTORDR2.ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
                & "   and SOTORDR2.ORDR_STATUS in ('O','P')" & vbCrLf _
                & IIf(optOpenPick.Value = "A", "", _
                  "   and SOTORDR1.ORDR_SHIP_DATE >= '" & Format(DTE0(0), "dd-MMM-yyyy") & "'" & vbCrLf _
                & "   and SOTORDR1.ORDR_SHIP_DATE <= '" & Format(DTE1(UBound(DTE1)), "dd-MMM-yyyy") & "'" & vbCrLf) _
                & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
            Fill_Records("SOTORDR2", "", True, ASCMAIN1.sql)

            Sort_grdColumns(grdSOTORDR2, "ORDR_NO")
            grdSOTORDR2.Text = "Sales Order Details for Item " & ITEM_CODE & ", Customer " & CUST_CODE

            splSOTORDR2.Panel2Collapsed = False
        End If
    End Sub

    Private Sub grdSOTWHODI_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTWHODI.AfterRowActivate
        Setup_grdSOTWHOD1()
        Setup_grdSOTORDR2()
    End Sub

    Private Sub grdICTITEM1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTWHODI.DoubleClickRow
        'If e.Row.IsDataRow Then
        '    Absx1.txtFor("COLLECTION_CODE").Text = e.Row.Cells("COLLECTION_CODE").Value
        '    Click_Command("View")
        'End If
    End Sub

    Sub Setup_grdSOTWHOD1()
        If grdSOTWHODI.ActiveRow Is Nothing OrElse Not grdSOTWHODI.ActiveRow.IsDataRow Then
            splSOTWHODI.Panel2Collapsed = True
        Else
            Dim ITEM_CODE As String = grdSOTWHODI.ActiveRow.Cells("ITEM_CODE").Value

            Dim dvw As DataView = Nothing

            dvw = DirectCast(grdSOTWHOD1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "' and ISNULL(ORDR_COUNT,0) <>0"
            Sort_grdColumns(grdSOTWHOD1, "CUST_CODE")
            grdSOTWHOD1.Text = "Customers Who Did Order Item " & ITEM_CODE

            dvw = DirectCast(grdSOTWHOD1_NOT.DataSource, DataTable).DefaultView
            dvw.RowFilter = "ITEM_CODE = '" & ITEM_CODE & "'"
            Sort_grdColumns(grdSOTWHOD1_NOT, "CUST_CODE")
            grdSOTWHOD1_NOT.Text = "Customers Who Did NOT Order Item " & ITEM_CODE

            If optSelection.Value = "I" Then
                grdSOTWHOD1_NOT.Text = "Customers Who Didn't Order Item " & ITEM_CODE & " (but have ordered from the List of Selected Items)"
            Else
                grdSOTWHOD1_NOT.Text = "Customers Who Didn't Order Item " & ITEM_CODE & " (but have ordered from the Collection " & COLLECTION_CODE & ")"
            End If

            splSOTWHODI.Panel2Collapsed = False
        End If
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Who Did & Who Didn't Order Report")

        Print_Report_Begin()
        ' CR_params.Add("NOTES", "1")
        Generate_Report("SORWHOD1", "", "")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_ICTITEM1()

        Dim COLLECTION_CODE As String = Absx1.txtFor("COLLECTION_CODE").Text
        rowICTCOLL1 = LookUp("ICTCOLL1", COLLECTION_CODE)
        If rowICTCOLL1 Is Nothing Then
            dst.Tables("SOTWHODI").Rows.Clear()
            grdSOTWHODI.Visible = False
        Else
            Create_Temp_Tables()
            Fill_Records("SOTWHODI")
            Sort_grdColumns(grdSOTWHODI, "ITEM_CODE")
            grdSOTWHODI.Text = "Items defined to Collection " & Absx1.txtFor("COLLECTION_CODE").Text
            grdSOTWHODI.Visible = True
        End If
    End Sub

    Sub Load_SOTWHODx()
        RYP0 = cbeYP0.Value
        RYP1 = cbeYP1.Value

        DTE0 = ASCMAIN1.Get_Dates(RYP0)
        DTE1 = ASCMAIN1.Get_Dates(RYP1)

        ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE, SOTORDR2.CUST_CODE" & vbCrLf _
           & ", Count (*) ORDR_COUNT" & vbCrLf _
           & ", Sum (Case when SOTORDR2.ORDR_QTY_SHIP <> 0 THEN 1 ELSE 0 END) SHIP_COUNT" & vbCrLf _
           & ", Sum (Case when SOTORDR2.ORDR_QTY_PICK <> 0 THEN 1 ELSE 0 END) PICK_COUNT" & vbCrLf _
           & ", Sum (Case when SOTORDR2.ORDR_QTY_OPEN <> 0 THEN 1 ELSE 0 END) OPEN_COUNT" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_QTY) ORDR_UNITS" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_QTY_SHIP) SHIP_UNITS" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_QTY_PICK) PICK_UNITS" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_QTY_OPEN) OPEN_UNITS" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_UNIT_PRICE * SOTORDR2.ORDR_QTY) ORDR_SALES" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_UNIT_PRICE * SOTORDR2.ORDR_QTY_SHIP) SHIP_SALES" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_UNIT_PRICE * SOTORDR2.ORDR_QTY_PICK) PICK_SALES" & vbCrLf _
           & ", Sum (SOTORDR2.ORDR_UNIT_PRICE * SOTORDR2.ORDR_QTY_OPEN) OPEN_SALES" & vbCrLf _
           & " from SOTORDR2" & vbCrLf _
           & " where SOTORDR2.ITEM_CODE in (Select ITEM_CODE from " & SOTWHODI & " where SEL = '1')" & vbCrLf
        ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE" & vbCrLf _
           & ", Sum (ORDR_COUNT) ORDR_COUNT" & vbCrLf _
           & ", Sum (SHIP_COUNT) SHIP_COUNT" & vbCrLf _
           & ", Sum (PICK_COUNT) PICK_COUNT" & vbCrLf _
           & ", Sum (OPEN_COUNT) OPEN_COUNT" & vbCrLf _
           & ", Sum (ORDR_UNITS) ORDR_UNITS" & vbCrLf _
           & ", Sum (SHIP_UNITS) SHIP_UNITS" & vbCrLf _
           & ", Sum (PICK_UNITS) PICK_UNITS" & vbCrLf _
           & ", Sum (OPEN_UNITS) OPEN_UNITS" & vbCrLf _
           & ", Sum (ORDR_SALES) ORDR_SALES" & vbCrLf _
           & ", Sum (SHIP_SALES) SHIP_SALES" & vbCrLf _
           & ", Sum (PICK_SALES) PICK_SALES" & vbCrLf _
           & ", Sum (OPEN_SALES) OPEN_SALES" & vbCrLf _
           & ", 0 ALLO_UNITS" & vbCrLf _
           & " from (" & vbCrLf _
           & ASCMAIN1.sql & vbCrLf _
           & "   and SOTORDR2.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" & vbCrLf _
           & "   and SOTORDR2.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'" & vbCrLf _
           & " group by SOTORDR2.ITEM_CODE, SOTORDR2.CUST_CODE" & vbCrLf _
           & " union " & vbCrLf _
           & IIf(optOpenPick.Value = "A", _
                 ASCMAIN1.sql, _
                 Replace(ASCMAIN1.sql, "from SOTORDR2", "from SOTORDR1,SOTORDR2")) & vbCrLf _
           & "   and SOTORDR2.ORDR_STATUS in ('O','P')" & vbCrLf _
           & IIf(optOpenPick.Value = "A", "", _
             "   and SOTORDR1.ORDR_SHIP_DATE >= '" & Format(DTE0(0), "dd-MMM-yyyy") & "'" & vbCrLf _
           & "   and SOTORDR1.ORDR_SHIP_DATE <= '" & Format(DTE1(UBound(DTE1)), "dd-MMM-yyyy") & "'" & vbCrLf _
           & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf) _
           & " group by SOTORDR2.ITEM_CODE, SOTORDR2.CUST_CODE" & vbCrLf _
           & ") group by ITEM_CODE, CUST_CODE"

        ASCDATA1.ExecuteSQL("Truncate Table " & SOTWHOD1)
        ASCDATA1.ExecuteSQL("Insert into " & SOTWHOD1 & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select Distinct SOTORDR2.CUST_CODE from SOTORDR2,ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
            & IIf(optSelection.Value = "I", _
                  "   and SOTORDR2.ITEM_CODE in (Select ITEM_CODE from " & SOTWHODI & " where SEL = '1')", _
                  "   and ICTITEM1.COLLECTION_CODE = '" & COLLECTION_CODE & "'") & vbCrLf _
            & "   and SOTORDR2.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" & vbCrLf _
            & "   and SOTORDR2.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'" & vbCrLf
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTWHODC)
        ASCDATA1.ExecuteSQL("Insert into " & SOTWHODC & " " & ASCMAIN1.sql)

        Dim sqlALLOx As String = "" _
            & "  from SOTALLO1,SOTALLO2" _
            & "  where SOTALLO1.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" _
            & "    and SOTALLO1.ITEM_CODE in (Select ITEM_CODE from " & SOTWHODI & " where SEL = '1')" _
            & "    and SOTALLO1.DATE_START <= '" & Format(DTE1(UBound(DTE1)), "dd-MMM-yyyy") & "'" _
            & "    and SOTALLO1.DATE_END >= '" & Format(DTE0(0), "dd-MMM-yyyy") & "'" 

        For Each rowSOTWHODI As DataRow In dst.Tables("SOTWHODI").Select("SEL = '1'")
            Dim ITEM_CODE As String = rowSOTWHODI.Item("ITEM_CODE")
            ASCMAIN1.sql = "Select '" & ITEM_CODE & "' ITEM_CODE, CUST_CODE from " & SOTWHODC & vbCrLf _
                & " where CUST_CODE in ((Select CUST_CODE from " & SOTWHODC & vbCrLf _
                & " union Select Distinct CUST_CODE " & sqlALLOx & ")" _
                & " minus Select CUST_CODE from " & SOTWHOD1 & " where ITEM_CODE = '" & ITEM_CODE & "')"
            ASCDATA1.ExecuteSQL("Insert into " & SOTWHOD1 & " (ITEM_CODE,CUST_CODE) " & ASCMAIN1.sql)
        Next

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is" & vbCrLf _
            & " Select SOTALLO1.ITEM_CODE, SOTALLO2.CUST_CODE, SOTALLO2.QTY_ALLO" & sqlALLOx & ";" & vbCrLf _
            & " Begin For R1 in C1 Loop" & vbCrLf _
            & "  Update " & SOTWHOD1 & " Set ALLO_UNITS = R1.QTY_ALLO" & vbCrLf _
            & "   where CUST_CODE = R1.CUST_CODE and ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & " End Loop; End; " & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Create_Temp_Tables()

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, '0' SEL" & vbCrLf _
            & " from ICTITEM1" & vbCrLf _
            & " where ICTITEM1.COLLECTION_CODE = '" & Absx1.txtFor("COLLECTION_CODE").Text & "'"

        If SOTWHODI = "" Then
            SOTWHODI = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTWHODI & " Add Primary Key (ITEM_CODE)")

            ASCMAIN1.sql = "Select CUST_CODE from SOTINVH2 where ROWNUM < 1"
            SOTWHODC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTWHODC & " Add Primary Key (CUST_CODE)")

            ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE from SOTINVH2 where ROWNUM < 1"
            SOTWHOD1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTWHOD1 & " Add Primary Key (ITEM_CODE,CUST_CODE)")
            ADD_COLS(SOTWHOD1)

            ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO from SOTINVH2 where ROWNUM < 1"
            SOTWHOD2 = ASCMAIN1.Temp_Table
                        ASCDATA1.ExecuteSQL("Alter Table " & SOTWHOD2 & " Add Primary Key (ITEM_CODE,CUST_CODE,CUST_STORE_NO)")
            ADD_COLS(SOTWHOD2)

        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTWHODI)
            ASCDATA1.ExecuteSQL("Insert into " & SOTWHODI & " " & ASCMAIN1.sql)
        End If

    End Sub

    Sub Add_Cols(TABLE_NAME As String)

        For Each DT As String In New String() {"COUNT", "UNITS", "SALES"}
            For Each DS As String In New String() {"ORDR", "SHIP", "PICK", "OPEN"}
                ASCDATA1.ExecuteSQL("Alter Table " & TABLE_NAME & " Add " & DS & "_" & DT & " NUMBER " & IIf(DT = "SALES", "(13,2)", "(8,0)"))
            Next
        Next
        ASCDATA1.ExecuteSQL("Alter Table " & TABLE_NAME & " Add ALLO_UNITS NUMBER (8,0)")
    End Sub

    Private Sub optWho_ValueChanged(sender As System.Object, e As System.EventArgs)
        If SELECTION_NO = 0 Then Exit Sub
        Setup_grdSOTWHOD1()
    End Sub

    Private Sub optSelection_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSelection.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Clear_Items()
    End Sub

    Sub Clear_Items()
        lblITEM_CODE.Visible = (optSelection.Value = "I")
        txtITEM_CODE.Visible = (optSelection.Value = "I")

        ASCMAIN1.sql = "Delete from " & SOTWHODI
        ASCDATA1.ExecuteSQL()
        Absx1.txtFor("COLLECTION_CODE").Text = ""
        Absx1.txtFor("ITEM_CODE").Text = ""
        dst.Tables("SOTWHODI").Rows.Clear()

        grdSOTWHODI.Visible = (optSelection.Value = "I")
        spl.Panel1Collapsed = (optSelection.Value = "I")

        If optSelection.Value = "I" Then
            grdSOTWHODI.Text = "Selected Items"
        End If
    End Sub

    Sub Add_Item(ITEM_CODE As String)
        ASCMAIN1.sql = "Insert into " & SOTWHODI _
            & " Select ITEM_CODE, '1' from ICTITEM1 where ITEM_CODE = '" & ITEM_CODE & "'" _
            & " and ITEM_CODE Not in (Select ITEM_CODE from " & SOTWHODI & ")"
        ASCDATA1.ExecuteSQL()
        Fill_Records("SOTWHODI")
        Absx1.txtFor("ITEM_CODE").Focus()
    End Sub
End Class