Public Class SOFORDF1
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim sqlICTITEM1 As String = ""
    Dim sqlICTITEM1_for_Allocations As String = ""
    Dim SOTORDR1_orig_column_count As Integer
    Dim HC_CODEs As New List(Of String)
    Dim BRAND_CODEs As New List(Of String)
    Dim ALLO_GROUP_CODEs As New List(Of String)

    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim xls_CBs As New List(Of String)
    Dim tblSOTORDF1 As DataTable

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "SOFORDFI" Then
            InquiryMode = True
        End If

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 12, -12) ' -11 - 12)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 12, -7) '  0 - 12)

        Dim Q As Integer = 1 + (Val(Now.Month) - 1) / 3
        Dim Y As Integer = 0
        If Q > 4 Then
            Q = 1
            Y = 1
        End If
        Dim MDY = Format(Now.Date.AddYears(Y), "MM/01/yyyy")
        Mid(MDY, 1, 2) = Format(1 + (Q - 1) * 3, "00")
        dteAllocation.Value = MDY

        Dim MDY2 = CDate(MDY).AddMonths(3).AddDays(-1)
        dteAllocation2.Value = MDY2

        With dst
            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_CITY, ARTCUST2.CUST_STORE_STATE" & vbCrLf _
                & ", ARTCUST2.SELL_CODE, SOTSELL1.SELL_NAME, SOTSELL1.REGION_CODE, SOTSREG1.REGION_DESC" & vbCrLf _
                & " from ARTCUST2,SOTSELL1,SOTSREG1" & vbCrLf _
                & " where SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" & vbCrLf _
                & "   and NVL(ARTCUST2.CUST_STORE_STATUS,'?') = 'A'"
            Create_TDA(.Tables.Add, "SOTORDF1", "**", 1, False, "", 2)
            With .Tables("SOTORDF1").Columns
                .Add("SEL")
                dst.Tables("SOTORDF1").Columns("SEL").DefaultValue = "1"
            End With
            Dim SOTORDR1_orig_column_count = dst.Tables("SOTORDF1").Columns.Count
            tblSOTORDF1 = dst.Tables("SOTORDF1").Copy

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.PROD_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.COST_CATGY_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.HC_CODE, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE, ICTITEM1.ITEM_ALT_SORT" & vbCrLf _
                & ", ICTITEM1.ITEM_STD_PACK_SLS, ICTITEM1.ITEM_SO_QTY_MIN, ICTITEM1.ITEM_SO_QTY_MULT" & vbCrLf _
                & ", ICTITEM1.ITEM_CODE_COMPARE_TO" & vbCrLf _
                & " from ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and NVL(ICTITEM1.ITEM_STATUS,'?') = 'A'"
            sqlICTITEM1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 0)
            ' NOTE THAT THIS TABLE CANNOT HAVE A KEY OF ITEM CODE SINCE WE ALLOWED MULTIPLE RECORDS FOR THE SAME ITEM FOR DIFFERENT ALLOCATIONS.
            ' PROBABLY SHOULD HAVE CHANGED THE NAME OF THE TABLE

            With .Tables("ICTITEM1").Columns
                .Add("SEL")
                dst.Tables("ICTITEM1").Columns("SEL").DefaultValue = "1"
                .Add("QTY_ALLO", GetType(System.Int64))
                .Add("ORDR_QTY", GetType(System.Int64))
                .Add("ORDR_QTY_OPEN", GetType(System.Int64))
                .Add("ORDR_QTY_PICK", GetType(System.Int64))
                .Add("ORDR_QTY_SHIP", GetType(System.Int64))
                .Add("ORDR_QTY_CANC", GetType(System.Int64))
                .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_OPEN,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
                .Add("QTY_OVER", GetType(System.Int64), "IIF(QTY_LEFT-ISNULL(ORDR_QTY_OPEN,0)>=0,0,ISNULL(ORDR_QTY_OPEN,0)-QTY_LEFT)")
                .Add("ALLO_CTL_NO")
                .Add("ALLO_NOTES")
                .Add("ALLO_GROUP_CODE")
                .Add("DATE_START", GetType(System.DateTime))
                .Add("DATE_END", GetType(System.DateTime))
            End With

            Create_TDA(.Tables.Add, "SOTALLO1", "*", 0, False, "", 1)
            Create_TDA(.Tables.Add, "SOTALLO2", "*", 0, False, "", 2)

            ASCMAIN1.sql = "Select BRAND_CODE, BRAND_NAME from ICTBRAN1 where BRAND_STATUS = 'A'"
            Create_TDA(.Tables.Add, "ICTBRAN1", "**", 0, False, "", 1)
            .Tables("ICTBRAN1").Columns.Add("SEL")
            .Tables("ICTBRAN1").Columns("SEL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "SOTALLOG", "*", 0, False)
            With .Tables("SOTALLOG").Columns
                .Add("SEL")
            End With
            .Tables("SOTALLOG").Columns("SEL").DefaultValue = "0"
        End With

        grdSOTORDF1.DataSource = dst.Tables("SOTORDF1")
        grdICTITEM1.DataSource = dst.Tables("ICTITEM1")

        grdICTBRAN1.DataSource = dst.Tables("ICTBRAN1")
        Fill_Records("ICTBRAN1")
        Sort_grdColumns(grdICTBRAN1, "BRAND_CODE")

        grdSOTALLOG.DataSource = dst.Tables("SOTALLOG")
        Fill_Records("SOTALLOG")
        dst.Tables("SOTALLOG").Rows.Add(New Object() {"", "Ungrouped", "A", DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, "1"})
        Sort_grdColumns(grdSOTALLOG, "ALLO_GROUP_CODE")


        Create_Summary(grdSOTORDF1, "CUST_STORE_NO", "Count")
        Create_Summary(grdSOTORDF1, "SEL")
        Create_Summary(grdICTITEM1, "ITEM_CODE", "Count")
        Create_Summary(grdICTITEM1, "SEL")


        With grdSOTORDF1.DisplayLayout.Bands(0)
            Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add()
            With G.Header.Appearance
                .BackColor = Drawing.Color.White
                .BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = Drawing.Color.LightGray
            End With

            G.Header.Caption = "Store"
            For Each COLUMN_NAME As String In New String() _
                {"SEL", "CUST_STORE_NO", "CUST_STORE_NAME", "CUST_STORE_CITY", "CUST_STORE_STATE", "SELL_CODE", "SELL_NAME", "REGION_CODE", "REGION_DESC"}
                Dim gcol As UltraWinGrid.UltraGridColumn = .Columns(COLUMN_NAME)
                gcol.Group = G
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = System.Drawing.Color.WhiteSmoke
                End If
                With gcol.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = Drawing.Color.LightGray
                End With
            Next
        End With


        With grdICTITEM1.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
            .Columns("ITEM_RETAIL_PRICE").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = System.Drawing.Color.WhiteSmoke
                End If
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    If New String() {"QTY_ALLO", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "QTY_LEFT", "QTY_BAL", "QTY_OVER"}.Contains(gcol.Key) Then
                        .BackColor2 = System.Drawing.Color.LightBlue
                    Else
                        .BackColor2 = System.Drawing.Color.LightGreen
                    End If
                End With
            Next
        End With


        grdSOTALLOG.DisplayLayout.UseFixedHeaders = True
        grdSOTALLOG.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        grdSOTALLOG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdSOTALLOG.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdSOTALLOG.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
            Next
            For Each COLUMN_NAME As String In New String() {"ALLO_GROUP_CODE", "ALLO_GROUP_DESC"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With

        '   dteAllocation.DateTime = Now.Date
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE")

                If grdICTBRAN1.ActiveRow IsNot Nothing AndAlso grdICTBRAN1.ActiveRow.DataChanged Then
                    'Dim X As CurrencyManager = Me.BindingContext(dst.Tables("ICTBRAN1"))
                    'X.EndCurrentEdit()
                    grdICTBRAN1.ActiveRow = Nothing
                    'grdICTBRAN1.UpdateData()
                    'grdICTBRAN1.ActiveRow.Update()
                    'Debug.Print("HAD TO UPDATE")
                    'If grdICTBRAN1.ActiveRow IsNot Nothing AndAlso grdICTBRAN1.ActiveRow.DataChanged Then
                    '    grdICTBRAN1.ActiveRow.CancelUpdate()
                    'End If
                    'grdICTBRAN1.Rows.Refresh(UltraWinGrid.RefreshRow.ReloadData)
                End If

                BRAND_CODEs.Clear()
                For Each row As DataRow In dst.Tables("ICTBRAN1").Select("SEL='1'")
                    BRAND_CODEs.Add(row.Item("BRAND_CODE"))
                Next
                If BRAND_CODEs.Count = 0 Then
                    EMsg &= vbCr & "You must select at least 1 Brand"
                End If

                ALLO_GROUP_CODEs.Clear()
                For Each row As DataRow In dst.Tables("SOTALLOG").Select("SEL = '1'")
                    ALLO_GROUP_CODEs.Add(row.Item(0))
                Next
                If ALLO_GROUP_CODEs.Count = 0 Then
                    EMsg &= vbCr & "You must select at least 1 Allocation Group (Perhaps Ungrouped)"
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

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Export"
                If chkMasterFile.Checked Then
                    Load_Excel_Promo()
                Else
                    Load_Excel()
                End If


            Case "Done"
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")

                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Export").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Export").Visible = ScreenMode
                End With

                .Groups("Retail Sales").Visible = ScreenMode
            End With
        End If

        If ScreenMode Then
            grdSOTORDF1.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdSOTORDF1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdSOTORDF1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        ' spl.Panel1Collapsed = ScreenMode
        splSOTORDF1.Visible = ScreenMode

        '  splSOTORDF1.Panel2Collapsed = True

        If ScreenMode Then
            grdICTBRAN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdSOTALLOG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            Clear_Record()
            grdICTBRAN1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdSOTALLOG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDF1", "ICTITEM1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        grdSOTORDF1.SuspendLayout()

        ASCMAIN1.Progress("Deleting HC Columns")

        dst.Tables("SOTORDF1").Rows.Clear()

        '  dst.Tables("SOTORDF1") = tblSOTORDF1.Copy
        With dst.Tables("SOTORDF1")
            For i As Integer = .Columns.Count - 1 To 0 Step -1
                If .Columns(i).ColumnName = "SEL" Then
                    Exit For
                Else
                    .Columns.Remove(.Columns(i).ColumnName)
                    ASCMAIN1.Progress("-", i)
                End If
            Next
        End With

        ASCMAIN1.Progress("Deleting Brand Groups")
        With grdSOTORDF1.DisplayLayout.Bands(0)
            If .Groups.Count > 1 Then
                For G As Integer = .Groups.Count - 1 To 1 Step -1
                    ASCMAIN1.Progress("-", G)
                    .Groups.Remove(G)
                Next
            End If
        End With

        ASCMAIN1.Progress("")

        grdSOTORDF1.ResumeLayout()

        For Each row As DataRow In dst.Tables("ICTBRAN1").Select("")
            row.Item("SEL") = "0"
        Next

        EnforceConstraints(True)

        Absx1.txtFor("CUST_CODE").Text = ""

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        CUST_CODE = Absx1.txtFor("CUST_CODE").Text

        EnforceConstraints(False)

        grdSOTORDF1.DisplayLayout.Bands(0).Summaries.Clear()

        With dst.Tables("SOTORDF1")
            If .Columns.Count > SOTORDR1_orig_column_count Then
                For dc As Integer = .Columns.Count - 1 To SOTORDR1_orig_column_count
                    .Columns.Remove(.Columns(dc).ColumnName)
                Next
            End If
        End With

        If grdSOTORDF1.DisplayLayout.Bands(0).Groups.Count > 1 Then
            For g As Integer = grdSOTORDF1.DisplayLayout.Bands(0).Groups.Count - 1 To 1
                grdSOTORDF1.DisplayLayout.Bands(0).Groups.Remove(g)
            Next
        End If

        Fill_Records("SOTORDF1", New String() {CUST_CODE})
        grdSOTORDF1.Text = "Spread Order Worksheet for " & CUST_CODE & ", Brands " & Join(BRAND_CODEs.ToArray, ",")
        Sort_grdColumns(grdSOTORDF1, "CUST_CODE,CUST_STORE_NO")


        ASCMAIN1.sql = sqlICTITEM1 _
            & "   and ICTCOLL1.BRAND_CODE in ('" & Join(BRAND_CODEs.ToArray, "','") & "')" _
            & IIf(optITEM_BASIC_PROMO.Value = "*", _
                  "", _
                  "   and ICTITEM1.ITEM_BASIC_PROMO = '" & optITEM_BASIC_PROMO.Value & "'") _
            & "   and ICTITEM1.ITEM_SNU_CODE = '" & optITEM_SNU_CODE.Value & "'"
        sqlICTITEM1_for_Allocations = ASCMAIN1.sql
        Fill_Records("ICTITEM1", "", True, ASCMAIN1.sql)

        grdICTITEM1.Text = "Active Items in Brands " & Join(BRAND_CODEs.ToArray, ",")
        Sort_grdColumns(grdICTITEM1, "ITEM_CODE")


        Get_Allocations()
        For Each row As DataRow In dst.Tables("ICTITEM1").Select("ISNULL(QTY_ALLO,0) = 0")
            row.Item("SEL") = "0"
        Next

        Create_Summary(grdSOTORDF1, "CUST_STORE_NO", "Count")
        Create_Summary(grdSOTORDF1, "SEL")

        Dim HCi As Integer = 0
        HC_CODEs.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct("ICTITEM1", New String() {"HC_CODE"}).Select("", "HC_CODE")
            Dim HC_CODE As String = row.Item("HC_CODE")
            HC_CODEs.Add(HC_CODE)
            HCi += 1
            Dim PFX As String = "HC" & Format(HCi, "00")
            With dst.Tables("SOTORDF1")
                .Columns.Add(PFX & "_RTL", GetType(System.Decimal))
                .Columns.Add(PFX & "_PCT", GetType(System.Decimal))
                .Columns.Add(PFX & "_AUTH")
                If CUST_CODE = "ULTA" Then
                    .Columns(PFX & "_AUTH").DefaultValue = "1"
                Else
                    .Columns(PFX & "_AUTH").DefaultValue = "0"
                End If
            End With

            ASCMAIN1.sql = "Select Distinct CUST_STORE_NO from SATAUTH1" & vbCrLf _
                & " where SATAUTH1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                & "   and SATAUTH1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                & "   and SATAUTH1.OPS_YYYYPP_OPENED is Not Null and SATAUTH1.OPS_YYYYPP_CLOSED is Null"
            For Each rowSATAUTH1 As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim CUST_STORE_NO As String = rowSATAUTH1.Item("CUST_STORE_NO")
                Dim rowSOTORDF1 As DataRow = dst.Tables("SOTORDF1").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                If rowSOTORDF1 IsNot Nothing Then
                    rowSOTORDF1.Item(PFX & "_AUTH") = "1"
                End If
            Next

            If CUST_CODE = "ULTA" Then
                For Each rowSOTORDF1 As DataRow In dst.Tables("SOTORDF1").Rows
                    rowSOTORDF1.Item(PFX & "_AUTH") = "1"
                Next
            End If

            With grdSOTORDF1.DisplayLayout.Bands(0)
                Dim G As UltraWinGrid.UltraGridGroup = .Groups.Add(HC_CODE)
                G.Header.Caption = HC_CODE
                With G.Header.Appearance
                    .BackColor = Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = Drawing.Color.Orange
                    .TextHAlign = HAlign.Center
                End With
                For Each SFX As String In New String() {"_RTL", "_PCT", "_AUTH"}
                    With .Columns(PFX & SFX)
                        .Group = G
                        .Hidden = False
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        With .Header.Appearance
                            .BackColor = Drawing.Color.White
                            .BackGradientStyle = GradientStyle.ForwardDiagonal
                            .BackColor2 = Drawing.Color.Violet
                        End With
                    End With
                Next
                With .Columns(PFX & "_RTL")
                    .Width = 70
                    .Header.Caption = "$Retail"
                    .Format = "#,##0"
                    .CellAppearance.TextHAlign = HAlign.Right
                    .Header.Appearance.TextHAlign = HAlign.Right
                End With
                With .Columns(PFX & "_PCT")
                    .Width = 45
                    .Header.Caption = "%Ttl"
                    .Format = "#.0"
                    .CellAppearance.TextHAlign = HAlign.Right
                    .Header.Appearance.TextHAlign = HAlign.Right
                End With
                With .Columns(PFX & "_AUTH")
                    .Width = 45
                    .Header.Caption = "Auth"
                    .Style = UltraWinGrid.ColumnStyle.CheckBox
                    .Header.Appearance.TextHAlign = HAlign.Center
                End With
            End With

            Create_Summary(grdSOTORDF1, PFX & "_RTL")
            Create_Summary(grdSOTORDF1, PFX & "_PCT")
            Create_Summary(grdSOTORDF1, PFX & "_AUTH")


        Next


        EnforceConstraints(True)
        'Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()
            Update_Record_TDA("SOTINVH1")
            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDF1, "SS", "Show Filter", "Show GroupBox", "Select All", "De-Select All")
        Load_Popup_Menu(grdICTITEM1, "SSBBB", "Show Filter", "Show GroupBox", "Select All", "De-Select All", "Item Status Inquiry")
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
                Case "grdSOTINVH1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim rowsSkipped As Int16 = 0

        Select Case e.Tool.Key

            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If Not grow.IsFilteredOut Then
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next


            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then

                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

    Public Overrides Sub cmb_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.cmb_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If COLUMN_NAME = "RYP0" Or COLUMN_NAME = "RYP1" Then
            If Absx1.cmbFor("RYP0").Value & "" <> "" And Absx1.cmbFor("RYP1").Value & "" <> "" Then
                Dim C As Integer = 1 + ASCMAIN1.Period_Diff(Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value)
                lblMonths.Text = CStr(C) & " Mos"
            Else
                lblMonths.Text = ""
            End If
        End If
    End Sub

#End Region

    Private Sub btnRetailSales_Click(sender As Object, e As EventArgs) Handles btnRetailSales.Click
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("Now Loading Retail Sales by Store")

        For Each rowSOTORDF1 As DataRow In dst.Tables("SOTORDF1").Select("")
            For I As Integer = 1 To HC_CODEs.Count
                Dim C As String = "HC" & Format(I, "00") & "_RTL"
                rowSOTORDF1.Item(C) = DBNull.Value
            Next
        Next

        Dim sql As String = ""
        Dim HCi As Integer = 0
        For Each HC_CODE As String In HC_CODEs
            HCi += 1
            sql &= ", Sum (DECODE(ICTCOLL1.HC_CODE,'" & HC_CODE & "',RETAIL_SALES,0)) HC" & Format(HCi, "00") & "_RTL" & vbCrLf
        Next

        ASCMAIN1.sql = "" _
            & "Select RSTRETL2.CUST_CODE,RSTRETL2.CUST_STORE_NO" & vbCrLf _
            & sql _
            & " from RSTRETL2,ICTCOLL1" & vbCrLf _
            & " where RSTRETL2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ICTCOLL1.BRAND_CODE in ('" & Join(BRAND_CODEs.ToArray, "','") & "')" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = RSTRETL2.COLLECTION_CODE" & vbCrLf _
            & "   and RSTRETL2.OPS_YYYYPP between '" & Absx1.cmbFor("RYP0").Value & "' and '" & Absx1.cmbFor("RYP1").Value & "'" & vbCrLf _
            & " group by RSTRETL2.CUST_CODE,RSTRETL2.CUST_STORE_NO"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "CUST_STORE_NO")
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
            Dim rowSOTORDF1 As DataRow = dst.Tables("SOTORDF1").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            If rowSOTORDF1 Is Nothing Then
                rowSOTORDF1 = dst.Tables("SOTORDF1").Rows.Find(New String() {CUST_CODE, "XXXXXX"})
                If rowSOTORDF1 Is Nothing Then
                    rowSOTORDF1 = dst.Tables("SOTORDF1").Rows.Add(New String() {CUST_CODE, "XXXXXX", "All Other Stores"})
                    Dim grow As UltraWinGrid.UltraGridRow = grdSOTORDF1.Rows.GetRowWithListIndex(dst.Tables("SOTORDF1").Rows.IndexOf(rowSOTORDF1))
                    grow.Fixed = True

                    grdSOTORDF1.DisplayLayout.Override.FixedRowStyle = Infragistics.Win.UltraWinGrid.FixedRowStyle.Bottom
                    grdSOTORDF1.DisplayLayout.Override.SpecialRowSeparator = UltraWinGrid.SpecialRowSeparator.FixedRows

                End If
            End If
            For I As Integer = 1 To HC_CODEs.Count
                Dim C As String = "HC" & Format(I, "00") & "_RTL"
                rowSOTORDF1.Item(C) = Val(rowSOTORDF1.Item(C) & "") + Val(row.Item(C) & "")
            Next
        Next

        HCi = 0
        For Each HC_CODE As String In HC_CODEs
            HCi += 1
            Dim PFX As String = "HC" & Format(HCi, "00")
            Dim TOTAL As Decimal = Val(dst.Tables("SOTORDF1").Compute("SUM(" & PFX & "_RTL)", "") & "")
            If TOTAL = 0 Then
            Else
                dst.Tables("SOTORDF1").Columns(PFX & "_PCT").Expression = "100*" & PFX & "_RTL/" & CStr(TOTAL)
            End If
        Next


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Excel_Promo()
        Dim Selected_HC_CODEs As New List(Of String)
        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangeCopyFrom As SpreadsheetGear.IRange
        Dim rangePaste_To As SpreadsheetGear.IRange
        Dim Start_Row As Integer = 12 'needs to be changed if adding anything
        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        'ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(dteAllocation.DateTime, "dd-MMM-yyyy") & "'"
        'Dim TYW As String = ASCDATA1.GetDataValue
        'Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", TYW)

        'Dim LYW As String = ASCMAIN1.Week_Calc(TYW, -52)

        'Dim TYW_start As String = Mid(TYW, 1, 4) & IIf(Mid(TYW, 5, 2) >= "27", "27", "01")
        'Dim LYW_start As String = ASCMAIN1.Week_Calc(TYW_start, -52)

   
        Dim TYP As String = Format(dteAllocation.DateTime, "yyyyMM")
        Dim TYP_start As String = Mid(TYP, 1, 4) & IIf(Mid(TYP, 5, 2) >= "07", "07", "01")
        Dim TYP_end As String = ASCMAIN1.Period_Calc(TYP_start, 5)
        Dim LYP As String = ASCMAIN1.Period_Calc(TYP, -12)
        Dim LYP_start As String = ASCMAIN1.Period_Calc(TYP_start, -12)
        Dim LYP_end As String = ASCMAIN1.Period_Calc(LYP_start, 5)

        Dim SEASON_CODE As String = Mid(TYP, 1, 4) & IIf(Mid(TYP, 5, 2) >= "07", "F", "S")

        Dim FILENAME_SOURCE As String = ASCMAIN1.Folders("SharedRoot") & "Templates\SORORDF1.xlsx"
        Dim FILENAME As String = ""

        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "gcv" Then
            FILENAME_SOURCE = "C:\Share\INT\Templates\SORORDF1.xlsx"
        End If

   
        'XLS_NO += 1
        'xls_name = ASCMAIN1.DBS_SESSION_ID
        'xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
        'FILENAME = xls_path & xls_name

        ' FileCopy(FILENAME_SOURCE, FILENAME)

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "MasterFile"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FileCopy(FILENAME_SOURCE, XLS_FILENAME)
                success = True

            Catch ex As Exception
                ' Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME)

        Selected_HC_CODEs.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ICTITEM1").Select("SEL = '1'"), "HC_CODE").Rows
            Selected_HC_CODEs.Add(row.Item("HC_CODE"))
        Next

        ASCMAIN1.sql = "Select Distinct ICTCOLL1.HC_CODE, SPTCWRXC.CHECKBOOK from ICTCOLL1,SPTCWRXC" & vbCrLf _
            & " where SPTCWRXC.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_STATUS = 'A'" & vbCrLf _
            & "   and (SPTCWRXC.COLLECTION_GENDER = 'U' or ICTCOLL1.COLLECTION_GENDER = SPTCWRXC.COLLECTION_GENDER)" & vbCrLf _
            & "   and SPTCWRXC.BRAND_CODE in ('" & Join(BRAND_CODEs.ToArray, "','") & "')" & vbCrLf _
            & "   and ICTCOLL1.HC_CODE in ('" & Join(Selected_HC_CODEs.ToArray, "','") & "')"

        xls_CBs.Clear()
        For Each rowCHECKBOOK1 As DataRow In ASCDATA1.GetDataTable.Rows
            xls_CBs.Add(rowCHECKBOOK1.Item("CHECKBOOK") & "-" & rowCHECKBOOK1.Item("HC_CODE"))
        Next

        If xls_CBs.Count = 0 Then
            MsgBox("No Items to Show by High-Collection/Checkbook")
        Else
            For x As Integer = 0 To xls_CBs.Count - 1
                Dim xls_CB As String = xls_CBs(x)
                Dim CHECKBOOK As String = Split(xls_CB, "-")(0)
                Dim HC_CODE As String = Split(xls_CB, "-")(1)
                Dim workSheetSummary As SpreadsheetGear.IWorksheet = oWB.Worksheets("Sheet1")

                oSheet = oWB.Worksheets.Add
                oSheet.Name = xls_CBs(x)

                rangeCopyFrom = workSheetSummary.Range("A:R")
                rangePaste_To = oSheet.Range("A:R")
                rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)

                oSheet.Cells(Start_Row, 8).Value = CHECKBOOK & " LY Freelance"
                oSheet.Cells(Start_Row, 9).Value = CHECKBOOK & "-" & HC_CODE & " TY Freelance"
                oSheet.Cells(Start_Row, 11).Value = CHECKBOOK & "-" & HC_CODE & " LY $"
                oSheet.Cells(Start_Row, 12).Value = CHECKBOOK & "-" & HC_CODE & " TY Plan"

                oSheet.Cells(0, 0).Value = "Prepared"
                oSheet.Cells(1, 0).Value = "Customer"
                oSheet.Cells(2, 0).Value = "Season"

                oSheet.Cells(1, 3).Value = "Chkbk"
                oSheet.Cells(2, 3).Value = "HC"

                oSheet.Cells(0, 1).Value = Format(Now, "MM/dd/yyyy HH:mm")
                oSheet.Cells(1, 1).Value = CUST_CODE
                oSheet.Cells(2, 1).Value = SEASON_CODE

                oSheet.Cells(1, 4).Value = CHECKBOOK
                oSheet.Cells(2, 4).Value = HC_CODE

                Dim ITEM_CODE_COMPARE_TO As String = ""

                Dim ITEM_CODEs As New List(Of String)
                Dim ITEM_CODE_COMPARE_TOs As New List(Of String)
                Dim ITEM_CODEs_LY As New Dictionary(Of String, String)

                Dim iITEM_CODE As Integer = 0
                Dim XR_source As String = Excel_Cell0(-1, 17 + 1) & ":" & Excel_Cell0(-1, 17 + 5)
                rangeCopyFrom = workSheetSummary.Range(XR_source)
                For Each row As DataRow In dst.Tables("ICTITEM1").Select("SEL = '1' and HC_CODE = '" & HC_CODE & "'", "ITEM_CODE")
                    iITEM_CODE += 1
                    Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                    ITEM_CODEs.Add(ITEM_CODE)
                    Dim XR As String = Excel_Cell0(-1, 17 + (iITEM_CODE - 1) * 5 + 1) & ":" & Excel_Cell0(-1, 17 + (iITEM_CODE - 1) * 5 + 1)
                    rangePaste_To = oSheet.Range(XR)
                    rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
                    Dim rowICTITEM1 As DataRow = row ' dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                    Dim QTY_ALLO As Int64 = Val(rowICTITEM1.Item("QTY_ALLO") & "")
                    'For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    '    Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")
                    '    Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").
                    'Next


                    oSheet.Cells(Start_Row - 12, 17 + (iITEM_CODE - 1) * 5 + 2).Value = ITEM_CODE
                    oSheet.Cells(Start_Row - 11, 17 + (iITEM_CODE - 1) * 5 + 2).Value = rowICTITEM1.Item("ITEM_EAN_CODE")
                    oSheet.Cells(Start_Row - 10, 17 + (iITEM_CODE - 1) * 5 + 2).Value = rowICTITEM1.Item("ITEM_ALT_SORT")
                    oSheet.Cells(Start_Row - 9, 17 + (iITEM_CODE - 1) * 5 + 2).Value = Format(Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & ""), "$#.00")
                    oSheet.Cells(Start_Row - 8, 17 + (iITEM_CODE - 1) * 5 + 2).Value = Format(Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & ""), "#,##0")
                    oSheet.Cells(Start_Row - 7, 17 + (iITEM_CODE - 1) * 5 + 2).Value = Format(Val(rowICTITEM1.Item("ITEM_SO_QTY_MIN") & ""), "#,##0") 'indexes need to be changed if adding
                    oSheet.Cells(Start_Row - 6, 17 + (iITEM_CODE - 1) * 5 + 2).Value = Format(QTY_ALLO, "#,##0")
                    oSheet.Cells(Start_Row - 5, 17 + (iITEM_CODE - 1) * 5 + 2).Value = Val(rowICTITEM1.Item("ORDR_QTY_SHIP") & "") + Val(rowICTITEM1.Item("ORDR_QTY_OPEN") & "") + Val(rowICTITEM1.Item("ORDR_QTY_PICK") & "")
                    oSheet.Cells(Start_Row - 0, 17 + (iITEM_CODE - 1) * 5 + 2).Value = rowICTITEM1.Item("ITEM_DESC")

                    ITEM_CODE_COMPARE_TO = rowICTITEM1.Item("ITEM_CODE_COMPARE_TO") & ""
                    If ITEM_CODE_COMPARE_TO = "" Then
                        ITEM_CODE_COMPARE_TO = ITEM_CODE
                        ITEM_CODE_COMPARE_TOs.Add(ITEM_CODE)
                        If Not ITEM_CODEs_LY.ContainsKey(ITEM_CODE) Then
                            ITEM_CODEs_LY.Add(ITEM_CODE, ITEM_CODE)
                        End If
                    Else
                        ''Dim rowICTITEM1_COMPARE_TO As DataRow = LookUp("ICTITEM1", ITEM_CODE_COMPARE_TO)
                        'oSheet.Cells(Start_Row - 11, 17 + (iITEM_CODE - 1) * 5 + 2 + 1).Value = ITEM_CODE_COMPARE_TO
                        ''oSheet.Cells(Start_Row - 0, 17 + (iITEM_CODE - 1) * 5 + 2 + 1).Value = rowICTITEM1_COMPARE_TO.Item("ITEM_DESC")
                        ITEM_CODE_COMPARE_TOs.Add(ITEM_CODE_COMPARE_TO)
                        If Not ITEM_CODEs_LY.ContainsKey(ITEM_CODE) Then
                            ITEM_CODEs_LY.Add(ITEM_CODE, ITEM_CODE_COMPARE_TO)
                        End If
                    End If
                Next


                'ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, SUM (LY_SELL_IN) LY_SELL_IN, SUM (LY_SELL_THRU) LY_SELL_THRU from (" & vbCrLf _
                '    & "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, SUM(ORDR_QTY_SHIP) LY_SELL_IN, 0 LY_SELL_THRU from SOTINVH2" & vbCrLf _
                '    & " where ITEM_CODE in ('" & Join(ITEM_CODE_COMPARE_TOs.ToArray, "','") & "')" & vbCrLf _
                '    & "   and ORDR_YYYYPP_UPDATED between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                '    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                '    & " union " & vbCrLf _
                '    & "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, 0 LY_SELL_IN, SUM (QTY_SOLD) LY_SELL_THRU from RSTRETL1" & vbCrLf _
                '    & " where ITEM_CODE in ('" & Join(ITEM_CODE_COMPARE_TOs.ToArray, "','") & "')" & vbCrLf _
                '    & "   and OPS_YYYYPP between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                '    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                '    & ") X" & vbCrLf _
                '    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO"

                ' SWITCH TO DOLLARS

                'ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, SUM (LY_SELL_IN) LY_SELL_IN, SUM (LY_SELL_THRU) LY_SELL_THRU from (" & vbCrLf _
                '    & "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, SUM(ORDR_QTY_SHIP * ITEM_RETAIL_PRICE) LY_SELL_IN, 0 LY_SELL_THRU from SOTINVH2" & vbCrLf _
                '    & " where ITEM_CODE in ('" & Join(ITEM_CODE_COMPARE_TOs.ToArray, "','") & "')" & vbCrLf _
                '    & "   and ORDR_YYYYPP_UPDATED between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                '    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                '    & " union " & vbCrLf _
                '    & "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, 0 LY_SELL_IN, SUM (AMT_SOLD) LY_SELL_THRU from RSTRETL1" & vbCrLf _
                '    & " where ITEM_CODE in ('" & Join(ITEM_CODE_COMPARE_TOs.ToArray, "','") & "')" & vbCrLf _
                '    & "   and OPS_YYYYPP between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                '    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                '    & ") X" & vbCrLf _
                '    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO"

                ' SWITCH BACK TO UNITS

                ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, SUM (LY_SELL_IN) LY_SELL_IN, SUM (LY_SELL_THRU) LY_SELL_THRU from (" & vbCrLf _
                    & "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, SUM(ORDR_QTY_SHIP) LY_SELL_IN, 0 LY_SELL_THRU from SOTINVH2" & vbCrLf _
                    & " where ITEM_CODE in ('" & Join(ITEM_CODE_COMPARE_TOs.ToArray, "','") & "')" & vbCrLf _
                    & "   and ORDR_YYYYPP_UPDATED between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                    & " union " & vbCrLf _
                    & "Select ITEM_CODE, CUST_CODE, CUST_STORE_NO, 0 LY_SELL_IN, SUM (QTY_SOLD) LY_SELL_THRU from RSTRETL1" & vbCrLf _
                    & " where ITEM_CODE in ('" & Join(ITEM_CODE_COMPARE_TOs.ToArray, "','") & "')" & vbCrLf _
                    & "   and OPS_YYYYPP between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO" & vbCrLf _
                    & ") X" & vbCrLf _
                    & " group by ITEM_CODE, CUST_CODE, CUST_STORE_NO"
                Dim tblSIST As DataTable = ASCDATA1.GetDataTable
                tblSIST.PrimaryKey = New DataColumn() {tblSIST.Columns("ITEM_CODE"), tblSIST.Columns("CUST_CODE"), tblSIST.Columns("CUST_STORE_NO")}

                Dim Row_Ctr As Int64 = 1

                ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
                    & ", Decode(ICTCOLL1.COLLECTION_GENDER,'W',ARTCUST2.CUST_STORE_CUST_RANK_W,ARTCUST2.CUST_STORE_CUST_RANK_M) CUST_STORE_CUST_RANK" & vbCrLf _
                    & ", '' FALL_LEGEND, '' RANK_LEGEND" & vbCrLf _
                    & ", SOTSELL1.SELL_NAME, SOTSREG1.REGION_DESC, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                    & ", SPTCWRXX.LY_MDL_ACT, SPTMBUDX.TY_MDL_BUD" & vbCrLf _
                    & ", RSTRETLX.LY_RTL_ACT, RSTBUDRX.TY_RTL_BUD" & vbCrLf _
                    & " from SATAUTH1, ARTCUST2, SOTSELL1, SOTSREG1, SPTCWRXC, ICTCOLL1, ICTCOLL0" & vbCrLf _
                    & ", (Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, Sum (RSTRETL1.AMT_SOLD) LY_RTL_ACT" & vbCrLf _
                    & " from RSTRETL1,ICTITEM1,ICTCOLL1" & vbCrLf _
                    & "    where RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "      and RSTRETL1.OPS_YYYYWW between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                    & "      and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
                    & "      and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & "      and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                    & "    group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO) RSTRETLX" & vbCrLf _
                    & ", (Select RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO, Sum (RSTBUDR1.BUDGET) TY_RTL_BUD" & vbCrLf _
                    & " from RSTBUDR1,ICTCOLL1" & vbCrLf _
                    & "    where RSTBUDR1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "      and RSTBUDR1.OPS_YYYYPP between '" & TYP_start & "' and '" & TYP_end & "'" & vbCrLf _
                    & "      and ICTCOLL1.COLLECTION_CODE = RSTBUDR1.COLLECTION_CODE" & vbCrLf _
                    & "      and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                    & "    group by RSTBUDR1.CUST_CODE, RSTBUDR1.CUST_STORE_NO) RSTBUDRX" & vbCrLf _
                    & ", (Select SPTCWRX2.CUST_CODE, SPTCWRX2.CUST_STORE_NO, Sum (SPTCWRX2.BILL_AMT) LY_MDL_ACT" & vbCrLf _
                    & " from SPTCWRX2" & vbCrLf _
                    & "    where SPTCWRX2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "      and SPTCWRX2.OPS_YYYYPP between '" & LYP_start & "' and '" & LYP_end & "'" & vbCrLf _
                    & "      and SPTCWRX2.CHECKBOOK = '" & CHECKBOOK & "'" & vbCrLf _
                    & "    group by SPTCWRX2.CUST_CODE, SPTCWRX2.CUST_STORE_NO) SPTCWRXX" & vbCrLf _
                    & ", (Select SPTMBUD1.CUST_CODE, SPTMBUD1.CUST_STORE_NO, Sum (SPTMBUD1.BUDGET) TY_MDL_BUD" & vbCrLf _
                    & " from SPTMBUD1,ICTCOLL1" & vbCrLf _
                    & "    where SPTMBUD1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "      and SPTMBUD1.OPS_YYYYPP between '" & TYP_start & "' and '" & TYP_end & "'" & vbCrLf _
                    & "      and ICTCOLL1.COLLECTION_CODE = SPTMBUD1.COLLECTION_CODE" & vbCrLf _
                    & "      and ICTCOLL1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                    & "    group by SPTMBUD1.CUST_CODE, SPTMBUD1.CUST_STORE_NO) SPTMBUDX" & vbCrLf _
                    & " where SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE " & vbCrLf _
                    & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE" & vbCrLf _
                    & "   and RSTRETLX.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
                    & "   and RSTRETLX.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
                    & "   and RSTBUDRX.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
                    & "   and RSTBUDRX.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
                    & "   and SPTCWRXX.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
                    & "   and SPTCWRXX.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
                    & "   and SPTMBUDX.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
                    & "   and SPTMBUDX.CUST_STORE_NO (+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
                    & "   and ARTCUST2.CUST_CODE = '" & CUST_CODE & "' " & vbCrLf _
                    & "   and SPTCWRXC.CHECKBOOK = '" & CHECKBOOK & "'" & vbCrLf _
                    & "   and ICTCOLL0.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE = ICTCOLL0.COLLECTION_CODE" & vbCrLf _
                    & "   and SATAUTH1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
                    & "   and SATAUTH1.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO" & vbCrLf _
                    & "   and SATAUTH1.HC_CODE = '" & HC_CODE & "'" & vbCrLf _
                    & "   and SATAUTH1.OPS_YYYYPP_OPENED is Not Null" & vbCrLf _
                    & "   and SATAUTH1.OPS_YYYYPP_CLOSED is Null"

                For Each rowData As DataRow In ASCDATA1.GetDataTable.Select("", "CUST_STORE_NO")
                    Dim CUST_STORE_NO As String = rowData.Item("CUST_STORE_NO") & ""

                    Row_Ctr += 1
                    For i As Integer = 0 To 9
                        oSheet.Cells(Start_Row + Row_Ctr, i).Value = rowData.Item(i)
                    Next
                    oSheet.Cells(Start_Row + Row_Ctr, 11).Value = rowData.Item("LY_RTL_ACT")
                    oSheet.Cells(Start_Row + Row_Ctr, 12).Value = rowData.Item("TY_RTL_BUD")

                    iITEM_CODE = 0
                    For Each ITEM_CODE As String In ITEM_CODEs
                        iITEM_CODE += 1
                        Dim ITEM_CODE_LY As String = ITEM_CODEs_LY(ITEM_CODE)
                        Dim rowItemData As DataRow = tblSIST.Rows.Find(New String() {ITEM_CODE_LY, CUST_CODE, CUST_STORE_NO})
                        If rowItemData IsNot Nothing Then
                            oSheet.Cells(Start_Row + Row_Ctr, 17 + (iITEM_CODE - 1) * 5 + 3).Value = rowItemData.Item("LY_SELL_IN")
                            oSheet.Cells(Start_Row + Row_Ctr, 17 + (iITEM_CODE - 1) * 5 + 4).Value = rowItemData.Item("LY_SELL_THRU")
                        End If

                        For Each Column_Formula As Integer In New Integer() {3, 4}
                            Dim CF As Integer = 17 + (iITEM_CODE - 1) * 5 + Column_Formula
                            oSheet.Cells(Start_Row + 1, CF).Formula = "=Sum(" & Excel_Cell0(Start_Row + 1 + 1, CF) & ":" & Excel_Cell0(Start_Row + Row_Ctr, CF) & ")"
                        Next
                    Next
                Next

                For Each Column_Formula As Integer In New Integer() {8, 9, 10, 11, 12, 13}
                    oSheet.Cells(Start_Row + 1, Column_Formula).Formula = "=Sum(" & Excel_Cell0(Start_Row + 1 + 1, Column_Formula) & ":" & Excel_Cell0(Start_Row + Row_Ctr, Column_Formula) & ")"
                Next

                For Each Column_Formula As Integer In New Integer() {10, 13, 14, 15, 16, 17}
                    'If Column_Formula = 17 Then
                    '    oSheet.Cells(0, Column_Formula).EntireColumn.NumberFormat = "#,##0%"
                    'Else
                    '    oSheet.Cells(0, Column_Formula).EntireColumn.NumberFormat = "$#,##0"
                    'End If
                    rangeCopyFrom = oSheet.Cells(Start_Row + 2, Column_Formula)
                    rangePaste_To = oSheet.Range(Start_Row + 3, Column_Formula, Start_Row + Row_Ctr, Column_Formula)
                    rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                Next

                Dim PROMO_FORMULA As String = ""
                iITEM_CODE = 0
                For Each ITEM_CODE As String In ITEM_CODEs
                    iITEM_CODE += 1

                    Dim RTL As String = Excel_Cell0(Start_Row - 8, 17 + (iITEM_CODE - 1) * 5 + 2, 2)
                    PROMO_FORMULA &= "+(" & Excel_Cell0(Start_Row + 2, 17 + (iITEM_CODE - 1) * 5 + 2) & "*" & RTL & "*" & "0.85" & ")"

                    For Each iC As Integer In New Integer() {1, 5}
                        Dim Column_Formula As Integer = 17 + (iITEM_CODE - 1) * 5 + iC
                        rangeCopyFrom = oSheet.Cells(Start_Row + 2, Column_Formula)
                        rangePaste_To = oSheet.Range(Start_Row + 3, Column_Formula, Start_Row + Row_Ctr, Column_Formula)
                        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                    Next
                Next

                Dim Column_Formula_all As Integer = 17 + (iITEM_CODE) * 5

                oSheet.Cells(Start_Row + 0, Column_Formula_all + 1).Value = "Promo $"
                oSheet.Cells(Start_Row + 0, Column_Formula_all + 2).Value = "Promotionality %"

                oSheet.Cells(Start_Row + 0, Column_Formula_all + 1).EntireColumn.NumberFormat = "$#,##0"
                oSheet.Cells(Start_Row + 0, Column_Formula_all + 2).EntireColumn.NumberFormat = "#,##0%"

                oSheet.Cells(Start_Row + 1, Column_Formula_all + 1).Formula = "=Sum(" & Excel_Cell0(Start_Row + 1 + 1, Column_Formula_all + 1) & ":" & Excel_Cell0(Start_Row + Row_Ctr, Column_Formula_all + 1) & ")"

 
                oSheet.Cells(Start_Row + 2, Column_Formula_all + 1).Formula = "=" & Mid(PROMO_FORMULA, 2)
                Dim FX As String = Excel_Cell0(Start_Row + 2, Column_Formula_all + 1) & "/" & Excel_Cell0(Start_Row + 2, 12)
                oSheet.Cells(Start_Row + 2, Column_Formula_all + 2).Formula = "=" & FX
                For LL As Integer = 1 To 2
                    rangeCopyFrom = oSheet.Cells(Start_Row + 2, Column_Formula_all + LL)
                    rangePaste_To = oSheet.Range(Start_Row + 3, Column_Formula_all + LL, Start_Row + Row_Ctr, Column_Formula_all + LL)
                    rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                    If LL = 2 Then
                        rangePaste_To = oSheet.Range(Start_Row + 1, Column_Formula_all + LL, Start_Row + 1, Column_Formula_all + LL)
                        rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)
                    End If
                Next

                For Each Column_Color As Integer In New Integer() {9, 12}
                    'oSheet.Cells.NumberFormat = "$###,###,##0"
                    range = oSheet.Cells(0, Column_Color, Start_Row + 3 + Row_Ctr, Column_Color)
                    range.Interior.Color = SpreadsheetGear.Colors.Yellow
                Next

                'For Each Column_Color As Integer In New Integer() {18, 23, 28}
                '    range = oSheet.Cells(0, Column_Color, Start_Row + 3 + Row_Ctr, Column_Color)
                '    range.Interior.Color = SpreadsheetGear.Colors.LightGray
                'Next
            Next

            oWB.Worksheets(0).Delete()
            oWB.Worksheets(0).Select()
            oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
            Show_Document(XLS_FILENAME)
        End If

        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Excel()


        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim rangeCopyFrom As SpreadsheetGear.IRange
        Dim rangePaste_To As SpreadsheetGear.IRange

        Dim XL_ROWS As Integer
        Dim XL_COLS As Integer

        ' Set up parameterized row and col settings

        XL_ROWS = dst.Tables("SOTORDF1").Rows.Count ' # of Rows in Store List
        XL_COLS = 0   ' # of numeric columns in Layout Selected

        ' Create Workbook

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Initializing Excel Objects", "")

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim FILENAME As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FILENAME = xls_path & "\" & xls_name & ".XLSx"

                If Not My.Computer.FileSystem.FileExists(FILENAME) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook()

        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i

        ' Add a worksheet for each row in the Report Summary

        Dim HCi As Integer = 0
        Dim Sheet1_used As Boolean = False

        For Each HC_CODE As String In HC_CODEs

            ASCMAIN1.Progress("-", HC_CODE)

            HCi += 1

            Dim ITEM_CODE_count As Integer = dst.Tables("ICTITEM1").Select("HC_CODE = '" & HC_CODE & "' and SEL = '1'").Length
            If ITEM_CODE_count <> 0 Then
                If Sheet1_used Then
                    oSheet = oWB.Worksheets.Add
                Else
                    oSheet = oWB.Sheets(0)
                    Sheet1_used = True
                End If

                'oSheet.Cells.Font.Name = "Verdana" ' "Times New Roman"
                oSheet.Name = HC_CODE & " Items"


                ' Load the DataTable into the Item Summary Sheet


                Dim TR As Integer = 1
                Dim TC As Integer = 3

                Load_DataTable_into_SGXLS(TR, TC, dst.Tables("ICTITEM1"), oSheet, grdICTITEM1, Nothing, "ITEM_CODE", "HC_CODE = '" & HC_CODE & "' and SEL = '1'")

                oSheet.Range(0, 0).EntireColumn.ColumnWidth = 10
                oSheet.Range(0, 1).EntireColumn.ColumnWidth = 25

                With oSheet.Cells(0, 0)
                    .Value = "Spread Order Worksheet"
                    .Font.Size = 14
                    .Font.Bold = True
                    .Font.Color = SpreadsheetGear.Colors.Purple
                End With

                Dim rowICTCOLL0 As DataRow = LookUp("ICTCOLL0", HC_CODE)
                Dim BRAND_CODE As String = rowICTCOLL0.Item("BRAND_CODE")
                Dim rowICTBRAN1 As DataRow = dst.Tables("ICTBRAN1").Rows.Find(BRAND_CODE)
                oSheet.Cells(1, 0).Value = BRAND_CODE
                oSheet.Cells(1, 1).Value = rowICTBRAN1.Item("BRAND_NAME")
                oSheet.Cells(2, 0).Value = CUST_CODE
                oSheet.Cells(2, 1).Value = Absx1.txtFor("CUST_NAME").Text
                oSheet.Cells(3, 0).Value = "Ship Date"
                'oSheet.Cells( 3, 1).Value = "?"
                oSheet.Cells(4, 0).Value = "Dept"
                'oSheet.Cells( 4, 1).Value = "?"

                '  oSheet.Cells(3, 4).Value = "High Collection " & HC_CODE & " Item List"

                oSheet = oWB.Worksheets.Add
                'oSheet.Cells.Font.Name = "Times New Roman"
                oSheet.Name = HC_CODE & " Stores"

                Dim RowStore As Integer = 17

                Dim grdfmt As New Dictionary(Of String, grdcolfmt)
                grdfmt_add(grdfmt, "CUST_STORE_NO", grdSOTORDF1)
                grdfmt_add(grdfmt, "CUST_STORE_NAME", grdSOTORDF1)
                grdfmt_add(grdfmt, "CUST_STORE_CITY", grdSOTORDF1)
                grdfmt_add(grdfmt, "CUST_STORE_STATE", grdSOTORDF1)
                grdfmt_add(grdfmt, "SELL_CODE", grdSOTORDF1)
                grdfmt_add(grdfmt, "SELL_NAME", grdSOTORDF1)
                grdfmt_add(grdfmt, "REGION_CODE", grdSOTORDF1)
                grdfmt_add(grdfmt, "REGION_DESC", grdSOTORDF1)
                grdfmt_add(grdfmt, "HC" & Format(HCi, "00") & "_RTL", grdSOTORDF1)
                grdfmt_add(grdfmt, "HC" & Format(HCi, "00") & "_PCT", grdSOTORDF1)
                Dim tbl_Stores As DataTable = Load_DataTable_into_SGXLS(RowStore, 1, dst.Tables("SOTORDF1"), oSheet, Nothing, grdfmt, "CUST_STORE_NO", "SEL = '1' AND HC" & Format(HCi, "00") & "_AUTH='1'")


                ' Total cells for Retail Sales and % to Total

                'worksheet.Cells[0, index].EntireColumn.NumberFormat = @"##0.00\%;[Red](##0.00\%)"

                With oSheet.Cells(RowStore - 1 - 1, 8)
                    .NumberFormat = "#,##0"
                    .Formula = "=SUM(" & Excel_Cell(RowStore + 1, 8 + 1) & ":" & Excel_Cell(RowStore + tbl_Stores.Rows.Count, 8 + 1) & ")"
                    .EntireColumn.ColumnWidth = 9
                End With
                With oSheet.Cells(RowStore - 1 - 1, 9)
                    .NumberFormat = "#,##0.0%"
                    .Formula = "=SUM(" & Excel_Cell(RowStore + 1, 9 + 1) & ":" & Excel_Cell(RowStore + tbl_Stores.Rows.Count, 9 + 1) & ")"
                    .EntireColumn.ColumnWidth = 6
                End With


                ' Total Qty Ordered Column

                With oSheet.Cells(RowStore - 1, 10)
                    .Value = "Total Qty"
                    .Interior.Color = SpreadsheetGear.Colors.LightGray
                    .NumberFormat = "@"
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Right
                End With

                With oSheet.Cells(RowStore - 1 + 1, 10)
                    .NumberFormat = "#,##0"
                    .Formula = "=SUM(" & Excel_Cell(RowStore - 1 + 1 + 1, 10 + 1 + 1) & ":" & Excel_Cell(RowStore - 1 + 1 + 1, 10 + ITEM_CODE_count + 1) & ")"
                    .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                End With

                If tbl_Stores.Rows.Count > 1 Then
                    rangeCopyFrom = oSheet.Range(Excel_Cell(RowStore - 1 + 1 + 1, 10 + 1) & ":" & Excel_Cell(RowStore - 1 + 1 + 1, 10 + 1))
                    rangePaste_To = oSheet.Range(Excel_Cell(RowStore - 1 + 1 + 1 + 1, 10 + 1) & ":" & Excel_Cell(RowStore - 1 + tbl_Stores.Rows.Count + 1, 10 + 1))
                    rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
                End If


                ' % to Total column - calculated using Total Retail Sales

                With oSheet.Cells(RowStore - 1 + 1, 9)
                    Dim Cell1 As String = Excel_Cell0(RowStore - 1 + 1, 8)
                    Dim Cell2 As String = Excel_Cell0(RowStore - 1 - 1, 8, 3)
                    .Formula = "=IF(" & Cell2 & "=0,0," & Cell1 & "/" & Cell2 & ")"
                    .NumberFormat = "#,##0.0%"
                End With

                If tbl_Stores.Rows.Count > 1 Then
                    rangeCopyFrom = oSheet.Range(RowStore - 1 + 1, 9)
                    rangePaste_To = oSheet.Range(RowStore - 1 + 2, 9, RowStore - 1 + tbl_Stores.Rows.Count, 9)
                    rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
                End If


                ' Order Qty Total

                Dim RowX As Integer = RowStore - 1 + 1 - 1
                Dim ColX As Integer = 10 + 1 + 1 - 1
                With oSheet.Cells(RowX - 4, ColX)
                    .NumberFormat = "#,##0"
                    .Formula = "=SUM(" & Excel_Cell0(RowX + 1, ColX) & ":" & Excel_Cell0(RowX + tbl_Stores.Rows.Count, ColX) & ")"
                    '.Interior.Color = SpreadsheetGear.Colors.LightGreen
                End With

                If ITEM_CODE_count > 1 Then
                    rangeCopyFrom = oSheet.Range(RowX - 4, ColX, RowX - 4, ColX)
                    rangePaste_To = oSheet.Range(RowX - 4, ColX + 1, RowX - 4, ColX + ITEM_CODE_count - 1)
                    rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.All, SpreadsheetGear.PasteOperation.None, False, False)
                End If

                With oSheet.Range(RowX - 4, ColX, RowX - 4, ColX + ITEM_CODE_count - 1)
                    .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                End With


                ' Border around Entry Area

                If tbl_Stores.Rows.Count > 0 Then
                    With oSheet.Range(RowX + 1, ColX - 1 + 1, RowX + tbl_Stores.Rows.Count, ColX - 1 + ITEM_CODE_count)
                        .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
                        .Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    End With
                End If



                ' Freeze Panes

                oSheet.Range(RowStore, 2).Select()
                oSheet.WindowInfo.FreezePanes = True
                oSheet.Range("A1:A1").Select()


                ' Setup Sheet Header Block

                oSheet.Range(0, 0).EntireColumn.ColumnWidth = 10
                oSheet.Range(0, 1).EntireColumn.ColumnWidth = 25

                With oSheet.Cells(0, 0)
                    .Value = "Spread Order Worksheet"
                    .Font.Size = 14
                    .Font.Bold = True
                    .Font.Color = SpreadsheetGear.Colors.Purple
                End With

                oSheet.Cells(1, 0).Value = BRAND_CODE
                oSheet.Cells(1, 1).Value = rowICTBRAN1.Item("BRAND_NAME")
                oSheet.Cells(2, 0).Value = CUST_CODE
                oSheet.Cells(2, 1).Value = Absx1.txtFor("CUST_NAME").Text
                oSheet.Cells(3, 0).Value = "Ship Date"
                'oSheet.Cells( 3, 1).Value = "?"
                oSheet.Cells(4, 0).Value = "Dept"
                'oSheet.Cells( 4, 1).Value = "?"


                ' Setup SKU Block

                Dim SKU_BLOCK As New List(Of String)
                SKU_BLOCK.Add("SKU|@")
                SKU_BLOCK.Add("EAN|@")
                SKU_BLOCK.Add("3PL Code|@")
                SKU_BLOCK.Add("Retail|$#,##0.00")
                SKU_BLOCK.Add("Description|@")
                SKU_BLOCK.Add("Notes|@")
                SKU_BLOCK.Add("Group|@")
                SKU_BLOCK.Add("Multiple|#,##0")
                SKU_BLOCK.Add("Minimum|#,##0")
                SKU_BLOCK.Add("Allocation|#,##0")
                SKU_BLOCK.Add("Shp+Opn+Pck|#,##0")
                SKU_BLOCK.Add("Balance|#,##0")
                SKU_BLOCK.Add("Order Qty|#,##0")
                SKU_BLOCK.Add("Remaining|#,##0")
                SKU_BLOCK.Add("Balance $Retail|$#,##0")
                SKU_BLOCK.Add("Balance $Net|$#,##0")

                Dim tbl As DataTable = New DataView(dst.Tables("ICTITEM1"), "HC_CODE = '" & HC_CODE & "' and SEL = '1'", "ITEM_CODE", DataViewRowState.CurrentRows).ToTable

                Dim CHI As Integer = 0
                Dim RHI As Integer = 10
                Dim CHX As Integer = -1
                For Each element As String In SKU_BLOCK
                    CHX += 1
                    oSheet.Cells(CHI + CHX, RHI).Value = Split(element, "|")(0)
                    oSheet.Range(CHI + CHX, RHI + 1, CHI + CHX, RHI + tbl.Rows.Count).NumberFormat = Split(element, "|")(1)
                Next

                With oSheet.Range(CHI + 0, RHI, CHI + CHX, RHI)
                    .Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
                    .NumberFormat = "@"
                    .Font.Color = SpreadsheetGear.Colors.Blue
                End With

                With oSheet.Range(CHI + 0, RHI + 1, CHI + 2, RHI + tbl.Rows.Count)
                    .HorizontalAlignment = SpreadsheetGear.HAlign.Center
                End With


                Dim RHX As Integer = RHI
                oSheet.Cells(CHX, RHX + 1).Formula = "=" & Excel_Cell(CHX + 9, RHX + 1) & "-" & Excel_Cell(CHX + 10, RHX + 1)
                Dim t1 As Integer = 5000
                For Each row As DataRow In tbl.Select("")
                    RHX += 1
                    CHX = CHI - 1
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = row.Item("ITEM_CODE")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = row.Item("ITEM_EAN_CODE")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = row.Item("ITEM_ALT_SORT")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = Val(row.Item("ITEM_RETAIL_PRICE") & "")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = row.Item("ITEM_DESC")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = row.Item("ALLO_NOTES")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = row.Item("ALLO_GROUP_CODE")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = Val(row.Item("ITEM_SO_QTY_MULT") & "")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = Val(row.Item("ITEM_SO_QTY_MIN") & "")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = Val(row.Item("QTY_ALLO") & "")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = Val(row.Item("ORDR_QTY_SHIP") & "") + Val(row.Item("ORDR_QTY_OPEN") & "") + Val(row.Item("ORDR_QTY_PICK") & "")
                    CHX += 1 : oSheet.Cells(CHX, RHX).Value = Val(row.Item("QTY_BAL") & "")
                    CHX += 2 : oSheet.Cells(CHX, RHX).Formula = "=" & Excel_Cell(CHI + 11 + 1, RHX + 1) & "-" & Excel_Cell(CHI + 12 + 1, RHX + 1) 'next 3 need to be adjusted if adding anything above
                    CHX += 1 : oSheet.Cells(CHX, RHX).Formula = "=" & Excel_Cell(CHI + 3 + 1, RHX + 1) & "*" & Excel_Cell(CHI + 11 + 1, RHX + 1)
                    CHX += 1 : oSheet.Cells(CHX, RHX).Formula = "=" & Excel_Cell(CHI + 14 + 1, RHX + 1) & "*" & 0.6
                    If row.Item("DATE_START") & "" <> "" And row.Item("DATE_END") & "" <> "" Then
                        oSheet.Cells(CHX + 1, RHX).Value = Format(row.Item("DATE_START"), "MM/dd") & "-" & Format(row.Item("DATE_END"), "MM/dd")
                    End If
                Next

                oSheet.Range(Excel_Cell0(-1, RHI) & ":" & Excel_Cell0(-1, RHX)).ColumnWidth = 13
                With oSheet.Range(0, RHI + 1, CHX, RHI + ITEM_CODE_count)
                    .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
                    .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
                End With

                oSheet.Range("C2:C2").Select()
                oSheet.WindowInfo.FreezePanes = True
                oSheet.Range("A1:A1").Select()
            End If
        Next

        oWB.Worksheets(0).Select()

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Get_Allocations()
        Dim DT As Date = dteAllocation.Value
        Dim DT2 As Date = dteAllocation2.Value
        Dim sqlSOTALLO1 As String = "Select * from SOTALLO1" & vbCrLf _
            & " where ITEM_CODE in (Select ITEM_CODE from (" & sqlICTITEM1_for_Allocations & "))" & vbCrLf _
            & "   and DATE_START <= '" & Format(DT2, "dd-MMM-yyyy") & "' and DATE_END >= '" & Format(DT, "dd-MMM-yyyy") & "'" & vbCrLf
        If ALLO_GROUP_CODEs.Count > 0 Then
            Dim SQLA As String = ""
            If ALLO_GROUP_CODEs.Contains("") Then
                SQLA = " OR ALLO_GROUP_CODE IS NULL"
            End If
            sqlSOTALLO1 &= "   and (NVL(ALLO_GROUP_CODE,'') in ('" & Join(ALLO_GROUP_CODEs.ToArray, "','") & "')" & SQLA & ")"
        End If
        '& "   and NVL(ALLO_GROUP_CODE,'') in ('" & Join(ALLO_GROUP_CODEs.ToArray, "','") & "')"
        Fill_Records("SOTALLO1", "", True, sqlSOTALLO1)

        ASCMAIN1.sql = "Select * from SOTALLO2" & vbCrLf _
            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ALLO_CTL_NO in (Select ALLO_CTL_NO from (" & sqlSOTALLO1 & "))"
        Fill_Records("SOTALLO2", "", True, ASCMAIN1.sql)

        For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("")

            Dim ITEM_CODE As String = rowSOTALLO1.Item("ITEM_CODE")
            Dim ALLO_CTL_NO As String = rowSOTALLO1.Item("ALLO_CTL_NO")

            Dim sqlw As String = "ITEM_CODE = '" & ITEM_CODE & "'"
            Dim rowICTITEM1s() As DataRow = dst.Tables("ICTITEM1").Select(sqlw, "ALLO_CTL_NO")
            If rowICTITEM1s.Length > 0 Then
                Dim rowICTITEM1 As DataRow = rowICTITEM1s(0)
                If rowICTITEM1.Item("ALLO_CTL_NO") & "" <> "" Then
                    Dim row2 As DataRow = dst.Tables("ICTITEM1").NewRow
                    row2.ItemArray = rowICTITEM1.ItemArray
                    row2.Item("ALLO_CTL_NO") = DBNull.Value
                    row2.Item("DATE_START") = DBNull.Value
                    row2.Item("DATE_END") = DBNull.Value
                    dst.Tables("ICTITEM1").Rows.Add(row2)
                    rowICTITEM1 = row2
                End If

                With rowICTITEM1
                    .Item("ALLO_CTL_NO") = rowSOTALLO1.Item("ALLO_CTL_NO")
                    .Item("DATE_START") = rowSOTALLO1.Item("DATE_START")
                    .Item("DATE_END") = rowSOTALLO1.Item("DATE_END")
                    .Item("ALLO_GROUP_CODE") = rowSOTALLO1.Item("ALLO_GROUP_CODE")
                    .Item("QTY_ALLO") = 0
                    Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
                    If rowSOTALLO2 IsNot Nothing Then
                        .Item("QTY_ALLO") = rowSOTALLO2.Item("QTY_ALLO")
                        .Item("ALLO_NOTES") = rowSOTALLO2.Item("ALLO_NOTES")
                    End If

                End With
            End If
        Next


        ASCMAIN1.sql = "Select ITEM_CODE, ALLO_CTL_NO" & vbCrLf _
            & ", SUM (ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from SOTORDR2" & vbCrLf _
            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ALLO_CTL_NO in (Select ALLO_CTL_NO from (" & sqlSOTALLO1 & "))" & vbCrLf _
            & " group by ITEM_CODE, ALLO_CTL_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "ITEM_CODE,ALLO_CTL_NO")
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")

            Dim sqlw As String = "ITEM_CODE = '" & ITEM_CODE & "' and ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
            Dim rowICTITEM1s() As DataRow = dst.Tables("ICTITEM1").Select(sqlw)
            If rowICTITEM1s.Length = 1 Then
                Dim rowICTITEM1 As DataRow = rowICTITEM1s(0)
                With rowICTITEM1
                    .Item("ORDR_QTY") = row.Item("ORDR_QTY")
                    .Item("ORDR_QTY_OPEN") = row.Item("ORDR_QTY_OPEN")
                    .Item("ORDR_QTY_PICK") = row.Item("ORDR_QTY_PICK")
                    .Item("ORDR_QTY_SHIP") = row.Item("ORDR_QTY_SHIP")
                    .Item("ORDR_QTY_CANC") = row.Item("ORDR_QTY_CANC")
                End With
            End If
        Next

        Dim dvw As DataView = dst.Tables("ICTITEM1").DefaultView
        dvw.RowFilter = "QTY_ALLO <> 0"
        Sort_grdColumns(grdICTITEM1, "ITEM_CODE")
    End Sub

    Private Sub UltraTabPageControl2_Paint(sender As Object, e As PaintEventArgs) Handles UltraTabPageControl2.Paint

    End Sub
End Class