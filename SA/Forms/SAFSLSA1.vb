Public Class SAFSLSA1
    Dim COLUMNS_sort() As String
    Dim TATSTAT1 As String
    Dim TATSTATX As String
    Dim SOTINVHX As String
    Dim sqlSOTINVHX As String
    Dim GMTCGMAX As String
    Dim sqlGMTCGMAX As String

    Dim COLUMN_NAMEs As New ArrayList
    Dim COLUMN_CAPTIONs As New ArrayList
    Dim COLUMN_NAME_by_Lvl() As String
    Dim COLUMN_CAPTION_by_Lvl() As String
    Dim G_by_Lvl() As Integer
    Dim tblASTDSQLA As DataTable
    Dim QCOLS As New Dictionary(Of String, String)
    Dim LVL As Int16

    Dim tblSATSLSW1 As DataTable

    Dim SCOPE() As String

    Dim YWP(,) As String
    Dim YWPD() As Date
    'Dim YWF(,) As String
    'Dim YWFD() As Date
    'Dim YWN(,) As String
    'Dim YWND() As Date
    Private WithEvents UltraTree_DropHightLight_DrawFilter As New UltraTree_DropHightLight_DrawFilter_Class()
    Dim US_STATES() As String
    ' Dim USmap As MapLayer

    Dim myCodes As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' These are the columns in the Data Block

        If MENU_ITEM_OBJECT = "SAFSLSA1" Then
            QCOLS.Add("TY_YTD_GRS", "TY YTD Grs")
            QCOLS.Add("TY_YTD_RTN", "TY YTD Rtn")
            QCOLS.Add("TY_YTD_NET", "TY YTD Net")
            QCOLS.Add("LY_YTD_GRS", "TY YTD Grs")
            QCOLS.Add("LY_YTD_RTN", "TY YTD Rtn")
            QCOLS.Add("LY_YTD_NET", "TY YTD Net")
            QCOLS.Add("TY_M01_GRS", "Jan")
            QCOLS.Add("TY_M02_GRS", "Feb")
            QCOLS.Add("TY_M03_GRS", "Mar")
            QCOLS.Add("TY_M04_GRS", "Apr")
            QCOLS.Add("TY_M05_GRS", "May")
            QCOLS.Add("TY_M06_GRS", "Jun")
            QCOLS.Add("TY_M07_GRS", "Jul")
            QCOLS.Add("TY_M08_GRS", "Aug")
            QCOLS.Add("TY_M09_GRS", "Sep")
            QCOLS.Add("TY_M10_GRS", "Oct")
            QCOLS.Add("TY_M11_GRS", "Nov")
            QCOLS.Add("TY_M12_GRS", "Dec")
        Else
            ASCMAIN1.sql = "Select COLUMN_NAME, COLUMN_CAPTION, COLUMN_SEQ from ASTDSQLS where FORM_NAME = '" & MENU_ITEM_OBJECT & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "COLUMN_SEQ")
                QCOLS.Add(row.Item("COLUMN_NAME"), row.Item("COLUMN_CAPTION") & "")
            Next
        End If

        ASCMAIN1.Get_Week_Range(-72, YWPD, YWP)
        'ASCMAIN1.Get_Week_Range(60, YWFD, YWF)
        'ASCMAIN1.Get_Week_Range(60, YWND, YWN, ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -52))

        With dst

            ' Main Query Table 

            ASCMAIN1.sql = ""
            For I As Int16 = 1 To 9
                ASCMAIN1.sql &= ",USER_NAME CODE" & CStr(I)
            Next
            For Each QCOL As String In QCOLS.Keys
                If QCOL.StartsWith("QTY_") Then
                    ASCMAIN1.sql &= ",1 " & QCOL
                Else
                    ASCMAIN1.sql &= ",0.01 " & QCOL
                End If
            Next
            ASCMAIN1.sql &= ",0.01 SORT_SEQ"
            ASCMAIN1.sql = "Select " & Mid(ASCMAIN1.sql, 2) & " FROM ASTUSER1 where ROWNUM < 1"
            TATSTAT1 = ASCMAIN1.Temp_Table
            TATSTATX = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select * from " & TATSTAT1
            Create_TDA(.Tables.Add, "TATSTAT1", "**", 0, False)
            .Tables("TATSTAT1").Columns.Add("DESC_VALUE")
            For I As Int16 = 1 To 9
                .Tables("TATSTAT1").Columns.Add("DESC" & CStr(I))
            Next
            '.Tables("TATSTAT1").Columns("QTY_AVA").Expression = "ISNULL(QTY_ONH,0) - ISNULL(QTY_COM,0) - ISNULL(QTY_PMH,0)  - ISNULL(QTY_QAH,0) - ISNULL(QTY_FDA,0)"

            ASCMAIN1.sql = "Select * from TATFAVE1"
            Create_TDA(.Tables.Add, "TATFAVE1", "**", 0)


            ' Detail Tables - need to make sure each table has the CODE fields

            'BRAND_CODE,COLLECTION_CODE,CUST_CODE,ITEM_BASIC_PROMO,ITEM_SNU_CODE,MARKET_CODE,SREP_CODE,TRADE_CLASS_CODE

            'Purchase Orders

            sqlSOTINVHX = "Select SOTINVH2.*" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
                & ", ICTITEM1.ITEM_SNU_CODE, SOTTCLS1.MARKET_CODE, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & " from SOTINVH2,ICTITEM1,ARTCUST1,SOTTCLS1,ICTCOLL1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE"
            SOTINVHX = ASCMAIN1.Temp_Table(sqlSOTINVHX & " and ROWNUM < 1")
            ASCMAIN1.sql = "Select * from " & SOTINVHX
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False)
            '.Tables("POTORDRX").Columns.Add("PO_CST_ORD", GetType(System.Decimal), "ISNULL(PO_QTY_ORD,0) * ISNULL(PO_COST,0)")
            '.Tables("POTORDRX").Columns.Add("PO_CST_OPN", GetType(System.Decimal), "ISNULL(PO_QTY_OPN,0) * ISNULL(PO_COST,0)")
            '.Tables("POTORDRX").Columns.Add("PO_RTL_ORD", GetType(System.Decimal), "ISNULL(PO_QTY_ORD,0) * ISNULL(RETAIL_PRICE,0)")
            '.Tables("POTORDRX").Columns.Add("PO_RTL_OPN", GetType(System.Decimal), "ISNULL(PO_QTY_OPN,0) * ISNULL(RETAIL_PRICE,0)")
            Add_Description_Columns("SOTINVHX")


            '' Inventory
            ''DEPT_CODE DGC_CODE DG_CODE MDSE_DIV_CODE OP_DIV_CODE OP_GRP_CODE REPT_DEPT_CODE SEASON_SEQ_NO STORE_NO

            'sqlGMTCGMAX = "Select GMTCGMB0.*" & vbCrLf _
            '    & ", SUBSTR(GMTCGMB0.DGC_CODE,1,1) DEPT_CODE, SUBSTR(GMTCGMB0.DGC_CODE,1,2) DG_CODE" & vbCrLf _
            '    & ", GMTREPT1.MDSE_DIV_CODE, GMTSTOR1.OP_DIV_CODE, POTORDR1.OP_GRP_CODE" & vbCrLf _
            '    & " from GMTCGMB0,GMTREPT1,GMTDGRP1" & vbCrLf _
            '    & " where GMTREPT1.REPT_DEPT_CODE = GMTDGRP1.REPT_DEPT_CODE" & vbCrLf _
            '    & "   and GMTDGRP1.DG_CODE = SUBSTR(GMTCGMB0.DGC_CODE,1,2)"
            'GMTCGMAX = ASCMAIN1.Temp_Table(sqlGMTCGMAX & " and ROWNUM < 1")
            'ASCMAIN1.sql = "Select * from " & GMTCGMAX
            'Create_TDA(.Tables.Add, "GMTCGMAX", "**", 0, False)
            ''.Tables("GMTCGMAX").Columns.Add("PO_CASES_AVA", GetType(System.Int32), "ISNULL(PO_CASES,0) - ISNULL(PO_CASES_PRESOLD,0)")
            'Add_Description_Columns("GMTCGMAX")


            With .Tables.Add("TATSTATC")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_NAME_CODE")
                .Columns.Add("COLUMN_NAME_DESC")
                .Columns.Add("TABLE_NAME_LOOKUP")
                .Columns.Add("ASTCODE1")
                .PrimaryKey = New DataColumn() {.Columns("COLUMN_NAME")}
            End With

            ASCMAIN1.sql = "Select USER_NAME COLUMN_NAME, USER_NAME CODE_VALUE, USER_NAME DESC_VALUE from ASTUSER1"
            Create_TDA(.Tables.Add, "TATSTATD", "**", 0, False, "", 2)
            .Tables("TATSTATD").Columns("DESC_VALUE").MaxLength = -1
        End With


        ' Description Sources and Work Tables
        'BRAND_CODE,COLLECTION_CODE,CUST_CODE,ITEM_BASIC_PROMO,ITEM_SNU_CODE,MARKET_CODE,SREP_CODE,TRADE_CLASS_CODE

        dst.Tables("TATSTATC").Rows.Clear()
        ' NEED TO GET THIS FROM THE REPORT DEFINITION AS WELL
        With dst.Tables("TATSTATC").Rows
            If MENU_ITEM_OBJECT = "SAFSLSA1" Then
                .Add(New String() {"BRAND_CODE", "BRAND_CODE", "BRAND_NAME", "ICTBRAN1"})
                .Add(New String() {"COLLECTION_CODE", "COLLECTION_CODE", "COLLECTION_NAME", "ICTCOLL1"})
                .Add(New String() {"CUST_CODE", "CUST_CODE", "CUST_NAME", "ARTCUST1"})
                .Add(New String() {"ITEM_BASIC_PROMO", "ITEM_BASIC_PROMO", "T_DESC", "ICTITEM1", "1"})
                .Add(New String() {"ITEM_SNU_CODE", "ITEM_SNU_CODE", "T_DESC", "ICTITEM1", "1"})
                .Add(New String() {"MARKET_CODE", "MARKET_CODE", "MARKET_DESC", "SOTMKTC1"})
                .Add(New String() {"SREP_CODE", "SREP_CODE", "SREP_NAME", "SOTSREP1"})
                .Add(New String() {"TRADE_CLASS_CODE", "TRADE_CLASS_CODE", "TRADE_CLASS_DESC", "SOTTCLS1"})
            ElseIf MENU_ITEM_OBJECT = "SAFCOMPA" Then
                .Add(New String() {"BRAND_CODE", "BRAND_CODE", "BRAND_NAME", "ICTBRAN1"})
                .Add(New String() {"COLLECTION_CODE", "COLLECTION_CODE", "COLLECTION_NAME", "ICTCOLL1"})
                .Add(New String() {"CUST_CODE", "CUST_CODE", "CUST_NAME", "ARTCUST1"})
                .Add(New String() {"CUST_STORE_GROUP", "CUST_STORE_GROUP", "CUST_STORE_GROUP_NAME", "ARTCUST8"})
                ' .Add(New String() {"CUST_STORE_LOCATION", "CUST_STORE_LOCATION", "CUST_STORE_LOCATION", "ARTCUSTM"})
                .Add(New String() {"CUST_STORE_NO", "CUST_STORE_NO", "CUST_STORE_NAME", "ARTCUST2"})
                .Add(New String() {"CUST_STORE_STATE", "STATE_CODE", "STATE_NAME", "TATSTATE"})
                '.Add(New String() {"CUST_STORE_ZIP_CODE", "ZIP", "ZIP", "ASTZDZIP"})
                .Add(New String() {"DEPT_CODE", "DEPT_CODE", "DEPT_DESC", "ICTDEPT1"})
                .Add(New String() {"DMA_CODE", "DMA", "NAME", "ASTZDDMA"})
                .Add(New String() {"HC_CODE", "HC_CODE", "HC_NAME", "ICTCOLL0"})
                .Add(New String() {"ITEM_BASIC_PROMO", "ITEM_BASIC_PROMO", "T_DESC", "ICTITEM1", "1"})
                .Add(New String() {"ITEM_CATGY_CODE", "ITEM_CATGY_CODE", "ITEM_CATGY_DESC", "ICTCATG1"})
                .Add(New String() {"ITEM_CLASS_CODE", "ITEM_CLASS_CODE", "ITEM_CLASS_DESC", "ICTCLAS1"})
                .Add(New String() {"ITEM_CODE", "ITEM_CODE", "ITEM_DESC", "ICTITEM1"})
                .Add(New String() {"ITEM_SNU_CODE", "ITEM_SNU_CODE", "T_DESC", "ICTITEM1", "1"})
                .Add(New String() {"MARKET_CODE", "MARKET_CODE", "MARKET_DESC", "SOTMKTC1"})
                .Add(New String() {"PROD_CODE", "PROD_CODE", "PROD_DESC", "ICTPROD1"})
                .Add(New String() {"SELL_CODE", "SELL_CODE", "SELL_NAME", "SOTSELL1"})
                .Add(New String() {"SREP_CODE", "SREP_CODE", "SREP_NAME", "SOTSREP1"})
                .Add(New String() {"TRADE_CLASS_CODE", "TRADE_CLASS_CODE", "TRADE_CLASS_DESC", "SOTTCLS1"})
            End If
        End With

        dst.Tables("TATSTATD").Rows.Clear()
        For Each rowTATSTATC As DataRow In dst.Tables("TATSTATC").Rows
            Dim COLUMN_NAME As String = rowTATSTATC.Item("COLUMN_NAME")
            Dim COLUMN_NAME_CODE As String = rowTATSTATC.Item("COLUMN_NAME_CODE")
            Dim COLUMN_NAME_DESC As String = rowTATSTATC.Item("COLUMN_NAME_DESC")
            Dim TABLE_NAME_LOOKUP As String = rowTATSTATC.Item("TABLE_NAME_LOOKUP") & ""
            Dim ASTCODE1 As String = rowTATSTATC.Item("ASTCODE1") & ""
            If ASTCODE1 = "1" Then
                Dim sqlPTW As String = " Where TABLE_NAME = '" & TABLE_NAME_LOOKUP & "'" _
                & " AND COLUMN_NAME = '" & COLUMN_NAME & "'"
                Fill_Records("TATSTATD", "", False, "Select '" & COLUMN_NAME & "' COLUMN_NAME, T_CODE CODE_VALUE, " _
                             & COLUMN_NAME_DESC & " DESC_VALUE from ASTCODE1" & sqlPTW)
            ElseIf TABLE_NAME_LOOKUP = "" Then
                Dim VL As ValueList = ASCMAIN1.ValueListFor(COLUMN_NAME)
                If VL IsNot Nothing Then
                    For Each VLI As ValueListItem In VL.ValueListItems
                        dst.Tables("TATSTATD").Rows.Add(New Object() {COLUMN_NAME, VLI.DataValue, VLI.DisplayText})
                    Next
                Else
                    Dim V As Dictionary(Of String, String) = CodeValues(COLUMN_NAME)
                    For Each DE As String In V.Keys
                        dst.Tables("TATSTATD").Rows.Add(New Object() {COLUMN_NAME, DE, V(DE)})
                    Next
                End If
            Else
                Fill_Records("TATSTATD", "", False, "Select '" & COLUMN_NAME & "' COLUMN_NAME, " & COLUMN_NAME_CODE & " CODE_VALUE, " & COLUMN_NAME_DESC & " DESC_VALUE from " & TABLE_NAME_LOOKUP)
            End If
        Next


        Absc1.grdSetup.DisplayLayout.Bands(0).Columns("GROUP_ALL_OTHERS").Hidden = True
        grdTATSTAT1.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show

        grdTATFAVE1.DataSource = dst.Tables("TATFAVE1")
        grdTATSTAT1.DataSource = dst.Tables("TATSTAT1")

        grdPOTORDRX.DataSource = dst.Tables("SOTINVHX")
        ' grdGMTCGMAX.DataSource = dst.Tables("GMTCGMAX")

        'Dim dvw As DataView = dst.Tables("SATCSLSS").DefaultView
        'dvw.RowFilter = "SALES <> 0"
        'grdSATCSLSS.DataSource = dvw

        'Me.UltraChart1.Data.DataBind()

        Format_grdTATSTAT1()

        'Create_Summary(grdSATCSLSS, "STATE_CODE", "Count")
        'Create_Summary(grdSATCSLSS, "SALES")

        Create_Summary(grdTATSTAT1, "DESC_VALUE", "Count")

        Create_Summary(grdPOTORDRX, "INV_NO", "Count")
        Create_Summary(grdPOTORDRX, New String() {"ORDR_QTY_SHIP"})
        'Create_Summary(grdGMTCGMAX, "STYLE_CODE", "Count")
        'Create_Summary(grdGMTCGMAX, New String() {"ON_HAND_UNITS_S", "ON_HAND_UNITS_W"})

        ' grdPOTORDRX.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

        For Each rowASTDSQLA As DataRow In Absc1.tblASTDSQLA.Rows
            Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                {grdPOTORDRX, grdGMTCGMAX}

                If (grd.DisplayLayout.Bands(0).Columns.Exists(COLUMN_NAME)) Then
                    With grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                        .CellAppearance.BackColor = Drawing.Color.Yellow
                        If New String() {"PROD_CODE", "SIZE_CODE", "PACK_CODE", "WHSE_CODE"}.Contains(COLUMN_NAME) Then
                            Dim COLUMN_NAME_DESC As String = Replace(COLUMN_NAME, "_CODE", "_DESC")
                            ' .Hidden = True

                            With grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME_DESC)
                                If COLUMN_NAME_DESC = "SIZE_DESC" Or COLUMN_NAME_DESC = "PACK_DESC" Then
                                    .Width = 50
                                Else
                                    .Width = 60
                                End If
                                .CellAppearance.BackColor = Drawing.Color.Yellow
                                '.Header.VisiblePosition = grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.VisiblePosition
                            End With
                        End If
                    End With
                End If
            Next
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                {grdPOTORDRX, grdGMTCGMAX}
            For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                If gcol.DataType.Equals(GetType(System.DateTime)) Then
                    gcol.Width = 80
                    gcol.Format = "MM/dd/yy"
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit

                End If
            Next
        Next

        ' ASCMAIN1.Add_Value_List(grdSOTINVH0, "CONTROLLED")

        ASCMAIN1.sql = "Select SALES_DIVISION_CODE, SALES_DIVISION_NAME from SOTSDIV1"
        Dim tbl As DataTable = ASCDATA1.GetDataTable()
        tbl.PrimaryKey = New DataColumn() {tbl.Columns(0)}
        cbeDIVISION_CODE.DataSource = tbl
        cbeDIVISION_CODE.DisplayMember = tbl.Columns(0).ColumnName
        cbeDIVISION_CODE.DisplayMember = tbl.Columns(1).ColumnName
        cbeDIVISION_CODE.SelectedItem = cbeDIVISION_CODE.Items(0)

        optCATEGORY_CODE.ValueList = ASCMAIN1.ValueListFor _
            ("", , New String() {":", "*:All"}, "Select * from (Select BRAND_CODE, INITCAP(BRAND_NAME) BRAND_NAME from ICTBRAN1 union Select '*' BRAND_CODE, 'All Brands' BRAND_NAME from Dual) order by 1")

        tab1.Tabs("Inquiry").Visible = False

        tvwSEQ.DrawFilter = UltraTree_DropHightLight_DrawFilter
        tvwSEQ.Override.SelectionType = UltraWinTree.SelectType.ExtendedAutoDrag

        With tvwSEQ
            .Appearances.Add("DropHighLightAppearance")
            With .Appearances("DropHighLightAppearance")
                .BackColor = System.Drawing.Color.Cyan
            End With

            .Override.SelectionType = UltraWinTree.SelectType.ExtendedAutoDrag

            .Override.CellClickAction = UltraWinTree.CellClickAction.Default
            .ViewStyle = UltraWinTree.ViewStyle.Standard
            .AllowDrop = True
            .Override.AllowCut = DefaultableBoolean.False
            .Override.ActiveNodeAppearance.BackColor = Drawing.Color.Green ' Blue
            .Override.ActiveNodeAppearance.ForeColor = Drawing.Color.White
        End With

        With chtSATCSLS1
            .Axis.X.ScrollScale.Visible = True
            .Axis.Y.ScrollScale.Visible = True

            .Axis.X.ScrollScale.Scale = 1 ' 0.25
            .Axis.Y.ScrollScale.Scale = 1 ' 0.25
            Me.trkbrXAxis.Value = .Axis.X.ScrollScale.Scale * 100
            Me.trkbrYAxis.Value = .Axis.Y.ScrollScale.Scale * 100
            .EnableCrossHair = True

            '.ColorModel.ModelStyle = ColorModels.CustomLinear '  CType(System.Enum.Parse(GetType(ColorModels), System.Enum.GetNames(GetType(ColorModels))(0)), ColorModels)
        End With

        Dim modelStyle As String() = System.Enum.GetNames(GetType(Infragistics.UltraChart.Shared.Styles.ColorModels))
        cbeColor.DataSource = modelStyle
        cbeColor.SelectedItem = cbeColor.Items(cbeColor.FindString(System.Enum.GetName(GetType(Infragistics.UltraChart.Shared.Styles.ColorModels), chtSATCSLS1.ColorModel.ModelStyle), 0))

        Check_Inquiry_Mode()

        Dim rowTATUSER1 As DataRow = LookUp("TATUSER1", ASCMAIN1.USER_ID, True)
        For Each dc As DataColumn In rowTATUSER1.Table.Columns
            myCodes.Add(dc.ColumnName, rowTATUSER1.Item(dc.ColumnName) & "")
        Next

        grdTATSTAT1.DisplayLayout.Override.ActiveRowAppearance.BackColor = System.Drawing.Color.LightGreen

        'If myCodes("DIVISION_CODE") <> "" Then
        '    cbeDIVISION_CODE.Value = myCodes("DIVISION_CODE")
        'End If

        optCATEGORY_CODE.Tag = "*"
    End Sub

    Sub Check_Inquiry_Mode()

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Generate"
                tblASTDSQLA = Absc1.grdSetupDataSource
                COLUMNS_sort = Absc1.grdSetupCOLUMNs
                If COLUMNS_sort Is Nothing Then
                    EMsg &= vbCr & "You Must Select Columns to Sort By"
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

            Case "Generate"

                EntryMode = "N"
                Load_Record()
                'Mode_Settings(True)

                optView.ValueList.ValueListItems.Item(2).Appearance.ForeColor = System.Drawing.Color.Gray
                If optView.Value = "G" Then
                    optView.Value = "S"
                End If

            Case "Clear"
                Clear_Settings()

            Case "Done"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Generate").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Clear").Settings.Enabled = not_iScreenMode
            .Groups("View").Visible = False
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        'dst.Tables("BATINVH0").Rows.Clear()
        EnforceConstraints(True)

        optCATEGORY_CODE.CheckedIndex = -1
        optCATEGORY_CODE.Value = "*"

        Setup_tab1()
        Setup_tabDetails()
        Setup_View()
        tabDetails.Tabs("Map").Visible = False '  resolve program exception

        If optCATEGORY_CODE.Tag = "*" Then
            Load_TATFAVE1()
            optCATEGORY_CODE.Tag = ""
        End If

    End Sub

    Sub Load_Record()

        Call Save_Header_Fields(UltraGroupBox1)

        ASCMAIN1.Progress("Now Loading Data")
        Me.Cursor = Cursors.WaitCursor

        EnforceConstraints(False)

        If EntryMode = "N" Then
        Else
        End If

        EnforceConstraints(True)

        Setup_DQ()

        Generate_Inquiry(True)
        tab1.SelectedTab = tab1.Tabs("Inquiry")

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Update_Record()

        BeginTrans()
        'INIT_LAST("BATINVH0", False, String.Empty, True)
        'Update_Record_TDA("BATINVH0")
        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Windows.Forms.Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "GROUP_NO"
                'sql_where = "GROUP_NO in (Select Distinct GROUP_NO from ICTPCAT1)"
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "Load"
                Dim BATCH_NO As String = Split(key, ":")(0)
                Absx1.txtFor("BATCH_NO").Text = BATCH_NO
                Click_Command(command)
        End Select

        Return return_key
    End Function
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(tvwDQ, "S", "Show Codes")
        Load_Popup_Menu(grdPOTORDRX, "SSBBS", "Show Filter", "Show GroupBox", "PO Inquiry")
        Load_Popup_Menu(grdGMTCGMAX, "SSBBS", "Show Filter", "Show GroupBox", "Style Inquiry")
        Load_Popup_Menu(grdSATCSLSS, "CC", "Best", "Worst")
        Load_Popup_Menu(grdTATSTAT1, "SSSSS", "Show Filter", "Show GroupBox", "Show Pins", _
                        "Show Details", "Show All Levels")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If e.SourceControl.Name = "tvwDQ" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Details") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Details"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Tag = "N"
            tlb_sbt.Checked = Not splInquiry.Panel2Collapsed
            tlb_sbt.Tag = ""
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdTATSTAT1"
                    'currentGridRow = dst.Tables("TATSTAT1").Rows.Find(grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text)
            End Select
        End If

    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            If GRDs.ContainsKey(Mid(e.Tool.OwningMenu.Key, 4)) Then
                grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
            End If
        End If

        Select Case e.Tool.Key

            Case "Show Codes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Codes(tlb_sbt.Checked)
                Exit Sub

            Case "Show All Levels"

                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag = "X" Then Exit Sub

                grdTATSTAT1.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
                'UltraExplorerBar1.Groups("View").Visible = Not tlb_sbt.Checked

                Click_Node(tvwDQ.ActiveNode)
                Exit Sub

            Case "Show Details"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "N" Then
                    splInquiry.Panel2Collapsed = Not splInquiry.Panel2Collapsed

                    If Not splInquiry.Panel2Collapsed Then
                        Setup_Details()
                        'splInquiry.Panel2Collapsed = False
                    Else
                        'UltraExplorerBar1.Groups("Maintain Lots").Visible = False
                    End If
                End If

            Case "Best"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorEnd

            Case "Worst"
                Dim tlb_cpt As UltraWinToolbars.PopupColorPickerTool = DirectCast(e.Tool, UltraWinToolbars.PopupColorPickerTool)
                tlb_cpt.ReplaceableColor = Me.UltraChart1.ColorModel.ColorBegin
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")
        End Select
    End Sub

    Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)

        If e.Tool.OwningMenu Is Nothing Then Exit Sub

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Best"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorEnd = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATCSLSS.DataBind()
                Application.DoEvents()
                grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)

            Case "Worst"
                Dim tlb_cpt As Infragistics.Win.UltraWinToolbars.PopupColorPickerTool _
                = DirectCast(e.Tool, Infragistics.Win.UltraWinToolbars.PopupColorPickerTool)
                Me.UltraChart1.ColorModel.ColorBegin = tlb_cpt.SelectedColor
                UltraChart1.DataBind()
                'grdSATCSLSS.DataBind()
                Application.DoEvents()
                grdSATCSLSS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
        End Select
    End Sub

    Private Sub tlb_AfterToolCloseup(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolDropdownEventArgs) Handles tlb.AfterToolCloseup
        Select Case e.SourceControl.Name
            Case "grdTATSTAT1"

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Clear")
                    Click_Command("Generate")
                End If
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "GROUP_NO"
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "PROM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "GROUP_NO"
            'If InquiryMode Then Click_Command("Load") Else Click_Command("Edit")
        End Select
    End Sub
#End Region

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        If Absx1.txtFor("SET_DESC").Text = "" Then
            MsgBox("You must enter a Description in order to Save", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        If False Then
            MsgBox("You Must Select at least 1 Code in order to Save", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        Else
            Dim SET_ID As String = grpFavorite.Tag
            Dim rowTATFAVE1 As DataRow = dst.Tables("TATFAVE1").Rows.Find(SET_ID)

            If SET_ID <> "" Then
                Dim iresponse As Int16 = MsgBox("Favorite " & "" & " already exists." & vbCrLf & vbCrLf & "Do you want to create a New Favorite?" & vbCrLf & vbCrLf & "Yes = Create a New Favorite" & vbCrLf & "No = Over-write existing Favorite" & vbCrLf & "Cancel = Cancel Save Request", MsgBoxStyle.YesNoCancel, "Favorite already exists")
                If iresponse = MsgBoxResult.Cancel Then
                    Exit Sub
                ElseIf iresponse = MsgBoxResult.Yes Then
                    SET_ID = ASCMAIN1.Next_Control_No("TATFAVE1.SET_ID")
                    rowTATFAVE1 = Nothing
                Else
                    If rowTATFAVE1 Is Nothing Then
                        Fill_Records("TATFAVE1", , , "SELECT * FROM TATFAVE1 WHERE SET_ID = '" & SET_ID & "'")
                        rowTATFAVE1 = dst.Tables("TATFAVE1").Rows.Find(SET_ID)
                    End If
                    If rowTATFAVE1 Is Nothing Then
                        SET_ID = ASCMAIN1.Next_Control_No("TATFAVE1.SET_ID")
                    End If
                End If
            Else
                SET_ID = ASCMAIN1.Next_Control_No("TATFAVE1.SET_ID")
            End If

            Dim t As DataTable = Absc1.grdSetupDataSource
            Dim DATASOURCE As String = ""
            For Each row As DataRow In t.Rows
                For i As Int16 = 0 To t.Columns.Count - 1
                    DATASOURCE &= vbTab & row.Item(i)
                Next
                DATASOURCE &= vbCrLf
            Next

            If rowTATFAVE1 Is Nothing Then
                rowTATFAVE1 = dst.Tables("TATFAVE1").NewRow()
                With rowTATFAVE1
                    .Item("SET_ID") = SET_ID
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                End With
                dst.Tables("TATFAVE1").Rows.Add(rowTATFAVE1)
            Else
                Dim row As DataRow = LookUp("TATFAVE1", SET_ID)
                If row Is Nothing Then
                    MsgBox("Favorite No Longer Exists in Database - Try Saving as New", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If
            End If
            With rowTATFAVE1
                .Item("SET_DESC") = Absx1.txtFor("SET_DESC").Text
                .Item("DIVISION_CODE") = cbeDIVISION_CODE.Value & ""
                .Item("SET_PUBLIC") = IIf(Absx1.chkFor("SET_PUBLIC").Checked, "1", "0")
                .Item("CATEGORY_CODE") = optCATEGORY_CODE.Value & ""
                .Item("DATASOURCE") = DATASOURCE
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                .Item("PROGRAM_CODE") = Absx1.txtFor("PROGRAM_CODE").Text
            End With
            Update_Record_TDA("TATFAVE1")

            Load_TATFAVE1()
            Load_Favorite(SET_ID)
        End If
    End Sub

    Sub Load_TATFAVE1()
        If IsLoading Then Exit Sub
        If SELECTION_NO = 0 Then Exit Sub

        If optCATEGORY_CODE.Value & "" = "" Then
            optCATEGORY_CODE.Value = "*"
        End If

        Dim Sql As String = " AND DIVISION_CODE = '" & cbeDIVISION_CODE.Value & "'"

        If optCATEGORY_CODE.Value <> "*" And optCATEGORY_CODE.Value <> "" Then
            Sql &= " AND CATEGORY_CODE = '" & optCATEGORY_CODE.Value & "'"
        End If

        Sql = "Select * from TATFAVE1 " _
        & " where (SET_PUBLIC = '1' or INIT_OPER = '" & ASCMAIN1.USER_ID & "')" _
        & Sql

        Fill_Records("TATFAVE1", , , Sql)
        Sort_grdColumns(grdTATFAVE1, "SET_DESC")
    End Sub

    Private Sub optCATEGORY_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCATEGORY_CODE.ValueChanged
        Load_TATFAVE1()
    End Sub

    Private Sub cbeDIVISION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeDIVISION_CODE.ValueChanged
        Load_TATFAVE1()
    End Sub

    Private Sub grdTATFAVE1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdTATFAVE1.AfterRowsDeleted
        Update_Record_TDA("TATFAVE1")
    End Sub

    Private Sub grdTATFAVE1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdTATFAVE1.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("INIT_OPER").Text <> ASCMAIN1.USER_ID Then
                MsgBox("You cannot delete Favorites which you did not Originate", MsgBoxStyle.OkOnly, "Cannot Perform Reqeusted Action")
                e.Cancel = True
                Exit For
            End If
        Next
    End Sub

    Private Sub grdTATFAVE1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdTATFAVE1.ClickCellButton
        Load_Favorite(e.Cell.Row.Cells("SET_ID").Value & "")
    End Sub

    Sub Setup_DQ()

        ' COLUMNS_sort contains the columns and the sequence of those columns from ABSC1
        ' we don't use COLUMNS_sort for anything else after we read its contents in this procedure
        ' COLUMN_NAMEs are the Columns from COLUMNS_sort, in index order
        ' COLUMN_CAPTIONs are the captions for the Columns from COLUMNS_sort, in index order
        ' COLUMN_NAMEs and COLUMN_CAPTIONs represent the original order and content from COLUMNS_sort
        '  and represent the context for CODE1,CODE2,.. in the temp table

        COLUMN_NAMEs.Clear()
        COLUMN_CAPTIONs.Clear()
        For Each COLUMN_NAME As String In COLUMNS_sort
            COLUMN_NAMEs.Add(COLUMN_NAME)
            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(COLUMN_NAME)
            COLUMN_CAPTIONs.Add(rowASTDSQLA.Item("COLUMN_CAPTION"))
        Next

        tvwSEQ.Nodes.Clear()
        'Dim anode_parent As UltraWinTree.UltraTreeNode = Nothing

        For i As Integer = 0 To COLUMN_NAMEs.Count - 1
            Dim anode As New UltraWinTree.UltraTreeNode
            anode.Text = COLUMN_CAPTIONs(i)
            anode.Key = COLUMN_NAMEs(i)
            tvwSEQ.Nodes.Add(anode)
        Next

        tvwSEQ.ExpandAll()
    End Sub

    Sub Generate_Inquiry(ByVal refresh_data As Boolean)

        ReDim COLUMN_NAME_by_Lvl(COLUMN_NAMEs.Count)
        ReDim COLUMN_CAPTION_by_Lvl(COLUMN_NAMEs.Count)
        ReDim G_by_Lvl(COLUMN_NAMEs.Count)
        ReDim SCOPE(COLUMN_NAMEs.Count)
        For G As Integer = 1 To COLUMN_NAMEs.Count
            Dim tnode As UltraWinTree.UltraTreeNode = tvwSEQ.GetNodeByKey(COLUMN_NAMEs(G - 1))
            Dim Lvl As Integer = tnode.Index + 1 ' tnode.Level + 1
            COLUMN_NAME_by_Lvl(Lvl) = COLUMN_NAMEs(G - 1)
            COLUMN_CAPTION_by_Lvl(Lvl) = tnode.Text
            G_by_Lvl(Lvl) = G
        Next

        If refresh_data Then

            ' Build Main Status Grid Data

            Absc1.Get_SQL("*")

            Dim sql_cols2 As String = ""
            Dim COLUMN_NAMEs_tree As String = ""
            Dim COLUMN_NAMEs_ordered_orig As String = ""
            Dim TABLE_NAMEs() As String = New String() {"ICTLOTD2", "POTORDR2", "ICTLOTD2"}
            For G As Integer = 1 To COLUMN_NAMEs.Count
                COLUMN_NAMEs_ordered_orig &= "CODE" & CStr(G) & ","
                COLUMN_NAMEs_tree &= "{TABLE_NAME}." & COLUMN_NAMEs(G - 1) & ","
                sql_cols2 &= Absc1.sql_SELECT_COLUMNs(G - 1) & ","
            Next

            ASCDATA1.ExecuteSQL("Truncate Table " & TATSTAT1)
            ASCDATA1.ExecuteSQL("Truncate Table " & TATSTATX)

            Dim sqlCols As String = ""
            For Each COLUMN_NAME In QCOLS.Keys
                sqlCols &= ", " & Absc1.sql_TABLE_NAME & "." & COLUMN_NAME
            Next
            sqlCols = Mid(sqlCols, 2)

            Dim sqlCols3 As String = ""
            If MENU_ITEM_OBJECT = "SAFSLSA1" Then
                Dim XX As String = "NVL(SOTINVH2.ORDR_QTY_SHIP, 0) * NVL(SOTINVH2.ORDR_UNIT_PRICE, 0)"
                sqlCols3 &= ", CASE WHEN SOTINVH2.ORDR_YYYYPP_UPDATED LIKE '2015%' AND SOTINVH2.INV_TYPE = 'I' THEN " & XX & " ELSE 0 END TY_YTD_GRS" & vbCrLf
                sqlCols3 &= ", CASE WHEN SOTINVH2.ORDR_YYYYPP_UPDATED LIKE '2015%' AND SOTINVH2.INV_TYPE = 'C' THEN " & XX & " ELSE 0 END TY_YTD_RTN" & vbCrLf
                sqlCols3 &= ", CASE WHEN SOTINVH2.ORDR_YYYYPP_UPDATED LIKE '2015%' THEN " & XX & " ELSE 0 END TY_YTD_NET" & vbCrLf
                sqlCols3 &= ", CASE WHEN SOTINVH2.ORDR_YYYYPP_UPDATED LIKE '2014%' AND SOTINVH2.INV_TYPE = 'I' THEN " & XX & " ELSE 0 END LY_YTD_GRS" & vbCrLf
                sqlCols3 &= ", CASE WHEN SOTINVH2.ORDR_YYYYPP_UPDATED LIKE '2014%' AND SOTINVH2.INV_TYPE = 'C' THEN " & XX & " ELSE 0 END LY_YTD_RTN" & vbCrLf
                sqlCols3 &= ", CASE WHEN SOTINVH2.ORDR_YYYYPP_UPDATED LIKE '2014%' THEN " & XX & " ELSE 0 END LY_YTD_NET" & vbCrLf
                For P As Integer = 1 To 12
                    sqlCols3 &= ", CASE WHEN SOTINVH2.ORDR_YYYYPP_UPDATED = '2014" & Format(P, "00") & "' THEN " & XX & " ELSE 0 END TY_M" & Format(P, "00") & "_GRS" & vbCrLf
                Next
                sqlCols3 = Mid(sqlCols3, 2)

            ElseIf MENU_ITEM_OBJECT = "SAFCOMPA" Then
                sqlCols3 = "1 RANK_LY,3 RANK_2Y,3 FULL_LY,3 FULL_2Y,3 PCT_LY2Y,3 RANK_TY,3 YTD_TY,33 YTD_LY,33 PCT_TYLY,3 GAP,333 TY_DOORS,3 LY_DOORS,3 PCT_TOT,3 PCT_CUM"
            End If

            ASCMAIN1.sql = "Insert into " & TATSTAT1 & vbCrLf _
                & " (" & COLUMN_NAMEs_ordered_orig & vbCrLf _
                & Replace(sqlCols, Absc1.sql_TABLE_NAME & ".", "") & ")" & vbCrLf _
                & " Select " & sql_cols2 & vbCrLf _
                & sqlCols3 & vbCrLf _
                & " from " & Absc1.sql_TABLE_NAME & Absc1.sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(Absc1.sql_JOIN & Absc1.sql_WHERE)

            ASCDATA1.ExecuteSQL()

            ' Purchase Orders

            'Absc1.Get_SQL("P")
            'ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVHX)
            'ASCMAIN1.sql = "Insert into " & SOTINVHX & " " & sqlSOTINVHX & Absc1.sql_WHERE
            'ASCDATA1.ExecuteSQL()

            '' Styles

            'Absc1.Get_SQL("H")
            'ASCDATA1.ExecuteSQL("Truncate Table " & GMTCGMAX)
            'ASCMAIN1.sql = "Insert into " & GMTCGMAX & " " & sqlGMTCGMAX & Absc1.sql_WHERE
            'ASCDATA1.ExecuteSQL()

            Dim C As String = ""
            For Each COLUMN_NAME In Split(COLUMN_NAMEs_tree, ",")
                If COLUMN_NAME <> "" Then
                    C &= COLUMN_NAME & ","
                End If
            Next
        Else
            tblSATSLSW1 = Nothing
        End If


        tblSATSLSW1 = Nothing
        'If InStr(Sql, COL) <> 0 Then
        '    ASCMAIN1.sql = "Select " & COL & vbCrLf & sql2 & Mid(Sql, P2, P3 - P2) & " group by " & COL
        '    tblSATSLSW1 = ASCDATA1.GetDataTable
        'End If
        Dim sql2 As String = "" _
                            & ",SUM (TY_M01_GRS) TY_M01_GRS" _
                            & ",SUM (TY_M02_GRS) TY_M02_GRS" _
                            & ",SUM (TY_M03_GRS) TY_M03_GRS" _
                            & ",SUM (TY_M04_GRS) TY_M04_GRS" _
                            & ",SUM (TY_M05_GRS) TY_M05_GRS" _
                            & ",SUM (TY_M06_GRS) TY_M06_GRS" _
                            & ",SUM (TY_M07_GRS) TY_M07_GRS" _
                            & ",SUM (TY_M08_GRS) TY_M08_GRS" _
                            & ",SUM (TY_M09_GRS) TY_M09_GRS" _
                            & ",SUM (TY_M10_GRS) TY_M10_GRS" _
                            & ",SUM (TY_M11_GRS) TY_M11_GRS" _
                            & ",SUM (TY_M12_GRS) TY_M12_GRS"
        If MENU_ITEM_OBJECT = "SAFCOMPA" Then
            sql2 = ", SUM(FULL_2Y) FULL_2Y, SUM (FULL_LY) FULL_LY, SUM (YTD_TY) YTD_TY"
        End If
        Dim COL As String = "CODE1"

        ASCMAIN1.sql = "Select " & COL & vbCrLf & sql2 & " from " & TATSTAT1 & " group by " & COL
        tblSATSLSW1 = ASCDATA1.GetDataTable
        If tblSATSLSW1 IsNot Nothing Then
            For I As Int16 = tblSATSLSW1.Rows.Count - 1 To 0 Step -1
                Dim row As DataRow = tblSATSLSW1.Rows(I)
                If row.Item(0) & "" = "" Then
                    row.Delete()
                End If
            Next
            tblSATSLSW1.PrimaryKey = New DataColumn() {tblSATSLSW1.Columns(0)}
        End If

        ' Show Grid
        tab1.Tabs("Inquiry").Visible = True
        If txtSET_DESC.Text = "" Then
            tab1.Tabs("Inquiry").Text = "Custom Inquiry"
        Else
            tab1.Tabs("Inquiry").Text = txtSET_DESC.Text
        End If
        Me.Text = Me.MENU_ITEM_DESC & " - " & tab1.Tabs("Inquiry").Text
        ASCMAIN1.Set_Form_Caption_on_Tab(Me)

        Application.DoEvents()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Collections into Selection Tree")

        With tvwDQ
            Dim rootColumnSet As UltraWinTree.UltraTreeColumnSet = .ColumnSettings.RootColumnSet
            rootColumnSet.Columns.Clear()
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count
                Dim column As UltraWinTree.UltraTreeNodeColumn = rootColumnSet.Columns.Add(COLUMN_NAME_by_Lvl(Lvl))
            Next
        End With

        Dim COLUMN_NAMEs_ordered As String = ""
        Dim CODE_COLUMNs_ordered As String = ""
        Dim ORDER_BY As String = ""
        For Lvl As Integer = 1 To COLUMN_NAMEs.Count
            COLUMN_NAMEs_ordered &= "," & COLUMN_NAME_by_Lvl(Lvl)
            CODE_COLUMNs_ordered &= ",CODE" & CStr(G_by_Lvl(Lvl))

            If COLUMN_NAME_by_Lvl(Lvl) = "SIZE_CODE" Then
                ASCMAIN1.sql = "Update " & TATSTAT1 & " TATSTAT1" _
                & " Set TATSTAT1.SORT_SEQ = (SELECT ICTSIZE1.RANK1 + ICTSIZE1.RANK2/1000000 from ICTSIZE1 " _
                & " where SIZE_CODE = " & "TATSTAT1.CODE" & CStr(G_by_Lvl(Lvl)) & ")"
                ASCDATA1.ExecuteSQL()
                ORDER_BY &= ",SORT_SEQ,CODE" & CStr(G_by_Lvl(Lvl))
            Else
                ORDER_BY &= ",CODE" & CStr(G_by_Lvl(Lvl))
            End If
        Next
        COLUMN_NAMEs_ordered = Mid(COLUMN_NAMEs_ordered, 2)
        CODE_COLUMNs_ordered = Mid(CODE_COLUMNs_ordered, 2)
        ORDER_BY = Mid(ORDER_BY, 2)

        Dim aNode As New Infragistics.Win.UltraWinTree.UltraTreeNode
        Dim CODE_VALUE_at_Lvl() As String = Nothing
        ReDim CODE_VALUE_at_Lvl(COLUMN_NAMEs.Count)

        Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "ABS\Menu\Tree\"

        tvwDQ.Nodes.Clear()

        Dim cur_Node_at_Lvl() As Infragistics.Win.UltraWinTree.UltraTreeNode
        ReDim cur_Node_at_Lvl(COLUMN_NAMEs.Count)
        If COLUMN_CAPTION_by_Lvl.Length = 1 Then
            aNode = tvwDQ.Nodes.Add("*", "All")
        Else
            aNode = tvwDQ.Nodes.Add("*", "All (" & COLUMN_CAPTION_by_Lvl(1) & ")")
        End If

        cur_Node_at_Lvl(0) = aNode

        ASCMAIN1.sql = "Select Distinct " & CODE_COLUMNs_ordered & ",SORT_SEQ from " & TATSTAT1
        Dim TBL As DataTable = ASCDATA1.GetDataTable
        Dim last_level_set As Integer = 0
        'If COLUMN_NAMEs.Count > 1 Then ' no nodes (other than All) when there is only 1 level

        Dim show_codes As Boolean = False
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Codes"), UltraWinToolbars.StateButtonTool)
        show_codes = tlb_sbt.Checked

        Dim images As New Dictionary(Of String, System.Drawing.Bitmap)
        images.Add("LEAF", ASCMAIN1.Get_Image(IMAGE_FOLDER, "ITEM_green"))
        images.Add("M", ASCMAIN1.Get_Image(IMAGE_FOLDER, "M"))
        images.Add("M_OPEN", ASCMAIN1.Get_Image(IMAGE_FOLDER, "M_OPEN"))

        For Each row As DataRow In TBL.Select("", ORDER_BY)
            For Lvl As Integer = 1 To COLUMN_NAMEs.Count - 1
                If CODE_VALUE_at_Lvl(Lvl) <> row.Item(Lvl - 1) & "" Or last_level_set < Lvl Then
                    last_level_set = Lvl
                    If Lvl = 1 Then
                        aNode = tvwDQ.Nodes.Add
                    Else
                        aNode = cur_Node_at_Lvl(Lvl - 1).Nodes.Add
                    End If
                    cur_Node_at_Lvl(Lvl) = aNode

                    If Lvl = COLUMN_NAMEs.Count Then
                        ' Dim KEY As String = row.Item("ITEM_CATGY_CODE") & "/" & row.Item("COLLECTION_CODE") & "/" & row.Item("ITEM_CLASS_CODE")
                        ' aNode.Key = KEY
                        ' IF WE EVER EXPAND UPON WHAT COLUMNS TO PLACE INTO THE KEY, WE NEED TO ALSO LOOK AT TXTFINDITEM_CODE
                    End If

                    Dim CAPTION As String = "?"
                    Dim COLUMN_NAME_CODE As String = COLUMN_NAME_by_Lvl(Lvl) ' Gs(Lvl - 1)

                    Dim rowTATSTATD As DataRow = dst.Tables("TATSTATD").Rows.Find(New Object() {COLUMN_NAME_CODE, row.Item(Lvl - 1) & ""})
                    If rowTATSTATD IsNot Nothing Then
                        CAPTION = rowTATSTATD.Item("DESC_VALUE") & ""
                    End If

                    'Dim rowTATSTATC As DataRow = dst.Tables("TATSTATC").Rows.Find(COLUMN_NAME_CODE)
                    'Dim COLUMN_NAME_DESC As String = rowTATSTATC.Item("COLUMN_NAME_DESC")
                    'Dim TABLE_NAME_LOOKUP As String = rowTATSTATC.Item("TABLE_NAME_LOOKUP")
                    'CAPTION = LookUp(TABLE_NAME_LOOKUP, row.Item(Lvl - 1) & "", True).Item(COLUMN_NAME_DESC) & ""

                    If CAPTION = "" Then
                        CAPTION = "?"
                    End If

                    'Select Case COLUMN_NAME_CODE
                    '    Case "PROD_CODE"
                    '        CAPTION = LookUp("ICTPROD1", row.Item(Lvl - 1), True).Item("PROD_DESC") & ""
                    '    Case Else
                    '        CAPTION = row.Item(Lvl - 1)

                    'End Select

                    'If CAPTION = "?" Then
                    '    aNode.Text = row.Item(Lvl - 1) & ""
                    'Else
                    '    If show_codes Then
                    '        aNode.Text = row.Item(Lvl - 1) & ":" & CAPTION
                    '    Else
                    '        aNode.Text = CAPTION
                    '    End If
                    'End If
                    If show_codes Then
                        aNode.Text = row.Item(Lvl - 1) & ":" & CAPTION
                    Else
                        aNode.Text = CAPTION
                    End If

                    aNode.Tag = row.Item(Lvl - 1) & ":" & CAPTION
                    aNode.Expanded = False

                    CODE_VALUE_at_Lvl(Lvl) = row.Item(Lvl - 1) & ""
                    If last_level_set = COLUMN_NAMEs.Count - 1 Then
                        aNode.LeftImages.Add(images("LEAF"))
                    Else
                        aNode.Override.NodeAppearance.Image = images("M")
                        aNode.Override.ExpandedNodeAppearance.Image = images("M_OPEN")
                    End If

                    For iLvl As Integer = 1 To Lvl
                        aNode.Cells(iLvl - 1).Value = CODE_VALUE_at_Lvl(iLvl)
                    Next
                End If
            Next
        Next
        'End If

        'If cur_Node_at_Lvl(1) IsNot Nothing Then
        '    cur_Node_at_Lvl(1).Expanded = True
        'End If

        Setup_View()
        Setup_tabDetails()

        If tvwDQ.Nodes.Count > 0 Then
            tvwDQ.ActiveNode = tvwDQ.Nodes(0)
            tvwDQ.Nodes(0).Selected = True
            Click_Node(tvwDQ.Nodes(0))
        End If

        Prepare_Charts()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Prepare_Charts()
        If tblSATSLSW1 Is Nothing Then
            Exit Sub
        End If

        'For Each ROW As DataRow In dst.Tables("SATCSLSS").Rows
        '    ROW.Item("SALES") = 0
        'Next
        'For Each row As DataRow In ASCDATA1.SelectDistinct("TATSTAT1", "CODE1").Rows
        '    Dim rowTATSTATE As DataRow = dst.Tables("TATSTATE").Rows.Find(row(0))
        '    If rowTATSTATE IsNot Nothing Then
        '        Dim rows() As DataRow = dst.Tables("SATCSLSS").Select _
        '        ("STATE_NAME = '" & rowTATSTATE("STATE_NAME") & "'")
        '        If rows.Length = 1 Then
        '            Dim SALES As Decimal = dst.Tables("SATCSLS1").Compute("SUM (P00)", "SUB_CODE_VALUE2 = '" & row(0) & "'")
        '            rows(0).Item("STATE_CODE") = rowTATSTATE("STATE_CODE")
        '            rows(0).Item("SALES") = SALES
        '        End If
        '    End If
        'Next

        '  Fill_Records("SATCSLSS")

        CreateGraph_SATCSLS1()
        CreateGraph_SATCSLS1_X()
        chtSATCSLS1.Visible = True
        chtSATCSLS1_X.Visible = True
    End Sub

    Private Sub tab1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab1.SelectedTabChanged
        Setup_tab1()
    End Sub

    Sub Setup_tab1()
        If SELECTION_NO = 0 Then Exit Sub
        If tab1.SelectedTab Is Nothing Then Exit Sub
        With UltraExplorerBar1
            .Groups("Category").Visible = (tab1.SelectedTab.Key = "Setup")
            .Groups("Division").Visible = (tab1.SelectedTab.Key = "Setup")
            '.Groups("View").Visible = (tab1.SelectedTab.Key = "Inquiry")
            .Groups("Sequence").Visible = (tab1.SelectedTab.Key = "Inquiry")
        End With
    End Sub

#Region "tvwSEQ"
    Private Sub tvwSEQ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwSEQ.Click

        'Try
        '    Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
        '    Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
        '    Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

        '    If tnode IsNot Nothing Then
        '        Click_Node(tvwSEQ.ActiveNode)
        '        tvwSEQ.SelectedNodes.Clear()
        '        tvwSEQ.ActiveNode.Selected = True
        '    End If
        'Catch ex As Exception
        'End Try
    End Sub

    Private Sub tvwSEQ_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwSEQ.DragDrop

        Dim SelectedNodes As UltraWinTree.SelectedNodesCollection
        Dim DropNode As UltraWinTree.UltraTreeNode

        DropNode = UltraTree_DropHightLight_DrawFilter.DropHightLightNode

        SelectedNodes = e.Data.GetData(GetType(UltraWinTree.SelectedNodesCollection))
        SelectedNodes = SelectedNodes.Clone()

        SelectedNodes.SortByPosition()

        Dim current_sort As String = 0
        For i As Integer = 0 To tvwSEQ.Nodes.Count - 1
            current_sort &= tvwSEQ.Nodes(i).Key
        Next

        'For i as Integer = 0 To SelectedNodes.Count - 1
        '    Node = SelectedNodes(i)
        '    Node.Reposition(DropNode, UltraWinTree.NodePosition.Previous)
        'Next
        'SelectedNodes(0).Reposition(DropNode, UltraWinTree.NodePosition.Previous)
        If DropNode.Key <> SelectedNodes(0).Key Then
            Select Case UltraTree_DropHightLight_DrawFilter.DropLinePosition
                Case DropLinePositionEnum.OnNode
                    'UltraExplorerBar1.Groups("Sequence").Text = "OnNode " & DropNode.Text
                    'For i As Integer = 0 To SelectedNodes.Count - 1
                    '    Node = SelectedNodes(i)
                    '    Node.Reposition(DropNode.Nodes)
                    'Next
                    SelectedNodes(0).Reposition(DropNode.Nodes)
                Case DropLinePositionEnum.BelowNode
                    'UltraExplorerBar1.Groups("Sequence").Text = "BelowNode " & DropNode.Text
                    'For i As Integer = 0 To SelectedNodes.Count - 1
                    '    Node = SelectedNodes(i)
                    '    Node.Reposition(DropNode, UltraWinTree.NodePosition.Next)
                    '    DropNode = Node
                    'Next
                    SelectedNodes(0).Reposition(DropNode, UltraWinTree.NodePosition.Next)
                Case DropLinePositionEnum.AboveNode ', DropLinePositionEnum.OnNode
                    'UltraExplorerBar1.Groups("Sequence").Text = "AboveNode " & DropNode.Text
                    'For i As Integer = 0 To SelectedNodes.Count - 1
                    '    Node = SelectedNodes(i)
                    '    Node.Reposition(DropNode, UltraWinTree.NodePosition.Previous)
                    'Next
                    SelectedNodes(0).Reposition(DropNode, UltraWinTree.NodePosition.Previous)
            End Select
        End If

        Dim revised_sort As String = 0
        For i As Integer = 0 To tvwSEQ.Nodes.Count - 1
            revised_sort &= tvwSEQ.Nodes(i).Key
        Next

        If current_sort <> revised_sort Then
            Absc1.Clear_grdSetup(False)
            For ii As Int16 = 0 To tvwSEQ.Nodes.Count - 1
                Absc1.Re_SEQ(tvwSEQ.Nodes(ii).Key, True)
            Next

            Generate_Inquiry(False)
        End If

        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()

        tvwSEQ.ActiveNode = SelectedNodes(0)
    End Sub

    Private Sub tvwSEQ_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwSEQ.DragLeave
        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
    End Sub

    Private Sub tvwSEQ_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwSEQ.DragOver
        Dim Node As UltraWinTree.UltraTreeNode
        Dim PointInTree As System.Drawing.Point

        With tvwSEQ
            PointInTree = .PointToClient(New System.Drawing.Point(e.X, e.Y))

            Node = .GetNodeFromPoint(PointInTree)

            If Node Is Nothing Then
                e.Effect = DragDropEffects.None
                UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
                Return
            End If

            If Me.IsParentNode(Node) And Me.IsParentNodeSelected(Me.tvwSEQ) Then
                If PointInTree.Y > (Node.Bounds.Top + 2) AndAlso PointInTree.Y < (Node.Bounds.Bottom - 2) Then
                    e.Effect = DragDropEffects.None
                    UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
                    Return
                End If
            End If

            UltraTree_DropHightLight_DrawFilter.SetDropHighlightNode(Node, PointInTree)
            e.Effect = DragDropEffects.Move
        End With
    End Sub

    Private Function IsParentNode(ByVal Node As UltraWinTree.UltraTreeNode) As Boolean
        'Dim Tag As String
        'Tag = Node.Tag
        'Return Split(Tag, Chr(1))(1) = "M"
    End Function

    Private Function IsParentNodeSelected(ByVal Tree As UltraWinTree.UltraTree) As Boolean
        'For Each SelectedNode As UltraWinTree.UltraTreeNode In Tree.SelectedNodes
        '    If Me.IsParentNode(SelectedNode) Then Return True
        'Next
        'Return False
    End Function

    Private Sub tvwSEQ_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvwSEQ.MouseUp
        tvwSEQ.SelectedNodes.Clear()
        Dim anode As Infragistics.Win.UltraWinTree.UltraTreeNode = tvwSEQ.GetNodeFromPoint(e.Location)
        If anode IsNot Nothing Then
            anode.Selected = True
            tvwSEQ.ActiveNode = anode
        End If
    End Sub

    Private Sub tvwSEQ_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles tvwSEQ.QueryContinueDrag
        If e.EscapePressed Then
            e.Action = DragAction.Cancel
            UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
        End If
    End Sub

    Private Sub tvwSEQ_SelectionDragStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwSEQ.SelectionDragStart
        tvwSEQ.DoDragDrop(tvwSEQ.SelectedNodes, DragDropEffects.Move)
    End Sub

    Private Sub UltraTree_DropHightLight_DrawFilter_Invalidate(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraTree_DropHightLight_DrawFilter.Invalidate
        tvwSEQ.Invalidate()
    End Sub

    Private Sub UltraTree_DropHightLight_DrawFilter_QueryStateAllowedForNode(ByVal sender As Object, ByVal e As UltraTree_DropHightLight_DrawFilter_Class.QueryStateAllowedForNodeEventArgs) Handles UltraTree_DropHightLight_DrawFilter.QueryStateAllowedForNode
        If Not IsParentNode(e.Node) Then
            e.StatesAllowed = DropLinePositionEnum.AboveNode Or DropLinePositionEnum.BelowNode
            UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 2
        Else
            If e.Node.Selected Then
                e.StatesAllowed = DropLinePositionEnum.AboveNode Or DropLinePositionEnum.BelowNode
                UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 2
            Else
                UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 3
            End If
        End If
    End Sub
#End Region

#Region "tvwDQ"
    Private Sub tvwDQ_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwDQ.Click

        Try
            Dim xx As System.Windows.Forms.MouseEventArgs = DirectCast(e, System.Windows.Forms.MouseEventArgs)
            Dim tt As UltraWinTree.UltraTree = DirectCast(sender, UltraWinTree.UltraTree)
            Dim tnode As UltraWinTree.UltraTreeNode = tt.GetNodeFromPoint(xx.X, xx.Y)

            If tnode IsNot Nothing Then
                Click_Node(tvwDQ.ActiveNode)
                tvwDQ.SelectedNodes.Clear()
                tvwDQ.ActiveNode.Selected = True
            End If


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
#End Region

    Sub Click_Node(ByVal tnode As UltraWinTree.UltraTreeNode)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Summary")

        If tnode IsNot Nothing Then
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show All Levels"), UltraWinToolbars.StateButtonTool)
            LVL = tnode.Level + 1
            If tnode.Key = "*" Then
                LVL = 0
            End If
            Dim COLS_Select As String = ""
            Dim COLS_Group_By As String = ""
            Dim COLS_Order_By As String = ""
            Dim sqlW As String = ""
            Dim CAPTION As String = ""

            Dim use_SORT_SEQ As Boolean = False

            If tlb_sbt.Checked Then
            Else
                For I As Int16 = 1 To 9
                    With grdTATSTAT1.DisplayLayout.Bands(0).Columns("DESC" & CStr(I))
                        .Hidden = True
                    End With
                Next
                grdTATSTAT1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Hidden = False
            End If

            For G As Integer = 1 To COLUMN_NAME_by_Lvl.Count - 1

                If G <= LVL + 1 Or tlb_sbt.Checked Then
                    COLS_Select &= ",CODE" & CStr(G_by_Lvl(G)) & " CODE" & CStr(G)
                    COLS_Group_By &= ",CODE" & CStr(G_by_Lvl(G))
                    If COLUMN_NAME_by_Lvl(G) = "SIZE_CODE" Then use_SORT_SEQ = True
                Else
                    COLS_Select &= ",NULL CODE" & CStr(G)
                End If
                If G <= LVL Then
                    If tnode.Cells(G - 1).Text = "" Then
                        sqlW &= " and CODE" & CStr(G_by_Lvl(G)) & " is Null"
                    Else
                        sqlW &= " and CODE" & CStr(G_by_Lvl(G)) & " = '" & tnode.Cells(G - 1).Text & "'"
                    End If

                    SCOPE(G) = tnode.Cells(G - 1).Text

                    Dim CODE_VALUE As String = tnode.Cells(G - 1).Text
                    Dim DESC_VALUE As String = Get_Description(COLUMN_NAME_by_Lvl(G), CODE_VALUE)
                    CAPTION &= ", " & COLUMN_CAPTION_by_Lvl(G) & " " & CODE_VALUE & ":" & DESC_VALUE
                End If

                COLS_Order_By &= ",CODE" & CStr(G)
                With grdTATSTAT1.DisplayLayout.Bands(0).Columns("CODE" & CStr(G))
                    Dim hideColumn As Boolean = (G < LVL + 1) Or (G > LVL + 1 And Not tlb_sbt.Checked)
                    .Hidden = hideColumn
                    .Header.Caption = COLUMN_CAPTION_by_Lvl(G)
                    If Not hideColumn Then
                        .Header.SetVisiblePosition(0, False)
                        grdTATSTAT1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.SetVisiblePosition(1, False)
                    End If
                End With
                If tlb_sbt.Checked Then
                    With grdTATSTAT1.DisplayLayout.Bands(0).Columns("DESC" & CStr(G))
                        .Hidden = (G < LVL + 1) Or (G > LVL + 1 And Not tlb_sbt.Checked)
                        .Header.Caption = COLUMN_CAPTION_by_Lvl(G)
                    End With
                    grdTATSTAT1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Hidden = True
                End If
            Next

            If LVL = 0 Then
                grdTATSTAT1.Text = "All" ' tvwDQ.Nodes(0).Text
            Else
                grdTATSTAT1.Text = Mid(CAPTION, 3)
            End If

            If COLUMN_NAME_by_Lvl.Count - 1 < 9 Then
                For G As Integer = COLUMN_NAME_by_Lvl.Count To 9
                    With grdTATSTAT1.DisplayLayout.Bands(0).Columns("CODE" & CStr(G))
                        .Hidden = True
                    End With
                    COLS_Select &= ",NULL CODE" & CStr(G)
                    With grdTATSTAT1.DisplayLayout.Bands(0).Columns("DESC" & CStr(G))
                        .Hidden = True
                    End With
                Next
            End If

            Dim sqlQ As String = ""
            For Each QCOL As String In QCOLS.Keys
                'what am i doing with new qty columns
                If QCOL <> "QTY_AVA" And _
                   QCOL <> "W1" And QCOL <> "W2" And QCOL <> "WKS" Then
                    sqlQ &= ",SUM (" & QCOL & ") " & QCOL
                End If
                If QCOL = "W1" Then sqlQ &= ",MIN(" & QCOL & ") " & QCOL
                If QCOL = "W2" Then sqlQ &= ",MAX(" & QCOL & ") " & QCOL
            Next
            ASCMAIN1.sql = "Select " & Mid(COLS_Select, 2) & sqlQ & IIf(use_SORT_SEQ, ",SORT_SEQ", "") & " from " _
            & TATSTAT1 & ASCMAIN1.SQL_Add_WHERE(sqlW) & " group by " & Mid(COLS_Group_By, 2) & IIf(use_SORT_SEQ, ",SORT_SEQ", "")
            Fill_Records("TATSTAT1", , , ASCMAIN1.sql)
            Dim ORDER_BY As String = Mid(COLS_Order_By, 2)
            ORDER_BY = Replace(ORDER_BY, "CODE1", "SORT_SEQ,CODE1")
            If COLUMN_NAME_by_Lvl(LVL + 1) = "SIZE_CODE" Then
                ORDER_BY = Replace(ORDER_BY, "CODE1", "SORT_SEQ,CODE1")
            End If
            Sort_grdColumns(grdTATSTAT1, ORDER_BY)

            For Each rowTATSTAT1 As DataRow In dst.Tables("TATSTAT1").Rows
                Dim CODE_VALUE As String = rowTATSTAT1.Item("CODE" & CStr(LVL + 1)) & ""
                If CODE_VALUE = "" Then
                    CODE_VALUE = "?"
                    rowTATSTAT1.Item("CODE" & CStr(LVL + 1)) = CODE_VALUE
                End If

                Dim DESC_VALUE As String = Get_Description(COLUMN_NAME_by_Lvl(LVL + 1), CODE_VALUE)
                rowTATSTAT1.Item("DESC_VALUE") = DESC_VALUE

                If tlb_sbt.Checked Then
                    For I As Int16 = 1 To COLUMN_NAME_by_Lvl.Count - 1
                        Dim CODE_VALUEX As String = rowTATSTAT1.Item("CODE" & CStr(I)) & ""
                        If CODE_VALUEX = "" Then
                            CODE_VALUEX = "?"
                            'rowTATSTAT1.Item("CODE" & CStr(I)) = CODE_VALUEX
                        End If
                        Dim DESC_VALUEX As String = Get_Description(COLUMN_NAME_by_Lvl(I), CODE_VALUEX)
                        rowTATSTAT1.Item("DESC" & CStr(I)) = DESC_VALUEX
                    Next
                End If
            Next

        End If
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Show_Codes(ByVal tf As Boolean)
        Show_Codes_for_Nodes(tf, tvwDQ.Nodes)
    End Sub

    Sub Show_Codes_for_Nodes(ByVal tf As Boolean, ByVal anodes As UltraWinTree.TreeNodesCollection)
        For Each cnode As UltraWinTree.UltraTreeNode In anodes
            If cnode.Key <> "*" Then
                Dim CAPTION As String = cnode.Tag & ""
                If tf Then
                    cnode.Text = CAPTION
                Else
                    cnode.Text = Split(CAPTION, ":")(1)
                End If

                If cnode.HasNodes Then
                    Show_Codes_for_Nodes(tf, cnode.Nodes)
                End If
            End If
        Next
    End Sub

    Private Sub grdTATSTAT1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdTATSTAT1.AfterRowActivate

        If grdTATSTAT1.ActiveRow.IsDataRow Then
            If optView.Value = "S" Then

                If Not splInquiry.Panel2Collapsed Then
                    Setup_Details()
                End If
            End If
        End If

    End Sub

    Private Sub grdTATSTAT1_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdTATSTAT1.DoubleClickCell
        Setup_Details()
        splInquiry.Panel2Collapsed = False

        Select Case e.Cell.Column.Key
            Case "CODE1", "CODE2", "CODE3", "CODE4", "CODE5", "CODE6", "CODE7", "CODE8", "CODE9"
                If tvwDQ.ActiveNode IsNot Nothing Then
                    'Dim pnode As UltraWinTree.UltraTreeNode
                    ' this code needs repair
                    ' 1) don't want same kind of code in then and else
                    ' 2) need a better way to locate the node corresponding to a row on the grid
                    If tvwDQ.ActiveNode.Key = "*" Then
                        For Each tnode As UltraWinTree.UltraTreeNode In tvwDQ.Nodes
                            If tnode.Tag Like e.Cell.Value & ":*" Then
                                Click_Node(tnode)
                            End If
                        Next
                    Else
                        For Each tnode As UltraWinTree.UltraTreeNode In tvwDQ.ActiveNode.Nodes
                            If tnode.Tag Like e.Cell.Value & ":*" Then
                                Click_Node(tnode)
                            End If
                        Next
                    End If
                End If
            Case "QTY_OPO", "QTY_OPO_ETD"
                tabDetails.SelectedTab = tabDetails.Tabs("Purchase Orders")
            Case "QTY_SL*", "QTY_SN*"
                tabDetails.SelectedTab = tabDetails.Tabs("Sales")
            Case Else
                tabDetails.SelectedTab = tabDetails.Tabs("Purchase Orders")
        End Select
    End Sub

    Function Build_SQLW() As String

        If grdTATSTAT1.ActiveRow Is Nothing Then
            Return ""
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting Detail Filter")

        Dim sqlw As String = ""
        Dim sqlw_main As String = ""
        Dim sqlw_subq As String = ""
        Dim sqlw_caption As String = ""
        Dim subQueryRequired As Boolean = False

        If Not grdTATSTAT1.DisplayLayout.GroupByBox.Hidden Then ' WE ARE SHOWING ALL LEVELS
            'NEEDS TO SUPPORT ALLOCATIONS
            For I As Int16 = 1 To COLUMN_NAME_by_Lvl.Length - 1
                Dim CODE_VALUE As String = grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(I)).Value & ""
                Dim CN As String = COLUMN_NAME_by_Lvl(I)
                If CODE_VALUE = "" Then
                    sqlw &= " and " & CN & " is Null"
                Else
                    sqlw &= " and " & CN & " = '" & CODE_VALUE & "'"
                End If
                sqlw_caption &= " and " & COLUMN_CAPTION_by_Lvl(I) & " = " & Get_Description(COLUMN_NAME_by_Lvl(I), CODE_VALUE, True)
            Next
        Else
            If LVL > 0 Then
                For I As Int16 = 1 To LVL
                    Dim colName As String = COLUMN_NAME_by_Lvl(I)
                    Dim colNameAllo As String = ""
                    Dim colValue As String = ""

                    If SCOPE(I) = "" Then
                        colValue = " is Null"
                    Else
                        colValue = " = '" & SCOPE(I) & "'"
                    End If

                    Select Case colName
                        Case "CUST_CODE"
                            subQueryRequired = True
                            colNameAllo = "DECODE(ICTLOTDA.CUST_CODE,NULL,'NC',ICTLOTDA.CUST_CODE)"
                            sqlw_subq &= " and (" & colNameAllo & colValue & " and ICTLOTDA.PROGRAM_CODE IS NULL ) "
                        Case Else
                            sqlw_main &= " and " & colName & colValue
                    End Select

                    sqlw_caption &= " and " & COLUMN_CAPTION_by_Lvl(I) & " = " & Get_Description(COLUMN_NAME_by_Lvl(I), SCOPE(I), True)
                Next
            End If
        End If

        Dim currentLevelColumn As String = COLUMN_NAME_by_Lvl(LVL + 1)
        Dim currentLevelFilter As String = ""

        If grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text = "" _
        Or grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text = "?" Then
            'need to find out if this is a SQ field or not
            currentLevelFilter = " is Null"
        Else
            currentLevelFilter = " = '" & grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text & "'"
        End If

        Dim currentLevelColumnAllo As String = ""
        Select Case currentLevelColumn
            Case "CUST_CODE"
                subQueryRequired = True
                currentLevelColumnAllo = "DECODE(ICTLOTDA.CUST_CODE,NULL,'NC',ICTLOTDA.CUST_CODE)"
                sqlw_subq &= " and (" & currentLevelColumnAllo & currentLevelFilter & " and ICTLOTDA.PROGRAM_CODE IS NULL ) "
            Case Else
                sqlw_main &= " and " & currentLevelColumn & currentLevelFilter
        End Select

        sqlw = ASCMAIN1.SQL_Add_WHERE(sqlw_main)
        sqlw_subq = ASCMAIN1.SQL_Add_WHERE(sqlw_subq)

        Return sqlw
    End Function

    Sub Setup_Details()
        If grdTATSTAT1.ActiveRow Is Nothing Then
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Details")

        Dim sqlw As String = ""
        Dim sqlw_main As String = ""
        Dim sqlw_subq As String = ""
        Dim sqlw_caption As String = ""
        Dim subQueryRequired As Boolean = False

        If Not grdTATSTAT1.DisplayLayout.GroupByBox.Hidden Then ' WE ARE SHOWING ALL LEVELS
            'NEEDS TO SUPPORT ALLOCATIONS
            For I As Int16 = 1 To COLUMN_NAME_by_Lvl.Length - 1
                Dim CODE_VALUE As String = grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(I)).Value & ""
                Dim CN As String = COLUMN_NAME_by_Lvl(I)
                If CODE_VALUE = "" Then
                    sqlw &= " and " & CN & " is Null"
                Else
                    sqlw &= " and " & CN & " = '" & CODE_VALUE & "'"
                End If
                sqlw_caption &= " and " & COLUMN_CAPTION_by_Lvl(I) & " = " & Get_Description(COLUMN_NAME_by_Lvl(I), CODE_VALUE, True)
            Next
        Else
            If LVL > 0 Then
                For I As Int16 = 1 To LVL
                    Dim colName As String = COLUMN_NAME_by_Lvl(I)
                    Dim colNameAllo As String = ""
                    Dim colValue As String = ""

                    If SCOPE(I) = "" Then
                        colValue = " is Null"
                    Else
                        colValue = " = '" & SCOPE(I) & "'"
                    End If

                    Select Case colName
                        Case "CUST_CODE"
                            subQueryRequired = True
                            colNameAllo = "DECODE(ICTLOTDA.CUST_CODE,NULL,'NC',ICTLOTDA.CUST_CODE)"
                            sqlw_subq &= " and (" & colNameAllo & colValue & " and ICTLOTDA.PROGRAM_CODE IS NULL ) "
                        Case Else
                            sqlw_main &= " and " & colName & colValue
                    End Select

                    sqlw_caption &= " and " & COLUMN_CAPTION_by_Lvl(I) & " = " & Get_Description(COLUMN_NAME_by_Lvl(I), SCOPE(I), True)
                Next
            End If
        End If

        Dim currentLevelColumn As String = COLUMN_NAME_by_Lvl(LVL + 1)
        Dim currentLevelFilter As String = ""

        If grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text = "" _
        Or grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text = "?" Then
            'need to find out if this is a SQ field or not
            currentLevelFilter = " is Null"
        Else
            currentLevelFilter = " = '" & grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text & "'"
        End If

        Dim currentLevelColumnAllo As String = ""
        Select Case currentLevelColumn
            Case "CUST_CODE"
                subQueryRequired = True
                currentLevelColumnAllo = "DECODE(ICTLOTDA.CUST_CODE,NULL,'NC',ICTLOTDA.CUST_CODE)"
                sqlw_subq &= " and (" & currentLevelColumnAllo & currentLevelFilter & " and ICTLOTDA.PROGRAM_CODE IS NULL ) "
            Case Else
                sqlw_main &= " and " & currentLevelColumn & currentLevelFilter
        End Select

        sqlw = ASCMAIN1.SQL_Add_WHERE(sqlw_main)
        sqlw_subq = ASCMAIN1.SQL_Add_WHERE(sqlw_subq)

        ' Purchase Orders

        If MENU_ITEM_OBJECT = "SAFSLSA1" Then
            ASCMAIN1.sql = "Select " & SOTINVHX & ".* from " & SOTINVHX & sqlw ' Replace(sqlw, "ICTLOTDX", POTORDRX)
            Fill_Records("SOTINVHX", , , ASCMAIN1.sql)
            Load_Descriptions("SOTINVHX")
            grdPOTORDRX.Text = "Sales Invoices" & sqlw_caption
        ElseIf MENU_ITEM_OBJECT = "SAFCOMPA" Then

        End If



        ' Styles

        'ASCMAIN1.sql = "Select " & GMTCGMAX & ".* from " & GMTCGMAX & sqlw ' Replace(sqlw, "ICTLOTDX", POTORDRX)
        'Fill_Records("GMTCGMAX", , , ASCMAIN1.sql)
        'Load_Descriptions("GMTCGMAX")
        'grdGMTCGMAX.Text = "Styles" & sqlw_caption


        EnforceConstraints(True)

        sqlw_caption &= " and " & COLUMN_CAPTION_by_Lvl(LVL + 1) & " = " & Get_Description(COLUMN_NAME_by_Lvl(LVL + 1), grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text, True)
        sqlw_caption = ASCMAIN1.SQL_Add_WHERE(sqlw_caption)

        Dim COLUMNS_out_of_forecast_scope As String = ""
        sqlw = ""
        If LVL > 0 Then
            For I As Int16 = 1 To LVL
                sqlw &= " and " & COLUMN_NAME_by_Lvl(I) & " = '" & SCOPE(I) & "'"
            Next
        End If
        sqlw &= " and " & COLUMN_NAME_by_Lvl(LVL + 1) & " = '" & grdTATSTAT1.ActiveRow.Cells("CODE" & CStr(LVL + 1)).Text & "'"
        sqlw = ASCMAIN1.SQL_Add_WHERE(sqlw)

        ' Hide columns that are part of scope

        If MENU_ITEM_OBJECT = "SAFSLSA1" Then

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdPOTORDRX} ' , grdGMTCGMAX}
                For Each rowASTDSQLA As DataRow In tblASTDSQLA.Rows
                    Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")
                    If COLUMN_NAME <> "CUST_CODE" Then
                        grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False

                        If New String() {"PROD_CODE"}.Contains(COLUMN_NAME) Then
                            COLUMN_NAME = Replace(COLUMN_NAME, "_CODE", "_DESC")
                            grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
                        End If
                    End If
                Next

                For I As Int16 = 1 To LVL + 1
                    Dim COLUMN_NAME As String = COLUMN_NAME_by_Lvl(I)
                    If COLUMN_NAME <> "CUST_CODE" Then
                        grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True

                        If New String() {"PROD_CODE"}.Contains(COLUMN_NAME) Then
                            COLUMN_NAME = Replace(COLUMN_NAME, "_CODE", "_DESC")
                            grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
                        End If
                    End If
                Next
            Next

        ElseIf MENU_ITEM_OBJECT = "SAFCOMPA" Then

        End If

        Setup_tabDetails()

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdPOTORDRX} ' , grdGMTCGMAX}
            'Dim visiblePositionMOP As Integer = grd.DisplayLayout.Bands(0).Columns("MOP").Header.VisiblePosition
            'For Each COLUMN_NAME As String In New String() {"DIVISION_CODE_O"}
            '    If (grd.DisplayLayout.Bands(0).Columns.Exists(COLUMN_NAME)) Then
            '        With grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
            '            .Hidden = False
            '            .Header.Appearance.TextHAlign = HAlign.Center
            '            .CellAppearance.TextHAlign = HAlign.Center
            '            .CellAppearance.BackColor = Drawing.Color.Yellow
            '            .Header.VisiblePosition = visiblePositionMOP
            '        End With
            '    End If
            'Next
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdPOTORDRX} ' , grdGMTCGMAX}
            For Each COLUMN_NAME As String In New String() {"PROD_DIV_CODE"}
                If (grd.DisplayLayout.Bands(0).Columns.Exists(COLUMN_NAME)) Then
                    With grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                        .Hidden = True
                    End With
                End If
            Next
        Next

        ASCMAIN1.Progress("")
    End Sub

    Private Sub optView_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optView.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_View()
    End Sub

    Sub Setup_View()
        If SELECTION_NO = 0 Then Exit Sub

        grdTATSTAT1.Visible = (optView.Value = "S")

        With UltraExplorerBar1
            .Groups("Sequence").Visible = tab1.SelectedTab IsNot Nothing AndAlso (tab1.SelectedTab.Key = "Inquiry" And (optView.Value = "S"))
        End With


        splInquiry.Panel2Collapsed = True

        'UltraExplorerBar1.Groups("Sales History Weeks").Visible = (optView.Value = "S")
        UltraExplorerBar1.Groups("Charts").Visible = False
    End Sub

    Function where_clause(ByVal DATA_SOURCE As String) As String
        Dim where As String = ""
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Select Case DATA_SOURCE
                Case "ICTLOTD2"
                Case "POTORDR2"

            End Select
            Absc1.SQLA_filter(COLUMN_NAME, "TAB", "COL")
        Next
        Return where
    End Function

    Sub Format_grdTATSTAT1()

        With grdTATSTAT1.DisplayLayout.Bands(0)

            If MENU_ITEM_OBJECT = "SAFSLSA1" Then
                .Columns("TY_YTD_GRS").CellAppearance.BackColor = Drawing.Color.LightBlue
                .Columns("TY_YTD_RTN").CellAppearance.BackColor = Drawing.Color.LightBlue
                .Columns("TY_YTD_NET").CellAppearance.BackColor = Drawing.Color.LightBlue

                For I As Int16 = 1 To 9
                    .Columns("CODE" & CStr(I)).Header.Appearance.BackColor2 = Drawing.Color.Yellow
                    .Columns("CODE" & CStr(I)).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                    .Columns("CODE" & CStr(I)).CellAppearance.BackColor = Drawing.Color.LightYellow
                Next
                For I As Int16 = 1 To 9
                    .Columns("DESC" & CStr(I)).Header.Appearance.BackColor2 = Drawing.Color.Yellow
                    .Columns("DESC" & CStr(I)).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                    .Columns("DESC" & CStr(I)).CellAppearance.BackColor = Drawing.Color.LightYellow
                Next
                .Columns("DESC_VALUE").Header.Appearance.BackColor2 = Drawing.Color.Yellow
                .Columns("DESC_VALUE").Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                .Columns("DESC_VALUE").CellAppearance.BackColor = Drawing.Color.LightYellow
            ElseIf MENU_ITEM_OBJECT = "SAFCOMPA" Then

                .Columns("RANK_LY").CellAppearance.BackColor = Drawing.Color.LightBlue
                .Columns("RANK_2Y").CellAppearance.BackColor = Drawing.Color.LightBlue

                .Columns("PCT_LY2Y").CellAppearance.BackColor = Drawing.Color.HotPink
                .Columns("PCT_TYLY").CellAppearance.BackColor = Drawing.Color.HotPink

                .Columns("GAP").CellAppearance.BackColor = Drawing.Color.LightGreen

                .Columns("PCT_TOT").CellAppearance.BackColor = Drawing.Color.Gold
                .Columns("PCT_CUM").CellAppearance.BackColor = Drawing.Color.LightGray

            End If


            'For I As Int16 = 0 To 6
            '    With .Columns("QTY_SL" & CStr(I))
            '        .Header.Appearance.BackColor2 = Drawing.Color.Violet ' DarkSeaGreen
            '        .Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '        .CellAppearance.BackColor = Drawing.Color.LightYellow
            '        .Hidden = True
            '    End With
            '    .Columns("QTY_SL0").CellAppearance.BackColor = Color.Aqua
            '    With .Columns("QTY_SN" & CStr(I))
            '        .Header.Appearance.BackColor2 = Drawing.Color.PowderBlue ' Orchid
            '        .Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '        .CellAppearance.BackColor = Drawing.Color.LightYellow
            '        .Hidden = True
            '    End With
            'Next


            'For Each COLUMN_NAME As String In New String() {"QTY_LASTX", "ROS_LASTX", "WOS_LASTX"}
            '    With .Columns(COLUMN_NAME)
            '        .Header.Appearance.BackColor2 = Drawing.Color.Violet
            '        .Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20
            '    End With
            'Next

        End With

        For Each QCOL As String In QCOLS.Keys
            With grdTATSTAT1.DisplayLayout.Bands(0).Columns(QCOL)
                .Format = "###,##0"
                .Header.Caption = QCOLS(QCOL)
                .Width = 65
                'If New String() {"ONH_UNITS_X_COST", "VAL_AVA", "VAL_OPO", "VAL_AFL", "VAL_XIT"}.Contains(QCOL) Then
                '    .Hidden = True
                '    .Width = 100
                '    .Format = "#,##0.00"
                'End If
            End With

            If QCOL = "AVG_COST" Or QCOL = "AVG_COST_ADJ" Then
                Create_Summary(grdTATSTAT1, QCOL, "Custom")
            Else
                Create_Summary(grdTATSTAT1, QCOL)
            End If
        Next

        With grdTATSTAT1.DisplayLayout.Bands("TATSTAT1")
            .Columns("CODE1").Header.Fixed = True
            .Columns("CODE2").Header.Fixed = True
            .Columns("CODE3").Header.Fixed = True
            .Columns("CODE4").Header.Fixed = True
            .Columns("CODE5").Header.Fixed = True
            .Columns("CODE6").Header.Fixed = True
            .Columns("CODE7").Header.Fixed = True
            .Columns("CODE8").Header.Fixed = True
            .Columns("CODE9").Header.Fixed = True
            .Columns("DESC_VALUE").Header.Fixed = True
            ' .Columns("QTY_AVA").Header.Fixed = True
        End With

    End Sub

    Sub Get_SQL_for_COLUMN_NAME(ByVal sender As Object, ByVal e As ABSC.ABSC.grdSetupClickCellEventArgs) _
   Handles Absc1.Get_SQL_for_COLUMN_NAME

        Dim PROD_CODEs As String = ""
        'If chkUseSmartLists.Checked Then
        '    PROD_CODEs = Absc1.SQLA("PROD_CODE", , True)
        'End If

        Select Case e.COLUMN_NAME
            ' NOTE COOL IS NOT IN SMART LISTS
            ' NOTE CATGY IS NOT IN SMART LISTS
            Case "COOL_COMPLIANT", "CATEGORY_CODE"
            Case "PRODUCT_TYPE"

                Dim ptSql As String = "Select T_CODE TYPE, T_DESC Description FROM" _
                & " ASTCODE1 WHERE TABLE_NAME = 'ASTCODE1' AND COLUMN_NAME = 'PRODUCT_TYPE'"
                e.SQL = ptSql
            Case Else

                Dim COLUMN_NAME_in_ICTLOTD2 As String = e.COLUMN_NAME
                Dim COLUMN_NAME_in_LOOKUP As String = e.COLUMN_NAME
                If e.COLUMN_NAME = "DIVISION_CODE_O" Then
                    COLUMN_NAME_in_ICTLOTD2 = "PROD_DIV_CODE"
                    COLUMN_NAME_in_LOOKUP = "DIVISION_CODE"
                End If
                If e.COLUMN_NAME = "DIVISION_CODE_G" Then
                    COLUMN_NAME_in_ICTLOTD2 = "DIVISION_CODE"
                    COLUMN_NAME_in_LOOKUP = "DIVISION_CODE"
                End If
                'If chkUseSmartLists.Checked And PROD_CODEs <> "" Then
                '    e.SQL = "Select * from (" & e.SQL & ") where " & COLUMN_NAME_in_LOOKUP & " in " _
                '    & " (Select Distinct " & COLUMN_NAME_in_ICTLOTD2 & " from ICTLOTD2 where PROD_CODE in (" & PROD_CODEs & "))"
                'End If
        End Select
    End Sub

    Private Sub grdTATFAVE1_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdTATFAVE1.DoubleClickRow
        Load_Favorite(e.Row.Cells("SET_ID").Value & "")
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
        UltraExplorerBar1.Groups("Charts").Visible = tabDetails.SelectedTab IsNot Nothing AndAlso (tabDetails.SelectedTab.Key = "Sales" Or tabDetails.SelectedTab.Key = "Map")
        UltraExplorerBar1.Groups("Sales History").Visible = tabDetails.SelectedTab IsNot Nothing AndAlso (tabDetails.SelectedTab.Key = "Sales" Or tabDetails.SelectedTab.Key = "Sales History")

        'If tabDetails.SelectedTab.Key = "Purchase Orders" Then
        '    grdPOTORDRX.Parent = tabDetails.Tabs("Purchase Orders").TabPage
        '    Setup_PO("O")
        'End If
    End Sub

    Sub Load_Favorite(ByVal SET_ID As String)

        Dim rowTATFAVE1 As DataRow = dst.Tables("TATFAVE1").Rows.Find(SET_ID)

        If Not rowTATFAVE1 Is Nothing Then
            Clear_Settings()
            Absx1.txtFor("SET_DESC").Text = rowTATFAVE1.Item("SET_DESC") & ""
            Absx1.chkFor("SET_PUBLIC").Checked = rowTATFAVE1.Item("SET_PUBLIC") & "" = "1"
            Absx1.txtFor("PROGRAM_CODE").Text = rowTATFAVE1.Item("PROGRAM_CODE") & ""

            grpFavorite.Text = "Favorite " & rowTATFAVE1.Item("SET_ID") & ""
            grpFavorite.Tag = SET_ID

            Dim t As DataTable = Absc1.grdSetupDataSource
            For Each rd As String In Split(rowTATFAVE1.Item("DATASOURCE"), vbCrLf)
                Dim f() As String = Split(rd, vbTab)
                If f.Length > 1 Then
                    Dim r As DataRow = t.Rows.Find(f(1))
                    For i As Int16 = 0 To t.Columns.Count - 1
                        If f(i + 1) = "" Then
                            r.Item(i) = DBNull.Value
                        Else
                            r.Item(i) = f(i + 1)
                        End If
                    Next
                End If
            Next
        End If
    End Sub

    Sub Clear_Settings()
        txtSET_DESC.Text = ""
        grpFavorite.Text = "Favorite"
        grpFavorite.Tag = ""
        Absc1.Clear_grdSetup(True)
        tab1.Tabs("Inquiry").Text = "Inquiry"
        tab1.Tabs("Inquiry").Visible = False
        tab1.SelectedTab = tab1.Tabs("Setup")

        Me.Text = Me.MENU_ITEM_DESC
        ASCMAIN1.Set_Form_Caption_on_Tab(Me)
    End Sub

    Private Sub grdTATSTAT1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdTATSTAT1.InitializeRow

        'For Each CU As String In New String() {"", "_UNITS"}
        '    For Each COLUMN_NAME As String In New String() _
        '        {"QTY_ONH", "QTY_OPO", "QTY_OPO_ETD", "QTY_AFL", "QTY_XIT", "QTY_FDA"}
        '        Dim COLUMN_NAME_COM As String = COLUMN_NAME & "_COM" & CU
        '        COLUMN_NAME = COLUMN_NAME & CU
        '        If COLUMN_NAME = "QTY_ONH" & CU Then COLUMN_NAME_COM = "QTY_COM" & CU

        '        If COLUMN_NAME = "QTY_ONH" & CU Then
        '            ' DO NOTHING, SINCE WE DO NOT SHOW QTY_ONH - QTY_COM IN THE QTY_ONH COLUMN, AS WE DO FOR THE OTHER SUPPLY-SIDE COLUMNS
        '        Else
        '             e.Row.Cells(COLUMN_NAME).Appearance.ForeColor = Color.Empty
        '        End If
        '    Next
        'Next

        'If Val(e.Row.Cells("QTY_OVR").Value & "") <> 0 Then
        '    e.Row.Cells("QTY_OVR").Appearance.BackColor = Color.Orange
        '    'e.Row.Cells("QTY_OVR").Appearance.ForeColor = Color.White
        'Else
        '    e.Row.Cells("QTY_OVR").Appearance.BackColor = Color.Empty
        '    'e.Row.Cells("QTY_OVR").Appearance.ForeColor = Color.Empty
        'End If
    End Sub

    Function Get_Description( _
    ByVal COLUMN_NAME As String, _
    ByVal CODE_VALUE As String, _
    Optional ByVal use_code_as_default_value As Boolean = False)
        Dim DESC_VALUE As String = IIf(use_code_as_default_value, CODE_VALUE, "")
        Dim rowTATSTATD As DataRow = dst.Tables("TATSTATD").Rows.Find _
                       (New String() {COLUMN_NAME, CODE_VALUE})
        If rowTATSTATD IsNot Nothing Then
            DESC_VALUE = rowTATSTATD.Item("DESC_VALUE") & ""
        Else
            If COLUMN_NAME = "CUST_CODE" And CODE_VALUE = "NC" Then
                DESC_VALUE = "Non-Customer"
            End If
        End If

        Return DESC_VALUE
    End Function

    Sub Setup_PO(ByVal OA As String)
        If OA = "NOT USED" Then
            Dim dvw As DataView = DirectCast(grdPOTORDRX.DataSource, DataTable).DefaultView
            If OA = "O" Then
                dvw.RowFilter = "ISNULL(IMPORT_NO,'?') = '?'"
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("IMPORT_NO").Hidden = True
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_ETD").Hidden = False
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY").Hidden = False
                Sort_grdColumns(grdPOTORDRX, "PO_ORDER_NO")
                grdPOTORDRX.Text = Replace(grdPOTORDRX.Text, "Purchase Orders Afloat", "Open Purchase Orders")
            Else
                dvw.RowFilter = "ISNULL(IMPORT_NO,'?') <> '?'"
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("IMPORT_NO").Hidden = False
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_ETD").Hidden = True
                grdPOTORDRX.DisplayLayout.Bands(0).Columns("PO_DATE_SHIP_BY").Hidden = True
                Sort_grdColumns(grdPOTORDRX, "IMPORT_NO")
                grdPOTORDRX.Text = Replace(grdPOTORDRX.Text, "Open Purchase Orders", "Purchase Orders Afloat")
            End If
        End If
    End Sub

    Public Overrides Function CustomSummary_End( _
    ByVal summarySettings As UltraWinGrid.SummarySettings, _
    ByVal rows As UltraWinGrid.RowsCollection, _
    ByVal CustomValue As Double, _
    ByVal grd As UltraWinGrid.UltraGrid) As Double
        Select Case grd.Name
            Case "grdICTLOTDX", "grdTATSTAT1"
                Dim FIELD_NAME As String = ""
                Select Case summarySettings.Key
                    Case "AVG_COST"
                        FIELD_NAME = "ONH_UNITS_X_COST"
                End Select

                'Dim OPT As String = "1"

                Dim QTY_ON_HAND_UNITS As Decimal = 0

                Dim EXT_X_COST As Decimal = 0
                Select Case grd.Name
                    Case "grdTATSTAT1"
                        'QTY_ON_HAND_UNITS = Val(dst.Tables("TATSTAT1").Compute("SUM (QTY_ONH_UNITS)", "") & "")
                        'EXT_X_COST = Val(dst.Tables("TATSTAT1").Compute("SUM (" & FIELD_NAME & ")", "") & "")
                    Case "grdICTLOTDX"
                        'If OPT = "1" Then
                        'Else
                        'End If
                        QTY_ON_HAND_UNITS = Val(dst.Tables("ICTLOTDX").Compute("SUM (QTY_ON_HAND_UNITS)", "ISNULL(ON_HOLD_FLAG,'M')<>'R'") & "")
                        EXT_X_COST = Val(dst.Tables("ICTLOTDX").Compute("SUM (" & FIELD_NAME & ")", "ISNULL(ON_HOLD_FLAG,'M')<>'R'") & "")
                End Select

                If QTY_ON_HAND_UNITS = 0 Then
                    CustomValue = 0
                Else
                    CustomValue = EXT_X_COST / QTY_ON_HAND_UNITS
                End If

            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub Add_Description_Columns(ByVal TABLE_NAME As String)
        With dst.Tables(TABLE_NAME).Columns
            For Each COLUMN_NAME In New String() {"PROD_DESC", "SIZE_DESC", "PACK_DESC", "WHSE_DESC", "WHSE_DESC_PROJ"}
                If Not .Contains(COLUMN_NAME) Then
                    .Add(COLUMN_NAME)
                End If
            Next
        End With
    End Sub

    Sub Load_Descriptions(ByVal TABLE_NAME As String)
        Dim CODE_DESC As New Dictionary(Of String, String)
        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            'For Each COLUMN_NAME As String In New String() _
            '{"PROD_CODE", "SIZE_CODE", "PACK_CODE", "WHSE_CODE", "WHSE_CODE_PROJ"}
            '    Dim COLUMN_NAME_TO_USE As String = COLUMN_NAME
            '    Dim C2 As String = COLUMN_NAME
            '    If TABLE_NAME = "POTORDRX" And COLUMN_NAME = "WHSE_CODE_PROJ" Then
            '        COLUMN_NAME_TO_USE = "WHSE_CODE_PROJ"
            '        C2 = "WHSE_CODE"
            '    End If
            '    If COLUMN_NAME = "WHSE_CODE_PROJ" And TABLE_NAME <> "POTORDRX" Then
            '        Exit For
            '    End If
            '    Dim COLUMN_NAME_KEY = C2 & ":" & row.Item(COLUMN_NAME_TO_USE)
            '    If Not CODE_DESC.ContainsKey(COLUMN_NAME_KEY) Then
            '        CODE_DESC.Add(COLUMN_NAME_KEY, Get_Description(C2, row.Item(COLUMN_NAME_TO_USE) & "", True))
            '    End If
            '    Dim COLUMN_NAME_DESC = Replace(COLUMN_NAME, "_CODE", "_DESC")
            '    row.Item(COLUMN_NAME_DESC) = CODE_DESC(COLUMN_NAME_KEY)
            'Next
        Next
        dst.Tables(TABLE_NAME).AcceptChanges()
    End Sub

    Function Prepare_for_PO(ByVal sql As String) As String
        sql = Replace(sql, "ICTLOTD2 ICTLOTD2", "ICTWHSE1,POTORDR1,POTORDR2 ICTLOTD2")
        sql = Replace(sql, "ICTLOTD2.WHSE_CODE (+) = SOTORDR3.WHSE_CODE", "POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO")
        sql = Replace(sql, "ICTLOTD2.LOT_NO (+) = SOTORDR3.LOT_NO", "ICTWHSE1.WHSE_CODE (+) = POTORDR1.WHSE_CODE")
        sql = Replace(sql, "ICTLOTD2.LOT_SEQ_NO (+) = SOTORDR3.LOT_SEQ_NO", "POTORDR2.PO_ORDER_NO = SOTORDR3.PO_ORDER_NO and POTORDR2.PO_ORDER_LNO = SOTORDR3.PO_ORDER_LNO and SOTORDR3.PRE_SALES_TYPE = 'P'")
        sql = Replace(sql, "ICTLOTD2.ORIG_CODE", "POTORDR1.ORIG_CODE")
        sql = Replace(sql, "ICTLOTD2.IMPORT_NO", "POTORDR1.IMPORT_NO")
        sql = Replace(sql, "ICTLOTD2.TERR_CODE", "ICTWHSE1.TERR_CODE")
        sql = Replace(sql, "ICTLOTD2.WHSE_CODE", "POTORDR1.WHSE_CODE")
        sql = Replace(sql, "ICTLOTD2.CLASS_CODE", "ICTPROD1.CLASS_CODE")
        sql = Replace(sql, "ICTLOTD2", "POTORDR2")
        sql = Replace(sql, "SOTORDR3.LOT_NO LOT_NO_SHOW", "DECODE(POTORDR1.IMPORT_NO,NULL,'POrder','Afloat') LOT_NO_SHOW")
        Return sql
    End Function

    Sub CreateGraph_SATCSLS1()

        Dim chtIsVisible As Boolean = chtSATCSLS1.Visible
        chtSATCSLS1.Visible = False

        chtSATCSLS1.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String

        chtSATCSLS1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1.LabelHash = labelHash

        chtSATCSLS1.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1.Tooltips.FormatString = "<HIGHLOW>"

        Dim DTY As New DataTable
        With DTY
            .Columns.Add("CODE")
            .Columns.Add("VALUE", GetType(System.Int64))
        End With

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("TATSTAT1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("TATSTAT1").Select("", "CODE1")
            RL(RLi) = row("CODE1") ' & ":" & row("DESC_VALUE")
            RL(RLi) = row("DESC_VALUE") & ""
            RLi += 1
            'DTY.Rows.Add(New Object() {row.Item("CODE1"), row.Item(IIf(optXY.Value = "X", "QTY_LASTX", "QTY_NEXTY"))})
            'this needs to know about LASTX_LY
            If MENU_ITEM_OBJECT = "SAFCOMPA" Then
                Select Case optXY.Value
                    Case "X"
                        DTY.Rows.Add(New Object() {row.Item("DESC_VALUE"), row.Item("FULL_2Y")})
                    Case "Y"
                        DTY.Rows.Add(New Object() {row.Item("DESC_VALUE"), row.Item("FULL_TY")})
                    Case "Z"
                        DTY.Rows.Add(New Object() {row.Item("DESC_VALUE"), row.Item("YTD_TY")})
                End Select
            Else
                Select Case optXY.Value
                    Case "X"
                        DTY.Rows.Add(New Object() {row.Item("DESC_VALUE"), row.Item("TY_YTD_GRS")})
                    Case "Y"
                        DTY.Rows.Add(New Object() {row.Item("DESC_VALUE"), row.Item("TY_YTD_RTN")})
                    Case "Z"
                        DTY.Rows.Add(New Object() {row.Item("DESC_VALUE"), row.Item("TY_YTD_NET")})
                End Select
            End If


        Next

        'SLS_LWK_UNITS NUMBER(7)
        'SLS_2WK_UNITS NUMBER(7)
        'SLS_3WK_UNITS NUMBER(7)
        'SLS_4WK_UNITS NUMBER(7)
        'SLS_5WK_UNITS NUMBER(7)

        'chtSATCSLS1.Data.SetRowLabels(RL)
        'chtSATCSLS1.Data.SetColumnLabels(CL)

        'chtSATCSLS1.DataSource = dst.Tables("SATCSLS1")
        chtSATCSLS1.DataSource = DTY
        chtSATCSLS1.PieChart.ColumnIndex = -1
        chtSATCSLS1.PieChart.OthersCategoryPercent = 2
        'chtSATCSLS1.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1.Data.IncludeColumn("P00", True)


        chtSATCSLS1.DataBind()

        chtSATCSLS1.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Sub CreateGraph_SATCSLS1_X()

        Dim weeks As Int16 = 0
        'Select Case optXY.Value
        '    Case "X"
        '        weeks = numLASTX.Value
        '    Case "Y"
        '        weeks = numNEXTY.Value
        '    Case "Z"
        '        weeks = numLASTX_LY.Value
        'End Select
        weeks = 12
        If MENU_ITEM_OBJECT = "SAFCOMPA" Then
            weeks = 3
        End If

        Dim chtIsVisible As Boolean = chtSATCSLS1_X.Visible
        chtSATCSLS1_X.Visible = False

        chtSATCSLS1_X.DataSource = Nothing

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Charting Data")

        Me.SuspendLayout()

        Dim RL() As String
        Dim CL() As String
        ReDim CL(weeks)

        'this will be necessary for line graph
        'For i As Integer = MOSMAX To 0 Step -1
        '    Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
        '    CL(MOSMAX - i) = Mid(L, 10, 6)
        '    grdSATCSLS1.DisplayLayout.Bands(0).Columns("M" & Format(i, "00")).Header.Caption = Mid(L, 10, 3)
        'Next
        If MENU_ITEM_OBJECT = "SAFSLSA1" Then
            For i As Integer = 1 To weeks
                'Dim L As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * i))
                'If optXY.Value = "X" Then
                '    CL(i - 1) = Format(YWPD(weeks - (i - 1)), "MM/dd")
                'Else
                '    CL(i - 1) = Format(YWPD(52 - (i - 1)), "MM/dd")
                'End If
                CL(i - 1) = grdTATSTAT1.DisplayLayout.Bands(0).Columns("TY_M" & Format(i, "00") & "_GRS").Header.Caption
            Next
        ElseIf MENU_ITEM_OBJECT = "SAFCOMPA" Then
            CL(1) = grdTATSTAT1.DisplayLayout.Bands(0).Columns("FULL_2Y").Header.Caption
            CL(2) = grdTATSTAT1.DisplayLayout.Bands(0).Columns("FULL_LY").Header.Caption
            CL(3) = grdTATSTAT1.DisplayLayout.Bands(0).Columns("YTD_TY").Header.Caption
        End If


        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.LabelPlusDataValue
        'chtICTINVAT.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom

        chtSATCSLS1_X.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        chtSATCSLS1_X.LabelHash = labelHash

        chtSATCSLS1_X.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        chtSATCSLS1_X.Tooltips.FormatString = "<HIGHLOW>"

        Dim DT As New DataTable
        DT.Columns.Add("CODE_VALUE")
        DT.Columns.Add("DESC_VALUE")
        For P As Integer = 1 To weeks
            DT.Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
        Next

        'SLS_LWK_UNITS NUMBER(7)
        'SLS_2WK_UNITS NUMBER(7)
        'SLS_3WK_UNITS NUMBER(7)
        'SLS_4WK_UNITS NUMBER(7)
        'SLS_5WK_UNITS NUMBER(7)

        Dim RLi As Integer = 0
        ReDim RL(dst.Tables("TATSTAT1").Rows.Count - 1)
        For Each row As DataRow In dst.Tables("TATSTAT1").Select("", "CODE1")
            RL(RLi) = row("CODE1") ' & ":" & row("DESC_VALUE")
            RL(RLi) = row("DESC_VALUE") & ""
            RLi += 1

            Dim rowDT As DataRow = DT.NewRow
            rowDT.Item("CODE_VALUE") = row("CODE1")
            rowDT.Item("DESC_VALUE") = row("DESC_VALUE")
            For P As Integer = 1 To weeks
                Dim rowSATSLSW1 As DataRow = tblSATSLSW1.Rows.Find(row("CODE1") & "")
                Dim VALUE As Decimal = 0
                If rowSATSLSW1 IsNot Nothing Then
                    If MENU_ITEM_OBJECT = "SAFSLSA1" Then

                        Select Case optXY.Value
                            'Case "X"
                            '    VALUE = Val(rowSATSLSW1.Item("LAST_" & Format(P, "00")) & "")
                            'Case "Y"
                            '    VALUE = Val(rowSATSLSW1.Item("NEXT_" & Format(P, "00")) & "")
                            'Case "Z"
                            '    VALUE = Val(rowSATSLSW1.Item("LAST_" & Format(P, "00")) & "_LY")
                            'Case "X"
                            '    VALUE = Val(rowSATSLSW1.Item("SLS_LWK_UNITS") & "")
                            'Case "Y"
                            '    VALUE = Val(rowSATSLSW1.Item("SLS_2WK_UNITS") & "")
                            'Case "Z"
                            '    VALUE = Val(rowSATSLSW1.Item("SLS_3WK_UNITS") & "")
                        End Select
                        VALUE = Val(rowSATSLSW1.Item("TY_M" & Format(P, "00") & "_GRS") & "")

                    ElseIf MENU_ITEM_OBJECT = "SAFCOMPA" Then

                        Select Case optXY.Value
                            'Case "X"
                            '    VALUE = Val(rowSATSLSW1.Item("LAST_" & Format(P, "00")) & "")
                            'Case "Y"
                            '    VALUE = Val(rowSATSLSW1.Item("NEXT_" & Format(P, "00")) & "")
                            'Case "Z"
                            '    VALUE = Val(rowSATSLSW1.Item("LAST_" & Format(P, "00")) & "_LY")
                            'Case "X"
                            '    VALUE = Val(rowSATSLSW1.Item("SLS_LWK_UNITS") & "")
                            'Case "Y"
                            '    VALUE = Val(rowSATSLSW1.Item("SLS_2WK_UNITS") & "")
                            'Case "Z"
                            '    VALUE = Val(rowSATSLSW1.Item("SLS_3WK_UNITS") & "")
                        End Select
                        VALUE = Val(rowSATSLSW1.Item(P) & "")

                    End If

                End If

                rowDT.Item("P" & Format(P, "00")) = VALUE
            Next
            DT.Rows.Add(rowDT)
        Next
        chtSATCSLS1_X.Data.SetRowLabels(RL)
        chtSATCSLS1_X.Data.SetColumnLabels(CL)

        chtSATCSLS1_X.DataSource = DT
        'chtSATCSLS1_X.Data.IncludeColumn("CODE_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("DESC_VALUE", False)
        'chtSATCSLS1_X.Data.IncludeColumn("P00", False)

        chtSATCSLS1_X.DataBind()

        chtSATCSLS1_X.Visible = chtIsVisible

        Me.ResumeLayout()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
        Application.DoEvents()
    End Sub

    Private Sub trkbrXAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrXAxis.Scroll
        chtSATCSLS1_X.Axis.X.ScrollScale.Scale = Me.trkbrXAxis.Value / 100.0
    End Sub

    Private Sub trkbrYAxis_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles trkbrYAxis.Scroll
        chtSATCSLS1_X.Axis.Y.ScrollScale.Scale = Me.trkbrYAxis.Value / 100.0
    End Sub

    Sub Set_Totals_ChartType()
        If Not chkTotalsChart3D.Checked Then
            chtSATCSLS1_X.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.LineChart
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATCSLS1.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.PieChart
                Case "DoughnutChart"
                    chtSATCSLS1.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.DoughnutChart
            End Select
        Else
            chtSATCSLS1_X.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.LineChart3D
            Select Case optTotalsChartType.Value
                Case "PieChart"
                    chtSATCSLS1.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.PieChart3D
                Case "DoughnutChart"
                    chtSATCSLS1.ChartType = Infragistics.UltraChart.Shared.Styles.ChartType.DoughnutChart3D
            End Select
        End If
    End Sub

    Private Sub cbeColor_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeColor.ValueChanged
        'chtSATCSLS1.ColorModel.ModelStyle = cbeColor.ValueMember
        'chtSATCSLS1_X.ColorModel.ModelStyle = Infragistics.UltraChart.Shared.Styles.ColorModels.PureRandom
        chtSATCSLS1.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(Infragistics.UltraChart.Shared.Styles.ColorModels), cbeColor.SelectedItem.ToString()), Infragistics.UltraChart.Shared.Styles.ColorModels)
        chtSATCSLS1_X.ColorModel.ModelStyle = CType(System.Enum.Parse(GetType(Infragistics.UltraChart.Shared.Styles.ColorModels), cbeColor.SelectedItem.ToString()), Infragistics.UltraChart.Shared.Styles.ColorModels)
    End Sub
End Class