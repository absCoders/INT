Public Class SAFSLSX1

    Dim DTE1 As Date
    Dim DTE2 As Date
    Dim SATSLSXC As String
    Dim SATSLSXI As String
    Dim SATSLSXS As String
    Dim FILENAME As String = ""

    Dim SOTINVHX As String
    Dim sqlSOTINVH1_S As String = ""
    ' ON HANDS ARE ALWAYS AS OF TODAY

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then

                ASCMAIN1.sql = "Select INV_TYPE, INV_NO from SOTINVH1 where ROWNUM < 1"
                SOTINVHX = ASCMAIN1.Temp_Table()
                ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHX & " Add Primary Key (INV_TYPE,INV_NO)")

                sqlSOTINVH1_S = "Select SOTINVH1.INV_TYPE,SOTINVH1.INV_NO,SOTINVH1.CUST_CODE,SOTINVH1.CUST_STORE_NO,ARTCUST2.CUST_STORE_NAME,SOTINVH1.ORDR_CUST_PO," _
                    & "SOTINVH1.ORDR_NO,SOTINVH1.WHSE_CODE,SOTINVH1.POST_CODE,SOTINVH1.TERM_CODE,SOTINVH1.SREP_CODE," _
                    & "SOTINVH1.REASON_CODE,SOTINVH1.CUST_BILL_TO_CUST,SOTINVH1.EVENT_CODE,SOTINVH1.INV_SALES,SOTINVH1.INV_COGS," _
                    & "SOTINVH1.INV_FREIGHT,SOTINVH1.INV_MISC_CHG,SOTINVH1.INV_TOTAL_AMOUNT,SOTINVH1.INV_DATE," _
                    & "SOTINVH1.ORDR_YYYYPP_UPDATED,SOTINVH1.INIT_DATE,SOTINVH1.INIT_OPER,SOTINVH1.REGISTER_XNO,SOTINVH1.INV_DATE_SHIPPED," _
                    & "SOTINVH1.INV_CARTONS,SOTINVH1.INV_WEIGHT,SOTINVH1.INV_BOL_NO,SOTINVH1.INV_PRO_NO," _
                    & "SOTINVH1.SHIP_VIA_DESC,SOTINVH1.INV_NO_CONS,SOTINVH1.SHIP_BOL_NO,SOTINVH1.PICK_NO," _
                    & "SOTINVH1.CUST_SHIP_TO_STATE,SOTINVH1.INV_COMMENT,SOTINVH1.SHIP_VIA_CODE," _
                    & "SOTINVH1.ORDR_TYPE_CODE,SOTINVH1.OPS_YYYYWW,SOTINVH1.ORDR_DEPT," _
                    & "SOTINVH1.REGISTER_DATE,SOTINVH1.INV_UNITS,SOTINVH1.SALES_DIVISION_CODE," _
                    & "SOTINVH1.MISC_CHG_CODE,SOTINVH1.REGISTER_IND,SOTINVH1.SHIP_FRT_AMT_ACCRUED," _
                    & "SOTINVH1.SHIP_FRT_AMT_ACTUAL,SOTINVH1.SELL_CODE,ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & ",SOTINVH2.INV_LNO,SOTINVH2.ITEM_CODE,ICTITEM1.ITEM_DESC,SOTINVH2.ORDR_UNIT_PRICE" & vbCrLf _
                    & ",SOTINVH2.ORDR_QTY_SHIP,SOTINVH2.ITEM_UNIT_COST,SOTINVH2.ITEM_RETAIL_PRICE" & vbCrLf _
                    & ",ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTITEM1.COST_CATGY_CODE, ICTCOLL1.BRAND_CODE, ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
                    & ",SOTINVH2.ORDR_UNIT_PRICE * SOTINVH2.ORDR_QTY_SHIP LINE_AMT" & vbCrLf _
                    & ",SOTINVH2.ITEM_RETAIL_PRICE * SOTINVH2.ORDR_QTY_SHIP LINE_AMT_RTL" & vbCrLf _
                    & " from " & SOTINVHX & " SOTINVHX,SOTINVH1,SOTINVH2,ICTITEM1,ICTCOLL1,ARTCUST2,ARTCUST1" & vbCrLf _
                    & " where SOTINVH1.INV_TYPE = SOTINVHX.INV_TYPE And SOTINVH1.INV_NO = SOTINVHX.INV_NO" & vbCrLf _
                    & "   And SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE And SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                    & "   And ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                    & "   And ARTCUST2.CUST_CODE (+) = SOTINVH1.CUST_CODE" & vbCrLf _
                    & "   And ARTCUST2.CUST_STORE_NO (+) = SOTINVH1.CUST_STORE_NO" & vbCrLf _
                    & "   And ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & "   And ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE"
                ASCMAIN1.sql = sqlSOTINVH1_S
                Create_TDA(.Tables.Add, "SOTINVH1_S", "**", 0, False, "", 0)
                '.Tables("SOTINVH1_S").Columns.Add("LINE_AMT", GetType(System.Decimal), "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
                '.Tables("SOTINVH1_S").Columns.Add("LINE_AMT_RTL", GetType(System.Decimal), "ORDR_QTY_SHIP * ITEM_RETAIL_PRICE")
            Else

                TAC.SACMAIN1.Create_Sales_Extract_Tables(Me, True, SATSLSXC, SATSLSXI, SATSLSXS)

                'ASCMAIN1.sql = "Select * from " & SATSLSXC
                'Create_TDA(.Tables.Add, "SATSLSXC", "**", 0, False, "", 0)

                'ASCMAIN1.sql = "Select * from " & SATSLSXI
                'Create_TDA(.Tables.Add, "SATSLSXI", "**", 0, False, "", 0)

                'ASCMAIN1.sql = "Select * from " & SATSLSXS
                'Create_TDA(.Tables.Add, "SATSLSXS", "**", 0, False, "", 0)
            End If
        End With

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then

            grdSOTINVH1_S.DataSource = dst.Tables("SOTINVH1_S")
            Create_Summary(grdSOTINVH1_S, "INV_NO", "Count")
            Create_Summary(grdSOTINVH1_S, "ORDR_QTY_SHIP")
            Create_Summary(grdSOTINVH1_S, "LINE_AMT")
        Else
            grdSATSLSXC.DataSource = dst.Tables("SATSLSXC")
            grdSATSLSXI.DataSource = dst.Tables("SATSLSXI")
            grdSATSLSXS.DataSource = dst.Tables("SATSLSXS")

            grdXLS.DataSource = dst.Tables("SATSLSXS")


            Create_Summary(grdSATSLSXC, grdSATSLSXC.DisplayLayout.Bands(0).Columns(0).Key, "Count")
            Create_Summary(grdSATSLSXI, grdSATSLSXI.DisplayLayout.Bands(0).Columns(0).Key, "Count")
            Create_Summary(grdSATSLSXS, grdSATSLSXS.DisplayLayout.Bands(0).Columns(0).Key, "Count")
            Create_Summary(grdXLS, grdXLS.DisplayLayout.Bands(0).Columns(0).Key, "Count")

        End If

        dteTo.Value = Now.AddDays(-1)
        dteFrom.Value = Now.AddDays(-7)

        ASCMAIN1.sql = "Select Distinct REGISTER_XNO from SOTINVH1" _
            & " where ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6) & "'" _
            & "   and ORDR_YYYYPP_UPDATED <= '" & ASCMAIN1.CYP & "'" _
            & "   and REGISTER_XNO is Not Null"

        Dim tbl As DataTable = New DataView(ASCDATA1.GetDataTable, "", "REGISTER_XNO DESC", DataViewRowState.CurrentRows).ToTable
        cbeXNO.DataSource = tbl
        cbeXNO.SelectedItem = cbeXNO.Items(0)
        'cbeXNO.Value = tbl.Rows(0)

        chkCGS.Visible = (ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT")
 
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 12, 0) ' -11 - 12)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 12, 0) '  0 - 12)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                DTE1 = dteFrom.Value
                DTE2 = dteTo.Value

                If optScope.Value = "X" Then
                    If cbeXNO.Text = "" Then
                        EMsg &= vbCr & "You must specify a valid Sales Journal XNO"
                    Else
                        ASCMAIN1.sql = "Select Count (*) from SOTINVH1 where REGISTER_XNO = '" & cbeXNO.Text & "'"
                        Dim recs As Integer = Val(ASCDATA1.GetDataValue)
                        If recs = 0 Then
                            EMsg &= vbCr & "No Records on file for Sales Journal XNO " & cbeXNO.Text
                        End If
                    End If
                ElseIf optScope.Value = "D" Then
                    If Format(DTE1, "yyyyMMdd") > Format(DTE2, "yyyyMMdd") Then
                        EMsg &= vbCr & "From Date cannot be later than To Date"
                    End If
                ElseIf optScope.Value = "P" Then
                    If ASCMAIN1.CLIENT = "AHA" Then
                        EMsg &= vbCr & "Period Range NOT Supported for this function (Israel uses dates)"
                        ' would need to change extract to work with periods
                    End If
                    If cmbFrom.Value > cmbTo.Value Then
                        EMsg &= vbCr & "From Period cannot be later than To Period"
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

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Show Temp Folder"
                Process.Start("explorer.exe", ASCMAIN1.Folders("Temp"))

            Case "Done"
                Mode_Settings(False)

            Case "ftp File"
                'ftp_File()
                TAC.SACMAIN1.ftp_BI_Files(Me)

            Case "Print"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("ftp File").Settings.Enabled = iScreenMode
                    .Items("Show Temp Folder").Visible = ScreenMode

                    If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                        .Items("ftp File").Visible = False
                        .Items("Show Temp Folder").Visible = False

                    End If

                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabData.Visible = tf
        UltraExplorerBarContainerControl1.Enabled = Not tf

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            For Each TABLE_NAME As String In New String() {"SOTINVH1_S"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        Else
            For Each TABLE_NAME As String In New String() {"SATSLSXC", "SATSLSXI", "SATSLSXS"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
        End If

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data from Database")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then

            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVHX)
            ASCMAIN1.sql = "Insert into " & SOTINVHX & vbCrLf _
                & " Select INV_TYPE, INV_NO from SOTINVH1" & vbCrLf _
                & " where " & IIf(optScope.Value = "D", _
                                    " INV_DATE  between '" & Format(dteFrom.Value, "dd-MMM-yyyy") & "' and '" & Format(dteTo.Value, "dd-MMM-yyyy") & "'", _
                                  IIf(optScope.Value = "P", _
                                    " ORDR_YYYYPP_UPDATED  between '" & cmbFrom.Value & "' and '" & cmbTo.Value & "'", _
                                    " REGISTER_XNO = '" & cbeXNO.Value & "'"))
            ASCDATA1.ExecuteSQL()

            If chkCGS.Checked Then
                tabData.Tabs("Customers").Visible = False
                tabData.Tabs("Items").Visible = False
                tabData.Tabs("Sales").Visible = False
                tabData.Tabs("Extracts").Visible = False
                tabData.Tabs("Shipments && Returns").Visible = False
                grpSSG.Dock = DockStyle.None
                grpSSG.Parent = tabData.Tabs("CGS").TabPage
                grpSSG.Dock = DockStyle.Fill

                ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE" & vbCrLf _
                & ",SOTINVH1.INV_NO" & vbCrLf _
                & ",SOTINVH1.CUST_CODE" & vbCrLf _
                & ",SOTINVH1.CUST_STORE_NO" & vbCrLf _
                & ",SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                & ",SOTINVH1.EVENT_CODE" & vbCrLf _
                & ",SOTINVH1.INV_DATE" & vbCrLf _
                & ",SOTINVH1.CUST_SHIP_TO_STATE" & vbCrLf _
                & ",SOTORDR5.CUST_CITY CUST_SHIP_TO_CITY" & vbCrLf _
                & ",SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
                & ",ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & ",SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                & ",ICTCOLL1.BRAND_CODE" & vbCrLf _
                & ",ICTCOLL1.HC_CODE" & vbCrLf _
                & ",SOTINVH2.ITEM_CODE" & vbCrLf _
                & ",ICTITEM1.ITEM_DESC" & vbCrLf _
                & ",ICTITEM1.PROD_CODE" & vbCrLf _
                & ",ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ",ICTITEM1.COST_CATGY_CODE" & vbCrLf _
                & ",ICTITEM1.ITEM_SNU_CODE" & vbCrLf _
                & ",ICTITEM1.COUNTRY_CODE ITEM_COUNTRY_CODE" & vbCrLf _
                & ",SOTINVH2.ITEM_RETAIL_PRICE" & vbCrLf _
                & ",SOTINVH2.ORDR_UNIT_PRICE" & vbCrLf _
                & ",SOTINVH2.ORDR_QTY_SHIP" & vbCrLf _
                & ",SOTINVH2.ITEM_UNIT_COST" & vbCrLf _
                & ",ICTCOSTA.ITEM_COST_TOOLG STD_TARIFF" & vbCrLf _
                & ",NVL(SOTINVH2.ITEM_UNIT_COST - ICTCOSTA.ITEM_COST_TOOLG, SOTINVH2.ITEM_UNIT_COST) UNIT_COST_NO_TARIFF" & vbCrLf _
                & ",SOTINVH2.ORDR_QTY_SHIP * NVL(SOTINVH2.ITEM_UNIT_COST - ICTCOSTA.ITEM_COST_TOOLG, SOTINVH2.ITEM_UNIT_COST) EXT_CGS_NO_TARIFF" & vbCrLf _
                & ",SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE SLS" & vbCrLf _
                & ",SOTINVH2.ORDR_QTY_SHIP * DECODE(SOTINVH2.WHSE_CODE,NULL,0,SOTINVH2.ITEM_UNIT_COST) CGS" & vbCrLf _
                & ",SOTINVH2.ACCT_CODE_SLS" & vbCrLf _
                & ",SOTINVH2.ACCT_CODE_CGS" & vbCrLf _
                & ",SOTINVH2.SEG2_CODE" & vbCrLf _
                & ",SOTINVH2.SEG3_CODE" & vbCrLf _
                & ",SOTINVH2.SEG4_CODE" & vbCrLf _
                & " from SOTINVH2, SOTINVH1, " & SOTINVHX & " SOTINVHX, SOTORDR5, ICTITEM1, ICTCOLL1, ARTCUST1, SOTTCLS1, ICTCOSTA" & vbCrLf _
                & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
                & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE = SOTINVHX.INV_TYPE" & vbCrLf _
                & "   and SOTINVH1.INV_NO = SOTINVHX.INV_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOSTA.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP = '" & IIf(optScope.Value = "D", Format(dteTo.Value, "yyyyMM"), cmbTo.Value) & "'" & vbCrLf _
                & "   and SOTORDR5.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'"

                Dim DataTable As DataTable = ASCDATA1.GetDataTable
                '  Brand	 	Hcollection	Trade Class	Chanel	Retail Price

                WorkbookView1.GetLock()

                'If ASCMAIN1.Running_in_VS Then
                '    Stop ' HIJACKING THIS SPREADSHEET FOR ANOTHER PURPOSE
                '    DataTable = ASCDATA1.GetDataTable("SELECT * FROM SPTCWRX2 WHERE OPS_YYYYPP < '201601'")
                '    Load_DataTable_into_SGXLS(1, 1, DataTable, WorkbookView1.ActiveWorksheet, Nothing, Nothing, "CTRL_NO", "")
                'Else
                '    Load_DataTable_into_SGXLS(1, 1, DataTable, WorkbookView1.ActiveWorksheet, Nothing, Nothing, "INV_TYPE,INV_NO", "")
                'End If
                Load_DataTable_into_SGXLS(1, 1, DataTable, WorkbookView1.ActiveWorksheet, Nothing, Nothing, "INV_TYPE,INV_NO", "")

                WorkbookView1.ReleaseLock()

            Else
                ASCMAIN1.sql = sqlSOTINVH1_S
                Dim DataTable As DataTable = ASCDATA1.GetDataTable
                Dim recs As Integer = DataTable.Rows.Count
                If recs > 100000 Then
                    If MsgBox("Too many records to display, do you want to export as a CSV?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Me.Cursor = Cursors.WaitCursor
                        Dim FILENAME As String = ASCMAIN1.Export_To_CSV(sqlSOTINVH1_S, recs, "SalesExtract")
                        Show_Document(FILENAME)
                        Me.Cursor = Cursors.Default
                    Else
                        Exit Sub
                    End If
                Else
                    Fill_Records("SOTINVH1_S")
                    Sort_grdColumns(grdSOTINVH1_S, "INV_TYPE,INV_NO")

                    tabData.Tabs("Customers").Visible = False
                    tabData.Tabs("Items").Visible = False
                    tabData.Tabs("Sales").Visible = False
                    tabData.Tabs("Extracts").Visible = False
                    tabData.Tabs("CGS").Visible = False

                    WorkbookView1.GetLock()
                    Load_DataTable_into_SGXLS(1, 1, dst.Tables("SOTINVH1_S"), WorkbookView1.ActiveWorksheet, grdSOTINVH1_S, Nothing, "INV_TYPE,INV_NO", "")
                    WorkbookView1.ReleaseLock()
                End If


            End If

        Else
            TAC.SACMAIN1.Create_Sales_Extract_Tables(Me, False, SATSLSXC, SATSLSXI, SATSLSXS, optScope.Value, cbeXNO.Text, DTE1, DTE2)

            Sort_grdColumns(grdSATSLSXC, "CUSTOMER_CODE,BRANCH_CODE")
            Sort_grdColumns(grdSATSLSXI, "ITEM_CODE")
            Sort_grdColumns(grdSATSLSXS, "IVNUM,LINE")

            tabData.Tabs("Shipments && Returns").Visible = False
            tabData.Tabs("CGS").Visible = False
        End If


        'FILENAME = Format(DTE1, "yyyyMMdd") & "_" & Format(DTE2, "yyyyMMdd")
        'TAC.TACMAIN1.Create_Zip_File(FILENAME, New String() {ASCMAIN1.Folders("Temp") & "customer.csv", _
        '                                                     ASCMAIN1.Folders("Temp") & "part.csv", _
        '                                                     ASCMAIN1.Folders("Temp") & "sales.csv"})

        'Show_Document(ASCMAIN1.Folders("Temp") & FILENAME & ".zip")

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSATSLSXC, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
        Load_Popup_Menu(grdSOTINVH1_S, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        'If tlb_pop.Tools.Exists("Include Inactive") Then
        'End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else

            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

#End Region

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Print_Report_End()
    End Sub

    Private Sub optScope_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optScope.ValueChanged
        lblXNO.Visible = (optScope.Value = "X")
        cbeXNO.Visible = (optScope.Value = "X")

        lblFrom.Visible = (optScope.Value = "D") Or (optScope.Value = "P")
        dteFrom.Visible = (optScope.Value = "D")
        lblTo.Visible = (optScope.Value = "D") Or (optScope.Value = "P")
        dteTo.Visible = (optScope.Value = "D")

        cmbFrom.Visible = (optScope.Value = "P")
        cmbTo.Visible = (optScope.Value = "P")

        cmdKSP.Visible = (optScope.Value = "P")
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        WorkbookView1.GetLock()
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("RSFSSPL1.XLSX_NO") & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        WorkbookView1.ReleaseLock()
    End Sub

    Private Sub chkCGS_CheckedChanged(sender As Object, e As EventArgs) Handles chkCGS.CheckedChanged

    End Sub

    Private Sub cmdKSP_Click(sender As Object, e As EventArgs) Handles cmdKSP.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now creating work tables")

        Dim YP1 As String = cmbFrom.Value
        Dim YP2 As String = cmbTo.Value

        Dim YP1Text As String = cmbFrom.Text
        Dim YP2Text As String = cmbTo.Text

        ASCMAIN1.sql = "" _
            & "Select" & vbCrLf _
            & "SOTRTRN1.OPS_YYYYPP, SOTRTRN1.INV_NO, SOTRTRN1.RTRN_NO, SOTRTRN1.RTRN_DATE" & vbCrLf _
            & ", SOTRTRN2.ITEM_CODE, SOTRTRN2.RTRN_QTY, SOTRTRN2.RTRN_PRICE, SOTRTRN2.ITEM_COST_STD" & vbCrLf _
            & ", SOTRTRN2.RTRN_QTY * SOTRTRN2.RTRN_PRICE AMT" & vbCrLf _
            & ", SOTRTRN2.RTRN_QTY * SOTRTRN2.ITEM_COST_STD CGR" & vbCrLf _
            & ", SOTRTRN1.CUST_NAME, SOTRTRN1.CUST_STORE_NO, SOTRTRN1.CUST_CLAIM_NO" & vbCrLf _
            & " from SOTRTRN1,SOTRTRN2" & vbCrLf _
            & " where SOTRTRN2.RTRN_NO = SOTRTRN1.RTRN_NO" & vbCrLf _
            & $"   and SOTRTRN1.OPS_YYYYPP >= '{YP1}' AND SOTRTRN1.OPS_YYYYPP <= '{YP2}'" & vbCrLf _
            & "   and SOTRTRN1.RTRN_AS_PO_REC = '1'"

        Dim tbl As DataTable = ASCDATA1.GetDataTable
        Dim frmmsg As New ASFMSGBF
        frmmsg.Show_grd(tbl, Me, $"KSP Returns entered as a PO from {YP1Text} to {YP2Text}")

        'EXCEL_SHEET = frmmsg.grow.Cells("TABLE_NAME").Text & "$"

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
End Class