Public Class RSFSIST1

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW As String
    Dim RSTSIST1 As String
    Dim ICTITEM1 As String
    Dim filter_desc As String
    Dim sql_filter As String
    Dim sql_joins As String
    Dim sql_tables As String
    Dim EDT852TC As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        With dst
            Create_TDA(.Tables.Add, "ASTFLTR1", "*", 0, False)
            .Tables("ASTFLTR1").Columns("CODE_VALUES").MaxLength = -1

            Create_RSTSIST1()
            ASCMAIN1.sql = "Select RSTSIST1.ITEM_CODE, RSTSIST1.CUST_CODE_S, ICTITEM1.ITEM_DESC" _
            & ", ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE" _
            & ", ICTITEM1.ITEM_PICTURE_FILENAME, ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.DEPT_CODE" _
            & ", ICTITEM1.CUST_CODE, ICTITEM1.LAUNCH_DATE" _
            & ", ICTITEM1.METAL_CLASS_CODE, ICTITEM1.ITEM_CLASS_CODE, PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE) PRICE_POINT_CODE" _
            & ", ICTITEM1.MATL_CODE, ICTITEM1.MATL_CATGY_CODE, ICTITEM1.STYLE_CODE" _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE" _
            & ", RSTSIST1.RU, RSTSIST1.RS, RSTSIST1.RQ, RSTSIST1.WU, RSTSIST1.WS " _
            & ", RSTSIST1.LRU, RSTSIST1.LRS, RSTSIST1.LRQ, RSTSIST1.LWU, RSTSIST1.LWS " _
            & " from " _
            & RSTSIST1 & " RSTSIST1, " & ICTITEM1 & " ICTITEM1 where ICTITEM1.ITEM_CODE = RSTSIST1.ITEM_CODE"
            Create_TDA(.Tables.Add, "RSTSIST1", "**", 0, False, , 2)
            .Tables("RSTSIST1").Columns.Add("ITEM_PICTURE", GetType(System.Byte()))

            ASCMAIN1.sql = "Select DECODE(CUST_CODE,'ECOMSALE10','ECOM',DECODE(EDI_CUST_BATCH_NO,'LGI','LGI','EDI')) CUST_TYPE, CUST_CODE" _
            & ", MAX (OPS_YYYYWW) OPS_YYYYWW, '0' SEL from EDT852T1 " _
            & " where COMPANY_CODE = '" & ASCMAIN1.SOLUTION & "' AND CUST_CODE IS NOT NULL" _
            & " group by DECODE(EDI_CUST_BATCH_NO,'LGI','LGI','EDI'),CUST_CODE"
            EDT852TC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TC & " Add Primary Key (CUST_TYPE, CUST_CODE)")
            Create_TDA(.Tables.Add("EDT852TC"), EDT852TC, "**", 0, True, , 2)
            '.Tables("EDT852TC").Columns.Add("SEL")


        End With

        grdRSTSIST1.DataSource = dst.Tables("RSTSIST1")

        Load_EDT852TC()
        grdEDT852TC.DataSource = dst.Tables("EDT852TC")

        With grdRSTSIST1.DisplayLayout.Bands("RSTSIST1")
            For Each COLUMN_NAME As String In New String() _
            {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "PROD_CODE", "METAL_CLASS_CODE", "ITEM_CLASS_CODE", "CUST_CODE", "LAUNCH_DATE", _
             "ITEM_CATGY_CODE", "DEPT_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE", "PRICE_POINT_CODE", "MATL_CODE", "MATL_CATGY_CODE", "STYLE_CODE"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
            Next
            .Columns("ITEM_CODE").Header.Fixed = True

            Create_Summary(grdRSTSIST1, "ITEM_CODE", "Count")
            For Each COLUMN_NAME As String In New String() {"RU", "RS", "RQ", "WU", "WS", "LRU", "LRS", "LRQ", "LWU", "LWS"}
                .Columns(COLUMN_NAME).Format = "###,##0"
                .Columns(COLUMN_NAME).Width = 100
                Create_Summary(grdRSTSIST1, COLUMN_NAME)
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = IIf(COLUMN_NAME.StartsWith("R") Or COLUMN_NAME.StartsWith("LR"), Color.LightGreen, Color.LightBlue)
            Next
        End With

        Absx1.cmbFor("RYP0").Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -11)
        Absx1.cmbFor("RYP1").Value = ASCMAIN1.CYP

        ASCMAIN1.Add_Value_List(grdRSTSIST1, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdRSTSIST1, "MATL_CODE")
        ASCMAIN1.Add_Value_List(grdRSTSIST1, "PRICE_POINT_CODE")

        TAC.RSCMAIN1.Load_Filter(New String() _
        {"TRADE_CLASS_CODE", "MARKET_CODE"}, grdASTFLTR1, Me)

        Show_Filter(grdRSTSIST1, True)
        grdRSTSIST1.DisplayLayout.GroupByBox.Hidden = False

        Sort_grdColumns(grdEDT852TC, "CUST_CODE")

        With grdRSTSIST1.DisplayLayout.Bands(0)
            .Columns("ITEM_PICTURE_FILENAME").Hidden = True
            .Columns("ITEM_PICTURE").Hidden = True
            .Columns("METAL_CLASS_CODE").Hidden = True
            .Columns("MATL_CODE").Hidden = True
            .Columns("MATL_CATGY_CODE").Hidden = True
        End With

        'ASCMAIN1.Add_Value_List(grdRSTSIST1, "MATL_CATGY_CODE", , New String() {":", "Z:Any"}, , "SELECT MATL_CATGY_CODE, MATL_CATGY_DESC FROM ICTMATLA")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                RYP0 = Absx1.cmbFor("RYP0").Value
                RYP1 = Absx1.cmbFor("RYP1").Value

                If RYP0 > RYP1 Then
                    EMsg &= vbCr & "Starting Period may not be later than Ending Period"
                End If

                If EMsg = "" Then
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP1)
                    ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYPP = '" & RYP1 & "'"
                    RYW = ASCDATA1.GetDataValue
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

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Scope").Visible = Not ScreenMode
                .Groups("Options").Visible = ScreenMode
                .Groups("Filters").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdRSTSIST1.Visible = tf
        Set_Read_Only_for_ctl(chkHISTCAT, ScreenMode)

        If ScreenMode Then
            grdEDT852TC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            Clear_Record()
            grdEDT852TC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("RSTSIST1").Rows.Clear()
        EnforceConstraints(True)
        grdRSTSIST1.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Set_CUST_TYPE()
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Save_Header_Fields(UltraGroupBox1)

        grdRSTSIST1.DisplayLayout.Bands(0).SortedColumns.Clear()

        EnforceConstraints(False)
        Create_RSTSIST1()
        Fill_Records("RSTSIST1")
        EnforceConstraints(True)

        Setup_Pictures()

        ASCMAIN1.Progress("Now Setting Up Screen")

        grdRSTSIST1.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdRSTSIST1, "ITEM_CODE")
        Show_Columns()

        grdRSTSIST1.DisplayLayout.Bands(0).Columns("CUST_CODE_S").Hidden = Not chkByCustomer.Checked

        Set_grdRSTSIST1_filter()

        If chkByCustomer.Checked Then
            grdRSTSIST1.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE_S", False, True)
        End If
        grdRSTSIST1.DisplayLayout.Bands(0).SortedColumns.Add("DEPT_CODE", False, True)

        With grdRSTSIST1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"LRU", "LRS", "LRQ", "LWU", "LWS"}
                .Columns(COLUMN_NAME).Hidden = Not chkShowLY.Checked
            Next
        End With

        grdRSTSIST1.Rows.Refresh(UltraWinGrid.RefreshRow.RefreshDisplay)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTSIST1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdEDT852TC, "BBBBB", "Select All", "Clear All", "Save List", "Load List", "Maintain Lists")
    End Sub

      Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)
        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                'Case "grdSATCSLS1"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All", "Clear All"
                Set_SEL(IIf(e.Tool.Key = "Select All", "1", "0"))

            Case "Load List"
                Dim LIST_CODE As String = View_Lookup(Nothing, "LIST_CODE", "", "", "COLUMN_NAME = 'CUST_CODE'")
                If LIST_CODE <> "" Then
                    ASCMAIN1.sql = "Select CODE_VALUE from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'"
                    Dim CUST_CODEs As New List(Of String)
                    For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                        CUST_CODEs.Add(row.Item("CODE_VALUE"))
                    Next
                    For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = '" & optCustomers.Value & "'")
                        row.Item("SEL") = "0"
                        If CUST_CODEs.Contains(row.Item("CUST_CODE")) Then
                            row.Item("SEL") = "1"
                        End If
                    Next
                End If

            Case "Save List"
                Dim PSCs As String = ""
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = '" & optCustomers.Value & "' and SEL = '1'")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    PSCs &= Chr(0) & CUST_CODE
                Next

                If PSCs.Length = 0 Then
                    MsgBox("No Customers Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE")
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = PSCs
                Using frmASFCODE1 As New ASFCODE1
                    frmASFCODE1.EntryMode = "S"
                    frmASFCODE1.ShowDialog()
                End Using

            Case "Maintain Lists"
                ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE")
                ASCMAIN1.CodeSelector.MultipleSelections = True
                Using frmASFCODE1 As New ASFCODE1
                    frmASFCODE1.ShowDialog()
                End Using

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "OPS_YYYYPP"
            '    If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select

    End Sub

#End Region

    Sub Create_RSTSIST1()

        If RYP1 <> "" Then
            Set_Filter()
        End If

        Dim sql_CUST_CODE As String = ""
        If chkByCustomer.Checked Or RYP1 = "" Then
            sql_CUST_CODE = ", X.CUST_CODE"
        Else
            sql_CUST_CODE = ", 'X'"
        End If

        Dim CONS_CUST As String = "CASE WHEN LENGTH(CUST_CODE) = 9 AND (CUST_CODE LIKE 'NEIMANM%' OR CUST_CODE LIKE 'SAKSFIF%') THEN SUBSTR(CUST_CODE,1,7) || '10' ELSE DECODE(CUST_CODE,'NEIMANPJ10','NEIMANM10',CUST_CODE) END"


        Dim sqlRSTSIST1 As String = "Select ITEM_CODE, " & CONS_CUST & " CUST_CODE_S" & vbCrLf _
        & ", SUM (RU) RU, SUM (RS) RS, SUM (RQ) RQ, SUM (WU) WU, SUM (WS) WS" & vbCrLf _
        & ", SUM (LRU) LRU, SUM (LRS) LRS, SUM (LRQ) LRQ, SUM (LWU) LWU, SUM (LWS) LWS" & vbCrLf _
        & " from (" & vbCrLf _
        & " Select X.ITEM_CODE" & sql_CUST_CODE & " CUST_CODE" & vbCrLf _
        & ", 0 RU, 0 RS, 0 RQ, SUM (X.ORDR_QTY_SHIP) WU, SUM (X.ORDR_QTY_SHIP * X.ORDR_UNIT_PRICE) WS" & vbCrLf _
        & ", 0 LRU, 0 LRS, 0 LRQ, 0 LWU, 0 LWS" & vbCrLf _
        & " from SOTINVH2 X" & sql_tables & vbCrLf _
        & " where X.ORDR_YYYYPP_UPDATED >= 'TYP000' and X.ORDR_YYYYPP_UPDATED <= 'TYP001'" & vbCrLf _
        & IIf(chkGross.Checked, " and X.INV_TYPE = 'I'", "") & vbCrLf _
        & sql_joins & sql_filter & vbCrLf _
        & " group by X.ITEM_CODE" & sql_CUST_CODE & vbCrLf _
        & " union " & vbCrLf _
        & " Select X.ITEM_CODE" & sql_CUST_CODE & " CUST_CODE" & vbCrLf _
        & ", 0 RU, 0 RS, 0 RQ, 0 WU, 0 WS" & vbCrLf _
        & ", 0 LRU, 0 LRS, 0 LRQ, SUM (X.ORDR_QTY_SHIP) LWU, SUM (X.ORDR_QTY_SHIP * X.ORDR_UNIT_PRICE) LWS" & vbCrLf _
        & " from SOTINVH2 X" & sql_tables & vbCrLf _
        & " where X.ORDR_YYYYPP_UPDATED >= 'LYP000' and X.ORDR_YYYYPP_UPDATED <= 'LYP001'" & vbCrLf _
        & IIf(chkGross.Checked, " and X.INV_TYPE = 'I'", "") & vbCrLf _
        & sql_joins & sql_filter & vbCrLf _
        & " group by X.ITEM_CODE" & sql_CUST_CODE & vbCrLf _
        & " union " & vbCrLf _
        & " Select X.ITEM_CODE" & sql_CUST_CODE & " CUST_CODE" & vbCrLf _
        & ", SUM (X.QTY_SOLD) RU, SUM (X.AMT_SOLD) RS, SUM (DECODE(X.OPS_YYYYWW,'TYW000',X.QTY_EOW,0)) RQ, 0 WU, 0 WS" & vbCrLf _
        & ", 0 LRU, 0 LRS, 0 LRQ, 0 LWU, 0 LWS" & vbCrLf _
        & " from RSTRETL1 X" & sql_tables & vbCrLf _
        & " where X.OPS_YYYYPP >= 'TYP000' and X.OPS_YYYYPP <= 'TYP001'" & vbCrLf _
        & sql_joins & sql_filter & vbCrLf _
        & " group by X.ITEM_CODE" & sql_CUST_CODE & vbCrLf _
        & " union " & vbCrLf _
        & " Select X.ITEM_CODE" & sql_CUST_CODE & " CUST_CODE" & vbCrLf _
        & ", 0 RU, 0 RS, 0 RQ, 0 WU, 0 WS" & vbCrLf _
        & ", SUM (X.QTY_SOLD) RU, SUM (X.AMT_SOLD) RS, SUM (DECODE(X.OPS_YYYYWW,'LYW000',X.QTY_EOW,0)) RQ, 0 WU, 0 WS" & vbCrLf _
        & " from RSTRETL1 X" & sql_tables & vbCrLf _
        & " where X.OPS_YYYYPP >= 'LYP000' and X.OPS_YYYYPP <= 'LYP001'" & vbCrLf _
        & sql_joins & sql_filter & vbCrLf _
        & " group by X.ITEM_CODE" & sql_CUST_CODE & vbCrLf _
        & ") group by ITEM_CODE, " & CONS_CUST

        If RYP1 = "" Then
            ASCMAIN1.sql = sqlRSTSIST1
            RSTSIST1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTSIST1 & " Add Primary Key (ITEM_CODE, CUST_CODE_S)")
            ICTITEM1 = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP1, chkHISTCAT.Checked)
        Else
            RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP1, chkHISTCAT.Checked, ICTITEM1)
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSIST1)
            ASCMAIN1.sql = sqlRSTSIST1
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "TYP000", RYP0)
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "TYP001", RYP1)
            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "TYW000", RYW)
            If chkShowLY.Checked Then
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "LYP000", ASCMAIN1.Period_Calc(RYP0, -12))
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "LYP001", ASCMAIN1.Period_Calc(RYP1, -12))
                ASCMAIN1.sql = Replace(ASCMAIN1.sql, "LYW000", ASCMAIN1.Week_Calc(RYW, -52))
            End If
            ASCDATA1.ExecuteSQL("Insert into " & RSTSIST1 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Private Sub grdRSTSIST1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTSIST1.InitializeRow

    End Sub

    Private Sub chkSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUnits.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkRetailSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRetail.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkWSSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkWS.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkSales_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSales.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Sub Show_Columns()

    End Sub

    Sub Setup_Pictures()
        Dim IMAGE_FOLDER As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then IMAGE_FOLDER = "C:\Documents and Settings\wjz\Desktop\Clients\JHI\Images\"

        For Each rowRSTSIST1 As DataRow In dst.Tables("RSTSIST1").Rows
            Dim ITEM_PICTURE_FILENAME As String = rowRSTSIST1.Item("ITEM_PICTURE_FILENAME") & ""

            'testing
            If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
                ITEM_PICTURE_FILENAME = "b.jpg"
            End If

            If ITEM_PICTURE_FILENAME <> "" Then
                Dim FILENAME As String = IMAGE_FOLDER & ITEM_PICTURE_FILENAME
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    rowRSTSIST1.Item("ITEM_PICTURE") = ASCMAIN1.GetImageData(FILENAME) 'ASCMAIN1.Get_Image(IMAGE_FOLDER, ITEM_PICTURE_FILENAME)
                End If
            End If
        Next
    End Sub

    Private Sub chkExclInactive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExclInactive.CheckedChanged
        Set_grdRSTSIST1_filter()
    End Sub

    Private Sub chkExclNoSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExclNoSales.CheckedChanged
        Set_grdRSTSIST1_filter()
    End Sub

    Sub Set_grdRSTSIST1_filter()
        If SELECTION_NO = 0 Then Exit Sub
        Dim dvw As DataView = DirectCast(grdRSTSIST1.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        Dim row_filters As String = ""
        If chkExclInactive.Checked Then
            sql &= " and ITEM_CATGY_CODE <> 'I'"
            row_filters &= ", Excluding Inactive"
        End If
        If chkExclNoSales.Checked Then
            sql &= " and (ISNULL(RU,0) <> 0 or ISNULL(WU,0) <> 0)"
            row_filters &= ", Excluding Items with No Sales"
        End If
        If sql <> "" Then
            sql = Mid(sql, 5)
        End If
        dvw.RowFilter = sql

        Dim MOs As String = Format(ASCMAIN1.Period_Diff(RYP0, RYP1) + 1, "00")
        grdRSTSIST1.Text = "Sell-In / Sell-Thru for the " & MOs & " Months from " & Absx1.cmbFor("RYP0").Text & " to " & Absx1.cmbFor("RYP1").Text & filter_desc & row_filters

    End Sub

    Sub Set_Filter()

        grdASTFLTR1.UpdateData()

        Dim filter_count As Integer = 0
        filter_desc = ""
        sql_filter = ""
        Dim sql_filters As New Dictionary(Of String, String)

        sql_joins = ""
        sql_tables = ""

        sql_filters.Clear()
        Dim sql As String = ""
        For Each rowASTFLTR1 As DataRow In dst.Tables("ASTFLTR1").Rows
            Dim CODE_VALUES As String = rowASTFLTR1.Item("CODE_VALUES") & ""
            Dim filter_COL As String = rowASTFLTR1.Item("COLUMN_NAME") & ""
            Dim filter_CAP As String = rowASTFLTR1.Item("COLUMN_CAPTION") & ""
            sql = ""
            If CODE_VALUES <> "" Then
                Dim TABLE_NAME As String = "X"
                Select Case filter_COL
                    Case "MARKET_CODE"
                        TABLE_NAME = "SOTTCLS1"
                        If Not sql_tables.Contains(",ARTCUST1") Then
                            sql_tables &= ",ARTCUST1"
                            sql_joins &= " and ARTCUST1.CUST_CODE = X.CUST_CODE"
                        End If
                        If Not sql_tables.Contains(",SOTTCLS1") Then
                            sql_tables &= ",SOTTCLS1"
                            sql_joins &= " and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE"
                        End If
                    Case "TRADE_CLASS_CODE"
                        TABLE_NAME = "ARTCUST1"
                        If Not sql_tables.Contains(",ARTCUST1") Then
                            sql_tables &= ",ARTCUST1"
                            sql_joins &= " and ARTCUST1.CUST_CODE = X.CUST_CODE"
                        End If
                End Select

                filter_count += 1
                If InStr(CODE_VALUES, ",") = 0 Then
                    sql = " and " & TABLE_NAME & "." & filter_COL & " = '" & CODE_VALUES & "'"
                    filter_desc &= "," & filter_CAP & " " & CODE_VALUES
                    'filter_desc = filter_CAP & " " & CODE_VALUES
                Else
                    sql = " and " & TABLE_NAME & "." & filter_COL & " in ('" & Replace(CODE_VALUES, ",", "','") & "')"
                    filter_desc &= "," & filter_CAP & " " & CODE_VALUES
                    'filter_desc = "Selected " & filter_CAP & "s"
                End If

            End If
            sql_filter &= sql
            sql_filters.Add(filter_COL, sql)
        Next

        ASCDATA1.ExecuteSQL("Delete from " & EDT852TC)
        dst.Tables("EDT852TC").AcceptChanges()
        For Each rowEDT852TC As DataRow In dst.Tables("EDT852TC").Select
            rowEDT852TC.SetAdded()
        Next
        Update_Record_TDA("EDT852TC")
        If chkAllCusts.Checked Then
            sql = " and X.CUST_CODE in (SELECT CUST_CODE from ARTCUST1 " _
            & " where TRADE_CLASS_CODE in ('NAT','IND')) " _
            & "   and X.CUST_CODE <> 'SAKSOFF5TH' and X.CUST_CODE <> 'NMLASTCALL' and X.CUST_CODE <> 'BLOUTLET10'"
            sql_filter &= sql
            sql_filters.Add("CUST_CODE", sql)
            filter_desc = ", All Customers except NEIMANM20,SAKSOFF5TH,NMLASTCALL,BLOUTLET10"
        Else
            If dst.Tables("EDT852TC").Select("SEL<>'1'").Length <> 0 Then
                filter_count += 1
                sql = " and X.CUST_CODE in (Select CUST_CODE from " & EDT852TC & " where SEL = '1')"
                sql_filter &= sql
                sql_filters.Add("CUST_CODE", sql)
            End If

            If filter_count = 0 And dst.Tables("EDT852TC").Select("SEL<>'1'").Length = 0 Then
                filter_desc = ", All Records"
            Else
                'filter_desc = Mid(filter_desc, 2)
                If filter_count > 1 Then
                    'filter_desc = "Filtered Records"
                End If
            End If
        End If


    End Sub

    Private Sub grdASTFLTR1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTFLTR1.BeforeRowUpdate

        Dim COLUMN_NAME As String = e.Row.Cells("COLUMN_NAME").Text
        Dim sql As String = ASCMAIN1.CodeSelector.Get_SQL(COLUMN_NAME)
        If sql <> "" Then
            Dim CODE_VALUES_new As String = ""
            Dim CODE_VALUES As String = e.Row.Cells("CODE_VALUES").Text
            Dim KEY_EXPRESSION As String = ASCMAIN1.CodeSelector.grdColumns(0).Item("COLUMN_NAME")

            If CODE_VALUES <> "" Then
                Dim CODE_VALUES_old As String = ""
                For Each txt As String In Split(Replace(CODE_VALUES, "'", ""), ",")
                    CODE_VALUES_old = CODE_VALUES_old & ",'" & ASCMAIN1.Format_Field(txt, COLUMN_NAME, , True) & "'"
                Next
                CODE_VALUES_old = Mid$(CODE_VALUES_old, 2)
                Dim where_or_and As String = " where "
                If ASCMAIN1.CodeSelector.tblASTVIEW1.Rows(0).Item("WHERE_CLAUSE") & "" <> "" Then
                    where_or_and = " and "
                End If

                For Each dr As DataRow In ASCDATA1.GetDataTable(sql & where_or_and & KEY_EXPRESSION & " IN (" & CODE_VALUES_old & ")").Rows
                    CODE_VALUES_new = CODE_VALUES_new & "," & dr.Item(0)
                Next
            End If

            CODE_VALUES_new = Mid(CODE_VALUES_new, 2)
            If CODE_VALUES_new <> CODE_VALUES Then
                grdASTFLTR1.DisplayLayout.Rows(e.Row.Index).Cells("CODE_VALUES").Value = CODE_VALUES_new
            End If
        End If
    End Sub

    Private Sub grdASTFLTR1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTFLTR1.ClickCellButton
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL(grdASTFLTR1.ActiveRow.Cells("COLUMN_NAME").Text)
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Replace(grdASTFLTR1.ActiveRow.Cells("CODE_VALUES").Text & "", ",", Chr(0))
            Dim F As New ASFCODE1
            F.ShowDialog()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                grdASTFLTR1.ActiveRow.Cells("CODE_VALUES").Value = Mid$(Replace(ASCMAIN1.CodeSelector.SelectedCodes0, Chr(0), ","), 2)
                grdASTFLTR1.UpdateData()
            End If
        End If
    End Sub

    Sub Load_EDT852TC()
        Fill_Records("EDT852TC")
        'ALL_EDI_CUSTS = dst.Tables("EDT852TC").Select("CUST_TYPE = 'EDI'").Length

        For Each rowEDT852TC As DataRow In dst.Tables("EDT852TC").Select("", "CUST_CODE")
            rowEDT852TC.Item("SEL") = "1"
            Dim CUST_CODE As String = rowEDT852TC.Item("CUST_CODE")
            If rowEDT852TC.Item("CUST_TYPE") = "EDI" Then
                'EDI_CUSTs.Add(CUST_CODE)
            ElseIf rowEDT852TC.Item("CUST_TYPE") = "ECOM" Then
                'ECOM_CUSTs.Add(CUST_CODE)
            Else
                'LGI_CUSTs.Add(CUST_CODE)
            End If
        Next
    End Sub

    Private Sub optCustomers_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCustomers.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_CUST_TYPE()
    End Sub

    Sub Set_CUST_TYPE()
        Dim dvw As DataView = DirectCast(grdEDT852TC.DataSource, DataTable).DefaultView
        dvw.RowFilter = "CUST_TYPE = '" & optCustomers.Value & "'"
    End Sub

    Sub Set_SEL(ByVal SEL As String)
        For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = '" & optCustomers.Value & "'")
            row.Item("SEL") = SEL
        Next
    End Sub
End Class