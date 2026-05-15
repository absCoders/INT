Imports Infragistics.Win.UltraWinGrid
Imports Microsoft.Exchange.WebServices.Data
Imports Microsoft.Identity.Client

Public Class WHFAUDIT

    Private wkTable As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        InquiryMode = MENU_ITEM_OBJECT = "WHFAUDII"

        With dst

            Get_PARM("WHTPARM1")
            Create_TDA(.Tables.Add, "WHTAUDT1", "*")

            wkTable = ASCMAIN1.Temp_Table("SELECT AUDIT_NO FROM WHTAUDT1 WHERE ROWNUM < 1")

            ASCMAIN1.sql = $"Select * from WHTAUDT2 where AUDIT_NO IN (SELECT AUDIT_NO FROM {wkTable})"
            Create_TDA(.Tables.Add, "WHTAUDT2", ASCMAIN1.sql, 0)

            ASCMAIN1.sql = $"Select * from WHTAUDT3 where AUDIT_NO IN (SELECT AUDIT_NO FROM {wkTable})"
            Create_TDA(.Tables.Add, "WHTAUDT3", ASCMAIN1.sql, 0)

            ASCMAIN1.sql = "SELECT DISTINCT NVL(SOTSHIP1.SHIP_REF, SOTSHIP1.BILL_OF_LADING_NO) PRO_NO, SOTSHIP1.BILL_OF_LADING_NO, SOTORDR0.CUST_CODE, SOTSHIP1.SHIP_DATE_SHIPPED, 
                                        SOTSHIP1.SHIP_VIA_CODE, SOTSHIP1.WHSE_CODE, SOTSVIA1.CARRIER_CODE, ICTWHSE1.LP_CODE, COUNT(DISTINCT SOTORDR0.ORDR_CUST_PO) NUM_POS
                                        FROM SOTSHIP1, SOTSVIA1, SOTORDR0, ARTCUST1, ICTWHSE1, WHTAUDT1
                                        WHERE SOTSHIP1.SHIP_STATUS = 'F'
                                        AND SOTSHIP1.SHIP_DATE_SHIPPED >= TRUNC(WHTAUDT1.SHIP_DATE_FROM)
                                        AND TRUNC(SOTSHIP1.SHIP_DATE_SHIPPED) <= TRUNC(WHTAUDT1.SHIP_DATE_TO)
                                        AND SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)
                                        AND SOTSHIP1.BILL_OF_LADING_NO IS NOT NULL
                                        AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO
                                        AND SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE
                                        AND ARTCUST1.TRADE_CLASS_CODE <> 'HDQ'
                                        AND SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE
                                        AND ICTWHSE1.LP_CODE IS NOT NULL
                                        AND WHTAUDT1.AUDIT_NO = :PARM1
                                        GROUP BY NVL(SOTSHIP1.SHIP_REF, SOTSHIP1.BILL_OF_LADING_NO), SOTSHIP1.BILL_OF_LADING_NO, SOTORDR0.CUST_CODE, SOTSHIP1.SHIP_DATE_SHIPPED, 
                                        SOTSHIP1.SHIP_VIA_CODE, SOTSHIP1.WHSE_CODE, SOTSVIA1.CARRIER_CODE, ICTWHSE1.LP_CODE"
            Create_TDA(.Tables.Add, "AuditData", ASCMAIN1.sql, 0, False, "V", 0)

            ASCMAIN1.sql = "SELECT DISTINCT NVL(SOTSHIP1.SHIP_REF, SOTSHIP1.BILL_OF_LADING_NO) PRO_NO, SOTSHIP1.BILL_OF_LADING_NO, SOTORDR0.CUST_CODE, SOTSHIP1.SHIP_DATE_SHIPPED, 
                                        SOTSHIP1.SHIP_VIA_CODE, SOTSHIP1.WHSE_CODE, SOTSVIA1.CARRIER_CODE, ICTWHSE1.LP_CODE, SOTORDR0.ORDR_CUST_PO
                                        FROM SOTSHIP1, SOTSVIA1, SOTORDR0, ARTCUST1, ICTWHSE1, WHTAUDT1
                                        WHERE SOTSHIP1.SHIP_STATUS = 'F'
                                        AND SOTSHIP1.SHIP_DATE_SHIPPED >= TRUNC(WHTAUDT1.SHIP_DATE_FROM)
                                        AND TRUNC(SOTSHIP1.SHIP_DATE_SHIPPED) <= TRUNC(WHTAUDT1.SHIP_DATE_TO)
                                        AND SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)
                                        AND SOTSHIP1.BILL_OF_LADING_NO IS NOT NULL
                                        AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO
                                        AND SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE
                                        AND ARTCUST1.TRADE_CLASS_CODE <> 'HDQ'
                                        AND SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE
                                        AND ICTWHSE1.LP_CODE IS NOT NULL
                                        AND WHTAUDT1.AUDIT_NO = :PARM1"
            Create_TDA(.Tables.Add, "AuditDetails", ASCMAIN1.sql, 0, False, "V", 0)

            ASCMAIN1.sql = "SELECT SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO, SOTINVH1.ORDR_CUST_PO, 
                                NVL(SOTINVH1.INV_SALES, 0) INV_SALES, NVL(SOTINVH1.INV_FREIGHT, 0) INV_FREIGHT, 
                                NVL(SOTINVH1.INV_MISC_CHG, 0) INV_MISC_CHG, NVL(SOTINVH1.INV_STAX, 0) INV_STAX, NVL(SOTINVH1.INV_TOTAL_AMOUNT, 0) INV_TOTAL_AMOUNT, 
                                SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.INV_BOL_NO, SOTINVH1.INV_PRO_NO
                                FROM SOTSHIP1, SOTSVIA1, SOTORDR0, ARTCUST1, ICTWHSE1, WHTAUDT1, SOTPICK1, SOTINVH1
                                WHERE SOTSHIP1.SHIP_STATUS = 'F'
                                AND SOTSHIP1.SHIP_DATE_SHIPPED >= TRUNC(WHTAUDT1.SHIP_DATE_FROM)
                                AND TRUNC(SOTSHIP1.SHIP_DATE_SHIPPED) <= TRUNC(WHTAUDT1.SHIP_DATE_TO)
                                AND SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)
                                AND SOTSHIP1.BILL_OF_LADING_NO IS NOT NULL
                                AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO
                                AND SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE
                                AND ARTCUST1.TRADE_CLASS_CODE <> 'HDQ'
                                AND SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE
                                AND ICTWHSE1.LP_CODE IS NOT NULL
                                AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO
                                AND SOTPICK1.PICK_STATUS = 'F'
                                AND SOTPICK1.INV_NO = SOTINVH1.INV_NO
                                AND WHTAUDT1.AUDIT_NO = :PARM1"
            Create_TDA(.Tables.Add, "AuditInvoices", ASCMAIN1.sql, 0, False, "V", 0)

        End With

        Create_Relation("WHTAUDT1", "WHTAUDT2", "AUDIT_NO", "AUDIT_NO")
        'Create_Relation("AuditData", "AuditDetails", "PRO_NO,BILL_OF_LADING_NO,CUST_CODE", "PRO_NO,BILL_OF_LADING_NO,CUST_CODE")

        With dst.Tables("WHTAUDT2")
            .Columns.Add("REVIEWS_1", GetType(Int32), "IIF(ISNULL(REVIEWER_1, '') = '', 0, 1)")
            .Columns.Add("REVIEWS_2", GetType(Int32), "IIF(ISNULL(REVIEWER_2, '') = '', 0, 1)")
        End With

        With dst.Tables("WHTAUDT1")
            '.Columns.Add("NUM_SHIPMENTS", GetType(Int32), "COUNT(CHILD(WHTAUDT1_WHTAUDT2).AUDIT_NO)")
            .Columns.Add("REVIEWS_1", GetType(Int32), "SUM(CHILD(WHTAUDT1_WHTAUDT2).REVIEWS_1)")
            .Columns.Add("REVIEWS_2", GetType(Int32), "SUM(CHILD(WHTAUDT1_WHTAUDT2).REVIEWS_2)")
        End With

        With dst.Tables("WHTAUDT3")
            .Columns.Add("ATTACHMENT", GetType(String))
        End With

        grdWHTAUDT1.DataSource = dst.Tables("WHTAUDT1")
        grdWHTAUDT3.DataSource = dst.Tables("WHTAUDT3")
        grdAuditData.DataSource = dst.Tables("AuditData")
        grdAuditDetails.DataSource = dst.Tables("AuditDetails")
        grdAuditInvoices.DataSource = dst.Tables("AuditInvoices")

        grdWHTAUDT1.Dock = DockStyle.Fill

        Create_Summary(grdWHTAUDT1, "AUDIT_NO", "Count")
        Create_Summary(grdWHTAUDT1, "NUM_SHIPMENTS", "Sum")
        Create_Summary(grdWHTAUDT1, "TOTAL_SHIPMENTS", "Sum")

        Create_Summary(grdWHTAUDT1, "WHSE_CODE", "Count", grdWHTAUDT1.DisplayLayout.Bands(1).Key)
        Create_Summary(grdWHTAUDT1, "CARTONS", "Sum", grdWHTAUDT1.DisplayLayout.Bands(1).Key)

        Create_Summary(grdAuditData, "PRO_NO", "Count")
        Create_Summary(grdAuditData, "NUM_POS", "Sum")
        Create_Summary(grdAuditDetails, "PRO_NO", "Count")

        Create_Summary(grdAuditInvoices, "INV_NO", "Count")
        Create_Summary(grdAuditInvoices, "INV_SALES", "Sum")
        Create_Summary(grdAuditInvoices, "INV_FREIGHT", "Sum")
        Create_Summary(grdAuditInvoices, "INV_MISC_CHG", "Sum")
        Create_Summary(grdAuditInvoices, "INV_STAX", "Sum")
        Create_Summary(grdAuditInvoices, "INV_TOTAL_AMOUNT", "Sum")

        ASCMAIN1.Add_Value_List(grdWHTAUDT1, "SHIP_VIA_CODE", Nothing, Nothing, 1)
        ASCMAIN1.Add_Value_List(grdAuditDetails, "SHIP_VIA_CODE")

        grdWHTAUDT1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdWHTAUDT1.DisplayLayout.Override.CellClickAction = CellClickAction.CellSelect
        For Each grdBand As Infragistics.Win.UltraWinGrid.UltraGridBand In grdWHTAUDT1.DisplayLayout.Bands
            For Each datacol As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdBand.Columns
                Select Case datacol.Key
                    Case "AUDIT_NOTES", "AUDIT_NOTES_2"
                        datacol.CellActivation = Activation.AllowEdit
                    Case Else
                        datacol.CellActivation = Activation.NoEdit
                End Select
            Next
        Next

        Dim SHIP_DATE_FROM As String = ASCDATA1.GetDataValue("SELECT MIN(SHIP_DATE_FROM) FROM WHTAUDT1") & String.Empty
        If SHIP_DATE_FROM.Length = 0 Then
            SHIP_DATE_FROM = "01/01/2024"
        End If

        dteSHIP_DATE_FROM.MinDate = CDate(SHIP_DATE_FROM).ToShortDateString
        dteSHIP_DATE_FROM.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)

        dteSHIP_DATE_TO.MinDate = dteSHIP_DATE_FROM.MinDate
        dteSHIP_DATE_TO.MaxDate = dteSHIP_DATE_FROM.MaxDate

        dteSHIP_DATE_TO.DateTime = DateTime.Now.ToLongDateString
        Select Case DateTime.Now.ToString("MMM").ToUpper
            Case "JAN", "FEB", "MAR"
                If dteSHIP_DATE_FROM.MinDate > CDate("01/01/" & DateTime.Now.ToString("yyyy")) Then
                    dteSHIP_DATE_FROM.DateTime = dteSHIP_DATE_FROM.MinDate
                Else
                    dteSHIP_DATE_FROM.DateTime = CDate("01/01/" & DateTime.Now.ToString("yyyy"))
                End If

            Case "APR", "MAY", "JUN"
                dteSHIP_DATE_FROM.DateTime = CDate("04/01/" & DateTime.Now.ToString("yyyy"))

            Case "JUL", "AUG", "SEP"
                dteSHIP_DATE_FROM.DateTime = CDate("07/01/" & DateTime.Now.ToString("yyyy"))

            Case "OCT", "NOV", "DEC"
                dteSHIP_DATE_FROM.DateTime = CDate("10/01/" & DateTime.Now.ToString("yyyy"))

            Case Else
                dteSHIP_DATE_FROM.DateTime = dteSHIP_DATE_TO.DateTime
        End Select

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

            Case "Done"

            Case "Import Emails"
                If InquiryMode Then
                    Exit Sub
                End If

                If MessageBox.Show("Do you want to Import Emails?", "Import Emails", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
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

            Case "Import Emails"
                ImportEmails()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Import Emails").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode

                If InquiryMode Then
                    .Groups("Screen Control").Items("Import Emails").Visible = False
                End If

                ' Filter Container
                UltraExplorerBarContainerControl1.Enabled = Not tf

            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdWHTAUDT1.Visible = True

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each tableName As String In New String() {"WHTAUDT1", "WHTAUDT2", "WHTAUDT3", "AuditData", "AuditDetails", "AuditInvoices"}
            dst.Tables(tableName).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdWHTAUDT3.Tag = String.Empty
        grdWHTAUDT1.Text = String.Empty
        grdAuditDetails.Text = "Shipping Audit Details broken out by PO"
        grdAuditInvoices.Text = "Shipping Audit Invoices"

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        EnforceConstraints(False)

        Dim SHIP_DATE_FROM As String = dteSHIP_DATE_FROM.DateTime.ToString("dd-MMM-yyyy")
        Dim SHIP_DATE_TO As String = dteSHIP_DATE_TO.DateTime.ToString("dd-MMM-yyyy")

        Dim grdCaption As String = "Shipping Audits"

        Select Case optReviews.Value
            Case "A"
                ASCMAIN1.sql = $"SELECT * FROM WHTAUDT1 WHERE TRUNC(SHIP_DATE_FROM) >= '{SHIP_DATE_FROM}' AND TRUNC(SHIP_DATE_TO) <= '{SHIP_DATE_TO}'"
                grdCaption = $"All Audits between {SHIP_DATE_FROM} and {SHIP_DATE_TO}"

            Case "1"
                ASCMAIN1.sql = "SELECT * FROM WHTAUDT1 WHERE AUDIT_NO IN (SELECT DISTINCT AUDIT_NO FROM WHTAUDT2 WHERE REVIEWER_1 IS NULL)"
                grdCaption = $"All Audits requiring a 1st Review"

            Case "2"
                ASCMAIN1.sql = $"SELECT * FROM WHTAUDT1 
                                WHERE AUDIT_NO IN (SELECT DISTINCT AUDIT_NO FROM WHTAUDT2 WHERE REVIEWER_2 IS NULL AND REVIEWER_1 IS NOT NULL)
                                AND TRUNC(SHIP_DATE_FROM) >= '{SHIP_DATE_FROM}' AND TRUNC(SHIP_DATE_TO) <= '{SHIP_DATE_TO}'"
                grdCaption = $"All Audits between {SHIP_DATE_FROM} and {SHIP_DATE_TO} requiring a 2nd Review"
        End Select

        ASCMAIN1.Progress("-", "WHTAUDT1")
        Fill_Records("WHTAUDT1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.Progress("-", "WHTAUDT2")
        ASCDATA1.ExecuteSQL($"TRUNCATE TABLE {wkTable}")
        ASCDATA1.ExecuteSQL($"INSERT INTO {wkTable} SELECT AUDIT_NO FROM ({ASCMAIN1.sql})")
        Fill_Records("WHTAUDT2")

        ASCMAIN1.Progress("-", "WHTAUDT3")
        Fill_Records("WHTAUDT3")

        For Each rowWHTAUDT3 As DataRow In dst.Tables("WHTAUDT3").Select("")
            rowWHTAUDT3.Item("ATTACHMENT") = Split(rowWHTAUDT3.Item("FILENAME") & "", "\")(Split(rowWHTAUDT3.Item("FILENAME") & "", "\").Length - 1)
        Next

        Sort_grdColumns(grdWHTAUDT1, "SHIP_DATE_FROM", False)
        Sort_grdColumns(grdWHTAUDT1, "LP_CODE,WHSE_CODE,SHIP_REF", False, 1)
        grdWHTAUDT1.Text = grdCaption

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()

        Try
            BeginTrans()
            Update_Record_TDA("WHTAUDT1")
            CommitTrans("Update Complete")
        Catch ex As Exception

            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        InquiryMode = MENU_ITEM_OBJECT = "WHFAUDII"

        If Not InquiryMode Then
            Load_Popup_Menu(grdWHTAUDT1, "SSPBBPBB", "Show Filter", "Show GroupBox", "View Shipment Data", "Auto Size Columns", "1st Review", "2nd Review")
        Else
            Load_Popup_Menu(grdWHTAUDT1, "SSPBB", "Show Filter", "Show GroupBox", "View Shipment Data", "Auto Size Columns")
        End If

        Load_Popup_Menu(grdAuditData, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
        Load_Popup_Menu(grdAuditDetails, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
        Load_Popup_Menu(grdAuditInvoices, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        MyBase.tlb_BeforeToolDropdown(sender, e)

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        'Dim tlb_btn As UltraWinToolbars.ButtonTool
        Dim tlb_btn1 As UltraWinToolbars.ButtonTool
        Dim tlb_btn2 As UltraWinToolbars.ButtonTool

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name
            Case grdWHTAUDT1.Name
                If InquiryMode Then
                    Exit Sub
                End If

                tlb_btn1 = DirectCast(tlb_pop.Tools("1st Review"), UltraWinToolbars.ButtonTool)
                tlb_btn2 = DirectCast(tlb_pop.Tools("2nd Review"), UltraWinToolbars.ButtonTool)

                Select Case grd.ActiveRow.Band.Key
                    Case grd.DisplayLayout.Bands(0).Key
                        tlb_btn1.SharedProps.Enabled = False
                        tlb_btn2.SharedProps.Enabled = False

                    Case grd.DisplayLayout.Bands(1).Key
                        ' Allow the user to Review and unreview the record

                        ' Need to refresh data
                        Dim AUDIT_NO As String = grd.ActiveRow.Cells("AUDIT_NO").Value & String.Empty
                        Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                        Dim rowWHTAUDT2 As DataRow = LookUp("WHTAUDT2", {AUDIT_NO, SHIP_BOL_NO})
                        If rowWHTAUDT2 IsNot Nothing Then
                            grd.ActiveRow.Cells("REVIEWER_1").Value = rowWHTAUDT2.Item("REVIEWER_1") & String.Empty

                            If IsDate(rowWHTAUDT2.Item("REVIEWER_1_DATE") & String.Empty) Then
                                grd.ActiveRow.Cells("REVIEWER_1_DATE").Value = rowWHTAUDT2.Item("REVIEWER_1_DATE") & String.Empty
                            Else
                                grd.ActiveRow.Cells("REVIEWER_1_DATE").Value = DBNull.Value
                            End If

                            grd.ActiveRow.Cells("REVIEWER_2").Value = rowWHTAUDT2.Item("REVIEWER_2") & String.Empty
                            If IsDate(rowWHTAUDT2.Item("REVIEWER_2_DATE") & String.Empty) Then
                                grd.ActiveRow.Cells("REVIEWER_2_DATE").Value = rowWHTAUDT2.Item("REVIEWER_2_DATE") & String.Empty
                            Else
                                grd.ActiveRow.Cells("REVIEWER_2_DATE").Value = DBNull.Value
                            End If
                        End If

                        If grd.ActiveRow.Cells("REVIEWER_2").Value & String.Empty <> String.Empty Then
                            tlb_btn1.SharedProps.Enabled = False
                            tlb_btn2.SharedProps.Enabled = True
                        ElseIf grd.ActiveRow.Cells("REVIEWER_1").Value & String.Empty <> String.Empty Then
                            tlb_btn1.SharedProps.Enabled = True
                            tlb_btn2.SharedProps.Enabled = True
                        Else
                            tlb_btn1.SharedProps.Enabled = True
                            tlb_btn2.SharedProps.Enabled = False
                        End If
                End Select

            Case Else
        End Select
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

            Case "Auto Size Columns"
                Me.Cursor = Cursors.WaitCursor
                grd.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                Me.Cursor = Cursors.Default

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "View Shipment Data"

                Try
                    Me.Cursor = Cursors.WaitCursor

                    ASCMAIN1.Progress("Loading Audit Data", "")
                    Dim AUDIT_NO As String = grd.ActiveRow.Cells("AUDIT_NO").Value & String.Empty
                    Fill_Records("AuditData", {AUDIT_NO})
                    grdAuditData.Text = $"Shipping Audit Data for Audit {AUDIT_NO}"

                    ASCMAIN1.Progress("Loading Audit Details", "")
                    Fill_Records("AuditDetails", {AUDIT_NO})
                    grdAuditDetails.Text = $"Shipping Audit Details broken out by PO for Audit {AUDIT_NO}"

                    ASCMAIN1.Progress("Loading Audit Invoices", "")
                    Fill_Records("AuditInvoices", {AUDIT_NO})
                    grdAuditInvoices.Text = $"Shipping Audit Invoices for Audit {AUDIT_NO}"

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    dst.Tables("AuditData").Rows.Clear()
                    dst.Tables("AuditDetails").Rows.Clear()
                    dst.Tables("AuditInvoices").Rows.Clear()
                Finally
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("", "")
                End Try

            Case "1st Review"
                If InquiryMode Then Exit Sub

                Dim AUDIT_NO As String = grd.ActiveRow.Cells("AUDIT_NO").Value & String.Empty
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                If Not ASCMAIN1.Logical_Lock("WHTAUDT1", AUDIT_NO) Then
                    Exit Sub
                End If

                Dim rowWHTAUDT2 As DataRow = LookUp("WHTAUDT2", {AUDIT_NO, SHIP_BOL_NO})
                If rowWHTAUDT2 Is Nothing Then
                    MessageBox.Show("Could Not Locate the selected Audit entry. Try again", "1St Review", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                grd.ActiveRow.Cells("REVIEWER_1").Value = rowWHTAUDT2.Item("REVIEWER_1") & String.Empty
                If IsDate(rowWHTAUDT2.Item("REVIEWER_1_DATE") & String.Empty) Then
                    grd.ActiveRow.Cells("REVIEWER_1_DATE").Value = rowWHTAUDT2.Item("REVIEWER_1_DATE") & String.Empty
                Else
                    grd.ActiveRow.Cells("REVIEWER_1_DATE").Value = DBNull.Value
                End If

                grd.ActiveRow.Cells("REVIEWER_2").Value = rowWHTAUDT2.Item("REVIEWER_2") & String.Empty
                If IsDate(rowWHTAUDT2.Item("REVIEWER_2_DATE") & String.Empty) Then
                    grd.ActiveRow.Cells("REVIEWER_2_DATE").Value = rowWHTAUDT2.Item("REVIEWER_2_DATE") & String.Empty
                Else
                    grd.ActiveRow.Cells("REVIEWER_2_DATE").Value = DBNull.Value
                End If

                If rowWHTAUDT2.Item("REVIEWER_1") & String.Empty <> String.Empty Then

                    If rowWHTAUDT2.Item("REVIEWER_1") & String.Empty <> ASCMAIN1.USER_ID Then
                        MessageBox.Show("You are Not permitted To Remove a review created by another user.", "1St Review", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If MessageBox.Show("The 1St Review was already completed. Do you want To remove the Review?", "1St Review", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    If rowWHTAUDT2.Item("REVIEWER_2") & String.Empty <> String.Empty Then
                        MessageBox.Show("You must delete the 2nd Review before you can delete the 1St Review.", "1St Review", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Try
                        BeginTrans()
                        ASCMAIN1.sql = "UPDATE WHTAUDT2 Set REVIEWER_1 = NULL, REVIEWER_1_DATE = Null WHERE AUDIT_NO = : PARM1 AND SHIP_BOL_NO = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {AUDIT_NO, SHIP_BOL_NO})
                        CommitTrans()
                        Dim row As DataRow = dst.Tables("WHTAUDT2").Select($"AUDIT_NO = '{AUDIT_NO}' and SHIP_BOL_NO = '{SHIP_BOL_NO}'")(0)

                        row.Item("REVIEWER_1") = DBNull.Value
                        row.Item("REVIEWER_1_DATE") = DBNull.Value
                        grdWHTAUDT1.UpdateData()

                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try
                Else

                    If MessageBox.Show("Do you want to do a 1st Review?", "1st Review", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    Try
                        BeginTrans()
                        ASCMAIN1.sql = "UPDATE WHTAUDT2 SET REVIEWER_1 = :PARM1, REVIEWER_1_DATE = SYSDATE WHERE AUDIT_NO = :PARM2 AND SHIP_BOL_NO = :PARM3"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", {ASCMAIN1.USER_ID, AUDIT_NO, SHIP_BOL_NO})
                        CommitTrans()
                        Dim row As DataRow = dst.Tables("WHTAUDT2").Select($"AUDIT_NO = '{AUDIT_NO}' and SHIP_BOL_NO = '{SHIP_BOL_NO}'")(0)

                        row.Item("REVIEWER_1") = ASCMAIN1.USER_ID
                        row.Item("REVIEWER_1_DATE") = DateTime.Now
                        grdWHTAUDT1.UpdateData()

                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try
                End If

                ASCMAIN1.MultiTask_Release()

            Case "2nd Review"
                If InquiryMode Then Exit Sub

                Dim AUDIT_NO As String = grd.ActiveRow.Cells("AUDIT_NO").Value & String.Empty
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                If Not ASCMAIN1.Logical_Lock("WHTAUDT1", AUDIT_NO) Then
                    Exit Sub
                End If

                Dim rowWHTAUDT2 As DataRow = LookUp("WHTAUDT2", {AUDIT_NO, SHIP_BOL_NO})
                If rowWHTAUDT2 Is Nothing Then
                    MessageBox.Show("Could Not Locate the selected Audit entry. Try again", "2st Review", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                grd.ActiveRow.Cells("REVIEWER_1").Value = rowWHTAUDT2.Item("REVIEWER_1") & String.Empty
                If IsDate(rowWHTAUDT2.Item("REVIEWER_1_DATE") & String.Empty) Then
                    grd.ActiveRow.Cells("REVIEWER_1_DATE").Value = rowWHTAUDT2.Item("REVIEWER_1_DATE") & String.Empty
                Else
                    grd.ActiveRow.Cells("REVIEWER_1_DATE").Value = DBNull.Value
                End If

                grd.ActiveRow.Cells("REVIEWER_2").Value = rowWHTAUDT2.Item("REVIEWER_2") & String.Empty
                If IsDate(rowWHTAUDT2.Item("REVIEWER_2_DATE") & String.Empty) Then
                    grd.ActiveRow.Cells("REVIEWER_2_DATE").Value = rowWHTAUDT2.Item("REVIEWER_2_DATE") & String.Empty
                Else
                    grd.ActiveRow.Cells("REVIEWER_2_DATE").Value = DBNull.Value
                End If

                If rowWHTAUDT2.Item("REVIEWER_2") & String.Empty <> String.Empty Then

                    If rowWHTAUDT2.Item("REVIEWER_2") & String.Empty <> ASCMAIN1.USER_ID Then
                        MessageBox.Show("You are not permitted to Remove a review created by another user.", "1st Review", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If MessageBox.Show("The 2nd Review was already completed. Do you want to remove the Review?", "2nd Review", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    Try
                        BeginTrans()
                        ASCMAIN1.sql = "UPDATE WHTAUDT2 SET REVIEWER_2 = NULL, REVIEWER_2_DATE = NULL WHERE AUDIT_NO = :PARM1 AND SHIP_BOL_NO = :PARM2"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {AUDIT_NO, SHIP_BOL_NO})
                        CommitTrans()

                        Dim row As DataRow = dst.Tables("WHTAUDT2").Select($"AUDIT_NO = '{AUDIT_NO}' and SHIP_BOL_NO = '{SHIP_BOL_NO}'")(0)

                        row.Item("REVIEWER_2") = DBNull.Value
                        row.Item("REVIEWER_2_DATE") = DBNull.Value
                        grdWHTAUDT1.UpdateData()

                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try
                Else

                    If MessageBox.Show("Do you want to do a 2nd Review?", "2nd Review", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    If rowWHTAUDT2.Item("REVIEWER_1") & String.Empty = String.Empty Then
                        MessageBox.Show("You must perform the 1st Review before the 2nd Review.", "2nd Review", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Try
                        BeginTrans()
                        ASCMAIN1.sql = "UPDATE WHTAUDT2 SET REVIEWER_2 = :PARM1, REVIEWER_2_DATE = SYSDATE WHERE AUDIT_NO = :PARM2 AND SHIP_BOL_NO = :PARM3"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", {ASCMAIN1.USER_ID, AUDIT_NO, SHIP_BOL_NO})
                        CommitTrans()

                        Dim row As DataRow = dst.Tables("WHTAUDT2").Select($"AUDIT_NO = '{AUDIT_NO}' and SHIP_BOL_NO = '{SHIP_BOL_NO}'")(0)

                        row.Item("REVIEWER_2") = ASCMAIN1.USER_ID
                        row.Item("REVIEWER_2_DATE") = DateTime.Now
                        grdWHTAUDT1.UpdateData()

                    Catch ex As Exception
                        Rollback(ex.Message)
                    End Try
                End If

                ASCMAIN1.MultiTask_Release()

        End Select

    End Sub


#End Region

#Region "Form Controls"

    Private Sub grdWHTAUDT1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdWHTAUDT1.AfterRowUpdate

        Try
            BeginTrans()
            Dim AUDIT_NO As String = e.Row.Cells("AUDIT_NO").Value & String.Empty
            Dim SHIP_BOL_NO As String = e.Row.Cells("SHIP_BOL_NO").Value & String.Empty
            Dim AUDIT_NOTES As String = e.Row.Cells("AUDIT_NOTES").Value & String.Empty
            Dim AUDIT_NOTES_2 As String = e.Row.Cells("AUDIT_NOTES_2").Value & String.Empty

            ASCMAIN1.sql = "UPDATE WHTAUDT2 SET AUDIT_NOTES = :PARM1, AUDIT_NOTES_2 = :PARM2 WHERE AUDIT_NO = :PARM3 AND SHIP_BOL_NO = :PARM4"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", {AUDIT_NOTES, AUDIT_NOTES_2, AUDIT_NO, SHIP_BOL_NO})
            CommitTrans()

        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

    Private Sub grdWHTAUDT1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdWHTAUDT1.AfterRowActivate

        Dim row As UltraGridRow = grdWHTAUDT1.ActiveRow
        If row Is Nothing Then Exit Sub
        If row.IsFilterRow OrElse row.IsGroupByRow OrElse row.IsAddRow Then Exit Sub

        If Not row.Cells.Exists("AUDIT_NO") Then Exit Sub

        Dim AUDIT_NO As String = row.Cells("AUDIT_NO").Value & String.Empty
        If AUDIT_NO = String.Empty Then Exit Sub

        If grdWHTAUDT3.Tag & String.Empty = AUDIT_NO Then Exit Sub

        Dim dvw As DataView = DirectCast(grdWHTAUDT3.DataSource, DataTable).DefaultView
        dvw.RowFilter = $"AUDIT_NO = '{AUDIT_NO.Replace("'", "''")}'"
        dvw.Sort = "FILENAME"
        grdWHTAUDT3.Tag = AUDIT_NO

    End Sub

    Private Sub grdWHTAUDT1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTAUDT1.InitializeRow

        If e.Row.Band.Key <> grdWHTAUDT1.DisplayLayout.Bands(1).Key Then
            Exit Sub
        End If

        If e.Row.Cells("INV_DATE").Text <> e.Row.Cells("SHIP_DATE").Text Then
            e.Row.Cells("INV_DATE").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub grdWHTAUDT3_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdWHTAUDT3.ClickCellButton

        Select Case e.Cell.Column.Key
            Case "FILENAME", "ATTACHMENT"
                Try
                    'Dim AUDIT_NO As String = e.Cell.Row.Cells("AUDIT_NO").Value & String.Empty
                    Dim FILENAME As String = e.Cell.Row.Cells("FILENAME").Value & String.Empty
                    Dim p As Process = Process.Start(FILENAME)

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "View file", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

        End Select

    End Sub


#End Region

#Region "Form Procedures"

    Private Async Sub ImportEmails()

        Dim lstErrors As New List(Of String)
        Try
            If Not ASCMAIN1.Logical_Lock("WHTAUDT3", "*",,,, 888) Then
                Exit Sub
            End If

            Dim numberDocsImported As Int16 = 0

            Dim WH_PARM_AUDIT_EMAIL_USER_ID As String = (ROWs("WHTPARM1").Item("WH_PARM_AUDIT_EMAIL_USER_ID") & String.Empty).ToString.ToLower
            Dim WH_PARM_AUDIT_EMAIL_PASSWORD As String = ROWs("WHTPARM1").Item("WH_PARM_AUDIT_EMAIL_PASSWORD") & String.Empty
            Dim EMAIL_PDF_DIRECTORY As String = ROWs("WHTPARM1").Item("EMAIL_PDF_DIRECTORY") & String.Empty

            Dim APPID As String = ROWs("WHTPARM1").Item("APPID") & String.Empty
            Dim TENANTID As String = ROWs("WHTPARM1").Item("TENANTID") & String.Empty
            Dim CLIENTSECRET As String = ROWs("WHTPARM1").Item("CLIENTSECRET") & String.Empty

            If WH_PARM_AUDIT_EMAIL_USER_ID.Length = 0 Then
                MessageBox.Show("Email User ID does not have a value", "Import Emails", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If EMAIL_PDF_DIRECTORY.Length = 0 Then
                MessageBox.Show("Email PDF Directory does not have a value", "Import Emails", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            If Not My.Computer.FileSystem.DirectoryExists(EMAIL_PDF_DIRECTORY) Then
                My.Computer.FileSystem.CreateDirectory(EMAIL_PDF_DIRECTORY)
            End If

            Dim service As ExchangeService = Nothing
            AltEwsServiceCredentials = New Credentials

            If APPID.Length > 0 AndAlso TENANTID.Length > 0 AndAlso CLIENTSECRET.Length > 0 Then
                With AltEwsServiceCredentials
                    .APPID = APPID
                    .CLIENTSECRET = CLIENTSECRET
                    .TENANTID = TENANTID
                    .USER_EMAIL = WH_PARM_AUDIT_EMAIL_USER_ID
                End With
            Else
                AltEwsServiceCredentials = Nothing
            End If

            If AltEwsServiceCredentials Is Nothing Then
                service = TAC.TACMAIN1.Get_EWS_Service(WH_PARM_AUDIT_EMAIL_USER_ID)
            Else
                service = Await GetAltEwsService()
            End If

            Dim eMailItems As FindItemsResults(Of Item) = service.FindItems(New FolderId(WellKnownFolderName.Inbox), New ItemView(15))
            If eMailItems.Count > 0 Then
                ASCMAIN1.Progress("Reading Emails", "")
                service.LoadPropertiesForItems(eMailItems, PropertySet.FirstClassProperties)
                Dim customPropertySet = New PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Subject, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.From)

                For Each msg As EmailMessage In eMailItems
                    If Not msg.HasAttachments Then
                        'Continue For
                    End If

                    If Not msg.Subject.ToUpper.Contains("Re: Audit".ToUpper) Then
                        Continue For
                    End If

                    Dim AuditNo As String = msg.Subject.ToUpper.Replace("Re: Audit".ToUpper, "").Trim
                    If AuditNo.Length < 6 Then
                        Continue For
                    End If

                    Dim EMAIL_FROM As String = msg.From.Name
                    Dim EMAIL_DATE As Date = msg.DateTimeSent
                    ASCMAIN1.Progress("Reading Emails " & msg.From.Name & " " & EMAIL_DATE, "")

                    AuditNo = AuditNo.Substring(0, 6)
                    EMAIL_PDF_DIRECTORY = ROWs("WHTPARM1").Item("EMAIL_PDF_DIRECTORY") & String.Empty

                    If Not My.Computer.FileSystem.DirectoryExists(IO.Path.Combine(EMAIL_PDF_DIRECTORY, AuditNo)) Then
                        My.Computer.FileSystem.CreateDirectory(IO.Path.Combine(EMAIL_PDF_DIRECTORY, AuditNo))
                        System.Threading.Thread.Sleep(3000)
                    End If

                    Dim AUDIT_LNO As String = String.Empty
                    For Each attachment As Microsoft.Exchange.WebServices.Data.Attachment In msg.Attachments
                        Dim FileAttachment As FileAttachment = Nothing

                        Try
                            FileAttachment = DirectCast(attachment, FileAttachment)
                        Catch ex As Exception
                            lstErrors.Add($"From: {EMAIL_FROM}, Audit: {AuditNo}, Date: {EMAIL_DATE}")
                            Continue For
                        End Try

                        AUDIT_LNO = Val(ASCMAIN1.Next_Control_No("WHTAUDT3.AUDIT_LNO")).ToString
                        Dim attachmentName As String = IO.Path.Combine(EMAIL_PDF_DIRECTORY, AuditNo, AUDIT_LNO & "_" & FileAttachment.Name)
                        If Not attachmentName.Contains(".") Then
                            attachmentName &= ".pdf"
                        End If
                        FileAttachment.Load(attachmentName)

                        ASCMAIN1.sql = "INSERT INTO WHTAUDT3 (AUDIT_NO, AUDIT_LNO, FILENAME, EMAIL_FROM, EMAIL_DATE) VALUES (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5)"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVD", {AuditNo, AUDIT_LNO, IO.Path.Combine(EMAIL_PDF_DIRECTORY, AuditNo, attachmentName), EMAIL_FROM, EMAIL_DATE})
                        numberDocsImported += 1
                    Next

                    AUDIT_LNO = Val(ASCMAIN1.Next_Control_No("WHTAUDT3.AUDIT_LNO") & String.Empty)
                    Dim emailName As String = "Email_" & AUDIT_LNO & ".eml"
                    msg.SaveToFile(IO.Path.Combine(EMAIL_PDF_DIRECTORY, AuditNo, emailName))

                    ASCMAIN1.sql = "INSERT INTO WHTAUDT3 (AUDIT_NO, AUDIT_LNO, FILENAME, EMAIL_FROM, EMAIL_DATE) VALUES (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5)"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVD", {AuditNo, AUDIT_LNO, IO.Path.Combine(EMAIL_PDF_DIRECTORY, AuditNo, emailName), EMAIL_FROM, EMAIL_DATE})
                    numberDocsImported += 1

                    msg.Move(New FolderId(WellKnownFolderName.DeletedItems))
                Next
            End If

            MessageBox.Show($"{numberDocsImported} Documents Imported.", "Import Emails", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "Import Emails", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.MultiTask_Release(,, 888)
            ASCMAIN1.Progress("", "")
            If lstErrors.Count > 0 Then
                Dim msg As String = "The following emails had errors on the import:" & Environment.NewLine & Environment.NewLine & String.Join(Environment.NewLine, lstErrors.ToArray)
                MessageBox.Show(msg, "Import", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        End Try
    End Sub

    Public Shared Function Get_EWS_Service(ByVal username As String,
                                           ByVal password As String,
                                           ByVal domain As String) As ExchangeService

        Dim service As ExchangeService = New ExchangeService()

        Try
            service.Url = New Uri("https://outlook.office365.com/EWS/Exchange.asmx")
            service.Credentials = New WebCredentials(username, password, domain)
            service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.SmtpAddress, username)
            service.HttpHeaders.Add("X-AnchorMailbox", username)

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then Stop
        End Try

        Return service

    End Function

    Private Sub optReviews_ValueChanged(sender As Object, e As EventArgs) Handles optReviews.ValueChanged
        grpDATE_RANGE.Enabled = optReviews.Value <> "1"
    End Sub

#End Region

#Region "EWS Alt Service"

    Private Class Credentials
        Public APPID As String
        Public TENANTID As String
        Public CLIENTSECRET As String
        Public USER_EMAIL As String
        Public authResult As AuthenticationResult
    End Class

    Private AltEwsServiceCredentials As New Credentials

    Private Async Function GetAltEwsService() As Threading.Tasks.Task(Of ExchangeService)

        Dim value As String = Await Get_Auth()
        Dim service As ExchangeService = New ExchangeService

        Try
            service = Get_EWS_Service(AltEwsServiceCredentials.USER_EMAIL)
        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then Stop
        End Try

        Return service

    End Function

    Async Function Get_Auth() As Threading.Tasks.Task(Of String)

        Dim authResult As AuthenticationResult = Nothing

        Try
            Dim appID As String = AltEwsServiceCredentials.APPID
            Dim tenantID As String = AltEwsServiceCredentials.TENANTID
            Dim clientSecret As String = AltEwsServiceCredentials.CLIENTSECRET

            Dim cca As IConfidentialClientApplication = ConfidentialClientApplicationBuilder _
            .Create(appID) _
                .WithClientSecret(clientSecret) _
                .WithTenantId(tenantID) _
                .Build()
            Dim ewsScopes As String() = New String() {"https://outlook.office365.com/.default"}
            authResult = Await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync()
            AltEwsServiceCredentials.authResult = authResult

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then Stop
        End Try

        If authResult IsNot Nothing Then
            Return authResult.ToString
        Else
            Return ""
        End If
    End Function

    Private Function Get_EWS_Service(USER_EMAIL As String) As ExchangeService

        Dim service As ExchangeService = New ExchangeService()

        Try
            Dim authResult As AuthenticationResult = AltEwsServiceCredentials.authResult

            service.Url = New Uri("https://outlook.office365.com/EWS/Exchange.asmx")
            service.Credentials = New OAuthCredentials(authResult.AccessToken)
            service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.SmtpAddress, USER_EMAIL)
            service.HttpHeaders.Add("X-AnchorMailbox", USER_EMAIL)

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then Stop
        End Try

        Return service

    End Function

#End Region

End Class