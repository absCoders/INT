Imports Infragistics.Win.UltraWinGrid
Imports System.Data
'Imports System.Data.DataSetExtensions   ' Needed for AsEnumerable() and Field(Of T)
Imports System.Linq

' ISSUE-6889: 3PL to 3PL transfers

Public Class ICFXFR3P

#Region "Form Variables"

    Private Const XFR_SOURCE As String = "I"
    Private Const transferWarehouse As String = "XIT"
    Private lstTransWhses As New List(Of String) From {"XIT"}
    Private sqlShipRecon As String = String.Empty

    ' ISSUE-7288
    Private ViewingReconciliation As Boolean = False

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.Running_in_VS Then
            ValidateTablesExist()
        End If

        With dst
            ASCMAIN1.sql = "SELECT EDTTRXN5.ITEM_CODE, ICTITEM1.ITEM_DESC, SUM(EDTTRXN5.TRAN_QTY) TRAN_QTY 
                                FROM EDTTRXN5, ICTITEM1 
                                WHERE EDTTRXN5.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                                AND EDTTRXN5.TRANS_DATE = :PARM1 
                                GROUP BY EDTTRXN5.ITEM_CODE, ICTITEM1.ITEM_DESC"
            Create_TDA(.Tables.Add, "EDTTRXN5_LK", ASCMAIN1.sql, 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT ICTXFRX2.ITEM_CODE, ICTITEM1.ITEM_DESC, 0 TRAN_QTY_5, 0 USED_QTY
                                FROM ICTXFRX1, ICTXFRX2, ICTITEM1
                                WHERE ICTXFRX1.CTL_NO = ICTXFRX2.CTL_NO
                                AND ICTXFRX1.BOL_NO = :PARM1 AND ICTXFRX1.CTL_NO = :PARM2
                                AND ICTXFRX2.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                                GROUP BY ICTXFRX1.BOL_NO, ICTXFRX2.ITEM_CODE, ICTITEM1.ITEM_DESC"
            Create_TDA(.Tables.Add, "ItemsSelected", ASCMAIN1.sql, 0, False, "VV", 1)
            With .Tables("ItemsSelected")
                Dim bolNoTotal As String = String.Empty
                For iLoop As Int16 = 1 To 30
                    .Columns.Add($"BOL_NO_{iLoop}", GetType(Int64))
                    bolNoTotal &= $" + ISNULL(BOL_NO_{iLoop}, 0)"
                Next
                .Columns.Add("BOL_NO_TOTAL", GetType(Int64), bolNoTotal.Substring(2))
                .Columns.Add("TRAN_DIFF", GetType(Int64), "ISNULL(TRAN_QTY_5, 0) - (ISNULL(BOL_NO_TOTAL, 0) + ISNULL(USED_QTY, 0))")
            End With

            ASCMAIN1.sql = "SELECT ICTXFRX1.BOL_NO, NVL(ICTXFRX1.PO_ORDER_NO, 'Unk') PO_ORDER_NO, ICTXFRX1.PO_TRANS_DATE, ICTXFRX1.WHSE_CODE, ICTXFRX1.WHSE_CODE_TO, SUM(ICTXFRX2.TRAN_QTY) TRAN_QTY
                                FROM ICTXFRX1, ICTXFRX2
                                WHERE ICTXFRX1.CTL_NO = ICTXFRX2.CTL_NO
                                AND ICTXFRX1.TRANS_DATE = :PARM1
                                AND ICTXFRX2.ITEM_CODE = :PARM2
                                GROUP BY ICTXFRX1.BOL_NO, ICTXFRX1.PO_ORDER_NO, ICTXFRX1.PO_TRANS_DATE, ICTXFRX1.WHSE_CODE, ICTXFRX1.WHSE_CODE_TO"
            Create_TDA(.Tables.Add, "ItemsRecon", ASCMAIN1.sql, 0, False, "VV", 2)

            Create_TDA(.Tables.Add, "ICTXFRX0", "*")
            Create_TDA(.Tables.Add, "ICTXFRX1", "*")
            Create_TDA(.Tables.Add, "ICTXFRX2", "*")
            .Tables("ICTXFRX2").Columns.Add("ITEM_DESC", GetType(String))
            Create_Relation("ICTXFRX1", "ICTXFRX2", "CTL_NO", "CTL_NO")

            Create_TDA(.Tables.Add, "ICTIXFR1", "*")
            Create_TDA(.Tables.Add, "ICTIXFR2", "*")
            .Tables("ICTIXFR2").Columns.Add("TOTAL_COSTS", GetType(System.Decimal), "ISNULL(XFR_QTY, 0) * ISNULL(ITEM_COST_STD, 0)")
            Create_TDA(.Tables.Add, "ICTWHSE1", "*")

            ASCMAIN1.sql = "SELECT * FROM ICTXFRX1 WHERE XFR_BATCH_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "ICTXFRX1_0", ASCMAIN1.sql, 0, False, "V")

            ASCMAIN1.sql = $"SELECT
                                X.*,
                                ICTXFRX3.NOTE
                            FROM (
                                SELECT
                                    WHSE_CODE,
                                    WHSE_CODE_TO,
                                    ITEM_CODE,
                                    ITEM_DESC,
                                    SUM(ADJ_QTY) AS ADJ_QTY,
                                    SUM(BOL_QTY) AS BOL_QTY,
                                    SUM(ADJ_QTY) + SUM(BOL_QTY) AS ADJ_DIFF,
                                    SUM(XFR_QTY) AS XFR_QTY,
                                    SUM(RECEIPT_QTY) AS RECEIPT_QTY,
                                    SUM(XFR_QTY) - SUM(RECEIPT_QTY) AS XFR_DIFF
                                FROM (
                                    SELECT
                                        'A' AS REC_TYPE,
                                        E.WHSE_CODE,
                                        E.WHSE_CODE_TO,
                                        E.ITEM_CODE,
                                        I.ITEM_DESC,
                                        SUM(E.TRAN_QTY) AS ADJ_QTY,
                                        0 AS BOL_QTY,
                                        0 AS XFR_QTY,
                                        0 AS RECEIPT_QTY
                                    FROM EDTTRXN5 E, ICTITEM1 I
                                    WHERE E.ITEM_CODE = I.ITEM_CODE (+)
                                    GROUP BY
                                        E.WHSE_CODE,
                                        E.WHSE_CODE_TO,
                                        E.ITEM_CODE,
                                        I.ITEM_DESC

                                    UNION ALL

                                    SELECT
                                        'E',
                                        X1.WHSE_CODE,
                                        DECODE(X1.WHSE_CODE_TO, 'ADSBLA', 'ADS', 'ADSMIN', 'ADS', X1.WHSE_CODE_TO),
                                        X2.ITEM_CODE,
                                        I.ITEM_DESC,
                                        0,
                                        SUM(X2.TRAN_QTY),
                                        0,
                                        0
                                    FROM ICTXFRX1 X1, ICTXFRX2 X2, ICTITEM1 I
                                    WHERE X1.CTL_NO = X2.CTL_NO
                                      AND X2.ITEM_CODE = I.ITEM_CODE (+)
                                    GROUP BY
                                        X1.WHSE_CODE,
                                        DECODE(X1.WHSE_CODE_TO, 'ADSBLA', 'ADS', 'ADSMIN', 'ADS', X1.WHSE_CODE_TO),
                                        X2.ITEM_CODE,
                                        I.ITEM_DESC

                                    UNION ALL

                                    SELECT
                                        'P',
                                        I1.WHSE_CODE,
                                        DECODE(I1.WHSE_CODE_TO, 'XIT', 'ADS', 'XITRTN', 'ADSRTN', I1.WHSE_CODE_TO),
                                        I2.ITEM_CODE,
                                        I.ITEM_DESC,
                                        0,
                                        0,
                                        SUM(I2.XFR_QTY),
                                        0
                                    FROM ICTIXFR1 I1, ICTIXFR2 I2, ICTITEM1 I
                                    WHERE I1.XFR_NO = I2.XFR_NO
                                      AND I1.XFR_SOURCE = :PARM1
                                      AND I1.WHSE_CODE_TO IN (SELECT * FROM TABLE(IN_LIST(:PARM2)))
                                      AND I2.ITEM_CODE = I.ITEM_CODE (+)
                                      AND I1.XFR_REF IS NOT NULL
                                    GROUP BY
                                        I1.WHSE_CODE,
                                        DECODE(I1.WHSE_CODE_TO, 'XIT', 'ADS', 'XITRTN', 'ADSRTN', I1.WHSE_CODE_TO),
                                        I2.ITEM_CODE,
                                        I.ITEM_DESC

                                    UNION ALL

                                    SELECT
                                        'R',
                                        DECODE(I1.WHSE_CODE, 'XIT', 'CLA', 'XITRTN', 'CLARTN', I1.WHSE_CODE),
                                        DECODE(I1.WHSE_CODE_TO, 'ADSBLA', 'ADS', 'ADSMIN', 'ADS', I1.WHSE_CODE_TO),
                                        I2.ITEM_CODE,
                                        I.ITEM_DESC,
                                        0,
                                        0,
                                        0,
                                        SUM(I2.XFR_QTY)
                                    FROM ICTIXFR1 I1, ICTIXFR2 I2, ICTITEM1 I
                                    WHERE I1.XFR_NO = I2.XFR_NO
                                      AND I1.XFR_SOURCE = :PARM1
                                      AND I1.WHSE_CODE IN (SELECT * FROM TABLE(IN_LIST(:PARM2)))
                                      AND I2.ITEM_CODE = I.ITEM_CODE (+)
                                      AND I1.XFR_REF IS NOT NULL
                                    GROUP BY
                                        DECODE(I1.WHSE_CODE, 'XIT', 'CLA', 'XITRTN', 'CLARTN', I1.WHSE_CODE),
                                        DECODE(I1.WHSE_CODE_TO, 'ADSBLA', 'ADS', 'ADSMIN', 'ADS', I1.WHSE_CODE_TO),
                                        I2.ITEM_CODE,
                                        I.ITEM_DESC
                                )
                                GROUP BY
                                    WHSE_CODE,
                                    WHSE_CODE_TO,
                                    ITEM_CODE,
                                    ITEM_DESC
                            ) X,
                            ICTXFRX3
                            WHERE X.WHSE_CODE = ICTXFRX3.WHSE_CODE (+)
                              AND X.WHSE_CODE_TO = ICTXFRX3.WHSE_CODE_TO (+)
                              AND X.ITEM_CODE = ICTXFRX3.ITEM_CODE (+)"
            sqlShipRecon = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SHIP_RECON", ASCMAIN1.sql, 0, False, "VV", 0)

            Create_TDA(.Tables.Add, "ICTPARM1", "*")
            Fill_Records("ICTPARM1", "", True, "SELECT * FROM ICTPARM1 WHERE IC_PARM_KEY = 'Z'")
            ' ISSUE-7288
            Create_TDA(.Tables.Add, "ICTXFRX3", "*")
        End With

        splImport.Dock = DockStyle.Fill
        grdSHIP_RECON.Dock = DockStyle.Fill
        grdSHIP_RECON.Visible = False

        grdSHIP_RECON.DataSource = dst.Tables("SHIP_RECON")
        Create_Summary(grdSHIP_RECON, "ITEM_CODE", "Count")
        For Each columnName As String In {"ADJ_QTY", "BOL_QTY", "ADJ_DIFF", "XFR_QTY", "RECEIPT_QTY", "XFR_DIFF"}
            Create_Summary(grdSHIP_RECON, columnName, "Sum")
        Next

        Create_Relation("ICTXFRX0", "ICTXFRX1_0", "XFR_BATCH_NO", "XFR_BATCH_NO")
        grdICTXFRX0.DataSource = dst.Tables("ICTXFRX0")
        Create_Summary(grdICTXFRX0, "XFR_BATCH_NO", "Count")

        grdItemsRecon.DataSource = dst.Tables("ItemsRecon")
        Create_Summary(grdItemsRecon, "BOL_NO", "Count")
        Create_Summary(grdItemsRecon, "TRAN_QTY", "Sum")
        dst.Tables("ICTXFRX1").Columns.Add("BOL_TRAN_QTY", GetType(Int64), "SUM(CHILD.TRAN_QTY)")

        grdICTXFRX1.DataSource = dst.Tables("ICTXFRX1")
        Create_Summary(grdICTXFRX1, "BOL_NO", "Count")
        Create_Summary(grdICTXFRX1, "BOL_TRAN_QTY", "Sum")

        grdICTXFRX2.DataSource = dst.Tables("ICTXFRX2")
        Create_Summary(grdICTXFRX2, "TRAN_QTY", "Sum")

        grdItemsSelected.DataSource = dst.Tables("ItemsSelected")
        Create_Summary(grdItemsSelected, "ITEM_CODE", "Count")
        Create_Summary(grdItemsSelected, "TRAN_QTY_5", "Sum")
        Create_Summary(grdItemsSelected, "TRAN_DIFF", "Sum")
        Create_Summary(grdItemsSelected, "BOL_NO_TOTAL", "Sum")
        Create_Summary(grdItemsSelected, "USED_QTY", "Sum")

        For iLoop As Int16 = 1 To 30
            Dim columnName As String = $"BOL_NO_{iLoop}"
            Create_Summary(grdItemsSelected, columnName, "Sum")
        Next
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Import Xls"

            Case "Load"

            Case "Update"
                ' ISSUE-7288
                If ViewingReconciliation Then
                    Exit Select
                End If

                dst.Tables("ICTXFRX1").AcceptChanges()
                If dst.Tables("ICTXFRX1").Rows.Count = 0 Then
                    EMsg &= vbCr & "The are no Imported Bols."
                    Exit Select
                End If

                Fill_Records("ICTWHSE1", String.Empty, True, "SELECT * FROM ICTWHSE1")
                For Each drICTXFRX1 As DataRow In dst.Tables("ICTXFRX1").Select("")
                    Dim WHSE_CODE_TO As String = drICTXFRX1.Item("WHSE_CODE_TO") & String.Empty
                    If WHSE_CODE_TO.Length = 0 Then
                        EMsg &= vbCr & "All Imported Bols require a To Warehouse."
                        Exit Select
                    End If

                    Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE_TO)
                    If drICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & $"To Warehouse {WHSE_CODE_TO} is an invalid warehouse"
                        Exit Select
                    End If

                    If drICTWHSE1.Item("WHSE_LOC_XFR") & String.Empty = String.Empty Then
                        EMsg &= vbCr & $"To Warehouse {WHSE_CODE_TO} is not assigned a Transfer Warehouse"
                        Exit Select
                    End If

                    drICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(drICTWHSE1.Item("WHSE_LOC_XFR") & String.Empty)
                    If drICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & $"To Warehouse {WHSE_CODE_TO} is assigned to an invalid Transfer Warehouse {drICTWHSE1.Item("WHSE_LOC_XFR") & String.Empty}"
                        Exit Select
                    End If
                Next

                If MessageBox.Show($"Do you want to {eItemKey}?", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

            Case "Cancel"
                'If MessageBox.Show($"Do you want to {eItemKey} any changes", eItemKey, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                '    Exit Sub
                'End If


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Cancel"
                Mode_Settings(False)

            Case "Import Xls"
                If ImportExcelFile() Then
                    Mode_Settings(True)
                End If

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "View Reconciliation"
                ViewReconciliation()
                Mode_Settings(True, "ViewReconciliation")
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Import Xls").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("View Reconciliation").Settings.Enabled = not_iScreenMode

                If MODE_description = "ViewReconciliation" Then
                    .Groups("Screen Control").Items("Import Xls").Settings.Enabled = DefaultableBoolean.False
                    ' ISSUE-7288
                    .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.True
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                    .Groups("Screen Control").Items("View Reconciliation").Settings.Enabled = DefaultableBoolean.False
                End If
            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"EDTTRXN5_LK",
                "ItemsSelected", "ItemsRecon", "ICTXFRX1_0",
                "ICTXFRX0", "ICTXFRX1", "ICTXFRX2",
                "ICTIXFR1", "ICTIXFR2", "SHIP_RECON", "ICTXFRX3"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Fill_Records("ICTXFRX0", "", True, "SELECT * FROM ICTXFRX0")
        Dim lstXfrBatchNos As New List(Of String)
        For Each drICTXFRX0 As DataRow In dst.Tables("ICTXFRX0").Select("")
            lstXfrBatchNos.Add(drICTXFRX0.Item("XFR_BATCH_NO"))
        Next
        Fill_Records("ICTXFRX1_0", String.Join(",", lstXfrBatchNos.ToArray))
        grdICTXFRX0.DisplayLayout.Bands(1).Hidden = False

        grdICTXFRX0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Sort_grdColumns(grdICTXFRX0, "TRANS_DATE".ToLower)
        EnforceConstraints(True)

        For iLoop As Int16 = 1 To 30
            Dim columnName As String = $"BOL_NO_{iLoop}"
            grdItemsSelected.DisplayLayout.Bands(0).Columns(columnName).Hidden = True
            grdItemsSelected.DisplayLayout.Bands(0).Columns(columnName).Header.VisiblePosition = iLoop + 4
            grdItemsSelected.DisplayLayout.Bands(0).Columns(columnName).Tag = String.Empty
        Next
        grdItemsSelected.DisplayLayout.Bands(0).Columns("BOL_NO_TOTAL").Header.VisiblePosition = grdItemsSelected.DisplayLayout.Bands(0).Columns.Count
        grdItemsSelected.DisplayLayout.Bands(0).Columns("TRAN_DIFF").Header.VisiblePosition = grdItemsSelected.DisplayLayout.Bands(0).Columns.Count
        grdItemsSelected.Text = $"Item Reconciliation"
        grdItemsSelected.DisplayLayout.Bands(0).Columns("BOL_NO_TOTAL").Hidden = True

        splImportDetails.Panel1Collapsed = False
        splImportDetails.Panel2Collapsed = True
        splItemRecon.Panel2Collapsed = False

        splImport.Visible = True
        grdSHIP_RECON.Visible = False

        ' ISSUE-7288
        ViewingReconciliation = False
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        ' ISSUE-7288
        If ViewingReconciliation Then
            Try
                BeginTrans()

                dst.Tables("ICTXFRX3").Rows.Clear()
                For Each dr As DataRow In dst.Tables("SHIP_RECON").Select("")
                    Dim NOTE As String = (dr.Item("NOTE") & String.Empty).ToString.Trim
                    If NOTE.Length = 0 Then
                        Continue For
                    End If

                    Dim WHSE_CODE As String = (dr.Item("WHSE_CODE") & String.Empty).ToString.Trim
                    Dim WHSE_CODE_TO As String = (dr.Item("WHSE_CODE_TO") & String.Empty).ToString.Trim
                    Dim ITEM_CODE As String = (dr.Item("ITEM_CODE") & String.Empty).ToString.Trim
                    dst.Tables("ICTXFRX3").Rows.Add({WHSE_CODE, WHSE_CODE_TO, ITEM_CODE, NOTE})
                Next

                ASCDATA1.ExecuteSQL("DELETE FROM ICTXFRX3")
                Update_Record_TDA("ICTXFRX3")
                CommitTrans()

            Catch ex As Exception
                Rollback(ex.Message)
            End Try
            Exit Sub
        End If

        Dim inTrans As Boolean = False
        Try
            For Each drICTXFRX1 As DataRow In dst.Tables("ICTXFRX1").Select("")
                drICTXFRX1.Item("XFR_NO") = ASCMAIN1.Next_Control_No("ICTIXFR1.XFR_NO")
            Next

            BeginTrans()
            inTrans = True

            dst.Tables("ICTXFRX1").AcceptChanges()
            dst.Tables("ICTXFRX2").AcceptChanges()

            For Each drICTXFRX1 As DataRow In dst.Tables("ICTXFRX1").Select("")
                Dim CTL_NO As String = drICTXFRX1.Item("CTL_NO")
                drICTXFRX1.SetAdded()

                For Each drICTXFRX2 As DataRow In dst.Tables("ICTXFRX2").Select($"CTL_NO = '{CTL_NO}'")
                    drICTXFRX2.SetAdded()
                Next
            Next

            Update_Record_TDA("ICTXFRX0")
            Update_Record_TDA("ICTXFRX1")
            Update_Record_TDA("ICTXFRX2")

            For Each drICTXFRX1 As DataRow In dst.Tables("ICTXFRX1").Select("")
                Dim CTL_NO As String = drICTXFRX1.Item("CTL_NO")
                Dim XFR_NO As String = drICTXFRX1.Item("XFR_NO")
                ASCMAIN1.Progress($"Create Transfer {XFR_NO}", "")
                If Not CreateTransfer(CTL_NO) Then
                    Throw New Exception($"Could Not Create Transfer")
                End If
            Next

            CommitTrans("Transfers Updated. After clicking 'OK' the PO will be sent to the warehouse.")
            inTrans = False

            For Each drICTXFRX1 As DataRow In dst.Tables("ICTXFRX1").Select("")
                Dim XFR_NO As String = drICTXFRX1.Item("XFR_NO")
                Dim BOL_NO As String = drICTXFRX1.Item("BOL_NO")

                ASCMAIN1.Progress($"Create Transfer Purchase Order for Transfer {XFR_NO}", "")
                If Not TransferPurchaseOrder(XFR_NO) Then
                    MessageBox.Show($"Could Not Create Transfer PurchaseOrder for BOL No: {BOL_NO}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    'MessageBox.Show($"PurchaseOrder transfered for BOL No: {BOL_NO}", "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Next

        Catch ex As Exception
            If inTrans Then
                Rollback(ex.Message)
            Else
                MessageBox.Show(ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Finally
            ASCMAIN1.Progress("", "")
        End Try
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTXFRX0, "SSBPB", "Show Filter", "Show GroupBox", "Auto Fit Columns", "Transfer PO")
        Load_Popup_Menu(grdICTXFRX1, "SSB", "Show Filter", "Show GroupBox", "Auto Fit Columns")
        Load_Popup_Menu(grdICTXFRX2, "SSB", "Show Filter", "Show GroupBox", "Auto Fit Columns")

        Load_Popup_Menu(grdItemsSelected, "SSB", "Show Filter", "Show GroupBox", "Auto Fit Columns")
        Load_Popup_Menu(grdItemsRecon, "SSB", "Show Filter", "Show GroupBox", "Auto Fit Columns")

        Load_Popup_Menu(grdSHIP_RECON, "SSB", "Show Filter", "Show GroupBox", "Auto Fit Columns")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case grdICTXFRX0.Name
                    tlb_btn = DirectCast(tlb.Tools("Transfer PO"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key AndAlso grd.ActiveRow.Cells("PO_ORDER_NO").Value & String.Empty = String.Empty

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Auto Fit Columns"
                AutoFitGridColumns(grd)

            Case "Create Transfer"
                Try
                    Dim CTL_NO As String = grd.ActiveRow.Cells("CTL_NO").Text
                    Dim BOL_NO As String = grd.ActiveRow.Cells("BOL_NO").Text
                    Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Text
                    Dim WHSE_CODE_TO As String = grd.ActiveRow.Cells("WHSE_CODE_TO").Text

                    If Not ASCMAIN1.Logical_Lock("ICTXFRX1", CTL_NO,,,, 888) Then
                        Exit Sub
                    End If

                    ASCMAIN1.sql = $"Select * FROM ICTXFRX1 WHERE CTL_NO = '{CTL_NO}' AND XFR_NO Is Null"
                    Dim drICTXFRX1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", {CTL_NO})

                    If drICTXFRX1 Is Nothing Then
                        MessageBox.Show($"BOL No. {BOL_NO} has no items pending transfer.", "Create Transfer", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim zMsg As String = $"Please confirm if you would like to initiate a Transfer PO from {WHSE_CODE} to {WHSE_CODE_TO} using BOL No: {BOL_NO}."
                    zMsg &= $" The items will be stored in the {transferWarehouse} warehouse until a receipt is received from {WHSE_CODE_TO}."

                    If MessageBox.Show(zMsg, "Create Transfer", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    Send3PLPurchaseOrder(CTL_NO)

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Create Transfer", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    ASCMAIN1.MultiTask_Release(,, 888)
                End Try

            Case "Transfer PO"
                Try
                    Dim CTL_NO As String = grd.ActiveRow.Cells("CTL_NO").Text
                    Dim BOL_NO As String = grd.ActiveRow.Cells("BOL_NO").Text
                    Dim XFR_NO As String = grd.ActiveRow.Cells("XFR_NO").Text

                    If Not ASCMAIN1.Logical_Lock("ICTIXFR1", XFR_NO,,,, 888) Then
                        Exit Sub
                    End If

                    If Not ASCMAIN1.Logical_Lock("ICTXFRX1", CTL_NO,,,, 888) Then
                        Exit Sub
                    End If

                    Dim drICTXFRX1 As DataRow = LookUp("ICTXFRX1", CTL_NO)
                    If drICTXFRX1 Is Nothing Then
                        MessageBox.Show($"Cannot locate Transfer BOL No: {BOL_NO}", "Transfer PO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If drICTXFRX1.Item("PO_ORDER_NO") & String.Empty <> String.Empty Then
                        MessageBox.Show($"BOL No: {BOL_NO} was already sent to the warehouse", "Transfer PO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        grd.ActiveRow.Cells("PO_ORDER_NO").Value = drICTXFRX1.Item("PO_ORDER_NO")
                        grd.ActiveRow.Cells("PO_TRANS_DATE").Value = drICTXFRX1.Item("PO_TRANS_DATE")
                        Exit Sub
                    End If

                    If TransferPurchaseOrder(XFR_NO) Then
                        drICTXFRX1 = LookUp("ICTXFRX1", CTL_NO)
                        If drICTXFRX1 IsNot Nothing Then
                            grd.ActiveRow.Cells("PO_ORDER_NO").Value = drICTXFRX1.Item("PO_ORDER_NO")
                            grd.ActiveRow.Cells("PO_TRANS_DATE").Value = drICTXFRX1.Item("PO_TRANS_DATE")
                        End If
                    End If

                Catch ex As Exception
                    MessageBox.Show(ex.Message, "Transfer PO", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Finally
                    ASCMAIN1.MultiTask_Release(,, 888)
                End Try

        End Select
    End Sub

    Private Sub AutoFitGridColumns(grd As UltraWinGrid.UltraGrid)
        For Each band As UltraGridBand In grd.DisplayLayout.Bands
            band.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
        Next
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If ScreenMode Then
            Exit Sub
        End If

        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    e.Handled = True
                    Me.ProcessTabKey(Not e.Shift)
                    Call Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "OPS_YYYYPP"
                Call Click_Command("Load")
        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdICTXFRX1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTXFRX1.AfterRowActivate

        Dim CTL_NO As String = grdICTXFRX1.ActiveRow.Cells("CTL_NO").Value & String.Empty
        Dim BOL_NO As String = grdICTXFRX1.ActiveRow.Cells("BOL_NO").Value & String.Empty
        Dim dv As DataView = DirectCast(grdICTXFRX2.DataSource, DataTable).DefaultView
        dv.RowFilter = $"CTL_NO = '{CTL_NO}'"
        dv.Sort = "ITEM_CODE"

        grdICTXFRX2.Text = $"Items for BOL No: {BOL_NO}"
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub ViewReconciliation()
        splImport.Visible = False
        grdSHIP_RECON.Visible = True

        Fill_Records("SHIP_RECON", {XFR_SOURCE, String.Join(",", lstTransWhses.ToArray)})
        Sort_grdColumns(grdSHIP_RECON, "ITEM_CODE")
        ' ISSUE-7288
        ViewingReconciliation = True
    End Sub

    Private Sub ValidateTablesExist()

        For Each tableName As String In {"ICTXFRX0", "ICTXFRX1", "ICTXFRX2"}

            ASCMAIN1.sql = "Select * from USER_TABLES WHERE TABLE_NAME = :PARM1"
            Dim rowTable As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", {tableName})
            If rowTable IsNot Nothing Then
                Continue For
            End If

            Select Case tableName
                Case "ICTXFRX0"
                    ASCMAIN1.sql = "CREATE TABLE ICTXFRX0(
                                XFR_BATCH_NO VARCHAR2(10),
                                TRANS_DATE DATE,
                                INIT_OPER VARCHAR2(20),
                                INIT_DATE DATE,
                                PRIMARY KEY (XFR_BATCH_NO))"

                Case "ICTXFRX1"
                    ASCMAIN1.sql = "CREATE TABLE ICTXFRX1(
                                CTL_NO VARCHAR2(10), 
                                BOL_NO VARCHAR2(20), 
                                TRANS_DATE DATE, 
                                FILENAME VARCHAR2(100), 
                                XFR_NO VARCHAR2(10), 
                                PO_ORDER_NO VARCHAR2(6), 
                                PO_TRANS_DATE DATE, 
                                WHSE_CODE VARCHAR2(6), 
                                WHSE_CODE_TO VARCHAR2(6), 
                                INIT_OPER VARCHAR2(20), 
                                INIT_DATE DATE, 
                                XFR_BATCH_NO VARCHAR2(10),
                                PRIMARY KEY (CTL_NO))"

                Case "ICTXFRX2"
                    ASCMAIN1.sql = "CREATE TABLE ICTXFRX2(
                            CTL_NO VARCHAR2(10), 
                            PALLET_NO NUMBER(4,0), 
                            ITEM_CODE VARCHAR2(25), 
                            TRAN_QTY NUMBER(7,0), 
                            TRAN_NOTE VARCHAR2(100))"

            End Select

            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            Select Case tableName
                Case "ICTXFRX2"
                    ASCDATA1.ExecuteSQL("CREATE INDEX I_ICTXFRX2_1 ON ICTXFRX2(CTL_NO)")
            End Select

        Next

        Try
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub VerifyTables()

        For Each tableName As String In {"ICTXFRX0", "ICTXFRX1", "ICTXFRX2"}

        Next
    End Sub

    Private Function GetDateFromWorksheetCell(worksheet As SpreadsheetGear.IWorksheet, cellAddress As String) As Date
        Dim cellValue As String = worksheet?.Cells(cellAddress)?.Value
        If cellValue IsNot Nothing Then
            If IsDate(worksheet?.Workbook?.NumberToDateTime(cellValue).Date) Then
                Return worksheet?.Workbook?.NumberToDateTime(cellValue).Date
            Else
                Return DateTime.Now
            End If
        Else
            Return DateTime.Now
        End If
    End Function

    Private Function ImportExcelFile() As Boolean

        ' Stephanie, I wrote some of the code for you. You need to open and read the Excel Workbook.
        ' Verify the Workbook will have only 1 sheet with BOL data.
        ' Look at SOFALLO2.Load_from_XLS to see code that readsd from a Excel Workbook

        ' Need to verify the date in the worksheet matches the selected date - Below
        ' Need to try and lock the date, warehouse
        Dim xfrDate As Date = Nothing
        Dim xfrWhseCode As String = dst.Tables("ICTPARM1").Rows(0).Item("IC_PARM_WHSE_CODE") & String.Empty
        Dim DATE_KEY As String = String.Empty
        Dim excelFilename As String = String.Empty

        If Not ASCMAIN1.Logical_Lock("ICTXFRX1", "IMPORT") Then
            Return False
        End If

        dst.Tables("ItemsSelected").Rows.Clear()

        Dim firstWorkBook As Boolean = True

        Try
            Using openFileDialog1 As New OpenFileDialog
                openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
                openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
                openFileDialog1.RestoreDirectory = True
                openFileDialog1.Multiselect = True

                If openFileDialog1.ShowDialog() <> DialogResult.OK Then
                    Return dst.Tables("ICTXFRX1").Rows.Count <> 0
                End If

                Dim INIT_DATE As Date = DateTime.Now
                Dim XFR_BATCH_NO As String = ASCMAIN1.Next_Control_No("ICTXFRX0.XFR_BATCH_NO")
                Dim rowICTXFRX0 As DataRow = dst.Tables("ICTXFRX0").NewRow
                rowICTXFRX0.Item("XFR_BATCH_NO") = XFR_BATCH_NO
                rowICTXFRX0.Item("INIT_DATE") = INIT_DATE
                rowICTXFRX0.Item("INIT_OPER") = ASCMAIN1.USER_ID
                dst.Tables("ICTXFRX0").Rows.Add(rowICTXFRX0)

                For Each excelFilename In openFileDialog1.FileNames

                    ASCMAIN1.Progress(excelFilename, "")

                    If Not My.Computer.FileSystem.FileExists(excelFilename) Then
                        MessageBox.Show($"Cannot locate file: {excelFilename}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Continue For
                    End If

                    Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(excelFilename)

                    If Not IsDate(GetDateFromWorksheetCell(workbook.Worksheets(0), "L3") & String.Empty) Then
                        MessageBox.Show($"Excel Workbook has an invalid Transaction Date: {excelFilename}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        workbook.Close()
                        workbook = Nothing
                        Continue For
                    End If

                    ' Take header info from the first worksheet only
                    Dim CTL_NO As String = ASCMAIN1.Next_Control_No("ICTXFRX1.CTL_NO")
                    ' 11/6/25 - BOL_NO now in J1 instead of M1, or sometimes K1. Trying to be a little flexible
                    Dim BOL_NO As String = workbook.Worksheets(0)?.Cells("J1")?.Value & String.Empty & workbook.Worksheets(0)?.Cells("K1")?.Value
                    Dim TRANS_DATE As Date = GetDateFromWorksheetCell(workbook.Worksheets(0), "L3")
                    Dim FILENAME As String = My.Computer.FileSystem.GetName(excelFilename)

                    If firstWorkBook Then
                        firstWorkBook = False
                        xfrDate = TRANS_DATE
                        DATE_KEY = xfrDate.ToString("yyyyMMdd")
                        rowICTXFRX0.Item("TRANS_DATE") = xfrDate.ToShortDateString
                        Fill_Record("EDTTRXN5_LK", {xfrDate.ToString("dd-MMM-yyyy")})
                    End If

                    If TRANS_DATE.ToString("yyyyMMdd") <> xfrDate.ToString("yyyyMMdd") Then
                        MessageBox.Show($"The selected Transfer Date does not match the date in the Excel Workbook: {excelFilename}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        workbook.Close()
                        workbook = Nothing
                        Continue For
                    End If

                    If BOL_NO.Length = 0 Then
                        MessageBox.Show($"Excel Workbook is missing the Bol No: {excelFilename}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        workbook.Close()
                        workbook = Nothing
                        Continue For
                    End If

                    Dim WHSE_CODE As String = ""
                    If FILENAME.Contains("-") Then
                        WHSE_CODE = FILENAME.Split("-")(0)
                        WHSE_CODE = WHSE_CODE.Substring(0, Math.Min(6, WHSE_CODE.Length))
                        If WHSE_CODE <> "CLA" AndAlso WHSE_CODE <> "CLARTN" Then
                            WHSE_CODE = "CLA"
                        End If
                    End If

                    Dim WHSE_CODE_TO As String = dst.Tables("ICTPARM1").Rows(0).Item("IC_PARM_WHSE_CODE") & String.Empty

                    ' See if the file was already imported.
                    If dst.Tables("ICTXFRX1").Select($"FILENAME = '{FILENAME}'").Length > 0 Then
                        MessageBox.Show($"File was already imported: {FILENAME}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        workbook.Close()
                        workbook = Nothing
                        Continue For
                    End If

                    If dst.Tables("ICTXFRX1").Select($"BOL_NO = '{BOL_NO}'").Length > 0 Then
                        MessageBox.Show($"BOL_NO {BOL_NO} was already imported: {FILENAME}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        workbook.Close()
                        workbook = Nothing
                        Continue For
                    End If

                    Dim drICTXFRX1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTXFRX1 WHERE FILENAME = :PARM1 OR BOL_NO = :PARM2", "VV", {FILENAME, BOL_NO})
                    If drICTXFRX1 IsNot Nothing Then
                        If drICTXFRX1.Item("FILENAME") = FILENAME Then
                            MessageBox.Show($"File was already imported: {FILENAME}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ElseIf drICTXFRX1.Item("BOL_NO") = BOL_NO Then
                            MessageBox.Show($"BOL No was already imported from File: {drICTXFRX1.Item("FILENAME")} on  {drICTXFRX1.Item("INIT_DATE")}", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        workbook.Close()
                        workbook = Nothing
                        Continue For
                    End If

                    ' Reminder - you may get a Clarins 3pl item code. You can convert it to an item by
                    ' looking at ICTITEM1.ITEM_ALT_SORT. If you get multiple hits then use the Active Item.
                    ' If you get multiple Active items then use the most recent LAST_DATE
                    ' If there are only Inactive items then also use the most recent LAST_DATE
                    ' You should validate all ITEM_CODES

                    ' CREATE A LIST OF ALL ITEMS IN THE WORKBOOK
                    Dim lstItemCodes As New List(Of String)

                    For Each worksheet As SpreadsheetGear.IWorksheet In workbook.Worksheets
                        ' start at Row 12 until we reach "TOTAL PALLETS ON LOAD" in col C
                        Dim startingRowNum As Int16 = 11

                        Dim psRange As SpreadsheetGear.IRange = worksheet.UsedRange
                        If Not (psRange Is Nothing) Then
                            Dim headerCell As SpreadsheetGear.IRange = psRange.Cells(0, 0)
                            If Not (headerCell.Value & String.Empty).StartsWith("Packing List", StringComparison.CurrentCultureIgnoreCase) Then
                                Exit For
                            End If
                            ' Determine if the used range is A or B?
                            Dim itemCodeColOffset As Int16 = 2
                            Dim itemQtyColOffset As Int16 = 3
                            If psRange.Address.StartsWith("$B") Then
                                itemCodeColOffset -= 1
                                itemQtyColOffset -= 1
                            End If

                            For Each row As SpreadsheetGear.IRange In psRange.Rows
                                Dim rowNum As Int16 = row.Row
                                If rowNum < startingRowNum Then
                                    Continue For
                                End If
                                Dim cell As SpreadsheetGear.IRange = row.Cells(0, itemCodeColOffset)
                                If (cell.Value & String.Empty).StartsWith("TOTAL", StringComparison.CurrentCultureIgnoreCase) Then
                                    Exit For
                                End If
                                Dim ITEM_CODE As String = row.Cells(0, itemCodeColOffset).Value & String.Empty
                                ITEM_CODE = ITEM_CODE.Trim
                                If ITEM_CODE.Length > 0 AndAlso Not lstItemCodes.Contains(ITEM_CODE) Then
                                    lstItemCodes.Add(ITEM_CODE)
                                End If
                            Next
                        End If
                    Next

                    ASCMAIN1.sql = "SELECT * FROM ICTITEM1 WHERE ITEM_CODE = :PARM1 OR ITEM_ALT_SORT = :PARM1"
                    Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN (SELECT * FROM TABLE(IN_LIST(:PARM1))) or ITEM_ALT_SORT IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))", "ICTITEM1", "C", {String.Join(",", lstItemCodes.ToArray)})

                    For Each worksheet As SpreadsheetGear.IWorksheet In workbook.Worksheets
                        ' start at Row 12 until we reach "TOTAL PALLETS ON LOAD" in col C
                        Dim startingRowNum As Int16 = 11
                        Dim psRange As SpreadsheetGear.IRange = worksheet.UsedRange
                        If Not (psRange Is Nothing) Then
                            ' Need to validate that we're actually looking at a packing list tab
                            Dim cell As SpreadsheetGear.IRange = psRange.Cells(0, 0)
                            If Not (cell.Value & String.Empty).StartsWith("Packing List", StringComparison.CurrentCultureIgnoreCase) Then
                                Exit For
                            End If
                            Dim itemCodeColOffset As Int16 = 2
                            Dim itemQtyColOffset As Int16 = 3
                            If psRange.Address.StartsWith("$B") Then
                                itemCodeColOffset -= 1
                                itemQtyColOffset -= 1
                            End If
                            For Each row As SpreadsheetGear.IRange In psRange.Rows
                                Dim rowNum As Int16 = row.Row
                                If rowNum < startingRowNum Then
                                    Continue For
                                End If
                                cell = row.Cells(0, itemCodeColOffset)
                                If (cell.Value & String.Empty).StartsWith("TOTAL", StringComparison.CurrentCultureIgnoreCase) Then
                                    Exit For
                                End If
                                ' Expect to find Pallet no in A (0), Item Code in C (2), Quantity in E (4)
                                ' 11/6/25 - Item Quantity now in D (3)
                                Dim ITEM_CODE As String = row.Cells(0, itemCodeColOffset).Value & String.Empty
                                ITEM_CODE = ITEM_CODE.Trim
                                If ITEM_CODE.Length = 0 Then
                                    'Stop
                                    ' Prob want a continue For to skip blank lines?
                                    Continue For
                                End If

                                Select Case tbl.Select($"ITEM_CODE = '{ITEM_CODE}' OR ITEM_ALT_SORT = '{ITEM_CODE}'").Length
                                    Case 0
                                        MessageBox.Show($"Cannot locate item {ITEM_CODE} in item master. File {FILENAME}. Line skipped.", "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                        Continue For
                                    Case 1
                                        ITEM_CODE = tbl.Select($"ITEM_CODE = '{ITEM_CODE}' OR ITEM_ALT_SORT = '{ITEM_CODE}'")(0).Item("ITEM_CODE")

                                    Case Else
                                        ' Look for an Active first, then an inactive item. Select by Last Date desc
                                        If tbl.Select($"ITEM_CODE = '{ITEM_CODE}' OR ITEM_ALT_SORT = '{ITEM_CODE}' and ITEM_STATUS = 'A'").Length > 0 Then
                                            For Each rowICTITEM1 As DataRow In tbl.Select("ITEM_STATUS = 'A'", "LAST_DATE DESC")
                                                ITEM_CODE = rowICTITEM1.Item("ITEM_CODE")
                                                Exit For
                                            Next
                                        ElseIf tbl.Select($"ITEM_CODE = '{ITEM_CODE}' OR ITEM_ALT_SORT = '{ITEM_CODE}' and ITEM_STATUS = 'I'").Length > 0 Then
                                            For Each rowICTITEM1 As DataRow In tbl.Select("ITEM_STATUS = 'I'", "LAST_DATE DESC")
                                                ITEM_CODE = rowICTITEM1.Item("ITEM_CODE")
                                                Exit For
                                            Next
                                        End If
                                End Select

                                Dim PALLET_NO As String = row.Cells(0, 0).Value & String.Empty
                                Dim TRAN_QTY As Int32 = row.Cells(0, itemQtyColOffset).Value
                                'Dim TRAN_NOTE As String = row.Cells(0, 6).Value & String.Empty

                                ' Do not let a long note blow up the routine. Truncate it
                                'TRAN_NOTE = TRAN_NOTE.Substring(0, Math.Min(TRAN_NOTE.Length, dst.Tables("ICTXFRX2").Columns("TRAN_NOTE").MaxLength))

                                If dst.Tables("ICTXFRX1").Rows.Find(CTL_NO) Is Nothing Then
                                    Dim rowICTXFRX1 As DataRow = dst.Tables("ICTXFRX1").NewRow
                                    rowICTXFRX1.Item("CTL_NO") = CTL_NO
                                    rowICTXFRX1.Item("BOL_NO") = BOL_NO
                                    rowICTXFRX1.Item("TRANS_DATE") = TRANS_DATE
                                    rowICTXFRX1.Item("FILENAME") = FILENAME
                                    'rowICTXFRX1.item("PO_ORDER_NO") = ""
                                    'rowICTXFRX1.item("PO_TRANS_DATE") = ""
                                    rowICTXFRX1.Item("WHSE_CODE") = WHSE_CODE
                                    rowICTXFRX1.Item("WHSE_CODE_TO") = WHSE_CODE_TO
                                    rowICTXFRX1.Item("INIT_DATE") = INIT_DATE
                                    rowICTXFRX1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                    rowICTXFRX1.Item("XFR_BATCH_NO") = XFR_BATCH_NO
                                    'rowICTXFRX1.item("XFR_NO") = ""
                                    dst.Tables("ICTXFRX1").Rows.Add(rowICTXFRX1)
                                End If

                                Dim rowICTXFRX2 As DataRow = dst.Tables("ICTXFRX2").NewRow
                                rowICTXFRX2.Item("CTL_NO") = CTL_NO
                                rowICTXFRX2.Item("PALLET_NO") = PALLET_NO
                                rowICTXFRX2.Item("ITEM_CODE") = ITEM_CODE
                                rowICTXFRX2.Item("TRAN_QTY") = TRAN_QTY
                                'rowICTXFRX2.Item("TRAN_NOTE") = TRAN_NOTE
                                dst.Tables("ICTXFRX2").Rows.Add(rowICTXFRX2)
                            Next
                        End If
                    Next

                    Dim lstItems As New List(Of String)
                    Dim tblICTITEM1 As DataTable = Nothing

                    For Each drItems As DataRow In dst.Tables("ICTXFRX2").Select($"CTL_NO = '{CTL_NO}'")
                        If Not lstItems.Contains(drItems.Item("ITEM_CODE") & "") Then
                            lstItems.Add(drItems.Item("ITEM_CODE"))
                        End If
                    Next

                    tblICTITEM1 = ASCDATA1.GetDataTable("SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))", "ICTITEM1", "C", {String.Join(",", lstItems.ToArray)})

                    Dim bolColumnName As String = String.Empty
                    For iLoop As Int16 = 1 To 30
                        bolColumnName = $"BOL_NO_{iLoop}"
                        If grdItemsSelected.DisplayLayout.Bands(0).Columns(bolColumnName).Tag = String.Empty Then
                            grdItemsSelected.DisplayLayout.Bands(0).Columns(bolColumnName).Tag = CTL_NO
                            Exit For
                        End If
                    Next
                    grdItemsSelected.DisplayLayout.Bands(0).Columns(bolColumnName).Hidden = False
                    grdItemsSelected.DisplayLayout.Bands(0).Columns(bolColumnName).Header.Caption = BOL_NO
                    grdItemsSelected.DisplayLayout.Bands(0).Columns(bolColumnName).Header.Appearance.BackColor = Color.LightGreen

                    For Each ITEM_CODE As String In lstItems
                        If dst.Tables("ICTXFRX2").Select($"ITEM_CODE = '{ITEM_CODE}'").Length > 0 Then

                            Dim drSelected As DataRow = Nothing
                            If dst.Tables("ItemsSelected").Select($"ITEM_CODE = '{ITEM_CODE}'").Length = 0 Then
                                drSelected = dst.Tables("ItemsSelected").NewRow
                                drSelected.Item("ITEM_CODE") = ITEM_CODE
                                dst.Tables("ItemsSelected").Rows.Add(drSelected)
                            Else
                                drSelected = dst.Tables("ItemsSelected").Select($"ITEM_CODE = '{ITEM_CODE}'")(0)
                            End If

                            Dim drICTITEM1 As DataRow = tblICTITEM1.Rows.Find({ITEM_CODE})
                            If drICTITEM1 IsNot Nothing Then
                                drSelected.Item("ITEM_DESC") = drICTITEM1.Item("ITEM_DESC")

                                For Each drICTXFRX2 As DataRow In dst.Tables("ICTXFRX2").Select($"ITEM_CODE = '{ITEM_CODE}'")
                                    drICTXFRX2.Item("ITEM_DESC") = drICTITEM1.Item("ITEM_DESC")
                                Next
                            End If

                            drSelected.Item(bolColumnName) = Val(dst.Tables("ICTXFRX2").Compute("SUM(TRAN_QTY)", $"CTL_NO = '{CTL_NO}' AND ITEM_CODE = '{ITEM_CODE}' ") & String.Empty)
                        End If
                    Next

                    ' Query to use to get previous transfered items/quantity
                    Dim sql As String = "SELECT ICTXFRX2.ITEM_CODE, SUM(ICTXFRX2.TRAN_QTY) USED_QTY
                                        FROM ICTXFRX1, ICTXFRX2
                                        WHERE ICTXFRX1.CTL_NO = ICTXFRX2.CTL_NO
                                        AND TRUNC(ICTXFRX1.TRANS_DATE) = :PARM1
                                        AND WHSE_CODE = :PARM2
                                        GROUP BY ICTXFRX2.ITEM_CODE"
                    Dim tblICTXFRX2 As DataTable = ASCDATA1.GetDataTable(sql, "", "VV", {TRANS_DATE.ToString("dd-MMM-yyyy"), xfrWhseCode})
                    For Each row As DataRow In tblICTXFRX2.Select("")
                        dst.Tables("ItemsSelected").Select($"ITEM_CODE = '{row.Item("ITEM_CODE")}'")(0).Item("USED_QTY") = Val(row.Item("USED_QTY") & String.Empty)
                    Next

                    For Each drEDTTRXN5_LK As DataRow In dst.Tables("EDTTRXN5_LK").Select("")
                        Dim ITEM_CODE As String = drEDTTRXN5_LK.Item("ITEM_CODE") & String.Empty
                        If dst.Tables("ItemsSelected").Select($"ITEM_CODE = '{ITEM_CODE}'").Length > 0 Then
                            dst.Tables("ItemsSelected").Select($"ITEM_CODE = '{ITEM_CODE}'")(0).Item("TRAN_QTY_5") = drEDTTRXN5_LK.Item("TRAN_QTY")
                        Else
                            Dim dr As DataRow = dst.Tables("ItemsSelected").NewRow
                            dr.Item("ITEM_CODE") = drEDTTRXN5_LK.Item("ITEM_CODE")
                            dr.Item("ITEM_DESC") = drEDTTRXN5_LK.Item("ITEM_DESC")
                            dr.Item("TRAN_QTY_5") = drEDTTRXN5_LK.Item("TRAN_QTY")
                            dst.Tables("ItemsSelected").Rows.Add(dr)
                        End If
                    Next

                    dst.Tables("ICTXFRX1").AcceptChanges()
                    dst.Tables("ICTXFRX2").AcceptChanges()
                    dst.Tables("ItemsSelected").AcceptChanges()

                    For Each grd As UltraGrid In {grdICTXFRX1, grdICTXFRX2, grdItemsSelected}
                        If grd.Rows.Count <= 50 Then
                            AutoFitGridColumns(grd)
                        End If
                    Next

                    Sort_grdColumns(grdICTXFRX1, "BOL_NO")
                    Sort_grdColumns(grdICTXFRX2, "ITEM_CODE")
                    Sort_grdColumns(grdItemsSelected, "ITEM_CODE")

                    grdICTXFRX0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdItemsSelected.DisplayLayout.Bands(0).Columns("BOL_NO_TOTAL").Header.VisiblePosition = grdItemsSelected.DisplayLayout.Bands(0).Columns.Count
                    grdItemsSelected.DisplayLayout.Bands(0).Columns("TRAN_DIFF").Header.VisiblePosition = grdItemsSelected.DisplayLayout.Bands(0).Columns.Count

                    grdItemsSelected.Text = $"Item Reconciliation for Warehouse {xfrWhseCode}, Transfer Date {xfrDate.ToString("MM/dd/yyyy")}"
                Next
            End Using

            SortColumnsByBolNo()
            splImportDetails.Panel1Collapsed = True
            splImportDetails.Panel2Collapsed = False
            splItemRecon.Panel2Collapsed = True
            grdItemsSelected.DisplayLayout.Bands(0).Columns("BOL_NO_TOTAL").Hidden = False

            Return True

        Catch ex As Exception
            MessageBox.Show(excelFilename & " - " & ex.Message, "Import Excel File", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Clear_Record()
            Return False
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Function

    Private Sub Send3PLPurchaseOrder(CTL_NO As String)
        ' This needs to create a Transfer from CLA -> ADS or CLARTN -> ADSRTN
        ' Then use the transfer to create an XML PO and send it to ADS
        Try
            Dim XFR_NO As String = String.Empty
            If CreateTransfer(CTL_NO) Then
                If TransferPurchaseOrder(XFR_NO) Then
                    MessageBox.Show("Purchase Order sent to 3PL.", "Create Transfer", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Could NOT send Purchase Order to 3PL.", "Create Transfer", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            Else
                MessageBox.Show("Could Not create Transfer.", "Create Transfer", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Send 3PL Purchase Order", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
        End Try
    End Sub

    Private Function CreateTransfer(ByVal CTL_NO As String) As Boolean

        Dim tblICTITEM1 As DataTable = ASCDATA1.GetDataTable("SELECT ICTITEM1.ITEM_CODE, ICTITEM1.COST_CATGY_CODE, ICTITEM1.PROD_CODE, ICTCOSTC.ITEM_COST_TOTAL
                                            FROM ICTITEM1, ICTCOSTC
                                            WHERE ICTITEM1.ITEM_CODE = ICTCOSTC.ITEM_CODE (+)
                                            AND ICTITEM1.ITEM_CODE IN (SELECT ITEM_CODE FROM ICTXFRX2 WHERE CTL_NO = :PARM1)", "", "V", {CTL_NO})
        tblICTITEM1.PrimaryKey = New DataColumn() {tblICTITEM1.Columns("ITEM_CODE")}

        Dim drICTXFRX1 As DataRow = dst.Tables("ICTXFRX1").Rows.Find(CTL_NO)
        Dim XFR_NO As String = drICTXFRX1.Item("XFR_NO")

        Dim WHSE_CODE As String = drICTXFRX1.Item("WHSE_CODE") & String.Empty
        Dim WHSE_CODE_TO As String = drICTXFRX1.Item("WHSE_CODE_TO") & String.Empty

        ' Always clear the table since we may procees multiple transfers and we call TAC.ICCMAIN1.Update_Transfer
        ' We only want one transfer in memory
        dst.Tables("ICTIXFR1").Rows.Clear()
        dst.Tables("ICTIXFR2").Rows.Clear()

        Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
        Dim locationSupport As Boolean = (drICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")

        drICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE_TO)

        Dim drICTIXFR1 As DataRow = dst.Tables("ICTIXFR1").NewRow
        drICTIXFR1.Item("XFR_NO") = XFR_NO
        drICTIXFR1.Item("XFR_DATE") = DateTime.Now.ToShortDateString
        drICTIXFR1.Item("WHSE_CODE") = WHSE_CODE
        drICTIXFR1.Item("WHSE_CODE_TO") = drICTWHSE1.Item("WHSE_LOC_XFR")
        drICTIXFR1.Item("XFR_NOTE") = ""
        drICTIXFR1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        drICTIXFR1.Item("INIT_DATE") = DateTime.Now
        drICTIXFR1.Item("REGISTER_IND") = "0"
        'drICTIXFR1.Item("REGISTER_XNO") = ""
        drICTIXFR1.Item("XFR_SOURCE") = XFR_SOURCE ' R, S, and E already exist
        drICTIXFR1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        drICTIXFR1.Item("TOTAL_COSTS") = "0"
        drICTIXFR1.Item("RTRN_NO") = ""
        drICTIXFR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        drICTIXFR1.Item("LAST_DATE") = DateTime.Now
        'drICTIXFR1.Item("REVERSED_BY_XFR_NO") = ""
        'drICTIXFR1.Item("REVERSES_XFR_NO") = ""
        'drICTIXFR1.Item("XFR_REF") = ""
        drICTIXFR1.Item("JOURNAL_IND") = "0"
        'drICTIXFR1.Item("JOURNAL_XNO") = ""
        dst.Tables("ICTIXFR1").Rows.Add(drICTIXFR1)

        Dim XFR_LNO As Int32 = 1
        For Each drICTXFRX2 As DataRow In dst.Tables("ICTXFRX2").Select($"CTL_NO = '{CTL_NO}'", "ITEM_CODE")

            Dim ITEM_CODE As String = drICTXFRX2.Item("ITEM_CODE")
            Dim drICTIXFR2 As DataRow = Nothing

            If dst.Tables("ICTIXFR2").Select($"ITEM_CODE = '{ITEM_CODE}'").Length > 0 Then
                drICTIXFR2 = dst.Tables("ICTIXFR2").Select($"ITEM_CODE = '{ITEM_CODE}'")(0)
                drICTIXFR2.Item("XFR_QTY") = drICTIXFR2.Item("XFR_QTY") + drICTXFRX2.Item("TRAN_QTY")
                Continue For
            End If

            drICTIXFR2 = dst.Tables("ICTIXFR2").NewRow
            drICTIXFR2.Item("XFR_NO") = XFR_NO
            drICTIXFR2.Item("XFR_LNO") = XFR_LNO
            XFR_LNO += 1
            drICTIXFR2.Item("ITEM_CODE") = drICTXFRX2.Item("ITEM_CODE")
            drICTIXFR2.Item("XFR_QTY") = drICTXFRX2.Item("TRAN_QTY")

            Dim drICTITEM1 As DataRow = tblICTITEM1.Rows.Find(drICTIXFR2.Item("ITEM_CODE"))
            If drICTITEM1 IsNot Nothing Then
                drICTIXFR2.Item("COST_CATGY_CODE") = drICTITEM1.Item("COST_CATGY_CODE")
                drICTIXFR2.Item("PROD_CODE") = drICTITEM1.Item("PROD_CODE")
                drICTIXFR2.Item("ITEM_COST_STD") = drICTITEM1.Item("ITEM_COST_TOTAL")
            End If

            drICTIXFR2.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            'drICTIXFR2.item("LOCATION_CODE") = ""
            'drICTIXFR2.Item("BAR_CODE") = ""
            dst.Tables("ICTIXFR2").Rows.Add(drICTIXFR2)
        Next

        drICTIXFR1.Item("TOTAL_COSTS") = Val(dst.Tables("ICTIXFR2").Compute("SUM(TOTAL_COSTS)", $"XFR_NO = '{XFR_NO}'") & String.Empty)

        Update_Record_TDA("ICTIXFR1")
        Update_Record_TDA("ICTIXFR2")

        TAC.ICCMAIN1.Update_Transfer(Me)

        If locationSupport Then
            Update_WHTLOCBX("T")
        End If

        drICTWHSE1 = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE_TO)
        If drICTWHSE1.Item("WHSE_LOCATOR") & "" = "1" Then
            Update_WHTLOCBX("X")
        End If

        Return True

    End Function

    Private Function TransferPurchaseOrder(ByVal XFR_NO As String) As Boolean

        Dim XMIT_NO As String = String.Empty

        ' Update all changed items
        Try
            XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "All", "ADS")
        Catch ex As Exception

        End Try

        XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "XFR_NO", XFR_NO, "ADS")

        If XMIT_NO.Length > 0 Then
            Dim XFR_REF As String = "X" & Val(XFR_NO).ToString.PadLeft(5, "0")
            ' Do this because the process sends back a Transmit No when a file is NOT sent.
            Dim drICTIXFR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTIXFR1 WHERE XFR_NO = :PARM1 AND XFR_NOTE = :PARM2", "VV", {XFR_NO, XFR_REF})

            If drICTIXFR1 Is Nothing Then
                Return False
            End If

            ASCMAIN1.sql = "UPDATE ICTIXFR1 SET XFR_REF = :PARM1, LAST_DATE = SYSDATE, LAST_OPER = :PARM2 WHERE XFR_NO = :PARM3"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", {XFR_REF, ASCMAIN1.USER_ID, XFR_NO})

            ASCMAIN1.sql = "UPDATE ICTXFRX1 SET PO_ORDER_NO = :PARM1, PO_TRANS_DATE = SYSDATE WHERE XFR_NO = :PARM2"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {XFR_REF, XFR_NO})

            Return True
        Else
            Return False
        End If

    End Function

    Sub Update_WHTLOCBX(TRAN_TYPE As String)
        Dim rowICTIXFR1 As DataRow = dst.Tables("ICTIXFR1").Rows(0)
        TAC.ICCMAIN1.Update_WHTLOCBX(TRAN_TYPE, rowICTIXFR1.Item("XFR_NO"))
    End Sub

    Private Sub grdICTXFRX1_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdICTXFRX1.BeforeRowUpdate
        Dim WHSE_CODE As String = e.Row.Cells("WHSE_CODE").Value & String.Empty
        Dim WHSE_CODE_TO As String = e.Row.Cells("WHSE_CODE_TO").Value & String.Empty

        WHSE_CODE_TO = WHSE_CODE_TO.Trim.ToUpper
        e.Row.Cells("WHSE_CODE_TO").Value = WHSE_CODE_TO

        If WHSE_CODE.Length = 0 OrElse WHSE_CODE_TO.Length = 0 Then
            MessageBox.Show("Form Warehouse and To Warehouse are required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If WHSE_CODE = WHSE_CODE_TO Then
            MessageBox.Show("Form Warehouse and To Warehouse must be different.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        Dim drICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE_TO)
        If drICTWHSE1 Is Nothing Then
            MessageBox.Show("To Warehouse is invalid.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

        If drICTWHSE1.Item("LP_CODE") & String.Empty = String.Empty Then
            MessageBox.Show("To Warehouse is not a 3PL warehouse.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
            Exit Sub
        End If

    End Sub

    Private Sub grdICTXFRX1_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdICTXFRX1.ClickCellButton

        If grdICTXFRX1.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "WHSE_CODE_TO"
                sql_where = "LP_CODE IS NOT NULL AND WHSE_LOC_XFR IS NOT NULL AND WHSE_STATUS='A'"
        End Select
        grdClickCellButton(grdICTXFRX1, sql_where, False, "WHSE_CODE_TO", "WHSE_CODE")
    End Sub

    Private DeletedRows As New List(Of String)
    Private Sub grdICTXFRX1_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdICTXFRX1.BeforeRowsDeleted

        DeletedRows.Clear()
        Dim lstBolNos As New List(Of String)
        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Rows
            lstBolNos.Add(row.Cells("BOL_NO").Value)
            DeletedRows.Add(row.Cells("CTL_NO").Value)
        Next

        e.DisplayPromptMsg = False

        Dim zmsg As String = "Do you want to Remove the following Bol Nos?" & Environment.NewLine & String.Join(Environment.NewLine, lstBolNos.ToArray)
        If MessageBox.Show(zmsg, "Remove Bol Nos", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
            DeletedRows.Clear()
            e.Cancel = True
        End If
    End Sub

    Private Sub grdICTXFRX1_AfterRowsDeleted(sender As Object, e As EventArgs) Handles grdICTXFRX1.AfterRowsDeleted

        For Each CTL_NO As String In DeletedRows
            For Each grdCOl As UltraGridColumn In grdItemsSelected.DisplayLayout.Bands(0).Columns
                If grdCOl.Tag = CTL_NO Then
                    Dim COLUMN_NAME As String = grdCOl.Key
                    For Each drItemsSelected As DataRow In dst.Tables("ItemsSelected").Select("")
                        drItemsSelected.Item(COLUMN_NAME) = 0
                    Next
                    grdCOl.Tag = String.Empty
                    grdCOl.Hidden = True
                End If
            Next
        Next

        SortColumnsByBolNo()
    End Sub

    Private Sub SortColumnsByBolNo()

        Dim dict As New Dictionary(Of String, String)
        For Each grdCOl As UltraGridColumn In grdItemsSelected.DisplayLayout.Bands(0).Columns
            If grdCOl.Tag <> String.Empty Then
                dict.Add(grdCOl.Header.Caption, grdCOl.Key)
            End If
        Next

        Dim sortedByKey = dict.OrderBy(Function(x) x.Key)
        Dim iLoop As Int16 = 1
        For Each kvp As KeyValuePair(Of String, String) In sortedByKey
            Dim COLUMN_NAME As String = kvp.Value
            grdItemsSelected.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.VisiblePosition = iLoop + 4
            iLoop += 1
        Next

        grdItemsSelected.DisplayLayout.Bands(0).Columns("BOL_NO_TOTAL").Header.VisiblePosition = grdItemsSelected.DisplayLayout.Bands(0).Columns.Count
        grdItemsSelected.DisplayLayout.Bands(0).Columns("TRAN_DIFF").Header.VisiblePosition = grdItemsSelected.DisplayLayout.Bands(0).Columns.Count
    End Sub

    Private Sub grdICTXFRX0_AfterRowActivate(sender As Object, e As EventArgs) Handles grdICTXFRX0.AfterRowActivate
        If grdICTXFRX0.ActiveRow Is Nothing Then
            Exit Sub
        End If

        Dim TRANS_DATE As Date = CDate(grdICTXFRX0.ActiveRow.Cells("TRANS_DATE").Value & String.Empty)

        dst.Tables("ItemsSelected").Rows.Clear()

        ASCMAIN1.sql = "SELECT ITEM_CODE, ITEM_DESC, SUM(TRAN_QTY_5) TRAN_QTY_5, SUM(USED_QTY) USED_QTY
                        FROM
                        (
                            SELECT ICTXFRX2.ITEM_CODE, ICTITEM1.ITEM_DESC, 0 TRAN_QTY_5, SUM(TRAN_QTY) USED_QTY
                                    FROM ICTXFRX1, ICTXFRX2, ICTITEM1
                                    WHERE ICTXFRX1.CTL_NO = ICTXFRX2.CTL_NO
                                    AND ICTXFRX1.TRANS_DATE = :PARM1
                                    AND ICTXFRX2.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                                    GROUP BY ICTXFRX1.BOL_NO, ICTXFRX2.ITEM_CODE, ICTITEM1.ITEM_DESC
                            UNION
                                SELECT EDTTRXN5.ITEM_CODE, ICTITEM1.ITEM_DESC, SUM(EDTTRXN5.TRAN_QTY) TRAN_QTY_5, 0 USED_QTY
                                FROM EDTTRXN5, ICTITEM1
                                WHERE EDTTRXN5.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                                AND TRANS_DATE = :PARM1
                                GROUP BY EDTTRXN5.ITEM_CODE, ICTITEM1.ITEM_DESC
                        )
                        GROUP BY ITEM_CODE, ITEM_DESC"
        Fill_Records("ItemsSelected", "", True, ASCMAIN1.sql.Replace(":PARM1", $"'{TRANS_DATE.ToString("dd-MMM-yyyy")}'"))
        Sort_grdColumns(grdItemsSelected, "ITEM_CODE")

        grdItemsSelected.Text = $"Item Reconciliation for Import Date {TRANS_DATE.ToShortDateString}"

    End Sub

    Private Sub grdItemsSelected_AfterRowActivate(sender As Object, e As EventArgs) Handles grdItemsSelected.AfterRowActivate

        If ScreenMode Then
            Exit Sub
        End If

        If grdICTXFRX0.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If grdItemsSelected.ActiveRow Is Nothing Then
            Exit Sub
        End If

        dst.Tables("ItemsRecon").Rows.Clear()
        Dim TRANS_DATE As Date = CDate(grdICTXFRX0.ActiveRow.Cells("TRANS_DATE").Value & String.Empty)
        Dim ITEM_CODE As String = grdItemsSelected.ActiveRow.Cells("ITEM_CODE").Value & String.Empty
        Fill_Records("ItemsRecon", {TRANS_DATE.ToString("dd-MMM-yyyy"), ITEM_CODE})

        grdItemsRecon.Text = $"Item Reconciliation for Import Date {TRANS_DATE.ToShortDateString}, Item {ITEM_CODE}"
        Sort_grdColumns(grdItemsRecon, "BOL_NO")

    End Sub

#End Region

End Class