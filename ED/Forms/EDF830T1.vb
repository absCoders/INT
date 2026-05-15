
Imports Infragistics.Win.UltraWinGrid

Public Class EDF830T1

#Region "Class Variables"

    Private EDI_DOC_SEQ_NO As String = String.Empty
    Private EDT830T1_WK As String = String.Empty
    Private EDT830T1_RS As String = String.Empty
    Private tblItems As DataTable = Nothing

    Private grdLayoutFile As String = String.Empty

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    ''' <summary>
    ''' ** Sets up required Data Tables and intializes form controls
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "EDT830T1", "*")
            EDT830T1_WK = ASCMAIN1.Temp_Table("SELECT * FROM EDT830T1 WHERE ROWNUM < 1")
            EDT830T1_RS = ASCMAIN1.Temp_Table("SELECT * FROM EDT830T1 WHERE ROWNUM < 1")
        End With

        grdEDT830T1.DataSource = dst.Tables("EDT830T1")
        spldata.Panel2Collapsed = True

        grdLayoutFile = IO.Path.Combine(ASCMAIN1.Folders("Work"), "grdEDT830T1.xml")
        If My.Computer.FileSystem.FileExists(grdLayoutFile) Then
            My.Computer.FileSystem.DeleteFile(grdLayoutFile)
        End If
        grdEDT830T1X.DisplayLayout.SaveAsXml(grdLayoutFile, PropertyCategories.All)

    End Sub

    ''' <summary>
    ''' Clear tables and controls based on the current state of the screen
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ClearRecord()

        EnforceConstraints(False)

        spldata.Panel1Collapsed = False
        spldata.Panel2Collapsed = True

        Fill_Records("EDT830T1", String.Empty, True, "SELECT * FROM EDT830T1")
        Sort_grdColumns(grdEDT830T1, "EDI_DOC_SEQ_NO".ToLower)

        If tblItems IsNot Nothing Then
            tblItems.Rows.Clear()
        End If

        EDI_DOC_SEQ_NO = String.Empty

        EnforceConstraints(True)
    End Sub

    ''' <summary>
    ''' Load up changes to go into 832 files
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub LoadRecord()

        ASCMAIN1.Progress("Evaluate 830 Records", String.Empty)

        ASCDATA1.ExecuteSQL($"TRUNCATE TABLE {EDT830T1_WK}")
        ASCDATA1.ExecuteSQL($"DROP TABLE {EDT830T1_WK}")

        ASCDATA1.ExecuteSQL($"TRUNCATE TABLE {EDT830T1_RS}")
        ASCDATA1.ExecuteSQL($"DROP TABLE {EDT830T1_RS}")

        ASCMAIN1.sql = $"CREATE TABLE {EDT830T1_WK} AS
                            SELECT T1.EDI_DOC_SEQ_NO, 
                            M1.CUST_CODE, I1.ITEM_CODE, I1.ITEM_DESC, I1.ITEM_RETAIL_PRICE, NVL(I1.ITEM_EAN_CODE, I1.ITEM_UPC_CODE) EAN_UPC,
                            0.00 DISC_OFF_MSRP, 0.00 NET_PRICE, T2.EDI_QTY_TOTAL NUM_STORES, TRUNC(T3.EDI_FST_END_DATE) FORECAST_WK_END, T3.EDI_FST_QTY
                            FROM EDT830T1 T1
                            JOIN EDT830T2 T2 ON T1.EDI_DOC_SEQ_NO=T2.EDI_DOC_SEQ_NO
                            JOIN EDT830T3 T3 ON T2.EDI_DOC_SEQ_NO=T3.EDI_DOC_SEQ_NO AND T2.EDI_DTL_SEQ=T3.EDI_DTL_SEQ
                            JOIN ICTITEM1 I1 ON (T2.EDI_EAN = I1.ITEM_EAN_CODE AND I1.ITEM_EAN_CODE IS NOT NULL) OR (T2.EDI_UPC = I1.ITEM_UPC_CODE AND I1.ITEM_UPC_CODE IS NOT NULL)
                            JOIN EDTTRPM1 M1 ON T1.EDI_TP_QUAL = TRIM(M1.EDI_TP_QUAL) AND TRIM(T1.EDI_TP_ID) = M1.EDI_TP_ID AND M1.EDI_DOC_NO = '830'
                            WHERE T1.EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}'
                            ORDER BY T3.EDI_DOC_SEQ_NO, T2.EDI_DTL_SEQ, T3.EDI_FST_SEQ"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCMAIN1.sql = $"BEGIN DECLARE CURSOR C1 IS
        '                    SELECT DISTINCT WK.CUST_CODE, WK.ITEM_CODE, S1.PRICE_BASE_DPCT
        '                    FROM {EDT830T1_WK} WK, ARTCUST1 T1, SOTPCLS1 S1
        '                    WHERE WK.CUST_CODE = T1.CUST_CODE
        '                    AND T1.PRICE_CLASS_CODE = S1.PRICE_CLASS_CODE;
        '                    BEGIN FOR R1 IN C1 LOOP
        '                        UPDATE {EDT830T1_WK} SET DISC_OFF_MSRP = R1.PRICE_BASE_DPCT WHERE CUST_CODE = R1.CUST_CODE AND ITEM_CODE = R1.ITEM_CODE;
        '                    END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCMAIN1.sql = $"UPDATE {EDT830T1_WK} SET NET_PRICE = ITEM_RETAIL_PRICE - (ITEM_RETAIL_PRICE * (DISC_OFF_MSRP / 100)) WHERE DISC_OFF_MSRP > 0"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCMAIN1.sql = $"UPDATE {EDT830T1_WK} SET NET_PRICE = (SELECT ITEM_PRICE FROM SOTPRIC2 WHERE PRICE_LIST_CODE = {EDT830T1_WK}.CUST_CODE AND ITEM_CODE = {EDT830T1_WK}.ITEM_CODE) WHERE NET_PRICE = 0"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCMAIN1.sql = $"UPDATE {EDT830T1_WK} SET DISC_OFF_MSRP = TRUNC((NET_PRICE / ITEM_RETAIL_PRICE) * 100)
        '                    WHERE ITEM_RETAIL_PRICE <> 0
        '                    AND DISC_OFF_MSRP = 0"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = $"DECLARE
                              v_sql VARCHAR2(32767);
                              v_cols VARCHAR2(32767);
                            BEGIN
                              SELECT LISTAGG('''' || FORECAST_WK_END || '''', ', ') WITHIN GROUP (ORDER BY FORECAST_WK_END)
                              INTO v_cols
                              FROM (SELECT DISTINCT FORECAST_WK_END FROM {EDT830T1_WK} order by FORECAST_WK_END);
                                v_sql := 'Create Table {EDT830T1_RS}  as Select * FROM {EDT830T1_WK}  pivot (sum(edi_fst_qty) FOR FORECAST_WK_END IN (' || V_COLS || '))';
                              EXECUTE IMMEDIATE v_sql;
                            end;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Dim wkTbl As DataTable = ASCDATA1.GetDataTable($"SELECT * FROM {EDT830T1_RS}", EDT830T1_RS)
        For Each col As DataColumn In wkTbl.Columns
            col.ColumnName = col.ColumnName.Replace("'", "")
        Next

        If wkTbl.Rows.Count > 0 Then
            Dim CUST_CODE As String = wkTbl.Rows(0).Item("CUST_CODE") & String.Empty
            If CUST_CODE.Length > 0 Then
                ASCMAIN1.sql = "SELECT DISTINCT SOTORDR2.ITEM_CODE
                                    FROM SOTORDR1, SOTORDR2
                                    WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO
                                    AND SOTORDR1.CUST_CODE = :PARM1"
                tblItems = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {CUST_CODE})
                tblItems.PrimaryKey = {tblItems.Columns("ITEM_CODE")}
            End If

            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim PRICE_LIST_CODE As String = rowARTCUST1.Item("PRICE_LIST_CODE") & ""
            Dim PRICE_LIST_CODE_ALLO As String = rowARTCUST1.Item("PRICE_LIST_CODE") & ""

            Dim CUST_CODE_ALLO As String = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
            If CUST_CODE_ALLO <> "" Then
                Dim rowARTCUST1_ALLO As DataRow = LookUp("ARTCUST1", CUST_CODE_ALLO)
                If rowARTCUST1_ALLO IsNot Nothing Then
                    PRICE_LIST_CODE_ALLO = rowARTCUST1_ALLO.Item("PRICE_LIST_CODE") & ""
                End If
            End If

            Dim PRICE_CLASS_CODE As String = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
            Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
            Dim PRICE_BASIS As String = rowSOTPCLS1.Item("PRICE_BASIS") & ""
            Dim PRICE_BASE_DPCT As Decimal = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
            Dim ORDR_DATE_BOOKED As Date = DateTime.Now
            Dim ITEM_RETAIL_PRICE As Decimal = 0

            Dim rowEDT830T1 As DataRow = dst.Tables("EDT830T1").Rows.Find(EDI_DOC_SEQ_NO)
            If rowEDT830T1 IsNot Nothing Then
                If rowEDT830T1.Item("TRANS_FST_DATE_START") & String.Empty <> String.Empty Then
                    If IsDate(rowEDT830T1.Item("TRANS_FST_DATE_START") & String.Empty) Then
                        ORDR_DATE_BOOKED = CDate(rowEDT830T1.Item("TRANS_FST_DATE_START") & String.Empty).ToShortDateString
                    End If
                End If
            End If

            For Each rowData As DataRow In wkTbl.Select("", "ITEM_CODE")
                Dim ITEM_CODE As String = rowData.Item("ITEM_CODE") & ""
                rowData.Item("ITEM_RETAIL_PRICE") = DBNull.Value
                rowData.Item("NET_PRICE") = DBNull.Value
                rowData.Item("DISC_OFF_MSRP") = DBNull.Value

                ASCMAIN1.Progress("-", ITEM_CODE)
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)

                If rowICTITEM1 IsNot Nothing Then
                    ITEM_RETAIL_PRICE = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                    Dim ITEM_RETAIL_PRICE_NEW As Decimal = Val(rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE") & "")
                    Dim ITEM_NEW_RETAIL_PRICE_DATE As String = rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE_DATE") & ""

                    If ITEM_NEW_RETAIL_PRICE_DATE <> "" And ITEM_RETAIL_PRICE_NEW <> 0 Then
                        If Format(ORDR_DATE_BOOKED, "yyyyMMdd") >= Format(CDate(ITEM_NEW_RETAIL_PRICE_DATE), "yyyyMMdd") Then
                            ITEM_RETAIL_PRICE = ITEM_RETAIL_PRICE_NEW
                        End If
                    End If

                    Dim NET_PRICE As Decimal = TAC.SOCMAIN1.Get_Price _
                                                        (Me,
                                                         PRICE_LIST_CODE,
                                                         PRICE_LIST_CODE_ALLO,
                                                         PRICE_BASIS,
                                                         PRICE_BASE_DPCT,
                                                         ITEM_CODE,
                                                         rowICTITEM1,
                                                         ORDR_DATE_BOOKED,
                                                         0)

                    rowData.Item("ITEM_RETAIL_PRICE") = ITEM_RETAIL_PRICE
                    rowData.Item("NET_PRICE") = NET_PRICE

                    If ITEM_RETAIL_PRICE > 0 Then
                        rowData.Item("DISC_OFF_MSRP") = ((ITEM_RETAIL_PRICE - NET_PRICE) / ITEM_RETAIL_PRICE) * 100
                    End If
                End If
            Next
        End If

        wkTbl.Columns.Add("TOTAL_NUM", GetType(Int32))
        wkTbl.Columns.Add("TOTAL_SALES", GetType(Double))

        grdEDT830T1X.DisplayLayout.Bands(0).Summaries.Clear()
        grdEDT830T1X.DataSource = Nothing
        grdEDT830T1X.DisplayLayout.Reset()
        grdEDT830T1X.DisplayLayout.LoadFromXml(grdLayoutFile, PropertyCategories.All)

        With grdEDT830T1X.DisplayLayout.Override
            .CellClickAction = CellClickAction.CellSelect
            .RowSelectors = DefaultableBoolean.True

            .AllowAddNew = AllowAddNew.No
            .AllowDelete = DefaultableBoolean.False
            .AllowUpdate = DefaultableBoolean.False
        End With

        grdEDT830T1X.DataSource = wkTbl
        Dim sqlTotal As String = String.Empty

        For Each grdCol As UltraWinGrid.UltraGridColumn In grdEDT830T1X.DisplayLayout.Bands(0).Columns
            Select Case grdCol.Key
                Case "EDI_DOC_SEQ_NO"
                    grdCol.Hidden = True

                Case "CUST_CODE", "ITEM_CODE", "ITEM_DESC"
                    grdCol.Header.Caption = StrConv(grdCol.Key.Replace("_", " "), VbStrConv.ProperCase)
                    grdCol.CellAppearance.BackColor = Drawing.Color.LightGray

                    If grdCol.Key = "ITEM_CODE" Then
                        Create_Summary(grdEDT830T1X, grdCol.Key, "Count")
                    End If

                Case "EAN_UPC"
                    grdCol.Header.Caption = "EAN / UPC"
                    grdCol.CellAppearance.BackColor = Drawing.Color.LightGray

                Case "ITEM_RETAIL_PRICE", "DISC_OFF_MSRP", "NET_PRICE"
                    grdCol.Header.Caption = StrConv(grdCol.Key.Replace("_", " "), VbStrConv.ProperCase)
                    grdCol.Format = "#,##0.00"
                    grdCol.CellAppearance.BackColor = Drawing.Color.LightGreen

                Case "NUM_STORES"
                    grdCol.Header.Caption = StrConv(grdCol.Key.Replace("_", " "), VbStrConv.ProperCase)
                    grdCol.Format = "#,##0"

                Case "FORECAST_WK_END"
                    grdCol.Header.Caption = StrConv(grdCol.Key.Replace("_", " "), VbStrConv.ProperCase)
                    'grdCol.Format = "MM/dd/yyyy"

                Case "EDI_FST_QTY"
                    grdCol.Header.Caption = "Forecast QTY"
                    grdCol.Format = "#,##0"

                Case "TOTAL_NUM"
                    grdCol.Header.Caption = "Total #"
                    grdCol.Format = "#,##0"
                    grdCol.CellAppearance.BackColor = Drawing.Color.LightGreen

                Case "TOTAL_SALES"
                    grdCol.Header.Caption = "Total $$"
                    grdCol.Format = "#,##0.00"
                    grdCol.CellAppearance.BackColor = Drawing.Color.LightGreen

                Case Else
                    If IsDate(grdCol.Key.Replace("'", "")) Then
                        grdCol.Header.Caption = grdCol.Key.Replace("'", "")
                        grdCol.Format = "#,##0"
                        Create_Summary(grdEDT830T1X, grdCol.Key, "Sum")

                        sqlTotal &= $"+ ISNULL([{grdCol.Key}], 0) "
                    End If

            End Select
        Next

        If sqlTotal.Length > 0 Then
            sqlTotal = sqlTotal.Substring(1).Trim
            With wkTbl
                .Columns("TOTAL_NUM").Expression = sqlTotal
                Create_Summary(grdEDT830T1X, "TOTAL_NUM", "Sum")

                .Columns("TOTAL_SALES").Expression = "ISNULL(TOTAL_NUM, 0) * ISNULL(NET_PRICE, 0)"
                Create_Summary(grdEDT830T1X, "TOTAL_SALES", "Sum")
            End With
        End If

        grdEDT830T1X.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
        grdEDT830T1X.DisplayLayout.UseFixedHeaders = True
        grdEDT830T1X.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.Fixed = True
        grdEDT830T1X.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True
        grdEDT830T1X.DisplayLayout.Bands(0).Columns("ITEM_DESC").Header.Fixed = True
        grdEDT830T1X.DisplayLayout.Bands(0).Columns("EAN_UPC").Header.Fixed = True

        grdEDT830T1X.DisplayLayout.Bands(0).Columns("TOTAL_NUM").Header.SetVisiblePosition(8, False)
        grdEDT830T1X.DisplayLayout.Bands(0).Columns("TOTAL_SALES").Header.SetVisiblePosition(9, False)

        spldata.Panel1Collapsed = True
        spldata.Panel2Collapsed = False
        ASCMAIN1.Progress("", String.Empty)

    End Sub

    ''' <summary>
    ''' Sets up screen based on the form modality, state and type of processing
    ''' </summary>
    ''' <param name="tf"></param>
    ''' <remarks></remarks>
    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_Description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        If ScreenMode Then
        Else
            ClearRecord()
        End If
    End Sub

    ''' <summary>
    ''' Validates data when a user selects a menu option
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty
        Dim sql As String = String.Empty

        Select Case eItemKey

            Case "Load"
                If MessageBox.Show($"Do you want to load 830 Entry {EDI_DOC_SEQ_NO}?", "Load", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

            Case "Cancel"

        End Select

        If EMsg <> String.Empty Then
            MessageBox.Show(EMsg, "Cannot Proceed", MessageBoxButtons.OK)
        Else
            Call Proceed(eItemKey)
        End If

    End Sub

    ''' <summary>
    ''' When the user selects a menu option perform the action
    ''' </summary>
    ''' <param name="eItemKey"></param>
    ''' <remarks></remarks>
    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                LoadRecord()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    ''' <summary>
    ''' Updates data based on the current state of the screen and the type of processing
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub UpdateRecord()

        Try
            BeginTrans()
            CommitTrans("Update Successful")
        Catch ex As Exception
            Rollback(ex.Message)

        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT830T1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdEDT830T1X, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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
                Case "grdARTCBDA1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case ""

            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case ""
        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdEDT830T1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdEDT830T1.DoubleClickRow

        If e.Row.IsDataRow Then
            EDI_DOC_SEQ_NO = e.Row.Cells("EDI_DOC_SEQ_NO").Text
            Click_Command("Load")
        End If

    End Sub

    Private Sub grdEDT830T1X_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdEDT830T1X.InitializeRow
        If tblItems IsNot Nothing Then
            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Text
            If tblItems.Rows.Find(ITEM_CODE) Is Nothing Then
                e.Row.Cells("ITEM_CODE").Appearance.BackColor = Drawing.Color.Yellow
                e.Row.Cells("ITEM_CODE").ToolTipText = "New Item"
            End If
        End If
    End Sub

    Private Sub grdEDT830T1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDT830T1.AfterRowActivate
        If grdEDT830T1.ActiveRow Is Nothing Then
            txtRawEDI.Clear()
            Exit Sub
        End If

        If Not grdEDT830T1.ActiveRow.IsDataRow Then
            txtRawEDI.Clear()
            Exit Sub
        End If

        Dim EDI_DOC_SEQ_NO As String = grdEDT830T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Text
        txtRawEDI.Text = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, "", "830")

    End Sub

#End Region

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub


End Class