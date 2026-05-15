Imports Infragistics.Win.UltraWinGrid

Public Class SOFSHIPT

    Private CUST_CODE As String = String.Empty
    Private ORDR_CUST_PO As String = String.Empty
    Private clsWHTSHIP1 As New TAC.WHCSHIP1
    Private podPath As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'TRACKING_STATUS
        '   S Shipped
        '   D Delivered
        '   L Lost(user will mark As lost)
        '   E Expired(more than 120 days without a status Of delivered)

        '   Optset B is for SHIP_BOL_NO
        '   OptSet U is for UPS Ship Date

        InquiryMode = MENU_ITEM_OBJECT = "SOFSHIPI"

        Get_PARM("ASTPARM1")

        With dst
            Create_TDA(.Tables.Add, "SOTSHIPT", "*")
            With .Tables("SOTSHIPT")
                .Columns.Add("CUST_CODE", GetType(String))
                .Columns.Add("CUST_NAME", GetType(String))
                .Columns.Add("ORDR_GROUP_NO", GetType(String))
                .Columns.Add("ORDR_CUST_PO", GetType(String))
                .Columns.Add("ORDR_NO", GetType(String))
                .Columns.Add("INV_NO", GetType(String))
                .Columns.Add("INV_DATE", GetType(Date))
                .Columns.Add("SHIP_TO_NAME", GetType(String))
            End With

            dst.Tables.Add("SOTSHIPX")
            With .Tables("SOTSHIPX")
                .Columns.Add("SHIP_BOL_NO", GetType(String))
                .Columns.Add("CUST_CODE", GetType(String))
                .Columns.Add("CUST_NAME", GetType(String))
                .Columns.Add("ORDR_GROUP_NO", GetType(String))
                .Columns.Add("ORDR_CUST_PO", GetType(String))
                .Columns.Add("NUM_PICK_NOS", GetType(Int16))
            End With

            ASCMAIN1.sql = "SELECT EDT945T1.EDI_PICK_NO PICK_NO, EDT945T2.EDI_SHIPPER_ID_NO TRACKING_NO, EDT945T2.STYLE_CODE, ICTITEM1.ITEM_DESC, EDT945T2.PICK_QTY,
                                SUM(EDT945T2.EDI_SHIP_QTY) EDI_SHIP_QTY
                                FROM EDT945T1, EDT945T2, ICTITEM1
                                WHERE EDT945T1.EDI_DOC_SEQ_NO = EDT945T2.EDI_DOC_SEQ_NO
                                AND EDI_PICK_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))
                                AND EDT945T2.STYLE_CODE = ICTITEM1.ITEM_CODE (+)
                                GROUP BY EDT945T1.EDI_PICK_NO, EDT945T2.EDI_SHIPPER_ID_NO , EDT945T2.STYLE_CODE, ICTITEM1.ITEM_DESC, EDT945T2.PICK_QTY"
            Create_TDA(.Tables.Add, "EDT945TX", ASCMAIN1.sql, 0, False, "V", 0)

            Create_TDA(.Tables.Add, "TATEVNT1", "*")
            Create_TDA(.Tables.Add, "SOTCARR3", "*")

            Fill_Records("SOTCARR3", String.Empty, True, "SELECT * FROM SOTCARR3 WHERE SERVER_TOKEN_URL IS NOT NULL AND CLIENT_ID IS NOT NULL AND CLIENT_SECRET IS NOT NULL")
        End With

        Create_Relation("SOTSHIPX", "SOTSHIPT", "SHIP_BOL_NO")
        With dst.Tables("SOTSHIPX")
            .Columns.Add("NUM_PACKAGES", GetType(Int16), "COUNT(CHILD.TRACKING_NO)")
            .Columns.Add("SHIP_DATE", GetType(Date), "MIN(CHILD.SHIP_DATE)")
            .Columns.Add("EXPECTED_DELIVERY_DATE", GetType(Date), "MAX(CHILD.EXPECTED_DELIVERY_DATE)")
            .Columns.Add("EMAIL_OPER", GetType(String), "MAX(CHILD.EMAIL_OPER)")
            .Columns.Add("EMAIL_DATE", GetType(Date), "MAX(CHILD.EMAIL_DATE)")
            .Columns.Add("SHIP_TO_NAME", GetType(String), "MAX(CHILD.SHIP_TO_NAME)")
        End With

        Create_Relation("SOTSHIPT", "EDT945TX", "PICK_NO,TRACKING_NO")

        grdSOTSHIPT.DataSource = dst.Tables("SOTSHIPX")
        ASCMAIN1.Add_Value_List(grdSOTSHIPT, "TRACKING_STATUS", , New String() {":", "S:Shipped", "D:Delivered", "L:Lost", "E:Expired", "W:Warning"}, 1)
        Create_Summary(grdSOTSHIPT, "SHIP_BOL_NO", "Count")

        dteStart.MaxDate = DateAdd(DateInterval.Day, 1, DateTime.Now)
        dteStart.MinDate = CDate("03/01/2025")
        dteStart.DateTime = DateTime.Now.AddDays(-5).ToShortDateString

        dteEnd.MaxDate = dteStart.MaxDate
        dteEnd.MinDate = dteStart.MinDate
        dteEnd.DateTime = DateTime.Now.ToShortDateString

        If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "edz" Then
            Stop
            podPath = IO.Path.Combine(ASCMAIN1.Folders("Archive"), "POD")
        Else
            podPath = IO.Path.Combine(ROWs("ASTPARM1").Item("AS_PARM_ARCHIVE_FOLDER"), "POD")
        End If

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                CUST_CODE = Absx1.txtFor("CUST_CODE").Text.Trim
                If CUST_CODE.Length > 0 Then
                    Validate_Code("CUST_CODE", False, True)
                End If

                ORDR_CUST_PO = txtCustomerPO.Text.Trim
                If ORDR_CUST_PO.Length > 0 AndAlso CUST_CODE.Length = 0 Then
                    EMsg &= vbCr & "Customer is required when providing a Customer PO."
                    Exit Select
                End If

                Select Case optStatus.Value
                    Case "S", "L", "E", "W"
                        ' Nothing to do

                    Case "D"
                        If DateDiff(DateInterval.Day, CDate(dteStart.DateTime.ToShortDateString), CDate(dteEnd.DateTime.ToShortDateString)) < 0 Then
                            EMsg &= vbCr & "End Date must be greater/equal Start Date"
                        End If
                End Select

            Case "Done"

            Case "Email POD"
                Dim SHIP_BOL_NO As String = grdSOTSHIPT.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim NUM_PACKAGES As Int16 = dst.Tables("SOTSHIPT").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' AND TRACKING_STATUS = 'D'").Count
                If NUM_PACKAGES = 0 Then
                    MessageBox.Show("The are no cartons with POD information.", "Email POD", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim zMsg As String = "Do you want to send POD email for:" & Environment.NewLine
                zMsg &= $"Shipment: {grdSOTSHIPT.ActiveRow.Cells("SHIP_BOL_NO").Value}" & Environment.NewLine
                zMsg &= $"Customer: {grdSOTSHIPT.ActiveRow.Cells("CUST_CODE").Value} {grdSOTSHIPT.ActiveRow.Cells("CUST_NAME").Value}" & Environment.NewLine
                zMsg &= $"P.O.: {grdSOTSHIPT.ActiveRow.Cells("ORDR_CUST_PO").Value}" & Environment.NewLine
                zMsg &= $"Order Group: {grdSOTSHIPT.ActiveRow.Cells("ORDR_GROUP_NO").Value}" & Environment.NewLine

                Select Case grdSOTSHIPT.ActiveRow.Band.Index
                    Case 0
                        zMsg &= $"Number Packages: {grdSOTSHIPT.ActiveRow.Cells("NUM_PACKAGES").Value}?"

                    Case 1
                        zMsg &= $"Tracking Number: {grdSOTSHIPT.ActiveRow.Cells("TRACKING_NO").Value}?"
                End Select

                If MessageBox.Show(zMsg, "Email POD", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Email POD"
                EmailPOD()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Email POD").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Email POD").Visible = Not InquiryMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        splPOD.Visible = ScreenMode
        Set_Read_Only(ebcOptions, ScreenMode)

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SOTSHIPT", "SOTSHIPX", "EDT945TX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        wbPOD.DocumentText = String.Empty

        txtCustomerPO.Clear()
        txtCUST_CODE.Clear()

        Clear_All_Filters(grdSOTSHIPT)

        If InquiryMode Then
            txtCUST_CODE.Text = "IPLBAE"
            txtCUST_CODE.Enabled = False
            Show_Filter(grdSOTSHIPT)
        End If
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        Dim DATE_S As String = String.Empty
        Dim DATE_E As String = String.Empty

        ASCMAIN1.sql = "SELECT SOTSHIPT.*, SOTORDR0.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO, ARTCUST1.CUST_NAME
                            , SOTPICK1.ORDR_NO, SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTORDR5.CUST_NAME SHIP_TO_NAME
                            From SOTSHIPT, SOTSHIP1, SOTORDR0, ARTCUST1, SOTPICK1, SOTINVH1, 
                            (SELECT ORDR_NO, CUST_NAME FROM SOTORDR5 WHERE CUST_ADDR_TYPE = 'ST') SOTORDR5
                            Where SOTSHIPT.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO
                            And SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO (+)
                            And SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE (+)
                            AND SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO (+)
                            AND SOTSHIPT.PICK_NO = SOTPICK1.PICK_NO (+)
                            AND SOTPICK1.INV_NO = SOTINVH1.INV_NO (+)
                            AND SOTPICK1.ORDR_NO = SOTORDR5.ORDR_NO (+)"

        If CUST_CODE.Length > 0 Then
            ASCMAIN1.sql &= $" And SOTORDR0.CUST_CODE = '{CUST_CODE}'"
            If ORDR_CUST_PO.Length > 0 Then
                ASCMAIN1.sql &= $" AND SOTORDR0.ORDR_CUST_PO = '{txtCustomerPO.Text.Trim}'"
            End If
        End If

        ' Customer PO gets all enteies for the Customer / PO 
        If ORDR_CUST_PO.Length = 0 Then
            ASCMAIN1.sql &= " AND SOTSHIPT.SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM SOTSHIPT "

            Select Case optStatus.Value
                Case "S", "L", "E"
                    ASCMAIN1.sql &= $" WHERE SOTSHIPT.TRACKING_STATUS = '{optStatus.Value}'"

                Case "D"
                    DATE_S = dteStart.DateTime.ToString("dd-MMM-yyyy")
                    DATE_E = dteEnd.DateTime.ToString("dd-MMM-yyyy")
                    ASCMAIN1.sql &= $" WHERE SOTSHIPT.TRACKING_STATUS = '{optStatus.Value}' and SOTSHIPT.DELIVERY_DATE BETWEEN '{DATE_S}' AND '{DATE_E}'"
            End Select

            ASCMAIN1.sql &= ")"
        End If

        Fill_Records("SOTSHIPT", String.Empty, True, ASCMAIN1.sql)

        Dim tblSOTSHIPx As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTSHIPT"), {"SHIP_BOL_NO", "CUST_CODE", "CUST_NAME", "ORDR_GROUP_NO", "ORDR_CUST_PO"})
        Dim tblPickNos As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTSHIPT"), {"SHIP_BOL_NO", "PICK_NO"})
        For Each rowSOTSHIPx As DataRow In tblSOTSHIPx.Select("")
            dst.Tables("SOTSHIPX").ImportRow(rowSOTSHIPx)
            Dim SHIP_BOL_NO As String = rowSOTSHIPx.Item("SHIP_BOL_NO") & String.Empty
            dst.Tables("SOTSHIPX").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}'")(0).Item("NUM_PICK_NOS") = tblPickNos.Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}'").Length
        Next

        Dim lstPickNos As New List(Of String)
        For Each rowSOTSHIPT As DataRow In dst.Tables("SOTSHIPT").Select("")
            Dim PICK_NO As String = rowSOTSHIPT.Item("PICK_NO") & ""
            If Not lstPickNos.Contains(PICK_NO) Then
                lstPickNos.Add(PICK_NO)
            End If
        Next
        ' Needed this to avoid error
        dst.Tables("EDT945TX").Rows.Clear()
        Dim lstRetrieve As New List(Of String)
        For Each PICK_NO As String In lstPickNos
            lstRetrieve.Add(PICK_NO)
            If lstRetrieve.Count >= 500 Then
                Fill_Records("EDT945TX", String.Join(",", lstRetrieve.ToArray), False)
                lstRetrieve.Clear()
            End If
        Next

        If lstRetrieve.Count > 0 Then
            Fill_Records("EDT945TX", String.Join(",", lstRetrieve.ToArray), False)
            lstRetrieve.Clear()
        End If

        For Each rowEDT945TX As DataRow In dst.Tables("EDT945TX").Select("")
            Dim PICK_NO As String = rowEDT945TX.Item("PICK_NO") & ""
            Dim TRACKING_NO As String = rowEDT945TX.Item("TRACKING_NO") & ""

            If dst.Tables("SOTSHIPT").Select($"PICK_NO = '{PICK_NO}' and TRACKING_NO = '{TRACKING_NO}'").Length = 0 Then
                rowEDT945TX.Delete()
            End If
        Next
        dst.Tables("EDT945TX").AcceptChanges()

        EnforceConstraints(True)
        Sort_grdColumns(grdSOTSHIPT, "SHIP_BOL_NO,CUST_CODE,ORDR_GROUP_NO")
        Sort_grdColumns(grdSOTSHIPT, "SHIP_BOL_NO,CUST_CODE,ORDR_GROUP_NO,TRACKING_NO",, 1)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPT, "SSSPBPBBPBBPBBPB", "Show Filter", "Show GroupBox", "Show Pins", "Auto Fit Columns", "Reset to Shipped", "Mark as Delivered", "Get Shipment POD", "Get Package POD", "View Shipment POD", "View Package POD", "Customer Order Inquiry")
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
                Case grdSOTSHIPT.Name
                    tlb_btn = DirectCast(tlb_pop.Tools("Reset to Shipped"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode AndAlso grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key AndAlso Not "SD".Contains(grd.ActiveRow.Cells("TRACKING_STATUS").Value & String.Empty)

                    tlb_btn = DirectCast(tlb_pop.Tools("Mark as Delivered"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode AndAlso grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key AndAlso Not "D".Contains(grd.ActiveRow.Cells("TRACKING_STATUS").Value & String.Empty)

                    tlb_btn = DirectCast(tlb_pop.Tools("View Shipment POD"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(0).Key AndAlso dst.Tables("SOTSHIPT").Select($"SHIP_BOL_NO = '{grd.ActiveRow.Cells("SHIP_BOL_NO").Value}' AND ISNULL(POD, '') <> ''").Length > 0

                    tlb_btn = DirectCast(tlb_pop.Tools("View Package POD"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key AndAlso grd.ActiveRow.Cells("POD").Value & String.Empty <> String.Empty

                    tlb_btn = DirectCast(tlb_pop.Tools("Get Shipment POD"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(0).Key AndAlso dst.Tables("SOTSHIPT").Select($"SHIP_BOL_NO = '{grd.ActiveRow.Cells("SHIP_BOL_NO").Value}' AND ISNULL(POD, '') = ''").Length > 0

                    tlb_btn = DirectCast(tlb_pop.Tools("Get Package POD"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key AndAlso grd.ActiveRow.Cells("POD").Value & String.Empty = String.Empty

                    tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key
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
            Case "Auto Fit Columns"
                Me.Cursor = Cursors.WaitCursor
                grd.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                Me.Cursor = Cursors.Default

            Case "Reset to Shipped"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim TRACKING_NO As String = grd.ActiveRow.Cells("TRACKING_NO").Value

                Try
                    BeginTrans()
                    ASCMAIN1.sql = "UPDATE SOTSHIPT SET TRACKING_STATUS = 'S' WHERE SHIP_BOL_NO = :PARM1 AND TRACKING_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {SHIP_BOL_NO, TRACKING_NO})
                    CommitTrans()

                    dst.Tables("SOTSHIPT").Rows.Find({SHIP_BOL_NO, TRACKING_NO}).Item("TRACKING_STATUS") = "S"

                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

            Case "Mark as Delivered"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim TRACKING_NO As String = grd.ActiveRow.Cells("TRACKING_NO").Value
                Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value

                Try
                    BeginTrans()
                    ASCMAIN1.sql = "UPDATE SOTSHIPT SET TRACKING_STATUS = 'D' WHERE SHIP_BOL_NO = :PARM1 AND TRACKING_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {SHIP_BOL_NO, TRACKING_NO})

                    ASCMAIN1.sql = "UPDATE SOTINVH1 SET ORDR_YYYYPP_DEL = :PARM1 WHERE INV_TYPE = 'I' AND INV_NO = (SELECT INV_NO FROM SOTPICK1 WHERE PICK_NO = :PARM2)"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {ASCMAIN1.CYP, PICK_NO})

                    CommitTrans()

                    dst.Tables("SOTSHIPT").Rows.Find({SHIP_BOL_NO, TRACKING_NO}).Item("TRACKING_STATUS") = "D"

                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

            Case "View Shipment POD"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                ViewShipmentPOD(SHIP_BOL_NO, "", PODTypes.Shipment)

            Case "View Package POD"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim TRACKING_NO As String = grd.ActiveRow.Cells("TRACKING_NO").Value
                ViewShipmentPOD(SHIP_BOL_NO, TRACKING_NO, PODTypes.TrackingNumber)

            Case "Get Shipment POD"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                ' Dim TRACKING_NO As String = grd.ActiveRow.Cells("TRACKING_NO").Value
                GetShipmentPOD(SHIP_BOL_NO, "", PODTypes.Shipment)
                ViewShipmentPOD(SHIP_BOL_NO, "", PODTypes.Shipment)
                grdSOTSHIPT.Rows.Refresh(RefreshRow.FireInitializeRow, True)

            Case "Get Package POD"
                Dim SHIP_BOL_NO As String = grd.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim TRACKING_NO As String = grd.ActiveRow.Cells("TRACKING_NO").Value
                GetShipmentPOD(SHIP_BOL_NO, TRACKING_NO, PODTypes.TrackingNumber)
                ViewShipmentPOD(SHIP_BOL_NO, TRACKING_NO, PODTypes.TrackingNumber)
                grdSOTSHIPT.Rows.Refresh(RefreshRow.FireInitializeRow, True)

            Case "Customer Order Inquiry"
                Dim ORDR_GROUP_NO As String = grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & ""
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Context_Launch("Select", CUST_CODE & ":" & ORDR_GROUP_NO, e.Tool.Key, "SOFCORD1")
        End Select
    End Sub


#End Region

#Region "ABSColumn Controls"

    Private Sub optStatus_ValueChanged(sender As Object, e As EventArgs) Handles optStatus.ValueChanged

        dteStart.Enabled = False
        dteEnd.Enabled = False

        Select Case optStatus.Value
            Case "S", "L", "E"

            Case "D"
                dteStart.Enabled = True
                dteEnd.Enabled = True
                dteStart.Focus()

        End Select

    End Sub

#End Region

#Region "Form Procedures"

    Private Enum PODTypes
        Shipment
        TrackingNumber
    End Enum

    Private Sub ViewShipmentPOD(ByVal SHIP_BOL_NO As String, ByVal TRACKING_NO As String, ByVal PODType As PODTypes)

        Try
            Dim tblSOTSHIPT As DataTable = Nothing
            Dim podData As String = String.Empty
            wbPOD.DocumentText = String.Empty

            Select Case PODType
                Case PODTypes.Shipment
                    ASCMAIN1.sql = "SELECT * FROM SOTSHIPT WHERE SHIP_BOL_NO = :PARM1"
                    tblSOTSHIPT = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSHIPT", "V", SHIP_BOL_NO)

                Case PODTypes.TrackingNumber
                    ASCMAIN1.sql = "SELECT * FROM SOTSHIPT WHERE SHIP_BOL_NO = :PARM1 AND TRACKING_NO = :PARM2"
                    tblSOTSHIPT = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSHIPT", "VV", {SHIP_BOL_NO, TRACKING_NO})
            End Select

            If tblSOTSHIPT.Rows.Count = 0 Then
                MessageBox.Show("Cannot locate the desired data.", "View Shipment POD", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            For Each rowSOTSHIPT As DataRow In tblSOTSHIPT.Select("", "TRACKING_NO")
                TRACKING_NO = rowSOTSHIPT.Item("TRACKING_NO") & String.Empty
                If My.Computer.FileSystem.FileExists(IO.Path.Combine(podPath, TRACKING_NO & ".html")) Then
                    Using sr As New IO.StreamReader(IO.Path.Combine(podPath, TRACKING_NO & ".html"))
                        podData &= sr.ReadToEnd & Environment.NewLine
                        sr.Close()
                        sr.Dispose()
                    End Using
                End If
            Next

            wbPOD.DocumentText = podData

        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "View Shipment POD", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub GetShipmentPOD(ByVal SHIP_BOL_NO As String, ByVal TRACKING_NO As String, ByVal PODType As PODTypes)
        Try
            Me.Cursor = Cursors.WaitCursor
            Dim tbl As DataTable = Nothing

            Select Case PODType
                Case PODTypes.Shipment
                    ASCMAIN1.sql = "SELECT * FROM SOTSHIPT WHERE SHIP_BOL_NO = :PARM1 AND POD IS NULL"
                    tbl = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSHIPT", "V", SHIP_BOL_NO)

                Case PODTypes.TrackingNumber
                    ASCMAIN1.sql = "SELECT * FROM SOTSHIPT WHERE SHIP_BOL_NO = :PARM1 AND TRACKING_NO = :PARM2 AND POD IS NULL"
                    tbl = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSHIPT", "VV", {SHIP_BOL_NO, TRACKING_NO})
            End Select

            If tbl.Rows.Count = 0 Then
                MessageBox.Show("PODs already exist.", "Get Shipment POD", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            ASCMAIN1.Progress("Get Shipment POD", "")

            For Each rowSOTSHIPT As DataRow In tbl.Select("", "TRACKING_NO")
                TRACKING_NO = rowSOTSHIPT.Item("TRACKING_NO") & String.Empty
                SHIP_BOL_NO = rowSOTSHIPT.Item("SHIP_BOL_NO") & String.Empty

                ASCMAIN1.Progress("-", TRACKING_NO)

                Select Case rowSOTSHIPT.Item("CARRIER_CODE") & String.Empty

                    Case "UPS"
                        If dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'UPS'").Length = 0 Then
                            Continue For
                        End If
                        Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = 'UPS'")(0)

                        Dim TRACKINFO As New TAC.WHCSHIP1.TRACKINFO
                        Dim PackagePODFile As String = System.IO.Path.Combine(podPath, TRACKING_NO & ".html")

                        If clsWHTSHIP1.UPSProofOfDelivery(TRACKING_NO, rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty, TRACKINFO, PackagePODFile) Then
                            Dim rowSOTSHIPT_grd As DataRow = dst.Tables("SOTSHIPT").Rows.Find({SHIP_BOL_NO, TRACKING_NO})
                            If IsDate(TRACKINFO.DELIVERY_DATE) Then
                                TRACKINFO.DELIVERY_DATE = CDate(TRACKINFO.DELIVERY_DATE).ToString("dd-MMM-yyyy HH:mm:ss")
                            Else
                                TRACKINFO.DELIVERY_DATE = ""
                            End If

                            If IsDate(TRACKINFO.EXPECTED_DELIVERY_DATE) Then
                                TRACKINFO.EXPECTED_DELIVERY_DATE = CDate(TRACKINFO.EXPECTED_DELIVERY_DATE).ToString("dd-MMM-yyyy")
                            Else
                                TRACKINFO.EXPECTED_DELIVERY_DATE = ""
                            End If

                            If rowSOTSHIPT_grd IsNot Nothing Then

                                If IsDate(TRACKINFO.DELIVERY_DATE) Then
                                    rowSOTSHIPT_grd.Item("DELIVERY_DATE") = TRACKINFO.DELIVERY_DATE
                                    rowSOTSHIPT_grd.Item("TRACKING_STATUS") = "D"
                                End If

                                If TRACKINFO.PACKAGE_POD.Length > 0 Then
                                    If My.Computer.FileSystem.FileExists(TRACKINFO.PACKAGE_POD) Then
                                        rowSOTSHIPT_grd.Item("POD") = TRACKINFO.PACKAGE_POD
                                    End If
                                End If

                                rowSOTSHIPT_grd.Item("PACKAGE_WEIGHT") = TRACKINFO.PACKAGE_WEIGHT
                                rowSOTSHIPT_grd.Item("SERVICE_DESC") = TRACKINFO.SERVICE_DESC

                                rowSOTSHIPT_grd.Item("RECEIVED_BY") = TRACKINFO.RECEIVED_BY
                                rowSOTSHIPT_grd.Item("DELIVERY_CITY") = TRACKINFO.DELIVERY_CITY
                                rowSOTSHIPT_grd.Item("DELIVERY_STATE") = TRACKINFO.DELIVERY_STATE
                                rowSOTSHIPT.Item("DELIVERY_ZIPCODE") = TRACKINFO.DELIVERY_ZIPCODE

                                rowSOTSHIPT_grd.Item("DELIVERY_LOCATION") = TRACKINFO.DELIVERY_LOCATION

                                If IsDate(TRACKINFO.EXPECTED_DELIVERY_DATE) Then
                                    rowSOTSHIPT_grd.Item("EXPECTED_DELIVERY_DATE") = TRACKINFO.EXPECTED_DELIVERY_DATE
                                End If

                                rowSOTSHIPT_grd.Item("LAST_PROCESSED") = DateTime.Now
                            End If

                        ElseIf clsWHTSHIP1.LastError.Length > 0 Then
                            MessageBox.Show(clsWHTSHIP1.LastError, "Get Shipment POD", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If

                        Try
                            BeginTrans()
                            Update_Record_TDA("SOTSHIPT")

                            'Dim rowSOTSHIPT_LK As DataRow = dst.Tables("SOTSHIPT").Rows.Find({SHIP_BOL_NO, TRACKING_NO})
                            'If rowSOTSHIPT_LK IsNot Nothing AndAlso rowSOTSHIPT_LK.Item("TRACKING_STATUS") = "D" Then
                            '    Dim PICK_NO As String = rowSOTSHIPT_LK.Item("PICK_NO") & String.Empty
                            '    Dim INV_NO As String = rowSOTSHIPT_LK.Item("INV_NO") & String.Empty
                            '    If INV_NO.Length = 0 Then
                            '        Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", {PICK_NO})
                            '        If rowSOTPICK1 IsNot Nothing Then
                            '            INV_NO = rowSOTPICK1.Item("INV_NO") & String.Empty
                            '        End If
                            '    End If

                            '    ASCMAIN1.sql = "UPDATE SOTCART1 SET DELIVERY_DATE = :PARM1 WHERE PICK_NO = :PARM2 AND CART_TRACKING_NO = :PARM3"
                            '    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVV", {CDate(rowSOTSHIPT_LK.Item("DELIVERY_DATE") & String.Empty), PICK_NO, TRACKING_NO})

                            '    ASCMAIN1.sql = "UPDATE SOTINVH1 SET DELIVERY_DATE = :PARM1 WHERE INV_TYPE = :PARM2 AND INV_NO = :PARM3 AND DELIVERY_DATE IS NULL"
                            '    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVV", {CDate(rowSOTSHIPT_LK.Item("DELIVERY_DATE") & String.Empty), "I", INV_NO})
                            'End If

                            CommitTrans()
                        Catch ex As Exception
                            Rollback(ex.Message)
                        End Try

                    Case "FEDEX"

                End Select
            Next

        Catch ex As Exception
            MessageBox.Show($"Error: {ex.Message}", "Get Shipment POD", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Private Sub EmailPOD()

        Try
            ASCMAIN1.Progress("Emailing POD")
            Dim dictOrderNos As New Dictionary(Of String, String)

            Select Case grdSOTSHIPT.ActiveRow.Band.Index
                Case 0
                    ' Nothing now, maybe show all Shipments for the Shipment
                    Dim SHIP_BOL_NO As String = grdSOTSHIPT.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty
                    For Each row As DataRow In dst.Tables("SOTSHIPT").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' and TRACKING_STATUS = 'D'", "TRACKING_NO")
                        If row.Item("POD") & String.Empty <> String.Empty Then
                            dictOrderNos.Add(row.Item("TRACKING_NO") & String.Empty, row.Item("ORDR_NO") & String.Empty)
                        End If
                    Next

                Case 1
                    If grdSOTSHIPT.ActiveRow.Cells("ORDR_NO").Value & String.Empty <> String.Empty AndAlso grdSOTSHIPT.ActiveRow.Cells("POD").Value & String.Empty <> String.Empty Then
                        dictOrderNos.Add(grdSOTSHIPT.ActiveRow.Cells("TRACKING_NO").Value & String.Empty, grdSOTSHIPT.ActiveRow.Cells("ORDR_NO").Value & String.Empty)
                    End If

            End Select

            If dictOrderNos.Count = 0 Then
                MessageBox.Show("There are no Shipments marked as Delivered.", "Email POD", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Dim POD_CONTENT As String = String.Empty
            For Each kvp As KeyValuePair(Of String, String) In dictOrderNos
                Dim TRACKING_NO As String = kvp.Key
                If My.Computer.FileSystem.FileExists(IO.Path.Combine(podPath, TRACKING_NO & ".HTML")) Then
                    Using sr As New IO.StreamReader(IO.Path.Combine(podPath, TRACKING_NO & ".HTML"))
                        POD_CONTENT &= sr.ReadToEnd
                        sr.Close()
                        sr.Dispose()
                    End Using
                End If
            Next

            If POD_CONTENT.Length = 0 Then
                MessageBox.Show("Cannot locate the Proof of Delivery data for the requested shipments.", "Email POD", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End If

            Using frmTAFSEND1 As New TAFSEND1(Me)
                With frmTAFSEND1
                    .EMAIL_KEY = ""

                    .SEND_FROM = ASCMAIN1.USER_EMAIL  ' "donotreply" & "@" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_DEFAULT_EMAIL_DOMAIN")  
                    .SEND_TOs = New Dictionary(Of String, String)
                    .SEND_TOs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)

                    .SEND_FROM_NAME = ASCMAIN1.USER_NAME
                    .SEND_TO = ""
                    .SEND_TO_NAME = ""

                    If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "edz" Then
                        .SEND_FROM = "nferon@interparfums.com"
                        .SEND_TOs = New Dictionary(Of String, String)
                        .SEND_TOs.Add("ewz@absolution.com", "Ed Zenker")
                    End If

                    .SEND_SUBJECT = "Proof Of Delivery"
                    .SEND_BODY = POD_CONTENT
                    .SEND_BODY_OVERRIDE = POD_CONTENT
                    '.SEND_ATTACHMENT = ""
                    .SEND_METHOD = "E"
                    .SEND_ENTITY_CAPTION = "Customer"
                    .SEND_ENTITY_TABLE = "ARTCUST1"

                    Me.Cursor = Cursors.WaitCursor

                    .ShowDialog()

                    Try
                        If .SEND_STATUS <> "C" Then
                            For Each kvp As KeyValuePair(Of String, String) In dictOrderNos
                                Dim TRACKING_NO As String = kvp.Key
                                Dim ORDR_NO As String = kvp.Value

                                dst.Tables("TATEVNT1").Rows.Add _
                                    (New Object() {"SOTORDR1", ORDR_NO,
                                    Now + ASCMAIN1.NowTSD, ASCMAIN1.USER_ID,
                                    "POD",
                                    $"Proof of Delivery Sent Tracking No: {TRACKING_NO}",
                                    .SEND_NO,
                                    "SOTSHIPT"})

                                Update_Record_TDA("TATEVNT1")

                                ASCMAIN1.sql = $"TRACKING_NO = '{TRACKING_NO}' AND ORDR_NO = '{ORDR_NO}'"
                                For Each rowSOTSHIPT As DataRow In dst.Tables("SOTSHIPT").Select(ASCMAIN1.sql)
                                    rowSOTSHIPT.Item("EMAIL_OPER") = ASCMAIN1.USER_ID
                                    rowSOTSHIPT.Item("EMAIL_DATE") = DateTime.Now
                                Next

                                Update_Record_TDA("SOTSHIPT")
                            Next
                        End If
                    Catch ex As Exception

                    Finally
                        dst.Tables("TATEVNT1").Rows.Clear()
                    End Try
                End With

                frmTAFSEND1.Dispose()
            End Using

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Send POD Email", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Private Sub grdSOTSHIPT_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTSHIPT.InitializeRow

        Select Case e.Row.Band.Key
            Case grdSOTSHIPT.DisplayLayout.Bands(0).Key, grdSOTSHIPT.DisplayLayout.Bands(1).Key
                Dim SHIP_BOL_NO As String = e.Row.Cells("SHIP_BOL_NO").Value & String.Empty
                If dst.Tables("SOTSHIPT").Select($"SHIP_BOL_NO = '{SHIP_BOL_NO}' AND TRACKING_STATUS = 'D'").Length > 0 Then
                    If optStatus.Value <> "D" AndAlso e.Row.Band.Key = grdSOTSHIPT.DisplayLayout.Bands(0).Key Then
                        e.Row.Cells("SHIP_BOL_NO").Appearance.BackColor = Drawing.Color.LightGreen
                    ElseIf optStatus.Value <> "D" AndAlso e.Row.Band.Key = grdSOTSHIPT.DisplayLayout.Bands(1).Key Then
                        e.Row.ParentRow.Cells("SHIP_BOL_NO").Appearance.BackColor = Drawing.Color.LightGreen
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "ORDR_CUST_PO"

                If Absx1.txtFor("CUST_CODE").TextLength = 0 Then
                    MessageBox.Show("You must provide a Customer Code", "Select Customer PO", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = $" CUST_CODE = '{Absx1.txtFor("CUST_CODE").Text}'"
        End Select
    End Sub

    Private Sub grdSOTSHIPT_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTSHIPT.AfterRowActivate
        wbPOD.DocumentText = String.Empty
    End Sub

#End Region

End Class