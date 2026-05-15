Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid
Imports System.Net

Public Class WHFSPCK1
    Dim WHSE_CODE As String = String.Empty
    Dim LP_CODE As String = String.Empty
    Dim rowICTWHSE1 As DataRow = Nothing
    Dim rowWHTTPLP1 As DataRow = Nothing

    Dim WHTSTYLX As String = String.Empty

    ' NEED TO MT AGAINST REL & DE-REL FOR THE WHSE SELECTED
    ' NEED TO MT FOR SENDITEMS
    ' DO NOT ALLOW PT PRINT FOR A LP WHSE
    ' CHECK MT ON DESIGN RECALL OF SHIPMENTS/PICK TICKETS
    ' CHECK EVENT PROCEDURE FIRING WHEN CLICKING CANCEL - WHEN CYCLING MODES TO FALSE, DONT WANT TO LOAD_SOTSHIPX

    Dim SOTSHIPX As String = String.Empty
    Dim Shipments As Integer = 0
    Dim LP_XNO As String = String.Empty
    Dim ASW As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            ASCMAIN1.sql = "Select SHIP_BOL_NO, '1' SEL, '1' EDI856, '1' SHIP_CART_REQD from SOTSHIP1 where ROWNUM < 1"
            SOTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add ORDR_ADDR_TYPE_ST VARCHAR2(2)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add CUST_CODE VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("CREATE INDEX I_" & SOTSHIPX & "_1 ON " & SOTSHIPX & " (ORDR_NO)")

            ASCMAIN1.sql = "Select SOTSHIP1.*, SOTSHIPX.SEL, SOTSHIPX.EDI856" & vbCrLf _
                & ", SOTSHIPX.ORDR_NO, SOTSHIP1.SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST" & vbCrLf _
                & ", WHTSHIPX.LP_XNO LP_XNO_XMIT" & vbCrLf _
                & ", SOTORDR0.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
                & ", SOTORDR0.ORDR_CUST_PO, SOTORDR0.ORDR_DATE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" & vbCrLf _
                & " from " & SOTSHIPX & " SOTSHIPX, SOTSHIP1, SOTORDR0, ARTCUST1, WHTSHIPX" & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
                & "   and WHTSHIPX.LP_XNO (+) = SOTSHIP1.LP_XNO" & vbCrLf _
                & "   and WHTSHIPX.SHIP_BOL_NO (+) = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_STATUS = 'P'"
            Create_TDA(.Tables.Add("SOTSHIPX"), SOTSHIPX, "**", 0, True, "", 1, "SEL")
            '.Tables("SOTSHIPX").Columns.Add("SEL")

            Create_TDA(.Tables.Add, "WHTSHIPX", "*")

            ASCMAIN1.sql = "Select ICTWHSE1.WHSE_CODE, ICTWHSE1.WHSE_DESC, ICTWHSE1.LP_CODE" & vbCrLf _
                & ", X.SHIPS" & vbCrLf _
                & "  from ICTWHSE1" & vbCrLf _
                & ", (Select SOTSHIP1.WHSE_CODE, Count (*) SHIPS" & vbCrLf _
                & "  from SOTSHIP1" & vbCrLf _
                & "where SOTSHIP1.SHIP_STATUS = 'P' and NVL(SOTSHIP1.LP_STATUS, '0') = '0'" & vbCrLf _
                & " group by SOTSHIP1.WHSE_CODE) X" & vbCrLf _
                & " where ICTWHSE1.LP_CODE is Not Null" & vbCrLf _
                & "   and X.WHSE_CODE (+) = ICTWHSE1.WHSE_CODE"
            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False)
            .Tables("ICTWHSEX").Columns("SHIPS").DataType = GetType(System.Int32)


            ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE FROM SOTPICK1, SOTORDR1 WHERE SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO AND SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC" _
                & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
                & ", SUM (SOTPICK2.PICK_QTY_CONF) PICK_QTY_CONF" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC) PICK_QTY_CANC" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK) PICK_QTY_BACK" _
                & ", SUM (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" _
                & ", SUM (SOTPICK2.PICK_QTY_BACK_REL) PICK_QTY_BACK_REL" _
                & " from SOTPICK2, SOTORDR2, SOTPICK1 " _
                & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
                & " and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO and SOTPICK1.SHIP_BOL_NO = :PARM1" _
                & " group by SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC"
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "V", 0)

            If ASCMAIN1.CLIENT = "AHA" Then
                ASCMAIN1.sql = "Select * from WHTLPXN1 where LP_XNO_SOURCE IN ('" & MENU_ITEM_OBJECT & "', 'SERVICE')" _
                    & " and INIT_DATE >= :PARM1 and INIT_DATE -1  < :PARM2"
            Else
                ASCMAIN1.sql = "Select * from WHTLPXN1 where LP_XNO_SOURCE = '" & MENU_ITEM_OBJECT & "'" _
                    & " and INIT_DATE >= :PARM1 and INIT_DATE -1  < :PARM2"
            End If
            Create_TDA(.Tables.Add, "WHTLPXN1", "**", 0, True, "DD", 1)

            Create_TDA(dst.Tables.Add, "TATCNTRY", "*", 0, False)

            .Tables.Add("TASKS")
            With .Tables("TASKS")
                .Columns.Add("SEQ_NO", GetType(Int32))
                .Columns.Add("TASK_TIME", GetType(DateTime))
                .Columns.Add("TASK_DESC", GetType(String))
            End With
        End With

        grdICTWHSEX.DataSource = dst.Tables("ICTWHSEX")
        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")
        grdWHTLPXN1.DataSource = dst.Tables("WHTLPXN1")

        grdTasks.DataSource = dst.Tables("TASKS")

        Fill_Records("ICTWHSEX")
        Fill_Records("TATCNTRY")

        Create_Summary(grdICTWHSEX, "WHSE_CODE", "Count")
        Create_Summary(grdICTWHSEX, "SHIPS")

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTSHIPX, "SEL")
        Create_Summary(grdSOTPICK1, "PICK_NO", "Count")

        Create_Summary(grdSOTPICKX, "ITEM_CODE", "Count")
        Create_Summary(grdSOTPICKX, New String() _
                       {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK" _
                       , "PICK_QTY_CANC_REL", "PICK_QTY_BACK_REL"})

        Create_Summary(grdWHTLPXN1, "LP_XNO", "Count")


        grdSOTSHIPX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        For Each C As UltraWinGrid.UltraGridColumn In grdSOTSHIPX.DisplayLayout.Bands(0).Columns
            If C.Key = "SEL" Then
                C.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                C.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        With grdSOTSHIPX.DisplayLayout.Bands("SOTSHIPX")
            .Columns("SHIP_BOL_NO").Header.Fixed = True
            .Columns("SEL").Header.Fixed = True
            .Columns("CUST_CODE").Header.Fixed = True
        End With
        With grdSOTPICKX.DisplayLayout.Bands("SOTPICKX")
            For Each COLUMN_NAME As String In New String() {"ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        calFrom.Value = Now.Date.AddDays(-10)
        calTo.Value = Now.Date

        optPending.ValueList.ValueListItems(1).DisplayText = "Transmitted"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCrLf & "You must specify a Warehouse"
                Else
                    rowICTWHSE1 = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCrLf & "Invalid Value specified for Warehouse"
                    Else
                        If rowICTWHSE1.Item("LP_CODE") & "" = "" Then
                            EMsg &= vbCr & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " is not set up as a 3PL"
                        Else
                            rowWHTTPLP1 = LookUp("WHTTPLP1", rowICTWHSE1.Item("LP_CODE"))
                            If rowWHTTPLP1 Is Nothing Then
                                EMsg &= vbCrLf & "Warehouse " & Absx1.txtFor("WHSE_CODE").Text & " Does NOT have a valid value specified for its 3PL"
                            End If
                        End If

                    End If
                End If

                If EMsg = "" Then
                    WHSE_CODE = rowICTWHSE1.Item("WHSE_CODE")
                    LP_CODE = rowICTWHSE1.Item("LP_CODE")

                    If Not ASCMAIN1.Logical_Lock("WHTSPCK1", WHSE_CODE) Then Exit Sub

                End If

            Case "Send"
                Shipments = dst.Tables("SOTSHIPX").Select("SEL = '1' And LP_XMIT_DATE is not Null").Length
                If Shipments > 0 Then
                    EMsg &= vbCr & "There are " & Shipments & " De-Released Shipments Selected. You are not permitted to select De-Released Shipments"
                    Exit Select
                End If

                Shipments = dst.Tables("SOTSHIPX").Select("SEL = '1'").Length
                If Shipments = 0 Then
                    EMsg &= vbCr & "No Shipments Selected"
                    Exit Select
                End If

                ' All shipments for an Order group must be selected
                If Not chkOrderGroup.Checked Then
                    Dim ORDR_GROUP_NO As String = String.Empty
                    For Each row As DataRow In dst.Tables("SOTSHIPX").Select("SEL = '1'", "ORDR_GROUP_NO")
                        If ORDR_GROUP_NO = row.Item("ORDR_GROUP_NO") & String.Empty Then
                            Continue For
                        End If

                        ORDR_GROUP_NO = row.Item("ORDR_GROUP_NO") & String.Empty
                        If dst.Tables("SOTSHIPX").Select("SEL <> '1' AND ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'").Length > 0 Then
                            EMsg &= vbCr & "Order Group (" & ORDR_GROUP_NO & ") does not have all its sales orders selected."
                        End If
                    Next
                End If

                If EMsg = "" Then
                    If optPending.Value = "0" Then
                        If MessageBox.Show("You are about to send " & Shipments & " Shipments Electronically over to the 3PL." _
                                    & vbCrLf _
                                    & "No Changes or De-Releases are Permitted" _
                                    & " to these Orders once they are sent to the 3PL" _
                                    & " without getting the 3PL to Void the corresponding Record in their System." _
                                    & vbCrLf _
                                    & vbCrLf & "OK To Proceed?", "Transmit", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    Else
                        EMsg &= vbCr & "Transmit option not available for Transmitted Sales Orders"
                    End If
                End If

            Case "Request 940 Cancel"
                Shipments = dst.Tables("SOTSHIPX").Select("SEL='1'").Length
                If Shipments = 0 Then
                    EMsg &= vbCr & "No Shipments Selected"
                End If

                If EMsg = "" Then
                    If MsgBox("You are about to send a Request to Cancel " & Shipments & " Shipments to the 3PL" _
                             & vbCrLf _
                             & vbCrLf & "You should have communicated with your CSR before doing this" _
                             & vbCrLf & " to make sure that these Pick Tickets are able to be Cancelled." _
                             & vbCrLf _
                             & vbCrLf & "Once you get a confirmation email, you should then De-Transmit these Shipments." _
                             & vbCrLf _
                             & vbCrLf & "OK To Proceed?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
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

            Case "Send"
                If SendShipmentsTo3PL(Absx1.txtFor("WHSE_CODE").Text) Then
                    Mode_Settings(False)
                End If

            Case "Request 940 Cancel"
                Request_940_Cancel()
                Load_SOTSHIPX()

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Send").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Transmission Controls").Visible = ScreenMode
                .Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tab0.Visible = Not ScreenMode
        splShipments.Visible = ScreenMode

        If ScreenMode Then
            grdSOTSHIPX.Dock = DockStyle.None
            grdSOTSHIPX.Parent = splShipments.Panel1
            grdSOTSHIPX.Dock = DockStyle.Fill
            grdSOTSHIPX.Text = "Shipments Pending Transmission to 3PL"
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("SEL").Hidden = False
            grdSOTSHIPX.Visible = True
        Else
            Clear_Record()
            grdSOTSHIPX.Dock = DockStyle.None
            grdSOTSHIPX.Parent = splTransmissions.Panel2
            grdSOTSHIPX.Dock = DockStyle.Fill
            grdSOTSHIPX.Text = "Shipments Transmitted"
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns("SEL").Hidden = True
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIPX", "SOTPICK1", "SOTPICKX", "WHTLPXN1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        optPending.Value = "0"
        Load_WHTLPXN1()
        Fill_Records("ICTWHSEX")
        Setup_tab0()
        UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
        Clear_All_Filters(grdSOTSHIPX)
        Clear_All_Filters(grdSOTPICK1)
        Clear_All_Filters(grdSOTPICKX)

        Sort_grdColumns(grdTasks, "SEQ_NO")

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Load_SOTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

    End Sub

    Sub Delete_Record()
        Call BeginTrans()
        Stop
        'Call Delete_Records("table")
        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTWHSEX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTSHIPX, "SSSPBBPBBB", "Show Filter", "Show GroupBox", "Show Pins" _
                        , "Select All", "De-Select All", "Select All in Group", "Select All for Customer", "Recall Shipment")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTSHIPX"
                    tlb_pop.Tools("Select All").SharedProps.Visible = True
                    tlb_pop.Tools("De-Select All").SharedProps.Visible = True
                    tlb_pop.Tools("Recall Shipment").SharedProps.Visible = Not ScreenMode

                    If optPending.Value = "D" Then
                        'tlb_pop.Tools("Select All").SharedProps.Visible = False
                    End If
                    tlb_pop.Tools("Recall Shipment").SharedProps.Visible = False

                    tlb_pop.Tools("Select All in Group").SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow)
                    tlb_pop.Tools("Select All for Customer").SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow)

                    If Not ScreenMode Then
                        If grd.ActiveRow.Cells("LP_XNO").Value & "" <> grd.ActiveRow.Cells("LP_XNO_XMIT").Value & "" _
                        Or grd.ActiveRow.Cells("LP_STATUS").Value & "" <> "1" Then
                            tlb_pop.Tools("Recall Shipment").SharedProps.Visible = False
                        End If
                    End If
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Select All"
                Dim selectAll As Int16 = 0

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    If EntryMode = "L" AndAlso optPending.Value = "0" Then
                        If grow.Cells("LP_XMIT_DATE").Value & String.Empty = String.Empty Then
                            grow.Cells("SEL").Value = "1"
                            selectAll += 1
                        End If
                    Else
                        grow.Cells("SEL").Value = "1"
                        selectAll += 1
                    End If
                    grow.Update()
                Next

                MsgBox("You have selected " & selectAll & " Records by Selecting All", MsgBoxStyle.OkOnly, "Verification")

            Case "De-Select All"

                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                    grow.Cells("SEL").Value = "0"
                    grow.Update()
                Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Recall Shipment"
                Dim SHIP_BOL_NO As String = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value
                Dim LP_XNO As String = grdSOTSHIPX.ActiveRow.Cells("LP_XNO").Value
                If MsgBox("Are you sure you want to Recall all Pick Tickets and Shipments for Shipment " & SHIP_BOL_NO,
                          MsgBoxStyle.YesNo,
                          "Verification to Recall Pick Tickets and Shipments from a 3PL") <> MsgBoxResult.Yes Then
                    Exit Sub
                End If

                Recall_Shipment(SHIP_BOL_NO, LP_XNO)

            Case "Select All in Group"
                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("ORDR_GROUP_NO = '" & grd.ActiveRow.Cells("ORDR_GROUP_NO").Value & "'")
                    'rowSOTWSHIPX.Item("SEL") = "1"
                    If EntryMode = "L" AndAlso optPending.Value = "0" Then
                        If rowSOTWSHIPX.Item("LP_XMIT_DATE") & String.Empty = String.Empty Then rowSOTWSHIPX.Item("SEL") = "1"
                    Else
                        rowSOTWSHIPX.Item("SEL") = "1"
                    End If
                Next

            Case "Select All for Customer"
                For Each rowSOTWSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("CUST_CODE = '" & grd.ActiveRow.Cells("CUST_CODE").Value & "'")
                    'rowSOTWSHIPX.Item("SEL") = "1"
                    If EntryMode = "L" AndAlso optPending.Value = "0" Then
                        If rowSOTWSHIPX.Item("LP_XMIT_DATE") & String.Empty = String.Empty Then rowSOTWSHIPX.Item("SEL") = "1"
                    Else
                        rowSOTWSHIPX.Item("SEL") = "1"
                    End If
                Next

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Click_Command("Load")
        End Select
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdICTWHSEX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTWHSEX.DoubleClickRow
        Absx1.txtFor("WHSE_CODE").Text = e.Row.Cells("WHSE_CODE").Text
        Click_Command("Load")
    End Sub

    Private Sub grdWHTLPXN1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdWHTLPXN1.AfterRowActivate
        Setup_WHTLPXN1()
    End Sub

    Private Sub grdSOTSHIPX_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTSHIPX.AfterRowActivate
        Setup_SOTSHIPX()
    End Sub

    Private Sub grdSOTSHIPX_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTSHIPX.BeforeCellUpdate
        If EntryMode = "L" AndAlso optPending.Value = "0" Then
            If e.Cell.Column.Key = "SEL" Then
                If e.Cell.Row.Cells("LP_XMIT_DATE").Value & String.Empty <> String.Empty Then
                    e.Cancel = True
                End If
            End If
        End If
    End Sub

    Private Sub grdSOTSHIPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTSHIPX.InitializeRow
        If Not ScreenMode And EntryMode = "" Then
            If grdWHTLPXN1.ActiveRow IsNot Nothing Then
                If e.Row.Cells("LP_XNO_XMIT").Value & "" <> grdWHTLPXN1.ActiveRow.Cells("LP_XNO").Value & "" Then
                    e.Row.CellAppearance.BackColor = Drawing.Color.Tomato
                Else
                    e.Row.CellAppearance.BackColor = Drawing.Color.Empty
                End If
            End If

        ElseIf EntryMode = "L" AndAlso optPending.Value = "0" Then
            If e.Row.Cells("LP_XMIT_DATE").Value & String.Empty <> String.Empty Then
                e.Row.CellAppearance.BackColor = Drawing.Color.Tomato
            Else
                e.Row.CellAppearance.BackColor = Drawing.Color.Empty
            End If
        Else
            e.Row.CellAppearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub optPending_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPending.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        ' dont want to be here if closing down
        Load_SOTSHIPX()
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Private Sub btnLoadHistory_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadHistory.Click
        Load_WHTLPXN1()
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub AddTask(ByVal TaskDescription As String)
        dst.Tables("TASKS").Rows.Add({dst.Tables("TASKS").Rows.Count + 1, DateTime.Now, TaskDescription})
    End Sub


    Public Function Build_List_of_Objects(Of C As {New})(sql As String) As List(Of C)

        Dim objList As New List(Of C)
        Dim ALL_COLUMNS As Dictionary(Of String, System.Reflection.FieldInfo) _
            = Get_Columns_from_Class(GetType(C))

        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)
        Dim row_count_total As Int32 = tbl.Rows.Count
        Dim row_counter As Int32 = 0

        For Each row As DataRow In tbl.Rows
            row_counter += 1

            Dim objItem As New C

            If 1 <> 1 Then
                ALL_COLUMNS = Get_Columns_from_Class(GetType(C))
            End If

            For Each COLUMN_NAME In ALL_COLUMNS.Keys
                If row.Item(COLUMN_NAME) & "" = "" Then
                Else
                    'Dim P As System.Reflection.MemberInfo = ALL_COLUMNS(COLUMN_NAME)

                    'If row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.DateTime" Then
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, row.Item(COLUMN_NAME), Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.String" Then
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, row.Item(COLUMN_NAME), Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Double" Then
                    '    Dim V As Decimal = Val(row.Item(COLUMN_NAME))
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, V, Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int32" Then
                    '    Dim V As Int32 = Val(row.Item(COLUMN_NAME))
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, V, Nothing)
                    'ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int16" Then
                    '    Dim V As Int16 = Val(row.Item(COLUMN_NAME))
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, V, Nothing)
                    'Else
                    '    P.GetType().GetProperty(COLUMN_NAME).SetValue(P, Val(row.Item(COLUMN_NAME)), Nothing)
                    'End If

                    If row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.DateTime" Then
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, row.Item(COLUMN_NAME))
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.String" Then
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, row.Item(COLUMN_NAME))
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Double" Then
                        Dim V As Decimal = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int32" Then
                        Dim V As Int32 = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    ElseIf row.Table.Columns(COLUMN_NAME).DataType.ToString = "System.Int16" Then
                        Dim V As Int16 = Val(row.Item(COLUMN_NAME))
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, V)
                    Else
                        ALL_COLUMNS(COLUMN_NAME).SetValue(objItem, Val(row.Item(COLUMN_NAME)))
                    End If

                End If
            Next
            objList.Add(objItem)

            'If row_counter > 100 Then
            '    Exit For
            'End If
        Next

        Return objList
    End Function

    Public Shared Function Get_Columns_from_Class(T As Type) _
        As Dictionary(Of String, System.Reflection.FieldInfo)

        Dim COLUMN_NAMEs As New Dictionary(Of String, System.Reflection.FieldInfo)
        ' Dim COLUMN_NAMEs As New Dictionary(Of String, System.Reflection.PropertyInfo)

        'Dim t As Type = XX.GetType
        Dim fieldName As String
        ' Dim propertyValue As Object

        ' Use each property of the business object passed in 
        'For Each pi As System.Reflection.PropertyInfo In _
        '        T.GetProperties(System.Reflection.BindingFlags.Instance Or _
        '                        System.Reflection.BindingFlags.Public Or _
        '                        System.Reflection.BindingFlags.NonPublic)
        '    ' Get the name and value of the property 
        '    If pi.Name <> "ExtensionData" Then
        '        fieldName = pi.Name
        '        COLUMN_NAMEs.Add(fieldName, pi)
        '    End If

        '    ' Get the value of the property 
        '    ' propertyValue = pi.GetValue(XX, Nothing)
        '    'Console.WriteLine(fieldName & ": " &
        '    'If(propertyValue Is Nothing, "Nothing", propertyValue.ToString))
        'Next

        For Each pi As System.Reflection.FieldInfo In
               T.GetFields(System.Reflection.BindingFlags.Instance Or
                               System.Reflection.BindingFlags.Public Or
                               System.Reflection.BindingFlags.NonPublic)
            If pi.MemberType = Reflection.MemberTypes.Field Then
                fieldName = pi.Name
                If fieldName <> "SQL" Then
                    ' Debug.Write(pi.Name & ":" & pi.MemberType.ToString)
                    COLUMN_NAMEs.Add(fieldName, pi)
                End If
            End If
        Next
        Return COLUMN_NAMEs
    End Function

    Sub Setup_WHTLPXN1()
        If grdWHTLPXN1.ActiveRow Is Nothing Then
            grdSOTSHIPX.Visible = False
        Else
            grdSOTSHIPX.Visible = True
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
            Dim LP_XNO As String = grdWHTLPXN1.ActiveRow.Cells("LP_XNO").Value
            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " (SHIP_BOL_NO) Select SHIP_BOL_NO from WHTSHIPX where LP_XNO = '" & LP_XNO & "'")
            Fill_Records("SOTSHIPX")
            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO")
        End If
    End Sub

    Sub Setup_tab0()
        UltraExplorerBar1.Groups("Transmission History").Visible = Not ScreenMode And tab0.SelectedTab.Key = "Transmissions"
    End Sub

    Sub Load_WHTLPXN1()
        Fill_Records("WHTLPXN1", New Object() {calFrom.Value, calTo.Value})
        Sort_grdColumns(grdWHTLPXN1, "LP_XNO".ToLower)
        Setup_WHTLPXN1()
    End Sub

    Sub Setup_SOTSHIPX()
        If grdSOTSHIPX.ActiveRow Is Nothing OrElse Not grdSOTSHIPX.ActiveRow.IsDataRow Then
            tabShipment.Visible = False
        Else
            tabShipment.Visible = True
            Dim SHIP_BOL_NO As String = grdSOTSHIPX.ActiveRow.Cells("SHIP_BOL_NO").Value & ""
            grdSOTPICK1.Text = "Pick Tickets for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICK1", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTPICK1, "PICK_NO")
            grdSOTPICKX.Text = "Item Summary for Shipment No " & SHIP_BOL_NO
            Fill_Records("SOTPICKX", SHIP_BOL_NO)
            Sort_grdColumns(grdSOTPICKX, "ITEM_CODE")
        End If
    End Sub

    Sub Recall_Shipment(SHIP_BOL_NO As String, LP_XNO As String)

        If Not ASCMAIN1.Logical_Lock("WHTSPCK1", WHSE_CODE) Then Exit Sub
        If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub

        Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NO)
        If rowSOTSHIP1.Item("LP_XNO") & "" <> LP_XNO _
        Or rowSOTSHIP1.Item("LP_STATUS") <> "1" Then
            ' DO NOTHING, SOMETHING HAS CHANGED
        Else
            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL where SHIP_BOL_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {SHIP_BOL_NO})
        End If

        Mode_Settings(False)

    End Sub

    Sub De_Transmit_3PL()

        Dim sqlSHIP_BOL_NO As String = ""
        For Each row As DataRow In dst.Tables("SOTSHIPX").Select("")
            Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO")
            sqlSHIP_BOL_NO &= ",'" & SHIP_BOL_NO & "'"
        Next
        sqlSHIP_BOL_NO = " where SHIP_BOL_NO in (" & Mid(sqlSHIP_BOL_NO, 2) & ")"
        Dim sqlPICK_NO As String = " where PICK_NO in (Select PICK_NO from ADS.SOTPICK1_3PL@ADSIIS" & sqlSHIP_BOL_NO & ")"


        For Each TABLE_NAME As String In New String() _
            {"SOTCART2", "SOTCART1", "SOTPICK2", "SOTPICK1", "SOTSHIP1"}

            ASCMAIN1.Progress("-", TABLE_NAME)

        Next
    End Sub

    Sub Write_Notes(EDI_OUTBOUND_DOC_NO As String, EDI_NTE_TYPE As String, NOTES As String)

        Dim EDI_NTE_SEQ_NO As Int32 = 0
        For Each NOTE As String In Split(NOTES, vbCrLf)
            NOTE = Trim(NOTE)
            Do While NOTE <> ""
                EDI_NTE_SEQ_NO += 1

                Dim EDI_NTE As String
                If NOTE.Length > 40 Then
                    EDI_NTE = Mid(NOTE, 1, 40)
                    NOTE = Mid(NOTE, 41)
                Else
                    EDI_NTE = NOTE
                    NOTE = ""
                End If

                EDI_NTE = Replace(EDI_NTE, "*", "@")

                Dim rowEDT940O4 As DataRow = dst.Tables("EDT940O4").NewRow
                With rowEDT940O4
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("EDI_NTE_TYPE") = EDI_NTE_TYPE
                    .Item("EDI_NTE_SEQ_NO") = EDI_NTE_SEQ_NO
                    .Item("EDI_NTE") = EDI_NTE
                End With
                dst.Tables("EDT940O4").Rows.Add(rowEDT940O4)
            Loop
        Next

    End Sub

    Sub Load_SOTSHIPX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Shipments Queue")

        If optPending.Value = "0" Then
            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & ", '0' SEL, SHIP_856_IND EDI856, SOTSHIP1.SHIP_CART_REQD, NULL ORDR_NO" & vbCrLf _
                & ", SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST, NULL CUST_CODE from SOTSHIP1" _
                & " where SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"

            UltraExplorerBar1.Groups("Screen Control").Items("Send").Text = "Transmit"
            ASCMAIN1.sql &= " and nvl(SOTSHIP1.LP_STATUS,'0') = '0'"
            grdSOTSHIPX.Text = "Shipments Pending Transmission to 3PL (" & WHSE_CODE & ")"

            UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.False
            UltraExplorerBar1.Groups("Screen Control").Items("Send").Settings.Enabled = DefaultableBoolean.True
        Else
            ASCMAIN1.sql = "Select SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                 & ", '0' SEL, SHIP_856_IND EDI856, SOTSHIP1.SHIP_CART_REQD, NULL ORDR_NO" & vbCrLf _
                 & ", SHIP_ADDR_TYPE ORDR_ADDR_TYPE_ST, NULL CUST_CODE from SOTSHIP1" _
                 & " where SOTSHIP1.WHSE_CODE = '" & WHSE_CODE & "'"

            ASCMAIN1.sql &= " and nvl(SOTSHIP1.LP_STATUS, '0') = '1'" & vbCrLf
            grdSOTSHIPX.Text = "Shipments Sent to 3PL (" & WHSE_CODE & ")"

            UltraExplorerBar1.Groups("Screen Control").Items("Request 940 Cancel").Settings.Enabled = DefaultableBoolean.True
            UltraExplorerBar1.Groups("Screen Control").Items("Send").Settings.Enabled = DefaultableBoolean.False
        End If

        For Each colName As String In New String() {"LP_XNO"} ', "LP_XMIT_DATE"}
            grdSOTSHIPX.DisplayLayout.Bands(0).Columns(colName).Hidden = optPending.Value = "0"
        Next

        ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
        ASCDATA1.ExecuteSQL("Insert into " & SOTSHIPX & " " & ASCMAIN1.sql)

        'ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
        '                    & "Set ORDR_NO = (Select Min (ORDR_NO) ORDR_NO from SOTPICK1 " _
        '                    & " where SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO)")
        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT SOTPICK1.SHIP_BOL_NO, MIN(SOTPICK1.ORDR_NO) ORDR_NO" _
            & " FROM SOTPICK1, " & SOTSHIPX & " SOTSHIPX " _
            & " WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIPX.SHIP_BOL_NO GROUP BY SOTPICK1.SHIP_BOL_NO;" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " UPDATE " & SOTSHIPX & " SET ORDR_NO = R1.ORDR_NO WHERE SHIP_BOL_NO = R1.SHIP_BOL_NO;" _
            & " END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        'ASCDATA1.ExecuteSQL("Update " & SOTSHIPX & " SOTSHIPX " _
        '            & "Set CUST_CODE = (Select CUST_CODE from SOTORDR1 where ORDR_NO = SOTSHIPX.ORDR_NO)")

        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
            & " SELECT SOTORDR1.ORDR_NO, SOTORDR1.CUST_CODE" _
            & " FROM SOTORDR1," & SOTSHIPX & " SOTSHIPX " _
            & " WHERE SOTORDR1.ORDR_NO = SOTSHIPX.ORDR_NO;" _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " UPDATE " & SOTSHIPX & " SET CUST_CODE = R1.CUST_CODE WHERE ORDR_NO = R1.ORDR_NO;" _
            & " END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)


        Fill_Records("SOTSHIPX")
        Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)

        Setup_SOTSHIPX()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Request_940_Cancel()

        ' ADS does not use 940's
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing to Cancel Shipments")

        Try
            Dim shipmentList As New List(Of String)

            For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
                Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("SHIP_BOL_NO")
                Dim ORDR_GROUP_NO As String = rowSOTSHIPX.Item("ORDR_GROUP_NO")

                If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then
                    Continue For
                End If

                Dim rowSOTSHIP1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTSHIP1 where SHIP_BOL_NO = :PARM1", "V", SHIP_BOL_NO)
                If rowSOTSHIP1 Is Nothing Then
                    MessageBox.Show("Shipment (" & SHIP_BOL_NO & ") cannot be found.", "940 Cancel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Continue For
                ElseIf rowSOTSHIP1.Item("SHIP_STATUS") & String.Empty <> "P" Then
                    MessageBox.Show("Shipment (" & SHIP_BOL_NO & ") is not 'In Pick'.", "940 Cancel", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Continue For
                End If

                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then
                    Continue For
                End If

                shipmentList.Add(SHIP_BOL_NO)
            Next

            Try
                BeginTrans()

                Dim listOfShipments As String = String.Join("', '", shipmentList.ToArray)
                listOfShipments = "'" & listOfShipments & "'"

                ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL, SHIP_PICK_PRINTED = NULL where SHIP_BOL_NO in (" & listOfShipments & ")"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                CommitTrans("Request Successful")

            Catch ex As Exception
                Rollback(ex.Message)
            End Try

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Cancel 940", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
            ASCMAIN1.MultiTask_Release()
        End Try

        '*******************************************************************************************************************
        '*******************************************************************************************************************
        '*******************************************************************************************************************
        '*******************************************************************************************************************
        '*******************************************************************************************************************


        If 1 = 1 Then
            Exit Sub
        End If

        BeginTrans()

        If Not dst.Tables.Contains("EDT940O1") Then
            Create_TDA(dst.Tables.Add, "EDT940O1", "*")
            Create_TDA(dst.Tables.Add, "EDT940O2", "*")
            Create_TDA(dst.Tables.Add, "EDT940O4", "*")
            Create_TDA(dst.Tables.Add, "EDT940O5", "*")
        Else
            For Each tableName As String In New String() {"EDT940O1", "EDT940O2", "EDT940O4", "EDT940O"}
                dst.Tables(tableName).Clear()
            Next
        End If

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
        Dim rowEDTTRPM1 As DataRow = LookUp("EDTTRPM1",
                                            New String() {rowICTWHSE1.Item("WHSE_EDI_QUAL"), rowICTWHSE1.Item("WHSE_EDI_ID"), "943"})


        For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
            Dim CUST_CODE As String = rowSOTSHIPX.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim rowEDTSLSP1 As DataRow = LookUp("EDTSLSP1", CUST_CODE)
            Dim SHIP_BOL_NO As String = rowSOTSHIPX.Item("SHIP_BOL_NO")
            Dim SHIP_VIA_CODE As String = rowSOTSHIPX.Item("SHIP_VIA_CODE") & ""
            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE, True)

            ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '0', LP_XNO = NULL, SHIP_PICK_PRINTED = NULL" _
               & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTPICK1.* from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P'"
            For Each rowSOTPICK1 As DataRow In ASCDATA1.GetDataTable.Select("")

                Dim EDI_OUTBOUND_DOC_NO As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                'Dim EDI_DOC_SEQ_NO As String = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""
                'Dim rowEDT850T1 As DataRow = LookUp("EDT850T1", EDI_DOC_SEQ_NO)

                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & ""
                Dim CUST_DC_NO As String = rowSOTORDR1.Item("CUST_DC_NO") & ""

                Dim rowEDT940O1 As DataRow = dst.Tables("EDT940O1").NewRow
                With rowEDT940O1
                    .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                    .Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("PICK_NO") = PICK_NO
                    .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                    .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                    .Item("EDI_SUPPLIER_NO") = rowARTCUST1.Item("CUST_VEND_REF")
                    .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                    .Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
                    .Item("ORDR_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
                    .Item("ORDR_PO_DATE") = rowSOTORDR1.Item("ORDR_DATE")

                    Dim FRT_TERMS As String = rowSOTSHIPX.Item("FRT_TERMS") & ""
                    Dim FRT_TERMS_EDI As String = ""
                    Select Case FRT_TERMS
                        Case "PPD", "PPA"
                            FRT_TERMS_EDI = "PP"
                        Case "COL"
                            FRT_TERMS_EDI = "CC"
                    End Select
                    .Item("FRT_TERMS") = FRT_TERMS_EDI

                    '.Item("EDI_TRANS_METH_CODE") = "?"
                    '.Item("EDI_SERVICE_LEVEL") = "?"
                    '.Item("EDI_TP_BILLING_ACCT") = "? ' IF FRT_TERMS WAS 3RD PARTY WE WOULD SEND THE 3RD PARTY ACCT NUMBER
                    .Item("EDI_SCAC_CODE") = "ROUT" ' rowSOTSVIA1.Item("SHIP_VIA_SCAC")

                    Dim EDI_DIVISION_CODE As String = rowICTWHSE1.Item("LP_WHSE_ID") & ""

                    ASCMAIN1.sql = "Select * from EDT940O1 where PICK_NO = '" & PICK_NO & "'"
                    Dim rowEDT940O1_prior As DataRow = ASCDATA1.GetDataRow
                    If rowEDT940O1_prior IsNot Nothing Then
                        EDI_DIVISION_CODE = rowEDT940O1_prior.Item("EDI_DIVISION_CODE") & ""
                    End If
                    .Item("EDI_DIVISION_CODE") = EDI_DIVISION_CODE

                    ' .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("CUST_VEND_REF")
                    .Item("EDI_LABEL_FORMAT") = rowARTCUST1.Item("LABEL_TEMPLATE_CODE")
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("ORDR_TYPE_CODE") = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                    'If rowEDT850T1 IsNot Nothing Then
                    '    .Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")
                    'End If
                    .Item("EDI_MERCH_TYPE") = rowSOTORDR1.Item("EDI_MERCH_TYPE")
                    .Item("ORDR_STATUS_CODE") = "V"
                End With
                dst.Tables("EDT940O1").Rows.Add(rowEDT940O1)

                ASCMAIN1.sql = "Insert into EDTSYSIH (COMPANY_CODE,EDI_OUTBOUND_DOC_NO,EDI_APPLICATION_ID,EDI_PROCESS_IND," _
                    & "EDI_OUR_ID,EDI_TP_ID,INIT_DATE,INIT_OPER)" _
                    & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,SYSDATE,'" & ASCMAIN1.USER_ID & "')"
                Dim EDI_APPLICATION_ID As String = "OW"
                Dim EDI_PROCESS_IND As String = "1"
                ' EDI_PROCESS_IND = "T"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVVVV",
                        New Object() {ASCMAIN1.DBS_COMPANY, EDI_OUTBOUND_DOC_NO, EDI_APPLICATION_ID, EDI_PROCESS_IND,
                                      rowEDTTRPM1.Item("EDI_OUR_ID"), rowICTWHSE1.Item("WHSE_EDI_ID")})
            Next
        Next

        Update_Record_TDA("EDT940O1")
        Update_Record_TDA("EDT940O2")
        Update_Record_TDA("EDT940O4")
        Update_Record_TDA("EDT940O5")

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Function SendShipmentsTo3PL(ByVal WHSE_CODE As String) As Boolean

        Try
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Preparing to Transfer Shipments")

            Dim clsSOCADSO1 As New TAC.SOCADSO1
            Dim listOfShipments As New List(Of String)
            Dim shipmentsSentTo3pl As String = String.Empty

            For Each rowSOTSHIPX As DataRow In dst.Tables("SOTSHIPX").Select("SEL='1'")
                listOfShipments.Add(rowSOTSHIPX.Item("SHIP_BOL_NO"))
            Next

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            If rowICTWHSE1 Is Nothing Then
                MessageBox.Show($"Unknown Warehouse {WHSE_CODE}", "Send Shipments To 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                SendShipmentsTo3PL = False
                Exit Function
            End If

            Dim LP_CODE As String = rowICTWHSE1.Item("LP_CODE") & String.Empty

            dst.Tables("TASKS").Rows.Clear()
            clsSOCADSO1.tblTasks = Nothing
            clsSOCADSO1.tblTasks = dst.Tables("TASKS")
            Sort_grdColumns(grdTasks, "SEQ_NO")

            Select Case LP_CODE

                Case "ADS"
                    AddTask("WHFSPCK1 Call Prepare ADS Sales Orders File")

                    If Not clsSOCADSO1.PrepareADSSalesOrdersFile(listOfShipments, WHSE_CODE) Then
                        MessageBox.Show("The following occured when releasing to Warehouse (" & WHSE_CODE & "): " & vbCr & clsSOCADSO1.LastError, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        SendShipmentsTo3PL = False
                        Exit Function
                    End If

                    shipmentsSentTo3pl = "'" & String.Join("', '", listOfShipments.ToArray) & "'"
                    SendShipmentsTo3PL = True

                Case "CLA"
                    AddTask("WHFSPCK1 Call Prepare Clarins File")
                    SendShipmentsTo3PL = clsSOCADSO1.Prepare_Clarins_File(listOfShipments)

                    If SendShipmentsTo3PL = False Then
                        MessageBox.Show("The following occured when releasing to Warehouse (" & WHSE_CODE & "): " & vbCr & clsSOCADSO1.LastError, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Function
                    End If

                    shipmentsSentTo3pl = "'" & String.Join("', '", listOfShipments.ToArray) & "'"
                    If clsSOCADSO1.LastError.Length > 0 Then
                        Dim zmsg As String = "Some Shipments were successfuly sent to the 3PL." & Environment.NewLine & Environment.NewLine
                        zmsg &= "However, the following occured when releasing to Warehouse (" & WHSE_CODE & "): " & vbCr & clsSOCADSO1.LastError
                        MessageBox.Show(zmsg, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If

                Case Else
                    MessageBox.Show($"Invalid Warehouse {WHSE_CODE}, LP Code {LP_CODE} - No Shipments processed", "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    SendShipmentsTo3PL = False
                    Exit Function
            End Select

            Try
                Select Case LP_CODE

                    Case "ADS"
                        LP_XNO = clsSOCADSO1.ReleaseShipmentNo

                        BeginTrans()

                        AddTask("WHFSPCK1 Update SOTSHIP1")

                        ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '1', LP_XNO = '" & LP_XNO & "', LP_XMIT_DATE = SYSDATE"
                        ASCMAIN1.sql &= ", SHIP_PICK_PRINTED = SYSDATE"
                        ASCMAIN1.sql &= " where SHIP_BOL_NO in (" & shipmentsSentTo3pl & ")"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                        AddTask("WHFSPCK1 Update SOTPICK1")

                        ASCMAIN1.sql = "Update SOTPICK1 Set PICK_PRINTED = SYSDATE, ORDR_NO_3PL = PICK_NO"
                        ASCMAIN1.sql &= ", PICK_PRINTED_OPER = '" & ASCMAIN1.USER_ID & "'"
                        ASCMAIN1.sql &= " where SHIP_BOL_NO in (" & shipmentsSentTo3pl & ")"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                        dst.Tables("WHTSHIPX").Rows.Clear()
                        dst.Tables("WHTLPXN1").Rows.Clear()

                        For Each shipBol As String In shipmentsSentTo3pl.Split(",")
                            dst.Tables("WHTSHIPX").Rows.Add(New Object() {LP_XNO, shipBol.Replace("'", "").Trim})
                        Next

                        AddTask("WHFSPCK1 Update WHTSHIPX")
                        Update_Record_TDA("WHTSHIPX")

                        Dim rowWHTLPXN1 As DataRow = dst.Tables("WHTLPXN1").NewRow
                        rowWHTLPXN1.Item("LP_XNO") = LP_XNO
                        rowWHTLPXN1.Item("LP_XNO_SOURCE") = MENU_ITEM_OBJECT
                        rowWHTLPXN1.Item("LP_XNO_RECORDS") = dst.Tables("WHTSHIPX").Rows.Count
                        rowWHTLPXN1.Item("LP_XNO_NOTES") = ""
                        rowWHTLPXN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowWHTLPXN1.Item("INIT_DATE") = DateTime.Now
                        dst.Tables("WHTLPXN1").Rows.Add(rowWHTLPXN1)

                        AddTask("WHFSPCK1 Update WHTLPXN1")
                        Update_Record_TDA("WHTLPXN1")

                        AddTask("WHFSPCK1 CommitTrans")
                        CommitTrans()
                        SendShipmentsTo3PL = True
                End Select

                Try
                    ' Show Address where the goods are shipping
                    Dim specialDeliveries As String = String.Empty
                    Dim rowSOTORDR5 As DataRow = Nothing
                    Dim addressLine As String = String.Empty
                    ',CUST_ADDR1,CUST_ADDR2,CUST_ADDR3,CUST_CITY,CUST_STATE,CUST_ZIP_CODE,CUST_COUNTRY
                    Dim sql As String = "Select * from SOTORDR5 where CUST_ADDR_TYPE = 'ST' and ORDR_NO = (SELECT ORDR_NO from SOTPICK1 where PICK_NO = :PARM1)"

                    AddTask("WHFSPCK1 GetPriorityDeliveries")
                    If clsSOCADSO1.GetPriorityDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are PRIORITY Deliveries"
                        For Each priorityDelivery As String In clsSOCADSO1.GetPriorityDeliveries
                            priorityDelivery = priorityDelivery.Trim
                            If priorityDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", priorityDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & priorityDelivery & addressLine & Environment.NewLine
                            End If
                        Next

                    End If

                    AddTask("WHFSPCK1 GetNextDayDeliveries")
                    If clsSOCADSO1.GetNextDayDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are Next Day Deliveries"
                        For Each nextDayDelivery As String In clsSOCADSO1.GetNextDayDeliveries
                            nextDayDelivery = nextDayDelivery.Trim
                            If nextDayDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", nextDayDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & nextDayDelivery & addressLine & Environment.NewLine
                            End If
                        Next
                    End If

                    AddTask("WHFSPCK1 GetSecondDayDeliveries")
                    If clsSOCADSO1.GetSecondDayDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are 2nd Day Deliveries"
                        For Each secondDayDelivery As String In clsSOCADSO1.GetSecondDayDeliveries
                            secondDayDelivery = secondDayDelivery.Trim
                            If secondDayDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", secondDayDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & secondDayDelivery & addressLine & Environment.NewLine
                            End If
                        Next
                    End If

                    AddTask("WHFSPCK1 GetPosDeliveries")
                    If clsSOCADSO1.GetPosDeliveries.Count > 0 Then
                        specialDeliveries &= Environment.NewLine & "The following shipments are POS Deliveries"
                        For Each PosDelivery As String In clsSOCADSO1.GetPosDeliveries
                            PosDelivery = PosDelivery.Trim
                            If PosDelivery.Length > 0 Then
                                rowSOTORDR5 = ASCDATA1.GetDataRow(sql, "V", PosDelivery)
                                addressLine = String.Empty
                                If rowSOTORDR5 IsNot Nothing Then
                                    addressLine = Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_NAME") & String.Empty
                                    For Each field As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3"}
                                        If rowSOTORDR5.Item(field) & String.Empty <> String.Empty Then
                                            addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item(field) & String.Empty
                                        End If
                                    Next
                                    addressLine &= Environment.NewLine & vbTab & vbTab & rowSOTORDR5.Item("CUST_CITY") & ", " & rowSOTORDR5.Item("CUST_STATE") & "  " & rowSOTORDR5.Item("CUST_ZIP_CODE")
                                End If
                                specialDeliveries &= Environment.NewLine & vbTab & vbTab & PosDelivery & addressLine & Environment.NewLine
                            End If
                        Next
                    End If

                    ' ISSUE-7297 - Message in Sales Order Email to ADS
                    ' 04/22/2026 - Cancelled Pick Tickets appeared in the email sent to 3PL. Added AND SOTPICK1.PICK_STATUS = 'P'
                    Dim pickData As String = String.Empty
                    sql = " SELECT SOTPICK1.PICK_NO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_NAME, SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_ADDR_TYPE_ST, SOTORDR1.ORDR_MESSAGE
                                FROM SOTORDR1, SOTPICK1
                                WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO 
                                AND SOTPICK1.PICK_STATUS = 'P'
                                AND SOTPICK1.SHIP_BOL_NO IN (" & shipmentsSentTo3pl & ")"

                    pickData = "<table cellpadding=""5"" cellspacing=""5"" width=""910"">" & Environment.NewLine

                    pickData &= "<tr>"
                    pickData &= "<td colspan=""6"" style=""font-weight:bold;"">Pick Ticket Detail:</td>"
                    pickData &= "</tr>" & Environment.NewLine

                    pickData &= "<tr>"
                    pickData &= "<th style=""text-align:left; width:100px;"">Pick No</th>"
                    pickData &= "<th style=""text-align:left; width:150px;"">Customer</th>"
                    pickData &= "<th style=""text-align:left; width:300px;"">Customer Name</th>"
                    pickData &= "<th style=""text-align:left; width:200px;"">PO</th>"
                    pickData &= "<th style=""text-align:left; width:200px;"">Message</th>"
                    pickData &= "<th style=""width:10px;""></th>"
                    pickData &= "</tr>" & Environment.NewLine

                    AddTask("WHFSPCK1 Create Email Pick Information")
                    For Each row As DataRow In ASCDATA1.GetDataTable(sql).Select("", "PICK_NO")
                        Dim CUST_CODE As String = row.Item("CUST_CODE")
                        Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")

                        Dim XREF_CUST_CODE_SHIP_TO As String = row.Item("CUST_CODE")
                        Dim XREF_CUST_STORE_NO_SHIP_TO As String = String.Empty

                        If WHSE_CODE = "CLA" Then
                            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

                            If rowARTCUST1.Item("CUST_SHIP_TO_MANUAL") & String.Empty = "1" Then
                                XREF_CUST_CODE_SHIP_TO = rowARTCUST1.Item("CUST_NO_3PL") & "/" & rowARTCUST1.Item("CUST_STORE_NO_3PL") & String.Empty
                            Else
                                XREF_CUST_CODE_SHIP_TO = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                                XREF_CUST_STORE_NO_SHIP_TO = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty

                                Select Case row.Item("ORDR_ADDR_TYPE_ST") & String.Empty
                                    Case "DC"
                                        Dim CUST_DC_NO As String = rowARTCUST2.Item("CUST_DC_NO") & String.Empty
                                        rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_DC_NO})

                                        ' Set to Blank so if the DC record is missing the Clarins Settings an error is generated
                                        XREF_CUST_CODE_SHIP_TO = String.Empty
                                        XREF_CUST_STORE_NO_SHIP_TO = String.Empty

                                        If rowARTCUST2 IsNot Nothing Then
                                            XREF_CUST_CODE_SHIP_TO = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                                            XREF_CUST_STORE_NO_SHIP_TO = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty
                                        End If
                                End Select
                            End If
                            XREF_CUST_CODE_SHIP_TO &= "/" & XREF_CUST_STORE_NO_SHIP_TO
                        Else
                            XREF_CUST_CODE_SHIP_TO = CUST_CODE & "/" & CUST_STORE_NO
                        End If

                        pickData &= "<tr>"
                        pickData &= "<td>" & WebUtility.HtmlEncode(row.Item("PICK_NO").ToString()) & "</td>"
                        pickData &= "<td>" & WebUtility.HtmlEncode(XREF_CUST_CODE_SHIP_TO) & "</td>"
                        pickData &= "<td>" & WebUtility.HtmlEncode(row.Item("CUST_NAME").ToString()) & "</td>"
                        pickData &= "<td>" & WebUtility.HtmlEncode(row.Item("ORDR_CUST_PO").ToString()) & "</td>"

                        ' ISSUE-7297 In the email, would we be able to highlight the message in yellow, so it catches ADS' eye?
                        'pickData &= "<td>" & WebUtility.HtmlEncode(row.Item("ORDR_MESSAGE").ToString()) & "</td>"
                        If Not String.IsNullOrWhiteSpace(row.Item("ORDR_MESSAGE").ToString()) Then
                            pickData &= "<td style='background-color: yellow;'>" & WebUtility.HtmlEncode(row.Item("ORDR_MESSAGE").ToString()) & "</td>"
                        Else
                            pickData &= "<td>" & WebUtility.HtmlEncode(row.Item("ORDR_MESSAGE").ToString()) & "</td>"
                        End If

                        pickData &= "<td></td>"
                        pickData &= "</tr>" & Environment.NewLine
                    Next
                    pickData &= "</table>"

                    AddTask("WHFSPCK1 Select MAX(LP_XNO) from WHTSHIPX")
                    ASCMAIN1.sql = "Select MAX(LP_XNO) from WHTSHIPX WHERE SHIP_BOL_NO IN (" & shipmentsSentTo3pl & ")"
                    LP_XNO = ASCDATA1.GetDataValue(ASCMAIN1.sql) & String.Empty

                    If LP_XNO.Length > 0 Then
                        AddTask("WHFSPCK1 Start EmailIPLBShipments")
                        EmailIPLBShipments(LP_XNO, WHSE_CODE, listOfShipments.Count, specialDeliveries, pickData)
                        AddTask("WHFSPCK1 End EmailIPLBShipments")
                    End If

                Catch ex As Exception
                    MessageBox.Show("The following error occurred: " & ex.Message, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try

            Catch ex As Exception
                Rollback(ex.Message)
                SendShipmentsTo3PL = False
            End Try

            MessageBox.Show("Shipments transferred successfully.", "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
            SendShipmentsTo3PL = False

        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("", "")
            AddTask("WHFSPCK1 Process Complete")
            Sort_grdColumns(grdTasks, "SEQ_NO")
        End Try

    End Function

    Private Sub EmailIPLBShipments(ByVal LP_XNO As String,
                                  ByVal WHSE_CODE As String,
                                  ByVal NumOfShipments As Int32,
                                  ByVal specialDeliveries As String,
                                  ByVal shipmentData As String)

        If LP_XNO.Length > 0 Then
            Dim emailTransferred As Boolean = False

            While Not emailTransferred
                Dim objASCNOTEE As New TAC.ASCNOTEE(ASCMAIN1.Folders, "SHIP_" & WHSE_CODE, dst)
                Dim emailNote As String = $"Batch number {LP_XNO} from IPLB has been uploaded and includes {NumOfShipments} shipments for warehouse {WHSE_CODE}." & Environment.NewLine

                If specialDeliveries.Length > 0 Then
                    emailNote &= Environment.NewLine & specialDeliveries & Environment.NewLine & Environment.NewLine & shipmentData
                Else
                    emailNote &= shipmentData
                End If

                objASCNOTEE.Note = emailNote
                objASCNOTEE.CreateComponents()
                objASCNOTEE.SetEmailSubject($"Shipment files sent to {WHSE_CODE}")
                objASCNOTEE.EmailDocument()

                If objASCNOTEE.lastError.Length = 0 Then
                    Exit While
                End If

                MessageBox.Show($"The following error occurred when generating the Shipments email: {objASCNOTEE.lastError}", "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)

                If MessageBox.Show("Do you want to try and resend the email? If you choose 'Yes' the system will wait 3 seconds then retry.", "Process Shipments", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    emailTransferred = True
                Else
                    System.Threading.Thread.Sleep(3000)
                End If
            End While
        End If

    End Sub


#End Region

End Class