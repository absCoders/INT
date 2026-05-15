Imports ABSolution

Public Class EDF945T1

    Private BOL_NO As String = String.Empty
    Private EDI_SHIPMENT_ID As String = String.Empty
    Private PICK_NO As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "EDT945T1", "SELECT * FROM EDT945T1 WHERE EDI_SHIPMENT_ID = :PARM1", 0, , "V")
            Create_TDA(.Tables.Add, "EDT945T2", "*")

            Create_TDA(.Tables.Add, "SOTPICK1", "*")

            ASCMAIN1.sql = "SELECT SOTORDR2.ITEM_CODE, SOTPICK2.* 
                                FROM SOTPICK2, SOTORDR2
                                WHERE SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO
                                AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO
                                AND SOTPICK2.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 0)

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)

            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            Fill_Records("SOTSVIA1", String.Empty, True, "SELECT * FROM SOTSVIA1")

            ASCMAIN1.sql = "select sotpickx.*, shiphdr.num_hdr
                                from
                                (select sotpick1.ship_bol_no, sotordr0.cust_code, sotordr0.ordr_cust_po, sotship1.SHIP_ADDR_CODE, count(*) num_pick 
                                        from sotpick1, sotship1, sotordr0
                                        where sotpick1.pick_status = 'P' 
                                        and sotpick1.ship_bol_no = sotship1.ship_bol_no
                                        and sotship1.ordr_group_no = sotordr0.ordr_group_no
                                        group by  sotpick1.ship_bol_no, sotordr0.cust_code, sotordr0.ordr_cust_po, sotship1.SHIP_ADDR_CODE) sotpickx,
                                (select ship_bol_no, count(*) num_hdr from conv.cfg_shiphdr group by ship_bol_no) shiphdr
                                where sotpickx.ship_bol_no = shiphdr.ship_bol_no (+)"
            Create_TDA(.Tables.Add, "SOTPICKX", "**", 0, False, "", 0)

            ASCMAIN1.sql = "SELECT SOTORDR2.ITEM_CODE, SOTPICK2.* 
                                FROM SOTPICK2, SOTORDR2
                                WHERE SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO
                                AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO
                                AND SOTPICK2.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK2_X", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "Select 0 MISSING, sotpick1.* from sotpick1 where ship_bol_no = :PARM1 and pick_status = 'P'"
            Create_TDA(.Tables.Add, "SOTPICK1_X", "**", 0, False, "V")

            ASCMAIN1.sql = "Select * from conv.cfg_shiphdr where ship_bol_no = :PARM1"
            Create_TDA(.Tables.Add, "SHIPHDR", "**", 0, False, "V")

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, CARTON.* 
                                    from conv.cfg_carton CARTON, ICTITEM1
                                    where CARTON.chinvn = :PARM1
                                    AND CARTON.CDITEM = ICTITEM1.ITEM_ALT_SORT (+)"
            Create_TDA(.Tables.Add, "CARTONS", "**", 0, False, "V")

            ASCMAIN1.sql = "Select * from EDT945T1 WHERE EDI_SHIPMENT_ID = :PARM1"
            Create_TDA(.Tables.Add, "EDT945T1_X", "**", 0, False, "V")

            Create_TDA(.Tables.Add("EDT945T2_X"), "EDT945T2", "*", 1)

            ASCMAIN1.sql = "SELECT * FROM CONV.CFG_SHIPHDR WHERE OHINVN = :PARM1"
            Create_TDA(.Tables.Add, "SHIPHDR_LK", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "SELECT SHIPDTL.*, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC FROM CONV.CFG_SHIPDTL SHIPDTL, ICTITEM1 WHERE SHIPDTL.SAINVN = :PARM1 AND SHIPDTL.SAITEM = ICTITEM1.ITEM_ALT_SORT (+)"
            Create_TDA(.Tables.Add, "SHIPDTL_LK", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "SELECT CARTON.*, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC FROM CONV.CFG_CARTON CARTON, ICTITEM1 WHERE CARTON.CHINVN = :PARM1 AND CARTON.CDITEM = ICTITEM1.ITEM_ALT_SORT (+)"
            Create_TDA(.Tables.Add, "CARTON_LK", "**", 0, False, "V", 0)

            ASCMAIN1.sql = "SELECT * FROM CONV.CFG_TRACK WHERE INVN = :PARM1"
            Create_TDA(.Tables.Add, "TRACK_LK", "**", 0, False, "V", 0)

            Create_TDA(.Tables.Add, "SOTSHIPB", "*")

            ASCMAIN1.sql = "SELECT * FROM SOTSHIP1 WHERE BILL_OF_LADING_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTSHIP1", ASCMAIN1.sql, 0, False, "V", 1)

        End With

        grdEDT945T1.DataSource = dst.Tables("EDT945T1")
        grdEDT945T2.DataSource = dst.Tables("EDT945T2")

        grdSOTPICKX.DataSource = dst.Tables("SOTPICKX")

        grdSOTPICK1_X.DataSource = dst.Tables("SOTPICK1_X")
        grdSOTPICK2_X.DataSource = dst.Tables("SOTPICK2_X")

        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")

        'Create_Relation("SHIPHDR", "CARTONS", "OHINVN", "CHINVN")
        grdSHIPHDR.DataSource = dst.Tables("SHIPHDR")
        grdCARTONS.DataSource = dst.Tables("CARTONS")

        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdSOTSHIPB.DataSource = dst.Tables("SOTSHIPB")

        grdEDT945T1_X.DataSource = dst.Tables("EDT945T1_X")
        grdEDT945T2_X.DataSource = dst.Tables("EDT945T2_X")

        Create_Summary(grdSOTPICKX, "SHIP_BOL_NO", "Count")
        Create_Summary(grdSOTPICK1_X, "PICK_NO", "Count")
        Create_Summary(grdSHIPHDR, "OHCONO", "Count")
        Create_Summary(grdEDT945T1_X, "EDI_DOC_SEQ_NO", "Count")

        Create_Summary(grdEDT945T1, "EDI_PICK_NO", "Count")
        Create_Summary(grdEDT945T2, "EDI_DOC_SEQ_NO", "Count")

        grdSHIPHDR_LK.DataSource = dst.Tables("SHIPHDR_LK")
        grdSHIPDTL_LK.DataSource = dst.Tables("SHIPDTL_LK")
        grdCARTON_LK.DataSource = dst.Tables("CARTON_LK")
        grdTRACK_LK.DataSource = dst.Tables("TRACK_LK")

        For Each grdcol As Infragistics.Win.UltraWinGrid.UltraGridColumn In grdEDT945T1.DisplayLayout.Bands(0).Columns
            grdcol.CellActivation = UltraWinGrid.Activation.NoEdit
        Next

        grdEDT945T1.DisplayLayout.UseFixedHeaders = True
        Dim iloop As Int16 = 0

        With grdEDT945T1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"EDI_PICK_NO", "EDI_BOL_NO", "EDI_MASTER_BOL_NO"}
                .Columns(COLUMN_NAME).Header.VisiblePosition = iloop
                iloop += 1
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            .Columns("EDI_MASTER_BOL_NO").CellAppearance.BackColor = Drawing.Color.LightGreen
            .Columns("EDI_BOL_NO").CellAppearance.BackColor = Drawing.Color.LightGreen
            .Columns("EDI_PICK_NO").Header.Caption = "Pick Ticket"
            .Columns("EDI_BOL_NO").Header.Caption = "BOL No"
            .Columns("EDI_MASTER_BOL_NO").Header.Caption = "Master Bol No"
        End With

        optProcess.CheckedIndex = 0
        optProcess.Enabled = ASCMAIN1.Running_in_VS

        splEDT945T1.Dock = DockStyle.Fill
        splEDT945T1.Visible = False

        splData.Dock = DockStyle.Fill
        splData.Visible = False

        splClarinsInvoice.Dock = DockStyle.Fill
        splClarinsInvoice.Visible = False

        splBolNo.Dock = DockStyle.Fill
        splBolNo.Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = String.Empty

        Select Case eItemKey

            Case "Load"

                Select Case optProcess.Value
                    Case "C"
                        Absx1.txtFor("EDI_SHIPMENT_ID").Text = Absx1.txtFor("EDI_SHIPMENT_ID").Text.Trim
                        Dim OHINVN As String = Absx1.txtFor("EDI_SHIPMENT_ID").Text.Trim

                        If OHINVN.Length = 0 Then
                            EMsg &= vbCr & $"Provide the Clarins Invoice No."
                            Exit Select
                        End If

                        If Not IsNumeric(OHINVN) Then
                            EMsg &= vbCr & $"The Clarins Invoice No. needs to be numeric."
                            Exit Select
                        End If

                    Case "P"

                        If ASCMAIN1.USER_ID <> "edz" Then
                            EMsg &= vbCr & "You are not permitted to use this option."
                            Exit Sub
                        End If

                        PICK_NO = Absx1.txtFor("EDI_SHIPMENT_ID").Text.Trim
                        PICK_NO = ASCMAIN1.Format_Field(PICK_NO, "PICK_NO")
                        Absx1.txtFor("EDI_SHIPMENT_ID").Text = PICK_NO

                        Fill_Records("SOTPICK1", New Object() {PICK_NO})
                        If dst.Tables("SOTPICK1").Rows.Count = 0 Then
                            EMsg &= vbCr & "Cannot locate the provided Pick Ticket."
                            Exit Select
                        ElseIf dst.Tables("SOTPICK1").Rows(0).Item("PICK_STATUS") & String.Empty <> "P" Then
                            EMsg &= vbCr & "The provided Pick Ticket is not in Pick."
                            Exit Select
                        End If

                        ASCMAIN1.sql = $"Select * from EDT945T1 where EDI_PICK_NO = '{PICK_NO}' AND NVL(EDI_PROCESS_IND, '0') in ('0', '1')"
                        Fill_Records("EDT945T1", String.Empty, True, ASCMAIN1.sql)
                        If dst.Tables("EDT945T1").Rows.Count > 0 Then
                            EMsg &= vbCr & $"The provided Pick Ticket {PICK_NO} already exists in the EDI 945 table."
                            Exit Select
                        End If

                        Dim ORDR_NO As String = dst.Tables("SOTPICK1").Rows(0).Item("ORDR_NO") & String.Empty
                        Fill_Records("SOTORDR1", ORDR_NO)
                        If dst.Tables("SOTORDR1").Rows.Count = 0 Then
                            EMsg &= vbCr & $"Cannot locate the Sales Order {ORDR_NO} for the provided Pick Ticket {PICK_NO}."
                            Exit Select
                        End If

                        'EDI_DOC_SEQ_NO
                        If dst.Tables("SOTORDR1").Rows(0).Item("EDI_DOC_SEQ_NO") & String.Empty <> String.Empty Then
                            Dim zMsg As String = $"The Sales Order {ORDR_NO} for the provided Pick Ticket {PICK_NO} is an EDI Order. Do you want to continue?"
                            If MessageBox.Show(zMsg, "Edi Order", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                                Exit Sub
                            End If
                        End If

                    Case "S"
                        EDI_SHIPMENT_ID = Absx1.txtFor("EDI_SHIPMENT_ID").Text.Trim
                        EDI_SHIPMENT_ID = ASCMAIN1.Format_Field(EDI_SHIPMENT_ID, "SHIP_BOL_NO")

                        Absx1.txtFor("EDI_SHIPMENT_ID").Text = EDI_SHIPMENT_ID

                        Fill_Records("EDT945T1", New Object() {EDI_SHIPMENT_ID})
                        If dst.Tables("EDT945T1").Rows.Count = 0 Then
                            EMsg &= vbCr & "Cannot locate the provided Shipment ID."
                        ElseIf dst.Tables("EDT945T1").Select("ISNULL(EDI_PROCESS_IND, '0') <> '0'").Length > 0 Then
                            MessageBox.Show("The provided Shipment ID has been processed, you will not be able to modify the Master Bol No.", "Load", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Else
                            If Not ASCMAIN1.Logical_Lock("EDT945T1", EDI_SHIPMENT_ID) Then
                                Exit Sub
                            End If
                        End If

                    Case "B"
                        BOL_NO = Absx1.txtFor("EDI_SHIPMENT_ID").Text.Trim
                        BOL_NO = ASCMAIN1.Format_Field(BOL_NO, "BOL_NO")

                        Absx1.txtFor("EDI_SHIPMENT_ID").Text = BOL_NO
                        Fill_Records("SOTSHIPB", New Object() {BOL_NO})
                        If dst.Tables("SOTSHIPB").Rows.Count = 0 Then
                            EMsg &= vbCr & "Cannot locate the provided BOL No."
                        Else
                            If Not ASCMAIN1.Logical_Lock("SOTSHIPB", BOL_NO) Then
                                Exit Sub
                            End If
                        End If
                End Select

            Case "Update"

                If optProcess.Value = "B" Then
                    If MessageBox.Show("Do you want to change the BOL No to avoid a conflict in 3PL Shipment Confirmation?", "Update BOL No", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If
                    Exit Select
                End If

                If optProcess.Value = "P" Then
                    Dim lstErrors As New List(Of String)

                    ' Warn the user that the 945 details do not match the Pick Ticket Details.
                    For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
                        Dim ITEM_CODE As String = rowSOTPICK2.Item("ITEM_CODE") & String.Empty
                        Dim ptQty As Int32 = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", $"ITEM_CODE = '{ITEM_CODE}'") & String.Empty)
                        Dim ediQty As Int32 = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_SHIP_QTY)", $"STYLE_CODE = '{ITEM_CODE}'") & String.Empty)

                        If ptQty <> ediQty Then
                            lstErrors.Add($"Item {ITEM_CODE} has a Pick of {ptQty} and EDI 945 of {ediQty}.")
                        End If
                    Next

                    If lstErrors.Count > 0 Then
                        If MessageBox.Show($"WARNING: Discrepancies exist.{Environment.NewLine}{Environment.NewLine}{String.Join(Environment.NewLine, lstErrors.ToArray)}{Environment.NewLine}{Environment.NewLine}Do you want to continue?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                End If

                Dim tbl As DataTable = ASCDATA1.SelectDistinct("EDT945T1", "EDI_MASTER_BOL_NO")
                If tbl.Rows.Count > 1 Then
                    EMsg &= vbCr & "All Shipment Pick Tickets must have the same EDI MASTER BOL NO."
                Else
                    If MessageBox.Show("Do you want to Update changes?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If optProcess.Value <> "C" AndAlso optProcess.Value <> "B" Then
                    If MessageBox.Show("Do you want to Cancel changes?", "Cancel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode

                Select Case optProcess.Value
                    Case "B"
                        .Groups("Screen Control").Items("Update").Text = "Change BOL No"
                    Case Else
                        .Groups("Screen Control").Items("Update").Text = "Update"
                End Select
            End With
        End If

        If ScreenMode Then
            If dst.Tables("EDT945T1").Select("ISNULL(EDI_PROCESS_IND, '0') <> '0'").Length > 0 OrElse optProcess.Value = "C" Then
                UltraExplorerBar1.Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
            End If
            splData.Visible = False
        Else
            Clear_Record()
            ASCMAIN1.Progress("Loading Pick Tickets in Pick", "")
            Fill_Records("SOTPICKX")
            splData.Visible = True
            ASCMAIN1.Progress("", "")
        End If

        Set_Read_Only(grdHeader, ScreenMode)

        splEDT945T1.Visible = ",P,S".Contains(optProcess.Value) AndAlso ScreenMode
        splClarinsInvoice.Visible = optProcess.Value = "C" AndAlso ScreenMode
        splBolNo.Visible = optProcess.Value = "B" AndAlso ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)

        For Each table As String In New String() {"EDT945T1", "EDT945T2",
            "SOTPICK1", "SOTPICK2", "SOTORDR1", "SOTORDR2",
            "SOTPICKX", "SOTPICK1_X", "SOTPICK2_X",
            "SHIPHDR", "CARTONS", "EDT945T1_X", "EDT945T2_X",
            "SHIPHDR_LK", "SHIPDTL_LK", "CARTON_LK", "TRACK_LK",
            "SOTSHIPB", "SOTSHIP1"}
            dst.Tables(table).Rows.Clear()
        Next

        txtEDI_SHIPMENT_ID.Clear()

        If optProcess.Enabled = False Then
            optProcess.Value = "S"
        End If

        grdSHIPHDR.Text = ""
        grdSOTPICK1_X.Text = ""
        grdSOTPICK2_X.Text = ""
        grdEDT945T1_X.Text = ""
        grdEDT945T2_X.Text = ""

        splPick945.Panel2Collapsed = True

        EnforceConstraints(True)

        BOL_NO = String.Empty
        EDI_SHIPMENT_ID = String.Empty
        PICK_NO = String.Empty

    End Sub

    Private Sub Load_Record()

        Static ResizeData As Boolean = False

        Try
            ASCMAIN1.Progress("Loading EDI 945 data")

            Select Case optProcess.Value
                Case "P"
                    GenerateEDT945Records(PICK_NO)

                Case "S"
                    If dst.Tables("EDT945T1").Select("ISNULL(EDI_PROCESS_IND, '0') <> '0'").Length > 0 Then
                        grdEDT945T1.DisplayLayout.Bands(0).Columns("EDI_MASTER_BOL_NO").CellActivation = UltraWinGrid.Activation.NoEdit
                        grdEDT945T1.DisplayLayout.Bands(0).Columns("EDI_BOL_NO").CellActivation = UltraWinGrid.Activation.NoEdit
                    Else
                        grdEDT945T1.DisplayLayout.Bands(0).Columns("EDI_MASTER_BOL_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                        grdEDT945T1.DisplayLayout.Bands(0).Columns("EDI_BOL_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If

                    grdEDT945T1.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

                    ASCMAIN1.sql = "SELECT * FROM EDT945T2 WHERE EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM EDT945T1 WHERE EDI_SHIPMENT_ID = '" & EDI_SHIPMENT_ID & "')"
                    Fill_Records("EDT945T2", String.Empty, True, ASCMAIN1.sql)

                Case "C"
                    Fill_Records("SHIPHDR_LK", New Object() {txtEDI_SHIPMENT_ID.Text})
                    Fill_Records("SHIPDTL_LK", New Object() {txtEDI_SHIPMENT_ID.Text})
                    Fill_Records("CARTON_LK", New Object() {txtEDI_SHIPMENT_ID.Text})
                    Fill_Records("TRACK_LK", New Object() {txtEDI_SHIPMENT_ID.Text})

                    'grdSHIPDTL_LK.DisplayLayout.UseFixedHeaders = True
                    'With grdSHIPDTL_LK.DisplayLayout.Bands(0)
                    '    .Columns("SAITEM").Header.Fixed = True
                    '    .Columns("ITEM_CODE").Header.Fixed = True
                    '    .Columns("ITEM_DESC").Header.Fixed = True
                    'End With

                    'grdCARTON_LK.DisplayLayout.UseFixedHeaders = True
                    'With grdCARTON_LK.DisplayLayout.Bands(0)
                    '    .Columns("CDITEM").Header.Fixed = True
                    '    .Columns("ITEM_CODE").Header.Fixed = True
                    '    .Columns("ITEM_DESC").Header.Fixed = True
                    'End With

                    Sort_grdColumns(grdSHIPDTL_LK, "SAITEM")
                    Sort_grdColumns(grdCARTON_LK, "CDITEM")

                    If Not ResizeData Then
                        grdSHIPHDR_LK.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                        grdSHIPDTL_LK.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                        grdCARTON_LK.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                        grdTRACK_LK.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
                        ResizeData = True

                        grdSHIPDTL_LK.DisplayLayout.Bands(0).Columns("SAITEM").Header.VisiblePosition = 0
                        grdSHIPDTL_LK.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.VisiblePosition = 1
                        grdSHIPDTL_LK.DisplayLayout.Bands(0).Columns("ITEM_DESC").Header.VisiblePosition = 2

                        grdCARTON_LK.DisplayLayout.Bands(0).Columns("CDITEM").Header.VisiblePosition = 0
                        grdCARTON_LK.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.VisiblePosition = 1
                        grdCARTON_LK.DisplayLayout.Bands(0).Columns("ITEM_DESC").Header.VisiblePosition = 2
                    End If

                Case "B"
                    Fill_Records("SOTSHIP1", New Object() {BOL_NO})

            End Select

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            Update_Record_TDA("EDT945T1")
            Dim userMessage As String = "Update Successful"

            For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("")
                Dim EDI_BOL_NO As String = rowEDT945T1.Item("EDI_BOL_NO") & String.Empty
                Dim EDI_PICK_NO As String = rowEDT945T1.Item("EDI_PICK_NO") & String.Empty
                Dim EDI_SHIPMENT_ID As String = rowEDT945T1.Item("EDI_SHIPMENT_ID") & String.Empty
                Dim sql As String = "UPDATE CONV.CFG_SHIPHDR SET OHMANN = :PARM1 WHERE ABSPICKNBR = :PARM2 AND SHIP_BOL_NO = :PARM3"
                ASCDATA1.ExecuteSQL(sql, "VVV", New Object() {EDI_BOL_NO, EDI_PICK_NO, EDI_SHIPMENT_ID})
            Next

            If optProcess.Value = "P" Then
                Dim EDI_DTL_SEQ As Int16 = 1
                For Each rowEDT945T2 As DataRow In dst.Tables("EDT945T2").Select("", "EDI_DTL_SEQ")
                    rowEDT945T2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                    EDI_DTL_SEQ += 1
                Next
                Update_Record_TDA("EDT945T2")

            ElseIf optProcess.Value = "B" Then
                Dim BOL_NO_NEW As String = String.Empty

                For iChar As Int16 = 65 To 90
                    Dim prefix As String = Microsoft.VisualBasic.Strings.Chr(iChar)
                    BOL_NO_NEW = prefix & BOL_NO.Substring(1)
                    Dim rowSOTSHIPB As DataRow = LookUp("SOTSHIPB", BOL_NO_NEW)
                    If rowSOTSHIPB Is Nothing Then
                        Exit For
                    End If
                    BOL_NO_NEW = String.Empty
                Next

                If BOL_NO_NEW.LENGTH > 0 Then
                    Dim sql As String = "UPDATE SOTSHIPB SET BOL_NO = :PARM1 WHERE BOL_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(sql, "VV", {BOL_NO_NEW, BOL_NO})

                    sql = "UPDATE SOTSHIP1 SET BILL_OF_LADING_NO = :PARM1 WHERE BILL_OF_LADING_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(sql, "VV", {BOL_NO_NEW, BOL_NO})

                    userMessage = $"The New BOL No is {BOL_NO_NEW}"
                Else
                    Throw New Exception($"Could not Update BOL No {BOL_NO}")
                End If
            End If

            CommitTrans(userMessage)

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTPICKX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSHIPHDR, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

        Load_Popup_Menu(grdSOTPICK1_X, "SSSPB", "Show Filter", "Show GroupBox", "Show Pins", "Generate EDI 945")
        Load_Popup_Menu(grdSOTPICK2_X, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

        Load_Popup_Menu(grdEDT945T1_X, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Set to Process")
        Load_Popup_Menu(grdEDT945T2_X, "SSS", "Show Filter", "Show GroupBox", "Show Pins")

        Load_Popup_Menu(grdEDT945T1, "B", "Set All to this BOL")

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

        Select Case grd.Name

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTPICK1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If e.Tool.OwningMenu IsNot Nothing Then
            grd = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        End If
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key

            Case "Generate EDI 945"
                Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Value & String.Empty
                optProcess.Value = "P"
                Absx1.txtFor("EDI_SHIPMENT_ID").Text = PICK_NO
                Click_Command("Load")

            Case "Set to Process"
                Dim EDI_DOC_SEQ_NO = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty
                Dim EDI_PROCESS_IND = grd.ActiveRow.Cells("EDI_PROCESS_IND").Value & String.Empty

                If EDI_PROCESS_IND = "1" OrElse EDI_PROCESS_IND = "3" Then
                    MessageBox.Show("Document does not have the proper Process Ind value", "Set to Process", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Dim rowEDT945T1 As DataRow = LookUp("EDT945T1", EDI_DOC_SEQ_NO)
                If rowEDT945T1 Is Nothing Then
                    MessageBox.Show("Could NOT locate the EDI Document in Oracle", "Set to Process", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                EDI_PROCESS_IND = rowEDT945T1.Item("PROCESS_IND") & String.Empty
                If EDI_PROCESS_IND = "1" OrElse EDI_PROCESS_IND = "3" Then
                    MessageBox.Show("Document does not have the proper Process Ind value", "Set to Process", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                ASCMAIN1.sql = "Update EDT945T1 SET EDI_PROCESS_IND = '1' WHERE EDI_DOC_SEQ_NO = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {EDI_DOC_SEQ_NO})
                grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value = "1"

            Case "Set All to this BOL"
                Dim EDI_BOL_NO As String = grd.ActiveRow.Cells("EDI_BOL_NO").Value & String.Empty
                Dim EDI_MASTER_BOL_NO As String = grd.ActiveRow.Cells("EDI_MASTER_BOL_NO").Value & String.Empty
                If EDI_BOL_NO.Length > 0 Then
                    If MessageBox.Show($"Do you want to set all recored to have BOL: {EDI_BOL_NO} and Master BOL: {EDI_MASTER_BOL_NO}?", "Set All to this BOL", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        Exit Sub
                    End If

                    For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("")
                        rowEDT945T1.Item("EDI_BOL_NO") = EDI_BOL_NO
                        rowEDT945T1.Item("EDI_MASTER_BOL_NO") = EDI_MASTER_BOL_NO
                    Next
                End If

        End Select
    End Sub

#End Region

    Private Sub txtEDI_SHIPMENT_ID_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEDI_SHIPMENT_ID.KeyPress
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Return) Then
            Click_Command("Load")
            e.Handled = True
        End If
    End Sub

    Private Sub grdEDT945T1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDT945T1.AfterRowActivate

        Dim viewEDT945T2 As New DataView(dst.Tables("EDT945T2"))

        Try
            Dim EDI_DOC_SEQ_NO As String = grdEDT945T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value

            viewEDT945T2.RowFilter = "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            viewEDT945T2.Sort = "EDI_DTL_SEQ"

        Catch ex As Exception
            viewEDT945T2.RowFilter = "EDI_DOC_SEQ_NO = '****'"
        End Try

        grdEDT945T2.DataSource = viewEDT945T2

    End Sub

    Private Sub optProcess_ValueChanged(sender As Object, e As EventArgs) Handles optProcess.ValueChanged

        lblType.Text = optProcess.CheckedItem.DisplayText

        'Select Case optProcess.Value
        '    Case "P"
        '        lblType.Text = "Pick Ticket"
        '    Case "S"
        '        lblType.Text = "Shipment ID"
        '    Case "C"
        '        lblType.Text = "Shipment ID"
        'End Select
    End Sub

    Private Sub grdSOTPICKX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICKX.AfterRowActivate

        If grdSOTPICKX.ActiveRow Is Nothing Then
            Exit Sub
        End If

        If Not grdSOTPICKX.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Dim SHIP_BOL_NO As String = grdSOTPICKX.ActiveRow.Cells("SHIP_BOL_NO").Value & String.Empty

        EnforceConstraints(False)
        dst.Tables("CARTONS").Rows.Clear()
        dst.Tables("SOTPICK2_X").Rows.Clear()
        dst.Tables("EDT945T2_X").Rows.Clear()

        grdSHIPHDR.Text = ""
        grdSOTPICK1_X.Text = "Pick Tickets"
        grdSOTPICK2_X.Text = "Pick Ticket Details"
        grdEDT945T1_X.Text = "EDI 945 Headers"
        grdEDT945T2_X.Text = "EDI 945 Details"
        grdCARTONS.Text = "Clarins Carton Data"

        Fill_Record("SOTPICK1_X", SHIP_BOL_NO)
        Fill_Record("SHIPHDR", SHIP_BOL_NO)
        Fill_Record("EDT945T1_X", SHIP_BOL_NO)

        EnforceConstraints(True)

        For Each rowSOTPICK1_X As DataRow In dst.Tables("SOTPICK1_X").Select("")
            Dim PICK_NO As String = rowSOTPICK1_X.Item("PICK_NO") & String.Empty
            If dst.Tables("SHIPHDR").Select($"ABSPICKNBR = '{PICK_NO}'").Length = 0 Then
                rowSOTPICK1_X.Item("MISSING") = 1
            Else
                rowSOTPICK1_X.Item("MISSING") = 0
            End If
        Next

        Sort_grdColumns(grdSOTPICK1_X, "MISSING")
        Sort_grdColumns(grdSHIPHDR, "ABSPICKNBR")

        grdSHIPHDR.Text = $"Clarins data for Shipment {SHIP_BOL_NO}"
        grdEDT945T1_X.Text = $"EDI 945 records for Shipment {SHIP_BOL_NO}"
        grdSOTPICK1_X.Text = $"Pick Tickets for Shipment {SHIP_BOL_NO}"

        If grdSHIPHDR.Rows.Count > 0 Then
            grdSHIPHDR.Selected.Rows.Add(grdSHIPHDR.Rows(0))
        End If

        Try
            If grdSOTPICKX.Rows.Count > 0 Then
                grdSOTPICKX.Selected.Rows.Clear()
                grdSOTPICKX.Selected.Rows.Add(grdSOTPICKX.Rows(0))
            End If

        Catch ex As Exception

        End Try

        If grdEDT945T1_X.Rows.Count > 0 Then
            grdEDT945T1_X.Selected.Rows.Add(grdEDT945T1_X.Rows(0))
        End If

        If grdSOTPICK1_X.Rows.Count > 0 Then
            grdSOTPICK1_X.Selected.Rows.Add(grdSOTPICK1_X.Rows(0))
        End If


    End Sub

    Private Sub grdSOTPICK1_X_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTPICK1_X.AfterRowActivate
        If grdSOTPICK1_X.ActiveRow Is Nothing Then
            Exit Sub
        End If

        Dim PICK_NO As String = grdSOTPICK1_X.ActiveRow.Cells("PICK_NO").Value & String.Empty
        Fill_Record("SOTPICK2_X", PICK_NO)

        grdSOTPICK2_X.Text = $"Pick Ticket Details for Pick No {PICK_NO}"
    End Sub

    Private Sub grdEDT945T1_X_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDT945T1_X.AfterRowActivate
        If grdEDT945T1_X.ActiveRow Is Nothing Then
            Exit Sub
        End If

        Dim EDI_DOC_SEQ_NO As String = grdEDT945T1_X.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty
        Fill_Record("EDT945T2_X", EDI_DOC_SEQ_NO)

        grdEDT945T2_X.Text = $"EDI 945 detail for Document {EDI_DOC_SEQ_NO}"

    End Sub

    Private Sub GenerateEDT945Records(ByVal PICK_NO As String)

        Fill_Records("SOTPICK1", PICK_NO)
        Fill_Records("SOTPICK2", PICK_NO)

        dst.Tables("EDT945T1").Rows.Clear()
        dst.Tables("EDT945T2").Rows.Clear()

        ASCMAIN1.sql = $"SELECT EDT945T1.* FROM EDT945T1, SOTPICK1
                            WHERE EDT945T1.EDI_SHIPMENT_ID = SOTPICK1.SHIP_BOL_NO
                            AND SOTPICK1.PICK_NO = '{PICK_NO}'"
        Dim rowShipment As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)

        Fill_Records("EDT945T1", String.Empty, True, $"Select * From EDT945T1 where EDI_PICK_NO = '{PICK_NO}' AND EDI_PROCESS_IND = '0'")
        Dim rowEDT945T1 As DataRow = Nothing
        Dim EDI_DOC_SEQ_NO As String = Nothing

        If dst.Tables("EDT945T1").Select($"EDI_PICK_NO = '{PICK_NO}'").Length > 0 Then
            rowEDT945T1 = dst.Tables("EDT945T1").Select($"EDI_PICK_NO = '{PICK_NO}'")(0)
            EDI_DOC_SEQ_NO = rowEDT945T1.Item("EDI_DOC_SEQ_NO") & String.Empty
        End If

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("", "PICK_NO")
            'Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO") & String.Empty
            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO") & String.Empty
            Dim SHIP_BOL_NO As String = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty

            Fill_Records("SOTORDR1", ORDR_NO)
            Fill_Records("SOTORDR2", ORDR_NO)
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)
            Dim EDI_DTL_SEQ As Int16 = 0

            If rowEDT945T1 Is Nothing Then
                EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDI_DOC_SEQ_NO")
                EDI_DOC_SEQ_NO = "1" & EDI_DOC_SEQ_NO.Substring(1) ' Leading 1 so no collisions with Gentran

                rowEDT945T1 = dst.Tables("EDT945T1").NewRow
                rowEDT945T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO

                'rowEDT945T1.ITEM("GEN_DOC_NO") = String.Empty
                'rowEDT945T1.ITEM("EDI_ISA_NO") = String.Empty
                'rowEDT945T1.ITEM("EDI_TP_QUAL") = String.Empty
                'rowEDT945T1.ITEM("EDI_TP_ID") = String.Empty
                'rowEDT945T1.ITEM("EDI_OUR_QUAL") = String.Empty
                'rowEDT945T1.ITEM("EDI_OUR_ID") = String.Empty
                'rowEDT945T1.ITEM("EDI_REPORTING_CODE") = String.Empty

                Dim EDI_SHIPMENT_DATE As String = String.Empty
                While Not IsDate(EDI_SHIPMENT_DATE)
                    EDI_SHIPMENT_DATE = InputBox("Enter Ship Date", "Shipment Date")
                End While

                rowEDT945T1.Item("EDI_SHIPMENT_DATE") = CDate(EDI_SHIPMENT_DATE).ToShortDateString
                rowEDT945T1.Item("EDI_SHIPMENT_ID") = rowSOTPICK1.Item("SHIP_BOL_NO")
                rowEDT945T1.Item("EDI_BOL_NO") = rowSOTPICK1.Item("SHIP_BOL_NO")
                rowEDT945T1.Item("EDI_MASTER_BOL_NO") = rowSOTPICK1.Item("SHIP_BOL_NO")
                rowEDT945T1.Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")

                rowEDT945T1.Item("EDI_PICK_NO") = rowSOTPICK1.Item("PICK_NO") & String.Empty

                rowEDT945T1.Item("EDI_ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
                rowEDT945T1.Item("EDI_DIVISION_CODE") = ASCMAIN1.DBS_COMPANY
                'rowEDT945T1.Item("EDI_BOL_NO") = rowHeader.Item("OHMANN")
                'rowEDT945T1.Item("EDI_MASTER_BOL_NO") = rowHeader.Item("OHMANN")

                'rowEDT945T1.Item("EDI_SHIPPER_ID_NO") = dst.Tables("TRACK").Compute("MAX(PROORTRK)", "[INVN] = '" & INVOICE_NUMBER & "'") & String.Empty
                ' 12/29/2015 as per Maria
                'rowEDT945T1.Item("SHIP_PICKUP_NO") = dst.Tables("TRACK").Compute("MAX([PUN])", "[INVN] = '" & INVOICE_NUMBER & "'") & String.Empty
                'rowEDT945T1.Item("SHIP_AUTH_NO") = dst.Tables("TRACK").Compute("MAX([AUTHN])", "[INVN] = '" & INVOICE_NUMBER & "'") & String.Empty

                rowEDT945T1.Item("EDI_FRT_COST") = 0
                rowEDT945T1.Item("EDI_ORDR_SHIP_DATE") = rowEDT945T1.Item("EDI_SHIPMENT_DATE")
                'rowEDT945T1.ITEM("EDI_TRANS_METH_CODE") = String.Empty

                Dim SHIP_VIA_CODE As String = rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty
                Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                If rowSOTSVIA1 IsNot Nothing Then
                    rowEDT945T1.Item("EDI_CARRIER_SCAC_CODE") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
                    rowEDT945T1.Item("EDI_CARRIER_NAME") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    rowEDT945T1.Item("EDI_FRT_TERMS") = rowSOTSVIA1.Item("FRT_TERMS") & String.Empty
                End If
                rowEDT945T1.Item("EDI_CARRIER_CODE") = SHIP_VIA_CODE

                rowEDT945T1.Item("EDI_TOTAL_UNITS_SHIPPED") = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", $"PICK_NO = '{PICK_NO}'") & String.Empty)
                rowEDT945T1.Item("EDI_TOTAL_ORDR_WEIGHT") = 0
                rowEDT945T1.Item("EDI_RECEIVED_DATE") = DateTime.Now
                rowEDT945T1.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                'rowEDT945T1.ITEM("EDI_TRAILER_NO") = String.Empty
                'rowEDT945T1.Item("EDI_LOAD_ID") = String.Empty
                dst.Tables("EDT945T1").Rows.Add(rowEDT945T1)
            End If

            If rowShipment IsNot Nothing Then
                For Each field As String In New String() {"EDI_BOL_NO", "EDI_MASTER_BOL_NO", "EDI_CARRIER_NAME", "EDI_CARRIER_CODE",
                        "EDI_CARRIER_SCAC_CODE", "EDI_RECEIVED_DATE", "SHIP_PICKUP_NO", "SHIP_AUTH_NO"}
                    rowEDT945T1.Item(field) = rowShipment.Item(field)
                Next
            End If

            Dim tblCarton As DataTable = Nothing
            Dim rowSHIPHDR As DataRow = ASCDATA1.GetDataRow("SELECT * FROM CONV.CFG_SHIPHDR WHERE ABSPICKNBR = :PARM1", "V", New Object() {PICK_NO})
            If rowSHIPHDR IsNot Nothing Then
                Dim OHINVN As String = rowSHIPHDR.Item("OHINVN") & String.Empty
                If OHINVN.Length > 0 Then
                    tblCarton = ASCDATA1.GetDataTable($"SELECT CARTON.*, ICTITEM1.ITEM_CODE 
                                                                    FROM CONV.CFG_CARTON CARTON, ICTITEM1
                                                                    WHERE CARTON.CHINVN = '{OHINVN}'
                                                                    AND CARTON.CDITEM = ICTITEM1.ITEM_ALT_SORT (+)")
                End If
            End If

            If tblCarton Is Nothing OrElse tblCarton.Rows.Count = 0 Then
                Continue For
            End If

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND PICK_QTY > 0", "PICK_LNO")
                rowSOTPICK2.Item("PICK_QTY_CONF") = 0
            Next

            EDI_DTL_SEQ = 0
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND PICK_QTY > 0", "PICK_LNO")
                Dim ORDR_LNO As Int16 = rowSOTPICK2.Item("ORDR_LNO")
                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty

                For Each rowCarton As DataRow In tblCarton.Select($"ITEM_CODE = '{ITEM_CODE}'", "CHCTNN")
                    Dim rowEDT945T2 As DataRow = dst.Tables("EDT945T2").NewRow
                    rowEDT945T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                    EDI_DTL_SEQ += 1
                    rowEDT945T2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                    rowEDT945T2.Item("EDI_CART_NO") = rowCarton.Item("CPUCCN")
                    rowEDT945T2.Item("EDI_SHIPMENT_STATUS_CODE") = "SH"
                    rowEDT945T2.Item("PICK_LNO") = rowSOTPICK2.Item("PICK_LNO")
                    rowEDT945T2.Item("PICK_QTY") = Val(rowCarton.Item("CDQTYS") & String.Empty) * Val(rowCarton.Item("CDSKMQ") & String.Empty)

                    Dim EDI_PICK_QTY As Int32 = Val(rowCarton.Item("CDQTYS") & String.Empty) * Val(rowCarton.Item("CDSKMQ") & String.Empty)
                    rowSOTPICK2.Item("PICK_QTY_CONF") = EDI_PICK_QTY
                    rowEDT945T2.Item("EDI_SHIP_QTY") = EDI_PICK_QTY
                    'rowEDT945T2.Item("EDI_DIFF_QTY") = String.Empty
                    'rowEDT945T2.Item("STYLE_UOM") = String.Empty
                    'rowEDT945T2.Item("UPC_CODE") = String.Empty
                    rowEDT945T2.Item("STYLE_CODE") = rowSOTORDR2.Item("ITEM_CODE")
                    'rowEDT945T2.Item("STYLE_DESC") = String.Empty
                    'rowEDT945T2.Item("UNIT_PRICE") = String.Empty
                    'rowEDT945T2.Item("EDI_STYLE") = String.Empty
                    'rowEDT945T2.Item("EDI_SKU") = String.Empty
                    'rowEDT945T2.Item("EDI_PROD_CLASS") = String.Empty
                    'rowEDT945T2.Item("NMFC") = String.Empty
                    'rowEDT945T2.Item("EDI_STYLE_DESC") = String.Empty
                    'rowEDT945T2.Item("EDI_RETAIL_PRICE") = String.Empty
                    'rowEDT945T2.Item("EDI_SIZE") = String.Empty
                    'rowEDT945T2.Item("EDI_COLOR") = String.Empty
                    'rowEDT945T2.Item("PACK_QTY") = String.Empty
                    'rowEDT945T2.Item("PACK_SIZE") = String.Empty
                    'rowEDT945T2.Item("PACK_UOM") = String.Empty
                    'rowEDT945T2.Item("EDI_ORDR_CUBE") = String.Empty
                    'rowEDT945T2.Item("EDI_CART_LENGTH") = rowCarton.Item("CARTON_LENGTH")
                    'rowEDT945T2.Item("EDI_CART_HEIGHT") = rowCarton.Item("CARTON_HEIGHT")
                    'rowEDT945T2.Item("EDI_CART_WIDTH") = rowCarton.Item("CARTON_WIDTH")
                    'rowEDT945T2.Item("EDI_SHIPPER_ID_NO") = rowCarton.Item("TRACKING_ID")
                    rowEDT945T2.Item("EDI_SHIPPER_ID_NO") = String.Empty
                    rowEDT945T2.Item("EDI_CART_WEIGHT") = Val(rowCarton.Item("CDWGHT") & String.Empty)
                    dst.Tables("EDT945T2").Rows.Add(rowEDT945T2)
                Next
            Next
        Next

        Dim manualEntry As Boolean = False

        ' Us the pick ticket if not Clarins data
        If dst.Tables("EDT945T2").Rows.Count = 0 Then
            Dim EDI_DTL_SEQ As Int16 = 0
            EDI_DOC_SEQ_NO = dst.Tables("EDT945T1").Rows(0).Item("EDI_DOC_SEQ_NO")
            manualEntry = True

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select($"PICK_NO = '{PICK_NO}' AND PICK_QTY > 0", "PICK_LNO")
                Dim ORDR_NO As String = rowSOTPICK2.Item("ORDR_NO")

                Dim ORDR_LNO As Int16 = rowSOTPICK2.Item("ORDR_LNO")
                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty

                Dim PICK_QTY As Int16 = Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty)

                Dim rowEDT945T2 As DataRow = dst.Tables("EDT945T2").NewRow
                rowEDT945T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                EDI_DTL_SEQ += 1
                rowEDT945T2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                rowEDT945T2.Item("EDI_CART_NO") = "CTN-" & ASCMAIN1.Next_Control_No("SOTCART1.CART_NO")
                rowEDT945T2.Item("EDI_SHIPMENT_STATUS_CODE") = "SH"
                rowEDT945T2.Item("PICK_LNO") = rowSOTPICK2.Item("PICK_LNO")
                rowEDT945T2.Item("PICK_QTY") = PICK_QTY

                rowSOTPICK2.Item("PICK_QTY_CONF") = PICK_QTY
                rowEDT945T2.Item("EDI_SHIP_QTY") = PICK_QTY
                'rowEDT945T2.Item("EDI_DIFF_QTY") = String.Empty
                'rowEDT945T2.Item("STYLE_UOM") = String.Empty
                'rowEDT945T2.Item("UPC_CODE") = String.Empty
                rowEDT945T2.Item("STYLE_CODE") = rowSOTORDR2.Item("ITEM_CODE")
                'rowEDT945T2.Item("STYLE_DESC") = String.Empty
                'rowEDT945T2.Item("UNIT_PRICE") = String.Empty
                'rowEDT945T2.Item("EDI_STYLE") = String.Empty
                'rowEDT945T2.Item("EDI_SKU") = String.Empty
                'rowEDT945T2.Item("EDI_PROD_CLASS") = String.Empty
                'rowEDT945T2.Item("NMFC") = String.Empty
                'rowEDT945T2.Item("EDI_STYLE_DESC") = String.Empty
                'rowEDT945T2.Item("EDI_RETAIL_PRICE") = String.Empty
                'rowEDT945T2.Item("EDI_SIZE") = String.Empty
                'rowEDT945T2.Item("EDI_COLOR") = String.Empty
                'rowEDT945T2.Item("PACK_QTY") = String.Empty
                'rowEDT945T2.Item("PACK_SIZE") = String.Empty
                'rowEDT945T2.Item("PACK_UOM") = String.Empty
                'rowEDT945T2.Item("EDI_ORDR_CUBE") = String.Empty
                'rowEDT945T2.Item("EDI_CART_LENGTH") = rowCarton.Item("CARTON_LENGTH")
                'rowEDT945T2.Item("EDI_CART_HEIGHT") = rowCarton.Item("CARTON_HEIGHT")
                'rowEDT945T2.Item("EDI_CART_WIDTH") = rowCarton.Item("CARTON_WIDTH")
                'rowEDT945T2.Item("EDI_SHIPPER_ID_NO") = rowCarton.Item("TRACKING_ID")
                rowEDT945T2.Item("EDI_SHIPPER_ID_NO") = String.Empty
                rowEDT945T2.Item("EDI_CART_WEIGHT") = 0
                dst.Tables("EDT945T2").Rows.Add(rowEDT945T2)
            Next

        End If

        For Each rowEDT945T1 In dst.Tables("EDT945T1").Select("")
            EDI_DOC_SEQ_NO = rowEDT945T1.Item("EDI_DOC_SEQ_NO")
            rowEDT945T1.Item("EDI_TOTAL_UNITS_SHIPPED") = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_SHIP_QTY)", $"EDI_DOC_SEQ_NO= '{EDI_DOC_SEQ_NO}'") & String.Empty)
            rowEDT945T1.Item("EDI_TOTAL_ORDR_WEIGHT") = Val(dst.Tables("EDT945T2").Compute("SUM(EDI_CART_WEIGHT)", $"EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}'") & String.Empty)
        Next

        grdEDT945T1.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdEDT945T2.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        grdSOTPICK2.DisplayLayout.PerformAutoResizeColumns(True, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

        splPick945.Panel2Collapsed = False

        If manualEntry Then
            MessageBox.Show("There were no details from Clarins; therefore, the EDI 945 used the Pick Ticket to ship complete", "EDI 945", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub grdSHIPHDR_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSHIPHDR.AfterRowActivate

        Dim OHINVN As String = grdSHIPHDR.ActiveRow.Cells("OHINVN").Value & String.Empty
        Fill_Record("CARTONS", OHINVN)
        grdCARTONS.Text = $"Clarins Cartons for Invoice {OHINVN}"

    End Sub

End Class