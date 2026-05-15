Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

' ISSUE-7043 Creation of ASNs instead of invoices 

Public Class POFSHIP1

    Private PO_SHIPMENT_NO As String
    Private CONTAINER_CTL_NO As String
    Private WHSE_CODE As String
    Private VEND_CODE As String
    Private rowPOTSHIP1 As DataRow

    Private Const VEND_CODE_DEFAULT As String = "IPSA"
    Private lst3PLs As New List(Of String) From {"ADS"}

    Private aswPOTSHIP1 As String = String.Empty
    Private aswPOTSHIP2 As String = String.Empty
    Private aswPOTSHIP3 As String = String.Empty

    Private processingImportedData As Boolean = False
    Private editableColumns As New List(Of String) From {"VESSEL_NAME", "CONTAINER_NO", "SHIP_DATE", "ETA_DATE", "WHSE_CODE", "BOL_NO", "SEAL_NO1", "SEAL_NO2", "SEAL_NO3", "SEAL_NO4", "SHIP_TYPE"}

    ' ISSUE-7230 ADS as the defaUlt warehouse
    'Private Const WHSE_CODE_DEFAULT As String = "ADS"
    Private ReadOnly Property WHSE_CODE_DEFAULT()
        Get
            Return ROWs("POTPARM1").Item("PO_PARM_WHSE_CODE") & String.Empty
        End Get
    End Property


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If ASCMAIN1.Running_in_VS Then
            ValidateTablesExist()
        End If

        With dst
            Get_PARM("POTPARM1")

            ASCMAIN1.sql = "Select * from POTSHIP1 where PO_SHIPMENT_STATUS = :PARM1"
            Create_TDA(.Tables.Add, "POTSHIP1_X", ASCMAIN1.sql, 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT * FROM POTSHIP2 WHERE PO_SHIPMENT_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "POTSHIP2_X", ASCMAIN1.sql, 0, False, "V")

            ASCMAIN1.sql = "SELECT POTSHIP3.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.VESSEL_NAME, ICTPINV1.PINV_NO, ICTPINV1.WHSE_CODE, ICTPINV1.VEND_CODE, ICTPINV1.PINV_STATUS
                                FROM POTSHIP3, ICTPINV1
                                WHERE PO_SHIPMENT_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))
                                AND POTSHIP3.INV_NUM = ICTPINV1.INV_NUM (+)
                                AND POTSHIP3.PACK_LIST_NO = ICTPINV1.PACK_LIST_NO (+)
                                AND POTSHIP3.CONTAINER_CTL_NO = ICTPINV1.CONTAINER_CTL_NO (+)"
            Create_TDA(.Tables.Add, "POTSHIP3_X", ASCMAIN1.sql, 0, False, "V")

            ASCMAIN1.sql = "SELECT POTASNM2.*, POTASNM1.CONTAINER_CTL_NO
                            FROM POTASNM1, POTASNM2
                            WHERE POTASNM1.ASN_NO = POTASNM2.ASN_NO
                            AND POTASNM1.ASN_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "POTASNM2_X", ASCMAIN1.sql, 0, False, "V")

            ASCMAIN1.sql = "SELECT POTASNM3.*, ICTITEM1.ITEM_DESC
                            FROM POTASNM3, ICTITEM1
                            WHERE ASN_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))
                            AND POTASNM3.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
            Create_TDA(.Tables.Add, "POTASNM3_X", ASCMAIN1.sql, 0, False, "V")

            Create_TDA(.Tables.Add, "POTASNM1", "*")
            Create_TDA(.Tables.Add, "POTASNM2", "*")
            Create_TDA(.Tables.Add, "POTASNM3", "*")

            Create_TDA(.Tables.Add, "POTSHIP1", "*")
            Create_TDA(.Tables.Add, "POTSHIP2", "*")
            Create_TDA(.Tables.Add, "POTSHIP3", "*")
            ' ICTPINV1.PO_ORDER_NO, ICTPINV1.VESSEL_NAME, ICTPINV1.PINV_NO, ICTPINV1.WHSE_CODE, ICTPINV1.VEND_CODE
            With dst.Tables("POTSHIP3")
                .Columns.Add("PO_ORDER_NO", GetType(String))
                .Columns.Add("VESSEL_NAME", GetType(String))
                .Columns.Add("PINV_NO", GetType(String))
                .Columns.Add("WHSE_CODE", GetType(String))
                .Columns.Add("VEND_CODE", GetType(String))
                .Columns.Add("WHSE_CODE_PO", GetType(String))
                .Columns.Add("PINV_STATUS", GetType(String))
            End With

            AUDIT.Add("POTSHIP1", "E")
            AUDIT.Add("POTSHIP2", "E")

            ASCMAIN1.sql = "SELECT ICTPINV1.*, POTORDR1.WHSE_CODE WHSE_CODE_PO
                            FROM ICTPINV1, POTORDR1
                            WHERE ICTPINV1.CONTAINER_CTL_NO IS NULL 
                            AND ICTPINV1.PINV_STATUS = 'O'
                            AND ICTPINV1.PO_ORDER_NO = POTORDR1.PO_ORDER_NO (+)
                            AND ICTPINV1.WHSE_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTPINV1", ASCMAIN1.sql, 0, True, "V")

            Create_TDA(.Tables.Add("POTSHIP1_IMP"), "POTSHIP1", "*")
            Create_TDA(.Tables.Add("POTSHIP2_IMP"), "POTSHIP2", "*")
            Create_TDA(.Tables.Add("POTSHIP3_IMP"), "POTSHIP3", "*")

            aswPOTSHIP1 = ASCMAIN1.Temp_Table("Select * from POTSHIP1 Where Rownum < 1")
            Create_TDA(.Tables.Add, aswPOTSHIP1, "*")

            aswPOTSHIP2 = ASCMAIN1.Temp_Table("Select * from POTSHIP2 Where Rownum < 1")
            Create_TDA(.Tables.Add, aswPOTSHIP2, "*")

            aswPOTSHIP3 = ASCMAIN1.Temp_Table("Select * from POTSHIP3 Where Rownum < 1")
            Create_TDA(.Tables.Add, aswPOTSHIP3, "*")

            Create_TDA(.Tables.Add("ICTWHSEL"), "ICTWHSEL", "*")
            Fill_Records("ICTWHSEL", String.Empty, True, "SELECT * FROM ICTWHSEL")
        End With

        Create_Relation("POTSHIP1_X", "POTSHIP2_X", "PO_SHIPMENT_NO", "PO_SHIPMENT_NO")
        Create_Relation("POTSHIP2_X", "POTSHIP3_X", "PO_SHIPMENT_NO,CONTAINER_CTL_NO", "PO_SHIPMENT_NO,CONTAINER_CTL_NO")

        Create_Relation("POTSHIP2_X", "POTASNM2_X", "CONTAINER_CTL_NO", "CONTAINER_CTL_NO")
        Create_Relation("POTASNM2_X", "POTASNM3_X", "ASN_NO,ASN_XMIT_NO", "ASN_NO,ASN_XMIT_NO")

        grdPOTSHIP1_X.DataSource = dst.Tables("POTSHIP1_X")
        Create_Summary(grdPOTSHIP1_X, "BOL_NO", "Count")
        Create_Summary(grdPOTSHIP1_X, "CONTAINER_NO", "Count", "POTSHIP1_X_POTSHIP2_X")
        Create_Summary(grdPOTSHIP1_X, "VEND_CODE", "Count", "POTSHIP2_X_POTSHIP3_X")

        ASCMAIN1.Add_Value_List(grdPOTSHIP1_X, "SHIP_TYPE", Nothing, {":", "B:Boat", "A:Air", "G:Ground"})
        ASCMAIN1.Add_Value_List(grdPOTSHIP1_X, "PO_SHIPMENT_STATUS", Nothing, {":", "O:Open", "I:Invoiced", "C:Closed, R:Received"})
        ASCMAIN1.Add_Value_List(grdPOTSHIP1_X, "PINV_STATUS", Nothing, {":", "O:Open", "C:Closed", "I:Invoiced", "R:Received"}, 2)

        grdICTPINV1.DataSource = dst.Tables("ICTPINV1")
        Create_Summary(grdICTPINV1, "PINV_NO", "Count")

        grdPOTSHIP3.DataSource = dst.Tables("POTSHIP3")
        Create_Summary(grdPOTSHIP3, "PINV_NO", "Count")
        ASCMAIN1.Add_Value_List(grdPOTSHIP3, "PINV_STATUS", Nothing, {":", "O:Open", "C:Closed", "I:Invoiced", "R:Received"})

        grdASTAUDT1.DataSource = dst.Tables("ASTAUDT1")
        For Each col As UltraGridColumn In grdASTAUDT1.DisplayLayout.Bands(0).Columns
            col.Header.Caption = StrConv(col.Header.Caption.Replace("_", " "), VbStrConv.ProperCase)
        Next

        dteETA_DATE.MinDate = CDate("01/01/2025")
        dteETA_DATE.MaxDate = DateAdd(DateInterval.Year, 3, DateTime.Now)

        dteSHIP_DATE.MinDate = dteETA_DATE.MinDate
        dteSHIP_DATE.MaxDate = dteETA_DATE.MaxDate

        TABLE_NAME = "POTSHIP1"
        Bind_Controls(grpDetails, "POTSHIP1")
        Bind_Controls(grpDetails, "POTSHIP2")
        Bind_Controls(UltraGroupBox1, "POTSHIP1")
        Bind_Controls(grpSeals, "POTSHIP1")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "View", "Edit"

                If PO_SHIPMENT_NO.Length = 0 Then
                    EMsg &= vbCr & "Double-click a container in the grid."
                    Exit Select
                End If

                If CONTAINER_CTL_NO.Length = 0 Then
                    EMsg &= vbCr & "Double-click a container in the grid."
                    Exit Select
                End If

                If eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then
                        Exit Sub
                    End If
                End If

                rowPOTSHIP1 = LookUp("POTSHIP1", PO_SHIPMENT_NO)

                If rowPOTSHIP1 Is Nothing Then
                    EMsg &= vbCr & "Missing or invalid PO Shipment No"
                ElseIf rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") & String.Empty = "C" AndAlso eItemKey = "Edit" Then
                    EMsg &= vbCr & "You may NOT Edit a Completed ASN"
                End If

                Dim drPOTSHIP2 As DataRow = LookUp("POTSHIP2", {PO_SHIPMENT_NO, CONTAINER_CTL_NO})
                If drPOTSHIP2 Is Nothing Then
                    EMsg &= vbCr & "Cannot locate the selected Container"
                End If

                If eItemKey = "Edit" Then
                    If EMsg.Length > 0 Then
                        ASCMAIN1.MultiTask_Release()
                    Else
                        ' See if an ASN was already sent
                        ASCMAIN1.sql = "SELECT POTASNM2.*
                                            FROM POTASNM1, POTASNM2
                                            WHERE POTASNM1.ASN_NO = POTASNM2.ASN_NO
                                            AND POTASNM1.CONTAINER_CTL_NO = :PARM1
                                            AND POTASNM2.TRANSMIT_DATE IS NOT NULL"
                        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {CONTAINER_CTL_NO})
                        If tbl.Rows.Count > 0 Then
                            EMsg &= vbCr & "You are not permitted to Edit a Container once an ASN has been sent."
                            ASCMAIN1.MultiTask_Release()
                        End If
                    End If
                End If

            Case "Done"

            Case "Update"

                If processingImportedData Then
                    ' Need to validate the data
                    For Each drPOTSHIP1 As DataRow In dst.Tables("POTSHIP1_X").Select("")
                        Dim PO_SHIPMENT_NO As String = drPOTSHIP1.Item("PO_SHIPMENT_NO") & String.Empty
                        Dim BOL_NO As String = drPOTSHIP1.Item("BOL_NO") & String.Empty
                        Dim VESSEL_NAME As String = drPOTSHIP1.Item("VESSEL_NAME") & String.Empty

                        BOL_NO = BOL_NO.Trim
                        If BOL_NO.Length = 0 Then
                            EMsg &= vbCr & "BOL No is required"
                        End If

                        Dim SHIP_DATE As String = drPOTSHIP1.Item("SHIP_DATE") & String.Empty
                        Dim ETA_DATE As String = drPOTSHIP1.Item("ETA_DATE") & String.Empty

                        If Not IsDate(SHIP_DATE) Then
                            EMsg &= vbCr & "Ship Date is not a valid date."
                        ElseIf Not IsDate(ETA_DATE) Then
                            EMsg &= vbCr & "ETA Date is not a valid date."
                        ElseIf CDate(SHIP_DATE) > CDate(ETA_DATE) Then
                            EMsg &= vbCr & "Ship date must be less equal ETA Date."
                        End If

                        VESSEL_NAME = VESSEL_NAME.Trim
                        If VESSEL_NAME.Length = 0 Then
                            EMsg &= vbCr & "Vessel is required"
                        End If

                        For Each drPOTSHIP2 As DataRow In dst.Tables("POTSHIP2_X").Select($"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'", "")
                            Dim CONTAINER_NO As String = drPOTSHIP2.Item("CONTAINER_NO") & String.Empty
                            CONTAINER_NO = CONTAINER_NO.Trim

                            Dim WHSE_CODE As String = drPOTSHIP2.Item("WHSE_CODE") & String.Empty
                            WHSE_CODE = WHSE_CODE.Trim

                            If CONTAINER_NO.Length = 0 Then
                                EMsg &= vbCr & "Container is required"
                            End If

                            If WHSE_CODE.Length = 0 Then
                                EMsg &= vbCr & "Container warehouse is required"
                            Else
                                Dim drICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                                If drICTWHSE1 Is Nothing Then
                                    EMsg &= vbCr & "Invalid or missing Container warehouse."
                                ElseIf drICTWHSE1.Item("LP_CODE") & String.Empty = String.Empty Then
                                    EMsg &= vbCr & "Container warehouses must be assigned assigned a 3PL Code."
                                ElseIf Not lst3PLs.Contains(drICTWHSE1.Item("LP_CODE") & String.Empty) Then
                                    EMsg &= vbCr & "Container warehouse must be an ADS warehouse."
                                End If
                            End If
                        Next

                        If EMsg.Length > 0 Then
                            Exit Select
                        End If
                    Next

                    If MessageBox.Show("Do you want to commit the imported LOA/BOL?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.No Then
                        Exit Sub
                    End If
                Else
                    If dst.Tables("POTSHIP3").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
                        EMsg &= vbCr & "You must have at least one detail line"
                    End If

                    If MyBase.Absx1.txtFor("CONTAINER_NO").TextLength = 0 Then
                        EMsg &= vbCr & "Container No is required"
                    End If
                    If MyBase.Absx1.txtFor("VESSEL_NAME").TextLength = 0 Then
                        EMsg &= vbCr & "Vessel Name is required"
                    End If

                    If Not MyBase.Absx1.dteFor("SHIP_DATE").IsDateValid Then
                        EMsg &= vbCr & "Ship Date is missing or invalid"
                    End If

                    If Not MyBase.Absx1.dteFor("ETA_DATE").IsDateValid Then
                        EMsg &= vbCr & "ETA Date is missing or invalid"
                    End If

                    If MyBase.Absx1.dteFor("SHIP_DATE").DateTime > MyBase.Absx1.dteFor("SHIP_DATE").DateTime Then
                        EMsg &= vbCr & "Ship Date must be less than ETA Date"
                    End If

                    If EMsg.Length = 0 Then
                        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                        Dim drICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

                        If drICTWHSE1 Is Nothing Then
                            EMsg &= vbCr & "Invalid or missing warehouse."
                            Exit Select
                        ElseIf drICTWHSE1.Item("LP_CODE") & String.Empty = String.Empty Then
                            EMsg &= vbCr & "The selected warehouse is not assigned a 3PL Code."
                            Exit Select
                        ElseIf Not lst3PLs.Contains(drICTWHSE1.Item("LP_CODE") & String.Empty) Then
                            EMsg &= vbCr & "The selected warehouse is not an ADS warehouse."
                            Exit Select
                        End If

                        ' Need to verify the PO is for LP Code ADS?????

                        If WHSE_CODE <> dst.Tables("POTSHIP2").Rows(0).Item("WHSE_CODE", DataRowVersion.Original) Then
                            Dim zMsg As String = $"You’ve changed the warehouse, so all Selected Invoices and their corresponding Purchase Orders will now be updated to use Warehouse Code {WHSE_CODE}."
                            zMsg &= Environment.NewLine & Environment.NewLine & "Do you want to proceed with this update?"

                            If MessageBox.Show(zMsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                                Exit Sub
                            End If
                        Else
                            If dst.Tables("POTSHIP3").Select($"WHSE_CODE <> '{WHSE_CODE}'").Length > 0 Then '  OR WHSE_CODE_PO <> '{WHSE_CODE}'
                                Dim zMsg As String = $"There are Selected Invoices or their corresponding Purchase Orders that are assigned a different warehouse. They will now be updated to use Warehouse Code {WHSE_CODE}."
                                zMsg &= Environment.NewLine & Environment.NewLine & "Do you want to proceed with this update?"

                                If MessageBox.Show(zMsg, "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                    Dim lstInvoices As New List(Of String)
                    Dim lstPOs As New List(Of String)
                    Dim tblICTPINV1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ICTPINV1 WHERE CONTAINER_CTL_NO = :PARM1", "", "V", {CONTAINER_CTL_NO})
                    For Each drICTPINV1 As DataRow In tblICTPINV1.Select("")
                        If Not lstInvoices.Contains(drICTPINV1.Item("PINV_NO")) Then
                            lstInvoices.Add(drICTPINV1.Item("PINV_NO"))
                        End If

                        If Not lstPOs.Contains(drICTPINV1.Item("PO_ORDER_NO")) Then
                            lstPOs.Add(drICTPINV1.Item("PO_ORDER_NO"))
                        End If
                    Next

                    For Each drPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
                        Dim PINV_NO As String = drPOTSHIP3.Item("PINV_NO") & String.Empty
                        Dim PO_ORDER_NO As String = drPOTSHIP3.Item("PO_ORDER_NO") & String.Empty

                        If Not lstInvoices.Contains(PINV_NO) Then
                            lstInvoices.Add(PINV_NO)
                        End If

                        If Not lstPOs.Contains(PO_ORDER_NO) Then
                            lstPOs.Add(PO_ORDER_NO)
                        End If
                    Next

                    For Each PINV_NO As String In lstInvoices
                        If Not ASCMAIN1.Logical_Lock("ICTPINV1", PINV_NO, False) Then
                            Exit Sub
                        End If
                    Next

                    For Each PO_ORDER_NO As String In lstPOs
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, False) Then
                            Exit Sub
                        End If
                    Next

                End If

            Case "Cancel"
                Dim qMsg As String = String.Empty
                If EntryMode = "N" Then
                    qMsg = "Do you want to cancel entry of a new PO Shipment?"
                Else
                    qMsg = "Do you want to cancel changes made to this PO Shipment?"
                End If

                If MessageBox.Show(qMsg, "Cancel Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

            Case "Import Manifest"

            Case "Create Air BOL"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                Update_Record()
                processingImportedData = False
                Mode_Settings(False)

            Case "Cancel"
                processingImportedData = False
                Mode_Settings(False)

            Case "Refresh"
                Refresh_Documents()

            Case "Import Manifest"
                ImportLOA()
                Refresh_Documents()

            Case "Create Air BOL"
                CreateAirBOL()
                Refresh_Documents()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                If (EntryMode = "V" AndAlso ScreenMode) Then
                    If rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                End If

                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Visible = False
                .Items("Done").Settings.Enabled = iScreenMode

                .Items("Edit").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Refresh").Settings.Enabled = not_iScreenMode
                .Items("Import Manifest").Settings.Enabled = not_iScreenMode
                .Items("Create Air BOL").Settings.Enabled = not_iScreenMode
            End With
            .Groups("Show if Entered in").Visible = Not ScreenMode
        End With

        lblStatus.Visible = ScreenMode
        splDetails.Visible = ScreenMode

        grdPOTSHIP1_X.Visible = Not ScreenMode
        splDetails.Visible = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            Select Case EntryMode
                Case "V"
                    Set_Read_Only(splVessel.Panel1, True)
                    txtNOTES.ReadOnly = True

                Case Else
                    Set_Read_Only(splVessel.Panel1, False)
                    txtNOTES.ReadOnly = False
            End Select
        Else
            Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)

        For Each tablename As String In New String() {"POTSHIP1_X", "POTSHIP2_X", "POTSHIP3_X",
                                        "POTSHIP1", "POTSHIP2", "POTSHIP3",
                                        "ICTPINV1", "POTASNM2_X", "POTASNM3_X",
                                        "POTSHIP1_IMP", "POTSHIP2_IMP", "POTSHIP3_IMP",
                                        aswPOTSHIP1, aswPOTSHIP2, aswPOTSHIP3}
            dst.Tables(tablename).Rows.Clear()
        Next

        Refresh_Documents()
        rowPOTSHIP1 = Nothing
        EnforceConstraints(True)

        PO_SHIPMENT_NO = String.Empty
        CONTAINER_CTL_NO = String.Empty
        VEND_CODE = String.Empty
        processingImportedData = False
        toggleAllowEdit(False)

        Absx1.txtFor("VEND_CODE").Text = VEND_CODE_DEFAULT
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        VEND_CODE = HFs("VEND_CODE")

        Fill_Records("POTSHIP1", PO_SHIPMENT_NO)
        Fill_Records("POTSHIP2", {PO_SHIPMENT_NO, CONTAINER_CTL_NO})

        'ASCMAIN1.sql = "SELECT POTSHIP3.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.VESSEL_NAME, ICTPINV1.PINV_NO, ICTPINV1.WHSE_CODE, ICTPINV1.VEND_CODE
        '                        FROM POTSHIP3, ICTPINV1
        '                        WHERE (POTSHIP3.PO_SHIPMENT_NO, POTSHIP3.CONTAINER_CTL_NO) IN (SELECT PO_SHIPMENT_NO, CONTAINER_CTL_NO FROM POTSHIP2 WHERE PO_SHIPMENT_NO = :PARM1 AND CONTAINER_CTL_NO = :PARM2)
        '                        AND POTSHIP3.INV_NUM = ICTPINV1.INV_NUM (+)
        '                        AND POTSHIP3.PACK_LIST_NO = ICTPINV1.PACK_LIST_NO (+)
        '                        AND POTSHIP3.CONTAINER_CTL_NO = ICTPINV1.CONTAINER_CTL_NO (+)"

        ASCMAIN1.sql = $"SELECT POTSHIP3.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.VESSEL_NAME, ICTPINV1.PINV_NO, ICTPINV1.WHSE_CODE, ICTPINV1.VEND_CODE, POTORDR1.WHSE_CODE WHSE_CODE_PO, ICTPINV1.PINV_STATUS
                                FROM POTSHIP3, ICTPINV1, POTORDR1
                                WHERE POTSHIP3.PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'
                                AND POTSHIP3.CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'
                                AND POTSHIP3.INV_NUM = ICTPINV1.INV_NUM (+)
                                AND POTSHIP3.PACK_LIST_NO = ICTPINV1.PACK_LIST_NO (+)
                                AND POTSHIP3.CONTAINER_CTL_NO = ICTPINV1.CONTAINER_CTL_NO (+)
                                AND ICTPINV1.PO_ORDER_NO = POTORDR1.PO_ORDER_NO (+)"
        Fill_Records("POTSHIP3", String.Empty, True, ASCMAIN1.sql)

        Dim rowPOTSHIP2 As DataRow = dst.Tables("POTSHIP2").Rows.Find({PO_SHIPMENT_NO, CONTAINER_CTL_NO})

        VEND_CODE = rowPOTSHIP1.Item("VEND_CODE") & String.Empty
        WHSE_CODE = rowPOTSHIP2.Item("WHSE_CODE") & String.Empty

        Fill_Records("ICTPINV1", New String() {WHSE_CODE})

        ASCMAIN1.sql = $"SELECT * FROM ASTAUDT1 WHERE TABLE_NAME = 'POTSHIP1' AND KEY_VALUE = '{PO_SHIPMENT_NO}' AND COLUMN_NAME <> 'LAST_DATE'"
        Fill_Records("ASTAUDT1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdASTAUDT1, "INIT_DATE,COLUMN_NAME")

        Sort_grdColumns(grdPOTSHIP3, "PINV_NO")
        Sort_grdColumns(grdICTPINV1, "PINV_NO")

        Select Case rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") & String.Empty
            Case "O"
                lblStatus.Text = "Open"
            Case "P"
                lblStatus.Text = "Transmitted"
            Case "C"
                lblStatus.Text = "Completed"
            Case Else
                lblStatus.Text = "Unknown Status (" & rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") & String.Empty & ")"
        End Select

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        If processingImportedData Then
            Try
                BeginTrans()

                dst.Tables("POTSHIP1_IMP").Rows.Clear()
                dst.Tables("POTSHIP2_IMP").Rows.Clear()
                dst.Tables("POTSHIP3_IMP").Rows.Clear()

                For Each drPOTSHIP1_X As DataRow In dst.Tables("POTSHIP1_X").Select("")
                    dst.Tables("POTSHIP1_IMP").ImportRow(drPOTSHIP1_X)
                Next

                For Each drPOTSHIP2_X As DataRow In dst.Tables("POTSHIP2_X").Select("")
                    dst.Tables("POTSHIP2_IMP").ImportRow(drPOTSHIP2_X)
                Next

                For Each drPOTSHIP3_X As DataRow In dst.Tables("POTSHIP3_X").Select("")
                    dst.Tables("POTSHIP3_IMP").ImportRow(drPOTSHIP3_X)
                Next

                For Each tableName As String In {"POTSHIP1_IMP", "POTSHIP2_IMP", "POTSHIP3_IMP"}
                    dst.Tables(tableName).AcceptChanges()
                    For Each dr As DataRow In dst.Tables(tableName).Select
                        dr.SetAdded()
                    Next
                Next

                UpdateLOAImportData()

                CommitTrans("LOA successfully imported.")
            Catch ex As Exception
                Rollback(ex.Message)
            End Try

            processingImportedData = False
            Exit Sub
        End If

        Try
            BeginTrans()

            INIT_LAST("POTSHIP1")

            Dim PO_SHIPMENT_DNO As Int16 = 1
            For Each drPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("", "PO_SHIPMENT_DNO")
                drPOTSHIP3.Item("PO_SHIPMENT_DNO") = PO_SHIPMENT_DNO
                PO_SHIPMENT_DNO += 1
            Next

            Update_Record_TDA("POTSHIP1")
            Update_Record_TDA("POTSHIP2", $"CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'")
            Update_Record_TDA("POTSHIP3", $"CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'")

            ASCMAIN1.sql = "UPDATE ICTPINV1 SET CONTAINER_CTL_NO = NULL, ASN_NO = NULL WHERE CONTAINER_CTL_NO = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {CONTAINER_CTL_NO})

            Dim WHSE_CODE As String = dst.Tables("POTSHIP2").Rows(0).Item("WHSE_CODE")
            Dim ASN_NO As String = dst.Tables("POTSHIP2").Rows(0).Item("ASN_NO")
            For Each drPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select("")
                Dim PINV_NO As String = drPOTSHIP3.Item("PINV_NO")
                Dim PO_ORDER_NO As String = drPOTSHIP3.Item("PO_ORDER_NO")

                ASCMAIN1.sql = "UPDATE ICTPINV1 SET CONTAINER_CTL_NO = :PARM1, ASN_NO = :PARM2 WHERE PINV_NO = :PARM3"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", {CONTAINER_CTL_NO, ASN_NO, PINV_NO})

                ' See if we need to change the Warehouse Code on ICTPINV1 and/or POTORDR1
                Dim drICTPINV1 As DataRow = LookUp("ICTPINV1", PINV_NO)

                If drICTPINV1.Item("WHSE_CODE") & String.Empty <> WHSE_CODE Then
                    ' Add Walter's Code from ICFPINV1
                    Write_Event_Log("ICTPINV1", PINV_NO, $"Change Warehouse from {drICTPINV1.Item("WHSE_CODE") } to {WHSE_CODE}")
                    ASCMAIN1.sql = "UPDATE ICTPINV1 SET WHSE_CODE = :PARM1 WHERE PINV_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", {WHSE_CODE, PINV_NO})
                End If
            Next

            MyBase.CommitTrans("Update successful")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try
    End Sub

    Sub Transmit_ASN()
        ' write a record to the queue that will push ASNs out to HG
    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As Control, COLUMN_NAME As String, ByRef Optional sql_where As String = "", ByRef Optional Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                sql_where = "LP_CODE = 'ADS' AND WHSE_STATUS = 'A'"

            Case "VEND_CODE"
                sql_where = "VEND_CODE = 'IPSA'"
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                Fill_Records("ICTPINV1", New String() {txtWHSE_CODE.Text})
        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        InquiryMode = MENU_ITEM_OBJECT = "POTSHIPI"

        If Not InquiryMode Then
            Load_Popup_Menu(grdICTPINV1, "SSPBB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Add Selected to Shipment")
        Else
            Load_Popup_Menu(grdICTPINV1, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
        End If

        If Not InquiryMode Then
            Load_Popup_Menu(grdPOTSHIP3, "SSPBB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Remove from Shipment")
        Else
            Load_Popup_Menu(grdPOTSHIP3, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
        End If

        If Not InquiryMode Then
            Load_Popup_Menu(grdPOTSHIP1_X, "SSPBBBB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Create ASN", "Delete BOL", "Reset ASN")
        Else
            Load_Popup_Menu(grdPOTSHIP1_X, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
        End If

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        MyBase.tlb_BeforeToolDropdown(sender, e)

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If TypeOf e.SourceControl IsNot UltraWinGrid.UltraGrid Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name
            Case grdPOTSHIP3.Name
                If tlb_pop.Tools.Exists("Remove from Shipment") Then
                    rowPOTSHIP1 = dst.Tables("POTSHIP1").Rows.Find(PO_SHIPMENT_NO)
                    tlb_btn = DirectCast(tlb_pop.Tools("Remove from Shipment"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                            AndAlso rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") & String.Empty = "O" _
                            AndAlso "EN".Contains(EntryMode) _
                            AndAlso grd.Selected.Rows.Count > 0
                End If

            Case grdICTPINV1.Name
                If tlb_pop.Tools.Exists("Add Selected to Shipment") Then
                    rowPOTSHIP1 = dst.Tables("POTSHIP1").Rows.Find(PO_SHIPMENT_NO)
                    tlb_btn = DirectCast(tlb_pop.Tools("Add Selected to Shipment"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                        AndAlso rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") & String.Empty = "O" _
                        AndAlso "EN".Contains(EntryMode) _
                        AndAlso grd.Selected.Rows.Count > 0
                End If

            Case grdPOTSHIP1_X.Name
                If tlb_pop.Tools.Exists("Create ASN") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Create ASN"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                        AndAlso grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key AndAlso Not processingImportedData
                End If

                If tlb_pop.Tools.Exists("Delete BOL") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Delete BOL"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                        AndAlso grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(0).Key AndAlso Not processingImportedData
                End If

                If tlb_pop.Tools.Exists("Reset ASN") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Reset ASN"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                        AndAlso grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key AndAlso Not processingImportedData
                End If

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

            Case "Reset ASN"
                ' ISSUE-7349 from Marissa:
                Dim CONTAINER_CTL_NO As String = grd.ActiveRow.Cells("CONTAINER_CTL_NO").Value & String.Empty
                Dim CONTAINER_NO As String = grd.ActiveRow.Cells("CONTAINER_NO").Value & String.Empty
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value & String.Empty
                Dim ASN_NO As String = grd.ActiveRow.Cells("ASN_NO").Value & String.Empty

                If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then
                    Exit Sub
                End If

                Dim drPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                If drPOTSHIP1 Is Nothing Then
                    MessageBox.Show($"Cannot locate PO Shipment No {PO_SHIPMENT_NO}.", "Reset ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                Dim drPOTSHIP2 As DataRow = LookUp("POTSHIP2", {PO_SHIPMENT_NO, CONTAINER_CTL_NO})
                If drPOTSHIP2 Is Nothing Then
                    MessageBox.Show($"Cannot locate PO Shipment No {PO_SHIPMENT_NO}, Container Control No {CONTAINER_CTL_NO}", "Reset ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                If drPOTSHIP2.Item("ASN_REQUIRES_XMIT") & String.Empty = "1" Then
                    MessageBox.Show($"The selected ASN has not been transferred.", "Reset ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                If MessageBox.Show($"ASN was already transmitted. You must notify your 3PL to cancel the previously sent ASN. Do you want to Reset ASN {ASN_NO}, Container {CONTAINER_NO}? This process will assign a new ASN No.", "Reset ASN", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                Try
                    BeginTrans()

                    Dim ASN_NO_NEW As String = ASCMAIN1.Next_Control_No("POTASNM1.ASN_NO")

                    ASCMAIN1.sql = "UPDATE POTSHIP1 SET PO_SHIPMENT_STATUS = 'O' WHERE PO_SHIPMENT_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    ASCMAIN1.sql = "UPDATE POTSHIP2 SET ASN_NO = :PARM1, ASN_REQUIRES_XMIT = '1', ASN_DATE = TRUNC(SYSDATE) WHERE ASN_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {ASN_NO_NEW, ASN_NO})

                    ASCMAIN1.sql = "DELETE FROM POTASNM1 WHERE ASN_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {ASN_NO})

                    ASCMAIN1.sql = "DELETE FROM POTASNM2 WHERE ASN_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {ASN_NO})

                    ASCMAIN1.sql = "DELETE FROM POTASNM3 WHERE ASN_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {ASN_NO})

                    ASCMAIN1.sql = "UPDATE ICTPINV1 SET SHIP_DATE = NULL, ETA_DATE = NULL, ASN_NO = :PARM1 WHERE ASN_NO = :PARM2"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {ASN_NO_NEW, ASN_NO})

                    CommitTrans($"ASN No {ASN_NO} reset to New ASN No {ASN_NO_NEW}.")
                Catch ex As Exception
                    Rollback(ex.Message)
                Finally
                    ASCMAIN1.MultiTask_Release()
                    Click_Command("Refresh")
                End Try

            Case "Delete BOL"
                ' ISSUE-7349 from Marissa:
                If InquiryMode Then
                    MessageBox.Show("Inquiry Mode does Not permit Deleting a BOL", "Delete BOL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                If InputBox("Enter Password", "Delete BOL") <> "Oliver0412" Then
                    MessageBox.Show("Invalid password.", "IDelete BOL", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Exit Sub
                End If

                Dim BOL_NO As String = grd.ActiveRow.Cells("BOL_NO").Value & String.Empty
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value & String.Empty

                If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then
                    Exit Sub
                End If

                Dim drPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                If drPOTSHIP1 Is Nothing Then
                    MessageBox.Show($"Cannot locate PO Shipment No {PO_SHIPMENT_NO}.", "Delete BOL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If

                If MessageBox.Show($"Do you want to Delete BOL {BOL_NO}? It is your responsibilty to notify the 3PL to delete any ASNs they refceived for this BOL.", "Delete BOL", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                Try
                    BeginTrans()

                    ASCMAIN1.sql = "DELETE FROM POTASNM1 WHERE ASN_NO IN
                                        (
                                        SELECT ASN_NO FROM POTSHIP2 WHERE PO_SHIPMENT_NO = :PARM1
                                        )"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    ASCMAIN1.sql = "DELETE FROM POTASNM2 WHERE ASN_NO IN
                                        (
                                        SELECT ASN_NO FROM POTSHIP2 WHERE PO_SHIPMENT_NO = :PARM1
                                        )"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    ASCMAIN1.sql = "DELETE FROM POTASNM3 WHERE ASN_NO IN
                                        (
                                        SELECT ASN_NO FROM POTSHIP2 WHERE PO_SHIPMENT_NO = :PARM1
                                        )"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    ASCMAIN1.sql = "UPDATE ICTPINV1 SET CONTAINER_CTL_NO = NULL, ASN_NO = NULL 
                                        WHERE ASN_NO IN 
                                        (
                                        SELECT ASN_NO FROM POTSHIP2 WHERE PO_SHIPMENT_NO = :PARM1
                                        )"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    ASCMAIN1.sql = "DELETE FROM POTSHIP1 WHERE PO_SHIPMENT_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    ASCMAIN1.sql = "DELETE FROM POTSHIP2 WHERE PO_SHIPMENT_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    ASCMAIN1.sql = "DELETE FROM POTSHIP3 WHERE PO_SHIPMENT_NO = :PARM1"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                    CommitTrans($"BOL NO {BOL_NO} has been deleted.")

                Catch ex As Exception
                    Rollback(ex.Message)
                Finally
                    ASCMAIN1.MultiTask_Release()
                    Click_Command("Refresh")
                End Try

            Case "Add Selected to Shipment"
                Try
                    If InquiryMode Then
                        MessageBox.Show("Inquiry Mode does Not permit 'Adding To an PO Shipment'", "Add Selected To Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") <> "O" Then
                        MessageBox.Show("This PO Shipment does Not have a status of Open", "Add Selected To Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If Not "E".Contains(EntryMode) Then
                        MessageBox.Show("Form State does Not permit 'Adding To an PO Shipment'", "Add Selected To Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim WHSE_CODE As String = String.Empty
                    Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("POTSHIP3").Select("ISNULL(WHSE_CODE, '') <> ''"), {"WHSE_CODE"})
                    Select Case tbl.Rows.Count
                        Case 0

                        Case 1
                            WHSE_CODE = tbl.Rows(0).Item("WHSE_CODE") & String.Empty
                        Case Else
                            MessageBox.Show("Selected Invoices contains more than 1 warehouse", "Add Selected to Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                    End Select

                    Me.Cursor = Cursors.WaitCursor
                    Dim lstPINV_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                        If grd.Selected.Rows.Contains(grow) Then
                            lstPINV_NOs.Add(grow.Cells("PINV_NO").Value & String.Empty)
                        End If
                    Next

                    For Each PINV_NO As String In lstPINV_NOs
                        ASCMAIN1.Progress("Processing Invoice(s)", PINV_NO)
                        For Each rowICTPINV1 As DataRow In dst.Tables("ICTPINV1").Select($"PINV_NO = '{PINV_NO}'")
                            If WHSE_CODE.Length = 0 Then
                                WHSE_CODE = rowICTPINV1.Item("WHSE_CODE") & String.Empty
                            End If

                            If dst.Tables("POTSHIP3").Select($"PINV_NO = '{PINV_NO}'").Length > 0 Then
                                Continue For
                            End If

                            If WHSE_CODE = rowICTPINV1.Item("WHSE_CODE") & String.Empty Then
                                Dim drPOTSHIP3 As DataRow = dst.Tables("POTSHIP3").NewRow
                                drPOTSHIP3.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                                drPOTSHIP3.Item("CONTAINER_CTL_NO") = CONTAINER_CTL_NO
                                drPOTSHIP3.Item("PO_SHIPMENT_DNO") = Val(dst.Tables("POTSHIP3").Compute("max(PO_SHIPMENT_DNO)", $"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}' and CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'") & String.Empty) + 1
                                drPOTSHIP3.Item("INV_NUM") = rowICTPINV1.Item("INV_NUM")
                                drPOTSHIP3.Item("PACK_LIST_NO") = rowICTPINV1.Item("PACK_LIST_NO")
                                drPOTSHIP3.Item("PO_ORDER_NO") = rowICTPINV1.Item("PO_ORDER_NO")
                                drPOTSHIP3.Item("VESSEL_NAME") = rowICTPINV1.Item("VESSEL_NAME")
                                drPOTSHIP3.Item("PINV_NO") = rowICTPINV1.Item("PINV_NO")
                                drPOTSHIP3.Item("WHSE_CODE") = rowICTPINV1.Item("WHSE_CODE")
                                drPOTSHIP3.Item("VEND_CODE") = rowICTPINV1.Item("VEND_CODE")
                                drPOTSHIP3.Item("WHSE_CODE_PO") = rowICTPINV1.Item("WHSE_CODE_PO")
                                dst.Tables("POTSHIP3").Rows.Add(drPOTSHIP3)
                            End If
                        Next
                    Next

                    dst.Tables("POTSHIP3").AcceptChanges()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("", "")
                End Try

            Case "Remove from Shipment"

                Try
                    If InquiryMode Then
                        MessageBox.Show("Inquiry Mode does Not permit Removing from a PO Shipment", "Remove from Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") <> "O" Then
                        MessageBox.Show("This PO Shipment does Not have a status Of Open", "Remove from Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If Not "E".Contains(EntryMode) Then
                        MessageBox.Show("Form State does Not permit Removing from a PO Shipment", "Remove from Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Me.Cursor = Cursors.WaitCursor
                    Dim lstPINV_NOs As New List(Of String)
                    Dim lstINV_NUMs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                        If grd.Selected.Rows.Contains(grow) Then
                            If grow.Cells("PINV_NO").Value & String.Empty <> String.Empty Then
                                lstPINV_NOs.Add(grow.Cells("PINV_NO").Value & String.Empty)
                            Else
                                lstINV_NUMs.Add(grow.Cells("INV_NUM").Value & String.Empty)
                            End If
                        End If
                    Next

                    For Each PINV_NO As String In lstPINV_NOs
                        ASCMAIN1.Progress("Processing Invoice(s)", PINV_NO)
                        For Each drPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select($"PINV_NO = '{PINV_NO}'")
                            drPOTSHIP3.Delete()
                        Next
                    Next

                    For Each INV_NUM As String In lstINV_NUMs
                        ASCMAIN1.Progress("Processing Invoice(s)", INV_NUM)
                        For Each drPOTSHIP3 As DataRow In dst.Tables("POTSHIP3").Select($"INV_NUM = '{INV_NUM}'")
                            drPOTSHIP3.Delete()
                        Next
                    Next

                    dst.Tables("POTSHIP3").AcceptChanges()

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("", "")
                End Try

            Case "Create ASN"
                Try
                    Dim CONTAINER_CTL_NO As String = grd.ActiveRow.Cells("CONTAINER_CTL_NO").Value & String.Empty
                    Dim CONTAINER_NO As String = grd.ActiveRow.Cells("CONTAINER_NO").Value & String.Empty
                    Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value & String.Empty
                    Dim ASN_NO As String = grd.ActiveRow.Cells("ASN_NO").Value & String.Empty
                    Dim ASN_REQUIRES_XMIT As String = grd.ActiveRow.Cells("ASN_REQUIRES_XMIT").Value & String.Empty

                    If ASN_REQUIRES_XMIT <> "1" Then
                        MessageBox.Show($"ASN was already transmitted.", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If Not ASCMAIN1.Logical_Lock("POTSHIP1", PO_SHIPMENT_NO) Then
                        Exit Sub
                    End If

                    Dim drPOTSHIP1 As DataRow = LookUp("POTSHIP1", PO_SHIPMENT_NO)
                    If drPOTSHIP1 Is Nothing Then
                        MessageBox.Show($"Cannot locate PO Shipment No {PO_SHIPMENT_NO}.", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    Dim drPOTSHIP2 As DataRow = LookUp("POTSHIP2", {PO_SHIPMENT_NO, CONTAINER_CTL_NO})
                    If drPOTSHIP2 Is Nothing Then
                        MessageBox.Show($"Cannot locate PO Shipment No {PO_SHIPMENT_NO}, Container Control No {CONTAINER_CTL_NO}", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    ' ISSUE-7327 ASN Invoice number issue. Verify the Invoice numbers on POTSHIP3 exist in ICTPINV1
                    Dim lstMissingPOs As New List(Of String)
                    Dim lstMissingInvoiceNos As New List(Of String)

                    ASCMAIN1.sql = "SELECT POTSHIP3.INV_NUM INV_NUM_3, ICTPINV1.*
                                        FROM POTSHIP3, ICTPINV1
                                        WHERE POTSHIP3.CONTAINER_CTL_NO = ICTPINV1.CONTAINER_CTL_NO (+)
                                        AND POTSHIP3.INV_NUM = ICTPINV1.INV_NUM (+)
                                        AND POTSHIP3.CONTAINER_CTL_NO = :PARM1"

                    Dim tblICTPINV1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {CONTAINER_CTL_NO})
                    If tblICTPINV1.Rows.Count = 0 Then
                        MessageBox.Show($"Cannot locate Invoices connected to Container {CONTAINER_NO}.", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    For Each drICTPINV1 As DataRow In tblICTPINV1.Select("")
                        If drICTPINV1.Item("INV_NUM") & String.Empty = String.Empty Then
                            lstMissingInvoiceNos.Add(drICTPINV1.Item("INV_NUM_3") & String.Empty)
                        ElseIf drICTPINV1.Item("PO_ORDER_NO") & String.Empty = String.Empty Then
                            lstMissingPOs.Add(drICTPINV1.Item("INV_NUM_3") & String.Empty)
                        End If
                    Next

                    Dim errMsg As String = String.Empty

                    If lstMissingInvoiceNos.Count > 0 Then
                        errMsg &= $"Cannot locate the following Invoices (ICTPINV1) connected to Container {CONTAINER_NO}: {String.Join(", ", lstMissingInvoiceNos.ToArray)}." & Environment.NewLine
                    End If

                    If lstMissingPOs.Count > 0 Then
                        errMsg &= $"The following Invoices (ICTPINV1) connected to Container {CONTAINER_NO} are not assigned a PO: {String.Join(", ", lstMissingPOs.ToArray)}."
                    End If

                    If errMsg.Length > 0 Then
                        MessageBox.Show(errMsg, "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    For Each drICTPINV1 As DataRow In tblICTPINV1.Select("")
                        Dim PO_ORDER_NO As String = drICTPINV1.Item("PO_ORDER_NO") & String.Empty
                        Dim INV_NUM As String = drICTPINV1.Item("INV_NUM") & String.Empty
                        Dim PINV_NO As String = drICTPINV1.Item("PINV_NO") & String.Empty

                        If PO_ORDER_NO.Length = 0 Then
                            MessageBox.Show($"Invoice {INV_NUM} is not assigned a Purchase Order No.", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            ASCMAIN1.MultiTask_Release()
                            Exit Sub
                        End If

                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                            Exit Sub
                        End If

                        If Not ASCMAIN1.Logical_Lock("ICTPINV1", PINV_NO) Then
                            Exit Sub
                        End If
                    Next

                    ASCMAIN1.sql = "SELECT * FROM ICTPINV1 WHERE CONTAINER_CTL_NO = :PARM1"
                    tblICTPINV1 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTPINV1", "V", {CONTAINER_CTL_NO})
                    Dim lstNotOpen As New List(Of String)
                    For Each drICTPINV1 As DataRow In tblICTPINV1.Select("PINV_STATUS <> 'O'")
                        lstNotOpen.Add(drICTPINV1.Item("INV_NUM") & String.Empty)
                    Next

                    If lstNotOpen.Count > 0 Then
                        MessageBox.Show($"The following Invoices are no Open: {String.Join(", '", lstNotOpen.ToArray)}.", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    ASCMAIN1.sql = "SELECT ICTPINV2.ITEM_CODE, ICTPINV1.WHSE_CODE, SUM(NVL(ICTPINV2.PINV_QTY, 0)) PINV_QTY, MAX(ICTPINV2.PINV_COST) PINV_COST
                                        FROM ICTPINV1, ICTPINV2
                                        WHERE ICTPINV1.PINV_NO = ICTPINV2.PINV_NO
                                        AND ICTPINV1.CONTAINER_CTL_NO = :PARM1
                                        AND ICTPINV1.RECEIPT_NO IS NULL
                                        GROUP BY ICTPINV2.ITEM_CODE, ICTPINV1.WHSE_CODE"
                    Dim tblICTPINV2 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {CONTAINER_CTL_NO})
                    If tblICTPINV2.Rows.Count = 0 Then
                        MessageBox.Show($"There are no open Invoice quantities to transfer.", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    If ASCDATA1.SelectDistinct(tblICTPINV2, {"WHSE_CODE"}).Rows.Count > 1 Then
                        MessageBox.Show($"There Invoices assigned to this contaier are for different warehouses.", "Create ASN", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    If MessageBox.Show($"Do you want to create an ASN for all open quantites for Continer {CONTAINER_NO}?", "Create ASN", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If

                    Dim ASN_XMIT_NO As String = String.Empty
                    Fill_Records("POTASNM1", "", True, $"SELECT * FROM POTASNM1 WHERE CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'")

                    Dim ASNexists As Boolean = dst.Tables("POTASNM1").Rows.Count = 1
                    ASN_XMIT_NO = ASCMAIN1.Next_Control_No("POTASNM2.ASN_XMIT_NO")

                    Try
                        BeginTrans()

                        If Not ASNexists Then
                            Dim drPOTASNM1 As DataRow = dst.Tables("POTASNM1").NewRow
                            drPOTASNM1.Item("ASN_NO") = ASN_NO
                            drPOTASNM1.Item("CONTAINER_CTL_NO") = CONTAINER_CTL_NO
                            dst.Tables("POTASNM1").Rows.Add(drPOTASNM1)
                        End If

                        Dim drPOTASNM2 As DataRow = dst.Tables("POTASNM2").NewRow
                        drPOTASNM2.Item("ASN_NO") = ASN_NO
                        drPOTASNM2.Item("ASN_XMIT_NO") = ASN_XMIT_NO
                        drPOTASNM2.Item("INIT_DATE") = DateTime.Now
                        drPOTASNM2.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        'drPOTASNM2.Item("TRANSMIT_DATE") = DateTime.Now
                        'drPOTASNM2.Item("TRANSMIT_OPER") = ASCMAIN1.USER_ID
                        dst.Tables("POTASNM2").Rows.Add(drPOTASNM2)

                        Dim ASN_LNO As Int32 = 1
                        For Each drICTPINV2 As DataRow In tblICTPINV2.Select("", "ITEM_CODE")
                            Dim drPOTASNM3 As DataRow = dst.Tables("POTASNM3").NewRow
                            drPOTASNM3.Item("ASN_NO") = ASN_NO
                            drPOTASNM3.Item("ASN_XMIT_NO") = ASN_XMIT_NO
                            drPOTASNM3.Item("ASN_LNO") = ASN_LNO
                            drPOTASNM3.Item("ITEM_CODE") = drICTPINV2.Item("ITEM_CODE")
                            drPOTASNM3.Item("ASN_QTY") = drICTPINV2.Item("PINV_QTY")
                            drPOTASNM3.Item("ASN_COST") = drICTPINV2.Item("PINV_COST")
                            dst.Tables("POTASNM3").Rows.Add(drPOTASNM3)
                            ASN_LNO += 1
                        Next

                        Update_Record_TDA("POTASNM1")
                        Update_Record_TDA("POTASNM2")
                        Update_Record_TDA("POTASNM3")

                        CommitTrans("The ASN has been created and will now be sent to the 3PL")

                    Catch ex As Exception
                        Rollback(ex.Message)
                        ASN_NO = String.Empty
                        Exit Sub
                    End Try

                    Dim WHSE_CODE As String = tblICTPINV2.Rows(0).Item("WHSE_CODE")
                    Dim drICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

                    Dim LP_CODE As String = drICTWHSE1.Item("LP_CODE")

                    Dim XMIT_NO As String = String.Empty

                    ' Update all changed items
                    Try
                        XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "All", LP_CODE)
                    Catch ex As Exception

                    End Try

                    XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "ASN", ASN_XMIT_NO, LP_CODE)
                    If XMIT_NO.Length > 0 Then

                        Dim dr As DataRow = ASCDATA1.GetDataRow("SELECT * FROM POTASNM2 WHERE ASN_XMIT_NO = :PARM1", "V", {ASN_XMIT_NO})

                        If dr IsNot Nothing Then
                            If dr.Item("TRANSMIT_OPER") & String.Empty <> String.Empty Then
                                ASCMAIN1.sql = "UPDATE POTASNM2 SET TRANSMIT_DATE = SYSDATE, TRANSMIT_OPER = :PARM1 WHERE ASN_NO = :PARM2 AND ASN_XMIT_NO = :PARM3"
                                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", {ASCMAIN1.USER_ID, ASN_NO, ASN_XMIT_NO})
                            End If
                        End If

                        ASCMAIN1.sql = "UPDATE POTSHIP2 SET ASN_REQUIRES_XMIT = '0' WHERE CONTAINER_CTL_NO = :PARM1"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {CONTAINER_CTL_NO})

                        If dst.Tables("POTSHIP2_X").Select($"CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'").Length > 0 Then
                            drPOTSHIP2 = dst.Tables("POTSHIP2_X").Select($"CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'")(0)
                            If drPOTSHIP2.Item("ASN_NO") & String.Empty = String.Empty Then
                                drPOTSHIP2.Item("ASN_NO") = ASN_NO
                                drPOTSHIP2.Item("ASN_DATE") = DateTime.Now
                            End If
                            drPOTSHIP2.Item("ASN_REQUIRES_XMIT") = "0"
                        End If

                        Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT * FROM POTSHIP2 WHERE ASN_REQUIRES_XMIT = '1' AND PO_SHIPMENT_NO = :PARM1", "", "V", {PO_SHIPMENT_NO})
                        If tbl.Rows.Count = 0 Then
                            ASCMAIN1.sql = "UPDATE POTSHIP1 SET PO_SHIPMENT_STATUS = 'P' where PO_SHIPMENT_NO = :PARM1"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})
                        End If

                        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS
                                            SELECT POTSHIP1.SHIP_DATE, POTSHIP1.ETA_DATE, POTSHIP1.VESSEL_NAME, POTSHIP1.BOL_NO, POTSHIP2.CONTAINER_NO, POTSHIP3.INV_NUM, POTSHIP3.PACK_LIST_NO
                                            FROM POTSHIP1, POTSHIP2, POTSHIP3
                                            WHERE POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO
                                            AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO
                                            AND POTSHIP2.CONTAINER_CTL_NO = POTSHIP3.CONTAINER_CTL_NO
                                            AND POTSHIP1.PO_SHIPMENT_NO = :PARM1
                                            AND POTSHIP2.CONTAINER_CTL_NO = :PARM2;
                                            BEGIN FOR R1 IN C1 LOOP
                                             UPDATE ICTPINV1 SET SHIP_DATE = R1.SHIP_DATE, ETA_DATE = R1.ETA_DATE, VESSEL_NAME = R1.VESSEL_NAME,
                                                BOL_NO = R1.BOL_NO, CONTAINER_NO = R1.CONTAINER_NO 
                                                WHERE INV_NUM = R1.INV_NUM AND PACK_LIST_NO = R1.PACK_LIST_NO;
                                            END LOOP; END; END;"
                        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {PO_SHIPMENT_NO, CONTAINER_CTL_NO})
                    End If

                    ASCMAIN1.sql = $"SELECT POTASNM2.*, POTASNM1.CONTAINER_CTL_NO
                            FROM POTASNM1, POTASNM2
                            WHERE POTASNM1.ASN_NO = POTASNM2.ASN_NO
                            AND POTASNM2.ASN_XMIT_NO = '{ASN_XMIT_NO}'"

                    Fill_Records("POTASNM2_X", String.Empty, False, ASCMAIN1.sql)

                    ASCMAIN1.sql = $"SELECT POTASNM3.*, ICTITEM1.ITEM_DESC
                            FROM POTASNM3, ICTITEM1
                            WHERE ASN_XMIT_NO = '{ASN_XMIT_NO}'
                            AND POTASNM3.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
                    Fill_Records("POTASNM3_X", String.Empty, False, ASCMAIN1.sql)

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    ASCMAIN1.MultiTask_Release()
                End Try

        End Select
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub ImportLOA()
        Try
            processingImportedData = False

            Dim importFileName As String = String.Empty
            Dim zMsg As String = String.Empty

            Dim openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel file To Import"
            openFileDialog1.Filter = "Excel files|*.xls;*.xlsx"
            openFileDialog1.RestoreDirectory = True
            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                importFileName = openFileDialog1.FileName.Trim
                If importFileName.Length = 0 Then
                    zMsg = "No file selected."
                ElseIf Not (importFileName.ToUpper.EndsWith(".XLS") Or importFileName.ToUpper.EndsWith(".XLSX")) Then
                    zMsg = "Non Excel file selected."
                End If
            Else
                zMsg = "No file selected."
            End If

            If zMsg.Length > 0 Then
                MessageBox.Show("Import LOA", zMsg)
                Exit Sub
            End If

            Dim filename As String = (System.IO.Path.GetFileName(importFileName)).ToUpper
            If filename.Length > dst.Tables("POTSHIP1").Columns("FILENAME").MaxLength Then
                filename = filename.Substring(0, dst.Tables("WHTMANI1").Columns("FILENAME").MaxLength).Trim
            End If

            ASCMAIN1.sql = "SELECT * FROM POTSHIP1 WHERE FILENAME = :PARM1"
            Dim tblWHTMANI1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "POTSHIP1", "V", New Object() {filename})
            If tblWHTMANI1 IsNot Nothing AndAlso tblWHTMANI1.Rows.Count > 0 Then
                Dim importedOn As String = String.Empty
                For Each row As DataRow In tblWHTMANI1.Select("", "INIT_DATE")
                    importedOn &= Environment.NewLine & row.Item("INIT_OPER") & " - " & row.Item("INIT_DATE")
                Next

                If MessageBox.Show($"File {filename} was imported {tblWHTMANI1.Rows.Count} time(s).{importedOn}{Environment.NewLine}{Environment.NewLine}Do you want to continue?", "Import Excel File", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If
            Else
                If MessageBox.Show($"Do you want to Import File {filename}?", "Import Excel File", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If
            End If

            Const CONT_LNO As Int16 = 0
            Const CONTAINER As Int16 = 1
            Const CONT_TYPE As Int16 = 2
            Const NAVA As Int16 = 3
            Const SEAL As Int16 = 4
            Const PACK_LIST As Int16 = 5
            Const INVOICE_NO As Int16 = 6
            Const PALLETS As Int16 = 7
            Const CARTONS As Int16 = 8
            Const WEIGHT As Int16 = 9
            Const CUBE As Int16 = 10

            Const VESSEL As Int16 = 2
            Const VOYAGE As Int16 = 2
            Const SCAC As Int16 = 2
            Const COMPANY As Int16 = 2
            Const BOL As Int16 = 4
            Const ETD As Int16 = 10
            Const ETA As Int16 = 10
            Const WHSE As Int16 = 10

            For Each tablename As String In New String() {"POTSHIP1_IMP", "POTSHIP2_IMP", "POTSHIP3_IMP"}
                dst.Tables(tablename).Rows.Clear()
            Next

            For Each tablename As String In New String() {aswPOTSHIP1, aswPOTSHIP2, aswPOTSHIP3}
                ASCDATA1.ExecuteSQL($"DELETE FROM {tablename}")
                dst.Tables(tablename).Rows.Clear()
            Next

            Dim oWB As SpreadsheetGear.IWorkbook
            oWB = SpreadsheetGear.Factory.GetWorkbook(importFileName)
            Dim worksheetIndex As Short = 0
            Dim ws As SpreadsheetGear.IWorksheet = oWB.Worksheets(worksheetIndex)
            Dim startRow As Short = 0
            Dim usedRanged As Int32 = ws.UsedRange.RowCount

            Dim rowPOTSHIP1 As DataRow = Nothing
            Dim PO_SHIPMENT_NO As String = String.Empty
            Dim loaWhse As String = String.Empty

            For excelRow As Int16 = 0 To ws.UsedRange.RowCount + 20

                Select Case excelRow
                    Case 0
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Loading Excel Workbook")
                        rowPOTSHIP1 = dst.Tables(aswPOTSHIP1).NewRow
                        PO_SHIPMENT_NO = ASCMAIN1.Next_Control_No("POTSHIP1.PO_SHIPMENT_NO")
                        rowPOTSHIP1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO

                        Dim VESSEL_NAME As String = ws.Cells(excelRow, VESSEL).Value & String.Empty
                        If VESSEL_NAME.Contains("=") Then
                            VESSEL_NAME = VESSEL_NAME.Substring(InStr(VESSEL_NAME, "=")).Trim
                        End If
                        VESSEL_NAME = VESSEL_NAME.Substring(0, Math.Min(dst.Tables(aswPOTSHIP1).Columns("VESSEL_NAME").MaxLength, VESSEL_NAME.Length))
                        rowPOTSHIP1.Item("VESSEL_NAME") = VESSEL_NAME

                        Dim SHIP_DATE As String = ws.Cells(excelRow, ETD).Text & String.Empty
                        If IsDate(SHIP_DATE) Then
                            rowPOTSHIP1.Item("SHIP_DATE") = CDate(SHIP_DATE).ToShortDateString
                        End If

                        rowPOTSHIP1.Item("VEND_CODE") = txtVEND_CODE.Text
                        rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") = "O"
                        rowPOTSHIP1.Item("INIT_DATE") = DateTime.Now
                        rowPOTSHIP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowPOTSHIP1.Item("LAST_DATE") = DateTime.Now
                        rowPOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowPOTSHIP1.Item("FILENAME") = filename
                        rowPOTSHIP1.Item("OPS_YYYYPP") = ASCMAIN1.CYP

                    Case 1
                        Dim VOYAGE_NO As String = ws.Cells(excelRow, VOYAGE).Value & String.Empty
                        If VOYAGE_NO.Contains("=") Then
                            VOYAGE_NO = VOYAGE_NO.Substring(InStr(VOYAGE_NO, "=")).Trim
                        End If
                        VOYAGE_NO = VOYAGE_NO.Substring(0, Math.Min(dst.Tables(aswPOTSHIP1).Columns("VOYAGE_NO").MaxLength, VOYAGE_NO.Length))
                        rowPOTSHIP1.Item("VOYAGE_NO") = VOYAGE_NO

                        Dim ETA_DATE As String = ws.Cells(excelRow, ETA).Text & String.Empty
                        If IsDate(ETA_DATE) Then
                            rowPOTSHIP1.Item("ETA_DATE") = CDate(ETA_DATE).ToShortDateString
                        End If

                    Case 2
                        Dim SCAC_CODE As String = ws.Cells(excelRow, SCAC).Value & String.Empty
                        If SCAC_CODE.Contains("=") Then
                            SCAC_CODE = SCAC_CODE.Substring(InStr(SCAC_CODE, "=")).Trim
                        End If
                        SCAC_CODE = SCAC_CODE.Substring(0, Math.Min(dst.Tables(aswPOTSHIP1).Columns("SCAC_CODE").MaxLength, SCAC_CODE.Length))
                        rowPOTSHIP1.Item("SCAC_CODE") = SCAC_CODE

                    Case 3
                        Dim COMPANY_NAME As String = ws.Cells(excelRow, COMPANY).Value & String.Empty
                        If COMPANY_NAME.Contains("=") Then
                            COMPANY_NAME = COMPANY_NAME.Substring(InStr(COMPANY_NAME, "=")).Trim
                        End If
                        COMPANY_NAME = COMPANY_NAME.Substring(0, Math.Min(dst.Tables(aswPOTSHIP1).Columns("COMPANY_NAME").MaxLength, COMPANY_NAME.Length))
                        rowPOTSHIP1.Item("COMPANY_NAME") = COMPANY_NAME

                        Dim BOL_NO As String = ws.Cells(excelRow, BOL).Value & String.Empty
                        If BOL_NO.Contains("=") Then
                            BOL_NO = BOL_NO.Substring(InStr(BOL_NO, "=")).Trim
                        End If
                        BOL_NO = BOL_NO.Substring(0, Math.Min(dst.Tables(aswPOTSHIP1).Columns("BOL_NO").MaxLength, BOL_NO.Length))
                        rowPOTSHIP1.Item("BOL_NO") = BOL_NO

                        Dim rowPOTSHIP1_LK As DataRow = LookUp("POTSHIP1", BOL_NO)
                        If rowPOTSHIP1_LK IsNot Nothing Then
                            MessageBox.Show($"BOL No {BOL_NO} was already imported.", "Import LOA", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If

                        Dim WHSE_NAME As String = ws.Cells(excelRow, WHSE).Value & String.Empty
                        If (dst.Tables("ICTWHSEL").Select("WHSE_DESC = '" & WHSE_NAME.ToUpper & "'").Length > 0) Then
                            loaWhse = dst.Tables("ICTWHSEL").Select("WHSE_DESC = '" & WHSE_NAME.ToUpper & "'")(0).Item("WHSE_CODE") & String.Empty
                        End If

                        dst.Tables(aswPOTSHIP1).Rows.Add(rowPOTSHIP1)

                    Case Else

                        If (ws.Cells(excelRow, CONTAINER).Value & String.Empty).ToString.ToUpper.Contains("CONTAINER") Then
                            Continue For
                        End If

                        If (ws.Cells(excelRow, CONTAINER).Value & String.Empty).ToString.Trim = String.Empty Then
                            Continue For
                        End If

                        Dim CONTAINER_NO As String = ws.Cells(excelRow, CONTAINER).Value & String.Empty
                        Dim CONTAINER_TYPE As String = ws.Cells(excelRow, CONT_TYPE).Value & String.Empty
                        Dim CONTAINER_CTL_NO As String = ASCMAIN1.Next_Control_No("POTSHIP2.CONTAINER_CTL_NO")

                        Dim rowPOTSHIP2 As DataRow = dst.Tables(aswPOTSHIP2).NewRow
                        rowPOTSHIP2.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                        rowPOTSHIP2.Item("CONTAINER_CTL_NO") = CONTAINER_CTL_NO

                        ' Default value
                        rowPOTSHIP2.Item("WHSE_CODE") = IIf(String.IsNullOrEmpty(loaWhse), WHSE_CODE_DEFAULT, loaWhse)

                        rowPOTSHIP2.Item("ASN_NO") = ASCMAIN1.Next_Control_No("POTASNM1.ASN_NO")
                        rowPOTSHIP2.Item("ASN_DATE") = DateTime.Now.ToShortDateString
                        rowPOTSHIP2.Item("ASN_REQUIRES_XMIT") = "1"

                        rowPOTSHIP2.Item("CONTAINER_NO") = CONTAINER_NO
                        rowPOTSHIP2.Item("CONTAINER_TYPE") = CONTAINER_TYPE
                        rowPOTSHIP2.Item("CONTAINER_LNO") = ws.Cells(excelRow, CONT_LNO).Value & String.Empty

                        rowPOTSHIP2.Item("INIT_DATE") = DateTime.Now
                        rowPOTSHIP2.Item("INIT_OPER") = ASCMAIN1.USER_ID

                        Dim NAVALOCK As String = (ws.Cells(excelRow, NAVA).Value & String.Empty).ToString.Trim
                        NAVALOCK = NAVALOCK.Substring(0, Math.Min(dst.Tables(aswPOTSHIP2).Columns("NAVALOCK").MaxLength, NAVALOCK.Length))

                        rowPOTSHIP2.Item("NAVALOCK") = NAVALOCK
                        dst.Tables(aswPOTSHIP2).Rows.Add(rowPOTSHIP2)

                        While (ws.Cells(excelRow, PACK_LIST).Value & String.Empty).ToString.Trim.ToUpper <> "TOTAL"
                            Dim INV_NUM As String = (ws.Cells(excelRow, INVOICE_NO).Value & String.Empty).ToString.Trim
                            If INV_NUM.Length > 0 Then
                                Dim rowPOTSHIP3 As DataRow = dst.Tables(aswPOTSHIP3).NewRow
                                rowPOTSHIP3.Item("PO_SHIPMENT_NO") = rowPOTSHIP2.Item("PO_SHIPMENT_NO")
                                rowPOTSHIP3.Item("CONTAINER_CTL_NO") = rowPOTSHIP2.Item("CONTAINER_CTL_NO")
                                rowPOTSHIP3.Item("PO_SHIPMENT_DNO") = Val(dst.Tables(aswPOTSHIP3).Compute("MAX(PO_SHIPMENT_DNO)", $"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}' AND CONTAINER_CTL_NO = {rowPOTSHIP3.Item("CONTAINER_CTL_NO")}") & String.Empty) + 1
                                rowPOTSHIP3.Item("INV_NUM") = INV_NUM
                                rowPOTSHIP3.Item("PACK_LIST_NO") = ws.Cells(excelRow, PACK_LIST).Value & String.Empty
                                rowPOTSHIP3.Item("PALLETS") = Val(ws.Cells(excelRow, PALLETS).Value & String.Empty)
                                rowPOTSHIP3.Item("CARTONS") = Val(ws.Cells(excelRow, CARTONS).Value & String.Empty)
                                rowPOTSHIP3.Item("WEIGHT") = Val(ws.Cells(excelRow, WEIGHT).Value & String.Empty)
                                rowPOTSHIP3.Item("SHIP_CUBE") = Val(ws.Cells(excelRow, CUBE).Value & String.Empty)
                                dst.Tables(aswPOTSHIP3).Rows.Add(rowPOTSHIP3)
                            End If

                            Dim SHIP_SEAL_NO As String = (ws.Cells(excelRow, SEAL).Value & String.Empty).ToString.Trim
                            If SHIP_SEAL_NO.Length > 0 Then
                                If rowPOTSHIP2.Item("SEAL_NO1") & String.Empty = String.Empty Then
                                    rowPOTSHIP2.Item("SEAL_NO1") = SHIP_SEAL_NO
                                ElseIf rowPOTSHIP2.Item("SEAL_NO2") & String.Empty = String.Empty Then
                                    rowPOTSHIP2.Item("SEAL_NO2") = SHIP_SEAL_NO
                                ElseIf rowPOTSHIP2.Item("SEAL_NO3") & String.Empty = String.Empty Then
                                    rowPOTSHIP2.Item("SEAL_NO3") = SHIP_SEAL_NO
                                ElseIf rowPOTSHIP2.Item("SEAL_NO4") & String.Empty = String.Empty Then
                                    rowPOTSHIP2.Item("SEAL_NO4") = SHIP_SEAL_NO
                                Else
                                    MessageBox.Show($"Currently the application permits 4 seals. Container {CONTAINER_NO} has more than 4 seals.", "Container Seal", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                End If
                            End If

                            excelRow += 1
                        End While
                End Select
            Next

            ws = Nothing
            oWB.Close()
            oWB = Nothing

            ' 10/18/2025 - Since the user will review the iomport before updating to Oracle = default to B for Boat
            rowPOTSHIP1.Item("SHIP_TYPE") = "B"
            ' Get the ship Type from the user
            'Using frmASFMSGBF As New ASFMSGBF
            '    Dim frtOption As Integer = frmASFMSGBF.Get_opt_from_User("Select ship Type", {"Air", "Boat", "Ground"}, 1, "Ship Type")
            '    Select Case frtOption
            '        Case 0
            '            rowPOTSHIP1.Item("SHIP_TYPE") = "A"
            '        Case 1
            '            rowPOTSHIP1.Item("SHIP_TYPE") = "B"
            '        Case 2
            '            rowPOTSHIP1.Item("SHIP_TYPE") = "G"
            '    End Select
            'End Using

            Try
                BeginTrans()

                Update_Record_TDA(aswPOTSHIP1)
                Update_Record_TDA(aswPOTSHIP2)
                Update_Record_TDA(aswPOTSHIP3)

                ASCMAIN1.sql = $"BEGIN DECLARE CURSOR C1 IS
                                    SELECT {aswPOTSHIP3}.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.INV_NUM INV_NUM_P, ICTPINV1.PACK_LIST_NO PACK_LIST_NO_P, ICTPINV1.PINV_NO, ICTPINV1.RECEIPT_NO RECEIPT_NO_I
                                    FROM {aswPOTSHIP3}, ICTPINV1
                                    WHERE ICTPINV1.INV_NUM LIKE '%' || {aswPOTSHIP3}.INV_NUM 
                                    AND ICTPINV1.PACK_LIST_NO LIKE '%' || {aswPOTSHIP3}.PACK_LIST_NO
                                    AND {aswPOTSHIP3}.PO_SHIPMENT_NO = :PARM1;
                                    BEGIN FOR R1 IN C1 LOOP
                                        UPDATE {aswPOTSHIP3} SET PACK_LIST_NO = NVL(R1.PACK_LIST_NO_P, PACK_LIST_NO), INV_NUM = NVL(R1.INV_NUM_P, INV_NUM), RECEIPT_NO = R1.RECEIPT_NO_I
                                        WHERE PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO 
                                        AND CONTAINER_CTL_NO = R1.CONTAINER_CTL_NO 
                                        AND PO_SHIPMENT_DNO = R1.PO_SHIPMENT_DNO 
                                        AND INV_NUM = R1.INV_NUM;                                    
                                END LOOP; END; END;"
                ' AND NVL(ICTPINV1.PINV_STATUS, 'O') = 'O'
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                ASCMAIN1.sql = $"BEGIN DECLARE CURSOR C1 IS
                                    SELECT {aswPOTSHIP3}.CONTAINER_CTL_NO, MIN(ICTPINV1.WHSE_CODE) WHSE_CODE
                                    FROM ICTPINV1, {aswPOTSHIP3}
                                    WHERE ICTPINV1.INV_NUM = {aswPOTSHIP3}.INV_NUM
                                    AND {aswPOTSHIP3}.PO_SHIPMENT_NO = :PARM1
                                    GROUP BY {aswPOTSHIP3}.CONTAINER_CTL_NO;
                                    BEGIN FOR R1 IN C1 LOOP
                                        UPDATE {aswPOTSHIP2} SET WHSE_CODE = R1.WHSE_CODE WHERE CONTAINER_CTL_NO = R1.CONTAINER_CTL_NO;
                                    END LOOP; END; END;"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})

                CommitTrans("Import Successful")
                processingImportedData = True

            Catch ex As Exception
                Rollback(ex.Message)
                processingImportedData = False
            End Try

        Catch ex As Exception
            MessageBox.Show("Import LOA", ex.Message)
        End Try

    End Sub

    Private Sub CreateAirBOL()
        Try
            processingImportedData = False

            For Each tablename As String In New String() {"POTSHIP1_IMP", "POTSHIP2_IMP", "POTSHIP3_IMP"}
                dst.Tables(tablename).Rows.Clear()
            Next

            For Each tablename As String In New String() {aswPOTSHIP1, aswPOTSHIP2, aswPOTSHIP3}
                ASCDATA1.ExecuteSQL($"DELETE FROM {tablename}")
                dst.Tables(tablename).Rows.Clear()
            Next

            Dim PO_SHIPMENT_NO As String = ASCMAIN1.Next_Control_No("POTSHIP1.PO_SHIPMENT_NO")
            rowPOTSHIP1 = dst.Tables(aswPOTSHIP1).NewRow
            rowPOTSHIP1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP1.Item("SHIP_DATE") = DateTime.Now.ToShortDateString
            rowPOTSHIP1.Item("VEND_CODE") = txtVEND_CODE.Text
            rowPOTSHIP1.Item("PO_SHIPMENT_STATUS") = "O"
            rowPOTSHIP1.Item("INIT_DATE") = DateTime.Now
            rowPOTSHIP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowPOTSHIP1.Item("LAST_DATE") = DateTime.Now
            rowPOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowPOTSHIP1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowPOTSHIP1.Item("SHIP_TYPE") = "A"
            dst.Tables(aswPOTSHIP1).Rows.Add(rowPOTSHIP1)


            Dim CONTAINER_CTL_NO As String = ASCMAIN1.Next_Control_No("POTSHIP2.CONTAINER_CTL_NO")

            Dim rowPOTSHIP2 As DataRow = dst.Tables(aswPOTSHIP2).NewRow
            rowPOTSHIP2.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
            rowPOTSHIP2.Item("CONTAINER_CTL_NO") = CONTAINER_CTL_NO
            rowPOTSHIP2.Item("WHSE_CODE") = WHSE_CODE_DEFAULT
            rowPOTSHIP2.Item("ASN_NO") = ASCMAIN1.Next_Control_No("POTASNM1.ASN_NO")
            rowPOTSHIP2.Item("ASN_DATE") = DateTime.Now.ToShortDateString
            rowPOTSHIP2.Item("ASN_REQUIRES_XMIT") = "1"
            rowPOTSHIP2.Item("CONTAINER_NO") = "CONT-1"
            rowPOTSHIP2.Item("CONTAINER_TYPE") = "Air"
            rowPOTSHIP2.Item("CONTAINER_LNO") = 1
            rowPOTSHIP2.Item("INIT_DATE") = DateTime.Now
            rowPOTSHIP2.Item("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables(aswPOTSHIP2).Rows.Add(rowPOTSHIP2)

            Try
                BeginTrans()

                Update_Record_TDA(aswPOTSHIP1)
                Update_Record_TDA(aswPOTSHIP2)
                Update_Record_TDA(aswPOTSHIP3)

                CommitTrans("Air Shipment Setup Successfully")
                processingImportedData = True

            Catch ex As Exception
                Rollback(ex.Message)
                processingImportedData = False
            End Try

        Catch ex As Exception
            MessageBox.Show("Create Air BOL", ex.Message)
        End Try

    End Sub

    Sub Refresh_Documents()

        Static Resizegrid As Boolean = True

        Try
            If Not dst.Tables.Contains("POTASNM3_X") Then
                Exit Sub
            End If

            EnforceConstraints(False)
            dst.Tables("POTASNM3_X").Rows.Clear()
            dst.Tables("POTASNM2_X").Rows.Clear()
            dst.Tables("POTSHIP3_X").Rows.Clear()
            dst.Tables("POTSHIP2_X").Rows.Clear()
            dst.Tables("POTSHIP1_X").Rows.Clear()

            Me.Cursor = Cursors.WaitCursor

            If processingImportedData Then
                grdPOTSHIP1_X.Text = "Imported Shipment"

                ASCMAIN1.sql = $"Select * from {aswPOTSHIP1}"
                Fill_Records("POTSHIP1_X", String.Empty, True, ASCMAIN1.sql)

                ASCMAIN1.sql = $"SELECT * FROM {aswPOTSHIP2}"
                Fill_Records("POTSHIP2_X", String.Empty, True, ASCMAIN1.sql)

                ASCMAIN1.sql = $"SELECT POTSHIP3.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.VESSEL_NAME, ICTPINV1.PINV_NO, ICTPINV1.WHSE_CODE, ICTPINV1.VEND_CODE, ICTPINV1.PINV_STATUS
                                FROM {aswPOTSHIP3} POTSHIP3, ICTPINV1
                                WHERE POTSHIP3.INV_NUM = ICTPINV1.INV_NUM (+)
                                AND POTSHIP3.PACK_LIST_NO = ICTPINV1.PACK_LIST_NO (+)"
                ' AND POTSHIP3.CONTAINER_CTL_NO = ICTPINV1.CONTAINER_CTL_NO (+)
                Fill_Records("POTSHIP3_X", String.Empty, True, ASCMAIN1.sql)

                EnforceConstraints(True)

                'grdPOTSHIP1_X.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
                Sort_grdColumns(grdPOTSHIP1_X, "BOL_NO")
                Clear_All_Filters(grdPOTSHIP1_X)

                With UltraExplorerBar1
                    With .Groups("Screen Control")
                        .Items("Update").Settings.Enabled = DefaultableBoolean.True
                        .Items("Cancel").Settings.Enabled = DefaultableBoolean.True
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                        .Items("View").Settings.Enabled = DefaultableBoolean.False
                        .Items("Done").Settings.Enabled = DefaultableBoolean.False
                        .Items("Import Manifest").Settings.Enabled = DefaultableBoolean.False
                        .Items("Refresh").Settings.Enabled = DefaultableBoolean.False
                        .Items("Create Air BOL").Settings.Enabled = DefaultableBoolean.False
                    End With
                End With

                toggleAllowEdit(True)
                For Each row As UltraGridRow In grdPOTSHIP1_X.Rows
                    If row.Band.Key = grdPOTSHIP1_X.DisplayLayout.Bands(0).Key Then
                        row.Expanded = True
                    Else
                        row.Expanded = False
                    End If
                Next

                Exit Sub
            End If

            Select Case optRecords.Value
                Case "O"
                    Fill_Records("POTSHIP1_X", optRecords.Value)
                Case "P"
                    Fill_Records("POTSHIP1_X", optRecords.Value)
                Case "Y"
                    ASCMAIN1.sql = "Select * from POTSHIP1 where OPS_YYYYPP = :PARM1"
                    Fill_Records("POTSHIP1_X", {cbeYP.Value}, True, ASCMAIN1.sql)
                Case "A"
                    Dim ASN_NO As String = txtBOL_ASN_NO.Text.ToUpper.Substring(1).PadLeft(10, "0")
                    ASCMAIN1.sql = "Select * from POTSHIP1 where PO_SHIPMENT_NO IN (SELECT PO_SHIPMENT_NO FROM POTSHIP2 WHERE ASN_NO = :PARM1)"
                    Fill_Records("POTSHIP1_X", {ASN_NO}, True, ASCMAIN1.sql)
                Case "B"
                    Dim BOL_NO As String = txtBOL_ASN_NO.Text.ToUpper
                    ASCMAIN1.sql = "Select * from POTSHIP1 where UPPER(BOL_NO) = :PARM1"
                    Fill_Records("POTSHIP1_X", {BOL_NO}, True, ASCMAIN1.sql)
                Case "I"
                    Dim INV_NUM As String = txtBOL_ASN_NO.Text.ToUpper
                    ASCMAIN1.sql = "Select * from POTSHIP1 where PO_SHIPMENT_NO IN (SELECT PO_SHIPMENT_NO FROM POTSHIP3 WHERE INV_NUM = :PARM1)"
                    Fill_Records("POTSHIP1_X", {INV_NUM}, True, ASCMAIN1.sql)
                Case Else
                    Exit Sub
            End Select

            Dim sql = String.Empty

            Select Case optRecords.Value
                Case "O"
                    grdPOTSHIP1_X.Text = "Open Shipments"

                Case "P"
                    grdPOTSHIP1_X.Text = "Transmitted Shipments"

                Case "Y"
                    grdPOTSHIP1_X.Text = $"Shipments created in period {cbeYP.Value}"

                Case "A"
                    grdPOTSHIP1_X.Text = $"ASN {txtBOL_ASN_NO.Text}"

                Case "B"
                    grdPOTSHIP1_X.Text = $"BOL No {txtBOL_ASN_NO.Text.ToUpper}"

                Case "I"
                    grdPOTSHIP1_X.Text = $"Invoice No {txtBOL_ASN_NO.Text.ToUpper}"

                Case Else
                    grdPOTSHIP1_X.Text = "No LOAs"
            End Select

            Dim lstPO_SHIPMENT_NOs As New List(Of String)
            For Each drPOTSHIP1 As DataRow In dst.Tables("POTSHIP1_X").Select("")
                lstPO_SHIPMENT_NOs.Add(drPOTSHIP1.Item("PO_SHIPMENT_NO"))
            Next
            Fill_Records("POTSHIP2_X", String.Join(",", lstPO_SHIPMENT_NOs.ToArray))
            Fill_Records("POTSHIP3_X", String.Join(",", lstPO_SHIPMENT_NOs.ToArray))

            Dim lstASN_NOs As New List(Of String)
            For Each drPOTSHIP2 As DataRow In dst.Tables("POTSHIP2_X").Select("")
                Dim ASN_NO As String = drPOTSHIP2.Item("ASN_NO") & String.Empty
                If ASN_NO.Length > 0 Then
                    lstASN_NOs.Add(ASN_NO)
                End If
            Next

            Fill_Records("POTASNM2_X", String.Join(",", lstASN_NOs.ToArray))
            Fill_Records("POTASNM3_X", String.Join(",", lstASN_NOs.ToArray))

            ' Resize only once
            If dst.Tables("POTSHIP1_X").Rows.Count > 0 Then
                If Resizegrid Then
                    grdPOTSHIP1_X.DisplayLayout.PerformAutoResizeColumns(False, PerformAutoSizeType.AllRowsInBand, True)
                    Resizegrid = False
                End If
            End If

            Sort_grdColumns(grdPOTSHIP1_X, "BOL_NO")
            Clear_All_Filters(grdPOTSHIP1_X)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            dst.Tables("POTASNM3_X").Rows.Clear()
            dst.Tables("POTASNM2_X").Rows.Clear()
            dst.Tables("POTSHIP3_X").Rows.Clear()
            dst.Tables("POTSHIP2_X").Rows.Clear()
            dst.Tables("POTSHIP1_X").Rows.Clear()
        Finally
            Me.Cursor = Cursors.Default
            EnforceConstraints(True)
        End Try
    End Sub

    Private Sub UpdateLOAImportData()

        Update_Record_TDA("POTSHIP1_IMP")
        Update_Record_TDA("POTSHIP2_IMP")
        Update_Record_TDA("POTSHIP3_IMP")

        For Each drPOTSHIP1 As DataRow In dst.Tables("POTSHIP1_IMP").Select("")
            Dim PO_SHIPMENT_NO As String = drPOTSHIP1.Item("PO_SHIPMENT_NO")

            ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS
                                    SELECT POTSHIP3.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.INV_NUM INV_NUM_P, ICTPINV1.PACK_LIST_NO PACK_LIST_NO_P, ICTPINV1.PINV_NO, ICTPINV1.RECEIPT_NO RECEIPT_NO_I, POTSHIP2.ASN_NO
                                    FROM POTSHIP3, ICTPINV1, POTSHIP2
                                    WHERE ICTPINV1.INV_NUM LIKE '%' || POTSHIP3.INV_NUM 
                                    AND ICTPINV1.PACK_LIST_NO LIKE '%' || POTSHIP3.PACK_LIST_NO
                                    AND POTSHIP2.PO_SHIPMENT_NO = POTSHIP3.PO_SHIPMENT_NO
                                    AND POTSHIP3.PO_SHIPMENT_NO = :PARM1;
                                    BEGIN FOR R1 IN C1 LOOP
                                        UPDATE POTSHIP3 SET PACK_LIST_NO = NVL(R1.PACK_LIST_NO_P, PACK_LIST_NO), INV_NUM = NVL(R1.INV_NUM_P, INV_NUM), RECEIPT_NO = R1.RECEIPT_NO_I
                                        WHERE PO_SHIPMENT_NO = R1.PO_SHIPMENT_NO 
                                        AND CONTAINER_CTL_NO = R1.CONTAINER_CTL_NO 
                                        AND PO_SHIPMENT_DNO = R1.PO_SHIPMENT_DNO 
                                        AND INV_NUM = R1.INV_NUM;
                                        UPDATE ICTPINV1 SET CONTAINER_CTL_NO = R1.CONTAINER_CTL_NO, ASN_NO = R1.ASN_NO WHERE PINV_NO = R1.PINV_NO AND CONTAINER_CTL_NO IS NULL;
                                END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})
            ' AND NVL(PINV_STATUS, 'O') = 'O'

            ' Need updated INV_NUM
            Fill_Records("POTSHIP3_IMP", String.Empty, True, $"SELECT * FROM POTSHIP3 WHERE PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'")

            For Each drPOTSHIP2 As DataRow In dst.Tables("POTSHIP2_IMP").Select($"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'")
                Dim CONTAINER_CTL_NO As String = drPOTSHIP2.Item("CONTAINER_CTL_NO") & String.Empty
                If CONTAINER_CTL_NO = String.Empty Then
                    CONTAINER_CTL_NO = ASCMAIN1.Next_Control_No("POTSHIP2.CONTAINER_CTL_NO")
                    drPOTSHIP2.Item("CONTAINER_CTL_NO") = CONTAINER_CTL_NO
                End If
                Dim WHSE_CODE As String = drPOTSHIP2.Item("WHSE_CODE") & String.Empty

                For Each drPOTSHIP3 As DataRow In dst.Tables("POTSHIP3_IMP").Select($"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}' AND CONTAINER_CTL_NO = '{CONTAINER_CTL_NO}'")
                    Dim INV_NUM As String = drPOTSHIP3.Item("INV_NUM") & String.Empty

                    If CONTAINER_CTL_NO.Length > 0 AndAlso WHSE_CODE.Length > 0 AndAlso INV_NUM.Length > 0 Then
                        Dim drICTPINV1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTPINV1 WHERE CONTAINER_CTL_NO = :PARM1 AND INV_NUM = :PARM2", "VV", {CONTAINER_CTL_NO, INV_NUM})
                        If drICTPINV1 IsNot Nothing Then
                            Dim PINV_NO As String = drICTPINV1.Item("PINV_NO") & String.Empty
                            Dim WHSE_CODE_ICTPINV1 As String = drICTPINV1.Item("WHSE_CODE") & String.Empty
                            If WHSE_CODE <> WHSE_CODE_ICTPINV1 Then
                                Write_Event_Log("ICTPINV1", PINV_NO, $"Change Warehouse from {WHSE_CODE_ICTPINV1} to {WHSE_CODE}")
                                ASCMAIN1.sql = "UPDATE ICTPINV1 SET WHSE_CODE = :PARM1 WHERE PINV_NO = :PARM2"
                                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {WHSE_CODE, PINV_NO})
                            End If
                        End If
                    End If
                Next
            Next
        Next

        Update_Record_TDA("POTSHIP2_IMP")

    End Sub

    Private Sub toggleAllowEdit(ByVal Editable As Boolean)

        ' Lock everything down
        grdPOTSHIP1_X.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdPOTSHIP1_X.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdPOTSHIP1_X.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        grdPOTSHIP1_X.DisplayLayout.Bands(1).Override.AllowAddNew = AllowAddNew.No
        grdPOTSHIP1_X.DisplayLayout.Bands(1).Override.AllowDelete = DefaultableBoolean.False
        grdPOTSHIP1_X.DisplayLayout.Bands(2).Override.AllowDelete = DefaultableBoolean.False

        ' Only allowing the columns that are editable on the Invoice Screen.
        For Each uBand As UltraGridBand In grdPOTSHIP1_X.DisplayLayout.Bands
            For Each gCol As UltraGridColumn In uBand.Columns
                If editableColumns.Contains(gCol.Key) AndAlso Editable AndAlso uBand.Key <> grdPOTSHIP1_X.DisplayLayout.Bands(2).Key Then
                    gCol.CellActivation = Activation.AllowEdit
                    gCol.CellAppearance.BackColor = Drawing.Color.LightGreen
                Else
                    gCol.CellActivation = Activation.NoEdit
                    gCol.CellAppearance.BackColor = Nothing
                End If
            Next
        Next

        If Not Editable Then
            Exit Sub
        End If

        grdPOTSHIP1_X.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

        ' Container Level
        'grdPOTSHIP1_X.DisplayLayout.Bands(1).Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop
        grdPOTSHIP1_X.DisplayLayout.Bands(1).Override.AllowDelete = DefaultableBoolean.True

        ' Invoice Level
        grdPOTSHIP1_X.DisplayLayout.Bands(2).Override.AllowDelete = DefaultableBoolean.True

    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdPOTSHIP1_X_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdPOTSHIP1_X.DoubleClickRow

        If processingImportedData Then
            Exit Sub
        End If

        If e.Row.Band.Key = grdPOTSHIP1_X.DisplayLayout.Bands(1).Key Then
            PO_SHIPMENT_NO = e.Row.Cells("PO_SHIPMENT_NO").Value & String.Empty
            CONTAINER_CTL_NO = e.Row.Cells("CONTAINER_CTL_NO").Value & String.Empty
            Click_Command("View")
        End If
    End Sub

    Private Sub grdPOTSHIP1_X_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdPOTSHIP1_X.InitializeLayout
        e.Layout.Bands(1).SummaryFooterCaption = "Container Totals"
        e.Layout.Bands(2).SummaryFooterCaption = "Invoice Totals"
    End Sub

    Private Sub grdICTPINV1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTPINV1.InitializeRow

        Dim PINV_NO As String = e.Row.Cells("PINV_NO").Value & String.Empty
        If dst.Tables("POTSHIP3").Select($"PINV_NO = '{PINV_NO}'").Length > 0 Then
            e.Row.Cells("PINV_NO").Appearance.BackColor = Drawing.Color.LightGreen
        End If

    End Sub

#End Region

#Region "Table Definitions"

    Private Sub ValidateTablesExist()

        For Each tableName As String In {"POTSHIP1", "POTSHIP2", "POTSHIP3", "POTASNM1", "POTASNM2", "POTASNM3"}

            ASCMAIN1.sql = "Select * from USER_TABLES WHERE TABLE_NAME = :PARM1"
            Dim rowTable As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", {tableName})
            If rowTable IsNot Nothing Then
                Continue For
            End If

            Select Case tableName
                Case "POTSHIP1"
                    ASCMAIN1.sql = "CREATE TABLE POTSHIP1(
                                PO_SHIPMENT_NO VARCHAR2(10),
                                VEND_CODE VARCHAR2(10),
                                VESSEL_NAME VARCHAR2(50),
                                VOYAGE_NO VARCHAR2(25),
                                BOL_NO VARCHAR2(20),
                                SCAC_CODE VARCHAR2(20),
                                COMPANY_NAME VARCHAR2(50),
                                SHIP_DATE DATE,
                                ETA_DATE DATE,
                                SHIP_TYPE VARCHAR2(1),
                                FILENAME VARCHAR2(250),
                                COUNTRY_CODE VARCHAR2(3),
                                PO_SHIPMENT_STATUS VARCHAR2(1),
                                INIT_DATE DATE,
                                INIT_OPER VARCHAR2(20),
                                LAST_DATE DATE,
                                LAST_OPER VARCHAR2(20),
                                OPS_YYYYPP VARCHAR2(6),
                                PORT_CODE VARCHAR2(6),
                                NOTES VARCHAR2(300),
                                PRIMARY KEY (PO_SHIPMENT_NO))"

                Case "POTSHIP2"
                    ASCMAIN1.sql = "CREATE TABLE POTSHIP2(
                                PO_SHIPMENT_NO VARCHAR2(10),
                                CONTAINER_CTL_NO VARCHAR2(10),
                                CONTAINER_NO VARCHAR2(20),
                                CONTAINER_LNO VARCHAR2(20),
				                WHSE_CODE VARCHAR2(6),
                                NAVALOCK VARCHAR2(20),
                                CONTAINER_TYPE VARCHAR2(20),
                                ASN_NO VARCHAR2(10),
                                ASN_DATE DATE,
                                ASN_REQUIRES_XMIT VARCHAR2(1),
                                INIT_DATE DATE,
                                INIT_OPER VARCHAR2(20),
                                LAST_DATE DATE,
                                LAST_OPER VARCHAR2(20),
                                SEAL_NO1 VARCHAR2(20),
                                SEAL_NO2 VARCHAR2(20),
                                SEAL_NO3 VARCHAR2(20),
                                SEAL_NO4 VARCHAR2(20),
                                PRIMARY KEY (PO_SHIPMENT_NO, CONTAINER_CTL_NO))"

                Case "POTSHIP3"
                    ASCMAIN1.sql = "CREATE TABLE POTSHIP3(
                                PO_SHIPMENT_NO VARCHAR2(10),
                                CONTAINER_CTL_NO VARCHAR2(10),
                                PO_SHIPMENT_DNO NUMBER(4),
                                INV_NUM VARCHAR2(20),
                                PACK_LIST_NO VARCHAR2(20),
                                RECEIPT_NO VARCHAR2(6),
                                PALLETS NUMBER(6),
                                CARTONS NUMBER(6),
                                WEIGHT NUMBER(12, 3),
                                SHIP_CUBE NUMBER(12,3),
				                PRIMARY KEY (PO_SHIPMENT_NO, CONTAINER_CTL_NO, PO_SHIPMENT_DNO))"

                Case "POTASNM1"
                    ASCMAIN1.sql = "CREATE TABLE POTASNM1(
                                ASN_NO VARCHAR2(10),
                                CONTAINER_CTL_NO VARCHAR2(10),
                                PRIMARY KEY(ASN_NO))"

                Case "POTASNM2"
                    ASCMAIN1.sql = "CREATE TABLE POTASNM2(
                                ASN_NO VARCHAR2(10),
                                ASN_XMIT_NO VARCHAR2(10),
                                INIT_DATE DATE,
                                INIT_OPER VARCHAR2(20),
                                TRANSMIT_DATE DATE,
                                TRANSMIT_OPER VARCHAR2(20),
                                PRIMARY KEY (ASN_NO, ASN_XMIT_NO))"

                Case "POTASNM3"
                    ASCMAIN1.sql = "CREATE TABLE POTASNM3(
                                ASN_NO VARCHAR2(10),
                                ASN_XMIT_NO VARCHAR2(10),
                                ASN_LNO NUMBER(5),
                                ITEM_CODE VARCHAR2(25),
                                ASN_QTY NUMBER(8),
                                ASN_COST NUMBER(12,6),
                                PRIMARY KEY (ASN_NO, ASN_XMIT_NO, ASN_LNO))"
            End Select

            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        Next

        Try
            ASCMAIN1.sql = "Select * from USER_TAB_COLUMNS WHERE TABLE_NAME = :PARM1"
            Dim tblUTC As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {"ICTPINV1"})
            If tblUTC.Select("COLUMN_NAME = 'CONTAINER_CTL_NO'").Length = 0 Then
                ASCMAIN1.sql = "ALTER TABLE ICTPINV1 ADD CONTAINER_CTL_NO VARCHAR2(10)"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            End If

            If tblUTC.Select("COLUMN_NAME = 'ASN_NO'").Length = 0 Then
                ASCMAIN1.sql = "ALTER TABLE ICTPINV1 ADD ASN_NO VARCHAR2(10)"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            End If
        Catch ex As Exception
            Stop
        End Try
    End Sub

    Private Sub grdPOTSHIP1_X_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTSHIP1_X.InitializeRow
        If e.Row.Band.Columns.Exists("ASN_REQUIRES_XMIT") Then
            If e.Row.Cells("ASN_REQUIRES_XMIT").Value & String.Empty = "1" Then
                e.Row.Cells("ASN_REQUIRES_XMIT").Appearance.BackColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Private Sub grdPOTSHIP1_X_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdPOTSHIP1_X.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdPOTSHIP1_X.ActiveCell.Column.Key
            Case "WHSE_CODE"
                sql_where = "LP_CODE IS NOT NULL"
        End Select

        grdClickCellButton(grdPOTSHIP1_X, sql_where, False)
    End Sub

    Private Sub optRecords_ValueChanged(sender As Object, e As EventArgs) Handles optRecords.ValueChanged
        Refresh_Documents()
    End Sub

    Private Sub txtBOL_ASN_NO_KeyDown(sender As Object, e As KeyEventArgs) Handles txtBOL_ASN_NO.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case optRecords.Value
                Case "A", "B"
                    Refresh_Documents()
            End Select
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(sender As Object, e As EventArgs) Handles cbeYP.ValueChanged
        Select Case optRecords.Value
            Case "Y"
                Refresh_Documents()
        End Select
    End Sub

#End Region

End Class