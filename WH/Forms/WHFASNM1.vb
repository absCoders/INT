Imports ABSolution
Imports Infragistics.Win.UltraWinGrid

' ISSUE-7043 Creation of ASNs instead of invoices 

Public Class WHFASNM1

    Private PO_SHIPMENT_NO As String
    Private WHSE_CODE As String
    Private VEND_CODE As String

    Private rowWHTASNM1 As DataRow
    Private BOL_NO As String
    Private CONTAINER_NO As String

    Private processingContainer As Boolean = False

    Private Const WHSE_CODE_DEFAULT As String = "CLA"
    Private Const VEND_CODE_DEFAULT As String = "IPSA"
    Private tblHeader As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ' Get_PARM("WHTPARM1")
        InquiryMode = MENU_ITEM_OBJECT = "WHFASNMI"

        If ASCMAIN1.Running_in_VS Then
            ValidateTablesExist()
        End If

        With dst

            Create_TDA(.Tables.Add, "WHTASNM1", "*")
            Create_TDA(.Tables.Add, "WHTASNM2", "*", 1)
            AUDIT.Add("WHTASNM1", "E")
            AUDIT.Add("WHTASNM2", "E")

            ASCMAIN1.sql = "Select * from WHTASNM1 where PO_SHIPMENT_STATUS = :PARM1"
            Create_TDA(.Tables.Add, "WHTASNMX", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "SELECT * FROM ICTPINV1 
                            WHERE (PO_SHIPMENT_NO = :PARM1 
                            OR 
                            (
                            VEND_CODE = :PARM2 
                            AND PO_SHIPMENT_NO IS NULL 
                            AND (INV_NUM, PACK_LIST_NO) NOT IN (SELECT INV_NUM, PACK_LIST_NO FROM WHTMANI3 WHERE BOL_NO <> :PARM3)
                            AND PINV_STATUS = 'O'
                            AND WHSE_CODE = :PARM4
                            ))"
            Create_TDA(.Tables.Add, "ICTPINV1", ASCMAIN1.sql, 0, True, "VVVV")
            AUDIT.Add("ICTPINV1", "E")
            Create_TDA(.Tables.Add("ICTPINV1_S"), "ICTPINV1", "*")

            tblHeader = ASCMAIN1.Temp_Table("SELECT BOL_NO FROM WHTMANI1 WHERE ROWNUM < 10")

            Create_TDA(.Tables.Add, "WHTMANI1", "*")
            Create_TDA(.Tables.Add, "WHTMANI2", "*")

            Create_TDA(.Tables.Add, "WHTMANI3", "*")
            .Tables("WHTMANI3").Columns.Add("PINV_NO", GetType(String))
            .Tables("WHTMANI3").Columns.Add("PO_ORDER_NO", GetType(String))
            .Tables("WHTMANI3").Columns.Add("PINV_STATUS", GetType(String))

            Create_TDA(.Tables.Add, "WHTMANI4", "*")
            Create_TDA(.Tables.Add("WHTMANI1_IMP"), "WHTMANI1", "*")
            Create_TDA(.Tables.Add("WHTMANI2_IMP"), "WHTMANI2", "*")
            Create_TDA(.Tables.Add("WHTMANI3_IMP"), "WHTMANI3", "*")
            Create_TDA(.Tables.Add("WHTMANI4_IMP"), "WHTMANI4", "*")
        End With

        Create_Relation("WHTMANI2", "WHTMANI4", "BOL_NO,CONTAINER_NO")
        Create_Relation("WHTMANI1", "WHTMANI2", "BOL_NO")
        Create_Relation("WHTMANI2", "WHTMANI3", "BOL_NO,CONTAINER_NO")
        With dst.Tables("WHTMANI2")
            .Columns.Add("NUM_INVOICES", GetType(Int32), "COUNT(CHILD(WHTMANI2_WHTMANI3).INV_NUM)")
            .Columns.Add("NUM_PALLETS", GetType(Int32), "SUM(CHILD(WHTMANI2_WHTMANI3).PALLETS)")
            .Columns.Add("NUM_CARTONS", GetType(Int32), "SUM(CHILD(WHTMANI2_WHTMANI3).CARTONS)")
            .Columns.Add("TOT_WEIGHT", GetType(Decimal), "SUM(CHILD(WHTMANI2_WHTMANI3).WEIGHT)")
            .Columns.Add("TOT_CUBE", GetType(Decimal), "SUM(CHILD(WHTMANI2_WHTMANI3).SHIP_CUBE)")
            .Columns.Add("NUM_SEALS", GetType(Decimal), "COUNT(CHILD(WHTMANI2_WHTMANI4).SHIP_SEAL_NO)")
        End With

        With dst.Tables("WHTMANI1")
            .Columns.Add("NUM_INVOICES", GetType(Int32), "SUM(CHILD(WHTMANI1_WHTMANI2).NUM_INVOICES)")
            .Columns.Add("NUM_PALLETS", GetType(Int32), "SUM(CHILD(WHTMANI1_WHTMANI2).NUM_PALLETS)")
            .Columns.Add("NUM_CARTONS", GetType(Int32), "SUM(CHILD(WHTMANI1_WHTMANI2).NUM_CARTONS)")
            .Columns.Add("TOT_WEIGHT", GetType(Decimal), "SUM(CHILD(WHTMANI1_WHTMANI2).TOT_WEIGHT)")
            .Columns.Add("TOT_CUBE", GetType(Decimal), "SUM(CHILD(WHTMANI1_WHTMANI2).TOT_CUBE)")
            .Columns.Add("NUM_SEALS", GetType(Int32), "SUM(CHILD(WHTMANI1_WHTMANI2).NUM_SEALS)")
            .Columns.Add("NUM_CONTAINERS", GetType(Int32), "COUNT(CHILD(WHTMANI1_WHTMANI2).CONTAINER_NO)")
        End With

        grdWHTMANI1.DataSource = dst.Tables("WHTMANI1")
        Create_Summary(grdWHTMANI1, "BOL_NO", "Count")
        ASCMAIN1.Add_Value_List(grdWHTMANI1, "SHIP_TYPE", Nothing, {":", "B:Boat", "A:Air", "G:Ground"})
        ASCMAIN1.Add_Value_List(grdWHTMANI1, "PINV_STATUS", Nothing, {":", "O:Open", "I:Invoiced", "C:Closed, R:Received"}, 3)

        grdWHTASNM2.DataSource = dst.Tables("WHTASNM2")

        grdWHTASNMX.DataSource = dst.Tables("WHTASNMX")
        Create_Summary(grdWHTASNMX, "PO_SHIPMENT_NO", "Count")
        ASCMAIN1.Add_Value_List(grdWHTASNMX, "PO_SHIPMENT_STATUS",, {":", "O:Open", "P:Transmitted", "C:Completed"})

        grdICTPINV1_A.DataSource = dst.Tables("ICTPINV1")
        Create_Summary(grdICTPINV1_A, "PINV_NO", "Count")

        grdICTPINV1_S.DataSource = dst.Tables("ICTPINV1_S")
        Create_Summary(grdICTPINV1_S, "PINV_NO", "Count")

        grdASTAUDT1.DataSource = dst.Tables("ASTAUDT1")
        For Each col As UltraGridColumn In grdASTAUDT1.DisplayLayout.Bands(0).Columns
            col.Header.Caption = StrConv(col.Header.Caption.Replace("_", " "), VbStrConv.ProperCase)
        Next

        grdWHTASNMX.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdWHTASNMX.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        grdWHTASNMX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

        dteETA_DATE.MinDate = CDate("01/01/2010")
        dteETA_DATE.MaxDate = DateAdd(DateInterval.Year, 3, DateTime.Now)

        dteSHIP_DATE.MinDate = dteETA_DATE.MinDate
        dteSHIP_DATE.MaxDate = dteETA_DATE.MaxDate

        TABLE_NAME = "WHTASNM1"
        Bind_Controls(grpDetails, "WHTASNM1")
        Bind_Controls(UltraGroupBox1, "WHTASNM1")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "New"

                Validate_Code("WHSE_CODE")
                Validate_Code("VEND_CODE")

                If EMsg.Length > 0 Then
                    Exit Select
                End If

                If Absx1.txtFor("WHSE_CODE").Text <> "CLA" Then
                    EMsg &= vbCr & "Currently only 3PL CLA is supported"
                End If

                If Absx1.txtFor("VEND_CODE").Text <> "IPSA" Then
                    EMsg &= vbCr & "Currently only Vendor IPSA is supported"
                End If

                If processingContainer Then
                    Dim rowWHTMANI2 As DataRow = LookUp("WHTMANI2", {BOL_NO, CONTAINER_NO})
                    If rowWHTMANI2 Is Nothing Then
                        EMsg &= vbCr & $"Cannot locate BOL No {BOL_NO}, Container {CONTAINER_NO}"
                    ElseIf rowWHTMANI2.Item("PO_SHIPMENT_NO") & String.Empty <> String.Empty Then
                        EMsg &= vbCr & $"BOL No {BOL_NO}, Container {CONTAINER_NO} is already assigned an PO Shipment No."
                    Else
                        ASCMAIN1.sql = "SELECT WHTMANI3.*, ICTPINV1.PINV_NO, ICTPINV1.PO_ORDER_NO, ICTPINV1.PINV_STATUS
                                FROM WHTMANI3, ICTPINV1
                                WHERE WHTMANI3.INV_NUM = ICTPINV1.INV_NUM (+)
                                AND WHTMANI3.PACK_LIST_NO = ICTPINV1.PACK_LIST_NO (+)
                                AND WHTMANI3.BOL_NO = :PARM1
                                AND WHTMANI3.CONTAINER_NO = :PARM2"
                        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", {BOL_NO, CONTAINER_NO})
                        For Each row As DataRow In tbl.Select("")
                            If row.Item("PO_ORDER_NO") & String.Empty = String.Empty Then
                                EMsg &= vbCr & "There are Invoices that are not assigned to Purchase Orders."
                                Exit Select
                            ElseIf row.Item("PINV_STATUS") & String.Empty <> "O" Then
                                EMsg &= vbCr & "There are Invoices that are not Open."
                                Exit Select
                            End If
                        Next
                    End If
                End If

            Case "View", "Edit"

                PO_SHIPMENT_NO = Absx1.txtFor("PO_SHIPMENT_NO").Text

                If eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("WHTASNM1", PO_SHIPMENT_NO) Then
                        Exit Sub
                    End If
                End If

                rowWHTASNM1 = Fill_Record("WHTASNM1", PO_SHIPMENT_NO)

                If rowWHTASNM1 Is Nothing Then
                    EMsg &= vbCr & "Missing or invalid PO Shipment No"
                ElseIf rowWHTASNM1.Item("PO_SHIPMENT_STATUS") & String.Empty = "C" AndAlso eItemKey = "Edit" Then
                    EMsg &= vbCr & "You may NOT Edit a Completed ASN"
                End If

                If eItemKey = "Edit" Then
                    If EMsg.Length > 0 Then
                        ASCMAIN1.MultiTask_Release()
                    End If
                End If

            Case "Done"

            Case "Update"
                If dst.Tables("ICTPINV1_S").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
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
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

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
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Refresh"
                Refresh_Documents()

            Case "Import Manifest"
                ImportManifest()
                Refresh_Documents()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowWHTASNM1.Item("PO_SHIPMENT_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If

                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode

                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Refresh").Settings.Enabled = not_iScreenMode

                .Items("Import Manifest").Settings.Enabled = not_iScreenMode

            End With
            .Groups("Show if Entered in").Visible = Not ScreenMode
        End With

        lblStatus.Visible = ScreenMode
        splDetails.Visible = ScreenMode

        tabOptions.Visible = Not ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            Select Case EntryMode
                Case "V"
                    grdWHTASNM2.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
                    grdWHTASNM2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    grdWHTASNM2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                    Set_Read_Only(splVessel.Panel1, True)
                    txtNOTES.ReadOnly = True

                Case Else
                    grdWHTASNM2.DisplayLayout.Override.AllowAddNew = AllowAddNew.FixedAddRowOnTop
                    grdWHTASNM2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                    grdWHTASNM2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

                    Set_Read_Only(splVessel.Panel1, False)
                    txtNOTES.ReadOnly = False

            End Select

        Else
            Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)

        For Each tablename As String In New String() {"WHTASNM1", "WHTASNM2", "ASTAUDT1",
                        "ICTPINV1", "ICTPINV1_S",
                        "WHTMANI1", "WHTMANI2", "WHTMANI3", "WHTMANI4",
                        "WHTMANI1_IMP", "WHTMANI2_IMP", "WHTMANI3_IMP", "WHTMANI4_IMP"}
            dst.Tables(tablename).Rows.Clear()
        Next

        Refresh_Documents()
        rowWHTASNM1 = Nothing
        EnforceConstraints(True)

        PO_SHIPMENT_NO = String.Empty
        BOL_NO = String.Empty
        VEND_CODE = String.Empty
        processingContainer = False

        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE_DEFAULT
        Absx1.txtFor("VEND_CODE").Text = VEND_CODE_DEFAULT

    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        VEND_CODE = HFs("VEND_CODE")
        WHSE_CODE = HFs("WHSE_CODE")

        If EntryMode = "N" Then
            PO_SHIPMENT_NO = ASCMAIN1.Next_Control_No("WHTASNM1.PO_SHIPMENT_NO")
            rowWHTASNM1 = dst.Tables("WHTASNM1").NewRow
            If processingContainer Then
                Fill_Records("WHTMANI1", "", True, $"SELECT * FROM WHTMANI1 WHERE BOL_NO = '{BOL_NO}'")
                Fill_Records("WHTMANI2", "", True, $"SELECT * FROM WHTMANI2 WHERE BOL_NO = '{BOL_NO}' AND CONTAINER_NO = '{CONTAINER_NO}'")
                Fill_Records("WHTMANI3", "", True, $"SELECT * FROM WHTMANI3 WHERE BOL_NO = '{BOL_NO}' AND CONTAINER_NO = '{CONTAINER_NO}'")
                Fill_Records("WHTMANI4", "", True, $"SELECT * FROM WHTMANI4 WHERE BOL_NO = '{BOL_NO}' AND CONTAINER_NO = '{CONTAINER_NO}'")

                Dim rowWHTMANI2 As DataRow = dst.Tables("WHTMANI2").Rows.Find({BOL_NO, CONTAINER_NO})
                rowWHTMANI2.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO

                Dim rowWHTMANI1 As DataRow = dst.Tables("WHTMANI1").Rows.Find(BOL_NO)
                With rowWHTASNM1
                    .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("VEND_CODE") = VEND_CODE
                    .Item("PO_SHIPMENT_STATUS") = "O"
                    .Item("VESSEL_NAME") = rowWHTMANI1.Item("VESSEL_NAME")
                    .Item("VOYAGE_NO") = rowWHTMANI1.Item("VOYAGE_NO")
                    .Item("BOL_NO") = rowWHTMANI1.Item("BOL_NO")
                    BOL_NO = rowWHTMANI1.Item("BOL_NO") & String.Empty
                    .Item("CONTAINER_NO") = rowWHTMANI2.Item("CONTAINER_NO")
                    .Item("IDOC_FILENAME") = rowWHTMANI1.Item("FILENAME")

                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DateTime.Now
                    .Item("SHIP_DATE") = rowWHTMANI1.Item("SHIP_DATE")
                    .Item("ETA_DATE") = rowWHTMANI1.Item("ETA_DATE")
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    .Item("SHIP_TYPE") = rowWHTMANI1.Item("SHIP_TYPE")
                End With

                ' As per the meeting on 01/15/2024 - permit adding other invoices.
                Fill_Records("ICTPINV1", New String() {CONTAINER_NO, VEND_CODE, BOL_NO, WHSE_CODE})

                For Each rowWHTMANI4 As DataRow In dst.Tables("WHTMANI4").Select("")
                    Dim rowWHTASNM2 As DataRow = dst.Tables("WHTASNM2").NewRow
                    rowWHTASNM2.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                    rowWHTASNM2.Item("SHIP_SEAL_NO") = rowWHTMANI4.Item("SHIP_SEAL_NO")
                    dst.Tables("WHTASNM2").Rows.Add(rowWHTASNM2)
                Next

                For Each rowWHTMANI3 As DataRow In dst.Tables("WHTMANI3").Select("")
                    Dim INV_NUM As String = rowWHTMANI3.Item("INV_NUM") & String.Empty
                    Dim PACK_LIST_NO As String = rowWHTMANI3.Item("PACK_LIST_NO") & String.Empty

                    If INV_NUM.Length = 0 OrElse PACK_LIST_NO.Length = 0 Then
                        Continue For
                    End If

                    ' Pre-select for the user
                    For Each rowICTPINV1 As DataRow In dst.Tables("ICTPINV1").Select($"INV_NUM = '{INV_NUM}' and PACK_LIST_NO = '{PACK_LIST_NO}'")
                        rowICTPINV1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                    Next
                Next
            Else
                With rowWHTASNM1
                    .Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("VEND_CODE") = VEND_CODE
                    .Item("PO_SHIPMENT_STATUS") = "O"
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DateTime.Now
                    .Item("SHIP_DATE") = DBNull.Value
                    .Item("ETA_DATE") = DBNull.Value
                    .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                End With
                Fill_Records("ICTPINV1", New String() {"", VEND_CODE, BOL_NO, WHSE_CODE})
            End If

            dst.Tables("WHTASNM1").Rows.Add(rowWHTASNM1)
        Else
            PO_SHIPMENT_NO = HFs("PO_SHIPMENT_NO")
            rowWHTASNM1 = Fill_Record("WHTASNM1", PO_SHIPMENT_NO)

            VEND_CODE = rowWHTASNM1.Item("VEND_CODE") & String.Empty
            WHSE_CODE = rowWHTASNM1.Item("WHSE_CODE") & String.Empty

            Fill_Records("ICTPINV1", New String() {PO_SHIPMENT_NO, VEND_CODE, BOL_NO, WHSE_CODE})
        End If

        If Not (EntryMode = "N" AndAlso processingContainer) Then
            Fill_Records("WHTASNM2", {PO_SHIPMENT_NO})
        End If

        dst.Tables("ICTPINV1_S").Rows.Clear()
        For Each rowICTPINV1 As DataRow In dst.Tables("ICTPINV1").Select("")
            If rowICTPINV1.Item("PO_SHIPMENT_NO") & String.Empty <> String.Empty Then
                dst.Tables("ICTPINV1_S").ImportRow(rowICTPINV1)
                rowICTPINV1.Delete()
            End If
        Next

        dst.Tables("ICTPINV1").AcceptChanges()
        dst.Tables("ICTPINV1_S").AcceptChanges()

        ASCMAIN1.sql = $"SELECT * FROM ASTAUDT1 WHERE TABLE_NAME = 'WHTASNM1' AND KEY_VALUE = '{PO_SHIPMENT_NO}' AND COLUMN_NAME <> 'LAST_DATE'"
        Fill_Records("ASTAUDT1", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdASTAUDT1, "INIT_DATE,COLUMN_NAME")

        Sort_grdColumns(grdICTPINV1_S, "PINV_NO")
        Sort_grdColumns(grdICTPINV1_A, "PINV_NO")

        Select Case rowWHTASNM1.Item("PO_SHIPMENT_STATUS") & String.Empty
            Case "O"
                lblStatus.Text = "Open"
            Case "P"
                lblStatus.Text = "Transmitted"
            Case "C"
                lblStatus.Text = "Comleted"
            Case Else
                lblStatus.Text = "Unknown Status (" & rowWHTASNM1.Item("PO_SHIPMENT_STATUS") & String.Empty & ")"
        End Select

        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()

            INIT_LAST("WHTASNM1")

            ' Merge data back into one table and update that table
            dst.Tables("ICTPINV1").AcceptChanges()
            dst.Tables("ICTPINV1_S").AcceptChanges()
            For Each rowICTPINV1 As DataRow In dst.Tables("ICTPINV1_S").Select("")
                dst.Tables("ICTPINV1").ImportRow(rowICTPINV1)
            Next

            dst.Tables("ICTPINV1").AcceptChanges()
            dst.Tables("ICTPINV1_S").Rows.Clear()

            For Each rowICTPINV1 As DataRow In dst.Tables("ICTPINV1").Select("")
                rowICTPINV1.AcceptChanges()
                rowICTPINV1.SetModified()
            Next

            Update_Record_TDA("WHTASNM1")
            Update_Record_TDA("WHTASNM2", $"PO_SHIPMENT_NO = '{PO_SHIPMENT_NO}'")
            Update_Record_TDA("ICTPINV1")

            ASCMAIN1.sql = $"BOL_NO = '{BOL_NO}' AND CONTAINER_NO = '{CONTAINER_NO}'"

            If processingContainer Then
                Update_Record_TDA("WHTMANI1")
                Update_Record_TDA("WHTMANI2", ASCMAIN1.sql)
                Update_Record_TDA("WHTMANI3", ASCMAIN1.sql)
                Update_Record_TDA("WHTMANI4", ASCMAIN1.sql)
            End If

            'ASCMAIN1.sql = $"BEGIN SOPORDR0_G('{ORDR_GROUP_NO}'); END;"
            'ASCDATA1.ExecuteSQL()

            Transmit_ASN()

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
                sql_where = "LP_CODE IS NOT NULL AND WHSE_STATUS = 'A'"

            Case "VEND_CODE"
                sql_where = "VEND_CODE = 'IPSA'"
        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        InquiryMode = MENU_ITEM_OBJECT = "WHFASNMI"

        If Not InquiryMode Then
            Load_Popup_Menu(grdICTPINV1_A, "SSPBB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Add Seleted to Shipment")
        Else
            Load_Popup_Menu(grdICTPINV1_A, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
        End If

        If Not InquiryMode Then
            Load_Popup_Menu(grdICTPINV1_S, "SSPBB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Remove from Shipment")
        Else
            Load_Popup_Menu(grdICTPINV1_S, "SSPB", "Show Filter", "Show GroupBox", "Auto Size Columns")
        End If

        If Not InquiryMode Then
            Load_Popup_Menu(grdWHTMANI1, "SSPBB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Create Shipment")
        Else
            Load_Popup_Menu(grdWHTMANI1, "SSPBB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Create Shipment")
        End If

        Load_Popup_Menu(grdWHTASNMX, "SSBPB", "Show Filter", "Show GroupBox", "Auto Size Columns", "Transmit ASN")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)

        MyBase.tlb_BeforeToolDropdown(sender, e)

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))

        Select Case e.SourceControl.Name
            Case grdICTPINV1_S.Name
                If tlb_pop.Tools.Exists("Remove from Shipment") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Remove from Shipment"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                            AndAlso rowWHTASNM1.Item("PO_SHIPMENT_STATUS") & String.Empty = "O" _
                            AndAlso "EN".Contains(EntryMode) _
                            AndAlso grd.Selected.Rows.Count > 0
                End If

            Case grdICTPINV1_A.Name
                If tlb_pop.Tools.Exists("Add Seleted to Shipment") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Add Seleted to Shipment"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                        AndAlso rowWHTASNM1.Item("PO_SHIPMENT_STATUS") & String.Empty = "O" _
                        AndAlso "EN".Contains(EntryMode) _
                        AndAlso grd.Selected.Rows.Count > 0
                End If

            Case grdWHTMANI1.Name
                If tlb_pop.Tools.Exists("Create Shipment") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Create Shipment"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Enabled = Not InquiryMode _
                        AndAlso grd.ActiveRow.Band.Key = grd.DisplayLayout.Bands(1).Key _
                        AndAlso grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value & String.Empty = String.Empty
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

            Case "Transmit ASN"
                ' Ask about retransmit XMT_DATE is not null. Check Status
                Dim PO_SHIPMENT_NO As String = grd.ActiveRow.Cells("PO_SHIPMENT_NO").Value & String.Empty
                Dim rowWHTASNM1 As DataRow = LookUp("WHTASNM1", PO_SHIPMENT_NO)
                Dim MT_LEVEL As Int16 = 987

                If rowWHTASNM1 Is Nothing Then
                    MessageBox.Show($"Cannot locate PO Shipment No {PO_SHIPMENT_NO}", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                Select Case rowWHTASNM1.Item("PO_SHIPMENT_STATUS") & String.Empty
                    Case "O" ' Open
                        If Not ASCMAIN1.Logical_Lock("WHTASNM1", PO_SHIPMENT_NO,,,, MT_LEVEL) Then
                            Exit Sub
                        End If

                        If MessageBox.Show($"Do you want to Send PO Shipment No {PO_SHIPMENT_NO} to the 3PL?", e.Tool.Key, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                            ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                            Exit Sub
                        End If

                        Try
                            BeginTrans()
                            ASCMAIN1.sql = "UPDATE WHTASNM1 SET XMIT_IND = '1' WHERE PO_SHIPMENT_NO = :PARM1"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})
                            CommitTrans("PO Shipment Updated")
                            grd.ActiveRow.Cells("XMIT_IND").Value = "1"
                            grd.UpdateData()
                        Catch ex As Exception
                            Rollback(ex.Message)
                        Finally
                            ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                        End Try

                    Case "P" ' Transmitted
                        If Not ASCMAIN1.Logical_Lock("WHTASNM1", PO_SHIPMENT_NO,,,, MT_LEVEL) Then
                            Exit Sub
                        End If

                        If MessageBox.Show($"Do you want to Resend PO Shipment No {PO_SHIPMENT_NO} to the 3PL?", e.Tool.Key, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                            Exit Sub
                        End If

                        Try
                            BeginTrans()
                            ASCMAIN1.sql = "UPDATE WHTASNM1 SET XMIT_IND = '1', XMIT_DATE = NULL WHERE PO_SHIPMENT_NO = :PARM1"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {PO_SHIPMENT_NO})
                            CommitTrans("PO Shipment Updated")
                            grd.ActiveRow.Cells("XMIT_IND").Value = "1"
                            grd.ActiveRow.Cells("XMIT_DATE").Value = DBNull.Value
                            grd.UpdateData()
                        Catch ex As Exception
                            Rollback(ex.Message)
                        Finally
                            ASCMAIN1.MultiTask_Release(,, MT_LEVEL)
                        End Try

                    Case "C" ' Completed
                        MessageBox.Show($"PO Shipment No {PO_SHIPMENT_NO} is complete.", e.Tool.Key, MessageBoxButtons.OK)
                        Exit Sub

                    Case Else
                        MessageBox.Show($"PO Shipment No {PO_SHIPMENT_NO} has an invalid Status", e.Tool.Key, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                End Select


            Case "Add Seleted to Shipment"
                Try
                    If InquiryMode Then
                        MessageBox.Show("Inquiry Mode does not permit Adding to an PO Shipment", "Add Seleted to Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If rowWHTASNM1.Item("PO_SHIPMENT_STATUS") <> "O" Then
                        MessageBox.Show("This PO Shipment does not have a status of Open", "Add Seleted to Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If Not "EN".Contains(EntryMode) Then
                        MessageBox.Show("Form State does not permit Adding to an PO Shipment", "Add Seleted to Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

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
                            rowICTPINV1.Item("PO_SHIPMENT_NO") = PO_SHIPMENT_NO
                            dst.Tables("ICTPINV1_S").ImportRow(rowICTPINV1)
                            rowICTPINV1.Delete()
                        Next
                    Next

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

                    If rowWHTASNM1.Item("PO_SHIPMENT_STATUS") <> "O" Then
                        MessageBox.Show("This PO Shipment does Not have a status Of Open", "Remove from Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    If Not "EN".Contains(EntryMode) Then
                        MessageBox.Show("Form State does Not permit Removing from a PO Shipment", "Remove from Shipment", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Me.Cursor = Cursors.WaitCursor
                    Dim lstPINV_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Rows.GetFilteredInNonGroupByRows
                        If grd.Selected.Rows.Contains(grow) Then
                            lstPINV_NOs.Add(grow.Cells("PINV_NO").Value & String.Empty)
                        End If
                    Next

                    For Each PINV_NO As String In lstPINV_NOs
                        ASCMAIN1.Progress("Processing Invoice(s)", PINV_NO)
                        For Each rowICTPINV1 As DataRow In dst.Tables("ICTPINV1_S").Select($"PINV_NO = '{PINV_NO}'")
                            rowICTPINV1.Item("PO_SHIPMENT_NO") = DBNull.Value
                            dst.Tables("ICTPINV1").ImportRow(rowICTPINV1)
                            rowICTPINV1.Delete()
                        Next
                    Next

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                Finally
                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("", "")
                End Try

            Case "Create Shipment"

                Try
                    BOL_NO = grd.ActiveRow.Cells("BOL_NO").Value & String.Empty
                    CONTAINER_NO = grd.ActiveRow.Cells("CONTAINER_NO").Value & String.Empty

                    If Not ASCMAIN1.Logical_Lock("WHTMANI1", BOL_NO) Then
                        Exit Sub
                    End If

                    If Not ASCMAIN1.Logical_Lock("WHTMANI2", CONTAINER_NO) Then
                        Exit Sub
                    End If

                    Dim rowWHTMANI2 As DataRow = LookUp("WHTMANI2", {BOL_NO, CONTAINER_NO})
                    If rowWHTMANI2 Is Nothing Then
                        MessageBox.Show($"Cannot Locate Container {CONTAINER_NO} on BOL {BOL_NO}.", "Create Shipment", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    ElseIf rowWHTMANI2.Item("PO_SHIPMENT_NO") & String.Empty <> String.Empty Then
                        MessageBox.Show($"Container {CONTAINER_NO} on BOL {BOL_NO} is already assigned an PO Shipment No.", "Create Shipment", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    Else
                        ASCMAIN1.sql = "SELECT INV_NUM, PACK_LIST_NO FROM WHTMANI3 WHERE BOL_NO = :PARM1 AND CONTAINER_NO = :PARM2
                                        MINUS
                                        SELECT INV_NUM, PACK_LIST_NO FROM ICTPINV1"
                        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", {BOL_NO, CONTAINER_NO})

                        If tbl.Rows.Count > 0 Then
                            Dim lst As New List(Of String)
                            For Each row As DataRow In tbl.Select("")
                                lst.Add(row.Item("INV_NUM") & ", " & row.Item("PACK_LIST_NO"))
                            Next

                            Dim uMsg As String = "The following Invoice No, Pack List No(s) cannot be found." & Environment.NewLine & String.Join(Environment.NewLine, lst.ToArray)
                            MessageBox.Show(uMsg, "Create Shipment", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            ASCMAIN1.MultiTask_Release()
                            Exit Sub
                        End If
                    End If

                    processingContainer = True
                    Click_Command("New")
                    processingContainer = True

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    ASCMAIN1.MultiTask_Release()
                End Try

        End Select
    End Sub

#End Region

#Region "Form Procedures"

    Private Sub ImportManifest()
        Try

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
                MessageBox.Show("Import Manifest", zMsg)
                Exit Sub
            End If

            Dim filename As String = (System.IO.Path.GetFileName(importFileName)).ToUpper
            If filename.Length > dst.Tables("WHTMANI1").Columns("FILENAME").MaxLength Then
                filename = filename.Substring(0, dst.Tables("WHTMANI1").Columns("FILENAME").MaxLength).Trim
            End If

            ASCMAIN1.sql = "SELECT * FROM WHTMANI1 WHERE FILENAME = :PARM1"
            Dim tblWHTMANI1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "WHTMANI1", "V", New Object() {filename})
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

            For Each tablename As String In New String() {"WHTMANI1_IMP", "WHTMANI2_IMP", "WHTMANI3_IMP", "WHTMANI4_IMP"}
                dst.Tables(tablename).Rows.Clear()
            Next

            Dim oWB As SpreadsheetGear.IWorkbook
            oWB = SpreadsheetGear.Factory.GetWorkbook(importFileName)
            Dim worksheetIndex As Short = 0
            Dim ws As SpreadsheetGear.IWorksheet = oWB.Worksheets(worksheetIndex)
            Dim startRow As Short = 0
            Dim usedRanged As Int32 = ws.UsedRange.RowCount

            Dim rowWHTMANI1 As DataRow = Nothing

            For excelRow As Int16 = 0 To ws.UsedRange.RowCount + 20

                Select Case excelRow
                    Case 0
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Loading Excel Workbook")
                        rowWHTMANI1 = dst.Tables("WHTMANI1_IMP").NewRow
                        Dim VESSEL_NAME As String = ws.Cells(excelRow, VESSEL).Value & String.Empty
                        If VESSEL_NAME.Contains("=") Then
                            VESSEL_NAME = VESSEL_NAME.Substring(InStr(VESSEL_NAME, "=")).Trim
                        End If

                        If VESSEL_NAME.Length > dst.Tables("WHTMANI1_IMP").Columns("VESSEL_NAME").MaxLength Then
                            VESSEL_NAME = VESSEL_NAME.Substring(0, dst.Tables("WHTMANI1_IMP").Columns("VESSEL_NAME").MaxLength)
                        End If
                        rowWHTMANI1.Item("VESSEL_NAME") = VESSEL_NAME

                        Dim SHIP_DATE As String = ws.Cells(excelRow, ETD).Text & String.Empty
                        If IsDate(SHIP_DATE) Then
                            rowWHTMANI1.Item("SHIP_DATE") = CDate(SHIP_DATE).ToShortDateString
                        End If

                        rowWHTMANI1.Item("INIT_DATE") = DateTime.Now
                        rowWHTMANI1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        rowWHTMANI1.Item("LAST_DATE") = DateTime.Now
                        rowWHTMANI1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                        rowWHTMANI1.Item("FILENAME") = filename
                        rowWHTMANI1.Item("OPS_YYYYPP") = ASCMAIN1.CYP

                    Case 1
                        Dim VOYAGE_NO As String = ws.Cells(excelRow, VOYAGE).Value & String.Empty
                        If VOYAGE_NO.Contains("=") Then
                            VOYAGE_NO = VOYAGE_NO.Substring(InStr(VOYAGE_NO, "=")).Trim
                        End If

                        If VOYAGE_NO.Length > dst.Tables("WHTMANI1_IMP").Columns("VOYAGE_NO").MaxLength Then
                            VOYAGE_NO = VOYAGE_NO.Substring(0, dst.Tables("WHTMANI1_IMP").Columns("VOYAGE_NO").MaxLength)
                        End If
                        rowWHTMANI1.Item("VOYAGE_NO") = VOYAGE_NO

                        Dim ETA_DATE As String = ws.Cells(excelRow, ETA).Text & String.Empty
                        If IsDate(ETA_DATE) Then
                            rowWHTMANI1.Item("ETA_DATE") = CDate(ETA_DATE).ToShortDateString
                        End If

                    Case 2
                        Dim SCAC_CODE As String = ws.Cells(excelRow, SCAC).Value & String.Empty
                        If SCAC_CODE.Contains("=") Then
                            SCAC_CODE = SCAC_CODE.Substring(InStr(SCAC_CODE, "=")).Trim
                        End If

                        If SCAC_CODE.Length > dst.Tables("WHTMANI1_IMP").Columns("SCAC_CODE").MaxLength Then
                            SCAC_CODE = SCAC_CODE.Substring(0, dst.Tables("WHTMANI1_IMP").Columns("SCAC_CODE").MaxLength)
                        End If
                        rowWHTMANI1.Item("SCAC_CODE") = SCAC_CODE

                    Case 3
                        Dim COMPANY_NAME As String = ws.Cells(excelRow, COMPANY).Value & String.Empty
                        If COMPANY_NAME.Contains("=") Then
                            COMPANY_NAME = COMPANY_NAME.Substring(InStr(COMPANY_NAME, "=")).Trim
                        End If

                        If COMPANY_NAME.Length > dst.Tables("WHTMANI1_IMP").Columns("COMPANY_NAME").MaxLength Then
                            COMPANY_NAME = COMPANY_NAME.Substring(0, dst.Tables("WHTMANI1_IMP").Columns("COMPANY_NAME").MaxLength)
                        End If
                        rowWHTMANI1.Item("COMPANY_NAME") = COMPANY_NAME

                        Dim BOL_NO As String = ws.Cells(excelRow, BOL).Value & String.Empty
                        If BOL_NO.Contains("=") Then
                            BOL_NO = BOL_NO.Substring(InStr(BOL_NO, "=")).Trim
                        End If

                        If BOL_NO.Length > dst.Tables("WHTMANI1_IMP").Columns("BOL_NO").MaxLength Then
                            BOL_NO = BOL_NO.Substring(0, dst.Tables("WHTMANI1_IMP").Columns("BOL_NO").MaxLength)
                        End If
                        rowWHTMANI1.Item("BOL_NO") = BOL_NO

                        Dim rowWHTMANI1_LK As DataRow = LookUp("WHTMANI1", BOL_NO)
                        If rowWHTMANI1_LK IsNot Nothing Then
                            MessageBox.Show($"BOL No {BOL_NO} was already imported.", "Import Manifest", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            Exit Sub
                        End If

                        dst.Tables("WHTMANI1_IMP").Rows.Add(rowWHTMANI1)

                    Case Else

                        If (ws.Cells(excelRow, CONTAINER).Value & String.Empty).ToString.ToUpper.Contains("CONTAINER") Then
                            Continue For
                        End If

                        If (ws.Cells(excelRow, CONTAINER).Value & String.Empty).ToString.Trim = String.Empty Then
                            Continue For
                        End If

                        Dim CONTAINER_NO As String = ws.Cells(excelRow, CONTAINER).Value & String.Empty
                        Dim CONTAINER_TYPE As String = ws.Cells(excelRow, CONT_TYPE).Value & String.Empty
                        Dim rowWHTMANI2 As DataRow = dst.Tables("WHTMANI2_IMP").NewRow
                        rowWHTMANI2.Item("BOL_NO") = rowWHTMANI1.Item("BOL_NO")
                        rowWHTMANI2.Item("CONTAINER_NO") = CONTAINER_NO
                        rowWHTMANI2.Item("CONTAINER_TYPE") = CONTAINER_TYPE

                        Dim CONTAINER_LNO As String = ws.Cells(excelRow, CONT_LNO).Value & String.Empty
                        CONTAINER_LNO = CONTAINER_LNO.Trim
                        If BOL_NO.Length = 0 Then
                            CONTAINER_LNO = "Carton " & Val(dst.Tables("WHTMANI2_IMP").Compute("MAX(CONTAINER_LNO)", "") & String.Empty) + 1
                        End If

                        rowWHTMANI2.Item("CONTAINER_LNO") = ws.Cells(excelRow, CONT_LNO).Value & String.Empty

                        Dim NAVALOCK As String = (ws.Cells(excelRow, NAVA).Value & String.Empty).ToString.Trim
                        If NAVALOCK.Length > dst.Tables("WHTMANI2_IMP").Columns("NAVALOCK").MaxLength Then
                            NAVALOCK = NAVALOCK.Substring(0, dst.Tables("WHTMANI1_IMP").Columns("NAVALOCK").MaxLength)
                        End If
                        rowWHTMANI2.Item("NAVALOCK") = NAVALOCK
                        dst.Tables("WHTMANI2_IMP").Rows.Add(rowWHTMANI2)

                        While (ws.Cells(excelRow, PACK_LIST).Value & String.Empty).ToString.Trim.ToUpper <> "TOTAL"
                            Dim INV_NUM As String = (ws.Cells(excelRow, INVOICE_NO).Value & String.Empty).ToString.Trim
                            If INV_NUM.Length > 0 Then
                                Dim rowWHTMANI3 As DataRow = dst.Tables("WHTMANI3_IMP").NewRow
                                rowWHTMANI3.Item("BOL_NO") = rowWHTMANI2.Item("BOL_NO")
                                rowWHTMANI3.Item("CONTAINER_NO") = rowWHTMANI2.Item("CONTAINER_NO")
                                rowWHTMANI3.Item("INV_NUM") = INV_NUM
                                rowWHTMANI3.Item("PACK_LIST_NO") = ws.Cells(excelRow, PACK_LIST).Value & String.Empty
                                rowWHTMANI3.Item("PALLETS") = Val(ws.Cells(excelRow, PALLETS).Value & String.Empty)
                                rowWHTMANI3.Item("CARTONS") = Val(ws.Cells(excelRow, CARTONS).Value & String.Empty)
                                rowWHTMANI3.Item("WEIGHT") = ws.Cells(excelRow, WEIGHT).Value & String.Empty
                                rowWHTMANI3.Item("SHIP_CUBE") = ws.Cells(excelRow, CUBE).Value & String.Empty
                                dst.Tables("WHTMANI3_IMP").Rows.Add(rowWHTMANI3)
                            End If

                            Dim SHIP_SEAL_NO As String = (ws.Cells(excelRow, SEAL).Value & String.Empty).ToString.Trim
                            If SHIP_SEAL_NO.Length > 0 Then
                                Dim rowWHTMANI4 As DataRow = dst.Tables("WHTMANI4_IMP").NewRow
                                rowWHTMANI4.Item("BOL_NO") = rowWHTMANI1.Item("BOL_NO")
                                rowWHTMANI4.Item("CONTAINER_NO") = rowWHTMANI2.Item("CONTAINER_NO")
                                rowWHTMANI4.Item("SHIP_SEAL_NO") = SHIP_SEAL_NO
                                dst.Tables("WHTMANI4_IMP").Rows.Add(rowWHTMANI4)
                            End If

                            excelRow += 1
                        End While
                End Select
            Next

            ws = Nothing
            oWB.Close()
            oWB = Nothing

            ' Get the ship Type from the user
            Using frmASFMSGBF As New ASFMSGBF
                Dim frtOption As Integer = frmASFMSGBF.Get_opt_from_User("Select ship Type", {"Air", "Boat", "Ground"}, 1, "Ship Type")
                Select Case frtOption
                    Case 0
                        rowWHTMANI1.Item("SHIP_TYPE") = "A"
                    Case 1
                        rowWHTMANI1.Item("SHIP_TYPE") = "B"
                    Case 2
                        rowWHTMANI1.Item("SHIP_TYPE") = "G"
                End Select
            End Using

            Try
                BeginTrans()
                Update_Record_TDA("WHTMANI1_IMP")
                Update_Record_TDA("WHTMANI2_IMP")
                Update_Record_TDA("WHTMANI3_IMP")
                Update_Record_TDA("WHTMANI4_IMP")

                Dim BOL_NO As String = dst.Tables("WHTMANI1_IMP").Rows(0).Item("BOL_NO") & String.Empty
                ASCDATA1.ExecuteSQL("UPDATE WHTMANI1 SET WHSE_CODE = 
                                    (
                                        SELECT MAX(POTORDR2.WHSE_CODE)
                                        FROM POTORDR2, ICTPINV1, WHTMANI3
                                        WHERE POTORDR2.PO_ORDER_NO = ICTPINV1.PO_ORDER_NO
                                        AND TO_NUMBER(ICTPINV1.INV_NUM) = TO_NUMBER(WHTMANI3.INV_NUM)
                                        AND TO_NUMBER(ICTPINV1.PACK_LIST_NO) = TO_NUMBER(WHTMANI3.PACK_LIST_NO) 
                                        AND WHTMANI3.BOL_NO = :PARM1
                                    ) WHERE BOL_NO = :PARM1", "V", {BOL_NO})

                ' Do not use this incase either field contains an Alpha character
                'ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS
                '                    SELECT WHTMANI3.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.INV_NUM INV_NUM_P, ICTPINV1.PACK_LIST_NO PACK_LIST_NO_P
                '                    FROM
                '                    (SELECT WHTMANI3.*, TO_NUMBER(WHTMANI3.INV_NUM) INV_NUM_N, TO_NUMBER(PACK_LIST_NO) PACK_LIST_NO_N FROM WHTMANI3) WHTMANI3,
                '                    (SELECT ICTPINV1.*, TO_NUMBER(ICTPINV1.INV_NUM) INV_NUM_N, TO_NUMBER(PACK_LIST_NO) PACK_LIST_NO_N FROM ICTPINV1) ICTPINV1
                '                    WHERE WHTMANI3.INV_NUM_N = ICTPINV1.INV_NUM_N (+)
                '                    AND WHTMANI3.PACK_LIST_NO_N = ICTPINV1.PACK_LIST_NO_N (+)
                '                    AND WHTMANI3.BOL_NO = :PARM1;
                '                    BEGIN FOR R1 IN C1 LOOP
                '                        UPDATE WHTMANI3 SET PACK_LIST_NO = NVL(R1.PACK_LIST_NO_P, PACK_LIST_NO), INV_NUM = NVL(R1.INV_NUM_P, INV_NUM)
                '                        WHERE BOL_NO = R1.BOL_NO AND CONTAINER_NO = R1.CONTAINER_NO AND INV_NUM = R1.INV_NUM;
                '                END LOOP; END; END;"
                'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {BOL_NO})

                ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS
                                    SELECT WHTMANI3.*, ICTPINV1.PO_ORDER_NO, ICTPINV1.INV_NUM INV_NUM_P, ICTPINV1.PACK_LIST_NO PACK_LIST_NO_P
                                    FROM WHTMANI3, ICTPINV1
                                    WHERE ICTPINV1.INV_NUM LIKE '%' || WHTMANI3.INV_NUM 
                                    AND ICTPINV1.PACK_LIST_NO LIKE '%' || WHTMANI3.PACK_LIST_NO
                                    AND NVL(PINV_STATUS, 'O') = 'O'
                                    AND WHTMANI3.BOL_NO = :PARM1;
                                    BEGIN FOR R1 IN C1 LOOP
                                        UPDATE WHTMANI3 SET PACK_LIST_NO = NVL(R1.PACK_LIST_NO_P, PACK_LIST_NO), INV_NUM = NVL(R1.INV_NUM_P, INV_NUM)
                                        WHERE BOL_NO = R1.BOL_NO AND CONTAINER_NO = R1.CONTAINER_NO AND INV_NUM = R1.INV_NUM;
                                END LOOP; END; END;"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {BOL_NO})

                CommitTrans("Import Successful")
            Catch ex As Exception
                Rollback(ex.Message)
            End Try

        Catch ex As Exception
            MessageBox.Show("Import Manifest", ex.Message)
        End Try

    End Sub

    Sub Refresh_Documents()

        Try
            EnforceConstraints(False)
            Me.Cursor = Cursors.WaitCursor

            Select Case optRecords.Value
                Case "O"
                    Fill_Records("WHTASNMX", optRecords.Value)
                    grdWHTASNMX.Text = "Open ASNs"
                Case "P"
                    Fill_Records("WHTASNMX", optRecords.Value)
                    grdWHTASNMX.Text = "Transmitted ASNs"
                Case "Y"
                    ASCMAIN1.sql = "Select * from WHTASNM1 where OPS_YYYYPP = :PARM1"
                    Fill_Records("WHTASNMX", {cbeYP.Value}, True, ASCMAIN1.sql)
                    grdWHTASNMX.Text = $"ASNs created in period {cbeYP.Value}"
                Case Else
                    dst.Tables("WHTASNMX").Rows.Clear()
                    grdWHTASNMX.Text = "No ASNs"
            End Select

            Sort_grdColumns(grdWHTASNMX, "PO_SHIPMENT_NO")
            Clear_All_Filters(grdWHTASNMX)


            dst.Tables("WHTMANI4").Rows.Clear()
            dst.Tables("WHTMANI3").Rows.Clear()
            dst.Tables("WHTMANI2").Rows.Clear()
            dst.Tables("WHTMANI1").Rows.Clear()

            Dim sql = String.Empty

            Select Case optRecords.Value
                Case "O"
                    grdWHTMANI1.Text = "Open Manifests"
                    sql = "SELECT DISTINCT WHTMANI1.BOL_NO 
                            FROM WHTMANI1, WHTASNM1
                            WHERE WHTMANI1.BOL_NO = WHTASNM1.BOL_NO (+)
                            AND WHTASNM1.XMIT_DATE IS NULL"

                Case "P"
                    grdWHTMANI1.Text = "Transmitted Manifests"
                    sql = "SELECT DISTINCT WHTMANI1.BOL_NO 
                            FROM WHTMANI1, WHTASNM1
                            WHERE WHTMANI1.BOL_NO = WHTASNM1.BOL_NO
                            AND WHTASNM1.XMIT_DATE IS NOT NULL"

                Case "Y"
                    grdWHTMANI1.Text = $"Manifests created in period {cbeYP.Value}"
                    sql = $"SELECT DISTINCT WHTMANI1.BOL_NO 
                            FROM WHTMANI1
                            WHERE OPS_YYYYPP = '{cbeYP.Value}'"

                Case Else
                    grdWHTMANI1.Text = "No Manifests"
            End Select

            If sql.Length > 0 Then
                If tblHeader.Length = 0 Then
                    tblHeader = ASCMAIN1.Temp_Table(sql)
                Else
                    ASCDATA1.ExecuteSQL($"TRUNCATE TABLE {tblHeader}")
                    ASCDATA1.ExecuteSQL($"INSERT INTO {tblHeader} {sql}")
                End If

                ASCMAIN1.sql = $"SELECT * FROM WHTMANI1 WHERE BOL_NO IN (SELECT BOL_NO FROM {tblHeader})"
                Fill_Records("WHTMANI1", "", True, ASCMAIN1.sql)

                ASCMAIN1.sql = $"SELECT WHTMANI2.*, WHTASNM1.XMIT_DATE
                                    FROM WHTMANI2, WHTASNM1
                                    WHERE WHTMANI2.PO_SHIPMENT_NO = WHTASNM1.PO_SHIPMENT_NO (+)
                                    AND WHTMANI2.BOL_NO IN (SELECT BOL_NO FROM {tblHeader})"
                Fill_Records("WHTMANI2", "", True, ASCMAIN1.sql)

                ASCMAIN1.sql = $"SELECT WHTMANI3.*, ICTPINV1.PINV_NO, ICTPINV1.PO_ORDER_NO, ICTPINV1.PINV_STATUS
                        FROM WHTMANI3, ICTPINV1
                        WHERE WHTMANI3.BOL_NO IN (SELECT BOL_NO FROM {tblHeader})
                        AND WHTMANI3.INV_NUM = ICTPINV1.INV_NUM (+)
                        AND WHTMANI3.PACK_LIST_NO = ICTPINV1.PACK_LIST_NO (+)
                        AND WHTMANI3.BOL_NO IN (SELECT BOL_NO FROM {tblHeader})"
                Fill_Records("WHTMANI3", "", True, ASCMAIN1.sql)

                ASCMAIN1.sql = $"SELECT * FROM WHTMANI4 WHERE BOL_NO IN (SELECT BOL_NO FROM {tblHeader})"
                Fill_Records("WHTMANI4", "", True, ASCMAIN1.sql)
            End If

            Sort_grdColumns(grdWHTMANI1, "BOL_NO")
            Clear_All_Filters(grdWHTMANI1)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
            EnforceConstraints(True)
        End Try
    End Sub

#End Region

#Region "Form Controls"

    Private Sub grdWHTASNMX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdWHTASNMX.DoubleClickRow
        txtPO_SHIPMENT_NO.Text = e.Row.Cells("PO_SHIPMENT_NO").Value
        Click_Command("View")
    End Sub

    Private Sub grdWHTASNM2_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdWHTASNM2.BeforeRowUpdate
        Dim SHIP_SEAL_NO As String = e.Row.Cells("SHIP_SEAL_NO").Value & String.Empty
        SHIP_SEAL_NO = SHIP_SEAL_NO.Trim

        If SHIP_SEAL_NO.Length = 0 Then
            MessageBox.Show("Seal No is required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        e.Row.Cells("PO_SHIPMENT_NO").Value = PO_SHIPMENT_NO
    End Sub

    Private Sub grdWHTMANI1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdWHTMANI1.InitializeRow

        Select Case e.Row.Band.Key
            Case grdWHTMANI1.DisplayLayout.Bands("WHTMANI1_WHTMANI2").Key
                Dim PO_SHIPMENT_NO As String = e.Row.Cells("PO_SHIPMENT_NO").Value & String.Empty
                If PO_SHIPMENT_NO.Length > 0 Then
                    e.Row.Appearance.BackColor = Drawing.Color.LightGreen
                End If
        End Select
    End Sub

#End Region

#Region "Table Definitions"

    Private Sub ValidateTablesExist()

        For Each tableName As String In {"WHTASNM1", "WHTASNM2", "WHTMANI1", "WHTMANI2", "WHTMANI3", "WHTMANI4"}

            ASCMAIN1.sql = "Select * from USER_TABLES WHERE TABLE_NAME = :PARM1"
            Dim rowTable As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", {tableName})
            If rowTable IsNot Nothing Then
                If tableName <> "ICTPINV1" Then
                    Continue For
                End If
            End If

            Select Case tableName
                Case "WHTASNM1"
                    ASCMAIN1.sql = "CREATE TABLE WHTASNM1(
                                PO_SHIPMENT_NO VARCHAR2(10),
                                VEND_CODE VARCHAR2(10),
                                WHSE_CODE VARCHAR2(6),
                                VESSEL_NAME VARCHAR2(50),
                                CONTAINER_NO VARCHAR2(20),
                                VOYAGE_NO VARCHAR2(25),
                                BOL_NO VARCHAR2(20),
                                SHIP_DATE DATE,
                                ETA_DATE DATE,
                                IDOC_FILENAME VARCHAR2(250),
                                PACK_LIST_NO VARCHAR2(10),
                                INV_NUM VARCHAR2(20),
                                INV_DATE DATE,
                                COUNTRY_CODE VARCHAR2(3),
                                PO_SHIPMENT_STATUS VARCHAR2(1),
                                XMIT_IND VARCHAR2(1),
                                XMIT_DATE DATE,
                                INIT_DATE DATE,
                                INIT_OPER VARCHAR2(20),
                                LAST_DATE DATE,
                                LAST_OPER VARCHAR2(20),
                                OPS_YYYYPP VARCHAR2(6),
                                PORT_CODE VARCHAR2(6),
                                NOTES VARCHAR2(300),
                                SHIP_TYPE VARCHAR2(1),
                                TRANSFER_ORDER VARCHAR2(1),
                                TRANSFER_ORDR_NO VARCHAR2(10),
                                PRIMARY KEY (PO_SHIPMENT_NO))"

                Case "WHTASNM2"
                    ASCMAIN1.sql = "CREATE TABLE WHTASNM2(
                                        PO_SHIPMENT_NO VARCHAR2(10),
                                        SHIP_SEAL_NO VARCHAR2(10),
                                        PRIMARY KEY(PO_SHIPMENT_NO, SHIP_SEAL_NO))"

                Case "WHTMANI1"
                    ASCMAIN1.sql = "CREATE TABLE WHTMANI1(
                                    BOL_NO VARCHAR2(20),
                                    VESSEL_NAME VARCHAR2(50),
                                    VOYAGE_NO VARCHAR2(25),
                                    SCAC_CODE VARCHAR2(20),
                                    COMPANY_NAME VARCHAR2(50),
                                    SHIP_DATE DATE,
                                    ETA_DATE DATE,
                                    WHSE_CODE VARCHAR2(6),
                                    SHIP_TYPE VARCHAR2(1),
                                    FILENAME VARCHAR2(250),
                                    INIT_OPER VARCHAR2(20),
                                    INIT_DATE DATE,
                                    LAST_OPER VARCHAR2(20),
                                    LAST_DATE DATE,
                                    OPS_YYYYPP VARCHAR2(8),
                                    PRIMARY KEY(BOL_NO))"

                Case "WHTMANI2"
                    ASCMAIN1.sql = "CREATE TABLE WHTMANI2(
                                    BOL_NO VARCHAR2(20),
                                    CONTAINER_NO VARCHAR2(20),
                                    CONTAINER_LNO VARCHAR2(20),
                                    NAVALOCK VARCHAR2(20),
                                    CONTAINER_TYPE VARCHAR2(20),
                                    PO_SHIPMENT_NO VARCHAR2(10),
                                    XMIT_DATE DATE,
                                    PRIMARY KEY(BOL_NO, CONTAINER_NO))"

                Case "WHTMANI3"
                    ASCMAIN1.sql = "CREATE TABLE WHTMANI3(
                                    BOL_NO VARCHAR2(20),
                                    CONTAINER_NO VARCHAR2(20),
                                    INV_NUM VARCHAR2(10),
                                    PACK_LIST_NO VARCHAR2(20),
                                    PALLETS NUMBER(6),
                                    CARTONS NUMBER(6),
                                    WEIGHT NUMBER(12, 3),
                                    SHIP_CUBE NUMBER(12,3),
                                    PRIMARY KEY (BOL_NO,CONTAINER_NO,INV_NUM))"

                Case "WHTMANI4"
                    ASCMAIN1.sql = "CREATE TABLE WHTMANI4(
                                    BOL_NO VARCHAR2(20),
                                    CONTAINER_NO VARCHAR2(20),
                                    SHIP_SEAL_NO VARCHAR2(10),
                                    PRIMARY KEY (BOL_NO,CONTAINER_NO,SHIP_SEAL_NO))"

            End Select

            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        Next

        Try
            ASCMAIN1.sql = "Select * from USER_TAB_COLUMNS WHERE TABLE_NAME = :PARM1 AND COLUMN_NAME = :PARM2"
            Dim rowTable As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", {"ICTPINV1", "PO_SHIPMENT_NO"})
            If rowTable Is Nothing Then
                ASCMAIN1.sql = "ALTER TABLE ICTPINV1 ADD PO_SHIPMENT_NO VARCHAR2(10)"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            End If
        Catch ex As Exception

        End Try
    End Sub

#End Region

End Class