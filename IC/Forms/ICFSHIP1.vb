Imports ABSolution

Public Class ICFSHIP1

    Private rowICTSHIP1 As DataRow = Nothing
    Private rowICTITEM1 As DataRow = Nothing
    Private rowARTCUST1 As DataRow = Nothing
    Private rowARTCUST2 As DataRow = Nothing

    Private SHIP_NO As String = String.Empty
    Private ITEM_CODE_last_entry As String = String.Empty
    Private PRICE_CLASS_CODE As String = String.Empty
    Private PRICE_BASIS As String = String.Empty
    Private PRICE_BASE_DPCT As Double = 0
    Private PRICE_LIST_CODE As String = String.Empty

    Dim PRICE_LIST_CODE_ALLO As String
    Dim CUST_CODE_ALLO As String

    Private sqlSOTPICK1 As String = String.Empty
    Private warehouseHasLocatorSystem As Boolean = False

    Private setup_reversal As Boolean = False
    Dim CUST_ADDR_cols As String()

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")

        With dst
            ASCMAIN1.sql = "SELECT * FROM ICTSHIP1 WHERE OPS_YYYYPP = :PARM1"
            MyBase.Create_TDA(.Tables.Add, "ICTSHIPX", ASCMAIN1.sql, 0, False, "V", 0)

            MyBase.Create_TDA(.Tables.Add, "ICTSHIP1", "*")

            ASCMAIN1.sql = "Select ICTSHIP2.*, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.CARTON_PACK_QTY, ICTITEM1.ITEM_UOM" _
                & ", ICTITEM1.ITEM_WEIGHT, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.SALES_DIVISION_CODE, ICTITEM1.ITEM_BIN, ICTITEM1.ITEM_COST_STD" _
                & " from ICTSHIP2, ICTITEM1" _
                & " where ICTITEM1.ITEM_CODE = ICTSHIP2.ITEM_CODE and ICTSHIP2.SHIP_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTSHIP2", "**", 0, True, "V", 2)
            .Tables("ICTSHIP2").Columns.Add("EXT_PRICE", GetType(System.Decimal), "ISNULL(SHIP_QTY, 0) * ISNULL(SHIP_PRICE, 0)")

            MyBase.Create_TDA(.Tables.Add, "SOTORDR1", "*")
            MyBase.Create_TDA(.Tables.Add, "SOTORDR2", "*")
            MyBase.Create_TDA(.Tables.Add, "SOTORDR5", "*")

            ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
              & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_NO_WEB, SOTORDR1.ORDR_SOURCE" & vbCrLf _
              & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
              & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FREIGHT" & vbCrLf _
              & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
              & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND, SOTORDR1.CCPA_NO CCPA_NO_ORDR" & vbCrLf _
              & " from SOTPICK1,SOTORDR1,SOTSHIP1 where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")

            ASCMAIN1.sql = "Select SOTPICK2.*, " & vbCrLf _
                & " SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, " & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTSHIP1" & vbCrLf _
                & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK2", "**")

            With .Tables("SOTPICK2").Columns
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))
            End With

            MyBase.Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            MyBase.Create_TDA(.Tables.Add, "SOTPICK0", "*")

            MyBase.Create_TDA(.Tables.Add, "SOTCART1", "*")
            MyBase.Create_TDA(.Tables.Add, "SOTCART2", "*")

            MyBase.Create_TDA(.Tables.Add, "SOTINVH1", "*")
            MyBase.Create_TDA(.Tables.Add, "SOTINVH2", "*")

            MyBase.Create_TDA(.Tables.Add, "ARTOPEN1", "*")
        End With

        grdICTSHIPX.DataSource = dst.Tables("ICTSHIPX")
        Create_Summary(grdICTSHIPX, "SHIP_NO", "Count")

        grdICTSHIP2.DataSource = dst.Tables("ICTSHIP2")
        Create_Summary(grdICTSHIP2, "SHIP_LNO", "Count")
        Create_Summary(grdICTSHIP2, "SHIP_QTY", "Sum")
        Create_Summary(grdICTSHIP2, "EXT_PRICE", "Sum")

        MyBase.Absx1.dteFor("SHIP_DATE").MaxDate = DateAdd(DateInterval.Month, 2, DateTime.Now)
        MyBase.Absx1.dteFor("SHIP_DATE").MinDate = CDate("01/01/2013")

        grdICTSHIPX.Parent = tab.Parent
        splDetails.Parent = tab.Parent
        tab.Visible = False

        Bind_Controls(UltraGroupBox1, "ICTSHIP1")
        Bind_Controls(grpDetails, "ICTSHIP1")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)


        CUST_ADDR_cols = {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3",
            "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY",
            "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "New"

                Validate_Code("WHSE_CODE")
                Validate_Code("CUST_CODE")
                Validate_Code("SHIP_VIA_CODE")

                rowARTCUST2 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_STORE_NO = :PARM2", _
                                                 "VV", New Object() {MyBase.Absx1.txtFor("CUST_CODE").Text, MyBase.Absx1.txtFor("CUST_STORE_NO").Text})

                If rowARTCUST2 Is Nothing Then
                    EMsg &= vbCr & "Invalid Customer/Store No combination"
                End If

            Case "View", "Edit"
                SHIP_NO = MyBase.Absx1.txtFor("SHIP_NO").Text.Trim
                SHIP_NO = ASCMAIN1.Format_Field(SHIP_NO, "SHIP_NO")

                rowICTSHIP1 = ASCDATA1.GetDataRow("SELECT * FROM ICTSHIP1 WHERE SHIP_NO = :PARM1", "V", New Object() {SHIP_NO})
                If rowICTSHIP1 Is Nothing Then
                    EMsg &= vbCr & "Missing or invalid Shipment No"
                ElseIf rowICTSHIP1.Item("SHIP_STATUS") & String.Empty <> "O" AndAlso eItemKey = "Edit" Then
                    EMsg &= vbCr & "You may edit only Open Inventory Shipments"
                End If

            Case "Done"

            Case "Reverse"

                Dim SHIP_NO As String = rowICTSHIP1.Item("SHIP_NO")
                If MsgBox("Do you want to set up an entry to Reverse Shipment " & SHIP_NO, vbYesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"
                If dst.Tables("ICTSHIP2").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
                    EMsg &= vbCr & "You must have at least one detail line"
                End If

                MyBase.Absx1.txtFor("SHIP_CUST_PO").Text = MyBase.Absx1.txtFor("SHIP_CUST_PO").Text.Trim
                If MyBase.Absx1.txtFor("SHIP_CUST_PO").TextLength = 0 Then
                    EMsg &= vbCr & "Customer PO is required"
                End If

                If Not MyBase.Absx1.dteFor("SHIP_DATE").IsDateValid Then
                    EMsg &= vbCr & "Transaction date is missing or invalid"
                End If

                Absx1.txtFor("STAX_CODE").Text = Absx1.txtFor("STAX_CODE").Text.Trim
                Absx1.txtFor("MISC_CHG_CODE").Text = Absx1.txtFor("MISC_CHG_CODE").Text.Trim

                If Absx1.txtFor("STAX_CODE").TextLength = 0 AndAlso Val(Absx1.numFor("SHIP_STAX").Value & "") <> 0 Then
                    EMsg &= vbCr & "You Must Specify a Tax Code"
                ElseIf Val(Absx1.numFor("SHIP_STAX").Value & "") <> 0 Then
                    Validate_Code("STAX_CODE")
                End If

                If Absx1.txtFor("MISC_CHG_CODE").TextLength = 0 AndAlso Val(Absx1.numFor("SHIP_MISC_CHG").Value & "") <> 0 Then
                    EMsg &= vbCr & "You Must Specify a Misc Chg Code"
                ElseIf Val(Absx1.numFor("SHIP_MISC_CHG").Value & "") <> 0 Then
                    Validate_Code("MISC_CHG_CODE")
                End If

                ' All items must have a bin or the Whse must have a value in WHSE_LOC_SHIP
                If warehouseHasLocatorSystem Then
                    If dst.Tables("ICTSHIP2").Select("ISNULL(ITEM_BIN,'') = '' or ITEM_BIN = ''", "", DataViewRowState.CurrentRows).Length > 0 Then
                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", MyBase.Absx1.txtFor("WHSE_CODE").Text)
                        If rowICTWHSE1 Is Nothing OrElse rowICTWHSE1.Item("WHSE_LOC_SHP") & String.Empty = String.Empty Then
                            EMsg &= vbCr & "There is at least one item without a bin assignment and the warehouse is not setup with a defualt Shipping Location"
                        End If
                    End If
                End If


                Dim DT As Date = Absx1.dteFor("SHIP_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                End If


            Case "Cancel"
                Dim qMsg As String = String.Empty
                If EntryMode = "N" Then
                    qMsg = "Do you want to cancel entry of a new shipment."
                Else
                    qMsg = "Do you want to cancel changes made to this shipment?"
                End If

                If MessageBox.Show(qMsg, "Cancel Changes", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)


            Case "Reverse"

                Dim SHIP_NO As String = rowICTSHIP1.Item("SHIP_NO")
                Dim SHIP_VIA_CODE As String = rowICTSHIP1.Item("SHIP_VIA_CODE")
                Dim CUST_STORE_NO As String = rowICTSHIP1.Item("CUST_STORE_NO")
                Dim CUST_CODE As String = rowICTSHIP1.Item("CUST_CODE")
                Dim WHSE_CODE As String = rowICTSHIP1.Item("WHSE_CODE")
                Dim INV_NO As String = rowICTSHIP1.Item("INV_NO")

                Mode_Settings(False)

                Absx1.txtFor("CUST_CODE").Text = CUST_CODE
                Absx1.txtFor("CUST_STORE_NO").Text = CUST_STORE_NO
                Absx1.txtFor("SHIP_VIA_CODE").Text = SHIP_VIA_CODE
                Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE

                setup_reversal = True

                Click_Command("New")
                 
                Dim rowICTSHIP1_to_reverse As DataRow = LookUp("ICTSHIP1", New String() {SHIP_NO})
                 
                For Each C As String In New String() _
                    {"SHIP_CUST_PO", "MISC_CHG_CODE", "STAX_CODE"}
                    Absx1.txtFor(C).Text = rowICTSHIP1_to_reverse.Item(C) & ""

                Next
                Absx1.numFor("SHIP_MISC_CHG").Value = -1 * Val(rowICTSHIP1_to_reverse.Item("SHIP_MISC_CHG"))
                Absx1.numFor("SHIP_FREIGHT").Value = -1 * Val(rowICTSHIP1_to_reverse.Item("SHIP_FREIGHT"))
                Absx1.numFor("SHIP_STAX").Value = -1 * Val(rowICTSHIP1_to_reverse.Item("SHIP_STAX"))

                Absx1.dteFor("SHIP_DATE").Value = rowICTSHIP1_to_reverse.Item("SHIP_DATE")

                Fill_Records("ICTSHIP2", SHIP_NO)
                For Each row As DataRow In dst.Tables("ICTSHIP2").Select()
                    row.Item("SHIP_NO") = Me.SHIP_NO
                    row.Item("SHIP_QTY") = -1 * Val(row.Item("SHIP_QTY"))
                    row.AcceptChanges()
                    row.SetAdded()
                Next

                CalcTotal()

                setup_reversal = False

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Update"
                CreateFinalizationData()
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print Invoice"
                Print_Invoice()

            Case "Load From P.O."
                LoadFromPurchaseOrder()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowICTSHIP1.Item("SHIP_STATUS") & "" = "O" Then
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
                .Items("Load From P.O.").Settings.Enabled = iScreenMode
                 
                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode
                .Items("Print Invoice").Visible = EntryMode = "V"

                .Items("Reverse").Visible = ScreenMode And (EntryMode = "V") And Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Load From P.O.").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode And Not setup_reversal
            End With
            .Groups("Show if Entered in").Visible = Not ScreenMode
        End With

        If (EntryMode = "N" OrElse EntryMode = "E") AndAlso Not InquiryMode AndAlso Not setup_reversal Then
            grdICTSHIP2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdICTSHIP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdICTSHIP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            Set_Read_Only(grpDetails, False)
        Else
            grdICTSHIP2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdICTSHIP2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdICTSHIP2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            Set_Read_Only(grpDetails, True)
        End If

        lblStatus.Visible = ScreenMode
        splDetails.Visible = ScreenMode
        grdICTSHIPX.Visible = Not ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Private Sub Clear_Record()
        MyBase.EnforceConstraints(False)

        For Each tablename As String In New String() {"ICTSHIP1", "ICTSHIP2", "SOTORDR1", "SOTORDR2", "SOTORDR5", _
                                                      "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2", _
                                                      "SOTINVH1", "SOTINVH2", "ARTOPEN1"}
            dst.Tables(tablename).Rows.Clear()
        Next

        Refresh_Documents()

        rowICTSHIP1 = Nothing
        rowARTCUST1 = Nothing
        rowARTCUST2 = Nothing
        SHIP_NO = String.Empty
        warehouseHasLocatorSystem = False

        MyBase.Absx1.txtFor("WHSE_CODE").Clear()
        MyBase.Absx1.txtFor("CUST_CODE").Clear()
        MyBase.Absx1.txtFor("CUST_STORE_NO").Clear()
        MyBase.Absx1.txtFor("SHIP_CUST_PO").Clear()
        MyBase.Absx1.txtFor("SHIP_NO").Clear()
        MyBase.Absx1.dteFor("SHIP_DATE").DateTime = DateTime.Now

        MyBase.EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        Save_Header_Fields(UltraGroupBox1)
        MyBase.EnforceConstraints(False)

        If EntryMode = "N" Then
            rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {HFs("CUST_CODE")})
            rowARTCUST2 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_STORE_NO = :PARM2", "VV", New Object() {HFs("CUST_CODE"), HFs("CUST_STORE_NO")})

            SHIP_NO = ASCMAIN1.Next_Control_No("ICTSHIP1.SHIP_NO")
            rowICTSHIP1 = dst.Tables("ICTSHIP1").NewRow
            rowICTSHIP1.Item("SHIP_NO") = SHIP_NO
            rowICTSHIP1.Item("SHIP_DATE") = (MyBase.Absx1.dteFor("SHIP_DATE").DateTime).ToString("MM/dd/yyyy")
            rowICTSHIP1.Item("CUST_CODE") = rowARTCUST1.Item("CUST_CODE")
            rowICTSHIP1.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
            rowICTSHIP1.Item("CUST_STORE_NO") = rowARTCUST2.Item("CUST_STORE_NO")
            rowICTSHIP1.Item("CUST_STORE_NAME") = rowARTCUST2.Item("CUST_STORE_NAME")
            rowICTSHIP1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            'rowICTSHIP1.Item("SHIP_CUST_PO") = HFs("SHIP_CUST_PO")
            rowICTSHIP1.Item("SHIP_VIA_CODE") = HFs("SHIP_VIA_CODE")
            rowICTSHIP1.Item("SHIP_STATUS") = "O"
            rowICTSHIP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTSHIP1.Item("INIT_DATE") = DateTime.Now
            rowICTSHIP1.Item("SHIP_FREIGHT") = 0
            rowICTSHIP1.Item("SHIP_MISC_CHG") = 0
            rowICTSHIP1.Item("SHIP_STAX") = 0
            dst.Tables("ICTSHIP1").Rows.Add(rowICTSHIP1)
        Else
            SHIP_NO = HFs("SHIP_NO")
            Fill_Records("ICTSHIP1", SHIP_NO)
            rowICTSHIP1 = dst.Tables("ICTSHIP1").Rows(0)

            rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {rowICTSHIP1.Item("CUST_CODE")})
            rowARTCUST2 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_STORE_NO = :PARM2", "VV", New Object() {rowICTSHIP1.Item("CUST_CODE"), rowICTSHIP1.Item("CUST_STORE_NO")})
        End If

        warehouseHasLocatorSystem = ASCDATA1.GetDataValue("SELECT WHSE_LOCATOR FROM ICTWHSE1 WHERE WHSE_CODE = '" & HFs("WHSE_CODE") & "'") & String.Empty = "1"

        grdICTSHIP2.DisplayLayout.Bands(0).Columns("ITEM_BIN").Hidden = Not warehouseHasLocatorSystem

        PRICE_CLASS_CODE = rowARTCUST1.Item("PRICE_CLASS_CODE") & String.Empty
        PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & String.Empty

        Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
        PRICE_BASIS = rowSOTPCLS1.Item("PRICE_BASIS") & ""
        PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")

        CUST_CODE_ALLO = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
        PRICE_LIST_CODE_ALLO = ""
        If CUST_CODE_ALLO <> "" Then
            Dim rowARTCUST1_ALLO As DataRow = LookUp("ARTCUST1", CUST_CODE_ALLO)
            If rowARTCUST1_ALLO IsNot Nothing Then
                PRICE_LIST_CODE_ALLO = rowARTCUST1_ALLO.Item("PRICE_LIST_CODE") & ""
            End If
        End If

        Fill_Records("ICTSHIP2", SHIP_NO)

        Select Case rowICTSHIP1.Item("SHIP_STATUS") & String.Empty
            Case "O"
                lblStatus.Text = "Open"
            Case "D"
                lblStatus.Text = "Deleted"
            Case "F"
                lblStatus.Text = "Finalized"
            Case Else
                lblStatus.Text = "Unknown Status (" & rowICTSHIP1.Item("SHIP_STATUS") & String.Empty & ")"
        End Select

        MyBase.EnforceConstraints(True)
        CalcTotal()

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Dim ORDR_GROUP_NO As String = Create_Order_Pick_Ship_Carton()

        Try
            MyBase.BeginTrans()

            ' Curently we are not creating these records in the database. I need to create them locally for the Invoice to be created.
            '"SOTORDR1", "SOTORDR2", "SOTORDR5",
            '"SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2",

            INIT_LAST("ICTSHIP1")

            For Each tablename As String In New String() {"ICTSHIP1", "ICTSHIP2"}
                Update_Record_TDA(tablename)
            Next
            For Each tablename As String In New String() {"SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2", "SOTSHIP1", "SOTPICK0"}
                Update_Record_TDA(tablename)
            Next

            ' Cleanup ditry data used to create keys in work tables
            'For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows
            '    rowSOTINVH1.Item("ORDR_NO") = String.Empty ' "0".PadLeft(10, "0")
            '    rowSOTINVH1.Item("PICK_NO") = String.Empty
            '    rowSOTINVH1.Item("SHIP_BOL_NO") = String.Empty
            'Next

            For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Rows
                Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows.Find(New String() {"I", rowICTSHIP1.Item("INV_NO")})
                Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
                rowARTOPEN1.Item("ORDR_NO") = ORDR_NO ' String.Empty  '"0".PadLeft(10, "0")
            Next

            For Each tablename As String In New String() {"SOTINVH1", "SOTINVH2", "ARTOPEN1"}
                Update_Record_TDA(tablename)
            Next

            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows

                ASCMAIN1.sql = "BEGIN SOPSTAT1('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "BEGIN ICPSHIP2('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()

                If warehouseHasLocatorSystem Then TAC.ICCMAIN1.Update_WHTLOCBX("S", rowSOTINVH1.Item("INV_NO"))
            Next

            ASCMAIN1.sql = $"BEGIN SOPORDR0_G('{ORDR_GROUP_NO}'); END;"
            ASCDATA1.ExecuteSQL()

            MyBase.CommitTrans("Update successful")

        Catch ex As Exception
            MyBase.Rollback(ex.Message)
        End Try
    End Sub

#End Region


    Function Create_Order_Pick_Ship_Carton() As String

        Dim CUST_CODE As String = rowICTSHIP1.Item("CUST_CODE")
        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
        Dim CUST_STORE_NO As String = rowICTSHIP1.Item("CUST_STORE_NO")
        Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

        Dim WHSE_CODE As String = rowICTSHIP1.Item("WHSE_CODE")

        Dim INV_NO As String = rowICTSHIP1.Item("INV_NO")
        Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows.Find(New String() {"I", INV_NO})
        Dim PICK_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTPICK0.PICK_BATCH_NO")
        Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
        Dim PICK_NO As String = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")

        Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
        Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
        rowSOTINVH1.Item("ORDR_NO") = ORDR_NO
        rowSOTINVH1.Item("PICK_NO") = PICK_NO
        rowSOTINVH1.Item("SHIP_BOL_NO") = SHIP_BOL_NO

        Dim ORDR_GROUP_NO As String = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")

        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow
        With rowSOTORDR1
            .Item("ORDR_NO") = ORDR_NO
            .Item("ORDR_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_STORE_NO") = CUST_STORE_NO
            .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
            .Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION")
            Dim CUST_BILL_TO_CUST As String = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
            If CUST_BILL_TO_CUST = "" Then
                CUST_BILL_TO_CUST = CUST_CODE
            End If
            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST


            .Item("CUST_DISC_PCT") = 0 ' NOT RIGHT

            .Item("PRICE_CLASS_CODE") = rowARTCUST1.Item("PRICE_CLASS_CODE")
            .Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE")
            .Item("ORDR_CUST_PO") = rowICTSHIP1.Item("SHIP_CUST_PO")
            .Item("ORDR_SHIP_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_CANCEL_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_SOURCE") = "K"
            .Item("ORDR_DATE_RECD") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_SHIP_TO") = "MK"
            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
            .Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS")
            .Item("POST_CODE") = rowARTCUST1.Item("POST_CODE")
            .Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
            .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
            .Item("WHSE_CODE") = rowICTSHIP1.Item("WHSE_CODE")
            .Item("BRAND_CODE") = ""
            .Item("EVENT_CODE") = ""
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("ORDR_REL_BATCH_NO") = ""
            .Item("ORDR_DATE_REL") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_STATUS") = "F"
            .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
            .Item("ORDR_ROTR") = "0"
            .Item("ORDR_ROTR_CODE") = ""
            .Item("CURR_CODE") = "USD"
            .Item("CURR_EXCH_RATE") = 1
            .Item("COLLECTION_CODE") = ""

            .Item("ORDR_PICK_SEQ") = 1
            .Item("ORDR_OVERRIDE_NOT_ALLOCATED") = "0"
            .Item("SALES_DIVISION_CODE") = ""
            .Item("ORDR_HOLD") = "0"
            .Item("ORDR_ADDR_TYPE_ST") = "MA"

            .Item("SHIP_VIA_CODE") = rowICTSHIP1.Item("SHIP_VIA_CODE")
            .Item("ORDR_TYPE_CODE") = "REG"

            .Item("ORDR_ORIG_SHIP_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_ORIG_CANCEL_DATE") = rowICTSHIP1.Item("SHIP_DATE")

            .Item("ORDR_DATE_BOOKED") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_MISC_CHG") = rowICTSHIP1.Item("SHIP_MISC_CHG")
            .Item("ORDR_STAX") = rowICTSHIP1.Item("SHIP_STAX")
            .Item("STAX_CODE") = rowICTSHIP1.Item("STAX_CODE")
            '.Item("MISC_CHG_CODE") = rowICTSHIP1.Item("MISC_CHG_CODE")
            .Item("ORDR_DATE_CLOSED") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_YYYYPP_CLOSED") = ASCMAIN1.CYP
            .Item("ORDR_HIGH_PRIORITY") = "0"
            .Item("PRICE_LIST_CODE") = rowARTCUST1.Item("PRICE_LIST_CODE")
            .Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE")
            .Item("ORDR_DATE_SHIPPED") = rowICTSHIP1.Item("SHIP_DATE")
        End With
        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

        Create_SOTORDR5(ORDR_NO, "BT", rowARTCUST1)
        Create_SOTORDR5(ORDR_NO, "ST", rowARTCUST2)

        Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").NewRow
        With rowSOTPICK0
            .Item("PICK_BATCH_NO") = PICK_BATCH_NO
            .Item("PICK_SHPS") = 1
            .Item("PICK_CTNS") = 1
            .Item("PICK_PKTS") = 1
            .Item("PICK_BATCH_STATUS") = "F"
            .Item("PICK_PKTS") = 1
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("PICK_SHIP_REL_DATE") = DATETIME_STAMP.Date
            .Item("WHSE_CODE") = WHSE_CODE
        End With
        dst.Tables("SOTPICK0").Rows.Add(rowSOTPICK0)

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
        With rowSOTSHIP1
            .Item("SHIP_BOL_NO") = SHIP_BOL_NO
            .Item("SHIP_DATE_SHIPPED") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("SHIP_VIA_CODE") = rowICTSHIP1.Item("SHIP_VIA_CODE")
            .Item("SHIP_REF") = ""
            .Item("SHIP_TOTAL_WGT") = 1
            .Item("SHIP_CNT_CARTONS") = 1
            .Item("SHIP_ADDR_TYPE") = "MA"
            .Item("SHIP_ADDR_CODE") = rowICTSHIP1.Item("CUST_STORE_NO")
            .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            .Item("SHIP_PICK_PRINTED") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("PICK_BATCH_NO") = PICK_BATCH_NO
            .Item("SHIP_STATUS") = "F"
            .Item("FRT_TERMS") = "PPD"
            .Item("WHSE_CODE") = rowICTSHIP1.Item("WHSE_CODE")
            .Item("INV_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
            .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
            .Item("ORDR_DEPT") = ""
            .Item("SHIPPED_ACTUAL") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_PICK_TYPE") = "P"
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
        End With
        dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)


        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        With rowSOTCART1
            .Item("CART_NO") = CART_NO
            .Item("PICK_NO") = PICK_NO
            .Item("CART_TOTAL_UNITS") = 0
            .Item("CART_TOTAL_WGT_CALC") = 0
        End With
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

        Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
        With rowSOTPICK1
            .Item("PICK_NO") = PICK_NO
            .Item("ORDR_NO") = ORDR_NO
            .Item("ORDR_PICK_SEQ") = 1
            .Item("PICK_STATUS") = "F"
            .Item("PICK_RELEASED") = DATETIME_STAMP
            .Item("PICK_PRINTED") = DATETIME_STAMP
            .Item("PICK_BATCH_NO") = PICK_BATCH_NO
            .Item("SHIP_BOL_NO") = SHIP_BOL_NO
            .Item("INV_NO") = INV_NO
            .Item("PICK_CNT_CARTONS") = 1
            .Item("PICK_TOTAL_WGT") = 1
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("PICK_PRINTED_OPER") = ASCMAIN1.USER_ID

            .Item("WHSE_CODE") = WHSE_CODE
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_STORE_NO") = CUST_STORE_NO
        End With
        dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

        For Each rowICTSHIP2 As DataRow In dst.Tables("ICTSHIP2").Select("", "SHIP_LNO")
            Dim ITEM_CODE As String = rowICTSHIP2.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            Dim SHIP_LNO As Integer = Val(rowICTSHIP2.Item("SHIP_LNO") & "")
            Dim SHIP_QTY As Integer = Val(rowICTSHIP2.Item("SHIP_QTY") & "")
            Dim SHIP_PRICE As Decimal = Val(rowICTSHIP2.Item("SHIP_PRICE") & "")

            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
            With rowSOTORDR2
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = SHIP_LNO
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                .Item("ORDR_UNIT_PRICE") = SHIP_PRICE
                .Item("ORDR_QTY") = SHIP_QTY
                .Item("ORDR_QTY_SHIP") = SHIP_QTY
                .Item("ORDR_QTY_ORIG") = SHIP_QTY
                .Item("ORDR_STATUS") = "F"

                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = CUST_STORE_NO
                .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
                .Item("WHSE_CODE") = rowICTSHIP1.Item("WHSE_CODE")
                .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                .Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                .Item("ORDR_UNIT_PRICE_CURR") = SHIP_PRICE
                .Item("ITEM_RETAIL_PRICE_CURR") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")

                '.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                .Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE")
                '.Item("ALLO_CTL_NO_REL") = ALLO_CTL_NO_REL
            End With
            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

            Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
            With rowSOTPICK2
                .Item("PICK_NO") = PICK_NO
                .Item("PICK_LNO") = SHIP_LNO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = SHIP_LNO
                .Item("PICK_QTY") = SHIP_QTY
                .Item("PICK_QTY_CONF") = SHIP_QTY
                .Item("PICK_UNIT_PRICE") = SHIP_PRICE
            End With
            dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)

            Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
            With rowSOTCART2
                .Item("CART_NO") = CART_NO
                .Item("CART_LNO") = SHIP_LNO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = SHIP_LNO
                .Item("QTY_PACKED") = SHIP_QTY
                .Item("ITEM_UPC_CODE") = rowICTITEM1.Item("ITEM_UPC_CODE")
                .Item("ITEM_EAN_CODE") = rowICTITEM1.Item("ITEM_EAN_CODE")
                .Item("ITEM_CODE") = ITEM_CODE
            End With
            dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
        Next

        Return ORDR_GROUP_NO
    End Function

    Sub Create_SOTORDR5(ORDR_NO As String, CUST_ADDR_TYPE As String, row As DataRow)
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
        With rowSOTORDR5
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
            For Each C As String In CUST_ADDR_cols
                Dim CC As String = C
                If CUST_ADDR_TYPE = "ST" Then
                    CC = Replace(CC, "CUST_", "CUST_STORE_")
                End If
                .Item(C) = row.Item(CC)
            Next
        End With
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
    End Sub

#Region "grdICTSHIP2"

    Private Sub grdICTSHIP2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSHIP2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "")
                If ITEM_CODE <> "" Then
                    e.Cell.Row.Cells("ITEM_UOM").Value = rowICTITEM1.Item("ITEM_UOM") & ""
                    e.Cell.Row.Cells("ITEM_BIN").Value = rowICTITEM1.Item("ITEM_BIN") & ""
                    e.Cell.Row.Cells("ITEM_COST_STD").Value = Val(rowICTITEM1.Item("ITEM_COST_STD") & "")
                    e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = Val(rowICTITEM1.Item("SALES_DIVISION_CODE") & "")
                    e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                    e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1.Item("ITEM_RETAIL_PRICE")

                    Dim SHIP_PRICE As Decimal = TAC.SOCMAIN1.Get_Price _
                                                     (Me, _
                                                      PRICE_LIST_CODE,
                                                      PRICE_LIST_CODE_ALLO, _
                                                      PRICE_BASIS, _
                                                      PRICE_BASE_DPCT, _
                                                      ITEM_CODE, _
                                                      rowICTITEM1, _
                                                      rowICTSHIP1.Item("SHIP_DATE"), 0)

                    e.Cell.Row.Cells("SHIP_PRICE").Value = SHIP_PRICE

                    e.Cell.Row.Cells("ITEM_SO_QTY_MULT").Value = rowICTITEM1.Item("ITEM_SO_QTY_MULT")
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = rowICTITEM1.Item("CARTON_PACK_QTY")
                End If


        End Select
    End Sub

    Private Sub grdICTSHIP2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTSHIP2.AfterRowActivate

        If Trim(grdICTSHIP2.ActiveRow.Cells("ITEM_CODE").Value & "") = "" And _
            (grdICTSHIP2.ActiveCell Is Nothing OrElse _
             grdICTSHIP2.ActiveCell.Column.Key <> "ITEM_CODE") _
        Then
            grdICTSHIP2.ActiveCell = grdICTSHIP2.ActiveRow.Cells("ITEM_CODE")
        End If

        If grdICTSHIP2.ActiveRow.IsAddRow Then
            With grdICTSHIP2.DisplayLayout.Bands(0)
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ITEM_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("SHIP_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
             End With
        Else
            With grdICTSHIP2.DisplayLayout.Bands(0)
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ITEM_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("SHIP_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            End With
        End If
    End Sub

    Private Sub grdICTSHIP2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdICTSHIP2.AfterRowsDeleted
        CalcTotal()
    End Sub

    Private Sub grdICTSHIP2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTSHIP2.AfterRowUpdate
        CalcTotal()
    End Sub

    Private Sub grdICTSHIP2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdICTSHIP2.BeforeCellUpdate

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim ITEM_CODE As String = Validate_Item(e.NewValue & "")
                If ITEM_CODE = "" Then
                    e.Cancel = True
                End If
        End Select

    End Sub

    Private Sub grdICTSHIP2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTSHIP2.BeforeExitEditMode
        If grdICTSHIP2.ActiveCell IsNot Nothing Then
            With grdICTSHIP2.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                    Case "SHIP_QTY"
                        ' rem out the next 8 lines
                        If .EditorResolved.Value & "" = "" _
                        Or Val(.EditorResolved.Value & "") < 0 _
                        Then
                            .EditorResolved.Value = 0
                        End If
                        If Val(.EditorResolved.Value & "") < 0 Then
                            .EditorResolved.Value = System.Math.Abs(Val(.EditorResolved.Value & ""))
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdICTSHIP2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTSHIP2.BeforeRowUpdate

        Validate_Columns("ITEM_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("SHIP_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        ITEM_CODE_last_entry = e.Row.Cells("ITEM_CODE").Value & ""

        If e.Row.IsAddRow Then
            e.Row.Cells("SHIP_NO").Value = SHIP_NO
            Dim ORDR_LNO As Int64 = Val(dst.Tables("ICTSHIP2").Compute("MAX(SHIP_LNO)", "") & "") + 1
            e.Row.Cells("SHIP_LNO").Value = ORDR_LNO
        End If
    End Sub

    Private Sub grdICTSHIP2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTSHIP2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "X"
                    If Val(.Cells("ORDR_QTY_CANC").Value & "") <> 0 Then
                        If MsgBox("Restore Cancelled Qty of " & .Cells("ORDR_QTY_CANC").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        .Cells("ORDR_QTY_OPEN").Value = Val(.Cells("ORDR_QTY_OPEN").Value & "") + Val(.Cells("ORDR_QTY_CANC").Value & "")
                        ' grdSOWORDR2_AfterColUpdate(.Cells("ORDR_QTY_OPEN").position)
                        .Update()
                    Else
                        If MsgBox("Cancel Remaining Qty Open of " & .Cells("ORDR_QTY_OPEN").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        .Cells("ORDR_QTY_OPEN").Value = "0"
                        ' grdSOWORDR2_AfterColUpdate(.Cells("ORDR_QTY_OPEN").position)
                        grdICTSHIP2.ActiveRow.Update()
                    End If

                Case "ITEM_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdICTSHIP2, sql_where)


            End Select
        End With

    End Sub

    Private Sub grdICTSHIP2_Error(sender As Object, e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTSHIP2.Error
        e.Cancel = True
    End Sub

    Private Sub grdICTSHIP2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTSHIP2.InitializeRow
        Dim ITEM_SO_QTY_MULT As Integer = Val(e.Row.Cells("ITEM_SO_QTY_MULT").Value & "")
        Dim SHIP_QTY As Integer = Val(e.Row.Cells("SHIP_QTY").Value & "")
        If ITEM_SO_QTY_MULT <> 0 AndAlso SHIP_QTY Mod ITEM_SO_QTY_MULT <> 0 Then
            e.Row.Cells("SHIP_QTY").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("SHIP_QTY").ToolTipText = "Order Qty not Divisible by Inner Pack Qty"
        Else
            e.Row.Cells("SHIP_QTY").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("SHIP_QTY").ToolTipText = ""
        End If

        If Val(e.Row.Cells("ITEM_COST_STD").Value & String.Empty) > Val(e.Row.Cells("SHIP_PRICE").Value & String.Empty) Then
            e.Row.Cells("ITEM_COST_STD").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("ITEM_COST_STD").ToolTipText = "Standard Cost greater than Price"
        Else
            e.Row.Cells("ITEM_COST_STD").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("ITEM_COST_STD").ToolTipText = ""
        End If
    End Sub

    Private Sub grdICTSHIP2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles grdICTSHIP2.KeyDown
        With grdICTSHIP2
            Try
                If e.KeyCode = Keys.F5 Then
                    Select Case .ActiveCell.Column.Key
                        Case "ITEM_CODE"
                            grdICTSHIP2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                            .ActiveCell.Value = ITEM_CODE_last_entry
                            .ActiveCell.SelStart = Len(grdICTSHIP2.ActiveCell.Text)
                    End Select
                End If
            Catch ex As Exception

            End Try
        End With
    End Sub

    Private Sub grdICTSHIPX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTSHIPX.DoubleClickRow
        If grdICTSHIPX.ActiveRow Is Nothing Then
            Exit Sub
        End If

        MyBase.Absx1.txtFor("SHIP_NO").Text = grdICTSHIPX.ActiveRow.Cells("SHIP_NO").Text
        Click_Command("View")

    End Sub

#End Region

#Region "Form Procedures"

    Private Function Validate_Item(ITEM_CODE_z As String) As String
        Dim E As String = ""

        Dim ITEM_CODE As String = ""
        rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE_z)

        If rowICTITEM1 Is Nothing Then
            E = "Item is Not on File" & vbCrLf
        Else
            If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
                E = "Item Status is not Active" & vbCrLf
            End If
            If rowICTITEM1.Item("ITEM_UOM") & "" = "" Then
                E = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTITEM1.Item("SALES_DIVISION_CODE") & "" = "" Then
                E = "Item does not have a valid Division Code" & vbCrLf
            End If
        End If

        If E <> "" And grdICTSHIP2.ActiveRow IsNot Nothing AndAlso grdICTSHIP2.ActiveRow.IsAddRow Then
            MessageBox.Show(E, "Invalid Item", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            If E = "" Then
                ITEM_CODE = rowICTITEM1.Item(0)
            End If
        End If
        Return ITEM_CODE
    End Function

    Private Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdICTSHIP2.ActiveRow
            Select Case COLUMN_NAME

                Case "ITEM_CODE"
                    If .Cells("ITEM_CODE").Text <> "" Then
                        Dim ITEM_CODE As String = Validate_Item(.Cells("ITEM_CODE").Value & "")
                        Cancel = (ITEM_CODE = "")
                    End If

                Case "SHIP_QTY"
                    If Trim(.Cells("ITEM_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If

                    If Trim(.Cells("SHIP_QTY").Value & "") = "" Then
                        MsgBox("Ship Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdICTSHIP2.ActiveCell = grdICTSHIP2.ActiveRow.Cells("SHIP_QTY")
                        Exit Sub
                    End If

                    ' rem out the next 4 lines
                    If Val(.Cells("SHIP_QTY").Value & "") < 0 Then
                        MsgBox("Ship Qty May Not be Negative", vbOKOnly, "Invalid Ship Quantity")
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Private Sub CreateFinalizationData()

        ' Create a temp SOTORDR1,2 SOTPICK1,2 SOTSHIP1 record for this Shipment so and invoice can be created.

        rowICTSHIP1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        rowICTSHIP1.Item("SHIP_STATUS") = "F"

        'SOTORDR1
        dst.Tables("SOTORDR1").Rows.Clear()
        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow

        Dim CUST_BILL_TO_CUST As String = rowARTCUST1.Item("CUST_BILL_TO_CUST") & String.Empty
        If CUST_BILL_TO_CUST = "" Then
            CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_CODE") & String.Empty
        End If
        With rowSOTORDR1
            .Item("ORDR_NO") = "S" & rowICTSHIP1.Item("SHIP_NO").ToString.Substring(1)
            .Item("ORDR_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("CUST_CODE") = rowICTSHIP1.Item("CUST_CODE")
            .Item("CUST_NAME") = rowICTSHIP1.Item("CUST_NAME")
            .Item("CUST_STORE_NO") = rowICTSHIP1.Item("CUST_STORE_NO")
            .Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION")
            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            .Item("CUST_DISC_PCT") = rowARTCUST1.Item("CUST_DISC_PCT")
            .Item("PRICE_CLASS_CODE") = rowARTCUST1.Item("PRICE_CLASS_CODE")
            .Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE")
            .Item("ORDR_CUST_PO") = rowICTSHIP1.Item("SHIP_CUST_PO")
            .Item("ORDR_SOURCE") = "K"
            .Item("ORDR_SHIP_TO") = "MK"
            .Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS")
            .Item("POST_CODE") = rowARTCUST1.Item("POST_CODE")
            .Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
            .Item("SREP_CODE") = IIf((rowARTCUST2.Item("SREP_CODE") & String.Empty).ToString.Trim.Length > 0, rowARTCUST2.Item("SREP_CODE"), rowARTCUST1.Item("SREP_CODE"))
            .Item("WHSE_CODE") = rowICTSHIP1.Item("WHSE_CODE")
            .Item("ORDR_STATUS") = "F"
            .Item("ORDR_GROUP_NO") = "O" & rowICTSHIP1.Item("SHIP_NO").ToString.Substring(1)
            .Item("CURR_CODE") = rowARTCUST1.Item("CURR_CODE")
            .Item("CURR_EXCH_RATE") = "1"
            .Item("ORDR_FREIGHT") = Val(MyBase.Absx1.numFor("SHIP_FREIGHT").Value & String.Empty)
            .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE")
            .Item("STAX_RATE") = 0
            .Item("STAX_CODE") = MyBase.Absx1.txtFor("STAX_CODE").Text
            .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
            .Item("SHIP_VIA_CODE") = rowICTSHIP1.Item("SHIP_VIA_CODE")
            .Item("ORDR_TYPE_CODE") = "REG"
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DateTime.Now
            .Item("LAST_DATE") = DateTime.Now
            .Item("ORDR_SHIP_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
            .Item("ORDR_MISC_CHG") = Val(MyBase.Absx1.numFor("SHIP_MISC_CHG").Value & String.Empty)
            .Item("ORDR_STAX") = Val(MyBase.Absx1.numFor("SHIP_STAX").Value & String.Empty)
            .Item("PRICE_LIST_CODE") = PRICE_LIST_CODE
            .Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE") & String.Empty
        End With
        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

        Dim ITEM_CODEs As String = String.Empty
        For Each rowSOTORDR2 As DataRow In dst.Tables("ICTSHIP2").Select("")
            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & ""
            ITEM_CODEs &= ",'" & ITEM_CODE & "'"
        Next

        ASCMAIN1.sql = "Select MAX(SALES_DIVISION_CODE) SALES_DIVISION_CODE from ICTITEM1 where ITEM_CODE IN (" & Mid$(ITEM_CODEs, 2) & ")"
        Dim rowICTITEM1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
        rowSOTORDR1.Item("SALES_DIVISION_CODE") = rowICTITEM1.Item("SALES_DIVISION_CODE")

        ' SOTPICK1
        dst.Tables("SOTPICK1").Clear()
        Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
        With rowSOTPICK1
            .Item("PICK_NO") = "P" & rowICTSHIP1.Item("SHIP_NO").ToString.Substring(1)
            .Item("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO")
            .Item("PICK_FREIGHT") = 0
            .Item("ORDR_PICK_SEQ") = 1
            .Item("PICK_STATUS") = "F"
            .Item("PICK_RELEASED") = DateTime.Now
            .Item("PICK_PRINTED") = DateTime.Now
            .Item("PICK_SHIPPED") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("SHIP_BOL_NO") = "B" & rowICTSHIP1.Item("SHIP_NO").ToString.Substring(1)

            .Item("INV_NO") = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            rowICTSHIP1.Item("INV_NO") = rowSOTPICK1.Item("INV_NO")

            .Item("PICK_CNT_CARTONS") = 1
            .Item("PICK_TOTAL_WGT") = 1
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DateTime.Now
            .Item("LAST_DATE") = DateTime.Now

            ' Items from SOTORDR1 stuffed into SOTPICK1 - The invoice class is expecting them
            .Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
            .Item("CUST_STORE_NO") = rowSOTORDR1.Item("CUST_STORE_NO")
            .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
            .Item("ORDR_NO_WEB") = rowSOTORDR1.Item("ORDR_NO_WEB")
            .Item("ORDR_SOURCE") = rowSOTORDR1.Item("ORDR_SOURCE")
            .Item("SALES_DIVISION_CODE") = rowSOTORDR1.Item("SALES_DIVISION_CODE")
            .Item("CUST_BILL_TO_CUST") = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
            .Item("POST_CODE") = rowSOTORDR1.Item("POST_CODE")
            .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
            .Item("ORDR_FREIGHT") = rowSOTORDR1.Item("ORDR_FREIGHT")
            .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
            .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
            .Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE")
            .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
            .Item("ORDR_INV_COMMENT") = rowSOTORDR1.Item("ORDR_INV_COMMENT")
            .Item("CUST_FACTOR_IND") = rowSOTORDR1.Item("CUST_FACTOR_IND")
            .Item("CCPA_NO_ORDR") = rowSOTORDR1.Item("CCPA_NO")
        End With
        dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

        ' SOTORDR2 / SOTPICK2
        Dim ORDR_LNO As Int16 = 0
        dst.Tables("SOTORDR2").Clear()
        dst.Tables("SOTPICK2").Clear()

        For Each rowICTSHIP2 As DataRow In dst.Tables("ICTSHIP2").Select("", "", DataViewRowState.CurrentRows)
            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
            With rowSOTORDR2
                .Item("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO")
                ORDR_LNO += 1
                .Item("ORDR_LNO") = ORDR_LNO
                .Item("ITEM_CODE") = rowICTSHIP2.Item("ITEM_CODE")
                .Item("ITEM_DESC") = rowICTSHIP2.Item("ITEM_DESC")
                .Item("ORDR_UNIT_PRICE") = rowICTSHIP2.Item("SHIP_PRICE")
                .Item("ORDR_QTY") = rowICTSHIP2.Item("SHIP_QTY")
                .Item("ORDR_QTY_OPEN") = 0
                .Item("ORDR_QTY_PICK") = 0
                .Item("ORDR_QTY_SHIP") = rowICTSHIP2.Item("SHIP_QTY")
                .Item("ORDR_QTY_CANC") = 0
                .Item("ORDR_QTY_ORIG") = rowICTSHIP2.Item("SHIP_QTY")
                .Item("ORDR_STATUS") = "F"
                .Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE")
                .Item("CUST_STORE_NO") = rowSOTORDR1.Item("CUST_STORE_NO")
                .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
                .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                .Item("ITEM_RETAIL_PRICE") = rowICTSHIP2.Item("ITEM_RETAIL_PRICE")
                .Item("ORDR_UNIT_PRICE_CURR") = rowICTSHIP2.Item("SHIP_PRICE")
                .Item("ITEM_RETAIL_PRICE_CURR") = rowICTSHIP2.Item("ITEM_RETAIL_PRICE")
            End With
            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

            Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
            With rowSOTPICK2
                .Item("PICK_NO") = rowSOTPICK1.Item("PICK_NO")
                .Item("PICK_LNO") = ORDR_LNO
                .Item("ORDR_NO") = rowSOTORDR2.Item("ORDR_NO")
                .Item("ORDR_LNO") = rowSOTORDR2.Item("ORDR_LNO")
                .Item("PICK_QTY") = rowSOTORDR2.Item("ORDR_QTY_SHIP")
                .Item("PICK_QTY_CONF") = rowSOTORDR2.Item("ORDR_QTY_SHIP")
                .Item("PICK_QTY_CANC") = 0
                .Item("PICK_QTY_BACK") = 0
                .Item("PICK_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                .Item("PICK_QTY_CANC_REL") = 0
                .Item("PICK_QTY_BACK_REL") = 0
                .Item("PICK_856_TD5_IND") = String.Empty

                ' Additional fields used in Invoice Creation class
                .Item("ITEM_CODE") = rowSOTORDR2.Item("ITEM_CODE")
                .Item("ITEM_DESC") = rowSOTORDR2.Item("ITEM_DESC")
                .Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                .Item("PICK_AMT") = rowSOTORDR2.Item("ORDR_QTY")
                .Item("PICK_AMT_CONF") = rowSOTORDR2.Item("ORDR_QTY_SHIP")
                .Item("PICK_AMT_CANC") = 0
                .Item("PICK_AMT_BACK") = 0
            End With
            dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
        Next

        dst.Tables("SOTORDR5").Rows.Clear()
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
        With rowSOTORDR5
            .Item("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO")
            .Item("CUST_ADDR_TYPE") = "BT"
            .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
            .Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1")
            .Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2")
            .Item("CUST_ADDR3") = rowARTCUST1.Item("CUST_ADDR3")
            .Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY")
            .Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE")
            .Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE")
            .Item("CUST_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY")
            .Item("CUST_CONTACT") = rowARTCUST1.Item("CUST_CONTACT")
            .Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE")
            .Item("CUST_EXT") = rowARTCUST1.Item("CUST_EXT")
            .Item("CUST_FAX") = rowARTCUST1.Item("CUST_FAX")
            .Item("CUST_EMAIL") = rowARTCUST1.Item("CUST_EMAIL")
            .Item("CUST_ADDR_CODE") = rowSOTORDR1.Item("CUST_CODE")
        End With
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

        rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
        With rowSOTORDR5
            .Item("ORDR_NO") = rowSOTORDR1.Item("ORDR_NO")
            .Item("CUST_ADDR_TYPE") = "ST"
            .Item("CUST_NAME") = rowARTCUST2.Item("CUST_STORE_NAME")
            .Item("CUST_ADDR1") = rowARTCUST2.Item("CUST_STORE_ADDR1")
            .Item("CUST_ADDR2") = rowARTCUST2.Item("CUST_STORE_ADDR2")
            .Item("CUST_ADDR3") = rowARTCUST2.Item("CUST_STORE_ADDR3")
            .Item("CUST_CITY") = rowARTCUST2.Item("CUST_STORE_CITY")
            .Item("CUST_STATE") = rowARTCUST2.Item("CUST_STORE_STATE")
            .Item("CUST_ZIP_CODE") = rowARTCUST2.Item("CUST_STORE_ZIP_CODE")
            .Item("CUST_COUNTRY") = rowARTCUST2.Item("CUST_STORE_COUNTRY")
            .Item("CUST_CONTACT") = rowARTCUST2.Item("CUST_STORE_CONTACT")
            .Item("CUST_PHONE") = rowARTCUST2.Item("CUST_STORE_PHONE")
            .Item("CUST_EXT") = rowARTCUST2.Item("CUST_STORE_EXT")
            .Item("CUST_FAX") = rowARTCUST2.Item("CUST_STORE_FAX")
            .Item("CUST_EMAIL") = rowARTCUST2.Item("CUST_STORE_EMAIL")
            .Item("CUST_ADDR_CODE") = rowSOTORDR1.Item("CUST_STORE_NO")
        End With
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)


        dst.Tables("SOTSHIP1").Rows.Clear()
        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
        With rowSOTSHIP1
            .Item("SHIP_BOL_NO") = rowSOTPICK1.Item("SHIP_BOL_NO")
            .Item("SHIP_DATE_SHIPPED") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("SHIP_VIA_CODE") = rowSOTORDR1.Item("SHIP_VIA_CODE")
            .Item("SHIP_TOTAL_WGT") = 1
            .Item("SHIP_CNT_CARTONS") = 1
            .Item("SHIP_ADDR_TYPE") = "MK"
            .Item("SHIP_ADDR_CODE") = rowSOTORDR1.Item("CUST_STORE_NO")
            .Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO")
            .Item("SHIP_STATUS") = "F"
            .Item("FRT_TERMS") = rowSOTORDR1.Item("FRT_TERMS")
            .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
            .Item("INV_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("INIT_DATE") = DateTime.Now
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DateTime.Now
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
            .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
            .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("SHIPPED_ACTUAL") = rowICTSHIP1.Item("SHIP_DATE")
            .Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE")
        End With
        dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)

        ' Create Invoice Records
        dst.Tables("SOTINVH1").Rows.Clear()
        dst.Tables("SOTINVH2").Rows.Clear()
        dst.Tables("ARTOPEN1").Rows.Clear()

        For Each rowSOTSHIP1 In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)

            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            rowSOTSHIP1.Item("INV_DATE") = rowICTSHIP1.Item("SHIP_DATE")
            rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = rowICTSHIP1.Item("SHIP_DATE")

            Dim SOCINVH1 As New TAC.SOCINVH1(dst.Tables("SOTINVH1"), dst.Tables("SOTINVH2"), _
                                              dst.Tables("SOTPICK1"), dst.Tables("SOTPICK2"), _
                                              dst.Tables("ARTOPEN1"), dst.Tables("SOTSHIP1"), _
                                              dst.Tables("SOTORDR5"), dst.Tables("SOTORDR1"), dst.Tables("SOTORDR2"))
            SOCINVH1.CreateInvoices(SHIP_BOL_NO, rowICTSHIP1.Item("CUST_CODE") & String.Empty)

            If Val(dst.Tables("SOTINVH1").Rows(0).Item("INV_MISC_CHG") & String.Empty) <> 0 Then
                dst.Tables("SOTINVH1").Rows(0).Item("MISC_CHG_CODE") = MyBase.Absx1.txtFor("MISC_CHG_CODE").Text
            End If
        Next

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty
            rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

            If rowSOTORDR1 IsNot Nothing Then
                rowSOTORDR1.Item("ORDR_DATE_CLOSED") = rowSOTINVH1.Item("INV_DATE_SHIPPED")
                rowSOTORDR1.Item("ORDR_YYYYPP_CLOSED") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")
            End If
        Next

    End Sub

    Private Sub Print_Invoice()
        Try
            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Now Preparing Invoice for Printing")

            Dim REPORT_NAME As String = "SORINVP1"
            Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
            If RPT = "" Then RPT = REPORT_NAME

            If Not REPORTS.ContainsKey(REPORT_NAME) Then
                REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
                REPORTS(REPORT_NAME).Prepare_dst(False, "")
            End If

            Dim sql As String = " and SOTINVH1.INV_TYPE = 'I' and SOTINVH1.INV_NO = '" & rowICTSHIP1.Item("INV_NO") & "'"
            Dim tempFileName As String = "Inv" & DateTime.Now.ToString("yyyyMMddHHmmss")

            REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
            Dim FILENAME As String = ""
            With REPORTS(REPORT_NAME).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "")
                Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
                FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
                .Print_Report_End(, True)
            End With

            Show_Document(FILENAME)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print Credit Memo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Try
            Me.Cursor = Cursors.WaitCursor
            Fill_Records("ICTSHIPX", cbeYP.Value)
            grdICTSHIPX.Text = "Entered in " & cbeYP.Text
            Sort_grdColumns(grdICTSHIPX, "SHIP_NO")
            Clear_All_Filters(grdICTSHIPX)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_FREIGHT", "SHIP_STAX", "SHIP_MISC_CHG"
                CalcTotal()
        End Select
    End Sub

    Private Sub CalcTotal()
        Absx1.numFor("SHIP_TOTAL").Value = _
            Val(Absx1.numFor("SHIP_FREIGHT").Value & "") + _
            Val(Absx1.numFor("SHIP_STAX").Value & "") + _
            Val(Absx1.numFor("SHIP_MISC_CHG").Value & "") + _
            Val(dst.Tables("ICTSHIP2").Compute("SUM(EXT_PRICE)", "") & String.Empty)
    End Sub

    Private Sub LoadFromPurchaseOrder()

        Try

            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("PO_ORDER_NO")
            If ASCMAIN1.CodeSelector.SQL <> "" Then
                ASCMAIN1.CodeSelector.MultipleSelections = False
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                Using F As New ASFCODE1
                    F.ShowDialog()
                End Using
                If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                    Dim PO_ORDER_NO As String = ASCMAIN1.CodeSelector.SelectedCode
                    Dim rowPOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM POTORDR1 WHERE PO_ORDER_NO = :PARM1", "V", New Object() {PO_ORDER_NO})
                    If rowPOTORDR1 Is Nothing Then
                        MessageBox.Show("Could not locate selected Purchase Order", "Load from P.O.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If

                    If rowPOTORDR1.Item("WHSE_CODE") & String.Empty <> MyBase.Absx1.txtFor("WHSE_CODE").Text Then
                        If MessageBox.Show("The selected P.O. is is for Warehouse " & rowPOTORDR1.Item("WHSE_CODE") & "." & Environment.NewLine _
                                            & "Do you want to continue?", "Load From P.O.", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    Dim tblPOTORDR2 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM POTORDR2 WHERE PO_ORDER_NO = :PARM1", "POTORDR2", "V", New Object() {PO_ORDER_NO})

                    Dim NUM_PO_LNO As Int16 = tblPOTORDR2.Rows.Count
                    Dim PO_QTY_REC As Int16 = tblPOTORDR2.Select("PO_QTY_REC > 0").Length
                    Dim PO_QTY_OPN As Int16 = tblPOTORDR2.Select("PO_QTY_REC = 0").Length

                    If PO_QTY_REC = 0 Then
                        MessageBox.Show("There are no items received on this P.O. Inport aborted!", "Load From P.O.", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        Exit Sub
                    End If

                    If PO_QTY_OPN > 0 Then
                        If MessageBox.Show("There are " & PO_QTY_OPN & " line items not received. These items will be skipped." & Environment.NewLine _
                            & "Continue with import?", "Load From P.O.", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    Me.Cursor = Cursors.WaitCursor
                    For Each rowPOTORDR2 As DataRow In tblPOTORDR2.Select("PO_QTY_REC > 0", "PO_ORDER_LNO")
                        grdICTSHIP2.DisplayLayout.Bands(0).AddNew()
                        With grdICTSHIP2.ActiveRow
                            .Cells("ITEM_CODE").Value = rowPOTORDR2.Item("ITEM_CODE")
                            .Cells("SHIP_QTY").Value = rowPOTORDR2.Item("PO_QTY_REC")
                            .Update()
                        End With
                    Next

                    Sort_grdColumns(grdICTSHIP2, "SHIP_LNO")

                    Me.Cursor = Cursors.Default
                    MessageBox.Show("Import Complete", "Import From P.O.", MessageBoxButtons.OK)
                End If
            End If
        Catch ex As Exception

        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

#End Region

End Class