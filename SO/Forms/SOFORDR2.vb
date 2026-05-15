Public Class SOFORDR2
    'Load_Events
    'Stop ' Use Copy-From Order Values for SREP_CODE, SREP2_CODE, SHIP_VIA_CODE
    'Stop ' ORDR_ADDR_TYPE_ST LOGIC NEEDED
    'Stop ' LOAD UP COMBO THAT OFFERS SELECTION IF BILLTO ADDRESSES
    ' PROBLEM WHERE QTY ORDR DID NOT MAKE IT INTO QTY OPEN FOR 1 ITEM on a ms order entry

    ' load hold codes grid with the following:
    'ORDR_REL_HOLD_CODES
    'ORDR_CRED_HOLD_CODES
    'ORDR_HOLD_INVTY
    'ORDR_HOLD_SALES

#Region "Declarations"
    Dim CUST_CODE As String
    Dim WHSE_CODE As String
    Dim REV_NO As Int32
    Dim ORDR_NO As String
    Dim ORDR_NO_to_copy As String
    Dim ORDR_GROUP_NO As String     ' ORDR_GROUP_NO for Order currently in process
    Dim ORDR_GROUP_NOs As New List(Of String)    ' ORDR_GROUP_NOs for Order(s) currently in process

    Dim EDI_JRNL_NO As String       ' EDI_JRNL_NO for Order currently in process
    Dim ORDR_CUST_PO As String      ' Customer's PO No
    Dim CUST_STORE_NO As String     ' Store Related to this Order
    Dim SREP_CODE As String         ' Orders Sales Rep Code
    Dim SREP2_CODE As String        ' Orders Sales Rep2 Code

    Dim rowSOTORDR1 As DataRow
    Dim rowSOTORDR0 As DataRow
    Dim rowARTCUST1 As DataRow      ' ARTCUST1 for the Sold-To
    Dim rowARTCUST1_BT As DataRow   ' ARTCUST1 for the Bill-To
    Dim rowARTCUST2 As DataRow      ' ARTCUST2 for Store
    Dim dvwSOTORDR5 As DataView
    Dim sqlSOTORDRB As String = ""

    Dim rowSOTORDR1po As DataRow
    Dim rowICTITEM1 As DataRow

    Dim PRICE_BASE_DPCT As Decimal
    Dim PRICE_CLASS_CODE As String
    Dim PRICE_BASIS As String
    Dim PRICE_LIST_CODE As String

    Dim REV_LNO As Integer = 0

    Dim disable_update As Boolean   ' flag indicating that the Update Button should be disabled
    Dim TOTAL_ORDR_AMT As Decimal ' Total Order Amount from EDI Header

    Dim ORDR_LNOs As New List(Of Int64) ' list of ORDR_LNOs that are deleted

    'Dim MASS_CHANGE_TEXT As String  ' Value to Propagate in Multiple Order Maintenance
    'Dim MASS_CHANGE_VALUE As String ' Value of Cloned Field Prior to Change

    'Dim do_not_update               ' to stop grdSOWORDR2_LostFocus from Updating

    Dim CUST_NAME As String         ' Sold-To Customer Name
    Dim CUST_BILL_TO_CUST As String ' Bill-To Customer Code
    Dim CUST_DC_NO As String        ' DC Related to the CUST_STORE_NO of this Order

    Dim restore_reservation As Boolean
    Dim sub_grid As String

    Dim multi_store_is_active As Boolean            ' Multi-Store Order mode
    Dim multi_store_changes_made_to_SOTORDRS As Boolean           ' Changes were made in grdSOTORDRS and need to get to grdSOWORDR2 without recursively going back to grdSOWORDRS
    Dim CUST_STORE_NOs_multi_store As New List(Of String)
    Dim ORDR_NOs_to_maintain As New List(Of String)

    Dim multiple_order_maintenance As Boolean = False
    Dim multiple_order_type As String = ""

    Dim msqty As Int64               ' Repeat MS qty when entering stores blindly
    Dim msqty_col As Int64        ' ORDR_LNO for MS qty
    Dim RANGE_TYPE As String
    Dim ALLOW_CHANGE_RANGE As String
    Dim ITEM_CODE_last_entry As String = ""

    Private clsShip As New TAC.WHCSHIP1
    Private shipPackageDetailList As New List(Of nsoftware.InShip.PackageDetail)

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        AUDIT.Add("ARTCUST2", "E")

        With dst

            ASCMAIN1.sql = "Select SOTORDR1.* from SOTORDR1 where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)

            ASCMAIN1.sql = "Select SOTORDR2.*, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.CARTON_PACK_QTY, ICTITEM1.ITEM_UOM, ICTITEM1.ITEM_WEIGHT from SOTORDR2,ICTITEM1" _
                & " where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE and SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, True, "V", 2)
            With .Tables("SOTORDR2").Columns
                .Add("ORDR_RETAIL_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ITEM_RETAIL_PRICE,0)")
                .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_PICK", GetType(System.Decimal), "ISNULL(ORDR_QTY_PICK,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_SHIP", GetType(System.Decimal), "ISNULL(ORDR_QTY_SHIP,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            End With


            Create_TDA(.Tables.Add, "SOTORDR4", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)

            With .Tables.Add("SOTORDRT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("KEY")}
            End With

            ASCMAIN1.sql = "Select SOTPICK1.*, SOTSHIP1.SHIP_DATE_SHIPPED, SOTINVH1.INV_FREIGHT, SOTINVH1.INV_TOTAL_AMOUNT" & vbCrLf _
                & " from SOTPICK1,SOTSHIP1,SOTINVH1 " & vbCrLf _
                & " where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTINVH1.INV_TYPE (+) = 'I'" & vbCrLf _
                & "   and SOTINVH1.INV_NO (+) = SOTPICK1.INV_NO" & vbCrLf _
                & "   and SOTPICK1.PICK_STATUS <> 'D'" & vbCrLf _
                & "   and SOTPICK1.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*" & vbCrLf _
                & " from SOTPICK2,SOTPICK1 " & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO and SOTPICK1.ORDR_NO = :PARM1" _
                & "   and SOTPICK1.PICK_STATUS <> 'D'"
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "V", 2)


            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            Create_Relation("SOTORDR2", "SOTPICK2", "ORDR_NO,ORDR_LNO")
            With .Tables("SOTPICK2").Columns
                .Add("ITEM_CODE", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).ITEM_CODE")
                .Add("ITEM_DESC", GetType(System.String), "PARENT(SOTORDR2_SOTPICK2).ITEM_DESC")
            End With

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from SOTCART1 " & vbCrLf _
                & " where SOTCART1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCART2.*" & vbCrLf _
                & " from SOTCART2,SOTCART1 " & vbCrLf _
                & " where SOTCART2.CART_NO = SOTCART1.CART_NO" _
                & "   and SOTCART1.PICK_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0, False, "V", 2)

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")

            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
                & ", SOTORDR2.ORDR_LNO MLNO, SOTORDR1.CUST_STORE_NO " & vbCrLf _
                & ", SOTORDR2.ITEM_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC, '1' SUB" & vbCrLf _
                & ", SOTORDR2.ITEM_CODE ITEM_CODE_ORIG, SOTORDR2.ORDR_QTY_OPEN ORDR_QTY_OPEN_ORIG" & vbCrLf _
                & ", SOTORDR2.ORDR_LNO ORDR_LNO_ORIG" & vbCrLf _
                & " from SOTORDR1,SOTORDR2 " & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR1.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRC", "**", 0, False, "V", 2)
            .Tables("SOTORDRC").Columns.Add("ORDR_QTY_CANC_CALC", GetType(System.Int64), "ISNULL(ORDR_QTY,0)-ISNULL(ORDR_QTY_PICK,0)-ISNULL(ORDR_QTY_OPEN,0)-ISNULL(ORDR_QTY_SHIP,0)")
            .Tables("SOTORDRC").Columns.Add("ORDR_QTY_CANC_CALC2", GetType(System.Int64), "IIF(ORDR_QTY_CANC_CALC<0,0,ORDR_QTY_CANC_CALC)")

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_LNO MLNO" & vbCrLf _
                & ", SOTORDR2.ITEM_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
                & ", 0 ORDR_CXL, 0 ORDR_SUB " & vbCrLf _
                & " from SOTORDR2 " & vbCrLf _
                & " where SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRD", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO" & vbCrLf _
                & ", SOTORDR2.ITEM_CODE" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTORDR2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                & " from SOTORDR2 " & vbCrLf _
                & " where SOTORDR2.ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRQ", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_RETAIL_PRICE, SOTORDR2.ORDR_UNIT_PRICE" _
                & ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_CANC, SOTORDR2.ORDR_QTY_ORIG" _
                & ", ICTITEM1.ITEM_UOM, ICTITEM1.ITEM_DESC, ICTITEM1.CARTON_PACK_QTY" _
                & " from SOTORDR2,ICTITEM1 where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTORDRI", "**", 0, False, "", 1)
            With .Tables("SOTORDRI")
                With .Columns
                    .Add("COL", GetType(System.Int32))
                    '.Add("ORDR_RETAIL_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ITEM_RETAIL_PRICE,0)")
                    .Add("ORDR_RETAIL_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ITEM_RETAIL_PRICE,0)")
                    .Add("ORDR_AMT", GetType(System.Decimal), "ISNULL(ORDR_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                    .Add("ORDR_AMT_OPEN", GetType(System.Decimal), "ISNULL(ORDR_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                    .Add("ORDR_AMT_CANC", GetType(System.Decimal), "ISNULL(ORDR_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")
                End With
                .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE")}
            End With

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO" & vbCrLf _
              & ", SOTORDR1.CUST_STORE_NO" & vbCrLf _
              & ", SOTORDR1.ORDR_SHIP_DATE" & vbCrLf _
              & ", SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
              & ", SOTORDR1.ORDR_CUST_PO" & vbCrLf _
              & ", SOTORDR1.ORDR_DEPT" & vbCrLf _
              & ", SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
              & ", SOTORDR1.ORDR_HOLD" & vbCrLf _
              & ", SOTORDR1.ORDR_OVERRIDE_NOT_ALLOCATED" & vbCrLf _
              & ", SOTORDR1.ORDR_STATUS" & vbCrLf _
              & ", SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
              & ", ARTCUST2.CUST_STORE_NAME" & vbCrLf _
              & ", ARTCUST2.CUST_STORE_LOCATION" & vbCrLf _
              & " from SOTORDR1,ARTCUST2 " & vbCrLf _
              & " where SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
              & "   and ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
              & "   and ARTCUST2.CUST_STORE_NO = SOTORDR1.CUST_STORE_NO" & vbCrLf
            sqlSOTORDRB = ASCMAIN1.sql
            ASCMAIN1.sql &= "   and SOTORDR1.ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRB", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "SOTORDR0", "*")
            With .Tables("SOTORDR0").Columns
                .Add("ORDR_HOLD")
                .Add("ORDR_OVERRIDE_NOT_ALLOCATED")
                .Add("REASON_CODE")
                .Add("ORDR_ADDR_TYPE_ST")
                .Add("ORDR_HOLD_REASON")
                .Add("ORDR_INV_COMMENT")
                .Add("ORDR_SPECIAL_INST")
                .Add("ORDR_MESSAGE")
                .Add("SREP_CODE")
                .Add("SREP2_CODE")
                .Add("TERM_CODE")
                .Add("SHIP_VIA_CODE")
                .Add("FRT_TERMS")
            End With


            ASCMAIN1.sql = "Select * from SOTORDRE where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRE", "**", 0, True, "V")
            .Tables("SOTORDRE").Columns.Add("ATTACHMENT_EXT")

            ASCMAIN1.sql = "Select * from SOTORDXR where ORDR_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDXR", "**", 0, True, "V")

            Create_Relation("SOTORDRD", "SOTORDRC", "MLNO")
            With .Tables("SOTORDRD").Columns
                .Add("ORDR_QTY_CALC", GetType(System.Int64), "SUM(CHILD(SOTORDRD_SOTORDRC).ORDR_QTY)")
                .Add("ORDR_QTY_OPEN_CALC", GetType(System.Int64), "SUM(CHILD(SOTORDRD_SOTORDRC).ORDR_QTY_OPEN)")
                .Add("ORDR_QTY_PICK_CALC", GetType(System.Int64), "SUM(CHILD(SOTORDRD_SOTORDRC).ORDR_QTY_PICK)")
                .Add("ORDR_QTY_SHIP_CALC", GetType(System.Int64), "SUM(CHILD(SOTORDRD_SOTORDRC).ORDR_QTY_SHIP)")
                .Add("ORDR_QTY_CANC_CALC", GetType(System.Int64), "SUM(CHILD(SOTORDRD_SOTORDRC).ORDR_QTY_CANC)")
            End With

            ASCMAIN1.sql = "Select CUST_STORE_NO, ORDR_CUST_PO from SOTORDR1 where ORDR_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRS", "**", 0, False, "V", 1)
            With .Tables("SOTORDRS")
                .Columns.Add("TOTAL_QTY", GetType(System.Int64))
                .Columns.Add("TOTAL_AMT", GetType(System.Decimal))
                ' .PrimaryKey = New DataColumn() {.Columns("CUST_STORE_NO")}
            End With

            ASCMAIN1.sql = "Select ARTCUST2.* from ARTCUST2" & vbCrLf _
                 & " where CUST_CODE = :PARM1" & vbCrLf _
                 & "   and CUST_STORE_NO = :PARM2" & vbCrLf
            Create_TDA(.Tables.Add, "ARTCUST2_BT", "**", 0, False, "VV", 0)
            'Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "VV", 2)
            Create_TDA(.Tables.Add, "ARTCUST2", "*")

            ASCMAIN1.sql = "Select SOTORDR5.* from SOTORDR5" & vbCrLf _
                 & " where ORDR_NO = :PARM1" & vbCrLf _
                 & "   and CUST_ADDR_TYPE = :PARM2" & vbCrLf
            Create_TDA(.Tables.Add, "SOTORDR5_BT", "**", 0, False, "VV", 0)

            With .Tables.Add("EDTDOCS1")
                .Columns.Add("EDI_DOC_ID")
                .Columns.Add("EDI_DOC_DATE", GetType(System.DateTime))
                .Columns.Add("EDI_DOC_SEQ_NO")
                .Columns.Add("FILENAME_GEN")
                .Columns.Add("FILENAME_ABS")
                .Columns.Add("EDI_DOC_TEXT")
                .Columns.Add("EDI_DOC_STATUS")
                .Columns.Add("EDI_DOC_DESC")
                .Columns.Add("EDI_ISA_NO")
            End With

            ASCMAIN1.sql = "Select * from SOTWORK1 where WO_REF_TYPE = 'S' and WO_REF_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTWORK1", "**", 0, , "V", 1)
            ASCMAIN1.sql = "Select * from SOTWORK2 where WO_NO in " _
                & " (Select WO_NO from SOTWORK1 where WO_REF_TYPE = 'S' and WO_REF_NO = :PARM1)"
            Create_TDA(.Tables.Add, "SOTWORK2", "**", 0, , "V", 1)
        End With


        grdSOTORDRB.DataSource = dst.Tables("SOTORDRB")
        grdSOTORDRI.DataSource = dst.Tables("SOTORDRI")

        grdEDTDOCS1.DataSource = dst.Tables("EDTDOCS1")
        grdSOTORDRX.DataSource = dst.Tables("SOTORDRX")
        grdSOTORDR2.DataSource = dst.Tables("SOTORDR2")
        grdSOTORDR4.DataSource = dst.Tables("SOTORDR4")
        grdSOTORDRE.DataSource = dst.Tables("SOTORDRE")
        grdSOTORDXR.DataSource = dst.Tables("SOTORDXR")
        grdSOTORDRS.DataSource = dst.Tables("SOTORDRS")

        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")

        grdSOTORDRT.DataSource = dst.Tables("SOTORDRT")

        grdSOTORDRX.DisplayLayout.UseFixedHeaders = True
        With grdSOTORDRX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "ORDR_DATE", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTORDRS.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "ORDR_CUST_PO", "TOTAL_AMT", "TOTAL_QTY"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With

        With grdSOTORDRB.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Header.Appearance.BackColor = Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_NO", "CUST_STORE_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_CUST_PO"}
                With .Columns(COLUMN_NAME)
                    .Header.Fixed = True
                    .Header.Appearance.BackColor2 = Drawing.Color.Gold
                End With
            Next
        End With

        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"ORDR_LNO", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTORDR2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If Not New String() {"ITEM_CODE", "ORDR_QTY", "ORDR_UNIT_PRICE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next

            For Each COLUMN_NAME As String In New String() {"ORDR_RETAIL_AMT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_LNO", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ITEM_RETAIL_PRICE", "ORDR_UNIT_PRICE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_LNO", "ITEM_DESC"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
            Next
        End With


        With grdSOTORDRI.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If Not New String() {"ORDR_UNIT_PRICE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
            Next

            .Columns("COL").Hidden = True

            .Columns("ORDR_QTY").Hidden = True
            .Columns("ORDR_QTY_CANC").Hidden = True
            .Columns("ORDR_AMT").Hidden = True
            .Columns("ORDR_AMT_CANC").Hidden = True

            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next

            For Each COLUMN_NAME As String In New String() {"ORDR_RETAIL_AMT"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_CANC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ITEM_UOM", "CARTON_PACK_QTY", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
            For Each COLUMN_NAME As String In New String() {"ITEM_RETAIL_PRICE", "ORDR_UNIT_PRICE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            Next
        End With

        Create_Summary(grdSOTORDRX, "ORDR_NO", "Count")

        Create_Summary(grdSOTORDR2, "ORDR_LNO", "Count")
        Create_Summary(grdSOTORDR2, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "ORDR_RETAIL_AMT", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC"})

        Create_Summary(grdSOTORDRB, "CUST_STORE_NO", "Count")

        Create_Summary(grdSOTORDRI, "ITEM_CODE", "Count")
        Create_Summary(grdSOTORDRI, New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_CANC", "ORDR_AMT", "ORDR_AMT_OPEN", "ORDR_AMT_CANC", "ORDR_RETAIL_AMT"})

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL", "CART_TOTAL_WGT_CALC"})

        Create_Summary(grdSOTORDRS, "CUST_STORE_NO", "Count")
        Create_Summary(grdSOTORDRS, "TOTAL_AMT", "TOTAL_QTY")

        Bind_Controls(grpBILLTO, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'BT'", "", DataViewRowState.CurrentRows))
        Bind_Controls(grpSOLDTO, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'BY'", "", DataViewRowState.CurrentRows))
        dvwSOTORDR5 = New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'ST'", "", DataViewRowState.CurrentRows)
        Bind_Controls(grpSHIPTO, "SOTORDR5", dvwSOTORDR5)
        Bind_Controls(grpSTORE, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'MK'", "", DataViewRowState.CurrentRows))
        Bind_Controls(grpDC, "SOTORDR5", New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'DC'", "", DataViewRowState.CurrentRows))
        Bind_Controls(frmShipToOption, "SOTORDR1")
        Bind_Controls(grpSOLDTO_Attributes, "ARTCUST1")
        Bind_Controls(frmSOTORDRD, "SOTORDR0")

        'Set_Read_Only(grpStatus, True)
        Set_Read_Only(grpBILLTO, True)
        Set_Read_Only(grpSOLDTO, True)
        Set_Read_Only(grpSHIPTO, True)
        Set_Read_Only(grpSTORE, True)
        Set_Read_Only(grpDC, True)

        Set_Read_Only_for_ctl(Absx1.optFor("ORDR_SOURCE"), True)

        lblINV_NO.Visible = InquiryMode
        txtINV_NO.Visible = InquiryMode

        Absx1.txtFor("ORDR_PRIORITY").Enabled = ASCMAIN1.USER_SECURITY_CODEs.Contains("X2")

        With dst.Tables("SOTORDRT").Rows
            .Add(New Object() {1, "Order", 0, 0})
            .Add(New Object() {2, "Open", 0, 0})
            .Add(New Object() {3, "Pick", 0, 0})
            .Add(New Object() {4, "Ship", 0, 0})
            .Add(New Object() {5, "Canc", 0, 0})
        End With
        Sort_grdColumns(grdSOTORDRT, "KEY", True)

        Show_Filter(grdSOTORDRX, True)
        grdSOTORDRX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTORDRB, "ORDR_STATUS")

        Check_InquiryMode()
        Set_Read_Only(splOrderInfo, True)


        ASCMAIN1.sql = "Select T_CODE, T_DESC from ASTCODE1" _
            & " where TABLE_NAME = 'SOTORDR1' and COLUMN_NAME = 'ORDR_SOURCE'" _
            & " and T_CODE Not in ('E','K')"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim T_CODE As String = row.Item("T_CODE")
            Dim T_DESC As String = row.Item("T_DESC")
            Absx1.optFor("ORDR_SOURCE").ValueList.ValueListItems.Add(New ValueListItem(T_CODE, T_DESC))
        Next
    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFORDRI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Dim rowARTCUST2 As DataRow = Nothing
                multiple_order_maintenance = False
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                            If rowARTCUST1.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowARTCUST1.Item("SREP_CODE") & "") Then

                                Dim found_store As Boolean = False
                                ASCMAIN1.sql = "Select Distinct SREP_CODE from ARTCUST2 where CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "' and CUST_STORE_NO = '" & Absx1.txtFor("CUST_STORE_NO").Text & "'"
                                For Each rowARTCUST2_SREP As DataRow In ASCDATA1.GetDataTable.Select("")
                                    If rowARTCUST2_SREP.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                        And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowARTCUST2_SREP.Item("SREP_CODE") & "") Then
                                    Else
                                        found_store = True
                                    End If
                                Next

                                If Not found_store Then
                                    EMsg &= vbCr & "Customer " & Absx1.txtFor("CUST_CODE").Text & " is not connected to Sales Rep code " & TAC.TACMAIN1.SREP_CODE
                                End If
                            End If
                        End If

                        If EMsg = "" Then
                            CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                            CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
                            If CUST_BILL_TO_CUST = "" Then
                                CUST_BILL_TO_CUST = CUST_CODE
                            End If

                            ' apostrophe in Cust PO causes ABSolution to crash when lookig to see if it is a duplicate PO entry
                            Absx1.txtFor("ORDR_CUST_PO").Text = Absx1.txtFor("ORDR_CUST_PO").Text.Trim.Replace("'", "")
                            ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                            If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                                EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                            End If

                            Dim PRICE_CLASS_CODE As String = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
                            If LookUp("SOTPCLS1", PRICE_CLASS_CODE) Is Nothing Then
                                EMsg &= vbCr & "Invalid or missing value for Price Class Code(" & PRICE_CLASS_CODE & ")"
                            End If
                        End If

                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If Absx1.txtFor("CUST_STORE_NO").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Store (Mark-For)"
                Else
                    rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, Absx1.txtFor("CUST_STORE_NO").Text})

                    If rowARTCUST2 Is Nothing Then
                        If Absx1.txtFor("CUST_STORE_NO").Text = "000000" Then
                            ASCMAIN1.sql = "Insert into ARTCUST2 (" & vbCrLf _
                                & "CUST_CODE, CUST_STORE_NO, CUST_STORE_NAME," & vbCrLf _
                                & "CUST_STORE_ADDR1, CUST_STORE_ADDR2, CUST_STORE_ADDR3," & vbCrLf _
                                & "CUST_STORE_CITY, CUST_STORE_STATE, CUST_STORE_ZIP_CODE," & vbCrLf _
                                & "CUST_STORE_COUNTRY, CUST_STORE_CONTACT, " & vbCrLf _
                                & "CUST_STORE_PHONE, CUST_STORE_EXT, CUST_STORE_FAX," & vbCrLf _
                                & "CUST_STORE_EMAIL, CUST_STORE_STATUS," & vbCrLf _
                                & "INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE)" & vbCrLf _
                                & "Select " & vbCrLf _
                                & "CUST_CODE, '000000'," & vbCrLf _
                                & "CUST_NAME, CUST_ADDR1, CUST_ADDR2, CUST_ADDR3," & vbCrLf _
                                & "CUST_CITY, CUST_STATE, CUST_ZIP_CODE," & vbCrLf _
                                & "CUST_COUNTRY, CUST_CONTACT," & vbCrLf _
                                & "CUST_PHONE, CUST_EXT, CUST_FAX," & vbCrLf _
                                & "CUST_EMAIL, 'A', " & vbCrLf _
                                & "'" & ASCMAIN1.USER_ID & "' INIT_OPER, '" & ASCMAIN1.USER_ID & "' LAST_OPER, SYSDATE INIT_DATE, SYSDATE LAST_DATE" & vbCrLf _
                                & " from ARTCUST1 WHERE CUST_CODE = '" & CUST_CODE & "'"
                            ASCDATA1.ExecuteSQL()
                            rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, Absx1.txtFor("CUST_STORE_NO").Text})
                        End If
                    End If

                    If rowARTCUST2 IsNot Nothing Then
                        CUST_STORE_NO = Absx1.txtFor("CUST_STORE_NO").Text
                    Else
                        EMsg &= vbCr & "No Record of Customer Store " & Absx1.txtFor("CUST_STORE_NO").Text
                    End If
                End If


                If EMsg = "" Then
                    ' Check Copy From Order to verify that it exists, and that it is with same customer
                    If Absx1.txtFor("ORDR_NO").Text <> "" Then
                        Dim rowSOTORDR1_CopyFrom As DataRow = LookUp("SOTORDR1", Absx1.txtFor("ORDR_NO").Text)
                        If rowSOTORDR1_CopyFrom Is Nothing Then
                            EMsg &= vbCr & "No Record of (Copy From) Order No " & Absx1.txtFor("ORDR_NO").Text
                        ElseIf rowSOTORDR1_CopyFrom.Item("CUST_CODE") <> CUST_CODE Then
                            EMsg &= vbCr & "Copy Order Feature works with Same Customer Only"
                        End If
                    End If

                    ' Customer must have a Sales Rep assigned, look at Ship to first.
                    SREP_CODE = rowARTCUST2.Item("SREP_CODE") & ""
                    If SREP_CODE.Length = 0 Then
                        SREP_CODE = rowARTCUST1.Item("SREP_CODE") & ""
                    End If

                    Dim rowSOTSREP1 As DataRow = Nothing
                    If SREP_CODE <> "" Then rowSOTSREP1 = LookUp("SOTSREP1", SREP_CODE)
                    If rowSOTSREP1 Is Nothing Then
                        EMsg &= vbCr & "This Customer Has No Sales Rep Assigned"
                    End If
                    SREP2_CODE = rowARTCUST1.Item("SREP2_CODE") & ""
                End If


                If EMsg = "" Then
                    ' Load Default values in for Selected Fields if we have seen this Customer PO before
                    ASCMAIN1.sql = "Select ORDR_GROUP_NO, CUST_STORE_NO, ORDR_SHIP_DATE, ORDR_CANCEL_DATE" & vbCrLf _
                        & ", ORDR_DATE, ORDR_DEPT, ORDR_SPECIAL_INST, FRT_TERMS, SALES_DIVISION_CODE" & vbCrLf _
                        & " from SOTORDR1 where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                        & " and ORDR_CUST_PO = '" & ORDR_CUST_PO & "'" & vbCrLf _
                        & " order by ORDR_SHIP_DATE DESC"
                    ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where ROWNUM < 2"
                    rowSOTORDR1po = ASCDATA1.GetDataRow

                    If rowSOTORDR1po IsNot Nothing Then
                        ASCMAIN1.sql = "Select ORDR_NO, ORDR_DATE from SOTORDR1 " & vbCrLf _
                            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                            & "   and CUST_STORE_NO = '" & CUST_STORE_NO & "'" & vbCrLf _
                            & "   and ORDR_CUST_PO = '" & ORDR_CUST_PO & "'" & vbCrLf _
                            & "   and ORDR_STATUS in ('O','P','F')"
                        Dim rowDup As DataRow = ASCDATA1.GetDataRow
                        If rowDup IsNot Nothing Then
                            If MsgBox("Same Customer PO has already been entered for Store " & CUST_STORE_NO _
                                      & vbCrLf & " (See Sales Order " & rowDup.Item("ORDR_NO") & " dated " & Format(rowDup.Item("ORDR_DATE"), "MM/dd/yyyy") & ")" _
                                      & vbCrLf & vbCrLf & "Are You Sure that you want to Proceed?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, _
                                      "Possible Order Duplication") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                    'If Not ASCMAIN1.Logical_Lock("SOTORDR1", CUST_CODE) Then Exit Sub
                End If

            Case "Edit", "View"
                multiple_order_maintenance = False

                CUST_CODE = ""
                ORDR_NO = ""

                If Absx1.txtFor("ORDR_NO").Text = "" Then
                    EMsg &= vbCr & "No Order No Specified"
                Else
                    ORDR_NO = Absx1.txtFor("ORDR_NO").Text
                    rowSOTORDR1 = LookUp("SOTORDR1", ORDR_NO)
                    If rowSOTORDR1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Sales Order No " & ORDR_NO
                    Else

                        If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                            If rowSOTORDR1.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowSOTORDR1.Item("SREP_CODE") & "") Then
                                EMsg &= vbCr & "Order " & ORDR_NO & " is not connected to Sales Rep code " & TAC.TACMAIN1.SREP_CODE
                            End If
                        End If

                        ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO")
                        ORDR_GROUP_NOs.Clear()
                        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                        CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
                        EDI_JRNL_NO = rowSOTORDR1.Item("EDI_JRNL_NO") & ""

                        ASCMAIN1.sql = "Select Count (*) ORDR_CNT" & vbCrLf _
                            & ", SUM(DECODE(ORDR_STATUS,'O',1,0)) O" _
                            & ", SUM(DECODE(ORDR_STATUS,'P',1,0)) P" _
                            & ", SUM(DECODE(ORDR_STATUS,'C',1,0)) C" _
                            & ", SUM(DECODE(ORDR_STATUS,'F',1,0)) F" _
                            & " from SOTORDR1" & vbCrLf _
                            & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                            & "   and ORDR_STATUS in ('O','C','P','F')"
                        Dim rowSTATS As DataRow = ASCDATA1.GetDataRow
                        If Val(rowSTATS.Item("ORDR_CNT") & "") > 1 Then
                            multiple_order_maintenance = True
                            multiple_order_type = "ORDR_GROUP_NO"
                        End If

                        If Not multiple_order_maintenance And EDI_JRNL_NO <> "" Then
                            ASCMAIN1.sql = "Select Count (*) ORDR_CNT" & vbCrLf _
                               & ", SUM(DECODE(ORDR_STATUS,'O',1,0)) O" _
                               & ", SUM(DECODE(ORDR_STATUS,'P',1,0)) P" _
                               & ", SUM(DECODE(ORDR_STATUS,'C',1,0)) C" _
                               & ", SUM(DECODE(ORDR_STATUS,'F',1,0)) F" _
                               & " from SOTORDR1" & vbCrLf _
                               & " where EDI_JRNL_NO = '" & EDI_JRNL_NO & "'" & vbCrLf _
                               & "   and ORDR_STATUS in ('O','C','P','F')"
                            rowSTATS = ASCDATA1.GetDataRow
                            If Val(rowSTATS.Item("ORDR_CNT") & "") > 1 Then
                                multiple_order_maintenance = True
                                multiple_order_type = "EDI_JRNL_NO"
                            End If
                        End If

                        If multiple_order_maintenance Then
                            ' PROBABLY WILL NEED A REVERSE CANCELLATION OPTION FOR ORDERS IN A MOG
                            If Not InquiryMode Then
                                If Val(rowSTATS.Item("O") & "") = 0 Then
                                    EMsg &= vbCr & "Sales Order No " & ORDR_NO & " belongs to a Multiple-Order Group" _
                                        & vbCr & "- No Orders are Open in that group"
                                End If
                            End If
                        Else
                            If rowSOTORDR1.Item("ORDR_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                                Select Case rowSOTORDR1.Item("ORDR_STATUS")
                                    Case "C"
                                        MsgBox("Sales Order No " & ORDR_NO & " has been Cancelled", MsgBoxStyle.OkOnly, _
                                                "Cannot Edit Order")
                                        If MsgBox("Re-Open Order for Processing", MsgBoxStyle.YesNo, _
                                                    "Answer Yes to Reverse Cancellation of this Order") = MsgBoxResult.No Then
                                            Exit Sub
                                        Else
                                            Select Case MsgBox("Reverse Cancellation of All Orders in this Group", _
                                                                MsgBoxStyle.YesNoCancel, _
                                                                "Yes for All Orders in Group, No for this order only")
                                                Case MsgBoxResult.Yes
                                                    ASCMAIN1.sql = "Select ORDR_NO, CUST_STORE_NO from SOTORDR1" _
                                                        & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
                                                        & "   and ORDR_STATUS = 'C'"
                                                    Dim dt As DataTable = ASCDATA1.GetDataTable
                                                    Using F As New ASFMSGBF
                                                        F.Show_grd(dt, Me, "The following Orders will be Restored")
                                                        If F.user_option = -1 Then
                                                            MsgBox("Restoration of Cancelled Orders was NOT Performed", MsgBoxStyle.OkOnly, "Please Note")
                                                        Else
                                                            If ASCMAIN1.Running_in_VS Then Stop ' send in list of orders
                                                            Reverse_Cancel("", ORDR_GROUP_NO, dt)
                                                            MsgBox("Restoration of Cancelled Orders is Complete", MsgBoxStyle.OkOnly, "Please Note")
                                                        End If
                                                    End Using

                                                Case MsgBoxResult.No
                                                    Reverse_Cancel(ORDR_NO, "")

                                                Case MsgBoxResult.Cancel
                                                    Exit Sub
                                            End Select

                                        End If

                                    Case "D"
                                        EMsg &= vbCr & "Sales Order No " & ORDR_NO & " has been Deleted"
                                    Case "P"
                                        EMsg &= vbCr & "Sales Order No " & ORDR_NO & " has been Completely Released for Picking"
                                    Case Else ' such as "F"
                                        EMsg &= vbCr & "Sales Order No " & ORDR_NO & " is No Longer Open"
                                End Select
                            End If
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    disable_update = False
                    'ASCMAIN1.sql = "Select Count (*) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS in ('O','P','F')"
                    'If Val(ASCDATA1.GetDataValue) > 1 Then
                    '    If MsgBox("You Cannot Make Changes to an Order belonging to a Multiple Order Group" _
                    '              & vbCr & vbCr & "However, you Can Delete or Cancel this Order (Store) using this Option.", _
                    '              MsgBoxStyle.OkCancel, _
                    '              "Note: Order " & ORDR_NO & " is part of Order Group (" & ORDR_GROUP_NO & ")") = MsgBoxResult.Cancel Then
                    '        Exit Sub
                    '    End If
                    '    disable_update = True
                    'End If
                    If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub

                    If multiple_order_maintenance Then
                        ASCMAIN1.sql = "Select ORDR_NO" & vbCrLf _
                           & " from SOTORDR1" & vbCrLf _
                           & IIf(multiple_order_type = "ORDR_GROUP_NO", _
                                 " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", _
                                 " where EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") & vbCrLf _
                           & "   and ORDR_STATUS in ('O','C','P','F')"
                        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                            If Not ASCMAIN1.Logical_Lock("SOTORDR1", row.Item("ORDR_NO")) Then
                                multiple_order_maintenance = False
                                Exit Sub
                            End If
                        Next
                    End If

                    If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOFOREL1", CUST_CODE) Then Exit Sub
                    If Not ASCMAIN1.Logical_Open("R", "SOROREL1") Then Exit Sub
                End If

            Case "Update"

                Dim TERM_TYPE As String = String.Empty

                If CUST_CODE = "CONSUMER" Then
                    If rowSOTORDR1.Item("ORDR_SOURCE") <> "W" Then
                        If Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "") <> 0 Then
                            EMsg &= vbCr & "Cannot Accept Revenue Orders for CONSUMER" & vbCr & " - you must use Web Site (Credit Card & Sales Tax Issues)"
                        End If
                    End If
                End If

                If EntryMode = "M" Or multiple_order_maintenance Then

                    Dim ORDR_SHIP_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value
                    Dim ORDR_CANCEL_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_CANCEL_DATE").Value

                    If Format(ORDR_SHIP_DATE, "yyyyMMdd") > Format(ORDR_CANCEL_DATE, "yyyyMMdd") Then
                        EMsg &= vbCr & "Ship Date Cannot be Later than Cancel Date"
                    End If

                    If Absx1.dteFor("SOTORDR0.ORDR_ARRIVAL_DATE").Value & "" <> "" And Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value & "" <> "" Then
                        If Format(Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value, "yyyyMMdd") > Format(Absx1.dteFor("SOTORDR0.ORDR_ARRIVAL_DATE").Value, "yyyyMMdd") Then
                            EMsg &= vbCr & "Arrival date must not be prior to Ship-By Date"
                        End If
                    End If

                    Dim rowICTWHSE1 As DataRow = Nothing
                    Dim WHSE_CODE As String = Absx1.txtFor("SOTORDR0.WHSE_CODE").Text
                    If WHSE_CODE = "" Then
                        EMsg &= vbCr & "Warehouse is required"
                    Else
                        rowICTWHSE1 = LookUp("ICTWHSE1", WHSE_CODE)
                        If rowICTWHSE1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Warehouse Code"
                        End If
                    End If

                    Dim FRT_TERMS As String = Absx1.txtFor("SOTORDR0.FRT_TERMS").Text
                    If FRT_TERMS = "" Then
                        EMsg &= vbCr & "Freight Terms are Mandatory"
                    Else
                        Dim row As DataRow = LookUp("ASTCODE1", New String() {"ARTCUST1", "FRT_TERMS", FRT_TERMS})
                        If row Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Freight Terms"
                        Else

                        End If
                    End If


                    Dim TERM_CODE As String = Absx1.txtFor("SOTORDR0.TERM_CODE").Text
                    If TERM_CODE = "" Then
                        EMsg &= vbCr & "Terms Code is required"
                    Else
                        LookUp("TATTERM1", TERM_CODE)
                        If cdr Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Terms Code"
                        Else
                            TERM_TYPE = cdr.Item("TERM_TYPE") & String.Empty
                        End If
                    End If


                    Dim SREP_CODE As String = Absx1.txtFor("SOTORDR0.SREP_CODE").Text
                    If SREP_CODE = "" Then
                        EMsg &= vbCr & "Sales Rep is required"
                    Else
                        LookUp("SOTSREP1", SREP_CODE)
                        If cdr Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Sales Rep Code"
                        End If
                    End If

                    Dim SREP2_CODE As String = Absx1.txtFor("SOTORDR0.SREP2_CODE").Text
                    If SREP2_CODE = "" Then
                    Else
                        LookUp("SOTSREP1", SREP2_CODE)
                        If cdr Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Sales Rep2 Code"
                        End If
                    End If

                    Dim SHIP_VIA_CODE As String = Absx1.txtFor("SOTORDR0.SHIP_VIA_CODE").Text
                    If SHIP_VIA_CODE = "" Then
                        EMsg &= vbCr & "Ship Via is required"
                    Else
                        Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                        If rowSOTSVIA1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value specified for Ship Via Code"
                        Else
                            If rowSOTSVIA1.Item("SHIP_VIA_STATUS") & "" <> "A" Then
                                EMsg &= vbCr & "Ship Via Code specified is Inactive"
                            End If
                        End If
                    End If

                    If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                        EMsg &= vbCr & "Customer PO is required"
                    End If


                Else
                    If Absx1.txtFor("FRT_TERMS").Text = "" Then
                        EMsg &= vbCr & "Freight Terms are Mandatory"
                    Else
                        Dim row As DataRow = LookUp("ASTCODE1", New String() {"ARTCUST1", "FRT_TERMS", Absx1.txtFor("FRT_TERMS").Text})
                        If row Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Freight Terms"
                        End If
                    End If

                    ' WHAT IS THIS SECTION DOING?
                    Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

                    If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                    Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                    Else
                        If EMsg = "" And Absx1.optFor("ORDR_SOURCE").Value <> "E" Then
                            If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                             > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                                EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                            End If
                        End If
                    End If

                    If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                        Dim ORDR_AMT As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
                        If System.Math.Round(TOTAL_ORDR_AMT, 2) _
                        <> System.Math.Round(ORDR_AMT, 2) Then
                            Dim EDI_Value_Remark As String
                            Using F As New ASFMSGBF
                                EDI_Value_Remark = F.Get_txt_from_User(String.Format( _
                                    "Total Value Of Order ({0}) Does Not Equal Original EDI Amount ({1})", _
                                        Format(ORDR_AMT, "#,###0.00"), _
                                        Format(TOTAL_ORDR_AMT, "#,###0.00")), _
                                    "Please Provide A Reason For This Change In Order To Proceed", False, 50)
                            End Using

                            If EDI_Value_Remark = "" Then
                                EMsg &= vbCr & "EDI Value Change Canceled Or Reason Not Provided."
                            Else
                                rowSOTORDR1.Item("EDI_VALUE_CHANGE_REMARK") = EDI_Value_Remark
                                rowSOTORDR1.Item("EDI_VALUE_CHANGE_OPER") = ASCMAIN1.USER_ID
                                rowSOTORDR1.Item("EDI_VALUE_CHANGE_DATE") = DATETIME_STAMP
                            End If
                        End If

                        Using dt As New DataTable
                            dt.Columns.Add("LNO", GetType(System.Int64))
                            dt.Columns.Add("TYPE")
                            dt.Columns.Add("ORDR_QTY", GetType(System.Int64))
                            dt.Columns.Add("ORDR_QTY_ORIG", GetType(System.Int64))
                            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select _
                                    ("ISNULL(ORDR_QTY,0) <> ISNULL(ORDR_QTY_ORIG,0)")
                                dt.Rows.Add(New Object() {rowSOTORDR2.Item("ORDR_LNO"), _
                                                            "Item " & rowSOTORDR2.Item("ITEM_CODE"), _
                                                            rowSOTORDR2.Item("ORDR_QTY"), _
                                                            rowSOTORDR2.Item("ORDR_QTY_ORIG")})
                            Next
                            If dt.Rows.Count <> 0 Then
                                Using F As New ASFMSGBF
                                    F.Show_grd(dt, Me, _
                                               "The following lines have an Order Qty that has changed from the Original Qty", _
                                               "Please Verify that it is OK to Continue")
                                    If F.user_option = -1 Then
                                        Exit Sub
                                    End If
                                End Using
                            End If
                        End Using
                    End If

                    If multi_store_is_active Then
                        If Absx1.optFor("ORDR_ADDR_TYPE_ST").Value = "DC" Then
                            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                                Dim CUST_STORE_NO As String = rowSOTORDRS.Item("CUST_STORE_NO") & ""
                                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                                If rowARTCUST2.Item("CUST_DC_NO") & "" = "" Then
                                    EMsg &= vbCr & "No DC set up for Store " & CUST_STORE_NO
                                Else
                                    Dim rowARTCUST2_DC As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, rowARTCUST2.Item("CUST_DC_NO")})
                                    If rowARTCUST2_DC Is Nothing OrElse rowARTCUST2_DC.Item("CUST_DC_IND") & "" <> "1" Then
                                        EMsg &= vbCr & "Invalid DC set up for Store " & CUST_STORE_NO
                                    End If
                                End If
                            Next
                        End If
                    End If

                    If Absx1.txtFor("TERM_CODE").Text = "" Then
                        EMsg &= vbCr & "Terms Code is required"
                    Else
                        Validate_Code("TERM_CODE")
                        If cdr IsNot Nothing Then
                            TERM_TYPE = cdr.Item("TERM_TYPE") & String.Empty
                        End If
                    End If

                    If Absx1.txtFor("SREP_CODE").Text = "" Then
                        EMsg &= vbCr & "Sales Rep is required"
                    Else
                        Validate_Code("SREP_CODE")
                    End If

                    Validate_Code("SREP2_CODE", False, True)
                    Validate_Code("WHSE_CODE")
                    Validate_Code("SHIP_VIA_CODE")

                    If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                        EMsg &= vbCr & "Customer PO is required"
                    End If

                    If grdSOTORDR2.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Items on Order"
                    Else
                        If Val(dst.Tables("SOTORDR2").Compute("COUNT(ORDR_LNO)", "ORDR_QTY > 0") & "") = 0 Then
                            EMsg &= vbCr & "No Items on Order with Qty >0"
                        End If

                        Dim ITEM_CODEs As String = ""
                        For Each TABLE_NAME As String In New String() {"SOTORDR2"}
                            For Each row As DataRow In ASCDATA1.SelectDistinct _
                                    (dst.Tables("SOTORDR2").Select("ITEM_CODE is Not Null"), "ITEM_CODE").Rows
                                Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
                                ITEM_CODEs &= ",'" & ITEM_CODE & "'"
                            Next
                        Next

                        If ITEM_CODEs <> "" Then
                            ASCMAIN1.sql = "Select Distinct SALES_DIVISION_CODE from ICTITEM1 where ITEM_CODE IN (" & Mid$(ITEM_CODEs, 2) & ")"
                            Dim rows() As DataRow = ASCDATA1.GetDataTable.Select
                            If rows.Length > 1 Then
                                If MsgBox("Order Contains Styles From Different Divisions" _
                                          & vbCrLf & "Are You Sure You Want To Continue?", MsgBoxStyle.YesNo, "Mixed Styles") = MsgBoxResult.No Then
                                    Exit Sub
                                End If
                            Else
                                If rows(0).Item("SALES_DIVISION_CODE") & "" <> Absx1.txtFor("SALES_DIVISION_CODE").Text And Absx1.txtFor("SALES_DIVISION_CODE").Text <> "" Then
                                    If MsgBox("Order Contains Styles From a Sales Divison Other Than " & Absx1.txtFor("SALES_DIVISION_CODE").Text _
                                              & vbCrLf & "Are You Sure You Want To Continue?", MsgBoxStyle.YesNo, "Mixed Styles") = MsgBoxResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            End If
                        End If
                    End If

                    If EMsg.Length = 0 Then

                        Dim ORDR_TOTAL_AMT As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & String.Empty) _
                                                + Val(rowSOTORDR1.Item("ORDR_FREIGHT") & String.Empty) + Val(rowSOTORDR1.Item("ORDR_STAX") & String.Empty)
                        If ORDR_TOTAL_AMT > 0 AndAlso TERM_TYPE = "D" _
                            AndAlso Not MyBase.Absx1.chkFor("ORDR_HOLD").Checked _
                            AndAlso (rowSOTORDR1.Item("CCPA_NO") & String.Empty).ToString.Length = 0 Then

                            If MessageBox.Show("Credit Card Terms require a Credit Card authorization." _
                               & vbCr & "Updating the order without Credit Card Authorization will cause the order to be held up in the release process." _
                               & vbCr & "Update Anyway?", "Credit Card", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTPICK1 where PICK_STATUS <> 'D' "
                If EntryMode = "M" Or multiple_order_maintenance Then
                    ASCMAIN1.sql &= " and ORDR_NO in (Select ORDR_NO from SOTORDR1 " _
                        & IIf(multiple_order_type = "ORDR_GROUP_NO", _
                              "where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", _
                              "where EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") _
                        & ")"
                Else
                    ASCMAIN1.sql &= " and ORDR_NO = '" & ORDR_NO & "'"
                End If

                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Order has been Released and/or Partially Shipped"
                Else
                    If EntryMode = "M" Or multiple_order_maintenance Then
                        If Absx1.txtFor("SOTORDR0.REASON_CODE").Text = "" Then
                            EMsg &= vbCr & "You must Specify a Reason Code when Deleting a Group of Orders"
                        Else
                            If LookUp("SOTREAS1", Absx1.txtFor("SOTORDR0.REASON_CODE").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                            End If
                        End If
                        If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'P'").Length <> 0 Then
                            EMsg &= vbCr & "No Orders May be in Pick when Deleting an Orders in a Group"
                        End If
                        If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'O'").Length = 0 Then
                            EMsg &= vbCr & "No Open Orders to Delete"
                        End If
                    Else
                        If Absx1.txtFor("REASON_CODE").Text = "" Then
                            EMsg &= vbCr & "You must Specify a Reason Code when Deleting an Order"
                        Else
                            If LookUp("SOTREAS1", Absx1.txtFor("REASON_CODE").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                            End If
                        End If
                    End If

                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Order as Deleted", _
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Start Multi-Store"
                If grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.DataChanged Then
                    grdSOTORDR2.ActiveRow.Update()
                End If

                If rowSOTORDR1.Item("CCPA_NO") & "" <> "" Then
                    EMsg &= vbCr & "Cannot do a Multi-Store Order which is associated with a Credit Card"
                End If

                If multi_store_is_active Then
                    If MsgBox("Are You Sure that you want to Clear All Multi-Store Entries", _
                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Setup_MS(False)
                    End If
                    Exit Sub
                End If

                If grdSOTORDR2.Rows.Count = 0 Then
                    EMsg &= vbCr & "You Must First Add Line Items to This Order"
                Else

                End If


            Case "Cancel Order"

                If EntryMode = "M" Or multiple_order_maintenance Then
                    If Absx1.txtFor("SOTORDR0.REASON_CODE").Text = "" Then
                        EMsg &= vbCr & "You must Specify a Reason Code when Cancelling an Order"
                    Else
                        If LookUp("SOTREAS1", Absx1.txtFor("SOTORDR0.REASON_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                        End If
                    End If
                    If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'P'").Length <> 0 Then
                        EMsg &= vbCr & "No Orders May be in Pick when Cancelling an Orders in a Group"
                    End If
                    If dst.Tables("SOTORDRB").Select("ORDR_STATUS = 'O'").Length = 0 Then
                        EMsg &= vbCr & "No Open Orders to Cancel"
                    End If
                Else
                    If Absx1.txtFor("REASON_CODE").Text = "" Then
                        EMsg &= vbCr & "You must specify a Reason Code when Cancelling an Order"
                    Else
                        If LookUp("SOTREAS1", Absx1.txtFor("REASON_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Reason Code"
                        End If
                    End If
                End If
                If EMsg = "" Then
                    If MsgBox("Do you want to Cancel (the remaining open balance on) this Order" _
                               & vbCrLf & "(Lost Sales will be charged)", _
                               vbYesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Credit Card"
                If Absx1.txtFor("ORDR_NO").Text = "" Then
                    EMsg &= vbCr & "No Order No Specified"
                Else
                    ASCMAIN1.sql = "Select * from ARTCCPA1" _
                    & " where ORDR_NO = '" & HFs("ORDR_NO") & "'" _
                    & " and CCPA_STATUS IN ('C','T')"
                    Dim rowARTCCPA1 As DataRow = ASCDATA1.GetDataRow
                    If rowARTCCPA1 IsNot Nothing Then
                        EMsg &= "CC Authorization for " & Format(Val(rowARTCCPA1.Item("CCPA_AMT") & ""), "$#.00") & " already recorded for this order"
                        Exit Select
                    End If

                    If multi_store_is_active Then
                        EMsg &= vbCr & "Cannot do a Multi-Store Order which is associated with a Credit Card"
                    End If
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                If EntryMode = "M" Or multiple_order_maintenance Then
                    Update_Record_multiple_order_group()
                    ' Update_MOM()
                Else
                    Update_Record()
                End If
                Mode_Settings(False)

            Case "Delete"
                'Delete_Record()
                Delete_Order()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Cancel Order"
                Cancel_Order()
                Mode_Settings(False)

            Case "Start Multi-Store"
                ReSelect_Stores()

            Case "Reset Qty's"
                If MsgBox("Are You Sure that you want to Reset the Qty's to Store " & Absx1.txtFor("CUST_STORE_NO").Text & "'s Values", _
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    Init_MultiStore()
                End If

            Case "Re-Select Stores"
                ReSelect_Stores()

            Case "Clear Zeroes"
                Clear_Zeroes()

            Case "Work Orders"
                Using F As New TAC.SOFWORK1(Me, "S", ORDR_NO, (EntryMode = "V" Or InquiryMode), _
                                            Absx1.txtFor("CUST_CODE").Text, _
                                            Absx1.txtFor("ORDR_CUST_PO").Text,
                                            Absx1.dteFor("ORDR_SHIP_DATE").Value, _
                                             Absx1.dteFor("ORDR_CANCEL_DATE").Value, _
                                            "Work Orders relating to Sales Order " & ORDR_NO)
                    F.ShowDialog()
                End Using


            Case "Credit Card"
                Credit_Card()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowSOTORDR1.Item("ORDR_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If

                    ' Credit Cards can be Authorized when in Open and Pick.
                    If ",O,P,".Contains(rowSOTORDR1.Item("ORDR_STATUS") & String.Empty) Then
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    If (EntryMode = "E" Or EntryMode = "N") Then
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Credit Card").Settings.Enabled = DefaultableBoolean.False
                    End If
                End If

                .Items("Update").Settings.Enabled = iScreenMode
                If disable_update Then
                    .Items("Update").Settings.Enabled = DefaultableBoolean.False
                End If
                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode

                .Items("Cancel Order").Settings.Enabled = iScreenMode

                '.Items("Pro-Rate").Visible = (EntryMode <> "V")

                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode
                .Items("Credit Card").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode
                .Items("Print").Visible = (EntryMode = "V" And ScreenMode) Or InquiryMode

                .Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode
                .Items("Delete").Visible = (EntryMode = "E") And Not InquiryMode
                .Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode) And Not InquiryMode

                .Items("Cancel Order").Visible = (EntryMode = "E")

                .Items("Work Orders").Text = "Work Orders" & IIf(dst.Tables("SOTWORK1").Rows.Count = 0, "", " (" & CStr(dst.Tables("SOTWORK1").Rows.Count) & ")")
                .Items("Work Orders").Visible = False ' ScreenMode And Not (ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN")

                If (EntryMode = "N") Then
                    .Items("Credit Card").Visible = (MENU_ITEM_OBJECT = "SOFORDR1") And Not disable_update
                    .Items("Credit Card").Settings.Enabled = iScreenMode
                End If

                If disable_update Or multiple_order_maintenance Or InquiryMode Or Not ScreenMode Then
                    .Items("Credit Card").Visible = False
                Else
                    .Items("Credit Card").Visible = True
                End If
            End With

            .Groups("Show Orders").Visible = Not ScreenMode

            .Groups("Multi-Store").Visible = False

            .Groups("Totals").Visible = ScreenMode
            .Groups("Copy").Visible = False '  Not ScreenMode And Not InquiryMode
            .Groups("Release Holds").Visible = ScreenMode And (EntryMode = "V") AndAlso (rowSOTORDR1.Item("ORDR_STATUS") = "O")

            With .Groups("Multi-Store")
                .Visible = (EntryMode = "N") And dst.Tables("ARTCUST2").Rows.Count > 1
                .Items("Start Multi-Store").Settings.Enabled = DefaultableBoolean.True
                .Items("Reset Qty's").Settings.Enabled = DefaultableBoolean.False
                .Items("Re-Select Stores").Settings.Enabled = DefaultableBoolean.False
                .Items("Clear Zeroes").Settings.Enabled = DefaultableBoolean.False
            End With

        End With

        btnOrderForm.Visible = (EntryMode = "N" Or EntryMode = "E")

        lblMultiStore.Visible = False
        lblStatus.Visible = ScreenMode

        'If CUST_CODE = "CONSUMER" Then
        '    Set_Read_Only(grpSHIPTO, False)
        'Else
        '    Set_Read_Only(grpSHIPTO, True)
        'End If
        chkEdit.Checked = False
        Setup_SHIPTO_Edit()
        chkEdit.Visible = (EntryMode = "N" Or EntryMode = "E") And Not multiple_order_maintenance And (CUST_CODE = "CONSUMER")

        Set_Read_Only(SplitContainer8.Panel1, True)

        tabMain.Tabs("Multi-Store").Visible = False
        tabMain.Tabs("Pick / Cartons").Visible = (EntryMode = "V") And dst.Tables("SOTPICK1").Rows.Count > 0

        grdSOTORDRX.Visible = Not tf
        tabMain.Visible = tf And Not ((EntryMode = "E") And multiple_order_maintenance)
        frmSOTORDRD.Visible = tf And ((EntryMode = "E") And multiple_order_maintenance)

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
            If PRICE_BASIS = "R" Then
                grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.NoEdit
                grdSOTORDRI.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.NoEdit
            Else
                grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSOTORDRI.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
            'grdSOTORDR2.DisplayLayout.Bands(0).Columns("VOUCHER_NO").Hidden = Not (EntryMode = "V")
            If EntryMode = "V" Then
                grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTORDR2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

                grdSOTORDR4.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTORDR4.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTORDR4.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                grdSOTORDRS.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTORDRS.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                Set_Read_Only_for_ctl(optShipTo, True)

                grdSOTORDRI.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTORDRI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTORDRI.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdSOTORDR2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTORDR2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True


                grdSOTORDR4.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdSOTORDR4.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTORDR4.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

                grdSOTORDRS.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTORDRS.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTORDRS.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

                Set_Read_Only(SplitContainer8.Panel1, False)

                grdSOTORDRI.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTORDRI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTORDRI.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False

                'If EntryMode <> "E" Then
                '    Set_Read_Only(grpSHIPTO, False)
                'End If
                Set_Read_Only_for_ctl(optShipTo, (CUST_DC_NO = ""))
            End If
            'If CUST_DC_NO = "" Then
            '    frmShipToOption.Enabled = False
            'Else
            '    frmShipToOption.Enabled = True
            'End If

            Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_INV_COMMENT"), (EntryMode = "V"))
            Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), (EntryMode = "V"))
            Set_Read_Only_for_ctl(Absx1.optFor("ORDR_SOURCE"), True)
            tabHeader.Tabs("EDI").Visible = Absx1.optFor("ORDR_SOURCE").Value = "E"
            tabHeader.Tabs("Audit Trail").Visible = Not (EntryMode = "N")
        Else
            Clear_Record()
        End If

        If ScreenMode Then
            With grdSOTORDR2.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() _
              {"ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", _
               "ORDR_AMT_OPEN", "ORDR_AMT_PICK", "ORDR_AMT_SHIP", "ORDR_AMT_CANC", "ORDR_RELEASE"}
                    .Columns(COLUMN_NAME).Hidden = (EntryMode = "N")
                Next
                '  splOrderDetails.Panel2Collapsed = (EntryMode = "N" Or EntryMode = "E")
                '  .Columns("X").Hidden = InquiryMode Or (EntryMode <> "E")
            End With

            chkManualPrice.Checked = False
            If ScreenMode Then
                Dim rowSOTTCLS1 As DataRow = LookUp("SOTTCLS1", rowSOTORDR1.Item("TRADE_CLASS_CODE") & "")
                chkManualPrice.Visible = (EntryMode = "N" Or EntryMode = "E") _
                 And (rowSOTTCLS1.Item("ALLOW_MANUAL_PRICE") & "" = "1")

            End If
        Else
        End If

        If Trim(ASCMAIN1.USER_CODES) = "FS" Then
            Absx1.chkFor("ORDR_OVERRIDE_NOT_ALLOCATED").Enabled = False
            chkManualPrice.Enabled = False
            Absx1.txtFor("TERM_CODE").Enabled = False
            Absx1.txtFor("SREP_CODE").Enabled = False
            Absx1.txtFor("SREP2_CODE").Enabled = False
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("ORDR_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""

        CUST_CODE = ""
        ORDR_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDR2", "SOTORDR4", "SOTORDR5", "SOTORDRE", "SOTORDXR", _
              "SOTPICK1", "SOTPICK2", "SOTORDRC", "SOTORDRD", "SOTORDRB", "SOTORDRQ", "SOTORDRI", _
             "SOTCART1", "SOTCART2", "SOTWORK1", "SOTWORK2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        multiple_order_maintenance = False


        Load_SOTORDRX()
        If multi_store_is_active Then Setup_MS(False)
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            Init_Record()
        Else
            rowSOTORDR1 = Fill_Record("SOTORDR1", ORDR_NO)
            Fill_Records("SOTORDR2", ORDR_NO)
            Fill_Records("SOTORDR4", ORDR_NO)
            Fill_Records("SOTORDR5", ORDR_NO)

            If multiple_order_maintenance Then
                If EntryMode = "E" Then
                    If multiple_order_type = "ORDR_GROUP_NO" Then
                        Fill_Records("SOTORDRB", ORDR_GROUP_NO)
                    Else
                        ASCMAIN1.sql = sqlSOTORDRB & " and SOTORDR1.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
                        Fill_Records("SOTORDRB", "", True, ASCMAIN1.sql)
                        For Each rowSOTORDRB As DataRow In dst.Tables("SOTORDRB").Select("")
                            Dim ORDR_GROUP_NO As String = rowSOTORDRB.Item("ORDR_GROUP_NO")
                            If Not ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                        Next
                    End If
                    rowSOTORDR0 = Fill_Record("SOTORDR0", ORDR_GROUP_NO)
                    rowSOTORDR0.Item("ORDR_HOLD") = rowSOTORDR1.Item("ORDR_HOLD")
                    rowSOTORDR0.Item("ORDR_ADDR_TYPE_ST") = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
                    rowSOTORDR0.Item("ORDR_OVERRIDE_NOT_ALLOCATED") = rowSOTORDR1.Item("ORDR_OVERRIDE_NOT_ALLOCATED")
                    rowSOTORDR0.Item("REASON_CODE") = rowSOTORDR1.Item("REASON_CODE")
                    rowSOTORDR0.Item("ORDR_HOLD_REASON") = rowSOTORDR1.Item("ORDR_HOLD_REASON")
                    rowSOTORDR0.Item("ORDR_ARRIVAL_DATE") = rowSOTORDR1.Item("ORDR_ARRIVAL_DATE")
                    rowSOTORDR0.Item("ORDR_SPECIAL_INST") = rowSOTORDR1.Item("ORDR_SPECIAL_INST")
                    rowSOTORDR0.Item("ORDR_INV_COMMENT") = rowSOTORDR1.Item("ORDR_INV_COMMENT")
                    rowSOTORDR0.Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
                    rowSOTORDR0.Item("SHIP_VIA_CODE") = rowSOTORDR1.Item("SHIP_VIA_CODE")
                    rowSOTORDR0.Item("FRT_TERMS") = rowSOTORDR1.Item("FRT_TERMS")
                    rowSOTORDR0.Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                    rowSOTORDR0.Item("SREP2_CODE") = rowSOTORDR1.Item("SREP2_CODE")
                    rowSOTORDR0.Item("ORDR_MESSAGE") = rowSOTORDR1.Item("ORDR_MESSAGE")

                    dst.Tables("SOTORDRQ").Rows.Clear()
                    For Each row As DataRow In dst.Tables("SOTORDRB").Select("")
                        Dim ORDR_NO_MS As String = row.Item("ORDR_NO")
                        Fill_Records("SOTORDRQ", ORDR_NO_MS, False)
                    Next

                    grdSOTORDRB.DisplayLayout.Bands(0).Summaries.Clear()
                    If dst.Tables("SOTORDRB").Columns.Contains("QTY_000") Then
                        dst.Tables("SOTORDRB").Columns("QTY_000").Expression = ""
                    End If

                    With dst.Tables("SOTORDRB")
                        For DCOLi As Integer = .Columns.Count - 1 To 0 Step -1
                            If .Columns(DCOLi).ColumnName.StartsWith("QTY_") Then

                                .Columns.Remove(.Columns(DCOLi).ColumnName)
                            End If
                        Next
                    End With

                    Create_Summary(grdSOTORDRB, "CUST_STORE_NO", "Count")

                    grdSOTORDRB.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTORDRB.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText

                    Dim ITEM_CODE As String = ""
                    Dim COL As Integer = 0

                    Dim C As DataColumn = dst.Tables("SOTORDRB").Columns.Add("QTY_" & Format(COL, "000"), GetType(System.Int64))
                    With grdSOTORDRB.DisplayLayout.Bands(0).Columns("QTY_" & Format(COL, "000"))
                        .Header.Caption = "Total"
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Width = 80
                        .Hidden = False
                        .Format = "#,##0"
                        .CellAppearance.BackColor = Drawing.Color.Beige
                        .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                        .Header.Appearance.BackColor = Drawing.Color.White
                        .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        .Header.Appearance.TextHAlign = HAlign.Right
                        .CellAppearance.TextHAlign = HAlign.Right
                        Create_Summary(grdSOTORDRB, "QTY_" & Format(COL, "000"))
                    End With

                    Dim T As String = ""
                    Dim CN As String = ""

                    For Each rowSOTORDRQ As DataRow In dst.Tables("SOTORDRQ").Select("", "ITEM_CODE")
                        Dim ORDR_NO As String = rowSOTORDRQ.Item("ORDR_NO")
                        If rowSOTORDRQ.Item("ITEM_CODE") <> ITEM_CODE Then
                            ITEM_CODE = rowSOTORDRQ.Item("ITEM_CODE")
                            COL += 1

                            Dim rowSOTORDRI As DataRow = dst.Tables("SOTORDRI").NewRow
                            With rowSOTORDRI
                                .Item("ITEM_CODE") = ITEM_CODE
                                .Item("COL") = COL
                                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                                .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                                .Item("ITEM_UOM") = rowICTITEM1.Item("ITEM_UOM")
                                .Item("CARTON_PACK_QTY") = rowICTITEM1.Item("CARTON_PACK_QTY")
                                .Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                                .Item("ORDR_UNIT_PRICE") = rowSOTORDRQ.Item("ORDR_UNIT_PRICE")
                            End With
                            dst.Tables("SOTORDRI").Rows.Add(rowSOTORDRI)

                            CN = "QTY_" & Format(COL, "000")
                            C = dst.Tables("SOTORDRB").Columns.Add(CN, GetType(System.Int64))
                            With grdSOTORDRB.DisplayLayout.Bands(0).Columns(CN)
                                .Header.Caption = ITEM_CODE
                                .Width = 80
                                .Hidden = False
                                .Format = "#,##0"
                                .Header.Appearance.TextHAlign = HAlign.Right
                                .CellAppearance.TextHAlign = HAlign.Right
                                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                                .Header.ToolTipText = ITEM_CODE & vbCrLf & rowICTITEM1.Item("ITEM_DESC") & vbCrLf & "Retail " & Format(Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & ""), "$#,##0.00")
                                Create_Summary(grdSOTORDRB, CN)
                                T &= " + ISNULL(" & CN & ",0)"
                                .Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                                .Header.Appearance.BackColor = Drawing.Color.White
                                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            End With
                        End If
                        Dim rowSOTORDRB As DataRow = dst.Tables("SOTORDRB").Rows.Find(ORDR_NO)
                        rowSOTORDRB.Item(CN) = rowSOTORDRQ.Item("ORDR_QTY_OPEN")
                    Next
                    Setup_Multiple_Order_Grid(False)
                    '  Fill_Records("SOTORDRQ", ORDR_GROUP_NO)
                    dst.Tables("SOTORDRB").Columns("QTY_000").Expression = Mid(T, 4)
                End If

            End If
        End If

        Totals_for_SOTORDRI()

        PRICE_LIST_CODE = rowSOTORDR1.Item("PRICE_LIST_CODE") & ""
        PRICE_CLASS_CODE = rowSOTORDR1.Item("PRICE_CLASS_CODE")
        Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
        PRICE_BASIS = rowSOTPCLS1.Item("PRICE_BASIS") & ""
        PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")

        Fill_Records("SOTORDRE", ORDR_NO)
        Sort_grdColumns(grdSOTORDRE, "INIT_DATE".ToLower)
        Fill_Records("SOTORDXR", ORDR_NO)
        Sort_grdColumns(grdSOTORDXR, "INIT_DATE".ToLower)

        Fill_Records("SOTWORK1", ORDR_NO)
        Fill_Records("SOTWORK2", ORDR_NO)

        CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
        CUST_STORE_NO = rowSOTORDR1.Item("CUST_STORE_NO") & ""
        CUST_DC_NO = rowSOTORDR1.Item("CUST_DC_NO") & ""
        ORDR_GROUP_NO = rowSOTORDR1.Item("ORDR_GROUP_NO") & ""
        ORDR_GROUP_NOs.Clear()
        ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)
        CUST_BILL_TO_CUST = rowSOTORDR1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST <> CUST_CODE Then
            rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
        Else
            rowARTCUST1_BT = rowARTCUST1
        End If

        Absx1.txtFor("CUST_ROUTING_INST").Text = rowARTCUST1.Item("CUST_ROUTING_INST") & ""

        'rowARTCUST2 = Fill_Record("ARTCUST2", New String() {CUST_CODE, "MK", CUST_STORE_NO})

        ' ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "' and NVL(CUST_DC_IND,'0') <> '1'"
        ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "'" ' TARGET EDI ORDERS MAINT - SO WHAT IF WE THINK OF DC AS A STORE? MK ADDRESS NOT LOADING (WITH DC AS A STORE) IF I DON'T MAKE THIS CHANGE

        If CUST_CODE = "CONSUMER" Then
            ASCMAIN1.sql &= " AND CUST_STORE_NO = '" & CUST_STORE_NO & "'"
        End If
        Fill_Records("ARTCUST2", "", True, ASCMAIN1.sql)
        rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})

        'ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "' and CUST_ADDR_TYPE = 'DC'"
        'Fill_Records("ARTCUST2", "", False, ASCMAIN1.sql)

        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
        ' Sort_grdColumns(grdSOTORDR3, "ORDR_DIV_CODE")
        Sort_grdColumns(grdSOTORDR4, "ORDR_CLNO")

        If EntryMode = "N" Then
            lblStatus.Text = "New Order"
        Else
            Select Case rowSOTORDR1.Item("ORDR_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "P"
                    lblStatus.Text = "In Pick"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
                Case "F"
                    lblStatus.Text = "Shipped"
            End Select
        End If

        sub_grid = ""

        'CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
        'If CURR_CODE = "" Then
        '    CURR_CODE = "USD"
        '    CURR_EXCH_RATE = 1
        'Else
        '    If rowARTCUST1.Item("CURR_CODE").Value = ROWs("ARTPARM1").Item("AR_PARM_CURR_CODE") Then
        '        CURR_CODE = "USD"
        '        CURR_EXCH_RATE = 1
        '    Else
        '        OraD.Parameters("CODE").Value = rowARTCUST1.Item("CURR_CODE")
        '        dynICTCURR1.Refresh()
        '        CURR_CODE = dynARTCUST1.ITEM("CURR_CODE")
        '        CURR_EXCH_RATE = dynICTCURR1.ITEM("CURR_EXCH_CUR")
        '    End If
        'End If

        'If EntryMode = "N" Then
        '    Absx1.optFor("ORDR_ADDR_TYPE_ST").Value = "DC"
        'End If

        CUST_STORE_NOs_multi_store.Clear()

        If InquiryMode Or EntryMode = "V" Then
            Fill_Records("SOTPICK1", ORDR_NO)
            Fill_Records("SOTPICK2", ORDR_NO)
            'Set_SOTPICK2()
            Set_SOTCART1()
        End If

        With grdSOTORDR2.DisplayLayout.Bands(0)
            If (EntryMode = "E" Or EntryMode = "N") And Val(rowSOTORDR1.Item("ORDR_PICK_SEQ") & "") = 0 Then
                .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                If EntryMode = "E" Then
                    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Else
                .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTORDR2}
            If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    .AllowUpdate = DefaultableBoolean.True
                    .AllowDelete = DefaultableBoolean.True
                End With
            Else
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowUpdate = DefaultableBoolean.False
                    .AllowDelete = DefaultableBoolean.False
                End With
            End If
        Next

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, False)
            'frmOtherInfo.Enabled = True
        Else
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, True)
            'frmOtherInfo.Enabled = false
        End If

        If grdSOTORDR2.Rows.Count = 0 Then
        Else
            If grdSOTORDR2.ActiveRow Is Nothing Then
                grdSOTORDR2.ActiveRow = grdSOTORDR2.Rows(0)
            End If
        End If

        If EntryMode = "E" Or EntryMode = "N" Then
            If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                Set_Read_Only_for_ctl(Absx1.txtFor("EDI_APPOINTMENT"), True)
                'frmShipToOption.Enabled = False
                Set_Read_Only_for_ctl(optShipTo, True)
            Else
                grdSOTORDR2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                Set_Read_Only_for_ctl(Absx1.txtFor("EDI_APPOINTMENT"), False)
                'frmShipToOption.Enabled = True
                Set_Read_Only_for_ctl(optShipTo, False)
            End If
        Else
            Set_Read_Only_for_ctl(Absx1.txtFor("EDI_APPOINTMENT"), True)
        End If

        Dim rowSOTORDR5 As DataRow = Nothing
        Dim CUST_ADDR_CODEs() As String = {"BY", "MK", "DC"}
        Dim CUST_ADDR_CODE As String = ""
        If EntryMode = "N" Then
            CUST_ADDR_CODEs = {"BY", "BT", "MK", "DC", "ST"}
        End If
        For Each CUST_ADDR_TYPE As String In CUST_ADDR_CODEs
            Dim row As DataRow = Nothing
            If CUST_ADDR_TYPE = "BY" Then
                row = rowARTCUST1
                CUST_ADDR_CODE = CUST_CODE
            ElseIf CUST_ADDR_TYPE = "BT" Then
                row = rowARTCUST1_BT
                CUST_ADDR_CODE = CUST_BILL_TO_CUST
            ElseIf CUST_ADDR_TYPE = "MK" Then
                row = rowARTCUST2
                CUST_ADDR_CODE = CUST_STORE_NO
            ElseIf CUST_ADDR_TYPE = "DC" Then
                row = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_DC_NO}, True)
                CUST_ADDR_CODE = CUST_DC_NO
            ElseIf CUST_ADDR_TYPE = "ST" Then
                Dim ORDR_ADDR_TYPE_ST As String = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST")
                CUST_ADDR_CODE = IIf(ORDR_ADDR_TYPE_ST = "DC", CUST_DC_NO, CUST_STORE_NO)
                row = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_ADDR_CODE}, True)
            End If
            rowSOTORDR5 = dst.Tables("SOTORDR5").NewRow
            With rowSOTORDR5
                .Item("ORDR_NO") = ORDR_NO
                .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
                .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
                If row IsNot Nothing Then
                    For Each COLUMN_NAME As String In New String() _
                        {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", _
                         "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}
                        Dim COLUMN_NAME_row As String = COLUMN_NAME
                        If CUST_ADDR_TYPE = "DC" Or CUST_ADDR_TYPE = "ST" Or CUST_ADDR_TYPE = "MK" Then
                            COLUMN_NAME_row = Replace(COLUMN_NAME, "CUST_", "CUST_STORE_")
                        End If
                        .Item(COLUMN_NAME) = row.Item(COLUMN_NAME_row)
                    Next
                End If
            End With
            dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
        Next
        txtCUST_ADDR_CODE_BT.Text = "000000"

        If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
            TOTAL_ORDR_AMT = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")
            'ASCMAIN1.sql = "Select ALLOW_CHANGE_RANGE from EDTTRPM1" _
            ' & " WHERE EDI_DOC_NO = '850'" _
            ' & " AND CUST_CODE = '" & CUST_CODE & "'"
            'ALLOW_CHANGE_RANGE = ASCDATA1.GetDataValue
        Else
            TOTAL_ORDR_AMT = 0
            'ALLOW_CHANGE_RANGE = ""
        End If

        Display_Totals()

        ' tabMain.SelectedTab = tabMain.Tabs("Names && Addresses")
        tabMain.SelectedTab = tabMain.Tabs("Order Header Info")

        EnforceConstraints(True)

        ASCMAIN1.sql = "Select Max (REV_NO) from SOTORDXR where ORDR_NO = '" & ORDR_NO & "'"
        REV_NO = Val(ASCDATA1.GetDataValue & "")
        If EntryMode = "N" Then
            lblORDR_NO.Text = "Order No - New"
            Record_Event("INIT", "Sales Order Entry Started")
        ElseIf EntryMode = "E" Then
            lblORDR_NO.Text = "Order No - Rev#" & CStr(REV_NO + 1)
            Record_Event("LAST", "Sales Order Edit Started")
        Else
            lblORDR_NO.Text = "Order No - Rev#" & CStr(REV_NO)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Init_Record()
        ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
        ORDR_GROUP_NO = ""
        ORDR_GROUP_NOs.Clear()

        rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
        With rowSOTORDR1
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_STORE_NO") = CUST_STORE_NO
            .Item("ORDR_CUST_PO") = ORDR_CUST_PO
            .Item("ORDR_DATE") = DATETIME_STAMP.Date
            .Item("ORDR_SOURCE") = "K"
            .Item("ORDR_STATUS") = "O"
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
            .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
            .Item("ORDR_TYPE_CODE") = "REG"

            Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
            If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
            If WHSE_CODE = "" Then WHSE_CODE = ""
            .Item("WHSE_CODE") = WHSE_CODE

            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

            If rowARTCUST2.Item("CUST_DC_IND") & "" = "1" Then
                .Item("ORDR_ADDR_TYPE_ST") = "DC"
                CUST_DC_NO = CUST_STORE_NO
            Else
                .Item("ORDR_ADDR_TYPE_ST") = "MK"
                CUST_DC_NO = rowARTCUST2.Item("CUST_DC_NO") & ""
                If CUST_DC_NO <> "" Then
                    Dim rowARTCUST2_DC As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_DC_NO})
                    If rowARTCUST2_DC IsNot Nothing AndAlso rowARTCUST2_DC.Item("CUST_DC_IND") & "" = "1" Then
                        .Item("ORDR_ADDR_TYPE_ST") = "DC"
                    Else
                        CUST_DC_NO = ""
                    End If
                End If
            End If
            .Item("CUST_DC_NO") = CUST_DC_NO
            .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date

            ' Sold To
            .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
            .Item("SREP_CODE") = SREP_CODE
            .Item("SREP2_CODE") = SREP2_CODE
            .Item("ORDR_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            .Item("EVENT_CODE") = rowARTCUST1.Item("EVENT_CODE") & ""

            .Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS") & ""
            .Item("ORDR_SPECIAL_INST") = rowARTCUST1.Item("CUST_SPECIAL_INST") & ""
            .Item("ORDR_INV_COMMENT") = rowARTCUST1.Item("CUST_INV_COMMENT") & ""
            .Item("SHIP_VIA_CODE") = rowARTCUST1.Item("SHIP_VIA_CODE") & ""

            .Item("CURR_CODE") = "USD"
            .Item("CURR_EXCH_RATE") = 1

            .Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""
            .Item("PRICE_CLASS_CODE") = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
            .Item("PRICE_LIST_CODE") = rowARTCUST1.Item("PRICE_LIST_CODE") & ""
            .Item("SALES_DIVISION_CODE") = "AV"

            ' Bill To
            If CUST_BILL_TO_CUST <> CUST_CODE Then
                rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
            Else
                rowARTCUST1_BT = rowARTCUST1
            End If
            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE") & ""
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE") & ""
            .Item("CUST_FACTOR_IND") = rowARTCUST1_BT.Item("CUST_FACTOR_IND") & ""

            ' Store
            If rowARTCUST2 IsNot Nothing Then
                .Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION") & ""
            End If

            ' Use Values from Previously Entered Order
            ' THIS CAUSED 2 CONSUMER ORDERS WITH THE SAME PO TO GET GROUPED - NO GOOD.  THATS WHY 1<>1
            If rowSOTORDR1po IsNot Nothing And 1 <> 1 Then
                .Item("ORDR_GROUP_NO") = rowSOTORDR1po.Item("ORDR_GROUP_NO")
                .Item("ORDR_SHIP_DATE") = rowSOTORDR1po.Item("ORDR_SHIP_DATE")
                .Item("ORDR_CANCEL_DATE") = rowSOTORDR1po.Item("ORDR_CANCEL_DATE")
                .Item("ORDR_DATE") = rowSOTORDR1po.Item("ORDR_DATE")
                .Item("ORDR_DEPT") = rowSOTORDR1po.Item("ORDR_DEPT")
                .Item("ORDR_SPECIAL_INST") = rowSOTORDR1po.Item("ORDR_SPECIAL_INST")
                .Item("FRT_TERMS") = rowSOTORDR1po.Item("FRT_TERMS")
            End If

            'If ASCMAIN1.Running_in_VS Then Stop ' Use Copy-From Order Values for SREP_CODE, SREP2_CODE, SHIP_VIA_CODE

            If ASCDATA1.GetDataValue("Select MIN(ORDR_NO) from SOTORDR1 where CUST_CODE = '" & CUST_CODE & "'") = "" Then
                .Item("ORDR_INITIAL") = "1"
            End If
        End With
        dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

        If ORDR_NO_to_copy <> "" Then
            For Each TABLE_NAME As String In New String() _
                    {"SOTORDR2", "SOTORDR3", "SOTORDR4", "SOTORDR5", "SOTORDR9"}
                Fill_Records("TABLE_NAME", ORDR_NO_to_copy)
                For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
                    row.Item("ORDR_NO") = ORDR_NO
                    If TABLE_NAME = "SOTORDR2" Then
                        row.Item("ORDR_QTY_ALLO") = 0
                        row.Item("ORDR_QTY_OPEN") = 0
                        row.Item("ORDR_QTY_PICK") = 0
                        row.Item("ORDR_QTY_SHIP") = 0
                        row.Item("ORDR_QTY_CANC") = 0
                        row.Item("ORDR_STATUS") = "O"
                        row.Item("ORDR_QTY_PRE_ALLO") = 0
                    End If
                Next
            Next
        End If
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"SOTORDR1", "SOTORDR2", "SOTORDR4", "SOTORDR5"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where ORDR_NO = '" & ORDR_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record_multiple_order_group()

        BeginTrans()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        ' Re-Evaluate Order Status after making changes

        For Each rowSOTORDRB As DataRow In dst.Tables("SOTORDRB").Select("")
            Dim ORDR_NO As String = rowSOTORDRB.Item("ORDR_NO")
            TAC.SOCMAIN1.Record_Event_SOTORDRE(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "Change", "Multi-Order")
            ' Record_Event("UPDT", "Sales Order Updated")

            Dim changes As Boolean = False

            Fill_Records("SOTORDR2", ORDR_NO)
            For Each ROWSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                Dim ITEM_CODE As String = ROWSOTORDR2.Item("ITEM_CODE")
                Dim rowSOTORDRI As DataRow = dst.Tables("SOTORDRI").Rows.Find(ITEM_CODE)
                Dim ORDR_UNIT_PRICE As Decimal = Val(ROWSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                Dim ORDR_UNIT_PRICE_NEW As Decimal = Val(rowSOTORDRI.Item("ORDR_UNIT_PRICE") & "")
                Dim COLI As Integer = Val(rowSOTORDRI.Item("COL") & "")
                Dim ORDR_QTY As Int64 = Val(ROWSOTORDR2.Item("ORDR_QTY") & "")
                Dim ORDR_QTY_OPEN As Int64 = Val(ROWSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                Dim ORDR_QTY_CANC As Int64 = Val(ROWSOTORDR2.Item("ORDR_QTY_CANC") & "")
                Dim ORDR_QTY_OPEN_NEW As Int64 = Val(rowSOTORDRB.Item("QTY_" & Format(COLI, "000")) & "")
                If ORDR_QTY_OPEN_NEW <> ORDR_QTY_OPEN Or ORDR_UNIT_PRICE <> ORDR_UNIT_PRICE_NEW Then
                    'ORDR_QTY_CANC = ORDR_QTY_CANC + (ORDR_QTY_OPEN - QTY)
                    ORDR_QTY_CANC = ORDR_QTY - ORDR_QTY_OPEN_NEW
                    ORDR_QTY_OPEN = ORDR_QTY_OPEN_NEW
                    If ORDR_QTY_CANC < 0 Then ORDR_QTY_CANC = 0
                    ROWSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY_OPEN
                    ROWSOTORDR2.Item("ORDR_QTY_CANC") = ORDR_QTY_CANC
                    ROWSOTORDR2.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_NEW
                    ROWSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_NEW
                    changes = True
                End If
            Next
            If changes Then
                Dependent_Updates(-1, ORDR_NO)
                Update_Record_TDA("SOTORDR2")
                Dependent_Updates(1, ORDR_NO)
            End If
        Next

        Dim ORDR_SHIP_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_SHIP_DATE").Value
        Dim ORDR_CANCEL_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_CANCEL_DATE").Value
        Dim ORDR_ARRIVAL_DATE As Date = Absx1.dteFor("SOTORDR0.ORDR_ARRIVAL_DATE").Value
        Dim ORDR_OVERRIDE_NOT_ALLOCATED As String = IIf(Absx1.chkFor("SOTORDR0.ORDR_OVERRIDE_NOT_ALLOCATED").Checked, "1", "0")

        If Format(ORDR_ARRIVAL_DATE, "MM/dd/yyyy") = "01/01/0001" Then
            ORDR_ARRIVAL_DATE = Nothing
        End If

        Dim ORDR_DEPT As String = Absx1.txtFor("SOTORDR0.ORDR_DEPT").Text
        Dim ORDR_HOLD As String = IIf(Absx1.chkFor("SOTORDR0.ORDR_HOLD").Checked, "1", "0")
        Dim ORDR_HOLD_REASON As String = Absx1.txtFor("SOTORDR0.ORDR_HOLD_REASON").Text

        'ORDR_ADDR_TYPE_ST
        Dim ORDR_SPECIAL_INST As String = Absx1.txtFor("SOTORDR0.ORDR_SPECIAL_INST").Text
        Dim ORDR_INV_COMMENT As String = Absx1.txtFor("SOTORDR0.ORDR_INV_COMMENT").Text
        Dim ORDR_MESSAGE As String = Absx1.txtFor("SOTORDR0.ORDR_MESSAGE").Text

        Dim SREP_CODE As String = Absx1.txtFor("SOTORDR0.SREP_CODE").Text
        Dim SREP2_CODE As String = Absx1.txtFor("SOTORDR0.SREP2_CODE").Text
        Dim TERM_CODE As String = Absx1.txtFor("SOTORDR0.TERM_CODE").Text
        Dim WHSE_CODE As String = Absx1.txtFor("SOTORDR0.WHSE_CODE").Text
        Dim SHIP_VIA_CODE As String = Absx1.txtFor("SOTORDR0.SHIP_VIA_CODE").Text
        Dim FRT_TERMS As String = Absx1.txtFor("SOTORDR0.FRT_TERMS").Text
        Dim REASON_CODE As String = Absx1.txtFor("SOTORDR0.REASON_CODE").Text

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCMAIN1.sql = "Update SOTORDR1 " & vbCrLf _
                & " Set ORDR_SHIP_DATE = :PARM1, ORDR_CANCEL_DATE = :PARM2, ORDR_ARRIVAL_DATE = :PARM3, ORDR_MESSAGE = :PARM4" & vbCrLf _
                & ", ORDR_DEPT = :PARM5, ORDR_OVERRIDE_NOT_ALLOCATED = :PARM6, ORDR_HOLD = :PARM7, ORDR_HOLD_REASON = :PARM8" & vbCrLf _
                & ", ORDR_SPECIAL_INST = :PARM9, ORDR_INV_COMMENT = :PARM10" & vbCrLf _
                & ", SREP_CODE = :PARM11, SREP2_CODE = :PARM12, TERM_CODE = :PARM13, WHSE_CODE = :PARM14" & vbCrLf _
                & ", SHIP_VIA_CODE = :PARM15, FRT_TERMS = :PARM16, REASON_CODE = :PARM17" & vbCrLf _
                & ", LAST_DATE = :PARM18, LAST_OPER = :PARM19" & vbCrLf _
                & " where ORDR_GROUP_NO = :PARM20"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DDDVVVVVVVVVVVVVVDVV", _
                                New Object() {ORDR_SHIP_DATE, ORDR_CANCEL_DATE, ORDR_ARRIVAL_DATE, ORDR_MESSAGE, _
                                              ORDR_DEPT, ORDR_OVERRIDE_NOT_ALLOCATED, ORDR_HOLD, ORDR_HOLD_REASON, _
                                              ORDR_SPECIAL_INST, ORDR_INV_COMMENT, _
                                              SREP_CODE, SREP2_CODE, TERM_CODE, WHSE_CODE, _
                                              SHIP_VIA_CODE, FRT_TERMS, REASON_CODE, _
                                              DATETIME_STAMP, ASCMAIN1.USER_ID, ORDR_GROUP_NO})

            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        Next


        CommitTrans("Update Complete")
    End Sub

    Sub Update_Record()

        REV_NO += 1
        REV_LNO = 0
        dst.Tables("SOTORDXR").Rows.Clear()

        BeginTrans()

        ' If ASCMAIN1.Running_in_VS Then Stop ' WORKING WITH ROWSOTORDR1 BELOW - ARE WE SYNCHED?

        Me.Cursor = Cursors.WaitCursor

        restore_reservation = True

        Dim ALL_ORDERS As New List(Of String)
        ALL_ORDERS.Add(ORDR_NO)

        Dim ORDR_NO_ORIG As String = ORDR_NO

        ASCMAIN1.Progress("Now Updating ...")

        If rowSOTORDR1.Item("ORDR_GROUP_NO") & "" = "" Then
            ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
            rowSOTORDR1.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            ORDR_GROUP_NOs.Clear()
            ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
        End If

        ' Load up SOTORDR5 with Bill-To and Ship-To Address

        If chkEdit.Checked Then
            Update_Ship_To_Address() ' (REV_NO, REV_LNO)
        End If

        If EntryMode = "N" Or EntryMode = "E" Then
            'ASCDATA1.DeleteRows("SOTORDR5", "CUST_ADDR_CODE <> 'BT' and CUST_ADDR_CODE <> 'ST'")
            dst.Tables("SOTORDR5").Rows.Clear()
            Update_SOTORDR5(ORDR_NO, "BT", txtCUST_ADDR_CODE_BT.Text, "")
            Update_SOTORDR5(ORDR_NO, rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & "", CUST_DC_NO, CUST_STORE_NO)
        End If


        ' Set ORDR_STATUS and ORDR_QTY_ORIG

        Dim SALES_DIVISION_CODE As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            rowSOTORDR2.Item("ORDR_STATUS") = rowSOTORDR1.Item("ORDR_STATUS")
            If EntryMode = "N" Then
                rowSOTORDR2.Item("ORDR_QTY_ORIG") = rowSOTORDR2.Item("ORDR_QTY")
                If SALES_DIVISION_CODE = "" Then
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", rowSOTORDR2.Item("ITEM_CODE"))
                    SALES_DIVISION_CODE = rowICTITEM1.Item("SALES_DIVISION_CODE") & ""
                End If
            End If
        Next

        ' Double-Check SALES_DIVISION_CODE

        If EntryMode = "N" Then
            If SALES_DIVISION_CODE <> rowSOTORDR1.Item("SALES_DIVISION_CODE") & "" And SALES_DIVISION_CODE <> "" Then
                rowSOTORDR1.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
            End If

            rowSOTORDR1.Item("ORDR_ORIG_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
            rowSOTORDR1.Item("ORDR_ORIG_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
        End If

        ' Copy Order to Multiple Stores if Multi-Store Mode is True

        If multi_store_is_active Then
            Clear_Zeroes()

            ' Traverse the Multi-Store grid, and Write ORDR1/2'S for each Store

            For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                Dim CUST_STORE_NO As String = rowSOTORDRS.Item("CUST_STORE_NO")
                Dim ORDR_CUST_PO As String = rowSOTORDRS.Item("ORDR_CUST_PO")

                If CUST_STORE_NO = Me.CUST_STORE_NO Then
                    rowSOTORDR1.Item("ORDR_CUST_PO") = ORDR_CUST_PO
                    'Sql = "Update SOWORDR1 set ORDR_CUST_PO = '" & ORDR_CUST_PO & "'"
                    'Sql = Sql & " where ORDR_NO = '" & ORDR_NO_ORIG & "'"
                    'AccD.Execute(Sql)
                Else
                    Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                    ALL_ORDERS.Add(ORDR_NO)
                    ASCMAIN1.Progress("-", ORDR_NO)
                    Dim row As DataRow = dst.Tables("SOTORDR1").NewRow
                    row.ItemArray = rowSOTORDR1.ItemArray
                    row.Item("ORDR_NO") = ORDR_NO
                    dst.Tables("SOTORDR1").Rows.Add(row)

                    Dim rowARTCUST2_MS As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO}, True)
                    Dim CUST_DC_NO As String = rowARTCUST2_MS.Item("CUST_DC_NO") & ""
                    row.Item("ORDR_NO") = ORDR_NO
                    row.Item("CUST_STORE_NO") = CUST_STORE_NO
                    row.Item("ORDR_CUST_PO") = ORDR_CUST_PO
                    row.Item("CUST_DC_NO") = CUST_DC_NO
                    row.Item("CUST_STORE_LOCATION") = rowARTCUST2_MS.Item("CUST_STORE_LOCATION")

                    Dim SREP_CODE As String = rowARTCUST2_MS.Item("SREP_CODE") & String.Empty
                    If SREP_CODE.Length = 0 Then SREP_CODE = rowARTCUST1.Item("SREP_CODE") & String.Empty
                    row.Item("SREP_CODE") = SREP_CODE
                    row.Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE") & String.Empty

                    Update_SOTORDR5(ORDR_NO, row.Item("ORDR_ADDR_TYPE_ST") & "", CUST_DC_NO, CUST_STORE_NO)

                    Dim ORDR_QTY_orig_store As Int64

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & Me.ORDR_NO & "'")
                        Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim ORDR_QTY As Int64 = Val(rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) & "")
                        If ORDR_QTY <> 0 Then
                            Dim row2 As DataRow = dst.Tables("SOTORDR2").NewRow
                            row2.ItemArray = rowSOTORDR2.ItemArray
                            ORDR_QTY_orig_store = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                            row2.Item("ORDR_NO") = ORDR_NO
                            row2.Item("ORDR_QTY") = ORDR_QTY
                            row2.Item("ORDR_QTY_OPEN") = ORDR_QTY
                            row2.Item("ORDR_QTY_ORIG") = ORDR_QTY
                            dst.Tables("SOTORDR2").Rows.Add(row2)
                        End If
                    Next

                End If
            Next

            Copy_Records(ALL_ORDERS)
        End If


        ' Remove all Order Details where no qty ordered - SOTORDR2 & SOTORDR3

        ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), "ISNULL(ORDR_QTY,0) = 0 and ISNULL(ORDR_QTY_OPEN,0) = 0 and ISNULL(ORDR_QTY_SHIP,0) = 0 and ISNULL(ORDR_QTY_PICK,0) = 0 and ISNULL(ORDR_QTY_CANC,0) = 0")

        ' Update all Currency Fields

        If "USD" <> "USD" Then
            Stop
            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                rowSOTORDR1.Item("CURR_CODE") = "USD" ' CURR_CODE
                rowSOTORDR1.Item("CURR_EXCH_RATE") = 1 ' CURR_EXCH_RATE
            Next
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2.Item("ORDR_UNIT_PRICE") / 1 ' CURR_EXCH_RATE
            Next
        End If


        If ALL_ORDERS.Count <> 0 Then
            ' BeginTrans()
            Record_Event("UPDT", "Sales Order Updated")
            If EntryMode = "E" Then Check_Changed_Fields()
            If EntryMode <> "N" Then Delete_Records()
            Dim SQLD As String = "ORDR_NO = '" & ORDR_NO & "'"
            INIT_LAST("SOTORDR1", False, , True)
            Update_Record_TDA("SOTORDR1", SQLD)
            Update_Record_TDA("SOTORDR2", SQLD)
            Update_Record_TDA("SOTORDR3", SQLD)
            Update_Record_TDA("SOTORDR4", SQLD)
            Update_Record_TDA("SOTORDR5", SQLD)
            Update_Record_TDA("SOTORDR9", SQLD)

            Update_Record_TDA("SOTORDXR")
            Update_Record_TDA("SOTORDRE")

            Update_Record_TDA("SOTWORK1")
            Update_Record_TDA("SOTWORK2")

            For Each ORDR_NOx As String In ALL_ORDERS
                Dependent_Updates(1, ORDR_NOx)
            Next

            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})

            '  CommitTrans()
        End If

        CommitTrans("Update Complete")
    End Sub

    Sub Copy_Records(ORDR_NOs_to_copy_to As List(Of String))

        For Each TABLE_NAME As String In New String() {"SOTORDR4", "SOTORDR5"}
            '  Sql = "Select * from " & z & " where ORDR_NO = '" & Mid$(ALL_ORDERS, 1, 10) & "'"

            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "'"
            If TABLE_NAME = "SOTORDR5" Then sqlw &= " and CUST_ADDR_TYPE = 'BT'"
            For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlw)
                For Each ORDR_NO_to_copy_to As String In ORDR_NOs_to_copy_to
                    Dim row2 As DataRow = dst.Tables(TABLE_NAME).NewRow
                    row2.ItemArray = row.ItemArray
                    row2.Item("ORDR_NO") = ORDR_NO_to_copy_to
                    If dst.Tables(TABLE_NAME).Rows.Find(New String() {row2.Item("ORDR_NO"), row2.Item("CUST_ADDR_TYPE")}) Is Nothing Then
                        dst.Tables(TABLE_NAME).Rows.Add(row2)
                    End If
                Next
            Next
        Next
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "CUST_CODE"

                If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                        sql_where = " and (ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')))"
                    Else
                        sql_where = " and (ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'))"
                    End If
                End If

            Case "CUST_STORE_NO"
                sql_where = "CUST_STORE_STATUS = 'A'"

            Case "SOTORDR5_BT.CUST_ADDR_CODE"
                If ASCMAIN1.Running_in_VS Then Stop
                Stop
                ' Sql = "Select CUST_ADDR_CODE, CUST_NAME FROM ARTCUST2 WHERE CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                ' Sql = Sql & " AND CUST_ADDR_TYPE = 'BT'"


            Case "ORDR_NO", "INV_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                If Absx1.txtFor("CUST_STORE_NO").Text <> "" And Absx1.txtFor("CUST_CODE").Text = "" Then
                    MsgBox("You must enter a Customer Code", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""


                If COLUMN_NAME = "ORDR_NO" Then
                    If InquiryMode Then
                    Else
                        sql_where &= " and SOTORDR1.ORDR_STATUS = 'O' "
                    End If
                End If
                If COLUMN_NAME = "INV_NO" Then
                    sql_where &= " and SOTORDR1.ORDR_STATUS = 'F' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTORDR1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("CUST_STORE_NO").Text <> "" Then
                    sql_where &= " and SOTORDR1.CUST_STORE_NO = '" & Absx1.txtFor("CUST_STORE_NO").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTORDR1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If

                'If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                '    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                '        sql_where = " and ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')"
                '    Else
                '        sql_where = " and ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'"
                '    End If
                'End If
                If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                    If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                        sql_where = " and (ARTCUST1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')))"
                    Else
                        sql_where = " and (ARTCUST1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'" _
                            & " or ARTCUST1.CUST_CODE in (Select Distinct CUST_CODE from ARTCUST2 where ARTCUST2.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'))"
                    End If
                End If

            Case "SHIP_VIA_CODE"
                If Not InquiryMode Then
                    sql_where &= "SHIP_VIA_STATUS = 'A'"
                End If
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("ORDR_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTORDR1"
            E.COLUMN_NAME = "ORDR_NO"
            E.CODE_VALUE = Absx1.txtFor("ORDR_NO").Text
            E.DESC_VALUE = "Sales Order"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDRX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Refresh")
        Load_Popup_Menu(grdSOTORDR2, "BB", "Item Status Inquiry", "Load Order Form")
        Load_Popup_Menu(grdSOTORDRS, "BB", "Set Customer PO to Value in Header", "Update Qty to All Stores")
        Load_Popup_Menu(grdSOTORDXR, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTCART1, "B", "Track Shipment")
        Load_Popup_Menu(grdSOTORDRB, "S", "Show Additional Header Fields")
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
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

                Case "grdSOTORDRS"
                    Dim show_qty_copy_option As Boolean = False
                    Dim ORDR_QTY As Int64 = 0
                    If grdSOTORDRS.ActiveCell IsNot Nothing Then
                        Dim COLUMN_NAME As String = grdSOTORDRS.ActiveCell.Column.Key
                        If COLUMN_NAME.StartsWith("QTY_") Then
                            show_qty_copy_option = True
                            ORDR_QTY = Val(grdSOTORDRS.ActiveCell.Value & "")
                        End If
                    End If
                    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = show_qty_copy_option
                    tlb_btn.SharedProps.Caption = "Update Qty to " & CStr(ORDR_QTY) & " for All Stores"

                Case "grdSOTORDR2"
                    tlb_btn = DirectCast(tlb.Tools("Load Order Form"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "N")


            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Set Customer PO to Value in Header"
                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                    rowSOTORDRS.Item("ORDR_CUST_PO") = Absx1.txtFor("ORDR_CUST_PO").Text
                Next

            Case "Show UPC/SKU"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Tag <> "X" Then
                    grdSOTORDR2.DisplayLayout.Bands(0).Columns("CUST_UPC").Hidden = Not tlb_sbt.Checked
                    grdSOTORDR2.DisplayLayout.Bands(0).Columns("CUST_SKU").Hidden = Not tlb_sbt.Checked
                End If


            Case "Load Order Form"
                Load_Order_Form()

            Case "Refresh"
                Load_SOTORDRX()

            Case "Show Additional Header Fields"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Setup_Multiple_Order_Grid(tlb_sbt.Checked)

            Case "Multiple Order Maintenance"
                If grdSOTORDRX.Selected.Rows.Count <> 0 Then
                    Dim ORDR_SHIP_DATE As Date = Nothing
                    Dim ORDR_CANCEL_DATE As Date = Nothing
                    Dim ORDR_HOLD As String = ""
                    Dim ORDR_OVERRIDE_NOT_ALLOCATED As String = ""
                    ORDR_NOs_to_maintain.Clear()
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRX.Selected.Rows
                        Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then
                            ORDR_NOs_to_maintain.Clear()
                            Exit For
                        Else
                            ORDR_NOs_to_maintain.Add(ORDR_NO)
                        End If
                        If grow.Cells("CUST_CODE").Value <> Absx1.txtFor("CUST_CODE").Text Then
                            MsgBox("Customer Code of each order selected" & vbCrLf & " must be the same as Customer Entered Above (" & Absx1.txtFor("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Proceed with Multiple Order Maintenance")
                            ORDR_NOs_to_maintain.Clear()
                            Exit For
                        End If
                        If ORDR_HOLD = "" Then
                            ORDR_HOLD = IIf(grow.Cells("ORDR_HOLD").Value & "" = "1", "1", "0")
                            ORDR_OVERRIDE_NOT_ALLOCATED = IIf(grow.Cells("ORDR_OVERRIDE_NOT_ALLOCATED").Value & "" = "1", "1", "0")
                            ORDR_SHIP_DATE = grow.Cells("ORDR_SHIP_DATE").Value
                            ORDR_CANCEL_DATE = grow.Cells("ORDR_CANCEL_DATE").Value
                        Else
                            If ORDR_HOLD <> IIf(grow.Cells("ORDR_HOLD").Value & "" = "1", "1", "0") _
                            Or ORDR_OVERRIDE_NOT_ALLOCATED <> IIf(grow.Cells("ORDR_OVERRIDE_NOT_ALLOCATED").Value & "" = "1", "1", "0") _
                            Or ORDR_SHIP_DATE <> grow.Cells("ORDR_SHIP_DATE").Value _
                            Or ORDR_CANCEL_DATE <> grow.Cells("ORDR_CANCEL_DATE").Value Then
                                MsgBox("On Hold Status, Allocation, and Shipping Window of all orders selected must be the same", MsgBoxStyle.OkOnly, "Cannot Proceed with Multiple Order Maintenance")
                                ORDR_NOs_to_maintain.Clear()
                                Exit For
                            End If
                        End If
                    Next

                    If ORDR_NOs_to_maintain.Count = 0 Then
                        ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    Else
                        Absx1.txtFor("ORDR_NO").Text = ORDR_NOs_to_maintain(0)
                        Click_Command("Edit")
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Sales Order Inquiry"
                Dim SO_ORDER_NO As String = grd.ActiveRow.Cells("SO_ORDER_NO").Text
                Dim rowSOTINVH1 As DataRow = LookUp("SOTINVH1", SO_ORDER_NO)
                If rowSOTINVH1 IsNot Nothing Then
                    Context_Launch("View", SO_ORDER_NO, e.Tool.Key, "SOFORDRI")
                End If

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Sub Selected Stores"

            Case "Add Line to Selected Stores"

            Case "Update Qty to All Stores"
                Dim ORDR_QTY As Int64 = Val(grdSOTORDRS.ActiveCell.Value & "")
                Dim COLUMN_NAME As String = grdSOTORDRS.ActiveCell.Column.Key
                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                    rowSOTORDRS.Item(COLUMN_NAME) = ORDR_QTY
                Next

            Case "Track Shipment"
                Dim CART_TRACKING_NO As String = ""
                Dim PICK_NO As String = ""
                If grd.ActiveRow.Band.Key = "SOTCART1" Then
                    CART_TRACKING_NO = grd.ActiveRow.Cells("CART_TRACKING_NO").Value & ""
                    PICK_NO = grd.ActiveRow.Cells("PICK_NO").Value & ""
                Else
                    CART_TRACKING_NO = grd.ActiveRow.ParentRow.Cells("CART_TRACKING_NO").Value & ""
                    PICK_NO = grd.ActiveRow.ParentRow.Cells("PICK_NO").Value & ""
                End If
                If CART_TRACKING_NO <> "" Then
                    ASCMAIN1.sql = "Select SOTSHIP1.SHIP_VIA_CODE from SOTSHIP1,SOTPICK1 where SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO and SOTPICK1.PICK_NO = '" & PICK_NO & "'"
                    Dim SHIP_VIA_CODE As String = ASCDATA1.GetDataValue
                    TAC.SOCMAIN1.Track_Shipment(SHIP_VIA_CODE, CART_TRACKING_NO)
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Load_SOTORDRX()
                End If
            Case "CUST_STORE_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Not InquiryMode _
                        And Absx1.txtFor("CUST_CODE").Text <> "" _
                        And Absx1.txtFor("CUST_STORE_NO").Text <> "" _
                        And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If
            Case "ORDR_CUST_PO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If InquiryMode Then
                        If Absx1.txtFor("CUST_CODE").Text <> "" And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                            ASCMAIN1.sql = "Select * from SOTORDR1 where CUST_CODE = :PARM1 and ORDR_CUST_PO = :PARM2 and ORDR_STATUS in ('O','P','F','C')"
                            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VV", New Object() {Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("ORDR_CUST_PO").Text})
                            Dim rows() As DataRow = Nothing
                            If Absx1.txtFor("CUST_STORE_NO").Text <> "" Then
                                rows = tbl.Select("CUST_STORE_NO = '" & Absx1.txtFor("CUST_STORE_NO").Text & "'")
                            Else
                                rows = tbl.Select("")
                            End If
                            If rows.Length > 0 Then
                                Dim ORDR_NO As String = rows(0).Item("ORDR_NO")
                                Absx1.txtFor("ORDR_NO").Text = ORDR_NO
                                Click_Command("View")
                            End If
                        End If
                    ElseIf Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("CUST_STORE_NO").Text <> "" _
                       And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If
            Case "ORDR_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View")
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTORDRX()

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 IsNot Nothing Then
                            ASCMAIN1.sql = "Select Count (*) STORES, Max (CUST_STORE_NO) CUST_STORE_NO from ARTCUST2" _
                                & " where CUST_CODE = :PARM1 and NVL(CUST_DC_IND,'0') <> '1'"
                            Dim row As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {CUST_CODE})
                            If Val(row.Item("STORES") & "") = 1 Then
                                Absx1.txtFor("CUST_STORE_NO").Text = row.Item("CUST_STORE_NO")
                                Absx1.txtFor("CUST_STORE_NO").Focus()
                            End If

                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTORDRX()
            Case "ORDR_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ORDR_ADDR_TYPE_ST"
                If Not Me.IsLoading Then
                    Set_Ship_To()
                End If
        End Select
    End Sub
#End Region

    Sub Load_SOTORDRX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor

        dst.Tables("SOTORDRX").Rows.Clear()

        ASCMAIN1.Progress("Now Building List of Open Sales Orders", "")
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If optShowOrders.Value = "A" And CUST_CODE = "" Then
            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_STATUS = 'O'"
            grdSOTORDRX.Text = "All Open Sales Orders"
        ElseIf optShowOrders.Value = "M" Then
            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_STATUS = 'O' and (INIT_OPER = '" & ASCMAIN1.USER_ID & "' or LAST_OPER = '" & ASCMAIN1.USER_ID & "')"
            grdSOTORDRX.Text = "Open Sales Orders entered or modified by Me"
        ElseIf optShowOrders.Value = "C" Or CUST_CODE <> "" Then
            ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_STATUS = 'O' and CUST_CODE = '" & CUST_CODE & "'"
            grdSOTORDRX.Text = "Open Sales Orders associated with " & CUST_CODE
        Else
            ASCMAIN1.sql = "Select * from SOTORDR1 where ROWNUM < 1"
        End If

        If Trim(ASCMAIN1.USER_CODES) = "FS" Then
            If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                ASCMAIN1.sql &= " and SOTORDR1.SREP_CODE in ('" & Join(TAC.TACMAIN1.SREP_CODEs.ToArray, "','") & "')"
            Else
                ASCMAIN1.sql &= " and SOTORDR1.SREP_CODE = '" & TAC.TACMAIN1.SREP_CODE & "'"
            End If
        End If

        Fill_Records("SOTORDRX", "", , ASCMAIN1.sql)
        Sort_grdColumns(grdSOTORDRX, "ORDR_NO".ToLower)
        grdSOTORDRX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdSOTORDRX.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Sub Print_Record()

        Dim REPORTFILE As String = "SORINVP1"
        If Not REPORTS.ContainsKey(reportFile) Then
            REPORTS.Add(reportFile, Load_rptClass(reportFile))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        If rowSOTORDR1.Item("ORDR_STATUS") = "F" Then
            Dim INV_NO As String = ASCDATA1.GetDataValue("Select INV_NO from SOTINVH1 where ORDR_NO = '" & ORDR_NO & "'")
            REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and SOTINVH1.INV_NO = '" & INV_NO & "'"})
            With REPORTS(REPORTFILE).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "0")
                .Generate_Report(REPORTFILE, "Sales Invoice", , True, , , , , False)
                .Print_Report_End()
            End With
        Else
            REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and ORDR_NO = '" & ORDR_NO & "'", "1", "O"})
            With REPORTS(REPORTFILE).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "0")
                .Generate_Report(REPORTFILE, "Sales Order Confirmation", , True, , , , , False)
                .Print_Report_End()
            End With
        End If


    End Sub

    Private Sub grdSOTORDRX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTORDRX.DoubleClickRow
        Absx1.txtFor("ORDR_NO").Text = e.Row.Cells("ORDR_NO").Value & String.Empty
        Click_Command("View")
    End Sub


#Region "VB6"

    Sub Update_MOM()
        Me.Cursor = Cursors.WaitCursor

        BeginTrans()

        Dim i As Integer

        Dim ORDR_ADDR_TYPE_ST_NEW As String
        Dim ORDR_ADDR_TYPE_ST_OLD As String
        Dim rowARTCUST2 As DataRow = Nothing

        ORDR_ADDR_TYPE_ST_NEW = Absx1.optFor("SOTORDRB.ORDR_ADDR_TYPE_ST").Value
        ORDR_ADDR_TYPE_ST_OLD = Absx1.optFor("SOTORDRB.ORDR_ADDR_TYPE_ST").Tag

        Dim tblARTCUST2 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1", "ARTCUST2", "V", MyBase.Absx1.txtFor("CUST_CODE").Text)

        If ASCMAIN1.Running_in_VS Then Stop ' WHY NOT USE ROWSOTORDRB
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
            rowSOTORDR1.Item("ORDR_SHIP_DATE") = Absx1.dteFor("SOTORDRB.ORDR_SHIP_DATE").Value
            rowSOTORDR1.Item("ORDR_CANCEL_DATE") = Absx1.dteFor("SOTORDRB.ORDR_CANCEL_DATE").Value
            rowSOTORDR1.Item("ORDR_DEPT") = Absx1.txtFor("SOTORDRB.ORDR_DEPT").Text

            ' Look for srep code in ship to table first
            rowARTCUST2 = tblARTCUST2.Rows.Find(New Object() {rowSOTORDR1.Item("CUST_CODE"), rowSOTORDR1.Item("CUST_STORE_NO")})

            If rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("SREP_CODE") & String.Empty <> String.Empty Then
                rowSOTORDR1.Item("SREP_CODE") = rowARTCUST2.Item("SREP_CODE") & String.Empty
            Else
                rowSOTORDR1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & String.Empty
            End If
            rowSOTORDR1.Item("SREP2_CODE") = Absx1.txtFor("SOTORDRB.SREP2_CODE").Text

            rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") = Absx1.optFor("SOTORDRB.ORDR_ADDR_TYPE_ST").Value
            If Absx1.chkFor("SOTORDRB.CUST_FACTOR_IND").Checked Then
                rowSOTORDR1.Item("CUST_FACTOR_IND") = "1"
            Else
                rowSOTORDR1.Item("CUST_FACTOR_IND") = "0"
            End If
            If Absx1.chkFor("SOTORDRB.ORDR_HOLD").Checked Then
                rowSOTORDR1.Item("ORDR_HOLD") = "1"
            Else
                rowSOTORDR1.Item("ORDR_HOLD") = "0"
            End If
            If Absx1.chkFor("SOTORDRB.ORDR_OVERRIDE_NOT_ALLOCATED").Checked Then
                rowSOTORDR1.Item("ORDR_OVERRIDE_NOT_ALLOCATED") = "1"
            Else
                rowSOTORDR1.Item("ORDR_OVERRIDE_NOT_ALLOCATED") = "0"
            End If
        Next

        If ORDR_ADDR_TYPE_ST_NEW <> ORDR_ADDR_TYPE_ST_OLD Then
            ASCMAIN1.sql = "Delete from SOTORDR5 where CODE = 'ST' and ORDR_NO in " _
                & " (Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Select ORDR_NO, CUST_DC_NO, CUST_STORE_NO " _
                & " from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim ORDR_NO As String = row.Item("ORDR_NO")
                If ASCMAIN1.Running_in_VS Then Stop ' WHAT CODE HAS MADE SURE THAT THE CUST_DC_NO IS ACCURATE FOR THIS STORE?
                Update_SOTORDR5(ORDR_NO, ORDR_ADDR_TYPE_ST_NEW, row.Item("CUST_DC_NO"), row.Item("CUST_STORE_NO"))
            Next
        End If

        Dim QTY As Int64

        ASCMAIN1.sql = "Select Distinct WHSE_CODE from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
        Dim WHSE_CODE As String = ASCDATA1.GetDataValue
        Dim ITEM_CODE As String = ""

        For Each rowSOTORDRC As DataRow In dst.Tables("SOTORDRC").Select _
            ("ORDR_QTY_OPEN <> ORDR_QTY_OPEN_ORIG or ITEM_CODE <> ITEM_CODE_ORIG")

            Dim rowSOTORDR2 As DataRow = Nothing

            If rowSOTORDRC.Item("ORDR_LNO_ORIG") & "" <> "" Then
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", rowSOTORDRC.Item("ITEM_CODE"))
                'Dim rowSOTORDR2_Orig As DataRow = LookUp("SOTORDR2", _
                '       New OBJECT() {rowSOTORDRC.Item("ORDR_NO"), rowSOTORDRC.Item("ORDR_LNO_ORIG")})
                Dim rowSOTORDR2_Orig As DataRow = dst.Tables("SOTORDR2").Rows.Find _
                    (New Object() {rowSOTORDRC.Item("ORDR_NO"), rowSOTORDRC.Item("ORDR_LNO_ORIG")})
                rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                With rowSOTORDR2
                    For i = 0 To rowSOTORDR2_Orig.Table.Columns.Count - 1
                        .Item(i) = rowSOTORDR2_Orig.Item(i)
                    Next i
                    .Item("ORDR_LNO") = rowSOTORDRC.Item("ORDR_LNO")
                    .Item("ITEM_CODE") = rowSOTORDRC.Item("ITEM_CODE")
                    .Item("ORDR_QTY") = rowSOTORDRC.Item("ORDR_QTY")
                    .Item("ORDR_QTY_PICK") = 0
                    .Item("ORDR_QTY_SHIP") = 0
                    .Item("ORDR_QTY_CANC") = 0
                    .Item("ITEM_SO_QTY_MULT") = rowICTITEM1.Item("ITEM_SO_QTY_MULT")
                    .Item("CARTON_PACK_QTY") = rowICTITEM1.Item("CARTON_PACK_QTY")
                    .Item("ITEM_UOM") = rowICTITEM1.Item("ITEM_UOM")
                    .Item("ITEM_WEIGHT") = Val(rowICTITEM1.Item("ITEM_WEIGHT") & String.Empty)
                    .Item("ORDR_EXTD_COST") = Val(rowSOTORDRC.Item("ORDR_QTY") & "") * Val(rowICTITEM1.Item("ITEM_COST_STD") & "")
                    .Item("ORDR_STATUS") = "O"
                    .Item("ORDR_QTY_ORIG") = 0
                    .Item("ORDR_QTY_PRE_ALLO") = ""
                End With

                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

            Else
                rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find _
                        (New Object() {rowSOTORDRC.Item("ORDR_NO"), rowSOTORDRC.Item("ORDR_LNO")})

                ITEM_CODE = rowSOTORDR2.Item("ITEM_CODE")
                QTY = -1 * Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                Update_ICTSTAT2(ITEM_CODE, WHSE_CODE, QTY)

                rowSOTORDR2.Item("ITEM_CODE") = rowSOTORDRC.Item("ITEM_CODE")
                rowSOTORDR2.Item("ITEM_CODE_SUB") = rowSOTORDRC.Item("ITEM_CODE_SUB")
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", rowSOTORDRC.Item("ITEM_CODE"))
                rowSOTORDR2.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDRC.Item("ORDR_QTY_CANC") & "")
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = Val(rowSOTORDRC.Item("ORDR_QTY_OPEN") & "")
            End If

            QTY = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            ITEM_CODE = rowSOTORDR2.Item("ITEM_CODE")
            Update_ICTSTAT2(ITEM_CODE, WHSE_CODE, QTY)
        Next

        Update_Record_TDA("SOTORDR2")

        Dim ORDR_STATUS As String
        ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR1.ORDR_STATUS" & vbCrLf _
            & " , Sum (SOTORDR2.ORDR_QTY_OPEN) OPEN" & vbCrLf _
            & " , Sum (SOTORDR2.ORDR_QTY_PICK) PICK" & vbCrLf _
            & " , Sum (SOTORDR2.ORDR_QTY_SHIP) SHIP" & vbCrLf _
            & " , Sum (SOTORDR2.ORDR_QTY_CANC) CANC" & vbCrLf _
            & " from SOTORDR2,SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & " group by SOTORDR1.ORDR_NO, SOTORDR1.ORDR_STATUS" & vbCrLf
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            If Val(row.Item("OPEN") & "") <> 0 Then
                ORDR_STATUS = "O"
            ElseIf Val(row.Item("PICK") & "") <> 0 Then
                ORDR_STATUS = "P"
            ElseIf Val(row.Item("SHIP") & "") <> 0 Then
                ORDR_STATUS = "F"
            Else
                ORDR_STATUS = "C"
            End If
            If ORDR_STATUS <> row.Item("ORDR_STATUS") Then
                ASCMAIN1.sql = "Update SOTORDR1 set ORDR_STATUS = '" & ORDR_STATUS & "'" _
                    & " where ORDR_NO = '" & row.Item("ORDR_NO") & "'"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Update SOTORDR2 set ORDR_STATUS = '" & ORDR_STATUS & "'" _
                    & " where ORDR_NO = '" & row.Item("ORDR_NO") & "'"
                ASCDATA1.ExecuteSQL()
            End If
        Next

        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        CommitTrans("Update Successful")
    End Sub

    Sub Update_ICTSTAT2(ITEM_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVNNNNNN", _
                           New Object() {ITEM_CODE, WHSE_CODE, _
                                         0, 0, 0, _
                                         QTY, 0, 0}, _
                           New String() {"ITEM_CODE_IN", "WHSE_CODE_IN", _
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in", _
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Sub Set_SOTCART1()
        If grdSOTPICK1.ActiveRow Is Nothing OrElse Not grdSOTPICK1.ActiveRow.IsDataRow Then
            grdSOTCART1.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value

            EnforceConstraints(False)
            dst.Tables("SOTCART2").Rows.Clear()
            dst.Tables("SOTCART1").Rows.Clear()
            Fill_Records("SOTCART1", PICK_NO)
            Fill_Records("SOTCART2", PICK_NO)
            EnforceConstraints(True)

            Sort_grdColumns(grdSOTCART1, "CART_NO")
            Set_SOTCART2()
            grdSOTCART1.Text = "Cartons for Pick Ticket " & PICK_NO
            grdSOTCART1.Visible = True
        End If
    End Sub

    Sub Set_SOTCART2()
        'If grdSOTCART1.ActiveRow Is Nothing OrElse Not grdSOTCART1.ActiveRow.IsDataRow Then
        '    'grdSOTCART2.Visible = False
        'Else
        '    Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value
        '    Fill_Records("SOTCART2", CART_NO)
        '    'Sort_grdColumns(grdSOTCART2, "CART_LNO")
        '    'grdSOTCART2.Text = "Cartons Details for Carton " & CART_NO
        '    'grdSOTCART2.Visible = True
        'End If
    End Sub

    Sub Update_SOTORDR5 _
    (ORDR_NO As String, ORDR_ADDR_TYPE_ST As String, CUST_DC_NO As String, CUST_STORE_NO As String)

        Dim CUST_NAME As String
        Dim CUST_EMAIL As String
        Dim CUST_ADDR_TYPE As String
        Dim CUST_ADDR_CODE As String

        Dim rowC As DataRow

        If ORDR_ADDR_TYPE_ST = "BT" Then
            Dim CUST_ADDR_CODE_BT As String = txtCUST_ADDR_CODE_BT.Text
            Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
            Dim rowARTCUST2_BT As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_ADDR_CODE_BT})
            rowC = rowARTCUST1_BT

            CUST_NAME = rowC.Item("CUST_NAME") & ""
            CUST_EMAIL = rowC.Item("CUST_EMAIL") & ""
            CUST_ADDR_TYPE = "BT"
            CUST_ADDR_CODE = CUST_CODE
            'If ASCMAIN1.Running_in_VS Then Stop ' WHY DO WE PUT CUST_DC_NO INTO CUST_ADDR_CODE FOR THE BT ADDRESS RECORD?
        Else
            Dim CUST_NAME_default As String = rowSOTORDR1.Item("CUST_NAME") & " "
            If ORDR_ADDR_TYPE_ST = "DC" Then
                CUST_ADDR_CODE = CUST_DC_NO
                CUST_NAME_default &= "DC #" & CUST_DC_NO
            Else
                CUST_ADDR_CODE = CUST_STORE_NO
                CUST_NAME_default &= "#" & CUST_STORE_NO
                Dim rowARTCUST2_this_CUST_STORE_NO As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                If rowARTCUST2_this_CUST_STORE_NO IsNot Nothing Then CUST_NAME_default = rowARTCUST2_this_CUST_STORE_NO.Item("CUST_STORE_NAME") & ""
            End If

            rowC = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_ADDR_CODE})

            Dim rowARTCUST2_DC As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_DC_NO})
            If rowARTCUST2_DC IsNot Nothing AndAlso rowARTCUST2_DC.Item("CUST_STORE_NAME") & "" <> "" Then
                CUST_NAME = rowARTCUST2_DC.Item("CUST_STORE_NAME")
            Else
                CUST_NAME = CUST_NAME_default
            End If

            CUST_EMAIL = ""

            CUST_ADDR_TYPE = "ST"
            'CUST_ADDR_CODE = rowARTCUST2_DC.Item("CUST_ADDR_CODE").Value
        End If

        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
        With rowSOTORDR5
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
            .Item("CUST_ADDR_CODE") = CUST_ADDR_CODE
            .Item("CUST_NAME") = CUST_NAME
            .Item("CUST_EMAIL") = CUST_EMAIL
            For Each COLUMN_NAME As String In New String() _
                {"CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", _
                 "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX"}
                Dim COLUMN_NAME_source As String = COLUMN_NAME
                If CUST_ADDR_TYPE = "ST" Then COLUMN_NAME_source = Replace(COLUMN_NAME_source, "CUST_", "CUST_STORE_")
                .Item(COLUMN_NAME) = rowC.Item(COLUMN_NAME_source)
            Next
        End With
        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
    End Sub

    Sub Cancel_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        restore_reservation = False

        Dim EMsg As String

        If (EntryMode = "M" Or multiple_order_maintenance) Then
            ASCMAIN1.sql = "Select ORDR_NO from SOTORDR1 " & vbCrLf _
               & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
               & IIf(multiple_order_type = "ORDR_GROUP_NO", _
                     " and ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", _
                     " and EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") & vbCrLf _
               & " and ORDR_STATUS = 'O'"
            Dim dt As DataTable = ASCDATA1.GetDataTable
            For Each row As DataRow In dt.Rows
                Cancel_Order_1(row.Item("ORDR_NO"))
            Next
            EMsg = CStr(dt.Rows.Count) & IIf(multiple_order_type = "ORDR_GROUP_NO", _
                                             " Orders from Order Group " & ORDR_GROUP_NO, _
                                             " Orders from EDI Journal " & EDI_JRNL_NO) & " have been Cancelled"
        Else
            Cancel_Order_1(ORDR_NO)
            EMsg = "Order " & ORDR_NO & " has been Cancelled"
        End If

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        Next

        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Cancel_Order_1(ORDR_NO As String)
        Dependent_Updates(-1, ORDR_NO)

        ASCMAIN1.sql = "Select Sum (ORDR_QTY_PICK) from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        Dim ORDR_STATUS As String = ""
        If Val(ASCDATA1.GetDataValue) <> 0 Then
            ORDR_STATUS = "P"
        Else
            ASCMAIN1.sql = "Select Sum (ORDR_QTY_SHIP) from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
            If Val(ASCDATA1.GetDataValue) <> 0 Then
                ORDR_STATUS = "F"
            Else
                ORDR_STATUS = "C"
            End If
        End If

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2" _
            & "    Set ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + NVL(R1.ORDR_QTY_OPEN,0)" _
            & "      , ORDR_QTY_OPEN = 0, ORDR_STATUS = '" & ORDR_STATUS & "'" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        TAC.SOCMAIN1.Record_Event_SOTORDRE(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "ORDCXL", "Order Cancelled")

        ASCMAIN1.sql = "Update SOTORDR1 Set REASON_CODE = :PARM1" _
            & ", ORDR_STATUS = :PARM2, ORDR_DATE_CLOSED = TRUNC(SYSDATE), ORDR_YYYYPP_CLOSED = :PARM3" _
            & " where ORDR_NO = :PARM4"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {Absx1.txtFor("REASON_CODE").Text, ORDR_STATUS, ASCMAIN1.CYP, ORDR_NO})
    End Sub

    Sub Reverse_Cancel(ORDR_NO As String, ORDR_GROUP_NO As String, Optional dt As DataTable = Nothing)
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        restore_reservation = False
        Dim EMsg As String

        If ORDR_NO <> "" Then
            Reverse_Cancel_1(ORDR_NO)
            EMsg = "Order " & ORDR_NO & " has been Re-Opened"
        Else
            If ASCMAIN1.Running_in_VS Then Stop ' do we look at a selected indicator?
            For Each row As DataRow In dt.Rows
                Reverse_Cancel_1(row.Item("ORDR_NO"))
            Next
            EMsg = CStr(dt.Rows.Count) & " Orders from Order Group " & ORDR_GROUP_NO & " have been Re-Opened"
        End If

        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        CommitTrans("Process Complete")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Reverse_Cancel_1(ORDR_NO As String)
        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2" _
            & "    Set ORDR_QTY_CANC = 0" _
            & "      , ORDR_QTY_OPEN = NVL(R1.ORDR_QTY_CANC,0), ORDR_STATUS = 'O'" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR1 Set REASON_CODE = Null" _
            & ", ORDR_STATUS = 'O', ORDR_DATE_CLOSED = Null, ORDR_YYYYPP_CLOSED = Null" _
            & " where ORDR_NO = :PARM1"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {ORDR_NO})

        Dependent_Updates(1, ORDR_NO)
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        restore_reservation = False
        Dim EMsg As String

        If (EntryMode = "M" Or multiple_order_maintenance) Then
            ASCMAIN1.sql = "Select ORDR_NO from SOTORDR1 " & vbCrLf _
               & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
               & IIf(multiple_order_type = "ORDR_GROUP_NO", _
                     " and ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", _
                     " and EDI_JRNL_NO = '" & EDI_JRNL_NO & "'") & vbCrLf _
               & " and ORDR_STATUS = 'O'"
            Dim dt As DataTable = ASCDATA1.GetDataTable
            For Each row As DataRow In dt.Rows
                Delete_Order_1(row.Item("ORDR_NO"))
            Next
            EMsg = CStr(dt.Rows.Count) & IIf(multiple_order_type = "ORDR_GROUP_NO", _
                                  " Orders from Order Group " & ORDR_GROUP_NO, _
                                  " Orders from EDI Journal " & EDI_JRNL_NO) & " have been marked as Deleted"
        Else
            Delete_Order_1(ORDR_NO)
            EMsg = "Order " & ORDR_NO & " has been marked as Deleted"
        End If

        For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
        Next

        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(ORDR_NO As String)
        Dependent_Updates(-1, ORDR_NO)

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 Is Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2" _
            & "    Set ORDR_QTY_OPEN = 0, ORDR_STATUS = '" & "D" & "'" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTORDR1 Set REASON_CODE = :PARM1" _
            & ", ORDR_STATUS = :PARM2, ORDR_DATE_CLOSED = TRUNC(SYSDATE), ORDR_YYYYPP_CLOSED = :PARM3" _
            & " where ORDR_NO = :PARM4"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {Absx1.txtFor("REASON_CODE").Text, "D", ASCMAIN1.CYP, ORDR_NO})
    End Sub

    Sub Dependent_Updates(S As Integer, ORDR_NO As String)

        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

            If S = -1 Then
            Else
                ' Update_Record_TDA("SOTORDR2")
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                ITEM_CODE = rowSOTORDR2.Item("ITEM_CODE")
                Update_ICTSTAT2(ITEM_CODE, WHSE_CODE, S * QTY_TO_COMMIT)
            End If
        Next
    End Sub

    Sub Display_Totals()

        Dim KEY As Int32 = 0
        For Each SFX As String In New String() {"", "OPEN", "PICK", "SHIP", "CANC"}
            If SFX <> "" Then SFX = "_" & SFX
            KEY += 1
            Dim rowSOTORDRT As DataRow = dst.Tables("SOTORDRT").Rows.Find(KEY)
            rowSOTORDRT.Item("QTY") = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY" & SFX & ")", "") & "")
            rowSOTORDRT.Item("AMT") = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT" & SFX & ")", "") & "")
        Next

        If multi_store_is_active Then
            Set_MS_TOTAL_AMT_Expression()
        End If
    End Sub

    Function Select_Item() As String

        Dim ITEM_CODE As String = ""

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("ITEM_CODE")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            ITEM_CODE = ASCMAIN1.CodeSelector.SelectedCode
        End If

        Return ITEM_CODE
    End Function

    Sub Setup_MS(tf As Boolean)

        UltraExplorerBar1.Groups("Multi-Store").Items("Start Multi-Store").Text = IIf(tf, "Clear Multi-Store", "Start Multi-Store")

        tabMain.Tabs("Multi-Store").Visible = tf
        ' UltraExplorerBar1.Groups("Multi-Store").Visible = tf
        lblMultiStore.Visible = tf


        multi_store_is_active = tf

        msqty = 0
        msqty_col = 0

        If multi_store_is_active Then
            Init_MultiStore()

            With UltraExplorerBar1.Groups("Multi-Store")
                .Items("Reset Qty's").Settings.Enabled = DefaultableBoolean.True
                .Items("Re-Select Stores").Settings.Enabled = DefaultableBoolean.True
                .Items("Clear Zeroes").Settings.Enabled = DefaultableBoolean.True
            End With

        Else
            CUST_STORE_NOs_multi_store.Clear()

            With UltraExplorerBar1.Groups("Multi-Store")
                .Items("Reset Qty's").Settings.Enabled = DefaultableBoolean.False
                .Items("Re-Select Stores").Settings.Enabled = DefaultableBoolean.False
                .Items("Clear Zeroes").Settings.Enabled = DefaultableBoolean.False
            End With
        End If
    End Sub

    Sub Init_MultiStore()
        Me.Cursor = Cursors.WaitCursor

        dst.Tables("SOTORDRS").Rows.Clear()

        ' Add a column for each style referenced in grdSOWORDR2
        dst.Tables("SOTORDRS").Columns("TOTAL_QTY").Expression = ""
        dst.Tables("SOTORDRS").Columns("TOTAL_AMT").Expression = ""
        For i As Integer = dst.Tables("SOTORDRS").Columns.Count - 1 To 0 Step -1
            Dim DC As DataColumn = dst.Tables("SOTORDRS").Columns(i)
            If DC.ColumnName = "TOTAL_AMT" Or DC.ColumnName = "TOTAL_QTY" Then
                Exit For
            Else
                Dim summary As UltraWinGrid.SummarySettings = grdSOTORDRS.DisplayLayout.Bands(0).Summaries(DC.ColumnName)
                grdSOTORDRS.DisplayLayout.Bands(0).Summaries.Remove(summary)

                dst.Tables("SOTORDRS").Columns.Remove(DC)
            End If
        Next

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select()
            Fill_Items(rowSOTORDR2.Item("ORDR_LNO"), rowSOTORDR2.Item("ITEM_CODE") & "", False)
        Next
        Set_MS_TOTAL_QTY_Expression()

        ' Add a row for each store referenced in CUST_STORE_NOs_multi_store
        Fill_Stores(CUST_STORE_NOs_multi_store)

        tabMain.SelectedTab = tabMain.Tabs("Multi-Store")

        lblMSCopyToStore.Visible = False
        txtMSCopyToStore.Visible = False
        optMSCopyToStore.Visible = False
        optMSCopyToStore.Tag = ""
        lblMSCopyToStore.Tag = ""

        Me.Cursor = Cursors.Default
    End Sub

    Sub Fill_Items(ORDR_LNO As Int64, ITEM_CODE As String, _
                    Optional reset_TOTAL_QTY As Boolean = True)

        Dim COLUMN_NAME As String = "QTY_" & Format(ORDR_LNO, "000")
        dst.Tables("SOTORDRS").Columns.Add(COLUMN_NAME, GetType(System.Int64))

        With grdSOTORDRS.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
            .Hidden = False
            .Width = grdSOTORDRS.DisplayLayout.Bands(0).Columns("TOTAL_QTY").Width
            .Format = "#,##0"
            .CellAppearance.TextHAlign = HAlign.Right
            .Header.Appearance.TextHAlign = HAlign.Right
            .Header.Caption = ITEM_CODE
            Create_Summary(grdSOTORDRS, COLUMN_NAME)
            .CellActivation = UltraWinGrid.Activation.AllowEdit
        End With

        If reset_TOTAL_QTY Then Set_MS_TOTAL_QTY_Expression(ORDR_LNO)
    End Sub

    Private Sub ReSelect_Stores()
        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_STORE_NO")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True

            ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "'" '  and CUST_ADDR_STATUS = 'A'"
            Dim TBL As DataTable = ASCDATA1.GetDataTable
            ASCMAIN1.CodeSelector.UseDataFromTable = TBL

            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Join(CUST_STORE_NOs_multi_store.ToArray, Chr(0))
            ' If ASCMAIN1.Running_in_VS Then Stop ' Sql = Sql & " and CUST_ADDR_STATUS = 'A'"
            Using F As New ASFCODE1
                F.ShowDialog()
            End Using
            If ASCMAIN1.CodeSelector.SelectedCodes.Count <> 0 Then
                CUST_STORE_NOs_multi_store.Clear()
                For Each CC As String In ASCMAIN1.CodeSelector.SelectedCodes
                    CUST_STORE_NOs_multi_store.Add(CC)
                Next
                If Not CUST_STORE_NOs_multi_store.Contains(Absx1.txtFor("CUST_STORE_NO").Text) Then
                    CUST_STORE_NOs_multi_store.Add(Absx1.txtFor("CUST_STORE_NO").Text)
                End If

                Dim C As New List(Of String)
                For Each CUST_STORE_NO As String In CUST_STORE_NOs_multi_store
                    C.Add(CUST_STORE_NO)
                Next

                If multi_store_is_active Then ' Delete and Add Stores to table as required
                    If dst.Tables("SOTORDRS").Rows.Count > 0 Then
                        For i As Integer = dst.Tables("SOTORDRS").Rows.Count - 1 To 0
                            Dim row As DataRow = dst.Tables("SOTORDRS").Rows(i)
                            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                            If Not CUST_STORE_NOs_multi_store.Contains(CUST_STORE_NO) Then
                                row.Delete()
                            Else
                                C.Remove(CUST_STORE_NO)
                            End If
                        Next
                    End If
                    If C.Count <> 0 Then
                        Fill_Stores(C)
                    End If
                Else
                    Setup_MS(True)
                End If
            End If
        End If
    End Sub

    Sub Fill_Stores(ByVal CUST_STORE_NOs As List(Of String))
        Dim QTYs As New Dictionary(Of Int64, Int64)
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            QTYs.Add(Val(rowSOTORDR2.Item("ORDR_LNO") & ""), Val(rowSOTORDR2.Item("ORDR_QTY") & ""))
        Next

        multi_store_changes_made_to_SOTORDRS = True

        For Each CUST_STORE_NO As String In CUST_STORE_NOs
            Dim rowSOTORDRS As DataRow = dst.Tables("SOTORDRS").Rows.Find(CUST_STORE_NO)
            If rowSOTORDRS Is Nothing Then
                rowSOTORDRS = dst.Tables("SOTORDRS").NewRow
                rowSOTORDRS.Item("CUST_STORE_NO") = CUST_STORE_NO
                rowSOTORDRS.Item("ORDR_CUST_PO") = Absx1.txtFor("ORDR_CUST_PO").Text
                For Each ORDR_LNO As Int64 In QTYs.Keys
                    If QTYs(ORDR_LNO) <> 0 Then
                        rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) = QTYs(ORDR_LNO)
                    End If
                Next
                dst.Tables("SOTORDRS").Rows.Add(rowSOTORDRS)
            End If
        Next

        With dst.Tables("SOTORDRS")
            .AcceptChanges()
            If .Rows.Count > 0 Then
                For R As Integer = .Rows.Count - 1 To 0 Step -1
                    Dim CUST_STORE_NO As String = .Rows(R).Item("CUST_STORE_NO")
                    If Not CUST_STORE_NOs.Contains(CUST_STORE_NO) Then
                        .Rows.Remove(.Rows(R))
                    End If
                Next
            End If
            .AcceptChanges()
        End With


        multi_store_changes_made_to_SOTORDRS = False

        Sort_grdColumns(grdSOTORDRS, "CUST_STORE_NO")
        Set_MS_TOTAL_AMT_Expression()

    End Sub

    Sub Set_MS_TOTAL_AMT_Expression()
        Dim EXP As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
            Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
            EXP &= "+ISNULL(QTY_" & Format(ORDR_LNO, "000") & ",0) * " & CStr(ORDR_UNIT_PRICE)
        Next
        dst.Tables("SOTORDRS").Columns("TOTAL_AMT").Expression = Mid(EXP, 2)
    End Sub

    Sub Set_MS_TOTAL_QTY_Expression(Optional ORDR_LNO As Int64 = 0)
        Dim EXP As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            EXP &= "+ISNULL(QTY_" & Format(Val(rowSOTORDR2.Item("ORDR_LNO") & ""), "000") & ",0)"
        Next
        If ORDR_LNO <> 0 Then
            If InStr(EXP, "QTY_" & Format(ORDR_LNO, "000")) = 0 Then
                EXP &= "+QTY_" & Format(ORDR_LNO, "000")
            End If
        End If

        dst.Tables("SOTORDRS").Columns("TOTAL_QTY").Expression = Mid(EXP, 2)
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTORDR2.ActiveRow
            Select Case COLUMN_NAME

                Case "ITEM_CODE"
                    If .Cells("ITEM_CODE").Text <> "" Then
                        Dim ITEM_CODE As String = Validate_Item(.Cells("ITEM_CODE").Value & "")
                        Cancel = (ITEM_CODE = "")
                    End If

                Case "ORDR_QTY"
                    If Trim(.Cells("ITEM_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(.Cells("ORDR_QTY").Value & "") = "" Then
                        MsgBox("Order Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY")
                        Exit Sub
                    End If
                    If Val(.Cells("ORDR_QTY").Value & "") < 0 Then
                        MsgBox("Order Qty May Not be Negative", vbOKOnly, "Invalid Order Quantity")
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Function Validate_Item(ITEM_CODE_z As String) As String
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

        If E <> "" And grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Item Code Entered is Invalid because ...")
        Else
            If E = "" Then
                ITEM_CODE = rowICTITEM1.Item(0)
            End If
        End If
        Return ITEM_CODE
    End Function
#End Region

#Region "grdSOTORDR2"

    Private Sub grdSOTORDR2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)
                If ITEM_CODE <> "" Then
                    e.Cell.Row.Cells("ITEM_UOM").Value = rowICTITEM1.Item("ITEM_UOM") & ""
                    e.Cell.Row.Cells("ITEM_WEIGHT").Value = Val(rowICTITEM1.Item("ITEM_WEIGHT") & "")
                    e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                    e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1.Item("ITEM_RETAIL_PRICE")

                    Dim ORDR_UNIT_PRICE As Decimal = TAC.SOCMAIN1.Get_Price _
                                                     (Me, _
                                                      PRICE_LIST_CODE, _
                                                      PRICE_BASIS, _
                                                      PRICE_BASE_DPCT, _
                                                      ITEM_CODE, _
                                                      rowICTITEM1, _
                                                      rowSOTORDR1.Item("ORDR_DATE_BOOKED"))

                    e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE

                    'Dim rowSOTPRIC2 As DataRow = Nothing
                    'If PRICE_LIST_CODE <> "" Then
                    '    rowSOTPRIC2 = LookUp("SOTPRIC2", New String() {PRICE_LIST_CODE, ITEM_CODE})
                    'End If
                    'If rowSOTPRIC2 IsNot Nothing Then
                    '    If rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") & "" <> "" AndAlso Format(rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE"), "yyyyMMdd") <= Format(rowSOTORDR1.Item("ORDR_DATE_BOOKED"), "yyyyMMdd") Then
                    '        e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = Val(rowSOTPRIC2.Item("ITEM_NEW_PRICE") & "")
                    '    Else
                    '        e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                    '    End If
                    'Else
                    '    If PRICE_BASIS = "R" Then
                    '        Dim ITEM_RETAIL_PRICE As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                    '        e.Cell.Row.Cells("ORDR_UNIT_PRICE").Value = ITEM_RETAIL_PRICE * (100 - PRICE_BASE_DPCT) / 100
                    '    End If
                    'End If

                    e.Cell.Row.Cells("ITEM_SO_QTY_MULT").Value = rowICTITEM1.Item("ITEM_SO_QTY_MULT")
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = rowICTITEM1.Item("CARTON_PACK_QTY")
                    '   e.Cell.Row.Cells("SALES_DIVISION_CODE").Value = rowICTITEM1.Item("SALES_DIVISION_CODE")
                End If

            Case "ORDR_QTY"
                grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value = grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value

            Case "ORDR_QTY_OPEN"
                grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") _
                    - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") _
                    - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_OPEN").Value & "") _
                    - Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "")
                If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value) < 0 Then
                    grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value = 0
                End If

        End Select
    End Sub

    Private Sub grdSOTORDR2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTORDR2.AfterRowActivate

        If Trim(grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value & "") = "" And _
            (grdSOTORDR2.ActiveCell Is Nothing OrElse _
             grdSOTORDR2.ActiveCell.Column.Key <> "ITEM_CODE") _
        Then
            grdSOTORDR2.ActiveCell = grdSOTORDR2.ActiveRow.Cells("ITEM_CODE")
        End If

        If grdSOTORDR2.ActiveRow.IsAddRow Then
            With grdSOTORDR2.DisplayLayout.Bands(0)

                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ITEM_DESC").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End With

        Else
            With grdSOTORDR2.DisplayLayout.Bands(0)
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("ITEM_DESC").CellActivation = UltraWinGrid.Activation.NoEdit


                If Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
                Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
                Then
                    ' Or Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value & "") <> 0 _
                    'Or absx1.txtfor("ORDR_SOURCE")).Text = "E" 'was also part of this
                    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("ORDR_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("ORDR_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End With

        End If
    End Sub

    Private Sub grdSOTORDR2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTORDR2.AfterRowsDeleted
        If multi_store_is_active Then ' And Not multi_store_changes_made_to_SOTORDRS Then
            dst.Tables("SOTORDRS").Columns("TOTAL_AMT").Expression = ""
            dst.Tables("SOTORDRS").Columns("TOTAL_QTY").Expression = ""
            For Each ORDR_LNO As Int64 In ORDR_LNOs
                Dim summary As UltraWinGrid.SummarySettings = grdSOTORDRS.DisplayLayout.Bands(0).Summaries("QTY_" & Format(ORDR_LNO, "000"))
                grdSOTORDRS.DisplayLayout.Bands(0).Summaries.Remove(summary)
                dst.Tables("SOTORDRS").Columns.Remove("QTY_" & Format(ORDR_LNO, "000"))
            Next
        End If

        Display_Totals()

        If grdSOTORDR2.Rows.Count = 0 Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = ""
        End If

        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")

    End Sub

    Private Sub grdSOTORDR2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDR2.AfterRowUpdate

        If multi_store_is_active Then
            If multi_store_changes_made_to_SOTORDRS Then
                multi_store_changes_made_to_SOTORDRS = False ' Turn off multix from BeforeUpdate
            Else
                multi_store_changes_made_to_SOTORDRS = True
                Dim rowSOTORDRS As DataRow = dst.Tables("SOTORDRS").Rows.Find(Absx1.txtFor("CUST_STORE_NO").Text)
                Dim ORDR_LNO As Int64 = Val(e.Row.Cells("ORDR_LNO").Value & "")
                rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) = Val(e.Row.Cells("ORDR_QTY").Value)
                multi_store_changes_made_to_SOTORDRS = False
            End If
        End If

        Display_Totals()

        'If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
        '    Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTITEM1.Item("SALES_DIVISION_CODE")
        'End If

        ' If e.Row.IsAddRow Then
        ' if we just added a row
        If EntryMode = "N" Or EntryMode = "E" Then
            If e.Row.Cells("ORDR_STATUS").Tag & "" = "Added" Then
                Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
                grdSOTORDR2.DisplayLayout.Bands(0).AddNew()
                e.Row.Cells("ORDR_STATUS").Tag = DBNull.Value
            End If
        End If
        ' End If

    End Sub

    Private Sub grdSOTORDR2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDR2.BeforeCellUpdate

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim ITEM_CODE As String = Validate_Item(e.NewValue & "")
                If ITEM_CODE = "" Then
                    e.Cancel = True
                End If
        End Select

    End Sub

    Private Sub grdSOTORDR2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTORDR2.BeforeExitEditMode
        If grdSOTORDR2.ActiveCell IsNot Nothing Then
            With grdSOTORDR2.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                    Case "ORDR_QTY"
                        If .EditorResolved.Value & "" = "" _
                        Or Val(.EditorResolved.Value & "") < 0 _
                        Then
                            .EditorResolved.Value = 0
                        End If
                        If Val(.EditorResolved.Value & "") < 0 Then
                            .EditorResolved.Value = System.Math.Abs(Val(.EditorResolved.Value & ""))
                        End If

                    Case "ORDR_QTY_OPEN"
                        If .EditorResolved.Value & "" = "" _
                        Or Val(.EditorResolved.Value & "") < 0 _
                        Then
                            .EditorResolved.Value = 0
                        End If
                        Dim q As Int64 = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_PICK").Value & "") _
                                       + Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_SHIP").Value & "") _
                                       + Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY_CANC").Value & "")
                        If Val(.EditorResolved.Value & "") > Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") - q Then
                            .EditorResolved.Value = Val(grdSOTORDR2.ActiveRow.Cells("ORDR_QTY").Value & "") - q
                        End If
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTORDR2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTORDR2.BeforeRowsDeleted

        ORDR_LNOs.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Not grow.IsAddRow Then
                If Val(grow.Cells("ORDR_QTY_PICK").Value & "") <> 0 _
         Or Val(grow.Cells("ORDR_QTY_SHIP").Value & "") <> 0 _
         Or Val(grow.Cells("ORDR_QTY_CANC").Value & "") <> 0 _
         Then
                    MsgBox("Cannot Delete a Line if it has ever been " & vbCr & "Picked, Shipped Or Cancelled" & vbCr & "Use the Cancel Button (x)")
                    e.Cancel = True
                    Exit Sub
                End If

                If Absx1.optFor("ORDR_SOURCE").Value = "E" Then
                    If ALLOW_CHANGE_RANGE <> "1" Then
                        MsgBox("Cannot Delete a Line from an EDI Order" & vbCrLf & "Use the Cancel Button (x)")
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
                ORDR_LNOs.Add(grow.Cells("ORDR_LNO").Value)
            End If
        Next
    End Sub

    Private Sub grdSOTORDR2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDR2.BeforeRowUpdate

        Validate_Columns("ITEM_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("ORDR_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        ITEM_CODE_last_entry = e.Row.Cells("ITEM_CODE").Value & ""

        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_NO").Value = ORDR_NO
            Dim ORDR_LNO As Int64 = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_LNO)", "") & "") + 1
            e.Row.Cells("ORDR_LNO").Value = ORDR_LNO
            e.Row.Cells("ORDR_QTY_ORIG").Value = e.Row.Cells("ORDR_QTY").Value
            e.Row.Cells("ORDR_STATUS").Value = "O"
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
            e.Row.Cells("CUST_STORE_NO").Value = Absx1.txtFor("CUST_STORE_NO").Text
            e.Row.Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
            e.Row.Cells("ORDR_STATUS").Tag = "Added"

            If multi_store_is_active Then
                multi_store_changes_made_to_SOTORDRS = True
                Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
                Fill_Items(ORDR_LNO, ITEM_CODE)
                For Each rowSOTORDRS As DataRow In dst.Tables("SOTORDRS").Select("")
                    rowSOTORDRS.Item("QTY_" & Format(ORDR_LNO, "000")) = e.Row.Cells("ORDR_QTY").Value
                Next
                'Set_MS_TOTAL_AMT_Expression() ' gonna happen in afterrowupdate -> DisplayTotals
                'multi_store_changes_made_to_SOTORDRS = False 
                ' Turn multi_store_changes_made_to_SOTORDRS off in the AfterUpdate event
            End If
        End If
    End Sub

    Private Sub grdSOTORDR2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDR2.ClickCellButton

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
                        grdSOTORDR2.ActiveRow.Update()
                    End If

                Case "ITEM_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTORDR2, sql_where)


            End Select
        End With

    End Sub

    Private Sub grdSOTORDR2_Error(sender As Object, e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSOTORDR2.Error
        e.Cancel = True
    End Sub

    Private Sub grdSOTORDR2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDR2.InitializeRow
        Dim ITEM_SO_QTY_MULT As Integer = Val(e.Row.Cells("ITEM_SO_QTY_MULT").Value & "")
        Dim ORDR_QTY As Integer = Val(e.Row.Cells("ORDR_QTY").Value & "")
        If ITEM_SO_QTY_MULT <> 0 AndAlso ORDR_QTY Mod ITEM_SO_QTY_MULT <> 0 Then
            e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Cells("ORDR_QTY").ToolTipText = "Order Qty not Divisible by Inner Pack Qty"
        Else
            e.Row.Cells("ORDR_QTY").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Cells("ORDR_QTY").ToolTipText = ""
        End If
    End Sub

    Private Sub grdSOTORDR2_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles grdSOTORDR2.KeyDown
        With grdSOTORDR2
            Try
                If e.KeyCode = Keys.F5 Then
                    Select Case .ActiveCell.Column.Key
                        Case "ITEM_CODE"
                            grdSOTORDR2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
                            .ActiveCell.Value = ITEM_CODE_last_entry
                            .ActiveCell.SelStart = Len(grdSOTORDR2.ActiveCell.Text)
                    End Select
                End If
            Catch ex As Exception

            End Try
        End With
    End Sub

#End Region

#Region "grdSOTORDR4"

    Private Sub grdSOTORDR4_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTORDR4.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("ORDR_CLNO").Value = Val(dst.Tables("SOTORDR4").Compute("MAX(ORDR_CLNO)", "") & "") + 1
            e.Row.Cells("ORDR_NO").Value = ORDR_NO
            e.Row.Cells("INIT_DATE").Value = DATETIME_STAMP
            e.Row.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
        Else
            e.Row.Cells("LAST_DATE").Value = DATETIME_STAMP
            e.Row.Cells("LAST_OPER").Value = ASCMAIN1.USER_ID
        End If
    End Sub
#End Region

#Region "grdSOTORDRS"

    Private Sub grdSOTORDRS_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTORDRS.AfterCellUpdate
        msqty = Val(e.Cell.Value & "")
        msqty_col = Val(Split(e.Cell.Column.Key, "_")(1))
        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, msqty_col})

        With optMSCopyToStore.ValueList
            .ValueListItems(0).DisplayText = "Copy All Items from Store " & e.Cell.Row.Cells("CUST_STORE_NO").Value
            .ValueListItems(0).Tag = e.Cell.Row.Cells("CUST_STORE_NO").Value
            .ValueListItems(1).DisplayText = "Copy Qty of " & CStr(msqty) & " to Item " & rowSOTORDR2.Item("ITEM_CODE")
            .ValueListItems(1).Tag = rowSOTORDR2.Item("ITEM_CODE")
            .ValueListItems(2).DisplayText = "Copy Qty of " & CStr(msqty) & " to All Items"
        End With

        lblMSCopyToStore.Visible = True
        txtMSCopyToStore.Visible = True
        optMSCopyToStore.Visible = True
    End Sub

    Private Sub grdSOTORDRS_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTORDRS.AfterRowsDeleted
        ' Set_MS_TOTAL_AMT_Expression()
    End Sub

    Private Sub grdSOTORDRS_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTORDRS.AfterRowUpdate
        '  Set_MS_TOTAL_AMT_Expression()

        If e.Row.Cells("CUST_STORE_NO").Value = Absx1.txtFor("CUST_STORE_NO").Text And Not multi_store_changes_made_to_SOTORDRS Then
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDR2.Rows
                If grow.IsAddRow Then
                Else
                    Dim ORDR_LNO As Int64 = Val(grow.Cells("ORDR_LNO").Value & "")
                    Dim ORDR_QTY As Int64 = Val(grow.Cells("ORDR_QTY").Value & "")
                    Dim QTY As Int64 = Val(e.Row.Cells("QTY_" & Format(ORDR_LNO, "000")).Value & "")
                    If Val(grow.Cells("ORDR_QTY").Value & "") <> QTY Then
                        multi_store_changes_made_to_SOTORDRS = True
                        grow.Cells("ORDR_QTY").Value = QTY
                        grow.Update()
                        multi_store_changes_made_to_SOTORDRS = False
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub grdSOTORDRS_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTORDRS.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.Cells("CUST_STORE_NO").Value = Absx1.txtFor("CUST_STORE_NO").Text Then
                MsgBox("You Cannot Delete the Main Store", MsgBoxStyle.OkOnly, "Cannot Delete")
                e.Cancel = True
                e.DisplayPromptMsg = False
            End If
        Next
    End Sub
#End Region

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCART1.AfterRowActivate
        Set_SOTCART2()
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        'Set_SOTPICK2()
        Set_SOTCART1()
    End Sub


    Sub Clear_Zeroes()
        Dim sqlw As String = ""
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
            Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
            Dim COLUMN_NAME As String = "QTY_" & Format(ORDR_LNO, "000")
            sqlw &= " and (ISNULL(" & COLUMN_NAME & ",0) = 0)"
        Next
        sqlw = "CUST_STORE_NO <> '" & Absx1.txtFor("CUST_STORE_NO").Text & "'" & sqlw
        ASCDATA1.DeleteRows(dst.Tables("SOTORDRS"), sqlw)
    End Sub

    Sub Load_Events()
        '    grdEvents.RemoveAll
        '    Call Load_Events_1("Entered", "INIT_DATE")
        '    Call Load_Events_1("Modified", "LAST_DATE")
        '    Call Load_Events_1("Released", "ORDR_DATE_REL")
        '    Call Load_Events_1("Pick Ticket", "ORDR_DATE_PICK_PRT")
        '    Call Load_Events_1("Packed", "ORDR_DATE_PACKED")
        '    Call Load_Events_1("Shipped", "ORDR_DATE_SHIPPED")
        '    Call Load_Events_1("Invoice", "ORDR_INV_DATE")
        '    Call Load_Events_1("Invoice Prt", "ORDR_DATE_INV_PRT")
        '    Call Load_Events_1("Updated", "ORDR_DATE_UPDATED")
        '    Call Load_Events_1("Cancelled", "ORDR_DATE_CANCELLED")
    End Sub

    Private Sub txtMSStore_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtMSCopyToStore.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtMSCopyToStore.Text = "" Or msqty_col = 0 Then
                MsgBox("You Must First Enter a Store No")
                Exit Sub
            End If

            Dim CUST_STORE_NO_Copy_To As String = txtMSCopyToStore.Text.PadLeft(6, "0")
            If dst.Tables("SOTORDRS").Rows.Find(CUST_STORE_NO_Copy_To) Is Nothing Then
                MsgBox("Store " & CUST_STORE_NO_Copy_To & " Not in Multi-Store Grid" & vbCrLf & "You must enter a Store No to Copy To from the Stores listed above")
                txtMSCopyToStore.Text = ""
                Exit Sub
            End If

            Dim CUST_STORE_NO_Copy_From As String = optMSCopyToStore.ValueList.ValueListItems(0).Tag ' grdSOTORDRS.ActiveRow.Cells("CUST_STORE_NO").Value
            If grdSOTORDRS.ActiveRow Is Nothing OrElse _
                grdSOTORDRS.ActiveRow.Cells("CUST_STORE_NO").Value <> CUST_STORE_NO_Copy_To Then
                For Each grow As UltraWinGrid.UltraGridRow In grdSOTORDRS.Rows
                    If grow.Cells("CUST_STORE_NO").Value = CUST_STORE_NO_Copy_To Then
                        grdSOTORDRS.ActiveRow = grow
                        Exit For
                    End If
                Next
            End If

            Dim rowSOTORDRS_Copy_From As DataRow = dst.Tables("SOTORDRS").Rows.Find(CUST_STORE_NO_Copy_From)
            lblMSCopyToStore.Tag = "x"

            ' Dim rowSOTORDRS As DataRow = dst.Tables("SOTORDRS").Rows.Find(txtMSCopyToStore.Text.PadLeft(6, "0"))
            If grdSOTORDRS.ActiveRow IsNot Nothing Then
                If optMSCopyToStore.Value = "Store" Then
                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                        Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim C As String = "QTY_" & Format(ORDR_LNO, "000")
                        grdSOTORDRS.ActiveRow.Cells(C).Value = rowSOTORDRS_Copy_From.Item(C)
                    Next
                    'For i As Integer = grdMS_Cols_Orig + 0 To dst.Tables("SOTORDRS").Columns.Count - 1
                    '    grdSOTORDRS.ActiveRow.Cells(i).Value = rowSOTORDRS_this.Item(i)
                    'Next i
                ElseIf optMSCopyToStore.Value = "Item" Then
                    grdSOTORDRS.ActiveRow.Cells("QTY_" & Format(msqty_col, "000")).Value = msqty
                ElseIf optMSCopyToStore.Value = "Qty" Then
                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("")
                        Dim ORDR_LNO As Int64 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim C As String = "QTY_" & Format(ORDR_LNO, "000")
                        grdSOTORDRS.ActiveRow.Cells(C).Value = msqty
                    Next
                    'For i As Integer = grdMS_Cols_Orig + 0 To dst.Tables("SOTORDRS").Columns.Count - 1
                    '    grdSOTORDRS.ActiveRow.Cells(i).Value = msqty
                    'Next i
                End If
                grdSOTORDRS.ActiveRow.Update()
                txtMSCopyToStore.Text = ""
            End If
        End If
    End Sub

    Function Check_Changed_Fields() As Boolean

        'REV_NO += 1
        'Dim REV_LNO As Integer = 0
        'dst.Tables("SOTORDXR").Rows.Clear()

        Dim LAST_DATE As Date = DATETIME_STAMP
        If EntryMode = "N" Then Stop

        Check_Changed_Fields = False

        ASCMAIN1.Progress("Logging Header Changes")

        TAC.SOCMAIN1.Log_Changes(Me, ORDR_NO, rowSOTORDR1, "SOTORDR1", Check_Changed_Fields, REV_NO, REV_LNO, LAST_DATE)

        ASCMAIN1.Progress("Logging Detail Changes")

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each rowSOTORDR2_orig As DataRow In dt.Rows
            Dim ORDR_LNO As Int64 = rowSOTORDR2_orig.Item("ORDR_LNO")
            Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
            If rowSOTORDR2 Is Nothing Then ' Line was Deleted
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowSOTORDR2_orig.Table.Columns(i).ColumnName
                    Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                    With rowSOTORDXR
                        .Item("REV_NO") = REV_NO
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("ORDR_NO") = ORDR_NO
                        .Item("ORDR_LNO") = ORDR_LNO
                        .Item("INIT_DATE") = LAST_DATE
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("OLD_VALUE") = rowSOTORDR2_orig.Item(COLUMN_NAME)
                        '.Item("NEW_VALUE") = ""
                        .Item("EMODE") = EntryMode
                        .Item("CONTEXT") = rowSOTORDR2_orig.Item("ITEM_CODE")
                    End With
                    dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                Next

                Check_Changed_Fields = True
            Else
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowSOTORDR2_orig.Table.Columns(i).ColumnName
                    If rowSOTORDR2.Item(COLUMN_NAME) & "" <> rowSOTORDR2_orig.Item(COLUMN_NAME) & "" Then
                        ' Value in Column was Changed
                        Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                        With rowSOTORDXR
                            .Item("REV_NO") = REV_NO
                            REV_LNO += 1
                            .Item("REV_LNO") = REV_LNO
                            .Item("ORDR_NO") = ORDR_NO
                            .Item("ORDR_LNO") = ORDR_LNO
                            .Item("INIT_DATE") = LAST_DATE
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("COLUMN_NAME") = COLUMN_NAME
                            .Item("OLD_VALUE") = rowSOTORDR2_orig.Item(COLUMN_NAME)
                            .Item("NEW_VALUE") = rowSOTORDR2.Item(COLUMN_NAME)
                            .Item("EMODE") = EntryMode
                            '  .Item("CONTEXT") = rowSOTORDR2_orig.Item("ITEM_CODE")
                        End With
                        dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                        Check_Changed_Fields = True
                    End If
                Next
            End If
        Next

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "", DataViewRowState.Added)
            Dim ORDR_LNO = rowSOTORDR2.Item("ORDR_LNO")
            ' For i As Integer = 0 To dt.Columns.Count - 1
            Dim COLUMN_NAME As String = "" ' dt.Columns(i).ColumnName
            Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
            With rowSOTORDXR
                .Item("REV_NO") = REV_NO
                REV_LNO += 1
                .Item("REV_LNO") = REV_LNO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = ORDR_LNO
                .Item("INIT_DATE") = LAST_DATE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("COLUMN_NAME") = COLUMN_NAME
                '.Item("OLD_VALUE") = ""
                .Item("NEW_VALUE") = "PO Line Added" ' rowSOTORDR2.Item(COLUMN_NAME)
                .Item("EMODE") = EntryMode
                .Item("CONTEXT") = rowSOTORDR2.Item("ITEM_CODE")
            End With
            dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
            Check_Changed_Fields = True
            'Next
        Next

        ASCMAIN1.Progress("")
        Return Check_Changed_Fields
    End Function

    Sub Record_Event(EVENT_CODE As String, EVENT_DESC As String)
        Dim rowSOTORDRE As DataRow = dst.Tables("SOTORDRE").NewRow
        rowSOTORDRE.Item("ORDR_NO") = ORDR_NO
        rowSOTORDRE.Item("INIT_DATE") = DATETIME_STAMP
        rowSOTORDRE.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowSOTORDRE.Item("EVENT_CODE") = EVENT_CODE
        rowSOTORDRE.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("SOTORDRE").Rows.Add(rowSOTORDRE)
    End Sub

    Private Sub optShowOrders_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShowOrders.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTORDRX()
    End Sub

    '#Region "VB6ICI"

    '    ' rotator order support
    '    ' print option in order entry

    '                qstd = Val(dynICTITEM1.Fields("ITEM_STD_PACK_SLS").Value & "")
    '                If (qstd = 0 And qord > 100) Or (qstd <> 0 And qord > qstd * 10) Then
    '                    MsgBox("Qty Entered is > 100 or > 10 * Std Pack. Please Check Qty", vbOKOnly, "Warning")
    '                End If

    '                Dim ITEM_STD_PACK_SLS As Long
    '                ITEM_STD_PACK_SLS = Val(dynICTITEM1.Fields("ITEM_STD_PACK_SLS").Value & "")
    '                Dim ITEM_SO_QTY_MIN As Integer
    '                ITEM_SO_QTY_MIN = Val(dynICTITEM1.Fields("ITEM_SO_QTY_MIN").Value & "")
    '                Dim SO_PARM_STD_PACK_MULT As Integer
    '                Dim SO_PARM_QTY_MAX As Integer
    '                If UID = "COS" Then
    '                    If ITEM_STD_PACK_SLS <> 0 Then
    '                        If Int(Val(grdSOWORDR2.Columns("ORDR_QTY").Text) / ITEM_STD_PACK_SLS) * ITEM_STD_PACK_SLS <> Val(grdSOWORDR2.Columns("ORDR_QTY").Text) Then
    '                            z = "Order Qty for this Item should be in Multiples of " & CStr(ITEM_STD_PACK_SLS) & "." & vbCr
    '                            z = z & "Do you want to override this rule?"
    '                            If MsgBox(z, vbYesNo + vbDefaultButton2, "Warning") = vbNo Then
    '                                Cancel = True
    '                                Exit Sub
    '                            Else
    '                                z = "Order Qty for this Item should be in Multiples of " & CStr(ITEM_STD_PACK_SLS)
    '                                If InputBox("Enter Password to Continue", z) <> dynSOTPARM1.Fields("SO_PARM_QTY_MAX_PWD").Value & "" Then
    '                                    Cancel = True
    '                                    Exit Sub
    '                                End If
    '                            End If
    '                        End If
    '                    End If

    '                    If ITEM_SO_QTY_MIN <> 0 Then
    '                        If Val(grdSOWORDR2.Columns("ORDR_QTY").Text) < ITEM_SO_QTY_MIN Then
    '                            BTN = True
    '                            z = "Minimum Order Qty for this Item is " & CStr(ITEM_SO_QTY_MIN) & "." & vbCr
    '                            z = z & "Do you want to override this rule?"
    '                            If MsgBox(z, vbYesNo + vbDefaultButton2, "Warning") = vbNo Then
    '                                BTN = False
    '                                Exit Sub
    '                            Else
    '                                z = "Minimum Order Qty for this Item is " & CStr(ITEM_SO_QTY_MIN)
    '                                If InputBox("Enter Password to Continue", z) <> dynSOTPARM1.Fields("SO_PARM_QTY_MAX_PWD").Value & "" Then
    '                                    BTN = False
    '                                    Exit Sub
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '                End If
    '        End Select
    '    End Sub

    '    Function Store_000000() As Boolean
    '        Dim dyn As OraDynaset

    '        Sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "'"
    '        Sql = Sql & " and CUST_STORE_NO = '000000'"
    '        dyn = OraD.CreateDynaset(Sql, 8&)
    '        Store_000000 = True
    '        If dyn.EOF Then
    '            iResponse = MsgBox("Would you like Store 000000 Created Automatically", vbYesNo, "Store 000000 Record is Not on File")
    '            If iResponse = vbNo Then
    '                Store_000000 = False
    '                Exit Function
    '            End If
    '            OraD.Parameters("CUST_CODE").Value = CUST_CODE
    '            dynARTCUST1.Refresh()
    '            dyn.AddNew()
    '            dyn.Fields("CUST_CODE").Value = CUST_CODE
    '            dyn.Fields("CUST_STORE_NO").Value = "000000"
    '            dyn.Fields("CUST_STORE_NAME").Value = dynARTCUST1.Fields("CUST_NAME").Value

    '            dyn.Fields("CUST_STORE_STATUS").Value = "A"
    '            dyn.Fields("SELL_CODE").Value = dynARTCUST1.Fields("SREP_CODE").Value

    '            dyn.Update()
    '        End If
    '        txtCode_LostFocus(1)

    '    End Function

    '#End Region


    Sub Load_Order_Form()
        'Dim FILENAME As String = ""
        'Dim dt As DataTable = Gembox_Import_Sheet_to_DataTable(0, FILENAME)
    End Sub

    Sub Set_Ship_To()
        'Synch_TABLE_NAME("SOTORDR5")
        Dim X As CurrencyManager = Me.BindingContext(dvwSOTORDR5)
        X.EndCurrentEdit()
        Dim rowSOTORDR5_ST As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, "ST"})
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, optShipTo.Value})
        For Each dc As DataColumn In dst.Tables("SOTORDR5").Columns
            If dc.ColumnName <> "ORDR_NO" And dc.ColumnName <> "CUST_ADDR_TYPE" Then
                rowSOTORDR5_ST.Item(dc.ColumnName) = rowSOTORDR5.Item(dc.ColumnName)
            End If
        Next
    End Sub

    Private Sub chkManualPrice_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkManualPrice.CheckedChanged
        If chkManualPrice.Checked Then
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTORDRI.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSOTORDR2.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTORDRI.DisplayLayout.Bands(0).Columns("ORDR_UNIT_PRICE").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub btnOrderForm_Click(sender As System.Object, e As System.EventArgs) Handles btnOrderForm.Click
        Using F As New TAC.SOFFORM1
            F.ORDR_FORM_CODE = ""
            F.PRICE_BASE_DPCT = PRICE_BASE_DPCT
            F.ShowDialog()
            If F.ORDR_FORM_CODE <> "" Then
                Add_Items(F.ORDR_FORM_CODE, F.dst.Tables("SOTFORM2"))
            End If
        End Using
    End Sub

    Sub Add_Items(ORDR_FORM_CODE As String, tbl As DataTable)
        If tbl.Select("ISNULL(ORDR_QTY,0)<>0").Length = 0 Then
            MsgBox("No Qty's Entered", MsgBoxStyle.OkOnly, "Cannot Add Items")
            Exit Sub
        End If

        For Each rowSOTFORM2 As DataRow In tbl.Select("ISNULL(ORDR_QTY,0)<>0", "ORDR_FORM_LNO")
            Add_grdSOTORDR2(rowSOTFORM2.Item("ITEM_CODE"), Val(rowSOTFORM2.Item("ORDR_UNIT_PRICE") & ""), Val(rowSOTFORM2.Item("ORDR_QTY") & ""))
        Next
        If grdSOTORDR2.ActiveRow IsNot Nothing Then
            grdSOTORDR2.ActiveRow.CancelUpdate()
        End If
        Sort_grdColumns(grdSOTORDR2, "ORDR_LNO")
    End Sub

    Function Add_grdSOTORDR2(ITEM_CODE As String, ORDR_UNIT_PRICE As Decimal, ORDR_QTY As Int64) As UltraWinGrid.UltraGridRow

        If grdSOTORDR2.ActiveRow IsNot Nothing AndAlso grdSOTORDR2.ActiveRow.IsAddRow Then
            grdSOTORDR2.ActiveRow = Nothing
        End If
        grdSOTORDR2.DisplayLayout.Bands(0).AddNew()
        With grdSOTORDR2.ActiveRow
            .Cells("ITEM_CODE").Value = ITEM_CODE
            .Cells("ORDR_UNIT_PRICE").Value = ORDR_UNIT_PRICE
            .Cells("ORDR_QTY").Value = ORDR_QTY
            .Update()
        End With
        Return grdSOTORDR2.ActiveRow
    End Function

    Sub Credit_Card()

        Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text.Trim
        Dim FRT_TERMS As String = MyBase.Absx1.txtFor("FRT_TERMS").Text.Trim
        Dim SHIP_VIA_CODE As String = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text.Trim
        Dim freightCost As Decimal = 0
        Dim rowSOTORDR1X As DataRow = Nothing

        ' Here are some rules
        If dst.Tables("SOTORDR1").Rows.Count = 0 Then
            MessageBox.Show("Invalid or Missing Sales Order Number.", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        rowSOTORDR1X = dst.Tables("SOTORDR1").Rows(0)

        If rowSOTORDR1X.Item("CCPA_NO") & String.Empty <> String.Empty Then
            MessageBox.Show("This sales order has an existing credit card authorization. You are not permitted to authorize additional funds.", _
                "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If rowSOTORDR1X.Item("ORDR_SOURCE") & String.Empty = "W" Then
            MessageBox.Show("Web sales credit card authorization was processed on the website. You are not permitted to authorize additional funds.", _
                "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Not ",O,P,".Contains(rowSOTORDR1X.Item("ORDR_STATUS") & String.Empty) Then
            MessageBox.Show("Only Open and In-Pick statuses can perform a credit card Authorization. If the order has been shipped, you may charge the credit card in Customer Inquiry.", _
                "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If EntryMode = "V" Then
            ASCMAIN1.sql = "Select Count (*) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS in ('O','P','F')"
            If Val(ASCDATA1.GetDataValue) > 1 Then
                MessageBox.Show("You Cannot perform credit card processing on a Multiple Order Group", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        EMsg = String.Empty
        If FRT_TERMS.Length > 0 Then
            If ASCDATA1.GetDataRow("select * from astcode1 where TABLE_NAME = 'ARTCUST1' AND COLUMN_NAME = 'FRT_TERMS' AND T_CODE = '" & FRT_TERMS & "'") Is Nothing Then
                EMsg &= vbCr & "Freight Terms are required to process a credit card."
            End If
        Else
            EMsg &= vbCr & "Freight Terms are required to process a credit card."
        End If

        If SHIP_VIA_CODE.Length > 0 Then
            If ASCDATA1.GetDataRow("SELECT * FROM SOTSVIA1 WHERE SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'") Is Nothing Then
                EMsg &= vbCr & "A Valid Ship Via Code is required for credit card processing."
            End If
        Else
            EMsg &= vbCr & "Ship Via Code is required for credit card processing."
        End If

        'If Not IsDate(MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value) Then
        '    EMsg &= vbCr & "Ship date is required for credit card processing."
        'End If

        'If CDate(MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value) < DateTime.Now Then
        '    If EntryMode <> "V" Then
        '        EMsg &= vbCr & "Ship date must be greater equal to today for credit card processing."
        '    End If
        ' End If

        If EMsg.Length > 0 Then
            MessageBox.Show(EMsg, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim rowSOTCARR1 As DataRow = ASCDATA1.GetDataRow("select sotcarr1.carrier_type" _
                                                         & " from sotsvia1, sotcarr1" _
                                                         & " where sotsvia1.carrier_code = sotcarr1.carrier_code" _
                                                         & " and ship_via_code = :PARM1", "V", New Object() {SHIP_VIA_CODE})


        ' Fedex, UPS and similar pay for freight when freight terms of PPA 
        If rowSOTCARR1.Item("CARRIER_TYPE") & String.Empty = "U" AndAlso FRT_TERMS.ToUpper = "PPA" Then
            'ASCMAIN1.Progress("Analyzing freight costs", "")
            ' try to get an estimated freight charge
            ' New Rule 1/24/2013. 20% or $20 the greater of the two
            freightCost = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & String.Empty) * 0.2
            If freightCost <= 20 Then
                freightCost = 10
            End If
        End If

        Dim ORDR_TOTAL_AMT As Decimal = 0

        ORDR_TOTAL_AMT += Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & String.Empty)
        ORDR_TOTAL_AMT += Val(rowSOTORDR1.Item("ORDR_FREIGHT") & String.Empty)
        ORDR_TOTAL_AMT += Val(rowSOTORDR1.Item("ORDR_STAX") & String.Empty)
        ORDR_TOTAL_AMT += freightCost

        If ORDR_TOTAL_AMT <= 0 Then
            MessageBox.Show("You cannot charge $0.00 for sales Order No: " & ORDR_NO, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If Not ASCMAIN1.Logical_Lock("ARTCUSTC", CUST_CODE, , , , 1) Then Exit Sub
        If Not ASCMAIN1.Logical_Open("ARTCCPA1", "*", , , , 1) Then Exit Sub
        Get_PARM("SOTPARM1")

        Using frmCCProcessor As New TAC.TAFCARDF(Me)
            frmCCProcessor.test_mode = ROWs("SOTPARM1").Item("SO_PARM_CC_TEST_MODE") & String.Empty = "1"
            frmCCProcessor.CUST_CODE = CUST_CODE
            frmCCProcessor.CCPA_REASON = "O"
            frmCCProcessor.ORDR_NO = ORDR_NO
            frmCCProcessor.TRAN_TYPE = "A"

            With frmCCProcessor.rowARTCCPA1
                .Item("CUST_CODE") = CUST_CODE
                .Item("CCPA_AMT") = ORDR_TOTAL_AMT
                .Item("CCPA_NOTE") = "Credit Card Order"
            End With

            Try
                frmCCProcessor.ShowDialog()
                Dim row As DataRow = ASCDATA1.GetDataRow("Select * from ARTCCPA1 where CCPA_NO = :PARM1", "V", New Object() {frmCCProcessor.CCPA_NO & String.Empty})
                If row IsNot Nothing AndAlso row.Item("CCPA_STATUS") & String.Empty = "T" Then
                    rowSOTORDR1.Item("CCPA_NO") = frmCCProcessor.CCPA_NO & String.Empty
                    rowSOTORDR1.Item("CC_TRANS_ID") = row.Item("TRANS_ID") & String.Empty
                    ASCDATA1.ExecuteSQL("UPDATE SOTORDR1 SET CCPA_NO = '" & rowSOTORDR1.Item("CCPA_NO") & "'" _
                                        & ", CC_TRANS_ID = '" & rowSOTORDR1.Item("CC_TRANS_ID") & "' WHERE ORDR_NO = '" & ORDR_NO & "'")
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        End Using

        ASCMAIN1.MultiTask_Release(, , 1)
        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub


    Private Function GetFreightCosts() As Decimal

        Dim freightCosts As Decimal = 0

        Try
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", MyBase.Absx1.txtFor("WHSE_CODE").Text)

            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text)
            Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE")
            Dim rowSOTCARR1 As DataRow = LookUp("SOTCARR1", CARRIER_CODE)
            Dim rowSOTCARR3 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTCARR3 WHERE CARRIER_CODE = :PARM1", "V", New Object() {CARRIER_CODE})

            Dim ShippingLabelDirectory As String = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)
            Dim isInternationalShipment As Boolean = False
            If rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty = String.Empty Then
                isInternationalShipment = False
            Else
                isInternationalShipment = Not (rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.ToUpper.StartsWith("US")
            End If

            ' See if the consumer uses their own account to pay for  freight
            If rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" Then
                Return 0
            End If

            Select Case rowSOTCARR1.Item("PROVIDER_TYPE")
                Case WHCSHIP1.ProviderTypeFedex
                    If Not isInternationalShipment Then
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpress)
                        clsShip.RequestedServiceType = nsoftware.InShip.ServiceTypes.stFedExGround
                    Else
                        clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.FederalExpressInternational)
                        clsShip.RequestedServiceType = nsoftware.InShip.ServiceTypes.stFedExInternationalGround
                    End If
                Case WHCSHIP1.ProviderTypeUPS
                    clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.UPS)
                    clsShip.RequestedServiceType = nsoftware.InShip.ServiceTypes.stUPSGround
                Case WHCSHIP1.ProviderTypeUSPS
                    clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.USPS)
                    Return 0
                Case WHCSHIP1.ProviderTypeCanada
                    clsShip = New TAC.WHCSHIP1(WHCSHIP1.ServiceProviders.CanadaPost)
                    Return 0
                Case Else
                    Return 0
            End Select

            With clsShip.Sender
                .FirstName = rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty
                .MiddleInitial = ""
                .LastName = ""
                .Address1 = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
                .Address2 = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
                .City = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
                .State = rowICTWHSE1.Item("WHSE_STATE") & String.Empty
                .ZipCode = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
                .CountryCode = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
                .Company = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                .Phone = rowICTWHSE1.Item("WHSE_PHONE") & String.Empty

                .IsResidental = False
                .IsPOBox = False

                If .Company.Length = 0 Then
                    .Company = (.FirstName & " " & .LastName).ToString.Trim
                End If
            End With

            With clsShip.Recipient
                .FirstName = rowSOTORDR5.Item("CUST_CONTACT") & String.Empty
                .MiddleInitial = ""
                .LastName = "" 'txtFromLastName.Text.Trim
                .Address1 = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                .Address2 = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                .City = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                .State = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                .ZipCode = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                .CountryCode = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
                If .CountryCode.Trim = String.Empty OrElse .CountryCode.Trim.ToUpper.StartsWith("US") Then
                    .CountryCode = "US"
                End If
                .Company = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                .Phone = "1234567890"
            End With

            Dim weight As Decimal = 0
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "", DataViewRowState.CurrentRows)
                If Val(rowSOTORDR2.Item("ITEM_WEIGHT") & String.Empty) = 0 Then
                    ' Assume 4 ounces iof there is no weight.
                    weight += 0.25 + Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                Else
                    weight += Val(rowSOTORDR2.Item("ITEM_WEIGHT") & String.Empty) + Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                End If
            Next

            shipPackageDetailList.Clear()
            Dim shipPackageDetail As New nsoftware.InShip.PackageDetail
            With shipPackageDetail
                .PackagingType = nsoftware.InShip.TPackagingTypes.ptYourPackaging
                .Weight = Convert.ToInt32(weight)
                .Length = 17.5
                .Width = 17.5
                .Height = 13.5
                .Id = "00000001"
            End With

            shipPackageDetailList.Add(shipPackageDetail)

            Try
                If ASCMAIN1.Running_in_VS Then
                    ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "C:\")
                End If
                If ShippingLabelDirectory.Length > 0 Then
                    If Not My.Computer.FileSystem.DirectoryExists(ShippingLabelDirectory) Then
                        My.Computer.FileSystem.CreateDirectory(ShippingLabelDirectory)
                    End If
                End If
            Catch ex As Exception
                ShippingLabelDirectory = String.Empty
            End Try

            If ShippingLabelDirectory.Length > 0 AndAlso Not ShippingLabelDirectory.EndsWith("\") Then
                ShippingLabelDirectory = ShippingLabelDirectory & "\"
            End If

            ' Credentials
            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

            If isInternationalShipment Then
                If Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty) > 0 Then
                    clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty)
                End If
            Else
                If Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty) > 0 Then
                    clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
                End If
            End If

            If isInternationalShipment Then
                clsShip.CommodityDetailList.Clear()
                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("", "", DataViewRowState.CurrentRows)
                    Dim CommodityDetail As New nsoftware.InShip.CommodityDetail
                    CommodityDetail.Description = rowSOTORDR2.Item("ITEM_DESC") & String.Empty
                    CommodityDetail.NumberOfPieces = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                    CommodityDetail.Quantity = Val(rowSOTORDR2.Item("ORDR_QTY") & String.Empty)
                    CommodityDetail.QuantityUnit = "EA"
                    CommodityDetail.UnitPrice = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                    CommodityDetail.Weight = rowSOTORDR2.Item("ITEM_WEIGHT")
                    CommodityDetail.Manufacturer = "US"
                    clsShip.CommodityDetailList.Add(CommodityDetail)
                Next
                clsShip.TotalCustomsValue = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & String.Empty)
            End If

            With clsShip
                .EzshipLabelImage = nsoftware.InShip.EzshipLabelImageTypes.itEltron
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = "X"
                If IsDate(MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value) AndAlso MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value >= DateTime.Now Then
                    .ShipDate = MyBase.Absx1.dteFor("ORDR_SHIP_DATE").Value
                Else
                    .ShipDate = DateTime.Now
                End If

            End With

            clsShip.PackageDetailList = shipPackageDetailList
            clsShip.GetRates()

            freightCosts = 0

            For Each Charge As KeyValuePair(Of Integer, Decimal) In clsShip.ShipmentListCharge
                freightCosts += Val(Charge.Value & String.Empty)
            Next

        Catch ex As Exception
            freightCosts = -1
        Finally

        End Try

        Return freightCosts

    End Function

    Sub Setup_Multiple_Order_Grid(show_header_fields As Boolean)
        For Each COLUMN_NAME As String In New String() _
            {"ORDR_NO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_CUST_PO", "ORDR_DEPT", "ORDR_ADDR_TYPE_ST", "ORDR_HOLD", "ORDR_OVERRIDE_NOT_ALLOCATED", "ORDR_STATUS"}
            grdSOTORDRB.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not show_header_fields
            ' "CUST_STORE_NAME", "CUST_STORE_LOCATION"
        Next
    End Sub

    'Sub Log_Changes( _
    '    row As DataRow, _
    '    TABLE_NAME As String, _
    '    ByRef Check_Changed_Fields As Boolean, _
    '    ByRef REV_LNO As Integer, _
    '    LAST_DATE As Date)

    '    For i As Integer = 0 To row.Table.Columns.Count - 1
    '        Dim COLUMN_NAME As String = dst.Tables(TABLE_NAME).Columns(i).ColumnName

    '        If row.Item(COLUMN_NAME) & "" _
    '        <> row.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
    '            Check_Changed_Fields = True
    '            ASCMAIN1.Progress("-", COLUMN_NAME)
    '            Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
    '            With rowSOTORDXR
    '                .Item("REV_NO") = REV_NO
    '                REV_LNO += 1
    '                .Item("REV_LNO") = REV_LNO
    '                .Item("ORDR_NO") = ORDR_NO
    '                .Item("ORDR_LNO") = 0
    '                .Item("INIT_DATE") = LAST_DATE
    '                .Item("INIT_OPER") = ASCMAIN1.USER_ID
    '                .Item("COLUMN_NAME") = COLUMN_NAME
    '                .Item("OLD_VALUE") = row.Item(COLUMN_NAME, DataRowVersion.Original)
    '                .Item("NEW_VALUE") = row.Item(COLUMN_NAME)
    '                .Item("EMODE") = EntryMode
    '            End With
    '            dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
    '            Check_Changed_Fields = True
    '        End If
    '    Next i
    'End Sub

    Private Sub chkEdit_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkEdit.CheckedChanged
        Setup_SHIPTO_Edit()
        chkEdit.Visible = Not chkEdit.Checked
    End Sub

    Sub Setup_SHIPTO_Edit()
        Set_Read_Only(grpSHIPTO, Not chkEdit.Checked)
        Set_Read_Only_for_ctl(chkEdit, False)
        chkSaveToMasterFile.Checked = False
        chkSaveToMasterFile.Visible = chkEdit.Checked
        chkSaveToMasterFile.Checked = True
        chkSaveToMasterFile.Enabled = False
    End Sub

    Sub Update_Ship_To_Address() ' (REV_NO As Integer, ByRef REV_LNO As Integer)

        Synch_TABLE_NAME("SOTORDR5")

        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'ST'")(0)
        TAC.SOCMAIN1.Record_Event_SOTORDRE(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "ADRCHG", "Ship-To Address Changed")
        TAC.SOCMAIN1.Log_Changes(Me, ORDR_NO, rowSOTORDR5, "SOTORDR5", False, REV_NO, REV_LNO, Now)

        ' Update_Record_TDA("SOTORDR5")

        If chkSaveToMasterFile.Checked Then
            dst.Tables("ARTCUST2").Rows.Clear()
            Dim rowARTCUST2 As DataRow = Fill_Record("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

            Dim COLUMN_NAMEs() As String = {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY", "CUST_CONTACT"}
            For Each COLUMN_NAME As String In COLUMN_NAMEs
                rowARTCUST2.Item(Replace(COLUMN_NAME, "CUST_", "CUST_STORE_")) = rowSOTORDR5.Item(COLUMN_NAME)
            Next

            Update_Record_TDA("ARTCUST2")
        End If

    End Sub

    Private Sub grdSOTORDRB_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTORDRB.BeforeCellUpdate
        If e.Cell.Column.Key.StartsWith("QTY_") Then
            Dim row As DataRow = dst.Tables("SOTORDRB").Rows.Find(e.Cell.Row.Cells("ORDR_NO").Value)
            If row.Item(e.Cell.Column.Key) & "" = "" Then
                e.Cancel = True
            End If
        End If
    End Sub

    Sub Totals_for_SOTORDRI()
        For Each rowSOTORDRI As DataRow In dst.Tables("SOTORDRI").Select("")
            Dim COL As Integer = Val(rowSOTORDRI.Item("COL") & "")
            Dim QTY As Int64 = Val(dst.Tables("SOTORDRB").Compute("SUM(QTY_" & Format(COL, "000") & ")", "") & "")
            rowSOTORDRI.Item("ORDR_QTY_OPEN") = QTY
            ' DON'T WE NEED A SWITCH TO UPDATE ORDR_QTY OR ORDR_QTY_OPEN?
        Next
    End Sub

    Private Sub tabMultiOrder_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMultiOrder.SelectedTabChanged
        If tabMultiOrder.SelectedTab.Key = "Items" Then
            Totals_for_SOTORDRI()
        End If
    End Sub

    Private Sub grdSOTORDRX_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTORDRX.InitializeLayout

    End Sub

    Private Sub grdSOTORDRX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRX.InitializeRow
        If e.Row.Cells("ORDR_HOLD").Value & "" = "1" Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub
End Class