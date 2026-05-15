Imports DPayments.DShippingSDK

Public Class SOFSHIP0

    ' proceed prereq - after maintenance, or confirmation,need to verify that auth amt on credit check and credit card has not been violated
    Private expSOTPICK1 As New Dictionary(Of String, String)
    Private tblTATSTATE As DataTable
    Private tblSHTSHIPS As DataTable
    Private sqlSOTINTL1 As String = String.Empty
    Private requestShippingOptions As Boolean = False

    Private validDates() As Date = TAC.SOCMAIN1.Validate_Invoice_Date(Nothing, 0, 1, Nothing)

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name
    Dim SHIP_BOL_NOs As New List(Of String)
    Dim ORDR_GROUP_NO As String
    Dim ORDR_CUST_PO As String
    Dim rowARTCUST1 As DataRow
    Dim rowSOTSHIP0 As DataRow
    Dim rowSOTSHIP0_ORIG As DataRow
    Dim clsPrice_Change As Price_Change = Nothing
    Dim sqlSOTPICK1 As String
    Dim sqlSOTPICK2 As String
    Dim sqlSOTSHIPX As String
    Dim edi_customer As Boolean
    Dim edi856_customer As Boolean
    Dim edi_order As Boolean
    Dim ORDR_SOURCE As String
    Dim SOTSHIP0 As String
    Dim CURR_CODE As String
    Dim CURR_EXCH_RATE As Decimal
    Dim MaintenanceMode As Boolean = False

    Dim ORDR_SHIP_DATE As Date
    Dim ORDR_CANCEL_DATE As Date

    Dim dvwSOTORDR5 As DataView

    Dim SOTSHIPX As String
    Dim ASW As New Dictionary(Of String, String)

    Private commonCarrier As Boolean = False
    Private CreditCardProcessor As TAC.TAFCARDF
    Private RecreateLabel As Boolean = False
    Private refreshScreen As Boolean = True

    Private WithEvents ultraComboPackage As Infragistics.Win.UltraWinGrid.UltraCombo = New Infragistics.Win.UltraWinGrid.UltraCombo

#End Region

#Region "ABS Standard Routines"

    Private Sub SOFSHIP0_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If ScannerInUse AndAlso Not txtPickNo.Focused Then
            txtPickNo.Focus()
        End If
    End Sub

    ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Check_InquiryMode()
        Dim shipPackageDetail As New DPayments.DShippingSDK.PackageDetail

        SOTSHIP0 = ASCMAIN1.Temp_Table("Select SHIP_BOL_NO from SOTSHIP1 where ROWNUM < 1")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIP0 & " Add Primary Key (SHIP_BOL_NO)")

        If MaintenanceMode Then
            ASCMAIN1.sql = "Select SHIP_BOL_NO, '1' SEL, '1' EDI856, '1' SHIP_CART_REQD from SOTSHIP1 where ROWNUM < 1"
            SOTSHIPX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add Primary Key (SHIP_BOL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add SHPCHG_REQ_NO VARCHAR2(10)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTSHIPX & " Add CUST_CODE VARCHAR2(10)")
        End If

        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")
        Get_PARM("ASTPARM1")

        ' Only Common Carriers on this screen
        ' 01/11/2013
        ' Debbie needs to Setup some Fedex/UPS in her screen SOFORDRB
        ' so I added a check on SOTSHIP1.BILL_OF_LADING_NO
        With dst
            sqlSOTSHIPX = "Select SOTSHIP1.*" _
                & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" _
                & " from SOTSHIP1,SOTORDR0,SOTSVIA1,SOTCARR1" _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & " and SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE" & vbCrLf _
                & " and SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE" & vbCrLf _
                & " and SOTSHIP1.BILL_OF_LADING_NO IS NULL"

            If Not InquiryMode Then
                sqlSOTSHIPX &= "" _
                    & "   and SOTSHIP1.SHIP_STATUS = 'P'" _
                    & "   and SOTSHIP1.SHIP_PICK_PRINTED is Not Null"
            End If

            ' Needed to Bill Amazon Shipments from another Warehouse
            If Not (ASCMAIN1.DBS_COMPANY = "AHA" OrElse ASCMAIN1.DBS_SERVER = "AHA") Then
                sqlSOTSHIPX &= " and SOTCARR1.CARRIER_TYPE = 'U'"
            Else
                sqlSOTSHIPX &= " and SOTSHIP1.WHSE_CODE IN (Select WHSE_CODE from ICTWHSE1 WHERE LP_WHSE_ID IS NULL)"
            End If

            ASCMAIN1.sql = sqlSOTSHIPX
            Create_TDA(.Tables.Add, "SOTSHIPX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*")
            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "", 1)
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP0", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTCART1", "*")
            .Tables("SOTCART1").Columns.Add("INSURANCE", GetType(System.Decimal))
            .Tables("SOTCART1").Columns("INSURANCE").DefaultValue = 0

            .Tables("SOTCART1").Columns.Add("REFERENCE1", GetType(System.String))
            .Tables("SOTCART1").Columns("REFERENCE1").DefaultValue = String.Empty
            .Tables("SOTCART1").Columns("REFERENCE1").MaxLength = 20

            .Tables("SOTCART1").Columns.Add("REFERENCE2", GetType(System.String))
            .Tables("SOTCART1").Columns("REFERENCE2").DefaultValue = String.Empty
            .Tables("SOTCART1").Columns("REFERENCE2").MaxLength = 20

            .Tables("SOTCART1").Columns.Add("WIDTH", GetType(System.Int16))
            .Tables("SOTCART1").Columns("WIDTH").DefaultValue = 0
            .Tables("SOTCART1").Columns.Add("LENGTH", GetType(System.Int16))
            .Tables("SOTCART1").Columns("LENGTH").DefaultValue = 0
            .Tables("SOTCART1").Columns.Add("HEIGHT", GetType(System.Int16))
            .Tables("SOTCART1").Columns("HEIGHT").DefaultValue = 0

            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
              & " from SOTCART2,SOTCART1 where SOTCART1.CART_NO = SOTCART2.CART_NO"
            Create_TDA(.Tables.Add, "SOTCART2", "**", 0)

            'ASCMAIN1.sql = "Select CUST_CODE, EDI_DOC_NO, EDI_STATUS" _
            '    & " from EDTTRPM1" _
            '    & " where CUST_CODE = :PARM1 AND COMPANY_CODE = '" & ASCMAIN1.CLIENT_CODE & "'"

            'Create_TDA(.Tables.Add, "EDTTRPMC", "**", 0, False, "V", 2)
            '.Tables("EDTTRPM1").PrimaryKey = New DataColumn() {.Tables("EDTTRPM1").Columns("CUST_CODE"), _
            '                                                   .Tables("EDTTRPM1").Columns("EDI_DOC_NO")}

            ASCMAIN1.sql = " Select EDTTRPM1.CUST_CODE, EDTTRPM1.EDI_DOC_NO, EDTTRPM1.EDI_STATUS" _
             & "  from EDTTRPM1, EDTTRPMC" _
             & "  where EDTTRPM1.EDI_TP_ID = EDTTRPMC.EDI_TP_ID" _
             & "  and EDTTRPM1.EDI_OUR_ID = EDTTRPMC.EDI_OUR_ID" _
             & "  and EDTTRPMC.COMPANY_CODE = '" & ASCMAIN1.SOLUTION & "'" _
             & "  and EDTTRPM1.CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "EDTTRPMC", "**", 0, False, "V", 2)

            With .Tables.Add("SOTCARTX")
                .Columns.Add("PICK_NO")
                .Columns.Add("ORDR_NO")
                .Columns.Add("ORDR_LNO", GetType(System.Int64))
                .Columns.Add("PICK_QTY_CONF", GetType(System.Int64), "")
                .Columns.Add("QTY_PACKED", GetType(System.Int64), "")
                .PrimaryKey = New DataColumn() {.Columns("PICK_NO"), .Columns("ORDR_NO"), .Columns("ORDR_LNO")}
            End With

            sqlSOTPICK1 = "Select SOTPICK1.*" & vbCrLf _
                & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_NO_WEB, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
                & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FREIGHT" & vbCrLf _
                & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND, SOTORDR1.CCPA_NO CCPA_NO_ORDR" & vbCrLf _
                & " from SOTPICK1, SOTORDR1, SOTSHIP1"
            ASCMAIN1.sql = sqlSOTPICK1 & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")
            dst.Tables("SOTPICK1").Columns.Add("SELECTED")
            'dst.Tables("SOTPICK1").Columns.Add("PPA_FREIGHT", GetType(System.Decimal))
            'dst.Tables("SOTPICK1").Columns("PPA_FREIGHT").DefaultValue = 0
            dst.Tables("SOTPICK1").Columns.Add("OUR_FREIGHT", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns("OUR_FREIGHT").DefaultValue = 0
            dst.Tables("SOTPICK1").Columns.Add("ORDR_STAX", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns.Add("STAX_CODE", GetType(System.String))

            Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")

            sqlSOTPICK2 = "Select SOTPICK2.*, " & vbCrLf _
                & " SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, " & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTSHIP1" & vbCrLf
            ASCMAIN1.sql = sqlSOTPICK2 & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK2", "**")
            With .Tables("SOTPICK2").Columns
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))
            End With

            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            With .Tables("SOTPICK1").Columns
                .Add("PICK_QTY", GetType(System.Int64))
                .Add("PICK_QTY_CONF", GetType(System.Int64))
                .Add("PICK_QTY_CANC", GetType(System.Int64))
                .Add("PICK_QTY_BACK", GetType(System.Int64))
                .Add("PICK_AMT", GetType(System.Decimal))
                .Add("PICK_AMT_CONF", GetType(System.Decimal))
                .Add("PICK_AMT_CANC", GetType(System.Decimal))
                .Add("PICK_AMT_BACK", GetType(System.Decimal))
            End With

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")

            Create_Relation("SOTCARTX", "SOTPICK2", "PICK_NO,ORDR_NO,ORDR_LNO")
            Create_Relation("SOTCARTX", "SOTCART2", "PICK_NO,ORDR_NO,ORDR_LNO")

            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_CALC", GetType(System.Int64))
            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_ORIG", GetType(System.Int64), "SUM(CHILD.QTY_PACKED_ORIG)")

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_WGT_CALC", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns.Add("PICK_CNT_CARTONS_CALC", GetType(System.Int64))
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_UNITS_CALC", GetType(System.Int64))

            With .Tables.Add("SOTCONFT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("KEY")}
            End With

            Create_TDA(.Tables.Add, "SOTSHIP3", "*")
            Create_TDA(.Tables.Add, "SOTSHIP4", "*")
            Create_TDA(.Tables.Add, "SOTSHIP6", "*")

            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)
            .Tables("SOTORDR5").Columns("CUST_ADDR_CODE").MaxLength = 10

            Create_TDA(.Tables.Add, "TATEVNT1", "*")
            Create_TDA(.Tables.Add, "SOTORDXR", "*")

            ' Credit Card Processing
            Create_TDA(.Tables.Add, "ARTCCPA1", "*")
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(.Tables.Add, "ARTCCPDA", "*")

            ' Shipping Label
            Create_TDA(.Tables.Add, "WHTSHIP1", "*")
            Create_TDA(.Tables.Add, "WHTSHIP2", "*")
            Create_TDA(.Tables.Add, "WHTSHIP3", "*")
            Create_TDA(.Tables.Add, "WHTSHIP5", "*")
            Create_TDA(.Tables.Add, "WHTSHIPC", "*")
            Create_TDA(.Tables.Add, "WHTSHIPS", "*")
            Create_TDA(.Tables.Add, "WHTSHIPP", "*")

            ' Carrier Tables
            Create_TDA(.Tables.Add, "SOTCARR1", "*")
            Create_TDA(.Tables.Add, "SOTCARR2", "*")
            Create_TDA(.Tables.Add, "SOTCARR3", "*")

            Fill_Records("SOTCARR1", "", True, "SELECT * FROM SOTCARR1")
            Fill_Records("SOTCARR2", "", True, "SELECT * FROM SOTCARR2")
            Fill_Records("SOTCARR3", "", True, "Select SOTCARR3.*, SOTCARR1.CARRIER_DESC From SOTCARR3, SOTCARR1 Where SOTCARR3.CARRIER_CODE = SOTCARR1.CARRIER_CODE")

            ASCMAIN1.sql = "SELECT SOTSVIA1.*, SOTCARR1.CARRIER_TYPE" _
                    & " FROM SOTSVIA1, SOTCARR1" _
                    & " WHERE SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE"
            Create_TDA(.Tables.Add, "SOTSVIA1", "**", 1, False, "", 1)
            Fill_Records("SOTSVIA1", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE" _
                & " FROM SOTORDR0, SOTSHIP1" _
                & " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO " _
                & " AND SOTSHIP1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTORDRX", ASCMAIN1.sql, 0, False, "V", 0)

            ASCMAIN1.sql = " SELECT DISTINCT SOTINVH1.ORDR_NO, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UPC_CODE, ICTITEM1.COUNTRY_CODE, " _
                & " ICTITEM1.ITEM_MATL_DESC, TATCNTRY.COUNTRY_NAME " _
                & " FROM ICTITEM1, TATCNTRY, SOTINVH2, SOTINVH1" _
                & " WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
                & " AND ICTITEM1.COUNTRY_CODE = TATCNTRY.COUNTRY_CODE (+) " _
                & " AND SOTINVH2.ITEM_CODE = ICTITEM1.ITEM_CODE"
            sqlSOTINTL1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTINTL1", ASCMAIN1.sql, 0, False, String.Empty, 0)

            ASCMAIN1.sql = " Select SOTCARR1.CARRIER_CODE, SOTCARR1.CARRIER_DESC, SOTCARR2.CARRIER_PROD_CODE, SOTCARR2.CARRIER_PROD_DESC" _
             & " , SOTSVIA1.SHIP_VIA_CODE, SOTSVIA1.SHIP_VIA_DESC, SOTCARR2.CARRIER_PROD_DESC TRANSIT_TIME, SOTCARR1.CARRIER_MIN_INS LIST_PRICE, SOTCARR1.CARRIER_MIN_INS OUR_PRICE" _
             & " FROM SOTCARR1, SOTCARR2, SOTSVIA1" _
             & " WHERE ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTCARRX", ASCMAIN1.sql, 0, False, String.Empty, 0)
            .Tables("SOTCARRX").Columns("SHIP_VIA_CODE").MaxLength = 100

        End With

        tblTATSTATE = ASCDATA1.GetDataTable("SELECT * FROM TATSTATE WHERE region_code is not null", "TATSTATE")
        tblSHTSHIPS = ASCDATA1.GetDataTable("SELECT * FROM SHTSHIPS", "SHTSHIPS")


        With ultraComboPackage.DisplayLayout.Bands(0)

            ultraComboPackage.Font = grdSOTCART1.Font
            ultraComboPackage.AutoCompleteMode = Infragistics.Win.AutoCompleteMode.Default
            ultraComboPackage.DropDownStyle = UltraWinGrid.UltraComboStyle.DropDownList

            .Columns.Add("PKG_CODE")
            .Columns("PKG_CODE").Header.Caption = "Code"
            .Columns("PKG_CODE").Width = 75

            .Columns.Add("PKG_DESC")
            .Columns("PKG_DESC").Header.Caption = "Desc"
            .Columns("PKG_DESC").Width = 75

            .Columns.Add("PKG_D")
            .Columns("PKG_D").Header.Caption = "L x W x H"
            .Columns("PKG_D").Width = 200

        End With

        ultraComboPackage.DataSource = ASCDATA1.GetDataTable("SELECT PKG_CODE, PKG_DESC, PKG_L || ' x ' ||  PKG_W || ' x ' || PKG_H PKG_D FROM WHTPKGM1")
        ultraComboPackage.ValueMember = "PKG_CODE"
        ultraComboPackage.DisplayMember = "PKG_DESC"
        grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_CODE").EditorComponent = ultraComboPackage

        grdSOTSHIPX.DataSource = dst.Tables("SOTSHIPX")
        grdSOTCONFT.DataSource = dst.Tables("SOTCONFT")
        grdSOTPICK1.DataSource = dst.Tables("SOTPICK1")
        grdSOTPICK2.DataSource = dst.Tables("SOTPICK2")
        grdSOTCART1.DataSource = dst.Tables("SOTCART1")
        grdSOTCART2.DataSource = dst.Tables("SOTCART2")
        grdSOTSHIP1.DataSource = dst.Tables("SOTSHIP1")
        grdSOTCARRX.DataSource = dst.Tables("SOTCARRX")

        dvwSOTORDR5 = New DataView(dst.Tables("SOTORDR5"), "CUST_ADDR_TYPE = 'ST'", "", DataViewRowState.CurrentRows)
        Bind_Controls(grpSHIPTO, "SOTORDR5", dvwSOTORDR5)

        Bind_Controls(txt3PAccountNo, "WHTSHIP3")
        Bind_Controls(txt3pCountry, "WHTSHIP3")
        Bind_Controls(txt3PZipCode, "WHTSHIP3")

        grdSOTSHIPX.DisplayLayout.UseFixedHeaders = True
        With grdSOTSHIPX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SHIP_BOL_NO", "CUST_CODE", "ORDR_CUST_PO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdSOTPICK1.DisplayLayout.UseFixedHeaders = True
        With grdSOTPICK1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SELECTED", "PICK_NO", "CUST_STORE_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdSOTPICK1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If Not MaintenanceMode And New String() {"SELECTED", "PICK_CNT_CARTONS", "PICK_FREIGHT", "PICK_TOTAL_WGT", "ORDR_INV_COMMENT"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    If New String() {"SELECTED"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.DarkGoldenrod
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    Else
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    End If
                ElseIf New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PICK_AMT", "PICK_AMT_CONF", "PICK_AMT_CANC", "PICK_AMT_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    'gcol.Format = "#,##0.00"
                ElseIf New String() {"ORDR_STAX", "STAX_CODE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightCoral
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
            If MaintenanceMode Then
                .Columns("PICK_QTY_CONF").Header.Caption = "uRevised"
                .Columns("PICK_AMT_CONF").Header.Caption = "$Revised"
            End If
        End With

        grdSOTPICK2.DisplayLayout.UseFixedHeaders = True
        With grdSOTPICK2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"PICK_LNO", "ITEM_CODE", "ITEM_DESC", "PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdSOTPICK2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If Not MaintenanceMode And New String() {"PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                ElseIf MaintenanceMode And New String() {"PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
            If MaintenanceMode Then
                .Columns("PICK_QTY_CONF").Header.Caption = "Revised"
                .Columns("PICK_QTY_BACK").Hidden = True
            End If
        End With

        grdSOTCART1.DisplayLayout.UseFixedHeaders = True
        With grdSOTCART1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CART_NO"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTCART1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"CART_FREIGHT", "CART_TOTAL_WGT_ACTUAL", "PKG_CODE", "PACKAGING_TYPE", "INSURANCE", _
                                 "WIDTH", "LENGTH", "HEIGHT", "CART_SEQ", "REFERENCE1", "REFERENCE2"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        grdSOTCART2.DisplayLayout.UseFixedHeaders = True
        With grdSOTCART2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CART_LNO", "ITEM_CODE", "QTY_PACKED"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With
        With grdSOTCART2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"QTY_PACKED"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        Create_Summary(grdSOTSHIPX, "SHIP_BOL_NO", "Count")

        Create_Summary(grdSOTPICK2, "PICK_LNO", "Count")
        Create_Summary(grdSOTPICK2, New String() _
            {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"})

        Create_Summary(grdSOTCART1, "CART_NO", "Count")
        Create_Summary(grdSOTCART1, New String() _
            {"CART_FREIGHT", "CART_TOTAL_UNITS", "CART_TOTAL_WGT_ACTUAL"})

        With dst.Tables("SOTCONFT").Rows
            .Add(New Object() {1, "Reld", 0, 0})
            .Add(New Object() {2, IIf(MaintenanceMode, "Revd", "Conf"), 0, 0})
            .Add(New Object() {3, "Canc", 0, 0})
            .Add(New Object() {4, "Back", 0, 0})
        End With

        Sort_grdColumns(grdSOTCONFT, "KEY", True)

        Show_Filter(grdSOTSHIPX, True)
        grdSOTSHIPX.DisplayLayout.GroupByBox.Hidden = False

        ASCMAIN1.Add_Value_List(grdSOTSHIP1, "PICK_STATUS", Nothing, Nothing, 1)

        calFrom.Value = Now.Date.AddDays(-30)
        calTo.Value = Now.Date

        ASCMAIN1.Add_Value_List(grdSOTPICK1, "PICK_STATUS", Nothing, New String() {":", "P:In Pick", "F:Shipped", "D:Deleted", "C:Cancelled"})

        Position_txtSTORE()
        '  SplitContainer1.Panel2Collapsed = True

        If MaintenanceMode Then
            tabSelect.Tabs("3PL Shipments").Visible = False
            splHeader.Panel1Collapsed = True
            chkBO.Visible = False
        Else
            lblReason.Visible = False
            txtReason.Visible = False
            lblContact.Visible = False
            txtContact.Visible = False
            lblemail.Visible = False
            txtemail.Visible = False
        End If

        If InquiryMode Then
            tabSelect.Tabs("3PL Shipments").Visible = False
            MyBase.Absx1.dteFor("INV_DATE").MinDate = CDate("01/01/2013")
            MyBase.Absx1.dteFor("INV_DATE").MaxDate = validDates(1)
        Else
            MyBase.Absx1.dteFor("INV_DATE").MinDate = validDates(0)
            MyBase.Absx1.dteFor("INV_DATE").MaxDate = validDates(1)
        End If
    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFSHIPI")
        MaintenanceMode = (MENU_ITEM_OBJECT = "SOFSHIPM")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Select"

                RecreateLabel = False

                ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIP0)
                SHIP_BOL_NOs.Clear()

                Dim SHIP_STATUS As String = ""
                Dim rowSOTSHIP1 As DataRow = Nothing

                If Absx1.txtFor("SHIP_BOL_NO").Text = "" Then
                    EMsg &= vbCr & "You must First Select a Shipment No"
                Else
                    Dim SHIP_BOL_NO As String = Absx1.txtFor("SHIP_BOL_NO").Text
                    rowSOTSHIP1 = LookUp("SOTSHIP1", SHIP_BOL_NO)
                    If rowSOTSHIP1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Shipment No (" & SHIP_BOL_NO & ")"
                    Else
                        ' Debbie works on these in her screen. Do not allow modifications on this scree.
                        ' When she deletes the BOL from her screen the BILL_OF_LADING_NO value will be null
                        ' and hten it can be processed on this screewn
                        If rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty <> String.Empty Then
                            EMsg &= vbCr & "Shipment No (" & SHIP_BOL_NO & ") cannot be processed on this screen since it is attached to a Bill of lading."
                            Exit Select
                        End If

                        SHIP_STATUS = rowSOTSHIP1.Item("SHIP_STATUS")
                        ORDR_GROUP_NO = rowSOTSHIP1.Item("ORDR_GROUP_NO")
                        If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then Exit Sub
                        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                        CUST_CODE = rowSOTORDR0.Item("CUST_CODE")
                        ORDR_CUST_PO = rowSOTORDR0.Item("ORDR_CUST_PO") & ""

                        If optShipmentSelection.Value = "S" Then
                            If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then Exit Sub
                            SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                            ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO & "')")
                        Else
                            ASCMAIN1.sql = "Select SHIP_BOL_NO from SOTSHIP1 " _
                            & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                            & "   and SHIP_STATUS = '" & SHIP_STATUS & "'"
                            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                                Dim SHIP_BOL_NO2 As String = row.Item("SHIP_BOL_NO")
                                If Not InquiryMode Then If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO2) Then Exit Sub
                                Me.SHIP_BOL_NOs.Add(SHIP_BOL_NO2)
                                ASCDATA1.ExecuteSQL("Insert into " & SOTSHIP0 & " (SHIP_BOL_NO) values ('" & SHIP_BOL_NO2 & "')")
                            Next
                        End If
                    End If
                End If

                If EMsg.Length = 0 AndAlso rowSOTSHIP1 IsNot Nothing Then
                    Dim SHIP_VIA_CODE As String = rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty

                    ASCMAIN1.sql = "Select SOTSVIA1.* from SOTSVIA1, SOTCARR1" _
                        & " where SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE" _
                        & " and SOTCARR1.CARRIER_TYPE = 'U' and SOTSVIA1.SHIP_VIA_CODE = :PARM1"
                    If ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", SHIP_VIA_CODE) Is Nothing Then
                        EMsg &= vbCr & "Invalid Carrier Type for the Ship Via: " & SHIP_VIA_CODE
                    End If
                End If

                ' allow user to call up a previously billed shipment - need to look at some other variable

                If EMsg = "" Then
                    For Each SHIP_BOL_NO As String In SHIP_BOL_NOs
                        rowSOTSHIP1 = LookUp("SOTSHIP1", SHIP_BOL_NO)
                        If Not InquiryMode Then
                            If rowSOTSHIP1.Item("SHIP_STATUS") <> SHIP_STATUS Then
                                EMsg &= vbCr & "Shipment Status Changed for Shipment " & SHIP_BOL_NO
                            End If
                            If rowSOTSHIP1.Item("SHIP_BOL_NO_REV") & "" <> "" Then
                                EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is a Part of a Shipment/Invoice Reversal"
                            End If
                            If SHIP_STATUS = "P" Then
                                If rowSOTSHIP1.Item("SHIP_PICK_PRINTED") & "" = "" Then
                                    EMsg &= vbCr & "Pick Tickets have not been Printed (yet) for Shipment " & SHIP_BOL_NO
                                End If
                            Else
                                Select Case SHIP_STATUS
                                    Case "F"
                                        If ASCMAIN1.USER_SECURITY_CODEs.Contains("WL") Then
                                            If MessageBox.Show("The requested shiment has been finalized. Do you want to get a new shipping label?", _
                                                                "New label", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                                Exit Sub
                                            Else
                                                RecreateLabel = True
                                                Exit Select
                                            End If
                                        End If

                                        EMsg &= vbCr & "The requested shiment has been finalized."

                                        If 1 = 1 Then Exit Select

                                        Dim rowSOTCTLU1 As DataRow = LookUp("SOTCTLU1", "Z")
                                        If rowSOTCTLU1.Item("CTL_UPDATE_REQ") & "" = "C" Then
                                            MsgBox("There Has Been A Confirm that has not been updated by the Sales Journal." _
                                                   & "Please Run Sales Journal Before Proceeding", _
                                                   MsgBoxStyle.OkOnly, "Sales Journal Update Required First")
                                            Exit Sub
                                        End If

                                        ASCMAIN1.sql = "Select" _
                                            & "  Sum (DECODE(ORDR_YYYYPP_UPDATED,NULL,1,0)) PENDING" _
                                            & ", Sum (DECODE(ORDR_YYYYPP_UPDATED,NULL,0,1)) UPDATED" _
                                            & " FROM SOTINVH1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                                            & " and INV_NO_REV is Null"
                                        Dim row As DataRow = ASCDATA1.GetDataRow
                                        If Val(row.Item("UPDATED") & "") = 0 And (rowSOTSHIP1.Item("REGISTER_XNO") & "") = "" Then
                                            If MsgBox(CStr(row.Item("PENDING")) & " Pick Ticket(s) were Confirmed in this Shipment" _
                                                      & vbCrLf & vbCrLf _
                                                      & "Do you want to De-Confirm all Pick Tickets on this Shipment?" _
                                                      & vbCrLf & vbCrLf _
                                                      & "Warning - This feature performs the following: " _
                                                      & vbCrLf _
                                                      & "  1) Deletes Invoice Header & Details" & vbCrLf _
                                                      & "  2) Resets Pick Tickets to 'Unconfirmed'" & vbCrLf _
                                                      & "  3) Resets Shipment to 'Unconfirmed'" _
                                                      & vbCrLf & vbCrLf _
                                                      & "Note: You would not be offered this option if any of the Invoices associated with these Pick Tickets were Updated into the A/R", _
                                                      MsgBoxStyle.YesNo + MsgBoxStyle.Question, _
                                                      "Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped") = MsgBoxResult.Yes Then
                                                Stop ' SEE WJZ FOR TESTING
                                                De_Confirm(SHIP_BOL_NO)
                                            End If
                                            Exit Sub
                                        Else
                                            If (rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "N" _
                                            And rowSOTSHIP1.Item("SHIP_810_BATCH_NO") & "" <> "") _
                                            Or MaintenanceMode Then
                                                EMsg &= vbCr & "EDI Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped & Updated; No further Corrections are Permitted."
                                            Else
                                                ASCMAIN1.sql = "Select Count (*) from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                                                Dim PICKS As Int32 = Val(ASCDATA1.GetDataValue)

                                                If MsgBox("Shipment " & SHIP_BOL_NO & " has been Confirmed as Shipped & Updated" _
                                                          & vbCrLf & vbCrLf _
                                                          & "Do you want to Reverse Invoices Generated for all " & CStr(PICKS) & " Pick Tickets on this Shipment?" _
                                                          & vbCrLf & vbCrLf _
                                                          & "Warning - This feature performs the following: " & vbCrLf _
                                                          & "  1) Creates Negative Invoices" & vbCrLf _
                                                          & "  2) Resets Pick Tickets to 'Unconfirmed'" & vbCrLf _
                                                          & "  3) Resets Shipment to 'Unconfirmed'" & vbCrLf & vbCrLf _
                                                          & "Note: You would not be offered this option if this Shipment had already been Reversed", _
                                                          MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, _
                                                          "Shipment " & SHIP_BOL_NO & " has been Confirmed & Posted") = MsgBoxResult.Yes Then
                                                    If MsgBox("Are You Sure?", MsgBoxStyle.YesNo, _
                                                              "Verification to Reverse Invoices") = MsgBoxResult.Yes Then
                                                        Dim INV_REVERSAL_REASON As String = ""
                                                        Using F As New ASFMSGBF
                                                            INV_REVERSAL_REASON = F.Get_txt_from_User("Please Enter the Reason and then Click OK to Proceed", "Enter the Reason for Reversing")
                                                        End Using
                                                        Stop ' SEE WJZ FOR TESTING
                                                        Reverse_Invoice(SHIP_BOL_NO, INV_REVERSAL_REASON)
                                                    End If
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                    Case Else
                                        EMsg &= vbCr & "Shipment " & SHIP_BOL_NO & " is No Longer Open"
                                End Select
                            End If
                        End If
                    Next
                End If

                'If EMsg = "" And Not MaintenanceMode And Not InquiryMode Then
                '    Dim rowSOTSHIP1 As DataRow = LookUp("SOTSHIP1", SHIP_BOL_NOs(0))
                '    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowSOTSHIP1.Item("WHSE_CODE"))
                'End If

                If EMsg <> "" Then
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Update"

                If dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')='1' and PICK_QTY_CONF<>0").Length = 0 Then
                    EMsg &= vbCr & "Cannot Update when nothing is confirmed as shipped."
                    EMsg &= vbCr & "- Use Cancel Shipment option"
                End If

                If dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')='1' and PICK_QTY_CONF=0 and (PICK_CNT_CARTONS<>0 or PICK_TOTAL_WGT<>0)").Length <> 0 Then
                    EMsg &= vbCr & "Some Pick Tickets have 0 qty confirmed as Shipped"
                    EMsg &= vbCr & "-  but Still have a non-Zero value for cartons or weight"
                End If

                ' although this only matters for edi customers, I think we should enforce the integrity
                Dim rowSOTCARTX_oobal As DataRow() = dst.Tables("SOTCARTX").Select("ISNULL(PICK_QTY_CONF,0) <> ISNULL(QTY_PACKED,0)")
                If rowSOTCARTX_oobal.Length <> 0 Then
                    EMsg &= vbCr & "Pick Ticket Detail Qty Confirmed out of balance with Carton Details"
                    EMsg &= vbCr & " (See Pick Ticket " & rowSOTCARTX_oobal(0).Item("PICK_NO") & " Line " & rowSOTCARTX_oobal(0).Item("ORDR_LNO")
                End If

                If MaintenanceMode Then
                    Dim rowSOTPICK2_higher As DataRow() = dst.Tables("SOTPICK2").Select("ISNULL(PICK_QTY_CONF,0) < 0 or ISNULL(PICK_QTY_CONF,0) > ISNULL(PICK_QTY,0) + ISNULL(PICK_QTY_CANC_REL,0)")
                    If rowSOTPICK2_higher.Length <> 0 Then
                        EMsg &= vbCr & "Pick Ticket Detail Qty cannot be revised upward."
                        EMsg &= vbCr & " (See Pick Ticket " & rowSOTPICK2_higher(0).Item("PICK_NO") & " Line " & rowSOTPICK2_higher(0).Item("PICK_LNO")
                    End If

                    If Format(dteORDR_CANCEL_DATE.Value, "yyyyMMdd") < Format(dteORDR_SHIP_DATE.Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Cancel Date may not be prior to Ship Date"
                    End If
                    If txtReason.Text = "" Or txtContact.Text = "" Then
                        EMsg &= vbCr & "Reason and Contact are Mandatory when making changes to a Shipment"
                    End If

                Else
                    If Absx1.dteFor("SHIP_DATE_SHIPPED").Value & "" = "" _
                        Or Absx1.dteFor("INV_DATE").Value & "" = "" Then
                        EMsg &= vbCr & "Date Shipped and Invoice Date are Mandatory"
                    Else
                        If Format(Absx1.dteFor("SHIP_DATE_SHIPPED").Value, "yyyyMMdd") _
                         > Format(Absx1.dteFor("INV_DATE").Value, "yyyyMMdd") Then
                            EMsg &= vbCr & "Invoice Date cannot be Prior to Date Shipped"
                        End If

                        If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") <> ASCMAIN1.CYM Then
                            If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, 1) Then
                                If EMsg.Length = 0 Then
                                    If MessageBox.Show("You are about to confirm a shippment that will be posted into the Next period. Do you want to continue?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                        Exit Sub
                                    End If
                                End If
                            Else
                                EMsg &= vbCr & "Invoice Date Not in Current Period"
                            End If
                        End If
                    End If

                    ' Invoice cannot be prior to today!!
                    If EMsg.Length = 0 Then
                        If ASCMAIN1.SOLUTION <> "AHA" Then
                            If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMMdd") < Format(DateTime.Now, "yyyyMMdd") Then
                                EMsg &= vbCr & "Invoice Date cannot be Prior to today."
                            End If

                            If Format(Absx1.dteFor("SHIP_DATE_SHIPPED").Value, "yyyyMMdd") < Format(DateTime.Now, "yyyyMMdd") Then
                                EMsg &= vbCr & "Ship Date cannot be Prior to today."
                            End If
                        End If

                        ' Ship Date must be in current or next period
                        If Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") < Format(validDates(0), "yyyyMM") _
                            OrElse Format(Absx1.dteFor("INV_DATE").Value, "yyyyMM") > Format(validDates(1), "yyyyMM") Then
                            EMsg &= vbCr & "Invoice Date must be between " & validDates(0) & " and " & validDates(1)
                        End If
                    End If

                    ' Is this only for Fedex??
                    'Get_PARM("SOTPARM1")
                    'If Absx1.dteFor("SHIP_DATE_SHIPPED").Value < ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE") Then
                    '    EMsg &= vbCr & "Ship date may not be before " & ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE")
                    'End If

                    If Absx1.txtFor("TERM_CODE").Text = "" Then
                        EMsg &= vbCr & "Terms Code is Required"
                    Else
                        If LookUp("TATTERM1", Absx1.txtFor("TERM_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Terms Code"
                        End If
                    End If
                    If Absx1.txtFor("SHIP_VIA_CODE").Text = "" Then
                        EMsg &= vbCr & "Ship Via Code is Required"
                    Else
                        If LookUp("SOTSVIA1", Absx1.txtFor("SHIP_VIA_CODE").Text) Is Nothing Then
                            EMsg &= vbCr & "Invalid Ship Via Code"
                        End If
                    End If
                    If Absx1.txtFor("FRT_TERMS").Text = "" Then
                        EMsg &= vbCr & "Frt Terms Code is Required"
                    Else
                        If LookUp("ASTCODE1", New String() {"ARTCUST1", "FRT_TERMS", Absx1.txtFor("FRT_TERMS").Text}) Is Nothing Then
                            EMsg &= vbCr & "Invalid Frt Terms Code"
                        End If
                    End If

                    If Absx1.txtFor("SREP_CODE").Text <> "" AndAlso LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Sales Rep Code"
                    End If
                    If Absx1.txtFor("SREP2_CODE").Text <> "" AndAlso LookUp("SOTSREP1", Absx1.txtFor("SREP2_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Sales Rep2 Code"
                    End If

                    If dst.Tables("SOTPICK2").Select("PICK_QTY_CANC <> 0").Length <> 0 Then
                        If Absx1.txtFor("REASON_CODE").Text = "" Then
                            EMsg &= vbCr & "Reason Code is Required when Cancelling Qty's on any Pick Ticket"
                        Else
                            If LookUp("SOTREAS1", Absx1.txtFor("REASON_CODE").Text) Is Nothing Then
                                EMsg &= vbCr & "Invalid Reason Code"
                            End If
                        End If
                    Else
                        If Absx1.txtFor("REASON_CODE").Text <> "" Then
                            EMsg &= vbCr & "Reason Code should NOT be specified unless Cancelling Qty's"
                        End If
                    End If

                    If dst.Tables("SOTPICK2").Select("PICK_QTY_CONF <> 0").Length <> 0 Then
                        If Absx1.txtFor("SHIP_VIA_CODE").Text = "" Then
                            'Or Absx1.txtFor("SHIP_REF").Text = "" Then
                            'EMsg &= vbCr & "Ship Via Code and Shippers Reference (Pro #) are Required"
                            EMsg &= vbCr & "Ship Via Code is Required"
                        Else
                            Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", Absx1.txtFor("SHIP_VIA_CODE").Text)
                            If rowSOTSVIA1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Ship Via Code"
                            Else
                                If edi_customer Then
                                    If rowSOTSVIA1.Item("SHIP_VIA_SCAC") & "" = "" Then
                                        EMsg &= vbCr & "Selected Shipper Requires SCAC Code For EDI Customers"
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' ***************************************************************************************************************
                    ' This is not the case for Ahava. We will use the SHIP_BOL_NO for the Bill_Of_Lading_no when transmiting the ASN
                    ' ***************************************************************************************************************
                    'If dst.Tables("SOTPICK2").Select("PICK_QTY_CONF <> 0").Length <> 0 And edi_order Then
                    '    If dst.Tables("SOTSHIP1").Rows.Count > 1 Then
                    '        If dst.Tables("SOTPICK1").Select("SELECTED = '1' and BILL_OF_LADING_NO is Null").Length <> 0 Then
                    '            EMsg &= vbCr & "BOL No is Mandatory for EDI Orders"
                    '        End If
                    '    Else
                    '        If Absx1.txtFor("BILL_OF_LADING_NO").Text = "" Then
                    '            EMsg &= vbCr & "BOL No is Mandatory for EDI Orders"
                    '        End If
                    '    End If
                    'End If

                    Dim sqlw As String = ""

                    If Absx1.txtFor("FRT_TERMS").Text <> "" Then
                        If Absx1.txtFor("FRT_TERMS").Text <> "PPA" Then
                            sqlw = "PICK_FREIGHT <> 0 and SELECTED = '1'"
                        Else
                            sqlw = "PICK_FREIGHT = 0 and SELECTED = '1'"
                        End If
                        If dst.Tables("SOTPICK1").Select(sqlw).Length > 0 Then
                            If Absx1.txtFor("FRT_TERMS").Text <> "PPA" Then
                                EMsg &= vbCr & "Freight Terms Code Specified does not permit Non-Zero Freight Amounts"
                            Else
                                EMsg &= vbCr & "Freight Terms Code Specified does not permit Zero Freight Amounts"
                            End If
                        End If
                    End If

                    If edi856_customer Then
                        sqlw = "PICK_TOTAL_UNITS_CALC <> PICK_QTY_CONF"
                        Dim rows() As DataRow = dst.Tables("SOTPICK1").Select(sqlw)
                        If rows.Length <> 0 Then
                            EMsg = EMsg & vbCr & CStr(rows.Length) & " Pick Ticket(s) not matching Carton Details (See PT#" & rows(0).Item(0) & ")"
                        End If
                    End If

                    'Check for Assortments Knocked Out of Balance.

                    If EMsg = "" Then
                        If (chkFactored.Checked And rowARTCUST1.Item("CUST_FACTOR_IND") & "" <> "1") _
                        Or (Not chkFactored.Checked And rowARTCUST1.Item("CUST_FACTOR_IND") & "" = "1") Then
                            If MsgBox("Factor Option is not in synch with Customer Master" & vbCrLf & "Continue Anyway", _
                                       MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If

                    'If dst.Tables("SOTPICK2").Select("ORDR_UNIT_PRICE = 0").Length <> 0 Then
                    '    If MsgBox("This Shipment Contains Items That have Zero Prices." _
                    '         & vbCrLf & "Are You Sure You Want To Update This?", MsgBoxStyle.YesNo, _
                    '         "Price Check") = MsgBoxResult.No Then
                    '        EMsg &= vbCr & "Cancelled By User Due To Zero Price."
                    '    End If
                    'End If
                End If

                ' If common carrier then We need to know the packaging attributes
                If Not commonCarrier Then
                    EMsg &= vbCr & "This screen is for Common Carriers only"
                Else
                    Dim CART_SEQ As Int16 = Val(dst.Tables("SOTCART1").Compute("MAX(CART_SEQ)", "") & String.Empty) + 1
                    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1'")
                        Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")

                        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'")
                            ' Make sure all cartons have a sequence number. This is how the shipping labels will print
                            If rowSOTCART1.Item("CART_SEQ") & String.Empty = String.Empty Then
                                rowSOTCART1.Item("CART_SEQ") = CART_SEQ
                                CART_SEQ += 1
                            End If
                            If rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty = String.Empty Then
                                EMsg &= vbCrLf & "Package type is required for all cartons"
                            ElseIf Val(rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty) = DPayments.DShippingSDK.TPackagingTypes.ptYourPackaging Then
                                If rowSOTCART1.Item("PKG_CODE") & String.Empty = String.Empty Then
                                    EMsg &= vbCrLf & "Package code is required for all 'Our Packaging' cartons"
                                ElseIf rowSOTCART1.Item("PKG_CODE") & String.Empty = "OTHER" Then
                                    If Val(rowSOTCART1.Item("LENGTH") & String.Empty) <= 0 _
                                             OrElse Val(rowSOTCART1.Item("WIDTH") & String.Empty) <= 0 _
                                            OrElse Val(rowSOTCART1.Item("HEIGHT") & String.Empty) <= 0 Then
                                        EMsg &= vbCr & "package Type Other requires Width, length and height to be set."
                                    End If
                                End If
                            End If

                            If Val(rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty) <= 0 Then
                                EMsg &= vbCrLf & "Package weight is required for all cartons"
                            Else
                                Dim pickWeight As Decimal = Val(rowSOTPICK1.Item("PICK_TOTAL_WGT") & String.Empty)
                                Dim totalCartWeight As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)

                                If totalCartWeight > pickWeight Then
                                    EMsg &= vbCrLf & "Package weight for a Pick Ticket must be greater equal total cartons weight (" & PICK_NO & ")"
                                End If
                            End If
                        Next
                    Next
                End If

                ' Prescreen generating Shipping label
                If EMsg.Length = 0 Then
                    Dim list As New List(Of String)
                    Dim ErrorMessage As String = String.Empty

                    If Not RequestShippingLabel(list, ErrorMessage, True) Then
                        If ErrorMessage.Length > 0 Then
                            EMsg &= vbCrLf & ErrorMessage
                        End If
                    End If
                End If

                If EMsg.Length = 0 Then
                    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')= '1' and ORDR_STAX > 0")
                        If (rowSOTPICK1.Item("STAX_CODE") & String.Empty).ToString = String.Empty Then
                            EMsg &= vbCr & "All Entries with a Sales Tax Amount require a Sales Tax Code"
                            Exit For
                        End If
                    Next
                End If

                If EMsg.Length = 0 Then
                    If Val(numInsureValue.Value & String.Empty) > 200 AndAlso Not chkInsureShipment.Checked Then
                        If MessageBox.Show("Do you want to continue without insuring the shipment.", "Insure", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                If EMsg.Length = 0 Then
                    Dim packageInsure As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(INSURANCE)", "") & String.Empty)
                    If chkInsureShipment.Checked Then
                        If packageInsure <= 0 Then
                            EMsg &= vbCr & "You chose to insure the shipment put you did not place Insurance values on the cartons."
                        Else
                            If MessageBox.Show("Do you want to insure the shipment for a total of $" & Format(packageInsure, "#,##0.00"), "Insure", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                                Exit Sub
                            End If
                        End If
                    ElseIf packageInsure > 0 Then
                        If MessageBox.Show("You provided insurance for the carton(s); however, you did not check 'Insure Package'. Do you want to continue without applying the insurance?", _
                             "Insure", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                End If


            Case "Cancel"
                If MsgBox("Are you sure that you want to Cancel?", _
                      MsgBoxStyle.YesNo, _
                      "Verification to Cancel working with this Record") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Cancel Shipment"

                ' You may use this option only if all of the Pick Tickets are already Cancelled

                If dst.Tables("SOTPICK1").Select("ISNULL(SELECTED,'0')<>'1'").Length <> 0 Then
                    EMsg &= vbCr & "All Pick Tickets must be Selected in order to Cancel the Entire Shipment"
                    EMsg &= vbCr & "You may not cancel some pick tickets (and leave others open) with this option"
                End If

                If dst.Tables("SOTPICK1").Select("PICK_STATUS <> 'C' AND PICK_QTY_CONF <> 0").Length <> 0 Then
                    EMsg &= vbCr & "Cancellation Not Permitted" & vbCrLf & " - Some Pick Tickets on this Shipment are NOT Cancelled"
                    EMsg &= vbCr & vbCr & "Click on Shipment (in Mass Changes) and then use the Cancel button"
                End If

                If EMsg = "" Then
                    If MsgBox("This option will Cancel this Shipment." _
                              & vbCrLf & vbCrLf & "Use this option to Cancel All Pick Tickets on this Shipment" _
                              & vbCrLf & " and also Cancel this Shipment." _
                              & vbCrLf _
                              & vbCrLf & "This option will NOT restore the Order back to an Open state." _
                              & vbCrLf & "This option will NOT cause any EDI documents to transmit." _
                              & vbCrLf & "This option will NOT create Invoices." _
                              & vbCrLf & vbCrLf & "If you want to cancel this shipment so that the orders are re-opened," _
                              & vbCrLf & " then use De-Release." _
                              & vbCrLf & vbCrLf & "Are you sure that you want to Cancel this Shipment?", _
                                  MsgBoxStyle.YesNo, _
                                  "WARNING: This Action is Permanent") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Add Carton"
                If MessageBox.Show("Do you want to add a carton to this shipment?", "Add Carton", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
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
            Case "Select"
                If InquiryMode Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                refreshScreen = True
                If MaintenanceMode Then
                    Update_Record_Maintenance()
                Else
                    Update_Record()
                End If
                If refreshScreen Then
                    Mode_Settings(False)
                End If

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done", "Cancel"
                Mode_Settings(False)

            Case "Cancel Shipment"
                Cancel_Shipment()
                Mode_Settings(False)

            Case "Force PTs to Balance"
                Force_PTs_to_Balance()

            Case "Force Cartons to Balance"
                Force_Cartons_to_Balance()

            Case "Add Carton"
                AddCarton()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Select").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Update").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("Cancel Shipment").Settings.Enabled = iScreenMode
                .Items("Done").Visible = InquiryMode Or (EntryMode = "L" And ScreenMode)
                .Items("Select").Visible = (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Add Carton").Settings.Enabled = iScreenMode

                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
                .Items("Cancel Shipment").Visible = (Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode) And Not MaintenanceMode) AndAlso ASCMAIN1.USER_SECURITY_CODEs.Contains("WL")
                .Items("Add Carton").Visible = Not InquiryMode And (Not (EntryMode = "L") Or Not ScreenMode)
            End With
            .Groups("Totals").Visible = ScreenMode
            .Groups("Special Operations").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode ' And Not MaintenanceMode

            If ASCMAIN1.DBS_COMPANY <> "VAN" Then
                .Groups("Special Operations").Visible = False
            End If
            .Groups("Special Operations").Items("Substitute").Visible = Not MaintenanceMode
            .Groups("Special Operations").Items("Add Line").Visible = Not MaintenanceMode
            '.Groups("Order Header Changes").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode And Not select_from_3PL_list And MaintenanceMode
            .Groups("Mass Changes").Visible = ScreenMode And (EntryMode <> "L") And Not InquiryMode

            .Groups("Shipment Status").Visible = Not ScreenMode And InquiryMode
            .Groups("Shipment Selection").Visible = Not ScreenMode

            .Groups("Special Operations").Items("Force PTs to Balance").Visible = ScreenMode And edi_order And edi856_customer
            .Groups("Special Operations").Items("Force Cartons to Balance").Visible = ScreenMode And ((edi_order And edi856_customer) Or MaintenanceMode)
        End With

        '  lblStatus.Visible = ScreenMode

        'grdSOTSHIPX.Visible = Not tf
        tabSelect.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), edi_order Or Not (EntryMode = "E" Or EntryMode = "N"))

        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CODE"), True)
        Set_Read_Only_for_ctl(Absx1.optFor("SHIP_ADDR_TYPE"), True)
        Set_Read_Only_for_ctl(Absx1.txtFor("SHIP_ADDR_CODE"), True)

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                {grdSOTPICK1, grdSOTPICK2, grdSOTCART1, grdSOTCART2}
                With grd.DisplayLayout.Override
                    .AllowAddNew = UltraWinGrid.AllowAddNew.No
                    .AllowDelete = DefaultableBoolean.False
                    If (EntryMode = "N" Or EntryMode = "E") Then
                        .AllowUpdate = DefaultableBoolean.True
                    Else
                        .AllowUpdate = DefaultableBoolean.False
                    End If
                End With
            Next
            Setup_SOTPICK1() ' because allowupdate is toggled based on status of active pick1 record
            If MaintenanceMode Then
                grdSOTCART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            End If

            Set_Read_Only(splHeader, Not (EntryMode = "N" Or EntryMode = "E"))
            If Not InquiryMode And Not MaintenanceMode Then
                Set_Read_Only(splHeader.Panel1, False)
            End If

            Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(grpSHIPTO, (edi_order AndAlso Not ASCMAIN1.USER_SECURITY_CODEs.Contains("WL")) Or Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(grpHeaderInfo, MaintenanceMode Or Not (EntryMode = "E" Or EntryMode = "N"))
            Set_Read_Only(grpShippingWindow, Not MaintenanceMode Or Not (EntryMode = "E" Or EntryMode = "N"))

            With grdSOTPICK2.DisplayLayout.Bands(0).Columns("ITEM_CODE")
                If (EntryMode = "N" Or EntryMode = "E") And (1 <> 1) Then ' HOW IN THE WORLD IS THIS TO BE PERMITTED?
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGray 'LightGreen
                End If
            End With

            With grdSOTPICK1.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT"}
                    .Columns(COLUMN_NAME).Hidden = edi_order And edi856_customer
                    If COLUMN_NAME <> "PICK_FREIGHT" Then .Columns(COLUMN_NAME & "_CALC").Hidden = Not (edi_order And edi856_customer)
                    ' NOTE THAT FRT IS NOT SHOWN IF edi_order And edi856_customer; ASSUMPTION IS THAT THERE WILL BE NO FRT IF EDI
                Next

                With .Columns("BILL_OF_LADING_NO")
                    If Not edi_order Or Absx1.optFor("SHIP_ADDR_TYPE").Value = "MK" And Not MaintenanceMode Then
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                End With
            End With

            chkBO.Checked = Not MaintenanceMode And Not (edi_customer) And (rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & "" = "1")
            Setup_BO()
        Else
            Clear_Record()
        End If

        Absx1.txtFor("SHIP_ADDR_CODE").Visible = Not ScreenMode Or (Absx1.optFor("SHIP_ADDR_TYPE").Value = "DC")
        'tabSOTPICK1.Tabs("Cartons").Visible = (edi856_customer And edi_order)  '  And Not MaintenanceMode

        If tabSOTPICK1.Tabs("Cartons").Visible Then
            grdSOTCART2.Parent = splSOTCART1.Panel2
            grdSOTCART2.DisplayLayout.Bands(0).Columns("CART_NO").Hidden = True
            'splSOTPICK1.Panel2Collapsed = True
        Else
            grdSOTCART2.Parent = splSOTPICK2.Panel2
            grdSOTCART2.DisplayLayout.Bands(0).Columns("CART_NO").Hidden = False
            'splSOTPICK1.Panel2Collapsed = False
        End If

        Position_txtSTORE()
        ' lblBILL_OF_LADING_NO.Visible = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)
        ' Absx1.txtFor("BILL_OF_LADING_NO").Visible = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)
        If Not InquiryMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("BILL_OF_LADING_NO"), (dst.Tables("SOTSHIP1").Rows.Count > 1))
        End If
        grdSOTPICK1.DisplayLayout.Bands(0).Columns("BILL_OF_LADING_NO").Hidden = Not (dst.Tables("SOTSHIP1").Rows.Count > 1)

        If ASCMAIN1.DBS_COMPANY = "RGI" Or ASCMAIN1.DBS_SERVER = "RGI" Then
            chkFactored.Visible = False
        End If
    End Sub

    Sub Clear_Record()

        Absx1.txtFor("PICK_NO").Clear()

        txtReason.Clear()
        txtContact.Clear()
        txtemail.Clear()

        txt3PAccountNo.Clear()
        txt3pCountry.Clear()
        txt3PZipCode.Clear()

        CUST_CODE = String.Empty
        ORDR_GROUP_NO = String.Empty
        ORDR_CUST_PO = String.Empty
        RecreateLabel = False
        requestShippingOptions = False

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTINVH1", "SOTINVH2", "SOTINTL1", _
             "SOTCART1", "SOTCART2", "SOTORDR1", "SOTORDR2", "SOTSHIP0", "SOTCARTX", "SOTCARRX", _
             "SOTSHIP3", "SOTSHIP4", "SOTSHIP6", "SOTORDR5", _
             "ARTCCPA1", "ARTCCPA2", "ARTCCPDA", "ARTOPEN1", _
             "WHTSHIP1", "WHTSHIP2", "WHTSHIP3", "WHTSHIP5", "WHTSHIPC", "WHTSHIPS", "WHTSHIPP"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_SOTSHIPX()

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Permit Price Change"), UltraWinToolbars.StateButtonTool)
        tlb_sbt.Checked = False

        optAddressType.Value = "O"
        chkInsureShipment.Checked = False
        chkInsureShipment_CheckedChanged(Nothing, Nothing)
        chkSaturday.Checked = False
        optPayor.Value = "O"
        chkSignature.Checked = False

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        ToggleDataTableExpressions(False)

        Get_PARM("SOTPARM1")

        If EntryMode = "N" Then
        Else
            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)

            Fill_Records("EDTTRPMC", CUST_CODE)
            edi_customer = (dst.Tables("EDTTRPMC").Rows.Count <> 0)
            Dim rowEDTTRPMC As DataRow = dst.Tables("EDTTRPMC").Rows.Find(New Object() {CUST_CODE, "856"})
            edi856_customer = rowEDTTRPMC IsNot Nothing AndAlso rowEDTTRPMC.Item("EDI_STATUS") & "" = "P"
            lblASN.Visible = (edi856_customer And edi_order)

            If MaintenanceMode Then
                Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                ORDR_SHIP_DATE = rowSOTORDR0.Item("ORDR_SHIP_DATE")
                ORDR_CANCEL_DATE = rowSOTORDR0.Item("ORDR_CANCEL_DATE")
                dteORDR_SHIP_DATE.Value = ORDR_SHIP_DATE
                dteORDR_CANCEL_DATE.Value = ORDR_CANCEL_DATE
                txtReason.Text = ""
            End If

            'ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"

            If Not RecreateLabel Then
                ASCMAIN1.sql = "Select * from SOTORDR1 Where ORDR_NO IN"
                ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & HFs("SHIP_BOL_NO") & "' AND PICK_STATUS = 'P')"
            Else
                ASCMAIN1.sql = "Select * from SOTORDR1 Where ORDR_NO IN"
                ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & HFs("SHIP_BOL_NO") & "' AND PICK_STATUS = 'F')"
            End If
            Fill_Records("SOTORDR1", String.Empty, True, ASCMAIN1.sql)

            'ASCMAIN1.sql = "Select * from SOTORDR2 WHERE ORDR_NO IN (Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')"
            If Not RecreateLabel Then
                ASCMAIN1.sql = "Select * from SOTORDR2 Where ORDR_NO IN"
                ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & HFs("SHIP_BOL_NO") & "' AND PICK_STATUS = 'P')"
            Else
                ASCMAIN1.sql = "Select * from SOTORDR2 Where ORDR_NO IN"
                ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & HFs("SHIP_BOL_NO") & "' AND PICK_STATUS = 'F')"
            End If
            Fill_Records("SOTORDR2", String.Empty, True, ASCMAIN1.sql)

            'ASCMAIN1.sql = "Select * from SOTORDR5 where CUST_ADDR_TYPE = 'ST' AND ORDR_NO IN (Select ORDR_NO from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "')"
            If Not RecreateLabel Then
                ASCMAIN1.sql = "Select * from SOTORDR5 Where ORDR_NO IN"
                ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & HFs("SHIP_BOL_NO") & "' AND PICK_STATUS = 'P')"
            Else
                ASCMAIN1.sql = "Select * from SOTORDR5 Where ORDR_NO IN"
                ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & HFs("SHIP_BOL_NO") & "' AND PICK_STATUS = 'F')"
            End If
            Fill_Records("SOTORDR5", "", True, ASCMAIN1.sql)

            'ASCMAIN1.sql = "Select Count (*) from SOTORDR1" _
            '    & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_SOURCE = 'E'"
            'edi_order = (Val(ASCDATA1.GetDataValue) <> 0)
            edi_order = dst.Tables("SOTORDR1").Select("ORDR_SOURCE = 'E'").Length > 0
            If edi_order Then
                lblSource.Text = "EDI"
            Else
                lblSource.Text = "Manual"
            End If

            'ASCMAIN1.sql = "Select Distinct ORDR_SOURCE from SOTORDR1" _
            '    & " where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            'Dim rowORDR_SOURCE As DataRow = ASCDATA1.GetDataRow
            'ORDR_SOURCE = rowORDR_SOURCE.Item("ORDR_SOURCE")

            ' Web orders are assumed to be residential addresses as default
            If dst.Tables("SOTORDR1").Select("ORDR_SOURCE = 'W'").Length > 0 Then
                optAddressType.Value = "R"
            End If

            Dim sqlwhere_SOTSHIP1 As String = "" _
                & "   and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & SOTSHIP0 & ")" & vbCrLf _
                & IIf(InquiryMode OrElse RecreateLabel, _
                      "", _
                      "" _
                        & "   and SOTSHIP1.SHIP_STATUS = 'P' and SOTSHIP1.SHIP_PICK_PRINTED is Not Null")

            If RecreateLabel Then
                ASCMAIN1.sql = "Select SOTSHIP1.*" _
                    & ",SOTORDR0.CUST_CODE,SOTORDR0.ORDR_SHIP_DATE,SOTORDR0.ORDR_CANCEL_DATE,SOTORDR0.ORDR_CUST_PO" _
                    & " from SOTSHIP1,SOTORDR0,SOTSVIA1,SOTCARR1" _
                    & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                    & " and SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE" & vbCrLf _
                    & " and SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE" & vbCrLf _
                    & " and SOTCARR1.CARRIER_TYPE = 'U'" & vbCrLf _
                    & " and SOTSHIP1.SHIP_BOL_NO = '" & HFs("SHIP_BOL_NO") & "'"
            Else
                ASCMAIN1.sql = sqlSOTSHIPX
            End If

            ASCMAIN1.sql &= sqlwhere_SOTSHIP1

            Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)
            If dst.Tables("SOTSHIP1").Rows.Count <> SHIP_BOL_NOs.Count Then Stop ' NEED AN ABORT LOAD FEATURE IN STDS

            ' Evaluate the Order FOB, IF ENPTY THEN FILL FROM SOTORDR1
            If dst.Tables("SOTSHIP1").Rows(0).Item("SHIP_FOB") & String.Empty = String.Empty Then
                txtSHIP_FOB.Text = dst.Tables("SOTORDR1").Compute("MAX(ORDR_FOB)", "") & String.Empty
                dst.Tables("SOTSHIP1").Rows(0).Item("SHIP_FOB") = txtSHIP_FOB.Text
            End If

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                rowSOTORDR1.Item("ORDR_FOB") = dst.Tables("SOTSHIP1").Rows(0).Item("SHIP_FOB") & String.Empty
            Next

            ASCMAIN1.sql = sqlSOTPICK1 & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                rowSOTPICK1.Item("SELECTED") = "1"
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                If rowSOTORDR1 IsNot Nothing Then
                    rowSOTPICK1.Item("ORDR_STAX") = Val(rowSOTORDR1.Item("ORDR_STAX") & String.Empty)
                    rowSOTPICK1.Item("STAX_CODE") = rowSOTORDR1.Item("STAX_CODE") & String.Empty
                End If
            Next

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                For Each COLUMN_NAME As String In New String() _
                    {"FRT_TERMS", "SHIP_VIA_CODE"}
                    If rowSOTSHIP1.Item(COLUMN_NAME) & "" = "" Then rowSOTSHIP1.Item(COLUMN_NAME) = rowARTCUST1.Item(COLUMN_NAME)
                Next
                If rowARTCUST1.Item("TERM_CODE") & "" <> "" Then
                    rowSOTSHIP1.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                    For Each rowSOTPICK1 As DataRow In rowSOTSHIP1.GetChildRows("SOTSHIP1_SOTPICK1")
                        rowSOTPICK1.Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                    Next
                End If
            Next

            Dim rowSOTPICK1_0 As DataRow = dst.Tables("SOTPICK1").Rows(0)
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                For Each COLUMN_NAME As String In New String() _
                    {"ORDR_DEPT", "SREP_CODE", "SREP2_CODE", "TERM_CODE"}
                    If rowSOTSHIP1.Item(COLUMN_NAME) & "" = "" Then rowSOTSHIP1.Item(COLUMN_NAME) = rowSOTPICK1_0.Item(COLUMN_NAME)
                Next
            Next

            Dim row As DataRow = dst.Tables("SOTSHIP1").Rows(0)
            rowSOTSHIP0 = dst.Tables("SOTSHIP0").NewRow
            For i As Integer = 0 To dst.Tables("SOTSHIP0").Columns.Count - 1
                rowSOTSHIP0.Item(i) = row.Item(i)
            Next
            dst.Tables("SOTSHIP0").Rows.Add(rowSOTSHIP0)

            rowSOTSHIP0_ORIG = dst.Tables("SOTSHIP0").NewRow
            rowSOTSHIP0_ORIG.ItemArray = rowSOTSHIP0.ItemArray

            chkFactored.Checked = (dst.Tables("SOTPICK1").Select("CUST_FACTOR_IND = '1'").Length > 0)

            ASCMAIN1.sql = sqlSOTPICK2 & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)

            Dim CART_SEQ As Int16 = 1
            'Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", "X1")
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
                ' default to packaging code 31 - Our package
                rowSOTCART1.Item("PACKAGING_TYPE") = 31
                rowSOTCART1.Item("CART_SEQ") = CART_SEQ

                ' 12/18/2015 - Non 3PL shipments defualt to boxes
                If ASCMAIN1.DBS_COMPANY = "AHA" OrElse ASCMAIN1.DBS_SERVER = "AHA" Then
                    rowSOTCART1.Item("PKG_CODE") = "X1"
                    rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = Val(rowSOTCART1.Item("CART_TOTAL_WGT_CALC") & String.Empty)
                End If

                CART_SEQ += 1
                If edi_order And edi856_customer Then
                    rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = rowSOTCART1.Item("CART_TOTAL_WGT_CALC")
                End If
            Next

            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
                & " from SOTCART2,SOTCART1,SOTPICK1,SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)

            'For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("")
            '    rowSOTCART2.Item("QTY_PACKED") = rowSOTCART2.Item("QTY_PACKED")
            'Next

            Dim INV_SALES As Decimal = 0
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
                rowSOTPICK2.Item("PICK_UNIT_PRICE") = rowSOTPICK2.Item("ORDR_UNIT_PRICE")
                'rowSOTPICK2.Item("PICK_QTY_CONF") = rowSOTPICK2.Item("PICK_QTY")
                'rowSOTPICK2.Item("PICK_QTY_CANC") = 0
                'rowSOTPICK2.Item("PICK_QTY_BACK") = 0

                Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
                Dim PICK_UNIT_PRICE As Decimal = Val(rowSOTPICK2.Item("PICK_UNIT_PRICE") & "")
                INV_SALES = INV_SALES + ORDR_QTY_SHIP * PICK_UNIT_PRICE
            Next
            numInsureValue.Value = INV_SALES

            Fill_Records("SOTORDRX", New Object() {HFs("SHIP_BOL_NO")})
            dteORDR_SHIP_DATE.Value = dst.Tables("SOTORDRX").Compute("MIN(ORDR_SHIP_DATE)", "")
            dteORDR_CANCEL_DATE.Value = dst.Tables("SOTORDRX").Compute("MIN(ORDR_CANCEL_DATE)", "")
        End If

        ToggleDataTableExpressions(True)

        dst.Tables("SOTPICK1").AcceptChanges()
        dst.Tables("SOTPICK2").AcceptChanges()

        Sort_grdColumns(grdSOTPICK1, "PICK_NO")
        Setup_SOTPICK1()

        clsPrice_Change = Nothing

        Select Case rowSOTSHIP0.Item("SHIP_STATUS")
            Case "P"
                lblStatus.Text = "In Pick"
                ' Set to ship date in SOTPARM1
                Get_PARM("SOTPARM1")
                If Not IsDate(Absx1.dteFor("SHIP_DATE_SHIPPED").Value & String.Empty) Then
                    Absx1.dteFor("SHIP_DATE_SHIPPED").Value = DateTime.Now ' ROWs("SOTPARM1").Item("SO_PARM_WH_SHIP_DATE")
                End If
                Absx1.dteFor("INV_DATE").Value = Absx1.dteFor("SHIP_DATE_SHIPPED").Value
            Case "F"
                lblStatus.Text = "Shipped"
            Case "C"
                lblStatus.Text = "Cancelled"
            Case Else
                lblStatus.Text = "Status Unknown"
        End Select

        If EntryMode = "L" Then
            lblINIT_DATE.Text = "Confirmed by " & rowSOTSHIP0.Item("LAST_OPER") & " on " & Format(rowSOTSHIP0.Item("LAST_DATE"), "MM/dd/yy HH:mm")
        Else
            lblINIT_DATE.Text = "Confirmed by " & ASCMAIN1.USER_ID & " on " & Format(Now, "MM/dd/yy HH:mm")
        End If

        Display_Totals()

        Dim GL_PARM_CURR_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & ""
        CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
        If CURR_CODE = "" Or CURR_CODE = GL_PARM_CURR_CODE Then
            CURR_CODE = GL_PARM_CURR_CODE
            CURR_EXCH_RATE = 1
        Else
            Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", rowARTCUST1.Item("CURR_CODE"))
            CURR_CODE = rowTATCURR1.Item("CURR_CODE")
            CURR_EXCH_RATE = rowTATCURR1.Item("CURR_EXCH_CUR")
        End If

        dst.Tables("SOTCARTX").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTPICK2"), New String() {"PICK_NO", "ORDR_NO", "ORDR_LNO"}).Rows
            dst.Tables("SOTCARTX").Rows.Add(New Object() {row.Item("PICK_NO"), row.Item("ORDR_NO"), row.Item("ORDR_LNO")})
        Next

        'grdSOTSHIP1.DisplayLayout.Bands(0).Summaries.Clear()

        grdSOTPICK1.DisplayLayout.Bands(0).Summaries.Clear()
        If dst.Tables("SOTPICK1").Rows.Count = 1 Then
            splSOTPICK1.SplitterDistance = 80 + grdSOTPICK1.Rows(0).Height * 1
            txtStore.Visible = False
        Else
            'CANT WE JUST CREATE THE SUMMARIES ONCE AND THEN HIDE THEM?
            Create_Summary(grdSOTPICK1, "PICK_NO", "Count")
            Create_Summary(grdSOTPICK1, New String() {"SELECTED", "PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT"})
            Create_Summary(grdSOTPICK1, New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"})
            Create_Summary(grdSOTPICK1, New String() {"PICK_AMT", "PICK_AMT_CONF", "PICK_AMT_CANC", "PICK_AMT_BACK"})
            splSOTPICK1.SplitterDistance = 80 + grdSOTPICK1.Rows(0).Height * 4
            txtStore.Visible = True
        End If

        If MaintenanceMode Then
            txtemail.Text = ASCMAIN1.USER_EMAIL
            txtContact.Text = ASCMAIN1.USER_NAME
        End If

        Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")
        Dim inTransaction As Boolean = False
        Dim shipLabels As List(Of String) = New List(Of String)
        Dim ErrorMessage As String = String.Empty

        Try
            refreshScreen = True

            ' Need to see if we needs to get labels from carrier.
            ' This is done first since some Freight Terms require the Customer Absorbs the Freight Cost

            ASCMAIN1.Progress("Generating Shipping label(s)", "")

            If RequestShippingLabel(shipLabels, ErrorMessage, False) Then
                ' May want to print these later - we will see
                ' Moved below incase of error create invoice the user does not ship shipment without billing 
            Else
                ' ErrorMessage should have the error text available
                ' What to do if an error occurs in the requesting of a label??
                If ErrorMessage.Length > 0 Then
                    MessageBox.Show("Update Aborted! The following error occurred when processing the shipping label: " & ErrorMessage, "Ship Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Else
                    MessageBox.Show("Update Aborted! An error occurred when processing the shipping label.", "Ship Label", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                refreshScreen = False
                Exit Sub
            End If

            If RecreateLabel Then
                For Each shippingLabel As String In shipLabels
                    If shippingLabel.Trim.Length > 0 Then PrintLabel(shippingLabel)
                Next
                Exit Sub
            End If

            Dim shippingFreight As Decimal = 0

            Dim INV_STAX As Decimal = 0
            Dim CreditCardProcessed As Boolean = True
            Dim ship_ref As String = String.Empty

            ' Capture Credit Card Approved $$
            Dim rowSOTPICK1 As DataRow = Nothing
            Try
                dst.Tables("TATEVNT1").Rows.Clear()
                For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1' AND ISNULL(CCPA_NO_ORDR, '') <> ''", "SHIP_BOL_NO")

                    ASCMAIN1.Progress("Processing Credit Card", "")

                    Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                    Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                    Dim CCPA_NO_ORDR As String = rowSOTPICK1.Item("CCPA_NO_ORDR")

                    CCPA_NO_ORDR = CCPA_NO_ORDR.Trim
                    If CCPA_NO_ORDR.Length = 0 Then Continue For

                    If rowSOTPICK1.Item("CCPA_NO") & String.Empty <> String.Empty Then
                        ' Credit card against pick ticket already processed. May have been an error in the code
                        Continue For
                    End If

                    Dim chargeAmount As Decimal = Val(rowSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)
                    INV_STAX = 0

                    shippingFreight = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Val(rowSOTPICK1.Item("ORDR_FREIGHT") & String.Empty) ' + Val(rowSOTPICK1.Item("PPA_FREIGHT") & String.Empty)
                    chargeAmount += shippingFreight

                    If chargeAmount > 0 Then
                        Try
                            Dim ResponseText As String = String.Empty
                            Dim CCPA_NO As String = ProcessCreditCardAuthorization(CCPA_NO_ORDR, chargeAmount, shippingFreight, INV_STAX, ResponseText)
                            CreditCardProcessed = CCPA_NO.Length > 0 AndAlso CreditCardProcessed

                            If CCPA_NO.Length > 0 Then
                                ' This is done to preserve credit card transactions if the code causes an error after this point
                                MyBase.BeginTrans()
                                rowSOTPICK1.Item("CCPA_NO") = CCPA_NO
                                ASCDATA1.ExecuteSQL("Update SOTPICK1 set CCPA_NO = '" & CCPA_NO & "' where PICK_NO = '" & PICK_NO & "'")

                                ' Record Transaction Number in Order Header. Will be placed in Invoice Header
                                Dim rowARTCCPA1 As DataRow = LookUp("ARTCCPA1", CCPA_NO)
                                If rowARTCCPA1 IsNot Nothing Then
                                    Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                                    rowSOTORDR1.Item("CC_TRANS_ID") = rowARTCCPA1.Item("TRANS_ID")
                                End If

                                ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " _
                                     & " values " _
                                     & "('SOTORDR1', '" & rowSOTPICK1.Item("ORDR_NO") & "', SYSDATE, '" & ASCMAIN1.USER_ID & "', 'CCCHG','Credit card charged: " & Format(chargeAmount, "#,##0.00") & "', NULL)"
                                ASCDATA1.ExecuteSQL()
                                MyBase.CommitTrans()
                            Else
                                MessageBox.Show("Credit Card Could not be captured for the following reason: " & ResponseText, "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                CreditCardProcessed = False
                                Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
                                row.Item("TABLE_NAME") = "SOTORDR1"
                                row.Item("TABLE_KEY") = rowSOTPICK1.Item("ORDR_NO")
                                row.Item("INIT_DATE") = DATETIME_STAMP
                                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                row.Item("EVENT_TYPE") = "CCP"
                                row.Item("EVENT_DESC") = "Credit Card Error: " & ResponseText
                                row.Item("EVENT_KEY") = ""
                                dst.Tables("TATEVNT1").Rows.Add(row)
                            End If

                        Catch ex As Exception
                            MyBase.Rollback(ex.Message)
                            CreditCardProcessed = False
                            Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
                            row.Item("TABLE_NAME") = "SOTORDR1"
                            row.Item("TABLE_KEY") = rowSOTPICK1.Item("ORDR_NO")
                            row.Item("INIT_DATE") = DATETIME_STAMP
                            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                            row.Item("EVENT_TYPE") = "CCP"
                            row.Item("EVENT_DESC") = "Credit Card Error: " & ex.Message
                            row.Item("EVENT_KEY") = ""
                            dst.Tables("TATEVNT1").Rows.Add(row)
                        End Try
                    End If
                Next
            Catch ex As Exception
                MessageBox.Show("The following error occurred when processing the credit card: " & ex.Message, "Credit Card Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                CreditCardProcessed = False
                Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
                row.Item("TABLE_NAME") = "SOTORDR1"
                row.Item("TABLE_KEY") = rowSOTPICK1.Item("ORDR_NO")
                row.Item("INIT_DATE") = DATETIME_STAMP
                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                row.Item("EVENT_TYPE") = "CCP"
                row.Item("EVENT_DESC") = "Credit Card Error: " & ex.Message
                row.Item("EVENT_KEY") = ""
                dst.Tables("TATEVNT1").Rows.Add(row)
            End Try

            If CreditCardProcessed = False Then
                MessageBox.Show("Update aborted! Credit card could not be authorized.", "Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Dim EVENT_DESC_LEN As Int16 = dst.Tables("TATEVNT1").Columns("EVENT_DESC").MaxLength
                For Each rowTATEVNT1 As DataRow In dst.Tables("TATEVNT1").Select()
                    If (rowTATEVNT1.Item("EVENT_DESC") & String.Empty).ToString.Length > EVENT_DESC_LEN Then
                        rowTATEVNT1.Item("EVENT_DESC") = (rowTATEVNT1.Item("EVENT_DESC") & String.Empty).ToString.Substring(0, EVENT_DESC_LEN)
                    End If
                Next

                Try
                    BeginTrans()
                    Update_Record_TDA("TATEVNT1")
                    CommitTrans()
                Catch ex As Exception
                    Rollback()
                End Try
                Exit Sub
            End If

            ASCMAIN1.Progress("Now Updating ...", "")

            UpdateOrderPicksAndCreateInvoices()

            inTransaction = True
            BeginTrans()

            Update_SOTORDR5()

            Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            Dim WHSE_PHYS_STATUS As String = rowICTWHSE1.Item("WHSE_PHYS_STATUS") & ""
            Dim WHSE_LOCATOR As Boolean = rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1"

            Dim old_new_bols As String = ""
            Dim SHIP_BOL_NO_new As String

            rowSOTSHIP0.Item("SHIP_VIA_CODE") = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
            rowSOTSHIP0.Item("FRT_TERMS") = MyBase.Absx1.txtFor("FRT_TERMS").Text
            rowSOTSHIP0.Item("BILL_OF_LADING_NO") = MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Text
            rowSOTSHIP0.Item("ORDR_DEPT") = MyBase.Absx1.txtFor("ORDR_DEPT").Text

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

                rowSOTSHIP1.Item("SHIP_VIA_CODE") = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
                rowSOTSHIP1.Item("FRT_TERMS") = MyBase.Absx1.txtFor("FRT_TERMS").Text
                rowSOTSHIP1.Item("BILL_OF_LADING_NO") = MyBase.Absx1.txtFor("BILL_OF_LADING_NO").Text
                rowSOTSHIP1.Item("ORDR_DEPT") = MyBase.Absx1.txtFor("ORDR_DEPT").Text

                rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Value
                rowSOTSHIP1.Item("INV_DATE") = MyBase.Absx1.dteFor("INV_DATE").Value

                ship_ref = String.Empty
                If dst.Tables("WHTSHIP1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'").Length > 0 Then
                    ship_ref = dst.Tables("WHTSHIP1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")(0).Item("MASTER_TRACKING_NO") & String.Empty
                End If
                ' Pro Number
                rowSOTSHIP1.Item("SHIP_REF") = ship_ref
                rowSOTSHIP0.Item("SHIP_REF") = ship_ref

                Dim INV_DATE As Date = rowSOTSHIP1.Item("INV_DATE")
                Dim ORDR_YYYYPP_UPDATED As String = ASCDATA1.GetDataValue("Select MIN(OPS_YYYYPP) from gltparm2 where prd_end_date >= '" & INV_DATE.ToString("dd-MMM-yyyy") & "'") & String.Empty
                If ORDR_YYYYPP_UPDATED.Length = 0 Then
                    ORDR_YYYYPP_UPDATED = ASCMAIN1.CYP
                End If

                Dim sqlw As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                Dim sqlw_selected As String = sqlw & " and SELECTED = '1'"
                Dim T As DataTable = dst.Tables("SOTPICK1") ' .Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'").CopyToDataTable
                Dim SHIP_CNT_CARTONS As Int64 = Val(T.Compute("SUM(PICK_CNT_CARTONS)", sqlw_selected) & "")
                Dim SHIP_TOTAL_WGT As Decimal = Val(T.Compute("SUM(PICK_TOTAL_WGT)", sqlw_selected) & "")
                Dim SHIP_TOTAL_FRT As Decimal = Val(T.Compute("SUM(PICK_FREIGHT)", sqlw_selected) & "")
                Dim PICKS_SEL As Int64 = Val(T.Compute("Count(PICK_NO)", sqlw_selected) & "")
                Dim PICKS As Int64 = Val(T.Compute("Count(PICK_NO)", sqlw) & "")

                With rowSOTSHIP1
                    If PICKS_SEL > 0 Then
                        .Item("OPS_YYYYPP") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                        .Item("SHIP_CNT_CARTONS") = SHIP_CNT_CARTONS
                        .Item("SHIP_TOTAL_WGT") = SHIP_TOTAL_WGT
                        .Item("SHIP_STATUS") = "F"
                        For Each COLUMN_NAME As String In New String() _
                            {"SHIP_VIA_CODE", "SHIP_DATE_SHIPPED", "INV_DATE", "REASON_CODE", "TERM_CODE", _
                             "SREP_CODE", "SREP2_CODE", "ORDR_DEPT", "SHIP_REF", "SHIP_MANIFEST_NO", "BILL_OF_LADING_NO", "FRT_TERMS"}
                            .Item(COLUMN_NAME) = rowSOTSHIP0.Item(COLUMN_NAME)
                        Next

                        If PICKS_SEL <> PICKS Then
                            Dim rowSOTSHIP1_P As DataRow = dst.Tables("SOTSHIP1").NewRow
                            With rowSOTSHIP1_P
                                For i As Integer = 0 To dst.Tables("SOTSHIP1").Columns.Count - 1
                                    .Item(i) = rowSOTSHIP1.Item(i)
                                Next i
                                SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
                                old_new_bols = old_new_bols & vbCr & SHIP_BOL_NO & " -> " & SHIP_BOL_NO_new
                                .Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
                                .Item("SHIP_CNT_CARTONS") = 0
                                .Item("SHIP_TOTAL_WGT") = 0
                                .Item("SHIP_STATUS") = "P"
                                .Item("OPS_YYYYPP") = ""
                                For Each COLUMN_NAME As String In New String() _
                                    {"SHIP_VIA_CODE", "SHIP_DATE_SHIPPED", "INV_DATE", "REASON_CODE", "TERM_CODE", _
                                     "SREP_CODE", "SREP2_CODE", "ORDR_DEPT", "SHIP_REF", "SHIP_MANIFEST_NO", "BILL_OF_LADING_NO", "FRT_TERMS"}
                                    .Item(COLUMN_NAME) = rowSOTSHIP0_ORIG.Item(COLUMN_NAME)
                                Next
                            End With
                            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1_P)
                            sqlw = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and SELECTED <> '1'"
                            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select(sqlw)
                                rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
                            Next
                        End If
                    Else
                        .Item("SHIP_CNT_CARTONS") = 0
                        .Item("SHIP_TOTAL_WGT") = 0
                        .Item("SHIP_VIA_CODE") = ""
                        .Item("SHIP_DATE_SHIPPED") = DBNull.Value
                        .Item("FRT_TERMS") = ""
                        .Item("SHIP_REF") = ""
                        .Item("SHIP_MANIFEST_NO") = ""
                        .Item("BILL_OF_LADING_NO") = ""
                    End If
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID
                End With

            Next

            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("SOTORDR2")
            Update_Record_TDA("SOTORDR5")
            Update_Record_TDA("TATEVNT1")
            Update_Record_TDA("SOTORDXR")

            INIT_LAST("SOTSHIP1", False, , True)
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                ASCMAIN1.Progress("Updating Shipment " & SHIP_BOL_NO, "")
                Delete_Records(SHIP_BOL_NO, False)
            Next

            ' Copy Work Table Contents to Oracle

            For Each TABLE_NAME As String In New String() {"SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTCART1", "SOTCART2"}
                dst.Tables(TABLE_NAME).AcceptChanges()
                For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")
                    row.SetAdded()
                Next
                Update_Record_TDA(TABLE_NAME)
            Next

            For Each TABLE_NAME As String In New String() {"SOTINVH1", "SOTINVH2", "ARTOPEN1"}
                Update_Record_TDA(TABLE_NAME)
            Next


            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows
                ASCMAIN1.sql = "BEGIN SOPSTAT1('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "BEGIN SOPSTAT2('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()

                If WHSE_LOCATOR Then
                    TAC.ICCMAIN1.Update_WHTLOCBX("S", rowSOTINVH1.Item("INV_NO"))
                End If

                ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV", _
                   New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")}, _
                   New String() {"INV_TYPE_IN", "INV_NO_IN"})

            Next

            Dim order_header_updates_required As Boolean = False
            Dim SQLX As String = ""
            For Each COL As String In New String() {"SREP_CODE", "SREP2_CODE", "TERM_CODE", "ORDR_DEPT"}
                If rowSOTSHIP0_ORIG.Item(COL) & "" <> rowSOTSHIP0.Item(COL) & "" Then
                    order_header_updates_required = True
                    Exit For
                End If
            Next

            ' Process each BOL, now that it is in Oracle
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

                ASCMAIN1.sql = "Update SOTORDR1" _
                       & " Set ORDR_STATUS = 'F'" & vbCrLf _
                       & " where ORDR_NO in " & vbCrLf _
                       & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                       & " where SHIP_BOL_NO = :PARM1)" & vbCrLf
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {rowSOTSHIP1.Item("SHIP_BOL_NO")})

                ASCMAIN1.sql = "Update SOTPICK1" _
                        & " Set PICK_STATUS = 'F'" & vbCrLf _
                        & " where PICK_NO in " & vbCrLf _
                        & " (Select PICK_NO from SOTPICK1 " & vbCrLf _
                        & " where SHIP_BOL_NO = :PARM1)" & vbCrLf
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {rowSOTSHIP1.Item("SHIP_BOL_NO")})

                If order_header_updates_required Then
                    ASCMAIN1.sql = "Update SOTORDR1 " _
                         & "Set  TERM_CODE = :PARM1, ORDR_DEPT = :PARM2" & vbCrLf _
                         & " where ORDR_NO in " & vbCrLf _
                         & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                         & " where SHIP_BOL_NO = :PARM3)" & vbCrLf
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV", _
                                        New Object() {rowSOTSHIP1.Item("TERM_CODE"), _
                                                      rowSOTSHIP1.Item("ORDR_DEPT"), _
                                                      rowSOTSHIP1.Item("SHIP_BOL_NO")})
                End If

                ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "'); END;"
                ASCDATA1.ExecuteSQL()
            Next

            If old_new_bols <> "" Then
                MsgBox(old_new_bols, vbOKOnly, "Unshipped P/T's on the following BOL's have been assigned a New BOL No")
            End If

            CommitTrans()

            For Each shippingLabel As String In shipLabels
                If shippingLabel.Trim.Length > 0 Then PrintLabel(shippingLabel)
            Next

            ' Print Customs report
            ' 12/18/2015 - extened the if statement for ahava
            If Not txtCUST_COUNTRY.Text.ToUpper.Trim.StartsWith("US") AndAlso txtCUST_COUNTRY.Text.Trim = "" _
                AndAlso Not (ASCMAIN1.DBS_COMPANY = "AHA" OrElse ASCMAIN1.DBS_SERVER = "AHA") Then

                Dim InvNos As String = String.Empty
                For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select()
                    InvNos &= ", '" & rowSOTINVH1.Item("INV_NO") & "'"
                Next
                InvNos = InvNos.Substring(1).Trim
                Fill_Records("SOTINTL1", String.Empty, True, sqlSOTINTL1 & " and SOTINVH1.INV_NO IN (" & InvNos & ")")

                Dim RPT_TITLE As String = "International Customs Report"
                Dim reportFile As String = "SORINTL1"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Preparing " & RPT_TITLE)

                Print_Report_Begin()
                CR_params.Add("SUBT", "")

                Generate_Report(reportFile, RPT_TITLE)

                If ASCMAIN1.LaserPrinterName.Length > 0 Then
                    Print_Report_End(True, , ASCMAIN1.LaserPrinterName)
                Else
                    Print_Report_End()
                End If

                MessageBox.Show("Customs documents sent to printer: " & ASCMAIN1.LaserPrinterName, "Customs", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End If

            EmailInvoice()

            '*****************************************************************
        Catch ex As Exception
            If inTransaction Then Rollback()
            MessageBox.Show("Update Error: " & ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try

    End Sub

    Sub Update_Transfer(INV_NO As String)
        Dim XFR_NO As String = ASCDATA1.ExecuteSF _
                           ("SOPSHIP1_XFR", New String() {"INV_NO_IN"}, New Object() {INV_NO})
    End Sub

    Sub Update_Record_Maintenance()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")


        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSHIPX)
            ASCDATA1.ExecuteSQL("Truncate Table " & ASW("SOTCART1_3PL"))
            ASCDATA1.ExecuteSQL("Truncate Table " & ASW("SOTCART2_3PL"))
        End If


        BeginTrans()

        Dim SHIP_CHGREQ_NOs As New Dictionary(Of String, String)

        Dim date_changed As Boolean = False

        If Format(ORDR_SHIP_DATE, "yyyyMMdd") <> Format(dteORDR_SHIP_DATE.Value, "yyyyMMdd") _
        Or Format(ORDR_CANCEL_DATE, "yyyyMMdd") <> Format(dteORDR_CANCEL_DATE.Value, "yyyyMMdd") Then
            date_changed = True
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                Dim SHIP_CHGREQ_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP3.SHIP_CHGREQ_NO")
                SHIP_CHGREQ_NOs.Add(SHIP_BOL_NO, SHIP_CHGREQ_NO)

                ASCMAIN1.sql = "Update SOTORDR1 " _
                    & "Set ORDR_SHIP_DATE = :PARM1, ORDR_CANCEL_DATE = :PARM2" & vbCrLf _
                    & " where ORDR_NO in " _
                    & " (Select ORDR_NO from SOTPICK1 where SHIP_BOL_NO = :PARM3)" & vbCrLf
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DDV", _
                                    New Object() {dteORDR_SHIP_DATE.Value, _
                                                  dteORDR_CANCEL_DATE.Value, _
                                                  SHIP_BOL_NO})
            Next
        End If

        Dim qty_changed As Boolean = False
        Dim price_changed As Boolean = False
        Dim price_changed_to_Range As Boolean = False

        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select _
            ("", "", DataViewRowState.ModifiedCurrent)
            With rowSOTPICK2
                Dim QTY As Int64 = Val(.Item("PICK_QTY") & "") - Val(.Item("PICK_QTY_CONF") & "")
                If QTY <> 0 Then
                    qty_changed = True
                    .Item("PICK_QTY_CANC_REL") = Val(.Item("PICK_QTY_CANC_REL") & "") + QTY
                    .Item("PICK_QTY") = .Item("PICK_QTY_CONF")
                End If
                .Item("PICK_QTY_CONF") = DBNull.Value
                .Item("PICK_QTY_CANC") = DBNull.Value
                .Item("PICK_QTY_BACK") = DBNull.Value

                Dim SHIP_BOL_NO As String = .GetParentRow("SOTPICK1_SOTPICK2").Item("SHIP_BOL_NO")
                Dim rowSOTSHIP4 As DataRow = dst.Tables("SOTSHIP4").NewRow
                With rowSOTSHIP4
                    Dim SHIP_CHGREQ_NO As String
                    If SHIP_CHGREQ_NOs.ContainsKey(SHIP_BOL_NO) Then
                        SHIP_CHGREQ_NO = SHIP_CHGREQ_NOs(SHIP_BOL_NO)
                    Else
                        SHIP_CHGREQ_NO = ASCMAIN1.Next_Control_No("SOTSHIP3.SHIP_CHGREQ_NO")
                        SHIP_CHGREQ_NOs.Add(SHIP_BOL_NO, SHIP_CHGREQ_NO)
                    End If

                    .Item("SHIP_CHGREQ_NO") = SHIP_CHGREQ_NO
                    .Item("PICK_NO") = rowSOTPICK2.Item("PICK_NO")
                    .Item("PICK_LNO") = rowSOTPICK2.Item("PICK_LNO")
                    .Item("PICK_QTY_OLD") = rowSOTPICK2.Item("PICK_QTY", DataRowVersion.Original)
                    .Item("PICK_QTY_NEW") = rowSOTPICK2.Item("PICK_QTY")
                    .Item("PICK_UNIT_PRICE_OLD") = rowSOTPICK2.Item("PICK_UNIT_PRICE", DataRowVersion.Original)
                    .Item("PICK_UNIT_PRICE_NEW") = rowSOTPICK2.Item("PICK_UNIT_PRICE")

                    If Val(.Item("PICK_UNIT_PRICE_OLD") & "") <> Val(.Item("PICK_UNIT_PRICE_NEW") & "") Then
                        price_changed = True
                    End If

                    dst.Tables("SOTSHIP4").Rows.Add(rowSOTSHIP4)
                End With
            End With
        Next

        For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select _
            ("", "", DataViewRowState.ModifiedCurrent)
            With rowSOTCART2
                Dim QTY As Int64 = Val(.Item("QTY_PACKED") & "") - Val(.Item("QTY_PACKED_ORIG") & "")
                If QTY <> 0 Then
                    qty_changed = True
                    Dim SHIP_BOL_NO As String = dst.Tables("SOTPICK1").Rows.Find(.Item("PICK_NO")).Item("SHIP_BOL_NO")
                    Dim rowSOTSHIP6 As DataRow = dst.Tables("SOTSHIP6").NewRow
                    With rowSOTSHIP6
                        Dim SHIP_CHGREQ_NO As String
                        If SHIP_CHGREQ_NOs.ContainsKey(SHIP_BOL_NO) Then
                            SHIP_CHGREQ_NO = SHIP_CHGREQ_NOs(SHIP_BOL_NO)
                        Else
                            SHIP_CHGREQ_NO = ASCMAIN1.Next_Control_No("SOTSHIP3.SHIP_CHGREQ_NO")
                            SHIP_CHGREQ_NOs.Add(SHIP_BOL_NO, SHIP_CHGREQ_NO)
                        End If

                        .Item("SHIP_CHGREQ_NO") = SHIP_CHGREQ_NO
                        .Item("CART_NO") = rowSOTCART2.Item("CART_NO")
                        .Item("CART_LNO") = rowSOTCART2.Item("CART_LNO")
                        .Item("QTY_PACKED_OLD") = rowSOTCART2.Item("QTY_PACKED", DataRowVersion.Original)
                        .Item("QTY_PACKED_NEW") = rowSOTCART2.Item("QTY_PACKED")
                        dst.Tables("SOTSHIP6").Rows.Add(rowSOTSHIP6)
                    End With
                End If
            End With
        Next

        Update_SOTORDR5()

        'Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        'Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)

        If qty_changed Then
            ' Retract the Qty In Pick for each BOL - Before
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                Dependent_Updates(-1, SHIP_BOL_NO)
            Next
        End If

        ' Send changes - only modified rows will be updated, price may be updated even though Qty was not changed
        For Each TABLE_NAME As String In New String() {"SOTPICK2", "SOTCART2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        ' Restore the Qty In Pick for each BOL - After
        If qty_changed Then
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
                Dependent_Updates(-1, SHIP_BOL_NO)
            Next
        End If

        Dim LP_XNO As String = ""

        Dim sqlSHIP_CHG_REQ_NOs As String = ""
        For Each SHIP_BOL_NO As String In SHIP_CHGREQ_NOs.Keys
            Dim SHIP_CHGREQ_NO As String = SHIP_CHGREQ_NOs(SHIP_BOL_NO)
            sqlSHIP_CHG_REQ_NOs &= ",'" & SHIP_CHGREQ_NO & "'"
            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
            Dim rowSOTSHIP3 As DataRow = dst.Tables("SOTSHIP3").NewRow
            With rowSOTSHIP3
                .Item("SHIP_CHGREQ_NO") = SHIP_CHGREQ_NO
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("ORDR_SHIP_DATE_OLD") = ORDR_SHIP_DATE
                .Item("ORDR_CANCEL_DATE_OLD") = ORDR_CANCEL_DATE
                .Item("ORDR_SHIP_DATE_NEW") = dteORDR_SHIP_DATE.Value
                .Item("ORDR_CANCEL_DATE_NEW") = dteORDR_CANCEL_DATE.Value
                .Item("SHIP_CHGREQ_REASON") = txtReason.Text
                .Item("SHIP_CHGREQ_CONTACT") = txtContact.Text
                .Item("SHIP_CHGREQ_EMAIL") = txtemail.Text
            End With
            dst.Tables("SOTSHIP3").Rows.Add(rowSOTSHIP3)
        Next

        For Each TABLE_NAME As String In New String() _
            {"SOTSHIP3", "SOTSHIP4", "SOTSHIP6"}
            Update_Record_TDA(TABLE_NAME)
        Next

        If sqlSHIP_CHG_REQ_NOs <> "" Then

            ASCMAIN1.sql = "" _
                 & "Begin" & vbCrLf _
                 & " Declare Cursor C1 is " & vbCrLf _
                 & "  Select SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO, SOTSHIP4.PICK_UNIT_PRICE_NEW" & vbCrLf _
                 & "  , SOTSHIP4.PICK_QTY_OLD - SOTSHIP4.PICK_QTY_NEW QTY" & vbCrLf _
                 & "   from SOTPICK2,SOTSHIP4" & vbCrLf _
                 & "   where SOTPICK2.PICK_NO = SOTSHIP4.PICK_NO AND SOTPICK2.PICK_LNO = SOTSHIP4.PICK_LNO" & vbCrLf _
                 & "     and SOTSHIP4.SHIP_CHGREQ_NO IN (" & Mid(sqlSHIP_CHG_REQ_NOs, 2) & ");" & vbCrLf _
                 & " Begin" & vbCrLf _
                 & "  For R1 IN C1 Loop" & vbCrLf _
                 & "   Update SOTORDR2" & vbCrLf _
                 & "    Set ORDR_UNIT_PRICE = R1.PICK_UNIT_PRICE_NEW, ORDR_UNIT_PRICE_CURR = R1.PICK_UNIT_PRICE_NEW" & vbCrLf _
                 & "    , ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) - R1.QTY, ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + R1.QTY" & vbCrLf _
                 & "    where ORDR_NO = R1.ORDR_NO AND ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
                 & "  End Loop;" & vbCrLf _
                 & " End;" & vbCrLf _
                 & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        ' Group Record
        ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
        ASCDATA1.ExecuteSQL()
        CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_SOTORDR5()

        If Me.BindingContext.Contains(dvwSOTORDR5) Then
            ' Without the next 2 lines, data in text boxes in single row datatables (like header tables) will not get written to Oracle
            Dim X As CurrencyManager = Me.BindingContext(dvwSOTORDR5)
            X.EndCurrentEdit()
        End If

        For Each rowSOTORDR5 As DataRow In dst.Tables("TATEVNT1").Select("", "", DataViewRowState.ModifiedCurrent)
            Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
            row.Item("TABLE_NAME") = "SOTORDR1"
            row.Item("TABLE_KEY") = rowSOTORDR5.Item("ORDR_NO")
            row.Item("INIT_DATE") = DATETIME_STAMP
            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
            row.Item("EVENT_TYPE") = "SHPMTC"
            row.Item("EVENT_DESC") = "Ship-To Address was Changed"
            row.Item("EVENT_KEY") = ""
            dst.Tables("TATEVNT1").Rows.Add(row)

            ' Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowTATEVNT1.Item("ORDR_NO"))
            Dim REV_NO As Int16 = Val(ASCDATA1.GetDataValue("SELECT MAX(REV_NO) FROM SOTORDXR WHERE ORDR_NO = '" & rowSOTORDR5.Item("ORDR_NO") & "'")) + 1
            Dim REV_LNO As Int16 = 0
            For Each DC As DataColumn In dst.Tables("SOTORDR5").Columns
                Dim COLUMN_NAME As String = DC.ColumnName

                If rowSOTORDR5.Item(COLUMN_NAME) & "" <> rowSOTORDR5.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                    Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                    rowSOTORDXR.Item("ORDR_NO") = rowSOTORDR5.Item("ORDR_NO")
                    rowSOTORDXR.Item("REV_NO") = REV_NO
                    REV_LNO += 1
                    rowSOTORDXR.Item("REV_LNO") = REV_LNO
                    rowSOTORDXR.Item("INIT_DATE") = DATETIME_STAMP
                    rowSOTORDXR.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowSOTORDXR.Item("COLUMN_NAME") = COLUMN_NAME
                    rowSOTORDXR.Item("OLD_VALUE") = rowSOTORDR5.Item(COLUMN_NAME, DataRowVersion.Original)
                    rowSOTORDXR.Item("NEW_VALUE") = rowSOTORDR5.Item(COLUMN_NAME)
                    dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                End If
            Next
        Next
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SHIP_BOL_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTRSRV1.RSRV_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTRSRV1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTRSRV1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If

            Case "CUST_ADDR_CODE"
                sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
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

            Case "Load", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("SHIP_BOL_NO").Text = key
                Click_Command("Load")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTSHIP1"
            E.COLUMN_NAME = "SHIP_BOL_NO"
            E.CODE_VALUE = Absx1.txtFor("SHIP_BOL_NO").Text
            E.DESC_VALUE = "Shipment"
            E.ATTACHMENT_NOTES = ""
        End If

        Return E
    End Function

    Private Sub AddCarton()

        If grdSOTPICK1.ActiveRow Is Nothing Then
            MessageBox.Show("You must select a pick ticket.", "Add Carton", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
        rowSOTCART1.Item("CART_NO") = CART_NO ' "NEW" & Format(CART_NO_new, "0000000")
        rowSOTCART1.Item("PICK_NO") = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

        rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
        rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = 0
        rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 0
        rowSOTCART1.Item("CART_TRACKING_NO") = String.Empty
        'rowSOTCART1.Item("CART_SEQ") = String.Empty
        'rowSOTCART1.Item("CART_MEMO") = String.Empty
        'rowSOTCART1.Item("CART_TYPE") = String.Empty
        rowSOTCART1.Item("PACKAGING_TYPE") = 31
        rowSOTCART1.Item("PKG_CODE") = String.Empty
        Dim CART_SEQ As Int16 = Val(dst.Tables("SOTCART1").Compute("MAX(CART_SEQ)", "PICK_NO = '" & grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value & "'") & String.Empty) + 1
        rowSOTCART1.Item("CART_SEQ") = CART_SEQ

        grdSOTPICK1.ActiveRow.Cells("PICK_CNT_CARTONS").Value = dst.Tables("SOTCART1").Select("PICK_NO = '" & rowSOTCART1.Item("PICK_NO") & "'").Length
    End Sub

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTPICK1, "BBBSBB", "Select All", "De-Select All", "Propagate Value", "Hide Details", "Sales Order Inquiry", "Add Sales Tax")
        Load_Popup_Menu(grdSOTPICK2, "BS", "Item Status Inquiry", "Permit Price Change")

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
            Case "grdSOTPICK2"
                tlb_pop.Tools("Permit Price Change").SharedProps.Visible = Not InquiryMode And (EntryMode = "E" Or EntryMode = "N")
            Case "grdSOTPICK1"
                tlb_pop.Tools("Select All").SharedProps.Visible = Not InquiryMode AndAlso (EntryMode = "E" OrElse EntryMode = "N") AndAlso Not MaintenanceMode
                tlb_pop.Tools("De-Select All").SharedProps.Visible = Not InquiryMode AndAlso (EntryMode = "E" OrElse EntryMode = "N") AndAlso Not MaintenanceMode
                tlb_pop.Tools("Add Sales Tax").SharedProps.Visible = Not InquiryMode AndAlso (EntryMode = "E" OrElse EntryMode = "N") AndAlso Not MaintenanceMode
                'tlb_pop.Tools("Add Carton").SharedProps.Visible = False ' Not InquiryMode AndAlso (EntryMode = "E" OrElse EntryMode = "N") AndAlso Not MaintenanceMode
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '  e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdSOTPICK1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Propagate Value"), UltraWinToolbars.ButtonTool)
                    If Not (EntryMode = "E" Or EntryMode = "N") Or grd.ActiveCell Is Nothing OrElse Not New String() {"PICK_CNT_CARTONS", "PICK_TOTAL_WGT", "PICK_FREIGHT", "BILL_OF_LADING_NO", "ORDR_INV_COMMENT"}.Contains(grd.ActiveCell.Column.Key) Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        tlb_btn.SharedProps.Visible = True
                        tlb_btn.SharedProps.Caption = "Propagate Values for " & grd.ActiveCell.Column.Header.Caption
                        tlb_btn.Tag = grd.ActiveCell.Column.Key
                    End If
                    For Each ToolKey As String In New String() {"Hide Details"} ' {"Select All", "De-Select All", "Hide Details"}
                        DirectCast(tlb_pop.Tools(ToolKey), UltraWinToolbars.ButtonTool).SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")
                    Next
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
            Case "Add Sales Tax"
                If grdSOTPICK1.ActiveRow Is Nothing Then
                    Exit Sub
                End If
                Dim ORDR_STAX As Decimal = Val(grdSOTPICK1.ActiveRow.Cells("ORDR_STAX").Text & String.Empty)

                Dim input As String = InputBox("Enter Sales Tax Amount", "Sale Tax", ORDR_STAX)
                input = input.Trim
                If input.Length = 0 Then Exit Sub

                If Not IsNumeric(input) OrElse Val(input < 0) Then
                    MessageBox.Show("Invalid Sales Tax Entry", "Sales Tax", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                Else
                    grdSOTPICK1.ActiveRow.Cells("ORDR_STAX").Value = ((Val(input) * 100) \ 1) / 100
                End If

                Dim STAX_CODE As String = grdSOTPICK1.ActiveRow.Cells("STAX_CODE").Text & String.Empty

                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("STAX_CODE")
                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = False
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                    Using F As New ASFCODE1
                        F.ShowDialog()
                    End Using
                    STAX_CODE = ASCMAIN1.CodeSelector.SelectedCode
                End If

                If STAX_CODE.Length > 0 Then
                    grdSOTPICK1.ActiveRow.Cells("STAX_CODE").Value = STAX_CODE
                End If

            Case "Select All", "De-Select All"
                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                    rowSOTPICK1.Item("SELECTED") = IIf(e.Tool.Key = "Select All", "1", "0")
                Next
                Display_Totals()

            Case "Hide Details"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                splSOTPICK1.Panel2Collapsed = tlb_sbt.Checked

            Case "Permit Price Change"
                tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                With grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_UNIT_PRICE")
                    If tlb_sbt.Checked Then
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        .CellAppearance.BackColor = Drawing.Color.Empty
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        .CellAppearance.BackColor = Drawing.Color.Beige
                    End If
                End With

            Case "Add Carton"
                If 1 = 1 Then Exit Select

                Dim PICK_NO As String = grd.ActiveRow.Cells("PICK_NO").Text
                Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)

                Dim CART_NO As String = "MYCARTEZ"

                Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
                rowSOTCART1.Item("CART_NO") = CART_NO
                rowSOTCART1.Item("CART_FREIGHT") = 0
                'rowSOTCART1.Item("CART_PACKER") = String.Empty
                'rowSOTCART1.Item("CART_PACKED") = String.Empty
                'rowSOTCART1.Item("CART_SHIPPED") = String.Empty
                rowSOTCART1.Item("PICK_NO") = PICK_NO
                rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
                rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") = 0
                rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 0
                'rowSOTCART1.Item("CART_TRACKING_NO") = String.Empty
                'rowSOTCART1.Item("CART_SEQ") = String.Empty
                'rowSOTCART1.Item("CART_MEMO") = String.Empty
                'rowSOTCART1.Item("CART_TYPE") = String.Empty
                'rowSOTCART1.Item("PACKAGING_TYPE") = String.Empty
                'rowSOTCART1.Item("PKG_CODE") = String.Empty
                dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)

                Dim CART_LNO As Int16 = 1
                Dim CART_TOTAL_UNITS As Int16 = 0
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
                    Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                    rowSOTCART2.Item("CART_NO") = CART_NO
                    rowSOTCART2.Item("CART_LNO") = CART_LNO
                    CART_LNO += 1
                    rowSOTCART2.Item("ORDR_NO") = rowSOTPICK1.Item("ORDR_NO")
                    rowSOTCART2.Item("ORDR_LNO") = rowSOTPICK2.Item("ORDR_LNO")
                    rowSOTCART2.Item("QTY_PACKED") = 0
                    CART_TOTAL_UNITS += rowSOTCART2.Item("QTY_PACKED")
                    rowSOTCART2.Item("ITEM_UPC_CODE") = String.Empty
                    rowSOTCART2.Item("ITEM_EAN_CODE") = String.Empty
                    rowSOTCART2.Item("ITEM_CODE") = rowSOTPICK2.Item("ITEM_CODE")
                    rowSOTCART2.Item("PICK_NO") = PICK_NO
                    rowSOTCART2.Item("QTY_PACKED_ORIG") = Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty)
                    dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)
                Next
                rowSOTCART1.Item("CART_TOTAL_UNITS") = CART_TOTAL_UNITS
                If dst.Tables("sotcartx").Rows.Count = 0 Then Stop
                'SOTCARTX_SOTCART2

        End Select

        If grd Is Nothing Then
            Exit Sub
        Else
            If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
                Exit Sub
            End If
        End If

        Select Case e.Tool.Key

            Case "Propagate Value"
                Dim COLUMN_NAME As String = e.Tool.Tag
                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
                    rowSOTPICK1.Item(COLUMN_NAME) = grdSOTPICK1.ActiveCell.Value
                Next

                Display_Totals()

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("Select", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Sales Order Inquiry"
                Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value & ""
                Context_Launch("Load", ORDR_NO, e.Tool.Key, "SOFORDRI")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTSHIPX()
                End If

            Case "SHIP_BOL_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("Select")
                End If

            Case "PICK_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Dim PICK_NO As String = Absx1.txtFor("PICK_NO").Text.Trim
                    PICK_NO = ASCMAIN1.Format_Field(PICK_NO, "PICK_NO")
                    Dim rowSOTPICK1 As DataRow = LookUp("SOTPICK1", PICK_NO)
                    If rowSOTPICK1 Is Nothing Then
                        MsgBox("Invalid Pick Ticket No Specified (" & PICK_NO & ")", MsgBoxStyle.OkOnly, "Cannot Peform Requested Action")
                    Else
                        Absx1.txtFor("SHIP_BOL_NO").Text = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                        Click_Command("Select")
                    End If
                End If

            Case "CUST_ADDR_CODE"
                e.SuppressKeyPress = True
                ' e.Handled = True

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTSHIPX()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTSHIPX()

            Case "SHIPMENT_NO"
                Click_Command("Select")

            Case "CUST_ADDR_CODE"
                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, "MK", Absx1.txtFor("CUST_ADDR_CODE").Text})
                If rowARTCUST2 IsNot Nothing Then
                    Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
                    Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Rows.Find(New String() {ORDR_NO, "ST"}) ' .Select(dvwSOTORDR5.RowFilter)
                    For Each COLUMN_NAME As String In New String() _
                        {"CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", _
                         "CUST_COUNTRY", "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}

                        If COLUMN_NAME = "CUST_ADDR3" Then
                        Else
                            'rowSOTORDR5.Item(COLUMN_NAME) = rowARTCUST2.Item(COLUMN_NAME)
                            If COLUMN_NAME = "CUST_EXT" Or COLUMN_NAME = "CUST_FAX" Or COLUMN_NAME = "CUST_EMAIL" Then
                            Else
                                If COLUMN_NAME = "CUST_PHONE" Then
                                    Absx1.medFor("SOTORDR5." & COLUMN_NAME).Value = rowARTCUST2.Item(COLUMN_NAME) & ""
                                Else
                                    Absx1.txtFor("SOTORDR5." & COLUMN_NAME).Text = rowARTCUST2.Item(COLUMN_NAME) & ""
                                End If
                            End If
                        End If
                    Next
                End If

        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_DATE_SHIPPED"
                If EntryMode = "E" Then
                    'If Absx1.dteFor("INV_DATE").Value & "" = "" Then
                    Absx1.dteFor("INV_DATE").Value = Absx1.dteFor("SHIP_DATE_SHIPPED").Value
                    'End If
                End If
        End Select
    End Sub
#End Region

    Sub Load_SOTSHIPX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If InquiryMode Then
            ASCMAIN1.sql = sqlSOTSHIPX _
                & IIf(CUST_CODE = "", "", " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'")

            Select Case optStatus.Value
                Case "RNP"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is Null"
                    grdSOTSHIPX.Text = "Shipments Released not Printed"
                Case "PNC"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'P'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_PICK_PRINTED is Not Null"
                    grdSOTSHIPX.Text = "Shipments Printed not Confirmed"
                Case "C"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_STATUS = 'F'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_DATE_SHIPPED >= '" & Format(calFrom.Value, "dd-MMM-yyyy") & "'"
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_DATE_SHIPPED <= '" & Format(calTo.Value, "dd-MMM-yyyy") & "'"
                    grdSOTSHIPX.Text = "Shipments Confirmed as Shipped between " & calFrom.Value & " and " & calTo.Value
            End Select

            If CUST_CODE <> "" Then grdSOTSHIPX.Text &= " associated with " & CUST_CODE
            Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)

            Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
        Else
            If CUST_CODE = "" Then
                Fill_Records("SOTSHIPX")
                grdSOTSHIPX.Text = "Unconfirmed Shipments"
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            Else
                ASCMAIN1.sql = sqlSOTSHIPX _
                    & " and SOTORDR0.CUST_CODE = '" & CUST_CODE & "'"
                Fill_Records("SOTSHIPX", "", , ASCMAIN1.sql)
                grdSOTSHIPX.Text = "Unconfirmed Shipments associated with " & CUST_CODE
                Sort_grdColumns(grdSOTSHIPX, "SHIP_BOL_NO".ToLower)
            End If
        End If

        grdSOTSHIPX.Visible = True
    End Sub

    Private Sub grdSOTSHIPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTSHIPX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("SHIP_BOL_NO").Text = e.Row.Cells("SHIP_BOL_NO").Value
            Click_Command("Select")
        End If
    End Sub

    Sub Display_Totals()
        Dim KEY As Int32 = 0
        For Each COL As String In New String() _
            {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"}
            KEY += 1
            Dim rowSOTCONFT As DataRow = dst.Tables("SOTCONFT").Rows.Find(KEY)
            rowSOTCONFT.Item("QTY") = Val(dst.Tables("SOTPICK1").Compute("SUM(" & COL & ")", "SELECTED = '1'") & "")
            rowSOTCONFT.Item("AMT") = Val(dst.Tables("SOTPICK1").Compute("SUM(" & Replace(COL, "QTY", "AMT") & ")", "SELECTED = '1'") & "")
        Next
    End Sub

#Region "grdSOTPICK1"

    Private Sub grdSOTPICK1_AfterColPosChanged(sender As Object, e As Infragistics.Win.UltraWinGrid.AfterColPosChangedEventArgs) Handles grdSOTPICK1.AfterColPosChanged
        Position_txtSTORE()
    End Sub

    Private Sub grdSOTPICK1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowActivate
        Setup_SOTPICK1()
        Position_txtSTORE()
    End Sub

    Private Sub grdSOTPICK1_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdSOTPICK1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICK1.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdSOTPICK1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICK1.BeforeRowActivate
        If grdSOTCART1.ActiveRow IsNot Nothing AndAlso grdSOTCART1.ActiveRow.DataChanged Then
            grdSOTCART1.ActiveRow.Update()
        End If
        If grdSOTPICK2.ActiveRow IsNot Nothing AndAlso grdSOTPICK2.ActiveRow.DataChanged Then
            grdSOTPICK2.ActiveRow.Update()
        End If
    End Sub


    Private Sub grdSOTPICK1_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPICK1.BeforeRowUpdate
        If e.Row.Cells("PICK_STATUS").Value <> "P" Then
            MsgBox("You are attempting to make changes to a Pick Ticket" _
                    & vbCrLf & " that is Not In-Pick", _
                    MsgBoxStyle.OkOnly, "Changes are not Permitted to Pick Tickets NOT In Pick")
            e.Row.CancelUpdate()
            e.Cancel = True
        End If
    End Sub

    Private Sub grdSOTPICK1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK1.InitializeRow
        If e.Row.Cells("PICK_STATUS").Value = "F" Then
            e.Row.Cells("PICK_STATUS").Appearance.ForeColor = Drawing.Color.Blue
        ElseIf e.Row.Cells("PICK_STATUS").Value <> "P" Then
            e.Row.Cells("PICK_STATUS").Appearance.ForeColor = Drawing.Color.Red
        End If
    End Sub

    Private Sub grdSOTPICK1_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles grdSOTPICK1.MouseUp
        If grdSOTPICK1.ActiveCell IsNot Nothing AndAlso grdSOTPICK1.ActiveCell.Column.Key = "SELECTED" Then
            grdSOTPICK1.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdSOTPICK1_SizeChanged(sender As Object, e As System.EventArgs) Handles grdSOTPICK1.SizeChanged
        Position_txtSTORE()
    End Sub

    Sub Position_txtSTORE()

        If Not ScreenMode Then Exit Sub
        Try
            txtStore.Parent = grdSOTPICK1
            Dim r As System.Drawing.Rectangle = grdSOTPICK1.ActiveRowScrollRegion.FirstRow.Cells("CUST_STORE_NO").GetUIElement().ClipRect
            txtStore.Width = grdSOTPICK1.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Header.SizeResolved.Width
            txtStore.Left = r.Left
            txtStore.Top = grdSOTPICK1.Top

        Catch ex As Exception

        End Try
    End Sub

#End Region

#Region "grdSOTPICK2"

    Private Sub grdSOTPICK2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPICK2.AfterCellUpdate
        With grdSOTPICK1.ActiveRow
            Select Case e.Cell.Column.Key
                Case "PICK_UNIT_PRICE"

                Case "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"
                    If e.Cell.Tag = "X" Then Exit Sub
                    e.Cell.Tag = "X"
                    Dim PICK_QTY As Int64 = Val(e.Cell.Row.Cells("PICK_QTY").Value & "")
                    Dim PICK_QTY_CONF As Int64 = Val(e.Cell.Row.Cells("PICK_QTY_CONF").Value & "")
                    Dim PICK_QTY_CANC As Int64 = Val(e.Cell.Row.Cells("PICK_QTY_CANC").Value & "")
                    Dim PICK_QTY_BACK As Int64 = Val(e.Cell.Row.Cells("PICK_QTY_BACK").Value & "")

                    If PICK_QTY_CONF < PICK_QTY Then
                        If e.Cell.Column.Key = "PICK_QTY_BACK" Then
                            PICK_QTY_CANC = PICK_QTY - PICK_QTY_CONF - PICK_QTY_BACK
                            If PICK_QTY_CANC < 0 Then
                                PICK_QTY_CANC = 0
                            End If
                            e.Cell.Row.Cells("PICK_QTY_CANC").Value = PICK_QTY_CANC
                        Else
                            If chkBO.Checked Then

                                PICK_QTY_BACK = PICK_QTY - PICK_QTY_CONF - PICK_QTY_CANC

                                If PICK_QTY_BACK < 0 Then
                                    PICK_QTY_BACK = 0
                                End If
                                e.Cell.Row.Cells("PICK_QTY_BACK").Value = PICK_QTY_BACK
                            Else

                                PICK_QTY_CANC = PICK_QTY - PICK_QTY_CONF

                                If PICK_QTY_CANC < 0 Then
                                    PICK_QTY_CANC = 0
                                End If
                                e.Cell.Row.Cells("PICK_QTY_CANC").Value = PICK_QTY_CANC
                            End If
                        End If
                    Else
                        If PICK_QTY_CONF >= PICK_QTY Then
                            e.Cell.Row.Cells("PICK_QTY_CANC").Value = 0
                            e.Cell.Row.Cells("PICK_QTY_BACK").Value = 0
                        End If
                    End If
                    e.Cell.Tag = ""
            End Select
        End With
    End Sub

    Private Sub grdSOTPICK2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTPICK2.AfterRowActivate
        Dim ITEM_CODE As String = grdSOTPICK2.ActiveRow.Cells("ITEM_CODE").Value
        optSCB.ValueList.ValueListItems(1).DisplayText = "Item " & ITEM_CODE
        optSCB.ValueList.ValueListItems(1).Tag = "ITEM_CODE = '" & ITEM_CODE & "'"

        With grdSOTPICK2.ActiveRow
            clsPrice_Change = New Price_Change
            clsPrice_Change.PICK_NO = .Cells("PICK_NO").Value
            clsPrice_Change.PICK_LNO = .Cells("PICK_LNO").Value
            clsPrice_Change.ITEM_CODE = .Cells("ITEM_CODE").Value
            clsPrice_Change.PICK_UNIT_PRICE = .Cells("PICK_UNIT_PRICE").Value
        End With

        Setup_SOTCART2_from_SOTPICK2()
    End Sub

    Private Sub grdSOTPICK2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTPICK2.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdSOTPICK2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTPICK2.AfterRowUpdate
        If clsPrice_Change Is Nothing Then Exit Sub
        If clsPrice_Change.PICK_UNIT_PRICE <> e.Row.Cells("PICK_UNIT_PRICE").Value _
            And clsPrice_Change.PICK_NO = e.Row.Cells("PICK_NO").Value _
            And clsPrice_Change.PICK_LNO = e.Row.Cells("PICK_LNO").Value Then
            Dim sqlw As String = "(PICK_NO <> '" & clsPrice_Change.PICK_NO & "' or PICK_LNO <> " & CStr(clsPrice_Change.PICK_LNO) & ")" _
             & " and PICK_UNIT_PRICE = " & CStr(clsPrice_Change.PICK_UNIT_PRICE) _
             & " and ITEM_CODE = '" & clsPrice_Change.ITEM_CODE & "'"
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Changing Price for All Lines with Same Item")

            SOTPICK1_Expressions(True)
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw)
                rowSOTPICK2.Item("PICK_UNIT_PRICE") = e.Row.Cells("PICK_UNIT_PRICE").Value
            Next
            SOTPICK1_Expressions(False)

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End If
        clsPrice_Change = Nothing

        If MaintenanceMode Then
            Dim PICK_QTY_CONF As Int32 = Val(grdSOTPICK2.ActiveRow.Cells("PICK_QTY_CONF").Value & "")
            Dim PICK_NO As String = grdSOTPICK2.ActiveRow.Cells("PICK_NO").Value
            Dim PICK_LNO As String = Val(grdSOTPICK2.ActiveRow.Cells("PICK_LNO").Value & "")
            Dim sqlw As String = "PICK_NO = '" & PICK_NO & "'  and ORDR_LNO = " & CStr(PICK_LNO)
            Dim QTY_PACKED As Int32 = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", sqlw) & "")
            Dim LINES As Int32 = dst.Tables("SOTCART2").Select(sqlw).Length
            If LINES = 1 And PICK_QTY_CONF <> QTY_PACKED Then
                Dim row As DataRow = dst.Tables("SOTCART2").Select(sqlw)(0)
                row.Item("QTY_PACKED") = PICK_QTY_CONF
            End If
        End If

        Display_Totals()
    End Sub

    Sub SOTPICK1_Expressions(remove_expressions As Boolean)
        If remove_expressions Then
            expSOTPICK1.Clear()
            For Each fCOLUMN_NAME As String In New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", _
                                                             "PICK_AMT", "PICK_AMT_CONF", "PICK_AMT_CANC", "PICK_AMT_BACK", _
                                                             "PICK_TOTAL_WGT_CALC", "PICK_CNT_CARTONS_CALC", "PICK_TOTAL_UNITS_CALC"}
                expSOTPICK1.Add(fCOLUMN_NAME, dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression)
                dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression = ""
            Next
        Else
            For Each fCOLUMN_NAME As String In expSOTPICK1.Keys
                dst.Tables("SOTPICK1").Columns(fCOLUMN_NAME).Expression = expSOTPICK1(fCOLUMN_NAME)
            Next
        End If
    End Sub

    Private Sub grdSOTPICK2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTPICK2.BeforeExitEditMode
        Select Case grdSOTPICK2.ActiveCell.Column.Key
            Case "PICK_UNIT_PRICE"
                If Val(grdSOTPICK2.ActiveCell.Value & "") < 0 Then
                    e.Cancel = True
                End If
            Case "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK"
                If Val(grdSOTPICK2.ActiveCell.Value & "") < 0 Then
                    e.Cancel = True
                End If
        End Select
    End Sub

    Private Sub grdSOTPICK2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTPICK2.BeforeRowsDeleted

    End Sub

    Private Sub grdSOTPICK2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTPICK2.BeforeRowUpdate

        If grdSOTPICK1.ActiveRow.Cells("PICK_STATUS").Value <> "P" Then
            MsgBox("You are attempting to make changes to a Pick Ticket" _
                    & vbCrLf & " that is Not In-Pick", _
                    MsgBoxStyle.OkOnly, "Changes are not Permitted to Pick Tickets NOT In Pick")
            e.Row.CancelUpdate()
            e.Cancel = True
            Exit Sub
        End If

        Dim PICK_QTY As Int64 = Val(e.Row.Cells("PICK_QTY").Value)
        Dim PICK_QTY_CONF As Int64 = Val(e.Row.Cells("PICK_QTY_CONF").Value & "")
        Dim PICK_QTY_CANC As Int64 = Val(e.Row.Cells("PICK_QTY_CANC").Value & "")
        Dim PICK_QTY_BACK As Int64 = Val(e.Row.Cells("PICK_QTY_BACK").Value & "")
        Dim PICK_QTY_CANC_REL As Int64 = Val(e.Row.Cells("PICK_QTY_CANC_REL").Value & "")

        If MaintenanceMode Then
            If Val(e.Row.Cells("PICK_QTY_CONF").Value & "") > PICK_QTY + PICK_QTY_CANC_REL Then
                e.Row.Cells("PICK_QTY_CONF").Value = PICK_QTY - PICK_QTY_CANC - PICK_QTY_BACK
            End If
        End If

        If PICK_QTY_CONF < 0 Or PICK_QTY_BACK > PICK_QTY Or PICK_QTY_BACK < 0 Or PICK_QTY_CANC > PICK_QTY Or PICK_QTY_CANC < 0 Then
            e.Cancel = True
            Exit Sub
        End If

        If chkBO.Checked Then
            PICK_QTY_BACK = PICK_QTY - PICK_QTY_CONF - PICK_QTY_CANC
            If PICK_QTY_BACK < 0 Then
                PICK_QTY_BACK = 0
            End If
        Else
            PICK_QTY_CANC = PICK_QTY - PICK_QTY_CONF - PICK_QTY_BACK
            If PICK_QTY_CANC < -1 * PICK_QTY_CANC_REL Then
                PICK_QTY_CANC = 0
            End If
        End If
        If PICK_QTY_CONF > PICK_QTY + PICK_QTY_CANC_REL Then
            PICK_QTY_CANC = 0
            PICK_QTY_BACK = 0
        End If
        e.Row.Cells("PICK_QTY_CANC").Value = PICK_QTY_CANC
        e.Row.Cells("PICK_QTY_BACK").Value = PICK_QTY_BACK
    End Sub

    Private Sub grdSOTPICK2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTPICK2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key

            End Select
        End With

    End Sub
#End Region

#Region "grdSOTCART1"

    Private Sub grdSOTCART1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART1.InitializeRow
        If Val(e.Row.Cells("CART_TOTAL_UNITS_CALC").Value & "") <> Val(e.Row.Cells("CART_TOTAL_UNITS_ORIG").Value & "") Then
            e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("CART_TOTAL_UNITS_CALC").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Sub grdSOTCART1_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCART1.AfterCellUpdate

        Dim displayBoxAttributes As Boolean = False

        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In grdSOTCART1.Rows
            If row.Cells("PKG_CODE").Text = "OTHER" Then
                displayBoxAttributes = True
                Exit For
            End If
        Next

        If displayBoxAttributes Then
            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").Hidden = False
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").Hidden = False
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").Hidden = False

            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").Hidden = True
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").Hidden = True
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").Hidden = True

            grdSOTCART1.DisplayLayout.Bands(0).Columns("WIDTH").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("LENGTH").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("HEIGHT").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdSOTCART1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTCART1.AfterRowActivate
        Setup_SOTCART2_from_SOTCART1()
    End Sub

    Private Sub grdSOTCART1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.AfterRowUpdate
        Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Text

        Dim pickWeight As Decimal = Val(grdSOTPICK1.ActiveRow.Cells("PICK_TOTAL_WGT").Value & String.Empty)
        Dim totalCartWeight As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)

        If totalCartWeight > pickWeight Then
            grdSOTPICK1.ActiveRow.Cells("PICK_TOTAL_WGT").Value = totalCartWeight
            grdSOTPICK1.UpdateData()
        End If

    End Sub

    Private Sub grdSOTCART1_BeforeRowActivate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTCART1.BeforeRowActivate
        If grdSOTCART2.ActiveRow IsNot Nothing AndAlso grdSOTCART2.ActiveRow.DataChanged Then
            grdSOTCART2.ActiveRow.Update()
        End If
    End Sub

    Private Sub grdSOTCART1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTCART1.ClickCellButton
        If e.Cell.Column.Key = "CART_TOTAL_WGT_ACTUAL" Then
            registeredWeight = 0
            RequestWeightFromScale()
            e.Cell.Value = registeredWeight
        End If
    End Sub

#End Region

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

    Private Sub Force_Cartons_to_Balance()
        If dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED").Length <> 0 Then
            If MsgBox("Pick Ticket Details Have Been Found To Be Out Of Balance With Carton Details." & _
                   vbCrLf & "This Update Will Change The Cartons To Force Them to be In Balance With " & _
                   vbCrLf & "The Pick Tickets!" & vbCrLf & _
                   vbCrLf & "Are You SURE This is what you want?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then

                For Each rowSOTCARTX As DataRow In dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED")
                    Dim PICK_QTY_CONF As Int64 = rowSOTCARTX.Item("PICK_QTY_CONF")
                    Dim QTY_PACKED As Int64 = rowSOTCARTX.Item("QTY_PACKED")
                    Dim QTY As Int64 = PICK_QTY_CONF - QTY_PACKED
                    Dim sqlw As String = "ORDR_NO = '" & rowSOTCARTX.Item("ORDR_NO") & "' and ORDR_LNO = " & rowSOTCARTX.Item("ORDR_LNO")
                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select(sqlw, "CART_NO DESC")
                        If QTY = 0 Then Exit For
                        QTY_PACKED = Val(rowSOTCART2.Item("QTY_PACKED") & "")
                        If QTY > 0 Then
                            rowSOTCART2.Item("QTY_PACKED") = QTY_PACKED + QTY
                            QTY = 0
                        Else
                            If QTY_PACKED > System.Math.Abs(QTY) Then
                                rowSOTCART2.Item("QTY_PACKED") = QTY_PACKED + QTY
                                QTY = 0
                            Else
                                rowSOTCART2.Item("QTY_PACKED") = 0
                                QTY = QTY + QTY_PACKED
                            End If
                        End If
                    Next
                Next
            End If
        End If
    End Sub

    Private Sub Force_PTs_to_Balance()
        If dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED").Length <> 0 Then
            If MsgBox("Pick Ticket Details Have Been Found To Be Out Of Balance With Carton Details." & _
                   vbCrLf & "This Update Will Change The Pick Tickets To Force Them to be In Balance With " & _
                   vbCrLf & "The Cartons!" & vbCrLf & _
                   vbCrLf & "Are You SURE This is what you want?", MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then

                Dim dt As New DataTable
                For Each DC As DataColumn In dst.Tables("SOTCARTX").Columns
                    dt.Columns.Add(DC.ColumnName, DC.DataType)
                Next

                For Each rowSOTCARTX As DataRow In dst.Tables("SOTCARTX").Select("PICK_QTY_CONF <> QTY_PACKED")
                    Dim PICK_QTY_CONF As Int64 = rowSOTCARTX.Item("PICK_QTY_CONF")
                    Dim QTY_PACKED As Int64 = rowSOTCARTX.Item("QTY_PACKED")
                    Dim QTY As Int64 = PICK_QTY_CONF - QTY_PACKED
                    If QTY < 0 And "I DON'T KNOW WHY OR WHAT WE ARE DOING HERE - DT IS A TEMP TABLE" = "" Then
                        dt.Rows.Add(rowSOTCARTX.ItemArray)
                    Else
                        Dim sqlw As String = "ORDR_NO = '" & rowSOTCARTX.Item("ORDR_NO") & "' and ORDR_LNO = " & rowSOTCARTX.Item("ORDR_LNO")
                        Dim rowSOTPICK2s() As DataRow = dst.Tables("SOTPICK2").Select(sqlw, "PICK_NO DESC")
                        rowSOTPICK2s(0).Item("PICK_QTY_CONF") = QTY_PACKED
                        rowSOTPICK2s(0).Item("PICK_QTY_CANC") = Val(rowSOTPICK2s(0).Item("PICK_QTY_CANC") & "") + (PICK_QTY_CONF - QTY_PACKED)
                    End If
                Next

                If dt.Rows.Count <> 0 Then
                    Using F As New ASFMSGBF
                        F.Show_grd(dt, Me, "The Following Pick Ticket Lines Were Confirmed Higher Than The Original Qty's Released", "")
                    End Using
                End If
            End If
        End If
    End Sub

    Private Sub chkBO_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBO.CheckedChanged
        If chkBO.Checked Then
            If edi_customer Then
                If MsgBox("Override the Rule?", MsgBoxStyle.YesNo, _
                          "EDI Customers do not Allow Back Orders") = MsgBoxResult.No Then
                    chkBO.Checked = False
                End If
            End If
        End If
        Setup_BO()
    End Sub

    Sub Setup_BO()
        If chkBO.Checked Then
            cmdBACK.Enabled = True
            With grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_QTY_BACK")
                .CellActivation = UltraWinGrid.Activation.AllowEdit
                .Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                .CellAppearance.BackColor = Drawing.Color.Empty
            End With
        Else
            cmdBACK.Enabled = False
            With grdSOTPICK2.DisplayLayout.Bands(0).Columns("PICK_QTY_BACK")
                .CellActivation = UltraWinGrid.Activation.NoEdit
                .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                .CellAppearance.BackColor = Drawing.Color.Beige
            End With
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_QTY_BACK <> 0")
                rowSOTPICK2.Item("PICK_QTY_CANC") = Val(rowSOTPICK2.Item("PICK_QTY_CANC") & "") + Val(rowSOTPICK2.Item("PICK_QTY_BACK") & "")
                rowSOTPICK2.Item("PICK_QTY_BACK") = 0
            Next
        End If
    End Sub

    Sub Setup_SOTPICK1()
        If grdSOTPICK1.ActiveRow Is Nothing Then
            tabSOTPICK1.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK1.ActiveRow.Cells("PICK_NO").Value
            Dim ORDR_NO As String = grdSOTPICK1.ActiveRow.Cells("ORDR_NO").Value
            Dim CUST_STORE_NO As String = grdSOTPICK1.ActiveRow.Cells("CUST_STORE_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTPICK2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            grdSOTPICK2.Text = "Item Details for Pick No " & PICK_NO & ", Store " & CUST_STORE_NO
            optSCB.ValueList.ValueListItems(2).DisplayText = "Pick Ticket " & PICK_NO
            optSCB.ValueList.ValueListItems(2).Tag = "PICK_NO = '" & PICK_NO & "'"

            dvwSOTORDR5.RowFilter = "CUST_ADDR_TYPE = 'ST' and ORDR_NO = '" & ORDR_NO & "'"

            dvw = DirectCast(grdSOTCART1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "'"
            grdSOTCART1.Text = "Cartons for " & PICK_NO
            Setup_SOTCART2_from_SOTCART1()

            tabSOTPICK1.Visible = True

            If (EntryMode = "N" Or EntryMode = "E") Then
                If grdSOTPICK1.ActiveRow.Cells("PICK_STATUS").Value <> "P" AndAlso Not RecreateLabel Then
                    grdSOTPICK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdSOTCART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                Else
                    grdSOTPICK2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTCART1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdSOTCART2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                End If
            End If
        End If
    End Sub

    Sub Setup_SOTCART2_from_SOTCART1()
        If Not tabSOTPICK1.Tabs("Cartons").Visible Then Exit Sub
        If grdSOTCART1.ActiveRow Is Nothing Then
            grdSOTCART2.Visible = False
        Else
            Dim CART_NO As String = grdSOTCART1.ActiveRow.Cells("CART_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "CART_NO = '" & CART_NO & "'"
            grdSOTCART2.Text = "Contents of Carton " & CART_NO
            grdSOTCART2.Visible = True
        End If
    End Sub

    Sub Setup_SOTCART2_from_SOTPICK2()
        If tabSOTPICK1.Tabs("Cartons").Visible Then Exit Sub
        If grdSOTPICK2.ActiveRow Is Nothing Then
            grdSOTCART2.Visible = False
        Else
            Dim PICK_NO As String = grdSOTPICK2.ActiveRow.Cells("PICK_NO").Value
            Dim PICK_LNO As Int32 = Val(grdSOTPICK2.ActiveRow.Cells("PICK_LNO").Value & "")
            Dim dvw As DataView = DirectCast(grdSOTCART2.DataSource, DataTable).DefaultView
            dvw.RowFilter = "PICK_NO = '" & PICK_NO & "' and ORDR_LNO = " & CStr(PICK_LNO)
            grdSOTCART2.Text = "Cartons containing Items Indicated on Pick Ticket " & PICK_NO & ", Line " & CStr(PICK_LNO)
            grdSOTCART2.Visible = True
        End If
    End Sub

    Sub De_Confirm(SHIP_BOL_NO As String)

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
        Dim ORDR_GROUP_NO As String = rowSOTSHIP1.Item("ORDR_GROUP_NO")

        BeginTrans()
        Dim sqlw As String = "(Select INV_NO from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        ASCDATA1.ExecuteSQL("Delete from SOTINVH2 where INV_TYPE = 'I' and INV_NO in " & sqlw)
        ASCDATA1.ExecuteSQL("Delete from SOTINVH1 where INV_TYPE = 'I' and INV_NO in " & sqlw)

        ASCMAIN1.sql = "Update SOTPICK1 Set PICK_STATUS = 'P', PICK_SHIPPED = NULL, INV_NO = NULL" _
            & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "',LAST_DATE = SYSDATE" _
            & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update SOTSHIP1 Set SHIP_STATUS = 'P', SHIP_DATE_SHIPPED = NULL, LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
        ASCDATA1.ExecuteSQL()

        'Update New Control File To Force SJ&U between Confirm and Deconfirm - WR - 20051024
        ASCMAIN1.sql = "Update SOTCTLU1" _
            & " SET CTL_UPDATE_REQ = 'D'" _
            & " WHERE UPPER(CTL_KEY) = 'Z'"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Shipment " & SHIP_BOL_NO & " has been Successfully De-Confirmed")
    End Sub

    Sub Reverse_Invoice(SHIP_BOL_NO As String, INV_REVERSAL_REASON As String)

        Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
        Dim ORDR_GROUP_NO As String = rowSOTSHIP1.Item("ORDR_GROUP_NO")
        Dim REGISTER_XNO As String = rowSOTSHIP1.Item("REGISTER_XNO")

        Dim SHIP_BOL_NOs As New List(Of String)

        ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
            & " from SOTSHIP1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
            & " and REGISTER_XNO = '" & REGISTER_XNO & "'" & vbCrLf _
            & " and SHIP_STATUS = 'F'" & vbCrLf _
            & " and SHIP_BOL_NO_REV IS NULL"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        If DT.Rows.Count > 1 Then
            ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("SHIP_BOL_NO")
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.UseDataFromTable = DT
            ASCMAIN1.CodeSelector.Caption = "Please Select the Shipments to Reverse"
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                For Each SHIP_BOL_NO In ASCMAIN1.CodeSelector.SelectedCodes
                    SHIP_BOL_NOs.Add(SHIP_BOL_NO)
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If

            If ASCMAIN1.USER_ID = "angela" Or ASCMAIN1.USER_ID = "pat" Then
                If MsgBox("There Are Multiple BOLS In This Confirmation." _
                          & vbCrLf & vbCrLf _
                          & "Please Verify That The Reversal Has Gone Through Correctly" _
                          & vbCrLf _
                          & "By looking In The Customer Inquiry Screen After Completion." _
                          & vbCrLf & vbCrLf _
                          & "Are You Ready To Proceed?", MsgBoxStyle.YesNo, "Multiple BOLs") = MsgBoxResult.No Then
                    Exit Sub
                End If
            Else
                MsgBox("There Are Multiple BOLs In This Confirmation." _
                       & vbCrLf & "Please See Lenora To Proceed.", _
                       MsgBoxStyle.OkOnly, "Multiple BOLs")
                Exit Sub
            End If
            'Stop CMDEXECUTE CHECK ASSUMED THAT THERE WOULD BE ONLY 1 BOL
            'See SHIP0_REV_NOTES.txt in the Misc folder for further instructions.
        Else
            SHIP_BOL_NOs.Add(SHIP_BOL_NO)
        End If

        For Each SHIP_BOL_NO In SHIP_BOL_NOs
            Reverse_Invoice_1(SHIP_BOL_NO, dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO), INV_REVERSAL_REASON)
        Next

        ' Group Record
        ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Record Successfully Updated")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Reverse_Invoice_1(SHIP_BOL_NO As String, rowSOTSHIP1 As DataRow, INV_REVERSAL_REASON As String)
        Dim SHIP_BOL_NO_new As String
        Dim REGISTER_XNO As String = rowSOTSHIP1.Item("REGISTER_XNO")
        Dim INV_DATE As Date = rowSOTSHIP1.Item("INV_DATE")

        Stop ' WHY REGISTER XNO IN THE WHERE CLAUSE?
        Stop ' ISNT THIS DATA ALL HERE BY NOW?
        'SQL = "Select * from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        'Call Ora_to_Acc(Nothing, "SOWPICK1", 0, "X", SQL)
        'SQL = "Select SOTPICK2.*, SOTORDR2.RANGE_STYLE_LNO, SOTORDR2.QTY_PER_PP"
        'SQL = SQL & " from SOTPICK2, SOTORDR2 "
        'SQL = SQL & " where SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO"
        'SQL = SQL & " AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO"
        'SQL = SQL & " AND PICK_NO in (Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        'Call Ora_to_Acc(Nothing, "SOWPICK2", 0, "X", SQL)
        'SQL = "Select * from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and REGISTER_XNO = '" & REGISTER_XNO & "'"
        'Call Ora_to_Acc(Nothing, "SOWINVH1", 0, "X", SQL)
        'SQL = "Select * from SOTINVH2 where INV_TYPE = 'I' and INV_NO in (Select INV_NO from SOTINVH1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and REGISTER_XNO = '" & REGISTER_XNO & "')"
        'Call Ora_to_Acc(Nothing, "SOWINVH2", 0, "X", SQL)

        SHIP_BOL_NO_new = ASCMAIN1.Next_Control_No("SHIP_BOL_NO")
        rowSOTSHIP1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
        rowSOTSHIP1.Item("SHIP_856_BATCH_NO") = "N"
        rowSOTSHIP1.Item("SHIP_810_BATCH_NO") = "N"
        rowSOTSHIP1.Item("REGISTER_XNO") = ""
        rowSOTSHIP1.Item("SHIP_BOL_NO_REV") = SHIP_BOL_NO

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
            rowSOTPICK1.Item("PICK_NO_REV") = rowSOTPICK1.Item("PICK_NO")
        Next
        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("")
            For Each COLUMN_NAME As String In New String() {"PICK_QTY", "PICK_QTY_CONF", "PICK_QTY_CANC", "PICK_QTY_BACK", "PICK_QTY_CANC_REL", "PICK_QTY_BACK_REL"}
                rowSOTPICK2.Item(COLUMN_NAME) = -1 * Val(rowSOTPICK2.Item(COLUMN_NAME) & "")
            Next
        Next

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            rowSOTINVH1.Item("SHIP_BOL_NO") = SHIP_BOL_NO_new
            rowSOTINVH1.Item("INV_NO_REV") = rowSOTINVH1.Item("INV_NO")
            rowSOTINVH1.Item("ORDR_DATE_UPDATED") = DBNull.Value
            rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = DBNull.Value
            rowSOTINVH1.Item("INV_810_BATCH_NO") = DBNull.Value
            For Each COLUMN_NAME As String In New String() _
                {"INV_SALES", "INV_COGS", "INV_FREIGHT", "INV_MISC_CHG", "INV_TOTAL_AMOUNT"}
                rowSOTINVH1.Item(COLUMN_NAME) = -1 * Val(rowSOTINVH1.Item(COLUMN_NAME) & "")
            Next

            Dim SALES_DIVISION_CODE As String = rowSOTINVH1.Item("SALES_DIVISION_CODE")
            Dim PICK_NO As String = rowSOTINVH1.Item("PICK_NO")
            Dim PICK_NO_new As String = ASCMAIN1.Next_Control_No("PICK_NO", 10)
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim INV_NO_new As String = ASCMAIN1.Next_Control_No("INV_NO_01")

            rowSOTINVH1.Item("INV_NO") = INV_NO_new
            Stop ' do we have a relationship with SOTINVH2 to propagate this change down

            rowSOTINVH1.Item("PICK_NO") = PICK_NO_new
            rowSOTINVH1.Item("INV_COMMENT") = INV_REVERSAL_REASON

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")
                rowSOTPICK1.Item("PICK_NO") = PICK_NO_new
                rowSOTPICK1.Item("INV_NO") = INV_NO_new
            Next
            'For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
            '    rowSOTPICK1.Item("PICK_NO") = PICK_NO_new
            'Next

            ASCMAIN1.sql = "Update SOTINVH1 set INV_NO_REV_BY = '" & INV_NO_new & "'" _
                & " where INV_TYPE = 'I' AND INV_NO = '" & INV_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next
        For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select("")
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY_SHIP"}
                rowSOTINVH2.Item(COLUMN_NAME) = -1 * Val(rowSOTINVH2.Item(COLUMN_NAME) & "")
            Next
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH1").Select("INV_NO_CONS is Not Null"), "INV_NO_CONS, SALES_DIVISION_CODE").Rows
            Dim SALES_DIVISION_CODE As String = row.Item("SALES_DIVISION_CODE")
            Dim INV_NO As String = row.Item("INV_NO_CONS")
            Dim INV_NO_new As String = ASCMAIN1.Next_Control_No("INV_NO_01")
            Dim sqlw As String = "SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "' and INV_NO_CONS = '" & INV_NO & "'"
            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select(sqlw)
                rowSOTINVH1.Item("INV_NO_CONS") = INV_NO_new
            Next
        Next

        INIT_LAST("SOTSHIP1", False, "SHIP_BOL_NO = '" & SHIP_BOL_NO_new & "'")
        INIT_LAST("SOTPICK1", False, "SHIP_BOL_NO = '" & SHIP_BOL_NO_new & "'")
        For Each TABLE_NAME As String In New String() {"SOTSHIP1", "SOTPICK1", "SOTPICK2", "SOTINVH1", "SOTINVH2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        ASCMAIN1.sql = "Update SOTPICK1 set PICK_STATUS = 'P', PICK_SHIPPED = NULL, INV_NO = NULL" _
            & ", LAST_OPER = '" & ASCMAIN1.USER_ID & "',LAST_DATE = SYSDATE" _
            & " WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SOTSHIP1 SET SHIP_STATUS = 'P'" _
            & ", SHIP_DATE_SHIPPED = NULL" _
            & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
            & ", REGISTER_XNO = NULL" _
            & ", SHIP_810_BATCH_NO = NULL" _
            & ", SHIP_856_BATCH_NO = NULL" _
            & " WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Cancel_Shipment()
        Me.Cursor = Cursors.WaitCursor

        Try
            BeginTrans()

            '-	SOTSHIP1.SHIP_STATUS, SOTPICK1.PICK_STATUS, SOTORDR1.ORDR_STATUS,SOTORDR2.ORDR_STATUS
            '-	SOTPICK2: PICK_QTY gets copied to PICK_QTY_CANC, and PICK_QTY_CONF = 0
            '-	SOTORDR2: ORDR_QTY_PICK gets aggregated into ORDR_QTY_CANC
            '-	Retract SOTPICK2.PICK_QTY from ICTSTAT2
            '-	There may be more, check out the code from the original form

            Dim SHIP_BOL_NO As String = String.Empty
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                SHIP_BOL_NO = rowSOTSHIP1.Item("SHIP_BOL_NO")
                Dependent_Updates(-1, SHIP_BOL_NO)

                ASCMAIN1.sql = "" _
                    & "Begin" _
                    & " Declare Cursor C1 is" _
                    & "  Select SOTPICK2.* from SOTPICK2 " _
                    & "   where SOTPICK2.PICK_NO in (Select PICK_NO from SOTPICK1" _
                    & "     where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P') for Update;" _
                    & " Begin" _
                    & "  For R1 in C1 Loop" _
                    & "   Update SOTORDR2 " _
                    & "    Set ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + R1.PICK_QTY" _
                    & "      , ORDR_QTY_PICK = NVL(ORDR_QTY_PICK,0) - R1.PICK_QTY" _
                    & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" _
                    & "   Update SOTPICK2 Set PICK_QTY_CANC = PICK_QTY, PICK_QTY_CONF = 0 where Current of C1;" _
                    & "  End Loop;" _
                    & " End;" _
                    & "End;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_STATUS = 'C' where ORDR_NO in " _
                    & " (Select ORDR_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P')" _
                    & " and ORDR_STATUS = 'P' and ORDR_QTY_OPEN = 0 and ORDR_QTY_PICK = 0 and ORDR_QTY_CANC <> 0"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "" _
                    & "Update SOTORDR1 " _
                    & " Set ORDR_STATUS = 'C' where ORDR_NO in (" _
                    & "Select ORDR_NO from (" _
                    & "Select ORDR_NO" _
                    & ", SUM (DECODE(ORDR_STATUS,'O',1,0)) O" _
                    & ", SUM (DECODE(ORDR_STATUS,'P',1,0)) P" _
                    & ", SUM (DECODE(ORDR_STATUS,'C',1,0)) C" _
                    & ", SUM (DECODE(ORDR_STATUS,'F',1,0)) F" _
                    & " from SOTORDR2 where ORDR_NO in " _
                    & "(Select ORDR_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P')" _
                    & " group by ORDR_NO" _
                    & ") where O = 0 and P = 0 and F = 0 and C <> 0)"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTPICK1 Set PICK_STATUS = 'C'" _
                    & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS = 'P'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTSHIP1" _
                    & " Set SHIP_STATUS = 'C', SHIP_856_BATCH_NO = 'N', SHIP_810_BATCH_NO = 'N'" _
                    & " where SHIP_BOL_NO = :PARM1"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", SHIP_BOL_NO)
                ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & ORDR_GROUP_NO & "'); END;"
                ASCDATA1.ExecuteSQL()
            Next

            CommitTrans("Shipment has been Cancelled")

        Catch ex As Exception
            Rollback(ex.Message)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Sub Delete_Records(SHIP_BOL_NO As String, Optional sDependUpds As Boolean = True)

        If sDependUpds Then
            Dependent_Updates(-1, SHIP_BOL_NO)
        End If

        Dim sqlw As String = "where CART_NO in (" _
            & " Select CART_NO from SOTCART1 where PICK_NO in (" _
            & " Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'))"
        ASCDATA1.ExecuteSQL("Delete from SOTCART2 " & sqlw)
        ASCDATA1.ExecuteSQL("Delete from SOTCART1 " & sqlw)

        ASCMAIN1.sql = "Delete from SOTPICK2 where PICK_NO in " _
            & " (Select PICK_NO from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from SOTPICK1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from SOTSHIP1 where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_ICTSTAT2(ITEM_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVNNNNNN", _
                           New Object() {ITEM_CODE, WHSE_CODE, _
                                         0, 0, 0, _
                                         0, QTY, 0}, _
                           New String() {"ITEM_CODE_IN", "WHSE_CODE_IN", _
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in", _
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})

    End Sub

    Sub Dependent_Updates(S As Integer, SHIP_BOL_NO As String)
        ' If ASCMAIN1.Running_in_VS Then Stop
        Dim PICK_QTY As Int64 = 0
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

        ASCMAIN1.sql = "Select SOTORDR2.ITEM_CODE" _
            & ", SUM (SOTPICK2.PICK_QTY) PICK_QTY" _
            & " from SOTORDR2,SOTPICK2,SOTPICK1" _
            & " where SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
            & "   and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" _
            & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
            & "   and SOTPICK1.PICK_STATUS = 'P'" _
            & " group by SOTORDR2.ITEM_CODE"

        For Each rowSOTPICK2X As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ITEM_CODE As String = rowSOTPICK2X.Item("ITEM_CODE")
            PICK_QTY = Val(rowSOTPICK2X.Item("PICK_QTY") & "")
            If PICK_QTY <> 0 Then
                Update_ICTSTAT2(ITEM_CODE, WHSE_CODE, S * PICK_QTY)
            End If
        Next
    End Sub

    Private Sub txtStore_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtStore.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            If txtStore.Text = "" Then
                MsgBox("You Must First Enter a Store No", MsgBoxStyle.OkOnly, "Cannot Locate Pick Ticket for Selected Store")
                Exit Sub
            Else
                txtStore.Text = txtStore.Text.PadLeft(6, "0")
            End If

            grdSOTPICK1.ActiveRow = Nothing
            grdSOTPICK1.Selected.Rows.Clear()
            For Each grow As UltraWinGrid.UltraGridRow In grdSOTPICK1.Rows
                If grow.Cells("CUST_STORE_NO").Value & "" = txtStore.Text Then
                    grdSOTPICK1.ActiveRow = grow
                    grow.Selected = True
                    Exit For
                End If
            Next
            If grdSOTPICK1.ActiveRow Is Nothing Then
                MsgBox("No Pick Ticket Found for Store " & txtStore.Text, MsgBoxStyle.OkOnly, "Cannot Locate Pick Ticket for Selected Store")
            End If
            txtStore.Text = ""
        End If
    End Sub

    Private Sub cmdSHIP_Click(sender As System.Object, e As System.EventArgs) Handles cmdSHIP.Click
        SCB("PICK_QTY_CONF")
    End Sub

    Private Sub cmdCANC_Click(sender As System.Object, e As System.EventArgs) Handles cmdCANC.Click
        SCB("PICK_QTY_CANC")
    End Sub

    Private Sub cmdBACK_Click(sender As System.Object, e As System.EventArgs) Handles cmdBACK.Click
        SCB("PICK_QTY_BACK")
    End Sub

    Sub SCB(COLUMN_NAME As String)
        Dim sqlw As String = ""
        If optSCB.Value = "SHIP_BOL_NO" Then
        ElseIf optSCB.Value = "ITEM_CODE" Then
            ' Stop
            sqlw = optSCB.ValueList.ValueListItems(1).Tag
        ElseIf optSCB.Value = "PICK_NO" Then
            sqlw = optSCB.ValueList.ValueListItems(2).Tag
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Changing Pick Ticket Details Indicated")

        SOTPICK1_Expressions(True)
        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select(sqlw)
            rowSOTPICK2.Item("PICK_QTY_CONF") = 0
            rowSOTPICK2.Item("PICK_QTY_CANC") = 0
            rowSOTPICK2.Item("PICK_QTY_BACK") = 0
            rowSOTPICK2.Item(COLUMN_NAME) = rowSOTPICK2.Item("PICK_QTY")
        Next
        SOTPICK1_Expressions(False)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        grdSOTPICK1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        Display_Totals()
    End Sub

    Private Sub grdSOTSHIP1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSOTSHIP1.InitializeLayout

    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        grpConfirmed.Visible = (optStatus.Value = "C")
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTSHIPX()
    End Sub

    Private Sub btnLoadHistory_Click(sender As System.Object, e As System.EventArgs) Handles btnLoadHistory.Click
        Load_SOTSHIPX()
    End Sub

    Sub ToggleDataTableExpressions(ByVal tf As Boolean)

        With dst.Tables("SOTPICK2")
            .Columns("PICK_AMT").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY,0)")
            .Columns("PICK_AMT_CONF").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_CONF,0)")
            .Columns("PICK_AMT_CANC").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_CANC,0)")
            .Columns("PICK_AMT_BACK").Expression = IIf(Not tf, "", "ISNULL(PICK_UNIT_PRICE,0) * ISNULL(PICK_QTY_BACK,0)")
        End With

        With dst.Tables("SOTCARTX")
            .Columns("PICK_QTY_CONF").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCARTX_SOTPICK2).PICK_QTY_CONF)")
            .Columns("QTY_PACKED").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCARTX_SOTCART2).QTY_PACKED)")
        End With

        With dst.Tables("SOTCART1")
            .Columns("CART_TOTAL_UNITS_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
        End With

        With dst.Tables("SOTPICK1")
            .Columns("PICK_TOTAL_WGT_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_ACTUAL)")
            .Columns("PICK_CNT_CARTONS_CALC").Expression = IIf(Not tf, "", "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
            .Columns("PICK_TOTAL_UNITS_CALC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_UNITS_CALC)")

            .Columns("PICK_QTY").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY)")
            .Columns("PICK_QTY_CONF").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_CONF)")
            .Columns("PICK_QTY_CANC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_CANC)")
            .Columns("PICK_QTY_BACK").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_QTY_BACK)")
            .Columns("PICK_AMT").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT)")
            .Columns("PICK_AMT_CONF").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_CONF)")
            .Columns("PICK_AMT_CANC").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_CANC)")
            .Columns("PICK_AMT_BACK").Expression = IIf(Not tf, "", "SUM(CHILD(SOTPICK1_SOTPICK2).PICK_AMT_BACK)")
        End With

    End Sub

    Private Function PICK_NO() As Object
        Throw New NotImplementedException
    End Function

    Private Sub optShipmentSelection_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optShipmentSelection.ValueChanged

    End Sub

    Private Sub grdSOTPICK2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTPICK2.InitializeRow
        If Val(e.Row.Cells("PICK_QTY").Value & "") <> Val(e.Row.Cells("PICK_QTY_CONF").Value & "") Then
            e.Row.Cells("PICK_QTY_CONF").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("PICK_QTY_CONF").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If

        If Val(e.Row.Cells("PICK_UNIT_PRICE").Value & "") <> Val(e.Row.Cells("ORDR_UNIT_PRICE").Value & "") Then
            e.Row.Cells("PICK_UNIT_PRICE").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("PICK_UNIT_PRICE").Appearance.ForeColor = Drawing.Color.Empty
        End If

    End Sub



    Private Sub grdSOTCART2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTCART2.InitializeRow
        If Val(e.Row.Cells("QTY_PACKED").Value & "") <> Val(e.Row.Cells("QTY_PACKED_ORIG").Value & "") Then
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Red
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
        Else
            e.Row.Cells("QTY_PACKED").Appearance.ForeColor = Drawing.Color.Empty
            e.Row.Appearance.BackColor = Drawing.Color.Empty
        End If
    End Sub

    Private Function ProcessCreditCardAuthorization(ByVal AUTH_CCPA_NO As String, _
                                                    ByVal ChargeAmount As Double, _
                                                    ByVal freightAmount As Decimal, _
                                                    ByVal salesTax As Decimal, _
                                                    ByRef ResponseText As String) As String

        Dim sql As String = String.Empty
        Dim ORDR_UNIT_PRICE As Decimal = 0

        AUTH_CCPA_NO = AUTH_CCPA_NO.Trim
        If AUTH_CCPA_NO.Length = 0 Then Return String.Empty

        If ChargeAmount <= 0 Then Return String.Empty
        ChargeAmount = Math.Round(ChargeAmount, 2)

        ASCMAIN1.Progress("Processing Credit Card", String.Empty)

        MyBase.Fill_Records("ARTCCPA1", AUTH_CCPA_NO)
        MyBase.Fill_Records("ARTCCPDA", AUTH_CCPA_NO)

        If dst.Tables("ARTCCPA1").Rows.Count <> 1 Then Return String.Empty
        If dst.Tables("ARTCCPDA").Rows.Count <> 1 Then Return String.Empty

        Dim rowARTCCPA1_AUTH As DataRow = dst.Tables("ARTCCPA1").Rows(0)

        Dim AUTH_RESPONSE_APPROVAL_CODE As String = (rowARTCCPA1_AUTH.Item("RESPONSE_APPROVAL_CODE") & String.Empty).ToString.Trim
        If AUTH_RESPONSE_APPROVAL_CODE.Length = 0 Then Return String.Empty

        Dim CCPA_NO As String = String.Empty


        Try
            Me.CreditCardProcessor = New TAC.TAFCARDF(Me)

            '******************************************************************************************************************
            ' Default to Authorize
            ' If it is the case the Invoice Total Amount is greater than the Approved Amount
            ' Then Void the Original Auth and process as a Sale.
            ' If the Void Authorize fails, continue and process as a sale.
            Dim chargeType As String = "A"

            Dim OriginalAuthAmount As Decimal = Val(rowARTCCPA1_AUTH.Item("CCPA_AMT") & String.Empty)
            If ChargeAmount > OriginalAuthAmount Then
                chargeType = "S"
                Try
                    CreditCardProcessor.MerchantSetup()
                    CreditCardProcessor.objCCProcessor.VoidTransaction(rowARTCCPA1_AUTH.Item("TRANS_ID") & String.Empty, "1")
                Catch ex As Exception
                    ResponseText = "Error trying to void initial CC Authorization: " & ex.Message
                    'Return String.Empty
                End Try
            End If
            '******************************************************************************************************************

            With Me.CreditCardProcessor
                .ORDR_NO = rowARTCCPA1_AUTH.Item("ORDR_NO") & String.Empty

                .objCCProcessor.TransactionNumber = ASCMAIN1.Next_Control_No("ARTCCPA1.TRANS_NUM")
                .objCCProcessor.TransactionAmount = ChargeAmount
                .objCCProcessor.CustomerCreditCard.CardNumber = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NO") & String.Empty

                Dim CUST_CREDIT_CARD_EXP_DATE As String = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_EXP_DATE") & String.Empty
                CUST_CREDIT_CARD_EXP_DATE = CUST_CREDIT_CARD_EXP_DATE.PadRight(4, "0")
                .objCCProcessor.CustomerCreditCard.CardExpMonth = CUST_CREDIT_CARD_EXP_DATE.Substring(0, 2)
                .objCCProcessor.CustomerCreditCard.CardExpYear = CUST_CREDIT_CARD_EXP_DATE.Substring(2)
                .objCCProcessor.ValidateCard()

                If chargeType = "S" Then
                    .objCCProcessor.CreditCardProcessingNo = CCPA_NO
                    .objCCProcessor.InternalReference = "Customer: " & rowARTCCPA1_AUTH.Item("CUST_CODE") & ", TransType: " & "S"

                    .objCCProcessor.CustomerCreditCard.CardHolderFirstName = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_NAME") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderLastName = "" 'Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text
                    .objCCProcessor.CustomerCreditCard.CardHolderAddress = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ADDR1") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderCity = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_CITY") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderState = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_STATE") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderZipCode = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_ZIP_CODE") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderCountry = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_COUNTRY") & String.Empty
                    .objCCProcessor.CustomerCreditCard.CardHolderTelephone = "" 'Absx1.txtFor("CUST_CREDIT_CARD_ADDR1").Text
                    .objCCProcessor.CustomerCreditCard.CardCVVData = rowARTCCPA1_AUTH.Item("CUST_CREDIT_CARD_VER_CODE") & String.Empty
                End If

                With .objCCProcessor.Level2Data
                    .Clear()

                    .CardType = CreditCardProcessor.objCCProcessor.CreditCardType

                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowARTCCPA1_AUTH.Item("ORDR_NO")) ' dst.Tables("SOTORDR1").Rows(0)
                    Dim rowSHIPTO As DataRow = dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)

                    If rowSHIPTO IsNot Nothing Then
                        .DestinationZip = rowSHIPTO.Item("CUST_ZIP_CODE") & String.Empty
                        .DestinationState = rowSHIPTO.Item("CUST_STATE") & String.Empty
                    End If

                    .DiscountAmount = 0
                    .FreightAmount = freightAmount
                    .InvoiceNumber = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    .OrderDate = rowSOTORDR1.Item("ORDR_DATE") & String.Empty
                    .PurchaseIdentifier = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    .TaxAmount = salesTax

                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                    If rowICTWHSE1 IsNot Nothing Then
                        .ShipFromZip = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
                    End If

                End With

                With .objCCProcessor.Level3Data
                    .Clear()
                    Dim ITEM_CODE As String = String.Empty
                    Dim Quantity As Integer = 0
                    Dim Description As String = String.Empty

                    For Each rowSOTPICK2 As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTPICK2"), New String() {"ITEM_CODE"}).Rows
                        ITEM_CODE = rowSOTPICK2.Item("ITEM_CODE") & String.Empty
                        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)

                        Dim rowSOTPICK2X As DataRow = dst.Tables("SOTPICK2").Select("ITEM_CODE = '" & ITEM_CODE & "'", "PICK_UNIT_PRICE DESC")(0)

                        ORDR_UNIT_PRICE = Val(rowSOTPICK2X.Item("PICK_UNIT_PRICE") & String.Empty)
                        'If ORDR_UNIT_PRICE <= 0 Then Continue For

                        Quantity = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY_CONF)", "ITEM_CODE = '" & ITEM_CODE & "' AND PICK_UNIT_PRICE = " & ORDR_UNIT_PRICE) & String.Empty)
                        If Quantity <= 0 Then Continue For

                        Dim level3 As New TAC.ARCCCARD.Level3
                        With level3
                            .Description = StrConv(rowICTITEM1.Item("ITEM_DESC") & String.Empty, VbStrConv.ProperCase)
                            .DiscountAmount = 0
                            .ProductCode = ITEM_CODE
                            .Quantity = Quantity
                            .TaxAmount = 0
                            .TaxType = TAC.ARCCCARD.TaxTypes.StateSalesTax
                            .UnitCost = ORDR_UNIT_PRICE
                            .Units = "each"
                            .Total = .Quantity * .UnitCost
                            .TaxAmount = Math.Round(.Total * .TaxRate / 100, 2, MidpointRounding.AwayFromZero)
                        End With
                        .Add(level3)
                    Next
                End With

                .rowARTCCPA1 = rowARTCCPA1_AUTH
                If chargeType = "S" Then
                    CCPA_NO = .CC_Sale(ChargeAmount)
                Else
                    CCPA_NO = .CC_Capture(ChargeAmount)
                End If

                ResponseText = .responseErrorMessage

                ' Need to see if the Authorizations fell out of scope. If so, then Do a sale
                If CCPA_NO = String.Empty AndAlso chargeType <> "S" AndAlso Me.CreditCardProcessor.objCCProcessor.NetworkResponse.ResponseCode = "3" Then
                    chargeType = "S"
                    CCPA_NO = .CC_Sale(ChargeAmount)
                    ResponseText = .responseErrorMessage
                End If

            End With

        Catch ex As Exception
            MessageBox.Show("The following error occurred processing a credit card: " & ex.Message, "Charge Credit Card", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            CreditCardProcessor.Dispose()
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

        Return CCPA_NO

    End Function

    ''' <summary>
    ''' Requests Shipping labels for carriers such as fedex, Ups, USPS
    ''' </summary>
    ''' <param name="ShippingLabels">String list of the labels created. Print these</param>
    ''' <param name="ErrorMessage">Error that occurred when processing request</param>
    ''' <param name="PreScreenForErrorsOnly">Boolen to determin if all requirements are meet. If true only evaluation is done.
    ''' Any errors or missing attributes will be returned in the ErrorMessage Parameter</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function RequestShippingLabel(ByRef ShippingLabels As List(Of String), ByRef ErrorMessage As String, ByVal PreScreenForErrorsOnly As Boolean) As Boolean

        'Do Not Print Labels for Ahava
        ' 12/18/2015 - Manually Process / Invoice Non ADS (3pl) shipments.
        If ASCMAIN1.DBS_COMPANY = "AHA" OrElse ASCMAIN1.DBS_SERVER = "AHA" Then
            Return True
        End If

        Dim createCarrierLabels As Boolean = False
        ErrorMessage = String.Empty

        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowSOTCARR1 As DataRow = Nothing
        Dim rowSOTPICK1 As DataRow = Nothing

        Dim SHIP_VIA_CODE As String = String.Empty
        Dim SHIP_PACKAGE_NO As Int64 = 0
        Dim pkgId As Int64 = 0

        Try
            rowSOTSHIP1 = dst.Tables("SOTSHIP1").Rows.Find(MyBase.Absx1.txtFor("SHIP_BOL_NO").Text)
            SHIP_VIA_CODE = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text ' rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty

            rowSOTSVIA1 = LookUp("SOTSVIA1", SHIP_VIA_CODE)
            If rowSOTSVIA1 IsNot Nothing Then
                rowSOTCARR1 = LookUp("SOTCARR1", rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty)
                If rowSOTCARR1 IsNot Nothing AndAlso rowSOTCARR1.Item("CARRIER_TYPE") = "U" Then
                    createCarrierLabels = True
                End If
            Else
                ErrorMessage = "Invalid or missing Ship Via for shipping label request"
            End If

        Catch ex As Exception
            ErrorMessage = "The following error occurred when evaluating a shipping label request: " & ex.Message
            Return False
        End Try

        ' Returns False since there is nothing to do. False with ErrorMessage indicates as an error occurred.
        If Not createCarrierLabels Then Return False

        RequestShippingLabel = True
        Try

            ' Load and validate Customer
            Dim CUST_CODE As String = MyBase.Absx1.txtFor("CUST_CODE").Text
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            If rowARTCUST1 Is Nothing Then
                ErrorMessage = "Invalid or missing Customer Code for shipping label request"
                Return False
            Else
                txt3PAccountNo.Text = txt3PAccountNo.Text.Trim
                txt3pCountry.Text = txt3pCountry.Text.Trim.ToUpper
                txt3PZipCode.Text = txt3PZipCode.Text.Trim

                ' Prepopulate ant Account numbers if the user did not provide them
                Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                    Case "F"
                        If txt3PAccountNo.TextLength = 0 Then txt3PAccountNo.Text = (rowARTCUST1.Item("FDX_ACCT_NO") & String.Empty).ToString.Trim
                        If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("CUST_COUNTRY") & String.Empty).ToString.Trim
                        If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
                    Case "U"
                        If txt3PAccountNo.TextLength = 0 Then txt3PAccountNo.Text = (rowARTCUST1.Item("UPS_ACCT_NO") & String.Empty).ToString.Trim
                        If txt3pCountry.TextLength = 0 Then txt3pCountry.Text = (rowARTCUST1.Item("CUST_COUNTRY") & String.Empty).ToString.Trim
                        If txt3PZipCode.TextLength = 0 Then txt3PZipCode.Text = (rowARTCUST1.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
                End Select
            End If

            Dim CARRIER_CODE As String = rowSOTSVIA1.Item("CARRIER_CODE") & String.Empty
            Dim CARRIER_PROD_CODE As String = rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty

            ' Load and Validate Carrier/Ship Method
            Dim rowSOTCARR2 As DataRow = LookUp("SOTCARR2", New String() {CARRIER_CODE, CARRIER_PROD_CODE})
            If rowSOTCARR2 Is Nothing Then
                ErrorMessage = "Invalid or missing Carrier / Ship Method combination for shipping label request"
                Return False
            End If

            ' Credentials
            Dim rowSOTCARR3 As DataRow = dst.Tables("SOTCARR3").Select("CARRIER_CODE = '" & CARRIER_CODE & "'")(0)
            Dim ShippingLabelDirectory As String = (rowSOTCARR1.Item("CARRIER_ARCHIVE_DIR") & String.Empty).ToString.Trim
            Dim PROVIDER_TYPE As String = (rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty).ToString.Trim

            If rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty = String.Empty Then
                ErrorMessage = "Invalid or missing Carrier Account Number for shipping label request"
                Return False
            End If

            Try
                If ASCMAIN1.Running_in_VS Then
                    'ShippingLabelDirectory = ShippingLabelDirectory.Replace("S:\", "N:\")
                    ShippingLabelDirectory = ShippingLabelDirectory.Replace(ASCMAIN1.Folders("SharedRoot"), "N:\")
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

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", MyBase.Absx1.txtFor("WHSE_CODE").Text)
            If rowICTWHSE1 Is Nothing Then
                ErrorMessage = "Invalid or missing Warehouse"
                Return False
            End If

            txtCUST_ADDR1.Text = txtCUST_ADDR1.Text.Trim
            txtCUST_ADDR2.Text = txtCUST_ADDR2.Text.Trim
            txtCUST_CITY.Text = txtCUST_CITY.Text.Trim
            txtCUST_CONTACT.Text = txtCUST_CONTACT.Text.Trim
            txtCUST_COUNTRY.Text = txtCUST_COUNTRY.Text.Trim
            txtCUST_NAME.Text = txtCUST_NAME.Text.Trim
            txtCUST_STATE.Text = txtCUST_STATE.Text.Trim.ToUpper
            txtCUST_ZIP_CODE.Text = txtCUST_ZIP_CODE.Text.Trim

            If txtCUST_ADDR1.TextLength = 0 AndAlso txtCUST_ADDR2.Text.Length = 0 Then
                ErrorMessage = "Invalid or missing Ship To Street Address"
                Return False
            ElseIf txtCUST_CITY.TextLength = 0 OrElse txtCUST_STATE.TextLength = 0 OrElse txtCUST_ZIP_CODE.TextLength = 0 Then
                ErrorMessage = "Invalid or missing Ship To City, State, Zip Code"
                Return False
            ElseIf txtCUST_COUNTRY.TextLength = 0 Then
                Dim STATE_CODE As String = txtCUST_STATE.Text
                Dim rowTATSTATE As DataRow = tblTATSTATE.Rows.Find(STATE_CODE)
                If rowTATSTATE IsNot Nothing Then
                    txtCUST_COUNTRY.Text = "US"
                Else
                    ErrorMessage = "Invalid or missing Country Code"
                    Return False
                End If
            End If

            ' look at ship via settings 

            If rowSOTSVIA1 IsNot Nothing Then
                If (rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1") _
                        AndAlso chkInsureShipment.Checked Then
                    ErrorMessage = "Ship Vias classified as Collect or Third party do not permit Insurance"
                    Return False
                End If

                txt3PAccountNo.Text = txt3PAccountNo.Text.Trim
                txt3pCountry.Text = txt3pCountry.Text.Trim.ToUpper
                txt3PZipCode.Text = txt3PZipCode.Text.Trim.PadLeft(5, "0")
                If txt3pCountry.TextLength = 0 OrElse txt3pCountry.Text.StartsWith("US") Then
                    txt3pCountry.Text = "US"
                End If

                Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                    Case "F" ' Fedex
                        If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse optPayor.Value = "C" OrElse optPayor.Value = "R" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                'ErrorMessage = "Fedex Collect type Ship Vias require an Account Code, Zip Code and Country Code in the customer master."
                                'Return False
                            End If

                        ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse optPayor.Value = "P" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "Fedex Third Party type Ship Vias require an Account Code, Zip Code and Country Code on Header tab."
                                Return False
                            End If
                        End If

                    Case "U" ' Ups
                        If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso rowSOTSVIA1.Item("REQUIRES_ACCT_NO") & String.Empty <> "1" Then
                            'clsShip.Payor = TPayorTypes.ptConsignee
                        ElseIf rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" OrElse optPayor.Value = "C" OrElse optPayor.Value = "R" Then
                            ' Use the Account Information on Customer Master ARTCUST1
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Collect type Ship Vias require an Account Code, Zip Code and Country Code in the customer master."
                                Return False
                            End If
                        ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" OrElse optPayor.Value = "P" Then
                            If txt3PAccountNo.TextLength = 0 _
                                OrElse txt3pCountry.TextLength = 0 _
                                OrElse txt3PZipCode.TextLength = 0 Then
                                ErrorMessage = "UPS Third Party type Ship Vias require an Account Code, Zip Code and Country Code on Header tab."
                                Return False
                            End If
                        End If
                End Select
            End If

            ' If Fedex and Collect must be a Ground delivery
            If rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty = "F" _
                AndAlso CARRIER_PROD_CODE <> "15" _
                AndAlso (rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso optPayor.Value <> "R") Then
                ErrorMessage = "Fedex Collect shipments must ship ground. Choose Recipient payor type on the 'Header Info' tab for non ground shipments."
                tabSOTPICK1.SelectedTab = tabSOTPICK1.Tabs("Header Info")
                Return False
            End If

            If PreScreenForErrorsOnly Then Return True

            '*******************************************************************************

            Dim isInternationalShipment As Boolean = False
            Dim fedexSmartPost As Int16 = 26

            Dim PICK_NO As String = String.Empty
            Dim ORDR_NO As String = String.Empty
            Dim ORDR_NO_WEB As String = String.Empty
            Dim ORDR_CUST_PO As String = String.Empty
            Dim SHIP_BOL_NO As String = Absx1.txtFor("SHIP_BOL_NO").Text

            Dim FRT_TERMS As String = Absx1.txtFor("FRT_TERMS").Text
            Dim PPA_FREIGHT As Decimal = 0
            Dim OUR_FREIGHT As Decimal = 0

            dst.Tables("WHTSHIP1").Rows.Clear()
            dst.Tables("WHTSHIP2").Rows.Clear()
            dst.Tables("WHTSHIP5").Rows.Clear()
            dst.Tables("WHTSHIPS").Rows.Clear()
            dst.Tables("WHTSHIPC").Rows.Clear()
            dst.Tables("WHTSHIPP").Rows.Clear()

            Dim SHIP_CNTL_NO As String = String.Empty 'ASCMAIN1.Next_Control_No("WHTSHIP1.SHIP_CNTL_NO")
            Dim clsShip As New TAC.WHCSHIP1

            ' Credentials
            clsShip.Server = rowSOTCARR1.Item("CARRIER_REMOTE_HOST_IP") & String.Empty
            clsShip.UserId = rowSOTCARR3.Item("SHIPPER_ID") & String.Empty
            clsShip.Password = rowSOTCARR3.Item("SHIPPER_PASSWORD") & String.Empty
            clsShip.AccountNumber = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
            clsShip.UPSAccessKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.FedexMeterNumber = rowSOTCARR3.Item("METER_NUMBER") & String.Empty
            clsShip.FedexDeveloperKey = rowSOTCARR3.Item("ACCESSLICENSENUMBER") & String.Empty
            clsShip.LabelStockType = (rowSOTCARR1.Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

            Dim rowWHTSHIP1 As DataRow = Nothing
            Dim rowWHTSHIP2 As DataRow = Nothing
            Dim rowWHTSHIP5 As DataRow = Nothing

            If Not requestShippingOptions Then
                rowWHTSHIP1 = dst.Tables("WHTSHIP1").NewRow
                SHIP_CNTL_NO = ASCMAIN1.Next_Control_No("WHTSHIP1.SHIP_CNTL_NO")
                rowWHTSHIP1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                rowWHTSHIP1.Item("CARRIER_CODE") = CARRIER_CODE
                rowWHTSHIP1.Item("CARRIER_PROD_CODE") = CARRIER_PROD_CODE
                rowWHTSHIP1.Item("CARRIER_ACCOUNT_NO") = rowSOTCARR3.Item("CARRIER_ACCOUNT_NO") & String.Empty
                dst.Tables("WHTSHIP1").Rows.Add(rowWHTSHIP1)

                rowWHTSHIP1.Item("STATUS") = "I"
                rowWHTSHIP1.Item("ERROR_MSG") = String.Empty
                rowWHTSHIP1.Item("SHIP_DATE") = CDate(MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Value).ToString("MM/dd/yyyy")
                rowWHTSHIP1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                rowWHTSHIP1.Item("OPS_YYYYWW") = ASCMAIN1.CYW
                rowWHTSHIP1.Item("CUST_CODE") = CUST_CODE
                rowWHTSHIP1.Item("INIT_DATE") = DATETIME_STAMP
                rowWHTSHIP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowWHTSHIP1.Item("LAST_DATE") = DATETIME_STAMP
                rowWHTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowWHTSHIP1.Item("MASTER_TRACKING_NO") = String.Empty
                rowWHTSHIP1.Item("CUSTOMS_VALUE") = 0
                rowWHTSHIP1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
                rowWHTSHIP1.Item("SHIP_VIA_CODE") = SHIP_VIA_CODE

                rowWHTSHIP1.Item("INSURED_VALUE") = 0
                rowWHTSHIP1.Item("INSURED_SHIPMENT") = IIf(MyBase.Absx1.chkFor("INSURED_SHIPMENT").Checked, "1", "0")
            End If

            ' Sender Information
            With clsShip.Sender
                .Company = (rowICTWHSE1.Item("WHSE_DESC") & String.Empty).ToString.Trim
                .FirstName = (rowICTWHSE1.Item("WHSE_CONTACT") & String.Empty).ToString.Trim
                .MiddleInitial = String.Empty
                .LastName = String.Empty
                .Address1 = (rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty).ToString.Trim
                .Address2 = (rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty).ToString.Trim
                .City = (rowICTWHSE1.Item("WHSE_CITY") & String.Empty).ToString.Trim
                .State = (rowICTWHSE1.Item("WHSE_STATE") & String.Empty).ToString.Trim
                .ZipCode = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Trim
                .CountryCode = (rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"

                If (rowICTWHSE1.Item("WHSE_CODE") & String.Empty) = "16" Then
                    .Phone = "8003367254" '(rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
                Else
                    .Phone = (rowICTWHSE1.Item("WHSE_PHONE") & String.Empty).ToString.Trim
                End If

                If Not requestShippingOptions Then
                    rowWHTSHIP5 = dst.Tables("WHTSHIP5").NewRow
                    rowWHTSHIP5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHIP5.Item("SHIP_ADDR_TYPE") = "SF"
                    rowWHTSHIP5.Item("SHIP_FIRST_NAME") = .FirstName
                    rowWHTSHIP5.Item("SHIP_LAST_NAME") = .LastName
                    rowWHTSHIP5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                    rowWHTSHIP5.Item("SHIP_PHONE") = .Phone
                    rowWHTSHIP5.Item("SHIP_FAX") = .Fax
                    rowWHTSHIP5.Item("SHIP_EMAIL") = .eMail
                    rowWHTSHIP5.Item("SHIP_COMPANY") = .Company
                    rowWHTSHIP5.Item("SHIP_ADDR1") = .Address1
                    rowWHTSHIP5.Item("SHIP_ADDR2") = .Address2
                    rowWHTSHIP5.Item("SHIP_CITY") = .City
                    rowWHTSHIP5.Item("SHIP_STATE") = .State
                    rowWHTSHIP5.Item("SHIP_ZIP_CODE") = .ZipCode
                    rowWHTSHIP5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                    rowWHTSHIP5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                    rowWHTSHIP5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                    dst.Tables("WHTSHIP5").Rows.Add(rowWHTSHIP5)
                End If

            End With

            ' Recipient
            With clsShip.Recipient
                .FirstName = IIf(txtCUST_CONTACT.TextLength > 0, txtCUST_CONTACT.Text, txtCUST_NAME.Text)
                .MiddleInitial = ""
                .LastName = ""

                .Address1 = txtCUST_ADDR1.Text
                .Address2 = txtCUST_ADDR2.Text
                .City = txtCUST_CITY.Text
                .State = txtCUST_STATE.Text
                .ZipCode = txtCUST_ZIP_CODE.Text
                .CountryCode = txtCUST_COUNTRY.Text.ToUpper.Replace(".", "")
                If .CountryCode.Trim.Length = 0 Then .CountryCode = "US"
                If .CountryCode = "USA" Then .CountryCode = "US"

                .Company = txtCUST_NAME.Text
                .Phone = mdtCUST_PHONE.Text

                If .Phone.Trim.Length = 0 Then
                    .Phone = clsShip.Sender.Phone
                End If

                If .Phone.Trim.Length = 0 Then
                    .Phone = "1234567890"
                End If

                .IsResidental = optAddressType.Value = "R"
                .IsPOBox = optAddressType.Value = "P"

                If Not requestShippingOptions Then
                    rowWHTSHIP5 = dst.Tables("WHTSHIP5").NewRow
                    rowWHTSHIP5.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHIP5.Item("SHIP_ADDR_TYPE") = "ST"
                    rowWHTSHIP5.Item("SHIP_FIRST_NAME") = .FirstName
                    rowWHTSHIP5.Item("SHIP_LAST_NAME") = .LastName
                    rowWHTSHIP5.Item("SHIP_MIDDLE_INIT") = .MiddleInitial
                    rowWHTSHIP5.Item("SHIP_PHONE") = .Phone
                    rowWHTSHIP5.Item("SHIP_FAX") = .Fax
                    rowWHTSHIP5.Item("SHIP_EMAIL") = .eMail
                    rowWHTSHIP5.Item("SHIP_COMPANY") = .Company
                    rowWHTSHIP5.Item("SHIP_ADDR1") = .Address1
                    rowWHTSHIP5.Item("SHIP_ADDR2") = .Address2
                    rowWHTSHIP5.Item("SHIP_CITY") = .City
                    rowWHTSHIP5.Item("SHIP_STATE") = .State
                    rowWHTSHIP5.Item("SHIP_ZIP_CODE") = .ZipCode
                    rowWHTSHIP5.Item("SHIP_COUNTRY_CODE") = .CountryCode
                    rowWHTSHIP5.Item("SHIP_RESIDENTIAL") = IIf(.IsResidental, "1", "0")
                    rowWHTSHIP5.Item("SHIP_PO_BOX") = IIf(.IsPOBox, "1", "0")
                    dst.Tables("WHTSHIP5").Rows.Add(rowWHTSHIP5)
                End If

                isInternationalShipment = (.CountryCode <> "US") OrElse (.CountryCode = "US" AndAlso .State = "PR")
            End With

            Select Case PROVIDER_TYPE
                Case WHCSHIP1.ProviderTypeFedex
                    If Not isInternationalShipment Then
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress
                    Else
                        clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpressInternational
                    End If
                Case WHCSHIP1.ProviderTypeUPS
                    clsShip.Service = WHCSHIP1.ServiceProviders.UPS
                Case WHCSHIP1.ProviderTypeUSPS
                    clsShip.Service = WHCSHIP1.ServiceProviders.USPS
                Case WHCSHIP1.ProviderTypeCanada
                    clsShip.Service = WHCSHIP1.ServiceProviders.CanadaPost
                Case Else
                    Return False
            End Select

            ' Build a package for each Carton for the current Pick Ticket
            ' Change as of 1/21/2013
            ' Some shipments are multi Pick Tickets and some Pick Tickets are combined into 1 carton.
            ' The carton sequence will be used to group pick tickets into one carton and also
            ' be used to identify the sequence the Shipping label will get printed
            ' The user is not permitted to deselect a pick ticket; therefore, no londfer need to use dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")
            clsShip.PackageDetailList.Clear()

            'For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")

            Dim cartSequenceNos As List(Of Int16) = New List(Of Int16)

            ' Commodities for international shipments
            clsShip.TotalCustomsValue = 0
            clsShip.CommodityDetailList.Clear()
            Dim COMMODITY_LNO As Int16 = 1
            Dim itemList As List(Of String) = New List(Of String)

            For Each rowSOTPICK1 In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "PICK_NO")

                PICK_NO = rowSOTPICK1.Item("PICK_NO") & String.Empty
                ORDR_NO = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                SHIP_BOL_NO = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                ORDR_NO_WEB = rowSOTPICK1.Item("ORDR_NO_WEB") & String.Empty
                ORDR_CUST_PO = rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                PPA_FREIGHT = 0
                OUR_FREIGHT = 0

                ' Get the Invoice Number now so we can put it on the label
                Dim INV_NO As String = String.Empty

                If Not requestShippingOptions Then
                    INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
                    rowSOTPICK1.Item("INV_NO") = INV_NO
                    rowSOTPICK1.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                End If

                ' See if we have cartons setup
                If dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'").Length = 0 Then
                    Continue For
                End If

                ' See if the carton has products
                If dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "' AND ISNULL(QTY_PACKED, 0) > 0 ").Length = 0 Then
                    Continue For
                End If

                For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("PICK_NO = '" & PICK_NO & "'", "CART_SEQ")
                    ' This is done to place multi pick tickets into one carton
                    Dim CART_SEQ As Int16 = rowSOTCART1.Item("CART_SEQ")
                    If cartSequenceNos.Contains(CART_SEQ) Then
                        Continue For
                    End If
                    cartSequenceNos.Add(CART_SEQ)

                    Dim PACKAGING_TYPE As String = rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty
                    Dim PKG_CODE As String = rowSOTCART1.Item("PKG_CODE") & String.Empty
                    Dim rowWHTPKGM1 As DataRow = LookUp("WHTPKGM1", PKG_CODE)
                    pkgId = CART_SEQ ' (Val(StrReverse(StrReverse(rowSOTCART1.Item("CART_NO").ToString).Substring(0, 8))))

                    Dim shipPackageDetail As New DPayments.DShippingSDK.PackageDetail
                    With shipPackageDetail
                        .PackagingType = Val(PACKAGING_TYPE)

                        ' This is done to place multi pick tickets into one carton. Need combined weight 
                        .Weight = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "CART_SEQ = " & CART_SEQ) & String.Empty)
                        If .Weight = 0 Then
                            .Weight = 1
                        End If

                        '*************************************
                        '        Convert to Ounces
                        '*************************************
                        .Weight = Convert.ToInt16(.Weight * 16)

                        If rowWHTPKGM1 IsNot Nothing Then
                            If rowWHTPKGM1.Item("PKG_CODE") & String.Empty = "OTHER" Then
                                .Length = Val(rowSOTCART1.Item("LENGTH") & String.Empty)
                                .Width = Val(rowSOTCART1.Item("WIDTH") & String.Empty)
                                .Height = Val(rowSOTCART1.Item("HEIGHT") & String.Empty)
                            Else
                                .Length = Val(rowWHTPKGM1.Item("PKG_L") & String.Empty)
                                .Width = Val(rowWHTPKGM1.Item("PKG_W") & String.Empty)
                                .Height = Val(rowWHTPKGM1.Item("PKG_H") & String.Empty)
                            End If
                        End If

                        Dim reference As String = String.Empty
                        Dim refCount As Int16 = 0

                        Select Case PROVIDER_TYPE
                            Case WHCSHIP1.ProviderTypeFedex
                                ' Fedex allows up to 3 References
                                If (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (rowSOTCART1.Item("REFERENCE1") & String.Empty).ToString
                                    refCount += 1
                                End If

                                If (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString.Trim.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & (rowSOTCART1.Item("REFERENCE2") & String.Empty).ToString
                                    refCount += 1
                                End If

                                ' Web order no is the ORDR_CUST_PO as well
                                If ORDR_NO_WEB.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & ORDR_NO_WEB
                                    refCount += 1
                                ElseIf ORDR_CUST_PO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; CR:" & ORDR_CUST_PO
                                    refCount += 1
                                End If

                                If INV_NO.Length > 0 AndAlso refCount < 3 Then
                                    reference &= "; IN:" & INV_NO
                                    refCount += 1
                                End If

                                If (rowSOTSHIP1.Item("ORDR_DEPT") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 3 Then
                                    reference &= "; DN:" & rowSOTSHIP1.Item("ORDR_DEPT") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                            Case WHCSHIP1.ProviderTypeUPS
                                ' Ups allows up to 2 References
                                If (rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                    reference &= "; ST:" & rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty
                                    refCount += 1
                                End If

                                If (rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty).ToString.Trim <> String.Empty AndAlso refCount < 2 Then
                                    reference &= "; PO:" & rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
                                    refCount += 1
                                End If

                                If reference.Length > 0 Then
                                    reference = reference.Substring(1).Trim
                                End If

                        End Select

                        .Reference = reference
                        .Id = pkgId.ToString("D8")

                        If chkInsureShipment.Checked Then
                            .InsuredValue = Val(rowSOTCART1.Item("INSURANCE") & String.Empty)
                            rowWHTSHIP1.Item("INSURED_VALUE") += Val(rowSOTCART1.Item("INSURANCE") & String.Empty)
                        End If

                    End With
                    clsShip.PackageDetailList.Add(shipPackageDetail)

                    If Not requestShippingOptions Then
                        rowWHTSHIP2 = dst.Tables("WHTSHIP2").NewRow
                        rowWHTSHIP2.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                        rowWHTSHIP2.Item("SHIP_PACKAGE_NO") = pkgId
                        rowWHTSHIP2.Item("HEIGHT") = shipPackageDetail.Height
                        rowWHTSHIP2.Item("INSURED_VALUE") = 0
                        rowWHTSHIP2.Item("LENGTH") = shipPackageDetail.Length
                        rowWHTSHIP2.Item("NET_CHARGE") = 0
                        rowWHTSHIP2.Item("PACKAGING_TYPE") = Val(shipPackageDetail.PackagingType)
                        rowWHTSHIP2.Item("TOTAL_DISCOUNT") = 0
                        rowWHTSHIP2.Item("TOTAL_SURCHARGES") = 0
                        rowWHTSHIP2.Item("TRACKING_NUMBER") = String.Empty
                        rowWHTSHIP2.Item("WEIGHT") = Convert.ToInt16(shipPackageDetail.Weight)
                        rowWHTSHIP2.Item("WIDTH") = shipPackageDetail.Width
                        rowWHTSHIP2.Item("TRACKING_NO") = String.Empty

                        rowWHTSHIP2.Item("CUST_REF") = ORDR_CUST_PO
                        rowWHTSHIP2.Item("INV_BOL_NO") = SHIP_BOL_NO
                        rowWHTSHIP2.Item("CART_NO") = rowSOTCART1.Item("CART_NO") & String.Empty
                        rowWHTSHIP2.Item("INV_NO") = INV_NO
                        rowWHTSHIP2.Item("PO_ORDER_NO") = String.Empty
                        rowWHTSHIP2.Item("DEPT_NO") = (rowSOTPICK1.Item("ORDR_DEPT") & String.Empty).ToString.Trim

                        dst.Tables("WHTSHIP2").Rows.Add(rowWHTSHIP2)
                    End If

                Next


                If isInternationalShipment Then
                    ' Set the Customs value
                    clsShip.TotalCustomsValue = Val(rowSOTPICK1.Item("PICK_AMT_CONF") & String.Empty)

                    For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select("PICK_NO = '" & PICK_NO & "'")
                        Dim ITEM_CODE As String = rowSOTCART2.Item("ITEM_CODE")
                        If itemList.Contains(ITEM_CODE) Then Continue For

                        itemList.Add(ITEM_CODE)
                        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                        ' Just in case a non item is permitted in the shipment
                        If rowICTITEM1 Is Nothing Then Continue For

                        Dim CommodityDetail As New DPayments.DShippingSDK.CommodityDetail
                        CommodityDetail.Description = rowICTITEM1.Item("ITEM_DESC") & String.Empty

                        Dim NumberOfPieces As Int16 = Val(dst.Tables("SOTCART2").Compute("SUM(QTY_PACKED)", "ITEM_CODE = '" & ITEM_CODE & "' and PICK_NO = '" & PICK_NO & "'") & String.Empty)

                        CommodityDetail.NumberOfPieces = NumberOfPieces
                        CommodityDetail.Quantity = NumberOfPieces
                        CommodityDetail.QuantityUnit = "EA"

                        Dim pickUnitPrice As Decimal = Val(dst.Tables("SOTPICK2").Compute("MAX(PICK_UNIT_PRICE)", "PICK_NO = '" & PICK_NO & "' AND ITEM_CODE = '" & ITEM_CODE & "'") & String.Empty)
                        CommodityDetail.UnitPrice = pickUnitPrice

                        CommodityDetail.Weight = Val(rowICTITEM1.Item("ITEM_WEIGHT") & String.Empty) ' Leave as pounds
                        CommodityDetail.Manufacturer = (rowICTITEM1.Item("COUNTRY_CODE") & String.Empty).ToString.ToUpper.Trim ' "US" '
                        If CommodityDetail.Manufacturer.Length = 0 Then
                            CommodityDetail.Manufacturer = "US"
                        End If
                        clsShip.CommodityDetailList.Add(CommodityDetail)

                        If Not requestShippingOptions Then
                            Dim rowWHTSHIPC As DataRow = dst.Tables("WHTSHIPC").NewRow
                            rowWHTSHIPC.Item("SHIP_CNTL_NO") = SHIP_CNTL_NO
                            rowWHTSHIPC.Item("COMMODITY_LNO") = COMMODITY_LNO
                            COMMODITY_LNO += 1
                            rowWHTSHIPC.Item("COMMODITY_DESC") = CommodityDetail.Description
                            rowWHTSHIPC.Item("NUM_PIECES") = CommodityDetail.NumberOfPieces
                            rowWHTSHIPC.Item("MANUFACTURER") = CommodityDetail.Manufacturer
                            rowWHTSHIPC.Item("HARMONIZED_CODE") = String.Empty
                            rowWHTSHIPC.Item("WEIGHT") = CommodityDetail.Weight
                            rowWHTSHIPC.Item("QUANTITY") = CommodityDetail.Quantity
                            rowWHTSHIPC.Item("QUANTITY_UOM") = CommodityDetail.QuantityUnit
                            rowWHTSHIPC.Item("UNIT_PRICE") = CommodityDetail.UnitPrice
                            dst.Tables("WHTSHIPC").Rows.Add(rowWHTSHIPC)
                        End If

                    Next
                End If
            Next  ' This is where the For Sotpick1, for sotcart1, for sotcart2 should end 

            If Val(numInsureValue.Value & String.Empty) > clsShip.TotalCustomsValue Then
                clsShip.TotalCustomsValue = numInsureValue.Value
            End If

            If chkSignature.Checked Then
                clsShip.SignatureRequired = True
            Else
                clsShip.SignatureRequired = False
            End If

            ' Shipping Method
            If isInternationalShipment Then
                clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE_INTL") & String.Empty)
            Else
                clsShip.RequestedServiceType = Val(rowSOTSVIA1.Item("CARRIER_PROD_CODE") & String.Empty)
            End If

            If clsShip.RequestedServiceType = fedexSmartPost Then
                clsShip.FedexSmartPost.HubId = rowSOTCARR3.Item("FEDEX_HUB_ID") & String.Empty
            End If

            'clsShip.DropOffType = FedexshipintlDropoffTypes.dtRegularPickup

            ' The COLLECT payment type is only supported in FedEx Ground services. The CONSIGNEE type is only supported in UPS service.

            ' For FedEx, when this field is set to a value other than 0 (ptSender), the AccountNumber and 
            ' CountryCode are required to be provided in the request as well. Otherwise, those will default to AccountNumber and CountryCode.

            ' For UPS, when set to ptSender, the AccountNumber is automatically set to AccountNumber. 
            ' When ptRecipient is specified, AccountNumber and ZipCode are required to be provided in the request. 
            ' For return international shipments, this option is invalid for transportation charges. 
            ' And, when ptThirdParty has been specified, the AccountNumber, ZipCode and CountryCode are 
            ' required to be provided in the request. When ptConsignee is specified, it indicates that UPS Consignee Billing 
            ' option is selected, no other fields need to be set. ptConsignee only applies to US/PR and PR/US shipment origins and destination. 

            ' Payor of the Shipmenet
            clsShip.Payor = TPayorTypes.ptSender


            Select Case rowSOTCARR1.Item("PROVIDER_TYPE") & String.Empty
                Case "F" ' Fedex
                    If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso optPayor.Value <> "R" Then
                        ' Use the Account Information on Customer Master ARTCUST1
                        clsShip.Payor = TPayorTypes.ptCollect
                        If (rowARTCUST1.Item("FDX_ACCT_NO") & String.Empty).ToString.Trim.Length > 0 Then
                            clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                            clsShip.PayorContact.CountryCode = txt3pCountry.Text
                            clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                            If clsShip.PayorContact.CountryCode = String.Empty Then
                                clsShip.PayorContact.CountryCode = "US"
                            End If
                        End If
                    ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "C" Then
                        clsShip.Payor = TPayorTypes.ptCollect
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "R" Then
                        clsShip.Payor = TPayorTypes.ptRecipient
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value = "P" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    End If

                Case "U" ' Ups
                    If rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" AndAlso rowSOTSVIA1.Item("REQUIRES_ACCT_NO") & String.Empty <> "1" Then
                        clsShip.Payor = TPayorTypes.ptConsignee
                    ElseIf rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1" Then
                        ' Use the Account Information on Customer Master ARTCUST1
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    ElseIf optPayor.Value <> "O" Then
                        clsShip.Payor = TPayorTypes.ptThirdParty
                        clsShip.PayorContact.AccountNumber = txt3PAccountNo.Text
                        clsShip.PayorContact.CountryCode = txt3pCountry.Text
                        clsShip.PayorContact.ZipCode = txt3PZipCode.Text
                        If clsShip.PayorContact.CountryCode = String.Empty Then
                            clsShip.PayorContact.CountryCode = "US"
                        End If
                    End If

            End Select

            Dim rowWHTSHIPP As DataRow

            If Not requestShippingOptions Then
                If clsShip.Payor <> TPayorTypes.ptSender Then
                    Dim rowWHTSHIP3 As DataRow = dst.Tables("WHTSHIP3").NewRow
                    rowWHTSHIP3("SHIP_CNTL_NO") = SHIP_CNTL_NO
                    rowWHTSHIP3("SHIP_BOL_NO") = SHIP_BOL_NO
                    rowWHTSHIP3("ACCOUNT_NO_3PL") = txt3PAccountNo.Text
                    rowWHTSHIP3("ZIP_CODE_3PL") = txt3PZipCode.Text
                    rowWHTSHIP3("COUNTRY_3PL") = txt3pCountry.Text
                    dst.Tables("WHTSHIP3").Rows.Add(rowWHTSHIP3)
                End If

                rowWHTSHIPP = dst.Tables("WHTSHIPP").NewRow
                rowWHTSHIPP("SHIP_CNTL_NO") = SHIP_CNTL_NO
                rowWHTSHIPP("PAYOR_TYPE") = "S"
                rowWHTSHIPP("PAYOR_ACCT_NO") = clsShip.PayorContact.AccountNumber & String.Empty
                rowWHTSHIPP("PAYOR_COUNTRY") = clsShip.PayorContact.CountryCode & String.Empty
                dst.Tables("WHTSHIPP").Rows.Add(rowWHTSHIPP)
            End If


            ' Payor of the Duties

            clsShip.DutiesPayor = TPayorTypes.ptSender
            If isInternationalShipment Then
                clsShip.DutiesPayor = clsShip.Payor
                clsShip.DutiesPayorContact.AccountNumber = clsShip.PayorContact.AccountNumber
                clsShip.DutiesPayorContact.CountryCode = clsShip.PayorContact.CountryCode
                clsShip.DutiesPayorContact.ZipCode = clsShip.PayorContact.ZipCode
            End If

            If Not requestShippingOptions Then
                rowWHTSHIPP = dst.Tables("WHTSHIPP").NewRow
                rowWHTSHIPP("SHIP_CNTL_NO") = SHIP_CNTL_NO
                rowWHTSHIPP("PAYOR_TYPE") = "D"
                rowWHTSHIPP("PAYOR_ACCT_NO") = clsShip.DutiesPayorContact.AccountNumber & String.Empty
                rowWHTSHIPP("PAYOR_COUNTRY") = clsShip.DutiesPayorContact.CountryCode & String.Empty
                dst.Tables("WHTSHIPP").Rows.Add(rowWHTSHIPP)
            End If

            With clsShip
                .EzshipLabelImage = DPayments.DShippingSDK.EzShipLabelImageTypes.itEltron
                .ShippingLabelDirectory = ShippingLabelDirectory
                .ShippingLabelPrefix = SHIP_CNTL_NO
                .ShipDate = dteSHIP_DATE_SHIPPED.DateTime.ToString("yyyy-MM-dd")
                If chkSaturday.Checked Then
                    clsShip.ShipmentSpecialServices = clsShip.ShipmentSpecialServices Or &H10000000L
                End If
            End With

            If Not requestShippingOptions Then
                Try
                    BeginTrans()
                    Update_Record_TDA("WHTSHIP1")
                    Update_Record_TDA("WHTSHIP2")
                    Update_Record_TDA("WHTSHIP3")
                    Update_Record_TDA("WHTSHIP5")
                    Update_Record_TDA("WHTSHIPS")
                    Update_Record_TDA("WHTSHIPP")
                    Update_Record_TDA("WHTSHIPC")
                    CommitTrans()
                Catch ex As Exception
                    ErrorMessage &= " " & ex.Message
                    Rollback()
                End Try
            End If


            If requestShippingOptions Then
                dst.Tables("SOTCARRX").Rows.Clear()

                ' ******************************** FEDEX ********************************
                Dim tblSOTCARR1 As DataTable = ASCDATA1.GetDataTable("select * from sotcarr1 where PROVIDER_TYPE = 'F'")
                Dim tblSOTCARR2 As DataTable = ASCDATA1.GetDataTable("Select * from sotcarr2 where carrier_code in (select carrier_code from sotcarr1 where PROVIDER_TYPE = 'F')")
                Dim tblSOTCARR3 As DataTable = ASCDATA1.GetDataTable("Select * from sotcarr3 where carrier_code in (select carrier_code from sotcarr1 where PROVIDER_TYPE = 'F')")
                Dim tblSOTSVIA1 As DataTable = ASCDATA1.GetDataTable("Select * from sotsvia1")

                clsShip.RequestedServiceType = ServiceTypes.stUnspecified

                ' Credentials
                clsShip.Server = tblSOTCARR1.Rows(0).Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                clsShip.UserId = tblSOTCARR3.Rows(0).Item("SHIPPER_ID") & String.Empty
                clsShip.Password = tblSOTCARR3.Rows(0).Item("SHIPPER_PASSWORD") & String.Empty
                clsShip.AccountNumber = tblSOTCARR3.Rows(0).Item("CARRIER_ACCOUNT_NO") & String.Empty
                clsShip.UPSAccessKey = tblSOTCARR3.Rows(0).Item("ACCESSLICENSENUMBER") & String.Empty
                clsShip.FedexMeterNumber = tblSOTCARR3.Rows(0).Item("METER_NUMBER") & String.Empty
                clsShip.FedexDeveloperKey = tblSOTCARR3.Rows(0).Item("ACCESSLICENSENUMBER") & String.Empty
                clsShip.LabelStockType = (tblSOTCARR1.Rows(0).Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

                'If isInternationalShipment Then
                '    clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpressInternational
                'Else
                clsShip.Service = WHCSHIP1.ServiceProviders.FederalExpress
                'End If

                clsShip.GetFedexRates()

                For Each reqSer As ServiceDetail In clsShip.RequestedServicesRates
                    Dim rowSOTCARRX As DataRow = dst.Tables("SOTCARRX").NewRow
                    rowSOTCARRX.Item("CARRIER_CODE") = tblSOTCARR1.Rows(0).Item("CARRIER_CODE") & String.Empty
                    rowSOTCARRX.Item("CARRIER_DESC") = tblSOTCARR1.Rows(0).Item("CARRIER_DESC") & String.Empty
                    rowSOTCARRX.Item("CARRIER_PROD_CODE") = Convert.ToInt16(reqSer.ServiceType)
                    rowSOTCARRX.Item("CARRIER_PROD_DESC") = StrConv(reqSer.ServiceTypeDescription.Replace("_", " "), VbStrConv.ProperCase)
                    rowSOTCARRX.Item("TRANSIT_TIME") = StrConv(reqSer.TransitTime.Replace("_", " "), VbStrConv.ProperCase)
                    rowSOTCARRX.Item("LIST_PRICE") = reqSer.ListNetCharge
                    rowSOTCARRX.Item("OUR_PRICE") = reqSer.AccountNetCharge

                    Dim SHIP_VIA_CODEs As String = String.Empty

                    For Each rowSOTSVIA1 In tblSOTSVIA1.Select("CARRIER_CODE = '" & rowSOTCARRX.Item("CARRIER_CODE") & "' AND CARRIER_PROD_CODE = '" & rowSOTCARRX.Item("CARRIER_PROD_CODE") & "' AND SHIP_VIA_STATUS = 'A'")
                        SHIP_VIA_CODEs &= ", " & rowSOTSVIA1.Item("SHIP_VIA_CODE")
                    Next
                    If SHIP_VIA_CODEs.Length > 0 Then SHIP_VIA_CODEs = SHIP_VIA_CODEs.Substring(1).Trim '

                    rowSOTCARRX.Item("SHIP_VIA_CODE") = SHIP_VIA_CODEs

                    dst.Tables("SOTCARRX").Rows.Add(rowSOTCARRX)
                Next

                ' ******************************** UPS ********************************
                tblSOTCARR1 = ASCDATA1.GetDataTable("select * from sotcarr1 where PROVIDER_TYPE = 'U'")
                tblSOTCARR2 = ASCDATA1.GetDataTable("Select * from sotcarr2 where carrier_code in (select carrier_code from sotcarr1 where PROVIDER_TYPE = 'U')")
                tblSOTCARR3 = ASCDATA1.GetDataTable("Select * from sotcarr3 where carrier_code in (select carrier_code from sotcarr1 where PROVIDER_TYPE = 'U')")

                ' Credentials
                clsShip.Server = tblSOTCARR1.Rows(0).Item("CARRIER_REMOTE_HOST_IP") & String.Empty
                If Not clsShip.Server.EndsWith("/") Then
                    clsShip.Server &= "/Rate"
                Else
                    clsShip.Server &= "Rate"
                End If

                clsShip.UserId = tblSOTCARR3.Rows(0).Item("SHIPPER_ID") & String.Empty
                clsShip.Password = tblSOTCARR3.Rows(0).Item("SHIPPER_PASSWORD") & String.Empty
                clsShip.AccountNumber = tblSOTCARR3.Rows(0).Item("CARRIER_ACCOUNT_NO") & String.Empty
                clsShip.UPSAccessKey = tblSOTCARR3.Rows(0).Item("ACCESSLICENSENUMBER") & String.Empty
                clsShip.FedexMeterNumber = tblSOTCARR3.Rows(0).Item("METER_NUMBER") & String.Empty
                clsShip.FedexDeveloperKey = tblSOTCARR3.Rows(0).Item("ACCESSLICENSENUMBER") & String.Empty
                clsShip.LabelStockType = (tblSOTCARR1.Rows(0).Item("LABEL_STOCK_TYPE") & String.Empty).ToString.Trim

                clsShip.RequestedServiceType = ServiceTypes.stUnspecified
                'If isInternationalShipment Then
                '    clsShip.Service = WHCSHIP1.ServiceProviders.UPSInternational
                'Else
                clsShip.Service = WHCSHIP1.ServiceProviders.UPS
                'End If
                clsShip.GetUPSRates()

                For Each reqSer As ServiceDetail In clsShip.RequestedServicesRates
                    Dim rowSOTCARRX As DataRow = dst.Tables("SOTCARRX").NewRow
                    rowSOTCARRX.Item("CARRIER_CODE") = tblSOTCARR1.Rows(0).Item("CARRIER_CODE") & String.Empty
                    rowSOTCARRX.Item("CARRIER_DESC") = tblSOTCARR1.Rows(0).Item("CARRIER_DESC") & String.Empty
                    rowSOTCARRX.Item("CARRIER_PROD_CODE") = Convert.ToInt16(reqSer.ServiceType)
                    rowSOTCARRX.Item("CARRIER_PROD_DESC") = StrConv(reqSer.ServiceTypeDescription.Replace("_", " "), VbStrConv.ProperCase)
                    rowSOTCARRX.Item("TRANSIT_TIME") = StrConv(reqSer.TransitTime.Replace("_", " "), VbStrConv.ProperCase)
                    rowSOTCARRX.Item("LIST_PRICE") = reqSer.ListNetCharge
                    rowSOTCARRX.Item("OUR_PRICE") = reqSer.AccountNetCharge

                    Dim SHIP_VIA_CODEs As String = String.Empty

                    For Each rowSOTSVIA1 In tblSOTSVIA1.Select("CARRIER_CODE = '" & rowSOTCARRX.Item("CARRIER_CODE") & "' AND CARRIER_PROD_CODE = '" & rowSOTCARRX.Item("CARRIER_PROD_CODE") & "' AND SHIP_VIA_STATUS = 'A'")
                        SHIP_VIA_CODEs &= ", " & rowSOTSVIA1.Item("SHIP_VIA_CODE")
                    Next
                    If SHIP_VIA_CODEs.Length > 0 Then SHIP_VIA_CODEs = SHIP_VIA_CODEs.Substring(1).Trim '

                    rowSOTCARRX.Item("SHIP_VIA_CODE") = SHIP_VIA_CODEs

                    dst.Tables("SOTCARRX").Rows.Add(rowSOTCARRX)
                Next

                Exit Function
            End If

            If clsShip.RequestLabel() Then

                rowWHTSHIP1.Item("ERROR_MSG") = clsShip.LastError & String.Empty
                rowWHTSHIP1.Item("STATUS") = "P"
                If rowWHTSHIP1 IsNot Nothing AndAlso (rowWHTSHIP1.Item("ERROR_MSG") & String.Empty).ToString.Length > 200 Then
                    rowWHTSHIP1.Item("ERROR_MSG") = rowWHTSHIP1("ERROR_MSG").ToString.Substring(0, 200).Trim
                End If
                rowWHTSHIP1.Item("MASTER_TRACKING_NO") = clsShip.MasterTrackingNumber & String.Empty

                For Each shipPackageDetail As DPayments.DShippingSDK.PackageDetail In clsShip.PackageDetailList
                    SHIP_PACKAGE_NO = Val(shipPackageDetail.Id)
                    If dst.Tables("WHTSHIP2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO, "").Length > 0 Then
                        rowWHTSHIP2 = dst.Tables("WHTSHIP2").Select("SHIP_PACKAGE_NO = " & SHIP_PACKAGE_NO)(0)
                        rowWHTSHIP2.Item("TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                        rowWHTSHIP2.Item("BASE_CHARGE") = Val(clsShip.ShipmentBaseCharge(SHIP_PACKAGE_NO) & String.Empty)
                        rowWHTSHIP2.Item("NET_CHARGE") = Val(clsShip.ShipmentNetCharge(SHIP_PACKAGE_NO) & String.Empty)
                        rowWHTSHIP2.Item("TOTAL_DISCOUNT") = Val(clsShip.ShipmentDiscountCharge(SHIP_PACKAGE_NO) & String.Empty)
                        rowWHTSHIP2.Item("TOTAL_SURCHARGES") = Val(clsShip.ShipmentSurCharge(SHIP_PACKAGE_NO) & String.Empty)

                        If clsShip.ShipmentListCharge.ContainsKey(SHIP_PACKAGE_NO) Then
                            rowWHTSHIP2.Item("LIST_PRICE") = Val(clsShip.ShipmentListCharge(SHIP_PACKAGE_NO) & String.Empty)
                        Else
                            rowWHTSHIP2.Item("LIST_PRICE") = rowWHTSHIP2.Item("NET_CHARGE")
                        End If
                        PPA_FREIGHT = Val(rowWHTSHIP2.Item("LIST_PRICE") & String.Empty)
                        OUR_FREIGHT = Val(rowWHTSHIP2.Item("NET_CHARGE") & String.Empty)

                        PICK_NO = String.Empty
                        rowSOTPICK1 = Nothing

                        ' We may have multi pick tickets in a single carton. This stamps them with the same tracking number
                        ' Spread the Customer Freight Cost and Our freight cost across the Pick Tickets
                        Dim numPickTickets As Int16 = dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO).Length
                        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("CART_SEQ = " & SHIP_PACKAGE_NO)
                            rowSOTCART1.Item("CART_TRACKING_NO") = shipPackageDetail.TrackingNumber & String.Empty
                            PICK_NO = rowSOTCART1.Item("PICK_NO") & String.Empty
                            rowSOTPICK1 = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                            If Absx1.txtFor("FRT_TERMS").Text = "PPA" AndAlso rowSOTPICK1("ORDR_SOURCE") & String.Empty <> "W" Then
                                rowSOTPICK1.Item("PICK_FREIGHT") = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Math.Round(PPA_FREIGHT / numPickTickets, 2)
                            End If
                            rowSOTPICK1.Item("OUR_FREIGHT") = Val(rowSOTPICK1.Item("OUR_FREIGHT") & String.Empty) + Math.Round(OUR_FREIGHT / numPickTickets, 2)
                        Next
                    End If

                    ShippingLabels.Add(shipPackageDetail.ShippingLabel)
                    ShippingLabels.Add(shipPackageDetail.CODLabel)
                    ShippingLabels.Add(shipPackageDetail.ReturnReceipt)
                Next

                Try
                    BeginTrans()
                    Update_Record_TDA("WHTSHIP1")
                    Update_Record_TDA("WHTSHIP2")
                    If RecreateLabel Then
                        Update_Record_TDA("SOTCART1")
                    End If
                    CommitTrans()
                Catch ex As Exception
                    ErrorMessage &= " " & ex.Message
                    Rollback()
                End Try

            Else
                ErrorMessage &= " " & clsShip.LastError
                RequestShippingLabel = False
            End If

        Catch ex As Exception
            ErrorMessage &= " " & ex.Message
            RequestShippingLabel = False
        End Try

        ErrorMessage = ErrorMessage.Trim

    End Function

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_VIA_CODE"
                Dim SHIP_VIA_CODE As String = Absx1.txtFor("SHIP_VIA_CODE").Text.Trim

                ASCMAIN1.Add_Value_List(grdSOTCART1, "PACKAGING_TYPE", "SELECT SOTCARR4.PACKAGE_CODE, SOTCARR4.PACKAGE_DESC" _
                                        & " FROM SOTSVIA1, SOTCARR4" _
                                        & " WHERE SOTCARR4.CARRIER_CODE = SOTSVIA1.CARRIER_CODE" _
                                        & " AND SOTSVIA1.SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'" _
                                        & " ORDER BY PACKAGE_CODE DESC")


                Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                commonCarrier = rowSOTSVIA1 IsNot Nothing AndAlso rowSOTSVIA1.Item("CARRIER_TYPE") & String.Empty = "U"

                grdSOTCART1.DisplayLayout.Bands(0).Columns("PACKAGING_TYPE").Hidden = Not commonCarrier
                grdSOTCART1.DisplayLayout.Bands(0).Columns("PKG_CODE").Hidden = Not commonCarrier
        End Select

    End Sub

#Region "Serial and Com Connections"

    ' Handles Keyboard wedge
    Private receivingWedgeScan As Boolean = False
    Private strWedgeScan As String = String.Empty
    Private registeredWeight As Decimal = 0

    ''' <summary>
    ''' Form activate - Calls to setup devices
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    ''' <summary>
    ''' Sets the Printer Settings
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetUpPortsAndPrinters()

        Dim tooltip As New System.Windows.Forms.ToolTip()

        '**************************
        '**    Laser Printer
        '**************************
        Try
            txtLaserPrinter.Text = ASCMAIN1.LaserPrinterIpAddress
            tooltip.SetToolTip(txtLaserPrinter, ASCMAIN1.LaserPrinterIpAddress)
            If ASCMAIN1.LaserPrinterIpAddress.Length = 0 Then
                txtLaserPrinter.Appearance.BackColor = Drawing.Color.Red
            Else
                txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow
                If Net.IPAddress.TryParse(ASCMAIN1.LaserPrinterIpAddress, Nothing) Then
                    txtLaserPrinter.Appearance.BackColor = Drawing.Color.Green
                End If
            End If

        Catch ex As Exception
            txtLaserPrinter.Appearance.BackColor = Drawing.Color.Yellow
            tooltip.SetToolTip(txtLaserPrinter, ex.Message)
        End Try


        '**************************
        '**    Label Printer Port
        '**************************        
        Try
            txtLabelPrinter.BackColor = Drawing.Color.Red

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtLabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            Else
                Me.txtLabelPrinter.Text = "No Port"
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            End If

            txtLabelPrinter.BackColor = Drawing.Color.Yellow
            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                If Not ASCMAIN1.Running_in_VS Then ASCMAIN1.LabelPrinterSerialPort.Open()
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                txtLabelPrinter.BackColor = Drawing.Color.Green
            End If

        Catch ex As Exception
            txtLabelPrinter.BackColor = Drawing.Color.Red
            tooltip.SetToolTip(txtLabelPrinter, ex.Message)
        End Try

        '**************************
        '**    Scale Port
        '************************** 
        'ASCMAIN1.ScaleWeightDelegate = AddressOf ProcessScaleData
        'Try
        '    txtScale.BackColor = Drawing.Color.Red

        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing Then
        '        txtScale.Text = ASCMAIN1.ScaleSerialPort.PortName
        '        tooltip.SetToolTip(txtScale, txtScale.Text)
        '    Else
        '        txtScale.Text = "No Port"
        '        tooltip.SetToolTip(txtScale, txtScale.Text)
        '    End If

        '    txtScale.BackColor = Drawing.Color.Yellow
        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing AndAlso Not ASCMAIN1.ScaleSerialPort.IsOpen Then
        '        If Not ASCMAIN1.Running_in_VS Then ASCMAIN1.ScaleSerialPort.Open()
        '    End If

        '    If ASCMAIN1.ScaleSerialPort IsNot Nothing AndAlso ASCMAIN1.ScaleSerialPort.IsOpen Then
        '        txtScale.BackColor = Drawing.Color.Green
        '    End If

        'Catch ex As Exception
        '    txtScale.BackColor = Drawing.Color.Red
        '    tooltip.SetToolTip(txtScale, ex.Message)
        'End Try

    End Sub

    ''' <summary>
    ''' Sends the Scanned Bar Code to the Appropriate Control
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ProcessScannedData(ByVal scannedData As String)

        If MdiParent.ActiveMdiChild Is Nothing Then Exit Sub
        If MdiParent.ActiveMdiChild.Name <> Me.Name Then Exit Sub

        Static dataReceived As String

        dataReceived += scannedData
        If InStr(dataReceived, Chr(13), CompareMethod.Text) = 0 Then
            Exit Sub
        End If

        Dim sender As Object = Nothing
        Dim e As New System.Windows.Forms.KeyEventArgs(Keys.Enter)

        ' Trim Off line feeds
        dataReceived = Replace(dataReceived, Chr(13), String.Empty)
        dataReceived = Replace(dataReceived, Chr(10), String.Empty)
        dataReceived = dataReceived.Trim

        ' Set Sender based on state of the screen

        ProcessEnterKeyStroke(sender, e)
        dataReceived = String.Empty
    End Sub

    ''' <summary>
    ''' Process keyboard 'Enter' key
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ProcessEnterKeyStroke(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)

        Select Case Absx1.GetABSColumnName(sender)
            Case "x"

            Case "y"
        End Select
    End Sub

    ''' <summary>
    ''' Request weight from scale
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub RequestWeightFromScale()

        Try
            registeredWeight = 0

            'If ASCMAIN1.ScaleSerialPort Is Nothing Then Exit Sub

            'If Not ASCMAIN1.ScaleSerialPort.IsOpen Then
            '    ASCMAIN1.ScaleSerialPort.Open()
            'End If

            '' Request the weight from the scale
            'Dim encoding As New System.Text.UTF8Encoding()
            'Dim inBuffer As Byte() = encoding.GetBytes("W")
            'ASCMAIN1.ScaleSerialPort.Write(inBuffer, 0, inBuffer.Length)

        Catch ex As Exception
            MessageBox.Show("Scale Weight Error: " & ex.Message, "Scale Weight", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Fires when the weight is requested from the scale
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ProcessScaleData(ByVal scaledata As String)

        If MdiParent.ActiveMdiChild Is Nothing Then Exit Sub
        If MdiParent.ActiveMdiChild.Name <> Me.Name Then Exit Sub

        Try
            'Dim length As Int16 = ASCMAIN1.ScaleSerialPort.BytesToRead
            'If length > 0 Then
            '    Dim numberOfBytesRead As Int16 = 0
            '    Dim readBuffer(length) As Byte
            '    numberOfBytesRead = ASCMAIN1.ScaleSerialPort.Read(readBuffer, 0, length)
            '    registeredWeight = Val(readBuffer)
            'End If
        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Sends data to the Label Printer
    ''' </summary>
    ''' <param name="LabelData"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PrintLabel(ByVal LabelData As String) As Boolean

        Try
            If ASCMAIN1.USER_ID = "edz" AndAlso ASCMAIN1.Running_in_VS Then
                ' Find Zebra printer

                Dim zebraPrinter As String = FindZebraPrinter()

                Dim vLabelPrinter As New ASCPRINT
                Return vLabelPrinter.SendStringToPrinter(zebraPrinter, LabelData)
            End If

            ASCMAIN1.LabelPrinterSerialPort.WriteLine(LabelData)

        Catch ex As Exception
            MessageBox.Show("Print Shipping Label Error: " & ex.Message)

        End Try

    End Function

    Private Shared Function FindZebraPrinter() As String
        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA450") Then
                Return printerName
            End If
        Next printerName

        For Each printerName As String In Drawing.Printing.PrinterSettings.InstalledPrinters
            If printerName.ToUpper.StartsWith("ZEBRA") Then
                Return printerName
            End If
        Next printerName

        Return ""
    End Function

    Private Sub UpdateOrderPicksAndCreateInvoices()

        Dim INV_DATE As Date = MyBase.Absx1.dteFor("INV_DATE").DateTime

        Dim ORDR_YYYYPP_UPDATED As String = ASCDATA1.GetDataValue("Select MIN(OPS_YYYYPP) from gltparm2 where prd_end_date >= '" & INV_DATE.ToString("dd-MMM-yyyy") & "'") & String.Empty
        If ORDR_YYYYPP_UPDATED.Length = 0 Then
            ORDR_YYYYPP_UPDATED = ASCMAIN1.CYP
        End If

        ' Update Sotordr1 and Sortordr2 and Possibly SOTPICK1,2 
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
            Dim ORDR_STAX As String = Val(rowSOTPICK1.Item("ORDR_STAX") & String.Empty)
            Dim STAX_CODE As String = rowSOTPICK1.Item("STAX_CODE") & String.Empty

            ' Update Cartom Weight based on the weight(s) provided by the user.
            ' Added 1/21/2013
            Dim PICK_TOTAL_WGT As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT

            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
                Dim ORDR_LNO As Int16 = rowSOTPICK2.Item("ORDR_LNO")

                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                rowSOTORDR2.Item("ORDR_QTY_PICK") = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty) - Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                If rowSOTORDR2.Item("ORDR_QTY_PICK") < 0 Then
                    rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                ElseIf rowSOTORDR2.Item("ORDR_QTY_PICK") > 0 Then
                    ' No back Orders
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & String.Empty)
                    rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                End If
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                rowSOTORDR2.Item("ORDR_QTY_SHIP") = Val(rowSOTORDR2.Item("ORDR_QTY_SHIP") & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY_CONF") & String.Empty)
                rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) + (Val(rowSOTPICK2.Item("PICK_QTY_CANC") & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY_BACK") & String.Empty))
                rowSOTORDR2.Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                rowSOTORDR2.Item("ORDR_STATUS") = "F"
            Next

            ' Finalize Sales Order and Pick Ticket
            dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("ORDR_STATUS") = "F"
            dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED ' ASCMAIN1.CYP
            dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("LAST_OPER") = ASCMAIN1.USER_ID
            dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("LAST_DATE") = DATETIME_STAMP
            dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("ORDR_FOB") = txtSHIP_FOB.Text.Trim

            If Val(dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("ORDR_STAX") & String.Empty) <> ORDR_STAX Then
                Dim row As DataRow = dst.Tables("TATEVNT1").Rows.Add
                row.Item("TABLE_NAME") = "SOTORDR1"
                row.Item("TABLE_KEY") = ORDR_NO
                row.Item("INIT_DATE") = DATETIME_STAMP
                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                row.Item("EVENT_TYPE") = "STAX"
                row.Item("EVENT_DESC") = "Sales Tax was changed from " _
                    & Val(dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("ORDR_STAX") & String.Empty) & " to " & ORDR_STAX
                row.Item("EVENT_KEY") = ""
            End If
            dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("ORDR_STAX") = ORDR_STAX
            If ORDR_STAX > 0 Then
                dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("STAX_CODE") = STAX_CODE
            End If

            If dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("SHIP_VIA_CODE") & String.Empty <> MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text Then
                Dim row As DataRow = dst.Tables("TATEVNT1").Rows.Add
                row.Item("TABLE_NAME") = "SOTORDR1"
                row.Item("TABLE_KEY") = ORDR_NO
                row.Item("INIT_DATE") = DATETIME_STAMP
                row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                row.Item("EVENT_TYPE") = "SHPMTC"
                row.Item("EVENT_DESC") = "Ship Via was changed from " _
                    & dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0).Item("SHIP_VIA_CODE") & " to " & MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text
                row.Item("EVENT_KEY") = ""
            End If
        Next

        ' Create Invoice Records
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "", DataViewRowState.CurrentRows)

            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            rowSOTSHIP1.Item("INV_DATE") = MyBase.Absx1.dteFor("INV_DATE").Value
            rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = MyBase.Absx1.dteFor("SHIP_DATE_SHIPPED").Value
            rowSOTSHIP1.Item("SHIP_FOB") = txtSHIP_FOB.Text.Trim

            Dim SOCINVH1 As New TAC.SOCINVH1(dst.Tables("SOTINVH1"), dst.Tables("SOTINVH2"), _
                                              dst.Tables("SOTPICK1"), dst.Tables("SOTPICK2"), _
                                              dst.Tables("ARTOPEN1"), dst.Tables("SOTSHIP1"), _
                                              dst.Tables("SOTORDR5"))
            SOCINVH1.CreateInvoices(SHIP_BOL_NO)
        Next

        ' New Logic
        ' As Per Walter 11/2/2015 email.
        'SOTORDR1.ORDR_DATE_CLOSED               <- SOTINVH1.INV_DATE_SHIPPED
        'SOTORDR1.ORDR_YYYYPP_CLOSED             <- SOTINVH1.ORDR_YYYYPP_UPDATED
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

            rowSOTORDR1.Item("ORDR_DATE_CLOSED") = rowSOTINVH1.Item("INV_DATE_SHIPPED")
            rowSOTORDR1.Item("ORDR_YYYYPP_CLOSED") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")
        Next

        If CURR_CODE = "" Or Val(CURR_EXCH_RATE) = 0 Then
            MessageBox.Show("*****************************************" & Environment.NewLine & _
                            "Please contact ABS about this shipment." & Environment.NewLine & _
                            "Let them know the Currency Code is blank!" & Environment.NewLine & _
                            "*****************************************", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub


#End Region

    Private Sub chkInsureShipment_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkInsureShipment.CheckedChanged
        If chkInsureShipment.Checked Then
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").Hidden = False
        Else
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").CellActivation = UltraWinGrid.Activation.NoEdit
            grdSOTCART1.DisplayLayout.Bands(0).Columns("INSURANCE").Hidden = True
        End If
    End Sub

    Private Sub UltraButton2_Click(sender As System.Object, e As System.EventArgs) Handles UltraButton2.Click
        Try
            ' send test label to printer
            Dim testlabel = TAC.TACMAIN1.CreateSampleLabel
            PrintLabel(testlabel)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Test Print", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Public Function EmailInvoice() As Boolean

        Try

            If dst.Tables("SOTINVH1").Rows.Count = 0 Then
                Return False
            End If

            If ASCMAIN1.DBS_COMPANY = "AHA" OrElse ASCMAIN1.DBS_SERVER = "AHA" Then
                Exit Function
            End If

            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows(0)
            Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO")
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim CONS_INV As String = "0"
            If rowSOTINVH1.Item("INV_NO_CONS") & String.Empty <> String.Empty Then
                CONS_INV = "1"
            End If

            ' See if the customer receives an acknowledgment
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
            Dim rowSOTSREP1 As DataRow = Nothing
            Dim salesRepEmail As String = String.Empty
            Dim custEmailShipAck As String = String.Empty

            If rowARTCUST1 IsNot Nothing Then
                custEmailShipAck = (rowARTCUST1.Item("CUST_EMAIL_SHIP_ACK") & "").ToString.Trim.ToUpper
            End If

            If rowARTCUST2 IsNot Nothing Then
                salesRepEmail = (rowARTCUST2.Item("CUST_STORE_EMAIL") & String.Empty).ToString.Trim.ToUpper
                rowSOTSREP1 = LookUp("SOTSREP1", rowARTCUST2.Item("SREP_CODE") & String.Empty)
                If rowSOTSREP1 IsNot Nothing Then
                    salesRepEmail = (rowSOTSREP1.Item("SREP_EMAIL") & String.Empty).ToString.Trim.ToUpper
                End If
            End If

            If custEmailShipAck.Length = 0 AndAlso salesRepEmail.Length = 0 Then
                Return True
            End If

            Dim invNos As String = String.Empty
            For Each row As DataRow In dst.Tables("SOTINVH1").Select("")
                invNos &= ", '" & row.Item("INV_NO") & "'"
            Next
            invNos = invNos.Substring(1).Trim

            Dim RPT As String = "SORINVP1"

            If Not REPORTS.ContainsKey(RPT) Then
                REPORTS.Add(RPT, Load_rptClass(RPT))
                REPORTS(RPT).Prepare_dst(False, "")
            End If

            REPORTS(RPT).Fill_Records_RPT(New String() {" and SOTINVH1.INV_NO IN (" & invNos & ")"})

            Dim REPORT_NO As String = ""

            With REPORTS(RPT).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", CONS_INV)
                REPORT_NO = .Generate_Report(RPT, "Invoice", , True, , , "PDF", INV_NO, False)
                .Print_Report_End(True, True)
            End With

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(INV_NO & ".pdf", ASCMAIN1.Folders("Temp") & INV_NO & ".pdf")

            Dim SUBJECT As String = "Invoice " & INV_NO

            Select Case ASCMAIN1.DBS_COMPANY
                Case "AHA"
                    SUBJECT = "Ahava Invoice " & INV_NO
                Case "INT"
                    SUBJECT = "Interparfums Invoice " & INV_NO
            End Select

            ' Concatentate and process all email addresses
            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            For Each emailAddress As String In (custEmailShipAck & ";" & salesRepEmail).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Return True
            End If

            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                    SUBJECT, "INV", True, False, CUST_CODE, rowARTCUST1.Item("CUST_NAME"), "Customer")

            EmailInvoice = True
        Catch ex As Exception
            EmailInvoice = False
        End Try

    End Function

    Private Sub btnShipOptions_Click(sender As System.Object, e As System.EventArgs) Handles btnShipOptions.Click
        Try
            requestShippingOptions = True
            Dim shipLabels As New List(Of String)
            Dim errorMessage As String = String.Empty

            RequestShippingLabel(shipLabels, errorMessage, False)

            If errorMessage.Length > 0 Then
                MessageBox.Show(errorMessage, "Ship Options", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            Sort_grdColumns(grdSOTCARRX, "CARRIER_CODE,CARRIER_PROD_CODE")

        Catch ex As Exception
            MessageBox.Show("The following error occurred: " & ex.Message, "Ship Options", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            requestShippingOptions = False
        End Try
    End Sub

End Class

Public Class Price_Change
    Public PICK_NO As String
    Public PICK_LNO As Int32
    Public ITEM_CODE As String
    Public PICK_UNIT_PRICE As Decimal
End Class