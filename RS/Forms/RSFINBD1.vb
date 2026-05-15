Imports ABSolution
Imports Infragistics.Win
Imports System.Windows.Forms

Public Class RSFINBD1
    Private _downloadedFiles As List(Of String) = New List(Of String)
    Private _remoteDirectoryFileList As List(Of String) = New List(Of String)
    Private Const _EDIFileExt As String = ".edi"
    Private Const _TextFileExt As String = ".txt"

    Private downloadFiles As Boolean = False

    Dim downloadFileLocation As String = ""
    Dim downloadFileLocationArchive As String = ""

    Dim ICTITEM1_TEMPsql As String = ""
    Dim ICTITEM1_TEMP As String = ""

    Dim ICTSTAT2_TEMPsql As String = ""
    Dim ICTSTAT2_TEMP As String = ""

    Dim ARTCUST1_TEMPsql As String = ""
    Dim ARTCUST1_TEMP As String = ""

    Dim SOTORDR1_TEMPsql As String = ""
    Dim SOTORDR1_TEMP As String = ""

    Dim SOTORDR2_TEMPsql As String = ""
    Dim SOTORDR2_TEMP As String = ""

    Dim SOTORDRPsql As String = ""
    Dim SOTORDRIsql As String = ""

    Dim shortFileName As String = ""

    Dim viewSOTORDR2_IN As DataView
    Dim viewSOTINVH2_IN As DataView
    Dim viewSOTINVH2_CR As DataView

    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("EDTPARM1")
        Get_PARM("RSTPARM1")
        downloadFileLocation = ROWs("RSTPARM1").Item("RS_PARM_INBOUND") & ""
        downloadFileLocationArchive = ROWs("RSTPARM1").Item("RS_PARM_INBOUND_ARCHIVE") & ""

        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, 0, 0)

        With dst
            With .Tables.Add("RSTFILE1")
                .Columns.Add("FILENAME")
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .Columns.Add("FILEDATETIME", GetType(System.DateTime))
            End With
        End With

        grdRSTFILE1.DataSource = dst.Tables("RSTFILE1")
        grdRSTFILE1.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdRSTFILE1.DisplayLayout.Bands(0).SortedColumns.Add("FILEDATETIME", False)

        Create_Summary(grdRSTFILE1, "FILENAME", "Count")

        With dst
            With .Tables.Add("ICTITEM1_IN")
                With .Columns
                    .Add("ITEM_CODE", GetType(System.String))
                    .Add("ITEM_DESC", GetType(System.String))
                    .Add("ITEM_DESC2", GetType(System.String))
                    .Add("ITEM_UOM", GetType(System.String))
                    '.Add("PRODUCT_GROUP_CODE", GetType(System.String))
                    .Add("COSTING_METHOD", GetType(System.String))
                    .Add("AVG_COST", GetType(System.Decimal))
                    .Add("UNIT_COST", GetType(System.Decimal))
                    .Add("ITEM_COST_STD", GetType(System.Decimal))
                    .Add("ITEM_RETAIL_PRICE", GetType(System.Decimal))
                    .Add("GEN_PROD_POST_GRP", GetType(System.String)) 'PROD CODE
                    .Add("INV_POST_GRP", GetType(System.String))      'COST_CODE
                    .Add("ITEM_UPC_EAN", GetType(System.String))
                    .Add("QTY_ON_HAND", GetType(System.Int64))
                    .Add("QTY_ON_SO", GetType(System.Int64))
                    .Add("QTY_AVAIL", GetType(System.Int64))
                    .Add("QTY_ON_PO", GetType(System.Int64))
                    .Add("PROD_LINE", GetType(System.String))         'COLL_CODE
                End With
            End With

            With .Tables("ICTITEM1_IN")
                .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE")}
            End With

            With .Tables.Add("ARTCUST1_IN")
                With .Columns
                    .Add("CUST_CODE", GetType(System.String))
                    .Add("NAME", GetType(System.String))
                    .Add("NAME2", GetType(System.String))
                    .Add("ADDR1", GetType(System.String))
                    .Add("ADDR2", GetType(System.String))
                    .Add("CITY", GetType(System.String))
                    .Add("STATE", GetType(System.String))
                    .Add("ZIP", GetType(System.String))
                    .Add("COUNTRY", GetType(System.String))
                    .Add("PHONE", GetType(System.String))
                    .Add("FAX", GetType(System.String))
                    .Add("CONTACT", GetType(System.String))
                    .Add("SALESPERSON", GetType(System.String))
                    .Add("DISC_PCT", GetType(System.String))
                    .Add("GEN_BUS_POSTING_GROUP", GetType(System.String))
                    .Add("CUSTOMER_POSTING_GROUP", GetType(System.String))
                    .Add("PAYMENT_TERMS_CODE", GetType(System.String))
                    .Add("SHIPPING_PAYMENT_TYPE", GetType(System.String))
                End With
            End With

            With .Tables("ARTCUST1_IN")
                .PrimaryKey = New DataColumn() {.Columns("CUST_CODE")}
            End With

            With .Tables.Add("SOTORDR1_IN")
                With .Columns
                    .Add("ORDR_NO", GetType(System.String))
                    .Add("SELL_TO_CUST_NO", GetType(System.String))
                    .Add("SELL_TO_NAME", GetType(System.String))
                    .Add("SELL_TO_ADDR", GetType(System.String))
                    .Add("SELL_TO_ADDR2", GetType(System.String))
                    .Add("SELL_TO_CITY", GetType(System.String))
                    .Add("SELL_TO_COUNTY", GetType(System.String))
                    .Add("SELL_TO_POST_CODE", GetType(System.String))
                    .Add("SELL_TO_CONTACT", GetType(System.String))
                    .Add("POSTING_DATE", GetType(System.DateTime))
                    .Add("ORDR_DATE", GetType(System.DateTime))
                    .Add("DOCUMENT_DATE", GetType(System.DateTime))
                    .Add("REQUESTED_DELIVERY_DATE", GetType(System.DateTime))
                    .Add("EXTERNAL_DOC_NO", GetType(System.String))
                    .Add("PAYMENT_TERMS_CODE", GetType(System.String))
                    .Add("SALESPERSON_CODE", GetType(System.String))
                    .Add("STATUS", GetType(System.String))
                    .Add("TOTAL_POSTED_PKG", GetType(System.Int64))
                    .Add("TOTAL_PKGS", GetType(System.Int64))
                    .Add("BILL_TO_CUST_NO", GetType(System.String))
                    .Add("BILL_TO_NAME", GetType(System.String))
                    .Add("BILL_TO_ADDR", GetType(System.String))
                    .Add("BILL_TO_ADDR2", GetType(System.String))
                    .Add("BILL_TO_CITY", GetType(System.String))
                    .Add("BILL_TO_COUNTY", GetType(System.String))
                    .Add("BILL_TO_POST_CODE", GetType(System.String))
                    .Add("BILL_TO_CONTACT_NO", GetType(System.String))
                    .Add("BILL_TO_CONTACT", GetType(System.String))
                    .Add("PAYMENT_TERMS_CODE_2", GetType(System.String))
                    .Add("DUE_DATE", GetType(System.DateTime))
                    .Add("PAYMENT_DISC_PCT", GetType(System.Int64))
                    .Add("PAYMENT_DISC_DATE", GetType(System.DateTime))
                    .Add("PAYMENT_METHOD_CODE", GetType(System.String))
                    .Add("TAX_AREA_CODE", GetType(System.String))
                    .Add("TAX_LIABLE", GetType(System.String))
                    .Add("SHIP_TO_CODE", GetType(System.String))
                    .Add("SHIP_TO_NAME", GetType(System.String))
                    .Add("SHIP_TO_ADDR", GetType(System.String))
                    .Add("SHIP_TO_ADDR2", GetType(System.String))
                    .Add("SHIP_TO_CITY", GetType(System.String))
                    .Add("SHIP_TO_COUNTY", GetType(System.String))
                    .Add("SHIP_TO_POST_CODE", GetType(System.String))
                    .Add("SHIP_TO_COUNTRY_CODE", GetType(System.String))
                    .Add("SHIP_TO_CONTACT", GetType(System.String))
                    .Add("SHIP_TO_SALESPERSON", GetType(System.String))
                    .Add("LOCATION_CODE", GetType(System.String))
                    .Add("SHIPMENT_METHOD_CODE", GetType(System.String))
                    .Add("SHIPPING_AGENT_CODE", GetType(System.String))
                    .Add("E_SHIP_AGENT_SVC", GetType(System.String))
                    .Add("SHIPMENT_DATE", GetType(System.DateTime))
                    .Add("SHIPPING_ADVICE", GetType(System.String))
                    .Add("EARLIEST_SHIP_DATE", GetType(System.DateTime))
                    .Add("CANCEL_AFTER_DATE", GetType(System.DateTime))
                    .Add("RESIDENTIAL_DELIVERY", GetType(System.String))
                    .Add("SHIPPING_ADVICE_2", GetType(System.String))
                    .Add("FREE_FREIGHT", GetType(System.String))
                    .Add("EDI_ORDER", GetType(System.String))
                    .Add("EDI_DEPT", GetType(System.String))
                    .Add("EDI_EXPECTED_DELIVERY_DATE", GetType(System.DateTime))
                    .Add("EDI_TRADE_PARTNER", GetType(System.String))
                    .Add("EDI_SELL_TO_CODE", GetType(System.String))
                    .Add("EDI_SHIP_FOR_CODE", GetType(System.String))
                    .Add("EDI_SHIP_TO_CODE", GetType(System.String))
                    .Add("EDI_CANCEL_AFTER_DATE", GetType(System.DateTime))
                    .Add("EXTERNAL_DOC_NO_2", GetType(System.String))
                    .Add("SHIP_FOR_CODE", GetType(System.String))
                    .Add("CREATED_BY", GetType(System.String))
                    .Add("CREATION_DATE", GetType(System.DateTime))
                    .Add("ON_HOLD", GetType(System.String))
                    .Add("RELEASED_BY", GetType(System.String))
                    .Add("RELEASED_DATE", GetType(System.DateTime))
                    .Add("OPENED_BY", GetType(System.String))
                    .Add("OPENED_DATE", GetType(System.DateTime))
                End With
            End With

            With .Tables("SOTORDR1_IN")
                .PrimaryKey = New DataColumn() {.Columns("ORDR_NO")}
            End With

            With .Tables.Add("SOTORDR2_IN")
                With .Columns
                    .Add("ORDR_NO", GetType(System.String))
                    .Add("ORDR_LNO", GetType(System.Int64)) 'assign
                    .Add("LINE_ITEM_TYPE", GetType(System.String))
                    .Add("ITEM_CODE", GetType(System.String))
                    .Add("DESCRIPTION", GetType(System.String))
                    .Add("QTY", GetType(System.Int64))
                    .Add("UNIT_PRICE", GetType(System.Decimal))
                    .Add("LINE_DISC_PCT", GetType(System.Int64))
                    .Add("LINE_AMT", GetType(System.Decimal))
                    .Add("QTY_TO_SHIP", GetType(System.Int64))
                    .Add("QTY_SHIPPED", GetType(System.Int64))
                    .Add("QTY_INVOICED", GetType(System.Int64))
                    .Add("UOM", GetType(System.String))
                    .Add("ORIG_QTY_ORD", GetType(System.Int64))
                    .Add("UNIT_COST", GetType(System.Decimal))
                    .Add("TAX_GROUP_CODE", GetType(System.String))
                    .Add("AMT_INCL_TAX", GetType(System.Decimal))
                    .Add("LINE_DISC_AMT", GetType(System.Decimal))
                    .Add("ALLOW_INV_DISC", GetType(System.String))
                    .Add("INV_DISC_AMT", GetType(System.Decimal))
                    .Add("PLANNED_DELIVERY_DATE", GetType(System.DateTime))
                    .Add("PLANNED_SHIPMENT_DATE", GetType(System.DateTime))
                    .Add("SHIPMENT_DATE", GetType(System.DateTime))
                End With
            End With

            With .Tables("SOTORDR2_IN")
                .PrimaryKey = New DataColumn() {.Columns("ORDR_NO"), .Columns("ORDR_LNO")}
            End With

            With .Tables.Add("SOTINVH1_IN")
                With .Columns
                    .Add("INV_NO", GetType(System.String))
                    .Add("SELL_TO_CUST_NO", GetType(System.String))
                    .Add("SELL_TO_NAME", GetType(System.String))
                    .Add("SELL_TO_ADDR", GetType(System.String))
                    .Add("SELL_TO_ADDR2", GetType(System.String))
                    .Add("SELL_TO_CITY", GetType(System.String))
                    .Add("SELL_TO_STATE", GetType(System.String))
                    .Add("SELL_TO_ZIP", GetType(System.String))
                    .Add("SELL_TO_CONTACT", GetType(System.String))
                    .Add("POSTING_DATE", GetType(System.DateTime))
                    .Add("DOCUMENT_DATE", GetType(System.DateTime))
                    .Add("ORDR_NO", GetType(System.String))
                    .Add("EXTERNAL_DOC_NO", GetType(System.String))
                    .Add("SALESPERSON_CODE", GetType(System.String))
                    .Add("NO_PRINTED", GetType(System.Int64))
                    .Add("TOTAL_PKGS", GetType(System.Int64))
                    .Add("TOTAL_WT", GetType(System.Int64))
                    .Add("AMOUNT_INCL_TAX", GetType(System.Double))
                    .Add("BILL_TO_CUST_NO", GetType(System.String))
                    .Add("BILL_TO_NAME", GetType(System.String))
                    .Add("BILL_TO_ADDR", GetType(System.String))
                    .Add("BILL_TO_ADDR2", GetType(System.String))
                    .Add("BILL_TO_CITY", GetType(System.String))
                    .Add("BILL_TO_STATE", GetType(System.String))
                    .Add("BILL_TO_ZIP", GetType(System.String))
                    .Add("BILL_TO_CONTACT_NO", GetType(System.String))
                    .Add("BILL_TO_CONTACT", GetType(System.String))
                    .Add("PROD_LINE_CODE", GetType(System.String))
                    .Add("COMPANY_CODE", GetType(System.String))
                    .Add("PAYMENT_TERMS_CODE", GetType(System.String))
                    .Add("DUE_DATE", GetType(System.DateTime))
                    .Add("PAYMENT_DISC_PCT", GetType(System.Int64))
                    .Add("PAYMENT_DISC_DATE", GetType(System.DateTime))
                    '.Add("PAYMENT_METHOD_CODE", GetType(System.String))
                    .Add("TAX_AREA_CODE", GetType(System.String))
                    .Add("TAX_LIABLE", GetType(System.String))
                    .Add("SHIP_TO_CODE", GetType(System.String))
                    .Add("SHIP_TO_NAME", GetType(System.String))
                    .Add("SHIP_TO_ADDR", GetType(System.String))
                    .Add("SHIP_TO_ADDR2", GetType(System.String))
                    .Add("SHIP_TO_CITY", GetType(System.String))
                    .Add("SHIP_TO_STATE", GetType(System.String))
                    .Add("SHIP_TO_ZIP", GetType(System.String))
                    .Add("SHIP_TO_COUNTRY_CODE", GetType(System.String))
                    .Add("SHIP_TO_CONTACT", GetType(System.String))
                    .Add("SHIP_TO_SALESPERSON", GetType(System.String))
                    .Add("LOCATION_CODE", GetType(System.String))
                    .Add("SHIPMENT_METHOD_CODE", GetType(System.String))
                    .Add("SHIPMENT_DATE", GetType(System.DateTime))
                    .Add("SHIP_DATE", GetType(System.DateTime))
                    .Add("SHIPPING_AGENT_CODE", GetType(System.String))
                    .Add("E_SHIP_AGENT_SVC", GetType(System.String))
                    .Add("RESIDENTIAL_DELIVERY", GetType(System.String))
                    .Add("SHIP_FOR_CODE", GetType(System.String))
                    .Add("SHIPPING_PAYMENT_TYPE", GetType(System.String))
                    .Add("SHIPPING_INSURANCE", GetType(System.String))
                    .Add("FREE_FREIGHT", GetType(System.String))
                    .Add("INVOICE_FOR_BOL", GetType(System.String))
                    .Add("INVOICE_FOR_SHIPMENT", GetType(System.String))
                    .Add("SHIPMENT_INV_OVERRIDE", GetType(System.String))
                    .Add("EDI_ORDER", GetType(System.String))
                    .Add("EDI_INV_GEN", GetType(System.String))
                    .Add("EDI_INV_GEN_DATE", GetType(System.DateTime))
                    .Add("EDI_TRADE_PARTNER", GetType(System.String))
                    .Add("EDI_SELL_TO_CODE", GetType(System.String))
                    .Add("EDI_SHIP_FOR_CODE", GetType(System.String))
                    .Add("EDI_SHIP_TO_CODE", GetType(System.String))
                    .Add("CANCEL_AFTER_DATE", GetType(System.DateTime))
                    .Add("EXTERNAL_SHIP_FOR_NO", GetType(System.String))
                    .Add("SHIP_FOR_CODE2", GetType(System.String))
                    .Add("CREATED_BY", GetType(System.String))
                    .Add("CREATION_DATE", GetType(System.DateTime))
                    .Add("RELEASED_BY", GetType(System.String))
                    .Add("RELEASED_DATE", GetType(System.DateTime))
                    .Add("OPENED_BY", GetType(System.String))
                    .Add("OPENED_DATE", GetType(System.DateTime))
                End With
            End With

            With .Tables("SOTINVH1_IN")
                .PrimaryKey = New DataColumn() {.Columns("INV_NO")}
            End With

            With .Tables.Add("SOTINVH2_IN")
                With .Columns
                    .Add("INV_NO", GetType(System.String))
                    .Add("INV_LNO", GetType(System.Int64)) 'assign
                    .Add("LINE_ITEM_TYPE", GetType(System.String))
                    .Add("ITEM_CODE", GetType(System.String))
                    .Add("DESCRIPTION", GetType(System.String))
                    .Add("RET_REASON_CODE", GetType(System.String))
                    .Add("PACKAGE_TRACKING_NO", GetType(System.String))
                    .Add("SHIPMENT_NO", GetType(System.String))
                    .Add("QTY", GetType(System.Int64))
                    .Add("QTY_ORD", GetType(System.Int64))
                    .Add("UOM_CODE", GetType(System.String))
                    .Add("UNIT_COST", GetType(System.Decimal))
                    .Add("TAX_GROUP_CODE", GetType(System.String))
                    .Add("UNIT_PRICE", GetType(System.Decimal))
                    .Add("LINE_AMT", GetType(System.Decimal))
                    .Add("AMT_INCL_TAX", GetType(System.Decimal))
                    .Add("LINE_DISC_PCT", GetType(System.Int64))
                    .Add("LINE_DISC_AMT", GetType(System.Decimal))
                    .Add("ALLOW_INV_DISC", GetType(System.String))
                    .Add("PROD_LINE_CODE", GetType(System.String))
                    .Add("COMPANY_CODE", GetType(System.String))
                End With
            End With

            With .Tables("SOTINVH2_IN")
                .PrimaryKey = New DataColumn() {.Columns("INV_NO"), .Columns("INV_LNO")}
            End With

            'Credits
            With .Tables.Add("SOTINVH1_CR")
                With .Columns
                    .Add("SELL_TO_CUST_NO", GetType(System.String))
                    .Add("CM_NO", GetType(System.String))
                    .Add("BILL_TO_CUST_NO", GetType(System.String))
                    .Add("BILL_TO_NAME", GetType(System.String))
                    .Add("BILL_TO_NAME2", GetType(System.String))
                    .Add("BILL_TO_ADDR", GetType(System.String))
                    .Add("BILL_TO_ADDR2", GetType(System.String))
                    .Add("BILL_TO_CITY", GetType(System.String))
                    .Add("BILL_TO_STATE", GetType(System.String)) 'STATE
                    .Add("BILL_TO_ZIP", GetType(System.String)) 'ZIP
                    .Add("BILL_TO_COUNTRY_CODE", GetType(System.String))
                    .Add("BILL_TO_CONTACT", GetType(System.String))
                    .Add("YOUR REFERENCE", GetType(System.String)) 'ADDED
                    .Add("SHIP_TO_CODE", GetType(System.String))
                    .Add("SHIP_TO_NAME", GetType(System.String))
                    .Add("SHIP_TO_NAME2", GetType(System.String))
                    .Add("SHIP_TO_ADDR", GetType(System.String))
                    .Add("SHIP_TO_ADDR2", GetType(System.String))
                    .Add("SHIP_TO_CITY", GetType(System.String))
                    .Add("SHIP_TO_STATE", GetType(System.String)) 'STATE
                    .Add("SHIP_TO_ZIP", GetType(System.String)) 'ZIP
                    .Add("SHIP_TO_COUNTRY_CODE", GetType(System.String))
                    .Add("SHIP_TO_CONTACT", GetType(System.String))
                    .Add("SHIP_TO_SALESPERSON", GetType(System.String))
                    .Add("POSTING_DATE", GetType(System.DateTime))
                    .Add("SHIPMENT_DATE", GetType(System.DateTime))
                    .Add("POSTING_DESC", GetType(System.String)) 'added
                    .Add("PAYMENT_TERMS_CODE", GetType(System.String))
                    .Add("DUE_DATE", GetType(System.DateTime))
                    .Add("PAYMENT_DISC_PCT", GetType(System.Int64))
                    .Add("PAYMENT_DISC_DATE", GetType(System.DateTime))
                    .Add("SHIPMENT_METHOD_CODE", GetType(System.String))
                    .Add("LOCATION_CODE", GetType(System.String))
                    .Add("SHORTCUT_DIM_1_CODE", GetType(System.String)) 'ADDED
                    .Add("SHORTCUT_DIM_2_CODE", GetType(System.String)) 'ADDED
                    .Add("CUSTOMER_POSTING_GROUP", GetType(System.String)) 'ADDED
                    .Add("CURRENCY_CODE", GetType(System.String)) 'ADDED
                    .Add("CURRENCY_FACTOR", GetType(System.String)) 'ADDED
                    .Add("CUST_PRICE_GROUP", GetType(System.String)) 'ADDED
                    .Add("PRICES_INCL_VAT", GetType(System.Double)) 'ADDED
                    .Add("INV_DISC_CODE", GetType(System.String)) 'ADDED
                    .Add("CUST_DISC_GROUP", GetType(System.String)) 'ADDED
                    .Add("LANGUAGE_CODE", GetType(System.String)) 'ADDED
                    .Add("SALESPERSON_CODE", GetType(System.String))
                    .Add("COMMENT", GetType(System.String)) 'ADDED
                    .Add("NO_PRINTED", GetType(System.Int64))
                    .Add("ON_HOLD", GetType(System.String)) 'ADDED
                    .Add("APPLIES_TO_DOC_TYPE", GetType(System.String)) 'ADDED
                    .Add("APPLIES_TO_DOC_NO", GetType(System.String)) 'ADDED
                    .Add("BALANCE_ACCT_NO", GetType(System.String)) 'ADDED
                    .Add("JOB_NO", GetType(System.String)) 'ADDED
                    .Add("AMOUNT", GetType(System.Double)) 'ADDED
                    .Add("AMOUNT_INCL_VAT", GetType(System.Double)) 'ADDED
                    .Add("VAT_REG_NO", GetType(System.String)) 'ADDED
                    .Add("RET_REASON_CODE", GetType(System.String)) 'ADDED
                    .Add("GEN_BUS_POSTING_GROUP", GetType(System.String)) 'ADDED
                    .Add("EU_3_PARTY_TRADE", GetType(System.String)) 'ADDED
                    .Add("TRANSACTION_TYPE", GetType(System.String)) 'ADDED
                    .Add("TRANSPORT_METHOD", GetType(System.String)) 'ADDED
                    .Add("VAT_COUNTRY_CODE", GetType(System.String)) 'ADDED
                    .Add("SELL_TO_NAME", GetType(System.String))
                    .Add("SELL_TO_NAME2", GetType(System.String))
                    .Add("SELL_TO_ADDR", GetType(System.String))
                    .Add("SELL_TO_ADDR2", GetType(System.String))
                    .Add("SELL_TO_CITY", GetType(System.String))
                    .Add("SELL_TO_CONTACT", GetType(System.String))
                    .Add("SELL_TO_STATE", GetType(System.String)) 'STATE
                    .Add("SELL_TO_ZIP", GetType(System.String)) 'ZIP
                    .Add("SELL_TO_COUNTRY_CODE", GetType(System.String))

                    'ADDED
                    .Add("BAL_ACCT_TYPE", GetType(System.String))
                    .Add("EXIT_POINT", GetType(System.String))
                    .Add("CORRECTION", GetType(System.String))
                    .Add("DOCUMENT_DATE", GetType(System.DateTime))
                    .Add("EXTERNAL_DOC_NO", GetType(System.String))

                    'ADDED
                    .Add("AREA", GetType(System.String))
                    '.Add("TRANS_SPEC", GetType(System.String))
                    .Add("PAYMENT_METH_CODE", GetType(System.String))
                    '.Add("PRE_ASSIGNED_NO_SERIES", GetType(System.String))
                    '.Add("NO_SERIES", GetType(System.String))
                     .Add("PRE_ASSIGNED_NO", GetType(System.String))
                    .Add("USER_ID", GetType(System.String))
                    .Add("SOURCE_CODE", GetType(System.String))
                    .Add("TAX_AREA_CODE", GetType(System.String))
                    .Add("TAX_LIABLE", GetType(System.String))
                    .Add("VAT_BUS_POSTING_GROUP", GetType(System.String))
                    .Add("VAT_BASE_DISC_PCT", GetType(System.String))
                    '.Add("CAMPAIGN_NO", GetType(System.String))
                    .Add("SELL_TO_CONTACT_NO", GetType(System.String))
                    .Add("BILL_TO_CONTACT_NO", GetType(System.String))
                    '.Add("RESPONSIBILITY_CTR", GetType(System.String))
                    '.Add("SERVICE_MGMT_DOC", GetType(System.String))
                    .Add("RETURN_ORDER_NO", GetType(System.String))
                    '.Add("RETURN_ORDER_NO_SERIES", GetType(System.String))
                    .Add("ALLOW_LINE_DISC", GetType(System.String))
                    '.Add("SHIP_TO_UPS_ZONE", GetType(System.String))
                    .Add("TAX_EXEMPT_NO", GetType(System.String))
                    '.Add("PRE_PACK_CONF_SENT", GetType(System.String))
                    '.Add("PRE_PACK_CONF_REQ", GetType(System.String))
                    '.Add("ORDER_SUBMITTED", GetType(System.String))
                    '.Add("EARLIEST_SHIP_DATE", GetType(System.DateTime))
                    .Add("EDI_DEPT", GetType(System.String))
                    '.Add("EDI_CANCEL_AFTER_DATE", GetType(System.DateTime))
                    '.Add("SENT_TO_QB", GetType(System.String))
                    '.Add("DATA_SENT_TO_QB", GetType(System.String))
                    .Add("EDI_ORDER", GetType(System.String))
                    '.Add("EDI_INTERNAL_DOC_NO", GetType(System.String))
                    '.Add("EDI_CM_GENERATED", GetType(System.String))
                    '.Add("EDI_CM_GEN_DATE", GetType(System.DateTime))
                    .Add("EDI_TRADE_PARTNER", GetType(System.String))
                    .Add("EDI_SELL_TO_CODE", GetType(System.String))
                    .Add("EDI_SHIP_FOR_CODE", GetType(System.String))
                    .Add("EDI_SHIP_TO_CODE", GetType(System.String))
                    .Add("DATE_SENT", GetType(System.DateTime))
                    .Add("TIME_SENT", GetType(System.DateTime))

                End With
            End With

            With .Tables("SOTINVH1_CR")
                .PrimaryKey = New DataColumn() {.Columns("CM_NO")}
            End With

            With .Tables.Add("SOTINVH2_CR")
                With .Columns
                    .Add("CM_NO", GetType(System.String))
                    .Add("CM_LNO", GetType(System.Int64)) 'assign
                    .Add("LINE_ITEM_TYPE", GetType(System.String))
                    .Add("ITEM_CODE", GetType(System.String))
                    .Add("DESCRIPTION", GetType(System.String))
                    .Add("RET_REASON_CODE", GetType(System.String))
                    .Add("PACKAGE_TRACKING_NO", GetType(System.String))
                    .Add("QTY", GetType(System.Int64))
                    .Add("UOM", GetType(System.String))
                    .Add("UNIT_COST", GetType(System.Decimal))
                    .Add("TAX_GROUP_CODE", GetType(System.String))
                    .Add("UNIT_PRICE", GetType(System.Decimal))
                    .Add("LINE_AMT", GetType(System.Decimal))
                    .Add("AMT_INCL_TAX", GetType(System.Decimal))
                    .Add("LINE_DISC_PCT", GetType(System.Int64))
                    .Add("LINE_DISC_AMT", GetType(System.Decimal))
                    .Add("ALLOW_INV_DISC", GetType(System.String))
                    .Add("PROD_LINE_CODE", GetType(System.String))
                    .Add("COMPANY_CODE", GetType(System.String))
                End With
            End With

            With .Tables("SOTINVH2_CR")
                .PrimaryKey = New DataColumn() {.Columns("CM_NO"), .Columns("CM_LNO")}
            End With
        End With

        ICTITEM1_TEMPsql = "Select * from ICTITEM1 where rownum < 1"
        ICTITEM1_TEMP = ASCMAIN1.Temp_Table(ICTITEM1_TEMPsql)
        Call Create_TDA(dst.Tables.Add, ICTITEM1_TEMP, "*", 0, True)

        ICTSTAT2_TEMPsql = "Select * from ICTSTAT2 where rownum < 1"
        ICTSTAT2_TEMP = ASCMAIN1.Temp_Table(ICTSTAT2_TEMPsql)
        Call Create_TDA(dst.Tables.Add, ICTSTAT2_TEMP, "*", 0, True)

        grdICTITEM1_IN.DataSource = dst.Tables("ICTITEM1_IN")
        grdICTITEM1_IN.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdICTITEM1_IN.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False)
        Create_Summary(grdICTITEM1_IN, "ITEM_CODE", "Count")

        'Items Pending Update (Add & Change)
        Call Create_TDA(dst.Tables.Add, "ICTITEM1_PEND", _
            "Select * from " & ICTITEM1_TEMP & " where ITEM_STATUS in ('A','C')", 0, False)
        grdICTITEM1_PEND.DataSource = dst.Tables("ICTITEM1_PEND")
        grdICTITEM1_PEND.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdICTITEM1_PEND.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False)
        Create_Summary(grdICTITEM1_PEND, "ITEM_CODE", "Count")

        ARTCUST1_TEMPsql = "Select * from ARTCUST1 where rownum < 1"
        ARTCUST1_TEMP = ASCMAIN1.Temp_Table(ARTCUST1_TEMPsql)
        Call Create_TDA(dst.Tables.Add, ARTCUST1_TEMP, "*", 0, True)

        grdARTCUST1_IN.DataSource = dst.Tables("ARTCUST1_IN")
        grdARTCUST1_IN.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdARTCUST1_IN.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False)
        Create_Summary(grdARTCUST1_IN, "CUST_CODE", "Count")

        'Customers Pending Update (Add & Change)
        Call Create_TDA(dst.Tables.Add, "ARTCUST1_PEND", _
            "Select * from " & ARTCUST1_TEMP & " where CUST_STATUS in ('A','C')", 0, False)
        grdARTCUST1_PEND.DataSource = dst.Tables("ARTCUST1_PEND")
        grdARTCUST1_PEND.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdARTCUST1_PEND.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False)
        Create_Summary(grdARTCUST1_PEND, "CUST_CODE", "Count")

        grdSOTORDR1_IN.DataSource = dst.Tables("SOTORDR1_IN")
        grdSOTORDR1_IN.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTORDR1_IN.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_NO", False)
        Create_Summary(grdSOTORDR1_IN, "ORDR_NO", "Count")

        grdSOTORDR2_IN.DataSource = dst.Tables("SOTORDR2_IN")
        grdSOTORDR2_IN.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTORDR2_IN.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_NO", False)
        'grdSOTORDR2_IN.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_LNO", False)
        Create_Summary(grdSOTORDR2_IN, "ORDR_LNO", "Count")

        viewSOTORDR2_IN = New DataView(dst.Tables("SOTORDR2_IN"))
        viewSOTORDR2_IN.RowFilter = "ORDR_NO = '**'"
        viewSOTORDR2_IN.Sort = "ORDR_LNO"
        grdSOTORDR2_IN.DataSource = viewSOTORDR2_IN

        grdSOTINVH1_IN.DataSource = dst.Tables("SOTINVH1_IN")
        grdSOTINVH1_IN.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTINVH1_IN.DisplayLayout.Bands(0).SortedColumns.Add("INV_NO", False)
        Create_Summary(grdSOTINVH1_IN, "INV_NO", "Count")
        Create_Summary(grdSOTINVH1_IN, "AMOUNT_INCL_TAX")

        grdSOTINVH2_IN.DataSource = dst.Tables("SOTINVH2_IN")
        grdSOTINVH2_IN.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTINVH2_IN.DisplayLayout.Bands(0).SortedColumns.Add("INV_NO", False)
        'grdSOTORDR2_IN.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_LNO", False)
        Create_Summary(grdSOTINVH2_IN, "INV_LNO", "Count")
        Create_Summary(grdSOTINVH2_IN, "LINE_AMT")

        viewSOTINVH2_IN = New DataView(dst.Tables("SOTINVH2_IN"))
        viewSOTINVH2_IN.RowFilter = "INV_NO = '**'"
        viewSOTINVH2_IN.Sort = "INV_LNO"
        grdSOTINVH2_IN.DataSource = viewSOTINVH2_IN

        'credits
        grdSOTINVH1_CR.DataSource = dst.Tables("SOTINVH1_CR")
        grdSOTINVH1_CR.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTINVH1_CR.DisplayLayout.Bands(0).SortedColumns.Add("CM_NO", False)
        Create_Summary(grdSOTINVH1_CR, "CM_NO", "Count")
        Create_Summary(grdSOTINVH1_CR, "AMOUNT")

        grdSOTINVH2_CR.DataSource = dst.Tables("SOTINVH2_CR")
        grdSOTINVH2_CR.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTINVH2_CR.DisplayLayout.Bands(0).SortedColumns.Add("CM_NO", False)
        'grdSOTORDR2_CR.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_LNO", False)
        Create_Summary(grdSOTINVH2_CR, "CM_LNO", "Count")
        Create_Summary(grdSOTINVH2_CR, "LINE_AMT")

        viewSOTINVH2_CR = New DataView(dst.Tables("SOTINVH2_CR"))
        viewSOTINVH2_CR.RowFilter = "CM_NO = '**'"
        viewSOTINVH2_CR.Sort = "CM_LNO"
        grdSOTINVH2_CR.DataSource = viewSOTINVH2_CR


        SOTORDR1_TEMPsql = "Select * from SOTORDR1 where rownum < 1"
        SOTORDR1_TEMP = ASCMAIN1.Temp_Table(SOTORDR1_TEMPsql)
        Call Create_TDA(dst.Tables.Add, SOTORDR1_TEMP, "*", 0, True)

        SOTORDR2_TEMPsql = "Select * from SOTORDR2 where rownum < 1"
        SOTORDR2_TEMP = ASCMAIN1.Temp_Table(SOTORDR2_TEMPsql)
        Call Create_TDA(dst.Tables.Add, SOTORDR2_TEMP, "*", 0, True)

        'Summary by PO
        SOTORDRPsql = "Select SOTORDR1.CUST_CODE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CUST_PO, " & vbCr _
        & "SOTORDR1.ORDR_YYYYPP_UPDATED, SOTORDR1.ORDR_CANCEL_DATE," & vbCr _
        & "SUM(SOTORDR2.ORDR_QTY) ORDR_QTY, SUM(SOTORDR2.ORDR_QTY*SOTORDR2.ORDR_UNIT_PRICE) EXT_PRICE" & vbCr _
        & " from " & SOTORDR1_TEMP & " SOTORDR1, " & SOTORDR2_TEMP & " SOTORDR2" & vbCr _
        & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCr _
        & " group by SOTORDR1.CUST_CODE,SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CUST_PO, " & vbCr _
        & "          SOTORDR1.ORDR_YYYYPP_UPDATED, SOTORDR1.ORDR_CANCEL_DATE" & vbCr _
        & " order by SOTORDR1.CUST_CODE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CUST_PO, " & vbCr _
        & "          SOTORDR1.ORDR_YYYYPP_UPDATED, SOTORDR1.ORDR_CANCEL_DATE"
        Call Create_TDA(dst.Tables.Add, "SOTORDRP", SOTORDRPsql, 0, False)
        grdSOTORDRP.DataSource = dst.Tables("SOTORDRP")
        'grdSOTORDRP.DisplayLayout.Bands(0).SortedColumns.Clear()
        'grdSOTORDRP.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False)
        Create_Summary(grdSOTORDRP, "ORDR_QTY")
        Create_Summary(grdSOTORDRP, "EXT_PRICE")

        'Summary by PO
        SOTORDRIsql = "Select SOTORDR1.CUST_CODE, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CUST_PO,  " & vbCr _
        & "SOTORDR1.ORDR_YYYYPP_UPDATED, SOTORDR1.ORDR_CANCEL_DATE, " & vbCr _
        & "SOTORDR2.ITEM_CODE,SOTORDR2.ITEM_DESC, SUM(SOTORDR2.ORDR_QTY) ORDR_QTY, " & vbCr _
        & "SUM(SOTORDR2.ORDR_QTY*SOTORDR2.ORDR_UNIT_PRICE) EXT_PRICE" & vbCr _
        & " from " & SOTORDR1_TEMP & " SOTORDR1, " & SOTORDR2_TEMP & " SOTORDR2" & vbCr _
        & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCr _
        & " group by SOTORDR1.CUST_CODE,SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CUST_PO, " & vbCr _
        & "          SOTORDR1.ORDR_YYYYPP_UPDATED, SOTORDR1.ORDR_CANCEL_DATE," & vbCr _
        & "          SOTORDR2.ITEM_CODE,SOTORDR2.ITEM_DESC " & vbCr _
        & " order by SOTORDR1.ORDR_YYYYPP_UPDATED, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.CUST_CODE, " & vbCr _
        & "          SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_CANCEL_DATE," & vbCr _
        & "          SOTORDR2.ITEM_CODE,SOTORDR2.ITEM_DESC "
        Call Create_TDA(dst.Tables.Add, "SOTORDRI", SOTORDRIsql, 0, False)
        grdSOTORDRI.DataSource = dst.Tables("SOTORDRI")
        'grdSOTORDRP.DisplayLayout.Bands(0).SortedColumns.Clear()
        'grdSOTORDRP.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False)
        Create_Summary(grdSOTORDRI, "ORDR_QTY")
        Create_Summary(grdSOTORDRI, "EXT_PRICE")

        Create_TDA(dst.Tables.Add, "GLTPARM2", "*", 0, False)

        Fill_Records("GLTPARM2")

        ASCMAIN1.sql = "Select * from SOTINVH1 where rownum < 1"
        Call Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, True)

        ASCMAIN1.sql = "Select * from SOTINVH2 where rownum < 1"
        Call Create_TDA(dst.Tables.Add, "SOTINVH2", "**", 0, True)

        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")


    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

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
                Call Setup_ASTPCTL1()

                Call Cleanup_Archive("CUST")
                Call Cleanup_Archive("SHIP")
                Call Cleanup_Archive("ITEM")

                Call Inbound_ftp()
                Call Process_EDI()

                Call Setup_RSTFILE1()
                Call Load_Items()
                Call Load_Customers()
                Call Load_Orders()
                Call Load_Invoices()
                Call Load_Credits()

                Call Load_Record()
                Call Mode_Settings(True)

                'Case "Import Raw EDI Files"
                '    Call Mode_Settings(True)
                '    Call Import_Raw_EDI()
                '    Call Mode_Settings(False)

                'Case "Load 852 Data", "Retract 852 Data", "Restore Deleted"
                '    Call Mode_Settings(True)
                '    Call Load_852_Data()
                '    Call Mode_Settings(False)

            Case "Cancel"
                Call Mode_Settings(False)

            Case "Update"
                Call Update_Record()
                Call Mode_Settings(False)

            Case "Print"
                Call Print_Report()


        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
            End With

        End If

        'Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'UltraTabControl1.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
            'Setup_tabMain()
        End If

    End Sub

    Sub Clear_Record()
        'Select_tabMain()

        dst.EnforceConstraints = False
        dst.Tables("RSTFILE1").Rows.Clear()
        dst.Tables("ICTITEM1_IN").Rows.Clear()
        dst.Tables("ICTITEM1_PEND").Rows.Clear()
        dst.Tables("ARTCUST1_IN").Rows.Clear()
        dst.Tables("ARTCUST1_PEND").Rows.Clear()
        dst.Tables("SOTORDR1_IN").Rows.Clear()
        dst.Tables("SOTORDR2_IN").Rows.Clear()
        dst.Tables("SOTINVH1_IN").Rows.Clear()
        dst.Tables("SOTINVH2_IN").Rows.Clear()
        dst.Tables("SOTINVH1_CR").Rows.Clear()
        dst.Tables("SOTINVH2_CR").Rows.Clear()
        dst.Tables("SOTORDRP").Rows.Clear()
        dst.Tables(ICTITEM1_TEMP).Rows.Clear()
        dst.Tables(ICTSTAT2_TEMP).Rows.Clear()
        dst.Tables(ARTCUST1_TEMP).Rows.Clear()
        dst.Tables(SOTORDR1_TEMP).Rows.Clear()
        dst.Tables(SOTORDR2_TEMP).Rows.Clear()
        dst.Tables("SOTINVH1").Rows.Clear()
        dst.Tables("SOTINVH2").Rows.Clear()

        dst.EnforceConstraints = True

    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading")

        Call Save_Header_Fields(UltraGroupBox1)
        dst.EnforceConstraints = False

        'Fill_Records("RSTFILE1")
        'Fill_Records("ICTITEM1_IN")

        'Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        'tlb_sbt = DirectCast(tlb.Tools("Show Summary Columns Only"), UltraWinToolbars.StateButtonTool)
        'tlb_sbt.Checked = True
        'the above three lines cause ERROR in tlb_ToolClick OBJECT NOT SET TO AN INSTANCE
        'Me.Show_Detail_or_Summary()

        dst.EnforceConstraints = True
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()

        'Add new items
        ASCMAIN1.sql = "Insert into ICTITEM1" & vbCr _
        & "(ITEM_CODE, ITEM_STATUS, ITEM_DESC, ITEM_DESC2, ITEM_UOM, " & vbCr _
        & " PROD_CODE, COLLECTION_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD, " & vbCr _
        & " ITEM_UPC_CODE, ITEM_EAN_CODE, DEPT_CODE, " & vbCr _
        & " ITEM_CLASS_CODE, INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE,ITEM_CATGY_CODE) " & vbCr _
        & " Select ITEM_CODE, 'A', ITEM_DESC, ITEM_DESC2, ITEM_UOM, " & vbCr _
        & " PROD_CODE, COLLECTION_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD, " & vbCr _
        & " ITEM_UPC_CODE, ITEM_EAN_CODE, (Select MIN(DEPT_CODE) from ICTDEPT1)," & vbCr _
        & " (Select MIN(ITEM_CLASS_CODE) from ICTCLAS1), '" & ASCMAIN1.USER_ID & "','" & ASCMAIN1.USER_ID & "'," & vbCr _
        & " SYSDATE , SYSDATE, 'E' " & vbCr _
        & " from " & ICTITEM1_TEMP & " where ITEM_STATUS = 'A' "
        ASCDATA1.ExecuteSQL()

        'Update existing items
        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select * from " & ICTITEM1_TEMP & " where ITEM_STATUS = 'C';" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update ICTITEM1 " _
        & "    set  ITEM_DESC=R1.ITEM_DESC, ITEM_DESC2=R1.ITEM_DESC2, ITEM_UOM=R1.ITEM_UOM," _
        & "         PROD_CODE=R1.PROD_CODE, ITEM_RETAIL_PRICE=R1.ITEM_RETAIL_PRICE, ITEM_COST_STD=R1.ITEM_COST_STD," _
        & "         DEPT_CODE=(Select MIN(DEPT_CODE) from ICTDEPT1), " _
        & "         ITEM_CATGY_CODE=(Select MIN(ITEM_CATGY_CODE) from ICTCATG1), " _
        & "         ITEM_EAN_CODE=R1.ITEM_EAN_CODE,ITEM_UPC_CODE=R1.ITEM_UPC_CODE, " _
        & "         COLLECTION_CODE=R1.COLLECTION_CODE, " _
        & "         LAST_OPER = '" & ASCMAIN1.USER_ID & "'," _
        & "         LAST_DATE = SYSDATE" _
        & "    where ITEM_CODE = R1.ITEM_CODE; " _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        'set ITEM_SNU_CODE
        ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_SNU_CODE = 'S' where nvl(ITEM_RETAIL_PRICE,0) <> '0'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_SNU_CODE = 'N' where nvl(ITEM_RETAIL_PRICE,0) = '0'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update ICTITEM1 Set ITEM_SNU_CODE = 'U' where nvl(ITEM_RETAIL_PRICE,0) = '0' and PROD_CODE='COMP'"
        ASCDATA1.ExecuteSQL()

        'Add new customers
        ASCMAIN1.sql = "Insert into ARTCUST1" & vbCr _
        & "(CUST_CODE, CUST_STATUS, CUST_NAME, CUST_ADDR1, CUST_ADDR2, " & vbCr _
        & " CUST_CITY, CUST_STATE, CUST_ZIP_CODE, CUST_COUNTRY, CUST_CONTACT, " & vbCr _
        & " CUST_PHONE, CUST_FAX, SREP_CODE, PRICE_CLASS_CODE, TERM_CODE, FRT_TERMS,  " & vbCr _
        & " CURR_CODE, TRADE_CLASS_CODE, CUST_CLASS_CODE, INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE) " & vbCr _
        & " Select CUST_CODE, 'A', CUST_NAME, CUST_ADDR1, CUST_ADDR2, " & vbCr _
        & " CUST_CITY, CUST_STATE, CUST_ZIP_CODE, CUST_COUNTRY, CUST_CONTACT, " & vbCr _
        & " CUST_PHONE, CUST_FAX, SREP_CODE, PRICE_CLASS_CODE, TERM_CODE, FRT_TERMS, " & vbCr _
        & " CURR_CODE, TRADE_CLASS_CODE, CUST_CLASS_CODE,'" & ASCMAIN1.USER_ID & "','" & ASCMAIN1.USER_ID & "'," & vbCr _
        & " SYSDATE , SYSDATE " & vbCr _
        & " from " & ARTCUST1_TEMP & " where CUST_STATUS = 'A' "
        ASCDATA1.ExecuteSQL()

        'Setup Store 000000 for customers without Store
        ASCMAIN1.sql = "Insert into ARTCUST2" & vbCr _
        & "(CUST_CODE, CUST_STORE_STATUS, CUST_STORE_NO, CUST_STORE_NAME, CUST_STORE_ADDR1," & vbCr _
        & " CUST_STORE_ADDR2, CUST_STORE_ADDR3, CUST_STORE_CITY, CUST_STORE_STATE, CUST_STORE_ZIP_CODE, " & vbCr _
        & " SELL_CODE, INIT_OPER, LAST_OPER, INIT_DATE, LAST_DATE) " & vbCr _
        & " Select CUST_CODE, 'A', '000000',CUST_NAME, CUST_ADDR1,   " & vbCr _
        & " CUST_ADDR2, CUST_ADDR3, CUST_CITY, CUST_STATE, CUST_ZIP_CODE, " & vbCr _
        & " '000', '" & ASCMAIN1.USER_ID & "','" & ASCMAIN1.USER_ID & "'," & vbCr _
        & " SYSDATE , SYSDATE " & vbCr _
        & " from ARTCUST1 where CUST_CODE in (" & vbCr _
        & "  Select CUST_CODE from ARTCUST1 " & vbCr _
        & "  minus " & vbCr _
        & "  Select Distinct CUST_CODE from ARTCUST2)"
        ASCDATA1.ExecuteSQL()

        'Update existing customers
        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select * from " & ARTCUST1_TEMP & " where CUST_STATUS = 'C';" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update ARTCUST1 " _
        & "    set  CUST_NAME=R1.CUST_NAME, CUST_ADDR1=R1.CUST_ADDR1, CUST_ADDR2=R1.CUST_ADDR2," _
        & "         CUST_CITY=R1.CUST_CITY, CUST_STATE=R1.CUST_STATE, CUST_ZIP_CODE=R1.CUST_ZIP_CODE," _
        & "         CUST_COUNTRY=R1.CUST_COUNTRY, CUST_CONTACT=R1.CUST_CONTACT," _
        & "         CUST_PHONE=R1.CUST_PHONE, CUST_FAX=R1.CUST_FAX, SREP_CODE=R1.SREP_CODE, " _
        & "         PRICE_CLASS_CODE=R1.PRICE_CLASS_CODE, TERM_CODE=R1.TERM_CODE, " _
        & "         FRT_TERMS=R1.FRT_TERMS, CURR_CODE=R1.CURR_CODE, " _
        & "         LAST_OPER = '" & ASCMAIN1.USER_ID & "'," _
        & "         LAST_DATE = SYSDATE" _
        & "    where CUST_CODE = R1.CUST_CODE; " _
        & "   Update ARTCUST2 " _
        & "    set  CUST_STORE_NAME=R1.CUST_NAME, CUST_STORE_ADDR1=R1.CUST_ADDR1, " _
        & "         CUST_STORE_ADDR2=R1.CUST_ADDR2, CUST_STORE_CITY=R1.CUST_CITY," _
        & "         CUST_STORE_STATE=R1.CUST_STATE, CUST_STORE_ZIP_CODE=R1.CUST_ZIP_CODE," _
        & "         LAST_OPER = '" & ASCMAIN1.USER_ID & "'," _
        & "         LAST_DATE = SYSDATE" _
        & "    where CUST_CODE = R1.CUST_CODE and CUST_STORE_NO = '000000'; " _
       & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        Call Update_Record_TDA("SOTINVH1")
        Call Update_Record_TDA("SOTINVH2")

        'Update ORDR_YYYYPP_UPDATED on SOTINVH1 AND SOTINVH2
        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select * from GLTPARM2 order by OPS_YYYYPP;" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update SOTINVH1" _
        & "    set  ORDR_YYYYPP_UPDATED = R1.OPS_YYYYPP" _
        & "    where INV_DATE <= R1.PRD_END_DATE and ORDR_YYYYPP_UPDATED = 'XXXXXX';" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select SOTINVH2.INV_NO INV_NO, SOTINVH1.ORDR_YYYYPP_UPDATED ORDR_YYYYPP_UPDATED " _
        & "   from SOTINVH2, SOTINVH1 " _
        & "   where SOTINVH2.INV_NO=SOTINVH1.INV_NO and SOTINVH2.ORDR_YYYYPP_UPDATED='XXXXXX';" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update SOTINVH2" _
        & "    set  ORDR_YYYYPP_UPDATED=R1.ORDR_YYYYPP_UPDATED" _
        & "    where INV_NO = R1.INV_NO;" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        'Update OPS_YYYYWW on SOTINVH1 AND SOTINVH2
        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select * from GLTPARM3 order by YYYYWW;" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update SOTINVH1" _
        & "    set  OPS_YYYYWW  = R1.YYYYWW" _
        & "    where INV_DATE <= R1.WEEK_END_DATE and OPS_YYYYWW IS NULL;" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select SOTINVH2.INV_NO INV_NO, SOTINVH1.OPS_YYYYWW OPS_YYYYWW " _
        & "   from SOTINVH2, SOTINVH1 " _
        & "   where SOTINVH2.INV_NO=SOTINVH1.INV_NO and SOTINVH2.OPS_YYYYWW IS NULL;" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update SOTINVH2" _
        & "    set  OPS_YYYYWW=R1.OPS_YYYYWW" _
        & "    where INV_NO = R1.INV_NO;" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        'Rebuild SATSSUM tables for the last 3 months plus 1 month
        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYPP between " _
        & " '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -3) & "' and '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, +1) & "';" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & " Begin SAPSSUMX(R1.OPS_YYYYPP); End;" _
        & "  End Loop;" _
        & " End; " _
        & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from ICTSTAT2"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into ICTSTAT2 Select * from " & ICTSTAT2_TEMP
        ASCDATA1.ExecuteSQL()

        Call CommitTrans("Update Complete")

        'Archive Invoice files processed
        For Each downloadedFileName As String In My.Computer.FileSystem.GetFiles(downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "Invoices*.txt")
            ASCMAIN1.Progress("Archiving File " & downloadedFileName)
            My.Computer.FileSystem.CopyFile(downloadedFileName, downloadFileLocationArchive & My.Computer.FileSystem.GetFileInfo(downloadedFileName).Name, True)
            My.Computer.FileSystem.DeleteFile(downloadedFileName)
        Next
        For Each downloadedFileName As String In My.Computer.FileSystem.GetFiles(downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "Credits*.txt")
            ASCMAIN1.Progress("Archiving File " & downloadedFileName)
            My.Computer.FileSystem.CopyFile(downloadedFileName, downloadFileLocationArchive & My.Computer.FileSystem.GetFileInfo(downloadedFileName).Name, True)
            My.Computer.FileSystem.DeleteFile(downloadedFileName)
        Next

        ASCMAIN1.Progress("")

    End Sub

    Sub Print_Report()
        Call Print_Report_Begin()
        Dim SUBT As String = ""

        Dim RecordSelectionFormula As String = ""
        Generate_Report("RSRINBDS", "Stock Status", SUBT, RecordSelectionFormula)
        Generate_Report("RSRINBDP", "Orders by Customer PO", SUBT, RecordSelectionFormula)
        Generate_Report("RSRINBDI", "Orders by Customer PO Item", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(grdEDT852T1, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdICTITEM1_IN, "SSSS", "Show Filter", "Show GroupBox", "Show Unavailable Only", "Show Summary Columns Only")
        Call Load_Popup_Menu(grdICTITEM1_PEND, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdARTCUST1_IN, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdARTCUST1_PEND, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdSOTORDR1_IN, "SS", "Show Filter", "Show GroupBox")
        'Call Load_Popup_Menu(grdSOTORDR2_IN, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdSOTORDRP, "SS", "Show Filter", "Show GroupBox")
        Call Load_Popup_Menu(grdSOTORDRI, "SS", "Show Filter", "Show GroupBox")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Dim tlb_sbt As UltraWinToolbars.StateButtonTool
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden

            Select Case e.SourceControl.Name
                Case "grdEDT852T1"
                Case "grdICTITEM1_IN"
                Case "grdICTITEM1_PEND"
                Case "grdARTCUST1_IN"
                Case "grdARTCUST1_PEND"
                Case "grdSOTORDR1_IN"
                Case "grdSOTORDRP"
                Case "grdSOTORDRI"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)
            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
            Case "Show Unavailable Only"
                Call Show_Unavailable_Only()
            Case "Show Summary Columns Only"
                Call Show_Detail_or_Summary()


        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.dte_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "DTE0", "DTE1"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub

        If tabMain.SelectedTab Is Nothing Then Exit Sub

        'With UltraExplorerBar1.Groups("Batch Control")
        '    .Items("Import Raw EDI Files").Visible = (tabMain.SelectedTab.Key = "Raw EDI Files")
        '    .Items("Load 852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "0")
        '    .Items("Retract 852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "1")
        '    .Items("Restore Deleted").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "D")
        'End With

        With UltraExplorerBar1

            '.Groups("Screen Control").Items("Print").Visible = _
            'tabMain.SelectedTab.Key = "Orders"

            '.Groups("Screen Control").Visible = (tabMain.SelectedTab.Key = "Inbound Files")

            'Select Case tabMain.SelectedTab.Key
            '    Case "Inbound Files"
            '        grdEDTJRNL1.Dock = DockStyle.None
            '        grdEDTJRNL1.Parent = tabRaw.Tabs("Raw Data").TabPage
            '        grdEDTJRNL1.Dock = DockStyle.Fill

            '        grpRawEDI.Dock = DockStyle.None
            '        grpRawEDI.Parent = tabRaw.Tabs("Raw EDI").TabPage
            '        grpRawEDI.Dock = DockStyle.Fill

            '        grdEDT852T1.Dock = DockStyle.None
            '        grdEDT852T1.Parent = tabRaw.Tabs("Documents").TabPage
            '        grdEDT852T1.Dock = DockStyle.Fill

            'End Select

            '.Groups("Raw EDI Files").Visible = (tabMain.SelectedTab.Key = "Raw EDI Files")
            '.Groups("852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data")
            '.Groups("Batch Control").Visible = (tabMain.SelectedTab.Key <> "Manual Entry")

        End With
    End Sub

    Sub Import_Raw_EDI()

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Processing EDI Raw Files")

        'Dim EDCIBND1 As New TAC.EDCIBND1()

        'dst.Tables("EDTFILE1").Rows.Clear()

        'Dim file_counter As Integer = 0

        'Dim wildcard As String = "*.edi" ' "MAIL_IN.TXT"

        'Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""
        'For Each FILE As String In My.Computer.FileSystem.GetFiles _
        '(ED_PARM_RAW_INBOUND, FileIO.SearchOption.SearchAllSubDirectories, wildcard)
        '    Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
        '    Dim rowEDTFILE1 As DataRow = dst.Tables("EDTFILE1").NewRow
        '    Dim FILENAME As String = Mid(FILEINFO.FullName, ED_PARM_RAW_INBOUND.Length + 2)

        '    Dim row As DataRow = LookUp("EDTFILE1", FILENAME)
        '    If row Is Nothing Then
        '        rowEDTFILE1.Item("EDI_FILENAME") = FILENAME
        '        rowEDTFILE1.Item("EDI_FILESIZE") = FILEINFO.Length
        '        rowEDTFILE1.Item("EDI_DATETIME") = FILEINFO.LastWriteTime
        '        dst.Tables("EDTFILE1").Rows.Add(rowEDTFILE1)

        '        file_counter += 1

        '        ASCMAIN1.Progress("Processing " & FILENAME)
        '        Dim EDI_JRNL_NOs As List(Of String)
        '        EDI_JRNL_NOs = EDCIBND1.Process_File( _
        '        ED_PARM_RAW_INBOUND, FILENAME, _
        '        FILEINFO, "852", _
        '        ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE") & "", _
        '        "", False)

        '        If EDI_JRNL_NOs.Count > 0 Then
        '            Dim rowEDTJRNL1 = LookUp("EDTJRNL1", EDI_JRNL_NOs(0))
        '            rowEDTFILE1.Item("EDI_JRNL_NO") = rowEDTJRNL1.item("EDI_JRNL_NO")
        '            rowEDTFILE1.Item("EDI_SENDER_QUAL") = rowEDTJRNL1.item("EDI_SENDER_QUAL")
        '            rowEDTFILE1.Item("EDI_SENDER_ID") = rowEDTJRNL1.item("EDI_SENDER_ID")
        '            rowEDTFILE1.Item("EDI_ISA_CTL_NO") = rowEDTJRNL1.item("EDI_ISA_CTL_NO")
        '            rowEDTFILE1.Item("EDI_ISA_CTL_DATE") = rowEDTJRNL1.item("EDI_ISA_CTL_DATE")
        '            ASCMAIN1.sql = "Select Count (*) from EDTJRNL3 where EDI_JRNL_NO = '" & EDI_JRNL_NOs(0) & "'"
        '            Dim DOC_EDI As Integer = ASCDATA1.GetDataValue
        '            ASCMAIN1.sql = "Select Count (*) from EDTJRNL3 where EDI_JRNL_NO = '" & EDI_JRNL_NOs(0) & "' and EDI_DOC_NO = '852'"
        '            Dim DOC_852 As Integer = ASCDATA1.GetDataValue
        '            rowEDTFILE1.Item("DOC_EDI") = DOC_EDI
        '            rowEDTFILE1.Item("DOC_852") = DOC_852
        '            rowEDTFILE1.Item("NOTES") = ""
        '        End If

        '        ASCMAIN1.Progress("Now Processing EDI Raw Files")
        '    End If

        'Next

        'Update_Record_TDA("EDTFILE1")

        'Me.Cursor = Cursors.Default
        'ASCMAIN1.Progress("")

        'MsgBox(CStr(file_counter) & " Files have been Imported", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Inbound_ftp()
        Dim HOST As String = ROWs("RSTPARM1").Item("RS_PARM_HOST_IP") & ""
        Dim USERNAME As String = ROWs("RSTPARM1").Item("RS_PARM_HOST_USER") & ""
        Dim PASSWORD As String = ROWs("RSTPARM1").Item("RS_PARM_HOST_PASS") & ""
        Dim REMOTEDIR As String = ROWs("RSTPARM1").Item("RS_PARM_HOST_DIR") & ""


        ASCMAIN1.Progress("Connect to FTP Server", String.Empty)
        Dim fileName As String = String.Empty
        Dim iFiles As Integer = 0

        Try
            If Not Ftp1.Connected Then
                Ftp1.RemoteHost = HOST
                Ftp1.User = USERNAME
                Ftp1.Password = PASSWORD
                Ftp1.Logon()
            End If

            If Not Ftp1.Connected Then
                MessageBox.Show("Could not connect to remote host to download files.", "FTP Files", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

            _downloadedFiles = New List(Of String)

            _remoteDirectoryFileList = New List(Of String)
            Try
                Ftp1.RemoteFile = String.Empty
                Ftp1.RemotePath = remoteDir
                Ftp1.ListDirectory()
            Catch ex As Exception
                ' just in case, no need to bomb out
                MessageBox.Show("Error: " & ex.Message & " - accessing " & remoteDir, "FTP", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            Ftp1.Overwrite = True

            iFiles = 1
            ASCMAIN1.Progress("-", iFiles.ToString & " / " & _remoteDirectoryFileList.Count)

            Dim delete_files As Boolean = True
            If ASCMAIN1.Running_in_VS Then
                If MsgBox("Delete files after pickup?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    delete_files = False
                End If
            End If
            ' Get the Data Files
            For Each remoteDirItem As String In _remoteDirectoryFileList
                fileName = remoteDirItem.Trim
                If fileName.EndsWith(_EDIFileExt) Or _
                    fileName.EndsWith(_TextFileExt) Then

                    ASCMAIN1.Progress(fileName, iFiles.ToString & " of " & _remoteDirectoryFileList.Count)
                    ' Download file
                    Ftp1.RemotePath = REMOTEDIR
                    Ftp1.RemoteFile = fileName
                    Ftp1.LocalFile = downloadFileLocation & My.Computer.FileSystem.GetName(fileName)
                    Ftp1.Download()
                    Ftp1.DoEvents()

                    If delete_files = True Then
                        Ftp1.DeleteFile(fileName)
                        Ftp1.DoEvents()
                    End If

                    ' keep a copy of the downloaded file's name
                    _downloadedFiles.Add(fileName)
                    downloadFiles = True

                    iFiles += 1
                End If
            Next

            Ftp1.Logoff()
            downloadFiles = True

        Catch ex As Exception
            MessageBox.Show(ex.Message, "FTP Files", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Try
                If Ftp1.Connected Then Ftp1.Logoff()
            Catch ex1 As Exception
                ' nothing
            End Try
            'Return False
        End Try

        ASCMAIN1.Progress(String.Empty, String.Empty)

    End Sub

    Sub Setup_EDTFILE1_Details()

        'Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""

        'Dim FILENAME As String = ED_PARM_RAW_INBOUND & "\" & grdRSTFILE1.ActiveRow.Cells("EDI_FILENAME").Text

        'Me.Cursor = Cursors.WaitCursor
        'ASCMAIN1.Progress("Now Loading Data for file")

        'Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        'grpRawEDI.Text = FI.Name & " " & Format$(FI.LastWriteTime, "MM/dd/yy HH:mm")
        'txtRawEDI.Text = ""
        'Using SR As New System.IO.StreamReader(FILENAME)
        '    Dim RAW As String = SR.ReadToEnd
        '    txtRawEDI.Text = Replace(RAW, Mid(RAW, 106, 1), vbCrLf)
        'End Using

        'Dim EDI_JRNL_NO As String = grdRSTFILE1.ActiveRow.Cells("EDI_JRNL_NO").Text
        'Dim sql As String = ""

        'dst.EnforceConstraints = False
        'sql = "Select EDT852T1.*, ARTCUST1.CUST_NAME, '0' SELECTED " _
        '& " from EDT852T1,ARTCUST1 " _
        '& " where EDT852T1.EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO FROM EDTJRNL1,EDTJRNL3 where EDTJRNL1.EDI_JRNL_NO = EDTJRNL3.EDI_JRNL_NO and EDTJRNL1.EDI_JRNL_NO = '" & EDI_JRNL_NO & "')" _
        '& " and ARTCUST1.CUST_CODE (+) = EDT852T1.CUST_CODE"
        'Fill_Records("EDT852T1", "", True, sql)

        'Fill_Records("EDTJRNL1", EDI_JRNL_NO)

        'sql = "Select EDTJRNL2.* from EDTJRNL2 " _
        '& " where EDTJRNL2.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
        'Fill_Records("EDTJRNL2", "", True, sql)

        'sql = "Select EDTJRNL3.* from EDTJRNL3 " _
        '& " where EDTJRNL3.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
        'Fill_Records("EDTJRNL3", "", True, sql)

        'dst.EnforceConstraints = True

        'Sort_grdColumns(grdEDT852T1, "EDI_DOC_SEQ_NO")

        'If grdEDTJRNL1.Rows.Count > 0 Then
        '    grdEDTJRNL1.Rows(0).ExpandAll()
        'End If

        'Me.Cursor = Cursors.Default
        'ASCMAIN1.Progress("")
    End Sub

    Sub Select_tabMain()
        'tabMain.VisibleTab
        'tabMain.Tabs(optMode.Value).visible = True
        'tabMain.SelectedTab = tabMain.Tabs(optMode.Value)
        tabMain.SelectedTab.Visible = True
        For Each TAB As UltraWinTabControl.UltraTab In tabMain.Tabs
            If Not TAB.Selected Then
                TAB.Visible = False
            End If
        Next
    End Sub

    Sub Setup_RSTFILE1()
        ASCMAIN1.Progress("Processing Files", "")

        Dim rowRSTFILE1 As DataRow = Nothing

        Get_Most_Recent("CUST")
        Get_Most_Recent("SHIP")
        Get_Most_Recent("ITEM")
        Get_Most_Recent("ORDERS")

        For Each downloadedFileName As String In My.Computer.FileSystem.GetFiles(downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
            rowRSTFILE1 = dst.Tables("RSTFILE1").NewRow

            shortFileName = My.Computer.FileSystem.GetFileInfo(downloadedFileName).Name
            rowRSTFILE1.Item("FILENAME") = shortFileName
            rowRSTFILE1.Item("FILESIZE") = My.Computer.FileSystem.GetFileInfo(downloadedFileName).Length
            'parse filename for date time
            If My.Computer.FileSystem.GetFileInfo(downloadedFileName).Extension = ".txt" _
            And (shortFileName.Substring(0, 4) = "Cust" Or _
                  shortFileName.Substring(0, 4) = "Ship" Or _
                  shortFileName.Substring(0, 4) = "Item") Then
                rowRSTFILE1.Item("FILEDATETIME") = shortFileName.Substring(4, 2) & "/" & _
                    shortFileName.Substring(6, 2) & "/" & shortFileName.Substring(8, 4) & " " & _
                    shortFileName.Substring(12, 2) & ":" & shortFileName.Substring(14, 2) & ":" & _
                    shortFileName.Substring(16, 2)
            Else
                rowRSTFILE1.Item("FILEDATETIME") = My.Computer.FileSystem.GetFileInfo(downloadedFileName).LastWriteTime

            End If
            dst.Tables("RSTFILE1").Rows.Add(rowRSTFILE1)
        Next
        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub Get_Most_Recent(ByVal FILETYPE As String)
        Dim master_filename As String = ""
        Dim current_filedate As String = ""
        Dim current_filename As String = ""
        Dim last_filename As String = ""
        Dim startpos As Integer = Len(FILETYPE)

        current_filename = ""
        current_filedate = ""
        master_filename = ""
        last_filename = ""
        'keep most recent item file - archive others
        For Each downloadedFileName As String In My.Computer.FileSystem.GetFiles(downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, FILETYPE & "*.txt")
            current_filename = My.Computer.FileSystem.GetFileInfo(downloadedFileName).Name
            current_filedate = current_filename.Substring(startpos + 4, 4) & current_filename.Substring(startpos, 4)
            ASCMAIN1.Progress("Processing " & current_filename)

            If master_filename.Length = 0 Then
                master_filename = current_filename
            Else
                If current_filedate > master_filename.Substring(startpos + 4, 4) & master_filename.Substring(startpos, 4) Then
                    last_filename = master_filename
                    master_filename = current_filename
                    'My.Computer.FileSystem.moveFile(downloadFileLocation & last_filename, downloadFileLocationArchive.Substring(0, _downloadFileLocationArchive.Length - 1))
                    My.Computer.FileSystem.CopyFile(downloadFileLocation & last_filename, downloadFileLocationArchive & last_filename, True)
                    My.Computer.FileSystem.DeleteFile(downloadFileLocation & last_filename)
                End If
            End If
        Next
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub Cleanup_Archive(ByVal FILETYPE As String)
        'keep most recent item file - archive others
        For Each downloadedFileName As String In My.Computer.FileSystem.GetFiles(downloadFileLocationArchive, FileIO.SearchOption.SearchTopLevelOnly, FILETYPE & "*.txt")
            ASCMAIN1.Progress("Processing " & downloadedFileName)
            If DateAdd("d", 30, My.Computer.FileSystem.GetFileInfo(downloadedFileName).CreationTime) < Today Then
                My.Computer.FileSystem.DeleteFile(downloadedFileName)
            End If
        Next
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub Ftp1_OnDirList(ByVal sender As System.Object, ByVal e As nsoftware.IPWorks.FtpDirListEventArgs) Handles Ftp1.OnDirList
        If Not e.IsDir Then
            _remoteDirectoryFileList.Add(e.FileName)
        End If
    End Sub

    Private Sub Process_EDI()
        ASCMAIN1.Progress("Processing EDI", "")
        Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
        (downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "*" & _EDIFileExt)
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            Dim rowRSTFILE1 As DataRow = dst.Tables("RSTFILE1").NewRow
            rowRSTFILE1 = dst.Tables("RSTFILE1").NewRow

            shortFileName = My.Computer.FileSystem.GetFileInfo(FILE).Name
            rowRSTFILE1.Item("FILENAME") = shortFileName
            rowRSTFILE1.Item("FILESIZE") = My.Computer.FileSystem.GetFileInfo(FILE).Length
            rowRSTFILE1.Item("FILEDATETIME") = My.Computer.FileSystem.GetFileInfo(FILE).LastWriteTime

            dst.Tables("RSTFILE1").Rows.Add(rowRSTFILE1)
            My.Computer.FileSystem.CopyFile(downloadFileLocation & shortFileName, downloadFileLocationArchive & shortFileName, True)
            My.Computer.FileSystem.CopyFile(downloadFileLocation & shortFileName, ED_PARM_RAW_INBOUND & "\" & shortFileName, True)
            My.Computer.FileSystem.DeleteFile(downloadFileLocation & shortFileName)
        Next
        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub Setup_ASTPCTL1()
        ASCMAIN1.sql = "UPDATE ASTPCTL1 SET CURR_WEEK = " _
        & "(SELECT SUBSTR(YYYYWW,5,2) FROM GLTPARM3 WHERE WEEK_END_DATE = " _
        & "(SELECT MIN(WEEK_END_DATE) FROM GLTPARM3 " _
        & " WHERE WEEK_END_DATE > SYSDATE))"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "UPDATE ASTPCTL1 SET CURR_PERIOD = " _
        & "(SELECT SUBSTR(YYYYPP,5,2) FROM GLTPARM3 WHERE WEEK_END_DATE =  " _
        & "(SELECT MIN(WEEK_END_DATE) FROM GLTPARM3 " _
        & " WHERE WEEK_END_DATE > SYSDATE))"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "UPDATE ASTPCTL1 SET CURR_YEAR = " _
        & "(SELECT SUBSTR(YYYYPP,1,4) FROM GLTPARM3 WHERE WEEK_END_DATE =  " _
        & "(SELECT MIN(WEEK_END_DATE) FROM GLTPARM3 " _
        & " WHERE WEEK_END_DATE > SYSDATE))"
        ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub Load_Items()
        ASCMAIN1.Progress("Loading Items", "")

        Dim item_data As String = ""
        Dim TEXT_FILENAME As String = ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
            (downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "item*.txt")
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            TEXT_FILENAME = My.Computer.FileSystem.GetFileInfo(FILE).FullName
        Next

        Using sr As New System.IO.StreamReader(TEXT_FILENAME)
            item_data = sr.ReadToEnd
        End Using

        Dim rows_imported As Integer = 0

        dst.Tables("ICTITEM1_IN").Rows.Clear()
        dst.Tables(ICTITEM1_TEMP).Rows.Clear()
        dst.Tables(ICTSTAT2_TEMP).Rows.Clear()

        For Each line As String In Split(item_data, vbCrLf)
            If line.Length > 1 Then
                Dim rowICTITEM1_IN As DataRow = dst.Tables("ICTITEM1_IN").NewRow
                rowICTITEM1_IN.Item("ITEM_CODE") = Trim(line.Substring(0, 20))
                rowICTITEM1_IN.Item("ITEM_DESC") = Trim(line.Substring(20, 30))
                rowICTITEM1_IN.Item("ITEM_DESC2") = Trim(line.Substring(50, 30))
                rowICTITEM1_IN.Item("ITEM_UOM") = Trim(line.Substring(80, 4))
                rowICTITEM1_IN.Item("AVG_COST") = Trim(line.Substring(108, 8))   'AVG_COST
                rowICTITEM1_IN.Item("UNIT_COST") = Trim(line.Substring(116, 8))   'UNIT_COST
                rowICTITEM1_IN.Item("ITEM_COST_STD") = Trim(line.Substring(124, 8))
                rowICTITEM1_IN.Item("ITEM_RETAIL_PRICE") = Trim(line.Substring(132, 8))
                rowICTITEM1_IN.Item("GEN_PROD_POST_GRP") = Trim(line.Substring(140, 10))
                rowICTITEM1_IN.Item("INV_POST_GRP") = Trim(line.Substring(150, 10))
                rowICTITEM1_IN.Item("ITEM_UPC_EAN") = Trim(line.Substring(160, 13))
                rowICTITEM1_IN.Item("QTY_ON_HAND") = Math.Abs(Val(Trim(line.Substring(173, 10).Replace(",", ""))))
                rowICTITEM1_IN.Item("QTY_ON_SO") = Math.Abs(Val(Trim(line.Substring(183, 10).Replace(",", ""))))
                rowICTITEM1_IN.Item("QTY_ON_PO") = Math.Abs(Val(Trim(line.Substring(193, 10).Replace(",", ""))))
                rowICTITEM1_IN.Item("QTY_AVAIL") = rowICTITEM1_IN.Item("QTY_ON_HAND") - rowICTITEM1_IN.Item("QTY_ON_SO")
                rowICTITEM1_IN.Item("PROD_LINE") = Trim(line.Substring(203, 20))
                dst.Tables("ICTITEM1_IN").Rows.Add(rowICTITEM1_IN)
                rows_imported += 1
            End If
        Next

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & ICTITEM1_TEMP)
        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & ICTSTAT2_TEMP)

        For Each rowICTITEM1_IN As DataRow In dst.Tables("ICTITEM1_IN").Select
            Dim rowICTITEM1_TEMP As DataRow = dst.Tables(ICTITEM1_TEMP).NewRow
            rowICTITEM1_TEMP.Item("ITEM_CODE") = rowICTITEM1_IN.Item("ITEM_CODE")
            rowICTITEM1_TEMP.Item("ITEM_DESC") = rowICTITEM1_IN.Item("ITEM_DESC")
            rowICTITEM1_TEMP.Item("ITEM_DESC2") = rowICTITEM1_IN.Item("ITEM_DESC2")
            rowICTITEM1_TEMP.Item("ITEM_UOM") = rowICTITEM1_IN.Item("ITEM_UOM")
            rowICTITEM1_TEMP.Item("ITEM_RETAIL_PRICE") = rowICTITEM1_IN.Item("ITEM_RETAIL_PRICE")
            rowICTITEM1_TEMP.Item("ITEM_COST_STD") = rowICTITEM1_IN.Item("ITEM_COST_STD")
            rowICTITEM1_TEMP.Item("PROD_CODE") = rowICTITEM1_IN.Item("GEN_PROD_POST_GRP")
            rowICTITEM1_TEMP.Item("COLLECTION_CODE") = rowICTITEM1_IN.Item("PROD_LINE")

            If rowICTITEM1_IN.Item("ITEM_UPC_EAN").ToString.Length = 13 Then
                rowICTITEM1_TEMP.Item("ITEM_EAN_CODE") = rowICTITEM1_IN.Item("ITEM_UPC_EAN")
            ElseIf rowICTITEM1_IN.Item("ITEM_UPC_EAN").ToString.Length = 12 Then
                rowICTITEM1_TEMP.Item("ITEM_UPC_CODE") = rowICTITEM1_IN.Item("ITEM_UPC_EAN")
            End If
            dst.Tables(ICTITEM1_TEMP).Rows.Add(rowICTITEM1_TEMP)

            Dim rowICTSTAT2_TEMP As DataRow = dst.Tables(ICTSTAT2_TEMP).NewRow
            rowICTSTAT2_TEMP.Item("ITEM_CODE") = rowICTITEM1_IN.Item("ITEM_CODE")
            rowICTSTAT2_TEMP.Item("WHSE_CODE") = "PDR"
            rowICTSTAT2_TEMP.Item("WHSE_QTY_ON_HAND") = rowICTITEM1_IN.Item("QTY_ON_HAND")
            rowICTSTAT2_TEMP.Item("WHSE_QTY_OPEN") = rowICTITEM1_IN.Item("QTY_ON_SO")
            rowICTSTAT2_TEMP.Item("WHSE_QTY_ONPO") = rowICTITEM1_IN.Item("QTY_ON_PO")
            dst.Tables(ICTSTAT2_TEMP).Rows.Add(rowICTSTAT2_TEMP)

        Next

        Call Update_Record_TDA(ICTITEM1_TEMP)
        Call Update_Record_TDA(ICTSTAT2_TEMP)

        'New Items
        ASCMAIN1.sql = " Update " & ICTITEM1_TEMP & " set ITEM_STATUS = 'A' WHERE ITEM_CODE IN " & vbCr _
        & " (Select ITEM_CODE from  " & ICTITEM1_TEMP & vbCr _
        & " minus " & vbCr _
        & " Select ITEM_CODE from ICTITEM1)"
        ASCDATA1.ExecuteSQL()

        'Update Items
        ASCMAIN1.sql = " Update " & ICTITEM1_TEMP & " set ITEM_STATUS = 'C' WHERE ITEM_CODE IN (" & vbCr _
        & " Select ITEM_CODE from (" & vbCr _
        & " Select ITEM_CODE, ITEM_DESC, ITEM_DESC2, ITEM_UOM, " & vbCr _
        & " PROD_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD, " & vbCr _
        & " ITEM_UPC_CODE, ITEM_EAN_CODE, COLLECTION_CODE" & vbCr _
        & " from ICTITEM1 " & vbCr _
        & " minus " & vbCr _
        & " Select ITEM_CODE, ITEM_DESC, ITEM_DESC2, ITEM_UOM, " & vbCr _
        & " PROD_CODE, ITEM_RETAIL_PRICE, ITEM_COST_STD, " & vbCr _
        & " ITEM_UPC_CODE, ITEM_EAN_CODE, COLLECTION_CODE " & vbCr _
        & " from " & ICTITEM1_TEMP & "))"
        ASCDATA1.ExecuteSQL()

        Fill_Records("ICTITEM1_PEND")

        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub Load_Customers()
        ASCMAIN1.Progress("Loading Customers", "")

        Dim cust_data As String = ""
        Dim TEXT_FILENAME As String = ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
            (downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "cust*.txt")
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            TEXT_FILENAME = My.Computer.FileSystem.GetFileInfo(FILE).FullName
        Next

        Using sr As New System.IO.StreamReader(TEXT_FILENAME)
            cust_data = sr.ReadToEnd
        End Using

        Dim rows_imported As Integer = 0

        dst.Tables("ARTCUST1_IN").Rows.Clear()
        For Each line As String In Split(cust_data, vbCrLf)
            If line.Length > 1 Then
                Dim rowARTCUST1_IN As DataRow = dst.Tables("ARTCUST1_IN").NewRow
                rowARTCUST1_IN.Item("CUST_CODE") = Trim(line.Substring(0, 20)) 'Limit to 10?
                rowARTCUST1_IN.Item("NAME") = Trim(line.Substring(20, 30))
                'rowARTCUST1_IN.Item("NAME2") = Trim(line.Substring(51, 30))
                rowARTCUST1_IN.Item("ADDR1") = Trim(line.Substring(80, 30))
                rowARTCUST1_IN.Item("ADDR2") = Trim(line.Substring(110, 30))
                rowARTCUST1_IN.Item("CITY") = Trim(line.Substring(140, 30))
                rowARTCUST1_IN.Item("STATE") = Trim(line.Substring(170, 2))
                rowARTCUST1_IN.Item("ZIP") = Trim(line.Substring(200, 20))
                rowARTCUST1_IN.Item("COUNTRY") = Trim(line.Substring(220, 10))
                rowARTCUST1_IN.Item("PHONE") = Trim(line.Substring(230, 30))
                rowARTCUST1_IN.Item("FAX") = Trim(line.Substring(260, 30))
                rowARTCUST1_IN.Item("CONTACT") = Trim(line.Substring(290, 30))
                rowARTCUST1_IN.Item("SALESPERSON") = Trim(line.Substring(320, 10))
                rowARTCUST1_IN.Item("DISC_PCT") = Val(Trim(line.Substring(330, 6)))
                rowARTCUST1_IN.Item("GEN_BUS_POSTING_GROUP") = Trim(line.Substring(336, 10))
                rowARTCUST1_IN.Item("CUSTOMER_POSTING_GROUP") = Trim(line.Substring(346, 10))
                rowARTCUST1_IN.Item("PAYMENT_TERMS_CODE") = Trim(line.Substring(356, 10))
                rowARTCUST1_IN.Item("SHIPPING_PAYMENT_TYPE") = Trim(line.Substring(366, 10))
                dst.Tables("ARTCUST1_IN").Rows.Add(rowARTCUST1_IN)
                rows_imported += 1
            End If
        Next

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & ARTCUST1_TEMP)

        For Each rowARTCUST1_IN As DataRow In dst.Tables("ARTCUST1_IN").Select
            Dim rowARTCUST1_TEMP As DataRow = dst.Tables(ARTCUST1_TEMP).NewRow
            rowARTCUST1_TEMP.Item("CUST_CODE") = (rowARTCUST1_IN.Item("CUST_CODE") & Space(10)).substring(0, 10).trim
            rowARTCUST1_TEMP.Item("CUST_NAME") = rowARTCUST1_IN.Item("NAME")
            rowARTCUST1_TEMP.Item("CUST_ADDR1") = rowARTCUST1_IN.Item("ADDR1")
            rowARTCUST1_TEMP.Item("CUST_ADDR2") = rowARTCUST1_IN.Item("ADDR2")
            rowARTCUST1_TEMP.Item("CUST_CITY") = rowARTCUST1_IN.Item("CITY")
            rowARTCUST1_TEMP.Item("CUST_STATE") = (rowARTCUST1_IN.Item("STATE") & Space(2)).Substring(0, 2).Trim
            rowARTCUST1_TEMP.Item("CUST_ZIP_CODE") = (rowARTCUST1_IN.Item("ZIP") & Space(10)).Substring(0, 10).Trim
            rowARTCUST1_TEMP.Item("CUST_COUNTRY") = rowARTCUST1_IN.Item("COUNTRY")
            rowARTCUST1_TEMP.Item("CUST_PHONE") = (rowARTCUST1_IN.Item("PHONE").ToString.Replace("-", "").ToString.Replace(" ", "") & Space(10)).Substring(0, 10).Trim
            rowARTCUST1_TEMP.Item("CUST_FAX") = (rowARTCUST1_IN.Item("FAX").ToString.Replace("-", "").ToString.Replace(" ", "") & Space(10)).Substring(0, 10).Trim
            rowARTCUST1_TEMP.Item("CUST_CONTACT") = rowARTCUST1_IN.Item("CONTACT")
            rowARTCUST1_TEMP.Item("SREP_CODE") = (rowARTCUST1_IN.Item("SALESPERSON") & Space(10)).Substring(0, 10).Trim
            rowARTCUST1_TEMP.Item("PRICE_CLASS_CODE") = (rowARTCUST1_IN.Item("DISC_PCT") & Space(6)).Substring(0, 6).Trim
            'rowARTCUST1_TEMP.Item("") = rowARTCUST1_IN.Item("GEN_BUS_POSTING_GROUP")
            'rowARTCUST1_TEMP.Item("") = rowARTCUST1_IN.Item("CUSTOMER_POSTING_GROUP")
            rowARTCUST1_TEMP.Item("TERM_CODE") = (rowARTCUST1_IN.Item("PAYMENT_TERMS_CODE") & Space(6)).Substring(0, 6).Trim
            If rowARTCUST1_IN.Item("SHIPPING_PAYMENT_TYPE") = "Prepaid" Then
                rowARTCUST1_TEMP.Item("FRT_TERMS") = "PP"
            Else
                rowARTCUST1_TEMP.Item("FRT_TERMS") = "CC"
            End If

            rowARTCUST1_TEMP.Item("CURR_CODE") = "USD"
            rowARTCUST1_TEMP.Item("TRADE_CLASS_CODE") = "IND"
            rowARTCUST1_TEMP.Item("CUST_CLASS_CODE") = "BOUTQ"

            dst.Tables(ARTCUST1_TEMP).Rows.Add(rowARTCUST1_TEMP)
        Next

        Call Update_Record_TDA(ARTCUST1_TEMP)

        'New Customers
        ASCMAIN1.sql = " Update " & ARTCUST1_TEMP & " set CUST_STATUS = 'A' WHERE CUST_CODE IN " & vbCr _
        & " (Select CUST_CODE from  " & ARTCUST1_TEMP & vbCr _
        & " minus " & vbCr _
        & " Select CUST_CODE from ARTCUST1)"
        ASCDATA1.ExecuteSQL()

        'Update Customers
        ASCMAIN1.sql = " Update " & ARTCUST1_TEMP & " set CUST_STATUS = 'C' WHERE CUST_CODE IN (" & vbCr _
        & " Select CUST_CODE from (" & vbCr _
        & " Select CUST_CODE, CUST_NAME, CUST_ADDR1, CUST_ADDR2 " & vbCr _
        & " , CUST_CITY, CUST_STATE, CUST_ZIP_CODE, CUST_COUNTRY" & vbCr _
        & ", CUST_PHONE, CUST_FAX" & vbCr _
        & ", CUST_CONTACT, SREP_CODE, PRICE_CLASS_CODE, TERM_CODE, FRT_TERMS" & vbCr _
        & " from ARTCUST1" & vbCr _
        & " minus " & vbCr _
        & " Select CUST_CODE, CUST_NAME, CUST_ADDR1, CUST_ADDR2" & vbCr _
        & ", CUST_CITY, CUST_STATE, CUST_ZIP_CODE, CUST_COUNTRY" & vbCr _
        & ",  CUST_PHONE, CUST_FAX" & vbCr _
        & ", CUST_CONTACT, SREP_CODE, PRICE_CLASS_CODE, TERM_CODE, FRT_TERMS" & vbCr _
        & " from " & ARTCUST1_TEMP & "))"
        ASCDATA1.ExecuteSQL()

        Fill_Records("ARTCUST1_PEND")

        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub Load_Orders()
        ASCMAIN1.Progress("Loading Orders", "")

        Dim order_data As String = ""
        Dim TEXT_FILENAME As String = ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
            (downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "Orders*.txt")
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            TEXT_FILENAME = My.Computer.FileSystem.GetFileInfo(FILE).FullName
        Next

        Using sr As New System.IO.StreamReader(TEXT_FILENAME)
            order_data = sr.ReadToEnd
        End Using

        Dim rows_imported As Integer = 0
        Dim lno As Integer = 1
        Dim last_ord As String = ""

        dst.Tables("SOTORDR1_IN").Rows.Clear()
        dst.Tables("SOTORDR2_IN").Rows.Clear()
        dst.Tables("SOTORDRP").Rows.Clear()

        ASCMAIN1.Progress("Loading Order File", "")

        For Each line As String In Split(order_data, vbCrLf)
            If line.Length > 1 Then
                If Trim(line.Substring(0, 5)) <> "Order" Then 'headers
                    Dim rowSOTORDR1_IN As DataRow = dst.Tables("SOTORDR1_IN").NewRow
                    rowSOTORDR1_IN.Item("ORDR_NO") = Trim(line.Substring(0, 20))
                    rowSOTORDR1_IN.Item("SELL_TO_CUST_NO") = Trim(line.Substring(20, 10))
                    rowSOTORDR1_IN.Item("SELL_TO_NAME") = Trim(line.Substring(40, 30))
                    rowSOTORDR1_IN.Item("SELL_TO_ADDR") = Trim(line.Substring(70, 30))
                    rowSOTORDR1_IN.Item("SELL_TO_ADDR2") = Trim(line.Substring(100, 30))
                    rowSOTORDR1_IN.Item("SELL_TO_CITY") = Trim(line.Substring(130, 30))
                    rowSOTORDR1_IN.Item("SELL_TO_COUNTY") = Trim(line.Substring(160, 20))
                    rowSOTORDR1_IN.Item("SELL_TO_POST_CODE") = Trim(line.Substring(180, 20))
                    rowSOTORDR1_IN.Item("SELL_TO_CONTACT") = Trim(line.Substring(200, 30))
                    rowSOTORDR1_IN.Item("POSTING_DATE") = System.DateTime.Parse(Trim(line.Substring(230, 10)))
                    If Trim(line.Substring(240, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("ORDR_DATE") = System.DateTime.Parse(Trim(line.Substring(240, 10)))
                    Else
                        rowSOTORDR1_IN.Item("ORDR_DATE") = DATETIME_STAMP.Date
                    End If
                    rowSOTORDR1_IN.Item("DOCUMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(250, 10)))
                    If Trim(line.Substring(260, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("REQUESTED_DELIVERY_DATE") = System.DateTime.Parse(Trim(line.Substring(260, 10)))
                    Else
                        rowSOTORDR1_IN.Item("REQUESTED_DELIVERY_DATE") = Null
                    End If
                    rowSOTORDR1_IN.Item("EXTERNAL_DOC_NO") = Trim(line.Substring(270, 30))
                    rowSOTORDR1_IN.Item("PAYMENT_TERMS_CODE") = Trim(line.Substring(300, 10))
                    rowSOTORDR1_IN.Item("SALESPERSON_CODE") = Trim(line.Substring(310, 10))
                    rowSOTORDR1_IN.Item("STATUS") = Trim(line.Substring(320, 10))
                    rowSOTORDR1_IN.Item("TOTAL_POSTED_PKG") = Math.Abs(Val(Trim(line.Substring(330, 3).Replace(",", ""))))
                    rowSOTORDR1_IN.Item("TOTAL_PKGS") = Math.Abs(Val(Trim(line.Substring(333, 3).Replace(",", ""))))
                    rowSOTORDR1_IN.Item("BILL_TO_CUST_NO") = Trim(line.Substring(336, 10))
                    rowSOTORDR1_IN.Item("BILL_TO_NAME") = Trim(line.Substring(356, 30))
                    rowSOTORDR1_IN.Item("BILL_TO_ADDR") = Trim(line.Substring(386, 30))
                    rowSOTORDR1_IN.Item("BILL_TO_ADDR2") = Trim(line.Substring(416, 30))
                    rowSOTORDR1_IN.Item("BILL_TO_CITY") = Trim(line.Substring(446, 30))
                    rowSOTORDR1_IN.Item("BILL_TO_COUNTY") = Trim(line.Substring(476, 30))
                    rowSOTORDR1_IN.Item("BILL_TO_POST_CODE") = Trim(line.Substring(506, 20))
                    rowSOTORDR1_IN.Item("BILL_TO_CONTACT_NO") = Trim(line.Substring(526, 20))
                    rowSOTORDR1_IN.Item("BILL_TO_CONTACT") = Trim(line.Substring(546, 30))
                    rowSOTORDR1_IN.Item("PAYMENT_TERMS_CODE_2") = Trim(line.Substring(576, 10))
                    If Trim(line.Substring(586, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("DUE_DATE") = System.DateTime.Parse(Trim(line.Substring(586, 10)))
                    Else
                        rowSOTORDR1_IN.Item("DUE_DATE") = Null
                    End If
                    rowSOTORDR1_IN.Item("PAYMENT_DISC_PCT") = Math.Abs(Val(Trim(line.Substring(596, 3).Replace(",", ""))))
                    If Trim(line.Substring(599, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("PAYMENT_DISC_DATE") = System.DateTime.Parse(Trim(line.Substring(599, 10)))
                    Else
                        rowSOTORDR1_IN.Item("PAYMENT_DISC_DATE") = Null
                    End If
                    rowSOTORDR1_IN.Item("PAYMENT_METHOD_CODE") = Trim(line.Substring(609, 10))
                    rowSOTORDR1_IN.Item("TAX_AREA_CODE") = Trim(line.Substring(619, 10))
                    rowSOTORDR1_IN.Item("TAX_LIABLE") = Trim(line.Substring(629, 3))
                    rowSOTORDR1_IN.Item("SHIP_TO_CODE") = Trim(line.Substring(632, 10))
                    rowSOTORDR1_IN.Item("SHIP_TO_NAME") = Trim(line.Substring(642, 30))
                    rowSOTORDR1_IN.Item("SHIP_TO_ADDR") = Trim(line.Substring(672, 30))
                    rowSOTORDR1_IN.Item("SHIP_TO_ADDR2") = Trim(line.Substring(702, 30))
                    rowSOTORDR1_IN.Item("SHIP_TO_CITY") = Trim(line.Substring(732, 30))
                    rowSOTORDR1_IN.Item("SHIP_TO_COUNTY") = Trim(line.Substring(762, 30))
                    rowSOTORDR1_IN.Item("SHIP_TO_POST_CODE") = Trim(line.Substring(792, 30))
                    rowSOTORDR1_IN.Item("SHIP_TO_COUNTRY_CODE") = Trim(line.Substring(812, 10))
                    rowSOTORDR1_IN.Item("SHIP_TO_CONTACT") = Trim(line.Substring(822, 30))
                    rowSOTORDR1_IN.Item("SHIP_TO_SALESPERSON") = Trim(line.Substring(852, 10))
                    rowSOTORDR1_IN.Item("LOCATION_CODE") = Trim(line.Substring(862, 10))
                    rowSOTORDR1_IN.Item("SHIPMENT_METHOD_CODE") = Trim(line.Substring(872, 20))
                    rowSOTORDR1_IN.Item("SHIPPING_AGENT_CODE") = Trim(line.Substring(892, 20))
                    rowSOTORDR1_IN.Item("E_SHIP_AGENT_SVC") = Trim(line.Substring(912, 20))
                    rowSOTORDR1_IN.Item("SHIPMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(932, 10)))
                    rowSOTORDR1_IN.Item("SHIPPING_ADVICE") = Trim(line.Substring(942, 20))
                    If Trim(line.Substring(962, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("EARLIEST_SHIP_DATE") = System.DateTime.Parse(Trim(line.Substring(962, 10)))
                    Else
                        rowSOTORDR1_IN.Item("EARLIEST_SHIP_DATE") = rowSOTORDR1_IN.Item("ORDR_DATE")
                    End If
                    If Trim(line.Substring(972, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("CANCEL_AFTER_DATE") = System.DateTime.Parse(Trim(line.Substring(972, 10)))
                    Else
                        rowSOTORDR1_IN.Item("CANCEL_AFTER_DATE") = Null
                    End If
                    rowSOTORDR1_IN.Item("RESIDENTIAL_DELIVERY") = Trim(line.Substring(982, 3))
                    rowSOTORDR1_IN.Item("SHIPPING_ADVICE_2") = Trim(line.Substring(985, 20))
                    rowSOTORDR1_IN.Item("FREE_FREIGHT") = Trim(line.Substring(1005, 3))
                    rowSOTORDR1_IN.Item("EDI_ORDER") = Trim(line.Substring(1008, 3))
                    rowSOTORDR1_IN.Item("EDI_DEPT") = Trim(line.Substring(1011, 10))
                    If Trim(line.Substring(1021, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("EDI_EXPECTED_DELIVERY_DATE") = System.DateTime.Parse(Trim(line.Substring(1021, 10)))
                    Else
                        rowSOTORDR1_IN.Item("EDI_EXPECTED_DELIVERY_DATE") = Null
                    End If
                    rowSOTORDR1_IN.Item("EDI_TRADE_PARTNER") = Trim(line.Substring(1031, 20))
                    rowSOTORDR1_IN.Item("EDI_SELL_TO_CODE") = Trim(line.Substring(1051, 10))
                    rowSOTORDR1_IN.Item("EDI_SHIP_FOR_CODE") = Trim(line.Substring(1061, 10))
                    rowSOTORDR1_IN.Item("EDI_SHIP_TO_CODE") = Trim(line.Substring(1071, 10))
                    If Trim(line.Substring(1081, 10)) <> "" Then
                        rowSOTORDR1_IN.Item("EDI_CANCEL_AFTER_DATE") = System.DateTime.Parse(Trim(line.Substring(1081, 10)))
                    Else
                        rowSOTORDR1_IN.Item("EDI_CANCEL_AFTER_DATE") = Null
                    End If
                    rowSOTORDR1_IN.Item("EXTERNAL_DOC_NO_2") = Trim(line.Substring(1091, 30))
                    rowSOTORDR1_IN.Item("SHIP_FOR_CODE") = Trim(line.Substring(1121, 10))
                    rowSOTORDR1_IN.Item("CREATED_BY") = Trim(line.Substring(1131, 10))
                    rowSOTORDR1_IN.Item("CREATION_DATE") = System.DateTime.Parse(Trim(line.Substring(1141, 10)))
                    rowSOTORDR1_IN.Item("ON_HOLD") = Trim(line.Substring(1151, 3))
                    rowSOTORDR1_IN.Item("RELEASED_BY") = Trim(line.Substring(1151, 10))
                    If Trim(line.Substring(1164, 30)) <> "" Then
                        rowSOTORDR1_IN.Item("RELEASED_DATE") = System.DateTime.Parse(Trim(line.Substring(1164, 30)))
                    Else
                        rowSOTORDR1_IN.Item("RELEASED_DATE") = Null
                    End If
                    rowSOTORDR1_IN.Item("OPENED_BY") = Trim(line.Substring(1194, 10))
                    If Trim(line.Substring(1204, 20)) <> "" Then
                        rowSOTORDR1_IN.Item("OPENED_DATE") = System.DateTime.Parse(Trim(line.Substring(1204, 20)))
                    Else
                        rowSOTORDR1_IN.Item("OPENED_DATE") = Null
                    End If

                    dst.Tables("SOTORDR1_IN").Rows.Add(rowSOTORDR1_IN)

                Else 'details
                    Dim rowSOTORDR2_IN As DataRow = dst.Tables("SOTORDR2_IN").NewRow
                    rowSOTORDR2_IN.Item("ORDR_NO") = Trim(line.Substring(10, 20))
                    If last_ord <> Trim(line.Substring(10, 20)) Then
                        lno = 1
                        last_ord = Trim(line.Substring(10, 20))
                    Else
                        lno = lno + 1
                    End If
                    rowSOTORDR2_IN.Item("ORDR_LNO") = lno
                    rowSOTORDR2_IN.Item("LINE_ITEM_TYPE") = Trim(line.Substring(30, 15))
                    rowSOTORDR2_IN.Item("ITEM_CODE") = Trim(line.Substring(45, 20))
                    rowSOTORDR2_IN.Item("DESCRIPTION") = Trim(line.Substring(65, 30))
                    rowSOTORDR2_IN.Item("QTY") = Math.Abs(Val(Trim(line.Substring(95, 6).Replace(",", ""))))
                    rowSOTORDR2_IN.Item("UNIT_PRICE") = Trim(line.Substring(101, 10))
                    rowSOTORDR2_IN.Item("LINE_DISC_PCT") = Math.Abs(Val(Trim(line.Substring(111, 3).Replace(",", ""))))
                    rowSOTORDR2_IN.Item("LINE_AMT") = Math.Abs(Val(Trim(line.Substring(114, 10).Replace(",", ""))))
                    rowSOTORDR2_IN.Item("QTY_TO_SHIP") = Math.Abs(Val(Trim(line.Substring(124, 6).Replace(",", ""))))
                    rowSOTORDR2_IN.Item("QTY_SHIPPED") = Math.Abs(Val(Trim(line.Substring(130, 6).Replace(",", ""))))
                    rowSOTORDR2_IN.Item("QTY_INVOICED") = Math.Abs(Val(Trim(line.Substring(136, 6).Replace(",", ""))))
                    rowSOTORDR2_IN.Item("UOM") = Trim(line.Substring(142, 6))
                    rowSOTORDR2_IN.Item("ORIG_QTY_ORD") = Math.Abs(Val(Trim(line.Substring(148, 6).Replace(",", ""))))
                    rowSOTORDR2_IN.Item("UNIT_COST") = Trim(line.Substring(154, 10))
                    rowSOTORDR2_IN.Item("TAX_GROUP_CODE") = Trim(line.Substring(164, 10))
                    rowSOTORDR2_IN.Item("AMT_INCL_TAX") = Trim(line.Substring(174, 10))
                    rowSOTORDR2_IN.Item("LINE_DISC_AMT") = Trim(line.Substring(184, 10))
                    rowSOTORDR2_IN.Item("ALLOW_INV_DISC") = Trim(line.Substring(194, 3))
                    rowSOTORDR2_IN.Item("INV_DISC_AMT") = Trim(line.Substring(197, 10))
                    If Trim(line.Substring(207, 10)) & "" <> "" Then
                        rowSOTORDR2_IN.Item("PLANNED_DELIVERY_DATE") = System.DateTime.Parse(Trim(line.Substring(207, 10)))
                    End If
                    If Trim(line.Substring(217, 10)) & "" <> "" Then
                        rowSOTORDR2_IN.Item("PLANNED_SHIPMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(217, 10)))
                    End If
                    If Trim(line.Substring(227, 10)) & "" <> "" Then
                        rowSOTORDR2_IN.Item("SHIPMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(227, 10)))
                    End If
                    dst.Tables("SOTORDR2_IN").Rows.Add(rowSOTORDR2_IN)
                End If
            End If
        Next

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTORDR1_TEMP)
        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & SOTORDR2_TEMP)

        ASCMAIN1.Progress("Loading Order Headers", "")
        For Each rowSOTORDR1_IN As DataRow In dst.Tables("SOTORDR1_IN").Select
            Dim rowSOTORDR1_TEMP As DataRow = dst.Tables(SOTORDR1_TEMP).NewRow
            rowSOTORDR1_TEMP.Item("ORDR_NO") = rowSOTORDR1_IN.Item("ORDR_NO")
            rowSOTORDR1_TEMP.Item("ORDR_DATE") = rowSOTORDR1_IN.Item("ORDR_DATE")
            rowSOTORDR1_TEMP.Item("CUST_CODE") = rowSOTORDR1_IN.Item("SELL_TO_CUST_NO")
            If rowSOTORDR1_IN.Item("EDI_ORDER") <> "Yes" Then
                If rowSOTORDR1_IN.Item("SHIP_TO_CODE") = "" Then
                    rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = "000000"
                Else
                    If Val(rowSOTORDR1_IN.Item("SHIP_TO_CODE")) > 0 Then
                        rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = Format(Val(rowSOTORDR1_IN.Item("SHIP_TO_CODE")), "000000").ToString.Substring(0, 6)
                    Else
                        rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = (rowSOTORDR1_IN.Item("SHIP_TO_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                End If
            Else
                'EDI Orders
                If rowSOTORDR1_IN.Item("SHIP_TO_CODE") = "" Then
                    rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = "000000"
                ElseIf Len(rowSOTORDR1_IN.Item("EDI_SHIP_FOR_CODE")) > 6 Then
                    If Val(rowSOTORDR1_IN.Item("SHIP_TO_CODE")) > 0 Then
                        rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = Format(Val(rowSOTORDR1_IN.Item("SHIP_TO_CODE")), "000000").ToString.Substring(0, 6)
                    Else
                        rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = (rowSOTORDR1_IN.Item("SHIP_TO_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                Else
                    If Val(rowSOTORDR1_IN.Item("EDI_SHIP_FOR_CODE")) > 0 Then
                        rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = Format(Val(rowSOTORDR1_IN.Item("EDI_SHIP_FOR_CODE")), "000000")
                    Else
                        rowSOTORDR1_TEMP.Item("CUST_STORE_NO") = rowSOTORDR1_IN.Item("EDI_SHIP_FOR_CODE").ToString.Trim
                    End If
                End If
            End If
            rowSOTORDR1_TEMP.Item("CUST_NAME") = rowSOTORDR1_IN.Item("SELL_TO_NAME")
            rowSOTORDR1_TEMP.Item("CUST_STORE_LOCATION") = rowSOTORDR1_IN.Item("SHIP_TO_NAME")
            rowSOTORDR1_TEMP.Item("CUST_BILL_TO_CUST") = rowSOTORDR1_IN.Item("BILL_TO_CUST_NO")
            'rowSOTORDR1_TEMP.Item("CUST_DISC_PCT") = rowSOTORDR1_IN.Item("")
            'rowSOTORDR1_TEMP.Item("PRICE_CLASS_CODE") = rowSOTORDR1_IN.Item("")
            'rowSOTORDR1_TEMP.Item("TRADE_CLASS_CODE") = rowSOTORDR1_IN.Item("")
            rowSOTORDR1_TEMP.Item("ORDR_CUST_PO") = rowSOTORDR1_IN.Item("EXTERNAL_DOC_NO")
            rowSOTORDR1_TEMP.Item("ORDR_SHIP_DATE") = rowSOTORDR1_IN.Item("EARLIEST_SHIP_DATE")
            rowSOTORDR1_TEMP.Item("ORDR_CANCEL_DATE") = rowSOTORDR1_IN.Item("EDI_CANCEL_AFTER_DATE")
            If rowSOTORDR1_IN.Item("EDI_ORDER") = "Yes" Then
                rowSOTORDR1_TEMP.Item("ORDR_SOURCE") = "E"
            Else
                rowSOTORDR1_TEMP.Item("ORDR_SOURCE") = "K"
            End If
            rowSOTORDR1_TEMP.Item("ORDR_DEPT") = rowSOTORDR1_IN.Item("EDI_DEPT")
            rowSOTORDR1_TEMP.Item("ORDR_DATE_RECD") = rowSOTORDR1_IN.Item("ORDR_DATE")
            'rowSOTORDR1_TEMP.Item("ORDR_SHIP_TO") = rowSOTORDR1_IN.Item("")
            'rowSOTORDR1_TEMP.Item("ORDR_YYYYPP_BOOKED") = rowSOTORDR1_IN.Item("")
            rowSOTORDR1_TEMP.Item("FRT_TERMS") = IIf(rowSOTORDR1_IN.Item("SHIPMENT_METHOD_CODE") = "", "CC", rowSOTORDR1_IN.Item("SHIPMENT_METHOD_CODE"))
            'rowSOTORDR1_TEMP.Item("POST_CODE") = rowSOTORDR1_IN.Item("")
            rowSOTORDR1_TEMP.Item("TERM_CODE") = rowSOTORDR1_IN.Item("PAYMENT_METHOD_CODE")
            If rowSOTORDR1_IN.Item("SALESPERSON_CODE") <> "" Then
                rowSOTORDR1_TEMP.Item("SREP_CODE") = rowSOTORDR1_IN.Item("SALESPERSON_CODE")
            Else
                rowSOTORDR1_TEMP.Item("SREP_CODE") = "000"
            End If

            rowSOTORDR1_TEMP.Item("SREP_CODE") = rowSOTORDR1_IN.Item("SALESPERSON_CODE")
            rowSOTORDR1_TEMP.Item("WHSE_CODE") = rowSOTORDR1_IN.Item("LOCATION_CODE")
            'rowSOTORDR1_TEMP.Item("BRAND_CODE") = rowSOTORDR1_IN.Item("")
            rowSOTORDR1_TEMP.Item("INIT_OPER") = rowSOTORDR1_IN.Item("CREATED_BY")
            'rowSOTORDR1_TEMP.Item("LAST_OPER") = rowSOTORDR1_IN.Item("")
            rowSOTORDR1_TEMP.Item("INIT_DATE") = rowSOTORDR1_IN.Item("CREATION_DATE")
            'rowSOTORDR1_TEMP.Item("LAST_DATE") = rowSOTORDR1_IN.Item("")
            rowSOTORDR1_TEMP.Item("ORDR_DATE_REL") = rowSOTORDR1_IN.Item("RELEASED_DATE")
            If rowSOTORDR1_IN.Item("STATUS") & "" = "Released" Then
                rowSOTORDR1_TEMP.Item("ORDR_STATUS") = "R"
            Else
                rowSOTORDR1_TEMP.Item("ORDR_STATUS") = "O"
            End If
            rowSOTORDR1_TEMP.Item("CURR_CODE") = "USD"
            rowSOTORDR1_TEMP.Item("CURR_EXCH_RATE") = "1"
            rowSOTORDR1_TEMP.Item("ORDR_SHIP_INSTR") = rowSOTORDR1_IN.Item("SHIPPING_AGENT_CODE") & " " & _
                                                rowSOTORDR1_IN.Item("E_SHIP_AGENT_SVC")
            If Val(rowSOTORDR1_IN.Item("EDI_SHIP_TO_CODE")) > 6 Then
                rowSOTORDR1_TEMP.Item("CUST_DC_NO") = String.Empty
            ElseIf Val(rowSOTORDR1_IN.Item("EDI_SHIP_TO_CODE")) > 0 Then
                rowSOTORDR1_TEMP.Item("CUST_DC_NO") = Format(Val(rowSOTORDR1_IN.Item("EDI_SHIP_TO_CODE")), "000000")

            Else
                rowSOTORDR1_TEMP.Item("CUST_DC_NO") = rowSOTORDR1_IN.Item("EDI_SHIP_TO_CODE")
            End If

            'rowSOTORDR1_TEMP.Item("ORDR_INVOICED") = rowSOTORDR1_IN.Item("")
            'rowSOTORDR1_TEMP.Item("ORDR_FREIGHT") = rowSOTORDR1_IN.Item("")

            dst.Tables(SOTORDR1_TEMP).Rows.Add(rowSOTORDR1_TEMP)
        Next

        Call Update_Record_TDA(SOTORDR1_TEMP)

        ASCMAIN1.Progress("Loading Order Details", "")
        For Each rowSOTORDR2_IN As DataRow In dst.Tables("SOTORDR2_IN").Select("LINE_ITEM_TYPE = 'Item'", "", DataViewRowState.CurrentRows)
            Dim rowSOTORDR2_TEMP As DataRow = dst.Tables(SOTORDR2_TEMP).NewRow
            rowSOTORDR2_TEMP.Item("ORDR_NO") = rowSOTORDR2_IN.Item("ORDR_NO")
            rowSOTORDR2_TEMP.Item("ORDR_LNO") = rowSOTORDR2_IN.Item("ORDR_LNO")
            rowSOTORDR2_TEMP.Item("ITEM_CODE") = rowSOTORDR2_IN.Item("ITEM_CODE")
            rowSOTORDR2_TEMP.Item("ITEM_DESC") = rowSOTORDR2_IN.Item("DESCRIPTION")
            rowSOTORDR2_TEMP.Item("ORDR_UNIT_PRICE") = rowSOTORDR2_IN.Item("UNIT_PRICE") * _
                                                ((100 - rowSOTORDR2_IN.Item("LINE_DISC_PCT")) * 0.01)
            rowSOTORDR2_TEMP.Item("ITEM_RETAIL_PRICE") = rowSOTORDR2_IN.Item("UNIT_PRICE")
            rowSOTORDR2_TEMP.Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2_IN.Item("UNIT_PRICE") * _
                                               ((100 - rowSOTORDR2_IN.Item("LINE_DISC_PCT")) * 0.01)
            rowSOTORDR2_TEMP.Item("ITEM_RETAIL_PRICE_CURR") = rowSOTORDR2_IN.Item("UNIT_PRICE")
            rowSOTORDR2_TEMP.Item("ORDR_QTY") = rowSOTORDR2_IN.Item("QTY")
            rowSOTORDR2_TEMP.Item("ORDR_QTY_OPEN") = rowSOTORDR2_IN.Item("QTY")
            rowSOTORDR2_TEMP.Item("ORDR_QTY_PICK") = rowSOTORDR2_IN.Item("QTY")
            rowSOTORDR2_TEMP.Item("ORDR_QTY_SHIP") = rowSOTORDR2_IN.Item("QTY_SHIPPED")
            rowSOTORDR2_TEMP.Item("ORDR_QTY_CANC") = 0
            rowSOTORDR2_TEMP.Item("ORDR_QTY_ORIG") = rowSOTORDR2_IN.Item("ORIG_QTY_ORD")
            rowSOTORDR2_TEMP.Item("ORDR_QTY_BACK") = rowSOTORDR2_IN.Item("QTY")
            rowSOTORDR2_TEMP.Item("ORDR_QTY_ORIG_BACK") = rowSOTORDR2_IN.Item("QTY")
            'rowSOTORDR2_TEMP.Item("ORDR_STATUS") = 
            'rowSOTORDR2_TEMP.Item("ORDR_RELEASE") = 
            'rowSOTORDR2_TEMP.Item("CUST_CODE") = 
            'rowSOTORDR2_TEMP.Item("CUST_STORE_NO") = 
            'rowSOTORDR2_TEMP.Item("SREP_CODE") = 
            'rowSOTORDR2_TEMP.Item("WHSE_CODE") = 
            'rowSOTORDR2_TEMP.Item("ORDR_YYYYPP_UPDATED") = 
            dst.Tables(SOTORDR2_TEMP).Rows.Add(rowSOTORDR2_TEMP)
        Next

        Call Update_Record_TDA(SOTORDR2_TEMP)

        'update ORDR_YYYYPP_UPDATED - period based on Ship Date
        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select * from GLTPARM2 order by OPS_YYYYPP;" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update " & SOTORDR1_TEMP _
        & "    set  ORDR_YYYYPP_UPDATED = R1.OPS_YYYYPP" _
        & "    where ORDR_SHIP_DATE <= R1.PRD_END_DATE and ORDR_YYYYPP_UPDATED is Null;" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        'update detail fields from header
        ASCMAIN1.sql = "Begin " _
        & " Declare Cursor C1 is Select ORDR_NO, CUST_CODE, CUST_STORE_NO, ORDR_STATUS, " _
        & "                SREP_CODE, WHSE_CODE, ORDR_YYYYPP_UPDATED from " & SOTORDR1_TEMP & ";" _
        & " Begin " _
        & "  For R1 in C1 Loop" _
        & "   Update " & SOTORDR2_TEMP _
        & "    set  CUST_CODE=R1.CUST_CODE, CUST_STORE_NO=R1.CUST_STORE_NO," _
        & "         ORDR_STATUS=R1.ORDR_STATUS, SREP_CODE=R1.SREP_CODE, WHSE_CODE=R1.WHSE_CODE," _
        & "         ORDR_YYYYPP_UPDATED=R1.ORDR_YYYYPP_UPDATED" _
        & "    where ORDR_NO = R1.ORDR_NO;" _
        & "  End Loop;" _
        & " End;" _
        & "End;"
        ASCDATA1.ExecuteSQL()

        Fill_Records("SOTORDRP")
        Fill_Records("SOTORDRI")

        ASCMAIN1.Progress("", "")

    End Sub

    Private Sub Load_Invoices()
        ASCMAIN1.Progress("Loading Invoices", "")

        dst.Tables("SOTINVH1_IN").Rows.Clear()
        dst.Tables("SOTINVH2_IN").Rows.Clear()

        Dim invoice_data As String = ""
        Dim TEXT_FILENAME As String = ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
            (downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "Invoices*.txt")
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            TEXT_FILENAME = My.Computer.FileSystem.GetFileInfo(FILE).FullName

            Using sr As New System.IO.StreamReader(TEXT_FILENAME)
                invoice_data = sr.ReadToEnd
            End Using

            Dim rows_imported As Integer = 0
            Dim lno As Integer = 1
            Dim last_inv As String = ""

            Dim SHIPMENT_NO As String = ""

            For Each line As String In Split(invoice_data, vbCrLf)
                If line.Length > 1 AndAlso Trim(line.Substring(50, 16)) <> "Deleted Document" Then

                    If Trim(line.Substring(50, 12)) = "Shipment No." Then
                        SHIPMENT_NO = Trim(line.Substring(63, 20)).Replace(":", "")
                    ElseIf Trim(line.Substring(20, 10)) <> "Item" _
                           And Trim(line.Substring(20, 10)) <> "Resource" _
                           And Trim(line.Substring(20, 10)) <> "Tax" Then
                        SHIPMENT_NO = ""
                    End If

                    If Trim(line.Substring(50, 12)) <> "Shipment No." _
                    And Trim(line.Substring(20, 30)) <> "" Then
                        If Trim(line.Substring(20, 10)) <> "Item" _
                        And Trim(line.Substring(20, 10)) <> "Resource" _
                        And Trim(line.Substring(20, 10)) <> "Tax" _
                        And Trim(line.Substring(20, 10)) <> "G/L Accoun" Then 'headers
                            Dim rowSOTINVH1_IN As DataRow = dst.Tables("SOTINVH1_IN").NewRow
                            rowSOTINVH1_IN.Item("INV_NO") = Trim(line.Substring(0, 20))
                            rowSOTINVH1_IN.Item("SELL_TO_CUST_NO") = Trim(line.Substring(20, 10))
                            rowSOTINVH1_IN.Item("SELL_TO_NAME") = Trim(line.Substring(40, 30))
                            rowSOTINVH1_IN.Item("SELL_TO_ADDR") = Trim(line.Substring(70, 30))
                            rowSOTINVH1_IN.Item("SELL_TO_ADDR2") = Trim(line.Substring(100, 30))
                            rowSOTINVH1_IN.Item("SELL_TO_CITY") = Trim(line.Substring(130, 30))
                            rowSOTINVH1_IN.Item("SELL_TO_STATE") = Trim(line.Substring(160, 20))
                            rowSOTINVH1_IN.Item("SELL_TO_ZIP") = Trim(line.Substring(180, 30))
                            'rowSOTINVH1_IN.Item("SELL_TO_COUNTY") = Trim(line.Substring(160, 20))
                            'rowSOTINVH1_IN.Item("SELL_TO_POST_CODE") = Trim(line.Substring(180, 20))
                            rowSOTINVH1_IN.Item("SELL_TO_CONTACT") = Trim(line.Substring(200, 30))
                            rowSOTINVH1_IN.Item("POSTING_DATE") = System.DateTime.Parse(Trim(line.Substring(230, 20)))
                            rowSOTINVH1_IN.Item("DOCUMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(250, 10)))
                            rowSOTINVH1_IN.Item("ORDR_NO") = Trim(line.Substring(260, 10))
                            'If Trim(line.Substring(260, 10)) <> "" Then
                            '    rowSOTINVH1_IN.Item("REQUESTED_DELIVERY_DATE") = System.DateTime.Parse(Trim(line.Substring(260, 10)))
                            'Else
                            '    rowSOTINVH1_IN.Item("REQUESTED_DELIVERY_DATE") = Null
                            'End If
                            rowSOTINVH1_IN.Item("EXTERNAL_DOC_NO") = Trim(line.Substring(270, 30))
                            rowSOTINVH1_IN.Item("SALESPERSON_CODE") = Trim(line.Substring(310, 10))
                            rowSOTINVH1_IN.Item("NO_PRINTED") = Math.Abs(Val(Trim(line.Substring(320, 10).Replace(",", ""))))
                            'rowSOTINVH1_IN.Item("STATUS") = Trim(line.Substring(320, 10))
                            'rowSOTINVH1_IN.Item("TOTAL_POSTED_PKG") = Math.Abs(Val(Trim(line.Substring(330, 3).Replace(",", ""))))
                            rowSOTINVH1_IN.Item("TOTAL_PKGS") = Math.Abs(Val(Trim(line.Substring(330, 3).Replace(",", ""))))
                            rowSOTINVH1_IN.Item("TOTAL_WT") = Math.Abs(Val(Trim(line.Substring(333, 3).Replace(",", ""))))
                            rowSOTINVH1_IN.Item("AMOUNT_INCL_TAX") = Val(Trim(line.Substring(336, 10).Replace(",", "")))
                            rowSOTINVH1_IN.Item("BILL_TO_CUST_NO") = Trim(line.Substring(356, 10))
                            rowSOTINVH1_IN.Item("BILL_TO_NAME") = Trim(line.Substring(376, 30))
                            rowSOTINVH1_IN.Item("BILL_TO_ADDR") = Trim(line.Substring(406, 30))
                            rowSOTINVH1_IN.Item("BILL_TO_ADDR2") = Trim(line.Substring(436, 30))
                            rowSOTINVH1_IN.Item("BILL_TO_CITY") = Trim(line.Substring(466, 30))
                            rowSOTINVH1_IN.Item("BILL_TO_STATE") = Trim(line.Substring(496, 30))
                            rowSOTINVH1_IN.Item("BILL_TO_ZIP") = Trim(line.Substring(526, 20))
                            'rowSOTINVH1_IN.Item("BILL_TO_POST_CODE") = Trim(line.Substring(506, 20))
                            rowSOTINVH1_IN.Item("BILL_TO_CONTACT_NO") = Trim(line.Substring(546, 20))
                            rowSOTINVH1_IN.Item("BILL_TO_CONTACT") = Trim(line.Substring(566, 30))
                            rowSOTINVH1_IN.Item("PROD_LINE_CODE") = Trim(line.Substring(596, 10))
                            rowSOTINVH1_IN.Item("COMPANY_CODE") = Trim(line.Substring(606, 10))
                            rowSOTINVH1_IN.Item("PAYMENT_TERMS_CODE") = Trim(line.Substring(616, 3))
                            If Trim(line.Substring(619, 10)) <> "" Then
                                rowSOTINVH1_IN.Item("DUE_DATE") = System.DateTime.Parse(Trim(line.Substring(619, 10)))
                            Else
                                rowSOTINVH1_IN.Item("DUE_DATE") = Null
                            End If
                            rowSOTINVH1_IN.Item("PAYMENT_DISC_PCT") = Val(Trim(line.Substring(629, 3).Replace(",", "")))
                            If Trim(line.Substring(639, 10)) <> "" Then
                                rowSOTINVH1_IN.Item("PAYMENT_DISC_DATE") = System.DateTime.Parse(Trim(line.Substring(639, 10)))
                            Else
                                rowSOTINVH1_IN.Item("PAYMENT_DISC_DATE") = Null
                            End If
                            'rowSOTINVH1_IN.Item("PAYMENT_METHOD_CODE") = Trim(line.Substring(609, 10))
                            rowSOTINVH1_IN.Item("TAX_AREA_CODE") = Trim(line.Substring(649, 10))
                            rowSOTINVH1_IN.Item("TAX_LIABLE") = Trim(line.Substring(629, 3))
                            rowSOTINVH1_IN.Item("SHIP_TO_CODE") = Trim(line.Substring(662, 10))
                            rowSOTINVH1_IN.Item("SHIP_TO_NAME") = Trim(line.Substring(672, 30))
                            rowSOTINVH1_IN.Item("SHIP_TO_ADDR") = Trim(line.Substring(702, 30))
                            rowSOTINVH1_IN.Item("SHIP_TO_ADDR2") = Trim(line.Substring(732, 30))
                            rowSOTINVH1_IN.Item("SHIP_TO_CITY") = Trim(line.Substring(762, 30))
                            rowSOTINVH1_IN.Item("SHIP_TO_STATE") = Trim(line.Substring(792, 30))
                            rowSOTINVH1_IN.Item("SHIP_TO_ZIP") = Trim(line.Substring(822, 30))
                            'rowSOTINVH1_IN.Item("SHIP_TO_POST_CODE") = Trim(line.Substring(792, 30))
                            rowSOTINVH1_IN.Item("SHIP_TO_COUNTRY_CODE") = Trim(line.Substring(842, 10))
                            rowSOTINVH1_IN.Item("SHIP_TO_CONTACT") = Trim(line.Substring(852, 30))
                            rowSOTINVH1_IN.Item("SHIP_TO_SALESPERSON") = Trim(line.Substring(882, 10))
                            rowSOTINVH1_IN.Item("LOCATION_CODE") = Trim(line.Substring(892, 10))
                            rowSOTINVH1_IN.Item("SHIPMENT_METHOD_CODE") = Trim(line.Substring(902, 20))
                            If Trim(line.Substring(922, 10)) <> "" Then
                                rowSOTINVH1_IN.Item("SHIPMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(922, 10)))
                            Else
                                rowSOTINVH1_IN.Item("SHIPMENT_DATE") = Null
                            End If
                            If Trim(line.Substring(932, 10)) <> "" Then
                                rowSOTINVH1_IN.Item("SHIP_DATE") = System.DateTime.Parse(Trim(line.Substring(932, 10)))
                            Else
                                rowSOTINVH1_IN.Item("SHIP_DATE") = Null
                            End If
                            rowSOTINVH1_IN.Item("SHIPPING_AGENT_CODE") = Trim(line.Substring(942, 10))
                            rowSOTINVH1_IN.Item("E_SHIP_AGENT_SVC") = Trim(line.Substring(952, 20))
                            'rowSOTINVH1_IN.Item("SHIPMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(932, 10)))
                            'rowSOTINVH1_IN.Item("SHIPPING_ADVICE") = Trim(line.Substring(942, 20))
                            rowSOTINVH1_IN.Item("RESIDENTIAL_DELIVERY") = Trim(line.Substring(972, 3))
                            'rowSOTINVH1_IN.Item("SHIPPING_ADVICE_2") = Trim(line.Substring(985, 20))
                            'If Trim(line.Substring(1021, 10)) <> "" Then
                            '    rowSOTINVH1_IN.Item("EDI_EXPECTED_DELIVERY_DATE") = System.DateTime.Parse(Trim(line.Substring(1021, 10)))
                            'Else
                            '    rowSOTINVH1_IN.Item("EDI_EXPECTED_DELIVERY_DATE") = Null
                            'End If
                            rowSOTINVH1_IN.Item("SHIP_FOR_CODE") = Trim(line.Substring(975, 20))
                            rowSOTINVH1_IN.Item("SHIPPING_PAYMENT_TYPE") = Trim(line.Substring(995, 10))
                            rowSOTINVH1_IN.Item("SHIPPING_INSURANCE") = Trim(line.Substring(1005, 30))
                            rowSOTINVH1_IN.Item("FREE_FREIGHT") = Trim(line.Substring(1035, 3))
                            rowSOTINVH1_IN.Item("INVOICE_FOR_BOL") = Trim(line.Substring(1038, 20))
                            rowSOTINVH1_IN.Item("INVOICE_FOR_SHIPMENT") = Trim(line.Substring(1058, 30))
                            rowSOTINVH1_IN.Item("SHIPMENT_INV_OVERRIDE") = Trim(line.Substring(1088, 5))
                            rowSOTINVH1_IN.Item("EDI_ORDER") = Trim(line.Substring(1093, 5))
                            rowSOTINVH1_IN.Item("EDI_INV_GEN") = Trim(line.Substring(1098, 5))
                            If Trim(line.Substring(1103, 10)) <> "" Then
                                rowSOTINVH1_IN.Item("EDI_INV_GEN_DATE") = System.DateTime.Parse(Trim(line.Substring(1103, 10)))
                            Else
                                rowSOTINVH1_IN.Item("EDI_INV_GEN_DATE") = Null
                            End If
                            rowSOTINVH1_IN.Item("EDI_TRADE_PARTNER") = Trim(line.Substring(1113, 20))
                            rowSOTINVH1_IN.Item("EDI_SELL_TO_CODE") = Trim(line.Substring(1133, 10))
                            rowSOTINVH1_IN.Item("EDI_SHIP_FOR_CODE") = Trim(line.Substring(1143, 10))
                            rowSOTINVH1_IN.Item("EDI_SHIP_TO_CODE") = Trim(line.Substring(1153, 10))
                            If Trim(line.Substring(1163, 10)) <> "" Then
                                rowSOTINVH1_IN.Item("CANCEL_AFTER_DATE") = System.DateTime.Parse(Trim(line.Substring(1163, 10)))
                            Else
                                rowSOTINVH1_IN.Item("CANCEL_AFTER_DATE") = Null
                            End If
                            rowSOTINVH1_IN.Item("EXTERNAL_SHIP_FOR_NO") = Trim(line.Substring(1173, 30))
                            rowSOTINVH1_IN.Item("SHIP_FOR_CODE2") = Trim(line.Substring(1203, 20))
                            rowSOTINVH1_IN.Item("CREATED_BY") = Trim(line.Substring(1213, 10))
                            If Trim(line.Substring(1223, 10)) <> "" Then
                                rowSOTINVH1_IN.Item("CREATION_DATE") = System.DateTime.Parse(Trim(line.Substring(1223, 10)))
                            Else
                                rowSOTINVH1_IN.Item("CREATION_DATE") = Null
                            End If
                            rowSOTINVH1_IN.Item("RELEASED_BY") = Trim(line.Substring(1233, 10))
                            If Trim(line.Substring(1243, 20)) <> "" Then
                                rowSOTINVH1_IN.Item("RELEASED_DATE") = System.DateTime.Parse(Trim(line.Substring(1243, 10)) & " " & _
                                                                                             Trim(line.Substring(1253, 10)))
                            Else
                                rowSOTINVH1_IN.Item("RELEASED_DATE") = Null
                            End If
                            rowSOTINVH1_IN.Item("OPENED_BY") = Trim(line.Substring(1273, 10))
                            If Trim(line.Substring(1283, 20)) <> "" Then
                                rowSOTINVH1_IN.Item("OPENED_DATE") = System.DateTime.Parse(Trim(line.Substring(1283, 10)) & " " & _
                                                                                           Trim(line.Substring(1293, 10)))
                            Else
                                rowSOTINVH1_IN.Item("OPENED_DATE") = Null
                            End If
                            dst.Tables("SOTINVH1_IN").Rows.Add(rowSOTINVH1_IN)

                        Else 'details
                            Dim rowSOTINVH2_IN As DataRow = dst.Tables("SOTINVH2_IN").NewRow
                            rowSOTINVH2_IN.Item("INV_NO") = Trim(line.Substring(0, 20))
                            If last_inv <> Trim(line.Substring(0, 20)) Then
                                lno = 1
                                last_inv = Trim(line.Substring(0, 20))
                            Else
                                lno = lno + 1
                            End If
                            rowSOTINVH2_IN.Item("INV_LNO") = lno
                            rowSOTINVH2_IN.Item("LINE_ITEM_TYPE") = Trim(line.Substring(20, 10))
                            rowSOTINVH2_IN.Item("ITEM_CODE") = Trim(line.Substring(30, 20))
                            rowSOTINVH2_IN.Item("DESCRIPTION") = Trim(line.Substring(50, 30))
                            rowSOTINVH2_IN.Item("RET_REASON_CODE") = Trim(line.Substring(80, 3))
                            rowSOTINVH2_IN.Item("PACKAGE_TRACKING_NO") = Trim(line.Substring(83, 32))
                            rowSOTINVH2_IN.Item("SHIPMENT_NO") = SHIPMENT_NO
                            rowSOTINVH2_IN.Item("QTY") = Val(Trim(line.Substring(115, 6).Replace(",", "")))
                            rowSOTINVH2_IN.Item("QTY_ORD") = Val(Trim(line.Substring(121, 6).Replace(",", "")))
                            rowSOTINVH2_IN.Item("UOM_CODE") = Trim(line.Substring(127, 6))
                            rowSOTINVH2_IN.Item("UNIT_COST") = Trim(line.Substring(133, 10))
                            rowSOTINVH2_IN.Item("TAX_GROUP_CODE") = Trim(line.Substring(143, 10))
                            rowSOTINVH2_IN.Item("UNIT_PRICE") = Trim(line.Substring(153, 10))
                            rowSOTINVH2_IN.Item("LINE_AMT") = Val(Trim(line.Substring(163, 10).Replace(",", "")))
                            rowSOTINVH2_IN.Item("LINE_DISC_PCT") = Val(Trim(line.Substring(183, 3).Replace(",", "")))
                            rowSOTINVH2_IN.Item("LINE_DISC_AMT") = Trim(line.Substring(186, 10))
                            rowSOTINVH2_IN.Item("ALLOW_INV_DISC") = Trim(line.Substring(196, 3))
                            rowSOTINVH2_IN.Item("PROD_LINE_CODE") = Trim(line.Substring(199, 10))
                            rowSOTINVH2_IN.Item("COMPANY_CODE") = Trim(line.Substring(209, 3))
                            dst.Tables("SOTINVH2_IN").Rows.Add(rowSOTINVH2_IN)
                        End If
                    End If
                End If

            Next
        Next

        For Each rowSOTINVH1_IN As DataRow In dst.Tables("SOTINVH1_IN").Select
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow

            Dim INV_NO As String = rowSOTINVH1_IN.Item("INV_NO")
            'rowSOTINVH1.Item("INV_NO") = rowSOTINVH1_IN.Item("INV_NO")
            rowSOTINVH1.Item("INV_NO") = Format(Val(rowSOTINVH1_IN.Item("INV_NO")), "0000000000")
            rowSOTINVH1.Item("CUST_CODE") = rowSOTINVH1_IN.Item("SELL_TO_CUST_NO")

            Dim CUST_STORE_NO As String = ""
            If rowSOTINVH1_IN.Item("EDI_ORDER") <> "Yes" Then
                If rowSOTINVH1_IN.Item("SHIP_TO_CODE") = "" Then
                    CUST_STORE_NO = "000000"
                Else
                    If Val(rowSOTINVH1_IN.Item("SHIP_TO_CODE")) > 0 Then
                        CUST_STORE_NO = Format(Val(rowSOTINVH1_IN.Item("SHIP_TO_CODE")), "000000").ToString.Substring(0, 6)
                    Else
                        CUST_STORE_NO = (rowSOTINVH1_IN.Item("SHIP_TO_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                End If
            Else
                'EDI Orders
                If rowSOTINVH1_IN.Item("SHIP_TO_CODE") = "" Then
                    CUST_STORE_NO = "000000"
                ElseIf Len(rowSOTINVH1_IN.Item("EDI_SHIP_FOR_CODE")) > 6 Then
                    If Val(rowSOTINVH1_IN.Item("SHIP_TO_CODE")) > 0 Then
                        CUST_STORE_NO = Format(Val(rowSOTINVH1_IN.Item("SHIP_TO_CODE")), "000000").ToString.Substring(0, 6)
                    Else
                        CUST_STORE_NO = (rowSOTINVH1_IN.Item("SHIP_TO_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                Else
                    If Val(rowSOTINVH1_IN.Item("EDI_SHIP_FOR_CODE")) > 0 Then
                        CUST_STORE_NO = Format(Val(rowSOTINVH1_IN.Item("EDI_SHIP_FOR_CODE")), "000000")
                    Else
                        CUST_STORE_NO = rowSOTINVH1_IN.Item("EDI_SHIP_FOR_CODE").ToString.Trim
                    End If
                End If
            End If
            If CUST_STORE_NO = "" Then
                CUST_STORE_NO = "000000"
            End If
            rowSOTINVH1.Item("CUST_STORE_NO") = CUST_STORE_NO
            rowSOTINVH1.Item("ORDR_CUST_PO") = rowSOTINVH1_IN.Item("EXTERNAL_DOC_NO")
            If Trim(rowSOTINVH1_IN.Item("LOCATION_CODE")) = "CAUD TX" Then
                rowSOTINVH1.Item("WHSE_CODE") = "TX"
            ElseIf Trim(rowSOTINVH1_IN.Item("LOCATION_CODE")) = "" Then
                rowSOTINVH1.Item("WHSE_CODE") = "PDR"
            Else
                rowSOTINVH1.Item("WHSE_CODE") = (rowSOTINVH1_IN.Item("LOCATION_CODE") & Space(6)).ToString.Substring(0, 6).Trim
            End If

            rowSOTINVH1.Item("ORDR_NO") = ""
            rowSOTINVH1.Item("POST_CODE") = ""
            If rowSOTINVH1_IN.Item("SALESPERSON_CODE") <> "" Then
                rowSOTINVH1.Item("SREP_CODE") = rowSOTINVH1_IN.Item("SALESPERSON_CODE")
            Else
                rowSOTINVH1.Item("SREP_CODE") = "000"
            End If

            rowSOTINVH1.Item("REASON_CODE") = ""
            rowSOTINVH1.Item("CUST_BILL_TO_CUST") = rowSOTINVH1_IN.Item("BILL_TO_CUST_NO")
            rowSOTINVH1.Item("BRAND_CODE") = ""
            rowSOTINVH1.Item("EVENT_CODE") = ""
            rowSOTINVH1.Item("INV_DATE") = rowSOTINVH1_IN.Item("DOCUMENT_DATE") 'POSTING_DATE?

            rowSOTINVH1.Item("CUST_BILL_TO_CUST") = rowSOTINVH1_IN.Item("BILL_TO_CUST_NO")

            rowSOTINVH1.Item("INIT_DATE") = rowSOTINVH1_IN.Item("CREATION_DATE")
            rowSOTINVH1.Item("INIT_OPER") = rowSOTINVH1_IN.Item("CREATED_BY")
            rowSOTINVH1.Item("CURR_CODE") = "USD"
            rowSOTINVH1.Item("CURR_EXCH_RATE") = "1"

            rowSOTINVH1.Item("INV_DATE_SHIPPED") = rowSOTINVH1_IN.Item("SHIPMENT_DATE") ' SHIP_DATE?
            rowSOTINVH1.Item("INV_CARTONS") = rowSOTINVH1_IN.Item("TOTAL_PKGS")
            rowSOTINVH1.Item("INV_WEIGHT") = rowSOTINVH1_IN.Item("TOTAL_WT")
            rowSOTINVH1.Item("INV_BOL_NO") = rowSOTINVH1_IN.Item("INVOICE_FOR_BOL")
            rowSOTINVH1.Item("INV_PRO_NO") = ""
            rowSOTINVH1.Item("SHIP_VIA_DESC") = ""
            rowSOTINVH1.Item("INV_PRO_NO") = ""
            rowSOTINVH1.Item("INV_NO_CONS") = ""
            rowSOTINVH1.Item("SHIPMENT_NO") = ""
            'rowSOTINVH1.Item("INV_NO_REV") = ""
            If Len(rowSOTINVH1_IN.Item("SHIP_TO_STATE")) <= 2 Then
                rowSOTINVH1.Item("CUST_SHIP_TO_STATE") = rowSOTINVH1_IN.Item("SHIP_TO_STATE")
            End If
            'rowSOTINVH1.Item("INV_NO_REV_BY") = ""
            rowSOTINVH1.Item("INV_COMMENT") = ""
            rowSOTINVH1.Item("INV_FREIGHT_TAX") = 0
            rowSOTINVH1.Item("SHIP_VIA_CODE") = ""
            rowSOTINVH1.Item("STAX_CODE") = ""
            rowSOTINVH1.Item("ORDER_TYPE_CODE") = ""
            rowSOTINVH1.Item("POSTRUN") = ""
            rowSOTINVH1.Item("TRANSID") = ""
            rowSOTINVH1.Item("BATCHID") = ""
            rowSOTINVH1.Item("INVCNUM") = ""
            rowSOTINVH1.Item("TRANSID") = ""
            rowSOTINVH1.Item("SHIP_TO_CITY") = rowSOTINVH1_IN.Item("SHIP_TO_CITY")

            Dim INV_SALES As Double = 0
            Dim INV_COGS As Double = 0
            Dim INV_FREIGHT As Double = 0
            Dim INV_STAX As Double = 0
            Dim INV_MISC_CHG As Double = 0
            Dim INV_TOTAL_AMOUNT As Double = 0

            For Each rowSOTINVH2_IN As DataRow In dst.Tables("SOTINVH2_IN").Select("INV_NO = '" & INV_NO & "'", "", DataViewRowState.CurrentRows)

                Dim LINE_AMOUNT As Double = rowSOTINVH2_IN.Item("LINE_AMT") '- rowSOTINVH2_IN.Item("LINE_DISC_AMT")
                'LINE AMOUNT IS ORDER QTY * PRICE (not shipped qty)

                Dim UNIT_PRICE As Decimal = rowSOTINVH2_IN.Item("UNIT_PRICE") * _
                                                       ((100 - rowSOTINVH2_IN.Item("LINE_DISC_PCT")) * 0.01)

                If rowSOTINVH2_IN.Item("LINE_ITEM_TYPE") = "Item" Then
                    Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").NewRow
                    'rowSOTINVH2.Item("INV_NO") = rowSOTINVH2_IN.Item("INV_NO")
                    rowSOTINVH2.Item("INV_NO") = Format(Val(rowSOTINVH2_IN.Item("INV_NO")), "0000000000")
                    rowSOTINVH2.Item("INV_LNO") = rowSOTINVH2_IN.Item("INV_LNO")
                    rowSOTINVH2.Item("ITEM_CODE") = rowSOTINVH2_IN.Item("ITEM_CODE")
                    rowSOTINVH2.Item("ORDR_UNIT_PRICE") = UNIT_PRICE
                    rowSOTINVH2.Item("ITEM_RETAIL_PRICE") = rowSOTINVH2_IN.Item("UNIT_PRICE")
                    rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") = UNIT_PRICE
                    rowSOTINVH2.Item("ITEM_RETAIL_PRICE_CURR") = rowSOTINVH2_IN.Item("UNIT_PRICE")
                    rowSOTINVH2.Item("ORDR_QTY_SHIP") = rowSOTINVH2_IN.Item("QTY")
                    rowSOTINVH2.Item("ITEM_UNIT_COST") = rowSOTINVH2_IN.Item("UNIT_COST")
                    rowSOTINVH2.Item("CUST_CODE") = rowSOTINVH1_IN.Item("SELL_TO_CUST_NO")
                    rowSOTINVH2.Item("CUST_STORE_NO") = CUST_STORE_NO
                    rowSOTINVH2.Item("SREP_CODE") = rowSOTINVH1_IN.Item("SALESPERSON_CODE")
                    If Trim(rowSOTINVH1_IN.Item("LOCATION_CODE")) = "CAUD TX" Then
                        rowSOTINVH2.Item("WHSE_CODE") = "TX"
                    ElseIf Trim(rowSOTINVH1_IN.Item("LOCATION_CODE")) = "" Then
                        rowSOTINVH2.Item("WHSE_CODE") = "PDR"
                    Else
                        rowSOTINVH2.Item("WHSE_CODE") = (rowSOTINVH1_IN.Item("LOCATION_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                    'rowSOTINVH2.Item("POSTRUN") = ""
                    'rowSOTINVH2.Item("TRANSID") = ""
                    'rowSOTINVH2.Item("INVCNUM") = ""

                    INV_COGS = INV_COGS + rowSOTINVH2_IN.Item("UNIT_COST") * rowSOTINVH2_IN.Item("QTY")
                    INV_SALES = INV_SALES + (rowSOTINVH2_IN.Item("QTY") * UNIT_PRICE)
                    INV_TOTAL_AMOUNT = INV_TOTAL_AMOUNT + (rowSOTINVH2_IN.Item("QTY") * UNIT_PRICE)

                    rowSOTINVH2.Item("INV_TYPE") = "I"
                    rowSOTINVH2.Item("ORDR_YYYYPP_UPDATED") = "XXXXXX"
                    rowSOTINVH2.Item("GLACCTSALES") = "4000"
                    dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)

                    'DO NOT WRITE SOTINVH2 FOR RESOURCE
                ElseIf rowSOTINVH2_IN.Item("LINE_ITEM_TYPE") = "Resource" Then
                    If rowSOTINVH2_IN.Item("ITEM_CODE") = "SHIPPING" Then
                        INV_FREIGHT = INV_FREIGHT + LINE_AMOUNT
                    ElseIf rowSOTINVH2_IN.Item("ITEM_CODE") = "STAX" Then
                        INV_STAX = INV_STAX + LINE_AMOUNT
                    End If
                Else
                    INV_MISC_CHG = INV_MISC_CHG + LINE_AMOUNT
                End If

            Next

            rowSOTINVH1.Item("INV_TYPE") = "I"
            rowSOTINVH1.Item("INV_SALES") = INV_SALES
            rowSOTINVH1.Item("INV_COGS") = INV_COGS
            rowSOTINVH1.Item("INV_FREIGHT") = INV_FREIGHT
            rowSOTINVH1.Item("INV_MISC_CHG") = INV_MISC_CHG
            rowSOTINVH1.Item("INV_STAX") = INV_STAX
            rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = INV_TOTAL_AMOUNT

            rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = "XXXXXX"
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)
        Next

        ASCMAIN1.Progress("", "")

    End Sub
    Private Sub Load_Credits()
        ASCMAIN1.Progress("Loading Credits", "")

        dst.Tables("SOTINVH1_CR").Rows.Clear()
        dst.Tables("SOTINVH2_CR").Rows.Clear()

        Dim invoice_data As String = ""
        Dim TEXT_FILENAME As String = ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
            (downloadFileLocation, FileIO.SearchOption.SearchTopLevelOnly, "Credits*.txt")
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            TEXT_FILENAME = My.Computer.FileSystem.GetFileInfo(FILE).FullName

            Using sr As New System.IO.StreamReader(TEXT_FILENAME)
                invoice_data = sr.ReadToEnd
            End Using

            Dim rows_imported As Integer = 0
            Dim lno As Integer = 1
            Dim last_crm As String = ""

            Dim SHIPMENT_NO As String = ""

            For Each line As String In Split(invoice_data, vbCrLf)
                If line.Length > 1 AndAlso Trim(line.Substring(50, 16)) <> "Deleted Document" Then

                    If Trim(line.Substring(50, 12)) = "Shipment No." Then
                        SHIPMENT_NO = Trim(line.Substring(63, 20)).Replace(":", "")
                    ElseIf Trim(line.Substring(20, 10)) <> "Item" _
                           And Trim(line.Substring(20, 10)) <> "Resource" _
                           And Trim(line.Substring(20, 10)) <> "Tax" Then
                        SHIPMENT_NO = ""
                    End If

                    If Trim(line.Substring(50, 12)) <> "Shipment No." _
                    And Trim(line.Substring(20, 30)) <> "" Then
                        If Trim(line.Substring(20, 10)) <> "Item" _
                        And Trim(line.Substring(20, 10)) <> "Resource" _
                        And Trim(line.Substring(20, 10)) <> "Tax" _
                        And Trim(line.Substring(20, 10)) <> "G/L Accoun" Then 'headers
                            Dim rowSOTINVH1_CR As DataRow = dst.Tables("SOTINVH1_CR").NewRow
                            rowSOTINVH1_CR.Item("SELL_TO_CUST_NO") = Trim(line.Substring(0, 10))
                            rowSOTINVH1_CR.Item("CM_NO") = Trim(line.Substring(21, 20))
                            rowSOTINVH1_CR.Item("BILL_TO_CUST_NO") = Trim(line.Substring(42, 10))
                            rowSOTINVH1_CR.Item("BILL_TO_NAME") = Trim(line.Substring(63, 30))
                            rowSOTINVH1_CR.Item("BILL_TO_ADDR") = Trim(line.Substring(125, 30))
                            rowSOTINVH1_CR.Item("BILL_TO_ADDR2") = Trim(line.Substring(156, 30))
                            rowSOTINVH1_CR.Item("BILL_TO_CITY") = Trim(line.Substring(187, 30))
                            rowSOTINVH1_CR.Item("BILL_TO_STATE") = Trim(line.Substring(1209, 30))   'STATE
                            rowSOTINVH1_CR.Item("BILL_TO_ZIP") = Trim(line.Substring(1188, 20)) 'ZIP
                            rowSOTINVH1_CR.Item("BILL_TO_COUNTRY_CODE") = Trim(line.Substring(1240, 10)) 'COUNTRY

                            rowSOTINVH1_CR.Item("SHIP_TO_CODE") = Trim(line.Substring(280, 10))
                            rowSOTINVH1_CR.Item("SHIP_TO_NAME") = Trim(line.Substring(291, 30))
                            rowSOTINVH1_CR.Item("SHIP_TO_NAME2") = Trim(line.Substring(322, 30))
                            rowSOTINVH1_CR.Item("SHIP_TO_ADDR") = Trim(line.Substring(353, 30))
                            rowSOTINVH1_CR.Item("SHIP_TO_ADDR2") = Trim(line.Substring(384, 30))
                            rowSOTINVH1_CR.Item("SHIP_TO_CITY") = Trim(line.Substring(415, 30))
                            rowSOTINVH1_CR.Item("SHIP_TO_STATE") = Trim(line.Substring(1335, 30))   'STATE
                            rowSOTINVH1_CR.Item("SHIP_TO_ZIP") = Trim(line.Substring(1314, 20)) 'ZIP
                            rowSOTINVH1_CR.Item("SHIP_TO_COUNTRY_CODE") = Trim(line.Substring(1366, 10)) 'COUNTRY
                            rowSOTINVH1_CR.Item("SHIP_TO_CONTACT") = Trim(line.Substring(446, 30))
                            rowSOTINVH1_CR.Item("POSTING_DATE") = System.DateTime.Parse(Trim(line.Substring(477, 20)))
                            If Trim(line.Substring(489, 10)) <> "" Then
                                rowSOTINVH1_CR.Item("SHIPMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(489, 10)))
                            Else
                                rowSOTINVH1_CR.Item("SHIPMENT_DATE") = Null
                            End If
                            rowSOTINVH1_CR.Item("POSTING_DESC") = Trim(line.Substring(501, 50))
                            rowSOTINVH1_CR.Item("PAYMENT_TERMS_CODE") = Trim(line.Substring(552, 3))
                            If Trim(line.Substring(563, 10)) <> "" Then
                                rowSOTINVH1_CR.Item("DUE_DATE") = System.DateTime.Parse(Trim(line.Substring(563, 10)))
                            Else
                                rowSOTINVH1_CR.Item("DUE_DATE") = Null
                            End If
                            rowSOTINVH1_CR.Item("PAYMENT_DISC_PCT") = Val(Trim(line.Substring(575, 3).Replace(",", "")))
                            'If Trim(line.Substring(639, 10)) <> "" Then
                            '    rowSOTINVH1_CR.Item("PAYMENT_DISC_DATE") = System.DateTime.Parse(Trim(line.Substring(639, 10)))
                            'Else
                            '    rowSOTINVH1_CR.Item("PAYMENT_DISC_DATE") = Null
                            'End If
                            rowSOTINVH1_CR.Item("LOCATION_CODE") = Trim(line.Substring(611, 10))
                            rowSOTINVH1_CR.Item("SHORTCUT_DIM_1_CODE") = Trim(line.Substring(643, 20)) 'USA
                            rowSOTINVH1_CR.Item("SHORTCUT_DIM_2_CODE") = Trim(line.Substring(664, 20)) 'CUSTOMER
                            rowSOTINVH1_CR.Item("CUSTOMER_POSTING_GROUP") = Trim(line.Substring(686, 10))
                            rowSOTINVH1_CR.Item("SALESPERSON_CODE") = Trim(line.Substring(764, 10))

                            rowSOTINVH1_CR.Item("APPLIES_TO_DOC_TYPE") = Trim(line.Substring(805, 10))
                            rowSOTINVH1_CR.Item("APPLIES_TO_DOC_NO") = Trim(line.Substring(816, 20))
                            rowSOTINVH1_CR.Item("AMOUNT") = Val(Trim(line.Substring(879, 10).Replace(",", "")))
                            rowSOTINVH1_CR.Item("AMOUNT_INCL_VAT") = Val(Trim(line.Substring(892, 10).Replace(",", "")))
                            rowSOTINVH1_CR.Item("RET_REASON_CODE") = Trim(line.Substring(936, 10))
                            rowSOTINVH1_CR.Item("GEN_BUS_POSTING_GROUP") = Trim(line.Substring(947, 10))
                            rowSOTINVH1_CR.Item("EU_3_PARTY_TRADE") = Trim(line.Substring(958, 10))
                            rowSOTINVH1_CR.Item("TRANSACTION_TYPE") = Trim(line.Substring(969, 10))
                            rowSOTINVH1_CR.Item("TRANSPORT_METHOD") = Trim(line.Substring(980, 10))
                            rowSOTINVH1_CR.Item("VAT_COUNTRY_CODE") = Trim(line.Substring(991, 10))

                            rowSOTINVH1_CR.Item("SELL_TO_NAME") = Trim(line.Substring(1002, 30))
                            rowSOTINVH1_CR.Item("SELL_TO_NAME2") = Trim(line.Substring(1033, 30))
                            rowSOTINVH1_CR.Item("SELL_TO_ADDR") = Trim(line.Substring(1064, 30))
                            rowSOTINVH1_CR.Item("SELL_TO_ADDR2") = Trim(line.Substring(1095, 30))
                            rowSOTINVH1_CR.Item("SELL_TO_CITY") = Trim(line.Substring(1126, 30))
                            rowSOTINVH1_CR.Item("SELL_TO_ZIP") = Trim(line.Substring(1251, 20)) 'ZIP
                            rowSOTINVH1_CR.Item("SELL_TO_STATE") = Trim(line.Substring(1272, 30)) 'STATE
                            rowSOTINVH1_CR.Item("SELL_TO_COUNTRY_CODE") = Trim(line.Substring(1303, 10))
                            rowSOTINVH1_CR.Item("SELL_TO_CONTACT") = Trim(line.Substring(1157, 30))

                            rowSOTINVH1_CR.Item("DOCUMENT_DATE") = System.DateTime.Parse(Trim(line.Substring(1410, 10)))
                            rowSOTINVH1_CR.Item("EXTERNAL_DOC_NO") = Trim(line.Substring(1422, 20))
                            rowSOTINVH1_CR.Item("PRE_ASSIGNED_NO") = Trim(line.Substring(1498, 20))
                            rowSOTINVH1_CR.Item("USER_ID") = Trim(line.Substring(1519, 20))
                            rowSOTINVH1_CR.Item("SOURCE_CODE") = Trim(line.Substring(1540, 10))
                            rowSOTINVH1_CR.Item("TAX_AREA_CODE") = Trim(line.Substring(1551, 20))
                            rowSOTINVH1_CR.Item("TAX_LIABLE") = Trim(line.Substring(1572, 10))
                            rowSOTINVH1_CR.Item("RETURN_ORDER_NO") = Trim(line.Substring(1692, 20))
                            rowSOTINVH1_CR.Item("ALLOW_LINE_DISC") = Trim(line.Substring(1724, 10))
                            rowSOTINVH1_CR.Item("EDI_ORDER") = Trim(line.Substring(1920, 10))
                            rowSOTINVH1_CR.Item("EDI_TRADE_PARTNER") = Trim(line.Substring(1965, 20))
                            rowSOTINVH1_CR.Item("EDI_SELL_TO_CODE") = Trim(line.Substring(1986, 20))
                            rowSOTINVH1_CR.Item("EDI_SHIP_TO_CODE") = Trim(line.Substring(2007, 20))
                            rowSOTINVH1_CR.Item("EDI_SHIP_FOR_CODE") = Trim(line.Substring(2028, 20))

                            dst.Tables("SOTINVH1_CR").Rows.Add(rowSOTINVH1_CR)

                        Else 'details
                            If Trim(line.Substring(20, 10)) <> "" Then
                                Dim rowSOTINVH2_CR As DataRow = dst.Tables("SOTINVH2_CR").NewRow
                                rowSOTINVH2_CR.Item("CM_NO") = Trim(line.Substring(0, 20))
                                If last_crm <> Trim(line.Substring(0, 20)) Then
                                    lno = 1
                                    last_crm = Trim(line.Substring(0, 20))
                                Else
                                    lno = lno + 1
                                End If
                                rowSOTINVH2_CR.Item("CM_LNO") = lno
                                rowSOTINVH2_CR.Item("LINE_ITEM_TYPE") = Trim(line.Substring(20, 10))
                                rowSOTINVH2_CR.Item("ITEM_CODE") = Trim(line.Substring(30, 20))
                                rowSOTINVH2_CR.Item("DESCRIPTION") = Trim(line.Substring(50, 30))
                                rowSOTINVH2_CR.Item("RET_REASON_CODE") = Trim(line.Substring(80, 3))
                                rowSOTINVH2_CR.Item("QTY") = Val(Trim(line.Substring(115, 6).Replace(",", "")))
                                rowSOTINVH2_CR.Item("UOM") = Trim(line.Substring(121, 6))
                                rowSOTINVH2_CR.Item("UNIT_COST") = Trim(line.Substring(127, 10))
                                rowSOTINVH2_CR.Item("TAX_GROUP_CODE") = Trim(line.Substring(137, 10))
                                rowSOTINVH2_CR.Item("UNIT_PRICE") = Trim(line.Substring(147, 10))
                                rowSOTINVH2_CR.Item("LINE_AMT") = Val(Trim(line.Substring(157, 10).Replace(",", "")))
                                rowSOTINVH2_CR.Item("AMT_INCL_TAX") = Val(Trim(line.Substring(167, 10).Replace(",", "")))
                                rowSOTINVH2_CR.Item("LINE_DISC_PCT") = Val(Trim(line.Substring(177, 3).Replace(",", "")))
                                rowSOTINVH2_CR.Item("LINE_DISC_AMT") = Val(Trim(line.Substring(181, 10).Replace(",", "")))
                                rowSOTINVH2_CR.Item("ALLOW_INV_DISC") = Trim(line.Substring(190, 3))
                                rowSOTINVH2_CR.Item("PROD_LINE_CODE") = Trim(line.Substring(193, 10))
                                rowSOTINVH2_CR.Item("COMPANY_CODE") = Trim(line.Substring(203, 3))
                                dst.Tables("SOTINVH2_CR").Rows.Add(rowSOTINVH2_CR)
                            End If
                        End If
                    End If
                End If

            Next
        Next

        For Each rowSOTINVH1_CR As DataRow In dst.Tables("SOTINVH1_CR").Select
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow

            Dim INV_NO As String = rowSOTINVH1_CR.Item("CM_NO")
            Dim INV_REVERSAL As Boolean

            'rowSOTINVH1.Item("INV_NO") = rowSOTINVH1_CR.Item("INV_NO")
            rowSOTINVH1.Item("INV_NO") = Format(Val(rowSOTINVH1_CR.Item("CM_NO")), "0000000000")
            rowSOTINVH1.Item("CUST_CODE") = rowSOTINVH1_CR.Item("SELL_TO_CUST_NO")

            Dim CUST_STORE_NO As String = ""
            If rowSOTINVH1_CR.Item("EDI_ORDER") <> "Yes" Then
                If rowSOTINVH1_CR.Item("SHIP_TO_CODE") = "" Then
                    CUST_STORE_NO = "000000"
                Else
                    If Val(rowSOTINVH1_CR.Item("SHIP_TO_CODE")) > 0 Then
                        CUST_STORE_NO = Format(Val(rowSOTINVH1_CR.Item("SHIP_TO_CODE")), "000000").ToString.Substring(0, 6)
                    Else
                        CUST_STORE_NO = (rowSOTINVH1_CR.Item("SHIP_TO_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                End If
            Else
                'EDI Orders
                If rowSOTINVH1_CR.Item("SHIP_TO_CODE") = "" Then
                    CUST_STORE_NO = "000000"
                ElseIf Len(rowSOTINVH1_CR.Item("EDI_SHIP_FOR_CODE")) > 6 Then
                    If Val(rowSOTINVH1_CR.Item("SHIP_TO_CODE")) > 0 Then
                        CUST_STORE_NO = Format(Val(rowSOTINVH1_CR.Item("SHIP_TO_CODE")), "000000").ToString.Substring(0, 6)
                    Else
                        CUST_STORE_NO = (rowSOTINVH1_CR.Item("SHIP_TO_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                Else
                    If Val(rowSOTINVH1_CR.Item("EDI_SHIP_FOR_CODE")) > 0 Then
                        CUST_STORE_NO = Format(Val(rowSOTINVH1_CR.Item("EDI_SHIP_FOR_CODE")), "000000")
                    Else
                        CUST_STORE_NO = rowSOTINVH1_CR.Item("EDI_SHIP_FOR_CODE").ToString.Trim
                    End If
                End If
            End If
            If CUST_STORE_NO = "" Then
                CUST_STORE_NO = "000000"
            End If
            rowSOTINVH1.Item("CUST_STORE_NO") = CUST_STORE_NO
            rowSOTINVH1.Item("ORDR_CUST_PO") = rowSOTINVH1_CR.Item("EXTERNAL_DOC_NO")
            rowSOTINVH1.Item("INV_COMMENT") = rowSOTINVH1_CR.Item("PRE_ASSIGNED_NO")

            If Trim(rowSOTINVH1_CR.Item("LOCATION_CODE")) = "CAUD TX" Then
                rowSOTINVH1.Item("WHSE_CODE") = "TX"
            ElseIf Trim(rowSOTINVH1_CR.Item("LOCATION_CODE")) = "" Then
                rowSOTINVH1.Item("WHSE_CODE") = "PDR"
            Else
                rowSOTINVH1.Item("WHSE_CODE") = (rowSOTINVH1_CR.Item("LOCATION_CODE") & Space(6)).ToString.Substring(0, 6).Trim
            End If

            rowSOTINVH1.Item("ORDR_NO") = ""
            rowSOTINVH1.Item("POST_CODE") = ""
            If rowSOTINVH1_CR.Item("SALESPERSON_CODE") <> "" Then
                rowSOTINVH1.Item("SREP_CODE") = rowSOTINVH1_CR.Item("SALESPERSON_CODE")
            Else
                rowSOTINVH1.Item("SREP_CODE") = "000"
            End If

            rowSOTINVH1.Item("REASON_CODE") = ""
            rowSOTINVH1.Item("CUST_BILL_TO_CUST") = rowSOTINVH1_CR.Item("BILL_TO_CUST_NO")
            rowSOTINVH1.Item("BRAND_CODE") = ""
            rowSOTINVH1.Item("EVENT_CODE") = ""
            rowSOTINVH1.Item("INV_DATE") = rowSOTINVH1_CR.Item("DOCUMENT_DATE") 'POSTING_DATE?

            rowSOTINVH1.Item("CUST_BILL_TO_CUST") = rowSOTINVH1_CR.Item("BILL_TO_CUST_NO")

            rowSOTINVH1.Item("INIT_DATE") = rowSOTINVH1_CR.Item("POSTING_DATE")
            rowSOTINVH1.Item("INIT_OPER") = rowSOTINVH1_CR.Item("USER_ID")
            rowSOTINVH1.Item("CURR_CODE") = "USD"
            rowSOTINVH1.Item("CURR_EXCH_RATE") = "1"

            'rowSOTINVH1.Item("INV_DATE_SHIPPED") = 'rowSOTINVH1_CR.Item("SHIPMENT_DATE") ' SHIP_DATE?
            'rowSOTINVH1.Item("INV_CARTONS") = "" 'rowSOTINVH1_CR.Item("TOTAL_PKGS")
            'rowSOTINVH1.Item("INV_WEIGHT") = "" 'rowSOTINVH1_CR.Item("TOTAL_WT")
            'rowSOTINVH1.Item("INV_BOL_NO") = "" 'rowSOTINVH1_CR.Item("INVOICE_FOR_BOL")
            rowSOTINVH1.Item("INV_PRO_NO") = ""
            rowSOTINVH1.Item("SHIP_VIA_DESC") = ""
            rowSOTINVH1.Item("INV_PRO_NO") = ""
            rowSOTINVH1.Item("INV_NO_CONS") = ""
            rowSOTINVH1.Item("SHIPMENT_NO") = ""
            rowSOTINVH1.Item("INV_NO_REV") = ""
            If Len(rowSOTINVH1_CR.Item("SHIP_TO_STATE")) <= 2 And Len(rowSOTINVH1_CR.Item("SHIP_TO_STATE")) > 0 Then
                rowSOTINVH1.Item("CUST_SHIP_TO_STATE") = rowSOTINVH1_CR.Item("SHIP_TO_STATE")
                rowSOTINVH1.Item("SHIP_TO_CITY") = rowSOTINVH1_CR.Item("SHIP_TO_CITY")
            ElseIf Len(rowSOTINVH1_CR.Item("BILL_TO_STATE")) <= 2 And Len(rowSOTINVH1_CR.Item("BILL_TO_STATE")) > 0 Then
                ' SOME RETURNS DON'T HAVE SHIP-TO INFO use BILL-TO info
                rowSOTINVH1.Item("CUST_SHIP_TO_STATE") = rowSOTINVH1_CR.Item("BILL_TO_STATE")
                rowSOTINVH1.Item("SHIP_TO_CITY") = rowSOTINVH1_CR.Item("BILL_TO_CITY")
            Else
                rowSOTINVH1.Item("CUST_SHIP_TO_STATE") = rowSOTINVH1_CR.Item("SELL_TO_STATE")
                rowSOTINVH1.Item("SHIP_TO_CITY") = rowSOTINVH1_CR.Item("SELL_TO_CITY")
            End If
            'rowSOTINVH1.Item("INV_NO_REV_BY") = ""
            rowSOTINVH1.Item("INV_FREIGHT_TAX") = 0
            rowSOTINVH1.Item("SHIP_VIA_CODE") = ""
            rowSOTINVH1.Item("STAX_CODE") = ""
            rowSOTINVH1.Item("ORDER_TYPE_CODE") = ""
            rowSOTINVH1.Item("POSTRUN") = ""
            rowSOTINVH1.Item("TRANSID") = ""
            rowSOTINVH1.Item("BATCHID") = ""
            rowSOTINVH1.Item("INVCNUM") = ""
            rowSOTINVH1.Item("TRANSID") = ""

            If (rowSOTINVH1_CR.Item("EXTERNAL_DOC_NO") & Space(3)).ToString.Substring(0, 3) = "129" Or _
                (rowSOTINVH1_CR.Item("PRE_ASSIGNED_NO") & Space(3)).ToString.Substring(0, 3) = "129" Then
                INV_REVERSAL = True
            Else
                INV_REVERSAL = False
            End If

            Dim INV_SALES As Double = 0
            Dim INV_COGS As Double = 0
            Dim INV_FREIGHT As Double = 0
            Dim INV_STAX As Double = 0
            Dim INV_MISC_CHG As Double = 0
            Dim INV_TOTAL_AMOUNT As Double = 0

            For Each rowSOTINVH2_CR As DataRow In dst.Tables("SOTINVH2_CR").Select("CM_NO = '" & INV_NO & "'", "", DataViewRowState.CurrentRows)

                Dim LINE_AMOUNT As Double = rowSOTINVH2_CR.Item("LINE_AMT") '- rowSOTINVH2_CR.Item("LINE_DISC_AMT")
                'LINE AMOUNT IS ORDER QTY * PRICE (not shipped qty)

                If rowSOTINVH2_CR.Item("LINE_ITEM_TYPE") = "Item" Then
                    Dim UNIT_PRICE As Decimal = rowSOTINVH2_CR.Item("UNIT_PRICE") * _
                                                        ((100 - rowSOTINVH2_CR.Item("LINE_DISC_PCT")) * 0.01)
                    Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").NewRow
                    rowSOTINVH2.Item("INV_NO") = Format(Val(rowSOTINVH2_CR.Item("CM_NO")), "0000000000")
                    rowSOTINVH2.Item("INV_LNO") = rowSOTINVH2_CR.Item("CM_LNO")
                    rowSOTINVH2.Item("ITEM_CODE") = rowSOTINVH2_CR.Item("ITEM_CODE")
                    rowSOTINVH2.Item("ORDR_UNIT_PRICE") = rowSOTINVH2_CR.Item("UNIT_PRICE") * _
                                                        ((100 - rowSOTINVH2_CR.Item("LINE_DISC_PCT")) * 0.01)
                    rowSOTINVH2.Item("ITEM_RETAIL_PRICE") = rowSOTINVH2_CR.Item("UNIT_PRICE")
                    rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") = rowSOTINVH2_CR.Item("UNIT_PRICE") * _
                                                       ((100 - rowSOTINVH2_CR.Item("LINE_DISC_PCT")) * 0.01)
                    rowSOTINVH2.Item("ITEM_RETAIL_PRICE_CURR") = rowSOTINVH2_CR.Item("UNIT_PRICE")
                    rowSOTINVH2.Item("ORDR_QTY_SHIP") = rowSOTINVH2_CR.Item("QTY") * -1
                    rowSOTINVH2.Item("ITEM_UNIT_COST") = rowSOTINVH2_CR.Item("UNIT_COST")
                    rowSOTINVH2.Item("CUST_CODE") = rowSOTINVH1_CR.Item("SELL_TO_CUST_NO")
                    rowSOTINVH2.Item("CUST_STORE_NO") = CUST_STORE_NO
                    rowSOTINVH2.Item("SREP_CODE") = rowSOTINVH1_CR.Item("SALESPERSON_CODE")
                    rowSOTINVH2.Item("RET_REASON_CODE") = rowSOTINVH2_CR.Item("RET_REASON_CODE")
                    If Trim(rowSOTINVH1_CR.Item("LOCATION_CODE")) = "CAUD TX" Then
                        rowSOTINVH2.Item("WHSE_CODE") = "TX"
                    ElseIf Trim(rowSOTINVH1_CR.Item("LOCATION_CODE")) = "" Then
                        rowSOTINVH2.Item("WHSE_CODE") = "PDR"
                    Else
                        rowSOTINVH2.Item("WHSE_CODE") = (rowSOTINVH1_CR.Item("LOCATION_CODE") & Space(6)).ToString.Substring(0, 6).Trim
                    End If
                    'rowSOTINVH2.Item("POSTRUN") = ""
                    'rowSOTINVH2.Item("TRANSID") = ""
                    'rowSOTINVH2.Item("INVCNUM") = ""

                    INV_COGS = INV_COGS + (rowSOTINVH2_CR.Item("UNIT_COST") * (rowSOTINVH2_CR.Item("QTY") * -1))
                    INV_SALES = INV_SALES + (UNIT_PRICE * (rowSOTINVH2_CR.Item("QTY") * -1))
                    INV_TOTAL_AMOUNT = INV_TOTAL_AMOUNT + (UNIT_PRICE * (rowSOTINVH2_CR.Item("QTY") * -1))

                    rowSOTINVH2.Item("ORDR_YYYYPP_UPDATED") = "XXXXXX"
                    rowSOTINVH2.Item("GLACCTSALES") = "4000"

                    'INV TYPE
                    If INV_REVERSAL = True Then
                        rowSOTINVH2.Item("INV_TYPE") = "I"
                    Else
                        rowSOTINVH2.Item("INV_TYPE") = "C"
                    End If
                    dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)

                    'DO NOT WRITE SOTINVH2 FOR RESOURCE
                ElseIf rowSOTINVH2_CR.Item("LINE_ITEM_TYPE") = "Resource" Then
                    If rowSOTINVH2_CR.Item("ITEM_CODE") = "SHIPPING" Then
                        INV_FREIGHT = INV_FREIGHT + (LINE_AMOUNT * -1)
                    ElseIf rowSOTINVH2_CR.Item("ITEM_CODE") = "STAX" Then
                        INV_STAX = INV_STAX + (LINE_AMOUNT * -1)
                    End If
                Else
                    INV_MISC_CHG = INV_MISC_CHG + (LINE_AMOUNT * -1)
                End If

            Next

            rowSOTINVH1.Item("INV_SALES") = INV_SALES
            rowSOTINVH1.Item("INV_COGS") = INV_COGS
            rowSOTINVH1.Item("INV_FREIGHT") = INV_FREIGHT
            rowSOTINVH1.Item("INV_MISC_CHG") = INV_MISC_CHG
            rowSOTINVH1.Item("INV_STAX") = INV_STAX
            rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = INV_TOTAL_AMOUNT

            If INV_REVERSAL = True Then
                rowSOTINVH1.Item("INV_TYPE") = "I"
            Else
                rowSOTINVH1.Item("INV_TYPE") = "C"
            End If

            rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED") = "XXXXXX"
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)
        Next

        ASCMAIN1.Progress("", "")

    End Sub

    Sub Show_Unavailable_Only()

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = _
            DirectCast(tlb.Tools("Show Unavailable Only"), UltraWinToolbars.StateButtonTool)

        If Not tlb_sbt.Checked Then
            grdICTITEM1_IN.DataSource = dst.Tables("ICTITEM1_IN")
        Else
            grdICTITEM1_IN.DataSource = New DataView(dst.Tables("ICTITEM1_IN"), "ISNULL(QTY_AVAIL,0) < 0", "", DataViewRowState.CurrentRows)
        End If

    End Sub

    Sub Show_Detail_or_Summary()

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = _
            DirectCast(tlb.Tools("Show Summary Columns Only"), UltraWinToolbars.StateButtonTool)

        For Each s As String In New String() _
        {"ITEM_UOM", "COSTING_METHOD", "AVG_COST", "UNIT_COST", "ITEM_COST_STD", _
         "ITEM_DESC2", "GEN_PROD_POST_GRP", "INV_POST_GRP"}
            grdICTITEM1_IN.DisplayLayout.Bands(0).Columns(s).Hidden = tlb_sbt.Checked
        Next

        'If dst.Tables("POTORDG1") IsNot Nothing AndAlso dst.Tables("POTORDG1").Rows.Count > 0 Then
        '    grdPOTORDG2.DisplayLayout.Bands(0).Columns("PATIENT_NAME").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") <> "R"
        '    grdPOTORDG2.DisplayLayout.Bands(0).Columns("ORDR_NO").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") <> "R"
        '    grdPOTORDG2.DisplayLayout.Bands(0).Columns("ORDR_LNO").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") <> "R"
        '    grdPOTORDG2.DisplayLayout.Bands(0).Columns("ITEM_MIN_QTY").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") = "R"
        '    grdPOTORDG2.DisplayLayout.Bands(0).Columns("ITEM_MAX_QTY").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") = "R"
        '    grdPOTORDG2.DisplayLayout.Bands(0).Columns("WOS_CALC").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") = "R"
        'End If

        'If tlb_sbt.Checked Then
        '    Select Case MyBase.Absx1.optFor("PO_PARM_DEF_METHOD").Value
        '        Case "A"
        '            grdPOTORDG2.DisplayLayout.Bands(0).Columns("ABC_QTY").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") = "R" 'Or False
        '        Case "M"
        '            grdPOTORDG2.DisplayLayout.Bands(0).Columns("MNX_QTY").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") = "R" 'Or False
        '        Case "W"
        '            grdPOTORDG2.DisplayLayout.Bands(0).Columns("WOS_QTY").Hidden = dst.Tables("POTORDG1").Rows(0).Item("PO_ORDER_TYPE") = "R" 'Or False
        '    End Select
        'End If

    End Sub

    Private Sub grdICTITEM1_IN_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTITEM1_IN.InitializeRow
        Call Set_Item_BackColor(sender, e)
    End Sub

    Sub Set_Item_BackColor(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs)
        If e.Row.IsDataRow Then
            If Val(e.Row.Cells("QTY_AVAIL").Value & "") < 0 Then
                e.Row.Cells("QTY_AVAIL").Appearance.BackColor = Drawing.Color.Red
                e.Row.Cells("QTY_AVAIL").Appearance.ForeColor = Drawing.Color.White
            Else
                e.Row.Cells("QTY_AVAIL").Appearance.BackColor = Drawing.Color.Empty
            End If
        End If
    End Sub
    'Private Sub grdSOTORDR1_IN_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTORDR1_IN.AfterRowActivate
    '    With grdSOTORDR2_IN
    '        If .ActiveRow Is Nothing OrElse .ActiveRow.IsGroupByRow Then
    '            Exit Sub
    '        Else
    '            Dim ORDR_NO As String = .ActiveRow.Cells("ORDR_NO").Text
    '            grdSOTORDR2_IN.Text = "Details for Order" & ORDR_NO
    '            Call Fill_Records("SOTORDR2_IN", New Object() {ORDR_NO})
    '            Call Sort_grdColumns(grdSOTORDR2_IN, "ORDR_LNO")
    '            grdSOTORDR2_IN.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals for Order " & .ActiveRow.Cells("ORDR_NO").Text
    '        End If
    '        'Call Fill_Records("PPTLBKPP", New Object() _
    '        '    {HFs("CUST_CODE"), .ActiveRow.Cells("PYMT_BATCH_NO").Text, .ActiveRow.Cells("PYMT_BATCH_LNO").Text})

    '    End With
    'End Sub

    Private Sub grdSOTORDR1_IN_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTORDR1_IN.AfterRowActivate
        Dim ORDR_NO As String = grdSOTORDR1_IN.ActiveRow.Cells("ORDR_NO").Value & String.Empty
        viewSOTORDR2_IN.RowFilter = "ORDR_NO = '" & ORDR_NO & "'"
        If grdSOTORDR2_IN.Rows.Count > 0 Then
            grdSOTORDR2_IN.Text = "Details for Order No. : " & ORDR_NO
        Else
            grdSOTORDR2_IN.Text = String.Empty
        End If
    End Sub
    Private Sub grdSOTINVH1_IN_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTINVH1_IN.AfterRowActivate
        Dim INV_NO As String = grdSOTINVH1_IN.ActiveRow.Cells("INV_NO").Value & String.Empty
        viewSOTINVH2_IN.RowFilter = "INV_NO = '" & INV_NO & "'"
        If grdSOTINVH2_IN.Rows.Count > 0 Then
            grdSOTINVH2_IN.Text = "Details for Invoice No. : " & INV_NO
        Else
            grdSOTINVH2_IN.Text = String.Empty
        End If
    End Sub

    Private Sub grdSOTINVH1_CR_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSOTINVH1_CR.AfterRowActivate
        Dim INV_NO As String = grdSOTINVH1_CR.ActiveRow.Cells("CM_NO").Value & String.Empty
        viewSOTINVH2_CR.RowFilter = "CM_NO = '" & INV_NO & "'"
        If grdSOTINVH2_CR.Rows.Count > 0 Then
            grdSOTINVH2_CR.Text = "Details for Credit No. : " & INV_NO
        Else
            grdSOTINVH2_CR.Text = String.Empty
        End If
    End Sub

End Class