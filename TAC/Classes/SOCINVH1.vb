Public Class SOCINVH1

    Private tblSOTINVH1 As DataTable = Nothing
    Private tblSOTINVH2 As DataTable = Nothing

    Private tblSOTRTRN1 As DataTable = Nothing
    Private tblSOTRTRN2 As DataTable = Nothing

    Private tblSOTPICK1 As DataTable = Nothing
    Private tblSOTPICK2 As DataTable = Nothing

    Private tblARTOPEN1 As DataTable = Nothing
    Private tblSOTSHIP1 As DataTable = Nothing

    Private tblSOTORDR1ise As DataTable = Nothing
    Private tblSOTORDR2ise As DataTable = Nothing
    Private tblSOTORDR5 As DataTable = Nothing
    Private tblSOTSVIA1 As DataTable = Nothing

    Private tblTATTERM1 As DataTable = Nothing
    Private rowGLTPARM1 As DataRow

    Private frmASFBASE0 As New ASFBASE0

    Private dictYYYYWW As New Dictionary(Of String, String)
    Private dictYYYYMM As New Dictionary(Of String, String)

    ''' <summary>
    ''' used to create invoices
    ''' </summary>
    ''' <param name="SOTINVH1"></param>
    ''' <param name="SOTINVH2"></param>
    ''' <param name="SOTPICK1"></param>
    ''' <param name="SOTPICK2"></param>
    ''' <param name="ARTOPEN1"></param>
    ''' <param name="SOTSHIP1"></param>
    ''' <param name="SOTORDR5"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef SOTINVH1 As DataTable, _
                   ByRef SOTINVH2 As DataTable, _
                   ByRef SOTPICK1 As DataTable, _
                   ByRef SOTPICK2 As DataTable, _
                   ByRef ARTOPEN1 As DataTable, _
                   ByRef SOTSHIP1 As DataTable, _
                   ByRef SOTORDR5 As DataTable)

        tblSOTINVH1 = SOTINVH1
        tblSOTINVH2 = SOTINVH2
        tblSOTPICK1 = SOTPICK1
        tblSOTPICK2 = SOTPICK2
        tblARTOPEN1 = ARTOPEN1
        tblSOTSHIP1 = SOTSHIP1
        tblSOTORDR5 = SOTORDR5

        InitializeClassVariables()
    End Sub

    ''' <summary>
    ''' Used to Create Credits from returns
    ''' </summary>
    ''' <param name="SOTINVH1"></param>
    ''' <param name="SOTINVH2"></param>
    ''' <param name="SOTRTRN1"></param>
    ''' <param name="SOTRTRN2"></param>
    ''' <param name="ARTOPEN1"></param>
    ''' <remarks></remarks>
    Public Sub New(ByRef SOTINVH1 As DataTable, _
                   ByRef SOTINVH2 As DataTable, _
                   ByRef SOTRTRN1 As DataTable, _
                   ByRef SOTRTRN2 As DataTable, _
                   ByRef ARTOPEN1 As DataTable)


        tblSOTINVH1 = SOTINVH1
        tblSOTINVH2 = SOTINVH2
        tblSOTRTRN1 = SOTRTRN1
        tblSOTRTRN2 = SOTRTRN2
        tblARTOPEN1 = ARTOPEN1

        InitializeClassVariables()

    End Sub

    ''' <summary>
    ''' used to create invoices
    ''' </summary>
    ''' <param name="SOTINVH1"></param>
    ''' <param name="SOTINVH2"></param>
    ''' <param name="SOTPICK1"></param>
    ''' <param name="SOTPICK2"></param>
    ''' <param name="ARTOPEN1"></param>
    ''' <param name="SOTSHIP1"></param>
    ''' <param name="SOTORDR5"></param>
    ''' <param name="SOTORDR1">If set and the SOTORDR1 entry for the ordr_no is sotpick1 cannot be found then this will be used</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef SOTINVH1 As DataTable, _
                   ByRef SOTINVH2 As DataTable, _
                   ByRef SOTPICK1 As DataTable, _
                   ByRef SOTPICK2 As DataTable, _
                   ByRef ARTOPEN1 As DataTable, _
                   ByRef SOTSHIP1 As DataTable, _
                   ByRef SOTORDR5 As DataTable, _
                   ByRef SOTORDR1 As DataTable, _
                   ByRef SOTORDR2 As DataTable)

        tblSOTINVH1 = SOTINVH1
        tblSOTINVH2 = SOTINVH2
        tblSOTPICK1 = SOTPICK1
        tblSOTPICK2 = SOTPICK2
        tblARTOPEN1 = ARTOPEN1
        tblSOTSHIP1 = SOTSHIP1
        tblSOTORDR5 = SOTORDR5
        tblSOTORDR1ise = SOTORDR1
        tblSOTORDR2ise = SOTORDR2

        InitializeClassVariables()
    End Sub

    ''' <summary>
    ''' used to create invoices
    ''' </summary>
    ''' <param name="SOTINVH1"></param>
    ''' <param name="SOTINVH2"></param>
    ''' <param name="SOTPICK1"></param>
    ''' <param name="SOTPICK2"></param>
    ''' <param name="ARTOPEN1"></param>
    ''' <param name="SOTSHIP1"></param>
    ''' <param name="SOTORDR5"></param>
    ''' <param name="SOTORDR1">If set and the SOTORDR1 entry for the ordr_no is sotpick1 cannot be found then this will be used</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef SOTINVH1 As DataTable, _
                   ByRef SOTINVH2 As DataTable, _
                   ByRef SOTPICK1 As DataTable, _
                   ByRef SOTPICK2 As DataTable, _
                   ByRef ARTOPEN1 As DataTable, _
                   ByRef SOTSHIP1 As DataTable, _
                   ByRef SOTORDR5 As DataTable, _
                   ByRef SOTORDR1 As DataTable)

        tblSOTINVH1 = SOTINVH1
        tblSOTINVH2 = SOTINVH2
        tblSOTPICK1 = SOTPICK1
        tblSOTPICK2 = SOTPICK2
        tblARTOPEN1 = ARTOPEN1
        tblSOTSHIP1 = SOTSHIP1
        tblSOTORDR5 = SOTORDR5
        tblSOTORDR1ise = SOTORDR1

        InitializeClassVariables()
    End Sub

    ''' <summary>
    ''' This would be used if you are going to use CreateConsolidatedInvoice Only
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New()

    End Sub

    Private Sub InitializeClassVariables()
        rowGLTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM1 WHERE GL_PARM_KEY = 'Z'")
        tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT * FROM SOTSVIA1", "SOTSVIA1")
        tblTATTERM1 = ASCDATA1.GetDataTable("SELECT * FROM TATTERM1", "TATTERM1")

        frmASFBASE0.ROWs = New Dictionary(Of String, DataRow)
        frmASFBASE0.ROWs.Add("GLTPARM1", ASCDATA1.GetDataRow("SELECT * FROM GLTPARM1 WHERE GL_PARM_KEY = 'Z'"))

        dictYYYYWW = New Dictionary(Of String, String)
        dictYYYYMM = New Dictionary(Of String, String)
    End Sub

    Public Function CreateInvoices(ByVal SHIP_BOL_NO As String) As Int16
        Return CreateInvoices(SHIP_BOL_NO, String.Empty)
    End Function

    Public Function CreateInvoices(ByVal SHIP_BOL_NO As String, ByVal CUST_CODE As String) As Int16

        Dim rowSOTINVH1 As DataRow = Nothing
        Dim rowSOTINVH2 As DataRow = Nothing
        Dim rowARTCUST1 As DataRow = Nothing
        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim tblEDTTRPMC As DataTable = Nothing

        Dim WHSE_CODE As String = String.Empty
        Dim numInvoices As Int16 = 0
        Dim ORDR_GROUP_NO As String = String.Empty
        Dim CURR_CODE As String = String.Empty
        Dim CURR_EXCH_RATE As Decimal = 1
        Dim INV_NO As String = String.Empty
        Dim PPA_FREIGHT As Decimal = 0

        Dim edi_customer As Boolean = False
        Dim edi856_customer As Boolean = False
        Dim isCustConsInv As Boolean = False
        Dim INV_NO_CONS As String = String.Empty
        Dim INV_TYPE As String = "I"
        Dim foreignExchange As Boolean = False

        rowSOTSHIP1 = tblSOTSHIP1.Rows.Find(SHIP_BOL_NO)
        If rowSOTSHIP1 Is Nothing Then
            Return numInvoices
        End If

        WHSE_CODE = rowSOTSHIP1.Item("WHSE_CODE")
        ORDR_GROUP_NO = rowSOTSHIP1.Item("ORDR_GROUP_NO")

        If CUST_CODE.Length = 0 Then
            Dim rowSOTORODR0 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO = :PARM1", "V", New Object() {ORDR_GROUP_NO})
            If rowSOTORODR0 Is Nothing Then
                Return numInvoices
            End If
            CUST_CODE = rowSOTORODR0.Item("CUST_CODE")
        End If

        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
        ' This means the customer has consolidated invoices by a shipment SOTSHIP1.SHIP_BOL_NO
        If rowARTCUST1.Item("CUST_CONS_INV") & String.Empty = "1" Then
            isCustConsInv = True
        End If

        tblEDTTRPMC = ASCDATA1.GetDataTable("Select * From EDTTRPM1 Where CUST_CODE = :PARM1", "EDTTRPMC", "V", New Object() {CUST_CODE})
        edi_customer = tblEDTTRPMC.Rows.Count <> 0
        edi856_customer = tblEDTTRPMC.Select("EDI_STATUS = 'P' AND EDI_DOC_NO = '856' AND CUST_CODE = '" & CUST_CODE & "'").Length > 0

        Dim GL_PARM_CURR_CODE As String = rowGLTPARM1.Item("GL_PARM_CURR_CODE") & ""
        CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
        If CURR_CODE = "" Or CURR_CODE = GL_PARM_CURR_CODE Then
            CURR_CODE = GL_PARM_CURR_CODE
        Else
            foreignExchange = True
        End If

        ' Standard routine to get the Currency Exchange Rate
        If foreignExchange Then
            CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(frmASFBASE0, CURR_CODE, CDate(DateTime.Now.ToString("MM/dd/yyyy")))
        End If

        Dim ORDR_YYYYPP_UPDATED As String = ConvertDateToPeriod(rowSOTSHIP1.Item("INV_DATE"))

        If Not IsDate(rowSOTSHIP1.Item("INV_DATE") & String.Empty) Then
            rowSOTSHIP1.Item("INV_DATE") = CDate(DateTime.Now.ToString("MM/dd/yyyy"))
        End If
        Dim YYYYWW As String = ConvertDateToWeek(rowSOTSHIP1.Item("INV_DATE"))

        For Each rowSOTPICK1 As DataRow In tblSOTPICK1.Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "' and PICK_STATUS <> 'C'", "PICK_NO")

            Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
            Dim PICK_QTY_CONF As Int32 = Val(tblSOTPICK2.Compute("SUM(PICK_QTY_CONF)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
            If PICK_QTY_CONF = 0 Then Continue For

            Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
            Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO = :PARM1", "V", New Object() {ORDR_NO})
            Dim tblSOTORDR2 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR2 WHERE ORDR_NO = :PARM1", "SOTORDR2", "V", New Object() {ORDR_NO})
            Dim tblICTITEM1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN (SELECT ITEM_CODE FROM SOTORDR2 WHERE ORDR_NO = :PARM1)", "ICTITEM1", "V", New Object() {ORDR_NO})

            ' Need data for Inventory Shipment Entry ICFSHIP1
            If tblSOTORDR2.Rows.Count = 0 Then
                If tblSOTORDR2ise.Select("ORDR_NO = '" & ORDR_NO & "'").Length > 0 Then
                    tblSOTORDR2 = tblSOTORDR2ise

                    Dim items As String = String.Empty
                    For Each row As DataRow In tblSOTORDR2ise.Select("")
                        items &= ", '" & row.Item("ITEM_CODE") & "'"
                    Next
                    items = items.Substring(1).Trim
                    tblICTITEM1 = ASCDATA1.GetDataTable("SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN (" & items & ")")
                End If
            End If


            ' Need row for Inventory Shipment Entry ICFSHIP1
            If rowSOTORDR1 Is Nothing AndAlso tblSOTORDR1ise IsNot Nothing Then
                If tblSOTORDR1ise.Select("ORDR_NO = '" & ORDR_NO & "'").Length > 0 Then
                    rowSOTORDR1 = tblSOTORDR1ise.Select("ORDR_NO = '" & ORDR_NO & "'")(0)
                End If
            End If

            If Val(rowSOTORDR1.Item("ORDR_STAX") & String.Empty) = 0 AndAlso _
                (rowSOTPICK1.Table.Columns.Contains("ORDR_STAX") AndAlso Val(rowSOTPICK1.Item("ORDR_STAX") & String.Empty) > 0) Then
                rowSOTORDR1.Item("ORDR_STAX") = rowSOTPICK1.Item("ORDR_STAX")
                rowSOTORDR1.Item("STAX_CODE") = rowSOTPICK1.Item("STAX_CODE")
            End If

            Dim SALES_DIVISION_CODE As String = rowSOTPICK1.Item("SALES_DIVISION_CODE") & String.Empty
            INV_NO = rowSOTPICK1.Item("INV_NO") & String.Empty

            If INV_NO.Length = 0 Then
                INV_NO = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
            End If

            ' Use the first Invoice as the consolidated invoice
            If isCustConsInv AndAlso INV_NO_CONS.Length = 0 Then
                INV_NO_CONS = INV_NO
            End If

            Dim INV_COGS As Decimal = 0
            Dim INV_SALES As Decimal = 0

            For Each rowSOTPICK2 As DataRow In tblSOTPICK2.Select("PICK_NO = '" & PICK_NO & "'")
                Dim ORDR_QTY_SHIP As Int32 = Val(rowSOTPICK2.Item("PICK_QTY_CONF") & "")
                If ORDR_QTY_SHIP = 0 Then Continue For

                Dim ITEM_CODE As String = rowSOTPICK2.Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = tblICTITEM1.Rows.Find(New Object() {ITEM_CODE})
                Dim ORDR_UNIT_COST As Decimal = Val(rowICTITEM1.Item("ITEM_COST_STD") & "")

                rowSOTINVH2 = tblSOTINVH2.NewRow
                Dim rowSOTORDR2 As DataRow = tblSOTORDR2.Rows.Find(New Object() {rowSOTPICK2.Item("ORDR_NO"), rowSOTPICK2.Item("ORDR_LNO")})
                With rowSOTINVH2
                    ' If Foreign Exchange then use Currency Prices 
                    ' As per Walter on 6/15/2016 - this is how to calculate the Order Unit Price
                    .Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR")
                    If foreignExchange Then
                        .Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR") * CURR_EXCH_RATE
                    Else
                        .Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                    End If

                    INV_SALES += ORDR_QTY_SHIP * .Item("ORDR_UNIT_PRICE")

                    .Item("INV_TYPE") = INV_TYPE
                    .Item("INV_NO") = INV_NO
                    .Item("INV_LNO") = rowSOTPICK2.Item("PICK_LNO")
                    .Item("ITEM_CODE") = ITEM_CODE
                    .Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP

                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = rowSOTPICK1.Item("CUST_STORE_NO")
                    .Item("WHSE_CODE") = WHSE_CODE
                    .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE") 'rowSOTSHIP1.Item("SREP_CODE")
                    .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP

                    .Item("ITEM_UNIT_COST") = ORDR_UNIT_COST

                    ' Use the Item Price from the sales order, not Item Master.
                    .Item("ITEM_RETAIL_PRICE_CURR") = rowSOTORDR2.Item("ITEM_RETAIL_PRICE_CURR")
                    .Item("ITEM_RETAIL_PRICE") = rowSOTORDR2.Item("ITEM_RETAIL_PRICE")
                    .Item("OPS_YYYYWW") = YYYYWW 'ASCMAIN1.CYW

                    ' Added on 01/18/2016
                    .Item("SELL_CODE") = rowSOTORDR1.Item("SELL_CODE")

                    INV_COGS += (ORDR_QTY_SHIP * ORDR_UNIT_COST)
                End With
                tblSOTINVH2.Rows.Add(rowSOTINVH2)
            Next

            rowSOTINVH1 = tblSOTINVH1.NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = INV_TYPE
                .Item("INV_NO") = INV_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = rowSOTPICK1.Item("CUST_STORE_NO")
                .Item("ORDR_CUST_PO") = rowSOTPICK1.Item("ORDR_CUST_PO")
                .Item("ORDR_NO") = rowSOTPICK1.Item("ORDR_NO")
                .Item("WHSE_CODE") = rowSOTSHIP1.Item("WHSE_CODE")
                .Item("POST_CODE") = rowSOTPICK1.Item("POST_CODE")
                .Item("TERM_CODE") = rowSOTSHIP1.Item("TERM_CODE")
                .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE") ' rowSOTSHIP1.Item("SREP_CODE")
                .Item("SREP2_CODE") = rowSOTSHIP1.Item("SREP2_CODE")
                .Item("REASON_CODE") = "SHP"
                .Item("CUST_BILL_TO_CUST") = rowSOTORDR1.Item("CUST_BILL_TO_CUST")
                .Item("ORDR_CREDIT_CARD_TYPE") = rowSOTORDR1.Item("ORDR_CREDIT_CARD_TYPE")

                ' Added on 01/18/2016
                .Item("SELL_CODE") = rowSOTORDR1.Item("SELL_CODE")

                .Item("BRAND_CODE") = String.Empty
                .Item("EVENT_CODE") = rowSOTORDR1.Item("EVENT_CODE") & String.Empty

                .Item("INV_SALES") = Math.Round(INV_SALES, 2, MidpointRounding.AwayFromZero)
                .Item("INV_COGS") = INV_COGS

                PPA_FREIGHT = 0
                If rowSOTPICK1.Table.Columns.Contains("PPA_FREIGHT") Then
                    PPA_FREIGHT = Val(rowSOTPICK1.Item("PPA_FREIGHT") & String.Empty)
                End If

                '****************************************************
                ' New fields as of 01/13/2013
                .Item("REGISTER_IND") = "0"
                .Item("SHIP_FRT_AMT_ACTUAL") = 0
                If rowSOTSHIP1.Item("SHIP_810_IND") & String.Empty = "1" Then
                    .Item("EDI_INV_IND") = "1"
                End If

                .Item("CCPA_NO") = rowSOTPICK1.Item("CCPA_NO")

                ' Transaction Number from CC sale, use this since we have this for both Web and in house charges
                .Item("CC_SALE_TRANS_ID") = rowSOTORDR1.Item("CC_TRANS_ID") & String.Empty

                ' 12/18/2015 - Used for Amazon on Web Import
                .Item("PARTNER_ORDR_NO") = rowSOTORDR1.Item("PARTNER_ORDR_NO") & String.Empty

                .Item("STAX_RATE") = rowSOTORDR1.Item("STAX_RATE")
                .Item("STAX_CODE") = rowSOTORDR1.Item("STAX_CODE")

                ' Added 6/9/2014 as per Walter
                .Item("INV_FOB") = rowSOTSHIP1.Item("SHIP_FOB")

                ' This is our freight cost returned from the Shipper (Fedex, UPS)
                .Item("SHIP_FRT_AMT_ACCRUED") = 0
                If rowSOTPICK1.Table.Columns.Contains("OUR_FREIGHT") Then
                    .Item("SHIP_FRT_AMT_ACCRUED") = Val(rowSOTPICK1.Item("OUR_FREIGHT") & String.Empty)
                End If
                '****************************************************

                ' 10/21/2025
                ' If not PPA then set the INV_FREIGHT = 0
                .Item("INV_FREIGHT") = 0
                If rowSOTSHIP1 IsNot Nothing Then
                    If rowSOTSHIP1.Item("FRT_TERMS") & String.Empty = "PPA" Then
                        .Item("INV_FREIGHT") = Val(rowSOTPICK1.Item("PICK_FREIGHT") & String.Empty) + Val(rowSOTPICK1.Item("ORDR_FREIGHT") & String.Empty) + PPA_FREIGHT
                    End If
                End If

                .Item("INV_MISC_CHG") = Val(rowSOTORDR1.Item("ORDR_MISC_CHG") & String.Empty)
                .Item("INV_STAX") = Val(rowSOTORDR1.Item("ORDR_STAX") & String.Empty)
                .Item("INV_TOTAL_AMOUNT") = Val(.Item("INV_SALES") & String.Empty) + Val(.Item("INV_FREIGHT") & String.Empty) + Val(.Item("INV_STAX") & String.Empty) + Val(.Item("INV_MISC_CHG") & String.Empty)

                If Not IsDate(rowSOTSHIP1.Item("INV_DATE") & String.Empty) Then
                    rowSOTSHIP1.Item("INV_DATE") = CDate(DateTime.Now.ToString("MM/dd/yyyy"))
                End If

                .Item("INV_DATE") = CDate(rowSOTSHIP1.Item("INV_DATE")).ToShortDateString
                .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                .Item("INIT_DATE") = DateTime.Now
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("REGISTER_XNO") = String.Empty
                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                .Item("INV_DATE_SHIPPED") = CDate(rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")).ToShortDateString

                ' If Send Invoices is set to Email or None then flag Invoice as printed. Also if Web Order then set to printed.
                If rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty = "E" OrElse rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty = "N" Then
                    .Item("INV_PRINTED") = "1"
                ElseIf rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty = "W" Then
                    .Item("INV_PRINTED") = "1"
                ElseIf .Item("INV_TOTAL_AMOUNT") = 0 Then
                    .Item("INV_PRINTED") = "1"
                End If

                .Item("INV_UNITS") = Val(tblSOTPICK2.Compute("SUM(PICK_QTY_CONF)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                .Item("INV_CARTONS") = Val(rowSOTPICK1.Item("PICK_CNT_CARTONS") & String.Empty)
                .Item("INV_WEIGHT") = Val(rowSOTPICK1.Item("PICK_TOTAL_WGT") & String.Empty)
                .Item("INV_BOL_NO") = rowSOTSHIP1.Item("BILL_OF_LADING_NO")
                .Item("INV_PRO_NO") = rowSOTSHIP1.Item("SHIP_REF")
                .Item("SHIP_VIA_DESC") = tblSOTSVIA1.Rows.Find(rowSOTSHIP1.Item("SHIP_VIA_CODE")).Item("SHIP_VIA_DESC")
                .Item("INV_NO_CONS") = INV_NO_CONS
                .Item("SHIP_BOL_NO") = rowSOTPICK1.Item("SHIP_BOL_NO")
                .Item("PICK_NO") = PICK_NO
                If tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'").Length > 0 Then
                    .Item("CUST_SHIP_TO_STATE") = tblSOTORDR5.Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")(0).Item("CUST_STATE") & String.Empty
                End If
                .Item("INV_COMMENT") = rowSOTPICK1.Item("ORDR_INV_COMMENT")
                '.Item("INV_FREIGHT_TAX") = 0
                .Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE")
                .Item("ORDR_TYPE_CODE") = rowSOTORDR1.Item("ORDR_TYPE_CODE")
                .Item("OPS_YYYYWW") = YYYYWW 'ASCMAIN1.CYW
                .Item("ORDR_DEPT") = rowSOTORDR1.Item("ORDR_DEPT")
                .Item("CUST_FACTOR_IND") = rowARTCUST1.Item("CUST_FACTOR_IND") & String.Empty
                .Item("ORDR_NO_WEB") = rowSOTORDR1.Item("ORDR_NO_WEB")
                .Item("SALES_DIVISION_CODE") = rowSOTPICK1.Item("SALES_DIVISION_CODE")

                ' Used to send tracking data to Web
                If rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty = "W" Then
                    .Item("ORDR_WEB_IND") = "1"
                End If

            End With
            tblSOTINVH1.Rows.Add(rowSOTINVH1)
            numInvoices += 1
            CreateOpenAR(INV_TYPE, INV_NO, CURR_EXCH_RATE)

            rowSOTPICK1.Item("INV_NO") = INV_NO
            rowSOTPICK1.Item("PICK_SHIPPED") = rowSOTSHIP1.Item("SHIP_DATE_SHIPPED")
            ' Set this outside the class. Some forms may update SOTPICK1, then do an update to other table based on PICK_STATUS
            'rowSOTPICK1.Item("PICK_STATUS") = "F"
            rowSOTPICK1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSOTPICK1.Item("LAST_DATE") = DateTime.Now
        Next

        Return numInvoices

    End Function

    ''' <summary>
    ''' Creates the Open AR Record for the Invoice
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CreateOpenAR(ByVal INV_TYPE As String, ByVal INV_NO As String, ByVal CURR_EXCH_RATE As Decimal)

        Dim rowARTOPEN1 As DataRow = tblARTOPEN1.NewRow
        Dim rowSOTINVH1 As DataRow = tblSOTINVH1.Rows.Find(New Object() {INV_TYPE, INV_NO})

        Dim INV_SALES_CURR As Decimal = 0
        For Each row As DataRow In tblSOTINVH2.Select("INV_TYPE = '" & INV_TYPE & "' and INV_NO = '" & INV_NO & "'")
            INV_SALES_CURR += (Val(row.Item("ORDR_QTY_SHIP") & String.Empty) * Val(row.Item("ORDR_UNIT_PRICE_CURR") & String.Empty))
        Next

        Dim foreignExchange As Boolean = False
        Dim GL_PARM_CURR_CODE As String = rowGLTPARM1.Item("GL_PARM_CURR_CODE") & ""
        Dim CURR_CODE As String = rowSOTINVH1.Item("CURR_CODE") & ""
        If CURR_CODE = "" OrElse CURR_CODE = GL_PARM_CURR_CODE Then
            CURR_CODE = GL_PARM_CURR_CODE
        Else
            foreignExchange = True
        End If

        ' 12/18/2015 - Used for Amazon on Web Import - PARTNER_ORDR_NO
        For Each fieldName As String In New String() _
                {"CUST_CODE", "INV_TYPE", "INV_DATE", "CUST_STORE_NO", "POST_CODE", _
                 "TERM_CODE", "SREP_CODE", "SREP2_CODE", "STAX_CODE", "ORDR_TYPE_CODE", _
                 "ORDR_NO", "INV_SALES", "INV_FREIGHT", "INV_STAX", "INV_TOTAL_AMOUNT", _
                 "REASON_CODE", "INIT_OPER", "INIT_DATE", "INV_MISC_CHG", "ORDR_TYPE_CODE", "SALES_DIVISION_CODE", "PARTNER_ORDR_NO", "INV_NO_CONS", "ORDR_CREDIT_CARD_TYPE"}
            rowARTOPEN1.Item(fieldName) = rowSOTINVH1.Item(fieldName)
        Next

        ' Added 11/2/2015
        If rowSOTINVH1.Item("CUST_BILL_TO_CUST") & String.Empty <> String.Empty Then
            rowARTOPEN1.Item("CUST_CODE") = rowSOTINVH1.Item("CUST_BILL_TO_CUST")
        End If

        rowARTOPEN1.Item("INV_TYPE") = INV_TYPE
        rowARTOPEN1.Item("INV_NUM") = rowSOTINVH1.Item("INV_NO")

        Dim INV_DUE_DATE As Date = SOCMAIN1.Calculate_INV_DUE_DATE(Nothing, rowSOTINVH1.Item("TERM_CODE") & String.Empty, rowSOTINVH1.Item("INV_DATE"))
        rowARTOPEN1.Item("INV_DUE_DATE") = INV_DUE_DATE.ToShortDateString
        rowARTOPEN1.Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO")
        rowARTOPEN1.Item("INV_BALANCE") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
        rowARTOPEN1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowARTOPEN1.Item("LAST_DATE") = DateTime.Now + ASCMAIN1.NowTSD

        rowARTOPEN1.Item("CUST_CODE_SO") = rowSOTINVH1.Item("CUST_CODE")
        rowARTOPEN1.Item("SEG2_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG2")
        rowARTOPEN1.Item("SEG3_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG3")
        rowARTOPEN1.Item("SEG4_CODE") = rowGLTPARM1.Item("GL_PARM_DEF_SEG4")
        rowARTOPEN1.Item("CURR_CODE") = rowGLTPARM1.Item("GL_PARM_CURR_CODE")
        rowARTOPEN1.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE

        rowARTOPEN1.Item("INV_SALES_CURR") = INV_SALES_CURR ' Math.Round(rowARTOPEN1.Item("INV_SALES") / CURR_EXCH_RATE, 2, MidpointRounding.AwayFromZero)

        rowARTOPEN1.Item("INV_DISC") = 0
        rowARTOPEN1.Item("INV_PMT") = 0
        rowARTOPEN1.Item("INV_DISC_TAKEN") = 0
        rowARTOPEN1.Item("INV_WRITE_OFF") = 0

        rowARTOPEN1.Item("INV_PMT_CURR") = 0
        rowARTOPEN1.Item("INV_DISC_TAKEN_CURR") = 0
        rowARTOPEN1.Item("INV_WRITE_OFF_CURR") = 0
        rowARTOPEN1.Item("OPS_YYYYPP") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")

        If foreignExchange Then
            If Val(rowARTOPEN1.Item("INV_DISC") & String.Empty) <> 0 Then
                Throw New Exception("Invoice Discount for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            End If
            If Val(rowARTOPEN1.Item("INV_FREIGHT") & String.Empty) <> 0 Then
                Throw New Exception("Invoice Freight for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            End If
            If Val(rowARTOPEN1.Item("INV_STAX") & String.Empty) <> 0 Then
                Throw New Exception("Invoice Sales Tax for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            End If
            If Val(rowARTOPEN1.Item("INV_MISC_CHG") & String.Empty) <> 0 Then
                Throw New Exception("Invoice Misc Charge for Foreign Currency <> 0. Consult ABS to see how this is to be handled.")
            End If

            rowARTOPEN1.Item("INV_DISC_CURR") = 0 'Math.Round(rowARTOPEN1.Item("INV_DISC") / CURR_EXCH_RATE, 2, MidpointRounding.AwayFromZero)
            rowARTOPEN1.Item("INV_FREIGHT_CURR") = 0 'Math.Round(rowARTOPEN1.Item("INV_FREIGHT") / CURR_EXCH_RATE, 2, MidpointRounding.AwayFromZero)
            rowARTOPEN1.Item("INV_STAX_CURR") = 0 'Math.Round(rowARTOPEN1.Item("INV_STAX") / CURR_EXCH_RATE, 2, MidpointRounding.AwayFromZero)
            rowARTOPEN1.Item("INV_MISC_CHG_CURR") = 0 'Math.Round(rowARTOPEN1.Item("INV_MISC_CHG") / CURR_EXCH_RATE, 2, MidpointRounding.AwayFromZero)
            rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_SALES_CURR") ' Math.Round(rowARTOPEN1.Item("INV_BALANCE") / CURR_EXCH_RATE, 2, MidpointRounding.AwayFromZero)
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowARTOPEN1.Item("INV_SALES_CURR") ' Math.Round(rowARTOPEN1.Item("INV_TOTAL_AMOUNT") / CURR_EXCH_RATE, 2, MidpointRounding.AwayFromZero)
        Else
            rowARTOPEN1.Item("INV_DISC_CURR") = rowARTOPEN1.Item("INV_DISC")
            rowARTOPEN1.Item("INV_FREIGHT_CURR") = rowARTOPEN1.Item("INV_FREIGHT")
            rowARTOPEN1.Item("INV_STAX_CURR") = rowARTOPEN1.Item("INV_STAX")
            rowARTOPEN1.Item("INV_MISC_CHG_CURR") = rowARTOPEN1.Item("INV_MISC_CHG")
            rowARTOPEN1.Item("INV_BALANCE_CURR") = rowARTOPEN1.Item("INV_BALANCE")
            rowARTOPEN1.Item("INV_TOTAL_AMOUNT_CURR") = rowARTOPEN1.Item("INV_TOTAL_AMOUNT")
        End If

        Dim rowTATTERM1 As DataRow = tblTATTERM1.Rows.Find(rowARTOPEN1.Item("TERM_CODE") & String.Empty)
        If rowTATTERM1 IsNot Nothing _
            AndAlso Val(rowTATTERM1.Item("TERM_DISC_PERC") & String.Empty) > 0 _
            AndAlso Val(rowTATTERM1.Item("TERM_DAYS_DISC") & String.Empty) > 0 Then

            rowARTOPEN1.Item("INV_DISC_DATE") = DateAdd(DateInterval.Day, Val(rowTATTERM1.Item("TERM_DAYS_DISC") & String.Empty), rowARTOPEN1.Item("INV_DATE")).ToShortDateString
        End If

        tblARTOPEN1.Rows.Add(rowARTOPEN1)

        ' Other Columns in ARTOPEN1 that are not filled by this procedure
        'rowARTOPEN1.Item("INV_DISC_DATE") = String.Empty
        'rowARTOPEN1.Item("APPLY_TO_INV_NUM") = String.Empty
        'rowARTOPEN1.Item("APPLY_TO_INV_TYPE") = String.Empty
        'rowARTOPEN1.Item("INV_LAST_PMT") = String.Empty
        'rowARTOPEN1.Item("INV_PMT") = String.Empty
        'rowARTOPEN1.Item("INV_DISC_TAKEN") = String.Empty
        'rowARTOPEN1.Item("INV_WRITE_OFF") = String.Empty
        'rowARTOPEN1.Item("INV_LAST_PMT_REF") = String.Empty
        'rowARTOPEN1.Item("INV_LAST_PMT_REF_DT") = String.Empty
        'rowARTOPEN1.Item("OPS_YYYYPP_F") = String.Empty
        'rowARTOPEN1.Item("GST_TAX") = String.Empty
        'rowARTOPEN1.Item("GST_TAX_CURR") = String.Empty
        'rowARTOPEN1.Item("ORDR_CREDIT_APPR_BY") = String.Empty
        'rowARTOPEN1.Item("ORDR_CREDIT_APPR_DATE") = String.Empty
        'rowARTOPEN1.Item("INV_NOTES") = String.Empty
        'rowARTOPEN1.Item("OPS_YYYYPP_PAID") = String.Empty

    End Sub

    ''' <summary>
    ''' Get all invoices where the Invoice Number or Consolidated Invoice Nuumber match the invoiceNumber provided
    ''' Creates a SOTINVH2 header, and SOTINVH2 detaail using the Consolidated Invoice Number
    ''' </summary>
    ''' <param name="invoiceNumber">Invoice number used to get all invoices in teh consolidation</param>
    ''' <param name="rowSOTINVH1">reference to SOTINVH1 datarow</param>
    ''' <param name="tblSOTINVH2">reference to SOTINVH2 datatable</param>
    ''' <returns>returns True if no errors; otherwise, returns false</returns>
    ''' <remarks></remarks>
    Public Function  CreateConsolidatedInvoice(ByVal invoiceNumber As String, ByRef rowSOTINVH1 As DataRow, ByRef tblSOTINVH2 As DataTable) As boolean

        Try

            ' Row and table must have the primay key fields
            If Not rowSOTINVH1.Table.Columns.Contains("INV_NO") Then
                Return False
            End If

            If Not rowSOTINVH1.Table.Columns.Contains("INV_TYPE") Then
                Return False
            End If

            If Not tblSOTINVH2.Columns.Contains("INV_NO") Then
                Return False
            End If

            If Not tblSOTINVH2.Columns.Contains("INV_TYPE") Then
                Return False
            End If

            Dim sqlInvoices As String = "Select * from Sotinvh1 where Inv_no = :PARM1 or INV_NO_CONS = :PARM2"

            Dim tblHeader As DataTable = ASCDATA1.GetDataTable(sqlInvoices, "SOTINVH1", _
                                                               "VV", New Object() {invoiceNumber, invoiceNumber})

            tblHeader.PrimaryKey = New System.Data.DataColumn() {tblHeader.Columns("INV_NO")}

            If tblHeader.Rows.Count <= 1 Then
                Return True
            End If


            Dim tblDetails As DataTable = ASCDATA1.GetDataTable("Select * from Sotinvh2 where Inv_no in (" & sqlInvoices.Replace("*", "INV_NO") & ")", "SOTINVH2", "VV", New Object() {invoiceNumber, invoiceNumber})
            tblDetails.PrimaryKey = New System.Data.DataColumn() {tblDetails.Columns("INV_TYPE"), tblDetails.Columns("INV_NO"), tblDetails.Columns("INV_LNO")}

            Dim tblSOTINVH2wk As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTINVH2 WHERE ROWNUM < 1", "SOTINVH2")

            ' use the Consolidated Invoice as the Header Row. Will modify $ and totals later
            Dim rowHeader As DataRow = tblHeader.Rows.Find(invoiceNumber)

            Dim INV_NO As String = rowHeader.Item("INV_NO")
            Dim INV_LNO As Int32 = 0

            Dim ITEM_CODE As String = String.Empty
            Dim firstRow As DataRow = Nothing
            For Each rowItem As DataRow In ASCDATA1.SelectDistinct(tblDetails, New String() {"ITEM_CODE"}).Rows
                ITEM_CODE = rowItem.Item("ITEM_CODE")
                firstRow = tblDetails.Select("ITEM_CODE = '" & ITEM_CODE & "'")(0)

                Dim ORDR_UNIT_PRICE As Decimal = 0
                Dim ORDR_QTY_SHIP As Int32 = 0

                For Each rowShipped As DataRow In tblDetails.Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    If Val(rowShipped.Item("ORDR_QTY_SHIP") & String.Empty) <> 0 Then
                        ORDR_QTY_SHIP += Val(rowShipped.Item("ORDR_QTY_SHIP") & String.Empty)
                        ORDR_UNIT_PRICE += Val(rowShipped.Item("ORDR_QTY_SHIP") & String.Empty) * Val(rowShipped.Item("ORDR_UNIT_PRICE") & String.Empty)
                    End If
                Next

                Dim rowSOTINVH2 As DataRow = tblSOTINVH2wk.NewRow
                INV_LNO += 1
                rowSOTINVH2.Item("INV_TYPE") = rowHeader.Item("INV_TYPE") & String.Empty
                rowSOTINVH2.Item("INV_NO") = INV_NO
                rowSOTINVH2.Item("INV_LNO") = INV_LNO
                rowSOTINVH2.Item("ITEM_CODE") = ITEM_CODE
                rowSOTINVH2.Item("ORDR_UNIT_PRICE") = Math.Round(ORDR_UNIT_PRICE / ORDR_QTY_SHIP, 2)
                rowSOTINVH2.Item("ORDR_QTY_SHIP") = ORDR_QTY_SHIP
                rowSOTINVH2.Item("CUST_CODE") = firstRow.Item("CUST_CODE")
                rowSOTINVH2.Item("CUST_STORE_NO") = firstRow.Item("CUST_STORE_NO")
                rowSOTINVH2.Item("WHSE_CODE") = firstRow.Item("WHSE_CODE")
                rowSOTINVH2.Item("SREP_CODE") = firstRow.Item("SREP_CODE")
                rowSOTINVH2.Item("ORDR_YYYYPP_UPDATED") = firstRow.Item("ORDR_YYYYPP_UPDATED")
                rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") = firstRow.Item("ORDR_UNIT_PRICE_CURR")
                rowSOTINVH2.Item("ITEM_UNIT_COST") = firstRow.Item("ITEM_UNIT_COST")
                rowSOTINVH2.Item("ITEM_RETAIL_PRICE") = firstRow.Item("ITEM_RETAIL_PRICE")
                rowSOTINVH2.Item("ITEM_RETAIL_PRICE_CURR") = firstRow.Item("ITEM_RETAIL_PRICE_CURR")
                rowSOTINVH2.Item("OPS_YYYYWW") = firstRow.Item("OPS_YYYYWW")
                tblSOTINVH2wk.Rows.Add(rowSOTINVH2)
            Next

            rowHeader.Item("INV_SALES") = Val(tblHeader.Compute("SUM(INV_SALES)", "") & String.Empty)
            rowHeader.Item("INV_COGS") = Val(tblHeader.Compute("SUM(INV_COGS)", "") & String.Empty)
            rowHeader.Item("INV_FREIGHT") = Val(tblHeader.Compute("SUM(INV_FREIGHT)", "") & String.Empty)
            rowHeader.Item("INV_MISC_CHG") = Val(tblHeader.Compute("SUM(INV_MISC_CHG)", "") & String.Empty)
            rowHeader.Item("INV_TOTAL_AMOUNT") = Val(tblHeader.Compute("SUM(INV_TOTAL_AMOUNT)", "") & String.Empty)
            rowHeader.Item("INV_CARTONS") = Val(tblHeader.Compute("SUM(INV_CARTONS)", "") & String.Empty)
            rowHeader.Item("INV_WEIGHT") = Val(tblHeader.Compute("SUM(INV_WEIGHT)", "") & String.Empty)
            rowHeader.Item("INV_STAX") = Val(tblHeader.Compute("SUM(INV_STAX)", "") & String.Empty)
            rowHeader.Item("INV_UNITS") = Val(tblDetails.Compute("SUM(ORDR_QTY_SHIP)", "") & String.Empty)
            'rowHeader.Item("MISC_CHG_CODE") = Val(tblHeader.Compute("SUM()", "MISC_CHG_CODE") & String.Empty)
            'rowHeader.Item("SHIP_FRT_AMT_ACCRUED") = Val(tblHeader.Compute("SUM(SHIP_FRT_AMT_ACCRUED)", "") & String.Empty)
            'rowHeader.Item("SHIP_FRT_AMT_ACTUAL") = Val(tblHeader.Compute("SUM(SHIP_FRT_AMT_ACTUAL)", "") & String.Empty)

            If Val(rowHeader.Item("INV_TOTAL_AMOUNT") & String.Empty) <> 0 Then
                rowHeader.Item("STAX_RATE") = Val(rowHeader.Item("INV_STAX") & String.Empty) / Val(rowHeader.Item("INV_TOTAL_AMOUNT") & String.Empty)
            Else
                rowHeader.Item("STAX_RATE") = Val(rowHeader.Item("INV_STAX") & String.Empty)
            End If

            ' Update header row
            For Each col As DataColumn In rowHeader.Table.Columns
                If rowSOTINVH1.Table.Columns.Contains(col.ColumnName) Then
                    rowSOTINVH1.Item(col.ColumnName) = rowHeader.Item(col.ColumnName)
                End If
            Next

            ' Get a list of column names in common
            Dim colList As New List(Of String)
            For Each col As DataColumn In tblDetails.Columns
                If tblSOTINVH2.Columns.Contains(col.ColumnName) Then
                    colList.Add(col.ColumnName)
                End If
            Next

            ' Create the invoice details
            For Each row As DataRow In tblSOTINVH2.Select("INV_NO = '" & invoiceNumber & "'")
                row.Delete()
            Next
            tblSOTINVH2.AcceptChanges()

            For Each row As DataRow In tblSOTINVH2wk.Select("")
                Dim rowSOTINVH2 As DataRow = tblSOTINVH2.NewRow
                For Each field As String In colList
                    rowSOTINVH2.Item(field) = row.Item(field)
                Next
                tblSOTINVH2.Rows.Add(rowSOTINVH2)
            Next

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    ''' <summary>
    '''  Creates a Credit based on a return
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateReturnsCredit(ByVal ORDR_TYPE_CODE As String) As Int16

        Dim numReturns As Int16 = 0
        Dim INV_TYPE As String = "C"

        Dim rowSOTINVH1 As DataRow = Nothing
        Dim rowSOTINVH2 As DataRow = Nothing
        Dim CURR_CODE As String = String.Empty
        Dim CURR_EXCH_RATE As Decimal = 1

        ORDR_TYPE_CODE = ORDR_TYPE_CODE.Trim
        If ORDR_TYPE_CODE.Length = 0 Then
            ORDR_TYPE_CODE = "REG"
        End If

        For Each rowSOTRTRN1 As DataRow In tblSOTRTRN1.Select("", "RTRN_NO")

            Dim RTRN_NO As String = rowSOTRTRN1.Item("RTRN_NO")
            Dim INV_COGS As Decimal = 0
            Dim RTRN_SALES As Decimal = 0
            Dim INV_NO As String = rowSOTRTRN1.Item("INV_NO") & String.Empty
            Dim RTRN_AS_PO_REC As String = rowSOTRTRN1.Item("RTRN_AS_PO_REC") & String.Empty

            ' Use the period on the Return.
            Dim ORDR_YYYYPP_UPDATED As String = rowSOTRTRN1.Item("OPS_YYYYPP") ' ConvertDateToPeriod(DateTime.Now)
            Dim CUST_CODE As String = rowSOTRTRN1.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
            Dim rowARTCUST2 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_STORE_NO = :PARM2", "VV", New Object() {CUST_CODE, rowSOTRTRN1.Item("CUST_STORE_NO") & String.Empty})

            Dim GL_PARM_CURR_CODE As String = rowGLTPARM1.Item("GL_PARM_CURR_CODE") & ""
            CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
            If CURR_CODE = "" Or CURR_CODE = GL_PARM_CURR_CODE Then
                CURR_CODE = GL_PARM_CURR_CODE
            Else
                Throw New Exception("Cannot Create Return for Foreign Currency Customer. Consult ABS to see how this is to be handled.")
            End If

            ' Standard routine to get the Currency Exchange Rate
            ' CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(frmASFBASE0, CURR_CODE, CDate(DateTime.Now.ToString("MM/dd/yyyy")))

            If RTRN_AS_PO_REC <> "1" Then
                For Each rowSOTRTRN2 As DataRow In tblSOTRTRN2.Select("RTRN_NO = '" & RTRN_NO & "'")
                    Dim ORDR_QTY_RTRN As Int32 = Val(rowSOTRTRN2.Item("RTRN_QTY_1") & "") + Val(rowSOTRTRN2.Item("RTRN_QTY_2") & "") + Val(rowSOTRTRN2.Item("RTRN_QTY_3") & "") + Val(rowSOTRTRN2.Item("RTRN_QTY_4") & "")
                    If ORDR_QTY_RTRN = 0 Then Continue For

                    Dim ITEM_CODE As String = rowSOTRTRN2.Item("ITEM_CODE")
                    Dim rowICTITEM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_CODE = :PARM1", "V", New Object() {ITEM_CODE})

                    Dim RTN_UNIT_COST As Decimal = Val(rowSOTRTRN2.Item("ITEM_COST_STD") & "")
                    Dim RTRN_PRICE As Decimal = Val(rowSOTRTRN2.Item("RTRN_PRICE") & "")
                    RTRN_SALES += ORDR_QTY_RTRN * RTRN_PRICE

                    rowSOTINVH2 = tblSOTINVH2.NewRow
                    With rowSOTINVH2
                        .Item("INV_TYPE") = INV_TYPE
                        .Item("INV_NO") = rowSOTRTRN1.Item("INV_NO")
                        .Item("INV_LNO") = rowSOTRTRN2.Item("RTRN_LNO")
                        .Item("ITEM_CODE") = ITEM_CODE
                        .Item("ORDR_UNIT_PRICE") = RTRN_PRICE
                        .Item("ORDR_QTY_SHIP") = ORDR_QTY_RTRN * -1

                        .Item("CUST_CODE") = CUST_CODE
                        .Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO")
                        If .Item("CUST_STORE_NO") & String.Empty = String.Empty Then
                            .Item("CUST_STORE_NO") = "000000"
                        End If
                        .Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
                        .Item("SREP_CODE") = rowSOTRTRN1.Item("SREP_CODE")
                        .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                        .Item("ORDR_UNIT_PRICE_CURR") = RTRN_PRICE
                        .Item("ITEM_UNIT_COST") = RTN_UNIT_COST
                        .Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                        .Item("ITEM_RETAIL_PRICE_CURR") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                        .Item("OPS_YYYYWW") = ASCMAIN1.CYW

                        INV_COGS += (ORDR_QTY_RTRN * RTN_UNIT_COST)
                    End With
                    tblSOTINVH2.Rows.Add(rowSOTINVH2)
                Next

                If tblSOTINVH2.Select("INV_NO = '" & INV_NO & "'").Length = 0 Then
                    Continue For
                End If
            End If

            rowSOTINVH1 = tblSOTINVH1.NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = INV_TYPE
                .Item("INV_NO") = INV_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = rowSOTRTRN1.Item("CUST_STORE_NO")
                If .Item("CUST_STORE_NO") & String.Empty = String.Empty Then
                    .Item("CUST_STORE_NO") = "000000"
                End If
                .Item("ORDR_CUST_PO") = rowSOTRTRN1.Item("CUST_CLAIM_NO")

                If ASCMAIN1.CLIENT = "AHA" And INV_TYPE = "C" Then
                    .Item("ORDR_NO") = rowSOTRTRN1.Item("RA_NO")
                End If

                If RTRN_AS_PO_REC <> "1" Then
                    .Item("WHSE_CODE") = rowSOTRTRN1.Item("WHSE_CODE")
                    .Item("INV_SALES") = Val(rowSOTRTRN1.Item("RTRN_SALES") & String.Empty) * -1
                    .Item("INV_COGS") = INV_COGS * -1
                    .Item("ORDR_TYPE_CODE") = ORDR_TYPE_CODE
                    .Item("OPS_YYYYWW") = ASCMAIN1.CYW

                    If rowARTCUST2 IsNot Nothing Then
                        .Item("CUST_SHIP_TO_STATE") = rowARTCUST2.Item("CUST_STORE_STATE") & String.Empty
                    End If

                    ASCMAIN1.sql = "Select * from ARTREASR where RA_REASON_CODE = :PARM1"
                    Dim rowARTREASR As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {rowSOTRTRN1.Item("REASON_CODE")})
                    .Item("REASON_CODE") = rowARTREASR.Item("REASON_CODE")

                    .Item("INV_FREIGHT") = Val(rowSOTRTRN1.Item("RTRN_FREIGHT") & String.Empty) * -1
                    .Item("INV_MISC_CHG") = Val(rowSOTRTRN1.Item("RTRN_HANDLING") & String.Empty) * -1
                Else
                    .Item("INV_MISC_CHG") = Val(rowSOTRTRN1.Item("RTRN_SALES") & String.Empty) * -1
                    .Item("ORDR_TYPE_CODE") = "TOP"

                    ' grab random item to get Collection and Sales Division
                    If tblSOTRTRN2.Select("RTRN_NO = '" & RTRN_NO & "'").Length > 0 Then
                        Dim ITEM_CODE As String = tblSOTRTRN2.Select("RTRN_NO = '" & RTRN_NO & "'")(0).Item("ITEM_CODE") & String.Empty
                        Dim rowICTITEM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_CODE = :PARM1", "V", ITEM_CODE)
                        If rowICTITEM1 IsNot Nothing Then
                            .Item("SALES_DIVISION_CODE") = rowICTITEM1.Item("SALES_DIVISION_CODE")
                            .Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                        End If
                    End If

                    .Item("REASON_CODE") = "JKS"
                    .Item("MISC_CHG_CODE") = "KS"
                End If

                .Item("POST_CODE") = rowSOTRTRN1.Item("POST_CODE")
                .Item("TERM_CODE") = rowARTCUST1.Item("TERM_CODE")
                .Item("SREP_CODE") = rowSOTRTRN1.Item("SREP_CODE")
                .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE")

                If rowARTCUST1.Item("CUST_BILL_TO_CUST") & String.Empty = String.Empty Then
                    .Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_CODE")
                Else
                    .Item("CUST_BILL_TO_CUST") = rowARTCUST1.Item("CUST_BILL_TO_CUST")
                End If

                rowSOTRTRN1.Item("CUST_BILL_TO_CUST") = .Item("CUST_BILL_TO_CUST")

                .Item("BRAND_CODE") = String.Empty
                .Item("EVENT_CODE") = String.Empty

                .Item("REGISTER_IND") = "0"
                .Item("SHIP_FRT_AMT_ACTUAL") = 0
                .Item("STAX_RATE") = 0
                .Item("STAX_CODE") = rowSOTRTRN1.Item("STAX_CODE")
                .Item("SHIP_FRT_AMT_ACCRUED") = 0


                .Item("INV_STAX") = Val(rowSOTRTRN1.Item("RTRN_STAX") & String.Empty) * -1
                .Item("INV_TOTAL_AMOUNT") = Val(rowSOTRTRN1.Item("RTRN_AMOUNT") & String.Empty) * -1
                .Item("INV_DATE") = rowSOTRTRN1.Item("RTRN_DATE") ' CDate(DateTime.Now.ToString("MM/dd/yyyy"))
                .Item("ORDR_YYYYPP_UPDATED") = ORDR_YYYYPP_UPDATED 'ASCMAIN1.CYP
                .Item("INIT_DATE") = DateTime.Now
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("REGISTER_XNO") = String.Empty
                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                '.Item("INV_DATE_SHIPPED") = .Item("INV_DATE")

                If .Item("INV_TOTAL_AMOUNT") = 0 Then
                    .Item("INV_PRINTED") = "1"
                End If

                .Item("INV_CARTONS") = 0
                .Item("INV_WEIGHT") = 0
                '.Item("INV_BOL_NO") = String.Empty
                '.Item("INV_PRO_NO") = String.Empty
                '.Item("SHIP_VIA_DESC") = String.Empty
                '.Item("INV_NO_CONS") = String.Empty
                '.Item("SHIP_BOL_NO") = String.Empty


                .Item("INV_COMMENT") = rowSOTRTRN1.Item("RTRN_NOTE")
                '.Item("INV_FREIGHT_TAX") = 0
                '.Item("SHIP_VIA_CODE") = String.Empty
                '.Item("ORDR_DEPT") = String.Empty
                '.Item("CUST_FACTOR_IND") = String.Empty
                '.Item("ORDR_NO_WEB") = String.Empty
                '.Item("SALES_DIVISION_CODE") = rowSOTRTRN1.Item("SALES_DIVISION_CODE")
            End With
            tblSOTINVH1.Rows.Add(rowSOTINVH1)

            numReturns += 1
            CreateOpenAR(INV_TYPE, INV_NO, CURR_EXCH_RATE)
        Next

        Return numReturns

    End Function

    ' Converts a date to a period
    Public Function ConvertDateToPeriod(inDate As Date) As String

        Dim yyyymm As String = String.Empty
        Dim stringdate As String = String.Empty

        Try
            stringdate = inDate.ToString("yyyyMMdd")
            If dictYYYYMM.ContainsKey(stringdate) Then
                Return dictYYYYMM(stringdate)
            End If

            yyyymm = ASCDATA1.GetDataValue("Select MIN(OPS_YYYYPP) from gltparm2 where prd_end_date >= '" & inDate.ToString("dd-MMM-yyyy") & "'") & String.Empty
        Catch ex As Exception
            yyyymm = ASCMAIN1.CYP
        End Try

        If Not dictYYYYMM.ContainsKey(stringdate) Then
            dictYYYYMM.Add(stringdate, yyyymm)
        End If

        Return yyyymm

    End Function

    Public Function ConvertDateToWeek(inDate As Date) As String

        Dim yyyyww As String = String.Empty
        Dim stringdate As String = String.Empty

        Try
            stringdate = inDate.ToString("yyyyMMdd")
            If dictYYYYWW.ContainsKey(stringdate) Then
                Return dictYYYYWW(stringdate)
            End If
            yyyyww = ASCDATA1.GetDataValue("Select MIN (YYYYWW) FROM GLTPARM3 WHERE WEEK_END_DATE >= '" & inDate.ToString("dd-MMM-yyyy") & "'") & String.Empty
        Catch ex As Exception
            yyyyww = ASCMAIN1.CYW
        End Try

        If Not dictYYYYWW.ContainsKey(stringdate) Then
            dictYYYYWW.Add(stringdate, yyyyww)
        End If

        Return yyyyww
    End Function

End Class
