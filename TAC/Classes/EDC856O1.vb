Public Class EDC856O1

    Private tblEDT856O1 As DataTable = Nothing
    Private tblEDT856O2 As DataTable = Nothing
    Private tblEDT856O3 As DataTable = Nothing
    Private tblEDT856O4 As DataTable = Nothing
    Private tblEDT856O5 As DataTable = Nothing
    Private tblEDTSYSIH As DataTable = Nothing

    Private rowARTPARM1 As DataRow = Nothing

    Private COMPANY_CODE As String = ASCMAIN1.DBS_COMPANY
    Private Const EDI_PROCESS_IND As String = "1"

    Private EDI_OUTBOUND_DOC_NO As String = String.Empty

    Private tblSOTSVIA1 As DataTable = Nothing
    Private tblTATTERM1 As DataTable = Nothing
    Private tblEDTTRPM1 As DataTable = Nothing
    Private tblEDTTRPM2 As DataTable = Nothing
    Private tblEDTTRPM3 As DataTable = Nothing
    Private tblWHTPKGM1 As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tblEDTSYSIHin">Reference to table EDTSYSIH</param>
    ''' <param name="tblEDT856O1in">Reference to table EDT856O1</param>
    ''' <param name="tblEDT856O2in">Reference to table EDT856O2</param>
    ''' <param name="tblEDT856O3in">Reference to table EDT856O3</param>
    ''' <param name="tblEDT856O4in">Reference to table EDT856O4</param>
    ''' <param name="tblEDT856O5in">Reference to table EDT856O5</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tblEDTSYSIHin As DataTable, _
                   ByRef tblEDT856O1in As DataTable, _
                   ByRef tblEDT856O2in As DataTable, _
                   ByRef tblEDT856O3in As DataTable, _
                   ByRef tblEDT856O4in As DataTable, _
                   ByRef tblEDT856O5in As DataTable)

        tblEDTSYSIH = tblEDTSYSIHin
        tblEDT856O1 = tblEDT856O1in
        tblEDT856O2 = tblEDT856O2in
        tblEDT856O3 = tblEDT856O3in
        tblEDT856O4 = tblEDT856O4in
        tblEDT856O5 = tblEDT856O5in

        EDI_OUTBOUND_DOC_NO = String.Empty
        tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT SOTSVIA1.*, SOTCARR1.CARRIER_TYPE FROM SOTSVIA1, SOTCARR1 WHERE SOTSVIA1.CARRIER_CODE = SOTCARR1.CARRIER_CODE (+)", "SOTSVIA1", String.Empty, Nothing)
        tblSOTSVIA1.PrimaryKey = New DataColumn() {tblSOTSVIA1.Columns("SHIP_VIA_CODE")}

        tblTATTERM1 = ASCDATA1.GetDataTable("SELECT * FROM TATTERM1", "TATTERM1", String.Empty, Nothing)
        tblEDTTRPM1 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM1 where EDI_DOC_NO = '856'", "EDTTRPM1", String.Empty, Nothing)
        tblEDTTRPM2 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM2 where EDI_DOC_NO = '856'", "EDTTRPM2", String.Empty, Nothing)
        tblEDTTRPM3 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM3 where EDI_DOC_NO = '856'", "EDTTRPM3", String.Empty, Nothing)
        tblWHTPKGM1 = ASCDATA1.GetDataTable("SELECT * FROM WHTPKGM1", "WHTPKGM1")
        rowARTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM ARTPARM1 WHERE AR_PARM_KEY = 'Z'")
        tblEDTSLSP1 = ASCDATA1.GetDataTable("SELECT * FROM EDTSLSP1")
        tblEDTSLSP1.PrimaryKey = New DataColumn() {tblEDTSLSP1.Columns("CUST_CODE")}

    End Sub

    ''' <summary>
    ''' Creates the EDT865 table entries
    ''' </summary>
    ''' <param name="SHIP_BOL_NO"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function CreateEDI856(ByVal SHIP_BOL_NO As String, ByRef ErrorMessage As String) As String

        ' What are differences between BOL and Master BOL??

        Dim rowARTCUST1 As DataRow = Nothing
        Dim rowARTCUST2 As DataRow = Nothing
        Dim rowEDT856O1 As DataRow = Nothing
        Dim rowEDT856O2 As DataRow = Nothing
        Dim rowEDT856O3 As DataRow = Nothing
        Dim rowEDT856O4 As DataRow = Nothing
        Dim rowEDT856O5 As DataRow = Nothing

        Dim rowEDT850T1 As DataRow = Nothing
        Dim rowEDT850T2 As DataRow = Nothing
        Dim rowEDTTRPM1 As DataRow = Nothing
        Dim rowICTWHSE1 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing
        Dim rowSOTPICK1 As DataRow = Nothing
        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim rowSOTSHIPB As DataRow = Nothing
        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowWHTPKGM1 As DataRow = Nothing
        Dim rowEDTSLSP1 As DataRow = Nothing

        Dim tblSOTCART1 As DataTable = Nothing
        Dim tblSOTCART2 As DataTable = Nothing
        Dim tblSOTORDR1 As DataTable = Nothing
        Dim tblSOTORDR5 As DataTable = Nothing
        Dim tblSOTPICK1 As DataTable = Nothing
        Dim tblSOTSHIP1 As DataTable = Nothing
        Dim tblARTCUST2 As DataTable = Nothing

        Dim BILL_OF_LADING_NO As String = String.Empty
        Dim CART_NO As String = String.Empty
        Dim CUST_CODE As String = String.Empty
        Dim EDI_CUSTOMER As String = String.Empty
        Dim EDI_DTL_SEQ As Int32 = 0
        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim PICK_NO As String = String.Empty
        Dim SHIP_856_BATCH_NO As String = String.Empty
        Dim rowSOTSHIP1X As DataRow = Nothing
        Dim EDI_PROMOTION As Boolean = False
        Dim EDI_DOC_SEQ_NO_PROMO As String = String.Empty

        Dim sql As String = String.Empty

        rowSOTSHIP1 = ASCDATA1.GetDataRow("Select * from SOTSHIP1 where SHIP_BOL_NO = :PARM1 and (SHIP_856_BATCH_NO IS NULL OR SHIP_856_BATCH_NO = '')", "V", New Object() {SHIP_BOL_NO})
        If rowSOTSHIP1 Is Nothing Then
            ErrorMessage = "Cannot locate shipment " & SHIP_BOL_NO
            Return String.Empty
        End If

        rowICTWHSE1 = ASCDATA1.GetDataRow("SELECT * FROM ICTWHSE1 WHERE WHSE_CODE = :PARM1", "V", New Object() {rowSOTSHIP1.Item("WHSE_CODE") & String.Empty})
        If rowICTWHSE1 Is Nothing Then
            ErrorMessage = "Cannot locate Warehouse Master for shipment " & SHIP_BOL_NO
            Return String.Empty
        End If

        BILL_OF_LADING_NO = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty

        ' Collection of Pick_Nos and Order Nos for this shipment
        If BILL_OF_LADING_NO.Length > 0 Then
            sql = " SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO "
            sql &= " FROM SOTSHIPB, SOTSHIP1, SOTPICK1"
            sql &= " WHERE SOTSHIPB.BOL_NO = SOTSHIP1.BILL_OF_LADING_NO"
            sql &= " AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
            sql &= " AND SOTSHIPB.BOL_NO = '" & BILL_OF_LADING_NO & "'"
            sql &= " AND SOTSHIP1.SHIP_856_IND = '1'"
            sql &= " AND SOTSHIP1.EDI_856_CREATED IS NULL"
        Else
            sql = " SELECT SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO "
            sql &= " FROM SOTSHIP1, SOTPICK1"
            sql &= " WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
            sql &= " AND SOTSHIP1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            sql &= " AND SOTSHIP1.SHIP_856_IND = '1'"
            sql &= " AND SOTSHIP1.EDI_856_CREATED IS NULL"
        End If

        Dim wktable As String = ASCMAIN1.Temp_Table(sql)

        If BILL_OF_LADING_NO.Length > 0 Then
            rowSOTSHIPB = ASCDATA1.GetDataRow("SELECT * FROM SOTSHIPB WHERE BOL_NO = '" & BILL_OF_LADING_NO & "'")
            rowSOTSHIP1X = ASCDATA1.GetDataRow("SELECT * FROM SOTSHIP1 WHERE BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "'")
        Else
            rowSOTSHIP1X = ASCDATA1.GetDataRow("SELECT * FROM SOTSHIP1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
        End If

        rowSOTSVIA1 = tblSOTSVIA1.Rows.Find(rowSOTSHIP1X.Item("SHIP_VIA_CODE") & String.Empty)

        tblSOTORDR1 = ASCDATA1.GetDataTable("SELECT DISTINCT(CUST_CODE) FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & wktable & ")")
        If tblSOTORDR1.Rows.Count = 0 Then
            ' Nothing Found
            ErrorMessage = "Cannot locate Orders for shipment " & SHIP_BOL_NO
            Return String.Empty
        ElseIf tblSOTORDR1.Rows.Count > 1 Then
            ' multilpe customers 
            ErrorMessage = "There are multiple customers associated with shipment " & SHIP_BOL_NO
            Return String.Empty
        Else
            CUST_CODE = tblSOTORDR1.Rows(0).Item("CUST_CODE") & String.Empty
        End If

        ' Code added 11/19/2017
        ' See if the 850s Customer Code was changed at import time
        Dim ORDR_GROUP_NO As String = rowSOTSHIP1X.Item("ORDR_GROUP_NO") & String.Empty
        sql = "Select * From SOTORDR0 Where ORDR_GROUP_NO = :PARM1"
        Dim rowSOTORDR0 As DataRow = ASCDATA1.GetDataRow(sql, "V", ORDR_GROUP_NO)
        If rowSOTORDR0 Is Nothing Then
            Return String.Empty
        End If

        ' Code added 11/29/2017
        ' See if the 850s Customer Code was changed at import time.
        rowEDTTRPM1 = Nothing
        Dim row850T1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = '" & rowSOTORDR0.Item("EDI_DOC_SEQ_NO") & "'")
        If row850T1 IsNot Nothing AndAlso row850T1.Item("CUST_CODE_OVERRIDE") & String.Empty <> String.Empty Then
            Dim TP_QUAL As String = row850T1.Item("EDI_TP_QUAL") & String.Empty
            Dim TP_ID As String = row850T1.Item("EDI_TP_ID") & String.Empty

            If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
                ' error message ?? This should never fire.
            Else
                rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
            End If
        End If

        If tblEDTTRPM1.Select("CUST_CODE = '" & CUST_CODE & "'").Length = 0 Then
            If tblEDTTRPM2.Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                Dim rowEDTTRPM2 As DataRow = tblEDTTRPM2.Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                Dim TP_QUAL As String = rowEDTTRPM2.Item("EDI_TP_QUAL") & String.Empty
                Dim TP_ID As String = rowEDTTRPM2.Item("EDI_TP_ID") & String.Empty

                If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
                    ErrorMessage = "Customer (" & CUST_CODE & ") associated with shipment " & SHIP_BOL_NO & " is not setup to receive 856 data."
                    Return String.Empty
                End If
                rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
            ElseIf tblEDTTRPM3.Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
                Dim rowEDTTRPM3 As DataRow = tblEDTTRPM3.Select("CUST_CODE = '" & CUST_CODE & "'")(0)
                Dim TP_QUAL As String = rowEDTTRPM3.Item("EDI_TP_QUAL") & String.Empty
                Dim TP_ID As String = rowEDTTRPM3.Item("EDI_TP_ID") & String.Empty

                If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
                    ErrorMessage = "Customer (" & CUST_CODE & ") associated with shipment " & SHIP_BOL_NO & " is not setup to receive 856 data."
                    Return String.Empty
                End If
                rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
            Else
                ErrorMessage = "Customer (" & CUST_CODE & ") associated with shipment " & SHIP_BOL_NO & " is not setup to receive 856 data."
                Return String.Empty
            End If
        Else
            rowEDTTRPM1 = tblEDTTRPM1.Select("CUST_CODE = '" & CUST_CODE & "'")(0)
        End If

        Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
        Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty
        EDI_CUSTOMER = rowEDTTRPM1.Item("CUST_CODE") & String.Empty

        rowEDTSLSP1 = tblEDTSLSP1.Rows.Find(CUST_CODE)
        Dim SHIP_ADDR_TYPE As String = rowSOTSHIP1X.Item("SHIP_ADDR_TYPE") & String.Empty
        Dim SHIP_ADDR_CODE As String = rowSOTSHIP1X.Item("SHIP_ADDR_CODE") & String.Empty
        Dim numChars As Int16 = 0

        If rowEDTSLSP1 IsNot Nothing Then
            Select Case SHIP_ADDR_TYPE
                Case "MK", "MA"
                    numChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
                Case "DC"
                    numChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & String.Empty)
            End Select
        End If

        If numChars > 0 And IsNumeric(SHIP_ADDR_CODE) Then
            SHIP_ADDR_CODE = SHIP_ADDR_CODE.PadLeft(numChars, "0")
            SHIP_ADDR_CODE = StrReverse(StrReverse(SHIP_ADDR_CODE).Substring(0, numChars))
        End If

        ' Reset this value for the MK in EDT865O5
        numChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)

        rowARTCUST1 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
        'If rowARTCUST1.Item("CUST_CONS_INV") & String.Empty = "1" Then
        '    Stop
        'End If

        tblARTCUST2 = ASCDATA1.GetDataTable("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1", "ARTCUST2", "V", New Object() {CUST_CODE})
        tblARTCUST2.PrimaryKey = New DataColumn() {tblARTCUST2.Columns("CUST_CODE"), tblARTCUST2.Columns("CUST_STORE_NO")}


        ' Load Cartons.
        sql = " SELECT * FROM SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTCART1 = ASCDATA1.GetDataTable(sql, "SOTCART1")
        If tblSOTCART1.Rows.Count = 0 Then
            ErrorMessage = "Cannot find cartons associated with shipment " & SHIP_BOL_NO
            Return String.Empty
        End If

        ' Load carton details
        sql = "SELECT SOTCART2.*, ICTITEM1.ITEM_STATUS, ICTITEM1.ITEM_DESC, SOTORDR2.ITEM_RETAIL_PRICE"
        sql &= ", SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_UNIT_PRICE, SOTORDR2.EDI_DOC_SEQ_NO, SOTORDR2.EDI_DTL_SEQ"
        sql &= " FROM SOTCART1, SOTCART2, ICTITEM1, SOTORDR2"
        sql &= " WHERE SOTCART1.CART_NO = SOTCART2.CART_NO"
        sql &= " AND SOTCART2.ITEM_CODE = ICTITEM1.ITEM_CODE"
        sql &= " AND SOTORDR2.ORDR_NO = SOTCART2.ORDR_NO AND SOTORDR2.ORDR_LNO = SOTCART2.ORDR_LNO"
        sql &= " AND SOTCART1.PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTCART2 = ASCDATA1.GetDataTable(sql, "SOTCART2")

        ' Load Pick Tickets
        sql = "SELECT SOTPICK1.*, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_DEPT, SOTORDR1.ORDR_DATE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_STATUS, SOTORDR1.ORDR_SHIP_TO"
        sql &= " FROM SOTORDR1, SOTPICK1"
        sql &= " WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
        sql &= " AND SOTPICK1.PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTPICK1 = ASCDATA1.GetDataTable(sql, "SOTPICK1")
        tblSOTPICK1.PrimaryKey = New DataColumn() {tblSOTPICK1.Columns("PICK_NO")}

        ' Load Shipping Addresses
        tblSOTORDR5 = ASCDATA1.GetDataTable("SELECT * FROM SOTORDR5 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & wktable & ")", "SOTORDR5")

        ' Load Shipment header
        sql = "SELECT DISTINCT SOTSHIP1.* FROM SOTSHIP1, SOTPICK1 WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO AND SOTPICK1.PICK_NO IN (SELECT PICK_NO FROM " & wktable & ")"
        tblSOTSHIP1 = ASCDATA1.GetDataTable(sql, "SOTSHIP1")
        tblSOTSHIP1.PrimaryKey = New DataColumn() {tblSOTSHIP1.Columns("SHIP_BOL_NO")}

        EDI_OUTBOUND_DOC_NO = Me.CreateEDTSYSIH(EDI_OUR_ID, EDI_TP_ID, "SH", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)
        SHIP_856_BATCH_NO = EDI_OUTBOUND_DOC_NO
        Dim EDI_ADR_SEQ As Int16 = 0

        ' Header data From SOTSHIPB, ICTWHSE1
        rowEDT856O1 = tblEDT856O1.NewRow
        rowEDT856O1.Item("COMPANY_CODE") = COMPANY_CODE
        rowEDT856O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
        rowEDT856O1.Item("WHSE_CITY") = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
        rowEDT856O1.Item("WHSE_STATE") = rowICTWHSE1.Item("WHSE_STATE") & String.Empty
        rowEDT856O1.Item("EDI_SHIP_CNT_CARTONS") = tblSOTCART1.Rows.Count
        ' Debbie does not want to enter the weights on the cartons so use weight from pick ticket - 1/30/2013
        rowEDT856O1.Item("EDI_SHIP_TOTAL_WGT") = Val(tblSOTPICK1.Compute("SUM(PICK_TOTAL_WGT)", "") & String.Empty) 'Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "") & String.Empty)
        rowEDT856O1.Item("EDI_REMIT_NAME") = rowARTPARM1.Item("AR_PARM_REMIT_NAME") & String.Empty
        rowEDT856O1.Item("EDI_TP_ID") = EDI_TP_ID
        rowEDT856O1.Item("EDI_SUPPLIER_NO") = rowEDTTRPM1.Item("EDI_ACCT_REF_NO") & String.Empty
        If rowEDT856O1.Item("EDI_SUPPLIER_NO") & String.Empty = String.Empty Then
            sql = " SELECT EDT850T1.EDI_SUPPLIER_NO "
            sql &= " FROM EDT850T1, SOTORDR1, SOTPICK1"
            sql &= " WHERE EDT850T1. EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO"
            sql &= " AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
            sql &= " AND SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            Dim row As DataRow = ASCDATA1.GetDataRow(sql)
            If row IsNot Nothing Then
                rowEDT856O1.Item("EDI_SUPPLIER_NO") = row.Item("EDI_SUPPLIER_NO") & String.Empty
            End If
        End If

        rowEDT856O1.Item("SHIP_856_BATCH_NO") = SHIP_856_BATCH_NO
        rowEDT856O1.Item("WHSE_ZIP_CODE") = rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty
        rowEDT856O1.Item("EDI_CUSTOMER") = EDI_CUSTOMER
        rowEDT856O1.Item("CARRIER_MODE") = rowSOTSVIA1.Item("CARRIER_TYPE") & String.Empty
        rowEDT856O1.Item("SHIP_ADDR_CODE") = SHIP_ADDR_CODE

        If BILL_OF_LADING_NO.Length > 0 Then
            rowEDT856O1.Item("BILL_OF_LADING_NO") = rowSOTSHIPB.Item("BOL_NO") & String.Empty ' IIf(rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty <> String.Empty, rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty, rowSOTSHIPB.Item("BOL_NO") & String.Empty)
            rowEDT856O1.Item("SHIP_BOL_NO") = rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty
            rowEDT856O1.Item("EDI_DATE_SHIPPED") = rowSOTSHIPB.Item("SHIPPED_ACTUAL")
            rowEDT856O1.Item("EDI_SCHED_DELIV_DATE") = rowSOTSHIPB.Item("SCHED_DELIV_DATE")
            rowEDT856O1.Item("FRT_TERMS") = IIf(rowSOTSHIPB.Item("FRT_TERMS") & String.Empty = "COL", "CC", "PP")
            rowEDT856O1.Item("EDI_SCAC_CODE") = rowSOTSHIPB.Item("SHIP_VIA_SCAC") & String.Empty
            rowEDT856O1.Item("EDI_PRO_NO") = rowSOTSHIPB.Item("SHIP_REF") & String.Empty
            rowEDT856O1.Item("SHIP_VIA_DESC") = rowSOTSHIPB.Item("SHIP_VIA_DESC") & String.Empty
            rowEDT856O1.Item("MASTER_BILL_OF_LADING_NO") = IIf(rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty <> String.Empty, rowSOTSHIPB.Item("MASTER_BOL_NO") & String.Empty, rowSOTSHIPB.Item("BOL_NO") & String.Empty)
            rowEDT856O1.Item("SHIP_ADDR_CODE") = tblSOTPICK1.Rows(0).Item("ORDR_SHIP_TO")
            'rowEDT856O1.Item("SHIP_MANIFEST_NO") = rowSOTSHIPB.Item("SHIP_MANIFEST_NO") & String.Empty
        Else
            rowEDT856O1.Item("BILL_OF_LADING_NO") = "9" & (rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty).ToString.Substring(1)
            rowEDT856O1.Item("SHIP_BOL_NO") = rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty
            rowEDT856O1.Item("EDI_DATE_SHIPPED") = rowSOTSHIP1X.Item("SHIP_DATE_SHIPPED")
            ' For now just add 3 days when no Bill of Lading
            rowEDT856O1.Item("EDI_SCHED_DELIV_DATE") = DateAdd(DateInterval.Day, 3, rowSOTSHIP1X.Item("SHIP_DATE_SHIPPED"))
            rowEDT856O1.Item("FRT_TERMS") = IIf(rowSOTSHIP1X.Item("FRT_TERMS") & String.Empty = "COL", "CC", "PP")
            If rowSOTSVIA1 IsNot Nothing Then
                rowEDT856O1.Item("EDI_SCAC_CODE") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
                rowEDT856O1.Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
            End If
            rowEDT856O1.Item("EDI_PRO_NO") = rowSOTSHIP1X.Item("SHIP_REF") & String.Empty
            rowEDT856O1.Item("MASTER_BILL_OF_LADING_NO") = rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty
            'rowEDT856O1.Item("SHIP_MANIFEST_NO") = rowSOTSHIPB.Item("SHIP_MANIFEST_NO") & String.Empty
        End If

        tblEDT856O1.Rows.Add(rowEDT856O1)

        Dim EDI_HL2_SEQ As Int16 = 0
        Dim EDI_HL3_SEQ As Int16 = 0
        Dim EDI_HL4_SEQ As Int16 = 0

        ' Pick Tickets
        For Each rowPICKNO As DataRow In ASCDATA1.GetDataTable("SELECT DISTINCT PICK_NO FROM " & wktable, wktable).Select("", "PICK_NO")
            PICK_NO = rowPICKNO.Item("PICK_NO")
            rowSOTPICK1 = tblSOTPICK1.Rows.Find(PICK_NO)
            SHIP_BOL_NO = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
            rowSOTSHIP1 = tblSOTSHIP1.Rows.Find(SHIP_BOL_NO)

            ' Pick Ticket / Invoice header data
            rowEDT856O2 = tblEDT856O2.NewRow
            rowEDT856O2.Item("COMPANY_CODE") = COMPANY_CODE
            rowEDT856O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            EDI_HL2_SEQ += 1
            EDI_ADR_SEQ = 0
            rowEDT856O2.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
            rowEDT856O2.Item("ORDR_CUST_PO") = rowSOTPICK1.Item("ORDR_CUST_PO") & String.Empty
            rowEDT856O2.Item("ORDR_DEPT") = rowSOTPICK1.Item("ORDR_DEPT") & String.Empty
            rowEDT856O2.Item("EDI_ORD_CNT_CARTONS") = tblSOTCART1.Select("PICK_NO = '" & PICK_NO & "'").Length
            ' Debbie does not want to enter the weights on the cartons so use weight from pick ticket - 1/30/2013
            rowEDT856O2.Item("EDI_ORD_TOTAL_WGT") = Val(rowSOTPICK1.Item("PICK_TOTAL_WGT") & String.Empty) '  Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
            rowEDT856O2.Item("ORDR_NO") = rowSOTPICK1.Item("ORDR_NO") & String.Empty
            rowEDT856O2.Item("PICK_NO") = rowSOTPICK1.Item("PICK_NO") & String.Empty
            rowEDT856O2.Item("PRO_NO") = rowSOTSHIP1.Item("SHIP_REF") & String.Empty
            rowEDT856O2.Item("INV_NO") = rowSOTPICK1.Item("INV_NO") & String.Empty
            rowEDT856O2.Item("ORDR_DATE") = rowSOTPICK1.Item("ORDR_DATE") & String.Empty
            rowEDT856O2.Item("EDI_ORDER_STATUS") = rowSOTPICK1.Item("ORDR_STATUS") & String.Empty
            rowEDT856O2.Item("CUST_STORE_NO") = rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty
            rowEDT856O2.Item("EDI_CUSTOMER") = EDI_CUSTOMER
            tblEDT856O2.Rows.Add(rowEDT856O2)

            ' These are set below when we get a value for EDI_DOC_SEQ_NO
            'rowEDT856O2.Item("EDI_PROMOTION") = String.Empty
            'rowEDT856O2.Item("EDI_MERCH_TYPE") = String.Empty
            EDI_PROMOTION = False

            ' Cartons for the above Pick Ticket
            EDI_HL3_SEQ = 0
            For Each rowSOTCART1 As DataRow In tblSOTCART1.Select("PICK_NO = '" & PICK_NO & "'", "CART_NO")
                CART_NO = rowSOTCART1.Item("CART_NO")

                rowEDT856O3 = tblEDT856O3.NewRow
                rowEDT856O3.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT856O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT856O3.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                EDI_HL3_SEQ += 1
                rowEDT856O3.Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                rowEDT856O3.Item("CART_NO") = CART_NO
                rowEDT856O3.Item("CART_TOTAL_WGT_ACTUAL") = Val(rowSOTCART1.Item("CART_TOTAL_WGT_ACTUAL") & String.Empty)
                rowEDT856O3.Item("CART_SEQ") = Val(rowSOTCART1.Item("CART_SEQ") & String.Empty)
                'SOTCART1.CART_TRACKING_NO -> EDT856O3.EDI_CTN_TRACKING_NUMBER
                ' 5/9/2016
                rowEDT856O3.Item("EDI_CTN_TRACKING_NUMBER") = rowSOTCART1.Item("CART_TRACKING_NO") & String.Empty


                If rowSOTCART1.Item("PACKAGING_TYPE") & String.Empty = 31 Then
                    rowWHTPKGM1 = tblWHTPKGM1.Rows.Find(rowSOTCART1.Item("PKG_CODE"))
                    If rowWHTPKGM1 IsNot Nothing Then
                        rowEDT856O3.Item("CARTON_LENGTH") = rowWHTPKGM1.Item("PKG_L")
                        rowEDT856O3.Item("CARTON_WIDTH") = rowWHTPKGM1.Item("PKG_W")
                        rowEDT856O3.Item("CARTON_HEIGHT") = rowWHTPKGM1.Item("PKG_H")
                    End If
                Else
                    'rowEDT856O3.Item("CARTON_LENGTH") = String.Empty
                    'rowEDT856O3.Item("CARTON_WIDTH") = String.Empty
                    'rowEDT856O3.Item("CARTON_HEIGHT") = String.Empty
                End If

                'rowEDT856O3.Item("CARTON_WGT_PER") = String.Empty
                tblEDT856O3.Rows.Add(rowEDT856O3)

                ' Carton Contents
                EDI_HL4_SEQ = 0
                For Each rowSOTCART2 As DataRow In tblSOTCART2.Select("CART_NO = '" & CART_NO & "'", "CART_LNO")
                    EDI_DOC_SEQ_NO = rowSOTCART2.Item("EDI_DOC_SEQ_NO") & String.Empty
                    EDI_DTL_SEQ = Val(rowSOTCART2.Item("EDI_DTL_SEQ") & String.Empty)

                    If EDI_PROMOTION = False Then
                        EDI_PROMOTION = True
                        ' prevents duplicate lookups
                        If EDI_DOC_SEQ_NO_PROMO <> EDI_DOC_SEQ_NO Then
                            rowEDT850T1 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = :PARM1", "V", New Object() {EDI_DOC_SEQ_NO})
                            EDI_DOC_SEQ_NO_PROMO = EDI_DOC_SEQ_NO
                        End If
                        If rowEDT850T1 IsNot Nothing Then
                            rowEDT856O2.Item("EDI_PROMOTION") = rowEDT850T1.Item("EDI_PROMOTION") & String.Empty
                            rowEDT856O2.Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE") & String.Empty
                            rowEDT856O2.Item("EDI_PO_TYPE") = rowEDT850T1.Item("EDI_PO_TYPE") & String.Empty
                        End If
                    End If

                    ' Can we add items to an EDI Order? If so, do we send the new items over
                    rowEDT850T2 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T2 WHERE EDI_DOC_SEQ_NO = :PARM1 AND EDI_DTL_SEQ = :PARM1", "VN", New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ})

                    rowEDT856O4 = tblEDT856O4.NewRow
                    rowEDT856O4.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT856O4.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT856O4.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
                    rowEDT856O4.Item("EDI_HL3_SEQ") = EDI_HL3_SEQ
                    EDI_HL4_SEQ += 1
                    rowEDT856O4.Item("EDI_HL4_SEQ") = EDI_HL4_SEQ
                    rowEDT856O4.Item("CART_LNO") = rowSOTCART2.Item("CART_LNO")
                    rowEDT856O4.Item("STYLE_UPC_CODE") = (rowEDT850T2.Item("EDI_UPC") & String.Empty).trim
                    rowEDT856O4.Item("STYLE_CODE") = (rowEDT850T2.Item("EDI_STYLE") & String.Empty).trim
                    rowEDT856O4.Item("STYLE_STATUS") = rowSOTCART2.Item("ITEM_STATUS") & String.Empty
                    rowEDT856O4.Item("PICK_QTY_CONF") = rowSOTCART2.Item("QTY_PACKED") & String.Empty
                    rowEDT856O4.Item("ORDR_QTY_ORIG") = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & String.Empty)
                    rowEDT856O4.Item("STYLE_DESC") = (rowSOTCART2.Item("ITEM_DESC") & String.Empty).ToString.PadRight(35, " ").Substring(0, 35).Trim
                    rowEDT856O4.Item("STYLE_RETAIL") = Val(rowSOTCART2.Item("ITEM_RETAIL_PRICE") & String.Empty)
                    rowEDT856O4.Item("ORIG_PRICE") = Val(rowEDT850T2.Item("EDI_PRICE") & String.Empty)
                    ' Added 6/24/2014
                    rowEDT856O4.Item("EDI_PO_LNO") = rowEDT850T2.Item("EDI_PO_LNO")
                    rowEDT856O4.Item("EDI_SKU") = (rowEDT850T2.Item("EDI_SKU") & String.Empty).trim
                    rowEDT856O4.Item("EDI_PO4_UOM") = (rowEDT850T2.Item("EDI_PO4_UOM") & String.Empty).trim
                    rowEDT856O4.Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
                    rowEDT856O4.Item("STYLE_GTIN_CODE") = (rowEDT850T2.Item("EDI_GTIN") & String.Empty).trim
                    rowEDT856O4.Item("COLOR_CODE") = (rowEDT850T2.Item("EDI_COLOR_CODE") & String.Empty).trim
                    rowEDT856O4.Item("EDI_SIZE_DESC") = (rowEDT850T2.Item("EDI_SIZE_DESC") & String.Empty).trim
                    rowEDT856O4.Item("EDI_ITEM") = rowSOTCART2.Item("ITEM_CODE") & String.Empty
                    rowEDT856O4.Item("EDI_EAN") = (rowEDT850T2.Item("EDI_EAN") & String.Empty).ToString.Trim
                    'In the 856 processing, please take EDT850T2. EDI_PO4_INNER and populate EDT856O4.EDI_PO4_INNER
                    rowEDT856O4.Item("EDI_PO4_INNER") = rowEDT850T2.Item("EDI_PO4_INNER")
                    rowEDT856O4.Item("EDI_SIZE_CODE") = rowEDT850T2.Item("EDI_SIZE_CODE")
                    rowEDT856O4.Item("EDI_COLOR_NAME") = rowEDT850T2.Item("EDI_COLOR_NAME")
                    tblEDT856O4.Rows.Add(rowEDT856O4)

                Next ' End SOTCART2
            Next ' End SOTCART1

            Dim CUST_STORE_NO As String = rowSOTPICK1.Item("CUST_STORE_NO") & String.Empty

            ' Only one entry for Ship To and Ship From
            If EDI_HL2_SEQ = 1 Then
                ' Ship From
                rowEDT856O5 = tblEDT856O5.NewRow
                rowEDT856O5.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT856O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT856O5.Item("EDI_HL2_SEQ") = 0 'EDI_HL2_SEQ
                EDI_ADR_SEQ += 1
                rowEDT856O5.Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                rowEDT856O5.Item("EDI_ADDR_TYPE") = "SF"
                rowEDT856O5.Item("EDI_CUST_NAME_ADR") = rowICTWHSE1.Item("WHSE_DESC") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS1") = rowICTWHSE1.Item("WHSE_ADDR1") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS2") = rowICTWHSE1.Item("WHSE_ADDR2") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS3") = rowICTWHSE1.Item("WHSE_ADDR3") & String.Empty
                rowEDT856O5.Item("EDI_CITY") = rowICTWHSE1.Item("WHSE_CITY") & String.Empty
                rowEDT856O5.Item("EDI_STATE") = rowICTWHSE1.Item("WHSE_STATE") & String.Empty
                rowEDT856O5.Item("EDI_ZIPCODE") = (rowICTWHSE1.Item("WHSE_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                rowEDT856O5.Item("EDI_COUNTRY") = rowICTWHSE1.Item("WHSE_COUNTRY") & String.Empty
                rowEDT856O5.Item("EDI_ADDR_CODE") = rowICTWHSE1.Item("WHSE_EDI_ID") & String.Empty
                'rowEDT856O5.Item("EDI_ADDR_CODE_QUAL") = rowICTWHSE1.Item("WHSE_EDI_QUAL") & String.Empty
                tblEDT856O5.Rows.Add(rowEDT856O5)

                ' Ship to
                rowSOTORDR5 = tblSOTORDR5.Rows.Find(New Object() {rowSOTPICK1.Item("ORDR_NO"), "ST"})
                If rowSOTORDR5 Is Nothing Then
                    rowSOTORDR5 = ASCDATA1.GetDataRow("select * from ARTCUST1 Where CUST_CODE = :PARM1", "V", CUST_CODE)
                    rowARTCUST2 = Nothing
                Else
                    rowARTCUST2 = tblARTCUST2.Rows.Find(New Object() {CUST_CODE, rowSOTORDR5.Item("CUST_ADDR_CODE") & String.Empty})
                End If

                rowEDT856O5 = tblEDT856O5.NewRow
                rowEDT856O5.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT856O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT856O5.Item("EDI_HL2_SEQ") = 0 'EDI_HL2_SEQ
                EDI_ADR_SEQ += 1
                rowEDT856O5.Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
                rowEDT856O5.Item("EDI_ADDR_TYPE") = "ST"
                rowEDT856O5.Item("EDI_CUST_NAME_ADR") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                rowEDT856O5.Item("EDI_ADDRESS3") = rowSOTORDR5.Item("CUST_ADDR3") & String.Empty
                rowEDT856O5.Item("EDI_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                rowEDT856O5.Item("EDI_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                rowEDT856O5.Item("EDI_ZIPCODE") = (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
                rowEDT856O5.Item("EDI_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty

                If rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("CUST_DC_ALIAS") & String.Empty <> String.Empty Then
                    rowEDT856O5.Item("EDI_ADDR_CODE") = rowARTCUST2.Item("CUST_DC_ALIAS") & String.Empty
                ElseIf rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty <> String.Empty Then
                    rowEDT856O5.Item("EDI_ADDR_CODE") = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty
                Else
                    rowEDT856O5.Item("EDI_ADDR_CODE") = SHIP_ADDR_CODE
                End If

                'rowEDT856O5.Item("EDI_ADDR_CODE_QUAL") = rowICTWHSE1.Item("WHSE_EDI_QUAL") & String.Empty
                tblEDT856O5.Rows.Add(rowEDT856O5)
                EDI_ADR_SEQ = 0
            End If

            ' Mark For
            rowARTCUST2 = tblARTCUST2.Rows.Find(New Object() {CUST_CODE, CUST_STORE_NO})

            rowEDT856O5 = tblEDT856O5.NewRow
            rowEDT856O5.Item("COMPANY_CODE") = COMPANY_CODE
            rowEDT856O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
            rowEDT856O5.Item("EDI_HL2_SEQ") = EDI_HL2_SEQ
            EDI_ADR_SEQ += 1

            rowEDT856O5.Item("EDI_ADR_SEQ") = EDI_ADR_SEQ
            rowEDT856O5.Item("EDI_ADDR_TYPE") = "MK"
            rowEDT856O5.Item("EDI_CUST_NAME_ADR") = rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty
            rowEDT856O5.Item("EDI_ADDRESS1") = rowARTCUST2.Item("CUST_STORE_ADDR1") & String.Empty
            rowEDT856O5.Item("EDI_ADDRESS2") = rowARTCUST2.Item("CUST_STORE_ADDR2") & String.Empty
            rowEDT856O5.Item("EDI_ADDRESS3") = rowARTCUST2.Item("CUST_STORE_ADDR3") & String.Empty
            rowEDT856O5.Item("EDI_CITY") = rowARTCUST2.Item("CUST_STORE_CITY") & String.Empty
            rowEDT856O5.Item("EDI_STATE") = rowARTCUST2.Item("CUST_STORE_STATE") & String.Empty
            rowEDT856O5.Item("EDI_ZIPCODE") = (rowARTCUST2.Item("CUST_STORE_ZIP_CODE") & String.Empty).ToString.Replace("-", "").Replace(" ", "")
            rowEDT856O5.Item("EDI_COUNTRY") = rowARTCUST2.Item("CUST_STORE_COUNTRY") & String.Empty

            If numChars > 0 AndAlso IsNumeric(CUST_STORE_NO) Then
                CUST_STORE_NO = CUST_STORE_NO.PadLeft(numChars, "0")
                CUST_STORE_NO = StrReverse(StrReverse(CUST_STORE_NO).Substring(0, numChars))
            End If

            If rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty <> String.Empty Then
                rowEDT856O5.Item("EDI_ADDR_CODE") = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty
            Else
                rowEDT856O5.Item("EDI_ADDR_CODE") = CUST_STORE_NO
            End If

            'rowEDT856O5.Item("EDI_ADDR_CODE_QUAL") = rowICTWHSE1.Item("WHSE_EDI_QUAL") & String.Empty
            tblEDT856O5.Rows.Add(rowEDT856O5)

        Next ' End PICK_NO 

        Return EDI_OUTBOUND_DOC_NO

    End Function

    Private Function CreateEDTSYSIH(ByVal EDI_OUR_ID As String, ByVal EDI_TP_ID As String, ByVal ediApplicationID As String, ByVal EDI_STATUS As String) As String

        CreateEDTSYSIH = String.Empty

        Dim ediOutboundDocNo As String = String.Empty
        ' Moved from up above
        ediOutboundDocNo = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")

        Dim rowEDTSYSIH As DataRow = tblEDTSYSIH.NewRow
        rowEDTSYSIH.Item("COMPANY_CODE") = COMPANY_CODE
        rowEDTSYSIH.Item("EDI_OUTBOUND_DOC_NO") = ediOutboundDocNo
        rowEDTSYSIH.Item("EDI_APPLICATION_ID") = ediApplicationID
        If EDI_STATUS = "P" Then
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = EDI_PROCESS_IND
        Else
            rowEDTSYSIH.Item("EDI_PROCESS_IND") = "T"
        End If
        rowEDTSYSIH.Item("EDI_OUR_ID") = EDI_OUR_ID
        rowEDTSYSIH.Item("EDI_TP_ID") = EDI_TP_ID
        rowEDTSYSIH.Item("INIT_DATE") = DateTime.Now
        rowEDTSYSIH.Item("INIT_OPER") = ASCMAIN1.USER_ID
        'rowEDTSYSIH.Item("LAST_DATE") = DateTime.Now
        'rowEDTSYSIH.Item("LAST_OPER") = ASCMAIN1.USER_ID
        tblEDTSYSIH.Rows.Add(rowEDTSYSIH)

        CreateEDTSYSIH = ediOutboundDocNo

    End Function

End Class
