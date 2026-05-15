Public Class EDC810O1

    Private tblEDT810O1 As DataTable = Nothing
    Private tblEDT810O2 As DataTable = Nothing
    Private tblEDT810O3 As DataTable = Nothing
    Private tblEDT810O5 As DataTable = Nothing
    Private tblEDTSYSIH As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing
    Private rowARTPARM1 As DataRow = Nothing
    Private rowGLTPARM1 As DataRow = Nothing
    Private GL_PARM_CURR_CODE As String = String.Empty


    Private COMPANY_CODE As String = ASCMAIN1.DBS_COMPANY
    Private Const EDI_PROCESS_IND As String = "1"

    Private EDI_OUTBOUND_DOC_NO As String = String.Empty

    Private tblSOTSVIA1 As DataTable = Nothing
    Private tblTATTERM1 As DataTable = Nothing
    Private tblEDTTRPM1 As DataTable = Nothing
    Private tblEDTTRPM2 As DataTable = Nothing
    Private tblEDTTRPM3 As DataTable = Nothing

    ''' <summary>
    ''' Creates the EDI 810 entry for a Shipment
    ''' </summary>
    ''' <param name="tblEDTSYSIHin">Reference to table EDTSYSIH</param>
    ''' <param name="tblEDT810O1in">Reference to table EDT810O1</param>
    ''' <param name="tblEDT810O2in">Reference to table EDT81002</param>
    ''' <param name="tblEDT810O3in">Reference to table EDT81003</param>
    ''' <param name="tblEDT810O5in">Reference to table EDT81005</param>
    ''' <remarks></remarks>
    Public Sub New(ByRef tblEDTSYSIHin As DataTable, _
                   ByRef tblEDT810O1in As DataTable, _
                   ByRef tblEDT810O2in As DataTable, _
                   ByRef tblEDT810O3in As DataTable, _
                   ByRef tblEDT810O5in As DataTable)

        tblEDTSYSIH = tblEDTSYSIHin
        tblEDT810O1 = tblEDT810O1in
        tblEDT810O2 = tblEDT810O2in
        tblEDT810O3 = tblEDT810O3in
        tblEDT810O5 = tblEDT810O5in

        EDI_OUTBOUND_DOC_NO = String.Empty
        tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT SOTSVIA1.* FROM SOTSVIA1", "SOTSVIA1", String.Empty, Nothing)
        tblTATTERM1 = ASCDATA1.GetDataTable("SELECT * FROM TATTERM1", "TATTERM1", String.Empty, Nothing)
        tblEDTTRPM1 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM1 where EDI_DOC_NO = '810'", "EDTTRPM1", String.Empty, Nothing)
        tblEDTTRPM2 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM2 where EDI_DOC_NO = '810'", "EDTTRPM2", String.Empty, Nothing)
        tblEDTTRPM3 = ASCDATA1.GetDataTable("SELECT * FROM EDTTRPM3 where EDI_DOC_NO = '810'", "EDTTRPM3", String.Empty, Nothing)
        rowARTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM ARTPARM1 WHERE AR_PARM_KEY = 'Z'")
        rowGLTPARM1 = ASCDATA1.GetDataRow("SELECT * FROM GLTPARM1 WHERE GL_PARM_KEY = 'Z'")
        GL_PARM_CURR_CODE = rowGLTPARM1.Item("GL_PARM_CURR_CODE") & String.Empty
        tblEDTSLSP1 = ASCDATA1.GetDataTable("SELECT * FROM EDTSLSP1")
        tblEDTSLSP1.PrimaryKey = New DataColumn() {tblEDTSLSP1.Columns("CUST_CODE")}

    End Sub

    Public Sub Create810(ByVal SHIP_BOL_NO As String, ByRef EDI_OUTBOUND_DOC_NO As String)

        Dim rowEDT810O1 As DataRow = Nothing
        Dim rowEDT810O2 As DataRow = Nothing
        Dim rowEDT810O3 As DataRow = Nothing
        Dim rowEDT810O5 As DataRow = Nothing

        Dim rowSOTSHIP1 As DataRow = Nothing
        Dim rowICTITEM1 As DataRow = Nothing
        Dim rowSOTORDR0 As DataRow = Nothing
        Dim rowSOTORDR1 As DataRow = Nothing
        Dim rowSOTORDR2 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing
        Dim rowSOTSVIA1 As DataRow = Nothing
        Dim rowTATTERM1 As DataRow = Nothing
        Dim rowEDTTRPM1 As DataRow = Nothing
        Dim rowSOTINVH1 As DataRow = Nothing
        Dim rowARTOPEN1 As DataRow = Nothing
        Dim rowARTCUST2 As DataRow = Nothing

        Dim tblSOTINVH2 As DataTable = Nothing
        Dim tblSOTORDR2 As DataTable = Nothing
        Dim tblSOTORDR5 As DataTable = Nothing
        Dim tblSOTCART1 As DataTable = Nothing
        Dim tblSOTSHIP1 As DataTable = Nothing

        Dim ITEM_CODE As String = String.Empty
        Dim PICK_NO As String = String.Empty
        Dim INV_NO As String = String.Empty
        Dim ORDR_NO As String = String.Empty
        Dim CUST_CODE As String = String.Empty

        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim EDI_DTL_SEQ As Int16 = 0
        Dim EDI_BUYER_ITEM As String = String.Empty
        Dim sql As String = String.Empty
        Dim BillOfLading As String = String.Empty

        Dim rowEDTSLSP1 As DataRow = Nothing
        Dim rowEDT850T1 As DataRow = Nothing
        Dim numNonZeroDollarInvoices As Int16 = 0
        Dim foreignExchange As Boolean = False

        tblSOTSHIP1 = ASCDATA1.GetDataTable("select * from SOTSHIP1 where SHIP_BOL_NO = :PARM1 and (SHIP_810_BATCH_NO IS NULL OR SHIP_810_BATCH_NO = '')", "SOTSHIP1", "V", New Object() {SHIP_BOL_NO})
        If tblSOTSHIP1.Rows.Count = 0 Then
            Exit Sub
        End If

        'If tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty <> String.Empty Then
        '    tblSOTSHIP1 = ASCDATA1.GetDataTable("select * from SOTSHIP1 where BILL_OF_LADING_NO = :PARM1", "SOTSHIP1", "V", New Object() {tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty})
        'End If

        If tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty <> String.Empty Then
            BillOfLading = tblSOTSHIP1.Rows(0).Item("BILL_OF_LADING_NO") & String.Empty
            ' Interparfums needs all the characters for Dillards
            If COMPANY_CODE = "AHA" Then
                BillOfLading = BillOfLading.PadLeft(10, "0")
                BillOfLading = StrReverse(StrReverse(BillOfLading).Substring(0, 10))
            End If
        Else
            BillOfLading = "9" & SHIP_BOL_NO.Substring(1)
        End If

        Dim ORDR_GROUP_NO As String = tblSOTSHIP1.Rows(0).Item("ORDR_GROUP_NO") & String.Empty
        If ORDR_GROUP_NO.Length = 0 Then
            Exit Sub
        End If

        sql = "Select * From SOTORDR0 Where ORDR_GROUP_NO = :PARM1"
        rowSOTORDR0 = ASCDATA1.GetDataRow(sql, "V", ORDR_GROUP_NO)
        If rowSOTORDR0 Is Nothing Then
            Exit Sub
        End If
        CUST_CODE = rowSOTORDR0.Item("CUST_CODE") & String.Empty

        Dim rowARTCUST1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CUST_CODE})
        If rowARTCUST1 Is Nothing Then
            Exit Sub
        End If

        Dim CUST_CONS_INV As Boolean = rowARTCUST1.Item("CUST_CONS_INV") & String.Empty = "1"
        Dim CUST_CODE_810 As String = String.Empty

        ' Code added 11/29/2017
        ' See if the 850s Customer Code was changed at import time.
        rowEDTTRPM1 = Nothing
        rowEDT850T1 = ASCDATA1.GetDataRow("SELECT * FROM EDT850T1 WHERE EDI_DOC_SEQ_NO = '" & rowSOTORDR0.Item("EDI_DOC_SEQ_NO") & "'")
        If rowEDT850T1 IsNot Nothing Then
            CUST_CODE_810 = rowEDT850T1.Item("CUST_CODE") & String.Empty
        End If

        'If rowEDT850T1 IsNot Nothing AndAlso rowEDT850T1.Item("CUST_CODE_OVERRIDE") & String.Empty <> String.Empty Then
        '    Dim TP_QUAL As String = rowEDT850T1.Item("EDI_TP_QUAL") & String.Empty
        '    Dim TP_ID As String = rowEDT850T1.Item("EDI_TP_ID") & String.Empty

        '    If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
        '        ' error message ?? This should never fire.
        '    Else
        '        rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
        '    End If
        'End If

        If tblEDTTRPM1.Select("CUST_CODE = '" & CUST_CODE_810 & "'").Length > 0 Then
            rowEDTTRPM1 = tblEDTTRPM1.Select("CUST_CODE = '" & CUST_CODE_810 & "'")(0)
        ElseIf tblEDTTRPM2.Select("CUST_CODE = '" & CUST_CODE_810 & "'").Length > 0 Then
            Dim rowEDTTRMP2 As DataRow = tblEDTTRPM2.Select("CUST_CODE = '" & CUST_CODE_810 & "'")(0)
            Dim TP_QUAL As String = rowEDTTRMP2.Item("EDI_TP_QUAL") & String.Empty
            Dim TP_ID As String = rowEDTTRMP2.Item("EDI_TP_ID") & String.Empty

            If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
                Exit Sub
            End If
            rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
        ElseIf tblEDTTRPM3.Select("CUST_CODE = '" & CUST_CODE_810 & "'").Length > 0 Then
            Dim rowEDTTRMP3 As DataRow = tblEDTTRPM3.Select("CUST_CODE = '" & CUST_CODE_810 & "'")(0)
            Dim TP_QUAL As String = rowEDTTRMP3.Item("EDI_TP_QUAL") & String.Empty
            Dim TP_ID As String = rowEDTTRMP3.Item("EDI_TP_ID") & String.Empty

            If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
                Exit Sub
            End If
            rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
        ElseIf tblEDTTRPM2.Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
            Dim rowEDTTRMP2 As DataRow = tblEDTTRPM2.Select("CUST_CODE = '" & CUST_CODE & "'")(0)
            Dim TP_QUAL As String = rowEDTTRMP2.Item("EDI_TP_QUAL") & String.Empty
            Dim TP_ID As String = rowEDTTRMP2.Item("EDI_TP_ID") & String.Empty

            If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
                Exit Sub
            End If
            rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
        ElseIf tblEDTTRPM3.Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
            Dim rowEDTTRMP3 As DataRow = tblEDTTRPM3.Select("CUST_CODE = '" & CUST_CODE & "'")(0)
            Dim TP_QUAL As String = rowEDTTRMP3.Item("EDI_TP_QUAL") & String.Empty
            Dim TP_ID As String = rowEDTTRMP3.Item("EDI_TP_ID") & String.Empty

            If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length = 0 Then
                Exit Sub
            End If
            rowEDTTRPM1 = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0)
        Else
            Exit Sub
        End If

        Dim EDI_OUR_ID As String = rowEDTTRPM1.Item("EDI_OUR_ID") & String.Empty
        Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID") & String.Empty

        rowEDTSLSP1 = tblEDTSLSP1.Rows.Find(CUST_CODE)
        If rowEDTSLSP1 Is Nothing AndAlso CUST_CODE_810.Length > 0 Then
            rowEDTSLSP1 = tblEDTSLSP1.Rows.Find(CUST_CODE_810)
        End If
        Dim SHIP_ADDR_TYPE As String = tblSOTSHIP1.Rows(0).Item("SHIP_ADDR_TYPE") & String.Empty
        Dim SHIP_ADDR_CODE As String = tblSOTSHIP1.Rows(0).Item("SHIP_ADDR_CODE") & String.Empty
        Dim mkNumChars As Int16 = 0
        Dim dcNumChars As Int16 = 0

        mkNumChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
        dcNumChars = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & String.Empty)

        If rowEDTSLSP1 IsNot Nothing Then
            Select Case SHIP_ADDR_TYPE
                Case "MK", "MA"
                    If mkNumChars > 0 AndAlso IsNumeric(SHIP_ADDR_CODE) Then
                        SHIP_ADDR_CODE = SHIP_ADDR_CODE.PadLeft(mkNumChars, "0")
                        SHIP_ADDR_CODE = StrReverse(StrReverse(SHIP_ADDR_CODE).Substring(0, mkNumChars))
                    End If

                Case "DC"
                    If dcNumChars > 0 AndAlso IsNumeric(SHIP_ADDR_CODE) Then
                        SHIP_ADDR_CODE = SHIP_ADDR_CODE.PadLeft(dcNumChars, "0")
                        SHIP_ADDR_CODE = StrReverse(StrReverse(SHIP_ADDR_CODE).Substring(0, dcNumChars))
                    End If
            End Select
        End If

        '************************************************************************************
        ' This is wehre there should be a check for consolidated invoices
        ' If so, call another sub procedure.
        '************************************************************************************

        ' set a default incase all invoices are $0.00

        EDI_OUTBOUND_DOC_NO = String.Empty ' "xxx"

        For Each rowSOTSHIP1 In tblSOTSHIP1.Rows

            If CUST_CONS_INV Then
                ' Get the Pick Ticket for the Lead Consolidated Invoice
                sql = " Select SOTPICK1.*"
                sql &= " from SOTPICK1, SOTSHIP1, SOTINVH1"
                sql &= " where SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
                sql &= " and SOTPICK1.INV_NO = SOTINVH1.INV_NO"
                sql &= " and SOTPICK1.SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'"
                sql &= " and SOTINVH1.INV_NO = SOTINVH1.INV_NO_CONS"
            Else
                sql = "SELECT * FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "' AND PICK_STATUS = 'F'"
            End If

            For Each rowSOTPICK1 As DataRow In ASCDATA1.GetDataTable(sql).Select("", "PICK_NO")

                PICK_NO = rowSOTPICK1.Item("PICK_NO") & String.Empty
                INV_NO = rowSOTPICK1.Item("INV_NO") & String.Empty
                ORDR_NO = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                foreignExchange = False

                rowSOTINVH1 = ASCDATA1.GetDataRow("Select * from SOTINVH1 where inv_no = :PARM1", "V", New Object() {INV_NO})

                If rowSOTINVH1 Is Nothing Then
                    rowARTOPEN1 = Nothing
                Else
                    rowARTOPEN1 = ASCDATA1.GetDataRow("select * from ARTOPEN1 where CUST_CODE = :PARM1 AND INV_NUM = :PARM2", "VV", New Object() {rowSOTINVH1.Item("CUST_CODE"), INV_NO})
                    If rowSOTINVH1.Item("CURR_CODE") & String.Empty <> String.Empty AndAlso rowSOTINVH1.Item("CURR_CODE") & String.Empty <> GL_PARM_CURR_CODE Then
                        foreignExchange = True
                    End If
                End If

                sql = "Select SOTINVH2.*, ICTITEM1.ITEM_UOM, ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_DESC"
                sql &= " FROM SOTINVH2, ICTITEM1"
                sql &= " WHERE SOTINVH2.ITEM_CODE = ICTITEM1.ITEM_CODE"
                sql &= " AND SOTINVH2.INV_NO = :PARM1"
                tblSOTINVH2 = ASCDATA1.GetDataTable(sql, "SOTINVH2", "V", New Object() {INV_NO})

                If Not CUST_CONS_INV Then
                    rowSOTORDR1 = ASCDATA1.GetDataRow("select * from SOTORDR1 where ordr_no = :PARM1", "V", New Object() {ORDR_NO})

                    sql = " SELECT SOTORDR2.*, EDT850T2.*"
                    sql &= " FROM SOTORDR2, EDT850T2 "
                    sql &= " WHERE SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO (+)"
                    sql &= " AND SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ (+)"
                    sql &= " AND ORDR_NO = :PARM1"
                    tblSOTORDR2 = ASCDATA1.GetDataTable(sql, "SOTORDR2", "V", New Object() {ORDR_NO})

                    'tblSOTORDR2 = ASCDATA1.GetDataTable("select * from SOTORDR2 where ordr_no = :PARM1", "SOTORDR2", "V", New Object() {ORDR_NO})
                    tblSOTORDR5 = ASCDATA1.GetDataTable("select * from SOTORDR5 where ordr_no = :PARM1", "SOTORDR5", "V", New Object() {ORDR_NO})
                    tblSOTCART1 = ASCDATA1.GetDataTable("select * from SOTCART1 where PICK_NO = :PARM1", "SOTCART1", "V", New Object() {PICK_NO})
                Else
                    Dim clsSOTINVH1 As New TAC.SOCINVH1
                    clsSOTINVH1.CreateConsolidatedInvoice(INV_NO, rowSOTINVH1, tblSOTINVH2)

                    rowSOTORDR1 = ASCDATA1.GetDataRow("select * from SOTORDR1 where ordr_no = :PARM1", "V", New Object() {ORDR_NO})

                    sql = " SELECT SOTORDR2.*, EDT850T2.*"
                    sql &= " FROM SOTORDR2, EDT850T2 "
                    sql &= " WHERE SOTORDR2.EDI_DOC_SEQ_NO = EDT850T2.EDI_DOC_SEQ_NO (+)"
                    sql &= " AND SOTORDR2.EDI_DTL_SEQ = EDT850T2.EDI_DTL_SEQ (+)"
                    sql &= " AND (SOTORDR2.ITEM_CODE, SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO) IN"

                    'sql = " SELECT * FROM SOTORDR2 WHERE (ITEM_CODE, ORDR_NO, ORDR_LNO) IN"
                    sql &= " ("
                    sql &= " SELECT ITEM_CODE, ORDR_NO, MAX(ORDR_LNO) ORDR_LNO"
                    sql &= " FROM SOTORDR2"
                    sql &= " WHERE (ITEM_CODE, ORDR_NO) IN"
                    sql &= " ("
                    sql &= " SELECT SOTORDR2.ITEM_CODE, MAX(SOTORDR2.ORDR_NO) ORDR_NO"
                    sql &= " FROM SOTORDR1, SOTORDR2, SOTINVH1"
                    sql &= " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
                    sql &= " AND SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO"
                    sql &= " AND SOTINVH1.INV_NO_CONS = :PARM1"
                    sql &= " GROUP BY SOTORDR2.ITEM_CODE"
                    sql &= " )"
                    sql &= " GROUP BY ITEM_CODE, ORDR_NO"
                    sql &= " )"

                    tblSOTORDR2 = ASCDATA1.GetDataTable(sql, "SOTORDR2", "V", New Object() {INV_NO})
                    tblSOTORDR5 = ASCDATA1.GetDataTable("select * from SOTORDR5 where ordr_no = :PARM1", "SOTORDR5", "V", New Object() {ORDR_NO})

                    sql = " Select SOTCART1.* "
                    sql &= " from SOTCART1, SOTPICK1, SOTINVH1"
                    sql &= " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO"
                    sql &= " and SOTPICK1.INV_NO = SOTINVH1.INV_NO"
                    sql &= " and SOTINVH1.INV_NO_CONS = :PARM1"
                    tblSOTCART1 = ASCDATA1.GetDataTable(sql, "SOTCART1", "V", New Object() {INV_NO})
                End If

                ' No $0.00 invoices need to get sent over
                If Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty) = 0 Then
                    numNonZeroDollarInvoices += 1
                    Continue For
                End If

                CUST_CODE = rowSOTINVH1.Item("CUST_CODE") & String.Empty

                rowSOTSVIA1 = tblSOTSVIA1.Rows.Find(rowSOTINVH1.Item("SHIP_VIA_CODE") & String.Empty)
                rowTATTERM1 = tblTATTERM1.Rows.Find(rowSOTINVH1.Item("TERM_CODE") & String.Empty)

                rowEDT810O1 = tblEDT810O1.NewRow

                EDI_OUTBOUND_DOC_NO = Me.CreateEDTSYSIH(EDI_OUR_ID, EDI_TP_ID, "IN", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)

                rowEDT810O1.Item("COMPANY_CODE") = COMPANY_CODE
                rowEDT810O1.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                rowEDT810O1.Item("EDI_INVOICE_DATE") = CDate(rowSOTINVH1.Item("INV_DATE") & String.Empty).ToString("MM/dd/yyyy")
                rowEDT810O1.Item("EDI_INVOICE_NUMBER") = rowSOTINVH1.Item("INV_NO")
                rowEDT810O1.Item("EDI_PO_DATE") = CDate(rowSOTORDR1.Item("ORDR_DATE") & String.Empty).ToString("MM/dd/yyyy")
                rowEDT810O1.Item("EDI_PO_NO") = rowSOTINVH1.Item("ORDR_CUST_PO")
                rowEDT810O1.Item("EDI_DEPT_NO") = rowSOTORDR1.Item("ORDR_DEPT")
                rowEDT810O1.Item("EDI_ORDER_NO") = rowSOTORDR1.Item("ORDR_NO")

                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty
                rowARTCUST2 = ASCDATA1.GetDataRow("SELECT * FROM ARTCUST2 WHERE CUST_CODE = :PARM1 AND CUST_STORE_NO = :PARM2", "VV", New Object() {CUST_CODE, CUST_STORE_NO})
                If mkNumChars > 0 AndAlso IsNumeric(CUST_STORE_NO) Then
                    CUST_STORE_NO = CUST_STORE_NO.PadLeft(mkNumChars, "0")
                    CUST_STORE_NO = StrReverse(StrReverse(CUST_STORE_NO).Substring(0, mkNumChars))
                End If

                If rowARTOPEN1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_TERMS_DISC_DATE") = rowARTOPEN1.Item("INV_DISC_DATE")
                End If

                rowEDT810O1.Item("EDI_BILL_TO") = CUST_STORE_NO
                rowEDT810O1.Item("EDI_SHIP_TO") = SHIP_ADDR_CODE
                rowEDT810O1.Item("EDI_MARK_FOR") = CUST_STORE_NO

                rowEDT810O1.Item("EDI_REMIT_TO_NAME") = rowARTPARM1.Item("AR_PARM_REMIT_NAME") & String.Empty
                rowEDT810O1.Item("EDI_REMIT_TO_ID") = rowARTPARM1.Item("AR_PARM_DUNS_NO") & String.Empty

                If Not foreignExchange Then
                    rowEDT810O1.Item("EDI_TOTAL_INV_AMT") = rowSOTINVH1.Item("INV_TOTAL_AMOUNT")
                Else
                    rowEDT810O1.Item("EDI_TOTAL_INV_AMT") = Val(ASCDATA1.GetDataValue("SELECT SUM(ORDR_QTY_SHIP * ORDR_UNIT_PRICE_CURR)" _
                                                                                      & " FROM SOTINVH2" _
                                                                                      & " WHERE INV_TYPE = '" & rowSOTINVH1.Item("INV_TYPE") & "'" _
                                                                                      & " AND INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "'") & String.Empty)
                End If
                rowEDT810O1.Item("EDI_BL_NO") = BillOfLading ' SHIP_BOL_NO
                rowEDT810O1.Item("EDI_FRT_TERMS") = IIf(rowSOTSHIP1.Item("FRT_TERMS") & String.Empty = "COL", "CC", "PP")
                'ALTER TABLE EDT810O1 ADD EDI_TERMS_DESC VARCHAR2(35);

                If rowSOTSVIA1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_ROUTING") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    rowEDT810O1.Item("EDI_SCAC_CODE") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
                    rowEDT810O1.Item("EDI_CARRIER_MODE") = rowSOTSVIA1.Item("CARRIER_MODE") & String.Empty
                Else
                    rowEDT810O1.Item("EDI_ROUTING") = "Unknown Shipper"
                    rowEDT810O1.Item("EDI_SCAC_CODE") = String.Empty
                    rowEDT810O1.Item("EDI_CARRIER_MODE") = "M"
                End If

                rowEDT810O1.Item("EDI_TOTAL_UNITS") = Val(tblSOTCART1.Compute("SUM(CART_TOTAL_UNITS)", "") & String.Empty)
                rowEDT810O1.Item("EDI_WEIGHT") = Val(tblSOTCART1.Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "") & String.Empty)

                ' Added 06/27/2024
                ' Clarins newest thing is to send us cartons with a weight of 99999 which blows up Gentran when processing the 810s.
                If (rowEDT810O1.Item("EDI_WEIGHT") & String.Empty) > 99999 Then
                    ASCMAIN1.sql = "SELECT SHIPHDR.OHWGHT EDI_WEIGHT
                                        FROM CONV.CFG_SHIPHDR SHIPHDR, SOTPICK1, SOTINVH1
                                        WHERE SOTPICK1.INV_NO = SOTINVH1.INV_NO
                                        AND SOTPICK1.PICK_NO = SHIPHDR.ABSPICKNBR
                                        AND SOTINVH1.INV_NO = :PARM1"
                    Dim rowShipHdr As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", {rowSOTINVH1.Item("INV_NO") & String.Empty})
                    If rowShipHdr IsNot Nothing Then
                        rowEDT810O1.Item("EDI_WEIGHT") = Val(rowShipHdr.Item("EDI_WEIGHT") & String.Empty)
                    End If

                    If (rowEDT810O1.Item("EDI_WEIGHT") & String.Empty) <= 0 Then
                        rowEDT810O1.Item("EDI_WEIGHT") = 9999
                    ElseIf (rowEDT810O1.Item("EDI_WEIGHT") & String.Empty) > 99999 Then
                        rowEDT810O1.Item("EDI_WEIGHT") = 9999
                    End If
                End If

                rowEDT810O1.Item("EDI_SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF") & String.Empty

                If rowTATTERM1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_TERMS_NET_DAYS") = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & String.Empty)
                    rowEDT810O1.Item("EDI_TERMS_DISC_DAYS_DUE") = Val(rowTATTERM1.Item("TERM_DAYS_DISC") & String.Empty)
                    rowEDT810O1.Item("EDI_TERMS_DISC_PCT") = Val(rowTATTERM1.Item("TERM_DISC_PERC") & String.Empty)
                    rowEDT810O1.Item("EDI_TERMS_DESC") = rowTATTERM1.Item("TERM_DESC") & String.Empty

                    Dim INV_DUE_DATE As Date
                    Dim DISC_DUE_DATE As Date
                    CALC_DUE_DATE(rowTATTERM1, rowSOTINVH1.Item("INV_DATE"), INV_DUE_DATE, DISC_DUE_DATE)
                    rowEDT810O1.Item("EDI_TERMS_DUE_DATE") = INV_DUE_DATE
                Else
                    rowEDT810O1.Item("EDI_TERMS_NET_DAYS") = 30 ' DEFAULT
                    rowEDT810O1.Item("EDI_TERMS_DISC_DAYS_DUE") = 30
                    rowEDT810O1.Item("EDI_TERMS_DISC_PCT") = 0
                    rowEDT810O1.Item("EDI_TERMS_DESC") = String.Empty
                    rowEDT810O1.Item("EDI_TERMS_DUE_DATE") = rowSOTINVH1.Item("INV_DATE")
                End If

                rowEDT810O1.Item("INIT_DATE") = DateTime.Now
                rowEDT810O1.Item("INIT_OPER") = ASCMAIN1.USER_ID

                rowEDT810O1.Item("EDI_CARTON_CT") = tblSOTCART1.Rows.Count

                ' This allows for Nordrack and other similar situations
                If tblEDTTRPM3.Select("CUST_CODE = '" & IIf(CUST_CODE_810.Length > 0, CUST_CODE_810, CUST_CODE) & "'").Length > 0 Then
                    Dim rowEDTTRMP3 As DataRow = tblEDTTRPM3.Select("CUST_CODE = '" & IIf(CUST_CODE_810.Length > 0, CUST_CODE_810, CUST_CODE) & "'")(0)
                    Dim TP_QUAL As String = rowEDTTRMP3.Item("EDI_TP_QUAL") & String.Empty
                    Dim TP_ID As String = rowEDTTRMP3.Item("EDI_TP_ID") & String.Empty

                    If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length > 0 Then
                        rowEDT810O1.Item("EDI_SUPPLIER_NO") = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0).Item("EDI_ACCT_REF_NO") & String.Empty
                    End If
                ElseIf tblEDTTRPM2.Select("CUST_CODE = '" & IIf(CUST_CODE_810.Length > 0, CUST_CODE_810, CUST_CODE) & "'").Length > 0 Then
                    Dim rowEDTTRMP2 As DataRow = tblEDTTRPM2.Select("CUST_CODE = '" & IIf(CUST_CODE_810.Length > 0, CUST_CODE_810, CUST_CODE) & "'")(0)
                    Dim TP_QUAL As String = rowEDTTRMP2.Item("EDI_TP_QUAL") & String.Empty
                    Dim TP_ID As String = rowEDTTRMP2.Item("EDI_TP_ID") & String.Empty

                    If tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'").Length > 0 Then
                        rowEDT810O1.Item("EDI_SUPPLIER_NO") = tblEDTTRPM1.Select("EDI_TP_QUAL = '" & TP_QUAL & "' and EDI_TP_ID = '" & TP_ID & "'")(0).Item("EDI_ACCT_REF_NO") & String.Empty
                    End If
                ElseIf tblEDTTRPM1.Select("CUST_CODE = '" & IIf(CUST_CODE_810.Length > 0, CUST_CODE_810, CUST_CODE) & "'").Length > 0 Then
                    rowEDT810O1.Item("EDI_SUPPLIER_NO") = tblEDTTRPM1.Select("CUST_CODE = '" & IIf(CUST_CODE_810.Length > 0, CUST_CODE_810, CUST_CODE) & "'")(0).Item("EDI_ACCT_REF_NO") & String.Empty
                End If

                'sql = " SELECT EDT850T1.* "
                'sql &= " FROM EDT850T1, SOTORDR1, SOTPICK1"
                'sql &= " WHERE EDT850T1. EDI_DOC_SEQ_NO = SOTORDR1.EDI_DOC_SEQ_NO"
                'sql &= " AND SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO"
                'sql &= " AND SOTPICK1.SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                'rowEDT850T1 = ASCDATA1.GetDataRow(sql)
                If rowEDT850T1 IsNot Nothing Then
                    rowEDT810O1.Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")
                    'If COMPANY_CODE = "INT" AndAlso rowEDT810O1.Item("EDI_SUPPLIER_NO") & String.Empty = String.Empty Then
                    ' Allow for Ahaha as well - march 2016
                    If rowEDT810O1.Item("EDI_SUPPLIER_NO") & String.Empty = String.Empty Then
                        rowEDT810O1.Item("EDI_SUPPLIER_NO") = rowEDT850T1.Item("EDI_SUPPLIER_NO") & String.Empty
                    End If
                End If

                tblEDT810O1.Rows.Add(rowEDT810O1)

                For Each rowSOTINVH2 As DataRow In tblSOTINVH2.Select("", "ITEM_CODE")
                    EDI_BUYER_ITEM = String.Empty
                    EDI_DOC_SEQ_NO = String.Empty
                    EDI_DTL_SEQ = 0
                    rowSOTORDR2 = Nothing

                    rowEDT810O2 = tblEDT810O2.NewRow
                    rowEDT810O2.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT810O2.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT810O2.Item("EDI_DOC_LNO") = rowSOTINVH2.Item("INV_LNO")
                    rowEDT810O2.Item("EDI_QTY_INVOICED") = rowSOTINVH2.Item("ORDR_QTY_SHIP")

                    If Not foreignExchange Then
                        rowEDT810O2.Item("EDI_UNIT_PRICE") = rowSOTINVH2.Item("ORDR_UNIT_PRICE")
                    Else
                        rowEDT810O2.Item("EDI_UNIT_PRICE") = rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR")
                    End If
                    ITEM_CODE = rowSOTINVH2.Item("ITEM_CODE")

                    ' Just incase one day there is an item on the invoice that is not on the order
                    ' I know it will Never ever ever happen - Yeah right!!
                    If tblSOTORDR2.Select("ITEM_CODE = '" & ITEM_CODE & "'").Length > 0 Then
                        rowSOTORDR2 = tblSOTORDR2.Select("ITEM_CODE = '" & ITEM_CODE & "'")(0)
                        EDI_DOC_SEQ_NO = rowSOTORDR2.Item("EDI_DOC_SEQ_NO") & String.Empty
                        EDI_DTL_SEQ = Val(rowSOTORDR2.Item("EDI_DTL_SEQ") & String.Empty)

                        If EDI_DOC_SEQ_NO.Length = 0 OrElse EDI_DTL_SEQ = 0 Then
                            Continue For
                        End If
                    End If

                    rowEDT810O2.Item("EDI_UOM") = (rowSOTORDR2.Item("EDI_PRICE_UOM") & String.Empty).ToString.Trim  ' EDI_PRICE_UOM / EDI_PO4_UOM
                    If rowEDT810O2.Item("EDI_UOM") & String.Empty = String.Empty Then
                        rowEDT810O2.Item("EDI_UOM") = "EA"
                    End If

                    rowEDT810O2.Item("EDI_BUYER_STYLE") = (rowSOTORDR2.Item("EDI_STYLE") & String.Empty).ToString.Trim
                    ' Added 6/23/2014
                    rowEDT810O2.Item("EDI_PO_LNO") = rowSOTORDR2.Item("EDI_PO_LNO")

                    ' If EDI did not come in with UPC Code then use our UPC Code
                    If (rowSOTORDR2.Item("EDI_UPC") & String.Empty).ToString.Trim.Length > 0 Then
                        rowEDT810O2.Item("EDI_ITEM_UP") = (rowSOTORDR2.Item("EDI_UPC") & String.Empty).ToString.Trim
                    Else
                        rowEDT810O2.Item("EDI_ITEM_UP") = (rowSOTINVH2.Item("ITEM_UPC_CODE") & String.Empty).ToString.Trim
                    End If

                    rowEDT810O2.Item("EDI_ITEM_EN") = (rowSOTORDR2.Item("EDI_EAN") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_ITEM_GTIN") = (rowSOTORDR2.Item("EDI_GTIN") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_BUYER_ITEM") = (rowSOTORDR2.Item("EDI_SKU") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_SIZE_CODE") = (rowSOTORDR2.Item("EDI_SIZE_CODE") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_SIZE_DESC") = (rowSOTORDR2.Item("EDI_SIZE_DESC") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_COLOR_CODE") = (rowSOTORDR2.Item("EDI_COLOR_CODE") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_COLOR_NAME") = (rowSOTORDR2.Item("EDI_COLOR_NAME") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_SELLER_ITEM") = rowSOTINVH2.Item("ITEM_CODE") & String.Empty
                    'rowEDT810O2.Item("EDI_ITEM_DESC") = (rowSOTORDR2.Item("EDI_ITEM_DESC") & String.Empty).ToString.Trim
                    If (rowSOTORDR2.Item("EDI_ITEM_DESC") & String.Empty).ToString.Trim.Length > 0 Then
                        rowEDT810O2.Item("EDI_ITEM_DESC") = (rowSOTORDR2.Item("EDI_ITEM_DESC") & String.Empty).ToString.Trim
                    Else
                        rowEDT810O2.Item("EDI_ITEM_DESC") = (rowSOTINVH2.Item("ITEM_DESC") & String.Empty).ToString.Trim
                    End If

                    rowEDT810O2.Item("EDI_PO4_UOM") = (rowSOTORDR2.Item("EDI_PO4_UOM") & String.Empty).ToString.Trim
                    rowEDT810O2.Item("EDI_PO4_QTY") = Val(rowSOTORDR2.Item("EDI_PO4_QTY") & String.Empty)
                    If Val(rowSOTORDR2.Item("EDI_PO4_QTY") & String.Empty) > 1 AndAlso rowSOTORDR2.Item("EDI_PO4_INNER") & String.Empty = String.Empty Then
                        rowEDT810O2.Item("EDI_PO4_INNER") = 1
                    ElseIf Val(rowSOTORDR2.Item("EDI_PO4_INNER") & String.Empty) > 0 Then
                        rowEDT810O2.Item("EDI_PO4_INNER") = Val(rowSOTORDR2.Item("EDI_PO4_INNER") & String.Empty)
                    End If

                    tblEDT810O2.Rows.Add(rowEDT810O2)
                Next

                Dim charge As Decimal = Val(rowSOTINVH1.Item("INV_FREIGHT") & String.Empty)
                If charge <> 0 Then
                    rowEDT810O3 = tblEDT810O3.NewRow
                    rowEDT810O3.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT810O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT810O3.Item("EDI_SAC_LNO") = 1
                    rowEDT810O3.Item("EDI_CHG_ALL_IND") = IIf(charge >= 0, "C", "A")
                    rowEDT810O3.Item("EDI_CHG_ALL_CODE") = "D240"
                    rowEDT810O3.Item("EDI_SAC_AMOUNT") = Math.Abs(charge)
                    rowEDT810O3.Item("EDI_SAC_DESC") = "FREIGHT"
                    tblEDT810O3.Rows.Add(rowEDT810O3)
                End If

                For Each rowSOTORDR5 In tblSOTORDR5.Select("")
                    rowEDT810O5 = tblEDT810O5.NewRow
                    rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT810O5.Item("EDI_ADDR_TYPE") = rowSOTORDR5.Item("CUST_ADDR_TYPE")
                    rowEDT810O5.Item("EDI_NAME") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                    rowEDT810O5.Item("EDI_ADDRESS1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                    rowEDT810O5.Item("EDI_ADDRESS2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                    rowEDT810O5.Item("EDI_ADDRESS3") = rowSOTORDR5.Item("CUST_ADDR3") & String.Empty
                    rowEDT810O5.Item("EDI_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                    rowEDT810O5.Item("EDI_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                    rowEDT810O5.Item("EDI_ZIPCODE") = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                    rowEDT810O5.Item("EDI_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty

                    Dim EDI_ADDR_CODE As String = IIf(rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT", rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty, String.Empty)
                    If rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT" Then
                        If mkNumChars > 0 AndAlso IsNumeric(EDI_ADDR_CODE) Then
                            EDI_ADDR_CODE = EDI_ADDR_CODE.PadLeft(mkNumChars, "0")
                            EDI_ADDR_CODE = StrReverse(StrReverse(EDI_ADDR_CODE).Substring(0, mkNumChars))
                        End If
                    ElseIf rowARTCUST2 IsNot Nothing AndAlso rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty <> String.Empty Then
                        EDI_ADDR_CODE = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty
                    Else
                        EDI_ADDR_CODE = SHIP_ADDR_CODE
                    End If

                    rowEDT810O5.Item("EDI_ADDR_CODE") = EDI_ADDR_CODE
                    'rowEDT810O5.Item("EDI_ADDR_CODE_QUAL") = IIf(rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT", "91", "92")
                    tblEDT810O5.Rows.Add(rowEDT810O5)
                Next

                ' Create MK record for Ahava - Currently used only for Steinmart, but wil generate for all customers
                If ASCMAIN1.SOLUTION = "AHA" AndAlso rowARTCUST2 IsNot Nothing Then
                    rowEDT810O5 = tblEDT810O5.NewRow
                    rowEDT810O5.Item("COMPANY_CODE") = COMPANY_CODE
                    rowEDT810O5.Item("EDI_OUTBOUND_DOC_NO") = EDI_OUTBOUND_DOC_NO
                    rowEDT810O5.Item("EDI_ADDR_TYPE") = "MK"
                    rowEDT810O5.Item("EDI_NAME") = rowARTCUST2.Item("CUST_STORE_NAME") & String.Empty
                    rowEDT810O5.Item("EDI_ADDRESS1") = rowARTCUST2.Item("CUST_STORE_ADDR1") & String.Empty
                    rowEDT810O5.Item("EDI_ADDRESS2") = rowARTCUST2.Item("CUST_STORE_ADDR2") & String.Empty
                    rowEDT810O5.Item("EDI_ADDRESS3") = rowARTCUST2.Item("CUST_STORE_ADDR3") & String.Empty
                    rowEDT810O5.Item("EDI_CITY") = rowARTCUST2.Item("CUST_STORE_CITY") & String.Empty
                    rowEDT810O5.Item("EDI_STATE") = rowARTCUST2.Item("CUST_STORE_STATE") & String.Empty
                    rowEDT810O5.Item("EDI_ZIPCODE") = rowARTCUST2.Item("CUST_STORE_ZIP_CODE") & String.Empty
                    rowEDT810O5.Item("EDI_COUNTRY") = rowARTCUST2.Item("CUST_STORE_COUNTRY") & String.Empty

                    Dim EDI_ADDR_CODE As String = rowARTCUST2.Item("CUST_STORE_NO") & String.Empty
                    If rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty <> String.Empty Then
                        EDI_ADDR_CODE = rowARTCUST2.Item("GLOBAL_LOCATION_NUMBER") & String.Empty
                    Else
                        If mkNumChars > 0 AndAlso IsNumeric(EDI_ADDR_CODE) Then
                            EDI_ADDR_CODE = EDI_ADDR_CODE.PadLeft(mkNumChars, "0")
                            EDI_ADDR_CODE = StrReverse(StrReverse(EDI_ADDR_CODE).Substring(0, mkNumChars))
                        End If
                    End If

                    rowEDT810O5.Item("EDI_ADDR_CODE") = EDI_ADDR_CODE
                    tblEDT810O5.Rows.Add(rowEDT810O5)

                End If
            Next
        Next

        If numNonZeroDollarInvoices > 0 AndAlso EDI_OUTBOUND_DOC_NO.Length = 0 Then
            EDI_OUTBOUND_DOC_NO = "Zero"
        End If
    End Sub

    ''' <summary>
    ''' Creates a record for table EDTSYSIH
    ''' </summary>
    ''' <param name="EDI_OUR_ID"></param>
    ''' <param name="EDI_TP_ID"></param>
    ''' <returns>The key field EDI_OUTBOUND_DOC_NO for table EDTSYSIH</returns>
    ''' <remarks></remarks>
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
        tblEDTSYSIH.Rows.Add(rowEDTSYSIH)

        CreateEDTSYSIH = ediOutboundDocNo

    End Function

    Private Sub CALC_DUE_DATE(ByVal rowTATTERM1 As DataRow, ByVal InvoiceDate As Date,
                      ByRef INV_DUE_DATE As Date, ByRef DISC_DUE_DATE As Date)

        Dim INV_BASE_DATE As Date = CDate(InvoiceDate)

        Select Case rowTATTERM1.Item("TERM_DUE_TYPE") & String.Empty
            Case "D"
                INV_DUE_DATE = INV_BASE_DATE.AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DUE") & ""))

            Case "E"
                Dim ADD_MONTHS_BASE As Integer = 1
                Dim TERM_CUTOFF_DAY As Integer = Val(rowTATTERM1.Item("TERM_CUTOFF_DAY") & "")
                Dim BASE_DD As Integer = Val(Format(INV_BASE_DATE, "dd"))
                Dim TERM_DAYS_DUE As Integer = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & "")
                Dim TERM_ADDL_MOS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_MOS") & "")
                Dim INV_BASE_DATEx As String = Format(INV_BASE_DATE, "MM/dd/yyyy")

                Select Case rowTATTERM1.Item("TERM_EOM_TYPE") & String.Empty
                    Case "F"
                        ASCMAIN1.sql = "Select GLTPARM2.* " _
                         & " from GLTPARM2 " _
                         & " where OPS_YYYYPP = " _
                         & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
                         & "  where GLTPARM2.PRD_END_DATE >= '" & Format(INV_BASE_DATE, "dd-MMM-yyyy") & "')"
                        Dim rowGLTPARM2 As DataRow = ASCDATA1.GetDataRow
                        Dim YYYYMM As String = ASCMAIN1.Get_YYYYMM(rowGLTPARM2.Item("OPS_YYYYPP"), 0)
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "C"
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        If TERM_DAYS_DUE > 31 Then TERM_DAYS_DUE = 30
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "S"
                        If BASE_DD <= TERM_CUTOFF_DAY _
                        And BASE_DD <= TERM_DAYS_DUE Then
                            ADD_MONTHS_BASE = 0
                        End If
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                End Select
                If TERM_ADDL_MOS > 0 Then
                    INV_DUE_DATE = INV_DUE_DATE.AddMonths(TERM_ADDL_MOS)
                End If
        End Select

        If Val(rowTATTERM1.Item("TERM_DISC_PERC") & "") <> 0 Then
            If rowTATTERM1.Item("TERM_DISC_ELIG_DUE") & String.Empty = "1" Then
                DISC_DUE_DATE = INV_DUE_DATE
            Else
                If Val(rowTATTERM1.Item("TERM_DISC_PERC") & "") <> 0 Then
                    DISC_DUE_DATE = DateValue(INV_DUE_DATE & "").AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DISC") & ""))
                End If
            End If
        End If
    End Sub


End Class
