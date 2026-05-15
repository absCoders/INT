Imports System.Text.RegularExpressions
Imports System.Linq

Public MustInherit Class ShippingLabel

    Public Sub PrintLabel(Optional ByVal printQty As Integer = 1)
        Dim labelData As Dictionary(Of String, DataRow) = GetLabelData()
        Dim labelTemplate As String = GetLabelTemplate()

        For i As Integer = 1 To printQty
            ChangeLabelData(labelData, i, printQty)
            Dim labeltoPrint As String = FillLabelTemplateWithData(labelTemplate, labelData)
            If Not SendToLabelPrinter(labeltoPrint) Then
                If MessageBox.Show("An error occurred printing the label. Do you want to Abort printing any other Address labels?" _
                                     , "Address Labels", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.Yes Then
                    Exit For
                End If
            End If
        Next
    End Sub

    Protected MustOverride Function GetLabelData() As Dictionary(Of String, DataRow)
    Protected MustOverride Function GetLabelTemplate() As String

    Protected Overridable Sub ChangeLabelData(ByVal labelData As Dictionary(Of String, DataRow), ByVal currentIndex As Integer, ByVal lastIndex As Integer)

    End Sub

    Private Function FillLabelTemplateWithData(labelTemplate As String, labelData As Dictionary(Of String, DataRow)) As String
        'Matches <<TABLE.COLUMN>...>, and if the value of TABLE.COLUMN is null, it omits this line from the ZPL
        'Used for hiding a section of label if the data is unavailable
        labelTemplate = Regex.Replace(labelTemplate, "\<\<(?<table>[\w_]+)\.(?<column>[\w_]+)\>(?<command>.*)\>", _
                        Function(m) If(labelData(m.Groups("table").Value).Item(m.Groups("column").Value) & "" = "",
                                       "",
                                       m.Groups("command").Value))

        'Regex matches {TABLE.COLUMN} and replaces with values from rowUCC128
        labelTemplate = Regex.Replace(labelTemplate, "\{(?<table>[\w_]+)\.(?<column>[\w_]+)\}", _
                        Function(m) labelData(m.Groups("table").Value).Item(m.Groups("column").Value) & "")

        Return labelTemplate
    End Function

    Public Shared Function SendToLabelPrinter(ByVal labelData As String) As Boolean
        Try
            If ASCMAIN1.Running_in_VS Then
                PrintShippingLabelFromDevMachine(labelData)
            Else
                ASCMAIN1.LabelPrinterSerialPort.WriteLine(labelData)
            End If

            Return True
        Catch ex As Exception
            MessageBox.Show("Error Printing Label: " & ex.Message)
            Return False
        End Try
    End Function

    Private Shared Function PrintShippingLabelFromDevMachine(ByVal labelData As String) As Boolean
        If ASCMAIN1.USER_ID <> "edz" Then
            If MessageBox.Show(labelData, "Continue with label print?", MessageBoxButtons.YesNo) = DialogResult.No Then
                Return True
            End If
        End If

        Dim zebraPrinter As String = FindZebraPrinter()

        Dim vLabelPrinter As New ASCPRINT
        Return vLabelPrinter.SendStringToPrinter(zebraPrinter, labelData)
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
End Class

Public Class CartonLabel
    Inherits ShippingLabel

    Private Property cartonNo

    Public Sub New(ByVal cartonNo As String)
        Me.cartonNo = cartonNo
    End Sub

    Protected Overrides Function GetLabelData() As System.Collections.Generic.Dictionary(Of String, System.Data.DataRow)
        ASCMAIN1.sql = "SELECT X.*,WH1.*,O5.*,AC1.CUST_VEND_REF,SUBSTR(X.CART_NO,11,9) CART_NO_9," _
                    & " SUBSTR(X.CART_NO,20,1) CART_NO_DIGIT," _
                    & " SUBSTR(X.CART_NO,5,6) CART_NO_PFX," _
                    & " TRUNC(SYSDATE) CURRENT_DATE, " _
                    & " X.CART_SERIAL_NO || ' of ' || X.CART_SEQ_MAX CART_1_OF_9 FROM" _
                    & " (SELECT ROW_NUMBER() OVER (ORDER BY C1.CART_NO) CART_SERIAL_NO,C1.CART_NO,C1.PICK_NO,C1.CART_TOTAL_UNITS, ET1.EDI_MERCH_TYPE, " _
                    & " COUNT(*) OVER () CART_SEQ_MAX,SUM(C2.QTY_PACKED) CART_QTY_PACKED, " _
                    & " MAX(I1.ITEM_CODE) ITEM_CODE, " _
                    & " MAX(TO_CHAR(I1.XMIT_DATE_3PL,'MM/DD/YYYY')) XMIT_DATE_3PL, " _
                    & " MAX(TO_CHAR(I1.XMIT_DATE_3PL,'MMDDYYYY')) XMIT_DATE_3PL_BC, " _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN 'Mixed' ELSE MAX(I1.ITEM_UPC_CODE) END ITEM_UPC_CODE, " _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN '' ELSE MAX(I1.ITEM_UPC_CODE) END ITEM_UPC_CODE_ONLY, " _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN 'Pick And Pack' ELSE MAX(I1.ITEM_DESC) END ITEM_DESC, " _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN '' ELSE TO_CHAR(SUM(C2.QTY_PACKED)) END CART_QTY_PACKED_ONLY," _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN '' ELSE TO_CHAR(MAX(ET2.EDI_PO4_QTY)) END EDI_PO4_QTY, " _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN '' ELSE MAX(ET2.EDI_SKU) END EDI_SKU, " _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN '' ELSE MAX(ET2.EDI_STYLE) END EDI_STYLE, " _
                    & " CASE WHEN COUNT(DISTINCT I1.ITEM_CODE) > 1 THEN '' ELSE MAX(ET2.EDI_ITEM) END EDI_ITEM, " _
                    & " MAX(CI.CUST_ITEM_CODE) CUST_ITEM_CODE, " _
                    & " MAX(O1.ORDR_CUST_PO) ORDR_CUST_PO, " _
                    & " MAX(O1.ORDR_NO) ORDR_NO, " _
                    & " MAX(O1.CUST_CODE) CUST_CODE, " _
                    & " MAX(O1.WHSE_CODE) WHSE_CODE, " _
                    & " MAX(O1.ORDR_SHIP_DATE) ORDR_SHIP_DATE FROM " _
                    & " SOTCART1 C1 JOIN SOTCART2 C2 ON (C1.CART_NO=C2.CART_NO) JOIN " _
                    & " SOTORDR1 O1 ON (C2.ORDR_NO=O1.ORDR_NO) JOIN " _
                    & " SOTORDR2 O2 ON (C2.ORDR_NO=O2.ORDR_NO AND C2.ORDR_LNO=O2.ORDR_LNO) JOIN " _
                    & " ICTITEM1 I1 ON (C2.ITEM_CODE=I1.ITEM_CODE) LEFT JOIN " _
                    & " EDT850T1 ET1 ON (O2.EDI_DOC_SEQ_NO=ET1.EDI_DOC_SEQ_NO) LEFT JOIN" _
                    & " EDT850T2 ET2 ON (O2.EDI_DOC_SEQ_NO=ET2.EDI_DOC_SEQ_NO AND O2.EDI_DTL_SEQ=ET2.EDI_DTL_SEQ) LEFT JOIN " _
                    & " SOTCITM1 CI ON (O2.ITEM_CODE=CI.ITEM_CODE AND O2.CUST_CODE=CI.CUST_CODE) " _
                    & " WHERE C1.PICK_NO=(SELECT PICK_NO FROM SOTCART1 WHERE CART_NO=:PARM1) " _
                    & " GROUP BY C1.CART_NO,C1.PICK_NO,C1.CART_TOTAL_UNITS,ET1.EDI_MERCH_TYPE) X" _
                    & " JOIN ICTWHSE1 WH1 ON (X.WHSE_CODE=WH1.WHSE_CODE) " _
                    & " JOIN SOTORDR5 O5 ON (X.ORDR_NO=O5.ORDR_NO AND O5.CUST_ADDR_TYPE='ST') " _
                    & " JOIN ARTCUST1 AC1 ON (X.CUST_CODE=AC1.CUST_CODE)" _
                    & " WHERE CART_NO=:PARM1"
        Dim rowSOTCART1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {cartonNo})

        ASCMAIN1.sql = "Select SOTPICK1.*, SOTORDR1.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME, ARTCUST2.GLOBAL_LOCATION_NUMBER, " _
                    & " NVL(SOTORDR0.ORDR_DEPT,SOTSHIP1.ORDR_DEPT) ORDR_DEPT,SOTSHIP1.*,SOTORDR0.CUST_CODE,SOTORDR0.ORDR_CUST_PO, " _
                    & " COALESCE(SOTSVIA1.SHIP_VIA_DESC,SOTSHIP1.SHIP_VIA_CODE,'') SHIP_VIA_DESC, " _
                    & " SOTSVIA1.SHIP_VIA_SCAC, SOTSHIP1.SHIP_REF, ARTCUST2.CUST_STORE_MARK_FOR," _
                    & " SUBSTR(SOTORDR1.CUST_STORE_NO,-1 * EDTSLSP1.NUMBER_CHARS_STORE) CUST_STORE_NO_X," _
                    & " SUBSTR(SOTORDR1.CUST_STORE_NO,-4) CUST_STORE_NO_4, EDTTRPM1.EDI_ACCT_REF_NO," _
                    & " COALESCE(SOTSHIPB.SHIPPED_ACTUAL,SOTSHIP1.SHIP_DATE_PLANNED,TRUNC(SYSDATE)) SHIP_DATE" _
                    & " from SOTCART1 " _
                    & " JOIN SOTPICK1 ON (SOTCART1.PICK_NO=SOTPICK1.PICK_NO)" _
                    & " JOIN SOTORDR1 ON (SOTPICK1.ORDR_NO=SOTORDR1.ORDR_NO)" _
                    & " JOIN SOTSHIP1 ON (SOTPICK1.SHIP_BOL_NO=SOTSHIP1.SHIP_BOL_NO)" _
                    & " LEFT JOIN ARTCUST2 ON (SOTORDR1.CUST_CODE=ARTCUST2.CUST_CODE AND SOTORDR1.CUST_STORE_NO=ARTCUST2.CUST_STORE_NO)" _
                    & " LEFT JOIN EDTSLSP1 ON (SOTORDR1.CUST_CODE=EDTSLSP1.CUST_CODE)" _
                    & " LEFT JOIN SOTSHIPB ON (SOTSHIP1.BILL_OF_LADING_NO=SOTSHIPB.BOL_NO)" _
                    & " LEFT JOIN SOTORDR0 ON (SOTSHIP1.ORDR_GROUP_NO=SOTORDR0.ORDR_GROUP_NO)" _
                    & " LEFT JOIN SOTSVIA1 ON (SOTSHIP1.SHIP_VIA_CODE=SOTSVIA1.SHIP_VIA_CODE)" _
                    & " LEFT JOIN EDTTRPM1 ON (SOTORDR1.CUST_CODE=EDTTRPM1.CUST_CODE AND EDI_DOC_NO=856) " _
                    & " where " _
                    & " SOTCART1.CART_NO=:PARM1"
        Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {cartonNo})

        Dim labelData As New Dictionary(Of String, DataRow)
        labelData.Add("SOTCART1", rowSOTCART1)
        labelData.Add("SOTPICK1", rowSOTPICK1)
        labelData.Add("SOTORDR1", rowSOTCART1)
        labelData.Add("SOTSHIPX", rowSOTPICK1)
        labelData.Add("ICTWHSE1", rowSOTCART1)
        labelData.Add("SOTORDR5", rowSOTCART1)
        labelData.Add("ARTCUST1", rowSOTCART1)
        Return labelData
    End Function

    Protected Overrides Function GetLabelTemplate() As String
        Dim labelTemplate As String = ASCDATA1.GetDataValue( _
            "SELECT UCC128_COMMANDS FROM " & _
            " SOTCART1 C1 JOIN " & _
            " SOTPICK1 P1 ON (C1.PICK_NO=P1.PICK_NO) JOIN " & _
            " SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) JOIN " & _
            " ARTCUST1 AC ON (O1.CUST_CODE=AC.CUST_CODE) JOIN " & _
            " SOTUCCL1 U1 ON (AC.LABEL_TEMPLATE_CODE=U1.LABEL_TEMPLATE_CODE) " & _
            " WHERE C1.CART_NO=:PARM1", "V", New Object() {cartonNo}) & ""
        If labelTemplate = "" Then Throw New Exception("No UCC128 label template assigned for this customer")
        Return labelTemplate
    End Function
End Class

Public Class AddressLabel
    Inherits ShippingLabel

    Private Property pickNo As String
    Private Property labelComment As String
    Private Property labelTemplateCode As String
    Private Property rowSOTORDR5 As DataRow
    Private Property startPos As Integer = 0
    Private Property totalLabels As Integer = 0

    Public Sub New(ByVal pickNo As String, ByVal labelComment As String, ByVal labelTemplateCode As String, ByVal rowSOTORDR5 As DataRow)
        Me.pickNo = pickNo
        Me.labelComment = labelComment
        Me.rowSOTORDR5 = rowSOTORDR5
        Me.labelTemplateCode = labelTemplateCode
    End Sub

    Protected Overrides Function GetLabelData() As System.Collections.Generic.Dictionary(Of String, System.Data.DataRow)
        'retrieve data for label based off pick #

        ASCMAIN1.sql = "SELECT P1.PICK_NO,P1.SHIP_BOL_NO,O1.CUST_STORE_NO FROM " _
            & " SOTPICK1 P1 JOIN SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) " _
            & " WHERE P1.PICK_NO=:PARM1"
        Dim rowSOTPICK1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {pickNo})

        ASCMAIN1.sql = "Select SOTSHIP1.*,SOTORDR0.CUST_CODE,SOTORDR0.ORDR_CUST_PO,SOTORDR0.ORDR_DEPT" _
               & ", COALESCE(SOTSVIA1.SHIP_VIA_DESC,SOTSHIP1.SHIP_VIA_CODE,'') SHIP_VIA_DESC" _
               & " from SOTSHIP1,SOTORDR0,SOTSVIA1 " _
               & " where SOTORDR0.ORDR_GROUP_NO (+) = SOTSHIP1.ORDR_GROUP_NO" _
               & "   and SOTSVIA1.SHIP_VIA_CODE (+) = SOTSHIP1.SHIP_VIA_CODE" _
               & "   and SOTSHIP1.SHIP_BOL_NO = :PARM1"
        Dim rowSOTSHIPX As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {rowSOTPICK1.Item("SHIP_BOL_NO")})
        rowSOTSHIPX.Table.Columns.Add("ORDR_COMMENTS")
        rowSOTSHIPX.Table.Columns.Add("ADDR_1_OF_9")
        rowSOTSHIPX.Item("ORDR_COMMENTS") = labelComment

        ASCMAIN1.sql = "SELECT W1.* FROM" _
                    & " SOTPICK1 P1 JOIN" _
                    & " SOTORDR1 O1 ON (P1.ORDR_NO=O1.ORDR_NO) JOIN" _
                    & " ICTWHSE1 W1 ON (O1.WHSE_CODE=W1.WHSE_CODE) WHERE P1.PICK_NO=:PARM1"
        Dim rowICTWHSE1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {pickNo})

        Dim labelData As New Dictionary(Of String, DataRow)
        labelData.Add("SOTCART1", rowSOTSHIPX)
        labelData.Add("SOTPICK1", rowSOTPICK1)
        labelData.Add("SOTSHIPX", rowSOTSHIPX)
        labelData.Add("ICTWHSE1", rowICTWHSE1)
        labelData.Add("SOTORDR5", rowSOTORDR5)
        Return labelData
    End Function

    Public Sub Set1of9(ByVal startPos As Integer, ByVal totalLabels As Integer)
        Me.startPos = startPos
        Me.totalLabels = totalLabels
    End Sub

    Protected Overrides Sub ChangeLabelData(labelData As Dictionary(Of String, DataRow), currentIndex As Integer, lastIndex As Integer)
        If labelData("SOTCART1").Table.Columns.Contains("ADDR_1_OF_9") Then
            labelData("SOTCART1").Item("ADDR_1_OF_9") = CStr(currentIndex + If(startPos > 0, startPos, 1) - 1) & " of " & CStr(If(totalLabels > 0, totalLabels, lastIndex))
        End If
    End Sub
    
    Protected Overrides Function GetLabelTemplate() As String
        Dim labelTemplate As String = ASCDATA1.GetDataValue( _
            "SELECT UCC128_COMMANDS FROM " & _
            " SOTUCCL1 U1 " & _
            " WHERE U1.LABEL_TEMPLATE_CODE=:PARM1", "V", New Object() {labelTemplateCode}) & ""
        If labelTemplate = "" Then Throw New Exception("Label Template '" & labelTemplateCode & "' not found")
        Return labelTemplate
    End Function
End Class