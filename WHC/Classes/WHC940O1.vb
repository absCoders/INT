Public Class WHC940O1
    ' Create Anticipated Receipts File (Outbound) 943

    Inherits WHC000O1

    Private SOHDR As String = String.Empty
    Private SODTL As String = String.Empty
    Private salesOrderFiles As New List(Of String)

    Private sqlPickTickets As String = String.Empty
    Private listOfShipments As New List(Of String)

    Private tblICTCOLL1 As DataTable = Nothing
    Private tblICTITEM1 As DataTable = Nothing
    Private tblTATTERM1 As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing
    Private tblWholesale As DataTable = Nothing

    Private wkDirectory As String = Nothing
    Private badOrderGroupNos As New List(Of String)

    Private arrivalDateCustomers As List(Of String)

    ' Archive the files
    Private subDir As String = DateTime.Now.ToString("yyyyMM")

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHC940O1"
        clsSuccessfulExecution = False

        Dim clsTACMAIN1 As New TACMAIN1
        arrivalDateCustomers = clsTACMAIN1.IPLBMacysCustomerCodes

        Create_Work_Table()

        wkDirectory = ASCMAIN1.Folders("Work")
        If wkDirectory.Length > 0 AndAlso Not wkDirectory.EndsWith("\") Then
            wkDirectory &= "\"
        End If

        With dst

            tblTasks.Rows.Clear()

            Addtask("WHC940O1 Start WHC940O1 Load")

            sqlPickTickets = "Select SOTPICK1.*, SOTSHIP1.ORDR_GROUP_NO, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO" & vbCrLf _
                 & " from SOTPICK1, SOTSHIP1, ICTWHSE1, SOTORDR1" & vbCrLf _
                 & " where  SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                 & " and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                 & " and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
                 & " and SOTPICK1.PICK_STATUS = 'P'" & vbCrLf _
                 & " and NVL(SOTSHIP1.LP_STATUS, '0') = '0'" & vbCrLf _
                 & " and SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE" & vbCrLf _
                 & " and ICTWHSE1.LP_CODE is Not Null"

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*")
            dst.Tables("SOTORDR2").Columns.Add("COLLECTION_CODE", GetType(System.String))
            Create_TDA(.Tables.Add, "SOTORDR5", "*")

            Create_TDA(.Tables.Add, "SOTSHIP1", "*")

            Create_TDA(.Tables.Add, "SOTPICK1", "*")
            dst.Tables("SOTPICK1").Columns.Add("XREF_CUST_CODE_SOLD_TO", GetType(System.String))
            dst.Tables("SOTPICK1").Columns.Add("XREF_CUST_STORE_NO_SOLD_TO", GetType(System.String))
            dst.Tables("SOTPICK1").Columns.Add("XREF_CUST_CODE_SHIP_TO", GetType(System.String))
            dst.Tables("SOTPICK1").Columns.Add("XREF_CUST_STORE_NO_SHIP_TO", GetType(System.String))

            Create_TDA(.Tables.Add, "SOTPICK2", "*")
            dst.Tables("SOTPICK2").Columns.Add("ITEM_CODE", GetType(System.String))
            dst.Tables("SOTPICK2").Columns.Add("ITEM_ALT_SORT", GetType(System.String))
            dst.Tables("SOTPICK2").Columns.Add("XREF_BRAND_CODE", GetType(System.String))
            dst.Tables("SOTPICK2").Columns.Add("COLLECTION_CODE", GetType(System.String))
            dst.Tables("SOTPICK2").Columns.Add("EDI_SKU", GetType(System.String))

            Create_TDA(.Tables.Add, "WHTSHIPX", "*")
            Create_TDA(.Tables.Add, "WHTLPXN1", "*")

            dst.Tables("SOTPICK2").Columns.Add("EXT_PICK_UNIT_PRICE", GetType(System.Decimal), "ISNULL(PICK_QTY, 0) * ISNULL(PICK_UNIT_PRICE, 0)")
            Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
            dst.Tables("SOTPICK1").Columns.Add("TOT_EXT_PICK_UNIT_PRICE", GetType(System.Decimal), "SUM(CHILD.EXT_PICK_UNIT_PRICE)")

            ASCMAIN1.sql = "SELECT CUST_CODE FROM ARTCUST1,SOTTCLS1" _
                & " WHERE SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" _
                & " AND NVL(SOTTCLS1.CHANNEL_CODE,'?') in ('2', '4')"
            tblWholesale = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            Create_TDA(.Tables.Add, "EDT850T1", "*")
            Create_TDA(.Tables.Add, "EDT850T2", "*")

            Addtask("WHC940O1 Finish WHC940O1 Load")
        End With

        Main_Process()
    End Sub

    Public Sub Main_Process()

        Try
            EnforceConstraints(False)

            ' Global Variable that determines the number of records processed.
            R = 0

            Addtask("WHC940O1 Start Main_Process")

            ASCMAIN1.sql = String.Empty
            Dim numRecords As Int16 = 0
            badOrderGroupNos.Clear()

            Select Case G.APP_CMD & String.Empty
                Case "WHSE"
                    Dim whse_code As String = G.APP_KEY & String.Empty
                    ASCMAIN1.sql = sqlPickTickets
                    ASCMAIN1.sql &= " and SOTSHIP1.WHSE_CODE = '" & whse_code & "'"

                Case "SHIP"
                    Dim shipments As String = G.APP_KEY & String.Empty
                    ASCMAIN1.sql = sqlPickTickets
                    ASCMAIN1.sql &= " and SOTSHIP1.SHIP_BOL_NO IN (" & shipments & ")"

                Case Else
                    clsErrorMessage.Add("Missing Command Parameter")
                    Exit Sub
            End Select

            Addtask("WHC940O1 Create tempTableSOTPICKX")
            Dim tempTableSOTPICKX As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            Addtask("WHC940O1 Fill SOTORDR1")
            Fill_Records("SOTORDR1", String.Empty, True, "Select * from SOTORDR1 where ORDR_NO in (select ORDR_NO from " & tempTableSOTPICKX & ")")

            Addtask("WHC940O1 Fill SOTORDR2")
            Fill_Records("SOTORDR2", String.Empty, True, "Select * from SOTORDR2 where ORDR_NO in (select ORDR_NO from " & tempTableSOTPICKX & ")")

            Addtask("WHC940O1 Fill SOTORDR5")
            Fill_Records("SOTORDR5", String.Empty, True, "Select * from SOTORDR5 where ORDR_NO in (select ORDR_NO from " & tempTableSOTPICKX & ") and CUST_ADDR_TYPE = 'ST'")

            Addtask("WHC940O1 Fill SOTSHIP1")
            Fill_Records("SOTSHIP1", String.Empty, True, "Select * from SOTSHIP1 where SHIP_BOL_NO in (select SHIP_BOL_NO from " & tempTableSOTPICKX & ")")

            Addtask("WHC940O1 Fill SOTPICK1")
            Fill_Records("SOTPICK1", String.Empty, True, "Select * from SOTPICK1 where PICK_NO in (select PICK_NO from " & tempTableSOTPICKX & ")")

            Addtask("WHC940O1 Fill SOTPICK2")
            Fill_Records("SOTPICK2", String.Empty, True, "Select * from SOTPICK2 where PICK_NO in (select PICK_NO from " & tempTableSOTPICKX & ")")

            Addtask("WHC940O1 Fill EDT850T1")
            Fill_Records("EDT850T1", String.Empty, True, "Select * from EDT850T1 where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from SOTORDR1 WHERE ORDR_NO IN (Select ORDR_NO from " & tempTableSOTPICKX & "))")

            Addtask("WHC940O1 Fill EDT850T2")
            Fill_Records("EDT850T2", String.Empty, True, "Select * from EDT850T2 where EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO from SOTORDR1 WHERE ORDR_NO IN (Select ORDR_NO from " & tempTableSOTPICKX & "))")

            EnforceConstraints(True)

            ' Validate the data before we Create files for the warehouse.
            Addtask("WHC940O1 Fill ICTCOLL1")
            ASCMAIN1.sql = "SELECT * FROM ICTCOLL1 WHERE COLLECTION_CODE IN (SELECT COLLECTION_CODE FROM ICTITEM1 WHERE ITEM_CODE IN " _
                & " (SELECT ITEM_CODE FROM SOTORDR2 WHERE ORDR_NO IN (SELECT ORDR_NO from " & tempTableSOTPICKX & ")))"
            tblICTCOLL1 = ASCDATA1.GetDataTable(ASCMAIN1.sql)

            Addtask("WHC940O1 Fill ICTITEM1")
            ASCMAIN1.sql = "SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN (SELECT ITEM_CODE FROM SOTORDR2 WHERE ORDR_NO IN (SELECT ORDR_NO from " & tempTableSOTPICKX & "))"
            tblICTITEM1 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTITEM1")

            Addtask("WHC940O1 Fill TATTERM1")
            ASCMAIN1.sql = "Select * from TATTERM1"
            tblTATTERM1 = ASCDATA1.GetDataTable(ASCMAIN1.sql, "TATTERM1")

            Addtask("WHC940O1 Fill EDTSLSP1")
            tblEDTSLSP1 = ASCDATA1.GetDataTable("Select * from EDTSLSP1")
            tblEDTSLSP1.PrimaryKey = New DataColumn() {tblEDTSLSP1.Columns("CUST_CODE")}

            ' Need to Lock sales orders and Shipment Records
            Dim unableToLock As String = String.Empty

            Addtask("WHC940O1 Start Locking Shipments")
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTSHIP1"), New String() {"SHIP_BOL_NO", "ORDR_GROUP_NO"}).Select("", "SHIP_BOL_NO")
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & String.Empty
                Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & String.Empty

                Addtask($"WHC940O1 Lock Shipment {SHIP_BOL_NO}")
                If Not ASCMAIN1.Logical_Lock("SOTSHIP1", SHIP_BOL_NO) Then
                    unableToLock = "Unable to Lock Ship Bill Of Lading: " & SHIP_BOL_NO
                    clsErrorMessage.Add(unableToLock)
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then
                        badOrderGroupNos.Add(ORDR_GROUP_NO)
                    End If
                Else
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then
                        If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO) Then
                            unableToLock = "Unable to Lock Order Group: " & ORDR_GROUP_NO
                            clsErrorMessage.Add(unableToLock)
                            If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then
                                badOrderGroupNos.Add(ORDR_GROUP_NO)
                            End If
                        End If
                    End If
                End If
            Next
            Addtask("WHC940O1 Finish Locking Shipments")

            ' validate the data to send to the Warehouse
            Addtask("WHC940O1 Start Validate the data to send to the Warehouse")
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("", "PICK_NO")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO") & String.Empty
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                Dim SHIP_BOL_NO As String = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty

                Addtask($"WHC940O1 Start Pick Ticket Header {PICK_NO}")

                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)

                Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty
                Dim ORDR_GROUP_NO As String = rowSOTORDR1.Item("ORDR_GROUP_NO") & String.Empty

                If dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                    clsErrorMessage.Add($"Shipment {SHIP_BOL_NO}, Order No {ORDR_NO}, Pick Ticket {PICK_NO} does not have a Ship To record in SOTORDR5")
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    Continue For
                End If
                Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "'")(0)

                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

                Dim XREF_CUST_CODE_SOLD_TO As String = String.Empty
                Dim XREF_CUST_STORE_NO_SOLD_TO As String = String.Empty

                Dim XREF_CUST_CODE_SHIP_TO As String = String.Empty
                Dim XREF_CUST_STORE_NO_SHIP_TO As String = String.Empty

                If dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "' AND PICK_QTY > 0").Length = 0 Then
                    clsErrorMessage.Add("No items to ship on Pick Ticket: " & PICK_NO)
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                End If

                If rowARTCUST1.Item("CUST_SHIP_TO_MANUAL") & String.Empty = "1" Then
                    XREF_CUST_CODE_SOLD_TO = rowARTCUST1.Item("CUST_NO_3PL") & String.Empty
                    XREF_CUST_STORE_NO_SOLD_TO = rowARTCUST1.Item("CUST_STORE_NO_3PL") & String.Empty

                    XREF_CUST_CODE_SHIP_TO = XREF_CUST_CODE_SOLD_TO
                    XREF_CUST_STORE_NO_SHIP_TO = XREF_CUST_STORE_NO_SOLD_TO
                Else
                    XREF_CUST_CODE_SOLD_TO = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty ' rowARTCUST1.Item("CUST_NO_3PL") & String.Empty
                    XREF_CUST_STORE_NO_SOLD_TO = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty ' "0"

                    XREF_CUST_CODE_SHIP_TO = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                    XREF_CUST_STORE_NO_SHIP_TO = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty

                    Select Case rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & String.Empty
                        Case "DC"
                            Dim CUST_DC_NO As String = rowARTCUST2.Item("CUST_DC_NO") & String.Empty
                            rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_DC_NO})

                            ' Set to Blank so if the DC record is missing the Clarins Settings
                            ' and error is generated
                            XREF_CUST_CODE_SHIP_TO = String.Empty
                            XREF_CUST_STORE_NO_SHIP_TO = String.Empty

                            If rowARTCUST2 IsNot Nothing Then
                                XREF_CUST_CODE_SHIP_TO = rowARTCUST2.Item("CUST_NO_3PL") & String.Empty
                                XREF_CUST_STORE_NO_SHIP_TO = rowARTCUST2.Item("CUST_STORE_NO_3PL") & String.Empty
                            End If
                    End Select
                End If

                XREF_CUST_CODE_SOLD_TO = XREF_CUST_CODE_SOLD_TO.Trim
                XREF_CUST_STORE_NO_SOLD_TO = XREF_CUST_STORE_NO_SOLD_TO.Trim
                If XREF_CUST_CODE_SOLD_TO.Length = 0 OrElse XREF_CUST_STORE_NO_SOLD_TO.Length = 0 Then
                    clsErrorMessage.Add("Missing Customer Code or Store No Cross Reference for Customer/Sold To (" & CUST_CODE & "/" & CUST_STORE_NO & ") on Pick Ticket: " & PICK_NO)
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    'Exit Sub
                End If

                XREF_CUST_CODE_SHIP_TO = XREF_CUST_CODE_SHIP_TO.Trim
                XREF_CUST_STORE_NO_SHIP_TO = XREF_CUST_STORE_NO_SHIP_TO.Trim
                If XREF_CUST_CODE_SHIP_TO.Length = 0 OrElse XREF_CUST_STORE_NO_SHIP_TO.Length = 0 Then
                    clsErrorMessage.Add("Missing Bill To Cross Reference for Customer/Ship To (" & CUST_CODE & "/" & CUST_STORE_NO & ") on Pick Ticket: " & PICK_NO)
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    'Exit Sub
                End If

                ' Validate Only EDI orders
                If arrivalDateCustomers.Contains(CUST_CODE) AndAlso rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty = "E" Then
                    ' Reverse POs do not use arrival date 
                    If Not IsDate(rowSOTORDR1.Item("ORDR_ARRIVAL_DATE") & String.Empty) AndAlso rowSOTORDR1.Item("REVERSE_PO") & String.Empty <> "1" Then
                        clsErrorMessage.Add("Missing Order Arrival Date on Sales Order: " & ORDR_NO)
                        If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    End If
                End If

                ' Must have complete address
                If rowARTCUST1.Item("CUST_SHIP_TO_MANUAL") & String.Empty = "1" OrElse rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & String.Empty = "MA" Then
                    For Each fieldname As String In New String() {"CUST_NAME", "CUST_ADDR1", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE"}
                        If (rowSOTORDR5.Item(fieldname) & String.Empty).ToString.Trim.Length = 0 Then
                            clsErrorMessage.Add("Invalid Ship To " & StrConv(fieldname.Replace("_", " "), VbStrConv.ProperCase) & " on Order Number: " & ORDR_NO)
                            If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                        End If
                    Next
                End If

                ' If Freight Terms equal PPA then the value in CUST_FRT_CHG_CODE must be D U or N
                Select Case rowSOTORDR1.Item("FRT_TERMS") & String.Empty
                    Case "PPA"
                        If Not "DUN".Contains(rowARTCUST1.Item("CUST_FRT_CHG_CODE") & String.Empty) Then
                            clsErrorMessage.Add("PPA freight terms requires the Customer Master have a valid Freight Charge Code. Sales Order: " & ORDR_NO)
                            If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                        End If
                End Select

                If CDate(rowSOTORDR1.Item("ORDR_CANCEL_DATE") & String.Empty) < CDate(DateTime.Now.ToShortDateString & String.Empty) Then
                    clsErrorMessage.Add("Cancel Date (" & rowSOTORDR1.Item("ORDR_CANCEL_DATE") & ") less than today's date on Order Number: " & ORDR_NO)
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                End If

                If rowSOTORDR1.Item("WHSE_CODE") & String.Empty <> "CLA" AndAlso rowSOTORDR1.Item("WHSE_CODE") & String.Empty <> "CLARTN" Then '
                    clsErrorMessage.Add("Invalid Warehouse (" & rowSOTORDR1.Item("WHSE_CODE") & ") on Order Number: " & ORDR_NO)
                    If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)

                End If

                rowSOTPICK1.Item("XREF_CUST_CODE_SOLD_TO") = XREF_CUST_CODE_SOLD_TO
                rowSOTPICK1.Item("XREF_CUST_STORE_NO_SOLD_TO") = XREF_CUST_STORE_NO_SOLD_TO
                rowSOTPICK1.Item("XREF_CUST_CODE_SHIP_TO") = XREF_CUST_CODE_SHIP_TO
                rowSOTPICK1.Item("XREF_CUST_STORE_NO_SHIP_TO") = XREF_CUST_STORE_NO_SHIP_TO

                Addtask($"WHC940O1 Start Pick Ticket Details {PICK_NO}")
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'", "PICK_LNO")

                    Dim ORDR_LNO As Int16 = rowSOTPICK2.Item("ORDR_LNO") & String.Empty

                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO)(0)
                    Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty
                    Dim rowICTITEM1 As DataRow = tblICTITEM1.Rows.Find(ITEM_CODE)

                    ' 06/08/2018 place vendor sku in the detail record
                    Dim EDI_DOC_SEQ_NO As String = rowSOTORDR2.Item("EDI_DOC_SEQ_NO") & String.Empty
                    Dim EDI_DTL_SEQ As Int16 = Val(rowSOTORDR2.Item("EDI_DTL_SEQ") & String.Empty)
                    Dim rowEDT850T2 As DataRow = Nothing
                    If EDI_DOC_SEQ_NO.Length > 0 AndAlso EDI_DTL_SEQ > 0 Then
                        rowEDT850T2 = dst.Tables("EDT850T2").Rows.Find(New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ})
                    End If

                    If rowICTITEM1 Is Nothing Then
                        clsErrorMessage.Add("Invalid Item (" & ITEM_CODE & ") on Order Number: " & ORDR_NO)
                        If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                        Continue For
                    End If

                    If rowICTITEM1.Item("HIDE_FROM_3PL") & String.Empty = "1" Then
                        clsErrorMessage.Add("Hide from 3PL Item Code for Item (" & ITEM_CODE & ") on Order Number: " & ORDR_NO)
                        If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    End If

                    Dim ITEM_ALT_SORT As String = rowICTITEM1.Item("ITEM_ALT_SORT") & String.Empty
                    ITEM_ALT_SORT = ITEM_ALT_SORT.Trim
                    If ITEM_ALT_SORT.Length = 0 Then
                        clsErrorMessage.Add("Missing Warehouse Item Code for Item (" & ITEM_CODE & ") on Order Number: " & ORDR_NO)
                        If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    End If

                    Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE") & String.Empty
                    COLLECTION_CODE = COLLECTION_CODE.Trim
                    If COLLECTION_CODE.Length = 0 Then
                        clsErrorMessage.Add("Missing Collection Code for Item (" & ITEM_CODE & ") on Order Number: " & ORDR_NO)
                        If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    End If

                    numRecords = tblICTCOLL1.Select("COLLECTION_CODE = '" & COLLECTION_CODE & "'").Length
                    Dim XREF_BRAND_CODE As String = String.Empty
                    Select Case numRecords
                        Case 0
                            clsErrorMessage.Add("No Cross Reference for Collection Code (" & COLLECTION_CODE & ")  for Item (" & ITEM_CODE & ") on Order Number: " & ORDR_NO)
                            If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                        Case 1
                            XREF_BRAND_CODE = tblICTCOLL1.Select("COLLECTION_CODE = '" & COLLECTION_CODE & "'")(0).Item("BRAND_CODE_3PL") & String.Empty
                            If tblICTCOLL1.Select("COLLECTION_CODE = '" & COLLECTION_CODE & "'")(0).Item("HIDE_FROM_3PL") & String.Empty = "1" Then
                                clsErrorMessage.Add("Hide From 3PL setting on Collection Code (" & COLLECTION_CODE & ")  for Item (" & ITEM_CODE & ") on Order Number: " & ORDR_NO)
                                If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                            End If
                        Case Else
                            clsErrorMessage.Add("Multiple Cross References for Collection Code (" & COLLECTION_CODE & ")  for Item (" & ITEM_CODE & ") on Order Number: " & ORDR_NO)
                            If Not badOrderGroupNos.Contains(ORDR_GROUP_NO) Then badOrderGroupNos.Add(ORDR_GROUP_NO)
                    End Select

                    rowSOTPICK2.Item("ITEM_CODE") = ITEM_CODE
                    rowSOTPICK2.Item("ITEM_ALT_SORT") = ITEM_ALT_SORT
                    rowSOTPICK2.Item("XREF_BRAND_CODE") = XREF_BRAND_CODE
                    rowSOTPICK2.Item("COLLECTION_CODE") = COLLECTION_CODE
                    If rowEDT850T2 IsNot Nothing Then
                        rowSOTPICK2.Item("EDI_SKU") = rowEDT850T2.Item("EDI_SKU") & String.Empty
                    End If
                    Addtask($"WHC940O1 End Pick Ticket Details {PICK_NO}")
                Next
                Addtask($"WHC940O1 Finish Pick Ticket Header {PICK_NO}")
            Next

            Addtask("Call WHC000O1.Update_Record in Main_Process")
            Update_Record()
            Addtask("WHC940O1 Finish WHC000O1.Update_Record in Main_Process")

            Addtask("WHC940O1 Finish Main_Process")

        Catch ex As Exception
            ErrorMessages.Add(ex.Message)
        End Try

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Try
            Addtask("WHC940O1 Enter Update_Create_File")

            ' Need to send updated Item Master Data
            Addtask("WHC940O1 Start CreateItemMasterExtract")
            If Not CreateItemMasterExtract() Then
                ' Exit Sub
            End If
            Addtask("WHC940O1 Finish CreateItemMasterExtract")

            ' Need to send Updated Customer Store Data
            Addtask("WHC940O1 Start CreateCustomerStoreExtract")
            If Not CreateCustomerStoreExtract() Then
                ' Exit Sub
            End If
            Addtask("WHC940O1 Finish CreateCustomerStoreExtract")

            ' 11/4/2015 - as per meeting at Clarins. Separate files by order group
            ' If any pick ticket fails the entire file is rejected
            Dim tblGroup As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTSHIP1"), "ORDR_GROUP_NO")

            Addtask("WHC940O1 Start Creating Order File")
            For Each rowGroup As DataRow In tblGroup.Select("", "ORDR_GROUP_NO")
                System.Threading.Thread.Sleep(1000)

                Dim ORDR_GROUP_NO As String = rowGroup.Item("ORDR_GROUP_NO") & String.Empty

                If badOrderGroupNos.Contains(ORDR_GROUP_NO) Then
                    Continue For
                End If

                Dim cTime As String = DateTime.Now.ToString("yyyyMMddHHmmss")
                Dim rowARTCUST1 As DataRow = Nothing
                SOHDR = "SOHDR" & "_" & cTime & ".CSV"
                SODTL = "SODTL" & "_" & cTime & ".CSV"

                Addtask($"WHC940O1 Creating Order File {SOHDR}")

                ' 01/16/2025 - Split shipments across multiple wareshouse = Amazon
                'Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)
                'If rowSOTORDR0 IsNot Nothing AndAlso rowSOTORDR0.Item("WHSE_CODE") & String.Empty <> "CLA" Then
                '    SOHDR = "SCHDR" & "_" & cTime & ".CSV"
                '    SODTL = "SCDTL" & "_" & cTime & ".CSV"
                'End If

                If dst.Tables("SOTSHIP1").Select($"ORDR_GROUP_NO = '{ORDR_GROUP_NO}'", "")(0).Item("WHSE_CODE") & String.Empty <> "CLA" Then
                    SOHDR = "SCHDR" & "_" & cTime & ".CSV"
                    SODTL = "SCDTL" & "_" & cTime & ".CSV"
                End If

                salesOrderFiles.Add(SOHDR)
                salesOrderFiles.Add(SODTL)

                Using swHeader As New System.IO.StreamWriter(wkDirectory & SOHDR)
                    Using swDetail As New System.IO.StreamWriter(wkDirectory & SODTL)

                        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", "")
                            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO") & String.Empty
                            Dim FRT_TERMS As String = rowSOTSHIP1.Item("FRT_TERMS") & String.Empty
                            Dim ORDR_DEPT As String = rowSOTSHIP1.Item("ORDR_DEPT") & String.Empty
                            Dim SHIP_VIA_CODE As String = rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty

                            If Not listOfShipments.Contains(SHIP_BOL_NO) Then
                                listOfShipments.Add(SHIP_BOL_NO)
                            End If

                            SHIP_VIA_CODE = SHIP_VIA_CODE.PadRight(4, " ").Substring(0, 4).Trim
                            FRT_TERMS = FRT_TERMS.PadRight(3, " ").Substring(0, 3).Trim
                            ORDR_DEPT = ORDR_DEPT.PadRight(20, " ").Substring(0, 20).Trim

                            If FRT_TERMS = "DEL" Then
                                FRT_TERMS = "PPD"
                            End If

                            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "PICK_NO")
                                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO") & String.Empty
                                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                                Dim EDI_STORE_NO As String = String.Empty
                                Dim EDI_DC_NO As String = String.Empty
                                Dim NUMBER_CHARS_STORE As Int16 = 0
                                Dim NUMBER_CHARS_DC As Int16 = 0

                                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)
                                Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")(0)

                                Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE") & String.Empty
                                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty
                                Dim CUST_DC_NO As String = rowSOTORDR1.Item("CUST_DC_NO") & String.Empty
                                Dim ORDR_SHIP_DATE As String = rowSOTORDR1.Item("ORDR_SHIP_DATE") & String.Empty
                                Dim ORDR_CUST_PO As String = rowSOTORDR1.Item("ORDR_CUST_PO") & String.Empty
                                Dim ORDR_TYPE_CODE As String = rowSOTORDR1.Item("ORDR_TYPE_CODE") & String.Empty
                                Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE") & String.Empty
                                Dim EDI_DOC_SEQ_NO As String = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & String.Empty

                                ORDR_CUST_PO = ORDR_CUST_PO.Replace(quo, "'")

                                If rowARTCUST1 Is Nothing OrElse rowARTCUST1.Item("CUST_CODE") & String.Empty <> CUST_CODE Then
                                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                                End If

                                ' Reverse POs do not use arrival date
                                If arrivalDateCustomers.Contains(CUST_CODE) AndAlso rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty = "E" AndAlso IsDate(rowSOTORDR1.Item("ORDR_ARRIVAL_DATE") & String.Empty) Then
                                    ORDR_SHIP_DATE = rowSOTORDR1.Item("ORDR_ARRIVAL_DATE") & String.Empty
                                End If

                                Dim PICK_TICKET_MESSAGE As String = (rowSOTORDR1.Item("ORDR_SPECIAL_INST") & String.Empty).ToString.Replace(quo, "'").Trim.Replace(vbCrLf, " ")
                                Dim INV_MESSAGE As String = (rowSOTORDR1.Item("ORDR_INV_COMMENT") & String.Empty).ToString.Replace(quo, "'").Trim.Replace(vbCrLf, " ")
                                Dim PACK_SLIP_MESSAGE As String = (rowSOTORDR1.Item("ORDR_MESSAGE") & String.Empty).ToString.Replace(quo, "'").Trim.Replace(vbCrLf, " ")

                                While PICK_TICKET_MESSAGE.Contains(Space(2))
                                    PICK_TICKET_MESSAGE = PICK_TICKET_MESSAGE.Replace(Space(2), Space(1))
                                End While

                                While INV_MESSAGE.Contains(Space(2))
                                    INV_MESSAGE = INV_MESSAGE.Replace(Space(2), Space(1))
                                End While

                                While PACK_SLIP_MESSAGE.Contains(Space(2))
                                    PACK_SLIP_MESSAGE = PACK_SLIP_MESSAGE.Replace(Space(2), Space(1))
                                End While

                                Dim CUST_COUNTRY As String = (rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty).ToString.Trim.ToUpper
                                Dim CUST_ZIP_CODE_LAST4 As String = (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.Trim
                                If CUST_ZIP_CODE_LAST4.Length > 5 AndAlso CUST_COUNTRY.StartsWith("US") Then
                                    CUST_ZIP_CODE_LAST4 = CUST_ZIP_CODE_LAST4.Substring(5)
                                    CUST_ZIP_CODE_LAST4 = CUST_ZIP_CODE_LAST4.Replace("-", "").Trim
                                    If CUST_ZIP_CODE_LAST4.Length <> 4 Then
                                        CUST_ZIP_CODE_LAST4 = String.Empty
                                    End If
                                ElseIf CUST_COUNTRY = "CAN" Then
                                    ' As per Bob. for Canada send first 3 chars in zip code and last 3 in zip4
                                    CUST_ZIP_CODE_LAST4 = Microsoft.VisualBasic.Mid(rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty, 4).ToString.Trim
                                Else
                                    CUST_ZIP_CODE_LAST4 = String.Empty
                                End If

                                Dim CUST_NAME As String = (rowSOTORDR5.Item("CUST_NAME") & String.Empty).ToString.Replace(quo, " ").Trim
                                If CUST_NAME.Length > 40 Then
                                    CUST_NAME = CUST_NAME.Substring(0, 40).Trim
                                End If

                                Addtask($"WHC940O1 Write Header Order File {SOHDR}")
                                With swHeader
                                    'ABS_Order_Number	Char	10				Header
                                    .Write(quo & ORDR_NO & quo)
                                    'Order_Date		Decimal	6,0	yymmdd			Header
                                    .Write(sep & CDate(rowSOTORDR1.Item("ORDR_DATE") & String.Empty).ToString("yyMMdd"))
                                    'Request_Ship_Date	Decimal	6,0	yymmdd			Header
                                    .Write(sep & CDate(ORDR_SHIP_DATE).ToString("yyMMdd"))
                                    'Customer_Number_1	Decimal	7,0	Clarins			Header
                                    .Write(sep & rowSOTPICK1.Item("XREF_CUST_CODE_SOLD_TO"))
                                    'Customer_Number_2	Decimal	3,0	Clarins			Header
                                    .Write(sep & rowSOTPICK1.Item("XREF_CUST_STORE_NO_SOLD_TO"))
                                    'PO_Number		Char	20				Header
                                    .Write(sep & quo & ORDR_CUST_PO.PadRight(20, " ").Substring(0, 20).Trim & quo)
                                    'Carrier_Code		Char	3				Header
                                    .Write(sep & quo & SHIP_VIA_CODE & quo)
                                    'Invoice_Number		Char	10	ABS			Header
                                    .Write(sep & quo & Space(10) & quo)
                                    'Merch Dollars		Decimal	9,2				Header
                                    .Write(sep & rowSOTPICK1.Item("TOT_EXT_PICK_UNIT_PRICE") & String.Empty)

                                    'Ship_To_Cust1		Decimal	7,0	Clarins			Header
                                    .Write(sep & rowSOTPICK1.Item("XREF_CUST_CODE_SHIP_TO"))
                                    'Ship_To_Cust2		Decimal	3,0	Clarins			Header
                                    .Write(sep & rowSOTPICK1.Item("XREF_CUST_STORE_NO_SHIP_TO"))
                                    'Ship_To_Address1	Char	30				Header
                                    .Write(sep & quo & (rowSOTORDR5.Item("CUST_ADDR1") & String.Empty).ToString.Replace(vbCrLf, " ").PadRight(30, " ").Replace(quo, " ").Substring(0, 30).Trim & quo)
                                    'Ship_To_Address2	Char	30				Header
                                    .Write(sep & quo & (rowSOTORDR5.Item("CUST_ADDR2") & String.Empty).ToString.Replace(vbCrLf, " ").PadRight(30, " ").Replace(quo, " ").Substring(0, 30).Trim & quo)
                                    'Ship_To_City		Char	25				Header
                                    .Write(sep & quo & (rowSOTORDR5.Item("CUST_CITY") & String.Empty).ToString.Replace(vbCrLf, " ").PadRight(25, " ").Replace(quo, " ").Substring(0, 25).Trim & quo)
                                    'Ship_To_State		Char	2				Header
                                    .Write(sep & quo & (rowSOTORDR5.Item("CUST_STATE") & String.Empty).ToString.Replace(vbCrLf, " ").PadRight(2, " ").Replace(quo, " ").Substring(0, 2).Trim & quo)

                                    'Ship_To_Zip		Char	10				Header
                                    ' As per Bob. for Canada send first 3 chars in zip code and last 3 in zip4
                                    If CUST_COUNTRY = "CAN" Then
                                        .Write(sep & quo & Microsoft.VisualBasic.Mid(rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty, 1, 3).PadRight(10, " ").Replace(quo, " ").Substring(0, 10).Trim & quo)
                                    Else
                                        .Write(sep & quo & (rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty).ToString.PadRight(10, " ").Replace(quo, " ").Substring(0, 10).Trim & quo)
                                    End If

                                    'Cancel_Date		Decimal	6,0	yymmdd		Header
                                    If arrivalDateCustomers.Contains(CUST_CODE) AndAlso CUST_CODE = "ULTA" _
                                        AndAlso CDate(ORDR_SHIP_DATE) > CDate(rowSOTORDR1.Item("ORDR_CANCEL_DATE") & String.Empty) Then
                                        .Write(sep & DateAdd(DateInterval.Day, 1, CDate(ORDR_SHIP_DATE)).ToString("yyMMdd"))
                                    Else
                                        .Write(sep & CDate(rowSOTORDR1.Item("ORDR_CANCEL_DATE") & String.Empty).ToString("yyMMdd"))
                                    End If

                                    'Clarins_Status		Char	1	Blank. Clarins Will Update	Header
                                    .Write(sep & quo & Space(1) & quo)
                                    'Clarins_Order_Number	Char	7	Blank. Clarins Will Update	Header
                                    .Write(sep & quo & Space(7) & quo)

                                    'Fatal_Error_Reason	Char	10	Blank. Clarins Will Update	Header
                                    .Write(sep & quo & Space(10) & quo)
                                    'Fatal_Error_Date	Decimal	6	0. Clarins Will Update	Header
                                    .Write(sep & "0")
                                    'Fatal_Error_Time	Decimal	6	0. Clarins Will Update	Header
                                    .Write(sep & "0")

                                    ' ---------------------------------------------------------------
                                    ' Data Needed by ABSolution
                                    ' ---------------------------------------------------------------

                                    'CUST_CODE		Char	10	ABS Customer Code	Header
                                    .Write(sep & quo & CUST_CODE.PadRight(10, " ").Substring(0, 10).Trim & quo)
                                    'CUST_STORE_NO		Char	6	ABS Store No		Header
                                    .Write(sep & quo & CUST_STORE_NO.PadRight(6, " ").Substring(0, 6).Trim & quo)
                                    'CUST_DC_NO		Char	6	ABS DC No		Header
                                    .Write(sep & quo & CUST_DC_NO.PadRight(6, " ").Substring(0, 6).Trim & quo)
                                    'PICK_NO			Char	10	ABS Pick Ticket No	Header
                                    .Write(sep & quo & PICK_NO & quo)
                                    'ORDR_GROUP_NO		Char	10	ABS Order Group No	Header
                                    .Write(sep & quo & ORDR_GROUP_NO & quo)
                                    'SHIP_BOL_NO		Char	10	ABS Ship Ctl No		Header
                                    .Write(sep & quo & SHIP_BOL_NO & quo)
                                    'FRT_TERMS		Char	3	PPD/COL/PPA		Header
                                    .Write(sep & quo & FRT_TERMS & quo)
                                    'ORDR_MISC_CHG
                                    .Write(sep & Val(rowSOTORDR1.Item("ORDR_MISC_CHG") & String.Empty))

                                    EDI_DC_NO = CUST_DC_NO
                                    EDI_STORE_NO = CUST_STORE_NO

                                    Dim rowEDTSLSP1 As DataRow = tblEDTSLSP1.Rows.Find(CUST_CODE)
                                    If rowEDTSLSP1 IsNot Nothing Then

                                        NUMBER_CHARS_STORE = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
                                        NUMBER_CHARS_DC = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & String.Empty)

                                        If NUMBER_CHARS_STORE < 1 Or NUMBER_CHARS_STORE > 6 Then
                                            NUMBER_CHARS_STORE = 4
                                        End If

                                        If NUMBER_CHARS_DC < 1 Or NUMBER_CHARS_DC > 6 Then
                                            NUMBER_CHARS_DC = 4
                                        End If

                                        If IsNumeric(EDI_STORE_NO) Then
                                            EDI_STORE_NO = StrReverse(StrReverse(EDI_STORE_NO.PadLeft(NUMBER_CHARS_STORE, "0")).Substring(0, NUMBER_CHARS_STORE))
                                        Else
                                            EDI_STORE_NO = EDI_STORE_NO.PadRight(NUMBER_CHARS_STORE, " ").Substring(0, NUMBER_CHARS_STORE)
                                        End If

                                        If IsNumeric(EDI_DC_NO) Then
                                            EDI_DC_NO = StrReverse(StrReverse(EDI_DC_NO.PadLeft(NUMBER_CHARS_DC, "0")).Substring(0, NUMBER_CHARS_DC))
                                        Else
                                            EDI_DC_NO = EDI_DC_NO.PadRight(NUMBER_CHARS_DC, " ").Substring(0, NUMBER_CHARS_DC)
                                        End If
                                    End If

                                    EDI_STORE_NO = EDI_STORE_NO.PadRight(6, " ").Substring(0, 6).Trim
                                    EDI_DC_NO = EDI_DC_NO.PadRight(6, " ").Substring(0, 6).Trim

                                    'EDI_STORE_NO		Char	6	Store No in EDI Format	Header
                                    .Write(sep & quo & EDI_STORE_NO & quo)
                                    'EDI_DC_NO		Char	6	DC No in EDI Format	Header
                                    .Write(sep & quo & EDI_DC_NO & quo)
                                    'ORDR_DEPT		Char	20	Department No		Header
                                    .Write(sep & quo & ORDR_DEPT & quo)
                                    ' Clarins Batch ID
                                    .Write(sep)
                                    ' Clarins Combine
                                    .Write(sep)

                                    'PICK_TICKET_MESSAGE		Char	250 Pick Ticket Message		Header
                                    .Write(sep & quo & PICK_TICKET_MESSAGE & quo)

                                    'INV_MESSAGE		Char	120     Invoice Message		Header
                                    ' Removed on 11/4/2015
                                    '.Write(sep & quo & INV_MESSAGE & quo)

                                    'PACK_SLIP_MESSAGE		Char	500 Pick Slip Message		Header
                                    .Write(sep & quo & PACK_SLIP_MESSAGE & quo)

                                    'ORDR_FOB		Char	40	FOB		Header
                                    .Write(sep & quo & (rowSOTORDR1.Item("EDI_MERCH_TYPE") & String.Empty).ToString.Replace(quo, "'") & quo)
                                    'CUST_VEND_REF		Char	20 Vendor		Header
                                    .Write(sep & quo & rowSOTORDR1.Item("CUST_VEND_REF") & String.Empty & quo)

                                    'SHPNAME	951	40	0	0	Ship_To_Name
                                    .Write(sep & quo & CUST_NAME & quo)
                                    'SHPZIP4	991	4	0	0	Ship_To_Zip4
                                    .Write(sep & quo & CUST_ZIP_CODE_LAST4 & quo)

                                    'Use_Shipping_Address
                                    If rowARTCUST1.Item("CUST_SHIP_TO_MANUAL") & String.Empty = "1" _
                                        OrElse rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & String.Empty = "MA" Then
                                        .Write(sep & quo & "Y" & quo)
                                    Else
                                        .Write(sep & quo & "N" & quo)
                                    End If

                                    'Clarins Freight Code 
                                    If rowARTCUST1.Item("CUST_FRT_CHG_CODE") & String.Empty = "Z" Then
                                        .Write(sep & quo & String.Empty & quo)
                                    Else
                                        .Write(sep & quo & rowARTCUST1.Item("CUST_FRT_CHG_CODE") & String.Empty & quo)
                                    End If

                                    ' New field 01/07/2016
                                    ' We need to add a new field to the order header for order type.
                                    ' 3 byte character. 
                                    ' Not being able to determine a wholesale order from a basic order is causing us problems.
                                    ' “W  ”      wholesale
                                    ' “B  ”       Basic
                                    ' "WP " - "NORDRACK", "MACYSBACK"
                                    ' As per Bob on 1/14/2016
                                    ' For future Nordstrom Rack orders, we require an order type of WP.
                                    ' Macy Backstage will also require an order type of WP.

                                    ' As per Bob on 1/20/2016
                                    ' in addition to the NORDRACK and MACYSBACK changes, please make the changes Bob is suggesting below.
                                    ' If ORDR_TYPE_CODE = ‘RTV’ then use IT (if coming from CLA) and Y (if coming from CLARTN).

                                    ' As per Bob on 09/11/2017
                                    ' Clarins Order type “Y” should not care about order multiples, minimums or maximums.
                                    ' so we need to put code in Order Release and Order entry to avoid these constraints
                                    ' but we beed to do that using values in SOTORDR1, which could be used to anticipdate
                                    ' clarins order type "Y"
                                    ' currently, 2 customer KTS and PERFUMECTR are coded with ORDR_CODE_3PL = 'Y'
                                    ' and RTV orders (except those out of whse CLA) are order type Y
                                    ' if the code below changes, then the code placed into SOE and SOR might also need review

                                    If ORDR_TYPE_CODE = "RTV" Then
                                        Select Case WHSE_CODE
                                            Case "CLA"
                                                .Write(sep & quo & "IT " & quo)
                                            Case "CLARTN"
                                                .Write(sep & quo & "Y  " & quo)
                                            Case Else
                                                ' Just a safety net
                                                .Write(sep & quo & "Y  " & quo)
                                        End Select
                                    ElseIf rowARTCUST1.Item("ORDR_CODE_3PL") & String.Empty <> String.Empty Then
                                        .Write(sep & quo & (rowARTCUST1.Item("ORDR_CODE_3PL") & String.Empty).ToString.Trim.PadRight(3, " ") & quo)
                                    ElseIf tblWholesale.Select("CUST_CODE = '" & CUST_CODE & "'").Length = 0 Then
                                        .Write(sep & quo & "B  " & quo)
                                    Else
                                        .Write(sep & quo & "W  " & quo)
                                    End If

                                    ' Added 01/11/2016
                                    Dim rowEDT850T1 As DataRow = dst.Tables("EDT850T1").Rows.Find(EDI_DOC_SEQ_NO)
                                    If rowEDT850T1 Is Nothing Then
                                        .Write(sep & quo & "  " & quo)
                                    Else
                                        .Write(sep & quo & (rowEDT850T1.Item("EDI_PO_TYPE") & String.Empty).ToString.PadLeft(2, " ") & quo)
                                    End If

                                    .WriteLine()
                                End With

                                Addtask($"WHC940O1 Write Detail Order File {SODTL}")
                                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "' AND PICK_QTY > 0", "PICK_LNO")
                                    With swDetail
                                        'ABS_Order_Number	Char	10		Detail
                                        .Write(quo & ORDR_NO & quo)
                                        'Line_Number	Decimal	4,0		Detail
                                        .Write(sep & Val(rowSOTPICK2.Item("ORDR_LNO") & String.Empty))
                                        'Item_Number	Char	6		Detail
                                        .Write(sep & quo & (rowSOTPICK2.Item("ITEM_ALT_SORT") & String.Empty).ToString.PadRight(6, " ").Substring(0, 6).Trim & quo)
                                        'Qty_Ordered	Decimal	9,0		Detail
                                        .Write(sep & Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty))
                                        'Price	Decimal	9,2		Detail
                                        .Write(sep & Val(rowSOTPICK2.Item("PICK_UNIT_PRICE") & String.Empty))
                                        'Brand	Decimal	3,0		Detail
                                        .Write(sep & Val(rowSOTPICK2.Item("XREF_BRAND_CODE") & String.Empty))
                                        'ABS_Item_Number	Char	25		Detail
                                        .Write(sep & quo & (rowSOTPICK2.Item("ITEM_CODE") & String.Empty).ToString.PadRight(25, " ").Substring(0, 25).Trim & quo)
                                        ' Clarins Batch ID
                                        .Write(sep)
                                        ' Clarins Combine
                                        .Write(sep)
                                        ' ERRREASON  A     10  
                                        .Write(sep)
                                        ' CUSTSKU    A     27    
                                        .Write(sep & quo & (rowSOTPICK2.Item("EDI_SKU") & String.Empty).ToString.PadRight(25, " ").Substring(0, 25).Trim & quo)

                                        .WriteLine()
                                    End With
                                Next
                            Next
                        Next
                        swDetail.Close()
                        swDetail.Dispose()
                    End Using
                    swHeader.Close()
                    swHeader.Dispose()
                End Using
            Next

            ' Global Variable that determine the number of records processed.
            ' Needed so Update_Archive is called from base class.
            R = salesOrderFiles.Count
        Catch ex As Exception
            ErrorMessages.Add(ex.Message)
        Finally
            Addtask("WHC940O1 Finish Update_Create_File")
        End Try

    End Sub

    Overrides Sub Update_Archive() 'ORDR_GIFT_MESSAGE'
        MyBase.Update_Archive()

        Dim processedShipments As String = "ShipmentsSentTo3pl:"

        Addtask("WHC940O1 Enter Update_Archive")

        If listOfShipments.Count = 0 Then
            clsErrorMessage.Add(processedShipments)
            Exit Sub
        End If

        ' Need to Update Tables before 
        Dim LP_XNO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.LP_XNO")

        Dim shipmentsSentTo3pl As String = String.Empty
        If listOfShipments.Count = 1 Then
            shipmentsSentTo3pl = "'" & listOfShipments(0) & "'"
        Else
            shipmentsSentTo3pl = "'" & String.Join("', '", listOfShipments.ToArray) & "'"
        End If

        processedShipments &= shipmentsSentTo3pl
        clsErrorMessage.Add(processedShipments)

        Addtask("WHC940O1 Update_Archive - Update SOTSHIP1")
        ASCMAIN1.sql = "Update SOTSHIP1 Set LP_STATUS = '1', LP_XNO = '" & LP_XNO & "', LP_XMIT_DATE = SYSDATE"
        ASCMAIN1.sql &= ", SHIP_PICK_PRINTED = SYSDATE"
        ASCMAIN1.sql &= " where SHIP_BOL_NO in (" & shipmentsSentTo3pl & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Addtask("WHC940O1 Update_Archive - Update SOTPICK1")
        ASCMAIN1.sql = "Update SOTPICK1 Set PICK_PRINTED = SYSDATE"
        ASCMAIN1.sql &= ", PICK_PRINTED_OPER = '" & ASCMAIN1.USER_ID & "'"
        ASCMAIN1.sql &= " where SHIP_BOL_NO in (" & shipmentsSentTo3pl & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        dst.Tables("WHTSHIPX").Rows.Clear()
        dst.Tables("WHTLPXN1").Rows.Clear()

        For Each shipBol As String In shipmentsSentTo3pl.Split(",")
            dst.Tables("WHTSHIPX").Rows.Add(New Object() {LP_XNO, shipBol.Replace("'", "").Trim})
        Next

        Addtask("WHC940O1 Update_Archive - Update WHTSHIPX")
        Update_Record_TDA("WHTSHIPX")

        Addtask("WHC940O1 Update_Archive - Update WHTLPXN1")
        Dim rowWHTLPXN1 As DataRow = dst.Tables("WHTLPXN1").NewRow
        rowWHTLPXN1.Item("LP_XNO") = LP_XNO
        rowWHTLPXN1.Item("LP_XNO_SOURCE") = MENU_ITEM_OBJECT
        rowWHTLPXN1.Item("LP_XNO_RECORDS") = dst.Tables("WHTSHIPX").Rows.Count
        rowWHTLPXN1.Item("LP_XNO_NOTES") = ""
        rowWHTLPXN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        rowWHTLPXN1.Item("INIT_DATE") = DateTime.Now
        dst.Tables("WHTLPXN1").Rows.Add(rowWHTLPXN1)
        Update_Record_TDA("WHTLPXN1")

        ' If the Archive directory does not exist then create it as a courtesy
        If Not My.Computer.FileSystem.DirectoryExists(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir) Then
            Addtask($"WHC940O1 Update_Archive - Create Archive Directory: {ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir}")
            My.Computer.FileSystem.CreateDirectory(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir)
            Addtask($"WHC940O1 Update_Archive - Create Archive Directory Completed")
        End If

        For Each salesOrderFilename As String In salesOrderFiles
            Addtask($"WHC940O1 Update_Archive - Start Move File {wkDirectory & salesOrderFilename} to {ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & salesOrderFilename}")
            My.Computer.FileSystem.MoveFile(wkDirectory & salesOrderFilename, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & salesOrderFilename)
            Addtask($"WHC940O1 Update_Archive - Finish File Move")
        Next

        Addtask("WHC940O1 Exit Update_Archive")

    End Sub

    Sub Create_Work_Table()

    End Sub

    Overrides Sub Post_Update_Archive()

        ' Added 09/26/2025 to prevent sending test environment data
        If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                MessageBox.Show("You are not in production. File transfer avoided.", "Update Archive", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        MyBase.Post_Update_Archive()

        Dim X As Date = Now

        Addtask("WHC940O1 Enter Post_Update_Archive")

        Dim filesUploaded As Boolean = True
        For Each salesOrderFilename As String In salesOrderFiles
            Try
                ' If a successful Oracle Update then Push the files to the sftp server
                Addtask($"WHC940O1 Start CopyFile {ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & salesOrderFilename} {sftp_folder & salesOrderFilename}")
                My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & salesOrderFilename, sftp_folder & salesOrderFilename)
                Addtask($"WHC940O1 End CopyFile {ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & salesOrderFilename} {sftp_folder & salesOrderFilename}")
            Catch ex As Exception
                ErrorMessages.Add("The shipment files were created and the shipment data was updated. However these files did not get uploaded" _
                                   & " to the sftp directory. (" & SOHDR & ", " & SODTL & ") : " & ex.Message)
                filesUploaded = False
            End Try
        Next

        Addtask("WHC940O1 Exit Post_Update_Archive")

        clsSuccessfulExecution = filesUploaded

    End Sub

    Private Function CreateItemMasterExtract() As Boolean
        Try
            Dim environ As New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "AC"
            environ.APP_KEY = ""
            environ.LP_CODE = "CLA"
            Dim clsWHC888O1 As New WHC888O1(environ)

            If clsWHC888O1.ErrorMessages.Count > 0 Then
                For Each EMsg As String In clsWHC888O1.ErrorMessages
                    EMsg = EMsg.Trim
                    If EMsg.Length > 0 Then
                        ErrorMessages.Add(EMsg)
                    End If
                Next
                Return False
            End If

            Return True

        Catch ex As Exception
            ErrorMessages.Add("Create Item Master Extract" & ex.Message)
            Return False
        Finally
        End Try

    End Function

    Private Function CreateCustomerStoreExtract() As Boolean

        Try
            Dim environ As New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "AC"
            environ.APP_KEY = ""
            Dim clsWHC999O1 As New WHC999O1(environ)

            If clsWHC999O1.ErrorMessages.Count > 0 Then
                For Each EMsg As String In clsWHC999O1.ErrorMessages
                    EMsg = EMsg.Trim
                    If EMsg.Length > 0 Then
                        ErrorMessages.Add(EMsg)
                    End If
                Next
                Return False
            End If

            Return True

        Catch ex As Exception
            ErrorMessages.Add("Create Customer Store Extract" & ex.Message)
            Return False
        Finally
        End Try

    End Function


End Class