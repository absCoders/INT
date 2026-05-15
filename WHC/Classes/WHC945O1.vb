Imports Oracle.ManagedDataAccess.Client

Public Class WHC945O1
    ' Create Anticipated Receipts File (Outbound) 943

    Inherits WHC000O1

    Private environ As ABSEnvironment
    Private clsWHCIMP01 As WHCIMPO1


    Private Enum ProcessState
        ProcessShipment = 1
        NoDetails = 2
        PickTicketNotInPick = 3
        MismatchDetailsAndTracking = 4
        InvalidOrderNumber = 5
        Converted = 6
        ProcessReturn = 7
        InvalidCustomer = 8
    End Enum

    Private Enum ProcessType
        Brands
        CancelledOrders
        CustomerMaster
        Initialize
        Inventory
        InventoryAdjustments
        Invoice
        ItemMaster
        NewCustomers
        OpenOrders
        OrderConfirmation
        Receipts
        Returns
        ShipmentConfirmation
        ShipRule
        ShipViaScac
        SubBrands
        Vendors
        DailyShip
    End Enum

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHC945O1"

        Create_Work_Table()

        With dst
            Create_TDA(.Tables.Add, "EDT945T1", "*")
            Create_TDA(.Tables.Add, "EDT945T2", "*")

            Create_TDA(.Tables.Add, "EDTRTRN1", "*")
            Create_TDA(.Tables.Add, "EDTRTRN2", "*")

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR5", "*")
            Create_TDA(.Tables.Add, "SOTPICK1", "*")
            Create_TDA(.Tables.Add, "SOTPICK2", "*", 1)
            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            Create_TDA(.Tables.Add, "SOTCART1", "*")

            Create_TDA(.Tables.Add, "SOTPICKC", "*")
            Create_TDA(.Tables.Add, "SOTPICKO", "*")

            Create_TDA(.Tables.Add, "SOTSHIP1", "*")

            Create_TDA(.Tables.Add, "EDTTRXN1", "*")
            Create_TDA(.Tables.Add, "EDTTRXNA", "*")

            CreateTableAdaptors(ProcessType.Initialize)

        End With

        Main_Process()

        DisposeOPD()
    End Sub

    Public Sub Main_Process()

        Try
            EnforceConstraints(False)

            ' Global Variable that determines the number of records processed.
            R = 0
            ASCMAIN1.sql = String.Empty

            Select Case G.APP_CMD & String.Empty

                Case "INT"
                    Try
                        ProcessClarinsBrands()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsBrands Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsSubBrands()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsSubBrands Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsInventory()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsInventory Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsInvoice()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsInvoice Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsShipRule()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsShipRule Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsNewCustomers()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsNewCustomers Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsItemMaster()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsItemMaster Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsShipViaScac()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsShipViaScac Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsOrderConfirmationFiles()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsOrderConfirmationFiles Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsCancellations()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsCancellations Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsShipmentConfirmationFiles()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsShipmentConfirmationFiles Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsOpenOrders()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsOpenOrders Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsReturnsTransactions()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsReturnsTransactions Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsReceipts()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsReceipts Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsVendors()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsVendors Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsCustomers()
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsCustomers Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsRetnsinvAbscnts()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsRetnsinv Error: " & ex.Message)
                    End Try

                    Try
                        ProcessClarinsDailyShip()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsDailyShip Error: " & ex.Message)
                    End Try

                Case "ADJ"
                    Try
                        ProcessClarinsAdjustments()
                        clsErrorMessage.Add(dst.Tables("EDTTRXN1").Rows.Count & " adjustments imported.")
                        Update_Record()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsAdjustments Error: " & ex.Message)
                    End Try

                Case "CASECOUNT"
                    Try
                        ProcessClarinsItemCaseCount()
                    Catch ex As Exception
                        clsErrorMessage.Add("ProcessClarinsItemCaseCount Error: " & ex.Message)
                    End Try

                Case Else
                    clsErrorMessage.Add("Missing Command Parameter")
                    Exit Sub
            End Select

            Update_Record()

        Catch ex As Exception
            ErrorMessages.Add(ex.Message)
        Finally
            ASCMAIN1.MultiTask_Release()
        End Try

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Try

        Catch ex As Exception
            ErrorMessages.Add(ex.Message)
        End Try

    End Sub

    Overrides Sub Update_Archive()
        MyBase.Update_Archive()

        ' Need to Update Tables before 
        ' Remove entries from the tables that have been converted to EDI 945 records.
        Dim invoiceList As New List(Of String)
        Dim returnsList As New List(Of String)
        Dim returnsOrderNoList As New List(Of String)
        Dim adjustmentsList As New List(Of String)

        If dst.Tables.Contains("SHIPHDR") Then
            For Each rowHeader As DataRow In dst.Tables("SHIPHDR").Select("PROCESS = " & ProcessState.Converted)
                Dim INVOICE_NUMBER As String = rowHeader.Item("OHINVN") & String.Empty
                If INVOICE_NUMBER.Length > 0 Then
                    invoiceList.Add(INVOICE_NUMBER)
                End If
            Next
        End If

        If dst.Tables.Contains("RTNHDR") Then
            For Each rowHeader As DataRow In dst.Tables("RTNHDR").Select("PROCESS = " & ProcessState.Converted)
                Dim INVOICE_NUMBER As String = rowHeader.Item("OHINVN") & String.Empty
                If INVOICE_NUMBER.Length > 0 Then
                    returnsList.Add(INVOICE_NUMBER)
                    returnsOrderNoList.Add(rowHeader.Item("OHORDN") & String.Empty)
                End If
            Next
        End If

        If invoiceList.Count > 0 Then

            Dim newList As New List(Of String)
            For ictr As Int16 = 0 To invoiceList.Count - 1
                newList.Add(invoiceList(ictr))

                If newList.Count = 500 Or ictr = invoiceList.Count - 1 Then
                    Dim sqlInvoices As String = "('" & String.Join("', '", newList) & "')"

                    ' Need to populate Clarins Invoice Number on Order Header
                    ASCMAIN1.sql = "Begin Declare Cursor C1 is Select * from CONV.CFG_SHIPHDR where OHINVN in " & sqlInvoices & ";" _
                        & " Begin for R1 in C1 loop " _
                        & "    Update SOTORDR1 SET PARTNER_ORDR_NO = R1.OHINVN WHERE ORDR_NO = (SELECT ORDR_NO FROM SOTPICK1 WHERE PICK_NO = R1.ABSPICKNBR); " _
                        & " End Loop; End; End;"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                    ASCMAIN1.sql = "UPDATE CONV.CFG_SHIPHDR SET PROCESS_IND = '1' where NVL(PROCESS_IND, '0') = '0' AND OHINVN in " & sqlInvoices
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                    ASCMAIN1.sql = "UPDATE CONV.CFG_SHIPDTL SET PROCESS_IND = '1' where NVL(PROCESS_IND, '0') = '0' AND SAINVN in " & sqlInvoices
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                    ASCMAIN1.sql = "UPDATE CONV.CFG_CARTON SET PROCESS_IND = '1' where NVL(PROCESS_IND, '0') = '0' AND CHINVN in " & sqlInvoices
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                    ASCMAIN1.sql = "UPDATE CONV.CFG_TRACK SET PROCESS_IND = '1' where NVL(PROCESS_IND, '0') = '0' AND INVN in " & sqlInvoices
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                    newList = New List(Of String)
                End If
            Next

            Update_Record_TDA("EDT945T1")
            Update_Record_TDA("EDT945T2")
            Update_Record_TDA("SOTORDR5")

            For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("")
                Dim EDI_BOL_NO As String = rowEDT945T1.Item("EDI_BOL_NO") & String.Empty
                Dim EDI_MASTER_BOL_NO As String = rowEDT945T1.Item("EDI_MASTER_BOL_NO") & String.Empty
                Dim SHIP_BOL_NO As String = rowEDT945T1.Item("EDI_SHIPMENT_ID") & String.Empty

                Dim SHIP_PICKUP_NO As String = rowEDT945T1.Item("SHIP_PICKUP_NO") & String.Empty
                Dim SHIP_AUTH_NO As String = rowEDT945T1.Item("SHIP_AUTH_NO") & String.Empty

                EDI_BOL_NO = EDI_BOL_NO.Trim
                EDI_MASTER_BOL_NO = EDI_MASTER_BOL_NO.Trim
                SHIP_PICKUP_NO = SHIP_PICKUP_NO.Trim
                SHIP_AUTH_NO = SHIP_AUTH_NO.Trim

                If EDI_BOL_NO.Length > 0 Then
                    ASCMAIN1.sql = "UPDATE SOTSHIP1 SET BILL_OF_LADING_NO = '" & EDI_BOL_NO & "' WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                End If

                If EDI_MASTER_BOL_NO.Length > 0 AndAlso EDI_MASTER_BOL_NO <> EDI_BOL_NO Then
                    ASCMAIN1.sql = "UPDATE SOTSHIP1 SET MASTER_SHIP_BOL_NO = '" & EDI_MASTER_BOL_NO & "' WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                End If

                If SHIP_PICKUP_NO.Length > 0 OrElse SHIP_AUTH_NO.Length > 0 Then
                    ASCMAIN1.sql = "UPDATE SOTSHIP1 SET SHIP_PICKUP_NO = '" & SHIP_PICKUP_NO.Replace("'", "") & "', SHIP_AUTH_NO = '" & SHIP_AUTH_NO & "' WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                End If
            Next

            dst.Tables("EDT945T1").Rows.Clear()
            dst.Tables("EDT945T2").Rows.Clear()
            dst.Tables("SOTORDR5").Rows.Clear()

            dst.Tables("SHIPHDR").Rows.Clear()
            dst.Tables("SHIPDTL").Rows.Clear()
            dst.Tables("CARTON").Rows.Clear()
            dst.Tables("TRACK").Rows.Clear()

        End If

        If returnsList.Count > 0 Then
            Dim sqlInvoices As String = "('" & String.Join("', '", returnsList) & "')"
            Dim sqlOrders As String = "('" & String.Join("', '", returnsOrderNoList) & "')"

            ASCMAIN1.sql = "Delete from CONV.CFG_RTNHDR where OHINVN in " & sqlInvoices
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Delete from CONV.CFG_RTNDTL where SAINVN in " & sqlInvoices
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Delete from CONV.CFG_RTNRA where ORDNUMBER in " & sqlOrders & " OR ORDNUMBER IS NULL"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            Update_Record_TDA("EDTRTRN1")
            Update_Record_TDA("EDTRTRN2")

            dst.Tables("EDTRTRN1").Rows.Clear()
            dst.Tables("EDTRTRN2").Rows.Clear()

            dst.Tables("RTNHDR").Rows.Clear()
            dst.Tables("RTNDTL").Rows.Clear()
            dst.Tables("RTNRA").Rows.Clear()
        End If

        ' Order Confirmation 
        If dst.Tables.Contains("SOTPICKC") Then
            If dst.Tables("SOTPICKC").Rows.Count > 0 Then

                Update_Record_TDA("SOTPICKC")

                Dim sql As String = String.Empty
                For Each rowSOTPICKC As DataRow In dst.Tables("SOTPICKC").Select("ISNULL(ERROR_REASON, '*') <> '*'")
                    Select Case rowSOTPICKC.Item("PICK_LNO")
                        Case "0"
                            sql = "UPDATE SOTPICK1 SET IMPORT_NO = '" & rowSOTPICKC.Item("IMPORT_NO") & "', " _
                                & " BATCH_ID = '" & rowSOTPICKC.Item("BATCH_ID") & "', " _
                                & " ORDR_NO_3PL = '" & rowSOTPICKC.Item("ORDR_NO_3PL") & "', " _
                                & " ERROR_REASON = '" & rowSOTPICKC.Item("ERROR_REASON") & "' " _
                                & " WHERE PICK_NO = '" & rowSOTPICKC.Item("PICK_NO") & "'"
                        Case Else
                            sql = "UPDATE SOTPICK2 SET ERROR_REASON = '" & rowSOTPICKC.Item("ERROR_REASON") & "'" _
                                 & " WHERE PICK_NO = '" & rowSOTPICKC.Item("PICK_NO") & "'" _
                                 & " AND PICK_LNO = " & rowSOTPICKC.Item("PICK_LNO")
                    End Select
                    ASCDATA1.ExecuteSQL(sql)
                Next

                ASCMAIN1.sql = "Delete from CONV.CFG_ABSOH"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                ASCMAIN1.sql = "Delete from CONV.CFG_ABSOD"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                dst.Tables("SOTPICKC").Rows.Clear()
            End If
        End If

        ' Open Orders - Only keep a current list
        If dst.Tables.Contains("SOTPICKO") Then
            If dst.Tables("SOTPICKO").Rows.Count > 0 Then
                ASCMAIN1.sql = "Delete from SOTPICKO"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                Update_Record_TDA("SOTPICKO")

                ASCMAIN1.sql = "Delete from CONV.CFG_OPNORDHED"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                ASCMAIN1.sql = "Delete from CONV.CFG_OPNORDDTL"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                ' These are pick tickets in Pick that Clarins did not provide in the Open Orders Data
                ASCMAIN1.sql = "INSERT INTO SOTPICKO SELECT 'C' || ROWNUM, PICK_NO, ORDR_NO, PICK_STATUS" _
                    & " FROM SOTPICK1, SOTSHIP1 " _
                    & " WHERE SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" _
                    & " AND SOTPICK1.PICK_STATUS = 'P'" _
                    & " AND SOTSHIP1.SHIP_STATUS = 'P'" _
                    & " AND SOTPICK1.PICK_NO NOT IN (SELECT PICK_NO FROM SOTPICKO where PICK_NO IS NOT NULL)" _
                    & " AND NVL(SOTSHIP1.LP_STATUS, '0') = '0' AND SOTSHIP1.LP_XNO IS NULL"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                'Grab the current status of the Pick Ticket
                ASCMAIN1.sql = "UPDATE SOTPICKO SET PICK_STATUS = (SELECT PICK_STATUS FROM SOTPICK1 WHERE PICK_NO = SOTPICKO.PICK_NO)"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                dst.Tables("SOTPICKO").Rows.Clear()
            End If
        End If

        ' Adjustments
        If dst.Tables.Contains("EDTTRXN1") Then
            If dst.Tables("EDTTRXN1").Rows.Count > 0 Then
                Update_Record_TDA("EDTTRXN1")
                dst.Tables("EDTTRXN1").Rows.Clear()

                For Each rowINVTRANS As DataRow In dst.Tables("INVTRANS").Select()
                    adjustmentsList.Add(rowINVTRANS.Item("PROC_KEY") & String.Empty)
                Next

            End If

            If dst.Tables("EDTTRXNA").Rows.Count > 0 Then
                Update_Record_TDA("EDTTRXNA")
                dst.Tables("EDTTRXNA").Rows.Clear()

                For Each rowINVTRANSA As DataRow In dst.Tables("INVTRANSA").Select()
                    adjustmentsList.Add(rowINVTRANSA.Item("PROC_KEY") & String.Empty)
                Next
            End If

            ' More than 1000 causes an Oracle Error
            If adjustmentsList.Count > 0 Then
                If adjustmentsList.Count <= 500 Then
                    Dim sqlAdjustments As String = "('" & String.Join("', '", adjustmentsList) & "')"
                    ASCDATA1.ExecuteSQL("Update CONV.CFG_INVTRANSADJ set PROCESSED_DATE = SYSDATE, PROCESSED_IND = '1' WHERE PROC_KEY IN " & sqlAdjustments)
                Else
                    Dim lstadj As New List(Of String)
                    For Each adj As String In adjustmentsList
                        lstadj.Add(adj)

                        If lstadj.Count >= 500 Then
                            Dim sqlAdjustments As String = "('" & String.Join("', '", lstadj) & "')"
                            ASCDATA1.ExecuteSQL("Update CONV.CFG_INVTRANSADJ set PROCESSED_DATE = SYSDATE, PROCESSED_IND = '1' WHERE PROC_KEY IN " & sqlAdjustments)
                            lstadj.Clear()
                        End If
                    Next

                    If lstadj.Count > 0 Then
                        Dim sqlAdjustments As String = "('" & String.Join("', '", lstadj) & "')"
                        ASCDATA1.ExecuteSQL("Update CONV.CFG_INVTRANSADJ set PROCESSED_DATE = SYSDATE, PROCESSED_IND = '1' WHERE PROC_KEY IN " & sqlAdjustments)
                        lstadj.Clear()
                    End If

                End If
            End If

        End If

        ' Receipts
        If dst.Tables.Contains("RECEIPTS") Then
            If dst.Tables("RECEIPTS").Rows.Count > 0 Then
                Update_Record_TDA("RECEIPTS")
                dst.Tables("RECEIPTS").Rows.Clear()

                ASCMAIN1.sql = " BEGIN DECLARE TRX_NO_X VARCHAR2(10); CURSOR C1 IS " _
                    & "  SELECT DISTINCT CONTAINER, PO, INVOICE  FROM CONV.CFG_RECEIPTSADJ;" _
                    & " BEGIN FOR R1 IN C1 LOOP" _
                    & " TRX_NO_X := TAPCTLN1('EDTTRXN1.TRX_NO',1);" _
                    & " INSERT INTO EDTTRXN1" _
                    & " SELECT TRX_NO_X TRX_NO, R.TRANSN TRX_LNO," _
                    & " 'REC' TRANS_TYPE, TRX_NO_X TRANS_NUM," _
                    & " TO_DATE(R.ENTRYDATE,'YYYY-MM-DD') TRANS_DATE," _
                    & " SUBSTR(R.PO,3) PO_ORDER_NO," _
                    & " R.CONTAINER BUYER," _
                    & " R.ITEMN REASON_CODE," _
                    & " R.SEAL OPERATOR," _
                    & " ICTITEM1.ITEM_CODE ITEM_CODE," _
                    & " R.INVOICE LOCATION," _
                    & " R.QTY TRAN_QTY," _
                    & " '0' PROCESS_IND," _
                    & " DECODE(R.WAREHOUSE,'RNG','CLA',R.WAREHOUSE) WHSE_CODE," _
                    & " SYSDATE IMPORT_DATE" _
                    & " FROM CONV.CFG_RECEIPTSADJ R, ICTITEM1" _
                    & " WHERE ICTITEM1.ITEM_ALT_SORT (+) = R.ITEMN" _
                    & " AND NVL(R.CONTAINER,'?') = NVL(R1.CONTAINER,'?')" _
                    & " AND NVL(R.PO,'?') = NVL(R1.PO,'?')" _
                    & " AND NVL(R.INVOICE,'?') = NVL(R1.INVOICE,'?');" _
                    & " END LOOP; END; END;"

                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                ASCDATA1.ExecuteSQL("DELETE FROM CONV.CFG_RECEIPTSADJ")

            End If
        End If

        ' Customer Master 
        If dst.Tables.Contains("CUSTO") Then
            If dst.Tables("CUSTO").Rows.Count > 0 Then

                ' Update only rows in the tables where the CUST_NO_3PL is NULL.
                ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM CONV.CFG_CUSTO WHERE NVL(TRNSTATUS, 'A') <> 'E';" _
                    & " BEGIN FOR R1 IN C1 LOOP " _
                    & " Update ARTCUST2 SET CUST_NO_3PL = TRIM(R1.CLRCUST1), CUST_STORE_NO_3PL = TRIM(R1.CLRCUST2) " _
                    & " WHERE CUST_CODE = TRIM(R1.IPLBCUST1) AND CUST_STORE_NO = TRIM(R1.IPLBCUST2) AND CUST_NO_3PL IS NULL" _
                    & " AND NVL(R1.CLRCUST1, '0') <> '0' AND NVL(R1.CLRCUST2, '0') <> '0';" _
                    & " END LOOP; END; END;"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                ASCMAIN1.sql = "Delete from CONV.CFG_CUSTO"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                dst.Tables("CUSTO").Rows.Clear()
            End If
        End If

        ' Item Master
        If dst.Tables.Contains("ITEMMAST") Then
            If dst.Tables("ITEMMAST").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("ITEMMAST").Rows.Clear()
            End If
        End If

        ' Ship Via Scac Codes
        If dst.Tables.Contains("SHPVIASCAC") Then
            If dst.Tables("SHPVIASCAC").Rows.Count > 0 Then
                ' Ony add the new ones.
                ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM CONV.CFG_SHPVIASCAC WHERE OHSHPR NOT IN (SELECT SHIP_VIA_CODE FROM SOTSVIA1);" _
                    & " BEGIN FOR R1 IN C1 LOOP" _
                    & "     INSERT INTO SOTSVIA1  (SHIP_VIA_CODE, SHIP_VIA_DESC, SHIP_VIA_SCAC, SHIP_VIA_STATUS) " _
                    & "     VALUES " _
                    & "     (R1.OHSHPR, R1.DESCRIPT, substr(R1.SCSCAC, 1, 4), 'A');" _
                    & " END LOOP; END; END;"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                ' Modify Description Changes.
                ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS SELECT * FROM CONV.CFG_SHPVIASCAC WHERE OHSHPR IN (SELECT SHIP_VIA_CODE FROM SOTSVIA1);" _
                    & " BEGIN FOR R1 IN C1 LOOP" _
                    & "     UPDATE SOTSVIA1 SET SHIP_VIA_DESC = R1.DESCRIPT WHERE SHIP_VIA_CODE = R1.OHSHPR ;" _
                    & " END LOOP; END; END;"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                ASCMAIN1.sql = "Delete from CONV.CFG_SHPVIASCAC"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                dst.Tables("SHPVIASCAC").Rows.Clear()
            End If
        End If

        If dst.Tables.Contains("BRAND") Then
            If dst.Tables("BRAND").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("BRAND").Rows.Clear()
            End If
        End If

        If dst.Tables.Contains("SUBBRAND") Then
            If dst.Tables("SUBBRAND").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("SUBBRAND").Rows.Clear()
            End If
        End If

        If dst.Tables.Contains("INVENTORY") Then
            If dst.Tables("INVENTORY").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("INVENTORY").Rows.Clear()
            End If
        End If

        If dst.Tables.Contains("INVOICE") Then
            If dst.Tables("INVOICE").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("INVOICE").Rows.Clear()
            End If
        End If

        If dst.Tables.Contains("SHIPRULE") Then
            If dst.Tables("SHIPRULE").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("SHIPRULE").Rows.Clear()
            End If
        End If

        If dst.Tables.Contains("VENDORS") Then
            If dst.Tables("VENDORS").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("VENDORS").Rows.Clear()
            End If
        End If

        If dst.Tables.Contains("CUSTOMERS") Then
            If dst.Tables("CUSTOMERS").Rows.Count > 0 Then
                ' Currently Do Nothing
                dst.Tables("CUSTOMERS").Rows.Clear()
            End If
        End If

    End Sub

    Sub Create_Work_Table()

    End Sub

    Overrides Sub Post_Update_Archive()
        MyBase.Post_Update_Archive()

        Try

        Catch ex As Exception

        End Try

    End Sub

    Private Function ProcessClarinsDailyShip() As Boolean

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_ID = ""
            environ.APP_KEY = "DAILYSHIP"
            clsWHCIMP01 = New WHCIMPO1(environ)

            If clsWHCIMP01.ErrorMessages.Count > 0 Then
                Return False
            End If

            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Daily Ship" & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Private Function ProcessClarinsCancellations() As Boolean

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_ID = ""
            environ.APP_KEY = "CNLORDS"
            clsWHCIMP01 = New WHCIMPO1(environ)

            If clsWHCIMP01.ErrorMessages.Count > 0 Then
                Return False
            End If

            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Order Cancellations" & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Private Function ProcessClarinsOrderConfirmationFiles() As Boolean

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "ABSOH,ABSOD"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.OrderConfirmation)

            If Not dst.Tables.Contains("ABSOH") Then
                Return True
            End If

            Fill_Records("ABSOH")
            Fill_Records("ABSOD")

            If dst.Tables("ABSOH").Rows.Count > 0 Then
                Dim IMPORT_NO As String = ASCMAIN1.Next_Control_No("SOTPICKC.IMPORT_NO")

                Dim sql As String = String.Empty

                For Each rowABSOH As DataRow In dst.Tables("ABSOH").Select("", "ERRREASON DESC")
                    sql = "IMPORT_NO = '" & IMPORT_NO & "' AND PICK_NO = '" & (rowABSOH.Item("ABSPICKNBR") & String.Empty).ToString.Trim.PadLeft(10, "0") & "'" _
                        & " AND PICK_LNO = 0 AND BATCH_ID = '" & rowABSOH.Item("BATCHID") & String.Empty & "'"

                    If dst.Tables("SOTPICKC").Select(sql).Length > 0 Then
                        Continue For
                    End If

                    Dim rowSOTPICKC As DataRow = dst.Tables("SOTPICKC").NewRow
                    rowSOTPICKC.Item("IMPORT_NO") = IMPORT_NO
                    rowSOTPICKC.Item("PICK_NO") = (rowABSOH.Item("ABSPICKNBR") & String.Empty).ToString.Trim.PadLeft(10, "0")
                    rowSOTPICKC.Item("PICK_LNO") = 0
                    rowSOTPICKC.Item("BATCH_ID") = rowABSOH.Item("BATCHID") & String.Empty
                    rowSOTPICKC.Item("ORDR_NO") = (rowABSOH.Item("ABSORDNBR") & String.Empty).ToString.Trim.PadLeft(10, "0")
                    rowSOTPICKC.Item("ORDR_NO_3PL") = rowABSOH.Item("CLRORDNBR") & String.Empty
                    rowSOTPICKC.Item("ERROR_REASON") = rowABSOH.Item("ERRREASON") & String.Empty
                    rowSOTPICKC.Item("ERROR_IND") = "0"
                    dst.Tables("SOTPICKC").Rows.Add(rowSOTPICKC)
                Next

                ' Extract Only the Detail Errors.
                For Each rowABSOD As DataRow In dst.Tables("ABSOD").Select("ISNULL(ERRREASON, '*') <> '*'")
                    Dim ERRREASON As String = rowABSOD.Item("ERRREASON") & String.Empty
                    ERRREASON = ERRREASON.Trim
                    If ERRREASON.Length = 0 Then Continue For

                    Dim ABSORDNBR As String = (rowABSOD.Item("ABSORDNBR") & String.Empty).ToString.Trim.PadLeft(10, "0")
                    If dst.Tables("ABSOH").Select("ABSORDNBR = '" & ABSORDNBR & "'").Length = 0 Then Continue For

                    Dim rowABSOH As DataRow = dst.Tables("ABSOH").Select("ABSORDNBR = '" & ABSORDNBR & "'")(0)

                    sql = "IMPORT_NO = '" & IMPORT_NO & "' AND PICK_NO = '" & (rowABSOH.Item("ABSPICKNBR") & String.Empty).ToString.Trim.PadLeft(10, "0") & "'" _
                            & " AND PICK_LNO = " & rowABSOD.Item("LINENBR") & " AND BATCH_ID = 'D" & rowABSOH.Item("BATCHID") & String.Empty & "'"

                    If dst.Tables("SOTPICKC").Select(sql).Length > 0 Then
                        Continue For
                    End If

                    Dim rowSOTPICKC As DataRow = dst.Tables("SOTPICKC").NewRow
                    rowSOTPICKC.Item("IMPORT_NO") = IMPORT_NO
                    rowSOTPICKC.Item("PICK_NO") = (rowABSOH.Item("ABSPICKNBR") & String.Empty).ToString.Trim.PadLeft(10, "0")
                    rowSOTPICKC.Item("PICK_LNO") = rowABSOD.Item("LINENBR")
                    rowSOTPICKC.Item("BATCH_ID") = "D" & rowABSOH.Item("BATCHID") & String.Empty
                    rowSOTPICKC.Item("ORDR_NO") = (rowABSOH.Item("ABSORDNBR") & String.Empty).ToString.Trim.PadLeft(10, "0")
                    rowSOTPICKC.Item("ORDR_NO_3PL") = rowABSOH.Item("CLRORDNBR") & String.Empty
                    rowSOTPICKC.Item("ERROR_REASON") = rowABSOD.Item("ERRREASON") & String.Empty
                    rowSOTPICKC.Item("ERROR_IND") = "0"
                    dst.Tables("SOTPICKC").Rows.Add(rowSOTPICKC)
                Next
            End If

            R = dst.Tables("SOTPICKC").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Order Confirmation " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try

    End Function

    Private Function ProcessClarinsShipmentConfirmationFiles() As Boolean

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "SHIPHDR,SHIPDTL,TRACK,CARTON"
            clsWHCIMP01 = New WHCIMPO1(environ)

            Dim rowEDT945T1 As DataRow = Nothing
            Dim rowEDT945T2 As DataRow = Nothing
            Dim rowSOTPICK2 As DataRow = Nothing
            Dim tblSOTPICK2 As DataTable = Nothing
            Dim rowSOTORDR1 As DataRow = Nothing
            Dim rowSOTORDR5 As DataRow = Nothing

            Dim EDI_DOC_SEQ_NO As String = String.Empty
            Dim EDI_DTL_SEQ As Int32 = 0
            Dim EDI_SHIP_QTY As Int32 = 0

            Dim ITEM_CODE As String = String.Empty
            Dim ORDR_QTY_SHIP As Int32 = 0
            Dim CART_NO As String = String.Empty

            Dim sql As String = String.Empty
            Dim filesProcessed As New List(Of String)
            Dim CART_NO_LENGTH As Int32 = 0

            CreateTableAdaptors(ProcessType.ShipmentConfirmation)

            If Not dst.Tables.Contains("SHIPHDR") Then
                Exit Function
            End If

            ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_SHIPHDR SET IMPORT_DATE = SYSDATE WHERE IMPORT_DATE IS NULL")

            ' 02/05/2019 - Do not process any shipments until all 4 files have arrived
            ' Resets previous marked issues as fresh so they can be reprocessed
            ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_SHIPHDR SET PROCESS_IND = NULL WHERE PROCESS_IND = 'W'")
            ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_SHIPDTL SET PROCESS_IND = NULL WHERE PROCESS_IND = 'W'")
            ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_CARTON SET PROCESS_IND = NULL WHERE PROCESS_IND = 'W'")
            ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_TRACK SET PROCESS_IND = NULL WHERE PROCESS_IND = 'W'")

            Dim wkTable As String = ASCMAIN1.Temp_Table("Select OHINVN INV_NO from CONV.CFG_SHIPHDR WHERE NVL(PROCESS_IND, '0') = '0'")

            ' Update fields that are not pulled in from Clarins
            sql = " UPDATE CONV.CFG_SHIPHDR SET SHIP_BOL_NO = (SELECT SHIP_BOL_NO FROM SOTPICK1 WHERE PICK_NO = CONV.CFG_SHIPHDR.ABSPICKNBR) WHERE SHIP_BOL_NO IS NULL AND OHINVN IN (SELECT INV_NO FROM " & wkTable & ")"
            ASCDATA1.ExecuteSQL(sql)

            sql = "UPDATE CONV.CFG_SHIPHDR SET CUST_CODE = (SELECT CUST_CODE FROM SOTORDR1 WHERE ORDR_NO = CONV.CFG_SHIPHDR.ABSORDNBR) WHERE CUST_CODE IS NULL AND OHINVN IN (SELECT INV_NO FROM " & wkTable & ")"
            ASCDATA1.ExecuteSQL(sql)

            sql = "UPDATE CONV.CFG_SHIPHDR SET SHIP_ADDR_CODE = (SELECT SHIP_ADDR_CODE FROM SOTSHIP1 WHERE SHIP_BOL_NO = CONV.CFG_SHIPHDR.SHIP_BOL_NO) WHERE SHIP_ADDR_CODE IS NULL AND OHINVN IN (SELECT INV_NO FROM " & wkTable & ")"
            ASCDATA1.ExecuteSQL(sql)

            ' 02/05/2019 - Do not process any shipments until all 4 files have arrived
            ' Strip off leading Alpha chars from the file name so the datetime stamp part is remaining
            sql = " UPDATE CONV.CFG_SHIPHDR SET IMPORT_FILENAME = SUBSTR(IMPORT_FILENAME, 8) WHERE UPPER(IMPORT_FILENAME) LIKE 'SHIPHDR%'"
            ASCDATA1.ExecuteSQL(sql)

            sql = " UPDATE CONV.CFG_SHIPDTL SET IMPORT_FILENAME = SUBSTR(IMPORT_FILENAME, 8) WHERE UPPER(IMPORT_FILENAME) LIKE 'SHIPDTL%'"
            ASCDATA1.ExecuteSQL(sql)

            sql = " UPDATE CONV.CFG_CARTON SET IMPORT_FILENAME = SUBSTR(IMPORT_FILENAME, 7) WHERE UPPER(IMPORT_FILENAME) LIKE 'CARTON%'"
            ASCDATA1.ExecuteSQL(sql)

            sql = " UPDATE CONV.CFG_TRACK SET IMPORT_FILENAME = SUBSTR(IMPORT_FILENAME, 6) WHERE UPPER(IMPORT_FILENAME) LIKE 'TRACK%'"
            ASCDATA1.ExecuteSQL(sql)

            Dim lstSHIP_BOL_NO As New List(Of String)
            Dim lstMANIFEST_NO As New List(Of String)

            ' New code to verify all invoices have a SHIPDTL, CARTON AND TRACK ENTRY.
            sql = "Select DISTINCT CUST_CODE, SHIP_ADDR_CODE, OHCSPO, OHMANN, IMPORT_FILENAME, SHIP_BOL_NO from CONV.CFG_SHIPHDR " _
                    & " WHERE OHINVN IN " _
                    & " (SELECT INV_NO FROM " & wkTable & ")" _
                    & " AND IMPORT_FILENAME NOT IN  " _
                    & " (SELECT IMPORT_FILENAME FROM CONV.CFG_SHIPDTL WHERE PROCESS_IND IS NULL)"
            Dim tblMissingData As DataTable = ASCDATA1.GetDataTable(sql)
            If tblMissingData.Rows.Count > 0 Then
                For Each rowgroup As DataRow In ASCDATA1.SelectDistinct(tblMissingData, New String() {"IMPORT_FILENAME", "CUST_CODE", "OHCSPO"}).Select("", "IMPORT_FILENAME, CUST_CODE, OHCSPO")
                    Dim IMPORT_FILENAME As String = rowgroup.Item("IMPORT_FILENAME") & String.Empty
                    Dim CUST_CODE As String = rowgroup.Item("CUST_CODE") & String.Empty
                    Dim OHCSPO As String = rowgroup.Item("OHCSPO") & String.Empty

                    Dim msg As String = "SHIPDTL" & rowgroup.Item("IMPORT_FILENAME") & " is missing. No Detail data for Customer: " & rowgroup.Item("CUST_CODE") & ", PO Number: " & rowgroup.Item("OHCSPO") & ": "

                    For Each row As DataRow In tblMissingData.Select("IMPORT_FILENAME = '" & IMPORT_FILENAME & "' and CUST_CODE = '" & CUST_CODE & "' and OHCSPO = '" & OHCSPO & "'", "SHIP_ADDR_CODE, OHMANN")

                        msg &= Environment.NewLine & Space(5) & "Ship To: " & row.Item("SHIP_ADDR_CODE") & ", Clarins Manifest No: " & row.Item("OHMANN")
                        Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & String.Empty
                        Dim MANIFEST_NO As String = row.Item("OHMANN") & String.Empty

                        If SHIP_BOL_NO.Length > 0 AndAlso Not lstSHIP_BOL_NO.Contains(SHIP_BOL_NO) Then
                            lstSHIP_BOL_NO.Add(SHIP_BOL_NO)
                        End If

                        If MANIFEST_NO.Length > 0 AndAlso Not lstMANIFEST_NO.Contains(MANIFEST_NO) Then
                            lstMANIFEST_NO.Add(MANIFEST_NO)
                        End If
                    Next
                    ErrorMessages.Add(msg)
                Next
            End If

            sql = "Select DISTINCT CUST_CODE, SHIP_ADDR_CODE, OHCSPO, OHMANN, IMPORT_FILENAME, SHIP_BOL_NO from CONV.CFG_SHIPHDR " _
                    & " WHERE OHINVN IN " _
                    & " (SELECT INV_NO FROM " & wkTable & ")" _
                    & " AND IMPORT_FILENAME NOT IN  " _
                    & " (SELECT IMPORT_FILENAME FROM CONV.CFG_CARTON WHERE PROCESS_IND IS NULL)"
            tblMissingData = ASCDATA1.GetDataTable(sql)
            If tblMissingData.Rows.Count > 0 Then
                For Each rowgroup As DataRow In ASCDATA1.SelectDistinct(tblMissingData, New String() {"IMPORT_FILENAME", "CUST_CODE", "OHCSPO"}).Select("", "IMPORT_FILENAME, CUST_CODE, OHCSPO")
                    Dim IMPORT_FILENAME As String = rowgroup.Item("IMPORT_FILENAME") & String.Empty
                    Dim CUST_CODE As String = rowgroup.Item("CUST_CODE") & String.Empty
                    Dim OHCSPO As String = rowgroup.Item("OHCSPO") & String.Empty

                    Dim msg As String = "CARTON" & rowgroup.Item("IMPORT_FILENAME") & " is missing. No Carton data for Customer: " & rowgroup.Item("CUST_CODE") & ", PO Number: " & rowgroup.Item("OHCSPO") & ": "

                    For Each row As DataRow In tblMissingData.Select("IMPORT_FILENAME = '" & IMPORT_FILENAME & "' and CUST_CODE = '" & CUST_CODE & "' and OHCSPO = '" & OHCSPO & "'", "SHIP_ADDR_CODE, OHMANN")

                        msg &= Environment.NewLine & Space(5) & "Ship To: " & row.Item("SHIP_ADDR_CODE") & ", Clarins Manifest No: " & row.Item("OHMANN")
                        Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & String.Empty
                        Dim MANIFEST_NO As String = row.Item("OHMANN") & String.Empty

                        If SHIP_BOL_NO.Length > 0 AndAlso Not lstSHIP_BOL_NO.Contains(SHIP_BOL_NO) Then
                            lstSHIP_BOL_NO.Add(SHIP_BOL_NO)
                        End If

                        If MANIFEST_NO.Length > 0 AndAlso Not lstMANIFEST_NO.Contains(MANIFEST_NO) Then
                            lstMANIFEST_NO.Add(MANIFEST_NO)
                        End If
                    Next
                    ErrorMessages.Add(msg)
                Next
            End If

            sql = "Select DISTINCT CUST_CODE, SHIP_ADDR_CODE, OHCSPO, OHMANN, IMPORT_FILENAME, SHIP_BOL_NO from CONV.CFG_SHIPHDR " _
                    & " WHERE OHINVN IN " _
                    & " (SELECT INV_NO FROM " & wkTable & ")" _
                    & " AND IMPORT_FILENAME NOT IN  " _
                    & " (SELECT IMPORT_FILENAME FROM CONV.CFG_TRACK WHERE PROCESS_IND IS NULL)"
            tblMissingData = ASCDATA1.GetDataTable(sql)
            If tblMissingData.Rows.Count > 0 Then
                For Each rowgroup As DataRow In ASCDATA1.SelectDistinct(tblMissingData, New String() {"IMPORT_FILENAME", "CUST_CODE", "OHCSPO"}).Select("", "IMPORT_FILENAME, CUST_CODE, OHCSPO")
                    Dim IMPORT_FILENAME As String = rowgroup.Item("IMPORT_FILENAME") & String.Empty
                    Dim CUST_CODE As String = rowgroup.Item("CUST_CODE") & String.Empty
                    Dim OHCSPO As String = rowgroup.Item("OHCSPO") & String.Empty

                    Dim msg As String = "TRACK" & rowgroup.Item("IMPORT_FILENAME") & " is missing. No Tracking data for Customer: " & rowgroup.Item("CUST_CODE") & ", PO Number: " & rowgroup.Item("OHCSPO") & ": "

                    For Each row As DataRow In tblMissingData.Select("IMPORT_FILENAME = '" & IMPORT_FILENAME & "' and CUST_CODE = '" & CUST_CODE & "' and OHCSPO = '" & OHCSPO & "'", "SHIP_ADDR_CODE, OHMANN")

                        msg &= Environment.NewLine & Space(5) & "Ship To: " & row.Item("SHIP_ADDR_CODE") & ", Clarins Manifest No: " & row.Item("OHMANN")
                        Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & String.Empty
                        Dim MANIFEST_NO As String = row.Item("OHMANN") & String.Empty

                        If SHIP_BOL_NO.Length > 0 AndAlso Not lstSHIP_BOL_NO.Contains(SHIP_BOL_NO) Then
                            lstSHIP_BOL_NO.Add(SHIP_BOL_NO)
                        End If

                        If MANIFEST_NO.Length > 0 AndAlso Not lstMANIFEST_NO.Contains(MANIFEST_NO) Then
                            lstMANIFEST_NO.Add(MANIFEST_NO)
                        End If
                    Next
                    ErrorMessages.Add(msg)
                Next
            End If

            Dim records As Int32 = 0
            For Each SHIP_BOL_NO As String In lstSHIP_BOL_NO
                If SHIP_BOL_NO.Length = 0 Then Continue For
                sql = "UPDATE CONV.CFG_SHIPHDR SET PROCESS_IND = 'W' WHERE OHINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "') AND PROCESS_IND IS NULL"
                records += ASCDATA1.ExecuteSQL(sql)
            Next

            For Each MANIFEST_NO As String In lstMANIFEST_NO
                If MANIFEST_NO.Length = 0 Then Continue For
                sql = "UPDATE CONV.CFG_SHIPHDR SET PROCESS_IND = 'W' WHERE OHINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE OHMANN = '" & MANIFEST_NO & "') AND PROCESS_IND IS NULL"
                records += ASCDATA1.ExecuteSQL(sql)
            Next

            If records > 0 Then
                ErrorMessages.Add(records & " Invoices were removed from Shipment Confirmation process because of missing files.")
            End If

            sql = "UPDATE CONV.CFG_SHIPDTL SET PROCESS_IND = 'W' WHERE SAINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE PROCESS_IND = 'W') AND PROCESS_IND IS NULL"
            ASCDATA1.ExecuteSQL(sql)

            sql = "UPDATE CONV.CFG_CARTON SET PROCESS_IND = 'W' WHERE CHINVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE PROCESS_IND = 'W') AND PROCESS_IND IS NULL"
            ASCDATA1.ExecuteSQL(sql)

            sql = "UPDATE CONV.CFG_TRACK SET PROCESS_IND = 'W' WHERE INVN IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE PROCESS_IND = 'W') AND PROCESS_IND IS NULL"
            ASCDATA1.ExecuteSQL(sql)

            ' Remove any shipment records from the work table where the PROCESS_IND = 'W'
            sql = "DELETE FROM " & wkTable & " WHERE INV_NO IN (SELECT OHINVN FROM CONV.CFG_SHIPHDR WHERE PROCESS_IND = 'W')"
            ASCDATA1.ExecuteSQL(sql)

            Fill_Records("SHIPHDR", String.Empty, True, "SELECT * FROM CONV.CFG_SHIPHDR WHERE OHINVN IN (SELECT INV_NO FROM " & wkTable & ") AND NVL(PROCESS_IND, '0') = '0'")
            Fill_Records("SHIPDTL", String.Empty, True, "SELECT * FROM CONV.CFG_SHIPDTL WHERE SAINVN IN (SELECT INV_NO FROM " & wkTable & ") AND NVL(PROCESS_IND, '0') = '0'")
            Fill_Records("CARTON", String.Empty, True, "SELECT * FROM CONV.CFG_CARTON WHERE CHINVN IN (SELECT INV_NO FROM " & wkTable & ") AND NVL(PROCESS_IND, '0') = '0'")
            Fill_Records("TRACK", String.Empty, True, "SELECT * FROM CONV.CFG_TRACK WHERE INVN IN (SELECT INV_NO FROM " & wkTable & ") AND NVL(PROCESS_IND, '0') = '0'")
            Fill_Records("SOTSVIA1", String.Empty, True, "SELECT * FROM SOTSVIA1")
            Dim listGroupNos As New List(Of String)

            CART_NO_LENGTH = dst.Tables("SOTCART1").Columns("CART_NO").MaxLength

            For Each rowHeader As DataRow In dst.Tables("SHIPHDR").Select("", "OHORDN")
                Dim ORDER_NUMBER As String = rowHeader.Item("OHORDN") & String.Empty
                Dim ORDR_NO As String = rowHeader.Item("ABSORDNBR") & String.Empty
                Dim PICK_NO As String = rowHeader.Item("ABSPICKNBR") & String.Empty
                Dim INVOICE_NUMBER As String = rowHeader.Item("OHINVN") & String.Empty
                Dim OHCSPO As String = rowHeader.Item("OHCSPO") & String.Empty

                rowHeader.Item("PROCESS") = ProcessState.ProcessShipment

                ' Total Item Units Shipped should be equal in the Details and Tracking files
                ' Compare Details against Tracking, then Tracking against Details
                If dst.Tables("SHIPDTL").Select("[SAINVN] = '" & INVOICE_NUMBER & "'").Length = 0 Then
                    ErrorMessages.Add("No details for header record. Order Number: " & ORDR_NO)
                    rowHeader.Item("PROCESS") = ProcessState.NoDetails
                    Continue For
                End If

                ' Pick Tick must exist and have a status of Pick
                Fill_Records("SOTPICK1", String.Empty, False, "SELECT * FROM SOTPICK1 WHERE PICK_NO = '" & PICK_NO & "'")
                Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)

                If rowSOTPICK1 Is Nothing OrElse rowSOTPICK1.Item("PICK_STATUS") & String.Empty <> "P" Then
                    ' Assume a Scrap shiment
                    If rowSOTPICK1 IsNot Nothing Then
                        ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_SHIPHDR SET PROCESS_IND = 'S' WHERE NVL(PROCESS_IND, '1') = '1' AND ABSPICKNBR = '" & PICK_NO & "'")
                    Else
                        ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_SHIPHDR SET PROCESS_IND = 'S' WHERE NVL(PROCESS_IND, '1') = '1' AND OHORDN = '" & ORDER_NUMBER & "' AND OHINVN = '" & INVOICE_NUMBER & "'")
                    End If
                    rowHeader.Item("PROCESS") = ProcessState.PickTicketNotInPick

                    'Fill_Records("SOTPICK1", String.Empty, False, "SELECT * FROM SOTPICK1 WHERE PICK_NO = '" & PICK_NO & "'")
                    Dim errorMessage As String = "Problem with Pick Ticket associated with Clarins Invoice Number: " & INVOICE_NUMBER & " / Order Number: " & ORDER_NUMBER _
                                                 & ", ABSolution PO Number: " & OHCSPO & ", this has been flagged as a SCRAP order. THIS SHIPMENT WILL NOT BE PROCESSED!!!"
                    If ORDR_NO.Length > 0 Then
                        errorMessage &= " - ABSolution Order Number: " & ORDR_NO
                    Else
                        errorMessage &= " - Cannot determine ABSolution Order Number"
                    End If

                    errorMessage &= Environment.NewLine
                    ErrorMessages.Add(errorMessage)

                    Continue For
                End If

                Fill_Records("SOTORDR1", ORDR_NO, False)
                Fill_Records("SOTORDR5", New Object() {ORDR_NO, "ST"}, False)

                If dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                    rowHeader.Item("PROCESS") = ProcessState.InvalidOrderNumber
                    ErrorMessages.Add("Invalid Order Number: " & ORDR_NO)
                    Continue For
                End If

                For Each rowDetails As DataRow In dst.Tables("SHIPDTL").Select("[SAINVN] = '" & INVOICE_NUMBER & "'")
                    ITEM_CODE = rowDetails.Item("SAITEM") & String.Empty
                    sql = "[SAINVN] = '" & INVOICE_NUMBER & "' AND SAITEM = '" & ITEM_CODE & "'"
                    ORDR_QTY_SHIP = Val(dst.Tables("SHIPDTL").Compute("SUM(SAQTYS)", sql) & String.Empty)

                    sql = "[CHINVN] = '" & INVOICE_NUMBER & "' AND CDITEM = '" & ITEM_CODE & "'"
                    If Val(dst.Tables("CARTON").Compute("SUM(WHSE_QTY_SHIP)", sql) & String.Empty) > ORDR_QTY_SHIP Then
                        ErrorMessages.Add("Detail and Tracking totals do not match for Sales Order: " & ORDR_NO & ", Item Code: " & ITEM_CODE & Environment.NewLine)
                        rowHeader.Item("PROCESS") = ProcessState.MismatchDetailsAndTracking
                    End If
                Next
            Next

            ' Need to get the cancelled orders into the tables to be processed
            For Each rowHeader As DataRow In dst.Tables("SHIPHDR").Select("PROCESS = " & ProcessState.ProcessShipment, "OHORDN")
                Dim ORDR_NO As String = rowHeader.Item("ABSORDNBR") & String.Empty
                'ORDR_NO = ASCMAIN1.Format_Field(ORDR_NO, "ORDR_NO")
                rowSOTORDR1 = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)

                If Not listGroupNos.Contains(rowSOTORDR1.Item("ORDR_GROUP_NO")) Then
                    listGroupNos.Add(rowSOTORDR1.Item("ORDR_GROUP_NO"))
                End If
            Next

            If listGroupNos.Count > 0 AndAlso dst.Tables.Contains("CNLORDS") Then
                ASCMAIN1.sql = " Select SOTPICK1.*, CNLORDS.ohordN, SOTORDR1.ORDR_CUST_PO, SOTSHIP1.SHIP_VIA_CODE" _
                    & " from SOTPICK1, SOTPICKO, CONV.CFG_CNLORDS CNLORDS, SOTSHIP1, SOTORDR1" _
                    & " where SOTPICK1.pick_no = SOTPICKO.pick_no" _
                    & " and SOTPICKO.ordr_no_3pl = CNLORDS.ohordN" _
                    & " and SOTPICK1.pick_status = 'P'" _
                    & " and SOTPICK1.SHIP_BOL_NO = sotship1.SHIP_BOL_NO" _
                    & " AND SOTSHIP1.SHIP_STATUS = 'P'" _
                    & " AND SOTSHIP1.ORDR_GROUP_NO IN (" & "'" & String.Join("', '", listGroupNos.ToArray) & "'" & ")" _
                    & " AND SOTPICK1.PICK_NO NOT IN (SELECT EDI_PICK_NO FROM EDT945T1)" _
                    & " AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"

                Dim tblSOTPICK1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                For Each rowSOTPICK1 As DataRow In tblSOTPICK1.Select("", "PICK_NO")
                    Dim rowHeader As DataRow = dst.Tables("SHIPHDR").NewRow

                    If dst.Tables("SOTPICK1").Rows.Find(rowSOTPICK1.Item("PICK_NO")) Is Nothing Then
                        Fill_Records("SOTPICK1", String.Empty, False, "SELECT * FROM SOTPICK1 WHERE PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "' and PICK_STATUS = 'P'")
                    End If

                    If dst.Tables("SOTORDR1").Rows.Find(rowSOTPICK1.Item("ORDR_NO")) Is Nothing Then
                        Fill_Records("SOTORDR1", rowSOTPICK1.Item("ORDR_NO"), False)
                    End If

                    If dst.Tables("SOTORDR5").Rows.Find(New Object() {rowSOTPICK1.Item("ORDR_NO"), "ST"}) Is Nothing Then
                        Fill_Records("SOTORDR5", New Object() {rowSOTPICK1.Item("ORDR_NO"), "ST"}, False)
                    End If

                    rowSOTORDR5 = dst.Tables("SOTORDR5").Rows.Find(New Object() {rowSOTPICK1.Item("ORDR_NO"), "ST"})

                    rowHeader.Item("OHCONO") = 14 ' Company Code
                    rowHeader.Item("OHORDN") = rowSOTPICK1.Item("OHORDN") ' Clarin's Order Number
                    rowHeader.Item("ABSORDNBR") = rowSOTPICK1.Item("ORDR_NO") ' ABS Order No.
                    rowHeader.Item("ABSPICKNBR") = rowSOTPICK1.Item("PICK_NO") ' ABS Pick No

                    rowHeader.Item("OHINVN") = String.Empty ' Inv Number
                    rowHeader.Item("OHSHPD") = DateTime.Now.ToString("yyMMdd") ' Ship Date - defualt to today
                    rowHeader.Item("OHCSPO") = rowSOTPICK1.Item("ORDR_CUST_PO") ' Customer PO
                    rowHeader.Item("OHMANN") = "" ' BOL Number
                    rowHeader.Item("OHFRCG") = 0 ' Freight Charge 
                    rowHeader.Item("OHSHPR") = rowSOTPICK1.Item("SHIP_VIA_CODE") ' Ship Via Code
                    rowHeader.Item("OHTUNT") = 0 ' Units Shipped
                    rowHeader.Item("OHWGHT") = 0 ' Weight Ship
                    rowHeader.Item("PROCESS") = ProcessState.ProcessShipment

                    ' Address Information.
                    rowHeader.Item("OHSNAM") = rowSOTORDR5.Item("CUST_NAME")
                    rowHeader.Item("OHSAD1") = rowSOTORDR5.Item("CUST_ADDR1")
                    rowHeader.Item("OHSAD2") = rowSOTORDR5.Item("CUST_ADDR2")
                    rowHeader.Item("OHSCTY") = rowSOTORDR5.Item("CUST_CITY")
                    rowHeader.Item("OHSSTA") = rowSOTORDR5.Item("CUST_STATE")
                    Dim CUST_ZIP_CODE As String = rowSOTORDR5.Item("CUST_ZIP_CODE")
                    If CUST_ZIP_CODE.Contains("-") Then
                        rowHeader.Item("OHSZIP") = CUST_ZIP_CODE.Split("-")(0)
                        rowHeader.Item("OHSZP2") = CUST_ZIP_CODE.Split("-")(1)
                    ElseIf CUST_ZIP_CODE.Length <= 5 Then
                        rowHeader.Item("OHSZIP") = CUST_ZIP_CODE
                    ElseIf CUST_ZIP_CODE.Length = 9 Then
                        rowHeader.Item("OHSZIP") = CUST_ZIP_CODE.Substring(0, 5).Trim
                        rowHeader.Item("OHSZP2") = CUST_ZIP_CODE.Substring(5).Trim
                    Else
                        rowHeader.Item("OHSZIP") = CUST_ZIP_CODE.Substring(0, 5).Trim
                    End If

                    dst.Tables("SHIPHDR").Rows.Add(rowHeader)
                Next
            End If

            For Each rowHeader As DataRow In dst.Tables("SHIPHDR").Select("PROCESS = " & ProcessState.ProcessShipment, "OHORDN")
                Dim ORDER_NUMBER As String = rowHeader.Item("OHORDN") & String.Empty
                Dim ORDR_NO As String = rowHeader.Item("ABSORDNBR") & String.Empty
                Dim PICK_NO As String = rowHeader.Item("ABSPICKNBR") & String.Empty

                Dim INVOICE_NUMBER As String = rowHeader.Item("OHINVN") & String.Empty
                CART_NO = String.Empty

                EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDI_DOC_SEQ_NO")
                EDI_DOC_SEQ_NO = "1" & EDI_DOC_SEQ_NO.Substring(1) ' Leading 1 so no collisions with Gentran
                EDI_DTL_SEQ = 0

                rowEDT945T1 = dst.Tables("EDT945T1").NewRow
                rowEDT945T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO

                'rowEDT945T1.ITEM("GEN_DOC_NO") = String.Empty
                'rowEDT945T1.ITEM("EDI_ISA_NO") = String.Empty
                'rowEDT945T1.ITEM("EDI_TP_QUAL") = String.Empty
                'rowEDT945T1.ITEM("EDI_TP_ID") = String.Empty
                'rowEDT945T1.ITEM("EDI_OUR_QUAL") = String.Empty
                'rowEDT945T1.ITEM("EDI_OUR_ID") = String.Empty
                'rowEDT945T1.ITEM("EDI_REPORTING_CODE") = String.Empty
                Dim OHSHPD As String = rowHeader.Item("OHSHPD") & String.Empty
                OHSHPD = OHSHPD.Substring(2, 2) & "/" & OHSHPD.Substring(4, 2) & "/" & OHSHPD.Substring(0, 2)
                rowEDT945T1.Item("EDI_SHIPMENT_DATE") = CDate(OHSHPD).ToShortDateString

                Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")(0)
                rowSOTORDR1 = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'")(0)

                rowEDT945T1.Item("EDI_SHIPMENT_ID") = rowSOTPICK1.Item("SHIP_BOL_NO")
                rowEDT945T1.Item("EDI_PICK_NO") = rowSOTPICK1.Item("PICK_NO") & String.Empty

                rowEDT945T1.Item("EDI_ORDR_CUST_PO") = rowHeader.Item("OHCSPO")
                rowEDT945T1.Item("EDI_DIVISION_CODE") = ASCMAIN1.DBS_COMPANY
                rowEDT945T1.Item("EDI_BOL_NO") = rowHeader.Item("OHMANN")
                rowEDT945T1.Item("EDI_MASTER_BOL_NO") = rowHeader.Item("OHMANN")

                rowEDT945T1.Item("EDI_SHIPPER_ID_NO") = dst.Tables("TRACK").Compute("MAX(PROORTRK)", "[INVN] = '" & INVOICE_NUMBER & "'") & String.Empty
                ' 12/29/2015 as per Maria
                rowEDT945T1.Item("SHIP_PICKUP_NO") = dst.Tables("TRACK").Compute("MAX([PUN])", "[INVN] = '" & INVOICE_NUMBER & "'") & String.Empty
                rowEDT945T1.Item("SHIP_AUTH_NO") = dst.Tables("TRACK").Compute("MAX([AUTHN])", "[INVN] = '" & INVOICE_NUMBER & "'") & String.Empty

                rowEDT945T1.Item("EDI_FRT_COST") = Val(rowHeader.Item("OHFRCG") & String.Empty)
                rowEDT945T1.Item("EDI_ORDR_SHIP_DATE") = CDate(OHSHPD).ToShortDateString
                'rowEDT945T1.ITEM("EDI_TRANS_METH_CODE") = String.Empty

                Dim SHIP_VIA_CODE As String = rowHeader.Item("OHSHPR") & String.Empty
                Dim rowSOTSVIA1 As DataRow = dst.Tables("SOTSVIA1").Rows.Find(SHIP_VIA_CODE)
                If rowSOTSVIA1 IsNot Nothing Then
                    rowEDT945T1.Item("EDI_CARRIER_SCAC_CODE") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty
                    rowEDT945T1.Item("EDI_CARRIER_NAME") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    rowEDT945T1.Item("EDI_FRT_TERMS") = rowSOTSVIA1.Item("FRT_TERMS") & String.Empty
                End If
                rowEDT945T1.Item("EDI_CARRIER_CODE") = SHIP_VIA_CODE

                rowEDT945T1.Item("EDI_TOTAL_UNITS_SHIPPED") = Val(rowHeader.Item("OHTUNT") & String.Empty)
                rowEDT945T1.Item("EDI_TOTAL_ORDR_WEIGHT") = Val(rowHeader.Item("OHWGHT") & String.Empty)
                rowEDT945T1.Item("EDI_RECEIVED_DATE") = DateTime.Now
                rowEDT945T1.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY

                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                rowSOTORDR5 = dst.Tables("SOTORDR5").Rows.Find(New Object() {ORDR_NO, "ST"})
                rowEDT945T1.Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE") & String.Empty

                ' 11/4/2015 - As per Walter, Update Ship To Address
                If Val(rowHeader.Item("OHTUNT") & String.Empty) > 0 Then
                    rowSOTORDR5.Item("CUST_NAME") = rowHeader.Item("OHSNAM")
                    rowSOTORDR5.Item("CUST_ADDR1") = rowHeader.Item("OHSAD1")
                    rowSOTORDR5.Item("CUST_ADDR2") = rowHeader.Item("OHSAD2")
                    rowSOTORDR5.Item("CUST_CITY") = rowHeader.Item("OHSCTY")
                    rowSOTORDR5.Item("CUST_STATE") = rowHeader.Item("OHSSTA")
                    Dim CUST_ZIP_CODE As String = rowHeader.Item("OHSZIP") & String.Empty
                    If rowHeader.Item("OHSZP2") & String.Empty <> String.Empty Then
                        CUST_ZIP_CODE &= "-" & rowHeader.Item("OHSZP2") & String.Empty
                    End If
                    rowSOTORDR5.Item("CUST_ZIP_CODE") = CUST_ZIP_CODE
                End If

                rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                'rowEDT945T1.ITEM("EDI_TRAILER_NO") = String.Empty
                'rowEDT945T1.Item("EDI_LOAD_ID") = String.Empty
                dst.Tables("EDT945T1").Rows.Add(rowEDT945T1)

                sql = "Select SOTPICK2.*, SOTORDR2.ITEM_CODE, ICTITEM1.ITEM_ALT_SORT"
                sql &= " FROM SOTPICK2, SOTORDR2, ICTITEM1"
                sql &= " WHERE SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO"
                sql &= " AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO"
                sql &= " AND SOTPICK2.PICK_NO = '" & PICK_NO & "'"
                sql &= " AND SOTORDR2.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
                tblSOTPICK2 = ASCDATA1.GetDataTable(sql)
                For Each row As DataRow In tblSOTPICK2.Rows
                    row.Item("PICK_QTY_CONF") = 0
                Next

                ' Canceled Pick Tickets
                If Val(rowHeader.Item("OHTUNT") & String.Empty) = 0 Then
                    For Each row As DataRow In tblSOTPICK2.Select("PICK_QTY > 0")
                        rowEDT945T2 = dst.Tables("EDT945T2").NewRow
                        rowEDT945T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                        EDI_DTL_SEQ += 1
                        rowEDT945T2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                        rowEDT945T2.Item("EDI_CART_NO") = String.Empty
                        rowEDT945T2.Item("EDI_SHIPMENT_STATUS_CODE") = "CN"
                        rowEDT945T2.Item("PICK_LNO") = row.Item("PICK_LNO")
                        rowEDT945T2.Item("PICK_QTY") = row.Item("PICK_QTY")
                        rowEDT945T2.Item("EDI_SHIP_QTY") = 0
                        rowEDT945T2.Item("STYLE_CODE") = row.Item("ITEM_CODE")
                        rowEDT945T2.Item("EDI_CART_WEIGHT") = 0
                        dst.Tables("EDT945T2").Rows.Add(rowEDT945T2)
                    Next

                    rowHeader.Item("PROCESS") = ProcessState.Converted
                    Continue For
                End If

                If dst.Tables("TRACK").Select("[INVN] = '" & INVOICE_NUMBER & "'").Length = 0 AndAlso _
                    dst.Tables("SHIPDTL").Select("[SAORDN] = '" & ORDER_NUMBER & "' AND ISNULL(SAQTYS, 0) > 0").Length > 0 Then
                    ErrorMessages.Add("No tracking information provided for Pick Ticket: " & PICK_NO)
                End If

                Dim processingOverageAcrossMultipleCartons As Boolean = False

                For Each rowCarton As DataRow In dst.Tables("CARTON").Select("[CHINVN] = '" & INVOICE_NUMBER & "'", "CPUCCN")
                    ITEM_CODE = rowCarton.Item("CDITEM")
                    ORDR_QTY_SHIP = Val(rowCarton.Item("WHSE_QTY_SHIP") & String.Empty)
                    processingOverageAcrossMultipleCartons = False
                    Dim CHCTN As String = rowCarton.Item("CHCTNN") & String.Empty

                    While ORDR_QTY_SHIP > 0
                        processingOverageAcrossMultipleCartons = False

                        sql = "ITEM_ALT_SORT = '" & ITEM_CODE & "' and PICK_QTY > 0 and PICK_QTY_CONF < PICK_QTY"
                        If tblSOTPICK2.Select(sql).Length = 0 Then
                            sql = "ITEM_ALT_SORT = '" & ITEM_CODE & "' and PICK_QTY > 0"
                            If tblSOTPICK2.Select(sql).Length = 0 Then
                                ErrorMessages.Add("Details Missing for Pick Ticket: " & ORDR_NO & ", Item Code: " & ITEM_CODE)
                                Exit Function
                            Else
                                processingOverageAcrossMultipleCartons = True
                            End If
                        End If

                        rowSOTPICK2 = tblSOTPICK2.Select(sql)(0)

                        rowEDT945T2 = dst.Tables("EDT945T2").NewRow
                        rowEDT945T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                        EDI_DTL_SEQ += 1
                        rowEDT945T2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                        rowEDT945T2.Item("EDI_CART_NO") = (rowCarton.Item("CPUCCN") & String.Empty).ToString.PadLeft(CART_NO_LENGTH, "0")
                        rowEDT945T2.Item("EDI_SHIPMENT_STATUS_CODE") = "SH"
                        rowEDT945T2.Item("PICK_LNO") = rowSOTPICK2.Item("PICK_LNO")
                        rowEDT945T2.Item("PICK_QTY") = rowSOTPICK2.Item("PICK_QTY")

                        If Not processingOverageAcrossMultipleCartons Then
                            EDI_SHIP_QTY = rowSOTPICK2.Item("PICK_QTY") - rowSOTPICK2.Item("PICK_QTY_CONF")
                            If ORDR_QTY_SHIP <= (rowSOTPICK2.Item("PICK_QTY") - rowSOTPICK2.Item("PICK_QTY_CONF")) Then
                                EDI_SHIP_QTY = ORDR_QTY_SHIP
                            End If
                            ORDR_QTY_SHIP -= EDI_SHIP_QTY
                            rowEDT945T2.Item("EDI_SHIP_QTY") = EDI_SHIP_QTY
                            rowSOTPICK2.Item("PICK_QTY_CONF") += EDI_SHIP_QTY
                        Else
                            EDI_SHIP_QTY = ORDR_QTY_SHIP
                            rowEDT945T2.Item("EDI_SHIP_QTY") = 0
                        End If

                        ' See if this is an over shipment.
                        If ORDR_QTY_SHIP > 0 Then
                            sql = "ITEM_ALT_SORT = '" & ITEM_CODE & "' and PICK_QTY > 0 and PICK_QTY_CONF < PICK_QTY"
                            If tblSOTPICK2.Select(sql).Length = 0 Then
                                rowSOTPICK2.Item("PICK_QTY_CONF") += ORDR_QTY_SHIP
                                rowEDT945T2.Item("EDI_SHIP_QTY") += ORDR_QTY_SHIP
                                ' return a message so it can print on the report.
                                Dim rowError As DataRow = ASCDATA1.GetDataRow("select * from sotordr1 where ordr_no = (select ordr_no from sotpick1 where pick_no = :parm1)", "V", rowSOTPICK2.Item("PICK_NO"))
                                If rowError IsNot Nothing Then
                                    ErrorMessages.Add("WARNING: Sales Order: " & rowError.Item("ORDR_NO") _
                                        & " for Customer: " & rowError.Item("CUST_CODE") & ", PO Number: " & rowError.Item("ORDR_CUST_PO") _
                                        & " has an over shipment of " & ORDR_QTY_SHIP & " pieces for item " & ITEM_CODE & ".")
                                End If
                                ORDR_QTY_SHIP = 0
                            End If
                        End If

                        'rowEDT945T2.Item("EDI_DIFF_QTY") = String.Empty
                        'rowEDT945T2.Item("STYLE_UOM") = String.Empty
                        'rowEDT945T2.Item("UPC_CODE") = String.Empty
                        rowEDT945T2.Item("STYLE_CODE") = rowSOTPICK2.Item("ITEM_CODE")
                        'rowEDT945T2.Item("STYLE_DESC") = String.Empty
                        'rowEDT945T2.Item("UNIT_PRICE") = String.Empty
                        'rowEDT945T2.Item("EDI_STYLE") = String.Empty
                        'rowEDT945T2.Item("EDI_SKU") = String.Empty
                        'rowEDT945T2.Item("EDI_PROD_CLASS") = String.Empty
                        'rowEDT945T2.Item("NMFC") = String.Empty
                        'rowEDT945T2.Item("EDI_STYLE_DESC") = String.Empty
                        'rowEDT945T2.Item("EDI_RETAIL_PRICE") = String.Empty
                        'rowEDT945T2.Item("EDI_SIZE") = String.Empty
                        'rowEDT945T2.Item("EDI_COLOR") = String.Empty
                        'rowEDT945T2.Item("PACK_QTY") = String.Empty
                        'rowEDT945T2.Item("PACK_SIZE") = String.Empty
                        'rowEDT945T2.Item("PACK_UOM") = String.Empty
                        'rowEDT945T2.Item("EDI_ORDR_CUBE") = String.Empty
                        'rowEDT945T2.Item("EDI_CART_LENGTH") = rowCarton.Item("CARTON_LENGTH")
                        'rowEDT945T2.Item("EDI_CART_HEIGHT") = rowCarton.Item("CARTON_HEIGHT")
                        'rowEDT945T2.Item("EDI_CART_WIDTH") = rowCarton.Item("CARTON_WIDTH")
                        'rowEDT945T2.Item("EDI_SHIPPER_ID_NO") = rowCarton.Item("TRACKING_ID")
                        rowEDT945T2.Item("EDI_SHIPPER_ID_NO") = dst.Tables("TRACK").Compute("MAX(PROORTRK)", "[CTNN] = '" & CHCTN & "' AND [INVN] = '" & INVOICE_NUMBER & "'") & String.Empty
                        rowEDT945T2.Item("EDI_CART_WEIGHT") = rowCarton.Item("CHAWGT")
                        dst.Tables("EDT945T2").Rows.Add(rowEDT945T2)
                    End While
                Next
                rowHeader.Item("PROCESS") = ProcessState.Converted
            Next

            ' Back fill EDI_BOL_NO for Cancelled Pick Tickets
            For Each rowEDT945T1 In dst.Tables("EDT945T1").Select("EDI_BOL_NO is Null or EDI_BOL_NO = ''")
                Dim EDI_SHIPMENT_ID As String = (rowEDT945T1.Item("EDI_SHIPMENT_ID") & String.Empty).ToString.Trim

                If EDI_SHIPMENT_ID.Length = 0 Then Continue For

                sql = "EDI_SHIPMENT_ID = '" & EDI_SHIPMENT_ID & "' AND NOT (EDI_BOL_NO is Null or EDI_BOL_NO = '')"
                If dst.Tables("EDT945T1").Select(sql).Length > 0 Then
                    Dim row As DataRow = dst.Tables("EDT945T1").Select(sql)(0)
                    rowEDT945T1.Item("EDI_BOL_NO") = row.Item("EDI_BOL_NO")
                    rowEDT945T1.Item("EDI_MASTER_BOL_NO") = row.Item("EDI_MASTER_BOL_NO")
                End If
            Next

            R = dst.Tables("EDT945T1").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Shipment Confirmation " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try

    End Function

    Public Function ProcessClarinsReturnsTransactions() As Boolean

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "RETURNS"
            environ.APP_KEY = "RTNHDR,RTNDTL,RTNRA"
            clsWHCIMP01 = New WHCIMPO1(environ)

            Dim rowEDTRTRN1 As DataRow = Nothing
            Dim rowEDTRTRN2 As DataRow = Nothing
            Dim rowICTITEM1 As DataRow = Nothing
            Dim EDI_DTL_SEQ As Int32 = 0

            CreateTableAdaptors(ProcessType.Returns)

            If Not dst.Tables.Contains("RTNHDR") Then
                Return False
            End If

            Fill_Records("RTNHDR")
            Fill_Records("RTNDTL")
            Fill_Records("RTNRA")

            For Each rowHeader As DataRow In dst.Tables("RTNHDR").Select("")

                rowHeader.Item("PROCESS") = ProcessState.ProcessReturn

                Dim OrderNumber As String = rowHeader.Item("OHORDN") & String.Empty
                Dim rowRTNRA As DataRow = Nothing
                If dst.Tables("RTNRA").Select("ORDNUMBER = '" & OrderNumber & "'").Length > 0 Then
                    rowRTNRA = dst.Tables("RTNRA").Select("ORDNUMBER = '" & OrderNumber & "'")(0)
                End If

                Dim INVOICE_NO As String = rowHeader.Item("OHINVN")
                rowEDTRTRN1 = dst.Tables("EDTRTRN1").NewRow
                rowEDTRTRN1.Item("EDI_DOC_SEQ_NO") = INVOICE_NO

                If rowRTNRA IsNot Nothing Then
                    Dim EDI_RA_NO As String = rowRTNRA.Item("RMA") & String.Empty
                    EDI_RA_NO = EDI_RA_NO.Trim
                    If IsNumeric(EDI_RA_NO) Then
                        EDI_RA_NO = ASCMAIN1.Format_Field(EDI_RA_NO, "RA_NO")
                    End If
                    rowEDTRTRN1.Item("EDI_RA_NO") = EDI_RA_NO
                    rowEDTRTRN1.Item("CUST_CLAIM_NO") = rowRTNRA.Item("CLAIM") & String.Empty
                End If
                rowEDTRTRN1.Item("EDI_RETURN_DATE") = DateTime.Now.ToShortDateString
                'rowEDTRTRN1.Item("EDI_OP_ID") = ""

                rowEDTRTRN1.Item("EDI_CUSTOMER_NO") = rowHeader.Item("OHCUS1")
                rowEDTRTRN1.Item("EDI_CUST_SHIP_TO") = rowHeader.Item("OHCUS2")

                Dim EDI_CUSTOMER_NO As String = rowEDTRTRN1.Item("EDI_CUSTOMER_NO")
                Dim EDI_CUST_SHIP_TO As String = rowEDTRTRN1.Item("EDI_CUST_SHIP_TO")

                ASCMAIN1.sql = "SELECT * FROM ARTCUST2 WHERE CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'"
                Dim tblCustomer As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST2")

                If tblCustomer.Rows.Count = 0 Then
                    ASCMAIN1.sql = "SELECT * FROM ARTCUST1 WHERE CUST_NO_3PL = '" & EDI_CUSTOMER_NO & "' AND CUST_STORE_NO_3PL = '" & EDI_CUST_SHIP_TO & "'"
                    tblCustomer = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST1")
                End If

                ' Check Store Map Exceptions
                If tblCustomer.Rows.Count = 0 Then
                    ASCMAIN1.sql = "SELECT * FROM TATXREFX WHERE CSCUS1 = '" & EDI_CUSTOMER_NO & "' AND CSCUS2 = '" & EDI_CUST_SHIP_TO & "'"
                    tblCustomer = ASCDATA1.GetDataTable(ASCMAIN1.sql, "TATXREFX")

                    If tblCustomer.Rows.Count > 0 Then
                        Dim CUST_CODE As String = tblCustomer.Rows(0).Item("CUST_CODE")
                        Dim CUST_STORE_NO As String = tblCustomer.Rows(0).Item("CUST_STORE_NO")
                        ASCMAIN1.sql = "SELECT * FROM ARTCUST2 WHERE CUST_CODE = '" & CUST_CODE & "' AND CUST_STORE_NO = '" & CUST_STORE_NO & "'"
                        tblCustomer = ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUST2")
                    End If
                End If

                Dim numRecords As Int16 = tblCustomer.Rows.Count
                Select Case numRecords
                    Case 0
                        ErrorMessages.Add("No Cross Reference for Customer/Ship To (" & EDI_CUSTOMER_NO & "/" & EDI_CUST_SHIP_TO & ") on Return: " & INVOICE_NO)
                        'rowHeader.Item("PROCESS") = ProcessState.InvalidCustomer
                        'Continue For
                    Case 1
                        rowEDTRTRN1.Item("CUST_CODE") = tblCustomer.Rows(0).Item("CUST_CODE") & String.Empty

                    Case Else
                        ErrorMessages.Add("Multiple Cross References for Customer/Ship To (" & EDI_CUSTOMER_NO & "/" & EDI_CUST_SHIP_TO & ") on Return: " & INVOICE_NO)
                        'rowHeader.Item("PROCESS") = ProcessState.InvalidCustomer
                        'Continue For
                End Select

                rowEDTRTRN1.Item("EDI_TOTAL_CARTONS") = Val(rowHeader.Item("OHNCTN") & String.Empty) ' OHNCTN

                'rowEDTRTRN1.Item("EDI_REC_BY_SHIPPER") = String.Empty
                'rowEDTRTRN1.Item("EDI_COMMENTS_1") = String.Empty
                'rowEDTRTRN1.Item("EDI_COMMENTS_2") = String.Empty
                rowEDTRTRN1.Item("EDI_DEPT") = rowHeader.Item("OHDEPT") & String.Empty
                ' rowEDTRTRN1.Item("EDI_POSTING_DATE") = String.Empty

                rowEDTRTRN1.Item("PROCESS_IND") = "0"
                rowEDTRTRN1.Item("INIT_DATE") = DateTime.Now
                rowEDTRTRN1.Item("WHSE_CODE") = "CLARTN"
                dst.Tables("EDTRTRN1").Rows.Add(rowEDTRTRN1)

                Dim MSMSGT As New List(Of String)

                EDI_DTL_SEQ = 0
                For Each rowDetail As DataRow In dst.Tables("RTNDTL").Select("[SAINVN] = '" & INVOICE_NO & "'")
                    ' Detail Record

                    If Val(rowDetail.Item("SAQTYS") & String.Empty) = 0 Then
                        Continue For
                    End If

                    Select Case rowDetail.Item("SACRMC") & String.Empty
                        Case "C48", "C49"
                            'C48 Destroyed In Field (Manual Entry)
                            'C49 Destroyed In Field (EDI)
                        Case "C15"
                            'C15 Return and put in stock
                        Case "C11"
                            'C11 Return and put in refurbish
                        Case "C10"
                            'C10 Return and trashed
                        Case Else
                            Continue For
                    End Select

                    rowEDTRTRN2 = dst.Tables("EDTRTRN2").NewRow
                    rowEDTRTRN2.Item("EDI_DOC_SEQ_NO") = INVOICE_NO

                    ' Clarins is sendind in multiple lines with the SALINE number but the SACRMC is different
                    'rowEDTRTRN2.Item("EDI_DTL_SEQ") = Val(rowDetail.Item("SALINE") & String.Empty)
                    EDI_DTL_SEQ += 10
                    rowEDTRTRN2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ

                    ' Need to Convert Clarins Item Code to ABSolution Code
                    rowEDTRTRN2.Item("EDI_ITEM_CODE") = rowDetail.Item("SAITEM") & String.Empty
                    If rowEDTRTRN2.Item("EDI_ITEM_CODE") & String.Empty <> String.Empty Then
                        rowICTITEM1 = ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_ALT_SORT = :PARM1", "V", New Object() {rowEDTRTRN2.Item("EDI_ITEM_CODE") & String.Empty})
                        If rowICTITEM1 IsNot Nothing Then
                            rowEDTRTRN2.Item("EDI_ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                        End If
                    End If

                    ' Initialize Columns to zero
                    rowEDTRTRN2.Item("EDI_QTY_RETURNED") = 0
                    rowEDTRTRN2.Item("EDI_QTY_BACK_TO_STOCK") = 0
                    rowEDTRTRN2.Item("EDI_QTY_IN_REPAIR") = 0
                    rowEDTRTRN2.Item("EDI_QTY_DAMAGED") = 0
                    rowEDTRTRN2.Item("EDI_QTY_AS_IS") = 0

                    rowEDTRTRN2.Item("EDI_REASON_CODE") = rowDetail.Item("SACRMC")

                    ' Clarins sends negative numbers, convert to absolute value
                    rowEDTRTRN2.Item("EDI_QTY_RETURNED") = Math.Abs(Val(rowDetail.Item("SAQTYS") & String.Empty))

                    'The field called scrap is for scrap orders not items and has never been used. You reported all returns as (STOCK). This is incorrect.
                    'There is a field called SACRMC and it contains the following codes:

                    ' C48 Destroyed In Field (Manual Entry)
                    ' C49 Destroyed In Field (EDI)
                    ' C11 Return and put in refurbish
                    ' C10 Return and trashed

                    ' This code provided by someone in Clarins Warehoue - they spoke to Walter and Lauren
                    ' C15 Return and put in stock

                    Select Case rowDetail.Item("SACRMC") & String.Empty
                        Case "C48", "C49"
                            'C48 Destroyed In Field (Manual Entry)
                            'C49 Destroyed In Field (EDI)
                            rowEDTRTRN2.Item("EDI_QTY_DAMAGED") = rowEDTRTRN2.Item("EDI_QTY_RETURNED")
                        Case "C15"
                            'C15 Return and put in stock
                            rowEDTRTRN2.Item("EDI_QTY_BACK_TO_STOCK") = rowEDTRTRN2.Item("EDI_QTY_RETURNED")
                        Case "C11"
                            'C11 Return and put in Refurbish
                            rowEDTRTRN2.Item("EDI_QTY_IN_REPAIR") = rowEDTRTRN2.Item("EDI_QTY_RETURNED")
                        Case "C10"
                            'C10 Return and trashed
                            rowEDTRTRN2.Item("EDI_QTY_DAMAGED") = rowEDTRTRN2.Item("EDI_QTY_RETURNED")
                    End Select

                    rowEDTRTRN2.Item("EDI_COMMENTS") = rowDetail.Item("MSMSGT") & String.Empty
                    dst.Tables("EDTRTRN2").Rows.Add(rowEDTRTRN2)
                Next
                rowHeader.Item("PROCESS") = ProcessState.Converted
            Next
            R = dst.Tables("EDTRTRN1").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Returns Transactions: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsAdjustments() As Boolean

        Dim errorLocation As String = String.Empty

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "INVTRANSADJ"
            clsWHCIMP01 = New WHCIMPO1(environ)
            CreateTableAdaptors(ProcessType.InventoryAdjustments)

            If Not dst.Tables.Contains("INVTRANS") Then
                Return True
            End If

            ASCDATA1.ExecuteSQL("UPDATE CONV.CFG_INVTRANSADJ SET PROC_KEY = to_char(systimestamp, 'yyyymmddhh24miss') || '_' || ROWNUM WHERE PROC_KEY IS NULL")
            Fill_Records("INVTRANS")
            Fill_Records("INVTRANSA")

            If dst.Tables("INVTRANS").Rows.Count = 0 AndAlso dst.Tables("INVTRANSA").Rows.Count = 0 Then
                Return True
            End If

            ' Do not process Type R's, only Type A's
            ' Within the A's, do not process reason codes A03 and A21, these are Make items and components, these are handled
            '   in a different process
            Dim TRX_NO As String = String.Empty
            If dst.Tables("INVTRANS").Rows.Count > 0 Then
                TRX_NO = ASCMAIN1.Next_Control_No("EDTTRXN1.TRX_NO")
            End If
            Dim TRX_LNO As Int32 = 0
            Dim TRANS_DATE As String = String.Empty
            Dim rowICTITEM1 As DataRow = Nothing

            Dim tblICTWHSE1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM ICTWHSE1")

            For Each rowINVTRANS As DataRow In dst.Tables("INVTRANS").Select("DTTTYP = 'A' AND ISNULL(DTADJC, '*') <> 'A03' AND ISNULL(DTADJC, '*') <> 'A21'")

                Dim rowEDTTRXN1 As DataRow = dst.Tables("EDTTRXN1").NewRow
                rowEDTTRXN1.Item("TRX_NO") = TRX_NO
                TRX_LNO += 1
                rowEDTTRXN1.Item("TRX_LNO") = TRX_LNO

                Select Case rowINVTRANS.Item("DTTTYP") & String.Empty

                    Case "A"
                        rowEDTTRXN1.Item("TRANS_TYPE") = "ADJ"
                    Case "R"
                        rowEDTTRXN1.Item("TRANS_TYPE") = "REC"
                End Select

                'DTASLE, DTROWN, DTBINN
                Dim TRANS_NUM As String = rowINVTRANS.Item("DTASLE") & "_" & rowINVTRANS.Item("DTROWN") & "_" & rowINVTRANS.Item("DTBINN") & String.Empty
                rowEDTTRXN1.Item("TRANS_NUM") = TRANS_NUM

                TRANS_DATE = rowINVTRANS.Item("DTDATE") & String.Empty
                If TRANS_DATE.Length = 6 Then
                    TRANS_DATE = TRANS_DATE.Substring(2, 2) & "/" & TRANS_DATE.Substring(4, 2) & "/" & TRANS_DATE.Substring(0, 2)
                Else
                    TRANS_DATE = DateTime.Now.ToString("MM/dd/yy")
                End If

                If IsDate(TRANS_DATE) Then
                    rowEDTTRXN1.Item("TRANS_DATE") = TRANS_DATE
                End If

                'rowEDTTRXN1.Item("PO_ORDER_NO") = String.Empty
                'rowEDTTRXN1.Item("BUYER") = String.Empty

                rowEDTTRXN1.Item("REASON_CODE") = rowINVTRANS.Item("DTADJC") & String.Empty
                rowEDTTRXN1.Item("OPERATOR") = rowINVTRANS.Item("DTUSER") & String.Empty
                rowEDTTRXN1.Item("ITEM_CODE") = rowINVTRANS.Item("DTITEM") & String.Empty

                rowICTITEM1 = ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_ALT_SORT = :PARM1", "V", New Object() {rowEDTTRXN1.Item("ITEM_CODE") & String.Empty})
                If rowICTITEM1 IsNot Nothing Then
                    rowEDTTRXN1.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE") & String.Empty
                    If rowEDTTRXN1.Table.Columns.Contains("ITEM_DESC") Then
                        rowEDTTRXN1.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                    End If
                End If

                rowEDTTRXN1.Item("LOCATION") = rowINVTRANS.Item("DTWHSI") & String.Empty
                rowEDTTRXN1.Item("TRAN_QTY") = Val(rowINVTRANS.Item("DTTQTY") & String.Empty)
                rowEDTTRXN1.Item("PROCESS_IND") = "0"

                If tblICTWHSE1.Select("LP_WHSE_ID = '" & rowEDTTRXN1.Item("LOCATION") & "'").Length > 0 Then
                    rowEDTTRXN1.Item("WHSE_CODE") = tblICTWHSE1.Select("LP_WHSE_ID = '" & rowEDTTRXN1.Item("LOCATION") & "'")(0).Item("WHSE_CODE")
                Else
                    rowEDTTRXN1.Item("WHSE_CODE") = rowEDTTRXN1.Item("LOCATION")
                End If

                If ASCMAIN1.DBS_COMPANY = "INT" OrElse ASCMAIN1.DBS_SERVER = "INT" Then
                    If rowEDTTRXN1.Item("WHSE_CODE") & String.Empty = "RNG" Then
                        rowEDTTRXN1.Item("WHSE_CODE") = "CLA"
                    End If

                    If rowEDTTRXN1.Item("WHSE_CODE") & String.Empty = "RTN" Then
                        rowEDTTRXN1.Item("WHSE_CODE") = "CLARTN"
                    End If
                End If

                rowEDTTRXN1.Item("IMPORT_DATE") = CDate(DateTime.Now.ToShortDateString)

                dst.Tables("EDTTRXN1").Rows.Add(rowEDTTRXN1)
            Next

            For Each rowINVTRANSA As DataRow In dst.Tables("INVTRANSA").Select("NOT (DTTTYP = 'A' AND ISNULL(DTADJC, '*') <> 'A03' AND ISNULL(DTADJC, '*') <> 'A21')")
                Dim rowEDTTRXNA As DataRow = dst.Tables("EDTTRXNA").NewRow

                For Each col As DataColumn In dst.Tables("EDTTRXNA").Columns
                    If dst.Tables("INVTRANSA").Columns.Contains(col.ColumnName) Then
                        rowEDTTRXNA.Item(col.ColumnName) = rowINVTRANSA.Item(col.ColumnName)
                    End If
                Next

                dst.Tables("EDTTRXNA").Rows.Add(rowEDTTRXNA)
            Next

            R = dst.Tables("EDTTRXN1").Rows.Count + dst.Tables("EDTTRXNA").Rows.Count

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Adjustments: " & ex.Message)
            Return False

        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsOpenOrders() As Boolean

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "OPNORDHED,OPNORDDTL"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.OpenOrders)

            If Not dst.Tables.Contains("OPNORDHED") Then
                Exit Function
            End If

            Fill_Records("OPNORDHED")
            Fill_Records("OPNORDDTL")

            If dst.Tables("OPNORDHED").Rows.Count = 0 Then
                Return True
            End If

            Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT * FROM SOTPICKC WHERE ORDR_NO_3PL IN (SELECT OHORDN FROM CONV.CFG_OPNORDHED)", "SOTPICKC")

            For Each rowOPNORDHED As DataRow In dst.Tables("OPNORDHED").Select()

                Dim ORDR_NO_3PL As String = rowOPNORDHED.Item("OHORDN") & String.Empty

                Dim rowSOTPICKO As DataRow = dst.Tables("SOTPICKO").NewRow
                rowSOTPICKO.Item("ORDR_NO_3PL") = ORDR_NO_3PL

                If tbl.Select("ORDR_NO_3PL = '" & ORDR_NO_3PL & "'").Length > 0 Then
                    Dim row As DataRow = tbl.Select("ORDR_NO_3PL = '" & ORDR_NO_3PL & "'")(0)
                    rowSOTPICKO.Item("PICK_NO") = row.Item("PICK_NO")
                    rowSOTPICKO.Item("ORDR_NO") = row.Item("ORDR_NO")
                End If

                dst.Tables("SOTPICKO").Rows.Add(rowSOTPICKO)
            Next

            R = dst.Tables("SOTPICKO").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Open Orders: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsReceipts() As Boolean

        Try
            R = 0
            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "RECEIPTSADJ"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.Receipts)

            If Not dst.Tables.Contains("RECEIPTS") Then
                Return True
            End If

            Fill_Records("RECEIPTS")
            If dst.Tables("RECEIPTS").Rows.Count = 0 Then
                Return True
            End If

            R = dst.Tables("RECEIPTS").Rows.Count

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Receipts: " & ex.Message)
            Return False

        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsNewCustomers() As Boolean

        Try
            R = 0

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "CUSTO"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.NewCustomers)

            If Not dst.Tables.Contains("CUSTO") Then
                Return True
            End If

            Fill_Records("CUSTO")
            R = dst.Tables("CUSTO").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins New Customers: " & ex.Message)
            Return False

        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try

    End Function

    Public Function ProcessClarinsItemMaster() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "ITEMMAST"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.ItemMaster)

            If Not dst.Tables.Contains("ITEMMAST") Then
                Return True
            End If

            Fill_Records("ITEMMAST")
            R = dst.Tables("ITEMMAST").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Item Master: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsShipViaScac() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "SHPVIASCAC"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.ShipViaScac)

            If Not dst.Tables.Contains("SHPVIASCAC") Then
                Return True
            End If

            Fill_Records("SHPVIASCAC")
            R = dst.Tables("SHPVIASCAC").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Ship Via Scac: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsBrands() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "BRANDIPLB"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.Brands)

            If Not dst.Tables.Contains("BRAND") Then
                Return True
            End If

            Fill_Records("BRAND")
            R = dst.Tables("BRAND").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Brands: " & ex.Message)
            Return False

        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsSubBrands() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "SUBBRANDIPLB"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.SubBrands)

            If Not dst.Tables.Contains("SUBBRAND") Then
                Return True
            End If

            Fill_Records("SUBBRAND")
            R = dst.Tables("SUBBRAND").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Sub Brands: " & ex.Message)
            Return False

        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsInventory() As Boolean

        Try

            ' Clean out exisitng files
            ASCDATA1.ExecuteSQL("DELETE FROM CONV.CFG_INVENTORYIPLB")

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "INVENTORYIPLB"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.Inventory)

            If Not dst.Tables.Contains("INVENTORY") Then
                Return True
            End If

            ' Currently we do not do anything with this data
            Fill_Records("INVENTORY", String.Empty, True, "Select * from CONV.CFG_INVENTORYIPLB WHERE ROWNUM < 1")
            R = dst.Tables("INVENTORY").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Inventory: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsInvoice() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "INVOICEIPLB"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.Invoice)

            If Not dst.Tables.Contains("INVOICE") Then
                Return True
            End If

            Fill_Records("INVOICE")
            R = dst.Tables("INVOICE").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Invoice: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsShipRule() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "SHIPRULEIPLB"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.ShipRule)

            If Not dst.Tables.Contains("SHIPRULE") Then
                Return True
            End If

            Fill_Records("SHIPRULE")
            R = dst.Tables("SHIPRULE").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Ship Rule: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsVendors() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "VENDORS"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.Vendors)

            If Not dst.Tables.Contains("VENDORS") Then
                Return True
            End If

            Fill_Records("VENDORS")
            R = dst.Tables("VENDORS").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Vendors: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsCustomers() As Boolean

        Try

            environ = New ABSEnvironment
            environ.DBS_COMPANY = G.DBS_COMPANY
            environ.DBS_SERVER = G.DBS_SERVER
            environ.DBS_PASSWORD = G.DBS_PASSWORD
            environ.CLIENT = G.CLIENT
            environ.THREAD_NO = 0
            environ.APP_ID = ""
            environ.APP_DESC = ""
            environ.USER_ID = ASCMAIN1.USER_ID
            environ.APP_CMD = "SHIP"
            environ.APP_KEY = "CUSTMASTIPLB"
            clsWHCIMP01 = New WHCIMPO1(environ)

            CreateTableAdaptors(ProcessType.CustomerMaster)

            If Not dst.Tables.Contains("CUSTOMERS") Then
                Return True
            End If

            'Fill_Records("CUSTOMERS")
            R = dst.Tables("CUSTOMERS").Rows.Count
            Return True

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Customers: " & ex.Message)
            Return False
        Finally
            environ = Nothing
            clsWHCIMP01 = Nothing
        End Try
    End Function

    Public Function ProcessClarinsItemCaseCount() As Boolean
        'File	    External Field Name	    Input Buffer Position    Field Length In Bytes	Number Of Digits	Decimal Positions	Field Text Description
        'IPCASEQTY	IPLBITEM	            1	                    15	                    0	                0	                IPLB_Item_Number
        'IPCASEQTY	CASEQTY	                16	                    9	                    9	                0	                IPLB_Case_Quantity

        Try
            Dim sql As String = String.Empty

            sql = "Select ITEM_ALT_SORT, NVL(CARTON_PACK_QTY, 1) CARTON_PACK_QTY From ICTITEM1 Where NVL(ITEM_STATUS, 'A') = 'A' And ITEM_ALT_SORT IS NOT Null And NVL(HIDE_FROM_3PL, '0') = '0'"
            Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)

            'IPCAS_yyyymmddhhmmss.CSV
            Dim fileName As String = "IPCAS_" & DateTime.Now.ToString("yyyyMMddhhmmss") & ".CSV"
            Dim subDir As String = DateTime.Now.ToString("yyyyMM")

            Dim dataDirectory As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\" _
                    & IIf(G.DBS_SERVER = "TST", "TEST", "PROD") & "\FROM_IPLB\"

            Using sw As New IO.StreamWriter(dataDirectory & fileName, False)
                For Each rowData As DataRow In tbl.Select("", "ITEM_ALT_SORT")
                    Dim ITEM_ALT_SORT As String = rowData.Item("ITEM_ALT_SORT") & String.Empty
                    Dim CARTON_PACK_QTY As String = Val(rowData.Item("CARTON_PACK_QTY") & String.Empty).ToString

                    sw.WriteLine(ITEM_ALT_SORT & "," & CARTON_PACK_QTY)
                Next
                sw.Close()

            End Using

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins Item Case Count: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function ProcessClarinsRetnsinvAbscnts() As Boolean
        Try

            Dim dataDirectory As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT") & "\cfg\" & IIf(G.DBS_SERVER = "TST", "TEST", "PROD") & "\FROM_CUSA\"
            Dim subDir As String = DateTime.Now.ToString("yyyyMM")
            Dim dataDirectoryArchive As String = dataDirectory & "Archive\" & subDir & "\"

            For Each dataFileName As String In My.Computer.FileSystem.GetFiles(dataDirectory, FileIO.SearchOption.SearchTopLevelOnly, "retnsinv*.*")
                Dim fileName As String = My.Computer.FileSystem.GetName(dataFileName)
                My.Computer.FileSystem.MoveFile(dataDirectory & fileName, dataDirectoryArchive & fileName, True)
            Next

            For Each dataFileName As String In My.Computer.FileSystem.GetFiles(dataDirectory, FileIO.SearchOption.SearchTopLevelOnly, "abscnts*.*")
                Dim fileName As String = My.Computer.FileSystem.GetName(dataFileName)
                My.Computer.FileSystem.MoveFile(dataDirectory & fileName, dataDirectoryArchive & fileName, True)
            Next

        Catch ex As Exception
            ErrorMessages.Add("Process Clarins retnsinv/abscnts File: " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub CreateTableAdaptors(ByVal pType As ProcessType)
        Try
            With dst

                If pType = ProcessType.ShipmentConfirmation OrElse pType = ProcessType.Initialize Then
                    Dim rowSHIPHDR As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_SHIPHDR'")

                    If rowSHIPHDR IsNot Nothing Then
                        If Not .Tables.Contains("SHIPHDR") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_SHIPHDR"
                            Create_TDA(.Tables.Add("SHIPHDR"), "CONV.CFG_SHIPHDR", ASCMAIN1.sql, 0, True)
                            dst.Tables("SHIPHDR").Columns.Add("PROCESS", GetType(System.Int16))
                        End If

                        If Not .Tables.Contains("SHIPDTL") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_SHIPDTL"
                            Create_TDA(.Tables.Add("SHIPDTL"), "CONV.CFG_SHIPDTL", ASCMAIN1.sql, 0, True)
                        End If

                        If Not .Tables.Contains("CARTON") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_CARTON"
                            Create_TDA(.Tables.Add("CARTON"), "CONV.CFG_CARTON", ASCMAIN1.sql, 0, True)
                            .Tables("CARTON").Columns.Add("WHSE_QTY_SHIP", GetType(System.Int32), "CDQTYS * CDSKMQ")
                        End If

                        If Not .Tables.Contains("TRACK") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_TRACK"
                            Create_TDA(.Tables.Add("TRACK"), "CONV.CFG_TRACK", ASCMAIN1.sql, 0, True)
                        End If
                    End If

                End If

                If pType = ProcessType.Returns OrElse pType = ProcessType.Initialize Then
                    Dim rowRTNHDR As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_RTNHDR'")

                    If rowRTNHDR IsNot Nothing Then
                        If Not .Tables.Contains("RTNHDR") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_RTNHDR"
                            Create_TDA(.Tables.Add("RTNHDR"), "CONV.CFG_RTNHDR", ASCMAIN1.sql, 0, True)
                            dst.Tables("RTNHDR").Columns.Add("PROCESS", GetType(System.Int16))
                        End If

                        If Not .Tables.Contains("RTNDTL") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_RTNDTL"
                            Create_TDA(.Tables.Add("RTNDTL"), "CONV.CFG_RTNDTL", ASCMAIN1.sql, 0, True)
                        End If

                        If Not .Tables.Contains("RTNRA") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_RTNRA"
                            Create_TDA(.Tables.Add("RTNRA"), "CONV.CFG_RTNRA", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.OrderConfirmation OrElse pType = ProcessType.Initialize Then
                    Dim rowABSOH As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_ABSOH'")

                    If rowABSOH IsNot Nothing Then
                        If Not .Tables.Contains("ABSOH") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_ABSOH"
                            Create_TDA(.Tables.Add("ABSOH"), "CONV.CFG_ABSOH", ASCMAIN1.sql, 0, True)
                        End If

                        If Not .Tables.Contains("ABSOD") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_ABSOD"
                            Create_TDA(.Tables.Add("ABSOD"), "CONV.CFG_ABSOD", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.OpenOrders OrElse pType = ProcessType.Initialize Then
                    Dim rowOPNORDHED As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_OPNORDHED'")

                    If rowOPNORDHED IsNot Nothing Then
                        If Not .Tables.Contains("OPNORDHED") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_OPNORDHED"
                            Create_TDA(.Tables.Add("OPNORDHED"), "CONV.CFG_OPNORDHED", ASCMAIN1.sql, 0, True)
                        End If

                        If Not .Tables.Contains("OPNORDDTL") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_OPNORDDTL"
                            Create_TDA(.Tables.Add("OPNORDDTL"), "CONV.CFG_OPNORDDTL", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.CancelledOrders OrElse pType = ProcessType.Initialize Then
                    Dim rowCNLORDS As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_CNLORDS'")

                    If rowCNLORDS IsNot Nothing Then
                        If Not .Tables.Contains("CNLORDS") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_CNLORDS"
                            Create_TDA(.Tables.Add("CNLORDS"), "CONV.CFG_CNLORDS", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.InventoryAdjustments OrElse pType = ProcessType.Initialize Then
                    Dim rowINVTRANS As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_INVTRANSADJ'")

                    If rowINVTRANS IsNot Nothing Then
                        ASCMAIN1.sql = "Select * from CONV.CFG_INVTRANSADJ WHERE ROWNUM < 1"
                        Dim tmpTable As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

                        Try
                            If Not tmpTable.Columns.Contains("PROC_KEY") Then
                                ASCDATA1.ExecuteSQL("ALTER TABLE CONV.CFG_INVTRANSADJ ADD PROC_KEY VARCHAR2(25)")
                                ASCDATA1.ExecuteSQL("ALTER TABLE CONV.CFG_INVTRANSADJ ADD PROCESSED_DATE DATE")
                                ASCDATA1.ExecuteSQL("ALTER TABLE CONV.CFG_INVTRANSADJ ADD PROCESSED_IND VARCHAR2(1)")
                            End If
                        Catch ex As Exception

                        End Try

                        If Not .Tables.Contains("INVTRANS") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_INVTRANSADJ WHERE (DTTTYP = 'A' AND NVL(DTADJC, '*') NOT IN ('A03', 'A21')) AND NVL(PROCESSED_IND, '0') = '0'"
                            Create_TDA(.Tables.Add("INVTRANS"), "CONV.CFG_INVTRANSADJ", ASCMAIN1.sql, 0, True)
                        End If

                        If Not .Tables.Contains("INVTRANSA") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_INVTRANSADJ WHERE NOT (DTTTYP = 'A' AND NVL(DTADJC, '*') NOT IN ('A03', 'A21')) AND NVL(PROCESSED_IND, '0') = '0'"
                            Create_TDA(.Tables.Add("INVTRANSA"), "CONV.CFG_INVTRANSADJ", ASCMAIN1.sql, 0, True)
                        End If

                    End If
                End If

                If pType = ProcessType.Receipts OrElse pType = ProcessType.Initialize Then
                    Dim rowINVTRANS As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_RECEIPTSADJ'")

                    If rowINVTRANS IsNot Nothing Then
                        If Not .Tables.Contains("RECEIPTS") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_RECEIPTSADJ"
                            Create_TDA(.Tables.Add("RECEIPTS"), "CONV.CFG_RECEIPTSADJ", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.NewCustomers OrElse pType = ProcessType.Initialize Then
                    Dim rowCUSTO As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_CUSTO'")

                    If rowCUSTO IsNot Nothing Then
                        If Not .Tables.Contains("CUSTO") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_CUSTO WHERE TRNSTATUS <> 'E'"
                            Create_TDA(.Tables.Add("CUSTO"), "CONV.CFG_CUSTO", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.ShipViaScac OrElse pType = ProcessType.Initialize Then
                    Dim rowSHPVIASCAC As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_SHPVIASCAC'")

                    If rowSHPVIASCAC IsNot Nothing Then
                        If Not .Tables.Contains("SHPVIASCAC") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_SHPVIASCAC"
                            Create_TDA(.Tables.Add("SHPVIASCAC"), "CONV.CFG_SHPVIASCAC", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.ItemMaster OrElse pType = ProcessType.Initialize Then
                    Dim rowITEMMAST As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_ITEMMAST'")

                    If rowITEMMAST IsNot Nothing Then
                        If Not .Tables.Contains("ITEMMAST") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_ITEMMAST"
                            Create_TDA(.Tables.Add("ITEMMAST"), "CONV.CFG_ITEMMAST", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.Brands OrElse pType = ProcessType.Initialize Then
                    Dim rowBRAND As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_BRANDIPLB'")

                    If rowBRAND IsNot Nothing Then
                        If Not .Tables.Contains("BRAND") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_BRANDIPLB"
                            Create_TDA(.Tables.Add("BRAND"), "CONV.CFG_BRANDIPLB", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.SubBrands OrElse pType = ProcessType.Initialize Then
                    Dim rowSUBBRAND As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_SUBBRANDIPLB'")

                    If rowSUBBRAND IsNot Nothing Then
                        If Not .Tables.Contains("SUBBRAND") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_SUBBRANDIPLB"
                            Create_TDA(.Tables.Add("SUBBRAND"), "CONV.CFG_SUBBRANDIPLB", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.Inventory OrElse pType = ProcessType.Initialize Then
                    Dim rowINVENTORY As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_INVENTORYIPLB'")

                    If rowINVENTORY IsNot Nothing Then
                        If Not .Tables.Contains("INVENTORY") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_INVENTORYIPLB"
                            Create_TDA(.Tables.Add("INVENTORY"), "CONV.CFG_INVENTORYIPLB", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.Invoice OrElse pType = ProcessType.Initialize Then
                    Dim rowINVOICE As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_INVOICEIPLB'")

                    If rowINVOICE IsNot Nothing Then
                        If Not .Tables.Contains("INVOICE") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_INVOICEIPLB"
                            Create_TDA(.Tables.Add("INVOICE"), "CONV.CFG_INVOICEIPLB", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.ShipRule OrElse pType = ProcessType.Initialize Then
                    Dim rowSHIPRULE As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_SHIPRULEIPLB'")

                    If rowSHIPRULE IsNot Nothing Then
                        If Not .Tables.Contains("SHIPRULE") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_SHIPRULEIPLB"
                            Create_TDA(.Tables.Add("SHIPRULE"), "CONV.CFG_SHIPRULEIPLB", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.Vendors OrElse pType = ProcessType.Initialize Then
                    Dim rowBRAND As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_VENDORS'")

                    If rowBRAND IsNot Nothing Then
                        If Not .Tables.Contains("VENDORS") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_VENDORS"
                            Create_TDA(.Tables.Add("VENDORS"), "CONV.CFG_VENDORS", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If

                If pType = ProcessType.CustomerMaster OrElse pType = ProcessType.Initialize Then
                    Dim rowBRAND As DataRow = ASCDATA1.GetDataRow("Select * from dba_tables where owner = 'CONV' and TABLE_NAME = 'CFG_CUSTMASTIPLB'")

                    If rowBRAND IsNot Nothing Then
                        If Not .Tables.Contains("CUSTOMERS") Then
                            ASCMAIN1.sql = "Select * from CONV.CFG_CUSTMASTIPLB"
                            Create_TDA(.Tables.Add("CUSTOMERS"), "CONV.CFG_CUSTMASTIPLB", ASCMAIN1.sql, 0, True)
                        End If
                    End If
                End If


            End With

        Catch ex As Exception
            If pType <> ProcessType.Initialize Then
                Throw New Exception(ex.Message)
            End If
        End Try
    End Sub

    Private Sub DisposeOPD()
        Try
            If clsASCBASE1 Is Nothing Then
                Exit Sub
            End If

            With clsASCBASE1

                If .CMDs IsNot Nothing AndAlso .CMDs.Count <> 0 Then
                    For Each CMD_key As String In .CMDs.Keys
                        Dim cmd As OracleCommand = .CMDs(CMD_key)
                        For Each param As OracleParameter In cmd.Parameters
                            param.Dispose()
                        Next
                        cmd.Dispose()
                    Next
                End If
                .CMDs = Nothing

                If .BA_CMDs IsNot Nothing AndAlso .BA_CMDs.Count <> 0 Then
                    For Each CMD_key As String In .BA_CMDs.Keys
                        Dim cmds() As OracleCommand = .BA_CMDs(CMD_key)
                        For Each cmd As OracleCommand In cmds
                            For Each param As OracleParameter In cmd.Parameters
                                param.Dispose()
                            Next
                            cmd.Dispose()
                        Next
                        cmds = Nothing
                    Next
                End If
                .BA_CMDs = Nothing

                If .TDAs IsNot Nothing Then
                    For Each tda As OracleDataAdapter In .TDAs.Values
                        tda.Dispose()
                    Next
                End If
                .TDAs = Nothing

                .Dispose()
            End With

            environ = Nothing
            clsWHCIMP01 = Nothing

        Catch ex As Exception

        End Try
    End Sub

End Class