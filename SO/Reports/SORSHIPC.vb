Public Class SORSHIPC

#Region "General Declarations"

    Private sqlWHT3PLS1 As String = String.Empty
    Private tblWHT3PLS1 As String = String.Empty
    Private SEQ_NO As Int16 = 0
    Private tblSOTSVIA1 As DataTable = Nothing

    Private sqlSOTPICK1 As String = String.Empty
    Private sqlSOTPICK2 As String = String.Empty
    Private sqlSOTSHIPX As String = String.Empty

    Private CURR_CODE As String = String.Empty
    Private CURR_EXCH_RATE As Decimal = 0

    Private CreditCardProcessor As TAC.TAFCARDF
    Private validDates() As Date = TAC.SOCMAIN1.Validate_Invoice_Date(Nothing, 0, 1, Nothing)
    Private clsSOCASDO1 As TAC.SOCADSO1
    Private filesProcessed As List(Of String)
    Private Company_Code As String = String.Empty

    Private tblEDT945T1 As String = String.Empty
    Private tblEDT945T2 As String = String.Empty

    Private tblSOTCART1 As String = String.Empty
    Private tblSOTCART2 As String = String.Empty
    ' SR-6549 - Lot Numbers on Shipments and Receipts
    Private tblSOTCART3 As String = String.Empty
    Private bolList As New List(Of String)

    Private tblSOTSHIP_CART As DataTable

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        MyBase.txtDescription.Enabled = False

        'If ASCMAIN1.Running_in_VS Then
        '    ASCMAIN1.absTimer.Enabled = True
        'End If


        If ASCMAIN1.DBS_COMPANY = "AHA" OrElse ASCMAIN1.DBS_SERVER = "AHA" Then
            Company_Code = "AHA"
        ElseIf ASCMAIN1.DBS_COMPANY = "INT" OrElse ASCMAIN1.DBS_SERVER = "INT" Then
            Company_Code = "INT"
        End If

        ' Did this so the Excel Workbook has the font as Black not a disabled look.
        grdSOTCART1.Parent = Me
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        Dim continueProcessing As Boolean = True
        SEQ_NO = 0
        RWU = "R"

        clsSOCASDO1 = New TAC.SOCADSO1
        Prepare_dst(False, sqlw)

        ' Added 3/7/2018
        If ASCMAIN1.CLIENT = "INT" Then
            If Not ASCMAIN1.Logical_Lock("CLARINS", "IMPORT") Then
                RWU = 0
                Exit Sub
            End If

            If Not ASCMAIN1.Logical_Lock("F", "EDF81056") Then
                RWU = 0
                Exit Sub
            End If
        End If

        GetShipmentData()

        Fill_Records_RPT()
        filesProcessed = New List(Of String)

        If dst.Tables("EDT945T1").Rows.Count > 0 Then
            ToggleDataTableExpressions(True)

            continueProcessing = ValidateShipments()

            If continueProcessing Then
                continueProcessing = dst.Tables("EDT945T1").Rows.Count > 0
            End If

            If continueProcessing Then
                continueProcessing = Load_Shipments()
            End If

            If continueProcessing Then
                continueProcessing = ProcessShipments()
            End If

            If continueProcessing Then
                continueProcessing = FillQuantityShipped()
            End If

            If continueProcessing Then
                continueProcessing = EvaluateLotNos()
            End If

            If Not continueProcessing Then
                Clear_Records()
            End If
        Else
            AddError("No shipments to Process. Try again later.")
        End If

        Try

            ' Send emails in Ahava Production Environment Only
            Dim production As Boolean = False

            If ASCMAIN1.DBS_COMPANY = "AHA" AndAlso ASCMAIN1.DBS_SERVER = "AHA" Then
                production = True
            End If

            If production Then
                Dim note As String = String.Empty
                If dst.Tables("EDTTRXN1").Select("TRANS_TYPE = 'REC'").Length > 0 Then
                    note &= Environment.NewLine & dst.Tables("EDTTRXN1").Select("TRANS_TYPE = 'REC'").Length & " Receipts imported from ADS."
                End If

                If dst.Tables("EDTTRXN1").Select("TRANS_TYPE = 'ADJ'").Length > 0 Then
                    note &= Environment.NewLine & dst.Tables("EDTTRXN1").Select("TRANS_TYPE = 'ADJ'").Length & " Adjustments imported from ADS."
                End If

                If dst.Tables("EDTRTRN1").Select("").Length > 0 Then
                    note &= Environment.NewLine & dst.Tables("EDTRTRN1").Select("").Length & " Returns imported from ADS."
                End If

                If note.Length > 0 Then
                    Dim objASCNOTE1 As New TAC.ASCNOTE1("CONF_ADS", Nothing)
                    objASCNOTE1.Note = note
                    objASCNOTE1.CreateComponents()
                    objASCNOTE1.EmailDocument()
                End If
            End If

        Catch ex As Exception

        End Try

        Try
            GetNoShipShipments()
        Catch ex As Exception
            AddError("Error calling GetNoShipShipments", "1")
        End Try

        Try

        Catch ex As Exception
            AddError("Error calling EvaluateLotNos", "1")
        End Try

        'Check_if_Empty("EDT945T1")

        If dst.Tables("ERRORS").Select("FYI = '0'").Length > 0 Then
            ' Disable Update Menu Option if there are errors
            RWU &= "0"
        ElseIf dst.Tables("ERRORS").Select("FYI = '1'").Length > 0 Then
            RWU = "R"
        Else
            AddError("No Errors", "1")
            RWU = "R"
        End If

    End Sub

    Public Overrides Sub Print_Report()

        Dim reports As New Dictionary(Of String, String)
        reports.Add("SOR3PLER", "3PL Processing Error Messages")
        reports.Add("SORSH3PL", "3PL Shipments Confirmation")

        If Company_Code = "INT" Then
            reports.Add("SORCN3PL", "3PL Canceled Quantities")
        End If

        reports.Add("SOR3PLIN", "3PL Inventory Shipped")
        reports.Add("EDRTRXN1", "3PL Inventory Receipts-Adjustments")
        reports.Add("EDRRTRN1", "3PL Customer Returns")

        If dst.Tables("LOT_NOS").Rows.Count > 0 Then
            reports.Add("SOR3PLLT", "Missing Lot Numbers")
        End If

        If Company_Code = "AHA" Then
            reports.Add("ARRCCPAC", "3PL Credit Card Processing")
        ElseIf Company_Code = "INT" Then
            reports.Add("SOR3PLNS", "Confirmed but not Invoiced")
        End If
        SUBT = String.Empty

        Dim report As KeyValuePair(Of String, String)
        For Each report In reports
            RPT = report.Key
            If RPT = "SORSH3PL" Then
                CR_params.Add("INV_TOTAL_AMOUNT_SUM", Val(dst.Tables("SOTINVH1").Compute("SUM(INV_TOTAL_AMOUNT)", "") & String.Empty))
            End If
            Generate_Report(RPT, report.Value, SUBT)
        Next


    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Done"
        End Select

    End Sub

    Private Sub Clear_Records()

        EnforceConstraints(False)
        ' SR-6549 - Lot Numbers on Shipments and Receipts
        For Each tableName As String In New String() {"EDT945T1", "EDT945T2",
                                                "SOTSHIP0", "SOTSHIP1", "SOTSHIPB", "SOTSHIP3", "SOTSHIP4", "SOTSHIP6",
                                                "SOTPICK1", "SOTPICK2",
                                                "SOTORDR1", "SOTORDR2", "SOTORDR5", "TATEVNT1",
                                                "SOTCART1", "SOTCART2", "SOTCART3", "SOTCARTX",
                                                "ARTCCPA1", "ARTCCPA2", "ARTCCPDA", "ARTCCPAC",
                                                "ICTSTAT2", "ICTWHSE1",
                                                "SOTINVH1", "SOTINVH2", "ARTOPEN1",
                                                "SOTORDR1_CANC", "SOTORDR2_CANC", "SOTPICK2_CANC", "SOTORDR1_NOSHIP", "LOT_NOS"}
            dst.Tables(tableName).Rows.Clear()
        Next
        bolList.Clear()
        EnforceConstraints(True)
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()
        Dim sqlw As String = String.Empty

        With dst
            sqlWHT3PLS1 = "Select SOTSHIP1.SHIP_BOL_NO, EDT945T1.EDI_DOC_SEQ_NO" _
                & ", EDT945T1.EDI_BOL_NO, EDT945T1.EDI_MASTER_BOL_NO, EDT945T1.EDI_SHIPMENT_DATE, EDT945T1.EDI_ORDR_SHIP_DATE" _
                & ", SOTSHIP1.ORDR_GROUP_NO" _
                & ", SOTPICK1.PICK_NO" _
                & ", SOTPICK1.ORDR_NO" _
                & ", SOTORDR1.ORDR_TYPE_CODE, SOTORDR1.CUST_CODE" _
                & ", NVL(EDT945T1.EDI_PROCESS_IND, '0') EDI_PROCESS_IND" _
                & " FROM EDT945T1, SOTPICK1, SOTSHIP1, SOTORDR1" _
                & " WHERE EDT945T1.EDI_PICK_NO = SOTPICK1.PICK_NO" _
                & " AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" _
                & " AND NVL(EDT945T1.EDI_PROCESS_IND, '0') = '0'" _
                & " AND SOTSHIP1.SHIP_STATUS = 'P'" _
                & " AND SOTPICK1.PICK_STATUS = 'P'" _
                & " AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"

            tblWHT3PLS1 = ASCMAIN1.Temp_Table(sqlWHT3PLS1 & " and rownum < 1")
            Create_TDA(dst.Tables.Add, "WHT3PLS1", "Select * from " & tblWHT3PLS1, 0, False, , 2)

            ' Used to capture data from ADS
            Create_TDA(.Tables.Add, "EDT945T1", "*")
            Create_TDA(.Tables.Add, "EDT945T2", "*")
            Create_TDA(.Tables.Add, "EDTTRXN1", "*")
            .Tables("EDTTRXN1").Columns.Add("ITEM_DESC", GetType(System.String))

            Create_TDA(.Tables.Add, "EDTRTRN1", "*")
            .Tables("EDTRTRN1").Columns.Add("CUST_NAME", GetType(System.String))
            Create_TDA(.Tables.Add, "EDTRTRN2", "*")

            .Tables.Add("EDTTRXNC")
            .Tables("EDTTRXNC").Columns.Add("TRANS_TYPE", GetType(System.String))
            .Tables("EDTTRXNC").Columns.Add("TRANS_DESC", GetType(System.String))

            .Tables("EDTTRXNC").Rows.Add(New Object() {"ADJ", "Adjustment"})
            .Tables("EDTTRXNC").Rows.Add(New Object() {"REC", "Receipts"})

            sqlSOTSHIPX = "Select SOTSHIP1.*" _
                & ", SOTORDR0.CUST_CODE, SOTORDR0.ORDR_SHIP_DATE, SOTORDR0.ORDR_CANCEL_DATE, SOTORDR0.ORDR_CUST_PO, '0' SELECTED" _
                & " from SOTSHIP1, SOTORDR0" _
                & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
                & " and SOTSHIP1.SHIP_STATUS = 'P'"
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, True, "", 1)
            ASCMAIN1.sql = sqlSOTSHIPX & " and ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTSHIP0", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "SOTSHIPB", "*")
            Create_TDA(.Tables.Add, "SOTSHIP3", "*")
            Create_TDA(.Tables.Add, "SOTSHIP4", "*")
            Create_TDA(.Tables.Add, "SOTSHIP6", "*")

            sqlSOTPICK1 = "Select SOTPICK1.*" & vbCrLf _
                 & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO, SOTORDR1.ORDR_NO_WEB, SOTORDR1.ORDR_SOURCE" & vbCrLf _
                 & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.CUST_BILL_TO_CUST" & vbCrLf _
                 & ", SOTORDR1.POST_CODE, SOTORDR1.WHSE_CODE, SOTORDR1.ORDR_FREIGHT" & vbCrLf _
                 & ", SOTORDR1.TERM_CODE, SOTORDR1.SREP_CODE, SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                 & ", SOTSHIP1.BILL_OF_LADING_NO, SOTORDR1.ORDR_INV_COMMENT, SOTORDR1.CUST_FACTOR_IND, SOTORDR1.CCPA_NO CCPA_NO_ORDR" & vbCrLf _
                 & " from SOTPICK1, SOTORDR1 ,SOTSHIP1 "
            ASCMAIN1.sql = sqlSOTPICK1 & " where ROWNUM < 1"
            Create_TDA(.Tables.Add, "SOTPICK1", "**")
            dst.Tables("SOTPICK1").Columns.Add("SELECTED")
            dst.Tables("SOTPICK1").Columns.Add("OUR_FREIGHT", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns("OUR_FREIGHT").DefaultValue = 0

            Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")

            sqlSOTPICK2 = "Select SOTPICK2.*, " & vbCrLf _
                & " SOTORDR2.ITEM_CODE, SOTORDR2.ITEM_DESC, " & vbCrLf _
                & " SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTSHIP1" & vbCrLf
            ASCMAIN1.sql = sqlSOTPICK2 & " where ROWNUM < 1" & vbCrLf
            Create_TDA(.Tables.Add, "SOTPICK2", "**")
            Create_TDA(.Tables.Add("SOTPICK2_CANC"), "SOTPICK2", "*")

            For Each table As String In New String() {"SOTPICK2", "SOTPICK2_CANC"}
                With .Tables(table).Columns
                    .Add("PICK_AMT", GetType(System.Decimal))
                    .Add("PICK_AMT_CONF", GetType(System.Decimal))
                    .Add("PICK_AMT_CANC", GetType(System.Decimal))
                    .Add("PICK_AMT_BACK", GetType(System.Decimal))
                    .Add("PICK_QTY_CANC_WSHE", GetType(System.Int32), "ISNULL(PICK_QTY, 0) - ISNULL(PICK_QTY_CONF, 0)")
                End With
            Next

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

            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*")

            Create_TDA(.Tables.Add("SOTORDR1_CANC"), "SOTORDR1", "*")
            Create_TDA(.Tables.Add("SOTORDR2_CANC"), "SOTORDR2", "*")

            Create_TDA(.Tables.Add("SOTORDR1_NOSHIP"), "SOTORDR1", "*")

            Create_TDA(.Tables.Add, "SOTORDR5", "*")
            Create_TDA(.Tables.Add, "TATEVNT1", "*")

            Create_TDA(.Tables.Add, "SOTTRAC1", "*")
            Create_TDA(.Tables.Add, "SOTTRAC2", "*")

            Create_TDA(.Tables.Add, "SOTCART1", "*")
            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
                & " from SOTCART2, SOTCART1 where SOTCART1.CART_NO = SOTCART2.CART_NO"
            Create_TDA(.Tables.Add, "SOTCART2", "**")

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            Create_TDA(.Tables.Add, "SOTCART3", "*")

            With .Tables.Add("SOTCARTX")
                .Columns.Add("PICK_NO", GetType(System.String))
                .Columns.Add("ORDR_NO", GetType(System.String))
                .Columns.Add("ORDR_LNO", GetType(System.Int64))
                .Columns.Add("PICK_QTY_CONF", GetType(System.Int64), "")
                .Columns.Add("QTY_PACKED", GetType(System.Int64), "")
                .PrimaryKey = New DataColumn() { .Columns("PICK_NO"), .Columns("ORDR_NO"), .Columns("ORDR_LNO")}
            End With

            Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
            Create_Relation("SOTCARTX", "SOTPICK2", "PICK_NO,ORDR_NO,ORDR_LNO")
            Create_Relation("SOTCARTX", "SOTCART2", "PICK_NO,ORDR_NO,ORDR_LNO")
            ' SR-6549 - Lot Numbers on Shipments and Receipts
            Create_Relation("SOTCART2", "SOTCART3", "CART_NO,CART_LNO")

            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_CALC", GetType(System.Int64))
            dst.Tables("SOTCART1").Columns.Add("CART_TOTAL_UNITS_ORIG", GetType(System.Int64), "SUM(CHILD.QTY_PACKED_ORIG)")

            Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_WGT_CALC", GetType(System.Decimal))
            dst.Tables("SOTPICK1").Columns.Add("PICK_CNT_CARTONS_CALC", GetType(System.Int64))
            dst.Tables("SOTPICK1").Columns.Add("PICK_TOTAL_UNITS_CALC", GetType(System.Int64))

            ASCMAIN1.sql = "SELECT ARTCCPA1.*, ARTCUST1.CUST_NAME FROM ARTCCPA1, ARTCUST1 WHERE ARTCCPA1.CUST_CODE = ARTCUST1.CUST_CODE (+)"
            Create_TDA(.Tables.Add, "ARTCCPA1", "**")
            Create_TDA(.Tables.Add, "ARTCCPA2", "*")
            Create_TDA(.Tables.Add, "ARTCCPDA", "*")
            Create_TDA(.Tables.Add, "ARTCCPAC", "*")
            Create_TDA(.Tables.Add, "ICTITEM1", "*")
            Fill_Records("ICTITEM1", String.Empty, True, "SELECT * FROM ICTITEM1")
            Create_TDA(.Tables.Add, "ICTSTAT2", "*")
            dst.Tables("ICTSTAT2").Columns.Add("QTY_SHIPPED", GetType(System.Int64))
            Create_TDA(.Tables.Add, "ICTWHSE1", "*")

            Create_TDA(.Tables.Add, "SOTINVH1", "*")
            Create_TDA(.Tables.Add, "SOTINVH2", "*")
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Fill_Records("ARTCUST1", String.Empty, True, "Select * from ARTCUST1")

            tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT * FROM SOTSVIA1", "SOTSVIA1")

            dst.Tables.Add("ERRORS")
            With dst.Tables("ERRORS")
                .Columns.Add("SEQ_NO", GetType(System.Int16))
                .Columns.Add("MESSAGE", GetType(System.String))
                .Columns.Add("FYI", GetType(System.String))
            End With

            dst.Tables.Add("LOT_NOS")
            With dst.Tables("LOT_NOS")
                .Columns.Add("CART_NO", GetType(System.String))
                .Columns.Add("CART_LNO", GetType(System.Int32))
                .Columns.Add("PICK_NO", GetType(System.String))
                .Columns.Add("ORDER_NO", GetType(System.String))
                .Columns.Add("ORDER_LNO", GetType(System.Int32))
                .Columns.Add("ITEM_CODE", GetType(System.String))
                .Columns.Add("QTY_PACKED", GetType(System.Int32))
                .Columns.Add("LOT_QTY", GetType(System.Int32))
            End With

            Get_PARM("SOTPARM1")

            ' SR-5537  04/04/2024
            ASCMAIN1.sql = "Select SOTCART1.*,SOTPICK1.SHIP_BOL_NO,SOTSHIP1.SHIP_ADDR_TYPE,SOTSHIP1.SHIP_ADDR_CODE, SOTORDR1.ORDR_CUST_PO" _
                        & " from SOTCART1, SOTPICK1, SOTSHIP1, SOTORDR1" _
                        & " where SOTPICK1.PICK_NO = SOTCART1.PICK_NO" _
                        & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
                        & "   and SOTSHIP1.SHIP_BOL_NO = :PARM1" _
                        & "   and SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTCART1X", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select SOTCART2.*" _
                        & " from SOTCART2, SOTCART1, SOTPICK1, SOTSHIP1" _
                        & " where SOTCART1.CART_NO = SOTCART2.CART_NO" _
                        & "   and SOTPICK1.PICK_NO = SOTCART1.PICK_NO" _
                        & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" _
                        & "   and SOTSHIP1.SHIP_BOL_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTCART2X", "**", 0, False, "V", 2)

            Create_Relation("SOTCART1X", "SOTCART2X", "CART_NO")

            grdSOTCART1.DataSource = dst.Tables("SOTCART1X")

            ' ISSUE-7230 ADS as the defualt warehouse
            Create_TDA(.Tables.Add("ICTWHSE1_LK"), "ICTWHSE1", "*")
            Fill_Records("ICTWHSE1_LK", "", True, "SELECT * FROM ICTWHSE1 WHERE LP_CODE IS NOT NULL")

        End With

        If perform_fill Then
            Fill_Records_RPT(sqlw)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""

        EnforceConstraints(False)

        Clear_Records()

        ' Get all the records for the shipments that need processing
        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & tblWHT3PLS1)
        ASCDATA1.ExecuteSQL("Insert  Into " & tblWHT3PLS1 & " " & sqlWHT3PLS1)

        ' **** Clean up data in bulk instead of individualy by SHIP_BOL_NO

        sql = " Select DISTINCT SOTSHIPB.BOL_NO, SOTSHIPB.CUST_CODE, " & tblWHT3PLS1 & ".CUST_CODE CUST_CODE_X"
        sql &= " FROM SOTSHIPB, " & tblWHT3PLS1
        sql &= " WHERE SOTSHIPB.BOL_NO =  " & tblWHT3PLS1 & ".EDI_BOL_NO"
        sql &= " AND " & tblWHT3PLS1 & ".EDI_BOL_NO IS NOT NULL"

        If ASCMAIN1.CLIENT = "AHA" Then
            sql &= " AND SOTSHIPB.BOL_NO <> 'CANCEL'"
        End If

        Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)
        For Each row As DataRow In tbl.Select("")
            If row.Item("CUST_CODE") & String.Empty = row.Item("CUST_CODE_X") & String.Empty Then
                If Not bolList.Contains(row.Item("BOL_NO") & String.Empty) Then
                    bolList.Add(row.Item("BOL_NO") & String.Empty)
                End If
            Else
                AddError("WARNING: EDI BOL NO " & row.Item("BOL_NO") & " was previously used by customer " & row.Item("CUST_CODE") _
                         & " and now it is used by " & row.Item("CUST_CODE_X") & ". Shipments will be skipped.", "1")

                sql = "DELETE FROM " & tblWHT3PLS1 & " where SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & tblWHT3PLS1 & " WHERE EDI_BOL_NO = '" & row.Item("BOL_NO") & "')"
                ASCDATA1.ExecuteSQL(sql)
            End If
        Next

        ' Validate Warehouse did not use multiple EDI_MASTER_BOL_NOs for the same SHIP_BOL_NO
        sql = "SELECT CUST_CODE, ORDR_GROUP_NO, SHIP_BOL_NO, COUNT(*) FROM" _
            & " (" _
            & "SELECT DISTINCT CUST_CODE, ORDR_GROUP_NO, SHIP_BOL_NO, EDI_MASTER_BOL_NO FROM " & tblWHT3PLS1 _
            & IIf(ASCMAIN1.CLIENT = "AHA", " WHERE EDI_MASTER_BOL_NO <> 'CANCEL'", "") _
            & " )" _
            & " GROUP BY CUST_CODE, ORDR_GROUP_NO, SHIP_BOL_NO" _
            & " HAVING COUNT(*) > 1"

        tbl = ASCDATA1.GetDataTable(sql)

        For Each rowData As DataRow In tbl.Select("", "CUST_CODE, ORDR_GROUP_NO")
            Dim CUST_CODE As String = rowData.Item("CUST_CODE") & String.Empty
            Dim SHIP_BOL_NO As String = rowData.Item("SHIP_BOL_NO") & String.Empty

            Dim ptMessage As String = "WARNING: Customer: " & CUST_CODE & ", Order Group No: " & rowData.Item("ORDR_GROUP_NO") & ", Shipment: " & SHIP_BOL_NO & "" _
                                      & " contains multiple Master BOL Nos. These records must have the same Master BOL No. This shipment will not be processed."
            AddError(ptMessage.Trim, "1")
            sql = "DELETE FROM " & tblWHT3PLS1 & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            ASCDATA1.ExecuteSQL(sql)
        Next

        ' Look for missing Pick Tickets for each Ship Bol No
        sql = " SELECT ABS.*, CLA.TO_PROCESS RECEIVED, NVL(ABS.TO_PROCESS, 0) - NVL(CLA.TO_PROCESS, 0) MISSING"
        sql &= " FROM"
        sql &= " (Select SOTPICK1.SHIP_BOL_NO, SOTSHIP1.ORDR_GROUP_NO, SOTORDR0.CUST_CODE, COUNT(*) TO_PROCESS"
        sql &= " 		FROM SOTPICK1, SOTSHIP1, SOTORDR0"
        sql &= " 		where SOTPICK1.PICK_STATUS = 'P'"
        sql &= " 		AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO"
        sql &= " 		AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO"
        sql &= " 		and SOTPICK1.SHIP_BOL_NO IN (Select SHIP_BOL_NO FROM " & tblWHT3PLS1 & ")"
        sql &= " 		GROUP BY SOTPICK1.SHIP_BOL_NO, SOTSHIP1.ORDR_GROUP_NO, SOTORDR0.CUST_CODE) ABS,"
        sql &= " (Select SHIP_BOL_NO, ORDR_GROUP_NO, CUST_CODE, COUNT(*) TO_PROCESS FROM  " & tblWHT3PLS1 & "  GROUP BY SHIP_BOL_NO, ORDR_GROUP_NO, CUST_CODE) CLA"
        sql &= " WHERE ABS.SHIP_BOL_NO = CLA.SHIP_BOL_NO (+)"
        sql &= " AND ABS.TO_PROCESS <> CLA.TO_PROCESS"

        tbl = ASCDATA1.GetDataTable(sql)
        Dim tblDIFF As DataTable = Nothing
        If tbl.Rows.Count > 0 Then
            sql = "Select SHIP_BOL_NO, PICK_NO FROM SOTPICK1 where PICK_STATUS = 'P' and SHIP_BOL_NO IN (Select SHIP_BOL_NO FROM " & tblWHT3PLS1 & ")"
            sql &= " MINUS "
            sql &= " Select SHIP_BOL_NO, PICK_NO FROM " & tblWHT3PLS1
            tblDIFF = ASCDATA1.GetDataTable(sql)
        End If

        For Each rowData As DataRow In tbl.Select("", "CUST_CODE, ORDR_GROUP_NO")
            Dim CUST_CODE As String = rowData.Item("CUST_CODE") & String.Empty
            Dim SHIP_BOL_NO As String = rowData.Item("SHIP_BOL_NO") & String.Empty
            Dim ORDR_GROUP_NO As String = rowData.Item("ORDR_GROUP_NO") & String.Empty

            Dim ptMessage As String = "WARNING: Customer: " & CUST_CODE & ", Order Group No: " & ORDR_GROUP_NO & ", Shipment: " & SHIP_BOL_NO & "" _
                          & " is missing " & rowData.Item("MISSING") & " of " & rowData.Item("TO_PROCESS") & " Pick Ticket(s). It will not be processed until all Pick Tickets have Ship Confirmation Data." _
                          & " The missing Pick Ticket(s) are: "

            For Each rowDiff As DataRow In tblDIFF.Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "PICK_NO")
                ptMessage &= rowDiff.Item("PICK_NO") & ", "
            Next

            ptMessage = ptMessage.Trim
            If ptMessage.EndsWith(",") Then
                ptMessage = ptMessage.Substring(0, ptMessage.Length - 1)
            End If

            AddError(ptMessage, "1")
            sql = "DELETE FROM " & tblWHT3PLS1 & " where SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"
            ASCDATA1.ExecuteSQL(sql)
        Next

        sql = " SELECT SOTSHIP1.SHIP_BOL_NO, ARTCUST2.*"
        sql &= " FROM ARTCUST2, SOTORDR5, SOTORDR1, SOTSHIP1"
        sql &= " WHERE SOTSHIP1.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO"
        sql &= " AND SOTORDR1.ORDR_NO = SOTORDR5.ORDR_NO"
        sql &= " AND CUST_ADDR_TYPE = 'ST'"
        sql &= " AND ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE"
        sql &= " AND ARTCUST2.CUST_STORE_NO = SOTORDR5.CUST_ADDR_CODE"
        sql &= " AND SOTSHIP1.SHIP_BOL_NO IN (SELECT SHIP_BOL_NO FROM " & tblWHT3PLS1 & ")"
        sql &= " AND ARTCUST2.CUST_SHIP_EMAIL IS NOT NULL"
        tblSOTSHIP_CART = ASCDATA1.GetDataTable(sql)

        Fill_Records("WHT3PLS1")

        sqlw = "Select * from EDT945T1 where EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblWHT3PLS1 & ")"
        tblEDT945T1 = ASCMAIN1.Temp_Table(sqlw)
        Fill_Records("EDT945T1", String.Empty, True, "SELECT * FROM " & tblEDT945T1)

        sqlw = "Select * from EDT945T2 where EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblWHT3PLS1 & ")"
        tblEDT945T2 = ASCMAIN1.Temp_Table(sqlw)
        Fill_Records("EDT945T2", String.Empty, True, "SELECT * FROM " & tblEDT945T2)

        Select Case Company_Code
            Case "INT"
                ' For now we will not allow the user to process an RTV
                For Each rowEDT945T1x As DataRow In ASCDATA1.GetDataTable("SELECT * FROM " & tblWHT3PLS1 _
                                              & " Where ORDR_TYPE_CODE = 'RTV'").Select("")
                    Dim ORDR_NO As String = rowEDT945T1x.Item("ORDR_NO")
                    AddError("Order No: " & ORDR_NO & " is for an RTV and will not be processed.", "1")

                    Dim EDI_DOC_SEQ_NO As String = rowEDT945T1x.Item("EDI_DOC_SEQ_NO")

                    For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                        rowEDT945T1.Delete()
                    Next

                    For Each rowEDT945T2 As DataRow In dst.Tables("EDT945T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                        rowEDT945T2.Delete()
                    Next

                    ASCDATA1.ExecuteSQL("DELETE FROM " & tblEDT945T1 & " WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    ASCDATA1.ExecuteSQL("DELETE FROM " & tblEDT945T2 & " WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                Next

                dst.Tables("EDT945T1").AcceptChanges()
                dst.Tables("EDT945T2").AcceptChanges()

        End Select

        EnforceConstraints(True)

    End Sub

    ''' <summary>
    ''' Process the shipments
    ''' </summary>
    ''' <remarks></remarks>
    Private Function ValidateShipments() As Boolean

        Try
            Dim sql As String = String.Empty
            Dim shipments As New List(Of String)
            Dim tbl As DataTable = Nothing
            Dim errorMessage As String = String.Empty

            Dim EDI_FRT_COST As Double = 0
            Dim EDI_SHIPMENT_ID As String = String.Empty
            Dim eMsg As String = String.Empty

            Dim lstMaster As New List(Of String)

            Dim tblWHT3PLSx As DataTable = ASCDATA1.SelectDistinct(dst.Tables("WHT3PLS1"), New String() {"EDI_MASTER_BOL_NO", "CUST_CODE"})

            For Each row As DataRow In tblWHT3PLSx.Select("", "EDI_MASTER_BOL_NO, CUST_CODE")
                Dim EDI_MASTER_BOL_NO As String = row.Item("EDI_MASTER_BOL_NO") & String.Empty
                Dim CUST_CODE As String = row.Item("CUST_CODE") & String.Empty

                If EDI_MASTER_BOL_NO.Length = 0 Then
                    Continue For
                End If

                ASCMAIN1.Progress("Validate Shipments: " & CUST_CODE & "/" & EDI_MASTER_BOL_NO)

                Dim bolList As New List(Of String)
                Dim exitInnerLoop As Boolean = False

                For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("EDI_MASTER_BOL_NO = '" & EDI_MASTER_BOL_NO & "' AND CUST_CODE = '" & CUST_CODE & "'", "EDI_SHIPMENT_ID")
                    If exitInnerLoop Then Exit For
                    Dim EDI_PICK_NO As String = rowEDT945T1.Item("EDI_PICK_NO")
                    EDI_SHIPMENT_ID = rowEDT945T1.Item("EDI_SHIPMENT_ID")

                    If lstMaster.Contains(EDI_SHIPMENT_ID) Then
                        Continue For
                    End If

                    lstMaster.Add(EDI_SHIPMENT_ID)

                    ' Need to see if there are freight issues
                    EDI_FRT_COST = Val(rowEDT945T1.Item("EDI_FRT_COST") & String.Empty)
                    EDI_SHIPMENT_ID = rowEDT945T1.Item("EDI_SHIPMENT_ID") & String.Empty
                    eMsg = String.Empty

                    If EDI_FRT_COST = 0 Then
                        ' CONSUMER may have $0.00 freight; therefore they are skipped.
                        sql = "Select * "
                        sql &= " from SOTSHIP1, SOTORDR1"
                        sql &= " where SOTSHIP1.ORDR_GROUP_NO = SOTORDR1.ORDR_GROUP_NO"
                        sql &= " and SOTORDR1.CUST_CODE <> 'CONSUMER'"
                        sql &= " and SOTSHIP1.FRT_TERMS = 'PPA'"
                        sql &= " and SOTSHIP1.SHIP_BOL_NO = '" & EDI_SHIPMENT_ID & "'"
                        sql &= " and NVL(SOTORDR1.ORDR_FREIGHT, 0) = 0"

                        Dim rowPPA As DataRow = ASCDATA1.GetDataRow(sql)

                        If rowPPA IsNot Nothing Then
                            eMsg = "PPA Freight Terms does not permit Zero Freight. Shipment: " & EDI_SHIPMENT_ID

                            ' If PPA check Ship Via Code to see if Collect or Third Party
                            Dim shipViaCode As String = rowEDT945T1.Item("EDI_CARRIER_SCAC_CODE") & String.Empty
                            If shipViaCode.Length > 0 Then
                                Dim rowSOTSVIA1 As DataRow = tblSOTSVIA1.Rows.Find(shipViaCode)
                                If rowSOTSVIA1 IsNot Nothing _
                                    AndAlso (rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty = "1" _
                                    OrElse rowSOTSVIA1.Item("COLLECT_IND") & String.Empty = "1") Then
                                    eMsg = String.Empty
                                End If
                            End If

                            If eMsg.Length > 0 Then
                                Using frmASFMSGBF As New ASFMSGBF
                                    Dim frtOption As Integer = frmASFMSGBF.Get_opt_from_User _
                                    ("Options for Customer: " & rowPPA.Item("CUST_CODE") & "/" & rowPPA.Item("ORDR_CUST_PO"), New String() {"Set to PPD Freight Terms", "Provide Freight Amount"},
                                     0, "PPA Freight Terms does not permit Zero Freight.")

                                    Select Case frtOption
                                        Case 0
                                            Dim SHIP_BOL_NO As String = rowPPA.Item("SHIP_BOL_NO") & String.Empty
                                            ASCDATA1.ExecuteSQL("UPDATE SOTSHIP1 SET FRT_TERMS = 'PPD' WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                                            eMsg = String.Empty

                                        Case 1

                                            Using frmASFMSGBF1 As New ASFMSGBF
                                                Dim userEdiFrtCost As Decimal = frmASFMSGBF1.Get_numdec_from_User(
                                                     "Provide Freight Cost.", "Freight Cost", 2500, 0, 0)

                                                If userEdiFrtCost > 0 Then
                                                    EDI_FRT_COST = userEdiFrtCost
                                                    rowEDT945T1.Item("EDI_FRT_COST") = EDI_FRT_COST
                                                    ASCDATA1.ExecuteSQL("UPDATE EDT945T1 SET EDI_FRT_COST = " & EDI_FRT_COST & " WHERE EDI_SHIPMENT_ID = '" & EDI_SHIPMENT_ID & "' and EDI_PROCESS_IND = '0'")
                                                    eMsg = String.Empty
                                                End If
                                            End Using

                                    End Select
                                End Using
                            End If

                            If eMsg.Length > 0 Then
                                AddError(eMsg, "1")
                                sql = "Update " & tblWHT3PLS1 & " set edi_process_ind = 'X' where EDI_MASTER_BOL_NO = '" & EDI_MASTER_BOL_NO & "' AND CUST_CODE = '" & CUST_CODE & "'"
                                ASCDATA1.ExecuteSQL(sql)
                                exitInnerLoop = True
                                Continue For
                            End If
                        End If
                    End If
                Next
            Next

            sql = "Delete from " & tblWHT3PLS1 & " where edi_process_ind = 'X'"
            ASCDATA1.ExecuteSQL(sql)

            ' Requery if shipments had errors and were deleted from the batch
            ASCMAIN1.Progress("Validate Shipments: WHT3PLS1")
            Fill_Records("WHT3PLS1")

            ASCMAIN1.Progress("Validate Shipments: EDT945T1")
            sql = "Select * from EDT945T1 where EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblWHT3PLS1 & ")"
            tblEDT945T1 = ASCMAIN1.Temp_Table(sql)
            Fill_Records("EDT945T1", String.Empty, True, "SELECT * FROM " & tblEDT945T1)
            'Fill_Records("EDT945T1", String.Empty, True, sql)

            ASCMAIN1.Progress("Validate Shipments: EDT945T2")
            sql = "Select * from EDT945T2 where EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblWHT3PLS1 & ")"
            tblEDT945T2 = ASCMAIN1.Temp_Table(sql)
            Fill_Records("EDT945T2", String.Empty, True, "SELECT * FROM " & tblEDT945T2)
            'Fill_Records("EDT945T2", String.Empty, True, sql)

            Return True

        Catch ex As Exception
            MessageBox.Show("Validate Shipments Error: " & ex.Message, "Validate Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function Load_Shipments() As Boolean

        Try
            EnforceConstraints(False)

            Dim rowSOTORDR1 As DataRow = Nothing

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Loading Shipment Data")

            tblSOTCART1 = ASCMAIN1.Temp_Table("SELECT * FROM SOTCART1 WHERE ROWNUM < 1")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & tblSOTCART1 & " ADD PRIMARY KEY(CART_NO)")

            tblSOTCART2 = ASCMAIN1.Temp_Table("SELECT * FROM SOTCART2 WHERE ROWNUM < 1")
            ASCDATA1.ExecuteSQL("ALTER TABLE " & tblSOTCART2 & " ADD PRIMARY KEY(CART_NO, CART_LNO)")
            ASCDATA1.ExecuteSQL("ALTER TABLE  " & tblSOTCART2 & " MODIFY CART_LNO NUMBER(5)")

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            tblSOTCART3 = ASCMAIN1.Temp_Table("SELECT * FROM SOTCART3 WHERE ROWNUM < 1")
            ASCDATA1.ExecuteSQL($"ALTER TABLE {tblSOTCART3} ADD PRIMARY KEY(CART_NO, CART_LNO, CART_LNO_SEQ)")
            ASCDATA1.ExecuteSQL($"ALTER TABLE {tblSOTCART3} MODIFY CART_LNO NUMBER(5)")
            ASCDATA1.ExecuteSQL($"ALTER TABLE {tblSOTCART3} MODIFY CART_LNO_SEQ NUMBER(5)")

            sql = " INSERT INTO " & tblSOTCART1
            sql &= " (CART_NO, PICK_NO, CART_TOTAL_UNITS, CART_TRACKING_NO, PACKAGING_TYPE, PKG_CODE, CART_TOTAL_WGT_ACTUAL, CART_TOTAL_WGT_CALC, CART_SEQ)"
            sql &= " SELECT X.*, X.CART_TOTAL_WGT_ACTUAL CART_TOTAL_WGT_CALC,  ROWNUM CART_SEQ FROM"
            sql &= " ("
            sql &= " SELECT DISTINCT EDT945T2.EDI_CART_NO CART_NO, EDT945T1.EDI_PICK_NO PICK_NO, 0 CART_TOTAL_UNITS"
            sql &= " , EDT945T2.EDI_SHIPPER_ID_NO CART_TRACKING_NO "
            sql &= " , '31' PACKAGING_TYPE, 'X2' PKG_CODE, MAX(NVL(EDI_CART_WEIGHT, 0)) CART_TOTAL_WGT_ACTUAL"
            sql &= " FROM EDT945T1, EDT945T2"
            sql &= " WHERE EDT945T1.EDI_DOC_SEQ_NO = EDT945T2.EDI_DOC_SEQ_NO"
            sql &= " AND EDT945T1.EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblWHT3PLS1 & ")"
            sql &= " AND EDT945T2.EDI_CART_NO IS NOT NULL"
            sql &= " GROUP BY EDT945T2.EDI_CART_NO, EDT945T1.EDI_PICK_NO, EDT945T2.EDI_SHIPPER_ID_NO"
            sql &= " ) X"
            ASCDATA1.ExecuteSQL(sql)

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            ' 03/19/2025 - changed 
            '   SELECT EDT945T2.EDI_CART_NO CART_NO, EDT945T2.EDI_DTL_SEQ CART_LNO
            ' to
            '   SELECT EDT945T2.EDI_CART_NO CART_NO, EDT945T2.EDI_DTL_SEQ CART_LNO
            sql = " Insert Into " & tblSOTCART2
            sql &= " (CART_NO, CART_LNO, ORDR_NO, ORDR_LNO, QTY_PACKED, ITEM_UPC_CODE, ITEM_EAN_CODE, ITEM_CODE)"
            sql &= " SELECT EDT945T2.EDI_CART_NO CART_NO, EDT945T2.EDI_DTL_SEQ CART_LNO"
            sql &= " , SOTPICK2.ORDR_NO, SOTPICK2.ORDR_LNO"
            sql &= " , NVL(EDT945T2.EDI_SHIP_QTY, 0) QTY_PACKED"
            sql &= " , ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE"
            sql &= " , EDT945T2.STYLE_CODE ITEM_CODE"
            sql &= " FROM EDT945T1, EDT945T2, SOTPICK2, ICTITEM1"
            sql &= " WHERE EDT945T1.EDI_DOC_SEQ_NO = EDT945T2.EDI_DOC_SEQ_NO"
            sql &= " AND EDT945T2.EDI_CART_NO IS NOT NULL"
            sql &= " AND SOTPICK2.PICK_NO = EDT945T1.EDI_PICK_NO"
            sql &= " AND SOTPICK2.PICK_LNO = EDT945T2.PICK_LNO"
            sql &= " AND EDT945T2.STYLE_CODE = ICTITEM1.ITEM_CODE (+)"
            sql &= " AND EDT945T1.EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM " & tblWHT3PLS1 & ")"
            ASCDATA1.ExecuteSQL(sql)

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            sql = $"Insert Into {tblSOTCART3}
                        (CART_NO, CART_LNO, CART_LNO_SEQ, LOT_QTY, LOT_NO, LOT_FIFO_DATE, LOT_SHELF_LIFE_DAYS, LOT_EXPIRATION_DATE)
                    SELECT T2.EDI_CART_NO CART_NO, T2.EDI_DTL_SEQ CART_LNO, T3.EDI_DTL_LNO CART_LNO_SEQ,
                    T3.EDI_QTY LOT_QTY, T3.EDI_LOT_NO LOT_NO, T3.EDI_FIFO_DATE LOT_FIFO_DATE,
                    T3.EDI_SHELF_LIFE_DAYS LOT_SHELF_LIFE_DAYS, T3.EDI_EXPIRATION_DATE LOT_EXPIRATION_DATE
                    FROM EDT945T2 T2, EDT945T3 T3
                    WHERE T2.EDI_DOC_SEQ_NO = T3.EDI_DOC_SEQ_NO
                    AND T2.EDI_DTL_SEQ = T3.EDI_DTL_SEQ
                    AND T2.EDI_DOC_SEQ_NO IN (SELECT EDI_DOC_SEQ_NO FROM {tblWHT3PLS1})"
            ASCDATA1.ExecuteSQL(sql)

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            ' 03/19/2025 - commented out becuse of change above.
            '' Start off all cartons with LNO 1
            'sql = " begin"
            'sql &= "    declare cursor C1 is select distinct cart_no from " & tblSOTCART2 & ";"
            'sql &= "    LNO NUMBER;"
            'sql &= "    EDI_DTL_SEQ NUMBER;"
            'sql &= "    begin for r1 in c1 loop"
            'sql &= "        LNO := 0;"
            'sql &= "        begin declare cursor c2 is select * from " & tblSOTCART2 & " where cart_no = r1.cart_no order by cart_lno;"
            'sql &= "            BEGIN FOR R2 IN C2 LOOP"
            'sql &= "                LNO := LNO + 1;"
            'sql &= "                UPDATE " & tblSOTCART2 & " SET CART_LNO = LNO WHERE CART_NO = R2.CART_NO AND CART_LNO = R2.CART_LNO;"
            'sql &= "            end loop;"
            'sql &= "            end; END;"
            'sql &= "    end loop;"
            'sql &= "end ; END;"
            'ASCDATA1.ExecuteSQL(sql)

            ASCDATA1.ExecuteSQL("Delete from " & tblSOTCART2 & " where QTY_PACKED = 0")
            ASCDATA1.ExecuteSQL("Delete from " & tblSOTCART1 & " where CART_NO NOT IN (SELECT CART_NO FROM " & tblSOTCART2 & ")")
            ' SR-6549 - Lot Numbers on Shipments and Receipts
            ASCDATA1.ExecuteSQL("Delete from " & tblSOTCART3 & " where CART_NO NOT IN (SELECT CART_NO FROM " & tblSOTCART2 & ")")
            ASCDATA1.ExecuteSQL("UPDATE " & tblSOTCART1 & " SET CART_TOTAL_UNITS = (SELECT SUM(QTY_PACKED) FROM " & tblSOTCART2 & " WHERE CART_NO = " & tblSOTCART1 & ".CART_NO)")

            Get_PARM("SOTPARM1")
            ' Refresh the list - the inporting of 945s creates ship vias on the fly.
            tblSOTSVIA1 = ASCDATA1.GetDataTable("SELECT * FROM SOTSVIA1", "SOTSVIA1")

            ASCMAIN1.sql = "Select * from SOTORDR1 Where ORDR_NO IN"
            ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO in (Select SHIP_BOL_NO from " & tblWHT3PLS1 & ") AND PICK_STATUS = 'P')"
            Fill_Records("SOTORDR1", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from SOTORDR2 Where ORDR_NO IN"
            ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO in (Select SHIP_BOL_NO from " & tblWHT3PLS1 & ") AND PICK_STATUS = 'P')"
            Fill_Records("SOTORDR2", String.Empty, True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select * from SOTORDR5 Where ORDR_NO IN"
            ASCMAIN1.sql &= " (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO in (Select SHIP_BOL_NO from " & tblWHT3PLS1 & ") AND PICK_STATUS = 'P')"
            Fill_Records("SOTORDR5", "", True, ASCMAIN1.sql)

            Dim sqlwhere_SOTSHIP1 As String = " and SOTSHIP1.SHIP_BOL_NO in (Select SHIP_BOL_NO from " & tblWHT3PLS1 & ")"
            ASCMAIN1.sql = sqlSOTSHIPX & sqlwhere_SOTSHIP1
            Fill_Records("SOTSHIP1", "", True, ASCMAIN1.sql)

            ' Added " AND SOTPICK1.PICK_STATUS = 'P'" on 11/4/2016 - At INT there was one de-released pick ticket in a group of 20.
            sqlwhere_SOTSHIP1 &= " AND SOTPICK1.PICK_STATUS = 'P' and SOTPICK1.PICK_NO IN (Select PICK_NO from " & tblWHT3PLS1 & ")"

            ASCMAIN1.sql = "SELECT * FROM ICTWHSE1 WHERE WHSE_CODE IN "
            ASCMAIN1.sql &= " (SELECT WHSE_CODE FROM SOTSHIP1 WHERE SHIP_BOL_NO in (Select SHIP_BOL_NO from " & tblWHT3PLS1 & "))"
            Fill_Records("ICTWHSE1", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT ICTSTAT2.*, 0 QTY_SHIPPED FROM ICTSTAT2 WHERE WHSE_CODE IN "
            ASCMAIN1.sql &= " (SELECT WHSE_CODE FROM SOTSHIP1 WHERE SHIP_BOL_NO in (Select SHIP_BOL_NO from " & tblWHT3PLS1 & "))"
            Fill_Records("ICTSTAT2", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = sqlSOTPICK1 & vbCrLf _
                & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1
            Fill_Records("SOTPICK1", "", True, ASCMAIN1.sql)

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_STATUS = 'P'")
                rowSOTPICK1.Item("SELECTED") = "1"
            Next

            ASCMAIN1.Progress("Load Shipment Records")
            Dim rowARTCUST1 As DataRow = Nothing
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("", "CUST_CODE")
                ASCMAIN1.Progress("-", rowSOTSHIP1.Item("CUST_CODE") & String.Empty)

                If rowARTCUST1 Is Nothing OrElse rowSOTSHIP1.Item("CUST_CODE") <> rowARTCUST1.Item("CUST_CODE") Then
                    rowARTCUST1 = LookUp("ARTCUST1", rowSOTSHIP1.Item("CUST_CODE"))
                End If

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

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                dst.Tables("SOTSHIP0").Rows.Add(rowSOTSHIP1.ItemArray)
                For Each rowSOTPICK1_0 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'")
                    For Each COLUMN_NAME As String In New String() _
                        {"ORDR_DEPT", "SREP_CODE", "SREP2_CODE", "TERM_CODE"}
                        If rowSOTSHIP1.Item(COLUMN_NAME) & "" = "" Then rowSOTSHIP1.Item(COLUMN_NAME) = rowSOTPICK1_0.Item(COLUMN_NAME)
                    Next
                Next
            Next

            ASCMAIN1.Progress("Load Shipment - Pick Ticket Details")
            ASCDATA1.ExecuteSQL("UPDATE SOTPICK2 SET PICK_QTY_CONF = 0 WHERE PICK_NO IN (SELECT PICK_NO FROM " & tblSOTCART1 & ")")

            ' Allow for no EDT945T2 records
            If ASCMAIN1.CLIENT = "AHA" Then
                sql = "Update SOTPICK2 SET PICK_QTY_CONF = 0 WHERE PICK_NO IN (SELECT PICK_NO FROM " & tblWHT3PLS1 & ") and PICK_NO IN (SELECT PICK_NO FROM SOTPICK1 WHERE PICK_STATUS = 'P')"
                ASCDATA1.ExecuteSQL(sql)
            End If

            ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
                & " SELECT SOTPICK2.PICK_NO, SOTPICK2.PICK_LNO, SUM(SOTCART2.QTY_PACKED) QTY_PACKED" _
                & " FROM " & tblSOTCART1 & " SOTCART1, " & tblSOTCART2 & " SOTCART2, SOTPICK2" _
                & " WHERE SOTCART1.CART_NO = SOTCART2.CART_NO" _
                & " AND SOTCART2.ORDR_NO = SOTPICK2.ORDR_NO" _
                & " AND SOTCART2.ORDR_LNO = SOTPICK2.ORDR_LNO" _
                & " AND SOTPICK2.PICK_NO = SOTCART1.PICK_NO" _
                & " GROUP BY SOTPICK2.PICK_NO, SOTPICK2.PICK_LNO;" _
                & " BEGIN FOR R1 IN C1 LOOP" _
                & "     UPDATE SOTPICK2 SET PICK_QTY_CONF = PICK_QTY_CONF + R1.QTY_PACKED WHERE PICK_NO = R1.PICK_NO AND PICK_LNO = R1.PICK_LNO;" _
                & " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = sqlSOTPICK2 & vbCrLf _
                & " where SOTPICK2.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1

            sql = "Begin Declare Cursor C1 is " & ASCMAIN1.sql & ";" _
                    & " Begin For R1 in C1 Loop" _
                    & "  Update sotpick2 set PICK_UNIT_PRICE = R1.ORDR_UNIT_PRICE where PICK_NO = R1.PICK_NO AND PICK_LNO = R1.PICK_LNO;" _
                    & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL(sql)
            Fill_Records("SOTPICK2", "", True, ASCMAIN1.sql)

            ASCMAIN1.Progress("Load Shipment - Carton Header")

            ' New code on 10/07/2020 to speed up Query
            Dim wkTable As String = ASCMAIN1.Temp_Table($"(Select DISTINCT {tblWHT3PLS1}.PICK_NO from {tblWHT3PLS1}, SOTPICK1 where SOTPICK1.PICK_NO = {tblWHT3PLS1}.PICK_NO and SOTPICK1.PICK_STATUS = 'P')")
            ASCDATA1.ExecuteSQL($"ALTER TABLE {wkTable} ADD PRIMARY KEY (PICK_NO)")

            ASCMAIN1.sql = "Select SOTCART1.*" & vbCrLf _
                & " from " & tblSOTCART1 & " SOTCART1 , SOTPICK1, SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1

            ' New code on 10/07/2020 to speed up Query
            ASCMAIN1.sql = $"Select SOTCART1.*
                    from {tblSOTCART1} SOTCART1, {wkTable} WHT3PLS1
                    where SOTCART1.PICK_NO = WHT3PLS1.PICK_NO"

            Fill_Records("SOTCART1", "", True, ASCMAIN1.sql)

            For Each row As DataRow In dst.Tables("SOTCART1").Select("")
                row.SetAdded()
            Next

            ASCMAIN1.Progress("Load Shipment - Carton Details")
            ASCMAIN1.sql = "Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG" & vbCrLf _
                & " from " & tblSOTCART2 & " SOTCART2," & tblSOTCART1 & " SOTCART1, SOTPICK1, SOTSHIP1" & vbCrLf _
                & " where SOTCART1.PICK_NO = SOTPICK1.PICK_NO" & vbCrLf _
                & "   and SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                & "   and SOTCART2.CART_NO = SOTCART1.CART_NO" & vbCrLf _
                & sqlwhere_SOTSHIP1

            ' New code on 10/07/2020 to speed up Query
            ASCMAIN1.sql = $"Select SOTCART2.*, SOTCART1.PICK_NO, SOTCART2.QTY_PACKED QTY_PACKED_ORIG
                            from  {tblSOTCART2} SOTCART2,  {tblSOTCART1} SOTCART1, {wkTable} WHT3PLS1 
                            where SOTCART1.CART_NO = SOTCART2.CART_NO 
                            and SOTCART1.PICK_NO = WHT3PLS1.PICK_NO"
            Fill_Records("SOTCART2", "", True, ASCMAIN1.sql)

            For Each row As DataRow In dst.Tables("SOTCART2").Select("")
                row.SetAdded()
            Next

            ' SR-6549 - Lot Numbers on Shipments and Receipts
            ASCMAIN1.sql = $"SELECT * FROM {tblSOTCART3}"
            Fill_Records("SOTCART3", "", True, ASCMAIN1.sql)

            For Each row As DataRow In dst.Tables("SOTCART3").Select("")
                row.SetAdded()
            Next

            Dim GL_PARM_CURR_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & ""
            'CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
            If CURR_CODE = "" Or CURR_CODE = GL_PARM_CURR_CODE Then
                CURR_CODE = GL_PARM_CURR_CODE
                CURR_EXCH_RATE = 1
            Else
                Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", rowARTCUST1.Item("CURR_CODE"))
                CURR_CODE = rowTATCURR1.Item("CURR_CODE")
                CURR_EXCH_RATE = rowTATCURR1.Item("CURR_EXCH_CUR")
            End If

            '******************************************************************************
            ' Delete original Cartons, New ones created using 945 data

            Dim EDI_TOTAL_ORDR_WEIGHT As Decimal = 0
            Dim EDI_DOC_SEQ_NO As String = String.Empty
            Dim PICK_NO As String = String.Empty
            Dim ORDR_NO As String = String.Empty
            Dim rowSOTSVIA1 As DataRow = Nothing
            Dim shipViaCode As String = String.Empty

            ' Need to create new cartons using edi data
            ASCMAIN1.Progress("Create New Cartons for Pick Tickets")
            For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("", "EDI_DOC_SEQ_NO")
                EDI_TOTAL_ORDR_WEIGHT = Val(rowEDT945T1.Item("EDI_TOTAL_ORDR_WEIGHT") & String.Empty)
                EDI_DOC_SEQ_NO = rowEDT945T1.Item("EDI_DOC_SEQ_NO") & String.Empty
                PICK_NO = rowEDT945T1.Item("EDI_PICK_NO") & String.Empty
                Dim EDI_CART_NO As String = String.Empty

                ASCMAIN1.Progress("Process Pick Ticket: " & PICK_NO, "")

                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & rowEDT945T1.Item("EDI_SHIPMENT_ID") & "'")(0)

                Select Case Company_Code
                    Case "AHA"
                        shipViaCode = rowEDT945T1.Item("EDI_CARRIER_SCAC_CODE") & String.Empty
                    Case "INT"
                        shipViaCode = rowEDT945T1.Item("EDI_CARRIER_CODE") & String.Empty
                End Select

                rowSOTSVIA1 = tblSOTSVIA1.Rows.Find(shipViaCode)

                If rowSOTSVIA1 Is Nothing Then
                    AddError("Error processing EDI Doc No/Pick No (" & EDI_DOC_SEQ_NO & "/" & PICK_NO & ") Invalid SCAC Code/Name: " & shipViaCode & "/" & rowEDT945T1.Item("EDI_CARRIER_NAME") & String.Empty)
                    Continue For
                End If

                rowEDT945T1.Item("EDI_PROCESS_IND") = "1"
                ' Initialize to 0. Fill back in with EDT945T2
                Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Select("PICK_NO = '" & PICK_NO & "'")(0)

                ' Update Shipment Header
                rowSOTSHIP1 = dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & rowSOTPICK1.Item("SHIP_BOL_NO") & "'")(0)
                rowSOTSHIP1.Item("BILL_OF_LADING_NO") = rowEDT945T1.Item("EDI_BOL_NO")
                rowSOTSHIP1.Item("SHIP_REF") = rowEDT945T1.Item("EDI_SHIPPER_ID_NO")

                ' Need to add freight to PPA (Pre Paid and Add); but not Web Orders. That freight is on the Order Header and the Invoice class
                ' adds the Order Freight and Pick Ticket freight together.
                ' 1/31/2013 - Look at Collect and Third party settings on the Ship Via.

                Select Case Company_Code
                    Case "AHA"
                        If rowSOTSHIP1.Item("FRT_TERMS") & String.Empty = "PPA" _
                            AndAlso rowSOTPICK1.Item("ORDR_SOURCE") & String.Empty <> "W" _
                            AndAlso rowSOTSVIA1.Item("THIRD_PARTY_IND") & String.Empty <> "1" _
                            AndAlso rowSOTSVIA1.Item("COLLECT_IND") & String.Empty <> "1" Then
                            ' Add ADS fright charge only if there is not freight on the Sales Order.
                            If Val(rowSOTPICK1.Item("ORDR_FREIGHT") & String.Empty) = 0 Then
                                rowSOTPICK1.Item("PICK_FREIGHT") = Val(rowEDT945T1.Item("EDI_FRT_COST") & String.Empty)
                            End If
                        End If

                        ' Use the ADS Feight Cost as Our Cost.
                        rowSOTPICK1.Item("OUR_FREIGHT") = Val(rowEDT945T1.Item("EDI_FRT_COST") & String.Empty)

                    Case "INT"
                        rowSOTPICK1.Item("PICK_FREIGHT") = Val(rowEDT945T1.Item("EDI_FRT_COST") & String.Empty)
                        rowSOTPICK1.Item("OUR_FREIGHT") = Val(rowEDT945T1.Item("EDI_FRT_COST") & String.Empty)
                End Select

                rowSOTSHIP1.Item("SHIP_VIA_CODE") = rowSOTSVIA1.Item("SHIP_VIA_CODE")
                ' need to update the Ship via Code on the sales order
                ORDR_NO = rowSOTPICK1.Item("ORDR_NO") & String.Empty
                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                If rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("SHIP_VIA_CODE") <> rowSOTSVIA1.Item("SHIP_VIA_CODE") Then
                    Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
                    row.Item("TABLE_NAME") = "SOTORDR1"
                    row.Item("TABLE_KEY") = ORDR_NO
                    row.Item("INIT_DATE") = DateTime.Now
                    row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    row.Item("EVENT_TYPE") = "SHPMTC"
                    row.Item("EVENT_DESC") = "Ship Via was changed from " & rowSOTORDR1.Item("SHIP_VIA_CODE") & " to " & rowSOTSVIA1.Item("SHIP_VIA_CODE")
                    row.Item("EVENT_KEY") = ""
                    dst.Tables("TATEVNT1").Rows.Add(row)
                    rowSOTORDR1.Item("SHIP_VIA_CODE") = rowSOTSVIA1.Item("SHIP_VIA_CODE")
                End If

                If IsDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE")) Then
                    rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = CDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE") & String.Empty).ToShortDateString
                    rowSOTSHIP1.Item("INV_DATE") = CDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE") & String.Empty).ToShortDateString
                    rowSOTSHIP1.Item("SHIPPED_ACTUAL") = CDate(rowEDT945T1.Item("EDI_SHIPMENT_DATE") & String.Empty).ToShortDateString
                Else
                    rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") = CDate(DateTime.Now.ToShortDateString)
                    rowSOTSHIP1.Item("INV_DATE") = CDate(DateTime.Now.ToShortDateString)
                    rowSOTSHIP1.Item("SHIPPED_ACTUAL") = CDate(DateTime.Now.ToShortDateString)
                End If

                rowSOTSHIP1.Item("SHIP_TOTAL_WGT") = EDI_TOTAL_ORDR_WEIGHT

                If rowSOTSHIP1.Item("SHIP_REF") & String.Empty = String.Empty Then
                    If dst.Tables("EDT945T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'", "EDI_SHIPPER_ID_NO").Length > 0 Then
                        rowSOTSHIP1.Item("SHIP_REF") = dst.Tables("EDT945T2").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'", "EDI_SHIPPER_ID_NO")(0).Item("EDI_SHIPPER_ID_NO")
                    End If
                End If

                Dim INV_DATE As Date = rowSOTSHIP1.Item("INV_DATE")
                Dim ORDR_YYYYPP_UPDATED As String = ASCDATA1.GetDataValue("Select MIN(OPS_YYYYPP) from gltparm2 where prd_end_date >= '" & INV_DATE.ToString("dd-MMM-yyyy") & "'") & String.Empty
                If ORDR_YYYYPP_UPDATED.Length = 0 Then
                    ORDR_YYYYPP_UPDATED = ASCMAIN1.CYP
                End If
                rowSOTSHIP1.Item("OPS_YYYYPP") = ORDR_YYYYPP_UPDATED

            Next

            ' Need to set PICK_QTY_CANC
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_QTY_CANC_WSHE > 0")
                rowSOTPICK2.Item("PICK_QTY_CANC") = Val(rowSOTPICK2.Item("PICK_QTY_CANC_WSHE") & String.Empty)

                dst.Tables("SOTPICK2_CANC").ImportRow(rowSOTPICK2)
                Dim ORDR_NOx As String = rowSOTPICK2.Item("ORDR_NO")
                Dim ORDR_LNOx As String = rowSOTPICK2.Item("ORDR_LNO")

                If dst.Tables("SOTORDR2_CANC").Rows.Find(New Object() {ORDR_NOx, ORDR_LNOx}) Is Nothing Then
                    dst.Tables("SOTORDR2_CANC").ImportRow(dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NOx, ORDR_LNOx}))
                End If

                If dst.Tables("SOTORDR1_CANC").Rows.Find(ORDR_NOx) Is Nothing Then
                    dst.Tables("SOTORDR1_CANC").ImportRow(dst.Tables("SOTORDR1").Rows.Find(ORDR_NOx))
                End If
            Next

            ' Look for over shipments
            For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_QTY_CONF > PICK_QTY")
                Dim ORDR_NOx As String = rowSOTPICK2.Item("ORDR_NO")
                Dim ORDR_LNOx As String = rowSOTPICK2.Item("ORDR_LNO")
                Dim PICK_QTY As Int32 = rowSOTPICK2.Item("PICK_QTY")
                Dim PICK_QTY_CONF As Int32 = rowSOTPICK2.Item("PICK_QTY_CONF")

                rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NOx)
                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find({ORDR_NOx, ORDR_LNOx})
                Dim msg As String = $"WARNING: Sales Order {ORDR_NOx}, Line No {ORDR_LNOx}, Item {rowSOTORDR2.Item("ITEM_CODE")} has Pick Qty of {PICK_QTY} and the warehouse shipped {PICK_QTY_CONF}"
                AddError(msg, "1")
            Next

            ASCMAIN1.Progress("Update Pick Ticket Weight", "")
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select()
                ASCMAIN1.Progress("-", rowSOTPICK1.Item("PICK_NO") & String.Empty)
                Dim PICK_TOTAL_WGT As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)
                rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT
                rowSOTPICK1.Item("PICK_CNT_CARTONS") = dst.Tables("SOTCART1").Select("PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'").Length
            Next

            ASCMAIN1.Progress("Create SOTCARTX Master Records", "")
            Dim counter As Int64 = 0
            dst.Tables("SOTCARTX").Rows.Clear()
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTPICK2"), New String() {"PICK_NO", "ORDR_NO", "ORDR_LNO"}).Rows
                counter += 1
                ASCMAIN1.Progress("-", counter)
                dst.Tables("SOTCARTX").Rows.Add(New Object() {row.Item("PICK_NO"), row.Item("ORDR_NO"), row.Item("ORDR_LNO")})
            Next

            ASCMAIN1.Progress("Update Shipment cartons", "")
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("")
                ASCMAIN1.Progress("-", rowSOTSHIP1.Item("SHIP_BOL_NO") & String.Empty)
                rowSOTSHIP1.Item("SHIP_CNT_CARTONS") = Val(dst.Tables("SOTPICK1").Compute("SUM(PICK_CNT_CARTONS_CALC)", "SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'") & String.Empty)
            Next

            ASCMAIN1.Progress("Enforce Constraints Integrity Check - This may take a few minutes.", "")
            EnforceConstraints(True)
            ASCMAIN1.Progress("")

            Return True
        Catch ex As Exception
            MessageBox.Show("Load Shipments Error: " & ex.Message, "Load Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function FillQuantityShipped() As Boolean

        Try
            Dim WHSE_CODE As String = String.Empty
            Dim ITEM_CODE As String = String.Empty
            Dim ORDR_QTY_SHIP As Int32 = 0
            Dim sql As String = String.Empty
            Dim rowICTSTAT2 As DataRow = Nothing

            ASCMAIN1.Progress("Fill Quantity Shipped")

            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH2").Select()
                WHSE_CODE = rowSOTINVH1.Item("WHSE_CODE") & String.Empty
                ITEM_CODE = rowSOTINVH1.Item("ITEM_CODE") & String.Empty
                ORDR_QTY_SHIP = Val(rowSOTINVH1.Item("ORDR_QTY_SHIP") & String.Empty)

                ASCMAIN1.Progress("-", ITEM_CODE)

                rowICTSTAT2 = dst.Tables("ICTSTAT2").Rows.Find(New Object() {ITEM_CODE, WHSE_CODE})

                If rowICTSTAT2 IsNot Nothing Then
                    rowICTSTAT2.Item("QTY_SHIPPED") = Val(rowICTSTAT2.Item("QTY_SHIPPED") & String.Empty) + ORDR_QTY_SHIP
                End If
            Next

            For Each rowICTSTAT2 In dst.Tables("ICTSTAT2").Select("QTY_SHIPPED = 0 ")
                rowICTSTAT2.Delete()
            Next

            dst.Tables("ICTSTAT2").AcceptChanges()

        Catch ex As Exception

        End Try

        Return True

    End Function

    Private Function EvaluateLotNos() As Boolean

        Try
            dst.Tables("LOT_NOS").Rows.Clear()

            If Val(DateTime.Now.ToString("yyyyMMdd")) < Val("20250601") Then
                Return True
            End If

            ASCMAIN1.Progress("Evaluate Lot Nos.", "")
            For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("", "PICK_NO")
                Dim CART_NO As String = rowSOTCART1.Item("CART_NO") & String.Empty
                Dim PICK_NO As String = rowSOTCART1.Item("PICK_NO") & String.Empty
                ASCMAIN1.Progress("Evaluate Lot Nos: " & CART_NO & "/" & PICK_NO, "")

                Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").Rows.Find(PICK_NO)
                If rowSOTPICK1 Is Nothing Then
                    Continue For
                End If
                Dim SHIP_BOL_NO As String = rowSOTPICK1.Item("SHIP_BOL_NO") & String.Empty
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                If rowSOTSHIP1 Is Nothing Then
                    Continue For
                End If

                ' ISSUE-7230 ADS as the defaUlt warehouse
                Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1_LK").Rows.Find(rowSOTSHIP1.Item("WHSE_CODE") & "")

                If drICTWHSE1 Is Nothing OrElse drICTWHSE1.Item("LP_CODE") & String.Empty <> "ADS" Then
                    Continue For
                End If

                For Each rowSOTCART2 As DataRow In dst.Tables("SOTCART2").Select($"CART_NO = '{CART_NO}'")
                    Dim CART_LNO As String = Val(rowSOTCART2.Item("CART_LNO") & String.Empty)
                    Dim QTY_PACKED As String = Val(rowSOTCART2.Item("QTY_PACKED") & String.Empty)
                    Dim ITEM_CODE As String = rowSOTCART2.Item("ITEM_CODE") & String.Empty
                    Dim ORDR_NO As String = rowSOTCART2.Item("ORDR_NO") & String.Empty
                    Dim ORDR_LNO As String = Val(rowSOTCART2.Item("ORDR_LNO") & String.Empty)

                    Dim LOT_QTY As Int32 = Val(dst.Tables("SOTCART3").Compute("SUM(LOT_QTY)", $"CART_NO = '{CART_NO}' AND CART_LNO = {CART_LNO}") & String.Empty)

                    If LOT_QTY <> QTY_PACKED Then
                        Dim rowLOT_NOS As DataRow = dst.Tables("LOT_NOS").NewRow
                        rowLOT_NOS.Item("CART_NO") = CART_NO
                        rowLOT_NOS.Item("CART_LNO") = CART_LNO
                        rowLOT_NOS.Item("PICK_NO") = PICK_NO
                        rowLOT_NOS.Item("ORDER_NO") = ORDR_NO
                        rowLOT_NOS.Item("ORDER_LNO") = ORDR_LNO
                        rowLOT_NOS.Item("ITEM_CODE") = ITEM_CODE
                        rowLOT_NOS.Item("QTY_PACKED") = QTY_PACKED
                        rowLOT_NOS.Item("LOT_QTY") = LOT_QTY
                        dst.Tables("LOT_NOS").Rows.Add(rowLOT_NOS)
                    End If
                Next
            Next

            If dst.Tables("LOT_NOS").Rows.Count > 0 Then
                AddError("WARNING: There are missing Lot Nos.", "1")
            End If

            Return True
        Catch ex As Exception
            MessageBox.Show("Evaluate Lot Nos Error: " & ex.Message, "Evaluate Lot Nos", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    '''  This gets the data ready to update.
    '''  Creates the invoices and updates the shipment, pick and order tables
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ProcessShipments() As Boolean

        Try
            Dim EMsg As String = String.Empty
            Dim rowSOTSHIP1 As DataRow = Nothing
            Dim rowARTCUST1 As DataRow = Nothing
            ASCMAIN1.Progress("Process Shipments")

            ToggleDataTableExpressions(True)
            Dim shipmentsWithErrors As New List(Of String)

            'EDI_TOTAL_UNITS_SHIPPED
            Select Case Company_Code
                Case "AHA"
                    'For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("EDI_PROCESS_IND = '1' and EDI_TOTAL_UNITS_SHIPPED = 0")
                    '    For Each row As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(SELECTED, '0') = '1' and PICK_NO = '" & rowEDT945T1.Item("EDI_PICK_NO") & "'")
                    '        dst.Tables("EDT945T1").Select("EDI_PICK_NO = '" & row.Item("PICK_NO") & "'")(0).Item("EDI_PROCESS_IND") = "0"
                    '        EMsg = "The following shipment did not have any units shipped. Please de-release the shipment: " & row.Item("SHIP_BOL_NO") & ", Order No: " & row.Item("ORDR_NO") & "" _
                    '            & ", Customer: " & row.Item("CUST_CODE") & ", Customer PO: " & row.Item("ORDR_CUST_PO")
                    '        row.Item("SELECTED") = "0"

                    '        If Not shipmentsWithErrors.Contains(row.Item("SHIP_BOL_NO")) Then
                    '            shipmentsWithErrors.Add(row.Item("SHIP_BOL_NO"))
                    '            AddError(EMsg, "1")
                    '        End If
                    '    Next
                    'Next

                    'For Each row As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(SELECTED, '0') = '1' and PICK_QTY_CONF = 0")
                    '    dst.Tables("EDT945T1").Select("EDI_PICK_NO = '" & row.Item("PICK_NO") & "'")(0).Item("EDI_PROCESS_IND") = "0"
                    '    EMsg = "Cannot Update when nothing is confirmed as shipped for shipment: " & row.Item("SHIP_BOL_NO") & ", Order No: " & row.Item("ORDR_NO") & "" _
                    '        & ", Customer: " & row.Item("CUST_CODE") & ", Customer PO: " & row.Item("ORDR_CUST_PO")
                    '    row.Item("SELECTED") = "0"

                    '    If Not shipmentsWithErrors.Contains(row.Item("SHIP_BOL_NO")) Then
                    '        shipmentsWithErrors.Add(row.Item("SHIP_BOL_NO"))
                    '        AddError(EMsg, "1")
                    '    End If
                    'Next

                    For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("EDI_PROCESS_IND = '1' and EDI_TOTAL_UNITS_SHIPPED = 0", "EDI_SHIPMENT_ID,EDI_PICK_NO")
                        Dim row As DataRow = dst.Tables("SOTPICK1").Select("ISNULL(SELECTED, '0') = '1' and PICK_NO = '" & rowEDT945T1.Item("EDI_PICK_NO") & "'")(0)

                        EMsg = "WARNING: The following Pick Ticket (" & row.Item("PICK_NO") & ") does not have any units shipped. You may need to de-release shipment: " & row.Item("SHIP_BOL_NO") _
                            & ", Order No: " & row.Item("ORDR_NO") & "" _
                            & ", Customer: " & row.Item("CUST_CODE") _
                            & ", Customer PO: " & row.Item("ORDR_CUST_PO")
                        'row.Item("PICK_STATUS") = "C"
                        AddError(EMsg, "1")
                    Next

                    For Each row As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(SELECTED, '0') = '1' and PICK_QTY_CONF = 0", "SHIP_BOL_NO,PICK_NO")
                        EMsg = "WARNING: The following Pick Ticket (" & row.Item("PICK_NO") & ") does not have any units shipped. You may need to de-release shipment: " & row.Item("SHIP_BOL_NO") _
                            & ", Order No: " & row.Item("ORDR_NO") & "" _
                            & ", Customer: " & row.Item("CUST_CODE") _
                            & ", Customer PO: " & row.Item("ORDR_CUST_PO")
                        'row.Item("PICK_STATUS") = "C"
                        AddError(EMsg, "1")
                    Next

                    Dim tbl As DataTable = ASCDATA1.SelectDistinct("EDT945T1", New String() {"EDI_SHIPMENT_ID"})
                    For Each row As DataRow In tbl.Select("", "EDI_SHIPMENT_ID")
                        Dim SHIP_BOL_NO As String = row.Item("EDI_SHIPMENT_ID") & String.Empty
                        If dst.Tables("EDT945T1").Select("EDI_PROCESS_IND = '1' and EDI_SHIPMENT_ID = '" & SHIP_BOL_NO & "' AND EDI_TOTAL_UNITS_SHIPPED > 0").Length = 0 Then
                            EMsg = "WARNING: Shipment (" & SHIP_BOL_NO & ") does not have any units shipped. You may need to de-release the shipment."
                            AddError(EMsg, "1")

                            If Not shipmentsWithErrors.Contains(SHIP_BOL_NO) Then
                                'shipmentsWithErrors.Add(SHIP_BOL_NO)
                            End If

                            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(SELECTED, '0') = '1' and SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                                rowSOTPICK1.Item("SELECTED") = "1"
                            Next
                        End If
                    Next
            End Select

            For Each row As DataRow In dst.Tables("SOTPICK1").Select("ISNULL(SELECTED, '0') = '1' and PICK_QTY_CONF = 0 and (PICK_CNT_CARTONS <> 0 or PICK_TOTAL_WGT <> 0)")
                dst.Tables("EDT945T1").Select("EDI_PICK_NO = '" & row.Item("PICK_NO") & "'")(0).Item("EDI_PROCESS_IND") = "0"
                EMsg = "Some Pick Tickets have 0 qty confirmed as Shipped - but Still have a non-Zero value for cartons or weight for shipment: " & row.Item("SHIP_BOL_NO") _
                    & ", Order No: " & row.Item("ORDR_NO") & "" _
                    & ", Customer: " & row.Item("CUST_CODE") & ", Customer PO: " & row.Item("ORDR_CUST_PO")

                If Not shipmentsWithErrors.Contains(row.Item("SHIP_BOL_NO")) Then
                    shipmentsWithErrors.Add(row.Item("SHIP_BOL_NO"))
                    AddError(EMsg, "1")
                End If
            Next

            ' although this only matters for edi customers, I think we should enforce the integrity
            Dim rowSOTCARTX_oobal As DataRow() = dst.Tables("SOTCARTX").Select("ISNULL(PICK_QTY_CONF,0) <> ISNULL(QTY_PACKED,0)")
            If rowSOTCARTX_oobal.Length <> 0 Then
                For iLoop As Int16 = 0 To rowSOTCARTX_oobal.Count - 1
                    EMsg = "Pick Ticket Detail Qty Confirmed out of balance with Carton Details"
                    EMsg &= " (See Pick Ticket " & rowSOTCARTX_oobal(iLoop).Item("PICK_NO") & ", Line " & rowSOTCARTX_oobal(iLoop).Item("ORDR_LNO") & " Confirmed " _
                        & Val(rowSOTCARTX_oobal(iLoop).Item("PICK_QTY_CONF") & String.Empty) & " Packed " & Val(rowSOTCARTX_oobal(iLoop).Item("QTY_PACKED") & String.Empty) & ")"
                    AddError(EMsg)
                Next
            End If

            Dim lstEDI_SHIPMENT_ID As New List(Of String)

            ASCMAIN1.Progress("Validate EDI BOL data")

            Dim tblEDT945T1 As DataTable = ASCDATA1.SelectDistinct(dst.Tables("EDT945T1"), New String() {"EDI_SHIPMENT_ID", "EDI_PROCESS_IND"})
            'For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("EDI_PROCESS_IND = '1'", "EDI_DOC_SEQ_NO")

            For Each row As DataRow In tblEDT945T1.Select("EDI_PROCESS_IND = '1'", "EDI_SHIPMENT_ID")
                Dim rowEDT945T1 As DataRow = dst.Tables("EDT945T1").Select("EDI_SHIPMENT_ID = '" & row.Item("EDI_SHIPMENT_ID") & "' and EDI_PROCESS_IND = '1'")(0)

                If lstEDI_SHIPMENT_ID.Contains(rowEDT945T1.Item("EDI_SHIPMENT_ID")) Then
                    Continue For
                End If
                ASCMAIN1.Progress("Ship Bol No: " & rowEDT945T1.Item("EDI_SHIPMENT_ID"), "")
                lstEDI_SHIPMENT_ID.Add(rowEDT945T1.Item("EDI_SHIPMENT_ID"))
                For Each rowSOTSHIP1 In dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & rowEDT945T1.Item("EDI_SHIPMENT_ID") & "'", "CUST_CODE")
                    rowSOTSHIP1.Item("SELECTED") = "1"
                    rowSOTSHIP1.Item("BILL_OF_LADING_NO") = rowEDT945T1.Item("EDI_BOL_NO")
                    ASCMAIN1.Progress("-", rowEDT945T1.Item("EDI_BOL_NO") & String.Empty)

                    If rowARTCUST1 Is Nothing OrElse rowSOTSHIP1.Item("CUST_CODE") <> rowARTCUST1.Item("CUST_CODE") Then
                        rowARTCUST1 = LookUp("ARTCUST1", rowSOTSHIP1.Item("CUST_CODE"))
                    End If

                    If Not IsDate(rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") & String.Empty) Then
                        EMsg = "Date Shipped missing for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    End If

                    If Not IsDate(rowSOTSHIP1.Item("INV_DATE") & String.Empty) Then
                        EMsg = "Invoice Date missing for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    End If

                    ' Allow Warerhouse to send previous month shipments within 5 days of the new month
                    If Format(rowSOTSHIP1.Item("INV_DATE"), "yyyyMM") < Format(validDates(0), "yyyyMM") Then
                        If Val(DateTime.Now.ToString("dd")) <= 5 Then
                            EMsg = $"Invoice Date of {rowSOTSHIP1.Item("INV_DATE")} was changed to {CDate(DateTime.Now.ToString("MM/dd/yyyy"))} for Shipment: {rowSOTSHIP1.Item("SHIP_BOL_NO")}, Order Group: {rowSOTSHIP1.Item("ORDR_GROUP_NO")}, Customer: {rowSOTSHIP1.Item("CUST_CODE")}"
                            AddError(EMsg, "1")
                            rowSOTSHIP1.Item("INV_DATE") = CDate(DateTime.Now.ToString("MM/dd/yyyy"))
                        End If
                    End If

                    If IsDate(rowSOTSHIP1.Item("SHIP_DATE_SHIPPED") & String.Empty) _
                        AndAlso IsDate(rowSOTSHIP1.Item("INV_DATE") & String.Empty) _
                        AndAlso Format(rowSOTSHIP1.Item("SHIP_DATE_SHIPPED"), "yyyyMMdd") _
                            > Format(rowSOTSHIP1.Item("INV_DATE"), "yyyyMMdd") Then
                        EMsg = "Invoice Date is Prior to Date Shipped for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    End If

                    ' Ship Date must be in current or next period
                    If Format(rowSOTSHIP1.Item("INV_DATE"), "yyyyMM") < Format(validDates(0), "yyyyMM") _
                        OrElse Format(rowSOTSHIP1.Item("INV_DATE"), "yyyyMM") > Format(validDates(1), "yyyyMM") Then
                        EMsg = "Invoice Date must be between " & validDates(0) & " and " & validDates(1) & " for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    End If

                    If rowSOTSHIP1.Item("TERM_CODE") & String.Empty = "" Then
                        EMsg = "Terms Code is Required for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    Else
                        If LookUp("TATTERM1", rowSOTSHIP1.Item("TERM_CODE")) Is Nothing Then
                            EMsg = "Invalid Terms Code for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                            AddError(EMsg, "1")
                            rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                            Continue For
                        End If
                    End If

                    If rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty = "" Then
                        EMsg = "Ship Via Code is Required for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg)
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    Else
                        If LookUp("SOTSVIA1", rowSOTSHIP1.Item("SHIP_VIA_CODE")) Is Nothing Then
                            EMsg = "Invalid Ship Via Code for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                            AddError(EMsg, "1")
                            rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                            Continue For
                        End If
                    End If

                    If rowSOTSHIP1.Item("FRT_TERMS") & String.Empty = "" Then
                        EMsg = "Frt Terms Code is Required for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    Else
                        If LookUp("ASTCODE1", New String() {"ARTCUST1", "FRT_TERMS", rowSOTSHIP1.Item("FRT_TERMS")}) Is Nothing Then
                            EMsg = "Invalid Frt Terms Code for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                            AddError(EMsg, "1")
                            rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                            Continue For
                        End If
                    End If

                    If rowSOTSHIP1.Item("SREP_CODE") & String.Empty <> "" AndAlso LookUp("SOTSREP1", rowSOTSHIP1.Item("SREP_CODE")) Is Nothing Then
                        EMsg = "Invalid Sales Rep Code for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    End If

                    If rowSOTSHIP1.Item("SREP2_CODE") & String.Empty <> "" AndAlso LookUp("SOTSREP1", rowSOTSHIP1.Item("SREP2_CODE")) Is Nothing Then
                        EMsg = "Invalid Sales Rep2 Code for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                        AddError(EMsg, "1")
                        rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                        Continue For
                    End If

                    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'")
                        If dst.Tables("SOTPICK2").Select("PICK_QTY_CANC <> 0 AND PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'").Length <> 0 Then
                            If rowSOTSHIP1.Item("REASON_CODE") & String.Empty = "" Then
                                Select Case Company_Code
                                    Case "INT"
                                        rowSOTSHIP1.Item("REASON_CODE") = "OOS"
                                    Case "AHA"
                                        rowSOTSHIP1.Item("REASON_CODE") = "OOS"
                                End Select
                            End If
                        End If

                        If dst.Tables("SOTPICK2").Select("PICK_QTY_CONF <> 0 AND PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'").Length <> 0 Then
                            If rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty = "" Then
                                EMsg = "Ship Via Code is Required for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                                AddError(EMsg, "1")
                                rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                                Continue For
                            Else
                                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty)
                                If rowSOTSVIA1 Is Nothing Then
                                    EMsg = "Invalid Ship Via Code for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                                    AddError(EMsg, "1")
                                    rowEDT945T1.Item("EDI_PROCESS_IND") = "0"
                                    Continue For
                                Else
                                    ' Scac Code is in the ASN file sent from ADS
                                    'If rowSOTSHIP1.Item("SHIP_856_IND") & String.Empty = "1" Then
                                    '    If rowSOTSVIA1.Item("SHIP_VIA_SCAC") & "" = "" Then
                                    '        EMsg = "Selected Shipper Requires SCAC Code For EDI Customers for Shipment: " & rowSOTSHIP1.Item("SHIP_BOL_NO")
                                    '        AddError(EMsg)
                                    '    End If
                                    'End If
                                End If
                            End If
                        End If
                        If rowEDT945T1.Item("EDI_PROCESS_IND") = "0" Then Exit For
                    Next
                    If rowEDT945T1.Item("EDI_PROCESS_IND") = "0" Then Exit For
                Next
            Next

            For Each rowEDT945T1 As DataRow In dst.Tables("EDT945T1").Select("EDI_PROCESS_IND = '0'", "EDI_DOC_SEQ_NO")
                Dim EDI_SHIPMENT_ID As String = rowEDT945T1.Item("EDI_SHIPMENT_ID")
                ASCMAIN1.Progress("Ship Bol No: " & rowEDT945T1.Item("EDI_SHIPMENT_ID"), "")

                If Not shipmentsWithErrors.Contains(EDI_SHIPMENT_ID) Then
                    EMsg = "Shipment (" & EDI_SHIPMENT_ID & ") was not process since at least one Pick Ticket on the Shipment could not be processed."
                    shipmentsWithErrors.Add(EDI_SHIPMENT_ID)
                    AddError(EMsg, "1")
                End If

                ' set all pick tickets for the Shipment to not process.
                For Each row As DataRow In dst.Tables("EDT945T1").Select("EDI_SHIPMENT_ID = '" & EDI_SHIPMENT_ID & "'")
                    row.Item("EDI_PROCESS_IND") = "0"
                Next

                ' Reject all changes to the Pick Tickets and Sales Order tables
                For Each rowSOTSHIP1X As DataRow In dst.Tables("SOTSHIP1").Select("SHIP_BOL_NO = '" & EDI_SHIPMENT_ID & "'")
                    rowSOTSHIP1X.Item("SELECTED") = "0"
                    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & EDI_SHIPMENT_ID & "'")
                        For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'")
                            rowSOTPICK2.RejectChanges()
                        Next
                        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("ORDR_NO = '" & rowSOTPICK1.Item("ORDR_NO") & "'")
                            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & rowSOTORDR1.Item("ORDR_NO") & "'")
                                rowSOTORDR2.RejectChanges()
                            Next
                            rowSOTORDR1.RejectChanges()
                        Next
                        rowSOTPICK1.RejectChanges()
                        rowSOTPICK1.Item("SELECTED") = "0"
                    Next
                Next
            Next

            If dst.Tables("EDT945T1").Select("EDI_PROCESS_IND = '1'").Length = 0 Then
                Return False
            End If

            ' Interparfums receives cancelled orders from the warehouse. They will be cancelled here.
            If Company_Code = "INT" Then
                For Each rowSOTSHIP1X As DataRow In dst.Tables("SOTSHIP1").Select("SELECTED = '1' AND SHIP_STATUS = 'P'", "SHIP_BOL_NO")
                    Dim SHIP_BOL_NO As String = rowSOTSHIP1X.Item("SHIP_BOL_NO") & String.Empty
                    Dim cancelShipment As Boolean = True

                    For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'")
                        Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")

                        Dim qtyConfirmed As Int64 = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY_CONF)", "PICK_NO = '" & PICK_NO & "'") & String.Empty)
                        If qtyConfirmed > 0 Then
                            cancelShipment = False
                        Else
                            rowSOTPICK1.Item("PICK_STATUS") = "C"
                        End If
                    Next

                    If cancelShipment Then
                        rowSOTSHIP1X.Item("SHIP_STATUS") = "C"
                    End If
                Next
            End If

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_STATUS = 'C'")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select("PICK_NO = '" & PICK_NO & "'")
                    rowSOTPICK2.Item("PICK_QTY_CONF") = 0
                    rowSOTPICK2.Item("PICK_QTY_CANC") = Val(rowSOTPICK2.Item("PICK_QTY") & String.Empty)
                Next
            Next

            sql = "(BILL_OF_LADING_NO is Null or BILL_OF_LADING_NO = '') AND SELECTED = '1' AND SHIP_STATUS = 'P'"
            For Each row As DataRow In dst.Tables("SOTSHIP1").Select(sql)
                AddError("The following Shipment is missing the Bill of Lading No. Customer :" & row.Item("CUST_CODE") & ", PO NO: " & row.Item("ORDR_CUST_PO") & " ABS Ship BOL: " & row.Item("SHIP_BOL_NO"))
            Next

            If dst.Tables("SOTSHIP1").Select(sql).Length > 0 Then
                Return False
            End If

            ' Create Invoices and update Pick Tickets and Sales Orders
            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTSHIP1").Select("SELECTED = '1' AND SHIP_STATUS = 'P'"), New String() {"BILL_OF_LADING_NO"}).Rows
                Dim BILL_OF_LADING_NO As String = row.Item("BILL_OF_LADING_NO") & String.Empty
                ASCMAIN1.Progress("Bill of Lading No: " & BILL_OF_LADING_NO, "")

                rowSOTSHIP1 = dst.Tables("SOTSHIP1").Select("BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "' AND SELECTED = '1'")(0)
                Dim CUST_CODE As String = rowSOTSHIP1.Item("CUST_CODE")
                Dim ORDR_NO As String = dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "' and PICK_STATUS = 'P'")(0).Item("ORDR_NO")

                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                Dim rowSOTSHIPB As DataRow = dst.Tables("SOTSHIPB").NewRow
                With rowSOTSHIPB
                    .Item("BOL_NO") = BILL_OF_LADING_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("BOL_DATE") = DateTime.Now.ToShortDateString
                    .Item("FRT_TERMS") = rowSOTSHIP1.Item("FRT_TERMS")
                    .Item("WHSE_CODE") = rowSOTSHIP1.Item("WHSE_CODE")
                    .Item("MASTER_BOL_NO") = String.Empty
                    .Item("MASTER_BOL") = "0"

                    Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", rowSOTSHIP1.Item("SHIP_VIA_CODE"))
                    .Item("SHIP_VIA_CODE") = rowSOTSHIP1.Item("SHIP_VIA_CODE")
                    .Item("SHIP_VIA_DESC") = rowSOTSVIA1.Item("SHIP_VIA_DESC") & String.Empty
                    .Item("SHIP_VIA_SCAC") = rowSOTSVIA1.Item("SHIP_VIA_SCAC") & String.Empty

                    Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO & "' AND CUST_ADDR_TYPE = 'ST'")(0)
                    .Item("SHIP_TO_NAME") = rowSOTORDR5.Item("CUST_NAME") & String.Empty
                    .Item("SHIP_TO_ADDR1") = rowSOTORDR5.Item("CUST_ADDR1") & String.Empty
                    .Item("SHIP_TO_ADDR2") = rowSOTORDR5.Item("CUST_ADDR2") & String.Empty
                    .Item("SHIP_TO_ADDR3") = rowSOTORDR5.Item("CUST_ADDR3") & String.Empty
                    .Item("SHIP_TO_CITY") = rowSOTORDR5.Item("CUST_CITY") & String.Empty
                    .Item("SHIP_TO_STATE") = rowSOTORDR5.Item("CUST_STATE") & String.Empty
                    .Item("SHIP_TO_ZIP_CODE") = rowSOTORDR5.Item("CUST_ZIP_CODE") & String.Empty
                    .Item("SHIP_TO_COUNTRY") = rowSOTORDR5.Item("CUST_COUNTRY") & String.Empty
                    .Item("SHIP_TO_CONTACT") = rowSOTORDR5.Item("CUST_CONTACT") & String.Empty
                    .Item("SHIP_TO_PHONE") = rowSOTORDR5.Item("CUST_PHONE") & String.Empty

                    .Item("FRT_3PY_NAME") = String.Empty
                    .Item("FRT_3PY_ADDR1") = String.Empty
                    .Item("FRT_3PY_ADDR2") = String.Empty
                    .Item("FRT_3PY_ADDR3") = String.Empty
                    .Item("FRT_3PY_CITY") = String.Empty
                    .Item("FRT_3PY_STATE") = String.Empty
                    .Item("FRT_3PY_ZIP_CODE") = String.Empty
                    .Item("FRT_3PY_COUNTRY") = String.Empty
                    .Item("FRT_3PY_CONTACT") = String.Empty
                    .Item("FRT_3PY_PHONE") = String.Empty

                    .Item("BOL_INST") = rowARTCUST1.Item("CUST_BOL_INST")
                    .Item("THIRD_PARTY") = "0"
                    .Item("SHIP_REF") = rowSOTSHIP1.Item("SHIP_REF")
                    .Item("SHIP_TRAILER_NO") = rowSOTSHIP1.Item("SHIP_TRAILER_NO")
                    .Item("SHIP_SEAL_NO") = rowSOTSHIP1.Item("SHIP_SEAL_NO")
                    .Item("BOL_STATUS") = "F"
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = .Item("INIT_DATE")
                    .Item("LAST_OPER") = .Item("INIT_OPER")
                    .Item("SHIPPED_ACTUAL") = rowSOTSHIP1.Item("SHIPPED_ACTUAL")
                    .Item("SHIP_TO_CODE") = String.Empty
                    .Item("FRT_3PY_CODE") = String.Empty
                    '.Item("BOL_PRINTED") = String.Empty
                    .Item("SHIP_LOAD_NO") = rowSOTSHIP1.Item("SHIP_LOAD_NO") & String.Empty
                    .Item("SHIP_APPT_NO") = rowSOTSHIP1.Item("SHIP_APPT_NO") & String.Empty

                    'SHIP_PICKUP_NO, SHIP_AUTH_NO
                    .Item("SHIP_PICKUP_NO") = rowSOTSHIP1.Item("SHIP_PICKUP_NO") & String.Empty
                    .Item("SHIP_AUTH_NO") = rowSOTSHIP1.Item("SHIP_AUTH_NO") & String.Empty

                    ' .Item("SCHED_DELIV_DATE") = String.Empty
                    ' .Item("SHIP_FREIGHT") = String.Empty
                    dst.Tables("SOTSHIPB").Rows.Add(rowSOTSHIPB)
                End With

                UpdateOrderPicksAndCreateInvoices(BILL_OF_LADING_NO)
                rowSOTSHIPB.Item("SHIP_FOB") = dst.Tables("SOTSHIP1").Select("BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "' and SELECTED = '1'", "", DataViewRowState.CurrentRows)(0).Item("SHIP_FOB") & String.Empty
            Next

            Return True
        Catch ex As Exception
            MessageBox.Show("Process Shipments Error: " & ex.Message, "Process Shipments", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Public Overrides Sub Update_Record()
        MyBase.Update_Record()

        ASCMAIN1.Progress("Shipment BOL", "")
        For Each rowSOTSHIPB As DataRow In dst.Tables("SOTSHIPB").Select()
            rowSOTSHIPB.Item("BOL_STATUS") = "F"
            Dim BILL_OF_LADING_NO As String = rowSOTSHIPB.Item("BOL_NO")
            ASCMAIN1.Progress("-", BILL_OF_LADING_NO)
            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "' and SELECTED = '1'")
                rowSOTSHIP1.Item("SHIP_STATUS") = "F"
                rowSOTSHIP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTSHIP1.Item("LAST_DATE") = DateTime.Now
            Next
        Next

        ' Update the tables
        ASCMAIN1.Progress("Shipment No", "")
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("SELECTED = '1'")

            ASCMAIN1.Progress("-", rowSOTSHIP1.Item("SHIP_BOL_NO"))
            'SHIP_856_IND, SHIP_810_IND
            If rowSOTSHIP1.Item("SHIP_856_IND") & String.Empty = "1" OrElse rowSOTSHIP1.Item("SHIP_810_IND") & String.Empty = "1" Then
                Continue For
            End If

            If rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty = String.Empty Then
                Continue For
            End If

            Select Case Company_Code
                Case "INT"
                    Dim BILL_OF_LADING_NO As String = rowSOTSHIP1.Item("BILL_OF_LADING_NO") & String.Empty
                    Dim WHSE_CODE As String = rowSOTSHIP1.Item("WHSE_CODE") & String.Empty

                    ' ISSUE-7230 ADS as the default warehouse
                    Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1_LK").Rows.Find(WHSE_CODE)
                    If drICTWHSE1 Is Nothing OrElse drICTWHSE1.Item("LP_CODE") & String.Empty <> "ADS" Then
                        'If WHSE_CODE <> "ADS" Then
                        If BILL_OF_LADING_NO.Length = 7 Then
                            BILL_OF_LADING_NO = "0677385000" & BILL_OF_LADING_NO

                            ' 2024-04-16
                            ' EDI Orders and Keyed in orders cause an issue since the EDI already contain the prefix 0677385000
                            If dst.Tables("SOTSHIPB").Select($"BOL_NO = '{BILL_OF_LADING_NO}'").Length = 0 Then
                                For Each rowSOTSHIPB As DataRow In dst.Tables("SOTSHIPB").Select("BOL_NO = '" & rowSOTSHIP1.Item("BILL_OF_LADING_NO") & "'", "", DataViewRowState.CurrentRows)
                                    rowSOTSHIPB.Item("BOL_NO") = BILL_OF_LADING_NO
                                Next
                            End If

                            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("INV_BOL_NO = '" & rowSOTSHIP1.Item("BILL_OF_LADING_NO") & "'", "", DataViewRowState.CurrentRows)
                                rowSOTINVH1.Item("INV_BOL_NO") = BILL_OF_LADING_NO
                            Next

                            rowSOTSHIP1.Item("BILL_OF_LADING_NO") = BILL_OF_LADING_NO
                        End If
                    End If
            End Select
        Next

        INIT_LAST("SOTSHIPB", False, , True)
        ASCMAIN1.Progress("Updating", "SOTSHIPB")

        'Set records to Modified
        For Each bol As String In bolList
            Dim rowSOTSHIPB As DataRow = dst.Tables("SOTSHIPB").Rows.Find(bol)
            If rowSOTSHIPB IsNot Nothing Then
                rowSOTSHIPB.AcceptChanges()
                rowSOTSHIPB.SetModified()
            End If
        Next


        ' Used try to prevent a query string error.
        ' If this fails and there is a duplicate BOL NO the error will get trapped in  Update_Record_TDA("SOTSHIPB")
        ' This lets us see the BOL No
        Dim lstErrors As New List(Of String)
        Try
            Dim lstBols As New List(Of String)
            For Each row As DataRow In dst.Tables("SOTSHIPB").Select("")
                Dim BOL_NO As String = row.Item("BOL_NO") & String.Empty
                If Not lstBols.Contains(BOL_NO) Then
                    lstBols.Add(BOL_NO)
                End If
            Next

            If lstBols.Count > 0 Then
                Dim wkTBL As String = ASCMAIN1.Temp_Table("SELECT BOL_NO FROM SOTSHIPB WHERE ROWNUM < 1")
                For Each BOL_NO As String In lstBols
                    ASCMAIN1.sql = $"INSERT INTO {wkTBL} VALUES (:PARM1)"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {BOL_NO})
                Next

                Dim sql As String = $"SELECT * FROM SOTSHIPB WHERE BOL_NO IN (SELECT BOL_NO FROM {wkTBL})"
                Dim tbl As DataTable = ASCDATA1.GetDataTable(sql)
                If tbl.Rows.Count > 0 Then
                    For Each row As DataRow In tbl.Select("", "BOL_NO")
                        lstErrors.Add($"Bol No: {row.Item("BOL_NO")}, Customer {row.Item("CUST_CODE")}, Bol Date {CDate(row.Item("SHIPPED_ACTUAL") & String.Empty).ToShortDateString}")
                    Next
                End If
            End If
        Catch ex As Exception
        End Try

        If lstErrors.Count > 0 Then
            Throw New Exception("The following BOLS were already used. History needs to be modified before you can proceed." & Environment.NewLine & Environment.NewLine & String.Join(Environment.NewLine, lstErrors.ToArray))
        End If

        Update_Record_TDA("SOTSHIPB")

        INIT_LAST("SOTSHIP1", False, , True)

        ASCMAIN1.Progress("Updating", "SOTSHIP1")
        Update_Record_TDA("SOTSHIP1")

        ' Remove existing Cartons from SOTCART1/2
        ASCMAIN1.sql = "Delete from SOTCART2 WHERE CART_NO IN (SELECT CART_NO from SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & tblSOTCART1 & "))"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Delete from SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & tblSOTCART1 & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Delete from SOTCART3 WHERE CART_NO IN (SELECT CART_NO from SOTCART1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & tblSOTCART1 & "))"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.Progress("Updating", "SOTCART1")
        Update_Record_TDA("SOTCART1")

        ASCMAIN1.Progress("Updating", "SOTCART2")
        Update_Record_TDA("SOTCART2")

        ASCMAIN1.Progress("Updating", "SOTCART3")
        Update_Record_TDA("SOTCART3")

        ' Need to make sure PT's have weights
        Dim lstPICK_NOS As New List(Of String)
        Dim lstCART_TRACKING_NOS As New List(Of String)


        ASCMAIN1.Progress("Updating Pick Ticket", "")
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select()
            Dim PICK_TOTAL_WGT As Decimal = Val(dst.Tables("SOTCART1").Compute("SUM(CART_TOTAL_WGT_ACTUAL)", "PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'") & String.Empty)
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = PICK_TOTAL_WGT
            ASCMAIN1.Progress("-", rowSOTPICK1.Item("PICK_NO"))
            lstPICK_NOS.Add(rowSOTPICK1.Item("PICK_NO"))
        Next

        ASCMAIN1.Progress("Updating", "ARTCCPAC")
        Update_Record_TDA("ARTCCPAC")

        ASCMAIN1.Progress("Updating", "EDT945T1")
        Update_Record_TDA("EDT945T1")

        ASCMAIN1.Progress("Updating", "SOTORDR1")
        Update_Record_TDA("SOTORDR1")

        ASCMAIN1.Progress("Updating", "SOTORDR2")
        Update_Record_TDA("SOTORDR2")

        ASCMAIN1.Progress("Updating", "TATEVNT1")
        Update_Record_TDA("TATEVNT1")

        ASCMAIN1.Progress("Updating", "SOTPICK1")
        Update_Record_TDA("SOTPICK1")

        ASCMAIN1.Progress("Updating", "SOTPICK2")
        Update_Record_TDA("SOTPICK2")

        ASCMAIN1.Progress("Updating SOTINVH1", "")

        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select()
            ' Need to set the invoices as finalized - Bristol no longer does end of day processing.
            If rowSOTINVH1.Item("ORDR_WEB_IND") & String.Empty = "1" Then
                rowSOTINVH1.Item("ORDR_WEB_IND") = "2"
            End If
            ASCMAIN1.Progress("-", rowSOTINVH1.Item("INV_NO"))
        Next

        ASCMAIN1.Progress("Updating", "SOTINVH1")
        Update_Record_TDA("SOTINVH1")

        ASCMAIN1.Progress("Updating", "SOTINVH2")
        Update_Record_TDA("SOTINVH2")

        ASCMAIN1.Progress("Updating", "ARTOPEN1")
        Update_Record_TDA("ARTOPEN1")

        '  UPDATE SOTSHIPT SET CART_NO
        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS
                            SELECT SOTSHIPT.SHIP_BOL_NO, SOTSHIPT.TRACKING_NO, SOTSHIPT.PICK_NO, SOTCART1.CART_NO
                                FROM SOTSHIPT, SOTCART1
                                WHERE SOTSHIPT.PICK_NO = SOTCART1.PICK_NO
                                AND SOTSHIPT.TRACKING_NO = SOTCART1.CART_TRACKING_NO
                                AND SOTSHIPT.CART_NO IS NULL;
                            BEGIN FOR R1 IN C1 LOOP
                                UPDATE SOTSHIPT SET CART_NO = R1.CART_NO WHERE SOTSHIPT.SHIP_BOL_NO = R1.SHIP_BOL_NO AND TRACKING_NO = R1.TRACKING_NO AND PICK_NO = R1.PICK_NO;
                            END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ' UPDATE SOTINVH1 SET ORDR_YYYYPP_DEL
        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS
                            SELECT SOTINVH1.INV_NO
                                FROM SOTSHIPT, SOTPICK1, SOTINVH1, SOTORDR1
                                WHERE SOTSHIPT.PICK_NO = SOTPICK1.PICK_NO
                                AND SOTPICK1.INV_NO = SOTINVH1.INV_NO
                                AND SOTINVH1.ORDR_YYYYPP_DEL IS NULL
                                AND SOTSHIPT.DELIVERY_DATE IS NOT NULL
                                AND SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO
                                AND SOTORDR1.FRT_TERMS = 'DEL';
                            BEGIN FOR R1 IN C1 LOOP
                                UPDATE SOTINVH1 SET ORDR_YYYYPP_DEL = :PARM1 WHERE INV_TYPE = 'I' AND INV_NO = R1.INV_NO AND ORDR_YYYYPP_DEL IS NULL;
                            END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", {ASCMAIN1.CYP})

        ' UPDATE SOTCART1 SET DELIVERY_DATE
        ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS
                            SELECT SOTSHIPT.TRACKING_NO, SOTSHIPT.PICK_NO, SOTCART1.CART_NO, SOTSHIPT.DELIVERY_DATE
                                FROM SOTSHIPT, SOTCART1
                                WHERE SOTSHIPT.PICK_NO = SOTCART1.PICK_NO
                                AND SOTSHIPT.TRACKING_NO = SOTCART1.CART_TRACKING_NO
                                AND SOTCART1.DELIVERY_DATE IS NULL
                                AND SOTSHIPT.DELIVERY_DATE IS NOT NULL;
                            BEGIN FOR R1 IN C1 LOOP
                                UPDATE SOTCART1 SET DELIVERY_DATE = R1.DELIVERY_DATE WHERE PICK_NO = R1.PICK_NO AND CART_TRACKING_NO = R1.TRACKING_NO AND DELIVERY_DATE IS NULL;
                            END LOOP; END; END;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = String.Empty

        ' Update Ahava UPS Tracking Numbers
        If Company_Code = "AHA" Then
            ' These have the same tracking Numbers. So flag so they do not get processed
            sql = " BEGIN DECLARE CURSOR C1 IS	"
            sql &= " 	  Select SOTORDR1.ORDR_NO, SOTPICK1.PICK_NO, SOTCART1.CART_NO, SOTCART1.CART_TRACKING_NO, SOTPICK1.INV_NO, SOTPICK1.SHIP_BOL_NO, SOTTRAC2.*"
            sql &= " 	  FROM SOTORDR1, SOTPICK1, SOTCART1, SOTTRAC2"
            sql &= " 	  WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO "
            sql &= " 	  AND SOTCART1.PICK_NO = SOTPICK1.PICK_NO "
            sql &= " 	  AND SOTORDR1.ORDR_STATUS= 'F' "
            sql &= " 	  AND SOTPICK1.PICK_STATUS = 'F' "
            sql &= " 	  AND NVL(SOTTRAC2.PROCESS_IND, '0') = '0'"
            sql &= " 	  AND SOTTRAC2.ORDR_NO_WEB = SOTORDR1.ORDR_CUST_PO"
            sql &= " 	  AND SOTTRAC2.TRACKING_NO = SOTCART1.CART_TRACKING_NO;  "
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= " 	  UPDATE SOTTRAC2 SET PROCESS_IND = 'S' WHERE TRACK_NO = R1.TRACK_NO AND TRACK_LNO = R1.TRACK_LNO;"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)

        End If

        ASCMAIN1.Progress("Updating Inventory", "")
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Rows

            ASCMAIN1.Progress("-", rowSOTINVH1.Item("INV_NO"))

            ASCMAIN1.sql = "BEGIN SOPSTAT1('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "BEGIN SOPSTAT2('" & rowSOTINVH1.Item("INV_TYPE") & "','" & rowSOTINVH1.Item("INV_NO") & "'); END;"
            ASCDATA1.ExecuteSQL()

            ' Is this needed
            'TAC.ICCMAIN1.Update_WHTLOCBX("S", rowSOTINVH1.Item("INV_NO"))

            ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV",
                   New Object() {rowSOTINVH1.Item("INV_TYPE"), rowSOTINVH1.Item("INV_NO")},
                   New String() {"INV_TYPE_IN", "INV_NO_IN"})

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("INV_NO = '" & rowSOTINVH1.Item("INV_NO") & "' AND PICK_QTY_CANC > 0")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                sql = "BEGIN DECLARE CURSOR C2 IS" _
                    & " SELECT SOTPICK2.*, SOTORDR2.ITEM_CODE, SOTORDR2.WHSE_CODE " _
                    & " FROM SOTPICK2, SOTORDR2" _
                    & " WHERE SOTPICK2.PICK_NO = '" & PICK_NO & "' AND SOTPICK2.PICK_QTY_CANC > 0" _
                    & " and SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO;" _
                    & " BEGIN FOR R2 IN C2 LOOP" _
                    & " UPDATE ICTSTAT2 SET" _
                    & " WHSE_QTY_PICK = NVL(WHSE_QTY_PICK, 0) + -1 * NVL(R2.PICK_QTY_CANC, 0)" _
                    & " WHERE ITEM_CODE = R2.ITEM_CODE" _
                    & " AND WHSE_CODE = R2.WHSE_CODE;" _
                    & " IF SQL%NOTFOUND THEN" _
                    & " INSERT INTO ICTSTAT2 (WHSE_CODE, ITEM_CODE, WHSE_QTY_PICK)" _
                    & " VALUES (R2.WHSE_CODE, R2.ITEM_CODE,  -1 * NVL(R2.PICK_QTY_CANC,0));" _
                    & " END IF; " _
                    & " END LOOP; END; END;"
                ASCDATA1.ExecuteSQL(sql)
            Next
        Next

        ' IPLB cancels POs in Shipment Confirmation
        If Company_Code = "INT" Then
            ASCMAIN1.Progress("Updating Canceled Pick Tickets", "")
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_STATUS = 'C'")

                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select SOTORDR1.WHSE_CODE" & vbCrLf _
                    & ", SOTORDR2.ITEM_CODE" & vbCrLf _
                    & ", SUM (NVL(SOTPICK2.PICK_QTY,0)) QTY" & vbCrLf _
                    & ", SUM (NVL(SOTPICK2.PICK_QTY_CANC,0)) QTY_CANC" & vbCrLf _
                    & ", SUM (NVL(SOTPICK2.PICK_QTY_BACK,0)) QTY_BACK" & vbCrLf _
                    & " from SOTORDR2,SOTPICK2,SOTPICK1,SOTORDR1 " & vbCrLf _
                    & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                    & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                    & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                    & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
                    & "   and SOTPICK1.PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'" & vbCrLf _
                    & " group by SOTORDR1.WHSE_CODE, SOTORDR2.ITEM_CODE;" & vbCrLf _
                    & " Begin For R1 in C1 Loop" & vbCrLf _
                    & " Update ICTSTAT2 " & vbCrLf _
                    & " Set WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) - R1.QTY" & vbCrLf _
                    & " where ITEM_CODE = R1.ITEM_CODE" & vbCrLf _
                    & "   and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                    & " If SQL%NOTFOUND Then" & vbCrLf _
                    & "   Insert into ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_PICK, WHSE_QTY_OPEN)" & vbCrLf _
                    & "   Values (R1.ITEM_CODE, R1.WHSE_CODE, -1 * R1.QTY, 0);" & vbCrLf _
                    & " End If;" & vbCrLf _
                    & " End Loop; End; End;"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME,TABLE_KEY,INIT_DATE,INIT_OPER,EVENT_TYPE,EVENT_DESC,EVENT_KEY) " & vbCrLf _
                    & " Select 'SOTORDR1',ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'CANC', 'Pick Ticket Canceled', PICK_NO" & vbCrLf _
                    & " from SOTPICK1 where PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update SOTORDR2 Set ORDR_STATUS = 'C'" & vbCrLf _
                    & " where ORDR_NO in" & vbCrLf _
                    & " (Select ORDR_NO from SOTPICK1 where PICK_NO = '" & rowSOTPICK1.Item("PICK_NO") & "')"
                ASCDATA1.ExecuteSQL()

            Next
        End If

        ' Process each BOL, now that it is in Oracle
        ' 3/2/2104 - Need to allow for new changes where the Pick Ticket can be cancelled and still assigned to the SOTSHIP1 record.
        ASCMAIN1.Progress("Updating Sub Tables", "")
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("SELECTED = '1'", "", DataViewRowState.CurrentRows)
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
            ASCMAIN1.Progress("-", SHIP_BOL_NO)

            ASCMAIN1.sql = "Update SOTPICK1" _
                   & " Set PICK_STATUS = 'F'" & vbCrLf _
                   & " where PICK_NO in " & vbCrLf _
                   & " (Select PICK_NO from SOTPICK1 " & vbCrLf _
                   & " where SHIP_BOL_NO = :PARM1 and PICK_STATUS = 'P')" & vbCrLf
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {SHIP_BOL_NO})

            ASCMAIN1.sql = "Update SOTORDR1" _
                    & " Set ORDR_STATUS = 'F'" & vbCrLf _
                    & " where ORDR_NO in " & vbCrLf _
                    & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                    & " where SHIP_BOL_NO = :PARM1 and PICK_STATUS = 'F')" & vbCrLf
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {SHIP_BOL_NO})

            ASCMAIN1.sql = "Update SOTORDR1 " _
                     & "Set TERM_CODE = :PARM1, ORDR_DEPT = :PARM2" & vbCrLf _
                     & " where ORDR_NO in " & vbCrLf _
                     & " (Select ORDR_NO from SOTPICK1 " & vbCrLf _
                     & " where SHIP_BOL_NO = :PARM3 and PICK_STATUS = 'F')" & vbCrLf
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVV",
                                New Object() {rowSOTSHIP1.Item("TERM_CODE"),
                                              rowSOTSHIP1.Item("ORDR_DEPT"),
                                              rowSOTSHIP1.Item("SHIP_BOL_NO")})

            ASCMAIN1.sql = "BEGIN SOPORDR0_G('" & rowSOTSHIP1.Item("ORDR_GROUP_NO") & "'); END;"
            ASCDATA1.ExecuteSQL()
        Next

        ' Process Transfers
        Dim rowSOTORDR1 As DataRow = Nothing
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select()
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
            rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

            If rowSOTORDR1 Is Nothing Then
                rowSOTORDR1 = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_NO = '" & ORDR_NO & "'")
            End If

            ' Process Transfer 
            If rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("ORDR_TYPE_CODE") & String.Empty = "XFR" Then
                Dim XFR_NO As String = ASCDATA1.ExecuteSF _
                       ("SOPSHIP1_XFR", New String() {"INV_NO_IN"}, New Object() {rowSOTINVH1.Item("INV_NO")})
            End If

            ' Process RTV 
            If rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("ORDR_TYPE_CODE") & String.Empty = "RTV" Then
                Dim RTV_NO As String = ASCDATA1.ExecuteSF _
                       ("SOPSHIP1_RTV", New String() {"INV_NO_IN"}, New Object() {rowSOTINVH1.Item("INV_NO")})
            End If

            ASCMAIN1.Progress("-", rowSOTINVH1.Item("INV_NO"))
        Next

        If Company_Code = "AHA" Then
            ' These are like tracking numbers that need the missing characters.
            sql = " BEGIN DECLARE CURSOR C1 IS	"
            sql &= " 	  Select SOTORDR1.ORDR_NO, SOTPICK1.PICK_NO, SOTCART1.CART_NO, SOTCART1.CART_TRACKING_NO, SOTPICK1.INV_NO, SOTPICK1.SHIP_BOL_NO, SOTTRAC2.*"
            sql &= " 	  FROM SOTORDR1, SOTPICK1, SOTCART1, SOTTRAC2"
            sql &= " 	  WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO "
            sql &= " 	  AND SOTCART1.PICK_NO = SOTPICK1.PICK_NO "
            sql &= " 	  AND SOTORDR1.ORDR_STATUS= 'F' "
            sql &= " 	  AND SOTPICK1.PICK_STATUS = 'F' "
            sql &= " 	  AND NVL(SOTTRAC2.PROCESS_IND, '0') = '0'"
            sql &= " 	  AND SOTTRAC2.ORDR_NO_WEB = SOTORDR1.ORDR_CUST_PO"
            sql &= " 	  AND SOTTRAC2.TRACKING_NO LIKE '%' || SOTCART1.CART_TRACKING_NO "
            sql &= " 	  AND SOTTRAC2.TRACKING_NO <> SOTCART1.CART_TRACKING_NO;  "
            sql &= " BEGIN FOR R1 IN C1 LOOP"
            sql &= " 	  UPDATE SOTCART1 SET CART_TRACKING_NO = R1.TRACKING_NO WHERE CART_NO = R1.CART_NO AND PICK_NO = R1.PICK_NO AND CART_TRACKING_NO = R1.CART_TRACKING_NO;"
            sql &= " 	  UPDATE SOTINVH1 SET INV_PRO_NO = R1.TRACKING_NO WHERE INV_TYPE = 'I' AND INV_NO = R1.INV_NO AND INV_PRO_NO = R1.CART_TRACKING_NO;"
            sql &= "   	  UPDATE SOTSHIP1 SET SHIP_REF = R1.TRACKING_NO WHERE SHIP_BOL_NO = R1.SHIP_BOL_NO AND SHIP_REF = R1.CART_TRACKING_NO;"
            sql &= " 	  UPDATE SOTTRAC2 SET PROCESS_IND = '1' WHERE TRACK_NO = R1.TRACK_NO AND TRACK_LNO = R1.TRACK_LNO;"
            sql &= " END LOOP; END; END;"
            ASCDATA1.ExecuteSQL(sql)
        End If

        ASCMAIN1.sql = "BEGIN SAPSSUMX('" & ASCMAIN1.CYP & "'); END;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Private Sub Integrity_Check()

        Dim intrans As Boolean = False
        Try
            ASCMAIN1.sql = "Select ITEM_CODE, WHSE_CODE" & vbCrLf _
                & ", SUM (PO_DTL) PO_DTL, SUM (PO_SUM) PO_SUM" & vbCrLf _
                & ", SUM (PP_DTL) PP_DTL, SUM (PP_SUM) PP_SUM" & vbCrLf _
                & ", SUM (SO_DTL) SO_DTL, SUM (SO_SUM) SO_SUM" & vbCrLf _
                & ", SUM (SP_DTL) SP_DTL, SUM (SP_SUM) SP_SUM" & vbCrLf _
                & ", SUM (PC_DTL) PC_DTL, SUM (PC_SUM) PC_SUM" & vbCrLf _
                & " from (" & vbCrLf _
                & "Select 'IC' TYPE, ITEM_CODE, WHSE_CODE" & vbCrLf _
                & ", 0 PO_DTL, SUM (WHSE_QTY_ONPO) PO_SUM" & vbCrLf _
                & ", 0 PP_DTL, SUM (WHSE_QTY_PLAN) PP_SUM" & vbCrLf _
                & ", 0 SO_DTL, SUM (WHSE_QTY_OPEN) SO_SUM" & vbCrLf _
                & ", 0 SP_DTL, SUM (WHSE_QTY_PICK) SP_SUM" & vbCrLf _
                & ", 0 PC_DTL, SUM (WHSE_QTY_COMM) PC_SUM" & vbCrLf _
                & " from ICTSTAT2" & vbCrLf _
                & " group by ITEM_CODE, WHSE_CODE" & vbCrLf _
                & " union" & vbCrLf _
                & "Select 'PO' TYPE, POTORDR2.ITEM_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
                & ", SUM (POTORDR2.PO_QTY_OPN) PO_DTL, 0 PO_SUM" & vbCrLf _
                & ", 0 PP_DTL, 0 PP_SUM" & vbCrLf _
                & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
                & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
                & ", 0 PC_DTL, 0 PC_SUM" & vbCrLf _
                & " from POTORDR2,POTORDR1 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & " group by POTORDR2.ITEM_CODE, POTORDR1.WHSE_CODE" & vbCrLf _
                & " union" & vbCrLf _
                & "Select 'PP' TYPE, DPTPLAN1.ITEM_CODE, DPTPLAN1.TO_WHSE WHSE_CODE" & vbCrLf _
                & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
                & ", SUM (0 * DPTPLAN1.QTY_PLANNED) PP_DTL, 0 PP_SUM" & vbCrLf _
                & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
                & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
                & ", 0 PC_DTL, 0 PC_SUM" & vbCrLf _
                & " from DPTPLAN1" & vbCrLf _
                & " group by DPTPLAN1.ITEM_CODE, DPTPLAN1.TO_WHSE" & vbCrLf _
                & " union" & vbCrLf _
                & "Select 'SO' TYPE, SOTORDR2.ITEM_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
                & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
                & ", 0 PP_DTL, 0 PP_SUM" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) SO_DTL, 0 SO_SUM" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) SP_DTL, 0 SP_SUM" & vbCrLf _
                & ", 0 PC_DTL, 0 PC_SUM" & vbCrLf _
                & " from SOTORDR2,SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "  and SOTORDR2.ORDR_STATUS <> 'D' and SOTORDR2.ORDR_STATUS <> 'C'" & vbCrLf _
                & " group by SOTORDR2.ITEM_CODE, SOTORDR1.WHSE_CODE" & vbCrLf _
                & " union" & vbCrLf _
                & "Select 'PC' TYPE, POTORDR9.ITEM_CODE, DECODE(POTORDR9.PO_ORDER_LNO,0,DPTPLAN1.AT_WHSE,POTORDR1.VEND_WHSE_CODE) WHSE_CODE" & vbCrLf _
                & ", 0 PO_DTL, 0 PO_SUM" & vbCrLf _
                & ", 0 PP_DTL, 0 PP_SUM" & vbCrLf _
                & ", 0 SO_DTL, 0 SO_SUM" & vbCrLf _
                & ", 0 SP_DTL, 0 SP_SUM" & vbCrLf _
                & ", SUM (POTORDR9.PO_QTY_COM) PP_DTL, 0 PP_SUM" & vbCrLf _
                & " from POTORDR9,POTORDR1,DPTPLAN1" & vbCrLf _
                & " where POTORDR1.PO_ORDER_NO (+) = POTORDR9.PO_ORDER_NO and DPTPLAN1.PLAN_NO (+) = POTORDR9.PO_ORDER_NO" & vbCrLf _
                & " group by POTORDR9.ITEM_CODE, DECODE(POTORDR9.PO_ORDER_LNO,0,DPTPLAN1.AT_WHSE,POTORDR1.VEND_WHSE_CODE)" & vbCrLf _
                & ")" & vbCrLf _
                & " group by ITEM_CODE, WHSE_CODE" & vbCrLf _
                & "having SUM (PO_DTL) <> SUM (PO_SUM) or SUM (PP_DTL) <> SUM (PP_SUM) or SUM (SO_DTL) <> SUM (SO_SUM) or SUM (SP_DTL) <> SUM (SP_SUM) or SUM (PC_DTL) <> SUM (PC_SUM)"

            Dim ICTSTATO As String = ASCMAIN1.Temp_Table
            Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT * FROM " & ICTSTATO)

            If tbl.Rows.Count > 0 Then
                BeginTrans()
                intrans = True
                ASCMAIN1.sql = "" _
                        & "Begin" & vbCrLf _
                        & " Declare" & vbCrLf _
                        & "  Cursor C1 is Select * from " & ICTSTATO & ";" & vbCrLf _
                        & " Begin" & vbCrLf _
                        & "  For R1 in C1 Loop" & vbCrLf _
                        & "   Update ICTSTAT2 Set " & vbCrLf _
                        & "    WHSE_QTY_ONPO = R1.PO_DTL" & vbCrLf _
                        & "  , WHSE_QTY_OPEN = R1.SO_DTL" & vbCrLf _
                        & "  , WHSE_QTY_PICK = R1.SP_DTL" & vbCrLf _
                        & "  , WHSE_QTY_COMM = R1.PC_DTL" & vbCrLf _
                        & "    where WHSE_CODE = R1.WHSE_CODE and ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
                        & "   If SQL%NOTFOUND Then" & vbCrLf _
                        & "    Insert into ICTSTAT2" & vbCrLf _
                        & "     (WHSE_CODE, ITEM_CODE, WHSE_QTY_ONPO, WHSE_QTY_PLAN, WHSE_QTY_OPEN, WHSE_QTY_PICK, WHSE_QTY_COMM)" & vbCrLf _
                        & "    Values (R1.WHSE_CODE, R1.ITEM_CODE, R1.PO_DTL, R1.PP_DTL, R1.SO_DTL, R1.SP_DTL, R1.PC_DTL);" & vbCrLf _
                        & "   End If;" & vbCrLf _
                        & "  End Loop;" & vbCrLf _
                        & " End;" & vbCrLf _
                        & "End;"
                ASCDATA1.ExecuteSQL()
                CommitTrans()
            End If
        Catch ex As Exception
            If intrans Then Rollback()
        End Try

    End Sub


    Public Overrides Sub Update_Record_Post_Commit()
        MyBase.Update_Record_Post_Commit()

        Try
            Integrity_Check()
            ASCMAIN1.Message = String.Empty

        Catch ex As Exception

        End Try

        Dim sql As String = String.Empty

        Dim production As Boolean = False
        If ASCMAIN1.DBS_COMPANY = "AHA" AndAlso ASCMAIN1.DBS_SERVER = "AHA" Then
            production = True
        End If

        If ASCMAIN1.DBS_COMPANY = "INT" AndAlso ASCMAIN1.DBS_SERVER = "INT" Then
            production = True
        End If

        If Not production Then
            Exit Sub
        End If

        ' Email Invoices
        Try
            Dim tblSOTINVH1 As DataTable = dst.Tables("SOTINVH1").Clone
            Dim tblSOTINVH1_GRATIS As DataTable = dst.Tables("SOTINVH1").Clone
            Dim rowSOTORDR1 As DataRow = Nothing

            For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("SHIP_STATUS = 'F'", "SHIP_BOL_NO")
                tblSOTINVH1.Rows.Clear()
                tblSOTINVH1_GRATIS.Rows.Clear()

                For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & rowSOTSHIP1.Item("SHIP_BOL_NO") & "'")
                    rowSOTORDR1 = dst.Tables("SOTORDR1").Rows.Find(rowSOTPICK1.Item("ORDR_NO") & String.Empty)

                    ' Do not Try To Print Transfer Invoices
                    If rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("ORDR_TYPE_CODE") & String.Empty = "XFR" Then
                        Continue For
                    End If

                    If rowSOTPICK1.Item("INV_NO") & String.Empty <> String.Empty Then
                        tblSOTINVH1.Rows.Add(dst.Tables("SOTINVH1").Select("INV_NO = '" & rowSOTPICK1.Item("INV_NO") & "'")(0).ItemArray)

                        If rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty = "G" Then
                            tblSOTINVH1_GRATIS.Rows.Add(dst.Tables("SOTINVH1").Select("INV_NO = '" & rowSOTPICK1.Item("INV_NO") & "'")(0).ItemArray)
                        End If
                    End If
                Next

                If tblSOTINVH1.Rows.Count > 0 Then
                    If Not ASCMAIN1.Running_in_VS Then
                        EmailInvoice(tblSOTINVH1)
                    End If
                End If

                If tblSOTINVH1_GRATIS.Rows.Count > 0 Then
                    If Not ASCMAIN1.Running_in_VS Then
                        EmailGratisInvoice(tblSOTINVH1_GRATIS)
                    End If
                End If
            Next

            ASCMAIN1.Message = String.Empty
        Catch ex As Exception
            MessageBox.Show("Error Emailing Invoices: " & ex.Message, "Email Invoices", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ASCMAIN1.Message = String.Empty
        End Try

        If Company_Code = "INT" Then
            For Each row As DataRow In tblSOTSHIP_CART.Select()
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & String.Empty
                Dim CUST_SHIP_EMAIL As String = row.Item("CUST_SHIP_EMAIL") & String.Empty
                Dim CUST_SHIP_EMAIL_EXCEL As String = row.Item("CUST_SHIP_EMAIL_EXCEL") & String.Empty

                ASCMAIN1.Progress("Generating Carton Report for shipment " & SHIP_BOL_NO)
                CUST_SHIP_EMAIL = CUST_SHIP_EMAIL.Trim

                Select Case CUST_SHIP_EMAIL_EXCEL
                    Case "1"
                        Email_Shipment_Excel(SHIP_BOL_NO, CUST_SHIP_EMAIL)
                    Case Else
                        Email_Shipment(SHIP_BOL_NO, CUST_SHIP_EMAIL)
                End Select
            Next
        End If

    End Sub

    ''' <summary>
    ''' Process Credirt Cards for Sales Orders with a Credit Card Authorization
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ChargeCreditCards(ByRef CCPA_NO_ORDR As String, ByVal INV_NO As String, ByVal Charge_Amount As Decimal) As Boolean

        Dim CreditCardProcessed As Boolean = True
        Dim ORDR_NO As String = String.Empty
        Dim processedCreditCard As Boolean = False

        Try
            dst.Tables("TATEVNT1").Rows.Clear()

            Dim rowSOTINVH1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTINVH1 where inv_type = 'I' and inv_no = '" & INV_NO & "'")
            ORDR_NO = rowSOTINVH1.Item("ORDR_NO") & String.Empty

            ASCMAIN1.Progress("Processing Credit Card", INV_NO)

            CCPA_NO_ORDR = CCPA_NO_ORDR.Trim
            If CCPA_NO_ORDR.Length = 0 Then Return True

            Dim INV_SALES As Decimal = Val(rowSOTINVH1.Item("INV_SALES") & String.Empty)
            Dim INV_STAX As Decimal = Val(rowSOTINVH1.Item("INV_STAX") & String.Empty)
            Dim INV_FREIGHT As Decimal = Val(rowSOTINVH1.Item("INV_FREIGHT") & String.Empty)
            INV_SALES += INV_FREIGHT

            If INV_SALES > 0 Then
                Try
                    Dim ResponseText As String = String.Empty
                    Dim CCPA_NO As String = ProcessCreditCardAuthorization(CCPA_NO_ORDR, INV_SALES, INV_FREIGHT, INV_STAX, ResponseText)
                    CreditCardProcessed = CCPA_NO.Length > 0 AndAlso CreditCardProcessed

                    If CCPA_NO.Length > 0 Then
                        processedCreditCard = True
                        ' This is done to preserve credit card transactions if the code causes an error after this point
                        MyBase.BeginTrans()
                        ' Record Transaction Number in Order Header. Will be placed in Invoice Header
                        Dim rowARTCCPA1 As DataRow = LookUp("ARTCCPA1", CCPA_NO)
                        If rowARTCCPA1 IsNot Nothing Then
                            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                            ASCMAIN1.sql = "Update SOTORDR1 SET CC_TRANS_ID = '" & rowARTCCPA1.Item("TRANS_ID") & "' WHERE ORDR_NO = '" & ORDR_NO & "'"
                            ASCDATA1.ExecuteSQL()
                        End If

                        ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " _
                             & " values " _
                             & "('SOTORDR1', '" & ORDR_NO & "', SYSDATE, '" & ASCMAIN1.USER_ID & "', 'CCCHG','Credit card charged: " & Format(INV_SALES, "#,##0.00") & "', NULL)"
                        ASCDATA1.ExecuteSQL()
                        MyBase.CommitTrans()

                    Else
                        Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
                        row.Item("TABLE_NAME") = "SOTORDR1"
                        row.Item("TABLE_KEY") = ORDR_NO
                        row.Item("INIT_DATE") = DATETIME_STAMP
                        row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        row.Item("EVENT_TYPE") = "CCP"
                        row.Item("EVENT_DESC") = "Credit Card Error: " & ResponseText
                        row.Item("EVENT_KEY") = ""
                        dst.Tables("TATEVNT1").Rows.Add(row)
                    End If

                Catch ex As Exception
                    MyBase.Rollback(ex.Message)
                    Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
                    row.Item("TABLE_NAME") = "SOTORDR1"
                    row.Item("TABLE_KEY") = ORDR_NO
                    row.Item("INIT_DATE") = DATETIME_STAMP
                    row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    row.Item("EVENT_TYPE") = "CCP"
                    row.Item("EVENT_DESC") = "Credit Card Error: " & ex.Message
                    row.Item("EVENT_KEY") = ""
                    dst.Tables("TATEVNT1").Rows.Add(row)
                End Try
            End If

        Catch ex As Exception
            Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
            row.Item("TABLE_NAME") = "SOTORDR1"
            row.Item("TABLE_KEY") = ORDR_NO
            row.Item("INIT_DATE") = DATETIME_STAMP
            row.Item("INIT_OPER") = ASCMAIN1.USER_ID
            row.Item("EVENT_TYPE") = "CCP"
            row.Item("EVENT_DESC") = "Credit Card Error: " & ex.Message
            row.Item("EVENT_KEY") = ""
            dst.Tables("TATEVNT1").Rows.Add(row)
        End Try

        ' Update sales order events with Credit card errors
        Dim EVENT_DESC_LEN As Int16 = dst.Tables("TATEVNT1").Columns("EVENT_DESC").MaxLength
        For Each row As DataRow In dst.Tables("TATEVNT1").Select()
            If (row.Item("EVENT_DESC") & String.Empty).ToString.Length > EVENT_DESC_LEN Then
                row.Item("EVENT_DESC") = (row.Item("EVENT_DESC") & String.Empty).ToString.Substring(0, EVENT_DESC_LEN)
            End If
        Next

        Try
            BeginTrans()
            Update_Record_TDA("TATEVNT1")
            CommitTrans()
        Catch ex As Exception
            Rollback()
        End Try

        Return processedCreditCard
    End Function

    ''' <summary>
    ''' Convert Credit Card Authorization to a Captured Sale - Communicate with Clearing house
    ''' </summary>
    ''' <param name="AUTH_CCPA_NO"></param>
    ''' <param name="ChargeAmount"></param>
    ''' <param name="freightAmount"></param>
    ''' <param name="salesTax"></param>
    ''' <param name="ResponseText"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ProcessCreditCardAuthorization(ByVal AUTH_CCPA_NO As String,
                                                ByVal ChargeAmount As Double,
                                                ByVal freightAmount As Decimal,
                                                ByVal salesTax As Decimal,
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
                    'ResponseText = "Error trying to void initial CC Authorization: " & ex.Message
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

                ' Needed for only FDMS
                If CreditCardProcessor.objCCProcessor.ProcessingType = TAC.ARCCCARD.ProcessingTypes.FDMS Then
                    With .objCCProcessor.Level2Data
                        .Clear()

                        .CardType = CreditCardProcessor.objCCProcessor.CreditCardType

                        Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", rowARTCCPA1_AUTH.Item("ORDR_NO")) ' dst.Tables("SOTORDR1").Rows(0)
                        Dim rowSHIPTO As DataRow = ASCDATA1.GetDataRow("select * from sotordr5 where CUST_ADDR_TYPE = 'ST' and ordr_no = '" & rowARTCCPA1_AUTH.Item("ORDR_NO") & "'")  'dst.Tables("SOTORDR5").Select("CUST_ADDR_TYPE = 'ST'")(0)

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

                        Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE") & String.Empty
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
                            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)

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
                End If

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
            ResponseText = "Error trying to process Credit Card: " & ex.Message
        Finally
            CreditCardProcessor.Dispose()
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

        Return CCPA_NO

    End Function

    ''' <summary>
    ''' Log errors in the data to prevent the report to offer the update
    ''' </summary>
    ''' <param name="ErrorMessage"></param>
    ''' <remarks></remarks>
    Private Sub AddError(ByVal ErrorMessage As String, Optional FYI As String = "0")
        Try
            If dst.Tables("ERRORS").Select("MESSAGE = '" & ErrorMessage.Replace("'", "") & "'").Length = 0 Then
                SEQ_NO += 1
                dst.Tables("ERRORS").Rows.Add(New Object() {SEQ_NO, ErrorMessage, FYI})
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' Update Sales Orders, Pick Tickets and Create Invoices for the Shipments
    ''' </summary>
    ''' <param name="BILL_OF_LADING_NO"></param>
    ''' <remarks></remarks>
    Private Sub UpdateOrderPicksAndCreateInvoices(ByVal BILL_OF_LADING_NO As String)

        ' Update Sotordr1 and Sortordr2 and Possibly SOTPICK1,2 
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "' and SELECTED = '1'", "", DataViewRowState.CurrentRows)
            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SHIP_BOL_NO = '" & SHIP_BOL_NO & "' AND SELECTED = '1'", "", DataViewRowState.CurrentRows)
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")
                Dim ORDR_NO As String = rowSOTPICK1.Item("ORDR_NO")
                Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)

                ' Finalize Sales Order and Pick Ticket
                If rowSOTPICK1.Item("PICK_STATUS") = "C" Then
                    rowSOTORDR1.Item("ORDR_STATUS") = "C"

                    Dim row As DataRow = dst.Tables("TATEVNT1").Rows.Add
                    row.Item("TABLE_NAME") = "SOTORDR1"
                    row.Item("TABLE_KEY") = ORDR_NO
                    row.Item("INIT_DATE") = DATETIME_STAMP
                    row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    row.Item("EVENT_TYPE") = "CANC"
                    row.Item("EVENT_DESC") = "Order Canceled by Warehouse"
                    row.Item("EVENT_KEY") = ""
                Else
                    rowSOTORDR1.Item("ORDR_STATUS") = "F"
                End If

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
                    rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & String.Empty) ' + (Val(rowSOTPICK2.Item("PICK_QTY_CANC") & String.Empty) + Val(rowSOTPICK2.Item("PICK_QTY_BACK") & String.Empty))

                    If rowSOTPICK1.Item("PICK_STATUS") = "C" Then
                        rowSOTORDR2.Item("ORDR_STATUS") = "C"
                    Else
                        rowSOTORDR2.Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                        rowSOTORDR2.Item("ORDR_STATUS") = "F"
                    End If
                Next

                If rowSOTPICK1.Item("PICK_STATUS") <> "C" Then
                    rowSOTORDR1.Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                End If
                rowSOTORDR1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                rowSOTORDR1.Item("LAST_DATE") = DATETIME_STAMP

                If (rowSOTORDR1.Item("SHIP_VIA_CODE") & String.Empty <> rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty) _
                    AndAlso rowSOTORDR1.Item("ORDR_STATUS") <> "C" Then
                    Dim row As DataRow = dst.Tables("TATEVNT1").Rows.Add
                    row.Item("TABLE_NAME") = "SOTORDR1"
                    row.Item("TABLE_KEY") = ORDR_NO
                    row.Item("INIT_DATE") = DATETIME_STAMP
                    row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    row.Item("EVENT_TYPE") = "SHPMTC"
                    row.Item("EVENT_DESC") = "Ship Via was changed from " _
                        & rowSOTORDR1.Item("SHIP_VIA_CODE") & " to " & rowSOTSHIP1.Item("SHIP_VIA_CODE") & String.Empty
                    row.Item("EVENT_KEY") = ""
                End If
            Next
        Next

        ' Create Invoice Records
        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("BILL_OF_LADING_NO = '" & BILL_OF_LADING_NO & "' AND SELECTED = '1'", "", DataViewRowState.CurrentRows)

            Dim SHIP_BOL_NO As String = rowSOTSHIP1.Item("SHIP_BOL_NO")
            If rowSOTSHIP1.Item("SHIP_FOB") & String.Empty = String.Empty Then
                rowSOTSHIP1.Item("SHIP_FOB") = ABSolution.ASCDATA1.GetDataValue("SELECT MAX(ORDR_FOB) FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTPICK1 WHERE SHIP_BOL_NO = '" & SHIP_BOL_NO & "')")
            End If

            Dim SOCINVH1 As New TAC.SOCINVH1(dst.Tables("SOTINVH1"), dst.Tables("SOTINVH2"),
                                              dst.Tables("SOTPICK1"), dst.Tables("SOTPICK2"),
                                              dst.Tables("ARTOPEN1"), dst.Tables("SOTSHIP1"),
                                              dst.Tables("SOTORDR5"))
            SOCINVH1.CreateInvoices(SHIP_BOL_NO)

            ' Log Credit Card Sales
            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("SELECTED = '1' AND SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "SHIP_BOL_NO")
                If rowSOTPICK1.Item("CCPA_NO_ORDR") & String.Empty = String.Empty Then
                    Continue For
                End If
                Dim INV_NO As String = rowSOTPICK1.Item("INV_NO")
                Dim INV_TOTAL_AMOUNT As Decimal = Val(dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'")(0).Item("INV_TOTAL_AMOUNT") & String.Empty)
                dst.Tables("ARTCCPAC").Rows.Add(New Object() {rowSOTPICK1.Item("PICK_NO"), rowSOTPICK1.Item("INV_NO"), DateTime.Now, rowSOTPICK1.Item("CCPA_NO_ORDR"), INV_TOTAL_AMOUNT})
                If dst.Tables("ARTCCPA1").Rows.Find(rowSOTPICK1.Item("CCPA_NO_ORDR")) Is Nothing Then
                    Fill_Records("ARTCCPA1", rowSOTPICK1.Item("CCPA_NO_ORDR"), False)
                End If
            Next
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
            ' New field 9/21/2016
            rowSOTORDR1.Item("ORDR_DATE_SHIPPED") = rowSOTINVH1.Item("INV_DATE_SHIPPED")
        Next

        If CURR_CODE = "" Or Val(CURR_EXCH_RATE) = 0 Then
            MessageBox.Show("*****************************************" & Environment.NewLine &
                            "Please contact ABS about this shipment." & Environment.NewLine &
                            "Let them know the Currency Code is blank!" & Environment.NewLine &
                            "*****************************************", "Shipment", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub

    ''' <summary>
    ''' Set Column Expressions
    ''' </summary>
    ''' <param name="tf"></param>
    ''' <remarks></remarks>
    Private Sub ToggleDataTableExpressions(ByVal tf As Boolean)

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

    ''' <summary>
    ''' Downloads the ADS Shipment Files, Inventory receipts/adjustments Files and the returns files.
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetShipmentData() As Boolean

        Try
            Dim clsSOCASDO1 As New TAC.SOCADSO1
            Dim sql As String = String.Empty

            ASCMAIN1.Progress("Process Ship Confirmation Data", "")
            dst.Tables("EDT945T1").Rows.Clear()
            dst.Tables("EDT945T2").Rows.Clear()

            Dim success As Boolean = False
            Select Case Company_Code

                Case "INT"
                    ' Clarins places the same Master BOL No on many different shipments (SHIP_BOL_NO). Modify them for IPLBAE
                    ' Set the BOL to the SHIP_BOL_NO to group these together for processing; otherwise, they get held up until all pick tickets are here from other shipments.
                    sql = "Update EDT945T1 SET EDI_BOL_NO = EDI_SHIPMENT_ID, EDI_MASTER_BOL_NO = EDI_SHIPMENT_ID WHERE CUST_CODE IN ('IPLBAE', 'IPLBEDU') AND EDI_PROCESS_IND = '0' AND EDI_SHIPMENT_ID IS NOT NULL"
                    ASCDATA1.ExecuteSQL(sql)
                    success = clsSOCASDO1.ProcessClarinsShipConfirmationData(dst.Tables("EDT945T1"), dst.Tables("EDT945T2"))
                Case Else
                    success = False
                    AddError("Error Processing Ship Confirmation Data - Invalid Company Code")
            End Select

            If Not success Then
                AddError("Error Processing Ship Confirmation Data: " & clsSOCASDO1.LastError)
            Else
                Try

                    If clsSOCASDO1.LastError.Length > 0 Then
                        If clsSOCASDO1.LastError.Contains(Environment.NewLine) Then
                            For Each eString As String In clsSOCASDO1.LastError.Split(Environment.NewLine)
                                eString = eString.Trim
                                If eString.Length = 0 Then Continue For
                                AddError(eString, "1")
                            Next
                        Else
                            AddError(clsSOCASDO1.LastError, "1")
                        End If
                    End If

                    'Need to make sure we do not have duplicate Pick Tickets already in EDT945T1
                    BeginTrans()
                    Update_Record_TDA("EDT945T1")
                    Update_Record_TDA("EDT945T2")

                    Select Case Company_Code
                        Case "AHA"
                            ' 12/20/2017 - Auto create 945 details (zero quantity shipped) for Headers without details.
                            '   This means the pick ticket was cancelled by ADS.
                            sql = "BEGIN DECLARE CURSOR C1 IS"
                            sql &= " SELECT EDI_DOC_SEQ_NO FROM EDT945T1 WHERE NVL(EDI_PROCESS_IND, '0') = '0'"
                            sql &= " MINUS "
                            sql &= " SELECT EDI_DOC_SEQ_NO FROM EDT945T2;"
                            sql &= " BEGIN FOR R1 IN C1 LOOP"
                            sql &= " INSERT INTO EDT945T2"
                            sql &= " (EDI_DOC_SEQ_NO, EDI_DTL_SEQ, EDI_CART_NO, EDI_SHIPMENT_STATUS_CODE,"
                            sql &= " PICK_LNO, PICK_QTY, EDI_SHIP_QTY, STYLE_CODE)"
                            sql &= " SELECT EDT945T1.EDI_DOC_SEQ_NO, ROWNUM EDI_DTL_SEQ, EDT945T1.EDI_DOC_SEQ_NO EDI_CART_NO,"
                            sql &= " 'SH' EDI_SHIPMENT_STATUS_CODE, SOTPICK2.PICK_LNO, SOTPICK2.PICK_QTY, 0 EDI_SHIP_QTY,"
                            sql &= " SOTORDR2.ITEM_CODE STYLE_CODE"
                            sql &= " FROM EDT945T1, SOTPICK2, SOTORDR2"
                            sql &= " WHERE EDT945T1.EDI_PICK_NO = SOTPICK2.PICK_NO"
                            sql &= " AND SOTPICK2.ORDR_NO = SOTORDR2.ORDR_NO"
                            sql &= " AND SOTPICK2.ORDR_LNO = SOTORDR2.ORDR_LNO"
                            sql &= " AND EDT945T1.EDI_DOC_SEQ_NO = R1.EDI_DOC_SEQ_NO"
                            sql &= " AND NVL(EDT945T1.EDI_PROCESS_IND, '0') = '0';"
                            sql &= " DELETE FROM SOTCART2 WHERE CART_NO IN (SELECT CART_NO FROM SOTCART1 WHERE PICK_NO = (SELECT EDI_PICK_NO FROM EDT945T1 WHERE EDT945T1.EDI_DOC_SEQ_NO = R1.EDI_DOC_SEQ_NO));"
                            sql &= " DELETE FROM SOTCART1 WHERE CART_NO IN (SELECT CART_NO FROM SOTCART1 WHERE PICK_NO = (SELECT EDI_PICK_NO FROM EDT945T1 WHERE EDT945T1.EDI_DOC_SEQ_NO = R1.EDI_DOC_SEQ_NO));"
                            sql &= " END LOOP; END; END;"
                            ASCDATA1.ExecuteSQL(sql)
                    End Select

                    'Need to make sure we do not have duplicate Pick Tickets already in EDT945T1
                    sql = "SELECT EDI_PICK_NO, COUNT(*) FROM EDT945T1"
                    sql &= " WHERE NVL(EDI_PROCESS_IND, '0') IN ('0', '1')"
                    sql &= " AND COMPANY_CODE = '" & Company_Code & "'"
                    sql &= " GROUP BY EDI_PICK_NO "
                    sql &= " HAVING COUNT(*) > 1"

                    For Each rowEDT945T1 As DataRow In ASCDATA1.GetDataTable(sql).Select("", "EDI_PICK_NO")
                        AddError("Pick Ticket " & rowEDT945T1.Item("EDI_PICK_NO") & " appears more than once in the EDI table and will be skipped.", "1")
                        sql = "Update EDT945T1 set EDI_PROCESS_IND = 'X' where NVL(EDI_PROCESS_IND, '0') = '0' AND EDI_PICK_NO = '" & rowEDT945T1.Item("EDI_PICK_NO") & "'"
                        ASCDATA1.ExecuteSQL(sql)
                    Next

                    ' Clean up canceled orders sent to us from Warehouse
                    sql = "  Select * from SOTPICK1 where PICK_STATUS = 'D'"
                    sql &= "  and PICK_NO in"
                    sql &= "  ("
                    sql &= "  Select EDI_PICK_NO from EDT945T1 where EDI_PROCESS_IND = '0'"
                    sql &= "  )"

                    For Each row As DataRow In ASCDATA1.GetDataTable(sql).Select("", "ORDR_NO")
                        AddError("Pick Ticket " & row.Item("PICK_NO") & " for Sales Order " & row.Item("ORDR_NO") & " was de-release; however, the warehouse sent a Shipment Confirmation. Shipment data was ignored.", "1")
                    Next

                    sql = "UPDATE EDT945T1 SET EDI_PROCESS_IND = 'D'"
                    sql &= " WHERE EDI_PICK_NO IN "
                    sql &= " ( "
                    sql &= " select PICK_NO from sotpick1 where pick_no in"
                    sql &= " ("
                    sql &= " select edi_pick_no from edt945t1 where edi_process_ind = '0'"
                    sql &= " )"
                    sql &= " AND PICK_STATUS = 'D'"
                    sql &= " ) "
                    ASCDATA1.ExecuteSQL(sql)

                    CommitTrans()
                Catch ex As Exception
                    Rollback("Error Updating Ship Confirmation Data: " & ex.Message)
                    dst.Tables("EDT945T1").Clear()
                    dst.Tables("EDT945T2").Clear()
                End Try
            End If

            ASCMAIN1.Progress("Process Inventory Transactions", "")
            dst.Tables("EDTTRXN1").Rows.Clear()
            Select Case Company_Code

                Case "INT"
                    If Not clsSOCASDO1.ProcessClarinsInventoryTransactions(dst.Tables("EDTTRXN1")) Then

                    End If
            End Select

            ASCMAIN1.Progress("Process Returns Transactions", "")
            dst.Tables("EDTRTRN1").Rows.Clear()
            dst.Tables("EDTRTRN2").Rows.Clear()

            success = False
            Select Case Company_Code
                Case "INT"
                    success = clsSOCASDO1.ProcessClarinsReturnsTransactions(dst.Tables("EDTRTRN1"), dst.Tables("EDTRTRN2"))
                Case Else
                    success = False
                    AddError("Error Processing Ship Confirmation Data - Invalid Company Code")
            End Select

            If Not success Then
                AddError("Error Processing Returns Transactions: " & clsSOCASDO1.LastError, "1")
                dst.Tables("EDTRTRN1").Rows.Clear()
                dst.Tables("EDTRTRN2").Rows.Clear()
            Else
                Try
                    BeginTrans()
                    Update_Record_TDA("EDTRTRN1")
                    Update_Record_TDA("EDTRTRN2")
                    CommitTrans()
                Catch ex As Exception
                    Rollback("Error Updating Returns Transactions: " & ex.Message)
                    dst.Tables("EDTRTRN1").Clear()
                    dst.Tables("EDTRTRN2").Clear()
                End Try
            End If

            GetShipmentData = True

        Catch ex As Exception
            AddError("DownLoad Files Error: " & ex.Message)
            GetShipmentData = False
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Function

    Private Function EmailGratisInvoice(ByRef tblSOTINVH1 As DataTable) As Boolean

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        Dim ErrorMessage As String = String.Empty

        Try
            If ASCMAIN1.Running_in_VS Then
                Stop
            End If

            If tblSOTINVH1 Is Nothing OrElse tblSOTINVH1.Rows.Count = 0 Then
                Return False
            End If

            For Each rowSOTINVH1 As DataRow In tblSOTINVH1.Select("")
                Dim INV_NO As String = rowSOTINVH1.Item("INV_NO") & String.Empty
                Dim INV_PRO_NO As String = rowSOTINVH1.Item("INV_PRO_NO") & String.Empty
                Dim SHIP_VIA_DESC As String = rowSOTINVH1.Item("SHIP_VIA_DESC") & String.Empty
                Dim INV_DATE_SHIPPED As String = rowSOTINVH1.Item("INV_DATE_SHIPPED") & String.Empty

                Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty
                If ORDR_NO.Length = 0 Then Continue For

                Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE") & String.Empty
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 Is Nothing Then
                    Continue For
                End If

                Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                If rowSOTORDR1 Is Nothing Then
                    Continue For
                End If

                If rowSOTORDR1.Item("ORDR_SOURCE") & String.Empty <> "G" Then
                    Continue For
                End If

                Dim rowSOTORDR5 As DataRow = LookUp("SOTORDR5", {ORDR_NO, "ST"})
                If rowSOTORDR5 Is Nothing Then
                    Continue For
                End If

                EMAIL_ADDRESSs.Clear()
                Dim CUST_EMAIL As String = rowSOTORDR5.Item("CUST_EMAIL") & String.Empty
                CUST_EMAIL = CUST_EMAIL.Trim

                If CUST_EMAIL.Length = 0 Then
                    Continue For
                End If

                If CUST_EMAIL.Length > 5 Then
                    EMAIL_ADDRESSs.Add(CUST_EMAIL, CUST_EMAIL)
                End If

                If EMAIL_ADDRESSs.Count = 0 Then
                    Continue For
                End If

                Dim RPT As String = "SORINVP1"

                If Not REPORTS.ContainsKey(RPT) Then
                    REPORTS.Add(RPT, Load_rptClass(RPT))
                    REPORTS(RPT).Prepare_dst(False, "")
                End If

                REPORTS(RPT).Fill_Records_RPT(New String() {$" and SOTINVH1.INV_NO = ('{INV_NO}')"})
                Dim REPORT_NO As String = ""

                Dim CustomerReport As String = RPT
                Select Case Company_Code
                    Case "INT"
                        CustomerReport = "SORINVPI"
                End Select

                Dim CONS_INV As String = "0"
                With REPORTS(RPT).clsASCBASE1
                    .Print_Report_Begin()
                    .CR_params.Add("SUBT", "")
                    .CR_params.Add("CONS_INV", CONS_INV)
                    REPORT_NO = .Generate_Report(CustomerReport, "Invoice", , True, , , "PDF", INV_NO, False)
                    .Print_Report_End(True, True)
                End With

                Dim ATTACHMENTs As New Dictionary(Of String, String)
                ATTACHMENTs.Add(INV_NO & ".pdf", ASCMAIN1.Folders("Temp") & INV_NO & ".pdf")

                Dim SUBJECT As String = "Gratis Invoice " & INV_NO

                Select Case Company_Code
                    Case "INT"
                        SUBJECT = "Interparfums Luxury Brands, Inc. Gratis Invoice: " & INV_NO
                End Select

                Dim EMAIL_BODY As String = $"We've shipped your order! Your world is about to smell a whole lot better." & vbCrLf & vbCrLf
                EMAIL_BODY &= $"You order shipped Via: {SHIP_VIA_DESC}"
                If IsDate(INV_DATE_SHIPPED) Then
                    EMAIL_BODY &= " on " & CDate(INV_DATE_SHIPPED).ToString("MM/dd/yyyy")
                End If

                If INV_PRO_NO.Length > 0 Then
                    EMAIL_BODY &= $", Tracking No: {INV_PRO_NO}"
                End If

                EMAIL_BODY &= "."

                EMAIL_BODY &= vbCrLf & vbCrLf

                ASCMAIN1.Message = String.Empty
                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    SUBJECT, "GRATIS_INV", True, False, CUST_CODE, rowARTCUST1.Item("CUST_NAME"), "Customer", EMAIL_BODY)

                If SEND_NO = String.Empty Then
                    EmailGratisInvoice = False
                    ErrorMessage = ASCMAIN1.Message
                    If ErrorMessage.Length = 0 Then
                        ErrorMessage = "Not able to Send Gratis Email."
                    End If

                    Select Case Company_Code
                        Case "INT"
                            MessageBox.Show("The following Error occurred Emailing Gratis Invoice to customer " & rowARTCUST1.Item("CUST_CODE") & " - " & rowARTCUST1.Item("CUST_NAME") & ": " & ErrorMessage,
                            "Email Invoices",
                             MessageBoxButtons.OK)

                    End Select
                Else
                    EmailGratisInvoice = True
                End If
            Next

        Catch ex As Exception
            EmailGratisInvoice = False
            ErrorMessage = ex.Message
        End Try

    End Function

    ''' <summary>
    ''' Emails Invoices to Customers and sales Reps.
    ''' </summary>
    ''' <param name="tblSOTINVH1"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function EmailInvoice(ByRef tblSOTINVH1 As DataTable) As Boolean

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        Dim ErrorMessage As String = String.Empty

        Try
            If ASCMAIN1.Running_in_VS Then
                Stop
            End If

            If tblSOTINVH1 Is Nothing OrElse tblSOTINVH1.Rows.Count = 0 Then
                Return False
            End If

            Dim rowSOTINVH1 As DataRow = tblSOTINVH1.Rows(0)
            Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE")
            Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO")
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim CONS_INV As String = "0"

            ASCMAIN1.Progress("Emailing Invoice", CUST_CODE)

            If rowSOTINVH1.Item("INV_NO_CONS") & String.Empty <> String.Empty Then
                CONS_INV = "1"
            End If

            ' See if the customer receives an acknowledgment
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
            Dim rowSOTSREP1 As DataRow = Nothing
            Dim salesRepEmail As String = String.Empty
            Dim custEmailShipAck As String = String.Empty
            Dim custCCEmailShipAck As String = String.Empty

            If rowARTCUST1 IsNot Nothing Then
                If rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty = "E" OrElse rowARTCUST1.Item("CUST_XMIT_INV_VIA") & String.Empty = "B" Then
                    custEmailShipAck = (rowARTCUST1.Item("CUST_INV_EMAIL") & "").ToString.Trim.ToUpper
                    custCCEmailShipAck = (rowARTCUST1.Item("CUST_INV_CC") & "").ToString.Trim.ToUpper
                ElseIf Company_Code = "INT" Then
                    ' ILPB does not email to Sales Reps
                    Return True
                End If

                If custEmailShipAck.Length = 0 AndAlso custCCEmailShipAck.Length > 0 Then
                    custEmailShipAck = custCCEmailShipAck
                    custCCEmailShipAck = String.Empty
                End If
            End If

            ' ILPB does not email to Sales Reps
            If rowARTCUST2 IsNot Nothing AndAlso Company_Code <> "INT" Then
                salesRepEmail = (rowARTCUST2.Item("CUST_STORE_EMAIL") & String.Empty).ToString.Trim.ToUpper
                rowSOTSREP1 = LookUp("SOTSREP1", rowARTCUST2.Item("SREP_CODE") & String.Empty)
                If rowSOTSREP1 IsNot Nothing Then
                    salesRepEmail = (rowSOTSREP1.Item("SREP_EMAIL") & String.Empty).ToString.Trim.ToUpper
                End If
            End If

            If custEmailShipAck.Length = 0 AndAlso salesRepEmail.Length = 0 Then
                Return True
            End If

            ' 09/12/2023
            ' Ed – please make the change the Nathan is asking for.
            ' He wants To avoid emailing invoices To a Single customer, MONTBLANC, If(And only If) the invoice total Is $0.
            Dim invNos As String = String.Empty
            For Each row As DataRow In tblSOTINVH1.Select("")
                Select Case Company_Code
                    Case "INT"
                        If row.Item("CUST_CODE") & String.Empty = "MONTBLANC" Then
                            If Val(row.Item("INV_TOTAL_AMOUNT") & String.Empty) = 0 Then
                                Continue For
                            End If
                        End If
                End Select
                invNos &= ", '" & row.Item("INV_NO") & "'"
            Next

            If invNos.Length = 0 Then
                Return True
            End If

            invNos = invNos.Substring(1).Trim

            ' Default Setting
            Dim RPT As String = "SORINVP1"

            If Not REPORTS.ContainsKey(RPT) Then
                REPORTS.Add(RPT, Load_rptClass(RPT))
                REPORTS(RPT).Prepare_dst(False, "")
            End If

            REPORTS(RPT).Fill_Records_RPT(New String() {" and SOTINVH1.INV_NO IN (" & invNos & ")"})

            Dim REPORT_NO As String = ""

            Dim CustomerReport As String = RPT
            Select Case Company_Code
                Case "INT"
                    CustomerReport = "SORINVPI"
            End Select

            With REPORTS(RPT).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", CONS_INV)
                REPORT_NO = .Generate_Report(CustomerReport, "Invoice", , True, , , "PDF", INV_NO, False)
                .Print_Report_End(True, True)
            End With

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(INV_NO & ".pdf", ASCMAIN1.Folders("Temp") & INV_NO & ".pdf")

            Dim SUBJECT As String = "Invoice " & INV_NO

            Select Case Company_Code
                Case "AHA"
                    SUBJECT = $"Ahava Invoice {INV_NO}"

                Case "INT"
                    ' INC-6121 different Invoice email subject for Bealls 09/25/2024
                    Select Case CUST_CODE
                        Case "BEALLSOUT", "BEALLSOUT4"
                            SUBJECT = $"CURRENT, Interparfums Luxury Brands, Inc., Duns (962481698), Invoice {INV_NO}"
                        Case Else
                            SUBJECT = $"Interparfums Luxury Brands, Inc. Invoice {INV_NO}"
                    End Select
            End Select

            ' Convert all Commas to Semi-Colons - Clean up the data provided by the user.
            custEmailShipAck = custEmailShipAck.Replace(",", ";")
            salesRepEmail = salesRepEmail.Replace(",", ";")
            custCCEmailShipAck = custCCEmailShipAck.Replace(",", ";")

            ' Concatentate and process all email addresses
            For Each emailAddress As String In (custEmailShipAck & ";" & salesRepEmail & ";" & custCCEmailShipAck).ToString.Split(";")
                emailAddress = emailAddress.Trim
                If emailAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(emailAddress) Then
                    EMAIL_ADDRESSs.Add(emailAddress, emailAddress)
                End If
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Return True
            End If

            ASCMAIN1.Message = String.Empty
            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    SUBJECT, "INV", True, False, CUST_CODE, rowARTCUST1.Item("CUST_NAME"), "Customer")

            If SEND_NO = String.Empty Then
                EmailInvoice = False
                ErrorMessage = ASCMAIN1.Message
                If ErrorMessage.Length = 0 Then
                    ErrorMessage = "Not able to Send Email."
                End If

                Select Case Company_Code
                    Case "INT"
                        MessageBox.Show("The following Error occurred emailing invoices to customer " & rowARTCUST1.Item("CUST_CODE") & " - " & rowARTCUST1.Item("CUST_NAME") & ": " & ErrorMessage,
                            "Email Invoices",
                             MessageBoxButtons.OK)

                End Select
            Else
                EmailInvoice = True
            End If

        Catch ex As Exception
            EmailInvoice = False
            ErrorMessage = ex.Message
        End Try

        ' Create an event for the emailing of invoices 
        For Each row As DataRow In tblSOTINVH1.Select("")

            For Each kvp As KeyValuePair(Of String, String) In EMAIL_ADDRESSs
                Dim v1 As String = kvp.Key
                Dim v2 As String = kvp.Value

                v2 = v2.Trim
                v2 = v2.Replace("'", "") ' Just incase
                If v2.Length = 0 Then Continue For

                Dim EVENT_DESC As String = String.Empty

                If ErrorMessage.Length = 0 Then
                    EVENT_DESC = "Invoice emailed to: " & v2
                Else
                    EVENT_DESC = "Error emailing " & v2 & " " & ErrorMessage
                End If

                EVENT_DESC = EVENT_DESC.Replace("'", "") ' Just incase
                If EVENT_DESC.Length > 500 Then
                    EVENT_DESC = EVENT_DESC.Substring(0, 500).Trim
                End If

                If EVENT_DESC.Length > 0 Then

                    ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) " & vbCrLf _
                            & " Select 'SOTORDR1', ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'EMAILINV', '" & EVENT_DESC & "', INV_NO" & vbCrLf _
                            & " from SOTINVH1 where INV_NO = '" & row.Item("INV_NO") & "'"
                    ASCDATA1.ExecuteSQL()
                End If
            Next
        Next

    End Function

    Private Sub GetNoShipShipments()

        Try
            If ASCMAIN1.CLIENT <> "INT" Then
                Exit Sub
            End If

            ASCMAIN1.sql = "(select conv.cfg_shiphdr.abspicknbr PICK_NO from conv.cfg_shiphdr" _
                & " minus " _
                & " select edi_pick_no PICK_NO from edt945t1)" _
                & " UNION" _
                & " SELECT EDI_PICK_NO PICK_NO FROM EDT945T1 WHERE EDI_PROCESS_IND = '0'"
            Dim noShip As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            ASCMAIN1.sql = " DELETE FROM " & noShip & " WHERE PICK_NO IN (SELECT PICK_NO FROM SOTPICK1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & noShip & ") AND PICK_STATUS NOT IN ('F', 'P'))"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE ORDR_NO IN" _
                & " (" _
                & " SELECT ORDR_NO FROM SOTPICK1 WHERE PICK_NO IN (SELECT PICK_NO FROM  " & noShip & ")" _
                & " )"
            Fill_Records("SOTORDR1_NOSHIP", String.Empty, True, ASCMAIN1.sql)

            ' Need to remove Orders Invoiceed in the session
            If dst.Tables("SOTINVH1").Rows.Count > 0 Then
                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1_NOSHIP").Select("")
                    Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO") & String.Empty
                    If dst.Tables("SOTINVH1").Select("ORDR_NO = '" & ORDR_NO & "'").Length > 0 Then
                        rowSOTORDR1.Delete()
                    End If
                Next
            End If

            dst.Tables("SOTORDR1_NOSHIP").AcceptChanges()

        Catch ex As Exception
            AddError("Get No-Ship Shipments Error: " & ex.Message, "1")

        End Try

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="SHIP_BOL_NO">Shipment Bol No</param>
    ''' <param name="emailAddress">Email addresses separated by semicolon ;</param>
    ''' <remarks></remarks>
    Private Sub Email_Shipment(ByVal SHIP_BOL_NO As String, ByVal emailAddress As String)

        Try

            If ASCMAIN1.Running_in_VS Then
                'Stop
            End If

            Dim RPT As String = "SORCART1"
            Dim SUBJECT As String = String.Empty

            Select Case ASCMAIN1.CLIENT
                Case "AHA"
                    SUBJECT = "Ahava Invoice " & SHIP_BOL_NO

                Case "INT"
                    'Ed.  Please do not forget that the subject line has to have:
                    'Interparfums Luxury Brands, Inc.
                    'And the PO number  Those are the two requirements.
                    Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                    If rowSOTSHIP1 Is Nothing Then
                        Exit Sub
                    End If
                    SUBJECT = "Interparfums Luxury Brands, Inc. PO " & rowSOTSHIP1.Item("ORDR_CUST_PO")

                Case Else
                    SUBJECT = "Shipment " & SHIP_BOL_NO
            End Select

            If Not REPORTS.ContainsKey(RPT) Then
                REPORTS.Add(RPT, Load_rptClass(RPT))
                REPORTS(RPT).Prepare_dst(False, "")
            End If

            REPORTS(RPT).Fill_Records_RPT(New String() {" and SOTSHIP1.SHIP_BOL_NO IN ('" & SHIP_BOL_NO & "')"})

            Dim REPORT_NO As String = ""

            With REPORTS(RPT).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                REPORT_NO = .Generate_Report(RPT, "Shipment Carton Information", , True, , , "PDF", SHIP_BOL_NO, False)
                .Print_Report_End(True, True)
            End With

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(SHIP_BOL_NO & ".pdf", ASCMAIN1.Folders("Temp") & SHIP_BOL_NO & ".pdf")

            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

            ' Concatentate and process all email addresses
            For Each sendToAddress As String In emailAddress.Split(";")
                sendToAddress = sendToAddress.Trim
                If sendToAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(sendToAddress) Then
                    EMAIL_ADDRESSs.Add(sendToAddress, sendToAddress)
                End If
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Exit Sub
            End If

            ASCMAIN1.Progress("Email Pack List", SHIP_BOL_NO)

            ASCMAIN1.Message = String.Empty
            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    SUBJECT, "PACKLIST", True, False, "SHIP_BOL_NO", SHIP_BOL_NO, "Shipment")

            Dim EVENT_DESC As String = $"Pack List emailed to: {emailAddress}"
            If SEND_NO.Length = 0 Then
                EVENT_DESC = $"Pack List NOT emailed to: {emailAddress}"
            End If
            EVENT_DESC = EVENT_DESC.Replace("'", "")
            If EVENT_DESC.Length > 500 Then
                EVENT_DESC = EVENT_DESC.Substring(0, 500).Trim
            End If

            Try
                ASCMAIN1.sql = $"INSERT INTO TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, FORM_NAME)  
                                    (Select 'SOTORDR1' TABLE_NAME, SOTPICK1.ORDR_NO TABLE_KEY,
                                    SYSDATE INIT_DATE, '{ASCMAIN1.USER_ID}' INIT_OPER,
                                    'SHPEML' EVENT_TYPE, '{EVENT_DESC}' EVENT_DESC, 'SORSHIPC' FORM_NAME
                                    FROM SOTSHIP1, SOTPICK1
                                    WHERE SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                                    AND SOTSHIP1.SHIP_BOL_NO = '{SHIP_BOL_NO }')"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Catch ex As Exception

            End Try

        Catch ex As Exception
            MessageBox.Show("Error generating Shipment email for shipment " & SHIP_BOL_NO & ". " & ex.Message, "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub Email_Shipment_Excel(ByVal SHIP_BOL_NO As String, ByVal emailAddress As String)

        ' SR-5537  04/04/2024

        Dim CUST_CODE As String = String.Empty
        Dim ORDR_CUST_PO As String = String.Empty

        Try
            If ASCMAIN1.Running_in_VS Then
                Stop
            End If

            Dim SUBJECT As String = String.Empty

            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
            If rowSOTSHIP1 Is Nothing Then
                Exit Sub
            End If

            CUST_CODE = rowSOTSHIP1.Item("CUST_CODE") & String.Empty
            ORDR_CUST_PO = rowSOTSHIP1.Item("ORDR_CUST_PO") & String.Empty

            Select Case ASCMAIN1.CLIENT
                Case "INT"
                    'Ed.  Please do not forget that the subject line has to have:
                    'Interparfums Luxury Brands, Inc.
                    'And the PO number  Those are the two requirements.
                    SUBJECT = $"Interparfums Luxury Brands, Inc. PO {ORDR_CUST_PO}"

                Case Else
                    SUBJECT = $"Shipment {SHIP_BOL_NO}, PO {ORDR_CUST_PO}"

            End Select

            dst.Tables("SOTCART2X").Rows.Clear()
            dst.Tables("SOTCART1X").Rows.Clear()
            Fill_Records("SOTCART1X", SHIP_BOL_NO)
            Fill_Records("SOTCART2X", SHIP_BOL_NO)
            grdSOTCART1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)

            grdSOTCART1.Text = $"Customer {rowSOTSHIP1.Item("CUST_CODE")}, P.O. {rowSOTSHIP1.Item("ORDR_CUST_PO")}, Cartons for Shipment {SHIP_BOL_NO}"
            grdSOTCART1.DisplayLayout.Bands(0).SortedColumns.Add("CART_NO", False)
            grdSOTCART1.DisplayLayout.Bands(1).SortedColumns.Add("CART_LNO", False)

            Set_Read_Only_for_ctl(grdSOTCART1, False)

            'For Each band As Infragistics.Win.UltraWinGrid.UltraGridBand In grdSOTCART1.DisplayLayout.Bands
            '    For Each col As Infragistics.Win.UltraWinGrid.UltraGridColumn In band.Columns
            '        col.CellAppearance.ForeColor = Drawing.Color.Black
            '    Next
            'Next
            Dim filenameAttach As String = SHIP_BOL_NO
            Dim OrdrCustPoAttach As String = ORDR_CUST_PO
            For Each ch As Char In "<>:""/\|?&*"
                OrdrCustPoAttach = OrdrCustPoAttach.Replace(ch, " ")
            Next

            If OrdrCustPoAttach.Length > 0 Then
                filenameAttach = "PO " & OrdrCustPoAttach & " packing list"
            End If

            Dim filename As String = System.IO.Path.Combine(ASCMAIN1.Folders("Work"), filenameAttach & ".XLS")
            Try
                If My.Computer.FileSystem.FileExists(filename) Then
                    My.Computer.FileSystem.DeleteFile(filename)
                End If
            Catch ex As Exception

            End Try
            Dim xlwb As Infragistics.Documents.Excel.Workbook = Export_to_Excel(grdSOTCART1, False, False, grdSOTCART1.Text)
            xlwb.Save(filename)

            Dim ATTACHMENTs As New Dictionary(Of String, String)
            ATTACHMENTs.Add(filenameAttach & ".XLS", filename)

            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

            ' Concatentate and process all email addresses
            For Each sendToAddress As String In emailAddress.Split(";")
                sendToAddress = sendToAddress.Trim
                If sendToAddress.Length > 5 AndAlso Not EMAIL_ADDRESSs.Keys.Contains(sendToAddress) Then
                    EMAIL_ADDRESSs.Add(sendToAddress, sendToAddress)
                End If
            Next

            If EMAIL_ADDRESSs.Count = 0 Then
                Exit Sub
            End If

            ASCMAIN1.Progress("Email Pack List", SHIP_BOL_NO)

            ASCMAIN1.Message = String.Empty
            Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                   (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    SUBJECT, "PACKLIST", True, False, "SHIP_BOL_NO", SHIP_BOL_NO, "Shipment")

            Dim EVENT_DESC As String = $"Pack List emailed to: {emailAddress}"

            If SEND_NO.Length = 0 Then
                EVENT_DESC = $"Pack List NOT emailed to: {emailAddress}"
            End If

            EVENT_DESC = EVENT_DESC.Replace("'", "")
            If EVENT_DESC.Length > 500 Then
                EVENT_DESC = EVENT_DESC.Substring(0, 500).Trim
            End If

            Try
                ASCMAIN1.sql = $"INSERT INTO TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, FORM_NAME)  
                                    (Select 'SOTORDR1' TABLE_NAME, SOTPICK1.ORDR_NO TABLE_KEY,
                                    SYSDATE INIT_DATE, '{ASCMAIN1.USER_ID}' INIT_OPER,
                                    'SHPEML' EVENT_TYPE, '{EVENT_DESC}' EVENT_DESC, 'SORSHIPC' FORM_NAME
                                    FROM SOTSHIP1, SOTPICK1
                                    WHERE SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                                    AND SOTSHIP1.SHIP_BOL_NO = '{SHIP_BOL_NO }')"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            Catch ex As Exception

            End Try

        Catch ex As Exception
            MessageBox.Show($"Error generating Shipment email for Shipment {SHIP_BOL_NO}, Customer {CUST_CODE}, PO {ORDR_CUST_PO} : {ex.Message}", "Email Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class