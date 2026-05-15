Public Class WHC850O1

    Inherits WHC000O1

    Private wkDirectory As String = Nothing
    Private ordrGroupNo As String = String.Empty
    Private tblEDTTRPM1 As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHC850O1"
        clsSuccessfulExecution = False

        Create_Work_Table()

        wkDirectory = ASCMAIN1.Folders("Work")
        If wkDirectory.Length > 0 AndAlso Not wkDirectory.EndsWith("\") Then
            wkDirectory &= "\"
        End If

        With dst

            Create_TDA(.Tables.Add, "SOTORDR0", "*")
            Create_TDA(.Tables.Add, "SOTORDR1", "*")
            Create_TDA(.Tables.Add, "SOTORDR2", "*")
            Create_TDA(.Tables.Add, "SOTORDR5", "*")

            Create_TDA(.Tables.Add, "EDT850T1", "*")
            Create_TDA(.Tables.Add, "EDT850T2", "*")
            Create_TDA(.Tables.Add, "EDT850T3", "*")
            Create_TDA(.Tables.Add, "EDT850T5", "*")

            Create_TDA(.Tables.Add, "ICTITEM1", "*")

        End With

        Main_Process()
    End Sub

    Public Sub Main_Process()

        Try
            EnforceConstraints(False)

            ' Global Variable that determines the number of records processed.
            R = 0

            Select Case G.APP_CMD & String.Empty
                Case "ORDR_GROUP_NO"
                    ordrGroupNo = G.APP_KEY & String.Empty

                Case Else
                    clsErrorMessage.Add("Missing Command Parameter")
                    Exit Sub
            End Select

            ASCMAIN1.sql = "SELECT ORDR_NO FROM SOTORDR1 WHERE ORDR_GROUP_NO = '" & ordrGroupNo & "'"
            Dim tempTableSOTPICKX As String = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

            Fill_Records("SOTORDR0", String.Empty, True, "Select * from SOTORDR0 where ORDR_GROUP_NO = '" & ordrGroupNo & "'")
            Fill_Records("SOTORDR1", String.Empty, True, "Select * from SOTORDR1 where ORDR_NO in (select ORDR_NO from " & tempTableSOTPICKX & ")")
            Fill_Records("SOTORDR2", String.Empty, True, "Select * from SOTORDR2 where ORDR_NO in (select ORDR_NO from " & tempTableSOTPICKX & ")")
            Fill_Records("SOTORDR5", String.Empty, True, "Select * from SOTORDR5 where ORDR_NO in (select ORDR_NO from " & tempTableSOTPICKX & ")")

            ASCMAIN1.sql = "SELECT * FROM ICTITEM1 WHERE ITEM_CODE IN " _
                & "( " _
                & " SELECT DISTINCT SOTORDR2.ITEM_CODE " _
                & " FROM SOTORDR1, SOTORDR2" _
                & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
                & " AND SOTORDR1.ORDR_GROUP_NO = '" & ordrGroupNo & "'" _
                & " )"
            Fill_Records("ICTITEM1", String.Empty, True, ASCMAIN1.sql)

            EnforceConstraints(True)

            If dst.Tables("SOTORDR1").Rows.Count = 0 Then
                clsSuccessfulExecution = True
                Exit Sub
            End If

            tblEDTSLSP1 = ASCDATA1.GetDataTable("Select * from EDTSLSP1")
            tblEDTSLSP1.PrimaryKey = New DataColumn() {tblEDTSLSP1.Columns("CUST_CODE")}

            Dim CUST_CODE As String = dst.Tables("SOTORDR1").Rows(0).Item("CUST_CODE")
            ASCMAIN1.sql = "SELECT * FROM EDTTRPM1 WHERE EDI_DOC_NO = '850' AND CUST_CODE = '" & CUST_CODE & "'"
            tblEDTTRPM1 = ASCDATA1.GetDataTable(ASCMAIN1.sql)
            If tblEDTTRPM1.Rows.Count = 0 Then
                clsSuccessfulExecution = True
                Exit Sub
            End If

            Update_Record()

        Catch ex As Exception
            ErrorMessages.Add(ex.Message)
        End Try

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Try

            Dim EDI_DOC_SEQ_NO As String = ASCMAIN1.Next_Control_No("EDT850T1.EDI_DOC_SEQ_NO")
            Dim rowSOTORDR0 As DataRow = dst.Tables("SOTORDR0").Rows(0)

            Dim EDI_ORD_QTY As Int32 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "") & String.Empty)
            Dim NUMBER_CHARS_STORE As Int32 = Val(tblEDTSLSP1.Rows(0).Item("NUMBER_CHARS_STORE") & String.Empty)
            Dim NUMBER_CHARS_DC As Int32 = Val(tblEDTSLSP1.Rows(0).Item("NUMBER_CHARS_DC") & String.Empty)

            ' Ceaate the EDI 850s
            Dim rowEDT850T1 As DataRow = dst.Tables("EDT850T1").NewRow
            rowEDT850T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
            'rowEDT850T1.Item("EDI_JRNL_NO") = String.Empty
            rowEDT850T1.Item("GEN_DOC_NO") = "9999999999"
            'rowEDT850T1.Item("EDI_ISA_NO") = String.Empty
            'rowEDT850T1.Item("EDI_TP_QUAL") = String.Empty
            'rowEDT850T1.Item("EDI_TP_ID") = String.Empty
            'rowEDT850T1.Item("EDI_OUR_QUAL") = String.Empty
            'rowEDT850T1.Item("EDI_OUR_ID") = String.Empty
            'rowEDT850T1.Item("EDI_CUSTOMER") = String.Empty
            'rowEDT850T1.Item("EDI_STORE") = String.Empty
            rowEDT850T1.Item("EDI_DEPARTMENT") = rowSOTORDR0.Item("ORDR_DEPT") & String.Empty
            rowEDT850T1.Item("EDI_PO_NO") = rowSOTORDR0.Item("ORDR_CUST_PO") & String.Empty
            'rowEDT850T1.Item("EDI_START_DATE") = String.Empty
            rowEDT850T1.Item("EDI_END_DATE") = rowSOTORDR0.Item("ORDR_CANCEL_DATE")
            'rowEDT850T1.Item("EDI_SHIP_DATE") = String.Empty
            'rowEDT850T1.Item("EDI_SHIP_DC") = String.Empty
            'rowEDT850T1.Item("EDI_CENTER_CODE") = String.Empty
            'rowEDT850T1.Item("EDI_SUPPLIER_NO") = String.Empty
            'rowEDT850T1.Item("EDI_PROMOTION") = String.Empty
            'rowEDT850T1.Item("EDI_BATCH_NO") = String.Empty
            'rowEDT850T1.Item("EDI_MERCH_TYPE") = String.Empty
            'rowEDT850T1.Item("EDI_FOB") = String.Empty
            'rowEDT850T1.Item("EDI_TERMS") = String.Empty
            'rowEDT850T1.Item("EDI_TERM_TYPE") = String.Empty
            'rowEDT850T1.Item("EDI_TERM_BASIS") = String.Empty
            'rowEDT850T1.Item("EDI_TERM_RATE") = String.Empty
            'rowEDT850T1.Item("EDI_TERM_DSCDAYS") = String.Empty
            'rowEDT850T1.Item("EDI_TERM_NETDAYS") = String.Empty
            'rowEDT850T1.Item("EDI_TERM_DESC") = String.Empty
            'rowEDT850T1.Item("EDI_TERM_DOM") = String.Empty
            rowEDT850T1.Item("EDI_PO_PURP") = "00"
            rowEDT850T1.Item("EDI_PO_DATE") = rowSOTORDR0.Item("ORDR_DATE")
            rowEDT850T1.Item("EDI_PO_TYPE") = "SA"
            'rowEDT850T1.Item("EDI_SHIPPER") = String.Empty
            rowEDT850T1.Item("EDI_RECEIVED_DATE") = rowSOTORDR0.Item("ORDR_DATE_RECD")
            'rowEDT850T1.Item("EDI_SHIP_ADDR_TYPE") = String.Empty
            'rowEDT850T1.Item("EDI_CURRENCY") = String.Empty
            rowEDT850T1.Item("EDI_PROCESS_IND") = "1"
            rowEDT850T1.Item("EDI_ARRIVAL_DATE") = rowSOTORDR0.Item("ORDR_ARRIVAL_DATE")
            rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") = rowSOTORDR0.Item("ORDR_LAST_ARRIVAL_DATE")

            ' Done Below
            'rowEDT850T1.Item("EDI_NO_OF_LINES") = String.Empty

            'rowEDT850T1.Item("EDI_PRICE_BRACKET_ID") = String.Empty
            'rowEDT850T1.Item("STORE_GLOBAL_LOCATION_NUMBER") = String.Empty
            'rowEDT850T1.Item("DC_GLOBAL_LOCATION_NUMBER") = String.Empty
            'rowEDT850T1.Item("EDI_APPOINTMENT") = String.Empty
            'rowEDT850T1.Item("EDI_CARTON") = String.Empty
            'rowEDT850T1.Item("EDI_CURR_CODE") = String.Empty
            'rowEDT850T1.Item("EDI_DEPT_DESC") = String.Empty
            rowEDT850T1.Item("EDI_ORD_QTY") = EDI_ORD_QTY
            'rowEDT850T1.Item("EDI_WEIGHT") = String.Empty
            rowEDT850T1.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
            rowEDT850T1.Item("CUST_CODE") = rowSOTORDR0.Item("CUST_CODE")
            rowEDT850T1.Item("INIT_DATE") = DateTime.Now
            rowEDT850T1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowEDT850T1.Item("LAST_DATE") = DateTime.Now
            rowEDT850T1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            'rowEDT850T1.Item("EDI_PO_RELEASE_NO") = String.Empty
            'rowEDT850T1.Item("EDI_PO_TOTAL_AMT") = String.Empty
            'rowEDT850T1.Item("EDI_CHAIN") = String.Empty
            'rowEDT850T1.Item("EDI_FACILITY") = String.Empty
            'rowEDT850T1.Item("EDI_CONTRACT_NO") = String.Empty
            dst.Tables("EDT850T1").Rows.Add(rowEDT850T1)

            Dim EDI_DTL_SEQ As Int32 = 0
            For Each rowICTITEM1 As DataRow In dst.Tables("ICTITEM1").Select("", "ITEM_CODE")
                Dim ITEM_CODE As String = rowICTITEM1.Item("ITEM_CODE")
                Dim rowEDT850T2 As DataRow = dst.Tables("EDT850T2").NewRow
                rowEDT850T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                EDI_DTL_SEQ += 1
                rowEDT850T2.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                rowEDT850T2.Item("GEN_DOC_NO") = "9999999999"
                'rowEDT850T2.Item("EDI_BRAND") = String.Empty
                rowEDT850T2.Item("EDI_ITEM") = ITEM_CODE
                'rowEDT850T2.Item("EDI_DIMENSION") = String.Empty
                'rowEDT850T2.Item("EDI_SIZE_DESC") = String.Empty
                rowEDT850T2.Item("EDI_UPC") = rowICTITEM1.Item("ITEM_UPC_CODE") & String.Empty
                'rowEDT850T2.Item("EDI_SKU") = String.Empty
                rowEDT850T2.Item("EDI_PO4_UOM") = rowICTITEM1.Item("ITEM_UOM") & String.Empty
                rowEDT850T2.Item("EDI_PRICE") = Val(dst.Tables("SOTORDR2").Compute("MAX(ORDR_UNIT_PRICE)", "ITEM_CODE = '" & ITEM_CODE & "'") & String.Empty)
                rowEDT850T2.Item("EDI_TOTAL_QTY") = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY)", "ITEM_CODE = '" & ITEM_CODE & "'") & String.Empty)
                'rowEDT850T2.Item("EDI_ITEM_NAME") = String.Empty
                'rowEDT850T2.Item("EDI_START_DATE") = String.Empty
                'rowEDT850T2.Item("EDI_END_DATE") = String.Empty
                'rowEDT850T2.Item("EDI_PO4_QTY") = String.Empty
                'rowEDT850T2.Item("EDI_ITEM_DESC") = String.Empty
                'rowEDT850T2.Item("EDI_PRICE_UOM") = String.Empty
                rowEDT850T2.Item("EDI_EAN") = rowICTITEM1.Item("ITEM_EAN_CODE") & String.Empty
                'rowEDT850T2.Item("EDI_PRICE_BRACKET_ID") = String.Empty
                'rowEDT850T2.Item("EDI_PO4_INNER") = String.Empty
                'rowEDT850T2.Item("EDI_GTIN") = String.Empty
                'rowEDT850T2.Item("EDI_COLOR_CODE") = String.Empty
                'rowEDT850T2.Item("EDI_COLOR_NAME") = String.Empty
                'rowEDT850T2.Item("EDI_DIVISION") = String.Empty
                'rowEDT850T2.Item("EDI_LBL_CODE") = String.Empty
                'rowEDT850T2.Item("EDI_STYLE") = String.Empty
                'rowEDT850T2.Item("EDI_STYLE_NAME") = String.Empty
                'rowEDT850T2.Item("EDI_RETAIL_PRICE") = String.Empty
                'rowEDT850T2.Item("EDI_TOTAL_AMT") = String.Empty
                'rowEDT850T2.Item("EDI_CARTON_GRP") = String.Empty
                'rowEDT850T2.Item("EDI_PO_LNO") = String.Empty
                dst.Tables("EDT850T2").Rows.Add(rowEDT850T2)
            Next

            rowEDT850T1.Item("EDI_NO_OF_LINES") = dst.Tables("EDT850T2").Rows.Count

            Dim EDI_SDQ_SEQ As Int32 = 0
            Dim rowEDT850T3 As DataRow = Nothing
            For Each rowEDT850T2 As DataRow In dst.Tables("EDT850T2").Select("", "EDI_DTL_SEQ")
                EDI_DTL_SEQ = rowEDT850T2.Item("EDI_DTL_SEQ")
                EDI_SDQ_SEQ = 0

                Dim EDI_ITEM As String = rowEDT850T2.Item("EDI_ITEM")
                Dim fieldNum As Int16 = 0

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ITEM_CODE = '" & EDI_ITEM & "' AND ORDR_QTY > 0", "CUST_STORE_NO")
                    Dim ORDR_NO As String = rowSOTORDR2.Item("ORDR_NO")
                    Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                    Dim ORDR_ADDR_TYPE_ST As String = rowSOTORDR1.Item("ORDR_ADDR_TYPE_ST") & String.Empty
                    Dim EDI_STORE As String = String.Empty
                    Dim storeMaxLen As Int32 = 0

                    EDI_STORE = rowSOTORDR1.Item("CUST_STORE_NO") & String.Empty
                    storeMaxLen = NUMBER_CHARS_STORE

                    If storeMaxLen > 0 And EDI_STORE.Length > storeMaxLen Then
                        EDI_STORE = StrReverse(StrReverse(EDI_STORE).Substring(0, storeMaxLen))
                    End If

                    If IsNumeric(EDI_STORE) AndAlso storeMaxLen > 0 Then
                        EDI_STORE = EDI_STORE.PadLeft(storeMaxLen, "0")
                    End If

                    fieldNum += 1

                    Select Case fieldNum
                        Case 1
                            rowEDT850T3 = dst.Tables("EDT850T3").NewRow
                            rowEDT850T3.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                            rowEDT850T3.Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                            EDI_SDQ_SEQ += 1
                            rowEDT850T3.Item("EDI_SDQ_SEQ") = EDI_SDQ_SEQ
                            rowEDT850T3.Item("GEN_DOC_NO") = "9999999999"
                            'rowEDT850T3.Item("EDI_SDQ_UOM") = String.Empty
                            'rowEDT850T3.Item("EDI_SDQ_QUAL") = String.Empty
                            dst.Tables("EDT850T3").Rows.Add(rowEDT850T3)

                    End Select

                    rowEDT850T3.Item("EDI_STORE_" & fieldNum.ToString("00")) = EDI_STORE
                    rowEDT850T3.Item("EDI_QTY_" & fieldNum.ToString("00")) = rowSOTORDR2.Item("ORDR_QTY")

                    If fieldNum = 10 Then
                        fieldNum = 0
                    End If

                Next
            Next

            ' Global Variable that determine the number of records processed.
            ' Needed so Update_Archive is called from base class.
            R = dst.Tables("EDT850T1").Rows.Count
        Catch ex As Exception
            ErrorMessages.Add(ex.Message)
        End Try

    End Sub

    Overrides Sub Update_Archive()
        MyBase.Update_Archive()

        For Each tableName As String In New String() {"EDT850T1", "EDT850T2", "EDT850T3", "EDT850T5"}
            Update_Record_TDA(tableName)
        Next

    End Sub

    Sub Create_Work_Table()

    End Sub

    Overrides Sub Post_Update_Archive()
        MyBase.Post_Update_Archive()

        Try
            clsSuccessfulExecution = True
        Catch ex As Exception
        End Try

    End Sub

End Class