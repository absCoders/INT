Public Class EDC855O1
    Private rowARTPARM1 As DataRow = Nothing
    Private Const EDI_PROCESS_IND As String = "1"

    Private EDI_OUTBOUND_DOC_NO As String = String.Empty

    Private tblSOTSVIA1 As DataTable = Nothing
    Private tblTATTERM1 As DataTable = Nothing
    Private tblEDTTRPM1 As DataTable = Nothing
    Private tblWHTPKGM1 As DataTable = Nothing
    Private tblEDTSLSP1 As DataTable = Nothing
    Private Shared tblICTITEM1 As DataTable = Nothing

    Private ASCDATA1 As New WHC.ASCDATA1
    Private ASCMAIN1 As New WHC.ASCMAIN1

    Public Sub New(ByRef CompanyCode As String, ByVal UserID As String)
        ASCMAIN1.DBS_COMPANY = CompanyCode
        ASCMAIN1.USER_ID = UserID
    End Sub

    Public Function CreateEDTSYSIH(ByRef dst As DataSet,
                                    ByVal ediOutboundDocNo As String,
                                    ByVal EDI_OUR_ID As String,
                                    ByVal EDI_TP_ID As String,
                                    ByVal ediApplicationID As String,
                                    ByVal EDI_STATUS As String) As String

        CreateEDTSYSIH = String.Empty

        Dim rowEDTSYSIH As DataRow = dst.Tables("EDTSYSIH").NewRow
        rowEDTSYSIH.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
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
        dst.Tables("EDTSYSIH").Rows.Add(rowEDTSYSIH)

        CreateEDTSYSIH = ediOutboundDocNo

    End Function

    Public Sub Generate_855(ByVal clsASCBASE1 As ASCBASE1,
                                   ByVal ORDR_GROUP_NO As String)

        Dim EDI_OUR_ID As String = String.Empty
        Dim EDI_DOC_SEQ_NO As String = String.Empty
        Dim EDI_PURP_CODE As String = String.Empty
        Dim Err_Msg As String = String.Empty

        ' New code for Lord and Taylor
        Dim populateDCandGtin As Boolean = False

        If clsASCBASE1.dst.Tables.Contains("EDT855O1") Then
            clsASCBASE1.dst.Tables("EDT855O1").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O2").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O3").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O5").Rows.Clear()
            clsASCBASE1.dst.Tables("EDT855O7").Rows.Clear()
            clsASCBASE1.dst.Tables("EDTSYSIH").Rows.Clear()
            ' New code for Lord and Taylor
            clsASCBASE1.dst.Tables("SOTSHIP1_855").Rows.Clear()
        Else
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O1", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O2", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O3", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O5", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDT855O7", "*")
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add, "EDTSYSIH", "*")
            ' New code for Lord and Taylor
            clsASCBASE1.Create_TDA(clsASCBASE1.dst.Tables.Add("SOTSHIP1_855"), "SOTSHIP1", "*")
        End If

        ASCMAIN1.sql = "Select * from SOTORDR1 R1, EDT850T1 T1" _
        & " Where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
        & " And R1.EDI_DOC_SEQ_NO is Not Null" _
        & " And R1.EDI_DOC_SEQ_NO = T1.EDI_DOC_SEQ_NO"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select DISTINCT ICTITEM1.* " _
            & " FROM ICTITEM1, SOTORDR1, SOTORDR2" _
            & " WHERE SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
            & " AND SOTORDR2.ITEM_CODE = ICTITEM1.ITEM_CODE" _
            & " AND SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
        tblICTITEM1 = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        ' New Reverse PO driven by EDTSLSP1
        Dim tblEDTSLSP1 As DataTable = ASCDATA1.GetDataTable("SELECT * FROM EDTSLSP1 WHERE ALLOW_REVERSE_PO = '1'")
        Dim CUST_CODE As String = String.Empty
        If rowSOTORDR1 IsNot Nothing Then
            CUST_CODE = rowSOTORDR1.Item("CUST_CODE") & String.Empty
        Else
            CUST_CODE = ASCDATA1.GetDataValue("SELECT CUST_CODE FROM SOTORDR0 WHERE ORDR_GROUP_NO = :PARM1", "V", New Object() {ORDR_GROUP_NO})
        End If

        ' New Code for Reverse 850
        If rowSOTORDR1 IsNot Nothing _
            AndAlso tblEDTSLSP1.Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 AndAlso rowSOTORDR1.Item("REVERSE_PO") & String.Empty <> "1" Then
            Exit Sub
        End If

        Dim REVERSE_PO As Boolean = False
        If rowSOTORDR1 IsNot Nothing Then
            REVERSE_PO = rowSOTORDR1.Item("REVERSE_PO") & String.Empty = "1"
        Else
            Dim rowSOTORDR1X As DataRow = ASCDATA1.GetDataRow("SELECT * FROM SOTORDR1 WHERE ORDR_GROUP_NO = :PARM1", "V", ORDR_GROUP_NO)
            REVERSE_PO = rowSOTORDR1X.Item("REVERSE_PO") & String.Empty = "1"
        End If

        If rowSOTORDR1 Is Nothing Then
            Exit Sub
        End If

        EDI_DOC_SEQ_NO = rowSOTORDR1.Item("EDI_DOC_SEQ_NO") & ""

        Dim EDI_Outbound_Doc_No As String = ASCMAIN1.Next_Control_No("EDTSYSIH.EDI_OUTBOUND_DOC_NO")
        EDI_OUR_ID = Replace(rowSOTORDR1.Item("EDI_OUR_ID"), " ", "")

        ' In the future JCP has different values or some of their stores; therefore, this first query may return the wrong credentials.
        ASCMAIN1.sql = "Select * from EDTTRPM1 M1" & vbCrLf _
        & " Where EDI_DOC_NO = '855'" & vbCrLf _
        & " And EDI_TP_QUAL = '" & rowSOTORDR1.Item("EDI_TP_QUAL") & "'" & vbCrLf _
        & " And EDI_TP_ID = rtrim('" & rowSOTORDR1.Item("EDI_TP_ID") & "')" & vbCrLf
        Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow

        ' Ahava MSS sends the PO but 855 goes back to VCS
        If rowEDTTRPM1 Is Nothing Then
            ASCMAIN1.sql = "Select * from EDTTRPM1" & vbCrLf _
            & " Where EDI_DOC_NO = '855'" & vbCrLf _
            & " And CUST_CODE = '" & rowSOTORDR1.Item("CUST_CODE") & "'"
            rowEDTTRPM1 = ASCDATA1.GetDataRow
        End If

        If rowEDTTRPM1 Is Nothing Then
            Exit Sub
        End If

        Dim EDI_TP_ID As String = rowEDTTRPM1.Item("EDI_TP_ID")
        Dim Ack_Type As String = "AC"

        ' New code for Lord and Taylor
        If EDI_TP_ID = "6111492199" Then
            populateDCandGtin = True
            ASCMAIN1.sql = "SELECT * FROM SOTSHIP1 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            clsASCBASE1.Fill_Records("SOTSHIP1_855", String.Empty, True, ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select IH.* from EDT855O1 O1, EDTSYSIH IH" & vbCrLf _
        & " Where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
        & " and O1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" & vbCrLf _
        & " And O1.COMPANY_CODE = IH.COMPANY_CODE" & vbCrLf _
        & " And O1.EDI_OUTBOUND_DOC_NO = IH.EDI_OUTBOUND_DOC_NO"
        Dim rowEDT855OC As DataRow = ASCDATA1.GetDataRow

        If rowEDT855OC IsNot Nothing Then
            EDI_PURP_CODE = "05"
        Else
            EDI_PURP_CODE = "00"
        End If

        Dim NUMBER_CHARS_STORE As Int32 = 4
        Dim NUMBER_CHARS_DC As Int32 = 4

        If tblEDTSLSP1.Select("CUST_CODE = '" & CUST_CODE & "'").Length > 0 Then
            Dim rowEDTSLSP1 As DataRow = tblEDTSLSP1.Select("CUST_CODE = '" & CUST_CODE & "'")(0)
            NUMBER_CHARS_STORE = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
            NUMBER_CHARS_DC = Val(rowEDTSLSP1.Item("NUMBER_CHARS_DC") & String.Empty)
        End If

        Dim rowEDT855O1 As DataRow = clsASCBASE1.dst.Tables("EDT855O1").NewRow  ' TBLs("EDT810O1").NewRow
        With rowEDT855O1
            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
            .Item("ORDR_CUST_PO") = rowSOTORDR1.Item("ORDR_CUST_PO")
            .Item("ORDR_PO_DATE") = rowSOTORDR1.Item("EDI_PO_DATE")
            .Item("ORDR_GROUP_NO") = rowSOTORDR1.Item("ORDR_GROUP_NO")
            .Item("REQUEST_DATE") = Format(Now, "dd-MMM-yy")
            .Item("TERM_CODE") = rowSOTORDR1.Item("TERM_CODE")
            .Item("ORDR_SHIP_DATE") = rowSOTORDR1.Item("ORDR_SHIP_DATE")
            .Item("ORDR_CANCEL_DATE") = rowSOTORDR1.Item("ORDR_CANCEL_DATE")
            .Item("AS_OF_DATE") = Format(Now, "dd-MMM-yy")
            .Item("INIT_DATE") = Format(Now, "dd-MMM-yy")
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("EDI_PURPOSE_CODE") = EDI_PURP_CODE

            .Item("EDI_START_DATE_ORIG") = rowSOTORDR1.Item("EDI_START_DATE")
            .Item("EDI_END_DATE_ORIG") = rowSOTORDR1.Item("EDI_END_DATE")
            .Item("EDI_SHIP_DATE_ORIG") = rowSOTORDR1.Item("EDI_SHIP_DATE")
            .Item("EDI_ARRIVAL_DATE_ORIG") = rowSOTORDR1.Item("EDI_ARRIVAL_DATE")

            If IsDate(rowSOTORDR1.Item("ORDR_ARRIVAL_DATE") & String.Empty) Then
                .Item("EDI_ARRIVAL_DATE") = rowSOTORDR1.Item("ORDR_ARRIVAL_DATE")
            Else
                .Item("EDI_ARRIVAL_DATE") = rowSOTORDR1.Item("EDI_ARRIVAL_DATE")
            End If

            .Item("EDI_ACK_TYPE") = Ack_Type
            .Item("EDI_SUPPLIER_NO") = rowSOTORDR1.Item("EDI_SUPPLIER_NO")
            .Item("EDI_DEPARTMENT") = rowSOTORDR1.Item("ORDR_DEPT")

            ' New code for Lord and Taylor
            If populateDCandGtin Then
                If clsASCBASE1.dst.Tables("SOTSHIP1_855").Select("SHIP_ADDR_TYPE = 'DC'").Length > 0 Then
                    Dim EDI_SHIP_DC As String = clsASCBASE1.dst.Tables("SOTSHIP1_855").Select("SHIP_ADDR_TYPE = 'DC'")(0).Item("SHIP_ADDR_CODE") & String.Empty
                    EDI_SHIP_DC = EDI_SHIP_DC.PadLeft(NUMBER_CHARS_DC, "0")
                    If EDI_SHIP_DC.Length > NUMBER_CHARS_DC Then
                        EDI_SHIP_DC = StrReverse(StrReverse(EDI_SHIP_DC).Substring(0, 4))
                    End If
                    .Item("EDI_SHIP_DC") = EDI_SHIP_DC
                Else
                    .Item("EDI_SHIP_DC") = "9999"
                End If
            End If
        End With
        clsASCBASE1.dst.Tables("EDT855O1").Rows.Add(rowEDT855O1)

        ASCMAIN1.sql = "Select * from EDT850T2" & vbCrLf _
        & " Where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        For Each rowEDT850T2 As DataRow In ASCDATA1.GetDataTable.Rows
            If Add_EDT855O2(clsASCBASE1, rowEDT850T2, ORDR_GROUP_NO, EDI_Outbound_Doc_No, EDI_TP_ID, REVERSE_PO) Then
                Add_EDT855O3(clsASCBASE1, rowEDT850T2, ORDR_GROUP_NO, EDI_Outbound_Doc_No)
            End If
        Next

        ASCMAIN1.sql = "Select * from EDT850T5 " & vbCrLf _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
        & " And EDI_ADDR_TYPE <> 'SF' "
        For Each rowEDT850T5 As DataRow In ASCDATA1.GetDataTable.Rows

            Dim rowEDT855O5 As DataRow = clsASCBASE1.dst.Tables("EDT855O5").NewRow  ' TBLs("EDT810O1").NewRow
            With rowEDT855O5
                .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                .Item("EDI_ADR_SEQ") = rowEDT850T5.Item("EDI_ADR_SEQ") & ""
                .Item("EDI_ADDR_TYPE") = rowEDT850T5.Item("EDI_ADDR_TYPE") & ""
                .Item("EDI_CUST_NAME_ADR") = rowEDT850T5.Item("EDI_CUST_NAME_ADR") & ""
                .Item("EDI_ADDRESS1") = rowEDT850T5.Item("EDI_ADDRESS1") & ""
                .Item("EDI_ADDRESS2") = rowEDT850T5.Item("EDI_ADDRESS2") & ""
                .Item("EDI_ADDRESS3") = ""
                .Item("EDI_CITY") = rowEDT850T5.Item("EDI_CITY") & ""
                .Item("EDI_STATE") = rowEDT850T5.Item("EDI_STATE") & ""
                .Item("EDI_ZIPCODE") = rowEDT850T5.Item("EDI_ZIPCODE") & ""
                .Item("EDI_COUNTRY") = rowEDT850T5.Item("EDI_COUNTRY") & ""
                .Item("EDI_ADDR_CODE") = rowEDT850T5.Item("EDI_ADDR_CODE") & ""
                .Item("EDI_ADDR_CODE_QUAL") = rowEDT850T5.Item("EDI_ADDR_CODE_QUAL") & ""
            End With
            clsASCBASE1.dst.Tables("EDT855O5").Rows.Add(rowEDT855O5)
        Next

        ASCMAIN1.sql = "Select * from EDT850T7 " & vbCrLf _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
        For Each rowEDT850T7 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowEDT855O7 As DataRow = clsASCBASE1.dst.Tables("EDT855O7").NewRow  ' TBLs("EDT810O1").NewRow
            With rowEDT855O7
                .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                .Item("SAH_SEQ_NO") = rowEDT850T7.Item("SAH_SEQ_NO") & ""
                .Item("SAH_ALLOW_IND") = rowEDT850T7.Item("SAH_ALLOW_IND") & ""
                .Item("SAH_ALLOW_CODE") = rowEDT850T7.Item("SAH_ALLOW_CODE") & ""
                .Item("SAH_AMOUNT") = Val(rowEDT850T7.Item("SAH_AMOUNT") & "")
                .Item("SAH_PERCENT_QUAL") = rowEDT850T7.Item("SAH_PERCENT_QUAL") & ""
                .Item("SAH_PERCENT") = rowEDT850T7.Item("SAH_PERCENT")
                .Item("SAH_RATE") = rowEDT850T7.Item("SAH_RATE")
                .Item("SAH_UOM_CODE") = rowEDT850T7.Item("SAH_UOM_CODE") & ""
                .Item("SAH_QTY") = rowEDT850T7.Item("SAH_QTY")
                .Item("SAH_HANDLING_CODE") = rowEDT850T7.Item("SAH_HANDLING_CODE") & ""
                .Item("SAH_DESC") = rowEDT850T7.Item("SAH_DESC") & ""
            End With
            clsASCBASE1.dst.Tables("EDT855O7").Rows.Add(rowEDT855O7)
        Next

        If REVERSE_PO AndAlso clsASCBASE1.dst.Tables("EDT855O2").Select("").Length = 0 Then
            For Each tablename As String In New String() {"EDT855O1", "EDT855O2", "EDT855O3", "EDT855O5", "EDT855O7", "EDTSYSIH"}
                clsASCBASE1.dst.Tables(tablename).Clear()
            Next
            Exit Sub
        End If

        clsASCBASE1.Update_Record_TDA("EDT855O1")
        clsASCBASE1.Update_Record_TDA("EDT855O2")
        clsASCBASE1.Update_Record_TDA("EDT855O3")
        clsASCBASE1.Update_Record_TDA("EDT855O5")
        clsASCBASE1.Update_Record_TDA("EDT855O7")

        EDI_Outbound_Doc_No = CreateEDTSYSIH(clsASCBASE1.dst, EDI_Outbound_Doc_No, EDI_OUR_ID, EDI_TP_ID, "PR", rowEDTTRPM1.Item("EDI_STATUS") & String.Empty)
        clsASCBASE1.Update_Record_TDA("EDTSYSIH")

    End Sub

    Public Function Add_EDT855O2(ByVal clsASCBASE1 As ASCBASE1,
                  ByVal rowEDT850T2 As DataRow,
                  ByVal Ordr_Group_No As String,
                  ByVal EDI_Outbound_Doc_No As String,
                  ByVal EDI_TP_ID As String,
                  ByVal REVERSE_PO As Boolean) As Boolean

        Dim EDI_PRICE_ACTUAL As Double = 0
        Dim EDI_QTY_OPEN As Double = 0
        Dim EDI_QTY_PICK As Long = 0
        Dim EDI_QTY_CANCEL As Double = 0

        Dim rowEDT855O2 As DataRow = clsASCBASE1.dst.Tables("EDT855O2").NewRow
        With rowEDT855O2
            ' Need Sum because we have Multi Store Sales Order Groups
            ASCMAIN1.sql = $"Select MAX(ORDR_UNIT_PRICE) ORDR_UNIT_PRICE, 
                                SUM(DECODE(SOTORDR1.ORDR_STATUS, 'P', ORDR_QTY_OPEN, 0)) ORDR_QTY_OPEN, 
                                SUM(DECODE(SOTORDR1.ORDR_STATUS, 'P', ORDR_QTY_PICK, ORDR_QTY_OPEN)) ORDR_QTY_PICK, 
                                SUM(ORDR_QTY_CANC) ORDR_QTY_CANC
                 From SOTORDR2, SOTORDR1
                 Where SOTORDR1.ORDR_GROUP_NO = '{Ordr_Group_No}'
                 And SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO
                 And EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")
            Dim rowSOTORDR2 As DataRow = ASCDATA1.GetDataRow

            If rowSOTORDR2 IsNot Nothing Then
                EDI_PRICE_ACTUAL = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                EDI_QTY_OPEN = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                EDI_QTY_PICK = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                EDI_QTY_CANCEL = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")
            Else
                EDI_PRICE_ACTUAL = 0
                EDI_QTY_OPEN = 0
                EDI_QTY_PICK = 0
                EDI_QTY_CANCEL = 0
            End If

            If REVERSE_PO AndAlso EDI_QTY_PICK = 0 Then
                Return False
            End If

            .Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
            .Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
            .Item("EDI_DTL_SEQ") = rowEDT850T2.Item("EDI_DTL_SEQ")
            .Item("EDI_TOTAL_QTY") = rowEDT850T2.Item("EDI_TOTAL_QTY")
            .Item("EDI_UOM") = rowEDT850T2.Item("EDI_PRICE_UOM") & ""
            .Item("EDI_PRICE") = rowEDT850T2.Item("EDI_PRICE")
            .Item("EDI_PO4_QTY") = rowEDT850T2.Item("EDI_PO4_QTY")
            .Item("EDI_PO4_INNER") = rowEDT850T2.Item("EDI_PO4_INNER")
            .Item("EDI_PO4_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
            .Item("EDI_ITEM") = rowEDT850T2.Item("EDI_ITEM")
            .Item("EDI_UPC") = rowEDT850T2.Item("EDI_UPC")

            ' Added on 12/24/2015 as per maria
            .Item("EDI_EAN") = rowEDT850T2.Item("EDI_EAN")

            .Item("EDI_SKU") = rowEDT850T2.Item("EDI_SKU")
            .Item("EDI_GTIN") = rowEDT850T2.Item("EDI_GTIN")
            .Item("EDI_ITEM_DESC") = rowEDT850T2.Item("EDI_ITEM_DESC")
            .Item("EDI_PO_LNO") = rowEDT850T2.Item("EDI_PO_LNO")
            .Item("EDI_PRICE_ACTUAL") = EDI_PRICE_ACTUAL
            .Item("EDI_QTY_OPEN") = EDI_QTY_OPEN
            .Item("EDI_QTY_PICK") = EDI_QTY_PICK
            .Item("EDI_QTY_CANC") = EDI_QTY_CANCEL
            .Item("EDI_DIMENSION") = rowEDT850T2.Item("EDI_DIMENSION") & ""
        End With
        clsASCBASE1.dst.Tables("EDT855O2").Rows.Add(rowEDT855O2)
        Return True
    End Function

    Public Sub Add_EDT855O3(ByVal clsASCBASE1 As ASCBASE1,
               ByVal rowEDT850T2 As DataRow,
               ByVal Ordr_Group_No As String,
               ByVal EDI_Outbound_Doc_No As String)

        Dim rowEDT855O3 As DataRow = Nothing
        Dim EDI_SDQ_SEQ As Int32 = 0
        Dim fieldNum As Int16 = 0
        Dim ITEM_CODE As String = rowEDT850T2.Item("EDI_ITEM") & String.Empty

        If ITEM_CODE.Length = 0 Then
            Dim ITEM_UPC_CODE As String = rowEDT850T2.Item("EDI_UPC") & String.Empty
            If ITEM_UPC_CODE.Length > 0 Then
                If tblICTITEM1.Select("ITEM_UPC_CODE = '" & ITEM_UPC_CODE & "'").Length > 0 Then
                    ITEM_CODE = tblICTITEM1.Select("ITEM_UPC_CODE = '" & ITEM_UPC_CODE & "'")(0).Item("ITEM_CODE") & String.Empty
                End If
            End If
        End If

        For Each rowSOTORDR1 As DataRow In clsASCBASE1.dst.Tables("SOTORDR1").Select("ORDR_GROUP_NO = '" & Ordr_Group_No & "'")

            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")
            ASCMAIN1.sql = "SELECT * FROM EDTSLSP1 WHERE CUST_CODE = '" & CUST_CODE & "'"
            Dim rowEDTSLSP1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql)
            Dim storeMaxLen As Int32 = Val(rowEDTSLSP1.Item("NUMBER_CHARS_STORE") & String.Empty)
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            Dim sql As String = String.Empty

            Select Case rowSOTORDR1.Item("ORDR_STATUS") & String.Empty
                Case "O"
                    sql = "ORDR_NO = '" & ORDR_NO & "'" _
                                & " AND ORDR_GROUP_NO = '" & Ordr_Group_No & "'" _
                                & " AND ORDR_QTY_OPEN > 0" _
                                & " AND ITEM_CODE = '" & ITEM_CODE & "'" _
                                & " AND EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "'" _
                                & " AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")
                Case Else
                    sql = "ORDR_NO = '" & ORDR_NO & "'" _
                                & " AND ORDR_GROUP_NO = '" & Ordr_Group_No & "'" _
                                & " AND ORDR_QTY_PICK > 0" _
                                & " AND ITEM_CODE = '" & ITEM_CODE & "'" _
                                & " AND EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "'" _
                                & " AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")

            End Select


            If Not clsASCBASE1.dst.Tables("SOTORDR2").Columns.Contains("ORDR_GROUP_NO") Then
                Select Case rowSOTORDR1.Item("ORDR_STATUS") & String.Empty
                    Case "O"
                        sql = "ORDR_NO = '" & ORDR_NO & "'" _
                                    & " AND ORDR_QTY_OPEN > 0" _
                                    & " AND ITEM_CODE = '" & ITEM_CODE & "'" _
                                    & " AND EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "'" _
                                    & " AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")

                    Case Else
                        sql = "ORDR_NO = '" & ORDR_NO & "'" _
                    & " AND ORDR_QTY_PICK > 0" _
                    & " AND ITEM_CODE = '" & ITEM_CODE & "'" _
                    & " AND EDI_DOC_SEQ_NO = '" & rowEDT850T2.Item("EDI_DOC_SEQ_NO") & "'" _
                    & " AND EDI_DTL_SEQ = " & rowEDT850T2.Item("EDI_DTL_SEQ")

                End Select

            End If

            For Each rowSOTORDR2 As DataRow In clsASCBASE1.dst.Tables("SOTORDR2").Select(sql, "CUST_STORE_NO")
                Dim EDI_STORE As String = rowSOTORDR2.Item("CUST_STORE_NO") & String.Empty

                If storeMaxLen > 0 And EDI_STORE.Length > storeMaxLen Then
                    EDI_STORE = StrReverse(StrReverse(EDI_STORE).Substring(0, storeMaxLen))
                End If

                If IsNumeric(EDI_STORE) AndAlso storeMaxLen > 0 Then
                    EDI_STORE = EDI_STORE.PadLeft(storeMaxLen, "0")
                End If

                fieldNum += 1

                Select Case fieldNum
                    Case 1
                        rowEDT855O3 = clsASCBASE1.dst.Tables("EDT855O3").NewRow
                        rowEDT855O3.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        rowEDT855O3.Item("EDI_OUTBOUND_DOC_NO") = EDI_Outbound_Doc_No
                        rowEDT855O3.Item("EDI_DTL_SEQ") = rowEDT850T2.Item("EDI_DTL_SEQ")
                        EDI_SDQ_SEQ += 1
                        rowEDT855O3.Item("EDI_SDQ_SEQ") = EDI_SDQ_SEQ
                        rowEDT855O3.Item("EDI_SDQ_UOM") = rowEDT850T2.Item("EDI_PO4_UOM")
                        'rowEDT855O3.Item("EDI_SDQ_QUAL") = String.Empty
                        clsASCBASE1.dst.Tables("EDT855O3").Rows.Add(rowEDT855O3)
                End Select

                rowEDT855O3.Item("EDI_STORE_" & fieldNum.ToString("00")) = EDI_STORE
                Select Case rowSOTORDR1.Item("ORDR_STATUS") & String.Empty
                    Case "O"
                        rowEDT855O3.Item("EDI_QTY_" & fieldNum.ToString("00")) = rowSOTORDR2.Item("ORDR_QTY_OPEN")
                    Case Else
                        rowEDT855O3.Item("EDI_QTY_" & fieldNum.ToString("00")) = rowSOTORDR2.Item("ORDR_QTY_PICK")
                End Select

                If fieldNum = 10 Then
                    fieldNum = 0
                End If
            Next
        Next

    End Sub

End Class
