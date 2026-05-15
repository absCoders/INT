Public Class SOCMAIN1

    Private ASCDATA1 As WHC.ASCDATA1
    Private ASCMAIN1 As WHC.ASCMAIN1
    Private CYP As String = String.Empty
    Private sql As String = String.Empty

    Public Sub New(ByVal inCYP As String)
        CYP = inCYP
    End Sub


    ''' <summary>
    ''' Validates High Collection Authorization
    ''' </summary>
    ''' <param name="tblSOTORDR2">Order Details</param>
    ''' <returns>Returns Error Message for Customer/Store/High Colection not Authorized</returns>
    ''' <remarks></remarks>
    Public Function ValidateAuthorizations(ByVal tblSOTORDR2 As DataTable) As String
        Dim salesOrderList As New List(Of String)
        Return ValidateAuthorizations(tblSOTORDR2, salesOrderList)
    End Function

    ''' <summary>
    ''' Validates High Collection Authorization
    ''' </summary>
    ''' <param name="tblSOTORDR2">Order Details</param>
    ''' <param name="salesOrderList">List to contain order numbers with Authorization problems</param>
    ''' <returns>Returns Error Message for Customer/Store/High Colection not Authorized</returns>
    ''' <remarks></remarks>
    Public Function ValidateAuthorizations(ByVal tblSOTORDR2 As DataTable, ByRef salesOrderList As List(Of String)) As String

        ValidateAuthorizations = String.Empty
        If salesOrderList Is Nothing Then
            salesOrderList = New List(Of String)
        End If

        Dim tblCust As DataTable = ASCDATA1.SelectDistinct(tblSOTORDR2, New String() {"CUST_CODE", "CUST_STORE_NO"})
        Dim custStore As String = String.Empty
        Dim custCodes As New List(Of String)

        For Each row As DataRow In tblCust.Select("")
            custStore &= ", ('" & row.Item("CUST_CODE") & "', '" & row.Item("CUST_STORE_NO") & "')"
            If Not custCodes.Contains(row.Item("CUST_CODE")) Then
                custCodes.Add(row.Item("CUST_CODE"))
            End If
        Next

        If custStore.Length = 0 Then
            ValidateAuthorizations = "Authorizations could not determine the Customer / Store No"
            Exit Function
        Else
            custStore = custStore.Substring(1).Trim
            custStore = "( " & custStore & " )"
        End If

        Dim tblItems As DataTable = ASCDATA1.SelectDistinct(tblSOTORDR2, New String() {"ITEM_CODE"})
        Dim items As String = String.Empty
        For Each row As DataRow In tblItems.Select("")
            items &= ", '" & row.Item("ITEM_CODE") & "'"
        Next
        If items.Length = 0 Then
            ValidateAuthorizations = "Authorizations could not determine the Items"
            Exit Function
        Else
            items = items.Substring(1).Trim
            items = "( " & items & " )"
        End If

        Dim nextPeriod As String = ASCMAIN1.Period_Calc(CYP, 1)

        Dim rowSOTPARM1 As DataRow = ASCDATA1.GetDataRow("Select * from SOTPARM1 where SO_PARM_KEY = 'Z'")
        Dim SO_PARM_AUTH_MOS_CLOSE As Integer = Val(rowSOTPARM1.Item("SO_PARM_AUTH_MOS_CLOSE") & "")
        Dim YPX As String = ASCMAIN1.Period_Calc(CYP, SO_PARM_AUTH_MOS_CLOSE)
        '& " AND SATAUTH1.OPS_YYYYPP_OPENED <= '" & ASCMAIN1.CYP & "'" _

        sql = " Select SATAUTH1.*,ICTITEM1.ITEM_CODE" _
            & " FROM SATAUTH1, ICTCOLL1, ICTITEM1" _
            & " WHERE SATAUTH1.HC_CODE = ICTCOLL1.HC_CODE" _
            & " AND ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
            & " AND (SATAUTH1.CUST_CODE, SATAUTH1.CUST_STORE_NO) in " & custStore _
            & " AND SATAUTH1.OPS_YYYYPP_OPENED <= '" & YPX & "'" _
            & " AND NVL(SATAUTH1.OPS_YYYYPP_CLOSED, '" & nextPeriod & "') > '" & CYP & "'" _
            & " AND ICTITEM1.ITEM_CODE IN " & items
        Dim tblSATAUTH1 As DataTable = ASCDATA1.GetDataTable(sql)

        sql = "SELECT DISTINCT ICTITEM1.ITEM_CODE, ICTCOLL1.HC_CODE" _
            & " FROM ICTITEM1, ICTCOLL1" _
            & " WHERE ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" _
            & " AND ICTITEM1.ITEM_CODE IN " & items
        Dim tblHC_CODES As DataTable = ASCDATA1.GetDataTable(sql)
        tblHC_CODES.PrimaryKey = New DataColumn() {tblHC_CODES.Columns("ITEM_CODE")}

        Dim errorList As New List(Of String)
        Dim errorMsg As String = String.Empty

        'Dim TRADE_CLASS_CODE As String = rowARTCUST1.Item("TRADE_CLASS_CODE")
        'Dim rowSOTTCLS1 As DataRow = Lookup("SOTTCLS1", TRADE_CLASS_CODE)

        sql = "Select ARTCUST1.CUST_CODE " _
            & " FROM ARTCUST1, SOTTCLS1" _
            & " WHERE ARTCUST1.TRADE_CLASS_CODE = SOTTCLS1.TRADE_CLASS_CODE" _
            & " AND AUTH_REQD = '1'" _
            & " AND CUST_CODE IN ('" & String.Join("', '", custCodes.ToArray) & "')"

        Dim tblARTCUST1 As DataTable = ASCDATA1.GetDataTable(sql)

        For Each rowARTCUST1 As DataRow In tblARTCUST1.Select("", "CUST_CODE")
            Dim CUST_CODEx As String = rowARTCUST1.Item("CUST_CODE")
            For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("CUST_CODE = '" & CUST_CODEx & "'", "CUST_STORE_NO,ITEM_CODE")
                Dim CUST_CODE As String = rowSOTORDR2.Item("CUST_CODE") & String.Empty
                Dim CUST_STORE_NO As String = rowSOTORDR2.Item("CUST_STORE_NO") & String.Empty
                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty

                sql = "CUST_CODE = '" & CUST_CODE & "' AND CUST_STORE_NO = '" & CUST_STORE_NO & "' AND ITEM_CODE = '" & ITEM_CODE & "'"
                If tblSATAUTH1.Select(sql, "").Length = 0 Then
                    Dim row As DataRow = tblHC_CODES.Rows.Find(ITEM_CODE)

                    If row IsNot Nothing Then
                        errorMsg = CUST_CODE & "/" & CUST_STORE_NO & " not Authorized for High Collection: " & row.Item("HC_CODE")
                    Else
                        errorMsg = CUST_CODE & "/" & CUST_STORE_NO & " not Authorized for Item: " & ITEM_CODE
                    End If

                    If Not errorList.Contains(errorMsg) Then
                        errorList.Add(errorMsg)
                    End If

                    If Not salesOrderList.Contains(rowSOTORDR2.Item("ORDR_NO")) Then
                        salesOrderList.Add(rowSOTORDR2.Item("ORDR_NO"))
                    End If
                End If
            Next
        Next

        If errorList.Count > 0 Then
            ValidateAuthorizations = String.Join(vbCrLf, errorList.ToArray)
        End If

    End Function

    Public Function ValidateOrderQtys(ByVal tblSOTORDR2 As DataTable, Optional FieldName As String = "ORDR_QTY") As String
        Dim salesOrderList As New List(Of String)
        Return ValidateOrderQtys(tblSOTORDR2, salesOrderList, FieldName)
    End Function

    Public Function ValidateOrderQtys(ByVal tblSOTORDR2 As DataTable, ByRef salesOrderList As List(Of String), Optional FieldName As String = "ORDR_QTY") As String

        ValidateOrderQtys = String.Empty

        Dim tblItems As DataTable = ASCDATA1.SelectDistinct(tblSOTORDR2, New String() {"ITEM_CODE"})
        Dim items As String = String.Empty
        For Each row As DataRow In tblItems.Select("")
            items &= ", '" & row.Item("ITEM_CODE") & "'"
        Next

        If items.Length > 0 Then
            items = items.Substring(1).Trim
            items = "( " & items & " )"
        End If

        sql = "Select * from ICTITEM1 WHERE ITEM_CODE IN  " & items
        Dim tblICTITEM1 As DataTable = ASCDATA1.GetDataTable(sql, "ICTITEM1")

        Dim errorList As New List(Of String)
        Dim errorMsg As String = String.Empty

        For Each rowSOTORDR2 As DataRow In tblSOTORDR2.Select("", "ITEM_CODE")
            Dim CUST_CODE As String = rowSOTORDR2.Item("CUST_CODE") & String.Empty
            If CUST_CODE.StartsWith("IPLB") Then
                ' LM SAYS TO AVOID QTY CHECKS FOR ALL CUSTOMERS BEGINNING WITH IPLB
                Continue For
            End If

            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & String.Empty
            Dim rowICTITEM1 As DataRow = tblICTITEM1.Rows.Find(ITEM_CODE)

            Dim ORDR_NO As String = rowSOTORDR2.Item("ORDR_NO") & String.Empty
            Dim ORDR_LNO As String = rowSOTORDR2.Item("ORDR_LNO")

            Dim ORDR_QTY As Int32 = Val(rowSOTORDR2.Item(FieldName) & String.Empty)
            If ORDR_QTY = 0 Then
                Continue For
            End If

            ' SO Min Qty
            Dim ITEM_SO_QTY_MIN As Int32 = Val(rowICTITEM1.Item("ITEM_SO_QTY_MIN") & String.Empty)
            ' SO Multiple
            Dim ITEM_SO_QTY_MULT As Int32 = Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & String.Empty)
            ' Inner Pack
            Dim ITEM_STD_PACK_SLS As Int32 = Val(rowICTITEM1.Item("ITEM_STD_PACK_SLS") & String.Empty)
            ' Allow Half Pack
            Dim ITEM_ALLOW_HALF_PACK As Boolean = (Val(rowICTITEM1.Item("ITEM_ALLOW_HALF_PACK") & String.Empty) = 1)

            ' Order Quantity meets Min Qty and Order Multiple restictions
            If ORDR_QTY >= ITEM_SO_QTY_MIN AndAlso ORDR_QTY Mod ITEM_SO_QTY_MULT = 0 Then
                Continue For
            End If

            ' If Half pack then only half pack, not a case and a half pack
            If ITEM_ALLOW_HALF_PACK AndAlso ORDR_QTY = ITEM_SO_QTY_MULT / 2 _
                    AndAlso ITEM_SO_QTY_MULT Mod 2 = 0 _
                    AndAlso ITEM_SO_QTY_MULT > 0 Then
                Continue For
            End If
            'If ITEM_ALLOW_HALF_PACK AndAlso ORDR_QTY = ITEM_STD_PACK_SLS / 2 _
            '        AndAlso ITEM_STD_PACK_SLS Mod 2 = 0 _
            '        AndAlso ITEM_STD_PACK_SLS > 0 Then
            '    Continue For
            'End If
            '

            errorMsg = "Order (" & ORDR_NO & "/" & ORDR_LNO & ") has an Invalid Order Qty for Item: " & ITEM_CODE _
                & ", Min Qty = " & ITEM_SO_QTY_MIN & ", Order Multiple = " & ITEM_SO_QTY_MULT

            If Not errorList.Contains(errorMsg) Then
                errorList.Add(errorMsg)
            End If

            If Not salesOrderList.Contains(ORDR_NO & "/" & ORDR_LNO) Then
                salesOrderList.Add(ORDR_NO & "/" & ORDR_LNO)
            End If
        Next

        If errorList.Count > 0 Then
            ValidateOrderQtys = String.Join(vbCrLf, errorList.ToArray)
        End If

    End Function

    Public Shared Function UPC( _
    ByRef frmASFBASE0 As WHC.ASCBASE0, _
    ByVal UPC_SEQUENCE_NO As String, _
    ByRef SO_PARM_UPC_VENDOR_ID As String, _
    Optional ByVal prefix_with_VENDOR_ID As Boolean = True) As String

        ' Note: Check Digit Calculation applies to the 19-digits prior to the check digit
        '       These 11 digits are made up from the 6 digit Vendor ID prepended to the 5 digit UPC Serial Number
        '       19 digits = '0000' + 6 digit SO_PARM_UPC_VENDOR_ID + 9 digit Carton Serial Number

        Dim Check_Digit_Seed As String

        If prefix_with_VENDOR_ID Then
            If SO_PARM_UPC_VENDOR_ID = "" Then
                SO_PARM_UPC_VENDOR_ID = frmASFBASE0.ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & ""
            End If

            If Len(UPC_SEQUENCE_NO) <> 5 Then
                If Len(UPC_SEQUENCE_NO) <> 9 Then
                    Stop
                End If
            End If

            Check_Digit_Seed = Mid(SO_PARM_UPC_VENDOR_ID, 1) & UPC_SEQUENCE_NO
        Else
            Check_Digit_Seed = UPC_SEQUENCE_NO
        End If

        Dim odd_digits As Integer
        Dim even_digits As Integer

        For i As Integer = 1 To Len(Check_Digit_Seed) Step 2
            odd_digits = odd_digits + Val(Mid(Check_Digit_Seed, 1, 1))
            Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            If Check_Digit_Seed <> "" Then
                even_digits = even_digits + Val(Mid(Check_Digit_Seed, 1, 1))
                Check_Digit_Seed = Mid(Check_Digit_Seed, 2)
            End If
        Next i

        Dim check_digit As Integer
        check_digit = (odd_digits * 3 + even_digits) Mod 10
        If check_digit <> 0 Then
            check_digit = 10 - check_digit
        End If

        If prefix_with_VENDOR_ID Then
            UPC = SO_PARM_UPC_VENDOR_ID & UPC_SEQUENCE_NO & Format(check_digit, "0")
        Else
            UPC = UPC_SEQUENCE_NO & Format(check_digit, "0")
        End If

    End Function


End Class
