Public Class SORCONF1
    Dim SOTCONF1 As String = ""
    Dim SOTCONF5 As String = ""
    Dim CUST_CODE As String = ""
    Dim SOTORDR0 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = SQL_in("ORDR_GROUP_NO", "SOTORDR0.ORDR_GROUP_NO")
        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTCONF3")
    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Temp_Data(sqlw)

        With dst
            ASCMAIN1.sql = "Select * from " & SOTORDR0
            Create_TDA(dst.Tables.Add, "SOTORDR0", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select Distinct CONFIG_NO, ITEM_CODE, QTY from " & SOTCONF1
            Create_TDA(dst.Tables.Add, "SOTCONF3", "**", 0, False, , 2)

            ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC from ICTITEM1" _
                & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & SOTCONF1 & ")"
            Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select * from " & SOTCONF5
            Create_TDA(dst.Tables.Add, "SOTCONF5", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select Distinct CONFIG_NO" & vbCrLf _
                & " from " & SOTCONF5 & vbCrLf
            Create_TDA(dst.Tables.Add, "SOTCONF4", "**", 0, False, , 1)
            dst.Tables("SOTCONF4").Columns.Add("STORES")
            dst.Tables("SOTCONF4").Columns.Add("STORE_COUNT", GetType(System.Int64))
        End With

        If perform_fill Then
            Fill_Records_RPT("")
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)
        If sqlw <> "" Then
            Create_Temp_Data(sqlw)
        End If
        EnforceConstraints(False)
        Fill_Records("SOTORDR0")
        Fill_Records("SOTCONF3")
        Fill_Records("SOTCONF4")
        Fill_Records("SOTCONF5")
        Fill_Records("ICTITEM1")

        EnforceConstraints(True)
 
        For Each rowSOTCONF5 As DataRow In dst.Tables("SOTCONF5").Select("", "CUST_STORE_NO")
            Dim rowSOTCONF4 As DataRow = dst.Tables("SOTCONF4").Rows.Find _
                                         (New String() {rowSOTCONF5.Item("CONFIG_NO")})
            rowSOTCONF4.Item("STORES") = rowSOTCONF4.Item("STORES") & rowSOTCONF5.Item("CUST_STORE_NO") & ", "
            rowSOTCONF4.Item("STORE_COUNT") = Val(rowSOTCONF4.Item("STORE_COUNT") & "") + 1
        Next
    End Sub

    Sub Create_Temp_Data(SQLW As String)

        ASCMAIN1.sql = "Select SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN, SUM (ORDR_QTY_PICK) ORDR_QTY_PICK, SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" _
            & " from SOTORDR0 " & ASCMAIN1.SQL_Add_WHERE(SQLW)
        Dim row As DataRow = ASCDATA1.GetDataRow
        Dim C As String = ""
        If Val(row.Item("ORDR_QTY_SHIP") & "") <> 0 Then
            C = "ORDR_QTY_SHIP"
        ElseIf Val(row.Item("ORDR_QTY_PICK") & "") <> 0 Then
            C = "ORDR_QTY_PICK"
        Else
            C = "ORDR_QTY_OPEN"
        End If

        ASCMAIN1.sql = "Select * from SOTORDR0" & ASCMAIN1.SQL_Add_WHERE(SQLW)
        If SOTORDR0 = "" Then
            SOTORDR0 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTORDR0)
            ASCDATA1.ExecuteSQL("Insert into " & SOTORDR0 & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select SOTORDR2.CUST_STORE_NO, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", SUM (" & C & ") QTY" & vbCrLf _
            & " from SOTORDR1,SOTORDR2" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO " & vbCrLf _
            & "   and SOTORDR1.ORDR_GROUP_NO in (Select ORDR_GROUP_NO from " & SOTORDR0 & ")" & vbCrLf _
            & " group by SOTORDR2.CUST_STORE_NO, SOTORDR2.ITEM_CODE"
        If SOTCONF1 = "" Then
            SOTCONF1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTCONF1 & " Add CONFIG_NO VARCHAR2(6)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCONF1)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCONF1 & " (CUST_STORE_NO, ITEM_CODE, QTY) " & ASCMAIN1.sql)
        End If

        Dim CONFIG_NO As Integer = 0
        Dim CONFIG_NOs As New Dictionary(Of Integer, String)
        ASCMAIN1.sql = "Select Distinct CUST_STORE_NO from " & SOTCONF1
        For Each rowCUST_STORE_NO As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim CONFIG_NO_for_this_store As Integer = 0
            Dim CUST_STORE_NO As String = rowCUST_STORE_NO.Item("CUST_STORE_NO")
            ' If CUST_STORE_NO = "007019" Then Stop
            If CONFIG_NO > 0 Then
                For I As Integer = 1 To CONFIG_NO
                    ASCMAIN1.sql = "" _
                        & "(" _
                        & "Select ITEM_CODE, QTY from " & SOTCONF1 _
                        & " where CUST_STORE_NO = '" & CUST_STORE_NO & "'" _
                        & " minus " _
                        & "Select ITEM_CODE, QTY from " & SOTCONF1 _
                        & " where CUST_STORE_NO = '" & CONFIG_NOs(I) & "'" _
                        & ") union (" _
                        & "Select ITEM_CODE, QTY from " & SOTCONF1 _
                        & " where CUST_STORE_NO = '" & CONFIG_NOs(I) & "'" _
                        & " minus " _
                        & "Select ITEM_CODE, QTY from " & SOTCONF1 _
                        & " where CUST_STORE_NO = '" & CUST_STORE_NO & "'" _
                        & ")"
                    ASCMAIN1.sql = "Select Count (*) from (" & ASCMAIN1.sql & ")"
                    Dim RECORDS_DIFFERENCE As Integer = Val(ASCDATA1.GetDataValue & "")
                    If RECORDS_DIFFERENCE = 0 Then
                        CONFIG_NO_for_this_store = I
                        Exit For
                    End If
                Next
            End If
            If CONFIG_NO_for_this_store = 0 Then
                CONFIG_NO += 1
                CONFIG_NO_for_this_store = CONFIG_NO
                CONFIG_NOs.Add(CONFIG_NO, CUST_STORE_NO)
            End If
            ASCMAIN1.sql = "Update " & SOTCONF1 _
                & " Set CONFIG_NO = '" & Format(CONFIG_NO_for_this_store, "000000") & "'" _
                & " where CUST_STORE_NO = '" & CUST_STORE_NO & "'"
            ASCDATA1.ExecuteSQL()
        Next

        ASCMAIN1.sql = "Select Distinct CUST_STORE_NO, CONFIG_NO" & vbCrLf _
            & " from " & SOTCONF1
        If SOTCONF5 = "" Then
            SOTCONF5 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCONF5)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCONF5 & " (CUST_STORE_NO,CONFIG_NO) " & ASCMAIN1.sql)
        End If

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            ASCMAIN1.sql = "Select Distinct CUST_CODE from SOTORDR0" & ASCMAIN1.SQL_Add_WHERE(SQL_in("ORDR_GROUP_NO", "SOTORDR0.ORDR_GROUP_NO"))
            Dim rows() As DataRow = ASCDATA1.GetDataTable.Select
            If rows.Length = 0 Then
                EMsg &= vbCr & "No Records Found"
            ElseIf rows.Length > 1 Then
                EMsg &= vbCr & "Multiple Customers found in Groups Selected"
            Else
                CUST_CODE = rows(0).Item("CUST_CODE")
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "ORDR_GROUP_NO"
                sqlw = " SOTORDR0.ORDR_DATE > SYSDATE -180"
        End Select
        Return sqlw
    End Function
End Class