Public Class SORCART1

#Region "General Declarations"

    Dim sqlReport As String = String.Empty

#End Region

    Dim sqlCART3 As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""

        sqlw &= SQL_in("ORDR_GROUP_NO", "SOTSHIP1.ORDR_GROUP_NO")
        sqlw &= SQL_in("SHIP_BOL_NO", "SOTSHIP1.SHIP_BOL_NO")
        sqlw &= SQL_in("PICK_NO", "SOTPICK1.PICK_NO")
        sqlw &= SQL_in("ORDR_NO", "SOTPICK1.ORDR_NO")
        sqlw &= SQL_in("LOT_NO", "SOTCART3.LOT_NO")

        sqlCART3 = SQL_in("LOT_NO", "SOTCART3.LOT_NO")

        sqlReport = sqlw

        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTCART1")
    End Sub

    Public Overrides Sub Print_Report()
        RPT = "SORCART1"
        CR_params.Add("SUBT", "")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If sqlReport.Length = 0 Then
                Dim sqlw As String = String.Empty

                sqlw &= SQL_in("ORDR_GROUP_NO", "SOTSHIP1.ORDR_GROUP_NO")
                sqlw &= SQL_in("SHIP_BOL_NO", "SOTSHIP1.SHIP_BOL_NO")
                sqlw &= SQL_in("ORDR_NO", "SOTPICK1.ORDR_NO")
                sqlw &= SQL_in("PICK_NO", "SOTPICK1.PICK_NO")
                sqlw &= SQL_in("LOT_NO", "SOTCART3.LOT_NO")
                sqlCART3 = SQL_in("LOT_NO", "SOTCART3.LOT_NO")

                If sqlw.Length = 0 Then
                    EMsg &= vbCr & "You must make a filter selection."
                End If
            End If

        End If
    End Sub

    Overrides Sub Update_Record()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst
            Create_TDA(.Tables.Add, "ARTCUST1", "*")
            Create_TDA(.Tables.Add, "SOTSHIP1", "*")
            Create_TDA(.Tables.Add, "SOTPICK1", "*")
            Create_TDA(.Tables.Add, "SOTCART1", "*")
            Create_TDA(.Tables.Add, "SOTCART2", "*")
            Create_TDA(.Tables.Add, "SOTCART3", "*")

            .Relations.Add(.Tables("SOTSHIP1").Columns("SHIP_BOL_NO"), .Tables("SOTPICK1").Columns("SHIP_BOL_NO"))
            .Relations.Add(.Tables("SOTPICK1").Columns("PICK_NO"), .Tables("SOTCART1").Columns("PICK_NO"))
            .Relations.Add(.Tables("SOTCART1").Columns("CART_NO"), .Tables("SOTCART2").Columns("CART_NO"))
            .Relations.Add({ .Tables("SOTCART2").Columns("CART_NO"), .Tables("SOTCART2").Columns("CART_LNO")}, { .Tables("SOTCART3").Columns("CART_NO"), .Tables("SOTCART3").Columns("CART_LNO")})

            With .Tables("SOTCART2")
                .Columns.Add("NUM_LOTS", GetType(Int16), "COUNT(CHILD.LOT_NO)")
            End With

            .Tables("SOTSHIP1").Columns.Add("CUST_CODE", GetType(System.String))
            .Tables("SOTSHIP1").Columns.Add("ORDR_CUST_PO", GetType(System.String))
            .Tables("SOTSHIP1").Columns.Add("TOTAL_CARTONS", GetType(System.Int32), "SUM(CHILD.PICK_CNT_CARTONS)")
            .Tables("SOTSHIP1").Columns.Add("TOTAL_WEIGHT", GetType(System.Decimal), "SUM(CHILD.PICK_TOTAL_WGT)")

            .Tables("SOTPICK1").Columns.Add("CUST_CODE", GetType(System.String))
            .Tables("SOTPICK1").Columns.Add("CUST_STORE_NO", GetType(System.String))

        End With

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = String.Empty

        If parms.Length > 0 Then
            sqlw = parms(0)
        Else
            sqlw = sqlReport
        End If

        If sqlw.Length = 0 Then
            Exit Sub
        End If

        Dim sql As String = $"Select DISTINCT SOTORDR0.CUST_CODE, SOTPICK1.PICK_NO, SOTPICK1.SHIP_BOL_NO, SOTCART1.CART_NO
                                FROM SOTORDR0, SOTSHIP1, SOTPICK1, SOTCART1, SOTCART2, SOTCART3
                                WHERE SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO
                                AND SOTSHIP1.SHIP_BOL_NO = SOTPICK1.SHIP_BOL_NO
                                AND SOTPICK1.PICK_NO = SOTCART1.PICK_NO (+)
                                AND SOTCART1.CART_NO = SOTCART2.CART_NO (+)
                                AND SOTCART2.CART_NO = SOTCART3.CART_NO (+) AND SOTCART2.CART_LNO = SOTCART3.CART_LNO (+)
                                {sqlw}"

        Dim wktable As String = ASCMAIN1.Temp_Table(sql)

        EnforceConstraints(False)

        sql = "SELECT * FROM ARTCUST1 WHERE CUST_CODE IN (SELECT DISTINCT CUST_CODE FROM " & wktable & ")"
        Fill_Records("ARTCUST1", String.Empty, True, sql)

        sql = "SELECT SOTSHIP1.*, SOTORDR0.CUST_CODE, SOTORDR0.ORDR_CUST_PO FROM SOTSHIP1, SOTORDR0 WHERE SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO AND SOTSHIP1.SHIP_BOL_NO IN (SELECT DISTINCT SHIP_BOL_NO FROM " & wktable & ")"
        Fill_Records("SOTSHIP1", String.Empty, True, sql)

        sql = "SELECT SOTPICK1.*, SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO FROM SOTPICK1, SOTORDR1 WHERE SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO AND SOTPICK1.PICK_NO IN (SELECT DISTINCT PICK_NO FROM " & wktable & ")"
        Fill_Records("SOTPICK1", String.Empty, True, sql)

        sql = $"SELECT * FROM SOTCART1 WHERE CART_NO IN (SELECT CART_NO FROM {wktable})"
        Fill_Records("SOTCART1", String.Empty, True, sql)

        sql = $"SELECT * FROM SOTCART2 WHERE CART_NO IN (SELECT CART_NO FROM {wktable})"
        Fill_Records("SOTCART2", String.Empty, True, sql)

        sql = $"SELECT * FROM SOTCART3 WHERE CART_NO IN (SELECT CART_NO FROM {wktable}) " & sqlCART3
        Fill_Records("SOTCART3", String.Empty, True, sql)


        EnforceConstraints(True)

    End Sub

End Class