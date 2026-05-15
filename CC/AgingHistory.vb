Public Class AgingHistory

    Private sCustomerCode As String = String.Empty
    Private sPeriod As String = String.Empty
    Private AbsCon As Object = Nothing ' New ABSConnector
    'Private AbsCon As ABSConnector = New ABSConnector

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        sCustomerCode = String.Empty
        sPeriod = String.Empty

    End Sub

#Region "Control Properties"

    ''' <summary>
    ''' Customer Code to get CC data for
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CustomerCode() As String
        Get
            Return sCustomerCode
        End Get
        Set(ByVal value As String)
            sCustomerCode = value
        End Set
    End Property

    ''' <summary>
    ''' Gives access to the User Controls Grid Control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property UserControlGrid() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return grdAgingHistoryControl
        End Get

    End Property

    ''' <summary>
    ''' Get / Sets the period to use for the data
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Period() As String
        Get
            Return sPeriod
        End Get
        Set(ByVal value As String)
            sPeriod = value
        End Set
    End Property

#End Region

#Region "Control Public Methods"

    Public Sub ClearData()
        With AbsCon.dst
            .Tables("ARTSTMT1").Rows.Clear()
            .Tables("ARTOPEN1").Rows.Clear()
        End With

    End Sub

    Public Sub DisplayData()
        DisplayData(sCustomerCode, sPeriod)
    End Sub

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DisplayData(ByVal CustomerCode As String, ByVal Period As String)

        Dim sql As String = String.Empty
        Dim fuelSurcharge As String = String.Empty
        Dim i As Integer = 0

        Dim yp As New List(Of String)
        Dim dt As New Dictionary(Of String, Date)
        Dim AGED_TOTALS() As Double
        Dim DAYS(4) As Integer
        Dim YP_3(,) As String

        Try
            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            If Period.Length = 0 Then
                Period = DateTime.Now.ToString("yyyyMM")
            End If

            ReDim YP_3(3, 1)
            sql = "Select OPS_YYYYPP, PRD_END_DATE " _
                & " from GLTPARM2 " _
                & " where OPS_YYYYPP <= :PARM1" _
                & "   and OPS_YYYYPP >= :PARM2 order by OPS_YYYYPP DESC"

            For Each row As DataRow In AbsCon.GetDataTable(sql, "tbl", "VV", New Object() {Period, ABSolution.ASCMAIN1.Period_Calc(Period, -3)}).Rows
                dt.Add(row.Item(0), row.Item(1))
                yp.Add(row.Item(0))
                YP_3(i, 0) = row.Item(0)
                YP_3(i, 1) = ABSolution.ASCMAIN1.Get_Legend(row.Item(0))
                i += 1
            Next

            Dim rowARTPARM1 As DataRow = AbsCon.GetDataRow("Select * from ARTPARM1")
            For i = 1 To 4
                DAYS(i) = Val(rowARTPARM1.Item("AR_PARM_AGE_CATG_" & CStr(i)) & "")
                grdAgingHistoryControl.DisplayLayout.Bands(0).Columns("AGE_" & i).Header.Caption = rowARTPARM1.Item("AR_PARM_AGE_CATG_DESC_" & CStr(i)) & String.Empty
            Next

            With AbsCon.dst
                If Not AbsCon.dst.Tables.Contains("ARTSTMT1") Then
                    sql = "SELECT ARTSTMT1.*, GLTPARM2.LEGEND" _
                        & " from ARTSTMT1, GLTPARM2 where GLTPARM2.OPS_YYYYPP (+) = ARTSTMT1.OPS_YYYYPP and ARTSTMT1.CUST_CODE = :PARM1"
                    AbsCon.Create_TDA(.Tables.Add, "ARTSTMT1", sql, 0, False, "V", 2)

                    .Tables("ARTSTMT1").Columns.Add("AGE_OPEN", GetType(System.Int64), "IIF(ISNULL(TOTAL_DUE,0)=0,0,ISNULL(DAY_DOLLARS,0)/ISNULL(TOTAL_DUE,0))")
                    .Tables("ARTSTMT1").Columns.Add("DAYS_PPD", GetType(System.Int64), "IIF(ISNULL(TOTAL_CLOSED,0)=0,0,ISNULL(DAY_DOLLARS_CLOSED,0)/ISNULL(TOTAL_CLOSED,0))")
                    If Not .Tables("ARTSTMT1").Columns.Contains("AGE_0") Then
                        .Tables("ARTSTMT1").Columns.Add("AGE_0", GetType(System.Decimal))
                    End If
                    .Tables("ARTSTMT1").Columns.Add("TOTAL_W_FUTURES", GetType(System.Decimal), "ISNULL(TOTAL_DUE,0) + ISNULL(AGE_0,0)")
                    grdAgingHistoryControl.DataSource = AbsCon.dst.Tables("ARTSTMT1")

                    grdAgingHistoryControl.DisplayLayout.UseFixedHeaders = True
                    With grdAgingHistoryControl.DisplayLayout.Bands(0)
                        .Columns("LEGEND").Header.Fixed = True
                    End With
                End If

                If Not AbsCon.dst.Tables.Contains("ARTOPEN1") Then

                    Dim tbl As DataTable = AbsCon.GetDataTable("Select * from ARTOPEN1 WHERE ROWNUM < 1")
                    If tbl.Columns.Contains("FUEL_SURCHARGE_INV") Then
                        fuelSurcharge = "ARTOPEN1.FUEL_SURCHARGE_INV"
                    Else
                        fuelSurcharge = "0 FUEL_SURCHARGE_INV"
                    End If

                    sql = "SELECT ARTOPEN1.CUST_CODE, ARTOPEN1.INV_TYPE, ARTOPEN1.INV_NUM, ARTOPEN1.INV_DATE" _
                        & ", ARTOPEN1.CUST_SHIP_TO_NO, ARTOPEN1.POST_CODE, ARTOPEN1.TERM_CODE, ARTOPEN1.INV_DUE_DATE, ARTOPEN1.ORDR_NO" _
                        & ", ARTOPEN1.INV_CUST_PO, ARTOPEN1.INV_FREIGHT, " & fuelSurcharge & ", ARTOPEN1.INV_TOTAL_AMOUNT, ARTOPEN1.INV_BALANCE, ARTOPEN1.ORDR_NO_WEB" _
                        & ", ARTOPEN1.INV_PROFIT_B2C, ARTOPEN1.INV_PMT, ARTOPEN1.INV_DISC_TAKEN, ARTOPEN1.INV_WRITE_OFF, ARTOPEN1.DIVISION_CODE" _
                        & ", (CASE WHEN ARTOPEN1.INV_DATE > '" & Format(dt(yp(1)), "dd-MMM-yyyy") & "' AND ARTOPEN1.INV_DATE <= '" & Format(dt(yp(0)), "dd-MMM-yyyy") & "' THEN '1' ELSE" _
                        & "  CASE WHEN ARTOPEN1.INV_DATE > '" & Format(dt(yp(2)), "dd-MMM-yyyy") & "' AND ARTOPEN1.INV_DATE <= '" & Format(dt(yp(1)), "dd-MMM-yyyy") & "' THEN '2' ELSE" _
                        & "  CASE WHEN ARTOPEN1.INV_DATE > '" & Format(dt(yp(3)), "dd-MMM-yyyy") & "' AND ARTOPEN1.INV_DATE <= '" & Format(dt(yp(2)), "dd-MMM-yyyy") & "' THEN '3' ELSE" _
                        & "  '4' END END END) AGE_BUCKET" _
                        & ", ARTOPEN1.ORDR_TYPE_CODE" _
                        & " from ARTOPEN1, SOTINVH1" _
                        & " where ARTOPEN1.CUST_CODE = :PARM1" _
                        & " and SOTINVH1.INV_TYPE (+) = ARTOPEN1.INV_TYPE" _
                        & " and SOTINVH1.INV_NO (+) = ARTOPEN1.INV_NUM"

                    AbsCon.Create_TDA(.Tables.Add, "ARTOPEN1", sql, 0, False, "V", 0)

                End If
            End With

            grdAgingHistoryControl.DisplayLayout.Rows.FixedRows.Clear()

            With AbsCon.dst
                AbsCon.Fill_Records("ARTSTMT1", CustomerCode)
                AbsCon.Fill_Records("ARTOPEN1", CustomerCode)

                ReDim AGED_TOTALS(5)
                'AgeOpenItemsbyDate(CustomerCode, Period, AGED_TOTALS, DAYS)
                AGED_TOTALS = TAC.ARCMAIN1.GetCustomerAgeOpenItems(CustomerCode)

                Dim rowARTSTMT1 As DataRow = .Tables("ARTSTMT1").NewRow
                rowARTSTMT1.Item("CUST_CODE") = CustomerCode
                rowARTSTMT1.Item("OPS_YYYYPP") = "999999"
                rowARTSTMT1.Item("LEGEND") = "Open AR"

                rowARTSTMT1.Item("TYP_I_OPEN") = .Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'I'")
                rowARTSTMT1.Item("TYP_R_OPEN") = .Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'R'")
                rowARTSTMT1.Item("TYP_C_OPEN") = .Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'C'")
                rowARTSTMT1.Item("TYP_D_OPEN") = .Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'D'")
                rowARTSTMT1.Item("TYP_B_OPEN") = .Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'B'")
                rowARTSTMT1.Item("TYP_O_OPEN") = .Tables("ARTOPEN1").Compute("SUM(INV_BALANCE)", "INV_TYPE = 'O'")

                rowARTSTMT1.Item("AGE_1") = AGED_TOTALS(1)
                rowARTSTMT1.Item("AGE_2") = AGED_TOTALS(2)
                rowARTSTMT1.Item("AGE_3") = AGED_TOTALS(3)
                rowARTSTMT1.Item("AGE_4") = AGED_TOTALS(4)
                rowARTSTMT1.Item("TOTAL_DUE") = AGED_TOTALS(0)
                rowARTSTMT1.Item("AGE_0") = AGED_TOTALS(5)

                .Tables("ARTSTMT1").Rows.Add(rowARTSTMT1)
            End With

            AbsCon.Sort_grdColumns(grdAgingHistoryControl, "OPS_YYYYPP".ToLower)

            Dim rowARTCUST1 As DataRow = AbsCon.GetDataRow("SELECT * FROM ARTCUST1 WHERE CUST_CODE = :PARM1", "V", New Object() {CustomerCode})

            If rowARTCUST1 IsNot Nothing Then
                grdAgingHistoryControl.Text = "AR Aging History for " & CustomerCode & ":" & rowARTCUST1.Item("CUST_NAME")
            Else
                grdAgingHistoryControl.Text = "AR Aging History for " & CustomerCode
            End If

            If grdAgingHistoryControl.Rows.Count > 0 Then
                grdAgingHistoryControl.DisplayLayout.Rows.FixedRows.Add(grdAgingHistoryControl.Rows(0))
            End If

        Catch e As MissingMethodException
            MessageBox.Show("DisplayData: " & e.Message)

        Catch ex As Exception
            MessageBox.Show("Display Error: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Internal Procedures"

    Private Sub AgeOpenItemsbyDate(ByVal CUST_CODE As String, ByVal Period As String, ByRef AGED_TOTALS() As Double, ByVal days() As Integer)

        Dim AGE_CATGY As String = String.Empty
        Dim AGE_WHERE As String = String.Empty
        Dim Total As Double = 0
        Dim AGE_AMT As Double = 0

        Dim DT(5) As String
        Dim SQL As String = String.Empty

        Dim AGE_DATE_COLUMN As String = "INV_DATE"

        ReDim AGED_TOTALS(5)
        AGED_TOTALS(5) = TAC.ARCMAIN1.GetCustomerAgeOpenItems(CUST_CODE)(5)

        Try
            For iLoop As Integer = 1 To 7
                SQL = "Select * from GLTPARM2 where OPS_YYYYPP = :PARM1"
                Dim rowGLTPARM2 As DataRow = AbsCon.GetDataRow(SQL, "V", New Object() {ABSolution.ASCMAIN1.Period_Calc(Period, -1 * iLoop)})
                Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE") & String.Empty
                If iLoop < 5 Then
                    DT(iLoop) = "'" & Format(PRD_END_DATE, "MM/dd/yyyy") & "'"
                End If
                AGE_WHERE = String.Empty

                Select Case iLoop
                    Case 1
                        AGE_CATGY = "Current"
                        AGE_WHERE = AGE_DATE_COLUMN & " > " & DT(iLoop)

                    Case 4
                        AGE_CATGY = "Over 90"
                        AGE_WHERE = AGE_DATE_COLUMN & " <= " & DT(iLoop - 1)

                    Case 5
                        AGE_CATGY = "Total Due"
                        AGE_AMT = Total

                    Case 6
                        AGE_CATGY = "Future"
                        AGE_AMT = AGED_TOTALS(5)

                    Case 7
                        AGE_CATGY = "Total AR"
                        AGE_AMT = Total + AGED_TOTALS(5)

                    Case Else
                        AGE_CATGY = Format(days(iLoop - 1) + 1, "00") & "-" & Format(days(iLoop), "00")
                        AGE_WHERE = AGE_DATE_COLUMN & " > " & DT(iLoop) & " and " & AGE_DATE_COLUMN & " <= " & DT(iLoop - 1)
                End Select

                If iLoop < 5 Then
                    AGE_AMT = Val(AbsCon.dst.Tables("ARTOPEN1").Compute("SUM (INV_BALANCE)", AGE_WHERE) & "")

                    'if in past then minus out futures, if negative then make 0
                    If iLoop > 1 AndAlso iLoop <= 4 AndAlso AGED_TOTALS(5) > 0 Then
                        'AGE_AMT -= AGED_TOTALS(5)
                        'If AGE_AMT < 0 Then AGE_AMT = 0
                        AGE_AMT = Math.Max(AGE_AMT - AGED_TOTALS(5), 0)
                    End If

                    Total += AGE_AMT
                    AGED_TOTALS(iLoop) = AGE_AMT
                End If

            Next

            AGED_TOTALS(0) = Total + AGED_TOTALS(5)

        Catch e As MissingMethodException
            MessageBox.Show("AgeOpenItemsbyDate: " & e.Message)

        Catch ex As Exception
            MessageBox.Show("AgeOpenItemsbyDate: " & ex.Message)

        End Try

    End Sub

#End Region


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class
