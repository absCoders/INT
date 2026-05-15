Imports System.Drawing
Imports System.Math
Imports Infragistics.Olap.FlatData
Imports Infragistics.Win.UltraWinPivotGrid.DataSelector

Public Class SPFOLAP1


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SPFCOOPI" Then
            InquiryMode = True
        End If


        With dst


        End With

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        Absx1.cmbFor("RYP0").Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -11)
        Absx1.cmbFor("RYP1").Value = ASCMAIN1.CYP

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"


            Case "Done"


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")


                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode

        tab0.Visible = ScreenMode


        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        Dim coopSQL As String = "SELECT C1.CUST_NAME, SC1.AUTH_DATE, SC1.AUTH_REQ_BY, SR.SREP_NAME, SELL.SELL_NAME, " & _
                                "  SC1.APPR_STATUS_CODE, SC1.FEATURE_DESC, SC1.BOOKING_NAME, SC1.CUST_AGR_RECD, " & _
                                "  SC1.PROOF_ADV_RECD, SC1.SAMPLE_RECD, SEAS.SEASON_DESC, VH1.VEHICLE_DESC, NVL(SC2.VEHICLE_CPM,0) VEHICLE_CPM, " & _
                                "  SC2.DROP_DATE, SC2.IN_STORE_DATE, NVL(SC2.QTY, 0) QTY, NVL(SC2.OTHER_COST,0) OTHER_COST, NVL(SC2.OPEN_AMT,0) OPEN_AMT, " & _
                                "  NVL(SC2.PAID_AMT,0) PAID_AMT, SC2.STATUS_CODE, NVL(SC2.PYMTS,0) PYMTS, SC2.OPS_YYYYPP, VH2.AD_SIZE_DESC, " & _
                                "  CO.COLLECTION_NAME, NVL(SC3.DIST_AMT,0) DIST_AMT " & _
                                "FROM SPTCOOP1 SC1 " & _
                                "JOIN SPTCOOP2 SC2  ON (SC2.AUTH_NO = SC1.AUTH_NO) " & _
                                "JOIN SPTCOOP3 SC3  ON (SC3.AUTH_NO = SC2.AUTH_NO AND SC3.AUTH_LNO = SC2.AUTH_LNO) " & _
                                "JOIN ARTCUST1 C1   ON (SC1.CUST_CODE = C1.CUST_CODE) " & _
                                "JOIN ICTCOLL1 CO   ON (SC3.COLLECTION_CODE = CO.COLLECTION_CODE) " & _
                                "JOIN SOTSELL1 SELL ON (SC1.SELL_CODE = SELL.SELL_CODE) " & _
                                "JOIN ICTSEAS1 SEAS ON (SC1.SEASON_CODE = SEAS.SEASON_CODE) " & _
                                "JOIN SPTAVEH1 VH1  ON (SC2.VEHICLE_CODE = VH1.VEHICLE_CODE) " & _
                                "JOIN SPTAVEH2 VH2  ON (SC2.VEHICLE_CODE = VH2.VEHICLE_CODE AND SC2.AD_SIZE_LNO = " & _
                                "  VH2.AD_SIZE_LNO) " & _
                                "JOIN SOTSREP1 SR ON (SC1.SREP_CODE = SR.SREP_CODE) " & _
                                "WHERE SC2.OPS_YYYYPP BETWEEN :PARM1 AND :PARM2"

        Dim dtCoop As DataTable = ASCDATA1.GetDataTable(coopSQL, "coop", "VV", New String() {Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value})

        Dim retailSQL As String = "SELECT R2TY.OPS_YYYYPP, C1.CUST_NAME, C2.CUST_STORE_NAME, COLL.COLLECTION_NAME, SR.SREP_NAME, " & _
                                "  SELL.SELL_NAME, SD.SALES_DIVISION_NAME, TC.TRADE_CLASS_DESC, BR.BRAND_NAME, " & _
                                "  R2TY.RETAIL_SALES, NVL(R2LY.RETAIL_SALES, 0) RETAIL_SALES_LY " & _
                                "FROM RSTRETL2 R2TY " & _
                                "LEFT OUTER JOIN RSTRETL2 R2LY ON (R2TY.CUST_CODE = R2LY.CUST_CODE AND R2TY.CUST_STORE_NO = " & _
                                "  R2LY.CUST_STORE_NO AND R2TY.COLLECTION_CODE = R2LY.COLLECTION_CODE AND PERIOD_CALC( " & _
                                "  R2TY.OPS_YYYYPP, - 12) = R2LY.OPS_YYYYPP) " & _
                                "JOIN ARTCUST1 C1   ON (R2TY.CUST_CODE = C1.CUST_CODE) " & _
                                "JOIN ARTCUST2 C2   ON (R2TY.CUST_CODE = C2.CUST_CODE AND R2TY.CUST_STORE_NO = C2.CUST_STORE_NO) " & _
                                "JOIN SOTSREP1 SR   ON (C1.SREP_CODE = SR.SREP_CODE) " & _
                                "JOIN SOTTCLS1 TC   ON (C1.TRADE_CLASS_CODE = TC.TRADE_CLASS_CODE) " & _
                                "JOIN SOTSDIV1 SD   ON (SR.SALES_DIVISION_CODE = SD.SALES_DIVISION_CODE) " & _
                                "JOIN SOTSELL1 SELL ON (C2.SELL_CODE = SELL.SELL_CODE) " & _
                                "JOIN ICTCOLL1 COLL ON (R2TY.COLLECTION_CODE = COLL.COLLECTION_CODE) " & _
                                "JOIN ICTBRAN1 BR   ON (SR.SALES_DIVISION_CODE = BR.SALES_DIVISION_CODE) " & _
                                "WHERE R2TY.OPS_YYYYPP BETWEEN :PARM1 AND :PARM2 " & _
                                "ORDER BY OPS_YYYYPP"

        Dim dtRetail As DataTable = ASCDATA1.GetDataTable(retailSQL, "retail", "VV", New String() {Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value})


        Dim dsListCoop As New List(Of CoOpData)
        For Each row As DataRow In dtCoop.Rows()
            dsListCoop.Add(New CoOpData() With {.Customer = row("CUST_NAME"), .AuthDate = If(row("AUTH_DATE") Is DBNull.Value, Nothing, row("AUTH_DATE")), .AuthReq = row("AUTH_REQ_BY"),
                                            .SalesRep = row("SREP_NAME"), .SellName = row("SELL_NAME"), .ApprStatus = row("APPR_STATUS_CODE"),
                                            .Feature = row("FEATURE_DESC"), .BookingName = row("BOOKING_NAME"), .AgreementReceived = row("CUST_AGR_RECD"),
                                            .ProofReceived = row("PROOF_ADV_RECD"), .SampleReceived = row("SAMPLE_RECD"), .Season = row("SEASON_DESC"),
                                            .Vehicle = row("VEHICLE_DESC"), .VehicleCPM = row("VEHICLE_CPM"),
                                            .Quantity = row("QTY"), .OtherCost = row("OTHER_COST"),
                                            .OpenAmt = row("OPEN_AMT"), .PaidAmt = row("PAID_AMT"), .Status = row("STATUS_CODE"),
                                            .Period = row("OPS_YYYYPP"), .AdSize = row("AD_SIZE_DESC"), .Collection = row("COLLECTION_NAME"), .DistAmt = row("DIST_AMT")})
        Next

        Dim dsListRetail As New List(Of RetailSales)
        For Each row As DataRow In dtRetail.Rows()
            dsListRetail.Add(New RetailSales() With {.Customer = row("CUST_NAME"), .Period = row("OPS_YYYYPP"), .Store = row("CUST_STORE_NAME"),
                                                    .Collection = row("COLLECTION_NAME"), .SellInRep = row("SREP_NAME"), .SellThruRep = row("SELL_NAME"),
                                                     .TradeClass = row("TRADE_CLASS_DESC"), .Brand = row("BRAND_NAME"), .Division = row("SALES_DIVISION_NAME"),
                                                     .RetailSales = row("RETAIL_SALES"), .RetailSalesLY = row("RETAIL_SALES_LY")})
        Next


        If dsListCoop.Count > 0 Then
            BindPivot(dsListCoop, UltraPivotGrid1, OlapDataSelector1)
        End If

        If dsListRetail.Count > 0 Then
            Dim settings As FlatDataSourceInitialSettings = New FlatDataSourceInitialSettings()
            settings.Rows = "[RetailSales].[Customer]"
            settings.Measures = "[Measures].[RetailSales], [Measures].[RetailSalesLY]"
            BindPivot(dsListRetail, upgRetailSales, odsRetailSales, settings)
        End If

        ASCMAIN1.Progress("")
    End Sub

    Private Sub BindPivot(dataToBind As IEnumerable, gridToBind As UltraWinPivotGrid.UltraPivotGrid, selectorToBind As OlapDataSelector, Optional settings As FlatDataSourceInitialSettings = Nothing)

        Dim ds As FlatDataSource
        If settings Is Nothing Then
            ds = New FlatDataSource(dataToBind, Nothing)
        Else
            ds = New FlatDataSource(dataToBind, Nothing, settings)
        End If


        Dim parameters = New CubeGenerationParameters()
        ds.GenerateCube(parameters)
        ds.InitializeAsync(gridToBind)
        ' Bind the PivotGrid and DataSelector to data.

        gridToBind.DataSource = ds
        selectorToBind.DataSource = ds
        ' Optinally, set compact mode on PivotGrid
        gridToBind.RowHeaderLayout = Infragistics.Win.UltraWinPivotGrid.RowHeaderLayout.Compact


    End Sub


    Private Sub UltraPivotGrid1_InitializeColumn(sender As Object, e As UltraWinPivotGrid.InitializeColumnEventArgs) Handles UltraPivotGrid1.InitializeColumn, upgRetailSales.InitializeColumn
        e.Column.CellAppearance.Normal.TextHAlign = HAlign.Right
    End Sub

    Private Sub SplitContainer1_Panel1_Paint(sender As Object, e As PaintEventArgs) Handles SplitContainer1.Panel1.Paint

    End Sub
End Class

Public Class CoOpData
    Public Property Customer As String
    Public Property AuthDate As Date
    Public Property AuthReq As String
    Public Property SalesRep As String
    Public Property SellName As String
    Public Property ApprStatus As String
    Public Property Feature As String
    Public Property BookingName As String
    Public Property AgreementReceived As Boolean
    Public Property ProofReceived As Boolean
    Public Property SampleReceived As Boolean
    Public Property Season As String
    Public Property Vehicle As String
    Public Property VehicleCPM As Integer
    'Public Property DropDate As Date?
    'Public Property InStoreDate As Date?
    Public Property Quantity As Integer
    Public Property OtherCost As Decimal
    Public Property OpenAmt As Decimal
    Public Property PaidAmt As Decimal
    Public Property Status As String
    Public Property Period As String
    Public Property AdSize As String
    Public Property Collection As String
    Public Property DistAmt As Decimal
End Class

Public Class RetailSales
    Property Period As String
    Property Customer As String
    Property Store As String
    Property Collection As String
    Property SellInRep As String
    Property SellThruRep As String
    Property Brand As String
    Property Division As String
    Property TradeClass As String
    Property RetailSales As Decimal
    Property RetailSalesLY As Decimal
End Class