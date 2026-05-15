Public Class SOCOREL1
    Inherits WHCRPTS1

    Public Overrides Property reportName As String

    Private PICK_BATCH_NO As String = ""    ' Pick Batch No; this Release will be defined by this control number

    Private SOTORDR0 As String
    Private SOTORDR1 As String
    Private SOTORDR2 As String

    Private SOTPICK1 As String
    Private SOTPICK2 As String
    Private SOTSHIP1 As String
    Private SOTCART1 As String
    Private SOTCART2 As String

    Private ARTCUST1 As String
    Private ARTCUST1_CG As String

    Private SOTALLOZ As String
    Private SOTALLOY As String
    Private SOTALLO1 As String
    Private SOTOREL1 As String
    Private SOTOREL2 As String
    Private POTORDRW As String
    Private DPTITMFX As String

    Private PICK_NO_seq As Int64 = 0        ' Temporary Pick Ticket Sequencer
    Private SHIP_BOL_NO_seq As Int64 = 0    ' Temporary Shipment Sequencer
    Private CART_NO_seq As Int32 = 0        ' Temporary Carton Sequencer

    Private OOBAL As Boolean = False
    Private SOCMAIN1 As WHC.SOCMAIN1
    Private EDC855O1 As WHC.EDC855O1

    Public Class Arguments
        Public AllocationOnly As Boolean = False
        Public CancelIfNoPick As Boolean = False
        Public Client As String = String.Empty
        Public ForcePick As Boolean = False
        Public FPDCancelFutureDays As Int32 = 0
        Public FPDShortHorDays As Int32 = 0
        Public IncludeWOQtyAsAvailableToShip As Boolean = False
        Public MoneyOnly As Boolean = False
        Public ReleaseCompleteOnly As Boolean = False
        Public ReleasePastCancel As Boolean = False
        Public ReleaseShipDate As Date = DateTime.Today
        Public ShowPastCancel As Boolean = False
        Public WarehouseCode As String = String.Empty
        Public WarehouseSelection As String = "A"
        Public ShipByDate As Date = DateTime.Today
        Public SQL_in As New Dictionary(Of String, String)
        Public tblASTDSQLA As New DataTable
        Public OrderSource As String = "*"
        Public XNO As String = "1234567"
    End Class

    Private InArguments As New Arguments

    ''' <summary>
    ''' Retruns the dataset genertaed by the class
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property reportDataset() As DataSet
        Get
            Return dst
        End Get
    End Property

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment, clientData As Dictionary(Of String, Object))
        MyBase.New(g)

        ' inTestMode = True
        Try

            MENU_ITEM_OBJECT = "SOROREL1"

            Get_PARM("SOTPARM1")
            Get_PARM("ICTPARM1")
            Get_PARM("POTPARM1")

            InArguments = New Arguments
            InArguments.WarehouseCode = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & String.Empty

            InArguments.FPDCancelFutureDays = Val(ROWs("SOTPARM1").Item("SO_PARM_CANCEL_FUTURE_DAYS") & String.Empty)
            InArguments.FPDShortHorDays = Val(ROWs("SOTPARM1").Item("SO_PARM_SHORT_HOR_DAYS") & "")

            Dim SO_PARM_RELEASE_DAYS_AHEAD As Int64 = Val(ROWs.Item("SOTPARM1").Item("SO_PARM_RELEASE_DAYS_AHEAD") & String.Empty)
            InArguments.ReleaseShipDate = Now.Date.AddDays(SO_PARM_RELEASE_DAYS_AHEAD)
            InArguments.AllocationOnly = False

            If clientData.ContainsKey("logsFolder") Then
                logsFolder = clientData("logsFolder")
            End If

            If inTestMode Then RecordLogEntry("InitializeReportSettings(clientData)")
            InitializeReportSettings(clientData)

            If inTestMode Then RecordLogEntry("LoadBaseReportData")
            LoadBaseReportData()

            If inTestMode Then RecordLogEntry("LoadReportData")
            LoadReportData()

            'Try
            '    If inTestMode Then RecordLogEntry("reportFilename = GenerateReport(clientData)")
            '    reportFilename = GenerateReport(clientData)

            'Catch ex As Exception
            '    clsErrorMessage.Add("Error generating report: " & ex.Message)
            'End Try
        Catch ex As Exception
            clsErrorMessage.Add("Error generating report: " & ex.Message)
        Finally
            ASCMAIN1.MultiTask_Release(, , 1)
            ASCMAIN1.MultiTask_Release()
        End Try

    End Sub

    Private Sub InitializeReportSettings(clientData As Dictionary(Of String, Object))

        If clientData IsNot Nothing AndAlso clientData.ContainsKey("ARGUMENTS") Then
            InArguments = clientData("ARGUMENTS")
        End If

        SOCMAIN1 = New WHC.SOCMAIN1(ASCMAIN1.CYP)
        EDC855O1 = New WHC.EDC855O1(InArguments.Client, ASCMAIN1.CYP)
        XNO = InArguments.XNO

    End Sub

    Public Sub LoadReportData()

        For Each TABLE_NAME As String In New String() {"SOTSREP1", "SOTMKTC1", "ICTBRAN1", "SOTSDIV1"}
            ASCMAIN1.sql = "Select " & TABLE_NAME & ".* from " & TABLE_NAME
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, TABLE_NAME, 1))
        Next

        ' Main Process

        PICK_NO_seq = 0
        SHIP_BOL_NO_seq = 0
        CART_NO_seq = 0

        If Order_Release() Then

            BeginTrans()
            Update_SOTORDRx()

            If Not InArguments.AllocationOnly Then
                Create_PICK_SHIP_CART()
                Update_Release() ' Update Pick Ticket, Shipment Control & Carton Tables
            End If
            CommitTrans()
        End If

        ASCMAIN1.MultiTask_Release(, , 1)
    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()
    End Sub

    Overrides Sub Update_Archive()
        MyBase.Update_Archive()
    End Sub

    Private Sub LoadBaseReportData()

    End Sub


#Region "Report Procedures"

    Private Function Order_Release() As Boolean

        OOBAL = Check_OOB_Styles()
        If OOBAL Then
            Return False
        End If

        ' Prepare Order Header & Detail Work Tables in Oracle

        ASCMAIN1.sql = "" _
            & "Select SOTORDR1.*, ARTCUST1.CUST_PRIORITY_CODE" & vbCrLf _
            & " from SOTORDR1, ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & " and SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
            & " and (SOTORDR1.ORDR_SHIP_DATE <= '" & Format(InArguments.ReleaseShipDate, "dd-MMM-yyyy") & "' or SOTORDR1.ORDR_PRIORITY = '1')"

        ' If we are not showing past cancel and not releasing past cancel, then get all orders past cancel out of the result set

        If Not InArguments.ShowPastCancel Then
            If Not InArguments.ReleasePastCancel Then
                ASCMAIN1.sql &= " and SOTORDR1.ORDR_CANCEL_DATE >= '" & Format(DATETIME_STAMP.AddDays(InArguments.FPDCancelFutureDays), "dd-MMM-yyyy") & "'"
            End If
        End If

        If InArguments.WarehouseCode <> "" Then
            ASCMAIN1.sql &= " and SOTORDR1.WHSE_CODE = '" & InArguments.WarehouseCode & "'"
        End If

        For Each key As String In InArguments.SQL_in.Keys
            Select Case key
                Case "SALES_DIVISION_CODE"
                    ASCMAIN1.sql &= InArguments.SQL_in("SALES_DIVISION_CODE")
                Case "CUST_CODE"
                    ASCMAIN1.sql &= InArguments.SQL_in("CUST_CODE")
                Case "TRADE_CLASS_CODE"
                    ASCMAIN1.sql &= InArguments.SQL_in("TRADE_CLASS_CODE")
                Case "ORDR_GROUP_NO"
                    ASCMAIN1.sql &= InArguments.SQL_in("ORDR_GROUP_NO")
                Case "ORDR_NO"
                    ASCMAIN1.sql &= InArguments.SQL_in("ORDR_NO")
            End Select
        Next

        If InArguments.OrderSource <> "*" Then
            ASCMAIN1.sql &= " and SOTORDR1.ORDR_SOURCE = '" & InArguments.OrderSource & "'"
        End If

        SOTORDR1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add ORDR_AMT_OPEN NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add ORDR_QTY_OPEN NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add ORDR_AMT_PICK NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add ORDR_QTY_PICK NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Dim sqlw As String = ""
        Dim sql_ITEM_CODE As String = SQLA("ITEM_CODE", InArguments.tblASTDSQLA)
        If sql_ITEM_CODE <> "" Then
            sqlw = "Select Distinct SOTORDR2.ORDR_NO" & vbCrLf _
                & " from SOTORDR2, " & SOTORDR1 & " SOTORDR1" & vbCrLf _
                & " where SOTORDR2.ORDR_STATUS = 'O'" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ITEM_CODE in (" & sql_ITEM_CODE & ")"
            If SQLA("ITEM_CODE", InArguments.tblASTDSQLA, "EXCLUDE") = "1" Then
                ASCMAIN1.sql = "Delete from " & SOTORDR1 & " where ORDR_NO in (" & sqlw & ")"
            Else
                ASCMAIN1.sql = "Delete from " & SOTORDR1 & " where ORDR_NO not in (" & sqlw & ")"
            End If
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Create Index I_" & SOTORDR1 & "_1 on " & SOTORDR1 & " (ORDR_GROUP_NO)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        If InArguments.MoneyOnly Then
            ASCMAIN1.sql = "Delete from " & SOTORDR1 & " where ORDR_GROUP_NO in (" & vbCrLf _
                & "Select ORDR_GROUP_NO from SOTORDR0 " & vbCrLf _
                & " where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")" & vbCrLf _
                & "   and ORDR_AMT < 1" & ")"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If

        ASCMAIN1.AnalyzeTable(SOTORDR1)

        ' attempt to lock all orders in queue, and exit if we are unsuccessful

        If Not InArguments.AllocationOnly Then
            ASCMAIN1.sql = "Select Distinct ORDR_GROUP_NO from " & SOTORDR1
            For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select()
                Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , False, 1) Then
                    ASCMAIN1.MultiTask_Release(, , 1)
                    clsErrorMessage.Add("Could Not Lock Order Group " & ORDR_GROUP_NO)
                    Return False
                End If
            Next
        End If

        ' Set up Order Details

        ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO,ARTCUST1.CUST_CODE) CUST_CODE_ALLO" _
            & " from SOTORDR2, " & SOTORDR1 & " SOTORDR1, ARTCUST1" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE"

        SOTORDR2 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add Primary Key (ORDR_NO, ORDR_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add PICK_QTY_CANC_REL NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add PICK_QTY_BACK_REL NUMBER (8,0)")

        ASCMAIN1.sql = "Select * from " & SOTORDR2
        Create_TDA(dst.Tables.Add("SOTORDR2"), SOTORDR2, "**", 0)
        Fill_Records("SOTORDR2")


        ' Set up Order Totals from Sum of Details, and populate Order Header

        ASCDATA1.ExecuteSQL("Update " & SOTORDR1 & " SOTORDR1 Set ORDR_QTY_OPEN = (Select Sum (NVL(ORDR_QTY_OPEN,0)) from " & SOTORDR2 & " where ORDR_NO = SOTORDR1.ORDR_NO)")
        ASCDATA1.ExecuteSQL("Update " & SOTORDR1 & " SOTORDR1 Set ORDR_AMT_OPEN = (Select Sum (NVL(ORDR_QTY_OPEN,0) * NVL(ORDR_UNIT_PRICE,0)) from " & SOTORDR2 & " where ORDR_NO = SOTORDR1.ORDR_NO)")

        ASCMAIN1.sql = "Select * from " & SOTORDR1
        Create_TDA(dst.Tables.Add("SOTORDR1"), SOTORDR1, "**", 0)
        Fill_Records("SOTORDR1")


        ' Set up Order Group Summary

        ASCMAIN1.sql = "Select SOTORDR0.*, ARTCUST1.CUST_PRIORITY_CODE" & vbCrLf _
            & " from SOTORDR0, ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
            & "   and ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")"
        SOTORDR0 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDERS_HELD NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDERS_RELEASED NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDERS_CANCELLED NUMBER (8,0)")
        ' ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDERS_CANCELLED NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add ORDR_REL_HOLD_CODES VARCHAR2(20)")

        ASCMAIN1.sql = "Select * from " & SOTORDR0
        Create_TDA(dst.Tables.Add("SOTORDR0"), SOTORDR0, "**", 0)
        Fill_Records("SOTORDR0")


        ' Item Parameters

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_ORDR_REL_CODE, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_NOT_ALLOCATED, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_WEIGHT" & vbCrLf _
            & ", ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ICTITEM1.CARTON_PACK_QTY" & vbCrLf _
            & ", ICTBRAN1.SALES_DIVISION_CODE, '0' ITEM_IS_ALLOCATED, ICTITEM1.ITEM_SO_QTY_MIN" & vbCrLf _
            & " from ICTITEM1,ICTCOLL1,ICTBRAN1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE in (Select Distinct ITEM_CODE from " & SOTORDR2 & ")"
        SOTOREL1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL1 & " Add Primary Key (ITEM_CODE)")

        Dim Z As String = Format(Now, "dd-MMM-yyyy")
        ASCMAIN1.sql = "Update " & SOTOREL1 _
            & " Set ITEM_IS_ALLOCATED = '1'" & vbCrLf _
            & " where ITEM_CODE in (Select Distinct ITEM_CODE from SOTALLO1" & vbCrLf _
            & " where ITEM_CODE in (Select ITEM_CODE from " & SOTOREL1 & ") and (DATE_START > '" & Z & "' or DATE_END > '" & Z & "'))"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from " & SOTOREL1
        Create_TDA(dst.Tables.Add("SOTOREL1"), SOTOREL1, "**", 0)
        Fill_Records("SOTOREL1")


        ' Set up Work Order Table

        ASCMAIN1.sql = "Select POTORDR2.ITEM_CODE, POTORDR2.WHSE_CODE" & vbCrLf _
            & ", SUM (POTORDR2.PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
            & " from POTORDR2, POTORDR1" & vbCrLf _
            & " where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_STATUS = 'O'" & vbCrLf _
            & "   and POTORDR1.PO_ORDER_TYPE = 'W'" & vbCrLf _
            & "   and POTORDR2.ITEM_CODE in (Select ITEM_CODE from " & SOTOREL1 & ")" & vbCrLf _
            & IIf(InArguments.IncludeWOQtyAsAvailableToShip, "", " and ROWNUM < 1") & vbCrLf _
            & " group by POTORDR2.ITEM_CODE, POTORDR2.WHSE_CODE"
        POTORDRW = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDRW & " Add Primary Key (ITEM_CODE, WHSE_CODE)")


        ' Item Status by Warehouse

        ASCMAIN1.sql = "Select ITEM_CODE, WHSE_CODE" & vbCrLf _
            & ", WHSE_QTY_ON_HAND QTY_ON_HAND, WHSE_QTY_OPEN QTY_OPEN, WHSE_QTY_PICK QTY_PICK, WHSE_QTY_ONPO QTY_ONPO" & vbCrLf _
            & " from ICTSTAT2 where (ITEM_CODE,WHSE_CODE)" & vbCrLf _
            & " in (Select Distinct ITEM_CODE,WHSE_CODE from " & SOTORDR2 & ")"
        SOTOREL2 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add Primary Key (ITEM_CODE, WHSE_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_WO NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_ATS NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_TO_PICK NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_TO_DE_COMMIT NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_UNRELEASED NUMBER (8,0)")

        If InArguments.IncludeWOQtyAsAvailableToShip Then
            ASCMAIN1.sql = "Update " & SOTOREL2 & " SOTOREL2" & vbCrLf _
                & " Set QTY_WO = (Select PO_QTY_OPN from " & POTORDRW & " POTORDRW" & vbCrLf _
                & " where ITEM_CODE = SOTOREL2.ITEM_CODE and WHSE_CODE = SOTOREL2.WHSE_CODE)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Update " & SOTOREL2 & " SOTOREL2" & vbCrLf _
            & " Set QTY_ATS = NVL(QTY_ON_HAND,0) - NVL(QTY_PICK,0) - NVL(QTY_WO,0)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from " & SOTOREL2
        Create_TDA(dst.Tables.Add("SOTOREL2"), SOTOREL2, "**", 0)
        Fill_Records("SOTOREL2")

        With dst.Tables.Add("SOTORELA")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("ORDR_QTY_PICK", GetType(System.Int64))
            .Columns.Add("ORDR_QTY_OPEN", GetType(System.Int64))
            .Columns.Add("ORDR_QTY_BACK", GetType(System.Int64))
            .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE")}
        End With


        ' Customer Master : Sold-To - and all related Bill-To and Credit Group Parents and Children

        Dim SQL_ARTCUST1 As String = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.CUST_ALLOW_BACKORDER, '0' CUST_CASE_PACK_ONLY" & vbCrLf _
            & ", ARTCUST1.CUST_PRIORITY_CODE, ARTCUST1.CUST_SALES_HOLD" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO,ARTCUST1.CUST_CODE) CUST_CODE_ALLO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_ALLOCATE_BY_STORE,'0') CUST_ALLOCATE_BY_STORE" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE) CUST_BILL_TO_CUST" & vbCrLf _
            & ", NVL(ARTCUST1_BT.CUST_CREDIT_GROUP_CUST,ARTCUST1_BT.CUST_CODE) CUST_CREDIT_GROUP_CUST" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE, SOTTCLS1.MARKET_CODE" & vbCrLf _
            & ", NVL(SOTMKTC1.RESTRICT_ORDER_RELEASE,'0') RESTRICT_ORDER_RELEASE" & vbCrLf _
            & " from ARTCUST1,SOTTCLS1,SOTMKTC1,ARTCUST1 ARTCUST1_BT" & vbCrLf _
            & " where SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = SOTTCLS1.MARKET_CODE" & vbCrLf _
            & "   and ARTCUST1_BT.CUST_CODE = NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE)" & vbCrLf

        'ASCMAIN1.sql = SQL_ARTCUST1 _
        '    & "   and NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE) in " & vbCrLf _
        '    & " (Select Distinct NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE) " & vbCrLf _
        '    & " from " & SOTORDR0 & " SOTORDR0, ARTCUST1 where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE)"

        ASCMAIN1.sql = "" _
            & SQL_ARTCUST1 _
            & "   and NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE) in " & vbCrLf _
            & " (Select Distinct NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE) " & vbCrLf _
            & " from " & SOTORDR0 & " SOTORDR0, ARTCUST1 where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE)" & vbCrLf _
            & " UNION " & vbCrLf _
            & SQL_ARTCUST1 _
            & "   and NVL(ARTCUST1.CUST_CREDIT_GROUP_CUST,NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE)) in " & vbCrLf _
            & " (Select Distinct NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE) " & vbCrLf _
            & " from " & SOTORDR0 & " SOTORDR0, ARTCUST1 where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE)"

        ARTCUST1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        Dim zz As String = "(" _
            & " Select Distinct CUST_BILL_TO_CUST CUST_CODE from " & ARTCUST1 & " union " & vbCrLf _
            & " Select Distinct ARTCUST1.CUST_CODE from ARTCUST1,ARTCUST1 ARTCUST1_BT where ARTCUST1_BT.CUST_CODE = ARTCUST1.CUST_BILL_TO_CUST and ARTCUST1_BT.CUST_CREDIT_GROUP_CUST in (Select CUST_CODE from " & ARTCUST1 & ") union " & vbCrLf _
            & " Select Distinct CUST_CREDIT_GROUP_CUST CUST_CODE from " & ARTCUST1 & vbCrLf _
            & ")" & vbCrLf

        ASCMAIN1.sql = "" _
            & "Select CUST_CODE from ARTCUST1" & vbCrLf _
            & " where CUST_CODE in " & zz & " or CUST_BILL_TO_CUST in " & zz & " or CUST_CREDIT_GROUP_CUST in " & zz _
            & " minus " & vbCrLf _
            & "Select CUST_CODE from " & ARTCUST1

        ASCMAIN1.sql = "Insert into " & ARTCUST1 & " " & vbCrLf _
            & SQL_ARTCUST1 _
            & "   and ARTCUST1.CUST_CODE in (" & vbCrLf _
            & ASCMAIN1.sql & vbCrLf _
            & ")"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Add Primary Key (CUST_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Add CUST_AMT_PICK NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Add CUST_AMT_OPEN NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Add CUST_AMT_PICK_NOW NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Add CUST_HOLDS_SALES VARCHAR2(20)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Add RELEASED VARCHAR2(1)")

        ASCDATA1.ExecuteSQL("Create Index I_" & ARTCUST1 & "_1 on " & ARTCUST1 & "(CUST_BILL_TO_CUST)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ARTCUST1 & "_2 on " & ARTCUST1 & "(CUST_CREDIT_GROUP_CUST)")

        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Modify CUST_BILL_TO_CUST Not Null")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1 & " Modify CUST_CREDIT_GROUP_CUST Not Null")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select SOTORDR0.CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR0.ORDR_AMT_OPEN) CUST_AMT_OPEN" & vbCrLf _
            & ", SUM (SOTORDR0.ORDR_AMT_PICK) CUST_AMT_PICK" & vbCrLf _
            & "   from SOTORDR0," & ARTCUST1 & " ARTCUST1" & vbCrLf _
            & "   where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
            & "   and (SOTORDR0.ORDR_CNT_OPEN <> 0 or SOTORDR0.ORDR_CNT_PICK <> 0)" & vbCrLf _
            & "   group by SOTORDR0.CUST_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ARTCUST1 & " ARTCUST1 Set" & vbCrLf _
            & "    CUST_AMT_OPEN = R1.CUST_AMT_OPEN" & vbCrLf _
            & "   ,CUST_AMT_PICK = R1.CUST_AMT_PICK" & vbCrLf _
            & "    where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Update " & ARTCUST1 & " Set CUST_HOLDS_SALES = NVL(CUST_HOLDS_SALES,'') || '1' where CUST_SALES_HOLD = '1'")

        ASCMAIN1.sql = "Select * from " & ARTCUST1
        Create_TDA(dst.Tables.Add("ARTCUST1"), ARTCUST1, "**", 0)
        Fill_Records("ARTCUST1")

        ASCMAIN1.sql = "Select * from TATTERM1"
        Create_TDA(dst.Tables.Add, "TATTERM1", "**", 0)
        Fill_Records("TATTERM1")

        ' Customer Master - Credit Group

        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.CUST_PD_GRACE_DAYS, ARTCUST1.CUST_PD_GRACE_PCT" & vbCrLf _
            & ", ARTCUST1.CUST_CREDIT_LIMIT, ARTCUST1.CUST_CREDIT_HOLD" & vbCrLf _
            & ", ARTCUST1.CUST_CRED_LIMIT_REV, ARTCUST1.CUST_CREDIT_RELEASE" & vbCrLf _
            & " from ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE in (Select Distinct CUST_CREDIT_GROUP_CUST from " & ARTCUST1 & ")"
        ARTCUST1_CG = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_CG & " Add Primary Key (CUST_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_CG & " Add CUST_AMT_PICK NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_CG & " Add CUST_AMT_OPEN NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_CG & " Add CUST_AMT_PICK_NOW NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_CG & " Add CUST_BALANCE NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_CG & " Add CUST_BALANCE_PD NUMBER (13,2)")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_CG & " Add CUST_HOLDS_CREDIT VARCHAR2(20)")

        ' INVOICES ONLY, BELOW, PROBABLY SHOULD BE PARAMETERIZED
        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select ARTCUST1.CUST_CREDIT_GROUP_CUST CUST_CODE" & vbCrLf _
            & ", Sum (ARTOPEN1.INV_BALANCE) CUST_BALANCE" & vbCrLf _
            & ", Sum (CASE WHEN ARTOPEN1.INV_BALANCE > 0 and ARTOPEN1.INV_TYPE = 'I' and ARTOPEN1.INV_DUE_DATE + NVL(ARTCUST1_CG.CUST_PD_GRACE_DAYS,0) +1 < SYSDATE THEN ARTOPEN1.INV_BALANCE ELSE 0 END) CUST_BALANCE_PD" & vbCrLf _
            & "   from ARTOPEN1," & ARTCUST1 & " ARTCUST1," & ARTCUST1_CG & " ARTCUST1_CG" & vbCrLf _
            & "   where ARTCUST1.CUST_CODE = ARTOPEN1.CUST_CODE" & vbCrLf _
            & "     and ARTCUST1_CG.CUST_CODE = ARTCUST1.CUST_CREDIT_GROUP_CUST" & vbCrLf _
            & IIf(InArguments.Client = "INT", "     and ARTOPEN1.INV_TYPE = 'I'" & vbCrLf, "") _
            & "   group by ARTCUST1.CUST_CREDIT_GROUP_CUST;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ARTCUST1_CG & " ARTCUST1 Set" & vbCrLf _
            & "    CUST_BALANCE = R1.CUST_BALANCE" & vbCrLf _
            & "   ,CUST_BALANCE_PD = R1.CUST_BALANCE_PD" & vbCrLf _
            & "    where CUST_CODE = R1.CUST_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Update " & ARTCUST1_CG & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'P' where NVL(CUST_BALANCE_PD,0) > NVL(CUST_CREDIT_LIMIT,0) * NVL(CUST_PD_GRACE_PCT,0) / 100 and NVL(CUST_CREDIT_RELEASE,'?') NOT IN ('I','N')")
        ASCDATA1.ExecuteSQL("Update " & ARTCUST1_CG & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'C' where NVL(CUST_CREDIT_HOLD,'0') = '1'")
        ASCDATA1.ExecuteSQL("Update " & ARTCUST1_CG & " Set CUST_HOLDS_CREDIT = NVL(CUST_HOLDS_CREDIT,'') || 'Z' where (NVL(CUST_CREDIT_LIMIT,0) <=0 or CUST_CRED_LIMIT_REV is Null or CUST_CRED_LIMIT_REV < SYSDATE) and NVL(CUST_CREDIT_RELEASE,'?') <> 'N'")

        ASCMAIN1.sql = "Select * from " & ARTCUST1_CG
        Create_TDA(dst.Tables.Add("ARTCUST1_CG"), ARTCUST1_CG, "**", 0)
        Fill_Records("ARTCUST1_CG")

        Create_Relation("ARTCUST1_CG", "ARTCUST1", "CUST_CODE", "CUST_CREDIT_GROUP_CUST")
        With dst.Tables("ARTCUST1_CG")
            .Columns("CUST_AMT_PICK").Expression = "SUM(CHILD(ARTCUST1_CG_ARTCUST1).CUST_AMT_PICK)"
            .Columns("CUST_AMT_OPEN").Expression = "SUM(CHILD(ARTCUST1_CG_ARTCUST1).CUST_AMT_OPEN)"
            .Columns("CUST_AMT_PICK_NOW").Expression = "SUM(CHILD(ARTCUST1_CG_ARTCUST1).CUST_AMT_PICK_NOW)"
        End With


        ' Prepare Allocation Status Tables

        ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from " & SOTORDR2 & " SOTORDR2," & SOTORDR1 & " SOTORDR1,ARTCUST1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTORDR2.ALLO_CTL_NO IS NOT NULL" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE)"

        ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTORDR2.ALLO_CTL_NO in (Select DISTINCT ALLO_CTL_NO from " & SOTORDR2 & ")" & vbCrLf _
            & "   and SOTORDR2.ORDR_STATUS in ('O','P','F')" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE)"

        ASCMAIN1.sql = "Select X.ALLO_CTL_NO" & vbCrLf _
            & ", X.CUST_CODE" & vbCrLf _
            & ", NVL(SOTALLO2.QTY_ALLO,0) QTY_ALLO" & vbCrLf _
            & ", X.ORDR_QTY" & vbCrLf _
            & ", X.ORDR_QTY_OPEN" & vbCrLf _
            & ", X.ORDR_QTY_PICK" & vbCrLf _
            & ", X.ORDR_QTY_SHIP" & vbCrLf _
            & ", X.ORDR_QTY_CANC" & vbCrLf _
            & ", 0 ORDR_QTY_PICK_NOW" & vbCrLf _
            & ", 0 ORDR_QTY_PICK_LATER" & vbCrLf _
            & " from (" & ASCMAIN1.sql & ") X,SOTALLO2" & vbCrLf _
            & " where SOTALLO2.ALLO_CTL_NO (+) = X.ALLO_CTL_NO" & vbCrLf _
            & "   and SOTALLO2.CUST_CODE (+) = X.CUST_CODE"
        SOTALLOZ = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOZ & " Add Primary Key (ALLO_CTL_NO,CUST_CODE)")

        ASCMAIN1.sql = "Select * from " & SOTALLOZ
        Create_TDA(dst.Tables.Add("SOTALLOZ"), SOTALLOZ, "**", 0)
        Fill_Records("SOTALLOZ")

        ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", SOTORDR2.CUST_STORE_NO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and NVL(ARTCUST1.CUST_ALLOCATE_BY_STORE,'0') = '1'" & vbCrLf _
            & "   and SOTORDR2.ALLO_CTL_NO in (Select DISTINCT ALLO_CTL_NO from " & SOTORDR2 & ")" & vbCrLf _
            & "   and SOTORDR2.ORDR_STATUS in ('O','P','F')" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE)" & vbCrLf _
            & ", SOTORDR2.CUST_STORE_NO"

        ASCMAIN1.sql = "Select X.ALLO_CTL_NO" & vbCrLf _
            & ", X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", NVL(SOTALLO3.QTY_ALLO,0) QTY_ALLO" & vbCrLf _
            & ", X.ORDR_QTY" & vbCrLf _
            & ", X.ORDR_QTY_OPEN" & vbCrLf _
            & ", X.ORDR_QTY_PICK" & vbCrLf _
            & ", X.ORDR_QTY_SHIP" & vbCrLf _
            & ", X.ORDR_QTY_CANC" & vbCrLf _
            & ", 0 ORDR_QTY_PICK_NOW" & vbCrLf _
            & ", 0 ORDR_QTY_PICK_LATER" & vbCrLf _
            & " from (" & ASCMAIN1.sql & ") X,SOTALLO3" & vbCrLf _
            & " where SOTALLO3.ALLO_CTL_NO (+) = X.ALLO_CTL_NO" & vbCrLf _
            & "   and SOTALLO3.CUST_CODE (+) = X.CUST_CODE" & vbCrLf _
            & "   and SOTALLO3.CUST_STORE_NO (+) = X.CUST_STORE_NO"
        SOTALLOY = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTALLOY & " Add Primary Key (ALLO_CTL_NO,CUST_CODE,CUST_STORE_NO)")

        ASCMAIN1.sql = "Select * from " & SOTALLOY
        Create_TDA(dst.Tables.Add("SOTALLOY"), SOTALLOY, "**", 0)
        Fill_Records("SOTALLOY")

        ASCMAIN1.sql = "Select * from SOTALLO1 where ALLO_CTL_NO in (Select Distinct ALLO_CTL_NO from " & SOTALLOZ & ")"
        Create_TDA(dst.Tables.Add("SOTALLO1"), SOTALLO1, "**", 0, False)
        Fill_Records("SOTALLO1")


        With dst.Tables.Add("SOTORELX")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("ORDR_GROUP_NO")
            .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE"), .Columns("ORDR_GROUP_NO")}
        End With

        With dst.Tables.Add("SOTORELC")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("CUST_CODE")
            .Columns.Add("ORDR_NO")
            .Columns.Add("ORDR_LNO", GetType(System.Int64))
            .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE"), _
                                            .Columns("CUST_CODE"), _
                                            .Columns("ORDR_NO"), _
                                            .Columns("ORDR_LNO")}
        End With


        ' Prepare Market Restrictions Table

        ASCMAIN1.sql = "Select MARKET_CODE, ITEM_CODE" & vbCrLf _
            & ", Sum (FORECAST) FORECAST, Sum (SHIPPED) SHIPPED, Sum (IN_PICK) IN_PICK, 0 RELEASED from (" & vbCrLf _
            & "Select DPTITMF1.MARKET_CODE, DPTITMF1.ITEM_CODE" & vbCrLf _
            & ", Sum (DPTITMF1.FORECAST) FORECAST, 0 SHIPPED, 0 IN_PICK" & vbCrLf _
            & " from DPTITMF1,SOTMKTC1" & vbCrLf _
            & " where SOTMKTC1.MARKET_CODE = DPTITMF1.MARKET_CODE" & vbCrLf _
            & "   and SOTMKTC1.RESTRICT_ORDER_RELEASE = '1'" & vbCrLf _
            & "   and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and OPS_YYYYPP_FC IN ('000000','" & ASCMAIN1.CYP & "')" & vbCrLf _
            & " group by DPTITMF1.MARKET_CODE, DPTITMF1.ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTTCLS1.MARKET_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", 0 FORECAST, Sum (ORDR_QTY_SHIP) SHIPPED, 0 IN_PICK" & vbCrLf _
            & " from SOTINVH2,ARTCUST1,SOTTCLS1,SOTMKTC1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE = SOTTCLS1.MARKET_CODE" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and SOTMKTC1.RESTRICT_ORDER_RELEASE = '1'" & vbCrLf _
            & " group by SOTTCLS1.MARKET_CODE, SOTINVH2.ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTTCLS1.MARKET_CODE, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", 0 FORECAST, 0 SHIPPED, Sum (ORDR_QTY_PICK) IN_PICK" & vbCrLf _
            & " from SOTORDR2,ARTCUST1,SOTTCLS1,SOTMKTC1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE = SOTTCLS1.MARKET_CODE" & vbCrLf _
            & "   and SOTORDR2.ORDR_STATUS = 'P'" & vbCrLf _
            & "   and SOTMKTC1.RESTRICT_ORDER_RELEASE = '1'" & vbCrLf _
            & " group by SOTTCLS1.MARKET_CODE, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ") group by MARKET_CODE, ITEM_CODE" & vbCrLf
        DPTITMFX = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ASCDATA1.ExecuteSQL("Alter Table " & DPTITMFX & " Add Primary Key (MARKET_CODE, ITEM_CODE)")

        ASCMAIN1.sql = "Select * from " & DPTITMFX
        Create_TDA(dst.Tables.Add("DPTITMFX"), DPTITMFX, "**", 0)
        Fill_Records("DPTITMFX")


        ' Perform Order Release

        Dim ORDERS_RELEASED_TOTAL As Decimal = 0
        ' Stop ' SOTORDR0 SHOULD FILL FROM TEMP SOTORDR0, AND UPDATE IT
        Fill_Records("SOTORDR0")
        Fill_Records("SOTORDR1")
        Fill_Records("SOTORDR2") ' NEEDS ORDR_GROUP_NO, CUST_CODE_ALLO

        For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select _
            ("", "ORDR_PRIORITY DESC, CUST_PRIORITY_CODE, CUST_CODE, ORDR_GROUP_NO")

            Dim ORDERS_RELEASED As Integer = 0
            Dim ORDERS_HELD As Integer = 0
            Dim ORDR_REL_HOLD_CODES As String = ""

            Dim CUST_CODE As String = rowSOTORDR0.Item("CUST_CODE")

            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            Dim CUST_CODE_ALLO As String = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
            Dim CUST_ALLOCATE_BY_STORE As String = rowARTCUST1.Item("CUST_ALLOCATE_BY_STORE") & ""
            Dim CUST_BILL_TO_CUST As String = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
            Dim CUST_CREDIT_GROUP_CUST As String = rowARTCUST1.Item("CUST_CREDIT_GROUP_CUST") & ""

            Dim MARKET_CODE As String = rowARTCUST1.Item("MARKET_CODE") & ""
            Dim CUST_CASE_PACK_ONLY As String = rowARTCUST1.Item("CUST_CASE_PACK_ONLY") & ""
            Dim CUST_ALLOW_BACKORDER As String = rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & ""
            Dim CUST_AMT_PICK As Decimal = Val(rowARTCUST1.Item("CUST_AMT_PICK") & "")
            Dim CUST_AMT_OPEN As Decimal = Val(rowARTCUST1.Item("CUST_AMT_OPEN") & "")

            Dim rowARTCUST1_CG As DataRow = dst.Tables("ARTCUST1_CG").Rows.Find(CUST_CREDIT_GROUP_CUST)
            Dim CUST_AMT_PICK_BT As Decimal = Val(rowARTCUST1_CG.Item("CUST_AMT_PICK") & "")
            Dim CUST_AMT_OPEN_BT As Decimal = Val(rowARTCUST1_CG.Item("CUST_AMT_OPEN") & "")
            Dim CUST_CREDIT_LIMIT As Decimal = Val(rowARTCUST1_CG.Item("CUST_CREDIT_LIMIT") & "")
            Dim CUST_BALANCE As Decimal = Val(rowARTCUST1_CG.Item("CUST_BALANCE") & "")

            Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO") & ""
            Dim ORDR_PRIORITY As String = rowSOTORDR0.Item("ORDR_PRIORITY") & ""

            Dim ORDR_NO_ctr As Integer = 0

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select _
                ("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'", "CUST_STORE_NO")
                ORDR_NO_ctr += 1
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")

                Dim ORDR_HOLD_INVTY As String = ""

                Dim ORDR_CRED_CLEARED As String = rowSOTORDR1.Item("ORDR_CRED_CLEARED") & ""
                Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE") & ""
                Dim ORDR_OVERRIDE_NOT_ALLOCATED As String = rowSOTORDR1.Item("ORDR_OVERRIDE_NOT_ALLOCATED") & ""

                Dim ORDR_HOLD_SALES As String = rowARTCUST1.Item("CUST_HOLDS_SALES") & ""
                If rowSOTORDR1.Item("ORDR_HOLD") & "" = "1" Then
                    ORDR_HOLD_SALES = "1"
                End If

                Dim ORDR_HOLD_CREDIT As String = rowARTCUST1_CG.Item("CUST_HOLDS_CREDIT") & ""

                Dim TERM_CODE As String = rowSOTORDR1.Item("TERM_CODE")
                Dim rowTATTERM1 As DataRow = dst.Tables("TATTERM1").Rows.Find(TERM_CODE)

                If rowTATTERM1.Item("TERM_BYPASS_CREDIT") & "" = "1" Or rowTATTERM1.Item("TERM_TYPE") & "" = "C" Then
                    ORDR_HOLD_CREDIT = Replace(ORDR_HOLD_CREDIT, "P", "")
                    ORDR_HOLD_CREDIT = Replace(ORDR_HOLD_CREDIT, "Z", "")
                End If

                If InArguments.ForcePick Then
                    ORDR_HOLD_SALES = ""
                    ORDR_HOLD_CREDIT = ""
                End If

                Dim ORDR_REL_HOLD_CODES_hdr As String = ""
                If ORDR_HOLD_SALES = "1" Then
                    ORDR_REL_HOLD_CODES_hdr = "S"
                End If

                If rowTATTERM1.Item("TERM_TYPE") & String.Empty = "D" _
                    AndAlso rowSOTORDR1.Item("CCPA_NO") & String.Empty = String.Empty _
                    AndAlso rowSOTORDR1.Item("CC_TRANS_ID") & String.Empty = String.Empty Then
                    ORDR_REL_HOLD_CODES_hdr &= "R"
                    ORDR_HOLD_SALES = "1"
                End If

                Dim THIS_ORDR_AMT_PICK As Decimal = 0
                Dim THIS_ORDR_QTY_PICK As Decimal = 0

                dst.Tables("SOTORELA").Rows.Clear()

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select _
                    ("ORDR_NO = '" & ORDR_NO & "'", "ORDR_LNO")

                    Dim ORDR_REL_HOLD_CODES_dtl As String = ""
                    Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")

                    Dim rowSOTOREL1 As DataRow = dst.Tables("SOTOREL1").Rows.Find(ITEM_CODE)

                    If rowSOTOREL1 Is Nothing Then
                        rowSOTOREL1 = dst.Tables("SOTOREL1").NewRow
                        rowSOTOREL1.Item("ITEM_CODE") = ITEM_CODE
                    End If

                    Dim rowSOTOREL2 As DataRow = dst.Tables("SOTOREL2").Rows.Find(New String() {ITEM_CODE, WHSE_CODE})
                    If rowSOTOREL2 Is Nothing Then
                        rowSOTOREL2 = dst.Tables("SOTOREL2").Rows.Add(New Object() {ITEM_CODE, WHSE_CODE})

                        ASCMAIN1.sql = "Select ITEM_CODE, WHSE_CODE" & vbCrLf _
                             & ", WHSE_QTY_ON_HAND QTY_ON_HAND, WHSE_QTY_OPEN QTY_OPEN, WHSE_QTY_PICK QTY_PICK" & vbCrLf _
                             & " from ICTSTAT2 where ITEM_CODE = :PARM1 and WHSE_CODE = :PARM2"
                        Dim rowICTSTAT2 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VV", New Object() {ITEM_CODE, WHSE_CODE})
                        If rowICTSTAT2 IsNot Nothing Then
                            rowSOTOREL2.Item("QTY_ON_HAND") = rowICTSTAT2.Item("QTY_ON_HAND")
                            rowSOTOREL2.Item("QTY_OPEN") = rowICTSTAT2.Item("QTY_OPEN")
                            rowSOTOREL2.Item("QTY_PICK") = rowICTSTAT2.Item("QTY_PICK")
                        End If
                    End If

                    Dim ITEM_ORDR_REL_CODE As String = rowSOTOREL1.Item("ITEM_ORDR_REL_CODE") & ""
                    Dim ITEM_SNU_CODE As String = rowSOTOREL1.Item("ITEM_SNU_CODE") & ""
                    Dim ITEM_BASIC_PROMO As String = rowSOTOREL1.Item("ITEM_BASIC_PROMO") & ""
                    Dim ITEM_NOT_ALLOCATED As String = rowSOTOREL1.Item("ITEM_NOT_ALLOCATED") & ""

                    Dim ITEM_IS_ALLOCATED As String = rowSOTOREL1.Item("ITEM_IS_ALLOCATED") & ""

                    If InArguments.ReleaseCompleteOnly Then
                        If ITEM_ORDR_REL_CODE = "R" Or ITEM_ORDR_REL_CODE = "S" Then
                            ITEM_ORDR_REL_CODE = ""
                        End If
                    End If

                    Dim ITEM_SO_QTY_MULT As Int64 = Val(rowSOTOREL1.Item("ITEM_SO_QTY_MULT") & "")
                    Dim ITEM_SO_QTY_MIN As Int64 = Val(rowSOTOREL1.Item("ITEM_SO_QTY_MIN") & "")

                    Dim ORDR_QTY_PICK As Int64 = 0
                    Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")

                    Dim rowSOTORELA As DataRow = dst.Tables("SOTORELA").Rows.Find(ITEM_CODE)
                    If rowSOTORELA Is Nothing Then
                        rowSOTORELA = dst.Tables("SOTORELA").Rows.Add(New Object() {ITEM_CODE, 0, 0, 0})
                    Else
                        'If InStr(ORDR_REL_HOLD_CODES_hdr, "2") = 0 Then
                        '    ORDR_HOLD_INVTY = "1"
                        '    ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "2"
                        'End If
                    End If

                    Dim ORDR_QTY_PICK_previous_lines As Int64 = Val(rowSOTORELA.Item("ORDR_QTY_PICK") & "")

                    If ITEM_ORDR_REL_CODE = "D" Or ORDR_QTY_OPEN <= 0 Then
                        ORDR_REL_HOLD_CODES_dtl = "D"
                    Else
                        If ITEM_ORDR_REL_CODE = "H" Then
                            ORDR_REL_HOLD_CODES_dtl = ITEM_ORDR_REL_CODE
                            ORDR_HOLD_INVTY = "1"
                            If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                            End If
                        Else
                            Dim ALLO_CTL_NO As String = rowSOTORDR2.Item("ALLO_CTL_NO") & ""
                            If ALLO_CTL_NO <> "" Then
                                If ORDR_OVERRIDE_NOT_ALLOCATED = "1" Or ITEM_NOT_ALLOCATED = "1" Then
                                    rowSOTORDR2.Item("ALLO_CTL_NO") = ""
                                    ALLO_CTL_NO = ""
                                End If
                            End If

                            If ALLO_CTL_NO = "" Then
                                If (ITEM_IS_ALLOCATED = "1" Or (ITEM_SNU_CODE = "S" And ITEM_BASIC_PROMO = "P") Or ITEM_SNU_CODE = "N") And ITEM_NOT_ALLOCATED <> "1" Then
                                    If ORDR_OVERRIDE_NOT_ALLOCATED = "1" Then
                                        ' LET THE ORDER GO
                                    Else
                                        dst.Tables("SOTORELC").Rows.Add(New String() {ITEM_CODE, _
                                              CUST_CODE_ALLO, _
                                              rowSOTORDR2.Item("ORDR_NO"), _
                                              rowSOTORDR2.Item("ORDR_LNO")})

                                        rowSOTORDR2.Item("ALLO_CTL_NO") = "0000000000"
                                        ORDR_REL_HOLD_CODES_dtl = "C"
                                        ORDR_HOLD_INVTY = "1"
                                        If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                            ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                        End If
                                    End If
                                End If
                            Else
                                Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                                If rowSOTALLO1 Is Nothing Then
                                    ASCMAIN1.sql = "Select * from SOTALLO1 where ALLO_CTL_NO = '" & ALLO_CTL_NO & "'"
                                    Fill_Records("SOTALLO1", "", False, ASCMAIN1.sql)
                                    rowSOTALLO1 = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                                End If
                                Dim rowSOTALLOZ As DataRow = dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO})
                                If rowSOTALLOZ Is Nothing Then
                                    rowSOTALLOZ = dst.Tables("SOTALLOZ").NewRow
                                    rowSOTALLOZ.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                                    rowSOTALLOZ.Item("CUST_CODE") = CUST_CODE_ALLO
                                    rowSOTALLOZ.Item("QTY_ALLO") = 0
                                    dst.Tables("SOTALLOZ").Rows.Add(rowSOTALLOZ)
                                End If

                                If Val(rowSOTALLOZ.Item("QTY_ALLO") & "") = 0 And ORDR_OVERRIDE_NOT_ALLOCATED = "1" Then
                                    ' LET THE ORDER GO
                                Else

                                    Dim enable_this_option As Boolean = False
                                    If Format(DATETIME_STAMP, "yyyyMMdd") < Format(rowSOTALLO1.Item("DATE_START"), "yyyyMMdd") _
                                        And enable_this_option Then
                                        ORDR_REL_HOLD_CODES_dtl = "P"
                                        ORDR_HOLD_INVTY = "1"
                                        If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                            ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                        End If
                                    End If

                                    Dim cBAL As Int64 = Val(rowSOTALLOZ.Item("QTY_ALLO") & "") _
                                         - Val(rowSOTALLOZ.Item("ORDR_QTY_SHIP") & "") _
                                         - Val(rowSOTALLOZ.Item("ORDR_QTY_PICK") & "") _
                                         - Val(rowSOTALLOZ.Item("ORDR_QTY_PICK_NOW") & "")
                                    If cBAL < 0 Then cBAL = 0

                                    '    If ASCMAIN1.Running_in_VS And cBAL - Val(rowSOTALLOZ.Item("ORDR_QTY_PICK_LATER") & "") < 0 Then Stop
                                    If InArguments.Client = "INT" Then
                                        cBAL = cBAL - Val(rowSOTALLOZ.Item("ORDR_QTY_PICK_LATER") & "")
                                    End If

                                    If ORDR_QTY_OPEN + ORDR_QTY_PICK_previous_lines > cBAL And rowSOTALLO1.Item("ALLOW_OVER") & "" <> "1" Then
                                        ORDR_REL_HOLD_CODES_dtl = "C"
                                        ORDR_HOLD_INVTY = "1"
                                        If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                            ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                        End If
                                        dst.Tables("SOTORELC").Rows.Add(New String() {rowSOTORDR2.Item("ITEM_CODE"), _
                                            CUST_CODE_ALLO, _
                                            rowSOTORDR2.Item("ORDR_NO"), _
                                            rowSOTORDR2.Item("ORDR_LNO")})
                                    End If


                                    If CUST_ALLOCATE_BY_STORE = "1" Then
                                        Dim rowSOTALLOY As DataRow = dst.Tables("SOTALLOY").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO, CUST_STORE_NO})
                                        If rowSOTALLOY Is Nothing Then
                                            rowSOTALLOY = dst.Tables("SOTALLOY").NewRow
                                            rowSOTALLOY.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                                            rowSOTALLOY.Item("CUST_CODE") = CUST_CODE_ALLO
                                            rowSOTALLOY.Item("CUST_STORE_NO") = CUST_STORE_NO
                                            rowSOTALLOY.Item("QTY_ALLO") = 0
                                            dst.Tables("SOTALLOY").Rows.Add(rowSOTALLOY)
                                        End If

                                        Dim csBAL As Int64 = Val(rowSOTALLOY.Item("QTY_ALLO") & "") _
                                             - Val(rowSOTALLOY.Item("ORDR_QTY_SHIP") & "") _
                                             - Val(rowSOTALLOY.Item("ORDR_QTY_PICK") & "") _
                                             - Val(rowSOTALLOY.Item("ORDR_QTY_PICK_NOW") & "")
                                        If csBAL < 0 Then csBAL = 0

                                        If ORDR_QTY_OPEN + ORDR_QTY_PICK_previous_lines > csBAL And rowSOTALLO1.Item("ALLOW_OVER") & "" <> "1" Then
                                            ORDR_REL_HOLD_CODES_dtl = "C"
                                            ORDR_HOLD_INVTY = "1"
                                            If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                                ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                            End If
                                            ' MAYBE WE NEED A DIFFERENT TABLE HERE
                                            If Not dst.Tables("SOTORELC").Rows.Contains(New String() {rowSOTORDR2.Item("ITEM_CODE"), _
                                                CUST_CODE_ALLO, _
                                                rowSOTORDR2.Item("ORDR_NO"), _
                                                rowSOTORDR2.Item("ORDR_LNO")}) Then
                                                dst.Tables("SOTORELC").Rows.Add(New String() {rowSOTORDR2.Item("ITEM_CODE"), _
                                                    CUST_CODE_ALLO, _
                                                    rowSOTORDR2.Item("ORDR_NO"), _
                                                    rowSOTORDR2.Item("ORDR_LNO")})
                                            End If

                                        End If
                                    End If

                                End If
                            End If

                            If rowARTCUST1.Item("RESTRICT_ORDER_RELEASE") & "" = "1" Then
                                Dim rowDPTITMFX As DataRow = dst.Tables("DPTITMFX").Rows.Find(New String() {MARKET_CODE, ITEM_CODE})
                                If rowDPTITMFX Is Nothing Then
                                    rowDPTITMFX = dst.Tables("DPTITMFX").Rows.Add(New Object() {MARKET_CODE, ITEM_CODE})
                                End If

                                If Val(rowDPTITMFX.Item("FORECAST") & "") _
                                 - Val(rowDPTITMFX.Item("SHIPPED") & "") _
                                 - Val(rowDPTITMFX.Item("IN_PICK") & "") _
                                 - Val(rowDPTITMFX.Item("RELEASED") & "") _
                                 - (ORDR_QTY_OPEN + ORDR_QTY_PICK_previous_lines) < 0 Then
                                    ORDR_REL_HOLD_CODES_dtl = "M"
                                    ORDR_HOLD_INVTY = "1"
                                    If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                        ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                    End If
                                End If
                            End If

                            If InArguments.ForcePick Then
                                ' Terms requiring CC Authorization do not honor Force Pick. user must chnage terms on sales order
                                If Not ORDR_REL_HOLD_CODES_hdr.Contains("R") Then
                                    ORDR_HOLD_INVTY = ""
                                    ORDR_REL_HOLD_CODES_hdr = ""
                                    ORDR_REL_HOLD_CODES_dtl = ""
                                End If
                            End If

                            ' Authorizations and Order Qty Validation

                            If InArguments.Client = "INT" Then
                                Dim tblSOTORDR2 As DataTable = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & rowSOTORDR2.Item("ORDR_LNO")).CopyToDataTable
                                Dim listData As New List(Of String)
                                Dim errorMessages As String = String.Empty
                                errorMessages = SOCMAIN1.ValidateAuthorizations(tblSOTORDR2, listData)
                                If listData.Count > 0 Then
                                    If (ORDR_REL_HOLD_CODES_dtl & String.Empty).ToString.Length = 0 Then ORDR_REL_HOLD_CODES_dtl = "E"
                                    If InStr(ORDR_REL_HOLD_CODES_hdr, "E") = 0 Then
                                        ORDR_REL_HOLD_CODES_hdr &= "E"
                                    End If
                                End If

                                listData.Clear()
                                errorMessages = SOCMAIN1.ValidateOrderQtys(tblSOTORDR2, listData, "ORDR_QTY_OPEN")
                                If listData.Count > 0 Then
                                    If (ORDR_REL_HOLD_CODES_dtl & String.Empty).ToString.Length = 0 Then ORDR_REL_HOLD_CODES_dtl = "Q"
                                    If InStr(ORDR_REL_HOLD_CODES_hdr, "Q") = 0 Then
                                        ORDR_REL_HOLD_CODES_hdr &= "Q"
                                    End If
                                End If
                            End If

                            If ORDR_REL_HOLD_CODES_dtl = "" Then
                                Dim ORDR_QTY_AVAIL As Int64 = Val(rowSOTOREL2.Item("QTY_ON_HAND") & "") _
                                                           - Val(rowSOTOREL2.Item("QTY_PICK") & "") _
                                                           - Val(rowSOTOREL2.Item("QTY_TO_PICK") & "") _
                                                           - ORDR_QTY_PICK_previous_lines
                                If ORDR_QTY_AVAIL >= ORDR_QTY_OPEN Or InArguments.ForcePick Then
                                    ORDR_QTY_PICK = ORDR_QTY_OPEN
                                    ORDR_REL_HOLD_CODES_dtl = "A"
                                Else ' We have an Inventory Shortage
                                    If (ITEM_ORDR_REL_CODE = "S" Or ITEM_ORDR_REL_CODE = "R") Then
                                        ORDR_REL_HOLD_CODES_dtl = ITEM_ORDR_REL_CODE
                                        Dim QTY_MIN As Int64 = 1
                                        If InArguments.Client = "INT" Then
                                            QTY_MIN = ITEM_SO_QTY_MIN
                                        End If
                                        If ORDR_QTY_AVAIL > 0 And ORDR_QTY_AVAIL >= QTY_MIN Then
                                            ORDR_QTY_PICK = ORDR_QTY_AVAIL
                                            If CUST_CASE_PACK_ONLY = "1" And ITEM_SO_QTY_MULT <> 0 Then
                                                If ORDR_QTY_PICK Mod ITEM_SO_QTY_MULT <> 0 Then
                                                    ORDR_QTY_PICK = ORDR_QTY_PICK - ORDR_QTY_PICK Mod ITEM_SO_QTY_MULT
                                                End If
                                            End If
                                        End If
                                    Else
                                        ORDR_REL_HOLD_CODES_dtl = "I"
                                        ORDR_HOLD_INVTY = "1"
                                        If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                            ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                    THIS_ORDR_AMT_PICK = THIS_ORDR_AMT_PICK + ORDR_QTY_PICK * Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")
                    THIS_ORDR_QTY_PICK = THIS_ORDR_QTY_PICK + ORDR_QTY_PICK

                    rowSOTORDR2.Item("ORDR_RELEASE") = ORDR_REL_HOLD_CODES_dtl
                    rowSOTORDR2.Item("ORDR_QTY_PICK") = ORDR_QTY_PICK

                    Dim ORDR_QTY_BACK As Int64 = 0
                    Dim ORDR_QTY_CANC As Int64 = 0
                    If ORDR_QTY_PICK <> ORDR_QTY_OPEN Then
                        If "" = "HELL FROZE OVER" And CUST_ALLOW_BACKORDER = "1" And (ORDR_REL_HOLD_CODES_dtl = "A" Or ORDR_REL_HOLD_CODES_dtl = "R") Then
                            ORDR_QTY_BACK = ORDR_QTY_OPEN - ORDR_QTY_PICK
                        Else
                            ORDR_QTY_CANC = ORDR_QTY_OPEN - ORDR_QTY_PICK
                        End If
                    End If
                    rowSOTORDR2.Item("PICK_QTY_BACK_REL") = ORDR_QTY_BACK
                    rowSOTORDR2.Item("PICK_QTY_CANC_REL") = ORDR_QTY_CANC

                    ' CAUSES MAJOR PROBLEMS WITH ORDR_QTY_CANC - IF THE ORDER DOES NOT RELEASE, THIS IS NOT BEING RESET
                    ' rowSOTORDR2.Item("ORDR_QTY_CANC") = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "") + ORDR_QTY_CANC

                    '  rowSOTORDR2.Item("ORDR_QTY_OPEN") = ORDR_QTY_BACK
                    ' DO THIS AT END OF PROCESS

                    rowSOTORDR2.Item("ORDR_STATUS") = "P"

                    If ORDR_REL_HOLD_CODES_dtl = "A" _
                    Or ORDR_REL_HOLD_CODES_dtl = "S" _
                    Or ORDR_REL_HOLD_CODES_dtl = "R" _
                    Or ORDR_REL_HOLD_CODES_dtl = "D" Then
                        With rowSOTORELA
                            .Item("ORDR_QTY_PICK") = Val(.Item("ORDR_QTY_PICK") & "") + ORDR_QTY_PICK
                            .Item("ORDR_QTY_OPEN") = Val(.Item("ORDR_QTY_OPEN") & "") + ORDR_QTY_OPEN
                            .Item("ORDR_QTY_BACK") = Val(.Item("ORDR_QTY_BACK") & "") + ORDR_QTY_BACK
                        End With
                    End If
                Next

                If THIS_ORDR_AMT_PICK = 0 Then                  ' If this is a $0 shipment, then
                    If InStr(ORDR_HOLD_CREDIT, "C") <> 0 Then   ' If the customer is on a Credit Hold
                        ORDR_HOLD_CREDIT = "C"                  ' Set Order to Held because of the Credit Hold
                    Else                                        '   this prevents shipment of free goods to customers on Credit Hold, and is not force-pickable
                        ORDR_HOLD_CREDIT = ""                   ' Forget any other type of Credit-Related Holds
                    End If
                Else
                    If rowTATTERM1.Item("TERM_BYPASS_CREDIT") & "" = "1" Or rowTATTERM1.Item("TERM_TYPE") & "" = "C" Then
                    Else
                        If THIS_ORDR_AMT_PICK + CUST_AMT_PICK_BT + CUST_BALANCE > CUST_CREDIT_LIMIT And CUST_CREDIT_LIMIT > 0 Then
                            If Not InArguments.ForcePick Then
                                If InStr(ORDR_HOLD_CREDIT, "L") = 0 Then
                                    ORDR_HOLD_CREDIT &= "L"
                                End If
                            End If
                        End If
                    End If

                End If

                If Not InArguments.ForcePick Then
                    If rowSOTORDR1.Item("ORDR_HOLD_SALES") & "" = "1" Or rowSOTORDR1.Item("ORDR_HOLD") & "" = "1" Then
                        ORDR_HOLD_SALES = "1"
                        ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "O"
                    End If
                    If Format(rowSOTORDR1.Item("ORDR_CANCEL_DATE"), "yyyyMMdd") _
                    < Format(DATETIME_STAMP.Date.AddDays(InArguments.FPDCancelFutureDays), "yyyyMMdd") And Not InArguments.ReleasePastCancel Then
                        ORDR_HOLD_SALES = "1"
                        ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "D"
                    End If
                End If

                If ORDR_HOLD_CREDIT <> "" Then
                    Dim i As Integer = 0
                    For i = 1 To Len(ORDR_HOLD_CREDIT)
                        If InStr(ORDR_CRED_CLEARED, Mid(ORDR_HOLD_CREDIT, i, 1)) = 0 Then
                            rowSOTORDR1.Item("ORDR_CRED_CLEARED") = ""
                            rowSOTORDR1.Item("ORDR_CRED_CLR_AUTH") = ""
                            rowSOTORDR1.Item("ORDR_CRED_CLR_BY") = ""
                            rowSOTORDR1.Item("ORDR_CRED_CLR_DATE") = DBNull.Value
                            i = -1
                            Exit For
                        End If
                    Next i
                    If i <> -1 Then
                        ORDR_HOLD_CREDIT = ""
                    End If
                End If

                If ORDR_HOLD_CREDIT <> "" Then
                    ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "Z"
                End If

                If ORDR_HOLD_INVTY = "" And THIS_ORDR_QTY_PICK = 0 Then
                    ORDR_HOLD_SALES = "1"
                    ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "N"
                End If

                If ORDR_REL_HOLD_CODES_hdr <> "" And ORDR_HOLD_SALES <> "1" Then
                    ORDR_HOLD_SALES = "1"
                End If

                ' Check to see if order is releasable
                ' remember that we may de-release all orders in the batch if not all made it

                If ORDR_HOLD_INVTY & ORDR_HOLD_SALES & ORDR_HOLD_CREDIT = "" Then

                    ORDERS_RELEASED = ORDERS_RELEASED + 1
                    rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = ""
                    rowSOTORDR1.Item("ORDR_CRED_HOLD_CODES") = ""
                    rowSOTORDR1.Item("ORDR_STATUS") = "P"
                    rowSOTORDR1.Item("ORDR_DATE_REL") = DATETIME_STAMP.Date
                    rowSOTORDR1.Item("ORDR_REL_BATCH_NO") = Mid(XNO, 5, 6)

                    For Each rowSOTORELA As DataRow In dst.Tables("SOTORELA").Select("")
                        Dim ITEM_CODE As String = rowSOTORELA.Item("ITEM_CODE")
                        Dim ORDR_QTY_PICK As Int64 = Val(rowSOTORELA.Item("ORDR_QTY_PICK") & "")
                        Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORELA.Item("ORDR_QTY_OPEN") & "")
                        Dim ORDR_QTY_BACK As Int64 = Val(rowSOTORELA.Item("ORDR_QTY_BACK") & "")
                        Item_Status_Qtys(ITEM_CODE, WHSE_CODE, _
                                         ORDR_QTY_OPEN, ORDR_QTY_PICK, ORDR_QTY_BACK, _
                                         rowARTCUST1.Item("RESTRICT_ORDER_RELEASE"), MARKET_CODE)
                    Next
                    CUST_AMT_PICK = CUST_AMT_PICK + THIS_ORDR_AMT_PICK
                    CUST_AMT_PICK_BT = CUST_AMT_PICK_BT + THIS_ORDR_AMT_PICK
                    ' NOTE: THIS CALC DOES NOT TAKE INTO ACCT THE BACK ORDER
                    CUST_AMT_OPEN = CUST_AMT_OPEN - 1 * Val(rowSOTORDR1.Item("ORDR_AMT_OPEN") & "")
                    CUST_AMT_OPEN_BT = CUST_AMT_OPEN_BT - 1 * Val(rowSOTORDR1.Item("ORDR_AMT_OPEN") & "")

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' and ISNULL(ALLO_CTL_NO,'') <> ''")
                        Dim ALLO_CTL_NO As String = rowSOTORDR2.Item("ALLO_CTL_NO") & ""
                        Dim rowSOTALLOZ As DataRow = dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO})
                        If rowSOTALLOZ IsNot Nothing Then
                            rowSOTALLOZ.Item("ORDR_QTY_PICK_NOW") = Val(rowSOTALLOZ.Item("ORDR_QTY_PICK_NOW") & "") + Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                        End If
                        If CUST_ALLOCATE_BY_STORE = "1" Then
                            Dim rowSOTALLOY As DataRow = dst.Tables("SOTALLOY").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO, CUST_STORE_NO})
                            If rowSOTALLOY IsNot Nothing Then
                                rowSOTALLOY.Item("ORDR_QTY_PICK_NOW") = Val(rowSOTALLOY.Item("ORDR_QTY_PICK_NOW") & "") + Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                            End If
                        End If
                    Next

                    rowSOTORDR1.Item("ORDR_QTY_PICK") = THIS_ORDR_QTY_PICK
                    rowSOTORDR1.Item("ORDR_AMT_PICK") = THIS_ORDR_AMT_PICK
                Else
                    ORDERS_HELD = ORDERS_HELD + 1
                    rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = ORDR_REL_HOLD_CODES_hdr
                    rowSOTORDR1.Item("ORDR_CRED_HOLD_CODES") = ORDR_HOLD_CREDIT

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
                        Dim ALLO_CTL_NO As String = rowSOTORDR2.Item("ALLO_CTL_NO") & ""
                        If ALLO_CTL_NO <> "" Then

                            Dim rowSOTALLOZ As DataRow = dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO})
                            If rowSOTALLOZ Is Nothing Then
                                rowSOTALLOZ = dst.Tables("SOTALLOZ").NewRow
                                rowSOTALLOZ.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                                rowSOTALLOZ.Item("CUST_CODE") = CUST_CODE_ALLO
                                dst.Tables("SOTALLOZ").Rows.Add(rowSOTALLOZ)
                            End If
                            With rowSOTALLOZ
                                .Item("ORDR_QTY_PICK_LATER") = Val(.Item("ORDR_QTY_PICK_LATER") & "") + Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                            End With

                            If CUST_ALLOCATE_BY_STORE = "1" Then
                                Dim rowSOTALLOY As DataRow = dst.Tables("SOTALLOY").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO, CUST_STORE_NO})
                                If rowSOTALLOY Is Nothing Then
                                    rowSOTALLOY = dst.Tables("SOTALLOY").NewRow
                                    rowSOTALLOY.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                                    rowSOTALLOY.Item("CUST_CODE") = CUST_CODE_ALLO
                                    rowSOTALLOY.Item("CUST_STORE_NO") = CUST_STORE_NO
                                    dst.Tables("SOTALLOY").Rows.Add(rowSOTALLOY)
                                End If
                                With rowSOTALLOY
                                    .Item("ORDR_QTY_PICK_LATER") = Val(.Item("ORDR_QTY_PICK_LATER") & "") + Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                                End With
                            End If
                        End If

                        With rowSOTORDR2
                            .Item("ORDR_QTY_PICK") = 0
                            .Item("ORDR_STATUS") = "O"
                            .Item("PICK_QTY_CANC_REL") = 0
                            .Item("PICK_QTY_BACK_REL") = 0
                        End With
                    Next

                    If ORDR_REL_HOLD_CODES_hdr <> "" Then
                        For i As Integer = 1 To Len(ORDR_REL_HOLD_CODES_hdr)
                            If InStr(ORDR_REL_HOLD_CODES, Mid$(ORDR_REL_HOLD_CODES_hdr, i, 1)) = 0 Then
                                ORDR_REL_HOLD_CODES = ORDR_REL_HOLD_CODES & Mid$(ORDR_REL_HOLD_CODES_hdr, i, 1)
                            End If
                        Next i
                    End If
                End If

            Next

            ORDERS_RELEASED_TOTAL = ORDERS_RELEASED_TOTAL + ORDERS_RELEASED

            rowSOTORDR0.Item("ORDERS_HELD") = ORDERS_HELD
            rowSOTORDR0.Item("ORDERS_RELEASED") = ORDERS_RELEASED
            rowSOTORDR0.Item("ORDR_REL_HOLD_CODES") = ORDR_REL_HOLD_CODES

            rowARTCUST1.Item("CUST_AMT_PICK") = CUST_AMT_PICK
            rowARTCUST1.Item("CUST_AMT_OPEN") = CUST_AMT_OPEN
            rowARTCUST1.Item("RELEASED") = "Y"
        Next


        ' For Groups which are partially released (i.e., some orders released, some orders held)
        ' 1) cancel unreleased orders if the only hold code is N, and released orders > held orders
        '       (this will cause the other orders to successfully release)
        ' 2) or else hold all orders in the group

        For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select _
            ("ORDERS_HELD <> 0 and ORDERS_RELEASED <> 0")
            Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO")
            Dim CUST_CODE As String = rowSOTORDR0.Item("CUST_CODE")
            Dim WHSE_CODE As String = rowSOTORDR0.Item("WHSE_CODE")
            Dim ORDR_REL_HOLD_CODES As String = rowSOTORDR0.Item("ORDR_REL_HOLD_CODES") & ""

            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            Dim CUST_CODE_ALLO As String = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
            Dim CUST_ALLOCATE_BY_STORE As String = rowARTCUST1.Item("CUST_ALLOCATE_BY_STORE") & ""
            Dim CUST_ALLOW_BACKORDER As String = rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & ""
            Dim CUST_AMT_PICK As Decimal = Val(rowARTCUST1.Item("CUST_AMT_PICK") & "")
            Dim CUST_AMT_OPEN As Decimal = Val(rowARTCUST1.Item("CUST_AMT_OPEN") & "")

            'If chkCancelIfNoPick.Checked And ORDR_REL_HOLD_CODES = "N" And _
            '    Val(rowSOTORDR0.Item("ORDERS_RELEASED") & "") > _
            '    Val(rowSOTORDR0.Item("ORDERS_HELD") & "") Then
            'LBM SAYS NUMBER OF ORDERS DON'T MATTER
            If InArguments.CancelIfNoPick And ORDR_REL_HOLD_CODES = "N" Then

                If "" = "HELL FROZE OVER" And CUST_ALLOW_BACKORDER = "1" Then
                    ' no action required
                Else
                    For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select _
                        ("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS = 'O'")
                        Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                        Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")

                        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
                            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
                            Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                            With dst.Tables("SOTOREL2").Rows.Find(New String() {ITEM_CODE, WHSE_CODE})
                                .Item("QTY_TO_DE_COMMIT") = Val(.Item("QTY_TO_DE_COMMIT") & "") + ORDR_QTY_OPEN
                            End With
                            With rowSOTORDR2
                                .Item("ORDR_QTY_CANC") = Val(.Item("ORDR_QTY_CANC") & "") + Val(.Item("ORDR_QTY_OPEN") & "")
                                .Item("PICK_QTY_CANC_REL") = Val(.Item("ORDR_QTY_OPEN") & "")
                                .Item("ORDR_QTY_OPEN") = 0
                                .Item("ORDR_STATUS") = "C"
                                .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                            End With
                        Next
                        With rowSOTORDR1
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                            .Item("LAST_DATE") = DATETIME_STAMP
                            .Item("ORDR_REL_BATCH_NO") = Mid(XNO, 5, 6)
                            .Item("ORDR_DATE_REL") = DATETIME_STAMP.Date
                            .Item("ORDR_STATUS") = "C"
                            .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
                        End With
                    Next

                    rowSOTORDR0.Item("ORDERS_CANCELLED") = Val(rowSOTORDR0.Item("ORDERS_HELD") & "")
                    rowSOTORDR0.Item("ORDERS_HELD") = 0
                End If
            Else

                For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select _
                    ("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_STATUS = 'P'")

                    Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                    Dim CUST_STORE_NO As String = rowSOTORDR1.Item("CUST_STORE_NO")

                    CUST_AMT_PICK -= Val(rowSOTORDR1.Item("ORDR_AMT_PICK") & "")
                    CUST_AMT_OPEN += Val(rowSOTORDR1.Item("ORDR_AMT_OPEN") & "")
                    ORDERS_RELEASED_TOTAL = ORDERS_RELEASED_TOTAL - 1

                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'")
                        Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")

                        Dim ORDR_QTY_PICK As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                        Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                        Dim ORDR_QTY_BACK As Int64 = Val(rowSOTORDR2.Item("PICK_QTY_BACK_REL") & "")
                        Item_Status_Qtys(ITEM_CODE, WHSE_CODE, _
                                         -1 * ORDR_QTY_OPEN, _
                                         -1 * ORDR_QTY_PICK, _
                                         -1 * ORDR_QTY_BACK, _
                                         rowARTCUST1.Item("RESTRICT_ORDER_RELEASE"), _
                                         rowARTCUST1.Item("MARKET_CODE"))

                        Dim ALLO_CTL_NO As String = rowSOTORDR2.Item("ALLO_CTL_NO") & ""
                        If ALLO_CTL_NO <> "" Then
                            With dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO})
                                .Item("ORDR_QTY_PICK_NOW") = Val(.Item("ORDR_QTY_PICK_NOW") & "") - ORDR_QTY_PICK
                                .Item("ORDR_QTY_PICK_LATER") = Val(.Item("ORDR_QTY_PICK_LATER") & "") + ORDR_QTY_PICK
                            End With
                            If CUST_ALLOCATE_BY_STORE = "1" Then
                                With dst.Tables("SOTALLOY").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO, CUST_STORE_NO})
                                    .Item("ORDR_QTY_PICK_NOW") = Val(.Item("ORDR_QTY_PICK_NOW") & "") - ORDR_QTY_PICK
                                    .Item("ORDR_QTY_PICK_LATER") = Val(.Item("ORDR_QTY_PICK_LATER") & "") + ORDR_QTY_PICK
                                End With
                            End If
                        End If
                        rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                        rowSOTORDR2.Item("ORDR_STATUS") = "O"
                    Next

                    If InStr(rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & "", "X") = 0 Then
                        rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") = rowSOTORDR1.Item("ORDR_REL_HOLD_CODES") & "X"
                    End If
                    rowSOTORDR1.Item("ORDR_STATUS") = "O"
                    rowSOTORDR1.Item("ORDR_DATE_REL") = DBNull.Value
                    rowSOTORDR1.Item("ORDR_REL_BATCH_NO") = DBNull.Value
                Next

                rowARTCUST1.Item("CUST_AMT_PICK") = CUST_AMT_PICK
                rowARTCUST1.Item("CUST_AMT_OPEN") = CUST_AMT_OPEN
                rowARTCUST1.Item("RELEASED") = "Y"

                rowSOTORDR0.Item("ORDERS_HELD") = Val(rowSOTORDR0.Item("ORDERS_HELD") & "") + Val(rowSOTORDR0.Item("ORDERS_RELEASED") & "")
                rowSOTORDR0.Item("ORDERS_RELEASED") = 0
            End If
        Next

        ' Misc Order Release Tasks

        'For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("ORDR_REL_BATCH_NO IsNot Null")
        '    With rowSOTORDR1
        '        .Item("ORDR_PICK_SEQ") = Val(.Item("ORDR_PICK_SEQ") & "") + 1
        '    End With
        'Next

        sqlw = "ORDR_STATUS = 'P'"
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
            With rowSOTORDR2
                .Item("ORDR_QTY_OPEN") = 0
            End With
        Next

        sqlw = "(ORDR_STATUS = 'P' or ORDR_STATUS = 'C') and (ORDR_RELEASE = 'R' or ORDR_RELEASE = 'S')"
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR2").Select(sqlw), New String() {"ORDR_NO"}).Rows
            With dst.Tables("SOTORDR1").Rows.Find(row.Item("ORDR_NO"))
                .Item("REORD_MEMO_IND") = "1"
            End With
        Next

        sqlw = "ORDR_STATUS = 'P' and (ISNULL(PICK_QTY_BACK_REL,0) <> 0 or ISNULL(PICK_QTY_CANC_REL,0) <> 0)"
        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
            With rowSOTORDR2
                .Item("ORDR_QTY_OPEN") = Val(.Item("PICK_QTY_BACK_REL") & "")
                .Item("ORDR_QTY_CANC") = Val(.Item("ORDR_QTY_CANC") & "") + Val(.Item("PICK_QTY_CANC_REL") & "")
            End With
        Next

        ' Tally the total qty, by item, on all Orders Not Released (Held Orders)

        sqlw = "ORDR_STATUS = 'O' and ORDR_QTY_OPEN <> 0"
        For Each rowSOTORDR2_unreleased As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("SOTORDR2").Select(sqlw), New String() {"ITEM_CODE", "WHSE_CODE"}).Rows
            Dim ITEM_CODE As String = rowSOTORDR2_unreleased.Item("ITEM_CODE")
            Dim WHSE_CODE As String = rowSOTORDR2_unreleased.Item("WHSE_CODE")
            Dim sqlw2 = "ORDR_STATUS = 'O' and ORDR_QTY_OPEN <> 0 and ITEM_CODE = '" & ITEM_CODE & "' and WHSE_CODE = '" & WHSE_CODE & "'"
            Dim ORDR_QTY_OPEN As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", sqlw2) & "")
            Dim rowSOTOREL2 As DataRow = dst.Tables("SOTOREL2").Rows.Find(New String() {ITEM_CODE, WHSE_CODE})
            rowSOTOREL2.Item("QTY_UNRELEASED") = Val(rowSOTOREL2.Item("QTY_UNRELEASED") & "") + ORDR_QTY_OPEN
        Next

        ' Create Report Table for Items Holding up Orders because of Customer Allocations

        ' including items with over-allocations obscured by a shortage or somesuch

        sqlw = "ISNULL(QTY_ALLO,0) < ISNULL(ORDR_QTY_PICK,0) + ISNULL(ORDR_QTY_SHIP,0) + ISNULL(ORDR_QTY_PICK_NOW,0) + ISNULL(ORDR_QTY_PICK_LATER,0)"
        For Each rowSOTALLOZ As DataRow In dst.Tables("SOTALLOZ").Select(sqlw)
            Dim CUST_CODE_ALLO As String = rowSOTALLOZ.Item("CUST_CODE")
            Dim ALLO_CTL_NO As String = rowSOTALLOZ.Item("ALLO_CTL_NO")
            If ALLO_CTL_NO = "0000000000" Then
                sqlw = "CUST_CODE_ALLO = '" & CUST_CODE_ALLO & "'"
            Else
                sqlw = "ALLO_CTL_NO = '" & ALLO_CTL_NO & "' and CUST_CODE_ALLO = '" & CUST_CODE_ALLO & "'"
            End If
            For Each row As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("SOTORDR2").Select(sqlw), New String() {"ITEM_CODE", "ORDR_GROUP_NO"}).Rows

                If dst.Tables("SOTORELX").Rows.Find(New String() {row.Item("ITEM_CODE"), row.Item("ORDR_GROUP_NO")}) Is Nothing Then
                    ' If ASCMAIN1.Running_in_VS Then Stop '  I WOULD BE SURPRISED IF WE HIT A ROW HERE
                    dst.Tables("SOTORELX").Rows.Add(New String() {row.Item("ITEM_CODE"), row.Item("ORDR_GROUP_NO")})
                End If
            Next
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("SOTORDR2").Select("ORDR_RELEASE='C'"), New String() {"ITEM_CODE", "ORDR_GROUP_NO"}).Rows
            If dst.Tables("SOTORELX").Rows.Find(New String() {row.Item("ITEM_CODE"), row.Item("ORDR_GROUP_NO")}) Is Nothing Then
                '  If ASCMAIN1.Running_in_VS Then Stop '  I WOULD BE SURPRISED IF WE HIT A ROW HERE
                dst.Tables("SOTORELX").Rows.Add(New String() {row.Item("ITEM_CODE"), row.Item("ORDR_GROUP_NO")})
            End If
        Next


        With dst.Tables.Add("SOTOREL9")
            .Columns.Add("MARKET_CODE")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("ORDR_NO")
            .Columns.Add("ORDR_LNO", GetType(System.Int64))
            .PrimaryKey = New DataColumn() {.Columns("MARKET_CODE"), _
                                            .Columns("ITEM_CODE"), _
                                            .Columns("ORDR_NO"), _
                                            .Columns("ORDR_LNO")}
        End With

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_RELEASE = 'M'")
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(rowSOTORDR2.Item("CUST_CODE"))
            dst.Tables("SOTOREL9").Rows.Add(New String() {rowARTCUST1.Item("MARKET_CODE"), _
                                                          rowSOTORDR2.Item("ITEM_CODE"), _
                                                          rowSOTORDR2.Item("ORDR_NO"), _
                                                          rowSOTORDR2.Item("ORDR_LNO")})
        Next

        Return True
    End Function

    Sub Update_SOTORDRx()

        For Each TABLE_NAME As String In New String() {"SOTORDR0", "SOTORDR1", "SOTORDR2"}
            Update_Record_TDA(TABLE_NAME)
        Next

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is Select * from " & SOTORDR1 & ";" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR1 Set" _
            & "      ORDR_REL_HOLD_CODES = R1.ORDR_REL_HOLD_CODES" _
            & "    , ORDR_CRED_HOLD_CODES = R1.ORDR_CRED_HOLD_CODES" _
            & "    , ORDR_CRED_CLEARED = R1.ORDR_CRED_CLEARED" _
            & "    , ORDR_CRED_CLR_AUTH = R1.ORDR_CRED_CLR_AUTH" _
            & "    , ORDR_CRED_CLR_BY = R1.ORDR_CRED_CLR_BY" _
            & "    , ORDR_CRED_CLR_DATE = R1.ORDR_CRED_CLR_DATE" _
            & "    where ORDR_NO = R1.ORDR_NO;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "Update " & SOTORDR1 & " Set REORD_MEMO_IND = '1'" _
            & " where ORDR_NO in (" _
            & " Select Distinct ORDR_NO from " & SOTORDR2 _
            & " where ORDR_RELEASE = 'R')"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is Select * from " & SOTORDR2 & ";" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   Update SOTORDR2 Set" _
            & "      ORDR_QTY_ALLO = R1.ORDR_QTY_ALLO" _
            & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        'ASCDATA1.ExecuteSQL(ASCMAIN1.sql )
        ' MAYBE UPDATE THESE ALSO?
        '& "    , ORDR_RELEASE = R1.ORDR_RELEASE" _
        '& "    , ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" _

    End Sub

    Sub Create_PICK_SHIP_CART()

        ' Create Pick Tickets, Shipment BOL, and Cartons
        '   do this for all Sales Orders scheduled to ship on or before SHIP_BY_DATE
        '   also, filter on Division, Customer, and Order Group

        Dim PICK_RELEASED As Date = DATETIME_STAMP.Date
        'PICK_NO_seq = 0

        '  If ASCMAIN1.Running_in_VS Then Stop ' LINE 486 REALLY NEEDS ITEMS IN A DATATABLE

        SOTPICK1 = Create_Temporary_Table("SOTPICK1", "PICK_NO")
        SOTPICK2 = Create_Temporary_Table("SOTPICK2", "PICK_NO,PICK_LNO")
        SOTSHIP1 = Create_Temporary_Table("SOTSHIP1", "SHIP_BOL_NO")
        SOTCART1 = Create_Temporary_Table("SOTCART1", "CART_NO")
        SOTCART2 = Create_Temporary_Table("SOTCART2", "CART_NO,CART_LNO")

        Create_TDA(dst.Tables.Add(), "SOTPICK0", "*")

        Create_Relation("SOTORDR1", "SOTPICK1", "ORDR_NO")

        With dst.Tables("SOTPICK2").Columns
            .Add("ITEM_CODE")
        End With

        With dst.Tables("SOTCART2").Columns
            .Add("ITEM_WEIGHT", GetType(System.Decimal))
        End With

        Dim SO_PARM_UPC_VENDOR_ID As String = ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID") & ""
        Dim SO_PARM_MAX_CARTON As Integer = Val(ROWs("SOTPARM1").Item("SO_PARM_MAX_CARTON") & "")

        Create_TDA(dst.Tables.Add, "SOTCSTP1", "*", 1, False)

        With dst.Tables.Add("SOTSHIP2")
            .Columns.Add("SHIP_BOL_NO")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("PICK_QTY", GetType(System.Int64))
            .PrimaryKey = New DataColumn() {.Columns("SHIP_BOL_NO"), .Columns("ITEM_CODE")}
        End With

        Dim rowARTCUST1 As DataRow
        Dim CUST_CODE As String = ""
        ' and ORDR_REL_HOLD_CODES is Null
        ' why isn't this ORDR_REL_BATCH_NO = '{current batch}'?
        For Each rowSOTORDR1_rel As DataRow In dst.Tables("SOTORDR1").Select _
                ("ORDR_REL_BATCH_NO is Not Null", _
                 "CUST_CODE, ORDR_GROUP_NO, ORDR_NO")
            If CUST_CODE <> rowSOTORDR1_rel.Item("CUST_CODE") Then
                CUST_CODE = rowSOTORDR1_rel.Item("CUST_CODE")
                rowARTCUST1 = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            End If

            Dim ORDR_NO As String = rowSOTORDR1_rel.Item("ORDR_NO")
            Dim ORDR_STATUS As String = rowSOTORDR1_rel.Item("ORDR_STATUS")
            Dim PICK_SEQ_NO As Integer = Val(rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") & "") + 1
            rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") = PICK_SEQ_NO

            Dim rowSOTPICK1 As DataRow = dst.Tables("SOTPICK1").NewRow
            PICK_NO_seq += 1
            Dim PICK_NO As String = "TEMP" & Format(PICK_NO_seq, "000000")
            With rowSOTPICK1
                .Item("PICK_NO") = PICK_NO
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_PICK_SEQ") = PICK_SEQ_NO
                If ORDR_STATUS = "C" Then
                    .Item("PICK_STATUS") = "C"
                Else
                    .Item("PICK_STATUS") = "P"
                End If

                .Item("PICK_RELEASED") = PICK_RELEASED
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("SHIP_BOL_NO") = "X" ' "TEMP000001" ' IS THIS OK?  DOING THIS JUST TO AVOID THE DATA RELATION ISSUE
            End With
            dst.Tables("SOTPICK1").Rows.Add(rowSOTPICK1)

            Dim sqlw As String = "ORDR_NO = '" & ORDR_NO & "'"
            For Each rowSOTORDR2_rel As DataRow In dst.Tables("SOTORDR2").Select(sqlw, "ORDR_LNO")
                Dim rowSOTPICK2 As DataRow = dst.Tables("SOTPICK2").NewRow
                With rowSOTPICK2
                    .Item("PICK_NO") = PICK_NO
                    .Item("PICK_LNO") = rowSOTORDR2_rel.Item("ORDR_LNO")
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = rowSOTORDR2_rel.Item("ORDR_LNO")
                    .Item("ITEM_CODE") = rowSOTORDR2_rel.Item("ITEM_CODE")
                    .Item("PICK_QTY") = rowSOTORDR2_rel.Item("ORDR_QTY_PICK")
                    .Item("PICK_QTY_CONF") = rowSOTORDR2_rel.Item("ORDR_QTY_PICK")
                    .Item("PICK_UNIT_PRICE") = rowSOTORDR2_rel.Item("ORDR_UNIT_PRICE")
                    .Item("PICK_QTY_CANC_REL") = rowSOTORDR2_rel.Item("PICK_QTY_CANC_REL")
                    .Item("PICK_QTY_BACK_REL") = rowSOTORDR2_rel.Item("PICK_QTY_BACK_REL")
                End With
                dst.Tables("SOTPICK2").Rows.Add(rowSOTPICK2)
            Next

            ' THIS CODE IGNORES BACK ORDER POSSIBILITY
            Dim TOTAL_OPEN As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_OPEN)", sqlw) & "") ' Total Units left OPEN in Order after Release
            Dim TOTAL_PICK As Int64 = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_QTY_PICK)", sqlw) & "") ' Total Units in PICK in Order after Release

            If TOTAL_OPEN = 0 Then
                If TOTAL_PICK = 0 Then
                    If Val(rowSOTORDR1_rel.Item("ORDR_PICK_SEQ") & "") <= 1 Then
                        rowSOTORDR1_rel.Item("ORDR_STATUS") = "C"
                    Else
                        rowSOTORDR1_rel.Item("ORDR_STATUS") = "F"
                    End If
                    ' note: the next 2 fields are not going back to oracle; 
                    ' this code was placed here to make cancel in order release equivalent to cancel in order entry; 
                    ' not really a problem since we are not using these 2 fields for anything yet
                    ' 11/13/16 - NOW WE ARE USING THESE, SO NOW WE UPDATE SOTORDR1
                    rowSOTORDR1_rel.Item("ORDR_DATE_CLOSED") = DATETIME_STAMP.Date
                    rowSOTORDR1_rel.Item("ORDR_YYYYPP_CLOSED") = ASCMAIN1.CYP
                Else
                    rowSOTORDR1_rel.Item("ORDR_STATUS") = "P"
                    rowSOTORDR1_rel.Item("ORDR_DATE_REL") = PICK_RELEASED.Date
                End If
            End If
        Next

        Create_Relation("SOTPICK1", "SOTPICK2", "PICK_NO")
        dst.Tables("SOTPICK2").Columns.Add("SHIP_BOL_NO", GetType(System.String), "PARENT(SOTPICK1_SOTPICK2).SHIP_BOL_NO")

        ' Create_Relation("SOTSHIP2", "SOTPICK2", "SHIP_BOL_NO,ITEM_CODE")
        'dst.Tables("SOTSHIP2").Columns.Add("PICK_QTY", GetType(System.Int64)) ' , "SUM(CHILD(SOTSHIP2_SOTPICK2).PICK_QTY)")

        ' when done, regroup pick tickets by group/dc to assign SHIP_BOL_NO

        Dim CART_LNO_seq As Int64 = 0 = 0
        SHIP_BOL_NO_seq = 0

        ' TROUBLE PRINTING PICK TICKETS IF WE ALLOW BREAKS ON SHIP VIA NOW
        ' & " SOWORDR1.SHIP_VIA_CODE"
        '            & ", NULL SHIP_VIA_CODE " & vbCrLf _


        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
            & ", DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',SOTORDR1.CUST_DC_NO,SOTORDR1.CUST_STORE_NO) SHIP_TO" & vbCrLf _
            & ", MIN (NVL(SOTORDR1.SHIP_VIA_CODE,ARTCUST1.SHIP_VIA_CODE)) SHIP_VIA_CODE" & vbCrLf _
            & ", MIN (SOTORDR1.TERM_CODE) TERM_CODE" & vbCrLf _
            & ", MIN (SOTORDR1.SREP_CODE) SREP_CODE" & vbCrLf _
            & ", MIN (SOTORDR1.FRT_TERMS) FRT_TERMS" & vbCrLf _
            & ", MIN (SOTORDR1.ORDR_DEPT) ORDR_DEPT" & vbCrLf _
            & " from " & SOTORDR1 & " SOTORDR1, ARTCUST1" & vbCrLf _
            & " where SOTORDR1.ORDR_REL_BATCH_NO is Not Null" & vbCrLf _
            & "   and (SOTORDR1.ORDR_REL_HOLD_CODES is Null OR SOTORDR1.ORDR_REL_HOLD_CODES = 'N')" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & " group by SOTORDR1.ORDR_GROUP_NO, SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
            & ", DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',SOTORDR1.CUST_DC_NO,SOTORDR1.CUST_STORE_NO)"

        Dim QTY_TO_PACK As Int64         ' Working Variable to Pack PICK_QTY in to Cartons
        Dim PACK_QTY As Int64            ' Qty to Pack into current carton
        Dim MAX_QTY_CTN_cust As Int64    ' Default for the Maximum Unit Count in a Carton, by Customer, defaulting to SO Parameter
        Dim MAX_QTY_CTN_pick As Int64    ' Default for the Maximum Unit Count in a Carton, for the current Pick Ticket, defaulting to MAX_QTY_CTN_cust, but set to lowest MAX_QTY_CTN of all mixable items on the Pick Ticket with a MAX_QTY_CTN
        Dim LAST_ITEM_CODE As String = ""        'Used to Track Last Item for Splitting options.
        Dim WHSE_CODE As String = ""
        Dim LP_STATUS As String = ""
        Dim ORDR_PICK_TYPE As String = ""
        Dim SHIP_CART_REQD As String = ""

        Dim rowEDTSLSP1 As DataRow = Nothing
        Dim CUST_856_IND As String = ""
        Dim CUST_810_IND As String = ""
        Dim NO_MIXED_CASES As String = ""

        CUST_CODE = ""
        For Each rowSHIPMENT As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("", "WHSE_CODE,CUST_CODE")
            Dim ORDR_ADDR_TYPE_ST As String = rowSHIPMENT.Item("ORDR_ADDR_TYPE_ST")
            Dim ORDR_GROUP_NO As String = rowSHIPMENT.Item("ORDR_GROUP_NO")
            Dim EDI_DOC_SEQ_NO As String = rowSHIPMENT.Item("EDI_DOC_SEQ_NO") & ""

            Dim SHIP_TO As String = rowSHIPMENT.Item("SHIP_TO") & ""
            ' If SHIP_TO = "" Then SHIP_TO = "MK" ' TESTING MK
            Dim SHIP_VIA_CODE As String = rowSHIPMENT.Item("SHIP_VIA_CODE") & ""
            Dim TERM_CODE As String = rowSHIPMENT.Item("TERM_CODE") & ""
            Dim FRT_TERMS As String = rowSHIPMENT.Item("FRT_TERMS") & ""
            Dim SREP_CODE As String = rowSHIPMENT.Item("SREP_CODE") & ""
            Dim ORDR_DEPT As String = rowSHIPMENT.Item("ORDR_DEPT") & ""
            If WHSE_CODE <> rowSHIPMENT.Item("WHSE_CODE") & "" Then
                WHSE_CODE = rowSHIPMENT.Item("WHSE_CODE") & ""
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                    LP_STATUS = "0"
                Else
                    LP_STATUS = ""
                End If
            End If

            If CUST_CODE <> rowSHIPMENT.Item("CUST_CODE") Then
                CUST_CODE = rowSHIPMENT.Item("CUST_CODE")

                rowARTCUST1 = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)

                'If rowARTCUST1.Item("CUST_SHIP_BY_CASE") & "" = "1" Then
                '    ORDR_PICK_TYPE = "C"
                'Else
                '    ORDR_PICK_TYPE = "P"
                'End If
                ORDR_PICK_TYPE = "P"
                'SHIP_CART_REQD = rowARTCUST1.Item("CUST_CART_REQD") & ""
                SHIP_CART_REQD = ""

                Dim rowSOTCSTP2 As DataRow = LookUp("SOTCSTP2", CUST_CODE)
                If rowSOTCSTP2 IsNot Nothing Then
                    MAX_QTY_CTN_cust = Val(rowSOTCSTP2.Item("MAX_QTY_CTN") & "")
                    NO_MIXED_CASES = rowSOTCSTP2.Item("NO_MIXED_CASES") & ""
                Else
                    MAX_QTY_CTN_cust = SO_PARM_MAX_CARTON
                    NO_MIXED_CASES = ""
                End If

                rowEDTSLSP1 = LookUp("EDTSLSP1", CUST_CODE)
                CUST_856_IND = ""
                CUST_810_IND = ""
                If rowEDTSLSP1 IsNot Nothing Then
                    If rowEDTSLSP1.Item("EDI_ID_856") & "" <> "" Then CUST_856_IND = "1"
                    If rowEDTSLSP1.Item("EDI_ID_810") & "" <> "" Then CUST_810_IND = "1"
                End If

            End If

            Dim SHIP_856_IND As String = ""
            Dim SHIP_810_IND As String = ""

            If EDI_DOC_SEQ_NO <> "" Then
                SHIP_810_IND = CUST_810_IND
                SHIP_856_IND = CUST_856_IND
            End If

            SHIP_BOL_NO_seq += 1
            Dim SHIP_BOL_NO As String = "TEMP" & Format$(SHIP_BOL_NO_seq, "000000")

            Dim sqlw As String = "ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
            sqlw &= " and ISNULL(ORDR_REL_BATCH_NO,'') <> ''"
            sqlw &= " and ORDR_ADDR_TYPE_ST = '" & ORDR_ADDR_TYPE_ST & "'"

            If ORDR_ADDR_TYPE_ST = "DC" Then
                sqlw &= " and ISNULL(CUST_DC_NO,'') = '" & SHIP_TO & "'"
            Else
                sqlw &= " and CUST_STORE_NO = '" & SHIP_TO & "'"
            End If

            ' TROUBLE PRINTING PICK TICKETS IF WE ALLOW BREAKS ON SHIP VIA NOW
            'If SHIP_VIA_CODE = "" Then
            '    & "   and SHIP_VIA_CODE is Null"
            'Else
            '    & "   and SHIP_VIA_CODE = '" & SHIP_VIA_CODE & "'"
            'End If

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(sqlw)
                Dim rowSOTPICK1 As DataRow = rowSOTORDR1.GetChildRows("SOTORDR1_SOTPICK1")(0)
                rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
            Next

            Dim sqlBOL As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"

            For Each rowSC As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("SOTPICK2").Select(sqlBOL), _
                 New String() {"ITEM_CODE"}).Rows

                Dim ITEM_CODE As String = rowSC.Item("ITEM_CODE")
                Dim sqlSC As String = " and ITEM_CODE = '" & ITEM_CODE & "'"

                Dim PICK_QTY As Int64 = Val(dst.Tables("SOTPICK2").Compute("SUM(PICK_QTY)", sqlBOL & sqlSC) & "")

                Dim rowSOTSHIP2 As DataRow = dst.Tables("SOTSHIP2").NewRow
                rowSOTSHIP2.Item("SHIP_BOL_NO") = SHIP_BOL_NO
                rowSOTSHIP2.Item("ITEM_CODE") = ITEM_CODE
                rowSOTSHIP2.Item("PICK_QTY") = PICK_QTY
                dst.Tables("SOTSHIP2").Rows.Add(rowSOTSHIP2)
            Next

            ' DCG SAYS TO CALCULATE THE SHIP_VIA HERE - BUT FROM WHAT, A ROUTING INSTRUCTION IN TEXT?

            Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").NewRow
            With rowSOTSHIP1
                .Item("SHIP_BOL_NO") = SHIP_BOL_NO
                .Item("SHIP_VIA_CODE") = SHIP_VIA_CODE
                .Item("SHIP_ADDR_TYPE") = ORDR_ADDR_TYPE_ST
                .Item("SHIP_ADDR_CODE") = SHIP_TO
                .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                .Item("SHIP_STATUS") = "P"
                .Item("TERM_CODE") = TERM_CODE
                .Item("SREP_CODE") = SREP_CODE
                .Item("FRT_TERMS") = FRT_TERMS
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LP_STATUS") = LP_STATUS
                .Item("ORDR_PICK_TYPE") = ORDR_PICK_TYPE
                .Item("SHIP_CART_REQD") = SHIP_CART_REQD
                .Item("ORDR_DEPT") = ORDR_DEPT

                .Item("SHIP_856_IND") = SHIP_856_IND
                .Item("SHIP_810_IND") = SHIP_810_IND
            End With
            dst.Tables("SOTSHIP1").Rows.Add(rowSOTSHIP1)

            MAX_QTY_CTN_pick = MAX_QTY_CTN_cust

            'Override if any combo exceeds max cartons to avoid looping.
            Dim MAX_PICK As Int64 = 0 ' Val(dst.Tables(IIf(ALLOW_SPLIT_TYPE = "S", "SOTSHIP3", "SOTSHIP2")) _
            '  .Compute("MAX(PICK_QTY)", "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'") & "")
            Dim CART_NO_ALL_lno As Integer = 0  ' Last Line Number Used for CART_NO_ALL
            Dim CART_PACK_QTY_ALL As Long       ' CART_PACK_QTY of the ALL Carton
            Dim CART_PACK_QTY As Long           ' Running Balance of Qty in Current Carton

            'CART_NO_seq = 0

            Dim single_carton_shipment As Boolean = False
            If SHIP_856_IND <> "1" And CUST_CODE <> "HSN" Then
                single_carton_shipment = True
            End If

            For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select _
                ("SHIP_BOL_NO = '" & SHIP_BOL_NO & "'", "PICK_NO")
                Dim PICK_NO As String = rowSOTPICK1.Item("PICK_NO")

                Dim CART_NO As String = ""
                Dim CART_NO_ALL As String = ""      ' Carton No to use for ALL Items except for those with specific Carton Requirements

                For Each rowSOTPICK2 As DataRow In dst.Tables("SOTPICK2").Select _
                    ("PICK_NO = '" & PICK_NO & "'", "PICK_LNO")
                    Dim ORDR_NO As String = rowSOTPICK2.Item("ORDR_NO")
                    Dim ORDR_LNO As Int32 = Val(rowSOTPICK2.Item("ORDR_LNO") & "")
                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, ORDR_LNO})
                    Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
                    Dim rowSOTOREL1 As DataRow = dst.Tables("SOTOREL1").Rows.Find(ITEM_CODE) ' LookUp("ICTITEM1", ITEM_CODE)
                    Dim CARTON_PACK_QTY As Int64 = Val(rowSOTOREL1.Item("CARTON_PACK_QTY") & "")
                    ' Dim MAX_QTY_CTN As Int64 = MAX_QTY_CTN_pick ' Maximum Unit Count in a Carton for a Customer
                    QTY_TO_PACK = Val(rowSOTPICK2.Item("PICK_QTY") & "")
                    Dim sqlx As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"

                    ' how does single_carton_shipment work with this flag?
                    If NO_MIXED_CASES = "1" Then
                        CART_NO_ALL = ""
                    End If

                    Do While QTY_TO_PACK <> 0
                        If Not single_carton_shipment And CARTON_PACK_QTY > 0 And QTY_TO_PACK >= CARTON_PACK_QTY Then
                            PACK_QTY = CARTON_PACK_QTY
                            CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                        Else
                            PACK_QTY = QTY_TO_PACK
                            If CART_NO_ALL <> "" Then
                                CART_NO = CART_NO_ALL
                                CART_LNO_seq = CART_NO_ALL_lno
                                CART_PACK_QTY = CART_PACK_QTY_ALL
                                If MAX_QTY_CTN_pick <> 0 And CART_PACK_QTY + PACK_QTY >= MAX_QTY_CTN_pick Then
                                    CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                    CART_NO_ALL = CART_NO
                                End If
                            Else
                                CART_NO = New_Carton(PICK_NO, CART_NO_seq) : CART_LNO_seq = 0 : CART_PACK_QTY = 0
                                CART_NO_ALL = CART_NO
                            End If
                        End If

                        QTY_TO_PACK = QTY_TO_PACK - PACK_QTY
                        CART_PACK_QTY = CART_PACK_QTY + PACK_QTY

                        CART_LNO_seq = CART_LNO_seq + 1
                        Dim rowSOTCART2 As DataRow = dst.Tables("SOTCART2").NewRow
                        With rowSOTCART2
                            .Item("CART_NO") = CART_NO
                            .Item("CART_LNO") = CART_LNO_seq
                            .Item("ORDR_NO") = ORDR_NO
                            .Item("ORDR_LNO") = ORDR_LNO
                            .Item("ITEM_CODE") = ITEM_CODE
                            .Item("QTY_PACKED") = PACK_QTY
                            .Item("ITEM_WEIGHT") = rowSOTOREL1.Item("ITEM_WEIGHT")
                        End With

                        dst.Tables("SOTCART2").Rows.Add(rowSOTCART2)

                        LAST_ITEM_CODE = ITEM_CODE

                        If CART_NO = CART_NO_ALL Then
                            CART_NO_ALL_lno = CART_LNO_seq
                            CART_PACK_QTY_ALL = CART_PACK_QTY
                        End If
                    Loop
                Next
            Next
        Next

        Create_Relation("SOTCART1", "SOTCART2", "CART_NO")
        dst.Tables("SOTCART2").Columns.Add("WGT", GetType(System.Decimal), "ISNULL(QTY_PACKED,0) * ISNULL(ITEM_WEIGHT,0)")
        dst.Tables("SOTCART1").Columns.Add("QTY", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).QTY_PACKED)")
        dst.Tables("SOTCART1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTCART1_SOTCART2).WGT)")
        For Each rowSOTCART1 As DataRow In dst.Tables("SOTCART1").Select("")
            rowSOTCART1.Item("CART_TOTAL_UNITS") = rowSOTCART1.Item("QTY")
            rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = rowSOTCART1.Item("WGT")
        Next

        Create_Relation("SOTPICK1", "SOTCART1", "PICK_NO")
        dst.Tables("SOTPICK1").Columns.Add("CTNS", GetType(System.Int64), "COUNT(CHILD(SOTPICK1_SOTCART1).CART_NO)")
        dst.Tables("SOTPICK1").Columns.Add("WGT", GetType(System.Int64), "SUM(CHILD(SOTPICK1_SOTCART1).CART_TOTAL_WGT_CALC)")
        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("")
            rowSOTPICK1.Item("PICK_CNT_CARTONS") = rowSOTPICK1.Item("CTNS")
            rowSOTPICK1.Item("PICK_TOTAL_WGT") = rowSOTPICK1.Item("WGT")
        Next

        For Each rowSOTPICK1 As DataRow In dst.Tables("SOTPICK1").Select("PICK_CNT_CARTONS = 0")
            rowSOTPICK1.Item("PICK_STATUS") = "C"
        Next

        Create_Relation("SOTSHIP1", "SOTPICK1", "SHIP_BOL_NO")
        dst.Tables("SOTPICK1").Columns.Add("PICKS_C", GetType(System.Int64), "IIF(PICK_STATUS = 'C',1,0)")
        dst.Tables("SOTPICK1").Columns.Add("PICKS_P", GetType(System.Int64), "IIF(PICK_STATUS = 'P',1,0)")
        dst.Tables("SOTSHIP1").Columns.Add("PICKS_C", GetType(System.Int64), "SUM(CHILD(SOTSHIP1_SOTPICK1).PICKS_C)")
        dst.Tables("SOTSHIP1").Columns.Add("PICKS_P", GetType(System.Int64), "SUM(CHILD(SOTSHIP1_SOTPICK1).PICKS_P)")

        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("PICKS_C >0 AND PICKS_P =0")
            rowSOTSHIP1.Item("SHIP_STATUS") = "C"
            WHSE_CODE = rowSOTSHIP1.Item("WHSE_CODE") & ""
            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            If rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                rowSOTSHIP1.Item("LP_STATUS") = "C"
            End If
        Next

    End Sub

    Function New_Carton(PICK_NO As String, ByRef CART_NO_seq As Int32) As String
        Dim rowSOTCART1 As DataRow = dst.Tables("SOTCART1").NewRow
        CART_NO_seq += 1
        Dim CART_NO As String = "TEMP" & Format(CART_NO_seq, "000000")
        rowSOTCART1.Item("CART_NO") = CART_NO
        rowSOTCART1.Item("PICK_NO") = PICK_NO
        rowSOTCART1.Item("CART_TOTAL_UNITS") = 0
        rowSOTCART1.Item("CART_TOTAL_WGT_CALC") = 0
        dst.Tables("SOTCART1").Rows.Add(rowSOTCART1)
        Return CART_NO
    End Function

    Function Check_OOB_Styles() As Boolean
        If 1 = 1 Then Return False
        If Not ASCMAIN1.Running_in_VS Then Return False

        ASCMAIN1.sql = "Select WHSE_CODE, ITEM_CODE" & vbCrLf _
            & ", Sum(ORDR_OPEN) AS ORDR_OPEN, Sum(STAT_OPEN) AS STAT_OPEN" & vbCrLf _
            & ", Sum(ORDR_PICK) AS ORDR_PICK, Sum(STAT_PICK) AS STAT_PICK from (" & vbCrLf _
            & "Select WHSE_CODE, ITEM_CODE" & vbCrLf _
            & ", 0 ORDR_OPEN, Sum (WHSE_QTY_OPEN) STAT_OPEN" & vbCrLf _
            & ", 0 ORDR_PICK, Sum (WHSE_QTY_PICK) STAT_PICK" & vbCrLf _
            & " from ICTSTAT2" & vbCrLf _
            & " group by WHSE_CODE, ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select WHSE_CODE, ITEM_CODE" & vbCrLf _
            & ", Sum (ORDR_QTY_OPEN) ORDR_OPEN, 0 STAT_OPEN" & vbCrLf _
            & ", Sum (ORDR_QTY_PICK) ORDR_PICK, 0 STAT_PICK" & vbCrLf _
            & " from SOTORDR2 where ORDR_STATUS in ('O','P')" & vbCrLf _
            & " group by WHSE_CODE, ITEM_CODE" & vbCrLf _
            & ") group by WHSE_CODE, ITEM_CODE" & vbCrLf _
            & " having Sum(NVL(ORDR_OPEN, 0)) <> SUM(NVL(STAT_OPEN, 0))" _
            & "     or Sum(NVL(ORDR_PICK, 0)) <> SUM(NVL(STAT_PICK, 0))"
        Dim tblSOTROOB1 As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTROOB1", 2)

        If tblSOTROOB1.Rows.Count > 0 Then
            dst.Tables.Add(tblSOTROOB1)
        End If

        Return (tblSOTROOB1.Rows.Count > 0)
    End Function

    'Public Overrides Sub Print_Report()

    '    ' If Absx1.chkFor("CHKCHECK_OOBAL").Checked Then
    '    If OOBAL Then
    '        Using F As New ASFMSGBF
    '            F.Show_grd(dst.Tables("SOTROOB1"), Me, "Items Out of Balance")
    '        End Using
    '        'RPT = "SORROOB1"
    '        'RPT_TITLE = "Item Status Out of Balance Report"
    '        'SUBT = "Please Forward A Copy of This Report to ABS"
    '        'Generate_Report(RPT, RPT_TITLE, SUBT)
    '        Exit Sub
    '    End If
    '    '  End If

    '    ' Get List of Warehouse Codes which had Shipments Released

    '    Dim WHSE_CODEs As New List(Of String)
    '    If SHIP_BOL_NO_seq > 0 Then
    '        For Each row As DataRow In ASCDATA1.SelectDistinct _
    '                 (dst.Tables("SOTSHIP1"), New String() {"WHSE_CODE"}).Select("", "WHSE_CODE")
    '            Dim WHSE_CODE As String = row.Item("WHSE_CODE")
    '            WHSE_CODEs.Add(WHSE_CODE)
    '        Next
    '    End If

    '    ' Determine if any price threshholds have been violated
    '    Dim net_price_below As Boolean = False
    '    Dim D As Integer = Val(ROWs("SOTPARM1").Item("SO_PARM_DISC_THRESH_WARN") & "")
    '    If D <> 0 Then
    '        For Each row As DataRow In ASCMAIN1.Distinct_Values("", "ITEM_SNU_CODE = 'S'", dst.Tables("SOTOREL1"), New String() {"ITEM_CODE"}).Select("")
    '            Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
    '            Dim ORDR_UNIT_PRICE As Decimal = Val(dst.Tables("SOTORDR2").Compute("MIN(ORDR_UNIT_PRICE)", "ITEM_CODE = '" & ITEM_CODE & "'") & "")

    '            Dim rowSOTOREL1 As DataRow = dst.Tables("SOTOREL1").Rows.Find(ITEM_CODE)
    '            Dim ITEM_RETAIL_PRICE As Decimal = Val(rowSOTOREL1.Item("ITEM_RETAIL_PRICE") & "")
    '            If ITEM_RETAIL_PRICE <> 0 Then
    '                If 100 * ORDR_UNIT_PRICE / ITEM_RETAIL_PRICE < D Then
    '                    net_price_below = True
    '                    Exit For
    '                End If
    '            End If
    '        Next
    '    End If


    '    ' Print Reports

    '    Dim rd(11) As String
    '    rd(0) = "Inventory Positions Report"
    '    rd(1) = "Released Orders Report"
    '    rd(2) = "Inventory Picking Requirements Report"
    '    rd(3) = "Inventory Shortage Report"
    '    rd(4) = "Credit Hold Report"
    '    rd(5) = IIf(Absx1.chkFor("CHKALLOCATION_ONLY").Checked, _
    '                "Un-Releasable Orders Report", _
    '                "Orders Not Released Report")
    '    rd(6) = "Item Allocations by Customer Exceeded Report"
    '    rd(7) = "Territory Allotments Exceeded Report"
    '    rd(8) = "Item Allocation Ship Dates > Release Date"
    '    rd(9) = "Marketing Forecast Exceeded Report"
    '    rd(11) = "Net Price below Threshold Report"
    '    For i As Integer = 0 To 11
    '        If (i = 7 Or i = 2 Or i = 8 Or i = 9 Or i = 10) _
    '        Or (i = 11 And Not net_price_below) _
    '        Then
    '            ' SKIP REPORTS
    '        Else
    '            RPT = "SOROREL" & Mid("0123456789AB", i + 1, 1) ' & Format(i, "0")
    '            RPT_TITLE = rd(i)
    '            SUBT = "Batch " & Mid(XNO, 5, 6) _
    '            & IIf(WHSE_CODE = "", "", "  Whse " & WHSE_CODE) _
    '            & "  Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy") _
    '            & IIf(Not Absx1.chkFor("CHKREL_PAST_CANCEL").Checked, " (Ignoring Orders Past Cancel Date)", "") _
    '            & IIf(chkAllocateNoRelease.Checked, " (Reports Only)", "")
    '            If (i = 6) Then
    '                CR_params.Add("SUP_DTLS", "Y")
    '            End If
    '            If i = 11 Then
    '                CR_params.Add("SO_PARM_DISC_THRESH_WARN", Val(ROWs("SOTPARM1").Item("SO_PARM_DISC_THRESH_WARN") & ""))
    '            End If
    '            Generate_Report(RPT, RPT_TITLE, SUBT)
    '        End If
    '    Next i
    'End Sub

    Private Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and COLUMN_NAME <> 'ITEM_CODE' and EXCLUDE = '1'").Length <> 0 Then
            '    EMsg &= vbCr & "You may not use Exclusion on any Filter except Order Group and Item"
            'End If

            'If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and EXCLUDE = '1'").Length <> 0 Then
            '    EMsg &= vbCr & "You may not use Exclusion on any Filter except Order Group"
            'End If

            Dim rowASTDSQLA As DataRow = InArguments.tblASTDSQLA.Rows.Find("ORDR_GROUP_NO")

            If InArguments.ForcePick Then
                If rowASTDSQLA.Item("CODE_VALUES") & "" = "" Then
                    EMsg &= vbCr & "You Must Select Specific Order Groups to Force Pick"
                ElseIf rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                    EMsg &= vbCr & "When Force Picking, you must Select (not Exclude) Order Groups"
                End If
            End If

            If InArguments.ReleasePastCancel Then
                If rowASTDSQLA.Item("CODE_VALUES") & "" = "" Then
                    EMsg &= vbCr & "You Must Select Specific Order Groups to Release Past Cancel"
                ElseIf rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                    EMsg &= vbCr & "When Releasing Past Cancel, you must Select (not Exclude) Order Groups"
                End If
            End If

            If InArguments.WarehouseSelection = "S" Then
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", InArguments.WarehouseCode)
                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Ship-From Warehouse Specified"
                End If
            End If
        End If
    End Sub

    Sub Update_Release()

        Dim PICK_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTPICK0.PICK_BATCH_NO")

        If SHIP_BOL_NO_seq > 0 Then

            For Each TABLE_NAME As String In New String() {"SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2", "SOTORDR1", "SOTORDR2"}
                Update_Record_TDA(TABLE_NAME)
            Next

            ASCDATA1.ExecuteSQL("Update " & SOTPICK1 & " set PICK_BATCH_NO = '" & PICK_BATCH_NO & "'")
            ASCDATA1.ExecuteSQL("Update " & SOTSHIP1 & " set PICK_BATCH_NO = '" & PICK_BATCH_NO & "'")

            For i As Int64 = 1 To SHIP_BOL_NO_seq
                Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
                For Each TABLE_NAME As String In New String() {SOTSHIP1, SOTPICK1}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " set SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                        & " where SHIP_BOL_NO = 'TEMP" & Format$(i, "000000") & "'")
                Next

                ' Needed for sending information to warehouse
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find("TEMP" & Format$(i, "000000"))
                rowSOTSHIP1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
            Next i

            For i As Int64 = 1 To CART_NO_seq
                Dim CART_NO As String = SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
                For Each TABLE_NAME As String In New String() {SOTCART1, SOTCART2}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " set CART_NO = '" & CART_NO & "'" _
                        & " where CART_NO = 'TEMP" & Format$(i, "000000") & "'")
                Next
            Next i

            For i As Int64 = 1 To PICK_NO_seq
                Dim PICK_NO As String = ASCMAIN1.Next_Control_No("SOTPICK1.PICK_NO")
                For Each TABLE_NAME As String In New String() {SOTPICK1, SOTPICK2, SOTCART1}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " set PICK_NO = '" & PICK_NO & "'" _
                        & " where PICK_NO = 'TEMP" & Format(i, "000000") & "'")
                Next
            Next i

            ASCDATA1.ExecuteSQL("Insert into SOTPICK1 Select * from " & SOTPICK1)
            ASCDATA1.ExecuteSQL("Insert into SOTPICK2 Select * from " & SOTPICK2)
            ASCDATA1.ExecuteSQL("Insert into SOTSHIP1 Select * from " & SOTSHIP1)
            ASCDATA1.ExecuteSQL("Insert into SOTCART1 Select * from " & SOTCART1)
            ASCDATA1.ExecuteSQL("Insert into SOTCART2 Select * from " & SOTCART2)

            Dim rowSOTPICK0 As DataRow = dst.Tables("SOTPICK0").NewRow
            With rowSOTPICK0
                .Item("PICK_BATCH_NO") = PICK_BATCH_NO
                .Item("PICK_SHPS") = SHIP_BOL_NO_seq
                .Item("PICK_CTNS") = CART_NO_seq
                .Item("PICK_PKTS") = PICK_NO_seq
                .Item("PICK_BATCH_STATUS") = "O" 'P'?
                .Item("WHSE_CODE") = InArguments.WarehouseCode
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("PICK_SHIP_REL_DATE") = InArguments.ShipByDate
                If InArguments.ForcePick Then
                    .Item("PICK_FORCED") = "1"
                End If
            End With
            dst.Tables("SOTPICK0").Rows.Add(rowSOTPICK0)
            Update_Record_TDA("SOTPICK0")
        End If

        ASCMAIN1.sql = "Update " & SOTORDR2 & " SOTORDR2 Set ORDR_STATUS = " _
            & " (Select ORDR_STATUS from " & SOTORDR1 & " SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select * from " & SOTORDR1 & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SOTORDR1 Set" & vbCrLf _
            & "      ORDR_PICK_SEQ = R1.ORDR_PICK_SEQ" & vbCrLf _
            & "    , ORDR_STATUS = R1.ORDR_STATUS" & vbCrLf _
            & "    , ORDR_DATE_CLOSED = R1.ORDR_DATE_CLOSED, ORDR_YYYYPP_CLOSED = R1.ORDR_YYYYPP_CLOSED, ORDR_DATE_REL = R1.ORDR_DATE_REL" & vbCrLf _
            & "    , ORDR_REL_HOLD_CODES = R1.ORDR_REL_HOLD_CODES" & vbCrLf _
            & "    , ORDR_REL_BATCH_NO = R1.ORDR_REL_BATCH_NO" & vbCrLf _
            & "    where ORDR_NO = R1.ORDR_NO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select * from " & SOTORDR2 & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SOTORDR2 Set" & vbCrLf _
            & "      ORDR_STATUS = R1.ORDR_STATUS" & vbCrLf _
            & "    , ORDR_QTY_OPEN = R1.ORDR_QTY_OPEN" & vbCrLf _
            & "    , ORDR_QTY_PICK = R1.ORDR_QTY_PICK" & vbCrLf _
            & "    , ORDR_QTY_CANC = R1.ORDR_QTY_CANC" & vbCrLf _
            & "    , ORDR_RELEASE = R1.ORDR_RELEASE" & vbCrLf _
            & "    , ALLO_CTL_NO_REL = ALLO_CTL_NO" & vbCrLf _
            & "    where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ' WHAT ABOUT UPDATING SOTRDRVX LIKE WE UPDATE SOTORDRX?

        ' note - we are not doing ICTSTAT2 ALLO - and maybe we should

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is " & vbCrLf _
            & " Select SOTORDR1.WHSE_CODE, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY) PICK_QTY" & vbCrLf _
            & ", Sum (SOTPICK2.PICK_QTY_CANC_REL) PICK_QTY_CANC_REL" & vbCrLf _
            & " from SOTORDR2,SOTORDR1," & SOTPICK2 & " SOTPICK2," & SOTPICK1 & " SOTPICK1 " & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
            & "   and SOTPICK1.PICK_NO = SOTPICK2.PICK_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " group by SOTORDR1.WHSE_CODE, SOTORDR2.ITEM_CODE;" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT2 SET WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) + NVL(R1.PICK_QTY,0), " & vbCrLf _
            & "                        WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN,0) - NVL(R1.PICK_QTY,0) - NVL(R1.PICK_QTY_CANC_REL,0)" & vbCrLf _
            & "   where ITEM_CODE = R1.ITEM_CODE" & vbCrLf _
            & "     and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select Distinct ORDR_GROUP_NO from SOTORDR1" & vbCrLf _
            & "   where ORDR_REL_BATCH_NO = '" & Mid(XNO, 5, 6) & "';" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "    SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)


        ' Only do this if we have updated SOTSHIP1

        ' Do not do this for JCPenney
        ASCMAIN1.sql = "Select Distinct SOTSHIP1X.ORDR_GROUP_NO " _
            & " from " & SOTSHIP1 & " SOTSHIP1X, SOTSHIP1, SOTORDR0 " _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1X.ORDR_GROUP_NO" _
            & "   and SOTSHIP1.SHIP_BOL_NO = SOTSHIP1X.SHIP_BOL_NO" _
            & "   and SOTORDR0.ORDR_SOURCE = 'E'" _
            & "   and SOTORDR0.CUST_CODE in (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '855' and EDI_TP_ID <> '6111355038')"
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select()
            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & ""
            EDC855O1.Generate_855(clsASCBASE1, ORDR_GROUP_NO)
        Next

    End Sub

    Sub Item_Status_Qtys( _
        ITEM_CODE As String, WHSE_CODE As String, _
        ORDR_QTY_OPEN As Int64, ORDR_QTY_PICK As Int64, ORDR_QTY_BACK As Int64, _
        RESTRICT_ORDER_RELEASE As String, MARKET_CODE As String)

        With dst.Tables("SOTOREL2").Rows.Find(New String() {ITEM_CODE, WHSE_CODE})
            .Item("QTY_TO_PICK") = Val(.Item("QTY_TO_PICK") & "") + ORDR_QTY_PICK
            .Item("QTY_TO_DE_COMMIT") = Val(.Item("QTY_TO_DE_COMMIT") & "") + ORDR_QTY_OPEN - ORDR_QTY_BACK
        End With

        If RESTRICT_ORDER_RELEASE = "1" Then
            With dst.Tables("DPTITMFX").Rows.Find(New String() {MARKET_CODE, ITEM_CODE})
                .Item("RELEASED") = Val(.Item("RELEASED") & "") + ORDR_QTY_PICK
            End With
        End If
    End Sub


#End Region

End Class
