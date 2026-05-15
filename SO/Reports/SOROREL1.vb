Public Class SOROREL1

#Region "Declarations"

    Dim PICK_BATCH_NO As String = ""    ' Pick Batch No; this Release will be defined by this control number
    Dim SHIP_BY_DATE As Date    ' Release all orders Scheduled to ship by
    Dim WHSE_CODE As String     ' Ship From Warehouse Code for Released Orders

    Dim SOTORDR0 As String
    Dim SOTORDR1 As String
    Dim SOTORDR2 As String
    Dim SOTORDR5 As String

    Dim SOTPICK1 As String
    Dim SOTPICK2 As String
    Dim SOTSHIP1 As String
    Dim SOTCART1 As String
    Dim SOTCART2 As String

    Dim ARTCUST1 As String
    Dim ARTCUST1_CG As String

    Dim SOTALLOZ As String
    Dim SOTALLOY As String
    Dim SOTALLO1 As String
    Dim SOTOREL1 As String
    Dim SOTOREL2 As String
    Dim POTORDRW As String
    Dim DPTITMFX As String

    Dim PICK_NO_seq As Int64 = 0        ' Temporary Pick Ticket Sequencer
    Dim SHIP_BOL_NO_seq As Int64 = 0    ' Temporary Shipment Sequencer
    Dim CART_NO_seq As Int32 = 0        ' Temporary Carton Sequencer

    Dim OOBAL As Boolean = False
    Dim isIPLBAE As Boolean = False

    Dim MAX_COLs As Integer = 200
    Dim MAX_RSCs As Integer = 500

    Dim sqlICTITEM1 As String
    Dim ICTITEM1 As String
    Dim sqlSOTALLOX As String
    Dim SOTALLOX As String
    Dim sqlSOTALLOW As String
    Dim SOTALLOW As String

    Dim SOTCSTOX As String
    Dim SOTCSTOI As String


#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("POTPARM1")


        Create_Work_Tables(True)


        Set_WHSE()
        Absx1.numFor("FPDCANCEL_FUTURE_DAYS").Value = Val(ROWs("SOTPARM1").Item("SO_PARM_CANCEL_FUTURE_DAYS") & "")
        Absx1.numFor("FPDSHORT_HOR_DAYS").Value = Val(ROWs("SOTPARM1").Item("SO_PARM_SHORT_HOR_DAYS") & "")

        dteSHIP_DATE.CalendarInfo.MaxSelectedDays = 1
        Dim SO_PARM_RELEASE_DAYS_AHEAD As Int64 = Val(ROWs.Item("SOTPARM1").Item("SO_PARM_RELEASE_DAYS_AHEAD") & "")
        dteSHIP_DATE.CalendarInfo.ActivateDay(Now.Date.AddDays(SO_PARM_RELEASE_DAYS_AHEAD))
        dteSHIP_DATE.CalendarInfo.SelectedDateRanges.Add(dteSHIP_DATE.CalendarInfo.ActiveDay.Date)

        If MENU_ITEM_OBJECT = "SORORELG" Then
            Absx1.chkFor("CHKALLOCATION_ONLY").Checked = True
            Absx1.chkFor("CHKALLOCATION_ONLY").Enabled = False
        End If

        If MENU_ITEM_OBJECT = "" Then
            chkAllocateNoRelease.Checked = True
            chkAllocateNoRelease.Enabled = False
        End If

        Set_Date()

    End Sub

    Sub Setup_Allocate_No_Release()
        grpReleaseOptions.Visible = Not chkAllocateNoRelease.Checked
        lblAllUnReleasable.Visible = Not chkAllocateNoRelease.Checked
        numHorizonDays.Visible = chkAllocateNoRelease.Checked
        lblHorizonDays.Visible = chkAllocateNoRelease.Checked
        grpREL_DATE.Visible = Not chkAllocateNoRelease.Checked
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        ' Dim ORDR_GROUP_NO_sql As String = SQLA("ORDR_GROUP_NO", , True)
        SHIP_BY_DATE = dteSHIP_DATE.CalendarInfo.ActiveDay.Date

        ASCMAIN1.Progress("Setting Up Orders (Demand)", "")

        For Each TABLE_NAME As String In New String() {"SOTSREP1", "SOTMKTC1", "ICTBRAN1", "SOTSDIV1"}
            ASCMAIN1.sql = "Select " & TABLE_NAME & ".* from " & TABLE_NAME
            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, TABLE_NAME, 1))
        Next



        Dim CSO_COLs As New List(Of String)
        CSO_COLs.Add("CSO_QTY_TOTAL")

        With dst
            Create_TDA(dst.Tables.Add, "SOTDSCI1", "*", 0)
            Fill_Records("SOTDSCI1")

            Create_TDA(.Tables.Add, "SOTCSTO1", "*", 1)

            Create_TDA(.Tables.Add, "ICTCOLL0", "*", 0, False)
            With .Tables("ICTCOLL0").Columns
                .Add("SEL")
            End With
            .Tables("ICTCOLL0").Columns("SEL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "SOTALLOG", "*", 0, False)
            With .Tables("SOTALLOG").Columns
                .Add("SEL")
            End With
            .Tables("SOTALLOG").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select SOTCSTO2.*, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_SO_QTY_MIN" & vbCrLf _
                & " from SOTCSTO2,ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTCSTO2.ITEM_CODE and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
            Create_TDA(.Tables.Add, "SOTCSTO2", "**", 1)
            With .Tables("SOTCSTO2").Columns
                .Add("ALLO_GROUP_CODE")

                .Add("QTY_ALLO", GetType(System.Int64))
                .Add("ORDR_QTY", GetType(System.Int64))
                .Add("ORDR_QTY_OPEN", GetType(System.Int64))
                .Add("ORDR_QTY_PICK", GetType(System.Int64))
                .Add("ORDR_QTY_SHIP", GetType(System.Int64))
                .Add("ORDR_QTY_CANC", GetType(System.Int64))
                .Add("QTY_LEFT", GetType(System.Int64))


                .Add("WHSE_QTY_ON_HAND", GetType(System.Int64))
                .Add("WHSE_QTY_ONPO", GetType(System.Int64))
                .Add("WHSE_QTY_OPEN", GetType(System.Int64))
                .Add("WHSE_QTY_PICK", GetType(System.Int64))

                .Add("CSO_COL", GetType(System.Int64))

                Dim CT As String = ""
                For c As Integer = 1 To MAX_RSCs
                    Dim CX As String = $"CSO_QTY_{Format(c, "000")}"
                    CT &= $"+ISNULL({CX},0)"
                    .Add(CX, GetType(System.Int64))
                Next
                .Add("CSO_QTY_TOTAL", GetType(System.Int64), Mid(CT, 2))
                .Add("QTY_BAL", GetType(System.Int64), "ISNULL(QTY_LEFT,0) - ISNULL(CSO_QTY_TOTAL,0)")

            End With

            'ASCMAIN1.sql = "Select SOTCSTO3.*" & vbCrLf _
            '    & " from SOTCSTO3" & vbCrLf _
            '    & " where SOTCSTO3.CSO_NO = :PARM1 and SOTCSTO3.ORDR_NO is Not Null"
            'Create_TDA(.Tables.Add, "SOTCSTO3", "**",,, "V", 2)

            ASCMAIN1.sql = "Select SOTCSTO3.*" & vbCrLf _
                & " from SOTCSTO3" & vbCrLf _
                & " where SOTCSTO3.ORDR_NO is Not Null"
            Create_TDA(.Tables.Add, "SOTCSTO3", "**", 1)

            With .Tables("SOTCSTO3").Columns
                .Add("CSO_RSC", GetType(System.Int64))

                Dim CT As String = ""
                For c As Integer = 1 To MAX_COLs
                    Dim CX As String = $"CSO_QTY_{Format(c, "000")}"
                    CT &= $"+ISNULL({CX},0)"
                    .Add(CX, GetType(System.Int64))
                    CSO_COLs.Add($"CSO_QTY_{Format(c, "000")}")
                Next
                .Add("CSO_QTY_TOTAL", GetType(System.Int64), Mid(CT, 2))
            End With


            Create_TDA(.Tables.Add, "SOTCSTO4", "*", 1)

            ASCMAIN1.sql = $"Select * from {ICTITEM1}"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", , False,, 1)

            Create_TDA(.Tables.Add, "ICTWHSE1", "*", 0, False)
            Fill_Records("ICTWHSE1")

            ASCMAIN1.sql = $"Select SOTCSTOX.*, SOTORDR0.ORDR_STATUS from {SOTCSTOX} SOTCSTOX,SOTORDR0" & vbCrLf _
                & " where SOTORDR0.ORDR_GROUP_NO (+) = SOTCSTOX.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTCSTOX", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from " & SOTCSTOI
            Create_TDA(.Tables.Add, "SOTCSTOI", "**", 0, False)


            ASCMAIN1.sql = $"Select SOTALLOX.*" & vbCrLf _
                & ", SOTALLOW.ORDR_QTY" & vbCrLf _
                & ", SOTALLOW.ORDR_QTY_OPEN" & vbCrLf _
                & ", SOTALLOW.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTALLOW.ORDR_QTY_SHIP" & vbCrLf _
                & ", SOTALLOW.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTALLOW.ORDR_SHIP_DATE_OPEN" & vbCrLf _
                & ", SOTALLOW.ORDR_SHIP_DATE_PICK" & vbCrLf _
                & ", SOTALLOW.ORDR_SHIP_DATE_SHIP" & vbCrLf _
                & $"from {SOTALLOX} SOTALLOX, {SOTALLOW} SOTALLOW" & vbCrLf _
                & " where SOTALLOW.ALLO_CTL_NO (+) = SOTALLOX.ALLO_CTL_NO"
            Create_TDA(.Tables.Add, "SOTALLOX", "**", 0, False, "", 1)
            With .Tables("SOTALLOX").Columns
                '.Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)-ISNULL(ORDR_QTY_OPEN,0)")
                '.Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
                .Add("QTY_BAL", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_BAL,0)-ISNULL(ORDR_QTY_OPEN,0)")
            End With

            ASCMAIN1.sql = $"Select * from {SOTALLOW}"
            Create_TDA(.Tables.Add, "SOTALLOW", "**", 0, False, "VDVV", 2)

            Dim tblORDR_HOLD_CODES As New DataTable("SOTORDR_HOLD_CODES")
            With tblORDR_HOLD_CODES
                .Columns.Add("ORDR_NO", GetType(String))
                .Columns.Add("ITEM_CODE", GetType(String))
                .Columns.Add("ORDR_REL_HOLD_CODES_dtl", GetType(String))
                .Columns.Add("ITEM_ORDR_REL_CODE", GetType(String))
            End With
            .Tables.Add(tblORDR_HOLD_CODES)
        End With

        grdSOTCSTO1.DataSource = dst.Tables("SOTCSTO1")
        grdSOTCSTO3.DataSource = dst.Tables("SOTCSTO3")

        With grdSOTCSTO3.DisplayLayout.Bands(0)
            .Columns("CSO_NO").Header.Fixed = True
            .Columns("CSO_ADDR_LNO").Header.Fixed = True
            .Columns("CSO_TYPE").Header.Fixed = True
            .Columns("ORDR_NO").Header.Fixed = True
            .Columns("CUST_NAME").Header.Fixed = True
            .Columns("CSO_QTY_TOTAL").Header.Fixed = True
            .Columns("CSO_QTY_TOTAL").Header.Caption = "Total"
            .Columns("CSO_QTY_TOTAL").Width = 100
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTCSTO3, grdSOTCSTO1}
            With grd.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Next
            End With
        Next

        grdSOTCSTO3.DisplayLayout.Bands(0).Summaries.Clear()
        Create_Summary(grdSOTCSTO3, "CSO_NO", "Count")
        Create_Summary(grdSOTCSTO3, CSO_COLs.ToArray)

        ' Main Process

        PICK_NO_seq = 0
        SHIP_BOL_NO_seq = 0
        CART_NO_seq = 0

        If Order_Release() Then ' will return true whether orders succesfully released or not, and returns false if something bad was detected like OOBAL

            If isIPLBAE Then

                dst.Tables("SOTCSTO1").Rows.Clear()
                dst.Tables("SOTCSTO2").Rows.Clear()
                dst.Tables("SOTCSTO3").Rows.Clear()
                dst.Tables("SOTCSTO4").Rows.Clear()

                For Each row As DataRow In dst.Tables("SOTORDR0").Select("")
                    Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                    ASCMAIN1.sql = $"Select CSO_NO from SOTCSTO1 where ORDR_GROUP_NO = '{ORDR_GROUP_NO}'"
                    Dim rowSOTCSTO1 As DataRow = ASCDATA1.GetDataRow
                    If rowSOTCSTO1 IsNot Nothing Then
                        Dim CSO_NO As String = rowSOTCSTO1.Item("CSO_NO")
                        Load_CSO(CSO_NO)
                    End If
                Next

                'tabCSO.Visible = True
                'tabCSO.Dock = DockStyle.Fill

                'Sort_grdColumns(grdSOTCSTO1, "CSO_NO")

            Else

                'tabCSO.Dock = DockStyle.None
                'tabCSO.Visible = False
            End If

            Update_SOTORDRx()

            ' this is where we should relocate Create_PICK_SHIP_CART because there is temp table ddl in Create_PICK_SHIP_CART
            ' but after examination, the tranaction is restarted after the DDL so because there are no meaningful updates recorded yet, no harm, no foul
            ' so I am leaving things alone for now.
            'If Not chkAllocateNoRelease.Checked Then
            '    Create_PICK_SHIP_CART()
            'End If

            BeginTrans()
            If Not chkAllocateNoRelease.Checked Then
                Create_PICK_SHIP_CART() ' there is DDL that occurs in here, but before any real updates, and it acts as a savepoint with a re-begintrans
                Update_Release() ' Update Pick Ticket, Shipment Control & Carton Tables
            End If
            CommitTrans()
        End If

    End Sub

    Sub Update_SOTORDRx()

        ASCMAIN1.Progress("Now Updating Oracle Tables", "Order Tables")

        For Each TABLE_NAME As String In New String() {"SOTORDR0", "SOTORDR1", "SOTORDR2", "SOTORDR5"}
            Update_Record_TDA(TABLE_NAME)
        Next

        For Each rowORDR_NO As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTORDR5"), "ORDR_NO").Select("")
            Dim ORDR_NO As String = rowORDR_NO.Item("ORDR_NO")
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")
            Dim ORDR_NO_ORIG As String = rowSOTORDR1.Item("ORDR_NO_ORIG")
            Dim CUST_CODE As String = rowSOTORDR1.Item("CUST_CODE")
            ASCMAIN1.Record_Event("SOTORDR1", ORDR_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "BACKORDR", "Back Order Created", ORDR_NO_ORIG)
            ASCMAIN1.Record_Event("SOTORDR1", ORDR_NO_ORIG, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "BACKORIG", "Originating Order for Back Order Created", ORDR_NO)

            If CUST_CODE & "" = "IPLBAE" Then
                Dim rowICTPINVX As DataRow = Fill_Record("ICTPINVX", New String() {ORDR_NO, WHSE_CODE})
                If rowICTPINVX IsNot Nothing AndAlso rowICTPINVX.Item("ETA2DC") & "" <> "" Then
                    rowSOTORDR1.Item("ORDR_SHIP_DATE") = rowICTPINVX.Item("ETA2DC")
                    rowSOTORDR1.Item("ORDR_CANCEL_DATE") = CDate(rowSOTORDR1.Item("ORDR_SHIP_DATE")).AddDays(7)
                Else
                    Dim rowPOTORDRX As DataRow = Fill_Record("POTORDRX", New String() {ORDR_NO, WHSE_CODE})
                    If rowPOTORDRX IsNot Nothing AndAlso rowPOTORDRX.Item("CONF") & "" <> "" Then
                        rowSOTORDR1.Item("ORDR_SHIP_DATE") = rowPOTORDRX.Item("CONF")
                        rowSOTORDR1.Item("ORDR_CANCEL_DATE") = CDate(rowSOTORDR1.Item("ORDR_SHIP_DATE")).AddDays(7)
                    End If
                End If
            End If
        Next
        Update_Record_TDA("SOTORDR1")

        'For Each rowSOTORDR1_AE As DataRow In dst.Tables("SOTORDR1").Select("ORDR_SOURCE = 'S' and ORDR_ROTR = '1' and ORDR_STATUS = 'P'")
        '    Dim ORDR_NO As String = rowSOTORDR1_AE.Item("ORDR_NO")
        '    Dim CUST_STORE_LOCATION As String = rowSOTORDR1_AE.Item("CUST_STORE_LOCATION") & ""
        '    Dim body As String = ""
        '    For Each rowSOTORDR2_AE As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' and ORDR_RELEASE = 'D'")
        '        Dim ITEM_CODE As String = rowSOTORDR2_AE.Item("ITEM_CODE")
        '        Dim ITEM_DESC As String = rowSOTORDR2_AE.Item("ITEM_DESC")
        '        Dim ORDR_QTY As Integer = Val(rowSOTORDR2_AE.Item("ORDR_QTY") & "")
        '        body = $"The following item has been deleted from your Car Stock Order for {CUST_STORE_LOCATION}"
        '        body &= $"<br/>Qty {Format(ORDR_QTY, "#,##0")} {vbTab}, Item {ITEM_CODE} {ITEM_DESC})"
        '    Next
        '    If body <> "" And chkAllocateNoRelease.Checked = False Then
        '        email_to_AE(ORDR_NO, body)
        '    End If

        'Next

        Dim ORDR_GROUP_NOs As New List(Of String)()
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("ORDR_STATUS = 'P'")
            Dim ORDR_GROUP_NO As String = rowSOTORDR1.Item("ORDR_GROUP_NO").ToString()

            If ORDR_GROUP_NOs.Contains(ORDR_GROUP_NO) Then Continue For
            ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)

            Dim ORDR_CUST_PO As String = rowSOTORDR1.Item("ORDR_CUST_PO").ToString()
            Dim CUST_STORE_LOCATION As String = rowSOTORDR1.Item("CUST_STORE_LOCATION").ToString()
            Dim itemsList As New List(Of String)()
            Dim hasCuts As Boolean = False

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_GROUP_NO = '{ORDR_GROUP_NO}' AND (ORDR_QTY_CANC > 0 OR ORDR_RELEASE = 'D')")
                Dim ORDR_NO As String = rowSOTORDR2.Item("ORDR_NO") & ""
                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE") & ""
                Dim CUST_CODE As String = rowSOTORDR2.Item("CUST_CODE") & ""
                Dim ITEM_DESC As String = rowSOTORDR2.Item("ITEM_DESC") & ""
                Dim ORDR_QTY As Integer = Val(rowSOTORDR2.Item("ORDR_QTY") & "")
                Dim ORDR_QTY_CANC As Integer = Val(rowSOTORDR2.Item("ORDR_QTY_CANC") & "")
                Dim CUST_STORE_NO As String = rowSOTORDR2.Item("CUST_STORE_NO") & ""
                ASCMAIN1.sql = $"SELECT CUST_STORE_NAME FROM ARTCUST2 WHERE CUST_STORE_NO = '{CUST_STORE_NO}' AND CUST_CODE = '{CUST_CODE}'"
                Dim CUST_STORE_NAME As String = ASCDATA1.GetDataValue(ASCMAIN1.sql) & ""
                Dim rowSOTDSCI1 As DataRow = dst.Tables("SOTDSCI1").Rows.Find(New String() {ITEM_CODE, CUST_CODE})
                Dim rowORDR_REL_HOLD_CODE As DataRow() = dst.Tables("SOTORDR_HOLD_CODES").Select($"ORDR_NO = '{ORDR_NO}' AND ITEM_CODE = '{ITEM_CODE}'")
                Dim ORDR_REL_HOLD_CODE As String = If(rowORDR_REL_HOLD_CODE.Length > 0, rowORDR_REL_HOLD_CODE(0).Item("ORDR_REL_HOLD_CODES_dtl").ToString(), "")
                Dim ITEM_ORDR_REL_CODE As String = rowORDR_REL_HOLD_CODE(0).Item("ITEM_ORDR_REL_CODE") & ""

                Dim REASON As String = ""
                If ORDR_REL_HOLD_CODE = "N" Then
                    REASON = "No lines would have released"
                ElseIf ITEM_ORDR_REL_CODE = "D" AndAlso rowSOTDSCI1 Is Nothing Then
                    REASON = "Item not excluded for customer (Rel Code D)"
                ElseIf ORDR_QTY_CANC > 0 Then
                    REASON = If(ORDR_QTY = ORDR_QTY_CANC, "Cancelled - All units cancelled", "Cancelled - Some units cancelled")
                End If

                Dim displayQty As Integer = If(ITEM_ORDR_REL_CODE = "D", ORDR_QTY, ORDR_QTY_CANC)

                itemsList.Add($"Order No: {ORDR_NO} | Qty {displayQty} | Item {ITEM_CODE} - {ITEM_DESC} | Store: {CUST_STORE_NAME} (Store No: {CUST_STORE_NO}){If(REASON <> "", $" | Reason: {REASON}", "")}")
                hasCuts = True
            Next

            If hasCuts AndAlso Not chkAllocateNoRelease.Checked Then
                Dim body As String = $"The following items have been deleted from your Purchase Order (PO: {ORDR_CUST_PO}):<br/><br/>{String.Join("<br/>", itemsList)}"
                email_PO_Cuts(ORDR_CUST_PO, body)
            End If
        Next


        ASCMAIN1.Progress("Now Updating Oracle Tables", "Holds Audit D")

        ASCMAIN1.sql = $"INSERT INTO SOTORELD (XNO, ORDR_NO, REPORT_DATE, CUST_CODE, ORDR_CUST_PO, ORDR_REL_HOLD_CODES)
            SELECT '{XNO}' AS XNO, ORDR_NO,  SYSDATE AS REPORT_DATE, CUST_CODE, ORDR_CUST_PO, ORDR_REL_HOLD_CODES
            FROM {SOTORDR1} SOTORDR1_TEMP 
            WHERE SOTORDR1_TEMP.ORDR_REL_HOLD_CODES IS NOT NULL"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Now Updating Oracle Tables", "Holds Audit C")
        ASCMAIN1.sql = $"INSERT INTO SOTORELC
                     SELECT XNO,
                            SYSDATE AS REPORT_DATE,
                            REPLACE(LTRIM(SYS_CONNECT_BY_PATH(ORDR_REL_HOLD_CODE, ','), ','), ',', '') AS ORDR_REL_HOLD_CODES,
                        NULL AS REPORT_NO
                     FROM (
                           SELECT XNO,
                                  ORDR_REL_HOLD_CODE,
                                  ROW_NUMBER() OVER (PARTITION BY XNO ORDER BY ORDR_REL_HOLD_CODE) AS rn,
                                  ROW_NUMBER() OVER (PARTITION BY XNO ORDER BY ORDR_REL_HOLD_CODE DESC) AS rn_desc
                           FROM (
                                 {String.Join(" UNION ", Enumerable.Range(1, 20).Select(Function(i) $"SELECT DISTINCT XNO, SUBSTR(ORDR_REL_HOLD_CODES, {i}, 1) AS ORDR_REL_HOLD_CODE FROM SOTORELD WHERE XNO = '{XNO}'" & vbCrLf))}
                                ) 
                           WHERE ORDR_REL_HOLD_CODE IS NOT NULL
                           ORDER BY ORDR_REL_HOLD_CODE
                          )
                     WHERE rn_desc = 1
                     START WITH rn = 1
                     CONNECT BY PRIOR XNO = XNO AND PRIOR rn = rn - 1"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Now Updating Oracle Tables", "Hold Codes")

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
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Now Updating Oracle Tables", "Re-Order Memo")

        ASCMAIN1.sql = "Update " & SOTORDR1 & " Set REORD_MEMO_IND = '1'" _
            & " where ORDR_NO in (" _
            & " Select Distinct ORDR_NO from " & SOTORDR2 _
            & " where ORDR_RELEASE = 'R')"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.Progress("Now Updating Oracle Tables", "Order Details")

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
        'ASCDATA1.ExecuteSQL()
        ' MAYBE UPDATE THESE ALSO?
        '& "    , ORDR_RELEASE = R1.ORDR_RELEASE" _
        '& "    , ORDR_RELEASE_AVAIL = R1.ORDR_RELEASE_AVAIL" _

        ASCMAIN1.sql = ASCMAIN1.Flattened_List("TABLE_NAME", "COLUMN_NAME", "USER_TAB_COLUMNS", ",", "TABLE_NAME = 'SOTORDR1'")
        Dim row1 As DataRow = ASCDATA1.GetDataRow
        Dim COLUMN_NAMES_SOTORDR1 As String = row1.Item("COLUMN_NAMES")
        COLUMN_NAMES_SOTORDR1 = ""
        ASCMAIN1.sql = "Select TABLE_NAME, COLUMN_NAME, COLUMN_ID from USER_TAB_COLUMNS where TABLE_NAME = 'SOTORDR1'"
        For Each row1 In ASCDATA1.GetDataTable().Select("", "COLUMN_ID")
            COLUMN_NAMES_SOTORDR1 &= "," & row1.Item("COLUMN_NAME")
        Next
        COLUMN_NAMES_SOTORDR1 = Mid(COLUMN_NAMES_SOTORDR1, 2)
        ASCDATA1.ExecuteSQL($"Insert into SOTORDR1 Select {COLUMN_NAMES_SOTORDR1} from {SOTORDR1} where ORDR_NO in (Select Distinct ORDR_NO from {SOTORDR5})")

        ASCMAIN1.sql = ASCMAIN1.Flattened_List("TABLE_NAME", "COLUMN_NAME", "USER_TAB_COLUMNS", ",", "TABLE_NAME = 'SOTORDR2'")
        Dim row2 As DataRow = ASCDATA1.GetDataRow
        Dim COLUMN_NAMES_SOTORDR2 As String = row2.Item("COLUMN_NAMES")
        COLUMN_NAMES_SOTORDR2 = ""
        ASCMAIN1.sql = "Select TABLE_NAME, COLUMN_NAME, COLUMN_ID from USER_TAB_COLUMNS where TABLE_NAME = 'SOTORDR2'"
        For Each row2 In ASCDATA1.GetDataTable().Select("", "COLUMN_ID")
            COLUMN_NAMES_SOTORDR2 &= "," & row2.Item("COLUMN_NAME")
        Next
        COLUMN_NAMES_SOTORDR2 = Mid(COLUMN_NAMES_SOTORDR2, 2)
        ASCDATA1.ExecuteSQL($"Insert into SOTORDR2 Select {COLUMN_NAMES_SOTORDR2} from {SOTORDR2} where ORDR_NO in (Select Distinct ORDR_NO from {SOTORDR5})")

        ASCDATA1.ExecuteSQL($"Insert into SOTORDR5 Select * from {SOTORDR5} where ORDR_NO in (Select Distinct ORDR_NO from {SOTORDR5})")

        ASCMAIN1.Progress("")

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

            If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" AndAlso (ORDR_NO = "0009718178" Or ORDR_NO = "0009718192") Then Stop
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


        Dim by_ORDR_NO As String = ", 'X'"
        If ASCMAIN1.CLIENT = "INT" Then
            ' MAYBE ADD ADD MCX AND NEXCOM - BUT ONLY OF SHIPPING UPS - AND WE DON'T KNOW THAT UNTIL CONF COMES BACK FROM CLARINS
            'ASCMAIN1.sql &= vbCrLf & ", DECODE(SOTORDR1.CUST_CODE,'IPLBAE',SOTORDR1.ORDR_NO,NULL)"
            by_ORDR_NO = ", DECODE(SOTORDR1.CUST_CODE,'IPLBAE',DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST,'DC',NULL,SOTORDR1.ORDR_NO),NULL)"
        End If


        ASCMAIN1.sql = "Select SOTORDR1.ORDR_GROUP_NO, SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE, SOTORDR1.CUST_CODE" & vbCrLf _
            & ", SOTORDR1.ORDR_ADDR_TYPE_ST" & vbCrLf _
            & ", DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC', SOTORDR1.CUST_DC_NO, 'MA', DECODE(SOTORDR1.CUST_CODE, 'MCX', SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO), SOTORDR1.CUST_STORE_NO) SHIP_TO" & vbCrLf _
            & by_ORDR_NO & " ORDR_NO" & vbCrLf _
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
            & ", DECODE(SOTORDR1.ORDR_ADDR_TYPE_ST, 'DC', SOTORDR1.CUST_DC_NO, 'MA', DECODE(SOTORDR1.CUST_CODE, 'MCX', SOTORDR1.CUST_DC_NO, SOTORDR1.CUST_STORE_NO), SOTORDR1.CUST_STORE_NO)" & vbCrLf _
            & by_ORDR_NO



        ASCMAIN1.sql = "Select X.*, EDT850T1.CUST_CODE_OVERRIDE, EDT850T1.CUST_CODE CUST_CODE_ORIGINAL" & vbCrLf _
            & " from EDT850T1, (" & ASCMAIN1.sql & ") X where EDT850T1.EDI_DOC_SEQ_NO (+) = X.EDI_DOC_SEQ_NO"


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
        For Each rowSHIPMENT As DataRow In ASCDATA1.GetDataTable.Select("", "WHSE_CODE,CUST_CODE")
            Dim ORDR_ADDR_TYPE_ST As String = rowSHIPMENT.Item("ORDR_ADDR_TYPE_ST")
            Dim ORDR_GROUP_NO As String = rowSHIPMENT.Item("ORDR_GROUP_NO")
            Dim EDI_DOC_SEQ_NO As String = rowSHIPMENT.Item("EDI_DOC_SEQ_NO") & ""

            Dim CUST_CODE_OVERRIDE As String = ""
            Dim CUST_CODE_ORIGINAL As String = ""

            CUST_CODE_OVERRIDE = rowSHIPMENT.Item("CUST_CODE_OVERRIDE") & ""
            CUST_CODE_ORIGINAL = rowSHIPMENT.Item("CUST_CODE_ORIGINAL") & ""


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

                If CUST_CODE_OVERRIDE <> "" Then
                    rowEDTSLSP1 = LookUp("EDTSLSP1", CUST_CODE_ORIGINAL)
                End If



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

            ' OrElse add 9/15/2022
            If ORDR_ADDR_TYPE_ST = "DC" OrElse (CUST_CODE = "MCX" AndAlso ORDR_ADDR_TYPE_ST = "MA") Then
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

            If ASCMAIN1.CLIENT = "INT" Then
                Dim ORDR_NO As String = rowSHIPMENT.Item("ORDR_NO") & ""
                If ORDR_NO <> "" And ORDR_NO <> "X" Then
                    sqlw &= " and ORDR_NO = '" & ORDR_NO & "'"
                End If
            End If

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select(sqlw)
                Dim rowSOTPICK1 As DataRow = rowSOTORDR1.GetChildRows("SOTORDR1_SOTPICK1")(0)
                rowSOTPICK1.Item("SHIP_BOL_NO") = SHIP_BOL_NO
            Next

            Dim sqlBOL As String = "SHIP_BOL_NO = '" & SHIP_BOL_NO & "'"

            For Each rowSC As DataRow In ASCDATA1.SelectDistinct _
                (dst.Tables("SOTPICK2").Select(sqlBOL),
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
            If rowSOTPICK1.Item("SHIP_BOL_NO") = "X" Then
                rowSOTPICK1.Delete()
            Else
                rowSOTPICK1.Item("PICK_STATUS") = "C"
            End If

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

        ' code to demonstrate that DDL within a BeginTans actually commits to a savepoint and restarts the transaction
        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        '    ASCMAIN1.sql = "UPDATE ICTSTAT2 SET WHSE_QTY_ON_HAND = WHSE_QTY_ON_HAND - 1000  WHERE ITEM_CODE = 'JLB306010SF'"
        '    ASCDATA1.ExecuteSQL()
        'End If


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
        ASCMAIN1.Progress("Now Checking Inventory Status Integrity", "")

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

    Public Overrides Sub Print_Report()

        ' If Absx1.chkFor("CHKCHECK_OOBAL").Checked Then
        If OOBAL Then
            Using F As New ASFMSGBF
                F.Show_grd(dst.Tables("SOTROOB1"), Me, "Items Out of Balance")
            End Using
            'RPT = "SORROOB1"
            'RPT_TITLE = "Item Status Out of Balance Report"
            'SUBT = "Please Forward A Copy of This Report to ABS"
            'Generate_Report(RPT, RPT_TITLE, SUBT)
            Exit Sub
        End If
        '  End If

        ' Get List of Warehouse Codes which had Shipments Released

        Dim WHSE_CODEs As New List(Of String)
        If SHIP_BOL_NO_seq > 0 Then
            For Each row As DataRow In ASCDATA1.SelectDistinct _
                     (dst.Tables("SOTSHIP1"), New String() {"WHSE_CODE"}).Select("", "WHSE_CODE")
                Dim WHSE_CODE As String = row.Item("WHSE_CODE")
                WHSE_CODEs.Add(WHSE_CODE)
            Next
        End If

        ' Determine if any price threshholds have been violated
        Dim net_price_below As Boolean = False
        Dim D As Integer = Val(ROWs("SOTPARM1").Item("SO_PARM_DISC_THRESH_WARN") & "")
        If D <> 0 Then
            For Each row As DataRow In ASCMAIN1.Distinct_Values("", "ITEM_SNU_CODE = 'S'", dst.Tables("SOTOREL1"), New String() {"ITEM_CODE"}).Select("")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
                Dim ORDR_UNIT_PRICE As Decimal = Val(dst.Tables("SOTORDR2").Compute("MIN(ORDR_UNIT_PRICE)", "ITEM_CODE = '" & ITEM_CODE & "'") & "")

                Dim rowSOTOREL1 As DataRow = dst.Tables("SOTOREL1").Rows.Find(ITEM_CODE)
                Dim ITEM_RETAIL_PRICE As Decimal = Val(rowSOTOREL1.Item("ITEM_RETAIL_PRICE") & "")
                If ITEM_RETAIL_PRICE <> 0 Then
                    If 100 * ORDR_UNIT_PRICE / ITEM_RETAIL_PRICE < D Then
                        net_price_below = True
                        Exit For
                    End If
                End If
            Next
        End If


        ' Print Reports
        Dim REPORT_NO As String = String.Empty
        Dim rd(11) As String
        rd(0) = "Inventory Positions Report"
        rd(1) = "Released Orders Report"
        rd(2) = "Inventory Picking Requirements Report"
        rd(3) = "Inventory Shortage Report"
        rd(4) = "Credit Hold Report"
        rd(5) = IIf(Absx1.chkFor("CHKALLOCATION_ONLY").Checked,
                    "Un-Releasable Orders Report",
                    "Orders Not Released Report")
        rd(6) = "Item Allocations by Customer Exceeded Report"
        rd(7) = "Territory Allotments Exceeded Report"
        rd(8) = "Item Allocation Ship Dates > Release Date"
        rd(9) = "Marketing Forecast Exceeded Report"
        rd(11) = "Net Price below Threshold Report"
        For i As Integer = 0 To 11
            If (i = 7 Or i = 2 Or i = 8 Or i = 9 Or i = 10) _
            Or (i = 11 And Not net_price_below) _
            Then
                ' SKIP REPORTS
            Else
                RPT = "SOROREL" & Mid("0123456789AB", i + 1, 1) ' & Format(i, "0")
                RPT_TITLE = rd(i)
                SUBT = "Batch " & Mid(XNO, 5, 6) _
                & IIf(WHSE_CODE = "", "", "  Whse " & WHSE_CODE) _
                & "  Release Date " & Format(SHIP_BY_DATE, "MM/dd/yy") _
                & IIf(Not Absx1.chkFor("CHKREL_PAST_CANCEL").Checked, " (Ignoring Orders Past Cancel Date)", "") _
                & IIf(chkAllocateNoRelease.Checked, " (Reports Only)", "")
                If (i = 6) Then
                    CR_params.Add("SUP_DTLS", "Y")
                End If
                If i = 11 Then
                    CR_params.Add("SO_PARM_DISC_THRESH_WARN", Val(ROWs("SOTPARM1").Item("SO_PARM_DISC_THRESH_WARN") & ""))
                End If
                If RPT = "SOROREL5" Then
                    REPORT_NO = Generate_Report(RPT, RPT_TITLE, SUBT)
                Else
                    Generate_Report(RPT, RPT_TITLE, SUBT)
                End If
            End If
        Next i
        ASCMAIN1.sql = $"UPDATE SOTORELC SET REPORT_NO = '{REPORT_NO}' WHERE XNO = '{XNO}'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            'If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and COLUMN_NAME <> 'ITEM_CODE' and EXCLUDE = '1'").Length <> 0 Then
            '    EMsg &= vbCr & "You may not use Exclusion on any Filter except Order Group and Item"
            'End If

            'If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' and EXCLUDE = '1'").Length <> 0 Then
            '    EMsg &= vbCr & "You may not use Exclusion on any Filter except Order Group"
            'End If

            Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("ORDR_GROUP_NO")

            If Absx1.chkFor("CHKFORCE_PICK").Checked Then
                If rowASTDSQLA.Item("CODE_VALUES") & "" = "" Then
                    EMsg &= vbCr & "You Must Select Specific Order Groups to Force Pick"
                ElseIf rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                    EMsg &= vbCr & "When Force Picking, you must Select (not Exclude) Order Groups"
                End If
            End If

            If Absx1.chkFor("CHKREL_PAST_CANCEL").Checked Then
                If rowASTDSQLA.Item("CODE_VALUES") & "" = "" Then
                    EMsg &= vbCr & "You Must Select Specific Order Groups to Release Past Cancel"
                ElseIf rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                    EMsg &= vbCr & "When Releasing Past Cancel, you must Select (not Exclude) Order Groups"
                End If
            End If

            If optWHSE.Value = "S" Then
                Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                If rowICTWHSE1 Is Nothing Then
                    EMsg &= vbCr & "Invalid Ship-From Warehouse Specified"
                End If
            End If

            Dim rowASTDSQLA_CUST_CODE As DataRow = tblASTDSQLA.Rows.Find("CUST_CODE")
            Dim CUST_CODEs_list() As String = {}
            If rowASTDSQLA_CUST_CODE.Item("CODE_VALUES") & "" <> "" Then
                CUST_CODEs_list = Split(rowASTDSQLA_CUST_CODE.Item("CODE_VALUES") & "", ",")
            End If
            'Dim CUST_CODEs As New List(Of String)


            Dim CUST_CODEs_from_ORDR_GROUP_NOs As New List(Of String)
            If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                Dim ORDR_GROUP_CODEs As String = $"'{Replace(rowASTDSQLA.Item("CODE_VALUES"), ",", "','")}'"
                ASCMAIN1.sql = $"Select Distinct CUST_CODE from SOTORDR0 where ORDR_GROUP_NO IN ({ORDR_GROUP_CODEs})"
                For Each row As DataRow In ASCDATA1.GetDataTable().Select()
                    Dim CUST_CODE_sel As String = row.Item("CUST_CODE")
                    CUST_CODEs_from_ORDR_GROUP_NOs.Add(CUST_CODE_sel)
                Next
            End If


            isIPLBAE = False

            'If ASCMAIN1.Running_in_VS And False AndAlso ASCMAIN1.USER_ID = "wjz" Then

            If CUST_CODEs_list.Length > 0 Then
                    If CUST_CODEs_list.Contains("IPLBAE") Then
                        If CUST_CODEs_list.Length > 1 Then
                            EMsg &= vbCr & "Orders for IPLBAE must be the only Customer in the release batch"
                        Else
                            If CUST_CODEs_from_ORDR_GROUP_NOs.Count > 1 Or (CUST_CODEs_from_ORDR_GROUP_NOs.Count = 1 And CUST_CODEs_from_ORDR_GROUP_NOs(0) <> "IPLBAE") Then
                                EMsg &= vbCr & "Customer IPLBAE was chosen, but non-IPLBAE Groups were chosen"
                            Else
                                isIPLBAE = True
                            End If
                        End If
                    End If
                End If

            If CUST_CODEs_from_ORDR_GROUP_NOs.Count > 0 Then
                If CUST_CODEs_from_ORDR_GROUP_NOs.Contains("IPLBAE") Then
                    If CUST_CODEs_from_ORDR_GROUP_NOs.Count > 1 Then
                        EMsg &= vbCr & "Orders for IPLBAE must be released as the only customer in the release batch"
                    End If
                Else ' IPLBAE was not chosen as a customer

                End If

                'If CUST_CODEs_from_ORDR_GROUP_NOs.Contains("IPLBAE") And CUST_CODEs_from_ORDR_GROUP_NOs.Count > 1 Then

                'Else
                '    If CUST_CODEs_list.Length > 1 Or (CUST_CODEs_list.Length = 1 AndAlso CUST_CODEs_list(0) <> "IPLBAE") Then
                '        EMsg &= vbCr & "An IPLBAE Group was chosen, but non-IPLBAE Customers were chosen"
                '    Else
                '        isIPLBAE = True
                '    End If
                'End If
            End If

            If isIPLBAE And EMsg = "" Then
                    If tblASTDSQLA.Select("COLUMN_NAME <> 'ORDR_GROUP_NO' AND COLUMN_NAME <> 'CUST_CODE' AND (ISNULL(CODE_VALUES,'?') <> '?' and CODE_VALUES <> '')").Length <> 0 Then
                        EMsg &= vbCr & "IPLBAE Customer or Group was chosen - no other filters are permitted when releasing IPLBAE"
                    End If
                End If
            'End If

            'EMsg &= vbCr & "PLACEHOLDER"
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Sub Update_Release()

        Dim PICK_BATCH_NO As String = ASCMAIN1.Next_Control_No("SOTPICK0.PICK_BATCH_NO")

        If SHIP_BOL_NO_seq > 0 Then

            For Each TABLE_NAME As String In New String() {"SOTPICK1", "SOTPICK2", "SOTSHIP1", "SOTCART1", "SOTCART2", "SOTORDR1", "SOTORDR2"}
                Update_Record_TDA(TABLE_NAME)
            Next

            ASCMAIN1.Progress("Updating Pick Ticket & Shipment Tables", "")

            ASCDATA1.ExecuteSQL("Update " & SOTPICK1 & " set PICK_BATCH_NO = '" & PICK_BATCH_NO & "'")
            ASCDATA1.ExecuteSQL("Update " & SOTSHIP1 & " set PICK_BATCH_NO = '" & PICK_BATCH_NO & "'")

            For i As Int64 = 1 To SHIP_BOL_NO_seq
                Dim SHIP_BOL_NO As String = ASCMAIN1.Next_Control_No("SOTSHIP1.SHIP_BOL_NO")
                For Each TABLE_NAME As String In New String() {SOTSHIP1, SOTPICK1}
                    ASCDATA1.ExecuteSQL("Update " & TABLE_NAME & " set SHIP_BOL_NO = '" & SHIP_BOL_NO & "'" _
                        & " where SHIP_BOL_NO = 'TEMP" & Format$(i, "000000") & "'")
                Next
            Next i

            For i As Int64 = 1 To CART_NO_seq
                Dim CART_NO As String = TAC.SOCMAIN1.UPC(Me, ASCMAIN1.Next_Control_No("SOTCART1.CART_NO"), "0000" & ROWs("SOTPARM1").Item("SO_PARM_UPC_VENDOR_ID"))
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
                .Item("WHSE_CODE") = WHSE_CODE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("PICK_SHIP_REL_DATE") = SHIP_BY_DATE
                If Absx1.chkFor("CHKFORCE_PICK").Checked Then
                    .Item("PICK_FORCED") = "1"
                End If
            End With
            dst.Tables("SOTPICK0").Rows.Add(rowSOTPICK0)
            Update_Record_TDA("SOTPICK0")
        End If

        ASCMAIN1.Progress("Updating Order Tables", "")

        ASCMAIN1.sql = "Update " & SOTORDR2 & " SOTORDR2 Set ORDR_STATUS = " _
            & " (Select ORDR_STATUS from " & SOTORDR1 & " SOTORDR1 where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO)"
        ASCDATA1.ExecuteSQL()

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
        ASCDATA1.ExecuteSQL()

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
        ASCDATA1.ExecuteSQL()

        ' WHAT ABOUT UPDATING SOTRDRVX LIKE WE UPDATE SOTORDRX?

        ASCMAIN1.Progress("Updating Item Status Tables", "")

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
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("Updating Order Group Summary", "")

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
        ASCDATA1.ExecuteSQL()


        ' Only do this if we have updated SOTSHIP1
        ASCMAIN1.Progress("Updating 855's", "")

        ' Do not do this for JCPenney
        ASCMAIN1.sql = "Select Distinct SOTSHIP1X.ORDR_GROUP_NO " _
            & " from " & SOTSHIP1 & " SOTSHIP1X, SOTSHIP1, SOTORDR0 " _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1X.ORDR_GROUP_NO" _
            & "   and SOTSHIP1.SHIP_BOL_NO = SOTSHIP1X.SHIP_BOL_NO" _
            & "   and SOTORDR0.ORDR_SOURCE = 'E'" _
            & "   and SOTORDR0.CUST_CODE in" _
            & " (" _
            & " select edtslsp1.CUST_CODE " _
            & " from edtslsp1,  edttrpm1" _
            & " where edtslsp1.EDI_ID_850 = edttrpm1.EDI_TP_ID" _
            & " And edtslsp1.EDI_QUAL_850 = edttrpm1.EDI_TP_QUAL" _
            & " and edttrpm1.EDI_DOC_NO = '855'" _
            & " and edttrpm1.EDI_TP_ID <> '6111355038'" _
            & ")"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & ""
            Dim CUST_CODE As String = dst.Tables("SOTORDR1").Select($"ORDR_GROUP_NO = '{ORDR_GROUP_NO}'", "")(0).Item("CUST_CODE") & String.Empty
            ' No longer needed for Amazon
            'Dim WHSE_CODE As String = ASCDATA1.GetDataValue($"SELECT WHSE_CODE FROM {SOTSHIP1} WHERE ORDR_GROUP_NO = :PARM1", "V", {ORDR_GROUP_NO})

            ' ISSUE-7373 - remove prohibition
            'Select Case CUST_CODE
            '    Case "AMAZON"
            '        ' Amazon may be shipped across multiple warehouses. 
            '        ' Amazon accepts only one 855 per P.O.
            '        ' We need to send the 855 after all warehouse orders are released
            '        ' Possible issues - Release to one Warehouse then cancel the orders for the other warehouse after releasing the first warehouse
            '        ASCMAIN1.sql = "SELECT DISTINCT ORDR_GROUP_NO, WHSE_CODE, ORDR_STATUS FROM SOTORDR1 WHERE ORDR_GROUP_NO = :PARM1"
            '        Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "V", {ORDR_GROUP_NO})
            '        If tbl.Select("ORDR_STATUS = 'O'").Length > 0 Then
            '            Continue For
            '        End If

            '    Case Else

            'End Select

            TAC.EDC855O1.Generate_855(clsASCBASE1, ORDR_GROUP_NO, String.Empty)
        Next

    End Sub

    Private Sub chkAllocateNoRelease_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkAllocateNoRelease.CheckedChanged
        Setup_Allocate_No_Release()
    End Sub

    Private Sub dteSHIP_DATE_AfterPerformAction(sender As Object, e As Infragistics.Win.UltraWinSchedule.AfterMonthViewMultiPerformActionEventArgs) Handles dteSHIP_DATE.AfterPerformAction
        Set_Date()
    End Sub

    Sub Set_Date()
        lblSHIP_DATE.Text = Format(dteSHIP_DATE.CalendarInfo.SelectedDateRanges(0).StartDate, "MM/dd/yy")
        lblAllUnReleasable.Text = "All Un-Releasable Orders will be Shown (with Ship Date On or Before " & lblSHIP_DATE.Text & ")"
    End Sub

    Private Sub dteSHIP_DATE_MouseUp(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dteSHIP_DATE.MouseUp
        Set_Date()
    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "ORDR_GROUP_NO"
                sqlw = " SOTORDR0.ORDR_CNT_OPEN > 0"
        End Select
        Return sqlw
    End Function

    Private Sub optWHSE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optWHSE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_WHSE()
    End Sub

    Sub Set_WHSE()
        If optWHSE.Value = "A" Then
            Absx1.txtFor("WHSE_CODE").Text = ""
            Absx1.txtFor("WHSE_CODE").Enabled = False
        Else
            Absx1.txtFor("WHSE_CODE").Text = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""
            Absx1.txtFor("WHSE_CODE").Enabled = True
        End If
    End Sub

    Function Order_Release() As Boolean

        OOBAL = Check_OOB_Styles()
        If OOBAL Then
            Return False
        End If

        ' Prepare Order Header & Detail Work Tables in Oracle

        ASCMAIN1.Progress("Setting Up Order Tables")

        ASCMAIN1.sql = "" _
            & "Select SOTORDR1.*, ARTCUST1.CUST_PRIORITY_CODE, EDT850T1.EDI_RECEIVED_DATE" & vbCrLf _
            & " from SOTORDR1, ARTCUST1, EDT850T1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
            & " and EDT850T1.EDI_DOC_SEQ_NO (+) = SOTORDR1.EDI_DOC_SEQ_NO" & vbCrLf _
            & " and SOTORDR1.ORDR_STATUS = 'O'" & vbCrLf _
            & " and (SOTORDR1.ORDR_SHIP_DATE <= '" & Format(SHIP_BY_DATE, "dd-MMM-yyyy") & "' or SOTORDR1.ORDR_PRIORITY = '1')"

        ' If we are not showing past cancel and not releasing past cancel, then get all orders past cancel out of the result set

        If Not chkShowPastCancel.Checked Then
            If Not Absx1.chkFor("CHKREL_PAST_CANCEL").Checked Then
                ASCMAIN1.sql &= " and SOTORDR1.ORDR_CANCEL_DATE >= '" & Format(DATETIME_STAMP.AddDays(Val(Absx1.numFor("FPDCANCEL_FUTURE_DAYS").Value & "")), "dd-MMM-yyyy") & "'"
            End If
        End If

        If WHSE_CODE <> "" Then
            ASCMAIN1.sql &= " and SOTORDR1.WHSE_CODE = '" & WHSE_CODE & "'"
        End If

        ASCMAIN1.sql &= SQL_in("SALES_DIVISION_CODE", "SOTORDR1.SALES_DIVISION_CODE")
        ASCMAIN1.sql &= SQL_in("CUST_CODE", "SOTORDR1.CUST_CODE")
        ASCMAIN1.sql &= SQL_in("TRADE_CLASS_CODE", "SOTORDR1.TRADE_CLASS_CODE")
        ASCMAIN1.sql &= SQL_in("ORDR_GROUP_NO", "SOTORDR1.ORDR_GROUP_NO")
        ASCMAIN1.sql &= SQL_in("ORDR_NO", "SOTORDR1.ORDR_NO")

        If optORDR_SOURCE.Value <> "*" Then
            ASCMAIN1.sql &= " and SOTORDR1.ORDR_SOURCE = '" & optORDR_SOURCE.Value & "'"
        End If

        ' ISSUE-7373 - did not pick up amazon order(s) when All warehouses selected
        'If optWHSE.Value = "A" Then 'If all warehouses selected, cust_code <> AMAZON as ordr groups can refer to different whses
        '    ASCMAIN1.sql &= " and SOTORDR1.CUST_CODE <> 'AMAZON'"
        'End If

        SOTORDR1 = ASCMAIN1.Temp_Table
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
        ASCDATA1.ExecuteSQL()

        Dim sqlw As String = ""
        Dim sql_ITEM_CODE As String = SQLA("ITEM_CODE")
        If sql_ITEM_CODE <> "" Then
            sqlw = "Select Distinct SOTORDR2.ORDR_NO" & vbCrLf _
                & " from SOTORDR2, " & SOTORDR1 & " SOTORDR1" & vbCrLf _
                & " where SOTORDR2.ORDR_STATUS = 'O'" & vbCrLf _
                & "   and SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ITEM_CODE in (" & sql_ITEM_CODE & ")"
            If SQLA("ITEM_CODE", "EXCLUDE") = "1" Then
                ASCMAIN1.sql = "Delete from " & SOTORDR1 & " where ORDR_NO in (" & sqlw & ")"
            Else
                ASCMAIN1.sql = "Delete from " & SOTORDR1 & " where ORDR_NO not in (" & sqlw & ")"
            End If
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Create Index I_" & SOTORDR1 & "_1 on " & SOTORDR1 & " (ORDR_GROUP_NO)"
        ASCDATA1.ExecuteSQL()

        If chkMoneyOnly.Checked Then
            ASCMAIN1.sql = "Delete from " & SOTORDR1 & " where ORDR_GROUP_NO in (" & vbCrLf _
                & "Select ORDR_GROUP_NO from SOTORDR0 " & vbCrLf _
                & " where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")" & vbCrLf _
                & "   and ORDR_AMT < 1" & ")"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.AnalyzeTable(SOTORDR1)

        ' attempt to lock all orders in queue, and exit if we are unsuccessful

        If Not chkAllocateNoRelease.Checked Then
            ASCMAIN1.sql = "Select Distinct ORDR_GROUP_NO from " & SOTORDR1
            For Each row As DataRow In ASCDATA1.GetDataTable.Select()
                Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO")
                If Not ASCMAIN1.Logical_Lock("SOTORDR0", ORDR_GROUP_NO, , False, 1) Then
                    ASCMAIN1.MultiTask_Release(, , 1)
                    MsgBox("Could Not Lock Order Group " & ORDR_GROUP_NO, MsgBoxStyle.OkOnly, "Release Terminated")
                    Exit Function
                End If
            Next
        End If

        ' Set up Order Details

        ASCMAIN1.sql = "Select SOTORDR2.*, SOTORDR1.ORDR_GROUP_NO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO,ARTCUST1.CUST_CODE) CUST_CODE_ALLO" _
            & " from SOTORDR2, " & SOTORDR1 & " SOTORDR1, ARTCUST1" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE"

        SOTORDR2 = ASCMAIN1.Temp_Table
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
        With dst.Tables("SOTORDR1")
            .Columns.Add("ORDR_REL_EXCEPTIONS")
        End With
        Fill_Records("SOTORDR1")


        ASCMAIN1.sql = "Select SOTORDR5.* from SOTORDR5 where ROWNUM < 1"
        SOTORDR5 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR5 & " Add Primary Key (ORDR_NO, CUST_ADDR_TYPE)")
        ASCMAIN1.sql = "Select * from " & SOTORDR5
        Create_TDA(dst.Tables.Add("SOTORDR5"), SOTORDR5, "**", 0)




        ASCMAIN1.sql = $"SELECT POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO
, LEAST(POTORDR2.PO_DATE_REQUIRED, NVL(POTORDR2.PO_DATE_ETD,POTORDR2.PO_DATE_REQUIRED)) EARLIER
, POTORDR2.PO_DATE_REQUIRED
            , POTORDR2.PO_DATE_REQUESTED
            , POTORDR2.PO_DATE_ETD
             from POTORDR2,POTORDR1 
             where POTORDR2.ITEM_CODE in (Select ITEM_CODE from {SOTORDR2} where ORDR_NO = :PARM1)
               and POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO
               and POTORDR2.PO_STATUS = 'O'
AND POTORDR2.WHSE_CODE = :PARM2"
        ASCMAIN1.sql = "Select MIN(PO_DATE_REQUIRED) CONF from (" & ASCMAIN1.sql & ") X"
        Create_TDA(dst.Tables.Add, "POTORDRX", "**", 0, False, "VV")


        ASCMAIN1.sql = $"Select ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO
                , ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE
                , SUM(ICTPINV2.PINV_QTY) INV_QTY
                 from ICTPINV1, ICTPINV2
                 where ICTPINV1.PINV_NO = ICTPINV2.PINV_NO
                 And ICTPINV1.PINV_STATUS = 'O'
                 AND ICTPINV2.ITEM_CODE in (Select ITEM_CODE from {SOTORDR2} where ORDR_NO = :PARM1)
 AND ICTPINV1.WHSE_CODE = :PARM2
                 group by ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO
                , ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE
"
        ASCMAIN1.sql = "Select MIN(ETA_DATE + 7) ETA2DC from (" & ASCMAIN1.sql & ") X"
        Create_TDA(dst.Tables.Add, "ICTPINVX", "**", 0, False, "VV")

        ' Set up Order Group Summary

        ASCMAIN1.sql = "Select SOTORDR0.*, ARTCUST1.CUST_PRIORITY_CODE" & vbCrLf _
            & " from SOTORDR0, ARTCUST1" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE" & vbCrLf _
            & "   and ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")"
        SOTORDR0 = ASCMAIN1.Temp_Table
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
        SOTOREL1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL1 & " Add Primary Key (ITEM_CODE)")

        Dim Z As String = Format(Now, "dd-MMM-yyyy")
        ASCMAIN1.sql = "Update " & SOTOREL1 _
            & " Set ITEM_IS_ALLOCATED = '1'" & vbCrLf _
            & " where ITEM_CODE in (Select Distinct ITEM_CODE from SOTALLO1" & vbCrLf _
            & " where ITEM_CODE in (Select ITEM_CODE from " & SOTOREL1 & ") and (DATE_START > '" & Z & "' or DATE_END > '" & Z & "'))"
        ASCDATA1.ExecuteSQL()

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
            & IIf(chkIncludeWOasATS.Checked, "", " and ROWNUM < 1") & vbCrLf _
            & " group by POTORDR2.ITEM_CODE, POTORDR2.WHSE_CODE"
        POTORDRW = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDRW & " Add Primary Key (ITEM_CODE, WHSE_CODE)")


        ' Item Status by Warehouse

        ASCMAIN1.sql = "Select ITEM_CODE, WHSE_CODE" & vbCrLf _
            & ", WHSE_QTY_ON_HAND QTY_ON_HAND, WHSE_QTY_OPEN QTY_OPEN, WHSE_QTY_PICK QTY_PICK, WHSE_QTY_ONPO QTY_ONPO" & vbCrLf _
            & " from ICTSTAT2 where (ITEM_CODE,WHSE_CODE)" & vbCrLf _
            & " in (Select Distinct ITEM_CODE,WHSE_CODE from " & SOTORDR2 & ")"
        SOTOREL2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add Primary Key (ITEM_CODE, WHSE_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_WO NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_ATS NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_TO_PICK NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_TO_DE_COMMIT NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTOREL2 & " Add QTY_UNRELEASED NUMBER (8,0)")

        If chkIncludeWOasATS.Checked = "1" Then
            ASCMAIN1.sql = "Update " & SOTOREL2 & " SOTOREL2" & vbCrLf _
                & " Set QTY_WO = (Select PO_QTY_OPN from " & POTORDRW & " POTORDRW" & vbCrLf _
                & " where ITEM_CODE = SOTOREL2.ITEM_CODE and WHSE_CODE = SOTOREL2.WHSE_CODE)"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Update " & SOTOREL2 & " SOTOREL2" & vbCrLf _
            & " Set QTY_ATS = NVL(QTY_ON_HAND,0) - NVL(QTY_PICK,0) - NVL(QTY_WO,0)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from " & SOTOREL2
        Create_TDA(dst.Tables.Add("SOTOREL2"), SOTOREL2, "**", 0)
        Fill_Records("SOTOREL2")

        With dst.Tables.Add("SOTORELA")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("ORDR_QTY_PICK", GetType(System.Int64))
            .Columns.Add("ORDR_QTY_OPEN", GetType(System.Int64))
            .Columns.Add("ORDR_QTY_BACK", GetType(System.Int64))
            .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE")}
        End With


        ' Customer Master : Sold-To - and all related Bill-To and Credit Group Parents and Children

        Dim SQL_ARTCUST1 As String = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
            & ", ARTCUST1.CUST_ALLOW_BACKORDER, '0' CUST_CASE_PACK_ONLY" & vbCrLf _
            & ", ARTCUST1.CUST_PRIORITY_CODE, ARTCUST1.CUST_SALES_HOLD" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_CODE_ALLO,ARTCUST1.CUST_CODE) CUST_CODE_ALLO" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_ALLOCATE_BY_STORE,'0') CUST_ALLOCATE_BY_STORE" & vbCrLf _
            & ", NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE) CUST_BILL_TO_CUST" & vbCrLf _
            & ", NVL(ARTCUST1_BT.CUST_CREDIT_GROUP_CUST,ARTCUST1_BT.CUST_CODE) CUST_CREDIT_GROUP_CUST" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,SOTTCLS1.MARKET_CODE)) MARKET_CODE" & vbCrLf _
            & ", DECODE(SOTMKTC1.MARKET_CODE_FC,NULL" & vbCrLf _
            & "   ,NVL(SOTMKTC1_CUST_CODE.RESTRICT_ORDER_RELEASE,'0')" & vbCrLf _
            & "   ,NVL(SOTMKTC1.RESTRICT_ORDER_RELEASE,'0')) RESTRICT_ORDER_RELEASE" & vbCrLf _
            & ", ARTCUST1.ORDR_CODE_3PL" & vbCrLf _
            & " from ARTCUST1,SOTTCLS1,SOTMKTC1 SOTMKTC1_CUST_CODE, SOTMKTC1,ARTCUST1 ARTCUST1_BT" & vbCrLf _
            & " where SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = ARTCUST1.CUST_CODE" & vbCrLf _
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

        ARTCUST1 = ASCMAIN1.Temp_Table

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
        ASCDATA1.ExecuteSQL()

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
        ASCDATA1.ExecuteSQL()

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
        ARTCUST1_CG = ASCMAIN1.Temp_Table
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
            & IIf(ASCMAIN1.CLIENT = "INT", "     and ARTOPEN1.INV_TYPE = 'I'" & vbCrLf, "") _
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
        ASCDATA1.ExecuteSQL()

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

        ASCMAIN1.Progress("Setting Up Allocation Work Tables")


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
        SOTALLOZ = ASCMAIN1.Temp_Table
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
        SOTALLOY = ASCMAIN1.Temp_Table
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
            .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE"), .Columns("ORDR_GROUP_NO")}
        End With

        With dst.Tables.Add("SOTORELC")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("CUST_CODE")
            .Columns.Add("ORDR_NO")
            .Columns.Add("ORDR_LNO", GetType(System.Int64))
            .PrimaryKey = New DataColumn() { .Columns("ITEM_CODE"),
                                            .Columns("CUST_CODE"),
                                            .Columns("ORDR_NO"),
                                            .Columns("ORDR_LNO")}
        End With


        ' Prepare Market Restrictions Table
        'NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?')))
        '& "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
        '& "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
        ',SOTMKTC1 SOTMKTC1_CUST_CODE

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
            & "Select NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", 0 FORECAST, Sum (ORDR_QTY_SHIP) SHIPPED, 0 IN_PICK" & vbCrLf _
            & " from SOTINVH2,ARTCUST1,SOTTCLS1,SOTMKTC1,SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "   and SOTMKTC1.RESTRICT_ORDER_RELEASE = '1'" & vbCrLf _
            & " group by NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTINVH2.ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", 0 FORECAST, 0 SHIPPED, Sum (ORDR_QTY_PICK) IN_PICK" & vbCrLf _
            & " from SOTORDR2,ARTCUST1,SOTTCLS1,SOTMKTC1,SOTMKTC1 SOTMKTC1_CUST_CODE" & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1_CUST_CODE.CUST_CODE (+) = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = NVL(SOTTCLS1.MARKET_CODE,'?')" & vbCrLf _
            & "   and SOTORDR2.ORDR_STATUS = 'P'" & vbCrLf _
            & "   and SOTMKTC1.RESTRICT_ORDER_RELEASE = '1'" & vbCrLf _
            & " group by NVL(SOTMKTC1_CUST_CODE.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,NVL(SOTTCLS1.MARKET_CODE,'?'))), SOTORDR2.ITEM_CODE" & vbCrLf _
            & ") group by MARKET_CODE, ITEM_CODE" & vbCrLf
        DPTITMFX = ASCMAIN1.Temp_Table
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

        ASCMAIN1.Progress("Released/Processed")
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

                Dim ORDR_TYPE_CODE As String = rowSOTORDR1.Item("ORDR_TYPE_CODE")

                Dim ORDR_HOLD_INVTY As String = ""

                Dim ORDR_CRED_CLEARED As String = rowSOTORDR1.Item("ORDR_CRED_CLEARED") & ""
                Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE") & ""
                Dim rowICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
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

                If Absx1.chkFor("CHKFORCE_PICK").Checked Then
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

                Dim thisOrderEverHeldForAllocations As Boolean = False
                Dim thisOrderEverHeldForShortage As Boolean = False
                Dim thisOrderEverHeldFor3PL As Boolean = False

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select _
                    ("ORDR_NO = '" & ORDR_NO & "'", "ORDR_LNO")

                    Dim ORDR_REL_HOLD_CODES_dtl As String = ""
                    Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
                    Dim ORDR_LNO As String = rowSOTORDR2.Item("ORDR_LNO")
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

                    If chkCompleteNoRS.Checked Then
                        If ITEM_ORDR_REL_CODE = "R" Or ITEM_ORDR_REL_CODE = "S" Then
                            ITEM_ORDR_REL_CODE = ""
                        End If
                    End If

                    If chkReleaseItemsOnHold.Checked And chkForcePick.Checked Then
                        If ITEM_ORDR_REL_CODE = "H" Then
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

                    Dim ORDR_DETAIL_LINE_QTY As Integer = Val(rowSOTORDR2.Item("ORDR_QTY"))
                    Dim MAX_DETAIL_LINE_QTY As Int32 = Val(rowICTWHSE1.Item("MAX_DETAIL_LINE_QTY") & "")
                    'if the line qty is greater than the max, hold code B
                    If MAX_DETAIL_LINE_QTY > 0 And ORDR_DETAIL_LINE_QTY > MAX_DETAIL_LINE_QTY Then
                        ORDR_REL_HOLD_CODES_dtl = "B"
                        If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                            ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                        End If
                        thisOrderEverHeldFor3PL = True
                    End If

                    If ITEM_ORDR_REL_CODE = "D" Then
                        Dim rowSOTDSCI1 As DataRow = dst.Tables("SOTDSCI1").Rows.Find(New String() {ITEM_CODE, CUST_CODE})
                        If rowSOTDSCI1 IsNot Nothing Then
                            'ITEM_ORDR_REL_CODE = "S" ISSUE-7279 1/9/2026
                            ITEM_ORDR_REL_CODE = ""
                        End If
                    End If
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
                            If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CH021A01" Then Stop
                            If ALLO_CTL_NO = "" Then
                                If (ITEM_IS_ALLOCATED = "1" Or (ITEM_SNU_CODE = "S" And ITEM_BASIC_PROMO = "P") Or ITEM_SNU_CODE = "N") And ITEM_NOT_ALLOCATED <> "1" Then
                                    If ORDR_OVERRIDE_NOT_ALLOCATED = "1" Then
                                        ' LET THE ORDER GO
                                    Else
                                        dst.Tables("SOTORELC").Rows.Add(New String() {ITEM_CODE,
                                              CUST_CODE_ALLO,
                                              rowSOTORDR2.Item("ORDR_NO"),
                                              rowSOTORDR2.Item("ORDR_LNO")})

                                        rowSOTORDR2.Item("ALLO_CTL_NO") = "0000000000"
                                        ORDR_REL_HOLD_CODES_dtl = "C"
                                        ORDR_HOLD_INVTY = "1"
                                        thisOrderEverHeldForAllocations = True
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
                                    If ASCMAIN1.CLIENT = "INT" Then
                                        cBAL = cBAL - Val(rowSOTALLOZ.Item("ORDR_QTY_PICK_LATER") & "")
                                    End If

                                    If ORDR_QTY_OPEN + ORDR_QTY_PICK_previous_lines > cBAL And rowSOTALLO1.Item("ALLOW_OVER") & "" <> "1" Then
                                        ORDR_REL_HOLD_CODES_dtl = "C"
                                        ORDR_HOLD_INVTY = "1"
                                        thisOrderEverHeldForAllocations = True
                                        If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                            ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                        End If
                                        dst.Tables("SOTORELC").Rows.Add(New String() {rowSOTORDR2.Item("ITEM_CODE"),
                                            CUST_CODE_ALLO,
                                            rowSOTORDR2.Item("ORDR_NO"),
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
                                            thisOrderEverHeldForAllocations = True
                                            If InStr(ORDR_REL_HOLD_CODES_hdr, ORDR_REL_HOLD_CODES_dtl) = 0 Then
                                                ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & ORDR_REL_HOLD_CODES_dtl
                                            End If
                                            ' MAYBE WE NEED A DIFFERENT TABLE HERE
                                            If Not dst.Tables("SOTORELC").Rows.Contains(New String() {rowSOTORDR2.Item("ITEM_CODE"),
                                                CUST_CODE_ALLO,
                                                rowSOTORDR2.Item("ORDR_NO"),
                                                rowSOTORDR2.Item("ORDR_LNO")}) Then
                                                dst.Tables("SOTORELC").Rows.Add(New String() {rowSOTORDR2.Item("ITEM_CODE"),
                                                    CUST_CODE_ALLO,
                                                    rowSOTORDR2.Item("ORDR_NO"),
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

                            If Absx1.chkFor("CHKFORCE_PICK").Checked Then
                                ' Terms requiring CC Authorization do not honor Force Pick. user must chnage terms on sales order
                                If Not ORDR_REL_HOLD_CODES_hdr.Contains("R") And Not ORDR_REL_HOLD_CODES_hdr.Contains("Q") Then

                                    If thisOrderEverHeldForAllocations Or thisOrderEverHeldForShortage Or thisOrderEverHeldFor3PL Then
                                        ' NEVER ALLOW FORCE PICK TO OVERRIDE ALLOCATIONS, INVENTORY SHORTAGE, OR 3PL QUANTITY CONSTRAINT
                                    Else
                                        ORDR_HOLD_INVTY = ""
                                        ORDR_REL_HOLD_CODES_hdr = ""
                                        ORDR_REL_HOLD_CODES_dtl = ""
                                    End If

                                    'Dim ORDR_REL_HOLD_CODES_dtl_original As String = ORDR_REL_HOLD_CODES_dtl
                                    'Dim ORDR_REL_HOLD_CODES_hdr_original As String = ORDR_REL_HOLD_CODES_hdr
                                    'ORDR_HOLD_INVTY = ""
                                    'ORDR_REL_HOLD_CODES_hdr = ""
                                    'ORDR_REL_HOLD_CODES_dtl = ""

                                    'If ORDR_REL_HOLD_CODES_dtl_original.Contains("C") Then
                                    '    ' if the order is being held because of Allocations, still hold it for Allocations when Force Picking
                                    '    ORDR_HOLD_INVTY = "1"
                                    '    ORDR_REL_HOLD_CODES_hdr &= "C"
                                    '    ORDR_REL_HOLD_CODES_dtl &= "C"
                                    'End If
                                End If
                            End If

                            ' Authorizations and Order Qty Validation

                            If ASCMAIN1.CLIENT = "INT" Then
                                Dim tblSOTORDR2 As DataTable = dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & rowSOTORDR2.Item("ORDR_LNO")).CopyToDataTable
                                Dim listData As New List(Of String)
                                Dim errorMessages As String = String.Empty
                                errorMessages = TAC.SOCMAIN1.ValidateAuthorizations(tblSOTORDR2, listData)
                                If listData.Count > 0 Then
                                    If (ORDR_REL_HOLD_CODES_dtl & String.Empty).ToString.Length = 0 Then ORDR_REL_HOLD_CODES_dtl = "E"
                                    If InStr(ORDR_REL_HOLD_CODES_hdr, "E") = 0 Then
                                        ORDR_REL_HOLD_CODES_hdr &= "E"
                                    End If

                                    If rowSOTORDR1.Item("ORDR_REL_EXCEPTIONS") & "" <> "" Then
                                        errorMessages = vbCrLf & errorMessages
                                    End If
                                    rowSOTORDR1.Item("ORDR_REL_EXCEPTIONS") = rowSOTORDR1.Item("ORDR_REL_EXCEPTIONS") & errorMessages
                                End If

                                If (ORDR_TYPE_CODE = "RTV" And WHSE_CODE <> "CLA") Or rowARTCUST1.Item("ORDR_CODE_3PL") & "" = "Y" Then
                                    ' 09/11/2017 Walter, can you make order type Y not care about order mins and multiples?
                                Else
                                    listData.Clear()
                                    errorMessages = TAC.SOCMAIN1.ValidateOrderQtys(tblSOTORDR2, listData, "ORDR_QTY_OPEN")
                                    If listData.Count > 0 Then
                                        If (ORDR_REL_HOLD_CODES_dtl & String.Empty).ToString.Length = 0 Then ORDR_REL_HOLD_CODES_dtl = "Q"
                                        If InStr(ORDR_REL_HOLD_CODES_hdr, "Q") = 0 Then
                                            ORDR_REL_HOLD_CODES_hdr &= "Q"
                                        End If
                                    End If
                                End If
                            End If

                            If ORDR_REL_HOLD_CODES_dtl = "" Then
                                Dim ORDR_QTY_AVAIL As Int64 = Val(rowSOTOREL2.Item("QTY_ON_HAND") & "") _
                                                           - Val(rowSOTOREL2.Item("QTY_PICK") & "") _
                                                           - Val(rowSOTOREL2.Item("QTY_TO_PICK") & "") _
                                                           - ORDR_QTY_PICK_previous_lines
                                If ORDR_QTY_AVAIL >= ORDR_QTY_OPEN Then '  Or Absx1.chkFor("CHKFORCE_PICK").Checked Then ' LM wants inventory shortage to be observed even when force picking
                                    ORDR_QTY_PICK = ORDR_QTY_OPEN
                                    ORDR_REL_HOLD_CODES_dtl = "A"
                                Else ' We have an Inventory Shortage
                                    If (ITEM_ORDR_REL_CODE = "S" Or ITEM_ORDR_REL_CODE = "R") Then
                                        ORDR_REL_HOLD_CODES_dtl = ITEM_ORDR_REL_CODE
                                        Dim QTY_MIN As Int64 = 1
                                        If ASCMAIN1.CLIENT = "INT" Then
                                            QTY_MIN = ITEM_SO_QTY_MIN
                                        End If
                                        If ORDR_QTY_AVAIL > 0 And ORDR_QTY_AVAIL >= QTY_MIN Then
                                            ORDR_QTY_PICK = ORDR_QTY_AVAIL
                                            'If CUST_CASE_PACK_ONLY = "1" And ITEM_SO_QTY_MULT <> 0 Then
                                            '    If ORDR_QTY_PICK Mod ITEM_SO_QTY_MULT <> 0 Then
                                            '        ORDR_QTY_PICK = ORDR_QTY_PICK - ORDR_QTY_PICK Mod ITEM_SO_QTY_MULT
                                            '    End If
                                            'End If
                                            ' LM email 08/11/22 : make a change to conform to OM when using R & S
                                            If ITEM_SO_QTY_MULT <> 0 Then
                                                If ORDR_QTY_PICK Mod ITEM_SO_QTY_MULT <> 0 Then
                                                    ORDR_QTY_PICK = ORDR_QTY_PICK - ORDR_QTY_PICK Mod ITEM_SO_QTY_MULT
                                                End If
                                            End If
                                        End If
                                    Else
                                        ORDR_REL_HOLD_CODES_dtl = "I"
                                        ORDR_HOLD_INVTY = "1"
                                        thisOrderEverHeldForShortage = True
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
                    rowSOTORDR2.Item("PICK_QTY_CANC_REL") = ORDR_QTY_CANC 'if code is R or A, no customer has 1 for C_A_B

                    Dim newRow As DataRow = dst.Tables("SOTORDR_HOLD_CODES").NewRow()
                    newRow("ORDR_NO") = ORDR_NO
                    newRow("ITEM_CODE") = ITEM_CODE
                    newRow("ORDR_REL_HOLD_CODES_dtl") = ORDR_REL_HOLD_CODES_dtl
                    newRow("ITEM_ORDR_REL_CODE") = ITEM_ORDR_REL_CODE
                    dst.Tables("SOTORDR_HOLD_CODES").Rows.Add(newRow)

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
                If THIS_ORDR_QTY_PICK > 0 Then
                    'if there was no single line item over the max qty for an order, we need to check the sum still
                    If InStr(ORDR_REL_HOLD_CODES_hdr, "B") = 0 Then
                        Dim MAX_SALES_ORDER_QTY As Int32 = Val(rowICTWHSE1.Item("MAX_SALES_ORDER_QTY") & "")
                        Dim ORDR_QTY As Integer = Val(dst.Tables("SOTORDR2").Compute("Sum(ORDR_QTY_OPEN)", $"ORDR_NO = '{ORDR_NO}' AND (ORDR_STATUS = 'O' OR ORDR_STATUS = 'P')") & "")
                        ' Check if the total order quantity exceeds the maximum allowed quantity, if so, code B
                        If MAX_SALES_ORDER_QTY > 0 And ORDR_QTY > MAX_SALES_ORDER_QTY Then
                            ORDR_REL_HOLD_CODES_hdr &= "B"
                        End If
                    End If
                End If

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
                            If Not Absx1.chkFor("CHKFORCE_PICK").Checked Then
                                If InStr(ORDR_HOLD_CREDIT, "L") = 0 Then
                                    ORDR_HOLD_CREDIT &= "L"
                                End If
                            End If
                        End If
                    End If

                End If

                ' LM 12/29/2021 
                ' - I don't want it to override the sales order hold.  I purposely have to put on hold for various Sales Planning reasons. Yes. Absolutely override credit when forced.
                ' - For release, when I force an order I do want the days in this box to be respected. There are times when I do release less than 5 days
                ' so I am remming out the chkForcePick logic, so that this always applies
                ' If Not Absx1.chkFor("CHKFORCE_PICK").Checked Then
                If rowSOTORDR1.Item("ORDR_HOLD_SALES") & "" = "1" Or rowSOTORDR1.Item("ORDR_HOLD") & "" = "1" Then
                    ORDR_HOLD_SALES = "1"
                    ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "O"
                End If
                If Format(rowSOTORDR1.Item("ORDR_CANCEL_DATE"), "yyyyMMdd") _
                    < Format(DATETIME_STAMP.Date.AddDays(Val(Absx1.numFor("FPDCANCEL_FUTURE_DAYS").Value & "")), "yyyyMMdd") And Not chkReleasePastCancel.Checked Then
                    ORDR_HOLD_SALES = "1"
                    ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "D"
                End If
                ' End If

                '12/04 LM = Can you please amend your code so that it is only for Macys/Macyscom?
                If CUST_CODE = "MACYS" Or CUST_CODE = "MACYSCOM" Then
                    '> We can never have a cancel date prior to an arrival date.
                    '  regardless if force picked
                    If rowSOTORDR1.Item("ORDR_ARRIVAL_DATE") & "" <> "" Then
                        If Format(rowSOTORDR1.Item("ORDR_CANCEL_DATE"), "yyyyMMdd") _
                         < Format(rowSOTORDR1.Item("ORDR_ARRIVAL_DATE"), "yyyyMMdd") Then
                            ORDR_HOLD_SALES = "1"
                            ORDR_REL_HOLD_CODES_hdr = ORDR_REL_HOLD_CODES_hdr & "D" ' SHOULD WE USE A DIFFERENT CODE FOR THESE?
                        End If
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

                If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" AndAlso CUST_CODE = "IPLBAE" Then
                    'If ORDR_HOLD_INVTY = "1" And ORDR_REL_HOLD_CODES_hdr = "I" Then
                    '    ORDR_HOLD_INVTY = ""
                    '    ORDR_HOLD_SALES = ""
                    'End If
                End If

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
                        Item_Status_Qtys(ITEM_CODE, WHSE_CODE,
                                         ORDR_QTY_OPEN, ORDR_QTY_PICK, ORDR_QTY_BACK,
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

                ASCMAIN1.Progress("-", CStr(ORDERS_RELEASED_TOTAL + ORDERS_RELEASED) & "/" & CStr(ORDR_NO_ctr) & ":" & CUST_CODE)
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


        If chkAllocateNoRelease.Checked Then
            ' do not release IPLBAE
        Else

            Dim ORDR_NO_AE_SAVED As String = ""
            ' TRY TO RELEASE AS MANY LINES AS WE HAVE STOCK FOR, CREATING NEW ORDERS(BOS) WITH THE REMAINING LINE ITEMS
            For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select("CUST_CODE = 'IPLBAE' and ORDERS_HELD <> 0")
                Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO")
                Dim CUST_CODE As String = rowSOTORDR0.Item("CUST_CODE")


                Dim WHSE_CODE As String = rowSOTORDR0.Item("WHSE_CODE")
                Dim ORDR_REL_HOLD_CODES As String = rowSOTORDR0.Item("ORDR_REL_HOLD_CODES") & ""

                ' Check to see if order is releasable
                ' remember that we may de-release all orders in the batch if not all made it

                ' NOTE = X CODES (AND MAYBE N CODES) HAVE NOT YET BEEN APPLIED 
                If ORDR_REL_HOLD_CODES = "I" Then ' Replace(Replace(Replace(ORDR_REL_HOLD_CODES, "I", ""), "X", ""), "N", "") = "" Then
                    Dim orders_with_Inventory_Shortage_hold As Integer = 0
                    Dim orders_with_no_hold As Integer = 0
                    Dim release_anyway As Boolean = True
                    For Each rowSOTORDR1_AE As DataRow In dst.Tables("SOTORDR1").Select($"ORDR_GROUP_NO = '{ORDR_GROUP_NO}'")
                        Dim ORDR_NO_AE As String = rowSOTORDR1_AE.Item("ORDR_NO")
                        Dim ORDR_REL_HOLD_CODES_AE As String = rowSOTORDR1_AE.Item("ORDR_REL_HOLD_CODES") & ""
                        If ORDR_REL_HOLD_CODES_AE = "I" Then
                            ' note that I may not be able to release if there are no releasable items on it
                            orders_with_Inventory_Shortage_hold += 1
                        ElseIf ORDR_REL_HOLD_CODES_AE = "" Then
                            orders_with_no_hold += 1
                        Else ' don't know about this hold
                            release_anyway = False
                            Exit For
                        End If
                    Next

                    If orders_with_Inventory_Shortage_hold = 0 Then
                        release_anyway = False
                    End If

                    If release_anyway Then

                        For Each rowSOTORDR1_AE As DataRow In dst.Tables("SOTORDR1").Select($"ORDR_GROUP_NO = '{ORDR_GROUP_NO}' AND ORDR_STATUS = 'O'")
                            Dim ORDR_NO_AE As String = rowSOTORDR1_AE.Item("ORDR_NO")
                            Dim ORDR_NO_AE_NEW As String = ""
                            Dim ORDR_STATUS_AE As String = rowSOTORDR1_AE.Item("ORDR_STATUS") & ""
                            Dim CUST_STORE_NO As String = rowSOTORDR1_AE.Item("CUST_STORE_NO")
                            Dim lines_Released As Int32 = 0
                            Dim lines_Short As Int32 = 0

                            ' 01/16/2026 - ISSUE-7316 IPLBAE can have multiple warehouses (SOTORDR1) in a group
                            If rowSOTORDR1_AE.Item("WHSE_CODE") & "" <> "" Then
                                WHSE_CODE = rowSOTORDR1_AE.Item("WHSE_CODE") & ""
                            End If

                            'Dim ORDR_AMT_PICK As Decimal = Val(rowSOTORDR1_AE.Item("ORDR_AMT_PICK") & "")
                            'Dim ORDR_QTY_PICK As Decimal = Val(rowSOTORDR1_AE.Item("ORDR_QTY_PICK") & "")
                            Dim was_anything_released As Boolean = False

                            Dim ORDR_REL_HOLD_CODES_AE As String = rowSOTORDR1_AE.Item("ORDR_REL_HOLD_CODES") & ""

                            For Each rowSOTORDR2_AE As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO_AE}' AND ORDR_QTY_OPEN > 0") ' PICK = 0?
                                With rowSOTORDR2_AE
                                    Dim ORDR_RELEASE As String = .Item("ORDR_RELEASE") & ""
                                    Dim ITEM_CODE As String = .Item("ITEM_CODE") & ""

                                    Dim release_what_we_can As Boolean = False
                                    If ORDR_RELEASE = "A" Then
                                        If ORDR_REL_HOLD_CODES_AE = "I" Then

                                            Dim rowSOTOREL2 As DataRow = dst.Tables("SOTOREL2").Rows.Find(New String() {ITEM_CODE, WHSE_CODE})

                                            Dim ORDR_QTY_AVAIL As Int64 = Val(rowSOTOREL2.Item("QTY_ON_HAND") & "") _
                                                                           - Val(rowSOTOREL2.Item("QTY_PICK") & "") _
                                                                           - Val(rowSOTOREL2.Item("QTY_TO_PICK") & "") _
                                                                           - 0 ' ORDR_QTY_PICK_previous_lines  we will assume that an item will appear once and only once per order

                                            Dim ORDR_QTY_PICK As Int32 = Val(.Item("ORDR_QTY_PICK") & "")
                                            Dim ORDR_QTY_OPEN As Int32 = Val(.Item("ORDR_QTY_OPEN") & "")
                                            Dim ORDR_UNIT_PRICE As Decimal = Val(.Item("ORDR_UNIT_PRICE") & "")

                                            If ORDR_QTY_PICK <> 0 Then
                                                If ASCMAIN1.Running_in_VS Then Stop ' WHY IS THIS ALREADY NON-0?
                                            End If

                                            If ORDR_QTY_AVAIL > ORDR_QTY_OPEN Then
                                                .Item("ORDR_QTY_PICK") = ORDR_QTY_OPEN
                                                .Item("ORDR_QTY_OPEN") = 0
                                                .Item("PICK_QTY_CANC_REL") = 0
                                                .Item("PICK_QTY_BACK_REL") = 0

                                                .Item("ORDR_STATUS") = "P"

                                                rowSOTORDR1_AE.Item("ORDR_QTY_PICK") = Val(rowSOTORDR1_AE.Item("ORDR_QTY_PICK") & "") + ORDR_QTY_OPEN
                                                rowSOTORDR1_AE.Item("ORDR_AMT_PICK") = Val(rowSOTORDR1_AE.Item("ORDR_AMT_PICK") & "") + ORDR_QTY_OPEN * ORDR_UNIT_PRICE

                                                lines_Released += 1
                                                release_what_we_can = True
                                            Else
                                                .Item("ORDR_RELEASE") = "I"
                                                ORDR_RELEASE = "I"
                                            End If

                                        End If
                                    End If

                                    If ORDR_RELEASE = "A" Then
                                        ' THESE FIELDS ARE ALREADY SET TO RELEASE - UNLESS THE ORDER WAS SET TO I
                                        'If ORDR_REL_HOLD_CODES_AE = "I" Then
                                        '    If Val(.Item("ORDR_QTY_PICK") & "") <> 0 Then
                                        '        If ASCMAIN1.Running_in_VS Then Stop ' WHY IS THIS ALREADY NON-0?
                                        '    End If
                                        '    .Item("ORDR_QTY_PICK") = .Item("ORDR_QTY_OPEN")
                                        '    .Item("ORDR_QTY_OPEN") = 0
                                        '    .Item("PICK_QTY_CANC_REL") = 0
                                        '    .Item("PICK_QTY_BACK_REL") = 0

                                        '    .Item("ORDR_STATUS") = "P"

                                        '    rowSOTORDR1_AE.Item("ORDR_QTY_PICK") = Val(rowSOTORDR1_AE.Item("ORDR_QTY_PICK") & "") + Val(.Item("ORDR_QTY_PICK") & "")
                                        '    rowSOTORDR1_AE.Item("ORDR_AMT_PICK") = Val(rowSOTORDR1_AE.Item("ORDR_AMT_PICK") & "") + Val(.Item("ORDR_QTY_PICK") & "") * Val(.Item("ORDR_UNIT_PRICE") & "")
                                        'End If

                                        'lines_Released += 1
                                    Else
                                        .Item("ORDR_QTY_PICK") = 0
                                        .Item("ORDR_STATUS") = "P"

                                        If ORDR_RELEASE = "D" Or ORDR_RELEASE = "S" Or ORDR_RELEASE = "R" Then
                                            .Item("PICK_QTY_CANC_REL") = .Item("ORDR_QTY_OPEN")
                                            .Item("PICK_QTY_BACK_REL") = 0
                                        Else
                                            .Item("PICK_QTY_CANC_REL") = 0
                                            .Item("PICK_QTY_BACK_REL") = .Item("ORDR_QTY_OPEN")

                                            If ORDR_NO_AE_NEW = "" Then
                                                If ORDR_NO_AE_SAVED <> "" Then
                                                    ORDR_NO_AE_NEW = ORDR_NO_AE_SAVED
                                                    ORDR_NO_AE_SAVED = ""
                                                Else
                                                    ORDR_NO_AE_NEW = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                                                End If

                                                Dim rowSOTORDR1_AE_NEW As DataRow = dst.Tables("SOTORDR1").NewRow
                                                rowSOTORDR1_AE_NEW.ItemArray = rowSOTORDR1_AE.ItemArray
                                                rowSOTORDR1_AE_NEW.Item("ORDR_NO") = ORDR_NO_AE_NEW
                                                rowSOTORDR1_AE_NEW.Item("ORDR_STATUS") = "O"
                                                rowSOTORDR1_AE_NEW.Item("INIT_DATE") = DATETIME_STAMP
                                                rowSOTORDR1_AE_NEW.Item("INIT_OPER") = ASCMAIN1.USER_ID
                                                rowSOTORDR1_AE_NEW.Item("ORDR_NO_ORIG") = ORDR_NO_AE
                                                rowSOTORDR1_AE_NEW.Item("ORDR_BACKORDER") = "1"

                                                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1_AE_NEW)

                                                Fill_Records("SOTORDR5", , False, $"Select * from SOTORDR5 where ORDR_NO = '{ORDR_NO_AE}'")
                                                For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO_AE}'")
                                                    rowSOTORDR5.Item("ORDR_NO") = ORDR_NO_AE_NEW
                                                    rowSOTORDR5.AcceptChanges()
                                                    rowSOTORDR5.SetAdded()
                                                Next
                                            End If

                                            Dim rowSOTORDR2_AE_NEW As DataRow = dst.Tables("SOTORDR2").NewRow
                                            rowSOTORDR2_AE_NEW.ItemArray = rowSOTORDR2_AE.ItemArray
                                            rowSOTORDR2_AE_NEW.Item("ORDR_NO") = ORDR_NO_AE_NEW
                                            rowSOTORDR2_AE_NEW.Item("ORDR_STATUS") = "O"
                                            rowSOTORDR2_AE_NEW.Item("ORDR_QTY") = .Item("ORDR_QTY_OPEN")
                                            rowSOTORDR2_AE_NEW.Item("ORDR_QTY_OPEN") = .Item("ORDR_QTY_OPEN")
                                            rowSOTORDR2_AE_NEW.Item("ORDR_QTY_PICK") = 0
                                            rowSOTORDR2_AE_NEW.Item("ORDR_QTY_CANC") = 0
                                            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2_AE_NEW)


                                            ' ORDR_STATUS REMAINS O
                                            lines_Short += 1
                                        End If

                                    End If

                                End With
                            Next

                            If lines_Released = 0 Then
                                For Each rowSOTORDR2_AE As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO_AE}' AND ORDR_STATUS = 'P'")
                                    rowSOTORDR2_AE.Item("PICK_QTY_CANC_REL") = 0
                                    rowSOTORDR2_AE.Item("PICK_QTY_BACK_REL") = 0
                                    rowSOTORDR2_AE.Item("ORDR_STATUS") = "O"
                                Next
                                If ORDR_NO_AE_NEW <> "" Then
                                    ORDR_NO_AE_SAVED = ORDR_NO_AE_NEW
                                    ' we may want to save this order_no_ae_new so that we don't waste so many
                                    ASCDATA1.DeleteRows(dst.Tables("SOTORDR1"), $"ORDR_NO = '{ORDR_NO_AE_NEW}'")
                                    ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), $"ORDR_NO = '{ORDR_NO_AE_NEW}'")
                                    ASCDATA1.DeleteRows(dst.Tables("SOTORDR5"), $"ORDR_NO = '{ORDR_NO_AE_NEW}'")
                                End If
                            Else

                                If lines_Released > 0 And ORDR_STATUS_AE = "O" Then ' DO THIS PROCESS FOR ALL ORDERS WHICH WERE ON HOLD BUT ONLY FOR THEIR RELEASABLE LINES

                                    rowSOTORDR0.Item("ORDR_REL_HOLD_CODES") = "" ' should htis wait until all orders are processed?
                                    rowSOTORDR0.Item("ORDERS_RELEASED") = Val(rowSOTORDR0.Item("ORDERS_RELEASED") & "") + 1
                                    rowSOTORDR0.Item("ORDERS_HELD") = Val(rowSOTORDR0.Item("ORDERS_HELD") & "") - 1

                                    Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
                                    Dim CUST_CODE_ALLO As String = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
                                    Dim CUST_ALLOCATE_BY_STORE As String = rowARTCUST1.Item("CUST_ALLOCATE_BY_STORE") & ""
                                    Dim CUST_ALLOW_BACKORDER As String = rowARTCUST1.Item("CUST_ALLOW_BACKORDER") & ""
                                    Dim CUST_AMT_PICK As Decimal = Val(rowARTCUST1.Item("CUST_AMT_PICK") & "")
                                    Dim CUST_AMT_OPEN As Decimal = Val(rowARTCUST1.Item("CUST_AMT_OPEN") & "")

                                    Dim ORDR_AMT_PICK_adj As Decimal = 0
                                    Dim ORDR_AMT_OPEN_adj As Decimal = 0

                                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO_AE & "' AND ORDR_RELEASE = 'A'")
                                        ' THIS LOOP DEALS WITH ORDERS MARKED AS I BUT WITH RELEASABLE INNER LINES
                                        Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
                                        Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTORDR2.Item("ORDR_UNIT_PRICE") & "")

                                        Dim ORDR_QTY_PICK As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_PICK") & "")
                                        Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                                        Dim ORDR_QTY_BACK As Int64 = Val(rowSOTORDR2.Item("PICK_QTY_BACK_REL") & "")

                                        ORDR_AMT_PICK_adj += ORDR_QTY_OPEN * ORDR_UNIT_PRICE
                                        ORDR_AMT_OPEN_adj += ORDR_QTY_OPEN * ORDR_UNIT_PRICE

                                        Item_Status_Qtys(ITEM_CODE, WHSE_CODE,
                                                                1 * ORDR_QTY_OPEN,
                                                                1 * ORDR_QTY_PICK,
                                                                1 * ORDR_QTY_BACK,
                                                                rowARTCUST1.Item("RESTRICT_ORDER_RELEASE"),
                                                                rowARTCUST1.Item("MARKET_CODE"))

                                        Dim ALLO_CTL_NO As String = rowSOTORDR2.Item("ALLO_CTL_NO") & ""
                                        If ALLO_CTL_NO <> "" Then
                                            With dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO})
                                                .Item("ORDR_QTY_PICK_NOW") = Val(.Item("ORDR_QTY_PICK_NOW") & "") - ORDR_QTY_PICK
                                                .Item("ORDR_QTY_PICK_LATER") = Val(.Item("ORDR_QTY_PICK_LATER") & "") + ORDR_QTY_PICK
                                            End With
                                            If CUST_ALLOCATE_BY_STORE = "1" Then
                                                Dim rowSOTALLOY As DataRow = dst.Tables("SOTALLOY").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO, CUST_STORE_NO})
                                                If rowSOTALLOY Is Nothing Then
                                                    rowSOTALLOY = dst.Tables("SOTALLOY").NewRow
                                                    With rowSOTALLOY
                                                        .Item("ALLO_CTL_NO") = ALLO_CTL_NO
                                                        .Item("CUST_CODE_ALLO") = CUST_CODE_ALLO
                                                        .Item("CUST_STORE_NO") = CUST_STORE_NO
                                                    End With
                                                    dst.Tables("SOTALLOY").Rows.Add(rowSOTALLOY)
                                                End If
                                                With rowSOTALLOY
                                                    .Item("ORDR_QTY_PICK_NOW") = Val(.Item("ORDR_QTY_PICK_NOW") & "") - ORDR_QTY_PICK
                                                    .Item("ORDR_QTY_PICK_LATER") = Val(.Item("ORDR_QTY_PICK_LATER") & "") + ORDR_QTY_PICK
                                                End With
                                            End If
                                        End If
                                        ' rowSOTORDR2.Item("ORDR_QTY_PICK") = 0
                                        ' rowSOTORDR2.Item("ORDR_STATUS") = "O"
                                    Next

                                    CUST_AMT_PICK += ORDR_AMT_PICK_adj
                                    CUST_AMT_OPEN -= ORDR_AMT_OPEN_adj
                                    ORDERS_RELEASED_TOTAL = ORDERS_RELEASED_TOTAL + 1

                                    rowARTCUST1.Item("CUST_AMT_PICK") = CUST_AMT_PICK
                                    rowARTCUST1.Item("CUST_AMT_OPEN") = CUST_AMT_OPEN
                                    rowARTCUST1.Item("RELEASED") = "Y"

                                    'rowSOTORDR0.Item("ORDERS_HELD") = Val(rowSOTORDR0.Item("ORDERS_HELD") & "") + Val(rowSOTORDR0.Item("ORDERS_RELEASED") & "")
                                    'rowSOTORDR0.Item("ORDERS_RELEASED") = 0

                                    rowSOTORDR1_AE.Item("ORDR_REL_BATCH_NO") = Mid(XNO, 5, 6)
                                    rowSOTORDR1_AE.Item("ORDR_DATE_REL") = DATETIME_STAMP.Date
                                    rowSOTORDR1_AE.Item("ORDR_REL_HOLD_CODES") = ""
                                    rowSOTORDR1_AE.Item("ORDR_STATUS") = "P"
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End If

        ASCMAIN1.Progress("De-Releasing Partially Released PO's")

        For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select _
            ("CUST_CODE <> 'IPLBAE' AND ORDERS_HELD <> 0 and ORDERS_RELEASED <> 0")
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
            If chkCancelIfNoPick.Checked And ORDR_REL_HOLD_CODES = "N" Then

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
                        Item_Status_Qtys(ITEM_CODE, WHSE_CODE,
                                         -1 * ORDR_QTY_OPEN,
                                         -1 * ORDR_QTY_PICK,
                                         -1 * ORDR_QTY_BACK,
                                         rowARTCUST1.Item("RESTRICT_ORDER_RELEASE"),
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
                .Item("ORDR_QTY_OPEN") = 0 ' NEW
                .Item("ORDR_QTY_CANC") = Val(.Item("ORDR_QTY_CANC") & "") + Val(.Item("PICK_QTY_CANC_REL") & "")
            End With
        Next

        ' Tally the total qty, by item, on all Orders Not Released (Held Orders)

        ASCMAIN1.Progress("Un-Released Invty Req")
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

        ASCMAIN1.Progress("Checking Holds for Item Allocation")

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


        ASCMAIN1.Progress("Checking Holds for Forecast Allocation")

        With dst.Tables.Add("SOTOREL9")
            .Columns.Add("MARKET_CODE")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("ORDR_NO")
            .Columns.Add("ORDR_LNO", GetType(System.Int64))
            .PrimaryKey = New DataColumn() { .Columns("MARKET_CODE"),
                                            .Columns("ITEM_CODE"),
                                            .Columns("ORDR_NO"),
                                            .Columns("ORDR_LNO")}
        End With

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_RELEASE = 'M'")
            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(rowSOTORDR2.Item("CUST_CODE"))
            dst.Tables("SOTOREL9").Rows.Add(New String() {rowARTCUST1.Item("MARKET_CODE"),
                                                          rowSOTORDR2.Item("ITEM_CODE"),
                                                          rowSOTORDR2.Item("ORDR_NO"),
                                                          rowSOTORDR2.Item("ORDR_LNO")})
        Next


        If dst.Tables("SOTOREL2").Select("ISNULL(QTY_ON_HAND,0) - ISNULL(QTY_PICK,0) - ISNULL(QTY_TO_PICK,0) < 0 AND ISNULL(QTY_TO_PICK,0) <> 0").Length > 0 Then
            MsgBox("There are items that are over-picked (compared to Qty Available)", MsgBoxStyle.OkOnly, "Cannot permit Sales Order Release to Contune")
            Using frm As New ASFMSGBF
                frm.Show_grd(dst.Tables("SOTOREL2").Select("ISNULL(QTY_ON_HAND,0) - ISNULL(QTY_PICK,0) - ISNULL(QTY_TO_PICK,0) < 0 AND ISNULL(QTY_TO_PICK,0) <> 0").CopyToDataTable, Me, "There are items that are over-picked (compared to Qty Available)")
            End Using

            Return False
        End If

        Return True
    End Function

    Sub Item_Status_Qtys(
        ITEM_CODE As String, WHSE_CODE As String,
        ORDR_QTY_OPEN As Int64, ORDR_QTY_PICK As Int64, ORDR_QTY_BACK As Int64,
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

    Private Sub chkReleasePastCancel_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkReleasePastCancel.CheckedChanged
        Setup_Future_Days()
    End Sub

    Sub Setup_Future_Days()
        numFutureDays.Visible = Not chkReleasePastCancel.Checked
        lblFutureDays.Visible = Not chkReleasePastCancel.Checked
    End Sub

    Private Sub dteSHIP_DATE_BeforeMonthScroll(sender As System.Object, e As Infragistics.Win.UltraWinSchedule.BeforeMonthScrollEventArgs) Handles dteSHIP_DATE.BeforeMonthScroll

    End Sub

    Sub Load_CSO(CSO_NO As String)
        Dim rowSOTCSTO1 As DataRow = Fill_Record("SOTCSTO1", CSO_NO, False)
        Dim SELL_CODE As String = rowSOTCSTO1.Item("SELL_CODE")
        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE)

        Dim CUST_CODE As String = rowSOTCSTO1.Item("CUST_CODE")
        Dim CUST_STORE_NO As String = "000" & SELL_CODE
        Dim DATE_START As Date = rowSOTCSTO1.Item("DATE_START")
        Dim WHSE_CODE As String = rowSOTCSTO1.Item("WHSE_CODE")

        Dim ORDR_GROUP_NO As String = rowSOTCSTO1.Item("ORDR_GROUP_NO") & ""
        Create_Work_Tables_SOTALLOX(CSO_NO, WHSE_CODE, CUST_CODE, CUST_STORE_NO, SELL_CODE, DATE_START, ORDR_GROUP_NO)

        Fill_Records("SOTCSTO2", CSO_NO, False)
        Fill_Records("SOTCSTO3", CSO_NO, False)
        Fill_Records("SOTCSTO4", CSO_NO, False)


        Fill_Records("SOTCSTOX")
        Fill_Records("SOTCSTOI")

        dst.Tables("SOTALLOX").Rows.Clear()
        Fill_Records("SOTALLOX")

        ASCMAIN1.sql = sqlICTITEM1 & $" and ICTITEM1.ITEM_CODE in (Select ITEM_CODE from SOTCSTO2 where CSO_NO = '{CSO_NO}')"
        Fill_Records("ICTITEM1",,, ASCMAIN1.sql)

        For Each rowSOTCSTO2 As DataRow In dst.Tables("SOTCSTO2").Select()
            With rowSOTCSTO2
                Dim ITEM_CODE As String = .Item("ITEM_CODE")
                Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                '.Item("CSO_QTY") = QTY
                .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                .Item("ITEM_SNU_CODE") = rowICTITEM1.Item("ITEM_SNU_CODE")
                .Item("ITEM_SO_QTY_MULT") = rowICTITEM1.Item("ITEM_SO_QTY_MULT")
                .Item("ITEM_SO_QTY_MIN") = rowICTITEM1.Item("ITEM_SO_QTY_MIN")
                .Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                .Item("HC_CODE") = rowICTITEM1.Item("HC_CODE")
                .Item("BRAND_CODE") = rowICTITEM1.Item("BRAND_CODE")
                Dim row() As DataRow = dst.Tables("SOTALLOX").Select($"ITEM_CODE = '{ITEM_CODE}'")
                If row.Length = 1 Then
                    .Item("ALLO_GROUP_CODE") = row(0).Item("ALLO_GROUP_CODE")
                    '.Item("CSO_QTY_ALLO") = Val(row(0).Item("QTY_ALLO") & "")
                End If
                Dim CSO_LNO As Integer = Val(.Item("CSO_LNO") & "")
                For Each rowSOTCSTO4 As DataRow In dst.Tables("SOTCSTO4").Select($"CSO_LNO = {CStr(CSO_LNO)}")
                    Dim CSO_QTY As Integer = Val(rowSOTCSTO4.Item("CSO_QTY") & "")
                    Dim CSO_ADDR_LNO As Integer = Val(rowSOTCSTO4.Item("CSO_ADDR_LNO") & "")
                    .Item($"CSO_QTY_{Format(CSO_ADDR_LNO, "000")}") = CSO_QTY

                    Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Rows.Find(New Object() {CSO_NO, CSO_ADDR_LNO})
                    rowSOTCSTO3.Item("CSO_QTY_" & Format(CSO_LNO, "000")) = CSO_QTY
                Next
            End With
        Next


        For Each rowSOTCSTO2 As DataRow In dst.Tables("SOTCSTO2").Select()
            Dim ITEM_CODE As String = rowSOTCSTO2.Item("ITEM_CODE")
            Dim ALLO_CTL_NO As String = rowSOTCSTO2.Item("ALLO_CTL_NO")
            Dim rowSOTALLOX As DataRow = dst.Tables("SOTALLOX").Rows.Find(ALLO_CTL_NO)
            If rowSOTALLOX IsNot Nothing Then
                For Each COL As String In New String() _
                    {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC" _
                    , "QTY_LEFT", "QTY_ALLO" _
                    , "WHSE_QTY_ON_HAND", "WHSE_QTY_ONPO", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"}
                    rowSOTCSTO2.Item(COL) = rowSOTALLOX.Item(COL)
                Next
            End If
        Next

        ' Load_SSG()

        ASCMAIN1.Progress("Address Columns")

        ASCMAIN1.Progress("Item Columns")
        Dim CSO_COL As Integer = 0
        For Each rowSOTCSTO2 As DataRow In dst.Tables("SOTCSTO2").Select("", "CSO_LNO")
            CSO_COL += 1
            Dim C As String = $"CSO_QTY_{Format(CSO_COL, "000")}"
            rowSOTCSTO2.Item("CSO_COL") = CSO_COL
        Next
    End Sub

    Sub Create_Work_Tables_SOTALLOX(CSO_NO As String, WHSE_CODE As String, CUST_CODE As String, CUST_STORE_NO As String, SELL_CODE As String, DATE_START As Date, Optional ORDR_GROUP_NO As String = "")

        'Dim DATE_START_since_oracle As String = Format(DATE_START_since, "dd-MMM-yyyy")
        'Dim DATE_START_until_oracle As String = Format(DATE_START_until, "dd-MMM-yyyy")
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLOX)
        ASCMAIN1.sql = Replace(Replace(Replace(Replace(sqlSOTALLOX, ":PARM1", $"'{CUST_CODE}'"), ":PARM2", $"'{SELL_CODE}'"), ":PARM3", $"'{CSO_NO}'"), ":PARM4", $"'{WHSE_CODE}'")
        ASCDATA1.ExecuteSQL("Insert into " & SOTALLOX & " " & ASCMAIN1.sql)

        Dim DATE_START_oracle As String = Format(DATE_START, "dd-MMM-yyyy")
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLOW)
        ASCMAIN1.sql = Replace(Replace(Replace(sqlSOTALLOW, ":PARM1", $"'{CUST_CODE}'"), ":PARM2", $"'{CUST_STORE_NO}'"), " group by", $" and ORDR_GROUP_NO <> '{ORDR_GROUP_NO}'  group by")
        ASCDATA1.ExecuteSQL("Insert into " & SOTALLOW & " " & ASCMAIN1.sql)
    End Sub

    Sub Create_Work_Tables(initialize As Boolean, Optional sqlw As String = "")

        If initialize Then
            sqlw = " and rownum < 1"
        End If

        ' & "   and SOTALLO1.DATE_START >= :PARM3" & vbCrLf _
        ' & "   and SOTALLO1.DATE_START <= :PARM4" & vbCrLf _
        If initialize Then
            ASCMAIN1.sql = "Select SOTALLO1.ALLO_CTL_NO, SOTALLO3.QTY_ALLO, SOTALLO3.ALLO_NOTES" & vbCrLf _
                & ", SOTALLO2.QTY_ALLO QTY_ALLO_AES, SOTALLO2.ALLO_NOTES ALLO_NOTES_AES" & vbCrLf _
                & ", SOTALLO1.ITEM_CODE, SOTALLO1.DATE_START, SOTALLO1.DATE_END, SOTALLO1.ALLO_GROUP_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
                & ", ICTSTAT2.WHSE_QTY_ON_HAND, ICTSTAT2.WHSE_QTY_ONPO, ICTSTAT2.WHSE_QTY_OPEN, ICTSTAT2.WHSE_QTY_PICK" & vbCrLf _
                & "from SOTALLO1, SOTALLO2, SOTALLO3, ICTITEM1, ICTCOLL1, ARTCUST2, ICTSTAT2" & vbCrLf _
                & " where SOTALLO1.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
                & "   And SOTALLO3.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
                & "   And SOTALLO3.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
                & "   And SOTALLO2.CUST_CODE = :PARM1" & vbCrLf _
                & "   and SOTALLO3.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.SELL_CODE = :PARM2" & vbCrLf _
                & "   and SOTALLO1.ALLO_CTL_NO in (Select Distinct ALLO_CTL_NO from SOTCSTO2 where CSO_NO = :PARM3)" & vbCrLf _
                & "   and ICTSTAT2.ITEM_CODE (+) = SOTALLO1.ITEM_CODE" & vbCrLf _
                & "   and ICTSTAT2.WHSE_CODE (+) = :PARM4" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
            sqlSOTALLOX = ASCMAIN1.sql
            ASCMAIN1.sql = Replace(Replace(Replace(Replace(sqlSOTALLOX, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''")
            SOTALLOX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {SOTALLOX} Add Primary Key (ALLO_CTL_NO)")
            ASCDATA1.ExecuteSQL($"Create Index I_{SOTALLOX}_1 on {SOTALLOX} (ITEM_CODE)")

            ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_OPEN" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_PICK" & vbCrLf _
                & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_SHIP" & vbCrLf _
                & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
                & $"   where SOTORDR2.ALLO_CTL_NO in  (Select ALLO_CTL_NO from {SOTALLOX})" & vbCrLf _
                & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
                & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
                & "     and SOTORDR2.CUST_CODE = :PARM1" & vbCrLf _
                & "     and SOTORDR2.CUST_STORE_NO = :PARM2" & vbCrLf _
                & "     and SOTORDR2.ORDR_STATUS IN ('O','P','F','C')" & vbCrLf _
                & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"
            sqlSOTALLOW = ASCMAIN1.sql
            ASCMAIN1.sql = Replace(Replace(Replace(Replace(sqlSOTALLOW, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''")
            SOTALLOW = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {SOTALLOW} Add Primary Key (ALLO_CTL_NO)")
        End If

        ASCMAIN1.sql = "Select SOTCSTO1.*" & vbCrLf _
            & " from SOTCSTO1" & ASCMAIN1.SQL_Add_WHERE(sqlw)

        If SOTCSTOX = "" Then
            SOTCSTOX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTCSTOX & " Add Primary Key (CSO_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCSTOX)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCSTOX & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select SOTCSTO2.ITEM_CODE, ICTITEM1.ITEM_DESC, SOTCSTO2.CSO_LNO, SOTCSTO1.*" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.HC_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " from ICTITEM1, ICTCOLL1, SOTCSTO2, SOTCSTO1, " & SOTCSTOX & " X " & vbCrLf _
            & " where SOTCSTO2.CSO_NO = X.CSO_NO " & vbCrLf _
            & "   And ICTITEM1.ITEM_CODE = SOTCSTO2.ITEM_CODE" & vbCrLf _
            & "   And ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   And SOTCSTO1.CSO_NO = SOTCSTO2.CSO_NO"

        If SOTCSTOI = "" Then
            SOTCSTOI = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTCSTOI & " Add Primary Key (CSO_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCSTOI)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCSTOI & " " & ASCMAIN1.sql)
        End If

        If initialize Then
            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_SO_QTY_MIN" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & " from ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
            sqlICTITEM1 = ASCMAIN1.sql
            ICTITEM1 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1)
            ASCDATA1.ExecuteSQL("Insert into " & ICTITEM1 & " " & sqlICTITEM1)
        End If

    End Sub

    Private Sub grdSOTCSTO1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTCSTO1.AfterRowActivate
        If grdSOTCSTO1.ActiveRow IsNot Nothing AndAlso grdSOTCSTO1.ActiveRow.IsDataRow Then

            Dim CSO_NO As String = grdSOTCSTO1.ActiveRow.Cells("CSO_NO").Value
            Dim dvw As DataView = DirectCast(grdSOTCSTO3.DataSource, DataTable).DefaultView
            dvw.RowFilter = $"CSO_NO = '{CSO_NO}'"
            Sort_grdColumns(grdSOTCSTO3, "CSO_ADDR_LNO")

            Dim CSO_LNO As Integer = 0
            With grdSOTCSTO3.DisplayLayout.Bands(0)
                For Each row As DataRow In dst.Tables("SOTCSTO2").Select($"CSO_NO = '{CSO_NO}'", "CSO_LNO")
                    CSO_LNO = Val(row.Item("CSO_LNO") & "")
                    Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
                    Dim c As String = $"CSO_QTY_{Format(CSO_LNO, "000")}"
                    With .Columns(c)
                        .Header.Caption = ITEM_CODE
                        .Hidden = False
                    End With
                Next
                For CS As Integer = CSO_LNO + 1 To MAX_COLs
                    Dim c As String = $"CSO_QTY_{Format(CS, "000")}"
                    With .Columns(c)
                        .Hidden = True
                    End With
                Next
            End With




            grdSOTCSTO3.Visible = True
        Else
            grdSOTCSTO3.Visible = False

        End If
    End Sub

    Sub email_PO_Cuts(ORDR_CUST_PO As String, body As String)

        Dim FILENAME As String = ""

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

        If ASCMAIN1.Running_in_VS Then
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
            'EMAIL_ADDRESSs.Add("nicholas@absolution.com", "Nicholas")
        Else
            If ASCMAIN1.DBS_SERVER = "INT" Then
                EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
                If Not EMAIL_ADDRESSs.ContainsValue("Nathan Feron") Then
                    EMAIL_ADDRESSs.Add("nferon@interparfums.com", "Nathan Feron")
                End If

                If Not EMAIL_ADDRESSs.ContainsValue("Erin Lepore") Then
                    EMAIL_ADDRESSs.Add("elepore@interparfums.com", "Erin Lepore")
                End If
            Else
                EMAIL_ADDRESSs.Add("dmoore@interparfums.com", "Dani Moore")
            End If
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        'ATTACHMENTs.Add("Car-Stock Order", FILENAME)

        Dim EMAIL_SUBJECT As String = $"PO {ORDR_CUST_PO} has had cuts"
        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    EMAIL_SUBJECT, "POCUTS", True, False, ORDR_CUST_PO, "ORDR_CUST_PO", "Purchase Order Items Cut", body)
    End Sub

    Private Sub chkForcePick_CheckedChanged(sender As Object, e As EventArgs) Handles chkForcePick.CheckedChanged
        chkReleaseItemsOnHold.Visible = chkForcePick.Checked
    End Sub
End Class