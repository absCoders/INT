Public Class RSRBOWS1
    Dim RSTBOWS1 As String = ""

    Dim RSTBOWS1_CUST_CODE As String
    Dim RSTBOWS1_CUST_STORE_NO As String
    Dim RSTBOWS1_ITEM_CODE As String
    ' Dim XCs As New Dictionary(Of String, String)
    Dim sole_AE As String = ""
    Dim sole_REGION As String = ""
    Dim hide_column As Boolean = False
    Dim SN_START As Date = Nothing
    Dim SN_END As Date = Nothing

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1)
        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, 0)
        ' Absx1.cmbFor("RYW").ReadOnly = True

        cbeLY.SelectedIndex = 0

        If ASCMAIN1.USER_CODES.Contains("FS") Then
            sole_AE = "?"
            Dim rowTATUSER1 As DataRow = LookUp("TATUSER1", ASCMAIN1.USER_ID)
            If rowTATUSER1 IsNot Nothing Then
                sole_AE = rowTATUSER1.Item("SELL_CODE") & ""
                If sole_AE = "" Then
                    sole_REGION = rowTATUSER1.Item("REGION_CODE") & ""
                    If sole_REGION = "" Then
                        sole_AE = "?"
                    End If
                End If
            End If

            If ASCMAIN1.CYW.StartsWith("2021") Then
                cbeLY.SelectedIndex = 1
            End If
            cbeLY.ReadOnly = True
        End If

        If sole_AE <> "" Or sole_REGION <> "" Then
            Dim TDY As Date = Now.Date
            Dim YYYY As String = Format(TDY, "yyyy")

            If Format(TDY, "MM") < "07" Then
                SN_START = CDate("01/01/" & YYYY)
                SN_END = CDate("06/30/" & YYYY)
            Else
                SN_START = CDate("07/01/" & YYYY)
                SN_END = CDate("12/31/" & YYYY)
            End If

            'UltraTabControl1.Tabs(0).Enabled = False
            'UltraTabControl1.SelectedTab = UltraTabControl1.Tabs(1)
            grpPERIOD_RANGE.Visible = False
            lblN.Visible = False
            numWEEKS.Visible = False

            Dim report_For As String = ""
            If sole_REGION <> "" Then
                report_For = "Report for Region " & sole_REGION
            Else
                report_For = "Report for AE " & sole_AE
            End If
            lblSole_AE.Text = report_For & "," _
                & " showing Saleable Promo items with allocations from " & SN_START & " to " & SN_END & "," _
                & " and Saleable Basic with Sell-In in past 6 weeks or Sell-Thru Last/Next Weeks TY/LY"
            lblSole_AE.Visible = True

        End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        Dim LYX As Integer = 1
        LYX = Val(cbeLY.Value & "")

        Create_Work_File()

        With dst
            ASCMAIN1.sql = "Select * from " & RSTBOWS1
            Create_TDA(.Tables.Add, "RSTBOWS1", "**", 0, False, "", 3)
            With .Tables("RSTBOWS1").Columns

                If sole_AE <> "" Or sole_REGION <> "" Then
                    '.Add("INXIT_3WKS", GetType(System.Decimal), "ISNULL(SHIPW1,0)+ISNULL(SHIPW2,0)+ISNULL(SHIPW3,0)")
                    '3WKS in PO transit: ShpL1 + ShpL2+ ShpL3
                    .Add("INXIT_3WKS", GetType(System.Decimal), "ISNULL(SHIPW0,0)+ISNULL(SHIPW1,0)+ISNULL(SHIPW2,0)")
                    '3WKS in PO transit: ShpL0 + ShpL1 + ShpL2 PER ANDREA & STEF 03/01/2024
                    .Add("LYN10", GetType(System.Decimal), "ISNULL(RSLYN01,0)+ISNULL(RSLYN02,0)+ISNULL(RSLYN03,0)+ISNULL(RSLYN04,0)+ISNULL(RSLYN05,0)+ISNULL(RSLYN06,0)+ISNULL(RSLYN07,0)+ISNULL(RSLYN08,0)+ISNULL(RSLYN09,0)+ISNULL(RSLYN10,0)")
                    'LY Next 10 wks SLS: LY-N01+LY-N02+LY-N03+LY-N04+LY-N05+LY-N06+LY-N07+LY-N08+LY-N09+LY-N10
                    .Add("TYL4", GetType(System.Decimal), "ISNULL(RSTYL01,0)+ISNULL(RSTYL02,0)+ISNULL(RSTYL03,0)+ISNULL(RSTYL04,0)")
                    'TY-last 4wks SLS: TY-L01+TY-L02+TY-L03+TY-L04
                    .Add("LYL4", GetType(System.Decimal), "ISNULL(RSLYL01,0)+ISNULL(RSLYL02,0)+ISNULL(RSLYL03,0)+ISNULL(RSLYL04,0)")
                    'LY-last 4wks SLS: LY-L01+ LY-L02+LY-L03+LY-L04
                    .Add("L4TYLY", GetType(System.Decimal), "IIF(LYL4>0,(TYL4-LYL4)/LYL4,0)")
                    'L4 % LY: =IF(R4>0,((Q4-R4)/R4),"")
                    .Add("TYL10", GetType(System.Decimal), "ISNULL(RSTYL01,0)+ISNULL(RSTYL02,0)+ISNULL(RSTYL03,0)+ISNULL(RSTYL04,0)+ISNULL(RSTYL05,0)+ISNULL(RSTYL06,0)+ISNULL(RSTYL07,0)+ISNULL(RSTYL08,0)+ISNULL(RSTYL09,0)+ISNULL(RSTYL10,0)")
                    'TY-last 10wks SLS: TY-L01+TY-L02+TY-L03+TY-L04+ TY-L05+TY-L06+TY-L07+TY-L08+TY-L09+TY-L10
                    .Add("LYL10", GetType(System.Decimal), "ISNULL(RSLYL01,0)+ISNULL(RSLYL02,0)+ISNULL(RSLYL03,0)+ISNULL(RSLYL04,0)+ISNULL(RSLYL05,0)+ISNULL(RSLYL06,0)+ISNULL(RSLYL07,0)+ISNULL(RSLYL08,0)+ISNULL(RSLYL09,0)+ISNULL(RSLYL10,0)")
                    'LY-last 10wks SLS: LY-L01+ LY-L02+LY-L03+LY-L04+LY-L05+LY-L06+LY-L07+LY-L08+LY-L09+LY-L10
                    .Add("L10TYLY", GetType(System.Decimal), "IIF(LYL10>0,(TYL10-LYL10)/LYL10,0)")
                    'L10 % LY: =IF(R4>0,((Q4-R4)/R4),"")
                End If

                .Add("WTD_VAR_PCT", GetType(System.Decimal), "IIF(ISNULL(RSLYL01,0)=0,0,(ISNULL(RSTYL01,0)-ISNULL(RSLYL01,0))/ISNULL(RSLYL01,0))")
                .Add("MTD_VAR_PCT", GetType(System.Decimal), "IIF(ISNULL(RSLYMTD,0)=0,0,(ISNULL(RSTYMTD,0)-ISNULL(RSLYMTD,0))/ISNULL(RSLYMTD,0))")
                .Add("STD_VAR_PCT", GetType(System.Decimal), "IIF(ISNULL(RSLYSTD,0)=0,0,(ISNULL(RSTYSTD,0)-ISNULL(RSLYSTD,0))/ISNULL(RSLYSTD,0))")
                .Add("YTD_VAR_PCT", GetType(System.Decimal), "IIF(ISNULL(RSLYYTD,0)=0,0,(ISNULL(RSTYYTD,0)-ISNULL(RSLYYTD,0))/ISNULL(RSLYYTD,0))")
                .Add("WXX_VAR_PCT", GetType(System.Decimal), "IIF(ISNULL(LYLXX,0)=0,0,(ISNULL(TYLXX,0)-ISNULL(LYLXX,0))/ISNULL(LYLXX,0))")
                .Add("WXX_LYNL_TREND", GetType(System.Decimal), "IIF(ISNULL(LYLXX,0)=0,0,ISNULL(LYNXX,0)/ISNULL(LYLXX,0))")
                .Add("NXX_QTY", GetType(System.Decimal), "ISNULL(LYNXX,0)*WXX_VAR_PCT")
                .Add("QTY_NEEDED_CALC", GetType(System.Decimal), "ISNULL(NXX_QTY,0)-ISNULL(QTY_EOW,0)-ISNULL(ORDR_QTY_OPEN,0)-ISNULL(SHIPW0,0)-ISNULL(SHIPW1,0)-ISNULL(SHIPW2,0)")
                .Add("QTY_NEEDED", GetType(System.Decimal), "IIF(QTY_NEEDED_CALC<0,0,QTY_NEEDED_CALC)")
                .Add("TY_L3_MOs", GetType(System.Decimal), "ISNULL(RSTYL00,0) + ISNULL(RSTYL01,0) + ISNULL(RSTYL02,0) + ISNULL(RSTYL03,0) + ISNULL(RSTYL04,0) + ISNULL(RSTYL05,0) + ISNULL(RSTYL06,0) + ISNULL(RSTYL07,0) + ISNULL(RSTYL08,0) + ISNULL(RSTYL09,0) + ISNULL(RSTYL10,0) + ISNULL(RSTYL11,0) + ISNULL(RSTYL12,0) + ISNULL(RSTYL13,0)")
                .Add("LY_L3_MOs", GetType(System.Decimal), "ISNULL(RSLYL00,0) + ISNULL(RSLYL01,0) + ISNULL(RSLYL02,0) + ISNULL(RSLYL03,0) + ISNULL(RSLYL04,0) + ISNULL(RSLYL05,0) + ISNULL(RSLYL06,0) + ISNULL(RSLYL07,0) + ISNULL(RSLYL08,0) + ISNULL(RSLYL09,0) + ISNULL(RSLYL10,0) + ISNULL(RSLYL11,0) + ISNULL(RSLYL12,0) + ISNULL(RSLYL13,0)")
                .Add("3MO_TREND_PCT", GetType(System.Decimal), "IIF(ISNULL(LY_L3_MOs,0)=0, 0, ((ISNULL(TY_L3_MOs,0) - ISNULL(LY_L3_MOs,0)) / ISNULL(LY_L3_MOs,0)))")

            End With

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, SREP_CODE, TRADE_CLASS_CODE" & vbCrLf _
                & " from ARTCUST1 where CUST_CODE in (Select Distinct CUST_CODE from " & RSTBOWS1 & ")"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE, ARTCUST2.SELL_CODE" & vbCrLf _
                & ", SOTSELL1.SELL_NAME, SOTSELL1.REGION_CODE, SOTSREG1.REGION_DESC" & vbCrLf _
                & " from ARTCUST2,SOTSELL1,SOTSREG1" & vbCrLf _
                & " where (ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO)" & vbCrLf _
                & " in (Select Distinct CUST_CODE, CUST_STORE_NO from " & RSTBOWS1 & ")" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE"
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "SELECT ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, " & vbCrLf _
               & "ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_CLASS_CODE, " & vbCrLf _
               & "ICTCLAS1.ITEM_CLASS_DESC " & vbCrLf _
               & "FROM ICTITEM1 " & vbCrLf _
               & "INNER JOIN ICTCLAS1 ON ICTITEM1.ITEM_CLASS_CODE = ICTCLAS1.ITEM_CLASS_CODE " & vbCrLf _
               & "WHERE ICTITEM1.ITEM_CODE IN (SELECT DISTINCT ITEM_CODE FROM " & RSTBOWS1 & ")"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from " & RSTBOWS1_CUST_CODE
            Create_TDA(.Tables.Add, "RSTBOWS1_CUST_CODE", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select * from " & RSTBOWS1_CUST_STORE_NO
            Create_TDA(.Tables.Add, "RSTBOWS1_CUST_STORE_NO", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select * from " & RSTBOWS1_ITEM_CODE
            Create_TDA(.Tables.Add, "RSTBOWS1_ITEM_CODE", "**", 0, False, "", 1)

            Create_Relation("RSTBOWS1_CUST_CODE", "RSTBOWS1", "CUST_CODE")
            With .Tables("RSTBOWS1").Columns
                '.Add("CUST_NAME", GetType(System.String), "PARENT(RSTBOWS1_CUST_CODE_RSTBOWS1).CUST_NAME")
                .Add("SREP_CODE", GetType(System.String), "PARENT(RSTBOWS1_CUST_CODE_RSTBOWS1).SREP_CODE")
                .Add("TRADE_CLASS_CODE", GetType(System.String), "PARENT(RSTBOWS1_CUST_CODE_RSTBOWS1).TRADE_CLASS_CODE")
            End With

            Create_Relation("RSTBOWS1_CUST_STORE_NO", "RSTBOWS1", "CUST_CODE,CUST_STORE_NO")
            With .Tables("RSTBOWS1").Columns
                .Add("CUST_STORE_NAME", GetType(System.String), "PARENT(RSTBOWS1_CUST_STORE_NO_RSTBOWS1).CUST_STORE_NAME")
                .Add("CUST_STORE_STATE", GetType(System.String), "PARENT(RSTBOWS1_CUST_STORE_NO_RSTBOWS1).CUST_STORE_STATE")
                .Add("SELL_CODE", GetType(System.String), "PARENT(RSTBOWS1_CUST_STORE_NO_RSTBOWS1).SELL_CODE")
                .Add("SELL_NAME", GetType(System.String), "PARENT(RSTBOWS1_CUST_STORE_NO_RSTBOWS1).SELL_NAME")
                .Add("REGION_CODE", GetType(System.String), "PARENT(RSTBOWS1_CUST_STORE_NO_RSTBOWS1).REGION_CODE")
                .Add("REGION_DESC", GetType(System.String), "PARENT(RSTBOWS1_CUST_STORE_NO_RSTBOWS1).REGION_DESC")
            End With

            Create_Relation("RSTBOWS1_ITEM_CODE", "RSTBOWS1", "ITEM_CODE")
            With .Tables("RSTBOWS1").Columns
                .Add("ITEM_DESC", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).ITEM_DESC")
                .Add("ITEM_RETAIL_PRICE", GetType(System.Decimal), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).ITEM_RETAIL_PRICE")
                .Add("COLLECTION_CODE", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).COLLECTION_CODE")
                .Add("PROD_CODE", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).PROD_CODE")
                .Add("COLLECTION_GENDER", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).COLLECTION_GENDER")
                .Add("ITEM_EAN_CODE", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).ITEM_EAN_CODE")
                .Add("ITEM_ALT_SORT", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).ITEM_ALT_SORT")
                '.Add("ITEM_CLASS_CODE", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).ITEM_CLASS_CODE")
                .Add("ITEM_CLASS_DESC", GetType(System.String), "PARENT(RSTBOWS1_ITEM_CODE_RSTBOWS1).ITEM_CLASS_DESC")

            End With

            With .Tables.Add("RSTBOWSC")
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_INDEX", GetType(System.Int32))
                .Columns.Add("COLUMN_NAME_LENGTH", GetType(System.Int32))
            End With
        End With

        EnforceConstraints(False)

        Fill_Records("RSTBOWS1")
        grdRSTBOWS1.DataSource = dst.Tables("RSTBOWS1")

        For Each DCOL As DataColumn In dst.Tables("RSTBOWS1").Columns
            If DCOL.DataType.ToString = "System.String" Then
            ElseIf DCOL.DataType.ToString = "System.DateTime" Then
                grdRSTBOWS1.DisplayLayout.Bands(0).Columns(DCOL.ColumnName).Format = "MM/dd/yy"
            Else
                grdRSTBOWS1.DisplayLayout.Bands(0).Columns(DCOL.ColumnName).Format = "#,##0"
            End If
        Next

        Fill_Records("RSTBOWS1_CUST_CODE")
        grdRSTBOWS1_CUST_CODE.DataSource = dst.Tables("RSTBOWS1_CUST_CODE")
        Sort_grdColumns(grdRSTBOWS1_CUST_CODE, "CUST_CODE")
        Initialize_grd(grdRSTBOWS1_CUST_CODE)

        Dim WKS As Integer = Val(numWEEKS.Value & "")
        Dim P As Integer = 0
        dst.Tables("RSTBOWSC").Rows.Clear()

        hide_column = False

        With grdRSTBOWS1.DisplayLayout.Bands(0)
            '.Columns("QTY_NEEDED_CALC").Hidden = True
            .Columns("WTD_VAR_PCT").Hidden = True

            P = -1
            grdfmt(.Columns("CUST_CODE"), P, 100, "Customer", , , True)
            grdfmt(.Columns("CUST_STORE_NO"), P, 70, "Store")
            grdfmt(.Columns("CUST_STORE_NAME"), P, 160, "Store Name")
            grdfmt(.Columns("CUST_STORE_STATE"), P, 40, "St")

            grdfmt(.Columns("SELL_CODE"), P, 60, "AE", Color.Gold)
            grdfmt(.Columns("SELL_NAME"), P, 60, "AE Name", Color.Gold)
            grdfmt(.Columns("REGION_CODE"), P, 60, "RAM Code", Color.Gold)
            grdfmt(.Columns("REGION_DESC"), P, 60, "RAM Name", Color.Gold)
            grdfmt(.Columns("SREP_CODE"), P, 60, "KAM", Color.Gold)
            grdfmt(.Columns("TRADE_CLASS_CODE"), P, 60, "TC", Color.Gold)

            grdfmt(.Columns("COLLECTION_CODE"), P, 60, "Coll", Color.LightBlue)
            grdfmt(.Columns("PROD_CODE"), P, 60, "Prod Cd", Color.LightBlue)
            grdfmt(.Columns("ITEM_CLASS_DESC"), P, 160, "Item Class Desc", Color.LightBlue)
            grdfmt(.Columns("COLLECTION_GENDER"), P, 70, "Coll Gender", Color.LightBlue)


            grdfmt(.Columns("ITEM_CODE"), P, 160, "Item Code", Color.LightBlue)

            grdfmt(.Columns("ITEM_ALT_SORT"), P, 70, "3PL Code", Color.LightBlue)

            grdfmt(.Columns("ITEM_DESC"), P, 120, "Item Description", Color.LightBlue)
            grdfmt(.Columns("ITEM_EAN_CODE"), P, 110, "EAN", Color.LightBlue)
            grdfmt(.Columns("ITEM_RETAIL_PRICE"), P, 70, "MSRP", Color.LightBlue, "###.00")

            grdfmt(.Columns("QTY_EOW"), P, 70, "EOW OH", Color.Violet, "#,##0")
            grdfmt(.Columns("ORDR_QTY_OPEN"), P, 70, "Opn+Pck", Color.Violet, "#,##0")
            grdfmt(.Columns("INV_DATE"), P, 80, "Last Ship", Color.Violet, "MM/dd/yy")

            ' THESE VARIABLES WERE USED TO FACILLITATE HEADNIG CHANGES AS PER SB
            Dim heading As String = ""
            Dim headingLY = Format(Val(Mid(RYW, 1, 4)) - 1 * LYX, "0000")

            If sole_AE <> "" Or sole_REGION <> "" Then
                heading = "3WKS in PO transit" : grdfmt(.Columns("INXIT_3WKS"), P, 200, heading, Color.LightGreen, "#,##0")
                heading = "LY Next 10 wks SLS" : heading = headingLY & " N10 wks SLS" : grdfmt(.Columns("LYN10"), P, 200, heading, Color.LightGreen, "#,##0")
                heading = "TY Last 4wks SLS" : grdfmt(.Columns("TYL4"), P, 200, heading, Color.LightGreen, "#,##0")
                heading = "LY Last 4wks SLS" : heading = headingLY & " Last 4wks SLS" : grdfmt(.Columns("LYL4"), P, 200, heading, Color.LightGreen, "#,##0")
                heading = "L4 % LY" : heading = "L4 % " & headingLY : grdfmt(.Columns("L4TYLY"), P, 150, heading, Color.LightGreen, "#,##0.0%")
                heading = "TY Last 10wks SLS" : grdfmt(.Columns("TYL10"), P, 200, heading, Color.LightGreen, "#,##0")
                heading = "LY Last 10wks SLS" : heading = headingLY & " Last 10wks SLS" : grdfmt(.Columns("LYL10"), P, 200, heading, Color.LightGreen, "#,##0")
                heading = "L10 % LY" : heading = "L10 % " & headingLY : grdfmt(.Columns("L10TYLY"), P, 150, heading, Color.LightGreen, "#,##0.0%")
            End If

            If sole_AE <> "" Or sole_REGION <> "" Then
                hide_column = True
            End If

            For w As Integer = 0 To 5
                Dim C As String = "SHIPW" & Format(w, "0")
                grdfmt(.Columns(C), P, 70, "Shp L" & Format(w, "0"), Color.Violet, "#,##0")
            Next
            grdfmt(.Columns("WXX_LYNL_TREND"), P, 70, "LYNL>", Color.Tan, "#,##0.0")
            grdfmt(.Columns("WXX_VAR_PCT"), P, 70, CStr(WKS) & "Wk Var", Color.Tan, "#,##0.0")
            grdfmt(.Columns("NXX_QTY"), P, 70, CStr(WKS) & "Wk Need", Color.Tan, "#,##0")
            grdfmt(.Columns("QTY_NEEDED_CALC"), P, 1, "Calc", Color.Tan, "#,##0")
            grdfmt(.Columns("QTY_NEEDED"), P, 70, "Boost", Color.Tan, "#,##0")
            For w As Integer = 0 To 13
                ' Dim YW As String = ASCMAIN1.Week_Calc(RYW, -1 * w)
                Dim C As String = "RSTYL" & Format(w, "00")
                grdfmt(.Columns(C), P, 70, "TY-L" & Format(w, "00"), Color.LightGreen, "#,##0")
            Next
            For w As Integer = 0 To 13
                Dim C As String = "RSLYL" & Format(w, "00")
                grdfmt(.Columns(C), P, 70, "LY-L" & Format(w, "00"), Color.Yellow, "#,##0")
            Next
            For w As Integer = 1 To 13
                Dim C As String = "RSLYN" & Format(w, "00")
                grdfmt(.Columns(C), P, 70, "LY-N" & Format(w, "00"), Color.LightPink, "#,##0")
            Next

            grdfmt(.Columns("RSTYMTD"), P, 70, "TY-MTD", Color.LightGreen, "#,##0")
            grdfmt(.Columns("RSLYMTD"), P, 70, "LY-MTD", Color.LightGreen, "#,##0")
            grdfmt(.Columns("MTD_VAR_PCT"), P, 70, "Mtd-Var%", Color.LightGreen, "#,##0.0")
            grdfmt(.Columns("TY_L3_MOs"), P, 70, "TY-L3-MOs", Color.Aquamarine, "#,##0")
            grdfmt(.Columns("LY_L3_MOs"), P, 70, "LY-L3-MOs", Color.Aquamarine, "#,##0")
            grdfmt(.Columns("3MO_TREND_PCT"), P, 70, "3Mo-%", Color.Aquamarine, "#,##0.0")
            If sole_AE <> "" Or sole_REGION <> "" Then
                hide_column = False
            End If

            heading = "TY-STD" : grdfmt(.Columns("RSTYSTD"), P, 70, heading, Color.LightBlue, "#,##0")
            heading = "LY-STD" : heading = headingLY & "-STD" : grdfmt(.Columns("RSLYSTD"), P, 70, heading, Color.LightBlue, "#,##0")
            heading = "Std-Var%" : heading = "Std-Var% " & headingLY : grdfmt(.Columns("STD_VAR_PCT"), P, 70, heading, Color.LightBlue, "#,##0.0")

            Dim sYW As String = ""
            sYW = RYW
            For W As Integer = 1 To 10
                sYW = ASCMAIN1.Week_Calc(sYW, -1)
                Dim rowWeek As DataRow = LookUp("GLTPARM3", sYW)
                grdfmt(.Columns("OHTYL" & Format(W, "00")), P, 70, "OH " & Format(rowWeek.Item("WEEK_END_DATE"), "MM/dd/yy"), Color.LightGreen, "#,##0")
            Next
            sYW = ASCMAIN1.Week_Calc(RYW, -52 * LYX)
            For w As Integer = 1 To 10
                sYW = ASCMAIN1.Week_Calc(sYW, -1)
                Dim rowWeek As DataRow = LookUp("GLTPARM3", sYW)
                grdfmt(.Columns("OHLYL" & Format(w, "00")), P, 70, "OH " & Format(rowWeek.Item("WEEK_END_DATE"), "MM/dd/yy"), Color.Yellow, "#,##0")
            Next
            sYW = ASCMAIN1.Week_Calc(RYW, -52 * LYX)
            For w As Integer = 1 To 10
                sYW = ASCMAIN1.Week_Calc(sYW, +1)
                Dim rowWeek As DataRow = LookUp("GLTPARM3", sYW)
                grdfmt(.Columns("OHLYN" & Format(w, "00")), P, 70, "OH " & Format(rowWeek.Item("WEEK_END_DATE"), "MM/dd/yy"), Color.LightPink, "#,##0")
            Next

            If sole_AE <> "" Or sole_REGION <> "" Then
                hide_column = True
            End If

            grdfmt(.Columns("RSTYYTD"), P, 70, "TY-YTD", Color.LimeGreen, "#,##0")
            grdfmt(.Columns("RSLYYTD"), P, 70, "LY-YTD", Color.LimeGreen, "#,##0")
            grdfmt(.Columns("YTD_VAR_PCT"), P, 70, "Ytd-Var%", Color.LimeGreen, "#,##0.0")

            grdfmt(.Columns("TYLXX"), P, 70, "TY-L" & Format(WKS, "00"), Color.Orange, "#,##0")
            grdfmt(.Columns("LYLXX"), P, 70, "LY-L" & Format(WKS, "00"), Color.Orange, "#,##0")
            grdfmt(.Columns("LYNXX"), P, 70, "LY-N" & Format(WKS, "00"), Color.Orange, "#,##0")

            If sole_AE <> "" Or sole_REGION <> "" Then
                hide_column = False
            End If

        End With

        Fill_Records("RSTBOWS1_CUST_STORE_NO")
        grdRSTBOWS1_CUST_STORE_NO.DataSource = dst.Tables("RSTBOWS1_CUST_STORE_NO")
        Sort_grdColumns(grdRSTBOWS1_CUST_STORE_NO, "CUST_CODE, CUST_STORE_NO")

        Fill_Records("RSTBOWS1_ITEM_CODE")
        grdRSTBOWS1_ITEM_CODE.DataSource = dst.Tables("RSTBOWS1_ITEM_CODE")
        Sort_grdColumns(grdRSTBOWS1_ITEM_CODE, "ITEM_CODE")

        Fill_Records("ARTCUST1")
        grdARTCUST1.DataSource = dst.Tables("ARTCUST1")
        Setup_grdRSTBOWS1()

        With grdARTCUST1.DisplayLayout.Bands(0)
            P = -1
            grdfmt(.Columns("CUST_CODE"), P, 100, "Customer", , True)
            grdfmt(.Columns("CUST_NAME"), P, 120, "Customer Name")
            grdfmt(.Columns("SREP_CODE"), P, 60, "KAM")
            grdfmt(.Columns("TRADE_CLASS_CODE"), P, 60, "TC")
        End With

        Fill_Records("ARTCUST2")
        Fill_Records("ICTITEM1")

        EnforceConstraints(True)

        tabMain.Visible = True

        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Other Run-Time Options")

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Creating Workbook")

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim FILENAME As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FILENAME = xls_path & "\" & xls_name & ".XLSx"

                If Not My.Computer.FileSystem.FileExists(FILENAME) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        Dim oWB As SpreadsheetGear.IWorkbook
        oWB = SpreadsheetGear.Factory.GetWorkbook()

        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i

        Dim SI As Integer = 0
        Dim oSheet As SpreadsheetGear.IWorksheet

        oSheet = oWB.Sheets(SI)
        oSheet.Name = "Parameters"
        Dim tbl As New DataTable
        With tbl.Columns
            .Add("RTL_WEEK", GetType(System.String))
            .Add("WEEK_END_DATE", GetType(System.DateTime))
            .Add("WEEKS_QTY_NEEDED", GetType(System.Int32))
        End With
        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        tbl.Rows.Add(New Object() {RYW, rowGLTPARM3.Item("WEEK_END_DATE"), numWEEKS.Value})
        Load_DataTable_into_SGXLS(1, 1, tbl, oSheet)

        For Each grow As UltraWinGrid.UltraGridRow In grdARTCUST1.Rows
            grow.Activate()
            SI += 1
            Dim CUST_CODE As String = grow.Cells("CUST_CODE").Value
            grdRSTBOWS1.Text = CUST_CODE
            Application.DoEvents()

            oSheet = oWB.Worksheets.Add

            Dim dvw As DataView = dst.Tables("RSTBOWS1").DefaultView
            dvw.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"
            dvw.Sort = "ITEM_CODE,CUST_STORE_NO"
            tbl = dvw.ToTable

            oSheet.Name = CUST_CODE

            Dim tbl3 As DataTable = Load_DataTable_into_SGXLS(1, 1, tbl, oSheet, grdRSTBOWS1)

            For Each rowRSTBOWSC As DataRow In dst.Tables("RSTBOWSC").Select("", "COLUMN_NAME_LENGTH DESC")
                Dim COLUMN_NAME As String = rowRSTBOWSC.Item("COLUMN_NAME")
                'If COLUMN_NAME.StartsWith("NXX") Or COLUMN_NAME.StartsWith("QTY_NEED") Then Stop
                Dim COLUMN_INDEX As Integer = Val(rowRSTBOWSC.Item("COLUMN_INDEX") & "")
                If dst.Tables("RSTBOWS1").Columns(COLUMN_NAME).Expression <> "" Then
                    If dst.Tables("RSTBOWS1").Columns(COLUMN_NAME).DataType.ToString = "System.String" _
                    Or New String() {"ITEM_RETAIL_PRICE"}.Contains(COLUMN_NAME) Then
                    Else
                        If (sole_AE <> "" Or sole_REGION <> "") And COLUMN_NAME <> "L4TYLY" And COLUMN_NAME <> "L10TYLY" Then ' And COLUMN_NAME <> "STD_VAR_PCT" Then
                            ' no formulae for this version
                        Else
                            Dim FX As String = Make_Formula(COLUMN_NAME)
                            Dim XCX As String = Excel_Cell(2, COLUMN_INDEX)
                            Dim XCY As String = Excel_Cell(1 + tbl.Select.Length, COLUMN_INDEX)
                            oSheet.Cells(XCX).Formula = FX
                            oSheet.Cells(XCX).Copy(oSheet.Range(XCX & ":" & XCY))
                        End If
                    End If
                End If
            Next

            With oSheet.Range(0, 0, 0, tbl3.Columns.Count - 1)
                .AutoFilter()
            End With

            If sole_AE <> "" Or sole_REGION <> "" Then
                oSheet.Cells("M1").EntireColumn.Delete() ' M
                oSheet.Cells("J1").EntireColumn.Delete() ' J
                oSheet.Cells("I1").EntireColumn.Delete() ' I
                oSheet.Cells("G1").EntireColumn.Delete() ' G
                oSheet.Cells("E1").EntireColumn.Delete() ' E
            End If
        Next

        If sole_AE = "" And sole_REGION = "" Then
            oSheet = oWB.Worksheets.Add
            oSheet.Name = "Customers"
            Load_DataTable_into_SGXLS(1, 1, dst.Tables("RSTBOWS1_CUST_CODE"), oSheet, grdRSTBOWS1_CUST_CODE, Nothing, "CUST_CODE")
            oSheet = oWB.Worksheets.Add
            oSheet.Name = "Stores"
            Load_DataTable_into_SGXLS(1, 1, dst.Tables("RSTBOWS1_CUST_STORE_NO"), oSheet, grdRSTBOWS1_CUST_STORE_NO, Nothing, "CUST_CODE,CUST_STORE_NO")
            oSheet = oWB.Worksheets.Add
            oSheet.Name = "Items"
            Load_DataTable_into_SGXLS(1, 1, dst.Tables("RSTBOWS1_ITEM_CODE"), oSheet, grdRSTBOWS1_ITEM_CODE, Nothing, "ITEM_CODE")
        End If

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        '   Show_Document(FILENAME)
        oWB = Nothing

        Add_Document_to_ASTSPRF1(FILENAME)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")

    End Sub

    Sub grdfmt(gcol As UltraWinGrid.UltraGridColumn, _
               ByRef Position As Integer, _
               Width As Integer, _
               Caption As String, _
               Optional colColor As System.Drawing.Color = Nothing, _
               Optional Mask As String = "", _
               Optional Fixed As Boolean = False)


        'P = -1
        'grdfmt(.Columns("CUST_CODE"), P, 100, "Customer", , , True)
        'grdfmt(.Columns("CUST_STORE_NO"), P, 70, "Store")
        'grdfmt(.Columns("ITEM_CODE"), P, 160, "Item Code", Color.LightBlue)
        'grdfmt(.Columns("ITEM_DESC"), P, 120, "Item Description", Color.LightBlue)
        'grdfmt(.Columns("ITEM_RETAIL_PRICE"), P, 70, "MSRP", Color.LightBlue, "###.00")
        'grdfmt(.Columns("QTY_EOW"), P, 70, "EOW OH", Color.Violet, "#,##0")
        'grdfmt(.Columns("INV_DATE"), P, 80, "Last Ship", Color.Violet, "MM/dd/yy")

        If gcol.Band.Key = "RSTBOWS1" Then
            dst.Tables("RSTBOWSC").Rows.Add(New Object() {gcol.Key, dst.Tables("RSTBOWSC").Rows.Count + 1, gcol.Key.Length})
        End If

        With gcol
            Position += 1 : .Header.SetVisiblePosition(Position, False)
            .Width = Width
            .Header.Caption = Caption
            If Mask <> "" Then .Format = Mask
            If Fixed Then .Header.Fixed = Fixed

            If colColor = Nothing Then
                colColor = System.Drawing.Color.LightGray
            End If

            With .Header.Appearance
                .BackColor = Color.White
                .BackGradientStyle = GradientStyle.ForwardDiagonal
                .BackColor2 = colColor
            End With

            If hide_column Then
                .Hidden = True
                '.Width = 1
            End If
        End With
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYW").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Week"
            End If
        End If
    End Sub

    Sub Create_Work_File()

        Get_SQL("*")

        Dim LYX As Integer = 1
        LYX = Val(cbeLY.Value & "")

        Dim sql_sole_AE As String = ""
        Dim sql_sole_AE_items As String = ""
        Dim sql_sole_AE_SP_Items As String = ""

        If sole_AE <> "" Or sole_REGION <> "" Then
            If sole_REGION <> "" Then
                sql_sole_AE = " and ARTCUST2.SELL_CODE in (Select SELL_CODE from SOTSELL1 where REGION_CODE = '" & sole_REGION & "')"
            Else
                sql_sole_AE = " and ARTCUST2.SELL_CODE = '" & sole_AE & "'"
            End If

            ' DATE_START IS PRIOR OR EQUAL TO END OF SEASON AND DATE_END IS LATER THAN OR EQUAL TO BEG OF SEASON
            Dim sqlAI As String = "Select Distinct ITEM_CODE from SOTALLO1" & vbCrLf _
                                  & " where DATE_START <= '" & Format(SN_END, "dd-MMM-yyyy") & "'" & vbCrLf _
                                  & "   and DATE_END >= '" & Format(SN_START, "dd-MMM-yyyy") & "'"
            sql_sole_AE_items = " and ICTITEM1.ITEM_SNU_CODE = 'S' and (ICTITEM1.ITEM_BASIC_PROMO = 'B' or ICTITEM1.ITEM_CODE in (" & sqlAI & "))"
            sql_sole_AE_SP_Items = " and ICTITEM1.ITEM_SNU_CODE = 'S' and ICTITEM1.ITEM_BASIC_PROMO = 'P' and ICTITEM1.ITEM_CODE in (" & sqlAI & ")"

        End If

        Dim YP1 As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12 * LYX)
        ASCMAIN1.sql = "" _
            & "Select Distinct RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, RSTRETL1.ITEM_CODE from RSTRETL1" & vbCrLf _
            & sql_TABLE_NAMEs & ASCMAIN1.SQL_Add_WHERE(sql_JOIN & sql_WHERE & sql_sole_AE & sql_sole_AE_items & $" and RSTRETL1.OPS_YYYYPP >='{YP1}'") & vbCrLf _
            & " union " & vbCrLf _
            & "Select Distinct SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE from SOTINVH2" & vbCrLf _
            & sql_TABLE_NAMEs & ASCMAIN1.SQL_Add_WHERE(Replace(sql_JOIN & sql_WHERE & $" AND SOTINVH2.INV_TYPE = 'I' and SOTINVH2.ORDR_YYYYPP_UPDATED >='{YP1}'" & sql_sole_AE & sql_sole_AE_items, "RSTRETL1", "SOTINVH2"))

        RSTBOWS1 = ASCMAIN1.Temp_Table

        Dim sqlCUST_CODE As String = "Select Distinct CUST_CODE from " & RSTBOWS1
        Dim sqlCUST_STORE_NO As String = "SELECT DISTINCT CUST_STORE_NO FROM " & RSTBOWS1

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE, SREP_CODE" & vbCrLf _
            & " from ARTCUST1 where CUST_CODE in (" & sqlCUST_CODE & ")"
        RSTBOWS1_CUST_CODE = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTBOWS1_CUST_CODE & " Add Primary Key (CUST_CODE)")

        ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_STATE, ARTCUST2.SELL_CODE" & vbCrLf _
            & ", SOTSELL1.SELL_NAME, SOTSELL1.REGION_CODE, SOTSREG1.REGION_DESC" & vbCrLf _
            & " from ARTCUST2,SOTSELL1,SOTSREG1" & vbCrLf _
            & " where CUST_CODE in (" & sqlCUST_CODE & ")" & vbCrLf _
            & " and CUST_STORE_NO in (" & sqlCUST_STORE_NO & ")" & vbCrLf _
            & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
            & IIf(sole_AE <> "", sql_sole_AE, "") _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSELL1.REGION_CODE"
        RSTBOWS1_CUST_STORE_NO = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTBOWS1_CUST_STORE_NO & " Add Primary Key (CUST_CODE, CUST_STORE_NO)")

        ASCMAIN1.sql = "Select Distinct ITEM_CODE from " & RSTBOWS1
        If sole_AE <> "" Or sole_REGION <> "" Then
            ASCMAIN1.sql &= " union Select ITEM_CODE from ICTITEM1 " & ASCMAIN1.SQL_Add_WHERE(sql_sole_AE_SP_Items)
        End If
        'ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, ICTITEM1.COLLECTION_CODE, PROD_CODE, COLLECTION_GENDER,ITEM_EAN_CODE, ITEM_ALT_SORT, ITEM_CLASS_CODE" & vbCrLf _
        '    & " from ICTITEM1,ICTCOLL1 where ICTCOLL1.COLLECTION_CODE(+) = ICTITEM1.COLLECTION_CODE AND ITEM_CODE in (" & ASCMAIN1.sql & ")"
        'RSTBOWS1_ITEM_CODE = ASCMAIN1.Temp_Table
        'ASCDATA1.ExecuteSQL("Alter Table " & RSTBOWS1_ITEM_CODE & " Add Primary Key (ITEM_CODE)")
        ASCMAIN1.sql = "SELECT ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, " & vbCrLf _
               & "ICTITEM1.COLLECTION_CODE, ICTITEM1.PROD_CODE, ICTCOLL1.COLLECTION_GENDER, " & vbCrLf _
               & "ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_ALT_SORT, ICTITEM1.ITEM_CLASS_CODE, " & vbCrLf _
               & "ICTCLAS1.ITEM_CLASS_DESC " & vbCrLf _
               & "FROM ICTITEM1 " & vbCrLf _
               & "INNER JOIN ICTCOLL1 ON ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE " & vbCrLf _
               & "LEFT JOIN ICTCLAS1 ON ICTITEM1.ITEM_CLASS_CODE = ICTCLAS1.ITEM_CLASS_CODE " & vbCrLf _
               & "WHERE ICTITEM1.ITEM_CODE IN (" & ASCMAIN1.sql & ")"
        RSTBOWS1_ITEM_CODE = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("ALTER TABLE " & RSTBOWS1_ITEM_CODE & " ADD PRIMARY KEY (ITEM_CODE)")


        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        Dim RYP As String = rowGLTPARM3("YYYYPP")

        Dim sqlTYL13 As String = ""
        Dim sqlTYL13_0 As String = ""
        Dim sqlTYL13_sum As String = ""
        For w As Integer = 0 To 13
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -1 * w)
            sqlTYL13 &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',QTY_SOLD,0)) RSTYL" & Format(w, "00") & vbCrLf
            sqlTYL13_0 &= ", 0 RSTYL" & Format(w, "00")
            sqlTYL13_sum &= ", Sum (RSTYL" & Format(w, "00") & ") RSTYL" & Format(w, "00") & vbCrLf
        Next
        For w As Integer = 1 To 10
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -1 * w)
            sqlTYL13 &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',QTY_EOW,0)) OHTYL" & Format(w, "00") & vbCrLf
            sqlTYL13_0 &= ", 0 OHTYL" & Format(w, "00")
            sqlTYL13_sum &= ", Sum (OHTYL" & Format(w, "00") & ") OHTYL" & Format(w, "00") & vbCrLf
        Next

        Dim sqlLYL13 As String = ""
        Dim sqlLYL13_0 As String = ""
        Dim sqlLYL13_sum As String = ""
        For w As Integer = 0 To 13
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -52 * LYX - 1 * w)
            sqlLYL13 &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',QTY_SOLD,0)) RSLYL" & Format(w, "00") & vbCrLf
            sqlLYL13_0 &= ", 0 RSLYL" & Format(w, "00")
            sqlLYL13_sum &= ", Sum (RSLYL" & Format(w, "00") & ") RSLYL" & Format(w, "00") & vbCrLf
        Next
        For w As Integer = 1 To 10
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -52 * LYX - 1 * w)
            sqlLYL13 &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',QTY_EOW,0)) OHLYL" & Format(w, "00") & vbCrLf
            sqlLYL13_0 &= ", 0 OHLYL" & Format(w, "00")
            sqlLYL13_sum &= ", Sum (OHLYL" & Format(w, "00") & ") OHLYL" & Format(w, "00") & vbCrLf
        Next

        Dim sqlLYN13 As String = ""
        Dim sqlLYN13_0 As String = ""
        Dim sqlLYN13_sum As String = ""
        For w As Integer = 1 To 13
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -52 * LYX + 1 * w)
            sqlLYN13 &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',QTY_SOLD,0)) RSLYN" & Format(w, "00") & vbCrLf
            sqlLYN13_0 &= ", 0 RSLYN" & Format(w, "00")
            sqlLYN13_sum &= ", Sum (RSLYN" & Format(w, "00") & ") RSLYN" & Format(w, "00") & vbCrLf
        Next
        For w As Integer = 1 To 10
            Dim YW As String = ASCMAIN1.Week_Calc(RYW, -52 * LYX + 1 * w)
            sqlLYN13 &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',QTY_EOW,0)) OHLYN" & Format(w, "00") & vbCrLf
            sqlLYN13_0 &= ", 0 OHLYN" & Format(w, "00")
            sqlLYN13_sum &= ", Sum (OHLYN" & Format(w, "00") & ") OHLYN" & Format(w, "00") & vbCrLf
        Next

        Dim RTYW_01 As String = Mid(RYW, 1, 4) & "01"
        Dim RTYW_01_STD As String = Mid(RYW, 1, 4) & "01"
        If Mid(RYW, 5, 2) >= "27" Then
            RTYW_01_STD = Mid(RTYW_01_STD, 1, 4) & "27"
        End If


        Dim sqlTYMSY As String = "" _
        & ", Sum (Decode(OPS_YYYYPP,'" & RYP & "',QTY_SOLD,0)) RSTYMTD" & vbCrLf _
        & ", Sum (Case when OPS_YYYYWW between '" & RTYW_01_STD & "' and '" & RYW & "' THEN QTY_SOLD ELSE 0 END) RSTYSTD" & vbCrLf _
        & ", Sum (Case when OPS_YYYYWW between '" & Mid(RYW, 1, 4) & "01" & "' and '" & RYW & "' THEN QTY_SOLD ELSE 0 END) RSTYYTD" & vbCrLf
        Dim sqlTYMSY_0 As String = ", 0 RSTYMTD, 0 RSTYSTD, 0 RSTYYTD"
        Dim sqlTYMSY_sum As String = ", Sum (RSTYMTD) RSTYMTD, Sum (RSTYSTD) RSTYSTD, Sum (RSTYYTD) RSTYYTD"

        ' NOTE - WE ARE NOT DOING -52 WEEKS FOR LY
        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -12 * LYX)
        Dim LYW As String = Format(Val(Mid(RYW, 1, 4)) - 1 * LYX, "0000") & Mid(RYW, 5, 2)
        Dim RLYW_01 As String = Format(Val(Mid(RYW, 1, 4)) - 1 * LYX, "0000") & "01"
        Dim RLYW_01_STD As String = Format(Val(Mid(RYW, 1, 4)) - 1 * LYX, "0000") & "01"
        If Mid(RYW, 5, 2) >= "27" Then
            RLYW_01_STD = Mid(RLYW_01_STD, 1, 4) & "27"
        End If

        Dim sqlLYMSY As String = "" _
        & ", Sum (Decode(OPS_YYYYPP,'" & LYP & "',QTY_SOLD,0)) RSLYMTD" & vbCrLf _
        & ", Sum (Case when OPS_YYYYWW between '" & RLYW_01_STD & "' and '" & LYW & "' THEN QTY_SOLD ELSE 0 END) RSLYSTD" & vbCrLf _
        & ", Sum (Case when OPS_YYYYWW between '" & Mid(LYW, 1, 4) & "01" & "' and '" & LYW & "' THEN QTY_SOLD ELSE 0 END) RSLYYTD" & vbCrLf
        Dim sqlLYMSY_0 As String = ", 0 RSLYMTD, 0 RSLYSTD, 0 RSLYYTD"
        Dim sqlLYMSY_sum As String = ", Sum (RSLYMTD) RSLYMTD, Sum (RSLYSTD) RSLYSTD, Sum (RSLYYTD) RSLYYTD"

        Dim sqlTYLXX_sum As String = ""
        Dim sqlLYLXX_sum As String = ""
        Dim sqlLYNXX_sum As String = ""
        Dim NUMWEEKS As Integer = Val(Absx1.numFor("NUMWEEKS").Value)
        For w As Integer = 1 To NUMWEEKS
            sqlTYLXX_sum &= "+ NVL(RSTYL" & Format(w, "00") & ",0)"
            sqlLYLXX_sum &= "+ NVL(RSLYL" & Format(w, "00") & ",0)"
            If w = 1 Then
                sqlLYNXX_sum &= "+ NVL(RSLYL00,0)"
            Else
                sqlLYNXX_sum &= "+ NVL(RSLYN" & Format(w - 1, "00") & ",0)"
            End If
        Next
        sqlTYLXX_sum = ", Sum (" & Mid(sqlTYLXX_sum, 3) & ") TYLXX"
        sqlLYLXX_sum = ", Sum (" & Mid(sqlLYLXX_sum, 3) & ") LYLXX"
        sqlLYNXX_sum = ", Sum (" & Mid(sqlLYNXX_sum, 3) & ") LYNXX"

        Dim RYW_ONH As String = RYW
        If RYW = ASCMAIN1.CYW Then
            RYW_ONH = ASCMAIN1.Week_Calc(RYW, -1)
        End If

        ' SPECIAL STARTING WEEK BECAUSE THE BLOCK THAT PREPARES TY-STD ALSO PREPARES PREVIOUS 13 WEEKS
        ' THIS HAS RESULTED IN 2 BUGS FOUND AT VARIOUS PARTS OF THE YEAR, AND THAT IS WHY THERE ARE 2 ADJUSTMENTS TO TYW1
        ' PROBABLY OUGHT TO HAVE SACRIFICED PERFORMANCE AND SEPARATED THAT QUERY INTO 2 PARTS TO AVOID THE TROUBLES
        Dim TYW1 As String = ASCMAIN1.Week_Calc(RYW, -13 + 1) ' RTYW_01
        If RYW_ONH < TYW1 Then
            TYW1 = RYW_ONH
        End If
        If RTYW_01 < TYW1 Then
            TYW1 = RTYW_01
        End If

        ' NOTE THAT THE BLOCK THAT PREPARES LY-STD DOES NOT NEED TO PREPARE ANY OF THE 13 WEEKS NUMBERS SINCE THAT IS TAKEN CARE OF BY A SPECIAL BLOCK A LITTLE LOWER
        'Dim LYW1 As String = ASCMAIN1.Week_Calc(LYW, -13 + 1) ' LTYW_01
        'If RLYW_01 < LYW1 Then
        '    LYW1 = RLYW_01
        'End If

        Dim sql_sole_AE_SP As String = ""
        If sql_sole_AE <> "" Then
            sql_sole_AE_SP = vbCrLf _
            & " union " & vbCrLf _
            & "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ICTITEM1.ITEM_CODE" & vbCrLf _
            & ", 0 QTY_EOW, 0 ORDR_QTY_OPEN, NULL INV_DATE" & vbCrLf _
            & ", 0 SHIPW0, 0 SHIPW1, 0 SHIPW2, 0 SHIPW3, 0 SHIPW4, 0 SHIPW5" & vbCrLf _
            & sqlTYL13_0 & sqlLYL13_0 & sqlLYN13_0 & sqlTYMSY_0 & sqlLYMSY_0 _
            & " from ARTCUST2,ICTITEM1 where " & Mid(sql_sole_AE, 5) & sql_sole_AE_SP_Items
        End If

        ASCMAIN1.sql = "Select X.CUST_CODE, X.CUST_STORE_NO, X.ITEM_CODE" & vbCrLf _
            & ", Sum (X.QTY_EOW) QTY_EOW, Sum (X.ORDR_QTY_OPEN) ORDR_QTY_OPEN, MAX(X.INV_DATE) INV_DATE" & vbCrLf _
            & ", Sum (X.SHIPW0) SHIPW0, Sum (X.SHIPW1) SHIPW1, Sum (X.SHIPW2) SHIPW2, Sum (X.SHIPW3) SHIPW3, Sum (X.SHIPW4) SHIPW4, Sum (X.SHIPW5) SHIPW5" & vbCrLf _
            & sqlTYL13_sum & sqlLYL13_sum & sqlLYN13_sum & sqlTYMSY_sum & sqlLYMSY_sum _
            & sqlTYLXX_sum & sqlLYLXX_sum & sqlLYNXX_sum _
            & " from ARTCUST2, SATAUTH1, ICTITEM1, ICTCOLL1, (" & vbCrLf _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & RYW_ONH & "',QTY_EOW,0)) QTY_EOW, 0 ORDR_QTY_OPEN, NULL INV_DATE" & vbCrLf _
            & ", 0 SHIPW0, 0 SHIPW1, 0 SHIPW2, 0 SHIPW3, 0 SHIPW4, 0 SHIPW5" & vbCrLf _
            & sqlTYL13 & sqlLYL13_0 & sqlLYN13_0 & sqlTYMSY & sqlLYMSY_0 _
            & " from RSTRETL1 where OPS_YYYYWW Between '" & TYW1 & "' and '" & RYW & "'" & vbCrLf _
            & " and CUST_CODE in (Select CUST_CODE from " & RSTBOWS1_CUST_CODE & ")" & vbCrLf _
            & " and ITEM_CODE in (Select ITEM_CODE from " & RSTBOWS1_ITEM_CODE & ")" & vbCrLf _
            & " and CUST_STORE_NO in (Select CUST_STORE_NO from " & RSTBOWS1_CUST_STORE_NO & ")" & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", 0 QTY_EOW, 0 ORDR_QTY_OPEN, NULL INV_DATE" & vbCrLf _
            & ", 0 SHIPW0, 0 SHIPW1, 0 SHIPW2, 0 SHIPW3, 0 SHIPW4, 0 SHIPW5" & vbCrLf _
            & sqlTYL13_0 & sqlLYL13_0 & sqlLYN13_0 & sqlTYMSY_0 & sqlLYMSY _
            & " from RSTRETL1 where OPS_YYYYWW Between '" & RLYW_01 & "' and '" & LYW & "'" & vbCrLf _
            & " and CUST_CODE in (Select CUST_CODE from " & RSTBOWS1_CUST_CODE & ")" & vbCrLf _
            & " and ITEM_CODE in (Select ITEM_CODE from " & RSTBOWS1_ITEM_CODE & ")" & vbCrLf _
            & " and CUST_STORE_NO in (Select CUST_STORE_NO from " & RSTBOWS1_CUST_STORE_NO & ")" & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", 0 QTY_EOW, 0 ORDR_QTY_OPEN, NULL INV_DATE" & vbCrLf _
            & ", 0 SHIPW0, 0 SHIPW1, 0 SHIPW2, 0 SHIPW3, 0 SHIPW4, 0 SHIPW5" & vbCrLf _
            & sqlTYL13_0 & sqlLYL13 & sqlLYN13 & sqlTYMSY_0 & sqlLYMSY_0 _
            & " from RSTRETL1 where OPS_YYYYWW Between '" & ASCMAIN1.Week_Calc(RYW, -52 * LYX - 12) & "' and '" & ASCMAIN1.Week_Calc(RYW, -52 * LYX + 13) & "'" & vbCrLf _
            & " and CUST_CODE in (Select CUST_CODE from " & RSTBOWS1_CUST_CODE & ")" & vbCrLf _
            & " and ITEM_CODE in (Select ITEM_CODE from " & RSTBOWS1_ITEM_CODE & ")" & vbCrLf _
            & " and CUST_STORE_NO in (Select CUST_STORE_NO from " & RSTBOWS1_CUST_STORE_NO & ")" & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, NVL(ICTITEM1.STYLE_CODE_NC,SOTINVH2.ITEM_CODE) ITEM_CODE" & vbCrLf _
            & ", 0 QTY_EOW, 0 ORDR_QTY_OPEN, MAX(SOTINVH1.INV_DATE) INV_DATE" & vbCrLf _
            & ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & ASCMAIN1.Week_Calc(RYW, -0) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SHIPW0" & vbCrLf _
            & ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & ASCMAIN1.Week_Calc(RYW, -1) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SHIPW1" & vbCrLf _
            & ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & ASCMAIN1.Week_Calc(RYW, -2) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SHIPW2" & vbCrLf _
            & ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & ASCMAIN1.Week_Calc(RYW, -3) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SHIPW3" & vbCrLf _
            & ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & ASCMAIN1.Week_Calc(RYW, -4) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SHIPW4" & vbCrLf _
            & ", Sum (Decode(SOTINVH2.OPS_YYYYWW,'" & ASCMAIN1.Week_Calc(RYW, -5) & "',SOTINVH2.ORDR_QTY_SHIP,0)) SHIPW5" & vbCrLf _
            & sqlTYL13_0 & sqlLYL13_0 & sqlLYN13_0 & sqlTYMSY_0 & sqlLYMSY_0 _
            & " from SOTINVH2,SOTINVH1,ICTITEM1 where SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(RYP, -1) & "' and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'" & vbCrLf _
            & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & " and SOTINVH2.INV_TYPE= 'I'" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & " and SOTINVH2.CUST_CODE in (Select CUST_CODE from " & RSTBOWS1_CUST_CODE & ")" & vbCrLf _
            & " and SOTINVH2.CUST_STORE_NO in (Select CUST_STORE_NO from " & RSTBOWS1_CUST_STORE_NO & ")" & vbCrLf _
            & " and NVL(ICTITEM1.STYLE_CODE_NC,SOTINVH2.ITEM_CODE) in (Select ITEM_CODE from " & RSTBOWS1_ITEM_CODE & ")" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, NVL(ICTITEM1.STYLE_CODE_NC,SOTINVH2.ITEM_CODE)" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, NVL(ICTITEM1.STYLE_CODE_NC,SOTORDR2.ITEM_CODE) ITEM_CODE" & vbCrLf _
            & ", 0 QTY_EOW, Sum (NVL(ORDR_QTY_OPEN,0) + NVL(ORDR_QTY_PICK,0)) ORDR_QTY_OPEN, NULL INV_DATE" & vbCrLf _
            & ", 0 SHIPW0, 0 SHIPW1, 0 SHIPW2, 0 SHIPW3, 0 SHIPW4, 0 SHIPW5" & vbCrLf _
            & sqlTYL13_0 & sqlLYL13_0 & sqlLYN13_0 & sqlTYMSY_0 & sqlLYMSY_0 _
            & " from SOTORDR2,ICTITEM1 where ORDR_STATUS >= 'O' and ORDR_STATUS <= 'P'" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
            & " and SOTORDR2.CUST_CODE in (Select CUST_CODE from " & RSTBOWS1_CUST_CODE & ")" & vbCrLf _
            & " and SOTORDR2.CUST_STORE_NO in (Select CUST_STORE_NO from " & RSTBOWS1_CUST_STORE_NO & ")" & vbCrLf _
            & " and NVL(ICTITEM1.STYLE_CODE_NC,SOTORDR2.ITEM_CODE) in (Select ITEM_CODE from " & RSTBOWS1_ITEM_CODE & ")" & vbCrLf _
            & " group by SOTORDR2.CUST_CODE, SOTORDR2.CUST_STORE_NO, NVL(ICTITEM1.STYLE_CODE_NC,SOTORDR2.ITEM_CODE)" & vbCrLf _
            & sql_sole_AE_SP _
            & ") X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and SATAUTH1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SATAUTH1.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & "   and SATAUTH1.HC_CODE = ICTCOLL1.HC_CODE" & vbCrLf _
            & "   and SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL" & vbCrLf _
            & "   and (SATAUTH1.OPS_YYYYPP_CLOSED IS NULL OR SATAUTH1.OPS_YYYYPP_CLOSED > '" & ASCMAIN1.CYP & "')" & vbCrLf _
            & "   and ARTCUST2.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = X.CUST_STORE_NO" & vbCrLf _
            & IIf(sole_AE <> "" Or sole_REGION <> "", sql_sole_AE & sql_sole_AE_items, "") _
            & "   and NVL(ARTCUST2.CUST_STORE_STATUS,'?') = 'A'" & vbCrLf _
            & "   group by X.CUST_CODE, X.CUST_STORE_NO, X.ITEM_CODE"

        ' LM email 07/20 asking why a future closed store did not appear on the BOW
        '& "   and SATAUTH1.OPS_YYYYPP_CLOSED IS NULL" & vbCrLf _
        '& "   and (SATAUTH1.OPS_YYYYPP_CLOSED IS NULL OR SATAUTH1.OPS_YYYYPP_CLOSED > '" & ASCMAIN1.CYP & "')" & vbCrLf _

        RSTBOWS1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTBOWS1 & " Add Primary Key (CUST_CODE,CUST_STORE_NO,ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Delete from " & RSTBOWS1 & " where ITEM_CODE in (Select ITEM_CODE from ICTITEM1 where ITEM_CODE in (Select Distinct ITEM_CODE from " & RSTBOWS1 & ") and NVL(ITEM_SNU_CODE,'?') <> 'S')")

        If sole_AE <> "" Or sole_REGION <> "" Then
            ASCMAIN1.sql = "" _
                & "DELETE FROM " & RSTBOWS1 & vbCrLf _
                & "WHERE (CUST_CODE, CUST_STORE_NO) IN (" & vbCrLf _
                & "SELECT CUST_CODE, CUST_STORE_NO FROM " & RSTBOWS1 & vbCrLf _
                & "MINUS" & vbCrLf _
                & "SELECT CUST_CODE, CUST_STORE_NO FROM " & RSTBOWS1_CUST_STORE_NO & ")"
            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = "" _
                & "DELETE FROM " & RSTBOWS1 & vbCrLf _
                & "WHERE (CUST_CODE, CUST_STORE_NO) IN (" & vbCrLf _
                & "SELECT CUST_CODE, CUST_STORE_NO FROM " & RSTBOWS1 & vbCrLf _
                & "MINUS" & vbCrLf _
                & "SELECT CUST_CODE, CUST_STORE_NO FROM ARTCUST2)"
            ASCDATA1.ExecuteSQL()
        End If

    End Sub

    Private Sub grdARTCUST1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdARTCUST1.AfterRowActivate
        Setup_grdRSTBOWS1()
    End Sub

    Sub Setup_grdRSTBOWS1()
        If grdARTCUST1.ActiveRow Is Nothing Then
            grdRSTBOWS1.Visible = False
        Else
            grdRSTBOWS1.Visible = True
            Dim CUST_CODE As String = grdARTCUST1.ActiveRow.Cells("CUST_CODE").Value
            Dim dvw As DataView = DirectCast(grdRSTBOWS1.DataSource, DataTable).DefaultView
            dvw.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"
            grdRSTBOWS1.Text = "Boost Order Worksheet Data for " & CUST_CODE
            Sort_grdColumns(grdRSTBOWS1, "ITEM_CODE,CUST_STORE_NO")
        End If
    End Sub

    Sub Initialize_grd(grd As UltraWinGrid.UltraGrid)
        grd.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.SortMulti
        ASCMAIN1.grdInitializeLayout(grd)
        grd.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grd, True)
    End Sub

    Function Make_Formula(COLUMN_NAME As String)
        Dim FX = dst.Tables("RSTBOWS1").Columns(COLUMN_NAME).Expression
        FX = Replace(FX, "IIF", "IF")
        For Each rowRSTBOWSC As DataRow In dst.Tables("RSTBOWSC").Select("", "COLUMN_NAME_LENGTH DESC")
            Dim XC As String = rowRSTBOWSC.Item("COLUMN_NAME")
            Dim COLUMN_INDEX As Integer = Val(rowRSTBOWSC.Item("COLUMN_INDEX") & "")
            Dim XCZ As String = "ISNULL(" & XC & ",0)"
            If InStr(FX, XCZ) <> 0 Then
                FX = Replace(FX, XCZ, Excel_Cell(2, COLUMN_INDEX))
            End If
            XCZ = XC
            If InStr(FX, XCZ) <> 0 Then
                FX = Replace(FX, XCZ, Excel_Cell(2, COLUMN_INDEX))
            End If
        Next
        Return "=" & FX
    End Function
End Class