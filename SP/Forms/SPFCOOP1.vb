Imports Infragistics.Win.UltraWinGrid
Imports System.IO
Public Class SPFCOOP1
    ' SHOW ACCRUAL HiSTORY
    Dim rowSPTCOOP1 As DataRow
    Dim AUTH_NO As String
    Dim AUTH_NO_new As String
    Dim STATUS_CODE As String

    Dim SPTCODE1 As String = ""

    ' DRAG DROP NEEDS TO UPDATE DATATABLE AND DATABAES

    Dim SCHED_events As New Dictionary(Of String, UltraWinSchedule.Appointment)
    Dim dvwSPTSCHD1 As DataView
    Dim apptEdit As Infragistics.Win.UltraWinSchedule.Appointment = Nothing
    Dim SALES_DIVISION_CODE As String

    Dim AUTH_APPR_NOTES As String

    Dim rowARTCUST1 As DataRow
    Dim rowSOTSDIV1 As DataRow

    Dim EXPENSE_TYPE_CODEs_I_may_approve As New List(Of String)
    Dim EXPENSE_TYPE_CODEs_I_may_verify As New List(Of String)


    Dim sqlSPTCOOPX As String
    Dim APPR_STATUS_CODE_BackColors As Dictionary(Of String, System.Drawing.Color)
    Dim APPR_STATUS_CODE_ForeColors As Dictionary(Of String, System.Drawing.Color)
    Dim update_with_approval As Boolean = False
    Dim appRedForeColor As New Infragistics.Win.Appearance
    Dim isSettled As Boolean = False

    Dim VERIFIED_AS_OPEN_COMMENTS_last As String
    Dim VERIFIED_AS_OPEN_COMMENTS_last_do_not_save As Boolean = False
    Dim isInitialScreenLoaded As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SPFCOOPI" Then
            InquiryMode = True
        End If

        appRedForeColor.ForeColor = Drawing.Color.Red

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
            .Items("Clone").Visible = Not InquiryMode And False ' sp email 02/14/19
            .Items("Create Template").Visible = Not InquiryMode
            .Items("Upload Template").Visible = Not InquiryMode
        End With

        Get_PARM("ICTPARM1")

        With dst
            sqlSPTCOOPX = "Select SPTCOOP1.*, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, SPTCOOP3.AUTH_LNO" & vbCrLf _
                & ", SPTCOOP3.FEATURE_DESC, SPTCOOP3.ITEM_CODE, SPTTYPE1.SECURITY_CODE" & vbCrLf _
                & ", ICTCOLL1.COLLECTION_NAME, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME" & vbCrLf _
                & " from SPTCOOP1,SPTCOOP3,ICTCOLL1,ICTBRAN1,SPTTYPE1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE (+) = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
                & "   and SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SPTTYPE1.EXPENSE_TYPE_CODE (+) = SPTCOOP1.EXPENSE_TYPE_CODE" & vbCrLf _
                & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO"
            ASCMAIN1.sql = sqlSPTCOOPX & "  and SPTCOOP1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOPX", "**", 0, False, "V")
            With .Tables("SPTCOOPX").Columns
                .Add("TOTAL_AMT", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
                .Add("DIST_PCT", GetType(System.Decimal), "IIF(ISNULL(TOTAL_AMT,0)=0,0,100*ISNULL(DIST_AMT,0)/ISNULL(TOTAL_AMT,0))")
                .Add("DIST_OPEN", GetType(System.Decimal), "ISNULL(OPEN_AMT,0)*DIST_PCT/100")
                .Add("DIST_PAID", GetType(System.Decimal), "ISNULL(PAID_AMT,0)*DIST_PCT/100")
                .Add("DIST_OPEN_AND_PAID", GetType(System.Decimal), "ISNULL(DIST_OPEN,0)+ISNULL(DIST_PAID,0)")
                .Add("DIST_REAL_EXPENSE", GetType(System.Decimal), "IIF(OPS_YYYYPP>='" & Mid(ASCMAIN1.CYP, 1, 4) & "01" & "',ISNULL(DIST_OPEN,0)+ISNULL(DIST_PAID,0),IIF(ISNULL(DIST_PAID,0)=0,ISNULL(DIST_OPEN,0),ISNULL(DIST_PAID,0)))")
            End With

            ASCMAIN1.sql = sqlSPTCOOPX
            Create_TDA(.Tables.Add, "SPTCOOPG", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "SPTCOOP1", "*")
            With .Tables("SPTCOOP1").Columns
                .Add("TOTAL_AMT", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
                .Add("CANCEL_AMT", GetType(System.Decimal), "IIF(ISNULL(TOTAL_AMT,0) - ISNULL(PAID_AMT,0) - ISNULL(OPEN_AMT,0) < 0,0,ISNULL(TOTAL_AMT,0) - ISNULL(PAID_AMT,0) - ISNULL(OPEN_AMT,0))")
            End With

            ASCMAIN1.sql = "Select SPTCOOP2.*, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                & " from SPTCOOP2,ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SPTCOOP2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SPTCOOP2", "**", 1)

            ASCMAIN1.sql = "Select SPTCOOP3.*, ICTCOLL1.COLLECTION_NAME, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME" & vbCrLf _
                & " from SPTCOOP3,ICTCOLL1,ICTBRAN1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE (+) = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE"
            Create_TDA(.Tables.Add, "SPTCOOP3", "**", 1)
            With .Tables("SPTCOOP3").Columns
                .Add("TOTAL_AMT", GetType(System.Decimal)) ', "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
                .Add("DIST_PCT", GetType(System.Decimal), "IIF(ISNULL(TOTAL_AMT,0)=0,0,100*ISNULL(DIST_AMT,0)/ISNULL(TOTAL_AMT,0))")
                .Add("OPEN_AMT", GetType(System.Decimal)) ', "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
                .Add("PAID_AMT", GetType(System.Decimal)) ', "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
                .Add("DIST_OPEN", GetType(System.Decimal), "ISNULL(OPEN_AMT,0)*DIST_PCT/100")
                .Add("DIST_PAID", GetType(System.Decimal), "ISNULL(PAID_AMT,0)*DIST_PCT/100")
                '.Add("DIST_OPEN_AND_PAID", GetType(System.Decimal), "ISNULL(DIST_OPEN,0)+ISNULL(DIST_PAID,0)")
                '.Add("DIST_REAL_EXPENSE", GetType(System.Decimal), "IIF(OPS_YYYYPP>='" & Mid(ASCMAIN1.CYP, 1, 4) & "01" & "',ISNULL(DIST_OPEN,0)+ISNULL(DIST_PAID,0),IIF(ISNULL(DIST_PAID,0)=0,ISNULL(DIST_OPEN,0),ISNULL(DIST_PAID,0)))")
            End With

            ASCMAIN1.sql = "Select SPTCOOP5.* from SPTCOOP5 where SPTCOOP5.AUTH_NO = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOP5", "**", 0, True, "V")
            With .Tables("SPTCOOP5").Columns
                .Add("MODEL_HRS_TOTAL", GetType(System.Decimal), "ISNULL(MODEL_HRS_01,0)+ISNULL(MODEL_HRS_02,0)+ISNULL(MODEL_HRS_03,0)+ISNULL(MODEL_HRS_04,0)+ISNULL(MODEL_HRS_05,0)+ISNULL(MODEL_HRS_06,0)+ISNULL(MODEL_HRS_07,0)+ISNULL(MODEL_HRS_08,0)+ISNULL(MODEL_HRS_09,0)+ISNULL(MODEL_HRS_10,0)")
                .Add("MODEL_AMT_TOTAL", GetType(System.Decimal), "ISNULL(MODEL_RATE,0) * MODEL_HRS_TOTAL")
            End With

            ASCMAIN1.sql = "Select SPTCOOP6.* from SPTCOOP6 where SPTCOOP6.AUTH_NO = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOP6", "**", 0, True, "V")

            ASCMAIN1.sql = "Select SPTCOOP7.* from SPTCOOP7 where SPTCOOP7.AUTH_NO = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOP7", "**", 0, True, "V")

            ASCMAIN1.sql = "Select SPTCOOP8.* from SPTCOOP8 where SPTCOOP8.AUTH_NO = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOP8", "**", 0, True, "V")

            ASCMAIN1.sql = "Select SPTCOOP9.*, ARTCUST2.CUST_STORE_LOCATION" & vbCrLf _
                & " from SPTCOOP9,ARTCUST2,SPTCOOP1" & vbCrLf _
                & " where SPTCOOP9.AUTH_NO = :PARM1" & vbCrLf _
                & "   and SPTCOOP1.AUTH_NO = SPTCOOP9.AUTH_NO" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SPTCOOP1.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SPTCOOP9.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SPTCOOP9", "**", 0, True, "V")
            For P As Integer = -6 To 3
                Dim C As String = ""
                If P <= 0 Then
                    C = "P" & Format(-1 * P, "00")
                Else
                    C = "N" & Format(P, "00")
                End If
                .Tables("SPTCOOP9").Columns.Add("RTL_LY_" & C, GetType(System.Decimal))
                .Tables("SPTCOOP9").Columns.Add("RTL_TY_" & C, GetType(System.Decimal))
            Next

            ASCMAIN1.sql = "Select SPTCOOPB.*" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_LOCATION, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & ", ARTCUST2.SELL_CODE, SOTSELL1.REGION_CODE" & vbCrLf _
                & " from SPTCOOPB,ARTCUST2,SPTCOOP1,SOTSELL1" & vbCrLf _
                & " where SPTCOOPB.AUTH_NO = :PARM1" & vbCrLf _
                & "   and SPTCOOP1.AUTH_NO = SPTCOOPB.AUTH_NO" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SPTCOOP1.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SPTCOOPB.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SPTCOOPB", "**", 0, True, "V", 2)
            With .Tables("SPTCOOPB")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "Select SPTSCHD1.*, TATWHOP1.WH_OPER_NAME, TATWHOP1.WH_OPER_GRP, 0 TOTAL_DAYS"
            ASCMAIN1.sql &= " from SPTSCHD1, TATWHOP1 "
            ASCMAIN1.sql &= " where TATWHOP1.WH_OPER_ID = SPTSCHD1.WH_OPER_ID"
            Create_TDA(.Tables.Add, "SPTSCHD1", "**", 0, True, String.Empty)
            .Tables("SPTSCHD1").Columns("WH_OPER_GRP").MaxLength = 10

            Create_TDA(.Tables.Add, "SPTSCHDL", "*")
            Fill_Records("SPTSCHDL", String.Empty, True, "SELECT * FROM SPTSCHDL")

            ASCMAIN1.sql = "Select CUST_NAME CODE_TYPE, CUST_NAME CODE_VALUE, CUST_NAME DESC_VALUE, '1' SEL from ARTCUST1 where ROWNUM < 1"
            SPTCODE1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SPTCODE1 & " Modify DESC_VALUE VARCHAR2(100)")
            ASCDATA1.ExecuteSQL("Alter Table " & SPTCODE1 & " Add Primary Key (CODE_TYPE, CODE_VALUE)")
            ASCDATA1.ExecuteSQL("Insert into " & SPTCODE1 & " Select 'CUST_CODE' CODE_TYPE, CUST_CODE, CUST_NAME, '1' SEL from ARTCUST1")
            ASCDATA1.ExecuteSQL("Insert into " & SPTCODE1 & " Select 'SREP_CODE' CODE_TYPE, SREP_CODE, SREP_NAME, '1' SEL from SOTSREP1 where SREP_TYPE = 'I'")
            ASCDATA1.ExecuteSQL("Insert into " & SPTCODE1 & " Select 'COLLECTION_CODE' CODE_TYPE, COLLECTION_CODE, COLLECTION_NAME, '1' SEL from ICTCOLL1")
            ASCDATA1.ExecuteSQL("Insert into " & SPTCODE1 & " Select 'BRAND_CODE' CODE_TYPE, BRAND_CODE, BRAND_NAME, '1' SEL from ICTBRAN1")

            ASCMAIN1.sql = "Select * from " & SPTCODE1
            Create_TDA(.Tables.Add("SPTCODE1"), SPTCODE1, "**", 0)

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'SPTCOOP1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V")
            .Tables("TATEVNT1").Columns.Add("ATTACHMENT_EXT")


            ASCMAIN1.sql = "Select * from ASTATTA2 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'AUTH_NO' and CODE_VALUE = :PARM1 AND NVL(ATTACHMENT_STATUS,'?') <> 'D'"
            Create_TDA(.Tables.Add, "ASTATTA2", "**", 0, True, "V")
            '  .Tables("TATEVNT1").Columns.Add("ATTACHMENT_EXT")

            ASCMAIN1.sql = "Select SPTPYMT2.*" _
                & ", SPTPYMT1.PYMT_TYPE, SPTPYMT1.PYMT_REF_NO, SPTPYMT1.PYMT_REF_DATE" _
                & ", SPTPYMT1.PYMT_REF_AMT PYMT_REF_AMT_TOTAL, SPTPYMT1.OPS_YYYYPP" _
                & ", SPTPYMT1.CUST_CODE, SPTPYMT1.VEND_CODE, SPTPYMT1.PYMT_CTL_NO" _
                & ", SPTPYMT1.REVERSED_BY_PYMT_NO, SPTPYMT1.REVERSED_PYMT_NO" _
                & " from SPTPYMT2, SPTPYMT1 " _
                & " where SPTPYMT1.PYMT_NO = SPTPYMT2.PYMT_NO and SPTPYMT2.AUTH_NO = :PARM1"
            Create_TDA(.Tables.Add, "SPTPYMTX", "**", 0, False, "V")
        End With

        Bind_Controls(grpTotals, "SPTCOOP1")

        Set_Read_Only(grpTotals, True)

        Fill_Records("SPTCODE1")

        grdSPTCOOP2.DataSource = dst.Tables("SPTCOOP2")
        grdSPTCOOP3.DataSource = dst.Tables("SPTCOOP3")
        grdSPTCOOP5.DataSource = dst.Tables("SPTCOOP5")
        grdSPTCOOP6.DataSource = dst.Tables("SPTCOOP6")
        grdSPTCOOP8.DataSource = dst.Tables("SPTCOOP8")
        grdSPTCOOP9.DataSource = dst.Tables("SPTCOOP9")
        grdSPTCOOPX.DataSource = dst.Tables("SPTCOOPX")
        grdSPTCOOPG.DataSource = dst.Tables("SPTCOOPG")
        grdSPTCODE1.DataSource = dst.Tables("SPTCODE1")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdASTAUDT1.DataSource = dst.Tables("ASTAUDT1")
        grdASTATTA2.DataSource = dst.Tables("ASTATTA2")
        grdSPTPYMTX.DataSource = dst.Tables("SPTPYMTX")
        grdSPTCOOPB.DataSource = dst.Tables("SPTCOOPB")

        If ASCMAIN1.CLIENT = "INT" Then
            tabPromo.Tabs("Featuring").Visible = False
        End If

        Create_Summary(grdSPTCOOPX, "AUTH_NO", "Count")
        Create_Summary(grdSPTCOOPX, New String() {"DIST_AMT", "OPEN_AMT", "PAID_AMT", "DIST_OPEN", "DIST_PAID", "DIST_OPEN_AND_PAID", "DIST_REAL_EXPENSE"})

        Create_Summary(grdSPTCOOPG, "AUTH_NO", "Count")
        Create_Summary(grdSPTCOOPG, New String() {"DIST_AMT", "OPEN_AMT", "PAID_AMT"})

        Create_Summary(grdSPTCOOP6, "AUTH_TNO", "Count")

        Create_Summary(grdSPTCOOP9, "CUST_STORE_NO", "Count")
        Create_Summary(grdSPTCOOP9, New String() {"IN_STORE"})

        Create_Summary(grdSPTCOOP5, "AUTH_SNO", "Count")
        Create_Summary(grdSPTCOOP5, New String() {"MODEL_HRS_TOTAL", "MODEL_AMT_TOTAL"})

        Create_Summary(grdSPTCOOP8, "AUTH_ENO", "Count")
        Create_Summary(grdSPTCOOP8, New String() {"AUTH_EXP_QTY", "AUTH_EXP_AMT"})

        Create_Summary(grdSPTCOOPB, "CUST_STORE_NO", "Count")
        Create_Summary(grdSPTCOOPB, New String() {"SEL"})

        Create_Summary(grdSPTCOOP3, "AUTH_LNO", "Count")
        Create_Summary(grdSPTCOOP3, New String() {"DIST_AMT", "DIST_OPEN", "DIST_PAID"})
        grdSPTCOOP3.DisplayLayout.Bands(0).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False

        Create_Summary(grdSPTCOOP2, "AUTH_LNO", "Count")
        grdSPTCOOP2.DisplayLayout.Bands(0).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False

        Create_Summary(grdSPTPYMTX, "PYMT_NO", "Count")
        Create_Summary(grdSPTPYMTX, "PYMT_REF_AMT")
        grdSPTPYMTX.DisplayLayout.Bands(0).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False

        With grdSPTCOOPG.DisplayLayout.Bands("SPTCOOPG")
            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTCOOP2.DisplayLayout.Bands("SPTCOOP2")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTCOOP3.DisplayLayout.Bands("SPTCOOP3")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"TOTAL_AMT", "DIST_AMT", "DIST_PCT", "DIST_OPEN", "DIST_PAID", "DIST_OPEN_AND_PAID"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTCOOP5.DisplayLayout.Bands("SPTCOOP5")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"AUTH_SNO", "MODEL_CODE", "MODEL_NAME", "MODEL_NOTE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                ElseIf New String() {"MODEL_RATE", "MODEL_AMT_TOTAL"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet
                End If
            Next

            For h As Integer = 1 To 10
                With .Columns("MODEL_HRS_" & Format(h, "00"))
                    .Width = 40
                    .Header.Caption = Format(h, "00")
                End With
            Next

            For Each C As String In New String() {"MODEL_HRS_TOTAL", "MODEL_AMT_TOTAL"}
                With .Columns(C)
                    .Width = 60
                    ' .Header.Caption = "Total " & Mid(C, 13, 3)
                End With
            Next

            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdASTAUDT1.DisplayLayout.Bands("ASTAUDT1")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Silver
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Silver
                End If
            Next
            '.Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdASTATTA2.DisplayLayout.Bands("ASTATTA2")
            .Override.AllowDelete = DefaultableBoolean.False
            ' IF YOU WANT TO ENABLE DELETE HERE,
            ' 1 YOU MUST OBSERVE WHAT HAPPENS WHEN YOU DELETE IN THE POPUP, AND
            ' 2 YOU MUST DECIDE IF IT IS RIGHT TO BE ABLE TO OFFER DELETE IN INQUIRY MODE

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Silver
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Silver
                End If
            Next
            '.Columns("AUTH_NO").Header.Fixed = True
        End With


        With grdTATEVNT1.DisplayLayout.Bands("TATEVNT1")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Silver
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Silver
                End If
            Next
            ' .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTCOOPX.DisplayLayout.Bands("SPTCOOPX")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"NOTES", "BOOKING_NAME", "VERIFIED_AS_OPEN_COMMENTS"}.Contains(GCOL.Key) Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

                If New String() {"OPEN_AMT", "PAID_AMT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"COLLECTION_CODE", "DIST_AMT", "DIST_PCT", "DIST_OPEN", "DIST_PAID", "DIST_OPEN_AND_PAID"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                ElseIf New String() {"AUTH_APPR_DATE", "AUTH_APPR_BY", "AUTH_APPR_AMT", "AUTH_APPR_NOTES"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightPink
                ElseIf New String() {"VERIFIED_AS_OPEN", "VERIFIED_AS_OPEN_NOTES", "VERIFIED_AS_OPEN_AMT", "VERIFIED_AS_OPEN_BY", "VERIFIED_AS_OPEN_DATE", "VERIFIED_AS_OPEN_COMMENTS"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Violet
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTCOOP6.DisplayLayout.Bands("SPTCOOP6")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"TASK_TYPE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTCOOP8.DisplayLayout.Bands("SPTCOOP8")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"AUTH_EXP_QTY", "AUTH_EXP_PRICE", "AUTH_EXP_AMT"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With

        With grdSPTCOOP9.DisplayLayout.Bands("SPTCOOP9")
            .LevelCount = 2
            .ColHeadersVisible = False

            Dim G As UltraWinGrid.UltraGridGroup

            G = .Groups.Add
            G.Header.Caption = "Store"
            G.Key = "STORE"
            G.Header.Appearance.BackColor = System.Drawing.Color.White
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            G.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
            For Each COLUMN_NAME As String In New String() {"CUST_STORE_NO", "CUST_STORE_LOCATION", "IN_STORE"}
                .Columns(COLUMN_NAME).Group = G
            Next

            G = .Groups.Add
            G.Header.Caption = "Yr"
            G.Key = "YEAR"
            G.Header.Appearance.BackColor = System.Drawing.Color.White
            G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            G.Header.Appearance.BackColor2 = System.Drawing.Color.Silver
            For Each COLUMN_NAME As String In New String() {"LY", "TY"}
                .Columns(COLUMN_NAME).Group = G
                .Columns(COLUMN_NAME).CellAppearance.TextHAlign = HAlign.Center
            Next
            .Columns("TY").Level = 1
            G.Width = 40
            G.Header.Appearance.TextHAlign = HAlign.Center

            Dim C As String = ""
            For P As Integer = -6 To 3
                If P = 0 Then
                    G = .Groups.Add
                    G.Header.Caption = "Goal"
                    G.Key = "GOAL"
                    G.Header.Appearance.BackColor = System.Drawing.Color.White
                    G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    G.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                    G.Header.Appearance.TextHAlign = HAlign.Right
                    .Columns("RETAIL_GOAL").Group = G
                    .Columns("RETAIL_GOAL").Format = "###,##0"
                    .Columns("RETAIL_GOAL").Level = 1
                    Create_Summary(grdSPTCOOP9, "RETAIL_GOAL")
                    G.Width = 60
                End If

                If P <= 0 Then
                    C = "P" & Format(-1 * P, "00")
                Else
                    C = "N" & Format(P, "00")
                End If
                G = .Groups.Add
                G.Header.Caption = C
                G.Key = C
                G.Header.Appearance.BackColor = System.Drawing.Color.White
                G.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If P >= 0 Then
                    G.Header.Appearance.BackColor2 = System.Drawing.Color.Pink
                Else
                    G.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                End If

                G.Header.Appearance.TextHAlign = HAlign.Right

                .Columns("RTL_LY_" & C).Group = G
                .Columns("RTL_LY_" & C).Format = "###,##0"
                Create_Summary(grdSPTCOOP9, "RTL_LY_" & C)

                .Columns("RTL_TY_" & C).Group = G
                .Columns("RTL_TY_" & C).Format = "###,##0"
                .Columns("RTL_TY_" & C).Level = 1
                Create_Summary(grdSPTCOOP9, "RTL_TY_" & C)

                G.Width = 60
            Next

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "IN_STORE" Or GCOL.Key = "RETAIL_GOAL" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With


        With grdSPTCOOPB.DisplayLayout.Bands("SPTCOOPB")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SEL" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            .Columns("AUTH_NO").Header.Fixed = True
        End With


        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOPX, "BOOKED_BY", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'BOOKED_BY'")

        ASCMAIN1.Add_Value_List(grdSPTCOOPG, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTCOOPG, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")

        ASCMAIN1.Add_Value_List(grdSPTPYMTX, "PYMT_TYPE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTPYMT1' and COLUMN_NAME = 'PYMT_TYPE'")
        'ASCMAIN1.Add_Value_List(Absx1.cbeFor("BOOKED_BY"), "BOOKED_BY", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'BOOKED_BY'")

        Absx1.cbeFor("BOOKED_BY").DataSource = ASCDATA1.GetDataTable("Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'BOOKED_BY'")
        'Absx1.cbeFor("BOOKED_BY").SelectedItem = cbeYP.Items(0)

        Set_Read_Only(grpTotals, True)
        grpHeader.Visible = False

        APPR_STATUS_CODE_BackColors = New Dictionary(Of String, Drawing.Color)

        APPR_STATUS_CODE_BackColors.Add("A", System.Drawing.Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("P", System.Drawing.Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("G", System.Drawing.Color.Empty)
        APPR_STATUS_CODE_BackColors.Add("X", System.Drawing.Color.Empty)

        APPR_STATUS_CODE_ForeColors = New Dictionary(Of String, Drawing.Color)

        APPR_STATUS_CODE_ForeColors.Add("A", System.Drawing.Color.Green)
        APPR_STATUS_CODE_ForeColors.Add("P", System.Drawing.Color.Purple)
        APPR_STATUS_CODE_ForeColors.Add("G", System.Drawing.Color.Blue)
        APPR_STATUS_CODE_ForeColors.Add("X", System.Drawing.Color.Red)

        For Each VLI As ValueListItem In Absx1.optFor("APPR_STATUS_CODE").ValueList.ValueListItems
            VLI.Appearance.ForeColor = APPR_STATUS_CODE_ForeColors(VLI.DataValue)
        Next

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 24) & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(24)

        grdSPTCOOPS.DataSource = dst.Tables("SPTSCHDL")

        eventMonthView.VisibleWeeks = 4

        Dim calLook As Infragistics.Win.UltraWinSchedule.UltraCalendarLook = New Infragistics.Win.UltraWinSchedule.UltraCalendarLook
        calLook.AlternateMonthAppearance.BackColor = System.Drawing.Color.LightYellow
        calLook.SelectedDayAppearance.BackColor = System.Drawing.Color.DarkSlateBlue
        calLook.SelectedDayAppearance.ForeColor = System.Drawing.Color.YellowGreen

        Me.eventMonthView.CalendarLook = calLook

        Show_Filter(grdSPTCOOPX, True)

        For Each gcol As UltraWinGrid.UltraGridColumn In grdSPTCOOPX.DisplayLayout.Bands(0).Columns
            If dst.Tables("SPTCOOPX").Columns(gcol.Key).DataType.ToString = "System.String" Then
                gcol.FilterOperatorDefaultValue = UltraWinGrid.FilterOperatorDefaultValue.Contains
            End If

            'If gcol.FilterOperatorDefaultValue = UltraWinGrid.FilterOperatorDefaultValue.StartsWith Then
            '    gcol.FilterOperatorDefaultValue = UltraWinGrid.FilterOperatorDefaultValue.Contains
            'End If
        Next

        dteStartDate.DateTime = DateAdd(DateInterval.Day, -30, DateTime.Now)
        dteEndDate.DateTime = DateAdd(DateInterval.Day, 30, DateTime.Now)

        ASCMAIN1.sql = "Select T_DESC AD_SIZE from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'AD_SIZE'"
        cmbAD_SIZE.DataSource = ASCDATA1.GetDataTable
        ' cbeAD_SIZE.ValueList = ""
        'ASCMAIN1.Add_Value_List(cbeAD_SIZE, "AD_SIZE")

        For Each tabC As String In New String() {"Tasks", "Store Goals", "Scheduling", "Expenses"}
            tabPromo.Tabs(tabC).Visible = False
        Next

        ASCMAIN1.sql = "Select * from SPTTYPE1"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim SECURITY_CODE As String = row.Item("SECURITY_CODE") & ""
            Dim EXPENSE_TYPE_CODE As String = row.Item("EXPENSE_TYPE_CODE") & ""

            If SECURITY_CODE <> "" And ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
                EXPENSE_TYPE_CODEs_I_may_verify.Add(EXPENSE_TYPE_CODE)
            End If

            'If ASCMAIN1.CLIENT = "INT" Then
            If SECURITY_CODE <> "" And ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
                EXPENSE_TYPE_CODEs_I_may_approve.Add(EXPENSE_TYPE_CODE)
            End If
            'Else
            '    If SECURITY_CODE = "" Then
            '        EXPENSE_TYPE_CODEs_I_may_approve.Add(EXPENSE_TYPE_CODE)
            '    Else
            '        If SECURITY_CODE <> "" And ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
            '            EXPENSE_TYPE_CODEs_I_may_approve.Add(EXPENSE_TYPE_CODE)
            '        End If
            '    End If
            'End If

        Next

        If EXPENSE_TYPE_CODEs_I_may_verify.Count = 0 Then
            chkVerify.Visible = False
        End If

        If ASCMAIN1.CLIENT = "INT" Then
            lblDATE_ACCRUE.Visible = False
            dteDATE_ACCRUE.Visible = False
            With grdSPTCOOPX.DisplayLayout.Bands(0)
                .Columns("DATE_ACCRUE").Hidden = True
                .Columns("OPS_YYYYPP_ACCRUE").Hidden = True
            End With
        End If

        If InquiryMode Then
            chkEditNotes.Visible = False
            chkEditNotes.Enabled = False

            chkEditVerComment.Visible = False
            chkEditVerComment.Enabled = False

            chkVerify.Visible = False
            chkVerify.Enabled = False
        End If

        tab0.Tabs("Tasks").Visible = False
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("CUST_CODE")

                If Absx1.dteFor("AUTH_DATE").Value & "" = "" Then
                    Absx1.dteFor("AUTH_DATE").Value = Now.Date
                End If

                Dim DT As Date = Absx1.dteFor("AUTH_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                End If

                If Absx1.txtFor("CUST_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If IsNothing(rowARTCUST1) Then
                        EMsg &= vbCr & "Customer Entered Is Not Valid"
                    Else
                        If rowARTCUST1.Item("CUST_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Customer Entered Is Not Active"

                        End If
                    End If
                End If

                'If Absx1.txtFor("SALES_DIVISION_CODE").Text.Length = 0 Then
                '    EMsg &= vbCr & "You must supply a Valid Division"
                'Else
                '    rowSOTSDIV1 = LookUp("SOTSDIV1", Absx1.txtFor("SALES_DIVISION_CODE").Text)
                '    If IsNothing(rowSOTSDIV1) Then
                '        EMsg &= vbCr & "Division Entered Is Not Valid"
                '        'Else
                '        '    If rowARTCUST1.Item("CUST_STATUS").ToString <> "A" Then
                '        '        EMsg &= vbCr & "Customer Entered Is Not Active"
                '        '    End If
                '    End If
                'End If

                If Absx1.txtFor("EXPENSE_TYPE_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Expense Type"
                Else
                    Validate_Code("EXPENSE_TYPE_CODE")
                End If

                'If EMsg = "" Then
                '    AUTH_NO = Absx1.txtFor("AUTH_NO").Text
                '    If Not ASCMAIN1.Logical_Lock("SPTCOOP1", AUTH_NO) Then
                '        Exit Sub
                '    End If
                'End If

            Case "View", "Edit"
                AUTH_NO = Absx1.txtFor("AUTH_NO").Text
                If AUTH_NO = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowSPTCOOP1 = LookUp("SPTCOOP1", AUTH_NO)
                    If rowSPTCOOP1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & AUTH_NO & " on File"
                    Else
                        If eItemKey = "Edit" Then
                            If rowSPTCOOP1.Item("APPR_STATUS_CODE") & "" = "A" Then
                                Dim EXPENSE_TYPE_CODE As String = rowSPTCOOP1.Item("EXPENSE_TYPE_CODE") & ""
                                If Not EXPENSE_TYPE_CODEs_I_may_approve.Contains(EXPENSE_TYPE_CODE) Then
                                    If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("P1") Then
                                        EMsg &= vbCr & "Auth No " & AUTH_NO & " has already been approved." & vbCr _
                                            & "No Changes (except for by Approver)"
                                    End If
                                End If
                            End If

                            ' as per AK - allow any one to edit an old Promo record 02/24/2023
                            'If Format(CDate(rowSPTCOOP1.Item("DATE_START")), "yyyy") < Mid(ASCMAIN1.CYP, 1, 4) Then
                            '    If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("P1") Then
                            '        EMsg &= vbCr & "You may not Edit Promo Events with a Start Date prior to the 1st of the Current Operations Year"
                            '    End If
                            'End If

                            If rowSPTCOOP1.Item("APPR_STATUS_CODE") = "X" Or rowSPTCOOP1.Item("STATUS_CODE") = "C" Then
                                If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("P1") Then
                                    'EMsg &= vbCr & "You may not Edit Promo Events that are Cancelled or Closed"
                                    MsgBox("Promo Event is Cancelled or Closed - Edits are Restricted", MsgBoxStyle.OkOnly, "Verificaiton")
                                End If
                            End If

                            If EMsg = "" Then
                                If Not ASCMAIN1.Logical_Lock("SPTCOOP1", AUTH_NO) Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If

                If optShow.Value = "G" Then
                    If rowSPTCOOP1.Item("APPR_STATUS_CODE") & "" <> "G" Then
                        EMsg &= "Event " & AUTH_NO & " is NOT Pending Approval"
                    End If
                End If

                If eItemKey = "Edit" And EMsg = "" Then
                    If rowSPTCOOP1.Item("EVENT_GROUP_NO") & "" <> "" Then
                        EMsg &= "Event " & AUTH_NO & " may NOT be maintained here - use Store Focus Events screen"
                    End If
                End If


                'If Not chkStartDate.Checked And Not IsDate(dteStartDate.DateTime) Then
                '    EMsg &= vbCr & "You must provide a valid Start Date."
                'ElseIf Not chkEnddate.Checked And Not IsDate(dteEndDate.DateTime) Then
                '    EMsg &= vbCr & " You must provide a valid End Date."
                'ElseIf Not chkStartDate.Checked AndAlso Not chkEnddate.Checked Then
                '    If dteStartDate.DateTime.Date > dteEndDate.DateTime.Date Then
                '        EMsg &= vbCr & "Start date must be less equal End date."
                '    End If
                'End If


            Case "Update"


                If Absx1.txtFor("SREP_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Sell-In Rep"
                Else
                    Dim row As DataRow = LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Sell-In Rep"
                    End If
                End If

                'If Absx1.txtFor("SELL_CODE").Text = "" Then
                '    EMsg &= vbCr & "You Must Specify a Sell-Thru Rep"
                'Else
                '    Dim row As DataRow = LookUp("SOTSELL1", Absx1.txtFor("SELL_CODE").Text)
                '    If row Is Nothing Then
                '        EMsg &= vbCr & "Invalid Value Specified for Sell-Thru Rep"
                '    End If
                'End If

                If Absx1.txtFor("EVENT_TYPE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify an Event Type Code"
                Else
                    Dim row As DataRow = LookUp("SPTEVNT1", Absx1.txtFor("EVENT_TYPE_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Event Type Code"
                    End If
                End If

                If Absx1.txtFor("VEHICLE_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify an Advertising Vehicle Code"
                Else
                    Dim row As DataRow = LookUp("SPTAVEH1", Absx1.txtFor("VEHICLE_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Advertising Vehicle Code"
                    Else
                        If ASCMAIN1.CLIENT = "AHA" Then
                            If Absx1.txtFor("VEHICLE_CODE").Text <> Absx1.txtFor("EXPENSE_TYPE_CODE").Text Then
                                EMsg &= vbCr & "Invalid Value Specified for Advertising Vehicle Code - MUST BE SAME AS EXPENSE TYPE"
                            End If
                        End If
                    End If
                End If



                'If Absx1.txtFor("FEATURE_DESC").Text = "" Then
                '    EMsg &= vbCr & "You Must Specify a Description of an Item, Collection or Brand being Featured"
                'End If

                If Absx1.txtFor("BOOKING_NAME").Text = "" Then
                    EMsg &= vbCr & "You Must Specify the Booking Name of this Sales Promotion Event"
                End If

                If grdSPTCOOP3.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Distribution Details Entered"

                    'Else
                    '    For Each rowSPTCOOP3 As DataRow In dst.Tables("SPTCOOP3").Select("", "", DataViewRowState.CurrentRows)
                    '        If rowSPTCOOP3.Item("COST_CATGY_CODE") & "" = "" Then
                    '            EMsg &= vbCr & "Unable to determine Cost Category for " & rowSPTCOOP3.Item("VEHICLE_CODE") & ""
                    '        End If
                    '        If rowSPTCOOP3.Item("PROD_CODE") & "" = "" Then
                    '            EMsg &= vbCr & "Unable to determine Product Code for " & rowSPTCOOP3.Item("VEHICLE_CODE") & ""
                    '        End If
                    '    Next
                Else

                    For Each row As DataRow In dst.Tables("SPTCOOP3").Select("")
                        Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE") & ""

                        'Dim DIST_AMT As Decimal = Val(row.Item("DIST_AMT") & "")
                        'Dim DIST_PAID As Decimal = Val(row.Item("DIST_PAID") & "")
                        'If DIST_AMT < DIST_PAID Then
                        '    EMsg &= vbCr & "Distribution Amount may not be less than Paid Amount"
                        'End If

                        If COLLECTION_CODE = "" Then
                            If Absx1.txtFor("EXPENSE_TYPE_CODE").Text <> "RTLEVENTS" Then
                                EMsg &= vbCr & "Invalid Expense Type for Blank Collection Distributions (" & Absx1.txtFor("EXPENSE_TYPE_CODE").Text & ")"
                            End If
                            If System.Math.Round(Val(Absx1.numFor("TOTAL_AMT").Value & ""), 2) <> 0 Then
                                EMsg &= vbCr & "Blank Collection Distributions not valid with $$$ Events"
                            End If
                        Else
                            If LookUp("ICTCOLL1", COLLECTION_CODE) Is Nothing Then
                                EMsg &= vbCr & "Invalid Collection Code (" & COLLECTION_CODE & ")"
                            End If
                        End If
                    Next
                End If

                Dim TOTAL_AMT As Decimal = System.Math.Round(Val(Absx1.numFor("TOTAL_AMT").Value & ""), 2)
                If Val(dst.Tables("SPTCOOP3").Compute("SUM(DIST_AMT)", "") & "") <> TOTAL_AMT Then
                    EMsg &= vbCr & "Total of Distribution does not agree with Total for Commitment"
                End If

                If dst.Tables("SPTCOOP3").Select("ISNULL(COLLECTION_CODE,'')='' AND ISNULL(DIST_AMT,0)<>0").Length <> 0 Then
                    EMsg &= vbCr & "You may not distribute an expense amount without specifying a Collection"
                End If

                Dim DATE_START As Date = Absx1.dteFor("DATE_START").Value
                If DATE_START & "" = "12:00:00 AM" Then
                    EMsg &= vbCr & "Start Date is Mandatory"
                Else
                    ' TAC.SOCMAIN1.Validate_Invoice_Date(DATE_START, 2, 1, EMsg)
                End If
                Dim DATE_END As Date = Absx1.dteFor("DATE_END").Value
                If DATE_END & "" = "12:00:00 AM" Then
                    EMsg &= vbCr & "End Date is Mandatory"
                Else
                    ' TAC.SOCMAIN1.Validate_Invoice_Date(DATE_END, 2, 1, EMsg)
                End If

                If DATE_START & "" <> "" And DATE_END & "" <> "" Then
                    If Format(DATE_START, "yyyyMMdd") > Format(DATE_END, "yyyyMMdd") Then
                        EMsg &= vbCr & "Start Date may not be later than End Date"
                    End If

                    If EntryMode = "N" Then
                        If Format(DATE_START, "yyyy") < Mid(ASCMAIN1.CYP, 1, 4) Then
                            EMsg &= vbCr & "Start Date may not be prior to start of Current Ops Year"
                        End If
                        If Format(DATE_START, "yyyy") > Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) + 1, "0000") Then
                            EMsg &= vbCr & "Start Date may not be later than last day of next Year"
                        End If
                    ElseIf EntryMode = "E" Then
                        Dim DATE_START_orig As Date = rowSPTCOOP1.Item("DATE_START", DataRowVersion.Original)
                        If Format(DATE_START, "yyyy") <> Format(DATE_START_orig, "yyyy") Then
                            EMsg &= vbCr & "The Year of the Start Date may not changed to a Different Year"
                        Else

                            ' the code below is not necessary since we do not permit changing the year to a different (ie, prior) year
                            'If Format(DATE_START, "yyyyMMdd") < Format(DATE_START_orig, "yyyyMMdd") Then
                            '    If Format(DATE_START, "yyyyMMdd") < ASCMAIN1.CYP & "01" Then
                            '        EMsg &= vbCr & "The Year of the Start Date may not changed to a Date Prior to the Start of the Year"
                            '    End If
                            'End If
                        End If

                    End If


                End If

                If ASCMAIN1.CLIENT = "INT" Then
                Else
                    Dim DATE_ACCRUE As Date = Absx1.dteFor("DATE_ACCRUE").Value
                    If DATE_ACCRUE & "" = "12:00:00 AM" Then
                        EMsg &= vbCr & "Accrual Date is Mandatory"
                    Else
                        ' TAC.SOCMAIN1.Validate_Invoice_Date(DATE_START, 2, 1, EMsg)
                    End If

                End If

                If Absx1.txtFor("SEASON_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Season"
                Else
                    Dim row As DataRow = LookUp("ICTSEAS1", Absx1.txtFor("SEASON_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Season"
                    Else
                        If EMsg = "" Then
                            Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
                            If Mid(SEASON_CODE, 1, 4) <> Format(DATE_START, "yyyy") Then
                                EMsg &= vbCr & "Season not congruous with Start Date"
                            Else
                                If Mid(SEASON_CODE, 5, 1) = "S" And Format(DATE_START, "MM") >= "07" Then
                                    EMsg &= vbCr & "Season not congruous with Start Date"
                                End If
                                If Mid(SEASON_CODE, 5, 1) = "F" And Format(DATE_START, "MM") < "07" Then
                                    EMsg &= vbCr & "Season not congruous with Start Date"
                                End If
                            End If

                        End If

                    End If
                End If

                Dim TOTAL_AMT_promo As Decimal = Val(dst.Tables("SPTCOOP3").Compute("SUM(DIST_AMT)", "") & "")
                Dim PAID_AMT_promo As Decimal = Val(Absx1.numFor("PAID_AMT").Value & "")
                Dim OPEN_AMT_promo As Decimal = TOTAL_AMT_promo - PAID_AMT_promo
                If OPEN_AMT_promo < 0 Then EMsg &= vbCr & "Cannot lower Total Distributed amount to less than amount Paid to date"

                update_with_approval = False
                Dim APPR_STATUS_CODE As String = Absx1.optFor("APPR_STATUS_CODE").Value
                Dim APPR_STATUS_CODE_OLD As String = ""
                If EntryMode = "N" Then
                Else
                    APPR_STATUS_CODE_OLD = rowSPTCOOP1.Item("APPR_STATUS_CODE", DataRowVersion.Original) & ""


                    If APPR_STATUS_CODE = "X" And APPR_STATUS_CODE_OLD = "A" Then
                        Dim msg = ""
                        If PAID_AMT_promo <> 0 Then
                            msg &= vbCrLf & " - Promo has Payments Recorded"
                        End If
                        If Format(DATE_START, "yyyyMM") < ASCMAIN1.CYP Then
                            msg &= vbCrLf & " - Promo has Already been Accrued (based on Start Date)"
                        End If
                        If msg <> "" Then
                            If MsgBox("This Event has already been recorded with Financial Activity" & msg & vbCrLf & vbCrLf & "Are you Sure that you want to Cancel this Event?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                EMsg &= vbCr & "Update to Cancelled Status has been avoided because" & msg
                            End If
                        End If
                    End If

                    If APPR_STATUS_CODE <> APPR_STATUS_CODE_OLD And APPR_STATUS_CODE_OLD = "A" Then
                        If rowSPTCOOP1.Item("VERIFIED_AS_OPEN_COMMENTS") & "" <> "" Then
                            If MsgBox("This Contract was Previously Approved" & vbCrLf & " and it has already been *Verified*." & vbCrLf & vbCrLf & "Are you sure that you want to reverse the Approved Status?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                EMsg &= vbCr & "Update has been Stopped because of Previous Verification"
                            End If
                        End If
                    End If

                End If
                If APPR_STATUS_CODE = "A" And APPR_STATUS_CODE <> APPR_STATUS_CODE_OLD Then
                    Dim EXPENSE_TYPE_CODE As String = Absx1.txtFor("EXPENSE_TYPE_CODE").Text
                    Dim rowSPTTYPE1 As DataRow = LookUp("SPTTYPE1", EXPENSE_TYPE_CODE)
                    Dim SECURITY_CODE As String = rowSPTTYPE1.Item("SECURITY_CODE") & ""
                    If SECURITY_CODE <> "" And Not ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
                        EMsg &= vbCr & "You do not have authorization to Approve"
                    Else
                        update_with_approval = True
                    End If
                End If

                If EMsg = "" And EntryMode <> "N" Then
                    Dim STATUS_CODE_check As String = Absx1.optFor("STATUS_CODE").Value
                    If APPR_STATUS_CODE <> "X" And STATUS_CODE_check = "C" And rowSPTCOOP1.Item("STATUS_CODE") & "" <> rowSPTCOOP1.Item("STATUS_CODE", DataRowVersion.Original) & "" Then
                        If MsgBox("Update Anyway?", MsgBoxStyle.YesNo, "Approval Status (" & APPR_STATUS_CODE & ") is not congruous with Accounting Status (" & STATUS_CODE_check & ")") = MsgBoxResult.No Then
                            EMsg &= vbCr & "You May Have to Cancel your Changes and Re-Enter them"
                        End If
                    End If
                End If

            Case "Clone"

                If rowSPTCOOP1.Item("EVENT_GROUP_NO") & "" <> "" Then
                    EMsg &= "Event " & AUTH_NO & " may NOT be maintained here - use Store Focus Events screen"
                End If

                If EMsg = "" Then
                    If MsgBox("Are you sure you want Clone this agreement to a new record?", MsgBoxStyle.YesNo,
          "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If


            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Approve"
                If Not ASCMAIN1.Logical_Lock("SPTCOOP1", AUTH_NO) Then
                    Exit Sub
                End If

                rowSPTCOOP1 = LookUp("SPTCOOP1", AUTH_NO)
                If rowSPTCOOP1.Item("APPR_STATUS_CODE") & "" <> "G" Then
                    EMsg &= "Event " & AUTH_NO & " is NOT Pending Approval"
                End If

                If EMsg = "" Then
                    Dim EXPENSE_TYPE_CODE As String = rowSPTCOOP1.Item("EXPENSE_TYPE_CODE") & ""
                    If Not EXPENSE_TYPE_CODEs_I_may_approve.Contains(EXPENSE_TYPE_CODE) Then
                        EMsg &= vbCr & "You do not have Authorization to Approve this record"
                    End If
                End If

                If EMsg = "" Then
                    Seek_Approval()
                    If AUTH_APPR_NOTES = "" Then
                        'ASCMAIN1.MultiTask_Release()
                        Exit Sub
                    End If
                Else
                    ASCMAIN1.MultiTask_Release()
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Create Template"
                Create_Template()
                Mode_Settings(False)

            Case "Upload Template"
                Upload_Template()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Clone"
                Clone_Record()
                Mode_Settings(False)
                Absx1.txtFor("AUTH_NO").Text = AUTH_NO_new
                Click_Command("View")

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Approve"
                Update_Approval()
                'ASCMAIN1.MultiTask_Release()
                Mode_Settings(False)

            Case "Refresh"
                Refresh_Documents()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Refresh").Visible = Not ScreenMode

                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Create Template").Settings.Enabled = not_iScreenMode
                    .Items("Upload Template").Settings.Enabled = not_iScreenMode
                    .Items("Create Template").Visible = Not InquiryMode
                    .Items("Upload Template").Visible = Not InquiryMode

                    If EntryMode = "V" And ScreenMode And (optShow.Value <> "G") Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    If EntryMode = "V" And ScreenMode And (optShow.Value <> "G") Then
                        .Items("Clone").Visible = Not InquiryMode And False ' sp email 02/14/19
                    Else
                        .Items("Clone").Visible = False
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" And EntryMode <> "E" Then
                        .Items("Update").Settings.Enabled = not_iScreenMode
                        .Items("Cancel").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Update").Settings.Enabled = iScreenMode
                        .Items("Cancel").Settings.Enabled = iScreenMode
                    End If

                    If ScreenMode And EntryMode <> "V" Then
                        .Items("Done").Settings.Enabled = not_iScreenMode
                    Else
                        .Items("Done").Settings.Enabled = iScreenMode
                    End If

                    .Items("Approve").Visible = (Not InquiryMode And EntryMode = "V" And ScreenMode And optShow.Value = "G") ' And ASCMAIN1.USER_SECURITY_CODEs.Contains("SP")

                End With

                .Groups("Totals").Visible = ScreenMode

                Setup_tab0()
            End With
        End If


        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        chkEditNotes.Checked = False
        chkEditVerComment.Checked = False
        If grdSPTCOOPX.Tag = "no refresh" Then

        Else
            chkVerify.Checked = False
        End If



        tab0.Visible = Not ScreenMode

        cmdRefresh.Enabled = Not tf
        grdSPTCOOPS.Enabled = Not tf

        If ScreenMode Then

            Absx1.optFor("APPR_STATUS_CODE").Enabled = True
            Absx1.dteFor("DATE_START").Enabled = True
            Absx1.dteFor("DATE_END").Enabled = True
            Absx1.dteFor("EVENT_DATE_CHANGED").Enabled = True

            Absx1.numFor("OTHER_COST").Enabled = True
            Absx1.numFor("VEHICLE_CPM").Enabled = True


            With grdSPTCOOP3.DisplayLayout.Bands(0)
                .Columns("DIST_OPEN").Hidden = (EntryMode = "E")
                .Columns("DIST_PAID").Hidden = (EntryMode = "E")
            End With

            If EntryMode = "E" Then
                If ASCMAIN1.USER_SECURITY_CODEs.Contains("P1") Then
                    Set_Read_Only_for_ctl(Absx1.optFor("STATUS_CODE"), False)
                Else
                    Set_Read_Only_for_ctl(Absx1.optFor("STATUS_CODE"), True)
                End If
            Else
                Set_Read_Only_for_ctl(Absx1.optFor("STATUS_CODE"), True)
            End If

            splSPTCOOP3.Panel2Collapsed = (EntryMode = "N" Or EntryMode = "E")

            grdSPTCOOP6.Parent = tabPromo.Tabs("Tasks").TabPage

            SET_SPTCODE1()

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                    .Items("Upload Template").Visible = False
                    .Items("Create Template").Visible = False
                End With
            End If

            If EntryMode = "V" Or optShow.Value = "G" Or rowSPTCOOP1.Item("APPR_STATUS_CODE") & "" <> "G" Then
                chkReady4Approval.Visible = False
            Else
                chkReady4Approval.Visible = True
                chkReady4Approval.Checked = False
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            Set_Read_Only(splSpend, (EntryMode = "V"))
            Set_Read_Only(splDocuments, (EntryMode = "V"))
            cmdBrowse.Visible = Not (EntryMode = "V")

            For Each COLUMN_NAME As String In New String() {"INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}
                grdSPTCOOP6.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = (EntryMode = "N" Or EntryMode = "E")
            Next
            grdSPTCOOP6.DisplayLayout.Bands(0).Columns("EDIT").Hidden = Not (EntryMode = "N" Or EntryMode = "E")


            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTCOOP2, grdSPTCOOP3, grdSPTCOOP5, grdSPTCOOP6, grdSPTCOOP8, grdSPTCOOP9, grdSPTCOOPB}
                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name = "grdSPTCOOP9" Or grd.Name = "grdSPTCOOPB" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            .AllowDelete = DefaultableBoolean.True
                        End If
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Else
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                End If
            Next




            If EntryMode = "E" Then
                If rowSPTCOOP1.Item("APPR_STATUS_CODE") = "X" Or rowSPTCOOP1.Item("STATUS_CODE") = "C" Then
                    If Not ASCMAIN1.USER_SECURITY_CODEs.Contains("P1") Then

                        'grdSPTCOOP3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                        'grdSPTCOOP3.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                        'grdSPTCOOP3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False

                        Absx1.optFor("APPR_STATUS_CODE").Enabled = False
                        Absx1.dteFor("DATE_START").Enabled = False
                        Absx1.dteFor("DATE_END").Enabled = False
                        Absx1.dteFor("EVENT_DATE_CHANGED").Enabled = False

                        Absx1.numFor("OTHER_COST").Enabled = False
                        Absx1.numFor("VEHICLE_CPM").Enabled = False

                    End If
                End If
            End If


        Else
            Clear_Record()
            grdSPTCOOP6.Parent = tab0.Tabs("Tasks").TabPage
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SPTCOOP1", "SPTCOOP2", "SPTCOOP3", "SPTCOOP5", "SPTCOOP6", "SPTCOOP7",
             "SPTCOOP8", "SPTCOOP9", "SPTCOOPB",
             "SPTSCHD1", "TATEVNT1", "ASTAUDT1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If grdSPTCOOPX.Tag & "" = "no refresh" Then
            grdSPTCOOPX.Tag = ""
            ' do nothing
            Dim AUTH_NO As String = grdSPTCOOPX.ActiveRow.Cells("AUTH_NO").Value
            ASCMAIN1.sql = sqlSPTCOOPX & $" And SPTCOOP1.AUTH_NO = '{AUTH_NO}'"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE") & ""

                Dim rowSPTCOOPX() As DataRow = dst.Tables("SPTCOOPX").Select($"AUTH_NO = '{AUTH_NO}' AND COLLECTION_CODE = '{COLLECTION_CODE}'")
                If rowSPTCOOPX.Length = 1 Then
                    For Each dcol As DataColumn In row.Table.Columns
                        rowSPTCOOPX(0).Item(dcol.ColumnName) = row.Item(dcol.ColumnName)
                    Next
                End If
            Next
        Else
            'Setup_cbeYP()
            cbeYP.Visible = (optShow.Value = "E")

            isSettled = True
            If Not isInitialScreenLoaded Then
                Refresh_Documents()
                isInitialScreenLoaded = True
            End If
            SET_SPTCODE1()
        End If


    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowSPTCOOP1 = dst.Tables("SPTCOOP1").NewRow
            AUTH_NO = ASCMAIN1.Next_Control_No("SPTCOOP1.AUTH_NO")
            rowSPTCOOP1.Item("AUTH_NO") = AUTH_NO
            rowSPTCOOP1.Item("CUST_CODE") = HFs("CUST_CODE")
            rowSPTCOOP1.Item("AUTH_DATE") = HFs("AUTH_DATE")
            rowSPTCOOP1.Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")
            rowSPTCOOP1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowSPTCOOP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSPTCOOP1.Item("INIT_DATE") = DATETIME_STAMP
            rowSPTCOOP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSPTCOOP1.Item("LAST_DATE") = DATETIME_STAMP
            rowSPTCOOP1.Item("APPR_STATUS_CODE") = "P"
            rowSPTCOOP1.Item("STATUS_CODE") = "O"

            rowSPTCOOP1.Item("EVENT_DATE_CHANGED") = DATETIME_STAMP.Date

            ' rowSPTCOOP1.Item("SALES_DIVISION_CODE") = HFs("SALES_DIVISION_CODE")
            rowSPTCOOP1.Item("EXPENSE_TYPE_CODE") = HFs("EXPENSE_TYPE_CODE")
            dst.Tables("SPTCOOP1").Rows.Add(rowSPTCOOP1)

            Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
            With row
                .Item("TABLE_NAME") = "SPTCOOP1"
                .Item("TABLE_KEY") = AUTH_NO
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("EVENT_TYPE") = "ENTRY"
                .Item("EVENT_DESC") = "Entry Started"
                .Item("EVENT_KEY") = ""
                .Item("FORM_NAME") = Me.Name
            End With
            dst.Tables("TATEVNT1").Rows.Add(row)

        Else
            rowSPTCOOP1 = Fill_Record("SPTCOOP1", Absx1.txtFor("AUTH_NO").Text)
            dst.AcceptChanges()

            rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
            '   rowSOTSDIV1 = LookUp("SOTSDIV1", Absx1.txtFor("SALES_DIVISION_CODE").Text)

            Fill_Records("TATEVNT1", AUTH_NO)
            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
        End If

        STATUS_CODE = rowSPTCOOP1.Item("STATUS_CODE")

        DisplayTotals()

        EnforceConstraints(False)
        Fill_Records("SPTCOOP2", AUTH_NO)
        Fill_Records("SPTCOOP3", AUTH_NO)
        Fill_Records("SPTCOOP5", AUTH_NO)
        Fill_Records("SPTCOOP6", AUTH_NO)
        Fill_Records("SPTCOOP7", AUTH_NO)
        Fill_Records("SPTCOOP8", AUTH_NO)
        Fill_Records("SPTCOOP9", AUTH_NO)
        Fill_Records("ASTATTA2", AUTH_NO)
        Fill_Records("SPTPYMTX", AUTH_NO)
        Sort_grdColumns(grdSPTPYMTX, "PYMT_NO")

        Fill_Records("SPTCOOPB", AUTH_NO)
        For Each row As DataRow In dst.Tables("SPTCOOPB").Select
            row.Item("SEL") = "1"
        Next
        dst.Tables("SPTCOOPB").AcceptChanges()
        Sort_grdColumns(grdSPTCOOPB, "CUST_STORE_NO")

        ASCMAIN1.sql = sqlSPTCOOPX & vbCrLf & $" and SPTCOOP1.AUTH_NO = '{AUTH_NO}'"
        For Each row As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
            Dim AUTH_LNO As Integer = Val(row.Item("AUTH_LNO") & "")

            Dim QTY As Int64 = Val(row.Item("QTY") & "")
            Dim VEHICLE_CPM As Decimal = Val(row.Item("VEHICLE_CPM") & "")
            Dim OTHER_COST As Decimal = Val(row.Item("OTHER_COST") & "")

            Dim TOTAL_AMT As Decimal = QTY * VEHICLE_CPM / 1000 + OTHER_COST
            Dim OPEN_AMT As Decimal = Val(row.Item("OPEN_AMT") & "")
            Dim PAID_AMT As Decimal = Val(row.Item("PAID_AMT") & "")
            Dim rowSPTCOOP3 As DataRow = dst.Tables("SPTCOOP3").Rows.Find(New Object() {AUTH_NO, AUTH_LNO})
            With rowSPTCOOP3
                .Item("TOTAL_AMT") = TOTAL_AMT
                .Item("OPEN_AMT") = OPEN_AMT
                .Item("PAID_AMT") = PAID_AMT
            End With
        Next
        ASCMAIN1.sql = "Select * from ASTAUDT1 where TABLE_NAME = 'SPTCOOP1' and KEY_VALUE = '" & AUTH_NO & "'"
        Fill_Records("ASTAUDT1", , True, ASCMAIN1.sql)
        ASCMAIN1.sql = "Select * from ASTAUDT1 where TABLE_NAME = 'SPTCOOP3' and KEY_VALUE = '" & AUTH_NO & "'"
        Fill_Records("ASTAUDT1", , False, ASCMAIN1.sql)

        Sort_grdColumns(grdASTAUDT1, "INIT_DATE".ToLower)

        EnforceConstraints(True)

        Setup_Retail_Weeks()


        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Dim OPS_YYYYPP As String = Format(rowSPTCOOP1.Item("DATE_START"), "yyyyMM")
        rowSPTCOOP1.Item("OPS_YYYYPP") = OPS_YYYYPP

        Dim OPS_YYYYPP_ACCRUE As String = ""
        If rowSPTCOOP1.Item("DATE_ACCRUE") & "" <> "" Then
            OPS_YYYYPP_ACCRUE = Format(rowSPTCOOP1.Item("DATE_ACCRUE"), "yyyyMM")
        End If
        rowSPTCOOP1.Item("OPS_YYYYPP_ACCRUE") = OPS_YYYYPP_ACCRUE

        BeginTrans()

        If EntryMode = "E" Then
            If STATUS_CODE <> Absx1.optFor("STATUS_CODE").Value Then
                If STATUS_CODE = "O" Then
                    ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CLSD", "Contract Closed", "")
                Else
                    ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "REOP", "Contract Re-Opened", "")
                End If
            End If

            If rowSPTCOOP1.Item("APPR_STATUS_CODE") & "" <> rowSPTCOOP1.Item("APPR_STATUS_CODE", DataRowVersion.Original) & "" Then
                ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPRSTA", $"Appr Status Manually Changed from {rowSPTCOOP1.Item("APPR_STATUS_CODE", DataRowVersion.Original)} to {rowSPTCOOP1.Item("APPR_STATUS_CODE")}", "")
            End If
        End If

        If EntryMode = "N" Then
            ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "ADD", "Contract Created", "")
        Else
            ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CHG", "Contract Updated", "")
        End If

        If EntryMode = "E" Then Check_Changed_Fields()
        If EntryMode <> "N" Then Delete_Records()

        Dim SQLD As String = "AUTH_NO = '" & AUTH_NO & "'"
        INIT_LAST("SPTCOOP1", False, , True)

        Update_Record_TDA("SPTCOOP1", SQLD)
        Update_Record_TDA("SPTCOOP2", SQLD)
        Update_Record_TDA("SPTCOOP3", SQLD)
        Update_Record_TDA("SPTCOOP5", SQLD)
        Update_Record_TDA("SPTCOOP6", SQLD)
        Update_Record_TDA("SPTCOOP7", SQLD)
        Update_Record_TDA("SPTCOOP8", SQLD)
        Update_Record_TDA("SPTCOOP9", SQLD)

        ASCDATA1.DeleteRows(dst.Tables("SPTCOOPB"), "ISNULL(SEL,'0')<>'1'")
        Update_Record_TDA("SPTCOOPB", SQLD)

        Update_Record_TDA("ASTAUDT1")
        '   Update_Record_TDA("TATEVNT1")

        If update_with_approval Then
            Update_Approval()
        End If

        CommitTrans("Update Complete")

    End Sub

    Sub Delete_Record()
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        ' Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"SPTCOOP1", "SPTCOOP2", "SPTCOOP3", "SPTCOOP5", "SPTCOOP6", "SPTCOOP7", "SPTCOOP8", "SPTCOOP9", "SPTCOOPB"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where AUTH_NO = '" & AUTH_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("AUTH_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SPTCOOP1"
            E.COLUMN_NAME = "AUTH_NO"
            E.CODE_VALUE = Absx1.txtFor("AUTH_NO").Text
            E.DESC_VALUE = Absx1.txtFor("CUST_CODE").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "SPTCOOP1"
        E.TABLE_KEY_CAPTION = "Co-Op Advertising Spend Authorization"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("AUTH_NO").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("CUST_CODE").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME

            'Case "PO_ORDER_NO"
            '    If InquiryMode Then
            '        If optStatus.Value = "O" Then
            '            sql_where = " AND PO_ORDER_NO in (Select DISTINCT PO_ORDER_NO from POTORDR2 where PO_STATUS = 'O') "
            '        End If
            '    Else
            '        sql_where = " AND PO_ORDER_NO in (Select DISTINCT PO_ORDER_NO from POTORDR2 where PO_STATUS = 'O') "
            '    End If
            '    If Absx1.txtFor("VEND_CODE").Text <> "" Then
            '        sql_where &= " AND VEND_CODE = '" & Replace(Absx1.txtFor("VEND_CODE").Text, "'", "") & "'"
            '    End If
            '    If Absx1.txtFor("PO_REFERENCE").Text <> "" Then
            '        ' HOW DO WE PROTECT AGAINST SINGLE QUOTES?
            '        sql_where &= " AND PO_REFERENCE like '" & Replace(Absx1.txtFor("PO_REFERENCE").Text, "'", "") & "%'"
            '    End If
            '    If Absx1.txtFor("PO_SPEC_ORDR_NO").Text <> "" Then
            '        sql_where &= " AND PO_SPEC_ORDR_NO like '" & Replace(Absx1.txtFor("PO_SPEC_ORDR_NO").Text, "'", "") & "%'"
            '    End If

            Case "VEND_CODE"
                sql_where = "VEND_TYPE = 'S'"
        End Select

    End Sub


#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTCOOPX, "SSBBBB", "Show Filter", "Show GroupBox", "Move to Pending", "Move to Preliminary", "Approve", "Verify as Open", "Zero Out (Close)", "Edit Promo", "Copy Last Comment")
        Load_Popup_Menu(grdSPTCODE1, "BB", "Select All", "De-Select All")
        Load_Popup_Menu(grdSPTCOOPG, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSPTCOOP9, "B", "Load Stores")
        Load_Popup_Menu(grdSPTCOOPB, "B", "Load All Stores")

    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        If e.SourceControl.Name = "eventMonthView" Then
            Dim umv As UltraWinSchedule.UltraMonthViewSingle = DirectCast(e.SourceControl, Infragistics.Win.UltraWinSchedule.UltraMonthViewSingle)
            If umv.CalendarInfo.SelectedAppointments.Count = 1 Then
                apptEdit = umv.CalendarInfo.SelectedAppointments(0)
            Else
                apptEdit = Nothing
            End If
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If


        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.SourceControl.Name

            Case "grdSPTCOOP9"
                tlb_btn = tlb_pop.Tools("Load Stores")
                If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_btn.SharedProps.Visible = True
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

            Case "grdSPTCOOPB"
                tlb_btn = tlb_pop.Tools("Load All Stores")
                If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_btn.SharedProps.Visible = True
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

            Case "grdSPTCOOPX"

                ' G = Pending
                ' P = Preliminary
                ' O = Open & Approved

                If Not InquiryMode And (optShow.Value = "P" Or optShow.Value = "G") Then
                    tlb_pop.Tools("Move to Pending").SharedProps.Visible = (optShow.Value = "P")
                    tlb_pop.Tools("Approve").SharedProps.Visible = True
                    tlb_pop.Tools("Move to Pending").SharedProps.Visible = False
                    tlb_pop.Tools("Move to Preliminary").SharedProps.Visible = False
                ElseIf Not InquiryMode And (optShow.Value = "A") Then
                    tlb_pop.Tools("Approve").SharedProps.Visible = False
                    tlb_pop.Tools("Move to Pending").SharedProps.Visible = True
                    tlb_pop.Tools("Move to Preliminary").SharedProps.Visible = True
                Else
                    tlb_pop.Tools("Move to Pending").SharedProps.Visible = False
                    tlb_pop.Tools("Approve").SharedProps.Visible = False
                    tlb_pop.Tools("Move to Pending").SharedProps.Visible = False
                    tlb_pop.Tools("Move to Preliminary").SharedProps.Visible = False
                End If

                tlb_pop.Tools("Verify as Open").SharedProps.Visible = Not InquiryMode And (optShow.Value = "O") And chkVerify.Checked
                tlb_pop.Tools("Zero Out (Close)").SharedProps.Visible = Not InquiryMode And (optShow.Value = "O") And chkVerify.Checked And ASCMAIN1.USER_SECURITY_CODEs.Contains("P1")
                tlb_pop.Tools("Edit Promo").SharedProps.Visible = Not InquiryMode And (optShow.Value = "O") And chkVerify.Checked

                tlb_pop.Tools("Copy Last Comment").SharedProps.Visible = Not InquiryMode And chkEditVerComment.Checked And (VERIFIED_AS_OPEN_COMMENTS_last <> "")



        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdSPTCOOP9"
                '    tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                '        tlb_btn.SharedProps.Visible = True
                '    Else
                '        tlb_btn.SharedProps.Visible = False
                '    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If e.Tool.Key = "Edit Appointment" Then

            'Dim appt As UltraWinSchedule.Appointment = eventMonthView.GetAppointmentFromPoint(e.Cursor.Position)
            If apptEdit IsNot Nothing Then
                Dim SCHED_NO As String = apptEdit.Tag
                Edit_Appointment(SCHED_NO)
            End If
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdSPTCODE1.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Load Stores"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

                ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "'"
                For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CUST_STORE_NO As String = ROW.Item("CUST_STORE_NO")
                    If dst.Tables("SPTCOOP9").Rows.Find(New String() {AUTH_NO, CUST_STORE_NO}) Is Nothing Then
                        Dim rowSPTCOOP9 As DataRow = dst.Tables("SPTCOOP9").NewRow
                        Dim rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                        rowSPTCOOP9.Item("AUTH_NO") = AUTH_NO
                        rowSPTCOOP9.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSPTCOOP9.Item("IN_STORE") = "0"
                        rowSPTCOOP9.Item("RETAIL_GOAL") = 0
                        rowSPTCOOP9.Item("CUST_STORE_LOCATION") = rowARTCUST2.ITEM("CUST_STORE_LOCATION")
                        dst.Tables("SPTCOOP9").Rows.Add(rowSPTCOOP9)
                    End If
                Next
                Sort_grdColumns(grdSPTCOOP9, "CUST_STORE_NO")


            Case "Load All Stores"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

                ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = '" & CUST_CODE & "'"
                For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CUST_STORE_NO As String = ROW.Item("CUST_STORE_NO")
                    If dst.Tables("SPTCOOPB").Rows.Find(New String() {AUTH_NO, CUST_STORE_NO}) Is Nothing Then
                        Dim rowSPTCOOPB As DataRow = dst.Tables("SPTCOOPB").NewRow
                        Dim rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                        rowSPTCOOPB.Item("AUTH_NO") = AUTH_NO
                        rowSPTCOOPB.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSPTCOOPB.Item("SEL") = "0"
                        rowSPTCOOPB.Item("CUST_STORE_LOCATION") = rowARTCUST2.ITEM("CUST_STORE_LOCATION")
                        rowSPTCOOPB.Item("CUST_STORE_NAME") = rowARTCUST2.ITEM("CUST_STORE_NAME")
                        dst.Tables("SPTCOOPB").Rows.Add(rowSPTCOOPB)
                    End If
                Next
                Sort_grdColumns(grdSPTCOOPB, "CUST_STORE_NO")

            Case "Copy Last Comment"

                If grdSPTCOOPX.Selected.Rows.Count = 0 Then
                    MsgBox("Nothing Selected to " & e.Tool.Key, MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                Else
                    VERIFIED_AS_OPEN_COMMENTS_last_do_not_save = True
                    For Each grow As UltraWinGrid.UltraGridRow In grdSPTCOOPX.Selected.Rows
                        grow.Cells("VERIFIED_AS_OPEN_COMMENTS").Value = VERIFIED_AS_OPEN_COMMENTS_last
                        grow.Update()
                    Next
                    VERIFIED_AS_OPEN_COMMENTS_last_do_not_save = False
                End If

            Case "Move to Pending", "Move to Preliminary", "Approve", "Verify as Open", "Zero Out (Close)", "Edit Promo"
                If grdSPTCOOPX.Selected.Rows.Count = 0 Then
                    If grdSPTCOOPX.ActiveRow IsNot Nothing Then
                        grdSPTCOOPX.ActiveRow.Selected = True
                    End If
                End If

                If grdSPTCOOPX.Selected.Rows.Count = 0 Then
                    MsgBox("Nothing Selected to " & e.Tool.Key, MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                Else
                    Dim cannot_proceed_msg As String = ""
                    Dim VERIFIED_AS_OPEN_NOTES As String = ""
                    Dim AUTH_NOs As New Dictionary(Of String, Decimal)
                    Dim can_proceed As Boolean = True
                    Dim TOTAL_AMTs As Decimal = 0

                    Dim AUTH_NOsVerified As New List(Of String)

                    For Each grow As UltraWinGrid.UltraGridRow In grdSPTCOOPX.Selected.Rows
                        Dim EXPENSE_TYPE_CODE As String = grow.Cells("EXPENSE_TYPE_CODE").Value
                        Dim AUTH_NO As String = grow.Cells("AUTH_NO").Value

                        If AUTH_NOs.ContainsKey(AUTH_NO) Then
                        Else
                            rowSPTCOOP1 = LookUp("SPTCOOP1", AUTH_NO)
                            Dim VEHICLE_CPM As Decimal = Val(rowSPTCOOP1.Item("VEHICLE_CPM") & "")
                            Dim QTY As Int64 = Val(rowSPTCOOP1.Item("QTY") & "")
                            Dim OTHER_COST As Decimal = Val(rowSPTCOOP1.Item("OTHER_COST") & "")

                            Dim TOTAL_AMT As Decimal = VEHICLE_CPM * QTY / 1000 + OTHER_COST
                            TOTAL_AMTs += TOTAL_AMT

                            Dim APPR_STATUS_CODE As String = rowSPTCOOP1.Item("APPR_STATUS_CODE") & ""
                            Dim VERIFIED_AS_OPEN As String = rowSPTCOOP1.Item("VERIFIED_AS_OPEN") & ""

                            ' CHECK THAT CURRENT APPR_STATUS_CODE MATCHES GRID STATUS
                            If grow.Cells("APPR_STATUS_CODE").Value <> APPR_STATUS_CODE Then
                                cannot_proceed_msg = vbCrLf & vbCrLf & "Promo " & AUTH_NO & " has changed Approval Status - refresh grid"
                                can_proceed = False
                                Exit For
                            End If

                            If e.Tool.Key = "Verify as Open" Or e.Tool.Key = "Zero Out (Close)" Then
                                If VERIFIED_AS_OPEN = "1" Then
                                    'MsgBox("Promo " & AUTH_NO & " is already Verified as Open", MsgBoxStyle.OkOnly, "Cannot Proceed")
                                    'can_proceed = False
                                    'Exit For
                                End If
                                If VERIFIED_AS_OPEN_NOTES = "" Then
                                    VERIFIED_AS_OPEN_NOTES = grow.Cells("VERIFIED_AS_OPEN_NOTES").Value & ""
                                End If
                            ElseIf e.Tool.Key = "Edit Promo" Then


                            Else

                                If rowSPTCOOP1.Item("EVENT_GROUP_NO") & "" <> "" Then
                                    cannot_proceed_msg = vbCrLf & vbCrLf & "At least 1 Contract (" & AUTH_NO & ") has an Event Group, And therefore may Not use " & e.Tool.Key
                                    can_proceed = False
                                End If

                                If rowSPTCOOP1.Item("VERIFIED_AS_OPEN_COMMENTS") & "" <> "" Then
                                    AUTH_NOsVerified.Add(AUTH_NO)
                                End If


                                If APPR_STATUS_CODE = "A" Then
                                    If e.Tool.Key = "Move to Pending" Or e.Tool.Key = "Move to Preliminary" Then
                                        ' ok to proceed
                                    Else
                                        cannot_proceed_msg = vbCrLf & vbCrLf & "Contract " & AUTH_NO & " may Not use " & e.Tool.Key
                                        can_proceed = False
                                    End If

                                ElseIf APPR_STATUS_CODE = "G" Or APPR_STATUS_CODE = "P" Then
                                    If e.Tool.Key = "Move to Pending" Or e.Tool.Key = "Move to Preliminary" Then
                                        cannot_proceed_msg = vbCrLf & vbCrLf & "Contract " & AUTH_NO & " Is Not Pending Approval"
                                        can_proceed = False
                                        Exit For
                                    Else
                                        ' ok to proceed
                                    End If
                                End If
                            End If

                            If e.Tool.Key = "Edit Promo" Then

                                If AUTH_NOs.Count > 0 Then
                                    cannot_proceed_msg = vbCrLf & vbCrLf & "You must Select 1 And only 1 Auth To Edit"
                                    'MsgBox("You must Select 1 And only 1 Auth To Edit", MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                                    can_proceed = False
                                    Exit For
                                Else
                                    AUTH_NOs.Add(AUTH_NO, TOTAL_AMT)
                                End If

                            Else

                                If ((e.Tool.Key = "Verify as Open" Or e.Tool.Key = "Zero Out (Close)") And EXPENSE_TYPE_CODEs_I_may_verify.Contains(EXPENSE_TYPE_CODE)) Or
                               (Not (e.Tool.Key = "Verify as Open" Or e.Tool.Key = "Zero Out (Close)") And EXPENSE_TYPE_CODEs_I_may_approve.Contains(EXPENSE_TYPE_CODE)) Then

                                    If Not ASCMAIN1.Logical_Lock("SPTCOOP1", AUTH_NO) Then
                                        cannot_proceed_msg = vbCrLf & vbCrLf & "Could Not lock Auth " & AUTH_NO
                                        can_proceed = False
                                        Exit For
                                    End If

                                    AUTH_NOs.Add(AUTH_NO, TOTAL_AMT)
                                Else
                                    can_proceed = False
                                    Exit For
                                End If
                            End If
                        End If

                    Next

                    If Not can_proceed Then
                        MsgBox("Cannot Proceed" & cannot_proceed_msg, MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                    Else

                        If e.Tool.Key = "Verify as Open" Then
                            AUTH_APPR_NOTES = ""


                            Dim LBL As String = "You are verifying " & CStr(AUTH_NOs.Count) & " Contracts at Once" _
                                                & vbCrLf & "Total Amount Is " & Format(TOTAL_AMTs, "$#,##0.00") _
                                                & vbCrLf & vbCrLf & "Enter Notes To Record With this Verification" _
                                                & IIf(chkKeepComment.Checked, "(Note - Contracts With Notes will retail their Notes)", "")


                            VERIFIED_AS_OPEN_NOTES = "Verified"

                            'If VERIFIED_AS_OPEN_NOTES = "" Then
                            '    VERIFIED_AS_OPEN_NOTES = "Verified"
                            'Else
                            '    VERIFIED_AS_OPEN_NOTES = ""
                            'End If

                            AUTH_APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, "OK To Verify the Open Amount Of this(these) Contract(s)?", False, 60, VERIFIED_AS_OPEN_NOTES)

                            If ASCMAIN1.response = -1 Then
                                ' USER CLICKED CANCEL
                            Else
                                If AUTH_APPR_NOTES <> "" Then
                                    BeginTrans()
                                    DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                                    For Each AUTH_NO As String In AUTH_NOs.Keys
                                        Verify_Record(AUTH_NO, AUTH_NOs(AUTH_NO))
                                    Next
                                    CommitTrans("Verification Complete")

                                    ' Refresh_Documents()
                                End If
                            End If


                        ElseIf e.Tool.Key = "Zero Out (Close)" Then
                            AUTH_APPR_NOTES = ""

                            Dim LBL As String = "You are Zeroing Out (Closing) " & vbCrLf & CStr(AUTH_NOs.Count) & " Contracts at Once" _
                                                & vbCrLf & "Total Amount Is " & Format(TOTAL_AMTs, "$#,##0.00") _
                                                & vbCrLf & vbCrLf & "Enter Notes To Record With this Closure"

                            AUTH_APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, "OK To Close the Open Amount Of this(these) Contract(s)?", False, 60, "Zero Out")

                            If ASCMAIN1.response = -1 Then
                                ' USER CLICKED CANCEL
                            Else
                                If AUTH_APPR_NOTES <> "" Then
                                    BeginTrans()
                                    DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                                    For Each AUTH_NO As String In AUTH_NOs.Keys
                                        Close_Record(AUTH_NO, AUTH_NOs(AUTH_NO))
                                    Next
                                    CommitTrans("Verification Complete")

                                    ' Refresh_Documents()
                                End If
                            End If

                        ElseIf e.Tool.Key = "Edit Promo" Then
                            grdSPTCOOPX.Tag = "no refresh"
                            Absx1.txtFor("AUTH_NO").Text = AUTH_NOs.Keys(0)
                            Click_Command("Edit")
                            If Not ScreenMode Then
                                grdSPTCOOPX.Tag = ""
                            End If
                        Else
                            AUTH_APPR_NOTES = ""

                            Dim actionMesg = ""
                            Dim APPROVAL_STATUS_new As String = ""
                            Dim defaultNote As String = ""
                            If e.Tool.Key = "Approve" Then
                                actionMesg = "Approving"
                                APPROVAL_STATUS_new = "A"
                                defaultNote = "Approved"
                            ElseIf e.Tool.Key = "Move to Pending" Then
                                actionMesg = "Moving To Pending"
                                APPROVAL_STATUS_new = "G"
                                defaultNote = "Moving To Pending"
                            ElseIf e.Tool.Key = "Move to Preliminary" Then
                                actionMesg = "Moving To Preliminary"
                                APPROVAL_STATUS_new = "P"
                                defaultNote = "Move to Preliminary"
                            Else
                                ' Not sure what else there would be
                            End If

                            If actionMesg = "" Then
                                MsgBox("There Is an issue With this Action: " & e.Tool.Key, MsgBoxStyle.OkOnly)
                            Else
                                actionMesg = Chr(34) & actionMesg & Chr(34)
                                Dim contracts As String = ""
                                If AUTH_NOs.Count = 1 Then
                                    contracts = CStr(AUTH_NOs.Count) & " Contract"
                                Else
                                    contracts = CStr(AUTH_NOs.Count) & " Contracts at Once"
                                End If



                                Dim lblv As String = ""
                                If AUTH_NOsVerified.Count > 1 And e.Tool.Key <> "Approve" Then
                                    lblv = vbCrLf & vbCrLf & "***************************" & vbCrLf & $"Please Note:" & vbCrLf & $"There are {AUTH_NOsVerified.Count}" & vbCrLf & " *Previously Verified*" & vbCrLf & "Contracts in the range selected" & vbCrLf & "***************************"
                                End If

                                Dim LBL As String = $"You are {actionMesg} " & vbCrLf & contracts _
                                                & vbCrLf & "Total Amount is " & Format(TOTAL_AMTs, "$#,##0.00") & lblv _
                                                & vbCrLf & vbCrLf & "Enter Notes to Record with this Action"

                                AUTH_APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, $"OK To '{e.Tool.Key}' this Contract?", False, 60, defaultNote)
                                If ASCMAIN1.response = -1 Then
                                    ' USER CLICKED CANCEL
                                Else
                                    If AUTH_APPR_NOTES <> "" Then
                                        BeginTrans()
                                        DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                                        For Each AUTH_NO As String In AUTH_NOs.Keys

                                            Dim rowSPTCOOP1 As DataRow = LookUp("SPTCOOP1", AUTH_NO)
                                            Dim APPROVAL_STATUS_orig As String = rowSPTCOOP1.Item("APPR_STATUS_CODE")

                                            If e.Tool.Key = "Approve" Then
                                                Approve_Record(AUTH_NO, AUTH_NOs(AUTH_NO), APPROVAL_STATUS_new, APPROVAL_STATUS_orig)
                                            ElseIf e.Tool.Key = "Move to Pending" Then
                                                Approve_Record(AUTH_NO, AUTH_NOs(AUTH_NO), APPROVAL_STATUS_new, APPROVAL_STATUS_orig)
                                            ElseIf e.Tool.Key = "Move to Preliminary" Then
                                                Approve_Record(AUTH_NO, AUTH_NOs(AUTH_NO), APPROVAL_STATUS_new, APPROVAL_STATUS_orig)
                                            Else
                                                MsgBox("There is an issue with this Action: " & e.Tool.Key, MsgBoxStyle.OkOnly)
                                            End If
                                        Next
                                        CommitTrans($"{e.Tool.Key} Complete")

                                        Refresh_Documents()
                                    End If
                                End If
                            End If

                        End If
                        ASCMAIN1.MultiTask_Release()
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            'Case "Item Status Inquiry"
            '    Dim VEHICLE_CODE As String = grd.ActiveRow.Cells("VEHICLE_CODE").Text
            '    Dim rowSPTAVEH1 As DataRow = LookUp("SPTAVEH1", VEHICLE_CODE)
            '    If rowSPTAVEH1 IsNot Nothing Then
            '        Context_Launch("View", VEHICLE_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "AUTH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEHICLE_CODE"
                Dim VEHICLE_CODE As String = Absx1.txtFor("VEHICLE_CODE").Text
                Dim AD_SIZE_APPL As String = "0"
                If VEHICLE_CODE <> "" Then
                    Dim rowSPTAVEH1 As DataRow = LookUp("SPTAVEH1", VEHICLE_CODE)
                    If rowSPTAVEH1 IsNot Nothing Then
                        AD_SIZE_APPL = rowSPTAVEH1.Item("AD_SIZE_APPL") & ""
                    End If
                End If
                cmbAD_SIZE.Visible = (AD_SIZE_APPL = "1")
                lblAD_SIZE.Visible = (AD_SIZE_APPL = "1")
                If AD_SIZE_APPL <> "1" Then
                    cmbAD_SIZE.Value = DBNull.Value
                End If
        End Select
    End Sub
    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "AUTH_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "QTY", "VEHICLE_CPM", "OTHER_COST"
                Dim QTY As Int64 = Val(Absx1.numFor("QTY").Value & "")
                Dim VEHICLE_CPM As Decimal = Val(Absx1.numFor("VEHICLE_CPM").Value & "")
                Dim OTHER_COST As Decimal = Val(Absx1.numFor("OTHER_COST").Value & "")
                Dim TOTAL_AMT As Decimal = QTY * VEHICLE_CPM / 1000 + OTHER_COST
                Absx1.numFor("TOTAL_AMT").Value = TOTAL_AMT
            Case "TOTAL_AMT"
                Calculate_OPEN_AMT()
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.opt_ValueChanged(sender, e)
        If Me.IsLoading Or Not ScreenMode Then Exit Sub

        Select Case Absx1.GetABSColumnName(sender)
            Case "APPR_STATUS_CODE"

                If Absx1.optFor("APPR_STATUS_CODE").Value = "X" Then
                    'Absx1.optFor("STATUS_CODE").Value = "C"
                    rowSPTCOOP1.Item("APPR_STATUS_CODE") = "X"
                    rowSPTCOOP1.Item("STATUS_CODE") = "C"
                    Calculate_OPEN_AMT()
                Else
                    'If rowSPTCOOP1.RowState = DataRowState.Modified Then
                    ' THE NEXT 2 LINES ARE MEANT TO PERMIT THE USER TO CHANGE THE APPR_STATUS OF AN EVENT FROM CANCELLED TO SOME OTHER STATUS AFTER INADVERTENTLY CLICKING CANCELLED (WHICH APPARENTLY PERMANENTLY CLOSED THE EVENT BECAUSE OF THE LINES ABOVE)
                    ' BUT IT IS INSTEAD PINNING THE APPR_STATUS TO CANCELLED, WHICH IS JUST AS WELL, SINCE WE DON'T WANT THE USER TO BE ABLE TO CHANGE THE APPR_STATUS TO PENDING OR SOMETHING WHILE THE STATUS_CODE = 'C'
                    'rowSPTCOOP1.Item("STATUS_CODE") = rowSPTCOOP1.Item("STATUS_CODE", DataRowVersion.Original)
                    'Absx1.optFor("STATUS_CODE").Value = rowSPTCOOP1.Item("STATUS_CODE", DataRowVersion.Original)
                    'Calculate_OPEN_AMT()
                    'End If
                End If

            Case "STATUS_CODE"
                Calculate_OPEN_AMT()

        End Select
    End Sub
    Sub Calculate_OPEN_AMT()
        If Not ScreenMode Then Exit Sub
        Dim TOTAL_AMT As Decimal = Val(Absx1.numFor("TOTAL_AMT").Value & "")
        Dim PAID_AMT As Decimal = Val(Absx1.numFor("PAID_AMT").Value & "")
        Dim OPEN_AMT As Decimal = TOTAL_AMT - PAID_AMT
        If OPEN_AMT < 0 Then OPEN_AMT = 0
        If Absx1.optFor("STATUS_CODE").Value & "" <> "O" Then
            OPEN_AMT = 0
        End If
        Absx1.numFor("OPEN_AMT").Value = OPEN_AMT
        Absx1.numFor("CANCEL_AMT").Value = TOTAL_AMT - OPEN_AMT - PAID_AMT
    End Sub
    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "DATE_START"
                If Absx1.dteFor("DATE_START").Value & "" = "" Then
                    Absx1.txtFor("OPS_YYYYWW").Text = ""
                Else
                    Dim DATE_START As Date = Absx1.dteFor("DATE_START").Value
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(DATE_START, "dd-MMM-yyyy") & "'"
                        Dim YW As String = ASCDATA1.GetDataValue
                        If YW <> "" Then
                            Absx1.txtFor("OPS_YYYYWW").Text = YW
                        End If
                    End If

                    'If ScreenMode And (EntryMode = "N") Then
                    '    If Absx1.dteFor("DATE_END").Value & "" = "" Then
                    '        Absx1.dteFor("DATE_END").Value = Absx1.dteFor("DATE_START").Value
                    '    End If
                    'End If
                End If
        End Select
    End Sub
#End Region

#Region "grdSPTCOOP2"

    Private Sub grdSPTCOOP2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Dim ITEM_CODE As String = CStr(e.Cell.Value & "").ToUpper

                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    If e.Cell.Value & "" <> ITEM_CODE Then
                        e.Cell.Row.Cells("ITEM_CODE").Value = ITEM_CODE
                    End If
                    e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                    e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
                End If
        End Select
    End Sub

    Private Sub grdSPTCOOP2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP2.AfterExitEditMode
        'Select Case grdSPTCOOP2.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdSPTCOOP2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP2.AfterRowActivate
        With grdSPTCOOP2.DisplayLayout.Bands(0)
            If grdSPTCOOP2.ActiveRow.IsAddRow Then
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSPTCOOP2.ActiveCell = grdSPTCOOP2.ActiveRow.Cells("ITEM_CODE")
                grdSPTCOOP2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub
    Private Sub grdSPTCOOP2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTCOOP2.BeforeExitEditMode
        If grdSPTCOOP2.ActiveCell Is Nothing Then Exit Sub
        With grdSPTCOOP2.ActiveCell
            Select Case .Column.Key
                Case "ITEM_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Item Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSPTCOOP2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTCOOP2.BeforeRowUpdate
        With grdSPTCOOP2
            If e.Row.Cells("ITEM_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                ' e.Cancel = True MAGIC UNICORN IS A COMMENTED LINE
            Else
                LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("AUTH_NO").Text = "" Then
                    .ActiveRow.Cells("AUTH_NO").Value = Absx1.CtlFor("AUTH_NO").Text
                    .ActiveRow.Cells("AUTH_LNO").Value = Val(dst.Tables("SPTCOOP2").Compute("Max(AUTH_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdSPTCOOP2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP2.ClickCellButton

        If grdSPTCOOP2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                sql_where = "ITEM_STATUS = 'A'"
        End Select
        grdClickCellButton(grdSPTCOOP2, sql_where, False)

    End Sub

    Private Sub grdSPTCOOP2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSPTCOOP2.Error
        grdSPTCOOP2.ActiveRow.CancelUpdate()
    End Sub

#End Region

#Region "grdSPTCOOP3"

    Private Sub grdSPTCOOP3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                '   If cdr IsNot Nothing Then
                Dim ITEM_CODE As String = CStr(e.Cell.Value & "").ToUpper
                ' Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)

                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    If e.Cell.Value & "" <> ITEM_CODE Then
                        e.Cell.Row.Cells("ITEM_CODE").Value = ITEM_CODE
                    End If
                    e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                    e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                    e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
                End If

                '   End If

            Case "COLLECTION_CODE"

                grdCodeDesc(grdSPTCOOP3, "ICTCOLL1", "COLLECTION_CODE", "COLLECTION_NAME")
                If cdr IsNot Nothing Then
                    Dim COLLECTION_CODE As String = e.Cell.Value
                    Dim BRAND_CODE As String = cdr.Item("BRAND_CODE")
                    Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", BRAND_CODE)
                    e.Cell.Row.Cells("BRAND_CODE").Value = BRAND_CODE
                    e.Cell.Row.Cells("BRAND_NAME").Value = rowICTBRAN1.Item("BRAND_NAME")
                Else
                    grdSPTCOOP3.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "DIST_AMT"


        End Select
    End Sub

    Private Sub grdSPTCOOP3_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP3.AfterExitEditMode
        'Select Case grdSPTCOOP3.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdSPTCOOP3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP3.AfterRowActivate
        With grdSPTCOOP3.DisplayLayout.Bands(0)
            If grdSPTCOOP3.ActiveRow.IsAddRow Then
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSPTCOOP3.ActiveCell = grdSPTCOOP3.ActiveRow.Cells("COLLECTION_CODE")
                grdSPTCOOP3.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub
    Private Sub grdSPTCOOP3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTCOOP3.BeforeExitEditMode
        If grdSPTCOOP3.ActiveCell Is Nothing Then Exit Sub
        With grdSPTCOOP3.ActiveCell
            Select Case .Column.Key
                Case "COLLECTION_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTCOLL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Collection Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSPTCOOP3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTCOOP3.BeforeRowUpdate
        With grdSPTCOOP3
            If e.Row.Cells("COLLECTION_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                ' e.Cancel = True MAGIC UNICORN IS A COMMENTED LINE
            Else
                LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Collection Code (" & e.Row.Cells("COLLECTION_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Val(e.Row.Cells("DIST_AMT").Text) < 0 Then 'If Val(e.Row.Cells("DIST_AMT").Text) = 0 Then
                MsgBox("Invalid Value entered for Distribution Amount (" & e.Row.Cells("DIST_AMT").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("AUTH_NO").Text = "" Then
                    .ActiveRow.Cells("AUTH_NO").Value = Absx1.CtlFor("AUTH_NO").Text
                    .ActiveRow.Cells("AUTH_LNO").Value = Val(dst.Tables("SPTCOOP3").Compute("Max(AUTH_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdSPTCOOP3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP3.ClickCellButton

        If grdSPTCOOP3.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
            Case "COLLECTION_CODE"
                sql_where = "COLLECTION_STATUS = 'A'"
        End Select
        grdClickCellButton(grdSPTCOOP3, sql_where, False)

    End Sub

    Private Sub grdSPTCOOP3_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSPTCOOP3.Error
        grdSPTCOOP3.ActiveRow.CancelUpdate()
    End Sub

#End Region

#Region "grdSPTCOOP5"

    Private Sub grdSPTCOOP5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP5.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "MODEL_CODE"

                'grdCodeDesc(grdSPTCOOP5, "?", "MODEL_CODE", "MODEL_NAME")
                'If cdr IsNot Nothing Then
                '    Dim AUTH_EXP_CATGY As String = e.Cell.Value
                'Else
                '    grdSPTCOOP5.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                'End If

            Case "ADJ_QTY"

        End Select
    End Sub

    Private Sub grdSPTCOOP5_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP5.AfterExitEditMode

        'Select Case grdSPTCOOP5.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdSPTCOOP5_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP5.AfterRowActivate
        With grdSPTCOOP5.DisplayLayout.Bands(0)
            If grdSPTCOOP5.ActiveRow.IsAddRow Then
                .Columns("MODEL_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSPTCOOP5.ActiveCell = grdSPTCOOP5.ActiveRow.Cells("MODEL_CODE")
                grdSPTCOOP5.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("MODEL_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub
    Private Sub grdSPTCOOP5_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTCOOP5.BeforeExitEditMode
        If grdSPTCOOP5.ActiveCell Is Nothing Then Exit Sub
        With grdSPTCOOP5.ActiveCell
            Select Case .Column.Key
                'Case "MODEL_CODE"
                '    If .Text <> "" Then
                '        If .Value IsNot Nothing Then
                '            .Value = .Text.ToUpper
                '        End If

                '    End If
                '    If .Text <> "" Then
                '        cdr = LookUp("?", .Text)
                '        If cdr Is Nothing Then
                '            ASCMAIN1.Progress("Invalid Model Code (" & .Text & ")")
                '            If .Value IsNot Nothing Then
                '                .Value = ""
                '            End If
                '            e.Cancel = True
                '        End If
                '    End If
            End Select
        End With
    End Sub

    Private Sub grdSPTCOOP5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTCOOP5.BeforeRowUpdate
        With grdSPTCOOP5
            'If e.Row.Cells("MODEL_CODE").Text = "" Then
            '    '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'Else
            '    LookUp("?", e.Row.Cells("MODEL_CODE").Text)
            '    If cdr Is Nothing Then
            '        MsgBox("Invalid Value entered for Model Code (" & e.Row.Cells("MODEL_CODE").Text & ")", _
            '               MsgBoxStyle.OkOnly, "Cannot Update Row")
            '        e.Cancel = True
            '    End If
            'End If

            'If Val(e.Row.Cells("ADJ_QTY").Text) = 0 Then
            '    MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("ADJ_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("AUTH_NO").Text = "" Then
                    .ActiveRow.Cells("AUTH_NO").Value = Absx1.CtlFor("AUTH_NO").Text
                    .ActiveRow.Cells("AUTH_SNO").Value = Val(dst.Tables("SPTCOOP5").Compute("Max(AUTH_SNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdSPTCOOP5_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP5.ClickCellButton

        If grdSPTCOOP5.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "AUTH_EXP_CATGY"
            Case "LOCATION_CODE"
                sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        End Select
        grdClickCellButton(grdSPTCOOP5, sql_where, False)

    End Sub

    Private Sub grdSPTCOOP5_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSPTCOOP5.Error
        grdSPTCOOP5.ActiveRow.CancelUpdate()
    End Sub

#End Region

#Region "grdSPTCOOP6"

    Private Sub grdSPTCOOP6_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP6.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "TASK_TYPE"

                grdCodeDesc(grdSPTCOOP6, "SPTTASK0", "TASK_TYPE", "TASK_DESC")
                If cdr IsNot Nothing Then
                    Dim TASK_TYPE As String = e.Cell.Value
                Else
                    grdSPTCOOP6.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "ADJ_QTY"

        End Select
    End Sub

    Private Sub grdSPTCOOP6_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP6.AfterExitEditMode

        'Select Case grdSPTCOOP6.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdSPTCOOP6_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP6.AfterRowActivate
        With grdSPTCOOP6.DisplayLayout.Bands(0)
            If grdSPTCOOP6.ActiveRow.IsAddRow Then
                For Each C As String In New String() {"TASK_TYPE", "TASK_DESC", "TASK_ASSIGNED_TO", "TASK_DUE_DATE"}
                    .Columns(C).CellActivation = UltraWinGrid.Activation.AllowEdit
                Next

                grdSPTCOOP6.ActiveCell = grdSPTCOOP6.ActiveRow.Cells("TASK_TYPE")
                grdSPTCOOP6.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                For Each C As String In New String() {"TASK_TYPE", "TASK_DESC", "TASK_ASSIGNED_TO", "TASK_DUE_DATE"}
                    .Columns(C).CellActivation = UltraWinGrid.Activation.NoEdit
                Next
            End If
        End With
    End Sub

    Private Sub grdSPTCOOP6_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTCOOP6.BeforeExitEditMode
        If grdSPTCOOP6.ActiveCell Is Nothing Then Exit Sub
        With grdSPTCOOP6.ActiveCell
            Select Case .Column.Key
                Case "TASK_TYPE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("SPTTASK0", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Task Type (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSPTCOOP6_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTCOOP6.BeforeRowUpdate
        With grdSPTCOOP6
            If e.Row.Cells("TASK_TYPE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("SPTTASK0", e.Row.Cells("TASK_TYPE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Task Type (" & e.Row.Cells("TASK_TYPE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            'If Val(e.Row.Cells("ADJ_QTY").Text) = 0 Then
            '    MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("ADJ_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("AUTH_NO").Text = "" Then
                    .ActiveRow.Cells("AUTH_NO").Value = Absx1.CtlFor("AUTH_NO").Text
                    .ActiveRow.Cells("AUTH_TNO").Value = Val(dst.Tables("SPTCOOP6").Compute("Max(AUTH_TNO)", "") & "") + 1
                    .ActiveRow.Cells("INIT_DATE").Value = DATETIME_STAMP
                    .ActiveRow.Cells("INIT_OPER").Value = ASCMAIN1.USER_ID
                End If
            End If
        End With
    End Sub

    Private Sub grdSPTCOOP6_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP6.ClickCellButton

        If grdSPTCOOP6.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "TASK_TYPE"
            Case "LOCATION_CODE"
                sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
            Case "EDIT"
                If e.Cell.Row.IsAddRow Then Exit Sub
                Dim AUTH_TNO As Integer = Val(e.Cell.Row.Cells("AUTH_TNO").Value & "")
                Edit_Task(AUTH_NO & ":" & CStr(AUTH_TNO))
                Exit Sub
        End Select
        grdClickCellButton(grdSPTCOOP6, sql_where, False)

    End Sub

    Private Sub grdSPTCOOP6_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSPTCOOP6.Error
        grdSPTCOOP6.ActiveRow.CancelUpdate()
    End Sub

#End Region

#Region "grdSPTCOOP8"

    Private Sub grdSPTCOOP8_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP8.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "AUTH_EXP_CATGY"

                grdCodeDesc(grdSPTCOOP8, "SPTXCAT1", "AUTH_EXP_CATGY", "AUTH_EXP_DESC")
                If cdr IsNot Nothing Then
                    Dim AUTH_EXP_CATGY As String = e.Cell.Value
                Else
                    grdSPTCOOP8.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "ADJ_QTY"

        End Select
    End Sub

    Private Sub grdSPTCOOP8_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP8.AfterExitEditMode

        'Select Case grdSPTCOOP8.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdSPTCOOP8_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOP8.AfterRowActivate
        With grdSPTCOOP8.DisplayLayout.Bands(0)
            If grdSPTCOOP8.ActiveRow.IsAddRow Then
                .Columns("AUTH_EXP_CATGY").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSPTCOOP8.ActiveCell = grdSPTCOOP8.ActiveRow.Cells("AUTH_EXP_CATGY")
                grdSPTCOOP8.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("AUTH_EXP_CATGY").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub
    Private Sub grdSPTCOOP8_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTCOOP8.BeforeExitEditMode
        If grdSPTCOOP8.ActiveCell Is Nothing Then Exit Sub
        With grdSPTCOOP8.ActiveCell
            Select Case .Column.Key
                Case "AUTH_EXP_CATGY"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("SPTXCAT1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Expense Category Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSPTCOOP8_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTCOOP8.BeforeRowUpdate
        With grdSPTCOOP8
            If e.Row.Cells("AUTH_EXP_CATGY").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("SPTXCAT1", e.Row.Cells("AUTH_EXP_CATGY").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Expense Category Code (" & e.Row.Cells("AUTH_EXP_CATGY").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            'If Val(e.Row.Cells("ADJ_QTY").Text) = 0 Then
            '    MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("ADJ_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("AUTH_NO").Text = "" Then
                    .ActiveRow.Cells("AUTH_NO").Value = Absx1.CtlFor("AUTH_NO").Text
                    .ActiveRow.Cells("AUTH_ENO").Value = Val(dst.Tables("SPTCOOP8").Compute("Max(AUTH_ENO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdSPTCOOP8_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTCOOP8.ClickCellButton

        If grdSPTCOOP8.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "AUTH_EXP_CATGY"
            Case "LOCATION_CODE"
                sql_where = "CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        End Select
        grdClickCellButton(grdSPTCOOP8, sql_where, False)

    End Sub

    Private Sub grdSPTCOOP8_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSPTCOOP8.Error
        grdSPTCOOP8.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()

    End Sub

    Private Sub grdSPTCOOPX_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdSPTCOOPX.BeforeRowUpdate
        If Update_Record_SPTCOOP1() Then
            ' update was successful
        Else
            e.Cancel = True
        End If
    End Sub

    Private Sub grdSPTCOOPX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTCOOPX.DoubleClickRow
        If e.Row.IsDataRow Then
            If chkEditNotes.Checked Or chkEditVerComment.Checked Or chkVerify.Checked Then
                Exit Sub
            End If
            Absx1.txtFor("AUTH_NO").Text = e.Row.Cells("AUTH_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()

        If Not isSettled Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting Up Lead Screen")

        EnforceConstraints(False)
        If optShow.Value = "G" Then
            ASCMAIN1.sql = sqlSPTCOOPX & " and SPTCOOP1.APPR_STATUS_CODE = 'G'" _
                & " and SPTCOOP1.STATUS_CODE = 'O'"
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            grdSPTCOOPX.Text = "Contracts Pending Approval"
        ElseIf optShow.Value = "P" Then
            ASCMAIN1.sql = sqlSPTCOOPX & " and SPTCOOP1.APPR_STATUS_CODE = 'P'" _
                & " and SPTCOOP1.STATUS_CODE = 'O'"
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            grdSPTCOOPX.Text = "Contracts not yet Ready for Approval (Preliminary)"
        ElseIf optShow.Value = "A" Then
            ASCMAIN1.sql = sqlSPTCOOPX & " and SPTCOOP1.APPR_STATUS_CODE = 'A'" _
                & " and SPTCOOP1.STATUS_CODE = 'O'"
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            grdSPTCOOPX.Text = "Contracts Open and Approved"
        ElseIf optShow.Value = "E" Then
            Dim YP As String = cbeYP.Value
            Fill_Records("SPTCOOPX", YP)
            grdSPTCOOPX.Text = "Entered in " & cbeYP.Text
        ElseIf optShow.Value = "O" Then
            ASCMAIN1.sql = sqlSPTCOOPX & " and SPTCOOP1.STATUS_CODE = 'O'"
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            grdSPTCOOPX.Text = "All Open"
        ElseIf optShow.Value = "X" Then
            ASCMAIN1.sql = Replace(sqlSPTCOOPX, "and SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')", "and SPTCOOP1.APPR_STATUS_CODE in ('X')")
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            grdSPTCOOPX.Text = "Cancelled"
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlSPTCOOPX
            Fill_Records("SPTCOOPX", "", True, ASCMAIN1.sql)
            grdSPTCOOPX.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdSPTCOOPX, "AUTH_NO".ToLower)

        Me.Get_SP_Events()

        Me.eventMonthView.CalendarInfo.SelectedDateRanges.Clear()
        Me.eventMonthView.CalendarInfo.SelectedDateRanges.Add(DateTime.Now, 0)

        chkEditNotes.Checked = False
        chkEditVerComment.Checked = False

        chkVerify.Checked = False
        chkVerify.Enabled = (optShow.Value = "O")

        Setup_SPTSCHD1()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub grdSPTCOOPX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOPX.InitializeRow
        If e.Row.IsFilterRow Or Not e.Row.IsDataRow Then
            Exit Sub
        End If

        With e.Row.Cells("APPR_STATUS_CODE")
            Select Case .Value & ""
                Case "A"
                    .Appearance.ForeColor = System.Drawing.Color.Green
                Case "P"
                    .Appearance.ForeColor = System.Drawing.Color.Purple
                Case "G"
                    .Appearance.ForeColor = System.Drawing.Color.Blue
                Case "X"
                    .Appearance.ForeColor = System.Drawing.Color.Red
            End Select
        End With

        Dim DIST_REAL_EXPENSE As Decimal = Val(e.Row.Cells("DIST_REAL_EXPENSE").Value & "")
        Dim DIST_OPEN_AND_PAID As Decimal = Val(e.Row.Cells("DIST_OPEN_AND_PAID").Value & "")
        If DIST_REAL_EXPENSE <> DIST_OPEN_AND_PAID Then
            e.Row.Cells("DIST_REAL_EXPENSE").Appearance = appRedForeColor
        Else
            ' e.Row.Cells("DIST_REAL_EXPENSE").Appearance.ForeColor = Drawing.Color.Empty
        End If



    End Sub

#Region "Form Controls"

    Private Sub eventMonthView_BeforeAppointmentEdit(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinSchedule.BeforeAppointmentEditEventArgs) Handles eventMonthView.BeforeAppointmentEdit
        e.Cancel = True
        'Edit_Appointment(e.Appointment.Tag)
    End Sub

    Private Sub eventMonthView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles eventMonthView.Click
        Setup_SPTSCHD1()
    End Sub

    Private Sub eventMonthView_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles eventMonthView.MouseDoubleClick

        If InquiryMode Then Exit Sub
        If Not e.Button = MouseButtons.Left Then Exit Sub

        Dim point As System.Drawing.Point
        point = New System.Drawing.Point(e.X, e.Y)

        ' Determine where in the control the right button was pressed
        Dim objAppointment As Infragistics.Win.UltraWinSchedule.Appointment
        Dim objNote As Infragistics.Win.UltraWinSchedule.Note
        Dim objWeek As Infragistics.Win.UltraWinSchedule.Week
        Dim objDay As Infragistics.Win.UltraWinSchedule.Day
        Dim objDayOfWeek As Infragistics.Win.UltraWinSchedule.DayOfWeek

        ' See is we clicked an Appointment
        objAppointment = Me.eventMonthView.GetAppointmentFromPoint(e.X, e.Y)
        objDay = Me.eventMonthView.GetDayFromPoint(e.X, e.Y)
        If objAppointment Is Nothing AndAlso objDay Is Nothing Then
            Exit Sub
        End If

        If objDay IsNot Nothing AndAlso (objDay.Date < eventMonthView.CalendarInfo.MinDate OrElse objDay.Date > eventMonthView.CalendarInfo.MaxDate) Then
            MessageBox.Show("The selected date (" & objDay.Date & ") is out of specified Date Range.", "Create Appointment", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        objNote = Me.eventMonthView.GetNoteFromPoint(e.X, e.Y)
        objWeek = Me.eventMonthView.GetWeekFromPoint(e.X, e.Y)
        objDay = Me.eventMonthView.GetDayFromPoint(e.X, e.Y)
        objDayOfWeek = Me.eventMonthView.GetDayOfWeekFromPoint(e.X, e.Y)

        eventMonthView.CalendarInfo.SelectedDateRanges.Clear()
        eventMonthView.CalendarInfo.ActiveDay = objDay

        eventMonthView.CalendarInfo.SelectedDateRanges.Clear()
        eventMonthView.CalendarInfo.SelectedDateRanges.Add(objDay.Date, 0)
        Setup_SPTSCHD1()

        If objAppointment IsNot Nothing Then
            Dim SCHED_NO As String = objAppointment.Tag
            Edit_Appointment(SCHED_NO)
            Exit Sub
        End If

        'Me.eventMonthView.CalendarInfo.DA()
        Using f As New SPFCOOP2
            f.sqlSPTCOOPX = sqlSPTCOOPX
            f.rowSPTCOOPX = dst.Tables("SPTCOOPX").NewRow
            f.rowSPTCOOPX.Item("AUTH_NO") = ASCMAIN1.Next_Control_No("SPTCOOP1.AUTH_NO")
            f.rowSPTCOOPX.Item("DATE_START") = objDay.Date 'eventMonthView.CalendarInfo.ActiveDay.Date
            f.rowSPTCOOPX.Item("DATE_END") = objDay.Date
            f.CUST_CODE = Absx1.txtFor("CUST_CODE").Text

            f.ShowDialog()

            If f.UPDATED Then
                Add_Appointment(f.rowSPTCOOPX)
                Dim SCHED_NO As String = f.rowSPTCOOPX.Item("SCHED_NO")
                Dim row As DataRow = dst.Tables("SPTSCHD1").Rows.Find(SCHED_NO)
                If row IsNot Nothing Then
                    row.ItemArray = f.rowSPTCOOPX.ItemArray
                Else
                    row = dst.Tables("SPTSCHD1").NewRow
                    row.ItemArray = f.rowSPTCOOPX.ItemArray
                    dst.Tables("SPTSCHD1").Rows.Add(row)
                End If
                dst.Tables("SPTSCHD1").AcceptChanges()
            End If
        End Using
    End Sub

    Private Sub grdSPTSCHDL_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTCOOPS.AfterRowsDeleted
        Try
            MyBase.BeginTrans()
            Update_Record_TDA("SPTSCHDL")
            MyBase.CommitTrans()
        Catch ex As Exception
            MessageBox.Show("Error Deleting Statuses: " & ex.Message, "Delete", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MyBase.Rollback()
        End Try

    End Sub

    Private Sub grdSPTSCHDL_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSPTCOOPS.AfterRowUpdate

        Dim SCHED_CODE As String = e.Row.Cells("SCHED_CODE").Value
        Dim argbBackColor() As String = e.Row.Cells("SCHED_BACKCOLOR").Value.ToString.Split(".")
        Dim argbForeColor() As String = e.Row.Cells("SCHED_FORECOLOR").Value.ToString.Split(".")
        Dim SCHED_NO As String = String.Empty

        Dim rowAppts As DataRow() = dst.Tables("SPTSCHD1").Select("SCHED_CODE = '" & SCHED_CODE & "'")
        If rowAppts IsNot Nothing AndAlso rowAppts.Length > 0 Then
            Dim tbl As DataTable = dst.Tables("SPTSCHD1").Clone
            For Each row As DataRow In rowAppts
                tbl.ImportRow(row)
            Next

            For Each appt As Infragistics.Win.UltraWinSchedule.Appointment In eventMonthView.CalendarInfo.Appointments
                SCHED_NO = appt.Tag

                If tbl.Select("SCHED_NO = '" & SCHED_NO & "'").Length > 0 Then
                    appt.Appearance.BackColor = System.Drawing.Color.FromArgb(argbBackColor(0), argbBackColor(1), argbBackColor(2), argbBackColor(3))
                    appt.Appearance.ForeColor = System.Drawing.Color.FromArgb(argbForeColor(0), argbForeColor(1), argbForeColor(2), argbForeColor(3))
                End If
            Next

        End If

        Try
            MyBase.BeginTrans()
            Update_Record_TDA("SPTSCHDL")
            MyBase.CommitTrans()
        Catch ex As Exception
            MessageBox.Show("Error Updating Statuses: " & ex.Message, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            MyBase.Rollback()
        End Try

        grdSPTCOOPG.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow, True)
    End Sub

    Private Sub grdSPTSCHDL_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTCOOPS.BeforeRowUpdate

        If e.Row.Cells("BACKCOLOR").Value Is Nothing Then
            e.Row.Cells("BACKCOLOR").Value = System.Drawing.Color.White
        End If

        If e.Row.Cells("FORECOLOR").Value Is Nothing Then
            e.Row.Cells("FORECOLOR").Value = System.Drawing.Color.White
        End If

        Dim BACKCOLOR As System.Drawing.Color = e.Row.Cells("BACKCOLOR").Value
        Dim FORECOLOR As System.Drawing.Color = e.Row.Cells("FORECOLOR").Value
        Dim SCHED_DESC As String = (e.Row.Cells("SCHED_DESC").Value & String.Empty).ToString.Trim

        If BACKCOLOR = FORECOLOR Then
            e.Cancel = True
            MessageBox.Show("Backcolor and Forecolor must be different.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        If SCHED_DESC.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Description is required.", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If

        Dim SCHED_BACKCOLOR As String = BACKCOLOR.A & "." & BACKCOLOR.R & "." & BACKCOLOR.G & "." & BACKCOLOR.B
        Dim SCHED_FORECOLOR As String = FORECOLOR.A & "." & FORECOLOR.R & "." & FORECOLOR.G & "." & FORECOLOR.B

        e.Row.Cells("SCHED_BACKCOLOR").Value = SCHED_BACKCOLOR
        e.Row.Cells("SCHED_FORECOLOR").Value = SCHED_FORECOLOR

        If e.Row.IsAddRow Then
            Dim codeFound As Boolean = False
            Dim SCHED_CODES As String = "ABCDEFGHIJLKMNOPQRSTUVWXYZ0123456789"
            For Each SCHED_CODE As Char In SCHED_CODES
                If dst.Tables("SPTSCHDL").Select("SCHED_CODE = '" & SCHED_CODE.ToString.Trim & "'").Length = 0 Then
                    e.Row.Cells("SCHED_CODE").Value = SCHED_CODE.ToString.Trim
                    codeFound = True
                    Exit For
                End If
            Next
            If Not codeFound Then
                e.Cancel = True
                MessageBox.Show("There are no codes available. Please contact ABS.", "Add Reason", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        End If
    End Sub

    Private Sub grdSPTSCHDL_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOPS.InitializeRow
        Dim argbBackColor() As String = e.Row.Cells("SCHED_BACKCOLOR").Value.ToString.Split(".")
        Dim argbForeColor() As String = e.Row.Cells("SCHED_FORECOLOR").Value.ToString.Split(".")

        e.Row.Cells("SCHED_BACKCOLOR").SelectedAppearance.ForeColor = System.Drawing.Color.Transparent
        e.Row.Cells("SCHED_FORECOLOR").SelectedAppearance.ForeColor = System.Drawing.Color.Transparent

        If e.Row.Cells("BACKCOLOR").Value IsNot Nothing Then
            Dim BACKCOLOR As System.Drawing.Color = e.Row.Cells("BACKCOLOR").Value
            Dim SCHED_BACKCOLOR As String = BACKCOLOR.A & "." & BACKCOLOR.R & "." & BACKCOLOR.G & "." & BACKCOLOR.B
            argbBackColor = SCHED_BACKCOLOR.Split(".")
        End If

        If e.Row.Cells("FORECOLOR").Value IsNot Nothing Then
            Dim FORECOLOR As System.Drawing.Color = e.Row.Cells("FORECOLOR").Value
            Dim SCHED_FORECOLOR As String = FORECOLOR.A & "." & FORECOLOR.R & "." & FORECOLOR.G & "." & FORECOLOR.B
            argbForeColor = SCHED_FORECOLOR.Split(".")
        End If

        If argbBackColor.Length <> 4 OrElse argbForeColor.Length <> 4 Then
            Exit Sub
        End If

        'With e.Row.Cells("SCHED_DESC")
        '    .Appearance.BackColor = System.Drawing.Color.FromArgb(argbBackColor(0), argbBackColor(1), argbBackColor(2), argbBackColor(3))
        '    e.Row.Cells("BACKCOLOR").Value = .Appearance.BackColor
        '    .Appearance.ForeColor = System.Drawing.Color.FromArgb(argbForeColor(0), argbForeColor(1), argbForeColor(2), argbForeColor(3))
        '    e.Row.Cells("FORECOLOR").Value = .Appearance.ForeColor
        'End With
    End Sub

    Private Sub cmdRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefresh.Click
        Get_SP_Events()
    End Sub

#End Region

#Region "Form Procedures"

    Sub Setup_SPTSCHD1()
        Dim SCHED_DATE As Date = eventMonthView.CalendarInfo.ActiveDay.Date
        '  dvwSPTSCHD1.RowFilter = "SCHED_DATE <= #" & Format(SCHED_DATE, "MM/dd/yyyy") & "# AND SCHED_DATE_END >= #" & Format(SCHED_DATE, "MM/dd/yyyy") & "#"
        Dim dvw As DataView = DirectCast(grdSPTCOOPG.DataSource, DataTable).DefaultView
        dvw.RowFilter = "DATE_START = #" & Format(SCHED_DATE, "MM/dd/yyyy") & "#"
        grdSPTCOOPG.Text = "Scheduled Events for " & Format(SCHED_DATE, "MM/dd/yyyy")
    End Sub

    Sub Remove_Appointment(ByVal SCHED_NO As String)
        eventMonthView.CalendarInfo.Appointments.Remove(SCHED_events(SCHED_NO))
        SCHED_events.Remove(SCHED_NO)
    End Sub

    Sub Load_Events_into_Calendar()
        eventMonthView.CalendarInfo.Appointments.Clear()
        For Each rowSPTCOOPG As DataRow In dst.Tables("SPTCOOPG").Rows
            Add_Appointment(rowSPTCOOPG)
        Next
    End Sub

    Sub Add_Appointment(ByVal rowSPTCOOPG As DataRow)
        Dim SUBJECT As String = rowSPTCOOPG.Item("CUST_CODE") _
                                & ":" & rowSPTCOOPG.Item("BRAND_CODE") _
                                & ":" & rowSPTCOOPG.Item("BOOKING_NAME")

        Dim APPR_STATUS_CODE As String = rowSPTCOOPG.Item("APPR_STATUS_CODE") & ""
        Dim EVENT_key As String = rowSPTCOOPG.Item("AUTH_NO") & "." & rowSPTCOOPG.Item("AUTH_LNO") & ""

        Dim appt As UltraWinSchedule.Appointment = Nothing 'New UltraWinSchedule.Appointment

        Dim DTE1 As Date = rowSPTCOOPG.Item("DATE_START")
        Dim DTE2 As Date = rowSPTCOOPG.Item("DATE_START")
        If rowSPTCOOPG.Item("DATE_END") & "" <> "" Then
            DTE2 = rowSPTCOOPG.Item("DATE_END")
        End If
        appt = eventMonthView.CalendarInfo.Appointments.Add(DTE1, DTE2, SUBJECT)
        appt.Tag = EVENT_key
        appt.Appearance.BackColor = APPR_STATUS_CODE_BackColors(APPR_STATUS_CODE)
        appt.Appearance.ForeColor = APPR_STATUS_CODE_ForeColors(APPR_STATUS_CODE)

        appt.Description = SUBJECT
        appt.AllDayEvent = True

        If SCHED_events.ContainsKey(EVENT_key) Then
            SCHED_events(EVENT_key) = appt
        Else
            SCHED_events.Add(EVENT_key, appt)
        End If
    End Sub

    Sub Edit_Appointment(ByVal EVENT_key As String)

        If EVENT_key.Length = 0 Then Exit Sub
        Dim AUTH_NO As String = Split(EVENT_key, ".")(0)
        Dim AUTH_LNO As Integer = Split(EVENT_key, ".")(1)
        Dim sql As String = "AUTH_NO = '" & AUTH_NO & "' and AUTH_LNO = " & CStr(AUTH_LNO)

        Using F As New SPFCOOP2
            Dim row As DataRow = dst.Tables("SPTCOOPG").Select(sql)(0)
            F.rowSPTCOOPX = row
            F.sqlSPTCOOPX = sqlSPTCOOPX
            F.CUST_CODE = F.rowSPTCOOPX.Item("CUST_CODE")
            F.VEHICLE_CODE = F.rowSPTCOOPX.Item("VEHICLE_CODE")
            F.SREP_CODE = F.rowSPTCOOPX.Item("SREP_CODE")
            F.BRAND_CODE = F.rowSPTCOOPX.Item("BRAND_CODE")
            F.ShowDialog()

            If F.UPDATED Then
                Remove_Appointment(EVENT_key)
                Add_Appointment(F.rowSPTCOOPX)
                row.ItemArray = F.rowSPTCOOPX.ItemArray
                dst.Tables("SPTCOOPG").AcceptChanges()
            End If
        End Using
    End Sub

    Sub Get_SP_Events()

        Update_Record_TDA("SPTCODE1")
        ASCMAIN1.sql = sqlSPTCOOPX
        For Each COLUMN_NAME As String In New String() {"CUST_CODE", "BRAND_CODE", "SREP_CODE", "COLLECTION_CODE"}
            Dim A As Integer = dst.Tables("SPTCODE1").Select("CODE_TYPE = '" & COLUMN_NAME & "'").Length
            Dim B As Integer = dst.Tables("SPTCODE1").Select("CODE_TYPE = '" & COLUMN_NAME & "' and SEL = '1'").Length
            If A <> B Then
                Dim TABLE_NAME As String = "SPTCOOP1"
                If COLUMN_NAME = "COLLECTION_CODE" Then
                    TABLE_NAME = "SPTCOOP3"
                ElseIf COLUMN_NAME = "BRAND_CODE" Then
                    TABLE_NAME = "ICTCOLL1"
                End If
                ASCMAIN1.sql &= " and " & TABLE_NAME & "." & COLUMN_NAME _
                    & " in (Select CODE_VALUE from " & SPTCODE1 & " where CODE_TYPE = '" & COLUMN_NAME & "' and SEL = '1')"
            End If

        Next

        If Not chkStartDate.Checked Then
            ASCMAIN1.sql &= " and SPTCOOP1.DATE_START >= '" & dteStartDate.DateTime.ToString("dd-MMM-yyyy") & "'"
            eventMonthView.CalendarInfo.MinDate = dteStartDate.DateTime.Date
        End If

        If Not chkEnddate.Checked Then
            ASCMAIN1.sql &= " and SPTCOOP1.DATE_START <= '" & dteEndDate.DateTime.ToString("dd-MMM-yyyy") & "'"
            eventMonthView.CalendarInfo.MaxDate = dteEndDate.DateTime.Date
        End If

        Fill_Records("SPTCOOPG", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdSPTCOOPG, "AUTH_NO".ToLower)

        If chkStartDate.Checked Then
            eventMonthView.CalendarInfo.MinDate = dst.Tables("SPTCOOPX").Compute("MIN(DATE_START)", "")
        End If

        If chkEnddate.Checked Then
            eventMonthView.CalendarInfo.MaxDate = dst.Tables("SPTCOOPX").Compute("MAX(DATE_START)", "")
        End If

        Load_Events_into_Calendar()

    End Sub

#End Region

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        With UltraExplorerBar1
            .Groups("Screen Control").Visible = ScreenMode Or (tab0.SelectedTab.Key = "Event && Expense Log")
            .Groups("Show").Visible = Not ScreenMode And (tab0.SelectedTab.Key = "Event && Expense Log")
            .Groups("Date Range").Visible = (tab0.SelectedTab.Key = "Calendar")
            .Groups("Status Legend").Visible = False ' (tab0.SelectedTab.Key = "Calendar")
            .Groups("Filters").Visible = (tab0.SelectedTab.Key = "Calendar")
        End With

        spl.Panel1Collapsed = Not (tab0.SelectedTab.Key = "Event && Expense Log")
    End Sub

    Private Sub optCODE_TYPE_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optCODE_TYPE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        SET_SPTCODE1()
    End Sub

    Sub SET_SPTCODE1()
        Dim dvw As DataView = DirectCast(grdSPTCODE1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "CODE_TYPE = '" & optCODE_TYPE.Value & "'"
        Sort_grdColumns(grdSPTCODE1, "CODE_VALUE")
        grdSPTCODE1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = optCODE_TYPE.Text
    End Sub

    Sub Manage_Document()

        Dim mail As New System.Net.Mail.MailMessage()
        Dim folder As String = ASCMAIN1.Folders("Archive") & "documents\"
        'folder = ASCMAIN1.Folders("Attach")
        If Not My.Computer.FileSystem.DirectoryExists(folder) Then
            My.Computer.FileSystem.CreateDirectory(folder)
        End If

        mail.Body = "Test Message Body"

        mail.From = New System.Net.Mail.MailAddress("wjz@absolution.com", "wjz")
        mail.Subject = "This is my Subject"
        mail.Sender = New System.Net.Mail.MailAddress("walter@absolution.com", "Walter")

        Dim EMAIL_LOGO As String = "INT.PNG"

        Dim SEND_BODY As String = "Hello" & vbCrLf & vbCrLf & "This is important"
        Dim SEND_FROM_SIGNATURE As String = "Walter J. Zielenski" & vbCrLf & "President"

        Dim plainView As System.Net.Mail.AlternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(SEND_BODY)
        Dim htmlView As System.Net.Mail.AlternateView
        If EMAIL_LOGO <> "" Then
            htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString("<img src=cid:logo>" & "<p>" & Replace(SEND_BODY & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE, vbCrLf, "<br>") & "</p>", Nothing, "text/html")
            Dim logo As New System.Net.Mail.LinkedResource(ASCMAIN1.Folders("Images") & "ABS\" & EMAIL_LOGO)
            logo.ContentId = "logo"
            htmlView.LinkedResources.Add(logo)
        Else
            htmlView = System.Net.Mail.AlternateView.CreateAlternateViewFromString("<p>" & SEND_BODY & vbCrLf & vbCrLf & SEND_FROM_SIGNATURE & "</p>", Nothing, "text/html")
        End If

        mail.AlternateViews.Add(plainView)
        mail.AlternateViews.Add(htmlView)

        Dim DOC_NO = ASCMAIN1.Next_Control_No("ASTDOCM1.DOC_NO")

        Dim EXT As String = ".eml"
        Try
            mail.Save(folder & DOC_NO & EXT)
        Catch ex As Exception

        End Try

        mail = Nothing

        '  DOC_NO = "0000000000"

        Dim appFileName = folder & DOC_NO & EXT ' ASCMAIN1.Folders("Work") & DOC_NO & "." & "eml"
        Dim p As Process = Process.Start(appFileName)

    End Sub

    'Dim outlookNameSpace As Microsoft.Office.Interop.Outlook.NameSpace
    'Dim inbox As Microsoft.Office.Interop.Outlook.MAPIFolder
    'Dim WithEvents items As Microsoft.Office.Interop.Outlook.Items

    'Private Sub ThisAddIn_Startup() Handles Me.Startup

    '    outlookNameSpace = Me.Application.GetNamespace("MAPI")
    '    inbox = _
    '        outlookNameSpace.GetDefaultFolder( _
    '        Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)
    '    items = inbox.Items

    'End Sub

    'Private Sub Items_ItemAdd(ByVal item As Object) Handles items.ItemAdd
    '    Dim filter As String = "USED CARS"
    '    If TypeOf (item) Is Microsoft.Office.Interop.Outlook.MailItem Then
    '        Dim mail As Microsoft.Office.Interop.Outlook.MailItem = item
    '        If mail.MessageClass = "IPM.Note" And _
    ' mail.Subject.ToUpper.Contains(filter.ToUpper) Then
    '            mail.Move(outlookNameSpace.GetDefaultFolder( _
    '                Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderJunk))
    '        End If
    '    End If

    'End Sub

    Function Check_Changed_Fields() As Boolean

        'TABLE_NAME,
        'KEY_VALUE,
        'COLUMN_NAME,
        'USER_ID,
        'INIT_DATE,
        'OLD_VALUE,
        'NEW_VALUE,
        'FM_MODE,
        'NOTES,
        'KEY_VALUE2,
        'KEY_LNO,
        'SESSION_NO,
        'SELECTION_NO,
        'XNO,

        Dim REV_NO As Integer = 0

        REV_NO += 1

        Dim LAST_DATE As Date = DATETIME_STAMP
        If EntryMode = "N" Then Stop
        Dim REV_LNO As Integer = 0

        Check_Changed_Fields = False

        dst.Tables("ASTAUDT1").Rows.Clear()

        Write_Audit_Trail(rowSPTCOOP1, EntryMode)

        For Each rowSPTCOOP3 As DataRow In dst.Tables("SPTCOOP3").Select("")
            Write_Audit_Trail(rowSPTCOOP3, EntryMode)
        Next

        ASCMAIN1.Progress("Logging Header Changes")


        'For i As Integer = 0 To rowSPTCOOP1.Table.Columns.Count - 1
        '    Dim COLUMN_NAME As String = dst.Tables("SPTCOOP1").Columns(i).ColumnName

        '    If rowSPTCOOP1.Item(COLUMN_NAME) & "" _
        '    <> rowSPTCOOP1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
        '        Check_Changed_Fields = True
        '        ASCMAIN1.Progress("-", COLUMN_NAME)
        '        Dim rowASTAUDT1 As DataRow = dst.Tables("ASTAUDT1").NewRow
        '        With rowASTAUDT1
        '            .Item("REV_NO") = REV_NO
        '            REV_LNO += 1
        '            .Item("REV_LNO") = REV_LNO
        '            .Item("KEY_VALUE") = AUTH_NO
        '            .Item("ORDR_LNO") = 0
        '            .Item("INIT_DATE") = LAST_DATE
        '            .Item("INIT_OPER") = ASCMAIN1.USER_ID
        '            .Item("COLUMN_NAME") = COLUMN_NAME
        '            .Item("OLD_VALUE") = rowSPTCOOP1.Item(COLUMN_NAME, DataRowVersion.Original)
        '            .Item("NEW_VALUE") = rowSPTCOOP1.Item(COLUMN_NAME)
        '            .Item("EMODE") = EntryMode
        '        End With
        '        dst.Tables("ASTAUDT1").Rows.Add(rowASTAUDT1)
        '        Check_Changed_Fields = True
        '    End If
        'Next i

        'ASCMAIN1.Progress("Logging Detail Changes")

        ASCMAIN1.Progress("")
        Return Check_Changed_Fields
    End Function

    Sub Clone_Record()
        BeginTrans()

        AUTH_NO_new = ASCMAIN1.Next_Control_No("SPTCOOP1.AUTH_NO")
        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SPTCOOP1", "SPTCOOP2", "SPTCOOP3", "SPTCOOP5", "SPTCOOP8", "SPTCOOP9", "SPTCOOPB"}
            dst.Tables(TABLE_NAME).AcceptChanges()

            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")

                row.Item("AUTH_NO") = AUTH_NO_new

                Select Case TABLE_NAME
                    Case "SPTCOOP1"
                        row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        row.Item("INIT_DATE") = DATETIME_STAMP
                        row.Item("LAST_OPER") = DBNull.Value
                        row.Item("LAST_DATE") = DBNull.Value
                        row.Item("APPR_STATUS_CODE") = "P"
                        row.Item("STATUS_CODE") = "O"
                        row.Item("AUTH_DATE") = DATETIME_STAMP.Date
                        row.Item("CUST_REF_NUM") = DBNull.Value
                        row.Item("CUST_AGR_RECD") = "0"
                        row.Item("PROOF_ADV_RECD") = "0"
                        row.Item("SAMPLE_RECD") = "0"
                        row.Item("PAID_AMT") = 0
                        Dim TOTAL_AMT As Decimal = Val(Absx1.numFor("TOTAL_AMT").Value & "")
                        row.Item("OPEN_AMT") = TOTAL_AMT

                        row.Item("PYMTS") = 0
                        row.Item("AUTH_APPR_DATE") = DBNull.Value
                        row.Item("AUTH_APPR_BY") = DBNull.Value
                        row.Item("AUTH_APPR_AMT") = DBNull.Value
                        row.Item("AUTH_APPR_NOTES") = DBNull.Value
                End Select

                row.AcceptChanges()
                row.SetAdded()
            Next
            Update_Record_TDA(TABLE_NAME)

            ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO_new, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CLONE", "Contract Cloned from " & AUTH_NO, "")
            ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CLONE", "Contract Cloned to " & AUTH_NO_new, "")

        Next
        EnforceConstraints(True)
        CommitTrans("Clone Successful - New Auth No:  " & AUTH_NO_new)
    End Sub

    Private Sub optShow_ValueChanged(sender As Object, e As EventArgs) Handles optShow.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_cbeYP()
    End Sub

    Sub Setup_cbeYP()
        cbeYP.Visible = (optShow.Value = "E")
        Refresh_Documents()
    End Sub

    Sub Setup_Retail_Weeks()
        If Absx1.txtFor("OPS_YYYYWW").Text = "" Then
            Absx1.txtFor("OPS_YYYYWW").Text = Set_OPS_YYYYWW()
        End If

        Dim OPS_YYYYWW As String = Absx1.txtFor("OPS_YYYYWW").Text

        If OPS_YYYYWW <> "" Then
            For P As Integer = -6 To 3
                Dim YW As String = ASCMAIN1.Week_Calc(OPS_YYYYWW, P)
                Dim row As DataRow = LookUp("GLTPARM3", YW)
                Dim C As String = ""
                If P <= 0 Then
                    C = "P" & Format(-1 * P, "00")
                Else
                    C = "N" & Format(P, "00")
                End If
                Dim LEGEND As String = row.Item("LEGEND") & ""
                With grdSPTCOOP9.DisplayLayout.Bands(0).Groups(C)
                    .Header.ToolTipText = LEGEND
                    .Header.Caption = "#" & Mid(LEGEND, 6, 2)
                End With
            Next
        End If

    End Sub

    Function Set_OPS_YYYYWW() As String
        If Absx1.dteFor("DATE_START").Value & "" = "" Then
            Return ""
        Else
            ASCMAIN1.sql = "Select MAX (YYYYWW) from GLTPARM3 where WEEK_END_DATE <= '" & Format(Absx1.dteFor("DATE_START").Value, "dd-MMM-yyyy") & "'"
            Dim OPS_YYYYWW As String = ASCDATA1.GetDataValue
            Return OPS_YYYYWW
        End If
    End Function

    Sub Seek_Approval()
        Dim TOTAL_AMT As Decimal = Val(Absx1.numFor("TOTAL_AMT").Value & "")
        Dim DATE_START As Date = Absx1.dteFor("DATE_START").Value
        Dim DATE_END As Date = Absx1.dteFor("DATE_END").Value

        Dim RETAIL_GOAL As Decimal = Val(dst.Tables("SPTCOOP9").Compute("SUM(RETAIL_GOAL)", "") & "")

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim VEHICLE_CODE As String = Absx1.txtFor("VEHICLE_CODE").Text

        AUTH_APPR_NOTES = ""

        Dim LBL As String = "Total Contract Amount is " & Format(TOTAL_AMT, "$#,##0.00") _
                                  & vbCrLf & "Event Vehicle to be used is " & VEHICLE_CODE _
                                  & vbCrLf & vbCrLf & "Booking Name is " & Absx1.txtFor("BOOKING_NAME").Text _
                                  & vbCrLf & vbCrLf & "Retail Goal is " & Format(RETAIL_GOAL, "#,##0") _
                                  & vbCrLf & "Customer is " & CUST_CODE _
                                  & vbCrLf & vbCrLf & "Event Date Range is " & Format(DATE_START, "MM/dd/yy") & " thru " & Format(DATE_END, "MM/dd/yy") _
                                  & vbCrLf & vbCrLf & "Enter Notes to Record with this Approval"

        AUTH_APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, "OK To Approve this Contract?", False, 60, "Approved")
    End Sub

    Sub Update_Approval()

        If Not update_with_approval Then BeginTrans()

        Dim TOTAL_AMT As Decimal = Val(Absx1.numFor("TOTAL_AMT").Value & "")
        Approve_Record(AUTH_NO, TOTAL_AMT, "A")

        If Not update_with_approval Then
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
            CommitTrans("Approval Complete")
        End If
    End Sub

    Sub Verify_Record(AUTH_NO As String, TOTAL_AMT As Decimal)

        ' the variable AUTH_APPR_NOTES really contains the verification notes at this point

        Dim VERIFIED_AS_OPEN_NOTES As String = AUTH_APPR_NOTES
        Dim rowSPTCOOP1 As DataRow = LookUp("SPTCOOP1", AUTH_NO)
        If chkKeepComment.Checked AndAlso rowSPTCOOP1.Item("VERIFIED_AS_OPEN_NOTES") & "" <> "" Then
            VERIFIED_AS_OPEN_NOTES = rowSPTCOOP1.Item("VERIFIED_AS_OPEN_NOTES") & ""
        End If

        ASCMAIN1.sql = "Update SPTCOOP1 Set VERIFIED_AS_OPEN_DATE = :PARM1, VERIFIED_AS_OPEN_BY = :PARM2, VERIFIED_AS_OPEN_AMT = :PARM3, VERIFIED_AS_OPEN_NOTES = :PARM4, VERIFIED_AS_OPEN = '1' where AUTH_NO = '" & AUTH_NO & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, VERIFIED_AS_OPEN_NOTES})

        For Each rowSPTCOOPX As DataRow In dst.Tables("SPTCOOPX").Select($"AUTH_NO = '{AUTH_NO}'")
            rowSPTCOOPX.Item("VERIFIED_AS_OPEN_DATE") = DATETIME_STAMP
            rowSPTCOOPX.Item("VERIFIED_AS_OPEN_BY") = ASCMAIN1.USER_ID
            rowSPTCOOPX.Item("VERIFIED_AS_OPEN_AMT") = TOTAL_AMT
            rowSPTCOOPX.Item("VERIFIED_AS_OPEN_NOTES") = VERIFIED_AS_OPEN_NOTES
            rowSPTCOOPX.Item("VERIFIED_AS_OPEN") = "1"
        Next

        ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "VER_O", "Verified for " & Format(TOTAL_AMT, "$#,##0.00") & "; " & VERIFIED_AS_OPEN_NOTES, "")
    End Sub

    Sub Close_Record(AUTH_NO As String, TOTAL_AMT As Decimal)
        'Stop
        ASCMAIN1.sql = "Update SPTCOOP1 Set LAST_DATE = :PARM1, LAST_OPER = :PARM2, OPEN_AMT = :PARM3, OTHER_COST = :PARM3, STATUS_CODE = :PARM4 where AUTH_NO = '" & AUTH_NO & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, 0, "C"})

        ASCMAIN1.sql = "Update SPTCOOP3 Set DIST_AMT = 0 where AUTH_NO = '" & AUTH_NO & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        For Each rowSPTCOOPX As DataRow In dst.Tables("SPTCOOPX").Select($"AUTH_NO = '{AUTH_NO}'")
            rowSPTCOOPX.Item("LAST_DATE") = DATETIME_STAMP
            rowSPTCOOPX.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSPTCOOPX.Item("OPEN_AMT") = 0
            rowSPTCOOPX.Item("DIST_AMT") = 0
            rowSPTCOOPX.Item("STATUS_CODE") = "C"
            ' rowSPTCOOPX.Delete()
        Next

        ' the variable AUTH_APPR_NOTES really contains the zeroing notes at this point
        ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "ZERO", "Closed for " & Format(TOTAL_AMT, "$#,##0.00") & "; " & AUTH_APPR_NOTES, "")
    End Sub

    Sub Approve_Record(AUTH_NO As String, TOTAL_AMT As Decimal, APPROVAL_STATUS As String, Optional APPROVAL_STATUS_orig As String = "")
        Select Case APPROVAL_STATUS
            Case "A"
                ASCMAIN1.sql = "Update SPTCOOP1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = 'A' where AUTH_NO = '" & AUTH_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})

                Dim rowSPTCOOP1 As DataRow = LookUp("SPTCOOP1", AUTH_NO)
                Dim EVENT_GROUP_NO As String = rowSPTCOOP1.Item("EVENT_GROUP_NO") & ""
                If EVENT_GROUP_NO <> "" Then
                    ASCMAIN1.sql = "Update SPTSFOC1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = 'A' where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})
                End If

                ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPR", "Approved for " & Format(TOTAL_AMT, "$#,##0.00") & "; " & AUTH_APPR_NOTES, "")

            Case Else

                ASCMAIN1.sql = $"Update SPTCOOP1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = '{APPROVAL_STATUS}' where AUTH_NO = '" & AUTH_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})

                Dim rowSPTCOOP1 As DataRow = LookUp("SPTCOOP1", AUTH_NO)

                Dim EVENT_GROUP_NO As String = rowSPTCOOP1.Item("EVENT_GROUP_NO") & ""
                If EVENT_GROUP_NO <> "" Then
                    ASCMAIN1.sql = $"Update SPTSFOC1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = '{APPROVAL_STATUS}' where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})
                End If

                ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPR", $"Approval Status Changed from {APPROVAL_STATUS_orig} to {APPROVAL_STATUS} for " & Format(TOTAL_AMT, "$#,##0.00") & "; " & AUTH_APPR_NOTES, "")

        End Select
    End Sub

    Private Sub grdSPTCOOP9_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSPTCOOP9.InitializeRow
        e.Row.Cells("TY").Value = "TY"
        e.Row.Cells("LY").Value = "LY"
    End Sub

    Sub Edit_Task(ByVal EVENT_key As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Editing Task")

        If EVENT_key.Length = 0 Then Exit Sub
        Dim AUTH_NO As String = Split(EVENT_key, ":")(0)
        Dim AUTH_TNO As Integer = Val(Split(EVENT_key, ":")(1))
        Dim rowSPTCOOP6 As DataRow = dst.Tables("SPTCOOP6").Rows.Find(New Object() {AUTH_NO, AUTH_TNO})
        ' Dim TASK_ID As String = rowSPTCOOP6.Item("TASK_ID")

        Using F As New SPFCOOPT

            F.frmASFBASE0 = Me
            F.rowSPTCOOP6 = rowSPTCOOP6
            F.CUST_CODE = Absx1.txtFor("CUST_CODE").Text
            F.VEHICLE_CODE = Absx1.txtFor("VEHICLE_CODE").Text
            'F.TASK_ASSIGNED_TO = F.rowPOTWPDM6.Item("TASK_ASSIGNED_TO")
            F.AUTH_NO = Absx1.txtFor("AUTH_NO").Text

            F.CUST_NAME = Absx1.txtFor("CUST_NAME").Text
            F.VEHICLE_DESC = Absx1.txtFor("VEHICLE_DESC").Text
            F.BOOKING_NAME = Absx1.txtFor("BOOKING_NAME").Text

            F.ShowDialog()

            If F.UPDATED Then
                'Dim rowPOTWPDM5 As DataRow = dst.Tables("POTWPDM5").Rows.Find(New Object() {STYLE_GROUP_NO, STEP_LNO})
                'rowPOTWPDM5.Item("STEP_STATUS") = Set_STEP_STATUS(STEP_LNO)
            End If
        End Using

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Set_STEP_STATUS(STEP_LNO As Int32) As String
        Dim STEP_STATUS As String = ""
        For Each rowPOTWPDM6_status As DataRow In dst.Tables("POTWPDM6").Select("STEP_LNO = " & CStr(STEP_LNO))
            If rowPOTWPDM6_status.Item("TASK_STATUS") & "" = "U" Then
                STEP_STATUS = "U"
                Exit For
            ElseIf rowPOTWPDM6_status.Item("TASK_STATUS") & "" = "O" Then
                STEP_STATUS = "O"
            ElseIf rowPOTWPDM6_status.Item("TASK_STATUS") & "" = "C" And STEP_STATUS = "" Then
                STEP_STATUS = "C"
            End If
        Next

        Return STEP_STATUS

    End Function

    Private Sub UltraButton2_Click(sender As Object, e As EventArgs) Handles UltraButton2.Click
        Dim workbook As SpreadsheetGear.IWorkbook = Nothing
        Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing

        If Not ASCMAIN1.Running_in_VS Then Exit Sub

        Dim FILENAME As String = "" '"C:\Users\wjz\Desktop\Interparfums\ASP\FieldCopy_CFG-IPLB_PromoExpenseMaster 7.14.15.xlsm"
        'FILENAME = "C:\Users\wjz\Desktop\Interparfums\ASP\FieldCopy_CFG-IPLB_PromoExpenseMaster 8.4.15.xlsm"
        'FILENAME = "C:\Users\wjz\Desktop\Interparfums\ASP\FieldCopy_CFG-IPLB_PromoExpenseMaster 8.27.15.xlsm"
        'FILENAME = "C:\Users\wjz\Desktop\Interparfums\ASP\Copy of IPLB PromoExpenseMaster.xlsm"
        'FILENAME = "C:\Users\wjz\Desktop\Interparfums\ASP\IPLB PromoExpenseMaster10.6.15.xlsm"
        'FILENAME = "C:\Users\wjz\Desktop\Interparfums\ASP\IPLB PromoExpenseMaster.xlsm"
        'FILENAME = "C:\Users\wjz\Desktop\Interparfums\ASP\Copy of IPLB PromoExpenseMaster.xlsm"
        FILENAME = "C:\Users\wjz\Desktop\Interparfums\ASP\IPLB PromoExpenseMaster.xlsm"
        workbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
        worksheet = workbook.Sheets("MASTER_LOG")

        ASCMAIN1.sql = "Delete from COOP"
        ASCDATA1.ExecuteSQL()

        If dst.Tables.Contains("COOP") Then
            dst.Tables("COOP").Rows.Clear()
        Else
            Create_TDA(dst.Tables.Add, "COOP", "*")
        End If

        ASCMAIN1.Progress("Now Converting")

        Dim R As Int64 = 4
        R += 2 ' AS OF 10/06 CONVERSION - 2 LINES WERE ADDED TO HEADER
        Do While worksheet.Cells(R + 1, 0).Value IsNot Nothing And worksheet.Cells(R + 1, 9).Value IsNot Nothing ' worksheet.Cells(R + 1, 1).Value IsNot Nothing
            ASCMAIN1.Progress("-", CStr(R))
            R += 1
            Dim CAL As String = worksheet.Cells(R, 1).Value & ""
            'If CAL.StartsWith("IPLB") Then
            Dim row As DataRow = dst.Tables("COOP").NewRow
            For C As Integer = 0 To dst.Tables("COOP").Columns.Count - 1
                Dim DTC As String = dst.Tables("COOP").Columns(C).DataType.ToString
                If DTC = "System.Decimal" Or DTC.StartsWith("System.Int") Then
                    row.Item(C) = Val(worksheet.Cells(R, C).Value & "")
                ElseIf DTC = "System.DateTime" Then
                    Dim DTT As String = worksheet.Cells(R, C).Text
                    If DTT <> "" Then
                        If Mid(DTT, 4, 1) = "," Then DTT = Mid(DTT, 6)
                        If DTT.Contains(",") Then DTT = Mid(DTT, 1, InStr(DTT, ",") - 1)
                        If DTT.Contains("/") And DTT <> "01/00/00" Then ' DTT <> "F" And DTT <> "Tue" And DTT <> "Sun" Then
                            row.Item(C) = DTT
                        End If
                    End If
                Else
                    row.Item(C) = worksheet.Cells(R, C).Value
                End If
            Next
            dst.Tables("COOP").Rows.Add(row)
            'End If
        Loop

        Update_Record_TDA("COOP")

        ASCMAIN1.sql = "update coop set retailer_vendor = (select CUST_CODE from tatxref2 where cscus1 = cus_1)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "update coop set retailer_vendor = 'SEPHORA' where cus_1 = '415002'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "update coop set retailer_vendor = 'IPLBMARKET' where cus_1 = '666666'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "update coop set br_num = '09' where br_num = '9'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "update coop set br_num = '67A' where br_num = '67'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "update coop set br_num = '27A' where br_num = '27'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "update coop set br_num = '71A' where br_num = '71'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "delete from coop where br_num = '77'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "update coop set brand = (select collection_code from tatxref1 where BLSUBBRND = br_num)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from coop where CLOSING_STATUS IS NULL"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from coop where ACCRUAL_DATE IS NULL"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "BEGIN DECLARE R NUMBER;" & vbCrLf _
            & "CURSOR C1 IS SELECT * FROM (SELECT * FROM COOP ORDER BY ENTRY_DATE) FOR UPDATE;" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "R := NVL(R,0) + 1;" & vbCrLf _
            & "UPDATE COOP SET GENDER = TRIM(TO_CHAR(R,'000000'))  WHERE CURRENT OF C1;" & vbCrLf _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        'ASCMAIN1.sql = "Delete from coop where ACCRUAL_DATE IS NULL"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "insert into INT_VEHMAP SELECT ADSIZE_VEHICLE, EXPENSE_TYPE, NULL FROM  (" _
            & " SELECT distinct ADSIZE_VEHICLE, EXPENSE_TYPE from COOP" _
            & " minus" _
            & " SELECT ADSIZE_VEHICLE, EXPENSE_TYPE from INT_VEHMAP)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "truncate table SPTCOOP1"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "truncate table SPTCOOP3"
        ASCDATA1.ExecuteSQL()

        '            & "DECODE(NVL(CIRC_QTY,0),0,0,DECODE(NVL(UNITCOST_SCENTEDPC_,0),0,NVL(EXPECTED_CHARGE_,0)/NVL(CIRC_QTY,0),NVL(UNITCOST_SCENTEDPC_,0)/1000) * 1000) VEHICLE_CPM," & vbCrLf _

        '            & "DECODE(CLOSING_STATUS,'ACCRUE','A','PREPAID','P','OPEN','G','?') APPR_STATUS_CODE," & vbCrLf _


        ASCMAIN1.sql = "" _
            & "INSERT INTO SPTCOOP1" & vbCrLf _
            & "SELECT GENDER AUTH_NO," & vbCrLf _
            & "RETAILER_VENDOR CUST_CODE," & vbCrLf _
            & "ENTRY_DATE AUTH_DATE," & vbCrLf _
            & "BRAND AUTH_REQ_BY," & vbCrLf _
            & "CASE WHEN KAD_KAM LIKE 'Robert%' then '531' ELSE" & vbCrLf _
            & "CASE WHEN KAD_KAM LIKE 'Jennifer%' then '533' ELSE" & vbCrLf _
            & "CASE WHEN KAD_KAM LIKE 'Yvonne%' then '539' ELSE '591' END END END SREP_CODE," & vbCrLf _
            & "'conv' INIT_OPER," & vbCrLf _
            & "'conv' LAST_OPER," & vbCrLf _
            & "SYSDATE INIT_DATE," & vbCrLf _
            & "SYSDATE LAST_DATE," & vbCrLf _
            & "REFNUM CUST_REF_NUM," & vbCrLf _
            & "NULL SELL_CODE," & vbCrLf _
            & "DECODE(TRIM(UPPER(C.STATUS)),'CANCELLED','X','APPROVED','A','PENDING','P','G') APPR_STATUS_CODE," & vbCrLf _
            & "BOOKNAME BOOKING_NAME," & vbCrLf _
            & "DECODE(AGREEMENT__RECEIVED,'YES','1','0') CUST_AGR_RECD," & vbCrLf _
            & "DECODE(PAA_RECEIVED,'YES','1','0') PROOF_ADV_RECD," & vbCrLf _
            & "DECODE(UPPER(EXAMPLE_RECVD),'YES','1','0') SAMPLE_RECD," & vbCrLf _
            & "CASE WHEN MONTHX < 7 THEN YEAR || 'S' ELSE YEAR || 'F' END SEASON_CODE," & vbCrLf _
            & "NULL SALES_DIVISION_CODE," & vbCrLf _
            & "DECODE(C.EXPENSE_TYPE," _
            & "'Coop','COOP'," _
            & "'IPLBEducation','TRAINMATER'," _
            & "'COOP','COOP'," _
            & "'Spec Mrktg','MAILERS'," _
            & "'IP Sales Incen','SLSINCENT'," _
            & "'Nat Media','NATMEDIA',     " _
            & "'RTLEVENTS','RTLEVENTS'," _
            & "'Scent','SCENT'," _
            & "'Visual','VISUAL'," _
            & "'scent','SCENT',nvl(C.EXPENSE_TYPE,'?')) EXPENSE_TYPE_CODE," & vbCrLf _
            & "NULL EVENT_TYPE_CODE," & vbCrLf _
            & "TRIM(TO_CHAR(ACCRUAL_DATE,'YYYYMM')) OPS_YYYYPP," & vbCrLf _
            & "NULL EVENT_QUALIFIER," & vbCrLf _
            & "UPPER(NVL(C.VEHICLE,I.VEHICLE_CODE)) VEHICLE_CODE," & vbCrLf _
            & "DECODE(NVL(CIRC_QTY,0),0,0,NVL(EXPECTED_CHARGE_,0)/NVL(CIRC_QTY,0) * 1000) VEHICLE_CPM," & vbCrLf _
            & "START_DATE DATE_START," & vbCrLf _
            & "END_DATE DATE_END," & vbCrLf _
            & "CIRC_QTY QTY," & vbCrLf _
            & "DECODE(NVL(CIRC_QTY,0),0,EXPECTED_CHARGE_,0) OTHER_COST," & vbCrLf _
            & "COMMENTSX NOTES," & vbCrLf _
            & "0 OPEN_AMT," & vbCrLf _
            & "0 PAID_AMT," & vbCrLf _
            & "0 PYMTS," & vbCrLf _
            & "DECODE(TRIM(UPPER(C.STATUS)),'CANCELLED','X','O') STATUS_CODE," & vbCrLf _
            & "NULL OPS_YYYYWW," & vbCrLf _
            & "NULL AUTH_APPR_DATE," & vbCrLf _
            & "NULL AUTH_APPR_BY," & vbCrLf _
            & "NULL AUTH_APPR_AMT," & vbCrLf _
            & "NULL AUTH_APPR_NOTES," & vbCrLf _
            & "CASE WHEN UPPER(C.ADSIZE_VEHICLE) <> 'SCENTED PAGE' AND (UPPER(C.ADSIZE_VEHICLE) LIKE '%PAGE%' OR UPPER(C.ADSIZE_VEHICLE) = 'OMNI') THEN SUBSTR(C.ADSIZE_VEHICLE,1,20) ELSE NULL END AD_SIZE" & vbCrLf _
            & " from COOP C, INT_VEHMAP I" & vbCrLf _
            & " where NVL(I.ADSIZE_VEHICLE,'?') = NVL(C.ADSIZE_VEHICLE,'?')" & vbCrLf _
            & "   and NVL(I.EXPENSE_TYPE,'?') = NVL(C.EXPENSE_TYPE,'?')"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "INSERT INTO SPTCOOP3 SELECT S.AUTH_NO, 1, I.ITEM_CODE, S.AUTH_REQ_BY , NVL(S.OTHER_COST,0) + NVL(S.VEHICLE_CPM,0) * NVL(S.QTY,0) / 1000, C.FEATURE FROM SPTCOOP1 S,COOP C, ICTITEM1 I where C.GENDER = S.AUTH_NO and I.ITEM_CODE (+) =  C.ITEM_CODE"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET OPS_YYYYWW = (SELECT MIN (YYYYWW) FROM GLTPARM3 WHERE WEEK_END_DATE >= SPTCOOP1.DATE_START) WHERE DATE_START IS NOT NULL"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "truncate table SPTPYMT1"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "truncate table SPTPYMT2"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "truncate table SPTPYMT3"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" & vbCrLf _
            & "begin declare P NUMBER; cursor c1 is" & vbCrLf _
            & "SELECT GENDER AUTH_NO,CLAIM_1,CLAIM_OR_INVOICE_1,BATCH_DATE1,BATCH_1, " & vbCrLf _
            & "AUTH_REQ_BY, BRAND_CODE, TRADE_CLASS_CODE, SPTCOOP1.CUST_CODE FROM (" & vbCrLf _
            & "SELECT GENDER,CLAIM_1,CLAIM_OR_INVOICE_1,BATCH_DATE1,BATCH_1 FROM COOP WHERE NVL(CLAIM_1,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_2,CLAIM_OR_INVOICE_2,BATCH_DATE2,BATCH_2 FROM COOP WHERE NVL(CLAIM_2,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_3,CLAIM_OR_INVOICE_3,BATCH_DATE3,BATCH_3 FROM COOP WHERE NVL(CLAIM_3,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_4,CLAIM_OR_INVOICE_4,BATCH_DATE4,BATCH_4 FROM COOP WHERE NVL(CLAIM_4,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_5,CLAIM_OR_INVOICE_5,BATCH_DATE5,BATCH_5 FROM COOP WHERE NVL(CLAIM_5,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_6,CLAIM_OR_INVOICE_6,BATCH_DATE6,BATCH_6 FROM COOP WHERE NVL(CLAIM_6,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_7,CLAIM_OR_INVOICE_7,BATCH_DATE7,BATCH_7 FROM COOP WHERE NVL(CLAIM_7,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_8,CLAIM_OR_INVOICE_8,BATCH_DATE8,BATCH_8 FROM COOP WHERE NVL(CLAIM_8,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_9,CLAIM_OR_INVOICE_9,BATCH_DATE8,BATCH_9 FROM COOP WHERE NVL(CLAIM_9,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_10,CLAIM_OR_INVOICE_10,BATCH_DATE8,BATCH_10 FROM COOP WHERE NVL(CLAIM_10,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_11,CLAIM_OR_INVOICE_11,BATCH_DATE8,BATCH_11 FROM COOP WHERE NVL(CLAIM_11,0) <> 0 UNION" & vbCrLf _
            & "SELECT GENDER,CLAIM_12,CLAIM_OR_INVOICE_12,BATCH_DATE8,BATCH_12 FROM COOP WHERE NVL(CLAIM_12,0) <> 0) X," & vbCrLf _
            & "SPTCOOP1,ARTCUST1,ICTCOLL1 where SPTCOOP1.AUTH_NO = X.GENDER " & vbCrLf _
            & "AND ARTCUST1.CUST_CODE (+) = SPTCOOP1.CUST_CODE " & vbCrLf _
            & "AND ICTCOLL1.COLLECTION_CODE (+) = SPTCOOP1.AUTH_REQ_BY;" & vbCrLf _
            & "BEGIN for r1 in c1 loop" & vbCrLf _
            & "P := NVL(P,0) + 1;" & vbCrLf _
            & "insert into SPTPYMT3 VALUES " & vbCrLf _
            & "(TRIM(TO_CHAR(P,'0000000000')), 1, R1.AUTH_NO, 1, 1, NULL, " & vbCrLf _
            & "R1.AUTH_REQ_BY, R1.CLAIM_1, R1.CLAIM_1, R1.BRAND_CODE, R1.TRADE_CLASS_CODE);" & vbCrLf _
            & "insert into SPTPYMT2 VALUES " & vbCrLf _
            & "(TRIM(TO_CHAR(P,'0000000000')), 1, R1.AUTH_NO, 1, R1.CLAIM_1, '0', R1.CLAIM_OR_INVOICE_1);" & vbCrLf _
            & "insert into SPTPYMT1 VALUES " & vbCrLf _
            & "(TRIM(TO_CHAR(P,'0000000000')), 'R', R1.CLAIM_OR_INVOICE_1, R1.BATCH_DATE1,R1.CLAIM_1,TO_CHAR(R1.BATCH_DATE1,'YYYYMM')," & vbCrLf _
            & "R1.CUST_CODE, NULL, '0000000000', '1','000000', 'conv',sysdate,'conv',sysdate,null,null," & vbCrLf _
            & "'1','0000000000',null,r1.batch_1);" & vbCrLf _
            & "end loop; end; end;"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET SALES_DIVISION_CODE = 'IP1', AUTH_REQ_BY = 'Marketing'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET SALES_DIVISION_CODE = 'IP2' where AUTH_NO in (Select Distinct AUTH_NO from SPTCOOP3,ICTCOLL1,ICTBRAN1 where ICTCOLL1.COLLECTION_CODE = SPTCOOP3.COLLECTION_CODE and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE and ICTBRAN1.SALES_DIVISION_CODE = 'IP2')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET SALES_DIVISION_CODE = 'IP3' where AUTH_NO in (Select Distinct AUTH_NO from SPTCOOP3,ICTCOLL1,ICTBRAN1 where ICTCOLL1.COLLECTION_CODE = SPTCOOP3.COLLECTION_CODE and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE and ICTBRAN1.SALES_DIVISION_CODE = 'IP3')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET OPEN_AMT =  NVL(OTHER_COST,0) + NVL(VEHICLE_CPM,0) * NVL(QTY,0) / 1000, PAID_AMT = 0"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET PAID_AMT =  (SELECT SUM (PYMT_REF_AMT) FROM SPTPYMT2 WHERE AUTH_NO = SPTCOOP1.AUTH_NO)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET OPEN_AMT = NVL(OPEN_AMT,0) - NVL(PAID_AMT,0)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET OPEN_AMT = 0 WHERE OPEN_AMT < 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "BEGIN DECLARE L NUMBER; CURSOR C1 IS" & vbCrLf _
            & "SELECT BOOKING_NAME, AUTH_DATE, CUST_CODE, VEHICLE_CODE, DATE_START, APPR_STATUS_CODE, QTY" & vbCrLf _
            & ", COUNT (*) AUTHS, MIN (AUTH_NO) A1, MAX (AUTH_NO) A2, MIN (VEHICLE_CPM) CPM1, MAX (VEHICLE_CPM) FROM SPTCOOP1" & vbCrLf _
            & "GROUP BY BOOKING_NAME, AUTH_DATE, CUST_CODE, VEHICLE_CODE, DATE_START, APPR_STATUS_CODE, QTY HAVING COUNT (*) > 1;" & vbCrLf _
            & "BEGIN FOR R1 IN C1 LOOP" & vbCrLf _
            & "BEGIN DECLARE CURSOR C2 IS " & vbCrLf _
            & "SELECT * FROM SPTCOOP1 WHERE CUST_CODE = R1.CUST_CODE AND AUTH_DATE = R1.AUTH_DATE" & vbCrLf _
            & " AND NVL(DATE_START,TRUNC(SYSDATE)) = NVL(R1.DATE_START,TRUNC(SYSDATE)) AND NVL(APPR_STATUS_CODE,'?') = NVL(R1.APPR_STATUS_CODE,'?') AND NVL(QTY,0) = NVL(R1.QTY,0)" & vbCrLf _
            & " AND VEHICLE_CODE = R1.VEHICLE_CODE AND BOOKING_NAME = R1.BOOKING_NAME AND AUTH_NO <> R1.A1;" & vbCrLf _
            & "BEGIN " & vbCrLf _
            & "L := 1; " & vbCrLf _
            & "FOR R2 IN C2 LOOP" & vbCrLf _
            & "L := L + 1;" & vbCrLf _
            & "UPDATE SPTCOOP1 SET OPEN_AMT = NVL(OPEN_AMT,0) + NVL(R2.OPEN_AMT,0),PAID_AMT = NVL(PAID_AMT,0) + NVL(R2.PAID_AMT,0)" & vbCrLf _
            & ", CUST_AGR_RECD = CASE WHEN CUST_AGR_RECD = '1' OR R2.CUST_AGR_RECD = '1' THEN '1' ELSE '0' END" & vbCrLf _
            & ", PROOF_ADV_RECD = CASE WHEN PROOF_ADV_RECD = '1' OR R2.PROOF_ADV_RECD = '1' THEN '1' ELSE '0' END" & vbCrLf _
            & ", SAMPLE_RECD = CASE WHEN SAMPLE_RECD = '1' OR R2.SAMPLE_RECD = '1' THEN '1' ELSE '0' END" & vbCrLf _
            & ", OTHER_COST = NVL(OTHER_COST,0) + NVL(R2.OTHER_COST,0), VEHICLE_CPM = NVL(VEHICLE_CPM,0) + NVL(R2.VEHICLE_CPM,0)" & vbCrLf _
            & "WHERE AUTH_NO = R1.A1;" & vbCrLf _
            & "DELETE FROM SPTCOOP1 WHERE AUTH_NO = R2.AUTH_NO;" & vbCrLf _
            & "UPDATE SPTCOOP3 SET AUTH_LNO = L WHERE AUTH_NO = R2.AUTH_NO;" & vbCrLf _
            & "UPDATE SPTPYMT3 SET AUTH_LNO = L WHERE AUTH_NO = R2.AUTH_NO;" & vbCrLf _
            & "UPDATE SPTCOOP3 SET AUTH_NO = R1.A1 WHERE AUTH_NO = R2.AUTH_NO;" & vbCrLf _
            & "UPDATE SPTPYMT3 SET AUTH_NO = R1.A1 WHERE AUTH_NO = R2.AUTH_NO;" & vbCrLf _
            & "UPDATE SPTPYMT2 SET AUTH_NO = R1.A1 WHERE AUTH_NO = R2.AUTH_NO;" & vbCrLf _
            & "END LOOP; END; END;" & vbCrLf _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET EXPENSE_TYPE_CODE = (SELECT EXPENSE_TYPE_CODE FROM SPTAVEH1 WHERE VEHICLE_CODE = SPTCOOP1.VEHICLE_CODE)"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET STATUS_CODE = 'C' WHERE STATUS_CODE = 'O' AND OPEN_AMT = 0"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET STATUS_CODE = 'C', OPEN_AMT = 0 WHERE OPEN_AMT < 1 AND STATUS_CODE = 'O'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET STATUS_CODE = 'C', OPEN_AMT = 0 WHERE APPR_STATUS_CODE = 'X'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE SPTCOOP1 SET STATUS_CODE = 'C', OPEN_AMT = 0 WHERE DATE_START < '01-JAN-2013'" ' ALBINA SAYS THAT THIS DATE MAY GO TO 01/01/2014 BY THE TIME WE GO LIVE
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "BEGIN DECLARE CURSOR C1 IS SELECT * FROM SPTPYMT2 ORDER BY PYMT_NO;" _
            & " BEGIN " _
            & " UPDATE SPTCOOP1 SET PYMTS = 0;" _
            & " FOR R1 IN C1 LOOP" _
            & " UPDATE SPTCOOP1 SET PYMTS = PYMTS + 1 WHERE AUTH_NO = R1.AUTH_NO;" _
            & " UPDATE SPTPYMT2 SET AUTH_PNO = (SELECT PYMTS FROM SPTCOOP1 WHERE AUTH_NO = R1.AUTH_NO)" _
            & " WHERE PYMT_NO = R1.PYMT_NO AND PYMT_LNO = R1.PYMT_LNO;" _
            & "END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Delete from SPTCOOP3 WHERE AUTH_NO IN (" & vbCrLf _
            & "SELECT AUTH_NO FROM SPTCOOP1 WHERE CUST_CODE IS NULL AND BOOKING_NAME IS NULL " & vbCrLf _
            & "AND NVL(OPEN_AMT,0) = 0 AND NVL(PAID_AMT,0) = 0)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Delete from SPTCOOP1 WHERE AUTH_NO IN (" & vbCrLf _
            & "SELECT AUTH_NO FROM SPTCOOP1 WHERE CUST_CODE IS NULL AND BOOKING_NAME IS NULL " & vbCrLf _
            & "AND NVL(OPEN_AMT,0) = 0 AND NVL(PAID_AMT,0) = 0)"
        ASCDATA1.ExecuteSQL()

        MsgBox("1) Fix AUTH_NOs (Sarah has used 5000s for Spring 2016) (TATCTLN1?) (did 5000s get used and not ROWNUM)" _
               & vbCrLf & "2) Check that all Customers and Expense Types and Vehicle Codes are valid" _
               & vbCrLf & "3) Sarah Loaded ITEM_CODE and VEHICLE in COOP table - check them" _
               & vbCrLf & "   - bad ITEM_CODEs were stripped out, " _
               & vbCrLf & "   - bad VEHICLEs were not stripped out",
               MsgBoxStyle.OkOnly, "Important Post Mortem")

        ASCMAIN1.sql = "UPDATE TATCTLN1 SET CTL_NO_LAST = (SELECT TO_NUMBER(MAX(AUTH_NO)) FROM SPTCOOP1) WHERE CTL_NO_TYPE = 'SPTCOOP1.AUTH_NO'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "UPDATE TATCTLN1 SET CTL_NO_LAST = (SELECT TO_NUMBER(MAX(PYMT_NO)) FROM SPTPYMT1) WHERE CTL_NO_TYPE = 'SPTPYMT1.PYMT_NO'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.Progress("")
        MsgBox("Done")
    End Sub

    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowse.Click
        Dim FILENAME As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Link"
            openFileDialog1.Filter = "xlsx files (*.xlsx)|*.xlsx"
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                FILENAME = openFileDialog1.FileName
            End If
        End Using

        If FILENAME <> "" Then
            Absx1.txtFor("EVENT_FILE_LINK").Text = FILENAME
        End If
    End Sub

    Private Sub dteDATE_END_GotFocus(sender As Object, e As EventArgs) Handles dteDATE_END.GotFocus
        If ScreenMode And (EntryMode = "N") And dteDATE_END.Value & "" = "" Then
            dteDATE_END.Value = dteDATE_START.Value
        End If
    End Sub

    Private Sub chkEditNotes_CheckedChanged(sender As Object, e As EventArgs) Handles chkEditNotes.CheckedChanged

        With grdSPTCOOPX.DisplayLayout
            If chkEditNotes.Checked Then
                .Override.AllowUpdate = DefaultableBoolean.True
                .Bands(0).Columns("NOTES").CellAppearance.BackColor = Drawing.Color.Yellow
                .Bands(0).Columns("BOOKING_NAME").CellAppearance.BackColor = Drawing.Color.Yellow
                .Bands(0).Columns("NOTES").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Bands(0).Columns("BOOKING_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                chkEditVerComment.Enabled = False
                chkVerify.Enabled = False
            Else
                .Override.AllowUpdate = DefaultableBoolean.False
                .Bands(0).Columns("NOTES").CellAppearance.BackColor = Drawing.Color.Empty
                .Bands(0).Columns("BOOKING_NAME").CellAppearance.BackColor = Drawing.Color.Empty
                .Bands(0).Columns("NOTES").CellActivation = UltraWinGrid.Activation.NoEdit
                .Bands(0).Columns("BOOKING_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
                chkEditVerComment.Enabled = True
                chkVerify.Enabled = (optShow.Value = "O")
            End If
        End With
    End Sub

    Function Update_Record_SPTCOOP1() As Boolean

        Dim success As Boolean = False

        Try
            Dim AUTH_NO As String = grdSPTCOOPX.ActiveRow.Cells("AUTH_NO").Value
            'Dim rowSPTCOOPX As DataRow = dst.Tables("SPTCOOPX").Rows.Find(AUTH_NO)
            If Not ASCMAIN1.Logical_Lock("SPTCOOP1", AUTH_NO) Then

                Return False
            Else
                dst.Tables("SPTCOOP1").Rows.Clear()
                Dim rowSPTCOOP1 As DataRow = Fill_Record("SPTCOOP1", AUTH_NO)

                rowSPTCOOP1.Item("NOTES") = grdSPTCOOPX.ActiveRow.Cells("NOTES").Value
                rowSPTCOOP1.Item("BOOKING_NAME") = grdSPTCOOPX.ActiveRow.Cells("BOOKING_NAME").Value
                rowSPTCOOP1.Item("VERIFIED_AS_OPEN_COMMENTS") = grdSPTCOOPX.ActiveRow.Cells("VERIFIED_AS_OPEN_COMMENTS").Value
                If VERIFIED_AS_OPEN_COMMENTS_last_do_not_save Then
                Else
                    VERIFIED_AS_OPEN_COMMENTS_last = rowSPTCOOP1.Item("VERIFIED_AS_OPEN_COMMENTS") & ""
                End If


                For Each row As DataRow In dst.Tables("SPTCOOPX").Select("AUTH_NO = '" & AUTH_NO & "'")
                    row.Item("NOTES") = rowSPTCOOP1.Item("NOTES")
                    row.Item("BOOKING_NAME") = rowSPTCOOP1.Item("BOOKING_NAME")
                    row.Item("VERIFIED_AS_OPEN_COMMENTS") = rowSPTCOOP1.Item("VERIFIED_AS_OPEN_COMMENTS")
                Next

                DATETIME_STAMP = Now + ASCMAIN1.NowTSD

                BeginTrans()
                Write_Audit_Trail(rowSPTCOOP1, "E")
                Update_Record_TDA("SPTCOOP1")
                CommitTrans()

                ASCMAIN1.MultiTask_Release()



                success = True
            End If
        Catch ex As Exception
            MsgBox("Error Occurred: " & ex.Message, MsgBoxStyle.OkOnly, "Cannot Update Row for Record " & AUTH_NO)
        End Try

        dst.Tables("SPTCOOP1").Rows.Clear()
        Return success
    End Function

    Private Sub chkEditVerComment_CheckedChanged(sender As Object, e As EventArgs) Handles chkEditVerComment.CheckedChanged

        With grdSPTCOOPX.DisplayLayout
            If chkEditVerComment.Checked Then
                .Override.AllowUpdate = DefaultableBoolean.True
                .Bands(0).Columns("VERIFIED_AS_OPEN_COMMENTS").CellAppearance.BackColor = Drawing.Color.Yellow
                .Bands(0).Columns("VERIFIED_AS_OPEN_COMMENTS").CellActivation = UltraWinGrid.Activation.AllowEdit
                chkEditNotes.Enabled = False
                chkVerify.Enabled = False
                VERIFIED_AS_OPEN_COMMENTS_last = ""
            Else
                .Override.AllowUpdate = DefaultableBoolean.False
                .Bands(0).Columns("VERIFIED_AS_OPEN_COMMENTS").CellAppearance.BackColor = Drawing.Color.Empty
                .Bands(0).Columns("VERIFIED_AS_OPEN_COMMENTS").CellActivation = UltraWinGrid.Activation.NoEdit
                chkEditNotes.Enabled = True
                chkVerify.Enabled = (optShow.Value = "O")
            End If
        End With
    End Sub

    Private Sub chkVerify_CheckedChanged(sender As Object, e As EventArgs) Handles chkVerify.CheckedChanged

        Dim sqlsecs = $"ISNULL(DIST_OPEN,0) <> 0 AND EXPENSE_TYPE_CODE in ('{Join(EXPENSE_TYPE_CODEs_I_may_verify.ToArray, "','")}')"
        If chkHideVerified.Checked Then
            sqlsecs &= " and ISNULL(VERIFIED_AS_OPEN,'0') <> '1'"
        End If

        Dim DVW As DataView = DirectCast(grdSPTCOOPX.DataSource, DataTable).DefaultView

        chkKeepComment.Visible = chkVerify.Checked
        chkHideVerified.Visible = chkVerify.Checked

        With grdSPTCOOPX.DisplayLayout
            If chkVerify.Checked Then

                DVW.RowFilter &= sqlsecs

                chkEditNotes.Enabled = False
                chkEditVerComment.Enabled = False

            Else

                DVW.RowFilter = Replace(DVW.RowFilter, sqlsecs, "")

                chkEditNotes.Enabled = True
                chkEditVerComment.Enabled = True
            End If
        End With
    End Sub

    Overrides Sub Attachments_Changed()
        Fill_Records("ASTATTA2", AUTH_NO)
    End Sub

    Private Sub chkHideVerified_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideVerified.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim sqlHide As String = " and ISNULL(VERIFIED_AS_OPEN,'0') <> '1'"

        Dim DVW As DataView = DirectCast(grdSPTCOOPX.DataSource, DataTable).DefaultView
        Dim sqlw As String = DVW.RowFilter
        If chkHideVerified.Checked Then
            sqlw &= sqlHide
        Else
            sqlw = Replace(sqlw, sqlHide, "")
        End If
        DVW.RowFilter = sqlw

    End Sub

    Private Sub grdASTATTA2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTATTA2.ClickCellButton
        Dim ATTACHMENT_NO As String = grdASTATTA2.ActiveRow.Cells("ATTACHMENT_NO").Text
        Dim ATTACHMENT_EXT As String = grdASTATTA2.ActiveRow.Cells("ATTACHMENT_EXT").Text.ToUpper
        Call ASCMAIN1.Launch_Attachment(ATTACHMENT_NO, ATTACHMENT_EXT)
    End Sub

    Private Sub cbeBOOKED_BY_KeyDown(sender As Object, e As KeyEventArgs) Handles cbeBOOKED_BY.KeyDown
        If e.KeyCode = Keys.Delete Then
            If EntryMode = "N" Or EntryMode = "E" Then
                cbeBOOKED_BY.Value = DBNull.Value
            End If
        End If
    End Sub

    Private Sub cbeBOOKED_BY_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cbeBOOKED_BY.KeyPress
        If e.KeyChar = Chr(Keys.Delete) Then
            e.Handled = True
        End If
        If EntryMode = "N" Or EntryMode = "E" Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub grdSPTCOOP3_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSPTCOOP3.InitializeRow
        If e.Row.IsDataRow Then
            'Dim DIST_AMT As Decimal = Val(e.Row.Cells("DIST_AMT").Value & "")
            'Dim DIST_PAID As Decimal = Val(e.Row.Cells("DIST_PAID").Value & "")
            'If DIST_AMT < DIST_PAID Then
            '    e.Row.Cells("DIST_AMT").Appearance.ForeColor = System.Drawing.Color.Red
            '    e.Row.Cells("DIST_AMT").ToolTipText = "Distribution Amount may not be less than Paid Amount"
            'Else
            '    e.Row.Cells("DIST_AMT").Appearance.ForeColor = System.Drawing.Color.Empty
            '    e.Row.Cells("DIST_AMT").ToolTipText = ""
            'End If
        End If

    End Sub
    Private Sub Upload_Template()
        Using openFileDialog As New OpenFileDialog
            openFileDialog.Title = "Select a Template File to Upload"
            openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|xls files (*.xls)|*.xls"
            openFileDialog.RestoreDirectory = True
            dst.Tables("SPTCOOP1").Rows.Clear()
            dst.Tables("SPTCOOP3").Rows.Clear()
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Dim filename As String = openFileDialog.FileName
                Process_Template(filename)
            End If
        End Using
    End Sub
    Private Function Confirm_Upload(ByVal filename As String, ByVal lineCount As Integer, ByVal AUTH_NO_COUNT As Integer, ByVal dollaramt As Decimal) As Boolean
        Dim authNoText As String = If(AUTH_NO_COUNT = 1, "authorization number", "authorization numbers")
        ' Format the dollar amount to include commas and exactly two decimal places
        Dim formattedDollarAmt As String = dollaramt.ToString("N2")

        ' Construct the message using the variables above to handle singular/plural forms and include the dollar amount
        Dim message As String = $"You are about to upload a spreadsheet with {lineCount} lines and {AUTH_NO_COUNT} {authNoText}, totalling ${formattedDollarAmt}. Are you sure you want to proceed?"
        Dim caption As String = "Confirm Upload"
        Dim result As DialogResult = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        Return result = DialogResult.Yes
    End Function
    Private Sub Process_Template(filename As String)
        If filename <> "" Then
            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(filename)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
            Dim t1 As DataTable = dst.Tables("SPTCOOP1").Clone
            Dim t3 As DataTable = dst.Tables("SPTCOOP3").Clone
            EnforceConstraints(False)

            Dim AUTH_NO_LIST As New List(Of String)()
            Dim r As Integer = 2
            ' previous trackers
            Dim PREV_START_DATE As String = String.Empty
            Dim PREV_END_DATE As String = String.Empty
            Dim PREV_CUST_CODE As String = String.Empty
            Dim PREV_OTHER_COST As String = String.Empty
            Dim PREV_COLL_CODE As String = String.Empty
            Dim PREV_BOOKING As String = String.Empty

            Dim TOTAL_DIST_AMT As Decimal = 0D
            Dim AUTH_LNO_COUNT As Integer = 0
            Dim AUTH_NO_COUNT As Integer = 0
            Dim IS_NEW_GROUP As Boolean = True
            Dim FORCE_NEW_HEADER_NEXT As Boolean = False

            Dim EXPENSE_TYPE_CODE As String = String.Empty
            Dim TOTAL_AMT As Decimal = 0
            Dim DATE_START1 As Date = Date.MinValue
            Dim DATE_END As Date = Date.MinValue
            Dim SEASON_CODE As String = String.Empty
            Dim CUST_CODE As String = String.Empty
            Dim COLLECTION_CODEs As New List(Of String)
            Dim CUSTOMER_CODES As New List(Of String)

            Do While Not String.IsNullOrEmpty(oSheet.Cells(r, 4).Text) ' Loop as long as there's a customer code
                Dim CURR_START_DATE As String = Trim(oSheet.Cells(r, 1).Text & "")
                Dim CURR_END_DATE As String = Trim(oSheet.Cells(r, 2).Text & "")
                Dim CURR_CUST_CODE As String = Trim(oSheet.Cells(r, 4).Text)
                Dim CURR_OTHER_COST As String = Trim(oSheet.Cells(r, 17).Text)
                Dim CURR_COLL_CODE As String = Trim(oSheet.Cells(r, 19).Text & "")
                Dim CURR_BOOKING As String = Trim(oSheet.Cells(r, 7).Value & "")
                Dim CURR_DIST_AMT As Decimal = 0D : Decimal.TryParse(oSheet.Cells(r, 21).Text, CURR_DIST_AMT)
                TOTAL_DIST_AMT += CURR_DIST_AMT

                Dim custChanged As Boolean = (String.Compare(CURR_CUST_CODE, PREV_CUST_CODE, True) <> 0)
                Dim startChanged As Boolean = (String.Compare(CURR_START_DATE, PREV_START_DATE, True) <> 0)
                Dim endChanged As Boolean = (String.Compare(CURR_END_DATE, PREV_END_DATE, True) <> 0)
                Dim bookingChanged As Boolean = (String.Compare(CURR_BOOKING, PREV_BOOKING, True) <> 0)

                Dim isZeroDist As Boolean = (CURR_DIST_AMT = 0D)

                'Dim baseBreak As Boolean =
                '(String.IsNullOrEmpty(CURR_START_DATE) OrElse startChanged _
                ' OrElse custChanged _
                ' OrElse (String.Compare(CURR_OTHER_COST, PREV_OTHER_COST, True) <> 0) _
                ' OrElse bookingChanged _
                ' OrElse (TOTAL_DIST_AMT = Val(PREV_OTHER_COST)))

                Dim baseBreak As Boolean =
                (String.IsNullOrEmpty(CURR_START_DATE) OrElse startChanged _
                 OrElse custChanged _
                 OrElse (String.Compare(CURR_OTHER_COST, PREV_OTHER_COST, True) <> 0) _
                 OrElse bookingChanged)

                '$0 rule: start a new header if Booking changed OR Start changed OR End changed
                Dim zeroBreak As Boolean = (bookingChanged OrElse startChanged OrElse endChanged)

                If FORCE_NEW_HEADER_NEXT Then
                    IS_NEW_GROUP = True
                    FORCE_NEW_HEADER_NEXT = False
                End If

                Dim startNewHeader As Boolean
                If IS_NEW_GROUP Then
                    startNewHeader = True
                Else
                    startNewHeader = If(isZeroDist, zeroBreak, baseBreak)
                End If

                If startNewHeader Then
                    ' New header/auth
                    AUTH_NO = ASCMAIN1.Next_Control_No("SPTCOOP1.AUTH_NO")
                    PREV_START_DATE = CURR_START_DATE
                    PREV_END_DATE = CURR_END_DATE
                    PREV_CUST_CODE = CURR_CUST_CODE
                    PREV_COLL_CODE = CURR_COLL_CODE
                    TOTAL_DIST_AMT = CURR_DIST_AMT
                    PREV_BOOKING = CURR_BOOKING
                    PREV_OTHER_COST = CURR_OTHER_COST
                    AUTH_LNO_COUNT = 1
                    AUTH_NO_COUNT += 1
                    IS_NEW_GROUP = False

                    Dim rowSPTCOOP1XLS As DataRow = t1.NewRow()
                    rowSPTCOOP1XLS.Item("AUTH_NO") = AUTH_NO
                    If Not AUTH_NO_LIST.Contains(AUTH_NO) Then AUTH_NO_LIST.Add(AUTH_NO)

                    If String.IsNullOrEmpty(CURR_START_DATE) AndAlso Not String.IsNullOrEmpty(PREV_START_DATE) Then
                        rowSPTCOOP1XLS("DATE_START") = PREV_START_DATE
                    ElseIf Not String.IsNullOrEmpty(CURR_START_DATE) Then
                        rowSPTCOOP1XLS("DATE_START") = CURR_START_DATE
                        PREV_START_DATE = CURR_START_DATE
                    End If

                    If Not String.IsNullOrEmpty(CURR_END_DATE) Then
                        rowSPTCOOP1XLS("DATE_END") = CURR_END_DATE
                    End If

                    Date.TryParse(Trim(oSheet.Cells(r, 1).Text), DATE_START1)
                    Date.TryParse(Trim(oSheet.Cells(r, 2).Text), DATE_END)

                    If DATE_START1 = Date.MinValue Then EMsg &= vbCr & "Start Date is Mandatory"
                    If DATE_END = Date.MinValue Then EMsg &= vbCr & "End Date is Mandatory"
                    If DATE_START1 <> Date.MinValue AndAlso DATE_END <> Date.MinValue Then
                        If DATE_START1 > DATE_END Then EMsg &= vbCr & "Start Date may not be later than End Date"
                    End If
                    If Format(DATE_START1, "yyyy") < Mid(ASCMAIN1.CYP, 1, 4) Then
                        EMsg &= vbCr & "Start Date may not be prior to start of Current Ops Year"
                    End If
                    If Format(DATE_START1, "yyyy") > Format(Val(Mid(ASCMAIN1.CYP, 1, 4)) + 1, "0000") Then
                        EMsg &= vbCr & "Start Date may not be later than last day of next Year"
                    End If

                    SEASON_CODE = Trim(oSheet.Cells(r, 3).Text)
                    rowSPTCOOP1XLS("SEASON_CODE") = SEASON_CODE
                    If String.IsNullOrEmpty(SEASON_CODE) Then
                        EMsg &= vbCr & "You Must Specify a Season"
                    ElseIf Mid(SEASON_CODE, 1, 4) <> Format(DATE_START1, "yyyy") OrElse
                       (Mid(SEASON_CODE, 5, 1) = "S" And Format(DATE_START1, "MM") >= "07") OrElse
                       (Mid(SEASON_CODE, 5, 1) = "F" And Format(DATE_START1, "MM") < "07") Then
                        EMsg &= vbCr & "Season not congruous with Start Date"
                    End If

                    CUST_CODE = CURR_CUST_CODE
                    rowSPTCOOP1XLS("CUST_CODE") = CUST_CODE
                    If String.IsNullOrEmpty(CUST_CODE) Then
                        EMsg &= vbCr & "You must supply a Valid Customer"
                    ElseIf Not CUSTOMER_CODES.Contains(CUST_CODE) Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 Is Nothing Then
                            EMsg &= vbCr & "Customer Entered Is Not Valid on line " & r - 1
                        ElseIf rowARTCUST1("CUST_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Customer Entered Is Not Active on line " & r - 1
                        Else
                            CUSTOMER_CODES.Add(CUST_CODE)
                        End If
                    End If

                    ASCMAIN1.sql = "SELECT SREP_CODE FROM ARTCUST1 WHERE CUST_CODE = '" & CUST_CODE & "'"
                    rowSPTCOOP1XLS("SREP_CODE") = ASCDATA1.GetDataValue() & String.Empty

                    EXPENSE_TYPE_CODE = Trim(oSheet.Cells(r, 5).Text)
                    rowSPTCOOP1XLS("EXPENSE_TYPE_CODE") = EXPENSE_TYPE_CODE

                    Dim VEHICLE_CODE As String = Trim(oSheet.Cells(r, 6).Value & "")
                    If VEHICLE_CODE = "" Then EMsg &= vbCr & "You Must Specify an Advertising Vehicle Code"
                    rowSPTCOOP1XLS.Item("VEHICLE_CODE") = VEHICLE_CODE

                    Dim BOOKING_NAME As String = Trim(oSheet.Cells(r, 7).Value & "")
                    If BOOKING_NAME = "" Then EMsg &= vbCr & "You Must Specify the Booking Name of this Sales Promotion Event on line " & r - 1
                    rowSPTCOOP1XLS.Item("BOOKING_NAME") = BOOKING_NAME

                    rowSPTCOOP1XLS.Item("BOOKED_BY") = Trim(oSheet.Cells(r, 8).Value & "")
                    Dim qty As String = Trim(oSheet.Cells(r, 9).Value & "")
                    If Not String.IsNullOrEmpty(qty) Then rowSPTCOOP1XLS.Item("QTY") = qty
                    rowSPTCOOP1XLS.Item("NOTES") = Trim(oSheet.Cells(r, 10).Value & "")

                    Dim EVENT_TYPE_CODE As String = Trim(oSheet.Cells(r, 11).Value & "")
                    If EVENT_TYPE_CODE = "" Then EMsg &= vbCr & "You Must Specify an Event Type Code on line " & r - 1
                    rowSPTCOOP1XLS.Item("EVENT_TYPE_CODE") = EVENT_TYPE_CODE

                    rowSPTCOOP1XLS.Item("AUTH_DATE") = Trim(oSheet.Cells(r, 12).Text & "")
                    rowSPTCOOP1XLS.Item("AUTH_REQ_BY") = Trim(oSheet.Cells(r, 13).Value & "")
                    rowSPTCOOP1XLS.Item("CUST_REF_NUM") = Trim(oSheet.Cells(r, 14).Value & "")
                    rowSPTCOOP1XLS.Item("EVENT_QUALIFIER") = Trim(oSheet.Cells(r, 15).Value & "")
                    rowSPTCOOP1XLS.Item("VERIFIED_AS_OPEN_COMMENTS") = Trim(oSheet.Cells(r, 16).Value & "")
                    rowSPTCOOP1XLS.Item("OTHER_COST") = Trim(oSheet.Cells(r, 17).Value & "")
                    rowSPTCOOP1XLS.Item("OPEN_AMT") = rowSPTCOOP1XLS.Item("OTHER_COST")
                    rowSPTCOOP1XLS.Item("INIT_OPER") = ASCMAIN1.USER_ID
                    rowSPTCOOP1XLS.Item("INIT_DATE") = DATETIME_STAMP
                    rowSPTCOOP1XLS.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    rowSPTCOOP1XLS.Item("LAST_DATE") = DATETIME_STAMP
                    rowSPTCOOP1XLS.Item("APPR_STATUS_CODE") = "P"
                    rowSPTCOOP1XLS.Item("STATUS_CODE") = "O"
                    rowSPTCOOP1XLS.Item("EVENT_DATE_CHANGED") = DATETIME_STAMP

                    If PREV_START_DATE <> "" And CURR_START_DATE <> "" Then
                        ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(rowSPTCOOP1XLS.Item("DATE_START"), "dd-MMM-yyyy") & "'"
                        Dim YW As String = ASCDATA1.GetDataValue
                        If YW <> "" Then rowSPTCOOP1XLS.Item("OPS_YYYYWW") = YW
                    End If
                    rowSPTCOOP1XLS.Item("OPS_YYYYPP") = Format(rowSPTCOOP1XLS.Item("DATE_START"), "yyyyMM")

                    t1.Rows.Add(rowSPTCOOP1XLS)
                Else
                    ' same header – next detail line
                    AUTH_LNO_COUNT += 1
                End If

                ' detail row (SPTCOOP3)
                Dim rowSPTCOOP3XLS As DataRow = t3.NewRow()
                rowSPTCOOP3XLS.Item("AUTH_NO") = AUTH_NO
                rowSPTCOOP3XLS.Item("AUTH_LNO") = AUTH_LNO_COUNT
                rowSPTCOOP3XLS.Item("ITEM_CODE") = Trim(oSheet.Cells(r, 18).Text & "")
                rowSPTCOOP3XLS.Item("COLLECTION_CODE") = Trim(oSheet.Cells(r, 19).Text & "")
                rowSPTCOOP3XLS.Item("FEATURE_DESC") = Trim(oSheet.Cells(r, 20).Text & "")
                rowSPTCOOP3XLS.Item("DIST_AMT") = Trim(oSheet.Cells(r, 21).Value & "")
                t3.Rows.Add(rowSPTCOOP3XLS)

                Dim headerAmt As Decimal = 0D
                Decimal.TryParse(PREV_OTHER_COST, headerAmt)

                If CURR_DIST_AMT <> 0D AndAlso headerAmt > 0D AndAlso Math.Round(TOTAL_DIST_AMT, 2) >= Math.Round(headerAmt, 2) Then
                    FORCE_NEW_HEADER_NEXT = True
                End If


                ' update previous trackers for next comparison
                PREV_CUST_CODE = CURR_CUST_CODE
                PREV_START_DATE = CURR_START_DATE
                PREV_END_DATE = CURR_END_DATE
                PREV_BOOKING = CURR_BOOKING
                PREV_OTHER_COST = CURR_OTHER_COST

                r += 1
            Loop

            If t3.Rows.Count = 0 Then
                EMsg &= vbCr & "No Distribution Details Entered"
            Else
                For Each row As DataRow In t3.Select("")
                    Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE") & ""
                    If COLLECTION_CODE = "" Then
                        If EXPENSE_TYPE_CODE <> "RTLEVENTS" Then
                            EMsg &= vbCr & "Invalid Expense Type for Blank Collection Distributions (" & Absx1.txtFor("EXPENSE_TYPE_CODE").Text & ")"
                        End If
                        If System.Math.Round(TOTAL_AMT, 2) <> 0 Then
                            EMsg &= vbCr & "Blank Collection Distributions not valid with $$$ Events"
                        End If
                    Else
                        If COLLECTION_CODEs.Contains(COLLECTION_CODE) Then
                            ' already validated
                        Else
                            If LookUp("ICTCOLL1", COLLECTION_CODE) Is Nothing Then
                                EMsg &= vbCr & "Invalid Collection Code (" & COLLECTION_CODE & ") on line " & r - 1
                            Else
                                COLLECTION_CODEs.Add(COLLECTION_CODE)
                            End If
                        End If
                    End If
                Next
            End If

            Dim distributionSum As Decimal = 0D
            If t3.Rows.Count > 0 Then
                Dim sumObj As Object = t3.Compute("SUM(DIST_AMT)", "")
                If Not IsDBNull(sumObj) AndAlso sumObj IsNot Nothing AndAlso sumObj.ToString() <> "" Then
                    Decimal.TryParse(sumObj.ToString(), distributionSum)
                End If
            End If

            Dim totalSum As Decimal = 0D
            If t1.Rows.Count > 0 Then
                Dim sumObj2 As Object = t1.Compute("SUM(OPEN_AMT)", "")
                If Not IsDBNull(sumObj2) AndAlso sumObj2 IsNot Nothing AndAlso sumObj2.ToString() <> "" Then
                    Decimal.TryParse(sumObj2.ToString(), totalSum)
                End If
            End If

            If Math.Round(distributionSum, 2) <> Math.Round(totalSum, 2) Then
                EMsg &= vbCr & "Total of Distribution does not agree with Total for Commitment"
            End If
            If t3.Select("ISNULL(COLLECTION_CODE,'')='' AND ISNULL(DIST_AMT,0)<>0").Length <> 0 Then
                EMsg &= vbCr & "You may not distribute an expense amount without specifying a Collection"
            End If

            If EMsg = "" Then
                For Each row As DataRow In t1.Select()
                    dst.Tables("SPTCOOP1").Rows.Add(row.ItemArray)
                Next
                For Each row As DataRow In t3.Select()
                    dst.Tables("SPTCOOP3").Rows.Add(row.ItemArray)
                Next
                EnforceConstraints(True)
                Dim lineCount As Integer = r - 2
                If Confirm_Upload(filename, lineCount, AUTH_NO_COUNT, distributionSum) Then
                    Perform_Update(AUTH_NO_LIST)
                End If
            Else
                MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
                Exit Sub
            End If
        End If
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Private Sub Perform_Update(ByVal AUTH_NO_LIST As List(Of String))
        BeginTrans()

        For Each AUTH_NO In AUTH_NO_LIST
            ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "ADD", "Contract Created", "")
        Next
        ' Record an event for each AUTH_NO
        Dim SQLD As String = "AUTH_NO = '" & AUTH_NO & "'"
        INIT_LAST("SPTCOOP1", False, , True)

        ' Update records in database
        Update_Record_TDA("SPTCOOP1", SQLD)
        Update_Record_TDA("SPTCOOP3", SQLD)

        CommitTrans("Update Complete")
    End Sub
    Public Sub Create_Template()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Creating New Template")
        Dim FILENAME As String
        Dim desktopPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop)
        Dim filePath As String = Path.Combine(desktopPath, "PromoEventTemplate - Blank.xlsx")
        ASCMAIN1.Progress("Now Creating Workbook")

        'if running on pdb then use the template on the desktop 
        If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then
            FILENAME = desktopPath & "\" & Me.Name & ".xlsx"
        Else
            FILENAME = ASCMAIN1.Folders("SharedRoot") & "Templates\" & Me.Name & ".xlsx"
        End If
        ' Create a new Excel application instance
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()

        ' Open the workbook
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = excelApp.Workbooks.Open(FILENAME)

        ' Access the "Lists" sheet
        Dim listsSheet As Microsoft.Office.Interop.Excel.Worksheet = wb.Sheets("Lists")
        listsSheet.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden

        ' Define your SQL queries and corresponding header names
        Dim queries As New List(Of String) From {
            "SELECT DISTINCT BOOKED_BY FROM SPTCOOP1 WHERE BOOKED_BY IS NOT NULL ORDER BY BOOKED_BY",
            "Select DISTINCT collection_code from ictcoll1 WHERE COLLECTION_CODE IS NOT NULL order by collection_code",
            "SELECT DISTINCT EVENT_TYPE_CODE FROM SPTCOOP1 WHERE EVENT_TYPE_CODE IS NOT NULL ORDER BY EVENT_TYPE_CODE",
            "SELECT DISTINCT CUST_CODE FROM ARTCUST1 WHERE CUST_CODE IS NOT NULL ORDER BY CUST_CODE",
            "SELECT DISTINCT EXPENSE_TYPE_CODE FROM SPTCOOP1 WHERE EXPENSE_TYPE_CODE IS NOT NULL ORDER BY EXPENSE_TYPE_CODE",
            "SELECT DISTINCT SEASON_CODE FROM SPTCOOP1 WHERE SEASON_CODE IS NOT NULL ORDER BY SEASON_CODE",
            "SELECT DISTINCT VEHICLE_CODE FROM SPTCOOP1 WHERE VEHICLE_CODE IS NOT NULL ORDER BY VEHICLE_CODE"
        }
        Dim headers As New List(Of String) From {
            "BOOKED_BY",
            "COLLECTION_CODE",
            "EVENT_TYPE_CODE",
            "CUST_CODE",
            "EXPENSE_TYPE_CODE",
            "SEASON_CODE",
            "VEHICLE_CODE"
        }
        For columnIndex As Integer = 0 To queries.Count - 1
            listsSheet.Range(listsSheet.Cells(1, columnIndex + 1), listsSheet.Cells(listsSheet.Rows.Count, columnIndex + 1)).ClearContents()
            listsSheet.Cells(1, columnIndex + 1).Value = headers(columnIndex)
            ASCMAIN1.sql = queries(columnIndex)

            ' Execute the SQL query and retrieve results
            Dim rowIndex As Integer = 2 ' Start from the second row for data
            For Each row As DataRow In ASCDATA1.GetDataTable.Select()
                Dim stuff As String = row.Item(0).ToString()

                ' Write the result to the sheet
                listsSheet.Cells(rowIndex, columnIndex + 1).Value = stuff
                rowIndex += 1
            Next
            Dim lastRow As Integer = listsSheet.Cells(listsSheet.Rows.Count, columnIndex + 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            Dim columnName As String = GetExcelColumnName(columnIndex + 1)
            Dim rangeName As String = headers(columnIndex)
            Dim rangeAddress As String = $"${columnName}$2:${columnName}${lastRow}"
            ' Create or update the named range
            Dim namedRange As Microsoft.Office.Interop.Excel.Name
            Try
                namedRange = wb.Names.Item(rangeName)
                namedRange.RefersTo = $"=Lists!{rangeAddress}"
            Catch ex As Exception
                wb.Names.Add(Name:=rangeName, RefersTo:=$"=Lists!{rangeAddress}")
            End Try

        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0
        Dim XLS_FILENAME As String = ""
        Do Until success
            Try
                XLS_NO += 1
                XLS_FILENAME = "UpdatedTemplate"
                XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"
                Dim objOpt As Object = Nothing ' Missing.Value
                wb.SaveAs(Path.Combine(desktopPath, XLS_FILENAME))
                success = True
                excelApp.Visible = True
            Catch ex As Exception
                ' Stop
            End Try
        Loop
    End Sub
    Private Function GetExcelColumnName(columnNumber As Integer) As String
        Dim columnName As String = ""
        While columnNumber > 0
            Dim modulo As Integer = (columnNumber - 1) Mod 26
            columnName = Convert.ToChar(65 + modulo).ToString() & columnName
            columnNumber = (columnNumber - modulo) \ 26
        End While
        Return columnName
    End Function
End Class