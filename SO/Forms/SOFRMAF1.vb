Public Class SOFRMAF1

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String             ' Sold-To Customer Name
    Dim CUST_BILL_TO_CUST As String
    Dim CUST_CLAIM_NO As String

    Dim PRICE_CLASS_CODE As String
    Dim PRICE_BASE_DPCT As Decimal
    Dim PRICE_BASIS As String
    Dim PRICE_LIST_CODE As String

    Dim PRICE_LIST_CODE_ALLO As String
    Dim CUST_CODE_ALLO As String

    Dim RA_NO As String
    Dim rowSOTRMAF1 As DataRow
    Dim RA_NO_EDI As String

    Dim rowARTCUST1 As DataRow          ' ARTCUST1 for the Sold-To
    Dim rowARTCUST1_BT As DataRow       ' ARTCUST1 for the Bill-To
    Dim rowICTITEM1 As DataRow
    Dim RA_LNOs As New List(Of Int64)   ' list of RA_LNOs that are deleted

    Dim blnReturnsHaveBeenApplied As Boolean
    Dim EDI_DOC_SEQ_NOs_no_company As New List(Of String)

    Dim SOTRMAFX As String
    Dim SOTRMAFI As String
    Dim blnAutoPilot180 As Boolean = False
    Dim blnAutoPilot812 As Boolean = False

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Check_InquiryMode()
        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Create_Work_Tables(" and rownum < 1")

        MakeTransparent(chk_showDetails180)
        SplitContainer4.Panel2Collapsed = True

        With dst

            ASCMAIN1.sql = "Select SOTRTRN1.*" _
                & " from SOTRTRN1 where SOTRTRN1.RA_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTRTRNX", "**", 0, False, "V")

            ASCMAIN1.sql = "Select * from " & SOTRMAFX
            Create_TDA(.Tables.Add, "SOTRMAFX", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from " & SOTRMAFI
            Create_TDA(.Tables.Add, "SOTRMAFI", "**", 0, False)

            Create_TDA(.Tables.Add, "SOTRMAF1", "*", 1)
            .Tables("SOTRMAF1").Columns.Add("RA_AMT", GetType(System.Decimal))
            .Tables("SOTRMAF1").Columns.Add("AR_PARM_KEY")

            ASCMAIN1.sql = "Select SOTRMAF2.*, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_SNU_CODE" _
            & " from SOTRMAF2,ICTITEM1" _
            & " where ICTITEM1.ITEM_CODE = SOTRMAF2.ITEM_CODE"
            Create_TDA(.Tables.Add, "SOTRMAF2", "**", 1)
            With .Tables("SOTRMAF2").Columns
                .Add("RA_AMT", GetType(System.Decimal), "ISNULL(RA_QTY,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_OPEN", GetType(System.Decimal), "ISNULL(RA_QTY_OPEN,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_USED", GetType(System.Decimal), "ISNULL(RA_QTY_USED,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_CANC", GetType(System.Decimal), "ISNULL(RA_QTY_CANC,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_RETAIL_EXT", GetType(System.Decimal), "IIF(ISNULL(RA_QTY,0)=0,RA_AMT/((100 - 0) / 100),ISNULL(RA_QTY,0) * ISNULL(RA_RETAIL,0))")
                .Add("QTY_EOW", GetType(System.Int64))
            End With

            With .Tables.Add("SOTRMAFT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() { .Columns("KEY")}
            End With

            Create_TDA(.Tables.Add, "ARTCUST1", "*", , False)
            Create_TDA(.Tables.Add, "ARTCUST2", "*", , False)
            Create_TDA(.Tables.Add, "ICTWHSE1", "*", , False)
            Create_TDA(.Tables.Add, "SOTSREP1", "*", , False)

            Create_TDA(.Tables.Add, "ARTOPEN1", "*", 0)
            Create_TDA(.Tables.Add, "SOTINVH1", "*", 0)
            Create_TDA(.Tables.Add, "SOTINVH2", "*", 0)

            With .Tables.Add("SOTRMAF0")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("ADDRESS0")
                .Columns.Add("ADDRESS1")
                .Columns.Add("ADDRESS2")
                .Columns.Add("ADDRESS3")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() { .Columns("AR_PARM_KEY")}
            End With

            ASCMAIN1.sql = "SELECT SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE, SOTINVH2.ORDR_UNIT_PRICE, SOTINVH2.ORDR_QTY_SHIP" _
             & " FROM SOTINVH1, SOTINVH2" _
             & " WHERE SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
             & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO" _
             & " AND SOTINVH1.INV_TYPE = 'I'" _
             & " AND SOTINVH1.CUST_CODE = :PARM1" _
             & " AND SOTINVH1.ORDR_CUST_PO = :PARM2"
            Create_TDA(.Tables.Add, "SOTINVH2X", ASCMAIN1.sql, 0, False, "VV", 0)

            ASCMAIN1.sql = "Select EDT180T1.* from EDT180T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            Create_TDA(.Tables.Add, "EDT180TX", ASCMAIN1.sql, 0, False, "VV", 0)
            .Tables("EDT180TX").Columns.Add("SEL")
            .Tables("EDT180TX").Columns("SEL").DefaultValue = "0"
            .Tables("EDT180TX").PrimaryKey = New DataColumn() { .Tables("EDT180TX").Columns("EDI_DOC_SEQ_NO")}

            ASCMAIN1.sql = "Select EDT180T1.* from EDT180T1 where rownum < 1"
            Create_TDA(.Tables.Add, "EDT180TY", ASCMAIN1.sql, 0, False, "VV", 0)
            .Tables("EDT180TY").Columns.Add("SEL")
            .Tables("EDT180TY").Columns("SEL").DefaultValue = "1"
            .Tables("EDT180TY").PrimaryKey = New DataColumn() { .Tables("EDT180TY").Columns("EDI_DOC_SEQ_NO")}


            'ASCMAIN1.sql = "Select Distinct EDT180T2.*" & vbCrLf _
            '    & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
            '    & " from EDT180T2,ICTITEM1" & vbCrLf _
            '    & " where EDT180T2.EDI_EAN is not null and ICTITEM1.ITEM_EAN_CODE (+) = EDT180T2.EDI_EAN" & vbCrLf _
            '    & "   and EDT180T2.EDI_DOC_SEQ_NO = :PARM1" & vbCrLf _
            '    & " union " & vbCrLf _
            '    & "Select Distinct EDT180T2.*" & vbCrLf _
            '    & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
            '    & " from EDT180T2,ICTITEM1" & vbCrLf _
            '    & " where EDT180T2.EDI_EAN is null and EDT180T2.EDI_UPC is not null and ICTITEM1.ITEM_UPC_CODE (+) = EDT180T2.EDI_UPC" & vbCrLf _
            '    & "   and EDT180T2.EDI_DOC_SEQ_NO = :PARM1"
            ASCMAIN1.sql = "Select Distinct EDT180T2.*" & vbCrLf _
                & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & " from EDT180T2,ICTITEM1" & vbCrLf _
                & " where EDT180T2.EDI_EAN is not null and ICTITEM1.ITEM_EAN_CODE (+) = EDT180T2.EDI_EAN" & vbCrLf _
                & "   and EDT180T2.EDI_DOC_SEQ_NO = :PARM1" & vbCrLf _
                & " union " & vbCrLf _
                & "Select Distinct EDT180T2.*" & vbCrLf _
                & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & " from EDT180T2,ICTITEM1" & vbCrLf _
                & " where EDT180T2.EDI_EAN is null and ICTITEM1.ITEM_UPC_CODE (+) = EDT180T2.EDI_UPC" & vbCrLf _
                & "   and EDT180T2.EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "EDT180T2", "**", 0, False, "V")
            .Tables("EDT180T2").Columns.Add("ORDR_UNIT_PRICE", GetType(System.Decimal))

            ASCMAIN1.sql = "Select EDT812T1.* from EDT812T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            Create_TDA(.Tables.Add, "EDT812TX", ASCMAIN1.sql, 0, False, "VV", 0)
            .Tables("EDT812TX").Columns.Add("SEL")
            .Tables("EDT812TX").Columns("SEL").DefaultValue = "0"
            .Tables("EDT812TX").PrimaryKey = New DataColumn() { .Tables("EDT812TX").Columns("EDI_DOC_SEQ_NO")}


            ASCMAIN1.sql = "Select Distinct EDT812T2.*" & vbCrLf _
                & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & " from EDT812T2,ICTITEM1" & vbCrLf _
                & " where EDT812T2.EDI_EAN is not null and ICTITEM1.ITEM_EAN_CODE (+) = EDT812T2.EDI_EAN" & vbCrLf _
                & "   and EDT812T2.EDI_DOC_SEQ_NO = :PARM1" & vbCrLf _
                & " union " & vbCrLf _
                & "Select Distinct EDT812T2.*" & vbCrLf _
                & ", ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & " from EDT812T2,ICTITEM1" & vbCrLf _
                & " where EDT812T2.EDI_EAN is null and ICTITEM1.ITEM_UPC_CODE (+) = EDT812T2.EDI_UPC" & vbCrLf _
                & "   and EDT812T2.EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "EDT812T2", "**", 0, False, "V")
            .Tables("EDT812T2").Columns.Add("ORDR_UNIT_PRICE", GetType(System.Decimal))


            ASCMAIN1.sql = "select T1.XNO, MAX(TRUNC(f1.INIT_DATE)) DATE_PROCESSED, COUNT(DISTINCT f1.RA_NO) DIF_COUNT, SUM(F2.RA_NET_PRICE*F2.RA_QTY) DIF_TOTAL_AMT, MAX(F1.INIT_OPER) INIT_OPER
                            from edt180t1 t1
                            join sotrmaf1 f1 on (f1.edi_doc_seq_no=t1.edi_doc_seq_no and f1.edi_doc_no = '180')
                            join sotrmaf2 f2 on (f1.ra_no=f2.ra_no)
                            where T1.XNO IS NOT NULL
                            group by T1.XNO"
            Create_TDA(.Tables.Add, "EDT180X0", "**", 0, False)
            .Tables("EDT180X0").Columns.Add("SEL")
            .Tables("EDT180X0").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "SELECT * FROM EDT180T2 WHERE EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "EDT180X2", "**", 0, False, "V")

            ASCMAIN1.sql = "select T1.XNO, MAX(TRUNC(f1.INIT_DATE)) DATE_PROCESSED, COUNT(DISTINCT F1.RA_NO) DIF_COUNT, SUM(F2.RA_NET_PRICE*F2.RA_QTY) DIF_TOTAL_AMT, MAX(f1.INIT_OPER) INIT_OPER
                            from edt812t1 t1
                            join sotrmaf1 f1 on (f1.edi_doc_seq_no=t1.edi_doc_seq_no and f1.edi_doc_no = '812')
                            join sotrmaf2 f2 on (f1.ra_no=f2.ra_no)
                            where T1.XNO IS NOT NULL
                            group by T1.XNO"
            Create_TDA(.Tables.Add, "EDT812X0", "**", 0, False)
            .Tables("EDT812X0").Columns.Add("SEL")
            .Tables("EDT812X0").Columns("SEL").DefaultValue = "0"


            ASCMAIN1.sql = "select t1.XNO,t1.EDI_DOC_SEQ_NO, t1.RMA_NUMBER, f1.CUST_CLAIM_NO, f1.RA_NO, f1.INV_NUM, f1.RA_DATE TRANS_DATE, t1.EDI_RECD_DATE, SUM(f2.RA_NET_PRICE*f2.RA_QTY) TOTAL_AMT,f1.INIT_OPER from edt180t1 t1
                        join sotrmaf1 f1 on (f1.EDI_DOC_SEQ_NO=t1.EDI_DOC_SEQ_NO and f1.edi_doc_no = '180')
                        join sotrmaf2 f2 on (f1.ra_no=f2.ra_no)
                        where (t1.XNO = :PARM1) OR (T1.CUST_CODE = :PARM2 AND F1.RA_DATE BETWEEN :PARM3 AND :PARM4 AND f1.CUST_CLAIM_NO LIKE NVL2(:PARM5,'%' || :PARM5 || '%', f1.CUST_CLAIM_NO))
                        group by t1.XNO, t1.EDI_DOC_SEQ_NO, f1.RA_NO, f1.INV_NUM, t1.RMA_NUMBER, f1.CUST_CLAIM_NO, f1.RA_DATE, t1.EDI_RECD_DATE,f1.INIT_OPER"
            Create_TDA(.Tables.Add, "EDT180X1", "**", 0, False, "VVDDV")

            ASCMAIN1.sql = "select t1.XNO,t1.EDI_DOC_SEQ_NO, t1.EDI_RMA_NO RMA_NUMBER, f1.CUST_CLAIM_NO, F1.RA_NO, f1.INV_NUM, f1.RA_DATE TRANS_DATE, t1.EDI_RECEIVED_DATE EDI_RECD_DATE, SUM(f2.RA_NET_PRICE*f2.RA_QTY) TOTAL_AMT,f1.INIT_OPER from edt812t1 t1
                        join sotrmaf1 f1 on (f1.EDI_DOC_SEQ_NO=t1.EDI_DOC_SEQ_NO and f1.edi_doc_no = '812')
                        join sotrmaf2 f2 on (f1.ra_no=f2.ra_no)
                        where (t1.XNO = :PARM1) OR (T1.CUST_CODE = :PARM2 AND F1.RA_DATE BETWEEN :PARM3 AND :PARM4 AND f1.CUST_CLAIM_NO LIKE NVL2(:PARM5,'%' || :PARM5 || '%', f1.CUST_CLAIM_NO))
                        group by t1.XNO, t1.EDI_DOC_SEQ_NO, f1.RA_NO, f1.INV_NUM, t1.EDI_RMA_NO, f1.CUST_CLAIM_NO, f1.RA_DATE, t1.EDI_RECEIVED_DATE,f1.INIT_OPER"
            Create_TDA(.Tables.Add, "EDT812X1", "**", 0, False, "VVDDV")


            ASCMAIN1.sql = "select DISTINCT C1.CUST_CODE,C1.CUST_NAME 
                            FROM SOTRMAF1 F1 
                            JOIN ARTCUST1 C1 ON (F1.CUST_CODE=C1.CUST_CODE) 
                            WHERE F1.RA_REASON_CODE='X' AND F1.EDI_DOC_SEQ_NO IS NOT NULL AND F1.EDI_DOC_NO=:PARM1 ORDER BY CUST_NAME"
            Create_TDA(.Tables.Add, "ARTCUSTS", "**", 0, False, "V")

        End With

        Fill_Records("ARTCUSTS", optDifReportType.Value)
        cbeSearchCust.DataSource = dst.Tables("ARTCUSTS")
        cbeSearchCust.SelectedIndex = 0
        cbeSearchCust.DisplayMember = "CUST_NAME"
        cbeSearchCust.ValueMember = "CUST_CODE"

        dteSearchFrom.DateTime = Now.Date.AddYears(-1)
        dteSearchTo.DateTime = Now.Date

        grdSOTRMAFX.DataSource = dst.Tables("SOTRMAFX")
        grdSOTRMAFI.DataSource = dst.Tables("SOTRMAFI")
        grdSOTRMAF2.DataSource = dst.Tables("SOTRMAF2")
        grdSOTRMAFT.DataSource = dst.Tables("SOTRMAFT")
        grdSOTRTRNX.DataSource = dst.Tables("SOTRTRNX")
        grdEDT180TX.DataSource = dst.Tables("EDT180TX")
        grdEDT180T2.DataSource = dst.Tables("EDT180X2")
        grdEDT812TX.DataSource = dst.Tables("EDT812TX")

        grdEDTXXXX0.DataSource = dst.Tables("EDT180X0")
        Sort_grdColumns(grdEDTXXXX0, "date_processed")


        grdEDTXXXX1.DataSource = dst.Tables("EDT180X1")
        Sort_grdColumns(grdEDTXXXX1, "CUST_CLAIM_NO")

        Create_Summary(grdEDT180TX, "EDI_DOC_SEQ_NO", "Count")
        Create_Summary(grdEDT180TX, New String() {"SEL"})
        Create_Summary(grdEDT180TX, New String() {"REQUIRES_RESOLUTION"})

        Create_Summary(grdEDTXXXX1, "EDI_DOC_SEQ_NO", "Count")
        Create_Summary(grdEDTXXXX1, "TOTAL_AMT", "Sum")

        Create_Summary(grdEDT812TX, "EDI_DOC_SEQ_NO", "Count")
        Create_Summary(grdEDT812TX, New String() {"SEL"})
        Create_Summary(grdEDT812TX, New String() {"REQUIRES_RESOLUTION"})

        Create_Summary(grdEDTXXXX0, New String() {"SEL"})

        Create_Summary(grdSOTRTRNX, "RTRN_NO", "Count")
        Create_Summary(grdSOTRTRNX, New String() {"RTRN_COSTS", "RTRN_SALES", "RTRN_AMOUNT"})


        With grdEDT180TX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            .Override.AllowUpdate = DefaultableBoolean.True
        End With

        With grdEDT812TX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            .Override.AllowUpdate = DefaultableBoolean.True
        End With

        With grdEDTXXXX0.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            .Override.AllowUpdate = DefaultableBoolean.True
        End With

        With grdSOTRTRNX.DisplayLayout.Bands("SOTRTRNX")
            .Columns("RTRN_NO").Header.Fixed = True
        End With

        grdSOTRMAFX.DisplayLayout.UseFixedHeaders = True
        With grdSOTRMAFX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RA_NO", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdSOTRMAF2.DisplayLayout.UseFixedHeaders = True
        With grdSOTRMAF2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RA_LNO", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTRMAFX, grdSOTRMAFI}
            With grd.DisplayLayout.Bands(0)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackColor = Drawing.Color.White
                    If New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        gcol.Width = 80
                    ElseIf New String() {"RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        gcol.Width = 90
                    ElseIf New String() {"RA_RETAIL_EXT"}.Contains(gcol.Key) Then
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        gcol.Width = 90
                    Else
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    End If
                Next
            End With
        Next


        With grdSOTRMAF2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"ITEM_CODE", "RA_QTY", "RA_QTY_OPEN", "RA_NET_PRICE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"ITEM_CODE", "ITEM_DESC", "ITEM_RETAIL_PRICE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"RA_NET_PRICE", "RA_RETAIL_EXT", "RA_LINE_AMT"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With

        Create_Summary(grdSOTRMAFX, "RA_NO", "Count")
        Create_Summary(grdSOTRMAFX, New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC",
                                          "RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC", "RA_RETAIL_EXT"})

        Create_Summary(grdSOTRMAFI, "RA_NO", "Count")
        Create_Summary(grdSOTRMAFI, New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC",
                                          "RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC", "RA_RETAIL_EXT"})

        Create_Summary(grdSOTRMAF2, "RA_LNO", "Count")
        Create_Summary(grdSOTRMAF2, New String() {"RA_QTY", "RA_QTY_OPEN", "RA_QTY_USED", "RA_QTY_CANC",
                                                  "RA_AMT", "RA_AMT_OPEN", "RA_AMT_USED", "RA_AMT_CANC", "RA_LINE_AMT", "RA_RETAIL_EXT"})

        With dst.Tables("SOTRMAFT").Rows
            .Add(New Object() {1, "Auth", 0, 0})
            .Add(New Object() {2, "Open", 0, 0})
            .Add(New Object() {3, "Used", 0, 0})
            .Add(New Object() {4, "Canc", 0, 0})
        End With
        Sort_grdColumns(grdSOTRMAFT, "KEY", True)

        Dim rowSOTRMAF0 As DataRow = dst.Tables("SOTRMAF0").NewRow
        With ROWs("ARTPARM1")
            rowSOTRMAF0.Item("AR_PARM_KEY") = "Z"
            rowSOTRMAF0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowSOTRMAF0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowSOTRMAF0.Item("REMIT2") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            rowSOTRMAF0.Item("REMIT3") = "Tel " & .Item("AR_PARM_REMIT_PHONE") & " Fax " & .Item("AR_PARM_REMIT_FAX")
            'rowSOTRMAF0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
        End With
        rowSOTRMAF0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        dst.Tables("SOTRMAF0").Rows.Add(rowSOTRMAF0)


        Show_Filter(grdSOTRMAFX, True)
        grdSOTRMAFX.DisplayLayout.GroupByBox.Hidden = False
        SplitContainer1.Panel2Collapsed = True

        Show_Filter(grdSOTRMAFI, True)
        grdSOTRMAFI.DisplayLayout.GroupByBox.Hidden = False

        '  ASCMAIN1.Add_Value_List(grdSOTRMAFX, "RA_REASON_CODE", Nothing, New String() {":", "D:Damaged", "X:Destroyed", "O:Overstock", "Z:Other"})
        ASCMAIN1.Add_Value_List(grdSOTRMAFX, "RA_REASON_CODE")
        ASCMAIN1.Add_Value_List(grdSOTRMAFX, "RA_STATUS", Nothing, New String() {":", "O:Open", "F:Completed", "D:Deleted", "C:Cancelled"})
        ASCMAIN1.Add_Value_List(grdSOTRMAFI, "RA_REASON_CODE")
        ASCMAIN1.Add_Value_List(grdSOTRMAFI, "RA_STATUS", Nothing, New String() {":", "O:Open", "F:Completed", "D:Deleted", "C:Cancelled"})
        ASCMAIN1.Add_Value_List(grdSOTRTRNX, "REASON_CODE", Nothing, New String() {":", "D:Damaged", "X:Destroyed", "O:Overstock", "Z:Other"})


        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
            tabRAMaster.Tabs("EDI 180s").Visible = False
            tabRAMaster.Tabs("EDI 812s").Visible = False
        End If

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            lblSALES_DIVISION_CODE.Visible = False
            txtSALES_DIVISION_CODE.Visible = False
            txtSALES_DIVISION_NAME.Visible = False
            lblSREP2_CODE.Visible = False
            txtSREP2_CODE.Visible = False
            txtSREP2_NAME.Visible = False
        End If
        If ASCMAIN1.CLIENT = "AHA" Then
            dteRAFrom.Value = Now.Date.AddDays(-365)
        Else
            dteRAFrom.Value = Now.Date.AddDays(-30)
        End If
        dteRATo.Value = Now.Date
        If ASCMAIN1.CLIENT = "AHA" Then
            Absx1.dteFor("RA_START_DATE").Visible = True
            lblRA_START_DATE.Visible = True
        Else
            Absx1.dteFor("RA_START_DATE").Visible = False
            lblRA_START_DATE.Visible = False
        End If

        grdSOTRMAF2.DisplayLayout.Bands(0).Columns("QTY_EOW").Hidden = Not (ASCMAIN1.CLIENT = "INT")
        If (ASCMAIN1.CLIENT = "INT") Then
            Create_Summary(grdSOTRMAF2, "QTY_EOW")
        End If

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFRMAFI")
    End Sub

    Public Overrides Function Audit_Context() As Audit_Entity

        Dim E As New Audit_Entity

        E.TABLE_NAME = "SOFRMAF1"
        E.TABLE_DESC = ""
        E.KEY_VALUE = ""
        E.KEY_DESC = ""

        Return E
    End Function

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    Else
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                blnReturnsHaveBeenApplied = False

                CUST_CLAIM_NO = Absx1.txtFor("CUST_CLAIM_NO").Text
                If CUST_CLAIM_NO = "" Then
                    If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                        ' SP SAYS NOT TO BOTHER
                    Else
                        EMsg &= vbCr & "You Must Provide a Value for Customer Claim No"
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTRMAF1", CUST_CODE) Then Exit Sub
                End If

            Case "Edit", "View"

                CUST_CODE = ""
                RA_NO = ""
                blnReturnsHaveBeenApplied = False
                Dim test As String = Absx1.txtFor("APPROVED_BY").Text
                If test.Length > 0 And Absx1.txtFor("RA_NO").Text = "" Then
                    View_RMA()
                End If
                If Absx1.txtFor("RA_NO").Text = "" Then
                    EMsg &= vbCr & "No Returns Authorization No Specified"
                Else
                    RA_NO = Absx1.txtFor("RA_NO").Text
                    Dim row As DataRow = LookUp("SOTRMAF1", RA_NO)
                    If row Is Nothing Then
                        EMsg &= vbCr & "No Record of Returns Authorization No " & RA_NO
                    Else
                        CUST_CODE = row.Item("CUST_CODE")

                        If eItemKey = "Edit" Then
                            If row.Item("RA_STATUS") & "" <> "O" Then
                                Dim msg As String = ""
                                Select Case row.Item("RA_STATUS")
                                    Case "C"
                                        msg = "Returns Authorization No " & RA_NO & " has been Cancelled"
                                    Case "D"
                                        msg = "Returns Authorization No " & RA_NO & " has been Deleted"
                                    Case Else ' such as "F"
                                        msg = "Returns Authorization No " & RA_NO & " is No Longer Open"
                                End Select

                                EMsg &= vbCr & msg 'no reversal until we understand if an RA is used to generate 1 or multiple Credit memos
                                'If row.Item("RA_REASON_CODE") & "" <> "X" Then
                                '    If MsgBox("Do you want to Re-Open it for Processing", _
                                '                       MsgBoxStyle.YesNo, _
                                '                      msg) = MsgBoxResult.Yes Then
                                '        ASCMAIN1.sql = "Update SOTRMAF1 Set RA_STATUS = 'O', RA_DATE_CLOSED = NULL" _
                                '            & " where RA_NO = '" & RA_NO & "'"
                                '        ASCDATA1.ExecuteSQL()
                                '        row = Fill_Record("SOTRMAF1", RA_NO)
                                '    Else
                                '        EMsg &= vbCr & msg
                                '    End If
                                'Else
                                '    ASCMAIN1.sql = "" _
                                '        & "Select 'O' SOURCE, ARTOPEN1.* from ARTOPEN1 where INV_TYPE = 'R' and ORDR_NO = '" & RA_NO & "'" _
                                '        & " union " _
                                '        & "Select 'X' SOURCE, ARTOPEN1.* from ARTOPENX where INV_TYPE = 'R' and ORDR_NO = '" & RA_NO & "'"
                                '    For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                                '        If row.Item("OPS_YYYYPP") & "" <> ASCMAIN1.CYP Then
                                '            EMsg &= vbCr & msg & vbCr & "RA No " & RA_NO & " has generated a Credit in a Prior Period (see Credit Memo " _
                                '           & row.Item("INV_NUM") & ")" & vbCr & "It may not be re-opened for processing"
                                '            Exit For
                                '        End If

                                '        If row.Item("SOURCE") & "" = "X" Then
                                '            EMsg &= vbCr & msg & vbCr & "Some Credits originating from this RA are No Longer on File; Cannot Reverse"
                                '            Exit For
                                '        Else
                                '            If row.Item("INV_LAST_PMT") & "" <> "" Or Val(row.Item("INV_PMT") & "") <> 0 Or _
                                '               Val(row.Item("INV_TOTAL_AMOUNT") & "") <> Val(row.Item("INV_BALANCE") & "") Or _
                                '               row.Item("INV_LAST_PMT_REF") & "" <> "" Or _
                                '               row.Item("INV_LAST_PMT_REF_DT") & "" <> "" Then
                                '                EMsg &= vbCr & msg & vbCr & "Some Credits originating from this RA have been used in Application; Cannot Reverse"
                                '            End If
                                '            Exit For
                                '        End If
                                '    Next
                                '    If EMsg = "" Then
                                '        If MsgBox("Do you want to Reverse them to Edit the RA", _
                                '                  MsgBoxStyle.YesNo, "RA No " & RA_NO & " has already been used to Generate Credits") = MsgBoxResult.YES Then
                                '            Reverse_Credit_Memo()
                                '        Else
                                '            EMsg &= vbCr & msg
                                '        End If
                                '    End If
                                'End If
                            Else
                                ' See if there are any Returns against the RMA, if so, you cannot edit
                                ' 01/07/2019 - As per Lauren
                                If ASCMAIN1.CLIENT <> "INT" Then
                                    ASCMAIN1.sql = "select * from sotrtrn1 where RA_NO = '" & RA_NO & "'"

                                    ' This new query allows for reversals
                                    ASCMAIN1.sql = "select item_code, sum(rtrn_qty) " _
                                        & " from sotrtrn2" _
                                        & " where rtrn_no in (select rtrn_no from sotrtrn1 where RA_NO = '" & RA_NO & "') " _
                                        & " group by item_code" _
                                        & " having  sum(rtrn_qty) > 0"

                                    If ASCDATA1.GetDataTable(ASCMAIN1.sql).Rows.Count > 0 Then
                                        MsgBox("Returns Authorization No " & RA_NO & " has Returns applied to it." _
                                               & vbCrLf & vbCrLf & "Editing will not be permitted, but you may Cancel remaining balance.",
                                               MsgBoxStyle.OkOnly, "Editing Not Permitted if Returns have been applied")
                                        blnReturnsHaveBeenApplied = True
                                        'EMsg = "Returns Authorization No " & RA_NO & " has Returns applied to it."
                                    End If
                                End If
                            End If

                            If EMsg.Length = 0 Then
                                If Not ASCMAIN1.Logical_Lock("SOTRMAF1", RA_NO) Then Exit Sub
                                If Not ASCMAIN1.Logical_Lock("SOTRMAF1", CUST_CODE) Then Exit Sub
                            End If
                        End If

                    End If
                End If

                If EMsg <> "" Then ASCMAIN1.MultiTask_Release()

            Case "Update"
                If Absx1.dteFor("RA_DATE").Value & "" = "" _
                    Or Absx1.dteFor("RA_EXPIRE").Value & "" = "" Then
                    EMsg &= vbCr & "RA Date and Expiration Date are Mandatory"
                Else
                    If Format(Absx1.dteFor("RA_DATE").Value, "yyyyMMdd") _
                     > Format(Absx1.dteFor("RA_EXPIRE").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Expiration Date cannot be Prior to RA Date"
                    End If
                End If
                If ASCMAIN1.CLIENT = "AHA" And Absx1.dteFor("RA_START_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "RA Start Date is Mandatory"
                End If

                If Absx1.txtFor("CUST_STORE_NO").Text = "" Then
                    ' NO BIGGIE
                Else
                    If LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, Absx1.txtFor("CUST_STORE_NO").Text}) Is Nothing Then
                        EMsg &= vbCr & "Invalid Store No"
                    End If
                End If

                If Absx1.optFor("RA_REASON_CODE").Value = "" Then
                    EMsg &= vbCr & "RA Reason is required"
                Else
                    If LookUp("ARTREASR", Absx1.optFor("RA_REASON_CODE").Value) Is Nothing Then
                        EMsg &= vbCr & "Invalid RA Reason"
                    Else
                        If Absx1.optFor("RA_REASON_CODE").Value = "X" Then

                            'If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                            '    If ASCMAIN1.Running_in_VS Then
                            '        Stop
                            '    Else
                            '        EMsg &= vbCr & "Do Not Enter DIF until after we go live"
                            '    End If
                            'End If

                            'Dim dt() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
                            'Dim RA As String = Format(Absx1.dteFor("RA_DATE").Value, "yyyyMMdd")
                            'If RA < Format(dt(1), "yyyyMMdd") _
                            'Or RA > Format(dt(dt.Length - 1), "yyyyMMdd") Then
                            '    EMsg &= vbCr & "Invalid Date for RA Credit (must be " & Format(dt(1), "MM/dd/yy") & " thru " & Format(dt(dt.Length - 1), "MM/dd/yy") & ")"
                            'End If

                            Dim DT As Date = Absx1.dteFor("RA_DATE").Value
                            If DT & "" = "" Then
                                EMsg &= vbCr & "Document Date is Mandatory"
                            Else
                                Dim EMsg2 As String = ""
                                TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg2)

                                If blnAutoPilot180 Or blnAutoPilot812 Then
                                    ' LET IT GO WITH THE DATE IN THE 180

                                    'Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
                                    'Dim DTE2USE As Date = rowGLTPARM2.Item("PRD_END_DATE")
                                    'If Format(Now.Date, "yyyyMMdd") < Format(DTE2USE, "yyyyMMdd") Then
                                    '    DTE2USE = Now.Date
                                    'End If
                                    'If MsgBox("Would you like to use " & Format(DTE2USE, "MM/dd/yyyy"), _
                                    '          MsgBoxStyle.YesNo, _
                                    '          "Cannot Post Credit with Date Indicated in 180") = MsgBoxResult.Yes Then
                                    '    Absx1.dteFor("RA_DATE").Value = DTE2USE
                                    '    Absx1.dteFor("RA_EXPIRE").Value = DTE2USE
                                    'Else
                                    '    EMsg &= EMsg2
                                    'End If
                                Else
                                    EMsg &= EMsg2
                                End If
                            End If

                            If EMsg = "" Then
                                If Absx1.txtFor("CUST_CLAIM_NO").Text = "" Then
                                    'If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                                    '    'sp says not to insist on this 11/17/15
                                    ' I think it is important for DIFs
                                    'Else
                                    EMsg &= vbCr & "Customer Claim No Required for Reason Code Indicated"
                                    'End If
                                Else
                                    ' LBM ASKED TO REMOVE THIS ON 02/21/17
                                    'ASCMAIN1.sql = "" _
                                    '    & "Select * from ARTOPEN1 where CUST_CODE = '" & CUST_BILL_TO_CUST & "'" _
                                    '    & " and INV_TYPE = 'C' and INV_CUST_PO = '" & Absx1.txtFor("CUST_CLAIM_NO").Text & "'" _
                                    '    & " union " _
                                    '    & "Select * from ARTOPENX where CUST_CODE = '" & CUST_BILL_TO_CUST & "'" _
                                    '    & " and INV_TYPE = 'C' and INV_CUST_PO = '" & Absx1.txtFor("CUST_CLAIM_NO").Text & "'"
                                    'Dim tbl As DataTable = ASCDATA1.GetDataTable

                                    'If tbl.Rows.Count <> 0 Then
                                    '    Dim row As DataRow = tbl.Select("", "INV_DATE DESC")(0)
                                    '    If MsgBox("Previously Generated Credit Memo " & row.Item("INV_NUM") & " dated " & Format(row.Item("INV_DATE"), "MM/dd/yyyy") & vbCrLf _
                                    '              & vbCrLf & "Proceed w/Generation of Credit Anyway?", _
                                    '              MsgBoxStyle.YesNo, _
                                    '              "Customer Claim No has already been used on an existing Credit") = MsgBoxResult.No Then
                                    '        Exit Sub
                                    '    End If
                                    'End If
                                End If
                            End If
                        End If
                    End If
                End If

                ' RELAXING THIS CONSTRAINT FOR NOW - CONVERTED DATA HAD NO CLAIM
                'If Absx1.txtFor("CUST_CLAIM_NO").Text = "" Then
                '    EMsg &= vbCr & "Customer Claim No is required"
                'End If

                If optB.Value = "X" Then
                    If Not blnAutoPilot180 Then
                        ASCMAIN1.sql = "Select * from EDTTRPM1" & vbCrLf _
                            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                            & "   and EDI_DOC_NO = '180'"
                        Dim row As DataRow = ASCDATA1.GetDataRow
                        If row IsNot Nothing Then
                            EMsg &= vbCr & "Cannot Enter an RA for Customers who transmit EDI Document 180s"
                        End If
                    End If
                    If Not blnAutoPilot812 Then
                        ASCMAIN1.sql = "Select * from EDTTRPM1" & vbCrLf _
                            & " where CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                            & "   and EDI_DOC_NO = '812'"
                        Dim row As DataRow = ASCDATA1.GetDataRow
                        If row IsNot Nothing Then
                            EMsg &= vbCr & "Cannot Enter an RA for Customers who transmit EDI Document 812s"
                        End If
                    End If
                End If

                If grdSOTRMAF2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Returns Authorization"
                Else
                    If Val(dst.Tables("SOTRMAF2").Compute("COUNT(RA_LNO)", "RA_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Items on Returns Authorization with Qty >0"
                    Else
                        ' check for no charge with price, or for saleable without price
                        Dim items_with_price_issues_S As String = ""
                        Dim items_with_price_issues_NON_S As String = ""
                        For Each rowSOTRMAF2 As DataRow In dst.Tables("SOTRMAF2").Select("(ITEM_SNU_CODE = 'S' and RA_NET_PRICE = 0) or (ITEM_SNU_CODE <> 'S' and RA_NET_PRICE <> 0)")
                            Dim ITEM_CODE As String = rowSOTRMAF2.Item("ITEM_CODE") & ""
                            Dim ITEM_SNU_CODE As String = rowSOTRMAF2.Item("ITEM_SNU_CODE") & ""
                            If ITEM_SNU_CODE = "S" Then
                                items_with_price_issues_S &= "," & ITEM_CODE
                            Else
                                items_with_price_issues_NON_S &= "," & ITEM_CODE
                            End If

                        Next

                        If items_with_price_issues_NON_S <> "" Or items_with_price_issues_S <> "" Then
                            If MsgBox("There are Items with Price anomalies." & vbCrLf _
                                      & vbCrLf & "Saleable with 0 price: " & Mid(items_with_price_issues_S, 2) _
                                      & vbCrLf & "No-Charge with non-0 price: " & Mid(items_with_price_issues_NON_S, 2) _
                                      & vbCrLf & vbCrLf & "OK to Update anyway?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                EMsg &= vbCr & "Please correct the Prices and then try Update again"
                            End If
                        End If
                    End If
                End If


                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    'SP DOES NOT CARE ABOUT DIV
                Else
                    If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
                        EMsg &= vbCr & "Sales Division is required"
                    Else
                        Validate_Code("SALES_DIVISION_CODE")
                    End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTRMAF2 where RA_QTY_USED <> 0"
                ASCMAIN1.sql &= " and RA_NO = '" & RA_NO & "'"

                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Returns Authorization has been Used - Delete not permitted"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Returns Authorization as Deleted",
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Cancel Balance"
                If EMsg = "" Then
                    If MsgBox("Do you want to Cancel (the remaining open balance on) this Returns Authorization",
                               vbYesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If


            Case "Reverse"

                If Not ASCMAIN1.Logical_Lock("SOTRMAF1", RA_NO) Then Exit Sub

                If rowSOTRMAF1.Item("RA_REASON_CODE") <> "X" Then EMsg &= vbCr & "Invalid RA Type for this operation"
                If rowSOTRMAF1.Item("INV_NUM_REV") & "" <> "" Then EMsg &= vbCr & "This Credit has already been reversed"

                If EMsg = "" Then
                    If MsgBox("This option will Reverse the Credit which was Generated by this RA." _
                              & vbCrLf & vbCrLf & "This action is permanent" _
                              & vbCrLf & " - an inverse Credit Memo will be generated" _
                              & vbCrLf & " - you will need to apply this inverse Credit" _
                              & vbCrLf & "   (perhaps against the original Credit)" _
                              & vbCrLf & "   in Payment Application." _
                              & vbCrLf & vbCrLf & "OK to Proceed with the Reversal of Credit Memo " & rowSOTRMAF1.Item("INV_NUM") & "?",
                              MsgBoxStyle.YesNo,
                              "Option to Reverse Credit Generated by Destroyed RA") = MsgBoxResult.No Then
                        EMsg &= vbCr & "RA DIF Credit " & rowSOTRMAF1.Item("INV_NUM") & " has NOT been Reversed."
                    End If
                End If

                If EMsg <> "" Then ASCMAIN1.MultiTask_Release()

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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Order()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)
                RMAno.Visible = True
                UltraLabel7.Visible = True
            Case "Cancel Balance"
                Cancel_Order()
                Mode_Settings(False)

            Case "Print Credit Memo"
                Print_Credit_Memo()

            Case "Reverse"
                Reverse_Credit()
                Mode_Settings(False)

            Case "Show EDI"
                If rowSOTRMAF1 Is Nothing Then Exit Sub

                Dim EDI_DOC_SEQ_NO As String = rowSOTRMAF1.Item("EDI_DOC_SEQ_NO") & ""
                Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, , "180")
                If RAW_EDI = "" Then
                    RAW_EDI = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, , "812")
                End If
                Using frm As New ASFTEXT1
                    frm.t = RAW_EDI
                    frm.Text = "Raw EDI for " & CUST_CODE & " Claim No " & rowSOTRMAF1.Item("CUST_CLAIM_NO")
                    frm.ShowDialog()
                End Using

            Case "Refresh"
                Fill_Records("EDT180TX")
                Sort_grdColumns(grdEDT180TX, "EDI_DOC_SEQ_NO")
                Fill_Records("EDT812TX")
                Sort_grdColumns(grdEDT812TX, "EDI_DOC_SEQ_NO")
                Fill_Records("SOTRMAFX")
                Fill_Records("SOTRMAFI")
                Fill_Records("EDT180X0")
                Fill_Records("EDT812X0")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowSOTRMAF1.Item("RA_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If
                .Items("Update").Settings.Enabled = iScreenMode

                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode
                .Items("Print").Settings.Enabled = iScreenMode

                .Items("Cancel Balance").Settings.Enabled = iScreenMode


                .Items("New").Visible = Not InquiryMode
                .Items("Edit").Visible = Not InquiryMode

                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                .Items("Print").Visible = (EntryMode = "V" And ScreenMode) ' False ' ScreenMode
                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode) And Not blnReturnsHaveBeenApplied
                .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E") And Not blnReturnsHaveBeenApplied
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Refresh").Visible = Not InquiryMode And Not ScreenMode

                .Items("Cancel Balance").Visible = (EntryMode = "E")
                .Items("Print Credit Memo").Visible = (InquiryMode OrElse EntryMode = "V" OrElse ScreenMode) _
                            AndAlso (rowSOTRMAF1 IsNot Nothing _
                                     AndAlso rowSOTRMAF1.Item("RA_STATUS") & String.Empty = "F" _
                                     AndAlso rowSOTRMAF1.Item("RA_REASON_CODE") & String.Empty = "X" _
                                     AndAlso rowSOTRMAF1.Item("INV_TYPE") & String.Empty <> String.Empty _
                                     AndAlso rowSOTRMAF1.Item("INV_NUM") & String.Empty <> String.Empty)

                .Items("Show EDI").Visible = .Items("Print Credit Memo").Visible AndAlso rowSOTRMAF1.Item("EDI_DOC_NO") & "" <> ""
                If .Items("Show EDI").Visible Then
                    .Items("Show EDI").Text = "Show " & rowSOTRMAF1.Item("EDI_DOC_NO")
                End If

                .Items("Reverse").Visible = Not InquiryMode And (EntryMode = "V") _
                    AndAlso (rowSOTRMAF1.Item("RA_REASON_CODE") & "" = "X" And rowSOTRMAF1.Item("INV_NUM_REV") & "" = "")

            End With

            .Groups("Totals").Visible = ScreenMode
            .Groups("Status").Visible = Not ScreenMode And InquiryMode
            Hide_DIF_Group()
        End With

        tabRA.Tabs("Returns").Visible = (EntryMode = "V" And ScreenMode)

        lblStatus.Visible = ScreenMode

        grdSOTRMAFX.Visible = Not tf
        tabRAMaster.Visible = Not tf

        lblCredit.Visible = ScreenMode AndAlso (rowSOTRMAF1.Item("INV_NUM") & "" <> "")
        lblINV_NUM.Visible = ScreenMode AndAlso (rowSOTRMAF1.Item("INV_NUM") & "" <> "")
        'If ScreenMode Then
        '    lblINV_NUM.Text = ""
        'End If
        lblReversed.Visible = ScreenMode AndAlso (rowSOTRMAF1.Item("INV_NUM_REV") & "" <> "")
        lblINV_NUM_REV.Visible = ScreenMode AndAlso (rowSOTRMAF1.Item("INV_NUM_REV") & "" <> "")

        Absx1.optFor("RA_REASON_CODE").Visible = ScreenMode

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("CUST_CLAIM_NO"), InquiryMode Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")) Or blnReturnsHaveBeenApplied)
        Dim RA_was_used As Boolean = (dst.Tables("SOTRMAF2").Select("RA_QTY_USED <> 0").Length > 0)
        ' isnt RA_was_used doing the same thing as blnReturnsHaveBeenApplied?
        Set_Read_Only_for_ctl(Absx1.optFor("RA_REASON_CODE"), InquiryMode Or RA_was_used Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")) Or blnReturnsHaveBeenApplied)
        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N") Or blnReturnsHaveBeenApplied)
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N") Or blnReturnsHaveBeenApplied)

        grdSOTRMAF2.DisplayLayout.Bands(0).Columns("X").Hidden = Not (ScreenMode And (EntryMode = "E" And Not blnReturnsHaveBeenApplied))
        grdSOTRMAF2.DisplayLayout.Bands(0).Columns("RA_LINE_AMT").Hidden = True ' THIS COL IS FOR A DIFFERENT DESIGN

        If ScreenMode Then

            If EntryMode = "V" Or blnReturnsHaveBeenApplied Then
                grdSOTRMAF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTRMAF2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTRMAF2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdSOTRMAF2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdSOTRMAF2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTRMAF2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

                If EntryMode <> "E" Then
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("X").Hidden = True
                Else
                    grdSOTRMAF2.DisplayLayout.Bands(0).Columns("X").Hidden = False
                End If
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("RA_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = CUST_CODE

        CUST_CODE = ""
        RA_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTRMAF1", "SOTRMAF2", "ARTOPEN1", "SOTINVH1", "EDT180X2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_SOTRMAFX()
        Load_EDTXXXT1("180")
        Load_EDTXXXT1("812")

        Fill_Records("EDT180X0")
        Fill_Records("EDT812X0")
        blnReturnsHaveBeenApplied = False
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then

            Do Until RA_NO <> ""
                RA_NO = ASCMAIN1.Next_Control_No("SOTRMAF1.RA_NO")
                Dim rowcheck As DataRow = LookUp("SOTRMAF1", RA_NO)
                If rowcheck IsNot Nothing Then
                    RA_NO = ""
                Else
                    Dim tblcheck As DataTable = ASCDATA1.GetDataTable($"SELECT * FROM SOTRMAF2 WHERE RA_NO = {RA_NO}")
                    If tblcheck.Rows.Count <> 0 Then
                        RA_NO = ""
                    End If
                End If
            Loop


            rowSOTRMAF1 = dst.Tables("SOTRMAF1").NewRow
            With rowSOTRMAF1
                .Item("RA_NO") = RA_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_CLAIM_NO") = CUST_CLAIM_NO
                .Item("RA_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                    .Item("RA_REASON_CODE") = "O"  ' SP DEFAULT
                End If
                Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
                If WHSE_CODE = "" Then WHSE_CODE = ""
                .Item("WHSE_CODE") = WHSE_CODE
                '  .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
                .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & ""
                .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE") & ""
            End With
            dst.Tables("SOTRMAF1").Rows.Add(rowSOTRMAF1)
        Else

            rowSOTRMAF1 = Fill_Record("SOTRMAF1", RA_NO)
        End If

        ' Load screen with items found on Invoices for the given Customer, Customer PO
        Fill_Records("SOTINVH2X", New Object() {CUST_CODE, CUST_CLAIM_NO})

        If EntryMode = "N" AndAlso dst.Tables("SOTINVH2X").Rows.Count > 0 Then
            Dim minCUST_STORE_NO As String = dst.Tables("SOTINVH2X").Compute("MIN(CUST_STORE_NO)", "") & String.Empty
            Dim maxCUST_STORE_NO As String = dst.Tables("SOTINVH2X").Compute("MAX(CUST_STORE_NO)", "") & String.Empty
            If minCUST_STORE_NO.Length > 0 AndAlso minCUST_STORE_NO = maxCUST_STORE_NO Then
                MyBase.Absx1.txtFor("CUST_STORE_NO").Text = minCUST_STORE_NO
            End If
        End If

        CUST_CODE = rowSOTRMAF1.Item("CUST_CODE")
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
        If CUST_BILL_TO_CUST = "" Then CUST_BILL_TO_CUST = CUST_CODE
        rowARTCUST1_BT = LookUp("ARTCUST1", CUST_BILL_TO_CUST)

        PRICE_CLASS_CODE = rowARTCUST1.Item("PRICE_CLASS_CODE")
        Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
        PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
        PRICE_BASIS = rowSOTPCLS1.Item("PRICE_BASIS") & ""
        PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & ""


        CUST_CODE_ALLO = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
        PRICE_LIST_CODE_ALLO = ""
        If CUST_CODE_ALLO <> "" Then
            Dim rowARTCUST1_ALLO As DataRow = LookUp("ARTCUST1", CUST_CODE_ALLO)
            If rowARTCUST1_ALLO IsNot Nothing Then
                PRICE_LIST_CODE_ALLO = rowARTCUST1_ALLO.Item("PRICE_LIST_CODE") & ""
            End If
        End If


        Setup_Price_Class()

        Fill_Records("SOTRMAF2", RA_NO)
        Sort_grdColumns(grdSOTRMAF2, "RA_LNO")

        lblINIT_DATE.Text = "RA Entered " & Format(rowSOTRMAF1.Item("INIT_DATE"), "MM/dd/yyyy") & " by " & rowSOTRMAF1.Item("INIT_OPER")

        If EntryMode = "N" Then
            lblStatus.Text = "New"
        Else
            Select Case rowSOTRMAF1.Item("RA_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
                Case "F"
                    lblStatus.Text = "Completed"
                Case Else
                    lblStatus.Text = "?"
            End Select
        End If

        With grdSOTRMAF2.DisplayLayout.Bands(0)
            If (EntryMode = "E" Or EntryMode = "N") Then
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                If EntryMode = "E" Then
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Else
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdSOTRMAF2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            End With
            grdSOTRMAF2.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, False)
        Else
            With grdSOTRMAF2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            grdSOTRMAF2.DisplayLayout.Bands(0).Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, True)
        End If

        If EntryMode = "V" Then
            Fill_Records("SOTRTRNX", RA_NO)
            grdSOTRTRNX.Text = "Returns Entered Referencing RA " & RA_NO
        End If

        If ASCMAIN1.CLIENT = "INT" Then
            Load_Retail_On_Hand()
        End If

        Display_Totals()
        EnforceConstraints(True)
        RMAno.Visible = False
        UltraLabel7.Visible = False
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, RA_NO)
        For Each TABLE_NAME As String In New String() _
            {"SOTRMAF1", "SOTRMAF2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where RA_NO = '" & RA_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        If EntryMode <> "N" Then Delete_Records()

        rowSOTRMAF1.Item("RA_AMT") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT)", "") & "")

        'If rowSOTRMAF1.Item("CUST_STORE_NO") & "" <> "" Then
        '    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {rowSOTRMAF1.Item("CUST_CODE"), rowSOTRMAF1.Item("CUST_STORE_NO")})
        '    If rowARTCUST2 IsNot Nothing Then
        '        rowSOTRMAF1.Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE")
        '    End If
        'End If

        If Absx1.optFor("RA_REASON_CODE").Value = "X" Then
            rowSOTRMAF1.Item("RA_STATUS") = "F"
            rowSOTRMAF1.Item("RA_DATE_CLOSED") = DATETIME_STAMP.Date
            For Each rowSOTRMAF2 As DataRow In dst.Tables("SOTRMAF2").Select("")
                rowSOTRMAF2.Item("RA_QTY_USED") = rowSOTRMAF2.Item("RA_QTY")
                rowSOTRMAF2.Item("RA_QTY_OPEN") = 0
            Next
            Record_AR_Item()
        End If

        INIT_LAST("SOTRMAF1", False, , True)
        Dim sqldelete As String = "RA_NO = '" & RA_NO & "'"
        Update_Record_TDA("SOTRMAF1", sqldelete)
        Update_Record_TDA("SOTRMAF2", sqldelete)
        Dependent_Updates(1, RA_NO)

        Dim msg As String = "Update Complete"
        If blnAutoPilot180 Then msg = ""
        If blnAutoPilot812 Then msg = ""
        CommitTrans(msg)

        ' See if we need to issue credit card credit.
        If dst.Tables("SOTINVH1").Rows.Count > 0 Then
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows(0)
            If rowSOTINVH1.Item("CC_SALE_TRANS_ID") & String.Empty <> String.Empty AndAlso Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty) <> 0 Then
                ASCMAIN1.Progress("Processing CC Credit", "")
                Dim errorMessage As String = String.Empty
                If Not SOCMAIN1.IssueCredit(rowSOTINVH1.Item("INV_NO"), errorMessage) Then
                    MessageBox.Show("Error Processing Credit Card Refund: " & errorMessage, "CC Credit", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                ASCMAIN1.Progress("", "")
            End If
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Reverse_Credit()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        rowSOTRMAF1.Item("RA_AMT") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT)", "") & "")

        Record_AR_Item(True)

        INIT_LAST("SOTRMAF1", False, , True)
        Update_Record_TDA("SOTRMAF1")

        CommitTrans("Reversal Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "RA_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("CUST_CLAIM_NO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a Claim No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTRMAF1.RA_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTRMAF1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("CUST_CLAIM_NO").Text <> "" Then
                    sql_where &= " and SOTRMAF1.CUST_CLAIM_NO = '" & Absx1.txtFor("CUST_CLAIM_NO").Text & "'"
                End If

            Case "CUST_STORE_NO"
                sql_where &= " and CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"

        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")
                RMAno.Visible = True
                UltraLabel7.Visible = True
            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("RA_NO").Text = key
                Click_Command("View")
                RMAno.Visible = False
                UltraLabel7.Visible = False
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTRMAF1"
            E.COLUMN_NAME = "RA_NO"
            E.CODE_VALUE = Absx1.txtFor("RA_NO").Text
            E.DESC_VALUE = "Returns Authorization"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRMAF1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTRMAFX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTRMAF2, "BBB", "Item Status Inquiry", "Load Retail On Hand", "Copy Retail On Hand to Qty Open", "Copy Retail On Hand to RA Qty", "Paste Items", "Load Store Items")
        Load_Popup_Menu(grdSOTRTRNX, "SS", "Show Filter", "Show GroupBox", "Show Return")
        Load_Popup_Menu(grdSOTRMAFI, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdEDT180TX, "SSSBBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Show Raw EDI", "Edit EDI Details", "Select All", "Select All for CUST_CODE", "De-Select All", "Report Selected", "Process Selected", "Delete Selected")
        Load_Popup_Menu(grdEDT812TX, "SSSBBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Show Raw EDI", "Edit EDI Details", "Select All", "Select All for CUST_CODE", "De-Select All", "Report Selected", "Process Selected", "Delete Selected")
        Load_Popup_Menu(grdEDTXXXX0, "SSPB", "Show Filter", "Show GroupBox", "Print Report")
        Load_Popup_Menu(grdEDTXXXX1, "SSPB", "Show Filter", "Show GroupBox", "RA Inquiry")
        Load_Popup_Menu(grdEDT180T2, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then grd = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdSOTRMAF2"
                tlb_btn = DirectCast(tlb_pop.Tools("Item Status Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "INT") And ScreenMode And grdSOTRMAF2.Rows.Count > 0

                tlb_btn = DirectCast(tlb_pop.Tools("Load Retail On Hand"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "INT") And (EntryMode = "N" Or EntryMode = "E") And grdSOTRMAF2.Rows.Count > 0

                tlb_btn = DirectCast(tlb_pop.Tools("Copy Retail On Hand to Qty Open"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "INT") And (EntryMode = "N" Or EntryMode = "E") And grdSOTRMAF2.Rows.Count > 0
                tlb_btn = DirectCast(tlb_pop.Tools("Copy Retail On Hand to RA Qty"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "INT") And (EntryMode = "N" Or EntryMode = "E") And grdSOTRMAF2.Rows.Count > 0

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Items"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "INT") And (EntryMode = "N" Or EntryMode = "E")
                tlb_btn = DirectCast(tlb_pop.Tools("Load Store Items"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (ASCMAIN1.CLIENT = "INT") And (EntryMode = "N" Or EntryMode = "E") And Trim(Absx1.txtFor("CUST_STORE_NO").Text) <> ""

            Case "grdEDT180TX", "grdEDT812TX"
                tlb_btn = DirectCast(tlb_pop.Tools("Process Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not InquiryMode
                tlb_btn = DirectCast(tlb_pop.Tools("Delete Selected"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not InquiryMode

                tlb_btn = DirectCast(tlb_pop.Tools("Select All for CUST_CODE"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Visible = True
                    tlb_btn.SharedProps.Caption = "Select All for " & grd.ActiveRow.Cells("CUST_CODE").Value
                    tlb_btn.Tag = grd.ActiveRow.Cells("CUST_CODE").Value
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Edit EDI Details"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    Dim reqRes As String = grd.ActiveRow.Cells("REQUIRES_RESOLUTION").Value & ""
                    tlb_btn.SharedProps.Visible = reqRes = "1"
                End If
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            If grd.Name <> "grdSOTRMAF2" Then e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Select All for CUST_CODE"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        If grow.Cells("CUST_CODE").Value = e.Tool.Tag & "" Then
                            grow.Cells("SEL").Value = "1"
                            grow.Update()
                        End If
                    End If
                Next

            Case "Print Report", "Report Selected"
                Dim EDTXXXTX As String = "EDT180TX"
                Dim EDTXXXT2 As String = "EDT180T2"
                If e.Tool.Key = "Print Report" Then
                    Dim XNOs As String = String.Join("','", dst.Tables("EDT180X0").AsEnumerable().Where(Function(row) row.Item("SEL") = "1").Select(Of String)(Function(row) row.Item("XNO").ToString))
                    ASCMAIN1.sql = $"Select EDT180T1.*, '1' SEL from EDT180T1 WHERE XNO IN ('{XNOs}')"
                    Fill_Records("EDT180TX", Temp_Select:=ASCMAIN1.sql)
                End If

                If grd.Name = "grdEDT812TX" Or (grd.Name = "grdEDTXXXX0" And optDifReportType.Value = "812") Then
                    EDTXXXTX = "EDT812TX"
                    EDTXXXT2 = "EDT812T2"
                End If

                If e.Tool.Key = "Print Report" Then
                    Dim editype As String = optDifReportType.Value
                    Dim XNOs As String = String.Join("','", dst.Tables("EDT180X0").AsEnumerable().Where(Function(row) row.Item("SEL") = "1").Select(Of String)(Function(row) row.Item("XNO").ToString))
                    ASCMAIN1.sql = $"Select EDT{editype}T1.*, '1' SEL from EDT{editype}T1 WHERE XNO IN ('{XNOs}')"
                    Fill_Records(EDTXXXTX, Temp_Select:=ASCMAIN1.sql)
                End If

                If dst.Tables(EDTXXXTX).Select("SEL='1'").Length = 0 And e.Tool.Key = "Report Selected" Then
                    MsgBox("No Records Selected", vbOKOnly, "Cannot Produce Report")
                Else
                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Loading Data")

                    dst.Tables(EDTXXXT2).Rows.Clear()
                    Dim CUST_CODE As String = ""
                    Dim PRICE_CLASS_CODE As String = ""
                    Dim PRICE_LIST_CODE As String = ""
                    Dim PRICE_LIST_CODE_ALLO As String = ""
                    Dim PRICE_BASIS As String = ""
                    Dim PRICE_BASE_DPCT As Decimal = 0
                    For Each row As DataRow In dst.Tables(EDTXXXTX).Select("SEL='1'", "CUST_CODE")
                        If row.Item("CUST_CODE") <> CUST_CODE Then
                            CUST_CODE = row.Item("CUST_CODE")
                            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                            PRICE_CLASS_CODE = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
                            Dim rowSOTPCLS1 As DataRow = LookUp("SOTPCLS1", PRICE_CLASS_CODE)
                            PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
                            PRICE_BASIS = rowSOTPCLS1.Item("PRICE_BASIS") & ""
                            PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & ""
                            Dim CUST_CODE_ALLO As String = rowARTCUST1.Item("CUST_CODE_ALLO") & ""
                            PRICE_LIST_CODE_ALLO = ""
                            If CUST_CODE_ALLO <> "" Then
                                Dim rowARTCUST1_ALLO As DataRow = LookUp("ARTCUST1", CUST_CODE_ALLO)
                                If rowARTCUST1_ALLO IsNot Nothing Then
                                    PRICE_LIST_CODE_ALLO = rowARTCUST1_ALLO.Item("PRICE_LIST_CODE") & ""
                                End If
                            End If
                        End If
                        Dim EDI_DOC_SEQ_NO = row.Item("EDI_DOC_SEQ_NO")

                        Dim rptDate As Date = Now.Date
                        If e.Tool.Key = "Print Report" Then
                            Dim RA_DATE As Date = ASCDATA1.GetDataValue("SELECT RA_DATE FROM SOTRMAF1 F1 WHERE F1.EDI_DOC_NO=:PARM1 AND F1.EDI_DOC_SEQ_NO=:PARM2", "VV", {optDifReportType.Value, row.Item("EDI_DOC_SEQ_NO")})
                            rptDate = RA_DATE
                        End If

                        Fill_Records(EDTXXXT2, EDI_DOC_SEQ_NO, False)
                        For Each row2 As DataRow In dst.Tables(EDTXXXT2) _
                            .Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' AND ITEM_CODE IS NOT NULL", "")
                            Dim ITEM_CODE As String = row2.Item("ITEM_CODE")
                            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                            Dim ITEM_RETAIL_PRICE As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                            Dim ORDR_UNIT_PRICE As Decimal = TAC.SOCMAIN1.Get_Price _
                                  (Me,
                                   PRICE_LIST_CODE,
                                   PRICE_LIST_CODE_ALLO,
                                   PRICE_BASIS,
                                   PRICE_BASE_DPCT,
                                   ITEM_CODE,
                                   rowICTITEM1,
                                   rptDate, ITEM_RETAIL_PRICE) ' MAYBE SHOULD USE RA_DATE - 60
                            row2.Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE
                            row2.Item("ITEM_RETAIL_PRICE") = ITEM_RETAIL_PRICE
                        Next
                    Next


                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Printing Report")

                    Print_Report_Begin()
                    Dim SUBT As String = ""

                    If grd.Name = "grdEDT180TX" Or (grd.Name = "grdEDTXXXX0" And optDifReportType.Value = "180") Then
                        Generate_Report("SORRMAFE", "Customer DIF-180 Report", SUBT, , , , True)
                    Else
                        Generate_Report("SORRMAFF", "Customer DIF-812 Report", SUBT, , , , True)
                    End If

                    Print_Report_End()

                    Me.Cursor = Cursors.Default
                    Click_Command("Refresh")
                    ASCMAIN1.Progress("")

                End If

            Case "Process Selected", "Delete Selected"
                Dim EDI As String = "180"
                Dim EDTXXXTX As String = "EDT180TX"
                Dim EDTXXXT1 As String = "EDT180T1"
                Dim EDTXXXT2 As String = "EDT180T2"
                If grd.Name = "grdEDT812TX" Then
                    EDI = "812"
                    EDTXXXTX = "EDT812TX"
                    EDTXXXT1 = "EDT812T1"
                    EDTXXXT2 = "EDT812T2"
                End If

                If dst.Tables(EDTXXXTX).Select("SEL='1'").Length = 0 Then
                    MsgBox("No Records Selected", vbOKOnly, "Cannot Process Records")
                Else

                    If MsgBox("Hi " & Split(ASCMAIN1.USER_NAME, " ")(0) & vbCrLf & vbCrLf & "You have Selected " & CStr(dst.Tables(EDTXXXTX).Select("SEL='1'").Length) & " Records to " & Split(e.Tool.Key, " ")(0) & vbCrLf & vbCrLf & "Do you want to Continue?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    Dim COUNTS(2) As Integer
                    COUNTS(0) = dst.Tables(EDTXXXTX).Select("SEL='1'").Length

                    Dim EDI_DOC_SEQ_NOs As New List(Of String)

                    For Each rowEDTXXXTX As DataRow In dst.Tables(EDTXXXTX).Select("SEL='1'", "CUST_CODE")
                        Dim EDI_DOC_SEQ_NO As String = rowEDTXXXTX.Item("EDI_DOC_SEQ_NO")
                        EDI_DOC_SEQ_NOs.Add(EDI_DOC_SEQ_NO)
                    Next

                    If EDI = "180" Then
                        blnAutoPilot180 = True
                    Else
                        blnAutoPilot812 = True
                    End If

                    Dim tbl As New DataTable
                    With tbl
                        .Columns.Add("EDI_DOC_SEQ_NO")
                        .Columns.Add("RA_NO")
                        .Columns.Add("RA_QTY", GetType(System.Int32))
                        .Columns.Add("QTY_180", GetType(System.Int32))
                        .Columns.Add("VARIANCES", GetType(System.Int32), "RA_QTY - QTY_180")
                    End With

                    Dim XNO As String = ""

                    If e.Tool.Key = "Process Selected" Then
                        XNO = ASCMAIN1.Next_Control_No($"{EDTXXXT1}.XNO")
                    End If

                    For Each EDI_DOC_SEQ_NO As String In EDI_DOC_SEQ_NOs
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Processing")

                        Dim processResult = Process_EDI(EDI, EDI_DOC_SEQ_NO)

                        If e.Tool.Key = "Delete Selected" Then
                            Click_Command("Cancel")
                            If Not ScreenMode Then
                                COUNTS(1) += 1
                                ASCMAIN1.sql = "Update " & EDTXXXT1 & " Set EDI_PROCESS_IND = 'D' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                                ASCDATA1.ExecuteSQL()
                            End If
                        Else
                            If Not processResult Then
                                Click_Command("Cancel")
                                ASCMAIN1.sql = "Update " & EDTXXXT1 & " Set REQUIRES_RESOLUTION = '1' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                                ASCDATA1.ExecuteSQL()
                            ElseIf dst.Tables("SOTRMAF2").Rows.Count = 0 Then
                                Click_Command("Cancel")
                                If Not ScreenMode Then
                                    COUNTS(2) += 1
                                    ASCMAIN1.sql = "Update " & EDTXXXT1 & " Set EDI_PROCESS_IND = 'Z' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                                    ASCDATA1.ExecuteSQL()
                                End If
                            Else
                                Click_Command("Update")


                                If Not ScreenMode Then
                                    COUNTS(1) += 1
                                    ASCMAIN1.sql = "Update " & EDTXXXT1 & " Set EDI_PROCESS_IND = '1', XNO = :PARM1 where EDI_DOC_SEQ_NO = :PARM2"
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", {XNO, EDI_DOC_SEQ_NO})

                                    ASCMAIN1.sql = $"SELECT SUM (RA_QTY) FROM SOTRMAF2 WHERE RA_NO= '{RA_NO_EDI}'"
                                    Dim RA_QTY As Integer = Val(ASCDATA1.GetDataValue)

                                    ASCMAIN1.sql = $"SELECT SUM (EDI_QTY) FROM EDT180T2 WHERE EDI_DOC_SEQ_NO= '{EDI_DOC_SEQ_NO}'"
                                    Dim QTY_180 As Integer = Val(ASCDATA1.GetDataValue)

                                    Dim row As DataRow = tbl.NewRow
                                    row.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                    row.Item("RA_NO") = RA_NO_EDI
                                    row.Item("RA_QTY") = RA_QTY
                                    row.Item("QTY_180") = QTY_180
                                    tbl.Rows.Add(row)
                                End If

                            End If
                        End If

                        Dim rowEDTXXXTX As DataRow = dst.Tables(EDTXXXTX).Rows.Find(EDI_DOC_SEQ_NO)
                        If rowEDTXXXTX IsNot Nothing Then rowEDTXXXTX.Delete()

                        If ScreenMode Then
                            MsgBox("Batch Process has been interrupted",
                                   MsgBoxStyle.OkOnly,
                                   "Exceptional Condition has been Encountered")
                            Exit For
                        End If

                    Next

                    If EDI = "180" Then
                        Dim rows() As DataRow = tbl.Select("RA_QTY <> QTY_180")
                        If rows.Length > 0 Then
                            MsgBox("Warning - RAs have been generated with a different qty than the originating 180")

                            Using FRM As New ASFMSGBF
                                Dim dvw As DataView = tbl.DefaultView
                                dvw.RowFilter = "RA_QTY <> QTY_180"
                                FRM.Show_grd(dvw.ToTable, Me, "180 converted to RA with Qty Discrepancies")
                            End Using

                        End If

                    End If

                    If EDI = "180" Then
                        blnAutoPilot180 = False
                    Else
                        blnAutoPilot812 = False
                    End If

                    If e.Tool.Key = "Delete Selected" Then
                        MsgBox("Records Selected = " & CStr(COUNTS(0)) _
                            & vbCrLf & "Records Deleted = " & CStr(COUNTS(1)), MsgBoxStyle.OkOnly, "Processing Complete")
                    Else
                        MsgBox("Records Selected = " & CStr(COUNTS(0)) _
                            & vbCrLf & "Records Updated = " & CStr(COUNTS(1)) _
                            & vbCrLf & "Records Zeroed = " & CStr(COUNTS(2)), MsgBoxStyle.OkOnly, "Processing Complete")
                    End If

                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")
                    Click_Command("Refresh")

                End If

            Case "Load Retail On Hand"

                Load_Retail_On_Hand()


            Case "Copy Retail On Hand to Qty Open", "Copy Retail On Hand to RA Qty"
                Dim grows() As UltraWinGrid.UltraGridRow
                ReDim grows(grdSOTRMAF2.Rows.Count)
                If grdSOTRMAF2.Selected.Rows.Count <> 0 Then
                    grdSOTRMAF2.Selected.Rows.CopyTo(grows, 0)
                Else
                    grdSOTRMAF2.Rows.CopyTo(grows, 0)
                End If

                Dim COL As String = "RA_QTY_OPEN"
                Dim COL_DESC As String = "Open RA Qty"
                If e.Tool.Key = "Copy Retail On Hand to RA Qty" Then
                    COL = "RA_QTY"
                    COL_DESC = "Total RA Qty"
                End If
                If MsgBox($"OK to copy the Qty@EOW to the {COL_DESC} field for " & IIf(grdSOTRMAF2.Selected.Rows.Count = 0, "All Rows?", "the " & grdSOTRMAF2.Selected.Rows.Count & " Selected Rows?"),
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                    For Each grow As UltraWinGrid.UltraGridRow In grows
                        If grow IsNot Nothing Then
                            Load_OH_into_RA(grow, COL)
                        End If
                    Next
                    grdSOTRMAF2.Selected.Rows.Clear()
                End If

            Case "Paste Items"
                Paste_Item_Codes_to_Add_to_RA()

            Case "Load Store Items"
                ASCMAIN1.Progress("Now Loading Items and EOW Qtys for Store " & Absx1.txtFor("CUST_STORE_NO").Text)
                grdSOTRMAF2.Visible = False
                Load_Retail_On_Hand(True)
                grdSOTRMAF2.Visible = True
                ASCMAIN1.Progress("")
                Sort_grdColumns(grdSOTRMAF2, "RA_LNO")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                If grd.ActiveRow.IsDataRow Then
                    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1 IsNot Nothing Then
                        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                    End If
                End If

            Case "Show Return"
                If grd.ActiveRow.IsDataRow Then
                    Dim RTRN_NO As String = grd.ActiveRow.Cells("RTRN_NO").Text
                    Dim rowSOTRTRN1 As DataRow = LookUp("SOTRTRN1", RTRN_NO)
                    If rowSOTRTRN1 IsNot Nothing Then
                        Context_Launch("View", RTRN_NO, e.Tool.Key, "SOFRTRN1")
                    End If
                End If

            Case "Show Raw EDI"
                Dim EDI As String = "180"
                If grd.Name = "grdEDT812TX" Then
                    EDI = "812"
                End If
                If grd.ActiveRow IsNot Nothing Then
                    Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                    Dim RAW_EDI As String = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, , EDI)
                    Using frm As New ASFTEXT1
                        frm.t = RAW_EDI
                        frm.Text = "Raw EDI for " & CUST_CODE & " RA No " & grd.ActiveRow.Cells("RMA_NUMBER").Value
                        frm.ShowDialog()
                    End Using
                End If

            Case "Edit EDI Details"
                Dim EDI As String = "180"
                If grd.Name = "grdEDT812TX" Then
                    EDI = "812"
                End If
                If grd.ActiveRow IsNot Nothing Then
                    Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & ""
                    Using frm As New SOFDIFX1(EDI, EDI_DOC_SEQ_NO)
                        frm.ShowDialog()
                    End Using
                End If

            Case "RA Inquiry"
                Dim RA_NO As String = grd.ActiveRow.Cells("RA_NO").Text
                Context_Launch("View", RA_NO, "Returns Authorization Inquiry", "SOFRMAFI")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTRMAFX()
                End If

            Case "CUST_CLAIM_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode _
                       And Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("CUST_CLAIM_NO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "RA_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If

            Case "APPROVED_BY"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    View_RMA()
                End If

            Case "RA_REASON_CODE"
                If Absx1.optFor("RA_REASON_CODE").Value = "X" And EntryMode = "N" Then
                    MsgBox("Credit is Issued Immediately when this Reason Code is Used",
                           MsgBoxStyle.OkOnly, "Please Note")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTRMAFX()

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTRMAFX()
            Case "RA_NO"
                Click_Command("View")
            Case "APPROVED_BY"
                View_RMA()
                'Click_Command("View")
        End Select
    End Sub
    Private Sub View_RMA()
        Dim rmaNumber As String = Absx1.txtFor("APPROVED_BY").Text
        Dim raNumbers As New List(Of String)
        ' Ensure that RMA Number is not empty
        If String.IsNullOrWhiteSpace(rmaNumber) Then
            MessageBox.Show("Please enter an RMA Number.")
            Return
        End If
        ' Query to get EDI_DOC_SEQ_NO
        ASCMAIN1.sql = $"SELECT EDI_DOC_SEQ_NO FROM EDT180T1 WHERE RMA_NUMBER LIKE " & "'%" & rmaNumber & "'"
        Dim edi_doc_seq_no As String = ASCDATA1.GetDataValue

        ' Check if EDI_DOC_SEQ_NO was retrieved
        If String.IsNullOrEmpty(edi_doc_seq_no) Then
            MessageBox.Show("EDI_DOC_SEQ_NO not found for the provided RMA Number.")
            Return
        End If

        ' Query to get RA_NO using EDI_DOC_SEQ_NO
        ASCMAIN1.sql = $"SELECT RA_NO FROM SOTRMAF1 WHERE EDI_DOC_SEQ_NO = " & "'" & edi_doc_seq_no & "'"
        Dim dataTable As DataTable = ASCDATA1.GetDataTable
        If dataTable IsNot Nothing AndAlso dataTable.Rows.Count > 0 Then
            ' Process each RA_NO
            For Each row As DataRow In dataTable.Rows
                Dim ra_no As String = row("RA_NO").ToString()
                raNumbers.Add(ra_no)
                'Absx1.txtFor("RA_NO").Text = e.Row.Cells("RA_NO").Value & String.Empty
                Absx1.txtFor("RA_NO").Text = ra_no & String.Empty
            Next
        Else
            MessageBox.Show("No RA_NO found for the provided EDI_DOC_SEQ_NO.")
        End If
        'raNumbers.Add("218098")
        If raNumbers.Count > 1 Then
            Dim raNumbersString As String = "'" & String.Join("', '", raNumbers) & "'"
            ASCMAIN1.sql = "SELECT SOTRMAF1.*, X.RA_QTY, X.RA_QTY_OPEN, X.RA_QTY_USED, X.RA_QTY_CANC, " &
                             "X.RA_AMT, X.RA_AMT_OPEN, X.RA_AMT_USED, X.RA_AMT_CANC, X.RA_RETAIL_EXT " &
                             "FROM (SELECT SOTRMAF2.RA_NO, Sum(SOTRMAF2.RA_QTY) RA_QTY, " &
                             "Sum(SOTRMAF2.RA_QTY_OPEN) RA_QTY_OPEN, Sum(SOTRMAF2.RA_QTY_USED) RA_QTY_USED, " &
                             "Sum(SOTRMAF2.RA_QTY_CANC) RA_QTY_CANC, " &
                             "Sum(NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT, " &
                             "Sum(NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_OPEN, " &
                             "Sum(NVL(SOTRMAF2.RA_QTY_USED,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_USED, " &
                             "Sum(NVL(SOTRMAF2.RA_QTY_CANC,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_CANC, " &
                             "Sum(NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_RETAIL,0)) RA_RETAIL_EXT " &
                             "FROM SOTRMAF2, SOTRMAF1 " &
                             "WHERE SOTRMAF2.RA_NO = SOTRMAF1.RA_NO AND SOTRMAF2.RA_NO IN (" & raNumbersString & ") " &
                             "GROUP BY SOTRMAF2.RA_NO) X, SOTRMAF1 " &
                             "WHERE X.RA_NO = SOTRMAF1.RA_NO AND SOTRMAF1.RA_NO IN (" & raNumbersString & ")"
            If SOTRMAFX = "" Then
                SOTRMAFX = ASCMAIN1.Temp_Table
                ASCDATA1.ExecuteSQL("Alter Table " & SOTRMAFX & " Add Primary Key (RA_NO)")
            Else
                ASCDATA1.ExecuteSQL("Truncate Table " & SOTRMAFX)
                ASCDATA1.ExecuteSQL("Insert into " & SOTRMAFX & " " & ASCMAIN1.sql)
            End If
            grdSOTRMAFX.Text = "Connected RAs"
            Fill_Records("SOTRMAFX")
            Sort_grdColumns(grdSOTRMAFX, "RA_NO".ToLower)
            grdSOTRMAFX.Visible = True
        End If
        'Click_Command("View")
    End Sub
    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTRMAFX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Data")

        Dim sqlw As String = ""
        If InquiryMode Then
            If optStatus.Value <> "A" Then
                sqlw &= " and SOTRMAF1.RA_STATUS = '" & optStatus.Value & "'"
                grdSOTRMAFI.Text = optStatus.Text & " by Item"
            End If
            If optStatus.Value = "F" Or optStatus.Value = "A" Then
                sqlw &= " and SOTRMAF1.RA_DATE >= '" & Format(dteRAFrom.Value, "dd-MMM-yyyy") & "'"
                sqlw &= " and SOTRMAF1.RA_DATE <= '" & Format(dteRATo.Value, "dd-MMM-yyyy") & "'"
                grdSOTRMAFI.Text &= " between " & Format(dteRAFrom.Value, "MM/dd/yyyy") & " and " & Format(dteRATo.Value, "MM/dd/yyyy")
            End If
        Else
            sqlw &= " and SOTRMAF1.RA_STATUS = 'O'"
        End If
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE = "" Then
            grdSOTRMAFX.Text = optStatus.Text
        Else
            sqlw &= " and CUST_CODE = '" & CUST_CODE & "'"
            grdSOTRMAFX.Text = optStatus.Text & " associated with " & CUST_CODE
            grdSOTRMAFI.Text &= " associated with Customer " & CUST_CODE
        End If

        If optStatus.Value = "F" Or optStatus.Value = "A" Then
            grdSOTRMAFX.Text &= "; RA's Dates between " & Format(dteRAFrom.Value, "MM/dd/yyyy") & " and " & Format(dteRATo.Value, "MM/dd/yyyy")
        End If

        Create_Work_Tables(sqlw)

        Fill_Records("SOTRMAFX")
        Sort_grdColumns(grdSOTRMAFX, "RA_NO".ToLower)
        grdSOTRMAFX.Visible = True

        Fill_Records("SOTRMAFI")
        Sort_grdColumns(grdSOTRMAFI, "RA_NO".ToLower)
        grdSOTRMAFI.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False, True)
        grdSOTRMAFX.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Create_Work_Tables(sqlw As String)
        Dim sqlTotals As String = "Select SOTRMAF2.RA_NO" _
    & ", Sum (SOTRMAF2.RA_QTY) RA_QTY" _
    & ", Sum (SOTRMAF2.RA_QTY_OPEN) RA_QTY_OPEN" _
    & ", Sum (SOTRMAF2.RA_QTY_USED) RA_QTY_USED" _
    & ", Sum (SOTRMAF2.RA_QTY_CANC) RA_QTY_CANC" _
    & ", Sum (NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT" _
    & ", Sum (NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_OPEN" _
    & ", Sum (NVL(SOTRMAF2.RA_QTY_USED,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_USED" _
    & ", Sum (NVL(SOTRMAF2.RA_QTY_CANC,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_CANC" _
    & ", Sum (NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_RETAIL,0)) RA_RETAIL_EXT" _
    & " from SOTRMAF2,SOTRMAF1 where SOTRMAF2.RA_NO = SOTRMAF1.RA_NO" & sqlw & " group by SOTRMAF2.RA_NO"

        ASCMAIN1.sql = "Select SOTRMAF1.*" & vbCrLf _
            & ", X.RA_QTY, X.RA_QTY_OPEN, X.RA_QTY_USED, X.RA_QTY_CANC" & vbCrLf _
            & ", X.RA_AMT, X.RA_AMT_OPEN, X.RA_AMT_USED, X.RA_AMT_CANC, X.RA_RETAIL_EXT" & vbCrLf _
            & " from (" & sqlTotals & ") X, SOTRMAF1" & ASCMAIN1.SQL_Add_WHERE(sqlw & " and X.RA_NO = SOTRMAF1.RA_NO")

        If SOTRMAFX = "" Then
            SOTRMAFX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTRMAFX & " Add Primary Key (RA_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTRMAFX)
            ASCDATA1.ExecuteSQL("Insert into " & SOTRMAFX & " " & ASCMAIN1.sql)
        End If



        ASCMAIN1.sql = "Select SOTRMAF2.ITEM_CODE, ICTITEM1.ITEM_DESC, SOTRMAF2.RA_LNO, SOTRMAF1.*" _
            & ", ICTITEM1.COLLECTION_CODE" _
            & ", ICTCOLL1.HC_CODE" _
            & ", ICTCOLL1.BRAND_CODE" _
            & ", SOTRMAF2.RA_QTY RA_QTY" _
            & ", SOTRMAF2.RA_RETAIL RA_RETAIL" _
            & ", SOTRMAF2.RA_NET_PRICE RA_NET_PRICE" _
            & ", SOTRMAF2.RA_QTY_OPEN RA_QTY_OPEN" _
            & ", SOTRMAF2.RA_QTY_USED RA_QTY_USED" _
            & ", SOTRMAF2.RA_QTY_CANC RA_QTY_CANC" _
            & ", NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_NET_PRICE,0) RA_AMT" _
            & ", NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(SOTRMAF2.RA_NET_PRICE,0) RA_AMT_OPEN" _
            & ", NVL(SOTRMAF2.RA_QTY_USED,0) * NVL(SOTRMAF2.RA_NET_PRICE,0) RA_AMT_USED" _
            & ", NVL(SOTRMAF2.RA_QTY_CANC,0) * NVL(SOTRMAF2.RA_NET_PRICE,0) RA_AMT_CANC" _
            & ", NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_RETAIL,0) RA_RETAIL_EXT" _
            & " from ICTITEM1, ICTCOLL1, SOTRMAF2, SOTRMAF1, " & SOTRMAFX & " X " _
            & " where SOTRMAF2.RA_NO = X.RA_NO " _
            & "   and ICTITEM1.ITEM_CODE = SOTRMAF2.ITEM_CODE" _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" _
            & "   and SOTRMAF1.RA_NO = SOTRMAF2.RA_NO"


        If SOTRMAFI = "" Then
            SOTRMAFI = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTRMAFI & " Add Primary Key (RA_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTRMAFI)
            ASCDATA1.ExecuteSQL("Insert into " & SOTRMAFI & " " & ASCMAIN1.sql)
        End If

    End Sub
    Sub Print_Record()

        ' To use the data layer and dst that is associated with this form

        Fill_Records("ARTCUST1", CUST_CODE)
        Fill_Records("ARTCUST2", New String() {CUST_CODE, Absx1.txtFor("CUST_STORE_NO").Text})
        Fill_Records("SOTSREP1", Absx1.txtFor("SREP_CODE").Text)
        Fill_Records("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)

        rowSOTRMAF1.Item("AR_PARM_KEY") = "Z"
        rowSOTRMAF1.Item("RA_AMT") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT)", "") & "")

        Print_Report_Begin()
        CR_params.Add("SUBT", "")
        Dim RPT As String = "SORRMAP1" ' unneccesary if Report Name is Like Form Name
        Generate_Report(RPT, "Returns Authorization", , , , , False)
        Print_Report_End()

    End Sub

    Private Sub Print_Credit_Memo()
        Try
            Me.Cursor = Cursors.WaitCursor

            ASCMAIN1.Progress("Now Preparing Credit Memo for Printing")

            Dim REPORT_NAME As String = "SORINVP1"
            Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
            If RPT = "" Then RPT = REPORT_NAME

            If Not REPORTS.ContainsKey(REPORT_NAME) Then
                REPORTS.Add(REPORT_NAME, Load_rptClass(REPORT_NAME))
                REPORTS(REPORT_NAME).Prepare_dst(False, "")
            End If

            Dim sql As String = " and SOTINVH1.INV_TYPE = '" & rowSOTRMAF1.Item("INV_TYPE") & "' and SOTINVH1.INV_NO = '" & rowSOTRMAF1.Item("INV_NUM") & "'"
            Dim tempFileName As String = "Memo" & DateTime.Now.ToString("yyyyMMddHHmmss")

            REPORTS(REPORT_NAME).Fill_Records_RPT(sql)
            Dim FILENAME As String = ""
            With REPORTS(REPORT_NAME).clsASCBASE1
                .Print_Report_Begin()
                .CR_params.Add("SUBT", "")
                .CR_params.Add("CONS_INV", "")
                Dim REPORT_NO As String = .Generate_Report(RPT, "", "", False, False, "", "PDF", tempFileName, False)
                FILENAME = .F.REPORT_FILENAMES(REPORT_NO)
                .Print_Report_End(, True)
            End With

            Show_Document(FILENAME)

        Catch ex As Exception
            MessageBox.Show(ex.Message, "Print Credit Memo", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub grdSOTRMAFX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRMAFX.DoubleClickRow

        ' Prevebt DBNUll error
        If e.Row Is Nothing OrElse e.Row.IsFilterRow Then
            Exit Sub
        End If

        Absx1.txtFor("RA_NO").Text = e.Row.Cells("RA_NO").Value & String.Empty
        Click_Command("View")
    End Sub

    Sub Cancel_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        Dim EMsg As String = ""
        If EntryMode = "E" Then
            Cancel_Order_1(RA_NO)
            EMsg = "Balance Open on Returns Authorization " & RA_NO & " has been Cancelled"
        End If

        'ASCDATA1.ExecuteSP("SOPRSRV0_G", "V", New Object() {RSRV_GROUP_NO}, New String() {"RSRV_GROUP_NO_IN"})
        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Cancel_Order_1(RA_NO As String)
        Dependent_Updates(-1, RA_NO)

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is Select * from SOTRMAF2 where RA_NO = '" & RA_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTRMAF2" _
            & "    Set RA_QTY_CANC = NVL(RA_QTY_CANC,0) + NVL(R1.RA_QTY_OPEN,0)" _
            & "      , RA_QTY_OPEN = 0" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()
        ', RA_STATUS = 'C'
        ASCMAIN1.sql = "Update SOTRMAF1 Set RA_STATUS = :PARM1" _
            & " where RA_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"F", RA_NO})
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        Dim EMsg As String = ""

        BeginTrans()

        If EntryMode = "E" Then
            Delete_Order_1(RA_NO)
            EMsg = "Returns Authorization No " & RA_NO & " has been marked as Deleted"
        End If

        CommitTrans(EMsg)
        'ASCDATA1.ExecuteSP("SOPRSRV0_G", "V", New Object() {RSRV_GROUP_NO}, New String() {"RSRV_GROUP_NO_IN"})
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(RA_NO As String)
        Dependent_Updates(-1, RA_NO)

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select * from SOTRMAF2" & vbCrLf _
            & "     where RA_NO = '" & RA_NO & "' for Update;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SOTRMAF2" & vbCrLf _
            & "    Set RA_QTY_CANC = NVL(RA_QTY_CANC,0) + NVL(R1.RA_QTY_OPEN,0)" & vbCrLf _
            & "   , RA_QTY_OPEN = 0" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTRMAF1 Set RA_STATUS = :PARM1" _
            & " where RA_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"D", RA_NO})
    End Sub

    Sub Dependent_Updates(S As Integer, RA_NO As String)

    End Sub

    Sub Display_Totals()
        Dim KEY As Int32 = 0
        For Each SFX As String In New String() {"", "OPEN", "USED", "CANC"}
            If SFX <> "" Then SFX = "_" & SFX
            KEY += 1
            Dim rowSOTRMAFT As DataRow = dst.Tables("SOTRMAFT").Rows.Find(KEY)
            rowSOTRMAFT.Item("QTY") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_QTY" & SFX & ")", "") & "")
            rowSOTRMAFT.Item("AMT") = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT" & SFX & ")", "") & "")
        Next
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTRMAF2.ActiveRow
            Select Case COLUMN_NAME
                Case "ITEM_CODE"
                    Dim ITEM_CODE As String = ""
                    If Trim(.Cells("ITEM_CODE").Value & "") <> "" Then
                        ITEM_CODE = Validate_Item(.Cells("ITEM_CODE").Value & "")
                    End If
                    Cancel = (ITEM_CODE = "")

                Case "RA_QTY"
                    If Trim(.Cells("ITEM_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Val(.Cells("RA_QTY").Value & "") = 0 Then
                        MsgBox("Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("RA_QTY")
                        Exit Sub
                    End If
                    If Val(.Cells("RA_QTY").Value & "") < 0 Then
                        MsgBox("Qty May Not be Negative", vbOKOnly, "Invalid Quantity")
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Function Validate_Item(ITEM_CODE_z As String) As String
        Dim EMsg As String = ""
        If ITEM_CODE_z = "" Then Return ""

        Dim ITEM_CODE As String = ""
        rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE_z)

        If rowICTITEM1 Is Nothing Then
            EMsg = "Item is Not on File" & vbCrLf
        Else
            'If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
            '    EMsg = "Item Status is not Active" & vbCrLf
            'End If
            If rowICTITEM1.Item("ITEM_UOM") & "" = "" Then
                EMsg = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            'If rowICTITEM1.Item("SALES_DIVISION_CODE") & "" = "" Then
            '    EMsg = "Item does not have a valid Division Code" & vbCrLf
            'End If
        End If

        If EMsg <> "" And grdSOTRMAF2.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Item Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                ITEM_CODE = rowICTITEM1.Item(0)
            End If
        End If
        Return ITEM_CODE
    End Function

#Region "grdSOTRMAF2"

    Private Sub grdSOTRMAF2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRMAF2.AfterCellUpdate
        With grdSOTRMAF2.ActiveRow
            Select Case e.Cell.Column.Key
                Case "ITEM_CODE"
                    Dim ITEM_CODE As String = Validate_Item(.Cells("ITEM_CODE").Value)
                    If ITEM_CODE <> "" Then
                        .Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                        .Cells("ITEM_SNU_CODE").Value = rowICTITEM1.Item("ITEM_SNU_CODE")
                        Dim ITEM_RETAIL_PRICE As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                        Dim RA_DATE As Date

                        ' get lowest price from invoice sales
                        If dst.Tables("SOTINVH2X").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length > 0 Then
                            Dim RA_NET_PRICE As Decimal = Val(dst.Tables("SOTINVH2X").Compute("MIN(ORDR_UNIT_PRICE)", "ITEM_CODE = '" & ITEM_CODE & "'") & String.Empty)
                            .Cells("RA_NET_PRICE").Value = RA_NET_PRICE
                        ElseIf rowSOTRMAF1.Item("RA_DATE") & "" = "" Then
                            MsgBox("Please specify an RA Date before entering Items", MsgBoxStyle.OkOnly, "Cannot Determine Price")
                        Else
                            RA_DATE = rowSOTRMAF1.Item("RA_DATE")

                            Dim ORDR_UNIT_PRICE As Decimal = TAC.SOCMAIN1.Get_Price _
                                                             (Me,
                                                              PRICE_LIST_CODE,
                                                              PRICE_LIST_CODE_ALLO,
                                                              PRICE_BASIS,
                                                              PRICE_BASE_DPCT,
                                                              ITEM_CODE,
                                                              rowICTITEM1,
                                                              rowSOTRMAF1.Item("RA_DATE"), ITEM_RETAIL_PRICE) ' MAYBE SHOULD USE RA_DATE - 60

                            .Cells("RA_NET_PRICE").Value = ORDR_UNIT_PRICE
                        End If

                        .Cells("RA_RETAIL").Value = ITEM_RETAIL_PRICE
                        If ASCMAIN1.CLIENT = "INT" Then
                            If ITEM_CODE <> "" Then .Cells("QTY_EOW").Value = Load_Retail_On_Hand(False, ITEM_CODE)
                        End If

                    End If

                Case "RA_QTY"
                    .Cells("RA_QTY_OPEN").Value = .Cells("RA_QTY").Value

                Case "RA_QTY_OPEN"
                    .Cells("RA_QTY_CANC").Value _
                        = Val(.Cells("RA_QTY").Value & "") _
                        - Val(.Cells("RA_QTY_USED").Value & "") _
                        - Val(.Cells("RA_QTY_OPEN").Value & "")
                    If Val(.Cells("RA_QTY_CANC").Value) < 0 Then
                        .Cells("RA_QTY_CANC").Value = 0
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSOTRMAF2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTRMAF2.AfterRowActivate

        If Trim(grdSOTRMAF2.ActiveRow.Cells("ITEM_CODE").Value & "") = "" And
            (grdSOTRMAF2.ActiveCell Is Nothing OrElse
             (grdSOTRMAF2.ActiveCell.Column.Key <> "ITEM_CODE")) _
        Then
            grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("ITEM_CODE")
            Exit Sub
        End If

        With grdSOTRMAF2.DisplayLayout.Bands(0)
            If grdSOTRMAF2.ActiveRow.IsAddRow Then
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                If grdSOTRMAF2.ActiveRow.Cells("ITEM_CODE").Value & "" = "" Then
                    grdSOTRMAF2.ActiveCell = grdSOTRMAF2.ActiveRow.Cells("ITEM_CODE")
                End If
            Else
                Validate_Item(grdSOTRMAF2.ActiveRow.Cells("ITEM_CODE").Value & "")
                .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                If Val(grdSOTRMAF2.ActiveRow.Cells("RA_QTY_USED").Value & "") <> 0 _
                Or Val(grdSOTRMAF2.ActiveRow.Cells("RA_QTY_CANC").Value & "") <> 0 _
                Then
                    .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    ' 01/07/2019 - As per Lauren
                    If ASCMAIN1.CLIENT = "INT" Then
                        .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    End If
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("RA_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("RA_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End If
        End With
    End Sub

    Private Sub grdSOTRMAF2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTRMAF2.AfterRowsDeleted
        Display_Totals()

        If grdSOTRMAF2.Rows.Count = 0 Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = ""
        End If
    End Sub

    Private Sub grdSOTRMAF2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTRMAF2.AfterRowUpdate
        Display_Totals()

        'If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
        '    Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTITEM1.Item("SALES_DIVISION_CODE")
        'End If
    End Sub

    Private Sub grdSOTRMAF2_BeforeCellCancelUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableCellEventArgs) Handles grdSOTRMAF2.BeforeCellCancelUpdate

    End Sub

    Private Sub grdSOTRMAF2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTRMAF2.BeforeCellUpdate

        If grdSOTRMAF2.ActiveCell IsNot Nothing Then
            With grdSOTRMAF2.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        If .Value & "" <> "" Then
                            Dim ITEM_CODE As String = Validate_Item(.Text)
                            If ITEM_CODE <> "" Then
                                If .Row.IsAddRow Then
                                    If dst.Tables("SOTRMAF2").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length <> 0 Then
                                        MsgBox("Item is already part of this RA Entry", MsgBoxStyle.OkOnly, "Cannot Add This Item")
                                        e.Cancel = True
                                        Exit Sub
                                    End If
                                End If
                            Else
                                e.Cancel = True
                            End If
                        End If
                End Select
            End With
        End If

    End Sub

    Private Sub grdSOTRMAF2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTRMAF2.BeforeExitEditMode
        If grdSOTRMAF2.ActiveCell IsNot Nothing Then
            With grdSOTRMAF2.ActiveCell
                Select Case .Column.Key
                    Case "ITEM_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTRMAF2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTRMAF2.BeforeRowsDeleted

        RA_LNOs.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If grow.IsAddRow Then
                e.Cancel = True
                grow.CancelUpdate()
                Exit Sub
            End If
            If Val(grow.Cells("RA_QTY_USED").Value & "") <> 0 _
            Or Val(grow.Cells("RA_QTY_CANC").Value & "") <> 0 _
            Then
                MsgBox("Cannot Delete a Line if it has ever been " & vbCr & "Used Or Cancelled" & vbCr & "Use the Cancel Button (x)")
                e.Cancel = True
                Exit Sub
            End If

            RA_LNOs.Add(grow.Cells("RA_LNO").Value)
        Next
    End Sub

    Private Sub grdSOTRMAF2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTRMAF2.BeforeRowUpdate

        Validate_Columns("ITEM_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("RA_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("RA_NO").Value = RA_NO
            Dim RA_LNO As Int64 = Val(dst.Tables("SOTRMAF2").Compute("MAX(RA_LNO)", "") & "") + 1
            e.Row.Cells("RA_LNO").Value = RA_LNO
        Else
            Dim RA_QTY As Int32 = Val(e.Row.Cells("RA_QTY").Value & String.Empty)
            Dim RA_QTY_OPEN As Int32 = Val(e.Row.Cells("RA_QTY_OPEN").Value & String.Empty)
            Dim RA_QTY_USED As Int32 = Val(e.Row.Cells("RA_QTY_USED").Value & String.Empty)
            Dim RA_QTY_CANC As Int32 = Val(e.Row.Cells("RA_QTY_CANC").Value & String.Empty)

            Dim totalConsumed As Int32 = RA_QTY_CANC + RA_QTY_USED
            If totalConsumed > RA_QTY Then
                MessageBox.Show("RA quantity must be greater equal than the total quantity of Used and Cancelled.", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
                Exit Sub
            End If
            If RA_QTY - RA_QTY_OPEN - RA_QTY_USED <> RA_QTY_CANC Then
                e.Row.Cells("RA_QTY_CANC").Value = RA_QTY - RA_QTY_OPEN - RA_QTY_USED
            Else
                ' THIS SHOULD NEVER WIND UP CHANGING RA_QTY_OPEN
                e.Row.Cells("RA_QTY_OPEN").Value = RA_QTY - (RA_QTY_USED + RA_QTY_CANC)
            End If
            'e.Row.Cells("RA_QTY_OPEN").Value = RA_QTY - (RA_QTY_USED + RA_QTY_CANC)
        End If
    End Sub

    Private Sub grdSOTRMAF2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRMAF2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "X"
                    If Val(.Cells("RA_QTY_CANC").Value) <> 0 Then
                        If MsgBox("Restore Cancelled Qty of " & .Cells("RA_QTY_CANC").Value,
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        .Cells("RA_QTY_OPEN").Value = Val(.Cells("RA_QTY_OPEN").Value & "") + Val(.Cells("RA_QTY_CANC").Value & "")

                        .Update()
                    Else
                        If MsgBox("Cancel Remaining Qty Open of " & .Cells("RA_QTY_OPEN").Value,
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        .Cells("RA_QTY_OPEN").Value = "0"
                        ' grdSOWRMAF2_AfterColUpdate(.Cells("RA_QTY_OPEN").position)
                        grdSOTRMAF2.ActiveRow.Update()
                    End If

                Case "ITEM_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTRMAF2, sql_where)
            End Select
        End With

    End Sub

#End Region

    Sub Load_Events()
        '    grdEvents.RemoveAll
        '    Call Load_Events_1("Entered", "INIT_DATE")
        '    Call Load_Events_1("Modified", "LAST_DATE")
        '    Call Load_Events_1("Released", "RSRV_DATE_REL")
    End Sub

    Function Record_AR_Item(Optional reverse_Credit As Boolean = False) As String
        Dim INV_NUM As String = ASCMAIN1.Next_Control_No("SOTINVH1.INV_NO")
        Dim RA_AMT As Decimal = Val(rowSOTRMAF1.Item("RA_AMT") & "")
        Dim S As Integer = -1
        If reverse_Credit Then S = 1

        Dim rowSOTTYPE1 As DataRow = LookUp("SOTTYPE1", ROWs("SOTPARM1").Item("SO_PARM_RA_DIF_TYPE"))

        Dim rowARTREASR As DataRow = LookUp("ARTREASR", rowSOTRMAF1.Item("RA_REASON_CODE"))
        Dim REASON_CODE As String = rowARTREASR.Item("REASON_CODE") ' WHY NOT USE SOTTYPE1 FOR THIS?  EVEN THO ARTREASR HAS REASON CODES FOR EACH TYPE OF RA, THEY NEVER GET USED - THEY JUST GET RETURNED

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        With rowARTOPEN1
            .Item("CUST_CODE") = CUST_BILL_TO_CUST
            .Item("INV_TYPE") = "C"
            .Item("INV_NUM") = INV_NUM
            .Item("INV_DATE") = rowSOTRMAF1.Item("RA_DATE")
            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE")
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE")
            .Item("INV_DUE_DATE") = rowSOTRMAF1.Item("RA_DATE")
            .Item("INV_CUST_PO") = rowSOTRMAF1.Item("CUST_CLAIM_NO")
            If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                .Item("INV_SALES") = S * RA_AMT
            Else
                .Item("INV_MISC_CHG") = S * RA_AMT
            End If
            .Item("INV_TOTAL_AMOUNT") = S * RA_AMT
            .Item("INV_BALANCE") = S * RA_AMT
            .Item("CUST_CODE_SO") = CUST_CODE

            .Item("REASON_CODE") = REASON_CODE ' rowSOTTYPE1.Item("REASON_CODE")

            .Item("ORDR_NO") = RA_NO
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("SALES_DIVISION_CODE") = rowSOTRMAF1.Item("SALES_DIVISION_CODE")
            .Item("SREP_CODE") = rowSOTRMAF1.Item("SREP_CODE")

            .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")

            .Item("INV_SALES_CURR") = .Item("INV_SALES")
            .Item("INV_DISC_CURR") = .Item("INV_DISC")
            .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
            .Item("INV_STAX_CURR") = .Item("INV_STAX")
            .Item("INV_MISC_CHG_CURR") = .Item("INV_MISC_CHG")
            .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT")
            .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")

            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1

            .Item("INV_NOTES") = rowSOTRMAF1.Item("RA_NOTES")
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            ' .Item("SELL_CODE") = rowSOTRMAF1.Item("SELL_CODE")
        End With
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
        Update_Record_TDA("ARTOPEN1")

        Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
        With rowSOTINVH1
            .Item("INV_TYPE") = "C"
            .Item("INV_NO") = INV_NUM
            .Item("CUST_CODE") = CUST_CODE
            .Item("CUST_STORE_NO") = rowSOTRMAF1.Item("CUST_STORE_NO")
            If .Item("CUST_STORE_NO") & "" = "" Then
                .Item("CUST_STORE_NO") = "000000"
            End If
            .Item("ORDR_CUST_PO") = rowSOTRMAF1.Item("CUST_CLAIM_NO")
            .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE")
            .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE")
            .Item("SREP_CODE") = rowSOTRMAF1.Item("SREP_CODE")
            .Item("REASON_CODE") = REASON_CODE ' rowSOTTYPE1.Item("REASON_CODE")
            .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
            If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                .Item("INV_SALES") = S * RA_AMT
            Else
                .Item("INV_MISC_CHG") = S * RA_AMT
                .Item("MISC_CHG_CODE") = rowSOTTYPE1.Item("MISC_CHG_CODE")
            End If
            .Item("INV_TOTAL_AMOUNT") = S * RA_AMT
            .Item("INV_DATE") = rowSOTRMAF1.Item("RA_DATE")
            .Item("ORDR_YYYYPP_UPDATED") = ASCMAIN1.CYP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            .Item("CURR_EXCH_RATE") = 1
            .Item("INV_COMMENT") = rowSOTRMAF1.Item("RA_NOTES")

            If Not reverse_Credit Then
                Dim ORDR_CUST_PO As String = (rowSOTRMAF1.Item("CUST_CLAIM_NO") & String.Empty).ToString.Trim
                If ORDR_CUST_PO.Length > 0 Then
                    Dim INV_NO_CR As String = String.Empty
                    Dim CC_SALE_TRANS_ID As String = String.Empty
                    SOCMAIN1.GetCreditCardSaleTransaction(CUST_CODE, ORDR_CUST_PO, INV_NO_CR, CC_SALE_TRANS_ID)
                    .Item("CC_SALE_TRANS_ID") = CC_SALE_TRANS_ID
                    .Item("INV_NO_CR") = INV_NO_CR
                End If
            End If

            .Item("ORDR_TYPE_CODE") = ROWs("SOTPARM1").Item("SO_PARM_RA_DIF_TYPE")

            .Item("OPS_YYYYWW") = ASCMAIN1.CYW
            .Item("REGISTER_IND") = "0"
            .Item("ORDR_DEPT") = rowSOTRMAF1.Item("ORDR_DEPT")
            .Item("SREP2_CODE") = rowSOTRMAF1.Item("SREP2_CODE")
            .Item("SALES_DIVISION_CODE") = rowSOTRMAF1.Item("SALES_DIVISION_CODE")

            Dim ITEM_CODE As String = dst.Tables("SOTRMAF2").Select("")(0).Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", rowICTITEM1.Item("COLLECTION_CODE"))
            .Item("BRAND_CODE") = rowICTCOLL1.Item("BRAND_CODE")
            ' MAYBE WE SHOULD PICK UP DIVISION HERE TOO?

            .Item("INV_UNITS") = dst.Tables("SOTRMAF2").Compute("Sum (RA_QTY)", "")

        End With
        dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)
        Update_Record_TDA("SOTINVH1")

        If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
            For Each rowSOTRMAF2 As DataRow In dst.Tables("SOTRMAF2").Select("")
                Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").NewRow
                With rowSOTINVH2
                    .Item("INV_TYPE") = "C"
                    .Item("INV_NO") = INV_NUM
                    .Item("INV_LNO") = rowSOTRMAF2.Item("RA_LNO")
                    .Item("ITEM_CODE") = rowSOTRMAF2.Item("ITEM_CODE")
                    .Item("ORDR_UNIT_PRICE") = rowSOTRMAF2.Item("RA_NET_PRICE")
                    .Item("ORDR_QTY_SHIP") = -1 * Val(rowSOTRMAF2.Item("RA_QTY") & "")
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = rowSOTINVH1.Item("CUST_STORE_NO")
                    .Item("SREP_CODE") = rowSOTINVH1.Item("SREP_CODE")
                    .Item("ORDR_YYYYPP_UPDATED") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")
                    .Item("ORDR_UNIT_PRICE_CURR") = rowSOTRMAF2.Item("RA_NET_PRICE")
                    .Item("ITEM_UNIT_COST") = 0
                    .Item("ITEM_RETAIL_PRICE") = rowSOTRMAF2.Item("RA_RETAIL")
                    .Item("ITEM_RETAIL_PRICE_CURR") = rowSOTRMAF2.Item("RA_RETAIL")
                    .Item("OPS_YYYYWW") = rowSOTINVH1.Item("OPS_YYYYWW")
                    .Item("SELL_CODE") = rowSOTINVH1.Item("SELL_CODE")
                End With
                dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)
            Next
            Update_Record_TDA("SOTINVH2")
        End If

        ASCDATA1.ExecuteSP("ARPCUST6_IC", "VV",
                           New Object() {"C", INV_NUM},
                           New String() {"INV_TYPE_IN", "INV_NO_IN"})

        If reverse_Credit Then
            rowSOTRMAF1.Item("INV_TYPE_REV") = "C"
            rowSOTRMAF1.Item("INV_NUM_REV") = INV_NUM
            rowSOTRMAF1.Item("OPS_YYYYPP_REV") = ASCMAIN1.CYP
        Else
            rowSOTRMAF1.Item("INV_TYPE") = "C"
            rowSOTRMAF1.Item("INV_NUM") = INV_NUM
            rowSOTRMAF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        End If

        Return INV_NUM
    End Function

    Sub Record_AR_Item_Reversal(row As DataRow)
        Dim INV_NUM As String = ASCMAIN1.Next_Control_No("AR_CR_MEMO_NO")
        Dim INV_TOTAL_AMOUNT As Decimal = Val(row.Item("INV_TOTAL_AMOUNT") & "")
        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        rowARTOPEN1.ItemArray = row.ItemArray
        With rowARTOPEN1
            .Item("INV_NUM").Value = INV_NUM
            .Item("INV_MISC_CHG").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_TOTAL_AMOUNT").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_BALANCE").Value = -1 * INV_TOTAL_AMOUNT

            .Item("INV_MISC_CHG_CURR").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_TOTAL_AMOUNT_CURR").Value = -1 * INV_TOTAL_AMOUNT
            .Item("INV_BALANCE_CURR").Value = -1 * INV_TOTAL_AMOUNT

            .Item("INIT_OPER").Value = ASCMAIN1.USER_ID
            .Item("INIT_DATE").Value = DATETIME_STAMP
        End With
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
        Update_Record_TDA("ARTOPEN1")
    End Sub

    Sub Setup_Price_Class()
        With grdSOTRMAF2.DisplayLayout.Bands(0)
            If PRICE_BASIS = "E" Then
                .Columns("RA_RETAIL").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("RA_RETAIL_EXT").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RA_NET_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("RA_RETAIL").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RA_RETAIL_EXT").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RA_NET_PRICE").CellActivation = UltraWinGrid.Activation.AllowEdit
            End If
        End With
        If PRICE_BASE_DPCT = 100 Then
            dst.Tables("SOTRMAF2").Columns("RA_RETAIL_EXT").Expression = "0"
        Else
            dst.Tables("SOTRMAF2").Columns("RA_RETAIL_EXT").Expression = "IIF(ISNULL(RA_QTY,0)=0,RA_AMT/((100 - " & CStr(PRICE_BASE_DPCT) & ") / 100),ISNULL(RA_QTY,0) * ISNULL(RA_RETAIL,0))"
        End If

    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTRMAFX()

        grpRADATE.Visible = (optStatus.Value = "F" Or optStatus.Value = "A")
    End Sub

    Private Sub grdSOTRTRNX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTRTRNX.InitializeRow
        If e.Row.Cells("REVERSED_BY_RTRN_NO").Value & "" <> "" Then
            e.Row.Cells("RTRN_NO").Appearance.ForeColor = System.Drawing.Color.Red
            e.Row.Cells("RTRN_NO").ToolTipText = "Reversed by Return No " & e.Row.Cells("REVERSED_BY_RTRN_NO").Value
        ElseIf e.Row.Cells("REVERSED_RTRN_NO").Value & "" <> "" Then
            e.Row.Cells("RTRN_NO").Appearance.ForeColor = System.Drawing.Color.Red
            e.Row.Cells("RTRN_NO").ToolTipText = "Reverses Return No " & e.Row.Cells("REVERSED_RTRN_NO").Value
        Else
            e.Row.Cells("RTRN_NO").Appearance.ForeColor = System.Drawing.Color.Empty
            e.Row.Cells("RTRN_NO").ToolTipText = ""
        End If
    End Sub

    Private Sub grdEDT812TX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs)
        If Not InquiryMode Then
            If e.Row.IsActiveRow And e.Row.IsDataRow Then
                Dim EDI_DOC_SEQ_NO As String = e.Row.Cells("EDI_DOC_SEQ_NO").Value
            End If
        End If
    End Sub

    Private Sub grdEDT180TX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs)
        If Not InquiryMode Then
            If e.Row.IsActiveRow And e.Row.IsDataRow Then
                Dim EDI_DOC_SEQ_NO As String = e.Row.Cells("EDI_DOC_SEQ_NO").Value
            End If
        End If
    End Sub

    Function Process_EDI(EDI As String, EDI_DOC_SEQ_NO As String) As Boolean
        Process_EDI = True
        Dim rowX As DataRow = LookUp("EDT" & EDI & "T1", EDI_DOC_SEQ_NO)
        With rowX
            If .Item("EDI_PROCESS_IND") & "" <> "0" And .Item("EDI_PROCESS_IND") & "" <> "" Then
                MsgBox("EDI Doc Seq No " & EDI_DOC_SEQ_NO & " has already been Processed", vbOKOnly, "Cannot Process " & EDI)
                Return True
            End If

            Dim CUST_CODE As String = rowX.Item("CUST_CODE")

            Absx1.txtFor("CUST_CODE").Text = CUST_CODE
            Dim CUST_CLAIM_NO As String = ""
            If EDI = "180" Then
                If CUST_CODE = "NORDSTROM" Then
                    CUST_CLAIM_NO = rowX.Item("RMA_NUMBER") & ""
                Else
                    CUST_CLAIM_NO = rowX.Item("TRANS_REF_NO") & ""
                End If
            Else
                CUST_CLAIM_NO = rowX.Item("EDI_CLAIM_NO") & ""
            End If

            If TAC.TACMAIN1.IPLBMacysCustomerCodes.Contains(CUST_CODE) Then
                CUST_CLAIM_NO = Mid(CUST_CLAIM_NO, CUST_CLAIM_NO.Length - 10 + 1, 10)
            End If

            If CUST_CODE = "NORDSTROM" Then
                If CUST_CLAIM_NO = "000000" And rowX.Item("RMA_NUMBER") & "" <> "" Then
                    CUST_CLAIM_NO = rowX.Item("RMA_NUMBER")
                End If
            End If

            Absx1.txtFor("CUST_CLAIM_NO").Text = CUST_CLAIM_NO
            'Absx1.txtFor("RA_NOTES").Text = e.Row.Cells("TRANS_REF_NO").Value
            Click_Command("New")

            If ScreenMode Then

                rowSOTRMAF1.Item("EDI_DOC_NO") = EDI
                rowSOTRMAF1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO

                Dim TRANS_DATE As Date

                If EDI = "180" Then
                    rowSOTRMAF1.Item("RA_NOTES") = rowX.Item("RMA_NUMBER")
                    TRANS_DATE = rowX.Item("TRANS_DATE")
                Else
                    rowSOTRMAF1.Item("RA_NOTES") = rowX.Item("EDI_FREIGHT_BILL_NO")
                    If rowX.Item("EDI_CLAIM_DATE") & "" = "" Then
                        TRANS_DATE = rowX.Item("EDI_CB_DATE")
                    Else
                        TRANS_DATE = rowX.Item("EDI_CLAIM_DATE")
                    End If
                End If

                Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
                Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

                If Format(TRANS_DATE, "yyyyMMdd") > Format(PRD_END_DATE, "yyyyMMdd") Then
                    rowSOTRMAF1.Item("RA_DATE") = PRD_END_DATE
                    rowSOTRMAF1.Item("RA_EXPIRE") = PRD_END_DATE
                Else
                    rowSOTRMAF1.Item("RA_DATE") = TRANS_DATE
                    rowSOTRMAF1.Item("RA_EXPIRE") = TRANS_DATE
                End If

                'If ASCMAIN1.CLIENT = "INT" Then
                '    If ASCMAIN1.CYP = "201601" Then
                '        If Format(rowSOTRMAF1.Item("RA_DATE"), "yyyyMMdd") < "20160101" Then
                '            rowSOTRMAF1.Item("RA_DATE") = CDate("01/01/2016")
                '        End If
                '        If Format(rowSOTRMAF1.Item("RA_EXPIRE"), "yyyyMMdd") < "20160101" Then
                '            rowSOTRMAF1.Item("RA_EXPIRE") = CDate("01/01/2016")
                '        End If
                '    End If
                'End If

                rowSOTRMAF1.Item("RA_REASON_CODE") = "X"

                RA_NO_EDI = rowSOTRMAF1.Item("RA_NO")

                'ASCMAIN1.sql = "Select * from EDT" & EDI & "T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                ' need to aggregate because 180s/812s sometimes repeat EANs on multiple lines, and RA's don't play that
                If EDI = "180" Then
                    ASCMAIN1.sql = "Select EDI_EAN, EDI_UPC, 0 EDI_PRICE, SUM (EDI_QTY) EDI_QTY from EDT180T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' group by EDI_EAN, EDI_UPC"
                Else
                    ASCMAIN1.sql = "Select EDI_EAN, EDI_UPC, EDI_PRICE, SUM (EDI_TOTAL_QTY) EDI_QTY from EDT812T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' group by EDI_EAN, EDI_UPC, EDI_PRICE"
                End If

                For Each row2 As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim ITEM_EAN_CODE As String = row2.Item("EDI_EAN") & ""
                    Dim ITEM_UPC_CODE As String = row2.Item("EDI_UPC") & ""
                    Dim rowICTITEM1 As DataRow = Nothing

                    If ITEM_EAN_CODE <> "" Then
                        ASCMAIN1.sql = "Select * from ICTITEM1 where ITEM_EAN_CODE = '" & ITEM_EAN_CODE & "'"
                        rowICTITEM1 = ASCDATA1.GetDataRow
                    ElseIf ITEM_UPC_CODE <> "" Then
                        ASCMAIN1.sql = "Select * from ICTITEM1 where ITEM_UPC_CODE = '" & ITEM_UPC_CODE & "'"
                        rowICTITEM1 = ASCDATA1.GetDataRow
                    End If

                    If rowICTITEM1 Is Nothing Then
                        Dim errorText As String = ""
                        If String.IsNullOrEmpty(ITEM_EAN_CODE) AndAlso String.IsNullOrEmpty(ITEM_UPC_CODE) Then
                            errorText = "Missing item EAN/UPC"
                        ElseIf Not String.IsNullOrEmpty(ITEM_EAN_CODE) Then
                            errorText = $"Invalid item EAN ({ITEM_EAN_CODE})"
                        ElseIf Not String.IsNullOrEmpty(ITEM_UPC_CODE) Then
                            errorText = $"Invalid item UPC ({ITEM_UPC_CODE})"
                        End If

                        MsgBox($"{errorText} on EDI Doc Seq No {EDI_DOC_SEQ_NO}", vbOKOnly, "Cannot Process " & EDI)
                        Return False
                    End If

                    If rowICTITEM1 IsNot Nothing Then

                        '**************************** TEMPORARY UNTIL WE GET AN ANSWER ON HOW TO DEAL WITH RA INTERACTION
                        Dim COLLECTION_CODE As String = rowICTITEM1.Item("COLLECTION_CODE") & ""
                        Dim ITEM_CODE As String = rowICTITEM1.Item("ITEM_CODE") & ""
                        Dim ITEM_DESC As String = rowICTITEM1.Item("ITEM_DESC") & ""
                        'If COLLECTION_CODE.StartsWith("LCS") And ITEM_DESC.StartsWith("XXX ") Then
                        '    MsgBox($"We have a LCS item with XXX in the description ({ITEM_CODE}) which has been sent in a {EDI} for DIF credit", vbOKOnly, "Please send Screenshot to ABS before clicking OK")
                        '    Throw New Exception($"We have a LCS item with XXX in the description ({ITEM_CODE}) which has been sent in a {EDI} for DIF credit")
                        '    Me.Close()
                        '    Exit Sub

                        'End If

                        If grdSOTRMAF2.ActiveRow IsNot Nothing AndAlso grdSOTRMAF2.ActiveRow.IsAddRow Then
                            grdSOTRMAF2.ActiveRow.CancelUpdate()
                        End If
                        grdSOTRMAF2.DisplayLayout.Bands(0).AddNew()
                        With grdSOTRMAF2.ActiveRow
                            .Cells("ITEM_CODE").Value = rowICTITEM1.Item("ITEM_CODE")
                            .Cells("RA_QTY").Value = row2.Item("EDI_QTY")
                            If EDI = "812" Then
                                Dim EDI_PRICE As Decimal = Val(row2.Item("EDI_PRICE") & "")
                                Dim NET_PRICE As Decimal = Val(.Cells("RA_NET_PRICE").Value & "")
                                .Cells("EDI_PRICE").Value = EDI_PRICE
                                .Cells("NET_PRICE").Value = NET_PRICE
                                If EDI_PRICE < NET_PRICE And EDI_PRICE > 0 Then
                                    .Cells("RA_NET_PRICE").Value = EDI_PRICE
                                End If
                            End If

                            .Update()
                        End With
                    End If
                Next
            End If
        End With
    End Function

    Sub Load_EDTXXXT1(EDI As String)

        Dim EDTXXXT1 As String = "EDT180T1"
        If EDI = "812" Then EDTXXXT1 = "EDT812T1"

        If Not InquiryMode Then
            ASCDATA1.ExecuteSQL("Update " & EDTXXXT1 & " Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_PROCESS_IND is Null")

            ASCMAIN1.sql = "Update " & EDTXXXT1 & " X Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
                   & " where EDI_OUR_ID = TRIM(X.EDI_OUR_ID) and EDI_TP_ID = TRIM(X.EDI_TP_ID))" _
                   & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Update " & EDTXXXT1 & " X Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
                   & " where EDI_TP_QUAL = X.EDI_TP_QUAL and EDI_TP_ID = X.EDI_TP_ID and EDI_DOC_NO = '" & EDI & "')" _
                   & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
            ASCDATA1.ExecuteSQL()

            If EDI = "812" Then

                ' right now DILLARDS is the only customer providing 812s - may need to rethink these if other customers start transmitting 812s
                ASCMAIN1.sql = "Update EDT812T1" & vbCrLf _
                    & "Set EDI_PROCESS_IND = 'C'" & vbCrLf _
                    & " where EDI_PROCESS_IND = '0' and EDI_CREDIT_DEBIT_FLAG = 'C'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Update EDT812T1 SET EDI_PROCESS_IND = 'N'" & vbCrLf _
                    & " where EDI_DOC_SEQ_NO IN (" & vbCrLf _
                    & "Select EDI_DOC_SEQ_NO from (" & vbCrLf _
                    & "Select EDT812T1.EDI_DOC_SEQ_NO" & vbCrLf _
                    & ", MIN (EDI_ADJ_REAS_CODE) A1" & vbCrLf _
                    & ", MAX (EDI_ADJ_REAS_CODE) A2" & vbCrLf _
                    & ", COUNT (*) RECORDS" & vbCrLf _
                    & " from GEN.EDT812T1,GEN.EDT812T2" & vbCrLf _
                    & " where EDT812T2.EDI_DOC_SEQ_NO = EDT812T1.EDI_DOC_SEQ_NO" & vbCrLf _
                    & "  and EDT812T1.EDI_PROCESS_IND = '0'" & vbCrLf _
                    & " group by EDT812T1.EDI_DOC_SEQ_NO" & vbCrLf _
                    & ") where A1 <> '93' and A1 <> 'GD')"

                '                   & ") where A1 <> '93' and A1 <> 'GD')"
                ' & " and EDT812T1.EDI_CLAIM_DATE >= '01-Dec-2023'" & vbCrLf _

                '     If ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, "GEN.", "GENINT.")
                ASCDATA1.ExecuteSQL()
            End If

            ASCMAIN1.sql = "Select * from " & EDTXXXT1 & " where EDI_PROCESS_IND = '0' and COMPANY_CODE is Null"
            If EDI_DOC_SEQ_NOs_no_company.Count <> 0 Then
                ASCMAIN1.sql &= " and EDI_DOC_SEQ_NO Not in ('" & Join(EDI_DOC_SEQ_NOs_no_company.ToArray, "','") & "')"
            End If
            Dim dt As DataTable = ASCDATA1.GetDataTable
            If dt.Rows.Count <> 0 Then
                For Each row As DataRow In dt.Rows
                    EDI_DOC_SEQ_NOs_no_company.Add(row.Item("EDI_DOC_SEQ_NO"))
                Next
                Using frm As New ASFMSGBF
                    frm.Show_grd(dt, Me, "EDI Transactions which could not be mapped to an ABSolution Company")
                End Using
            End If

        End If

        If EDI = "180" Then
            Fill_Records("EDT180TX")
            Sort_grdColumns(grdEDT180TX, "EDI_DOC_SEQ_NO")
        Else
            Fill_Records("EDT812TX")
            Sort_grdColumns(grdEDT812TX, "EDI_DOC_SEQ_NO")
        End If
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        Load_SOTRMAFX()
    End Sub

    Private Sub grdSOTRMAF2_DoubleClickCell(sender As Object, e As UltraWinGrid.DoubleClickCellEventArgs) Handles grdSOTRMAF2.DoubleClickCell
        If Not e.Cell.Row.IsAddRow Then
            If EntryMode = "N" Or EntryMode = "E" Then
                If e.Cell.Column.Key = "RA_QTY_OPEN" Then
                    Load_OH_into_RA(e.Cell.Row, "RA_QTY_OPEN")
                ElseIf e.Cell.Column.Key = "RA_QTY" Then
                    Load_OH_into_RA(e.Cell.Row, "RA_QTY")
                End If

            End If
        End If
    End Sub

    Function Add_SOTRMAF2(ITEM_CODE As String, QTY As Int64) As UltraWinGrid.UltraGridRow

        If grdSOTRMAF2.ActiveRow IsNot Nothing AndAlso grdSOTRMAF2.ActiveRow.IsAddRow Then
            grdSOTRMAF2.ActiveRow.CancelUpdate()
        End If
        grdSOTRMAF2.DisplayLayout.Bands(0).AddNew()
        With grdSOTRMAF2.ActiveRow
            .Cells("ITEM_CODE").Value = ITEM_CODE
            .Cells("RA_QTY").Value = QTY
            .Update()
        End With

        Return grdSOTRMAF2.ActiveRow
    End Function

    Function Load_Retail_On_Hand(Optional add_Item As Boolean = False, Optional single_ITEM_CODE As String = "") As Int32
        Dim YW As String = ASCMAIN1.Week_Calc(ASCMAIN1.CYW, -1)

        If single_ITEM_CODE = "" Then
            For Each row As DataRow In dst.Tables("SOTRMAF2").Select()
                row.Item("QTY_EOW") = DBNull.Value
            Next
        End If

        Dim qty As Int32 = 0

        'If ASCMAIN1.Running_in_VS AndAlso ASCMAIN1.USER_ID = "wjz" Then
        '    'Stop
        '    YW = "202447"
        'End If

        Dim CUST_STORE_NO As String = Absx1.txtFor("CUST_STORE_NO").Text
        ASCMAIN1.sql = "Select RSTRETL1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
            & ", SUM (QTY_SOLD) QTY_SOLD, SUM (QTY_EOW) QTY_EOW" & vbCrLf _
            & " from RSTRETL1,ICTITEM1" & vbCrLf _
            & " where RSTRETL1.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYWW = '" & YW & "'" & vbCrLf _
            & IIf(CUST_STORE_NO = "", "", "   and RSTRETL1.CUST_STORE_NO = '" & CUST_STORE_NO & "'" & vbCrLf) _
            & IIf(single_ITEM_CODE = "", "", "   and RSTRETL1.ITEM_CODE = '" & single_ITEM_CODE & "'" & vbCrLf) _
            & " group by RSTRETL1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_BASIC_PROMO"
        For Each row As DataRow In ASCDATA1.GetDataTable().Select()
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            If add_Item Then ASCMAIN1.Progress("-", ITEM_CODE)
            Dim found As Boolean = False
            Dim QTY_EOW As Int64 = Val(row.Item("QTY_EOW") & "")
            If QTY_EOW > 0 Then
                If single_ITEM_CODE <> "" Then
                    qty = QTY_EOW
                    Exit For
                End If
                For Each rowSOTRMAF2 As DataRow In dst.Tables("SOTRMAF2").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    rowSOTRMAF2.Item("QTY_EOW") = QTY_EOW
                    found = True
                Next
                If add_Item And Not found Then
                    Dim grow As UltraWinGrid.UltraGridRow = Add_SOTRMAF2(ITEM_CODE, QTY_EOW)
                    grow.Cells("QTY_EOW").Value = QTY_EOW
                    grow.Update()
                End If
            End If
        Next

        Return qty
    End Function

    Function Load_OH_into_RA(grow As UltraWinGrid.UltraGridRow, COL As String) As UltraWinGrid.UltraGridRow

        Dim RA_QTY As Int64 = Val(grow.Cells("RA_QTY").Value & "")
        Dim RA_QTY_OPEN As Int64 = Val(grow.Cells("RA_QTY_OPEN").Value & "")
        Dim RA_QTY_USED As Int64 = Val(grow.Cells("RA_QTY_USED").Value & "")
        Dim RA_QTY_CANC As Int64 = Val(grow.Cells("RA_QTY_CANC").Value & "")
        Dim QTY_EOW As Int64 = Val(grow.Cells("QTY_EOW").Value & "")

        Dim DIFF As Int64 = QTY_EOW - RA_QTY_OPEN

        grow.Cells(COL).Value = QTY_EOW ' grow.Cells("RA_QTY_OPEN").Value = QTY_EOW

        If COL = "RA_QTY_OPEN" Then
            RA_QTY_OPEN = QTY_EOW
            If RA_QTY < RA_QTY_OPEN + RA_QTY_USED + RA_QTY_CANC Then
                grow.Cells("RA_QTY").Value = RA_QTY_OPEN + RA_QTY_USED + RA_QTY_CANC
            End If
        Else
            RA_QTY = QTY_EOW
            If RA_QTY <> RA_QTY_OPEN + RA_QTY_USED + RA_QTY_CANC Then
                grow.Cells("RA_QTY_CANC").Value = RA_QTY - (RA_QTY_OPEN + RA_QTY_USED) ' + RA_QTY_CANC
            End If
        End If



        grow.Update()

        Return grow
    End Function

    Sub Paste_Item_Codes_to_Add_to_RA()

        Using FRM As New ASFMSGBF
            Dim ITEM_CODEs As String = FRM.Get_txtblock_from_User("Item Codes", "Paste Valid Item Codes into the Textbox Below")
            If ITEM_CODEs <> "" Then
                Dim BAD_ITEM_CODEs As String = ""
                Dim ADDS As Int64 = 0
                Dim SKIPS As Int64 = 0
                Dim BADS As Int64 = 0
                For Each ITEM_CODE As String In Split(ITEM_CODEs, vbCrLf)
                    If LookUp("ICTITEM1", ITEM_CODE) Is Nothing Then
                        BAD_ITEM_CODEs &= ITEM_CODE & vbCrLf
                        BADS += 1
                    Else
                        If dst.Tables("SOTRMAF2").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length <> 0 Then
                            ' ITEM IS ALREADY ON RA
                            SKIPS += 1
                        Else
                            Dim grow As UltraWinGrid.UltraGridRow = Add_SOTRMAF2(ITEM_CODE, 1)
                            If grow Is Nothing Then
                                BADS += 1
                            Else
                                ADDS += 1
                            End If

                        End If

                    End If
                Next

                Sort_grdColumns(grdSOTRMAF2, "RA_LNO")

                If BAD_ITEM_CODEs <> "" Then
                    MsgBox("The following Item Codes were Invalid:" & vbCrLf & vbCrLf & BAD_ITEM_CODEs,
                           MsgBoxStyle.OkOnly, "Verification")
                End If
                If BADS = 0 And SKIPS = 0 Then
                    MsgBox("All " & Split(ITEM_CODEs, vbCrLf).Count & " Items added Successfully", MsgBoxStyle.OkOnly, "Verification")
                Else
                    MsgBox(CStr(ADDS) & " Items added Successfully, " & CStr(BADS + SKIPS) & " items bad or skipped", MsgBoxStyle.OkOnly, "Verification")
                End If
            End If

        End Using
    End Sub

    Private Sub UltraLabel7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub SOFRMAF1_DoubleClick(sender As Object, e As EventArgs) Handles Me.DoubleClick

    End Sub

    Private Sub tab_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab.SelectedTabChanged
        If e.Tab Is tab.Tabs(0) Then
            ' Show the RMAno textbox
            RMAno.Visible = True
            UltraLabel7.Visible = True
        Else
            ' Hide the RMAno textbox
            RMAno.Visible = False
            UltraLabel7.Visible = False
        End If
    End Sub

    Private Sub grdEDTXXXX0_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDTXXXX0.AfterRowActivate
        If grdEDTXXXX0.ActiveRow.IsAddRow Then Return

        Dim ediType As String = optDifReportType.Value
        Fill_Records($"EDT{ediType}X1", {grdEDTXXXX0.ActiveRow.Cells("XNO").Value, Nothing, Nothing, Nothing, Nothing})
        Refresh_DIF_Data("1")
        grdEDTXXXX1.Text = $"DIFs for Batch {grdEDTXXXX0.ActiveRow.Cells("XNO").Value}"
    End Sub

    Private Sub optDifReportType_ValueChanged(sender As Object, e As EventArgs) Handles optDifReportType.ValueChanged
        If Me.SELECTION_NO = 0 Then
            Exit Sub
        End If

        Dim ediType As String = optDifReportType.Value
        'grdEDTXXXX0.DataSource = dst.Tables($"EDT{ediType}X0")
        'grdEDTXXXX1.DataSource = dst.Tables($"EDT{ediType}X1")

        Dim selectedCustomer As String = cbeSearchCust.Value
        Fill_Records("ARTCUSTS", optDifReportType.Value)

        If dst.Tables("ARTCUSTS").Select($"CUST_CODE='{selectedCustomer}'").Count Then
            cbeSearchCust.Value = selectedCustomer
        Else
            cbeSearchCust.SelectedIndex = 0
        End If


        Fill_Records($"EDT{ediType}X0")
        If dst.Tables($"EDT{ediType}X0").Rows.Count = 0 Then
            dst.Tables($"EDT{ediType}X1").Clear()
        End If

        Refresh_DIF_Data("0")
        Refresh_DIF_Data("1")
    End Sub

    Private Sub Refresh_DIF_Data(tableSfx As String)
        Dim ediType As String = optDifReportType.Value

        EnforceConstraints(False)


        If ediType = "812" Then
            dst.Tables($"EDT180X{tableSfx}").Clear()
            For Each ROW As DataRow In dst.Tables($"EDT812X{tableSfx}").Select
                dst.Tables($"EDT180X{tableSfx}").Rows.Add(ROW.ItemArray)
            Next
        End If

        EnforceConstraints(True)
    End Sub

    Private Sub tabRAMaster_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabRAMaster.SelectedTabChanged
        Hide_DIF_Group()
    End Sub

    Private Sub Hide_DIF_Group()

        If Me.SELECTION_NO = 0 Then
            Exit Sub
        End If

        UltraExplorerBar1.Groups("DIF Report").Visible = (tabRAMaster.SelectedTab.Key = "DIF Report")
        UltraExplorerBar1.Groups("DIF Search").Visible = (tabRAMaster.SelectedTab.Key = "DIF Report")
        UltraExplorerBar1.Groups("Status").Visible = (tabRAMaster.SelectedTab.Key <> "DIF Report")
    End Sub

    Private Sub btnSearchDif_Click(sender As Object, e As EventArgs) Handles btnSearchDif.Click
        ASCMAIN1.Progress("Searching...")
        splDifReport.Panel1Collapsed = True
        Fill_Records($"EDT{optDifReportType.Value}X1", {Nothing, cbeSearchCust.Value, dteSearchFrom.DateTime, dteSearchTo.DateTime, txtSearchClaimNo.Text})

        grdEDTXXXX1.Text = $"{cbeSearchCust.Text} {optDifReportType.Text} from {dteSearchFrom.DateTime.Date.ToShortDateString()} to {dteSearchTo.DateTime.Date.ToShortDateString()}{If(Not String.IsNullOrEmpty(txtSearchClaimNo.Text), ", Claim No contains " & txtSearchClaimNo.Text, "")}"

        ASCMAIN1.Progress("")
    End Sub

    Private Sub dteSearchFrom_ValueChanged(sender As Object, e As EventArgs) Handles dteSearchFrom.ValueChanged
        If dteSearchFrom.DateTime > dteSearchTo.DateTime Then
            dteSearchTo.DateTime = dteSearchFrom.DateTime
        End If
    End Sub

    Private Sub dteSearchTo_ValueChanged(sender As Object, e As EventArgs) Handles dteSearchTo.ValueChanged
        If dteSearchTo.DateTime < dteSearchFrom.DateTime Then
            dteSearchFrom.DateTime = dteSearchTo.DateTime
        End If
    End Sub

    Private Sub btnClearSearch_Click(sender As Object, e As EventArgs) Handles btnClearSearch.Click
        dst.Tables($"EDT{optDifReportType.Value}X1").Clear()
        splDifReport.Panel1Collapsed = False
        If grdEDTXXXX0.Rows.Count > 0 Then
            grdEDTXXXX0.ActiveRow = Nothing
            grdEDTXXXX0.Rows(0).Activate()
        End If
    End Sub

    Private Sub grdEDTXXXX1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDTXXXX1.DoubleClickRow

        If e.Row Is Nothing OrElse e.Row.IsFilterRow Then
            Exit Sub
        End If

        Absx1.txtFor("RA_NO").Text = e.Row.Cells("RA_NO").Value & String.Empty
        Click_Command("View")
    End Sub

    Private Sub chk_showDetails180_CheckedChanged(sender As Object, e As EventArgs) Handles chk_showDetails180.CheckedChanged
        If chk_showDetails180.Checked Then
            SplitContainer4.Panel2Collapsed = False
        Else
            SplitContainer4.Panel2Collapsed = True
        End If
    End Sub
    Private Sub grdEDT180TX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdEDT180TX.AfterRowActivate
        If grdEDT180TX.ActiveRow IsNot Nothing And grdEDT180TX.ActiveRow.IsDataRow And chk_showDetails180.Checked Then
            Dim selectedRow As Infragistics.Win.UltraWinGrid.UltraGridRow = grdEDT180TX.ActiveRow
            Dim EDI_DOC_SEQ_NO As String = selectedRow.Cells("EDI_DOC_SEQ_NO").Value.ToString()
            Fill_Records("EDT180X2", New Object() {EDI_DOC_SEQ_NO})
        End If
    End Sub

End Class