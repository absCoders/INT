Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class POFORDR1
    Dim rowPOTORDR1 As DataRow
    Dim rowAPTVEND1 As DataRow
    Dim rowICTITEM1 As DataRow

    Dim PO_ORDER_NO As String
    Dim VEND_CODE As String
    Dim ITEM_CODE As String

    Dim POTORDRX As String
    Dim POTORDRT As String
    Dim sqlPOTORDRT As String
    Dim POTORDRN As String
    Dim sqlPOTORDRN As String

    Dim PO_DATE_REQUIRED_last As String
    Dim PO_DATE_REQUESTED_last As String
    Dim WHSE_CODE_last As String
    Dim APPR_NOTES As String

    Dim VEND_BUYER_PURCH_LIMIT As Decimal = 0
    Dim preapproval_applied As Boolean = False
    Dim PO_PARM_PINV_LT As Integer
    Dim PO_PARM_PINV_PORT As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "POFORDRI" Then
            InquiryMode = True
        End If

        Get_PARM("GLTPARM1")
        Get_PARM("POTPARM1")

        PO_PARM_PINV_LT = Val(ROWs("POTPARM1").Item("PO_PARM_PINV_LT") & "")
        PO_PARM_PINV_PORT = ROWs("POTPARM1").Item("PO_PARM_PINV_PORT") & ""

        Load_POTORDRX()

        With dst

            ASCMAIN1.sql = "Select DPTPLAN1.*, ICTITEM1.ITEM_DESC" _
            & " from DPTPLAN1, ICTITEM1 where ICTITEM1.ITEM_CODE = DPTPLAN1.ITEM_CODE and DPTPLAN1.VEND_CODE = :PARM1"
            Create_TDA(.Tables.Add, "DPTPLAN1", "**", 0, True, "V", 1)
            With .Tables("DPTPLAN1")
                .Columns.Add("DELETE")
            End With

            ASCMAIN1.sql = "Select POTORDR1.*" _
            & " from POTORDR1 where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, True, "V", 1)

            ASCMAIN1.sql = "Select POTORDRX.*" & vbCrLf _
                & ", POTORDRT.PO_QTY_ORD, POTORDRT.PO_QTY_REC, POTORDRT.PO_QTY_INV, POTORDRT.PO_QTY_OPN" & vbCrLf _
                & ", POTORDRT.PO_AMT_ORD, POTORDRT.PO_AMT_REC, POTORDRT.PO_AMT_INV, POTORDRT.PO_AMT_OPN" & vbCrLf _
                & ", POTORDRN.PO_NINV_LINES, POTORDRN.PO_NINV_QTY_OPN, POTORDRN.PO_NINV_AMT_OPN" & vbCrLf _
                & $" from {POTORDRX} POTORDRX, {POTORDRT} POTORDRT, {POTORDRN} POTORDRN" & vbCrLf _
                & " where POTORDRT.PO_ORDER_NO (+) = POTORDRX.PO_ORDER_NO and POTORDRN.PO_ORDER_NO (+) = POTORDRX.PO_ORDER_NO"
            Create_TDA(.Tables.Add, "POTORDRX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select POTORDR2.*, ICTITEM1.ITEM_EAN_CODE" _
            & " from POTORDR2, ICTITEM1 where POTORDR2.ITEM_CODE = ICTITEM1.ITEM_CODE AND PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select ICTIREC2.*, ICTIREC1.RECEIPT_DATE, ICTIREC1.SOURCE_DOC_NO" _
            & " from ICTIREC2,ICTIREC1 where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO and ICTIREC2.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIRECX", "**", 0, False, "V", 2)

            ASCMAIN1.sql = "Select POTORDR9.*, ICTITEM1.ITEM_DESC" _
            & " from POTORDR9,ICTITEM1 where POTORDR9.PO_ORDER_NO = :PARM1" _
            & " and ICTITEM1.ITEM_CODE = POTORDR9.ITEM_CODE"
            Create_TDA(.Tables.Add, "POTORDR9", "**", 0, True, "V", 3)

            With .Tables("POTORDR2").Columns
                .Add("PO_QTY_OPN_CALC", GetType(System.Int64), "IIF(ISNULL(PO_QTY_REC,0)>ISNULL(PO_QTY_ORD,0) OR ISNULL(PO_STATUS,'')='C',0,ISNULL(PO_QTY_ORD,0)-ISNULL(PO_QTY_REC,0))")
                .Add("PO_AMT_ORD", GetType(System.Decimal), "ISNULL(PO_QTY_ORD,0) * ISNULL(PO_COST,0)")
                .Add("PO_AMT_REC", GetType(System.Decimal), "ISNULL(PO_QTY_REC,0) * ISNULL(PO_COST,0)")
                .Add("PO_AMT_INV", GetType(System.Decimal), "ISNULL(PO_QTY_INV,0) * ISNULL(PO_COST,0)")
                .Add("PO_AMT_OPN", GetType(System.Decimal), "ISNULL(PO_QTY_OPN,0) * ISNULL(PO_COST,0)")
                .Add("PO_AMT_OPN_CALC", GetType(System.Decimal), "ISNULL(PO_QTY_OPN_CALC,0) * ISNULL(PO_COST,0)")
            End With

            ASCMAIN1.sql = "Select POTORDR5.*" _
                & " from POTORDR5 where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR5", "**", 0, True, "V", 2)

            ASCMAIN1.sql = "Select POTORDR6.*" _
                & " from POTORDR6 where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDR6", "**", 0, True, "V", 3)

            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'POTORDR1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V")
            .Tables("TATEVNT1").Columns.Add("ATTACHMENT_EXT")

            ASCMAIN1.sql = "Select * from POTORDXR where PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "POTORDXR", "**", 0, True, "V")

            Create_TDA(.Tables.Add, "POTORDRH", "*", 1)

            Create_TDA(.Tables.Add, "ICTPORT2", "*", 0, False)
            Fill_Records("ICTPORT2")

            Create_TDA(.Tables.Add, "ICTCCLS1", "*", 0, False)
            Fill_Records("ICTCCLS1")

            .Tables.Add("POTORDRT")
            With .Tables("POTORDRT")
                .Columns.Add("STATUS")
                .Columns.Add("QTY_I", GetType(System.Int32))
                .Columns.Add("AMT_I", GetType(System.Decimal))
                .Columns.Add("QTY_N", GetType(System.Int32))
                .Columns.Add("AMT_N", GetType(System.Decimal))
                .Columns.Add("AMT_T", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "Select BM_ISSUE_NO, BM_ISSUE_DATE, BM_ISSUE_COMMENT, BM_ISSUE_TYPE, BM_ISSUE_VCOST" & vbCrLf _
                & " from BMTMAIN2 where BM_PROD_ITEM = :PARM1" & vbCrLf _
                & " and BM_ISSUE_NO <> '00'" & vbCrLf _
                & " and BM_ISSUE_TYPE = :PARM2"
            Create_TDA(.Tables.Add, "BMTMAIN2", "**", 0, False, "VV")


            ASCMAIN1.sql = "Select ICTIREC1.* from ICTIREC1 where PO_ORDER_NO = :PARM1"
            If ASCMAIN1.CLIENT = "INT" Then
                ASCMAIN1.sql = "Select ICTIREC1.* from ICTIREC1 where RECEIPT_NO in (Select Distinct RECEIPT_NO from ICTIREC2 where PO_ORDER_NO = :PARM1)"
            End If
            Create_TDA(.Tables.Add, "ICTIREC1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select ICTIREC2.*, ICTITEM1.ITEM_DESC, ICTIREC2.PO_COST * ICTIREC2.QTY_REC EXT_PO_COST " _
             & " from ICTIREC2,ICTITEM1 " _
             & " where ICTIREC2.ITEM_CODE = ICTITEM1.ITEM_CODE " _
             & "   and ICTIREC2.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIREC2", "**", 0, False, "V", 2)

            Create_Relation("ICTIREC1", "ICTIREC2", "RECEIPT_NO")

            .Tables("ICTIREC1").Columns.Add("EXT_PO_COST", GetType(System.Decimal), "SUM (CHILD.EXT_PO_COST)")

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO" & vbCrLf _
                & ", APTINVH1.INV_NUM, APTINVH1.INV_DATE, APTINVH1.INV_AMT" & vbCrLf _
                & ", APTINVH1.INV_STATUS, APTINVH1.CHECK_NUM, APTINVH1.CHECK_DATE" & vbCrLf _
                & " from APTINVH1 where APTINVH1.VOUCHER_NO in " & vbCrLf _
                & " (Select Distinct APTINVH5.VOUCHER_NO from APTINVH5,ICTIREC2" & vbCrLf _
                & " where ICTIREC2.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
                & "   and ICTIREC2.PO_ORDER_NO = :PARM1)"
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select X.*" & vbCrLf _
                & ", ICTIREC1.RECEIPT_DATE, ICTIREC1.QTY_REC, ICTIREC1.AMT_REC" & vbCrLf _
                & " from ICTIREC1,(" & vbCrLf _
                & "Select APTINVH5.VOUCHER_NO, APTINVH5.RECEIPT_NO" & vbCrLf _
                & ", SUM (NVL(APTINVH5.INV_QTY,0)) QTY_INV" & vbCrLf _
                & ", SUM (NVL(APTINVH5.INV_QTY,0) * NVL(APTINVH5.INV_COST,0)) AMT_INV" & vbCrLf _
                & ", SUM (NVL(APTINVH5.VAR_QTY,0)) VAR_QTY, SUM (NVL(APTINVH5.VAR_AMT,0)) VAR_AMT" & vbCrLf _
                & " from APTINVH5,ICTIREC1" & vbCrLf _
                & " where ICTIREC1.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
                & "   and ICTIREC1.PO_ORDER_NO = :PARM1" & vbCrLf _
                & " group by APTINVH5.VOUCHER_NO, APTINVH5.RECEIPT_NO) X" & vbCrLf _
                & " where ICTIREC1.RECEIPT_NO = X.RECEIPT_NO"
            Create_TDA(.Tables.Add, "APTINVH5_SUM", "**", 0, False, "V", 2)

            Create_Relation("APTINVH1", "APTINVH5_SUM", "VOUCHER_NO")


            ASCMAIN1.sql = "Select APTINVH5.*, APTINVH1.INV_NUM, APTINVH1.INV_DATE" & vbCrLf _
                & " from APTINVH5,ICTIREC1,APTINVH1" & vbCrLf _
                & " where ICTIREC1.RECEIPT_NO = APTINVH5.RECEIPT_NO" & vbCrLf _
                & "   and APTINVH1.VOUCHER_NO = APTINVH5.VOUCHER_NO" & vbCrLf _
                & "   and ICTIREC1.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "APTINVH5", "**", 0, False, "V", 2)

            Create_Relation("ICTIRECX", "APTINVH5", "RECEIPT_NO, RECEIPT_LNO")

            ASCMAIN1.sql = "SELECT
                ICTPINV1.PINV_NO, ICTPINV1.WHSE_CODE,ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV1.PO_ORDER_NO, ICTPINV2.ITEM_CODE, SUM(ICTPINV2.PINV_QTY) INV_QTY
                FROM
                ICTPINV1, ICTPINV2
                WHERE
                ICTPINV1.PINV_NO = ICTPINV2.PINV_NO
                and ICTPINV1.PINV_STATUS = 'O'
                AND ICTPINV2.PO_ORDER_NO = :PARM1
                GROUP BY ICTPINV1.PINV_NO, ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV1.PO_ORDER_NO, ICTPINV2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTPINV1", "**", 0, False, "V", 6)

            ASCMAIN1.sql = "Select ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO, ICTPINV2.PINV_LNO
                , ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE
                , ICTPINV2.PINV_QTY INV_QTY, ICTITEM1.PORT_CODE
                 from ICTPINV1, ICTPINV2, ICTITEM1
                 where ICTPINV1.PINV_NO = ICTPINV2.PINV_NO
                 And ICTPINV1.PINV_STATUS = 'O'
And ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE
                AND ICTPINV2.PO_ORDER_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTPINVD", "**", 0, False, "V", 7)
            With dst.Tables("ICTPINVD")
                .Columns("VESSEL_NAME").AllowDBNull = True
                .Columns("INV_DATE").AllowDBNull = True

                .Columns.Add("ETA_DATE_DC", GetType(System.DateTime))
                .Columns.Add("OPO_QTY", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO
                , ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE
                , ICTPINV2.PINV_QTY INV_QTY, ICTITEM1.PORT_CODE, ICTPINV2.RECEIPT_NO, ICTPINV2.RECEIPT_LNO
                 from ICTPINV1, ICTPINV2, ICTITEM1
                 where ICTPINV1.PINV_NO = ICTPINV2.PINV_NO
                 And ICTPINV1.PINV_STATUS <> 'C'
And ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE
                AND ICTPINV2.PO_ORDER_NO = :PARM1 AND ICTPINV1.RECEIPT_NO IS NOT NULL"
            Create_TDA(.Tables.Add, "ICTPINVR", "**", 0, False, "V", 6)
            With dst.Tables("ICTPINVR")
                .Columns("VESSEL_NAME").AllowDBNull = True
                .Columns("INV_DATE").AllowDBNull = True

                .Columns.Add("ETA_DATE_DC", GetType(System.DateTime))
                .Columns.Add("OPO_QTY", GetType(System.Int32))
            End With
            Create_Relation("ICTIRECX", "ICTPINVR", "RECEIPT_NO, RECEIPT_LNO")

            'Create_Relation("POTORDR2", "ICTPINVD", "PO_ORDER_NO, PO_ORDER_LNO")

        End With

        grdDPTPLAN1.DataSource = dst.Tables("DPTPLAN1")
        grdPOTORDRX.DataSource = dst.Tables("POTORDRX")
        grdPOTORDRH.DataSource = dst.Tables("POTORDRH")
        grdPOTORDR2.DataSource = dst.Tables("POTORDR2")
        grdICTIRECX.DataSource = dst.Tables("ICTIRECX")
        grdPOTORDR9.DataSource = dst.Tables("POTORDR9")
        grdPOTORDR5.DataSource = dst.Tables("POTORDR5")
        grdPOTORDR6.DataSource = dst.Tables("POTORDR6")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")
        grdPOTORDXR.DataSource = dst.Tables("POTORDXR")
        grdPOTORDRT.DataSource = dst.Tables("POTORDRT")

        grdICTIREC1.DataSource = dst.Tables("ICTIREC1")
        grdAPTINVH1.DataSource = dst.Tables("APTINVH1")
        grdICTPINV1.DataSource = dst.Tables("ICTPINV1")
        grdICTPINVD.DataSource = dst.Tables("ICTPINVD")

        Create_Summary(grdPOTORDR2, "PO_ORDER_LNO", "Count")
        Create_Summary(grdPOTORDR2, New String() {"PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_INV", "PO_QTY_OPN", "PO_QTY_OPN_CALC", "PO_AMT_ORD", "PO_QTY_BACKORDER"})

        Create_Summary(grdICTIRECX, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIRECX, New String() {"QTY_REC", "EXT_COST_MATLS", "TRAN_PV", "TRAN_MV", "TRAN_CV"})

        Create_Summary(grdPOTORDR9, "ITEM_CODE", "Count")
        Create_Summary(grdPOTORDR9, New String() {"PO_QTY_COM"})

        Create_Summary(grdPOTORDRX, "PO_ORDER_NO", "Count")

        Create_Summary(grdICTIREC1, "RECEIPT_NO", "Count")
        Create_Summary(grdICTIREC1, New String() {"QTY_REC", "AMT_REC", "QTY_INV", "AMT_INV", "EXT_PO_COST"})

        Create_Summary(grdAPTINVH1, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVH1, New String() {"INV_AMT"})

        Create_Summary(grdDPTPLAN1, "PLAN_NO", "Count")
        Create_Summary(grdDPTPLAN1, "QTY_PLANNED")


        Create_Summary(grdICTPINVD, "INV_NUM", "Count")
        Create_Summary(grdICTPINVD, New String() {"INV_QTY", "OPO_QTY"})
        With grdICTPINVD.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.LightGray
                If gcol.Key = "INV_QTY" Or gcol.Key = "OPO_QTY" Then
                    gcol.Format = "###,##0"
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                Else

                End If
            Next

        End With

        Create_Summary(grdICTPINV1, "INV_NUM", "Count")
        Create_Summary(grdICTPINV1, New String() {"INV_QTY"})
        With grdICTPINV1.DisplayLayout.Bands(0)
            .Columns("PO_ORDER_NO").Hidden = True
            .Columns("INV_QTY").Format = "###,##0"
            .Columns("INV_QTY").Header.Appearance.BackColor = Color.White
            .Columns("INV_QTY").Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            .Columns("INV_QTY").Header.Appearance.BackColor2 = Color.LightBlue
        End With

        With grdPOTORDRX.DisplayLayout.Bands("POTORDRX")
            .Columns("PO_ORDER_NO").Header.Fixed = True
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VEND_NAME").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"PO_QTY_ORD", "PO_QTY_REC", "PO_QTY_INV", "PO_QTY_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 70
                    gcol.Format = "#,##0"
                    Create_Summary(grdPOTORDRX, gcol.Key)
                ElseIf New String() {"PO_AMT_ORD", "PO_AMT_REC", "PO_AMT_INV", "PO_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Width = 80
                    gcol.Format = "#,##0"
                    Create_Summary(grdPOTORDRX, gcol.Key)
                ElseIf New String() {"PO_DATE_SHIPPED", "PO_DATE_ETA", "PO_SHIP_VESSEL", "CONTAINER_NO"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Turquoise
                ElseIf gcol.Key.StartsWith("PO_APPR") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                ElseIf gcol.Key.StartsWith("PO_XMIT") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Orange
                ElseIf New String() {"PO_NINV_LINES", "PO_NINV_QTY_OPN", "PO_NINV_AMT_OPN"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    gcol.Format = "#,##0"
                    If gcol.Key = "PO_NINV_AMT_OPN" Then gcol.Format = "#,##0.00"
                    Create_Summary(grdPOTORDRX, gcol.Key)
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdPOTORDR5.DisplayLayout.Bands("POTORDR5")
            If ASCMAIN1.CLIENT = "INT" Then
                .Columns("PO_NINV_PRICE").MaskInput = "nnnn.nnnn"
                .Columns("PO_NINV_PRICE").Format = "#.0000"
            End If
        End With



        With grdPOTORDR2.DisplayLayout.Bands("POTORDR2")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                'If New String() {"RECEIPT_NO", "RECEIPT_LNO", "RECEIPT_DATE", "QTY_REC"}.Contains(gcol.Key) Then
                'gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                'End If
                If New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
            .Columns("PO_ORDER_LNO").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
            .Columns("PO_COST").Format = "#.0000"
            .Columns("PO_COST").MaskInput = "nnnnnn.nnnnnn"
            For Each COLUMN_NAME As String In New String() {"PO_ORDER_LNO", "ITEM_DESC", "ITEM_UOM", "PO_QTY_OPN_CALC", "PO_AMT_ORD", "BM_ISSUE_NO", "BM_ISSUE_DATE", "PO_QTY_CXL", "ITEM_CODE_ALT", "ITEM_EAN_CODE"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            For Each COLUMN_NAME As String In New String() {"PO_DATE_ETD", "PO_DATE_ETD_NOTES"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next
            .Columns("ITEM_CODE_ALT").Hidden = Not (ASCMAIN1.CLIENT = "INT")

        End With

        For Each COLUMN_NAME As String In New String() _
        {"PO_QTY_ORD", "PO_QTY_OPN_CALC", "PO_QTY_BACKORDER"}
            With grdPOTORDR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .CellAppearance.BackColor = Drawing.Color.Yellow
            End With
        Next
        For Each COLUMN_NAME As String In New String() _
        {"PO_DATE_BACKORDER", "PO_QTY_BACKORDER"}
            With grdPOTORDR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                .Hidden = True
            End With
        Next

        With grdICTIRECX.DisplayLayout.Bands("ICTIRECX")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                'If New String() {"RECEIPT_NO", "RECEIPT_LNO", "RECEIPT_DATE", "QTY_REC"}.Contains(gcol.Key) Then
                'gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                'End If
            Next
            For Each COLUMN_NAME As String In New String() {"RECEIPT_NO", "RECEIPT_LNO", "RECEIPT_DATE", "QTY_REC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdICTIRECX.DisplayLayout.Bands("ICTIRECX_APTINVH5")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
            'For Each COLUMN_NAME As String In New String() {"RECEIPT_NO", "RECEIPT_LNO", "RECEIPT_DATE", "QTY_REC"}
            '    .Columns(COLUMN_NAME).Header.Fixed = True
            'Next
        End With

        With grdICTIRECX.DisplayLayout.Bands("ICTIRECX_ICTPINVR")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
            'For Each COLUMN_NAME As String In New String() {"RECEIPT_NO", "RECEIPT_LNO", "RECEIPT_DATE", "QTY_REC"}
            '    .Columns(COLUMN_NAME).Header.Fixed = True
            'Next
        End With

        With grdICTIREC1.DisplayLayout.Bands("ICTIREC1")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                If New String() {"QTY_REC", "AMT_REC", "EXT_PO_COST"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"QTY_INV", "AMT_INV"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                End If
            Next
        End With
        With grdICTIREC1.DisplayLayout.Bands("ICTIREC1_ICTIREC2")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                If New String() {"QTY_REC", "AMT_REC", "EXT_PO_COST"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"QTY_INV", "INV_COST", "TRAN_PV"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                End If
            Next
        End With


        With grdAPTINVH1.DisplayLayout.Bands("APTINVH1")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                If New String() {"CHECK_NUM", "CHECK_DATE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Gold
                End If
            Next
        End With

        With grdAPTINVH1.DisplayLayout.Bands("APTINVH1_APTINVH5_SUM")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                If New String() {"QTY_REC", "AMT_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"QTY_INV", "AMT_INV"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                End If
            Next
        End With



        With grdDPTPLAN1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackColor2 = Color.LightBlue
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "DELETE" Or gcol.Key = "DATE_REQUIRED" Or gcol.Key = "QTY_PLANNED" Then
                    gcol.CellActivation = Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                Else
                    gcol.CellActivation = Activation.NoEdit
                End If
            Next
        End With
        Show_Filter(grdDPTPLAN1, True)

        MakeTransparent(chkEditPlans)

        Set_Read_Only(grpVendor, True)


        ASCMAIN1.Add_Value_List(grdAPTINVH1, "INV_STATUS", , New String() {":", "H:Hold", "O:Open", "P:Paid", "D:Deleted"})
        ASCMAIN1.Add_Value_List(grdPOTORDRX, "PO_STATUS", , New String() {":", "O:Open", "C:Closed", "D:Deleted"})
        ASCMAIN1.Add_Value_List(grdPOTORDR2, "PO_STATUS", , New String() {":", "O:Open", "C:Closed"})

        Dim rowPOTBUYR1 As DataRow = LookUp("POTBUYR1", ASCMAIN1.USER_ID)
        If rowPOTBUYR1 IsNot Nothing Then
            VEND_BUYER_PURCH_LIMIT = Val(rowPOTBUYR1.Item("VEND_BUYER_PURCH_LIMIT") & "")
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "New"
                Validate_Vendor()
                'Validate_Code("VEND_CODE")

                If Absx1.dteFor("PO_DATE_ORDERED").DateTime & "" = "" Then
                    EMsg &= vbCr & "PO Date is Required"
                End If
                If optPO_TYPE.Value & "" = "" Then
                    EMsg &= vbCr & "You must choose Make or Buy for PO Type"
                End If
                If EMsg = "" Then
                    If rowAPTVEND1.Item("VEND_WHSE_CODE") & "" = "" And optPO_TYPE.Value = "M" Then
                        EMsg &= vbCr & "Cannot start a Make PO with Vendor " & VEND_CODE & vbCr & " because there is no Warehouse Associated with " & VEND_CODE
                    End If
                    If rowAPTVEND1.Item("VEND_COUNTRY") & "" = "" Then
                        EMsg &= vbCr & "Cannot start a PO with a Vendor (" & VEND_CODE & ") without a Country"
                    End If
                End If

                If Absx1.chkFor("PO_DISASSEMBLY_IND").Checked Then
                    If Absx1.optFor("PO_TYPE").Value & "" = "M" Then
                        EMsg &= vbCr & "Dis-Assembly may not be selected for Make POs (Current State)"
                    End If
                End If
                ' email to lbm 11/15
                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock("POTORDR1_V", VEND_CODE) Then Exit Sub
                'End If

                If Absx1.optFor("PO_ORDER_TYPE").Value & "" = "" Then
                    EMsg &= vbCr & "PO Order Type must be specified"
                ElseIf Absx1.optFor("PO_ORDER_TYPE").Value & "" = "R" Then
                    If optPO_TYPE.Value & "" <> "M" Then
                        EMsg &= vbCr & "PO Order Type Re-Work must be type Make"
                    End If
                    If chkDisAssembly.Checked Then
                        EMsg &= vbCr & "PO Order Type Re-Work may not be associated with Dis-Assembly"
                    End If
                End If



            Case "Edit", "View"
                If Absx1.txtFor("PO_ORDER_NO").Text = "" Then
                    EMsg &= vbCr & "No PO No Specified"
                Else
                    PO_ORDER_NO = Absx1.txtFor("PO_ORDER_NO").Text
                    rowPOTORDR1 = LookUp("POTORDR1", PO_ORDER_NO)
                    If rowPOTORDR1 Is Nothing Then
                        EMsg &= vbCr & "Invalid PO No (" & PO_ORDER_NO & ")"
                    Else
                        VEND_CODE = rowPOTORDR1.Item("VEND_CODE")
                    End If
                End If

                If EMsg = "" Then
                    Dim PO_STATUS As String = rowPOTORDR1.Item("PO_STATUS") & ""

                    If optStatus.Value = "a" Then
                        If PO_STATUS = "C" Then
                            EMsg &= vbCr & "PO No (" & PO_ORDER_NO & ") is no longer Open - no need to approve"
                        End If
                    End If

                    If eItemKey = "Edit" Then


                        If PO_STATUS = "C" Then
                            If MsgBox("PO " & PO_ORDER_NO & " has been Closed" _
                                      & vbCrLf & "Do you want to Re-Open it?",
                                      MsgBoxStyle.YesNo, "Option to Re-Open a Closed PO") = MsgBoxResult.Yes Then
                                ASCMAIN1.sql = "Update POTORDR1 Set PO_STATUS = 'O' where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_STATUS = 'C'"
                                ASCDATA1.ExecuteSQL()

                                ' ASCDATA1.ExecuteSQL("Update POTORDR2 set PO_QTY_OPN = NVL(PO_QTY_OPN,0) + NVL(PO_QTY_CXL,0) where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_STATUS = 'C'")
                                ASCDATA1.ExecuteSQL("Update POTORDR2 set PO_QTY_CXL = 0, PO_STATUS = 'O' where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_STATUS = 'C'")

                                'ASCMAIN1.sql = "Update POTORDR2 Set PO_STATUS = 'O' where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_STATUS = 'C'"
                                'ASCDATA1.ExecuteSQL()

                                ' NEED TO LEAVE PO_QTY_OPN ALONE UNLESS YOU CODE A FIX TO ICTSTAT2

                                TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "REOPEN", "PO Re-Opened")
                                '  Record_Event("RE_OPEN", "PO Re-Opened")
                                rowPOTORDR1 = LookUp("POTORDR1", PO_ORDER_NO)
                                PO_STATUS = rowPOTORDR1.Item("PO_STATUS") & ""
                            End If
                        End If

                        If PO_STATUS <> "O" Then


                            Select Case PO_STATUS
                                Case "C"
                                    EMsg &= EMsg & vbCr & "PO No " & PO_ORDER_NO & " has been Cancelled or Closed"
                                Case "D"
                                    EMsg &= EMsg & vbCr & "PO No " & PO_ORDER_NO & " has been Deleted"
                                Case Else
                                    EMsg &= EMsg & vbCr & "PO No " & PO_ORDER_NO & " is No Longer Open"
                            End Select
                        End If
                    End If

                    'If rowPOTORDR1.Item("PO_AUTO_GEN") & "" <> "" Then
                    '    EMsg = EMsg & vbCr & "Changes are Not Permitted to Automatically Generated PO's"
                    'End If
                End If

                If EMsg = "" Then
                    PO_ORDER_NO = Absx1.txtFor("PO_ORDER_NO").Text
                    If eItemKey = "Edit" Then
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                        ' email to lbm 11/15
                        'If Not ASCMAIN1.Logical_Lock("POTORDR1_V", VEND_CODE) Then Exit Sub
                    End If
                End If

            Case "Update"

                If EntryMode = "E" Then
                    'Dim always As Boolean = True ' always ask for revision reason
                    'If (Absx1.txtFor("PO_REVISION_NOTE").Text = _
                    'rowPOTORDR1.Item("PO_REVISION_NOTE", DataRowVersion.Original) & "" _
                    'Or Absx1.txtFor("PO_REVISION_NOTE").Text = "") _
                    'And (always Or rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1") Then
                    '    EMsg &= vbCr & "You Must Enter the Reason for Changing this PO in the Reason for Revision"
                    'End If
                    'tabMain.SelectedTab = tabMain.Tabs("Totals")
                    'tabNotes.SelectedTab = tabNotes.Tabs("Reason for Revision")

                    If rowPOTORDR1.Item("PO_PRINTED_IND", DataRowVersion.Original) & "" = "1" And Absx1.txtFor("PO_REVISION_NOTE").Text = "" Then
                        EMsg &= vbCr & "You Must Enter the Reason for Changing this PO in the Reason for Revision"
                    End If
                    tabMain.SelectedTab = tabMain.Tabs("Totals")
                    tabNotes.SelectedTab = tabNotes.Tabs("Reason for Revision")
                End If

                'If EntryMode = "N" Then
                '    If Absx1.txtFor("BUYER_CODE").Text = "" Then
                '        EMsg &= vbCr & "You Must Specify the Buyer"
                '    End If
                'End If

                If EMsg = "" Then
                    If Absx1.dteFor("PO_DATE_REQUIRED").Value & "" = "" _
                    Or Absx1.dteFor("PO_DATE_CANCEL").Value & "" = "" Then
                        EMsg = EMsg & vbCr & "Date Required (Confirmed) & Cancel-By Date are Required"
                    Else
                        If Format(Absx1.dteFor("PO_DATE_REQUIRED").Value, "yyyyMMdd") >
                           Format(Absx1.dteFor("PO_DATE_CANCEL").Value, "yyyyMMdd") Then
                            EMsg = EMsg & vbCr & "Cancel Date cannot be Prior to Date Required (Confirmed)"
                        End If
                        If Format(Absx1.dteFor("PO_DATE_REQUIRED").Value, "yyyyMMdd") <
                           Format(Absx1.dteFor("PO_DATE_ORDERED").Value, "yyyyMMdd") Then
                            EMsg = EMsg & vbCr & "Date Required (Confirmed) cannot be Prior to Order Date"
                        End If
                    End If
                End If

                Validate_Code("WHSE_CODE")
                Validate_Code("MARKET_CODE")
                Validate_Code("TERM_CODE")
                '  Validate_Code("FRT_TERMS")
                Validate_Code("FRT_CLASS_CODE",, True)
                Validate_Code("TRF_CLASS_CODE",, True)

                If Absx1.txtFor("VEND_WHSE_CODE").Text <> "" Then
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("VEND_WHSE_CODE").Text)
                    If rowICTWHSE1 Is Nothing Then
                        EMsg &= vbCr & "Vendor Whse Specified is an Invalid Whse Code"
                    End If
                End If

                If dst.Tables("POTORDR2").Select("").Length = 0 AndAlso dst.Tables("POTORDR5").Select("").Length = 0 Then
                    EMsg &= vbCr & "No PO Details"
                Else

                    Dim ITEM_CODEs_on_po As New List(Of String)
                    Dim ITEM_CODEs_on_po_more_than_once As New List(Of String)
                    For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                        Dim ITEM_CODE As String = rowPOTORDR2.Item("ITEM_CODE")
                        Dim PO_STATUS As String = rowPOTORDR2.Item("PO_STATUS")
                        Dim BM_ISSUE_NO As String = rowPOTORDR2.Item("BM_ISSUE_NO") & ""
                        Dim BM_ISSUE_SEL As String = rowPOTORDR2.Item("BM_ISSUE_SEL") & ""

                        If ITEM_CODEs_on_po.Contains(ITEM_CODE) Then
                            ' for use as a warning that AN ITEM IS DUPLICATED ON THE ORDER
                            ITEM_CODEs_on_po_more_than_once.Add(ITEM_CODE)
                        Else
                            ITEM_CODEs_on_po.Add(ITEM_CODE)
                        End If

                        If PO_STATUS = "O" Then
                            If BM_ISSUE_NO <> "" Then
                                Dim row As DataRow = LookUp("BMTMAIN2", New String() {ITEM_CODE, BM_ISSUE_NO})
                                If row Is Nothing Then
                                    EMsg &= vbCr & $"Invalid BM Issue {BM_ISSUE_NO} selected for Item {ITEM_CODE}"
                                End If
                            Else
                                If BM_ISSUE_SEL = "1" Then
                                    EMsg &= vbCr & $"You must select a valid BM Issue even when selecting to use the most current issue for Item {ITEM_CODE}"
                                End If
                            End If
                        End If

                        If ASCMAIN1.CLIENT = "INT" Then
                            Dim WHSE_CODE2 As String = rowPOTORDR2.Item("WHSE_CODE") & ""
                            If WHSE_CODE2 <> "" Then

                            End If
                            If (Absx1.txtFor("WHSE_CODE").Text = "CLA" And WHSE_CODE2 <> "CLA") _
                            Or (Absx1.txtFor("WHSE_CODE").Text <> "CLA" And WHSE_CODE2 = "CLA") Then
                                EMsg &= vbCr & "Cannot Mix CLA Warehouse with other warehouses in same PO"
                            End If
                        End If
                    Next

                    If ITEM_CODEs_on_po_more_than_once.Count > 0 Then
                        Dim z As String = "Items duplicated on this PO: " & Join(ITEM_CODEs_on_po_more_than_once.ToArray, ",")
                        If ASCMAIN1.CLIENT = "INT" Then
                            If optPO_ORDER_TYPE.Value = "R" Then
                                EMsg &= vbCr & "Items duplicated on this PO (not permitted on a Re-Work PO): " & Join(ITEM_CODEs_on_po_more_than_once.ToArray, ",")
                            Else
                                If MsgBox(z & vbCrLf & vbCrLf & "Continue with Update?", MsgBoxStyle.YesNo, "Duplicate Items") = MsgBoxResult.No Then
                                    EMsg &= vbCr & "Items duplicated on this PO: " & Join(ITEM_CODEs_on_po_more_than_once.ToArray, ",")
                                End If
                            End If
                        Else
                            If MsgBox(z & vbCrLf & vbCrLf & "Continue with Update?", MsgBoxStyle.YesNo, "Duplicate Items") = MsgBoxResult.No Then
                                EMsg &= vbCr & "Items duplicated on this PO: " & Join(ITEM_CODEs_on_po_more_than_once.ToArray, ",")
                            End If
                        End If
                    End If

                    If EMsg = "" Then
                        If ASCMAIN1.CLIENT = "INT" Then
                            If Absx1.txtFor("WHSE_CODE").Text = "CLA" And (optPO_TYPE.Value = "B" Or Absx1.txtFor("VEND_WHSE_CODE").Text <> "CLA") Then
                                Dim row As DataRow = LookUp("APTVEND1", VEND_CODE)
                                If row.Item("VEND_SUPPLIER_ID") & "" = "" Then
                                    EMsg &= vbCr & "Vendor " & VEND_CODE & " does not have a Clarins Vendor ID - and this is required if PO ships to CLA"
                                End If
                            End If
                        End If
                    End If
                End If

                Dim line_count_with_BM As Integer = dst.Tables("POTORDR2").Select("ISNULL(BM_ISSUE_NO,'') <> '' or ISNULL(BM_ISSUE_SEL,'0') = '1'").Length
                Dim line_count As Integer = dst.Tables("POTORDR2").Select("").Length
                If line_count_with_BM > 0 And Absx1.txtFor("VEND_WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "There are MAKE items on PO, but Vendor does NOT have a Warehouse"
                Else
                    If line_count_with_BM = 0 And Absx1.txtFor("VEND_WHSE_CODE").Text <> "" Then
                        EMsg &= vbCr & "There are no items with a valid BM on PO, but Vendor Warehouse is Specified (implying Production)."
                    Else
                        If line_count_with_BM <> 0 And line_count <> line_count_with_BM Then
                            EMsg &= vbCr & "Cannot mix Make & Buy Items on Same PO"
                        End If
                    End If
                End If

                ' This is done several lines above, Have items in POTORDR5 but not POTORDR2.
                'If dst.Tables("POTORDR2").Select("").Length = 0 Then
                '    EMsg &= vbCr & "No Details for PO"
                'End If

                EMsg &= TAC.ICCMAIN1.Check_Standard_Cost_Initialization(Me, "POTORDR2")

                preapproval_applied = False


                Dim ITEM_CODEs As New List(Of String)
                For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("", "ITEM_CODE")
                    Dim ITEM_CODE As String = rowPOTORDR2.Item("ITEM_CODE")
                    If Not ITEM_CODEs.Contains(ITEM_CODE) Then
                        ITEM_CODEs.Add(ITEM_CODE)
                    End If
                Next

                If Absx1.txtFor("FRT_CLASS_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify a Freight Class Code for this Purchase" & vbCrLf & $" suggestions from Item Std Costs: {Join(Get_Class_Codes("ITEM_COST_FRT_CLASS", ITEM_CODEs), "','")}"
                End If

                If Absx1.txtFor("TRF_CLASS_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify a Tariff Class Code for this Purchase" & vbCrLf & $" suggestions from Item Std Costs: {Join(Get_Class_Codes("ITEM_COST_TRF_CLASS", ITEM_CODEs), "','")}"
                End If

                If optPO_ORDER_TYPE.Value = "R" Then
                    'If dst.Tables("POTORDR2").Select("BM_ISSUE_SEL = '1'").Length > 0 Then
                    '    EMsg &= vbCr & "Cannot select 'Curr BM' for a Re-Work PO - you must select a specific Re-Work BM"
                    'End If
                    For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                        Dim PO_ORDER_LNO As Int32 = Val(rowPOTORDR2.Item("PO_ORDER_LNO") & "")
                        Dim ITEM_CODE As String = rowPOTORDR2.Item("ITEM_CODE") & ""
                        Dim BM_ISSUE_NO As String = rowPOTORDR2.Item("BM_ISSUE_NO") & ""
                        If BM_ISSUE_NO = "" Then
                            EMsg &= vbCr & $"No Re-Work BM Issue Selected for Item {ITEM_CODE} on Line {CStr(PO_ORDER_LNO)}"
                        Else
                            Dim rowBMTMAIN2 As DataRow = LookUp("BMTMAIN2", New String() {ITEM_CODE, BM_ISSUE_NO})
                            If rowBMTMAIN2 Is Nothing Then
                                EMsg &= vbCr & $"Invalid BM Issue Selected ({BM_ISSUE_NO}) for Item {ITEM_CODE} on Line {CStr(PO_ORDER_LNO)}"
                            Else
                                Dim BM_ISSUE_TYPE As String = rowBMTMAIN2.Item("BM_ISSUE_TYPE") & ""
                                If BM_ISSUE_TYPE <> "R" Then
                                    EMsg &= vbCr & $"Invalid Type ({BM_ISSUE_TYPE}) for BM Issue Selected ({BM_ISSUE_NO}) for Item {ITEM_CODE} on Line {CStr(PO_ORDER_LNO)}"
                                End If
                            End If
                        End If
                    Next
                End If

                If EMsg = "" Then
                    Dim PO_AMT_ORD As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
                    Dim PO_NINV_AMOUNT As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMOUNT)", "") & "")
                    Dim PO_TOTAL_AMT As Decimal = PO_AMT_ORD + PO_NINV_AMOUNT

                    If ROWs("POTPARM1").Item("PO_PARM_APPR_REQD") & "" = "1" Then

                        APPR_NOTES = ""
                        If chkReadyForApproval.Checked Then
                            Dim PO_APPR_AMOUNT As Decimal = Val(rowPOTORDR1.Item("PO_APPR_AMOUNT") & "")
                            Dim VEND_BUYER_PURCH_LIMIT_BUFPCT As Decimal = 0
                            Dim VEND_BUYER_PURCH_LIMIT_BUFAMT As Decimal = 0
                            If rowPOTORDR1.Item("PO_APPR_BY") & "" <> "" Then
                                Dim rowPOTBUYR1 As DataRow = LookUp("POTBUYR1", rowPOTORDR1.Item("PO_APPR_BY"))
                                If rowPOTBUYR1 IsNot Nothing Then
                                    VEND_BUYER_PURCH_LIMIT_BUFPCT = Val(rowPOTBUYR1.Item("VEND_BUYER_PURCH_LIMIT_BUFPCT") & "")
                                    VEND_BUYER_PURCH_LIMIT_BUFAMT = Val(rowPOTBUYR1.Item("VEND_BUYER_PURCH_LIMIT_BUFAMT") & "")
                                End If
                            End If
                            Dim BUFAMT As Decimal = VEND_BUYER_PURCH_LIMIT_BUFPCT * PO_APPR_AMOUNT / 100
                            If BUFAMT > VEND_BUYER_PURCH_LIMIT_BUFAMT Then
                                BUFAMT = VEND_BUYER_PURCH_LIMIT_BUFAMT
                            End If
                            If rowPOTORDR1.Item("PO_APPR_BY") & "" <> "" And PO_APPR_AMOUNT > 0 _
                                And PO_TOTAL_AMT <= PO_APPR_AMOUNT + BUFAMT Then
                                APPR_NOTES = "Previously Approved for " & Format(Val(rowPOTORDR1.Item("PO_APPR_AMOUNT") & ""), "$#,##0.00")
                                preapproval_applied = True
                                'rowPOTORDR1.Item("PO_APPR_DATE") = rowPOTORDR1.Item("PO_APPR_DATE", DataRowVersion.Original)
                                'rowPOTORDR1.Item("PO_APPR_BY") = rowPOTORDR1.Item("PO_APPR_BY", DataRowVersion.Original)
                                'rowPOTORDR1.Item("PO_APPR_AMOUNT") = rowPOTORDR1.Item("PO_APPR_AMOUNT", DataRowVersion.Original)
                                'rowPOTORDR1.Item("PO_APPR_NOTES") = rowPOTORDR1.Item("PO_APPR_NOTES", DataRowVersion.Original)
                            Else
                                Dim PO_PARM_APPR_LIMIT As Decimal = Val(ROWs("POTPARM1").Item("PO_PARM_APPR_LIMIT") & "")
                                If PO_TOTAL_AMT <= PO_PARM_APPR_LIMIT Or PO_TOTAL_AMT <= VEND_BUYER_PURCH_LIMIT Then
                                    Seek_Approval()
                                End If
                            End If

                            If APPR_NOTES = "" Then
                                rowPOTORDR1.Item("PO_APPR_DATE") = DBNull.Value
                                rowPOTORDR1.Item("PO_APPR_BY") = DBNull.Value
                                rowPOTORDR1.Item("PO_APPR_AMOUNT") = DBNull.Value
                                rowPOTORDR1.Item("PO_APPR_NOTES") = DBNull.Value
                            End If
                        End If
                    End If
                End If

            Case "Delete"
                ASCMAIN1.sql = "Select Count (*) from ICTIREC2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                If Val(ASCDATA1.GetDataValue) > 0 Then
                    EMsg &= vbCr & "You May Not Delete an Order which has been Received or Invoiced"
                End If
                ASCMAIN1.sql = "Select Count (*) from ICTPINV1 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                If Val(ASCDATA1.GetDataValue) > 0 Then
                    EMsg &= vbCr & "You May Not Delete an Order which has been entered in the Receiving Advice screen"
                End If

                If EMsg = "" Then
                    If MsgBox("Do you want to Delete this Order",
                              MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                              "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    If InputBox("Enter the word DELETE using CAPITAL LETTERS to Proceed", "Verification") <> "DELETE" Then
                        Exit Sub
                    End If
                End If

            Case "Cancel Order"
                Dim msg As String = ""
                ASCMAIN1.sql = "Select Count (*) from ICTIREC2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                If Val(ASCDATA1.GetDataValue) > 0 Then
                    msg = "Do you want to cancel the remaining open balance on this Order"
                Else
                    msg = "Do you wish to cancel this order"
                End If

                If MsgBox(msg, MsgBoxStyle.Question + MsgBoxStyle.YesNo,
                          "Confirmation") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Transmit"
                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                    Exit Sub
                End If

                Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                If rowPOTORDR1.Item("PO_STATUS") & "" <> "O" _
                    Or rowPOTORDR1.Item("PO_APPR_BY") & "" = "" _
                          Or rowPOTORDR1.Item("PO_XMIT_IND") & "" = "1" _
                    Or rowPOTORDR1.Item("PO_APPR_PENDING") & "" = "0" Then
                    EMsg &= vbCr & "PO " & PO_ORDER_NO & " is No Longer Pending Transmit"
                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Re-Transmit"

            Case "Approve"
                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                    Exit Sub
                End If


                Seek_Approval()

                If APPR_NOTES = "" Then
                    'ASCMAIN1.MultiTask_Release()
                    Exit Sub
                End If
                'PO TOTAL BELOW DOES NOT INCLUDE NINV
                'If MsgBox("Total Purchase Order Amount is " & Format(PO_AMT_ORD, "$#,##0.00") _
                '          & vbCrLf & "Total Amount Open to Ship is " & Format(PO_AMT_OPN, "$#,##0.00") _
                '          & vbCrLf & vbCrLf & "Terms are " & Absx1.txtFor("TERM_DESC").Text _
                '          & vbCrLf & vbCrLf & "Total Units Ordered are " & Format(PO_QTY_ORD, "#,##0") _
                '          & vbCrLf & "Total Units Open are " & Format(PO_QTY_OPN, "#,##0") _
                '          & vbCrLf & vbCrLf & "Arrival Date Range is " & Format(PO_DATE_REQUIRED_MIN, "MM/dd/yy") & " thru " & Format(PO_DATE_REQUIRED_MAX, "MM/dd/yy") _
                '          & vbCrLf & vbCrLf & "OK To Approve this Purchase?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                '    ASCMAIN1.MultiTask_Release()
                '    Exit Sub
                'End If

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

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                If APPR_NOTES <> "" Then
                    Update_Approval()
                End If
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Cancel Order"
                Cancel_Order()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "email"
                email_PO()

            Case "Transmit"
                Transmit_PO()
                ' ASCMAIN1.MultiTask_Release()
                Mode_Settings(False)

            Case "Re-Transmit"
                Transmit_PO(True)
                Mode_Settings(False)

            Case "Approve"
                Update_Approval()
                'ASCMAIN1.MultiTask_Release()
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If EntryMode = "V" And Not InquiryMode Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel Order").Settings.Enabled = iScreenMode
                    .Items("Approve").Settings.Enabled = iScreenMode
                    .Items("email").Settings.Enabled = iScreenMode
                    .Items("Transmit").Settings.Enabled = iScreenMode
                    .Items("Re-Transmit").Settings.Enabled = iScreenMode
                    .Items("Status Updates").Settings.Enabled = iScreenMode

                    .Items("New").Visible = (Not InquiryMode)
                    .Items("Edit").Visible = (Not InquiryMode)
                    .Items("Done").Visible = Not (EntryMode = "N" Or EntryMode = "E")
                    .Items("Print").Visible = (InquiryMode Or EntryMode = "V")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel Order").Visible = (EntryMode = "E")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Delete").Visible = (EntryMode = "E")

                    .Items("Approve").Visible = (Not InquiryMode And EntryMode = "V" And ScreenMode And optStatus.Value = "a") _
                        And (ASCMAIN1.USER_SECURITY_CODEs.Contains("OM") Or VEND_BUYER_PURCH_LIMIT <> 0)

                    .Items("email").Visible = ScreenMode And Not (EntryMode = "N")
                    .Items("Transmit").Visible = (Not InquiryMode And EntryMode = "V" And ScreenMode And optStatus.Value = "t")
                    .Items("Re-Transmit").Visible = (EntryMode = "V" And (ScreenMode AndAlso rowPOTORDR1("PO_XMIT_IND") & "" = "1"))
                    .Items("Status Updates").Visible = False ' (Not InquiryMode And EntryMode = "V" And ScreenMode And optStatus.Value = "O")
                End With
                .Groups("PO Status").Visible = Not ScreenMode
                .Groups("Find by Receipt").Visible = Not ScreenMode
            End With
        End If

        If ScreenMode Then
            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
                    {grdPOTORDR2, grdPOTORDR5}
                If InquiryMode Or (EntryMode = "V") Then
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                Else
                    grd.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                    grd.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
                    grd.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                End If
            Next

            Dim ever_received As Boolean = (dst.Tables("POTORDR2").Select("PO_QTY_REC <> 0").Length <> 0)

            With grdPOTORDR2.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"BM_ISSUE_DATE", "PO_DATE_COMPSDUE"} ' "BM_ISSUE_NO"
                    .Columns(COLUMN_NAME).Hidden = Not (Absx1.optFor("PO_TYPE").Value = "M") Or (EntryMode <> "E" And EntryMode <> "V")
                Next
                For Each COLUMN_NAME As String In New String() {"PO_QTY_OPN_CALC"}
                    .Columns(COLUMN_NAME).Hidden = (EntryMode = "N")
                Next
                For Each COLUMN_NAME As String In New String() {"BM_ISSUE_SEL"}
                    .Columns(COLUMN_NAME).Hidden = Not (Absx1.optFor("PO_TYPE").Value = "M")
                Next
                For Each COLUMN_NAME As String In New String() {"PO_QTY_REC", "PO_QTY_INV"}
                    .Columns(COLUMN_NAME).Hidden = (EntryMode = "N" Or (EntryMode = "E" And Not ever_received))
                Next
                For Each COLUMN_NAME As String In New String() {"PO_STATUS"}
                    .Columns(COLUMN_NAME).Hidden = (EntryMode = "N")
                Next
            End With

            lblVEND_WHSE_CODE.Visible = (Absx1.optFor("PO_TYPE").Value = "M")
            txtVEND_WHSE_CODE.Visible = (Absx1.optFor("PO_TYPE").Value = "M")
            txtVEND_WHSE_DESC.Visible = (Absx1.optFor("PO_TYPE").Value = "M")

            tabNotes.Tabs("Reason for Revision").Visible = (EntryMode = "E")

            ' Set_Read_Only(UltraTabPageControl4, InquiryMode)
            Set_Read_Only(grpShipTo, InquiryMode)
            Set_Read_Only(splNotes, InquiryMode)
        End If

        tabPOTORDR2.Tabs("Component Commitments").Visible = (Absx1.optFor("PO_TYPE").Value = "M")
        tabPOTORDR2.Tabs("Invoices").Visible = (EntryMode = "V")

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdPOTORDRX.Visible = Not ScreenMode
        tabMain.Visible = ScreenMode

        lblStatus.Visible = ScreenMode
        lblRevision.Visible = ScreenMode

        If ScreenMode Then
            Set_Read_Only(grpShipTo, InquiryMode Or (EntryMode = "V"))
            Set_Read_Only(tabNotes, InquiryMode Or (EntryMode = "V"))
            Set_Read_Only(splPOTORDR1, InquiryMode Or (EntryMode = "V"))

            btnRequiredDate.Visible = (EntryMode = "N" Or EntryMode = "E")
        Else
            Clear_Record()
            btnRequiredDate.Visible = False
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"POTORDR1", "POTORDR2", "ICTIRECX", "POTORDR9", "POTORDR5", "POTORDR6",
             "TATEVNT1", "POTORDRT", "POTORDXR", "POTORDRH", "ICTIREC1", "ICTIREC2",
             "APTINVH5_SUM", "APTINVH1", "APTINVH5", "ICTPINV1", "ICTPINVD", "ICTPINVR"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        APPR_NOTES = ""
        preapproval_applied = False
        Absx1.txtFor("VEND_CODE").Text = ""
        Absx1.txtFor("RECEIPT_NO").Text = ""
        chkEditPlans.Checked = False
        optPO_ORDER_TYPE.Value = "P"
        Load_POTORDRX()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            rowPOTORDR1 = dst.Tables("POTORDR1").NewRow
            PO_ORDER_NO = ASCMAIN1.Next_Control_No("POTORDR1.PO_ORDER_NO")
            rowPOTORDR1.Item("PO_ORDER_NO") = PO_ORDER_NO
            rowPOTORDR1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowPOTORDR1.Item("PO_TYPE") = HFs("PO_TYPE")
            rowPOTORDR1.Item("PO_DATE_ORDERED") = DATETIME_STAMP.Date
            Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", HFs("VEND_CODE"))
            For Each COLUMN_NAME As String In New String() _
            {"VEND_NAME", "VEND_ADDR1", "VEND_ADDR2", "VEND_ADDR3" _
             , "VEND_CITY", "VEND_STATE", "VEND_ZIP_CODE", "VEND_COUNTRY" _
             , "VEND_PHONE", "VEND_EXT", "VEND_FAX", "VEND_EMAIL", "VEND_CONTACT" _
             , "TERM_CODE", "VEND_WHSE_CODE", "VEND_BUYER_CODE"}
                rowPOTORDR1.Item(COLUMN_NAME) = rowAPTVEND1.Item(COLUMN_NAME)
            Next

            If rowPOTORDR1.Item("PO_TYPE") = "B" Then
                rowPOTORDR1.Item("VEND_WHSE_CODE") = ""
            End If

            rowPOTORDR1.Item("PO_CONTACT") = rowAPTVEND1.Item("VEND_PURCH_CONTACT")
            rowPOTORDR1.Item("PO_FOB_DESC") = rowAPTVEND1.Item("VEND_PURCH_FOB_DESC")
            rowPOTORDR1.Item("PO_SHIP_VIA") = rowAPTVEND1.Item("VEND_PURCH_SHIP_VIA")
            rowPOTORDR1.Item("PO_ORDR_NOTES_INTERNAL") = rowAPTVEND1.Item("VEND_PURCH_COMMENT")

            'Dim COST_CLASS_CODE As String = rowAPTVEND1.Item("COST_CLASS_CODE") & ""
            'rowPOTORDR1.Item("COST_CLASS_CODE") = COST_CLASS_CODE
            'rowPOTORDR1.Item("COST_LIST_CODE") = rowAPTVEND1.Item("COST_LIST_CODE")
            'If COST_CLASS_CODE <> "" Then
            '    Dim rowICTCCLS1 As DataRow = dst.Tables("ICTCCLS1").Rows.Find(COST_CLASS_CODE)
            '    rowPOTORDR1.Item("COST_BASE_PCT_OF_MSRP") = rowICTCCLS1.Item("COST_BASE_PCT_OF_MSRP")
            'End If

            'rowPOTORDR1.Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            'rowPOTORDR1.Item("CURR_EXCH_RATE") = 1
            dst.Tables("POTORDR1").Rows.Add(rowPOTORDR1)

            ' rowPOTORDR1.Item("WHSE_CODE") = DBNull.Value
            'rowPOTORDR1.Item("WHSE_DESC") = "See PO Details for Ship-To Address"
            If rowPOTORDR1.Item("PO_FOB_DESC") & "" = "" Then rowPOTORDR1.Item("PO_FOB_DESC") = ROWs("POTPARM1").Item("PO_PARM_FOB")
            If rowPOTORDR1.Item("PO_SHIP_VIA") & "" = "" Then rowPOTORDR1.Item("PO_SHIP_VIA") = ROWs("POTPARM1").Item("PO_PARM_SHIP_VIA")
            rowPOTORDR1.Item("PO_DATE_ORDERED") = DATETIME_STAMP.Date
            rowPOTORDR1.Item("PO_STATUS") = "O"
            rowPOTORDR1.Item("PO_ORDER_TYPE") = HFs("PO_ORDER_TYPE") ' "P"
            rowPOTORDR1.Item("MARKET_CODE") = "DPT"
            lblStatus.Text = "New PO"
            lblRevision.Text = ""
            rowPOTORDR1.Item("PO_DISASSEMBLY_IND") = HFs("PO_DISASSEMBLY_IND")

            rowPOTORDR1.Item("TRF_CLASS_CODE") = rowAPTVEND1.Item("TRF_CLASS_CODE")
            'rowPOTORDR1.Item("VEND_CODE_TRF") = rowAPTVEND1.Item("VEND_CODE_TRF")

            rowPOTORDR1.Item("FRT_CLASS_CODE") = rowAPTVEND1.Item("FRT_CLASS_CODE")
            'rowPOTORDR1.Item("VEND_CODE_FRT") = rowAPTVEND1.Item("VEND_CODE_FRT")

        Else
            rowPOTORDR1 = Fill_Record("POTORDR1", PO_ORDER_NO)
            Dim PO_STATUS As String = rowPOTORDR1.Item("PO_STATUS")
            Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
            lblStatus.Text = IIf(PO_STATUS = "O", "Open", "Closed")
            If PO_STATUS = "O" Then
                If rowPOTORDR1.Item("PO_XMIT_IND") & "" = "1" Then
                    lblStatus.Text &= ", Transmitted"
                ElseIf rowPOTORDR1.Item("PO_APPR_BY") & "" <> "" Then
                    lblStatus.Text &= ", Pending Transmit"
                ElseIf rowPOTORDR1.Item("PO_APPR_PENDING") & "" = "1" Then
                    lblStatus.Text &= ", Pending Approval"
                Else
                    lblStatus.Text &= ", Work in Process"
                End If
            End If
            lblRevision.Text = IIf(PO_HDR_CTR_REV = 0, "Original", "Rev #" & CStr(PO_HDR_CTR_REV))

            If EntryMode = "E" Then
                If rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1" Then
                    rowPOTORDR1.Item("PO_PRINTED_IND") = "0"
                    rowPOTORDR1.Item("PO_XMIT_IND") = "0"
                    PO_HDR_CTR_REV += 1
                    rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV
                    rowPOTORDR1.Item("PO_REVISION_NOTE") = ""
                End If
                lblRevision.Text = "Rev#" & CStr(PO_HDR_CTR_REV)
                Record_Event("LAST", "PO Edit Started")
            Else
                lblRevision.Text = "Rev#" & CStr(PO_HDR_CTR_REV)
            End If
        End If

        Fill_Records("POTORDRH", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDRH, "PO_HDR_CTR_REV")

        Fill_Records("POTORDR2", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDR2, "PO_ORDER_LNO")
        Fill_Records("ICTIRECX", PO_ORDER_NO)
        Sort_grdColumns(grdICTIRECX, "RECEIPT_NO")
        Fill_Records("POTORDR9", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDR9, "ITEM_CODE")

        Fill_Records("POTORDR5", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDR5, "PO_NINV_LNO")
        Fill_Records("POTORDR6", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDR6, "PO_NINV_SLNO")

        Fill_Records("TATEVNT1", PO_ORDER_NO)
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
        Fill_Records("POTORDXR", PO_ORDER_NO)
        Sort_grdColumns(grdPOTORDXR, "INIT_DATE".ToLower)

        Fill_Records("ICTIREC2", PO_ORDER_NO)
        Fill_Records("ICTIREC1", PO_ORDER_NO)
        Sort_grdColumns(grdICTIREC1, "RECEIPT_NO")
        Fill_Records("APTINVH5_SUM", PO_ORDER_NO)
        Fill_Records("APTINVH1", PO_ORDER_NO)
        Sort_grdColumns(grdAPTINVH1, "VOUCHER_NO")
        grdAPTINVH1.Rows.ExpandAll(True)

        Fill_Records("APTINVH5", PO_ORDER_NO)

        Fill_Records("ICTPINV1", PO_ORDER_NO)
        Sort_grdColumns(grdICTPINV1, "PO_ORDER_NO".ToLower & "," & "INV_DATE".ToLower)

        Fill_Records("DPTPLAN1", VEND_CODE)
        Sort_grdColumns(grdDPTPLAN1, "ITEM_CODE,DATE_REQUIRED")

        If EntryMode = "V" Then

            Fill_Records("ICTPINVR", PO_ORDER_NO)
            Fill_Records("ICTPINVD", PO_ORDER_NO)

            For Each row As DataRow In dst.Tables("POTORDR2").Select("")
                Dim INV_QTY As Integer = Val(dst.Tables("ICTPINVD").Compute("SUM(INV_QTY)", $"PO_ORDER_LNO = {row.Item("PO_ORDER_LNO")}") & "")
                Dim PO_QTY_OPN As Integer = Val(row.Item("PO_QTY_OPN") & "")
                If PO_QTY_OPN <> INV_QTY Then
                    Dim rowICTPINVD As DataRow = dst.Tables("ICTPINVD").NewRow
                    With rowICTPINVD
                        .Item("PO_ORDER_NO") = row.Item("PO_ORDER_NO")
                        .Item("PO_ORDER_LNO") = row.Item("PO_ORDER_LNO")
                        .Item("CONTAINER_NO") = "Qty Open"
                        .Item("INV_NUM") = "Not Inv"
                        .Item("ETA_DATE") = row.Item("PO_DATE_REQUIRED")
                        .Item("OPO_QTY") = PO_QTY_OPN - INV_QTY
                        .Item("WHSE_CODE") = row.Item("WHSE_CODE")
                        .Item("PINV_LNO") = 0
                    End With
                    dst.Tables("ICTPINVD").Rows.Add(rowICTPINVD)
                End If
            Next

            For Each row As DataRow In dst.Tables("ICTPINVD").Select("CONTAINER_NO IS NULL")
                With row
                    Dim WHSE_CODE As String = row.Item("WHSE_CODE") & ""
                    ' ISSUE-7230 Clarins to ADS
                    'If ASCMAIN1.CLIENT = "INT" Then
                    '    If WHSE_CODE <> "CLA" Or WHSE_CODE <> "ADS" Then
                    '        WHSE_CODE = "CLA"
                    '    End If
                    'End If
                    If WHSE_CODE.Length = 0 Then
                        WHSE_CODE = ROWs("POTPARM1").Item("PO_PARM_WHSE_CODE")
                    End If

                    Dim PORT_CODE As String = row.Item("PORT_CODE") & ""
                    If PORT_CODE = "" Then PORT_CODE = PO_PARM_PINV_PORT
                    Dim rowICTPORT2 As DataRow = dst.Tables("ICTPORT2").Rows.Find(New String() {PORT_CODE, WHSE_CODE})
                    Dim PINV_LT As Int32 = PO_PARM_PINV_LT
                    If rowICTPORT2 IsNot Nothing Then
                        PINV_LT = Val(rowICTPORT2.Item("ETD_TO_ETA") & "")
                    End If
                    Dim INV_DATE As Date = CDate(row.Item("INV_DATE"))
                    .Item("CONTAINER_NO") = $"Inv Date +{CStr(PINV_LT)}"
                    .Item("ETA_DATE") = INV_DATE.AddDays(PINV_LT)
                End With
            Next

            For Each ROW As DataRow In dst.Tables("ICTPINVD").Select("")
                If ROW.Item("ETA_DATE") & "" <> "" Then
                    Dim ETA_DATE As Date = ROW.Item("ETA_DATE")
                    Dim ETA_DATE_DC As Date = ETA_DATE
                    For I As Integer = 1 To 5
                        ETA_DATE_DC = ETA_DATE_DC.AddDays(1)
                        If ETA_DATE_DC.DayOfWeek = DayOfWeek.Saturday Or ETA_DATE_DC.DayOfWeek = DayOfWeek.Sunday Then
                            I = I - 1
                        End If
                    Next
                    ROW.Item("ETA_DATE_DC") = ETA_DATE_DC
                End If
            Next
        End If

        Setup_grdICTIRECX()
        Setup_grdPOTORDR9()
        Setup_grdPOTORDR6()
        Setup_grdICTPINVD()

        EnforceConstraints(True)

        If EntryMode = "N" Then
            Record_Event("INIT", "PO Entry Started")
        ElseIf EntryMode = "E" Then
            rowPOTORDR1.Item("PO_REVISION_NOTE") = DBNull.Value
            Record_Event("LAST", "PO Edit Started")
        End If
        Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)

        PO_DATE_REQUIRED_last = Absx1.dteFor("PO_DATE_REQUIRED").Value & ""
        WHSE_CODE_last = Absx1.txtFor("WHSE_CODE").Text

        txtVEND_NAME.Text = Absx1.txtFor("VEND_NAME").Text

        Display_Totals()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            BeginTrans()

            If ASCMAIN1.DBS_COMPANY = "INT" Then
                ' NOT YET
            Else
                For Each rowDPTPLAN1 As DataRow In dst.Tables("DPTPLAN1").Select("DELETE = '1'")
                    rowDPTPLAN1.Delete()
                Next
                Update_Record_TDA("DPTPLAN1")
            End If

            For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                rowPOTORDR2.Item("PO_QTY_OPN") = rowPOTORDR2.Item("PO_QTY_OPN_CALC")
                If Val(rowPOTORDR2.Item("PO_QTY_OPN") & "") = 0 And rowPOTORDR2.Item("PO_STATUS") & "" = "O" Then
                    rowPOTORDR2.Item("PO_STATUS") = "C"
                End If
            Next

            Dim PO_QTY_OPN As Int64 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", "") & "")
            Dim PO_QTY_REC As Int64 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_REC)", "") & "")
            Dim PO_NINV_QTY_OPN As Int64 = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_QTY_OPN)", "") & "")
            Dim PO_NINV_AMT_OPN As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMT_OPN)", "") & "")

            If PO_QTY_OPN = 0 And PO_NINV_QTY_OPN = 0 Then
                rowPOTORDR1.Item("PO_STATUS") = "C"
            Else
                rowPOTORDR1.Item("PO_STATUS") = "O"
            End If

            If rowPOTORDR1.Item("PO_DISASSEMBLY_IND") & "" = "1" Then
                rowPOTORDR1.Item("PO_XMIT_IND") = "1"
            End If

            If EntryMode <> "N" Then
                Dependent_Updates(-1, PO_ORDER_NO)

                'If rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1" Then
                '    rowPOTORDR1.Item("PO_HDR_CTR_REV") = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "") + 1
                '    rowPOTORDR1.Item("PO_PRINTED_IND") = DBNull.Value
                '    rowPOTORDR1.Item("PO_XMIT_IND") = "0"
                '    'rowPOTORDR1.Item("PO_APPR_NOTES") = DBNull.Value
                '    'rowPOTORDR1.Item("PO_APPR_BY") = DBNull.Value
                '    'rowPOTORDR1.Item("PO_APPR_DATE") = DBNull.Value
                '    'rowPOTORDR1.Item("PO_APPR_AMOUNT") = DBNull.Value
                'End If

                If preapproval_applied Then
                Else
                    rowPOTORDR1.Item("PO_APPR_NOTES") = DBNull.Value
                    rowPOTORDR1.Item("PO_APPR_BY") = DBNull.Value
                    rowPOTORDR1.Item("PO_APPR_DATE") = DBNull.Value
                    rowPOTORDR1.Item("PO_APPR_AMOUNT") = DBNull.Value
                End If

                If Check_Changed_Fields() Then
                    'Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
                    'rowPOTORDR1.Item("PO_HDR_CTR_REV") = PO_HDR_CTR_REV + 1
                End If
            End If

            INIT_LAST("POTORDR1", True, , True)
            Record_Event("UPDT", "PO Updated")

            Dim sql_delete As String = "PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            Update_Record_TDA("POTORDR1")

            If rowPOTORDR1.Item("VEND_WHSE_CODE") & "" = "" Then
                For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                    rowPOTORDR2.Item("BM_ISSUE_NO") = DBNull.Value
                    rowPOTORDR2.Item("BM_ISSUE_DATE") = DBNull.Value
                    rowPOTORDR2.Item("BM_ISSUE_SEL") = DBNull.Value
                Next
            Else

            End If

            Update_Record_TDA("POTORDR2", sql_delete)
            Update_Record_TDA("POTORDRH")
            Update_Record_TDA("POTORDR5", sql_delete)

            Dependent_Updates(1, PO_ORDER_NO)

            Update_Record_TDA("POTORDXR")
            Update_Record_TDA("TATEVNT1")

            ASCMAIN1.sql = "SELECT * FROM (" & vbCrLf _
                & "SELECT ITEM_CODE, WHSE_CODE, SUM (PO) PO, SUM (IC) IC FROM (" & vbCrLf _
                & "SELECT ITEM_CODE, WHSE_CODE, 0 PO, WHSE_QTY_ONPO IC FROM ICTSTAT2" & vbCrLf _
                & "WHERE WHSE_QTY_ONPO <> 0" & vbCrLf _
                & "UNION" & vbCrLf _
                & "SELECT ITEM_CODE, WHSE_CODE, SUM (PO_QTY_OPN) PO, 0 IC" & vbCrLf _
                & "FROM POTORDR2 WHERE PO_QTY_OPN <> 0 GROUP BY ITEM_CODE, WHSE_CODE" & vbCrLf _
                & ")  GROUP BY ITEM_CODE, WHSE_CODE" & vbCrLf _
                & ") WHERE PO <> IC"
            Dim tbl As DataTable = ASCDATA1.GetDataTable

            If tbl.Rows.Count <> 0 Then
                Using frm As New ASFMSGBF
                    frm.Show_grd(tbl, Me, "Inventory Status")
                End Using
            End If

            CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Sub Cancel_Order()
        If EntryMode = "N" Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Cancelling PO")
        BeginTrans()
        Cancel_Order_1(PO_ORDER_NO)
        CommitTrans("PO " & PO_ORDER_NO & " has been Cancelled")
    End Sub

    Sub Cancel_Order_1(PO_ORDER_NO As String)

        Dependent_Updates(-1, PO_ORDER_NO)

        Dim sqlw As String = " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        Dim sqlw2 As String = " and PO_QTY_OPN <> 0"
        ASCDATA1.ExecuteSQL("Update POTORDR2 set PO_QTY_CXL = NVL(PO_QTY_CXL,0) + NVL(PO_QTY_OPN,0) " & sqlw & sqlw2)
        ASCDATA1.ExecuteSQL("Update POTORDR2 set PO_QTY_OPN = 0, PO_STATUS = 'C' " & sqlw & sqlw2)
        ASCDATA1.ExecuteSQL("Update POTORDR5 set PO_NINV_QTY_OPN = 0, PO_NINV_AMT_OPN = 0, PO_NINV_STATUS = 'C' " & sqlw)

        If rowPOTORDR1.Item("PO_PRINTED_IND") & "" = "1" Then
            rowPOTORDR1.Item("PO_HDR_CTR_REV") = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "") + 1
            rowPOTORDR1.Item("PO_PRINTED_IND") = DBNull.Value
        End If
        rowPOTORDR1.Item("PO_STATUS") = "C"
        rowPOTORDR1.Item("PO_DATE_CANCELLED") = DATETIME_STAMP.Date
        Update_Record_TDA("POTORDR1")

        TAC.TACMAIN1.Record_Event("POTORDR1", PO_ORDER_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "PO-CXL", "PO Cancelled")
    End Sub

    Sub Delete_Record()
        If EntryMode = "N" Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Deleting PO")
        BeginTrans()
        Dependent_Updates(-1, PO_ORDER_NO)
        Delete_Records("POTORDR1")
        Delete_Records("POTORDR2")
        Delete_Records("POTORDR5")

        Record_Event("DELE", "PO Deleted")
        Update_Record_TDA("TATEVNT1")

        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'")
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
                Absx1.txtFor("PO_ORDER_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "POTORDR1"
            E.COLUMN_NAME = "PO_ORDER_NO"
            E.CODE_VALUE = HFs("PO_ORDER_NO")
            E.DESC_VALUE = HFs("VEND_CODE")
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "POTORDR1"
        E.TABLE_KEY_CAPTION = "Purchase Order"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("PO_ORDER_NO").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("VEND_CODE").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E" Or EntryMode = "A")
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
        Load_Popup_Menu(grdDPTPLAN1, "S", "Show Filter")
        Load_Popup_Menu(grdPOTORDRX, "SSSBB", "Show Filter", "Show GroupBox", "Show Pins", "Re-Send 943", "Mark as Transmitted")
        Load_Popup_Menu(grdPOTORDRH, "B", "Show PO")
        Load_Popup_Menu(grdTATEVNT1, "B", "Show email")
        Load_Popup_Menu(grdPOTORDR2, "B", "Item Status Inquiry", "Split Order Line")
        Load_Popup_Menu(grdPOTORDR9, "B", "Item Status Inquiry")
        Load_Popup_Menu(grdICTIRECX, "B", "PO Receipts Inquiry", "Voucher Inquiry")
        Load_Popup_Menu(grdICTIREC1, "BB", "PO Receipts Inquiry")
        Load_Popup_Menu(grdAPTINVH1, "B", "Voucher Inquiry")
        Load_Popup_Menu(grdICTPINV1, "SS", "Show Filter", "Show GroupBox", "Pre-Receiving Advice Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

                Case "grdTATEVNT1"
                    tlb_btn = DirectCast(tlb_pop.Tools("Show email"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (grd.ActiveRow.Cells("EVENT_TYPE").Value = "POXMIT"))

                Case "grdPOTORDRX"
                    tlb_btn = DirectCast(tlb_pop.Tools("Re-Send 943"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (Not InquiryMode And (optStatus.Value = "t" Or optStatus.Value = "T" Or optStatus.Value = "O")))
                    If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                        tlb_btn.SharedProps.Visible = False
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Mark As Transmitted"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow IsNot Nothing AndAlso (Not InquiryMode And (optStatus.Value = "t")))
                    If ASCMAIN1.CLIENT = "AHA" Then
                        tlb_btn.SharedProps.Visible = False
                    End If

                Case "grdPOTORDR2"
                    tlb_btn = DirectCast(tlb_pop.Tools("Split Order Line"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (Not InquiryMode And EntryMode <> "V")
                    tlb_btn.SharedProps.Visible = False ' DISABLING THIS UNTIL WE GET MORE DESIGN SPECS

                Case "grdICTIRECX"
                    tlb_btn = DirectCast(tlb_pop.Tools("PO Receipts Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Index = 0)
                    tlb_btn = DirectCast(tlb_pop.Tools("Voucher Inquiry"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Index = 1)

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Pre-Receiving Advice Inquiry"
                Dim PINV_NO As String = grd.ActiveRow.Cells("PINV_NO").Text
                Context_Launch("View", PINV_NO, e.Tool.Key, "ICFPINVI", "F", "POE")

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "POE")

            Case "PO Receipts Inquiry"
                Dim RECEIPT_NO As String = grd.ActiveRow.Cells("RECEIPT_NO").Text
                Context_Launch("View", RECEIPT_NO, e.Tool.Key, "ICFIRECI")

            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Text
                Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Show email"
                If grd.ActiveRow.Cells("EVENT_TYPE").Value & "" = "POXMIT" Then
                    Dim FILENAME As String = grd.ActiveRow.Cells("EVENT_KEY").Value & ".EML"
                    Show_Document(ASCMAIN1.Folders("Archive") & "\email\Sent\" & FILENAME)
                End If

            Case "Show PO"
                Dim FILENAME As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value & "_" & CStr(Val(grd.ActiveRow.Cells("PO_HDR_CTR_REV").Value & "")) & ".PDF"
                Show_Document(ASCMAIN1.Folders("Archive") & "PO\" & FILENAME)

            Case "Re-Send 943" ' NOTE - THIS IS CODED FOR POS HEADED TO CLA ONLY
                If grd.Selected.Rows.Count = 0 AndAlso grd.ActiveRow IsNot Nothing Then
                    grd.ActiveRow.Selected = True
                End If

                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("No POs Selected", MsgBoxStyle.OkOnly, "Cannot Re-Send 943")
                Else
                    If MsgBox("OK Re-Send 943(s) To the Warehouse For the " & CStr(grd.Selected.Rows.Count) & " PO(s) Selected?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Me.Cursor = Cursors.WaitCursor
                        ASCMAIN1.Progress("Now Sending 943 Data")

                        Dim PO_ORDER_NOs As New List(Of String)
                        For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                            PO_ORDER_NO = grow.Cells("PO_ORDER_NO").Value
                            Dim row As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", row.Item("WHSE_CODE") & "")

                            If rowICTWHSE1 IsNot Nothing AndAlso rowICTWHSE1.Item("LP_CODE") & String.Empty <> String.Empty Then
                                If rowICTWHSE1.Item("LP_CODE") & String.Empty = "ADS" Then
                                    If row.Item("PO_TYPE") = "M" Then
                                        MessageBox.Show($"PO {PO_ORDER_NO} is a Make PO and will be skipped. You must Transmit Make POs.", "Re-Send 943", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        Continue For
                                    End If
                                End If
                                ASCMAIN1.sql = "Select Count (*) from POTORDR2 where PO_ORDER_NO = : PARM1 and PO_QTY_OPN <> 0"
                                Dim PO_lines As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {PO_ORDER_NO})
                                If PO_lines = 0 Then
                                    MsgBox("There are no Open Inventory Lines on PO " & PO_ORDER_NO, MsgBoxStyle.OkOnly, "PO should NOT be transmitted to Clarins")
                                Else
                                    PO_ORDER_NOs.Add(PO_ORDER_NO)
                                End If
                            Else
                                MsgBox("PO " & PO_ORDER_NO & " is not assigned to a 3PL - cannot send", MsgBoxStyle.OkOnly, "Transmission Request Cancelled")
                                PO_ORDER_NOs.Clear()
                                Exit For
                            End If
                        Next

                        If PO_ORDER_NOs.Count > 0 Then
                            Transmit_PO_to_3PL(PO_ORDER_NOs, True, True)
                            MsgBox(CStr(PO_ORDER_NOs.Count) & " 943(s) for Selected PO(s) have been queued up for sftp", MsgBoxStyle.OkOnly, "Success")
                        End If

                        Me.Cursor = Cursors.Default
                        ASCMAIN1.Progress("")
                    End If
                End If

            Case "Mark as Transmitted"

                If grdPOTORDRX.Selected.Rows.Count = 0 Then
                    If grdPOTORDRX.ActiveRow IsNot Nothing AndAlso grdPOTORDRX.ActiveRow.IsDataRow AndAlso Not grdPOTORDRX.ActiveRow.IsFilterRow Then
                        grdPOTORDRX.ActiveRow.Selected = True
                    End If
                End If

                If grdPOTORDRX.Selected.Rows.Count = 0 Then
                    MsgBox("No POs Selected", MsgBoxStyle.OkOnly, "Cannot Perform Mark as Transmitted")
                    Exit Sub
                End If

                Dim POs As New List(Of String)
                For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDRX.Selected.Rows
                    If grow.IsDataRow And Not grow.IsFilterRow Then
                        Dim PO_ORDER_NO As String = grow.Cells("PO_ORDER_NO").Value

                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, False, True, True, 1) Then
                            Exit Sub
                        End If
                        POs.Add(PO_ORDER_NO)
                    End If
                Next

                If POs.Count = 0 Then
                    MsgBox("No POs Selected", MsgBoxStyle.OkOnly, "Cannot Perform Mark as Transmitted")
                    Exit Sub
                End If

                If MsgBox($"OK to Mark {CStr(POs.Count)} POs as Transmitted?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "Verification") = MsgBoxResult.Yes Then
                Else
                    Exit Sub
                End If

                ' BeginTrans()

                EMsg = ""
                For Each PO_ORDER_NO_toTransmit As String In POs
                    Me.PO_ORDER_NO = PO_ORDER_NO_toTransmit
                    Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
                    If rowPOTORDR1.Item("PO_STATUS") & "" <> "O" _
                        Or rowPOTORDR1.Item("PO_APPR_BY") & "" = "" _
                              Or rowPOTORDR1.Item("PO_XMIT_IND") & "" = "1" _
                        Or rowPOTORDR1.Item("PO_APPR_PENDING") & "" = "0" Then
                        EMsg &= vbCr & "PO " & PO_ORDER_NO & " is No Longer Pending Transmit"
                    Else
                        Transmit_PO(, True)
                    End If
                Next

                ' CommitTrans(CStr(POs.Count) & " POs have been marked as Transmitted" & EMsg)
                MsgBox(CStr(POs.Count) & " POs have been marked as Transmitted" & EMsg, vbOKOnly, "Verification")

                ASCMAIN1.MultiTask_Release(,, 1)

            Case "Split Order Line"

                Dim PO_QTY_BACKORDER As Integer = Val(grd.ActiveRow.Cells("PO_QTY_BACKORDER").Value & "")
                Dim PO_DATE_BACKORDER As String = grd.ActiveRow.Cells("PO_DATE_BACKORDER").Text
                If PO_QTY_BACKORDER = 0 Or PO_DATE_BACKORDER = "" Then
                    MsgBox("There must be a BackOrder Qty and a BackOrder Date to Split a PO Detail Line", MsgBoxStyle.OkOnly, "Cannot Split PO Detail Line")
                    Exit Sub
                End If


                If PO_QTY_BACKORDER >= Val(grd.ActiveRow.Cells("PO_QTY_OPN_CALC").Value & "") Then
                    MsgBox("Backorder Qty (" & CStr(Val(grd.ActiveRow.Cells("PO_QTY_BACKORDER").Value & "")) & ") Must be less than the Qty Open (" & CStr(Val(grd.ActiveRow.Cells("PO_QTY_OPN_CALC").Value & "")) & ")", 0, "Cannot Split Line")
                    Exit Sub
                End If


                Dim PO_ORDR_LNO As Integer = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", "") & "") + 1
                Dim rowPOTORDR2 As DataRow = Nothing
                rowPOTORDR2 = dst.Tables("POTORDR2").NewRow
                With rowPOTORDR2
                    .Item("PO_ORDER_NO") = grd.ActiveRow.Cells("PO_ORDER_NO").Value & ""
                    .Item("PO_ORDER_LNO") = PO_ORDR_LNO
                    .Item("ITEM_CODE") = grd.ActiveRow.Cells("ITEM_CODE").Value & ""
                    .Item("ITEM_DESC") = grd.ActiveRow.Cells("ITEM_DESC").Value & ""
                    .Item("ITEM_UOM") = grd.ActiveRow.Cells("ITEM_UOM").Value & ""
                    .Item("ITEM_EAN_CODE") = grd.ActiveRow.Cells("ITEM_EAN_CODE").Value & ""

                    .Item("ITEM_PCT_ALLOW_OVER") = Val(grd.ActiveRow.Cells("ITEM_PCT_ALLOW_OVER").Value & "")
                    .Item("ITEM_PCT_ALLOW_UNDER") = Val(grd.ActiveRow.Cells("ITEM_PCT_ALLOW_UNDER").Value & "")
                    .Item("PO_COST") = grd.ActiveRow.Cells("PO_COST").Value & ""
                    .Item("PO_QTY_ORD") = PO_QTY_BACKORDER
                    .Item("WHSE_CODE") = grd.ActiveRow.Cells("WHSE_CODE").Value & ""
                    .Item("PO_DATE_REQUIRED") = PO_DATE_BACKORDER
                    .Item("PO_STATUS") = grd.ActiveRow.Cells("PO_STATUS").Value & ""
                    .Item("PO_QTY_BACKORDER") = 0
                    .Item("PO_DATE_BACKORDER") = System.DBNull.Value

                    ' DO ALL FIELDS  ?

                    Dim BM_ISSUE_DATE As String = grd.ActiveRow.Cells("BM_ISSUE_DATE").Text
                    Dim PO_DATE_REQUIRED_MRP As String = grd.ActiveRow.Cells("PO_DATE_REQUIRED_MRP").Text
                    Dim PO_DATE_COMPSDUE As String = grd.ActiveRow.Cells("PO_DATE_COMPSDUE").Text

                    .Item("PO_QTY_OPN") = PO_QTY_BACKORDER
                    .Item("PO_PRICE_VAR_REASON") = grd.ActiveRow.Cells("PO_PRICE_VAR_REASON").Value & ""
                    .Item("BM_ISSUE_NO") = grd.ActiveRow.Cells("BM_ISSUE_NO").Value & ""
                    If BM_ISSUE_DATE <> "" Then
                        .Item("BM_ISSUE_DATE") = BM_ISSUE_DATE
                    End If
                    .Item("BM_ISSUE_SEL") = grd.ActiveRow.Cells("BM_ISSUE_SEL").Value & ""
                    .Item("PO_ITEM_NOTE") = grd.ActiveRow.Cells("PO_ITEM_NOTE").Value & ""
                    .Item("PO_AUTO_PRD_SUB") = grd.ActiveRow.Cells("PO_AUTO_PRD_SUB").Value & ""
                    If PO_DATE_REQUIRED_MRP <> "" Then
                        .Item("PO_DATE_REQUIRED_MRP") = PO_DATE_REQUIRED_MRP
                    End If
                    If PO_DATE_COMPSDUE <> "" Then
                        .Item("PO_DATE_COMPSDUE") = PO_DATE_COMPSDUE
                    End If
                    .Item("ITEM_CODE_ALT") = grd.ActiveRow.Cells("ITEM_CODE_ALT").Value & ""

                End With
                dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2)

                grd.ActiveRow.Cells("PO_QTY_ORD").Value = grd.ActiveRow.Cells("PO_QTY_ORD").Value - PO_QTY_BACKORDER
                grd.ActiveRow.Cells("PO_QTY_BACKORDER").Value = 0
                grd.ActiveRow.Cells("PO_DATE_BACKORDER").Value = Null
                grd.ActiveRow.Update()

                '  rowPOTORDR1.Item("PO_APPR_PENDING") & "" = "0"
                '   grd.ActiveRow.Cells("PO_QTY_BACKORDER").Text = 15


                'For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDR2.Selected.Rows
                '    If grow.IsDataRow And Not grow.IsFilterRow Then
                '        Dim PO_QTY_BACKORDER As Integer = grow.Cells("PO_QTY_BACKORDER").Value
                '        Dim BACKORDER_DATE As Date = grow.Cells("BACKORDER_DATE").Value & ""
                '        If Val(PO_QTY_BACKORDER) = 0 Or BACKORDER_DATE = "" Then
                '            MsgBox("There must be a BO Qty and a BO Date to split a PO Detail LineE", MsgBoxStyle.OkOnly, "Cannot Split PO Detail Line")
                '            Exit Sub


                '        End If
                '    End If
                'Next

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEND_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Click_Command("Load", e)
                    Load_POTORDRX()
                End If

            Case "PO_ORDER_NO"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If InquiryMode Then
                        Click_Command("View", e)
                    Else
                        Click_Command("View") '  Click_Command("Edit", e)
                    End If
                End If

            Case "RECEIPT_NO"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim row As DataRow = LookUp("ICTIREC1", Absx1.txtFor("RECEIPT_NO").Text)
                    If row IsNot Nothing Then
                        Absx1.txtFor("PO_ORDER_NO").Text = row.Item("PO_ORDER_NO") & ""
                        Click_Command("View")
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "VEND_CODE"
                If InquiryMode Then
                    'Click_Command("Load")
                Else
                    'Click_Command("New")
                End If
                Load_POTORDRX()

            Case "PO_ORDER_NO"
                If InquiryMode Then
                    Click_Command("View")
                Else
                    Click_Command("View") '  Click_Command("Edit")
                End If

            Case "RECEIPT_NO"
                Dim row As DataRow = LookUp("ICTIREC1", Absx1.txtFor("RECEIPT_NO").Text)
                If row IsNot Nothing Then
                    Absx1.txtFor("PO_ORDER_NO").Text = row.Item("PO_ORDER_NO") & ""
                    Click_Command("View")
                End If

        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "VEND_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("VEND_CODE").Text <> "" Then
                        LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                        If cdr IsNot Nothing Then
                            If cdr.Item("VEND_WHSE_CODE") & "" <> "" Then
                                Absx1.optFor("PO_TYPE").Value = "M"
                            Else
                                Absx1.optFor("PO_TYPE").Value = "B"
                            End If
                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "grdPOTORDR2"

    Private Sub grdPOTORDR2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR2.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                Validate_Item(e.Cell.Value & "")
                If ITEM_CODE <> "" Then
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", New String() {ITEM_CODE}) ' {e.Cell.Value})
                    e.Cell.Row.Cells("ITEM_UOM").Value = rowICTITEM1.Item("ITEM_UOM") & ""
                    e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                    e.Cell.Row.Cells("ITEM_PCT_ALLOW_OVER").Value = Val(rowICTITEM1.Item("ITEM_PCT_ALLOW_OVER") & "")
                    e.Cell.Row.Cells("ITEM_CODE_ALT").Value = rowICTITEM1.Item("ITEM_CODE_ALT") & ""
                    e.Cell.Row.Cells("ITEM_UOM").Value = rowICTITEM1.Item("ITEM_UOM") & ""
                    e.Cell.Row.Cells("ITEM_EAN_CODE").Value = rowICTITEM1.Item("ITEM_EAN_CODE") & ""
                    e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = rowICTITEM1.Item("ITEM_RETAIL_PRICE")

                    If Absx1.optFor("PO_TYPE").Value = "M" Then
                        Setup_BM()
                        If dst.Tables("BMTMAIN2").Select("BM_ISSUE_NO <> '00'").Length > 0 Then
                            e.Cell.Row.Cells("BM_ISSUE_SEL").Value = "1"
                        End If
                    End If

                    If e.Cell.Row.IsAddRow Then


                        If WHSE_CODE_last <> "" Then
                            e.Cell.Row.Cells("WHSE_CODE").Value = WHSE_CODE_last
                        ElseIf Absx1.txtFor("WHSE_CODE").Text <> "" Then
                            e.Cell.Row.Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
                        End If

                        If PO_DATE_REQUESTED_last <> "" Then
                            e.Cell.Row.Cells("PO_DATE_REQUESTED").Value = CDate(PO_DATE_REQUESTED_last)
                        ElseIf Absx1.dteFor("PO_DATE_REQUIRED").Value & "" <> "" Then
                            e.Cell.Row.Cells("PO_DATE_REQUESTED").Value = Absx1.dteFor("PO_DATE_REQUIRED").Value
                        End If

                        If PO_DATE_REQUIRED_last <> "" Then
                            e.Cell.Row.Cells("PO_DATE_REQUIRED").Value = CDate(PO_DATE_REQUIRED_last)
                        ElseIf Absx1.dteFor("PO_DATE_REQUIRED").Value & "" <> "" Then
                            e.Cell.Row.Cells("PO_DATE_REQUIRED").Value = Absx1.dteFor("PO_DATE_REQUIRED").Value
                        End If

                        If optPO_ORDER_TYPE.Value = "R" Then
                            For Each rowBMTMAIN2 As DataRow In dst.Tables("BMTMAIN2").Select("", "BM_ISSUE_NO DESC")
                                e.Cell.Row.Cells("PO_COST").Value = rowBMTMAIN2.Item("BM_ISSUE_VCOST")
                                Exit For
                            Next
                        Else
                            e.Cell.Row.Cells("PO_COST").Value = Get_Price(Format(e.Cell.Row.Cells("PO_DATE_REQUIRED").Value, "yyyyMM"))
                        End If

                        '' Will comment these out
                        'Dim PO_COST_CLASS As Decimal = -1
                        'Dim PO_COST_LIST As Decimal = -1

                        'Dim PO_DATE_ORDERED As Date = rowPOTORDR1.Item("PO_DATE_ORDERED")
                        'Dim OPS_YYYYPP As String = Format(PO_DATE_ORDERED, "yyyyMM")

                        'Dim ITEM_RETAIL_PRICE As Decimal = Val(e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value & "")
                        '' NEXT LINE ONLY IF WE WANT Get_PO_Cost to get prior period ITEM_RETAIL_PRICE
                        'If OPS_YYYYPP <> ASCMAIN1.CYP Then
                        '    ITEM_RETAIL_PRICE = -1 * Val(e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value & "")
                        'End If
                        ''Dim ITEM_RETAIL_PRICE As Decimal = -1 
                        ''Dim ITEM_RETAIL_PRICE As Decimal = Val(e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value & "")


                        'Get_PO_Cost(rowPOTORDR1, ITEM_CODE, ITEM_RETAIL_PRICE, PO_COST_CLASS, PO_COST_LIST)
                        '' May need to add parameter to above - we are in grdPOTORDR2 - has the requested and confirmed dates (use confirmed)
                        'If PO_COST_CLASS <> -1 Then e.Cell.Row.Cells("PO_COST_CLASS").Value = PO_COST_CLASS
                        'If PO_COST_LIST <> -1 Then e.Cell.Row.Cells("PO_COST_LIST").Value = PO_COST_LIST
                        '' NEXT LINE ONLY IF WE WANT Get_PO_Cost to get prior period ITEM_RETAIL_PRICE
                        'If ITEM_RETAIL_PRICE > 0 Then e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = ITEM_RETAIL_PRICE
                        ''If ITEM_RETAIL_PRICE <> -1 Then e.Cell.Row.Cells("ITEM_RETAIL_PRICE").Value = ITEM_RETAIL_PRICE

                        'If PO_COST_LIST <> -1 Then
                        '    e.Cell.Row.Cells("PO_COST").Value = PO_COST_LIST
                        'ElseIf PO_COST_CLASS <> -1 Then
                        '    e.Cell.Row.Cells("PO_COST").Value = PO_COST_CLASS
                        'End If

                    End If
                End If

            Case "PO_DATE_REQUESTED"
                If e.Cell.Row.IsAddRow Then
                    e.Cell.Row.Cells("PO_DATE_REQUIRED").Value = e.Cell.Value
                End If

        End Select
    End Sub

    Private Sub grdPOTORDR2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTORDR2.AfterExitEditMode
        'With grdPOTORDR2

        '    Select Case .ActiveCell.Column.Key
        '        Case "ITEM_CODE"
        '            If .ActiveCell.Text <> "" Then
        '                'cdr = LookUp("ICTITEM1", .ActiveCell.Text)
        '                '.ActiveRow.Cells("ITEM_DESC").Value = cdr.Item("ITEM_DESC") & ""
        '                '.ActiveRow.Cells("ITEM_UOM").Value = cdr.Item("ITEM_UOM") & ""
        '                '.ActiveRow.Cells("PO_COST").Value = Val(cdr.Item("ITEM_COST_STD") & "")
        '            End If

        '        Case "WHSE_CODE"
        '            'If .ActiveCell.Text <> "" Then
        '            '    '.ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
        '            '    cdr = LookUp("ICTWHSE1", .ActiveCell.Text)
        '            'End If
        '    End Select
        ' End With
    End Sub

    Private Sub grdPOTORDR2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTORDR2.AfterRowActivate
        Setup_grdICTIRECX()
        Setup_grdPOTORDR9()
        Setup_grdICTPINVD()

        If (EntryMode = "E" Or EntryMode = "N") And ASCMAIN1.DBS_COMPANY <> "INT" Then
            splPOTORDR2details.Panel1Collapsed = True
            splPOTORDR2details.Panel2Collapsed = False
            splDetails.Panel2Collapsed = False
        Else
            splPOTORDR2details.Panel1Collapsed = False
            splPOTORDR2details.Panel2Collapsed = True
        End If

        With grdPOTORDR2.ActiveRow
            With grdPOTORDR2.DisplayLayout.Bands(0)
                Dim ITEM_CODE As String = grdPOTORDR2.ActiveRow.Cells("ITEM_CODE").Value & ""
                .Columns("BM_ISSUE_NO").ValueList = Get_BMs(ITEM_CODE)
                If grdPOTORDR2.ActiveRow.IsAddRow Then
                    .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit

                    For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_DATE_REQUIRED", "PO_DATE_REQUESTED", "WHSE_CODE", "ITEM_PCT_ALLOW_OVER", "ITEM_PCT_ALLOW_UNDER", "PO_COST", "PO_ITEM_NOTE", "BM_ISSUE_NO", "BM_ISSUE_SEL"}
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                    Next
                Else
                    .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit

                    For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_DATE_REQUIRED", "PO_DATE_REQUESTED", "WHSE_CODE", "ITEM_PCT_ALLOW_OVER", "ITEM_PCT_ALLOW_UNDER", "PO_COST", "PO_ITEM_NOTE", "BM_ISSUE_NO", "BM_ISSUE_SEL"}
                        If grdPOTORDR2.ActiveRow.Cells("PO_STATUS").Value = "C" Then
                            .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                        Else
                            .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                        End If
                    Next
                End If

            End With

            If .Cells("ITEM_CODE").Value & "" = "" Then
                grdPOTORDR2.ActiveCell = .Cells("ITEM_CODE")
                Exit Sub
            End If

            'If .IsAddRow Then
            '    If Absx1.txtFor("WHSE_CODE").Text <> "" Then .Cells("WHSE_CODE").Value = Absx1.txtFor("WHSE_CODE").Text
            '    If Absx1.dteFor("PO_DATE_REQUIRED").Value & "" <> "" Then .Cells("PO_DATE_REQUIRED").Value = Absx1.dteFor("PO_DATE_REQUIRED").Value
            'End If
        End With
    End Sub

    Private Sub grdPOTORDR2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTORDR2.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdPOTORDR2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDR2.AfterRowUpdate
        Display_Totals()
        WHSE_CODE_last = e.Row.Cells("WHSE_CODE").Text
        PO_DATE_REQUIRED_last = e.Row.Cells("PO_DATE_REQUIRED").Value & ""
        PO_DATE_REQUESTED_last = e.Row.Cells("PO_DATE_REQUESTED").Value & ""
    End Sub

    Private Sub grdPOTORDR2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTORDR2.BeforeExitEditMode

        If e.CancellingEditOperation Then Exit Sub
        With grdPOTORDR2.ActiveCell
            Select Case .Column.Key

                Case "ITEM_CODE"

                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            Dim ITEM_CODE As String = .Text
                            grdPOTORDR2.ActiveRow.Cells("BM_ISSUE_NO").Column.ValueList = Get_BMs(ITEM_CODE)
                        End If
                    End If

                Case "WHSE_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTWHSE1", .Value)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If

                Case "PO_STATUS"

                    Dim NEW_STATUS As String = grdPOTORDR2.ActiveCell.EditorResolved.Value & ""

                    For Each COLUMN_NAME As String In New String() {"PO_QTY_ORD", "PO_DATE_REQUIRED", "PO_DATE_REQUESTED", "WHSE_CODE", "ITEM_PCT_ALLOW_OVER", "ITEM_PCT_ALLOW_UNDER", "PO_COST", "PO_ITEM_NOTE"}
                        If NEW_STATUS = "C" Then
                            grdPOTORDR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                        Else
                            grdPOTORDR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                        End If
                    Next

                Case Else
                    ' e.Cancel = Validate_Columns_2(.Column.Key)
            End Select
        End With

    End Sub

    Private Sub grdPOTORDR2_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTORDR2.BeforeRowsDeleted

        For Each grow As UltraWinGrid.UltraGridRow In grdPOTORDR2.Selected.Rows
            If grow.IsAddRow Then
            Else
                ASCMAIN1.sql = "Select Count (*) from ICTIREC2" _
    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(grow.Cells("PO_ORDER_LNO").Value)
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    MsgBox("PO Line " & grow.Cells("PO_ORDER_LNO").Value & " has had Receipts and/or Invoices Posted Against it",
                           MsgBoxStyle.OkOnly, "Cannot Delete this Record")
                    e.Cancel = True
                    Exit For
                End If

                '09/13/2023 coded around Closed (ie reversed) ICTPINV1 records so that we can delete the PO Line
                ASCMAIN1.sql = "Select Count (*) from ICTPINV2, ICTPINV1" _
    & $" where ICTPINV1.PINV_NO = ICTPINV2.PINV_NO and ICTPINV1.PINV_STATUS <> 'C' and ICTPINV2.PO_ORDER_NO = '{PO_ORDER_NO}' and ICTPINV2.PO_ORDER_LNO = {CStr(grow.Cells("PO_ORDER_LNO").Value)}"
                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    MsgBox("PO Line " & grow.Cells("PO_ORDER_LNO").Value & " has had a PO Receiving Advice Posted Against it",
                           MsgBoxStyle.OkOnly, "Cannot Delete this Record")
                    e.Cancel = True
                    Exit For
                End If


            End If
        Next
    End Sub

    Private Sub grdPOTORDR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTORDR2.BeforeRowUpdate

        e.Cancel = Validate_Columns_2("ITEM_CODE")
        If Not e.Cancel Then e.Cancel = Validate_Columns_2("PO_QTY_ORD")
        If Not e.Cancel Then e.Cancel = Validate_Columns_2("PO_DATE_REQUIRED")
        If Not e.Cancel Then e.Cancel = Validate_Columns_2("PO_DATE_REQUESTED")
        If Not e.Cancel Then e.Cancel = Validate_Columns_2("WHSE_CODE")
        If Not e.Cancel Then e.Cancel = Validate_Columns_2("PO_QTY_BACKORDER")
        If Not e.Cancel Then e.Cancel = Validate_Columns_2("PO_DATE_BACKORDER")



        If e.Row.Cells("BM_ISSUE_SEL").Value & "" = "1" And e.Row.Cells("BM_ISSUE_NO").Value & "" = "" Then
            If e.Row.IsAddRow Then
                Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value & ""
                Dim BM_ISSUE_TYPE As String = "A"
                If optPO_ORDER_TYPE.Value = "R" Then BM_ISSUE_TYPE = "R"
                ASCMAIN1.sql = $"Select MAX(BM_ISSUE_NO) from BMTMAIN2 where BM_PROD_ITEM = '{ITEM_CODE}' and BM_ISSUE_TYPE = '{BM_ISSUE_TYPE}'"
                Dim BM_ISSUE_NO As String = ASCDATA1.GetDataValue
                If BM_ISSUE_NO <> "" Then
                    e.Row.Cells("BM_ISSUE_NO").Value = BM_ISSUE_NO
                Else
                    e.Cancel = True
                End If
            Else
                e.Cancel = True
            End If
        End If

        'If Not e.Cancel Then
        '    e.Row.Cells("PO_QTY_OPN").Value = e.Row.Cells("PO_QTY_OPN_CALC").Value
        'End If

        ' ENABLE THE BELOW IF YOU WANT TO INTERACTIVELY WARN ABOUT ITEM DUPLICATION - PROCEED PREREQ HAS A SIMILAR SECTION
        'If e.Row.IsAddRow And Not e.Cancel Then
        '    Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value & ""
        '    If dst.Tables("POTORDR2").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length <> 0 Then
        '        e.Cancel = True
        '        MsgBox("Item " & ITEM_CODE & " is already on PO", MsgBoxStyle.OkOnly, "Duplicate Items are NOT Allowed on a Single PO")
        '    End If
        'End If

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("PO_ORDER_NO").Value = Absx1.txtFor("PO_ORDER_NO").Text
            Dim PO_ORDER_LNO As Int32 = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", "") & "") + 1
            e.Row.Cells("PO_ORDER_LNO").Value = PO_ORDER_LNO
            e.Row.Cells("PO_STATUS").Value = "O"
        End If
    End Sub

    Private Sub grdPOTORDR2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR2.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdPOTORDR2.ActiveCell.Column.Key
            Case "ITEM_CODE"

            Case "WHSE_CODE"

            Case "CANCEL"
                e.Cell.Row.Cells("PO_STATUS").Value = "C"
        End Select

        grdClickCellButton(grdPOTORDR2, sql_where, False)
    End Sub

    Function Validate_Columns_2(COLUMN_NAME As String) As Boolean

        Dim Cancel As Boolean = False

        With grdPOTORDR2.ActiveRow
            Select Case COLUMN_NAME
                Case "ITEM_CODE"
                    If .Cells("ITEM_CODE").Value & "" <> "" Then
                        Validate_Item(.Cells("ITEM_CODE").Value)
                        If ITEM_CODE = "" Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If


                Case "PO_DATE_REQUIRED"
                    If Trim$(.Cells("PO_DATE_REQUIRED").Value & "") = "" Then
                        grdPOTORDR2.ActiveCell = .Cells("PO_DATE_REQUIRED")
                        Cancel = True
                    End If

                Case "PO_DATE_REQUESTED"
                    If Trim$(.Cells("PO_DATE_REQUESTED").Value & "") = "" Then
                        grdPOTORDR2.ActiveCell = .Cells("PO_DATE_REQUESTED")
                        Cancel = True
                    End If

                Case "PO_QTY_ORD"

                    If Trim$(.Cells("PO_QTY_ORD").Value & "") = "" Then
                        MsgBox("Order Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                    End If
                    If Val(.Cells("PO_QTY_ORD").Value & "") < 0 Then
                        MsgBox("Order Qty May Not be Negative", vbOKOnly, "Invalid Order Quantity")
                        Cancel = True
                    End If
                    If Val(.Cells("PO_QTY_ORD").Value & "") = 0 Then
                        MsgBox("Order Qty May Not be Zero", 0, "Cannot Update Record")
                        Cancel = True
                    End If
                    Dim q As Int64 = Val(rowICTITEM1.Item("ITEM_STD_PACK_PUR") & "")
                    If q <> 0 Then
                        If Val(.Cells("PO_QTY_ORD").Value & "") Mod q <> 0 Then
                            MsgBox("Order Qty is not an even multiple of Std Pack (" & CStr(q) & "). Please Check", 0, "Warning")
                            ' Cancel = True
                        End If
                    End If
                    q = Val(rowICTITEM1.Item("ITEM_PO_QTY_MULT") & "")
                    If q <> 0 Then
                        If Val(.Cells("PO_QTY_ORD").Value & "") Mod q <> 0 Then
                            MsgBox("Order Qty is not evenly divisible by PO Multiple (" & CStr(q) & "). Please Check", 0, "Warning")
                            ' Cancel = True
                        End If
                    End If
                    If Val(.Cells("PO_QTY_ORD").Value & "") < Val(.Cells("PO_QTY_REC").Value & "") _
                    Or Val(.Cells("PO_QTY_ORD").Value & "") < Val(.Cells("PO_QTY_INV").Value & "") Then
                        MsgBox("Order Qty May Not be Less than Qty Already Received or Invoiced", 0, "Cannot Update Record")
                        Cancel = True
                    End If
                    If Val(.Cells("PO_QTY_BACKORDER").Value & "") <> 0 Then
                        If Val(.Cells("PO_QTY_ORD").Value & "") <= Val(.Cells("PO_QTY_BACKORDER").Value & "") Then
                            MsgBox("Order Qty May Not be less than or equal to PO BackOrder Qty", vbOKOnly, "Invalid Order Quantity")
                            Cancel = True
                        End If
                    End If

                Case "WHSE_CODE"
                    Dim WHSE_CODE As String = .Cells("WHSE_CODE").Value & ""
                    If WHSE_CODE = "" OrElse LookUp("ICTWHSE1", WHSE_CODE) Is Nothing Then
                        grdPOTORDR2.ActiveCell = .Cells("WHSE_CODE")
                        Cancel = True
                    End If

                Case "PO_QTY_BACKORDER"
                    If Val(.Cells("PO_QTY_BACKORDER").Value & "") <> 0 Then
                        If Val(.Cells("PO_QTY_BACKORDER").Value & "") >= Val(.Cells("PO_QTY_ORD").Value & "") Then
                            MsgBox("Backorder Qty (" & CStr(Val(.Cells("PO_QTY_BACKORDER").Value & "")) & ") May not be greater than or equal to the Qty Ord (" & CStr(Val(.Cells("PO_QTY_OPN_CALC").Value & "")) & ")", 0, "Cannot Update Record")
                            Cancel = True
                        End If

                    End If


                Case "PO_DATE_BACKORDER"
                    Dim BO_DATE As String = .Cells("PO_DATE_BACKORDER").Value & ""
                    Dim REQ_DATE As String = .Cells("PO_DATE_REQUIRED").Value & ""
                    If BO_DATE = REQ_DATE Then
                        ' If Trim$(.Cells("PO_DATE_BACKORDER").Value & "") <> "" And Trim$(.Cells("BACKORDER_DATE").Value & "") = Trim$(.Cells("PO_DATE_REQUIRED").Value & "") Then
                        MsgBox("Backorder Date cannot be the same as Date Required", 0, "Cannot Update Record")
                        Cancel = True
                    End If

                    ' 
            End Select
        End With
        Return Cancel
    End Function

    Sub Setup_grdICTIRECX()

        If EntryMode <> "V" Then
            splDetails.Panel2Collapsed = True
            Exit Sub
        End If
        If grdPOTORDR2.ActiveRow Is Nothing OrElse (Not grdPOTORDR2.ActiveRow.IsDataRow Or grdPOTORDR2.ActiveRow.IsAddRow) Then
            splDetails.Panel2Collapsed = True
        Else
            Dim dvw As DataView = DirectCast(grdICTIRECX.DataSource, DataTable).DefaultView
            Dim PO_ORDER_LNO As Integer = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value)
            dvw.RowFilter = "PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            Sort_grdColumns(grdICTIRECX, "RECEIPT_NO")
            tabPOTORDR2.Tabs("PO Receipt Details").Text = "PO Receipt Details for Line " & CStr(PO_ORDER_LNO)
            'grdICTIRECX.Text = "PO Receipt Details for Line " & CStr(PO_ORDER_LNO)
            splDetails.Panel2Collapsed = False
            grdICTIRECX.Rows.ExpandAll(True)
        End If
    End Sub

    Sub Setup_grdPOTORDR9()

        If EntryMode <> "V" Then
            splDetails.Panel2Collapsed = True
            Exit Sub
        End If
        If grdPOTORDR2.ActiveRow Is Nothing OrElse (Not grdPOTORDR2.ActiveRow.IsDataRow Or grdPOTORDR2.ActiveRow.IsAddRow) Then
            splDetails.Panel2Collapsed = True
        Else
            Dim dvw As DataView = DirectCast(grdPOTORDR9.DataSource, DataTable).DefaultView
            Dim PO_ORDER_LNO As Integer = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value)
            dvw.RowFilter = "PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            Sort_grdColumns(grdPOTORDR9, "ITEM_CODE")
            tabPOTORDR2.Tabs("Component Commitments").Text = "Component Commitments for Line " & CStr(PO_ORDER_LNO)
            'grdICTIRECX.Text = "Component Committments for Line " & CStr(PO_ORDER_LNO)
            splDetails.Panel2Collapsed = False
        End If
    End Sub

    Sub Setup_grdICTPINVD()
        If EntryMode <> "V" Then
            splDetails.Panel2Collapsed = True
            Exit Sub
        End If
        If grdPOTORDR2.ActiveRow Is Nothing OrElse (Not grdPOTORDR2.ActiveRow.IsDataRow Or grdPOTORDR2.ActiveRow.IsAddRow) Then
            splDetails.Panel2Collapsed = True
        Else
            Dim dvw As DataView = DirectCast(grdICTPINVD.DataSource, DataTable).DefaultView
            Dim PO_ORDER_LNO As Integer = Val(grdPOTORDR2.ActiveRow.Cells("PO_ORDER_LNO").Value)
            dvw.RowFilter = "PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            Sort_grdColumns(grdICTPINVD, "ETA_DATE_DC")
            tabPOTORDR2.Tabs("Invoices").Text = "Invoices for Line " & CStr(PO_ORDER_LNO)
            splDetails.Panel2Collapsed = False
        End If
    End Sub
#End Region

#Region "grdPOTORDR5"

    Private Sub grdPOTORDR5_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR5.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "NINV_CODE"
                Dim NINV_CODE As String = e.Cell.Value & ""
                Dim rowPOTNINV1 As DataRow = LookUp("POTNINV1", NINV_CODE)
                If rowPOTNINV1 IsNot Nothing Then
                    e.Cell.Row.Cells("PO_NINV_DESC").Value = rowPOTNINV1.Item("NINV_DESC") & ""
                    'If e.Cell.Row.IsAddRow Then
                    '    e.Cell.Row.Cells("PO_COST").Value = Get_Price()
                    '    grdPOTORDR2.ActiveCell = grdPOTORDR2.ActiveRow.Cells("PO_COST")
                    'End If
                End If

            Case "PO_NINV_QTY"
                Dependent_Fields_5()
            Case "PO_NINV_PRICE"
                Dependent_Fields_5()
            Case "PO_NINV_AMOUNT"
                Dim PO_NINV_QTY As Int64 = Val(e.Cell.Row.Cells("PO_NINV_QTY").Value & "")
                Dim PO_NINV_PRICE As Decimal = Val(e.Cell.Row.Cells("PO_NINV_PRICE").Value & "")
                Dim PO_NINV_AMOUNT As Decimal = Val(e.Cell.Row.Cells("PO_NINV_AMOUNT").Value & "")
                If System.Math.Abs(PO_NINV_QTY * PO_NINV_PRICE - PO_NINV_AMOUNT) > 0.005 Then
                    e.Cell.Row.Cells("PO_NINV_QTY").Value = 0
                    e.Cell.Row.Cells("PO_NINV_QTY_OPN").Value = 0
                    e.Cell.Row.Cells("PO_NINV_PRICE").Value = 0
                    Dim PO_NINV_AMT_REC As Decimal = Val(e.Cell.Row.Cells("PO_NINV_AMT_REC").Value & "")
                    Dim PO_NINV_AMT_INV As Decimal = Val(e.Cell.Row.Cells("PO_NINV_AMT_INV").Value & "")
                    Dim PO_NINV_AMT_CLS As Decimal = IIf(PO_NINV_AMT_REC > PO_NINV_AMT_INV, PO_NINV_AMT_REC, PO_NINV_AMT_INV)
                    Dim PO_NINV_AMT_OPN As Decimal = PO_NINV_AMOUNT - PO_NINV_AMT_CLS
                    If PO_NINV_AMT_OPN < 0 Then PO_NINV_AMT_OPN = 0
                    e.Cell.Row.Cells("PO_NINV_AMT_OPN").Value = PO_NINV_AMT_OPN
                End If
            Case "PO_NINV_DATE_REQ"
                Dependent_Fields_5()
        End Select
    End Sub

    Private Sub grdPOTORDR5_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTORDR5.AfterExitEditMode
        With grdPOTORDR5
            Select Case .ActiveCell.Column.Key
                Case "BM_COMP_ITEM"
                    If .ActiveCell.Text <> "" Then
                        '.ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
                        cdr = LookUp("ICTITEM1", .ActiveCell.Text)
                        .ActiveRow.Cells("ITEM_DESC").Value = cdr.Item("ITEM_DESC") & ""
                        .ActiveRow.Cells("ITEM_UOM").Value = cdr.Item("ITEM_UOM") & ""
                        .ActiveRow.Cells("ITEM_COST_STD").Value = Val(cdr.Item("ITEM_COST_STD") & "")
                        .ActiveRow.Cells("ITEM_COST_WASTE_PCT").Value = Val(cdr.Item("ITEM_COST_WASTE_PCT") & "")
                        .ActiveRow.Cells("VEND_ITEM_CODE").Value = cdr.Item("VEND_ITEM_CODE") & ""
                    End If
            End Select
        End With
    End Sub

    Private Sub grdPOTORDR5_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDR5.AfterRowActivate
        With grdPOTORDR5.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"PO_NINV_DESC", "PO_NINV_QTY", "PO_NINV_PRICE", "PO_NINV_DATE_REQ", "PO_NINV_AMOUNT"}
                If grdPOTORDR5.ActiveRow.Cells("PO_NINV_STATUS").Value & "" = "C" Then
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next

            If Val(grdPOTORDR5.ActiveRow.Cells("PO_NINV_CTR_REC").Value & "") <> 0 _
            Or Val(grdPOTORDR5.ActiveRow.Cells("PO_NINV_CTR_INV").Value & "") <> 0 Then
                If Val(grdPOTORDR5.ActiveRow.Cells("PO_NINV_QTY").Value & "") <> 0 Then
                    .Columns("PO_NINV_AMOUNT").CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns("PO_NINV_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("PO_NINV_PRICE").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End If
        End With

        With grdPOTORDR5.ActiveRow
            If .Cells("PO_NINV_DESC").Value & "" = "" Then
                grdPOTORDR5.ActiveCell = .Cells("PO_NINV_DESC")
            End If
        End With

        Setup_grdPOTORDR6()
    End Sub

    Private Sub grdPOTORDR5_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTORDR5.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdPOTORDR5_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDR5.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdPOTORDR5_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTORDR5.BeforeExitEditMode

        If e.CancellingEditOperation Then Exit Sub
        With grdPOTORDR5.ActiveCell
            Select Case .Column.Key
                Case "PO_NINV_DATE_REQ"
                    ' .Value = DATE_CHECK(.Value, Cancel)
                Case Else
                    ' e.Cancel = Validate_Columns_5(.Column.Key)
            End Select
        End With
    End Sub

    Private Sub grdPOTORDR5_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTORDR5.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Val(grow.Cells("PO_NINV_CTR_REC").Value & "") <> 0 _
            Or Val(grow.Cells("PO_NINV_CTR_INV").Value & "") <> 0 Then
                MsgBox("PO Line " & grow.Cells("PO_ORDER_LNO").Value & " has had Receipts and/or Invoices Posted Against it",
                        MsgBoxStyle.OkOnly, "Cannot Delete this Record")
                e.Cancel = True
            End If
        Next
    End Sub

    Private Sub grdPOTORDR5_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTORDR5.BeforeRowUpdate
        'If Validate_Columns_5("PO_NINV_QTY") Then e.Cancel = True
        'If Validate_Columns_5("PO_NINV_DATE_REQ") Then e.Cancel = True
        Dim NINV_CODE As String = e.Row.Cells("NINV_CODE").Value & ""
        If NINV_CODE <> "" Then
            Dim rowPOTNINV1 As DataRow = LookUp("POTNINV1", NINV_CODE)
            If rowPOTNINV1 Is Nothing Then
                e.Cancel = True
            End If
        End If

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("PO_NINV_LNO").Value = Val(dst.Tables("POTORDR5").Compute("MAX(PO_NINV_LNO)", "") & "") + 1
            e.Row.Cells("PO_ORDER_NO").Value = PO_ORDER_NO
            e.Row.Cells("PO_NINV_STATUS").Value = "O"

            If e.Row.Cells("PO_NINV_DATE_REQ").Value & "" = "" And Absx1.dteFor("PO_DATE_REQUIRED").Value & "" <> "" Then
                e.Row.Cells("PO_NINV_DATE_REQ").Value = Absx1.dteFor("PO_DATE_REQUIRED").Value
            End If
        End If
    End Sub

    Private Sub grdPOTORDR5_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR5.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdPOTORDR5.ActiveCell.Column.Key
            Case "CANCEL"
                e.Cell.Row.Cells("PO_NINV_STATUS").Value = "C"
            Case Else
                grdClickCellButton(grdPOTORDR5, sql_where, False)
        End Select
    End Sub

    Sub Dependent_Fields_5()
        With grdPOTORDR5.ActiveRow
            Dim PO_NINV_QTY As Int64 = Val(.Cells("PO_NINV_QTY").Value & "")
            Dim PO_NINV_PRICE As Decimal = Val(.Cells("PO_NINV_PRICE").Value & "")
            Dim PO_NINV_STATUS As String = .Cells("PO_NINV_STATUS").Value & ""

            If PO_NINV_QTY <> 0 Or PO_NINV_PRICE <> 0 Then
                .Cells("PO_NINV_AMOUNT").Value = PO_NINV_QTY * PO_NINV_PRICE

                Dim PO_NINV_QTY_REC As Int64 = Val(.Cells("PO_NINV_QTY_REC").Value & "")
                Dim PO_NINV_QTY_INV As Int64 = Val(.Cells("PO_NINV_QTY_INV").Value & "")
                Dim PO_NINV_QTY_CLS As Int64 = PO_NINV_QTY_REC
                If PO_NINV_QTY_INV > PO_NINV_QTY_REC Then PO_NINV_QTY_CLS = PO_NINV_QTY_INV
                Dim PO_NINV_QTY_OPN As Int64 = PO_NINV_QTY - PO_NINV_QTY_CLS
                If PO_NINV_QTY_OPN < 0 Or PO_NINV_STATUS = "C" Then PO_NINV_QTY_OPN = 0
                .Cells("PO_NINV_QTY_OPN").Value = PO_NINV_QTY_OPN
                .Cells("PO_NINV_AMT_OPN").Value = PO_NINV_QTY_OPN * PO_NINV_PRICE
            Else
                Dim PO_NINV_AMOUNT As Int64 = Val(.Cells("PO_NINV_AMOUNT").Value & "")
                Dim PO_NINV_AMT_REC As Int64 = Val(.Cells("PO_NINV_AMT_REC").Value & "")
                Dim PO_NINV_AMT_INV As Int64 = Val(.Cells("PO_NINV_AMT_INV").Value & "")
                Dim PO_NINV_AMT_CLS As Int64 = PO_NINV_AMT_REC
                If PO_NINV_AMT_INV > PO_NINV_AMT_REC Then PO_NINV_AMT_CLS = PO_NINV_AMT_INV
                Dim PO_NINV_AMT_OPN As Int64 = PO_NINV_AMOUNT - PO_NINV_AMT_CLS
                If PO_NINV_AMT_OPN < 0 Or PO_NINV_STATUS = "C" Then PO_NINV_AMT_OPN = 0
                .Cells("PO_NINV_AMT_OPN").Value = PO_NINV_AMT_OPN
            End If
        End With
    End Sub

    Sub Setup_grdPOTORDR6()
        If grdPOTORDR5.ActiveRow Is Nothing OrElse (Not grdPOTORDR5.ActiveRow.IsDataRow Or grdPOTORDR5.ActiveRow.IsAddRow) Then
            splNonInventory.Panel2Collapsed = True
        Else
            Dim dvw As DataView = DirectCast(grdPOTORDR6.DataSource, DataTable).DefaultView
            Dim PO_NINV_LNO As Integer = Val(grdPOTORDR5.ActiveRow.Cells("PO_NINV_LNO").Value)
            dvw.RowFilter = "PO_NINV_LNO = " & CStr(PO_NINV_LNO)
            grdPOTORDR6.Text = "Non-Inventory Receipt / Invoice Details for Line " & CStr(PO_NINV_LNO)
            splNonInventory.Panel2Collapsed = False
        End If
    End Sub
#End Region

#Region "grdPOTORDR6"

    Private Sub grdPOTORDR6_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR6.AfterCellUpdate
        Select Case e.Cell.Column.Key

        End Select
    End Sub

    Private Sub grdPOTORDR6_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTORDR6.AfterExitEditMode
        With grdPOTORDR6
            Select Case .ActiveCell.Column.Key

            End Select
        End With
    End Sub

    Private Sub grdPOTORDR6_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdPOTORDR6.AfterRowActivate
        With grdPOTORDR6.DisplayLayout.Bands(0)

        End With

        With grdPOTORDR6.ActiveRow

        End With
    End Sub

    Private Sub grdPOTORDR6_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdPOTORDR6.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdPOTORDR6_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdPOTORDR6.AfterRowUpdate
        Display_Totals()
    End Sub

    Private Sub grdPOTORDR6_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdPOTORDR6.BeforeExitEditMode
        If e.CancellingEditOperation Then Exit Sub
        With grdPOTORDR6.ActiveCell
            Select Case .Column.Key

            End Select
        End With
    End Sub

    Private Sub grdPOTORDR6_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdPOTORDR6.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In e.Rows

        Next
    End Sub

    Private Sub grdPOTORDR6_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdPOTORDR6.BeforeRowUpdate
        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("PO_ORDER_NO").Value = PO_ORDER_NO
            e.Row.Cells("PO_NINV_LNO").Value = grdPOTORDR5.ActiveRow.Cells("PO_NINV_LNO").Value
            e.Row.Cells("PO_NINV_SLNO").Value = Val(dst.Tables("POTORDR6").Compute("MAX(PO_NINV_SLNO)", "") & "") + 1
        End If
    End Sub

    Private Sub grdPOTORDR6_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdPOTORDR6.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdPOTORDR6.ActiveCell.Column.Key

            Case Else
                grdClickCellButton(grdPOTORDR6, sql_where, False)
        End Select
    End Sub
#End Region

    Private Sub chkShipToRel_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShipToRel.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        For Each COLUMN_NAME In New String() _
        {"SHIP_TO_NAME", "SHIP_TO_ADDR1", "SHIP_TO_ADDR2", "SHIP_TO_ADDR3" _
         , "SHIP_TO_CITY", "SHIP_TO_STATE", "SHIP_TO_ZIP_CODE", "SHIP_TO_COUNTRY" _
         , "SHIP_TO_PHONE", "SHIP_TO_EXT", "SHIP_TO_FAX", "SHIP_TO_CONTACT", "SHIP_TO_EMAIL"}
            If COLUMN_NAME = "SHIP_TO_PHONE" Or COLUMN_NAME = "SHIP_TO_FAX" Then
                Absx1.medFor(COLUMN_NAME).ReadOnly = chkShipToRel.Checked
            Else
                Absx1.txtFor(COLUMN_NAME).ReadOnly = chkShipToRel.Checked
            End If
        Next
    End Sub

    Sub Display_Totals()
        Dim ORD_QTY_I As Int32 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_ORD)", "") & "")
        Dim REC_QTY_I As Int32 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_REC)", "") & "")
        Dim INV_QTY_I As Int32 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_INV)", "") & "")
        Dim OPN_QTY_I As Int32 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN_CALC)", "") & "")


        Dim ORD_AMT_I As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
        Dim REC_AMT_I As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_REC)", "") & "")
        Dim INV_AMT_I As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_INV)", "") & "")
        Dim OPN_AMT_I As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_OPN_CALC)", "") & "")

        Dim ORD_QTY_N As Int32 = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_QTY)", "") & "")
        Dim REC_QTY_N As Int32 = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_QTY_REC)", "") & "")
        Dim INV_QTY_N As Int32 = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_QTY_INV)", "") & "")
        Dim OPN_QTY_N As Int32 = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_QTY_OPN)", "") & "")

        Dim ORD_AMT_N As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMOUNT)", "") & "")
        Dim REC_AMT_N As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMT_REC)", "") & "")
        Dim INV_AMT_N As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMT_INV)", "") & "")
        Dim OPN_AMT_N As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMT_OPN)", "") & "")

        grdPOTORDRT.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdPOTORDRT.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.ExternalSortSingle
        With dst.Tables("POTORDRT").Rows
            .Clear()
            .Add(New Object() {"Ordered", ORD_QTY_I, ORD_AMT_I, ORD_QTY_N, ORD_AMT_N, ORD_AMT_I + ORD_AMT_N})
            .Add(New Object() {"Received", REC_QTY_I, REC_AMT_I, REC_QTY_N, REC_AMT_N, REC_AMT_I + REC_AMT_N})
            .Add(New Object() {"Invoiced", INV_QTY_I, INV_AMT_I, INV_QTY_N, INV_AMT_N, INV_AMT_I + INV_AMT_N})
            .Add(New Object() {"Open", OPN_QTY_I, OPN_AMT_I, OPN_QTY_N, OPN_AMT_N, OPN_AMT_I + OPN_AMT_N})
        End With
    End Sub

    Sub Dependent_Updates(S As Integer, kv As String)

        If S = -1 Then
            TAC.POCMAIN1.Production_Commit(-1, PO_ORDER_NO)
            TAC.POCMAIN1.ICTSTAT2_PO(-1, PO_ORDER_NO)
        Else
            TAC.POCMAIN1.ICTSTAT2_PO(1, PO_ORDER_NO)
            If rowPOTORDR1.Item("VEND_WHSE_CODE") & "" <> "" Then
                TAC.POCMAIN1.Update_POTORDR9(Me, PO_ORDER_NO, rowPOTORDR1.Item("VEND_WHSE_CODE") & "")
            End If
        End If
    End Sub

    Function Get_Price(OPS_YYYYPP As String) As Decimal
        ' this routine needs to be adjusted for for curr
        Dim ITEM_COST_VCURR As Decimal = 0
        ' Retrieve cost from ICTCOSTP
        ASCMAIN1.sql = $"select * from ICTCOSTP where OPS_YYYYPP<='{OPS_YYYYPP}' and item_code='{ITEM_CODE}' ORDER BY OPS_YYYYPP DESC FETCH FIRST 1 ROW ONLY"
        Dim rowICTCOSTP As DataRow = ASCDATA1.GetDataRow()
        If rowICTCOSTP IsNot Nothing Then
            ITEM_COST_VCURR = Val(rowICTCOSTP.Item("ITEM_COST_VCOST") & "")
            ' when 0 rows, move on to costf
        Else
            ' If not available, retrieve cost from ICTCOSTF

            ' If not available, retrieve cost from ICTCOSTC

            Dim rowICTCOSTF As DataRow = LookUp("ICTCOSTF", ITEM_CODE)
            If rowICTCOSTF Is Nothing Then
                Dim rowICTCOSTC As DataRow = LookUp("ICTCOSTC", ITEM_CODE, True)
                ITEM_COST_VCURR = Val(rowICTCOSTC.Item("ITEM_COST_VCURR") & "")
            Else
                ITEM_COST_VCURR = Val(rowICTCOSTF.Item("ITEM_COST_VCURR") & "")
            End If
        End If

        Return ITEM_COST_VCURR
    End Function

    Sub Setup_BM()
        Dim BM_ISSUE_TYPE As String = "A"
        If optPO_ORDER_TYPE.Value = "R" Then BM_ISSUE_TYPE = "R"
        Fill_Records("BMTMAIN2", New String() {ITEM_CODE, BM_ISSUE_TYPE})
        '  Sort_grdColumns(GRDBMTMAIN2, "BM_ISSUE_NO".ToLower)
    End Sub

    Sub Load_POTORDRX()
        If Me.IsClosing Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If POTORDRX = "" Then
            ASCMAIN1.sql = "Select POTORDR1.* from POTORDR1 where ROWNUM < 1"
            POTORDRX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & POTORDRX & " Add Primary Key (PO_ORDER_NO)")

            sqlPOTORDRT = "Select POTORDR2.PO_ORDER_NO" & vbCrLf _
                & ", SUM (PO_QTY_ORD) PO_QTY_ORD" & vbCrLf _
                & ", SUM (PO_QTY_REC) PO_QTY_REC" & vbCrLf _
                & ", SUM (PO_QTY_INV) PO_QTY_INV" & vbCrLf _
                & ", SUM (PO_QTY_OPN) PO_QTY_OPN" & vbCrLf _
                & ", SUM (NVL(PO_QTY_ORD,0) * NVL(PO_COST,0)) PO_AMT_ORD" & vbCrLf _
                & ", SUM (NVL(PO_QTY_REC,0) * NVL(PO_COST,0)) PO_AMT_REC" & vbCrLf _
                & ", SUM (NVL(PO_QTY_INV,0) * NVL(PO_COST,0)) PO_AMT_INV" & vbCrLf _
                & ", SUM (NVL(PO_QTY_OPN,0) * NVL(PO_COST,0)) PO_AMT_OPN" & vbCrLf _
                & " from POTORDR2 where PO_ORDER_NO in (Select PO_ORDER_NO from " & POTORDRX & ")" & vbCrLf _
                & " group by POTORDR2.PO_ORDER_NO"
            POTORDRT = ASCMAIN1.Temp_Table(sqlPOTORDRT)
            ASCDATA1.ExecuteSQL("Alter Table " & POTORDRT & " Add Primary Key (PO_ORDER_NO)")

            sqlPOTORDRN = "Select POTORDR5.PO_ORDER_NO" & vbCrLf _
                & ", COUNT (*) PO_NINV_LINES" & vbCrLf _
                & ", SUM (PO_NINV_QTY_OPN) PO_NINV_QTY_OPN" & vbCrLf _
                & ", SUM (PO_NINV_AMT_OPN) PO_NINV_AMT_OPN" & vbCrLf _
                & " from POTORDR5 where PO_ORDER_NO in (Select PO_ORDER_NO from " & POTORDRX & ")" & vbCrLf _
                & " group by POTORDR5.PO_ORDER_NO"
            POTORDRN = ASCMAIN1.Temp_Table(sqlPOTORDRN)
            ASCDATA1.ExecuteSQL("Alter Table " & POTORDRN & " Add Primary Key (PO_ORDER_NO)")

            Exit Sub
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & POTORDRX)
            Dim sqlw As String = ""
            Select Case optStatus.Value
                Case "A"
                    grdPOTORDRX.Text = "All POs"
                Case "O"
                    sqlw &= " and PO_STATUS = 'O'"
                    grdPOTORDRX.Text = "Open POs"
                Case "W"
                    sqlw &= " and PO_STATUS = 'O' and NVL(PO_APPR_PENDING,'0') = '0' and PO_APPR_BY is Null"
                    grdPOTORDRX.Text = "Open POs Not Ready for Approval"
                Case "a"
                    sqlw &= " and PO_STATUS = 'O' and NVL(PO_APPR_PENDING,'0') = '1' and PO_APPR_BY is Null"
                    grdPOTORDRX.Text = "Open POs Pending Approval"
                Case "t"
                    sqlw &= " and PO_STATUS = 'O' and PO_APPR_BY is Not Null and NVL(PO_XMIT_IND,0) = '0'"
                    ' sqlw &= " and PO_STATUS = 'O' and PO_APPR_BY is Not Null and NVL(PO_XMIT_IND,0) = '0' and not (VEND_CODE = 'IPSA' and WHSE_CODE = 'ADS')"
                    grdPOTORDRX.Text = "Open POs Pending Transmit"
                Case "T"
                    sqlw &= " and PO_STATUS = 'O' and PO_APPR_BY is Not Null and NVL(PO_XMIT_IND,0) = '1'"
                    grdPOTORDRX.Text = "Open POs Already Transmitted"
            End Select

            If Absx1.txtFor("VEND_CODE").Text <> "" And Absx1.txtFor("VEND_NAME").Text <> "" Then
                sqlw &= " and VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
                grdPOTORDRX.Text &= ", Supplier " & Absx1.txtFor("VEND_CODE").Text
            End If

            ASCMAIN1.sql = "Select * from POTORDR1" & ASCMAIN1.SQL_Add_WHERE(sqlw)
            ASCDATA1.ExecuteSQL("Insert into " & POTORDRX & " " & ASCMAIN1.sql)

            ASCDATA1.ExecuteSQL("Truncate Table " & POTORDRT)
            ASCDATA1.ExecuteSQL("Insert into " & POTORDRT & " " & sqlPOTORDRT)

            ASCDATA1.ExecuteSQL("Truncate Table " & POTORDRN)
            ASCDATA1.ExecuteSQL("Insert into " & POTORDRN & " " & sqlPOTORDRN)
        End If

        Fill_Records("POTORDRX")
        Sort_grdColumns(grdPOTORDRX, "PO_ORDER_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Validate_Item(ITEM_CODE_z As String) As Boolean
        ITEM_CODE = ""
        rowICTITEM1 = Nothing
        If ITEM_CODE_z = "" Then Return False

        Dim E As String = ""
        rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE_z)

        If rowICTITEM1 Is Nothing Then
            E = "Item is Not on File" & vbCrLf
        Else
            If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
                E = "Item Status is not Active" & vbCrLf
            End If
            If rowICTITEM1.Item("ITEM_UOM") & "" = "" Then
                E = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTITEM1.Item("COLLECTION_CODE") & "" = "" Then
                E = "Item does not have a valid Collection Code" & vbCrLf
            End If
        End If

        If E <> "" And grdPOTORDR2.ActiveRow.IsAddRow Then
            MsgBox(E, MsgBoxStyle.OkOnly, "Item Code entered is invalid because ...")
        Else
            If E = "" Then
                ITEM_CODE = rowICTITEM1.Item("ITEM_CODE")
            End If
        End If
        Return (ITEM_CODE <> "")
    End Function

    Function Validate_Vendor() As Boolean
        VEND_CODE = ""
        rowAPTVEND1 = Nothing

        If Absx1.txtFor("VEND_CODE").Text = "" Then
            Return False
        End If

        rowAPTVEND1 = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)

        If rowAPTVEND1 Is Nothing Then
            EMsg &= vbCr & "Vendor is Not on File" & vbCrLf
        Else
            If rowAPTVEND1.Item("VEND_STATUS") & "" <> "A" Then
                EMsg &= vbCr & "Vendor Status is not Active" & vbCrLf
            End If
            If rowAPTVEND1.Item("VEND_TYPE") & "" <> "S" Then
                EMsg &= vbCr & "Vendor is not set up as a Supplier" & vbCrLf
            End If
        End If

        If EMsg = "" Then
            VEND_CODE = rowAPTVEND1.Item("VEND_CODE")
            'If rowAPTVEND1.Item("VEND_WHSE_CODE") & "" <> "" Then
            '    optPO_TYPE.CheckedIndex = 0 ' "M"
            'Else
            '    optPO_TYPE.CheckedIndex = 1 ' "B"
            'End If
        End If

        Return (VEND_CODE <> "")
    End Function

    Function Check_Changed_Fields() As Boolean

        Dim PO_HDR_CTR_REV As Integer = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")
        PO_HDR_CTR_REV += 1

        Dim LAST_DATE As Date = DATETIME_STAMP
        If EntryMode = "N" Then Stop
        Dim REV_LNO As Integer = 0

        Check_Changed_Fields = False

        dst.Tables("POTORDXR").Rows.Clear()

        ASCMAIN1.Progress("Logging Header Changes")

        For i As Integer = 0 To rowPOTORDR1.Table.Columns.Count - 1
            Dim COLUMN_NAME As String = dst.Tables("POTORDR1").Columns(i).ColumnName

            If rowPOTORDR1.Item(COLUMN_NAME) & "" _
            <> rowPOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                Check_Changed_Fields = True
                ASCMAIN1.Progress("-", COLUMN_NAME)
                Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
                With rowPOTORDXR
                    .Item("REV_NO") = PO_HDR_CTR_REV
                    REV_LNO += 1
                    .Item("REV_LNO") = REV_LNO
                    .Item("PO_ORDER_NO") = PO_ORDER_NO
                    .Item("PO_ORDER_LNO") = 0
                    .Item("INIT_DATE") = LAST_DATE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("COLUMN_NAME") = COLUMN_NAME
                    .Item("OLD_VALUE") = rowPOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original)
                    .Item("NEW_VALUE") = rowPOTORDR1.Item(COLUMN_NAME)
                    .Item("EMODE") = EntryMode
                End With
                dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                Check_Changed_Fields = True
            End If
        Next i

        ASCMAIN1.Progress("Logging Detail Changes")

        ASCMAIN1.sql = "Select * from POTORDR2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)

        For Each rowPOTORDR2_orig As DataRow In dt.Rows
            Dim PO_ORDER_LNO As Int64 = rowPOTORDR2_orig.Item("PO_ORDER_LNO")
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {PO_ORDER_NO, PO_ORDER_LNO})
            If rowPOTORDR2 Is Nothing Then ' Line was Deleted
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowPOTORDR2_orig.Table.Columns(i).ColumnName
                    Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
                    With rowPOTORDXR
                        .Item("REV_NO") = PO_HDR_CTR_REV
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("PO_ORDER_NO") = PO_ORDER_NO
                        .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                        .Item("INIT_DATE") = LAST_DATE
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = COLUMN_NAME
                        .Item("OLD_VALUE") = rowPOTORDR2_orig.Item(COLUMN_NAME)
                        '.Item("NEW_VALUE") = ""
                        .Item("EMODE") = EntryMode
                    End With
                    dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                Next

                Check_Changed_Fields = True
            Else
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim COLUMN_NAME As String = rowPOTORDR2_orig.Table.Columns(i).ColumnName
                    If rowPOTORDR2.Item(COLUMN_NAME) & "" <> rowPOTORDR2_orig.Item(COLUMN_NAME) & "" Then
                        ' Value in Column was Changed
                        Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
                        With rowPOTORDXR
                            .Item("REV_NO") = PO_HDR_CTR_REV
                            REV_LNO += 1
                            .Item("REV_LNO") = REV_LNO
                            .Item("PO_ORDER_NO") = PO_ORDER_NO
                            .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                            .Item("INIT_DATE") = LAST_DATE
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("COLUMN_NAME") = COLUMN_NAME
                            .Item("OLD_VALUE") = rowPOTORDR2_orig.Item(COLUMN_NAME)
                            .Item("NEW_VALUE") = rowPOTORDR2.Item(COLUMN_NAME)
                            .Item("EMODE") = EntryMode
                        End With
                        dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
                        Check_Changed_Fields = True
                    End If
                Next
            End If
        Next

        For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("", "", DataViewRowState.Added)
            Dim PO_ORDER_LNO = rowPOTORDR2.Item("PO_ORDER_LNO")
            ' For i As Integer = 0 To dt.Columns.Count - 1
            Dim COLUMN_NAME As String = "" ' dt.Columns(i).ColumnName
            Dim rowPOTORDXR As DataRow = dst.Tables("POTORDXR").NewRow
            With rowPOTORDXR
                .Item("REV_NO") = PO_HDR_CTR_REV
                REV_LNO += 1
                .Item("REV_LNO") = REV_LNO
                .Item("PO_ORDER_NO") = PO_ORDER_NO
                .Item("PO_ORDER_LNO") = PO_ORDER_LNO
                .Item("INIT_DATE") = LAST_DATE
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("COLUMN_NAME") = COLUMN_NAME
                '.Item("OLD_VALUE") = ""
                .Item("NEW_VALUE") = "PO Line Added" ' rowPOTORDR2.Item(COLUMN_NAME)
                .Item("EMODE") = EntryMode
            End With
            dst.Tables("POTORDXR").Rows.Add(rowPOTORDXR)
            Check_Changed_Fields = True
            'Next
        Next

        ASCMAIN1.Progress("")
        Return Check_Changed_Fields
    End Function

    Sub Record_Event(EVENT_TYPE As String, EVENT_DESC As String)
        Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
        row.Item("TABLE_NAME") = "POTORDR1"
        row.Item("TABLE_KEY") = PO_ORDER_NO
        row.Item("INIT_DATE") = DATETIME_STAMP
        row.Item("INIT_OPER") = ASCMAIN1.USER_ID
        row.Item("EVENT_TYPE") = EVENT_TYPE
        row.Item("EVENT_DESC") = EVENT_DESC
        dst.Tables("TATEVNT1").Rows.Add(row)
    End Sub

    Private Sub grdPOTORDRX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTORDRX.DoubleClickRow
        Absx1.txtFor("PO_ORDER_NO").Text = e.Row.Cells("PO_ORDER_NO").Value
        Click_Command("View")
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing PO")
        Print_PO(PO_ORDER_NO)

        'Print_Report_Begin()
        '' CR_params.Add("NOTES", "1")
        'Generate_Report("POROPRT1", "Purchase Order", "")
        'Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_POTORDRX()
    End Sub

    Sub Update_Approval()

        Dim PO_AMT_ORD As Decimal = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
        Dim PO_NINV_AMOUNT As Decimal = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMOUNT)", "") & "")
        Dim PO_TOTAL_AMT As Decimal = PO_AMT_ORD + PO_NINV_AMOUNT

        BeginTrans()

        If preapproval_applied Then
            ' leave the original approval intact
        Else
            ASCMAIN1.sql = "Update POTORDR1 Set PO_APPR_DATE = :PARM1, PO_APPR_BY = :PARM2, PO_APPR_AMOUNT = :PARM3, PO_APPR_NOTES = :PARM4 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, PO_TOTAL_AMT, APPR_NOTES})
        End If

        Dim row As DataRow = dst.Tables("TATEVNT1").NewRow
        row.Item("TABLE_NAME") = "POTORDR1"
        row.Item("TABLE_KEY") = PO_ORDER_NO
        row.Item("INIT_DATE") = DATETIME_STAMP
        row.Item("INIT_OPER") = ASCMAIN1.USER_ID
        row.Item("EVENT_TYPE") = "POAPPR"
        row.Item("EVENT_DESC") = "Approved for " & Format(PO_TOTAL_AMT, "$#,##0.00") & "; " & APPR_NOTES
        dst.Tables("TATEVNT1").Rows.Add(row)

        Update_Record_TDA("TATEVNT1")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
        CommitTrans("Approval Complete")

        preapproval_applied = False
    End Sub

    Function Print_BM(ITEM_CODE As String,
                      BM_ISSUE_STATUS As String,
                      TRANSMIT_KITTING As Boolean,
                      make_pdf As Boolean,
                      FILENAME_body As String) As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing BMs")

        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
        Dim WHSE_CODE As String = rowPOTORDR1.Item("WHSE_CODE") & String.Empty
        Dim drICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim REPORT_NO As String = ""

        Dim LP_CODE As String = drICTWHSE1.Item("LP_CODE") & String.Empty
        Select Case LP_CODE
            Case "ADS"
                ' Need to call the new class
                ' ISSUE 7142
                If TRANSMIT_KITTING Then

                    Dim PlannedQuantity As Int32 = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", $"ITEM_CODE = '{ITEM_CODE}'") & String.Empty)
                    If PlannedQuantity > 0 Then
                        Dim APP_KEY As String = ITEM_CODE & "," & BM_ISSUE_STATUS & "," & PlannedQuantity & "," & rowPOTORDR1.Item("PO_ORDER_NO")

                        Dim XMIT_NO As String = String.Empty
                        XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "All", LP_CODE)

                        ' Only send BM when we have not sent the Item in the WWIMPZMFG*.XML file.
                        ' We may put a flag in BMTMAIN1
                        XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHCKITO1", "BM", APP_KEY, LP_CODE)

                        Dim rowWHT3PLX1 As DataRow = LookUp("WHT3PLX1", XMIT_NO)
                        If rowWHT3PLX1 IsNot Nothing AndAlso Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
                            ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)
                                        Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_KIT','Kitting PO Transmitted to {LP_CODE}', '{XMIT_NO}'
                                                From POTORDR1 where PO_ORDER_NO = '{PO_ORDER_NO}'"
                            ASCDATA1.ExecuteSQL()
                        End If

                    End If
                End If

            Case Else

        End Select

        Dim REPORTFILE As String = "BMRLIST1"
        Dim RPT As String = REPORTFILE

        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and BM_PROD_ITEM = '" & ITEM_CODE & "'", BM_ISSUE_STATUS})

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")

            .CR_params.Add("RUNQTY", "0")
            .CR_params.Add("STATUS", "0")
            .CR_params.Add("COSTED_BOM", "0")

            .CR_params.Add("NOTES", "1")
            .CR_params.Add("COMPNOTES", "1")

            If make_pdf Then
                REPORT_NO = .Generate_Report(RPT, "Bill of Materials", , True, , , "PDF", FILENAME_body, False)
            Else
                REPORT_NO = .Generate_Report(RPT, "Bill of Materials", , True, , , , , False)
            End If
            .Print_Report_End(make_pdf, make_pdf)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return REPORT_NO
    End Function

    Function Print_PO(PO_ORDER_NO As String, Optional make_pdf As Boolean = False, Optional FILENAME_body As String = "") As String

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing POs")

        Dim REPORTFILE As String = "POROPRT1"
        Dim RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
        If RPT = "" Then RPT = REPORTFILE

        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it

        REPORTS(REPORTFILE).Fill_Records_RPT(New String() {" and PO_ORDER_NO = '" & PO_ORDER_NO & "'"})

        'To fill the report's dataset with data from this form's dataset:
        'With REPORTS(REPORTFILE).clsASCBASE1
        '    .EnforceConstraints(False)
        '    For Each TABLE_NAME As String In New String() {"SOTPPDI1", "SOTPPDI2", "SOTPPDI3", "SOTINVH1", "SOTSVIA1"}
        '        .dst.Tables(TABLE_NAME).Rows.Clear()
        '        Dim SQL As String = ""
        '        If TABLE_NAME = "SOTINVH1" Then
        '            SQL = "ORDR_NO = '" & ORDR_NO & "'"
        '        End If

        '        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(Sql)
        '            Dim rowr As DataRow = .dst.Tables(TABLE_NAME).NewRow
        '            If TABLE_NAME = "SOTPPDI2" Or TABLE_NAME = "SOTPPDI3" Or TABLE_NAME = "SOTINVH1" Then

        '                For I As Integer = 0 To .dst.Tables(TABLE_NAME).Columns.Count - 1
        '                    Dim COLUMN_NAME As String = .dst.Tables(TABLE_NAME).Columns(I).ColumnName
        '                    rowr.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
        '                Next
        '            Else
        '                rowr.ItemArray = row.ItemArray
        '            End If
        '            .dst.Tables(TABLE_NAME).Rows.Add(rowr)
        '        Next
        '    Next
        '    .EnforceConstraints(True)
        'End With

        Dim REPORT_NO As String = ""

        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .CR_params.Add("FORM_TYPE", "P")
            If make_pdf Then
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , "PDF", FILENAME_body, False)
            Else
                REPORT_NO = .Generate_Report(RPT, "Purchase Order", , True, , , , , False)
            End If
            .Print_Report_End(make_pdf, make_pdf)
        End With

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return REPORT_NO
    End Function

    Sub Transmit_PO(Optional retransmit As Boolean = False, Optional mark_as_transitted As Boolean = False)

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
        Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")

        Dim VEND_CODE As String = rowPOTORDR1.Item("VEND_CODE")
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

        '' 07/17/2025 - No IPSA POs to ADS at this time
        'If VEND_CODE = "IPSA" Then
        '    Dim WHSE_CODE As String = rowPOTORDR1.Item("WHSE_CODE") & ""
        '    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        '    If rowICTWHSE1 IsNot Nothing AndAlso rowICTWHSE1.Item("LP_CODE") & String.Empty = "ADS" Then
        '        MessageBox.Show($"PO No {PO_ORDER_NO} is for Vendor {VEND_CODE}. Currently {VEND_CODE} POs may not be sent to 3PL {rowICTWHSE1.Item("LP_CODE")}. PO will not be transmitted.", "Transmit PO To 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        Exit Sub
        '    End If
        'End If

        Dim PO_ORDER_NO_PDF As String = PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV)
        Dim REPORT_NO As String = Print_PO(PO_ORDER_NO, True, PO_ORDER_NO_PDF)
        ATTACHMENTs.Add(PO_ORDER_NO_PDF & ".pdf", ASCMAIN1.Folders("Temp") & PO_ORDER_NO_PDF & ".pdf")

        If optPO_TYPE.Value = "M" Then
            ' ISSUE-7044 - Balance POs/ASNs
            SendPOCancellation(PO_ORDER_NO)
            For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                Dim ITEM_CODE As String = rowPOTORDR2.Item("ITEM_CODE")
                If Not ATTACHMENTs.ContainsKey(ITEM_CODE & ".pdf") Then
                    Print_BM(ITEM_CODE, "C", True, True, ITEM_CODE)
                    ATTACHMENTs.Add(ITEM_CODE & ".pdf", ASCMAIN1.Folders("Temp") & ITEM_CODE & ".pdf")
                End If
            Next
        End If

        Dim SUBJECT As String = ""
        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
            SUBJECT = "Ahava PO " & PO_ORDER_NO & IIf(PO_HDR_CTR_REV = 0, "", "; Revision " & CStr(PO_HDR_CTR_REV))
        Else
            SUBJECT = "PO " & PO_ORDER_NO & IIf(PO_HDR_CTR_REV = 0, "", "; Revision " & CStr(PO_HDR_CTR_REV))
        End If

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")
        EMAIL_ADDRESSs.Add(rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "", rowAPTVEND1.Item("VEND_PURCH_CONTACT") & "")

        Dim SEND_NO As String = ""

        If mark_as_transitted Then
        Else
            SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                        SUBJECT, "PO", False, True, VEND_CODE, rowAPTVEND1.Item("VEND_NAME"), "Supplier")
        End If

        If mark_as_transitted Or SEND_NO <> "" Then

            ' something caused the transaction to commit when I did 2 POs, probably a temp table when the report was generated
            ' so we might as well atomically commit
            ' If Not mark_as_transitted Then BeginTrans()
            BeginTrans()

            If Not retransmit Then

                My.Computer.FileSystem.CopyFile(
                    ASCMAIN1.Folders("Temp") & PO_ORDER_NO_PDF & ".PDF",
                    ASCMAIN1.Folders("Archive") & "PO\" & PO_ORDER_NO_PDF & ".PDF", True)

                ASCMAIN1.sql = "Update POTORDR1 " & vbCrLf _
                    & " Set PO_PRINTED_IND = '1', PO_DATE_PRINTED = SYSDATE" & vbCrLf _
                    & ", PO_XMIT_IND = '1', PO_XMIT_BY = '" & ASCMAIN1.USER_ID & "', PO_XMIT_DATE = SYSDATE, PO_XMIT_XNO = '" & XNO & "'" & vbCrLf _
                    & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                ASCDATA1.ExecuteSQL()
                Dim rowPOTORDRX As DataRow = dst.Tables("POTORDRX").Rows.Find(PO_ORDER_NO)
                rowPOTORDRX.Delete()

                ASCMAIN1.sql = "Insert into POTORDRZ (PO_ORDER_NO,PO_HDR_CTR_REV,PO_ORDER_LNO" & vbCrLf _
                    & ",ITEM_CODE,PO_QTY_ORD,PO_COST,PO_DATE_REQUIRED,PO_STATUS,CARTON_PACK_QTY)" & vbCrLf _
                    & " Select POTORDR2.PO_ORDER_NO, NVL(POTORDR1.PO_HDR_CTR_REV,0), POTORDR2.PO_ORDER_LNO" & vbCrLf _
                    & ", POTORDR2.ITEM_CODE, POTORDR2.PO_QTY_ORD" & vbCrLf _
                    & ", POTORDR2.PO_COST, POTORDR2.PO_DATE_REQUIRED, POTORDR2.PO_STATUS, ICTITEM1.CARTON_PACK_QTY" & vbCrLf _
                    & " from POTORDR1,POTORDR2,ICTITEM1" & vbCrLf _
                    & " where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
                    & "   and POTORDR1.PO_ORDER_NO = '" & PO_ORDER_NO & "'" & vbCrLf _
                    & "   and ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = "Insert into POTORDRH (PO_ORDER_NO,PO_HDR_CTR_REV,PO_REVISION_NOTE,INIT_OPER,INIT_DATE,LAST_OPER,LAST_DATE)" & vbCrLf _
                    & " Select PO_ORDER_NO, NVL(PO_HDR_CTR_REV,0), DECODE(NVL(PO_HDR_CTR_REV,0),0,'Original',PO_REVISION_NOTE), LAST_OPER, LAST_DATE, LAST_OPER, LAST_DATE" & vbCrLf _
                    & " from POTORDR1" & vbCrLf _
                    & " where POTORDR1.PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                ASCDATA1.ExecuteSQL()
            End If


            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'POXMIT'," & IIf(mark_as_transitted, "'PO Marked as Transmitted'", "'PO Transmitted'") & ", '" & SEND_NO & "'" _
                & " from POTORDR1 " & vbCrLf _
                & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            ASCDATA1.ExecuteSQL()


            If rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "" = "" Then
                Dim rowTATSEND1 As DataRow = LookUp("TATSEND1", SEND_NO)
                ASCMAIN1.sql = "Update APTVEND1 Set VEND_PURCH_EMAIL = :PARM1 where VEND_CODE = :PARM2"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {rowTATSEND1.Item("SEND_TO"), VEND_CODE})
            End If

            ' not transmitting to CLA because PO is a Make - need to change that logic to B or (CLARINSUSA and M)
            ' where do we re-transmit ADS POs
            ' note to Lauren - I don't think that we transmit Make POs to ADS - should we? LBM says no

            If ASCMAIN1.CLIENT = "INT" And Not mark_as_transitted Then

                ' If rowPOTORDR1.Item("WHSE_CODE") & "" = "CLA" And (rowPOTORDR1.Item("PO_TYPE") = "B" Or rowPOTORDR1.Item("VEND_WHSE_CODE") & "" <> "CLA") Then
                'If rowPOTORDR1.Item("WHSE_CODE") & "" = "CLA" Or rowPOTORDR1.Item("WHSE_CODE") & "" = "ADS" Then

                Dim drICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowPOTORDR1.Item("WHSE_CODE") & "")

                ' Change for ADS 07/16/2025
                If drICTWHSE1 IsNot Nothing AndAlso drICTWHSE1.Item("LP_CODE") & String.Empty <> String.Empty Then
                    If MsgBox("Re-Transmit PO to 3PL also?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Dim PO_ORDER_NOs As New List(Of String)

                        ASCMAIN1.sql = "Select Count (*) from POTORDR2 where PO_ORDER_NO = :PARM1 and PO_QTY_OPN <> 0"
                        Dim PO_lines As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New Object() {PO_ORDER_NO})
                        If PO_lines = 0 Then
                            MsgBox("There are no Open Inventory Lines on PO " & PO_ORDER_NO, MsgBoxStyle.OkOnly, "PO should NOT be transmitted to 3PL")
                        Else
                            PO_ORDER_NOs.Add(PO_ORDER_NO)
                            Transmit_PO_to_3PL(PO_ORDER_NOs, False, True)
                        End If
                    End If
                End If
            End If

            ' If Not mark_as_transitted Then CommitTrans()
            CommitTrans()

        End If
    End Sub

    Sub email_PO()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now setting up email")

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
        Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")

        Dim VEND_CODE As String = rowPOTORDR1.Item("VEND_CODE")
        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

        Dim PO_ORDER_NO_PDF As String = PO_ORDER_NO & "_" & CStr(PO_HDR_CTR_REV)
        Dim REPORT_NO As String = Print_PO(PO_ORDER_NO, True, PO_ORDER_NO_PDF)
        ATTACHMENTs.Add(PO_ORDER_NO_PDF & ".pdf", ASCMAIN1.Folders("Temp") & PO_ORDER_NO_PDF & ".pdf")

        If optPO_TYPE.Value = "M" Then
            For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                Dim ITEM_CODE As String = rowPOTORDR2.Item("ITEM_CODE")
                If Not ATTACHMENTs.ContainsKey(ITEM_CODE & ".pdf") Then
                    Print_BM(ITEM_CODE, "C", False, True, ITEM_CODE)
                    ATTACHMENTs.Add(ITEM_CODE & ".pdf", ASCMAIN1.Folders("Temp") & ITEM_CODE & ".pdf")
                End If
            Next
        End If

        Dim SUBJECT As String = ""
        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
            SUBJECT = "Ahava PO " & PO_ORDER_NO & IIf(PO_HDR_CTR_REV = 0, "", "; Revision " & CStr(PO_HDR_CTR_REV))
        Else
            SUBJECT = "PO " & PO_ORDER_NO & IIf(PO_HDR_CTR_REV = 0, "", "; Revision " & CStr(PO_HDR_CTR_REV))
        End If

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        'EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski")
        EMAIL_ADDRESSs.Add(rowAPTVEND1.Item("VEND_PURCH_EMAIL") & "", rowAPTVEND1.Item("VEND_PURCH_CONTACT") & "")

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
        SUBJECT, "PO", False, True, VEND_CODE, rowAPTVEND1.Item("VEND_NAME"), "Supplier")

        If SEND_NO <> "" Then
            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                & $" Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_EML','PO emailed', '{SEND_NO}'" & vbCrLf _
                & " from POTORDR1 " & vbCrLf _
                & $" where PO_ORDER_NO = '{PO_ORDER_NO}'"
            ASCDATA1.ExecuteSQL()

            Fill_Records("TATEVNT1", PO_ORDER_NO)
            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub


    Sub Transmit_PO_to_3PL(PO_ORDER_NOs As List(Of String), Optional skip_items As Boolean = False, Optional skip_invoices As Boolean = False)
        'Transmit_Document("WHC", "WHC943O1", "SPI", HFs("PINV_NO"))
        'Transmit_Document("WHC", "WHC943O1", "MPO", "132365,159053")
        'Dim MPOs As String = ""
        'ASCMAIN1.sql = "SELECT PO_ORDER_NO FROM POTORDR1 WHERE PO_STATUS = 'O'"
        'For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PO_ORDER_NO")
        '    MPOs &= "," & row.Item(0)
        'Next
        'TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "MPO", Mid(MPOs, 2))


        ' all calls to this method now skip invoices, as per BD - once we send the invoice, no need to resend the invoice.

        Dim XMIT_NO As String = ""
        Dim rowWHT3PLX1 As DataRow = Nothing

        Dim sqlPOs As String = "'" & Join(PO_ORDER_NOs.ToArray, "','") & "'"

        ASCMAIN1.sql = $"Select DISTINCT ICTWHSE1.LP_CODE
                from POTORDR1,ICTWHSE1
                where ICTWHSE1.WHSE_CODE = POTORDR1.WHSE_CODE
                  and POTORDR1.PO_ORDER_NO in ({sqlPOs})"

        For Each rowLP_CODE As DataRow In ASCDATA1.GetDataTable("").Select("")
            Dim LP_CODE As String = rowLP_CODE.Item("LP_CODE")

            Dim LP_POs As New List(Of String)
            ASCMAIN1.sql = $"Select POTORDR1.PO_ORDER_NO, POTORDR1.VEND_CODE 
                            from POTORDR1,ICTWHSE1
                            where ICTWHSE1.WHSE_CODE = POTORDR1.WHSE_CODE
                            and POTORDR1.PO_ORDER_NO in ({sqlPOs})
                            and ICTWHSE1.LP_CODE = '{LP_CODE}'"
            For Each rowPO_ORDER_NO As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim PO_ORDER_NO As String = rowPO_ORDER_NO.Item("PO_ORDER_NO")
                Dim VEND_CODE As String = rowPO_ORDER_NO.Item("VEND_CODE") & String.Empty

                ' 07/17/2025 - No IPSA POs to ADS at this time
                If LP_CODE = "ADS" AndAlso VEND_CODE = "IPSA" Then
                    MessageBox.Show($"PO No {PO_ORDER_NO} is for Vendor {VEND_CODE}. Currently {VEND_CODE} POs may not be sent to 3PL {LP_CODE}. PO will not be transmitted.", "Transmit PO To 3PL", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    If PO_ORDER_NOs.Contains(PO_ORDER_NO) Then
                        PO_ORDER_NOs.Remove(PO_ORDER_NO)
                    End If
                Else
                    LP_POs.Add(PO_ORDER_NO)
                End If
            Next

            If LP_POs.Count = 0 Then
                Exit Sub
            End If

            Dim sqlLP_POs As String = "'" & Join(LP_POs.ToArray, "','") & "'"

            If Not skip_items Then

                XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "", LP_CODE)
                rowWHT3PLX1 = LookUp("WHT3PLX1", XMIT_NO)

                If Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
                    ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)
                            Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_888','Item Master Transmitted to {LP_CODE}', '{XMIT_NO}'
                            From POTORDR1 where PO_ORDER_NO In ({sqlLP_POs})"
                    ASCDATA1.ExecuteSQL()
                End If

                System.Threading.Thread.Sleep(1500) ' TO ENSURE UNIQUE DATETIME STAMP
            End If

            ' ISSUE-7044 - Balance POs/ASNs
            For Each PO As String In LP_POs
                SendPOCancellation(PO)
            Next

            If LP_POs.Count = 1 Then
                XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "SPO", LP_POs(0), LP_CODE)
            Else
                XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "MPO", Join(LP_POs.ToArray, ","), LP_CODE)
            End If

            rowWHT3PLX1 = LookUp("WHT3PLX1", XMIT_NO)
            If Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
                For Each PO_ORDER_NO As String In LP_POs
                    ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)
                    Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_943','PO Transmitted to {LP_CODE}', '{XMIT_NO}'
                            From POTORDR1 where PO_ORDER_NO In ({sqlLP_POs})"
                    ASCDATA1.ExecuteSQL()
                Next
            End If

            System.Threading.Thread.Sleep(1500) ' TO ENSURE UNIQUE DATETIME STAMP

            If Not skip_invoices Then
                Dim MPIs As String = ""
                'ASCMAIN1.sql = "Select PINV_NO FROM ICTPINV1 WHERE PINV_DATE > '01-AUG-2015'"
                'ASCMAIN1.sql = "SELECT PINV_NO FROM ICTPINV1 WHERE PINV_STATUS = 'O'"
                Dim POs() As String = PO_ORDER_NOs.ToArray
                ASCMAIN1.sql = "Select Distinct PINV_NO from ICTPINV2 where PO_ORDER_NO in ('" & Join(POs, "','") & "')"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "PINV_NO")
                    MPIs &= "," & row.Item(0)
                Next

                ' Change for ADS 07/16/2025, force developer to supply the LP CODE
                XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "MPI", Mid(MPIs, 2), LP_CODE)
                rowWHT3PLX1 = LookUp("WHT3PLX1", XMIT_NO)
                For Each PINV_NO As String In Split(Mid(MPIs, 2), ",")
                    If Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
                        ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)
                                Select 'POTORDR1', POTORDR1.PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_943I','PO/Invoice {PINV_NO} Transmitted to {LP_CODE}', '{XMIT_NO}'
                                from POTORDR1,ICTPINV1 where POTORDR1.PO_ORDER_NO = ICTPINV1.PO_ORER_NO and ICTPINV1.PINV_NO = '{PINV_NO}'"
                        ASCDATA1.ExecuteSQL()
                    End If
                Next
            End If

        Next
    End Sub

    Sub Seek_Approval()


        Dim PO_AMT_ORD As Decimal = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
        Dim PO_AMT_OPN As Decimal = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_OPN)", "") & "")
        Dim PO_QTY_ORD As Int64 = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_ORD)", "") & "")
        Dim PO_QTY_OPN As Int64 = 0 ' Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", "") & "")
        Dim PO_DATE_REQUIRED_MIN As Date = Nothing ' dst.Tables("POTORDR2").Compute("MIN(PO_DATE_REQUIRED)", "")
        Dim PO_DATE_REQUIRED_MAX As Date = Nothing ' dst.Tables("POTORDR2").Compute("MAX(PO_DATE_REQUIRED)", "")
        Dim PO_NINV_AMOUNT As Decimal = 0
        Dim PO_TOTAL_AMT = 0

        If dst.Tables("POTORDR2").Select("", "", DataViewRowState.CurrentRows).Length > 0 Then
            PO_AMT_ORD = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_ORD)", "") & "")
            PO_AMT_OPN = Val(dst.Tables("POTORDR2").Compute("SUM(PO_AMT_OPN)", "") & "")
            PO_QTY_ORD = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_ORD)", "") & "")
            PO_QTY_OPN = Val(dst.Tables("POTORDR2").Compute("SUM(PO_QTY_OPN)", "") & "")
            PO_DATE_REQUIRED_MIN = dst.Tables("POTORDR2").Compute("MIN(PO_DATE_REQUIRED)", "")
            PO_DATE_REQUIRED_MAX = dst.Tables("POTORDR2").Compute("MAX(PO_DATE_REQUIRED)", "")


        ElseIf dst.Tables("POTORDR5").Select("", "", DataViewRowState.CurrentRows).Length > 0 Then
            'PO_AMT_ORD = Val(dst.Tables("POTORDR5").Compute("SUM(PO_AMT_ORD)", "") & "")
            'PO_AMT_OPN = Val(dst.Tables("POTORDR5").Compute("SUM(PO_AMT_OPN)", "") & "")
            'PO_QTY_ORD = Val(dst.Tables("POTORDR5").Compute("SUM(PO_QTY_ORD)", "") & "")
            'PO_QTY_OPN = Val(dst.Tables("POTORDR5").Compute("SUM(PO_QTY_OPN)", "") & "")
            'PO_DATE_REQUIRED_MIN = dst.Tables("POTORDR2").Compute("MIN(PO_NINV_DATE_REQ)", "")
            'PO_DATE_REQUIRED_MAX = dst.Tables("POTORDR2").Compute("MAX(PO_NINV_DATE_REQ)", "")
            'PO_TOTAL_AMT = PO_AMT_ORD
        End If

        APPR_NOTES = ""

        PO_NINV_AMOUNT = Val(dst.Tables("POTORDR5").Compute("SUM(PO_NINV_AMOUNT)", "") & "")
        PO_TOTAL_AMT = PO_AMT_ORD + PO_NINV_AMOUNT

        Dim PO_PARM_APPR_LIMIT As Decimal = Val(ROWs("POTPARM1").Item("PO_PARM_APPR_LIMIT") & "")

        If PO_TOTAL_AMT > VEND_BUYER_PURCH_LIMIT And PO_TOTAL_AMT > PO_PARM_APPR_LIMIT And Not ASCMAIN1.USER_SECURITY_CODEs.Contains("OM") Then
            MsgBox("Total Purchase Order Amount is " & Format(PO_TOTAL_AMT, "$#,##0.00") & ", which is above your approval limit.", MsgBoxStyle.OkOnly, "Cannot Approve this Purchase")
        Else
            Dim LBL As String = "Total Purchase Order Amount is " & Format(PO_TOTAL_AMT, "$#,##0.00") _
                                      & vbCrLf & "Total Amount Open to Ship is " & Format(PO_AMT_OPN, "$#,##0.00") _
                                      & vbCrLf & vbCrLf & "Terms are " & Absx1.txtFor("TERM_DESC").Text _
                                      & vbCrLf & vbCrLf & "Total Units Ordered are " & Format(PO_QTY_ORD, "#,##0") _
                                      & vbCrLf & "Total Units Open are " & Format(PO_QTY_OPN, "#,##0") _
                                      & vbCrLf & vbCrLf & "Arrival Date Range is " & Format(PO_DATE_REQUIRED_MIN, "MM/dd/yy") & " thru " & Format(PO_DATE_REQUIRED_MAX, "MM/dd/yy") _
                                      & vbCrLf & vbCrLf & "Enter Notes to Record with this Approval"

            APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, "OK To Approve this Purchase?", False, 60, "Approved")
        End If
    End Sub

    Private Sub btnRequiredDate_Click(sender As Object, e As EventArgs) Handles btnRequiredDate.Click
        If Absx1.dteFor("PO_DATE_REQUIRED").Value & "" = "" Then
            Exit Sub
        Else
            Dim PO_DATE_REQUIRED As Date = CDate(Absx1.dteFor("PO_DATE_REQUIRED").Value)
            'Dim PO_DATE_REQUESTED As Date = CDate(Absx1.dteFor("PO_DATE_REQUESTED").Value)
            If MsgBox($"OK to Copy Date Required {Format(PO_DATE_REQUIRED, "MM/dd/yyyy")} to All Details?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("")
                    rowPOTORDR2.Item("PO_DATE_REQUIRED") = PO_DATE_REQUIRED
                    rowPOTORDR2.Item("PO_DATE_REQUESTED") = PO_DATE_REQUIRED
                Next
            End If
        End If
    End Sub

    Function Get_BMs(ITEM_CODE As String) As ValueList
        Dim BMs As New ValueList
        If EntryMode = "E" Or EntryMode = "N" Then
            Dim BM_ISSUE_TYPE As String = "A"
            If optPO_ORDER_TYPE.Value = "R" Then BM_ISSUE_TYPE = "R"
            ASCMAIN1.sql = "Select BM_ISSUE_NO, BM_ISSUE_COMMENT from BMTMAIN2" & vbCrLf _
                & $" where BM_PROD_ITEM = '{ITEM_CODE}' and BM_ISSUE_NO <> '00' and BM_ISSUE_TYPE = '{BM_ISSUE_TYPE}'"
            For Each row As DataRow In ASCDATA1.GetDataTable().Select("", "BM_ISSUE_NO DESC")
                Dim BM_ISSUE_NO As String = row.Item("BM_ISSUE_NO")
                Dim BM_ISSUE_COMMENT As String = row.Item("BM_ISSUE_COMMENT")
                Dim VLI As New ValueListItem(BM_ISSUE_NO, BM_ISSUE_NO)
                BMs.ValueListItems.Add(BM_ISSUE_NO, BM_ISSUE_NO)
            Next
        End If
        Return BMs
    End Function

    Private Sub grdPOTORDR2_KeyDown(sender As Object, e As KeyEventArgs) Handles grdPOTORDR2.KeyDown
        With grdPOTORDR2
            If e.KeyCode = Windows.Forms.Keys.Delete Then
                If .ActiveCell Is Nothing Then
                    Exit Sub
                Else
                    If .ActiveCell.Value & "" <> "" Then
                        .ActiveCell.Value = ""
                        .ActiveRow.Update()
                    Else
                        .ActiveCell.CancelUpdate()
                    End If
                End If
            End If
        End With

    End Sub

    Private Sub grdICTPINVD_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdICTPINVD.InitializeRow
        Dim CONTAINER_NO As String = e.Row.Cells("CONTAINER_NO").Value & ""
        Dim ETA As String = "ETA_DATE_DC"
        If CONTAINER_NO.StartsWith("Inv Date +LT") Then
            e.Row.Cells(ETA).Appearance.ForeColor = System.Drawing.Color.Red
        ElseIf CONTAINER_NO = "Qty Open" Then
            e.Row.Appearance.BackColor = System.Drawing.Color.LightGreen
        Else
            e.Row.Appearance.BackColor = System.Drawing.Color.Empty
        End If

        Dim OPO_QTY As Int32 = Val(e.Row.Cells("OPO_QTY").Value & "")
        If OPO_QTY < 0 Then
            e.Row.Cells(ETA).Appearance.BackColor = System.Drawing.Color.Red
            e.Row.Cells(ETA).Appearance.ForeColor = System.Drawing.Color.White
        End If
    End Sub


    Private Sub grdDPTPLAN1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdDPTPLAN1.DoubleClickRow
        If Not e.Row.IsDataRow Then
            Exit Sub
        End If

        Dim PLAN_NO As String = e.Row.Cells("PLAN_NO").Value
        Dim rowDPTPLAN1 As DataRow = dst.Tables("DPTPLAN1").Rows.Find(PLAN_NO)
        If grdPOTORDR2.ActiveRow.IsAddRow AndAlso grdPOTORDR2.ActiveRow.DataChanged Then
            grdPOTORDR2.ActiveRow.CancelUpdate()
        End If
        With grdPOTORDR2.DisplayLayout.Bands(0).AddNew
            .Cells("ITEM_CODE").Value = rowDPTPLAN1.Item("ITEM_CODE")
            .Cells("PO_DATE_REQUIRED").Value = rowDPTPLAN1.Item("DATE_REQUIRED")
            .Cells("PO_QTY_ORD").Value = rowDPTPLAN1.Item("QTY_PLANNED")
            .Update()
        End With
        rowDPTPLAN1.Item("DELETE") = "1"
    End Sub

    Private Sub chkEditPlans_CheckedChanged(sender As Object, e As EventArgs) Handles chkEditPlans.CheckedChanged
        With grdDPTPLAN1.DisplayLayout.Override
            If chkEditPlans.Checked Then
                .AllowAddNew = AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.False
            Else
                .AllowAddNew = AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End If
        End With


        With grdDPTPLAN1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.CellActivation = Activation.AllowEdit Then
                    If chkEditPlans.Checked Then
                        gcol.CellAppearance.BackColor = System.Drawing.Color.Yellow
                    Else
                        gcol.CellAppearance.BackColor = System.Drawing.Color.Empty
                    End If
                End If
            Next

        End With

    End Sub

    Function Get_Class_Codes(COLUMN_NAME As String, ITEM_CODEs As List(Of String)) As String()

        Dim Class_Codes As New List(Of String)
        ASCMAIN1.sql = $"Select Distinct {COLUMN_NAME} from ICTCOSTF where ITEM_CODE in ('{Join(ITEM_CODEs.ToArray, "','")}')"
        ASCMAIN1.sql &= " UNION " & Replace(ASCMAIN1.sql, "ICTCOSTF", "ICTCOSTC")
        Dim sqlSuggestions As String = $"Select Distinct {COLUMN_NAME} from ({ASCMAIN1.sql})"

        For Each rowF As DataRow In ASCDATA1.GetDataTable.Select("", COLUMN_NAME)
            Dim Class_Code As String = rowF.Item(COLUMN_NAME) & ""
            Class_Codes.Add(Class_Code)
        Next

        Return Class_Codes.ToArray
    End Function

    Private Sub SendPOCancellation(ByVal PO_ORDER_NO As String)
        ' ISSUE-7044 - Balance POs/ASNs
        Try
            Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)
            If rowPOTORDR1 Is Nothing Then
                Exit Sub
            End If

            Dim PO_HDR_CTR_REV As Int32 = Val(rowPOTORDR1.Item("PO_HDR_CTR_REV") & "")

            If PO_HDR_CTR_REV > 0 Then
                Try
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowPOTORDR1.Item("WHSE_CODE") & "")
                    If rowICTWHSE1 IsNot Nothing AndAlso rowICTWHSE1.Item("LP_CODE") & String.Empty <> String.Empty Then
                        ' Currently only for ADS
                        Select Case rowICTWHSE1.Item("LP_CODE") & String.Empty
                            Case "ADS"
                                ASCMAIN1.Progress("Sending 3PL Request to cancel previous PO transmission", "")
                                Dim transPoNo As String = rowPOTORDR1.Item("PO_ORDER_NO")
                                If PO_HDR_CTR_REV - 1 > 0 Then
                                    transPoNo = rowPOTORDR1.Item("PO_ORDER_NO") & "_" & (PO_HDR_CTR_REV - 1)
                                End If

                                Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "CANC", transPoNo, rowICTWHSE1.Item("LP_CODE"))
                                Dim rowWHT3PLX1 As DataRow = LookUp("WHT3PLX1", XMIT_NO)
                                If rowWHT3PLX1 IsNot Nothing AndAlso Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
                                    ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)
                                                        Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_CREQ','Sent Cancel Request to {rowICTWHSE1.Item("LP_CODE")} for {transPoNo}', '{XMIT_NO}'
                                                        From POTORDR1 where PO_ORDER_NO = '{rowPOTORDR1.Item("PO_ORDER_NO")}'"
                                    ASCDATA1.ExecuteSQL()
                                End If
                        End Select
                    End If
                Catch ex As Exception
                    MessageBox.Show($"PO {PO_ORDER_NO} Error: {ex.Message}", "Cancel Request Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If

        Catch ex As Exception
            MessageBox.Show($"PO {PO_ORDER_NO} Error: {ex.Message}", "Cancel Request Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            ASCMAIN1.Progress("", "")
        End Try

    End Sub

    Sub Get_PO_Cost(rowPOTORDR1 As DataRow, ITEM_CODE As String,
                    ByRef ITEM_RETAIL_PRICE As Decimal, ByRef PO_COST_CLASS As Decimal, ByRef PO_COST_LIST As Decimal)

        Dim VEND_CODE As String = rowPOTORDR1.Item("VEND_CODE")
        Dim PO_DATE_ORDERED As Date = rowPOTORDR1.Item("PO_DATE_ORDERED") & ""
        Dim PO_DATE_REQUIRED As Date = rowPOTORDR1.Item("PO_DATE_REQUIRED") & "" ' Replace this with new parameter from POTORDR2
        Dim COST_CLASS_CODE As String = rowPOTORDR1.Item("COST_CLASS_CODE") & ""
        Dim COST_LIST_CODE As String = rowPOTORDR1.Item("COST_LIST_CODE") & ""
        Dim COST_BASE_PCT_OF_MSRP As Decimal = Val(rowPOTORDR1.Item("COST_BASE_PCT_OF_MSRP") & "")

        Try




            'Dim OPS_YYYYPP As String = Format(PO_DATE_ORDERED, "yyyyMM")

            'If ITEM_RETAIL_PRICE < 0 Then
            '    Dim rowICTRETLA As DataRow = LookUp("ICTRETLA", New String() {ITEM_CODE, OPS_YYYYPP})
            '    If rowICTRETLA Is Nothing Then
            '        ITEM_RETAIL_PRICE = -1 * ITEM_RETAIL_PRICE
            '    Else
            '        ITEM_RETAIL_PRICE = Val(rowICTRETLA.Item("ITEM_RETAIL_PRICE") & "")
            '    End If
            'End If

            'If COST_LIST_CODE <> "" Then
            '    Dim row As DataRow = Nothing
            '    If OPS_YYYYPP = ASCMAIN1.CYP Then
            '        row = LookUp("ICTCLST2", New String() {COST_LIST_CODE, ITEM_CODE})
            '    Else
            '        row = LookUp("ICTCLST4", New String() {OPS_YYYYPP, COST_LIST_CODE, ITEM_CODE})
            '    End If

            '    If row IsNot Nothing Then
            '        PO_COST_LIST = System.Math.Round(Val(row.Item("ITEM_VCOST") & ""), 2)
            '    End If
            'End If

            'If COST_CLASS_CODE <> "" AndAlso ITEM_RETAIL_PRICE > 0 Then
            '    Dim rowICTCCLS1 As DataRow = dst.Tables("ICTCCLS1").Rows.Find(COST_CLASS_CODE)
            '    If rowICTCCLS1.Item("COST_BASIS") & "" = "R" Then
            '        PO_COST_CLASS = System.Math.Round(ITEM_RETAIL_PRICE * COST_BASE_PCT_OF_MSRP / 100, 2)
            '    End If

            'End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub grdPOTORDR2_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdPOTORDR2.InitializeRow
        Dim PO_COST_CLASS As Decimal = Val(e.Row.Cells("PO_COST_CLASS").Value & "")
        Dim PO_COST_LIST As Decimal = Val(e.Row.Cells("PO_COST_LIST").Value & "")
        Dim PO_COST As Decimal = Val(e.Row.Cells("PO_COST").Value & "")

        e.Row.Cells("PO_COST").Appearance.ForeColor = System.Drawing.Color.Empty
        If Absx1.txtFor("COST_LIST_CODE").Text <> "" AndAlso PO_COST_LIST <> 0 AndAlso PO_COST_LIST <> PO_COST Then
            e.Row.Cells("PO_COST").Appearance.ForeColor = System.Drawing.Color.Red
            e.Row.Cells("PO_COST").ToolTipText = "PO Cost List Cost is " & Format(PO_COST_LIST, "#,##0.0000")
        ElseIf Absx1.txtFor("COST_CLASS_CODE").Text <> "" AndAlso PO_COST_LIST = 0 And PO_COST_CLASS <> 0 AndAlso PO_COST_CLASS <> PO_COST Then
            e.Row.Cells("PO_COST").Appearance.ForeColor = System.Drawing.Color.Red
            e.Row.Cells("PO_COST").ToolTipText = "PO Cost Class Cost is " & Format(PO_COST_CLASS, "#,##0.0000")
        End If
    End Sub
End Class