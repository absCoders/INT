Imports System.Drawing
Imports System.Math
Imports Infragistics.Win.UltraWinGrid

Public Class ICFPINV1

    ' Transmit_888_ADS - REFERS TO WHC88802 WHICH DOES NOT EXIST YET
    ' Transmit_943_ADS - REFERS TO WHC94302 WHICH DOES NOT EXIST YET

    Dim rowICTPINV1 As DataRow
    Dim rowICTWHSE1 As DataRow
    Dim reversal_update As Boolean = False

    Dim PO_ORDER_NO_IDOC As String

    Dim IDOC_mode As String
    Dim IDOC_SEQ_NO As String = ""
    Dim IDOC_DATA_AMT As Decimal
    Dim IDOC_DATA_KEY As String
    Dim IDOC_DATA_DATE As Date
    Dim IDOC_DATA_REF As String
    Dim IDOC_DATA_REF_ORIG As String
    Dim IDOC_DATA_TYPE As String
    Dim IDOC_DATA_REF_INV As String
    Dim IDOC_DATA_PACK_LIST_NO As String
    Dim IDOC_import As Boolean
    Dim DELETED_INV_LINES As New List(Of String)
    Dim DEL_LINE As String = ""
    Dim DEL_ITEM_CODE As String = ""
    Dim DEL_ITEM_DESC As String = ""


    Dim rowCopyFrom As DataRow = Nothing
    Dim editableColumns() As String = New String() {"VESSEL_NAME", "CONTAINER_NO", "SHIP_DATE", "ETA_DATE", "BOL_NO"}
    Dim editColumn As String = ""

    Dim IDOC_FILENAME As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFPINVI" Then
            InquiryMode = True
        End If

        If ASCMAIN1.Running_in_VS Then
            cmdPackList.Visible = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        AUDIT.Add("ICTPINV1", "*")

        With dst
            ASCMAIN1.sql = "Select ICTPINV1.*" _
            & " from ICTPINV1 where ICTPINV1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "ICTPINVX", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "ICTPINV1", "*")

            ASCMAIN1.sql = "Select ICTPINV2.*, ICTITEM1.ITEM_DESC, POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, ICTCOSTF.ITEM_COST_VCOST FUT_VCOST" & vbCrLf _
                & " , ICTCOSTF.ITEM_COST_TRF_CLASS FUT_TRF_CLASS, ICTCOSTC.ITEM_COST_VCOST CUR_VCOST" & vbCrLf _
                & " , ICTCOSTC.ITEM_COST_TRF_CLASS CUR_TRF_CLASS, POTORDR1.TRF_CLASS_CODE PO_TRF_CLASS" & vbCrLf _
                & " from ICTPINV2,ICTITEM1,POTORDR2,ICTCOSTF, ICTCOSTC, POTORDR1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE" & vbCrLf _
                & " And POTORDR2.PO_ORDER_NO (+) = ICTPINV2.PO_ORDER_NO" & vbCrLf _
                & " And POTORDR2.PO_ORDER_LNO (+) = ICTPINV2.PO_ORDER_LNO" & vbCrLf _
                & " And ICTCOSTF.ITEM_CODE(+) = POTORDR2.ITEM_CODE" & vbCrLf _
                & " And ICTCOSTC.ITEM_CODE(+) = POTORDR2.ITEM_CODE" & vbCrLf _
                & " and POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO (+)"
            Create_TDA(.Tables.Add, "ICTPINV2", "**", 1)
            With .Tables("ICTPINV2")
                .Columns.Add("EXT_COST_VCOST", GetType(System.Decimal), "ISNULL(PINV_QTY,0) * ISNULL(PO_COST,0)")
                .Columns.Add("EXT_COST_TOTAL", GetType(System.Decimal), "ISNULL(PINV_QTY,0) * ISNULL(ITEM_COST_STD,0)")
                .Columns.Add("EXT_COST_INV", GetType(System.Decimal), "ISNULL(PINV_QTY,0) * ISNULL(PINV_COST,0)")
            End With

            ASCMAIN1.sql = "Select ICTPINV2.*, ICTITEM1.ITEM_DESC, POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_DATE_REQUIRED" & vbCrLf _
                        & ", ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV1.PACK_LIST_NO, ICTPINV1.WHSE_CODE, ICTPINV1.VOUCHER_NO" & vbCrLf _
                        & ", ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE" & vbCrLf _
                        & ", ICTPINV1.ASN_NO" & vbCrLf _
                        & " from ICTPINV1,ICTPINV2,ICTITEM1,POTORDR2 where " & vbCrLf _
                        & " ICTPINV1.PINV_NO = ICTPINV2.PINV_NO " & vbCrLf _
                        & " And ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE" & vbCrLf _
                        & " And POTORDR2.PO_ORDER_NO (+) = ICTPINV2.PO_ORDER_NO" & vbCrLf _
                        & " And POTORDR2.PO_ORDER_LNO (+) = ICTPINV2.PO_ORDER_LNO" & vbCrLf _
                        & " And ICTPINV1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "ICTPINVY", "**", 0, False, "V")
            With .Tables("ICTPINVY")
                .Columns.Add("EXT_COST_VCOST", GetType(System.Decimal), "ISNULL(PINV_QTY,0) * ISNULL(PO_COST,0)")
                .Columns.Add("EXT_COST_TOTAL", GetType(System.Decimal), "ISNULL(PINV_QTY,0) * ISNULL(ITEM_COST_STD,0)")
                .Columns.Add("EXT_COST_INV", GetType(System.Decimal), "ISNULL(PINV_QTY,0) * ISNULL(PINV_COST,0)")
            End With

            .Tables.Add("ICTPINV0")
            .Tables("ICTPINV0").Columns.Add("KEY")
            .Tables("ICTPINV0").Columns.Add("DESCRIPTION")

            ASCMAIN1.sql = "Select * from ICTREAS1"
            Create_TDA(.Tables.Add, "ICTREAS1", "**", 0, False)

            ASCMAIN1.sql = "Select * from ICTCLAS1"
            Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)

            TAC.TACMAIN1.Get_Unprocessed_IDOCs(Me)

            TAC.TACMAIN1.Get_Deleted_IDOCs(Me)

            'With .Tables.Add("TATIDOCU")
            '    .Columns.Add("FILENAME")
            '    .Columns.Add("FILESIZE", GetType(System.Int64))
            '    .Columns.Add("FILEDATE", GetType(System.DateTime))
            '    .Columns.Add("FILENAME_SHORT")

            '    .Columns.Add("INV_NUM")
            '    .Columns.Add("INV_DATE", GetType(System.DateTime))
            '    .Columns.Add("INV_AMT", GetType(System.Decimal))
            '    .Columns.Add("PO_ORDER_NO")
            '    .Columns.Add("PINV_TYPE")
            '    .Columns.Add("PINV_REF_INV")
            'End With


            With .Tables.Add("ICTPINVV")
                .Columns.Add("PINV_NO")
                .Columns.Add("INV_NUM")
                .Columns.Add("INV_DATE", GetType(System.DateTime))
                .Columns.Add("VESSEL_NAME")
                .Columns.Add("BOL_NO")
                .Columns.Add("CONTAINER_NO")
                .Columns.Add("SHIP_DATE", GetType(System.DateTime))
                .Columns.Add("ETA_DATE", GetType(System.DateTime))
                .Columns.Add("ERROR")
            End With

            ASCMAIN1.sql = "Select * from TATIDOCP"
            Create_TDA(.Tables.Add, "TATIDOCP", "**", 0)

            Create_TDA(.Tables.Add, "TATIDOC1", "*", 1, False)
            Create_TDA(.Tables.Add, "TATIDOC2", "*", 1, False)
            Create_TDA(.Tables.Add, "TATIDOC3", "*", 1, False)

            Fill_Records("TATIDOC1", "INVOIC01")
            Fill_Records("TATIDOC2", "INVOIC01")
            Fill_Records("TATIDOC3", "INVOIC01")

            Create_TDA(.Tables.Add, "POTORDR2", "*", 1, True)


            ASCMAIN1.sql = "Select * from TATEVNT1 where TABLE_NAME = 'ICTPINV1' and TABLE_KEY = :PARM1"
            Create_TDA(.Tables.Add, "TATEVNT1", "**", 0, True, "V")

            ASCMAIN1.sql = "Select * from ICTITLP2 where LP_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTITLP2", "**", 0, False, "V")
        End With

        Set_Read_Only(grpTotals, True)

        Fill_Records("ICTREAS1")
        Fill_Records("ICTCLAS1")

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)

        ' cbe.DataSource = ASCDATA1.GetDataTable("Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdICTPINV0.DataSource = dst.Tables("ICTPINV0")
        grdICTPINV2.DataSource = dst.Tables("ICTPINV2")
        grdICTPINVX.DataSource = dst.Tables("ICTPINVX")
        grdICTPINVY.DataSource = dst.Tables("ICTPINVY")
        grdTATIDOCU.DataSource = dst.Tables("TATIDOCU")
        grdTATIDOCD.DataSource = dst.Tables("TATIDOCD")
        grdTATIDOCP.DataSource = dst.Tables("TATIDOCP")
        grdTATEVNT1.DataSource = dst.Tables("TATEVNT1")

        Create_Summary(grdTATIDOCU, "FILENAME_SHORT", "Count")

        Create_Summary(grdICTPINVX, "PINV_NO", "Count")
        Create_Summary(grdICTPINVX, New String() {"TOTAL_COSTS", "PO_TOTAL_COSTS", "PINV_TOTAL_COSTS"})

        Create_Summary(grdICTPINV2, "PINV_LNO", "Count")
        Create_Summary(grdICTPINV2, New String() {"PINV_QTY", "EXT_COST_VCOST", "EXT_COST_TOTAL", "EXT_COST_INV", "PO_QTY_ORD", "PO_QTY_OPN"})

        Create_Summary(grdICTPINVY, "PINV_LNO", "Count")
        Create_Summary(grdICTPINVY, New String() {"PINV_QTY", "EXT_COST_VCOST", "EXT_COST_TOTAL", "EXT_COST_INV", "PO_QTY_ORD", "PO_QTY_OPN"})

        With grdICTPINV2.DisplayLayout.Bands(0)
            .Columns("PINV_LNO").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With

        grdICTPINVX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        With grdICTPINVX.DisplayLayout.Bands(0)
            .Columns("PINV_NO").Header.Fixed = True
            .Columns("WHSE_CODE").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If editableColumns.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

        With grdICTPINVY.DisplayLayout.Bands(0)
            '.Columns("PINV_NO").Hidden = False
            '.Columns("PINV_NO").Header.Caption = "Advice No"
            '.Columns("PINV_NO").Header.Fixed = True
            '.Columns("PINV_LNO").Header.Fixed = True
            '.Columns("ITEM_CODE").Header.Fixed = True
            '.Columns("ITEM_DESC").Header.Fixed = True

            'Dim vp As Integer = .Columns("PINV_NOTE").Header.VisiblePosition
            ''.Columns("PO_ORDER_NO").Header.Caption = "PO Order No"
            '.Columns("PO_ORDER_NO").Header.VisiblePosition = vp
            '.Columns("PO_ORDER_LNO").Header.VisiblePosition = vp + 1
            '.Columns("PO_DATE_REQUIRED").Header.Caption = "Date Req"
            '.Columns("PO_DATE_REQUIRED").Header.VisiblePosition = vp + 2
            '.Columns("INV_NUM").Header.Caption = "Inv Num"
            '.Columns("INV_NUM").Header.VisiblePosition = vp + 3
            '.Columns("WHSE_CODE").Header.Caption = "Whse Code"
            '.Columns("WHSE_CODE").Header.VisiblePosition = vp + 4
            '.Columns("VOUCHER_NO").Header.Caption = "Voucher No"
            '.Columns("VOUCHER_NO").Header.VisiblePosition = vp + 5
        End With

        grdICTPINV0.DisplayLayout.Bands(0).ColHeadersVisible = False

        Set_Read_Only(grpTotals, True)
        Set_Read_Only(grpIDOCValues, True)
        Set_Read_Only_for_ctl(Absx1.txtFor("IDOC_DATA_REF"), False)

        grpHeader.Visible = False
        SplitContainer1.Panel1Collapsed = True ' until we need more header data

        If ASCMAIN1.Running_in_VS And 1 <> 1 Then
            btnGenerateReceipt.Visible = True
        End If

        MakeTransparent(chkEnableEditToShipment)
        MakeTransparent(chkShowDelIDOC)

        MakeTransparent(lblCopyFrom)

        chkEnableEditToShipment.Visible = Not InquiryMode
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                ' Validate_Code("WHSE_CODE")

                If Not IDOC_import Then
                    EMsg &= vbCr & "New Records may be established using IDOC Import only"
                End If

                Dim DT As Date = Absx1.dteFor("PINV_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Document Date is Mandatory"
                Else
                    If Not IDOC_import Then
                        TAC.SOCMAIN1.Validate_Invoice_Date(DT, 0, 0, EMsg)
                    End If
                End If

                If Absx1.txtFor("VEND_CODE").Text = "" Then
                    EMsg &= vbCr & "You must supply a Valid Supplier"
                Else
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", Absx1.txtFor("VEND_CODE").Text)
                    If IsNothing(rowAPTVEND1) Then
                        EMsg &= vbCr & "Supplier Entered Is Not Valid"
                    Else
                        ASCMAIN1.sql = "Select Count (*) from POTORDR1 where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "' and PO_STATUS = 'O'"
                        If Val(ASCDATA1.GetDataValue) = 0 Then
                            EMsg &= vbCr & "No Open POs on file with Supplier " & Absx1.txtFor("VEND_CODE").Text
                        End If
                    End If
                End If

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You must supply a Valid Warehouse"
                Else
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    Else
                        If rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        End If
                    End If
                End If

                EMsg &= TAC.ICCMAIN1.Check_Standard_Cost_Initialization(Me, "ICTPINV2")

                If EMsg.Length = 0 Then
                    If PO_ORDER_NO_IDOC <> "" Then
                        If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO_IDOC) Then
                            PO_ORDER_NO_IDOC = ""
                            Exit Sub
                        End If
                        If Not ASCMAIN1.Logical_Lock("ICTPINV1_I", IDOC_DATA_KEY) Then
                            PO_ORDER_NO_IDOC = ""
                            Exit Sub
                        End If
                    End If
                Else
                    PO_ORDER_NO_IDOC = ""
                End If

            Case "View"
                If Absx1.txtFor("PINV_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTPINV1 = LookUp("ICTPINV1", Absx1.txtFor("PINV_NO").Text)
                    If rowICTPINV1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("PINV_NO").Text & " on File"
                    End If
                End If

            Case "Update"

                Dim TRANSMIT_ALL_ITEMS As String = "1"
                Dim LP_CODE As String = ""

                dst.Tables("ICTITLP2").Rows.Clear()

                If grdICTPINV2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    If WHSE_CODE = "" Then
                        EMsg &= vbCr & "Warehouse Code is Required"
                    Else
                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                        If rowICTWHSE1 Is Nothing Then
                            EMsg &= vbCr & $"Warehouse Code {WHSE_CODE} is Invalid"
                        Else
                            LP_CODE = rowICTWHSE1.Item("LP_CODE") & ""
                            If LP_CODE = "" Then

                                If MsgBox($"Warehouse Code {WHSE_CODE} is not a 3PL" & vbCrLf & vbCrLf & "Continue with Update?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                    EMsg &= vbCr & $"Warehouse Code {WHSE_CODE} is not a 3PL"
                                End If
                            Else
                                Dim rowWHTTPLP1 As DataRow = LookUp("WHTTPLP1", LP_CODE)
                                TRANSMIT_ALL_ITEMS = rowWHTTPLP1.Item("TRANSMIT_ALL_ITEMS") & ""
                                If TRANSMIT_ALL_ITEMS <> "1" Then
                                    Fill_Records("ICTITLP2", LP_CODE)
                                End If
                            End If
                        End If
                    End If

                    Dim ITEM_CODE_with_duplicate_PO_ORDER_LNO As String = ""
                    For Each rowICTPINV2 As DataRow In dst.Tables("ICTPINV2").Select("", "", DataViewRowState.CurrentRows)
                        Dim ITEM_CODE As String = rowICTPINV2.Item("ITEM_CODE")
                        Dim PO_ORDER_NO As String = rowICTPINV2.Item("PO_ORDER_NO") & ""
                        Dim PO_ORDER_LNO As Integer = Val(rowICTPINV2.Item("PO_ORDER_LNO") & "")
                        Dim PINV_LNO As Integer = Val(rowICTPINV2.Item("PINV_LNO") & "")

                        If rowICTPINV2.Item("COST_CATGY_CODE") & "" = "" Then
                            EMsg &= vbCr & $"Unable to determine Cost Category for {ITEM_CODE}"
                        End If
                        If rowICTPINV2.Item("PROD_CODE") & "" = "" Then
                            EMsg &= vbCr & $"Unable to determine Product Code for {ITEM_CODE}"
                        End If

                        Dim sqlw As String = $"ITEM_CODE = '{ITEM_CODE}' and PO_ORDER_NO = '{PO_ORDER_NO}' and PO_ORDER_LNO = {CStr(PO_ORDER_LNO)} and PINV_LNO <> {CStr(PINV_LNO)}"

                        If dst.Tables("ICTPINV2").Select(sqlw).Length > 0 Then
                            ITEM_CODE_with_duplicate_PO_ORDER_LNO &= "," & ITEM_CODE
                        End If

                        Dim PINV_QTY As Int64 = Val(rowICTPINV2.Item("PINV_QTY") & "")
                        If PINV_QTY <= 0 Then
                            EMsg &= vbCr & $"Bad Qty for item {ITEM_CODE}"
                        End If

                        If TRANSMIT_ALL_ITEMS <> "1" Then
                            Dim rowICTITLP2 As DataRow = dst.Tables("ICTITLP2").Rows.Find(New String() {LP_CODE, ITEM_CODE})
                            If rowICTITLP2 Is Nothing Then
                                EMsg &= vbCr & $"Item {ITEM_CODE} is not on the list for 3PL {LP_CODE}"
                            End If
                        End If
                    Next

                    If dst.Tables("ICTPINV2").Select("ISNULL(PINV_NOTE,'') <> ''").Length <> 0 Then
                        EMsg &= vbCr & "Unable to Update - see errors in Notes column"
                    End If

                    If EMsg = "" Then
                        If ITEM_CODE_with_duplicate_PO_ORDER_LNO <> "" Then
                            If MsgBox("Some Items on this Receiving Advice were found" _
                                        & vbCrLf & " on more than one Invoice line" _
                                        & vbCrLf & " in connection with the same PO Order and PO Line." _
                                        & vbCrLf & vbCrLf & "OK to Continue?",
                                        MsgBoxStyle.YesNo, "Items found received more than once against Same PO/Line") = MsgBoxResult.No Then
                                Exit Sub

                            End If


                            If MsgBox("Are you sure that you want to Continue?" _
                                       & vbCrLf & vbCrLf & "*** This might cause issues at 3PL ***" _
                                      & vbCrLf & vbCrLf & "Note: You might be able to right click and merge the line (if costs are same)",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If

                        If dst.Tables("ICTPINV2").Select("ISNULL(PO_QTY_OPN,0) = 0").Length <> 0 Then
                            ' this next check might not stop duplicate receipts of partial qtys
                            ' wait for NF to ask for this
                            ' maybe we need to permit this condition to do over-receipt, or reverse receipts?
                            ' or maybe we need to record an event
                            Dim MSG As String = "Some Items On this Receiving Advice were found In the PO With Zero Qty Open."
                            Dim MSGCODE As String = "ZERSOM"
                            If dst.Tables("ICTPINV2").Select("ISNULL(PO_QTY_OPN,0) <> 0").Length <> 0 Then
                                '    EMsg &= vbCr & "All Items on this Receiving Advice were found with Zero Qty Open."
                                MSG = "All Items On this Receiving Advice were found With Zero Qty Open."
                                MSGCODE = "ZERALL"
                            End If

                            If MsgBox(MSG & vbCrLf & vbCrLf & "OK To Continue?",
                                MsgBoxStyle.YesNo, "Items found With Zero Qty Open On PO") = MsgBoxResult.No Then
                                Exit Sub
                            End If

                            Dim PINV_NO As String = rowICTPINV1.Item("PINV_NO")
                            ASCMAIN1.Record_Event("ICTPINV1", PINV_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, MSGCODE, $"Continue w/Zero Qty On Some/All Items", PINV_NO)

                        End If
                    End If

                End If

                EMsg &= TAC.ICCMAIN1.Check_Standard_Cost_Initialization(Me, "ICTPINV2")
                ' CHECK FOR ITEM BEING RECEIVED TWICE ON SAME RECEIPT?


                If Absx1.dteFor("SHIP_DATE").Value & "" <> "" And Absx1.dteFor("ETA_DATE").Value & "" <> "" Then
                    Dim SHIP_DATE As Date = Absx1.dteFor("SHIP_DATE").Value
                    Dim ETA_DATE As Date = Absx1.dteFor("ETA_DATE").Value
                    If Format(ETA_DATE, "yyyyMMdd") < Format(SHIP_DATE, "yyyyMMdd") Then
                        EMsg &= vbCr & "ETA Date must Not be prior To Ship Date"
                    End If
                End If

                If EMsg = "" Then
                    Dim msg As String = Check_Qty("ICTPINV2", Absx1.txtFor("WHSE_CODE").Text, "PINV_QTY", 1)
                    If msg <> "" Then
                        If MsgBox(msg & vbCr & vbCr & "OK To Continue Anyway?",
                                    MsgBoxStyle.YesNo,
                                    "The following Items Do Not have Sufficent Qty Open On PO For this Transaction") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Reverse"

                'If rowICTPINV1.Item("PINV_SOURCE") & "" = "I" Then
                '    EMsg &= vbCr & "Cannot Reverse an IDOC Entry"
                'End If

                'ASCMAIN1.sql = "Select Distinct PO_ORDER_NO from ICTPINV2 where PINV_NO = '" & Absx1.txtFor("PINV_NO").Text & "'"
                'For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                '    Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                '    If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then Exit Sub
                '    Dim rowPOTORDR1 As DataRow = LookUp("POTORDR1", PO_ORDER_NO)

                '    ASCMAIN1.sql = "Select * from ICTIREC2 where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                '    Dim rows As DataRow = ASCDATA1.GetDataRow
                '    If rows IsNot Nothing Then
                '        ' in the future, we would need to know if it was this invoice that was received
                '        ' right now we are doing reversals programmatically by stepping thru code 
                '        ' budget has not been approved for a fully tested reversal UI
                '        EMsg &= vbCr & "PO " & PO_ORDER_NO & " has had Receipts Posted to it - Reversal not permitted"
                '    End If

                'Next


                If rowICTPINV1.Item("VOUCHER_NO") & "" <> "" Then
                    EMsg &= vbCr & $"This IDOC has been Invoiced (see Voucher {rowICTPINV1.Item("VOUCHER_NO")} - Reversal not permitted"
                End If
                If rowICTPINV1.Item("RECEIPT_NO") & "" <> "" Then
                    EMsg &= vbCr & $"This IDOC has been Received (see Receipt {rowICTPINV1.Item("RECEIPT_NO")} - Reversal not permitted"
                End If
                If rowICTPINV1.Item("PINV_STATUS") & "" <> "O" Then
                    EMsg &= vbCr & $"The status of this IDOC is not Open - Reversal not permitted"
                End If


                ' see email to LBM 05/24/2022 - after reverse receipt, the voucher and receipt was reversed
                ' shouldn't receipt reversal have removed data from ICTPINV2.RECEIPT_NO and ICTPINV2.RECEIPT_LNO?
                ' or maybe ignore the receipt referenced by checking to see if it was reversed? <- i selected this path

                ' Select ICTPINV2.RECEIPT_NO, ICTPINV2.RECEIPT_LNO, ICTIREC1.REVERSED_BY_RECEIPT_NO
                '  from ICTPINV2,ICTIREC1 where PINV_NO = '008458' and ICTPINV2.RECEIPT_NO is Not Null
                ' and ICTIREC1.RECEIPT_NO = ICTPINV2.RECEIPT_NO
                ' RECEIP RECEIPT_LNO            REVERS
                ' ------ ---------------------- ------
                ' 012063 1                      012128

                ' on 05/25, I lust skipped around this block - I am not sure what to do in code, but need to reverse this IDOC so that we can re-import it

                ' ASCMAIN1.sql = "Select Distinct PO_ORDER_NO from ICTPINV2 where PINV_NO = '" & Absx1.txtFor("PINV_NO").Text & "' and RECEIPT_NO is Not Null"
                ASCMAIN1.sql = "Select Distinct ICTPINV2.PO_ORDER_NO from ICTPINV2,ICTIREC1" & vbCrLf _
                    & " where ICTPINV2.PINV_NO = '" & Absx1.txtFor("PINV_NO").Text & "'" & vbCrLf _
                    & " and ICTPINV2.RECEIPT_NO is Not Null" & vbCrLf _
                    & " and ICTIREC1.RECEIPT_NO = ICTPINV2.RECEIPT_NO" & vbCrLf _
                    & " and ICTIREC1.REVERSED_BY_RECEIPT_NO is Null"

                Dim rowPOs() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                If rowPOs.Length <> 0 Then
                    EMsg &= vbCr & $"At least 1 PO in connection with this IDOC has been Received (see PO {rowPOs(0).Item("PO_ORDER_NO")} - Reversal not permitted"
                End If

                Dim IDOC_FILENAME As String = rowICTPINV1.Item("IDOC_FILENAME") & ""
                Dim sftp_folder As String = "" _
            & IIf(ASCMAIN1.Running_in_VS And 1 = 1, "C:\Users\wjz\Desktop\Interparfums", ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT")) _
            & "\IPSA\" _
            & IIf(ASCMAIN1.DBS_SERVER = "TST" Or ASCMAIN1.DBS_COMPANY = "TST", "TEST", "PROD") _
            & "\FROM_IPSA\IDOC\"


                If Not My.Computer.FileSystem.FileExists(sftp_folder & "archive\" & IDOC_FILENAME) Then
                    EMsg &= vbCr & $"Cannot find IDOC file {IDOC_FILENAME} in the archive folder - Reversal not permitted"
                End If



                If EMsg = "" Then
                    If MessageBox.Show("Are you sure you want to reverse this Entry?", "Confirm Reversal",
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                        Exit Sub
                    End If
                End If

            Case "Re-Transmit"

                If MessageBox.Show("Are you sure you want to Re-Transmit this Entry?",
                                   "Confirm Re-Transmit",
                                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If



            Case "IDOC Import"

                IDOC_DATA_REF = Absx1.txtFor("IDOC_DATA_REF").Text
                If IDOC_DATA_REF = "" Then
                    EMsg &= vbCr & "You must provide a valid PO No"
                Else
                    If IDOC_DATA_REF_ORIG <> IDOC_DATA_REF Then
                        If MessageBox.Show("You have changed the PO referred to in this IDOC." & vbCrLf & "Do you want to Continue?",
                       "Confirm Change to PO",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                            Exit Sub
                        End If
                    End If

                    ASCMAIN1.sql = "Select * from TATIDOCP where IDOC_DATA_KEY = '" & IDOC_DATA_KEY & "'"
                    Dim row As DataRow = ASCDATA1.GetDataRow

                    If row IsNot Nothing Then

                        ASCMAIN1.sql = "Select * from ICTPINV1 where PINV_REF_INV = '" & IDOC_DATA_KEY & "' and PINV_STATUS <> 'C'"
                        Dim row2 As DataRow = ASCDATA1.GetDataRow
                        If row2 IsNot Nothing Then
                            EMsg &= vbCr & "IDOC " & IDOC_DATA_KEY & " appears to have already been imported"
                        End If

                    End If
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

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)
                If IDOC_mode = "U" Then
                    Click_Command("IDOC Done")

                    tab0.SelectedTab = tab0.Tabs("IDOCs")
                    Setup_tab0()
                End If

            Case "Reverse"
                reversal_update = True
                Reverse_Record()
                reversal_update = False

                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)
                If IDOC_mode <> "" Then
                    Click_Command("IDOC Cancel")
                    tab0.SelectedTab = tab0.Tabs("IDOCs")
                End If

            Case "Re-Transmit"

                Dim PO_ORDER_NO As String = Absx1.txtFor("PO_ORDER_NO").Text
                Dim PINV_NO As String = Absx1.txtFor("PINV_NO").Text
                Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                ReTransmit(PO_ORDER_NO, PINV_NO, WHSE_CODE)

            Case "IDOC Refresh"
                Refresh_IDOCs()

            Case "IDOC View"

            Case "IDOC Import"

                ' shouldn't some of this happen in Proceed_Prereq?

                Dim rowtest As DataRow = LookUp("POTORDR1", IDOC_DATA_REF)
                If rowtest Is Nothing Then
                    MsgBox("Invalid PO Referenced in IDOC: " & IDOC_DATA_REF)
                    Exit Sub
                End If

                tab0.SelectedTab = tab0.Tabs("Pre-Invoices")
                'rowICTPINV1.Item("WHSE_CODE") = "CLA"
                'rowICTPINV1.Item("VEND_CODE") = "IPSA"
                'rowICTPINV1.Item("INV_NUM") = IDOC_DATA_KEY
                'rowICTPINV1.Item("INV_DATE") = IDOC_DATA_DATE
                Absx1.txtFor("WHSE_CODE").Text = rowtest.Item("WHSE_CODE") ' "CLA"
                Absx1.txtFor("VEND_CODE").Text = "IPSA"
                Absx1.txtFor("INV_NUM").Text = IDOC_DATA_KEY
                Absx1.dteFor("PINV_DATE").Value = Now.Date


                IDOC_import = True
                PO_ORDER_NO_IDOC = IDOC_DATA_REF

                Click_Command("New")

                If EntryMode = "N" Then
                    Fill_Records("POTORDR2", PO_ORDER_NO_IDOC)

                    rowICTPINV1.Item("INV_NUM") = Trim(IDOC_DATA_KEY)
                    rowICTPINV1.Item("INV_DATE") = IDOC_DATA_DATE
                    rowICTPINV1.Item("PO_ORDER_NO") = PO_ORDER_NO_IDOC
                    rowICTPINV1.Item("PINV_SOURCE") = "I"
                    rowICTPINV1.Item("PINV_SOURCE_DOC") = IDOC_DATA_KEY

                    rowICTPINV1.Item("PINV_TYPE") = IDOC_DATA_TYPE
                    rowICTPINV1.Item("PINV_REF_INV") = IDOC_DATA_REF_INV

                    rowICTPINV1.Item("PACK_LIST_NO") = IDOC_DATA_PACK_LIST_NO


                    Dim UM_QTY_warning As String = ""

                    ' Segment definition E2EDP01008 Released since Release 700 , Segment length: 0543
                    ' For Each row As DataRow In dst.Tables("E2EDP01008").Select("", "POSEX")
                    ' Segment definition E2EDP01010 Released since Release 740 , Segment length: 0585
                    ' Segment definition E2EDP01012 Released since Release 752 , Segment length: 0767

                    For Each row As DataRow In dst.Tables("E2EDP01012").Select("", "POSEX")
                        Dim ITEM_CODE As String = ""
                        Dim rowICTPINV2 As DataRow = dst.Tables("ICTPINV2").NewRow
                        Dim skip_line As Boolean = False
                        With rowICTPINV2
                            .Item("PINV_NO") = Absx1.txtFor("PINV_NO").Text
                            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                            .Item("PINV_LNO") = Val(row.Item("POSEX") & "")
                            Dim PINV_QTY As Int64 = Val(row.Item("MENGE") & "")
                            .Item("PINV_QTY") = PINV_QTY
                            ITEM_CODE = ""
                            ' For Each row2 As DataRow In row.GetChildRows("E2EDP01008_E2EDP19002")
                            'For Each row2 As DataRow In row.GetChildRows("E2EDP01012_E2EDP19002")
                            For Each row2 As DataRow In row.GetChildRows("E2EDP01012_E2EDP19003")
                                If row2.Item("QUALF") = "002" Then
                                    ITEM_CODE = Trim(row2.Item("IDTNR") & "")
                                    .Item("ITEM_CODE") = ITEM_CODE
                                End If
                            Next
                            Dim net_price_used As Boolean = False
                            'For Each row3 As DataRow In row.GetChildRows("E2EDP01008_E2EDP26")
                            For Each row3 As DataRow In row.GetChildRows("E2EDP01012_E2EDP26")
                                If row3.Item("QUALF") = "003" Then
                                    Dim TOTAL As Decimal = Val(row3.Item("BETRG") & "")
                                    .Item("PINV_COST") = TOTAL / PINV_QTY
                                    net_price_used = True
                                End If
                                If row3.Item("QUALF") = "010" Then
                                    If net_price_used Then
                                    Else
                                        Dim TOTAL As Decimal = Val(row3.Item("BETRG") & "")
                                        Dim DISCOUNT As Decimal = 0
                                        'For Each row3_015 As DataRow In row.GetChildRows("E2EDP01008_E2EDP26")
                                        For Each row3_015 As DataRow In row.GetChildRows("E2EDP01012_E2EDP26")
                                            If row3_015.Item("QUALF") = "015" Or row3_015.Item("QUALF") = "012" Then
                                                Dim D As String = Trim(row3_015.Item("BETRG") & "")
                                                If D.EndsWith("-") Then
                                                    D = "-" & Mid(D, 1, D.Length - 1)
                                                End If
                                                DISCOUNT = Val(D)
                                                TOTAL = TOTAL + DISCOUNT
                                                Exit For
                                            End If
                                        Next
                                        .Item("PINV_COST") = TOTAL / PINV_QTY
                                    End If
                                End If
                            Next

                            If ITEM_CODE = "" Then
                                .Item("PINV_NOTE") = "No Item Code Found"
                            Else
                                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                                If rowICTITEM1 Is Nothing Then
                                    .Item("PINV_NOTE") = "Invalid Item Code"
                                Else
                                    Dim MENEE As String = row.Item("MENEE") & ""
                                    Dim MENGE As Decimal = Val(row.Item("MENGE") & "")

                                    If (MENEE <> "KGM" And MENEE <> "PCE") Or Math.Round(MENGE, 0) <> PINV_QTY Then
                                        skip_line = True
                                        'UM_QTY_warning &= vbCr & "Item " & ITEM_CODE & "; UM:" & MENEE & ", Qty:" & CStr(MENGE)
                                    End If

                                    If Not skip_line Then


                                        Dim sqlI As String = "((ITEM_CODE = '" & ITEM_CODE & "' and ISNULL(ITEM_CODE_ALT,'')='') OR (ITEM_CODE_ALT = '" & ITEM_CODE & "'))"
                                        Dim rowPOTORDR2() As DataRow = dst.Tables("POTORDR2").Select(sqlI)
                                        If rowPOTORDR2.Length > 1 Then
                                            Dim rowPOTORDR2X() As DataRow = dst.Tables("POTORDR2").Select(sqlI & " and PO_COST = " & CStr(Val(.Item("PINV_COST") & "")))

                                            If rowPOTORDR2X.Length = 1 Then
                                                rowPOTORDR2 = dst.Tables("POTORDR2").Select(sqlI & " and PO_ORDER_LNO = " & rowPOTORDR2X(0).Item("PO_ORDER_LNO"))
                                            End If
                                        End If
                                        If rowPOTORDR2.Length > 1 Then
                                            ' ? ORDR OR OpenAccess ?
                                            Dim rowPOTORDR2X() As DataRow = dst.Tables("POTORDR2").Select(sqlI & " and PO_QTY_OPN = " & CStr(Val(.Item("PINV_QTY") & "")), "PO_DATE_REQUIRED ASC")
                                            If rowPOTORDR2X.Length > 0 Then
                                                rowPOTORDR2 = dst.Tables("POTORDR2").Select(sqlI & " and PO_ORDER_LNO = " & rowPOTORDR2X(0).Item("PO_ORDER_LNO"))
                                            End If
                                        End If

                                        If rowPOTORDR2.Length = 0 Then
                                            .Item("PINV_NOTE") = "No open PO Lines w/Item"
                                        ElseIf rowPOTORDR2.Length > 1 Then
                                            .Item("PINV_NOTE") = "Multiple Lines with Item"
                                        Else
                                            Dim r As Integer = 0

                                            'If rowPOTORDR2.Length > 1 Then
                                            '    Dim Ls As String = ""
                                            '    For Each rowL As DataRow In rowPOTORDR2
                                            '        Ls &= "," & rowL.Item("PO_ORDER_LNO")
                                            '    Next
                                            '    Dim rL As String = ""
                                            '    Do
                                            '        rL = InputBox("Select one of the following PO lines: " & Mid(Ls, 2), _
                                            '                                    "Multiple Lines with Item", Split(Ls, ",")(1))
                                            '    Loop While Val(rL) = 0 Or Not Split(Ls, ",").Contains(CStr(rL))
                                            '    r = Val(rL)
                                            'End If

                                            Dim rowICTCOSTC As DataRow = LookUp("ICTCOSTC", ITEM_CODE)
                                            Dim rowICTCOSTF As DataRow = LookUp("ICTCOSTF", ITEM_CODE) ' might not be here

                                            If rowPOTORDR2(r).Item("ITEM_CODE") <> ITEM_CODE Then
                                                Dim ITEM_CODE_ALT As String = rowPOTORDR2(r).Item("ITEM_CODE_ALT") & ""
                                                If ITEM_CODE_ALT = ITEM_CODE Then
                                                    Dim ITEM_CODE_PO As String = rowPOTORDR2(r).Item("ITEM_CODE")
                                                    rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE_PO)
                                                    rowICTCOSTC = LookUp("ICTCOSTC", ITEM_CODE_PO)
                                                    rowICTCOSTF = LookUp("ICTCOSTF", ITEM_CODE_PO)
                                                    .Item("ITEM_CODE") = ITEM_CODE_PO
                                                Else
                                                    .Item("PINV_NOTE") = "PO Item is " & rowPOTORDR2(r).Item("ITEM_CODE")
                                                End If
                                            End If

                                            .Item("ITEM_COST_STD") = rowICTITEM1.Item("ITEM_COST_STD")
                                            .Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE")
                                            .Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
                                            .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")

                                            If (rowICTCOSTC IsNot Nothing) Then
                                                .Item("CUR_VCOST") = rowICTCOSTC.Item("ITEM_COST_VCOST")
                                                .Item("CUR_TRF_CLASS") = rowICTCOSTC.Item("ITEM_COST_TRF_CLASS")
                                            End If
                                            If (rowICTCOSTF IsNot Nothing) Then
                                                .Item("FUT_VCOST") = rowICTCOSTF.Item("ITEM_COST_VCOST")
                                                .Item("FUT_TRF_CLASS") = rowICTCOSTF.Item("ITEM_COST_TRF_CLASS")
                                            End If

                                            .Item("PO_ORDER_NO") = rowPOTORDR2(r).Item("PO_ORDER_NO")
                                            .Item("PO_ORDER_LNO") = rowPOTORDR2(r).Item("PO_ORDER_LNO")
                                            .Item("PO_COST") = rowPOTORDR2(r).Item("PO_COST")
                                            .Item("PO_TRF_CLASS") = rowtest.Item("TRF_CLASS_CODE")
                                            .Item("PO_QTY_ORD") = rowPOTORDR2(r).Item("PO_QTY_ORD")
                                            .Item("PO_QTY_OPN") = rowPOTORDR2(r).Item("PO_QTY_OPN")

                                            If rowPOTORDR2(r).Item("PO_DATE_BACKORDER") & "" <> "" And rowPOTORDR2(r).Item("PO_QTY_OPN") <> Val(rowICTPINV2.Item("PINV_QTY") & "") And rowPOTORDR2(r).Item("PO_QTY_OPN") > Val(rowICTPINV2.Item("PINV_QTY") & "") Then
                                                Dim PO_QTY_BACKORDER As Integer = Val(rowPOTORDR2(r).Item("PO_QTY_BACKORDER") & "")
                                                'If PO_QTY_BACKORDER = 0 Then
                                                PO_QTY_BACKORDER = Val(rowPOTORDR2(r).Item("PO_QTY_OPN") & "") - Val(rowICTPINV2.Item("PINV_QTY") & "")
                                                'End If
                                                ' Update Current Line
                                                Dim PO_DATE_BACKORDER As String = rowPOTORDR2(r).Item("PO_DATE_BACKORDER")
                                                rowPOTORDR2(r).Item("PO_QTY_ORD") = Val(rowPOTORDR2(r).Item("PO_QTY_ORD") & "") - PO_QTY_BACKORDER
                                                rowPOTORDR2(r).Item("PO_QTY_OPN") = Val(rowPOTORDR2(r).Item("PO_QTY_OPN") & "") - PO_QTY_BACKORDER
                                                rowPOTORDR2(r).Item("PO_QTY_BACKORDER") = 0
                                                rowPOTORDR2(r).Item("PO_DATE_BACKORDER") = System.DBNull.Value

                                                'Update New POTORDR2 Split Line
                                                Dim PO_ORDER_NO_SPLIT As String = rowPOTORDR2(r).Item("PO_ORDER_NO")
                                                '  Dim PO_ORDR_LNO As Integer = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", "") & "") + 1
                                                Dim PO_ORDR_LNO As Integer = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", $"PO_ORDER_NO = '{PO_ORDER_NO_SPLIT}'")) + 1
                                                Dim rowPOTORDR2X As DataRow = Nothing
                                                rowPOTORDR2X = dst.Tables("POTORDR2").NewRow
                                                With rowPOTORDR2X
                                                    .Item("PO_ORDER_NO") = rowPOTORDR2(r).Item("PO_ORDER_NO")
                                                    .Item("PO_ORDER_LNO") = PO_ORDR_LNO
                                                    .Item("ITEM_CODE") = rowPOTORDR2(r).Item("ITEM_CODE")
                                                    .Item("ITEM_DESC") = rowPOTORDR2(r).Item("ITEM_DESC")
                                                    .Item("ITEM_UOM") = rowPOTORDR2(r).Item("ITEM_UOM")
                                                    .Item("ITEM_PCT_ALLOW_OVER") = rowPOTORDR2(r).Item("ITEM_PCT_ALLOW_OVER")
                                                    .Item("ITEM_PCT_ALLOW_UNDER") = rowPOTORDR2(r).Item("ITEM_PCT_ALLOW_UNDER")
                                                    .Item("PO_COST") = rowPOTORDR2(r).Item("PO_COST")
                                                    .Item("PO_QTY_ORD") = PO_QTY_BACKORDER
                                                    .Item("WHSE_CODE") = rowPOTORDR2(r).Item("WHSE_CODE")
                                                    .Item("PO_DATE_REQUIRED") = PO_DATE_BACKORDER
                                                    .Item("PO_STATUS") = rowPOTORDR2(r).Item("PO_STATUS")
                                                    .Item("PO_QTY_BACKORDER") = 0
                                                    .Item("PO_DATE_BACKORDER") = System.DBNull.Value

                                                    'Dim BM_ISSUE_DATE As String = rowPOTORDR2(r).Item("PO_ORDER_NO")
                                                    'Dim PO_DATE_REQUIRED_MRP As String = rowPOTORDR2(r).Item("PO_ORDER_NO")
                                                    'Dim PO_DATE_COMPSDUE As String = rowPOTORDR2(r).Item("PO_ORDER_NO")

                                                    .Item("PO_QTY_OPN") = PO_QTY_BACKORDER
                                                    .Item("PO_PRICE_VAR_REASON") = rowPOTORDR2(r).Item("PO_PRICE_VAR_REASON")
                                                    .Item("BM_ISSUE_NO") = rowPOTORDR2(r).Item("BM_ISSUE_NO")
                                                    .Item("BM_ISSUE_DATE") = rowPOTORDR2(r).Item("BM_ISSUE_DATE")
                                                    .Item("BM_ISSUE_SEL") = rowPOTORDR2(r).Item("BM_ISSUE_SEL")
                                                    .Item("PO_ITEM_NOTE") = rowPOTORDR2(r).Item("PO_ITEM_NOTE")
                                                    .Item("PO_AUTO_PRD_SUB") = rowPOTORDR2(r).Item("PO_AUTO_PRD_SUB")
                                                    .Item("PO_DATE_COMPSDUE") = rowPOTORDR2(r).Item("PO_DATE_COMPSDUE")
                                                    .Item("ITEM_CODE_ALT") = rowPOTORDR2(r).Item("ITEM_CODE_ALT")
                                                End With
                                                dst.Tables("POTORDR2").Rows.Add(rowPOTORDR2X)
                                            End If

                                        End If
                                    End If

                                End If
                            End If

                        End With

                        If Not skip_line Then
                            dst.Tables("ICTPINV2").Rows.Add(rowICTPINV2)
                        End If

                    Next

                    If UM_QTY_warning <> "" Then
                        MsgBox("There may be issues with this invoice: " & UM_QTY_warning, MsgBoxStyle.OkOnly, "Warning: Check UM & Qty")
                    End If

                    Sort_grdColumns(grdICTPINV2, "PINV_LNO")
                Else
                    IDOC_import = False
                    PO_ORDER_NO_IDOC = ""
                End If

            Case "IDOC Cancel"
                IDOC_Modes("")
                Refresh_IDOCs()

            Case "IDOC Done"
                IDOC_Modes("")
                Refresh_IDOCs()

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If ScreenMode And EntryMode <> "N" Then
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

                    .Items("Re-Transmit").Visible = (ScreenMode And InquiryMode) AndAlso rowICTWHSE1.Item("LP_CODE") & "" <> ""

                    .Items("Reverse").Visible = (ScreenMode AndAlso EntryMode = "V") And Not InquiryMode _
                        AndAlso rowICTPINV1 IsNot Nothing _
                        AndAlso rowICTPINV1.Item("REVERSED_BY_PINV_NO") Is DBNull.Value _
                        AndAlso rowICTPINV1.Item("REVERSES_PINV_NO") Is DBNull.Value

                    '   .Items("Reverse").Visible = False ' not ready for this
                End With

                .Groups("Record Filter").Visible = Not ScreenMode ' And InStr(ASCMAIN1.USER_SECURITY_CODEs, "X5") <> 0
                .Groups("Totals").Visible = False ' ScreenMode
                .Groups("Events").Visible = ScreenMode And (EntryMode <> "N")

                .Groups("Invoice Info").Visible = ScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        tab0.Visible = Not ScreenMode

        tab0.Tabs("IDOCs").Visible = Not InquiryMode

        If ScreenMode Then
            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpInvoiceInfo, InquiryMode Or IDOC_import)
            Set_Read_Only(grpInvoiceInfo, True)

            grdICTPINV2.DisplayLayout.Bands(0).Columns("RECEIPT_NO").Hidden = (EntryMode = "N")
            grdICTPINV2.DisplayLayout.Bands(0).Columns("RECEIPT_LNO").Hidden = (EntryMode = "N")
            grdICTPINV2.DisplayLayout.Bands(0).Columns("PINV_NOTE").Hidden = Not (EntryMode = "N") ' USED FOR ERRORS

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            If EntryMode = "N" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPINV2}
                    With grd.DisplayLayout.Override
                        '.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Next
            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPINV2}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next

            End If
            If EntryMode = "N" Then
                'Set_Read_Only_for_ctl(Absx1.txtFor("INV_NUM"), False)
                Set_Read_Only_for_ctl(Absx1.dteFor("PINV_DATE"), False)
            End If
            If EntryMode = "N" Then ' If EntryMode = "N" Or EntryMode = "E" Then
                Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), False)
            End If

        Else
            Clear_Record()

            chkAllowDN.Checked = False
            chkEnableEditToShipment.Checked = False
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTPINV0", "ICTPINV1", "ICTPINV2", "TATEVNT1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Refresh_Documents()

        Absx1.txtFor("WHSE_CODE").Text = ""
        Absx1.dteFor("PINV_DATE").Value = Format(Now, "MM/dd/yyyy")
        Absx1.txtFor("PINV_NO").Text = ""

        lblCopyFrom.Visible = False
        chkShowDelIDOC.Checked = False
        SplitContainer3.Panel2Collapsed = True

        DELETED_INV_LINES = New List(Of String)

        IDOC_SEQ_NO = ""
        IDOC_import = False
        PO_ORDER_NO_IDOC = ""
        Setup_tab0()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowICTPINV1 = dst.Tables("ICTPINV1").NewRow
            rowICTPINV1.Item("PINV_NO") = ASCMAIN1.Next_Control_No("ICTPINV1.PINV_NO")
            rowICTPINV1.Item("WHSE_CODE") = HFs("WHSE_CODE")
            rowICTPINV1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowICTPINV1.Item("PINV_DATE") = HFs("PINV_DATE")
            rowICTPINV1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowICTPINV1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTPINV1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTPINV1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTPINV1.Item("LAST_DATE") = DATETIME_STAMP
            rowICTPINV1.Item("REGISTER_IND") = "0"
            'rowICTPINV1.Item("JOURNAL_IND") = "0"
            rowICTPINV1.Item("PINV_STATUS") = "O"
            rowICTPINV1.Item("PINV_TYPE") = "I"
            rowICTPINV1.Item("IDOC_FILENAME") = IDOC_FILENAME
            dst.Tables("ICTPINV1").Rows.Add(rowICTPINV1)
        Else
            rowICTPINV1 = Fill_Record("ICTPINV1", Absx1.txtFor("PINV_NO").Text)
            dst.AcceptChanges()

            Fill_Records("TATEVNT1", Absx1.txtFor("PINV_NO").Text)
            Sort_grdColumns(grdTATEVNT1, "INIT_DATE".ToLower)


            With dst.Tables("ICTPINV0").Rows
                .Add(New String() {"Entered", Format(rowICTPINV1.Item("INIT_DATE"), "MM/dd/yy hh:mm tt")})
                .Add(New String() {"By", rowICTPINV1.Item("INIT_OPER")})
                .Add(New String() {"Source", rowICTPINV1.Item("INV_NUM") & ""})
                If rowICTPINV1.Item("PINV_STATUS") & "" = "I" Then
                    .Add(New String() {"Status", "Invoiced"})
                    .Add(New String() {"Voucher", rowICTPINV1.Item("VOUCHER_NO") & ""})
                    .Add(New String() {"Receipt", rowICTPINV1.Item("RECEIPT_NO") & ""})
                ElseIf rowICTPINV1.Item("PINV_STATUS") & "" = "R" Then
                    .Add(New String() {"Status", "Received"})
                    .Add(New String() {"Receipt", rowICTPINV1.Item("RECEIPT_NO") & ""})
                ElseIf rowICTPINV1.Item("PINV_STATUS") & "" = "C" Then
                    .Add(New String() {"Status", "Closed"})
                Else
                    .Add(New String() {"Status", "Open"})
                End If
                .Add(New String() {"PO No", rowICTPINV1.Item("PO_ORDER_NO") & ""})
                If rowICTPINV1.Item("REVERSED_BY_PINV_NO") & "" <> "" Then
                    Dim row As DataRow = LookUp("ICTPINV1", rowICTPINV1.Item("REVERSED_BY_PINV_NO"))
                    .Add(New String() {"Reversed", Format(rowICTPINV1.Item("LAST_DATE"), "MM/dd/yy hh:mm tt")})
                    .Add(New String() {"By", rowICTPINV1.Item("LAST_OPER")})
                    .Add(New String() {"using", rowICTPINV1.Item("REVERSED_BY_PINV_NO")})
                ElseIf rowICTPINV1.Item("REVERSES_PINV_NO") & "" <> "" Then
                    .Add(New String() {"Reverses", rowICTPINV1.Item("REVERSES_PINV_NO")})
                End If
            End With
        End If

        rowICTWHSE1 = LookUp("ICTWHSE1", rowICTPINV1.Item("WHSE_CODE"))
        Fill_Records("ICTPINV2", Absx1.txtFor("PINV_NO").Text)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        ' BE AWARE THAT THERE IS AN INVOICE REVERSAL PROCESS 

        BeginTrans()

        Dim PO_ORDER_NO As String = rowICTPINV1.Item("PO_ORDER_NO")
        Dim WHSE_CODE As String = rowICTPINV1.Item("WHSE_CODE")
        Dim PINV_NO As String = rowICTPINV1.Item("PINV_NO")
        Dim PINV_TYPE As String = rowICTPINV1.Item("PINV_TYPE")
        Dim PINV_REF_INV As String = rowICTPINV1.Item("PINV_REF_INV") & ""

        If IDOC_import Then
            IDOC_SEQ_NO = ASCMAIN1.Next_Control_No("TATIDOCP.IDOC_SEQ_NO")

            Dim rowTATIDOCP As DataRow = dst.Tables("TATIDOCP").NewRow
            With rowTATIDOCP
                .Item("IDOC_SEQ_NO") = IDOC_SEQ_NO
                .Item("IDOC_TABLE") = "INVOIC01"
                .Item("IDOC_FILENAME") = grdTATIDOCU.ActiveRow.Cells("FILENAME").Value
                .Item("IDOC_FILEDATE") = grdTATIDOCU.ActiveRow.Cells("FILEDATE").Value
                .Item("IDOC_FILENAME") = grdTATIDOCU.ActiveRow.Cells("FILENAME").Value
                .Item("IDOC_DATA_AMT") = IDOC_DATA_AMT
                .Item("IDOC_DATA_KEY") = IDOC_DATA_KEY
                .Item("IDOC_DATA_DATE") = IDOC_DATA_DATE
                .Item("IDOC_DATA_REF") = IDOC_DATA_REF
                .Item("IDOC_DATA_TYPE") = IDOC_DATA_TYPE
                .Item("IDOC_DATA_REF_INV") = IDOC_DATA_REF_INV
                .Item("IDOC_DATA_PACK_LIST_NO") = IDOC_DATA_PACK_LIST_NO

            End With
            dst.Tables("TATIDOCP").Rows.Add(rowTATIDOCP)

            Update_Record_TDA("TATIDOCP")
            Dim FILENAME As String = grdTATIDOCU.ActiveRow.Cells("FILENAME").Value
            Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
            My.Computer.FileSystem.MoveFile(FILENAME, fi.Directory.ToString & "\archive\" & fi.Name)

            ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
                & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'IDOC','IPSA Invoice Imported', '" & PINV_NO & "'" _
                & " from POTORDR1 " & vbCrLf _
                & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
            ASCDATA1.ExecuteSQL()

        End If

        rowICTPINV1.Item("PO_TOTAL_COSTS") = Val(dst.Tables("ICTPINV2").Compute("SUM(EXT_COST_VCOST)", "") & "")
        rowICTPINV1.Item("PINV_TOTAL_COSTS") = Val(dst.Tables("ICTPINV2").Compute("SUM(EXT_COST_INV)", "") & "")
        rowICTPINV1.Item("TOTAL_COSTS") = Val(dst.Tables("ICTPINV2").Compute("SUM(EXT_COST_TOTAL)", "") & "")

        Update_Record_TDA("ICTPINV1")
        Update_Record_TDA("ICTPINV2")


        If dst.Tables("POTORDR2").Select("", "", DataViewRowState.Added + DataViewRowState.ModifiedCurrent).Length > 0 Then
            Update_Record_TDA("POTORDR2")
        End If

        If DELETED_INV_LINES.Count > 0 Then
            For Each DELLINE As String In DELETED_INV_LINES
                DELLINE = DELLINE.Trim
                ASCMAIN1.Record_Event("ICTPINV1", PINV_NO, "", Now, ASCMAIN1.USER_ID, "DelInvLine", DELLINE, "")
                ' If DELLINE.Length = 0 Then Continue For
                ' eCheckResponse.ErrorMessage = Environment.NewLine & errorMsg
            Next
        End If

        'If reversal_update Then Set_Up_Reversal()
        'ICCMAIN1.Update_Receipt(Me)

        CommitTrans("Update Complete")

        ' NEED TO DO THIS AFTER THE COMMIT - BECAUSE THE WHC MODULE RUNS IN A SEPARATE THREAD WITH ITS OWN DB CONNECTION

        If IDOC_import And (IDOC_DATA_TYPE = "I" Or IDOC_DATA_TYPE = "D") Then

            Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
            Dim LP_CODE As String = rowICTWHSE1.Item("LP_CODE") & ""

            If LP_CODE = "" Then
                MsgBox($"No 3PL Code set up for Warehouse {WHSE_CODE} - Nothing Transmitted", MsgBoxStyle.OkOnly, "Verification")
            Else

                ' Send Item Master
                Transmit_888(PO_ORDER_NO, LP_CODE)

                System.Threading.Thread.Sleep(1500) ' TO ENSURE UNIQUE DATETIME STAMP

                ' Send IPSA Invoice
                Transmit_943(PO_ORDER_NO, PINV_NO, LP_CODE)

                MsgBox("PO Receiving Advice has been successfully Transmitted", MsgBoxStyle.OkOnly, "Verification")
            End If

        End If

        'TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "SPI", "002001")
        'TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "SPI", "002004")
        'TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "SPI", "002005")

    End Sub

    Sub Reverse_Record()

        BeginTrans()

        Dim PO_ORDER_NO As String = rowICTPINV1.Item("PO_ORDER_NO")
        Dim WHSE_CODE As String = rowICTPINV1.Item("WHSE_CODE")
        Dim PINV_NO As String = rowICTPINV1.Item("PINV_NO")
        Dim PINV_TYPE As String = rowICTPINV1.Item("PINV_TYPE")
        Dim PINV_REF_INV As String = rowICTPINV1.Item("PINV_REF_INV") & ""
        Dim IDOC_FILENAME As String = rowICTPINV1.Item("IDOC_FILENAME") & ""

        ' Stop
        ' ICTPINV1 record status should reflect reversed, and not remain open
        ' ICTPINV1 record status of open should be a pre-requisite
        ' problem still because po line does not want to be deleted because there has been IDOC activity posted against it

        Dim sftp_folder As String = "" _
            & IIf(ASCMAIN1.Running_in_VS And 1 = 1, "C:\Users\wjz\Desktop\Interparfums\", ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT")) _
            & "\IPSA\" _
            & IIf(ASCMAIN1.DBS_SERVER = "TST" Or ASCMAIN1.DBS_COMPANY = "TST", "TEST", "PROD") _
            & "\FROM_IPSA\IDOC\"

        'Dim FILENAME As String = "INVOIC_01487976"
        '' IN THE REAL SOLUTION WITH USER UI, WE SHOULD ADD FILENAME TO ICTPINV1, AS WELL AS IDOC_SEQ_NO
        '' TODAY, WE NEED TO SET THIS VARIABLE MANUALLY WITH 


        My.Computer.FileSystem.MoveFile(sftp_folder & "archive\" & IDOC_FILENAME, sftp_folder & IDOC_FILENAME)


        ASCMAIN1.sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
            & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'IDOC','IPSA Invoice Reversed', '" & PINV_NO & "'" _
            & " from POTORDR1 " & vbCrLf _
            & " where PO_ORDER_NO = '" & PO_ORDER_NO & "'"
        ASCDATA1.ExecuteSQL()

        rowICTPINV1.Item("PINV_STATUS") = "C"
        rowICTPINV1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTPINV1.Item("LAST_DATE") = DATETIME_STAMP

        Update_Record_TDA("ICTPINV1")
        Update_Record_TDA("ICTPINV2")


        CommitTrans("Reversal Complete")


    End Sub

    Sub Transmit_888(PO_ORDER_NO As String, LP_CODE As String)
        Dim XMIT_NO As String = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC888O1", "AC", "", LP_CODE)
        Dim rowWHT3PLX1 As DataRow = LookUp("WHT3PLX1", XMIT_NO)
        If Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
            ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)
            Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_888','Item Master Transmitted to {LP_CODE}', '{XMIT_NO}'
            From POTORDR1 Where PO_ORDER_NO = '{PO_ORDER_NO}'"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Sub Transmit_943(PO_ORDER_NO As String, PINV_NO As String, LP_CODE As String)

        Dim XMIT_NO As String
        Dim rowWHT3PLX1 As DataRow

        XMIT_NO = TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "SPI", PINV_NO, LP_CODE)
        rowWHT3PLX1 = LookUp("WHT3PLX1", XMIT_NO)
        If Val(rowWHT3PLX1.Item("XMIT_RECORDS") & "") <> 0 Then
            ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) 
                  Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_943I','PO/Invoice Transmitted to {LP_CODE}', '{XMIT_NO}'
                    From POTORDR1 where POTORDR1.PO_ORDER_NO = '{PO_ORDER_NO}'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY) 
                  Select 'ICTPINV1', PINV_NO, SYSDATE, '{ASCMAIN1.USER_ID}', 'PO_943I','PO/Invoice Transmitted to {LP_CODE}', '{XMIT_NO}'
                    From ICTPINV1 where ICTPINV1.PINV_NO = '{PINV_NO}'"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
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
                Absx1.txtFor("PINV_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            'Case "VEND_CODE"
            '    sql_where = "VEND_TYPE = 'S'"
            Case "WHSE_CODE"
                sql_where = "LP_CODE IS NOT NULL AND WHSE_STATUS = 'A'"
        End Select
    End Sub

    Public Overrides Function Events_Context() As Events_Entity

        Dim E As New Events_Entity

        E.TABLE_NAME = "ICTPINV1"
        E.TABLE_KEY_CAPTION = "PO Receiving Advice"
        Dim grd As UltraGrid = Nothing
        If Me.tab0.SelectedTab IsNot Nothing Then
            If Me.tab0.SelectedTab.Key = "Pre-Invoices" Then
                grd = grdICTPINVX
            ElseIf Me.tab0.SelectedTab.Key = "Pre-Invoice Details" Then
                grd = grdICTPINVY
            End If
        End If
        If grd IsNot Nothing AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow Then ' If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = grd.ActiveRow.Cells("PINV_NO").Text
            E.TABLE_KEY_DESC = ""
            E.TABLE_KEY_locked = True ' ScreenMode And (EntryMode = "E")
        Else
            E.enabled = False
            E.read_only = True
            E.TABLE_KEY_locked = False
            E.TABLE_KEY = ""
        End If

        Return E
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTPINVX, "SSBBBBBBB", "Show Filter", "Show GroupBox", "Close", "Select Container Info", "Change Whse", "Copy From", "Copy To", "Re-Transmit", "Apply ETA to all in Vessel")
        Load_Popup_Menu(grdICTPINVY, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTPINV2, "BBB", "Item Status Inquiry", "PO Inquiry", "Merge Line")
        Load_Popup_Menu(grdTATIDOCU, "B", "Delete")
        Load_Popup_Menu(grdTATIDOCD, "B", "Recover")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
            Case "grdICTPINVX"
                tlb_btn = tlb.Tools("Close")
                tlb_btn.SharedProps.Visible = Not ScreenMode AndAlso Not InquiryMode AndAlso optFilter.Value = "O" AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow
                tlb_btn = tlb.Tools("Select Container Info")
                tlb_btn.SharedProps.Visible = Not ScreenMode AndAlso Not InquiryMode AndAlso optFilter.Value = "O" AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow

                tlb_btn = tlb.Tools("Change Whse")
                tlb_btn.SharedProps.Visible = Not chkEnableEditToShipment.Checked And Not InquiryMode AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow

                tlb_btn = tlb.Tools("Close")
                tlb_btn.SharedProps.Visible = Not chkEnableEditToShipment.Checked And Not InquiryMode AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow
                tlb_btn = tlb.Tools("Select Container Info")
                tlb_btn.SharedProps.Visible = Not chkEnableEditToShipment.Checked And Not InquiryMode AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow
                tlb_btn = tlb.Tools("Copy From")
                tlb_btn.SharedProps.Visible = Not chkEnableEditToShipment.Checked And Not InquiryMode AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow
                tlb_btn = tlb.Tools("Copy To")
                tlb_btn.SharedProps.Visible = Not chkEnableEditToShipment.Checked And Not InquiryMode AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow
                tlb_btn = tlb.Tools("Re-Transmit")
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow
                tlb_btn = tlb.Tools("Apply ETA to all in Vessel")
                tlb_btn.SharedProps.Visible = chkEnableEditToShipment.Checked And Not InquiryMode AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow AndAlso optFilter.Value = "O"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdICTPINVX"
                    tlb_btn = tlb.Tools("Re-Transmit")
                    If Not (grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow) Then
                        tlb_btn.SharedProps.Visible = False
                    Else
                        Dim PINV_NO As String = grd.ActiveRow.Cells("PINV_NO").Value
                        Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Value
                        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
                        Dim LP_CODE As String = rowICTWHSE1.Item("LP_CODE") & ""
                        tlb_btn.SharedProps.Visible = (LP_CODE <> "")
                    End If
            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Re-Transmit"
                If MessageBox.Show("Are you sure you want to Re-Transmit this Entry?",
                   "Confirm Re-Transmit",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
                    Exit Sub
                End If

                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value
                Dim PINV_NO As String = grd.ActiveRow.Cells("PINV_NO").Value
                Dim WHSE_CODE As String = grd.ActiveRow.Cells("WHSE_CODE").Value

                ReTransmit(PO_ORDER_NO, PINV_NO, WHSE_CODE)

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "PO Inquiry"
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
                Context_Launch("View", PO_ORDER_NO, e.Tool.Key, "POFORDRI")

            Case "Delete"

                If grdTATIDOCU.Selected.Rows.Count <> 0 Then

                    If MsgBox("Are you sure that you want to Delete Invoices ( " & grdTATIDOCU.Selected.Rows.Count & " ) Selected, ",
                                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                    '    Dim NUMBEROFDELTE As Integer = grdTATIDOCU.Selected.Rows

                    For Each grow As UltraWinGrid.UltraGridRow In grdTATIDOCU.Selected.Rows
                        Dim FILENAME As String = grow.Cells("FILENAME").Value
                        Dim INV_NUM As String = grow.Cells(4).Value

                        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                        If My.Computer.FileSystem.FileExists(fi.Directory.ToString & "\Deleted\" & fi.Name) Then
                            My.Computer.FileSystem.DeleteFile(fi.Directory.ToString & "\Deleted\" & fi.Name)
                        End If
                        My.Computer.FileSystem.MoveFile(FILENAME, fi.Directory.ToString & "\Deleted\" & fi.Name)

                    Next

                    MsgBox("( " & grdTATIDOCU.Selected.Rows.Count & " ) Invoices Selected have been moved to a Deleted Archive", MsgBoxStyle.OkOnly, "Deleted Succesfully")
                    Refresh_IDOCs()

                End If


                ''Dim FILENAME As String = grdTATIDOCU.ActiveRow.Cells("FILENAME").Value
                ''Dim INV_NUM As String = grdTATIDOCU.ActiveRow.Cells(4).Value
                ''If MsgBox("Are you sure that you want to Delete Invoice No " & INV_NUM,
                ''                      MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                ''    Exit Sub
                ''End If

                ''Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                ''My.Computer.FileSystem.MoveFile(FILENAME, fi.Directory.ToString & "\Deleted\" & fi.Name)

                ''MsgBox("Invoice No " & INV_NUM & " has been moved to a Deleted Archive", MsgBoxStyle.OkOnly, "Deleted Succesfully")
                ''Refresh_IDOCs()

            Case "Recover"

                Dim FILENAME As String = grdTATIDOCD.ActiveRow.Cells("FILENAME").Value
                Dim INV_NUM As String = grdTATIDOCD.ActiveRow.Cells(4).Value
                Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                Dim FILE_RECOVERY_PATH As String = fi.Directory.ToString
                FILE_RECOVERY_PATH = Replace(FILE_RECOVERY_PATH, "\Deleted", "\")

                My.Computer.FileSystem.MoveFile(FILENAME, FILE_RECOVERY_PATH & "\" & fi.Name)

                MsgBox("Invoice No " & INV_NUM & " has been Recovered", MsgBoxStyle.OkOnly, "Verification")
                Refresh_IDOCs()

            Case "Select Container Info"
                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    'openFileDialog1.InitialDirectory = ASCMAIN1.Folders("Work")
                    openFileDialog1.Title = "Locate the workbook containing the data to Import"
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsX files (*.xlsx)|*.xlsx"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True
                    If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                        ImportContainerInfo(FILENAME)
                    End If
                End Using

            Case "Change Whse"

                If ASCMAIN1.Logical_Lock("F", "ICFIREC1", False,,, 10) Then
                    Dim PINV_NO As String = grdICTPINVX.ActiveRow.Cells("PINV_NO").Value & ""
                    Dim PO_ORDER_NO As String = grdICTPINVX.ActiveRow.Cells("PO_ORDER_NO").Value & ""
                    If ASCMAIN1.Logical_Lock("ICTPINV1", PINV_NO, False,,, 10) Then
                        If ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO, False,,, 10) Then
                            grdICTPINVX.ActiveRow.Cells("WHSE_CODE").Appearance.BackColor = Color.Turquoise
                            Try
                                ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("WHSE_CODE")
                                If ASCMAIN1.CodeSelector.SQL <> "" Then
                                    ASCMAIN1.CodeSelector.MultipleSelections = False
                                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = ""
                                    Using F As New ASFCODE1
                                        F.ShowDialog()
                                    End Using
                                    If ASCMAIN1.CodeSelector.SelectedCode <> "" Then
                                        Dim WHSE_CODE As String = ASCMAIN1.CodeSelector.SelectedCode
                                        DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                                        Dim rowICTPINV1 As DataRow = Fill_Record("ICTPINV1", PINV_NO)

                                        BeginTrans()
                                        Write_Event_Log("ICTPINV1", PINV_NO, $"Change Warehouse from {rowICTPINV1.Item("WHSE_CODE") } to {WHSE_CODE}")
                                        rowICTPINV1.Item("WHSE_CODE") = WHSE_CODE
                                        Update_Record_TDA("ICTPINV1")

                                        grdICTPINVX.ActiveRow.Cells("WHSE_CODE").Value = WHSE_CODE
                                        grdICTPINVX.ActiveRow.Update()
                                        dst.Tables("ICTPINVX").AcceptChanges()

                                        CommitTrans("Change Whse Complete")
                                    End If
                                End If
                            Catch ex As Exception

                            Finally
                                Me.Cursor = Cursors.Default
                            End Try

                            grdICTPINVX.ActiveRow.Cells("WHSE_CODE").Appearance.BackColor = Color.Empty

                        End If
                    End If

                    ASCMAIN1.MultiTask_Release(,, 10)
                End If

            Case "Close"
                Dim PINV_NO As String = grd.ActiveRow.Cells("PINV_NO").Value
                If MsgBox("OK to Close PO Receiving Advice No " & PINV_NO, vbYesNo, "Verification") = MsgBoxResult.Yes Then
                    If Not ASCMAIN1.Logical_Lock("ICTPINV1", PINV_NO) Then
                        MsgBox("Cannot Obtain Exclusive Access to PO Receiving Advice No " & PINV_NO, vbOKOnly, "Cannot Perform Action Requested")
                    Else
                        BeginTrans()
                        ASCMAIN1.sql = "Update ICTPINV1 Set PINV_STATUS = 'C'" & vbCrLf _
                            & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf _
                            & " where PINV_NO = '" & PINV_NO & "' and PINV_STATUS = 'O'"
                        Dim R As Integer = ASCDATA1.ExecuteSQL()
                        If R = 1 Then
                            Dim rowICTPINV1 As DataRow = LookUp("ICTPINV1", PINV_NO)
                            ASCMAIN1.Record_Event("ICTPINV1", PINV_NO, "", rowICTPINV1.Item("LAST_DATE"), rowICTPINV1.Item("LAST_OPER"), "CLOSE", "Close", "")
                            CommitTrans("PO Receiving Advice No " & PINV_NO & " has been Closed")
                            Refresh_Documents()
                        Else
                            MsgBox("Update Failed to Close PO Receiving Advice No " & PINV_NO, vbOKOnly, "Cannot Perform Action Requested")
                            Rollback()
                        End If
                        ASCMAIN1.MultiTask_Release()
                    End If
                End If

            Case "Merge Line"

                If grd.Selected.Rows.Count > 0 Then
                    MsgBox("Do not select lines to Merge." & vbCrLf & "Just right click on the line to Merge" & vbCrLf & " with another line with the same Item Code, PO and Cost", MsgBoxStyle.YesNo, "Cannot Merge when Selecting Lines")
                    Exit Sub
                End If
                Dim PINV_NO As String = grd.ActiveRow.Cells("PINV_NO").Value
                Dim PINV_LNO As Integer = Val(grd.ActiveRow.Cells("PINV_LNO").Value & "")
                Dim PINV_QTY As Integer = Val(grd.ActiveRow.Cells("PINV_QTY").Value & "")
                Dim PINV_COST As Decimal = Val(grd.ActiveRow.Cells("PINV_COST").Value & "")
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Value
                Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Value
                Dim PO_ORDER_LNO As Integer = Val(grd.ActiveRow.Cells("PO_ORDER_LNO").Value & "")
                Dim sql = $"ITEM_CODE = '{ITEM_CODE}' and PO_ORDER_NO = '{PO_ORDER_NO}' and PO_ORDER_LNO = {CStr(PO_ORDER_LNO)} and PINV_LNO <> {CStr(PINV_LNO)}  and PINV_COST = {CStr(PINV_COST)}"
                Dim row() As DataRow = dst.Tables("ICTPINV2").Select(sql, "PINV_QTY DESC")
                If row.Length = 0 Then
                    MsgBox("No other Line to Merge into", vbOKOnly, "Cannot Perform Action Requested")
                Else
                    If MsgBox("OK to merge line " & CStr(PINV_LNO) & " with line " & row(0).Item("PINV_LNO") & "?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        row(0).Item("PINV_QTY") = Val(row(0).Item("PINV_QTY") & "") + PINV_QTY
                        'grd.ActiveRow.Cells("PINV_QTY").Value = 0
                        'grd.ActiveRow.Update()
                        'grd.ActiveRow.Tag = "AllowDeleteIfZero"
                        grd.ActiveRow.Delete(False)
                    End If
                End If

            Case "Copy From"
                If grdICTPINVX.ActiveRow Is Nothing Then
                    Exit Sub
                ElseIf grdICTPINVX.ActiveCell Is Nothing Then
                    Exit Sub
                End If
                'If grdICTPINVX.Selected.Rows.Count <> 1 OrElse grdICTPINVX.ActiveRow IsNot Nothing And Not grdICTPINVX.ActiveRow.Equals(grdICTPINVX.Selected.Rows(0)) Then
                '    grdICTPINVX.Selected.Rows.Clear()

                '    If grdICTPINVX.ActiveRow IsNot Nothing Then
                '        grdICTPINVX.ActiveRow.Selected = True
                '    End If
                'End If

                'If grdICTPINVX.Selected.Rows.Count = 1 Then
                Dim COLUMN_NAME As String = grdICTPINVX.ActiveCell.Column.Key
                If editableColumns.Contains(COLUMN_NAME) Then
                    editColumn = COLUMN_NAME
                    Dim PINV_NO As String = grdICTPINVX.ActiveRow.Cells("PINV_NO").Value
                    Dim rowICTPINVX As DataRow = dst.Tables("ICTPINVX").Rows.Find(PINV_NO)
                    rowCopyFrom = dst.Tables("ICTPINVX").NewRow
                    rowCopyFrom.ItemArray = rowICTPINVX.ItemArray
                    lblCopyFrom.Text = "Copy " & grdICTPINVX.ActiveCell.Column.Header.Caption & " " & rowCopyFrom.Item(COLUMN_NAME)
                    lblCopyFrom.Visible = True
                End If
                'End If

            Case "Copy To"

                If rowCopyFrom Is Nothing Then Exit Sub

                If grdICTPINVX.Selected.Rows.Count = 0 Then
                    If grdICTPINVX.ActiveRow IsNot Nothing Then
                        grdICTPINVX.ActiveRow.Selected = True
                    End If
                End If

                If grdICTPINVX.Selected.Rows.Count <> 0 Then
                    Dim PINV_NOs As New List(Of String)

                    For Each grow As UltraWinGrid.UltraGridRow In grdICTPINVX.Selected.Rows

                        Dim PINV_NO As String = grow.Cells("PINV_NO").Value
                        If Not ASCMAIN1.Logical_Lock("ICTPINV1", PINV_NO) Then
                            PINV_NOs.Clear()
                            Exit For
                        End If
                        PINV_NOs.Add(PINV_NO)
                    Next

                    If PINV_NOs.Count > 0 Then
                        dst.Tables("ICTPINV1").Rows.Clear()

                        BeginTrans()

                        For Each PINV_NO As String In PINV_NOs
                            Dim rowICTPINVX As DataRow = dst.Tables("ICTPINVX").Rows.Find(PINV_NO)
                            Dim rowICTPINV1 As DataRow = Fill_Record("ICTPINV1", PINV_NO, False, False)
                            For Each C As String In editableColumns
                                If C = editColumn Then
                                    rowICTPINVX.Item(C) = rowCopyFrom(C)
                                    rowICTPINV1.Item(C) = rowCopyFrom(C)
                                End If
                            Next

                            'ASCMAIN1.sql = "Update ICTPINV1 Set VESSEL_NAME = :PARM1, CONTAINER_NO = :PARM2, SHIP_DATE = :PARM3, ETA_DATE = :PARM4, BOL_NO = :PARM5" & vbCrLf _
                            '    & " where PINV_NO = :PARM6"
                            'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVDDVV", New Object() _
                            '                    {rowCopyFrom.Item("VESSEL_NAME"),
                            '                    rowCopyFrom.Item("CONTAINER_NO"),
                            '                    rowCopyFrom.Item("SHIP_DATE"),
                            '                    rowCopyFrom.Item("ETA_DATE"),
                            '                    rowCopyFrom.Item("BOL_NO"),
                            '                    PINV_NO})
                        Next

                        Update_Record_TDA("ICTPINV1")
                        CommitTrans($"Copy to {PINV_NOs.Count} Records Successful")

                        lblCopyFrom.Visible = False
                        editColumn = ""
                        rowCopyFrom = Nothing

                        dst.Tables("ICTPINV1").Rows.Clear()

                    End If

                    ASCMAIN1.MultiTask_Release()
                End If

            Case "Apply ETA to all in Vessel"
                If chkEnableEditToShipment.Checked Then
                    If grdICTPINVX.Selected.Rows.Count = 0 Then
                        If grdICTPINVX.ActiveRow IsNot Nothing Then
                            grdICTPINVX.ActiveRow.Selected = True
                        End If
                    End If

                    If grdICTPINVX.Selected.Rows.Count = 1 Then
                        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

                        Dim sourcePINV As String = grdICTPINVX.ActiveRow.Cells("PINV_NO")?.Value
                        Dim sourcePINVRow As DataRow = dst.Tables("ICTPINVX").Rows.Find(sourcePINV)
                        Dim vesselName As String = sourcePINVRow("VESSEL_NAME") & ""
                        If vesselName = "" Then
                            Exit Sub
                        End If

                        Dim newETADate As Date? = sourcePINVRow("ETA_DATE")
                        Dim invsToChange As DataRow() = dst.Tables("ICTPINVX").Select($"VESSEL_NAME='{vesselName}'")
                        Dim PINV_NOs As New List(Of String)
                        For Each invToChange As DataRow In invsToChange
                            Dim PINV_NO As String = invToChange("PINV_NO")
                            If sourcePINV <> PINV_NO Then
                                If Not ASCMAIN1.Logical_Lock("ICTPINV1", PINV_NO) Then
                                    PINV_NOs.Clear()
                                    Exit For
                                End If

                                PINV_NOs.Add(PINV_NO)
                            End If
                        Next

                        If PINV_NOs.Count > 0 Then
                            dst.Tables("ICTPINV1").Rows.Clear()

                            BeginTrans()

                            For Each PINV_NO As String In PINV_NOs
                                Dim rowICTPINVX As DataRow = dst.Tables("ICTPINVX").Rows.Find(PINV_NO)
                                Dim rowICTPINV1 As DataRow = Fill_Record("ICTPINV1", PINV_NO, False, False)
                                rowICTPINVX("ETA_DATE") = newETADate
                                rowICTPINV1("ETA_DATE") = newETADate
                                rowICTPINV1("INIT_DATE") = DATETIME_STAMP
                                rowICTPINV1("LAST_OPER") = ASCMAIN1.USER_ID
                            Next


                            Update_Record_TDA("ICTPINV1")
                            CommitTrans($"ETA Date applied to {PINV_NOs.Count} Additional Records Successfully")


                            dst.Tables("ICTPINV1").Rows.Clear()

                        End If

                        ASCMAIN1.MultiTask_Release()
                    End If
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                'If e.KeyCode = Windows.Forms.Keys.Enter Then
                '    If Not InquiryMode Then
                '        Click_Command("New", e)
                '    End If
                'End If
            Case "PINV_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "PINV_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTPINV2"

    Private Sub grdICTPINV2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPINV2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdICTPINV2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    Dim ITEM_CODE As String = e.Cell.Value
                    Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("PROD_CODE").Value = cdr.Item("PROD_CODE")

                    'Dim rowICTSTAT2 = Fill_Record("ICTSTAT2", New String() {ITEM_CODE, WHSE_CODE}, True)
                    Dim COST_CATGY_CODE As String = cdr.Item("COST_CATGY_CODE") & ""
                    Dim PROD_CODE As String = cdr.Item("PROD_CODE") & ""
                    Dim ITEM_COST_STD As Decimal = Val(cdr.Item("ITEM_COST_STD") & "")
                    e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE
                    e.Cell.Row.Cells("PROD_CODE").Value = PROD_CODE
                    e.Cell.Row.Cells("ITEM_COST_STD").Value = ITEM_COST_STD
                Else
                    grdICTPINV2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "PINV_QTY"

        End Select
    End Sub

    Private Sub grdICTPINV2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPINV2.AfterExitEditMode

    End Sub

    Private Sub grdICTPINV2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPINV2.AfterRowsDeleted
        If DEL_LINE <> "" Then
            DELETED_INV_LINES.Add($"Line No {DEL_LINE} Deleted Item {DEL_ITEM_CODE} - {DEL_ITEM_DESC}")
        End If

        DisplayTotals()
    End Sub

    Private Sub grdICTPINV2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTPINV2.AfterRowUpdate
        DisplayTotals()
    End Sub


    Private Sub grdICTPINV2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTPINV2.BeforeExitEditMode
        If grdICTPINV2.ActiveCell Is Nothing Then Exit Sub
        With grdICTPINV2.ActiveCell
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

    Private Sub grdICTPINV2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTPINV2.BeforeRowUpdate
        With grdICTPINV2
            If e.Row.Cells("ITEM_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            'If Val(e.Row.Cells("QTY_REC").Text) = 0 Then
            '    MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("QTY_REC").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
            '    e.Cancel = True
            'End If
            If Val(e.Row.Cells("PINV_QTY").Text) = 0 Then
                MsgBox("Invalid Value entered for Qty (" & e.Row.Cells("PINV_QTY").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If


            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PINV_NO").Text = "" Then
                    .ActiveRow.Cells("PINV_NO").Value = Absx1.CtlFor("PINV_NO").Text
                    .ActiveRow.Cells("PINV_LNO").Value = Val(dst.Tables("ICTPINV2").Compute("Max(PINV_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdICTPINV2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPINV2.ClickCellButton

        If grdICTPINV2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
            Case "LOCATION_CODE"
                sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdICTPINV2, sql_where, False)

    End Sub

    Private Sub grdICTPINV2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTPINV2.Error
        grdICTPINV2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        'Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTPINV2").Compute("SUM(LINE_COSTS)", "") & "")
        'Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

    Private Sub grdICTPINVX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPINVX.DoubleClickRow
        If e.Row.IsDataRow Then
            If chkEnableEditToShipment.Checked Then
                ' DO NOT CALL UP RECORD
            Else
                Absx1.txtFor("PINV_NO").Text = e.Row.Cells("PINV_NO").Text
                Click_Command("View")
            End If
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor

        If optFilter.Value = "E" Then
            Dim YP As String = cbeYP.Value
            Fill_Records("ICTPINVX", YP)
            Fill_Records("ICTPINVY", YP)
        Else
            Dim SQLW As String = ""
            Select Case optFilter.Value
                Case "O"
                    SQLW = " where ICTPINV1.PINV_STATUS = 'O'"
                Case "N"
                    SQLW = " where ICTPINV1.RECEIPT_NO IS NULL AND ICTPINV1.PINV_STATUS <> 'C'"
            End Select

            ASCMAIN1.sql = "Select ICTPINV1.* from ICTPINV1 " & SQLW
            Fill_Records("ICTPINVX", "", True, ASCMAIN1.sql)

            ASCMAIN1.sql = "Select ICTPINV2.*, ICTITEM1.ITEM_DESC, POTORDR2.PO_QTY_ORD, POTORDR2.PO_QTY_OPN, POTORDR2.PO_DATE_REQUIRED" & vbCrLf _
                        & ", ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV1.PACK_LIST_NO, ICTPINV1.WHSE_CODE, ICTPINV1.VOUCHER_NO" & vbCrLf _
                        & ", ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO, ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE" & vbCrLf _
                        & ", ICTPINV1.ASN_NO" & vbCrLf _
                        & " from ICTPINV1,ICTPINV2,ICTITEM1,POTORDR2  " & vbCrLf _
                        & SQLW & IIf(SQLW <> "", " and ", " where ") & vbCrLf _
                        & " ICTPINV1.PINV_NO = ICTPINV2.PINV_NO " & vbCrLf _
                        & " and ICTITEM1.ITEM_CODE = ICTPINV2.ITEM_CODE" & vbCrLf _
                        & " and POTORDR2.PO_ORDER_NO (+) = ICTPINV2.PO_ORDER_NO" & vbCrLf _
                        & " and POTORDR2.PO_ORDER_LNO (+) = ICTPINV2.PO_ORDER_LNO"
            Fill_Records("ICTPINVY", "", True, ASCMAIN1.sql)

        End If

        Sort_grdColumns(grdICTPINVX, "PINV_NO".ToLower)
        grdICTPINVX.Text = "Entered in " & cbeYP.Text

        Sort_grdColumns(grdICTPINVY, "PINV_NO".ToLower)
        grdICTPINVY.Text = "Entered in " & cbeYP.Text
    End Sub

    Function Check_Qty(ByVal TABLE_NAME As String,
                       ByVal WHSE_CODE As String,
                       ByVal QTY_FIELD As String,
                       ByVal S As Integer) As String

        Dim msg As String = ""

        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
            Dim PO_ORDER_LNO As Integer = Val(row.Item("PO_ORDER_LNO") & "")
            Dim QTY As Integer = row.Item(QTY_FIELD)
            ASCMAIN1.sql = "Select * from POTORDR2 where ITEM_CODE = '" & ITEM_CODE & "' and PO_ORDER_NO = '" & PO_ORDER_NO & "' and PO_ORDER_LNO = " & CStr(PO_ORDER_LNO)
            Dim rowPOTORDR2 As DataRow = ASCDATA1.GetDataRow
            Dim PO_QTY_OPN As Integer = 0
            If rowPOTORDR2 IsNot Nothing Then
                PO_QTY_OPN = Val(rowPOTORDR2.Item("PO_QTY_OPN") & "")
            End If
            If PO_QTY_OPN + S * QTY < 0 Then
                msg &= vbCr & Format("Item " & ITEM_CODE & " has only " & CStr(PO_QTY_OPN) & " Open On Order")
            End If
        Next

        Return msg
    End Function

    Sub Set_Up_Reversal()

        ' this procedure was referenced in a remark where we were originally using Update_Record with a boolean
        ' I don't think we use this any longer - now that we have a dedicated Reverse_Record procedure

        Dim REVERSED_BY_PINV_NO As String = ASCMAIN1.Next_Control_No("ICTPINV1.PINV_NO")

        Dim rowICTPINV1_orig As DataRow = dst.Tables("ICTPINV1").NewRow
        rowICTPINV1_orig.ItemArray = rowICTPINV1.ItemArray

        rowICTPINV1 = dst.Tables("ICTPINV1").Rows(0)
        rowICTPINV1.Item("REVERSED_BY_PINV_NO") = REVERSED_BY_PINV_NO
        rowICTPINV1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        rowICTPINV1.Item("LAST_DATE") = DATETIME_STAMP
        Update_Record_TDA("ICTPINV1")

        rowICTPINV1.ItemArray = rowICTPINV1_orig.ItemArray
        rowICTPINV1.AcceptChanges()
        rowICTPINV1.SetAdded()

        With rowICTPINV1
            .Item("REVERSES_PINV_NO") = .Item("PINV_NO")
            .Item("PINV_NO") = REVERSED_BY_PINV_NO
            .Item("OPS_YYYYPP") = ASCMAIN1.CYP
            .Item("PINV_DATE") = DATETIME_STAMP.Date
            '  .Item("TOTAL_COSTS") *= -1

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
            .Item("REGISTER_IND") = "0"
            .Item("REGISTER_XNO") = DBNull.Value
        End With

        For Each row As DataRow In dst.Tables("ICTPINV2").Rows
            row.Item("PINV_NO") = REVERSED_BY_PINV_NO
            If row.Item("QTY_REC") IsNot DBNull.Value Then
                row.Item("QTY_REC") *= -1
            End If

            row.Item("OPS_YYYYPP") = ASCMAIN1.CYP

            row.AcceptChanges()
            row.SetAdded()
        Next

    End Sub

    Sub Receive_PO_Line(row As DataRow, Optional QTY_REC As Int16 = 0)
        grdICTPINV2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
        Dim ITEM_CODE As String = row.Item("ITEM_CODE")
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        grdICTPINV2.DisplayLayout.Bands(0).AddNew()
        With grdICTPINV2.ActiveRow
            .Cells("ITEM_CODE").Value = row.Item("ITEM_CODE")

            If QTY_REC <> 0 Then
                .Cells("QTY_REC").Value = QTY_REC
            Else
                .Cells("QTY_REC").Value = row.Item("PO_QTY_OPN")
            End If
            .Cells("PO_ORDER_NO").Value = row.Item("PO_ORDER_NO")
            .Cells("PO_ORDER_LNO").Value = row.Item("PO_ORDER_LNO")
            .Cells("PO_QTY_ORD").Value = row.Item("PO_QTY_ORD")
            .Cells("PO_QTY_OPN").Value = row.Item("PO_QTY_OPN")
            .Cells("ITEM_COST_STD").Value = rowICTITEM1.Item("ITEM_COST_STD")
            .Cells("PO_COST").Value = row.Item("PO_COST")
            .Cells("ITEM_UOM").Value = row.Item("ITEM_UOM")
            Dim rowPOTORDR1 As DataRow = dst.Tables("POTORDR1").Rows.Find(New Object() {row.Item("PO_ORDER_NO")})
            If rowPOTORDR1 Is Nothing Then rowPOTORDR1 = Fill_Record("POTORDR1", New Object() {row.Item("PO_ORDER_NO")}, False, False)
            Dim rowPOTORDR2 As DataRow = dst.Tables("POTORDR2").Rows.Find(New Object() {row.Item("PO_ORDER_NO"), row.Item("PO_ORDER_LNO")})
            If rowPOTORDR2 Is Nothing Then rowPOTORDR2 = Fill_Record("POTORDR2", New Object() {row.Item("PO_ORDER_NO"), Val(row.Item("PO_ORDER_LNO"))}, False, False)

            If rowPOTORDR2.Item("BM_ISSUE_SEL") & "" = "1" Or rowPOTORDR2.Item("BM_ISSUE_NO") & "" <> "" Then
                .Cells("VEND_WHSE_CODE").Value = rowPOTORDR1.Item("VEND_WHSE_CODE")
                .Cells("BM_ISSUE_SEL").Value = rowPOTORDR2.Item("BM_ISSUE_SEL")
                .Cells("BM_ISSUE_NO").Value = rowPOTORDR2.Item("BM_ISSUE_NO")
            End If


            .Update()
        End With
        grdICTPINV2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No

    End Sub

    Private Sub grdICTPINVX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTPINVX.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("REVERSED_BY_PINV_NO").Value & "" <> "" Then
                e.Row.Cells("REVERSED_BY_PINV_NO").Appearance.ForeColor = Color.Red
                e.Row.Cells("PINV_NO").Appearance.ForeColor = Color.Red
            ElseIf e.Row.Cells("REVERSES_PINV_NO").Value & "" <> "" Then
                e.Row.Cells("REVERSES_PINV_NO").Appearance.ForeColor = Color.Red
                e.Row.Appearance.BackColor = Color.Yellow
            End If
        End If
    End Sub

    Private Sub optFilter_ValueChanged(sender As Object, e As EventArgs) Handles optFilter.ValueChanged
        cbeYP.Visible = (optFilter.Value = "E")

        If Me.SELECTION_NO = 0 Then Exit Sub

        Me.Refresh_Documents()
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        Setup_tab0()
        If tab0.SelectedTab.Key = "IDOCs" Then
            If dst.Tables("TATIDOCU").Rows.Count = 0 And dst.Tables("TATIDOCP").Rows.Count = 0 Then
                IDOC_Modes("")
                Refresh_IDOCs()
            End If
        End If

    End Sub

    Sub Setup_tab0()
        UltraExplorerBar1.Groups("Screen Control").Visible = (tab0.SelectedTab.Key = "Pre-Invoices" Or tab0.SelectedTab.Key = "Pre-Invoice Details")
        UltraExplorerBar1.Groups("Record Filter").Visible = (tab0.SelectedTab.Key = "Pre-Invoices" Or tab0.SelectedTab.Key = "Pre-Invoice Details")
        UltraExplorerBar1.Groups("IDOCs").Visible = (tab0.SelectedTab.Key = "IDOCs")
        UltraExplorerBar1.Groups("IDOC Values").Visible = (tab0.SelectedTab.Key = "IDOCs") And Not (IDOC_mode = "")
        UltraExplorerBar1.Groups("Import DR Note").Visible = (tab0.SelectedTab.Key = "IDOCs") And (IDOC_mode = "")

        spl.Panel1Collapsed = (tab0.SelectedTab.Key = "IDOCs")
    End Sub

    Sub Refresh_IDOCs()

        TAC.TACMAIN1.Get_Unprocessed_IDOCs(Me)

        Fill_Records("TATIDOCP")
        Sort_grdColumns(grdTATIDOCP, "IDOC_SEQ_NO".ToLower)

        If chkShowDelIDOC.Checked Then
            TAC.TACMAIN1.Get_Deleted_IDOCs(Me)
        End If


        IDOC_Modes("")
    End Sub

    Sub Process_IDOC(FILENAME As String, PINV_TYPE As String)

        EnforceConstraints(False)

        For Each row As DataRow In dst.Tables("TATIDOC2").Select("")
            Dim IDOC_TABLE As String = row.Item("IDOC_TABLE")
            Dim IDOC_ID As String = row.Item("IDOC_ID")
            Dim IDOC_SEGMENT_DESC As String = row.Item("IDOC_SEGMENT_DESC")
            Dim IDOC_SEGMENT As String = row.Item("IDOC_SEGMENT")

            If dst.Tables.Contains(IDOC_SEGMENT) Then
                dst.Tables(IDOC_SEGMENT).Rows.Clear()
            Else

                With dst.Tables.Add(IDOC_SEGMENT)

                    If IDOC_SEGMENT <> "Control" And IDOC_SEGMENT <> "Data" And IDOC_SEGMENT <> "Status" Then
                        For Each row2 As DataRow In dst.Tables("TATIDOC3") _
                            .Select("IDOC_TABLE = '" & IDOC_TABLE & "' and IDOC_SEGMENT = '" & "Data" & "'", "IDOC_DATUM_NO")
                            '.Select("IDOC_TABLE = '" & IDOC_TABLE & "' and IDOC_ID = '" & "RECORD" & "' and IDOC_SEGMENT = '" & "Data" & "'", "IDOC_DATUM_NO")
                            Dim IDOC_DATUM_NAME As String = row2.Item("IDOC_DATUM_NAME")
                            If IDOC_DATUM_NAME = "SDATA" Then Exit For

                            Dim IDOC_DATUM_DESC As String = row2.Item("IDOC_DATUM_DESC")
                            Dim IDOC_DATUM_LENGTH As Integer = Val(row2.Item("IDOC_DATUM_LENGTH") & "")
                            .Columns.Add(IDOC_DATUM_NAME)
                            .Columns(IDOC_DATUM_NAME).MaxLength = IDOC_DATUM_LENGTH
                            If IDOC_DATUM_DESC = "IDOC" Then
                                .Columns(IDOC_DATUM_NAME).Caption = IDOC_DATUM_NAME
                            Else
                                .Columns(IDOC_DATUM_NAME).Caption = IDOC_DATUM_DESC
                            End If
                        Next
                    End If


                    For Each row2 As DataRow In dst.Tables("TATIDOC3") _
                        .Select("IDOC_TABLE = '" & IDOC_TABLE & "' and IDOC_SEGMENT = '" & IDOC_SEGMENT & "'", "IDOC_DATUM_NO")
                        '.Select("IDOC_TABLE = '" & IDOC_TABLE & "' and IDOC_ID = '" & IDOC_ID & "' and IDOC_SEGMENT = '" & IDOC_SEGMENT & "'", "IDOC_DATUM_NO")
                        Dim IDOC_DATUM_NAME As String = row2.Item("IDOC_DATUM_NAME")
                        Dim IDOC_DATUM_DESC As String = row2.Item("IDOC_DATUM_DESC")
                        Dim IDOC_DATUM_LENGTH As Integer = Val(row2.Item("IDOC_DATUM_LENGTH") & "")
                        .Columns.Add(IDOC_DATUM_NAME)
                        .Columns(IDOC_DATUM_NAME).MaxLength = IDOC_DATUM_LENGTH
                        If IDOC_DATUM_DESC = "IDOC" Then
                            .Columns(IDOC_DATUM_NAME).Caption = IDOC_DATUM_NAME
                        Else
                            .Columns(IDOC_DATUM_NAME).Caption = IDOC_DATUM_DESC
                        End If
                    Next
                End With
                Dim t As UltraWinTabControl.UltraTab = tabIDOC.Tabs.Add
                t.Text = IDOC_SEGMENT
                Dim grd As New UltraWinGrid.UltraGrid
                grd.Text = IDOC_SEGMENT & ":" & IDOC_SEGMENT_DESC
                grd.Parent = t.TabPage
                grd.Visible = True
                grd.Dock = DockStyle.Fill
                grd.DataSource = dst.Tables(IDOC_SEGMENT)
                For Each gcol As UltraWinGrid.UltraGridColumn In grd.DisplayLayout.Bands(0).Columns
                    gcol.Header.Caption = dst.Tables(IDOC_SEGMENT).Columns(gcol.Key).Caption
                Next

            End If
        Next

        EnforceConstraints(True)

        tabIDOC.Visible = True

        Using SR As New System.IO.StreamReader(FILENAME)

            Dim TT As String = SR.ReadToEnd
            Dim CC() As String = Split(TT, vbLf)

            Dim T As String = ""
            Dim IDOC_TABLE As String = "INVOIC01"
            Dim IDOC_ID As String = ""
            Dim IDOC_SEGMENT As String = ""

            Dim SegDocs As New Dictionary(Of String, String)
            Dim SegRels As New Dictionary(Of String, String)

            Dim C As Integer = -1
            Do Until C >= CC.Length - 1
                C += 1 : T = CC(C)
                If T = "" Then Exit Do
                If C = 0 Then
                    IDOC_SEGMENT = "Control"
                Else
                    IDOC_SEGMENT = Trim(Mid(T, 1, 30))
                End If

                Dim p As Integer = 1
                With dst.Tables(IDOC_SEGMENT)
                    Dim row As DataRow = .NewRow
                    For i As Integer = 0 To .Columns.Count - 1
                        Dim l As Integer = .Columns(i).MaxLength
                        row.Item(i) = Mid(T, p, l)
                        p += l
                    Next
                    .Rows.Add(row)
                    ' If IDOC_SEGMENT.StartsWith("E2EDP26") Then Stop
                    If IDOC_SEGMENT <> "Control" And IDOC_SEGMENT <> "Data" And IDOC_SEGMENT <> "Status" Then
                        SegDocs.Add(row.Item("SEGNUM"), IDOC_SEGMENT)
                        If Not SegRels.ContainsKey(IDOC_SEGMENT) Then
                            SegRels.Add(IDOC_SEGMENT, row.Item("PSGNUM"))
                            If row.Item("PSGNUM") = "000000" Then
                                If dst.Relations.Contains("Control" & "_" & IDOC_SEGMENT) Then
                                Else
                                    Create_Relation("Control", IDOC_SEGMENT, "DOCNUM", "DOCNUM")
                                End If
                            Else
                                If dst.Relations.Contains(SegDocs(row.Item("PSGNUM")) & "_" & IDOC_SEGMENT) Then
                                Else
                                    Create_Relation(SegDocs(row.Item("PSGNUM")), IDOC_SEGMENT, "SEGNUM", "PSGNUM")

                                End If
                            End If
                        End If
                    End If

                End With

                If IDOC_SEGMENT <> "Control" And IDOC_SEGMENT <> "Data" And IDOC_SEGMENT <> "Status" Then
                    p = 1
                    With dst.Tables("Data")
                        Dim rowData As DataRow = .NewRow
                        For i As Integer = 0 To .Columns.Count - 1
                            Dim l As Integer = .Columns(i).MaxLength
                            rowData.Item(i) = Mid(T, p, l)
                            p += l
                        Next
                        .Rows.Add(rowData)
                    End With
                End If


            Loop
        End Using


        Dim rowE2EDK02_009 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '009'")(0)
        Dim rowE2EDK02_001 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '001'")(0)
        Dim rowE2EDK02_012s() As DataRow = dst.Tables("E2EDK02").Select("QUALF = '012'")
        Dim rowE2EDK02_012 As DataRow = Nothing

        If rowE2EDK02_012s.Length > 0 Then
            rowE2EDK02_012 = rowE2EDK02_012s(0)
        End If

        Dim rowE2EDK02_087 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '087'")(0)
        Dim rowE2EDS01 As DataRow = dst.Tables("E2EDS01").Select("SUMID = '010'")(0)

        IDOC_DATA_KEY = Trim(rowE2EDK02_009.Item("BELNR"))
        Dim YYYYMMDD As String = Trim(rowE2EDK02_009.Item("DATUM"))
        IDOC_DATA_DATE = CDate(Mid(YYYYMMDD, 5, 2) & "/" & Mid(YYYYMMDD, 7, 2) & "/" & Mid(YYYYMMDD, 1, 4))
        IDOC_DATA_AMT = Val(rowE2EDS01.Item("SUMME") & "")
        IDOC_DATA_REF = Trim(rowE2EDK02_087.Item("BELNR"))
        IDOC_DATA_REF_ORIG = IDOC_DATA_REF

        IDOC_DATA_PACK_LIST_NO = ""
        If rowE2EDK02_012 IsNot Nothing Then
            IDOC_DATA_PACK_LIST_NO = Trim(rowE2EDK02_012.Item("BELNR") & "")
        End If

        If PINV_TYPE = "I" Then
            Dim rowE2EDK02_017 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '017'")(0)
            IDOC_DATA_REF_INV = Trim(rowE2EDK02_017.Item("BELNR"))
        Else
            IDOC_DATA_REF_INV = Trim(rowE2EDK02_009.Item("BELNR"))
            ' IF WE WERE THINKING OF USING THE ORIGINAL INVOICE NO (WHICH WE ARE NOT)
            'Dim MESSAGE As String = Trim(rowE2EDK02_001.Item("BELNR"))
            'IDOC_DATA_REF_INV = "0" & Mid(MESSAGE, 9, 9) ' Invoice 120116629/price discrepancy
            IDOC_DATA_REF = ""
        End If


        IDOC_DATA_TYPE = PINV_TYPE
        'Dim BSART As String = rowE2EDK01005.Item("BSART")
        'IDOC_DATA_TYPE = IIf(BSART = "CRME", "C", "I")
        'If IDOC_DATA_TYPE = "I" And IDOC_DATA_REF_INV <> IDOC_DATA_KEY Then
        '    IDOC_DATA_TYPE = "D"
        'End If

        Absx1.txtFor("IDOC_DATA_KEY").Text = IDOC_DATA_KEY
        Absx1.dteFor("IDOC_DATA_DATE").Value = IDOC_DATA_DATE
        Absx1.numFor("IDOC_DATA_AMT").Value = IDOC_DATA_AMT
        Absx1.txtFor("IDOC_DATA_REF").Text = IDOC_DATA_REF

        Absx1.optFor("IDOC_DATA_TYPE").Value = IDOC_DATA_TYPE
        Absx1.txtFor("IDOC_DATA_REF_INV").Text = IDOC_DATA_REF_INV
        Absx1.txtFor("IDOC_DATA_PACK_LIST_NO").Text = IDOC_DATA_PACK_LIST_NO


        optIDOC_DATE_TYPE.Value = PINV_TYPE
    End Sub

    Private Sub grdTATIDOCU_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdTATIDOCU.DoubleClickRow
        If e.Row.IsDataRow Then
            Dim FILENAME As String = e.Row.Cells("FILENAME").Value & ""
            Dim PINV_TYPE As String = e.Row.Cells("PINV_TYPE").Value & ""
            If PINV_TYPE = "C" Or (PINV_TYPE = "D" And Not chkAllowDN.Checked) Then
                MsgBox("Credit and Debit Memos need to be handled manually." _
                       & vbCrLf & "Obtain the printed Credit Document and make adjustments as required." _
                       & vbCrLf & "You can delete this record from the screen using the right-click Context Menu.", MsgBoxStyle.OkOnly, "Cannot Pull in IPSA IDOC Credit using this Function")
            ElseIf PINV_TYPE <> "I" And PINV_TYPE <> "D" Then
                MsgBox("Unknown Type (" & PINV_TYPE & ")" _
                        & vbCrLf & "You can delete this record from the screen using the right-click Context Menu.", MsgBoxStyle.OkOnly, "Cannot Pull in IPSA IDOC Credit using this Function")

            Else


                'If PINV_TYPE = "D" Then
                '    If txtPO_ORDER_NO_DN.Text = "" Then
                '        MsgBox("For DR Notes, you must provide a valid PO No", MsgBoxStyle.OkOnly, "Cannot Process")
                '        Exit Sub
                '    Else
                '        Dim rowPO As DataRow = LookUp("POTORDR1", txtPO_ORDER_NO_DN.Text)
                '        If rowPO Is Nothing OrElse rowPO.Item("VEND_CODE") <> "IPSA" Then
                '            MsgBox("For DR Notes, you must provide a valid PO No", MsgBoxStyle.OkOnly, "Cannot Process")
                '            Exit Sub
                '        End If
                '    End If
                'End If


                If My.Computer.FileSystem.FileExists(FILENAME) Then

                    Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                    IDOC_FILENAME = FI.Name

                    Process_IDOC(FILENAME, PINV_TYPE)
                    IDOC_Modes("U")
                Else
                    MsgBox("Cannot Find IDOC File " & FILENAME, MsgBoxStyle.OkOnly, "Cannot Process")
                End If
            End If
        End If
    End Sub

    Private Sub grdTATIDOCP_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdTATIDOCP.DoubleClickRow

        'If e.Row.IsDataRow Then
        '    Dim FILENAME As String = e.Row.Cells("FILENAME").Value & ""
        '    Process_IDOC(FILENAME)

        '    IDOC_Modes("P")
        'End If
    End Sub

    Sub IDOC_Modes(IDOC_mode_new As String)

        IDOC_mode = IDOC_mode_new

        tabMainIDOC.Tabs("IDOC Files Received").Visible = (IDOC_mode = "")
        tabMainIDOC.Tabs("IDOC Contents").Visible = Not (IDOC_mode = "")
        UltraExplorerBar1.Groups("IDOC Values").Visible = Not (IDOC_mode = "")

        UltraExplorerBar1.Groups("Import DR Note").Visible = (IDOC_mode = "")

        With UltraExplorerBar1.Groups("IDOCs")
            If IDOC_mode = "" Then
                .Items("IDOC Refresh").Settings.Enabled = DefaultableBoolean.True
                .Items("IDOC Import").Settings.Enabled = DefaultableBoolean.False
                .Items("IDOC Cancel").Settings.Enabled = DefaultableBoolean.False
                .Items("IDOC Done").Settings.Enabled = DefaultableBoolean.False


            Else
                .Items("IDOC Refresh").Settings.Enabled = DefaultableBoolean.False
                If IDOC_mode = "U" Then
                    .Items("IDOC Import").Settings.Enabled = DefaultableBoolean.True
                    .Items("IDOC Cancel").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Items("IDOC Done").Settings.Enabled = DefaultableBoolean.True
                End If
            End If
        End With
    End Sub

    Private Sub grdICTPINV2_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdICTPINV2.InitializeRow
        Dim PO_COST As Decimal = Val(e.Row.Cells("PO_COST").Value & "")
        Dim PINV_COST As Decimal = Val(e.Row.Cells("PINV_COST").Value & "")
        Dim FUT_COST As Decimal = Val(e.Row.Cells("FUT_VCOST").Value & "")
        Dim CUR_COST As Decimal = Val(e.Row.Cells("CUR_VCOST").Value & "")
        Dim PO_TRF_CLASS As String = e.Row.Cells("PO_TRF_CLASS").Value & ""
        Dim FUT_TRF_CLASS As String = e.Row.Cells("FUT_TRF_CLASS").Value & ""
        Dim CUR_TRF_CLASS As String = e.Row.Cells("CUR_TRF_CLASS").Value & ""

        If IsDBNull(e.Row.Cells("FUT_VCOST").Value) Then
            If PO_COST <> CUR_COST Then
                e.Row.Cells("PO_COST").Appearance.ForeColor = Color.Red
                e.Row.Cells("PO_COST").ToolTipText = "PO Cost does not agree with Cur VCost"
            Else
                e.Row.Cells("PO_COST").Appearance.ForeColor = Color.Black
                e.Row.Cells("PO_COST").ToolTipText = String.Empty
            End If
            If PO_TRF_CLASS <> CUR_TRF_CLASS Then
                e.Row.Cells("PO_TRF_CLASS").Appearance.ForeColor = Color.Red
                e.Row.Cells("PO_TRF_CLASS").ToolTipText = "PO Tariff Class does not agree with Cur Tariff Class"
            Else
                e.Row.Cells("PO_TRF_CLASS").Appearance.ForeColor = Color.Black
                e.Row.Cells("PO_TRF_CLASS").ToolTipText = String.Empty
            End If
        Else
            If PO_COST <> FUT_COST Then
                e.Row.Cells("PO_COST").Appearance.ForeColor = Color.Red
                e.Row.Cells("PO_COST").ToolTipText = "PO Cost does not agree with Fut VCost"
            Else
                e.Row.Cells("PO_COST").Appearance.ForeColor = Color.Black
                e.Row.Cells("PO_COST").ToolTipText = String.Empty
            End If
            If PO_TRF_CLASS <> FUT_TRF_CLASS Then
                e.Row.Cells("PO_TRF_CLASS").Appearance.ForeColor = Color.Red
                e.Row.Cells("PO_TRF_CLASS").ToolTipText = "PO Tariff Class does not agree with Fut Tariff Class"
            Else
                e.Row.Cells("PO_TRF_CLASS").Appearance.ForeColor = Color.Black
                e.Row.Cells("PO_TRF_CLASS").ToolTipText = String.Empty
            End If
        End If


        If PINV_COST <> PO_COST Then
            e.Row.Cells("PINV_COST").Appearance.ForeColor = Color.Red
            e.Row.Cells("PINV_COST").ToolTipText = "Invoice Cost does not agree with PO Cost"
        Else
            e.Row.Cells("PINV_COST").Appearance.ForeColor = Color.Black
            e.Row.Cells("PINV_COST").ToolTipText = String.Empty
        End If

        Dim PINV_NOTE As String = e.Row.Cells("PINV_NOTE").Value & ""
        If PINV_NOTE <> "" Then
            e.Row.Cells("PINV_NOTE").Appearance.ForeColor = Color.Red
        Else
            e.Row.Cells("PINV_NOTE").Appearance.ForeColor = Color.Empty
        End If
    End Sub

    Private Sub btnGenerateReceipt_Click(sender As Object, e As EventArgs) Handles btnGenerateReceipt.Click

        Stop

        ' Change for ADS 07/16/2025, force developer to supply the LP CODE
        TAC.ICCMAIN1.Transmit_Document("WHC", "WHC943O1", "SPI", "002005", "CLA")


        Stop

        ' put POs to generate receipts for
        ' code will use ICTPINV2 to set up a receipt
        ' note that if there are multiple ICTPINV1s for a PO, that multiple receipts will be generated

        ' !!!
        ' NOTE AS OF DEC 2025 - this may need to have the warehouse changed if run in the future as CLA is no longer the default
        ASCMAIN1.sql = "insert into EDTTRXN1" & vbCrLf _
            & "SELECT " & vbCrLf _
            & "ICTPINV2.PINV_NO TRX_NO," & vbCrLf _
            & "ICTPINV2.PINV_LNO TRX_LNO," & vbCrLf _
            & "'REC' TRANS_TYPE," & vbCrLf _
            & "SUBSTR(ICTPINV1.INV_NUM,4) TRANS_NUM," & vbCrLf _
            & "ICTPINV1.INV_DATE," & vbCrLf _
            & "ICTPINV1.PO_ORDER_NO," & vbCrLf _
            & "NULL BUYER," & vbCrLf _
            & "NULL REASON_CODE," & vbCrLf _
            & "'XXX' OPERATOR," & vbCrLf _
            & "ICTPINV2.ITEM_CODE," & vbCrLf _
            & "NULL LOCATION," & vbCrLf _
            & "ICTPINV2.PINV_QTY TRAN_QTY," & vbCrLf _
            & "'0' PROCESS_IND," & vbCrLf _
            & "'CLA' WHSE_CODE" & vbCrLf _
            & "FROM ICTPINV2,ICTPINV1 " & vbCrLf _
            & "WHERE ICTPINV1.PO_ORDER_NO IN ('132419')" & vbCrLf _
            & "AND ICTPINV2.PINV_NO = ICTPINV1.PINV_NO" & vbCrLf _
            & "AND ICTPINV2.PINV_QTY <> 0"
        ASCDATA1.ExecuteSQL()
    End Sub

    Private Sub optIDOC_DATE_TYPE_ValueChanged(sender As Object, e As EventArgs) Handles optIDOC_DATE_TYPE.ValueChanged
        lblIDOC_REF_INV.Visible = (optIDOC_DATE_TYPE.Value & "" = "D") Or (optIDOC_DATE_TYPE.Value & "" = "C")
        txtIDOC_REF_INV.Visible = (optIDOC_DATE_TYPE.Value & "" = "D") Or (optIDOC_DATE_TYPE.Value & "" = "C")
    End Sub

    Private Sub grdTATIDOCU_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs)

    End Sub

    Private Sub grdTATIDOCU_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdTATIDOCU.InitializeRow
        If e.Row.IsDataRow Then
            Dim PINV_TYPE As String = e.Row.Cells("PINV_TYPE").Value & ""

            If PINV_TYPE <> "I" Then
                e.Row.Cells("PINV_TYPE").Appearance.BackColor = Color.Yellow
                If PINV_TYPE = "C" Or PINV_TYPE = "D" Then
                    e.Row.Cells("PINV_TYPE").Appearance.ForeColor = Color.Red
                End If
            End If
        End If
    End Sub

    Private Sub grdICTPINVX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPINVX.InitializeLayout

    End Sub

    Private Sub grdICTPINVY_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPINVY.InitializeLayout

    End Sub


    Private Sub grdICTPINVX_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdICTPINVX.DragEnter
        If grdICTPINVX.AllowDrop Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub grdICTPINVX_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles grdICTPINVX.DragDrop
        eDND = e
        ' Process_DragDrop()

        Dim files() As String = eDND.Data.GetData(DataFormats.FileDrop)
        If files.Count <> 1 Then
            MsgBox("Please Drag and Drop 1 file", MsgBoxStyle.OkOnly, "Cannot Import File")
            Exit Sub
        End If

        If MsgBox($"OK to Import Data from {files(0)}", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        Dim FILENAME As String = files(0)
        ImportContainerInfo(FILENAME)
    End Sub

    Sub ImportContainerInfo(FILENAME As String)

        If chkEnableEditToShipment.Checked Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading XLS file selected")

        Try
            Dim oWB As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
            Dim oSheet As SpreadsheetGear.IWorksheet = oWB.Worksheets(0)
            Dim range As SpreadsheetGear.IRange = Nothing

            Dim lastrow As Integer = oSheet.UsedRange.RowCount

            dst.Tables("ICTPINVV").Rows.Clear()

            Dim r As Integer = 1
            Dim rowICTPINVV As DataRow = Nothing
            Dim rows_loaded As Integer = 0

            Do While r < lastrow Or oSheet.Cells(r, 0).Value & "" <> "" '  AndAlso (oSheet.Cells(r, 0).Value & "" <> "" Or oSheet.Cells(r + 1, 0).Value & "" <> "" Or oSheet.Cells(r, 1).Value & "" <> "" Or oSheet.Cells(r + 1, 1).Value & "" <> "")
                If CStr(oSheet.Cells(r, 0).Value & "").StartsWith("*") Or (oSheet.Cells(r, 0).Value & "" = "" And oSheet.Cells(r, 1).Value & "" = "") Then

                Else
                    Dim row_loaded As Boolean = False

                    Try
                        Dim INVOICEs As String = oSheet.Cells(r, 3).Value & ""
                        Dim INVOICE_NOs() As String = Split(INVOICEs & "-", "-")
                        For Each INVOICE_NO As String In INVOICE_NOs
                            If INVOICE_NO <> "" Then

                                rowICTPINVV = dst.Tables("ICTPINVV").NewRow
                                rowICTPINVV.Item("PINV_NO") = "XLS Row " & Format(r + 1, "000000")
                                rowICTPINVV.Item("ERROR") = "No Matching Record Found"

                                dst.Tables("ICTPINVV").Rows.Add(rowICTPINVV)

                                With rowICTPINVV
                                    .Item("INV_NUM") = INVOICE_NO
                                    Dim INV_NUM As String = Split(INVOICE_NO & "-", "-")(0)
                                    If Len(INV_NUM) <= 10 Then
                                        INV_NUM = Format(Val(INV_NUM), "0000000000")
                                        .Item("INV_NUM") = INV_NUM

                                        ASCMAIN1.sql = $"Select * from ICTPINV1 where INV_NUM = '{INV_NUM}' and (PINV_STATUS = 'I' or PINV_STATUS = 'O')"
                                        Dim rowICTPINV1 As DataRow = ASCDATA1.GetDataRow
                                        If rowICTPINV1 IsNot Nothing Then
                                            .Item("PINV_NO") = rowICTPINV1.Item("PINV_NO")
                                            .Item("INV_DATE") = rowICTPINV1.Item("INV_DATE")
                                            .Item("ERROR") = DBNull.Value
                                        End If
                                    End If

                                    .Item("VESSEL_NAME") = oSheet.Cells(r, 7).Value & ""
                                    .Item("BOL_NO") = oSheet.Cells(r, 0).Value & ""
                                    .Item("CONTAINER_NO") = oSheet.Cells(r, 1).Value & ""
                                    .Item("SHIP_DATE") = CDate(oWB.NumberToDateTime(oSheet.Cells(r, 5).Value & ""))
                                    .Item("ETA_DATE") = CDate(oWB.NumberToDateTime(oSheet.Cells(r, 6).Value & ""))

                                End With
                                row_loaded = True
                            End If
                        Next

                        If row_loaded Then rows_loaded += 1

                    Catch ex As Exception
                        rowICTPINVV.Item("ERROR") = ex.Message
                    End Try
                End If


                r = r + 1
            Loop

            Using frm As New ASFMSGBF
                frm.Show_grd(dst.Tables("ICTPINVV"), Me, "Imported Data")
            End Using

            Dim rows_to_be_updated As Integer = dst.Tables("ICTPINVV").Select("ISNULL(ERROR,'')=''").Length

            If rows_to_be_updated > 0 Then

                If MsgBox("OK to Update Database with Imported Data?" _
                          & vbCrLf & "(Note: rows with Errors will not be updated)" _
                          & vbCrLf & vbCrLf & $"Last Row evaluated: {CStr(r + 1)}" & vbCrLf & $"{rows_loaded} Excel rows loaded" & vbCrLf & $"{rows_to_be_updated} invoices will be updated.", MsgBoxStyle.YesNo, "") = MsgBoxResult.Yes Then
                    dst.Tables("ICTPINV1").Rows.Clear()
                    DATETIME_STAMP = Now + ASCMAIN1.NowTSD

                    For Each rowICTPINVV In dst.Tables("ICTPINVV").Select("ISNULL(ERROR,'')=''")
                        Dim PINV_NO As String = rowICTPINVV.Item("PINV_NO")
                        Dim rowICTPINV1 As DataRow = Fill_Record("ICTPINV1", PINV_NO, False, False)
                        With rowICTPINV1
                            .Item("VESSEL_NAME") = rowICTPINVV.Item("VESSEL_NAME")
                            .Item("BOL_NO") = rowICTPINVV.Item("BOL_NO")
                            .Item("CONTAINER_NO") = rowICTPINVV.Item("CONTAINER_NO")
                            .Item("SHIP_DATE") = rowICTPINVV.Item("SHIP_DATE")
                            .Item("ETA_DATE") = rowICTPINVV.Item("ETA_DATE")

                            .Item("LAST_DATE") = DATETIME_STAMP
                            .Item("LAST_OPER") = ASCMAIN1.USER_ID
                        End With
                    Next

                    Update_Record_TDA("ICTPINV1")

                    MsgBox("Update Complete", MsgBoxStyle.OkOnly, "Success")

                    dst.Tables("ICTPINV1").Rows.Clear()

                End If

            End If
        Catch ex As Exception
            MsgBox($"Error encountered: {ex.Message}", MsgBoxStyle.OkOnly, "Cannot Import from File")
        End Try

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Process_DragDrop2()

        'Application.DoEvents()

        'Dim files() As String = eDND.Data.GetData(DataFormats.FileDrop)
        'If files.Length = 1 Then
        '    Dim FILENAME As String = files(0)
        '    If FILENAME.ToUpper Like "*.XLS" Or FILENAME.ToUpper Like "*.XLSX" Then
        '        Try
        '            Load_Workbook(FILENAME)
        '        Catch ex As Exception
        '            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Cannot Load from Workbook")
        '        End Try
        '    End If
        'End If

        'Application.DoEvents()

    End Sub

    Private Sub chkEnableEditToShipment_CheckedChanged(sender As Object, e As EventArgs) Handles chkEnableEditToShipment.CheckedChanged
        If chkEnableEditToShipment.Checked Then
            grdICTPINVX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            If Not ASCMAIN1.Logical_Lock("F", "ICFIREC1", False,,, 10) Then
                chkEnableEditToShipment.Checked = False
                Exit Sub
            End If
            grdICTPINVX.DisplayLayout.ScrollStyle = ScrollStyle.Immediate
            grdICTPINVX.DisplayLayout.ColScrollRegions(0).Scroll(ColScrollAction.Right)
        Else
            grdICTPINVX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            ASCMAIN1.MultiTask_Release(,, 10)
            grdICTPINVX.DisplayLayout.ScrollStyle = ScrollStyle.Immediate
            grdICTPINVX.DisplayLayout.ColScrollRegions(0).Scroll(ColScrollAction.Left)
        End If

        For Each c As String In editableColumns
            grdICTPINVX.DisplayLayout.Bands(0).Columns(c).CellAppearance.BackColor = IIf(chkEnableEditToShipment.Checked, Color.Yellow, Color.Empty)
        Next
    End Sub

    Private Sub grdICTPINVX_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdICTPINVX.AfterRowUpdate

        If chkEnableEditToShipment.Checked Then
            DATETIME_STAMP = Now + ASCMAIN1.NowTSD
            Dim PINV_NO As String = grdICTPINVX.ActiveRow.Cells("PINV_NO").Value
            Dim rowICTPINV1 As DataRow = Fill_Record("ICTPINV1", PINV_NO)
            For Each C As String In editableColumns
                rowICTPINV1.Item(C) = grdICTPINVX.ActiveRow.Cells(C).Value
            Next
            Update_Record_TDA("ICTPINV1")
        End If
    End Sub
    Private Sub chkShowDelIDOC_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowDelIDOC.CheckedChanged
        ' grdTATIDOCD.Visible = chkShowDelIDOC.Checked
        SplitContainer3.Panel2Collapsed = Not chkShowDelIDOC.Checked

        Refresh_IDOCs()
    End Sub

    Private Sub grdICTPINV2_BeforeRowsDeleted(sender As Object, e As BeforeRowsDeletedEventArgs) Handles grdICTPINV2.BeforeRowsDeleted
        For Each grow As UltraWinGrid.UltraGridRow In grdICTPINV2.Selected.Rows
            If grow.Cells("PINV_QTY").Value = 0 Then
                DEL_LINE = grow.Cells("PINV_LNO").Value & ""
                DEL_ITEM_CODE = grow.Cells("ITEM_CODE").Value & ""
                DEL_ITEM_DESC = grow.Cells("ITEM_DESC").Value & ""
                '  DELETED_INV_LINES.Add($"Line No {DEL_LINE} Deleted Item {DEL_ITEM_CODE} - {DEL_ITEM_DESC}")
            Else
                MsgBox("Cannot Delete Invoice Line " & grow.Cells("PINV_LNO").Value & " with Pinv Qty (" & Format(grow.Cells("PINV_QTY").Value, "##,###,##0") & ")", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                e.Cancel = True
                Exit For
            End If
        Next
    End Sub

    Function ReTransmit(PO_ORDER_NO As String, PINV_NO As String, WHSE_CODE As String) As Boolean

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        Dim LP_CODE As String = rowICTWHSE1.Item("LP_CODE") & ""

        If LP_CODE = "" Then
            MsgBox($"No 3PL Code set up for Warehouse {WHSE_CODE}", MsgBoxStyle.OkOnly, "Cannot Re-Transmit")
            Return False
        Else

            If LP_CODE = "ADS" Then ' ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "wjz" Or ASCMAIN1.USER_ID = "dgj") Then
                ' do this only if ADS or testing with CLA - retransmit does not normally send item master to CLA

                ' Send Item Master
                Transmit_888(PO_ORDER_NO, LP_CODE)

                System.Threading.Thread.Sleep(1500) ' TO ENSURE UNIQUE DATETIME STAMP
            End If

            Transmit_943(PO_ORDER_NO, PINV_NO, LP_CODE)
            MsgBox("PO Receiving Advice has been successfully Re-Transmitted", MsgBoxStyle.OkOnly, "Verification")
            Return True
        End If

    End Function

    Private Sub cmdPackList_Click(sender As Object, e As EventArgs) Handles cmdPackList.Click
        Dim folder As String = "\\Nymain-abs-iis1\sftp\IPSA\PROD\FROM_IPSA\IDOC\archive\"

        ASCMAIN1.sql = "Select * from TATIDOCP WHERE IDOC_FILEDATE > SYSDATE -365 * 5 AND IDOC_DATA_PACK_LIST_NO IS NULL"
        Fill_Records("TATIDOCP", ,, ASCMAIN1.sql)
        '        ASCMAIN1.sql = "select TATIDOCP.*, SUBSTR(IDOC_FILENAME,1+INSTR(IDOC_FILENAME,'\',-1)) LEAF from TATIDOCP"
        For Each row As DataRow In dst.Tables("TATIDOCP").Select("") '  ASCDATA1.GetDataTable.Select("")
            Dim IDOC_FILENAME As String = row.Item("IDOC_FILENAME") & ""
            IDOC_DATA_PACK_LIST_NO = row.Item("IDOC_DATA_PACK_LIST_NO") & ""
            Dim IDOC_DATA_TYPE As String = row.Item("IDOC_DATA_TYPE") & ""
            If IDOC_DATA_PACK_LIST_NO = "" And IDOC_FILENAME <> "" And IDOC_DATA_TYPE = "I" Then
                Dim FN() As String = Split(IDOC_FILENAME, "\")
                Dim LEAF As String = FN(FN.Length - 1)
                Dim FILENAME As String = folder & LEAF


                If My.Computer.FileSystem.FileExists(FILENAME) Then

                    Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
                    IDOC_FILENAME = FI.Name

                    ASCMAIN1.Progress("-", Mid(LEAF, 8))

                    Try
                        Process_IDOC(FILENAME, IDOC_DATA_TYPE)
                        row.Item("IDOC_DATA_PACK_LIST_NO") = IDOC_DATA_PACK_LIST_NO
                    Catch ex As Exception

                    End Try

                Else
                    ' MsgBox("Cannot Find IDOC File " & FILENAME, MsgBoxStyle.OkOnly, "Cannot Process")
                End If


            End If
        Next

        Update_Record_TDA("TATIDOCP")

        MsgBox("Packing List Load Complete", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Private Sub chkEnableEditToShipment_EnabledChanged(sender As Object, e As EventArgs) Handles chkEnableEditToShipment.EnabledChanged

    End Sub
End Class