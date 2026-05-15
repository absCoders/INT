Imports System.IO

Public Class WHC943O1

    ' Create Anticipated Receipts File (Outbound) 943

    Inherits WHC000O1

    ' THESE ARE THE FILENAMES USED FOR THE CLA INTERFACE
    Dim POHDR As String = "POHDR" & "_" & DTS & ".CSV"
    Dim PODTL As String = "PODTL" & "_" & DTS & ".CSV"

    Private Const PORetransmitChar As String = "-"

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHC943O1"

        Create_Work_Table()

        With dst
            ASCMAIN1.sql = "Select ICTPINV1.INV_NUM, POTORDR1.PO_ORDER_NO, POTORDR1.PO_DATE_ORDERED, ICTPINV1.PINV_NO, ICTPINV1.INV_DATE, POTORDR1.PO_HDR_CTR_REV" & vbCrLf _
                & " from ICTPINV1,POTORDR1"
            Create_TDA(.Tables.Add, "POHDR", "**", 0, False, "", 2)
            .Tables("POHDR").Columns("PINV_NO").AllowDBNull = True

            ASCMAIN1.sql = "Select ICTPINV1.INV_NUM, POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
                & ", POTORDR2.ITEM_CODE, POTORDR2.PO_QTY_OPN, POTORDR2.PO_COST" & vbCrLf _
                & ", POTORDR2.PO_DATE_REQUIRED, POTORDR2.PO_STATUS" & vbCrLf _
                & ", ICTITEM1.ITEM_ALT_SORT, ICTITEM1.ITEM_EAN_CODE" & vbCrLf _
                & " from POTORDR2,ICTITEM1,ICTPINV1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE"
            Create_TDA(.Tables.Add, "PODTL", "**", 0, False, "", 3)

            With .Tables("PODTL").Columns
                .Add("PO_AMT_OPN", GetType(System.Decimal), "ISNULL(PO_QTY_OPN,0) * ISNULL(PO_COST,0)")
            End With

            Create_Relation("POHDR", "PODTL", "INV_NUM,PO_ORDER_NO")

            With .Tables("POHDR").Columns
                .Add("PO_AMT_OPN", GetType(System.Decimal), "SUM(CHILD.PO_AMT_OPN)")
            End With

            Create_TDA(.Tables.Add, "ICTIXFR1", "*")

            ASCMAIN1.sql = "SELECT ICTIXFR2.*, ICTITEM1.ITEM_ALT_SORT, ICTITEM1.ITEM_EAN_CODE
                                FROM ICTIXFR2, ICTITEM1
                                where ICTIXFR2.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                                AND ICTIXFR2.XFR_NO = :PARM1"
            Create_TDA(.Tables.Add, "ICTIXFR2", ASCMAIN1.sql, 0, False, "V", 2)

            ASCMAIN1.sql = "SELECT POTSHIP2.WHSE_CODE, POTSHIP2.CONTAINER_NO, POTSHIP1.INIT_DATE, POTASNM3.*, ICTITEM1.ITEM_ALT_SORT, ICTITEM1.ITEM_EAN_CODE
                                FROM POTSHIP1, POTSHIP2, POTASNM1, POTASNM3, ICTITEM1
                                WHERE POTSHIP1.PO_SHIPMENT_NO = POTSHIP2.PO_SHIPMENT_NO
                                AND POTSHIP2.CONTAINER_CTL_NO = POTASNM1.CONTAINER_CTL_NO
                                AND POTASNM1.ASN_NO = POTASNM3.ASN_NO
                                AND POTASNM3.ASN_XMIT_NO = :PARM1
                                AND POTASNM3.ITEM_CODE = ICTITEM1.ITEM_CODE (+)"
            Create_TDA(.Tables.Add, "POTASNM3", ASCMAIN1.sql, 0, False, "V")
        End With

        Main_Process()
    End Sub

    Public Sub Main_Process()

        Dim POHDR_PO As String = "Select 0 INV_NUM, POTORDR1.PO_ORDER_NO, POTORDR1.PO_DATE_ORDERED, NULL PINV_NO, NULL INV_DATE" & vbCrLf _
            & ", APTVEND1.VEND_SUPPLIER_ID, POTORDR1.VEND_NAME, POTORDR1.VEND_ADDR1, POTORDR1.VEND_ADDR2" & vbCrLf _
            & ", POTORDR1.VEND_CITY, POTORDR1.VEND_STATE, POTORDR1.VEND_ZIP_CODE, POTORDR1.VEND_COUNTRY, POTORDR1.PO_HDR_CTR_REV" & vbCrLf _
            & " from POTORDR1,APTVEND1 where APTVEND1.VEND_CODE = POTORDR1.VEND_CODE and POTORDR1.PO_ORDER_NO = {0}"

        ' NEW
        Dim PODTL_PO As String = "Select 0 INV_NUM, POTORDR2.PO_ORDER_NO, MAX(POTORDR2.PO_ORDER_LNO) PO_ORDER_LNO" & vbCrLf _
            & ", POTORDR2.ITEM_CODE, SUM(POTORDR2.PO_QTY_OPN) PO_QTY_OPN, MAX(POTORDR2.PO_COST) PO_COST" & vbCrLf _
            & ", MIN(POTORDR2.PO_DATE_REQUIRED) PO_DATE_REQUIRED, MIN(POTORDR2.PO_STATUS) PO_STATUS" & vbCrLf _
            & ", MIN(ICTITEM1.ITEM_ALT_SORT) ITEM_ALT_SORT, MIN(NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE)) ITEM_EAN_CODE" & vbCrLf _
            & " from POTORDR2,ICTITEM1" & vbCrLf _
            & " where POTORDR2.PO_ORDER_NO = {0}" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
            & " GROUP BY 0, POTORDR2.PO_ORDER_NO,  POTORDR2.ITEM_CODE"

        ' OLD
        'Dim PODTL_PO As String = "Select 0 INV_NUM, POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
        '    & ", POTORDR2.ITEM_CODE, POTORDR2.PO_QTY_OPN, POTORDR2.PO_COST" & vbCrLf _
        '    & ", POTORDR2.PO_DATE_REQUIRED, POTORDR2.PO_STATUS" & vbCrLf _
        '    & ", ICTITEM1.ITEM_ALT_SORT, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE" & vbCrLf _
        '    & " from POTORDR2,ICTITEM1" & vbCrLf _
        '    & " where POTORDR2.PO_ORDER_NO = {0}" & vbCrLf _
        '    & "   and ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE"

        Dim POHDR_PI As String = "Select ICTPINV1.INV_NUM, ICTPINV1.PO_ORDER_NO, ICTPINV1.INV_DATE PO_DATE_ORDERED, ICTPINV1.PINV_NO, ICTPINV1.INV_DATE" & vbCrLf _
            & ", APTVEND1.VEND_SUPPLIER_ID, POTORDR1.VEND_NAME, POTORDR1.VEND_ADDR1, POTORDR1.VEND_ADDR2" & vbCrLf _
            & ", POTORDR1.VEND_CITY, POTORDR1.VEND_STATE, POTORDR1.VEND_ZIP_CODE, POTORDR1.VEND_COUNTRY, POTORDR1.PO_HDR_CTR_REV" & vbCrLf _
            & " from ICTPINV1,POTORDR1,APTVEND1 where APTVEND1.VEND_CODE = POTORDR1.VEND_CODE and POTORDR1.PO_ORDER_NO = ICTPINV1.PO_ORDER_NO and ICTPINV1.PINV_NO = {0}"

        ' NEED TO CHANGE PO_DATE_REQUIRED TO AN ANTICIPATED RECEIPT DATE FROM THE PRE-INVOICE DATA

        Dim PODTL_PI As String = "Select ICTPINV1.INV_NUM, POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ", POTORDR2.ITEM_CODE, ICTPINV2.PINV_QTY PO_QTY_OPN, POTORDR2.PO_COST" & vbCrLf _
            & ", POTORDR2.PO_DATE_REQUIRED, ICTPINV1.PINV_STATUS PO_STATUS" & vbCrLf _
            & ", ICTITEM1.ITEM_ALT_SORT, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE" & vbCrLf _
            & " from POTORDR2,ICTITEM1,ICTPINV1,ICTPINV2" & vbCrLf _
            & " where ICTPINV1.PINV_NO = {0}" & vbCrLf _
            & "   and ICTPINV2.PINV_NO = ICTPINV1.PINV_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = ICTPINV2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = ICTPINV2.PO_ORDER_LNO" & vbCrLf _
            & "   and ICTPINV2.PINV_QTY <> 0" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE"

        ' to accommodate when IPSA invoices the same PO line on 2 lines

        PODTL_PI = "Select ICTPINV1.INV_NUM, POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ", POTORDR2.ITEM_CODE, sum (ICTPINV2.PINV_QTY) PO_QTY_OPN, POTORDR2.PO_COST" & vbCrLf _
            & ", POTORDR2.PO_DATE_REQUIRED, ICTPINV1.PINV_STATUS PO_STATUS" & vbCrLf _
            & ", ICTITEM1.ITEM_ALT_SORT, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE" & vbCrLf _
            & " from POTORDR2,ICTITEM1,ICTPINV1,ICTPINV2" & vbCrLf _
            & " where ICTPINV1.PINV_NO = {0}" & vbCrLf _
            & "   and ICTPINV2.PINV_NO = ICTPINV1.PINV_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_NO = ICTPINV2.PO_ORDER_NO" & vbCrLf _
            & "   and POTORDR2.PO_ORDER_LNO = ICTPINV2.PO_ORDER_LNO" & vbCrLf _
            & "   and ICTPINV2.PINV_QTY <> 0" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
            & " group by " & vbCrLf _
            & " ICTPINV1.INV_NUM, POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ", POTORDR2.ITEM_CODE, POTORDR2.PO_COST" & vbCrLf _
            & ", POTORDR2.PO_DATE_REQUIRED, ICTPINV1.PINV_STATUS" & vbCrLf _
            & ", ICTITEM1.ITEM_ALT_SORT, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE)"


        EnforceConstraints(False)

        If G.APP_CMD = "" Then
            ' do whatever

        ElseIf G.APP_CMD = "SPO" Then ' Single PO

            Fill_Records("POHDR", "", True, Replace(POHDR_PO, "= {0}", "= '" & G.APP_KEY & "'"))
            Fill_Records("PODTL", "", True, Replace(PODTL_PO, "= {0}", "= '" & G.APP_KEY & "'"))
            'dst.Tables("POHDR").Columns("PINV_NO").AllowDBNull = True

        ElseIf G.APP_CMD = "SPI" Then ' Single IPSA Invoice

            Fill_Records("POHDR", "", True, Replace(POHDR_PI, "= {0}", "= '" & G.APP_KEY & "'"))
            Fill_Records("PODTL", "", True, Replace(PODTL_PI, "= {0}", "= '" & G.APP_KEY & "'"))

        ElseIf G.APP_CMD = "MPO" Then ' Multiple POs

            Fill_Records("POHDR", "", True, Replace(POHDR_PO, "= {0}", "in ('" & Replace(G.APP_KEY, ",", "','") & "')"))
            Fill_Records("PODTL", "", True, Replace(PODTL_PO, "= {0}", "in ('" & Replace(G.APP_KEY, ",", "','") & "')"))

        ElseIf G.APP_CMD = "MPI" Then ' Multiple IPSA Invoices

            Fill_Records("POHDR", "", True, Replace(POHDR_PI, "= {0}", "in ('" & Replace(G.APP_KEY, ",", "','") & "')"))
            Fill_Records("PODTL", "", True, Replace(PODTL_PI, "= {0}", "in ('" & Replace(G.APP_KEY, ",", "','") & "')"))

        ElseIf G.APP_CMD = "XFR_NO" Then
            ConvertTransferToPurchaseOrder(G.APP_KEY)

        ElseIf G.APP_CMD = "ASN" Then
            CreateASN(G.APP_KEY)

        ElseIf G.APP_CMD = "CANC" Then
            ' ISSUE-7044 - balance POs/ASNs
        End If

        EnforceConstraints(True)

        Update_Record()

    End Sub

    Private Sub CreateASN(ASN_XMIT_NO As String)

        Fill_Records("POTASNM3", ASN_XMIT_NO)
        If dst.Tables("POTASNM3").Rows.Count = 0 Then
            Exit Sub
        End If

        Dim drHeader As DataRow = dst.Tables("POTASNM3").Rows(0)

        Dim INV_NUM_PREFIX As String = String.Empty
        Select Case drHeader.Item("WHSE_CODE") & String.Empty
            Case "ADSMIN", "ADS"
                INV_NUM_PREFIX = "-M"
            Case "ADSBLA"
                INV_NUM_PREFIX = "-B"
        End Select

        Dim PO_ORDER_NO As String = "A" & Val(drHeader.Item("ASN_NO") & String.Empty).ToString.PadLeft(5, "0")
        Dim INV_NUM As String = drHeader.Item("CONTAINER_NO") & INV_NUM_PREFIX

        Dim drPOHDR As DataRow = dst.Tables("POHDR").NewRow
        drPOHDR.Item("INV_NUM") = INV_NUM
        drPOHDR.Item("PO_ORDER_NO") = PO_ORDER_NO
        drPOHDR.Item("PO_DATE_ORDERED") = drHeader.Item("INIT_DATE")
        drPOHDR.Item("PINV_NO") = DBNull.Value
        drPOHDR.Item("INV_DATE") = DateTime.Now.ToShortDateString
        dst.Tables("POHDR").Rows.Add(drPOHDR)

        Dim PO_ORDER_LNO As Int16 = 1
        For Each drPOTASNM3 As DataRow In dst.Tables("POTASNM3").Select("", "ITEM_CODE")
            Dim drPODTL As DataRow = dst.Tables("PODTL").NewRow
            drPODTL.Item("INV_NUM") = INV_NUM
            drPODTL.Item("PO_ORDER_NO") = PO_ORDER_NO
            drPODTL.Item("PO_ORDER_LNO") = PO_ORDER_LNO
            PO_ORDER_LNO += 1
            drPODTL.Item("ITEM_CODE") = drPOTASNM3.Item("ITEM_CODE")
            drPODTL.Item("PO_QTY_OPN") = drPOTASNM3.Item("ASN_QTY")
            drPODTL.Item("PO_COST") = drPOTASNM3.Item("ASN_COST")
            drPODTL.Item("PO_DATE_REQUIRED") = DateTime.Now.AddDays(3).ToShortDateString
            drPODTL.Item("PO_STATUS") = "O"
            drPODTL.Item("ITEM_ALT_SORT") = drPOTASNM3.Item("ITEM_ALT_SORT")
            drPODTL.Item("ITEM_EAN_CODE") = drPOTASNM3.Item("ITEM_EAN_CODE")
            dst.Tables("PODTL").Rows.Add(drPODTL)
        Next

    End Sub

    Private Sub ConvertTransferToPurchaseOrder(ByVal XFR_NO As String)

        Fill_Records("ICTIXFR1", "", True, $"SELECT * FROM ICTIXFR1 WHERE XFR_NO = '{XFR_NO}'")
        Fill_Records("ICTIXFR2", XFR_NO)

        'ASCMAIN1.sql = "Select ICTPINV1.INV_NUM, POTORDR1.PO_ORDER_NO, POTORDR1.PO_DATE_ORDERED, ICTPINV1.PINV_NO, ICTPINV1.INV_DATE" & vbCrLf _
        '    & " from ICTPINV1,POTORDR1"
        'Create_TDA(.Tables.Add, "POHDR", "**", 0, False, "", 2)
        '.Tables("POHDR").Columns("PINV_NO").AllowDBNull = True

        'ASCMAIN1.sql = "Select ICTPINV1.INV_NUM, POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO" & vbCrLf _
        '    & ", POTORDR2.ITEM_CODE, POTORDR2.PO_QTY_OPN, POTORDR2.PO_COST" & vbCrLf _
        '    & ", POTORDR2.PO_DATE_REQUIRED, POTORDR2.PO_STATUS" & vbCrLf _
        '    & ", ICTITEM1.ITEM_ALT_SORT, ICTITEM1.ITEM_EAN_CODE" & vbCrLf _
        '    & " from POTORDR2,ICTITEM1,ICTPINV1" & vbCrLf _
        '    & " where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE"
        'Create_TDA(.Tables.Add, "PODTL", "**", 0, False, "", 3)

        Dim drICTIXFR1 As DataRow = dst.Tables("ICTIXFR1").Rows(0)
        Dim PO_ORDER_NO As String = "X" & Val(XFR_NO).ToString.PadLeft(5, "0")
        Dim INV_NUM As String = "XFR" & XFR_NO

        Dim drPOHDR As DataRow = dst.Tables("POHDR").NewRow
        drPOHDR.Item("INV_NUM") = INV_NUM
        drPOHDR.Item("PO_ORDER_NO") = PO_ORDER_NO
        drPOHDR.Item("PO_DATE_ORDERED") = drICTIXFR1.Item("XFR_DATE")
        drPOHDR.Item("PINV_NO") = DBNull.Value
        drPOHDR.Item("INV_DATE") = DateTime.Now.ToShortDateString
        dst.Tables("POHDR").Rows.Add(drPOHDR)

        For Each drICTIXFR2 As DataRow In dst.Tables("ICTIXFR2").Select("", "XFR_LNO")
            Dim drPODTL As DataRow = dst.Tables("PODTL").NewRow
            drPODTL.Item("INV_NUM") = INV_NUM
            drPODTL.Item("PO_ORDER_NO") = PO_ORDER_NO
            drPODTL.Item("PO_ORDER_LNO") = drICTIXFR2.Item("XFR_LNO")
            drPODTL.Item("ITEM_CODE") = drICTIXFR2.Item("ITEM_CODE")
            drPODTL.Item("PO_QTY_OPN") = drICTIXFR2.Item("XFR_QTY")
            drPODTL.Item("PO_COST") = drICTIXFR2.Item("ITEM_COST_STD")
            drPODTL.Item("PO_DATE_REQUIRED") = DateTime.Now.AddDays(3).ToShortDateString
            drPODTL.Item("PO_STATUS") = "O"
            drPODTL.Item("ITEM_ALT_SORT") = drICTIXFR2.Item("ITEM_ALT_SORT")
            drPODTL.Item("ITEM_EAN_CODE") = drICTIXFR2.Item("ITEM_EAN_CODE")
            dst.Tables("PODTL").Rows.Add(drPODTL)
        Next

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Select Case LP_CODE

            Case "CLA"
                Using swPOHDR As New System.IO.StreamWriter(POHDR)
                    Using swPODTL As New System.IO.StreamWriter(PODTL)
                        For Each rowPOHDR As DataRow In dst.Tables("POHDR").Select("", "INV_NUM,PO_ORDER_NO")
                            'R += 1
                            Dim PO_ORDER_NO As String = rowPOHDR.Item("PO_ORDER_NO") & ""
                            If PO_ORDER_NO.Length <> 6 Then Stop
                            Dim INV_NUM As String = rowPOHDR.Item("INV_NUM") & ""
                            Dim INV_DATE As String = "00000000"
                            Dim INVO As String = ""
                            Dim INV_NUM_AS400 As String = INV_NUM
                            If INV_NUM_AS400 <> "0" Then
                                If INV_NUM_AS400.Length > 7 Then INV_NUM_AS400 = Mid(INV_NUM_AS400, INV_NUM_AS400.Length - 7 + 1, 7)
                                INV_DATE = Format(rowPOHDR.Item("PO_DATE_ORDERED"), "yyyyMMdd")
                                INVO = "INVO"
                            End If

                            swPOHDR.WriteLine(INV_NUM_AS400 _
                                              & sep & quo & "00" & PO_ORDER_NO & quo _
                                              & sep & INV_DATE _
                                              & sep & CStr(Val(Format(DATETIME_STAMP, "yyyyMMdd"))) _
                                              & sep & CStr(Val(Format(DATETIME_STAMP, "HHmmss"))) _
                                              & sep & quo & INVO & quo _
                                              & sep & CStr(Val(rowPOHDR.Item("PO_AMT_OPN") & "")) _
                                              & sep & quo & "USD" & quo _
                                              & sep & rowPOHDR.Item("VEND_SUPPLIER_ID") _
                                              & sep & quo & SpaceIfNull(Mid(rowPOHDR.Item("VEND_NAME") & "", 1, 30)) & quo _
                                              & sep & quo & SpaceIfNull(Mid(rowPOHDR.Item("VEND_ADDR1") & "", 1, 30)) & quo _
                                              & sep & quo & SpaceIfNull(Mid(rowPOHDR.Item("VEND_ADDR2") & "", 1, 30)) & quo _
                                              & sep & quo & SpaceIfNull(Mid(rowPOHDR.Item("VEND_CITY") & "", 1, 21)) & quo _
                                              & sep & quo & SpaceIfNull(rowPOHDR.Item("VEND_STATE") & "") & quo _
                                              & sep & quo & SpaceIfNull(Mid(rowPOHDR.Item("VEND_ZIP_CODE") & "", 1, 5)) & quo _
                                              & sep & quo & SpaceIfNull(rowPOHDR.Item("VEND_COUNTRY") & "") & quo
                                              )

                            Dim sqlw As String = "ISNULL(INV_NUM,'') = '" & INV_NUM & "' and PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                            For Each rowPODTL As DataRow In dst.Tables("PODTL").Select(sqlw, "PO_ORDER_LNO")
                                R += 1
                                swPODTL.WriteLine("D" _
                                                  & sep & INV_NUM_AS400 _
                                                  & sep & quo & "00" & rowPODTL.Item("PO_ORDER_NO") & quo _
                                                  & sep & "0" _
                                                  & sep & CStr(Val(rowPODTL.Item("PO_ORDER_LNO") & "")) _
                                                  & sep & quo & SpaceIfNull(rowPODTL.Item("ITEM_ALT_SORT") & "") & quo _
                                                  & sep & quo & SpaceIfNull(rowPODTL.Item("ITEM_CODE") & "") & quo _
                                                  & sep & SpaceIfNull(rowPODTL.Item("ITEM_EAN_CODE") & "") _
                                                  & sep & CStr(Val(rowPODTL.Item("PO_QTY_OPN") & "")) _
                                                  & sep & CStr(Val(rowPODTL.Item("PO_COST") & "")) _
                                                  & sep & Format(rowPODTL.Item("PO_DATE_REQUIRED"), "yyyyMMdd") _
                                                  & sep & quo & SpaceIfNull(rowPODTL.Item("PO_STATUS") & "") & quo _
                                                  & sep & "0" _
                                                  & sep & CStr(Val(Format(DATETIME_STAMP, "yyyyMMdd"))) _
                                                  & sep & CStr(Val(Format(DATETIME_STAMP, "HHmmss")))
                                                  )
                            Next
                        Next
                    End Using
                End Using

            Case "ADS"

                'Dim Folder As String = "\\Nymain-fs01\abs-files\INT\" & "XSDs\"
                'Dim Folder As String = "\\PROD-ABS-TS\abs-files\INT\" & "XSDs\"


                Dim Folder As String = ASCMAIN1.Folders("SharedRoot") & "XSDs\"

                If ASCMAIN1.Running_in_VS Then
                    Stop
                    Folder = "C:\SFTP\cfg\PROD\XSDs\"
                End If

                ' 11/24/2025 - Po Deletion request, PO changed and we need to send an updated PO
                ' ISSUE-7044 - balance POs/ASNs
                If G.APP_CMD = "CANC" Then
                    Dim xml As String = $"<RecordsToDelete>
                                <Record>
                                    <RecordType>PurchaseOrder</RecordType>
                                    <Key1>{G.APP_KEY}</Key1>
                                </Record>
                            </RecordsToDelete>"

                    FILENAME_TO_SEND = $"DeletionRequest_{XMIT_NO}.XML"
                    Using sw As New StreamWriter(FILENAME_TO_SEND)
                        sw.Write(xml)
                        sw.Close()
                        sw.Dispose()
                    End Using

                    R = 1
                    Exit Sub
                End If

                Dim XSD_name As String = "WWIMPZPOH"
                Dim dstPOH As New DataSet
                dstPOH.ReadXmlSchema(Folder & XSD_name & ".XSD")
                dstPOH.EnforceConstraints = False
                'dstPOH.ReadXml(Folder & "WWIMPZPOH_SAMPLE.XML")

                dstPOH.Tables("Line").Rows.Clear()
                Dim t As String = "PO"
                dstPOH.Tables(t).Rows.Clear()

                For Each rowPOHDR As DataRow In dst.Tables("POHDR").Select("", "INV_NUM,PO_ORDER_NO")

                    Dim PINV_NO As String = rowPOHDR.Item("PINV_NO") & ""
                    Dim PO_ORDER_NO As String = rowPOHDR.Item("PO_ORDER_NO")
                    Dim INV_DATE As Date = rowPOHDR.Item("PO_DATE_ORDERED") '  ALIASED IN SQL STMT rowPOHDR.Item("INV_DATE")
                    Dim INV_NUM As String = rowPOHDR.Item("INV_NUM") & ""

                    ' ISSUE-7044 - balance POs/ASNs
                    Dim PO_HDR_CTR_REV As String = ""
                    If dst.Tables("POHDR").Columns.Contains("PO_HDR_CTR_REV") Then
                        PO_HDR_CTR_REV = rowPOHDR.Item("PO_HDR_CTR_REV") & ""
                    End If

                    Dim rowPOH As DataRow = dstPOH.Tables(t).NewRow
                    With rowPOH
                        '.Item("Supplier") = row.Item("VEND_CODE")
                        .Item("Supplier") = "MAIN"
                        .Item("ExpectedReceiptDate") = INV_DATE.AddDays(30)
                        .Item("PurchaseOrderNumber") = INV_NUM
                        .Item("InternalReference") = rowPOHDR.Item("PO_ORDER_NO")
                    End With
                    dstPOH.Tables(t).Rows.Add(rowPOH)
                    Dim PO_ID As Int64 = Val(rowPOH.Item("PO_ID") & "")

                    ' Change for ADS 07/16/2025
                    If PINV_NO.Length > 0 Then
                        ASCMAIN1.sql = $"Select ICTPINV2.*, POTORDR2.PO_DATE_REQUIRED
                                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                                        from POTORDR2,ICTITEM1,ICTPINV2
                                        where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE
                                        and POTORDR2.PO_ORDER_NO = ICTPINV2.PO_ORDER_NO
                                        and POTORDR2.PO_ORDER_LNO = ICTPINV2.PO_ORDER_LNO
                                        AND ICTPINV2.PINV_NO = '{PINV_NO}'"
                    Else
                        ASCMAIN1.sql = $"Select POTORDR2.ITEM_CODE, POTORDR2.PO_QTY_OPN PINV_QTY, POTORDR2.PO_ORDER_LNO, POTORDR2.PO_ORDER_LNO PINV_LNO, POTORDR2.PO_DATE_REQUIRED
                                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                                        from POTORDR2, ICTITEM1 
                                        where ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE
                                        AND POTORDR2.PO_ORDER_NO = '{PO_ORDER_NO}' 
                                        AND POTORDR2.PO_QTY_OPN > 0"

                        If G.APP_CMD = "XFR_NO" Then
                            ASCMAIN1.sql = $"Select ICTIXFR2.ITEM_CODE, ICTIXFR2.XFR_QTY PINV_QTY, ICTIXFR2.XFR_LNO PO_ORDER_LNO, ICTIXFR2.XFR_LNO PINV_LNO, ICTIXFR1.XFR_DATE PO_DATE_REQUIRED
                                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                                        from ICTIXFR1, ICTIXFR2, ICTITEM1 
                                        where ICTIXFR1.XFR_NO = ICTIXFR2.XFR_NO
                                        AND ICTITEM1.ITEM_CODE = ICTIXFR2.ITEM_CODE
                                        AND ICTIXFR1.XFR_NO = '{G.APP_KEY}'"

                            rowPOH.Item("InternalReference") = INV_NUM
                        End If

                        If G.APP_CMD = "ASN" Then
                            ASCMAIN1.sql = $"Select POTASNM3.ITEM_CODE, POTASNM3.ASN_QTY PINV_QTY, POTASNM3.ASN_LNO PO_ORDER_LNO, POTASNM3.ASN_LNO PINV_LNO, TRUNC(SYSDATE) PO_DATE_REQUIRED
                                        , ICTITEM1.NRF_SIZE_CODE, ICTITEM1.SIZE_CODE, ICTITEM1.NRF_COLOR_CODE, ICTITEM1.COLOR_CODE
                                        from POTASNM3, ICTITEM1 
                                        where  POTASNM3.ITEM_CODE = ICTITEM1.ITEM_CODE (+)
                                        AND POTASNM3.ASN_XMIT_NO = '{G.APP_KEY}'"
                            rowPOH.Item("InternalReference") = INV_NUM
                        End If

                        rowPOH.Item("PurchaseOrderNumber") = PO_ORDER_NO
                        ' ISSUE-7044 - balance POs/ASNs
                        If PO_HDR_CTR_REV.Length > 0 Then
                            rowPOH.Item("PurchaseOrderNumber") &= PORetransmitChar & PO_HDR_CTR_REV
                        End If
                    End If

                    'Dim sqlw As String = "ISNULL(INV_NUM,'') = '" & INV_NUM & "' and PO_ORDER_NO = '" & PO_ORDER_NO & "'"
                    'For Each rowPODTL As DataRow In dst.Tables("PODTL").Select(sqlw, "PO_ORDER_LNO")
                    '    Stop
                    'Next

                    For Each row2 As DataRow In ASCDATA1.GetDataTable(ASCMAIN1.sql).Select("")
                        Dim ITEM_CODE As String = row2.Item("ITEM_CODE")
                        Dim PINV_QTY As Int32 = Val(row2.Item("PINV_QTY") & "")
                        Dim PINV_LNO As Int32 = Val(row2.Item("PINV_LNO") & "")
                        Dim PO_ORDER_LNO As Int32 = Val(row2.Item("PO_ORDER_LNO") & "")
                        Dim CustomerOrderRef As String = CStr(PO_ORDER_LNO)

                        If row2.Item("PO_DATE_REQUIRED") & "" <> "" Then
                            If rowPOH.Item("ExpectedReceiptDate") & "" = "" Then
                                rowPOH.Item("ExpectedReceiptDate") = row2.Item("PO_DATE_REQUIRED")
                            End If
                        End If

                        Dim rowLine As DataRow = dstPOH.Tables("Line").NewRow
                        With rowLine
                            .Item("StockNumber") = ITEM_CODE
                            .Item("Quantity") = PINV_QTY
                            .Item("CustomerOrderRef") = CustomerOrderRef
                            .Item("EDISize") = row2.Item("NRF_SIZE_CODE")
                            .Item("EDIColor") = row2.Item("NRF_COLOR_CODE")
                            .Item("PO_ID") = PO_ID
                        End With
                        dstPOH.Tables("Line").Rows.Add(rowLine)

                        R += 1
                    Next
                Next

                If R > 0 Then
                    FILENAME_TO_SEND = $"{XSD_name}_{XMIT_NO}.XML"
                    dstPOH.WriteXml(FILENAME_TO_SEND)
                End If

        End Select

    End Sub

    Overrides Sub Update_Archive()

        ' Added 09/26/2025 to prevent sending test environment data
        If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                MessageBox.Show("You are not in production. File transfer avoided.", "Update Archive", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        MyBase.Update_Archive()

        Dim subDir As String = DateTime.Now.ToString("yyyyMM")
        ' If the Archive direcory does not exist then create it as a courtesy
        If Not My.Computer.FileSystem.DirectoryExists(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir) Then
            My.Computer.FileSystem.CreateDirectory(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir)
        End If

        Select Case LP_CODE

            Case "CLA"

                If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
                    ' do not send
                Else
                    My.Computer.FileSystem.CopyFile(POHDR, sftp_folder & POHDR)
                    My.Computer.FileSystem.CopyFile(PODTL, sftp_folder & PODTL)
                End If

                My.Computer.FileSystem.MoveFile(POHDR, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & POHDR)
                My.Computer.FileSystem.MoveFile(PODTL, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & PODTL)

            Case "ADS"

                If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
                    ' do not send
                Else
                    sftp_put("ADS", True, FILENAME_TO_SEND, FILENAME_TO_SEND)
                End If

                My.Computer.FileSystem.MoveFile(FILENAME_TO_SEND, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & FILENAME_TO_SEND)

                ' Do this because the process sends back a Transmit No when a file is NOT sent.
                If G.APP_CMD = "XFR_NO" Then
                    If G.APP_KEY.Length > 0 Then
                        If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & FILENAME_TO_SEND) Then
                            Dim XFR_NOTE As String = "X" & Val(G.APP_KEY).ToString.PadLeft(5, "0")
                            ASCDATA1.ExecuteSQL("UPDATE ICTIXFR1 SET XFR_NOTE = :PARM1 WHERE XFR_NO = :PARM2", "VV", {XFR_NOTE, G.APP_KEY})
                        End If
                    End If
                ElseIf G.APP_CMD = "ASN" Then
                    If G.APP_KEY.Length > 0 Then
                        If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & FILENAME_TO_SEND) Then
                            ASCDATA1.ExecuteSQL("UPDATE POTASNM2 SET TRANSMIT_OPER = :PARM1 WHERE ASN_XMIT_NO = :PARM2", "VV", {ASCMAIN1.USER_ID, G.APP_KEY})
                        End If

                    End If
                End If
        End Select

    End Sub

    Sub Create_Work_Table()

    End Sub
End Class