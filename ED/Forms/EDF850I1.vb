Imports Infragistics.Win.UltraWinGrid

Public Class EDF850I1

    ' Temporary Oacle Tables
    Dim ROWSOTSVIA3
    Dim EDT850T1 As String
    Dim SOTORDR1 As String
    Dim SOTORDR2 As String

    ' Datarow

    Dim rowEDTTRPM1 As DataRow
    Dim rowEDT850T1 As DataRow
    Dim rowEDT850T2 As DataRow
    Dim rowEDTSLSP1 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim rowARTCUST1bt As DataRow
    Dim rowARTCUST2 As DataRow
    Dim rowARTCUST2_DC As DataRow
    Dim rowICTITEM1 As DataRow
    Dim rowSOTPCLS1 As DataRow

    ' EDI Control

    Dim EDI_ERRORs As New Dictionary(Of String, Int32)
    Dim EDI_DOC_SEQ_NO_ok As Boolean = False
    Dim EDI_DOC_SEQ_NO As String
    Dim EDI_DTL_SEQ As Int32
    Dim EDI_TP_QUAL As String
    Dim EDI_TP_ID As String

    Dim EDI_ACTION As String
    Dim EDI_REPLACED_VALUE As String

    Dim EDI_DOC_SEQ_NOs_no_company As New List(Of String)

    Dim multi_customer As Boolean
    Dim CUST_CODE_multi As String
    Dim CUST_CODEs As New Dictionary(Of String, List(Of String))

    ' Order Data

    Dim ORDR_CUST_PO As String
    Dim CUST_CODE As String
    Dim CUST_STORE_NO As String

    Dim WHSE_CODE As String
    Dim WHSE_CODE_customer As String

    Dim CUST_BILL_TO_CUST As String
    Dim ORDR_SHIP_TO As String
    Dim TERM_CODE As String
    Dim CUST_DC_NO As String
    Dim ORDR_DEPT As String
    Dim ORDR_EDI_810 As String
    Dim ORDR_EDI_856 As String
    Dim ORDR_DATE As Date
    Dim ORDR_CANCEL_DATE As Date
    Dim ORDR_SHIP_DATE As Date
    Dim ORDR_ARRIVAL_DATE As Date
    Dim ORDR_LAST_ARRIVAL_DATE As Date

    Dim PRICE_BASIS As String
    Dim PRICE_BASE_DPCT As Decimal
    Dim PRICE_LIST_CODE As String

    Dim CUST_CODE_OVERRIDE As String
    Dim PRICE_LIST_CODE_OVERRIDE As String

    Dim CURR_CODE As String = "USD"
    Dim CURR_EXCH_RATE As Decimal = 1

    Dim ITEM_CODE As String
    Dim ITEM_INACTIVE As String
    Dim ASST_ITEM_CODE As String
    Dim saleable_item As Boolean

    Dim TRANSIT_BUS_DAYS As Int32
    Dim CUST_EDI_DTS_FLAG As String
    Dim FRT_TERMS As String

    Dim ITEM_OK_IF_ONE_MATCH As Boolean = True
    Dim skip_store As Boolean
    Dim Bad_Data_Cond_05 As Boolean
    Dim skip_item As Boolean
    Dim skipped_items As New List(Of String)

    Dim sdqs_are_present As Boolean

    Dim RESOLUTIONS As New Dictionary(Of String, String)

    Dim ACTIONS_A As New List(Of String)
    Dim ACTIONS_E As New List(Of String)
    Dim ACTIONS_S As New List(Of String)
    Dim ACTIONS_R As New List(Of String)

    '01 N TRADING PARTNER SETUP REQUIRED 
    '27 A CHECK OTHER UNPROCESSED EDI ORDERS FOR SAME TP ID AND SAME CUSTOMER PO
    '28 A CHECK THIS PO FOR DUPLICATE ITEMS IN ORDER DETAIL
    '04 N CUSTOMER MASTER SETUP REQUIRED
    '35 N CUSTOMER MASTER SETUP REQUIRED
    '36 N CUSTOMER MASTER SETUP REQUIRED
    '03 N CUSTOMER MASTER SETUP REQUIRED
    '91 N CUSTOMER MASTER SETUP REQUIRED
    '09 A ACCEPT EDI DATA, GET AN EXTENSION FROM CUSTOMER, AND CHANGE CANCEL DATE
    '08 N CUSTOMER MUST RETRANSMIT WITH SHIP DATE
    '21 N CUSTOMER MASTER SETUP REQUIRED
    '11 A MAP EDI TERMS TO A VALID AR TERMS CODE
    '12 N MAP EDI TERMS TO A VALID AR TERMS CODE
    '13 N CUSTOMER MASTER SETUP REQUIRED
    '14 A CHECK ORDER HISTORY FOR THIS PO


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("EDTPARM1")
        Get_PARM("GLTPARM1")

        Create_Temp_Table()

        With dst
            ASCMAIN1.sql = "Select EDT850T1.*, ARTCUST1.CUST_NAME" & vbCrLf _
                & " from " & EDT850T1 & " EDT850T1,ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = EDT850T1.CUST_CODE"
            Create_TDA(.Tables.Add, "EDT850T1", "**", 0, True, "", 1)
            .Tables("EDT850T1").Columns.Add("SEL")
            .Tables("EDT850T1").Columns("SEL").DefaultValue = "0"
            .Tables("EDT850T1").Columns.Add("RESULT")
            .Tables("EDT850T1").Columns.Add("ORDERS", GetType(System.Int32))
            .Tables("EDT850T1").Columns.Add("AMOUNT", GetType(System.Decimal))
            '.Tables("EDT850T1").Columns.Add("CUST_CODE_OVERRIDE")
            '.Tables("EDT850T1").Columns.Add("PRICE_LIST_CODE_OVERRIDE")

            Dim TBL As DataTable = dst.Tables("EDT850T1").Clone
            TBL.TableName = "EDT850T1_IMPORTED"
            dst.Tables.Add(TBL)

            ASCMAIN1.sql = "Select EDT850T1.*, ARTCUST1.CUST_NAME" & vbCrLf _
              & " from EDT850T1,ARTCUST1" & vbCrLf _
              & " where ARTCUST1.CUST_CODE (+) = EDT850T1.CUST_CODE" & vbCrLf _
              & "   and EDT850T1.EDI_PROCESS_IND in ('1','C')" & vbCrLf _
              & "   and EDT850T1.EDI_PO_DATE between :PARM1 and :PARM2" & vbCrLf _
              & "   and EDT850T1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" & vbCrLf _
              & "   and (EDT850T1.CUST_CODE = :PARM3 or :PARM3 IS NULL)"

            Create_TDA(.Tables.Add, "EDT850T1_ARCHIVED", "**", 0, False, "DDV", 1)
            .Tables("EDT850T1_ARCHIVED").Columns.Add("SEL")
            .Tables("EDT850T1_ARCHIVED").Columns("SEL").DefaultValue = "0"
            .Tables("EDT850T1_ARCHIVED").Columns.Add("RESULT")
            .Tables("EDT850T1_ARCHIVED").Columns.Add("ORDERS", GetType(System.Int32))
            .Tables("EDT850T1_ARCHIVED").Columns.Add("AMOUNT", GetType(System.Decimal))
            '.Tables("EDT850T1_ARCHIVED").Columns.Add("CUST_CODE_OVERRIDE")
            '.Tables("EDT850T1_ARCHIVED").Columns.Add("PRICE_LIST_CODE_OVERRIDE")

            'TBL = dst.Tables("EDT850T1").Clone
            'TBL.TableName = "EDT850T1_ARCHIVED"
            'dst.Tables.Add(TBL)

            Create_TDA(.Tables.Add, "EDT850T2", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T3", "*", 1, False, "", 3)
            Create_TDA(.Tables.Add, "EDT850T4", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T5", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T6", "*", 1, False, "", 3)
            Create_TDA(.Tables.Add, "EDT850T7", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "EDT850T8", "*", 1, False, "", 3)
            Create_TDA(.Tables.Add, "EDT850T9", "*", 1, False, "", 2)

            Create_TDA(.Tables.Add, "EDT850TE", "*", 0, True, "", 2)
            .Tables("EDT850TE").Columns.Add("RESOLUTION")

            ASCMAIN1.sql = "Select * from " & SOTORDR1
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0)
            ASCMAIN1.sql = "Select * from " & SOTORDR2
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0)
            .Tables("SOTORDR2").Columns.Add("ORDR_AMT", GetType(System.Decimal), "ORDR_QTY * ORDR_UNIT_PRICE")

            Create_TDA(.Tables.Add, "SOTORDR4", "*")

            Create_TDA(.Tables.Add, "ARTCUST2", "*", 1, True, "", 2)

            Create_TDA(.Tables.Add, "SOTCITM1", "*", 1, False, "", 2)
            Create_TDA(.Tables.Add, "ICTITEM1", "*", 0, False)

            'ASCMAIN1.sql = "" _
            '    & "Select 'UPC' UPC_EAN, ITEM_UPC_CODE, ITEM_CODE from ICTITEM1 where ITEM_UPC_CODE is not Null" _
            '    & " union " _
            '    & "Select 'EAN' UPC_EAN, ITEM_EAN_CODE ITEM_UPC_CODE, ITEM_CODE from ICTITEM1 where ITEM_EAN_CODE is not Null"
            'Create_TDA(.Tables.Add, "ICTITEMX", "**", 0, False, "", 2)

            ASCMAIN1.sql = "" _
                & "Select ITEM_UPC_CODE, ITEM_CODE from ICTITEM1 where ITEM_UPC_CODE is not Null" _
                & " union " _
                & "Select ITEM_EAN_CODE ITEM_UPC_CODE, ITEM_CODE from ICTITEM1 where ITEM_EAN_CODE is not Null"
            ASCMAIN1.sql = "Select ITEM_UPC_CODE, MAX (ITEM_CODE) ITEM_CODE from (" & ASCMAIN1.sql & ") X group by ITEM_UPC_CODE"
            Create_TDA(.Tables.Add, "ICTITEMX", "**", 0, False, "", 1)


            For Each TABLE_NAME As String In New String() {"SOTPRIC1", "SOTPRIC2"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                Create_TDA(dst.Tables.Add, TABLE_NAME, "**", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            With .Tables.Add("EDT850TS")
                .Columns.Add("ST_KEY")
                .Columns.Add("EDI_STORE")
                .Columns.Add("EDI_SHIP_DC")
                .Columns.Add("EDI_ADDR_CODE")
                .Columns.Add("CUST_STORE_NO")
                .Columns.Add("CUST_DC_NO")
                .PrimaryKey = New DataColumn() { .Columns("ST_KEY")}
            End With

            ASCMAIN1.sql = "Select EDT850T2.EDI_DOC_SEQ_NO, EDT850T2.EDI_DTL_SEQ" & vbCrLf _
                & ", SOTORDR2.ITEM_CODE, EDT850T2.EDI_UPC, SOTPCLS1.PRICE_BASE_DPCT" & vbCrLf _
                & ", EDT850T2.EDI_PRICE PRICE, EDT850T2.EDI_PRICE" & vbCrLf _
                & ", EDT850T2.EDI_PRICE EDI_PRICE_CURR, EDT850T2.EDI_PRICE ITEM_PRICE_CURR" & vbCrLf _
                & " from SOTORDR2, SOTPCLS1, EDT850T2"
            Create_TDA(.Tables.Add, "EDTSDQT0", "**", 0, False, "", 4)

            ASCMAIN1.sql = "Select EDT850T1.EDI_DOC_SEQ_NO, EDT850T1.EDI_STORE" & vbCrLf _
                & ", EDT850T2.EDI_DTL_SEQ, ICTITEM1.ITEM_CODE, EDT850T2.EDI_ITEM SLN_PARENT_ITEM_CODE" & vbCrLf _
                & ", EDT850T2.EDI_ITEM, EDT850T2.EDI_UPC, EDT850T2.EDI_SKU, EDT850T2.EDI_SIZE_DESC" & vbCrLf _
                & ", 0 SLN_PARENT_ITEM_QTY, 0 SLN_PARENT_INNER_PACK_QTY, 0 EDI_SLN_SEQ" & vbCrLf _
                & ", 0 QTY, 0 ASST_PARENT_QTY, 0 ASST_PARENT_PRICE, EDT850T1.EDI_SHIP_DC, ICTITEM1.ITEM_CODE ASST_ITEM_CODE" & vbCrLf _
                & ", EDT850T2.EDI_ITEM SLN_PARENT_ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ICTITEM1.SALES_DIVISION_CODE, ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_DC_NO " & vbCrLf _
                & " from EDT850T2, EDT850T1, ICTITEM1, ICTCOLL1, ICTBRAN1, SOTSDIV1, ARTCUST2"
            Create_TDA(.Tables.Add, "EDTSDQT1", "**", 0, False, "", 4) ' 6)
            .Tables("EDTSDQT1").Columns("ASST_ITEM_CODE").AllowDBNull = True

            ASCMAIN1.sql = "Select SOTORDR1.CUST_CODE, EDT850T2.EDI_UPC, EDT850T2.EDI_SKU, EDT850T2.EDI_ITEM, EDT850T2.EDI_EAN" & vbCrLf _
                & ", SOTORDR2.ITEM_CODE, 0.01 CUST_PRICE, 0.01 ITEM_PRICE_CURR" & vbCrLf _
                & ", 0.01 PRICE_DISC, 0.01 RETAIL_PRICE, 0.01 RETAIL_PRICE_CURR" & vbCrLf _
                & ", SOTORDR1.COLLECTION_CODE, SOTORDR1.BRAND_CODE, SOTORDR1.SALES_DIVISION_CODE" & vbCrLf _
                & " from SOTORDR2, EDT850T2, SOTORDR1"
            Create_TDA(.Tables.Add, "EDTITEMX", "**", 0, False, "", 0)

            '        Call Create_Index("EDWITEMX", "I_EDWITEMX_1", "CUST_CODE,EDI_UPC,EDI_SKU,EDI_ITEM")
            '        Call Create_Index("EDWITEMX", "I_EDWITEMX_2", "CUST_CODE,ITEM_CODE")

            Create_TDA(.Tables.Add, "SOTSVIAX", "*", 0, False)
            Fill_Records("SOTSVIAX")

            For Each TABLE_NAME As String In New String() _
                {"SOTPCLS1", "EDTTRPM1", "TATTERM1", "SOTSVIA1", "SOTSVIA2", "SOTSVIA3", "EDTTERM1", "SOTSDIV1", "EDTUPCX1", "EDTTRPM3"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, TABLE_NAME = "EDTTERM1")
                Fill_Records(TABLE_NAME)
            Next

            ' ISSUE-7230 
            Create_TDA(.Tables.Add, "ICTWHSE1", "*")
            Fill_Records("ICTWHSE1", "", True, "SELECT * FROM ICTWHSE1 WHERE LP_CODE IS NOT NULL")

            ASCMAIN1.sql = "SELECT ICTWHSE2.*, ICTWHSE1.LP_CODE
                            FROM ICTWHSE2, ICTWHSE1
                            WHERE ICTWHSE2.WHSE_CODE = ICTWHSE1.WHSE_CODE
                            AND ICTWHSE1.LP_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTWHSE2", ASCMAIN1.sql, 0, False, "V")

        End With

        Fill_Records("ICTITEMX")

        grdEDT850T2.DataSource = dst.Tables("EDT850T2")
        grdEDT850T3.DataSource = dst.Tables("EDT850T3")
        grdEDT850T4.DataSource = dst.Tables("EDT850T4")
        grdEDT850T5.DataSource = dst.Tables("EDT850T5")
        grdEDT850T6.DataSource = dst.Tables("EDT850T6")
        grdEDT850TE.DataSource = dst.Tables("EDT850TE")
        grdEDT850T1.DataSource = dst.Tables("EDT850T1")
        grdEDTTERM1.DataSource = dst.Tables("EDTTERM1")

        Create_Summary(grdEDT850T1, "CUST_CODE", "Count")
        Create_Summary(grdEDT850T1, New String() {"SEL", "ORDERS", "AMOUNT"})

        Create_Summary(grdEDT850TE, "EDI_COND_DESC", "Count")

        grdEDT850T1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdEDT850T1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        grdEDT850T1.DisplayLayout.UseFixedHeaders = True
        With grdEDT850T1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White

                If gcol.Key = "SEL" Or gcol.Key = "CUST_CODE_OVERRIDE" Or gcol.Key = "PRICE_LIST_CODE_OVERRIDE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If

            Next

            For Each COLUMN_NAME As String In New String() {"SEL", "RESULT", "CUST_CODE", "CUST_NAME", "EDI_DEPARTMENT", "EDI_PO_NO", "EDI_PO_DATE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
            'For Each COLUMN_NAME As String In New String() {"EDI_DOC_SEQ_NO", "CUST_CODE", "ORDR_CUST_PO", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            '    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Violet
            'Next
            For Each COLUMN_NAME As String In New String() {"EDI_PO_NO", "EDI_START_DATE", "EDI_END_DATE", "EDI_SHIP_DATE", "EDI_SHIP_DC", "EDI_STORE", "EDI_PO_DATE", "EDI_ARRIVAL_DATE", "EDI_LAST_ARRIVAL_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            Next
            For Each COLUMN_NAME As String In New String() {"EDI_PO_PURP", "EDI_PO_TYPE", "EDI_ORD_QTY"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
            Next
            For Each COLUMN_NAME As String In New String() {"EDI_TERMS", "EDI_TERM_TYPE", "EDI_TERM_BASIS", "EDI_TERM_RATE", "EDI_TERM_DSCDAYS", "EDI_TERM_NETDAYS", "EDI_TERM_DESC", "EDI_TERM_DOM"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
            Next
            For Each COLUMN_NAME As String In New String() {"EDI_DOC_SEQ_NO", "GEN_DOC_NO", "EDI_ISA_NO", "EDI_TP_QUAL", "EDI_TP_ID", "EDI_RECEIVED_DATE"}
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
            Next
        End With

        Show_Filter(grdEDT850T1, True)
        grdEDT850T1.DisplayLayout.GroupByBox.Hidden = False


        With grdEDT850TE.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor = Drawing.Color.White

                If gcol.Key = "EDI_ACTION" Or gcol.Key = "EDI_REPLACED_VALUE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "ORDR_CUST_PO", "CUST_STORE_NO", "EDI_COND_CODE", "EDI_COND_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next

        End With


        Setup_RESOLUTIONS()

        ASCMAIN1.Add_Value_List(grdEDT850TE, "EDI_ACTION", Nothing, New String() {":", "N:Do Nothing", "S:Skip Record", "E:Use Value Expected", "A:Use Value Received", "R:Use Replacement Value"})
        ' ASCMAIN1.Add_Value_List(grdEDT850T1, "ORDR_PICK_TYPE", Nothing, New String() {":", "P:Pick&Pack", "C:Full Case"})
        'ASCMAIN1.Add_Value_List(grdEDT850T1, "WO_TYPE", "Select WO_TYPE, WO_TYPE_DESC from SOTWORKT")
        ASCMAIN1.Add_Value_List(Me.grdEDT850T1, "EDI_PROCESS_IND", Nothing, New String() {":", "0:Not Imported", "1:Imported", "C:Deleted", "T:Tests", "X:Voided"})


        dteFrom.Value = Now.Date.AddDays(-30)
        dteTo.Value = Now.Date

        If ASCMAIN1.CLIENT = "INT" Then
        Else
            TAC.TACMAIN1.Update_Forex()
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Refresh"


            Case "Process Orders"
                Dim SEL_CNT As Int32 = dst.Tables("EDT850T1").Select("SEL='1'").Length
                If SEL_CNT = 0 Then
                    EMsg &= vbCr & "You Must Select at Least 1 Order to Process"
                End If

                ' Added 02/21/2023
                If optACTION.Value = "I" Then
                    If dst.Tables("EDT850T1").Select("SEL='1' and ISNULL(CUST_CODE, '') = ''").Length > 0 Then
                        EMsg &= vbCr & "Some of the POs selected do not have an assigned Customer - import not permitted"
                    End If
                End If

                If dst.Tables("EDT850T1").Select("SEL='1' and EDI_PO_TYPE = 'TR'").Length <> 0 Then
                    If optACTION.Value = "I" Then
                        EMsg &= vbCr & "Some of the POs selected are Type TR (Terminate) - import not permitted"
                    End If
                End If

                If EMsg = "" Then
                    Select Case optACTION.Value
                        Case "I"
                            If MsgBox("This action will Import the " & CStr(SEL_CNT) & " Order(s) Selected." _
                                  & vbCrLf & vbCrLf & "OK to Continue?",
                                  MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                                Exit Sub
                            End If

                        Case Else
                            ' ISSUE-7241 - type DELETE in da box
                            Dim uMsg As String = $"This action will Delete the {CStr(SEL_CNT)} Order(s) Selected."
                            uMsg &= Environment.NewLine & Environment.NewLine
                            uMsg &= "Please type 'Delete' to continue."

                            Dim response As String = InputBox(uMsg, "Delete")
                            If response.ToUpper <> "DELETE" Then
                                Exit Sub
                            End If
                    End Select
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
            Case "Refresh"
                Load_EDT850T1()

            Case "Process Orders"
                Mode_Settings(True)
                If optACTION.Value = "D" Then
                    Delete_Orders()
                    optACTION.Value = "I"
                Else
                    Import_Orders()
                End If

                Mode_Settings(False)
                If grdEDT850T1.ActiveRow Is Nothing And grdEDT850T1.Rows.Count > 0 Then
                    grdEDT850T1.Rows(0).Activate()
                End If
                'Load_EDT850T1()
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Refresh").Settings.Enabled = not_iScreenMode
                    .Items("Process Orders").Settings.Enabled = not_iScreenMode
                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        ' grdEDT850T1.Visible = Not tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
                {"EDT850T1", "EDT850T2", "EDT850T3", "EDT850T4", "EDT850T5", "EDT850T6", "EDT850T7", "EDT850T8", "EDT850T9",
                 "SOTORDR1", "SOTORDR2", "ARTCUST2", "ICTITEM1", "SOTSVIAX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        tabMain.SelectedTab = tabMain.Tabs("Orders Waiting to be Imported")
        Setup_tabMain()
        Load_EDT850T1()
    End Sub

    Sub Load_Record()
        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        CommitTrans("Delete")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

    Sub Print_Record()
        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Generate_Report("PORWREC2")
        'Print_Report_End()
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "USER_ID"
                sql_where = ""
            Case "CUST_CODE"
                sql_where = "CUST_CODE in (Select Distinct CUST_CODE from EDT850T1)"
        End Select
    End Sub

    Public Overrides Function Events_Context() As Events_Entity

        Dim E As New Events_Entity

        E.TABLE_NAME = "EDT850T1"
        E.TABLE_KEY_CAPTION = "EDI Document 850"
        If grdEDT850T1.ActiveRow IsNot Nothing AndAlso grdEDT850T1.ActiveRow.IsDataRow Then ' If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = grdEDT850T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Text
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

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT850T1, "SSSBBBBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select Today's", "Select All", "De-Select All", "Select All for Customer", "Customer Order Inquiry", "Copy Value", "Paste Value", "Generate 855 as Rejection", "Reset Customer Code", "Export to Pivot")
        Load_Popup_Menu(grdEDT850T5, "BB", "Add as Store", "Add as DC")
        Load_Popup_Menu(grdEDT850T2, "B", "Item Status Inquiry")
        Load_Popup_Menu(grdEDT850TE, "SSSBBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Do Nothing", "Skip Record", "Use Value Expected", "Use Value Received", "Use Replacement Value", "Export to Pivot")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing

        If e.SourceControl.Name = "grdEDT850T1_EmbeddableTextBox" Then
            grd = grdEDT850T1
        Else
            Try
                grd = GRDs(Mid(e.SourceControl.Name, 4))
            Catch ex As Exception

            End Try
        End If

        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Select Case e.SourceControl.Name
            'Case "grdPOTORDRR"
            '    If EntryMode = "V" Then e.Cancel = True

        End Select

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name
            Case "grdEDT850T1"
                tlb_btn = DirectCast(tlb_pop.Tools("Customer Order Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow

                Dim CUST_CODE_855 As String = grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty
                Dim is855Customer As Boolean = dst.Tables("EDTTRPM1").Select($"CUST_CODE = '{CUST_CODE_855}' AND EDI_DOC_NO = '855'").Length > 0

                tlb_btn = DirectCast(tlb_pop.Tools("Generate 855 as Rejection"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = is855Customer AndAlso grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsDataRow

                tlb_btn = DirectCast(tlb_pop.Tools("Select All for Customer"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                    tlb_btn.SharedProps.Visible = True
                    tlb_btn.SharedProps.Caption = "Select All for " & CUST_CODE
                End If


                tlb_btn = DirectCast(tlb_pop.Tools("Copy Value"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = False
                Dim CNCV As String = tlb_btn.Tag
                tlb_btn.Tag = ""
                If grd.ActiveCell IsNot Nothing And grd.Selected.Rows.Count = 0 Then
                    Dim COLUMN_CAPTION As String = grd.ActiveCell.Column.Header.Caption
                    Dim COLUMN_NAME As String = grd.ActiveCell.Column.Key
                    Dim COLUMN_VALUE As String = grd.ActiveCell.Value & ""
                    If (COLUMN_NAME = "CUST_CODE_OVERRIDE" Or COLUMN_NAME = "PRICE_LIST_CODE_OVERRIDE") And COLUMN_VALUE <> "" Then
                        tlb_btn.Tag = COLUMN_CAPTION & ":" & COLUMN_NAME & ":" & COLUMN_VALUE
                        tlb_btn.SharedProps.Caption = "Copy " & COLUMN_CAPTION & " Value " & COLUMN_VALUE
                        tlb_btn.SharedProps.Visible = True
                    End If
                End If

                tlb_btn = DirectCast(tlb_pop.Tools("Paste Value"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = False
                If grd.Selected.Rows.Count <> 0 And CNCV <> "" Then
                    Dim COLUMN_CAPTION As String = Split(CNCV, ":")(0)
                    Dim COLUMN_NAME As String = Split(CNCV, ":")(1)
                    Dim COLUMN_VALUE As String = Split(CNCV, ":")(2)
                    If (COLUMN_NAME = "CUST_CODE_OVERRIDE" Or COLUMN_NAME = "PRICE_LIST_CODE_OVERRIDE") Then
                        tlb_btn.Tag = COLUMN_CAPTION & ":" & COLUMN_NAME & ":" & COLUMN_VALUE
                        tlb_btn.SharedProps.Caption = "Paste " & COLUMN_CAPTION & " Value " & COLUMN_VALUE
                        tlb_btn.SharedProps.Visible = True
                    End If
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

            Case "Generate 855 as Rejection"
                Try
                    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty
                    Dim is855Customer As Boolean = dst.Tables("EDTTRPM1").Select($"CUST_CODE = '{CUST_CODE}' AND EDI_DOC_NO = '855'").Length > 0

                    If Not is855Customer Then
                        MessageBox.Show($"Customer {CUST_CODE} is not an EDI 855 customer.", "Generate 855", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit Sub
                    End If

                    Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty
                    Dim numRecs As Int32 = 0

                    If EDI_DOC_SEQ_NO.Length > 0 Then
                        TAC.EDC855O1.Generate_855_Cancelling_850(clsASCBASE1, EDI_DOC_SEQ_NO)
                        numRecs = dst.Tables("EDT855O1").Rows.Count
                    End If

                    ' INC-5864 
                    Dim lstOrdrCustPo As List(Of String) = dst.Tables("EDT855O1").AsEnumerable().Select(Function(x) x.Item("ORDR_CUST_PO").ToString).ToList
                    MessageBox.Show($"{numRecs} EDI 855 records generated as Rejection.{Environment.NewLine}PO(s){Environment.NewLine}{String.Join(Environment.NewLine, lstOrdrCustPo.ToArray)}", "Rejection", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Catch ex As Exception
                    MessageBox.Show($"Error Generating EDI 855s {ex.Message}", "", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Try


            Case "Select Today's"
                Dim sqlw As String = "EDI_RECEIVED_DATE > '" & Format(DATETIME_STAMP, "MM/dd/yy") & "'"
                For Each row As DataRow In dst.Tables("EDT850T1").Select(sqlw)
                    row.Item("SEL") = "1"
                Next

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdEDT850T1.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next

            Case "Do Nothing", "Skip Record", "Use Value Expected", "Use Value Received", "Use Replacement Value"
                If grdEDT850TE.ActiveRow IsNot Nothing AndAlso Not grdEDT850TE.ActiveRow.Selected Then
                    grdEDT850TE.ActiveRow.Selected = True
                End If
                For Each grow As UltraWinGrid.UltraGridRow In grdEDT850TE.Selected.Rows
                    Dim EDI_ACTION As String = "N"
                    If e.Tool.Key = "Use Value Received" Then
                        EDI_ACTION = "A"
                    ElseIf e.Tool.Key = "Use Value Expected" Then
                        EDI_ACTION = "E"
                    ElseIf e.Tool.Key = "Use Replacement Value" Then
                        EDI_ACTION = "R"
                    ElseIf e.Tool.Key = "Skip Record" Then
                        EDI_ACTION = "S"
                    End If
                    grow.Cells("EDI_ACTION").Value = EDI_ACTION
                    grow.Update()
                Next

            Case "Paste Value"

                tlb_btn = DirectCast(e.Tool, UltraWinToolbars.ButtonTool)
                Dim CNCV As String = tlb_btn.Tag

                If grd.Selected.Rows.Count <> 0 And CNCV <> "" Then
                    Dim COLUMN_CAPTION As String = Split(CNCV, ":")(0)
                    Dim COLUMN_NAME As String = Split(CNCV, ":")(1)
                    Dim COLUMN_VALUE As String = Split(CNCV, ":")(2)

                    For Each grow As UltraWinGrid.UltraGridRow In grdEDT850T1.Selected.Rows
                        grow.Cells(COLUMN_NAME).Value = COLUMN_VALUE
                        grow.Update()
                    Next
                    tlb_btn.Tag = ""
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            ' Added 02/21/2023
            Case "Reset Customer Code"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & String.Empty
                Dim EDI_PO_NO As String = grd.ActiveRow.Cells("EDI_PO_NO").Value & String.Empty
                Dim EDI_DOC_SEQ_NO As String = grd.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value & String.Empty

                Dim zMsg As String = $"Do you want to reset the Customer Code for EDI Doc No {EDI_DOC_SEQ_NO}, Customer {CUST_CODE}, PO {EDI_PO_NO}?"
                If MessageBox.Show(zMsg, "Reset Customer Code", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = DialogResult.No Then
                    Exit Sub
                End If

                Try
                    BeginTrans()
                    'ASCMAIN1.sql = "Update EDT850T1 Set EDI_PROCESS_IND = null, CUST_CODE = null where EDI_PO_NO = :PARM1 and EDI_PROCESS_IND = '0'"
                    'ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {EDI_PO_NO})

                    ASCMAIN1.sql = $"Update EDT850T1 Set EDI_PROCESS_IND = null, LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}', CUST_CODE = null where EDI_DOC_SEQ_NO = :PARM1 and EDI_PROCESS_IND = '0'"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New Object() {EDI_DOC_SEQ_NO})

                    DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                    Write_Event_Log("EDT850T1", EDI_DOC_SEQ_NO, "Reset Customer Code")

                    grd.ActiveRow.Cells("CUST_CODE").Value = String.Empty
                    grd.ActiveRow.Cells("CUST_NAME").Value = String.Empty
                    grd.ActiveRow.Cells("EDI_PROCESS_IND").Value = String.Empty
                    grd.ActiveRow.Cells("SEL").Value = "0"

                    CommitTrans("Update successful. Please refresh the data.")

                Catch ex As Exception
                    Rollback(ex.Message)
                End Try

            Case "Select All for Customer"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                Dim sqlw As String = "CUST_CODE = '" & CUST_CODE & "'"
                Dim tbl As DataTable = DirectCast(grdEDT850T1.DataSource, DataTable)

                For Each row As DataRow In tbl.Select(sqlw)
                    row.Item("SEL") = "1"
                Next

            Case "Customer Order Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value
                Context_Launch("Select", CUST_CODE, e.Tool.Key, "SOFCORD1")

            Case "Item Status Inquiry"
                Dim ITEM_UPC_CODE As String = grd.ActiveRow.Cells("EDI_UPC").Value
                Dim row As DataRow = dst.Tables("ICTITEMX").Rows.Find(ITEM_UPC_CODE)
                If row IsNot Nothing Then
                    Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                Else
                    MsgBox("Cannot Locate Item for UPC " & ITEM_UPC_CODE, MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                End If

            Case "Add as Store", "Add as DC"

                Dim CUST_CODE As String = grdEDT850T1.ActiveRow.Cells("CUST_CODE").Value & ""
                If CUST_CODE = "" Then Exit Sub
                Dim EDI_ADDR_CODE_QUAL As String = grd.ActiveRow.Cells("EDI_ADDR_CODE_QUAL").Value & ""
                Dim EDI_ADDR_CODE As String = grd.ActiveRow.Cells("EDI_ADDR_CODE").Value & ""
                Dim CUST_STORE_NO As String = Format_Store(EDI_ADDR_CODE)

                Dim GLN As String = ""

                'If Len(CUST_STORE_NO) > 6 Then
                '    Dim CN As String = grd.ActiveRow.Cells("EDI_CUST_NAME_ADR").Value
                '    If CN.StartsWith("WAL-MART DC ") Then
                '        CUST_STORE_NO = Mid(CN, 13, 5)
                '        GLN = EDI_ADDR_CODE
                '    End If
                'End If

                If Len(CUST_STORE_NO) > 6 Then
                    Dim CUST_STORE_NO_entry As String = ASCMAIN1.Get_txt_from_User("Customer Address Code (" & CUST_STORE_NO & ") appears to be a GLN",
                                                                 "Enter Store No to use for this Address", , 6, "000000")
                    'MsgBox("Invalid Value for Customer Address Code (" & CUST_STORE_NO & ")", _
                    '       MsgBoxStyle.OkOnly, _
                    '       "Cannot Perform Requested Action")
                    If CUST_STORE_NO_entry = "" Then
                        Exit Sub
                    Else
                        GLN = EDI_ADDR_CODE
                        CUST_STORE_NO = CUST_STORE_NO_entry.PadLeft(6, "0")
                    End If

                End If

                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                If rowARTCUST2 IsNot Nothing Then
                    MsgBox("Address Record is alread on file for " & CUST_CODE & "-" & CUST_STORE_NO)
                Else
                    If ASCMAIN1.Logical_Lock("ARTCUST1", CUST_CODE, False, True, True, 1) Then
                        If MsgBox("OK to Add Address Record for " & CUST_CODE & "-" & CUST_STORE_NO,
                                  MsgBoxStyle.YesNo, "Verification to " & e.Tool.Key) = MsgBoxResult.Yes Then
                            dst.Tables("ARTCUST2").Rows.Clear()
                            rowARTCUST2 = dst.Tables("ARTCUST2").NewRow
                            With rowARTCUST2
                                .Item("CUST_CODE") = CUST_CODE
                                .Item("CUST_STORE_NO") = CUST_STORE_NO
                                .Item("CUST_STORE_NAME") = grd.ActiveRow.Cells("EDI_CUST_NAME_ADR").Value
                                .Item("CUST_STORE_ADDR1") = grd.ActiveRow.Cells("EDI_ADDRESS1").Value
                                .Item("CUST_STORE_ADDR2") = grd.ActiveRow.Cells("EDI_ADDRESS2").Value
                                .Item("CUST_STORE_CITY") = grd.ActiveRow.Cells("EDI_CITY").Value
                                .Item("CUST_STORE_STATE") = grd.ActiveRow.Cells("EDI_STATE").Value
                                .Item("CUST_STORE_ZIP_CODE") = grd.ActiveRow.Cells("EDI_ZIPCODE").Value
                                .Item("CUST_STORE_COUNTRY") = grd.ActiveRow.Cells("EDI_COUNTRY").Value
                                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                .Item("INIT_DATE") = DATETIME_STAMP
                                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                .Item("LAST_DATE") = DATETIME_STAMP
                                .Item("GLOBAL_LOCATION_NUMBER") = GLN
                                If e.Tool.Key = "Add as DC" Then
                                    .Item("CUST_DC_IND") = "1"
                                Else
                                    '  .Item("CUST_DC_NO") = CUST_DC_NO
                                End If
                            End With
                            dst.Tables("ARTCUST2").Rows.Add(rowARTCUST2)
                            Update_Record_TDA("ARTCUST2")
                            MsgBox("Address Record for " & CUST_CODE & "-" & CUST_STORE_NO & " has been added")
                        End If
                        ASCMAIN1.MultiTask_Release("", 0, 1)
                    End If
                End If

            Case "Export to Pivot"
                Dim dt As DataTable = DirectCast(grd.DataSource, DataTable)
                Dim pivotColumns() As String
                Dim includeTotals As Boolean = True
                Dim columnFormats As New Dictionary(Of String, (String, String))
                Dim columnTypes As String

                If grd.Name = "grdEDT850TE" Then
                    pivotColumns = {"CUST_CODE", "ORDR_CUST_PO", "EDI_COND_DESC"}
                    columnTypes = "AAA" 'all rowfields
                    includeTotals = False
                    For Each column As String In pivotColumns
                        If grd.DisplayLayout.Bands(0).Columns.Exists(column) Then
                            Dim gridColumn As Infragistics.Win.UltraWinGrid.UltraGridColumn = grd.DisplayLayout.Bands(0).Columns(column)
                            columnFormats(column) = (gridColumn.Header.Caption, gridColumn.Format)
                        End If
                    Next

                ElseIf grd.Name = "grdEDT850T1" Then
                    pivotColumns = {"CUST_CODE", "EDI_PO_NO", "EDI_START_DATE", "EDI_ARRIVAL_DATE", "EDI_END_DATE", "AMOUNT"}
                    columnTypes = "AAAAAB"
                    For Each column As String In pivotColumns
                        If grd.DisplayLayout.Bands(0).Columns.Exists(column) Then
                            Dim gridColumn As Infragistics.Win.UltraWinGrid.UltraGridColumn = grd.DisplayLayout.Bands(0).Columns(column)
                            columnFormats(column) = (gridColumn.Header.Caption, gridColumn.Format)
                        End If
                    Next
                Else
                    MsgBox("Grid not recognized for exporting to Pivot.")
                    Exit Sub
                End If

                If dt IsNot Nothing Then
                    TAC.TACMAIN1.Create_Pivot(Me, dt, pivotColumns, columnTypes, columnFormats, includeTotals)
                End If
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "PO_SHIPMENT_NO"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Call Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "PO_SHIPMENT_NO"
            '    Call Click_Command("View")
        End Select
    End Sub
#End Region

    Sub Load_EDT850T1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Data")

        Create_Temp_Table()
        Fill_Records("EDT850T1")

        ' Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT COUNT(*) NUM_RECORDS FROM EDT850T1 WHERE  COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' AND NVL(EDI_PROCESS_IND,'0') = '0' AND EDI_TP_ID ='6111350003' and EDI_PO_TYPE IN ('KC', 'KB','BK')")
        Dim tbl As DataTable = ASCDATA1.GetDataTable("SELECT COUNT(*) NUM_RECORDS FROM EDT850T1 WHERE  COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' AND NVL(EDI_PROCESS_IND,'0') = '0' AND EDI_TP_ID ='6111350003' and EDI_PO_TYPE IN ('KC', 'KB','BK')")

        Dim numRecords As Int16 = Val(tbl.Rows(0).Item("NUM_RECORDS") & String.Empty)
        If numRecords > 0 Then
            MessageBox.Show("There are " & numRecords & " transactions that need Acknowledgment", "850 Ack", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("EDT850TE"), New String() {"EDI_DOC_SEQ_NO"}).Rows
            Dim rowEDT850T1 As DataRow = dst.Tables("EDT850T1").Rows.Find(row.Item("EDI_DOC_SEQ_NO"))
            If rowEDT850T1 IsNot Nothing Then rowEDT850T1.Item("RESULT") = "Rejected"
        Next
        Sort_grdColumns(grdEDT850T1, "EDI_DOC_SEQ_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdEDT850T1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdEDT850T1.AfterRowActivate
        Setup_grdEDT850T1()
    End Sub

    Sub Setup_grdEDT850T1()
        If grdEDT850T1.ActiveRow Is Nothing OrElse Not grdEDT850T1.ActiveRow.IsDataRow Then
            tabEDT850T1.Visible = False
        Else
            tabEDT850T1.Visible = True
            EnforceConstraints(False)
            Dim EDI_DOC_SEQ_NO As String = grdEDT850T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value

            Fill_Records("EDT850T2", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T2, "EDI_DTL_SEQ")
            grdEDT850T2.Text = "Line Item Order Details for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T3", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T3, "EDI_DTL_SEQ")
            grdEDT850T3.Text = "Store/Qty Order Details for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T4", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T4, "EDI_CMT_SEQ")
            grdEDT850T4.Text = "Comments for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T5", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T5, "EDI_ADR_SEQ")
            grdEDT850T5.Text = "Address Details for " & EDI_DOC_SEQ_NO

            Fill_Records("EDT850T6", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT850T6, "EDI_DTL_SEQ")
            grdEDT850T6.Text = "Assortment Details for " & EDI_DOC_SEQ_NO

            If tabEDT850T1.SelectedTab.Key = "Raw EDI Data" Then
                txtRaw.Text = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO)
            End If

            EnforceConstraints(True)
        End If
    End Sub

    Sub Create_Temp_Table()

        ASCDATA1.ExecuteSQL("Update EDT850T1 Set EDI_PROCESS_IND = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_PROCESS_IND is Null")

        ASCMAIN1.sql = "Update EDT850T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
               & " where EDI_OUR_ID = TRIM(EDT850T1.EDI_OUR_ID) and EDI_TP_ID = TRIM(EDT850T1.EDI_TP_ID))" _
               & " where EDI_PROCESS_IND = '0' and COMPANY_CODE IS NULL"
        ASCDATA1.ExecuteSQL()

        ' 03/03/2025 - Permit using EDI_CHAIN
        ' INC 6407 Neiman2 Customer Override
        ASCMAIN1.sql = "Update EDT850T1 Set CUST_CODE = (Select CUST_CODE from EDTSLSP1
                    where EDI_QUAL_850 = EDT850T1.EDI_TP_QUAL and EDI_ID_850 = EDT850T1.EDI_TP_ID and EDI_CHAIN = EDT850T1.EDI_CHAIN)
                    where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Update EDT850T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM3" _
               & " where EDI_TP_QUAL = EDT850T1.EDI_TP_QUAL and EDI_TP_ID = EDT850T1.EDI_TP_ID and EDI_DOC_NO = '850' and EDI_STORE = (SELECT MIN(EDI_STORE_01) EDI_STORE FROM EDT850T3 WHERE EDI_DOC_SEQ_NO =  EDT850T1.EDI_DOC_SEQ_NO))" _
               & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT850T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM2" _
               & " where EDI_TP_QUAL = EDT850T1.EDI_TP_QUAL and EDI_TP_ID = EDT850T1.EDI_TP_ID and EDI_DOC_NO = '850' and EDI_DEPT_NO = EDT850T1.EDI_DEPARTMENT)" _
               & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ' FOR KSPOUTLETS
        ASCMAIN1.sql = "Update EDT850T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM3" _
               & " where EDI_TP_QUAL = EDT850T1.EDI_TP_QUAL and EDI_TP_ID = EDT850T1.EDI_TP_ID and EDI_DOC_NO = '850' and EDI_STORE = EDT850T1.EDI_SHIP_DC)" _
               & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT850T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
               & " where EDI_TP_QUAL = EDT850T1.EDI_TP_QUAL and EDI_TP_ID = EDT850T1.EDI_TP_ID and EDI_DOC_NO = '850')" _
               & " where EDI_PROCESS_IND = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from EDT850T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE is Null"
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

        If EDT850T1 = "" Then
            ASCMAIN1.sql = "Select * from EDT850T1 where ROWNUM < 1"
            EDT850T1 = ASCMAIN1.Temp_Table

            ASCMAIN1.sql = "Select * from SOTORDR1 where ROWNUM < 1"
            SOTORDR1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")

            ASCMAIN1.sql = "Select * from SOTORDR2 where ROWNUM < 1"
            SOTORDR2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR2 & " Add Primary Key (ORDR_NO,ORDR_LNO)")

        Else
            ASCMAIN1.sql = "Truncate Table " & EDT850T1
            ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Select * from EDT850T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' and EDI_PO_TYPE NOT IN ('KC', 'KB','BK')"
            ASCMAIN1.sql = "Select * from EDT850T1 where EDI_PROCESS_IND = '0' and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' and EDI_PO_TYPE NOT IN ('KC', 'KB')"
            ASCDATA1.ExecuteSQL("Insert into " & EDT850T1 & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = $"Select CUST_CODE, EDI_PO_TYPE, COUNT (*) EDI850S, MIN (EDI_DOC_SEQ_NO) MINEDI, MAX (EDI_DOC_SEQ_NO) MAXEDI
from EDT850T1
where EDI_PROCESS_IND = '0' and COMPANY_CODE = '{ASCMAIN1.DBS_COMPANY}'
and EDI_PO_TYPE IN ('KC', 'KB', 'BK')
group by CUST_CODE, EDI_PO_TYPE"
            Dim tbl2 As DataTable = ASCDATA1.GetDataTable
            If tbl2.Rows.Count > 0 Then
                Using frm As New ASFMSGBF
                    frm.Show_grd(tbl2, Me, "Warning: EDI 850 Records with Type KC, KB, or BK")
                End Using
            End If
        End If
    End Sub

    Sub Delete_Orders()

        BeginTrans()
        For Each rowEDT850T1 As DataRow In dst.Tables("EDT850T1").Select("SEL = '1'")
            Dim EDI_DOC_SEQ_NO As String = rowEDT850T1.Item("EDI_DOC_SEQ_NO")
            ASCMAIN1.sql = "Update EDT850T1" _
                & $" Set EDI_PROCESS_IND = 'C', LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}' where EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}'"
            ASCDATA1.ExecuteSQL()
            DATETIME_STAMP = Now + ASCMAIN1.NowTSD
            Write_Event_Log("EDT850T1", EDI_DOC_SEQ_NO, "Deleted")
        Next
        CommitTrans("Selected EDI Orders have been Deleted")

    End Sub

    Sub Asst_SDQT0()

    End Sub

    Sub Asst_Item_SDQ()

    End Sub

    Sub Import_Orders()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Importing EDI Orders")

        Fill_Records("ICTITEMX")

        ' Loop thru and Process/Cancel Selected EDI Orders

        Dim ORDERS_PROCESSED As Int32 = 0
        Dim ORDERS_IMPORTED As Int32 = 0
        Dim ORDERS_REJECTED As Int32 = 0

        Dim sqlw As String = ""
        Dim EDI_RECEIVED_VALUE As String = ""
        CUST_CODE = ""

        Dim EDI_ISA_NO As String = ""
        Dim EDI_JRNL_NO As String = "" ' THIS IS LIKE A ORDR_GROUP_NO FOR ISA'S
        Dim CUST_CODE_last As String = ""

        For Each rowEDT850T1 In dst.Tables("EDT850T1").Select("SEL = '1'", "CUST_CODE,EDI_ISA_NO,EDI_DOC_SEQ_NO")
            ORDERS_PROCESSED += 1
            sdqs_are_present = False

            skipped_items.Clear()

            EDI_ISA_NO = rowEDT850T1.Item("EDI_ISA_NO")
            EDI_JRNL_NO = rowEDT850T1.Item("EDI_JRNL_NO") & "" ' until I discuss this with Maria.  Also see commented out code below
            EDI_DOC_SEQ_NO = rowEDT850T1.Item("EDI_DOC_SEQ_NO")
            EDI_TP_QUAL = rowEDT850T1.Item("EDI_TP_QUAL") & ""
            EDI_TP_ID = rowEDT850T1.Item("EDI_TP_ID") & ""
            ORDR_CUST_PO = rowEDT850T1.Item("EDI_PO_NO") & ""
            ORDR_DEPT = rowEDT850T1.Item("EDI_DEPARTMENT") & ""
            ORDR_DATE = rowEDT850T1.Item("EDI_PO_DATE")

            CUST_CODE = rowEDT850T1.Item("CUST_CODE") & ""
            CUST_CODE_OVERRIDE = rowEDT850T1.Item("CUST_CODE_OVERRIDE") & ""
            PRICE_LIST_CODE_OVERRIDE = rowEDT850T1.Item("PRICE_LIST_CODE_OVERRIDE") & ""
            ' CUST_CODE_multi = CUST_CODE ' THERE ARE SOME ROUTINES THAT USE CUST_CODE_multi, THAT ARE CALLED BEFORE THE CUST_CODE_multi LOOP BELOW

            If CUST_CODE_OVERRIDE <> "" Then
                Dim rowARTCUST1_override As DataRow = LookUp("ARTCUST1", CUST_CODE_OVERRIDE)
                If rowARTCUST1_override Is Nothing Then
                    MsgBox("Invalid Value Specified for Overriding Customer " & CUST_CODE_OVERRIDE & " in EDI Document " & EDI_DOC_SEQ_NO,
                           MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Exit For
                End If
                CUST_CODE = CUST_CODE_OVERRIDE

            End If

            CUST_CODE_multi = CUST_CODE ' THERE ARE SOME ROUTINES THAT USE CUST_CODE_multi, THAT ARE CALLED BEFORE THE CUST_CODE_multi LOOP BELOW


            If PRICE_LIST_CODE_OVERRIDE <> "" Then
                Dim rowSOTPRIC1_override As DataRow = LookUp("SOTPRIC1", PRICE_LIST_CODE_OVERRIDE)
                If rowSOTPRIC1_override Is Nothing Then
                    MsgBox("Invalid Value Specified for Overriding Price List " & PRICE_LIST_CODE_OVERRIDE & " in EDI Document " & EDI_DOC_SEQ_NO,
                           MsgBoxStyle.OkOnly, "Cannot Proceed")
                    Exit For
                End If
            End If

            If CUST_CODE <> CUST_CODE_last Then
                CUST_CODE_last = CUST_CODE
                Fill_Records("ARTCUST2", CUST_CODE)
                Fill_Records("SOTCITM1", CUST_CODE)
            End If

            ASCMAIN1.Progress("-", EDI_DOC_SEQ_NO)

            Dim EDI_STORE As String = rowEDT850T1.Item("EDI_STORE") & ""

            CUST_STORE_NO = ""
            CUST_DC_NO = ""

            If EDI_STORE <> "" Then
                CUST_STORE_NO = Format_Store(EDI_STORE)
            End If

            'dst.Tables("EDT850TS").Rows.Clear()
            'dst.Tables("EDTITEMX").Rows.Clear()
            'dst.Tables("EDTSDQT1").Rows.Clear()
            'dst.Tables("EDTSDQT0").Rows.Clear()

            EDI_ERRORs.Clear()
            EDI_DOC_SEQ_NO_ok = True

            Dim EDI_ERROR_NO As Int32 = 0
            For Each row As DataRow In dst.Tables("EDT850TE").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'", "EDI_ERROR_NO")
                If row.Item("EDI_ACTION") = "N" Then
                    row.Delete()
                Else
                    EDI_ERROR_NO += 1
                    row.Item("EDI_ORDER_COUNT") = 0
                    row.Item("EDI_ERROR_NO") = EDI_ERROR_NO
                    Dim EDI_ERROR As String = row.Item("EDI_COND_DESC") & vbTab & row.Item("EDI_COND_CODE") & vbTab & row.Item("EDI_RECEIVED_VALUE")
                    EDI_ERRORs.Add(EDI_ERROR, EDI_ERROR_NO)
                End If
            Next

            rowEDTTRPM1 = Get_EDTTRPM1()

            sqlw = String.Format("EDI_DOC_SEQ_NO <> '{0}' and EDI_TP_QUAL = '{1}' and EDI_TP_ID = '{2}' and EDI_PO_NO = '{3}' and ISNULL(EDI_STORE,'') = '{4}'",
                                  EDI_DOC_SEQ_NO, EDI_TP_QUAL, EDI_TP_ID, ORDR_CUST_PO, EDI_STORE)

            If dst.Tables("EDT850T1").Select(sqlw).Length <> 0 Then
                '  If ASCMAIN1.Running_in_VS Then Stop
                Bad_Data(EDI_COND_DESC:="Possible PO Duplication within Import",
                            EDI_COND_CODE:="27",
                            EDI_RECEIVED_VALUE:=EDI_DOC_SEQ_NO)
            End If

            If EDI_DOC_SEQ_NO_ok Then

                Fill_Records("EDT850T2", EDI_DOC_SEQ_NO)

                ASCMAIN1.sql = "Select EDI_ITEM, EDI_UPC, EDI_SKU, EDI_EAN, EDI_SKU EDI_GTIN, EDI_LN_SHIP_DC, Count (*)" _
                    & " from EDT850T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
                    & " group by EDI_ITEM, EDI_UPC, EDI_SKU, EDI_EAN, EDI_GTIN, EDI_LN_SHIP_DC having Count (*) > 1"
                Dim TBL As DataTable = ASCDATA1.GetDataTable
                If TBL.Rows.Count = 0 Then
                    ' IF NO DUPLICATES, THEN CHECK AGAIN WITH JUST EDI_ITEM, EDI_UPC AND EDI_SKU
                    ASCMAIN1.sql = "Select EDI_ITEM, EDI_UPC, EDI_EAN, EDI_LN_SHIP_DC, Count (*)" _
                        & " from EDT850T2 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
                        & " group by EDI_ITEM, EDI_UPC, EDI_EAN, EDI_LN_SHIP_DC having Count (*) > 1"
                    TBL = ASCDATA1.GetDataTable
                End If
                For Each row As DataRow In TBL.Rows ' For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                    If row.Item("EDI_ITEM") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "Item: " & row.Item("EDI_ITEM")
                    ElseIf row.Item("EDI_UPC") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "UPC: " & row.Item("EDI_UPC")
                    ElseIf row.Item("EDI_SKU") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "SKU: " & row.Item("EDI_SKU")
                    ElseIf row.Item("EDI_EAN") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "EAN: " & row.Item("EDI_EAN")
                    ElseIf row.Item("EDI_GTIN") & "" <> "" Then
                        EDI_RECEIVED_VALUE = "GTIN: " & row.Item("EDI_GTIN")
                    End If
                    Bad_Data(EDI_COND_DESC:="Item Duplication within PO",
                            EDI_COND_CODE:="28",
                            EDI_RECEIVED_VALUE:=EDI_RECEIVED_VALUE)
                    If EDI_ACTION = "S" Then
                        skip_item = True
                    Else
                    End If
                Next

                Fill_Records("EDT850T3", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T4", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T5", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T6", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T7", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T8", EDI_DOC_SEQ_NO)
                Fill_Records("EDT850T9", EDI_DOC_SEQ_NO)

                If dst.Tables("EDT850T3").Rows.Count <> 0 Then
                    sdqs_are_present = True
                End If
            End If

            Dim EDI_SHIP_DC As String = ""
            Dim EDI_ADDR_CODE As String = ""

            If EDI_DOC_SEQ_NO_ok Then
                Get_ARTCUST1()

                Dim EDI_PO_PURP As String = rowEDT850T1.Item("EDI_PO_PURP") & ""
                If EDI_PO_PURP = "01" Then
                    Bad_Data(EDI_COND_DESC:="This is a Cancellation Order", EDI_COND_CODE:="43", EDI_RECEIVED_VALUE:=EDI_PO_PURP)
                End If

                If rowEDT850T1.Item("EDI_SHIP_DC") & "" <> "" Then
                    EDI_SHIP_DC = rowEDT850T1.Item("EDI_SHIP_DC") & ""
                Else
                    EDI_SHIP_DC = rowEDT850T1.Item("EDI_CENTER_CODE") & ""
                End If

                If EDI_SHIP_DC = "" And EDI_STORE = "" Then
                    Dim rowEDT850T5s() As DataRow = dst.Tables("EDT850T5").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_ADDR_TYPE = 'ST'")
                    If rowEDT850T5s.Length > 0 Then
                        EDI_ADDR_CODE = rowEDT850T5s(0).Item("EDI_ADDR_CODE") & ""
                    End If
                End If

                If EDI_SHIP_DC <> "" Then
                    ORDR_SHIP_TO = "DC"
                End If
                If EDI_STORE <> "" Or EDI_SHIP_DC <> "" Or EDI_ADDR_CODE <> "" Then
                    Get_Ship_To(EDI_STORE, EDI_SHIP_DC, EDI_ADDR_CODE)
                End If

                ' Load values for Ship & Cancel Dates, and check for Reasonableness

                ORDR_CANCEL_DATE = Nothing
                ORDR_SHIP_DATE = Nothing
                ORDR_ARRIVAL_DATE = Nothing
                ORDR_LAST_ARRIVAL_DATE = Nothing

                If rowEDT850T1.Item("EDI_ARRIVAL_DATE") & "" <> "" Then
                    ORDR_ARRIVAL_DATE = rowEDT850T1.Item("EDI_ARRIVAL_DATE") & ""
                End If
                If rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") & "" <> "" Then
                    ORDR_LAST_ARRIVAL_DATE = rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") & ""
                End If

                If rowEDT850T1.Item("EDI_END_DATE") & "" <> "" Then
                    ORDR_CANCEL_DATE = rowEDT850T1.Item("EDI_END_DATE")
                Else
                    If rowEDT850T1.Item("EDI_LAST_ARRIVAL_DATE") & "" <> "" Then
                        If CUST_CODE = "ULTA" Then 'As per Ulta: It’s not business days, it’s just +/- 3 days
                            'ORDR_CANCEL_DATE = ASCMAIN1.DateDiff_Weekday(ORDR_LAST_ARRIVAL_DATE, -1 * TRANSIT_BUS_DAYS)
                            ORDR_CANCEL_DATE = DateAdd("d", -1 * TRANSIT_BUS_DAYS, ORDR_LAST_ARRIVAL_DATE)
                        Else
                            ORDR_CANCEL_DATE = ASCMAIN1.DateDiff_Weekday(ORDR_LAST_ARRIVAL_DATE, -1 * TRANSIT_BUS_DAYS)
                        End If

                    ElseIf rowEDT850T1.Item("EDI_ARRIVAL_DATE") & "" <> "" Then
                        If CUST_CODE = "ULTA" Then 'As per Ulta: It’s not business days, it’s just +/- 3 days
                            ORDR_CANCEL_DATE = DateAdd("d", -1 * TRANSIT_BUS_DAYS, ORDR_ARRIVAL_DATE)
                        Else
                            ORDR_CANCEL_DATE = ASCMAIN1.DateDiff_Weekday(ORDR_ARRIVAL_DATE, -1 * TRANSIT_BUS_DAYS)
                        End If
                    End If
                End If

                If rowEDT850T1.Item("EDI_START_DATE") & "" <> "" Then
                    ORDR_SHIP_DATE = rowEDT850T1.Item("EDI_START_DATE")
                ElseIf rowEDT850T1.Item("EDI_SHIP_DATE") & "" <> "" Then
                    ORDR_SHIP_DATE = rowEDT850T1.Item("EDI_SHIP_DATE")
                ElseIf CUST_CODE = "ULTA" And rowEDT850T1.Item("EDI_ARRIVAL_DATE") & "" <> "" Then
                    ORDR_SHIP_DATE = DateAdd("d", -1 * TRANSIT_BUS_DAYS - 3, ORDR_ARRIVAL_DATE)
                ElseIf Not (ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA") _
                    AndAlso (TAC.TACMAIN1.IPLBMacysCustomerCodes.Contains(CUST_CODE) _
                    And rowEDT850T1.Item("EDI_ARRIVAL_DATE") & "" <> "") Then
                    ORDR_SHIP_DATE = ORDR_ARRIVAL_DATE
                    'For I As Integer = 1 To TRANSIT_BUS_DAYS
                    '    ORDR_SHIP_DATE = ORDR_SHIP_DATE.AddDays(-1)
                    '    If ORDR_SHIP_DATE.DayOfWeek = DayOfWeek.Sunday Then
                    '        ORDR_SHIP_DATE = ORDR_ARRIVAL_DATE.AddDays(-2)
                    '    ElseIf ORDR_ARRIVAL_DATE.DayOfWeek = DayOfWeek.Sunday Then
                    '        ORDR_SHIP_DATE = ORDR_ARRIVAL_DATE.AddDays(-1)
                    '    End If
                    'Next
                    '' ORDR_SHIP_DATE = DateAdd("d", -1 * TRANSIT_BUS_DAYS, ORDR_ARRIVAL_DATE)

                ElseIf Format(ORDR_CANCEL_DATE, "yyyyMMdd") <> "00010101" Then
                    ORDR_SHIP_DATE = rowEDT850T1.Item("EDI_PO_DATE")
                End If

                If Format(ORDR_CANCEL_DATE, "yyyyMMdd") = "00010101" Then
                    If Format(ORDR_SHIP_DATE, "yyyyMMdd") <> "00010101" Then
                        ORDR_CANCEL_DATE = ORDR_SHIP_DATE.AddDays(6)
                    End If
                End If

                If ASCMAIN1.CLIENT = "INT" Then
                    If CUST_CODE = "ULTA" Then
                        'As per MM/LBM: 02/21/17
                        '•         Ship Date should be Ship Date -1
                        '•         Cancel Date – if calculated to earlier than Ship Date, then set Cancel Date to Ship Date +1

                        ' lbm email 07/24/18 : Use EDI_START_DATE for ORDR_SHIP_DATE, and then add 2 business days for ORDR_CANCEL_DATE

                        ' ORDR_SHIP_DATE = ORDR_SHIP_DATE.AddDays(-1)
                        'If Format(ORDR_SHIP_DATE, "yyyyMMdd") >= Format(ORDR_CANCEL_DATE, "yyyyMMdd") Then
                        '    ORDR_CANCEL_DATE = AddDaysSkipWeekends(ORDR_SHIP_DATE, 2)
                        'End If


                        ' update 01/03/2019
                        ' Lauren - I remember.  Thanks…..but……as much as their document states collect vendors should use ship date it is the arrival date that counts so can we go back?  I found this out after the fact.  Sorry.
                        ' Maria  I think we need to comment out lines 1090 to 1105

                        'ORDR_SHIP_DATE = rowEDT850T1.Item("EDI_START_DATE")
                        'ORDR_CANCEL_DATE = AddDaysSkipWeekends(ORDR_SHIP_DATE, 2)

                        ' AS PER LM 02/27/19
                        ' ORDR_CANCEL_DATE = ORDR_SHIP_DATE.AddDays(2)

                        ' AS PER NF March 2025:
                        ' Need to set order cancel date with new logic instead of above, and then overwrite the order fields for the ship (Cancel - 6 calendar) and arrival (cancel + transit).
                        ' Test by blanking out edt850t1.EDI_START_DATE on a record and reprocess - make sure it's po date + 11
                        Dim DC_TRANSIT_CAL_DAYS As Int16 = 0

                        ROWSOTSVIA3 = dst.Tables("SOTSVIA3").Rows.Find(New String() {WHSE_CODE, CUST_CODE, CUST_DC_NO})
                        If ROWSOTSVIA3 Is Nothing Then
                            DC_TRANSIT_CAL_DAYS = 0
                            Bad_Data(EDI_COND_DESC:="DC " & CUST_DC_NO & " is Missing Transit Days", EDI_COND_CODE:="48", EDI_RECEIVED_VALUE:="")
                        Else
                            DC_TRANSIT_CAL_DAYS = Val(ROWSOTSVIA3.Item("TRANSIT_BUS_DAYS") & "") ' column is named for business days but we are storing and need calendar days for Ulta

                        End If

                        Dim MinCancelDate As Date = rowEDT850T1.Item("EDI_PO_DATE").AddDays(11)
                        If IsDate(rowEDT850T1.Item("EDI_START_DATE") & "") AndAlso (MinCancelDate < CDate(rowEDT850T1.Item("EDI_START_DATE"))) Then
                            MinCancelDate = CDate(rowEDT850T1.Item("EDI_START_DATE"))
                        End If
                        Dim DaysToAdd As Int16 = 0
                        If (Not (String.IsNullOrEmpty(ROWSOTSVIA3.Item("PICKUP_DAY") & ""))) Then
                            Dim DC_PICKUP_DAY_OF_WEEK As DayOfWeek = Val(ROWSOTSVIA3.Item("PICKUP_DAY"))
                            DaysToAdd = DC_PICKUP_DAY_OF_WEEK - MinCancelDate.DayOfWeek
                        End If

                        DaysToAdd = If(DaysToAdd < 0, 7 + DaysToAdd, DaysToAdd)

                        ORDR_CANCEL_DATE = MinCancelDate.AddDays(DaysToAdd)
                        ORDR_SHIP_DATE = ORDR_CANCEL_DATE.AddDays(-6)
                        ORDR_ARRIVAL_DATE = ORDR_CANCEL_DATE.AddDays(DC_TRANSIT_CAL_DAYS)

                    End If

                    If CUST_CODE = "AAFES" Then
                        'As per MM/LBM: 10/30/17
                        'The rule should be from arrival date.  4 day travel time to Mississippi.  Arrival is 11/19.  We have a 48 allowable window before or after.  So I would make the rule:  
                        'For start ship – backout 6 business days from arrival
                        'For cancel date – backout 4 business days from arrival
                        If Format(ORDR_ARRIVAL_DATE, "yyyyMMdd") <> "00010101" Then
                            'ORDR_SHIP_DATE = ORDR_ARRIVAL_DATE.AddDays(-1 * (2 + TRANSIT_BUS_DAYS))
                            'ORDR_CANCEL_DATE = ORDR_ARRIVAL_DATE.AddDays(-1 * (0 + TRANSIT_BUS_DAYS))

                            ' LM email 02/18
                            '  - USE above code to set ORDR_SHIP_DATE from EDI_START_DATE
                            '    If rowEDT850T1.Item("EDI_START_DATE") & "" <> "" Then ORDR_SHIP_DATE = rowEDT850T1.Item("EDI_START_DATE")
                            'ORDR_SHIP_DATE = ASCMAIN1.DateDiff_Weekday(ORDR_ARRIVAL_DATE, -1 * (2 + TRANSIT_BUS_DAYS))

                            ' LM email 02/18 = add 1 day
                            'ORDR_CANCEL_DATE = ASCMAIN1.DateDiff_Weekday(ORDR_ARRIVAL_DATE, -1 * (0 + TRANSIT_BUS_DAYS))
                            'ORDR_CANCEL_DATE = ASCMAIN1.DateDiff_Weekday(ORDR_ARRIVAL_DATE, -1 * (0 + TRANSIT_BUS_DAYS - 1))
                            ' 02/25/2020 W/LBM - USE ARRIVAL DATE - (TRANSIT + 1) FOR SHIP DATE, USE SHIP+3 FOR CANCEL
                            ORDR_SHIP_DATE = ORDR_ARRIVAL_DATE.AddDays(-1 * (TRANSIT_BUS_DAYS + 1))
                            ORDR_CANCEL_DATE = ORDR_SHIP_DATE.AddDays(3)

                        End If
                    End If

                    If CUST_CODE = "BELK" Or CUST_CODE = "NEIMANDIR" Or CUST_CODE = "NEIMAN" Or CUST_CODE = "DILLARDS" Then
                        ' ISSUE-7280 1/8/26 - per NF, include Dillard's in this logic

                        ' I would do the Friday. Thanks. 
                        ' On Nov 27, 2018, at 9:27 AM, Maria Mattina <maria@absolution.com> wrote:
                        ' Lauren-
                        ' What if the date they send is on Monday? Do we back it up to Friday? Or Sunday?
                        ' Maria Mattina
                        ' From: Lauren MARINELLI <lmarinelli@interparfums.com>
                        ' Date: Monday, November 12, 2018 at 12:42 PM
                        ' To: "Maria (ABS)" <maria@absolution.com>
                        ' Cc: Walter Zielenski <wjz@absolution.com>, Kara Savasta <ksavasta@interparfums.com>
                        ' Subject: belk and neiman
                        ' Hi. I need  you to move up Belk’s  and Neiman, Neimandir cancel dates – 1 day please.
                        ' So if they send a cancel date of November 19th it should read as November 18th.

                        ORDR_CANCEL_DATE = ORDR_CANCEL_DATE.AddDays(-1)
                        If ORDR_CANCEL_DATE.DayOfWeek = DayOfWeek.Sunday Then
                            ORDR_CANCEL_DATE = ORDR_CANCEL_DATE.AddDays(-2)
                        End If
                        If ORDR_CANCEL_DATE.DayOfWeek = DayOfWeek.Saturday Then
                            ORDR_CANCEL_DATE = ORDR_CANCEL_DATE.AddDays(-1)
                        End If
                    End If
                End If

                If Format(ORDR_CANCEL_DATE, "yyyyMMdd") < Format(Now, "yyyyMMdd") Then
                    EDI_RECEIVED_VALUE = Format$(ORDR_CANCEL_DATE, "MM/dd/yyyy")
                    EDI_REPLACED_VALUE = Bad_Data(
                        EDI_COND_DESC:="Cancel Date is past",
                        EDI_COND_CODE:="09",
                        EDI_RECEIVED_VALUE:=EDI_RECEIVED_VALUE)
                    If EDI_RECEIVED_VALUE <> "" Then
                        ORDR_CANCEL_DATE = DateValue(EDI_RECEIVED_VALUE)
                    End If
                End If


                If Format(ORDR_SHIP_DATE, "yyyyMMdd") = "00010101" Then
                    EDI_REPLACED_VALUE = Bad_Data(
                                           EDI_COND_DESC:="Missing Ship Date",
                                           EDI_COND_CODE:="08",
                                           EDI_RECEIVED_VALUE:=Format$(ORDR_SHIP_DATE, "MM/dd/yyyy"))
                    ' we do not allow accepting bad edi value for this field, so follownig lines are not used
                    If EDI_RECEIVED_VALUE <> "" Then
                        ORDR_SHIP_DATE = DateValue(EDI_RECEIVED_VALUE)
                    End If
                Else
                    If Format(ORDR_CANCEL_DATE, "yyyyMMdd") = "00010101" Then
                        ORDR_CANCEL_DATE = ORDR_SHIP_DATE.AddDays(7)
                    End If
                End If

                If CUST_CODE <> "" Then
                    Get_Terms()
                    Check_for_Possible_Order_Duplication()
                End If



                multi_customer = False
                CUST_CODEs.Clear()
                CUST_CODEs.Add(CUST_CODE, New List(Of String))

                If CUST_CODE = "NEIMAN" Then
                    ASCMAIN1.sql = ""
                    For i As Integer = 1 To 10
                        ASCMAIN1.sql &= " union Select Distinct EDI_STORE_" & Format(i, "00") & " EDI_STORE_XX" _
                            & " from EDT850T3 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    Next
                    ASCMAIN1.sql = "Select * from (" & Mid(ASCMAIN1.sql, 8) & ") where EDI_STORE_XX is Not Null"
                    For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                        Dim EDI_STORE_XX As String = row.Item("EDI_STORE_XX")
                        Dim rowEDTTRPM3 As DataRow = dst.Tables("EDTTRPM3").Rows.Find(New String() {EDI_TP_QUAL, EDI_TP_ID, "850", EDI_STORE_XX})
                        If rowEDTTRPM3 IsNot Nothing Then
                            Dim CUST_CODE_XX As String = rowEDTTRPM3.Item("CUST_CODE") & ""
                            If CUST_CODE_XX = CUST_CODE Or CUST_CODE_XX = "" Then
                                Throw New Exception("Bad Customer in EDTTRPM3 (" & CUST_CODE_XX & ")")
                            End If
                            If CUST_CODE_XX <> "" Then
                                If Not CUST_CODEs.ContainsKey(CUST_CODE_XX) Then
                                    CUST_CODEs.Add(CUST_CODE_XX, New List(Of String))
                                    Fill_Records("ARTCUST2", CUST_CODE_XX, False)
                                    Fill_Records("SOTCITM1", CUST_CODE_XX, False)
                                End If

                                CUST_CODEs(CUST_CODE_XX).Add(EDI_STORE_XX)
                                CUST_CODEs(CUST_CODE).Add(EDI_STORE_XX)
                            End If
                        End If
                    Next

                    If CUST_CODEs.Count > 1 Then multi_customer = True
                End If

                Dim ORDR_GROUP_NOs As New List(Of String)
                Dim ORDR_GROUP_NOs_ADS As New List(Of String)

                dst.Tables("SOTORDR1").Rows.Clear()
                dst.Tables("SOTORDR2").Rows.Clear()
                dst.Tables("SOTORDR4").Rows.Clear()

                For Each CUST_CODE_multi In CUST_CODEs.Keys

                    dst.Tables("EDT850TS").Rows.Clear()
                    dst.Tables("EDTITEMX").Rows.Clear()
                    dst.Tables("EDTSDQT1").Rows.Clear()
                    dst.Tables("EDTSDQT0").Rows.Clear()

                    Dim rowARTCUST1_multi As DataRow = LookUp("ARTCUST1", CUST_CODE_multi)

                    Process_EDT850T2(EDI_STORE, EDI_SHIP_DC, EDI_ADDR_CODE)


                    If dst.Tables("EDT850TS").Rows.Count = 0 Then
                        If EDI_STORE = "" Then
                            If CUST_CODEs.Count <> 0 And CUST_CODE = CUST_CODE_multi Then
                                ' DO NOTHING
                            Else
                                Bad_Data(EDI_COND_DESC:="No Store found for Order, Verify EDI Store",
                                        EDI_COND_CODE:="40",
                                        EDI_RECEIVED_VALUE:="Missing Store")
                            End If
                        End If
                    Else
                        If ASCMAIN1.CLIENT = "AHA" Or ASCMAIN1.CLIENT = "INT" Then ' DOING THIS FOR AHA ONLY AT THIS POINT
                            Dim rowEDT850T5s() As DataRow = dst.Tables("EDT850T5").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_ADDR_TYPE = 'ST'")
                            If dst.Tables("EDT850TS").Rows.Count = 1 And rowEDT850T5s.Length = 1 Then ' AT SOME POINT WE NEED TO CHECK ALL SHIPPING ADDRESSES

                                Dim rowEDT850TS As DataRow = dst.Tables("EDT850TS").Rows(0)
                                Dim CUST_STORE_NO As String = rowEDT850TS.Item("CUST_STORE_NO")
                                Dim CUST_DC_NO As String = rowEDT850TS.Item("CUST_DC_NO") & ""
                                Dim rowARTCUST2 As DataRow = Nothing
                                If CUST_DC_NO <> "" Then
                                    rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_DC_NO})
                                Else
                                    rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                                End If

                                If rowARTCUST2 IsNot Nothing Then
                                    If ASCMAIN1.CLIENT = "AHA" Then
                                        Dim ABS_CSZ As String = rowARTCUST2.Item("CUST_STORE_CITY") & ":" & rowARTCUST2.Item("CUST_STORE_STATE") & ":" & rowARTCUST2.Item("CUST_STORE_ZIP_CODE")
                                        Dim EDI_CSZ As String = rowEDT850T5s(0).Item("EDI_CITY") & ":" & rowEDT850T5s(0).Item("EDI_STATE") & ":" & rowEDT850T5s(0).Item("EDI_ZIPCODE")
                                        ABS_CSZ = Replace(ABS_CSZ, "-", "")
                                        EDI_CSZ = Replace(EDI_CSZ, "-", "")
                                        If ABS_CSZ.ToUpper <> EDI_CSZ.ToUpper And EDI_CSZ <> "::" Then
                                            Bad_Data(EDI_COND_DESC:="Mis-match in Ship-To Address", EDI_COND_CODE:="49", EDI_RECEIVED_VALUE:=EDI_CSZ, EDI_EXPECTED_VALUE:=ABS_CSZ)
                                            'If EDI_ACTION = "S" Then
                                            '    skip_store = True
                                            'Else
                                            '    Bad_Data_Cond_05 = True
                                            'End If
                                        End If
                                    ElseIf ASCMAIN1.CLIENT = "INT" Then
                                        Dim ABS_ZIP As String = rowARTCUST2.Item("CUST_STORE_ZIP_CODE").ToString()
                                        Dim ABS_COUNTRY As String = rowARTCUST2.Item("CUST_STORE_COUNTRY").ToString()
                                        If Not String.IsNullOrEmpty(ABS_ZIP) Then
                                            Select Case ABS_ZIP.Length
                                                Case 5 '12345 - ZIP CODE WITH 5 DIGITS
                                                    ' No change needed
                                                Case 6 'NIC123 - Canadian zip with no spaces
                                                    'if the ABS Zip is 6 digits but not canadian, something is wrong
                                                    If ABS_COUNTRY.ToUpper() <> "CA" AndAlso ABS_COUNTRY.ToUpper() <> "CAN" Then
                                                        ABS_ZIP = ""
                                                    End If
                                                Case 7 'NIC 123 (CANADIAN)
                                                    'if it's 7 digits and canadian, we just drop the space for comparison
                                                    If ABS_COUNTRY.ToUpper() = "CA" Or ABS_COUNTRY.ToUpper() = "CAN" Then
                                                        ABS_ZIP = ABS_ZIP.Replace(" ", "")
                                                    End If
                                                Case 9 '123456789
                                                    ABS_ZIP = ABS_ZIP.Substring(0, 5)
                                                Case 10
                                                    If ABS_ZIP.Contains("-") Then '12345-6789
                                                        ABS_ZIP = ABS_ZIP.Split("-"c)(0)
                                                    End If
                                                Case Else
                                                    ABS_ZIP = "" ' Invalid or unexpected length, treat as empty
                                            End Select
                                        Else
                                            ABS_ZIP = ""
                                        End If

                                        Dim EDI_ZIP As String = rowEDT850T5s(0).Item("EDI_ZIPCODE").ToString()
                                        Dim EDI_COUNTRY As String = rowEDT850T5s(0).Item("EDI_COUNTRY").ToString()
                                        If Not String.IsNullOrEmpty(EDI_ZIP) Then
                                            Select Case EDI_ZIP.Length
                                                Case 5 '12345
                                                    ' No change needed
                                                Case 6 '012345
                                                    'if it isn't canadian, then it is a 5 digit zip with a leading zero
                                                    If EDI_COUNTRY.ToUpper() <> "CA" AndAlso EDI_COUNTRY.ToUpper() <> "CAN" Then
                                                        If EDI_ZIP.StartsWith("0") Then
                                                            EDI_ZIP = EDI_ZIP.Substring(1)
                                                        End If
                                                    End If
                                                Case 7 'NIC 123 (CANADIAN)
                                                    If EDI_COUNTRY.ToUpper() = "CA" Or EDI_COUNTRY.ToUpper() = "CAN" Then
                                                        EDI_ZIP = EDI_ZIP.Replace(" ", "")
                                                        'other country w 6 digit zip and a leading zero
                                                    Else
                                                        EDI_ZIP = EDI_ZIP.Substring(1)
                                                    End If
                                                Case 8 '0NIC 123 (canadian)
                                                    If EDI_COUNTRY.ToUpper() = "CA" Or EDI_COUNTRY.ToUpper() = "CAN" Then
                                                        EDI_ZIP = EDI_ZIP.Substring(1).Replace(" ", "")
                                                        'country with 7 digit zip and a leading zero
                                                    Else
                                                        EDI_ZIP = EDI_ZIP.Substring(1)
                                                    End If
                                                Case 9 '123456789
                                                    EDI_ZIP = EDI_ZIP.Substring(0, 5)
                                                Case 10
                                                    If EDI_ZIP.Contains("-") Then '12345-6789
                                                        EDI_ZIP = EDI_ZIP.Split("-"c)(0)
                                                    Else '0123456789
                                                        EDI_ZIP = EDI_ZIP.Substring(1, 5)
                                                    End If
                                                Case Else
                                                    EDI_ZIP = "" ' Invalid or unexpected length, treat as empty
                                            End Select
                                        Else
                                            EDI_ZIP = ""
                                        End If

                                        If Not String.IsNullOrEmpty(ABS_ZIP) And Not String.IsNullOrEmpty(EDI_ZIP) Then
                                            If ABS_ZIP.ToUpper <> EDI_ZIP.ToUpper Then
                                                Bad_Data(EDI_COND_DESC:="Ship to zip code (" & EDI_ZIP & ") does not match ABS (" & ABS_ZIP & ")", EDI_COND_CODE:="42", EDI_RECEIVED_VALUE:=EDI_ZIP, EDI_EXPECTED_VALUE:=ABS_ZIP)
                                            End If
                                        End If
                                    End If
                                Else
                                    ' 02/07/2025 - Invalid Store caused an error message but the code kept executing and SOTORDR5 was not created for a batch of Bloomies Orders.
                                    ' There is code where 'Rollback is commented out that caused the error. It was chnaged to exit the loop, not create orders and let the user know of the error.
                                    Dim sql As String = $"EDI_DOC_SEQ_NO = '{rowEDT850T1.Item("EDI_DOC_SEQ_NO") & String.Empty}' 
                                                            AND (CUST_STORE_NO = '{If(EDI_SHIP_DC.Length > 0, EDI_SHIP_DC, CUST_STORE_NO)}' OR EDI_RECEIVED_VALUE = '{If(EDI_SHIP_DC.Length > 0, EDI_SHIP_DC, CUST_STORE_NO)}')
                                                            AND EDI_COND_CODE = '05' AND EDI_ACTION = 'S'"
                                    If dst.Tables("EDT850TE").Select(sql).Length > 0 Then
                                        EDI_DOC_SEQ_NO_ok = False
                                    End If
                                End If
                            End If
                        End If
                    End If

                    ' If the Journal Passed all tests, then generate orders from SDQ's
                    If EDI_DOC_SEQ_NO_ok Then

                        Dim order_count As Int32 = dst.Tables("EDT850TS").Rows.Count
                        If order_count <> 0 Then

                            Dim ORDER_COUNTER As Int32 = 0

                            Try
                                'BeginTrans()

                                ' PROBABLY DON'T WANT TO WASTE ORDER NUMBERS, BUT IF WE GET THIS FAR, WE ARE MORE THAN LIKELY GENERATING ORDERS

                                Dim ORDR_NO_next As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO", order_count)
                                Dim ORDR_NO_beg = ORDR_NO_next
                                Dim ORDR_NO_end = ORDR_NO_next

                                ' CHECK LEAD ITEM TO SEE IF IT IS MARKED FOR ADS, AND IF IT IS, CHANGE WHSE TO ADS
                                ' THIS CODE SHOULD EVENTUALLY CHECK ANY LP_CODE (NOT JUST ADS)
                                ' THIS CODE SHOULD CHECK ALL ITEMS AND PRODUCE AN ERROR IF THE REMAINING ITEMS DON'T FOLLOW THE RULE FOR THE LEAD ITEM
                                ' NEED TO GET CLA RECORDS OUT OF ICTITLP2
                                ' AT SOME POINT DEAL WITH ISSUE WHERE GET_SHIP_TO RELIES ON WHSE CODE TO GET TRANSIT DAYS, AND PRE-FLIGHT MAY RESOLVE DIFFERENTLY THAN WHEN CALLING GET_SHIP_TO AFTER CHANGINT THE WHSE_CODE

                                ' This code does not use the lead item.
                                ' If any items are in ICTWHSE2 then the order is changed to that warehouse for that LP Code

                                ' Code changed for ISSUE-7230 
                                ' ISSUE-7373 Removed if statement that disallows Amazon orders from coming through the warehouse assignment logic
                                ' NF no longer wants Amazon orders split - they should be already warehouse-specific because IPLB now has two
                                ' Amazon vendor accounts. Logic is the same as for other customers.
                                If WHSE_CODE_customer.Length = 0 Then
                                        Dim lstWarehouses As New List(Of String)
                                        For Each rowEDTSDQT1 As DataRow In ASCDATA1.SelectDistinct(dst.Tables("EDTSDQT1"), {"ITEM_CODE"}).Rows
                                            Dim ITEM_CODE As String = rowEDTSDQT1.Item("ITEM_CODE")
                                            Dim drICTWHSE1 As DataRow = dst.Tables("ICTWHSE1").Rows.Find(WHSE_CODE)
                                            If drICTWHSE1 IsNot Nothing Then
                                                Dim LP_CODE As String = drICTWHSE1.Item("LP_CODE") & String.Empty
                                                If LP_CODE.Length > 0 Then
                                                    If dst.Tables("ICTWHSE2").Rows.Count = 0 OrElse dst.Tables("ICTWHSE2").Select("LP_CODE = '{LP_CODE}'").Length = 0 Then
                                                        Fill_Records("ICTWHSE2", {LP_CODE})

                                                        ' This prevent running a query tht returns 0 rows multiple times
                                                        If dst.Tables("ICTWHSE2").Rows.Count = 0 Then
                                                            dst.Tables("ICTWHSE2").Rows.Add("X@@X", "X@@X")
                                                            dst.Tables("ICTWHSE2").Rows(0).Item("LP_CODE") = LP_CODE
                                                        End If
                                                    End If


                                                    If dst.Tables("ICTWHSE2").Select($"ITEM_CODE = '{ITEM_CODE}' AND LP_CODE = '{LP_CODE}'").Length > 0 Then
                                                        WHSE_CODE = dst.Tables("ICTWHSE2").Select($"ITEM_CODE = '{ITEM_CODE}' AND LP_CODE = '{LP_CODE}'")(0).Item("WHSE_CODE")
                                                    End If
                                                    If Not lstWarehouses.Contains(WHSE_CODE) Then
                                                        lstWarehouses.Add(WHSE_CODE)
                                                    End If
                                                End If
                                            End If
                                        Next

                                        If lstWarehouses.Count > 1 Then
                                            MessageBox.Show($"WARNING: Customer {CUST_CODE}, P.O. {ORDR_CUST_PO} has items in the following Warehouses: {String.Join(", ", lstWarehouses.ToArray)}. Orders set to use Warehouse {WHSE_CODE}", "Import Orders", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                        End If

                                    End If


                                    ITEM_CODE = ""

                                Dim ORDR_GROUP_NO As String = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                                Dim ORDR_LNO As Int32 = 0
                                Dim blnAMZN_AND_ADS As Boolean = False

                                ASCMAIN1.Progress("Generating Orders", "")

                                For Each rowEDT850TS As DataRow In dst.Tables("EDT850TS").Select("", "CUST_STORE_NO") '("", "EDI_STORE")
                                    EDI_STORE = rowEDT850TS.Item("EDI_STORE")
                                    EDI_SHIP_DC = rowEDT850TS.Item("EDI_SHIP_DC")
                                    EDI_ADDR_CODE = rowEDT850TS.Item("EDI_ADDR_CODE")
                                    CUST_STORE_NO = rowEDT850TS.Item("CUST_STORE_NO")
                                    CUST_DC_NO = rowEDT850TS.Item("CUST_DC_NO") & ""
                                    Dim SALES_DIVISION_CODE As String = "" '?
                                    Get_Ship_To(EDI_STORE, EDI_SHIP_DC, EDI_ADDR_CODE)

                                    Dim ORDR_NO As String = ORDR_NO_next
                                    ORDR_NO_end = ORDR_NO_next
                                    ORDR_NO_next = Format(Val(ORDR_NO_next) + 1, "0000000000")
                                    ORDR_LNO = 0

                                    ASCMAIN1.Progress("-", ORDR_NO)

                                    ' Write Order Header

                                    Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow
                                    With rowSOTORDR1
                                        ORDER_COUNTER = ORDER_COUNTER + 1
                                        .Item("ORDR_NO") = ORDR_NO
                                        .Item("ORDR_DATE") = ORDR_DATE
                                        .Item("CUST_CODE") = CUST_CODE_multi
                                        .Item("CUST_NAME") = rowARTCUST1_multi.Item("CUST_NAME")
                                        .Item("CUST_STORE_NO") = CUST_STORE_NO
                                        .Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION") & ""
                                        .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                                        .Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE

                                        If Not (ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA") _
                                            AndAlso TAC.TACMAIN1.IPLBMacysCustomerCodes.Contains(CUST_CODE) Then
                                            Dim ORDR_SHIP_DATE_CALC As Date = ORDR_SHIP_DATE
                                            For I As Integer = 1 To TRANSIT_BUS_DAYS
                                                ORDR_SHIP_DATE_CALC = ORDR_SHIP_DATE_CALC.AddDays(-1)
                                                If ORDR_SHIP_DATE_CALC.DayOfWeek = DayOfWeek.Sunday Then
                                                    ORDR_SHIP_DATE_CALC = ORDR_SHIP_DATE_CALC.AddDays(-2)
                                                ElseIf ORDR_SHIP_DATE_CALC.DayOfWeek = DayOfWeek.Sunday Then
                                                    ORDR_SHIP_DATE_CALC = ORDR_SHIP_DATE_CALC.AddDays(-1)
                                                End If
                                            Next
                                            .Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE_CALC
                                        End If

                                        .Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
                                        .Item("ORDR_ORIG_SHIP_DATE") = ORDR_SHIP_DATE
                                        .Item("ORDR_ORIG_CANCEL_DATE") = ORDR_CANCEL_DATE
                                        .Item("POST_CODE") = rowARTCUST1bt.Item("POST_CODE")
                                        .Item("TERM_CODE") = TERM_CODE
                                        .Item("SHIP_VIA_CODE") = rowARTCUST1_multi.Item("SHIP_VIA_CODE")

                                        If CUST_DC_NO <> "" And ORDR_SHIP_TO = "DC" Then
                                            Dim SHIP_VIA_CODE As String = rowARTCUST2_DC.Item("CUST_STORE_SHIP_VIA_CODE") & ""
                                            If SHIP_VIA_CODE <> "" Then
                                                .Item("SHIP_VIA_CODE") = SHIP_VIA_CODE
                                            End If
                                        End If

                                        Dim CUST_STORE_STATE As String = ""
                                        If rowARTCUST2_DC IsNot Nothing Then
                                            CUST_STORE_STATE = rowARTCUST2_DC.Item("CUST_STORE_STATE") & ""
                                        Else
                                            CUST_STORE_STATE = rowARTCUST2.Item("CUST_STORE_STATE") & ""
                                        End If

                                        Dim rowSOTSVIAX As DataRow = LookUp("SOTSVIAX", New String() {WHSE_CODE, CUST_STORE_STATE})
                                        If rowSOTSVIAX IsNot Nothing Then
                                            .Item("SHIP_VIA_CODE") = rowSOTSVIAX.Item("SHIP_VIA_CODE").ToString()
                                        End If

                                        .Item("EDI_MERCH_TYPE") = rowEDT850T1.Item("EDI_MERCH_TYPE")
                                        .Item("CUST_VEND_REF") = rowARTCUST1bt.Item("CUST_VEND_REF")

                                        If ASCMAIN1.CLIENT = "AHA" Then
                                            If rowARTCUST2.Item("SREP_CODE") & "" <> "" Then
                                                .Item("SREP_CODE") = rowARTCUST2.Item("SREP_CODE")
                                                .Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE")
                                            Else
                                                .Item("SREP_CODE") = rowARTCUST1_multi.Item("SREP_CODE")
                                            End If
                                        Else
                                            .Item("SREP_CODE") = rowARTCUST1_multi.Item("SREP_CODE")
                                            .Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE")
                                        End If

                                        .Item("SREP2_CODE") = rowARTCUST1_multi.Item("SREP2_CODE")
                                        .Item("WHSE_CODE") = WHSE_CODE
                                        .Item("ORDR_TYPE_CODE") = "REG"

                                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                        .Item("LAST_OPER") = ASCMAIN1.USER_ID
                                        .Item("INIT_DATE") = DATETIME_STAMP
                                        .Item("LAST_DATE") = DATETIME_STAMP
                                        .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date
                                        .Item("ORDR_DEPT") = ORDR_DEPT
                                        .Item("ORDR_SHIP_TO") = ORDR_SHIP_TO
                                        .Item("ORDR_ADDR_TYPE_ST") = ORDR_SHIP_TO

                                        .Item("ORDR_STATUS") = "O"
                                        .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
                                        .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
                                        .Item("ORDR_PRIORITY") = rowARTCUST1_multi.Item("CUST_PRIORITY_CODE") & ""

                                        .Item("EDI_JRNL_NO") = EDI_JRNL_NO
                                        .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                        .Item("CUST_DC_NO") = CUST_DC_NO
                                        .Item("CUST_BILL_TO_CUST") = CUST_BILL_TO_CUST
                                        .Item("ORDR_SOURCE") = "E"
                                        .Item("PRICE_CLASS_CODE") = rowARTCUST1_multi.Item("PRICE_CLASS_CODE") & ""
                                        .Item("TRADE_CLASS_CODE") = rowARTCUST1_multi.Item("TRADE_CLASS_CODE") & ""
                                        .Item("PRICE_LIST_CODE") = PRICE_LIST_CODE_OVERRIDE
                                        .Item("CURR_CODE") = CURR_CODE
                                        .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                                        .Item("CUST_DISC_PCT") = PRICE_BASE_DPCT
                                        .Item("CUST_DC_NO") = CUST_DC_NO
                                        '.Item("ITEM_BRAND_CODE") = ITEM_BRAND_CODE
                                        SALES_DIVISION_CODE = "?"
                                        .Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                                        .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO

                                        FRT_TERMS = rowARTCUST1_multi.Item("FRT_TERMS") & ""
                                        .Item("FRT_TERMS") = FRT_TERMS
                                        .Item("ORDR_FOB") = rowARTCUST1_multi.Item("CUST_FOB") & ""
                                        .Item("ORDR_EDI_810") = ORDR_EDI_810
                                        .Item("ORDR_EDI_856") = ORDR_EDI_856
                                        If Format(ORDR_LAST_ARRIVAL_DATE, "yyyyMMdd") <> "00010101" Then .Item("ORDR_LAST_ARRIVAL_DATE") = ORDR_LAST_ARRIVAL_DATE
                                        If Format(ORDR_ARRIVAL_DATE, "yyyyMMdd") <> "00010101" Then .Item("ORDR_ARRIVAL_DATE") = ORDR_ARRIVAL_DATE
                                        .Item("ORDR_BILL_SHIP_TO") = rowARTCUST1_multi.Item("CUST_BILL_SHIP_TO") & ""

                                        If (rowSOTORDR1.Item("CUST_CODE") & "" = "AMAZON" And rowSOTORDR1.Item("WHSE_CODE") & "" <> "ADS") Then
                                            blnAMZN_AND_ADS = True
                                        End If
                                    End With
                                    dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                                    ITEM_CODE = ""

                                    For Each rowEDTSDQT1 As DataRow In dst.Tables("EDTSDQT1").Select("EDI_STORE = '" & EDI_STORE & "'", "EDI_DTL_SEQ")
                                        Dim EDI_DTL_SEQ As Int32 = Val(rowEDTSDQT1.Item("EDI_DTL_SEQ") & "")
                                        Dim ITEM_CODE As String = rowEDTSDQT1.Item("ITEM_CODE") & ""
                                        Dim EDI_UPC As String = rowEDTSDQT1.Item("EDI_UPC") & ""
                                        Dim EDI_SKU As String = rowEDTSDQT1.Item("EDI_SKU") & ""
                                        EDI_DTL_SEQ = Val(rowEDTSDQT1.Item("EDI_DTL_SEQ") & "")
                                        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                                        Dim rowEDTSDQT0 As DataRow = dst.Tables("EDTSDQT0").Rows.Find(New String() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ, ITEM_CODE, EDI_UPC})

                                        If SALES_DIVISION_CODE = "?" Then
                                            SALES_DIVISION_CODE = rowICTITEM1.Item("SALES_DIVISION_CODE")
                                            rowSOTORDR1.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
                                        End If

                                        Dim QTY As Int32 = Val(rowEDTSDQT1.Item("QTY") & "")
                                        ' NOTE - rowEDTSDQT0 is Nothing if we skip an item
                                        Dim PRICE As Decimal = Val(rowEDTSDQT0.Item("PRICE") & "")
                                        'If rowEDTSDQT0 IsNot Nothing Then
                                        '    PRICE = Val(rowEDTSDQT0.Item("PRICE") & "")
                                        'End If
                                        ORDR_LNO = ORDR_LNO + 1

                                        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
                                        With rowSOTORDR2
                                            .Item("ORDR_NO") = ORDR_NO
                                            .Item("ORDR_LNO") = ORDR_LNO
                                            .Item("ITEM_CODE") = ITEM_CODE
                                            .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                                            .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                                            .Item("ORDR_UNIT_PRICE") = rowEDTSDQT0.Item("PRICE")
                                            .Item("ORDR_QTY") = QTY
                                            .Item("ORDR_QTY_OPEN") = QTY
                                            .Item("ORDR_QTY_ORIG") = QTY
                                            .Item("ORDR_STATUS") = "O"
                                            .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                            .Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                                            .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                                            .Item("SELL_CODE") = rowSOTORDR1.Item("SELL_CODE")
                                            .Item("CUST_CODE") = rowSOTORDR1.Item("CUST_CODE") ' CUST_CODE_multi
                                            .Item("CUST_STORE_NO") = rowSOTORDR1.Item("CUST_STORE_NO") ' CUST_STORE_NO
                                            .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE") ' WHSE_CODE
                                            .Item("ORDR_UNIT_PRICE_CURR") = rowEDTSDQT0.Item("ITEM_PRICE_CURR")

                                            If ASCMAIN1.CLIENT = "AHA" Then
                                                If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then

                                                    If CUST_CODE <> "SEPHORACAN" Then
                                                        MsgBox("Pricing in Foreign Currency was built around SEPHORACAN rules = needs to be reviewed for " & CUST_CODE)
                                                        Stop
                                                    End If

                                                    If PRICE_BASIS = "L' THEN" Then
                                                        ' NOT SURE WHY WE NEED THIS FOR L AND DON'T NEED IT FOR R
                                                        Dim ORDR_UNIT_PRICE_CURR As Decimal = rowEDTSDQT0.Item("ITEM_PRICE_CURR")
                                                        ORDR_UNIT_PRICE_CURR = rowEDTSDQT0.Item("PRICE")
                                                        .Item("ORDR_UNIT_PRICE_CURR") = ORDR_UNIT_PRICE_CURR
                                                        .Item("ORDR_UNIT_PRICE") = ORDR_UNIT_PRICE_CURR * CURR_EXCH_RATE
                                                    End If
                                                End If
                                            End If

                                            If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                                                Dim ITEM_RETAIL_PRICE_CURR As Decimal = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
                                                Dim rowSOTPRIC2 As DataRow = LookUp("SOTPRIC2", New String() {CURR_CODE, ITEM_CODE})
                                                ITEM_RETAIL_PRICE_CURR = 0
                                                If rowSOTPRIC2 IsNot Nothing Then
                                                    ITEM_RETAIL_PRICE_CURR = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                                                End If
                                                .Item("ITEM_RETAIL_PRICE_CURR") = ITEM_RETAIL_PRICE_CURR
                                            Else
                                                .Item("ORDR_UNIT_PRICE_CURR") = .Item("ORDR_UNIT_PRICE")
                                                .Item("ITEM_RETAIL_PRICE_CURR") = .Item("ITEM_RETAIL_PRICE")
                                            End If
                                        End With
                                        dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                                        If rowSOTORDR1.Item("COLLECTION_CODE") & "" = "" Then
                                            rowSOTORDR1.Item("COLLECTION_CODE") = rowEDTSDQT1.Item("COLLECTION_CODE")
                                            rowSOTORDR1.Item("BRAND_CODE") = rowEDTSDQT1.Item("BRAND_CODE")
                                        End If
                                    Next
                                Next


                                ORDR_GROUP_NOs.Add(ORDR_GROUP_NO)
                                ' ISSUE-7373 - this collection drives what orders are put through the amazon splitting stored procedure, which NF no longer wants. 
                                ' Likely can remove the code that sets this boolean, initializes the list of strings, and uses it as well.
                                'If blnAMZN_AND_ADS Then
                                '    ORDR_GROUP_NOs_ADS.Add(ORDR_GROUP_NO)
                                'End If

                            Catch ex As Exception
                                ' 02/07/2025 - Need to exit the Next since SOTORDR5 records were not created for a skipped invalid Store no
                                MessageBox.Show($"Error Occurred in EDI Import: Customer {CUST_CODE}, Customer PO: {ORDR_CUST_PO} - {ex.Message}", "Import Orders", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                rowEDT850T1.Item("RESULT") = "Rejected"
                                EDI_DOC_SEQ_NO_ok = False
                                Exit For

                                'MsgBox(ex.Message, vbOKOnly, "Error Occurred in EDI Import")
                                'Rollback()
                            End Try
                        End If
                    Else
                        rowEDT850T1.Item("RESULT") = "Rejected"
                        EDI_DOC_SEQ_NO_ok = False
                        Exit For
                    End If

                Next


                If EDI_DOC_SEQ_NO_ok Then

                    If skipped_items.Count <> 0 Then
                        MsgBox("Items Skipped: " & Join(skipped_items.ToArray, ","), vbOKOnly, "Verification")
                    End If

                    ORDERS_IMPORTED += 1

                    BeginTrans()

                    dst.Tables("EDT850T1").AcceptChanges()
                    rowEDT850T1.Item("EDI_PROCESS_IND") = "1"
                    rowEDT850T1.Item("LAST_DATE") = DATETIME_STAMP
                    rowEDT850T1.Item("LAST_OPER") = ASCMAIN1.USER_ID
                    Update_Record_TDA("EDT850T1")

                    'ASCMAIN1.sql = "Update EDT850T1" _
                    '    & " Set EDI_PROCESS_IND = '1' where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    'ASCDATA1.ExecuteSQL()

                    Dim addl_log As String = ""
                    If rowEDT850T1.Item("CUST_CODE_OVERRIDE") & "" <> "" Then
                        addl_log &= ", Customer " & rowEDT850T1.Item("CUST_CODE_OVERRIDE")
                    End If
                    If rowEDT850T1.Item("PRICE_LIST_CODE_OVERRIDE") & "" <> "" Then
                        addl_log &= ", Price List " & rowEDT850T1.Item("PRICE_LIST_CODE_OVERRIDE")
                    End If
                    DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                    Write_Event_Log("EDT850T1", EDI_DOC_SEQ_NO, "Imported" & addl_log)

                    Update_Record_TDA("SOTORDR1")
                    Update_Record_TDA("SOTORDR2")

                    rowEDT850T1.Item("RESULT") = "Imported"
                    rowEDT850T1.Item("ORDERS") = dst.Tables("SOTORDR1").Select("").Length  ' order_count

                    Dim AMOUNT As Decimal = Val(dst.Tables("SOTORDR2").Compute("SUM(ORDR_AMT)", "") & "")

                    Dim row As DataRow = dst.Tables("EDT850T1_IMPORTED").Rows.Find(EDI_DOC_SEQ_NO)
                    If row IsNot Nothing Then
                        row.Delete()
                    End If

                    rowEDT850T1.Item("AMOUNT") = AMOUNT
                    dst.Tables("EDT850T1_IMPORTED").Rows.Add(rowEDT850T1.ItemArray)

                    For Each ORDR_GROUP_NO As String In ORDR_GROUP_NOs

                        Update_ICTSTAT2(ORDR_GROUP_NO, 1)
                        ASCDATA1.ExecuteSQL("Begin SOPORDR0_G('" & ORDR_GROUP_NO & "'); End;")

                        ASCMAIN1.sql = "Insert into SOTORDR5 " & vbCrLf _
                            & " Select SOTORDR1.ORDR_NO, 'ST'" & vbCrLf _
                            & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_ADDR1, ARTCUST2.CUST_STORE_ADDR2, ARTCUST2.CUST_STORE_ADDR3" & vbCrLf _
                            & ", ARTCUST2.CUST_STORE_CITY, ARTCUST2.CUST_STORE_STATE, ARTCUST2.CUST_STORE_ZIP_CODE" & vbCrLf _
                            & ", ARTCUST2.CUST_STORE_COUNTRY, ARTCUST2.CUST_STORE_CONTACT, ARTCUST2.CUST_STORE_PHONE" & vbCrLf _
                            & ", ARTCUST2.CUST_STORE_EXT,ARTCUST2.CUST_STORE_FAX, ARTCUST2.CUST_STORE_EMAIL" & vbCrLf _
                            & ", ARTCUST2.CUST_STORE_NO CUST_ADDR_CODE" & vbCrLf _
                            & " from ARTCUST2, SOTORDR1" & vbCrLf _
                            & " where ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                            & "   and ARTCUST2.CUST_STORE_NO = SOTORDR1.CUST_" & IIf(ORDR_SHIP_TO = "MK", "STORE", "DC") & "_NO" & vbCrLf _
                            & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                            & "   and SOTORDR1.ORDR_STATUS = 'O'"
                        ASCDATA1.ExecuteSQL()
                    Next


                    'ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, 'ST'" & vbCrLf _
                    '    & ", ARTCUST2.CUST_STORE_NAME, ARTCUST2.CUST_STORE_ADDR1, ARTCUST2.CUST_STORE_ADDR2, ARTCUST2.CUST_STORE_ADDR3" & vbCrLf _
                    '    & ", ARTCUST2.CUST_STORE_CITY, ARTCUST2.CUST_STORE_STATE, ARTCUST2.CUST_STORE_ZIP_CODE" & vbCrLf _
                    '    & ", ARTCUST2.CUST_STORE_COUNTRY, ARTCUST2.CUST_STORE_CONTACT, ARTCUST2.CUST_STORE_PHONE" & vbCrLf _
                    '    & ", ARTCUST2.CUST_STORE_EXT,ARTCUST2.CUST_STORE_FAX, ARTCUST2.CUST_STORE_EMAIL" & vbCrLf _
                    '    & ", ARTCUST2.CUST_STORE_NO CUST_ADDR_CODE" & vbCrLf _
                    '    & " from ARTCUST2, SOTORDR1" & vbCrLf _
                    '    & " where ARTCUST2.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                    '    & "   and ARTCUST2.CUST_STORE_NO = SOTORDR1.CUST_" & IIf(ORDR_SHIP_TO = "MK", "STORE", "DC") & "_NO" & vbCrLf _
                    '    & "   and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" & vbCrLf _
                    '    & "   and SOTORDR1.ORDR_STATUS = 'O'"
                    'Dim TBLTEST As DataTable = ASCDATA1.GetDataTable


                    Dim ORDR_CLNO As Integer = 0
                    For Each rowEDT850T4 As DataRow In dst.Tables("EDT850T4").Select("")
                        ORDR_CLNO += 1
                        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                            Dim rowSOTORDR4 As DataRow = dst.Tables("SOTORDR4").NewRow
                            rowSOTORDR4.Item("ORDR_NO") = ORDR_NO
                            rowSOTORDR4.Item("ORDR_CLNO") = ORDR_CLNO
                            rowSOTORDR4.Item("ORDR_COMMENT") = rowEDT850T4.Item("EDI_CMMNT")
                            dst.Tables("SOTORDR4").Rows.Add(rowSOTORDR4)
                        Next
                    Next

                    Update_Record_TDA("SOTORDR4")

                    'OUR CODE CAN LOOP THROUGH ORDR GROUP NOS HERE
                    'SHOULD WE PUT IT ABOVE THE SOTORDR4 STUFF?
                    For Each ORDR_GROUP_NO_ADS As String In ORDR_GROUP_NOs_ADS
                        Update_ICTSTAT2(ORDR_GROUP_NO_ADS, -1)
                        ASCDATA1.ExecuteSQL("Begin SOPORDRX_SPLIT_ADS('" & ORDR_GROUP_NO_ADS & "'); End;")
                        Update_ICTSTAT2(ORDR_GROUP_NO_ADS, 1)
                        ASCDATA1.ExecuteSQL("Begin SOPORDR0_G('" & ORDR_GROUP_NO_ADS & "'); End;")
                    Next

                    CommitTrans()

                    ASCDATA1.DeleteRows(dst.Tables("EDT850TE"), "EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")

                Else


                End If



            End If
        Next

        ORDERS_REJECTED = ORDERS_PROCESSED - ORDERS_IMPORTED
        MsgBox(CStr(ORDERS_PROCESSED) & " Order(s) Processed, " & vbCrLf &
               CStr(ORDERS_IMPORTED) & " Order(s) Imported, " & vbCrLf &
               CStr(ORDERS_REJECTED) & " Order(s) Rejected", MsgBoxStyle.OkOnly, "Processing Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Bad_Data(
        EDI_COND_DESC As String,
        EDI_COND_CODE As String,
        EDI_RECEIVED_VALUE As String,
        Optional EDI_EXPECTED_VALUE As String = "",
        Optional EDI_REFERENCE As String = "") As String

        Dim EDI_REPLACED_VALUE As String = ""
        EDI_ACTION = ""

        Dim EDI_ERROR As String = EDI_COND_DESC & vbTab & EDI_COND_CODE & vbTab & EDI_RECEIVED_VALUE
        If EDI_ERRORs.ContainsKey(EDI_ERROR) Then
            Dim EDI_ERROR_NO As Int32 = EDI_ERRORs(EDI_ERROR)
            Dim rowEDT850TE As DataRow = dst.Tables("EDT850TE").Rows.Find(New Object() {EDI_DOC_SEQ_NO, EDI_ERROR_NO})

            EDI_ACTION = rowEDT850TE.Item("EDI_ACTION") & ""
            Select Case EDI_ACTION
                Case "A" ' Use Received Value (ie, Accept EDI Data)
                    If ACTIONS_A.Contains("*") Or ACTIONS_A.Contains(EDI_COND_CODE) Then
                        EDI_REPLACED_VALUE = EDI_RECEIVED_VALUE
                    Else
                        EDI_ACTION = "N"
                    End If

                Case "E" ' Use Expected Value
                    If ACTIONS_E.Contains("*") Or ACTIONS_E.Contains(EDI_COND_CODE) Then

                        If ASCMAIN1.DBS_COMPANY = "SLP" Then
                            EDI_REPLACED_VALUE = EDI_EXPECTED_VALUE
                        Else
                            If ASCMAIN1.CLIENT = "INT" AndAlso ((EDI_COND_CODE = "57" Or EDI_COND_CODE = "17") AndAlso EDI_EXPECTED_VALUE <> 0) Then
                                ' DO NOT PERMIT USE EXPECTED VALUE AS PER LM - OTHERWISE CUSTOMER (IE AMAZON) WILL CHARGE BACK
                                ' EXCEPT FOR WHEN A RETAILER SENDS .01 (LIKE LORD) FOR NC ITEMS
                            Else
                                EDI_REPLACED_VALUE = EDI_EXPECTED_VALUE
                            End If
                        End If

                    Else
                        EDI_ACTION = "N"
                    End If

                Case "S" ' Skip Record
                    If ACTIONS_S.Contains("*") Or ACTIONS_S.Contains(EDI_COND_CODE) Then
                        EDI_REPLACED_VALUE = "Skip"
                    Else
                        EDI_ACTION = "N"
                    End If

                Case "R" ' Use Replacement Value
                    If ACTIONS_R.Contains("*") Or ACTIONS_R.Contains(EDI_COND_CODE) Then
                        EDI_REPLACED_VALUE = rowEDT850TE.Item("EDI_REPLACED_VALUE") & ""
                    Else
                        EDI_ACTION = "N"
                    End If
            End Select

            If EDI_REPLACED_VALUE = "" Then
                rowEDT850TE.Item("EDI_ORDER_COUNT") = Val(rowEDT850TE.Item("EDI_ORDER_COUNT") & "") + 1
            End If
        Else
            Dim rowEDT850TE As DataRow = dst.Tables("EDT850TE").NewRow
            Dim EDI_ERROR_NO As Int32 = EDI_ERRORs.Count + 1
            EDI_ERRORs.Add(EDI_ERROR, EDI_ERROR_NO)
            With rowEDT850TE
                .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                .Item("EDI_ERROR_NO") = EDI_ERROR_NO
                .Item("EDI_COND_DESC") = EDI_COND_DESC
                .Item("EDI_COND_CODE") = EDI_COND_CODE
                .Item("EDI_EXPECTED_VALUE") = EDI_EXPECTED_VALUE
                .Item("EDI_RECEIVED_VALUE") = EDI_RECEIVED_VALUE
                .Item("EDI_ACTION") = "N"
                .Item("EDI_ORDER_COUNT") = 1
                .Item("EDI_TP_QUAL") = EDI_TP_QUAL
                .Item("EDI_TP_ID") = EDI_TP_ID
                .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = CUST_STORE_NO
                .Item("EDI_REFERENCE") = EDI_REFERENCE
                If RESOLUTIONS.ContainsKey(EDI_COND_CODE) Then
                    .Item("RESOLUTION") = RESOLUTIONS(EDI_COND_CODE)
                End If

            End With
            dst.Tables("EDT850TE").Rows.Add(rowEDT850TE)
        End If

        If EDI_REPLACED_VALUE = "" Then
            EDI_DOC_SEQ_NO_ok = False
        End If

        Return EDI_REPLACED_VALUE
    End Function

    Function Get_EDTTRPM1() As DataRow
        ORDR_EDI_810 = "0"
        ORDR_EDI_856 = "0"
        CUST_EDI_DTS_FLAG = "0"

        rowEDTTRPM1 = dst.Tables("EDTTRPM1").Rows.Find(New String() {EDI_TP_QUAL, EDI_TP_ID, "850"})
        If rowEDTTRPM1 Is Nothing Then
            Bad_Data(EDI_COND_DESC:="Unrecognized Sender ID", EDI_COND_CODE:="01", EDI_RECEIVED_VALUE:=EDI_TP_QUAL & "-" & EDI_TP_ID)
        Else
            rowEDTSLSP1 = LookUp("EDTSLSP1", CUST_CODE, True)
            If rowEDTSLSP1 Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Unrecognized Sender ID", EDI_COND_CODE:="01", EDI_RECEIVED_VALUE:=EDI_TP_QUAL & "-" & EDI_TP_ID)
            Else
                ORDR_EDI_810 = IIf(rowEDTSLSP1.Item("EDI_ID_810") & "" = "", "0", "1")
                ORDR_EDI_856 = IIf(rowEDTSLSP1.Item("EDI_ID_856") & "" = "", "0", "1")
                CUST_EDI_DTS_FLAG = rowEDTSLSP1.Item("EDI_DTS_IND") & ""
            End If
        End If

        Return rowEDTTRPM1
    End Function

    Sub Get_ARTCUST1()

        ' WHAT ABOUT CUST_STATUS?

        CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        CURR_EXCH_RATE = 1

        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
        If rowARTCUST1 Is Nothing Then
            Bad_Data(EDI_COND_DESC:="Customer Record Missing or Not Active", EDI_COND_CODE:="02", EDI_RECEIVED_VALUE:=CUST_CODE)
        Else
            If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                Bad_Data(EDI_COND_DESC:="Customer PO is Required, but Missing", EDI_COND_CODE:="04", EDI_RECEIVED_VALUE:=ORDR_CUST_PO)
            End If
            If rowARTCUST1.Item("FRT_TERMS") & "" = "" Then
                Bad_Data(EDI_COND_DESC:="Customer Missing Freight Terms", EDI_COND_CODE:="35", EDI_RECEIVED_VALUE:="")
            End If

            If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                Dim SHIP_VIA_CODE As String = rowARTCUST1.Item("SHIP_VIA_CODE")

                Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                If rowSOTSVIA1 Is Nothing Then
                    Bad_Data(EDI_COND_DESC:="Customer Missing Valid Ship Via Code", EDI_COND_CODE:="44", EDI_RECEIVED_VALUE:="")
                Else
                    If rowSOTSVIA1.Item("FRT_TERMS") & "" <> "" Then
                        If rowSOTSVIA1.Item("FRT_TERMS") & "" <> rowARTCUST1.Item("FRT_TERMS") & "" Then
                            Bad_Data(EDI_COND_DESC:="Freight Terms associated with Ship Via Code specified is " & rowSOTSVIA1.Item("FRT_TERMS"), EDI_COND_CODE:="45", EDI_RECEIVED_VALUE:="")
                        End If
                    End If
                End If
            End If

            If rowARTCUST1.Item("CUST_ROUTING_INST") & "" = "" Then
                Bad_Data(EDI_COND_DESC:="Customer Missing Routing Instructions", EDI_COND_CODE:="36", EDI_RECEIVED_VALUE:="")
            End If

            WHSE_CODE_customer = ""
            If rowARTCUST1.Item("WHSE_CODE") & "" = "" Then
                WHSE_CODE = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""
            Else
                WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
                WHSE_CODE_customer = WHSE_CODE
            End If

            CUST_BILL_TO_CUST = rowARTCUST1.Item("CUST_BILL_TO_CUST") & ""
            If CUST_BILL_TO_CUST = "" Then
                CUST_BILL_TO_CUST = CUST_CODE
            End If

            rowARTCUST1bt = LookUp("ARTCUST1", CUST_BILL_TO_CUST)
            If rowARTCUST1bt Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Bill-To Customer Missing or not Active", EDI_COND_CODE:="03", EDI_RECEIVED_VALUE:=CUST_BILL_TO_CUST)
            End If

            rowSOTPCLS1 = dst.Tables("SOTPCLS1").Rows.Find(rowARTCUST1.Item("PRICE_CLASS_CODE") & "")
            If rowSOTPCLS1 Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Missing or Invalid Price Class Code for Customer " & CUST_CODE, EDI_COND_CODE:="91", EDI_RECEIVED_VALUE:=rowARTCUST1.Item("PRICE_CLASS_CODE") & "")
            Else
                PRICE_BASIS = rowSOTPCLS1.Item("PRICE_BASIS") & ""
                PRICE_BASE_DPCT = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
            End If

            PRICE_LIST_CODE = rowARTCUST1.Item("PRICE_LIST_CODE") & ""

            CURR_CODE = rowARTCUST1.Item("CURR_CODE") & ""
            If CURR_CODE = "" Then CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                CURR_EXCH_RATE = 1
            Else
                CURR_EXCH_RATE = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, Now.Date)

            End If
        End If
    End Sub

    Function Get_Ship_To(EDI_STORE As String, EDI_SHIP_DC As String, EDI_ADDR_CODE As String) As String

        Dim skip_store As Boolean = False
        Dim Bad_Data_Cond_05 As Boolean = False

        If EDI_STORE <> "" Then
            If Len(EDI_STORE) < 6 And IsNumeric(EDI_STORE) Then
                CUST_STORE_NO = Format(Val(EDI_STORE), "000000")
            Else
                CUST_STORE_NO = EDI_STORE
            End If
        Else
            CUST_STORE_NO = ""
        End If
        If Len(CUST_STORE_NO) > 6 Then
            CUST_STORE_NO = ""
        End If
        'CUST_STORE_NAME = ""
        CUST_DC_NO = ""

        rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE_multi, CUST_STORE_NO})

        Dim CUST_DC_NO_GLN As String = ""
        If rowARTCUST2 Is Nothing And EDI_SHIP_DC <> "" Then
            Dim row() As DataRow = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_multi & "' AND GLOBAL_LOCATION_NUMBER = '" & EDI_SHIP_DC & "'")
            If row IsNot Nothing AndAlso row.Length > 0 Then
                rowARTCUST2 = row(0)
                CUST_DC_NO_GLN = rowARTCUST2.Item("CUST_STORE_NO")
            End If
        End If

        If EDI_STORE <> "" Then
            Dim row() As DataRow = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_multi & "' AND GLOBAL_LOCATION_NUMBER = '" & EDI_STORE & "'")
            If row IsNot Nothing AndAlso row.Length > 0 Then
                rowARTCUST2 = row(0)
            End If
        End If

        If rowARTCUST2 Is Nothing Then
            Dim row() As DataRow = Nothing
            If EDI_STORE <> "" Then
                row = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_multi & "' AND GLOBAL_LOCATION_NUMBER = '" & EDI_STORE & "'")
            Else
                row = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_multi & "' AND GLOBAL_LOCATION_NUMBER = '" & EDI_ADDR_CODE & "'")
            End If
            If row.Length > 0 Then
                rowARTCUST2 = row(0)
            End If
        End If

        If rowARTCUST2 Is Nothing And EDI_SHIP_DC <> "" Then
            If EDI_STORE <> "" Then
                ' Get_Ship_To is called at the EDI header level, and also for each SDQ.
                ' when called at the SDQ level, or even at the header level if the EDI document has a store defined,
                ' if we cannot resolve the store, it is important to leave the error condition
                ' and not to simply use the DC as a stand-in for the store, which is what the ELSE branch below will do.
                ' So the If statement above is meant to preserve the error if the present call has a value for EDI_STORE.
            Else
                rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE_multi, EDI_SHIP_DC})
                If rowARTCUST2 Is Nothing Then

                    ' Dim EDI_SHIP_DC_PADDED As String = Format(Val(EDI_SHIP_DC), "000000")
                    ' THE PRECEDING LINE TOOK DC T303 AND MADE IT 000000 - NO GOOD
                    Dim EDI_SHIP_DC_PADDED As String = EDI_SHIP_DC.PadLeft(6, "0")
                    rowARTCUST2 = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE_multi, EDI_SHIP_DC_PADDED})
                End If
            End If
        End If

        If rowARTCUST2 Is Nothing Then
            Bad_Data(EDI_COND_DESC:="Invalid Customer Store", EDI_COND_CODE:="05", EDI_RECEIVED_VALUE:=IIf(EDI_STORE = "", EDI_SHIP_DC, EDI_STORE))
            If EDI_ACTION = "S" Then
                skip_store = True
            Else
                Bad_Data_Cond_05 = True
            End If
        End If

        Dim rowSOTSVIA2 As DataRow
        Dim rowSOTSVIA3 As DataRow

        If Not Bad_Data_Cond_05 And Not skip_store Then

            rowSOTSVIA2 = dst.Tables("SOTSVIA2").Rows.Find _
                (New String() {WHSE_CODE, rowARTCUST2.Item("CUST_STORE_STATE") & ""})

            If rowSOTSVIA2 Is Nothing Then
                TRANSIT_BUS_DAYS = 0
            Else
                TRANSIT_BUS_DAYS = Val(rowSOTSVIA2.Item("TRANSIT_BUS_DAYS") & "")
            End If

            If CUST_EDI_DTS_FLAG <> "1" And rowARTCUST2.Item("CUST_DC_NO") & "" = "" And rowARTCUST2.Item("CUST_DC_IND") & "" <> "1" Then
                Bad_Data(EDI_COND_DESC:="Store " & CUST_STORE_NO & " is Missing DC Code", EDI_COND_CODE:="37", EDI_RECEIVED_VALUE:="")
            End If

            CUST_STORE_NO = rowARTCUST2.Item("CUST_STORE_NO")
            'CUST_STORE_NAME = rowARTCUST2.Item("CUST_STORE_NAME") & ""
            If rowARTCUST2.Item("CUST_DC_IND") & "" = "1" Then
                CUST_DC_NO = rowARTCUST2.Item("CUST_STORE_NO") & ""
            Else
                If IsNumeric(rowARTCUST2.Item("CUST_DC_NO")) Then
                    CUST_DC_NO = Format(Val(rowARTCUST2.Item("CUST_DC_NO") & ""), "000000")
                Else
                    CUST_DC_NO = rowARTCUST2.Item("CUST_DC_NO") & ""
                End If
                If CUST_DC_NO_GLN <> "" And CUST_DC_NO_GLN <> CUST_DC_NO Then
                    Bad_Data(EDI_COND_DESC:="Store " & CUST_STORE_NO & " refers to DC " & CUST_DC_NO & " and Order refers to " & CUST_DC_NO_GLN, EDI_COND_CODE:="41", EDI_RECEIVED_VALUE:="")
                End If
            End If

            If EDI_SHIP_DC <> "" Or CUST_EDI_DTS_FLAG <> "1" Then
                If EDI_SHIP_DC = "" Then
                    EDI_SHIP_DC = CUST_DC_NO
                End If
                If IsNumeric(EDI_SHIP_DC) And Len(EDI_SHIP_DC) <= 6 Then
                    CUST_DC_NO = EDI_SHIP_DC.PadLeft(6, "0")
                    'Else
                    '    CUST_DC_NO = EDI_SHIP_DC
                End If
                ORDR_SHIP_TO = "DC"
            Else
                ORDR_SHIP_TO = "MK"
            End If


            If CUST_DC_NO <> "" And ORDR_SHIP_TO = "DC" Then

                rowARTCUST2_DC = dst.Tables("ARTCUST2").Rows.Find(New String() {CUST_CODE_multi, CUST_DC_NO})
                If rowARTCUST2_DC Is Nothing Then
                    Dim row() As DataRow = Nothing
                    If EDI_SHIP_DC <> "" Then
                        row = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_multi & "' AND GLOBAL_LOCATION_NUMBER = '" & EDI_SHIP_DC & "'")
                    ElseIf EDI_STORE <> "" Then
                        row = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_multi & "' AND GLOBAL_LOCATION_NUMBER = '" & EDI_STORE & "'")
                    Else
                        row = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE_multi & "' AND GLOBAL_LOCATION_NUMBER = '" & EDI_ADDR_CODE & "'")
                    End If
                    If row.Length > 0 Then
                        rowARTCUST2_DC = row(0)
                    End If
                End If

                If rowARTCUST2_DC Is Nothing Then
                    Bad_Data(EDI_COND_DESC:="DC not on file", EDI_COND_CODE:="07", EDI_RECEIVED_VALUE:=CUST_DC_NO)
                    CUST_DC_NO = "000000"
                Else
                    CUST_DC_NO = rowARTCUST2_DC.Item("CUST_STORE_NO") & ""
                    ORDR_SHIP_TO = "DC"
                    rowSOTSVIA2 = dst.Tables("SOTSVIA2").Rows.Find _
                       (New String() {WHSE_CODE, rowARTCUST2_DC.Item("CUST_STORE_STATE") & ""})

                    If rowSOTSVIA2 Is Nothing Then
                        TRANSIT_BUS_DAYS = 0
                    Else
                        TRANSIT_BUS_DAYS = Val(rowSOTSVIA2.Item("TRANSIT_BUS_DAYS") & "")
                    End If
                End If

                ' if there is a DC defined, and we are shipping to DC, then look for an overriding value for SHIP_VIA
                ' ISSUE-7330 (Erin’s ABS is blowing up with a null ref). 01/29/2026
                Dim SHIP_VIA_CODE As String = String.Empty
                If rowARTCUST2_DC IsNot Nothing Then
                    SHIP_VIA_CODE = rowARTCUST2_DC.Item("CUST_STORE_SHIP_VIA_CODE") & ""
                End If

                If SHIP_VIA_CODE <> "" Then
                    Dim rowSOTSVIA1 As DataRow = LookUp("SOTSVIA1", SHIP_VIA_CODE)
                    If rowSOTSVIA1 Is Nothing Then
                        Bad_Data(EDI_COND_DESC:="Customer DC has an invalid Overriding Ship Via Code", EDI_COND_CODE:="44", EDI_RECEIVED_VALUE:="")
                        ' should we be usng code 44 or a new exception?
                    Else
                        If rowSOTSVIA1.Item("FRT_TERMS") & "" <> "" Then
                            If rowSOTSVIA1.Item("FRT_TERMS") & "" <> rowARTCUST1.Item("FRT_TERMS") & "" Then
                                Bad_Data(EDI_COND_DESC:="Freight Terms associated with DC Overriding Ship Via Code specified is " & rowSOTSVIA1.Item("FRT_TERMS"), EDI_COND_CODE:="45", EDI_RECEIVED_VALUE:="")
                            End If
                        End If
                    End If
                End If

            End If

            If CUST_DC_NO = "" And ORDR_SHIP_TO = "DC" Then
                Bad_Data(EDI_COND_DESC:="Store " & CUST_STORE_NO & " not linked to DC", EDI_COND_CODE:="33", EDI_RECEIVED_VALUE:=ORDR_SHIP_TO)
            Else

                If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                Else
                    If TAC.TACMAIN1.IPLBMacysCustomerCodes.Contains(CUST_CODE) Then
                        rowSOTSVIA3 = dst.Tables("SOTSVIA3").Rows.Find _
                      (New String() {WHSE_CODE, CUST_CODE, CUST_DC_NO})
                        If rowSOTSVIA3 Is Nothing Then
                            TRANSIT_BUS_DAYS = 0
                            Bad_Data(EDI_COND_DESC:="DC " & CUST_DC_NO & " is Missing Transit Days", EDI_COND_CODE:="48", EDI_RECEIVED_VALUE:="")
                        Else
                            TRANSIT_BUS_DAYS = Val(rowSOTSVIA3.Item("TRANSIT_BUS_DAYS") & "")
                        End If
                        '   If CUST_DC_NO = "AZ" Then Stop
                    End If

                    'SOTSVIA3
                    'WHSE_CODE                    K VARCHAR2(6)
                    'CUST_CODE                    K VARCHAR2(10)
                    'CUST_STORE_NO                K VARCHAR2(6)
                    'TRANSIT_BUS_DAYS NUMBER

                End If
            End If
        End If

        Return CUST_STORE_NO

    End Function

    Function Get_Terms() As String

        TERM_CODE = ""
        If rowARTCUST1bt Is Nothing Then Return TERM_CODE

        If rowARTCUST1bt.Item("TERM_CODE") & "" = "" Then
            Bad_Data(EDI_COND_DESC:="Bill-To Customer Missing Terms", EDI_COND_CODE:="21", EDI_RECEIVED_VALUE:="")
        Else
            TERM_CODE = rowARTCUST1bt.Item("TERM_CODE") & ""
        End If

        Dim EDI_TERMS_key As String = "" _
            & rowEDT850T1.Item("EDI_TERMS") _
            & rowEDT850T1.Item("EDI_TERM_TYPE") _
            & rowEDT850T1.Item("EDI_TERM_BASIS") _
            & rowEDT850T1.Item("EDI_TERM_RATE") _
            & rowEDT850T1.Item("EDI_TERM_DSCDAYS") _
            & rowEDT850T1.Item("EDI_TERM_NETDAYS") _
            & rowEDT850T1.Item("EDI_TERM_DESC") _
            & rowEDT850T1.Item("EDI_TERM_DOM")

        If EDI_TERMS_key <> "" Then
            Dim EDI_TERM_CODE As String = ""

            Dim rowEDTTERM1 As DataRow
            For Each row As DataRow In dst.Tables("EDTTERM1").Select("")
                If EDI_TERMS_key = "" _
                & row.Item("EDI_TERMS") _
                & row.Item("EDI_TERM_TYPE") _
                & row.Item("EDI_TERM_BASIS") _
                & row.Item("EDI_TERM_RATE") _
                & row.Item("EDI_TERM_DSCDAYS") _
                & row.Item("EDI_TERM_NETDAYS") _
                & row.Item("EDI_TERM_DESC") _
                & row.Item("EDI_TERM_DOM") Then
                    EDI_TERM_CODE = row.Item("EDI_TERM_CODE")
                    TERM_CODE = row.Item("TERM_CODE") & ""
                    Exit For
                End If
            Next

            If TERM_CODE = "" Then
                rowEDTTERM1 = dst.Tables("EDTTERM1").NewRow
                With rowEDTTERM1
                    EDI_TERM_CODE = ASCMAIN1.Next_Control_No("EDTTERM1.EDI_TERM_CODE")
                    .Item("EDI_TERM_CODE") = EDI_TERM_CODE
                    .Item("EDI_TERMS") = rowEDT850T1.Item("EDI_TERMS")
                    .Item("EDI_TERM_TYPE") = rowEDT850T1.Item("EDI_TERM_TYPE")
                    .Item("EDI_TERM_BASIS") = rowEDT850T1.Item("EDI_TERM_BASIS")
                    .Item("EDI_TERM_RATE") = rowEDT850T1.Item("EDI_TERM_RATE")
                    .Item("EDI_TERM_DSCDAYS") = rowEDT850T1.Item("EDI_TERM_DSCDAYS")
                    .Item("EDI_TERM_NETDAYS") = rowEDT850T1.Item("EDI_TERM_NETDAYS")
                    .Item("EDI_TERM_DESC") = rowEDT850T1.Item("EDI_TERM_DESC")
                    .Item("EDI_TERM_DOM") = rowEDT850T1.Item("EDI_TERM_DOM")
                End With
                dst.Tables("EDTTERM1").Rows.Add(rowEDTTERM1)
            End If
            If TERM_CODE = "" Then
                Bad_Data(EDI_COND_DESC:="EDI Terms File", EDI_COND_CODE:="11", EDI_RECEIVED_VALUE:=EDI_TERM_CODE)
            Else
                Dim rowTATTERM1 As DataRow = dst.Tables("TATTERM1").Rows.Find(TERM_CODE)
                If rowTATTERM1 Is Nothing Then
                    Bad_Data(EDI_COND_DESC:="AR Terms File", EDI_COND_CODE:="12", EDI_RECEIVED_VALUE:=TERM_CODE)
                Else
                    If TERM_CODE <> rowARTCUST1bt.Item("TERM_CODE") & "" Then
                        TERM_CODE = Bad_Data(EDI_COND_DESC:="EDI Terms do not match Bill-To Customer", EDI_COND_CODE:="13", EDI_RECEIVED_VALUE:=TERM_CODE, EDI_EXPECTED_VALUE:=rowARTCUST1bt.Item("TERM_CODE") & "")
                    End If
                End If
            End If
        End If

        Return TERM_CODE
    End Function

    Sub Check_for_Possible_Order_Duplication(Optional CUST_STORE_NO As String = "")
        If ORDR_CUST_PO <> "" Then
            ASCMAIN1.sql = "Select MAX (ORDR_NO) from SOTORDR1 where CUST_CODE = :PARM1 and ORDR_CUST_PO = :PARM2" _
                & " and ORDR_STATUS in ('O','P','C','F')"
            Dim ORDR_NO As String
            If CUST_STORE_NO = "" Then
                ORDR_NO = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {CUST_CODE, ORDR_CUST_PO})
            Else
                ASCMAIN1.sql &= " and CUST_STORE_NO = :PARM3"
                ORDR_NO = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VVV", New String() {CUST_CODE, ORDR_CUST_PO, CUST_STORE_NO})
            End If
            If ORDR_NO <> "" Then
                Bad_Data(EDI_COND_DESC:="Possible Order Duplication", EDI_COND_CODE:="14", EDI_RECEIVED_VALUE:=ORDR_CUST_PO)
            End If
        End If
    End Sub

    Sub Process_EDT850T2(EDI_STORE As String, EDI_SHIP_DC As String, EDI_ADDR_CODE As String)

        ' Loop thru 850T2 - each record will represent a Single Item/Style/UPC, etc

        'sdq_stores.Clear()

        Dim EDI_SHIP_DC_passed_in As String = EDI_SHIP_DC
        For Each rowEDT850T2 In dst.Tables("EDT850T2").Select("")
            ' Dim skip_item As Boolean = False
            skip_item = False
            EDI_SHIP_DC = EDI_SHIP_DC_passed_in
            If EDI_SHIP_DC = "" And rowEDT850T2.Item("EDI_LN_SHIP_DC") & "" <> "" Then
                EDI_SHIP_DC = rowEDT850T2.Item("EDI_LN_SHIP_DC")
            End If
            Dim EDI_DETL_QTY As Int64
            Dim EDI_PRICE As Decimal
            Dim EDI_PRICE_CURR As Decimal

            Dim ASST_ITEM_CODE As String
            Dim ASST_SKU As String
            Dim ASST_UPC As String

            Dim T2_PO4_QTY As Int32
            Dim T6_PO4_QTY As Int32
            Dim PO4_OW As Int32

            Dim EDI_PRICE_UOM As String = rowEDT850T2.Item("EDI_PRICE_UOM") & ""
            Dim use_case_pack As Boolean = False

            EDI_DTL_SEQ = Val(rowEDT850T2.Item("EDI_DTL_SEQ") & "")
            Dim rowEDT850T6s() As DataRow = dst.Tables("EDT850T6").Select("EDI_DTL_SEQ = " & CStr(EDI_DTL_SEQ))
            If rowEDT850T6s.Length <> 0 Then
                For Each rowEDT850T6 As DataRow In rowEDT850T6s
                    With rowEDT850T6
                        If .Item("EDI_PO4_QTY") & "" = "" Then
                            T6_PO4_QTY = 1
                        Else
                            If .Item("EDI_PO4_INNER") & "" = "" Then
                                T6_PO4_QTY = Val(.Item("EDI_PO4_QTY"))
                            Else
                                T6_PO4_QTY = Val(.Item("EDI_PO4_QTY")) * Val(.Item("EDI_PO4_INNER"))
                            End If
                        End If

                        If EDI_PRICE_UOM = "CA" Or EDI_PRICE_UOM = "CS" Then
                            If T6_PO4_QTY = 1 Then
                                use_case_pack = True
                            End If
                        End If

                        Dim EDI_UPC As String = Trim(.Item("EDI_SLN_UPC") & "")
                        Dim EDI_EAN As String = Trim(.Item("EDI_SLN_EAN") & "")
                        Dim EDI_SKU As String = Trim(.Item("EDI_SLN_SKU") & "")
                        Dim EDI_ITEM As String = Trim(.Item("EDI_SLN_ITEM") & "")
                        Dim EDI_STYLE As String = Trim(.Item("EDI_SLN_STYLE") & "")

                        PO4_OW = T6_PO4_QTY
                        PO4_OverWrite(EDI_UPC, EDI_SKU, EDI_ITEM, PO4_OW, T2_PO4_QTY, T6_PO4_QTY)
                        Dim EDI_SIZE_DESC As String = ""
                        If .Item("EDI_SLN_ITEM") & "" <> "" Then
                            EDI_SIZE_DESC = .Item("EDI_SLN_SIZE_DESC") & ""
                        Else
                            EDI_SIZE_DESC = rowEDT850T2.Item("EDI_SIZE_DESC") & ""
                        End If
                        If rowEDTSLSP1.Item("EDI_SLN_TOT_IND") & "" = "1" Then
                            EDI_DETL_QTY = .Item("EDI_SLN_QTY")
                            EDI_PRICE = Val(.Item("EDI_SLN_PRICE") & "")
                            EDI_PRICE_CURR = EDI_PRICE
                        Else
                            EDI_DETL_QTY = .Item("EDI_SLN_QTY") * T6_PO4_QTY
                            EDI_PRICE = Val(.Item("EDI_SLN_PRICE") & "") / T6_PO4_QTY
                            EDI_PRICE_CURR = EDI_PRICE
                        End If

                        ASST_ITEM_CODE = ""
                        ASST_SKU = ""
                        ASST_UPC = ""

                        Process_Items_Pre(EDI_UPC, EDI_EAN, EDI_SKU, EDI_ITEM, EDI_STYLE, EDI_DETL_QTY, EDI_STORE, EDI_SHIP_DC, EDI_ADDR_CODE, use_case_pack, EDI_PRICE, EDI_PRICE_CURR)

                    End With
                Next
            Else

                If rowEDT850T2.Item("EDI_PO4_QTY") & "" = "" Then
                    T2_PO4_QTY = 0
                Else
                    If rowEDT850T2.Item("EDI_PO4_INNER") & "" = "" Then
                        T2_PO4_QTY = Val(rowEDT850T2.Item("EDI_PO4_QTY")) & ""
                    Else
                        T2_PO4_QTY = Val(rowEDT850T2.Item("EDI_PO4_QTY")) * Val(rowEDT850T2.Item("EDI_PO4_INNER"))
                    End If
                End If

                If EDI_PRICE_UOM = "CA" Or EDI_PRICE_UOM = "CS" Then
                    If T2_PO4_QTY = 0 Then
                        use_case_pack = True
                    End If
                End If

                Dim EDI_GTIN As String = Trim(rowEDT850T2.Item("EDI_GTIN") & "")
                Dim EDI_UPC As String = Trim(rowEDT850T2.Item("EDI_UPC") & "")
                Dim EDI_EAN As String = Trim(rowEDT850T2.Item("EDI_EAN") & "")
                Dim EDI_STYLE As String = Trim(rowEDT850T2.Item("EDI_STYLE") & "")

                If EDI_GTIN <> "" Then
                    Dim rowICTGTIN1 As DataRow = LookUp("ICTGTIN1", EDI_GTIN)
                    If rowICTGTIN1 IsNot Nothing Then
                        EDI_UPC = rowICTGTIN1.Item("GTIN_ROOT_UPC") & rowICTGTIN1.Item("GTIN_ROOT_EAN") & ""
                    End If
                End If

                If EDI_GTIN <> "" And EDI_UPC = "" Then
                    If Len(EDI_GTIN) = 14 And EDI_GTIN.StartsWith("00") Then
                        EDI_UPC = Mid(EDI_GTIN, 3)
                    End If
                End If
                If EDI_GTIN <> "" And EDI_UPC = "" And EDI_EAN = "" Then
                    If Len(EDI_GTIN) = 14 And EDI_GTIN.StartsWith("0") Then
                        EDI_EAN = Mid(EDI_GTIN, 2)
                    End If
                End If

                Dim EDI_SKU As String = Trim(rowEDT850T2.Item("EDI_SKU") & "")
                Dim EDI_ITEM As String = Trim(rowEDT850T2.Item("EDI_ITEM") & "")
                Dim EDI_SIZE_DESC As String = rowEDT850T2.Item("EDI_SIZE_DESC") & ""
                PO4_OW = T2_PO4_QTY
                PO4_OverWrite(EDI_UPC, EDI_SKU, EDI_ITEM, PO4_OW, T2_PO4_QTY, T6_PO4_QTY)
                ' If rowEDT850T2.Item("EDI_PO4_UOM") & "" = "EA" Then
                If rowEDT850T2.Item("EDI_PO4_UOM") & "" = "EA" Or rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "EA" Then
                    'If rowEDTSLSP1.Item("EDI_PO1_TOT_IND") & "" = "1" And rowEDT850T2.Item("EDI_PO4_UOM") & "" = "EA" Then
                    EDI_DETL_QTY = rowEDT850T2.Item("EDI_TOTAL_QTY")
                    EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "")
                    EDI_PRICE_CURR = EDI_PRICE
                Else
                    ' Notice bellow is checking T2_PO4_QTY, should it check PO1_TOT_IND? or T2_PO4_QTY <> 1?
                    'If T2_PO4_QTY = 1 And (rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CA" Or rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CS") Then
                    '    EDI_PRICE = rowEDT850T2.Item("EDI_PRICE") / T2_PO4_QTY
                    'Else
                    If T2_PO4_QTY = 0 Then
                        EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "")
                    Else
                        EDI_PRICE = Val(rowEDT850T2.Item("EDI_PRICE") & "") / T2_PO4_QTY
                    End If
                    EDI_PRICE_CURR = EDI_PRICE
                    'End If
                    If dst.Tables("EDT850T3").Select("EDI_DTL_SEQ = " & CStr(EDI_DTL_SEQ)).Length > 0 Then
                        If T2_PO4_QTY <> 0 Then
                            EDI_DETL_QTY = T2_PO4_QTY
                        Else
                            EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "")
                        End If
                    Else
                        If T2_PO4_QTY <> 0 Then
                            EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "") * T2_PO4_QTY
                        Else
                            EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "")
                        End If
                    End If
                End If

                ASST_ITEM_CODE = ""
                ASST_SKU = ""
                ASST_UPC = ""
                Process_Items_Pre(EDI_UPC, EDI_EAN, EDI_SKU, EDI_ITEM, EDI_STYLE, EDI_DETL_QTY, EDI_STORE, EDI_SHIP_DC, EDI_ADDR_CODE, use_case_pack, EDI_PRICE, EDI_PRICE_CURR)
            End If
        Next
    End Sub

    Sub Process_Items(
        EDI_UPC As String,
        EDI_EAN As String,
        EDI_SKU As String,
        EDI_ITEM As String,
        EDI_STYLE As String,
        EDI_PRICE As Decimal,
        EDI_PRICE_CURR As Decimal)

        ITEM_INACTIVE = False

        Dim PRICE As Decimal = 0
        Dim CUST_PRICE As Decimal = 0
        'Dim EDI_PRICE As Decimal = 0
        'Dim EDI_PRICE_CURR As Decimal = 0
        Dim ITEM_PRICE_CURR As Decimal = 0
        Dim RETAIL_PRICE As Decimal
        Dim PRICE_DISC As Decimal

        Dim sqlw As String = String.Format("CUST_CODE = '{0}' and EDI_UPC = '{1}' and EDI_SKU = '{2}' and EDI_ITEM = '{3}' and EDI_EAN = '{4}'", CUST_CODE, EDI_UPC, EDI_SKU, EDI_ITEM, EDI_EAN)
        Dim rowEDTITEMXs() As DataRow = dst.Tables("EDTITEMX").Select(sqlw)
        If rowEDTITEMXs.Length <> 0 And (EDI_UPC <> "" Or EDI_SKU <> "" Or EDI_ITEM <> "" Or EDI_EAN <> "") Then
            With rowEDTITEMXs(0)
                CUST_PRICE = Val(.Item("CUST_PRICE") & "")
                ITEM_CODE = .Item("ITEM_CODE")
                Dim SALES_DIVISION_CODE As String = .Item("SALES_DIVISION_CODE") & ""
                rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(.Item("ITEM_CODE"))
                If rowICTITEM1 Is Nothing Then
                    Bad_Data(EDI_COND_DESC:="Item Not in Work Table",
                              EDI_COND_CODE:="18",
                              EDI_RECEIVED_VALUE:=ITEM_CODE)
                End If
                If rowICTITEM1.Item("ITEM_SNU_CODE") & "" = "S" Then
                    saleable_item = True
                Else
                    saleable_item = False
                End If
            End With
        Else
            Get_ICTITEM1(EDI_UPC, EDI_EAN, EDI_SKU, EDI_ITEM, EDI_STYLE,
                         PRICE, CUST_PRICE, EDI_PRICE, EDI_PRICE_CURR, ITEM_PRICE_CURR, RETAIL_PRICE, PRICE_DISC)
            If ITEM_CODE = "" And ASST_ITEM_CODE = "" Then
                EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="EAN/UPC Code not found",
                                            EDI_COND_CODE:="15",
                                            EDI_RECEIVED_VALUE:=IIf(EDI_UPC <> "", EDI_UPC, EDI_EAN), EDI_REFERENCE:=Mid$(":", 1, System.Math.Sign(Len(EDI_UPC))) & EDI_SKU & Mid$(":", 1, System.Math.Sign(Len(EDI_UPC & EDI_UPC))) & EDI_ITEM)
                If EDI_ACTION = "S" Then
                    ITEM_INACTIVE = True
                    Return
                Else
                    EDI_UPC = EDI_REPLACED_VALUE
                    Get_ICTITEM1(EDI_UPC, EDI_EAN, EDI_SKU, EDI_ITEM, EDI_STYLE,
                         PRICE, CUST_PRICE, EDI_PRICE, EDI_PRICE_CURR, ITEM_PRICE_CURR, RETAIL_PRICE, PRICE_DISC)
                End If
            Else
                If EDI_EAN <> "" And EDI_ITEM <> "" And ITEM_CODE <> "" Then
                    If EDI_ITEM.ToUpper <> ITEM_CODE Then ' ULTA SENT IN A LOWERCASE VERSION OF THE ITEM CODE - LM EMAIL 11/03/2020
                        Dim rowEDTUPCX1 As DataRow = dst.Tables("EDTUPCX1").Rows.Find(New String() {CUST_CODE, EDI_EAN})
                        If rowEDTUPCX1 IsNot Nothing AndAlso rowEDTUPCX1.Item("ITEM_CODE") & "" = ITEM_CODE Then
                            ' ignore the anomaly on mis-matched item codes - the exception table overrides this check
                        Else
                            EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Item Mapped to EAN does not agree with EDI Item",
                          EDI_COND_CODE:="46",
                          EDI_RECEIVED_VALUE:=EDI_ITEM, EDI_EXPECTED_VALUE:=ITEM_CODE)
                            If EDI_ACTION = "S" Then
                                skip_item = True
                            Else
                            End If
                        End If
                    End If
                End If

            End If
        End If

        Dim rowEDTSDQT0 As DataRow = dst.Tables("EDTSDQT0").Rows.Find(New Object() {EDI_DOC_SEQ_NO, EDI_DTL_SEQ, ITEM_CODE, EDI_UPC})
        If rowEDTSDQT0 Is Nothing Then

            PRICE = CUST_PRICE
            If EDI_PRICE < 0.01 Or (PRICE = 0 And EDI_PRICE = 0.01) Then
                EDI_PRICE = PRICE
                EDI_PRICE_CURR = ITEM_PRICE_CURR
            End If


            If (rowEDTSLSP1.Item("EDI_RETAIL_PRICE_FLAG") & "" = "1" And PRICE_BASIS <> "R") Then
                If (EDI_PRICE <> PRICE And System.Math.Abs(EDI_PRICE - RETAIL_PRICE) > 0.01) Then
                    If PRICE = -1 Then
                    Else
                        ' ALREADY NOTED THAT WE DO NOT HAVE A PRICE LIST RECORD
                        EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Item " & ITEM_CODE & "" & " Price s/b " & Format(PRICE, ".00"),
                         EDI_COND_CODE:="17",
                         EDI_RECEIVED_VALUE:=Format(EDI_PRICE, ".00"), EDI_EXPECTED_VALUE:=Format(PRICE, ".00"))

                        If ASCMAIN1.CLIENT = "INT" Then
                            If PRICE <> 0 Then
                                EDI_REPLACED_VALUE = ""
                                EDI_ACTION = ""
                            End If
                        End If

                        If EDI_ACTION = "S" Then
                            skip_item = True
                            Return
                        End If
                        If EDI_REPLACED_VALUE <> "" Then
                            If PRICE <> Val(EDI_REPLACED_VALUE) Then
                                PRICE = Val(EDI_REPLACED_VALUE)
                                RETAIL_PRICE = PRICE / ((100 - PRICE_DISC) / 100)
                                ITEM_PRICE_CURR = PRICE / CURR_EXCH_RATE
                            End If
                        End If
                    End If
                End If

            Else

                If (EDI_PRICE <> PRICE And System.Math.Abs(EDI_PRICE - PRICE) > 0.01) Then ' And PRICE <> 0 Then
                    EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Item " & ITEM_CODE & "" & " Price s/b " & Format(PRICE, ".00"),
                                 EDI_COND_CODE:="17",
                                 EDI_RECEIVED_VALUE:=Format(EDI_PRICE, ".00"), EDI_EXPECTED_VALUE:=Format(PRICE, ".00"))

                    If ASCMAIN1.CLIENT = "INT" Then

                        If EDI_ACTION = "S" Then
                            skip_item = True
                            Return
                        End If

                        If PRICE <> 0 Then
                            EDI_REPLACED_VALUE = ""
                            EDI_ACTION = ""
                        End If
                    End If

                    If EDI_ACTION = "S" Then
                        skip_item = True
                        Return
                    End If
                    If EDI_REPLACED_VALUE <> "" Then
                        If EDI_ACTION = "E" Then
                            If EDI_PRICE <> Val(EDI_REPLACED_VALUE) Then
                                EDI_PRICE = Val(EDI_REPLACED_VALUE)
                                ' RETAIL_PRICE = PRICE / ((100 - PRICE_DISC) / 100)
                                'ITEM_PRICE_CURR = PRICE / CURR_EXCH_RATE
                            End If
                        Else
                            If PRICE <> Val(EDI_REPLACED_VALUE) Then
                                PRICE = Val(EDI_REPLACED_VALUE)
                                RETAIL_PRICE = PRICE / ((100 - PRICE_DISC) / 100)
                                ITEM_PRICE_CURR = PRICE / CURR_EXCH_RATE
                            End If
                        End If
                    End If
                End If
            End If

            If saleable_item And PRICE = 0 Then
                EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Item " & ITEM_CODE & " is Saleable with a Price of $0",
                       EDI_COND_CODE:="38",
                       EDI_RECEIVED_VALUE:=Format(PRICE, ".00"), EDI_REFERENCE:=EDI_UPC)

                If ASCMAIN1.CLIENT = "INT" Then
                    If PRICE <> 0 Then
                        EDI_REPLACED_VALUE = ""
                        EDI_ACTION = ""
                    End If
                End If


                If EDI_ACTION = "S" Then
                    skip_item = True
                    Return
                End If
            End If

            ' THIS NEXT IF STMT NEEDS TO BE WORKED ON
            If ASCMAIN1.CLIENT = "AHA" And CURR_CODE = "CAD" Then
                If (rowEDTSLSP1.Item("EDI_RETAIL_PRICE_FLAG") & "" = "1" And PRICE_BASIS <> "R") Then
                    ' PRICE = EDI_PRICE - PRICE IS FROM THE PRICE LIST AND EDI_PRICE IS RETAIL IN CURR_CODE
                    ITEM_PRICE_CURR = EDI_PRICE_CURR
                Else
                    PRICE = EDI_PRICE
                    ITEM_PRICE_CURR = EDI_PRICE_CURR
                End If

            Else
                PRICE = EDI_PRICE
                ITEM_PRICE_CURR = EDI_PRICE_CURR
            End If

            Write_SDQT0(EDI_UPC, PRICE, EDI_PRICE, EDI_PRICE_CURR, ITEM_PRICE_CURR)
        Else
            If Val(rowEDTSDQT0.Item("EDI_PRICE") & "") <> EDI_PRICE Then
                EDI_REPLACED_VALUE = Bad_Data(EDI_COND_DESC:="Same Item (" & EDI_UPC & ") Diff Lines, Diff Prices",
                            EDI_COND_CODE:="10",
                            EDI_RECEIVED_VALUE:=rowEDTSDQT0.Item("EDI_PRICE") & " vs " & EDI_PRICE)
            End If
        End If

    End Sub

    Sub Process_Items_Pre(
        EDI_UPC As String,
        EDI_EAN As String,
        EDU_SKU As String,
        EDI_ITEM As String,
        EDI_STYLE As String,
        EDI_DETL_QTY As Int32,
        EDI_STORE As String,
        EDI_SHIP_DC As String,
        EDI_ADDR_CODE As String,
        use_case_pack As Boolean,
        EDI_PRICE As Decimal,
        EDI_PRICE_CURR As Decimal)

        Dim QTY As Int32 = 0

        Dim EDI_STOREs As List(Of String) = CUST_CODEs(CUST_CODE_multi)

        Dim found_sdq As Boolean = False
        ' Note that we do not even look at EDT850T2.EDI_UPC & EDI_SKU
        For Each rowEDT850T3 As DataRow In dst.Tables("EDT850T3").Select("EDI_DTL_SEQ = " & CStr(EDI_DTL_SEQ))
            With rowEDT850T3

                Dim first_sdqi As Integer = 0

                For sdqi As Int32 = 1 To 10

                    Dim EDI_STORE_XX_column As String = "EDI_STORE_" & Format(sdqi, "00")
                    Dim skip_store_XX As Boolean = False
                    Dim EDI_STORE_XX As String = .Item(EDI_STORE_XX_column) & ""

                    If CUST_CODEs.Count <> 0 Then
                        If CUST_CODE_multi <> CUST_CODE Then
                            If Not EDI_STOREs.Contains(EDI_STORE_XX) Then
                                skip_store_XX = True
                            End If
                        Else
                            ' note that for the main customer, the list of stores is the list of stores to skip
                            If EDI_STOREs.Contains(EDI_STORE_XX) Then
                                skip_store_XX = True
                            End If
                        End If
                    End If

                    ' 02/07/2025 - Invalid Store caused an error message but the code kept executing and SOTORDR5 was not created for a batch of Bloomies Orders.
                    ' There is code where 'Rollback is commented out that caused the error. It was changed to exit the loop, not create orders and let the user know of the error.
                    Dim sql As String = $"EDI_DOC_SEQ_NO = '{rowEDT850T3.Item("EDI_DOC_SEQ_NO") & String.Empty}' AND CUST_STORE_NO = '{Format_Store(EDI_STORE_XX)}' AND EDI_COND_CODE = '05' AND EDI_ACTION = 'S'"
                    If dst.Tables("EDT850TE").Select(sql).Length > 0 Then
                        skip_store_XX = True
                        skip_store = True
                        Bad_Data_Cond_05 = True
                    End If

                    If EDI_STORE_XX <> "" And Not skip_store_XX Then

                        If first_sdqi = 0 Then first_sdqi = sdqi

                        EDI_STORE = Format_Store(EDI_STORE_XX)
                        skip_store = False
                        Bad_Data_Cond_05 = False

                        Dim ST_KEY = "MK" & EDI_STORE
                        If dst.Tables("EDT850TS").Rows.Find(ST_KEY) Is Nothing Then
                            ' Check for Possible Order Duplication
                            If Not skip_store And Not Bad_Data_Cond_05 Then
                                Get_Ship_To(EDI_STORE, EDI_SHIP_DC, "")
                                dst.Tables("EDT850TS").Rows.Add(New String() {ST_KEY, EDI_STORE, EDI_SHIP_DC, "", CUST_STORE_NO, CUST_DC_NO})
                                Check_for_Possible_Order_Duplication(CUST_STORE_NO)
                            End If
                        End If

                        If sdqi = first_sdqi And Not found_sdq Then ' If sdqi = 1 And Not found_sdq Then
                            found_sdq = True
                            Process_Items(EDI_UPC, EDI_EAN, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_PRICE, EDI_PRICE_CURR)
                            If skip_item Then
                                Exit For
                            End If
                            If ITEM_INACTIVE = True Then
                                Exit For
                            End If
                        End If

                        If rowEDTSLSP1.Item("EDI_SDQ_TOT_IND") & "" <> "1" Then
                            QTY = .Item("EDI_QTY_" & Format(sdqi, "00"))
                        Else
                            QTY = .Item("EDI_QTY_" & Format(sdqi, "00")) * EDI_DETL_QTY
                        End If
                        Write_SDQ(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STORE, EDI_SHIP_DC, QTY)
                    End If
                Next
            End With
        Next

        If found_sdq = False Then
            If ASCMAIN1.CLIENT = "INT" Then

                If sdqs_are_present Then
                    If CUST_CODE <> "NEIMAN" Then MsgBox("Please Inform ABS that you see this message - need to check EDI raw" & vbCrLf & "SDQs are present, but not found for " & "MK" & EDI_STORE & "DC" & EDI_SHIP_DC & "GLN" & EDI_ADDR_CODE, MsgBoxStyle.OkOnly, CUST_CODE & " EDI Doc Seq " & EDI_DOC_SEQ_NO)
                    Exit Sub
                End If

                If CUST_CODE = "NEIMAN" Then
                    MsgBox("NEIMAN sdq anomaly - please call ABS", MsgBoxStyle.OkOnly, CUST_CODE & " EDI Doc Seq " & EDI_DOC_SEQ_NO)
                    Exit Sub
                End If
                'MsgBox("Please Inform ABS that you see this message - need to check EDI raw", MsgBoxStyle.OkOnly, CUST_CODE & " EDI Doc Seq " & EDI_DOC_SEQ_NO)
                'Exit Sub
            End If


            ' MsgBox("Please Inform ABS that you see this message - need to check EDI raw", MsgBoxStyle.OkOnly, CUST_CODE & " EDI Doc Seq " & EDI_DOC_SEQ_NO)

            Dim ST_KEY = "MK" & EDI_STORE & "DC" & EDI_SHIP_DC & "GLN" & EDI_ADDR_CODE
            If dst.Tables("EDT850TS").Rows.Find(ST_KEY) Is Nothing Then 'If dst.Tables("EDT850TS").Rows.Find(EDI_STORE) Is Nothing Then
                ' THE LINE BELOW WAS ADDED TO CATCH A PROBLEM WITH AHA AAFES EDI_DOC_SEQ_NO 0000007892 - WE MAY ALWAYS WANT THIS CALL, BUT IN THE CASE OF AAFES BOTH CUST_STORE_NO AND CUST_DC_NO WERE EMPTY, SO I PRECEDED WITH THE IF 
                If CUST_STORE_NO = "" And CUST_DC_NO = "" Then Get_Ship_To(EDI_STORE, EDI_SHIP_DC, "")

                dst.Tables("EDT850TS").Rows.Add(New String() {ST_KEY, EDI_STORE, EDI_SHIP_DC, EDI_ADDR_CODE, CUST_STORE_NO, CUST_DC_NO})
                'dst.Tables("EDT850TS").Rows.Add(New String() {EDI_STORE, EDI_SHIP_DC, CUST_STORE_NO, CUST_DC_NO})
            End If
            Process_Items(EDI_UPC, EDI_EAN, EDU_SKU, EDI_ITEM, EDI_STYLE, EDI_PRICE, EDI_PRICE_CURR)
            If skip_item Then
                Return
            End If
            If ITEM_INACTIVE Then
                Return
            End If

            If use_case_pack Then
                Dim CARTON_PACK_QTY As Integer = Val(rowICTITEM1.Item("CARTON_PACK_QTY") & "")
                QTY = EDI_DETL_QTY * CARTON_PACK_QTY
            Else
                QTY = EDI_DETL_QTY
            End If

            If ITEM_CODE <> "" Or ASST_ITEM_CODE <> "" Then
                Write_SDQ(EDI_UPC, EDU_SKU, EDI_ITEM, EDI_STORE, EDI_SHIP_DC, QTY)
            End If
        End If

    End Sub

    Sub Write_SDQT0(EDI_UPC As String, PRICE As Decimal, EDI_PRICE As Decimal, EDI_PRICE_CURR As Decimal, ITEM_PRICE_CURR As Decimal)
        If ITEM_CODE <> "" Then
            Dim rowEDTSDQT0 As DataRow = dst.Tables("EDTSDQT0").NewRow
            With rowEDTSDQT0
                .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                .Item("EDI_UPC") = EDI_UPC
                .Item("PRICE") = PRICE
                .Item("EDI_PRICE") = EDI_PRICE
                .Item("EDI_PRICE_CURR") = EDI_PRICE_CURR
                .Item("ITEM_PRICE_CURR") = ITEM_PRICE_CURR
            End With
            dst.Tables("EDTSDQT0").Rows.Add(rowEDTSDQT0)
        End If
    End Sub

    Sub Get_ICTITEM1(
        EDI_UPC As String,
        EDI_EAN As String,
        EDI_SKU As String,
        EDI_ITEM As String,
        EDI_STYLE As String,
        ByRef PRICE As Decimal,
        ByRef CUST_PRICE As Decimal,
        ByRef EDI_PRICE As Decimal,
        ByRef EDI_PRICE_CURR As Decimal,
        ByRef ITEM_PRICE_CURR As Decimal,
        ByRef RETAIL_PRICE As Decimal,
        ByRef PRICE_DISC As Decimal)

        ITEM_CODE = ""

        Dim ALL_EDI_ITEM_ok As Boolean = False

        Dim EDI_UPC_ok As Boolean = (EDI_UPC = "")
        Dim EDI_EAN_ok As Boolean = (EDI_EAN = "")
        Dim EDI_SKU_ok As Boolean = (EDI_SKU = "")
        Dim EDI_ITEM_ok As Boolean = (EDI_ITEM = "")

        Dim EDI_UPC_ITEM As String = IIf(EDI_UPC_ok, "none", "")
        Dim EDI_EAN_ITEM As String = IIf(EDI_EAN_ok, "none", "")
        Dim EDI_SKU_ITEM As String = IIf(EDI_SKU_ok, "none", "")
        Dim EDI_ITM_ITEM As String = IIf(EDI_ITEM_ok, "none", "")

        CUST_PRICE = 0

        ' Locate Item via UPC, if supplied

        ' DO NOT USE THE EDI_UPC AS IF IT WERE A CUSTOMER ITEM CODE
        'If Not EDI_UPC_ok Then
        '    Dim rowSOTCITM1 As DataRow = dst.Tables("SOTCITM1").Rows.Find(New String() {CUST_CODE, EDI_UPC})
        '    If rowSOTCITM1 IsNot Nothing Then
        '        If ITEM_CODE = "" Or rowSOTCITM1.Item("FORCE_SUB") & "" = "1" Then
        '            ITEM_CODE = rowSOTCITM1.Item("ITEM_CODE") & ""
        '        End If
        '        EDI_UPC_ITEM = rowSOTCITM1.Item("ITEM_CODE") & ""
        '        'CUST_PRICE = rowSOTCITM1.Item("CUST_PRICE") & ""
        '        EDI_UPC_ok = True
        '        ALL_EDI_ITEM_ok = True
        '    End If
        'End If

        If Not EDI_UPC_ok Then
            Dim rowEDTUPCX1 As DataRow = dst.Tables("EDTUPCX1").Rows.Find(New String() {CUST_CODE, EDI_UPC})
            If rowEDTUPCX1 IsNot Nothing Then
                Dim row As DataRow = LookUp("ICTITEM1", rowEDTUPCX1.Item("ITEM_CODE") & "")
                If row IsNot Nothing Then
                    ITEM_CODE = rowEDTUPCX1.Item("ITEM_CODE")
                    EDI_UPC_ITEM = ITEM_CODE
                    EDI_UPC_ok = True
                    ALL_EDI_ITEM_ok = True
                End If
            End If
        End If

        If Not EDI_UPC_ok Then
            Dim rowICTITEMX As DataRow = dst.Tables("ICTITEMX").Rows.Find(EDI_UPC)
            If rowICTITEMX Is Nothing And Len(EDI_UPC) = 11 Then
                Dim rowICTITEMX11() As DataRow = dst.Tables("ICTITEMX").Select("ITEM_UPC_CODE LIKE '" & EDI_UPC & "*'")
                If rowICTITEMX11.Length = 1 Then
                    rowICTITEMX = rowICTITEMX11(0)
                End If
            End If

            ' LBM 08/12/2019 - IF AMAZON, AND IF 12 CHAR UPC (IE, ONE THAT DOES NOT HAVE A FORCED LEADING 0), THEN REJECT
            If rowICTITEMX IsNot Nothing Then
                If ASCMAIN1.CLIENT = "INT" Then
                    If CUST_CODE = "AMAZON" Then
                        If EDI_STYLE.Length = 12 Then
                            rowICTITEMX = Nothing
                        End If
                    End If
                End If
            End If


            If rowICTITEMX IsNot Nothing Then
                ITEM_CODE = rowICTITEMX.Item("ITEM_CODE")
                EDI_UPC_ITEM = ITEM_CODE
                EDI_UPC_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If


        If Not EDI_EAN_ok Then
            Dim rowICTITEMX As DataRow = dst.Tables("ICTITEMX").Rows.Find(EDI_EAN)
            If rowICTITEMX Is Nothing And Len(EDI_EAN) = 12 Then
                Dim rowICTITEMX12() As DataRow = dst.Tables("ICTITEMX").Select("ITEM_UPC_CODE LIKE '" & EDI_EAN & "*'")
                If rowICTITEMX12.Length = 1 Then
                    rowICTITEMX = rowICTITEMX12(0)
                End If
            End If
            If rowICTITEMX IsNot Nothing Then
                ITEM_CODE = rowICTITEMX.Item("ITEM_CODE")
                EDI_EAN_ITEM = ITEM_CODE
                EDI_EAN_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        ' Locate Item via Customer Item Code, if supplied

        If Not EDI_SKU_ok Then
            Dim rowSOTCITM1 As DataRow = dst.Tables("SOTCITM1").Rows.Find(New String() {CUST_CODE, EDI_SKU})
            If rowSOTCITM1 IsNot Nothing Then
                If ITEM_CODE = "" Or rowSOTCITM1.Item("FORCE_SUB") & "" = "1" Then
                    ITEM_CODE = rowSOTCITM1.Item("ITEM_CODE") & ""
                End If
                EDI_SKU_ITEM = rowSOTCITM1.Item("ITEM_CODE") & ""
                'CUST_PRICE = rowSOTCITM1.Item("CUST_PRICE") & ""
                EDI_SKU_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        If Not EDI_ITEM_ok Then
            Dim rowSOTCITM1 As DataRow = dst.Tables("SOTCITM1").Rows.Find(New String() {CUST_CODE, EDI_ITEM})
            If rowSOTCITM1 IsNot Nothing Then
                If ITEM_CODE = "" Or rowSOTCITM1.Item("FORCE_SUB") & "" = "1" Then
                    ITEM_CODE = rowSOTCITM1.Item("ITEM_CODE") & ""
                End If
                EDI_ITM_ITEM = rowSOTCITM1.Item("ITEM_CODE") & ""
                'CUST_PRICE = rowSOTCITM1.Item("CUST_PRICE") & ""
                EDI_ITEM_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If


        ' Locate Item using ITEM_CODE
        If ITEM_CODE = "" Then
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(New String() {EDI_ITEM})
            If rowICTITEM1 Is Nothing Then
                rowICTITEM1 = LookUp("ICTITEM1", EDI_ITEM)
            End If
            If rowICTITEM1 IsNot Nothing Then
                ITEM_CODE = EDI_ITEM
                EDI_ITM_ITEM = EDI_ITEM
                EDI_ITEM_ok = True
                ALL_EDI_ITEM_ok = True
            End If
        End If

        ' Stop ' ASSUMES THAT GLOBAL ITEM_CODE IS IN USE
        Check_Item(EDI_UPC_ok, EDI_UPC, EDI_UPC_ITEM, ALL_EDI_ITEM_ok)
        Check_Item(EDI_SKU_ok, EDI_SKU, EDI_SKU_ITEM, ALL_EDI_ITEM_ok)
        Check_Item(EDI_ITEM_ok, EDI_ITEM, EDI_ITM_ITEM, ALL_EDI_ITEM_ok)

        If Not EDI_UPC_ok And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) And ITEM_CODE <> "" Then
            Bad_Data(EDI_COND_DESC:="EDI UPC " & EDI_UPC & " missing for Item " & ITEM_CODE, EDI_COND_CODE:="23", EDI_RECEIVED_VALUE:=EDI_UPC)
        End If
        If Not EDI_SKU_ok And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) And ITEM_CODE <> "" Then
            Bad_Data(EDI_COND_DESC:="EDI SKU " & EDI_SKU & " missing for Item " & ITEM_CODE, EDI_COND_CODE:="24", EDI_RECEIVED_VALUE:=EDI_SKU)
        End If
        If Not EDI_ITEM_ok And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI Item " & EDI_ITEM & " missing for Item " & ITEM_CODE, EDI_COND_CODE:="25", EDI_RECEIVED_VALUE:=EDI_ITEM)
        End If

        If EDI_UPC_ITEM <> "none" And EDI_UPC_ITEM <> ITEM_CODE And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI UPC " & EDI_UPC & " does not match Item " & ITEM_CODE, EDI_COND_CODE:="30", EDI_RECEIVED_VALUE:=EDI_UPC_ITEM)
        End If
        If EDI_SKU_ITEM <> "none" And EDI_SKU_ITEM <> ITEM_CODE And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI SKU " & EDI_SKU & " does not match Item " & ITEM_CODE, EDI_COND_CODE:="31", EDI_RECEIVED_VALUE:=EDI_SKU_ITEM)
        End If
        If EDI_ITM_ITEM <> "none" And EDI_ITM_ITEM <> ITEM_CODE And (Not ITEM_OK_IF_ONE_MATCH Or Not ALL_EDI_ITEM_ok) Then
            Bad_Data(EDI_COND_DESC:="EDI Item " & EDI_ITEM & " does not match Item " & ITEM_CODE, EDI_COND_CODE:="32", EDI_RECEIVED_VALUE:=EDI_ITM_ITEM)
        End If

        If ITEM_CODE <> "" Then

            ' If ASCMAIN1.Running_in_VS And ITEM_CODE.StartsWith("9") Then Stop

            rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            If rowICTITEM1 Is Nothing Then
                rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    rowICTITEM1 = Write_Item_to_MDB(rowICTITEM1)
                End If
            End If

            If rowICTITEM1.Item("ITEM_SNU_CODE") & "" = "S" Then
                saleable_item = True
            Else
                saleable_item = False
            End If

            Dim rowSOTSDIV1 As DataRow = dst.Tables("SOTSDIV1").Rows.Find(rowICTITEM1.Item("SALES_DIVISION_CODE") & "")
            If rowSOTSDIV1 Is Nothing Then
                Bad_Data(EDI_COND_DESC:="Missing/Bad Sales Division for Item " & ITEM_CODE, EDI_COND_CODE:="34", EDI_RECEIVED_VALUE:=rowICTITEM1.Item("SALES_DIVISION_CODE") & "")
            Else
                'If rowSOTSDIV1.Item("STATUS") & "" <> "A" Then
                '    Bad_Data(EDI_COND_DESC:="Inactive Sales Division for Item " & ITEM_CODE, EDI_COND_CODE:="19", EDI_RECEIVED_VALUE:=rowICTITEM1.Item("SALES_DIVISION_CODE") & "")
                '    If EDI_ACTION = "S" Then
                '        Stop
                '        ITEM_INACTIVE = True
                '        Return
                '    End If
                'End If
            End If


            'EDI_PRICE_CURR = 0
            'EDI_PRICE = 0
            Dim T2_PO4_QTY As Int32 = 0
            Dim EDI_DETL_QTY As Int32 = 0
            Dim T6_PO4_QTY As Int32 = 0

            If rowEDTSLSP1.Item("EDI_PO1_TOT_IND") & "" <> "1" Then
                If T2_PO4_QTY <> 0 And T2_PO4_QTY <> Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "") Then
                    Bad_Data(EDI_COND_DESC:="PO4 (" & CStr(T2_PO4_QTY) & ") <> Case Qty (" & CStr(Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")) & ") " & ITEM_CODE,
                             EDI_COND_CODE:="55",
                             EDI_RECEIVED_VALUE:=CStr(T2_PO4_QTY))
                    'If Good_Data_Value <> "" Then
                    '    ' PROCEED
                    'End If
                    EDI_DETL_QTY = EDI_DETL_QTY / T2_PO4_QTY * Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                    EDI_PRICE = EDI_PRICE * T2_PO4_QTY / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                    EDI_PRICE_CURR = EDI_PRICE_CURR * T2_PO4_QTY / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                End If

                If T6_PO4_QTY <> 0 And T6_PO4_QTY <> Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "") Then
                    Bad_Data(EDI_COND_DESC:="PO6 (" & CStr(T6_PO4_QTY) & ") <> Case Qty (" & CStr(Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")) & ") " & ITEM_CODE,
                             EDI_COND_CODE:="56",
                             EDI_RECEIVED_VALUE:=CStr(T6_PO4_QTY))
                    'If Good_Data_Value <> "" Then
                    '    ' PROCEED
                    'End If
                    EDI_DETL_QTY = EDI_DETL_QTY / T6_PO4_QTY * Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                    EDI_PRICE = EDI_PRICE * T6_PO4_QTY / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                    EDI_PRICE_CURR = EDI_PRICE_CURR * T6_PO4_QTY / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                End If
            End If

            If T2_PO4_QTY = 0 And Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "") <> 0 And (rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CA" Or rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "CS" Or (rowEDT850T2.Item("EDI_PRICE_UOM") & "" = "" And (rowEDT850T2.Item("EDI_PO4_UOM") & "" = "CA" Or rowEDT850T2.Item("EDI_PO4_UOM") & "" = "CS"))) Then
                ' FOR SOME REASON, DRUGSTORE ORDERS ARE FALLING INTO THIS ROUTINE, AFTER WE HAVE ALREADY TAKEN THE PRICE AND DIVIDED IT BY CASE QTY
                ' MAYBE i NEED TO SET A FLAG, BUT BUT NOT SURE WHAT FLAG NEEDS TO BE SET
                ' MAYBE T2_PO4_QTY NEEDS TO BE PASSED IN HERE, BECAUSE WE ALREADY USED IT TO DIVIDE PRICE
                ' SEE DRUGSTORE ORDER 6733416 EDI DOC SEQ NO = 0000028015
                If CUST_CODE <> "DRUGSTORE" And CUST_CODE <> "CVS" And CUST_CODE <> "MEIJER" Then
                    ' COMPARE WHATEVER EDI'S REQUIRE THIS ROUTINE TO MEIJER PO 206836492
                    TAC.TACMAIN1.Record_Event("EDT850T2", EDI_DOC_SEQ_NO, Now, ASCMAIN1.USER_ID, "CASE", "Case Pricing Routine", "EDF850I1")
                    EDI_PRICE_CURR = EDI_PRICE / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                    EDI_PRICE = EDI_PRICE * CURR_EXCH_RATE / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                    EDI_DETL_QTY = Val(rowEDT850T2.Item("EDI_TOTAL_QTY") & "") * Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                End If
            End If

            If T2_PO4_QTY = 0 Then
                T2_PO4_QTY = 1
                '            T2_PO4_QTY = dynICTITEM1.Item("ITEM_SO_QTY_MULT")
            End If

            If rowEDTSLSP1.Item("EDI_RETAIL_PRICE_FLAG") & "" = "1" And PRICE_BASIS = "R" Then
                EDI_PRICE_CURR = EDI_PRICE - (EDI_PRICE * (PRICE_BASE_DPCT / 100))
                EDI_PRICE = EDI_PRICE * CURR_EXCH_RATE - (EDI_PRICE * CURR_EXCH_RATE * (PRICE_BASE_DPCT / 100))
                'Else
                '    EDI_PRICE_CURR = EDI_PRICE
                '    EDI_PRICE = EDI_PRICE * CURR_EXCH_RATE
            End If

            PRICE = 0
            RETAIL_PRICE = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
            Dim ITEM_RETAIL_PRICE_NEW As Decimal = Val(rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE") & "")
            Dim ITEM_NEW_RETAIL_PRICE_DATE As String = rowICTITEM1.Item("ITEM_NEW_RETAIL_PRICE_DATE") & ""

            If CURR_CODE <> "USD" Then
                Dim rowSOTPRIC2 As DataRow = LookUp("SOTPRIC2", New String() {CURR_CODE, ITEM_CODE})
                If rowSOTPRIC2 Is Nothing Then
                    RETAIL_PRICE = 0
                    ITEM_RETAIL_PRICE_NEW = 0
                    ITEM_NEW_RETAIL_PRICE_DATE = ""
                Else
                    RETAIL_PRICE = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                    ITEM_RETAIL_PRICE_NEW = Val(rowSOTPRIC2.Item("ITEM_NEW_PRICE") & "")
                    ITEM_NEW_RETAIL_PRICE_DATE = rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") & ""
                End If
            End If

            If ITEM_NEW_RETAIL_PRICE_DATE <> "" Then
                If Format(ORDR_SHIP_DATE, "yyyyMMdd") >= Format(CDate(ITEM_NEW_RETAIL_PRICE_DATE), "yyyyMMdd") Then
                    RETAIL_PRICE = ITEM_RETAIL_PRICE_NEW
                End If
            End If

            ITEM_PRICE_CURR = 0
            PRICE_DISC = 0

            Select Case PRICE_BASIS
                Case "R"
                    PRICE = (RETAIL_PRICE - (RETAIL_PRICE * (PRICE_BASE_DPCT / 100))) * CURR_EXCH_RATE
                    ITEM_PRICE_CURR = RETAIL_PRICE - (RETAIL_PRICE * (PRICE_BASE_DPCT / 100))

                    PRICE = System.Math.Round(PRICE, 2, MidpointRounding.AwayFromZero)
                    ITEM_PRICE_CURR = System.Math.Round(ITEM_PRICE_CURR, 2, MidpointRounding.AwayFromZero)

                    PRICE_DISC = PRICE_BASE_DPCT
                    CUST_PRICE = PRICE

                    If PRICE_LIST_CODE_OVERRIDE <> "" Then

                        Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE_OVERRIDE, ITEM_CODE})

                        If rowSOTPRIC2 IsNot Nothing Then
                            PRICE = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                            ITEM_PRICE_CURR = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                            PRICE_DISC = 0
                            CUST_PRICE = PRICE
                        End If

                    Else
                        If PRICE_LIST_CODE <> "" Then
                            Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, ITEM_CODE})

                            If rowSOTPRIC2 IsNot Nothing Then
                                PRICE = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                                ITEM_PRICE_CURR = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                                PRICE_DISC = 0
                                CUST_PRICE = PRICE
                            End If
                        End If
                    End If


                Case "L"
                    ' SOTPRIC1 & SOTPRIC2 HAVE DIFFERENT LAYOUTS THAN THIS CODE SUGGESTS.
                    'Stop
                    Dim rowSOTPRIC2 As DataRow = Nothing

                    If PRICE_LIST_CODE_OVERRIDE <> "" Then
                        rowSOTPRIC2 = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE_OVERRIDE, ITEM_CODE})
                    Else
                        rowSOTPRIC2 = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, ITEM_CODE})
                    End If

                    If rowSOTPRIC2 IsNot Nothing Then

                        If rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE") & "" <> "" _
                            AndAlso Format(rowSOTPRIC2.Item("ITEM_NEW_PRICE_DATE"), "yyyyMMdd") <= Format(ORDR_SHIP_DATE, "yyyyMMdd") Then
                            PRICE = Val(rowSOTPRIC2.Item("ITEM_NEW_PRICE") & "")
                        Else
                            PRICE = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                        End If

                        ' PRICE = rowSOTPRIC2.Item("ITEM_PRICE")

                        CUST_PRICE = PRICE
                        ITEM_PRICE_CURR = PRICE '  rowSOTPRIC2.Item("ITEM_PRICE")
                        PRICE_DISC = 0
                    Else
                        'Dim rowSOTPRIC1 As DataRow = dst.Tables("SOTPRIC1").Rows.Find(New String() {ITEM_CODE})
                        'If rowSOTPRIC1 IsNot Nothing Then
                        '    PRICE = rowSOTPRIC1.Item("ITEM_PRICE")
                        '    CUST_PRICE = PRICE
                        '    ITEM_PRICE_CURR = rowSOTPRIC1.Item("ITEM_PRICE")
                        '    PRICE_DISC = 0
                        'Else
                        PRICE = -1
                        CUST_PRICE = -1
                        Bad_Data(EDI_COND_DESC:="No Price List record for Item " & ITEM_CODE,
                                 EDI_COND_CODE:="39",
                                 EDI_RECEIVED_VALUE:=ITEM_CODE)
                        'If Good_Data_Value <> "" Then
                        '    'ORDR_CUST_PO = Good_Data_Value
                        'End If
                        'End If
                    End If
                    'price = lookup from price file

                Case "E"
                    PRICE = rowEDT850T2.Item("EDI_PRICE") * CURR_EXCH_RATE / T2_PO4_QTY
                    CUST_PRICE = PRICE
                    ITEM_PRICE_CURR = rowEDT850T2.Item("EDI_PRICE") / T2_PO4_QTY
                    PRICE_DISC = 0

            End Select

            If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                If CUST_PRICE <> EDI_PRICE Then
                    Dim PRICE_TO_USE As String = Bad_Data(EDI_COND_DESC:="Price Transmitted (" & Format(EDI_PRICE, "#.00") & ") <> Net Price (" & Format(CUST_PRICE, "#.00") & ") " & ITEM_CODE,
                             EDI_COND_CODE:="57",
                             EDI_RECEIVED_VALUE:=CStr(EDI_PRICE),
                             EDI_EXPECTED_VALUE:=CStr(CUST_PRICE))

                    ' not here
                    If EDI_ACTION = "S" Then ' LBM wants to skip items that at NC - she says she can skip saleable (I cannot replicate that)
                        skip_item = True
                    Else
                    End If

                    If PRICE_TO_USE <> "" And PRICE_TO_USE <> "Skip" Then
                        If ASCMAIN1.CLIENT = "INT" Then
                            If CUST_PRICE <> 0 Then
                                EDI_REPLACED_VALUE = ""
                                EDI_ACTION = ""
                            End If
                        End If
                        'If ASCMAIN1.CLIENT = "INT" Then
                        '    If Val(PRICE_TO_USE) <> CUST_PRICE Then
                        '        PRICE_TO_USE = CStr(CUST_PRICE)
                        '    End If
                        'End If

                        CUST_PRICE = Val(PRICE_TO_USE)
                        EDI_PRICE = Val(PRICE_TO_USE)
                    End If
                    'EDI_PRICE = EDI_PRICE * T6_PO4_QTY / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                    'EDI_PRICE_CURR = EDI_PRICE_CURR * T6_PO4_QTY / Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "")
                End If
            End If

            Dim rowEDTITEMX As DataRow = dst.Tables("EDTITEMX").NewRow
            With rowEDTITEMX
                .Item("CUST_CODE") = CUST_CODE
                .Item("EDI_UPC") = EDI_UPC
                .Item("EDI_SKU") = EDI_SKU
                .Item("EDI_ITEM") = EDI_ITEM
                .Item("EDI_EAN") = EDI_EAN
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("CUST_PRICE") = CUST_PRICE
                .Item("ITEM_PRICE_CURR") = ITEM_PRICE_CURR
                .Item("PRICE_DISC") = PRICE_DISC
                .Item("RETAIL_PRICE") = RETAIL_PRICE
                .Item("RETAIL_PRICE_CURR") = RETAIL_PRICE / CURR_EXCH_RATE
                '.Item("SALES_DIVISION_CODE") = SALES_DIVISION_CODE
            End With
            dst.Tables("EDTITEMX").Rows.Add(rowEDTITEMX)
        End If

    End Sub

    Sub Check_Item(ByRef EDI_XXX_ok, ByRef EDI_XXX, ByRef EDI_XXX_ITEM, ByRef ALL_EDI_ITEM_ok)

        Dim rowICTITEM1 As DataRow
        If EDI_XXX_ok = False Then
            rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(EDI_XXX)
            If rowICTITEM1 IsNot Nothing Then
                If ITEM_CODE = "" Then
                    ITEM_CODE = rowICTITEM1.Item("ITEM_CODE") & ""
                End If
                EDI_XXX_ITEM = rowICTITEM1.Item("ITEM_CODE") & ""
                EDI_XXX_ok = True
                ALL_EDI_ITEM_ok = True
            Else
                rowICTITEM1 = LookUp("ICTITEM1", EDI_XXX)
                If rowICTITEM1 IsNot Nothing Then
                    rowICTITEM1 = Write_Item_to_MDB(rowICTITEM1)
                    EDI_XXX_ITEM = ITEM_CODE
                    EDI_XXX_ok = True
                    ALL_EDI_ITEM_ok = True
                End If
            End If
        End If
    End Sub

    Sub PO4_OverWrite(
        EDI_UPC As String,
        EDI_SKU As String,
        EDI_ITEM As String,
        PO4_OW As Int32,
        T2_PO4_QTY As Int32,
        T6_PO4_QTY As Int32)

        Dim rowSOTCITM1 As DataRow = dst.Tables("SOTCITM1").Rows.Find(New String() {CUST_CODE, EDI_UPC})
        If rowSOTCITM1 IsNot Nothing Then
            rowSOTCITM1 = dst.Tables("SOTCITM1").Rows.Find(New String() {CUST_CODE, EDI_ITEM})
        ElseIf rowSOTCITM1 IsNot Nothing Then
            rowSOTCITM1 = dst.Tables("SOTCITM1").Rows.Find(New String() {CUST_CODE, EDI_SKU})
        End If

        If rowSOTCITM1 IsNot Nothing Then
            If PO4_OW = Val(rowSOTCITM1.Item("CUST_PO4_QTY") & "") + 0 Then
                T2_PO4_QTY = Val(rowSOTCITM1.Item("ITEM_PO4_QTY") & "")
                T6_PO4_QTY = Val(rowSOTCITM1.Item("ITEM_PO4_QTY") & "")
            End If
        End If
    End Sub

    Function Write_Item_to_MDB(ROW As DataRow) As DataRow
        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").NewRow
        rowICTITEM1.ItemArray = ROW.ItemArray

        If Val(rowICTITEM1.Item("ITEM_SO_QTY_MULT") & "") = 0 Then
            If Val(rowICTITEM1.Item("CARTON_PACK_QTY") & "") <> 0 Then
                rowICTITEM1.Item("ITEM_SO_QTY_MULT") = Val(rowICTITEM1.Item("CARTON_PACK_QTY") & "")
            Else
                rowICTITEM1.Item("ITEM_SO_QTY_MULT") = 1
            End If
        End If
        dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)

        Return rowICTITEM1
    End Function

    Sub Write_SDQ(
        EDI_UPC As String,
        EDI_SKU As String,
        EDI_ITEM As String,
        EDI_STORE As String,
        EDI_SHIP_DC As String, QTY As Int32)

        If ITEM_CODE = "" Then
            Return
        End If

        If skip_item Then ' ADDED BY WJZ TO ADDRESS VONMAUR PO 
            If Not skipped_items.Contains(ITEM_CODE) Then skipped_items.Add(ITEM_CODE)
            Return
        End If

        If skip_store Or Bad_Data_Cond_05 Then
            Return
        End If

        Dim rowEDTSDQT1 As DataRow = dst.Tables("EDTSDQT1").Rows.Find(New Object() {EDI_DOC_SEQ_NO, EDI_STORE, EDI_DTL_SEQ, ITEM_CODE})
        If rowEDTSDQT1 IsNot Nothing Then
            rowEDTSDQT1.Item("QTY") = Val(rowEDTSDQT1.Item("QTY") & "") + QTY
        Else
            rowEDTSDQT1 = dst.Tables("EDTSDQT1").NewRow

            '     If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "80862" Then Stop

            With rowEDTSDQT1
                .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                .Item("EDI_STORE") = EDI_STORE
                .Item("ITEM_CODE") = ITEM_CODE
                .Item("EDI_ITEM") = EDI_ITEM
                .Item("ASST_ITEM_CODE") = ASST_ITEM_CODE
                .Item("QTY") = QTY
                .Item("EDI_DTL_SEQ") = EDI_DTL_SEQ
                .Item("EDI_SHIP_DC") = EDI_SHIP_DC
                .Item("EDI_UPC") = EDI_UPC
                .Item("EDI_SKU") = EDI_SKU

                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = CUST_STORE_NO
                .Item("CUST_DC_NO") = CUST_DC_NO
            End With

            dst.Tables("EDTSDQT1").Rows.Add(rowEDTSDQT1)
        End If
    End Sub

    Sub Update_ICTSTAT2(ORDR_GROUP_NO As String, S As Integer)
        ASCMAIN1.sql = "" _
            & "Begin" _
            & " Declare Cursor C1 is Select SOTORDR2.WHSE_CODE, SOTORDR2.ITEM_CODE, Sum (SOTORDR2.ORDR_QTY) ORDR_QTY" _
            & "  from SOTORDR1,SOTORDR2" _
            & "  where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" _
            & "    and SOTORDR1.ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'" _
            & "  group by SOTORDR2.WHSE_CODE, SOTORDR2.ITEM_CODE;" _
            & " Begin" _
            & "  For R1 in C1 Loop" _
            & "   ICPSTAT2(R1.ITEM_CODE,R1.WHSE_CODE,0,0,0,R1.ORDR_QTY * " & S & ",0,0);" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()
    End Sub

    Private Sub grdEDT850TE_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdEDT850TE.AfterRowActivate

    End Sub

    Function Format_Store(EDI_STORE As String)
        Dim CUST_STORE_NO As String = EDI_STORE
        If IsNumeric(EDI_STORE) Then
            CUST_STORE_NO = Format(Val(EDI_STORE), "000000")
        End If
        Return CUST_STORE_NO
    End Function

    Private Sub tabMain_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()

        ' reparenting a grid with summaries is causing a wierd internal infragistics error

        If tabMain.SelectedTab.Key = "Orders Waiting to be Imported" Then

            grdEDT850T1.DisplayLayout.Bands(0).Summaries.Clear()

            grdEDT850T1.DataSource = dst.Tables("EDT850T1")
            grdEDT850T1.Parent = splEDT850T1.Panel1
            With grdEDT850T1.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = False
                .Columns("ORDERS").Hidden = True
                .Columns("AMOUNT").Hidden = True
            End With

            If grdEDT850T1.DisplayLayout.Bands(0).Summaries.Count = 0 Then
                Create_Summary(grdEDT850T1, "CUST_CODE", "Count")
                Create_Summary(grdEDT850T1, New String() {"SEL", "ORDERS", "AMOUNT"})
            End If

        ElseIf tabMain.SelectedTab.Key = "Imported Orders" Then

            grdEDT850T1.DisplayLayout.Bands(0).Summaries.Clear()

            grdEDT850T1.DataSource = dst.Tables("EDT850T1_IMPORTED")
            grdEDT850T1.Parent = tabMain.Tabs("Imported Orders").TabPage
            With grdEDT850T1.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = True
                .Columns("ORDERS").Hidden = False
                .Columns("AMOUNT").Hidden = False
            End With

            If grdEDT850T1.DisplayLayout.Bands(0).Summaries.Count = 0 Then
                Create_Summary(grdEDT850T1, "CUST_CODE", "Count")
                Create_Summary(grdEDT850T1, New String() {"SEL", "ORDERS", "AMOUNT"})
            End If

        ElseIf tabMain.SelectedTab.Key = "Archived EDI" Then

            grdEDT850T1.DisplayLayout.Bands(0).Summaries.Clear()

            dst.Tables("EDT850T1_ARCHIVED").Rows.Clear()
            grdEDT850T1.DataSource = dst.Tables("EDT850T1_ARCHIVED")
            grdEDT850T1.Parent = tabMain.Tabs("Archived EDI").TabPage
            With grdEDT850T1.DisplayLayout.Bands(0)
                .Columns("SEL").Hidden = False
                .Columns("ORDERS").Hidden = True
                .Columns("AMOUNT").Hidden = True
            End With

        End If

        UltraExplorerBar1.Groups("Screen Control").Visible = (tabMain.SelectedTab.Key = "Orders Waiting to be Imported")
        UltraExplorerBar1.Groups("Action").Visible = (tabMain.SelectedTab.Key = "Orders Waiting to be Imported")
        UltraExplorerBar1.Groups("EDIs Already Imported").Visible = (tabMain.SelectedTab.Key = "Archived EDI")
    End Sub

    Private Sub grdEDT850TE_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT850TE.InitializeRow
        'If e.Row.Cells("EDI_ACTION").Value & "" = "A" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Orange
        'ElseIf e.Row.Cells("EDI_ACTION").Value & "" = "E" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Blue
        'ElseIf e.Row.Cells("EDI_ACTION").Value & "" = "R" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
        'ElseIf e.Row.Cells("EDI_ACTION").Value & "" = "S" Then
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Pink
        'Else
        '    e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Empty
        'End If

        Dim EDI_ACTION As String = e.Row.Cells("EDI_ACTION").Value & ""
        Dim EDI_COND_CODE As String = e.Row.Cells("EDI_COND_CODE").Value & ""

        e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Empty

        If EDI_ACTION = "A" Then
            If Not ACTIONS_A.Contains("*") And Not ACTIONS_A.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
        ElseIf EDI_ACTION = "E" Then
            If Not ACTIONS_E.Contains("*") And Not ACTIONS_E.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
            If ASCMAIN1.DBS_COMPANY = "SLP" Then
            Else
                If ASCMAIN1.CLIENT = "INT" Then
                    If (EDI_COND_CODE = "17" Or EDI_COND_CODE = "57") Then
                        Dim EDI_EXPECTED_VALUE As Decimal = Val(e.Row.Cells("EDI_EXPECTED_VALUE").Value & "")
                        If EDI_EXPECTED_VALUE = 0 Then
                            ' OK
                        Else
                            e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
                        End If
                    End If
                End If
            End If

        ElseIf EDI_ACTION = "S" Then
            If Not ACTIONS_S.Contains("*") And Not ACTIONS_S.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
        ElseIf EDI_ACTION = "R" Then
            If Not ACTIONS_R.Contains("*") And Not ACTIONS_R.Contains(EDI_COND_CODE) Then
                e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red
            End If
        End If

        If e.Row.Cells("EDI_ACTION").Appearance.ForeColor = Drawing.Color.Red Then
            e.Row.Cells("EDI_ACTION").ToolTipText = "Action is not permitted"
        Else
            e.Row.Cells("EDI_ACTION").ToolTipText = ""
        End If

        If e.Row.Cells("RESOLUTION").Value & "" <> "" Then
            e.Row.Cells("EDI_COND_DESC").ToolTipText = e.Row.Cells("RESOLUTION").Value
        End If
    End Sub

    Private Sub grdEDT850T1_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdEDT850T1.ClickCellButton
        Dim sql_where As String = ""

        Select Case e.Cell.Column.Key

            Case "CUST_CODE_OVERRIDE"
                grdClickCellButton(grdEDT850T1, sql_where, , "CUST_CODE_OVERRIDE", "CUST_CODE")

            Case "PRICE_LIST_CODE_OVERRIDE"
                grdClickCellButton(grdEDT850T1, sql_where, , "PRICE_LIST_CODE_OVERRIDE", "PRICE_LIST_CODE")
        End Select
    End Sub

    Private Sub grdEDT850T1_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDT850T1.InitializeLayout

    End Sub

    Private Sub grdEDT850T1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT850T1.InitializeRow
        If Not e.Row.IsDataRow Then
            Exit Sub
        Else

            Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & ""
            Dim EDI_PO_PURP As String = e.Row.Cells("EDI_PO_PURP").Value & ""
            Dim EDI_TP_ID As String = e.Row.Cells("EDI_TP_ID").Value & ""
            Dim EDI_PROCESS_IND As String = e.Row.Cells("EDI_PROCESS_IND").Value & ""

            If e.Row.Cells("RESULT").Value & "" = "Rejected" Then
                e.Row.Cells("RESULT").Appearance.ForeColor = Drawing.Color.Red
            ElseIf e.Row.Cells("RESULT").Value & "" = "Imported" Then
                e.Row.Cells("RESULT").Appearance.ForeColor = Drawing.Color.Blue
            Else
                e.Row.Cells("RESULT").Appearance.ForeColor = Drawing.Color.Empty
            End If
            If e.Row.Cells("EDI_PO_TYPE").Value & "" = "TR" Then
                e.Row.Cells("EDI_PO_TYPE").Appearance.ForeColor = Drawing.Color.Red
            End If


            If EDI_PO_PURP = "07" Then
                If EDI_TP_ID = "007942915" Then
                    e.Row.Cells("EDI_PO_PURP").Appearance.ForeColor = Drawing.Color.Red
                    e.Row.Cells("EDI_PO_PURP").ToolTipText = "This is a Full Replacement Order - delete the originally imported order before processing this order"
                End If
            ElseIf EDI_PO_PURP = "01" Then
                e.Row.Cells("EDI_PO_PURP").Appearance.BackColor = Drawing.Color.Red
                e.Row.Cells("EDI_PO_PURP").Appearance.ForeColor = Drawing.Color.White
                e.Row.Cells("EDI_PO_PURP").ToolTipText = "This is a Cancellation Advise - not a new EDI PO"

            End If

            With e.Row.Cells("EDI_PROCESS_IND").Appearance
                If EDI_PROCESS_IND = "1" Then
                    .ForeColor = Drawing.Color.Blue
                ElseIf EDI_PROCESS_IND = "C" Then
                    .ForeColor = Drawing.Color.Red
                ElseIf EDI_PROCESS_IND <> "0" Then
                    .ForeColor = Drawing.Color.DarkOrange
                Else
                    .ForeColor = Drawing.Color.Empty
                End If
            End With

        End If
    End Sub

    Sub Setup_RESOLUTIONS()

        ' add code to RESOLUTIONS to show the description that should appear in the Rejections Screen
        ' ie, the resolution requires some action in some other program, like Customer Master
        RESOLUTIONS.Add("01", "Trading Partner Setup Required")
        RESOLUTIONS.Add("03", "Customer Master Setup Required")
        RESOLUTIONS.Add("04", "Customer Master Setup Required")
        RESOLUTIONS.Add("08", "Customer must re-transmit with a Ship Date")
        RESOLUTIONS.Add("09", "Accept EDI Data, Get an Extension from Customer, and Change Cancel Date in Sales Order Entry")
        RESOLUTIONS.Add("11", "Map EDI Terms to a Valid AR Terms Code")
        RESOLUTIONS.Add("12", "Map EDI Terms to a Valid AR Terms Code")
        RESOLUTIONS.Add("13", "Customer Master Setup Required")
        RESOLUTIONS.Add("14", "Check Customer Order History for this PO")
        RESOLUTIONS.Add("21", "Customer Master Setup Required")
        RESOLUTIONS.Add("27", "Check other unprocessed EDI Orders for Same TP ID and Same Customer PO")
        RESOLUTIONS.Add("28", "Check this PO for Duplicate Items in Order Detail")
        RESOLUTIONS.Add("34", "Item Master Setup Required")
        RESOLUTIONS.Add("35", "Customer Master Setup Required")
        RESOLUTIONS.Add("36", "Customer Master Setup Required")
        RESOLUTIONS.Add("41", "Customer Master Setup Required")
        RESOLUTIONS.Add("48", "DC Transit Business Days Setup Required")
        RESOLUTIONS.Add("56", "hI mOM")
        RESOLUTIONS.Add("91", "Customer Master Setup Required")
        RESOLUTIONS.Add("42", "Verify the ABS Store/DC Address Record")
        RESOLUTIONS.Add("43", "Cancellation Order (Purpose Code = 01)")

        ' add code to ACTIONS_A if it is ok to use the value Received

        ACTIONS_A.Add("04")
        ACTIONS_A.Add("06")
        ACTIONS_A.Add("09")
        ACTIONS_A.Add("13")
        ACTIONS_A.Add("14")
        If ASCMAIN1.CLIENT = "INT" Then
        Else
            ACTIONS_A.Add("17")
        End If

        ACTIONS_A.Add("27")
        ACTIONS_A.Add("28")
        ACTIONS_A.Add("33")
        ACTIONS_A.Add("38")
        ACTIONS_A.Add("39")
        ACTIONS_A.Add("40")
        ACTIONS_A.Add("55")
        ACTIONS_A.Add("56")
        ACTIONS_A.Add("42")

        If ASCMAIN1.CLIENT = "INT" Then
        Else
            ACTIONS_A.Add("57")
        End If


        ' add code to ACTIONS_S if it is ok to Skip the record

        ACTIONS_S.Add("05")
        ACTIONS_S.Add("15")
        ACTIONS_S.Add("17")

        If ASCMAIN1.CLIENT = "INT" Then
            ACTIONS_S.Add("57")
            ACTIONS_S.Add("46") ' 07/20/18 EMAIL LBM Before I skip a record I have communication in writing to remove from the PO from the retailer.
            ACTIONS_S.Add("28") ' 08/12/19 EMAIL LBM Amazon is a problem retailer, and they transmitted the upc and the ean version of the upc.  I need to skip the record to import the order
        Else
        End If

        ACTIONS_S.Add("23")
        ACTIONS_S.Add("24")
        ACTIONS_S.Add("25")
        ACTIONS_S.Add("30")
        ACTIONS_S.Add("31")
        ACTIONS_S.Add("32")
        ACTIONS_S.Add("38")

        ' add code to ACTIONS_R if it is ok to Use Replacement Value

        ACTIONS_R.Add("15")

        ACTIONS_E.Add("*")

    End Sub

    Private Sub tabEDT850T1_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabEDT850T1.SelectedTabChanged
        If tabEDT850T1.SelectedTab.Key = "Raw EDI Data" Then
            If grdEDT850T1.ActiveRow IsNot Nothing AndAlso grdEDT850T1.ActiveRow.IsDataRow Then
                Dim EDI_DOC_SEQ_NO As String = grdEDT850T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Value
                txtRaw.Text = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO)
            End If
        End If
    End Sub

    Private Sub cmdFetchEDIs_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetchEDIs.Click
        ' ISSUE-6973 - cust code no longer required
        'If txtCUST_CODE.Text = "" Then
        '    MsgBox("You Must First Select a Customer and a Date Range", MsgBoxStyle.OkOnly, "Cannot Fetch")
        '    Exit Sub
        'End If

        Load_EDT850T1_ARCHIVED()

    End Sub

    Private Sub cmdReQueue_Click(sender As System.Object, e As System.EventArgs) Handles cmdReQueue.Click
        If dst.Tables("EDT850T1_ARCHIVED").Select("SEL = '1'").Length = 0 Then
            MsgBox("Nothing Selected", MsgBoxStyle.OkOnly, "Cannot Re-Queue")
            Exit Sub
        End If

        Dim I As Integer = 0
        For Each rowEDT850T1_ARCHIVED As DataRow In dst.Tables("EDT850T1_ARCHIVED").Select("SEL = '1'")
            Dim EDI_DOC_SEQ_NO As String = rowEDT850T1_ARCHIVED.Item("EDI_DOC_SEQ_NO")
            ASCDATA1.ExecuteSQL($"Update EDT850T1 Set EDI_PROCESS_IND = '0', LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}' where EDI_DOC_SEQ_NO = '{EDI_DOC_SEQ_NO}'")
            I += 1
            DATETIME_STAMP = Now + ASCMAIN1.NowTSD
            Write_Event_Log("EDT850T1", EDI_DOC_SEQ_NO, "Re-Queued")
        Next

        MsgBox(CStr(I) & " EDI Documents have been Re-Queued", MsgBoxStyle.OkOnly, "Verification")

        tabMain.SelectedTab = tabMain.Tabs("Orders Waiting to be Imported")
        Load_EDT850T1()

    End Sub

    Sub Load_EDT850T1_ARCHIVED()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing Data")

        Fill_Records("EDT850T1_ARCHIVED", New Object() {dteFrom.Value, dteTo.Value, txtCUST_CODE.Text})

        Sort_grdColumns(grdEDT850T1, "EDI_DOC_SEQ_NO".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdEDT850TE_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdEDT850TE.DoubleClickRow
        If e.Row.IsDataRow Then
            If grdEDT850TE.ActiveCell IsNot Nothing AndAlso grdEDT850TE.ActiveCell.Column.Key = "CUST_CODE" Then
                For Each grow As UltraWinGrid.UltraGridRow In grdEDT850T1.Rows
                    If grow.Cells("EDI_DOC_SEQ_NO").Value = e.Row.Cells("EDI_DOC_SEQ_NO").Value Then
                        grdEDT850T1.ActiveRow = grow
                        grow.Selected = True
                        grdEDT850T1.ActiveRowScrollRegion.ScrollRowIntoView(grow)
                        Exit For
                    End If
                Next
                tabMain.SelectedTab = tabMain.Tabs("Orders Waiting to be Imported")
            End If
        End If
    End Sub
    Public Shared Function AddDaysSkipWeekends(d As Date, days As Integer) As Date
        While days > 0
            If Weekday(d) > 1 And Weekday(d) < 7 Then 'Monday to Friday
                days = days - 1
            End If
            d = DateAdd("d", 1, d)
        End While
        Return d
    End Function

    Private Sub grdEDT850T1_BeforeExitEditMode(sender As Object, e As BeforeExitEditModeEventArgs) Handles grdEDT850T1.BeforeExitEditMode
        Try
            If grdEDT850T1.ActiveCell IsNot Nothing Then
                With grdEDT850T1.ActiveCell
                    Select Case .Column.Key
                        Case "PRICE_LIST_CODE_OVERRIDE"
                            .EditorResolved.Value = (.EditorResolved.Value & "").ToString.ToUpper '  ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                        Case "CUST_CODE_OVERRIDE"
                            .EditorResolved.Value = (.EditorResolved.Value & "").ToString.ToUpper '  ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)

                    End Select
                End With
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class