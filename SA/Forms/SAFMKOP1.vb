Imports Infragistics.Win.UltraWinGrid

Public Class SAFMKOP1
    Dim rowSATMKOP1 As DataRow
    Dim TRACKER_NO As String
    Dim sqlSATMKOPX As String
    Dim sqlICTITEM1 As String = ""
    Dim ICTITEM1 As String = ""
    Dim ICTSTATX As String
    Dim sqlSPTCOOPX As String
    '  Dim update_with_approval As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'If MENU_ITEM_OBJECT = "SAFMKOP1" Then
        '    InquiryMode = True
        'End If

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
            '    .Items("Generate Excel").Visible = Not InquiryMode
        End With

        With UltraExplorerBar1.Groups("Excel Options")

            .Items("Generate Excel").Visible = Not InquiryMode
            .Items("Save As New").Visible = Not InquiryMode
            '.Visible = True
            'With .Items.Add("Generate Excel")
            '    .Text = .Key
            'End With
            'With .Items.Add("Save As New")
            '    .Text = .Key
            'End With

        End With

        Get_PARM("ICTPARM1")
        Get_PARM("SPTPARM1")

        ' Create Temp

        ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 where ROWNUM < 1"
        ICTITEM1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1 & " add Primary Key (ITEM_CODE)")

        With dst

            ASCMAIN1.sql = "Select POTORDR2.ITEM_CODE,POTORDR2.PO_ORDER_NO,POTORDR2.PO_QTY_OPN,POTORDR2.PO_DATE_REQUIRED FROM POTORDR1,POTORDR2 WHERE" & vbCrLf _
                & " POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO And POTORDR1.PO_STATUS = 'O' And PO_QTY_OPN <> 0" & vbCrLf _
                & " And ITEM_CODE IN " & vbCrLf _
                & "(Select DISTINCT ITEM_CODE FROM " & ICTITEM1 & ")"
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, False, "")


            ASCMAIN1.sql = "Select ICTPINV2.ITEM_CODE,ICTPINV1.WHSE_CODE, ICTPINV1.INV_NUM, ICTPINV1.INV_DATE, ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO, ICTPINV1.VESSEL_NAME, ICTPINV1.BOL_NO," & vbCrLf _
                & " ICTPINV1.CONTAINER_NO, ICTPINV1.SHIP_DATE, ICTPINV1.ETA_DATE, ICTPINV2.PINV_QTY FROM ICTPINV1, ICTPINV2 WHERE" & vbCrLf _
                & " ICTPINV1.PINV_NO = ICTPINV2.PINV_NO " & vbCrLf _
                & " And ICTPINV1.PINV_STATUS = 'O' " & vbCrLf _
                & " And ITEM_CODE IN " & vbCrLf _
                & "(Select DISTINCT ITEM_CODE FROM " & ICTITEM1 & ")"
            Create_TDA(.Tables.Add, "ICTPINV2", "**", 0, False, "")

            ASCMAIN1.sql = "SELECT  ICTSTAT2.ITEM_CODE,Sum(WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND,sum(WHSE_QTY_ONPO) WHSE_QTY_ONPO,sum(WHSE_QTY_PLAN) WHSE_QTY_PLAN,sum(WHSE_QTY_OPEN) WHSE_QTY_OPEN, sum(WHSE_QTY_PICK) WHSE_QTY_PICK, sum(WHSE_QTY_COMM) WHSE_QTY_COMM, sum(WHSE_QTY_HOLD) WHSE_QTY_HOLD" & vbCrLf _
                & " From ICTSTAT2 Where" & vbCrLf _
                & " ICTSTAT2.WHSE_CODE In (Select WHSE_CODE From ICTWHSE1 Where LP_CODE Is Not Null And NVL(WHSE_MRP_EXC_IND,'0') = '0') And " & vbCrLf _
                & " ICTSTAT2.ITEM_CODE IN " & vbCrLf _
                & "(Select DISTINCT ITEM_CODE FROM " & ICTITEM1 & ") GROUP BY ICTSTAT2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTSTATX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTALLO1.ITEM_CODE,SOTALLO1.ALLO_CTL_NO,SOTALLO1.QTY_ALLO_PLAN,SOTALLO1.DATE_START,SOTALLO1.DATE_END FROM SOTALLO1  WHERE" & vbCrLf _
                & "  DATE_END >= '" & Format(Now, "dd-MMM-yyyy") & "'" & vbCrLf _
                & " And ITEM_CODE IN " & vbCrLf _
                & "(Select DISTINCT ITEM_CODE FROM " & ICTITEM1 & ")"
            Create_TDA(.Tables.Add, "SOTALLO1", "**", 0, False, "")


            ''sqlSPTCOOPX = "Select SPTCOOP1.*, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, SPTCOOP3.AUTH_LNO" & vbCrLf _
            ''    & ", SPTCOOP3.FEATURE_DESC, SPTCOOP3.ITEM_CODE" & vbCrLf _
            ''    & ", ICTCOLL1.COLLECTION_NAME, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME" & vbCrLf _
            ''    & " from SPTCOOP1,SPTCOOP3,ICTCOLL1,ICTBRAN1" & vbCrLf _
            ''    & " where ICTCOLL1.COLLECTION_CODE (+) = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
            ''    & "   and SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
            ''    & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
            ''    & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO"
            ''    & "(Select DISTINCT ITEM_CODE FROM " & ICTITEM1 & ")"
            ''ASCMAIN1.sql = sqlSPTCOOPX & "  and SPTCOOP1.OPS_YYYYPP = :PARM1"

            sqlSPTCOOPX = "Select SPTCOOP1.*, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, SPTCOOP3.AUTH_LNO, SPTCOOP3.ITEM_CODE" & vbCrLf _
                 & " from SPTCOOP1,SPTCOOP3" & vbCrLf _
                 & "   where SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
                 & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO" & vbCrLf _
                 & " and  DATE_END >= '" & Format(Now, "dd-MMM-yyyy") & "'" & vbCrLf _
                 & " And SPTCOOP3.ITEM_CODE In " & vbCrLf _
                 & "(Select DISTINCT ITEM_CODE FROM " & ICTITEM1 & ")"
            ASCMAIN1.sql = sqlSPTCOOPX & "  And SPTCOOP1.OPS_YYYYPP = :PARM1"


            Create_TDA(.Tables.Add, "SPTCOOP1", "**", 0, False, "V")


            .Tables("ICTSTATX").Columns.Add("WHSE_QTY_NETA", GetType(System.Int64), "ISNULL(WHSE_QTY_ON_HAND,0) - ISNULL(WHSE_QTY_PICK,0)")
            .Tables("ICTSTATX").Columns.Add("WHSE_QTY_ATS", GetType(System.Int64), "ISNULL(WHSE_QTY_NETA,0) - ISNULL(WHSE_QTY_OPEN,0) - ISNULL(WHSE_QTY_COMM,0)")

            ' NVL(ICTITEM1.ITEM_STATUS,'?') = 'A'"
            sqlSATMKOPX = "Select * from SATMKOP1"
            ASCMAIN1.sql = sqlSATMKOPX
            Create_TDA(.Tables.Add, "SATMKOPX", "**", 0, False)

            ASCMAIN1.sql = "Select * from SATMKOP1 where TRACKER_NO = :PARM1"
            Create_TDA(.Tables.Add, "SATMKOP1", "**", 0, True, "V")

            ASCMAIN1.sql = "Select * from SATMKOP2 where TRACKER_NO = :PARM1"
            Create_TDA(.Tables.Add, "SATMKOP2", "**", 0, True, "V")

            ASCMAIN1.sql = "Select SATMKOP3.*,ICTITEM1.ITEM_DESC DESC1 from SATMKOP3,ICTITEM1 where ICTITEM1.ITEM_CODE = SATMKOP3.ITEM_CODE AND TRACKER_NO = :PARM1"
            Create_TDA(.Tables.Add, "SATMKOP3", "**", 0, True, "V")
            With .Tables("SATMKOP3")
                ' ADD ALL FIELDS IN SPREADSHEET
                '      .Columns.Add("DESC")
                '  .Columns.Add("AL_NOS")
                .Columns.Add("ALL_QTYS")
                .Columns.Add("ALL_DATES")
                .Columns.Add("PO_NOS")
                .Columns.Add("PO_QTYS")
                .Columns.Add("PO_DATES")
                .Columns.Add("INV_NOS")
                .Columns("INV_NOS").DefaultValue = ""
                .Columns.Add("INV_QTYS")
                .Columns("INV_QTYS").DefaultValue = 0
                .Columns.Add("INV_DATES")
                .Columns("INV_DATES").DefaultValue = ""
                .Columns.Add("INV_ETADATES")
                .Columns("INV_ETADATES").DefaultValue = ""
                .Columns.Add("NOT_INV")
                .Columns("NOT_INV").DefaultValue = 0
                .Columns.Add("ON_HAND")
                .Columns("ON_HAND").DefaultValue = 0
                .Columns.Add("AVL2SELL")
                .Columns("AVL2SELL").DefaultValue = 0
                .Columns.Add("PROD_CODE")
                .Columns.Add("COST_CATGY_CODE")
                .Columns.Add("PROMO_CUST")
                .Columns.Add("PROMO_DATES")
                .Columns.Add("PROMO_BOOKING_NAME")

                'INV_ETA_TEXT
            End With




            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.PROD_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.COST_CATGY_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.HC_CODE, NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) ITEM_EAN_CODE, ICTITEM1.ITEM_ALT_SORT" & vbCrLf _
                & ", ICTITEM1.ITEM_STD_PACK_SLS, ICTITEM1.ITEM_SO_QTY_MIN, ICTITEM1.ITEM_SO_QTY_MULT" & vbCrLf _
                & ", ICTITEM1.ITEM_CODE_COMPARE_TO" & vbCrLf _
                & " from ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & "   and NVL(ICTITEM1.ITEM_STATUS,'?') = 'A'"
            sqlICTITEM1 = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 0)




        End With
        grdSATMKOPX.DataSource = dst.Tables("SATMKOPX")
        grdSATMKOP2.DataSource = dst.Tables("SATMKOP2")
        grdSATMKOP3.DataSource = dst.Tables("SATMKOP3")
        grdICTITEM1.DataSource = dst.Tables("ICTITEM1")


        Create_Summary(grdSATMKOPX, "TRACKER_NO", "Count")
        Create_Summary(grdSATMKOP2, "TRACKER_NO", "Count")



        With grdSATMKOP3.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("DESC1").Header.Fixed = True

        End With

        With grdICTITEM1.DisplayLayout.Bands(0)
            .Columns("ITEM_CODE").Header.Fixed = True
            .Override.AllowAddNew = AllowAddNew.No
            .Override.AllowUpdate = DefaultableBoolean.False
            .Override.AllowDelete = DefaultableBoolean.False
        End With




        '   grdSATMKOP2.DisplayLayout.Bands(0).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False


        For Each COLUMN_NAME As String In New String() _
                {"PO_NOS", "PO_QTYS", "PO_DATES"}

            With grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Appearance
                .BackColor = System.Drawing.Color.White
                .BackColor2 = Drawing.Color.LightSkyBlue
                .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
            ' If COLUMN_NAME = "CUST_STORE_NAME" Then grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
        Next

        For Each COLUMN_NAME As String In New String() _
                {"INV_NOS", "INV_QTYS", "INV_DATES", "INV_ETADATES"}

            With grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Appearance
                .BackColor = System.Drawing.Color.White
                .BackColor2 = Drawing.Color.LightGreen
                .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
            ' If COLUMN_NAME = "CUST_STORE_NAME" Then grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
        Next


        For Each COLUMN_NAME As String In New String() _
                {"NOT_INV", "ON_HAND", "AVL2SELL", "ALL_QTYS", "ALL_DATES"}

            With grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Appearance
                .BackColor = System.Drawing.Color.White
                .BackColor2 = Drawing.Color.LightPink
                .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
            ' If COLUMN_NAME = "CUST_STORE_NAME" Then grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
        Next

        For Each COLUMN_NAME As String In New String() _
                {"NOT_INV", "ON_HAND", "AVL2SELL", "ALL_QTYS", "ALL_DATES"}

            With grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Appearance
                .BackColor = System.Drawing.Color.White
                .BackColor2 = Drawing.Color.LightPink
                .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
            ' If COLUMN_NAME = "CUST_STORE_NAME" Then grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
        Next

        For Each COLUMN_NAME As String In New String() _
                {"PROMO_CUST", "PROMO_DATES", "PROMO_BOOKING_NAME"}

            With grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Appearance
                .BackColor = System.Drawing.Color.White
                .BackColor2 = Drawing.Color.Lavender
                .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
            End With
            ' If COLUMN_NAME = "CUST_STORE_NAME" Then grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
        Next

        With grdICTITEM1.DisplayLayout.Bands("ICTITEM1")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                    If New String() {"QTY_ALLO", "ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC", "QTY_LEFT", "QTY_BAL", "QTY_OVER"}.Contains(gcol.Key) Then
                        .BackColor2 = System.Drawing.Color.LightBlue
                    Else
                        .BackColor2 = System.Drawing.Color.LightGreen
                    End If
                End With
            Next

        End With


        With grdSATMKOP2.DisplayLayout.Bands("SATMKOP2")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            '  .Columns("TRACKING_NO").Header.Fixed = True
        End With

        With grdSATMKOPX.DisplayLayout.Bands("SATMKOPX")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal

                If New String() {"TRACKER_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor = System.Drawing.Color.LightGreen
                ElseIf New String() {"COLLECTION_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            ' .Columns("TRACKER_NO").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdSATMKOP2, "GROUP_TYPE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SATMKOP2' and COLUMN_NAME = 'GROUP_TYPE'")

        grpHeader.Visible = False
        SplitContainer2.Panel2Collapsed = True ' until we need more header data

        Show_Filter(grdSATMKOPX, False)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"

                If Absx1.txtFor("TRACKER_DESC").Text.Length = 0 Then
                    EMsg &= vbCr & "You must enter a Description for this Tracking No"
                End If

            Case "View", "Edit"
                'TRACKER_NO = Absx1.txtFor("TRACKER_NO").Text
                'If TRACKER_NO = "" Then
                '    EMsg &= vbCr & "You must specify a Tracking No to View"
                'Else

                'End If

                If Absx1.txtFor("TRACKER_NO").Text = "" Then
                    EMsg &= vbCr & "No Tracker No Specified"
                Else
                    TRACKER_NO = Absx1.txtFor("TRACKER_NO").Text
                    rowSATMKOP1 = LookUp("SATMKOP1", TRACKER_NO)
                    If rowSATMKOP1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Tracker No (" & TRACKER_NO & ")"
                    Else
                        'VEND_CODE = rowPOTORDR1.Item("VEND_CODE")
                    End If
                End If



            Case "Update"

                If Absx1.txtFor("TRACKER_NO").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Tracking No"
                Else
                    '  Dim row As DataRow = LookUp("SATMKOP1", Absx1.txtFor("TRACKER_NO").Text)
                    ' If IsNothing(row) Then
                    'EMsg &= vbCr & "Tracking No Entered Is Not Valid"
                    'Else
                    'End If
                End If


                If Absx1.txtFor("TRACKER_DESC").Text.Length = 0 Then
                    EMsg &= vbCr & "You must enter a Description for this Tracking No"
                End If


            Case "Generate Excel"
                If MsgBox("Are you sure you want to produce Spreadsheet?", MsgBoxStyle.YesNo,
                          "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Save As New"
                If MsgBox("Are you sure you want to Create a New Tracking No From this Tracking No", MsgBoxStyle.YesNo,
                          "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If





            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo,
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
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
                CALCULATE_DATA()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Generate Excel"
                CALCULATE_DATA()
                Generate_Excel()
                Mode_Settings(False)
                Absx1.txtFor("TRACKER_NO").Text = TRACKER_NO
                Click_Command("View")

            Case "Save As New"
                CREATE_NEW_TRACKING()
                EntryMode = "V"
                Load_Record()
                CALCULATE_DATA()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)
        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode

                    If EntryMode = "V" And ScreenMode Then ' And (optShow.Value <> "G") Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
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
                End With

                'With .Groups("Excel Options")
                '    If EntryMode = "V" And ScreenMode Then ' And (optShow.Value <> "G") Then
                '        .Items("Generate Excel").Visible = True And Not InquiryMode
                '    Else
                '        .Items("Generate Excel").Visible = False
                '    End If

                'End With




                If (EntryMode = "V") And ScreenMode Then ' And (optShow.Value <> "G") Then
                        UltraExplorerBar1.Groups("Excel Options").Visible = True And Not InquiryMode
                    Else
                        UltraExplorerBar1.Groups("Excel Options").Visible = False
                    End If

                ''If (EntryMode = "V" Or EntryMode = "E") And ScreenMode Then
                ''    grdClickCellButton(grdSATMKOP2, "", True)

                ''End If

                For Each COLUMN_NAME As String In New String() _
             {"PO_NOS", "PO_QTYS", "PO_DATES", "INV_NOS", "INV_QTYS", "INV_DATES", "INV_ETADATES",
              "NOT_INV", "ON_HAND", "AVL2SELL", "PROD_CODE", "COST_CATGY_CODE", "ALL_QTYS", "ALL_DATES", "PROMO_CUST", "PROMO_DATES", "PROMO_BOOKING_NAME"}

                    If (EntryMode = "V") And ScreenMode Then ' And (optShow.Value <> "G") Then
                        grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
                    Else
                        grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
                    End If
                    If COLUMN_NAME = "COST_CATGY_CODE" Or COLUMN_NAME = "PROD_CODE" Then
                        grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
                    End If
                    ' If COLUMN_NAME = "CUST_STORE_NAME" Then grdSATMKOP3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
                Next




            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode
        grdSATMKOPX.Visible = Not ScreenMode

        If ScreenMode Then

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            If EntryMode = "E" Or EntryMode = "N" Then
                Set_Read_Only_for_ctl(Absx1.txtFor("TRACKER_DESC"), False)
            End If

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATMKOP2, grdSATMKOP3}
                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name <> "grdSATMKOP3" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            '    .AllowUpdate = DefaultableBoolean.True
                        End If
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                        SplitContainer4.Panel2Collapsed = False ' until we need more header data

                    End With
                Else
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                        SplitContainer4.Panel2Collapsed = True ' until we need more header data

                    End With
                End If
            Next


        Else
            Clear_Record()
            '     Fill_Records("SATMKOP1", TRACKER_NO)
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SATMKOP1", "SATMKOP2", "SATMKOP3", "ICTITEM1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        SplitContainer2.Panel2Collapsed = True ' until we need more header data

        ASCMAIN1.sql = sqlSATMKOPX
        Fill_Records("SATMKOPX", "", True, ASCMAIN1.sql)
        Sort_grdColumns(grdSATMKOPX, "TRACKER_NO".ToLower)

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowSATMKOP1 = dst.Tables("SATMKOP1").NewRow
            TRACKER_NO = ASCMAIN1.Next_Control_No("SATMKOP1.TRACKER_NO")
            rowSATMKOP1.Item("TRACKER_NO") = TRACKER_NO
            rowSATMKOP1.Item("TRACKER_DESC") = HFs("TRACKER_DESC")
            rowSATMKOP1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSATMKOP1.Item("INIT_DATE") = DATETIME_STAMP
            rowSATMKOP1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSATMKOP1.Item("LAST_DATE") = DATETIME_STAMP


            dst.Tables("SATMKOP1").Rows.Add(rowSATMKOP1)

        Else
            rowSATMKOP1 = Fill_Record("SATMKOP1", Absx1.txtFor("TRACKER_NO").Text)
            '  dst.AcceptChanges()
        End If


        '   STATUS_CODE = rowSATMKOP1.Item("STATUS_CODE")

        EnforceConstraints(False)

        '    Fill_Records("SATMKOP1", TRACKER_NO)

        Sort_grdColumns(grdSATMKOPX, "TRACKER_NO")



        Fill_Records("SATMKOP2", TRACKER_NO)
        Sort_grdColumns(grdSATMKOP2, "TRACKER_NO")

        Fill_Records("SATMKOP3", TRACKER_NO)
        Sort_grdColumns(grdSATMKOP3, "TRACKER_NO")
        If EntryMode <> "N" And grdSATMKOP2.Rows.Count <> 0 Then
            SETUP_ITEMS()
        End If


        SplitContainer2.Panel1Collapsed = True ' until we need more header data
        EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        For Each rowSATMKOP2 As DataRow In dst.Tables("SATMKOP2").Select("")
            Dim GROUP_TYPE As String = rowSATMKOP2.Item("GROUP_TYPE") & ""

            If GROUP_TYPE = "B" Then
                rowSATMKOP2.Item("HC_CODE") = ""
                rowSATMKOP2.Item("COLLECTION_CODE") = ""
            ElseIf GROUP_TYPE = "H" Then
                rowSATMKOP2.Item("BRAND_CODE") = ""
                rowSATMKOP2.Item("COLLECTION_CODE") = ""
            ElseIf GROUP_TYPE = "C" Then
                rowSATMKOP2.Item("HC_CODE") = ""
                rowSATMKOP2.Item("BRAND_CODE") = ""
            End If

        Next

        For Each r As DataRow In dst.Tables("SATMKOP1").Select("")
            r("LAST_OPER") = ASCMAIN1.USER_ID
            r("LAST_DATE") = DATETIME_STAMP
        Next

        BeginTrans()

        '   If EntryMode <> "N" Then Delete_Records()

        'Dim SQLD As String = "TRACKER_NO = '" & TRACKER_NO & "'"
        '' INIT_LAST("SPTSFOC1", False, , True)

        'Update_Record_TDA("SATMKOP1", SQLD)
        'Update_Record_TDA("SATMKOP2", SQLD)
        'Update_Record_TDA("SATMKOP3", SQLD)


        'Fill_Records("SATMKOP1", TRACKER_NO)


        Update_Record_TDA("SATMKOP1")
        Update_Record_TDA("SATMKOP2")
        Update_Record_TDA("SATMKOP3")

        Update_Record_TDA("ASTAUDT1")

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
            {"SATMKOP1", "SATMKOP2", "SATMKOP3"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where TRACKER_NO = '" & TRACKER_NO & "'"
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
                Absx1.txtFor("EVENT_GROUP_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SATMKOP1"
            E.COLUMN_NAME = "TRACKER_NO"
            E.CODE_VALUE = Absx1.txtFor("TRACKER_NO").Text
            E.DESC_VALUE = Absx1.txtFor("TRACKER_DESC").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        'E.TABLE_NAME = "SPTSFOC1"
        'E.TABLE_KEY_CAPTION = "Store Focus Events"
        'If ScreenMode Then
        '    E.enabled = True
        '    E.read_only = False
        '    E.TABLE_KEY = Absx1.txtFor("EVENT_GROUP_NO").Text '  HFs("CUST_CODE")
        '    E.TABLE_KEY_DESC = Absx1.txtFor("VEHICLE_CODE").Text
        '    E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E")
        'Else
        '    E.enabled = False
        '    E.read_only = True
        '    E.TABLE_KEY_locked = False
        '    E.TABLE_KEY = ""
        'End If

        Return E
    End Function

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "VEHICLE_CODE"
                sql_where = "VEHICLE_CODE in ('BF','MA')"
        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        '  Load_Popup_Menu(grdSPTSFOCX, "SS", "Show Filter", "Show GroupBox", "Approve") ', "Move to Pending")
        '   Load_Popup_Menu(grdSPTSFOC9, "SSBBBB", "Show Filter", "Show GroupBox", "Load Stores w/Attribute", "Load All Stores", "Select All", "De-Select All")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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

        Select Case e.SourceControl.Name

            'Case "grdSPTSFOC9"
            '    tlb_btn = tlb_pop.Tools("Load All Stores")
            '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
            '        tlb_btn.SharedProps.Visible = True
            '    Else
            '        tlb_btn.SharedProps.Visible = False
            '    End If
            '    tlb_btn = tlb_pop.Tools("Load Stores w/Attribute")
            '    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
            '        tlb_btn.SharedProps.Visible = True
            '    Else
            '        tlb_btn.SharedProps.Visible = False
            '    End If

            Case "grdSPTSFOCX"

                'If Not InquiryMode Then ' And (optShow.Value = "P" Or optShow.Value = "G") Then
                '    'tlb_pop.Tools("Move to Pending").SharedProps.Visible = (optShow.Value = "P")
                '    tlb_pop.Tools("Approve").SharedProps.Visible = True
                'Else
                '    'tlb_pop.Tools("Move to Pending").SharedProps.Visible = False
                '    tlb_pop.Tools("Approve").SharedProps.Visible = False
                'End If



        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdSPTSFOC9"
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

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now executing " & e.Tool.Key)

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Load All Stores"



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
            Case "TRACKER_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If

            Case "ITEM_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Me.ProcessEnterKeyStroke(Absx1.txtFor("ITEM_CODE").Text.Trim)
                    '   timItemCode.Start()
                End If
        End Select

    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.txt_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "VEHICLE_CODE"
                Dim VEHICLE_CODE As String = Absx1.txtFor("VEHICLE_CODE").Text
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "EVENT_GROUP_NO"
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
            Case "TOTAL_AMT"
                ' Calculate_OPEN_AMT()
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Select Case Absx1.GetABSColumnName(sender)
            Case "APPR_STATUS_CODE"
                If Absx1.optFor("APPR_STATUS_CODE").Value = "X" Then
                    Absx1.optFor("STATUS_CODE").Value = "C"
                Else

                End If

        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "AUTH_DATE"
                If Absx1.dteFor("AUTH_DATE").Value & "" = "" Then
                    Absx1.txtFor("OPS_YYYYWW").Text = ""
                Else
                    Dim DATE_START As Date = Absx1.dteFor("AUTH_DATE").Value
                    If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                        ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(DATE_START, "dd-MMM-yyyy") & "'"
                        Dim YW As String = ASCDATA1.GetDataValue
                        If YW <> "" Then
                            Absx1.txtFor("OPS_YYYYWW").Text = YW
                        End If
                    End If
                End If
        End Select
    End Sub
#End Region

#Region "grdSPTSFOC3"

    Private Sub grdSPTSFOC3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
        'Select Case e.Cell.Column.Key
        '    Case "ITEM_CODE"
        '        '   If cdr IsNot Nothing Then
        '        Dim ITEM_CODE As String = CStr(e.Cell.Value & "").ToUpper
        '        ' Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)

        '        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        '        If rowICTITEM1 IsNot Nothing Then
        '            If e.Cell.Value & "" <> ITEM_CODE Then
        '                e.Cell.Row.Cells("ITEM_CODE").Value = ITEM_CODE
        '            End If
        '            e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
        '            e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
        '            e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
        '        End If

        '        '   End If

        '    Case "COLLECTION_CODE"

        '        grdCodeDesc(grdSATMKOP2, "ICTCOLL1", "COLLECTION_CODE", "COLLECTION_NAME")
        '        If cdr IsNot Nothing Then
        '            Dim COLLECTION_CODE As String = e.Cell.Value
        '            Dim BRAND_CODE As String = cdr.Item("BRAND_CODE")
        '            Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", BRAND_CODE)
        '            e.Cell.Row.Cells("BRAND_CODE").Value = BRAND_CODE
        '            e.Cell.Row.Cells("BRAND_NAME").Value = rowICTBRAN1.Item("BRAND_NAME")
        '        Else
        '            grdSATMKOP2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
        '        End If

        'End Select
    End Sub

    Private Sub grdSPTSFOC3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs)
        'With grdSATMKOP2.DisplayLayout.Bands(0)
        '    If grdSATMKOP2.ActiveRow.IsAddRow Then
        '        .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
        '        grdSATMKOP2.ActiveCell = grdSATMKOP2.ActiveRow.Cells("COLLECTION_CODE")
        '        grdSATMKOP2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
        '    Else
        '        .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
        '    End If
        'End With
    End Sub

    Private Sub grdSPTSFOC3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs)
        'If grdSATMKOP2.ActiveCell Is Nothing Then Exit Sub
        'With grdSATMKOP2.ActiveCell
        '    Select Case .Column.Key
        '        Case "COLLECTION_CODE"
        '            If .Text <> "" Then
        '                If .Value IsNot Nothing Then
        '                    .Value = .Text.ToUpper
        '                End If

        '            End If
        '            If .Text <> "" Then
        '                cdr = LookUp("ICTCOLL1", .Text)
        '                If cdr Is Nothing Then
        '                    ASCMAIN1.Progress("Invalid Collection Code (" & .Text & ")")
        '                    If .Value IsNot Nothing Then
        '                        .Value = ""
        '                    End If
        '                    e.Cancel = True
        '                End If
        '            End If
        '    End Select
        'End With
    End Sub

    Private Sub grdSPTSFOC3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs)
        'With grdSATMKOP2
        '    If e.Row.Cells("COLLECTION_CODE").Text = "" Then
        '        '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
        '        ' e.Cancel = True MAGIC UNICORN IS A COMMENTED LINE
        '    Else
        '        LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Text)
        '        If cdr Is Nothing Then
        '            MsgBox("Invalid Value entered for Collection Code (" & e.Row.Cells("COLLECTION_CODE").Text & ")",
        '                   MsgBoxStyle.OkOnly, "Cannot Update Row")
        '            e.Cancel = True
        '        End If
        '    End If

        '    If e.Cancel Then
        '        e.Row.CancelUpdate()
        '    End If

        '    If Not e.Cancel Then
        '        If e.Row.Cells("EVENT_GROUP_NO").Text = "" Then
        '            .ActiveRow.Cells("EVENT_GROUP_NO").Value = Absx1.CtlFor("EVENT_GROUP_NO").Text
        '            .ActiveRow.Cells("EVENT_GROUP_LNO").Value = Val(dst.Tables("SPTSFOC3").Compute("Max(EVENT_GROUP_LNO)", "") & "") + 1
        '        End If
        '    End If
        'End With
    End Sub

    Private Sub grdSPTSFOC3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)

        'If grdSATMKOP2.ActiveRow Is Nothing Then Exit Sub

        'Dim sql_where As String = ""
        'Select Case e.Cell.Column.Key
        '    Case "ITEM_CODE"
        '    Case "COLLECTION_CODE"
        '        sql_where = "COLLECTION_STATUS = 'A'"
        'End Select
        'grdClickCellButton(grdSATMKOP2, sql_where, False)

    End Sub



#End Region

    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs)
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

    Private Sub grdSATMKOPX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSATMKOPX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("TRACKER_NO").Text = e.Row.Cells("TRACKER_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdSATMKOP2_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSATMKOP2.InitializeLayout

    End Sub

    Private Sub grdSATMKOP2_ClickCell(sender As Object, e As ClickCellEventArgs) Handles grdSATMKOP2.ClickCell

    End Sub

    Private Sub ProcessEnterKeyStroke(ByVal scannedData As String)

        Dim ITEM_CODE As String = scannedData.Trim
        Dim rowSOTRTRN2 As DataRow = Nothing
        Dim itemFound As Boolean = True
        Dim tblICTITEM1 As DataTable = Nothing

        If dst.Tables("SATMKOP3").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length = 0 Then
            ' VALIDATE AND ADD ITEM
            tblICTITEM1 = ASCDATA1.GetDataTable("Select * from ICTITEM1 WHERE ITEM_CODE = :PARM1", "", "V", ITEM_CODE)

            If tblICTITEM1.Rows.Count = 0 Then
                tblICTITEM1 = ASCDATA1.GetDataTable("Select * from ICTITEM1 WHERE ITEM_UPC_CODE = :PARM1", "", "V", ITEM_CODE)
                If tblICTITEM1.Rows.Count = 0 Then
                    MessageBox.Show("Invalid Item (" & ITEM_CODE & ").", "Validate Item", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    itemFound = False
                ElseIf tblICTITEM1.Rows.Count > 1 Then
                    MessageBox.Show("Multiple Items found for the provided UPC Code(" & ITEM_CODE & ").", "Validate Item", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    itemFound = False
                Else
                    ITEM_CODE = tblICTITEM1.Rows(0).Item("ITEM_CODE")
                End If
            End If

            If itemFound Then
                Dim GROUP_ID As String = grdSATMKOP2.ActiveRow.Cells("GROUP_ID").Value & ""



                If dst.Tables("SATMKOP3").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length = 0 Then
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE, True)

                    grdSATMKOP3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                    grdSATMKOP3.DisplayLayout.Bands(0).AddNew()
                    With grdSATMKOP3.ActiveRow
                        .Cells("ITEM_CODE").Value = ITEM_CODE
                        .Cells("DESC1").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                        .Cells("TRACKER_NO").Value = TRACKER_NO
                        .Cells("GROUP_ID").Value = GROUP_ID
                        .Update()
                    End With
                    grdSATMKOP3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                End If
            End If
        End If

        txtITEM_CODE.Clear()
    End Sub



    Private Sub grdICTITEM1_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdICTITEM1.DoubleClickRow
        If ScreenMode Then

            Dim ITEM_CODE As String = e.Row.Cells("ITEM_CODE").Value
            Dim GROUP_ID As String = grdSATMKOP2.ActiveRow.Cells("GROUP_ID").Value & ""


            Dim rowSATMKOP3s() As DataRow = dst.Tables("SATMKOP3").Select("ITEM_CODE = '" & ITEM_CODE & "'")
            If rowSATMKOP3s.Length = 0 Then
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE, True)

                grdSATMKOP3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.Yes
                grdSATMKOP3.DisplayLayout.Bands(0).AddNew()
                With grdSATMKOP3.ActiveRow
                    .Cells("ITEM_CODE").Value = ITEM_CODE
                    .Cells("DESC1").Value = rowICTITEM1.Item("ITEM_DESC") & ""
                    .Cells("TRACKER_NO").Value = TRACKER_NO
                    .Cells("GROUP_ID").Value = GROUP_ID
                    .Update()
                End With
                '  grdSATMKOP3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            Else
                If rowSATMKOP3s.Length = 1 Then
                    If rowSATMKOP3s(0).Item("GROUP_ID") = GROUP_ID Then
                        MsgBox($"Item {ITEM_CODE} already exists in Group ID {GROUP_ID}", MsgBoxStyle.YesNo, "Verification")
                    Else
                        If MsgBox($"Item {ITEM_CODE} already exists in Group ID {rowSATMKOP3s(0).Item("GROUP_ID")}" & vbCrLf & vbCrLf & $"OK to move this item to Group ID {GROUP_ID}?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            rowSATMKOP3s(0).Item("GROUP_ID") = GROUP_ID
                            For Each grow As UltraWinGrid.UltraGridRow In grdSATMKOP3.Rows
                                If grow.Cells("ITEM_CODE").Value = ITEM_CODE Then
                                    grow.Activate()
                                End If
                            Next
                        End If
                    End If
                Else
                    MsgBox("Item already exists multiple times in Tracker " & TRACKER_NO, MsgBoxStyle.OkOnly, "Please Call ABS")
                End If
            End If


            'Dim rowSOTALLO1 As DataRow = Add_Item(ITEM_CODE, True)
            'If rowSOTALLO1 IsNot Nothing Then
            '    Add_Allocation_to_Grid(rowSOTALLO1)
            '    MsgBox("A New Allocation record has been added for Item " & ITEM_CODE, MsgBoxStyle.OkOnly, "Verification")
            'End If
        End If
    End Sub
    Sub Generate_Excel()
        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim oSheetMAST As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        'Dim rangeCopyFrom As SpreadsheetGear.IRange
        'Dim rangePaste_To As SpreadsheetGear.IRange


        Dim Start_Row As Integer = 5

        Dim FILENAME_TEMPLATE As String = "OPSTRACKER.xlsx"

        Dim FILENAME_SOURCE As String = ""

        If ASCMAIN1.Running_in_VS And (ASCMAIN1.USER_ID = "wjz" Or ASCMAIN1.USER_ID = "dgj") Then
            FILENAME_SOURCE = "C:\INT\templates\" & FILENAME_TEMPLATE
        Else
            FILENAME_SOURCE = ASCMAIN1.Folders("SharedRoot") & "Templates\" & FILENAME_TEMPLATE
        End If

        Dim XLS_FILENAME As String = ""

        ASCMAIN1.Progress("Now Creating Custom XLS Workbook")

        If FILENAME_TEMPLATE = "" Then
            oWB = SpreadsheetGear.Factory.GetWorkbook()
            oSheet = oWB.Worksheets.Add
            oSheet.Name = "DGJ"
            XLS_FILENAME = ASCMAIN1.Folders("Work") & XNO & ".xlsx"
            oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Else
            Dim success As Boolean = False
            Dim XLS_NO As Integer = 0
            Do Until success
                Try
                    XLS_NO += 1
                    XLS_FILENAME = ASCMAIN1.Folders("Work") & "OPS Tracker" & TRACKER_NO
                    XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"
                    FileCopy(FILENAME_SOURCE, XLS_FILENAME)
                    success = True


                Catch ex As Exception
                    ' Stop
                End Try
                If XLS_NO = 10 Then
                    Exit Sub
                End If
            Loop
            oWB = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME)
            oSheetMAST = oWB.Worksheets(0)
        End If


        '  Dim NEW_SHEET As Boolean = False

        For Each rowSATMKOP2 As DataRow In dst.Tables("SATMKOP2").Select("", "TRACKER_NO")
            Dim TRACKER_NO As String = rowSATMKOP2.Item("TRACKER_NO") & ""
            Dim GROUP_ID As String = rowSATMKOP2.Item("GROUP_ID") & ""
            Dim GROUP_DESC As String = rowSATMKOP2.Item("GROUP_DESC") & ""

            ' CREATE NEW SHEET IF NEW_SHEET = true
            oSheet = oWB.Worksheets.Add
            oSheet.Name = GROUP_ID

            oSheetMAST.Cells("A:Z").Copy(oSheet.Cells("A:Z"))
            With oSheet.Cells(0, 0)
                .Font.Color = SpreadsheetGear.Colors.White
                .Font.Size = 20
                .Value = GROUP_DESC
            End With

            'RIP THROUGH Details

            Dim rx As Integer = 4
            Dim cx As Integer = 0

            Dim PROD_CODE_DISPLAY As String = ""

            For Each rowSATMKOP3 As DataRow In dst.Tables("SATMKOP3").Select("GROUP_ID = '" & GROUP_ID & "'", "COST_CATGY_CODE DESC, PROD_CODE")
                rx += 1
                cx = -1

                Dim PROD_CODE As String = rowSATMKOP3.Item("PROD_CODE")
                Dim rowICTPROD1 As DataRow = LookUp("ICTPROD1", PROD_CODE, True)

                If PROD_CODE_DISPLAY <> PROD_CODE Then
                    oSheet.Cells(rx, 0).Value = rowICTPROD1.Item("PROD_CODE") & ""
                    oSheet.Cells(rx, 0).Font.Color = SpreadsheetGear.Colors.White
                    oSheet.Cells(rx, 0).Interior.Color = SpreadsheetGear.Colors.Black
                    rx += 1
                End If
                PROD_CODE_DISPLAY = PROD_CODE



                Dim ITEM_CODE As String = rowSATMKOP3.Item("ITEM_CODE")
                '   Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE, True)
                cx += 1 : oSheet.Cells(rx, cx).Value = rowSATMKOP3.Item("DESC1") & ""
                cx += 1 : oSheet.Cells(rx, cx).Value = ITEM_CODE
                cx += 2
                cx += 1 : oSheet.Cells(rx, cx).Value = rowSATMKOP3.Item("NEEDED_QTYS") & ""
                oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.Beige
                cx += 1 : oSheet.Cells(rx, cx).Value = rowSATMKOP3.Item("US_SHIP_DATE") & ""
                oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.Beige

                '          Dim ALL_TEXT As String = rowSATMKOP3.Item("ALL_NOS") & ""
                Dim ALL_DATE_TEXT As String = rowSATMKOP3.Item("ALL_DATES") & ""
                Dim ALL_QTY_TEXT As String = rowSATMKOP3.Item("ALL_QTYS") & ""

                If ALL_QTY_TEXT <> "" Then
                    'range = oSheet.Cells(rx, cx + 1)
                    'range.WrapText = True

                    'cx += 1 : oSheet.Cells(rx, cx).Value = ALL_TEXT
                    'oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightSkyBlue

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = ALL_QTY_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = ALL_DATE_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink
                Else
                    cx += 2
                End If


                '   cx += 2

                'PO DATA


                Dim PO_TEXT As String = rowSATMKOP3.Item("PO_NOS") & ""
                Dim PO_DATE_TEXT As String = rowSATMKOP3.Item("PO_DATES") & ""
                Dim PO_QTY_TEXT As String = rowSATMKOP3.Item("PO_QTYS") & ""


                ''For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                ''    PO_TEXT = PO_TEXT & vbLf & rowPOTORDR2.Item(1) & ""
                ''    PO_QTY_TEXT = PO_QTY_TEXT & vbLf & rowPOTORDR2.Item(2) & ""
                ''    PO_DATE_TEXT = PO_DATE_TEXT & vbLf & rowPOTORDR2.Item(3) & ""
                ''Next

                If PO_TEXT <> "" Then
                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = PO_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightSkyBlue

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = PO_QTY_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightSkyBlue

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = PO_DATE_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightSkyBlue
                Else
                    cx += 3
                End If

                Dim PROMO_CUST As String = rowSATMKOP3.Item("PROMO_CUST") & ""
                Dim PROMO_DATE As String = rowSATMKOP3.Item("PROMO_DATES") & ""
                Dim PROMO_BOOKING_NAME As String = rowSATMKOP3.Item("PROMO_BOOKING_NAME") & ""


                If PROMO_DATE <> "" Then
                    'range = oSheet.Cells(rx, cx + 1)
                    'range.WrapText = True

                    'cx += 1 : oSheet.Cells(rx, cx).Value = ALL_TEXT
                    'oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightSkyBlue

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = PROMO_BOOKING_NAME
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.Lavender

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = PROMO_DATE
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.Lavender

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = PROMO_CUST
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.Lavender

                Else
                    cx += 3
                End If

                'cx += 3
                ' ICTPINV1

                Dim INV_TEXT As String = rowSATMKOP3.Item("INV_NOS") & ""
                Dim INV_DATE_TEXT As String = rowSATMKOP3.Item("INV_DATES") & ""
                Dim INV_QTY_TEXT As String = rowSATMKOP3.Item("INV_QTYS") & ""
                Dim INV_ETA_TEXT As String = rowSATMKOP3.Item("INV_ETADATES") & ""
                Dim TOT_INV As Double = 0



                ''For Each rowICTPINV2 As DataRow In dst.Tables("ICTPINV2").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                ''    INV_TEXT = INV_TEXT & vbLf & rowICTPINV2.Item("INV_NUM") & ""
                ''    INV_DATE_TEXT = INV_DATE_TEXT & vbLf & rowICTPINV2.Item("INV_DATE") & ""
                ''    INV_QTY_TEXT = INV_QTY_TEXT & vbLf & rowICTPINV2.Item("PINV_QTY") & ""
                ''    INV_ETA_TEXT = INV_ETA_TEXT & vbLf & rowICTPINV2.Item("ETA_DATE") & ""
                ''    TOT_INV = TOT_INV + Val(rowICTPINV2.Item("PINV_QTY") & "")
                ''Next

                If INV_TEXT <> "" Then
                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = INV_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightGreen
                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = INV_DATE_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightGreen

                    range = oSheet.Cells(rx, cx + 1)
                    range.WrapText = True

                    cx += 1 : oSheet.Cells(rx, cx).Value = INV_QTY_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightGreen

                    cx += 1 : oSheet.Cells(rx, cx).Value = INV_ETA_TEXT
                    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightGreen
                Else
                    cx += 4

                End If
                ' On Hand
                Dim WHSE_QTY_NOT_INV As String = ""
                Dim WHSE_QTY_ON_HAND As String = ""
                Dim WHSE_QTY_ATS As String = ""
                Dim AMT_NOT_INV As Double = 0


                'ASCMAIN1.sql = " ITEM_CODE = '" & ITEM_CODE & "'"
                'Dim rows() As DataRow = dst.Tables("ICTSTATX").Select(ASCMAIN1.sql)
                'If ROWs.Length = 0 Then
                '    EMsg &= vbCr & "No Record of (Apply-To) Invoice "
                'End If

                AMT_NOT_INV = Val(rowSATMKOP3.Item("NOT_INV") & "") - TOT_INV
                WHSE_QTY_NOT_INV = AMT_NOT_INV
                WHSE_QTY_ON_HAND = rowSATMKOP3.Item("ON_HAND") & ""
                WHSE_QTY_ATS = rowSATMKOP3.Item("AVL2SELL") & ""

                cx += 1 : oSheet.Cells(rx, cx).Value = AMT_NOT_INV
                oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink
                cx += 4
                cx += 1 : oSheet.Cells(rx, cx).Value = WHSE_QTY_ON_HAND
                oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink
                cx += 1 : oSheet.Cells(rx, cx).Value = WHSE_QTY_ATS
                oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink

                cx += 1 : oSheet.Cells(rx, cx).Value = rowSATMKOP3.Item("COMMENTS") & ""
                oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.Beige


                '  rowSATMKOP3.Item("NEEDED_QTYS") & ""


                ''For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                ''    AMT_NOT_INV = Val(rowICTSTATX.Item("WHSE_QTY_ONPO") & "") - TOT_INV
                ''    WHSE_QTY_NOT_INV = AMT_NOT_INV
                ''    WHSE_QTY_ON_HAND = rowICTSTATX.Item("WHSE_QTY_ON_HAND") & ""
                ''    WHSE_QTY_ATS = rowICTSTATX.Item("WHSE_QTY_ATS") & ""

                ''    cx += 1 : oSheet.Cells(rx, cx).Value = AMT_NOT_INV
                ''    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink
                ''    cx += 4
                ''    cx += 1 : oSheet.Cells(rx, cx).Value = WHSE_QTY_ON_HAND
                ''    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink
                ''    cx += 1 : oSheet.Cells(rx, cx).Value = WHSE_QTY_ATS
                ''    oSheet.Cells(rx, cx).Interior.Color = SpreadsheetGear.Colors.LightPink
                ''Next

            Next
        Next

        oWB.Worksheets(0).Delete()

        oWB.Save()
        Show_Document(XLS_FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


    End Sub

    Sub CALCULATE_DATA()

        Dim SQL1 As String = "SELECT DISTINCT ITEM_CODE FROM SATMKOP3 WHERE TRACKER_NO = " & TRACKER_NO
        ASCDATA1.ExecuteSQL("Delete from " & ICTITEM1)
        ASCDATA1.ExecuteSQL("Insert into " & ICTITEM1 & " " & SQL1)

        Fill_Records("POTORDR2")
        Fill_Records("ICTPINV2")
        Fill_Records("ICTSTATX")
        Fill_Records("SOTALLO1")





        Dim YP As String = ASCMAIN1.CYP
        Fill_Records("SPTCOOP1", YP)



        For Each ROW As DataRow In dst.Tables("ICTPINV2").Select("")
            If ROW.Item("ETA_DATE") & "" <> "" Then
                Dim ETA_DATE As Date = ROW.Item("ETA_DATE")
                Dim ETA_DATE_DC As Date = ETA_DATE
                For I As Integer = 1 To 5
                    ETA_DATE_DC = ETA_DATE_DC.AddDays(1)
                    If ETA_DATE_DC.DayOfWeek = DayOfWeek.Saturday Or ETA_DATE_DC.DayOfWeek = DayOfWeek.Sunday Then
                        I = I - 1
                    End If
                Next
                ROW.Item("ETA_DATE") = ETA_DATE_DC
            End If
        Next


        For Each rowSATMKOP3 As DataRow In dst.Tables("SATMKOP3").Select("")
            'For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
            '    rowSATMKOP3.Item(COLUMN_NAME) = rowSATMKOP3.Item(COLUMN_NAME)
            'Next


            '           For Each rowARTOPEN1 As DataRow In dst.Tables("ARTOPEN1").Select("", "INV_DATE,INV_TYPE,INV_NUM")
            Dim ITEM_CODE As String = rowSATMKOP3.Item("ITEM_CODE")
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE, True)

            'PO DATA


            Dim PO_TEXT As String = ""
            Dim PO_DATE_TEXT As String = ""
            Dim PO_QTY_TEXT As String = ""


            For Each rowPOTORDR2 As DataRow In dst.Tables("POTORDR2").Select("ITEM_CODE = '" & ITEM_CODE & "'", "PO_ORDER_NO ASC")
                If PO_TEXT <> "" Then
                    PO_TEXT = PO_TEXT & vbLf & ""
                    PO_QTY_TEXT = PO_QTY_TEXT & vbLf & ""
                    PO_DATE_TEXT = PO_DATE_TEXT & vbLf & ""
                End If
                PO_TEXT = PO_TEXT & rowPOTORDR2.Item(1) ' & vbLf & ""
                PO_QTY_TEXT = PO_QTY_TEXT & rowPOTORDR2.Item(2) ' & vbLf & ""
                PO_DATE_TEXT = PO_DATE_TEXT & rowPOTORDR2.Item(3) ' & vbLf & ""
            Next

            If PO_TEXT <> "" Then
                    ' WRITE OUT PO 
                Else
                End If

                ' ICTPINV1

                Dim INV_TEXT As String = ""
                Dim INV_DATE_TEXT As String = ""
                Dim INV_QTY_TEXT As String = ""
                Dim INV_ETA_TEXT As String = ""
                Dim TOT_INV As Double = 0



            For Each rowICTPINV2 As DataRow In dst.Tables("ICTPINV2").Select("ITEM_CODE = '" & ITEM_CODE & "'", "INV_NUM ASC")
                If INV_TEXT <> "" Then
                    INV_TEXT = INV_TEXT & vbLf & ""
                    INV_DATE_TEXT = INV_DATE_TEXT & vbLf & ""
                    INV_QTY_TEXT = INV_QTY_TEXT & vbLf & ""
                    INV_ETA_TEXT = INV_ETA_TEXT & vbLf & ""
                End If

                INV_TEXT = INV_TEXT & rowICTPINV2.Item("INV_NUM") ' & vbLf & ""
                INV_DATE_TEXT = INV_DATE_TEXT & rowICTPINV2.Item("INV_DATE") '& vbLf & ""
                INV_QTY_TEXT = INV_QTY_TEXT & rowICTPINV2.Item("PINV_QTY") '& vbLf & ""
                INV_ETA_TEXT = INV_ETA_TEXT & rowICTPINV2.Item("ETA_DATE") '& vbLf & ""
                TOT_INV = TOT_INV + Val(rowICTPINV2.Item("PINV_QTY") & "")
            Next

            If INV_TEXT <> "" Then

                Else
                End If
                ' On Hand
                Dim WHSE_QTY_NOT_INV As String = ""
                Dim WHSE_QTY_ON_HAND As String = ""
                Dim WHSE_QTY_ATS As String = ""
                Dim AMT_NOT_INV As Double = 0


                ASCMAIN1.sql = " ITEM_CODE = '" & ITEM_CODE & "'"
                Dim rows() As DataRow = dst.Tables("ICTSTATX").Select(ASCMAIN1.sql)
                If rows.Length = 0 Then
                    EMsg &= vbCr & "No Record of (Apply-To) Invoice "
                End If


                For Each rowICTSTATX As DataRow In dst.Tables("ICTSTATX").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    AMT_NOT_INV = Val(rowICTSTATX.Item("WHSE_QTY_ONPO") & "") - TOT_INV
                    WHSE_QTY_NOT_INV = AMT_NOT_INV
                    WHSE_QTY_ON_HAND = rowICTSTATX.Item("WHSE_QTY_ON_HAND") & ""
                    WHSE_QTY_ATS = rowICTSTATX.Item("WHSE_QTY_ATS") & ""
                Next


                'Allocation DATA


                Dim ALL_TEXT As String = ""
                Dim ALL_DATE_TEXT As String = ""
                Dim ALL_QTY_TEXT As String = ""

                For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    If ALL_TEXT <> "" Then
                        ALL_TEXT = ALL_TEXT & vbLf & ""
                        ALL_QTY_TEXT = ALL_QTY_TEXT & vbLf & ""
                        ALL_DATE_TEXT = ALL_DATE_TEXT & vbLf & ""
                    End If
                    ALL_TEXT = ALL_TEXT & rowSOTALLO1.Item(1) ' & vbLf & ""
                    ALL_QTY_TEXT = ALL_QTY_TEXT & rowSOTALLO1.Item(2) ' & vbLf & ""
                    ALL_DATE_TEXT = ALL_DATE_TEXT & Format(rowSOTALLO1.Item(3), "MM/dd/yyyyy") & " - " & Format(rowSOTALLO1.Item(4), "MM/dd/yyyyy")
                Next

                Dim PROMO_DATES As String = ""
                Dim PROMO_CUST As String = ""
                Dim PROMO_BOOKING_NAME As String = ""

                For Each rowSPTCOOP1 As DataRow In dst.Tables("SPTCOOP1").Select("ITEM_CODE = '" & ITEM_CODE & "'")
                    If PROMO_DATES <> "" Then
                        PROMO_DATES = PROMO_DATES & vbLf & ""
                        PROMO_CUST = PROMO_CUST & vbLf & ""
                        PROMO_BOOKING_NAME = PROMO_BOOKING_NAME & vbLf & ""
                    End If
                    PROMO_CUST = PROMO_CUST & rowSPTCOOP1.Item("CUST_CODE") ' & vbLf & ""
                    PROMO_DATES = PROMO_DATES & Format(rowSPTCOOP1.Item("DATE_START"), "MM/dd/yyyyy") & " - " & Format(rowSPTCOOP1.Item("DATE_END"), "MM/dd/yyyyy")
                    PROMO_BOOKING_NAME = PROMO_BOOKING_NAME & rowSPTCOOP1.Item("BOOKING_NAME") ' & vbLf & ""

                Next



            rowSATMKOP3.Item("DESC1") = rowICTITEM1.Item("ITEM_DESC") & ""
            rowSATMKOP3.Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE") & ""
                rowSATMKOP3.Item("COST_CATGY_CODE") = rowICTITEM1.Item("COST_CATGY_CODE") & ""
                rowSATMKOP3.Item("PO_NOS") = PO_TEXT
                rowSATMKOP3.Item("PO_QTYS") = PO_QTY_TEXT
                rowSATMKOP3.Item("PO_DATES") = PO_DATE_TEXT
                rowSATMKOP3.Item("INV_NOS") = INV_TEXT
                rowSATMKOP3.Item("INV_QTYS") = INV_QTY_TEXT
                rowSATMKOP3.Item("INV_DATES") = INV_DATE_TEXT
                rowSATMKOP3.Item("INV_ETADATES") = INV_ETA_TEXT

                rowSATMKOP3.Item("NOT_INV") = WHSE_QTY_NOT_INV
                rowSATMKOP3.Item("ON_HAND") = WHSE_QTY_ON_HAND
                rowSATMKOP3.Item("AVL2SELL") = WHSE_QTY_ATS

                rowSATMKOP3.Item("ALL_QTYS") = ALL_QTY_TEXT
                rowSATMKOP3.Item("ALL_DATES") = ALL_DATE_TEXT

                rowSATMKOP3.Item("PROMO_CUST") = PROMO_CUST
                rowSATMKOP3.Item("PROMO_DATES") = PROMO_DATES
                rowSATMKOP3.Item("PROMO_BOOKING_NAME") = PROMO_BOOKING_NAME




                ' WRITE UPDATED 

            Next

    End Sub

    Private Sub grdSATMKOP2_BeforeExitEditMode(sender As Object, e As BeforeExitEditModeEventArgs) Handles grdSATMKOP2.BeforeExitEditMode
        If e.CancellingEditOperation Then Exit Sub

        With grdSATMKOP2.ActiveCell
            Select Case .Column.Key

                Case "ITEM_CODE"

                    'If .Text <> "" Then
                    '    .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    '    cdr = LookUp("ICTITEM1", .Text)
                    '    If cdr Is Nothing Then
                    '        ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                    '        .Value = ""
                    '        e.Cancel = True
                    '    Else
                    '        Dim ITEM_CODE As String = .Text
                    '        grdPOTORDR2.ActiveRow.Cells("BM_ISSUE_NO").Column.ValueList = Get_BMs(ITEM_CODE)
                    '    End If
                    'End If

                Case "COLLECTION_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTCOLL1", .Value)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With

    End Sub

    Private Sub grdSATMKOP2_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdSATMKOP2.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdSATMKOP2.ActiveCell.Column.Key
            Case "ITEM_CODE"

            Case "COLLECTION_CODE"

            Case "CANCEL"
        End Select

        grdClickCellButton(grdSATMKOP2, sql_where, False)
    End Sub

    Private Sub grdSATMKOP2_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSATMKOP2.BeforeRowUpdate

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("TRACKER_NO").Value = Absx1.txtFor("TRACKER_NO").Text
            'Dim PO_ORDER_LNO As Int32 = Val(dst.Tables("POTORDR2").Compute("MAX(PO_ORDER_LNO)", "") & "") + 1
            'e.Row.Cells("PO_ORDER_LNO").Value = PO_ORDER_LNO
            'e.Row.Cells("PO_STATUS").Value = "O"
        End If
    End Sub

    Sub SETUP_ITEMS()
        Dim GROUP_ID As String = grdSATMKOP2.ActiveRow.Cells("GROUP_ID").Value & ""

        '       Fill_Records("ICTPHYJM", LOCATION_CODE)


        grdSATMKOP3.Text = "Items for Group ID" & " " & grdSATMKOP2.ActiveRow.Cells("GROUP_ID").Value


        Dim dvw As DataView = DirectCast(grdSATMKOP3.DataSource, DataTable).DefaultView
        dvw.RowFilter = "GROUP_ID = '" & GROUP_ID & "'"

        Dim GROUP_TYPE As String = grdSATMKOP2.ActiveRow.Cells("GROUP_TYPE").Value & ""
        Dim BRAND_CODE As String = grdSATMKOP2.ActiveRow.Cells("BRAND_CODE").Value & ""
        Dim HC_CODE As String = grdSATMKOP2.ActiveRow.Cells("HC_CODE").Value & ""
        Dim COLLECTION_CODE As String = grdSATMKOP2.ActiveRow.Cells("COLLECTION_CODE").Value & ""

        If GROUP_TYPE = "B" Then
            ASCMAIN1.sql = sqlICTITEM1 _
            & "   and ICTCOLL1.BRAND_CODE = '" & BRAND_CODE & "'"
            grdICTITEM1.Text = "Items for Brand Code" & " " & BRAND_CODE & "'"
            grdSATMKOP2.ActiveRow.Cells("HC_CODE").Value = ""
            grdSATMKOP2.ActiveRow.Cells("COLLECTION_CODE").Value = ""


        ElseIf GROUP_TYPE = "H" Then
            ASCMAIN1.sql = sqlICTITEM1 _
             & "   and ICTCOLL1.HC_CODE = '" & HC_CODE & "'"
            grdICTITEM1.Text = "Items for High Collection" & " " & HC_CODE & "'"

        ElseIf GROUP_TYPE = "C" Then
            ASCMAIN1.sql = sqlICTITEM1 _
             & "   and ICTCOLL1.COLLECTION_CODE = '" & COLLECTION_CODE & "'"
            grdICTITEM1.Text = "Items for Collection Code" & " " & COLLECTION_CODE & "'"

        End If

        If grdSATMKOP2.ActiveRow.Cells(1).Value & "" <> "" And GROUP_TYPE <> "" Then
            Fill_Records("ICTITEM1", "", True, ASCMAIN1.sql)
        End If
        '    grdICTITEM1.Text = "Active Items in Brands " & Join(BRAND_CODEs.ToArray, ",")
        Sort_grdColumns(grdICTITEM1, "ITEM_CODE")

    End Sub

    Private Sub grdSATMKOP2_DoubleClick(sender As Object, e As EventArgs) Handles grdSATMKOP2.DoubleClick
        Dim GROUP_TYPE As String = grdSATMKOP2.ActiveRow.Cells("GROUP_TYPE").Value & ""
        If GROUP_TYPE <> "" Then
            SETUP_ITEMS()
        End If

    End Sub

    Private Sub grdSATMKOP2_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSATMKOP2.AfterRowActivate
        SETUP_ITEMS()
    End Sub
    Sub CREATE_NEW_TRACKING()

        BeginTrans()

        TRACKER_NO = ASCMAIN1.Next_Control_No("SATMKOP1.TRACKER_NO")


        rowSATMKOP1.AcceptChanges()
        rowSATMKOP1.SetAdded()

        For Each rowSATMKOP1 As DataRow In dst.Tables("SATMKOP1").Select("")
            rowSATMKOP1.Item("TRACKER_NO") = TRACKER_NO
            rowSATMKOP1("INIT_OPER") = ASCMAIN1.USER_ID
            rowSATMKOP1("INIT_DATE") = DATETIME_STAMP
            rowSATMKOP1("LAST_OPER") = ASCMAIN1.USER_ID
            rowSATMKOP1("LAST_DATE") = DATETIME_STAMP
        Next

        For Each rowSATMKOP2 As DataRow In dst.Tables("SATMKOP2").Select("")
            rowSATMKOP2.Item("TRACKER_NO") = TRACKER_NO
            rowSATMKOP2.AcceptChanges()
            rowSATMKOP2.SetAdded()

        Next

        For Each rowSATMKOP3 As DataRow In dst.Tables("SATMKOP3").Select("")
            rowSATMKOP3.Item("TRACKER_NO") = TRACKER_NO
            rowSATMKOP3.AcceptChanges()
            rowSATMKOP3.SetAdded()
        Next

        Update_Record_TDA("SATMKOP1")
        Update_Record_TDA("SATMKOP2")
        Update_Record_TDA("SATMKOP3")

        Update_Record_TDA("ASTAUDT1")

        CommitTrans("Tracker No " & TRACKER_NO & " has been created")

        Absx1.txtFor("TRACKER_NO").Value = TRACKER_NO

    End Sub

    Private Sub UltraTextEditor4_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor4.ValueChanged

    End Sub
End Class