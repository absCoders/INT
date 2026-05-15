Imports Infragistics.Win.UltraWinGrid

Public Class APFACRB1

    Dim APTACRX1 As String = String.Empty
    Dim sqlAPTACRX1 As String
    Dim APTACRCX As String = String.Empty
    Dim sqlAPTACRCX As String

    Dim CTL_NO_ADJs As New Dictionary(Of String, Decimal)
    Dim CTL_NO_BOLs As New Dictionary(Of String, String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "APFACRBI" Then
            InquiryMode = True
        End If

        If Not InquiryMode Then
            AUDIT.Add("APTACRC1", "E")
        End If

        tabMatch.Tabs(1).Visible = False

        Get_PARM("APTPARM1")

        With dst

            sqlAPTACRX1 = "Select APTACRC1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
                & " from APTACRC1, ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE (+) = APTACRC1.ITEM_CODE" & vbCrLf
            ASCMAIN1.sql = sqlAPTACRX1 & " and ROWNUM < 1"
            APTACRX1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add Primary Key (CTL_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add DEF_TOLERANCE NUMBER(6,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add COST_ACC_TOTAL NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add TPV_ADJ_TOTAL NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRX1 & " Add LNO_COUNT NUMBER(5,0)")


            sqlAPTACRCX = "Select APTACRC0.CTL_NO_MATCHED" & vbCrLf _
                & " from APTACRC0"
            ASCMAIN1.sql = sqlAPTACRCX & " where ROWNUM < 1"
            APTACRCX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add Primary Key (CTL_NO_MATCHED)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add DEF_TOLERANCE NUMBER(6,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add COST_ACT_TOTAL NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add COST_ACC_TOTAL NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add TPV_ADJ_TOTAL NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add CALC_VARIANCE NUMBER(13,2)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add LNO_COUNT NUMBER(5,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & APTACRCX & " Add SOURCE_DOC_NO VARCHAR2(20)")


            ASCMAIN1.sql = $"Select APTACRC0.CTL_NO_MATCHED " & vbCrLf &
                $", {APTACRCX}.DEF_TOLERANCE" & vbCrLf &
                $", {APTACRCX}.LNO_COUNT" & vbCrLf &
                $", {APTACRCX}.COST_ACT_TOTAL" & vbCrLf &
                $", {APTACRCX}.COST_ACC_TOTAL" & vbCrLf &
                $", CASE WHEN NVL({APTACRCX}.LNO_COUNT, 0) > 0 THEN NVL({APTACRCX}.COST_ACT, 0) - NVL({APTACRCX}.COST_ACC_TOTAL, 0) ELSE NULL END as CALC_VARIANCE" & vbCrLf &
                $", {APTACRCX}.TPV_ADJ_TOTAL" & vbCrLf &
                $", APTACRC0.SOURCE_DOC_NO" & vbCrLf &
                $" from APTACRC0, {APTACRCX} " & vbCrLf &
                $" where APTACRC0.ACCRUAL_CODE ='TRF'" & vbCrLf &
                $" And {APTACRCX}.CTL_NO_MATCHED = APTACRC0.CTL_NO_MATCHED"
            sqlAPTACRCX = ASCMAIN1.sql

            Create_TDA(.Tables.Add, "APTACRC0", "*", 0, True)
            With .Tables("APTACRC0")
                .Columns("VAR_OK").DefaultValue = "0"
                .Columns.Add("PREV_PPD", GetType(System.Decimal))
                .Columns.Add("PREV_ACC", GetType(System.Decimal))
                .Columns.Add("PREV_VAR", GetType(System.Decimal))
            End With

            ASCMAIN1.sql = "Select * from APTACRC0"
            Create_TDA(.Tables.Add, "APTACRCR", "**", 0, False)

            Create_Relation("APTACRC0", "APTACRCR", "CTL_NO_MATCHED", "CTL_NO_MATCHED_NEXT")

            With .Tables("APTACRC0")
                .Columns("PREV_PPD").Expression = "SUM(CHILD.COST_ACT_TOTAL)"
                .Columns("PREV_ACC").Expression = "SUM(CHILD.COST_ACC_TOTAL)"
                .Columns("PREV_VAR").Expression = "SUM(CHILD.CALC_VARIANCE)"
            End With

            ASCMAIN1.sql = $"Select APTACRC1.* " & vbCrLf &
                $", {APTACRX1}.DEF_TOLERANCE" & vbCrLf &
                $", {APTACRX1}.LNO_COUNT" & vbCrLf &
                $", CASE WHEN NVL({APTACRX1}.LNO_COUNT, 0) > 0 THEN NVL({APTACRX1}.COST_ACT, 0) - NVL({APTACRX1}.COST_ACC_TOTAL, 0) ELSE NULL END as CALC_VARIANCE" & vbCrLf &
                $", {APTACRX1}.TPV_ADJ_TOTAL" & vbCrLf &
                $" from APTACRC1, {APTACRX1} " & vbCrLf &
                $" where APTACRC1.ACCRUAL_CODE ='TRF' AND APTACRC1.CTL_STATUS = '1' AND NVL(APTACRC1.PPD_IND, '0') = '1' AND NVL(APTACRC1.PPD_MATCHED, '0') = :PARM1" & vbCrLf &
                $" And ({APTACRX1}.BOL_NO = APTACRC1.SOURCE_DOC_NO or {APTACRX1}.BOL_NO like '%' || APTACRC1.SOURCE_DOC_NO ) And {APTACRX1}.PPD_IND = '1'"

            'Create_TDA(.Tables.Add, "APTACRC1", "**", 0, True, "V", 1)

            Create_TDA(.Tables.Add, "APTACRCM", "**", 0, False, "V", 1)


            ASCMAIN1.sql = "Select APTACRX1.*" & vbCrLf _
                & ", (NVL(ICTIREC2.PO_COST,0) - NVL(ICTCOSTA.ITEM_COST_VCOST,0)) * NVL(ICTIREC2.QTY_REC,0) TRAN_PV_REC" & vbCrLf _
                & ", (CASE WHEN NVL(ICTIREC2.ACCRUAL_STATUS,'0') = '1' THEN NVL(ICTIREC2.AMT_INV,0) - NVL(ICTIREC2.PO_COST,0) * NVL(ICTIREC2.QTY_INV,0) ELSE 0 END) TRAN_PV_INV" & vbCrLf _
                & ", ICTIREC2.QTY_REC, ICTIREC2.PO_COST" & vbCrLf _
                & $"  from {APTACRX1} APTACRX1, ICTIREC2, ICTCOSTA" & vbCrLf _
                & " where APTACRX1.BOL_NO = :PARM1" & vbCrLf _
                & "   and ICTIREC2.RECEIPT_NO = APTACRX1.RECEIPT_NO" & vbCrLf _
                & "   and ICTIREC2.RECEIPT_LNO = APTACRX1.RECEIPT_LNO" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP (+) = APTACRX1.OPS_YYYYPP" & vbCrLf _
                & "   and ICTCOSTA.ITEM_CODE (+) = APTACRX1.ITEM_CODE"

            ' Create_TDA(.Tables.Add("APTACRX1"), APTACRX1, "**", 0, True, "V", 1)

            Create_TDA(.Tables.Add, "APTACRCT", "**", 0, False, "V", 1)

            ASCMAIN1.sql = Replace(ASCMAIN1.sql, "APTACRX1.BOL_NO = :PARM1", "APTACRX1.CTL_NO_MATCHED is Null")
            Create_TDA(.Tables.Add, "APTACRCO", "**", 0, False, "", 1)

            ASCMAIN1.sql = "SELECT * from ASTAUDT1 where TABLE_NAME='APTACRC1' AND KEY_VALUE=:PARM1"
            Create_TDA(.Tables.Add, "ASTAUDTX", "**", 0, False, "V")


            With dst.Tables.Add("APTACRXA")
                .Columns.Add("CTL_NO_MATCHED", GetType(System.String))
                .Columns.Add("COST_CATGY_CODE", GetType(System.String))
                .Columns.Add("SOURCE_DOC_NO", GetType(System.String))
                '.Columns.Add("CTL_NO_MATCHED", GetType(System.String))
                .Columns.Add("COST_ACC", GetType(System.Decimal))
                .Columns.Add("COST_VAR", GetType(System.Decimal))
                .Columns("COST_VAR").DefaultValue = 0
                .PrimaryKey = New DataColumn() { .Columns("CTL_NO_MATCHED"), .Columns("COST_CATGY_CODE")}
            End With
        End With

        grdAPTACRC0.DataSource = dst.Tables("APTACRC0")
        With grdAPTACRC0.DisplayLayout
            .Override.ExpansionIndicator = ShowExpansionIndicator.CheckOnDisplay

            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            If InquiryMode Then
                .Override.AllowUpdate = DefaultableBoolean.False
            Else
                .Override.AllowUpdate = DefaultableBoolean.True
            End If
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide

            With .Bands(1)
                .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                .Override.AllowDelete = DefaultableBoolean.False
                .Override.AllowUpdate = DefaultableBoolean.False
            End With
        End With
        With grdAPTACRC0.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If New String() {"SOURCE_DOC_NO", "VAR_OK", "NOTES"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf New String() {"PREV_PPD", "PREV_ACC", "PREV_VAR"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.PaleVioletRed
                End If
            Next
        End With
        'Create_Summary(grdAPTACRC0, "CTL_NO_MATCHED", "Count")
        Create_Summary(grdAPTACRC0, "SOURCE_DOC_NO", "Count")

        grdAPTACRCM.DataSource = dst.Tables("APTACRCM")
        With grdAPTACRCM.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            If InquiryMode Then
                .Override.AllowUpdate = DefaultableBoolean.False
            Else
                .Override.AllowUpdate = DefaultableBoolean.True
            End If
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
        End With
        With grdAPTACRCM.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If New String() {"SOURCE_DOC_NO"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With
        Create_Summary(grdAPTACRCM, "CTL_NO", "Count")
        Create_Summary(grdAPTACRCM, "COST_ACT")



        grdAPTACRCT.DataSource = dst.Tables("APTACRCT")
        With grdAPTACRCT.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            If InquiryMode Then
                .Override.AllowUpdate = DefaultableBoolean.False
            Else
                .Override.AllowUpdate = DefaultableBoolean.True
            End If
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
        End With
        With grdAPTACRCT.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If New String() {"TPV_ADJ"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                End If
            Next
        End With
        Create_Summary(grdAPTACRCT, "CTL_NO", "Count")
        Create_Summary(grdAPTACRCT, New String() {"COST_ACC", "TPV_ADJ", "TRAN_PV_REC", "TRAN_PV_INV", "COST_VAR_ITEM"})


        grdAPTACRCO.DataSource = dst.Tables("APTACRCO")
        With grdAPTACRCO.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            If InquiryMode Then
                .Override.AllowUpdate = DefaultableBoolean.False
            Else
                .Override.AllowUpdate = DefaultableBoolean.True
            End If
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
        End With
        With grdAPTACRCO.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                'If New String() {"TPV_ADJ"}.Contains(gcol.Key) Then
                '    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                '    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                'End If
            Next
        End With
        Create_Summary(grdAPTACRCO, "CTL_NO", "Count")
        Create_Summary(grdAPTACRCO, New String() {"COST_ACC", "TPV_ADJ", "TRAN_PV_REC", "TRAN_PV_INV"})


        'grdAPTACRC1.DataSource = dst.Tables("APTACRC1")
        'With grdAPTACRC1.DisplayLayout
        '    .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .Override.AllowDelete = DefaultableBoolean.False
        '    If InquiryMode Then
        '        .Override.AllowUpdate = DefaultableBoolean.False
        '    Else
        '        .Override.AllowUpdate = DefaultableBoolean.True
        '    End If
        '    .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
        'End With
        'With grdAPTACRC1.DisplayLayout.Bands(0)
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
        '        gcol.Header.Appearance.BackColor = Drawing.Color.White
        '        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        '        If New String() {"SOURCE_DOC_NO", "VAR_OK", "NOTES"}.Contains(gcol.Key) Then
        '            gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
        '            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
        '        End If
        '    Next
        'End With
        'Create_Summary(grdAPTACRC1, "CTL_NO", "Count")
        'Create_Summary(grdAPTACRC1, "TPV_ADJ_TOTAL")


        'grdAPTACRX1.DataSource = dst.Tables("APTACRX1")
        'With grdAPTACRX1.DisplayLayout
        '    .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
        '    .Override.AllowDelete = DefaultableBoolean.False
        '    If InquiryMode Then
        '        .Override.AllowUpdate = DefaultableBoolean.False
        '    Else
        '        .Override.AllowUpdate = DefaultableBoolean.True
        '    End If
        '    .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
        'End With
        'With grdAPTACRX1.DisplayLayout.Bands(0)
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
        '        gcol.Header.Appearance.BackColor = Drawing.Color.White
        '        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        '        If New String() {"TPV_ADJ"}.Contains(gcol.Key) Then
        '            gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
        '            gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
        '        End If
        '    Next
        'End With
        'Create_Summary(grdAPTACRX1, "CTL_NO", "Count")
        'Create_Summary(grdAPTACRX1, New String() {"COST_ACC", "TPV_ADJ", "TRAN_PV_REC", "TRAN_PV_INV"})


        grdASTAUDTX.DataSource = dst.Tables("ASTAUDTX")
        With grdASTAUDTX.DisplayLayout
            .Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            .Override.AllowDelete = DefaultableBoolean.False
            .Override.AllowUpdate = DefaultableBoolean.False
            .NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Hide
        End With
        With grdASTAUDTX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            Next
        End With

        MakeTransparent(chkShowReToggle)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey
            Case "Load"
                If Not (InquiryMode) Then
                    If optStatus.Value = "M" Then
                    Else
                        ASCMAIN1.sql = "Select Distinct(VEND_CODE_ACC) from APTACRC1 WHERE ACCRUAL_CODE='TRF' AND CTL_STATUS= '1' AND NVL(PPD_IND, '0') = '1' AND NVL(PPD_MATCHED, '0') = 0"
                        Dim tblVEND_CODE_ACCs As DataTable = ASCDATA1.GetDataTable()
                        For Each rowVEND_CODE_ACCs As DataRow In tblVEND_CODE_ACCs.Select("")
                            If Not ASCMAIN1.Logical_Lock("APTVEND1", rowVEND_CODE_ACCs("VEND_CODE_ACC")) Then
                                Exit Sub
                            End If
                        Next
                    End If
                End If

            Case "Update"

                'Stop ' EMsg &= "Variance entries require a Note." & vbCrLf

                If dst.Tables("APTACRC0").Select("VAR_OK = '1' and ISNULL(NOTES,'') = ''").Length > 0 Then
                    EMsg &= vbCr & "OK to Variances must be accompanied by a Note"
                End If

                For Each row As DataRow In dst.Tables("APTACRC0").Select("ISNULL(NOTES,'') = ''", "", DataRowState.Modified)
                    If row.Item("SOURCE_DOC_NO", DataRowVersion.Current) & "" <> row.Item("SOURCE_DOC_NO", DataRowVersion.Original) & "" Then
                        EMsg &= vbCr & "All changes must be accompanied by a Note"
                        Exit For
                    End If
                Next

                For Each row As DataRow In dst.Tables("APTACRC0").Select("VAR_OK = '1' AND ISNULL(CALC_VARIANCE,0) >  ABS(ISNULL(VAR_TOLERANCE,0))")
                    Dim SOURCE_DOC_NO As String = row.Item("SOURCE_DOC_NO") & ""
                    EMsg &= vbCr & $"Variance has Changed since OK'd - please toggle Var OK to ack BOL {SOURCE_DOC_NO}"
                Next

                'Stop

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Load_Record()
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

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("Load").Settings.Enabled = not_iScreenMode
                If Not (InquiryMode) Then
                    .Items("Update").Settings.Enabled = iScreenMode
                Else
                    .Items("Update").Settings.Enabled = DefaultableBoolean.False
                End If
                .Items("Cancel").Settings.Enabled = iScreenMode

                .Items("Update").Visible = (tf And optStatus.Value = "U") And Not InquiryMode
                .Items("Cancel").Visible = (tf And optStatus.Value = "U") And Not InquiryMode
                .Items("Done").Visible = (tf And (optStatus.Value = "M" Or InquiryMode))
            End With

        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Set_Read_Only(optStatus, ScreenMode)

        grdAPTACRC1.Visible = ScreenMode
        tabDetails.Visible = ScreenMode

        tabMatch.Visible = ScreenMode

        If ScreenMode Then
            If Not InquiryMode Then
                If optStatus.Value = "M" Then
                    grdAPTACRC1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdAPTACRX1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdAPTACRC0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdAPTACRCT.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                    grdAPTACRCM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                Else
                    grdAPTACRC1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdAPTACRX1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdAPTACRC0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdAPTACRCT.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                    grdAPTACRCM.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                End If

                grdAPTACRC0.Text = $"{optStatus.Text} Prepaid Tariff Records"
                grdAPTACRC1.Text = $"{optStatus.Text} Prepaid Tariff Records"
            End If


            For Each COL As String In New String() {"OPS_YYYYPP_MATCHED", "INIT_DATE", "INIT_OPER", "PPD_MATCHED_XNO"} ' , "CTL_NO_MATCHED"
                grdAPTACRC0.DisplayLayout.Bands(0).Columns(COL).Hidden = Not (optStatus.Value = "M")
                grdAPTACRC0.DisplayLayout.Bands(1).Columns(COL).Hidden = Not (optStatus.Value = "M")
            Next

            tabAccruals.Tabs("Orphans").Visible = Not (optStatus.Value = "M")
            chkShowReToggle.Visible = False ' Not (optStatus.Value = "M")
        Else
            Clear_Record()

            grdAPTACRC0.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdAPTACRCM.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdAPTACRCT.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdAPTACRCO.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

            chkShowReToggle.Checked = False
            chkShowReToggle.Visible = False
        End If

    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Unmatched Prepaid Tariff Records")

        CTL_NO_ADJs.Clear()
        CTL_NO_BOLs.Clear()

        Dim sqlw As String = $"APTACRC1.ACCRUAL_CODE = 'TRF' and "
        If optStatus.Value = "M" Then
            'sqlw &= $"(APTACRC1.CTL_STATUS = '1' and APTACRC1.PPD_IND = '1' and NVL(APTACRC1.PPD_MATCHED,'0') = '1')"
            sqlw &= $"(APTACRC1.CTL_STATUS = '1' and APTACRC1.PPD_MATCHED_XNO is Not Null)"
        Else
            sqlw &= $"(APTACRC1.CTL_STATUS = '0' or (APTACRC1.CTL_STATUS = '1' and APTACRC1.PPD_IND = '1' and NVL(APTACRC1.PPD_MATCHED,'0') = '0'))"
        End If



        ' Create_TDA(dst.Tables.Add, "GLTINTF1", "*") Not sure if I need this
        ASCDATA1.ExecuteSQL($"Truncate Table {APTACRX1}")
        ' ASCDATA1.ExecuteSQL("Insert into " & APTACRX1 & " " & Replace(sqlAPTACRX1, " from ", ", NULL BOL_NO, NULL DEF_TOLERANCE, NULL COST_ACC_TOTAL, NULL TPV_ADJ_TOTAL, NULL LNO_COUNT from ") & "   and " & sqlw)
        ASCDATA1.ExecuteSQL("Insert into " & APTACRX1 & " " & Replace(sqlAPTACRX1, " from ", ", NULL DEF_TOLERANCE, NULL COST_ACC_TOTAL, NULL TPV_ADJ_TOTAL, NULL LNO_COUNT from ") & "   and " & sqlw)

        Dim COL As String = "SOURCE_DOC_NO"
        If optStatus.Value = "M" Then
            COL = "CTL_NO_MATCHED"
        Else
            ASCMAIN1.sql = $"UPDATE {APTACRX1} SET BOL_NO_MATCHED = SOURCE_DOC_NO WHERE PPD_IND = '1'"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Begin Declare XNO Number (3,0); CTL_NO_MATCHED_XNO VARCHAR2(10);" & vbCrLf _
                & $" Cursor C1 Is Select Distinct SOURCE_DOC_NO from {APTACRX1} where PPD_IND = '1';" & vbCrLf _
                & " Begin For R1 in C1 Loop" & vbCrLf _
                & "  XNO := NVL(XNO,0) + 1;" & vbCrLf _
                & "  CTL_NO_MATCHED_XNO := 'M' || TRIM(TO_CHAR(XNO,'000000000'));" & vbCrLf _
                & $"  Update {APTACRX1} Set CTL_NO_MATCHED = CTL_NO_MATCHED_XNO" & vbCrLf _
                & $"  where PPD_IND = '1'" & vbCrLf _
                & "     and SOURCE_DOC_NO = R1.SOURCE_DOC_NO;" & vbCrLf _
                & $"  Update {APTACRX1} Set CTL_NO_MATCHED = CTL_NO_MATCHED_XNO, BOL_NO_MATCHED = R1.SOURCE_DOC_NO" & vbCrLf _
                & $"  where PPD_IND = '0'" & vbCrLf _
                & "     and (NVL(BOL_NO,SOURCE_DOC_NO) = R1.SOURCE_DOC_NO OR SOURCE_DOC_NO LIKE '%' || R1.SOURCE_DOC_NO);" & vbCrLf _
                & " End Loop; End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            COL = "BOL_NO_MATCHED"
        End If

        ASCMAIN1.sql = $"Select {COL} SOURCE_DOC_NO" & vbCrLf _
            & ", COUNT (*) RECS" & vbCrLf _
            & ", SUM (CASE WHEN NVL(PPD_IND,'0') = '1' THEN 1 ELSE 0 END) PPDS" & vbCrLf _
            & $" from {APTACRX1} APTACRX1" & vbCrLf _
            & $"GROUP BY {COL}" & vbCrLf _
            & "HAVING COUNT (*) > 1"
        ASCMAIN1.sql = $"Select * from ({ASCMAIN1.sql}) where PPDS >= 1"

        Dim w As String = "(SOURCE_DOC_NO = R1.SOURCE_DOC_NO OR SOURCE_DOC_NO LIKE '%' || R1.SOURCE_DOC_NO)"
        If optStatus.Value = "M" Then
            w = "CTL_NO_MATCHED = R1.SOURCE_DOC_NO"
        End If

        If optStatus.Value = "M" Then
        Else
            Fill_Records("APTACRCO")
            ASCMAIN1.sql = $"Delete from {APTACRX1} where CTL_NO_MATCHED is Null"
            grdAPTACRCO.Text = "Orphan Un-Matched Accruals"
            Sort_grdColumns(grdAPTACRCO, "SOURCE_DOC_NO")
        End If

        ' (SOURCE_DOC_NO = R1.SOURCE_DOC_NO OR SOURCE_DOC_NO LIKE '%' || R1.SOURCE_DOC_NO)
        ' CTL_NO_MATCHED = R1.SOURCE_DOC_NO
        If w = "CTL_NO_MATCHED = R1.SOURCE_DOC_NO" Then
            w = "CTL_NO_MATCHED = R1.CTL_NO_MATCHED"
        End If
        w = "CTL_NO_MATCHED = R1.CTL_NO_MATCHED" ' NOW THAT WE ARE USING CTL_NO_MATCHED FOR Matched as well as Un-Matched

        ASCMAIN1.sql = $"UPDATE {APTACRX1} R1 SET DEF_TOLERANCE = " & vbCrLf _
            & $"(SELECT COUNT (*) * .01  FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0')" & vbCrLf _
            & " where PPD_IND = '1'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRX1} R1 SET LNO_COUNT = " & vbCrLf _
            & $"(SELECT COUNT (*) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0')" & vbCrLf _
            & " where PPD_IND = '1'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRX1} R1 SET COST_ACC_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (COST_ACC) from {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0'" & vbCrLf _
            & "HAVING COUNT(*) > 0)" & vbCrLf _
            & " where PPD_IND = '1' AND NVL(LNO_COUNT, 0) > 0"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRX1} R1 SET TPV_ADJ_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (TPV_ADJ)  FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0'" & vbCrLf _
            & "HAVING COUNT(*) > 0)" & vbCrLf _
            & " where PPD_IND = '1' AND NVL(LNO_COUNT, 0) > 0"
        ASCDATA1.ExecuteSQL()



        ASCDATA1.ExecuteSQL("Truncate Table " & APTACRCX)
        'ASCDATA1.ExecuteSQL("Insert into " & APTACRCX & " (CTL_NO_MATCHED, SOURCE_DOC_NO) " & "Select CTL_NO_MATCHED, SOURCE_DOC_NO from APTACRC0")
        ASCDATA1.ExecuteSQL($"Insert into {APTACRCX} (CTL_NO_MATCHED, SOURCE_DOC_NO) Select Distinct CTL_NO_MATCHED, SOURCE_DOC_NO from {APTACRX1} where PPD_IND = '1'")

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET DEF_TOLERANCE = " & vbCrLf _
            & $"(SELECT COUNT (*) * .01  FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET LNO_COUNT = " & vbCrLf _
            & $"(SELECT COUNT (*) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET COST_ACC_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (COST_ACC) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET COST_ACT_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (COST_ACT) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '1') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET TPV_ADJ_TOTAL = " & vbCrLf _
            & $"(SELECT SUM (TPV_ADJ) FROM {APTACRX1}" & vbCrLf _
            & $"where {w}" & vbCrLf _
            & " AND NVL(PPD_IND,'0') = '0') "
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"UPDATE {APTACRCX} R1 SET CALC_VARIANCE = " & vbCrLf _
            & $"Case WHEN NVL(LNO_COUNT, 0) > 0 THEN NVL(COST_ACT_TOTAL, 0) - NVL(COST_ACC_TOTAL, 0) ELSE NULL END"
        ASCDATA1.ExecuteSQL()


        Dim PPD_MATCHED As String = IIf(optStatus.Value = "M", "1", "0")

        EnforceConstraints(False)

        If optStatus.Value = "M" Then
            Fill_Records("APTACRC0")
        Else
            Dim sqlT As String = $"Select CTL_NO_MATCHED, ACCRUAL_CODE, SOURCE_DOC_NO" & vbCrLf _
                & $", MAX(NOTES) NOTES, MAX(NVL(VAR_OK,'0')) VAR_OK, MAX(VAR_TOLERANCE) VAR_TOLERANCE" & vbCrLf _
                & $" from {APTACRX1} where ACCRUAL_CODE = 'TRF' and PPD_IND = '1'" & vbCrLf _
                & " group by CTL_NO_MATCHED, ACCRUAL_CODE, SOURCE_DOC_NO"
            ASCMAIN1.sql = $"Select X.CTL_NO_MATCHED, X.ACCRUAL_CODE, X.NOTES, NULL PPD_MATCHED_XNO, X.SOURCE_DOC_NO, NULL OPS_YYYYPP_MATCHED, NULL INIT_DATE, NULL INIT_OPER" & vbCrLf &
                $", {APTACRCX}.DEF_TOLERANCE" & vbCrLf &
                $", {APTACRCX}.LNO_COUNT" & vbCrLf &
                $", {APTACRCX}.COST_ACT_TOTAL" & vbCrLf &
                $", {APTACRCX}.COST_ACC_TOTAL" & vbCrLf &
                $", CASE WHEN NVL({APTACRCX}.LNO_COUNT, 0) > 0 THEN NVL({APTACRCX}.COST_ACT_TOTAL, 0) - NVL({APTACRCX}.COST_ACC_TOTAL, 0) ELSE NULL END as CALC_VARIANCE" & vbCrLf &
                $", {APTACRCX}.TPV_ADJ_TOTAL" & vbCrLf &
                $", X.VAR_OK" & vbCrLf &
                $", X.VAR_TOLERANCE" & vbCrLf &
                $" from ({sqlT}) X, {APTACRCX} " & vbCrLf &
                $" where X.ACCRUAL_CODE ='TRF'" & vbCrLf &
                $" And {APTACRCX}.CTL_NO_MATCHED = X.CTL_NO_MATCHED"

            Fill_Records("APTACRC0", , , ASCMAIN1.sql)
        End If

        'Stop ' 1-TIME UPDATE
        'Update_Record_TDA("APTACRC0", "1=1")
        ' UPDATE APTACRC0 X SET INIT_DATE = sysdate, INIT_OPER = 'wjz', PPD_MATCHED_XNO = (SELECT PPD_MATCHED_XNO FROM WJZ_APTACRC0 WHERE CTL_NO_MATCHED = X.CTL_NO_MATCHED)

        Fill_Records("APTACRCT")
        Fill_Records("APTACRCM")

        Get_PREV_APTACRC0


        EnforceConstraints(True)

        Sort_grdColumns(grdAPTACRC0, "SOURCE_DOC_NO")
        Setup_grdAPTACRC0()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
        {"APTACRC0", "APTACRCR", "APTACRCM", "APTACRCT"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        EnforceConstraints(True)
        'Clear_All_Filters(grdAPTACRC1)
        Clear_All_Filters(grdAPTACRC0)
    End Sub

    Private Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Updating...")

        'For Each CTL_NO As String In CTL_NO_ADJs
        '    Fill_Records("APTACRC1",, False, $"Select * from APTACRC1 where CTL_NO = '{CTL_NO}'")
        '    Dim rowAPTACRC1 As DataRow = dst.Tables("APTACRC1").Rows.Find(CTL_NO)
        '    Dim rowAPTACRX1 As DataRow = dst.Tables("APTACRX1").Rows.Find(CTL_NO)
        '    rowAPTACRC1.Item("TPV_ADJ") = rowAPTACRX1.Item("TPV_ADJ")
        'Next

        MyBase.BeginTrans()

        Dim R As Integer = 0

        For Each rowAPTACRC0 As DataRow In dst.Tables("APTACRC0").Select("", "", DataRowState.Modified)
            Dim CTL_NO_MATCHED As String = rowAPTACRC0.Item("CTL_NO_MATCHED") & ""
            Dim BOL_NO_orig As String = rowAPTACRC0.Item("SOURCE_DOC_NO", DataRowVersion.Original) & ""
            Dim BOL_NO_curr As String = rowAPTACRC0.Item("SOURCE_DOC_NO", DataRowVersion.Current) & ""
            If BOL_NO_curr <> BOL_NO_orig Then
                ' update aptacrc1 for both PPD=1 and PPD=0
                ASCMAIN1.sql = "Update APTACRC1 Set SOURCE_DOC_NO = :PARM1 where PPD_IND = '1' AND NVL(PPD_MATCHED,'0') = '0' and SOURCE_DOC_NO = :PARM2"
                R = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {BOL_NO_curr, BOL_NO_orig})
                'Stop

                ASCMAIN1.sql = "Select APTACRX1.* " _
                    & $" from {APTACRX1} APTACRX1 WHERE NVL(PPD_IND,'0') = '0' and CTL_NO_MATCHED = '{CTL_NO_MATCHED}'"
                Fill_Records("APTACRCT", ,, ASCMAIN1.sql) ' Accruals
                For Each rowAPTACRCT As DataRow In dst.Tables("APTACRCT").Select("")
                    Dim CTL_NO As String = rowAPTACRCT.Item("CTL_NO") & ""
                    'ASCMAIN1.sql = "Update APTACRC1 Set SOURCE_DOC_NO = :PARM1 where NVL(PPD_IND,'0') = '0' AND CTL_STATUS = '0' and CTL_NO = :PARM2"
                    ASCMAIN1.sql = "Update APTACRC1 Set BOL_NO = :PARM1 where NVL(PPD_IND,'0') = '0' AND CTL_STATUS = '0' and CTL_NO = :PARM2"
                    R = ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New String() {BOL_NO_curr, CTL_NO})
                    'Stop
                Next
            End If

            Dim VAR_OK_orig As String = rowAPTACRC0.Item("VAR_OK", DataRowVersion.Original) & ""
            Dim VAR_OK_curr As String = rowAPTACRC0.Item("VAR_OK", DataRowVersion.Current) & ""
            Dim NOTES_orig As String = rowAPTACRC0.Item("NOTES", DataRowVersion.Original) & ""
            Dim NOTES_curr As String = rowAPTACRC0.Item("NOTES", DataRowVersion.Current) & ""
            Dim VAR_TOLERANCE_curr As Decimal = Val(rowAPTACRC0.Item("VAR_TOLERANCE", DataRowVersion.Current) & "")
            Dim VAR_TOLERANCE_orig As Decimal = Val(rowAPTACRC0.Item("VAR_TOLERANCE", DataRowVersion.Original) & "")
            Dim VAR_TOLERANCE As Decimal = VAR_TOLERANCE_curr
            Dim CALC_VARIANCE As Decimal = Val(rowAPTACRC0.Item("CALC_VARIANCE", DataRowVersion.Current) & "")
            If VAR_OK_curr <> VAR_OK_orig Or NOTES_curr <> NOTES_orig Then
                If VAR_OK_curr <> "1" Then
                    VAR_TOLERANCE = 0
                End If
                ASCMAIN1.sql = "Update APTACRC1 Set VAR_OK = :PARM1, NOTES = :PARM2, VAR_TOLERANCE = :PARM3" & vbCrLf _
                    & " where PPD_IND = '1' AND NVL(PPD_MATCHED,'0') = '0' and SOURCE_DOC_NO = :PARM4"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVNV", New Object() {VAR_OK_curr, NOTES_curr, VAR_TOLERANCE, BOL_NO_curr})
            End If
        Next

        For Each CTL_NO As String In CTL_NO_ADJs.Keys
            Dim TPV_ADJ As Decimal = CTL_NO_ADJs(CTL_NO)
            ASCMAIN1.sql = "Update APTACRC1 Set TPV_ADJ = :PARM1 where CTL_NO = :PARM2"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NV", New Object() {TPV_ADJ, CTL_NO})
        Next
        For Each CTL_NO As String In CTL_NO_BOLs.Keys
            Dim BOL_NO As String = CTL_NO_BOLs(CTL_NO)
            ASCMAIN1.sql = "Update APTACRC1 Set BOL_NO = :PARM1 where CTL_NO = :PARM2"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "NV", New Object() {BOL_NO, CTL_NO})
        Next

        'Update_Record_TDA("APTACRC1")

        'Stop
        MyBase.CommitTrans("Update Complete")
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdAPTACRC1, "SSB", "Show Filter", "Show GroupBox", "Attachments")
        'Load_Popup_Menu(grdAPTACRX1, "SSB", "Show Filter", "Show GroupBox", "PO Receipts Inquiry")

        Load_Popup_Menu(grdAPTACRC0, "SSB", "Show Filter", "Show GroupBox", "Attachments")
        Load_Popup_Menu(grdAPTACRCM, "SSB", "Show Filter", "Show GroupBox", "Attachments")
        Load_Popup_Menu(grdAPTACRCT, "SSBB", "Show Filter", "Show GroupBox", "PO Receipts Inquiry", "Detach from BOL")
        Load_Popup_Menu(grdAPTACRCO, "SSB", "Show Filter", "Show GroupBox", "Attach to BOL")

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

        Select Case grd.Name
            '    Case "grdATEVNTX"
            '        tlb_btn = DirectCast(tlb_pop.Tools("Refresh"), UltraWinToolbars.ButtonTool)
            '        tlb_btn.SharedProps.Visible = (EntryMode = "Edit" Or EntryMode = "New")

            Case "grdAPTACRCO"
                tlb_btn = DirectCast(tlb_pop.Tools("Attach to BOL"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not (InquiryMode Or optStatus.Value = "M")

            Case "grdAPTACRCT"
                tlb_btn = DirectCast(tlb_pop.Tools("Detach from BOL"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = Not (InquiryMode Or optStatus.Value = "M")

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If
        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        'Select Case e.Tool.Key
        '    Case "Refresh"
        '        If grd.Name = "grdTATEVNTX" Then
        '            Fill_APTACRC1()
        '        End If
        'End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Attachments"
                If grd.Name = "grdAPTACRC1" Then

                    Dim ENTITY As New Dropped_On_Entity
                    ENTITY.TABLE_NAME = "APTACRC1"
                    ENTITY.COLUMN_NAME = "CTL_NO"
                    ENTITY.CODE_VALUE = grd.ActiveRow.Cells("CTL_NO").Value
                    ENTITY.READ_ONLY = InquiryMode
                    Dim F As New ASFATTA1
                    F.ENTITY = ENTITY
                    F.ShowDialog()
                    F.Dispose()

                    'tvwTATCONV1.ActiveNode.Cells("CONV_ATTACHMENTS").Value = Get_CONV_ATTACHMENTS(tvwTATCONV1.ActiveNode.Cells("CONV_NO").Value)
                End If

            Case "PO Receipts Inquiry"
                Dim RECEIPT_NO As String = grd.ActiveRow.Cells("RECEIPT_NO").Text
                Context_Launch("View", RECEIPT_NO, e.Tool.Key, "ICFIRECI")

            Case "Attach to BOL"
                If grdAPTACRC0.ActiveRow IsNot Nothing AndAlso grdAPTACRC0.ActiveRow.IsDataRow Then

                    If grdAPTACRCO.Selected.Rows.Count = 0 Then
                        If grdAPTACRCO.ActiveRow IsNot Nothing Then
                            grdAPTACRCO.ActiveRow.Selected = True
                        End If
                    End If
                    If grdAPTACRCO.Selected.Rows.Count <> 0 Then
                        Dim BOL_NO As String = grdAPTACRC0.ActiveRow.Cells("SOURCE_DOC_NO").Value & ""
                        Dim CTL_NO_MATCHED As String = grdAPTACRC0.ActiveRow.Cells("CTL_NO_MATCHED").Value & ""

                        Dim RECORDS As Int32 = grdAPTACRCO.Selected.Rows.Count
                        Dim COST_ACT_TOTAL As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("COST_ACT_TOTAL").Value & "")
                        Dim COST_ACC_TOTAL As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("COST_ACC_TOTAL").Value & "")
                        Dim TPV_ADJ_TOTAL As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("TPV_ADJ_TOTAL").Value & "")
                        Dim LNO_COUNT As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("LNO_COUNT").Value & "")

                        If MsgBox($"OK to Attach {grdAPTACRCO.Selected.Rows.Count} Accrual records selected to BOL No {BOL_NO}", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            For Each grow As UltraWinGrid.UltraGridRow In grdAPTACRCO.Selected.Rows
                                Dim CTL_NO As String = grow.Cells("CTL_NO").Value

                                COST_ACC_TOTAL += Val(grow.Cells("COST_ACC").Value & "")
                                TPV_ADJ_TOTAL += Val(grow.Cells("TPV_ADJ_TOTAL").Value & "")
                                LNO_COUNT += 1

                                ASCMAIN1.sql = $"Update {APTACRX1} Set BOL_NO = '{BOL_NO}', CTL_NO_MATCHED = '{CTL_NO_MATCHED}' where CTL_NO = '{CTL_NO}'"
                                ASCDATA1.ExecuteSQL()
                                'Dim rowAPTACRX1 As DataRow = dst.Tables("APTACRX1").Rows.Find(CTL_NO)
                                'rowAPTACRX1.Item("TPV_ADJ") = e.Row.Cells("TPV_ADJ").Value
                                'Update_Record_TDA("APTACRX1")

                                Dim rowAPTACRCO As DataRow = dst.Tables("APTACRCO").Rows.Find(CTL_NO)
                                rowAPTACRCO.Delete()
                                'grdAPTACRCO.ActiveRow.Cells("BOL_NO").Value = BOL_NO
                                'grdAPTACRCO.ActiveRow.Cells("CTL_NO_MATCHED").Value = CTL_NO_MATCHED
                                'grdAPTACRCO.ActiveRow.Update()

                                If CTL_NO_BOLs.ContainsKey(CTL_NO) Then
                                    CTL_NO_BOLs(CTL_NO) = BOL_NO
                                Else
                                    CTL_NO_BOLs.Add(CTL_NO, BOL_NO)
                                End If
                            Next

                            grdAPTACRC0.ActiveRow.Cells("COST_ACC_TOTAL").Value = COST_ACC_TOTAL
                            grdAPTACRC0.ActiveRow.Cells("TPV_ADJ_TOTAL").Value = TPV_ADJ_TOTAL
                            grdAPTACRC0.ActiveRow.Cells("LNO_COUNT").Value = LNO_COUNT
                            grdAPTACRC0.ActiveRow.Cells("DEF_TOLERANCE").Value = LNO_COUNT * 0.01
                            grdAPTACRC0.ActiveRow.Cells("CALC_VARIANCE").Value = COST_ACT_TOTAL - COST_ACC_TOTAL
                            grdAPTACRC0.ActiveRow.Update()

                            grdAPTACRC0.Rows.Refresh(RefreshRow.ReloadData)

                            Setup_grdAPTACRC0()
                            MsgBox($"{RECORDS} Orphan Records selected have been Attached to BOL {BOL_NO}", MsgBoxStyle.OkOnly, "Confirmation")
                        End If
                    End If
                End If

            Case "Detach from BOL"
                If grdAPTACRC0.ActiveRow IsNot Nothing AndAlso grdAPTACRC0.ActiveRow.IsDataRow Then

                    If grdAPTACRCT.Selected.Rows.Count = 0 Then
                        If grdAPTACRCT.ActiveRow IsNot Nothing Then
                            grdAPTACRCT.ActiveRow.Selected = True
                        End If
                    End If
                    If grdAPTACRCT.Selected.Rows.Count <> 0 Then

                        Dim BOL_NO As String = grdAPTACRC0.ActiveRow.Cells("SOURCE_DOC_NO").Value & ""
                        Dim CTL_NO_MATCHED As String = grdAPTACRC0.ActiveRow.Cells("CTL_NO_MATCHED").Value & ""

                        For Each grow As UltraWinGrid.UltraGridRow In grdAPTACRCT.Selected.Rows
                            If grow.Cells("BOL_NO").Value & "" <> BOL_NO Then
                                MsgBox($"Accruals must have been previously attached to BOL {BOL_NO} in order to Detach", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                                Exit Sub
                            End If
                        Next

                        Dim RECORDS As Int32 = grdAPTACRCT.Selected.Rows.Count
                        Dim COST_ACT_TOTAL As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("COST_ACT_TOTAL").Value & "")
                        Dim COST_ACC_TOTAL As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("COST_ACC_TOTAL").Value & "")
                        Dim TPV_ADJ_TOTAL As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("TPV_ADJ_TOTAL").Value & "")
                        Dim LNO_COUNT As Decimal = Val(grdAPTACRC0.ActiveRow.Cells("LNO_COUNT").Value & "")

                        If MsgBox($"OK to Detach {RECORDS} Accrual records selected from BOL No {BOL_NO}", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                            For Each grow As UltraWinGrid.UltraGridRow In grdAPTACRCT.Selected.Rows
                                Dim CTL_NO As String = grow.Cells("CTL_NO").Value

                                COST_ACC_TOTAL -= Val(grow.Cells("COST_ACC").Value & "")
                                TPV_ADJ_TOTAL -= Val(grow.Cells("TPV_ADJ_TOTAL").Value & "")
                                LNO_COUNT -= 1

                                ASCMAIN1.sql = $"Update {APTACRX1} Set BOL_NO = '', CTL_NO_MATCHED = '' where CTL_NO = '{CTL_NO}'"
                                ASCDATA1.ExecuteSQL()

                                Dim rowAPTACRCT As DataRow = dst.Tables("APTACRCT").Rows.Find(CTL_NO)
                                dst.Tables("APTACRCO").Rows.Add(rowAPTACRCT.ItemArray)

                                If CTL_NO_BOLs.ContainsKey(CTL_NO) Then
                                    CTL_NO_BOLs.Remove(CTL_NO)
                                Else
                                    CTL_NO_BOLs.Add(CTL_NO, "")
                                End If
                            Next

                            grdAPTACRC0.ActiveRow.Cells("COST_ACC_TOTAL").Value = COST_ACC_TOTAL
                            grdAPTACRC0.ActiveRow.Cells("TPV_ADJ_TOTAL").Value = TPV_ADJ_TOTAL
                            grdAPTACRC0.ActiveRow.Cells("LNO_COUNT").Value = LNO_COUNT
                            grdAPTACRC0.ActiveRow.Cells("DEF_TOLERANCE").Value = LNO_COUNT * 0.01
                            grdAPTACRC0.ActiveRow.Cells("CALC_VARIANCE").Value = COST_ACT_TOTAL - COST_ACC_TOTAL
                            grdAPTACRC0.ActiveRow.Update()

                            grdAPTACRC0.Rows.Refresh(RefreshRow.ReloadData)

                            Setup_grdAPTACRC0()
                            MsgBox($"{RECORDS} Records selected have been Detached from BOL {BOL_NO}", MsgBoxStyle.OkOnly, "Confirmation")
                        End If
                    End If
                End If

        End Select
    End Sub

#End Region


    Private Sub grdAPTACRC1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdAPTACRC1.AfterRowActivate
        If grdAPTACRC1.ActiveRow.IsGroupByRow Then
            tabDetails.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            Dim BOL_NO As String = grdAPTACRC1.ActiveRow.Cells("SOURCE_DOC_NO").Text
            Dim CTL_NO As String = grdAPTACRC1.ActiveRow.Cells("CTL_NO").Text

            grdAPTACRX1.Text = $"Accrued Tariff Details for BOL {BOL_NO}"

            If optStatus.Value = "M" Then
                Dim PPD_MATCHED_XNO As String = grdAPTACRC1.ActiveRow.Cells("PPD_MATCHED_XNO").Text
                ASCMAIN1.sql = "Select APTACRX1.* " _
                & $" from {APTACRX1} APTACRX1 WHERE PPD_MATCHED_XNO = '{PPD_MATCHED_XNO}' AND NVL(PPD_IND,'0') = '0' and BOL_NO_MATCHED = '{BOL_NO}'"
                Fill_Records("APTACRX1", ,, ASCMAIN1.sql)
            Else
                Fill_Records("APTACRX1", New String() {BOL_NO})
            End If

            grdASTAUDTX.Text = $"Audit Trail Details for BOL {BOL_NO}"
            Fill_Records("ASTAUDTX", New String() {CTL_NO})

            For Each row As DataRow In dst.Tables("APTACRX1").Select("")
                Fill_Records("ASTAUDTX", New String() {row.Item("CTL_NO")}, False)
            Next
            tabDetails.Visible = True
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub grdAPTACRC1_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdAPTACRC1.InitializeRow
        Dim variance As Decimal = Val(e.Row.Cells("DEF_TOLERANCE").Value) & ""
        If System.Math.Abs(Val(e.Row.Cells("CALC_VARIANCE").Value & "")) > System.Math.Abs(variance) Then
            e.Row.Appearance.BackColor = System.Drawing.Color.LightPink
            e.Row.ToolTipText = "Calculated Variances is greater than Tolerance"
        Else
            e.Row.Appearance.BackColor = System.Drawing.Color.Empty
            e.Row.ToolTipText = ""
        End If
    End Sub


    'Private Sub grdAPTACRX1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdAPTACRX1.AfterRowUpdate
    '    Dim TPV_ADJ_TOTAL As Decimal = Val(dst.Tables("APTACRX1").Compute("SUM(TPV_ADJ)", "") & "")
    '    Dim CTL_NO As String = grdAPTACRX1.ActiveRow.Cells("CTL_NO").Value
    '    CTL_NO_ADJs.Add(CTL_NO)
    '    Update_Record_TDA("APTACRX1")
    '    grdAPTACRC1.ActiveRow.Cells("TPV_ADJ_TOTAL").Value = TPV_ADJ_TOTAL
    '    grdAPTACRC1.ActiveRow.Update()
    'End Sub

    Sub Setup_grdAPTACRC0()
        If grdAPTACRC0.ActiveRow.IsGroupByRow Then
            tabDetails.Visible = False
            splMatch.Panel2.Visible = False
        Else
            splMatch.Panel2.Visible = True
            Me.Cursor = Cursors.WaitCursor
            Dim CTL_NO_MATCHED As String = grdAPTACRC0.ActiveRow.Cells("CTL_NO_MATCHED").Text
            Dim SOURCE_DOC_NO As String = grdAPTACRC0.ActiveRow.Cells("SOURCE_DOC_NO").Text

            Dim sfx As String = ""
            If grdAPTACRC0.ActiveRow.Band.Index = 1 Then
                sfx = " - Previous Match " & grdAPTACRC0.ActiveRow.Cells("CTL_NO_MATCHED").Text
            End If
            grdAPTACRCM.Text = $"Pre-Paid Tariff Details for BOL {SOURCE_DOC_NO}" & sfx
            grdAPTACRCT.Text = $"Accrued Tariff Details for BOL {SOURCE_DOC_NO}" & sfx

            ' & $" from {APTACRX1} APTACRX1, ICTIREC2, ICTCOSTA" & vbCrLf _
            If optStatus.Value = "M" Or grdAPTACRC0.ActiveRow.Band.Index = 1 Then
                ASCMAIN1.sql = "Select APTACRX1.* " & vbCrLf _
                    & ", (NVL(ICTIREC2.PO_COST,0) - NVL(ICTCOSTA.ITEM_COST_VCOST,0)) * NVL(ICTIREC2.QTY_REC,0) TRAN_PV_REC" & vbCrLf _
                    & ", (CASE WHEN NVL(ICTIREC2.ACCRUAL_STATUS,'0') = '1' THEN NVL(ICTIREC2.AMT_INV,0) - NVL(ICTIREC2.PO_COST,0) * NVL(ICTIREC2.QTY_INV,0) ELSE 0 END) TRAN_PV_INV" & vbCrLf _
                    & ", ICTIREC2.QTY_REC, ICTIREC2.PO_COST" & vbCrLf _
                    & $" from APTACRC1 APTACRX1, ICTIREC2, ICTCOSTA" & vbCrLf _
                    & $" where NVL(PPD_IND,'0') = '0' and CTL_NO_MATCHED = '{CTL_NO_MATCHED}'" & vbCrLf _
                    & "   and ICTIREC2.RECEIPT_NO = APTACRX1.RECEIPT_NO" & vbCrLf _
                    & "   and ICTIREC2.RECEIPT_LNO = APTACRX1.RECEIPT_LNO" & vbCrLf _
                    & "   and ICTCOSTA.OPS_YYYYPP (+) = APTACRX1.OPS_YYYYPP" & vbCrLf _
                    & "   and ICTCOSTA.ITEM_CODE (+) = APTACRX1.ITEM_CODE"
                Fill_Records("APTACRCT", ,, ASCMAIN1.sql) ' Accruals
            Else
                ASCMAIN1.sql = "Select APTACRX1.* " & vbCrLf _
                    & ", (NVL(ICTIREC2.PO_COST,0) - NVL(ICTCOSTA.ITEM_COST_VCOST,0)) * NVL(ICTIREC2.QTY_REC,0) TRAN_PV_REC" & vbCrLf _
                    & ", (CASE WHEN NVL(ICTIREC2.ACCRUAL_STATUS,'0') = '1' THEN NVL(ICTIREC2.AMT_INV,0) - NVL(ICTIREC2.PO_COST,0) * NVL(ICTIREC2.QTY_INV,0) ELSE 0 END) TRAN_PV_INV" & vbCrLf _
                    & ", ICTIREC2.QTY_REC, ICTIREC2.PO_COST" & vbCrLf _
                    & $" from {APTACRX1} APTACRX1, ICTIREC2, ICTCOSTA" & vbCrLf _
                    & $" where NVL(PPD_IND,'0') = '0' and CTL_NO_MATCHED = '{CTL_NO_MATCHED}'" & vbCrLf _
                    & "   and ICTIREC2.RECEIPT_NO = APTACRX1.RECEIPT_NO" & vbCrLf _
                    & "   and ICTIREC2.RECEIPT_LNO = APTACRX1.RECEIPT_LNO" & vbCrLf _
                    & "   and ICTCOSTA.OPS_YYYYPP (+) = APTACRX1.OPS_YYYYPP" & vbCrLf _
                    & "   and ICTCOSTA.ITEM_CODE (+) = APTACRX1.ITEM_CODE"
                Fill_Records("APTACRCT", ,, ASCMAIN1.sql) ' Accruals
            End If

            If optStatus.Value = "M" Or grdAPTACRC0.ActiveRow.Band.Index = 1 Then
                ASCMAIN1.sql = "Select APTACRC1.* " _
                    & $" from APTACRC1 WHERE NVL(PPD_IND,'0') = '1' and CTL_NO_MATCHED = '{CTL_NO_MATCHED}'"
                Fill_Records("APTACRCM", ,, ASCMAIN1.sql) ' Pre-Paid
            Else
                ASCMAIN1.sql = "Select APTACRX1.* " _
                    & $" from {APTACRX1} APTACRX1 WHERE NVL(PPD_IND,'0') = '1' and CTL_NO_MATCHED = '{CTL_NO_MATCHED}'"
                Fill_Records("APTACRCM", ,, ASCMAIN1.sql) ' Pre-Paid
            End If

            If optStatus.Value = "M" Or grdAPTACRC0.ActiveRow.Band.Index = 1 Then
                ' no need to calc variance
            Else
                'If grdAPTACRC0.ActiveRow.Index = 0 Then
                Variances_by_Line_Item(CTL_NO_MATCHED)
                'End If
            End If

            tabAccruals.Tabs("Matched").Selected = True

            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub grdAPTACRC0_AfterRowActivate(sender As Object, e As EventArgs) Handles grdAPTACRC0.AfterRowActivate
        If grdAPTACRC0.ActiveRow IsNot Nothing AndAlso grdAPTACRC0.ActiveRow.IsDataRow AndAlso Not grdAPTACRC0.ActiveRow.IsGroupByRow Then
            If grdAPTACRC0.ActiveRow.Band.Index = 0 AndAlso True Then

            End If
            Setup_grdAPTACRC0()
            splMatch.Panel2.Visible = True
        Else
            splMatch.Panel2.Visible = False
        End If

    End Sub

    Private Sub grdAPTACRC0_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdAPTACRC0.InitializeRow
        If e.Row.Band.Index = 1 Then Exit Sub

        If optStatus.Value = "M" Then
        Else
            With e.Row.Cells("CALC_VARIANCE")
                Dim DEF_TOLERANCE As Decimal = Val(e.Row.Cells("DEF_TOLERANCE").Value & "")
                Dim VAR_OK As String = e.Row.Cells("VAR_OK").Value & ""
                Dim CALC_VARIANCE As Decimal = Val(e.Row.Cells("CALC_VARIANCE").Value & "")
                Dim VAR_TOLERANCE As Decimal = Val(e.Row.Cells("VAR_TOLERANCE").Value & "")
                Dim LNO_COUNT As Integer = Val(e.Row.Cells("LNO_COUNT").Value & "")

                e.Row.Cells("CALC_VARIANCE").Appearance.BackColor = System.Drawing.Color.Empty
                e.Row.Cells("CALC_VARIANCE").ToolTipText = ""
                e.Row.Cells("SOURCE_DOC_NO").Appearance.BackColor = System.Drawing.Color.Empty
                e.Row.Cells("SOURCE_DOC_NO").ToolTipText = ""
                e.Row.Cells("CTL_NO_MATCHED").Appearance.BackColor = System.Drawing.Color.Empty
                e.Row.Cells("CTL_NO_MATCHED").ToolTipText = ""
                e.Row.Cells("VAR_OK").Appearance.BackColor = System.Drawing.Color.Empty
                e.Row.Cells("VAR_OK").ToolTipText = ""

                If LNO_COUNT > 0 Then
                    If System.Math.Abs(CALC_VARIANCE) > DEF_TOLERANCE And VAR_OK <> "1" Then
                        'If VAR_OK <> "1" Then
                        e.Row.Cells("CALC_VARIANCE").Appearance.BackColor = System.Drawing.Color.LightPink
                        e.Row.Cells("CALC_VARIANCE").ToolTipText = "Calculated Variance is greater than Tolerance - need to OK to Match"
                        'End If
                    Else
                        If System.Math.Abs(CALC_VARIANCE) > VAR_TOLERANCE Then
                            If VAR_OK <> "1" Then
                                ' NOT OK FOR MATCH REPORT
                            Else
                                e.Row.Cells("VAR_OK").Appearance.BackColor = System.Drawing.Color.Red
                                e.Row.Cells("VAR_OK").ToolTipText = $"Variance {Format(System.Math.Abs(CALC_VARIANCE), "$#,##0.00")} is Greater than the Previous value {Format(VAR_TOLERANCE, "$#,##0.00")} - Toggle to Reset"

                                e.Row.Cells("CTL_NO_MATCHED").Appearance.BackColor = System.Drawing.Color.LightBlue
                                e.Row.Cells("CTL_NO_MATCHED").ToolTipText = "This BOL is a candidate for the PPD Match Report - if OK is Toggled"
                            End If
                        Else
                            e.Row.Cells("CTL_NO_MATCHED").Appearance.BackColor = System.Drawing.Color.LightGreen
                            e.Row.Cells("CTL_NO_MATCHED").ToolTipText = "This BOL is a candidate for the PPD Match Report"
                        End If
                    End If

                End If
            End With

        End If

        With e.Row.Cells("LNO_COUNT")
            If System.Math.Abs(Val(.Value & "")) = 0 Then
                .Appearance.ForeColor = System.Drawing.Color.Red
                .ToolTipText = "No Accruals Matching this BOL (so far)"
            Else
                .Appearance.ForeColor = System.Drawing.Color.Empty
                .ToolTipText = ""
            End If
        End With

        With e.Row.Cells("SOURCE_DOC_NO")
            Dim CTL_NO_MATCHED As String = e.Row.Cells("CTL_NO_MATCHED").Value & ""
            Dim rowAPTACRC0 As DataRow = dst.Tables("APTACRC0").Rows.Find(CTL_NO_MATCHED)
            If rowAPTACRC0 IsNot Nothing Then
                Dim BOL_NO_orig As String = rowAPTACRC0.Item("SOURCE_DOC_NO", DataRowVersion.Original) & ""
                If .Value & "" <> BOL_NO_orig Then
                    .Appearance.ForeColor = System.Drawing.Color.Red
                    .ToolTipText = $"Original BOL Value was {BOL_NO_orig}"
                Else
                    .Appearance.ForeColor = System.Drawing.Color.Empty
                    .ToolTipText = ""
                End If
            End If
        End With
    End Sub

    Private Sub grdAPTACRCT_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdAPTACRCT.AfterRowUpdate
        Dim TPV_ADJ_TOTAL As Decimal = Val(dst.Tables("APTACRCT").Compute("SUM(TPV_ADJ)", "") & "")
        Dim TPV_ADJ As Decimal = Val(e.Row.Cells("TPV_ADJ").Value & "")
        Dim CTL_NO As String = grdAPTACRCT.ActiveRow.Cells("CTL_NO").Value
        If CTL_NO_ADJs.ContainsKey(CTL_NO) Then
            CTL_NO_ADJs(CTL_NO) = TPV_ADJ
        Else
            CTL_NO_ADJs.Add(CTL_NO, TPV_ADJ)
        End If

        ASCMAIN1.sql = $"Update {APTACRX1} Set TPV_ADJ = {CStr(TPV_ADJ)} where CTL_NO = '{CTL_NO}'"
        ASCDATA1.ExecuteSQL()
        'Dim rowAPTACRX1 As DataRow = dst.Tables("APTACRX1").Rows.Find(CTL_NO)
        'rowAPTACRX1.Item("TPV_ADJ") = e.Row.Cells("TPV_ADJ").Value
        'Update_Record_TDA("APTACRX1")
        grdAPTACRC0.ActiveRow.Cells("TPV_ADJ_TOTAL").Value = TPV_ADJ_TOTAL
        grdAPTACRC0.ActiveRow.Update()
    End Sub

    Private Sub grdAPTACRC0_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdAPTACRC0.BeforeRowUpdate
        Dim VAR_OK As String = e.Row.Cells("VAR_OK").Text
        If VAR_OK = "1" Then
            e.Row.Cells("VAR_TOLERANCE").Value = System.Math.Abs(Val(e.Row.Cells("CALC_VARIANCE").Value & ""))
        Else
            e.Row.Cells("VAR_TOLERANCE").Value = 0
        End If
    End Sub

    Private Sub chkShowReToggle_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowReToggle.CheckedChanged

        Dim dvw As DataView = DirectCast(grdAPTACRC0.DataSource, DataTable).DefaultView

        Dim RR() As DataRow = dst.Tables("APTACRC0").Select("VAR_OK = '1' and ABS(ISNULL(CALC_VARIANCE,0)) > ISNULL(VAR_TOLERANCE,0)")
        Dim RR1() As DataRow = dst.Tables("APTACRC0").Select("VAR_OK = '1'")
        Dim RR2() As DataRow = dst.Tables("APTACRC0").Select("VAR_TOLERANCE < ABS(CALC_VARIANCE)")
        Dim RR3() As DataRow = dst.Tables("APTACRC0").Select("VAR_OK = '1' and ABS(ISNULL(CALC_VARIANCE,0)) = 24.48")
        Dim RR4() As DataRow = dst.Tables("APTACRC0").Select("VAR_OK = '1' and ABS(ISNULL(VAR_TOLERANCE,0)) = 24")
        If chkShowReToggle.Checked Then
            dvw.RowFilter = "VAR_OK = '1' and VAR_TOLERANCE < ABS(CALC_VARIANCE)"
            'dvw.RowFilter = "VAR_OK = '1'"
        Else
            dvw.RowFilter = ""
        End If

    End Sub

    Sub Variances_by_Line_Item(CTL_NO_MATCHED As String)
        If CTL_NO_MATCHED = "" Then
            Exit Sub
        End If

        Dim rowAPTACRC0 As DataRow = dst.Tables("APTACRC0").Rows.Find(CTL_NO_MATCHED)
        Dim SOURCE_DOC_NO As String = rowAPTACRC0.Item("SOURCE_DOC_NO")

        Dim CALC_VARIANCE As Decimal = Val(rowAPTACRC0.Item("CALC_VARIANCE") & "")
        Dim TPV_ADJ_TOTAL As Decimal = Val(rowAPTACRC0.Item("TPV_ADJ_TOTAL") & "")
        Dim CALC_VARIANCE_NO_ADJ = CALC_VARIANCE - TPV_ADJ_TOTAL
        If CALC_VARIANCE_NO_ADJ = 0 Then

            For Each rowAPTACRCT As DataRow In dst.Tables("APTACRCT").Select($"CTL_NO_MATCHED = '{CTL_NO_MATCHED}'")
                rowAPTACRCT.Item("COST_VAR_ITEM") = DBNull.Value
            Next

        Else

            Dim COST_VARs As New Dictionary(Of String, Decimal)
            Dim COST_VAR_ADJs As New Dictionary(Of String, Decimal)

            Dim COST_ACC_TOTAL As Decimal = 0 ' Val(rowAPTACRC0.Item("COST_ACC_TOTAL") & "")

            ' Spread by BOL/COST_CATGY_CODE

            dst.Tables("APTACRXA").Rows.Clear()
            For Each rowAPTACRCT As DataRow In dst.Tables("APTACRCT").Select($"CTL_NO_MATCHED = '{CTL_NO_MATCHED}'", "COST_ACC DESC")
                Dim TPV_ADJ As Decimal = Val(rowAPTACRCT.Item("TPV_ADJ") & "")
                If TPV_ADJ <> 0 Then
                    rowAPTACRCT.Item("COST_VAR_ITEM") = TPV_ADJ
                Else
                    Dim COST_CATGY_CODE As String = rowAPTACRCT.Item("COST_CATGY_CODE")
                    Dim rowAPTACRXA As DataRow = dst.Tables("APTACRXA").Rows.Find(New String() {CTL_NO_MATCHED, COST_CATGY_CODE})
                    If rowAPTACRXA Is Nothing Then
                        rowAPTACRXA = dst.Tables("APTACRXA").NewRow
                        rowAPTACRXA.Item("CTL_NO_MATCHED") = CTL_NO_MATCHED
                        rowAPTACRXA.Item("COST_CATGY_CODE") = COST_CATGY_CODE
                        rowAPTACRXA.Item("SOURCE_DOC_NO") = SOURCE_DOC_NO
                        dst.Tables("APTACRXA").Rows.Add(rowAPTACRXA)
                    End If
                    Dim COST_ACC As Decimal = Val(rowAPTACRCT.Item("COST_ACC") & "")
                    rowAPTACRXA.Item("COST_ACC") = Val(rowAPTACRXA.Item("COST_ACC") & "") + COST_ACC
                    COST_ACC_TOTAL += COST_ACC
                End If
            Next

            If COST_ACC_TOTAL <> 0 Then
                For Each rowAPTACRXA As DataRow In dst.Tables("APTACRXA").Select
                    Dim COST_CATGY_CODE As String = rowAPTACRXA.Item("COST_CATGY_CODE")
                    Dim COST_ACC As Decimal = Val(rowAPTACRXA.Item("COST_ACC") & "")
                    Dim COST_VAR_SNU As Decimal = System.Math.Round(CALC_VARIANCE_NO_ADJ * COST_ACC / COST_ACC_TOTAL, 2)
                    COST_VARs.Add(COST_CATGY_CODE, COST_VAR_SNU)
                Next
            End If


            ' Spread by ITEM

            Dim blnIsolate_ADJ As Boolean = True

            For Each SNU As String In COST_VARs.Keys
                Dim COST_VAR_SNU As Decimal = COST_VARs(SNU)
                Dim sqlwSNU As String = $"CTL_NO_MATCHED = '{CTL_NO_MATCHED}' and COST_CATGY_CODE = '{SNU}' and ISNULL(PPD_IND,'0') = '0'"

                Dim COST_VAR_ADJ_TOTAL As Decimal = 0

                If blnIsolate_ADJ Then
                    Dim rowAPTACRCT_ADJs() As DataRow = dst.Tables("APTACRCT").Select(sqlwSNU & " and ISNULL(TPV_ADJ,0) <> 0")
                    For Each rowAPTACRCT As DataRow In rowAPTACRCT_ADJs
                        Dim COST_VAR_ITEM As Decimal = Val(rowAPTACRCT.Item("TPV_ADJ") & "")
                        rowAPTACRCT.Item("COST_VAR_ITEM") = COST_VAR_ITEM
                        COST_VAR_ADJ_TOTAL += COST_VAR_ITEM
                    Next
                End If

                Dim rowAPTACRCTs() As DataRow = dst.Tables("APTACRCT").Select(sqlwSNU & IIf(blnIsolate_ADJ, " and ISNULL(TPV_ADJ,0) = 0", ""), "COST_ACC DESC")
                'COST_VAR_SNU -= COST_VAR_ADJ_TOTAL

                If rowAPTACRCTs.Length = 0 Then
                    rowAPTACRCTs = dst.Tables("APTACRCT").Select(sqlwSNU, "COST_ACC DESC")
                End If

                If rowAPTACRCTs.Length = 1 Then
                    rowAPTACRCTs(0).Item("COST_VAR_ITEM") = COST_VAR_SNU
                Else
                    Dim COST_ACC_TOTAL_SNU As Decimal = Val(dst.Tables("APTACRCT").Compute("SUM(COST_ACC)", sqlwSNU) & "")
                    Dim COST_VAR_TOTAL As Decimal = 0
                    Dim COST_VAR_ITEM As Decimal = 0
                    For Each rowAPTACRCT As DataRow In rowAPTACRCTs
                        Dim COST_ACC As Decimal = Val(rowAPTACRCT.Item("COST_ACC") & "")
                        COST_VAR_ITEM = System.Math.Round(COST_VAR_SNU * COST_ACC / COST_ACC_TOTAL_SNU, 2)
                        If COST_VAR_ITEM <> 0 Then
                            rowAPTACRCT.Item("COST_VAR_ITEM") = COST_VAR_ITEM
                            COST_VAR_TOTAL += COST_VAR_ITEM
                        End If
                        If COST_VAR_TOTAL = COST_VAR_SNU Then
                            Exit For
                        End If
                    Next
                    If COST_VAR_TOTAL <> COST_VAR_SNU Then
                        rowAPTACRCTs(0).Item("COST_VAR_ITEM") = Val(rowAPTACRCTs(0).Item("COST_VAR_ITEM") & "") + COST_VAR_SNU - COST_VAR_TOTAL
                    End If
                End If
            Next
        End If
    End Sub

    Sub Get_PREV_APTACRC0()

        If optStatus.Value = "M" Then
            ASCMAIN1.sql = $"Select * from APTACRC0 where CTL_NO_MATCHED_NEXT is Not Null"
            Fill_Records("APTACRCR", , , ASCMAIN1.sql)
        Else
            ASCMAIN1.sql = $"Select Distinct SOURCE_DOC_NO from {APTACRX1} where PPD_IND = '1'"
            ASCMAIN1.sql = $"Select * from APTACRC0 where CTL_NO_MATCHED_NEXT is null and SOURCE_DOC_NO in ({ASCMAIN1.sql})"
            Fill_Records("APTACRCR", , , ASCMAIN1.sql)

            For Each rowAPTACRCR As DataRow In dst.Tables("APTACRCR").Select("")
                Dim CTL_NO_MATCHED_PREV As String = rowAPTACRCR.Item("CTL_NO_MATCHED")
                Dim SOURCE_DOC_NO As String = rowAPTACRCR.Item("SOURCE_DOC_NO")
                Dim rowAPTACRC0 As DataRow = dst.Tables("APTACRC0").Select($"SOURCE_DOC_NO = '{SOURCE_DOC_NO}'")(0)
                Dim CTL_NO_MATCHED_NEXT As String = rowAPTACRC0.Item("CTL_NO_MATCHED")
                rowAPTACRCR.Item("CTL_NO_MATCHED_NEXT") = CTL_NO_MATCHED_NEXT
                rowAPTACRC0.Item("CTL_NO_MATCHED_PREV") = CTL_NO_MATCHED_PREV

                ASCMAIN1.sql = $"Select APTACRC1.*, ICTITEM1.ITEM_DESC" & vbCrLf _
                    & ", NULL DEF_TOLERANCE, NULL COST_ACC_TOTAL, NULL TPV_ADJ_TOTAL, NULL LNO_COUNT" & vbCrLf _
                    & " from APTACRC1, ICTITEM1" & vbCrLf _
                    & " where ICTITEM1.ITEM_CODE (+) = APTACRC1.ITEM_CODE" & vbCrLf _
                    & $" and CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
                ASCMAIN1.sql = $"Insert into {APTACRX1} {ASCMAIN1.sql}"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = $"Update {APTACRX1} Set CTL_NO_MATCHED = '{CTL_NO_MATCHED_NEXT}' where CTL_NO_MATCHED = '{CTL_NO_MATCHED_PREV}'"
                ASCDATA1.ExecuteSQL()

                ASCMAIN1.sql = $"Select Sum (COST_ACT) COST_ACT_TOTAL" & vbCrLf _
                    & $" from {APTACRX1} where CTL_NO_MATCHED = '{CTL_NO_MATCHED_NEXT}' and NVL(PPD_IND,'0') = '1'"
                Dim rowM As DataRow = ASCDATA1.GetDataRow
                Dim COST_ACT_TOTAL As Decimal = Val(rowM.Item("COST_ACT_TOTAL") & "")
                rowAPTACRC0.Item("COST_ACT_TOTAL") = COST_ACT_TOTAL

                ASCMAIN1.sql = $"Select Count(*) LNO_COUNT, Sum (COST_ACC) COST_ACC_TOTAL, Sum (TPV_ADJ_TOTAL) TPV_ADJ_TOTAL" & vbCrLf _
                    & $" from {APTACRX1} where CTL_NO_MATCHED = '{CTL_NO_MATCHED_NEXT}' and NVL(PPD_IND,'0') = '0'"
                Dim rowT As DataRow = ASCDATA1.GetDataRow
                Dim LNO_COUNT As Int32 = Val(rowT.Item("LNO_COUNT") & "")
                Dim COST_ACC_TOTAL As Decimal = Val(rowT.Item("COST_ACC_TOTAL") & "")
                Dim TPV_ADJ_TOTAL As Decimal = Val(rowT.Item("TPV_ADJ_TOTAL") & "")
                rowAPTACRC0.Item("LNO_COUNT") = LNO_COUNT
                rowAPTACRC0.Item("COST_ACC_TOTAL") = COST_ACC_TOTAL
                rowAPTACRC0.Item("TPV_ADJ_TOTAL") = TPV_ADJ_TOTAL
                rowAPTACRC0.Item("DEF_TOLERANCE") = 0.01 * LNO_COUNT
                rowAPTACRC0.Item("CALC_VARIANCE") = COST_ACT_TOTAL - COST_ACC_TOTAL
                ' $", CASE WHEN NVL({APTACRCX}.LNO_COUNT, 0) > 0 THEN NVL({APTACRCX}.COST_ACT_TOTAL, 0) - NVL({APTACRCX}.COST_ACC_TOTAL, 0) ELSE NULL END as CALC_VARIANCE" & vbCrLf &

            Next
        End If
    End Sub
End Class