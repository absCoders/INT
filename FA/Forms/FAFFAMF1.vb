Public Class FAFFAMF1
    Dim ASSET_NO As String

    Dim rowFATFAMF1 As DataRow
    Dim sqlFATFATRX As String

    Dim YYYY_FIRST As Integer = 0
    Dim YYYY_LAST As Integer = Val(Mid(ASCMAIN1.CYP, 1, 4))


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "FAFFAMFI" Then
            InquiryMode = True
        End If

        Get_PARM("GLTPARM1")

        AUDIT.Add("FATFAMF1", "*")
        AUDIT.Add("FATFAMF2", "*")

        With dst

            Create_TDA(.Tables.Add, "FATFAMF1", "*")
            Create_TDA(.Tables.Add, "FATFAMF2", "*")

            ASCMAIN1.sql = "Select ASTAUDT1.*" & vbCrLf _
                & " from ASTAUDT1" & vbCrLf _
                & " where TABLE_NAME = 'FATFAMF1'"
            Create_TDA(.Tables.Add, "ASTAUDTX", "**", 0, False)

            ASCMAIN1.sql = "Select FATFAMF1.*" & vbCrLf _
                & " from FATFAMF1"
            Create_TDA(.Tables.Add, "FATFAMFX", "**", 0, False, "", 1)
            With .Tables("FATFAMFX")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "Select FATFATR1.*" & vbCrLf _
                & " from FATFATR1" & vbCrLf _
                & " where ASSET_NO = :PARM1"
            Create_TDA(.Tables.Add, "FATFATR1", "**", 0, False, "V", 1)
            With .Tables("FATFATR1")
                .Columns.Add("ASSET_AMT", GetType(System.Decimal), "IIF(ASSET_TRN_TYPE = 'CAP',ASSET_TRN_AMT,NULL)")
                .Columns.Add("ASSET_DEP", GetType(System.Decimal), "IIF(ASSET_TRN_TYPE = 'EXP',ASSET_TRN_AMT,NULL)")
                .Columns.Add("ASSET_WOF", GetType(System.Decimal), "IIF(ASSET_TRN_TYPE = 'WOF',ASSET_TRN_AMT,NULL)")
                .Columns.Add("ASSET_BAL", GetType(System.Decimal))
            End With


            ASCMAIN1.sql = "SELECT ASSET_NO, SUBSTR(OPS_YYYYPP,1,4) YEAR" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'01',ASSET_TRN_AMT,NULL)) M01" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'02',ASSET_TRN_AMT,NULL)) M02" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'03',ASSET_TRN_AMT,NULL)) M03" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'04',ASSET_TRN_AMT,NULL)) M04" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'05',ASSET_TRN_AMT,NULL)) M05" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'06',ASSET_TRN_AMT,NULL)) M06" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'07',ASSET_TRN_AMT,NULL)) M07" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'08',ASSET_TRN_AMT,NULL)) M08" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'09',ASSET_TRN_AMT,NULL)) M09" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'10',ASSET_TRN_AMT,NULL)) M10" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'11',ASSET_TRN_AMT,NULL)) M11" & vbCrLf _
                & ", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'12',ASSET_TRN_AMT,NULL)) M12" & vbCrLf _
                & " from FATFATR1" & vbCrLf _
                & " where (ASSET_TRN_TYPE = 'EXP' OR ASSET_TRN_TYPE = 'WOF')" & vbCrLf _
                & "   and ASSET_NO = :PARM1" & vbCrLf _
                & " group by ASSET_NO, SUBSTR(OPS_YYYYPP,1,4)"
            sqlFATFATRX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "FATFATRX", "**", 0, False, "V", 2)
            With .Tables("FATFATRX")
                .Columns.Add("BEG_BAL", GetType(System.Decimal))
                .Columns.Add("TOTAL_DWA", GetType(System.Decimal), "ISNULL(M01,0)+ISNULL(M02,0)+ISNULL(M03,0)+ISNULL(M04,0)+ISNULL(M05,0)+ISNULL(M06,0)+ISNULL(M07,0)+ISNULL(M08,0)+ISNULL(M09,0)+ISNULL(M10,0)+ISNULL(M11,0)+ISNULL(M12,0)")
                .Columns.Add("END_BAL", GetType(System.Decimal), "ISNULL(BEG_BAL,0) - TOTAL_DWA")
            End With

            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO, APTINVH2.VOUCHER_LNO, APTINVH1.OPS_YYYYPP" & vbCrLf _
                & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME, APTINVH1.INV_DATE, APTINVH1.INV_NUM" & vbCrLf _
                & ", APTINVH1.INV_STATUS, APTINVH1.INV_REF, APTINVH1.INIT_OPER, APTINVH1.INIT_DATE" & vbCrLf _
                & ", APTINVH2.INV_LINE_AMT" & vbCrLf _
                & ", APTINVH2.ACCT_CODE, X.ASSET_CLASS_CODE" & vbCrLf _
                & " from APTINVH1, APTINVH2, APTVEND1, (Select ACCT_CODE_CAP ACCT_CODE, MIN (ASSET_CLASS_CODE) ASSET_CLASS_CODE from FATFACL1 group by ACCT_CODE_CAP) X" & vbCrLf _
                & " where APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" & vbCrLf _
                & "   and APTINVH1.OPS_YYYYPP >= :PARM1" & vbCrLf _
                & "   and APTINVH2.ACCT_CODE IN (Select ACCT_CODE_CAP from FATFACL1)" & vbCrLf _
                & "   and X.ACCT_CODE = APTINVH2.ACCT_CODE" & vbCrLf _
                & "   and APTVEND1.VEND_CODE = APTINVH1.VEND_CODE" & vbCrLf _
                & " UNION" & vbCrLf _
                & "Select 'GLJE' || GLTDETL1.JOURNAL_NO, GLTDETL1.JOURNAL_LNO, GLTDETL1.OPS_YYYYPP" & vbCrLf _
                & ", NULL VEND_CODE, NULL VEND_NAME, GLTDETL1.DETL_CTL_DATE INV_DATE, GLTDETL1.DETL_CVX_REF_NO INV_NUM" & vbCrLf _
                & ", NULL INV_STATUS, GLTJRNL1.JOURNAL_DESC INV_REF, GLTJRNL1.INIT_OPER, GLTJRNL1.INIT_DATE" & vbCrLf _
                & ", GLTDETL1.DETL_POSTING_AMT INV_LINE_AMT" & vbCrLf _
                & ", GLTDETL1.ACCT_CODE, X.ASSET_CLASS_CODE" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1, (Select ACCT_CODE_CAP ACCT_CODE, MIN (ASSET_CLASS_CODE) ASSET_CLASS_CODE from FATFACL1 group by ACCT_CODE_CAP) X" & vbCrLf _
                & " where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_TYPE = 'GLJE'" & vbCrLf _
                & "   and GLTDETL1.OPS_YYYYPP >=  :PARM1" & vbCrLf _
                & "   and GLTDETL1.ACCT_CODE IN (Select ACCT_CODE_CAP from FATFACL1)" & vbCrLf _
                & "   and X.ACCT_CODE = GLTDETL1.ACCT_CODE"
            ASCMAIN1.sql = $"Select X.*, NVL(FATFAMF2.IGNORE,'0') IGNORE from ({ASCMAIN1.sql}) X, FATFAMF2" & vbCrLf _
                & " where FATFAMF2.VOUCHER_NO (+) = X.VOUCHER_NO" & vbCrLf _
                & "   and FATFAMF2.VOUCHER_LNO (+) = X.VOUCHER_LNO"
            Create_TDA(.Tables.Add, "APTINVHX", "**", 0, False, "V", 2)
            With .Tables("APTINVHX")
                .Columns("IGNORE").DefaultValue = "0"
                .Columns.Add("ASSET_AMT", GetType(System.Decimal))
                .Columns.Add("ASSET_REQ", GetType(System.Decimal), "IIF(ISNULL(IGNORE,'0')='1',0,ISNULL(INV_LINE_AMT,0) - ISNULL(ASSET_AMT,0))")
            End With
        End With

        grdFATFAMFX.DataSource = dst.Tables("FATFAMFX")
        grdFATFATR1.DataSource = dst.Tables("FATFATR1")
        grdFATFATRX.DataSource = dst.Tables("FATFATRX")
        grdASTAUDTX.DataSource = dst.Tables("ASTAUDTX")

        grdAPTINVHX.DataSource = dst.Tables("APTINVHX")

        Create_Summary(grdFATFAMFX, "ASSET_NO", "Count")
        Create_Summary(grdFATFAMFX, New String() {"SEL"})
        Create_Summary(grdFATFAMFX, New String() {"ASSET_AMT", "ASSET_DEP", "ASSET_WOF", "ASSET_BAL"})

        Create_Summary(grdFATFATR1, "ASSET_TRN_NO", "Count")
        Create_Summary(grdFATFATR1, New String() {"ASSET_TRN_AMT"})
        Create_Summary(grdFATFATR1, New String() {"ASSET_AMT", "ASSET_DEP", "ASSET_WOF", "ASSET_BAL"})

        Create_Summary(grdFATFATRX, New String() {"M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12", "TOTAL_DWA"})

        Create_Summary(grdAPTINVHX, "VOUCHER_NO", "Count")
        Create_Summary(grdAPTINVHX, New String() {"INV_LINE_AMT", "ASSET_AMT", "ASSET_REQ"})

        grdFATFAMFX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdFATFAMFX.DisplayLayout.UseFixedHeaders = True
        With grdFATFAMFX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    Select Case gcol.Key
                        Case "SEL", "ASSET_ACTION"
                            .BackColor2 = System.Drawing.Color.Orange
                            If gcol.Key = "SEL" Then
                                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                            End If
                        Case "ASSET_AMT", "ASSET_DEP", "ASSET_WOF", "ASSET_BAL"
                            .BackColor2 = System.Drawing.Color.LightGreen
                        Case "VEND_CODE", "VEND_NAME", "VOUCHER_NO", "INVOICE_NOTES"
                            .BackColor2 = System.Drawing.Color.LightGoldenrodYellow
                        Case "OPS_YYYYPP", "ASSET_DATE"
                            .BackColor2 = System.Drawing.Color.LightPink
                        Case "OPS_YYYYPP_IN_SERVICE", "ASSET_DATE_IN_SERVICE"
                            .BackColor2 = System.Drawing.Color.Gold
                        Case "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"
                            .BackColor2 = System.Drawing.Color.LightGray
                        Case Else
                            .BackColor2 = System.Drawing.Color.LightBlue
                    End Select
                    If New String() {"ASSET_ACTION", "ASSET_STATUS"}.Contains(gcol.Key) Then
                        gcol.CellAppearance.TextHAlign = HAlign.Center
                        gcol.Header.Appearance.TextHAlign = HAlign.Center
                    End If
                End With
            Next
            .Columns("SEL").Header.Fixed = True
            .Columns("ASSET_NO").Header.Fixed = True
            .Columns("ASSET_DESC").Header.Fixed = True
        End With

        grdFATFATR1.DisplayLayout.UseFixedHeaders = True
        With grdFATFATR1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    Select Case gcol.Key
                        Case "ASSET_TRN_AMT"
                            .BackColor2 = System.Drawing.Color.Orange
                        Case "ASSET_AMT", "ASSET_DEP", "ASSET_WOF", "ASSET_BAL"
                            If gcol.Key = "ASSET_AMT" Or gcol.Key = "ASSET_BAL" Then
                                gcol.Hidden = True
                            End If
                            .BackColor2 = System.Drawing.Color.LightGreen
                        Case "VOUCHER_NO", "VOUCHER_LNO", "JOURNAL_NO", "JOURNAL_LNO"
                            .BackColor2 = System.Drawing.Color.LightGoldenrodYellow
                        Case "OPS_YYYYPP"
                            .BackColor2 = System.Drawing.Color.LightPink
                        Case "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"
                            .BackColor2 = System.Drawing.Color.LightGray
                        Case Else
                            .BackColor2 = System.Drawing.Color.LightBlue
                    End Select

                End With
            Next
            .Columns("ASSET_TRN_NO").Header.Fixed = True
            .Columns("ASSET_TRN_DATE").Header.Fixed = True
            .Columns("ASSET_TRN_TYPE").Header.Fixed = True
            .Columns("ASSET_TRN_AMT").Header.Fixed = True
        End With

        grdFATFATRX.DisplayLayout.UseFixedHeaders = True
        With grdFATFATRX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    Select Case gcol.Key
                                                'Case "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12"
                        '    .BackColor2 = System.Drawing.Color.LightGreen
                        Case "BEG_BAL", "TOTAL_DWA", "END_BAL"
                            .BackColor2 = System.Drawing.Color.LightPink
                            gcol.Width = 95
                        Case "YEAR"
                            .BackColor2 = System.Drawing.Color.LightPink
                        Case Else
                            .BackColor2 = System.Drawing.Color.LightBlue
                    End Select
                End With
            Next

            For m As Integer = 1 To 12
                Dim MM As String = Format(m, "00")
                Dim c As String = "M" & MM
                .Columns(c).Header.Caption = Mid(Format(CDate($"{MM}/01/2023"), "dd-MMM-yyyy"), 4, 3)

                .Columns(c).Width = 80
                .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
            Next
            .Columns("YEAR").Header.Fixed = True
        End With

        grdAPTINVHX.DisplayLayout.UseFixedHeaders = True
        grdAPTINVHX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        With grdAPTINVHX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "IGNORE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    Select Case gcol.Key
                        Case Else
                            .BackColor2 = System.Drawing.Color.LightBlue
                    End Select
                End With
            Next
        End With

        grdFATFAMFX.DisplayLayout.GroupByBox.Hidden = False
        Show_Filter(grdFATFAMFX, True)

        Show_Filter(grdASTAUDTX, True)

        Bind_Controls(grpAMTS, "FATFAMF1")

        MakeTransparent(chkEdit)
        MakeTransparent(chkHideIgnored)
        MakeTransparent(chkHideDisposed)

        ASCMAIN1.Add_Value_List(grdFATFAMFX, "ASSET_STATUS", , New String() {":", "C:Capitalized", "D:Depreciated", "X:Disposed"})
        ASCMAIN1.Add_Value_List(grdFATFAMFX, "ASSET_ACTION", , New String() {":", "N:No Action", "D:Depreciate", "W:Write-Off"})
        ASCMAIN1.Add_Value_List(grdFATFAMFX, "ASSET_CLASS_CODE", "Select ASSET_CLASS_CODE, ASSET_CLASS_DESC from FATFACL1")
        ASCMAIN1.Add_Value_List(grdFATFAMFX, "ASSET_DEPR_CODE", "Select ASSET_DEPR_CODE, ASSET_DEPR_DESC from FATDEPM1")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                'If Absx1.txtFor("ASSET_DESC").Text = "" Then
                '    EMsg &= "You must enter an Asset Description for a New Asset"
                'End If

            Case "Refresh"
                If chkEdit.Checked Then
                    EMsg &= "You may not Refresh while Editing / Scheduling AP Items"
                End If

            Case "Edit", "View"
                Validate_Code("ASSET_NO")

                If EMsg = "" Then
                    ASSET_NO = Absx1.txtFor("ASSET_NO").Text

                    If chkEdit.Checked Then
                        EMsg &= "Finish Edits to AP Distributions grid before calling up Specific Assets to View/Edit"
                    End If

                    If eItemKey = "View" Then
                        rowFATFAMF1 = LookUp("FATFAMF1", New String() {ASSET_NO})
                        If rowFATFAMF1 Is Nothing Then
                            EMsg &= "No record on file for Fixed Asset " & ASSET_NO
                        End If
                    ElseIf eItemKey = "Edit" Then
                        If rowFATFAMF1.Item("ASSET_STATUS") = "X" Then
                            EMsg &= $"Fixed Asset {ASSET_NO} has already been Disposed"
                        End If

                    End If
                End If

                'If EMsg = "" And eItemKey = "Edit" Then

                '    If rowFATFAMF1.Item("ITEM_STATUS") & "" <> "A" Then
                '        EMsg &= vbCr & "Item Status is not Active"
                '    End If
                'End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("FATFAMF1", Absx1.txtFor("ASSET_NO").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"

                If Absx1.optFor("ASSET_ACTION").Value = "N" Then
                    If Val(Absx1.numFor("ASSET_BAL").Value & "") <> 0 Then
                        EMsg &= vbCr & "'No Action' is not a valid Action while there is a non-zero Asset Balance"
                    End If
                End If
                'If Format(Absx1.dteFor("ASSET_DATE_IN_SERVICE").Value, "yyyyMMdd") < Format(Absx1.dteFor("ASSET_DATE").Value, "yyyyMMdd") Then
                '    EMsg &= vbCr & "Date In Service may not be prior to Date Added (Capitalized)"
                'End If
                'If Absx1.txtFor("OPS_YYYYPP_IN_SERVICE").Text < Absx1.txtFor("OPS_YYYYPP").Text Then
                '    EMsg &= vbCr & "YP In Service may not be prior to YP Capitalized"
                'End If

                If Absx1.txtFor("OPS_YYYYPP_IN_SERVICE_NEW").Text & "" <> "" Then
                    Dim OPS_YYYPP_IN_SERVICE_NEW As String = Absx1.txtFor("OPS_YYYYPP_IN_SERVICE_NEW").Text
                    Dim OPS_YYYPP_IN_SERVICE As String = Absx1.txtFor("OPS_YYYYPP_IN_SERVICE").Text
                    Dim X As Integer = ASCMAIN1.Period_Diff(OPS_YYYPP_IN_SERVICE, OPS_YYYPP_IN_SERVICE_NEW)
                    If System.Math.Abs(X) > 12 Then
                        EMsg &= vbCr & "New YP In Service may not be more than 12 months away from current YP In Service"
                    End If
                End If


                If Absx1.txtFor("ASSET_CLASS_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify an Asset Class Code"
                Else
                    Dim rowFATFACL1 As DataRow = LookUp("FATFACL1", Absx1.txtFor("ASSET_CLASS_CODE").Text)
                    If rowFATFACL1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Asset Class Code"
                    End If
                End If

                If Absx1.txtFor("ASSET_DEPR_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Depreciation Method Code"
                End If

            Case "Delete"
                If MsgBox("Are you sure that you want to Delete this Asset?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
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

            Case "Refresh"
                Refresh_Records()

            Case "New"
                'EntryMode = "N"
                'Load_Record()
                'Mode_Settings(True)
                add_to_FA_Schedule


            Case "Edit", "View"
                If eItemKey = "View" Then
                    EntryMode = "V"
                Else
                    EntryMode = "E"
                End If
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Export"

                Export_Fixed_Assets_Schedule

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If not_iScreenMode Then
                        .Items("New").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("New").Settings.Enabled = not_iScreenMode
                    End If

                    If EntryMode = "V" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If

                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode

                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode

                    .Items("Refresh").Visible = Not ScreenMode
                    .Items("New").Visible = (Not InquiryMode) And Not ScreenMode
                    .Items("Edit").Visible = (Not InquiryMode)
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")

                    .Items("Update").Visible = ScreenMode And (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = ScreenMode And (Not InquiryMode And EntryMode <> "V")
                    .Items("Delete").Visible = ScreenMode And (Not InquiryMode And EntryMode <> "V")

                    .Items("Export").Visible = Not ScreenMode
                End With

                .Groups("Amounts").Visible = ScreenMode
                .Groups("Fixed Assets Schedule").Visible = False '  Not ScreenMode

            End With
        End If

        'With grdFATFAMFX.DisplayLayout.Override
        '    If InquiryMode Or (EntryMode = "V") Then
        '        .AllowAddNew = UltraWinGrid.AllowAddNew.No
        '        .AllowDelete = DefaultableBoolean.False
        '        .AllowUpdate = DefaultableBoolean.False
        '    Else
        '        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
        '        .AllowDelete = DefaultableBoolean.True
        '        .AllowUpdate = DefaultableBoolean.True
        '    End If

        'End With

        grpAsset.Visible = ScreenMode
        grpAssetAdj.Visible = ScreenMode
        chkNewInService.Visible = ScreenMode And Not InquiryMode And EntryMode = "E"



        Set_Read_Only(grpAMTS, True)
        Set_Read_Only(grpFATFAMF1A, True)
        Set_Read_Only(grpFATFAMF1B, Not (EntryMode = "E"))

        If EntryMode = "E" AndAlso rowFATFAMF1.Item("ASSET_STATUS") = "C" AndAlso Val(rowFATFAMF1.Item("ASSET_BAL") & "") > 0 Then
            Set_Read_Only_for_ctl(Absx1.optFor("ASSET_ACTION"), False)
        Else
            Set_Read_Only_for_ctl(Absx1.optFor("ASSET_ACTION"), True)
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Set_Read_Only(chkNewInService, Not chkNewInService.Visible)
        'Set_Read_Only(grpAssetAdj, Not chkNewInService.Visible)

        'grdFATFAMFX.Visible = Not ScreenMode
        splFA.Visible = Not ScreenMode
        splFATFAMF1.Visible = ScreenMode

        If ScreenMode Then
            Dim depreciation_has_happened As Boolean = (dst.Tables("FATFATR1").Rows.Count > 0)
            Dim Set_to_ReadOnly As Boolean = (InquiryMode Or EntryMode = "V") Or depreciation_has_happened
            Set_Read_Only_for_ctl(Absx1.numFor("ASSET_LIFE_MOS"), Set_to_ReadOnly)
            Set_Read_Only_for_ctl(Absx1.numFor("ASSET_AMT"), Set_to_ReadOnly)
            Set_Read_Only_for_ctl(Absx1.txtFor("ASSET_CLASS_CODE"), Set_to_ReadOnly)

            Set_Read_Only_for_ctl(Absx1.dteFor("ASSET_DATE"), Set_to_ReadOnly)
            Set_Read_Only_for_ctl(Absx1.dteFor("ASSET_DATE_IN_SERVICE"), Set_to_ReadOnly)
            UltraExplorerBar1.Groups("Screen Control").Items("Delete").Visible = Not Set_to_ReadOnly
        Else
            Clear_Record()
            ' Absx1.txtFor("ASSET_DESC").ReadOnly = False
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"FATFAMFX", "FATFAMF1", "FATFATR1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("ASSET_NO").Text = ""

        chkNewInService.Checked = False

        If chkHideDisposed.Tag & "" = "" Then
            chkHideDisposed.Checked = True
            chkHideDisposed.Tag = "H"
        End If

        Refresh_Records()

    End Sub

    Sub Refresh_Records()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Refreshing")

        Fill_Records("FATFAMFX")
        Filter_Disposed()

        Sort_grdColumns(grdFATFAMFX, "ASSET_NO".ToLower)

        If dst.Tables("FATFAMFX").Rows.Count > 0 Then
            YYYY_FIRST = Val(Mid(dst.Tables("FATFAMFX").Compute("MIN(OPS_YYYYPP)", ""), 1, 4))
            YYYY_LAST = Val(Mid(ASCMAIN1.CYP, 1, 4))
        Else
            YYYY_FIRST = Val(Mid(ASCMAIN1.CYP, 1, 4))
            YYYY_LAST = Val(Mid(ASCMAIN1.CYP, 1, 4))
        End If


        'If cbeSCHED_YYYY_FROM.Items.Count = 0 Then
        '    For YYYY_X As Integer = YYYY_FIRST To YYYY_LAST
        '        Dim YYYY As String = Format(YYYY_X, "0000")
        '        cbeSCHED_YYYY_FROM.Items.Add(YYYY)
        '        cbeSCHED_YYYY_TO.Items.Add(YYYY)
        '    Next
        '    cbeSCHED_YYYY_FROM.Value = cbeSCHED_YYYY_FROM.Items(0) ' Format(YYYY_LAST, "0000")
        '    cbeSCHED_YYYY_TO.Value = Format(YYYY_LAST, "0000")
        'End If

        Refresh_APTINVHX()

        Fill_Records("ASTAUDTX")
        Sort_grdColumns(grdASTAUDTX, "INIT_DATE".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub


    Sub Refresh_APTINVHX()

        Dim YP_SINCE As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6)
        Fill_Records("APTINVHX", YP_SINCE)
        'Sort_grdColumns(grdAPTINVHX, "VOUCHER_NO,VOUCHER_LNO")
        grdAPTINVHX.Text = "AP Distributions to Fixed Asset GL Accounts since " & YP_SINCE

        For Each rowAPTINVHX As DataRow In dst.Tables("APTINVHX").Select("")
            Dim VOUCHER_NO As String = rowAPTINVHX.Item("VOUCHER_NO")
            Dim VOUCHER_LNO As String = Val(rowAPTINVHX.Item("VOUCHER_LNO") & "")
            Dim ASSET_CLASS_CODE As String = rowAPTINVHX.Item("ASSET_CLASS_CODE")

            Dim ASSET_AMT As Decimal = Val(dst.Tables("FATFAMFX").Compute("SUM(ASSET_AMT)", $"VOUCHER_NO = '{VOUCHER_NO}' AND VOUCHER_LNO = {CStr(VOUCHER_LNO)}") & "")
            rowAPTINVHX.Item("ASSET_AMT") = ASSET_AMT
            rowAPTINVHX.Item("ASSET_CLASS_CODE") = ASSET_CLASS_CODE
        Next

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then
            ASSET_NO = ASCMAIN1.Next_Control_No("FATFAMF1.ASSET_NO")
            Dim ASSET_DESC As String = Absx1.txtFor("ASSET_DESC").Text
            rowFATFAMF1 = dst.Tables("FATFAMF1").NewRow
            rowFATFAMF1.Item("ASSET_NO") = ASSET_NO
            rowFATFAMF1.Item("ASSET_DESC") = ASSET_DESC ' HFs("ASSET_DESC")
            rowFATFAMF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowFATFAMF1.Item("ASSET_AMT") = 0

            rowFATFAMF1.Item("ASSET_STATUS") = "C"
            rowFATFAMF1.Item("ASSET_DATE") = Now.Date
            rowFATFAMF1.Item("ASSET_DATE_IN_SERVICE") = Now.Date
            rowFATFAMF1.Item("INIT_DATE") = DATETIME_STAMP
            rowFATFAMF1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowFATFAMF1.Item("ASSET_ACTION") = "D"
            rowFATFAMF1.Item("OPS_YYYYPP_IN_SERVICE") = ASCMAIN1.CYP
            rowFATFAMF1.Item("ASSET_DEPR_CODE") = "SL"

            dst.Tables("FATFAMF1").Rows.Add(rowFATFAMF1)

            Absx1.txtFor("ASSET_NO").Text = ASSET_NO
            Absx1.txtFor("ASSET_DESC").Text = ASSET_DESC
        Else
            rowFATFAMF1 = Fill_Record("FATFAMF1", ASSET_NO)
        End If



        Fill_Records("FATFATR1", ASSET_NO)
        Sort_grdColumns(grdFATFATR1, "ASSET_TRN_DATE".ToLower)

        Fill_Records("FATFATRX", ASSET_NO)
        Sort_grdColumns(grdFATFATRX, "YEAR", True)

        'Dim ASSET_AMT As String = Val(rowFATFAMF1.Item("ASSET_AMT") & "")
        Dim ASSET_AMT As Decimal = 0
        If rowFATFAMF1 IsNot Nothing Then
            ASSET_AMT = Val(rowFATFAMF1.Item("ASSET_AMT") & "")
        End If

        Dim YEAR As String = Mid(ASCMAIN1.CYP, 1, 4)
        If rowFATFAMF1 IsNot Nothing Then
            YEAR = Mid(rowFATFAMF1.Item("OPS_YYYYPP"), 1, 4)
        End If

        Dim rowFATFATRX As DataRow = dst.Tables("FATFATRX").Rows.Find(New String() {ASSET_NO, YEAR})

        If rowFATFATRX Is Nothing Then
            rowFATFATRX = dst.Tables("FATFATRX").NewRow
            With rowFATFATRX
                .Item("ASSET_NO") = ASSET_NO
                .Item("YEAR") = YEAR
                .Item("BEG_BAL") = ASSET_AMT
            End With
            dst.Tables("FATFATRX").Rows.Add(rowFATFATRX)
        End If
        rowFATFATRX.Item("BEG_BAL") = ASSET_AMT

        Dim END_BAL As Decimal = Val(rowFATFATRX.Item("END_BAL") & "")

        For Each rowFATFATRX In dst.Tables("FATFATRX").Select($"YEAR > '{YEAR}'", "YEAR")
            rowFATFATRX.Item("BEG_BAL") = END_BAL
            END_BAL = Val(rowFATFATRX.Item("END_BAL") & "")
        Next

        Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()
        Synch_TABLE_NAME("FATFAMF1")

        rowFATFAMF1.Item("LAST_DATE") = DATETIME_STAMP
        rowFATFAMF1.Item("LAST_OPER") = ASCMAIN1.USER_ID

        Update_Record_TDA("FATFAMF1")

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        'Stop ' WHEN WOULD WE BE PERMITTING THIS?
        BeginTrans()
        Delete_Records("FATFAMF1")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME & " where ASSET_NO = '" & ASSET_NO & "'")
        Dim rowFATFAMF1 As DataRow = dst.Tables("FATFAMF1").Rows.Find(ASSET_NO)
        rowFATFAMF1.Delete()

        'rowFATFAMF1.Item("ASSET_STATUS") = "X"
        'rowFATFAMF1.Item("OPS_YYYYPP_DISPOSED") = ASCMAIN1.CYP
        'rowFATFAMF1.Item("ASSET_AMT") = 0
        'rowFATFAMF1.Item("ASSET_BAL") = 0
        'rowFATFAMF1.Item("ASSET_ACTION") = "N"

        'rowFATFAMF1.Item("LAST_DATE") = DATETIME_STAMP
        'rowFATFAMF1.Item("LAST_OPER") = ASCMAIN1.USER_ID
        Update_Record_TDA("FATFAMF1")
    End Sub

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "FATFAMF1"
            E.COLUMN_NAME = "ASSET_NO"
            E.CODE_VALUE = Absx1.txtFor("ASSET_NO").Text
            E.DESC_VALUE = "Fixed Asset"
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = (EntryMode <> "E")
        End If

        Return E
    End Function

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("ASSET_NO").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdFATFAMFX, "SSBBBBBBB", "Show Filter", "Show GroupBox" _
                        , "Select All", "De-Select All", "Select Selected", "De-Select Selected" _
                        , "Set Selected to N (No Action)", "Set Selected to D (Depreciate)", "Set Selected to W (Write-Off)")
        'Load_Popup_Menu(grdFATFATRX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdFATFATR1, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdAPTINVHX, "SSBB", "Show Filter", "Show GroupBox", "Voucher Inquiry", "Add to Fixed Asset Schedule")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdFATFAMFX"
                    tlb_btn = DirectCast(tlb_pop.Tools("Select All"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                    tlb_btn = DirectCast(tlb_pop.Tools("De-Select All"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                    tlb_btn = DirectCast(tlb_pop.Tools("Select Selected"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And (grdFATFAMFX.Selected.Rows.Count <> 0) And Not InquiryMode
                    tlb_btn = DirectCast(tlb_pop.Tools("De-Select Selected"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And (grdFATFAMFX.Selected.Rows.Count <> 0) And Not InquiryMode

                    tlb_btn = DirectCast(tlb_pop.Tools("Set Selected to N (No Action)"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                    tlb_btn = DirectCast(tlb_pop.Tools("Set Selected to D (Depreciate)"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode
                    tlb_btn = DirectCast(tlb_pop.Tools("Set Selected to W (Write-Off)"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode

                Case "grdAPTINVHX"
                    tlb_btn = DirectCast(tlb_pop.Tools("Add to Fixed Asset Schedule"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = Not ScreenMode And Not InquiryMode And Not chkEdit.Checked

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsFilteredOut Then
                    Else
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next

            Case "Select Selected", "De-Select Selected"
                If grd.Selected.Rows.Count = 0 Then
                    MsgBox("No Rows Selected")
                Else
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select Selected", "1", "0")
                        grow.Update()
                    Next
                End If


            Case "Set Selected to N (No Action)", "Set Selected to D (Depreciate)", "Set Selected to W (Write-Off)"
                Dim SEL_COUNT As Integer = dst.Tables("FATFAMFX").Select("SEL = '1'").Length

                If SEL_COUNT = 0 Then
                    MsgBox("No Assets Selected")
                End If

                Dim ASSET_ACTION As String = Mid(e.Tool.Key, 17, 1)
                dst.Tables("FATFAMF1").Rows.Clear()
                Dim ASSET_NOs As New List(Of String)
                Dim ASSET_NOs_cannot_change As New List(Of String)
                For Each row As DataRow In dst.Tables("FATFAMFX").Select("SEL = '1'")
                    Dim ASSET_NO As String = row.Item("ASSET_NO")

                    Dim rowFATFAMF1 As DataRow = Fill_Record("FATFAMF1", ASSET_NO, False, False)
                    If rowFATFAMF1.Item("ASSET_STATUS") & "" <> "C" Or Val(rowFATFAMF1.Item("ASSET_BAL") & "") = 0 Then
                        ASSET_NOs_cannot_change.Add(ASSET_NO)
                        rowFATFAMF1.Delete()
                    Else
                        ASSET_NOs.Add(ASSET_NO)
                    End If
                Next

                If ASSET_NOs_cannot_change.Count <> 0 Then
                    If MsgBox(Join(ASSET_NOs_cannot_change.ToArray, ",") & vbCrLf & vbCrLf & "OK to Continue after De-Selecting these Assets?",
                              MsgBoxStyle.YesNo,
                              "The following Assets are not eligible for this Action") = MsgBoxResult.No Then
                        dst.Tables("FATFAMF1").Rows.Clear()
                        Exit Sub
                    Else
                        For Each ASSET_NO In ASSET_NOs_cannot_change
                            Dim rowFATFAMFX As DataRow = dst.Tables("FATFAMFX").Rows.Find(ASSET_NO)
                            rowFATFAMFX.Item("SEL") = "0"
                        Next
                    End If
                End If

                If ASSET_NOs.Count = 0 Then
                    MsgBox("No (eligible) Assets Selected")
                    dst.Tables("FATFAMF1").Rows.Clear()
                    Exit Sub
                End If

                dst.Tables("FATFAMF1").Rows.Clear()
                For Each ASSET_NO In ASSET_NOs
                    If Not ASCMAIN1.Logical_Lock("FATFAMF1", ASSET_NO,, True, True, 1) Then

                        dst.Tables("FATFAMF1").Rows.Clear()
                        Exit Sub
                    Else
                        Dim rowFATFAMF1 As DataRow = Fill_Record("FATFAMF1", ASSET_NO, False, False)
                        rowFATFAMF1.Item("ASSET_ACTION") = ASSET_ACTION
                    End If
                Next

                If MsgBox($"OK to Change the Asset Action" & vbCrLf & $" for the {ASSET_NOs.Count} Selected Assets" & vbCrLf & $" to { Mid(e.Tool.Key, 17)}?",
                      MsgBoxStyle.YesNo,
                      "Verification") = MsgBoxResult.No Then
                    dst.Tables("FATFAMF1").Rows.Clear()
                    Exit Sub
                Else
                    BeginTrans()
                    Update_Record_TDA("FATFAMF1")
                    dst.Tables("FATFAMF1").Rows.Clear()
                    CommitTrans("Actions Updated")

                    Refresh_Records
                End If
            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Inventory Status"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("BM_COMP_ITEM").Value
            '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            '    If rowICTITEM1 IsNot Nothing Then
            '        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Value
                Dim rowAPTINVH1 As DataRow = LookUp("APTINVH1", VOUCHER_NO)
                If rowAPTINVH1 IsNot Nothing Then
                    Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                End If

            Case "Add to Fixed Asset Schedule"

                If grd.ActiveRow.IsFilterRow Then Exit Sub
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Value
                Dim VOUCHER_LNO As Integer = Val(grd.ActiveRow.Cells("VOUCHER_LNO").Value & "")

                Dim rowAPTINVHX As DataRow = dst.Tables("APTINVHX").Rows.Find(New Object() {VOUCHER_NO, VOUCHER_LNO})

                If Val(rowAPTINVHX.Item("ASSET_REQ") & "") = 0 Then
                    MsgBox("Nothing needs to be scheduled for this record", MsgBoxStyle.OkOnly, "Cannot Add 0 to Fixed Asset Schedule")
                    Exit Sub
                End If

                Using frm As New TAC.FAFFAMFA

                    frm.rowAPTINVHX = rowAPTINVHX
                    frm.ASSET_AMT = Val(rowAPTINVHX.Item("ASSET_REQ") & "")
                    frm.ShowDialog()

                    If frm.ASSET_NO <> "" Then
                        Refresh_Records()
                    End If
                End Using

        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ASSET_DATE"
                If ScreenMode And EntryMode = "E" Then
                    Dim ASSET_DATE As Date = Absx1.dteFor("ASSET_DATE").Value
                    Dim OPS_YYYYPP As String = Format(ASSET_DATE, "yyyyMM")
                    Absx1.txtFor("OPS_YYYYPP").Text = OPS_YYYYPP
                End If
            Case "ASSET_DATE_IN_SERVICE"
                If ScreenMode And EntryMode = "E" Then
                    Dim ASSET_DATE_IN_SERVICE As Date = Absx1.dteFor("ASSET_DATE_IN_SERVICE").Value
                    Dim OPS_YYYYPP_IN_SERVICE As String = Format(ASSET_DATE_IN_SERVICE, "yyyyMM")
                    Absx1.txtFor("OPS_YYYYPP_IN_SERVICE").Text = OPS_YYYYPP_IN_SERVICE
                End If
            Case "ASSET_DATE_IN_SERVICE_NEW"
                If ScreenMode And EntryMode = "E" Then
                    If Absx1.dteFor("ASSET_DATE_IN_SERVICE_NEW").Value & "" = "" Then
                        Absx1.dteFor("ASSET_DATE_IN_SERVICE_NEW").Value = DBNull.Value
                        Absx1.txtFor("OPS_YYYYPP_IN_SERVICE_NEW").Text = ""
                        Absx1.numFor("IN_SERVICE_MOS_ADJ").Value = 0
                    Else
                        Dim OPS_YYYYPP_IN_SERVICE As String = Absx1.txtFor("OPS_YYYYPP_IN_SERVICE").Text
                        Dim ASSET_DATE_IN_SERVICE_NEW As Date = Absx1.dteFor("ASSET_DATE_IN_SERVICE_NEW").Value
                        Dim OPS_YYYYPP_IN_SERVICE_NEW As String = Format(ASSET_DATE_IN_SERVICE_NEW, "yyyyMM")
                        Absx1.txtFor("OPS_YYYYPP_IN_SERVICE_NEW").Text = OPS_YYYYPP_IN_SERVICE_NEW
                        Absx1.numFor("IN_SERVICE_MOS_ADJ").Value = ASCMAIN1.Period_Diff(OPS_YYYYPP_IN_SERVICE_NEW, OPS_YYYYPP_IN_SERVICE)
                    End If
                End If
        End Select
    End Sub
    Public Overrides Sub num_ValueChanged(sender As Object, e As EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ASSET_AMT"
                If ScreenMode And EntryMode = "E" Then
                    Dim ASSET_AMT As Decimal = Val(Absx1.numFor("ASSET_AMT").Value & "")
                    Dim ASSET_BAL As Decimal = ASSET_AMT
                    Absx1.numFor("ASSET_BAL").Value = ASSET_BAL
                End If

            Case "IN_SERVICE_MOS_ADJ"
                If ScreenMode And EntryMode = "E" Then
                    Dim IN_SERVICE_MOS_ADJ As Integer = Val(Absx1.numFor("IN_SERVICE_MOS_ADJ").Value & "")
                    If IN_SERVICE_MOS_ADJ = 0 Then
                        Absx1.dteFor("ASSET_DATE_IN_SERVICE_NEW").Value = DBNull.Value
                        Absx1.txtFor("OPS_YYYYPP_IN_SERVICE_NEW").Text = ""
                    Else
                        If Absx1.dteFor("ASSET_DATE_IN_SERVICE").Value & "" = "" Then
                        Else
                            Dim ASSET_DATE_IN_SERVICE As Date = Absx1.dteFor("ASSET_DATE_IN_SERVICE").Value
                            Absx1.dteFor("ASSET_DATE_IN_SERVICE_NEW").Value = ASSET_DATE_IN_SERVICE.AddMonths(-1 * IN_SERVICE_MOS_ADJ)
                        End If

                    End If
                End If

        End Select
    End Sub

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ASSET_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
                End If
        End Select
    End Sub

    Public Overrides Sub txt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "ASSET_CLASS_CODE"
                If EntryMode = "N" Then
                    Get_Life_from_Class()
                End If
        End Select
    End Sub


    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ASSET_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "ASSET_NO"
                If EntryMode = "" Then
                    If Absx1.txtFor("ASSET_NO").Text <> "" Then
                        LookUp("FATFAMF1", Absx1.txtFor("ASSET_NO").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub

#End Region

    Private Sub grdFATFAMFX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdFATFAMFX.DoubleClickRow
        If e.Row Is Nothing OrElse e.Row.IsFilterRow OrElse e.Row.IsAddRow OrElse e.Row.IsSummaryRow Then Exit Sub

        Absx1.txtFor("ASSET_NO").Text = e.Row.Cells("ASSET_NO").Value
        If InquiryMode Then
            Click_Command("View")
        Else
            Click_Command("View")
        End If
    End Sub

    Sub Display_Totals()
        'Dim rowICTCOSTS As DataRow = dst.Tables("ICTCOSTS").Rows.Find("2")
        'With dst.Tables("BMTMAIN3")
        '    rowICTCOSTS.Item("VCOST") = Val(.Compute("SUM(VCOST)", "") & "")
        '    rowICTCOSTS.Item("LANDG") = Val(.Compute("SUM(LANDG)", "") & "")
        '    rowICTCOSTS.Item("TOOLG") = Val(.Compute("SUM(TOOLG)", "") & "")
        '    rowICTCOSTS.Item("OVRHD") = Val(.Compute("SUM(OVRHD)", "") & "")
        '    rowICTCOSTS.Item("TOTAL") = Val(.Compute("SUM(TOTAL)", "") & "")

        'End With
    End Sub

    Public Overrides Function CustomStringSummary_End(
        ByVal summarySettings As UltraWinGrid.SummarySettings,
        ByVal rows As UltraWinGrid.RowsCollection,
        ByVal CustomValue As String,
        ByVal grd As UltraWinGrid.UltraGrid) As String

        Select Case grd.Name
            Case "grdICTCOSTS"
                Dim KEY As String = summarySettings.Key
                CustomValue = "Totals"
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Private Sub grdFATFATRX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdFATFATRX.InitializeLayout

    End Sub

    Private Sub grdFATFATRX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdFATFATRX.DoubleClickRow
        If e.Row.IsDataRow Then
            Show_Filter(grdFATFATR1, True)

            Dim YEAR As String = e.Row.Cells("YEAR").Value
            grdFATFATR1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            grdFATFATR1.DisplayLayout.Bands(0).ColumnFilters("OPS_YYYYPP_IN_SERVICE").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.StartsWith, YEAR)

            Sort_grdColumns(grdFATFATR1, "OPS_YYYYPP_IN_SERVICE")
        End If

    End Sub

    Private Sub grdFATFAMFX_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdFATFAMFX.InitializeRow
        If e.Row.IsDataRow Then
            Dim ASSET_ACTION As String = e.Row.Cells("ASSET_ACTION").Value
            If ASSET_ACTION = "W" Then
                e.Row.Cells("ASSET_ACTION").Appearance.ForeColor = System.Drawing.Color.Red
            ElseIf ASSET_ACTION = "N" Then
                e.Row.Cells("ASSET_ACTION").Appearance.ForeColor = System.Drawing.Color.Gray
            End If
            Dim ASSET_STATUS As String = e.Row.Cells("ASSET_STATUS").Value
            If ASSET_STATUS = "D" Then
                e.Row.Cells("ASSET_STATUS").Appearance.ForeColor = System.Drawing.Color.Red
            End If
        End If
    End Sub

    Sub Export_Fixed_Assets_Schedule()

        If YYYY_LAST < YYYY_FIRST Then
            MsgBox("Year To cannot be prior to Year From", MsgBoxStyle.OkOnly, "Cannot Export")
            Exit Sub
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Exporting Fixed Assets Schedule")

        grdFATFAMFX.Rows.ColumnFilters.ClearAllFilters()
        grdFATFAMFX.DisplayLayout.Bands(0).SortedColumns.Clear()

        Sort_grdColumns(grdFATFAMFX, "ASSET_NO")

        Dim RT As Integer = grdFATFAMFX.Rows.Count

        With grdFATFAMFX.DisplayLayout.Bands(0)
            For Each col As String In New String() {"SEL", "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}
                .Columns(col).Hidden = True
            Next
        End With

        Dim filename As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No(Me.Name & ".XLS") & ".XLS"
        Dim xlwb As Infragistics.Documents.Excel.Workbook = Export_to_Excel(grdFATFAMFX, False, False, "Fixed Asset Schedule")
        xlwb.Save(filename)

        Dim workbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(filename)
        Dim worksheet As SpreadsheetGear.IWorksheet = workbook.Worksheets(0)
        Dim range As SpreadsheetGear.IRange = Nothing

        Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
        Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

        With workbook.Worksheets(0).Cells
            .Font.Size = 10
            .Font.Name = "Verdana"
        End With

        Dim r As Integer = 4
        Dim c As Integer = 0

        Dim r0 As Integer = r
        worksheet.Cells(r + RT + 1, 0).EntireRow.Delete()
        worksheet.Cells(r + RT + 1, 0).EntireRow.Delete()

        For c = 6 To 9
            worksheet.Cells(r - 1, c).Formula = $"=SUBTOTAL(9,{Excel_Cell0(r + 1, c)}:{ Excel_Cell0(r + RT, c)})"
            worksheet.Cells(r - 1, c).EntireColumn.NumberFormat = "#,##0.00"
        Next

        c = 18
        Dim c0 As Integer = c
        Dim Depreciation_TY As String = ""

        For Y As Integer = YYYY_FIRST To YYYY_LAST
            Dim YYYY As String = Format(Y, "0000")
            'c += 1
            Dim sqlX As String = "Select ASSET_NO" & vbCrLf
            For M As Integer = 1 To 12
                c += 1
                Dim d As Date = CDate($"{CStr(M)}/01/{CStr(Y)}")
                Dim dt As String = Format(d, "MM/dd/yyyy") ' & Format(d, "MMM-yy")
                worksheet.Cells(r, c).Value = dt
                worksheet.Cells(r - 1, c).Formula = $"=SUBTOTAL(9,{ Excel_Cell0(r + 1, c)}:{Excel_Cell0(r + RT, c)})"
                sqlX &= $", SUM (DECODE(SUBSTR(OPS_YYYYPP,5,2),'{Format(M, "00")}',ASSET_TRN_AMT,NULL)) M{Format(M, "00")}" & vbCrLf
            Next
            sqlX &= $" from FATFATR1 where (ASSET_TRN_TYPE = 'EXP' OR ASSET_TRN_TYPE = 'WOF') and OPS_YYYYPP LIKE '{Format(Y, "0000")}%' group by ASSET_NO"
            'Dim sqlD As String = "Select FATFAMF1.ASSET_NO, X.M01, X.M02, X.M03, X.M04, X.M05, X.M06, X.M07, X.M08, X.M09, X.M10, X.M11, X.M12" & vbCrLf
            Dim sqlD As String = "Select X.M01, X.M02, X.M03, X.M04, X.M05, X.M06, X.M07, X.M08, X.M09, X.M10, X.M11, X.M12" & vbCrLf
            sqlD &= $"from FATFAMF1" & vbCrLf
            sqlD &= $", ({sqlX}) X" & vbCrLf
            sqlD &= $"where X.ASSET_NO (+) = FATFAMF1.ASSET_NO" & vbCrLf
            If chkHideDisposed.Checked Then
                sqlD &= $"and FATFAMF1.ASSET_STATUS <> 'X'" & vbCrLf
            End If
            sqlD &= " order by FATFAMF1.ASSET_NO"

            c += 1
            worksheet.Cells(r, c).Value = "Net Book Value " & Format(Y, "0000")
            worksheet.Cells(r - 1, c).Formula = $"=SUBTOTAL(9,{Excel_Cell0(r + 1, c)}:{Excel_Cell0(r + RT, c)})"

            Dim c_Init As Integer = 6 ' Asset Value
            Dim XC_Init As String = Excel_Cell0(r + 1, c_Init)
            Dim c_YP As Integer = 14 ' Col for YP in Service
            Dim XC_Year As String = $"Mid({Excel_Cell0(r + 1, c_YP)},1,4)"
            'Dim c_Prev As Integer = c - 13 ' Asset Value Previous Year
            'Dim XC_Prev As String = Excel_Cell0(r + 1, c_Prev)
            'Dim c_TY_Beg As String = $"IF({XC_Year}='{YYYY}',{XC_Init},IF({XC_Year}>'{YYYY}',0,{XC_Prev}))"
            Dim c_TY_Beg As String = $"IF({XC_Year}>'{YYYY}',0,{XC_Init}"
            Dim Depreciation_YYYY As String = $"SUM({Excel_Cell0(r + 1, c - 12)}:{Excel_Cell0(r + 1, c - 1)})"
            Depreciation_TY &= "-" & Depreciation_YYYY
            Dim f As String = $"={c_TY_Beg}{Depreciation_TY})"
            worksheet.Cells(r + 1, c).Formula = Replace(f, "'", Chr(34))

            '=IF(MID(N6,1,4)="2014",G6,IF(MID(N6,1,4)>"2014",0,S6))-SUM(T6:AE6)

            rangeCopyFrom = worksheet.Range(r + 1, c)
            rangePaste_To = worksheet.Range(r + 1, c, r + RT, c)
            rangeCopyFrom.Copy(rangePaste_To, SpreadsheetGear.PasteType.Formulas, SpreadsheetGear.PasteOperation.None, False, False)


            worksheet.Cells(0, c - 12, 0, c - 1).EntireColumn.Group()

            Dim tbl As DataTable = ASCDATA1.GetDataTable(sqlD)
            worksheet.Cells(r0 + 1, c - 12, r0 + RT, c - 1).CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)
        Next

        With worksheet.Cells(0, c0 + 1, 0, c)
            .ColumnWidth = 12
        End With

        With worksheet.Cells(r - 1, c0 + 1, r - 1, c)
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
        End With

        With worksheet.Cells(r, 0, r, c0)
            .Interior.Color = SpreadsheetGear.Colors.LightGray
        End With

        With worksheet.Cells(r, c0 + 1, r, c)
            .HorizontalAlignment = SpreadsheetGear.HAlign.Center
            .Interior.Color = SpreadsheetGear.Colors.LightGray
            .NumberFormat = "MMM'yy"
            .VerticalAlignment = SpreadsheetGear.VAlign.Center
        End With



        'worksheet.Names.Add("Data", "=" & Excel_Cell0(r0 + 1, 0) & ":" & Excel_Cell0(r0 + RT, c))
        workbook.Names.Add("Data", "='Fixed Asset Schedule'!" & Excel_Cell0(r0, 0, 3) & ":" & Excel_Cell0(r0 + RT, c, 3))
        '='Fixed Asset Schedule'!$A$5:$ES$346
        With worksheet.Cells(r, 0, r, c)
            .AutoFilter()
        End With

        With worksheet.Cells(r - 1, c0 + 1, r - 1, c)
            .NumberFormat = "#,##0.00;#,##0.00;" ' "#,##0.00"
        End With

        With worksheet.Cells(r + 1, c0 + 1, r + RT, c)
            .NumberFormat = "#,##0.00;#,##0.00;" ' "#,##0.00"
        End With

        worksheet.Outline.ShowLevels(0, 1)


        worksheet.Cells(r0 + 1, 3).Activate()
        worksheet.WindowInfo.FreezePanes = True

        workbook.Save()

        With grdFATFAMFX.DisplayLayout.Bands(0)
            For Each col As String In New String() {"SEL", "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}
                .Columns(col).Hidden = False
            Next
        End With

        workbook.Close()
        workbook = Nothing
        Show_Document(filename)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub chkEdit_CheckedChanged(sender As Object, e As EventArgs) Handles chkEdit.CheckedChanged
        If chkEdit.Checked Then
            If Not ASCMAIN1.Logical_Lock("FATFAMF2", "*",,,, 1) Then
                chkEdit.Checked = False
                Exit Sub
            End If

            Refresh_APTINVHX()

            Dim dvw As DataView = DirectCast(grdAPTINVHX.DataSource, DataTable).DefaultView
            dvw.RowFilter = "IGNORE = '1' or ASSET_REQ <> 0"

            grdAPTINVHX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdAPTINVHX.DisplayLayout.Bands(0).Columns("IGNORE").CellAppearance.BackColor = System.Drawing.Color.LightBlue

            dst.Tables("FATFAMF2").Rows.Clear()

            For Each rowAPTINVHX As DataRow In dst.Tables("APTINVHX").Select("")
                Dim VOUCHER_NO As String = rowAPTINVHX.Item("VOUCHER_NO")
                Dim VOUCHER_LNO As Integer = Val(rowAPTINVHX("VOUCHER_LNO") & "")
                Fill_Record("FATFAMF2", New Object() {VOUCHER_NO, VOUCHER_LNO}, False, False)
            Next

            chkHideIgnored.Checked = False
        Else

            BeginTrans()
            Update_Record_TDA("FATFAMF2")
            CommitTrans()

            ASCMAIN1.MultiTask_Release(,, 1)
            Dim dvw As DataView = DirectCast(grdAPTINVHX.DataSource, DataTable).DefaultView
            dvw.RowFilter = ""

            grdAPTINVHX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdAPTINVHX.DisplayLayout.Bands(0).Columns("IGNORE").CellAppearance.BackColor = System.Drawing.Color.Empty
        End If
    End Sub


    Private Sub grdAPTINVHX_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdAPTINVHX.AfterRowUpdate

    End Sub

    Private Sub grdFATFAMFX_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdFATFAMFX.InitializeLayout

    End Sub

    Private Sub grdFATFAMFX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdFATFAMFX.AfterRowActivate
        If grdFATFAMFX.ActiveRow IsNot Nothing AndAlso grdFATFAMFX.ActiveRow.IsDataRow And Not grdFATFAMFX.ActiveRow.IsFilterRow Then
            If tabFA.SelectedTab.Key = "Fixed Assets Audit Trail" Then
                Filter_to_Asset(grdFATFAMFX.ActiveRow.Cells("ASSET_NO").Value)
            End If
        End If

    End Sub

    Sub Filter_to_Asset(ASSET_NO As String)
        grdASTAUDTX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdASTAUDTX.DisplayLayout.Bands(0).ColumnFilters("KEY_VALUE").FilterConditions.Add(UltraWinGrid.FilterComparisionOperator.Equals, ASSET_NO)
    End Sub

    Private Sub tabFA_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabFA.SelectedTabChanged
        If Me.SELECTION_NO = 0 Then Exit Sub

        If tabFA.SelectedTab.Key = "Fixed Assets Audit Trail" Then
            If grdFATFAMFX.ActiveRow IsNot Nothing AndAlso grdFATFAMFX.ActiveRow.IsDataRow And Not grdFATFAMFX.ActiveRow.IsFilterRow Then
                Filter_to_Asset(grdFATFAMFX.ActiveRow.Cells("ASSET_NO").Value)
            Else
                grdASTAUDTX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
            End If
        End If
    End Sub

    Private Sub grdAPTINVHX_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdAPTINVHX.BeforeRowUpdate

        Dim VOUCHER_NO As String = e.Row.Cells("VOUCHER_NO").Value 'grdAPTINVHX.ActiveRow.Cells("VOUCHER_NO").Value '
        Dim VOUCHER_LNO As Integer = Val(e.Row.Cells("VOUCHER_LNO").Value & "") ' Val(grdAPTINVHX.ActiveRow.Cells("VOUCHER_LNO").Value & "") ' 
        If VOUCHER_NO = "" Then Exit Sub ' WHY IS THIS HAPPENING

        Dim rowFATFAMF2 As DataRow = dst.Tables("FATFAMF2").Rows.Find({VOUCHER_NO, VOUCHER_LNO})
        If rowFATFAMF2 Is Nothing Then
            rowFATFAMF2 = dst.Tables("FATFAMF2").NewRow
            rowFATFAMF2.Item("VOUCHER_NO") = VOUCHER_NO
            rowFATFAMF2.Item("VOUCHER_LNO") = VOUCHER_LNO
            dst.Tables("FATFAMF2").Rows.Add(rowFATFAMF2)
        End If
        If e.Row.Cells("IGNORE").Text = "Checked" Or e.Row.Cells("IGNORE").Text = "1" Then
            rowFATFAMF2.Item("IGNORE") = "1"
        ElseIf e.Row.Cells("IGNORE").Text = "UnChecked" Or e.Row.Cells("IGNORE").Text = "0" Then
            rowFATFAMF2.Item("IGNORE") = "0"
        Else
            rowFATFAMF2.Item("IGNORE") = e.Row.Cells("IGNORE").Text
        End If


    End Sub

    Private Sub chkHideIgnored_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideIgnored.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Filter_Ignored()
    End Sub

    Sub Filter_Ignored()
        Dim dvw As DataView = dst.Tables("APTINVHX").DefaultView
        If chkHideIgnored.Checked Then
            dvw.RowFilter = "ISNULL(IGNORE,'0')='0'"
        Else
            dvw.RowFilter = ""
        End If
    End Sub

    Private Sub chkHideDisposed_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideDisposed.CheckedChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Filter_Disposed()
    End Sub

    Sub Filter_Disposed()
        Dim dvw As DataView = DirectCast(grdFATFAMFX.DataSource, DataTable).DefaultView
        If chkHideDisposed.Checked Then
            dvw.RowFilter = "ASSET_STATUS <> 'X'"
        Else
            dvw.RowFilter = ""
        End If
    End Sub

    Private Sub chkNewInService_CheckedChanged(sender As Object, e As EventArgs) Handles chkNewInService.CheckedChanged

        grpAssetAdj.Enabled = chkNewInService.Checked
        Set_Read_Only(grpAssetAdj, Not chkNewInService.Checked)
    End Sub

    Function Get_Life_from_Class() As Integer
        Dim ASSET_CLASS_CODE As String = Absx1.txtFor("ASSET_CLASS_CODE").Text
        Dim ASSET_LIFE_MOS As Integer

        Dim rowFATFACL1 As DataRow = LookUp("FATFACL1", ASSET_CLASS_CODE)
        If rowFATFACL1 IsNot Nothing Then
            ASSET_LIFE_MOS = Val(rowFATFACL1.Item("ASSET_LIFE_MOS") & "")
            Absx1.numFor("ASSET_LIFE_MOS").Value = ASSET_LIFE_MOS
        End If

        Return ASSET_LIFE_MOS
    End Function


    Sub Add_to_FA_Schedule()

        Using frm As New TAC.FAFFAMFA

            frm.rowAPTINVHX = Nothing
            frm.ASSET_AMT = 0

            'ASSET_NO = ASCMAIN1.Next_Control_No("FATFAMF1.ASSET_NO")
            'Dim ASSET_DESC As String = Absx1.txtFor("ASSET_DESC").Text
            'rowFATFAMF1 = dst.Tables("FATFAMF1").NewRow
            'rowFATFAMF1.Item("ASSET_NO") = ASSET_NO
            'rowFATFAMF1.Item("ASSET_DESC") = ASSET_DESC ' HFs("ASSET_DESC")
            'rowFATFAMF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            'rowFATFAMF1.Item("ASSET_AMT") = 0

            'rowFATFAMF1.Item("ASSET_STATUS") = "C"
            'rowFATFAMF1.Item("ASSET_DATE") = Now.Date
            'rowFATFAMF1.Item("ASSET_DATE_IN_SERVICE") = Now.Date
            'rowFATFAMF1.Item("INIT_DATE") = DATETIME_STAMP
            'rowFATFAMF1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            'rowFATFAMF1.Item("ASSET_ACTION") = "D"
            'rowFATFAMF1.Item("OPS_YYYYPP_IN_SERVICE") = ASCMAIN1.CYP
            'rowFATFAMF1.Item("ASSET_DEPR_CODE") = "SL"

            'dst.Tables("FATFAMF1").Rows.Add(rowFATFAMF1)

            'Absx1.txtFor("ASSET_NO").Text = ASSET_NO
            'Absx1.txtFor("ASSET_DESC").Text = ASSET_DESC


            frm.ShowDialog()

            If frm.ASSET_NO <> "" Then
                Refresh_Records()
            End If
        End Using
    End Sub
End Class