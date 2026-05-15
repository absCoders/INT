Imports Infragistics.Win.UltraWinGrid
Public Class WHFROUT1
    Dim SHIP_TO_TYPE As String
    Dim SHIP_TO_KEY As String
    Dim SHIP_TO_CODE_PFX As String
    Dim sqlSOTSHIPG As String
    Dim sqlSOTSHIPY As String
    Dim rowSOTSHIPE As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SOTSHIPG.*" & vbCrLf _
                & ", DECODE(SOTSHIPD.SHIP_TO_TYPE,'C',ARTCUST1.CUST_NAME,'W',ICTWHSE1.WHSE_DESC,'S',ARTCUST2.CUST_STORE_NAME,'V',APTVEND1.VEND_NAME,'?') SHIP_TO_NAME" & vbCrLf _
                & " from SOTSHIPG,SOTSHIPD,SOTSHIPE,ARTCUST1,ARTCUST2,APTVEND1,ICTWHSE1" & vbCrLf _
                & " where NVL(SOTSHIPG.ROUTING_RULE_STATUS,'?') = :PARM1" & vbCrLf _
                & "   and (:PARM2 = '*' or SOTSHIPE.SHIP_TO_CODE_PFX = :PARM2)" & vbCrLf _
                & "   and SOTSHIPD.SHIP_TO_CODE (+) = SOTSHIPG.SHIP_TO_CODE" & vbCrLf _
                & "   and SOTSHIPD.SHIP_TO_CODE_PFX = SOTSHIPE.SHIP_TO_CODE_PFX" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE (+) = SOTSHIPD.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE (+) = SOTSHIPD.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO (+) = SOTSHIPD.CUST_STORE_NO" & vbCrLf _
                & "   and APTVEND1.VEND_CODE (+) = SOTSHIPD.VEND_CODE" & vbCrLf _
                & "   and ICTWHSE1.WHSE_CODE (+) = SOTSHIPD.WHSE_CODE"

            sqlSOTSHIPG = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTSHIPG", "**", 0, True, "VV", 1)

            ASCMAIN1.sql = "Select SOTSHIPY.*" & vbCrLf _
                & ", SOTSHIPS.DEL_METHOD, SOTSHIPS.SHIP_VIA_DESC, SOTSHIPS.SHIP_VIA_SCAC, SOTSHIPS.CARRIER, SOTSHIPS.SHIP_METHOD" & vbCrLf _
                & " from SOTSHIPG,SOTSHIPY,SOTSHIPD, SOTSHIPE,SOTSHIPS" & vbCrLf _
                & " where NVL(SOTSHIPG.ROUTING_RULE_STATUS,'?') = :PARM1" & vbCrLf _
                & "   and (:PARM2 = '*' or SOTSHIPE.SHIP_TO_CODE_PFX = :PARM2)" & vbCrLf _
                & "   and SOTSHIPD.SHIP_TO_CODE (+) = SOTSHIPG.SHIP_TO_CODE" & vbCrLf _
                & "   and SOTSHIPD.SHIP_TO_CODE_PFX = SOTSHIPE.SHIP_TO_CODE_PFX" & vbCrLf _
                & "   and SOTSHIPG.ROUTING_RULE_NO = SOTSHIPY.ROUTING_RULE_NO" & vbCrLf _
                & "   and SOTSHIPS.SHIP_VIA (+) = SOTSHIPY.SHIP_VIA"

            sqlSOTSHIPY = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTSHIPY", "**", 0, True, "VV", 0)

            Create_Relation("SOTSHIPG", "SOTSHIPY", "ROUTING_RULE_NO")

            ASCMAIN1.sql = sqlSOTSHIPG
            Create_TDA(.Tables.Add, "SOTSHIPG_INACTIVE", "**", 0, False, "VV", 1)

            ASCMAIN1.sql = sqlSOTSHIPY
            Create_TDA(.Tables.Add, "SOTSHIPY_INACTIVE", "**", 0, False, "VV", 0)

            Create_Relation("SOTSHIPG_INACTIVE", "SOTSHIPY_INACTIVE", "ROUTING_RULE_NO")

            ASCMAIN1.sql = "Select ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_CITY, ARTCUST2.CUST_STORE_STATE" & vbCrLf _
                & ", ARTCUST2.CUST_DC_IND, ARTCUST2.CUST_DC_NO, SOTSHIPD.SHIP_TO_CODE" & vbCrLf _
                & " from ARTCUST2, SOTSHIPD" & vbCrLf _
                & " where ARTCUST2.CUST_CODE = :PARM1" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SOTSHIPD.CUST_CODE (+)" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SOTSHIPD.CUST_STORE_NO (+)"
            Create_TDA(.Tables.Add, "ARTCUST2_SHIPTO", "**", 0, False, "V", 2)
            With dst.Tables("ARTCUST2_SHIPTO")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

            Create_TDA(.Tables.Add, "SOTSHIPE", "*") ' PROBABLY NEED A GRID FOR THIS ONE TOO

            ASCMAIN1.sql = "Select SOTSHIPD.*" & vbCrLf _
                & ", DECODE(SOTSHIPD.SHIP_TO_TYPE,'C',ARTCUST1.CUST_NAME,'W',ICTWHSE1.WHSE_DESC,'S',ARTCUST2.CUST_STORE_NAME,'V',APTVEND1.VEND_NAME,'?') SHIP_TO_NAME" & vbCrLf _
                & " from SOTSHIPD,ARTCUST1,ICTWHSE1,APTVEND1,ARTCUST2" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = SOTSHIPD.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE (+) = SOTSHIPD.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO (+) = SOTSHIPD.CUST_STORE_NO" & vbCrLf _
                & "   and APTVEND1.VEND_CODE (+) = SOTSHIPD.VEND_CODE" & vbCrLf _
                & "   and ICTWHSE1.WHSE_CODE (+) = SOTSHIPD.WHSE_CODE"
            Create_TDA(.Tables.Add, "SOTSHIPD", "*", 0, False)

            ASCMAIN1.sql = "Select SOTSHIPF.*" & vbCrLf _
                & " from SOTSHIPF" & vbCrLf
            Create_TDA(.Tables.Add, "SOTSHIPF", "*", 0, False)

            Create_Relation("SOTSHIPD", "SOTSHIPF", "SHIP_TO_CODE")

            Create_TDA(.Tables.Add, "SOTSHIPU", "*", 0, False)
            Fill_Records("SOTSHIPU")


            ASCMAIN1.sql = "SELECT SHIP_TO_CODE, SHIPTO1, SHIPTO2
            , SUM (CASE WHEN FR_CTN = 1 THEN 1 ELSE 0 END) CTN1
            , SUM (CASE WHEN FR_CTN > 0 THEN 1 ELSE 0 END) CTN_RNG
            , SUM (CASE WHEN FR_CTN = 0 THEN 1 ELSE 0 END) LBS_RNG
            FROM SOTSHIPG_JIC WHERE SHIP_TO_CODE IN (SELECT DISTINCT SHIP_TO_CODE FROM SOTSHIPG WHERE FR_CTN > 0)
            GROUP BY SHIP_TO_CODE, SHIPTO1, SHIPTO2"
            ASCMAIN1.sql = $"Select X.*, SOTSHIPD.CUST_CODE, SOTSHIPD.CUST_STORE_NO, ARTCUST2.CUST_STORE_NAME
            from ARTCUST2, SOTSHIPD, ({ASCMAIN1.sql}) X where SOTSHIPD.SHIP_TO_CODE (+) = X.SHIP_TO_CODE 
                and ARTCUST2.CUST_CODE (+) = SOTSHIPD.CUST_CODE and ARTCUST2.CUST_STORE_NO (+) = SOTSHIPD.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SOTSHIPG_CTN", "**", 0, False)
            Fill_Records("SOTSHIPG_CTN")

            ASCMAIN1.sql = "SELECT SOTSHIPG.*
            FROM SOTSHIPG_JIC SOTSHIPG WHERE SHIP_TO_CODE IN (SELECT DISTINCT SHIP_TO_CODE FROM SOTSHIPG_JIC WHERE FR_CTN > 0)"
            Create_TDA(.Tables.Add, "SOTSHIPG_ALL", "**", 0, False)
            Fill_Records("SOTSHIPG_ALL")

            Create_Relation("SOTSHIPG_CTN", "SOTSHIPG_ALL", "SHIP_TO_CODE")

        End With

        grdSOTSHIPU.DataSource = dst.Tables("SOTSHIPU")
        grdSOTSHIPD.DataSource = dst.Tables("SOTSHIPD")
        'Show_Filter(grdSOTSHIPD, True)

        grdSOTSHIPG_CTN.DataSource = dst.Tables("SOTSHIPG_CTN")
        Show_Filter(grdSOTSHIPG_CTN, True)
        Sort_grdColumns(grdSOTSHIPG_CTN, "SHIPTO1")
        Sort_grdColumns(grdSOTSHIPG_CTN, "ROUTING_RULE_NO",, 1)

        grdSOTSHIPG.DataSource = dst.Tables("SOTSHIPG")
        Create_Summary(grdSOTSHIPG, "ROUTING_RULE_NO", "Count")
        Show_Filter(grdSOTSHIPG, False)

        grdSOTSHIPG_INACTIVE.DataSource = dst.Tables("SOTSHIPG_INACTIVE")
        Create_Summary(grdSOTSHIPG_INACTIVE, "ROUTING_RULE_NO", "Count")
        Show_Filter(grdSOTSHIPG_INACTIVE)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTSHIPG, grdSOTSHIPG_INACTIVE}

            With grd.DisplayLayout.Bands(0)

                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    gcol.CellActivation = Activation.NoEdit
                    If grd.Name = "grdSOTSHIPG" And gcol.Key <> "SOTSHIPG_SOTSHIPY" Then
                        With grdSOTSHIPG_INACTIVE.DisplayLayout.Bands(0).Columns(gcol.Key)
                            .Header.Caption = gcol.Header.Caption
                            .Width = gcol.Width
                            .Style = gcol.Style
                            .Format = gcol.Format
                            .CellAppearance.TextHAlign = gcol.CellAppearance.TextHAlign
                            .Header.Appearance.TextHAlign = gcol.Header.Appearance.TextHAlign
                        End With
                    End If
                Next

                Dim g As New UltraWinGrid.UltraGridGroup

                g = .Groups.Add("RECORD_KEYS")
                g.Header.Fixed = True
                g.Header.Caption = "Record Keys"
                For Each c As String In New String() {"ROUTING_RULE_NO", "ROUTING_RULE_STATUS", "SHIP_TO_CODE", "SHIP_TO_NAME"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Fixed = True
                Next

                'g = .Groups.Add("ROUTING_WGTS")
                'g.Header.Caption = "Weights"
                'For Each c As String In New String() {"FR_WGT", "TO_WGT"}
                '    .Columns(c).Group = g
                '    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.Yellow
                '    If grd.Name = "grdSOTSHIPG" Then .Columns(c).CellActivation = Activation.AllowEdit
                'Next
                'g.Hidden = True

                'g = .Groups.Add("ROUTING_RESULTS")
                'g.Header.Caption = "Routing Guidance"
                'For Each c As String In New String() {"DEL_METHOD", "SHIP_VIA", "SHIP_VIA_DESC", "CARRIER", "SHIP_VIA_SCAC", "SHIP_METHOD"}
                '    .Columns(c).Group = g
                '    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                '    If grd.Name = "grdSOTSHIPG" Then .Columns(c).CellActivation = Activation.AllowEdit
                'Next
                'g.Hidden = True

                g = .Groups.Add("AS400 Key")
                g.Header.Caption = "AS400 Ship-To"
                For Each c As String In New String() {"SHIPTO1", "SHIPTO2"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Next
                'g.Hidden = True

                g = .Groups.Add("CHANGES")
                g.Header.Caption = "Dates Changed"
                For Each c As String In New String() {"INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}
                    .Columns(c).Group = g
                    .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.Gold
                Next
                g.Hidden = True

                For Each grp As UltraWinGrid.UltraGridGroup In .Groups
                    grp.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                    grp.Header.Appearance.BackColor = System.Drawing.Color.White
                    grp.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Next

                'ASCMAIN1.Add_Value_List(grd, "SHIP_RULE_CODE", Nothing, New String() {":", "C:Consecutive", "N:Non-Consecutive", "R:Specified Days"})
                'ASCMAIN1.Add_Value_List(grd, "ROUTING_HOL_EXC", Nothing, New String() {":", "E:1 Bus Day Before", "S:Next Sched Day", "L:1 Bus Day Later"})
                'ASCMAIN1.Add_Value_List(grd, "PICK_UP_HOL_EXC", Nothing, New String() {":", "E:1 Bus Day Before", "S:Next Sched Day", "L:1 Bus Day Later"})

                ASCMAIN1.Add_Value_List(grd, "ROUTING_RULE_STATUS", Nothing, New String() {":", "A:Active", "I:Retired"})

            End With

            If grd.Name = "grdSOTSHIPG" Then
                With grd.DisplayLayout.Bands(1)

                    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                        gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                        gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                        gcol.CellActivation = Activation.NoEdit
                        If grd.Name = "grdSOTSHIPG" Then ' And gcol.Key <> "SOTSHIPG_SOTSHIPY" Then
                            'With grdSOTSHIPG_INACTIVE.DisplayLayout.Bands(1).Columns(gcol.Key)
                            '    .Header.Caption = gcol.Header.Caption
                            '    .Width = gcol.Width
                            '    .Style = gcol.Style
                            '    .Format = gcol.Format
                            '    .CellAppearance.TextHAlign = gcol.CellAppearance.TextHAlign
                            '    .Header.Appearance.TextHAlign = gcol.Header.Appearance.TextHAlign
                            'End With
                        End If
                    Next

                    Dim g As New UltraWinGrid.UltraGridGroup

                    g = .Groups.Add("RECORD_KEYS")
                    g.Header.Fixed = True
                    g.Header.Caption = "Record Keys"
                    For Each c As String In New String() {"ROUTING_RULE_NO"}
                        .Columns(c).Group = g
                        .Columns(c).Header.Fixed = True
                    Next
                    g.Hidden = True

                    g = .Groups.Add("ROUTING_WGTS")
                    g.Header.Caption = "Weight Min"
                    For Each c As String In New String() {"FR_WGT"}
                        .Columns(c).Group = g
                        .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.Yellow
                        If grd.Name = "grdSOTSHIPG" Then .Columns(c).CellActivation = Activation.AllowEdit
                    Next

                    g = .Groups.Add("ROUTING_RESULTS")
                    g.Header.Caption = "Routing Guidance"
                    For Each c As String In New String() {"DEL_METHOD", "SHIP_VIA", "SHIP_VIA_DESC", "CARRIER", "SHIP_VIA_SCAC", "SHIP_METHOD"}
                        .Columns(c).Group = g
                        .Columns(c).Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                        If grd.Name = "grdSOTSHIPG" Then .Columns(c).CellActivation = Activation.AllowEdit
                    Next

                    For Each grp As UltraWinGrid.UltraGridGroup In .Groups
                        grp.Header.Appearance.BackColor2 = System.Drawing.Color.LightGray
                        grp.Header.Appearance.BackColor = System.Drawing.Color.White
                        grp.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    Next
                End With
            End If
        Next

        Bind_Controls(UltraGroupBox1, "SOTSHIPE")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                SHIP_TO_KEY = Absx1.txtFor("SHIP_TO_KEY").Text
                SHIP_TO_TYPE = optSHIP_TO_TYPE.Value
                SHIP_TO_CODE_PFX = Absx1.txtFor("SHIP_TO_CODE_PFX").Text

                ' Dim rowSOTSHIPE As DataRow = LookUp("SOTSHIPE", SHIP_TO_CODE_PFX)
                rowSOTSHIPE = Fill_Record("SOTSHIPE", SHIP_TO_CODE_PFX)
                If rowSOTSHIPE Is Nothing Then
                    EMsg &= vbCr & "Invalid Ship-To Code Prefix"
                End If

                If EMsg = "" Then
                    ASCMAIN1.Logical_Lock("SOTSHIPE", SHIP_TO_CODE_PFX)
                End If

            Case "Validate All"
                EMsg = Validate_Active_Records()
                If EMsg = "" Then
                    MsgBox("All Routing Records are Valid", MsgBoxStyle.OkOnly, "Validation Complete")
                Else
                    Using frm As New ASFMSGBF
                        EMsg = Replace(EMsg, vbCr, "</br>")
                        frm.Show_Formatted_txt("Data Validation Issues", EMsg, Me)
                        EMsg = ""
                    End Using
                End If

            Case "Done"

            Case "Update"

                EMsg = Validate_Active_Records()

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
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Validate All").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Visible = tf
            .Groups("Screen Control").Items("Update").Visible = tf
        End With

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        grpLead.Visible = Not ScreenMode
        ' SplitContainer1.Visible = ScreenMode
        grpShipTo.Visible = ScreenMode

        txtSHIP_VIA.Visible = ScreenMode
        lblSHIP_VIA.Visible = ScreenMode
        txtSHIP_VIA_DESC.Visible = ScreenMode
        txtCARRIER_ACCT_TYPE.Visible = ScreenMode
        lblCARRIER_ACCT_TYPE.Visible = ScreenMode

        If Not ScreenMode Then
            Clear_Record()
            'grdSOTSHIPG.Parent = grpLead
            'SplitContainer2.Parent = grpLead
            tabMain.Parent = grpLead
            grdSOTSHIPG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            SplitContainer1.Panel2Collapsed = True

        Else

            'grdSOTSHIPG.Parent = SplitContainer1.Panel1
            'SplitContainer2.Parent = grpShipTo '  UltraTabPageControl2
            tabMain.Parent = grpShipTo
            SplitContainer1.Panel2Collapsed = False
            grdSOTSHIPG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True

            'Dim rowSOTSHIPE As DataRow = LookUp("SOTSHIPE", SHIP_TO_CODE_PFX)

            Set_Read_Only_for_ctl(txtCARRIER_ACCT_TYPE, False)
            Set_Read_Only_for_ctl(txtSHIP_VIA, False)

            Select Case SHIP_TO_TYPE
                Case "C"
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", rowSOTSHIPE.Item("CUST_CODE"))
                    Absx1.txtFor("SHIP_TO_KEY_NAME").Text = rowARTCUST1.Item("CUST_NAME")
                Case "W"
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", rowSOTSHIPE.Item("WHSE_CODE"))
                    Absx1.txtFor("SHIP_TO_KEY_NAME").Text = rowICTWHSE1.Item("WHSE_DESC")
                Case "V"
                    Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1 ", rowSOTSHIPE.Item("VEND_CODE"))
                    Absx1.txtFor("SHIP_TO_KEY_NAME").Text = rowAPTVEND1.Item("VEND_NAME")
            End Select
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTSHIPG", "SOTSHIPG_INACTIVE", "SOTSHIPY"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("SHIP_TO_KEY").Text = ""
        Absx1.txtFor("SHIP_TO_KEY_NAME").Text = ""
        Absx1.txtFor("SHIP_TO_CODE_PFX").Text = ""

        Absx1.txtFor("SHIP_VIA").Text = ""
        Absx1.txtFor("CARRIER_ACCT_TYPE").Text = ""

        Refresh_Records()

        Clear_All_Filters(grdSOTSHIPG)
        Clear_All_Filters(grdSOTSHIPG_INACTIVE)
    End Sub

    Sub Load_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data ...")

        EnforceConstraints(False)

        Absx1.txtFor("SHIP_TO_KEY").Text = SHIP_TO_KEY
        Absx1.txtFor("SHIP_TO_CODE_PFX").Text = SHIP_TO_CODE_PFX
        Absx1.optFor("SHIP_TO_TYPE").Value = SHIP_TO_TYPE

        Dim SHIP_VIA As String = rowSOTSHIPE.Item("SHIP_VIA") & ""
        Dim CARRIER_ACCT_TYPE As String = rowSOTSHIPE.Item("CARRIER_ACCT_TYPE") & ""
        Absx1.txtFor("SHIP_VIA").Text = SHIP_VIA
        Absx1.txtFor("CARRIER_ACCT_TYPE").Text = CARRIER_ACCT_TYPE


        ASCMAIN1.sql = $"Select * from SOTSHIPD where SHIP_TO_CODE_PFX = '{SHIP_TO_CODE_PFX}'"
        Fill_Records("SOTSHIPD",,, ASCMAIN1.sql)

        ASCMAIN1.sql = $"Select * from SOTSHIPF where SHIP_TO_CODE like '{SHIP_TO_CODE_PFX}%'"
        Fill_Records("SOTSHIPF",,, ASCMAIN1.sql)

        Sort_grdColumns(grdSOTSHIPD, "SHIP_TO_CODE")
        grdSOTSHIPD.Rows.ExpandAll(True)

        Fill_Records("SOTSHIPG", New String() {"A", SHIP_TO_CODE_PFX})
        Fill_Records("SOTSHIPG_INACTIVE", New String() {"I", SHIP_TO_CODE_PFX})

        Fill_Records("SOTSHIPY", New String() {"A", SHIP_TO_CODE_PFX})
        Fill_Records("SOTSHIPY_INACTIVE", New String() {"I", SHIP_TO_CODE_PFX})

        Sort_grdColumns(grdSOTSHIPG, "SHIP_TO_CODE")
        Sort_grdColumns(grdSOTSHIPG, "FR_WGT",, 1)

        grdSOTSHIPG.Rows.ExpandAll(True)
        dst.AcceptChanges()

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Updating...")

        ' CUT-OFF TIME?

        For Each rowSOTSHIPG As DataRow In dst.Tables("SOTSHIPG").Select("ROUTING_RULE_STATUS = 'A'", "", DataViewRowState.ModifiedCurrent)
            Dim ROUTING_RULE_NO As String = rowSOTSHIPG.Item("ROUTING_RULE_NO")

            Dim rowSOTSHIPG_CHANGED As DataRow = dst.Tables("SOTSHIPG").NewRow
            rowSOTSHIPG_CHANGED.ItemArray = rowSOTSHIPG.ItemArray
            rowSOTSHIPG_CHANGED.Item("ROUTING_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPG.ROUTING_RULE_NO")
            dst.Tables("SOTSHIPG").Rows.Add(rowSOTSHIPG_CHANGED)

            Dim rowSOTSHIPG_ORIG As DataRow = dst.Tables("SOTSHIPG").Select($"ROUTING_RULE_NO = '{ROUTING_RULE_NO}'", "", DataViewRowState.OriginalRows)(0)
            rowSOTSHIPG_ORIG.RejectChanges()
            rowSOTSHIPG_ORIG.Item("ROUTING_RULE_STATUS") = "I"
        Next

        For Each rowSOTSHIPG_INACTIVE As DataRow In dst.Tables("SOTSHIPG_INACTIVE").Select
            Dim rowSOTSHIPG As DataRow = dst.Tables("SOTSHIPG").NewRow
            rowSOTSHIPG.ItemArray = rowSOTSHIPG_INACTIVE.ItemArray
            dst.Tables("SOTSHIPG").Rows.Add(rowSOTSHIPG)
        Next

        Update_Record_TDA("SOTSHIPG", $"SHIP_TO_CODE like '{SHIP_TO_CODE_PFX}%'")

        MsgBox("Update Complete", MsgBoxStyle.OkOnly, "Verification")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "SHIP_TO_KEY"
                sql_where &= $"SHIP_TO_TYPE = '{optSHIP_TO_TYPE.Value}'"

            Case "SHIP_TO_CODE_PFX"
                sql_where &= $"SHIP_TO_TYPE = '{optSHIP_TO_TYPE.Value}'"

            Case "SHIP_TO_CODE"
                sql_where &= "CUST_STORE_NO IS NULL"

        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTSHIPG, "SSBBB", "Show Filter", "Show GroupBox", "Retire", "Create Default Rule for Key", "Copy Rule to Key", "Create Default Rule for New Ship-Tos", "Copy Rule to New Ship-Tos")
        Load_Popup_Menu(grdSOTSHIPG_INACTIVE, "SSB", "Show Filter", "Show GroupBox", "Re-Activate")
        Load_Popup_Menu(grdSOTSHIPD, "SS", "Show Filter", "Show GroupBox")
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
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdSOTSHIPG"
                    tlb_btn = DirectCast(tlb_pop.Tools("Retire"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"

                    tlb_btn = DirectCast(tlb_pop.Tools("Create Default Rule for Key"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"
                    If tlb_btn.SharedProps.Visible Then
                        tlb_btn.SharedProps.Caption = $"Create Default Rule for {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}"
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Rule to Key"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"
                    If tlb_btn.SharedProps.Visible Then
                        tlb_btn.SharedProps.Caption = $"Copy Rule to {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}"
                    End If

                    tlb_btn = DirectCast(tlb_pop.Tools("Create Default Rule for New Ship-Tos"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E" AndAlso optSHIP_TO_TYPE.Value = "C"

                    tlb_btn = DirectCast(tlb_pop.Tools("Copy Rule to New Ship-Tos"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E" AndAlso optSHIP_TO_TYPE.Value = "C"

                Case "grdSOTSHIPG_INACTIVE"
                    tlb_btn = DirectCast(tlb_pop.Tools("Re-Activate"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = ScreenMode AndAlso EntryMode = "E"
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case ""

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Create Default Rule for Key"
                Dim rowKey() As DataRow = dst.Tables("SOTSHIPG").Select($"SHIP_TO_CODE = '{SHIP_TO_CODE_PFX}D'")

                Dim ROUTING_RULE_NO As String = ""
                If rowKey.Length > 0 Then
                    ROUTING_RULE_NO = rowKey(0).Item("ROUTING_RULE_NO")
                    MsgBox($"There is already a Call-In Rule ({ROUTING_RULE_NO})" & vbCrLf & $" set up for the {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}")
                    Exit Sub
                End If

                Dim SHIP_TO_CODE As String = SHIP_TO_CODE_PFX & "D"
                Dim SHIP_TO_NAME As String = Absx1.txtFor("SHIP_TO_KEY_NAME").Text
                Create_Default_ROUTING_rule(SHIP_TO_CODE, SHIP_TO_NAME)

                Sort_grdColumns(grdSOTSHIPG, "SHIP_TO_CODE")
                Sort_grdColumns(grdSOTSHIPG, "FR_WGT", True, 1)

            Case "Copy Rule to Key"
                Dim rowKey() As DataRow = dst.Tables("SOTSHIPG").Select($"SHIP_TO_CODE = '{SHIP_TO_CODE_PFX}D'")

                Dim ROUTING_RULE_NO As String = ""
                If rowKey.Length > 0 Then
                    ROUTING_RULE_NO = rowKey(0).Item("ROUTING_RULE_NO")
                    MsgBox($"There is already a Call-In Rule ({ROUTING_RULE_NO})" & vbCrLf & $" set up for the {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}")
                    Exit Sub
                End If

                If grdSOTSHIPG.Selected.Rows.Count = 0 And grdSOTSHIPG.ActiveRow IsNot Nothing Then
                    grdSOTSHIPG.ActiveRow.Selected = True
                End If

                If grdSOTSHIPG.Selected.Rows.Count > 1 Then
                    MsgBox($"You may select only 1 row to copy to {optSHIP_TO_TYPE.Text} {SHIP_TO_KEY}")
                    Exit Sub
                End If

                ROUTING_RULE_NO = grdSOTSHIPG.Selected.Rows(0).Cells("ROUTING_RULE_NO").Value
                Dim rowSOTSHIPG_RULE As DataRow = dst.Tables("SOTSHIPG").Rows.Find(ROUTING_RULE_NO)

                Dim rowSOTSHIPG As DataRow = dst.Tables("SOTSHIPG").NewRow
                rowSOTSHIPG.ItemArray = rowSOTSHIPG_RULE.ItemArray
                rowSOTSHIPG.Item("ROUTING_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPG.ROUTING_RULE_NO")
                rowSOTSHIPG.Item("SHIP_TO_CODE") = SHIP_TO_CODE_PFX & "D"
                rowSOTSHIPG.Item("SHIP_TO_NAME") = Absx1.txtFor("SHIP_TO_KEY_NAME").Text
                rowSOTSHIPG.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPG.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPG.Item("INIT_DATE") = DATETIME_STAMP
                rowSOTSHIPG.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTSHIPG.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTSHIPG.Item("LAST_OPER") = ASCMAIN1.USER_ID

                dst.Tables("SOTSHIPG").Rows.Add(rowSOTSHIPG)
                Sort_grdColumns(grdSOTSHIPG, "SHIP_TO_CODE")

            Case "Retire"

                If grdSOTSHIPG.Selected.Rows.Count = 0 And grdSOTSHIPG.ActiveRow IsNot Nothing Then
                    grdSOTSHIPG.ActiveRow.Selected = True
                End If

                If grdSOTSHIPG.Selected.Rows.Count > 0 Then
                    If MsgBox($"OK to Retired the {grdSOTSHIPG.Selected.Rows.Count} selected row(s)?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        Dim ROUTING_RULE_NOs As New List(Of String)
                        For Each grow As UltraWinGrid.UltraGridRow In grdSOTSHIPG.Selected.Rows
                            Dim ROUTING_RULE_NO As String = grow.Cells("ROUTING_RULE_NO").Value
                            ROUTING_RULE_NOs.Add(ROUTING_RULE_NO)
                        Next
                        For Each ROUTING_RULE_NO As String In ROUTING_RULE_NOs
                            Dim rowSOTSHIPG_all As DataRow = dst.Tables("SOTSHIPG").Rows.Find(ROUTING_RULE_NO)
                            If rowSOTSHIPG_all.RowState = DataRowState.Added Then
                                rowSOTSHIPG_all.Delete()
                                'rowSOTSHIPG_all.AcceptChanges()
                            Else
                                Dim rowSOTSHIPG As DataRow = dst.Tables("SOTSHIPG").Select($"ROUTING_RULE_NO = '{ROUTING_RULE_NO}'", "", DataViewRowState.OriginalRows)(0)
                                'If Not rowSOTSHIPG.RowState = DataRowState.Added Then
                                rowSOTSHIPG.RejectChanges()
                                    Dim rowSOTSHIPG_INACTIVE As DataRow = dst.Tables("SOTSHIPG_INACTIVE").NewRow
                                    rowSOTSHIPG_INACTIVE.ItemArray = rowSOTSHIPG.ItemArray
                                    dst.Tables("SOTSHIPG_INACTIVE").Rows.Add(rowSOTSHIPG_INACTIVE)
                                    rowSOTSHIPG_INACTIVE.AcceptChanges()
                                    rowSOTSHIPG_INACTIVE.Item("ROUTING_RULE_STATUS") = "I"
                                    rowSOTSHIPG_INACTIVE.Item("LAST_DATE") = DATETIME_STAMP
                                    rowSOTSHIPG_INACTIVE.Item("LAST_OPER") = ASCMAIN1.USER_ID
                                'End If
                                rowSOTSHIPG.Delete()
                                rowSOTSHIPG.AcceptChanges()
                            End If
                        Next
                        Sort_grdColumns(grdSOTSHIPG_INACTIVE, "SHIP_TO_CODE")
                    End If
                End If

            Case "Re-Activate"

                If grdSOTSHIPG_INACTIVE.Selected.Rows.Count = 0 And grdSOTSHIPG_INACTIVE.ActiveRow IsNot Nothing Then
                    grdSOTSHIPG_INACTIVE.ActiveRow.Selected = True
                End If

                If grdSOTSHIPG_INACTIVE.Selected.Rows.Count > 0 Then

                    Dim ROUTING_RULE_NOs As New List(Of String)
                    For Each grow As UltraWinGrid.UltraGridRow In grdSOTSHIPG_INACTIVE.Selected.Rows
                        Dim ROUTING_RULE_NO As String = grow.Cells("ROUTING_RULE_NO").Value
                        Dim SHIP_TO_CODE As String = grow.Cells("SHIP_TO_CODE").Value
                        If dst.Tables("SOTSHIPG").Select($"SHIP_TO_CODE = '{SHIP_TO_CODE}'").Length > 0 Then
                            MsgBox($"Cannot Re-Activate Call-In Rule {ROUTING_RULE_NO}" & vbCrLf & vbCrLf & $"An Active Call-In Rule already exists" & $" for Ship-To {SHIP_TO_CODE}")
                            Exit Sub
                        End If
                        ROUTING_RULE_NOs.Add(ROUTING_RULE_NO)
                    Next

                    If MsgBox($"OK to Re-Activate the {ROUTING_RULE_NOs.Count} selected row(s)?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                        For Each ROUTING_RULE_NO As String In ROUTING_RULE_NOs
                            Dim rowSOTSHIPG_INACTIVE As DataRow = dst.Tables("SOTSHIPG_INACTIVE").Rows.Find(ROUTING_RULE_NO)
                            If Not rowSOTSHIPG_INACTIVE.RowState = DataRowState.Added Then
                                Dim rowSOTSHIPG As DataRow = dst.Tables("SOTSHIPG").NewRow
                                rowSOTSHIPG.ItemArray = rowSOTSHIPG_INACTIVE.ItemArray
                                dst.Tables("SOTSHIPG").Rows.Add(rowSOTSHIPG)
                                rowSOTSHIPG.AcceptChanges()
                                rowSOTSHIPG.Item("ROUTING_RULE_STATUS") = "A"
                                rowSOTSHIPG.Item("LAST_DATE") = DATETIME_STAMP
                                rowSOTSHIPG.Item("LAST_OPER") = ASCMAIN1.USER_ID
                            End If
                            rowSOTSHIPG_INACTIVE.Delete()
                        Next
                        Sort_grdColumns(grdSOTSHIPG, "SHIP_TO_CODE")
                    End If
                End If

            Case "Create Default Rule for New Ship-Tos", "Copy Rule to New Ship-Tos"

                Dim rowSOTSHIPG As DataRow = Nothing
                If e.Tool.Key = "Copy Rule to New Ship-Tos" Then
                    If grdSOTSHIPG.Selected.Rows.Count = 0 And grdSOTSHIPG.ActiveRow IsNot Nothing Then
                        grdSOTSHIPG.ActiveRow.Selected = True
                    End If

                    If grdSOTSHIPG.Selected.Rows.Count = 0 Then Exit Sub

                    Dim ROUTING_RULE_NO As String = grdSOTSHIPG.ActiveRow.Cells("ROUTING_RULE_NO").Value
                    rowSOTSHIPG = dst.Tables("SOTSHIPG").Rows.Find(ROUTING_RULE_NO)

                End If

                Setup_ARTCUST2_SHIPTO(rowSOTSHIPG)

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SHIP_TO_KEY"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Validate_SHIP_TO_KEY() Then
                        Click_Command("Load")
                    End If
                End If

            Case "SHIP_TO_CODE_PFX"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Validate_SHIP_TO_CODE_PFX() Then
                        Click_Command("Load")
                    End If
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "SHIP_TO_KEY"
                If Validate_SHIP_TO_KEY() Then
                    Click_Command("Load")
                End If

            Case "SHIP_TO_CODE_PFX"
                If Validate_SHIP_TO_CODE_PFX() Then
                    Click_Command("Load")
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SHIP_TO_KEY"
                If Validate_SHIP_TO_KEY() Then
                    Click_Command("Load")
                End If

            Case "SHIP_TO_CODE_PFX"
                If Validate_SHIP_TO_CODE_PFX() Then
                    Click_Command("Load")
                End If
        End Select
    End Sub

#End Region

    Sub Refresh_Records()

        EnforceConstraints(False)
        Fill_Records("SOTSHIPG", New String() {"A", "*"})
        Fill_Records("SOTSHIPY", New String() {"A", "*"})
        EnforceConstraints(False)

        Sort_grdColumns(grdSOTSHIPG, "SHIP_TO_CODE")
        Sort_grdColumns(grdSOTSHIPG, "FR_WGT", True, 1)

        grdSOTSHIPG.Rows.ExpandAll(True)

        ASCMAIN1.sql = $"Select * from SOTSHIPD"
        Fill_Records("SOTSHIPD",,, ASCMAIN1.sql)

        ASCMAIN1.sql = $"Select * from SOTSHIPF"
        Fill_Records("SOTSHIPF",,, ASCMAIN1.sql)

        Sort_grdColumns(grdSOTSHIPD, "SHIP_TO_CODE")
        grdSOTSHIPD.Rows.ExpandAll(True)

    End Sub

    Private Sub grdSOTSHIPG_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTSHIPG.DoubleClickRow
        If Not ScreenMode Then
            If e.Row.IsDataRow Then
                Dim SHIP_TO_CODE As String = ""
                If e.Row.Band.Index = 1 Then
                    SHIP_TO_CODE = e.Row.ParentRow.Cells("SHIP_TO_CODE").Value
                Else
                    SHIP_TO_CODE = e.Row.Cells("SHIP_TO_CODE").Value
                End If

                Dim rowSOTSHIPD As DataRow = LookUp("SOTSHIPD", SHIP_TO_CODE)
                Dim SHIP_TO_CODE_PFX As String = rowSOTSHIPD.Item("SHIP_TO_CODE_PFX")

                Dim rowSOTSHIPE As DataRow = LookUp("SOTSHIPE", SHIP_TO_CODE_PFX)
                Dim SHIP_TO_KEY As String = rowSOTSHIPE.Item("SHIP_TO_KEY")
                Dim SHIP_TO_TYPE As String = rowSOTSHIPE.Item("SHIP_TO_TYPE")

                Absx1.txtFor("SHIP_TO_KEY").Text = SHIP_TO_KEY
                Absx1.txtFor("SHIP_TO_CODE_PFX").Text = SHIP_TO_CODE_PFX
                Absx1.optFor("SHIP_TO_TYPE").Value = SHIP_TO_TYPE

                Click_Command("Load")
            End If
        End If

    End Sub

    Function Validate_SHIP_TO_KEY() As Boolean

        Dim SHIP_TO_KEY As String = Absx1.txtFor("SHIP_TO_KEY").Text
        Dim SHIP_TO_TYPE As String = Absx1.optFor("SHIP_TO_TYPE").Value

        ASCMAIN1.sql = "Select * from SOTSHIPE where SHIP_TO_TYPE = :PARM1 and SHIP_TO_KEY = :PARM2"
        Dim rowSOTSHIPE As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "VV", New String() {SHIP_TO_TYPE, SHIP_TO_KEY})

        If rowSOTSHIPE Is Nothing Then
            Return False
        Else
            Dim SHIP_TO_CODE_PFX As String = rowSOTSHIPE.Item("SHIP_TO_CODE_PFX")
            Absx1.txtFor("SHIP_TO_CODE_PFX").Text = SHIP_TO_CODE_PFX
            Return True
        End If
    End Function

    Function Validate_SHIP_TO_CODE_PFX() As Boolean

        Dim SHIP_TO_CODE_PFX As String = Absx1.txtFor("SHIP_TO_CODE_PFX").Text

        ASCMAIN1.sql = "Select * from SOTSHIPE where SHIP_TO_CODE_PFX = :PARM1"
        Dim rowSOTSHIPE As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "V", New String() {SHIP_TO_CODE_PFX})

        If rowSOTSHIPE Is Nothing Then
            Return False
        Else
            Dim SHIP_TO_KEY As String = rowSOTSHIPE.Item("SHIP_TO_KEY")
            Absx1.txtFor("SHIP_TO_KEY").Text = SHIP_TO_KEY
            Return True
        End If
    End Function

    Sub Setup_ARTCUST2_SHIPTO(rowSOTSHIPG_RULE As DataRow)
        Dim CUST_CODE As String = Absx1.txtFor("SHIP_TO_KEY").Text
        Fill_Records("ARTCUST2_SHIPTO", CUST_CODE)

        For Each T As String In New String() {"SOTSHIPG", "SOTSHIPG_INACTIVE"}
            For Each rowT As DataRow In dst.Tables(T).Select("")
                Dim SHIP_TO_CODE As String = rowT.Item("SHIP_TO_CODE")
                Dim CUST_STORE_NO As String = Mid(SHIP_TO_CODE, 9)
                If CUST_STORE_NO <> "" Then
                    Dim rowARTCUST2_SHIPTO As DataRow = dst.Tables("ARTCUST2_SHIPTO").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
                    rowARTCUST2_SHIPTO.Delete()
                End If
            Next
        Next
        dst.Tables("ARTCUST2_SHIPTO").AcceptChanges()

        Dim CUST_STORE_NOs As List(Of String) = Select_ARTCUST2_SHIPTOs(dst.Tables("ARTCUST2_SHIPTO"))

        For Each CUST_STORE_NO As String In CUST_STORE_NOs
            Dim rowARTCUST2_SHIPTO As DataRow = dst.Tables("ARTCUST2_SHIPTO").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO})
            Dim SHIP_TO_CODE As String = rowARTCUST2_SHIPTO.Item("SHIP_TO_CODE")
            Dim CUST_STORE_NAME As String = rowARTCUST2_SHIPTO.Item("CUST_STORE_NAME")
            If rowSOTSHIPG_RULE Is Nothing Then
                Create_Default_ROUTING_rule(SHIP_TO_CODE, CUST_STORE_NAME)
            Else
                Dim rowSOTSHIPG As DataRow = dst.Tables("SOTSHIPG").NewRow
                rowSOTSHIPG.ItemArray = rowSOTSHIPG_RULE.ItemArray
                rowSOTSHIPG.Item("ROUTING_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPG.ROUTING_RULE_NO")
                rowSOTSHIPG.Item("SHIP_TO_CODE") = SHIP_TO_CODE_PFX & "D" & CUST_STORE_NO
                rowSOTSHIPG.Item("SHIP_TO_NAME") = CUST_STORE_NAME

                rowSOTSHIPG.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPG.Item("SHIPTO1") = DBNull.Value
                rowSOTSHIPG.Item("INIT_DATE") = DATETIME_STAMP
                rowSOTSHIPG.Item("INIT_OPER") = ASCMAIN1.USER_ID
                rowSOTSHIPG.Item("LAST_DATE") = DATETIME_STAMP
                rowSOTSHIPG.Item("LAST_OPER") = ASCMAIN1.USER_ID

                dst.Tables("SOTSHIPG").Rows.Add(rowSOTSHIPG)
            End If
        Next
        Sort_grdColumns(grdSOTSHIPG, "SHIP_TO_CODE")

    End Sub

    Function Select_ARTCUST2_SHIPTOs(tbl As DataTable) As List(Of String)

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("CUST_STORE_NO_CALL_IN")
        Dim CUST_STORE_NOs As New List(Of String)

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            ASCMAIN1.CodeSelector.UseDataFromTable = tbl
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                For Each CUST_STORE_NO As String In ASCMAIN1.CodeSelector.SelectedCodes
                    CUST_STORE_NOs.Add(CUST_STORE_NO)
                Next
            End If
        End If

        Return CUST_STORE_NOs
    End Function

    Sub Create_Default_ROUTING_rule(SHIP_TO_CODE As String, SHIP_TO_NAME As String)
        Dim rowSOTSHIPG As DataRow = dst.Tables("SOTSHIPG").NewRow
        With rowSOTSHIPG
            .Item("ROUTING_RULE_NO") = ASCMAIN1.Next_Control_No("SOTSHIPG.ROUTING_RULE_NO")
            .Item("ROUTING_RULE_STATUS") = "A"
            .Item("SHIP_TO_CODE") = SHIP_TO_CODE
            .Item("SHIP_TO_NAME") = SHIP_TO_NAME

            .Item("FR_WGT") = 1
            .Item("TO_WGT") = 999999

            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DATETIME_STAMP
            .Item("LAST_OPER") = ASCMAIN1.USER_ID
        End With
        dst.Tables("SOTSHIPG").Rows.Add(rowSOTSHIPG)
    End Sub

    Function Validate_Active_Records() As String

        Dim EMsg As String = ""

        For Each rowSOTSHIPG As DataRow In dst.Tables("SOTSHIPG").Select

            Dim ROUTING_RULE_NO As String = rowSOTSHIPG.Item("ROUTING_RULE_NO") & ""

            ' MAKE SURE RULES ARE CONTIGUOUS

            'If SHIP_RULE_CODE = "R" And PICK_UP_DAYS = "0000000" Then
            '    EMsg &= vbCr & $"{ROUTING_RULE_NO} Ship Rule Code (R) requires that there are Specific Pick-Up Days specifed"
            'End If
        Next

        Return EMsg
    End Function

    Private Sub grdSOTSHIPG_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTSHIPG.InitializeRow
        If e.Row.IsDataRow And e.Row.Band.Index = 1 Then
            With e.Row.Cells("DEL_METHOD")
                If .Value & "" = "SPS" Then
                    .Appearance.ForeColor = System.Drawing.Color.Red
                ElseIf .Value & "" = "LTL" Then
                    .Appearance.ForeColor = System.Drawing.Color.Blue
                Else
                    .Appearance.ForeColor = System.Drawing.Color.Empty
                End If
            End With
        End If
    End Sub

    'Sub Create_Bands()

    '    dst.Tables("SOTSHIPG_WGT").Rows.Clear()
    '    For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTSHIPG"), "SHIP_TO_CODE").Select
    '        Dim SHIP_TO_CODE As String = row.Item("SHIP_TO_CODE")
    '        Dim ROUTING_RULE_NO As String = ""
    '        Dim FR_WGT As Int32 = 0
    '        Dim FR_CTN As Int32 = 0

    '        For Each row2 As DataRow In dst.Tables("SOTSHIPG").Select($"SHIP_TO_CODE = '{SHIP_TO_CODE}'", "FR_WGT,FR_CTN")
    '            If FR_WGT = 0 Or (FR_WGT >= Val(row2.Item("FR_WGT")) And FR_CTN >= Val(row2.Item("FR_CTN"))) Then
    '                FR_WGT = Val(row2.Item("FR_WGT"))
    '                FR_CTN = Val(row2.Item("FR_CTN"))
    '                ROUTING_RULE_NO = ""
    '            End If
    '            If ROUTING_RULE_NO = "" Then
    '                ROUTING_RULE_NO = row2.Item("ROUTING_RULE_NO")
    '            End If

    '            dst.Tables("SOTSHIPG_WGT").Rows.Add(New Object() _
    '                {ROUTING_RULE_NO, row2.Item("FR_WGT"), row2.Item("DEL_METHOD"), row2.Item("SHIP_VIA"), row2.Item("SHIP_VIA_DESC"), row2.Item("SHIP_VIA_SCAC"), row2.Item("CARRIER"), row2.Item("SHIP_METHOD")})
    '            If ROUTING_RULE_NO <> row2.Item("ROUTING_RULE_NO") Then
    '                row2.Delete()
    '            End If
    '        Next
    '    Next
    'End Sub
End Class