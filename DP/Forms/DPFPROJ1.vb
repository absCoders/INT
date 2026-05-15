Public Class DPFPROJ1

    Dim RYP1 As String
    Dim RYP2 As String
    Dim RYP3 As String

    Dim YP() As String
    Dim YM() As String

    Dim ACT_M As Integer ' Last Month of Actuals
    Dim LCK_M As Integer ' Last Locked Month

    Dim OPS_YYYYPP_LCK As String
    Dim OPS_YYYYPP_LCK_LEGEND As String

    Dim DPTPROJX As String

    Dim sqlICTPRICX As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")

        Dim P As Integer = 12 - Val(Mid(ASCMAIN1.CYP, 5, 2))
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, P + 24, P)
        Set_cmbYP("RYP2", ASCMAIN1.CYP, -36, P + 24, -1)
        Set_cmbYP("RYP3", ASCMAIN1.CYP, -36, P + 24, -1)

        With dst
            Create_TDA(.Tables.Add, "ICTBRAN1", "*")
            Create_TDA(.Tables.Add, "ICTITEM1", "*")
            Create_TDA(.Tables.Add, "ICTPRICA", "*")
            Create_TDA(.Tables.Add, "DPTITMF1", "*")

            sqlICTPRICX = "Select X.*" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTITEM1.PG_STAT_FACTOR, ICTITEM1.ITEM_RETAIL_PRICE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " from ICTITEM1,ICTCOLL1,(Select ITEM_CODE" & vbCrLf
            For I As Integer = 1 To 12
                sqlICTPRICX &= ", SUM (DECODE(OPS_YYYYPP,'" & Format(I, "000000") & "',ITEM_UNIT_PRICE,0)) PRICE_" & Format(I, "00") & vbCrLf
            Next
            sqlICTPRICX &= " from ICTPRICA" & vbCrLf _
            & " where OPS_YYYYPP >= '000001' and OPS_YYYYPP <= '000012'" & vbCrLf _
            & " group by ICTPRICA.ITEM_CODE) X " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE " & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE"
            ASCMAIN1.sql = sqlICTPRICX
            Create_TDA(.Tables.Add, "ICTPRICX", "**", 0, False, , 1)

            Create_DPTPROJX(True)

            ASCMAIN1.sql = "Select MARKET_CODE, ITEM_CODE, ITEM_PRICE"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", Sum (QS" & Format(M, "00") & ") QS" & Format(M, "00")
                ASCMAIN1.sql &= ", Sum (AS" & Format(M, "00") & ") AS" & Format(M, "00")
                ASCMAIN1.sql &= ", Sum (QF" & Format(M, "00") & ") QF" & Format(M, "00")
                ASCMAIN1.sql &= ", Sum (AF" & Format(M, "00") & ") AF" & Format(M, "00")
            Next
            ASCMAIN1.sql &= ", Sum (QSOP) QSOP, Sum (ASOP) ASOP, Sum (SSOP) SSOP"
            ASCMAIN1.sql &= ", Sum (QBPE) QBPE, Sum (ABPE) ABPE, Sum (SBPE) SBPE"
            ASCMAIN1.sql &= " from " & DPTPROJX & " group by MARKET_CODE, ITEM_CODE, ITEM_PRICE"

            ASCMAIN1.sql = "Select X.*" _
            & ", ICTCOLL1.BRAND_CODE, ICTITEM1.COLLECTION_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.PG_STAT_FACTOR, ICTITEM1.ITEM_RETAIL_PRICE" _
            & " from ICTITEM1,ICTCOLL1" _
            & ", (" & ASCMAIN1.sql & ") X" _
            & " where X.ITEM_CODE = ICTITEM1.ITEM_CODE " _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE "
            Create_TDA(.Tables.Add, "DPTPROJ1", "**", 0, False, , 2)

            For M As Integer = 1 To 12
                .Tables("DPTPROJ1").Columns.Add("DS" & Format(M, "00"), GetType(System.Decimal))
                .Tables("DPTPROJ1").Columns.Add("DF" & Format(M, "00"), GetType(System.Decimal))
                .Tables("DPTPROJ1").Columns("DF" & Format(M, "00")).ReadOnly = False
                .Tables("DPTPROJ1").Columns.Add("SS" & Format(M, "00"), _
                                                 GetType(System.Decimal), _
                                                 "QS" & Format(M, "00") & " * ISNULL(PG_STAT_FACTOR,0) / 1000")
                .Tables("DPTPROJ1").Columns.Add("SF" & Format(M, "00"), _
                                                 GetType(System.Decimal), _
                                                 "QF" & Format(M, "00") & " * ISNULL(PG_STAT_FACTOR,0) / 1000")
            Next
            For Each SF As String In New String() {"S", "F"}
                For Each QAS As String In New String() {"Q", "A", "S", "D"}
                    Dim CT As String = QAS & SF
                    .Tables("DPTPROJ1").Columns.Add(CT & "Q1", GetType(System.Decimal), Replace("ISNULL(XX01,0) + ISNULL(XX02,0) + ISNULL(XX03,0)", "XX", CT))
                    .Tables("DPTPROJ1").Columns.Add(CT & "Q2", GetType(System.Decimal), Replace("ISNULL(XX04,0) + ISNULL(XX05,0) + ISNULL(XX06,0)", "XX", CT))
                    .Tables("DPTPROJ1").Columns.Add(CT & "Q3", GetType(System.Decimal), Replace("ISNULL(XX07,0) + ISNULL(XX08,0) + ISNULL(XX09,0)", "XX", CT))
                    .Tables("DPTPROJ1").Columns.Add(CT & "Q4", GetType(System.Decimal), Replace("ISNULL(XX10,0) + ISNULL(XX11,0) + ISNULL(XX12,0)", "XX", CT))
                    .Tables("DPTPROJ1").Columns.Add(CT & "Q0", GetType(System.Decimal), Replace("XXQ1 + XXQ2 + XXQ3 + XXQ4", "XX", CT))
                Next
            Next

            For Each TOT As String In New String() {"FCR", "ACT", "PROJ", "SOP", "SOPD", "BPE", "BPED"}
                If TOT <> "SOP" And TOT <> "BPE" Then
                    For Each QAS As String In New String() {"Q", "A", "S"}
                        Dim CT As String = QAS & TOT
                        Dim DC As DataColumn = .Tables("DPTPROJ1").Columns.Add(CT, GetType(System.Decimal))
                        Select Case TOT
                            Case "PROJ"
                                DC.Expression = "ISNULL(" & QAS & "FCR,0) + ISNULL(" & QAS & "ACT,0)"
                            Case "SOPD"
                                DC.Expression = "ISNULL(" & QAS & "PROJ,0) - ISNULL(" & QAS & "SOP,0)"
                            Case "BPED"
                                DC.Expression = "ISNULL(" & QAS & "PROJ,0) - ISNULL(" & QAS & "BPE,0)"
                        End Select
                    Next
                End If
                .Tables("DPTPROJ1").Columns.Add("D" & TOT, GetType(System.Decimal))
            Next

            Create_Relation("ICTPRICX", "DPTPROJ1", "ITEM_CODE")
            For M As Integer = 1 To 12
                .Tables("DPTPROJ1").Columns.Add("PRICE_" & Format(M, "00"), _
                                                GetType(System.Decimal), _
                                                "PARENT.PRICE_" & Format(M, "00"))
            Next
            .Tables("DPTPROJ1").Columns.Add("NET_PRICE", GetType(System.Decimal), "PARENT.PRICE_" & Format(12, "00"))
            .Tables("DPTPROJ1").Columns.Add("DISC_PCT", GetType(System.Decimal), "IIF(ISNULL(ITEM_RETAIL_PRICE,0)=0,0,100*(ISNULL(ITEM_RETAIL_PRICE,0) - ISNULL(NET_PRICE,0)) / ISNULL(ITEM_RETAIL_PRICE,0))")

            Create_TDA(.Tables.Add, "DPTITMF3", "*")

            .Tables.Add("DPTPROJS")

            With .Tables.Add("SOTMKTCA")
                .Columns.Add("MARKET_CODE")
                .Columns.Add("MARKET_DESC")
                .Columns.Add("SEL")
            End With

            Dim TOTX As String = ""
            Dim TOTCOLS As String = ""
            For I As Integer = 1 To 18
                TOTX &= "+FC" & Format(I, "00")
                Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, I - 1)
                TOTCOLS &= ", SUM (DECODE(OPS_YYYYPP_FC,'" & YP & "',NVL(FORECAST,0),0)) FC" & Format(I, "00") & vbCrLf
            Next

            ASCMAIN1.sql = "Select 'US' COL1, 'R' COL2, 'SECMARKETS' COL3, 'IT' COL4, ITEM_CODE " & vbCrLf _
            & TOTCOLS _
            & " from DPTITMF1 where MARKET_CODE in (Select MARKET_CODE from SOTMKTC1 where CUST_CODE is Not Null)" & vbCrLf _
            & " and OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " group by ITEM_CODE"
            Create_TDA(.Tables.Add, "DPTITMFS", "**", , False)
            '.Tables("DPTITMFS").Columns.Add("TOTAL", GetType(System.Int64), Mid(TOTX, 2))
        End With

        grdICTPRICX.DataSource = dst.Tables("ICTPRICX")
        Create_Summary(grdICTPRICX, "ITEM_CODE", "Count")

        Dim COLUMNS_shown As String() = {"MARKET_CODE", "BRAND_CODE", "COLLECTION_CODE", "ITEM_CODE", "ITEM_DESC", "ITEM_RETAIL_PRICE", "NET_PRICE", "DISC_PCT"}

        grdDPTPROJ1.DataSource = dst.Tables("DPTPROJ1")
        For Each gcol As UltraWinGrid.UltraGridColumn In grdDPTPROJ1.DisplayLayout.Bands(0).Columns
            Dim COL As String = gcol.Key
            If COLUMNS_shown.Contains(COL) _
            Or Mid(COL, 1, 1) = "D" And ((Mid(COL, 3) >= "01" And Mid(COL, 3) <= "12") _
                                         Or (Mid(COL, 3) >= "Q0" And Mid(COL, 3) <= "Q4") _
                                         Or New String() {"FCR", "ACT", "PROJ", "SOP", "SOPD", "BPE", "BPED"}.Contains(Mid(COL, 2))) Then
                gcol.Hidden = False

                Dim DC As DataColumn = dst.Tables("DPTPROJ1").Columns(COL)
                dst.Tables("DPTPROJS").Columns.Add(DC.ColumnName, DC.DataType)

            Else
                gcol.Hidden = True
            End If
        Next

        spl.Panel1Collapsed = True

        grdDPTPROJS.DataSource = dst.Tables("DPTPROJS")
        grdSOTMKTCA.DataSource = dst.Tables("SOTMKTCA")
        grdDPTITMFS.DataSource = dst.Tables("DPTITMFS")

        With grdDPTITMFS.DisplayLayout.Bands(0)
            .Columns("COL1").Header.Caption = "US"
            .Columns("COL2").Header.Caption = "R"
            .Columns("COL3").Header.Caption = "SECMARKETS"
            .Columns("COL4").Header.Caption = "IT"
            .Columns("ITEM_CODE").Header.Caption = "Item"
            '.Columns("TOTAL").Header.Caption = "Total"
            'Create_Summary(grdDPTITMFS, "ITEM_CODE", "Count")
            For i As Integer = 1 To 18
                Dim YP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, i - 1)
                Dim ROWGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YP, False, True)
                ' .Columns("FC" & Format(i, "00")).Header.Caption = LEGEND
                .Columns("FC" & Format(i, "00")).Header.Caption = Format(ROWGLTPARM2.Item("PRD_END_DATE"), "01.MM.yyyy")
                .Columns("FC" & Format(i, "00")).Format = "#####0"
                ' Create_Summary(grdDPTITMFS, "FC" & Format(i, "00"))
            Next
        End With

      
        'cmbAddMARKET_CODE.DataSource = ASCDATA1.GetDataTable("Select MARKET_CODE, MARKET_DESC from SOTMKTC1 order by MARKET_CODE")

        Format_grd(grdDPTPROJ1)
        Format_grd(grdDPTPROJS)

        ASCMAIN1.Add_Value_List(grdDPTPROJ1, "BRAND_CODE", "Select BRAND_CODE, BRAND_NAME from ICTBRAN1")
        ASCMAIN1.Add_Value_List(grdDPTPROJ1, "COLLECTION_CODE", "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1")
        ASCMAIN1.Add_Value_List(grdDPTPROJS, "BRAND_CODE", "Select BRAND_CODE, BRAND_NAME from ICTBRAN1")
        ASCMAIN1.Add_Value_List(grdDPTPROJS, "COLLECTION_CODE", "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1")
        ASCMAIN1.Add_Value_List(grdICTPRICX, "BRAND_CODE", "Select BRAND_CODE, BRAND_NAME from ICTBRAN1")
        ASCMAIN1.Add_Value_List(grdICTPRICX, "COLLECTION_CODE", "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
            Case "Edit"
                ' If ACT_M > 11 Then
                If LCK_M >= 12 Then
                    EMsg &= vbCr & "No Editable Periods on Display"
                End If

                If ACT_M < 0 Or ACT_M > 11 Then
                    EMsg &= vbCr & "No Editable Periods on Display"
                End If

                If Not ASCMAIN1.Logical_Lock("DPTITMF1", "*") Then
                    Exit Sub
                End If

            Case "Unlock MMM'YY"

                If OPS_YYYYPP_LCK < ASCMAIN1.CYP Then
                    EMsg &= vbCr & "You Cannot Unlock a period Prior to the Current Period"
                End If

                If EMsg = "" Then

                    If MsgBox("Are you Sure that you want to Unlock Forecasts for " & OPS_YYYYPP_LCK_LEGEND & "?" _
                              & vbCrLf _
                              & vbCrLf & "Click Yes to Delete all Locked Forecasts for " & OPS_YYYYPP_LCK_LEGEND & "." _
                              & vbCrLf _
                              & vbCrLf & "The current values for " & OPS_YYYYPP_LCK_LEGEND & "'s Forecast" _
                              & vbCrLf & " (which may have been changed since the period was locked)" _
                              & vbCrLf & " will remain as the current values for " & OPS_YYYYPP_LCK_LEGEND & "." _
                              & vbCrLf _
                              & vbCrLf & "You may then make edits are required to the forecast values for " & OPS_YYYYPP_LCK_LEGEND _
                              & vbCrLf & " and then re-lock it when you are ready." _
                              & vbCrLf _
                              & vbCrLf & "OK to Unlock the Forecast for " & OPS_YYYYPP_LCK_LEGEND & "?", _
                              MsgBoxStyle.YesNo, _
                              "Option to Unlock Foreacasts for a Period") <> MsgBoxResult.Yes Then
                        Exit Sub
                    End If
                End If


            Case "Update"
                If LCK_M < 0 Or LCK_M > 11 Then
                    EMsg &= vbCr & "No Updates Permitted - entire 12 Month Period is Locked"
                End If

            Case "Export Summary"
                If Not grdDPTPROJS.Visible Then
                    EMsg &= vbCr & "You must generate a Summary before using this option"
                End If

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
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Eliminate 0's"

                If MsgBox("This command will eliminate all records in the Projected Shipments grid" _
                          & vbCrLf & " which have a 0 value for both Shipments and Forecast" _
                          & vbCrLf & " for all 12 months shown, and then eliminate all Items" _
                          & vbCrLf & " from the Prices grid where there are no corresponding" _
                          & vbCrLf & " records in the Projected Shipments grid", _
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If

                Dim sqlw As String = ""
                For i As Integer = 1 To 12
                    sqlw &= "and ISNULL(QS" & Format(i, "00") & ",0) = 0"
                    sqlw &= "and ISNULL(QF" & Format(i, "00") & ",0) = 0"
                Next
                ASCDATA1.DeleteRows("DPTPROJ1", Mid(sqlw, 5))

                For Each rowICTPRICX As DataRow In dst.Tables("ICTPRICX").Select
                    If rowICTPRICX.GetChildRows("ICTPRICX_DPTPROJ1").Length = 0 Then
                        rowICTPRICX.Item("ITEM_DESC") = "DELETE_ME"
                    End If
                Next
                ASCDATA1.DeleteRows("ICTPRICX", "ITEM_DESC = 'DELETE_ME'")

                MsgBox("0's have been Eliminated", MsgBoxStyle.OkOnly, "Success")


            Case "Unlock MMM'YY"
                Unlock_Period()
                Mode_Settings(False)

            Case "18-Month by Item"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Generating 18-Month Forecast by Item")

                Fill_Records("DPTITMFS")
                grdDPTITMFS.Visible = True

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")


            Case "Export Summary"

                Export_Summary()
         
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode

                If ScreenMode And EntryMode = "L" Then
                    .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.False
                End If
                If ScreenMode And EntryMode <> "L" Then
                    .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                    .Groups("Screen Control").Items("Eliminate 0's").Settings.Enabled = iScreenMode
                Else
                    .Groups("Screen Control").Items("Update").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Cancel").Settings.Enabled = DefaultableBoolean.False
                    .Groups("Screen Control").Items("Eliminate 0's").Settings.Enabled = DefaultableBoolean.False
                End If
                If ScreenMode And EntryMode = "L" Then
                    .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                Else
                    .Groups("Screen Control").Items("Done").Settings.Enabled = DefaultableBoolean.False
                End If


                .Groups("Screen Control").Items("18-Month by Item").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Export Summary").Settings.Enabled = iScreenMode


                .Groups("Scope").Visible = Not ScreenMode

                .Groups("Add Item").Visible = ScreenMode And EntryMode = "E"

                '.Groups("Summarize By").Visible = ScreenMode And (EntryMode = "L")
                '.Groups("Data Options").Visible = ScreenMode
                '.Groups("Prices").Visible = False

                .Groups("Screen Control").Items("Unlock MMM'YY").Visible = Not ScreenMode

            End With
        End If

        tabMain.SelectedTab = tabMain.Tabs("Projections")
        Setup_tabMain()

        grdDPTITMFS.Visible = False

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        If EntryMode = "E" Then
            grdICTPRICX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdDPTPROJ1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            grdDPTPROJ1.DisplayLayout.Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
        Else
            grdICTPRICX.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
            grdDPTPROJ1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If

        'chkLock.Visible = ScreenMode And (EntryMode = "E")

        'grdDPTPROJ1.Visible = ScreenMode
        'grdDPTPROJS.Visible = ScreenMode
        tabMain.Visible = ScreenMode
        Set_Upload()

        Absx1.cmbFor("RYP2").ReadOnly = True

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        dst.Tables("DPTPROJ1").Rows.Clear()
        dst.Tables("DPTPROJS").Rows.Clear()
        dst.Tables("ICTPRICX").Rows.Clear()
        dst.Tables("DPTITMFS").Rows.Clear()
        EnforceConstraints(True)

        chkShowFCOnly.Checked = False
        Clear_Summary()

        grdDPTPROJ1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdDPTPROJS.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        grdICTPRICX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        Get_LCK()

    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Data ...")

        Call Save_Header_Fields(UltraGroupBox1)

        RYP1 = Absx1.cmbFor("RYP1").Value
        RYP2 = Absx1.cmbFor("RYP2").Value
        RYP3 = Absx1.cmbFor("RYP3").Value

        ReDim YP(12)
        YP(12) = RYP1
        For I As Integer = 1 To 11
            YP(I) = ASCMAIN1.Period_Calc(RYP1, -12 + I)
        Next
        ReDim YM(12)
        For I As Integer = 1 To 12
            YM(I) = ASCMAIN1.Get_Legend(YP(I), False, True)
        Next

        Dim ACT_D As Integer = ASCMAIN1.Period_Diff(ASCMAIN1.CYP, RYP1)
        If ACT_D >= 12 Then
            'ACT_M = -1
            ACT_M = 0 ' MAKE BELIEVE THAT THE PERIOD JUST TO THE LEFT OF THE 12 MONTHS SELECTED IS THE LAST ACTUAL PERIOD
        ElseIf ACT_D < 0 Then
            ACT_M = 12
        Else
            ACT_M = 12 - ACT_D - 1
        End If

        Get_LCK()

        Dim LEGEND As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(OPS_YYYYPP_LCK, 1), False, True)
        Dim LCK_D As Integer = ASCMAIN1.Period_Diff(OPS_YYYYPP_LCK, RYP1)
        If LCK_D >= 12 Then
            LCK_M = 0
        ElseIf LCK_D < 0 Then
            LCK_M = 12
        Else
            LCK_M = 12 - LCK_D
        End If

        chkLock.Checked = False
        If LCK_D > 0 And LCK_M <= 11 Then
            chkLock.Visible = True
            chkLock.Text = "Lock " & LEGEND & " Upon Update"
        Else
            chkLock.Visible = False
        End If

        EnforceConstraints(False)
        Create_DPTPROJX(False)
        For M As Integer = 1 To 12
            dst.Tables("DPTPROJ1").Columns("AF" & Format(M, "00")).Expression = ""
        Next
        Fill_Records("DPTPROJ1")
        'If LCK_M < 12 Then
        '    For M As Integer = LCK_M + 1 To 12
        '        dst.Tables("DPTPROJ1").Columns("AF" & Format(M, "00")).Expression = "QF" & Format(M, "00") & " * ITEM_PRICE / 1000"
        '    Next
        'End If


        Dim SQL As String = sqlICTPRICX
        For M As Integer = 1 To 12
            SQL = Replace(SQL, "'" & Format(M, "000000") & "'", "'" & YP(M) & "'")
        Next
        Fill_Records("ICTPRICX", "", , Sql)

        Dim PYP As String = ASCMAIN1.Period_Calc(YP(1), -1)
        For Each rowDPTPROJ1 As DataRow In dst.Tables("DPTPROJ1").Rows
            Dim ITEM_CODE As String = rowDPTPROJ1.Item("ITEM_CODE")
            Dim rowICTPRICX As DataRow = dst.Tables("ICTPRICX").Rows.Find(ITEM_CODE)
            If rowICTPRICX Is Nothing Then
                rowICTPRICX = dst.Tables("ICTPRICX").NewRow
                rowICTPRICX.Item("ITEM_CODE") = ITEM_CODE
                rowICTPRICX.Item("ITEM_DESC") = rowDPTPROJ1.Item("ITEM_DESC")
                rowICTPRICX.Item("COLLECTION_CODE") = rowDPTPROJ1.Item("COLLECTION_CODE")
                rowICTPRICX.Item("PG_STAT_FACTOR") = rowDPTPROJ1.Item("PG_STAT_FACTOR")
                rowICTPRICX.Item("ITEM_RETAIL_PRICE") = rowDPTPROJ1.Item("ITEM_RETAIL_PRICE")
                rowICTPRICX.Item("BRAND_CODE") = rowDPTPROJ1.Item("BRAND_CODE")
                Dim rowICTPRICA As DataRow = LookUp("ICTPRICA", New String() {ITEM_CODE, PYP})
                If rowICTPRICA IsNot Nothing Then
                    For M As Integer = 1 To 12
                        rowICTPRICX.Item("PRICE_" & Format(M, "00")) = rowICTPRICA.Item("ITEM_UNIT_PRICE")
                    Next
                End If
                dst.Tables("ICTPRICX").Rows.Add(rowICTPRICX)
            End If
        Next
        Sort_grdColumns(grdICTPRICX, "BRAND_CODE,COLLECTION_CODE,ITEM_CODE")

        EnforceConstraints(True)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdDPTPROJ1, grdDPTPROJS}
            With grd.DisplayLayout.Bands(0)
                For M As Integer = 1 To 12
                    .Columns("DF" & Format(M, "00")).Header.Caption = YM(M) & vbCrLf & "F/C"
                    .Columns("DS" & Format(M, "00")).Header.Caption = YM(M) & vbCrLf & "Act"

                    'If YP(M) > RYP2 Then
                    '    .Groups("P" & Format(M, "00")).Header.Appearance.BackColor2 = Color.Yellow
                    'Else
                    '    .Groups("P" & Format(M, "00")).Header.Appearance.BackColor2 = Color.LightBlue
                    'End If
                    '.Groups("P" & Format(M, "00")).Header.Appearance.BackGradientStyle = GradientStyle.GlassBottom20 ' GradientStyle.ForwardDiagonal
                Next
            End With
        Next

        With grdICTPRICX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "PG_STAT_FACTOR", "BRAND_CODE", "ITEM_RETAIL_PRICE"}
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
                If COLUMN_NAME <> "ITEM_DESC" Then
                    .Columns(COLUMN_NAME).Width = 80
                End If
            Next

            For M As Integer = 1 To 12
                With .Columns("PRICE_" & Format(M, "00"))
                    .Header.Caption = YM(M)
                    .Width = 60
                    .Format = "###,##0.00"
                    If M > LCK_M Then 'If YP(M) >= ASCMAIN1.CYP Then
                        .CellActivation = UltraWinGrid.Activation.AllowEdit
                        .CellAppearance.BackColor = Color.Empty
                        .CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
                    Else
                        .CellActivation = UltraWinGrid.Activation.NoEdit
                        .CellAppearance.BackColor = Color.LightGray
                    End If
                End With
            Next
        End With


        Dim SQL_FCR As String = ""
        Dim SQL_ACT As String = ""
        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdDPTPROJ1, grdDPTPROJS}
            SQL_FCR = ""
            SQL_ACT = ""
            For M As Integer = 1 To 12
                For Each QAS As String In New String() {"D", "Q"}
                    With grdDPTPROJ1.DisplayLayout.Bands(0).Columns(QAS & "F" & Format(M, "00"))
                        If M <= ACT_M Then
                            If QAS = "D" Then
                                SQL_ACT &= "+M" & Format(M, "00")
                            End If
                            .CellActivation = UltraWinGrid.Activation.NoEdit
                            .CellAppearance.BackColor = Color.LightGray
                        Else
                            If QAS = "D" Then
                                SQL_FCR &= "+M" & Format(M, "00")
                            Else
                            End If
                            If M > LCK_M Then
                                .CellActivation = UltraWinGrid.Activation.AllowEdit
                            Else
                                .CellActivation = UltraWinGrid.Activation.NoEdit
                                '.CellAppearance.BackColor = Color.LightSalmon
                                .CellAppearance.BackColor = Color.Turquoise
                            End If
                        End If
                        If QAS = "Q" Then
                            Dim DCOL As UltraWinGrid.UltraGridColumn = grdDPTPROJ1.DisplayLayout.Bands(0).Columns("D" & "F" & Format(M, "00"))
                            .Header.VisiblePosition = DCOL.Header.VisiblePosition + 1
                            .Width = DCOL.Width
                            .Header.Caption = DCOL.Header.Caption
                            .Format = DCOL.Format
                        End If

                        If M > LCK_M Then
                            .Header.Appearance.BackColor2 = Color.Gold
                        Else
                            .Header.Appearance.BackColor2 = Color.Turquoise
                        End If
                    End With
                Next
            Next
        Next
       

        For Each QAS As String In New String() {"Q", "A", "S"}
            dst.Tables("DPTPROJ1").Columns(QAS & "FCR").Expression = Mid(Replace(SQL_FCR, "+M", "+" & QAS & "F"), 2)
            dst.Tables("DPTPROJ1").Columns(QAS & "ACT").Expression = Mid(Replace(SQL_ACT, "+M", "+" & QAS & "S"), 2)
        Next

        For M As Integer = 1 To 12
            dst.Tables("DPTPROJ1").Columns("AF" & Format(M, "00")).Expression _
                            = "PRICE_" & Format(M, "00") & " * QF" & Format(M, "00") & " / 1000"
        Next

        Set_Display()
        Setup_Summary()


        dst.Tables("SOTMKTCA").Rows.Clear()
        ASCMAIN1.sql = "Select MARKET_CODE, MARKET_DESC from SOTMKTC1"
        If optMarkets.Value = "S" Then
            ASCMAIN1.sql &= " where MARKET_CODE in (Select MARKET_CODE from SOTMKTC1 where CUST_CODE is Not Null)"
        End If
        For Each rowSOTMKTC1 As DataRow In ASCDATA1.GetDataTable.Rows
            dst.Tables("SOTMKTCA").Rows.Add(New Object() {rowSOTMKTC1.Item("MARKET_CODE"), rowSOTMKTC1.Item("MARKET_DESC"), "0"})
        Next
        Sort_grdColumns(grdSOTMKTCA, "MARKET_CODE")

        dst.Tables("ICTPRICX").AcceptChanges()
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim sqlDelete As String = "OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        sqlDelete &= " and OPS_YYYYPP_FC >= '" & YP(LCK_M + 1) & "'" ' YP(ACT_M + 1) & "'"
        sqlDelete &= " and OPS_YYYYPP_FC <= '" & YP(12) & "'"
        If optMarkets.Value = "S" Then
            sqlDelete &= " and MARKET_CODE in (Select MARKET_CODE from SOTMKTC1 where CUST_CODE is Not Null)"
        End If

        dst.Tables("DPTITMF1").Rows.Clear()
        For Each rowDPTPROJ1 As DataRow In dst.Tables("DPTPROJ1").Select
            For M As Integer = LCK_M + 1 To 12 ' ACT_M + 1 To 12
                If Val(rowDPTPROJ1.Item("QF" & Format(M, "00")) & "") > 0 Then
                    Dim rowDPTITMF1 As DataRow = dst.Tables("DPTITMF1").NewRow
                    rowDPTITMF1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
                    rowDPTITMF1.Item("ITEM_CODE") = rowDPTPROJ1.Item("ITEM_CODE")
                    rowDPTITMF1.Item("MARKET_CODE") = rowDPTPROJ1.Item("MARKET_CODE")
                    rowDPTITMF1.Item("OPS_YYYYPP_FC") = YP(M)
                    rowDPTITMF1.Item("FORECAST") = Val(rowDPTPROJ1.Item("QF" & Format(M, "00")) & "")
                    dst.Tables("DPTITMF1").Rows.Add(rowDPTITMF1)
                End If
            Next
        Next
        Update_Record_TDA("DPTITMF1", sqlDelete)

        sqlDelete = ""
        sqlDelete &= "OPS_YYYYPP >= '" & YP(LCK_M + 1) & "'" ' YP(ACT_M + 1) & "'"
        sqlDelete &= " and OPS_YYYYPP <= '" & YP(12) & "'"

        dst.Tables("ICTPRICA").Rows.Clear()
        For Each rowICTPRICX As DataRow In dst.Tables("ICTPRICX").Select
            For M As Integer = LCK_M + 1 To 12 ' ACT_M + 1 To 12
                If Val(rowICTPRICX.Item("PRICE_" & Format(M, "00")) & "") > 0 Then
                    Dim rowICTPRICA As DataRow = dst.Tables("ICTPRICA").NewRow
                    rowICTPRICA.Item("ITEM_CODE") = rowICTPRICX.Item("ITEM_CODE")
                    rowICTPRICA.Item("OPS_YYYYPP") = YP(M)
                    rowICTPRICA.Item("ITEM_UNIT_PRICE") = rowICTPRICX.Item("PRICE_" & Format(M, "00"))
                    dst.Tables("ICTPRICA").Rows.Add(rowICTPRICA)
                End If
            Next
        Next
        Update_Record_TDA("ICTPRICA", sqlDelete)


        If chkLock.Checked Then
            dst.Tables("DPTITMF3").Rows.Clear()
            For Each rowDPTPROJ1 As DataRow In dst.Tables("DPTPROJ1").Rows
                Dim rowDPTITMF3 As DataRow = dst.Tables("DPTITMF3").NewRow
                rowDPTITMF3.Item("OPS_YYYYPP") = YP(LCK_M + 1)
                rowDPTITMF3.Item("ITEM_CODE") = rowDPTPROJ1.Item("ITEM_CODE")
                rowDPTITMF3.Item("MARKET_CODE") = rowDPTPROJ1.Item("MARKET_CODE")
                rowDPTITMF3.Item("QPROJ") = rowDPTPROJ1.Item("QF" & Format(LCK_M + 1, "00"))
                rowDPTITMF3.Item("APROJ") = rowDPTPROJ1.Item("AF" & Format(LCK_M + 1, "00"))
                rowDPTITMF3.Item("SPROJ") = rowDPTPROJ1.Item("SF" & Format(LCK_M + 1, "00"))
                dst.Tables("DPTITMF3").Rows.Add(rowDPTITMF3)
            Next
            Update_Record_TDA("DPTITMF3")
        End If

        CommitTrans("Update Complete")

    End Sub
#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdDPTPROJS, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdDPTPROJ1, "SSSS", "Show Filter", "Show GroupBox", "Show Pins", "Show Codes", "Show Appearance Designer")
        Load_Popup_Menu(grdICTPRICX, "SS", "Show Filter", "Show GroupBox", "Item Master")
        Load_Popup_Menu(grdSOTMKTCA, "BB", "Select All", "Clear All")
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
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If
        If tlb_pop.Tools.Exists("Show Pins") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Pins"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not (grd.DisplayLayout.Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.None)
        End If
        If tlb_pop.Tools.Exists("Show Codes") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Codes"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Bands(0).Columns("BRAND_CODE").ValueList Is Nothing)
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdDPTFCSTD"
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

            Case "Show Pins"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    grd.DisplayLayout.Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.Button
                Else
                    grd.DisplayLayout.Override.FixedHeaderIndicator = UltraWinGrid.FixedHeaderIndicator.None
                End If

            Case "Show Codes"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    grd.DisplayLayout.Bands(0).Columns("BRAND_CODE").ValueList = Nothing
                    grd.DisplayLayout.Bands(0).Columns("COLLECTION_CODE").ValueList = Nothing
                Else
                    grd.DisplayLayout.Bands(0).Columns("BRAND_CODE").ValueList = grd.DisplayLayout.ValueLists("BRAND_CODE")
                    grd.DisplayLayout.Bands(0).Columns("COLLECTION_CODE").ValueList = grd.DisplayLayout.ValueLists("COLLECTION_CODE")
                End If


            Case "Show Appearance Designer"
                Using frm As New ASFAPPD1
                    frm.grd = grd
                    'frm.gcol = grd.ActiveCell.Column
                    frm.ShowDialog()
                End Using
                Exit Sub

            Case "Select All", "Clear All"
                Set_SEL(IIf(e.Tool.Key = "Select All", "1", "0"))


        End Select

        Select Case grd.Name
            'Case "grdDPTFCSTD"
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Item Master"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", Column_Values("ITEM_CODE", ITEM_CODE), "Item Master", "ICTITEM1")

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "BRAND_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        'Call Click_Command("Load", e)
            '    End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "BRAND_CODE"
            '    'Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            'Case "BRAND_CODE"
            '    If EntryMode = "" Then
            '        If Absx1.txtFor("BRAND_CODE").Text <> "" Then
            '            Call LookUp("ICTBRAN1", Absx1.txtFor("BRAND_CODE").Text)
            '            If cdr IsNot Nothing Then

            '            End If
            '        End If
            '    End If

        End Select
    End Sub

#End Region


    Sub Create_DPTPROJX(ByVal initialize As Boolean)
        If initialize Then

            ASCMAIN1.sql = "Select SOTMKTC1.MARKET_CODE, SATSSUMI.ITEM_CODE, SATSSUMI.CUST_CODE"
            ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP / SATSSUMI.ORDR_AMT_SHIP ITEM_PRICE"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", SATSSUMI.ORDR_QTY_SHIP QS" & Format(M, "00")
                ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP AS" & Format(M, "00")
                ASCMAIN1.sql &= ", SATSSUMI.ORDR_QTY_SHIP QF" & Format(M, "00")
                ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP AF" & Format(M, "00")
            Next
            ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP QSOP"
            ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP ASOP"
            ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP SSOP"
            ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP QBPE"
            ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP ABPE"
            ASCMAIN1.sql &= ", SATSSUMI.ORDR_AMT_SHIP SBPE"
            ASCMAIN1.sql &= " from SATSSUMI,SOTMKTC1 where ROWNUM < 1"
            DPTPROJX = ASCMAIN1.Temp_Table

            'ASCMAIN1.sql = "Alter Table " & DPTPROJX & " Modify MARKET_CODE NOT NULL"
            'ASCDATA1.ExecuteSQL()
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & DPTPROJX)

            ' Shipments up to and including RYP2

            ASCMAIN1.sql = "Select '?' MARKET_CODE, ITEM_CODE, CUST_CODE, 0 ITEM_PRICE"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", Sum (DECODE(OPS_YYYYPP,'" & YP(M) & "',ORDR_QTY_SHIP,0)) QS" & Format(M, "00")
                ASCMAIN1.sql &= ", Sum (DECODE(OPS_YYYYPP,'" & YP(M) & "',ORDR_AMT_SHIP / 1000,0)) AS" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 QF" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AF" & Format(M, "00")
            Next
            ASCMAIN1.sql &= ", 0 QSOP, 0 ASOP, 0 SSOP, 0 QBPE, 0 ABPE, 0 SBPE"
            ASCMAIN1.sql &= " from SATSSUMI where OPS_YYYYPP >= '" & YP(1) & "' and OPS_YYYYPP <= '" & RYP2 & "'" _
            & " and NVL(ORDR_QTY_SHIP,0) <> 0"
            ASCMAIN1.sql &= " group by ITEM_CODE, CUST_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & DPTPROJX & " " & ASCMAIN1.sql)

            ' Determine Market of Customer 

            ASCDATA1.ExecuteSQL("Update " & DPTPROJX & " DPTPROJX " _
                                 & " Set MARKET_CODE = (Select MARKET_CODE from SOTTCLS1,ARTCUST1 " _
                                 & " where SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE " _
                                 & "   and ARTCUST1.CUST_CODE = DPTPROJX.CUST_CODE)")

            ' Changeable Forecasts from RYP2 + 1 to RYP1

            ASCMAIN1.sql = "Select MARKET_CODE, ITEM_CODE, NULL CUST_CODE, 0 ITEM_PRICE"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", 0 QS" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AS" & Format(M, "00")
                ASCMAIN1.sql &= ", Sum (DECODE(OPS_YYYYPP_FC,'" & YP(M) & "',FORECAST,0)) QF" & Format(M, "00")
                'ASCMAIN1.sql &= ", Sum (DECODE(OPS_YYYYPP,'" & YP(M) & "',FORECAST * X,0)) AF" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AF" & Format(M, "00")
            Next
            ASCMAIN1.sql &= ", 0 QSOP, 0 ASOP, 0 SSOP, 0 QBPE, 0 ABPE, 0 SBPE"
            ASCMAIN1.sql &= " from DPTITMF1 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" _
            & " and FORECAST <> 0" _
            & " and OPS_YYYYPP_FC > '" & RYP2 & "' and OPS_YYYYPP_FC <= '" & RYP1 & "'"
            ASCMAIN1.sql &= " group by MARKET_CODE, ITEM_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & DPTPROJX & " " & ASCMAIN1.sql)

            ' Locked-In Forecasts from 1 to RYP2

            ASCMAIN1.sql = "Select MARKET_CODE, ITEM_CODE, NULL CUST_CODE, 0 ITEM_PRICE"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", 0 QS" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AS" & Format(M, "00")
                ASCMAIN1.sql &= ", Sum (DECODE(OPS_YYYYPP,'" & YP(M) & "',QPROJ,0)) QF" & Format(M, "00")
                ASCMAIN1.sql &= ", Sum (DECODE(OPS_YYYYPP,'" & YP(M) & "',APROJ,0)) AF" & Format(M, "00")
            Next
            ASCMAIN1.sql &= ", 0 QSOP, 0 ASOP, 0 SSOP, 0 QBPE, 0 ABPE, 0 SBPE"
            ASCMAIN1.sql &= " from DPTITMF3 " _
            & " where OPS_YYYYPP >= '" & YP(1) & "' and OPS_YYYYPP <= '" & RYP2 & "'" _
            & " and (NVL(QPROJ,0) <> 0 OR NVL(QPROJ,0) <> 0 OR NVL(QPROJ,0) <> 0)"
            ASCMAIN1.sql &= " group by MARKET_CODE, ITEM_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & DPTPROJX & " " & ASCMAIN1.sql)

            ' SOP

            ASCMAIN1.sql = "Select MARKET_CODE, ITEM_CODE, NULL CUST_CODE, 0 ITEM_PRICE"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", 0 QS" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AS" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 QF" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AF" & Format(M, "00")
            Next
            ASCMAIN1.sql &= ", SUM (QPROJ) QSOP, SUM (APROJ) ASOP, SUM (SPROJ) SSOP"
            ASCMAIN1.sql &= ", 0 QBPE, 0 ABPE, 0 SBPE"
            ASCMAIN1.sql &= " from DPTITMF3 " _
            & " where OPS_YYYYPP = '" & OPS_YYYYPP_LCK & "'" _
            & " and (NVL(QPROJ,0) <> 0 OR NVL(QPROJ,0) <> 0 OR NVL(QPROJ,0) <> 0)"
            ASCMAIN1.sql &= " group by MARKET_CODE, ITEM_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & DPTPROJX & " " & ASCMAIN1.sql)


            ' BPE

            ASCMAIN1.sql = "Select MARKET_CODE, ITEM_CODE, NULL CUST_CODE, 0 ITEM_PRICE"
            For M As Integer = 1 To 12
                ASCMAIN1.sql &= ", 0 QS" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AS" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 QF" & Format(M, "00")
                ASCMAIN1.sql &= ", 0 AF" & Format(M, "00")
            Next
            ASCMAIN1.sql &= ", 0 QSOP, 0 ASOP, 0 SSOP"
            ASCMAIN1.sql &= ", SUM (QPROJ) QBPE, SUM (APROJ) ABPE, SUM (SPROJ) SBPE"
            ASCMAIN1.sql &= " from DPTITMF3 " _
            & " where OPS_YYYYPP = '" & RYP3 & "'" _
            & " and (NVL(QPROJ,0) <> 0 OR NVL(QPROJ,0) <> 0 OR NVL(QPROJ,0) <> 0)"
            ASCMAIN1.sql &= " group by MARKET_CODE, ITEM_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & DPTPROJX & " " & ASCMAIN1.sql)


            ' Determine Customer of Market

            ASCDATA1.ExecuteSQL("Update " & DPTPROJX & " DPTPROJX " _
                     & " Set CUST_CODE = (Select CUST_CODE from SOTMKTC1 " _
                     & " where SOTMKTC1.MARKET_CODE = DPTPROJX.MARKET_CODE)")


            If optMarkets.Value = "S" Then
                ASCMAIN1.sql = "Delete from " & DPTPROJX & " where MARKET_CODE in (Select MARKET_CODE from SOTMKTC1 where CUST_CODE is Null)"
                ASCDATA1.ExecuteSQL()
            End If

            ASCMAIN1.sql = "Update " & DPTPROJX & " X Set ITEM_PRICE = (SELECT ITEM_PRICE from SOTPRIC2 where SOTPRIC2.ITEM_CODE = X.ITEM_CODE and SOTPRIC2.CUST_CODE = X.CUST_CODE) where CUST_CODE is Not NULL"
            'ASCMAIN1.sql = "Update " & DPTPROJX & " X Set ITEM_PRICE = (SELECT ITEM_RETAIL_PRICE * .6 from ICTITEM1 where ICTITEM1.ITEM_CODE = X.ITEM_CODE) where CUST_CODE is NULL"
            ASCMAIN1.sql = "Update " & DPTPROJX & " X Set ITEM_PRICE = (SELECT ITEM_RETAIL_PRICE * .5 from ICTITEM1 where ICTITEM1.ITEM_CODE = X.ITEM_CODE) where CUST_CODE is NULL"

            ASCMAIN1.sql = "Update " & DPTPROJX & " X Set MARKET_CODE = NVL(MARKET_CODE,'?'),  BRAND_CODE = NVL(BRAND_CODE,'?'),  COLLECTION_CODE = NVL(COLLECTION_CODE,'?'),  ITEM_CODE = NVL(ITEM_CODE,'?')"

        End If

    End Sub

    Sub Set_Display()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting Data")

        Dim QAS As String = optData.Value
        For Each SF As String In New String() {"S", "F"}
            For M As Integer = 1 To 12
                dst.Tables("DPTPROJ1").Columns("D" & SF & Format(M, "00")).Expression = QAS & SF & Format(M, "00")
            Next
        Next


        For Each TOT As String In New String() {"FCR", "ACT", "PROJ", "SOP", "SOPD", "BPE", "BPED"}
            dst.Tables("DPTPROJ1").Columns("D" & TOT).Expression = QAS & TOT
        Next

        For M As Integer = 1 To 12
            With grdDPTPROJ1.DisplayLayout.Bands(0)
                .Columns("QF" & Format(M, "00")).Hidden = True
                If optData.Value = "Q" And EntryMode = "E" Then
                    .Columns("DF" & Format(M, "00")).Hidden = True
                    .Columns("QF" & Format(M, "00")).Hidden = False
                Else
                    .Columns("DF" & Format(M, "00")).Hidden = False
                End If
            End With
        Next

        Dim Actuals_Message As String = "" _
        & " - " & IIf(ACT_M = 0, "No Actuals in the 12-Month Period Selected", IIf(ACT_M > 12, "Actuals in All Months Shown", "Actuals Thru " & YM(ACT_M))) _
        & " - " & IIf(LCK_M <= 0, "No Locked Months in the 12-Month Period Selected", IIf(LCK_M > 12, "All Months Shown are Locked", "Locked Thru " & YM(LCK_M)))
        grdDPTPROJ1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdDPTPROJ1.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        grdDPTPROJ1.Text = "Projected Shipments (" & optData.Text & ")" & Actuals_Message

        grdDPTPROJS.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
        Dim sqld As String = grdDPTPROJS.Tag
        grdDPTPROJS.Text = "Projected Shipments Summary by " & Mid(sqld, 2) & " (" & optData.Text & ")" & Actuals_Message

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub optData_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optData.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Display()
    End Sub

    Private Sub cmdSummary_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSummary.Click
        Setup_Summary()
    End Sub

    Sub Setup_Summary()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting up Summary")

        If (chkSumMarket.Checked And chkSumBrand.Checked And chkSumCollection.Checked And chkSumItem.Checked) _
        Or (Not chkSumMarket.Checked And Not chkSumBrand.Checked And Not chkSumCollection.Checked And Not chkSumItem.Checked) Then
            splMain.Panel1Collapsed = True
        Else
            EnforceConstraints(False)
            If dst.Relations.Contains("DPTPROJS_DPTPROJ1") Then
                dst.Relations.Remove("DPTPROJS_DPTPROJ1")
                dst.Tables("DPTPROJ1").Constraints.Remove("DPTPROJS_DPTPROJ1")
                dst.Tables("DPTPROJS").Constraints.Remove(dst.Tables("DPTPROJS").Constraints(0))
                'dst.Tables("DPTPROJS").Constraints
                'Stop
            End If

            Dim SQL As String = ""
            Dim SQLD As String = ""
            If chkSumMarket.Checked Then SQL &= ",MARKET_CODE" : SQLD &= "," & chkSumMarket.Text
            If chkSumBrand.Checked Then SQL &= ",BRAND_CODE" : SQLD &= "," & chkSumBrand.Text
            If chkSumCollection.Checked Then SQL &= ",COLLECTION_CODE" : SQLD &= "," & chkSumCollection.Text
            If chkSumItem.Checked Then SQL &= ",ITEM_CODE" : SQLD &= "," & chkSumItem.Text
            grdDPTPROJS.Tag = SQLD
            grdDPTPROJS.Text = "Projected Shipments Summary by " & Mid(SQLD, 2) & " (" & optData.Text & ")"

            Create_Relation("DPTPROJS", "DPTPROJ1", Mid(SQL, 2))

            'Dim NSC() As String = {"MARKET_CODE", "ITEM_CODE", "BRAND_CODE", "COLLECTION_CODE", "ITEM_DESC"}

            For Each DC As DataColumn In dst.Tables("DPTPROJS").Columns
                If DC.DataType.Equals(GetType(System.String)) _
                Or DC.ColumnName = "ITEM_RETAIL_PRICE" Or DC.ColumnName = "NET_PRICE" Or DC.ColumnName = "DISC_PCT" Then
                Else
                    DC.Expression = "SUM(CHILD." & DC.ColumnName & ")"
                End If
                'If Not NSC.Contains(DC.ColumnName) Then

                'End If
            Next

            dst.Tables("DPTPROJS").Rows.Clear()
            For Each ROW As DataRow In ASCDATA1.SelectDistinct(dst.Tables("DPTPROJ1").Select, Split(Mid(SQL, 2), ",")).Rows
                Dim rowDPTPROJS As DataRow = dst.Tables("DPTPROJS").NewRow
                For Each COLUMN_NAME As String In Mid(SQL, 2).Split(",")
                    rowDPTPROJS.Item(COLUMN_NAME) = ROW.Item(COLUMN_NAME)
                Next
                dst.Tables("DPTPROJS").Rows.Add(rowDPTPROJS)
                'Stop
            Next
            EnforceConstraints(True)

            grdDPTPROJ1.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

            splMain.SplitterDistance = splMain.Height / 2
            splMain.Panel1Collapsed = False
        End If
        Setup_grdDPTPROJ1()


        Dim vcols(grdDPTPROJ1.DisplayLayout.Bands(0).Columns.Count - 1) As String
        For Each gcol As UltraWinGrid.UltraGridColumn In grdDPTPROJ1.DisplayLayout.Bands(0).Columns
            If gcol.Header.VisiblePosition >= 0 Then
                vcols(gcol.Header.VisiblePosition) = gcol.Key
            End If
        Next
        Dim vp As Integer = -1
        For i As Integer = 0 To vcols.Length - 1
            If dst.Tables("DPTPROJS").Columns.Contains(vcols(i)) Then
                vp += 1
                grdDPTPROJS.DisplayLayout.Bands(0).Columns(vcols(i)).Header.VisiblePosition = vp
            End If
        Next

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdDPTPROJS_AfterColRegionScroll(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ColScrollRegionEventArgs) Handles grdDPTPROJS.AfterColRegionScroll

    End Sub

    Private Sub grdDPTPROJS_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDPTPROJS.AfterRowActivate
        Setup_grdDPTPROJ1()
    End Sub

    Sub Setup_grdDPTPROJ1()
        Dim dvw As DataView = DirectCast(grdDPTPROJ1.DataSource, DataTable).DefaultView
        If splMain.Panel1Collapsed Then
            dvw.RowFilter = ""
        Else
            Dim SQL As String = ""

            If grdDPTPROJS.ActiveRow IsNot Nothing Then
                If grdDPTPROJS.ActiveRow.IsGroupByRow Then
                    Dim gbyrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grdDPTPROJS.ActiveRow, UltraWinGrid.UltraGridGroupByRow)
                    Do While gbyrow IsNot Nothing
                        Dim COLUMN_NAME As String = gbyrow.Column.Key
                        SQL &= " AND " & COLUMN_NAME & IIf(gbyrow.Value & "" = "", " IS NULL", " = '" & gbyrow.Value & "'")
                        gbyrow = gbyrow.ParentRow
                    Loop
                ElseIf grdDPTPROJS.ActiveRow.IsDataRow Then
                    For Each COLUMN_NAME As String In New String() {"BRAND_CODE", "COLLECTION_CODE", "MARKET_CODE", "ITEM_CODE"}
                        If grdDPTPROJS.ActiveRow.Cells(COLUMN_NAME).Value & "" <> "" Then
                            SQL &= " AND " & COLUMN_NAME & " = '" & grdDPTPROJS.ActiveRow.Cells(COLUMN_NAME).Value & "'"
                        End If
                    Next
                End If
            End If

            dvw.RowFilter = Mid(SQL, 5)

        End If
    End Sub

    Sub Format_grd(ByVal grd As UltraWinGrid.UltraGrid)
        grd.DisplayLayout.Override.GroupByRowDescriptionMask = "[value]"

        With grd.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"MARKET_CODE", "BRAND_CODE", "COLLECTION_CODE", "ITEM_CODE", "ITEM_DESC"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.Beige
                If COLUMN_NAME <> "ITEM_DESC" Then
                    .Columns(COLUMN_NAME).Width = 80
                End If
            Next
            For Each COLUMN_NAME As String In New String() _
                {"ITEM_RETAIL_PRICE", "NET_PRICE", "DISC_PCT"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightGray
            Next

            For M As Integer = 1 To 12
                For Each SF As String In New String() {"F", "S"}
                    For Each QAS As String In New String() {"D", "Q"}
                        If grd.Name = "grdDPTPROJ1" Or QAS = "D" Then
                            COLUMN_NAME = QAS & SF & Format(M, "00")
                            .Columns(COLUMN_NAME).Width = 80
                            '.Columns(COLUMN_NAME).Group = G
                            .Columns(COLUMN_NAME).Format = "###,##0"
                            Create_Summary(grd, COLUMN_NAME)

                            If SF = "F" Then
                                .Columns(COLUMN_NAME).Header.Caption = "F/C"
                                '.Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightGray
                            Else
                                .Columns(COLUMN_NAME).Header.Caption = "Act"
                                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightGreen
                            End If
                        End If

                    Next
                Next
            Next

            For Q As Integer = 0 To 4
                For Each SF As String In New String() {"F", "S"}
                    COLUMN_NAME = "D" & SF & "Q" & Format(Q, "0")
                    .Columns(COLUMN_NAME).Width = 80
                    '.Columns(COLUMN_NAME).Group = G
                    .Columns(COLUMN_NAME).Format = "###,##0"
                    Create_Summary(grd, COLUMN_NAME)

                    Dim HDG As String = ""
                    If Q = 0 Then
                        HDG = "Total"
                    Else
                        HDG = "Q" & Format(Q, "0")
                    End If
                    If SF = "F" Then
                        .Columns(COLUMN_NAME).Header.Caption = HDG & vbCrLf & "F/C"
                        '.Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightGray
                    Else
                        .Columns(COLUMN_NAME).Header.Caption = HDG & vbCrLf & "Act"
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightGreen
                    End If
                Next
            Next

            For Each TOT As String In New String() {"FCR", "ACT", "PROJ", "SOP", "SOPD", "BPE", "BPED"}
                COLUMN_NAME = "D" & TOT
                .Columns(COLUMN_NAME).Width = 80
                '.Columns(COLUMN_NAME).Group = G
                .Columns(COLUMN_NAME).Format = "###,##0"
                Create_Summary(grd, COLUMN_NAME)
                '.Columns(COLUMN_NAME).Header.Caption = ""
                .Columns(COLUMN_NAME).CellAppearance.BackColor = Color.LightGreen
            Next

        End With

        Create_Summary(grd, "ITEM_CODE", "Count")

    End Sub

    Sub Get_LCK()
        ASCMAIN1.sql = "Select Max (OPS_YYYYPP) from DPTITMF3"
        OPS_YYYYPP_LCK = ASCDATA1.GetDataValue
        OPS_YYYYPP_LCK_LEGEND = ASCMAIN1.Get_Legend(OPS_YYYYPP_LCK, False, True)
        UltraExplorerBar1.Groups("Screen Control").Items("Unlock MMM'YY").Text = "Unlock " & OPS_YYYYPP_LCK_LEGEND
    End Sub

    Sub Unlock_Period()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Deleteing Locked Forecast Data for " & OPS_YYYYPP_LCK_LEGEND)

        ASCMAIN1.sql = "Delete from DPTITMF3 where OPS_YYYYPP = '" & OPS_YYYYPP_LCK & "'"
        ASCDATA1.ExecuteSQL()


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        MsgBox("Period " & OPS_YYYYPP_LCK_LEGEND & " has been successfully Unlocked", _
                MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Private Sub tabMain_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()

        With UltraExplorerBar1
            .Groups("Summarize By").Visible = ScreenMode And (EntryMode = "L") And (tabMain.SelectedTab.Key = "Projections")
            .Groups("Data Options").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Projections")

            .Groups("Prices").Visible = ScreenMode And (tabMain.SelectedTab.Key = "Prices") And (EntryMode = "E")
        End With
    End Sub

    Private Sub grdICTPRICX_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPRICX.AfterCellUpdate
        If e.Cell.Row.IsFilterRow Then Exit Sub

        If chkCopyPrices.Checked And chkCopyPrices.Tag = "" Then
            Dim MM As Integer = Val(Mid(e.Cell.Column.Key, 7, 2))
            If MM < 12 Then
                chkCopyPrices.Tag = "X"
                For M As Integer = MM + 1 To 12
                    e.Cell.Row.Cells("PRICE_" & Format(M, "00")).Value = e.Cell.Value
                Next
                chkCopyPrices.Tag = ""
            End If
        End If
    End Sub

    Private Sub grdICTPRICX_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTPRICX.AfterRowUpdate
        'For Each rowDPTPROJ1 As DataRow In dst.Tables("DPTPROJ1").Select("")
        '    For M As Integer = ACT_M + 1 To 12
        '        rowDPTPROJ1.Item("AF" & Format(M, "00")) _
        '        = Val(rowDPTPROJ1.Item("QF" & Format(M, "00")) & "") _
        '        * Val(e.Row.Cells("PRICE_" & Format(M, "00")).Value & "")
        '    Next
        'Next
    End Sub

    Private Sub grdICTPRICX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPRICX.InitializeLayout

    End Sub

    Private Sub cmdSummaryClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSummaryClear.Click
        Clear_Summary()
    End Sub

    Sub Clear_Summary()
        chkSumBrand.Checked = False
        chkSumCollection.Checked = False
        chkSumItem.Checked = False
        chkSumMarket.Checked = False
        Setup_Summary()
    End Sub

    Private Sub cmdUploadFromTemplate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUploadFromTemplate.Click
        If cmbMARKET_CODE.Value & "" = "" Then
            MsgBox("Select a Market Code prior to Uploading a Forecast from a Template", MsgBoxStyle.OkOnly, "Cannot Upload without first Selecting a Market")
            Exit Sub
        End If

        If LCK_M >= 12 Then
            MsgBox("There are no Un-Locked Months in the 12-Months View Selected", MsgBoxStyle.OkOnly, "Cannot Upload without first Selecting a Market")
            Exit Sub
        End If

        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            openFileDialog1.Title = "Select a Template containing Forecasts to Upload"
            openFileDialog1.Filter = "Excel files (*.xls)|*.xls"
            openFileDialog1.FilterIndex = 2
            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim FILENAME As String = openFileDialog1.FileName
                Try
                    Dim strConnection As String = "Provider=Microsoft.Jet.OleDb.4.0;" & _
                "data source=" & FILENAME & ";" & _
                "Extended Properties=Excel 8.0;"
                    Dim objConnection As New System.Data.OleDb.OleDbConnection(strConnection)
                    objConnection.Open()
                    Dim dbSchema As DataTable = objConnection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, Nothing)
                    If dbSchema.Rows.Count = 0 Then
                        MsgBox("No Sheets Found")
                        Exit Sub
                    End If
                    Dim SHEET_NAME As String = dbSchema.Rows(0).Item("TABLE_NAME")
                    'If SHEET_NAME.EndsWith("$") Or (SHEET_NAME.StartsWith("'") And SHEET_NAME.EndsWith("$'")) Then
                    '    If SHEET_NAME.StartsWith("'") Then
                    '        SHEET_NAME = Mid(SHEET_NAME, 2, Len(SHEET_NAME) - 2)
                    '    End If
                    '    SHEET_NAME = Mid(SHEET_NAME, 1, Len(SHEET_NAME) - 1)
                    'End If

                    Dim strSQL As String = "SELECT * FROM [" & SHEET_NAME & "]"
                    Dim objCommand As New System.Data.OleDb.OleDbCommand(strSQL, objConnection)
                    Dim objAdapter As New System.Data.OleDb.OleDbDataAdapter(strSQL, objConnection)
                    Dim dt As New DataTable
                    objAdapter.FillSchema(dt, SchemaType.Source)
                    objAdapter.Fill(dt)
                    objConnection.Close()

                    'If dt.Columns.Count <> 13 Then
                    '    MsgBox("Template does not have 13 columns (Item Code and 12 months of Forecast)")
                    '    Exit Sub
                    'End If

                    Dim tblBAD_ITEMS As New DataTable
                    tblBAD_ITEMS.Columns.Add("ITEM_CODE")

                    For Each row As DataRow In dt.Rows
                        Dim ITEM_CODE As String = row.Item(0) & ""
                        If ITEM_CODE <> "" Then
                            Dim bad_item As Boolean = False

                            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                            If rowICTITEM1 Is Nothing Then
                                rowICTITEM1 = LookUp("ICTITEM1", "0" & ITEM_CODE)
                                If rowICTITEM1 IsNot Nothing Then
                                    ITEM_CODE = "0" & ITEM_CODE
                                End If
                            End If

                            Dim rowICTPRICX As DataRow = dst.Tables("ICTPRICX").Rows.Find(ITEM_CODE)
                            If rowICTPRICX Is Nothing Then
                                If rowICTITEM1 Is Nothing Then
                                    bad_item = True
                                    'tblBAD_ITEMS.Rows.Add(New String() {ITEM_CODE})
                                End If
                            End If

                            If bad_item Then
                                tblBAD_ITEMS.Rows.Add(New String() {ITEM_CODE})
                            End If
                        End If
                    Next

                    If tblBAD_ITEMS.Rows.Count <> 0 Then
                        Using frm As New ASFMSGBF
                            frm.Show_grd(tblBAD_ITEMS, Me, "Invalid Item Codes")
                            Exit Sub
                        End Using
                    End If

                    Dim MM As Integer = LCK_M + 1 ' Val(Mid(ASCMAIN1.CYP, 5, 2))
                    If MsgBox("This procedure will Replace all of the Unit Forecast values " _
                              & vbCrLf & " from " & YM(MM) & " to " & YM(12) & " for Market " & cmbMARKET_CODE.Value _
                              & vbCrLf & " with the Forecast Quantities found in the Selected Template" _
                              & vbCrLf & IIf(optReplace.Value = "M", " for the Entire Market", " for only the Items found in the Spreadsheet") _
                              & vbCrLf _
                              & vbCrLf & "Do you wish to Proceed?", _
                              MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                    If optReplace.Value = "M" Then
                        For Each rowDPTPROJ1 As DataRow In dst.Tables("DPTPROJ1").Select("MARKET_CODE = '" & cmbMARKET_CODE.Value & "'")
                            For M As Integer = MM To 12
                                rowDPTPROJ1.Item("QF" & Format(M, "00")) = 0
                            Next
                        Next
                    End If

                    Dim rows_updated As Integer = 0

                    For Each row As DataRow In dt.Rows
                        Dim ITEM_CODE As String = row.Item(0) & ""

                        If ITEM_CODE <> "" Then
                            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                            If rowICTITEM1 Is Nothing Then
                                rowICTITEM1 = LookUp("ICTITEM1", "0" & ITEM_CODE)
                                If rowICTITEM1 IsNot Nothing Then
                                    ITEM_CODE = "0" & ITEM_CODE
                                End If
                            End If

                            Dim rowICTPRICX As DataRow = dst.Tables("ICTPRICX").Rows.Find(ITEM_CODE)
                            If rowICTPRICX Is Nothing Then
                                rowICTPRICX = dst.Tables("ICTPRICX").NewRow
                                rowICTPRICX.Item("ITEM_CODE") = ITEM_CODE
                                rowICTPRICX.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                                rowICTPRICX.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                                rowICTPRICX.Item("PG_STAT_FACTOR") = rowICTITEM1.Item("PG_STAT_FACTOR")
                                rowICTPRICX.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                                Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", rowICTITEM1.Item("COLLECTION_CODE") & "")
                                If rowICTCOLL1 IsNot Nothing Then
                                    rowICTPRICX.Item("BRAND_CODE") = rowICTCOLL1.Item("BRAND_CODE")
                                End If
                                dst.Tables("ICTPRICX").Rows.Add(rowICTPRICX)
                            End If
                            Dim rowDPTPROJ1 As DataRow = dst.Tables("DPTPROJ1").Rows.Find(New String() {cmbMARKET_CODE.Value, ITEM_CODE})
                            If rowDPTPROJ1 Is Nothing Then
                                rowDPTPROJ1 = dst.Tables("DPTPROJ1").NewRow
                                rowDPTPROJ1.Item("MARKET_CODE") = cmbMARKET_CODE.Value
                                For Each COLUMN_NAME As String In New String() _
                                    {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "PG_STAT_FACTOR", "ITEM_RETAIL_PRICE", "BRAND_CODE"}
                                    rowDPTPROJ1.Item(COLUMN_NAME) = rowICTPRICX.Item(COLUMN_NAME)
                                Next
                                dst.Tables("DPTPROJ1").Rows.Add(rowDPTPROJ1)
                            End If

                            For M As Integer = MM To 12
                                rowDPTPROJ1.Item("QF" & Format(M, "00")) = Val(row.Item(M) & "")
                                'rowDPTPROJ1.Item("AF" & Format(M, "00")) = Val(row.Item(M) & "") & Val(rowDPTPROJ1.Item("PRICE_" & Format(M, "00")) & "")
                            Next

                            rows_updated += 1
                        End If

                    Next

                    MsgBox("Forecasts from Template have been Loaded (" & CStr(rows_updated) & " rows Updated)", MsgBoxStyle.OkOnly, "Success")

                Catch ex As Exception
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error trying to process Template")
                End Try

            End If
        End Using

    End Sub

    Private Sub chkShowFCOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowFCOnly.CheckedChanged

        If chkShowFCOnly.Checked Then
            Clear_Summary()
        End If

        cmdSummary.Enabled = Not chkShowFCOnly.Checked


        'grdDPTPROJ1.Visible = False

        With grdDPTPROJ1.DisplayLayout.Bands(0)
            For Q As Integer = 0 To 4
                .Columns("DSQ" & Format(Q, "0")).Hidden = chkShowFCOnly.Checked
                If Q <> 0 Then
                    .Columns("DFQ" & Format(Q, "0")).Hidden = chkShowFCOnly.Checked
                End If
            Next
            For M As Integer = 1 To 12
                .Columns("DS" & Format(M, "00")).Hidden = chkShowFCOnly.Checked
            Next
            For Each TOT As String In New String() {"FCR", "ACT", "PROJ", "SOP", "SOPD", "BPE", "BPED"}
                .Columns("D" & TOT).Hidden = chkShowFCOnly.Checked
            Next
        End With

        'grdDPTPROJ1.Visible = True

        Set_Upload()
    End Sub

    Sub Set_Upload()
        UltraExplorerBar1.Groups("Upload from Template").Visible = chkShowFCOnly.Checked And EntryMode = "E"
        lblMARKET_CODE.Visible = chkShowFCOnly.Checked And EntryMode = "E"
        cmbMARKET_CODE.Visible = chkShowFCOnly.Checked And EntryMode = "E"
        cmdUploadFromTemplate.Visible = chkShowFCOnly.Checked And EntryMode = "E"
        lblReplace.Visible = chkShowFCOnly.Checked And EntryMode = "E"
        optReplace.Visible = chkShowFCOnly.Checked And EntryMode = "E"
    End Sub

    Private Sub grdDPTPROJ1_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdDPTPROJ1.AfterRowUpdate
    End Sub

    Private Sub grdDPTPROJ1_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdDPTPROJ1.BeforeRowUpdate
        'For M As Integer = ACT_M + 1 To 12
        '    e.Row.Cells("AF" & Format(M, "00")).Value _
        '    = Val(e.Row.Cells("QF" & Format(M, "00")).Value & "") _
        '    * Val(e.Row.Cells("PRICE_" & Format(M, "00")).Value)
        'Next
    End Sub

    Private Sub grdDPTPROJ1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdDPTPROJ1.InitializeLayout

    End Sub

    Private Sub grdICTPRICX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTPRICX.InitializeRow
        If e.Row.IsDataRow Then
            Try
                Dim rowICTPRICX As DataRow = dst.Tables("ICTPRICX").Rows.Find(e.Row.Cells("ITEM_CODE").Value)
                If rowICTPRICX IsNot Nothing Then
                    For I As Integer = 2 To 12
                        If Val(rowICTPRICX.Item("PRICE_" & Format(I, "00")) & "") _
                        <> Val(rowICTPRICX.Item("PRICE_" & Format(I - 1, "00")) & "") Then
                            e.Row.Cells("PRICE_" & Format(I, "00")).Appearance.BackColor = Color.Yellow
                        Else
                            e.Row.Cells("PRICE_" & Format(I, "00")).Appearance.BackColor = Color.Empty
                        End If
                    Next
                    If rowICTPRICX.RowState = DataRowState.Modified Then
                        For I As Integer = 1 To 12
                            If Val(rowICTPRICX.Item("PRICE_" & Format(I, "00"), DataRowVersion.Current) & "") _
                            <> Val(rowICTPRICX.Item("PRICE_" & Format(I, "00"), DataRowVersion.Original) & "") Then
                                e.Row.Cells("PRICE_" & Format(I, "00")).Appearance.ForeColor = Color.Red
                            Else
                                e.Row.Cells("PRICE_" & Format(I, "00")).Appearance.ForeColor = Color.Empty
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub grdDPTPROJS_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdDPTPROJS.InitializeLayout

    End Sub

    Private Sub grdDPTPROJS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTPROJS.InitializeRow

    End Sub

    Private Sub grdDPTPROJ1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdDPTPROJ1.InitializeRow
        If e.Row.IsDataRow Then
            Try
                Dim rowDPTPROJ1 As DataRow = dst.Tables("DPTPROJ1").Rows.Find _
                    (New String() {e.Row.Cells("MARKET_CODE").Value, e.Row.Cells("ITEM_CODE").Value})
                If rowDPTPROJ1 IsNot Nothing Then
                    If rowDPTPROJ1.RowState = DataRowState.Added Then
                        For I As Integer = 1 To 12
                            e.Row.Cells("QF" & Format(I, "00")).Appearance.ForeColor = Color.Red
                        Next

                    End If
                    If rowDPTPROJ1.RowState = DataRowState.Modified Then
                        For I As Integer = 1 To 12
                            If Val(rowDPTPROJ1.Item("QF" & Format(I, "00"), DataRowVersion.Current) & "") _
                            <> Val(rowDPTPROJ1.Item("QF" & Format(I, "00"), DataRowVersion.Original) & "") Then
                                e.Row.Cells("QF" & Format(I, "00")).Appearance.ForeColor = Color.Red
                            Else
                                e.Row.Cells("QF" & Format(I, "00")).Appearance.ForeColor = Color.Empty
                            End If
                        Next
                    End If
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub cmdAddItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAddItem.Click
        'If cmbAddMARKET_CODE.Value & "" = "" Or txtAddITEM_CODE.Text = "" Then
        '    MsgBox("You Must Specify a Market Code and an Item Code to Add", MsgBoxStyle.OkOnly, "Cannot Add")
        '    Exit Sub
        'End If
        'Dim rowSOTMKTC1 As DataRow = LookUp("SOTMKTC1", cmbAddMARKET_CODE.Value)
        'If rowSOTMKTC1 Is Nothing Then
        '    MsgBox("You Must Specify a Valid Market Code", MsgBoxStyle.OkOnly, "Cannot Add")
        '    Exit Sub
        'End If
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", txtAddITEM_CODE.Text)
        If rowICTITEM1 Is Nothing Then
            MsgBox("You Must Specify a Valid Item Code", MsgBoxStyle.OkOnly, "Cannot Add")
            Exit Sub
        End If

        Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", rowICTITEM1.Item("COLLECTION_CODE") & "")

        If dst.Tables("SOTMKTCA").Select("SEL = '1'").Length = 0 Then
            MsgBox("You Must Select at least 1 Market Code", MsgBoxStyle.OkOnly, "Cannot Add")
            Exit Sub
        End If

        For Each rowSOTMKTCA As DataRow In dst.Tables("SOTMKTCA").Select("SEL = '1'")
            Dim rowDPTPROJ1 As DataRow = dst.Tables("DPTPROJ1").Rows.Find _
                (New String() {rowSOTMKTCA.Item("MARKET_CODE"), rowICTITEM1.Item("ITEM_CODE")})

            If rowDPTPROJ1 IsNot Nothing Then
                MsgBox("Record Already Exists for Item " & rowICTITEM1.Item("ITEM_CODE") & " in Market " & rowSOTMKTCA.Item("MARKET_CODE"), MsgBoxStyle.OkOnly, "Cannot Add")
                Exit Sub
            End If
        Next

        EnforceConstraints(False)

        For Each rowSOTMKTCA As DataRow In dst.Tables("SOTMKTCA").Select("SEL = '1'")
            Dim rowDPTPROJ1 As DataRow = dst.Tables("DPTPROJ1").Rows.Find _
                (New String() {rowSOTMKTCA.Item("MARKET_CODE"), rowICTITEM1.Item("ITEM_CODE")})

            rowDPTPROJ1 = dst.Tables("DPTPROJ1").NewRow
            rowDPTPROJ1.Item("MARKET_CODE") = rowSOTMKTCA.Item("MARKET_CODE")
            rowDPTPROJ1.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE")
            rowDPTPROJ1.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            rowDPTPROJ1.Item("PG_STAT_FACTOR") = rowICTITEM1.Item("PG_STAT_FACTOR")
            rowDPTPROJ1.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
            rowDPTPROJ1.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
            If rowICTCOLL1 IsNot Nothing Then
                rowDPTPROJ1.Item("BRAND_CODE") = rowICTCOLL1.Item("BRAND_CODE")
            End If
            dst.Tables("DPTPROJ1").Rows.Add(rowDPTPROJ1)

            Dim rowICTPRICX As DataRow = dst.Tables("ICTPRICX").Rows.Find(rowICTITEM1.Item("ITEM_CODE"))
            If rowICTPRICX Is Nothing Then
                rowICTPRICX = dst.Tables("ICTPRICX").NewRow
                rowICTPRICX.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE")
                rowICTPRICX.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                rowICTPRICX.Item("PG_STAT_FACTOR") = rowICTITEM1.Item("PG_STAT_FACTOR")
                rowICTPRICX.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                rowICTPRICX.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                If rowICTCOLL1 IsNot Nothing Then
                    rowICTPRICX.Item("BRAND_CODE") = rowICTCOLL1.Item("BRAND_CODE")
                End If
                dst.Tables("ICTPRICX").Rows.Add(rowICTPRICX)
            End If

        Next


        EnforceConstraints(True)

        ''grdDPTPROJ1.Selected.Rows.Clear()
        'For Each grow As UltraWinGrid.UltraGridRow In grdDPTPROJ1.Rows
        '    If grow.Cells("MARKET_CODE").Value = rowSOTMKTC1.Item("MARKET_CODE") And grow.Cells("ITEM_CODE").Value = rowICTITEM1.Item("ITEM_CODE") Then
        '        grdDPTPROJ1.ActiveRow = grow
        '        Exit For
        '    End If
        'Next
        MsgBox("Item " & rowICTITEM1.Item("ITEM_CODE") & " was Successfully Added to (" & CStr(dst.Tables("SOTMKTCA").Select("SEL = '1'").Length) & ") Market(s) Selected", MsgBoxStyle.OkOnly, "Success")

    End Sub

    Private Sub txtAddITEM_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAddITEM_CODE.ValueChanged

    End Sub

    Private Sub UltraLabel5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraLabel5.Click

    End Sub

    Sub Set_SEL(ByVal SEL As String)
        For Each row As DataRow In dst.Tables("SOTMKTCA").Select("")
            row.Item("SEL") = SEL
        Next
    End Sub

    Sub Export_Summary()
        Dim COLS As New List(Of String)
        With grdDPTPROJS.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() _
                {"ITEM_CODE", "ITEM_DESC", "ITEM_RETAIL_PRICE", "NET_PRICE", "DISC_PCT"}
                If Not .Columns(COLUMN_NAME).Hidden Then
                    .Columns(COLUMN_NAME).Hidden = True
                    COLS.Add(COLUMN_NAME)
                End If
            Next

            If Not chkSumBrand.Checked And Not .Columns("BRAND_CODE").Hidden Then
                .Columns("BRAND_CODE").Hidden = True
                COLS.Add("BRAND_CODE")
            End If
            If Not chkSumCollection.Checked And Not .Columns("COLLECTION_CODE").Hidden Then
                .Columns("COLLECTION_CODE").Hidden = True
                COLS.Add("COLLECTION_CODE")
            End If
            If Not chkSumMarket.Checked And Not .Columns("MARKET_CODE").Hidden Then
                .Columns("MARKET_CODE").Hidden = True
                COLS.Add("MARKET_CODE")
            End If
            If Not chkSumItem.Checked And Not .Columns("ITEM_CODE").Hidden Then
                .Columns("ITEM_CODE").Hidden = True
                COLS.Add("ITEM_CODE")
            End If

            For I As Integer = 1 To 12
                If I <= LCK_M Then
                    COLUMN_NAME = "DF" & Format(I, "00")
                Else
                    COLUMN_NAME = "DS" & Format(I, "00")
                End If
                .Columns(COLUMN_NAME).Hidden = True
                COLS.Add(COLUMN_NAME)
            Next

            For Each COLUMN_NAME As String In New String() _
                {"DSOPD", "DBPED" _
                 , "DSQ1", "DSQ2", "DSQ3", "DSQ4", "DSQ0", "DFQ1", "DFQ2", "DFQ3", "DFQ4", "DFQ0" _
                 , "DFCR", "DACT", "DSOP", "DBPE"}
                If Not .Columns(COLUMN_NAME).Hidden Then
                    .Columns(COLUMN_NAME).Hidden = True
                    COLS.Add(COLUMN_NAME)
                End If
            Next

            .ColHeadersVisible = False
            Export_to_Excel(grdDPTPROJS, , , , "0")
            .ColHeadersVisible = True
            For Each COLUMN_NAME As String In COLS
                .Columns(COLUMN_NAME).Hidden = False
            Next
        End With
    End Sub
End Class