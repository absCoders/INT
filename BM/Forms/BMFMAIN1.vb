Public Class BMFMAIN1
    Dim BM_PROD_ITEM As String
    Dim BM_ISSUE_NO As String
    Dim rowBMTMAIN1 As DataRow
    Dim rowBMTMAIN2 As DataRow
    Dim rowICTCOSTC_product As DataRow

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "BMFMAINI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")

        With dst

            ASCMAIN1.sql = "Select BMTMAIN1.*" _
                & " from BMTMAIN1 where BM_PROD_ITEM = :PARM1"
            Create_TDA(.Tables.Add, "BMTMAIN1", "**", 0, True, "V", 1)
            .Tables("BMTMAIN1").Columns.Add("IMAGE", GetType(System.Byte()))

            ASCMAIN1.sql = "Select BMTMAIN2.*" _
                & " from BMTMAIN2 where BM_PROD_ITEM = :PARM1 and BM_ISSUE_NO = :PARM2"
            Create_TDA(.Tables.Add, "BMTMAIN2", "**", 0, True, "VN", 2)

            ASCMAIN1.sql = "Select BMTMAIN1.*" & vbCrLf _
                & ",ICTITEM1.ITEM_DESC,ICTITEM1.ITEM_UOM,ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ",ICTITEM1.ITEM_SNU_CODE,ICTITEM1.ITEM_PLAN_MAKE_BUY" _
                & " from BMTMAIN1,ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = BMTMAIN1.BM_PROD_ITEM"
            Create_TDA(.Tables.Add, "BMTMAINX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select BMTMAIN2.*" & vbCrLf _
                & " from BMTMAIN2" & vbCrLf
            Create_TDA(.Tables.Add, "BMTMAINY", "**", 0, False, "", 2)

            Create_Relation("BMTMAINX", "BMTMAINY", "BM_PROD_ITEM")

            ASCMAIN1.sql = "Select BMTMAIN3.*" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_STD" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_WASTE_PCT" & vbCrLf _
                & ", ICTITEM1.ITEM_PLAN_WASTE_PCT" & vbCrLf _
                & ", ICTITEM1.VEND_CODE" & vbCrLf _
                & ", ICTITEM1.VEND_ITEM_CODE" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_VCOST" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_LANDG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOOLG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_OVRHD" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOTAL" & vbCrLf _
                & " from BMTMAIN3,ICTITEM1,ICTCOSTC" & vbCrLf _
                & " where BMTMAIN3.BM_PROD_ITEM = :PARM1" & vbCrLf _
                & " and BMTMAIN3.BM_ISSUE_NO = :PARM2" & vbCrLf _
                & " and ICTITEM1.ITEM_CODE = BMTMAIN3.BM_COMP_ITEM" & vbCrLf _
                & " and ICTCOSTC.ITEM_CODE (+) = BMTMAIN3.BM_COMP_ITEM"
            Create_TDA(.Tables.Add, "BMTMAIN3", "**", 0, True, "VN", 3)
            Dim CALC As String = "ISNULL(BM_QTY_PER_ASSY,0) * ISNULL(?,0) * (1 + ISNULL(ITEM_COST_WASTE_PCT,0)/100)"
            With .Tables("BMTMAIN3")
                .Columns.Add("EXT_COST", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_STD"))
                .Columns.Add("VCOST", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_VCOST"))
                .Columns.Add("LANDG", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_LANDG"))
                .Columns.Add("TOOLG", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_TOOLG"))
                .Columns.Add("OVRHD", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_OVRHD"))
                .Columns.Add("TOTAL", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_TOTAL"))
                .Columns.Add("QTY_ON_HAND", GetType(System.Int32))
                .Columns.Add("QTY_ONPO", GetType(System.Int32))
                .Columns.Add("QTY_PLAN", GetType(System.Int32))
                .Columns.Add("QTY_OPEN", GetType(System.Int32))
                .Columns.Add("QTY_PICK", GetType(System.Int32))
                .Columns.Add("QTY_COMM", GetType(System.Int32))
                .Columns.Add("QTY_OPEN_PICK", GetType(System.Int32), "ISNULL(QTY_OPEN,0)+ISNULL(QTY_PICK,0)")
                .Columns.Add("QTY_AVA", GetType(System.Int32), "ISNULL(QTY_ON_HAND,0)+ISNULL(QTY_OPEN,0)+ISNULL(QTY_PLAN,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_OPEN,0)-ISNULL(QTY_PICK,0)")
                .Columns.Add("BM_COMPONENT_SORT")
            End With

            ASCMAIN1.sql = "Select ICTCOSTC.*" _
                & " from ICTCOSTC where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTC", "**", 0, False, "V", 1)

            '                & ", ICTCOSTC.ITEM_COST_TOTAL TOTAL" & vbCrLf _

            ASCMAIN1.sql = "Select 'X' COST_TYPE" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_VCOST VCOST" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_LANDG LANDG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOOLG TOOLG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_OVRHD OVRHD" & vbCrLf _
                & " from ICTCOSTC"
            Create_TDA(.Tables.Add, "ICTCOSTS", "**", 0, False, "", 1)
            With .Tables("ICTCOSTS")
                .Columns.Add("TOTAL", GetType(System.Decimal), "ISNULL(VCOST,0)+ISNULL(LANDG,0)+ISNULL(TOOLG,0)+ISNULL(OVRHD,0)")
            End With

            Dim sqlICTSTAT2 As String = "Select ITEM_CODE" & vbCrLf _
                & ", Sum (WHSE_QTY_ON_HAND) QTY_ON_HAND" & vbCrLf _
                & ", Sum (WHSE_QTY_ONPO) QTY_ONPO" & vbCrLf _
                & ", Sum (WHSE_QTY_PLAN) QTY_PLAN" & vbCrLf _
                & ", Sum (WHSE_QTY_OPEN) QTY_OPEN" & vbCrLf _
                & ", Sum (WHSE_QTY_PICK) QTY_PICK" & vbCrLf _
                & ", Sum (WHSE_QTY_COMM) QTY_COMM" & vbCrLf _
                & " from ICTSTAT2 where ITEM_CODE = :PARM1 group by ITEM_CODE"

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM, ICTITEM1.ITEM_COST_STD" & vbCrLf _
                & ", ICTITEM1.ITEM_PLAN_MAKE_BUY, ICTITEM1.ITEM_PLAN_WASTE_PCT, ICTITEM1.VEND_ITEM_CODE" & vbCrLf _
                & ", X.QTY_ON_HAND, X.QTY_ONPO, X.QTY_OPEN, X.QTY_COMM, X.QTY_PICK, X.QTY_PLAN" & vbCrLf _
                & " from ICTITEM1, (" & sqlICTSTAT2 & ") X" & vbCrLf _
                & " where X.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "V", 1)
            .Tables("ICTITEM1").Columns.Add("QTY_OPEN_PICK", GetType(Int64), "ISNULL(QTY_OPEN,0)+ISNULL(QTY_COMM,0)+ISNULL(QTY_PICK,0)")
            .Tables("ICTITEM1").Columns.Add("QTY_AVA", GetType(Int64), "ISNULL(QTY_ON_HAND,0)+ISNULL(QTY_OPEN,0)+ISNULL(QTY_PLAN,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_OPEN,0)-ISNULL(QTY_PICK,0)")
        End With

        grdBMTMAINX.DataSource = dst.Tables("BMTMAINX")
        grdBMTMAIN3.DataSource = dst.Tables("BMTMAIN3")
        grdICTCOSTS.DataSource = dst.Tables("ICTCOSTS")

        'Create_Summary(grdBMTMAINX, "BM_ISSUE_NO", "Count")

        Create_Summary(grdBMTMAIN3, "BM_COMP_ITEM", "Count")
        Create_Summary(grdBMTMAIN3, "EXT_COST")
        Create_Summary(grdBMTMAIN3, New String() {"VCOST", "LANDG", "TOOLG", "OVRHD", "TOTAL"})

        grdICTCOSTS.DisplayLayout.Bands(0).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False
        grdICTCOSTS.DisplayLayout.Bands(0).Override.SummaryValueAppearance.BackColor = Drawing.Color.LightGray

        Create_Summary(grdICTCOSTS, New String() {"VCOST", "LANDG", "TOOLG", "OVRHD", "TOTAL"})
        Create_Summary(grdICTCOSTS, "COST_TYPE", "CustomString")


        grdBMTMAIN3.DisplayLayout.UseFixedHeaders = True
        With grdBMTMAIN3.DisplayLayout.Bands(0)
            .Columns("BM_COMP_ITEM").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
            .Columns("BM_SEQ").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"ITEM_DESC", "ITEM_UOM", "ITEM_PLAN_WASTE_PCT"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.Beige
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"VEND_CODE", "VEND_ITEM_CODE"}.Contains(gcol.Key) Then
                    '.Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightPink
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"ITEM_COST_STD", "ITEM_COST_WASTE_PCT", "EXT_COST"}.Contains(gcol.Key) Then
                    '.Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightPink
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"QTY_ON_HAND", "QTY_ONPO", "QTY_PLAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_OPEN_PICK", "QTY_AVA"}.Contains(gcol.Key) Then
                    '.Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.LightBlue
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                ElseIf New String() {"VCOST", "LANDG", "TOOLG", "OVRHD", "TOTAL"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        With grdICTCOSTS.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"VCOST", "LANDG", "TOOLG", "OVRHD", "TOTAL"}.Contains(gcol.Key) Then
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With


        grdBMTMAINX.DisplayLayout.UseFixedHeaders = True
        With grdBMTMAINX.DisplayLayout.Bands("BMTMAINX")
            .Columns("BM_PROD_ITEM").Header.Fixed = True
            .Columns("ITEM_DESC").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdBMTMAIN3, "BM_WHEN_EXHAUSTED", , New String() {":", "A", "B", "C", "D"})
        ASCMAIN1.Add_Value_List(grdBMTMAIN3, "BM_REPLACE_WITH", , New String() {":", "A", "B", "C", "D"})
        ASCMAIN1.Add_Value_List(grdICTCOSTS, "COST_TYPE", , New String() {":", "1:Direct", "2:Materials"})
        ASCMAIN1.Add_Value_List(grdBMTMAINX, "BM_ISSUE_TYPE", , New String() {":", "A:Assy/Mfg", "R:Re-Work"}, 1)

        Bind_Controls(splBMTMAIN2.Panel1, "BMTMAIN2")
        Bind_Controls(grpBM_ISSUE_TEXT, "BMTMAIN2")

        Set_Read_Only_for_ctl(Absx1.chkFor("BM_ISSUE_USE_FOR_STD"), True)

          Show_grdBMTMAIN3_Columns()
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"
                Validate_Code("BM_PROD_ITEM")

                If EMsg = "" Then
                    BM_PROD_ITEM = Absx1.txtFor("BM_PROD_ITEM").Text
                    BM_ISSUE_NO = Absx1.txtFor("BM_ISSUE_NO").Text
                    If BM_ISSUE_NO = "" Or eItemKey = "Edit" Then BM_ISSUE_NO = "00"
                    If eItemKey = "View" Then
                        Dim rowBMTMAIN2 As DataRow = LookUp("BMTMAIN2", New String() {BM_PROD_ITEM, BM_ISSUE_NO})
                        If rowBMTMAIN2 Is Nothing Then
                            EMsg &= "No record on file for " & BM_PROD_ITEM & ", Issue #" & BM_ISSUE_NO
                        End If
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", BM_PROD_ITEM)
                    If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
                        EMsg &= vbCr & "Item Status is not Active"
                    End If
                End If

                If EMsg = "" And eItemKey = "Edit" Then
                    If Not ASCMAIN1.Logical_Lock("BMTMAIN1", Absx1.txtFor("BM_PROD_ITEM").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Cancel"
                If MsgBox("OK to Lose Changes?", MsgBoxStyle.YesNo, _
                          "You may have made Changes") = MsgBoxResult.No Then
                    Exit Sub
                End If

            Case "Update"
                ' Stop ' CHECK STUFF
                If Absx1.txtFor("BM_ISSUE_COMMENT").Text = "" Then
                    EMsg &= vbCr & "Issue Comment is Required"
                End If
                ' MATCHING XR
                ' IF USE FOR STD THEN ALL COMPS MUST HAVE A STD COST
                ' AT LEAST 1 COMP NOT A VSM

                If Absx1.chkFor("BM_ISSUE_USE_FOR_STD").Checked Then
                    If optSaveChangesAs.Value = "W" Then
                        EMsg &= vbCr & "Cannot use Working Issue for Std Costing"
                    End If
                End If

                If optSaveChangesAs.Value = "P" Then
                    Dim BM_WHEN_EXHAUSTEDs As String = ""
                    Dim BM_REPLACE_WITHs As String = ""

                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                        (dst.Tables("BMTMAIN3").Select("ISNULL(BM_WHEN_EXHAUSTED,'') <> ''"), New String() {"BM_WHEN_EXHAUSTED"}).Rows
                        BM_WHEN_EXHAUSTEDs &= row.Item("BM_WHEN_EXHAUSTED")
                    Next
                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                        (dst.Tables("BMTMAIN3").Select("ISNULL(BM_REPLACE_WITH,'') <> ''"), New String() {"BM_REPLACE_WITH"}).Rows
                        BM_REPLACE_WITHs &= row.Item("BM_REPLACE_WITH")
                    Next
                    If BM_WHEN_EXHAUSTEDs <> BM_REPLACE_WITHs Then
                        EMsg &= vbCr & "When-Exhausted and Replace-With Codes Mismatch"
                    End If
                End If

                For Each rowBMTMAIN3 As DataRow In dst.Tables("BMTMAIN3").Select("ISNULL(BM_VEND_SUPP_MATL,'0') = '0'")
                    Dim BM_COMP_ITEM As String = rowBMTMAIN3.Item("BM_COMP_ITEM")
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", BM_COMP_ITEM)
                    If rowICTITEM1.Item("ITEM_COST_STATUS") & "" = "P" Then
                        If optSaveChangesAs.Value = "P" Then
                            EMsg &= vbCr & "Compoment Item " & BM_COMP_ITEM & " does Not have a Standard Cost"
                        Else
                            MsgBox("Item " & BM_COMP_ITEM & "does not have a Standard Cost", MsgBoxStyle.OkOnly, "Warning")
                        End If
                    End If
                    If rowICTITEM1.Item("ITEM_STATUS") & "" <> "A" Then
                        EMsg &= vbCr & "Item " & BM_COMP_ITEM & " is Not Active"
                    End If
                Next

                If dst.Tables("BMTMAIN3").Select("ISNULL(BM_VEND_SUPP_MATL,'0') = '0'").Length = 0 Then
                    EMsg &= vbCr & "No BOM Details; Update Denied"
                End If

                If optType.Value = "R" Then
                    If Absx1.chkFor("BM_ISSUE_USE_FOR_STD").Checked Then
                        EMsg &= vbCr & "You May NOT Use a Re-Work Type BOM for Std Costing"
                    End If
                End If

                If dst.Tables("BMTMAIN3").Select($"BM_COMP_ITEM = '{BM_PROD_ITEM}'").Length > 0 Then
                    If optType.Value <> "R" Then
                        EMsg &= vbCr & "The BM may list the Product as a Component only if the BM Type is Re-Work"
                    End If
                Else
                    If optType.Value = "R" Then
                        EMsg &= vbCr & "The BM must list the Product as a Component if the BM Type is Re-Work"
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

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Record()
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

                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("Edit").Visible = (Not InquiryMode)
                    .Items("Done").Visible = (EntryMode = "N" Or EntryMode <> "E")
                    .Items("Print").Visible = (InquiryMode Or EntryMode = "V")
                    .Items("Update").Visible = (Not InquiryMode And EntryMode <> "V")
                    .Items("Cancel").Visible = (Not InquiryMode And EntryMode <> "V")
                End With
                .Groups("Save Changes As").Visible = (EntryMode = "E")
                .Groups("Print Options").Visible = .Groups("Screen Control").Items("Print").Visible And ScreenMode
                .Groups("Item Image").Visible = ScreenMode
            End With
        End If

        If InquiryMode Or (EntryMode = "V") Then
            grdBMTMAIN3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdBMTMAIN3.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            grdBMTMAIN3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            grdBMTMAIN3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdBMTMAIN3.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            grdBMTMAIN3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(splBMTMAIN2.Panel1, InquiryMode Or EntryMode = "V")
        Set_Read_Only(grpBM_ISSUE_TEXT, InquiryMode Or EntryMode = "V")
        grdBMTMAINX.Visible = Not ScreenMode
        splBMTMAIN2.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"BMTMAIN1", "BMTMAIN2", "BMTMAIN3", "ICTCOSTC", "ICTITEM1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Fill_Records("BMTMAINX")
        Fill_Records("BMTMAINY")
        Sort_grdColumns(grdBMTMAINX, "BM_PROD_ITEM")
        Sort_grdColumns(grdBMTMAINX, "BM_ISSUE_NO", , 1)

        EnforceConstraints(True)

        Absx1.txtFor("BM_PROD_ITEM").Text = ""
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        rowBMTMAIN1 = Fill_Record("BMTMAIN1", BM_PROD_ITEM)
        If rowBMTMAIN1 Is Nothing Then
            rowBMTMAIN1 = dst.Tables("BMTMAIN1").NewRow
            rowBMTMAIN1.Item("BM_PROD_ITEM") = BM_PROD_ITEM
            rowBMTMAIN1.Item("BM_ISSUE_COUNTER") = 0
            'rowBMTMAIN1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            'rowBMTMAIN1.Item("INIT_DATE") = DATETIME_STAMP
            dst.Tables("BMTMAIN1").Rows.Add(rowBMTMAIN1)
        End If

        rowBMTMAIN2 = Fill_Record("BMTMAIN2", New String() {BM_PROD_ITEM, BM_ISSUE_NO})
        If rowBMTMAIN2 Is Nothing Then
            If BM_ISSUE_NO <> "00" Then Stop ' should never need to create an issue other than the working issue
            rowBMTMAIN2 = dst.Tables("BMTMAIN2").NewRow
            rowBMTMAIN2.Item("BM_PROD_ITEM") = BM_PROD_ITEM
            rowBMTMAIN2.Item("BM_ISSUE_NO") = BM_ISSUE_NO
            rowBMTMAIN2.Item("BM_ISSUE_DATE") = DATETIME_STAMP.Date
            rowBMTMAIN2.Item("BM_ISSUE_COMMENT") = "Initial Issue"
            rowBMTMAIN2.Item("BM_ISSUE_USE_FOR_STD") = "0"
            rowBMTMAIN2.Item("BM_ISSUE_TYPE") = "A"
            rowBMTMAIN2.Item("BM_ISSUE_VCOST") = 0
            dst.Tables("BMTMAIN2").Rows.Add(rowBMTMAIN2)
        End If

        Fill_Records("BMTMAIN3", New String() {BM_PROD_ITEM, BM_ISSUE_NO})
        Sort_grdColumns(grdBMTMAIN3, "BM_SEQ")
        EnforceConstraints(True)

        optSaveChangesAs.Value = "W"
        optSaveChangesAs.ValueList.ValueListItems(1).DisplayText = "Permanent Issue #" & CStr(Val(rowBMTMAIN1.Item("BM_ISSUE_COUNTER") & "") + 1)

        dst.Tables("ICTCOSTS").Rows.Clear()
        rowICTCOSTC_product = LookUp("ICTCOSTC", BM_PROD_ITEM, True)
        dst.Tables("ICTCOSTS").Rows.Add(New Object() {"1", rowICTCOSTC_product.Item("ITEM_COST_VCOST") _
                                                         , rowICTCOSTC_product.Item("ITEM_COST_LANDG") _
                                                         , rowICTCOSTC_product.Item("ITEM_COST_TOOLG") _
                                                         , rowICTCOSTC_product.Item("ITEM_COST_OVRHD")})
        '_
        ', Val(rowICTCOSTC_product.Item("ITEM_COST_VCOST") & "") +
        '  Val(rowICTCOSTC_product.Item("ITEM_COST_LANDG") & "") +
        '  Val(rowICTCOSTC_product.Item("ITEM_COST_TOOLG") & "") +
        '  Val(rowICTCOSTC_product.Item("ITEM_COST_OVRHD") & "")})
        dst.Tables("ICTCOSTS").Rows.Add(New Object() {"2"})

        For Each rowBMTMAIN3 As DataRow In dst.Tables("BMTMAIN3").Select("")
            Dim BM_COMP_ITEM As String = rowBMTMAIN3.Item("BM_COMP_ITEM")
            Fill_Records("ICTCOSTC", BM_COMP_ITEM, False)
        Next
        Fill_Records("ICTCOSTC", BM_PROD_ITEM, False)
 
        Dim IMAGE_NAME As String = BM_PROD_ITEM
        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        Dim imgba() As Byte = Nothing
        picItemImage.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
        UltraExplorerBar1.Groups("Item Image").Text = "Item Image " & BM_PROD_ITEM
        'rowBMTMAIN1.Item("IMAGE") = imgba

        Display_Totals()
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()
        Synch_TABLE_NAME("BMTMAIN2")

        Re_Sequence()

        Dim sql_delete As String = "BM_PROD_ITEM = '" & BM_PROD_ITEM & "' and BM_ISSUE_NO = '" & BM_ISSUE_NO & "'"

        If rowBMTMAIN2.Item("BM_ISSUE_TYPE") & "" = "R" Then
        Else
            rowBMTMAIN2.Item("BM_ISSUE_VCOST") = DBNull.Value
        End If

        INIT_LAST("BMTMAIN1")
        Update_Record_TDA("BMTMAIN1")
        Update_Record_TDA("BMTMAIN2")
        Update_Record_TDA("BMTMAIN3", sql_delete)

        If optSaveChangesAs.Value = "P" Then
            Dim BM_ISSUE_COUNTER As Int64 = Val(rowBMTMAIN1.Item("BM_ISSUE_COUNTER") & "") + 1
            rowBMTMAIN1.Item("BM_ISSUE_COUNTER") = BM_ISSUE_COUNTER

            Dim BM_ISSUE_NO_P As String = Format(BM_ISSUE_COUNTER, "00")
            rowBMTMAIN2.Item("BM_ISSUE_NO") = BM_ISSUE_NO_P
            rowBMTMAIN2.AcceptChanges()
            rowBMTMAIN2.SetAdded()

            For Each rowBMTMAIN3 As DataRow In dst.Tables("BMTMAIN3").Select("")
                rowBMTMAIN3.Item("BM_ISSUE_NO") = BM_ISSUE_NO_P
                rowBMTMAIN3.AcceptChanges()
                rowBMTMAIN3.SetAdded()
            Next

            Update_Record_TDA("BMTMAIN1")
            Update_Record_TDA("BMTMAIN2")
            Update_Record_TDA("BMTMAIN3")
        End If

        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        Stop ' WHEN WOULD WE BE PERMITTING THIS?
        BeginTrans()
        Delete_Records("BMTMAIN1")
        Delete_Records("BMTMAIN2")
        Delete_Records("BMTMAIN3")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where BM_PROD_ITEM = '" & BM_PROD_ITEM & "'" _
            & IIf(TABLE_NAME = "BMTMAIN1", "", "   and BM_ISSUE_NO = '" & BM_ISSUE_NO & "'"))
    End Sub

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "BMTMAIN1"
            E.COLUMN_NAME = "BM_PROD_ITEM"
            E.CODE_VALUE = Absx1.txtFor("BM_PROD_ITEM").Text
            E.DESC_VALUE = "Bill of Materials"
            E.ATTACHMENT_NOTES = ""
            'If rowPOTORDR1.Item("PO_STATUS_CODE") & "" <> "O" Or rowPOTORDR1.Item("PO_OK_TO_PAY") & "" = "1" Then
            '    E.RESTRICTIONS = "D"
            'End If
            E.READ_ONLY = (EntryMode <> "E")
        End If

        Return E
    End Function

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command(command)

            Case "View"
                Absx1.txtFor("BM_PROD_ITEM").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdBMTMAINX, "S", "Show Filter")
        Load_Popup_Menu(grdBMTMAIN3, "BBSSSB", "Copy From ...", "Re-Sequence", _
                        "Show Costing Data", "Show Qty Data", "Show Supplier Data", _
                        "Inventory Status")
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
                Case "grdBMTMAIN3"
                    tlb_btn = DirectCast(tlb_pop.Tools("Copy From ..."), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E")

                    tlb_btn = DirectCast(tlb_pop.Tools("Re-Sequence"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (EntryMode = "E")

                    tlb_sbt = DirectCast(tlb_pop.Tools("Show Qty Data"), UltraWinToolbars.StateButtonTool)
                    tlb_sbt.SharedProps.Visible = (EntryMode = "V")
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Costing Data", "Show Qty Data", "Show Supplier Data"
                Load_BM_with_Status()
                Show_grdBMTMAIN3_Columns()

            Case "Copy From ..."
                Copy_From_Issue()

            Case "Re-Sequence"
                Re_Sequence()

            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Inventory Status"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("BM_COMP_ITEM").Value
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BM_PROD_ITEM"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    'Click_Command("Load", e)
                    Show_Issues(Absx1.txtFor("BM_PROD_ITEM").Text)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BM_PROD_ITEM"
                Show_Issues(Absx1.txtFor("BM_PROD_ITEM").Text)
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME

            Case "BM_PROD_ITEM"
                If EntryMode = "" Then
                    If Absx1.txtFor("BM_PROD_ITEM").Text <> "" Then
                        LookUp("ICTITEM1", Absx1.txtFor("BM_PROD_ITEM").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

    Sub Show_Issues(ByVal ITEM_CODE As String)
        Absx1.txtFor("BM_ISSUE_NO").Text = ""
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE) '  LookUp("ICTITEM1", ITEM_CODE, True)

        If rowICTITEM1 Is Nothing Then
            If ITEM_CODE <> "" Then
                MsgBox("Invalid Item Code " & ITEM_CODE, MsgBoxStyle.OkOnly, "Cannot Show BM Issues")
            End If
            Exit Sub
        End If

        Absx1.txtFor("ITEM_DESC").Text = rowICTITEM1.Item("ITEM_DESC") & ""
        Absx1.txtFor("ITEM_UOM").Text = rowICTITEM1.Item("ITEM_UOM") & ""

        Dim dvw As DataView = dst.Tables("BMTMAINX").DefaultView
        dvw.RowFilter = "BM_PROD_ITEM = '" & ITEM_CODE & "'"

        'Fill_Records("BMTMAIN2", ITEM_CODE)
        If ITEM_CODE = "" Then
            grdBMTMAINX.Text = "All Bill of Material Items & Issues"
        Else
            grdBMTMAINX.Text = "BM Issues on file for Item " & ITEM_CODE
        End If
        'Sort_grdColumns(grdBMTMAINX, "BM_ISSUE_NO")
    End Sub

    Sub Show_grdBMTMAIN3_Columns()
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool
        With grdBMTMAIN3.DisplayLayout.Bands(0)
            tlb_sbt = DirectCast(tlb.Tools("Show Costing Data"), UltraWinToolbars.StateButtonTool)
            .Columns("ITEM_COST_STD").Hidden = Not tlb_sbt.Checked
            .Columns("ITEM_COST_WASTE_PCT").Hidden = Not tlb_sbt.Checked
            .Columns("EXT_COST").Hidden = Not tlb_sbt.Checked
            For Each COLUMN_NAME As String In New String() {"TOTAL", "VCOST", "LANDG", "TOOLG", "OVRHD"}
                .Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
            Next
            splIssue.Panel1Collapsed = tlb_sbt.Checked
            ' grdICTCOSTS.Visible = tlb_sbt.Checked
            ' splBMTMAIN3.Panel2Collapsed = Not tlb_sbt.Checked

            tlb_sbt = DirectCast(tlb.Tools("Show Qty Data"), UltraWinToolbars.StateButtonTool)
            If EntryMode = "E" And tlb_sbt.Checked Then
                tlb_sbt.Checked = False
            End If
            For Each COLUMN_NAME As String In New String() {"QTY_ON_HAND", "QTY_ONPO", "QTY_PLAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM", "QTY_OPEN_PICK", "QTY_AVA"}
                If COLUMN_NAME = "QTY_OPEN" Or COLUMN_NAME = "QTY_PICK" Then
                    .Columns(COLUMN_NAME).Hidden = True
                Else
                    .Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                End If
            Next

            tlb_sbt = DirectCast(tlb.Tools("Show Supplier Data"), UltraWinToolbars.StateButtonTool)
            .Columns("VEND_CODE").Hidden = Not tlb_sbt.Checked
            .Columns("VEND_ITEM_CODE").Hidden = Not tlb_sbt.Checked
        End With
    End Sub

#Region "grdBMTMAIN3"

    Private Sub grdBMTMAIN3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdBMTMAIN3.AfterCellUpdate
        If e.Cell.Column.Key = "BM_COMP_ITEM" Then
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", New String() {e.Cell.Value})
            If rowICTITEM1 IsNot Nothing Then
                e.Cell.Row.Cells("ITEM_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                e.Cell.Row.Cells("ITEM_UOM").Value = rowICTITEM1.Item("ITEM_UOM")
                e.Cell.Row.Cells("ITEM_COST_STD").Value = rowICTITEM1.Item("ITEM_COST_STD")
                e.Cell.Row.Cells("ITEM_COST_WASTE_PCT").Value = rowICTITEM1.Item("ITEM_COST_WASTE_PCT")
                e.Cell.Row.Cells("ITEM_PLAN_WASTE_PCT").Value = rowICTITEM1.Item("ITEM_PLAN_WASTE_PCT")
                e.Cell.Row.Cells("VEND_CODE").Value = rowICTITEM1.Item("VEND_CODE")
                e.Cell.Row.Cells("VEND_ITEM_CODE").Value = rowICTITEM1.Item("VEND_ITEM_CODE")

                Dim rowICTCOSTC As DataRow = LookUp("ICTCOSTC", New String() {e.Cell.Value}, True)
                e.Cell.Row.Cells("ITEM_COST_TOTAL").Value = rowICTCOSTC.Item("ITEM_COST_TOTAL")
                e.Cell.Row.Cells("ITEM_COST_VCOST").Value = rowICTCOSTC.Item("ITEM_COST_VCOST")
                e.Cell.Row.Cells("ITEM_COST_LANDG").Value = rowICTCOSTC.Item("ITEM_COST_LANDG")
                e.Cell.Row.Cells("ITEM_COST_TOOLG").Value = rowICTCOSTC.Item("ITEM_COST_TOOLG")
                e.Cell.Row.Cells("ITEM_COST_OVRHD").Value = rowICTCOSTC.Item("ITEM_COST_OVRHD")
            End If
        End If
    End Sub

    Private Sub grdBMTMAIN3_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdBMTMAIN3.AfterExitEditMode
        'With grdBMTMAIN3
        '    Select Case .ActiveCell.Column.Key
        '        Case "BM_COMP_ITEM"
        '            If .ActiveCell.Text <> "" Then
        '                '.ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)
        '                cdr = LookUp("ICTITEM1", .ActiveCell.Text)



        '            End If
        '    End Select
        'End With
    End Sub

    Private Sub grdBMTMAIN3_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdBMTMAIN3.AfterRowActivate

        With grdBMTMAIN3.DisplayLayout.Bands(0)
            If grdBMTMAIN3.ActiveRow.IsAddRow Then
                .Columns("BM_COMP_ITEM").CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                .Columns("BM_COMP_ITEM").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdBMTMAIN3_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdBMTMAIN3.AfterRowsDeleted
        Display_Totals()
    End Sub

    Private Sub grdBMTMAIN3_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdBMTMAIN3.AfterRowUpdate
        Display_Totals()

        Dim BM_COMP_ITEM As String = e.Row.Cells("BM_COMP_ITEM").Value
        If dst.Tables("ICTCOSTC").Rows.Find(BM_COMP_ITEM) Is Nothing Then
            Fill_Records("ICTCOSTC", BM_COMP_ITEM, False)
        End If
    End Sub

    Private Sub grdBMTMAIN3_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdBMTMAIN3.BeforeCellUpdate
        e.Cancel = Validate_Columns_BMTMAIN3(e.Cell.Column.Key, e.NewValue)
    End Sub

    Private Sub grdBMTMAIN3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdBMTMAIN3.BeforeExitEditMode

        If e.CancellingEditOperation Then Exit Sub
        With grdBMTMAIN3.ActiveCell
            Select Case .Column.Key

                Case "BM_COMP_ITEM"

                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With

    End Sub

    Private Sub grdBMTMAIN3_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdBMTMAIN3.BeforeRowsDeleted

    End Sub

    Private Sub grdBMTMAIN3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdBMTMAIN3.BeforeRowUpdate

        e.Cancel = Validate_Columns_BMTMAIN3("BM_COMP_ITEM", e.Row.Cells("BM_COMP_ITEM").Value)
        If e.Cancel = False Then
            e.Cancel = Validate_Columns_BMTMAIN3("BM_QTY_PER_ASSY", e.Row.Cells("BM_COMP_ITEM").Value)
        End If

        If e.Row.Cells("BM_WHEN_EXHAUSTED").Value & "" <> "" And _
             e.Row.Cells("BM_REPLACE_WITH").Value & "" <> "" Then
            'MsgBox("An Item Can't be assigned both an Exhaust and Replace Code", vbOKOnly, "Pleas verify")
            e.Cancel = False
        End If

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("BM_PROD_ITEM").Value = Absx1.txtFor("BM_PROD_ITEM").Text
            e.Row.Cells("BM_ISSUE_NO").Value = Absx1.txtFor("BM_ISSUE_NO").Text
            e.Row.Cells("BM_SEQ").Value = Val(dst.Tables("BMTMAIN3").Compute("MAX(BM_SEQ)", "") & "") + 10
        End If
    End Sub

    Private Sub grdBMTMAIN3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdBMTMAIN3.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdBMTMAIN3.ActiveCell.Column.Key
            Case "BM_COMP_ITEM"
        End Select

        grdClickCellButton(grdBMTMAIN3, sql_where, False)
    End Sub

    Function Validate_Columns_BMTMAIN3(COLUMN_NAME As String, CELLVALUE As Object) As Boolean

        Select Case COLUMN_NAME
            Case "BM_COMP_ITEM"
                Dim BM_COMP_ITEM As String = CELLVALUE & ""
                If BM_COMP_ITEM = Absx1.txtFor("BM_PROD_ITEM").Text Then
                    If Absx1.optFor("BM_ISSUE_TYPE").Value <> "R" Then
                        MsgBox("Product Item may not be used as one of its Component Items unless the BM Type is Re-Work", MsgBoxStyle.OkOnly, "Cannot Update Record")
                        Return True
                    End If
                End If
                If grdBMTMAIN3.ActiveRow.IsAddRow Then
                    If dst.Tables("BMTMAIN3").Rows.Find(New String() {BM_PROD_ITEM, BM_ISSUE_NO, BM_COMP_ITEM}) IsNot Nothing Then
                        MsgBox("Component Item " & BM_COMP_ITEM & " is already on Bill of Materials", MsgBoxStyle.OkOnly, "Cannot Update Record")
                        Return True
                    End If
                End If

                'Case "BM_QTY_PER_ASSY"
                '    If Val(CELLVALUE & "") <= 0 Then
                '        MsgBox("Qty per Assembly May Not be Zero", MsgBoxStyle.OkOnly, "Cannot Update Record")
                '        Return True
                '    End If

            Case "BM_SEQ"
                If Val(CELLVALUE & "") < 0 Or Val(CELLVALUE & "") > 999 Then
                    MsgBox("Sequence must be between 0 and 999", MsgBoxStyle.OkOnly, "Cannot Update Record")
                    Return True
                End If

            Case Else
                Return False
        End Select
    End Function

#End Region

    Private Sub grdBMTMAINX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdBMTMAINX.DoubleClickRow
        If e.Row.Band.Key = "BMTMAINX" Then
            Absx1.txtFor("BM_PROD_ITEM").Text = e.Row.Cells("BM_PROD_ITEM").Value
            Absx1.txtFor("BM_ISSUE_NO").Text = "00"
            If InquiryMode Then
                Click_Command("View")
            Else
                Click_Command("View")
            End If
        Else
            Absx1.txtFor("BM_PROD_ITEM").Text = e.Row.Cells("BM_PROD_ITEM").Value
            Absx1.txtFor("BM_ISSUE_NO").Text = e.Row.Cells("BM_ISSUE_NO").Value
            Click_Command("View")
        End If
    End Sub

    Sub Re_Sequence()
        Dim BM_SEQ As Integer = 0

        For Each rowBMTMAIN3 As DataRow In dst.Tables("BMTMAIN3").Select("", "BM_SEQ")
            BM_SEQ = Val(rowBMTMAIN3.Item("BM_SEQ") & "")
            rowBMTMAIN3.Item("BM_COMPONENT_SORT") = Format(BM_SEQ, "000")
        Next

        Dim BM_WHEN_EXHAUSTEDs As String = ""
        For Each row As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("BMTMAIN3").Select("ISNULL(BM_WHEN_EXHAUSTED,'') <> ''"), New String() {"BM_WHEN_EXHAUSTED"}).Rows
            BM_WHEN_EXHAUSTEDs &= row.Item("BM_WHEN_EXHAUSTED")
        Next

        If BM_WHEN_EXHAUSTEDs <> "" Then
            For I As Integer = 1 To Len(BM_WHEN_EXHAUSTEDs)
                Dim BM_WHEN_EXHAUSTED As String = Mid(BM_WHEN_EXHAUSTEDs, I, 1)
                Dim BM_SEQ_first As Integer = -1

                For Each row As DataRow In dst.Tables("BMTMAIN3").Select _
                    ("BM_WHEN_EXHAUSTED = '" & BM_WHEN_EXHAUSTED & "'", "BM_SEQ")
                    If BM_SEQ_first = -1 Then
                        BM_SEQ_first = Val(row.Item("BM_SEQ") & "")
                    End If
                    BM_SEQ = Val(row.Item("BM_SEQ") & "")
                    row.Item("BM_COMPONENT_SORT") = Format(BM_SEQ_first, "000") & BM_WHEN_EXHAUSTED & "1" & Format(BM_SEQ, "000")
                Next
                For Each row As DataRow In dst.Tables("BMTMAIN3").Select _
                    ("BM_REPLACE_WITH = '" & BM_WHEN_EXHAUSTED & "'", "BM_SEQ")
                    BM_SEQ = Val(row.Item("BM_SEQ") & "")
                    row.Item("BM_COMPONENT_SORT") = Format(BM_SEQ_first, "000") & BM_WHEN_EXHAUSTED & "2" & Format(BM_SEQ, "000")
                Next
            Next
        End If

        BM_SEQ = 0
        For Each row As DataRow In dst.Tables("BMTMAIN3").Select("", "BM_COMPONENT_SORT")
            BM_WHEN_EXHAUSTEDs &= row.Item("BM_WHEN_EXHAUSTED")
            BM_SEQ += 10
            row.Item("BM_SEQ") = BM_SEQ
        Next

        Sort_grdColumns(grdBMTMAIN3, "BM_SEQ")
    End Sub

    Private Sub Copy_From_Issue()

        ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("BM_ISSUE_NO")
        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = False
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()
            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Copying")
                Dim BM_PROD_ITEM_copy_from As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("BM_PROD_ITEM")
                Dim BM_ISSUE_NO_copy_from As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("BM_ISSUE_NO")
                Copy_BM_ISSUE_NO(BM_PROD_ITEM_copy_from, BM_ISSUE_NO_copy_from)
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")
            End If
        End If

    End Sub

    Sub Copy_BM_ISSUE_NO(BM_PROD_ITEM_to_copy As String, BM_ISSUE_NO_to_copy As String)

        Dim DT As New DataTable
        With DT
            .Columns.Add("ITEM_CODE")
            .Columns.Add("MESSAGE")
        End With
        Dim all_items_were_copied As Boolean = True

        ASCMAIN1.sql = "Select * from BMTMAIN3 " _
            & " where BM_PROD_ITEM = '" & BM_PROD_ITEM_to_copy & "'" _
            & "   and BM_ISSUE_NO = '" & BM_ISSUE_NO_to_copy & "'"

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "BM_SEQ")
            If grdBMTMAIN3.ActiveRow IsNot Nothing And grdBMTMAIN3.ActiveRow.DataChanged Then
                grdBMTMAIN3.ActiveRow.CancelUpdate()
            End If
            Dim BM_COMP_ITEM As String = row.Item("BM_COMP_ITEM")
            If BM_COMP_ITEM <> BM_PROD_ITEM And _
                dst.Tables("BMTMAIN3").Rows.Find(New String() {BM_PROD_ITEM, BM_ISSUE_NO, BM_COMP_ITEM}) Is Nothing Then
                With grdBMTMAIN3.DisplayLayout.Bands(0).AddNew()
                    .Cells("BM_COMP_ITEM").Value = BM_COMP_ITEM
                    .Cells("BM_QTY_PER_ASSY").Value = row.Item("BM_QTY_PER_ASSY")
                    .Cells("BM_VEND_SUPP_MATL").Value = row.Item("BM_VEND_SUPP_MATL")
                    .Cells("BM_WHEN_EXHAUSTED").Value = row.Item("BM_WHEN_EXHAUSTED")
                    .Cells("BM_REPLACE_WITH").Value = row.Item("BM_REPLACE_WITH")
                    .Update()
                    'DT.Rows.Add(New String() {BM_COMP_ITEM, ""})
                    'If .DataChanged Then
                    '    DT.Rows.Add(New String() {BM_COMP_ITEM, "Not Copied - Data Error"})
                    'Else
                    '    DT.Rows.Add(New String() {BM_COMP_ITEM, ""})
                    'End If
                End With
            Else
                DT.Rows.Add(New String() {BM_COMP_ITEM, "Not Copied - Existing Item"})
                all_items_were_copied = False
            End If
        Next
        If grdBMTMAIN3.ActiveRow IsNot Nothing And grdBMTMAIN3.ActiveRow.DataChanged Then
            grdBMTMAIN3.ActiveRow.CancelUpdate()
        End If

        Sort_grdColumns(grdBMTMAIN3, "BM_SEQ")
        Display_Totals()

        If all_items_were_copied Then
            MsgBox("All Component Items were Sucessfully Copied", MsgBoxStyle.OkOnly, "Copy Complete")
        Else
            Using F As New ASFMSGBF
                F.Show_grd(DT, Me, IIf(all_items_were_copied, "Copy Complete", "Copy Complete - Not All Items were Copied"))
            End Using
        End If
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Bill of Materials")

        Re_Sequence()
        Get_Item_and_Qtys()
        Print_Report_Begin()

        CR_params.Add("RUNQTY", IIf(Not optPrintOptions.Value = "Q", "0", CStr(numRunQty.Value & "")))
        CR_params.Add("STATUS", IIf(Not optPrintOptions.Value = "Q", "0", "1"))
        CR_params.Add("COSTED_BOM", IIf(optPrintOptions.Value = "C", "1", "0"))

        CR_params.Add("NOTES", "1")
        CR_params.Add("COMPNOTES", "1")

        Generate_Report("BMRLIST1", "Bill of Materials", "")
        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Display_Totals()
        If Me.SELECTION_NO = 0 Or EntryMode = "" Then Exit Sub
        If dst.Tables("BMTMAIN3").Rows.Count = 0 Then Return

        Dim rowICTCOSTS As DataRow = dst.Tables("ICTCOSTS").Rows.Find("2")
        With dst.Tables("BMTMAIN3")
            rowICTCOSTS.Item("VCOST") = Val(.Compute("SUM(VCOST)", "") & "")
            rowICTCOSTS.Item("LANDG") = Val(.Compute("SUM(LANDG)", "") & "")
            rowICTCOSTS.Item("TOOLG") = Val(.Compute("SUM(TOOLG)", "") & "")
            rowICTCOSTS.Item("OVRHD") = Val(.Compute("SUM(OVRHD)", "") & "")
            'rowICTCOSTS.Item("TOTAL") = Val(.Compute("SUM(TOTAL)", "") & "")
        End With

        rowICTCOSTS = dst.Tables("ICTCOSTS").Rows.Find("1")
        If optType.Value = "R" Then
            rowICTCOSTS.Item("VCOST") = Val(numBM_ISSUE_VCOST.Value & "")
            rowICTCOSTS.Item("LANDG") = 0
            rowICTCOSTS.Item("TOOLG") = 0
            rowICTCOSTS.Item("OVRHD") = 0
        Else
            rowICTCOSTS.Item("VCOST") = Val(rowICTCOSTC_product.Item("ITEM_COST_VCOST") & "")
            rowICTCOSTS.Item("LANDG") = Val(rowICTCOSTC_product.Item("ITEM_COST_LANDG") & "")
            rowICTCOSTS.Item("TOOLG") = Val(rowICTCOSTC_product.Item("ITEM_COST_TOOLG") & "")
            rowICTCOSTS.Item("OVRHD") = Val(rowICTCOSTC_product.Item("ITEM_COST_OVRHD") & "")
        End If
    End Sub

    Private Sub optSaveChangesAs_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSaveChangesAs.ValueChanged
        If optType.Value = "R" Then
            Absx1.chkFor("BM_ISSUE_USE_FOR_STD").Checked = False
        Else
            Absx1.chkFor("BM_ISSUE_USE_FOR_STD").Checked = (optSaveChangesAs.Value = "P")
        End If

    End Sub

    Public Overrides Function CustomStringSummary_End( _
        ByVal summarySettings As UltraWinGrid.SummarySettings, _
        ByVal rows As UltraWinGrid.RowsCollection, _
        ByVal CustomValue As String, _
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

    Sub Load_BM_with_Status()
        Get_Item_and_Qtys()
        For Each rowBMTMAIN3 As DataRow In dst.Tables("BMTMAIN3").Select("")
            Dim ITEM_CODE As String = rowBMTMAIN3.Item("BM_COMP_ITEM")
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            For Each COLUMN_NAME As String In New String() {"QTY_ON_HAND", "QTY_ONPO", "QTY_PLAN", "QTY_OPEN", "QTY_PICK", "QTY_COMM"}
                rowBMTMAIN3.Item(COLUMN_NAME) = rowICTITEM1.Item(COLUMN_NAME)
            Next
        Next
    End Sub

    Sub Get_Item_and_Qtys()

        dst.Tables("ICTITEM1").Rows.Clear()
        For Each row As DataRow In dst.Tables("BMTMAIN3").Select("")
            Dim BM_COMP_ITEM As String = row.Item("BM_COMP_ITEM")
            Fill_Records("ICTITEM1", BM_COMP_ITEM, False)
        Next
        Fill_Records("ICTITEM1", BM_PROD_ITEM, False)
    End Sub

    Private Sub optPrintOptions_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optPrintOptions.ValueChanged
        lblRunQty.Visible = (optPrintOptions.Value = "Q")
        numRunQty.Visible = (optPrintOptions.Value = "Q")
    End Sub

    Private Sub optType_ValueChanged(sender As Object, e As EventArgs) Handles optType.ValueChanged
        lblBM_ISSUE_VCOST.Visible = (optType.Value = "R")
        numBM_ISSUE_VCOST.Visible = (optType.Value = "R")

        Display_Totals()
    End Sub

    Private Sub numBM_ISSUE_VCOST_ValueChanged(sender As Object, e As EventArgs) Handles numBM_ISSUE_VCOST.ValueChanged
        Display_Totals()
    End Sub
End Class