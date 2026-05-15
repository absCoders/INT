Public Class SPFSFOC1
    Dim rowSPTSFOC1 As DataRow
    Dim EVENT_GROUP_NO As String
    Dim EVENT_GROUP_NO_new As String
    Dim STATUS_CODE As String

    'Dim SPTCODE1 As String = ""
    'Dim SALES_DIVISION_CODE As String
    Dim AUTH_APPR_NOTES As String
    Dim APPR_STATUS_CODE As String

    Dim EXPENSE_TYPE_CODEs_I_may_approve As New List(Of String)
    'Dim APPR_STATUS_CODE_BackColors As Dictionary(Of String, System.Drawing.Color)
    'Dim APPR_STATUS_CODE_ForeColors As Dictionary(Of String, System.Drawing.Color)
    Dim update_with_approval As Boolean = False

    Dim sqlSPTSFOCX As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "SPFSFOCI" Then
            InquiryMode = True
        End If

        With UltraExplorerBar1.Groups("Screen Control")
            .Items("New").Visible = Not InquiryMode
            .Items("Edit").Visible = Not InquiryMode
            .Items("Update").Visible = Not InquiryMode
            .Items("Cancel").Visible = Not InquiryMode
            .Items("Clone").Visible = Not InquiryMode
        End With

        Get_PARM("ICTPARM1")
        Get_PARM("SPTPARM1")

        With dst
            sqlSPTSFOCX = "Select SPTSFOC1.*, SPTSFOC3.COLLECTION_CODE, SPTSFOC3.EVENT_GROUP_LNO" & vbCrLf _
                & ", SPTSFOC3.FEATURE_DESC, SPTSFOC3.ITEM_CODE" & vbCrLf _
                & ", ICTCOLL1.COLLECTION_NAME, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME" & vbCrLf _
                & " from SPTSFOC1,SPTSFOC3,ICTCOLL1,ICTBRAN1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE (+) = SPTSFOC3.COLLECTION_CODE" & vbCrLf _
                & "   and SPTSFOC1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SPTSFOC3.EVENT_GROUP_NO = SPTSFOC1.EVENT_GROUP_NO"
            ASCMAIN1.sql = sqlSPTSFOCX & "  and SPTSFOC1.OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "SPTSFOCX", "**", 0, False, "V")

            ASCMAIN1.sql = sqlSPTSFOCX
            Create_TDA(.Tables.Add, "SPTSFOCG", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "SPTSFOC1", "*")

            ASCMAIN1.sql = "Select SPTSFOC3.*, ICTCOLL1.COLLECTION_NAME, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME" & vbCrLf _
                & " from SPTSFOC3,ICTCOLL1,ICTBRAN1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE (+) = SPTSFOC3.COLLECTION_CODE" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE"
            Create_TDA(.Tables.Add, "SPTSFOC3", "**", 1)

            ASCMAIN1.sql = "Select SPTSFOC9.*" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_LOCATION, ARTCUST2.CUST_STORE_NAME" & vbCrLf _
                & ", ARTCUST2.SELL_CODE, SOTSELL1.REGION_CODE, ARTCUST2.SELL_CODE_AC" & vbCrLf _
                & " from SPTSFOC9,ARTCUST2,SPTSFOC1,SOTSELL1" & vbCrLf _
                & " where SPTSFOC9.EVENT_GROUP_NO = :PARM1" & vbCrLf _
                & "   and SPTSFOC1.EVENT_GROUP_NO = SPTSFOC9.EVENT_GROUP_NO" & vbCrLf _
                & "   and SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_CODE = SPTSFOC9.CUST_CODE" & vbCrLf _
                & "   and ARTCUST2.CUST_STORE_NO = SPTSFOC9.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "SPTSFOC9", "**", 0, True, "V")
            With .Tables("SPTSFOC9")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
            End With

            ASCMAIN1.sql = "Select * from SPTCOOP1 where EVENT_GROUP_NO = :PARM1"
            Create_TDA(.Tables.Add, "SPTCOOP1", "**", 0, True, "V")

            'ASCMAIN1.sql = "Select * from SPTCOOP3 where AUTH_NO in (Select AUTH_NO from SPTCOOP1 where EVENT_GROUP_NO = :PARM1)"
            'Create_TDA(.Tables.Add, "SPTCOOP3", "**", 0, True, "V")
            Create_TDA(.Tables.Add, "SPTCOOP3", "*")

            'ASCMAIN1.sql = "Select * from SPTCOOPB where AUTH_NO in (Select AUTH_NO from SPTCOOP1 where EVENT_GROUP_NO = :PARM1)"
            'Create_TDA(.Tables.Add, "SPTCOOPB", "**", 0, True, "V")
            Create_TDA(.Tables.Add, "SPTCOOPB", "*")
        End With

        grdSPTSFOC3.DataSource = dst.Tables("SPTSFOC3")
        grdSPTSFOC9.DataSource = dst.Tables("SPTSFOC9")
        grdSPTSFOCX.DataSource = dst.Tables("SPTSFOCX")

        Create_Summary(grdSPTSFOCX, "EVENT_GROUP_NO", "Count")

        Create_Summary(grdSPTSFOC9, "CUST_STORE_NO", "Count")
        Create_Summary(grdSPTSFOC9, New String() {"SEL"})

        Create_Summary(grdSPTSFOC3, "EVENT_GROUP_LNO", "Count")
        grdSPTSFOC3.DisplayLayout.Bands(0).Override.SummaryFooterCaptionVisible = DefaultableBoolean.False

        With grdSPTSFOC3.DisplayLayout.Bands("SPTSFOC3")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"", ""}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("EVENT_GROUP_NO").Header.Fixed = True
        End With

        With grdSPTSFOCX.DisplayLayout.Bands("SPTSFOCX")
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                GCOL.Header.Appearance.BackColor = System.Drawing.Color.White
                GCOL.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"EVENT_GROUP_NO"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                ElseIf New String() {"COLLECTION_CODE"}.Contains(GCOL.Key) Then
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                Else
                    GCOL.Header.Appearance.BackColor2 = System.Drawing.Color.LightBlue
                End If
            Next
            .Columns("EVENT_GROUP_NO").Header.Fixed = True
        End With

        With grdSPTSFOC9.DisplayLayout.Bands("SPTSFOC9")

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                If GCOL.Key = "SEL" Then
                    GCOL.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    GCOL.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
            '.Columns("EVENT_GROUP_NO").Header.Fixed = True
        End With

        ASCMAIN1.Add_Value_List(grdSPTSFOCX, "APPR_STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'APPR_STATUS_CODE'")
        ASCMAIN1.Add_Value_List(grdSPTSFOCX, "STATUS_CODE", "Select T_CODE, T_DESC from ASTCODE1 where TABLE_NAME = 'SPTCOOP1' and COLUMN_NAME = 'STATUS_CODE'")


        ASCMAIN1.sql = "Select * from SPTTYPE1"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim SECURITY_CODE As String = row.Item("SECURITY_CODE") & ""
            Dim EXPENSE_TYPE_CODE As String = row.Item("EXPENSE_TYPE_CODE") & ""
            If ASCMAIN1.CLIENT = "INT" Then
                If SECURITY_CODE <> "" And ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
                    EXPENSE_TYPE_CODEs_I_may_approve.Add(EXPENSE_TYPE_CODE)
                End If
            Else
                If SECURITY_CODE = "" Then
                    EXPENSE_TYPE_CODEs_I_may_approve.Add(EXPENSE_TYPE_CODE)
                Else
                    If SECURITY_CODE <> "" And ASCMAIN1.USER_SECURITY_CODEs.Contains(SECURITY_CODE) Then
                        EXPENSE_TYPE_CODEs_I_may_approve.Add(EXPENSE_TYPE_CODE)
                    End If
                End If
            End If

        Next


        grpHeader.Visible = False

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 24) & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(24)

        Show_Filter(grdSPTSFOCX, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                ' Validate_Code("ATTR_CODE")

                If Absx1.txtFor("ATTR_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Attribute Code"
                Else
                    Dim row As DataRow = LookUp("SPTATTR1", Absx1.txtFor("ATTR_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Attribute Code Entered Is Not Valid"
                    Else
                        If row.Item("ATTR_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Attribute Status Is Not Active"
                        End If
                    End If
                End If

                Dim DT As Date = Absx1.dteFor("AUTH_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Event Date is Mandatory"
                Else
                    ' TAC.SOCMAIN1.Validate_Invoice_Date(DT, 12, 1, EMsg)
                End If

                If Absx1.txtFor("VEHICLE_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Vehicle Code"
                Else
                    Dim row As DataRow = LookUp("SPTAVEH1", Absx1.txtFor("VEHICLE_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Vehicle (Type) Entered Is Not Valid"
                    Else
                        If Absx1.txtFor("VEHICLE_CODE").Text <> "BF" And Absx1.txtFor("VEHICLE_CODE").Text <> "MA" Then
                            EMsg &= vbCr & "Valid Vehicle (Types) are BF and MA"
                        End If
                        'If rowARTCUST1.Item("VEHICLE_STATUS").ToString <> "A" Then
                        '    EMsg &= vbCr & "Vehicle (Type) Entered Is Not Active"
                        'End If
                    End If
                End If

                If Absx1.txtFor("BOOKING_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "You must enter a Description for this Event"
                End If

            Case "View", "Edit"
                EVENT_GROUP_NO = Absx1.txtFor("EVENT_GROUP_NO").Text
                If EVENT_GROUP_NO = "" Then
                    EMsg &= vbCr & "You must specify a Event Group No to View"
                Else
                    rowSPTSFOC1 = LookUp("SPTSFOC1", EVENT_GROUP_NO)
                    If rowSPTSFOC1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Event Group " & EVENT_GROUP_NO & " on File"
                    Else

                        If eItemKey = "Edit" Then

                            If rowSPTSFOC1.Item("APPR_STATUS_CODE") & "" = "A" Then
                                Dim EXPENSE_TYPE_CODE As String = rowSPTSFOC1.Item("EXPENSE_TYPE_CODE") & ""
                                If Not EXPENSE_TYPE_CODEs_I_may_approve.Contains(EXPENSE_TYPE_CODE) Then
                                    EMsg &= vbCr & "Event Group No " & EVENT_GROUP_NO & " has already been approved." & vbCr & "No Changes (except for by Approver)"
                                End If
                            End If

                            If EMsg = "" Then
                                If Not ASCMAIN1.Logical_Lock("SPTSFOC1", EVENT_GROUP_NO) Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                End If

            Case "Update"


                If Absx1.txtFor("ATTR_CODE").Text.Length = 0 Then
                    EMsg &= vbCr & "You must supply a Valid Attribute Code"
                Else
                    Dim row As DataRow = LookUp("SPTATTR1", Absx1.txtFor("ATTR_CODE").Text)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Attribute Code Entered Is Not Valid"
                    Else
                        If row.Item("ATTR_STATUS").ToString <> "A" Then
                            EMsg &= vbCr & "Attribute Status Is Not Active"
                        End If
                    End If
                End If

                Dim DT As Date = Absx1.dteFor("AUTH_DATE").Value
                If DT & "" = "" Then
                    EMsg &= vbCr & "Event Date is Mandatory"
                Else
                    '  TAC.SOCMAIN1.Validate_Invoice_Date(DT, 2, 1, EMsg)
                End If

                Dim VEHICLE_CODE As String = Absx1.txtFor("VEHICLE_CODE").Text
                If VEHICLE_CODE = "" Then
                    EMsg &= vbCr & "You must supply a Valid Vehicle Code"
                Else
                    Dim row As DataRow = LookUp("SPTAVEH1", VEHICLE_CODE)
                    If IsNothing(row) Then
                        EMsg &= vbCr & "Vehicle (Type) Entered Is Not Valid"
                    Else
                        If VEHICLE_CODE <> "BF" And VEHICLE_CODE <> "MA" Then
                            EMsg &= vbCr & "Valid Vehicle (Types) are BF and MA"
                        End If
                        'If rowARTCUST1.Item("VEHICLE_STATUS").ToString <> "A" Then
                        '    EMsg &= vbCr & "Vehicle (Type) Entered Is Not Active"
                        'End If
                    End If
                End If

                If Absx1.txtFor("BOOKING_NAME").Text.Length = 0 Then
                    EMsg &= vbCr & "You must enter a Description for this Event"
                End If

                If grdSPTSFOC3.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items or Collection Details Entered"
                Else

                End If


                If VEHICLE_CODE = "BF" Then ' OK TO HAVE NO STORES FOR BF
                Else
                    If grdSPTSFOC9.Rows.Count = 0 Then
                        EMsg &= vbCr & "No Stores Listed"
                    Else
                        If dst.Tables("SPTSFOC9").Select("SEL = '1'").Length = 0 Then
                            EMsg &= vbCr & "No Stores Selected"
                        End If
                    End If
                End If


                If Absx1.txtFor("SEASON_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Specify a Season"
                Else
                    Dim row As DataRow = LookUp("ICTSEAS1", Absx1.txtFor("SEASON_CODE").Text)
                    If row Is Nothing Then
                        EMsg &= vbCr & "Invalid Value Specified for Season"
                    Else
                        If EMsg = "" Then
                            Dim SEASON_CODE As String = Absx1.txtFor("SEASON_CODE").Text
                            If Mid(SEASON_CODE, 1, 4) <> Format(DT, "yyyy") Then
                                EMsg &= vbCr & "Season not congruous with Start Date"
                            Else
                                If Mid(SEASON_CODE, 5, 1) = "S" And Format(DT, "MM") >= "07" Then
                                    EMsg &= vbCr & "Season not congruous with Event Date"
                                End If
                                If Mid(SEASON_CODE, 5, 1) = "F" And Format(DT, "MM") < "07" Then
                                    EMsg &= vbCr & "Season not congruous with Event Date"
                                End If
                            End If

                        End If

                    End If
                End If


            Case "Clone"
                If MsgBox("Are you sure you want Clone this agreement to a new record?", MsgBoxStyle.YesNo,
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
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Clone"
                Clone_Record()
                Mode_Settings(False)
                Absx1.txtFor("EVENT_GROUP_NO").Text = EVENT_GROUP_NO_new
                Click_Command("View")

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

                    If EntryMode = "V" And ScreenMode Then ' And (optShow.Value <> "G") Then
                        .Items("Clone").Visible = True And Not InquiryMode
                    Else
                        .Items("Clone").Visible = False
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

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode
        cmdBrowse.Visible = (EntryMode = "E" Or EntryMode = "N")

        grdSPTSFOCX.Visible = Not ScreenMode

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
                Set_Read_Only_for_ctl(Absx1.txtFor("BOOKING_NAME"), False)
                Set_Read_Only_for_ctl(Absx1.dteFor("AUTH_DATE"), False)
            End If

            For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTSFOC3, grdSPTSFOC9}
                If EntryMode = "N" Or EntryMode = "E" Then
                    With grd.DisplayLayout.Override
                        If grd.Name = "grdSPTSFOC9" Then
                            .AllowAddNew = UltraWinGrid.AllowAddNew.No
                            .AllowDelete = DefaultableBoolean.False
                        Else
                            .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                            .AllowDelete = DefaultableBoolean.True
                        End If
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Else
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                End If
            Next


        Else
            Clear_Record()
            grdSPTSFOC9.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SPTSFOC1", "SPTSFOC3", "SPTSFOC9", "SPTCOOP1", "SPTCOOP3", "SPTCOOPB"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Setup_cbeYP()
        Refresh_Documents()

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            rowSPTSFOC1 = dst.Tables("SPTSFOC1").NewRow
            EVENT_GROUP_NO = ASCMAIN1.Next_Control_No("SPTSFOC1.EVENT_GROUP_NO")
            rowSPTSFOC1.Item("EVENT_GROUP_NO") = EVENT_GROUP_NO
            rowSPTSFOC1.Item("VEHICLE_CODE") = HFs("VEHICLE_CODE")
            rowSPTSFOC1.Item("ATTR_CODE") = HFs("ATTR_CODE")
            rowSPTSFOC1.Item("BOOKING_NAME") = HFs("BOOKING_NAME")
            rowSPTSFOC1.Item("AUTH_DATE") = HFs("AUTH_DATE")
            rowSPTSFOC1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
            rowSPTSFOC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowSPTSFOC1.Item("INIT_DATE") = DATETIME_STAMP
            rowSPTSFOC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowSPTSFOC1.Item("LAST_DATE") = DATETIME_STAMP
            rowSPTSFOC1.Item("APPR_STATUS_CODE") = "P"
            rowSPTSFOC1.Item("STATUS_CODE") = "O"
            rowSPTSFOC1.Item("EXPENSE_TYPE_CODE") = "RTLEVENTS"

            rowSPTSFOC1.Item("EVENT_DATE_CHANGED") = DATETIME_STAMP.Date

            Dim DT As Date = Absx1.dteFor("AUTH_DATE").Value
            Dim SEASON_CODE As String = Format(DT, "yyyy") & IIf(Format(DT, "MM") >= "07", "F", "S")
            rowSPTSFOC1.Item("SEASON_CODE") = SEASON_CODE

            dst.Tables("SPTSFOC1").Rows.Add(rowSPTSFOC1)

        Else
            rowSPTSFOC1 = Fill_Record("SPTSFOC1", Absx1.txtFor("EVENT_GROUP_NO").Text)
            dst.AcceptChanges()
        End If

        STATUS_CODE = rowSPTSFOC1.Item("STATUS_CODE")

        EnforceConstraints(False)

        Fill_Records("SPTSFOC3", EVENT_GROUP_NO)
        Sort_grdColumns(grdSPTSFOC3, "EVENT_GROUP_LNO")

        Fill_Records("SPTSFOC9", EVENT_GROUP_NO)
        For Each row As DataRow In dst.Tables("SPTSFOC9").Select
            row.Item("SEL") = "1"
        Next
        dst.Tables("SPTSFOC9").AcceptChanges()
        grdSPTSFOC9.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdSPTSFOC9, "CUST_CODE,CUST_STORE_NO")

        EnforceConstraints(True)

        Setup_Retail_Weeks()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        BeginTrans()

        Dim OPS_YYYYPP As String = Format(rowSPTSFOC1.Item("AUTH_DATE"), "yyyyMM")
        rowSPTSFOC1.Item("OPS_YYYYPP") = OPS_YYYYPP
        rowSPTSFOC1.Item("OPS_YYYYWW") = Set_OPS_YYYYWW()
        rowSPTSFOC1.Item("DATE_START") = rowSPTSFOC1.Item("AUTH_DATE")
        rowSPTSFOC1.Item("DATE_END") = rowSPTSFOC1.Item("AUTH_DATE")

        If EntryMode = "E" Then
            If STATUS_CODE <> Absx1.optFor("STATUS_CODE").Value Then
                If STATUS_CODE = "O" Then
                    ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CLSD", "Contract Closed", "")
                Else
                    ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "REOP", "Contract Re-Opened", "")
                End If
            End If

            If rowSPTSFOC1.Item("APPR_STATUS_CODE") & "" <> rowSPTSFOC1.Item("APPR_STATUS_CODE", DataRowVersion.Original) & "" Then
                ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPRSTA", $"Appr Status Manually Changed from {rowSPTSFOC1.Item("APPR_STATUS_CODE", DataRowVersion.Original)} to {rowSPTSFOC1.Item("APPR_STATUS_CODE")}", "")

                ASCMAIN1.sql = $"Select * from SPTCOOP1 where EVENT_GROUP_NO = '{EVENT_GROUP_NO}'"
                Dim rowSPTCOOP1() As DataRow = ASCDATA1.GetDataTable().Select("")
                Dim AUTH_NO As String = rowSPTCOOP1(0).Item("AUTH_NO") & ""
                ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPRSTA", $"Appr Status Manually Changed from {rowSPTSFOC1.Item("APPR_STATUS_CODE", DataRowVersion.Original)} to {rowSPTSFOC1.Item("APPR_STATUS_CODE")}", "")
            End If

        End If

        If EntryMode = "N" Then
            ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "ADD", "Contract Created", "")
        Else
            ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CHG", "Contract Updated", "")
        End If

        If EntryMode <> "N" Then Delete_Records()

        Dim SQLD As String = "EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
        INIT_LAST("SPTSFOC1", False, , True)

        Update_Record_TDA("SPTSFOC1", SQLD)
        Update_Record_TDA("SPTSFOC3", SQLD)

        ASCDATA1.DeleteRows(dst.Tables("SPTSFOC9"), "ISNULL(SEL,'0')<>'1'")
        Update_Record_TDA("SPTSFOC9", SQLD)


        ' Update SPTCOOP1/3/B

        ASCMAIN1.sql = "Update SPTCOOP1 Set STATUS_CODE = 'C' where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from SPTCOOP3 where AUTH_NO in (Select AUTH_NO from SPTCOOP1 where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "')"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from SPTCOOPB where AUTH_NO in (Select AUTH_NO from SPTCOOP1 where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "')"
        ASCDATA1.ExecuteSQL()

        Fill_Records("SPTCOOP1", EVENT_GROUP_NO)

        For Each rowC As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SPTSFOC9").Select("SEL='1'"), "CUST_CODE").Select
            Dim CUST_CODE As String = rowC.Item("CUST_CODE")
            Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
            Dim rowSPTCOOP1s() As DataRow = dst.Tables("SPTCOOP1").Select("CUST_CODE = '" & CUST_CODE & "'")
            Dim rowSPTCOOP1 As DataRow = Nothing
            Dim AUTH_NO As String = ""
            If rowSPTCOOP1s.Length <> 0 Then
                rowSPTCOOP1 = rowSPTCOOP1s(0)
                AUTH_NO = rowSPTCOOP1.Item("AUTH_NO")
            Else
                rowSPTCOOP1 = dst.Tables("SPTCOOP1").NewRow
                AUTH_NO = ASCMAIN1.Next_Control_No("SPTCOOP1.AUTH_NO")
                With rowSPTCOOP1
                    .Item("AUTH_NO") = AUTH_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("APPR_STATUS_CODE") = "P"
                    .Item("EXPENSE_TYPE_CODE") = rowSPTSFOC1.Item("EXPENSE_TYPE_CODE")
                    .Item("VEHICLE_CODE") = rowSPTSFOC1.Item("VEHICLE_CODE")
                    '.Item("ATTR_CODE") = rowSPTSFOC1.Item("ATTR_CODE")
                    .Item("EVENT_GROUP_NO") = EVENT_GROUP_NO
                End With
                dst.Tables("SPTCOOP1").Rows.Add(rowSPTCOOP1)
            End If

            'AUTH_APPR_DATE
            'AUTH_APPR_BY
            'AUTH_APPR_AMT
            'AUTH_APPR_NOTES

            With rowSPTCOOP1
                .Item("AUTH_DATE") = rowSPTSFOC1.Item("AUTH_DATE")
                .Item("DATE_START") = rowSPTSFOC1.Item("AUTH_DATE")
                .Item("DATE_END") = rowSPTSFOC1.Item("AUTH_DATE")
                .Item("BOOKING_NAME") = rowSPTSFOC1.Item("BOOKING_NAME")
                .Item("SEASON_CODE") = rowSPTSFOC1.Item("SEASON_CODE")
                .Item("SALES_DIVISION_CODE") = rowSPTSFOC1.Item("SALES_DIVISION_CODE")
                .Item("OPS_YYYYPP") = rowSPTSFOC1.Item("OPS_YYYYPP")
                .Item("OPS_YYYYWW") = rowSPTSFOC1.Item("OPS_YYYYWW")
                .Item("EVENT_TYPE_CODE") = rowSPTSFOC1.Item("EVENT_TYPE_CODE")
                .Item("NOTES") = rowSPTSFOC1.Item("NOTES")
                .Item("EVENT_FILE_LINK") = rowSPTSFOC1.Item("EVENT_FILE_LINK")
                .Item("STATUS_CODE") = rowSPTSFOC1.Item("STATUS_CODE")
                .Item("APPR_STATUS_CODE") = rowSPTSFOC1.Item("APPR_STATUS_CODE")

                .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE")

                .Item("EVENT_DATE_CHANGED") = rowSPTSFOC1.Item("EVENT_DATE_CHANGED")

                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
            End With

            For Each rowSPTSFOC3 As DataRow In dst.Tables("SPTSFOC3").Select("", "")
                Dim rowSPTCOOP3 As DataRow = dst.Tables("SPTCOOP3").NewRow
                With rowSPTCOOP3
                    .Item("AUTH_NO") = AUTH_NO
                    .Item("AUTH_LNO") = rowSPTSFOC3.Item("EVENT_GROUP_LNO")
                    .Item("ITEM_CODE") = rowSPTSFOC3.Item("ITEM_CODE")
                    .Item("COLLECTION_CODE") = rowSPTSFOC3.Item("COLLECTION_CODE")
                    .Item("FEATURE_DESC") = rowSPTSFOC3.Item("FEATURE_DESC")
                End With
                dst.Tables("SPTCOOP3").Rows.Add(rowSPTCOOP3)
            Next

            For Each rowSPTSFOC9 As DataRow In dst.Tables("SPTSFOC9").Select("CUST_CODE = '" & CUST_CODE & "' and SEL='1'", "")
                Dim rowSPTCOOPB As DataRow = dst.Tables("SPTCOOPB").NewRow
                With rowSPTCOOPB
                    .Item("AUTH_NO") = AUTH_NO
                    .Item("CUST_STORE_NO") = rowSPTSFOC9.Item("CUST_STORE_NO")
                End With
                dst.Tables("SPTCOOPB").Rows.Add(rowSPTCOOPB)
            Next

        Next

        Update_Record_TDA("SPTCOOP1")
        Update_Record_TDA("SPTCOOP3")
        Update_Record_TDA("SPTCOOPB")

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
            {"SPTSFOC1", "SPTSFOC3", "SPTSFOC9"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
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
            E.TABLE_NAME = "SPTSFOC1"
            E.COLUMN_NAME = "EVENT_GROUP_NO"
            E.CODE_VALUE = Absx1.txtFor("EVENT_GROUP_NO").Text
            E.DESC_VALUE = Absx1.txtFor("VEHICLE_CODE").Text
            E.ATTACHMENT_NOTES = ""
            E.READ_ONLY = False
        End If

        Return E
    End Function

    Public Overrides Function Log_Context() As Log_Entity

        Dim E As New Log_Entity

        E.TABLE_NAME = "SPTSFOC1"
        E.TABLE_KEY_CAPTION = "Store Focus Events"
        If ScreenMode Then
            E.enabled = True
            E.read_only = False
            E.TABLE_KEY = Absx1.txtFor("EVENT_GROUP_NO").Text '  HFs("CUST_CODE")
            E.TABLE_KEY_DESC = Absx1.txtFor("VEHICLE_CODE").Text
            E.TABLE_KEY_locked = ScreenMode And (EntryMode = "E")
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
            Case "VEHICLE_CODE"
                sql_where = "VEHICLE_CODE in ('BF','MA')"
        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSPTSFOCX, "SS", "Show Filter", "Show GroupBox", "Approve", "Move to Pending") ', "Move to Pending")
        Load_Popup_Menu(grdSPTSFOC9, "SSBBBB", "Show Filter", "Show GroupBox", "Load Stores w/Attribute", "Load All Stores", "Select All", "De-Select All")
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

            Case "grdSPTSFOC9"
                tlb_btn = tlb_pop.Tools("Load All Stores")
                If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_btn.SharedProps.Visible = True
                Else
                    tlb_btn.SharedProps.Visible = False
                End If
                tlb_btn = tlb_pop.Tools("Load Stores w/Attribute")
                If ScreenMode And (EntryMode = "N" Or EntryMode = "E") Then
                    tlb_btn.SharedProps.Visible = True
                Else
                    tlb_btn.SharedProps.Visible = False
                End If

            Case "grdSPTSFOCX"

                If Not InquiryMode Then ' And (optShow.Value = "P" Or optShow.Value = "G") Then
                    'tlb_pop.Tools("Move to Pending").SharedProps.Visible = (optShow.Value = "P")
                    tlb_pop.Tools("Approve").SharedProps.Visible = True
                Else
                    'tlb_pop.Tools("Move to Pending").SharedProps.Visible = False
                    tlb_pop.Tools("Approve").SharedProps.Visible = False
                End If



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
                For Each grow As UltraWinGrid.UltraGridRow In grdSPTSFOC9.Rows
                    grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                    grow.Update()
                Next
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Load All Stores"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading All Stores")

                ASCMAIN1.sql = "Select ARTCUST2.*,SOTSELL1.REGION_CODE" & vbCrLf _
                    & ",ARTCUST1.TRADE_CLASS_CODE,SOTTCLS1.CHANNEL_CODE,ARTCUST2.SELL_CODE_AC" & vbCrLf _
                    & " from ARTCUST2,ARTCUST1,SOTTCLS1,SOTSELL1" & vbCrLf _
                    & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
                    & "   And SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                    & "   And SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                    & "   And SOTTCLS1.CHANNEL_CODE = '1'"
                For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim CUST_CODE As String = ROW.Item("CUST_CODE")
                    Dim CUST_STORE_NO As String = ROW.Item("CUST_STORE_NO")
                    If dst.Tables("SPTSFOC9").Rows.Find(New String() {EVENT_GROUP_NO, CUST_CODE, CUST_STORE_NO}) Is Nothing Then
                        Dim rowSPTSFOC9 As DataRow = dst.Tables("SPTSFOC9").NewRow
                        ' Dim rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                        rowSPTSFOC9.Item("EVENT_GROUP_NO") = EVENT_GROUP_NO
                        rowSPTSFOC9.Item("CUST_CODE") = CUST_CODE
                        rowSPTSFOC9.Item("CUST_STORE_NO") = CUST_STORE_NO
                        rowSPTSFOC9.Item("SEL") = "0"
                        rowSPTSFOC9.Item("CUST_STORE_LOCATION") = ROW.Item("CUST_STORE_LOCATION")
                        rowSPTSFOC9.Item("CUST_STORE_NAME") = ROW.Item("CUST_STORE_NAME")
                        rowSPTSFOC9.Item("SELL_CODE") = ROW.Item("SELL_CODE")
                        rowSPTSFOC9.Item("REGION_CODE") = ROW.Item("REGION_CODE")
                        rowSPTSFOC9.Item("SELL_CODE_AC") = ROW.Item("SELL_CODE_AC")
                        dst.Tables("SPTSFOC9").Rows.Add(rowSPTSFOC9)
                    End If
                Next
                Sort_grdColumns(grdSPTSFOC9, "CUST_CODE, CUST_STORE_NO")

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Load Stores w/Attribute"
                If dst.Tables("SPTSFOC3").Select("").Length = 0 Then
                    MsgBox("No Items or Collections entered", MsgBoxStyle.OkOnly, "Cannot Infer Brand/Gender")
                Else
                    Dim BGs As New List(Of String)
                    For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SPTSFOC3"), New String() {"COLLECTION_CODE"}).Select("")
                        Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")
                        Dim rowICTCOLL1 As DataRow = LookUp("ICTCOLL1", COLLECTION_CODE)
                        Dim BRAND_CODE As String = rowICTCOLL1.Item("BRAND_CODE")
                        Dim BRAND_GENDER As String = rowICTCOLL1.Item("COLLECTION_GENDER")
                        Dim BG As String = BRAND_CODE & "_" & BRAND_GENDER
                        If Not BGs.Contains(BG) Then
                            BGs.Add(BG)
                        End If
                    Next

                    Dim ATTR_CODE As String = Absx1.txtFor("ATTR_CODE").Text
                    For Each BG As String In BGs
                        Dim BRAND_CODE As String = Split(BG, "_")(0)
                        Dim BRAND_GENDER As String = Split(BG, "_")(1)

                        ASCMAIN1.sql = "Select ARTCUST2.*,SOTSELL1.REGION_CODE" & vbCrLf _
                            & ",ARTCUST1.TRADE_CLASS_CODE,SOTTCLS1.CHANNEL_CODE,ARTCUST2.SELL_CODE_AC" & vbCrLf _
                            & " from ARTCUST2,ARTCUST1,SOTTCLS1,SOTSELL1" & vbCrLf _
                            & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
                            & "   And SOTSELL1.SELL_CODE (+) = ARTCUST2.SELL_CODE" & vbCrLf _
                            & "   And SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                            & "   And SOTTCLS1.CHANNEL_CODE = '1'" & vbCrLf _
                            & " and (ARTCUST2.CUST_CODE, ARTCUST2.CUST_STORE_NO) in " & vbCrLf _
                            & " (Select CUST_CODE, CUST_STORE_NO From SPTATTR2" & vbCrLf _
                            & " where BRAND_CODE = '" & BRAND_CODE & "' and BRAND_GENDER = '" & BRAND_GENDER & "' and ATTR_CODE = '" & ATTR_CODE & "')"

                        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("")
                            Dim CUST_CODE As String = ROW.Item("CUST_CODE")
                            Dim CUST_STORE_NO As String = ROW.Item("CUST_STORE_NO")
                            If dst.Tables("SPTSFOC9").Rows.Find(New String() {EVENT_GROUP_NO, CUST_CODE, CUST_STORE_NO}) Is Nothing Then
                                Dim rowSPTSFOC9 As DataRow = dst.Tables("SPTSFOC9").NewRow
                                ' Dim rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                                rowSPTSFOC9.Item("EVENT_GROUP_NO") = EVENT_GROUP_NO
                                rowSPTSFOC9.Item("CUST_CODE") = CUST_CODE
                                rowSPTSFOC9.Item("CUST_STORE_NO") = CUST_STORE_NO
                                rowSPTSFOC9.Item("SEL") = "1"
                                rowSPTSFOC9.Item("CUST_STORE_LOCATION") = ROW.Item("CUST_STORE_LOCATION")
                                rowSPTSFOC9.Item("CUST_STORE_NAME") = ROW.Item("CUST_STORE_NAME")
                                rowSPTSFOC9.Item("SELL_CODE") = ROW.Item("SELL_CODE")
                                rowSPTSFOC9.Item("REGION_CODE") = ROW.Item("REGION_CODE")
                                rowSPTSFOC9.Item("SELL_CODE_AC") = ROW.Item("SELL_CODE_AC")
                                dst.Tables("SPTSFOC9").Rows.Add(rowSPTSFOC9)
                            End If
                        Next
                        Sort_grdColumns(grdSPTSFOC9, "CUST_CODE, CUST_STORE_NO")

                    Next
                End If



            Case "Move to Pending", "Approve"
                If grdSPTSFOCX.Selected.Rows.Count = 0 Then
                    If grdSPTSFOCX.ActiveRow IsNot Nothing Then
                        grdSPTSFOCX.ActiveRow.Selected = True
                    End If
                End If

                If grdSPTSFOCX.Selected.Rows.Count = 0 Then
                    MsgBox("Nothing Selected to " & e.Tool.Key, MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                Else
                    Dim cannot_proceed_msg As String = ""
                    Dim EVENT_GROUP_NOs As New Dictionary(Of String, Decimal)
                    Dim can_proceed As Boolean = True
                    Dim TOTAL_AMTs As Decimal = 0

                    Dim AUTH_NOsVerified As New List(Of String)

                    For Each grow As UltraWinGrid.UltraGridRow In grdSPTSFOCX.Selected.Rows
                        Dim EXPENSE_TYPE_CODE As String = grow.Cells("EXPENSE_TYPE_CODE").Value
                        Dim EVENT_GROUP_NO As String = grow.Cells("EVENT_GROUP_NO").Value

                        If EVENT_GROUP_NOs.ContainsKey(EVENT_GROUP_NO) Then
                        Else
                            rowSPTSFOC1 = LookUp("SPTSFOC1", EVENT_GROUP_NO)
                            Dim APPR_STATUS_CODE As String = rowSPTSFOC1.Item("APPR_STATUS_CODE") & ""


                            If APPR_STATUS_CODE = "A" Then
                                If e.Tool.Key = "Move to Pending" Or e.Tool.Key = "Move to Preliminary" Then
                                    ' ok to proceed
                                Else
                                    cannot_proceed_msg = vbCrLf & vbCrLf & "Event Group " & EVENT_GROUP_NO & " may Not use " & e.Tool.Key
                                    can_proceed = False
                                End If

                            ElseIf APPR_STATUS_CODE = "G" Or APPR_STATUS_CODE = "P" Then
                                'cannot_proceed_msg = vbCrLf & vbCrLf & "Event Group " & EVENT_GROUP_NO & " Is Not Pending Approval"
                                'can_proceed = False
                                'Exit For
                                If e.Tool.Key = "Move to Pending" Or e.Tool.Key = "Move to Preliminary" Then
                                    cannot_proceed_msg = vbCrLf & vbCrLf & "Event Group " & EVENT_GROUP_NO & " may Not use " & e.Tool.Key
                                    can_proceed = False
                                    Exit For
                                Else
                                    ' ok to proceed
                                End If
                            End If

                            'If APPR_STATUS_CODE <> "G" And APPR_STATUS_CODE <> "P" Then
                            '    MsgBox("Store Focus Event " & EVENT_GROUP_NO & " is NOT Pending Approval", MsgBoxStyle.OkOnly, "Cannot Proceed")
                            '    can_proceed = False
                            '    Exit For
                            'End If

                            ASCMAIN1.sql = $"Select * from SPTCOOP1 where EVENT_GROUP_NO = '{EVENT_GROUP_NO}'"
                            Dim rowSPTCOOP1() As DataRow = ASCDATA1.GetDataTable().Select("")
                            If rowSPTCOOP1(0).Item("VERIFIED_AS_OPEN_COMMENTS") & "" <> "" Then
                                AUTH_NOsVerified.Add(EVENT_GROUP_NO)
                            End If

                            If EXPENSE_TYPE_CODEs_I_may_approve.Contains(EXPENSE_TYPE_CODE) Then

                                If Not ASCMAIN1.Logical_Lock("SPTSFOC1", EVENT_GROUP_NO) Then
                                    can_proceed = False
                                    Exit For
                                End If

                                EVENT_GROUP_NOs.Add(EVENT_GROUP_NO, 0)
                            Else
                                can_proceed = False
                                Exit For
                            End If
                        End If

                    Next

                    If Not can_proceed Then
                        MsgBox("Cannot Proceed" & cannot_proceed_msg, MsgBoxStyle.OkOnly, "Cannot " & e.Tool.Key)
                    Else

                        AUTH_APPR_NOTES = ""

                        Dim actionMesg = ""
                        Dim APPROVAL_STATUS_new As String = ""
                        Dim defaultNote As String = ""
                        If e.Tool.Key = "Approve" Then
                            actionMesg = "Approving"
                            APPROVAL_STATUS_new = "A"
                            defaultNote = "Approved"
                        ElseIf e.Tool.Key = "Move to Pending" Then
                            actionMesg = "Moving To Pending"
                            APPROVAL_STATUS_new = "G"
                            defaultNote = "Moving To Pending"
                        ElseIf e.Tool.Key = "Move to Preliminary" Then
                            actionMesg = "Moving To Preliminary"
                            APPROVAL_STATUS_new = "P"
                            defaultNote = "Move to Preliminary"
                        Else
                            ' Not sure what else there would be
                        End If


                        If actionMesg = "" Then
                            MsgBox("There Is an issue With this Action: " & e.Tool.Key, MsgBoxStyle.OkOnly)
                        Else
                            actionMesg = Chr(34) & actionMesg & Chr(34)
                            Dim contracts As String = ""
                            If EVENT_GROUP_NOs.Count = 1 Then
                                contracts = CStr(EVENT_GROUP_NOs.Count) & " Contract"
                            Else
                                contracts = CStr(EVENT_GROUP_NOs.Count) & " Contracts at Once"
                            End If

                            Dim lblv As String = ""
                            If AUTH_NOsVerified.Count > 1 And e.Tool.Key <> "Approve" Then
                                lblv = vbCrLf & vbCrLf & "***************************" & vbCrLf & $"Please Note:" & vbCrLf & $"There are {AUTH_NOsVerified.Count}" & vbCrLf & " *Previously Verified*" & vbCrLf & "Contracts in the range selected" & vbCrLf & "***************************"
                            End If

                            Dim LBL As String = $"You are {actionMesg} " & vbCrLf & Contracts _
                                                & vbCrLf & "Total Amount is " & Format(TOTAL_AMTs, "$#,##0.00") & lblv _
                                                & vbCrLf & vbCrLf & "Enter Notes to Record with this Action"

                            AUTH_APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, $"OK To '{e.Tool.Key}' these Events?", False, 60, defaultNote)


                            'Dim LBL As String = "You are approving " & CStr(EVENT_GROUP_NOs.Count) & " Store Focus Events at Once" _
                            '                    & vbCrLf & vbCrLf & "Enter Notes to Record with this Approval"

                            'AUTH_APPR_NOTES = ASCMAIN1.Get_txt_from_User(LBL, "OK To Approve these Events?", False, 60, "Approved")

                            If AUTH_APPR_NOTES <> "" Then
                                BeginTrans()
                                DATETIME_STAMP = Now + ASCMAIN1.NowTSD
                                For Each EVENT_GROUP_NO As String In EVENT_GROUP_NOs.Keys

                                    Dim rowSPTSFOC1 As DataRow = LookUp("SPTSFOC1", EVENT_GROUP_NO)
                                    Dim APPROVAL_STATUS_orig As String = rowSPTSFOC1.Item("APPR_STATUS_CODE")

                                    Approve_Record(EVENT_GROUP_NO, APPROVAL_STATUS_new, APPROVAL_STATUS_orig, EVENT_GROUP_NOs(EVENT_GROUP_NO))
                                Next
                                CommitTrans($"{e.Tool.Key} Complete")

                                Refresh_Documents()
                            End If

                        End If
                        ASCMAIN1.MultiTask_Release()
                    End If
                End If

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
            Case "EVENT_GROUP_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("View", e)
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

    Private Sub grdSPTSFOC3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTSFOC3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                '   If cdr IsNot Nothing Then
                Dim ITEM_CODE As String = CStr(e.Cell.Value & "").ToUpper
                ' Dim ITEM_CODE As String = Validate_Item(e.Cell.Value & "") ' grdSOTORDR2.ActiveRow.Cells("ITEM_CODE").Value)

                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    If e.Cell.Value & "" <> ITEM_CODE Then
                        e.Cell.Row.Cells("ITEM_CODE").Value = ITEM_CODE
                    End If
                    e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                    e.Cell.Row.Cells("FEATURE_DESC").Value = rowICTITEM1.Item("ITEM_DESC")
                    e.Cell.Row.Cells("COLLECTION_CODE").Value = rowICTITEM1.Item("COLLECTION_CODE")
                End If

                '   End If

            Case "COLLECTION_CODE"

                grdCodeDesc(grdSPTSFOC3, "ICTCOLL1", "COLLECTION_CODE", "COLLECTION_NAME")
                If cdr IsNot Nothing Then
                    Dim COLLECTION_CODE As String = e.Cell.Value
                    Dim BRAND_CODE As String = cdr.Item("BRAND_CODE")
                    Dim rowICTBRAN1 As DataRow = LookUp("ICTBRAN1", BRAND_CODE)
                    e.Cell.Row.Cells("BRAND_CODE").Value = BRAND_CODE
                    e.Cell.Row.Cells("BRAND_NAME").Value = rowICTBRAN1.Item("BRAND_NAME")
                Else
                    grdSPTSFOC3.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

        End Select
    End Sub

    Private Sub grdSPTSFOC3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdSPTSFOC3.AfterRowActivate
        With grdSPTSFOC3.DisplayLayout.Bands(0)
            If grdSPTSFOC3.ActiveRow.IsAddRow Then
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdSPTSFOC3.ActiveCell = grdSPTSFOC3.ActiveRow.Cells("COLLECTION_CODE")
                grdSPTSFOC3.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdSPTSFOC3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSPTSFOC3.BeforeExitEditMode
        If grdSPTSFOC3.ActiveCell Is Nothing Then Exit Sub
        With grdSPTSFOC3.ActiveCell
            Select Case .Column.Key
                Case "COLLECTION_CODE"
                    If .Text <> "" Then
                        If .Value IsNot Nothing Then
                            .Value = .Text.ToUpper
                        End If

                    End If
                    If .Text <> "" Then
                        cdr = LookUp("ICTCOLL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid Collection Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSPTSFOC3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSPTSFOC3.BeforeRowUpdate
        With grdSPTSFOC3
            If e.Row.Cells("COLLECTION_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                ' e.Cancel = True MAGIC UNICORN IS A COMMENTED LINE
            Else
                LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Collection Code (" & e.Row.Cells("COLLECTION_CODE").Text & ")",
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("EVENT_GROUP_NO").Text = "" Then
                    .ActiveRow.Cells("EVENT_GROUP_NO").Value = Absx1.CtlFor("EVENT_GROUP_NO").Text
                    .ActiveRow.Cells("EVENT_GROUP_LNO").Value = Val(dst.Tables("SPTSFOC3").Compute("Max(EVENT_GROUP_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdSPTSFOC3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSPTSFOC3.ClickCellButton

        If grdSPTSFOC3.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
            Case "COLLECTION_CODE"
                sql_where = "COLLECTION_STATUS = 'A'"
        End Select
        grdClickCellButton(grdSPTSFOC3, sql_where, False)

    End Sub

    Private Sub grdSPTSFOC3_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdSPTSFOC3.Error
        grdSPTSFOC3.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Private Sub grdSPTSFOCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSPTSFOCX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("EVENT_GROUP_NO").Text = e.Row.Cells("EVENT_GROUP_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Refresh_Documents()
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        EnforceConstraints(False)
        If optShow.Value = "E" Then
            Dim YP As String = cbeYP.Value
            Fill_Records("SPTSFOCX", YP)
            grdSPTSFOCX.Text = "Entered in " & cbeYP.Text
        ElseIf optShow.Value = "All" Then
            ASCMAIN1.sql = sqlSPTSFOCX
            Fill_Records("SPTSFOCX", "", True, ASCMAIN1.sql)
            grdSPTSFOCX.Text = "All"
        End If
        EnforceConstraints(True)

        Sort_grdColumns(grdSPTSFOCX, "EVENT_GROUP_NO".ToLower)
    End Sub

    Private Sub grdSPTSFOCX_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTSFOCX.InitializeRow
        With e.Row.Cells("APPR_STATUS_CODE")
            Select Case .Value & ""
                Case "A"
                    .Appearance.ForeColor = System.Drawing.Color.Green
                Case "P"
                    .Appearance.ForeColor = System.Drawing.Color.Purple
                Case "G"
                    .Appearance.ForeColor = System.Drawing.Color.Blue
                Case "X"
                    .Appearance.ForeColor = System.Drawing.Color.Red
            End Select

        End With
    End Sub

    Sub Clone_Record()
        BeginTrans()

        EVENT_GROUP_NO_new = ASCMAIN1.Next_Control_No("SPTSFOC1.EVENT_GROUP_NO")
        EnforceConstraints(False)

        For Each TABLE_NAME As String In New String() {"SPTSFOC1", "SPTSFOC3", "SPTSFOC9"}
            dst.Tables(TABLE_NAME).AcceptChanges()

            For Each row As DataRow In dst.Tables(TABLE_NAME).Select("")

                row.Item("EVENT_GROUP_NO") = EVENT_GROUP_NO_new

                Select Case TABLE_NAME
                    Case "SPTSFOC1"
                        row.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        row.Item("INIT_DATE") = DATETIME_STAMP
                        row.Item("LAST_OPER") = DBNull.Value
                        row.Item("LAST_DATE") = DBNull.Value
                        row.Item("APPR_STATUS_CODE") = "P"
                        row.Item("STATUS_CODE") = "O"
                        row.Item("AUTH_DATE") = DATETIME_STAMP.Date
                        row.Item("STATUS_CODE") = "O"
                        row.Item("AUTH_APPR_DATE") = DBNull.Value
                        row.Item("AUTH_APPR_BY") = DBNull.Value
                        row.Item("AUTH_APPR_AMT") = DBNull.Value
                        row.Item("AUTH_APPR_NOTES") = DBNull.Value

                End Select

                row.AcceptChanges()
                row.SetAdded()
            Next
            Update_Record_TDA(TABLE_NAME)
            ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "CLONE", "Contract Cloned from " & EVENT_GROUP_NO, "")

        Next
        EnforceConstraints(True)
        CommitTrans("Clone Successful - New Auth No: " & EVENT_GROUP_NO_new)
    End Sub

    Private Sub optShow_ValueChanged(sender As Object, e As EventArgs) Handles optShow.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_cbeYP()
    End Sub

    Sub Setup_cbeYP()
        cbeYP.Visible = (optShow.Value = "E")
        Refresh_Documents()
    End Sub

    Sub Setup_Retail_Weeks()
        If Absx1.txtFor("OPS_YYYYWW").Text = "" Then
            Absx1.txtFor("OPS_YYYYWW").Text = Set_OPS_YYYYWW()
        End If
    End Sub

    Function Set_OPS_YYYYWW() As String
        If Absx1.dteFor("AUTH_DATE").Value & "" = "" Then
            Return ""
        Else
            ASCMAIN1.sql = "Select Min (YYYYWW) from GLTPARM3 where WEEK_END_DATE >= '" & Format(Absx1.dteFor("AUTH_DATE").Value, "dd-MMM-yyyy") & "'"
            Dim OPS_YYYYWW As String = ASCDATA1.GetDataValue

            Return OPS_YYYYWW
        End If
    End Function

    Private Sub cmdBrowse_Click(sender As Object, e As EventArgs) Handles cmdBrowse.Click
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


    Sub Approve_Record(EVENT_GROUP_NO As String, APPROVAL_STATUS As String, APPROVAL_STATUS_orig As String, TOTAL_AMT As Decimal)

        ASCMAIN1.sql = $"Select AUTH_NO from SPTCOOP1 where EVENT_GROUP_NO = '{EVENT_GROUP_NO}'"
        Dim rowSPTCOOP1() As DataRow = ASCDATA1.GetDataTable().Select("")
        Dim AUTH_NO As String = rowSPTCOOP1(0).Item("AUTH_NO") & ""

        Select Case APPROVAL_STATUS
            Case "A"
                ASCMAIN1.sql = $"Update SPTSFOC1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = '{APPROVAL_STATUS}' where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})

                ASCMAIN1.sql = $"Update SPTCOOP1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = '{APPROVAL_STATUS}' where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})

                ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPR", "Approved" & AUTH_APPR_NOTES, "")
                ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPR", "Approved for " & Format(TOTAL_AMT, "$#,##0.00") & "; " & AUTH_APPR_NOTES, "")

            Case Else

                ASCMAIN1.sql = $"Update SPTSFOC1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = '{APPROVAL_STATUS}' where EVENT_GROUP_NO = '" & EVENT_GROUP_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})

                ASCMAIN1.sql = $"Update SPTCOOP1 Set AUTH_APPR_DATE = :PARM1, AUTH_APPR_BY = :PARM2, AUTH_APPR_AMT = :PARM3, AUTH_APPR_NOTES = :PARM4, APPR_STATUS_CODE = '{APPROVAL_STATUS}' where AUTH_NO = '" & AUTH_NO & "'"
                ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "DVNV", New Object() {DATETIME_STAMP, ASCMAIN1.USER_ID, TOTAL_AMT, AUTH_APPR_NOTES})

                ASCMAIN1.Record_Event("SPTSFOC1", EVENT_GROUP_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPR", $"Approval Status Changed from {APPROVAL_STATUS_orig} to {APPROVAL_STATUS} for " & Format(TOTAL_AMT, "$#,##0.00") & "; " & AUTH_APPR_NOTES & AUTH_APPR_NOTES, "")
                ASCMAIN1.Record_Event("SPTCOOP1", AUTH_NO, "", DATETIME_STAMP, ASCMAIN1.USER_ID, "APPR", $"Approval Status Changed from {APPROVAL_STATUS_orig} to {APPROVAL_STATUS} for " & Format(TOTAL_AMT, "$#,##0.00") & "; " & AUTH_APPR_NOTES, "")

        End Select

    End Sub

End Class