Imports System.Drawing
Imports System.Math

Public Class ICFPHYC1
    Dim rowICTPHYC1 As DataRow
    Dim location_support As Boolean = False
    Dim rowICTWHSE1 As DataRow
    Dim WHSE_CODE As String
    Dim TICKET_NO As String

    ' NOTE THAT IF WE DO NOT INITIALIZE COUNTS TABLES AT MONTH END, THAT THIS SCREEN WILL SHOW COUNTS (WHICH IS USEFUL) AFTER THE PI HAS BEEN POSTED
    '  HOWEVER, THE VARIANCE WILL WORK ONLY FOR LCOATABLE WHSES SINCE WE COMPARE TO WHTLOCB0 (SNAPSHOT BY LOCATION), 
    '  AND WE DIDN'T EVEN THINK TO INITIALIZE THAT TABLE.  THE BOOK INVENTORY WILL SHOW WITH BAD DATA FOR NON-LOCATABLE WHSES, SINCE IT IS LOOKING AT ICTSTAT1.
    '  BUT THIS MIGHT BE EASILY FIXED BY USING THE YP OF THE LAST PI UPDATE

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If MENU_ITEM_OBJECT = "ICFPHYCI" Then
            InquiryMode = True
        End If

        Get_PARM("ICTPARM1")
        Get_PARM("WHTPARM1")

        With dst
            ASCMAIN1.sql = "Select ICTPHYC1.*, X.ITEM_CODE, X.TOTAL_COUNT" _
            & " from ICTPHYC1, (Select WHSE_CODE, TICKET_NO, Min (ITEM_CODE) ITEM_CODE" _
            & ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) TOTAL_COUNT" _
            & " from ICTPHYC2 where WHSE_CODE = :PARM1 group by WHSE_CODE, TICKET_NO) X" _
            & " where ICTPHYC1.WHSE_CODE = :PARM1" _
            & "   and X.WHSE_CODE (+) = ICTPHYC1.WHSE_CODE" _
            & "   and X.TICKET_NO (+) = ICTPHYC1.TICKET_NO"
            Create_TDA(.Tables.Add, "ICTPHYCX", "**", 0, False, "V")
            .Tables("ICTPHYCX").Columns("TOTAL_COUNT").DataType = GetType(System.Int64)

            ASCMAIN1.sql = "Select X.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_COST_STD, X.BOOK, X.PHYS from ICTITEM1, (" & vbCrLf _
                & "Select ITEM_CODE, Sum (BOOK) BOOK, Sum (PHYS) PHYS from (" & vbCrLf _
                & "Select ITEM_CODE, 0 BOOK, Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS" & vbCrLf _
                & " from ICTPHYC2 where WHSE_CODE = :PARM1 group by ITEM_CODE" _
                & " union " & vbCrLf _
                & "Select ITEM_CODE, Sum (LOCATION_QTY) BOOK, 0 PHYS" & vbCrLf _
                & " from WHTLOCB0 where WHSE_CODE = :PARM2 group by ITEM_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select ITEM_CODE, WHSE_QTY_BEG BOOK, 0 PHYS" & vbCrLf _
                & " from ICTSTAT1 where WHSE_CODE = :PARM3 and OPS_YYYYPP = :PARM4" & vbCrLf _
                & " union " & vbCrLf _
                & "Select ITEM_CODE, -1 * WHSE_QTY_PHY BOOK, NULL PHYS" & vbCrLf _
                & " from ICTSTAT1 where WHSE_CODE = :PARM5 and OPS_YYYYPP = :PARM6 and WHSE_QTY_PHY <> 0" & vbCrLf _
                & ") group by ITEM_CODE" & vbCrLf _
                & ") X where X.ITEM_CODE = ICTITEM1.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTPHYCV", "**", 0, False, "VVVVVV")
            .Tables("ICTPHYCV").Columns("BOOK").DataType = GetType(System.Int64)
            .Tables("ICTPHYCV").Columns("PHYS").DataType = GetType(System.Int64)
            .Tables("ICTPHYCV").Columns.Add("VARIANCE", GetType(System.Int64), "ISNULL(PHYS,0) - ISNULL(BOOK,0)")
            .Tables("ICTPHYCV").Columns.Add("VARIANCE_COST", GetType(System.Int64), "ISNULL(ITEM_COST_STD,0) * (ISNULL(PHYS,0) - ISNULL(BOOK,0))")
            Create_TDA(.Tables.Add, "ICTWHSE1", "*")

            Create_TDA(.Tables.Add, "ICTPHYC1", "*")

            ASCMAIN1.sql = "Select ICTPHYC2.*, ICTITEM1.ITEM_DESC" _
                & " from ICTPHYC2,ICTITEM1 where ICTITEM1.ITEM_CODE = ICTPHYC2.ITEM_CODE"
            Create_TDA(.Tables.Add, "ICTPHYC2", "**", 2)
            .Tables("ICTPHYC2").Columns.Add("TOTAL_COUNT", GetType(System.Int64), "ISNULL(COUNT_CTNS,0) * ISNULL(CARTON_PACK_QTY,0) + ISNULL(COUNT_LOOSE,0)")

            ASCMAIN1.sql = "Select ICTPHYC2.*" _
                & ", ICTPHYC1.LOCATION_CODE, ICTPHYC1.COUNT_BY, ICTPHYC1.INIT_OPER, ICTPHYC1.INIT_DATE" _
                & " from ICTPHYC2,ICTPHYC1" _
                & " where ICTPHYC2.WHSE_CODE = :PARM1 and ICTPHYC2.ITEM_CODE = :PARM1" _
                & "   and ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE" _
                & "   and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO"
            Create_TDA(.Tables.Add, "ICTPHYCI", "**", 0, False, "VV", 3)
            .Tables("ICTPHYCI").Columns.Add("TOTAL_COUNT", GetType(System.Int64), "ISNULL(COUNT_CTNS,0) * ISNULL(CARTON_PACK_QTY,0) + ISNULL(COUNT_LOOSE,0)")

            ASCMAIN1.sql = "Select WHTLOCB0.*" _
                & " from WHTLOCB0 where WHSE_CODE = :PARM1 and ITEM_CODE = :PARM2"
            Create_TDA(.Tables.Add, "WHTLOCB0", "**", 0, False, "VV", 4)

            ASCMAIN1.sql = "Select * from WHTLOCM1"
            Create_TDA(.Tables.Add, "WHTLOCM1", "**", 0, False)

            'ASCMAIN1.sql = "Select * from ICTCLAS1"
            'Create_TDA(.Tables.Add, "ICTCLAS1", "**", 0, False)
        End With

        Fill_Records("WHTLOCM1")
        '  Fill_Records("ICTCLAS1")

        ' cbe.DataSource = ASCDATA1.GetDataTable("Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grdWHTLOCB0.DataSource = dst.Tables("WHTLOCB0")
        grdICTPHYC2.DataSource = dst.Tables("ICTPHYC2")
        grdICTPHYCI.DataSource = dst.Tables("ICTPHYCI")
        grdICTPHYCX.DataSource = dst.Tables("ICTPHYCX")
        grdICTPHYCV.DataSource = dst.Tables("ICTPHYCV")

        Create_Summary(grdICTPHYCX, "TICKET_NO", "Count")
        Create_Summary(grdICTPHYCX, "TOTAL_COUNT")

        Create_Summary(grdICTPHYCV, "ITEM_CODE", "Count")
        Create_Summary(grdICTPHYCV, New String() {"BOOK", "PHYS", "VARIANCE", "VARIANCE_COST"})

        Create_Summary(grdICTPHYC2, "TICKET_LNO", "Count")
        Create_Summary(grdICTPHYC2, "COUNT_CTNS")
        Create_Summary(grdICTPHYC2, "TOTAL_COUNT")

        Create_Summary(grdICTPHYCI, "TICKET_NO", "Count")
        Create_Summary(grdICTPHYCI, "COUNT_CTNS")
        Create_Summary(grdICTPHYCI, "TOTAL_COUNT")

        Create_Summary(grdWHTLOCB0, "LOCATION_CODE", "Count")
        Create_Summary(grdWHTLOCB0, "LOCATION_QTY")



        With grdICTPHYC2.DisplayLayout.Bands("ICTPHYC2")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns

                If gcol.Key = "ITEM_CODE" Or gcol.Key = "COUNT_CTNS" Or gcol.Key = "COUNT_LOOSE" Or gcol.Key = "CARTON_PACK_QTY" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Color.Beige
                End If

                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "COUNT_CTNS" Or gcol.Key = "COUNT_CTNS" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                    'ElseIf gcol.Key = "PHYS" Or gcol.Key = "BOOK" Then
                    '    gcol.Header.Appearance.BackColor = Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If

            Next
            .Columns("TICKET_LNO").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
        End With

        With grdICTPHYCV.DisplayLayout.Bands("ICTPHYCV")
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If gcol.Key = "VARIANCE" Or gcol.Key = "VARIANCE_COST" Then
                    gcol.Header.Appearance.BackColor2 = Color.Orange
                ElseIf gcol.Key = "PHYS" Or gcol.Key = "BOOK" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightBlue
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
            .Columns("ITEM_CODE").Header.Fixed = True
        End With

        With grdICTPHYCX.DisplayLayout.Bands("ICTPHYCX")
            .Columns("TICKET_NO").Header.Fixed = True
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "TICKET_NO" Or gcol.Key = "COUNTED_BY" Or gcol.Key = "LOCATION_CODE" Then
                    gcol.Header.Appearance.BackColor2 = Color.Pink
                ElseIf gcol.Key = "ITEM_CODE" Or gcol.Key = "TOTAL_COUNT" Then
                    gcol.Header.Appearance.BackColor2 = Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Color.LightGray
                End If
            Next
        End With


        With grdICTPHYCI.DisplayLayout.Bands("ICTPHYCI")

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.Turquoise
            Next
        End With


        With grdWHTLOCB0.DisplayLayout.Bands("WHTLOCB0")

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                gcol.Header.Appearance.BackColor2 = Color.Yellow
            Next
        End With

        'With grdICTPHYCX.DisplayLayout.Bands("ICTPHYCX")
        '    .Columns("TICKET_NO").Header.Fixed = True
        '    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
        '        gcol.Header.Appearance.BackColor = Color.White
        '        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
        '        If gcol.Key = "TICKET_NO" Or gcol.Key = "COUNTED_BY" Or gcol.Key = "LOCATION_CODE" Then
        '            gcol.Header.Appearance.BackColor = Color.Pink
        '        ElseIf gcol.Key = "ITEM_CODE" Or gcol.Key = "TOTAL_COUNT" Then
        '            gcol.Header.Appearance.BackColor = Color.LightGreen
        '        Else
        '            gcol.Header.Appearance.BackColor = Color.LightGray
        '        End If
        '    Next
        'End With


        'ASCMAIN1.Add_Value_List(grdICTPHYCX, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")
        'ASCMAIN1.Add_Value_List(grdICTPHYCV, "REASON_CODE", "Select REASON_CODE, REASON_DESC from ICTREAS1 order by REASON_DESC")

        grpHeader.Visible = False

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("WHSE_CODE")

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify a Valid Warehouse"
                Else
                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    ElseIf rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        'ElseIf rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                        '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Counts Entry Allowed"
                    ElseIf rowICTWHSE1.Item("WHSE_PHYS_STATUS") & "" <> "C" Then
                        EMsg &= vbCr & "Warehouse has not been Initialized for Physical Counts Entry"
                    End If
                End If

                If Absx1.txtFor("TICKET_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Ticket"
                Else
                    Dim rowICTPHYC1 As DataRow = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("TICKET_NO").Text})
                    If rowICTPHYC1 IsNot Nothing Then
                        Click_Command("View")
                        Exit Sub
                    Else
                        If Not ASCMAIN1.Logical_Lock("ICTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & Absx1.txtFor("TICKET_NO").Text) Then
                            Exit Sub
                        End If
                    End If
                End If



            Case "Edit"
                Dim rowICTPHYC1 As DataRow = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("TICKET_NO").Text})
                If rowICTPHYC1 Is Nothing Then
                    EMsg &= "Ticket is not on File"
                Else

                    Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text)
                    If IsNothing(rowICTWHSE1) Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Valid"
                    ElseIf rowICTWHSE1.Item("WHSE_STATUS").ToString <> "A" Then
                        EMsg &= vbCr & "Warehouse Entered Is Not Active"
                        'ElseIf rowICTWHSE1.Item("LP_CODE") & "" <> "" Then
                        '    EMsg &= vbCr & "Warehouse Entered Is A 3PL.  No Counts Entry Allowed"
                    ElseIf rowICTWHSE1.Item("WHSE_PHYS_STATUS") & "" <> "C" Then
                        EMsg &= vbCr & "Warehouse has not been Initialized for Physical Counts Entry"
                    End If

                    If Not ASCMAIN1.Logical_Lock("ICTPHYC1", Absx1.txtFor("WHSE_CODE").Text & ":" & Absx1.txtFor("TICKET_NO").Text) Then
                        Exit Sub
                    End If
                End If

            Case "View"
                If Absx1.txtFor("TICKET_NO").Text = "" Then
                    EMsg &= vbCr & "You must specify a Document No to View"
                Else
                    rowICTPHYC1 = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("TICKET_NO").Text})
                    If rowICTPHYC1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Document " & Absx1.txtFor("TICKET_NO").Text & " on File"
                    End If
                End If

            Case "Update"

                If location_support Then
                    If Absx1.txtFor("LOCATION_CODE").Text = "" Then
                        EMsg &= vbCr & "You Must Specify a Location"
                    Else
                        Dim rowWHTLOCM1 As DataRow = LookUp("WHTLOCM1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("LOCATION_CODE").Text})
                        If rowWHTLOCM1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Value Specified for Location"
                        End If
                    End If

                    If Absx1.txtFor("COUNT_BY").Text = "" Then
                        EMsg &= vbCr & "You Must enter either Notes or Initials of the person who did the count"
                    End If
                End If

                If grdICTPHYC2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Details Entered"
                Else
                    'For Each rowICTPHYC2 As DataRow In dst.Tables("ICTPHYC2").Select("", "", DataViewRowState.CurrentRows)
                    '    If rowICTPHYC2.Item("COST_CATGY_CODE") & "" = "" Then
                    '        EMsg &= vbCr & "Unable to determine Cost Category for " & rowICTPHYC2.Item("ITEM_CODE") & ""
                    '    End If
                    '    If rowICTPHYC2.Item("PROD_CODE") & "" = "" Then
                    '        EMsg &= vbCr & "Unable to determine Product Code for " & rowICTPHYC2.Item("ITEM_CODE") & ""
                    '    End If
                    'Next
                End If

            Case "Delete"
                If MessageBox.Show("Are you sure you want to Delete this Entry?", "Confirm Deletion", _
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then
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

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "View"
                EntryMode = "V"
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


            Case "By Ticket"
                Print_Counts("T")
            Case "By Location"
                Print_Counts("L")
            Case "By Item"
                Print_Counts("S")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode

                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If


                    If ScreenMode And (EntryMode <> "N" And EntryMode <> "E") Then
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

                    .Items("Delete").Visible = (ScreenMode And EntryMode = "E")
                End With

                '  .Groups("Variances").Visible = ScreenMode And (EntryMode = "V")
                .Groups("Count Reports").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = ScreenMode
        grpHeader.Visible = ScreenMode

        tab0.Visible = Not ScreenMode

        If ScreenMode Then
            Absx1.txtFor("LOCATION_CODE").Visible = location_support
            Absx1.txtFor("LOCATION_DESC").Visible = location_support
            lblLOCATION_CODE.Visible = location_support

            If InquiryMode Then
                With UltraExplorerBar1.Groups("Screen Control")
                    .Items("New").Visible = False
                    .Items("Update").Visible = False
                    .Items("Cancel").Visible = False
                End With
            End If

            Set_Read_Only(grpHeader, (EntryMode = "V"))
            ' Set_Read_Only(SplitContainer2, (EntryMode = "V"))
            If EntryMode = "N" Or EntryMode = "E" Then
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPHYC2}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End With
                Next
                With grdICTPHYC2.DisplayLayout.Bands(0)
                    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.LightYellow
                    .Columns("COUNT_CTNS").CellAppearance.BackColor = Color.LightYellow
                    .Columns("COUNT_LOOSE").CellAppearance.BackColor = Color.LightYellow
                End With

            Else
                For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdICTPHYC2, grdICTPHYCI}
                    With grd.DisplayLayout.Override
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    End With
                Next
                'With grdICTPHYC2.DisplayLayout.Bands(0)
                '    .Columns("ITEM_CODE").CellAppearance.BackColor = Color.Empty
                '    .Columns("COUNT_CTNS").CellAppearance.BackColor = Color.Empty
                'End With
            End If

            If grdICTPHYC2.ActiveRow Is Nothing Then
                Setup_ICTPHYC2("")
            End If
            Setup_WHTLOCB0(False)

            Absx1.txtFor("LOCATION_CODE").Focus()
            Setup_tab0()
        Else
            Clear_Record()
            tab0.SelectedTab = tab0.Tabs("Tickets")
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTPHYC1", "ICTPHYC2", "ICTPHYCI"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Absx1.txtFor("WHSE_CODE").Text = WHSE_CODE
        Absx1.txtFor("TICKET_NO").Text = ""

        If WHSE_CODE = "" Then
            Absx1.txtFor("WHSE_CODE").Focus()
        Else
            Absx1.txtFor("TICKET_NO").Focus()
        End If

        Refresh_Documents()
        Setup_tab0()
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        TICKET_NO = Absx1.txtFor("TICKET_NO").Text

        If EntryMode = "N" Then
            rowICTPHYC1 = dst.Tables("ICTPHYC1").NewRow
            rowICTPHYC1.Item("WHSE_CODE") = WHSE_CODE
            rowICTPHYC1.Item("TICKET_NO") = TICKET_NO ' ASCMAIN1.Next_Control_No("ICTPHYC1.TICKET_NO")

            rowICTPHYC1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            rowICTPHYC1.Item("INIT_DATE") = DATETIME_STAMP
            rowICTPHYC1.Item("LAST_OPER") = ASCMAIN1.USER_ID
            rowICTPHYC1.Item("LAST_DATE") = DATETIME_STAMP
            dst.Tables("ICTPHYC1").Rows.Add(rowICTPHYC1)
        Else
            Fill_Record("ICTPHYC1", New String() {WHSE_CODE, TICKET_NO})
            dst.AcceptChanges()
        End If

        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        location_support = (rowICTWHSE1.Item("WHSE_LOCATOR") & "" = "1")
        With grdICTPHYC2.DisplayLayout.Bands(0)
            .Columns("BAR_CODE").Hidden = Not location_support
        End With

        Fill_Records("ICTPHYC2", New String() {WHSE_CODE, TICKET_NO})

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        Update_Record_TDA("ICTPHYC1")
        Update_Record_TDA("ICTPHYC2")
        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()

        Delete_Records("ICTPHYC1")
        Delete_Records("ICTPHYC2")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where WHSE_CODE = '" & WHSE_CODE & "' and TICKET_NO = '" & TICKET_NO & "'")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdICTPHYCX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdICTPHYC2, "B", "Item Status Inquiry")
        Load_Popup_Menu(grdICTPHYCV, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
        Load_Popup_Menu(grdWHTLOCB0, "BB", "Location Inquiry", "Show 0's")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                Case "grdWHTLOCB0"
                    '  tlb_sbt = DirectCast(tlb.Tools("Show 0s"), UltraWinToolbars.StateButtonTool)

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Acknowledge w/Notes"
            '    Log_SetMode(True, True)
            Case "Show 0's"
                '  tlb_sbt = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Setup_WHTLOCB0(True)

            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                If rowICTITEM1 IsNot Nothing Then
                    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                End If

            Case "Location Inquiry"
                'Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                'Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                'If rowICTITEM1 IsNot Nothing Then
                '    Context_Launch("Select", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                'End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "WHSE_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Refresh_Documents()
                    'If Not InquiryMode Then
                    '    Click_Command("New", e)
                    'End If
                End If
            Case "TICKET_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("WHSE_CODE").Text <> "" Then
                        Dim row As DataRow = LookUp("ICTPHYC1", New String() {Absx1.txtFor("WHSE_CODE").Text, Absx1.txtFor("WHSE_CODE").Text})
                        If row IsNot Nothing Then
                            Click_Command("View", e)
                        Else
                            Click_Command("New", e)
                        End If
                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "WHSE_CODE"
                Refresh_Documents()
                'If Not InquiryMode Then
                '    Click_Command("New")
                'End If
            Case "TICKET_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "WHSE_CODE"
                Refresh_Documents()
            Case "COUNT_BY"
                grdICTPHYC2.Focus()
                If grdICTPHYC2.ActiveRow Is Nothing Then
                    If grdICTPHYC2.Rows.Count = 0 Then
                        grdICTPHYC2.DisplayLayout.Bands(0).AddNew()
                    End If
                End If
                If grdICTPHYC2.ActiveRow IsNot Nothing Then
                    grdICTPHYC2.ActiveCell = grdICTPHYC2.ActiveRow.Cells("ITEM_CODE")
                End If

        End Select
    End Sub

    Public Overrides Sub num_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.num_ValueChanged(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "INV_FREIGHT"
        End Select
    End Sub

#End Region

#Region "grdICTPHYC2"

    Private Sub grdICTPHYC2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPHYC2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"

                grdCodeDesc(grdICTPHYC2, "ICTITEM1", "ITEM_CODE", "ITEM_DESC")
                If cdr IsNot Nothing Then
                    Dim ITEM_CODE As String = e.Cell.Value
                    'Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
                    e.Cell.Row.Cells("CARTON_PACK_QTY").Value = cdr.Item("CARTON_PACK_QTY")

                    'Dim COST_CATGY_CODE As String = cdr.Item("COST_CATGY_CODE") & ""
                    'Dim PROD_CODE As String = cdr.Item("PROD_CODE") & ""
                    'Dim ITEM_COST_STD As Decimal = Val(cdr.Item("ITEM_COST_STD") & "")
                    'e.Cell.Row.Cells("COST_CATGY_CODE").Value = COST_CATGY_CODE
                    'e.Cell.Row.Cells("PROD_CODE").Value = PROD_CODE
                    'e.Cell.Row.Cells("ITEM_COST_STD").Value = ITEM_COST_STD

                    Setup_ICTPHYC2(ITEM_CODE)

                Else
                    grdICTPHYC2.PerformAction(UltraWinGrid.UltraGridAction.PrevCellByTab)
                End If

            Case "COUNT_CTNS"

        End Select
    End Sub

    Private Sub grdICTPHYC2_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPHYC2.AfterExitEditMode
        'Select Case grdICTPHYC2.ActiveCell.Column.Key

        'End Select
    End Sub

    Private Sub grdICTPHYC2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPHYC2.AfterRowActivate
        With grdICTPHYC2.DisplayLayout.Bands(0)
            If grdICTPHYC2.ActiveRow.IsAddRow Then
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdICTPHYC2.ActiveCell = grdICTPHYC2.ActiveRow.Cells("ITEM_CODE")
                grdICTPHYC2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        Dim ITEM_CODE As String = grdICTPHYC2.ActiveRow.Cells("ITEM_CODE").Value & ""
        Setup_ICTPHYC2(ITEM_CODE)
        'If EntryMode = "V" Then
        '    Show_Variances()
        'End If
    End Sub

    Private Sub grdICTPHYC2_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdICTPHYC2.AfterRowsDeleted
        DisplayTotals()
    End Sub

    Private Sub grdICTPHYC2_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdICTPHYC2.AfterRowUpdate
        DisplayTotals()
    End Sub

    Private Sub grdICTPHYC2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdICTPHYC2.BeforeExitEditMode
        If grdICTPHYC2.ActiveCell Is Nothing Then Exit Sub
        With grdICTPHYC2.ActiveCell
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
                            ASCMAIN1.Progress("Invalid Style Code (" & .Text & ")")
                            If .Value IsNot Nothing Then
                                .Value = ""
                            End If
                            e.Cancel = True
                        End If
                    End If
                    'Case "BAR_CODE"
                    '    If location_support Then
                    '        If .Text <> "" Then
                    '            If .Value IsNot Nothing Then
                    '                .Value = .Text.ToUpper
                    '            End If

                    '        End If
                    '        If .Text <> "" Then
                    '            cdr = LookUp("WHTBARC1", .Text)
                    '            If cdr Is Nothing Then
                    '                ASCMAIN1.Progress("Invalid Bar Code (" & .Text & ")")
                    '                If .Value IsNot Nothing Then
                    '                    .Value = ""
                    '                End If
                    '                e.Cancel = True
                    '            End If
                    '        End If
                    '    End If
            End Select
        End With
    End Sub

    Private Sub grdICTPHYC2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdICTPHYC2.BeforeRowUpdate
        With grdICTPHYC2
            If e.Row.Cells("ITEM_CODE").Text = "" Then
                '                MsgBox("Missing Value for Item Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            Else
                LookUp("ICTITEM1", e.Row.Cells("ITEM_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Item Code (" & e.Row.Cells("ITEM_CODE").Text & ")", _
                           MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If location_support Then
                'If e.Row.Cells("BAR_CODE").Text = "" Then
                '    e.Cancel = True
                'Else
                '    LookUp("WHTBARC1", e.Row.Cells("BAR_CODE").Text)
                '    If cdr Is Nothing Then
                '        MsgBox("Invalid Value entered for Bar Code (" & e.Row.Cells("BAR_CODE").Text & ")", _
                '               MsgBoxStyle.OkOnly, "Cannot Update Row")
                '        e.Cancel = True
                '    End If
                'End If

            End If

            If Val(e.Row.Cells("COUNT_CTNS").Value & "") = 0 And Val(e.Row.Cells("COUNT_LOOSE").Value & "") = 0 Then
                'MsgBox("Invalid Value entered for Count", MsgBoxStyle.OkOnly, "Cannot Update Row")
                e.Cancel = True
            End If

            If e.Cancel Then
                e.Row.CancelUpdate()
            End If

            If Not e.Cancel Then
                If e.Row.Cells("TICKET_NO").Text = "" Then
                    .ActiveRow.Cells("WHSE_CODE").Value = WHSE_CODE
                    .ActiveRow.Cells("TICKET_NO").Value = Absx1.CtlFor("TICKET_NO").Text
                    .ActiveRow.Cells("TICKET_LNO").Value = Val(dst.Tables("ICTPHYC2").Compute("Max(TICKET_LNO)", "") & "") + 1
                End If
            End If
        End With
    End Sub

    Private Sub grdICTPHYC2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdICTPHYC2.ClickCellButton

        If grdICTPHYC2.ActiveRow Is Nothing Then Exit Sub

        Dim sql_where As String = ""
        Select Case e.Cell.Column.Key
            Case "ITEM_CODE"
                'Case "LOCATION_CODE"
                '    sql_where = "WHSE_CODE = '" & Absx1.txtFor("WHSE_CODE").Text & "'"
        End Select
        grdClickCellButton(grdICTPHYC2, sql_where, False)

    End Sub

    Private Sub grdICTPHYC2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdICTPHYC2.Error
        grdICTPHYC2.ActiveRow.CancelUpdate()
    End Sub

#End Region

    Sub DisplayTotals()
        'Dim TOTAL_COSTS As Decimal = Val(dst.Tables("ICTPHYC2").Compute("SUM(LINE_COSTS)", "") & "")
        'Absx1.numFor("TOTAL_COSTS").Value = TOTAL_COSTS
    End Sub

    Private Sub grdICTPHYCX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPHYCX.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Text
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTPHYCV_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdICTPHYCV.AfterRowActivate
        Dim ITEM_CODE As String = ""
        If grdICTPHYCV.ActiveRow IsNot Nothing AndAlso grdICTPHYCV.ActiveRow.IsDataRow Then
            ITEM_CODE = grdICTPHYCV.ActiveRow.Cells("ITEM_CODE").Value
        End If
        Setup_ICTPHYC2(ITEM_CODE)
    End Sub

    Private Sub grdICTPHYCV_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPHYCV.DoubleClickRow
        'If e.Row.IsDataRow Then
        '    Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Text
        '    Click_Command("View")
        'End If
    End Sub

    Private Sub optVariances_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optVariances.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Variances()
    End Sub

    Sub Show_Variances()
        Dim dvw As DataView = DirectCast(grdICTPHYCV.DataSource, DataTable).DefaultView
        If optVariances.Value = "A" Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "VARIANCE <> 0"
        End If
    End Sub

    Sub Refresh_Documents()
        Me.Cursor = Cursors.WaitCursor
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        Fill_Records("ICTPHYCX", WHSE_CODE)
        grdICTPHYCX.Text = "Physical Counts for Warehouse " & WHSE_CODE
        tab0.SelectedTab = tab0.Tabs("Tickets")
    End Sub

    Sub Setup_ICTPHYC2(ITEM_CODE As String)
        If ITEM_CODE = "" Then
            splItemDetails.Visible = False
        Else
            splItemDetails.Visible = True
            Fill_Records("ICTPHYCI", New String() {WHSE_CODE, ITEM_CODE})
            grdICTPHYCI.Text = "Tickets with Item " & ITEM_CODE
            Fill_Records("WHTLOCB0", New String() {WHSE_CODE, ITEM_CODE})
            grdWHTLOCB0.Text = "Book Inventory by Location for Item " & ITEM_CODE
            Setup_WHTLOCB0(False)
        End If

    End Sub

    Sub Setup_WHTLOCB0(Show_0s As Boolean)
        Dim dvw As DataView = DirectCast(grdWHTLOCB0.DataSource, DataTable).DefaultView
        If Show_0s Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "LOCATION_QTY <> 0"
        End If
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As System.Object, e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()

    End Sub

    Sub Setup_tab0()
        If tab0.SelectedTab.Key = "Variances" Then
            If Load_Variances() Then
                splItemDetails.Parent = splICTPHYCV.Panel2
                Show_Variances()
            Else
                tab0.SelectedTab = tab0.Tabs("Tickets")
            End If
        Else
            splItemDetails.Parent = splICTPHYC2.Panel2
        End If
        UltraExplorerBar1.Groups("Variances").Visible = (tab0.SelectedTab.Key = "Variances") And Not ScreenMode
    End Sub

    Private Sub btnRefresh_Click(sender As System.Object, e As System.EventArgs) Handles btnRefresh.Click
        Load_Variances()
        Show_Variances()
    End Sub

    Function Load_Variances() As Boolean

        WHSE_CODE = Absx1.txtFor("WHSE_CODE").Text
        Dim rowICTWHSE1 As DataRow = LookUp("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then
            MsgBox("You Must Pick a Valid Warehouse Code", MsgBoxStyle.OkOnly, "Cannot Show Variance")
            Return False
        End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Compiling Variances")

        If rowICTWHSE1.Item("WHSE_LOCATOR") & "" <> "1" Then
            Fill_Records("ICTPHYCV", New String() {WHSE_CODE, "", WHSE_CODE, ASCMAIN1.CYP, WHSE_CODE, ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)})
        Else
            Fill_Records("ICTPHYCV", New String() {WHSE_CODE, WHSE_CODE, "", "", "", ""})
        End If
        Sort_grdColumns(grdICTPHYCV, "ITEM_CODE")

        If grdICTPHYCV.ActiveRow Is Nothing Then Setup_ICTPHYC2("")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Return True

    End Function

    Private Sub grdICTPHYCI_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdICTPHYCI.DoubleClickRow
        If Not ScreenMode Then
            '  Absx1.txtFor("WHSE_CODE").Text = ""
            Absx1.txtFor("TICKET_NO").Text = e.Row.Cells("TICKET_NO").Value & ""
            Click_Command("View")
        End If
    End Sub

    Private Sub grdICTPHYCI_InitializeLayout(sender As System.Object, e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdICTPHYCI.InitializeLayout

    End Sub

    Private Sub grdICTPHYC2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdICTPHYC2.InitializeRow
        Dim COUNT_LOOSE As Int64 = Val(e.Row.Cells("COUNT_LOOSE").Value & "")
        Dim CARTON_PACK_QTY As Int64 = Val(e.Row.Cells("CARTON_PACK_QTY").Value & "")
        If COUNT_LOOSE >= CARTON_PACK_QTY And CARTON_PACK_QTY <> 0 And CARTON_PACK_QTY <> 1 Then
            e.Row.Cells("COUNT_LOOSE").Appearance.ForeColor = Color.Red
            e.Row.Cells("COUNT_LOOSE").ToolTipText = "Loose Count is greater than or equal to Carton Pack Qty"
        End If
    End Sub

    Sub Print_Counts(BY As String)

        Dim RPT As String = ""
        Dim RPT_TITLE As String = ""
        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text
        If WHSE_CODE = "" Then Exit Sub
        Dim rowICTWHSE1 As DataRow = Fill_Record("ICTWHSE1", WHSE_CODE)
        If rowICTWHSE1 Is Nothing Then Exit Sub

        ASCMAIN1.sql = "Select * from ICTPHYC1 where WHSE_CODE = '" & WHSE_CODE & "'"
        Fill_Records("ICTPHYC1", "", True, ASCMAIN1.sql)


        ASCMAIN1.sql = "Select ICTPHYC2.*, ICTITEM1.ITEM_DESC" _
            & " from ICTPHYC2,ICTITEM1 where ICTITEM1.ITEM_CODE = ICTPHYC2.ITEM_CODE" _
            & " and ICTPHYC2.WHSE_CODE = '" & WHSE_CODE & "'"
        Fill_Records("ICTPHYC2", "", True, ASCMAIN1.sql)

        Select Case BY
            Case "T"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Ticket"
            Case "L"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Location"
            Case "S"
                RPT = "ICRPHYC1"
                RPT_TITLE = "Physical Counts by Item"
        End Select

        'Synch_TABLE_NAME("ICTSTYL1")
        Print_Report_Begin()
        CR_params.Add("SORT_BY", BY)
        Generate_Report(RPT, RPT_TITLE, "")
        Print_Report_End()

    End Sub
End Class