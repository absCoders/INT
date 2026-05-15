Public Class SOFRSRV1
    'Load_Events

#Region "Declarations"
    Dim CUST_CODE As String
    Dim CUST_NAME As String         ' Sold-To Customer Name

    Dim RSRV_NO As String
    Dim ORDR_CUST_PO As String      ' Customer's PO No
    Dim SREP_CODE As String         ' Orders Sales Rep Code
    Dim SREP2_CODE As String        ' Orders Sales Rep2 Code

    Dim rowSOTRSRV1 As DataRow
    Dim rowARTCUST1 As DataRow      ' ARTCUST1 for the Sold-To
    Dim rowICTSTYL1 As DataRow
    Dim RSRV_LNOs As New List(Of Int64) ' list of RSRV_LNOs that are deleted

    Dim COLOR_CODEs As New List(Of String)    ' table of COLOR_CODEs associated with a STYLE_CODE
#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")

        With dst
            ASCMAIN1.sql = "Select SOTRSRV1.* from SOTRSRV1 where RSRV_NO = :PARM1"
            Create_TDA(.Tables.Add, "SOTRSRVX", "**", 0, False, "V", 1)

            Create_TDA(.Tables.Add, "SOTRSRV1", "*", 1)

            ASCMAIN1.sql = "Select SOTRSRV2.*, ICTSTYL1.STYLE_DESC, ICTCOLR1.COLOR_DESC" _
            & " from SOTRSRV2,ICTSTYL1,ICTCOLR1" _
            & " where ICTSTYL1.STYLE_CODE = SOTRSRV2.STYLE_CODE" _
            & "   and ICTCOLR1.COLOR_CODE = SOTRSRV2.COLOR_CODE"
            Create_TDA(.Tables.Add, "SOTRSRV2", "**", 1)

            .Tables("SOTRSRV2").Columns.Add("RSRV_AMT", GetType(System.Decimal), "ISNULL(RSRV_QTY,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            .Tables("SOTRSRV2").Columns.Add("RSRV_AMT_OPEN", GetType(System.Decimal), "ISNULL(RSRV_QTY_OPEN,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            .Tables("SOTRSRV2").Columns.Add("RSRV_AMT_ALLO", GetType(System.Decimal), "ISNULL(RSRV_QTY_ALLO,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            .Tables("SOTRSRV2").Columns.Add("RSRV_AMT_USED", GetType(System.Decimal), "ISNULL(RSRV_QTY_USED,0) * ISNULL(ORDR_UNIT_PRICE,0)")
            .Tables("SOTRSRV2").Columns.Add("RSRV_AMT_CANC", GetType(System.Decimal), "ISNULL(RSRV_QTY_CANC,0) * ISNULL(ORDR_UNIT_PRICE,0)")

            With .Tables.Add("SOTRSRVT")
                .Columns.Add("KEY", GetType(System.Int32))
                .Columns.Add("STATUS")
                .Columns.Add("QTY", GetType(System.Int32))
                .Columns.Add("AMT", GetType(System.Decimal))
                .PrimaryKey = New DataColumn() {.Columns("KEY")}
            End With

            Create_TDA(.Tables.Add, "ARTCUST1", "*", 1)

            ASCMAIN1.sql = "Select * from ICTCOLR1"
            Create_TDA(.Tables.Add, "ICTCOLR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ICTSTYC1.COLOR_CODE, ICTCOLR1.COLOR_DESC" _
                & " from ICTSTYC1,ICTCOLR1 where ICTSTYC1.STYLE_CODE = :PARM1" _
                & "  and ICTCOLR1.COLOR_CODE = ICTSTYC1.COLOR_CODE"
            Create_TDA(.Tables.Add, "ICTCOLRS", "**", 0, False, "V", 1)

            'ASCMAIN1.sql = "Select * from SOTRSRV7 where RSRV_GROUP_NO = :PARM1 " & vbCrLf _
            '    & "   and SOTRSRV7.STYLE_CODE = :PARM2 " & vbCrLf _
            '    & "   and SOTRSRV7.COLOR_CODE = :PARM3" & vbCrLf _
            '    & "   and SOTRSRV7.ALLO_BATCH_NO is Null"
            'Create_TDA(.Tables.Add, "SOTRSRV7", "**", 0, True, "VVV", 1)
        End With

        grdSOTRSRVX.DataSource = dst.Tables("SOTRSRVX")
        grdSOTRSRV2.DataSource = dst.Tables("SOTRSRV2")
        grdSOTRSRVT.DataSource = dst.Tables("SOTRSRVT")
 
        grdSOTRSRVX.DisplayLayout.UseFixedHeaders = True
        With grdSOTRSRVX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RSRV_NO", "CUST_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        grdSOTRSRV2.DisplayLayout.UseFixedHeaders = True
        With grdSOTRSRV2.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"RSRV_LNO", "STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With

        With grdSOTRSRV2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If New String() {"STYLE_CODE", "COLOR_CODE", "RSRV_QTY", "RSRV_QTY_OPEN", "ORDR_UNIT_PRICE", "RSRV_PRIORITY", "RSRV_PRIORITY_DATE", "RSRV_DEMAND_DATE"}.Contains(gcol.Key) Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = Drawing.Color.Beige
                End If
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                If New String() {"RSRV_QTY", "RSRV_QTY_OPEN", "RSRV_QTY_ALLO", "RSRV_QTY_USED", "RSRV_QTY_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"RSRV_AMT", "RSRV_AMT_OPEN", "RSRV_AMT_ALLO", "RSRV_AMT_USED", "RSRV_AMT_CANC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                ElseIf New String() {"STYLE_CODE", "STYLE_DESC", "COLOR_CODE", "COLOR_DESC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.Pink
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End If
            Next
        End With

        Create_Summary(grdSOTRSRVX, "RSRV_NO", "Count")

        Create_Summary(grdSOTRSRV2, "RSRV_LNO", "Count")
        Create_Summary(grdSOTRSRV2, New String() {"RSRV_QTY", "RSRV_QTY_OPEN", "RSRV_QTY_ALLO", "RSRV_QTY_USED", "RSRV_QTY_CANC", "RSRV_AMT"})

        With dst.Tables("SOTRSRVT").Rows
            .Add(New Object() {1, "Rsrv", 0, 0})
            .Add(New Object() {2, "Open", 0, 0})
            .Add(New Object() {3, "Allo", 0, 0})
            .Add(New Object() {4, "Used", 0, 0})
            .Add(New Object() {5, "Canc", 0, 0})
        End With
        Sort_grdColumns(grdSOTRSRVT, "KEY", True)

        Show_Filter(grdSOTRSRVX, True)
        grdSOTRSRVX.DisplayLayout.GroupByBox.Hidden = False

        SplitContainer1.Panel2Collapsed = True

        Check_InquiryMode()

    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFRSRVI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("CUST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify a Customer"
                Else
                    rowARTCUST1 = LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                    If rowARTCUST1 IsNot Nothing Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        ORDR_CUST_PO = Absx1.txtFor("ORDR_CUST_PO").Text
                        If ORDR_CUST_PO = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                            EMsg &= vbCr & "You Must Provide a Value for Customer PO"
                        End If
                    Else
                        EMsg &= vbCr & "No Record of Customer " & Absx1.txtFor("CUST_CODE").Text
                    End If
                End If

                If EMsg = "" Then
                    ' Customer must have a Sales Rep assigned
                    SREP_CODE = rowARTCUST1.Item("SREP_CODE") & ""
                    Dim rowSOTSREP1 As DataRow = Nothing
                    If SREP_CODE = "" Then
                        EMsg &= vbCr & "This Customer Has No Sales Rep Assigned"
                    Else
                        rowSOTSREP1 = LookUp("SOTSREP1", rowARTCUST1.Item("SREP_CODE") & "")
                        If rowSOTSREP1 Is Nothing Then
                            EMsg &= vbCr & "This Customer has an Invalid Sales Rep Assigned (" & SREP_CODE & ")"
                        End If
                    End If
                    SREP2_CODE = rowARTCUST1.Item("SREP2_CODE") & ""
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                End If

            Case "Edit", "View"

                CUST_CODE = ""
                RSRV_NO = ""

                If Absx1.txtFor("RSRV_NO").Text = "" Then
                    EMsg &= vbCr & "No Reservation No Specified"
                Else
                    RSRV_NO = Absx1.txtFor("RSRV_NO").Text
                    rowSOTRSRV1 = LookUp("SOTRSRV1", RSRV_NO)
                    If rowSOTRSRV1 Is Nothing Then
                        EMsg &= vbCr & "No Record of Reservation No " & RSRV_NO
                    Else
                        CUST_CODE = rowSOTRSRV1.Item("CUST_CODE")
                        If rowSOTRSRV1.Item("RSRV_STATUS") & "" <> "O" And eItemKey = "Edit" Then
                            Select Case rowSOTRSRV1.Item("RSRV_STATUS")
                                Case "C"
                                    EMsg &= vbCr & "Reservation No " & RSRV_NO & " has been Cancelled"
                                Case "D"
                                    EMsg &= vbCr & "Reservation No " & RSRV_NO & " has been Deleted"
                                Case Else ' such as "F"
                                    EMsg &= vbCr & "Reservation No " & RSRV_NO & " is No Longer Open"
                            End Select
                        End If
                    End If
                End If

                If EMsg = "" And EntryMode = "E" Then
                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", RSRV_NO) Then Exit Sub
                    If Not ASCMAIN1.Logical_Lock("SOTRSRV1", CUST_CODE) Then Exit Sub
                End If

            Case "Update"
                If Absx1.dteFor("ORDR_SHIP_DATE").Value & "" = "" _
                    Or Absx1.dteFor("ORDR_CANCEL_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Ship Date and Cancel Date are Mandatory"
                Else
                    If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") _
                     > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                        EMsg &= vbCr & "Cancel Date cannot be Prior to Ship Date"
                    End If
                End If

                If Absx1.txtFor("SREP_CODE").Text = "" Then
                    EMsg &= vbCr & "Sales Rep is required"
                Else
                    If LookUp("SOTSREP1", Absx1.txtFor("SREP_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Sales Rep"
                    End If
                End If

                If Absx1.txtFor("WHSE_CODE").Text = "" Then
                    EMsg &= vbCr & "Warehouse is required"
                Else
                    If LookUp("ICTWHSE1", Absx1.txtFor("WHSE_CODE").Text) Is Nothing Then
                        EMsg &= vbCr & "Invalid Warehouse"
                    End If
                End If

                If Absx1.txtFor("ORDR_CUST_PO").Text = "" And rowARTCUST1.Item("CUST_PO_REQD") & "" = "1" Then
                    EMsg &= vbCr & "Customer PO is required"
                End If
                If grdSOTRSRV2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Reservation"
                Else
                    If Val(dst.Tables("SOTRSRV2").Compute("COUNT(RSRV_LNO)", "RSRV_QTY > 0") & "") = 0 Then
                        EMsg &= vbCr & "No Items on Reservation with Qty >0"
                    End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTRSRV2 where RSRV_QTY_USED <> 0"
                ASCMAIN1.sql &= " and RSRV_NO = '" & RSRV_NO & "'"

                If Val(ASCDATA1.GetDataValue) <> 0 Then
                    EMsg &= vbCr & "Reservation has been Used"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Reservation as Deleted", _
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If


            Case "Cancel Balance"
                If EMsg = "" Then
                    If MsgBox("Do you want to Cancel (the remaining open balance on) this Reservation", _
                               vbYesNo, "Confirmation") = MsgBoxResult.No Then
                        Exit Sub
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

            Case "New"
                EntryMode = "N"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Delete"
                Delete_Order()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Cancel Balance"
                Cancel_Order()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
            If (EntryMode = "V" And ScreenMode) Then
                If rowSOTRSRV1.Item("RSRV_STATUS") & "" = "O" Then
                    .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.True
                Else
                    .Groups("Screen Control").Items("Edit").Settings.Enabled = DefaultableBoolean.False
                End If
            Else
                .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
            End If
            .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Delete").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("View").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Cancel Balance").Settings.Enabled = iScreenMode

            .Groups("Screen Control").Items("Done").Visible = (EntryMode = "V" And ScreenMode)
            .Groups("Screen Control").Items("Print").Visible = False ' ScreenMode
            .Groups("Screen Control").Items("Update").Visible = (Not (EntryMode = "V") Or Not ScreenMode)
            .Groups("Screen Control").Items("Delete").Visible = (EntryMode = "E")
            .Groups("Screen Control").Items("Cancel").Visible = (Not (EntryMode = "V") Or Not ScreenMode)

            .Groups("Screen Control").Items("Cancel Balance").Visible = (EntryMode = "E")

            .Groups("Totals").Visible = ScreenMode
        End With

        lblStatus.Visible = ScreenMode

        grdSOTRSRVX.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("ORDR_CUST_PO"), InquiryMode Or (ScreenMode And Not (EntryMode = "E" Or EntryMode = "N")))
        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))

        If ScreenMode Then
            If EntryMode = "V" Then
                grdSOTRSRV2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
                grdSOTRSRV2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
                grdSOTRSRV2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Else
                grdSOTRSRV2.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                grdSOTRSRV2.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
                grdSOTRSRV2.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True

                If EntryMode <> "E" Then
                    grdSOTRSRV2.DisplayLayout.Bands(0).Columns("X").Hidden = True
                Else
                    grdSOTRSRV2.DisplayLayout.Bands(0).Columns("X").Hidden = False
                End If
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("RSRV_NO").Text = ""
        Absx1.txtFor("CUST_CODE").Text = ""

        CUST_CODE = ""
        RSRV_NO = ""

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTRSRV1", "SOTRSRV2"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)
        Load_SOTRSRVX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        If EntryMode = "N" Then
            RSRV_NO = ASCMAIN1.Next_Control_No("RSRV_NO")

            rowSOTRSRV1 = dst.Tables("SOTRSRV1").NewRow
            With rowSOTRSRV1
                .Item("RSRV_NO") = RSRV_NO
                .Item("CUST_CODE") = CUST_CODE
                .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                .Item("RSRV_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID

                Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                If rowARTCUST1.Item("WHSE_CODE") & "" <> "" Then WHSE_CODE = rowARTCUST1.Item("WHSE_CODE")
                If WHSE_CODE = "" Then WHSE_CODE = ""
                .Item("WHSE_CODE") = WHSE_CODE
                '  .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
                .Item("SREP_CODE") = SREP_CODE
                .Item("SREP2_CODE") = SREP2_CODE
                .Item("RSRV_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
            End With
            dst.Tables("SOTRSRV1").Rows.Add(rowSOTRSRV1)

        Else
            rowSOTRSRV1 = Fill_Record("SOTRSRV1", RSRV_NO)
        End If

        CUST_CODE = rowSOTRSRV1.Item("CUST_CODE")
        rowARTCUST1 = Fill_Record("ARTCUST1", CUST_CODE)

        Fill_Records("SOTRSRV2", RSRV_NO)
        Sort_grdColumns(grdSOTRSRV2, "RSRV_LNO")

        lblINIT_DATE.Text = "Entered on " & Format(rowSOTRSRV1.Item("INIT_DATE"), "MM/dd/yyyy")

        If EntryMode = "N" Then
            lblStatus.Text = "New Order"
        Else
            Select Case rowSOTRSRV1.Item("RSRV_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
            End Select
        End If

        With grdSOTRSRV2.DisplayLayout.Bands(0)
            If (EntryMode = "E" Or EntryMode = "N") Then
                .Columns("RSRV_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                If EntryMode = "E" Then
                    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Else
                .Columns("RSRV_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            With grdSOTRSRV2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            End With
            grdSOTRSRV2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, False)
        Else
            With grdSOTRSRV2.DisplayLayout.Override
                .AllowAddNew = UltraWinGrid.AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            grdSOTRSRV2.DisplayLayout.Bands(0).Columns("STYLE_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
            Set_Read_Only(splHeader, True)
        End If

        Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, RSRV_NO)
        For Each TABLE_NAME As String In New String() _
            {"SOTRSRV1", "SOTRSRV2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where RSRV_NO = '" & RSRV_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating ...")

        BeginTrans()

        If EntryMode <> "N" Then Delete_Records()

        If EntryMode = "N" Then
            rowSOTRSRV1.Item("ORDR_ORIG_SHIP_DATE") = rowSOTRSRV1.Item("ORDR_SHIP_DATE")
            rowSOTRSRV1.Item("ORDR_ORIG_CANCEL_DATE") = rowSOTRSRV1.Item("ORDR_CANCEL_DATE")
        End If

        INIT_LAST("SOTRSRV1", False, , True)
        Dim sqldelete As String = "RSRV_NO = '" & RSRV_NO & "'"
        Update_Record_TDA("SOTRSRV1", sqldelete)
        Update_Record_TDA("SOTRSRV2", sqldelete)
        Dependent_Updates(1, RSRV_NO)

        CommitTrans("Update Complete")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "RSRV_NO"

                If Absx1.txtFor("CUST_CODE").Text = "" And Absx1.txtFor("ORDR_CUST_PO").Text = "" Then
                    MsgBox("You must enter a Customer Code or a PO No", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTRSRV1.RSRV_STATUS = 'O' "
                End If

                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    sql_where &= " and SOTRSRV1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
                End If
                If Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                    sql_where &= " and SOTRSRV1.ORDR_CUST_PO = '" & Absx1.txtFor("ORDR_CUST_PO").Text & "'"
                End If
        End Select
    End Sub

    Public Overrides Function Remote_Control( _
    ByVal command As String, _
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                'If ScreenMode Then
                '    Click_Command("Done")
                'End If

                Absx1.txtFor("RSRV_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTRSRV1"
            E.COLUMN_NAME = "RSRV_NO"
            E.CODE_VALUE = Absx1.txtFor("RSRV_NO").Text
            E.DESC_VALUE = "Reservation"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTRSRV1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTRSRVX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTRSRV2, "BB", "Style Status Inquiry", "Style Multi-Color")
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

        Select Case grd.Name
            Case "grdSOTRSRV2"
                tlb_btn = DirectCast(tlb_pop.Tools("Style Multi-Color"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = (EntryMode = "E" Or EntryMode = "N")

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
           

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Style Status Inquiry"
                Dim STYLE_CODE As String = grd.ActiveRow.Cells("STYLE_CODE").Text
                Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                If rowICTSTYL1 IsNot Nothing Then
                    Context_Launch("Select", STYLE_CODE, e.Tool.Key, "ICFSTAT1")
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Load_SOTRSRVX()
                End If

            Case "ORDR_CUST_PO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode _
                       And Absx1.txtFor("CUST_CODE").Text <> "" _
                       And Absx1.txtFor("ORDR_CUST_PO").Text <> "" Then
                        Click_Command("New")
                    End If
                End If
            Case "RSRV_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If Not ScreenMode Then
                    Load_SOTRSRVX()

                    Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
                    If CUST_CODE <> "" Then
                        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Load_SOTRSRVX()
            Case "RSRV_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub
#End Region

    Sub Load_SOTRSRVX()
        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        If CUST_CODE = "" Then
            ASCMAIN1.sql = "Select * from SOTRSRV1 where RSRV_STATUS = 'O'"
            Fill_Records("SOTRSRVX", "", , ASCMAIN1.sql)
            grdSOTRSRVX.Text = "Open Reservations"
            Sort_grdColumns(grdSOTRSRVX, "RSRV_NO".ToLower)
        Else
            ASCMAIN1.sql = "Select * from SOTRSRV1 where RSRV_STATUS = 'O'" _
                & " and CUST_CODE = '" & CUST_CODE & "'"
            Fill_Records("SOTRSRVX", "", , ASCMAIN1.sql)
            grdSOTRSRVX.Text = "Open Reservations associated with " & CUST_CODE
            Sort_grdColumns(grdSOTRSRVX, "RSRV_NO".ToLower)
        End If
        grdSOTRSRVX.Visible = True
    End Sub

    Sub Print_Record()
        ' NOTE THAT THIS PRINT ROUTINE WAS USING THE DATA LAYER & DST THAT IS ASSOCIATED WITH THIS FORM
        ' AND SHOULD BE USING THE DATALAYER OF SORUSED1

        'Fill_Records("SOTSVIA1", USED_CODE)

        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Dim RPT As String = "SORUSED1" ' unneccesary if Report Name is Like Form Name
        'Generate_Report(RPT, "USEDper Invoice Report", , , , , False)
        'Print_Report_End()


        Dim REPORTFILE As String = "SORUSED1"
        If Not REPORTS.ContainsKey(REPORTFILE) Then
            REPORTS.Add(REPORTFILE, Load_rptClass(REPORTFILE))
            REPORTS(REPORTFILE).Prepare_dst(False, "")
        End If

        'To fill the report's dataset with data from Oracle, 
        ' set the parameter array to values that the Fill_Records_RPT method expects, and then call it
        'REPORTS(REPORTFILE).Fill_Records_RPT(New String() {"USEDPER_CTL_NO = '" & USEDPER_CTL_NO & "'"})

        'To fill the report's dataset with data from this form's dataset:
        With REPORTS(REPORTFILE).clsASCBASE1
            .EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SOTPPDI1", "SOTPPDI2", "SOTPPDI3", "SOTINVH1", "SOTSVIA1"}
                .dst.Tables(TABLE_NAME).Rows.Clear()
                Dim SQL As String = ""
                If TABLE_NAME = "SOTINVH1" Then
                    SQL = "RSRV_NO = '" & RSRV_NO & "'"
                End If

                For Each row As DataRow In dst.Tables(TABLE_NAME).Select(SQL)
                    Dim rowr As DataRow = .dst.Tables(TABLE_NAME).NewRow
                    If TABLE_NAME = "SOTPPDI2" Or TABLE_NAME = "SOTPPDI3" Or TABLE_NAME = "SOTINVH1" Then

                        For I As Integer = 0 To .dst.Tables(TABLE_NAME).Columns.Count - 1
                            Dim COLUMN_NAME As String = .dst.Tables(TABLE_NAME).Columns(I).ColumnName
                            rowr.Item(COLUMN_NAME) = row.Item(COLUMN_NAME)
                        Next
                    Else
                        rowr.ItemArray = row.ItemArray
                    End If
                    .dst.Tables(TABLE_NAME).Rows.Add(rowr)
                Next
            Next
            .EnforceConstraints(True)
        End With
        With REPORTS(REPORTFILE).clsASCBASE1
            .Print_Report_Begin()
            .CR_params.Add("SUBT", "")
            .Generate_Report("SORUSED1", "USEDper Invoice Report", , True, , , , , False)
            .Print_Report_End()
        End With

    End Sub

    Private Sub grdSOTRSRVX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTRSRVX.DoubleClickRow
        Absx1.txtFor("RSRV_NO").Text = e.Row.Cells("RSRV_NO").Value
        Click_Command("View")
    End Sub

    Sub Update_ICTSTAT2(STYLE_CODE As String, COLOR_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVVNNNNNN", _
                           New Object() {STYLE_CODE, COLOR_CODE, WHSE_CODE, _
                                         0, 0, 0, _
                                         QTY, 0, 0}, _
                           New String() {"STYLE_CODE_IN", "COLOR_CODE_IN", "WHSE_CODE_IN", _
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in", _
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Sub Cancel_Order()
        Me.Cursor = Cursors.WaitCursor
        BeginTrans()

        Dim EMsg As String = ""
        If EntryMode = "E" Then
            Cancel_Order_1(RSRV_NO)
            EMsg = "Reservation " & RSRV_NO & " has been Cancelled"
        End If

        'ASCDATA1.ExecuteSP("SOPRSRV0_G", "V", New Object() {RSRV_GROUP_NO}, New String() {"RSRV_GROUP_NO_IN"})
        CommitTrans(EMsg)
        Me.Cursor = Cursors.Default
    End Sub

    Sub Cancel_Order_1(RSRV_NO As String)
        Dependent_Updates(-1, RSRV_NO)

        ASCMAIN1.sql = "" _
            & "Begin " _
            & " Declare Cursor C1 is Select * from SOTRSRV2 where RSRV_NO = '" & RSRV_NO & "' for Update;" _
            & " Begin " _
            & "  For R1 in C1 Loop" _
            & "   Update SOTRSRV2" _
            & "    Set RSRV_QTY_CANC = NVL(RSRV_QTY_CANC,0) + NVL(R1.RSRV_QTY_OPEN,0)" _
            & "      , RSRV_QTY_OPEN = 0" _
            & "    where Current of C1;" _
            & "  End Loop;" _
            & " End;" _
            & "End;"
        ASCDATA1.ExecuteSQL()
        ', RSRV_STATUS = 'C'
        ASCMAIN1.sql = "Update SOTRSRV1 Set RSRV_STATUS = :PARM1" _
            & " where RSRV_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"C", RSRV_NO})
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        Dim EMsg As String = ""

        BeginTrans()

        If EntryMode = "E" Then
            Delete_Order_1(RSRV_NO)
            EMsg = "Reservation No " & RSRV_NO & " has been marked as Deleted"
        End If

        CommitTrans(EMsg)
        'ASCDATA1.ExecuteSP("SOPRSRV0_G", "V", New Object() {RSRV_GROUP_NO}, New String() {"RSRV_GROUP_NO_IN"})
        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(RSRV_NO As String)
        Dependent_Updates(-1, RSRV_NO)

        ASCMAIN1.sql = "" _
            & "Begin " & vbCrLf _
            & " Declare Cursor C1 is" & vbCrLf _
            & "  Select * from SOTRSRV2" & vbCrLf _
            & "     where RSRV_NO = '" & RSRV_NO & "' for Update;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update SOTRSRV2" & vbCrLf _
            & "    Set RSRV_QTY_CANC = NVL(RSRV_QTY_CANC,0) + NVL(R1.RSRV_QTY_OPEN,0)" & vbCrLf _
            & "   , RSRV_QTY_OPEN = 0" & vbCrLf _
            & "    where Current of C1;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SOTRSRV1 Set RSRV_STATUS = :PARM1" _
            & " where RSRV_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"D", RSRV_NO})
    End Sub

    Sub Dependent_Updates(S As Integer, RSRV_NO As String)

        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from SOTRSRV1 where RSRV_NO = '" & RSRV_NO & "'"
        Dim rowSOTRSRV1 As DataRow = ASCDATA1.GetDataRow
        Dim WHSE_CODE As String = rowSOTRSRV1.Item("WHSE_CODE") & ""
        If WHSE_CODE = "" Then WHSE_CODE = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE") & ""

        ASCMAIN1.sql = "Select * from SOTRSRV2 where RSRV_NO = '" & RSRV_NO & "'"
        For Each rowSOTRSRV2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim STYLE_CODE As String = rowSOTRSRV2.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowSOTRSRV2.Item("COLOR_CODE")

            QTY_TO_COMMIT = Val(rowSOTRSRV2.Item("RSRV_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                STYLE_CODE = rowSOTRSRV2.Item("STYLE_CODE")
                COLOR_CODE = rowSOTRSRV2.Item("COLOR_CODE")
                Update_ICTSTAT2(STYLE_CODE, COLOR_CODE, WHSE_CODE, S * QTY_TO_COMMIT)
            End If
        Next
    End Sub

    Sub Display_Totals()

        Dim KEY As Int32 = 0
        For Each SFX As String In New String() {"", "OPEN", "ALLO", "USED", "CANC"}
            If SFX <> "" Then SFX = "_" & SFX
            KEY += 1
            Dim rowSOTRSRVT As DataRow = dst.Tables("SOTRSRVT").Rows.Find(KEY)
            rowSOTRSRVT.Item("QTY") = Val(dst.Tables("SOTRSRV2").Compute("SUM(RSRV_QTY" & SFX & ")", "") & "")
            rowSOTRSRVT.Item("AMT") = Val(dst.Tables("SOTRSRV2").Compute("SUM(RSRV_AMT" & SFX & ")", "") & "")
        Next
    End Sub

    Sub Validate_Columns(COLUMN_NAME As String, ByRef Cancel As Boolean)
        With grdSOTRSRV2.ActiveRow
            Select Case COLUMN_NAME
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = ""
                    If Trim(.Cells("STYLE_CODE").Value & "") <> "" Then
                        STYLE_CODE = Validate_Style(.Cells("STYLE_CODE").Value & "")
                    End If
                    Cancel = (STYLE_CODE = "")

                Case "COLOR_CODE"
                    If .Cells("COLOR_CODE").Value & "" <> "" Then
                        If Not COLOR_CODEs.Contains(.Cells("COLOR_CODE").Value & "") Then
                            Cancel = True
                        End If
                    Else
                        Cancel = True
                    End If
                Case "RSRV_QTY"
                    If Trim(.Cells("STYLE_CODE").Value & "") = "" Then
                        Cancel = True
                        Exit Sub
                    End If
                    If Trim(.Cells("RSRV_QTY").Value & "") = "" Then
                        MsgBox("Qty Not Specified", vbOKOnly, "Cannot Update Record")
                        Cancel = True
                        grdSOTRSRV2.ActiveCell = grdSOTRSRV2.ActiveRow.Cells("RSRV_QTY")
                        Exit Sub
                    End If
                    If Val(.Cells("RSRV_QTY").Value & "") < 0 Then
                        MsgBox("Qty May Not be Negative", vbOKOnly, "Invalid Quantity")
                        Cancel = True
                    End If
            End Select
        End With
    End Sub

    Function Validate_Style(STYLE_CODE_z As String) As String
        Dim EMsg As String = ""
        If STYLE_CODE_z = "" Then Return ""

        Dim STYLE_CODE As String = ""
        rowICTSTYL1 = LookUp("ICTSTYL1", STYLE_CODE_z)

        If rowICTSTYL1 Is Nothing Then
            EMsg = "Style is Not on File" & vbCrLf
        Else
            If rowICTSTYL1.Item("STYLE_STATUS") & "" <> "A" Then
                EMsg = "Item Status is not Active" & vbCrLf
            End If
            If rowICTSTYL1.Item("STYLE_UOM") & "" = "" Then
                EMsg = "Item does not have a valid Unit of Measure" & vbCrLf
            End If
            If rowICTSTYL1.Item("SALES_DIVISION_CODE") & "" = "" Then
                EMsg = "Item does not have a valid Division Code" & vbCrLf
            End If
        End If

        If EMsg = "" Then
            COLOR_CODEs.Clear()
            Fill_Records("ICTCOLRS", STYLE_CODE_z)
            For Each row As DataRow In dst.Tables("ICTCOLRS").Select("")
                Dim COLOR_CODE As String = row.Item("COLOR_CODE")
                COLOR_CODEs.Add(COLOR_CODE)
            Next
        End If

        If EMsg <> "" And grdSOTRSRV2.ActiveRow.IsAddRow Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Style Code Entered is Invalid because ...")
        Else
            If EMsg = "" Then
                STYLE_CODE = rowICTSTYL1.Item(0)
            End If
        End If
        Return STYLE_CODE
    End Function

#Region "grdSOTRSRV2"

    Private Sub grdSOTRSRV2_AfterCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRSRV2.AfterCellUpdate
        With grdSOTRSRV2.ActiveRow
            Select Case e.Cell.Column.Key
                Case "STYLE_CODE"
                    Dim STYLE_CODE As String = Validate_Style(.Cells("STYLE_CODE").Value)
                    If STYLE_CODE <> "" Then
                        .Cells("STYLE_DESC").Value = rowICTSTYL1.Item("STYLE_DESC")
                        If COLOR_CODEs.Count = 1 Then
                            .Cells("COLOR_CODE").Value = COLOR_CODEs(0)
                        End If
                    End If

                Case "COLOR_CODE"
                    Dim COLOR_CODE As String = e.Cell.Value & ""
                    If COLOR_CODE <> "" Then
                        Dim rowICTCOLR1 As DataRow = LookUp("ICTCOLR1", COLOR_CODE)
                        If rowICTCOLR1 IsNot Nothing Then
                            .Cells("COLOR_DESC").Value = rowICTCOLR1.Item("COLOR_DESC")
                        End If
                    End If

                Case "RSRV_QTY"
                    .Cells("RSRV_QTY_OPEN").Value = .Cells("RSRV_QTY").Value

                Case "RSRV_QTY_OPEN"
                    .Cells("RSRV_QTY_CANC").Value _
                        = Val(.Cells("RSRV_QTY").Value & "") _
                        - Val(.Cells("RSRV_QTY_USED").Value & "") _
                        - Val(.Cells("RSRV_QTY_OPEN").Value & "")
                    If Val(.Cells("RSRV_QTY_CANC").Value) < 0 Then
                        .Cells("RSRV_QTY_CANC").Value = 0
                    End If
            End Select
        End With
    End Sub

    Private Sub grdSOTRSRV2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSOTRSRV2.AfterRowActivate

        If Trim(grdSOTRSRV2.ActiveRow.Cells("STYLE_CODE").Value & "") = "" And _
            (grdSOTRSRV2.ActiveCell Is Nothing OrElse _
             (grdSOTRSRV2.ActiveCell.Column.Key <> "STYLE_CODE")) _
        Then
            grdSOTRSRV2.ActiveCell = grdSOTRSRV2.ActiveRow.Cells("STYLE_CODE")
            Exit Sub
        End If

        If grdSOTRSRV2.ActiveRow.IsAddRow Then
            grdSOTRSRV2.DisplayLayout.Bands(0).Columns("RSRV_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
            grdSOTRSRV2.DisplayLayout.Bands(0).Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.NoEdit
            If grdSOTRSRV2.ActiveRow.Cells("STYLE_CODE").Value & "" = "" Then
                grdSOTRSRV2.ActiveCell = grdSOTRSRV2.ActiveRow.Cells("STYLE_CODE")
            End If
        Else
            With grdSOTRSRV2.DisplayLayout.Bands(0)
                Validate_Style(grdSOTRSRV2.ActiveRow.Cells("STYLE_CODE").Value & "")

                If Val(grdSOTRSRV2.ActiveRow.Cells("RSRV_QTY_USED").Value & "") <> 0 _
                Or Val(grdSOTRSRV2.ActiveRow.Cells("RSRV_QTY_CANC").Value & "") <> 0 _
                Then
                    .Columns("RSRV_QTY").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("RSRV_QTY").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("RSRV_QTY_OPEN").CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            End With
        End If
    End Sub

    Private Sub grdSOTRSRV2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdSOTRSRV2.AfterRowsDeleted
        Display_Totals()

        If grdSOTRSRV2.Rows.Count = 0 Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = ""
        End If
    End Sub

    Private Sub grdSOTRSRV2_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdSOTRSRV2.AfterRowUpdate
        Display_Totals()

        If Absx1.txtFor("SALES_DIVISION_CODE").Text = "" Then
            Absx1.txtFor("SALES_DIVISION_CODE").Text = rowICTSTYL1.Item("SALES_DIVISION_CODE")
        End If
    End Sub

    Private Sub grdSOTRSRV2_BeforeCellUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdSOTRSRV2.BeforeCellUpdate
     
    End Sub

    Private Sub grdSOTRSRV2_BeforeExitEditMode(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdSOTRSRV2.BeforeExitEditMode
        If grdSOTRSRV2.ActiveCell IsNot Nothing Then
            With grdSOTRSRV2.ActiveCell
                Select .Column.Key
                    Case "STYLE_CODE", "COLOR_CODE"
                        .EditorResolved.Value = ASCMAIN1.Format_Field(.EditorResolved.Value & "", .Column.Key)
                End Select
            End With
        End If
    End Sub

    Private Sub grdSOTRSRV2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdSOTRSRV2.BeforeRowsDeleted

        RSRV_LNOs.Clear()

        For Each grow As UltraWinGrid.UltraGridRow In e.Rows
            If Val(grow.Cells("RSRV_QTY_USED").Value & "") <> 0 _
            Or Val(grow.Cells("RSRV_QTY_CANC").Value & "") <> 0 _
            Then
                MsgBox("Cannot Delete a Line if it has ever been " & vbCr & "Used Or Cancelled" & vbCr & "Use the Cancel Button (x)")
                e.Cancel = True
                Exit Sub
            End If

            RSRV_LNOs.Add(grow.Cells("RSRV_LNO").Value)
        Next
    End Sub

    Private Sub grdSOTRSRV2_BeforeRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTRSRV2.BeforeRowUpdate

        Validate_Columns("STYLE_CODE", e.Cancel)
        If Not e.Cancel Then
            Validate_Columns("COLOR_CODE", e.Cancel)
        End If
        If Not e.Cancel Then
            Validate_Columns("RSRV_QTY", e.Cancel) 'THIS IS DONE IN BEFORECOLUPDATE
        End If

        If e.Cancel = True Then
            Exit Sub
        End If

        If e.Row.IsAddRow Then
            e.Row.Cells("RSRV_NO").Value = RSRV_NO
            Dim RSRV_LNO As Int64 = Val(dst.Tables("SOTRSRV2").Compute("MAX(RSRV_LNO)", "") & "") + 1
            e.Row.Cells("RSRV_LNO").Value = RSRV_LNO
        End If
    End Sub

    Private Sub grdSOTRSRV2_ClickCellButton(sender As Object, e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdSOTRSRV2.ClickCellButton

        With e.Cell.Row
            Select Case e.Cell.Column.Key
                Case "X"
                    If Val(.Cells("RSRV_QTY_CANC").Value) <> 0 Then
                        If MsgBox("Restore Cancelled Qty of " & .Cells("RSRV_QTY_CANC").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowSOTRSRVR As DataRow In dst.Tables("SOTRSRVR").Select("RSRV_LNO = " & .Cells("RSRV_LNO").Value)
                                rowSOTRSRVR.Item("RSRV_QTY_OPEN") = Val(rowSOTRSRVR.Item("RSRV_QTY_OPEN") & "") + Val(rowSOTRSRVR.Item("RSRV_QTY_CANC") & "")
                                rowSOTRSRVR.Item("RSRV_QTY_CANC") = 0
                            Next
                        End If
                        .Cells("RSRV_QTY_OPEN").Value = Val(.Cells("RSRV_QTY_OPEN").Value & "") + Val(.Cells("RSRV_QTY_CANC").Value & "")
                        ' grdSOWRSRV2_AfterColUpdate(.Cells("RSRV_QTY_OPEN").position)
                        .Update()
                    Else
                        If MsgBox("Cancel Remaining Qty Open of " & .Cells("RSRV_QTY_OPEN").Value, _
                                  MsgBoxStyle.YesNo, "Option to Restore Qty Cancelled") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                        If .Cells("RANGE_STYLE_CODE").Value & "" <> "" Then
                            For Each rowSOTRSRVR As DataRow In dst.Tables("SOTRSRVR").Select("RSRV_LNO = " & .Cells("RSRV_LNO").Value)
                                rowSOTRSRVR.Item("RSRV_QTY_OPEN") = 0
                                Dim RSRV_QTY_CANC As Int64 = Val(rowSOTRSRVR.Item("RSRV_QTY") & "") _
                                                           - Val(rowSOTRSRVR.Item("RSRV_QTY_USED") & "") _
                                                           - Val(rowSOTRSRVR.Item("RSRV_QTY_OPEN") & "")
                                rowSOTRSRVR.Item("RSRV_QTY_CANC") = IIf(RSRV_QTY_CANC < 0, 0, RSRV_QTY_CANC)
                            Next
                        End If
                        .Cells("RSRV_QTY_OPEN").Value = "0"
                        ' grdSOWRSRV2_AfterColUpdate(.Cells("RSRV_QTY_OPEN").position)
                        grdSOTRSRV2.ActiveRow.Update()
                    End If

                Case "STYLE_CODE"
                    Dim sql_where As String = ""
                    grdClickCellButton(grdSOTRSRV2, sql_where)

                Case "COLOR_CODE"
                    Dim sql_where As String = "COLOR_CODE IN (Select COLOR_CODE from ICTSTYC1 where STYLE_CODE ='" & grdSOTRSRV2.ActiveRow.Cells("STYLE_CODE").Value & "')"

                    grdClickCellButton(grdSOTRSRV2, sql_where)
            End Select
        End With

    End Sub
#End Region

    Sub Load_Events()
        '    grdEvents.RemoveAll
        '    Call Load_Events_1("Entered", "INIT_DATE")
        '    Call Load_Events_1("Modified", "LAST_DATE")
        '    Call Load_Events_1("Released", "RSRV_DATE_REL")
        '    Call Load_Events_1("ALLO Ticket", "RSRV_DATE_ALLO_PRT")
        '    Call Load_Events_1("Packed", "RSRV_DATE_PACKED")
        '    Call Load_Events_1("USEDped", "RSRV_DATE_USEDPED")
        '    Call Load_Events_1("Invoice", "RSRV_INV_DATE")
        '    Call Load_Events_1("Invoice Prt", "RSRV_DATE_INV_PRT")
        '    Call Load_Events_1("Updated", "RSRV_DATE_UPDATED")
        '    Call Load_Events_1("Cancelled", "RSRV_DATE_CANCELLED")
    End Sub

    Sub Add_Colors(STYLE_CODE As String, tbl As DataTable)
        If tbl.Select("ISNULL(QTY,0)<>0").Length = 0 Then
            MsgBox("No Qty's Entered", MsgBoxStyle.OkOnly, "Cannot Add Colors")
            Exit Sub
        End If

        For Each rowICTCOLRM As DataRow In tbl.Select("ISNULL(QTY,0)<>0", "COLOR_CODE")
            grdSOTRSRV2.DisplayLayout.Bands(0).AddNew()
            With grdSOTRSRV2.ActiveRow
                .Cells("STYLE_CODE").Value = STYLE_CODE
                .Cells("COLOR_CODE").Value = rowICTCOLRM.Item("COLOR_CODE")
                .Cells("RSRV_QTY").Value = rowICTCOLRM.Item("QTY")
                .Update()
            End With
        Next
        Sort_grdColumns(grdSOTRSRV2, "RSRV_LNO")
    End Sub
End Class