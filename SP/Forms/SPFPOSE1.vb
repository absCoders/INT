Public Class SPFPOSE1

    Dim DEPT_CODE As String
    Dim PROMOTION_WEEK As String
    Dim SEASON_CODE As String
    Dim SEASON_SEQ_NO As String
    Dim VEND_CODE As String
    Dim DGC_CODE As String
    Dim PROMOTION_CODE As String

    Dim SPTSPXP1 As String
    Dim GMTSTYL1 As String

    Dim REGION_CODEs As New Dictionary(Of Integer, String)
    Dim iREGION_CODEs As New List(Of Integer)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Load_GMTSCOLX(True)

        With dst

            ASCMAIN1.sql = "Select * from SPTSPXP1 where ROWNUM < 1"
            SPTSPXP1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SPTSPXP1 & " Add Primary Key (DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE,PROMOTION_WEEK,REGION_CODE)")
            Create_TDA(.Tables.Add("SPTSPXP1"), SPTSPXP1, "*")

            ASCMAIN1.sql = "Select VEND_CODE CODE_VALUE, VEND_NAME DESC_VALUE from APTVEND1"
            Create_TDA(.Tables.Add, "SPTSPXPC", "**", 0, False, , 1)
            .Tables("SPTSPXPC").Columns.Add("RECORDS", GetType(System.Int32))

            ASCMAIN1.sql = "Select * from GMTSTYL1 where ROWNUM < 1"
            GMTSTYL1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & GMTSTYL1 & " Add Primary Key (DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE)")

            ASCMAIN1.sql = "Select DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE from " & GMTSTYL1
            Create_TDA(.Tables.Add, "SPTSPXPR", "**", 0, False, , 4)
            With dst.Tables("SPTSPXPR").Columns
                For I As Integer = 1 To 9
                    .Add("PROMO_CODE_" & Format(I, "0"))
                Next
            End With

            ASCMAIN1.sql = "Select * from GMTREGN1"
            Create_TDA(.Tables.Add, "GMTREGN1", "**", 0, False)
            .Tables("GMTREGN1").Columns.Add("SEL")
            .Tables("GMTREGN1").Columns.Add("I", GetType(System.Int16))

            ASCMAIN1.sql = "Select * from SPTPROM1"
            Create_TDA(.Tables.Add, "SPTPROM1", "**", 0, False)
            '  .Tables("SPTPROM1").Columns.Add("PROMOTION_DESC")
        End With

        grdSPTSPXPC.DataSource = dst.Tables("SPTSPXPC")
        grdSPTSPXPR.DataSource = dst.Tables("SPTSPXPR")
        grdSPTPROM1.DataSource = dst.Tables("SPTPROM1")
        grdGMTREGN1.DataSource = dst.Tables("GMTREGN1")
        Sort_grdColumns(grdGMTREGN1, "REGION_CODE")

        Fill_Records("SPTPROM1")

        Create_Summary(grdSPTSPXPC, "CODE_VALUE", "Count")

        Create_Summary(grdSPTSPXPR, "DGC_CODE", "Count")
        '  Create_Summary(grdSPTSPXP1, New String() {"OH_STR", "OH_WHS", "QTY_ORD", "QTY_OPN_1", "QTY_OPN_2", "QTY_REC"})


        Dim iREGION_CODE As Integer = 0
        Fill_Records("GMTREGN1")
        For Each rowGMTREGN1 As DataRow In dst.Tables("GMTREGN1").Select("", "REGION_CODE")
            rowGMTREGN1.Item("SEL") = "1"
            iREGION_CODE += 1
            rowGMTREGN1.Item("I") = iREGION_CODE
            iREGION_CODEs.Add(iREGION_CODE)
            REGION_CODEs.Add(iREGION_CODE, rowGMTREGN1.Item("REGION_CODE"))
        Next

        With grdSPTSPXPR.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If New String() {"DGC_CODE", "VEND_CODE", "STYLE_CODE", "COLOR_CODE"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                ElseIf gcol.Key.StartsWith("PROMO_CODE_") Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    gcol.Width = 50
                    Dim I As Integer = Val(Mid(gcol.Key, 12, 1))
                    If I > REGION_CODEs.Count Then
                        gcol.Hidden = True
                    Else
                        gcol.Header.Caption = REGION_CODEs(I)
                    End If
                End If
            Next
        End With


        Show_Filter(grdSPTSPXPC, True)
        Show_Filter(grdSPTSPXPR, True)


        optWeek.ValueList.ValueListItems(0).DisplayText = "TW (" & Mid(ASCMAIN1.CYW, 5, 2) & ")"
        optWeek.ValueList.ValueListItems(1).DisplayText = "NW (" & Mid(ASCMAIN1.Week_Calc(ASCMAIN1.CYW, 1), 5, 2) & ")"

        Set_optStyles_Dependencies()

        '    ASCMAIN1.Add_Value_List(grdPOTORDRX, "SEASON_SEQ_NO", "SELECT SEASON_SEQ_NO, SEASON_CODE FROM GMTSEAS1 WHERE SEASON_ACTIVE = '1'")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Edit", "View"

                DEPT_CODE = ""
                PROMOTION_WEEK = ""
                VEND_CODE = ""
                DGC_CODE = ""
                SEASON_CODE = ""
                SEASON_SEQ_NO = ""

                If Validate_Code("DEPT_CODE") Then DEPT_CODE = Absx1.txtFor("DEPT_CODE").Text
                If Validate_Code("PROMOTION_WEEK") Then PROMOTION_WEEK = Absx1.txtFor("PROMOTION_WEEK").Text

                'If Validate_Code("SEASON_CODE", , True) Then SEASON_CODE = Absx1.txtFor("SEASON_CODE").Text
                ASCMAIN1.sql = "SELECT * FROM GMTSEAS1 WHERE SEASON_ACTIVE = '1' AND SEASON_CODE = :PARM1"
                Dim rowGMTSEAS1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {Absx1.txtFor("SEASON_CODE").Text}) ' LookUp("GMTSEAS1", SEASON_CODE)
                If rowGMTSEAS1 IsNot Nothing Then
                    SEASON_CODE = rowGMTSEAS1.Item("SEASON_CODE")
                    SEASON_SEQ_NO = rowGMTSEAS1.Item("SEASON_SEQ_NO")
                Else
                    EMsg &= vbCr & "Invalid Value specified for Season " & Absx1.txtFor("SEASON_CODE").Text
                End If

                If Absx1.txtFor("VEND_CODE").Text = "" And Absx1.txtFor("DGC_CODE").Text = "" Then
                    EMsg &= vbCr & "You must specify either a Vendor or a DGC before Proceeding"
                ElseIf Absx1.txtFor("VEND_CODE").Text <> "" And Absx1.txtFor("DGC_CODE").Text <> "" Then
                    EMsg &= vbCr & "You must specify one or the other, but not both: Vendor or DGC"
                Else
                    If Validate_Code("VEND_CODE", , True) Then VEND_CODE = Absx1.txtFor("VEND_CODE").Text
                    If Validate_Code("DGC_CODE", , True) Then DGC_CODE = Absx1.txtFor("DGC_CODE").Text
                End If

                If eItemKey = "Edit" Then
                    If EMsg = "" Then
                        If Not ASCMAIN1.Logical_Lock("SPTSPXP1", DEPT_CODE & ":" & PROMOTION_WEEK) Then
                            Exit Sub
                        End If
                    End If
                End If


            Case "Update"

                If Absx1.txtFor("STYLE_COLLECTION_DESC").Text = "{New Style Collection}" Then
                    EMsg &= "Invalid Description for a Style Collection"
                End If

                If Absx1.txtFor("STYLE_COLLECTION_DESC").Text = "" Then
                    EMsg &= vbCr & "A Style Collection Description is Mandatory"
                Else

                End If

            Case "Delete"

                If MsgBox("Do you really want to Delete the entire Style Collection?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Done", "Cancel"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    If (EntryMode = "V" And ScreenMode) Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = not_iScreenMode
                    End If
                    .Items("Update").Settings.Enabled = iScreenMode

                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    '   .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Edit").Visible = Not InquiryMode

                    .Items("Print").Visible = ScreenMode
                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                    ' .Items("Print").Visible = (EntryMode = "V" And ScreenMode) ' False ' ScreenMode
                    .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                    .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                    .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                End With

                .Groups("Regions").Visible = ScreenMode
                .Groups("Promotion Codes").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        'Set_Read_Only_for_ctl(Absx1.txtFor("STYLE_COLLECTION_DESC"), EntryMode = "V")
        'Set_Read_Only_for_ctl(Absx1.txtFor("VEND_BUYER_CODE"), EntryMode = "V")

        'Absx1.txtFor("VEND_BUYER_CODE").Visible = ScreenMode

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSPTSPXPR, grdGMTREGN1}
            With grd.DisplayLayout.Override
                If EntryMode = "E" Or EntryMode = "N" Then
                    .AllowUpdate = DefaultableBoolean.True
                Else
                    .AllowUpdate = DefaultableBoolean.False
                End If
            End With
        Next

        grdGMTSCOLX.Visible = Not ScreenMode
        splSPTXPSP1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"SPTSPXP1", "SPTSPXPC", "SPTSPXPR"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        '  dst.Relations.Remove("SPTSPXPC_SPTSPXPR")
        EnforceConstraints(True)

        '   Load_GMTSCOLX(False)
    End Sub

    Sub Print_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        'Print_Report_Begin()
        'CR_params.Add("SUBT", STYLE_COLLECTION_CODE & "-" & Absx1.txtFor("STYLE_COLLECTION_DESC").Text)
        'Generate_Report("GMRSCOL1", "Style Collection Members")

        'If optIC.Value = "A" Then
        '    CR_params.Add("SUBT", STYLE_COLLECTION_CODE & "-" & Absx1.txtFor("STYLE_COLLECTION_DESC").Text)
        '    CR_params.Add("XTD", optWMS.Text)
        '    Generate_Report("GMRSCOL2", "Style Collection Status & Activity")
        'End If

        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTSPXP1)
        ASCMAIN1.sql = "SELECT SPTSPXP1.*" _
            & " from GMTSTYL1,SPTSPXP1" _
            & " where SPTSPXP1.DGC_CODE = GMTSTYL1.DGC_CODE" _
            & "   and SPTSPXP1.VEND_CODE = GMTSTYL1.VEND_CODE" _
            & "   and SPTSPXP1.STYLE_CODE = GMTSTYL1.STYLE_CODE" _
            & "   and SPTSPXP1.COLOR_CODE = GMTSTYL1.COLOR_CODE"
        If VEND_CODE <> "" Then
            ASCMAIN1.sql &= "   and SPTSPXP1.VEND_CODE = '" & VEND_CODE & "'"
        End If
        If DGC_CODE <> "" Then
            ASCMAIN1.sql &= "   and SPTSPXP1.DGC_CODE = '" & DGC_CODE & "'"
        End If
        If SEASON_CODE <> "" Then
            ASCMAIN1.sql &= "   and GMTSTYL1.SEASON_CODE = '" & SEASON_CODE & "'"
        End If

        ASCDATA1.ExecuteSQL("Insert into " & SPTSPXP1 & " " & ASCMAIN1.sql)
        Fill_Records("SPTSPXP1")


        ASCDATA1.ExecuteSQL("Truncate Table " & GMTSTYL1)

        'ASCMAIN1.sql = "Select GMTSTYL1.DGC_CODE, GMTSTYL1.VEND_CODE, GMTSTYL1.STYLE_CODE, GMTSTYL1.COLOR_CODE from GMTSTYL1" _
        ASCMAIN1.sql = "Select GMTSTYL1.* from GMTSTYL1" _
            & " where GMTSTYL1.DGC_CODE like '" & DEPT_CODE & "%'"
        If VEND_CODE <> "" Then
            ASCMAIN1.sql &= "   and GMTSTYL1.VEND_CODE = '" & VEND_CODE & "'"
        End If
        If DGC_CODE <> "" Then
            ASCMAIN1.sql &= "   and GMTSTYL1.DGC_CODE = '" & DGC_CODE & "'"
        End If
        If SEASON_CODE <> "" Then
            ASCMAIN1.sql &= "   and GMTSTYL1.SEASON_CODE = '" & SEASON_CODE & "'"
        End If

        ASCDATA1.ExecuteSQL("Insert into " & GMTSTYL1 & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select GMTSTYL1.DGC_CODE, GMTSTYL1.VEND_CODE, GMTSTYL1.STYLE_CODE, GMTSTYL1.COLOR_CODE" & vbCrLf
        For I As Integer = 1 To REGION_CODEs.Count
            ASCMAIN1.sql &= ", MAX (DECODE(SPTSPXP1.REGION_CODE,'" & REGION_CODEs(I) & "',PROMOTION_CODE,NULL)) PROMO_CODE_" & Format(I, "0") & vbCrLf
        Next
        ASCMAIN1.sql &= " from " & GMTSTYL1 & " GMTSTYL1, " & SPTSPXP1 & " SPTSPXP1" & vbCrLf _
            & " where SPTSPXP1.DGC_CODE (+) = GMTSTYL1.DGC_CODE" & vbCrLf _
            & "   and SPTSPXP1.VEND_CODE (+) = GMTSTYL1.VEND_CODE" & vbCrLf _
            & "   and SPTSPXP1.STYLE_CODE (+) = GMTSTYL1.STYLE_CODE" & vbCrLf _
            & "   and SPTSPXP1.COLOR_CODE (+) = GMTSTYL1.COLOR_CODE" & vbCrLf _
            & " group by GMTSTYL1.DGC_CODE, GMTSTYL1.VEND_CODE, GMTSTYL1.STYLE_CODE, GMTSTYL1.COLOR_CODE"

        Fill_Records("SPTSPXPR", "", True, ASCMAIN1.sql)


        If VEND_CODE = "" And DGC_CODE = "" Then
            splSPTXPSP1.Panel1Collapsed = True
            Fill_Records("SPTSPXPR")
            Sort_grdColumns(grdSPTSPXPR, "DGC_CODE")
        Else
            splSPTXPSP1.Panel1Collapsed = False
            If DGC_CODE <> "" Then
                ASCMAIN1.sql = "Select VEND_CODE CODE_VALUE, VEND_NAME DESC_VALUE" _
                    & " from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from " & GMTSTYL1 & ")"
                grdSPTSPXPC.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption = "Vendor"
                grdSPTSPXPC.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = "Name"
                grdSPTSPXPC.Text = "Vendors with Styles in DGC " & DGC_CODE & IIf(SEASON_CODE = "", "", ", Season " & SEASON_CODE)
                ' Create_Relation("SPTSPXPC", "SPTSPXPR", "CODE_VALUE", "DGC_CODE")
            ElseIf VEND_CODE <> "" Then
                ASCMAIN1.sql = "Select DGC_CODE CODE_VALUE, DGC_DESC DESC_VALUE" _
                    & " from GMTDGCM1 where DGC_CODE in (Select Distinct DGC_CODE from " & GMTSTYL1 & ")"
                grdSPTSPXPC.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption = "DGC"
                grdSPTSPXPC.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = "Description"
                grdSPTSPXPC.Text = "DGCs with Styles in Vendor " & VEND_CODE & IIf(SEASON_CODE = "", "", ", Season " & SEASON_CODE)
                ' Create_Relation("SPTSPXPC", "SPTSPXPR", "CODE_VALUE", "DGC_CODE")
            End If
            Fill_Records("SPTSPXPC", "", True, ASCMAIN1.sql)
            Sort_grdColumns(grdSPTSPXPC, "CODE_VALUE")

            For Each rowSPTSPXPC As DataRow In dst.Tables("SPTSPXPC").Select("")
                Dim sql As String = ""
                If DGC_CODE <> "" Then
                    sql = "VEND_CODE = '" & rowSPTSPXPC.Item("CODE_VALUE") & "'"
                Else
                    sql = "DGC_CODE = '" & rowSPTSPXPC.Item("CODE_VALUE") & "'"
                End If
                rowSPTSPXPC.Item("RECORDS") = dst.Tables("SPTSPXPR").Select(sql).Length
            Next
        End If



        EnforceConstraints(True)

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
        ' Dependent_Updates(-1, ORDR_NO)
        For Each TABLE_NAME As String In New String() _
            {"GMTSCOL1", "GMTSCOL2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        'ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where STYLE_COLLECTION_CODE = '" & STYLE_COLLECTION_CODE & "'"
        'ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        BeginTrans()

        ASCMAIN1.sql = "Delete from SPTSPXP1 where DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE,PROMOTION_WEEK in " _
            & " (Select DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE,PROMOTION_WEEK from " & SPTSPXP1 & ")"
        ASCDATA1.ExecuteSQL()

        Update_Record_TDA("SPTSPXP1")

        ASCMAIN1.sql = "Insert into SPTSPXP1 Select * from " & SPTSPXP1
        ASCDATA1.ExecuteSQL()

        CommitTrans("Update Complete")
    End Sub


    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "VEND_CODE"
                Dim DEPT_CODE As String = Absx1.txtFor("DEPT_CODE").Text
                If DEPT_CODE <> "" Then sql_where = " and APTVEND1.VEND_CODE in (Select Distinct VEND_CODE from GMTSTYL1 where DGC_CODE like '" & DEPT_CODE & "%')"

            Case "DGC_CODE"
                Dim DEPT_CODE As String = Absx1.txtFor("DEPT_CODE").Text
                If DEPT_CODE <> "" Then sql_where = " and GMTDGCM1.DGC_CODE in (Select Distinct DGC_CODE from GMTSTYL1 where DGC_CODE like '" & DEPT_CODE & "%')"
        End Select
    End Sub



#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGMTSCOLX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSPTSPXPR, "BBBB", "Set Selected", "Set All", "Clear Selected", "Clear All")
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

        Select Case e.SourceControl.Name
            Case "grdSPTSPXPR"
                If EntryMode <> "E" Then
                    e.Cancel = True
                Else
                    PROMOTION_CODE = grdSPTPROM1.ActiveRow.Cells("PROMOTION_CODE").Value & ""
                    tlb_btn = DirectCast(tlb.Tools("Set Selected"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Caption = "Set Selected to " & PROMOTION_CODE
                    tlb_btn = DirectCast(tlb.Tools("Set All"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Caption = "Set All to " & PROMOTION_CODE
                End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        End If

        Select Case e.SourceControl.Name
            'Case "grdSOTINVHX"
            '    e.Tool.ToolbarsManager.Tools("Sales Order Inquiry").SharedProps.Visible = True
        End Select
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "Set All", "Clear All", "Set Selected", "Clear Selected"
                Dim PROMO_CODE As String = ""
                If e.Tool.Key.StartsWith("Set") Then
                    PROMO_CODE = PROMOTION_CODE
                End If
                For Each grow As UltraWinGrid.UltraGridRow In grdSPTSPXPR.Rows
                    If e.Tool.Key.EndsWith("All") Or grow.Selected Then
                        For Each iREGION_CODE As Integer In iREGION_CODEs
                            grow.Cells("PROMO_CODE_" & CStr(iREGION_CODE)).Value = PROMO_CODE
                        Next
                    End If
                    grow.Update()
                Next

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "SKU_NUMBER"
                If e.KeyCode = Windows.Forms.Keys.Enter Then

                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SKU_NUMBER"

        End Select
    End Sub
#End Region

    Sub Load_GMTSCOLX(initialize As Boolean)

    End Sub

    Private Sub grdSPTSPXPC_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdSPTSPXPC.AfterRowActivate
        Setup_grdSPTSPXPR()
    End Sub

    Sub Setup_grdSPTSPXPR()
        If grdSPTSPXPC.ActiveRow Is Nothing Then
            grdSPTSPXPR.Visible = False
        Else
            grdSPTSPXPR.Visible = True
            Dim CODE_VALUE As String = grdSPTSPXPC.ActiveRow.Cells("CODE_VALUE").Value & ""

            Dim dvw As DataView = DirectCast(grdSPTSPXPR.DataSource, DataTable).DefaultView
            If VEND_CODE <> "" Then
                dvw.RowFilter = "DGC_CODE = '" & CODE_VALUE & "'"
                grdSPTSPXPR.Text = "Styles for DGC " & CODE_VALUE
            Else
                dvw.RowFilter = "VEND_CODE = '" & CODE_VALUE & "'"
                grdSPTSPXPR.Text = "Styles for Vendor " & CODE_VALUE
            End If
        End If
    End Sub

    Private Sub grdSPTPROM1_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSPTPROM1.InitializeRow
        Dim PROMOTION_CODE As String = e.Row.Cells("PROMOTION_CODE").Value
        Dim PROMOTION_TYPE As String = e.Row.Cells("PROMOTION_TYPE").Value
        Dim PROMOTION_AMT As Decimal = Val(e.Row.Cells("PROMOTION_AMT").Value & "")
        Dim DIVISOR_TYPEB As Integer = Val(e.Row.Cells("DIVISOR_TYPEB").Value & "")
        Dim PROMOTION_DESC As String = ""

        Select Case PROMOTION_TYPE
            Case "%"
                PROMOTION_DESC = CStr(PROMOTION_AMT) & "% Off"
            Case "$"
                PROMOTION_DESC = Format(PROMOTION_AMT, "$##.00") & " Off"
            Case "P"
                PROMOTION_DESC = Format(PROMOTION_AMT, "$##.00")
            Case "D"
                PROMOTION_DESC = "Buy " & CStr(DIVISOR_TYPEB) & " Get 1 " & CStr(PROMOTION_AMT) & " for " & PROMOTION_CODE
            Case "B"
                PROMOTION_DESC = "Buy " & CStr(DIVISOR_TYPEB) & " Get 1 " & IIf(PROMOTION_AMT = 0, " Free", CStr(PROMOTION_AMT) & "% Off")
            Case "X"
                PROMOTION_DESC = "0"
            Case "Z"
                PROMOTION_DESC = "Remove"
        End Select

        e.Row.Cells("PROMOTION_DESC").Value = PROMOTION_DESC
    End Sub

    Private Sub grdGMTREGN1_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGMTREGN1.AfterRowUpdate
        Dim SEL As String = e.Row.Cells("SEL").Value & ""
        Dim REGION_CODE As String = e.Row.Cells("REGION_CODE").Value & ""
        Dim iREGION_CODE As Integer = Val(e.Row.Cells("I").Value & "" & "")
        grdSPTSPXPR.DisplayLayout.Bands(0).Columns("PROMO_CODE_" & CStr(iREGION_CODE)).Hidden = (SEL <> "1")
        If SEL = "1" Then
            iREGION_CODEs.Add(iREGION_CODE)
        Else
            iREGION_CODEs.Remove(iREGION_CODE)
        End If
    End Sub

    Private Sub optStyles_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStyles.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_optStyles_Dependencies()
    End Sub

    Sub Set_optStyles_Dependencies()
        Absx1.txtFor("VEND_CODE").Visible = (optStyles.Value = "VEND_CODE")
        Absx1.txtFor("VEND_NAME").Visible = (optStyles.Value = "VEND_CODE")
        Absx1.txtFor("DGC_CODE").Visible = (optStyles.Value = "DGC_CODE")
        Absx1.txtFor("DGC_DESC").Visible = (optStyles.Value = "DGC_CODE")
        Absx1.txtFor("DG_CODE").Visible = (optStyles.Value = "DG_CODE")
        Absx1.txtFor("DG_DESC").Visible = (optStyles.Value = "DG_CODE")
        Absx1.txtFor("SSR_XNO").Visible = (optStyles.Value = "SSR_XNO")
        Absx1.txtFor("SSR_XNO_DESC").Visible = (optStyles.Value = "SSR_XNO")

        Select Case optStyles.Value
            Case "VEND_CODE"
                lblCode.Text = "Vendor"
                lblDesc.Text = "Name"
            Case "DGC_CODE"
                lblCode.Text = "DGC"
                lblDesc.Text = "Description"
            Case "DG_CODE"
                lblCode.Text = "DG"
                lblDesc.Text = "Description"
            Case "SSR_XNO"
                lblCode.Text = "XNO"
                lblDesc.Text = "Report Sub-Title"
        End Select
    End Sub
End Class