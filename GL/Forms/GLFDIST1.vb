Public Class GLFDIST1

    Dim GLTDISTX As String
    Dim sqlGLTDIST2 As String
    Dim DIST_CODE As String
    Dim GLTDIST2 As String

    'Dim GYP As String
    'Dim inquiry As Boolean
    'Dim ACCT_CODE As String
    Dim ACCT_SEG_ID As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")

        Load_GLTDISTX(True)

        With dst
            ASCMAIN1.sql = "Select * from " & GLTDISTX
            Create_TDA(.Tables.Add, "GLTDISTX", "**", 0, False, "", 1)

            Create_TDA(.Tables.Add, "GLTDIST1", "*")

            Create_TDA(.Tables.Add, "GLTDIST2", "*", 1)
            .Tables("GLTDIST2").Columns("DIST_PCT").DefaultValue = 0

            Create_Relation("GLTDIST1", "GLTDIST2", "DIST_CODE")
            .Tables("GLTDIST1").Columns.Add("DIST_PCT_TOTAL", GetType(System.Decimal), "SUM(CHILD(GLTDIST1_GLTDIST2).DIST_PCT)")
            .Tables("GLTDIST2").Columns.Add("DIST_PCT_TOTAL", GetType(System.Decimal), "PARENT(GLTDIST1_GLTDIST2).DIST_PCT_TOTAL")
            .Tables("GLTDIST2").Columns.Add("DIST_PCT_CALC", GetType(System.Decimal), "IIF(DIST_PCT_TOTAL=0,0,DIST_PCT * 100 / DIST_PCT_TOTAL")

            ASCMAIN1.sql = "Select GLTDIST3.*, GLTACCT1.ACCT_DESC" & vbCrLf _
             & " from GLTDIST3, GLTACCT1 where GLTDIST3.DIST_CODE = :PARM1" & vbCrLf _
             & " and GLTACCT1.ACCT_CODE = GLTDIST3.ACCT_CODE"
            Create_TDA(.Tables.Add, "GLTDIST3", "**", 0, False, "V", 2)

            Create_TDA(.Tables.Add, "GLTDIST4", "*", 1)

            '  Create_TDA(.Tables.Add, "GLTSEGM1", "*", 0)
        End With

        '  Fill_Records("GLTSEGM1")

        Create_Relation("GLTDISTX", "GLTDIST2", "DIST_CODE")

        grdGLTDISTX.DataSource = dst.Tables("GLTDISTX")
        grdGLTDIST2.DataSource = dst.Tables("GLTDIST2")
        grdGLTDIST4.DataSource = dst.Tables("GLTDIST4")

        Create_Summary(grdGLTDISTX, "DIST_CODE", "Count")

        Create_Summary(grdGLTDIST2, "DIST_CODE", "Count")
        Create_Summary(grdGLTDIST2, New String() {"BASIS_AMT", "BASIS_PCT", "DIST_PCT"})

        Create_Summary(grdGLTDIST4, "DIST_CODE", "Count")
        Create_Summary(grdGLTDIST4, New String() {"DIST_PCT"})

        With grdGLTDIST2.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                If New String() {"QTY_ORD", "QTY_OPN_1", "QTY_OPN_2", "QTY_REC"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightPink
                    gcol.Width = 50
                End If
                If New String() {"OH_STR", "OH_WHS"}.Contains(gcol.Key) Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    gcol.Width = 50
                End If
            Next
        End With

        Show_Filter(grdGLTDISTX, True)

        Dim VL As New ValueList
        For i As Integer = 2 To 4
            Dim zz As String = "SEG" & CStr(i) & "_CODE"
            Dim z As String = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & ""
            If z <> "" Then
                Dim VLI As New ValueListItem(zz, z)
                VL.ValueListItems.Add(VLI)
                grdGLTDIST4.DISPLAYLAYOUT.BANDS(0).COLUMNS(zz).HEADER.CAPTION = z
            Else
                grdGLTDIST4.DISPLAYLAYOUT.BANDS(0).COLUMNS(zz).HIDDEN = True
            End If
        Next i
        optSegment.ValueList = VL
 
        'ASCMAIN1.Add_Value_List(grdGMTCGMAX, "SEASON_SEQ_NO", "SELECT SEASON_SEQ_NO, SEASON_CODE FROM GMTSEAS1 WHERE SEASON_ACTIVE = '1'")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("DIST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Supply a Valid Code"
                Else
                    DIST_CODE = Absx1.txtFor("DIST_CODE").Text
                    Dim rowGLTDIST1 As DataRow = LookUp("GLTDIST1", DIST_CODE)
                    If rowGLTDIST1 IsNot Nothing Then
                        EMsg &= vbCr & "Code " & DIST_CODE & " Already Exists"
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("GLTDIST1", DIST_CODE) Then
                        Exit Sub
                    End If
                End If

            Case "View"

            Case "Edit", "View"

                If Absx1.txtFor("DIST_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must Supply a Valid Code"
                Else
                    DIST_CODE = Absx1.txtFor("DIST_CODE").Text
                    Dim rowGLTDIST1 As DataRow = LookUp("GLTDIST1", DIST_CODE)
                    If rowGLTDIST1 Is Nothing Then
                        EMsg &= vbCr & "Code " & DIST_CODE & " Does Not Exist"
                    End If
                End If

                If EMsg = "" Then
                    If eItemKey = "Edit" Then
                        If Not ASCMAIN1.Logical_Lock("GLTDIST1", DIST_CODE) Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Update"

                If Absx1.txtFor("DIST_DESC").Text = "" Then
                    EMsg &= vbCr & "A Description is Mandatory"
                Else
                    If optDistribution.Value = "S" Then
                        Normalize(False, "")
                    Else
                        Normalize(False, "2")
                        Normalize(False, "3")
                        Normalize(False, "4")
                    End If

                    If optDistribution.Value = "S" Then
                        For Each row As DataRow In ASCDATA1.SelectDistinct("GLTDIST4", New String() {"SEG2_CODE", "SEG3_CODE", "SEG4_CODE"}).Select("")
                            Dim SEG2_CODE As String = row.Item("SEG2_CODE")
                            Dim SEG3_CODE As String = row.Item("SEG3_CODE")
                            Dim SEG4_CODE As String = row.Item("SEG4_CODE")
                            Dim sqlw As String = "SEG2_CODE = '" & SEG2_CODE & "' and SEG3_CODE = '" & SEG3_CODE & "' and SEG4_CODE = '" & SEG4_CODE & "'"
                            Dim C As Integer = Val(dst.Tables("GLTDIST4").Compute("Count(SEG2_CODE)", sqlw) & "")
                            If C > 1 Then
                                EMsg = EMsg & vbCr & "Duplicate Entry for " & SEG2_CODE & ":" & SEG3_CODE & ":" & SEG4_CODE
                            End If
                        Next
                    End If
                End If

            Case "Delete"

                If MsgBox("Do you really want to Delete the entire Distribution?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

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

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "Normalize to 100%"
                Normalize(True, ACCT_SEG_ID)

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
                    .Items("New").Settings.Enabled = not_iScreenMode
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

                    .Items("Excel").Visible = ScreenMode
                    .Items("New").Visible = Not InquiryMode
                    .Items("Edit").Visible = Not InquiryMode

                    .Items("Print").Visible = ScreenMode
                    .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                    ' .Items("Print").Visible = (EntryMode = "V" And ScreenMode) ' False ' ScreenMode
                    .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                    .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                    .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                End With

                .Groups("Show by Store for").Visible = ScreenMode
                .Groups("Add SKU").Visible = ScreenMode And (EntryMode <> "V")
                '.Groups("Style Group").Visible = ScreenMode

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(Absx1.txtFor("DIST_DESC"), EntryMode = "V")
        Set_Read_Only_for_ctl(Absx1.txtFor("VEND_BUYER_CODE"), EntryMode = "V")

        With grdGLTDIST2.DisplayLayout.Override
            If EntryMode = "E" Or EntryMode = "N" Then
                .AllowDelete = DefaultableBoolean.True
            Else
                .AllowDelete = DefaultableBoolean.False
            End If
        End With

        grdGLTDISTX.Visible = Not ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLTDIST1", "GLTDIST2", "GMTCGMAX", "GMTCGMA0", "GLTDISTX", "POTORDRX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Load_GLTDISTX(False)
    End Sub

    Sub Print_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Report")

        Print_Report_Begin()
        CR_params.Add("SUBT", DIST_CODE & "-" & Absx1.txtFor("DIST_DESC").Text)
        Generate_Report("GMRSCOL1", "Style Collection Members")

        Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Record()
        ASCMAIN1.Progress("Now Loading Data")
        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        If EntryMode = "N" Then

            Dim rowGLTDIST1 As DataRow = dst.Tables("GLTDIST1").NewRow
            DIST_CODE = ASCMAIN1.Next_Control_No("GLTDIST1.DIST_CODE")
            rowGLTDIST1.Item("DIST_CODE") = DIST_CODE
            rowGLTDIST1.Item("DIST_DESC") = HFs("DIST_DESC")
            rowGLTDIST1.Item("INIT_DATE") = DATETIME_STAMP
            rowGLTDIST1.Item("INIT_OPER") = ASCMAIN1.USER_ID
            dst.Tables("GLTDIST1").Rows.Add(rowGLTDIST1)
 
 
        Else
            Fill_Records("GLTDIST1", DIST_CODE)
        End If

        'Sql = "Select '" & DIST_CODE & "' DIST_CODE, GLTSEGM1.ACCT_SEG_ID, GLTSEGM1.ACCT_SEG_CODE, "
        'Sql = Sql & " GLTDIST2.BASIS_AMT, GLTDIST2.BASIS_PCT, GLTDIST2.DIST_PCT, GLTSEGM1.ACCT_SEG_DESC"
        'Sql = Sql & " from GLTDIST2, GLTSEGM1 where GLTDIST2.DIST_CODE (+) = '" & DIST_CODE & "'"
        'Sql = Sql & " and GLTDIST2.ACCT_SEG_CODE (+) = GLTSEGM1.ACCT_SEG_CODE"

        Fill_Records("GLTDIST2", DIST_CODE)
        Sort_grdColumns(grdGLTDIST2, "ACCT_SEG_CODE")

        Fill_Records("GLTDIST3", DIST_CODE)
        Fill_Records("GLTDIST4", DIST_CODE)

        EnforceConstraints(True)

        Set_DISTX()
        Setup_Method()
        Set_Segment()

        grdGLTDIST2.Text = "Distribution " & DIST_CODE & ":" & Absx1.txtFor("DIST_DESC").Text

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
            {"GLTDIST1", "GLTDIST2", "GLTDIST3", "GLTDIST4"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where DIST_CODE = '" & DIST_CODE & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        BeginTrans()

        Delete_Records()

        'If EntryMode = "N" Then
        '    Dim rowGLTDIST1 As DataRow = dst.Tables("GLTDIST1").NewRow
        '    rowGLTDIST1.Item("DIST_CODE") = DIST_CODE
        'End If

        ASCDATA1.DeleteRows("GLTDIST2", "DIST_PCT = 0")
        ASCDATA1.DeleteRows("GLTDIST4", "DIST_PCT = 0")

        INIT_LAST("GLTDIST1", True)

        Update_Record_TDA("GLTDIST1")
        Update_Record_TDA("GLTDIST2")
        Update_Record_TDA("GLTDIST3")
        Update_Record_TDA("GLTDIST4")

        CommitTrans("Update Complete")
    End Sub
#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTDISTX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdGLTDIST4, "B", "Get Accounts")
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
            'Case "grdGLTDIST2"
            '    e.Tool.ToolbarsManager.Tools("Add Styles from Style Master").SharedProps.Visible = (EntryMode <> "V")
            '    e.Tool.ToolbarsManager.Tools("Add Styles from Open POs").SharedProps.Visible = (EntryMode <> "V")
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
            Case "Add Styles from Style Master"

                'Dim sql_where As String = ""
                'ASCMAIN1.CodeSelector.SQL = ASCMAIN1.CodeSelector.Get_SQL("LONG_SKU", , sql_where)

                'If ASCMAIN1.CodeSelector.SQL <> "" Then
                '    ASCMAIN1.CodeSelector.MultipleSelections = True
                '    Dim F As New ASFCODE1
                '    F.ShowDialog()
                '    F.Dispose()
                '    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                '        Me.Cursor = Cursors.WaitCursor
                '        ASCMAIN1.Progress("Now Loading Styles")

                '        grdGLTDIST2.Visible = False
                '        For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows

                '            Dim DGC_CODE As String = row.Item("DGC_CODE")
                '            Dim VEND_CODE As String = row.Item("VEND_CODE")
                '            Dim STYLE_CODE As String = row.Item("STYLE_CODE")
                '            Dim COLOR_CODE As String = row.Item("COLOR_CODE")

                '            Dim rowGLTDIST2 As DataRow = dst.Tables("GLTDIST2").Rows.Find(New String() {DIST_CODE, DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
                '            If rowGLTDIST2 IsNot Nothing Then
                '                MsgBox("Style " & DGC_CODE & "-" & VEND_CODE & "-" & STYLE_CODE & "-" & COLOR_CODE & " is already in Style Collection", MsgBoxStyle.OkOnly, "Cannot Add Style")
                '            Else
                '                Add_Style(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
                '            End If
                '        Next
                '        grdGLTDIST2.Visible = True
                '        Me.Cursor = Cursors.Default
                '        ASCMAIN1.Progress("")
                '    End If
                'End If
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Text
                Context_Launch("Select Customer", CUST_CODE, e.Tool.Key, "ARFCINQ1")

            Case "Get Accounts"
                Get_Accounts()
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)

            Case "SKU_NUMBER"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Dim SKU_NUMBER As String = Absx1.txtFor("SKU_NUMBER").Text
                    If SKU_NUMBER <> "" Then
                        Add_SKU(SKU_NUMBER)

                    End If
                End If
        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SKU_NUMBER"
                Dim SKU_NUMBER As String = Absx1.txtFor("SKU_NUMBER").Text
                If SKU_NUMBER <> "" Then
                    Add_SKU(SKU_NUMBER)
                End If
        End Select
    End Sub

    Public Overrides Sub num_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs)
        MyBase.num_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case ""
                If e.KeyCode = Keys.Enter Then
                End If
        End Select
    End Sub
#End Region

    Sub Load_GLTDISTX(initialize As Boolean)
        ASCMAIN1.sql = "Select GLTDIST1.DIST_CODE, GLTDIST1.DIST_DESC, Count (*) LSKUS" & vbCrLf _
            & ",GLTDIST1.INIT_DATE,GLTDIST1.INIT_OPER,GLTDIST1.LAST_DATE,GLTDIST1.LAST_OPER,GLTDIST1.VEND_BUYER_CODE" & vbCrLf _
            & " from GLTDIST1,GLTDIST2 where GLTDIST1.DIST_CODE = GLTDIST2.DIST_CODE" & vbCrLf _
            & " group by GLTDIST1.DIST_CODE, GLTDIST1.DIST_DESC" & vbCrLf _
            & ",GLTDIST1.INIT_DATE,GLTDIST1.INIT_OPER,GLTDIST1.LAST_DATE,GLTDIST1.LAST_OPER,GLTDIST1.VEND_BUYER_CODE"
        If initialize Then
            GLTDISTX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & GLTDISTX & " Add Primary Key (DIST_CODE)")

            ASCMAIN1.sql = "Select * from GLTDIST2 where ROWNUM < 1"
            GLTDIST2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & GLTDIST2 & " Add Primary Key (DIST_CODE,DGC_CODE,VEND_CODE,STYLE_CODE,COLOR_CODE)")

        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & GLTDISTX)
            ASCDATA1.ExecuteSQL("Insert into " & GLTDISTX & " " & ASCMAIN1.sql)

            EnforceConstraints(False)
            Fill_Records("GLTDISTX")

            ASCMAIN1.sql = Replace(sqlGLTDIST2, GLTDIST2 & " GLTDIST2", "GLTDIST2")
            Fill_Records("GLTDIST2", "", , ASCMAIN1.sql)



            EnforceConstraints(True)
            Sort_grdColumns(grdGLTDISTX, "DIST_DESC")
        End If
    End Sub

    Private Sub grdGLTDISTX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTDISTX.DoubleClickRow
        DIST_CODE = e.Row.Cells("DIST_CODE").Value
        Click_Command("View")
    End Sub

    Sub Add_SKU(SKU_NUMBER As String)
        Dim rowGMTSKUF1 As DataRow = LookUp("GMTSKUF1", SKU_NUMBER)
        If rowGMTSKUF1 Is Nothing Then
            MsgBox("Invalid Value Specified for SKU (" & SKU_NUMBER & ")", MsgBoxStyle.OkOnly, "Cannot Add SKU")
        Else
            Dim DGC_CODE As String = rowGMTSKUF1.Item("DGC_CODE")
            Dim VEND_CODE As String = rowGMTSKUF1.Item("VEND_CODE")
            Dim STYLE_CODE As String = rowGMTSKUF1.Item("STYLE_CODE")
            Dim COLOR_CODE As String = rowGMTSKUF1.Item("COLOR_CODE")

            Dim rowGLTDIST2 As DataRow = dst.Tables("GLTDIST2").Rows.Find(New String() {DIST_CODE, DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
            If rowGLTDIST2 IsNot Nothing Then
                MsgBox("SKU " & SKU_NUMBER & " is already in Style Collection", MsgBoxStyle.OkOnly, "Cannot Add SKU")
            Else
                Add_Style(DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE)
            End If
            Absx1.txtFor("SKU_NUMBER").Text = ""
        End If
        Application.DoEvents()
        Absx1.txtFor("SKU_NUMBER").Focus()
    End Sub

    Sub Add_Style(DGC_CODE As String, VEND_CODE As String, STYLE_CODE As String, COLOR_CODE As String)
        Dim rowGLTDIST2 As DataRow = dst.Tables("GLTDIST2").NewRow

        Dim rowGMTSTYL1 As DataRow = LookUp("GMTSTYL1", New String() {DGC_CODE, VEND_CODE, STYLE_CODE, COLOR_CODE})
        If rowGMTSTYL1 IsNot Nothing Then
            rowGLTDIST2 = dst.Tables("GLTDIST2").NewRow
            For Each dcol As DataColumn In dst.Tables("GLTDIST2").Columns
                If dcol.ColumnName = "DIST_CODE" Then
                    rowGLTDIST2.Item("DIST_CODE") = DIST_CODE
                Else
                    If New String() {"OH_STR", "OH_WHS", "QTY_ORD", "QTY_OPN_1", "QTY_OPN_2", "QTY_REC", "INIT_DATE", "INIT_OPER", "LAST_DATE", "LAST_OPER"}.Contains(dcol.ColumnName) Then
                    Else
                        rowGLTDIST2.Item(dcol.ColumnName) = rowGMTSTYL1.Item(dcol.ColumnName)
                    End If
                End If
            Next
        Else
            ASCMAIN1.sql = "Select POTORDR2.DGC_CODE, POTORDR1.VEND_CODE, POTORDR2.STYLE_CODE, POTORDR2.COLOR_CODE" & vbCrLf _
                & ", POTORDR2.RETAIL_PRICE CUR_RETAIL, POTORDR2.COST LST_UNIT_COST, GMTSEAS1.SEASON_CODE" & vbCrLf _
                & ", POTORDR2.STYLE_DESC DESCRIPTION, POTORDR2.SCALE_CODE" & vbCrLf _
                & " from POTORDR1,POTORDR2,GMTSEAS1 where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
                & "   and GMTSEAS1.SEASON_SEQ_NO = POTORDR1.SEASON_SEQ_NO" & vbCrLf _
                & "   and POTORDR2.DGC_CODE = :PARM1 and POTORDR2.STYLE_CODE = :PARM2" & vbCrLf _
                & "   and POTORDR2.COLOR_CODE = :PARM3 and POTORDR1.VEND_CODE = :PARM4"

            Dim row() As DataRow = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVVV", New Object() {DGC_CODE, STYLE_CODE, COLOR_CODE, VEND_CODE}).Select("")

            If row.Length > 0 Then
                rowGLTDIST2 = dst.Tables("GLTDIST2").NewRow
                For Each dcol As DataColumn In dst.Tables("GLTDIST2").Columns
                    If dcol.ColumnName = "DIST_CODE" Then
                        rowGLTDIST2.Item("DIST_CODE") = DIST_CODE
                    Else
                        If row(0).Table.Columns.Contains(dcol.ColumnName) Then
                            rowGLTDIST2.Item(dcol.ColumnName) = row(0).Item(dcol.ColumnName)
                        End If
                    End If
                Next
            End If
        End If

        If rowGLTDIST2 IsNot Nothing Then
            dst.Tables("GLTDIST2").Rows.Add(rowGLTDIST2)
            Update_Record_TDA("GLTDIST2")

        End If

    End Sub

    Private Sub grdGLTDIST2_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdGLTDIST2.AfterRowActivate

    End Sub

    Private Sub grdGLTDIST2_AfterRowsDeleted(sender As Object, e As System.EventArgs) Handles grdGLTDIST2.AfterRowsDeleted

    End Sub

    Private Sub grdGLTDIST2_BeforeRowsDeleted(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdGLTDIST2.BeforeRowsDeleted

    End Sub

    Private Sub grdGLTDIST2_InitializeRow(sender As Object, e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTDIST2.InitializeRow

    End Sub






    Private Sub chkShowNonZeroOnly_Click()
        Set_DISTX()
    End Sub

    Sub Set_DISTX()
        Dim sql As String = ""
        If optDistribution.Value = "S" Then
            Dim dvw As DataView = DirectCast(grdGLTDIST4.DataSource, DataTable).DefaultView
            If chkShowNonZeroOnly.checked Then
                sql = "DIST_PCT <> 0"
            End If
            dvw.RowFilter = sql
        Else
            Dim dvw As DataView = DirectCast(grdGLTDIST2.DataSource, DataTable).DefaultView

            sql = "ACCT_SEG_ID = '" & ACCT_SEG_ID & "'"
            If chkShowNonZeroOnly.checked Then
                sql &= " and DIST_PCT <> 0"
            End If
            dvw.RowFilter = sql
        End If
    End Sub

    Sub Get_Accounts()
        Add_Codes(grdGLTDIST2, "GLTDIST3", "ACCT_CODE", "GL Accounts")
    End Sub

    Sub Normalize(show_error_message As Boolean, ACCT_SEG_ID As String)
        Dim TABLE_NAME As String
        If optDistribution.Value = "S" Then
            TABLE_NAME = "GLWDIST4"
        Else
            TABLE_NAME = "GLWDIST2"
        End If

        Dim sqlx As String
        sqlx = "DIST_PCT <> 0 "
        If optDistribution.Value = "M" Then
            sqlx = sqlx & " and ACCT_SEG_ID = '" & ACCT_SEG_ID & "'"
        End If

        For Each row As DataRow In dst.Tables(TABLE_NAME).Select(sqlx)
            row.Item("DIST_PCT") = row.Item("DIST_PCT_CALC")
        Next

        'T = Val(dst.Tables(TABLE_NAME).Compute("SUM(DIST_PCT)", sqlx) & "") - 100
        'Dim row As DataRow = dst.Tables(TABLE_NAME).Select(sqlx, "DIST_PCT" & IIf(T > 0, " DESC", ""))(0)
        'row.Item("DIST_PCT") = Val(row.Item("DIST_PCT") & "") - T
    End Sub

    Sub Zero_Basis()
        For Each rowGLTDIST2 As DataRow In dst.Tables("GLTDIST2").Select("ACCT_SEG_ID = '" & ACCT_SEG_ID & "'")
            rowGLTDIST2.Item("BASIS_AMT") = 0
            rowGLTDIST2.Item("BASIS_PCT") = 0
            rowGLTDIST2.Item("DIST_PCT") = 0
        Next
    End Sub
     
    Sub Set_Segment()
        ACCT_SEG_ID = optSegment.Value
        grdGLTDIST2.DisplayLayout.Bands(0).Columns("ACCT_SEG_CODE").Header.Caption = optSegment.Text
        Set_DISTX()
    End Sub

#Region "grdGLTDIST2"
    Private Sub grdGLTDIST2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ACCT_SEG_CODE"
                Dim ACCT_SEG_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdGLTDIST2, "GLTSEGM1", "ACCT_SEG_CODE", "ACCT_SEG_DESC")
        End Select
    End Sub

    Private Sub grdGLTDIST2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTDIST2.BeforeRowUpdate
        With grdGLTDIST2
            If Not e.Cancel Then
                If e.Row.Cells("DIST_CODE").Text = "" Then
                    .ActiveRow.Cells("DIST_CODE").Value = Absx1.CtlFor("DIST_CODE").Text
                End If
            End If
        End With
    End Sub

    Private Sub grdGLTDIST2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST2.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdGLTDIST2, sql_where, sql_where <> "")
    End Sub
#End Region

#Region "grdGLTDIST3"
    Private Sub grdGLTDIST3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST3.AfterCellUpdate

        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""
                grdCodeDesc(grdGLTDIST4, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
        End Select
    End Sub

    Private Sub grdGLTDIST3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTDIST3.BeforeRowUpdate
        With grdGLTDIST3
            If Not e.Cancel Then
                If e.Row.Cells("DIST_CODE").Text = "" Then
                    .ActiveRow.Cells("DIST_CODE").Value = Absx1.CtlFor("DIST_CODE").Text
                End If
            End If
        End With
    End Sub

    Private Sub grdGLTDIST3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST3.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdGLTDIST3, sql_where, sql_where <> "")
    End Sub
#End Region

#Region "grdGLTDIST4"
    Private Sub grdGLTDIST4_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST4.AfterCellUpdate

        Select Case e.Cell.Column.Key

        End Select
    End Sub

    Private Sub grdGLTDIST4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTDIST4.BeforeRowUpdate
        With grdGLTDIST4
            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If Not e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    If e.Row.Cells(COLUMN_NAME).Text = "" Then
                        e.Cancel = True
                    Else
                        LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                        If cdr Is Nothing Then
                            MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                            e.Cancel = True
                        Else
                            If cdr.Item("ACCT_SEG_NO_GL") & "" = "1" Then
                                e.Cancel = True
                            End If
                        End If
                    End If
                End If
            Next


            If Not e.Cancel Then
                If e.Row.Cells("DIST_CODE").Text = "" Then
                    .ActiveRow.Cells("DIST_CODE").Value = Absx1.CtlFor("DIST_CODE").Text
                    .ActiveRow.Cells("DIST_LNO").Value = Val(dst.Tables("GLTDIST4").Compute("Max(DIST_LNO)", "") & "") + 1
                End If
            End If
        End With

    End Sub

    Private Sub grdGLTDIST4_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTDIST4.ClickCellButton
        Dim sql_where As String = ""
        Dim z As String = e.Cell.Column.Key
        If z = "SEG2_CODE" Or z = "SEG3_CODE" Or z = "SEG4_CODE" Then
            Dim i As Integer = Val(Mid$(z, 4, 1))
            sql_where = "ACCT_SEG_ID = '" & CStr(i) & "' and (ACCT_SEG_NO_GL <> '1' or ACCT_SEG_NO_GL is Null)"

        End If
        grdClickCellButton(grdGLTDIST4, sql_where, sql_where <> "")
    End Sub
#End Region

    Sub Setup_Method()
        Dim tf As Boolean
        If optDistribution.Value = "S" Then
            tf = False
            splDistributionMethod.Panel2Collapsed = False
            splDistributionMethod.Panel1Collapsed = True
            splDistribution.Panel2Collapsed = True
        Else
            tf = True
            splDistributionMethod.Panel1Collapsed = False
            splDistributionMethod.Panel2Collapsed = True
            splDistribution.Panel2Collapsed = False
        End If

        optSegment.Visible = tf
        Absx1.txtFor("OPS_YYYYPP_1").Visible = tf
        Absx1.txtFor("OPS_YYYYPP_2").Visible = tf

        If optDistribution.Value = "S" Then
        Else
            Set_Segment()
        End If

    End Sub

    Private Sub Get_Basis()

        If Absx1.txtFor("OPS_YYYYPP_1").Text = "" Or Absx1.txtFor("OPS_YYYYPP_2").Text = "" Then
            MsgBox("You Must Specify Starting and Ending Periods", MsgBoxStyle.OkOnly, "Cannot Get Actuals to Use in Basis")
            Exit Sub
        End If
        If grdGLTDIST3.Rows.Count = 0 Then
            MsgBox("You Must Specify Accounts", MsgBoxStyle.OkOnly, "Cannot Get Actuals to Use in Basis")
            Exit Sub
        End If

        Dim OPS_YYYYPP_1 As String = Absx1.txtFor("OPS_YYYYPP_1").Text
        Dim OPS_YYYYPP_2 As String = Absx1.txtFor("OPS_YYYYPP_2").Text

        If OPS_YYYYPP_1 > OPS_YYYYPP_2 Then
            MsgBox("Starting Period must not be later than Ending Period", MsgBoxStyle.OkOnly, "Cannot Get Actuals to Use in Basis")
            Exit Sub
        End If

        Dim ACCT_CODEs As String = ""
        For Each rowGLTDIST3 As DataRow In dst.Tables("GLTDIST3").Select("", "ACCT_CODE")
            ACCT_CODEs = ACCT_CODEs & ",'" & rowGLTDIST3.Item("ACCT_CODE") & "'"
        Next
        ACCT_CODEs = Mid$(ACCT_CODEs, 2)

        For Each rowGLTDIST2 As DataRow In dst.Tables("GLTDIST2").Select("ACCT_SEG_ID = '" & ACCT_SEG_ID & "'")
            rowGLTDIST2.Item("BASIS_AMT") = 0
            rowGLTDIST2.Item("BASIS_PCT") = 0
            rowGLTDIST2.Item("DIST_PCT") = 0
        Next

        Dim T As Decimal = 0

        Dim z As String = "SEG" & ACCT_SEG_ID & "_CODE"
        ASCMAIN1.sql = "Select " & z & ", SUM (DETL_POSTING_AMT) TOTAL" & vbCrLf _
            & " from GLTDETL1 " & vbCrLf _
            & " where OPS_YYYYPP between '" & OPS_YYYYPP_1 & "' and '" & OPS_YYYYPP_2 & "'" & vbCrLf _
            & " and ACCT_CODE in (" & ACCT_CODEs & ")" & vbCrLf _
            & " group by " & z
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ACCT_SEG_CODE As String = row.Item("ACCT_SEG_CODE")
            Dim TOTAL As String = Val(row.Item("TOTAL") & "")
            Dim rowGLTDIST2 As DataRow = dst.Tables("GLTDIST2").Rows.Find(New String() {DIST_CODE, ACCT_SEG_ID, ACCT_SEG_CODE})
            If rowGLTDIST2 IsNot Nothing Then
                rowGLTDIST2.Item("BASIS_AMT") = TOTAL
                T += TOTAL
            End If
        Next

        For Each rowGLTDIST2 As DataRow In dst.Tables("GLTDIST2").Select("ACCT_SEG_ID = '" & ACCT_SEG_ID & "'")
            Dim BASIS_AMT As Decimal = Val(rowGLTDIST2.Item("BASIS_AMT") & "")
            rowGLTDIST2.Item("BASIS_PCT") = 100 * BASIS_AMT / T
            rowGLTDIST2.Item("DIST_PCT") = 100 * BASIS_AMT / T
        Next

        Normalize(False, "2")
        Normalize(False, "3")
        Normalize(False, "4")
    End Sub

    Private Sub optDistribution_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optDistribution.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_Method()
    End Sub

    Private Sub optSegment_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optSegment.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_Segment()
    End Sub

    Private Sub cmdZeroBasis_Click(sender As System.Object, e As System.EventArgs) Handles cmdZeroBasis.Click
        Zero_Basis()
    End Sub
     
    Private Sub grdGLTDIST2_BeforeDisplayDataErrorTooltip(sender As Object, e As Infragistics.Win.UltraWinGrid.BeforeDisplayDataErrorTooltipEventArgs) Handles grdGLTDIST2.BeforeDisplayDataErrorTooltip

    End Sub
End Class