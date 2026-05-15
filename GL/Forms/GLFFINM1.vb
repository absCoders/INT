Public Class GLFFINM1

    Dim row_Node As UltraWinTree.UltraTreeNode
    Dim anode As UltraWinTree.UltraTreeNode
    Dim STMT_LINE_NO_ctr As Integer = 0
    Dim setting_up As Boolean = False
    Dim A234 As String = "ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE"
    Dim GLTACCT2 As String = ""
    Dim GLTACCT3 As String = ""
    Dim GLTACCT2AND3 As String = ""
    Dim GLTFINR3 As String = ""
    Dim GLTFINRD As String = ""
    Dim GLTACCTX As String = ""
    Dim GLTACCTL As String = ""
    Dim ACCT_TYPEs As New Dictionary(Of String, String)
    Dim NEW_LNO As New Dictionary(Of Integer, Integer)
    Private WithEvents UltraTree_DropHightLight_DrawFilter As New UltraTree_DropHightLight_DrawFilter_Class()
    Dim IMAGE_FOLDER As String = ASCMAIN1.Folders("Images") & "COLUMN_NAME\STMT_LINE_TYPE\"
    Dim sub_total_map_generated_and_no_changes_have_been_made_yet As Boolean = False


#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ACCT_TYPEs.Add("B", "('A','L','E')")
        ACCT_TYPEs.Add("I", "('I','X')")

        With dst
            Create_TDA(.Tables.Add, "GLTFINR1", "*")

            Create_TDA(.Tables.Add, "GLTFINR2", "*", 1)

            ASCMAIN1.sql = "Select GLTFINR3.*, GLTACCT1.ACCT_DESC from GLTFINR3,GLTACCT1 where GLTACCT1.ACCT_CODE = GLTFINR3.ACCT_CODE"
            Create_TDA(.Tables.Add, "GLTFINR3", "**", 1)
            ASCMAIN1.sql = "Select GLTFINR4.*, GLTSEGM1.ACCT_SEG_DESC from GLTFINR4,GLTSEGM1 where GLTSEGM1.ACCT_SEG_ID = GLTFINR4.ACCT_SEG_ID and GLTSEGM1.ACCT_SEG_CODE = GLTFINR4.ACCT_SEG_CODE"
            Create_TDA(.Tables.Add, "GLTFINR4", "**", 1)

            ASCMAIN1.sql = "Select GLTFINR2.*, GLTFINR2.STMT_LINE_DESC DESCRIPTION, '0' REG_1, '0' REG_2, '0' REG_3, '0' REG_4, '0' REG_5, '0' REG_6, TATWORK1.W_AMT AMT_1, TATWORK1.W_AMT AMT_2, TATWORK1.W_AMT AMT_3, TATWORK1.W_INT STMT_LINE_NO2 from GLTFINR2,TATWORK1"
            Create_TDA(.Tables.Add, "GLTFINRM", "**", 0, False, "", 2)
            .Tables("GLTFINRM").Columns("DESCRIPTION").MaxLength = -1

            ASCMAIN1.sql = "Select GLTFINR3.ACCT_CODE, GLTFINR3.STMT_LINE_NO from GLTFINR3 where ROWNUM < 1"
            GLTACCTX = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTACCTX & " Add Primary Key (ACCT_CODE, STMT_LINE_NO)")
            ASCMAIN1.sql = "Select * from " & GLTACCTX
            Create_TDA(.Tables.Add, "GLTACCTX", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select GLTACCT1.ACCT_CODE, GLTACCT1.ACCT_DESC, SUBSTR(TATWORK1.W_TXT,1,100) STMT_LINE_NOS from GLTACCT1,TATWORK1 where ROWNUM < 1"
            GLTACCTL = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTACCTL & " Add Primary Key (ACCT_CODE)")
            ASCMAIN1.sql = "Select * from " & GLTACCTL
            Create_TDA(.Tables.Add, "GLTACCTL", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select GLTFINR3.*, GLTACCT1.ACCT_DESC, TATWORK1.W_AMT AMT_1, TATWORK1.W_AMT AMT_2, TATWORK1.W_AMT AMT_3 from GLTFINR3,GLTACCT1,TATWORK1"
            GLTFINRD = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
            ASCMAIN1.sql = "Alter Table " & GLTFINRD & " Add Primary Key (STMT_CODE, STMT_LINE_NO, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
            ASCMAIN1.sql = "Select * from " & GLTFINRD
            Create_TDA(.Tables.Add, "GLTFINRD", "**", 0, False, "", 6)

            ASCMAIN1.sql = "Select * from GLTFINR3 where ROWNUM < 1"
            GLTFINR3 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
            ASCMAIN1.sql = "Alter Table " & GLTFINR3 & " Add Primary Key (STMT_CODE, STMT_LINE_NO, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE)"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

            ASCMAIN1.sql = "Select X.ACCT_CODE, X.SEG2_CODE, X.SEG3_CODE, X.SEG4_CODE, GLTACCT1.ACCT_DESC from " _
                & " (Select " & A234 & " from " & GLTFINR3 & " group by " & A234 & " having Count (*) > 1)" _
                & " X, GLTACCT1 where GLTACCT1.ACCT_CODE = X.ACCT_CODE"
            Create_TDA(.Tables.Add, "GLTFINRX", "**", 0, False, "", 4)
            ASCMAIN1.sql = "Select X.ACCT_CODE, X.SEG2_CODE, X.SEG3_CODE, X.SEG4_CODE, GLTACCT1.ACCT_DESC from " _
                & " ((Select Distinct " & A234 & " from GLTACCT3 union Select Distinct " & A234 & " from GLTACCT2) minus Select Distinct " & A234 & " from " & GLTFINR3 & ")" _
                & " X, GLTACCT1 where GLTACCT1.ACCT_CODE = X.ACCT_CODE"
            Create_TDA(.Tables.Add, "GLTFINRY", "**", 0, False, "", 4)

            Dim x As String = Get_SelectCommand("GLTFINRY")

            .Relations.Add("GLTFINRD", _
            New DataColumn() {.Tables("GLTFINRM").Columns("STMT_CODE"), .Tables("GLTFINRM").Columns("STMT_LINE_NO")}, _
            New DataColumn() {.Tables("GLTFINRD").Columns("STMT_CODE"), .Tables("GLTFINRD").Columns("STMT_LINE_NO")})

            With .Tables.Add("GLTACCTS")
                .Columns.Add("ACCT_CODE")
                .Columns.Add("STMT_LINE_NO", GetType(System.Int64))
                .Columns.Add("STMT_LINE_DESC")
            End With
            Create_Relation("GLTACCTL", "GLTACCTS", "ACCT_CODE")

        End With

        ASCMAIN1.sql = "Select STMT_CODE,STMT_DESC from GLTFINR1 order by STMT_CODE"
        cbeSTMT_CODE_COPY.DataSource = ASCDATA1.GetDataTable

        Create_Lookup("GLTACCT1")
        Create_Lookup("GLTSEGM1")
        Create_Lookup("GLTFINR1")

        'grdGLTFINR2.DataSource = DVWs("GLTFINR2")
        grdGLTFINR3.DataSource = DVWs("GLTFINR3")
        grdGLTFINR4.DataSource = DVWs("GLTFINR4")
        grdGLTFINRM.DataSource = DVWs("GLTFINRM")

        grdGLTACCTL.DataSource = dst.Tables("GLTACCTL")
        grdGLTFINRX.DataSource = dst.Tables("GLTFINRX")
        grdGLTFINRY.DataSource = dst.Tables("GLTFINRY")

        Get_PARM("GLTPARM1")
        For i As Integer = 2 To 4
            Dim z As String = "SEG" & CStr(i)
            grdGLTFINRM.DisplayLayout.Bands("GLTFINRD").Columns("SEG" & CStr(i) & "_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & ""

            If ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & "" = "" Then
                grdGLTFINR3.DisplayLayout.Bands("GLTFINR3").Columns(z & "_CODE").Hidden = True
                Absx1.chkFor("BY_SEG" & CStr(i)).Visible = False
                grdGLTFINRX.DisplayLayout.Bands("GLTFINRX").Columns(z & "_CODE").Hidden = True
                grdGLTFINRY.DisplayLayout.Bands("GLTFINRY").Columns(z & "_CODE").Hidden = True
            Else
                grdGLTFINR3.DisplayLayout.Bands("GLTFINR3").Columns(z & "_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & ""
                optSEGS.ValueList.ValueListItems.Add(CStr(i), ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC"))
                Absx1.chkFor("BY_SEG" & CStr(i)).Text = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC")
                grdGLTFINRX.DisplayLayout.Bands("GLTFINRX").Columns(z & "_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & ""
                grdGLTFINRY.DisplayLayout.Bands("GLTFINRY").Columns(z & "_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & ""
            End If
        Next

        Set_cmbYP("RYP", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -24, 12, 0)

        Call Bind_Controls(Me, "GLTFINR2", TBLs("GLTFINR2"))

        With tvwGLTFINR2

            .Appearances.Add("DropHighLightAppearance")
            With .Appearances("DropHighLightAppearance")
                .BackColor = System.Drawing.Color.Cyan
            End With

            .DrawFilter = UltraTree_DropHightLight_DrawFilter
            .Override.SelectionType = UltraWinTree.SelectType.ExtendedAutoDrag

            .Override.CellClickAction = UltraWinTree.CellClickAction.Default
            .ViewStyle = UltraWinTree.ViewStyle.Standard
            .AllowDrop = True
            .Override.AllowCut = DefaultableBoolean.True
            .Override.AllowCopy = DefaultableBoolean.True
            .Override.AllowPaste = DefaultableBoolean.True
            .Override.AllowDelete = DefaultableBoolean.True
            '.Override.ActiveNodeAppearance.BackColor = Drawing.Color.Yellow
            .Override.ActiveNodeAppearance.BackColor = Drawing.Color.Blue
            .Override.ActiveNodeAppearance.ForeColor = Drawing.Color.White
        End With

        If optSEGS.ValueList.ValueListItems.Count = 0 Then
            grdGLTFINR4.Visible = False
            cmdSEGS.Visible = False
        Else
            optSEGS.CheckedIndex = 0
        End If


        With grdGLTFINRM.DisplayLayout.Bands("GLTFINRM")
            .SortedColumns.Clear()
            .SortedColumns.Add("STMT_LINE_NO", False)
            .Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select
            For i As Integer = 1 To 6
                With .Columns("REG_" & CStr(i))
                    .Header.Caption = CStr(i)
                    .Header.Appearance.TextHAlign = HAlign.Center
                    .CellAppearance.TextHAlign = HAlign.Center
                    .Style = UltraWinGrid.ColumnStyle.Button
                    .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                    .Width = 30
                    '.CellButtonAppearance.BackColor = Drawing.Color.BlanchedAlmond
                    '.CellButtonAppearance.ForeColor = Drawing.Color.Azure

                End With
            Next
            .Columns("AMT_1").Header.Caption = "TY_MTD"
            .Columns("AMT_2").Header.Caption = "TY_YTD"
            .Columns("AMT_3").Header.Caption = "LY_YTD"
            .Columns("AMT_1").Format = "##,##0"
            .Columns("AMT_2").Format = "##,##0"
            .Columns("AMT_3").Format = "##,##0"

            For Each c As String In New String() {"STMT_LINE_TYPE", "STMT_LINE_PRINT", "STMT_LINE_DC"}
                .Columns(c).Header.Appearance.TextHAlign = HAlign.Center
                .Columns(c).CellAppearance.TextHAlign = HAlign.Center
                .Columns(c).Width = 80
            Next
        End With

        With grdGLTFINRM.DisplayLayout.Bands("GLTFINRD")
            .SortedColumns.Clear()
            .SortedColumns.Add("ACCT_CODE", False)
            .SortedColumns.Add("SEG2_CODE", False)
            .SortedColumns.Add("SEG3_CODE", False)
            .SortedColumns.Add("SEG4_CODE", False)
            .Columns("AMT_1").Header.Caption = "TY_MTD"
            .Columns("AMT_2").Header.Caption = "TY_YTD"
            .Columns("AMT_3").Header.Caption = "LY_YTD"
            .Columns("AMT_1").Format = "##,##0"
            .Columns("AMT_2").Format = "##,##0"
            .Columns("AMT_3").Format = "##,##0"
        End With

        Create_Summary(grdGLTACCTL, "ACCT_CODE", "Count")
        Create_Summary(grdGLTFINRX, "ACCT_CODE", "Count")
        Create_Summary(grdGLTFINRY, "ACCT_CODE", "Count")

        grdGLTFINRM.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay
        'grdGLTFINR2.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Call Validate_Code("STMT_CODE", True)

                If Absx1.optFor("STMT_TYPE").CheckedIndex = -1 Then
                    EMsg &= vbCr & "Statement Type Must be Defined"
                End If

                If Absx1.txtFor("STMT_DESC").Text = "" Then
                    EMsg &= vbCr & "Statement Description Must be Entered"
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("STMT_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Edit"
                Call Validate_Code("STMT_CODE")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("STMT_CODE").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

            Case "Cancel"
                If MsgBox("OK to Lose any Changes Made to Statement " & HFs("STMT_CODE"), _
                        MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

            Case "Delete"
                If MsgBox("OK to Delete Entire Statement " & HFs("STMT_CODE"), _
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

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

            Case "Cancel"
                Mode_Settings(False)

            Case "Delete"
                If EntryMode = "E" Then
                    Call Delete_Record()
                End If
                Mode_Settings(False)

            Case "Sub-Total Map"
                Generate_Map()
                UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Sub-Total Map")
                sub_total_map_generated_and_no_changes_have_been_made_yet = True

            Case "Integrity Checks"
                Generate_Map()
                Load_GL_Data("I")
                grdGLTACCTL.Visible = True
                grdGLTFINRX.Visible = True
                grdGLTFINRY.Visible = True
                UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Integrity Checks")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Delete").Settings.Enabled = iScreenMode

                .Groups("Generate").Items("Sub-Total Map").Settings.Enabled = iScreenMode
                .Groups("Generate").Items("Integrity Checks").Settings.Enabled = iScreenMode

                .Groups("Screen Control").Items("Cancel").Visible = Not This_Record_Inquiry_Only
                .Groups("Screen Control").Items("Update").Visible = Not ScreenMode Or Not This_Record_Inquiry_Only
                .Groups("Screen Control").Items("Delete").Visible = Not ScreenMode Or Not This_Record_Inquiry_Only
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)
        Absx1.txtFor("STMT_DESC").ReadOnly = False
        Absx1.cbeFor("STMT_CODE_COPY").ReadOnly = False
        UltraTabControl1.Visible = tf

        grdGLTACCTL.Visible = False
        grdGLTFINRX.Visible = False
        grdGLTFINRY.Visible = False

        cbeSTMT_CODE_COPY.Visible = (EntryMode = "N")
        lblSTMT_CODE_COPY.Visible = (EntryMode = "N")


        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        setting_up = True

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"GLTFINR1", "GLTFINR2", "GLTFINR3", "GLTFINR4", _
                                                       "GLTACCTL", "GLTACCTS", "GLTFINRX", "GLTFINRY"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        Call Setup_Tab()
        grdGLTFINRM.Visible = False
        grpFetch.Visible = False

        setting_up = False
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        Dim rowGLTFINR1 As DataRow = Fill_Record("GLTFINR1", HFs("STMT_CODE"), EntryMode = "N")
        If EntryMode = "N" Then
            rowGLTFINR1.Item("STMT_TYPE") = HFs("STMT_TYPE")
            rowGLTFINR1.Item("STMT_DESC") = HFs("STMT_DESC")
        Else
            Save_Header_Fields(UltraGroupBox1)
        End If
        Fill_Records("GLTFINR2", HFs("STMT_CODE"))
        Fill_Records("GLTFINR3", HFs("STMT_CODE"))
        Fill_Records("GLTFINR4", HFs("STMT_CODE"))

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Setup_Tree()
    End Sub

    Sub Delete_Record()
        Call BeginTrans()

        For Each TABLE_NAME As String In New String() {"GLTFINR1", "GLTFINR2", "GLTFINR3", "GLTFINR4"}
            ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where STMT_CODE = :PARM1"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", HFs("STMT_CODE"))
        Next

        Call CommitTrans("Delete Complete")
    End Sub

    Sub Update_Record()
        Call BeginTrans()

        Call Generate_Map()

        dst.EnforceConstraints = False

        Call ReNumber_Lines("GLTFINR2")
        Call ReNumber_Lines("GLTFINR3")
        Call ReNumber_Lines("GLTFINR4")

        Call Update_Record_TDA("GLTFINR1")
        Call Update_Record_TDA("GLTFINR2", "Delete from GLTFINR2 where STMT_CODE = '" & HFs("STMT_CODE") & "'")
        Call Update_Record_TDA("GLTFINR3", "Delete from GLTFINR3 where STMT_CODE = '" & HFs("STMT_CODE") & "'")
        Call Update_Record_TDA("GLTFINR4", "Delete from GLTFINR4 where STMT_CODE = '" & HFs("STMT_CODE") & "'")
        dst.EnforceConstraints = True

        Call CommitTrans("Update Complete")
    End Sub

    Sub ReNumber_Lines(ByVal TABLE_NAME As String)
        For Each row As DataRow In dst.Tables(TABLE_NAME).Select("", "", DataViewRowState.CurrentRows)
            row.Item("STMT_LINE_NO") = -1 * row.Item("STMT_LINE_NO")
        Next
        For Each row As DataRow In dst.Tables(TABLE_NAME).Select("", "", DataViewRowState.CurrentRows)
            If NEW_LNO.ContainsKey(-1 * row.Item("STMT_LINE_NO")) Then
                row.Item("STMT_LINE_NO") = NEW_LNO(-1 * row.Item("STMT_LINE_NO"))
                If TABLE_NAME = "GLTFINR2" Then
                    Dim rowGLTFINRM As DataRow = dst.Tables("GLTFINRM").Rows.Find(New Object() {HFs("STMT_CODE"), row.Item("STMT_LINE_NO")})
                    row.Item("STMT_LINE_LEVEL") = rowGLTFINRM.Item("STMT_LINE_LEVEL")
                End If
            Else
                row.Delete()
            End If
        Next
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special( _
    ByVal ctl As Control, _
    ByVal COLUMN_NAME As String, _
    Optional ByRef sql_where As String = "", _
    Optional ByRef cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STMT_CODE"
                If Absx1.optFor("STMT_TYPE").CheckedIndex <> -1 Then
                    sql_where = "STMT_TYPE = '" & Absx1.optFor("STMT_TYPE").Value & "'"
                End If
        End Select
    End Sub

#End Region

#Region "Popup Menus"
    Overrides Sub Load_Popup_Menus()
        Call Load_Popup_Menu(tvwGLTFINR2, "BBB", "Insert Above", "Insert Below", "Insert Within")
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        STMT_LINE_NO_ctr = STMT_LINE_NO_ctr + 1
        Dim STMT_LINE_DESC As String = "{New Menu Item}"

        Dim rowGLTFINR2 As DataRow = dst.Tables("GLTFINR2").NewRow
        rowGLTFINR2.Item("STMT_CODE") = HFs("STMT_CODE")
        rowGLTFINR2.Item("STMT_LINE_NO") = STMT_LINE_NO_ctr
        rowGLTFINR2.Item("STMT_LINE_TYPE") = "D"
        rowGLTFINR2.Item("STMT_LINE_DESC") = STMT_LINE_DESC
        rowGLTFINR2.Item("STMT_LINE_PRINT") = "P"
        rowGLTFINR2.Item("STMT_LINE_DC") = "D"
        rowGLTFINR2.Item("STMT_LINE_ACCTS") = "S"
        rowGLTFINR2.Item("STMT_LINE_SEG2_SEL") = "A"
        rowGLTFINR2.Item("STMT_LINE_SEG3_SEL") = "A"
        rowGLTFINR2.Item("STMT_LINE_SEG4_SEL") = "A"

        For Each c As String In New String() _
            {"STMT_BOLD_LINE", "STMT_SKIP_LINE", "STMT_DRAW_LINE", "STMT_SUBT_SHOW", "STMT_LINE_REF_PCT", "STMT_LINE_REF_SET", _
           "STMT_SUBT_ADD1", "STMT_SUBT_ADD2", "STMT_SUBT_ADD3", "STMT_SUBT_ADD4", "STMT_SUBT_ADD5", "STMT_SUBT_ADD6"}
            rowGLTFINR2.Item(c) = "0"
        Next

        dst.Tables("GLTFINR2").Rows.Add(rowGLTFINR2)

        If tvwGLTFINR2.Nodes.Count = 0 Then
            anode = tvwGLTFINR2.Nodes.Add(CStr(STMT_LINE_NO_ctr), STMT_LINE_DESC)
        Else
            Select Case e.Tool.Key
                Case "Insert Above"
                    If tvwGLTFINR2.ActiveNode Is Nothing Then
                        Exit Sub
                    Else
                        If tvwGLTFINR2.ActiveNode.IsRootLevelNode Then
                            anode = tvwGLTFINR2.Nodes.Insert(tvwGLTFINR2.ActiveNode.Index, CStr(STMT_LINE_NO_ctr), STMT_LINE_DESC)
                        Else
                            anode = tvwGLTFINR2.ActiveNode.Parent.Nodes.Insert(tvwGLTFINR2.ActiveNode.Index, CStr(STMT_LINE_NO_ctr), STMT_LINE_DESC)
                        End If
                    End If

                Case "Insert Below"
                    If tvwGLTFINR2.ActiveNode Is Nothing Then
                        Exit Sub
                    Else
                        If tvwGLTFINR2.ActiveNode.IsRootLevelNode Then
                            anode = tvwGLTFINR2.Nodes.Insert(tvwGLTFINR2.ActiveNode.Index + 1, CStr(STMT_LINE_NO_ctr), STMT_LINE_DESC)
                        Else
                            anode = tvwGLTFINR2.ActiveNode.Parent.Nodes.Insert(tvwGLTFINR2.ActiveNode.Index + 1, CStr(STMT_LINE_NO_ctr), STMT_LINE_DESC)
                        End If
                    End If

                Case "Insert Within"
                    anode = tvwGLTFINR2.ActiveNode.Nodes.Add(CStr(STMT_LINE_NO_ctr), STMT_LINE_DESC)
                    anode.Parent.Expanded = True
            End Select
        End If


        'anode.Key = STMT_LINE_NO_ctr
        'anode.Text = "{New Menu Item}"

        Call Setup_Node("D")
        'anode.Expanded = False

        'tvwGLTFINR2.SelectedNodes.Clear()
        'tvwGLTFINR2.ActiveNode = anode
        anode.BeginEdit() ' .BeginCellEdit(anode.Cells("MENU_ITEM_DESC").Column)
    End Sub

#End Region
#Region "grdGLTFINR3"
    '    Private Sub grdGLTFINR3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTFINR3.AfterCellUpdate
    '        Select Case e.Cell.Column.Key
    '            Case "CUST_CODE"
    '                If e.Cell.Text = "" Then
    '                    grdGLTFINR3.ActiveRow.Cells("NON_AR").Value = "1"
    '                Else
    '                    grdGLTFINR3.ActiveRow.Cells("NON_AR").Value = "0"
    '                    grdCodeDesc(grdGLTFINR3, "ARTCUST1", "CUST_CODE", "CUST_NAME")
    '                End If

    '            Case "NON_AR"
    '                If e.Cell.Text = "1" Then
    '                    grdGLTFINR3.ActiveRow.Cells("CUST_CODE").Value = ""
    '                End If
    '        End Select
    '    End Sub

    '    Private Sub grdGLTFINR3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTFINR3.AfterRowActivate
    '        With grdGLTFINR3.DisplayLayout.Bands(0)
    '            If grdGLTFINR3.ActiveRow.IsAddRow Then
    '                .Columns("NON_AR").CellActivation = UltraWinGrid.Activation.AllowEdit
    '                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
    '                .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
    '                grdGLTFINR3.ActiveCell = grdGLTFINR3.ActiveRow.Cells("CUST_CODE")
    '                grdGLTFINR3.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
    '            Else
    '                .Columns("NON_AR").CellActivation = UltraWinGrid.Activation.NoEdit
    '                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
    '                If grdGLTFINR3.ActiveRow.Cells("NON_AR").Text = "1" Then
    '                    .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
    '                Else
    '                    .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
    '                End If
    '            End If
    '        End With
    '    End Sub

    '    Private Sub grdGLTFINR3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTFINR3.BeforeRowUpdate
    '        With grdGLTFINR3
    '            If e.Row.Cells("NON_AR").Text = "1" Then
    '                If e.Row.Cells("CUST_NAME").Text = "" Then
    '                    MsgBox("You Must Enter a Name for Non-AR Payments", MsgBoxStyle.OkOnly, "Cannot Update Row")
    '                    e.Cancel = True
    '                End If
    '            Else
    '                If e.Row.Cells("CUST_CODE").Text = "" Then
    '                    MsgBox("Missing Value for Customer Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
    '                    e.Cancel = True
    '                Else
    '                    Call LookUp("ARTCUST1", "CUST_CODE")
    '                    If cdr Is Nothing Then
    '                        MsgBox("Invalid Value entered for Customer Code (" & e.Row.Cells("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
    '                        e.Cancel = True
    '                    End If
    '                End If
    '            End If

    '            If Not e.Cancel Then
    '                If e.Row.Cells("PYMT_BATCH_NO").Text = "" Then
    '                    .ActiveRow.Cells("PYMT_BATCH_NO").Value = Absx1.CtlFor("PYMT_BATCH_NO").Text
    '                    .ActiveRow.Cells("PYMT_BATCH_LNO").Value = Val(dst.Tables("GLTFINR3").Compute("Max(PYMT_BATCH_LNO)", "") & "") + 1
    '                End If
    '            End If

    '        End With
    '    End Sub

    '    Private Sub grdGLTFINR3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTFINR3.ClickCellButton
    '        Dim sql_where As String = ""
    '        Call grdClickCellButton(grdGLTFINR3, sql_where, sql_where <> "")
    '    End Sub

    '    Private Sub grdGLTFINR3_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdGLTFINR3.Error
    '        grdGLTFINR3.ActiveRow.CancelUpdate()
    '    End Sub


    '    Private Sub grdGLTFINR3_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTFINR3.InitializeLayout

    '    End Sub

    '    Private Sub grdGLTFINR3_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTFINR3.InitializeRow

    '    End Sub

    '    Private Sub grdGLTFINR3_AfterRowCancelUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdGLTFINR3.AfterRowCancelUpdate

    '    End Sub

    '    Private Sub grdGLTFINR3_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdGLTFINR3.BeforeExitEditMode
    '        If grdGLTFINR3.ActiveRow.Cells("CUST_NAME").Text = "" Then
    '            'e.Cancel = True
    '        End If
    '        'Stop
    '    End Sub

    '    Private Sub grdGLTFINR3_BeforeRowDeactivate(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles grdGLTFINR3.BeforeRowDeactivate

    '    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "STMT_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Me.UltraGroupBox1.Select() ' to force txt_Leave event to fire, for formatting
                    Call Click_New_or_Edit()
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "STMT_CODE"
                Call Click_New_or_Edit()
        End Select
    End Sub
#End Region

    Sub Click_New_or_Edit()
        If LookUp("GLTFINR1", Absx1.txtFor("STMT_CODE").Text) Is Nothing Then
            Call Click_Command("New")
        Else
            Call Click_Command("Edit")
        End If
    End Sub

#Region "tvwGLTFINR2"

    Private Sub tvwGLTFINR2_AfterActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.NodeEventArgs) Handles tvwGLTFINR2.AfterActivate
        Call Activate_Node()
    End Sub

    Private Sub tvwGLTFINR2_AfterLabelEdit(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTree.NodeEventArgs) Handles tvwGLTFINR2.AfterLabelEdit
        Dim ML As Integer = DVWs("GLTFINR2").Table.Columns("STMT_LINE_DESC").MaxLength
        Dim STMT_LINE_DESC As String = e.TreeNode.Text

        If e.TreeNode.Text.Length > ML Then
            MsgBox("Maximum number of Characters (" & CStr(ML) & ") exceeded in Label Description (" & e.TreeNode.Text & ")", _
                   MsgBoxStyle.OkOnly, _
                   "Description will be Trucated")
            STMT_LINE_DESC = Mid(STMT_LINE_DESC, 1, ML)
        End If
        DVWs("GLTFINR2")(0).Item("STMT_LINE_DESC") = STMT_LINE_DESC
        sub_total_map_generated_and_no_changes_have_been_made_yet = False
    End Sub

    Private Sub tvwGLTFINR2_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwGLTFINR2.DragDrop
        Dim Node As UltraWinTree.UltraTreeNode
        Dim SelectedNodes As UltraWinTree.SelectedNodesCollection
        Dim DropNode As UltraWinTree.UltraTreeNode
        Dim i As Integer

        DropNode = UltraTree_DropHightLight_DrawFilter.DropHightLightNode

        SelectedNodes = e.Data.GetData(GetType(UltraWinTree.SelectedNodesCollection))
        SelectedNodes = SelectedNodes.Clone()

        SelectedNodes.SortByPosition()

        Select Case UltraTree_DropHightLight_DrawFilter.DropLinePosition
            Case DropLinePositionEnum.OnNode
                For i = 0 To SelectedNodes.Count - 1
                    Node = SelectedNodes(i)
                    Node.Reposition(DropNode.Nodes)
                Next
            Case DropLinePositionEnum.BelowNode
                For i = 0 To SelectedNodes.Count - 1
                    Node = SelectedNodes(i)
                    Node.Reposition(DropNode, UltraWinTree.NodePosition.Next)
                    DropNode = Node
                Next
            Case DropLinePositionEnum.AboveNode
                For i = 0 To SelectedNodes.Count - 1
                    Node = SelectedNodes(i)
                    Node.Reposition(DropNode, UltraWinTree.NodePosition.Previous)
                Next
        End Select

        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
    End Sub

    Private Sub tvwGLTFINR2_DragLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwGLTFINR2.DragLeave
        UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
    End Sub

    Private Sub tvwGLTFINR2_DragOver(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles tvwGLTFINR2.DragOver
        Dim Node As UltraWinTree.UltraTreeNode
        Dim PointInTree As System.Drawing.Point

        With tvwGLTFINR2
            PointInTree = .PointToClient(New System.Drawing.Point(e.X, e.Y))

            Node = .GetNodeFromPoint(PointInTree)

            If Node Is Nothing Then
                e.Effect = DragDropEffects.None
                UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
                Return
            End If

            If Me.IsParentNode(Node) And Me.IsParentNodeSelected(Me.tvwGLTFINR2) Then
                If PointInTree.Y > (Node.Bounds.Top + 2) AndAlso PointInTree.Y < (Node.Bounds.Bottom - 2) Then
                    e.Effect = DragDropEffects.None
                    UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
                    Return
                End If
            End If

            'If IsAnyParentSelected(Node) Then
            '    e.Effect = DragDropEffects.None
            '    UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
            '    Return
            'End If

            UltraTree_DropHightLight_DrawFilter.SetDropHighlightNode(Node, PointInTree)
            e.Effect = DragDropEffects.Move
        End With
    End Sub

    Private Sub tvwGLTFINR2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvwGLTFINR2.MouseUp
        tvwGLTFINR2.SelectedNodes.Clear()
        Dim anode As Infragistics.Win.UltraWinTree.UltraTreeNode = tvwGLTFINR2.GetNodeFromPoint(e.Location)
        If anode IsNot Nothing Then
            anode.Selected = True
            tvwGLTFINR2.ActiveNode = anode
        End If
    End Sub

    Private Sub tvwGLTFINR2_QueryContinueDrag(ByVal sender As Object, ByVal e As System.Windows.Forms.QueryContinueDragEventArgs) Handles tvwGLTFINR2.QueryContinueDrag
        If e.EscapePressed Then
            e.Action = DragAction.Cancel
            UltraTree_DropHightLight_DrawFilter.ClearDropHighlight()
        End If
    End Sub

    Private Sub tvwGLTFINR2_SelectionDragStart(ByVal sender As Object, ByVal e As System.EventArgs) Handles tvwGLTFINR2.SelectionDragStart
        tvwGLTFINR2.DoDragDrop(tvwGLTFINR2.SelectedNodes, DragDropEffects.Move)
    End Sub
#End Region


    Private Function IsParentNode(ByVal Node As UltraWinTree.UltraTreeNode) As Boolean
        Dim Tag As String
        Tag = Node.Tag
        If Tag Is Nothing Then
            Return False
        Else
            Return Split(Tag, Chr(1))(1) = "M"
        End If

    End Function

    Private Function IsParentNodeSelected(ByVal Tree As UltraWinTree.UltraTree) As Boolean
        For Each SelectedNode As UltraWinTree.UltraTreeNode In Tree.SelectedNodes
            If Me.IsParentNode(SelectedNode) Then Return True
        Next
        Return False
    End Function

    Private Function IsAnyParentSelected(ByVal Node As UltraWinTree.UltraTreeNode) As Boolean
        Dim ParentNode As UltraWinTree.UltraTreeNode

        ParentNode = Node.Parent
        Do Until ParentNode Is Nothing
            If ParentNode.Selected Then Return True
            ParentNode = ParentNode.Parent
        Loop
        Return False
    End Function

    Private Sub UltraTree_DropHightLight_DrawFilter_Invalidate(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraTree_DropHightLight_DrawFilter.Invalidate
        tvwGLTFINR2.Invalidate()
    End Sub

    Private Sub UltraTree_DropHightLight_DrawFilter_QueryStateAllowedForNode(ByVal sender As Object, ByVal e As UltraTree_DropHightLight_DrawFilter_Class.QueryStateAllowedForNodeEventArgs) Handles UltraTree_DropHightLight_DrawFilter.QueryStateAllowedForNode
        If Not IsParentNode(e.Node) Then
            e.StatesAllowed = DropLinePositionEnum.AboveNode Or DropLinePositionEnum.BelowNode
            UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 2
        Else
            If e.Node.Selected Then
                e.StatesAllowed = DropLinePositionEnum.AboveNode Or DropLinePositionEnum.BelowNode
                UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 2
            Else
                UltraTree_DropHightLight_DrawFilter.EdgeSensitivity = e.Node.Bounds.Height / 3
            End If
        End If
    End Sub

    Sub Setup_Node(ByVal STMT_LINE_TYPE As String)
        If anode Is Nothing Then Exit Sub
        If STMT_LINE_TYPE = "H" Then
            anode.Override.NodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, STMT_LINE_TYPE)
            anode.Override.ExpandedNodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, STMT_LINE_TYPE & "_EXP")
            anode.Override.NodeAppearance.FontData.Bold = DefaultableBoolean.True
        ElseIf STMT_LINE_TYPE = "D" Or STMT_LINE_TYPE = "S" Then
            If anode IsNot Nothing Then
                anode.Override.NodeAppearance.Image = ASCMAIN1.Get_Image(IMAGE_FOLDER, STMT_LINE_TYPE)
            End If
        Else
            Stop
        End If
    End Sub

    Private Sub SplitContainer1_Panel2_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles SplitContainer1.Panel2.Paint

    End Sub

    Private Sub grdGLTFINR3_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTFINR3.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = e.Cell.Value & ""
                'e.Cell.Value = ASCMAIN1.Format_Field(ACCT_CODE, e.Cell.Column.Key)

                grdCodeDesc(grdGLTFINR3, "GLTACCT1", "ACCT_CODE", "ACCT_DESC")
                For i As Integer = 2 To 4
                    If e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Text = "" Then
                        e.Cell.Row.Cells("SEG" & CStr(i) & "_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                    End If
                Next
        End Select
    End Sub

    Private Sub grdGLTFINR3_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTFINR3.AfterExitEditMode
        Select Case grdGLTFINR3.ActiveCell.Column.Key
            Case "ACCT_CODE"
                Dim ACCT_CODE As String = grdGLTFINR3.ActiveCell.Text
                If ACCT_CODE <> "" Then
                    grdGLTFINR3.ActiveCell.Value = ASCMAIN1.Format_Field(ACCT_CODE, grdGLTFINR3.ActiveCell.Column.Key)
                End If

        End Select
    End Sub

    Private Sub grdGLTFINR3_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTFINR3.AfterRowActivate
        With grdGLTFINR3
            If .ActiveRow.IsAddRow Then
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .ActiveCell = grdGLTFINR3.ActiveRow.Cells("ACCT_CODE")
                .PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .DisplayLayout.Bands(0).Columns("ACCT_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        End With
    End Sub

    Private Sub grdGLTFINR3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdGLTFINR3.BeforeRowUpdate
        With grdGLTFINR3
            If e.Row.Cells("ACCT_CODE").Text = "" Then
                e.Cancel = True
            Else
                Call LookUp("GLTACCT1", e.Row.Cells("ACCT_CODE").Text)
                If cdr Is Nothing Then
                    MsgBox("Invalid Value entered for Acct Code (" & e.Row.Cells("ACCT_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            Dim COLUMN_NAME As String
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If e.Row.Cells(COLUMN_NAME).Column.Hidden Then
                    .ActiveRow.Cells(COLUMN_NAME).Value = ROWs("GLTPARM1").Item("GL_PARM_DEF_" & Mid(COLUMN_NAME, 1, 4)) & ""
                Else
                    If e.Row.Cells(COLUMN_NAME).Value & "" = "" Then
                        e.Cancel = True
                    Else
                        If e.Row.Cells(COLUMN_NAME).Value & "" = "*" Then
                            ' all values
                        Else
                            Call LookUp("GLTSEGM1", New String() {CStr(i), e.Row.Cells(COLUMN_NAME).Text})
                            If cdr Is Nothing Then
                                MsgBox("Invalid Value entered for " & e.Row.Cells(COLUMN_NAME).Column.Header.Caption & " (" & e.Row.Cells(COLUMN_NAME).Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                                e.Cancel = True
                            End If
                        End If
                    End If
                End If
            Next

            If Not e.Cancel Then
                If e.Row.Cells("STMT_CODE").Text = "" Then
                    .ActiveRow.Cells("STMT_CODE").Value = HFs("STMT_CODE")
                    .ActiveRow.Cells("STMT_LINE_NO").Value = anode.Key
                End If
            End If
        End With

    End Sub

    Private Sub grdGLTFINR3_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTFINR3.ClickCellButton
        If grdGLTFINR3.ActiveRow.IsAddRow Then
            Dim sql_where As String = ""
            Call grdClickCellButton(grdGLTFINR3, sql_where, sql_where <> "")
        End If
    End Sub

    Private Sub optSTMT_LINE_TYPE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSTMT_LINE_TYPE.ValueChanged
        If optSTMT_LINE_TYPE.Value Is Nothing Then
            Exit Sub
        End If
        Panel1.Visible = Not (optSTMT_LINE_TYPE.Value = "H")
        grpSTMT_LINE_DC.Visible = Not (optSTMT_LINE_TYPE.Value = "H")
        grpGLTFINR3.Visible = optSTMT_LINE_TYPE.Value = "D"
        grpSTMT_SUBT_SHOW.Visible = optSTMT_LINE_TYPE.Value = "S"
        grpSTMT_LINE_ACCTS.Visible = optSTMT_LINE_TYPE.Value = "D"
        UltraExplorerBar1.Groups("Account Selectivity").Expanded = optSTMT_LINE_TYPE.Value = "D"
        Call Setup_Node(optSTMT_LINE_TYPE.Value)

    End Sub

    Private Sub optSTMT_LINE_ACCTS_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSTMT_LINE_ACCTS.ValueChanged

        lblSTMT_LINE_ACCT_RANGE1.Visible = (optSTMT_LINE_ACCTS.Value = "R")
        lblSTMT_LINE_ACCT_RANGE2.Visible = (optSTMT_LINE_ACCTS.Value = "R")
        txtSTMT_LINE_ACCT_RANGE1.Visible = (optSTMT_LINE_ACCTS.Value = "R")
        txtSTMT_LINE_ACCT_RANGE2.Visible = (optSTMT_LINE_ACCTS.Value = "R")

        chkACCT_CODE_SEL.Visible = (optSTMT_LINE_ACCTS.Value = "S")
        cmdACCT_CODE_ADD.Visible = (optSTMT_LINE_ACCTS.Value = "S") Or (optSTMT_LINE_ACCTS.Value = "R")

        grdGLTFINR3.Visible = (optSTMT_LINE_ACCTS.Value = "S") Or (optSTMT_LINE_ACCTS.Value = "R") Or (optSTMT_LINE_ACCTS.Value = "X")
        splSEGS.Visible = (optSTMT_LINE_ACCTS.Value = "S") Or (optSTMT_LINE_ACCTS.Value = "R")

        grdGLTFINR3.DisplayLayout.Bands(0).Columns("SEG2_CODE").Hidden = Not (optSTMT_LINE_ACCTS.Value = "X" And ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "" <> "")
        grdGLTFINR3.DisplayLayout.Bands(0).Columns("SEG3_CODE").Hidden = Not (optSTMT_LINE_ACCTS.Value = "X" And ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "" <> "")
        grdGLTFINR3.DisplayLayout.Bands(0).Columns("SEG4_CODE").Hidden = Not (optSTMT_LINE_ACCTS.Value = "X" And ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "" <> "")


        If optSTMT_LINE_ACCTS.Value = "X" Then
            grdGLTFINR3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
            grdGLTFINR3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Else
            grdGLTFINR3.DisplayLayout.Override.AllowAddNew = UltraWinGrid.AllowAddNew.No
            grdGLTFINR3.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If
    End Sub

    Private Sub txtACCT_CODE_TO_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSTMT_LINE_ACCT_RANGE2.ValueChanged

    End Sub

    Private Sub UltraTabControl1_ActiveTabChanged(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinTabControl.ActiveTabChangedEventArgs) Handles UltraTabControl1.ActiveTabChanged
        Call Setup_Tab()
    End Sub

    Sub Setup_Tab()
        Dim TAB As String = ""
        If UltraTabControl1.ActiveTab IsNot Nothing And EntryMode <> "" Then
            TAB = UltraTabControl1.ActiveTab.Key
        End If

        With UltraExplorerBar1
            .Groups("Account Selectivity").Visible = (TAB = "Line Definition")
            .Groups("Legend").Visible = (TAB = "Sub-Total Map")
        End With
    End Sub

    Sub Setup_grdGLTFINRM(ByVal TF As Boolean)
        With grdGLTFINRM.DisplayLayout.Bands("GLTFINRM")
            .Columns("AMT_1").Hidden = TF
            .Columns("AMT_2").Hidden = TF
            .Columns("AMT_3").Hidden = TF
        End With
        With grdGLTFINRM.DisplayLayout.Bands("GLTFINRD")
            .Columns("AMT_1").Hidden = TF
            .Columns("AMT_2").Hidden = TF
            .Columns("AMT_3").Hidden = TF
        End With
        If TF Then
            grpFetch.Text = ""
        End If
        grpFetch.Visible = Not TF
    End Sub

    Private Sub Generate_Map()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Generating Map")

        EnforceConstraints(False)

        dst.Tables("GLTFINRM").Rows.Clear()
        NEW_LNO.Clear()
        Navigate_Tree(tvwGLTFINR2.Nodes, 1, 0)

        dst.Tables("GLTFINRD").Rows.Clear()
        For Each row As DataRow In dst.Tables("GLTFINR3").Select("", "", DataViewRowState.CurrentRows)
            If NEW_LNO.ContainsKey(row.Item("STMT_LINE_NO")) Then
                Dim rowGLTFINRD As DataRow = dst.Tables("GLTFINRD").NewRow
                rowGLTFINRD.Item("STMT_CODE") = row.Item("STMT_CODE")
                rowGLTFINRD.Item("STMT_LINE_NO") = NEW_LNO(row.Item("STMT_LINE_NO"))
                rowGLTFINRD.Item("ACCT_CODE") = row.Item("ACCT_CODE")
                rowGLTFINRD.Item("SEG2_CODE") = row.Item("SEG2_CODE")
                rowGLTFINRD.Item("SEG3_CODE") = row.Item("SEG3_CODE")
                rowGLTFINRD.Item("SEG4_CODE") = row.Item("SEG4_CODE")
                rowGLTFINRD.Item("ACCT_DESC") = LookUp("GLTACCT1", row.Item("ACCT_CODE")).Item("ACCT_DESC") & ""
                dst.Tables("GLTFINRD").Rows.Add(rowGLTFINRD)
            Else
                row.Delete()
                ' THIS FIX WAS PUT IN PLACE BECAUSE WE GOT AN ERROR AFTER DELETING A NODE FROM THE TREE
                ' PERHAPS THE DELETE TO GLTINFR3 SHOULD HAPPEN IN THE TREE'S AFTER DELETE EVENT
                ' THERE MAY BE OTHER TABLES THAT NEED TO BE CLEANED UP AS WELL
            End If
        Next

        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        Call Setup_grdGLTFINRM(True)
        grdGLTFINRM.Visible = True
        grpFetch.Visible = True

        tvwGLTFINR2.ActiveNode = tvwGLTFINR2.Nodes(0)
    End Sub

    Sub Navigate_Tree( _
    ByVal n As UltraWinTree.TreeNodesCollection, _
    ByVal LEVEL As Integer, _
    ByRef STMT_LINE_NO As Integer)

        For Each nod As UltraWinTree.UltraTreeNode In n
            Dim rowGLTFINRM As DataRow = dst.Tables("GLTFINRM").NewRow
            Dim row As DataRow = _
                dst.Tables("GLTFINR2").Rows.Find _
                (New String() {HFs("STMT_CODE"), nod.Key})
            rowGLTFINRM.ItemArray = row.ItemArray
            STMT_LINE_NO = STMT_LINE_NO + 1
            NEW_LNO.Add(nod.Key, STMT_LINE_NO)
            rowGLTFINRM.Item("STMT_LINE_NO2") = nod.Key
            rowGLTFINRM.Item("STMT_LINE_NO") = STMT_LINE_NO
            rowGLTFINRM.Item("STMT_LINE_LEVEL") = LEVEL
            If row.Item("STMT_LINE_TYPE") & "" <> "H" Then

                For I As Integer = 1 To 6
                    Dim REG As String = "STMT_SUBT_ADD" & CStr(I)
                    If row.Item(REG) & "" = "1" Then
                        rowGLTFINRM.Item("REG_" & CStr(I)) = "+"
                    End If
                Next
                Dim STMT_SUBT_SHOW As Integer = Val(row.Item("STMT_SUBT_SHOW") & "")

                If row.Item("STMT_LINE_TYPE") & "" = "S" Then
                    If STMT_SUBT_SHOW <> 0 Then
                        Dim REG_X As String = "REG_" & CStr(STMT_SUBT_SHOW)
                        If rowGLTFINRM.Item(REG_X) & "" = "+" Then
                            rowGLTFINRM.Item(REG_X) = "="
                        Else
                            rowGLTFINRM.Item(REG_X) = "*"
                        End If
                    End If
                End If
            End If

            Dim STMT_LINE_DESC As String = Space((LEVEL - 1) * 3) & row.Item("STMT_LINE_DESC") & ""
            Dim ML As Integer = dst.Tables("GLTFINRM").Columns("STMT_LINE_DESC").MaxLength
            If Len(STMT_LINE_DESC) > ML Then
                STMT_LINE_DESC = Mid(STMT_LINE_DESC, 1, ML)
            End If
            rowGLTFINRM.Item("DESCRIPTION") = STMT_LINE_DESC

            If row.Item("STMT_LINE_ACCTS") & "" = "R" Then
                Call Create_Range(row.Item("STMT_LINE_NO"), row.Item("STMT_LINE_ACCT_RANGE1"), row.Item("STMT_LINE_ACCT_RANGE2"))
            End If

            dst.Tables("GLTFINRM").Rows.Add(rowGLTFINRM)

            If nod.Nodes.Count > 0 Then
                Navigate_Tree(nod.Nodes, LEVEL + 1, STMT_LINE_NO)
            End If
        Next
    End Sub

    Private Sub cmdACCT_CODE_ADD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdACCT_CODE_ADD.Click

        Select Case optSTMT_LINE_ACCTS.Value
            Case "S"

                If chkACCT_CODE_SEL.Checked Then
                    Call Load_GL_Data("A")
                End If

                VIEW_NAME = "ACCT_CODE"

                Dim Z As String = ""
                For i As Integer = 0 To DVWs("GLTFINR3").Count - 1
                    Z &= Chr(0) & DVWs("GLTFINR3")(i).Item("ACCT_CODE")
                Next

                Dim sql_where As String = "ACCT_TYPE in (" & IIf(HFs("STMT_TYPE") = "B", "'A','L','E'", "'I','X'") & ")"

                If chkACCT_CODE_SEL.Checked Then
                    sql_where &= " and ACCT_CODE not in (SELECT DISTINCT ACCT_CODE from " & GLTACCTX & " minus SELECT DISTINCT ACCT_CODE from " & GLTACCTX & " where STMT_LINE_NO = " & anode.Key & ")"
                End If


                ASCMAIN1.CodeSelector.SQL = _
                ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, "", sql_where)

                If ASCMAIN1.CodeSelector.SQL <> "" Then
                    ASCMAIN1.CodeSelector.MultipleSelections = True
                    ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Z
                    Dim F As New ASFCODE1
                    F.ShowDialog()
                    F.Dispose()
                    If ASCMAIN1.CodeSelector.Selections <> 0 Then
                        For i As Integer = DVWs("GLTFINR3").Count - 1 To 0 Step -1
                            DVWs("GLTFINR3").Delete(i)
                        Next

                        For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                            'Dim rowGLTFINR3 As DataRow = DVWs("GLTFINR3").AddNew.Row
                            Dim rowGLTFINR3 As DataRow = TBLs("GLTFINR3").NewRow
                            With rowGLTFINR3
                                .Item("STMT_CODE") = HFs("STMT_CODE")
                                .Item("STMT_LINE_NO") = anode.Key
                                .Item("ACCT_CODE") = row.Item("ACCT_CODE")
                                .Item("ACCT_DESC") = row.Item("ACCT_DESC")
                                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                            End With
                            'rowGLTFINR3.EndEdit()
                            TBLs("GLTFINR3").Rows.Add(rowGLTFINR3)
                        Next
                    End If
                End If


            Case "R"
                Call Create_Range(anode.Key, Absx1.txtFor("STMT_LINE_ACCT_RANGE1").Text, Absx1.txtFor("STMT_LINE_ACCT_RANGE2").Text)
        End Select

    End Sub

    Private Sub cmdMapFetch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMapFetch.Click
        Call Load_GL_Data("F")
    End Sub

    Sub Create_Range( _
    ByVal STMT_LINE_NO As Integer, _
    ByVal STMT_LINE_ACCT_RANGE1 As String, _
    ByVal STMT_LINE_ACCT_RANGE2 As String)

        DVWs("GLTFINR3").RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
        For i As Integer = DVWs("GLTFINR3").Count - 1 To 0 Step -1
            DVWs("GLTFINR3").Delete(i)
        Next

        ASCMAIN1.sql = "Select * from GLTACCT1 " _
                     & " where ACCT_CODE >= '" & STMT_LINE_ACCT_RANGE1 & "'" _
                     & "   and ACCT_CODE <= '" & STMT_LINE_ACCT_RANGE2 & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowGLTFINR3 As DataRow = TBLs("GLTFINR3").NewRow
            With rowGLTFINR3
                .Item("STMT_CODE") = HFs("STMT_CODE")
                .Item("STMT_LINE_NO") = STMT_LINE_NO
                .Item("ACCT_CODE") = row.Item("ACCT_CODE")
                .Item("ACCT_DESC") = LookUp("GLTACCT1", row.Item("ACCT_CODE")).Item("ACCT_DESC")
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
            End With
            TBLs("GLTFINR3").Rows.Add(rowGLTFINR3)
        Next

    End Sub

    Sub Load_GL_Data(Optional ByVal MODE As String = "")

        'MODE = "F" (Fetch all GL Data)
        'MODE = "A" (Accounts Not Used)
        'MODE = "I" (Integrity Checks)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading GL Data")

        Dim Z As String = cmbRYP.Text
        Dim RYP As String = Mid(Z, 1, 4) & Mid(Z, 6, 2)
        Dim P As Integer = Val(Mid(RYP, 5, 2))
        Dim TY As String = Mid(RYP, 1, 4)
        Dim LY As String = Format(Val(TY) - 1, "0000")

        Dim TT As String
        If optMapActBud.Value = "B" Then
            TT = GL_Prep(LY, TY, True, , , , GLTACCT2)
            If GLTACCT2 = "" Then
                GLTACCT2 = TT
            End If
        Else
            TT = GL_Prep(LY, TY, False, , , , GLTACCT3)
            If GLTACCT3 = "" Then
                GLTACCT3 = TT
            End If
        End If

        Dim sql As String = ""
        Dim sql_GLTACCTX As String

        If GLTACCT2AND3 = "" Then
            sql = "Select X.*, GLTACCT1.ACCT_TYPE from GLTACCT1, (Select Distinct " & A234 & " from GLTACCT2 union Select Distinct " & A234 & " from GLTACCT3) X where X.ACCT_CODE = GLTACCT1.ACCT_CODE"
            GLTACCT2AND3 = ASCMAIN1.Temp_Table(sql)
        End If

        If MODE = "F" Or MODE = "I" Then
            sql = "Truncate Table " & GLTFINR3
            ASCDATA1.ExecuteSQL(sql)
        End If

        sql = "Truncate Table " & GLTACCTX
        ASCDATA1.ExecuteSQL(sql)

        'dst.Tables("GLTFINRM").Rows.Clear()
        'NEW_LNO.Clear()
        'Call Navigate_Tree(tvwGLTFINR2.Nodes, 1, 0)

        Dim STMT_LINE_NO_ALL_ELSE As Integer = 0
        Dim STMT_LINE_NO As Integer = 0
        For Each row As DataRow In TBLs("GLTFINR2").Select("STMT_LINE_TYPE = 'D'", "STMT_LINE_NO", DataViewRowState.CurrentRows)
            STMT_LINE_NO = Val(row.Item("STMT_LINE_NO"))
            sql = ""
            sql_GLTACCTX = ""

            Dim sqlx As String = ""
            If NEW_LNO.ContainsKey(STMT_LINE_NO) Then
                sqlx = "Select Distinct '" & HFs("STMT_CODE") & "' STMT_CODE, " & NEW_LNO(STMT_LINE_NO) & " STMT_LINE_NO, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE "
            End If
            Dim sqlf As String
            If MODE = "F" Then
                sqlf = " from " & TT
            Else
                sqlf = " FROM " & GLTACCT2AND3
            End If

            Select Case row.Item("STMT_LINE_ACCTS") & ""
                Case "S"
                    Dim dvwGLTFINR3 As New DataView(dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
                    Z = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        Z &= ",'" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                    Next
                    If Z <> "" Then
                        sql_GLTACCTX = " where ACCT_CODE in (" & Mid(Z, 2) & ")"

                        sql = sqlx & sqlf & " where ACCT_CODE in (" & Mid(Z, 2) & ")"
                        For s As Integer = 2 To 4

                            If row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "S" _
                            Or row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "X" Then
                                Dim dvwGLTFINR4 As New DataView(dst.Tables("GLTFINR4"))
                                dvwGLTFINR4.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO) & " and ACCT_SEG_ID = '" & CStr(s) & "'"
                                Z = ""
                                For i As Integer = 0 To dvwGLTFINR4.Count - 1
                                    Z &= ",'" & dvwGLTFINR4(i).Item("ACCT_SEG_CODE") & "'"
                                Next
                                If Z <> "" Then
                                    sql &= " and SEG" & CStr(s) & "_CODE" _
                                        & IIf(row.Item("STMT_LINE_SEG" & CStr(s) & "_SEL") & "" = "X", " NOT", "") _
                                        & " in (" & Mid(Z, 2) & ")"
                                End If
                            End If

                        Next
                    End If
                Case "R"
                    sql_GLTACCTX = " where ACCT_CODE >= '" & row.Item("STMT_LINE_ACCT_RANGE1") & "' and ACCT_CODE <= '" & row.Item("STMT_LINE_ACCT_RANGE2") & "'"
                    sql = sqlx & sqlf & sql_GLTACCTX
                Case "I"
                    If HFs("STMT_TYPE") = "B" Then
                        sql = sqlx & sqlf & " where ACCT_TYPE in " & ACCT_TYPEs("I")
                        'sql_GLTACCTX = " where ACCT_TYPE in " & ACCT_TYPEs("I")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "B"
                    If HFs("STMT_TYPE") = "I" Then
                        sql = sqlx & sqlf & " where ACCT_TYPE in " & ACCT_TYPEs("B")
                        'sql_GLTACCTX = " where ACCT_TYPE in " & ACCT_TYPEs("B")
                    Else
                        STMT_LINE_NO_ALL_ELSE = STMT_LINE_NO
                    End If

                Case "X"
                    Dim dvwGLTFINR3 As New DataView(dst.Tables("GLTFINR3"))
                    dvwGLTFINR3.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO) & " AND (SEG2_CODE <> '*' AND SEG3_CODE <> '*' AND SEG4_CODE <> '*')"
                    Z = ""
                    For i As Integer = 0 To dvwGLTFINR3.Count - 1
                        Z &= ",('" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                        Z &= ",'" & dvwGLTFINR3(i).Item("SEG2_CODE") & "'"
                        Z &= ",'" & dvwGLTFINR3(i).Item("SEG3_CODE") & "'"
                        Z &= ",'" & dvwGLTFINR3(i).Item("SEG4_CODE") & "')"
                        sql_GLTACCTX &= ",'" & dvwGLTFINR3(i).Item("ACCT_CODE") & "'"
                    Next

                    If (MODE = "F" Or MODE = "I") Then
                        Dim dvwGLTFINR3wc As New DataView(dst.Tables("GLTFINR3"))
                        dvwGLTFINR3wc.RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO) & " AND (SEG2_CODE = '*' or SEG3_CODE = '*' or SEG4_CODE = '*')"
                        For i As Integer = 0 To dvwGLTFINR3wc.Count - 1
                            Dim sqlWildCard As String = ""
                            Dim sqlWildCard_where As String = " where (ACCT_CODE"
                            Dim sqlWildCard_in As String = "(('" & dvwGLTFINR3wc(i).Item("ACCT_CODE") & "'"
                            sql_GLTACCTX &= ",'" & dvwGLTFINR3wc(i).Item("ACCT_CODE") & "'"
                            If dvwGLTFINR3wc(i).Item("SEG2_CODE") <> "*" Then
                                sqlWildCard_where &= ",SEG2_CODE"
                                sqlWildCard_in &= ",'" & dvwGLTFINR3wc(i).Item("SEG2_CODE") & "'"
                            End If
                            If dvwGLTFINR3wc(i).Item("SEG3_CODE") <> "*" Then
                                sqlWildCard_where &= ",SEG3_CODE"
                                sqlWildCard_in &= ",'" & dvwGLTFINR3wc(i).Item("SEG3_CODE") & "'"
                            End If
                            If dvwGLTFINR3wc(i).Item("SEG4_CODE") <> "*" Then
                                sqlWildCard_where &= ",SEG4_CODE"
                                sqlWildCard_in &= ",'" & dvwGLTFINR3wc(i).Item("SEG4_CODE") & "'"
                            End If
                            sqlWildCard_where &= ")"
                            sqlWildCard_in &= "))"
                            sqlWildCard = "Insert into " & GLTFINR3 & " " & sqlx & sqlf & sqlWildCard_where & " in " & sqlWildCard_in

                            ASCDATA1.ExecuteSQL(sqlWildCard)
                        Next
                    End If
                    If Z <> "" Then
                        sql = sqlx & sqlf & " where (ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE) in (" & Mid(Z, 2) & ")"
                        sql_GLTACCTX = " where ACCT_CODE in (" & Mid(sql_GLTACCTX, 2) & ")"
                    End If
            End Select

            If sql <> "" And (MODE = "F" Or MODE = "I") Then
                sql = "Insert into " & GLTFINR3 & " " & sql
                ASCDATA1.ExecuteSQL(sql)
            End If

            If sql_GLTACCTX <> "" Then
                Dim STMT_LINE_NO_X As Int32 = STMT_LINE_NO
                If NEW_LNO.ContainsKey(STMT_LINE_NO) Then
                    STMT_LINE_NO_X = NEW_LNO(STMT_LINE_NO)
                End If
                sql = "Insert into " & GLTACCTX & " Select ACCT_CODE, " & CStr(STMT_LINE_NO_X) & " STMT_LINE_NO from GLTACCT1 " & sql_GLTACCTX
                ASCDATA1.ExecuteSQL(sql)
            End If
        Next

        If STMT_LINE_NO_ALL_ELSE <> 0 Then
            If MODE = "F" Or MODE = "I" Then
                sql = "Select '" & HFs("STMT_CODE") & "' STMT_CODE, " & CStr(STMT_LINE_NO_ALL_ELSE) & " STMT_LINE_NO, ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE " _
                    & " from (" _
                    & "Select DISTINCT TT.ACCT_CODE,TT.SEG2_CODE,TT.SEG3_CODE,TT.SEG4_CODE from GLTACCT1," & TT & " TT where TT.ACCT_CODE = GLTACCT1.ACCT_CODE and GLTACCT1.ACCT_TYPE in " & ACCT_TYPEs(HFs("STMT_TYPE")) _
                    & " MINUS " _
                    & "Select DISTINCT ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE from " & GLTFINR3 _
                    & ")"
                sql = "Insert into " & GLTFINR3 & " " & sql
                ASCDATA1.ExecuteSQL(sql)
            End If

            sql = "Select ACCT_CODE, " & CStr(STMT_LINE_NO_ALL_ELSE) & " STMT_LINE_NO " _
                & " from (" _
                & "Select DISTINCT ACCT_CODE from GLTACCT1 where ACCT_TYPE in " & ACCT_TYPEs(HFs("STMT_TYPE")) _
                & " MINUS " _
                & "Select DISTINCT ACCT_CODE from " & GLTACCTX _
                & ")"
            sql = "Insert into " & GLTACCTX & " " & sql
            ASCDATA1.ExecuteSQL(sql)
        End If

        If MODE = "F" Or MODE = "I" Then

            Call ASCMAIN1.AnalyzeTable(GLTFINR3)

            'dst.Tables("GLTFINRI").Rows.Clear()
            'Call Fill_Records("GLTFINRI")

            ASCDATA1.ExecuteSQL("Truncate Table " & GLTFINRD)

            Dim sqlp As String = "NVL(TT.ACCT_BEG_BAL,0)"
            For i As Integer = 1 To P
                sqlp &= " + NVL(TT.ACCT_" & IIf(optMapActBud.Value = "A", "ACT", "BUD") & "_P" & Format(i, "00") & ",0)"
            Next
            sql = "Select GLTFINR3.STMT_CODE, GLTFINR3.STMT_LINE_NO"
            sql = sql & ", GLTFINR3.ACCT_CODE"
            Dim gby_SEGS As String = ""
            For i As Integer = 2 To 4
                If Absx1.chkFor("BY_SEG" & CStr(i)).Checked Or MODE = "I" Then
                    sql = sql & ", GLTFINR3.SEG" & CStr(i) & "_CODE"
                    gby_SEGS = gby_SEGS & ", GLTFINR3.SEG" & CStr(i) & "_CODE"
                    grdGLTFINRM.DisplayLayout.Bands("GLTFINRD").Columns("SEG" & CStr(i) & "_CODE").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & ""
                Else
                    sql = sql & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "' SEG" & CStr(i) & "_CODE"
                    grdGLTFINRM.DisplayLayout.Bands("GLTFINRD").Columns("SEG" & CStr(i) & "_CODE").Header.Caption = ""
                End If
            Next
            sql = sql & ", NULL ACCT_DESC"
            sql = sql & ", SUM (DECODE(TT.ACCT_YEAR,'" & TY & "',NVL(TT.ACCT_" & IIf(optMapActBud.Value = "A", "ACT", "BUD") & "_P" & Format(P, "00") & ",0),0)) AMT_1"
            sql = sql & ", SUM (CASE WHEN TT.ACCT_YEAR = '" & TY & "' THEN " & sqlp & " ELSE 0 END) AMT_2"
            sql = sql & ", SUM (CASE WHEN TT.ACCT_YEAR = '" & LY & "' THEN " & sqlp & " ELSE 0 END) AMT_3"
            sql = sql & " from " & GLTFINR3 & " GLTFINR3," & TT & " TT"
            sql = sql & " where GLTFINR3.ACCT_CODE = TT.ACCT_CODE"
            sql = sql & "   and GLTFINR3.SEG2_CODE = TT.SEG2_CODE"
            sql = sql & "   and GLTFINR3.SEG3_CODE = TT.SEG3_CODE"
            sql = sql & "   and GLTFINR3.SEG4_CODE = TT.SEG4_CODE"
            sql = sql & " group by GLTFINR3.STMT_CODE, GLTFINR3.STMT_LINE_NO"
            sql = sql & ", GLTFINR3.ACCT_CODE"
            sql = sql & gby_SEGS
            'sql = sql & ", GLTFINR3.SEG2_CODE, GLTFINR3.SEG3_CODE, GLTFINR3.SEG4_CODE"

            sql = "Insert into " & GLTFINRD & " " & sql
            ASCDATA1.ExecuteSQL(sql)

            Call ASCMAIN1.AnalyzeTable(GLTFINRD)

            sql = "Delete from " & GLTFINRD & " where AMT_1 = 0 and AMT_2 = 0 and AMT_3 = 0"
            ASCDATA1.ExecuteSQL(sql)


            sql = "Update " & GLTFINRD & " set ACCT_DESC = (Select ACCT_DESC from GLTACCT1 where ACCT_CODE = " & GLTFINRD & ".ACCT_CODE)"
            ASCDATA1.ExecuteSQL(sql)

            sql = "Select * from " & GLTFINRD
            'sql = "Select X.*, GLTACCT1.ACCT_DESC from (" & sql & ") X, GLTACCT1 where GLTACCT1.ACCT_CODE = X.ACCT_CODE and AMT_1 <> 0 or AMT_2 <> 0 or AMT_3 <> 0"
            'sql = "Select * from (" & sql & ") where AMT_1 <> 0 or AMT_2 <> 0 or AMT_3 <> 0"
            dst.Tables("GLTFINRD").Rows.Clear()
            dst.EnforceConstraints = False
            Set_SelectCommand("GLTFINRD", sql)
            Call Fill_Records("GLTFINRD")
        End If

        If MODE = "F" Then
            Dim B(6, 3) As Double
            For Each row As DataRow In TBLs("GLTFINRM").Select("STMT_LINE_TYPE = 'D' OR STMT_LINE_TYPE = 'S'", "STMT_LINE_NO", DataViewRowState.CurrentRows)
                STMT_LINE_NO = Val(row.Item("STMT_LINE_NO"))
                If row.Item("STMT_LINE_TYPE") = "D" Then
                    Dim sqlx As String = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
                    row.Item("AMT_1") = Val(TBLs("GLTFINRD").Compute("SUM (AMT_1)", sqlx) & "")
                    row.Item("AMT_2") = Val(TBLs("GLTFINRD").Compute("SUM (AMT_2)", sqlx) & "")
                    row.Item("AMT_3") = Val(TBLs("GLTFINRD").Compute("SUM (AMT_3)", sqlx) & "")
                Else
                    Dim STMT_SUBT_SHOW As Integer = Val(row.Item("STMT_SUBT_SHOW") & "")
                    If STMT_SUBT_SHOW <> 0 Then
                        row.Item("AMT_1") = B(STMT_SUBT_SHOW, 1)
                        row.Item("AMT_2") = B(STMT_SUBT_SHOW, 2)
                        row.Item("AMT_3") = B(STMT_SUBT_SHOW, 3)
                        B(STMT_SUBT_SHOW, 1) = 0
                        B(STMT_SUBT_SHOW, 2) = 0
                        B(STMT_SUBT_SHOW, 3) = 0
                    End If
                End If
                For I As Integer = 1 To 6
                    Dim STMT_SUBT_ADD As Integer = Val(row.Item("STMT_SUBT_ADD" & CStr(I)) & "")
                    If row.Item("STMT_SUBT_ADD" & CStr(I)) & "" = "1" Then
                        B(I, 1) += Val(row.Item("AMT_1") & "")
                        B(I, 2) += Val(row.Item("AMT_2") & "")
                        B(I, 3) += Val(row.Item("AMT_3") & "")
                    End If
                Next

                If row.Item("STMT_LINE_DC") = "C" Then
                    row.Item("AMT_1") = -1 * Val(row.Item("AMT_1") & "")
                    row.Item("AMT_2") = -1 * Val(row.Item("AMT_2") & "")
                    row.Item("AMT_3") = -1 * Val(row.Item("AMT_3") & "")

                    For Each rowGLTFINRD As DataRow In dst.Tables("GLTFINRD").Select("STMT_LINE_NO = " & CStr(STMT_LINE_NO), "", DataViewRowState.CurrentRows)
                        rowGLTFINRD.Item("AMT_1") = -1 * Val(rowGLTFINRD.Item("AMT_1") & "")
                        rowGLTFINRD.Item("AMT_2") = -1 * Val(rowGLTFINRD.Item("AMT_2") & "")
                        rowGLTFINRD.Item("AMT_3") = -1 * Val(rowGLTFINRD.Item("AMT_3") & "")
                    Next
                End If
            Next

            Dim MASK As String = ""
            MASK = "###,##0.00;(###,##0.00);###,##0.00"
            'MASK = "(###,##0.00);###,##0.00;###,##0.00"

            With grdGLTFINRM.DisplayLayout
                .Bands("GLTFINRM").Columns("AMT_1").Format = MASK
                .Bands("GLTFINRM").Columns("AMT_2").Format = MASK
                .Bands("GLTFINRM").Columns("AMT_3").Format = MASK
                .Bands("GLTFINRD").Columns("AMT_1").Format = MASK
                .Bands("GLTFINRD").Columns("AMT_2").Format = MASK
                .Bands("GLTFINRD").Columns("AMT_3").Format = MASK
            End With
            grdGLTFINRM.Refresh()

            Setup_grdGLTFINRM(False)
            'grdGLTFINRM.DisplayLayout.Bands("GLTFINRM").Header.Caption = "Actuals" & " for " & cmbRYP.Text
            '        optMapActBud.ValueList.ValueListItems(optMapActBud.Value)
            grpFetch.Text = optMapActBud.ValueList.ValueListItems(optMapActBud.CheckedIndex).DisplayText & " for " & cmbRYP.Text
        End If

        If MODE = "I" Then
            ASCMAIN1.sql = "Select GLTFINR3.ACCT_CODE, GLTFINR3.SEG2_CODE, GLTFINR3.SEG3_CODE, GLTFINR3.SEG4_CODE, GLTACCT1.ACCT_DESC from GLTFINR3,GLTACCT1 where GLTACCT1.ACCT_CODE = GLTFINR3.ACCT_CODE"

            ASCDATA1.ExecuteSQL("Truncate Table " & GLTACCTL)
            sql = ASCMAIN1.SQL_CodeList(GLTACCTX, "ACCT_CODE", "STMT_LINE_NO")
            sql = "Insert into " & GLTACCTL & " Select X.ACCT_CODE, GLTACCT1.ACCT_DESC, X.STMT_LINE_NOS from GLTACCT1, (" & sql & ") X WHERE X.ACCT_CODE = GLTACCT1.ACCT_CODE"
            ASCDATA1.ExecuteSQL(sql)
            Fill_Records("GLTACCTL")
            dst.Tables("GLTACCTS").Rows.Clear()
            For Each rowGLTACCTL As DataRow In dst.Tables("GLTACCTL").Rows
                For Each SLNO As String In Split(rowGLTACCTL.Item("STMT_LINE_NOS") & "", ",")
                    STMT_LINE_NO = Val(SLNO)
                    Dim ROWGLTFINR2 As DataRow = dst.Tables("GLTFINR2").Rows.Find(New Object() {HFs("STMT_CODE"), STMT_LINE_NO})
                    dst.Tables("GLTACCTS").Rows.Add(New String() {rowGLTACCTL.Item("ACCT_CODE"), STMT_LINE_NO, ROWGLTFINR2.Item("STMT_LINE_DESC")})
                Next
            Next

            'ASCDATA1.ExecuteSQL("Delete from " & GLTFINRX)

            Fill_Records("GLTFINRX")

            Dim sqlY As String = Get_SelectCommand("GLTFINRY")
            Set_SelectCommand("GLTFINRY", sqlY & " AND GLTACCT1.ACCT_TYPE IN " & ACCT_TYPEs(HFs("STMT_TYPE")))
            Fill_Records("GLTFINRY")
            Set_SelectCommand("GLTFINRY", sqlY)

        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub optSEGS_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optSEGS.ValueChanged
        Call Setup_GLTFINR4()
    End Sub

    Sub Setup_GLTFINR4()
        If anode Is Nothing Then
            Exit Sub
        End If
        If optSEGS.Value Is Nothing Then
            Absx1.optFor("STMT_LINE_SEG2_SEL").Visible = False
            Absx1.optFor("STMT_LINE_SEG3_SEL").Visible = False
            Absx1.optFor("STMT_LINE_SEG4_SEL").Visible = False
        Else
            Absx1.optFor("STMT_LINE_SEG2_SEL").Visible = (optSEGS.Value = "2")
            Absx1.optFor("STMT_LINE_SEG3_SEL").Visible = (optSEGS.Value = "3")
            Absx1.optFor("STMT_LINE_SEG4_SEL").Visible = (optSEGS.Value = "4")

            DVWs("GLTFINR4").RowFilter = "STMT_LINE_NO = " & anode.Key & " and ACCT_SEG_ID = '" & optSEGS.Value & "'"
        End If

    End Sub

    Private Sub cmdSEGS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSEGS.Click
        Call Select_SEGS(True)
    End Sub

    Sub Select_SEGS(Optional ByVal set_OptionSet As Boolean = True)
        VIEW_NAME = "SEG" & optSEGS.Value & "_CODE"

        ASCMAIN1.CodeSelector.SQL = _
        ASCMAIN1.CodeSelector.Get_SQL(VIEW_NAME, "", "")

        If ASCMAIN1.CodeSelector.SQL <> "" Then
            ASCMAIN1.CodeSelector.MultipleSelections = True
            Dim Z As String = ""
            For i As Integer = 0 To DVWs("GLTFINR4").Count - 1
                Z &= Chr(0) & DVWs("GLTFINR4")(i).Item("ACCT_SEG_CODE")
            Next
            ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = Z
            Dim F As New ASFCODE1
            F.ShowDialog()
            F.Dispose()

            If ASCMAIN1.CodeSelector.Selections <> 0 Then
                For i As Integer = DVWs("GLTFINR4").Count - 1 To 0 Step -1
                    DVWs("GLTFINR4").Delete(i)
                Next

                For Each row As DataRow In ASCMAIN1.CodeSelector.SelectedRows
                    'Dim rowGLTFINR3 As DataRow = DVWs("GLTFINR3").AddNew.Row

                    Dim rowGLTFINR4 As DataRow = TBLs("GLTFINR4").Rows.Find _
                            (New String() {HFs("STMT_CODE"), anode.Key, optSEGS.Value, row.Item("ACCT_SEG_CODE")})
                    If rowGLTFINR4 Is Nothing Then
                        rowGLTFINR4 = TBLs("GLTFINR4").NewRow
                        With rowGLTFINR4
                            .Item("STMT_CODE") = HFs("STMT_CODE")
                            .Item("STMT_LINE_NO") = anode.Key
                            .Item("ACCT_SEG_ID") = optSEGS.Value
                            .Item("ACCT_SEG_CODE") = row.Item("ACCT_SEG_CODE")
                            .Item("ACCT_SEG_DESC") = row.Item("ACCT_SEG_DESC")
                        End With
                        TBLs("GLTFINR4").Rows.Add(rowGLTFINR4)
                    End If
                    'rowGLTFINR3.EndEdit()

                Next

                ' this should not be necessary - the optionset should be taking care of the underlying data value, but it is not

                Dim STMT_LINE_SEGX_SEL As String = "STMT_LINE_SEG" & optSEGS.Value & "_SEL"
                If DVWs("GLTFINR2")(0).Item(STMT_LINE_SEGX_SEL) = "A" Then
                    DVWs("GLTFINR2")(0).Item(STMT_LINE_SEGX_SEL) = "S"
                End If
                If set_OptionSet Then
                    Select Case optSEGS.Value
                        Case "2"
                            If optSTMT_LINE_SEG2_SEL.CheckedIndex = 0 Then
                                optSTMT_LINE_SEG2_SEL.CheckedIndex = 1
                            End If
                        Case "3"
                            If optSTMT_LINE_SEG3_SEL.CheckedIndex = 0 Then
                                optSTMT_LINE_SEG3_SEL.CheckedIndex = 1
                            End If
                        Case "4"
                            If optSTMT_LINE_SEG4_SEL.CheckedIndex = 0 Then
                                optSTMT_LINE_SEG4_SEL.CheckedIndex = 1
                            End If
                    End Select
                End If
            End If
        End If
    End Sub

    Private Sub optSTMT_LINE_SEGS_SEL_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSTMT_LINE_SEG2_SEL.ValueChanged, optSTMT_LINE_SEG3_SEL.ValueChanged, optSTMT_LINE_SEG4_SEL.ValueChanged

        If setting_up Then
            Exit Sub
        End If

        Dim opt As UltraWinEditors.UltraOptionSet = DirectCast(sender, UltraWinEditors.UltraOptionSet)
        If opt.CheckedIndex = 0 Then
            For i As Integer = DVWs("GLTFINR4").Count - 1 To 0 Step -1
                DVWs("GLTFINR4").Delete(i)
            Next
        Else
            If grdGLTFINR4.Rows.Count = 0 Then
                Call Select_SEGS(False)
                If grdGLTFINR4.Rows.Count = 0 Then
                    opt.CheckedIndex = 0
                End If
            End If
        End If
    End Sub

    Private Sub grdGLTFINR4_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTFINR4.AfterRowsDeleted
        If grdGLTFINR4.Rows.Count = 0 Then

            ' this should not be necessary - the optionset should be taking care of the underlying data value, but it is not

            Dim STMT_LINE_SEGX_SEL As String = "STMT_LINE_SEG" & optSEGS.Value & "_SEL"
            DVWs("GLTFINR2")(0).Item(STMT_LINE_SEGX_SEL) = "A"

            If optSEGS.Value = "2" Then
                optSTMT_LINE_SEG2_SEL.CheckedIndex = 0

            End If
            If optSEGS.Value = "3" Then
                optSTMT_LINE_SEG3_SEL.CheckedIndex = 0
            End If
            If optSEGS.Value = "4" Then
                optSTMT_LINE_SEG4_SEL.CheckedIndex = 0
            End If
        End If
    End Sub

    Private Sub grdGLTFINR3_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdGLTFINR3.Leave
        If grdGLTFINR3.ActiveRow Is Nothing Then
            Exit Sub
        End If
        grdGLTFINR3.RowUpdateCancelAction = UltraWinGrid.RowUpdateCancelAction.CancelUpdate

        If grdGLTFINR3.ActiveRow.DataChanged Then
            Try
                grdGLTFINR3.UpdateData()
            Catch ex As Exception
                grdGLTFINR3.ActiveRow.CancelUpdate()
            End Try
        End If
        grdGLTFINR3.RowUpdateCancelAction = UltraWinGrid.RowUpdateCancelAction.RetainDataAndActivation
    End Sub

    Private Sub grdGLTFINRM_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdGLTFINRM.ClickCellButton

        If Not sub_total_map_generated_and_no_changes_have_been_made_yet Then
            MsgBox("Changes have been made to the Format - Please Refresh Sub-Total Map", MsgBoxStyle.OkOnly, "Cannot Continue with Edits until Refreshed")
            Exit Sub
        End If

        If grdGLTFINRM.ActiveRow.Cells("STMT_LINE_TYPE").Text = "H" Then
            Exit Sub
        End If

        Dim ADDR As String
        If grdGLTFINRM.ActiveRow.Cells("STMT_LINE_TYPE").Text = "D" Then
            ADDR = "+"
        Else
            ADDR = "+=*"
        End If
        Dim REG As String = grdGLTFINRM.ActiveCell.Column.Key
        REG = Mid(REG, Len(REG), 1)

        Dim i As Integer = InStr(ADDR, e.Cell.Text)
        If Trim(e.Cell.Text) = "" Then i = 0
        Dim new_value As String = Mid(ADDR & " ", i + 1, 1)
        grdGLTFINRM.ActiveCell.Value = new_value
        Dim row As DataRow = dst.Tables("GLTFINR2").Rows.Find(New Object() {HFs("STMT_CODE"), grdGLTFINRM.ActiveRow.Cells("STMT_LINE_NO2").Text})
        If grdGLTFINRM.ActiveRow.Cells("STMT_LINE_TYPE").Text = "D" Then
            If new_value = "+" Then
                row.Item("STMT_SUBT_ADD" & REG) = "1"
            Else
                row.Item("STMT_SUBT_ADD" & REG) = ""
            End If
        ElseIf grdGLTFINRM.ActiveRow.Cells("STMT_LINE_TYPE").Text = "S" Then
            If new_value = "+" Then
                row.Item("STMT_SUBT_ADD" & REG) = "1"
                If row.Item("STMT_SUBT_SHOW") = REG Then
                    grdGLTFINRM.ActiveRow.Cells("STMT_SUBT_SHOW").Value = ""
                    row.Item("STMT_SUBT_SHOW") = ""
                End If
            ElseIf new_value = "=" Then
                row.Item("STMT_SUBT_ADD" & REG) = "1"
                Dim REG_OLD As String = row.Item("STMT_SUBT_SHOW") & ""
                If REG_OLD <> "" And REG_OLD <> REG And REG_OLD <> "0" Then
                    grdGLTFINRM.ActiveRow.Cells("REG_" & REG_OLD).Value = ""
                End If
                grdGLTFINRM.ActiveRow.Cells("STMT_SUBT_SHOW").Value = REG
                row.Item("STMT_SUBT_SHOW") = REG
            ElseIf new_value = "*" Then
                row.Item("STMT_SUBT_ADD" & REG) = ""
                Dim REG_OLD As String = row.Item("STMT_SUBT_SHOW") & ""
                If REG_OLD <> "" And REG_OLD <> REG And REG_OLD <> "0" Then
                    grdGLTFINRM.ActiveRow.Cells("REG_" & REG_OLD).Value = ""
                End If
                row.Item("STMT_SUBT_SHOW") = REG
                grdGLTFINRM.ActiveRow.Cells("STMT_SUBT_SHOW").Value = REG
            ElseIf new_value = " " Then
                row.Item("STMT_SUBT_ADD" & REG) = ""
                If row.Item("STMT_SUBT_SHOW") = REG Then
                    row.Item("STMT_SUBT_SHOW") = ""
                    grdGLTFINRM.ActiveRow.Cells("STMT_SUBT_SHOW").Value = ""
                End If
            End If
        End If

        ' grdGLTFINRM.ActiveRow.Update()
    End Sub

    Private Sub grdGLTFINRM_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTFINRM.InitializeLayout

        With grdGLTFINRM.DisplayLayout
            .Bands("GLTFINRM").Columns("STMT_LINE_NO").Header.Fixed = True
            .Bands("GLTFINRM").Columns("DESCRIPTION").Header.Fixed = True
            .Bands("GLTFINRD").Columns("ACCT_CODE").Header.Fixed = True
            .Bands("GLTFINRD").Columns("ACCT_DESC").Header.Fixed = True
        End With

    End Sub

    Sub Activate_Node()
        setting_up = True
        anode = tvwGLTFINR2.ActiveNode
        'If tvwGLTFINR2.Nodes.Count = 0 Then Exit Sub
        If anode Is Nothing Then Exit Sub
        Dim STMT_LINE_NO As Integer = Val(tvwGLTFINR2.ActiveNode.Key)
        DVWs("GLTFINR2").RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)
        DVWs("GLTFINR3").RowFilter = "STMT_LINE_NO = " & CStr(STMT_LINE_NO)

        Me.BindingContext(TBLs("GLTFINR2")).Position = _
            TBLs("GLTFINR2").Rows.IndexOf(TBLs("GLTFINR2").Rows.Find _
            (New String() {HFs("STMT_CODE"), CStr(STMT_LINE_NO)}))

        Setup_GLTFINR4()
        setting_up = False
    End Sub

    Private Sub tvwGLTFINR2_AfterSelect(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTree.SelectEventArgs) Handles tvwGLTFINR2.AfterSelect

    End Sub

    Private Sub grdGLTACCTL_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTACCTL.DoubleClickRow
        If e.Row.Band.Key = "GLTACCTL_GLTACCTS" Then
            Dim STMT_LINE_NO As Int64 = Val(e.Row.Cells("STMT_LINE_NO").Value & "")
            tvwGLTFINR2.GetNodeByKey(CStr(STMT_LINE_NO))
            Dim node As UltraWinTree.UltraTreeNode = tvwGLTFINR2.GetNodeByKey(CStr(STMT_LINE_NO))
            If node IsNot Nothing Then
                tvwGLTFINR2.ActiveNode = tvwGLTFINR2.GetNodeByKey(CStr(STMT_LINE_NO))
                UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Line Definition")
            End If

        End If
    End Sub

    Private Sub grdGLTACCTL_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTACCTL.InitializeLayout

    End Sub

    Private Sub cbeSTMT_CODE_COPY_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbeSTMT_CODE_COPY.AfterCloseUp
        Dim STMT_CODE_COPY As String = cbeSTMT_CODE_COPY.Value & ""
        If STMT_CODE_COPY <> "" Then
            Dim rowGLTFINR1 As DataRow = LookUp("GLTFINR1", STMT_CODE_COPY)
            If rowGLTFINR1.Item("STMT_TYPE") & "" <> Absx1.optFor("STMT_TYPE").Value Then
                MsgBox("Non-Matching Statement Type", MsgBoxStyle.OkOnly, "Cannot Copy from Statement " & STMT_CODE_COPY)
            Else
                If MsgBox("OK to lose all changes and copy definition from Statement " & STMT_CODE_COPY, _
                          MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

                    'Dim X As CurrencyManager = Me.BindingContext(dst.Tables("GLTFINR1"))
                    'X.EndCurrentEdit()

                    Absx1.txtFor("STMT_DESC").Text = rowGLTFINR1.Item("STMT_DESC")
                    tvwGLTFINR2.Visible = False

                    For Each TABLE_NAME As String In New String() {"GLTFINR2", "GLTFINR3", "GLTFINR4"}
                        Fill_Records(TABLE_NAME, STMT_CODE_COPY)
                        For Each row As DataRow In dst.Tables(TABLE_NAME).Select
                            row.Item("STMT_CODE") = HFs("STMT_CODE")
                            row.AcceptChanges()
                            row.SetAdded()
                        Next
                    Next

                    Setup_Tree()
                    tvwGLTFINR2.Visible = True

                End If
            End If
        End If

        cbeSTMT_CODE_COPY.Value = ""
        cbeSTMT_CODE_COPY.Visible = False
        lblSTMT_CODE_COPY.Visible = False

    End Sub

    Sub Setup_Tree()

        Dim N() As UltraWinTree.UltraTreeNode
        ReDim N(0)

        STMT_LINE_NO_ctr = 0
        tvwGLTFINR2.Nodes.Clear()
        For Each rowGLTFINR2 As DataRow In dst.Tables("GLTFINR2").Select("", "STMT_LINE_NO")
            Dim STMT_LINE_LEVEL As Integer = Val(rowGLTFINR2.Item("STMT_LINE_LEVEL") & "")
            STMT_LINE_NO_ctr = rowGLTFINR2.Item("STMT_LINE_NO")
            If STMT_LINE_LEVEL > UBound(N) Then
                ReDim Preserve N(STMT_LINE_LEVEL)
            End If
            If STMT_LINE_LEVEL = 1 Then
                anode = tvwGLTFINR2.Nodes.Add(rowGLTFINR2.Item("STMT_LINE_NO"), rowGLTFINR2.Item("STMT_LINE_DESC"))
            Else
                anode = N(STMT_LINE_LEVEL - 1).Nodes.Add(rowGLTFINR2.Item("STMT_LINE_NO"), rowGLTFINR2.Item("STMT_LINE_DESC") & "")
            End If
            N(STMT_LINE_LEVEL) = anode
            Call Setup_Node(rowGLTFINR2.Item("STMT_LINE_TYPE"))
        Next

        If tvwGLTFINR2.Nodes.Count > 0 Then
            tvwGLTFINR2.ActiveNode = tvwGLTFINR2.Nodes(0)
            'Call Activate_Node()
        End If

        grdGLTFINR3.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdGLTFINR3.DisplayLayout.Bands(0).SortedColumns.Add("ACCT_CODE", False)

        Setup_grdGLTFINRM(True)
        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Line Definition")
        Setup_Tab()
    End Sub
End Class