Imports Infragistics.Win.UltraWinGrid

Public Class GLFACTI1
    Dim ACCT_CODE_last As String = ""
    Dim ACCT_CODE_nav As String = ""
    Dim rowGLTACCT1 As DataRow
    Dim GLTACCTU As String = ""
    Dim sqlX As String = ""
    Dim sqlWhereClause As String = "{sql where}" & vbCrLf

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")

        With dst

            .Tables.Add("GLTYEAR1")
            With .Tables("GLTYEAR1")
                .Columns.Add("ACCT_YEAR", GetType(System.String))
                .Columns.Add("SEG2_CODE", GetType(System.String))
                .Columns.Add("SEG3_CODE", GetType(System.String))
                .Columns.Add("SEG4_CODE", GetType(System.String))
                .PrimaryKey = New DataColumn() { .Columns("ACCT_YEAR"), .Columns("SEG2_CODE"), .Columns("SEG3_CODE"), .Columns("SEG4_CODE")}
                .Columns.Add("ACCT_BEG_BAL", GetType(System.Decimal))
                .Columns.Add("ACCT_END_BAL", GetType(System.Decimal))
                .Columns.Add("ACCT_NET_ACT", GetType(System.Decimal))
                .Columns.Add("ACCT_TRANSACTIONS", GetType(System.Int64))

                .Columns("ACCT_END_BAL").Expression = "ISNULL(ACCT_BEG_BAL,0) + ISNULL(ACCT_NET_ACT,0)"

            End With

            Create_GLTACCTU(True)
            ASCMAIN1.sql = "Select GLTACCT1.*, X.JNO_MIN, X.JNO_MAX, X.TRANS, X.YP_MIN, X.YP_MAX" & vbCrLf _
                & " from GLTACCT1, (Select ACCT_CODE, MIN(JOURNAL_NO) JNO_MIN, MAX(JOURNAL_NO) JNO_MAX, MIN (OPS_YYYYPP) YP_MIN, MAX (OPS_YYYYPP) YP_MAX, COUNT (*) TRANS" & vbCrLf _
                & " from GLTDETL1 group by ACCT_CODE) X" & vbCrLf _
                & " where X.ACCT_CODE (+) = GLTACCT1.ACCT_CODE"
            Create_TDA(.Tables.Add, "GLTACCTX", "**", 0, False, "", 1)
            With .Tables("GLTACCTX")
                .Columns("TRANS").DataType = GetType(System.Int64)
                .Columns.Add("USES", GetType(System.Int32))
                .Columns.Add("USES_DESC")
            End With
            ASCMAIN1.sql = "Select GLTACCTU.*, X.MENU_ITEM_DESC" & vbCrLf _
                & " from " & GLTACCTU & " GLTACCTU, (Select MENU_ITEM_OBJECT, MIN (MENU_ITEM_DESC) MENU_ITEM_DESC from ASTMENU1 group by MENU_ITEM_OBJECT) X" & vbCrLf _
                & " where X.MENU_ITEM_OBJECT (+) = GLTACCTU.TBL"
            Create_TDA(.Tables.Add, "GLTACCTU", "**", 0, False, "", 0)
            Create_Relation("GLTACCTX", "GLTACCTU", "ACCT_CODE")


            Create_TDA(.Tables.Add, "GLTACCT1", "*", -1, False)
            Create_TDA(.Tables.Add, "GLTACCT3", "*", 1, False)

            ASCMAIN1.sql = "Select GLTACCT3.ACCT_YEAR, GLTACCT3.ACCT_CODE, GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE, GLTJRNL1.JOURNAL_TYPE from GLTACCT3,GLTJRNL1"
            Create_TDA(.Tables.Add, "GLTACCT3_TYPE", "**", 0, False, "", 6)

            With .Tables("GLTACCT3_TYPE")
                For P As Integer = 1 To 12
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Decimal))
                Next
            End With

            Create_Relation("GLTACCT3", "GLTACCT3_TYPE", "ACCT_YEAR,ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE")

            ASCMAIN1.sql = "Select DISTINCT SUBSTR(GLTPARM2.OPS_YYYYPP,1,4) ACCT_YEAR" _
                & ", GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE" _
                & ", GLTPARM2.* from GLTPARM2,GLTACCT3 "
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "", 5)
            With .Tables("GLTPARM2")
                .Columns.Add("ACCT_BEG_BAL", GetType(System.Decimal))
                .Columns.Add("ACCT_END_BAL", GetType(System.Decimal))
                .Columns.Add("ACCT_NET_ACT", GetType(System.Decimal))
                .Columns.Add("ACCT_TRANSACTIONS", GetType(System.Int64))

                .Columns("ACCT_END_BAL").Expression = "ISNULL(ACCT_BEG_BAL,0) + ISNULL(ACCT_NET_ACT,0)"

            End With

            ASCMAIN1.sql = "Select GLTDETL1.*, GLTJRNL1.JOURNAL_DESC, GLTJRNL1.JOURNAL_TYPE" _
                & ", GLTDETL1.SEG2_CODE SEG2_CODE_JOIN, GLTDETL1.SEG3_CODE SEG3_CODE_JOIN, GLTDETL1.SEG4_CODE SEG4_CODE_JOIN" _
                & " from GLTDETL1,GLTJRNL1 where GLTDETL1.ACCT_CODE = :PARM1 " _
                & " and GLTDETL1.OPS_YYYYPP >= :PARM2 AND GLTDETL1.OPS_YYYYPP <= :PARM3" _
                & " and GLTJRNL1.JOURNAL_NO (+) = GLTDETL1.JOURNAL_NO"
            Create_TDA(.Tables.Add, "GLTDETL1", "**", 0, False, "VVV", 3)

            Create_Relation("GLTYEAR1", "GLTPARM2", "ACCT_YEAR,SEG2_CODE,SEG3_CODE,SEG4_CODE")
            Create_Relation("GLTPARM2", "GLTDETL1", "OPS_YYYYPP,SEG2_CODE,SEG3_CODE,SEG4_CODE", "OPS_YYYYPP,SEG2_CODE_JOIN,SEG3_CODE_JOIN,SEG4_CODE_JOIN")

            .Tables("GLTPARM2").Columns("ACCT_NET_ACT").Expression = "SUM(CHILD.DETL_POSTING_AMT)"
            .Tables("GLTPARM2").Columns("ACCT_TRANSACTIONS").Expression = "COUNT(CHILD.OPS_YYYYPP)"

            .Tables("GLTYEAR1").Columns("ACCT_NET_ACT").Expression = "SUM(CHILD.ACCT_NET_ACT)"
            .Tables("GLTYEAR1").Columns("ACCT_TRANSACTIONS").Expression = "SUM(CHILD.ACCT_TRANSACTIONS)"

            .Tables.Add(ASCDATA1.GetDataTable("Select ACCT_SEG_CODE, ACCT_SEG_DESC, '0' SEL from GLTSEGM1 where ACCT_SEG_ID = '2'", "GLTSEGM2", 1, True))
            .Tables("GLTSEGM2").Columns("SEL").ReadOnly = False
            .Tables.Add(ASCDATA1.GetDataTable("Select ACCT_SEG_CODE, ACCT_SEG_DESC, '0' SEL from GLTSEGM1 where ACCT_SEG_ID = '3'", "GLTSEGM3", 1, True))
            .Tables("GLTSEGM3").Columns("SEL").ReadOnly = False
            .Tables.Add(ASCDATA1.GetDataTable("Select ACCT_SEG_CODE, ACCT_SEG_DESC, '0' SEL from GLTSEGM1 where ACCT_SEG_ID = '4'", "GLTSEGM4", 1, True))
            .Tables("GLTSEGM4").Columns("SEL").ReadOnly = False


            Create_TDA(.Tables.Add, "GLTTYPE1", "*", 0, False)
            Fill_Records("GLTTYPE1")
        End With

        grdSEG2.DataSource = dst.Tables("GLTSEGM2")
        grdSEG2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSEG2.DisplayLayout.Bands(0).SortedColumns.Add("ACCT_SEG_CODE", False)
        grdSEG3.DataSource = dst.Tables("GLTSEGM3")
        grdSEG3.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSEG3.DisplayLayout.Bands(0).SortedColumns.Add("ACCT_SEG_CODE", False)
        grdSEG4.DataSource = dst.Tables("GLTSEGM4")
        grdSEG4.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSEG4.DisplayLayout.Bands(0).SortedColumns.Add("ACCT_SEG_CODE", False)

        grdGLTACCT3.DataSource = dst.Tables("GLTACCT3")
        grdGLTDETL1.DataSource = dst.Tables("GLTYEAR1")
        grdGLTACCTX.DataSource = dst.Tables("GLTACCTX")

        With grdGLTACCTX.DisplayLayout.Bands(0)
            .Columns("TRANS").Format = "#,##0"
            .Columns("ACCT_CODE").Header.Fixed = True
            .Columns("ACCT_STATUS").Header.Fixed = True
            .Columns("ACCT_DESC").Header.Fixed = True
        End With
        Create_Summary(grdGLTACCTX, "ACCT_CODE", "Count")


        With grdGLTACCT3.DisplayLayout.Bands(0)
            For i As Integer = 1 To 12
                Dim z As String = "ACCT_ACT_P" & Format$(i, "00")
                Dim zz As String = ASCMAIN1.Get_Legend(Mid$(ASCMAIN1.CYP, 1, 4) & Format$(i, "00"))
                .Columns(z).Header.Caption = Mid$(zz, 10, 3)
                .Columns(z).Format = "###,##0.00"
            Next
            .Columns("ACCT_BEG_BAL").Format = "###,##0.00"
        End With

        Create_Summary(grdGLTACCT3, "ACCT_BEG_BAL")
        For i As Integer = 1 To 12
            Create_Summary(grdGLTACCT3, "ACCT_ACT_P" & Format(i, "00"))
        Next
        For i As Integer = 1 To 12
            Create_Summary(grdGLTACCT3, "P" & Format(i, "00"), , "GLTACCT3_GLTACCT3_TYPE")
        Next

        grdGLTACCT3.DisplayLayout.Bands("GLTACCT3_GLTACCT3_TYPE").SummaryFooterCaption = "Activity Totals for: [ACCT_YEAR]"

        Create_Summary(grdGLTDETL1, "DETL_POSTING_AMT", , "GLTPARM2_GLTDETL1")
        Create_Summary(grdGLTDETL1, "JOURNAL_NO", "Count", "GLTPARM2_GLTDETL1")

        Create_Summary(grdGLTDETL1, "ACCT_NET_ACT", , "GLTYEAR1")
        Create_Summary(grdGLTDETL1, "ACCT_TRANSACTIONS", , "GLTYEAR1")
        Create_Summary(grdGLTDETL1, "ACCT_NET_ACT", , "GLTYEAR1_GLTPARM2")
        Create_Summary(grdGLTDETL1, "ACCT_TRANSACTIONS", , "GLTYEAR1_GLTPARM2")

        For i As Integer = 2 To 4
            COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
            With grdGLTDETL1.DisplayLayout
                .Bands("GLTYEAR1").Columns(COLUMN_NAME).Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & ""
                .Bands("GLTYEAR1_GLTPARM2").Columns(COLUMN_NAME).Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & ""
                .Bands("GLTPARM2_GLTDETL1").Columns(COLUMN_NAME).Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC") & ""
                .Bands("GLTYEAR1").Columns(COLUMN_NAME).Width = 60
                .Bands("GLTYEAR1_GLTPARM2").Columns(COLUMN_NAME).Width = 60
                .Bands("GLTPARM2_GLTDETL1").Columns(COLUMN_NAME).Width = 60
            End With
        Next

        grdGLTDETL1.DisplayLayout.Bands("GLTYEAR1").SummaryFooterCaption = "Totals"
        grdGLTDETL1.DisplayLayout.Bands("GLTYEAR1_GLTPARM2").SummaryFooterCaption = "Year Totals"
        grdGLTDETL1.DisplayLayout.Bands("GLTPARM2_GLTDETL1").SummaryFooterCaption = "Period Totals"

        Breakout_By()

        'Set_SEGs(grdGLTACCT3, "GLTACCT3")
        Set_SEGS(grdGLTDETL1, "GLTPARM2_GLTDETL1")

        With tabSEGs
            For I As Integer = 2 To 4
                Dim Z As String = "SEG" & CStr(I)
                If ROWs("GLTPARM1").Item("GL_PARM_" & Z & "_DESC") & "" = "" Then
                    .Tabs(I - 2).Visible = False
                Else
                    .Tabs(I - 2).Text = ROWs("GLTPARM1").Item("GL_PARM_" & Z & "_DESC")
                End If
            Next
        End With

        optRed.Value = ROWs("GLTPARM1").Item("GL_PARM_SHOW_RED")
        'optRed.Value = "C" ' strange how this seems to work here in form load but not always
        optWith.Value = ROWs("GLTPARM1").Item("GL_PARM_SHOW_WITH")

        Dim CURR_YEAR = Mid(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 1, 4)
        Absx1.chkFor("NEXT_YEAR").Text = Format(Val(CURR_YEAR) + 1, "0000")
        Absx1.chkFor("CURR_YEAR").Text = CURR_YEAR
        Absx1.chkFor("LAST_YEAR").Text = Format(Val(CURR_YEAR) - 1, "0000")
        Absx1.chkFor("PRIOR_YEARS").Text = "Prior" ' "<" & Format(Val(CURR_YEAR) - 1, "0000")
        If CURR_YEAR < Mid(ASCMAIN1.CYP, 1, 4) Then
            Absx1.chkFor("NEXT_YEAR").Checked = True
        End If

        'ASCMAIN1.Add_Value_List(grdGLTACCTX, "ACCT_SEG2_MAND")
        'ASCMAIN1.Add_Value_List(grdGLTACCTX, "ACCT_SEG3_MAND")
        'ASCMAIN1.Add_Value_List(grdGLTACCTX, "ACCT_SEG4_MAND")

        If ASCMAIN1.CLIENT = "AHA" Then
            chkSEG2_CODE.Checked = False
            chkSEG3_CODE.Checked = False
            chkSEG4_CODE.Checked = False
        End If
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("ACCT_CODE")

            Case "Next Account", "Previous Account"
                If eItemKey = "Next Account" Then
                    ASCMAIN1.sql = "Select * from (Select ACCT_CODE from GLTACCT1 where ACCT_CODE > '" & ACCT_CODE_last.PadRight(6, "0") & "' order by ACCT_CODE) where ROWNUM <2"
                Else
                    ASCMAIN1.sql = "Select * from (Select ACCT_CODE from GLTACCT1 where ACCT_CODE < '" & ACCT_CODE_last.PadRight(6, "0") & "' order by ACCT_CODE DESC) where ROWNUM <2"
                    ASCMAIN1.sql = "Select * from (Select ACCT_CODE from GLTACCT1 where ACCT_CODE < '" & ACCT_CODE_last & "' order by ACCT_CODE DESC) where ROWNUM <2"
                End If
                ACCT_CODE_nav = ASCDATA1.GetDataValue
                If ACCT_CODE_nav = "" Then
                    EMsg &= vbCr & "Could Not Determine the " & eItemKey
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
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                'grdGLTACCT3.DisplayLayout.Override.SummaryDisplayArea = UltraWinGrid.SummaryDisplayAreas.InGroupByRows _
                '                                                      + UltraWinGrid.SummaryDisplayAreas.GroupByRowsFooter
                'Exit Sub

                Mode_Settings(False)

            Case "Next Account", "Previous Account"
                Click_Command("Done")
                Absx1.txtFor("ACCT_CODE").Text = ACCT_CODE_nav
                Click_Command("Load")

            Case "Generate Summary"

                Create_GLTACCTU(False)
                EnforceConstraints(False)
                Fill_Records("GLTACCTU")
                Fill_Records("GLTACCTX")


                For Each rowGLTACCTU As DataRow In dst.Tables("GLTACCTU").Select("")
                    Dim USE_DESC As String = rowGLTACCTU.Item("TBL") & ":" & rowGLTACCTU.Item("COL") & ":" & rowGLTACCTU.Item("KEYS")

                    Dim ACCT_CODE As String = rowGLTACCTU.Item("ACCT_CODE")
                    Dim rowGLTACCTX As DataRow = dst.Tables("GLTACCTX").Rows.Find(ACCT_CODE)
                    If rowGLTACCTX Is Nothing Then
                        rowGLTACCTU.Item("KEYS") = DBNull.Value
                        ' PROB SHOULD MAKE A MENTION TO USER
                    Else
                        rowGLTACCTX.Item("USES") = Val(rowGLTACCTX.Item("USES") & "") + 1
                        If Val(rowGLTACCTX.Item("USES")) > 1 Then
                            rowGLTACCTX.Item("USES_DESC") &= vbCrLf
                        End If
                        rowGLTACCTX.Item("USES_DESC") = rowGLTACCTX.Item("USES_DESC") & USE_DESC
                    End If
                Next
                ASCDATA1.DeleteRows("GLTACCTU", "ISNULL(KEYS,'') = ''")

                EnforceConstraints(True)

                Sort_grdColumns(grdGLTACCTX, "ACCT_CODE")
                Sort_grdColumns(grdGLTACCTX, "TBL,COL", , 1)
                ' grdGLTACCTX.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand
                grdGLTACCTX.DisplayLayout.Override.ExpansionIndicator = UltraWinGrid.ShowExpansionIndicator.CheckOnDisplay

                grdGLTACCTX.Visible = True

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Generate Summary").Visible = Not ScreenMode
                End With

                .Groups("Account Segments").Visible = tf
                .Groups("Options").Visible = False ' tf
                .Groups("Breakout By").Visible = True ' tf
            End With
        End If

        Setup_tabGLTACCT1()
        Setup_Details()

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        'bsx1.txtFor("JOURNAL_DESC").ReadOnly = False

        grpGLTACCT1.Visible = tf
        'Absx1.chkFor("ACCT_SUB_CTL").Visible = tf
        'Absx1.chkFor("ACCT_POST_SUMMARY").Visible = tf

        grdGLTACCTX.Visible = Not tf And dst IsNot Nothing AndAlso (dst.Tables("GLTACCTX").Rows.Count > 0)
        tabGLTACCT1.Visible = tf

        If ScreenMode Then
            MENU_ITEM_DESC = "Account Inquiry - " & HFs("ACCT_CODE")
        Else
            MENU_ITEM_DESC = "Account Inquiry"
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"GLTACCT1", "GLTACCT3", "GLTACCT3_TYPE", "GLTYEAR1", "GLTPARM2", "GLTDETL1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If ASCMAIN1.CLIENT = "AHA" Then
        Else
            If chkSEG2_CODE.Visible Then chkSEG2_CODE.Checked = True
            If chkSEG3_CODE.Visible Then chkSEG3_CODE.Checked = True
            If chkSEG4_CODE.Visible Then chkSEG4_CODE.Checked = True
        End If

        grdGLTDETLX.Tag = ""

        'Absx1.txtFor("ACCT_CODE").Text = ""
        'optShowBA.Value = "A"
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        rowGLTACCT1 = Fill_Record("GLTACCT1", HFs("ACCT_CODE"))

        Select Case rowGLTACCT1("ACCT_TYPE")
            Case "A"
                lblACCT_TYPE.Text = "Type: " & "Asset"
            Case "L"
                lblACCT_TYPE.Text = "Type: " & "Liability"
            Case "E"
                lblACCT_TYPE.Text = "Type: " & "Equity"
            Case "I"
                lblACCT_TYPE.Text = "Type: " & "Income"
            Case "X"
                lblACCT_TYPE.Text = "Type: " & "Expense"
        End Select

        ACCT_CODE_last = HFs("ACCT_CODE")

        Show_With()

        Load_Data()
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Load"
                Absx1.txtFor("ACCT_CODE").Text = key
                Click_Command(command)
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "GLTACCT1"
            E.COLUMN_NAME = "ACCT_CODE"
            E.CODE_VALUE = Absx1.txtFor("ACCT_CODE").Text
            E.DESC_VALUE = "GL Account"
            E.ATTACHMENT_NOTES = ""
            'If rowSOTORDR1.Item("STATUS") & "" <> "0" Then
            '    E.RESTRICTIONS = "D"
            'End If
            'E.READ_ONLY = True
        End If

        Return E
    End Function

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "ACCT_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "ACCT_CODE"
                Call Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub chk_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        MyBase.chk_CheckedChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "NEXT_YEAR", "CURR_YEAR", "LAST_YEAR", "PRIOR_YEARS"
                If ScreenMode Then Set_Year_Filters()
        End Select
    End Sub
#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTACCTX, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Show Record")
        Load_Popup_Menu(grdGLTDETL1, "SBBB", "Show Filter", "Account Inquiry", "Show Journal", "Voucher Inquiry")
        Load_Popup_Menu(grdGLTACCT3, "SSS", "Show Filter", "Show GroupBox", "Show Pins", "Show Record")

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
            Case "grdGLTDETL1"
                tlb_btn = DirectCast(tlb_pop.Tools("Show Journal"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = False
                tlb_btn = DirectCast(tlb_pop.Tools("Account Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = False
                tlb_btn = DirectCast(tlb_pop.Tools("Voucher Inquiry"), UltraWinToolbars.ButtonTool)
                tlb_btn.SharedProps.Visible = False
        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

            Select Case e.SourceControl.Name
                Case "grdGLTDETL1"
                    If grd.ActiveRow.Band.Key = "GLTPARM2_GLTDETL1" Then
                        tlb_btn = DirectCast(tlb_pop.Tools("Show Journal"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = True
                    End If
                    If grd.ActiveRow.Band.Key = "GLTPARM2_GLTDETL1" Then
                        tlb_btn = DirectCast(tlb_pop.Tools("Account Inquiry"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = True
                    End If
                    If grd.ActiveRow.Band.Key = "GLTPARM2_GLTDETL1" Then
                        If grd.ActiveRow.Cells("JOURNAL_TYPE").Value & "" = "APIN" Then
                            tlb_btn = DirectCast(tlb_pop.Tools("Voucher Inquiry"), UltraWinToolbars.ButtonTool)
                            tlb_btn.SharedProps.Visible = True
                        End If
                    End If
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            'Case "Select Selected", "De-Select Selected"
            '    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
            '        grow.Cells("TRAN_SEL").Value = IIf(e.Tool.Key = "De-Select Selected", "0", "1")
            '        grow.Update()
            '    Next

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            '    If rowICTITEM1 IsNot Nothing Then
            '        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
            '    End If

            Case "Show Record"
                Dim MENU_ITEM_OBJECT As String = ""
                Dim MENU_ITEM_DESC As String = ""
                Dim COLUMN_NAME_KEY As String = ""
                If grd.ActiveRow.Band.Key = "GLTACCTX" Then
                Else
                    MENU_ITEM_OBJECT = grd.ActiveRow.Cells("TBL").Value
                    MENU_ITEM_DESC = grd.ActiveRow.Cells("MENU_ITEM_DESC").Value & ""
                    COLUMN_NAME_KEY = grd.ActiveRow.Cells("TABLE_KEY").Value & ""
                End If
                If MENU_ITEM_OBJECT <> "" And MENU_ITEM_DESC <> "" Then
                    Dim KEYS As String = grd.ActiveRow.Cells("KEYS").Value & ""
                    If KEYS <> "" Then
                        Dim KEY As String = Split(KEYS, ",")(0)
                        Dim KEYfm As New Dictionary(Of String, Object)
                        KEYfm.Add(COLUMN_NAME_KEY, KEY)
                        Context_Launch("View", KEYfm, MENU_ITEM_DESC, MENU_ITEM_OBJECT)
                    End If

                End If

            Case "Account Inquiry"
                Dim ACCT_CODE As String = grd.ActiveRow.Cells("ACCT_CODE").Value
                Context_Launch("Load", ACCT_CODE, e.Tool.Key, "GLFACTI1")

            Case "Show Journal"
                Dim JOURNAL_NO As String = grd.ActiveRow.Cells("JOURNAL_NO").Value
                Dim rowGLTJRNL1 As DataRow = LookUp("GLTJRNL1", JOURNAL_NO)
                If rowGLTJRNL1 IsNot Nothing Then
                    Context_Launch("View", JOURNAL_NO, e.Tool.Key, "GLFJRNL1")
                End If

            Case "Voucher Inquiry"
                Dim VOUCHER_NO As String = grd.ActiveRow.Cells("DETL_CTL_NO").Value & ""
                If VOUCHER_NO <> "" Then
                    Context_Launch("Load", VOUCHER_NO, e.Tool.Key, "APFINVHI")
                End If

        End Select
    End Sub

#End Region

    Sub Load_Data()
        If SELECTION_NO = 0 Then Exit Sub
        'If dst.Tables.Count = 0 Or EntryMode = "" Then
        '    Exit Sub
        'End If

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Activity")

        'EnforceConstraints(False)
        Load_GLTACCT3()
        Load_GLTDETL1()
        'EnforceConstraints(True)
        Set_Year_Filters()

        ' Show_Red() ' this has been disabled because it makes load times excessive (like acct 311000 at INT)
        Show_With()

        ASCMAIN1.Progress("")
        Me.Cursor = Cursors.Default
    End Sub

    Sub Load_GLTACCT3()

        EnforceConstraints(False)

        If dst.Tables.Count = 0 Or EntryMode = "" Then
            Exit Sub
        End If

        Dim sql As String = "Select ACCT_CODE"
        Dim gby As String = ""
        Dim COLUMN_NAME As String
        For i As Integer = 2 To 4
            COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
            If Absx1.chkFor(COLUMN_NAME).Checked Then
                sql = sql & ", " & COLUMN_NAME & " " & COLUMN_NAME
                gby = gby & ", " & COLUMN_NAME
                grdGLTACCT3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = False
                grdGLTACCT3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC")
                grdGLTACCT3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Width = 60
            Else
                sql = sql & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "' " & COLUMN_NAME
                grdGLTACCT3.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
            End If
        Next
        sql = sql & ", ACCT_YEAR"
        For i As Integer = 0 To 12
            If i = 0 Then
                COLUMN_NAME = "ACCT_BEG_BAL"
            Else
                COLUMN_NAME = "ACCT_ACT_P" & Format(i, "00")
            End If
            sql = sql & ", Sum (NVL(" & COLUMN_NAME & ",0)) " & COLUMN_NAME
        Next
        sql = sql & " from GLTACCT3 where ACCT_CODE = '" & HFs("ACCT_CODE") & "'"
        sql = sql & " group by ACCT_CODE" & gby & ", ACCT_YEAR"

        Fill_Records("GLTACCT3", "", , sql)






        sql = "Select GLTJRNL1.JOURNAL_TYPE, GLTDETL1.ACCT_CODE"
        gby = ""
        For i As Integer = 2 To 4
            COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
            If Absx1.chkFor(COLUMN_NAME).Checked Then
                sql = sql & ", GLTDETL1." & COLUMN_NAME & " " & COLUMN_NAME
                gby = gby & ", GLTDETL1." & COLUMN_NAME
                grdGLTACCT3.DisplayLayout.Bands(1).Columns(COLUMN_NAME).Hidden = False
                grdGLTACCT3.DisplayLayout.Bands(1).Columns(COLUMN_NAME).Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_SEG" & CStr(i) & "_DESC")
                grdGLTACCT3.DisplayLayout.Bands(1).Columns(COLUMN_NAME).Width = 60
            Else
                sql = sql & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "' " & COLUMN_NAME
                grdGLTACCT3.DisplayLayout.Bands(1).Columns(COLUMN_NAME).Hidden = True
            End If
        Next
        sql = sql & ", SUBSTR(GLTDETL1.OPS_YYYYPP,1,4) ACCT_YEAR"
        For i As Integer = 0 To 12
            If i = 0 Then
                'COLUMN_NAME = "ACCT_BEG_BAL"
                'sql = sql & ", NULL " & COLUMN_NAME
            Else
                COLUMN_NAME = "GLTDETL1.DETL_POSTING_AMT"
                sql = sql & ", Sum (DECODE(SUBSTR(GLTDETL1.OPS_YYYYPP,5,2),'" & Format(i, "00") & "',NVL(" & COLUMN_NAME & ",0),0)) P" & Format(i, "00")
            End If
        Next
        sql = sql & " from GLTDETL1,GLTJRNL1 where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO and GLTDETL1.ACCT_CODE = '" & HFs("ACCT_CODE") & "'"
        sql = sql & " group by GLTJRNL1.JOURNAL_TYPE, GLTDETL1.ACCT_CODE" & gby & ", SUBSTR(GLTDETL1.OPS_YYYYPP,1,4)"

        Fill_Records("GLTACCT3_TYPE", "", , sql)






        Dim YYYY As String = Mid(ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP"), 1, 4)

        Dim ACCT_TYPE As String = rowGLTACCT1.Item("ACCT_TYPE") & ""
        If ACCT_TYPE = "A" Or ACCT_TYPE = "L" Or ACCT_TYPE = "E" Then

            If Mid(ASCMAIN1.CYP, 1, 4) > YYYY Then
                For Y As Integer = Val(YYYY) + 1 To Val(Mid(ASCMAIN1.CYP, 1, 4))
                    For Each row As DataRow In ASCDATA1.SelectDistinct _
                        (dst.Tables("GLTACCT3").Select("ACCT_YEAR = '" & Format(Y - 1, "0000") & "'"),
                         New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"}).Rows
                        Dim rowGLTACCT3 As DataRow = dst.Tables("GLTACCT3").Rows.Find _
                            (New String() {row.Item("ACCT_CODE"),
                                           row.Item("SEG2_CODE"),
                                           row.Item("SEG3_CODE"),
                                           row.Item("SEG4_CODE"),
                                           Format(Y, "0000")})
                        If rowGLTACCT3 Is Nothing Then
                            rowGLTACCT3 = dst.Tables("GLTACCT3").NewRow
                            With rowGLTACCT3
                                .Item("ACCT_CODE") = row.Item("ACCT_CODE")
                                .Item("SEG2_CODE") = row.Item("SEG2_CODE")
                                .Item("SEG3_CODE") = row.Item("SEG3_CODE")
                                .Item("SEG4_CODE") = row.Item("SEG4_CODE")
                                .Item("ACCT_YEAR") = Format(Y, "0000")
                            End With
                            dst.Tables("GLTACCT3").Rows.Add(rowGLTACCT3)
                        End If
                    Next
                Next
            End If

            For Each rowGLTACCT3_NY As DataRow In
                dst.Tables("GLTACCT3").Select _
                ("ACCT_YEAR > '" & YYYY & "'", "ACCT_YEAR")
                Dim ACCT_END_BAL As Decimal = 0
                Dim rowGLTACCT3_PY As DataRow = dst.Tables("GLTACCT3").Rows.Find _
                    (New Object() {rowGLTACCT3_NY.Item("ACCT_CODE"),
                                   rowGLTACCT3_NY.Item("SEG2_CODE"),
                                   rowGLTACCT3_NY.Item("SEG3_CODE"),
                                   rowGLTACCT3_NY.Item("SEG4_CODE"),
                                   Format(Val(rowGLTACCT3_NY.Item("ACCT_YEAR")) - 1, "0000")})
                If rowGLTACCT3_PY IsNot Nothing Then
                    ACCT_END_BAL = Val(rowGLTACCT3_PY.Item("ACCT_BEG_BAL") & "")
                    For i As Integer = 1 To 12
                        ACCT_END_BAL += Val(rowGLTACCT3_PY.Item("ACCT_ACT_P" & Format(i, "00")) & "")
                    Next
                End If
                rowGLTACCT3_NY.Item("ACCT_BEG_BAL") = ACCT_END_BAL
            Next
        End If

        EnforceConstraints(True)

        Sort_grdColumns(grdGLTACCT3, "ACCT_YEAR,SEG2_CODE,SEG3_CODE,SEG4_CODE")

        If optShowBA.Value <> "A" Then
            Set_GLTACCT3()
        End If

    End Sub

    Sub Load_GLTDETL1()

        EnforceConstraints(False)

        dst.Tables("GLTYEAR1").Rows.Clear()
        dst.Tables("GLTPARM2").Rows.Clear()

        Dim BY(4) As Boolean
        For i As Integer = 2 To 4
            COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
            BY(i) = Absx1.chkFor(COLUMN_NAME).Checked
            With grdGLTDETL1.DisplayLayout
                If BY(i) Then
                    .Bands("GLTYEAR1").Columns(COLUMN_NAME).Hidden = False
                    .Bands("GLTYEAR1_GLTPARM2").Columns(COLUMN_NAME).Hidden = False
                    '.Bands("GLTDETL1").Columns(COLUMN_NAME).Hidden = False
                Else
                    .Bands("GLTYEAR1").Columns(COLUMN_NAME).Hidden = True
                    .Bands("GLTYEAR1_GLTPARM2").Columns(COLUMN_NAME).Hidden = True
                    '.Bands("GLTDETL1").Columns(COLUMN_NAME).Hidden = True
                End If
            End With
        Next

        Dim YYYY_min As String = ""
        Dim YYYY_max As String = ""
        For Each rowGLTACCT3 As DataRow In dst.Tables("GLTACCT3").Rows
            Dim ACCT_YEAR As String = rowGLTACCT3.Item("ACCT_YEAR")
            If ACCT_YEAR > YYYY_max Or YYYY_max = "" Then YYYY_max = ACCT_YEAR
            If ACCT_YEAR < YYYY_min Or YYYY_min = "" Then YYYY_min = ACCT_YEAR
            Dim rowGLTYEAR1 As DataRow = dst.Tables("GLTYEAR1").NewRow
            rowGLTYEAR1.Item("ACCT_YEAR") = ACCT_YEAR
            rowGLTYEAR1.Item("SEG2_CODE") = rowGLTACCT3.Item("SEG2_CODE")
            rowGLTYEAR1.Item("SEG3_CODE") = rowGLTACCT3.Item("SEG3_CODE")
            rowGLTYEAR1.Item("SEG4_CODE") = rowGLTACCT3.Item("SEG4_CODE")
            dst.Tables("GLTYEAR1").Rows.Add(rowGLTYEAR1)
            For Each row As DataRow In ASCDATA1.GetDataTable("Select * from GLTPARM2 where OPS_YYYYPP >= '" & ACCT_YEAR & "01" & "'  and OPS_YYYYPP <= '" & ACCT_YEAR & "12" & "'").Rows
                Dim rowGLTPARM2 As DataRow = dst.Tables("GLTPARM2").NewRow
                rowGLTPARM2.Item("ACCT_YEAR") = ACCT_YEAR
                rowGLTPARM2.Item("SEG2_CODE") = rowGLTACCT3.Item("SEG2_CODE")
                rowGLTPARM2.Item("SEG3_CODE") = rowGLTACCT3.Item("SEG3_CODE")
                rowGLTPARM2.Item("SEG4_CODE") = rowGLTACCT3.Item("SEG4_CODE")
                For I As Integer = 0 To row.ItemArray.Length - 1
                    rowGLTPARM2.Item(4 + I) = row.Item(I)
                Next
                dst.Tables("GLTPARM2").Rows.Add(rowGLTPARM2)
            Next
        Next

        dst.Tables("GLTPARM2").Columns("ACCT_NET_ACT").Expression = ""
        dst.Tables("GLTPARM2").Columns("ACCT_TRANSACTIONS").Expression = ""

        Fill_Records("GLTDETL1", New String() {HFs("ACCT_CODE"), YYYY_min & "01", YYYY_max & "12"})

        For Each row As DataRow In dst.Tables("GLTDETL1").Rows
            For i As Integer = 2 To 4
                COLUMN_NAME = "SEG" & CStr(i) & "_CODE"
                If BY(i) Then
                    row.Item(COLUMN_NAME & "_JOIN") = row.Item(COLUMN_NAME)
                Else
                    row.Item(COLUMN_NAME & "_JOIN") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i))
                End If
            Next
        Next


        dst.Tables("GLTPARM2").Columns("ACCT_NET_ACT").Expression = "SUM(CHILD.DETL_POSTING_AMT)"
        dst.Tables("GLTPARM2").Columns("ACCT_TRANSACTIONS").Expression = "COUNT(CHILD.OPS_YYYYPP)"

        For Each rowGLTYEAR1 As DataRow In dst.Tables("GLTYEAR1").Rows
            Dim ACCT_YEAR As String = rowGLTYEAR1.Item("ACCT_YEAR")
            Dim ACCT_BEG_BAL As Decimal = Val(dst.Tables("GLTACCT3").Compute("SUM (ACCT_BEG_BAL)",
                "ACCT_YEAR = '" & ACCT_YEAR & "'" _
                & " and SEG2_CODE = '" & rowGLTYEAR1.Item("SEG2_CODE") & "'" _
                & " and SEG3_CODE = '" & rowGLTYEAR1.Item("SEG3_CODE") & "'" _
                & " and SEG4_CODE = '" & rowGLTYEAR1.Item("SEG4_CODE") & "'") & "")
            For Each rowGLTPARM2 As DataRow In dst.Tables("GLTPARM2").Select("OPS_YYYYPP like '" & ACCT_YEAR & "*'" _
                & " and SEG2_CODE = '" & rowGLTYEAR1.Item("SEG2_CODE") & "'" _
                & " and SEG3_CODE = '" & rowGLTYEAR1.Item("SEG3_CODE") & "'" _
                & " and SEG4_CODE = '" & rowGLTYEAR1.Item("SEG4_CODE") & "'" _
                , "OPS_YYYYPP")

                Dim OPS_YYYYPP As String = rowGLTPARM2.Item("OPS_YYYYPP")
                Dim P As Integer = Val(Mid(OPS_YYYYPP, 5, 2))

                rowGLTPARM2.Item("ACCT_BEG_BAL") = ACCT_BEG_BAL
                If P = 1 Then
                    rowGLTYEAR1.Item("ACCT_BEG_BAL") = ACCT_BEG_BAL
                End If
                ACCT_BEG_BAL = Val(rowGLTPARM2.Item("ACCT_END_BAL") & "")
            Next
        Next

        EnforceConstraints(True)


        With grdGLTDETL1.DisplayLayout.Bands("GLTYEAR1")
            .SortedColumns.Clear()
            .SortedColumns.Add("ACCT_YEAR", False)
            .SortedColumns.Add("SEG2_CODE", False)
            .SortedColumns.Add("SEG3_CODE", False)
            .SortedColumns.Add("SEG4_CODE", False)
        End With

        With grdGLTDETL1.DisplayLayout.Bands("GLTYEAR1_GLTPARM2")
            .SortedColumns.Clear()
            .SortedColumns.Add("OPS_YYYYPP", False)
            .SortedColumns.Add("SEG2_CODE", False)
            .SortedColumns.Add("SEG3_CODE", False)
            .SortedColumns.Add("SEG4_CODE", False)
        End With

        With grdGLTDETL1.DisplayLayout.Bands("GLTPARM2_GLTDETL1")
            .SortedColumns.Clear()
            .SortedColumns.Add("OPS_YYYYPP", False)
            .SortedColumns.Add("SEG2_CODE", False)
            .SortedColumns.Add("SEG3_CODE", False)
            .SortedColumns.Add("SEG4_CODE", False)
            .SortedColumns.Add("JOURNAL_NO", False)
            .SortedColumns.Add("JOURNAL_LNO", False)
        End With

    End Sub
    Private Sub chkCardView_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCardView.CheckedChanged

        If chkCardView.Checked Then
            If grdGLTACCT3.ActiveRow IsNot Nothing AndAlso grdGLTACCT3.ActiveRow.IsGroupByRow Then
                grdGLTACCT3.ActiveRow.ExpandAll()
            End If
        End If

        grdGLTACCT3.DisplayLayout.Bands(0).CardView = chkCardView.Checked
        For i As Integer = 0 To 12
            Dim COLUMN_NAME As String
            If i = 0 Then
                COLUMN_NAME = "ACCT_BEG_BAL"
            Else
                COLUMN_NAME = "ACCT_ACT_P" & Format(i, "00")
            End If

            With grdGLTACCT3.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                If chkCardView.Checked Then
                    .Header.Appearance.TextHAlign = HAlign.Left
                Else
                    .Header.Appearance.TextHAlign = HAlign.Right
                End If
            End With
        Next
        '        grdGLTACCT3.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand

    End Sub

    Sub Add_GLTYEAR1(ByVal ACCT_YEAR As String)
        Dim dr As DataRow = dst.Tables("GLTYEAR1").NewRow
        dr.Item("ACCT_YEAR") = ACCT_YEAR
        dst.Tables("GLTYEAR1").Rows.Add(dr)
    End Sub

    Private Sub optRed_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRed.ValueChanged
        If Not ScreenMode Then Exit Sub
        Show_Red()
    End Sub

    Sub Show_Red()
        Dim ACCT_DR_CR_IND As String = Absx1.optFor("ACCT_DR_CR_IND").Value ' THIS DOESNOT WORK WHEN CALLED FROM LOAD_RECORD
        ACCT_DR_CR_IND = dst.Tables("GLTACCT1").Rows(0).Item("ACCT_DR_CR_IND") & ""

        Dim r As Long
        Dim COLUMN_NAME As String

        r = 0
        For Each gr As UltraWinGrid.UltraGridRow In grdGLTACCT3.DisplayLayout.Bands("GLTACCT3").GetRowEnumerator(UltraWinGrid.GridRowType.DataRow)
            'If Not gr.IsGroupByRow Then
            'If gr.IsSummaryRow Then Stop
            For I As Integer = 0 To 12
                COLUMN_NAME = IIf(I = 0, "ACCT_BEG_BAL", "ACCT_ACT_P" & Format(I, "00"))
                gr.Cells(COLUMN_NAME).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(gr.Cells(COLUMN_NAME).Value & ""))
                If r = 0 Then
                    gr.Band.Summaries(COLUMN_NAME).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(grdGLTACCT3.Rows.SummaryValues(COLUMN_NAME).Value & ""))
                End If
            Next
            r = r + 1
            'End If
        Next

        r = 0
        For Each gr As UltraWinGrid.UltraGridRow In grdGLTDETL1.Rows
            For Each COLUMN_NAME In New String() {"ACCT_BEG_BAL", "ACCT_END_BAL", "ACCT_NET_ACT"}
                If gr.Cells IsNot Nothing Then

                    gr.Cells(COLUMN_NAME).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(gr.Cells(COLUMN_NAME).Value & ""))
                    If r = 0 And COLUMN_NAME = "ACCT_NET_ACT" Then
                        gr.Band.Summaries(COLUMN_NAME).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(grdGLTDETL1.Rows.SummaryValues(COLUMN_NAME).Value & ""))
                    End If
                    For Each gr2 As UltraWinGrid.UltraGridRow In gr.ChildBands(0).Rows
                        Dim r2 As Long = 0
                        For Each COLUMN_NAME2 As String In New String() {"ACCT_BEG_BAL", "ACCT_END_BAL", "ACCT_NET_ACT"}
                            gr2.Cells(COLUMN_NAME).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(gr2.Cells(COLUMN_NAME).Value & ""))
                            If r2 = 0 And COLUMN_NAME2 = "ACCT_NET_ACT" Then
                                gr2.Band.Summaries(COLUMN_NAME2).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(gr.ChildBands(0).Rows.SummaryValues(COLUMN_NAME2).Value & ""))
                            End If

                            For Each gr3 As UltraWinGrid.UltraGridRow In gr2.ChildBands(0).Rows
                                Dim r3 As Long = 0
                                For Each COLUMN_NAME3 As String In New String() {"DETL_POSTING_AMT"}
                                    gr3.Cells(COLUMN_NAME3).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(gr3.Cells(COLUMN_NAME3).Value & ""))
                                    If r3 = 0 And COLUMN_NAME3 = "DETL_POSTING_AMT" Then
                                        Try
                                            gr3.Band.Summaries(COLUMN_NAME3).Appearance.ForeColor = Get_CellColor(ACCT_DR_CR_IND, Val(gr2.ChildBands(0).Rows.SummaryValues(COLUMN_NAME3).Value & ""))
                                        Catch ex As Exception

                                        End Try
                                    End If
                                Next
                                r3 = r3 + 1
                            Next
                        Next
                        r2 = r2 + 1
                    Next
                End If
            Next
            r = r + 1
        Next
    End Sub

    Function Get_CellColor(ByVal ACCT_DR_CR_IND As String, ByVal AMT As Decimal) As System.Drawing.Color
        If (optRed.Value = "C" And AMT < 0) _
        Or (optRed.Value = "X" And
           ((ACCT_DR_CR_IND = "D" And AMT < 0) _
          Or ACCT_DR_CR_IND = "C" And AMT > 0)) Then
            Return Drawing.Color.Red
        Else
            Return Drawing.Color.Black
        End If
    End Function

    Private Sub optWith_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optWith.ValueChanged
        If Not ScreenMode Then Exit Sub
        Call Show_With()
    End Sub

    Sub Show_With()
        Dim MASK As String = ""
        MASK = "###,##0.00;(###,##0.00);###,##0.00"
        Select Case optWith.Value
            Case "B"
                MASK = "###,##0.00;(###,##0.00);###,##0.00"
            Case "M"
                MASK = "###,##0.00"
            Case "C"
                MASK = "###,##0.00DR;###,##0.00CR"
            Case "X"
                Dim ACCT_DR_CR_IND As String = rowGLTACCT1.Item("ACCT_DR_CR_IND")
                If ACCT_DR_CR_IND = "C" Then
                    MASK = "(###,##0.00);###,##0.00;###,##0.00"
                Else
                    MASK = "###,##0.00;(###,##0.00);###,##0.00"
                End If
        End Select

        'MASK = Replace(Replace(MASK, "#", "n"), "0", "n")

        With grdGLTACCT3
            For I As Integer = 0 To 12
                Dim COLUMN_NAME As String = ""
                If I = 0 Then
                    COLUMN_NAME = "ACCT_BEG_BAL"
                Else
                    COLUMN_NAME = "ACCT_ACT_P" & Format(I, "00")
                End If

                .DisplayLayout.Bands(0).Columns(COLUMN_NAME).Format = MASK
                If .DisplayLayout.Bands(0).Summaries.Count <> 0 Then
                    .DisplayLayout.Bands(0).Summaries(COLUMN_NAME).DisplayFormat = "{0:" & MASK & "}"
                End If
            Next
            .Refresh()
        End With

        With grdGLTDETL1
            For B As Integer = 0 To 1
                For Each COLUMN_NAME In New String() {"ACCT_BEG_BAL", "ACCT_END_BAL", "ACCT_NET_ACT"}
                    .DisplayLayout.Bands(B).Columns(COLUMN_NAME).Format = MASK
                    If .DisplayLayout.Bands(B).Summaries.Count <> 0 And COLUMN_NAME = "ACCT_NET_ACT" Then
                        .DisplayLayout.Bands(B).Summaries(COLUMN_NAME).DisplayFormat = "{0:" & MASK & "}"
                    End If
                Next
            Next
            COLUMN_NAME = "DETL_POSTING_AMT"
            .DisplayLayout.Bands("GLTPARM2_GLTDETL1").Columns(COLUMN_NAME).Format = MASK
            If .DisplayLayout.Bands("GLTPARM2_GLTDETL1").Summaries.Count <> 0 And COLUMN_NAME = "DETL_POSTING_AMT" Then
                .DisplayLayout.Bands("GLTPARM2_GLTDETL1").Summaries(COLUMN_NAME).DisplayFormat = "{0:" & MASK & "}"
            End If
            .Refresh()
        End With
    End Sub

    Private Sub grdGLTDETL1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdGLTDETL1.InitializeRow
        If e.Row.Band.Key = "GLTDETL1" Then
            If e.Row.Cells("DETL_CVX_NO").Text <> "" Then
                Select Case e.Row.Cells("DETL_CVX_TYPE").Text
                    Case "V"
                        e.Row.Cells("DETL_CVX_NAME").Value = LookUp("APTVEND1", e.Row.Cells("DETL_CVX_NO").Text, True).Item("VEND_NAME")
                    Case "C"
                        e.Row.Cells("DETL_CVX_NAME").Value = LookUp("ARTCUST1", e.Row.Cells("DETL_CVX_NO").Text, True).Item("CUST_NAME")
                End Select
            End If
        End If
    End Sub

    Private Sub grdGLTACCT3_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTACCT3.InitializeLayout
        With grdGLTACCT3.DisplayLayout.Bands(0)
            .Columns("SEG2_CODE").Header.Fixed = True
            .Columns("SEG3_CODE").Header.Fixed = True
            .Columns("SEG4_CODE").Header.Fixed = True
            .Columns("ACCT_YEAR").Header.Fixed = True
        End With
        With grdGLTACCT3.DisplayLayout.Bands(1)
            .Columns("SEG2_CODE").Header.Fixed = True
            .Columns("SEG3_CODE").Header.Fixed = True
            .Columns("SEG4_CODE").Header.Fixed = True
            .Columns("JOURNAL_TYPE").Header.Fixed = True
            .ColHeadersVisible = False
        End With

    End Sub

    Private Sub chkSEG2_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSEG2_CODE.CheckedChanged
        If ScreenMode Then Load_Data()
    End Sub

    Private Sub chkSEG3_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSEG3_CODE.CheckedChanged
        If ScreenMode Then Load_Data()
    End Sub

    Private Sub chkSEG4_CODE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSEG4_CODE.CheckedChanged
        If ScreenMode Then Load_Data()
    End Sub

    Private Sub optShowBA_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optShowBA.ValueChanged
        If Not ScreenMode Then Exit Sub
        Set_GLTACCT3()
    End Sub

    Sub Set_GLTACCT3()
        'Stop
        Dim P As Integer = 5 ' 0 based position of ACCT_BEG_BAL, which is just prior to P01
        For Each rowGLTACCT3 As DataRow In dst.Tables("GLTACCT3").Rows
            If optShowBA.Value = "B" Then
                For i As Integer = 1 To 13
                    rowGLTACCT3(P + i) = Val(rowGLTACCT3(P + i) & "") + Val(rowGLTACCT3(P + i - 1) & "")
                Next
            Else
                For i As Integer = 13 To 1 Step -1
                    rowGLTACCT3(P + i) = Val(rowGLTACCT3(P + i) & "") - Val(rowGLTACCT3(P + i - 1) & "")
                Next
            End If
        Next

        Show_Red()
    End Sub

    Sub Set_Year_Filters()

        Dim sql As String = ""

        If Absx1.chkFor("NEXT_YEAR").Checked _
        And Absx1.chkFor("CURR_YEAR").Checked _
        And Absx1.chkFor("LAST_YEAR").Checked _
        And Absx1.chkFor("PRIOR_YEARS").Checked Then
        Else
            If Absx1.chkFor("NEXT_YEAR").Checked Then
                sql &= " OR ACCT_YEAR = '" & Absx1.chkFor("NEXT_YEAR").Text & "'"
            End If
            If Absx1.chkFor("CURR_YEAR").Checked Then
                sql &= " OR ACCT_YEAR = '" & Absx1.chkFor("CURR_YEAR").Text & "'"
            End If
            If Absx1.chkFor("LAST_YEAR").Checked Then
                sql &= " OR ACCT_YEAR = '" & Absx1.chkFor("LAST_YEAR").Text & "'"
            End If
            If Absx1.chkFor("PRIOR_YEARS").Checked Then
                sql &= " OR ACCT_YEAR < '" & Absx1.chkFor("LAST_YEAR").Text & "'"
            End If

            If sql = "" Then
                sql = "ACCT_YEAR = '0000'"
            Else
                sql = Mid(sql, 5)
            End If
        End If

        Dim dvw As DataView

        dvw = DirectCast(grdGLTACCT3.DataSource, DataTable).DefaultView
        dvw.RowFilter = sql

        dvw = dst.Tables("GLTYEAR1").DefaultView
        dvw.RowFilter = sql

    End Sub

    Private Sub grdGLTACCTX_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdGLTACCTX.DoubleClickRow

        If e.Row.IsDataRow Then
            Absx1.txtFor("ACCT_CODE").Text = e.Row.Cells("ACCT_CODE").Value
            Click_Command("Load")
        End If

    End Sub

    Sub Create_GLTACCTU(initialize As Boolean)
        ASCMAIN1.sql = "Select X.*, K.COLUMN_NAME_KEY" & vbCrLf _
            & " from (Select TABLE_NAME, COLUMN_NAME" & vbCrLf _
            & " from USER_TAB_COLUMNS WHERE COLUMN_NAME LIKE '%ACCT%' AND DATA_LENGTH = 6" & vbCrLf _
            & " and LENGTH(TABLE_NAME) = 8" & vbCrLf _
            & " and (TABLE_NAME NOT LIKE 'GLT%' OR TABLE_NAME = 'GLTPARM1')" & vbCrLf _
            & " and TABLE_NAME NOT LIKE 'MET%'" & vbCrLf _
            & " and TABLE_NAME NOT LIKE 'ASW%'" & vbCrLf _
            & " and TABLE_NAME NOT IN (" & vbCrLf _
            & "'APTINVH2','ARTPYMT4','ARTPYMT5','ICTIADJ3','ICTIREC3','SOTCARR3','POTNINV1'," & vbCrLf _
            & "'ICTIXFR3','ICTPINV3','ICTTRAN5','SOTRTRN3','SOTINVHG')" & vbCrLf _
            & ") X, " & vbCrLf _
            & "(Select TABLE_NAME, COLUMN_NAME COLUMN_NAME_KEY" & vbCrLf _
            & " from USER_TAB_COLUMNS WHERE COLUMN_ID = 1) K" & vbCrLf _
            & " where K.TABLE_NAME = X.TABLE_NAME"

        Dim sqlT As String = "Select ACCT_CODE, 'TABLE_NAME' TBL, 'COLUMN_NAME' COL, 'TABLE_KEY_COLUMN' TABLE_KEY" & vbCrLf _
            & ", ltrim(sys_connect_by_path(COLUMN_NAME_KEY,','),',') KEYS" & vbCrLf _
            & " from (select ACCT_CODE, COLUMN_NAME_KEY, row_number() over(partition by ACCT_CODE" & vbCrLf _
            & " order by COLUMN_NAME_KEY) rn, row_number() over(partition by ACCT_CODE" & vbCrLf _
            & " order by COLUMN_NAME_KEY desc) rn_desc from " & vbCrLf _
            & "(Select Distinct COLUMN_NAME ACCT_CODE, COLUMN_NAME_KEY from TABLE_NAME where " & IIf(initialize, "ROWNUM <1 and ", "") & "COLUMN_NAME is Not Null))" & vbCrLf _
            & " where rn_desc = 1 start with rn = 1 connect by prior ACCT_CODE = ACCT_CODE and prior rn = rn-1"

        Dim sqlTX As String = ""

        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim COLUMN_NAME As String = row.Item("COLUMN_NAME")
            Dim COLUMN_NAME_KEY As String = row.Item("COLUMN_NAME_KEY")
            Dim TABLE_NAME As String = row.Item("TABLE_NAME")
            sqlTX &= " union " & vbCrLf _
                & Replace(Replace(Replace(Replace(sqlT,
                                             "COLUMN_NAME_KEY", COLUMN_NAME_KEY),
                                             "COLUMN_NAME", COLUMN_NAME),
                                             "TABLE_NAME", TABLE_NAME),
                                         "TABLE_KEY_COLUMN", COLUMN_NAME_KEY)
            If initialize Then
                Exit For
            End If
        Next

        sqlTX = Mid(sqlTX, 10)

        If GLTACCTU = "" Then
            GLTACCTU = ASCMAIN1.Temp_Table(sqlTX)
            'ASCDATA1.ExecuteSQL("Alter Table " & GLTACCTU & " Add Primary Key ()")
            ASCDATA1.ExecuteSQL("Alter Table " & GLTACCTU & " MODIFY COL VARCHAR2(100)")
            ASCDATA1.ExecuteSQL("Alter Table " & GLTACCTU & " MODIFY TABLE_KEY VARCHAR2(100)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & GLTACCTU)
            ASCDATA1.ExecuteSQL("Insert into " & GLTACCTU & " " & sqlTX)
        End If

    End Sub

    Private Sub tabGLTACCT1_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabGLTACCT1.SelectedTabChanged
        Setup_tabGLTACCT1()
    End Sub

    Sub Setup_tabGLTACCT1()
        grpSummary.Visible = ScreenMode AndAlso (tabGLTACCT1.SelectedTab.Key = "Summary")
    End Sub

    Private Sub chkDetails_CheckedChanged(sender As Object, e As EventArgs) Handles chkDetails.CheckedChanged
        Setup_Details()
    End Sub

    Sub Setup_Details()
        splGLTACCT1.Panel2Collapsed = Not (chkDetails.Checked)
        grdGLTDETLX.Visible = False
        If ScreenMode Then
            Load_Details()
        End If
    End Sub

    Private Sub grdGLTACCT3_AfterRowActivate(sender As Object, e As EventArgs) Handles grdGLTACCT3.AfterRowActivate
        If ScreenMode Then
            If grdGLTACCT3.ActiveRow.Band.Key = "GLTACCT3_GLTACCT3_TYPE" Then
                Load_Details()
            End If
        End If
    End Sub

    Sub Load_Details()
        If grdGLTACCT3.ActiveRow Is Nothing OrElse grdGLTACCT3.ActiveRow.Band.Key <> "GLTACCT3_GLTACCT3_TYPE" Or splGLTACCT1.Panel2Collapsed Then Exit Sub
        If grdGLTACCT3.ActiveRow.IsFilterRow Then Exit Sub
        Dim ACCT_YEAR As String = grdGLTACCT3.ActiveRow.Cells("ACCT_YEAR").Value
        Dim JOURNAL_TYPE As String = grdGLTACCT3.ActiveRow.Cells("JOURNAL_TYPE").Value

        Dim SEG2_CODE As String = grdGLTACCT3.ActiveRow.Cells("SEG2_CODE").Value
        Dim SEG3_CODE As String = grdGLTACCT3.ActiveRow.Cells("SEG3_CODE").Value
        Dim SEG4_CODE As String = grdGLTACCT3.ActiveRow.Cells("SEG4_CODE").Value

        Dim sql_SEG As String = ""
        If chkSEG2_CODE.Checked Then sql_SEG &= " and NVL(X.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') = '" & SEG2_CODE & "'"
        If chkSEG3_CODE.Checked Then sql_SEG &= " and NVL(X.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') = '" & SEG3_CODE & "'"
        If chkSEG4_CODE.Checked Then sql_SEG &= " and NVL(X.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') = '" & SEG4_CODE & "'"

        Dim C As String = ""
        If grdGLTACCT3.ActiveCell IsNot Nothing Then
            C = grdGLTACCT3.ActiveCell.Column.Key
        End If

        Dim OPS_YYYYPP As String = ""
        If C.StartsWith("ACCT_ACT_P") Then
            Dim P As Integer = Mid(C, 10, 2)
            If P >= 1 And P <= 12 Then
                OPS_YYYYPP = ACCT_YEAR & Format(P, "00")
            End If
        End If

        If grdGLTDETLX.Tag = ACCT_YEAR & ":" & JOURNAL_TYPE & ":" & SEG2_CODE & ":" & SEG3_CODE & ":" & SEG4_CODE Then Exit Sub

        Dim rowGLTTYPE1 As DataRow = dst.Tables("GLTTYPE1").Rows.Find(JOURNAL_TYPE)
        Dim JOURNAL_TYPE_DESC As String = ""
        If rowGLTTYPE1 IsNot Nothing Then JOURNAL_TYPE_DESC = rowGLTTYPE1.Item("JOURNAL_TYPE_DESC") & ""

        grdGLTDETLX.Text = "Transaction Details for " & Absx1.txtFor("ACCT_CODE").Text & " in " & ACCT_YEAR & " for Journal Type " & JOURNAL_TYPE & " " & JOURNAL_TYPE_DESC

        grdGLTDETLX.DataSource = Nothing

        ASCMAIN1.sql = ""

        Dim ACCT_CODE As String = Absx1.txtFor("ACCT_CODE").Text

        sqlX = GLCMAIN1.Load_Details(ACCT_YEAR, JOURNAL_TYPE, ACCT_CODE, sql_SEG, sqlWhereClause)

        If ASCMAIN1.sql = "" Then
            grdGLTDETLX.Visible = False
        Else
            grdGLTDETLX.DisplayLayout.Bands(0).Summaries.Clear()
            'grdGLTDETLX.DisplayLayout.CaptionVisible = DefaultableBoolean.Default
            'grdGLTDETLX.DisplayLayout.CaptionAppearance .TextHAlign =HAlign.Left 
            grdGLTDETLX.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
            grdGLTDETLX.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show

            Dim TBL As DataTable = ASCDATA1.GetDataTable
            TBL.TableName = "GLTDETLX"
            grdGLTDETLX.DataSource = TBL
            ASCMAIN1.grdInitializeLayout(grdGLTDETLX)
            grdGLTDETLX.Visible = True
        End If

        With grdGLTDETLX.DisplayLayout.Bands(0)
            .Tag = ""

            If JOURNAL_TYPE = "APCD" Then Headings("Vendor", 100, "Bank", 60, "", 0, "Bank Desc", 100)
            If JOURNAL_TYPE = "APIN" Then Headings("Vendor", 100, "Post", 60, "", 0, "Vendor Name", 100)
            If JOURNAL_TYPE = "ARCR" Then Headings("Customer", 100, "RC", 60, "", 0, "Description", 100)
            If JOURNAL_TYPE = "GLJE" Then Headings("Journal", 100, "Comment", 60, "Detail", 60, "Description", 100)
            If JOURNAL_TYPE = "OPCJ" Then Headings("Customer", 100, "Prod", 60, "", 0, "Description", 100)
            If JOURNAL_TYPE = "OPSJ" Then Headings("Customer", 100, "Prod", 60, "", 0, "Description", 100)
            If JOURNAL_TYPE = "SPAC" Then Headings("Customer", 100, "ASP Code", 60, "", 0, "Description", 100)
            If JOURNAL_TYPE = "SPCA" Then Headings("Customer", 100, "Type Code", 60, "", 0, "Description", 100)
            If JOURNAL_TYPE = "SPCP" Then Headings("Customer", 100, "Type Code", 60, "", 0, "Description", 100)
            If JOURNAL_TYPE = "OPRA" Then Headings("Customer", 100, "Type Code", 60, "", 0, "Description", 100)
            If .Tag = "" Then
            Else

                For I As Integer = 1 To 12
                    .Columns("AMT_" & Format(I, "00")).Header.Caption = grdGLTACCT3.DisplayLayout.Bands(0).Columns("ACCT_ACT_P" & Format(I, "00")).Header.Caption
                    .Columns("AMT_" & Format(I, "00")).Width = grdGLTACCT3.DisplayLayout.Bands(0).Columns("ACCT_ACT_P" & Format(I, "00")).Width
                    Create_Summary(grdGLTDETLX, "AMT_" & Format(I, "00"))
                Next

                .Columns("AMT").Header.Caption = "Total"
                .Columns("AMT").Width = 100
                Create_Summary(grdGLTDETLX, "AMT")
            End If
        End With

        grdGLTDETLX.Tag = ACCT_YEAR & ":" & JOURNAL_TYPE & ":" & SEG2_CODE & ":" & SEG3_CODE & ":" & SEG4_CODE
    End Sub

    Sub Headings(CODE1_hdg As String, CODE1_width As Integer,
                 CODE2_hdg As String, CODE2_width As Integer,
                 CODE3_hdg As String, CODE3_width As Integer,
                 DESC_hdg As String, DESC_width As Integer)

        With grdGLTDETLX.DisplayLayout.Bands(0)
            .Tag = "X"

            .Columns("CODE1_VALUE").Header.Caption = CODE1_hdg
            .Columns("CODE1_VALUE").Width = CODE1_width
            .Columns("CODE1_VALUE").Hidden = (CODE1_hdg = "")

            .Columns("CODE2_VALUE").Header.Caption = CODE2_hdg
            .Columns("CODE2_VALUE").Width = CODE2_width
            .Columns("CODE2_VALUE").Hidden = (CODE2_hdg = "")

            .Columns("CODE3_VALUE").Header.Caption = CODE3_hdg
            .Columns("CODE3_VALUE").Width = CODE3_width
            .Columns("CODE3_VALUE").Hidden = (CODE3_hdg = "")

            .Columns("DESC_VALUE").Header.Caption = DESC_hdg
            .Columns("DESC_VALUE").Width = DESC_width
        End With

    End Sub

    Private Sub grdGLTDETLX_DoubleClickCell(sender As Object, e As DoubleClickCellEventArgs) Handles grdGLTDETLX.DoubleClickCell

        If grdGLTACCT3.ActiveRow Is Nothing OrElse grdGLTACCT3.ActiveRow.Band.Key <> "GLTACCT3_GLTACCT3_TYPE" Then
            Exit Sub
        End If

        Dim JOURNAL_TYPE As String = grdGLTACCT3.ActiveRow.Cells("JOURNAL_TYPE").Value

        Dim C As String = e.Cell.Column.Key
        If C.Length <> 6 Or Not C.StartsWith("AMT_") Then
            Exit Sub
        End If

        Dim ACCT_YEAR As String = grdGLTACCT3.ActiveRow.Cells("ACCT_YEAR").Value

        Dim OPS_YYYYPP = ACCT_YEAR & Mid(C, 5, 2)
        Dim CODE1_VALUE As String = ""
        Dim CODE2_VALUE As String = ""
        Dim CODE3_VALUE As String = ""
        CODE1_VALUE = e.Cell.Row.Cells("CODE1_VALUE").Value & ""
        CODE2_VALUE = e.Cell.Row.Cells("CODE2_VALUE").Value & ""
        CODE3_VALUE = e.Cell.Row.Cells("CODE3_VALUE").Value & ""

        GLCMAIN1.Load_Sub_Details(JOURNAL_TYPE, ACCT_YEAR, OPS_YYYYPP, sqlX, sqlWhereClause, CODE1_VALUE, CODE2_VALUE, CODE3_VALUE)

        If ASCMAIN1.sql <> "" Then
            Using frm As New ASFMSGBF
                Dim tbl As DataTable = ASCDATA1.GetDataTable()
                tbl.TableName = "GLTDETLY"
                frm.Width = Me.Width * 0.9
                frm.Show_grd(tbl, Me, grdGLTDETLX.Text & " - " & CODE1_VALUE & ":" & CODE2_VALUE & ":" & OPS_YYYYPP, "GLTDETLY")
            End Using
        End If

    End Sub

    Overrides Sub Format_grd_ASFMSGBF(ByVal grd As UltraWinGrid.UltraGrid, ByVal grdCode As String)

        Dim JOURNAL_TYPE As String = grdGLTACCT3.ActiveRow.Cells("JOURNAL_TYPE").Value

        Select Case grdCode
            Case "GLTDETLY"
                With grd.DisplayLayout.Override
                    .CellClickAction = UltraWinGrid.CellClickAction.CellSelect
                    .ActiveRowAppearance.BackColor = Drawing.Color.Empty
                    .ActiveRowAppearance.ForeColor = Drawing.Color.Empty
                End With
                With grd.DisplayLayout.Bands(0)
                    .Override.CellClickAction = UltraWinGrid.CellClickAction.CellSelect
                    ' .Columns("COLUMN_NAME").Hidden = True
                    With .Columns("OPS_YYYYPP")
                        .Header.Caption = "YP"
                        .Width = 70
                    End With

                    Select Case JOURNAL_TYPE


                        Case "APCD"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Vendor"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "Bank Code"
                                .Width = 100
                            End With
                            With .Columns("CODE3_VALUE")
                                .Hidden = True
                            End With

                            With .Columns("PYMT_METHOD")
                                .Header.Caption = "Method"
                                .Width = 70
                            End With
                            With .Columns("CHECK_NUM")
                                .Header.Caption = "Check No"
                                .Width = 100
                            End With
                            With .Columns("CHECK_DATE")
                                .Header.Caption = "Chk Date"
                                .Width = 100
                            End With


                        Case "APIN"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Vendor"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "Post Code"
                                .Width = 100
                            End With
                            With .Columns("CODE3_VALUE")
                                .Hidden = True
                            End With

                            With .Columns("INV_TYPE")
                                .Header.Caption = "T"
                                .Width = 30
                            End With
                            With .Columns("INV_NUM")
                                .Header.Caption = "Invoice No"
                                .Width = 100
                            End With

                            With .Columns("INV_DATE")
                                .Header.Caption = "Inv Date"
                                .Width = 100
                            End With
                            With .Columns("INV_REF")
                                .Header.Caption = "Ref"
                                .Width = 100
                            End With
                            With .Columns("CHECK_NUM")
                                .Header.Caption = "Check No"
                                .Width = 100
                            End With
                            With .Columns("CHECK_DATE")
                                .Header.Caption = "Chk Date"
                                .Width = 100
                            End With


                        Case "ARCR"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Customer"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "RC"
                                .Width = 50
                                '.Style = UltraWinGrid.ColumnStyle.CheckBox
                            End With
                            With .Columns("CODE3_VALUE")
                                .Hidden = True
                            End With
                            With .Columns("BANK_CODE")
                                .Header.Caption = "Bank"
                                .Width = 70
                            End With
                            With .Columns("PYMT_SOURCE")
                                .Header.Caption = "Source"
                                .Width = 60
                            End With
                            With .Columns("CUST_PYMT_REF_NO")
                                .Header.Caption = "Cust Ref"
                                .Width = 90
                            End With
                            With .Columns("CUST_PYMT_REF_DATE")
                                .Header.Caption = "Ref Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("PYMT_BATCH_NO")
                                .Header.Caption = "Batch No"
                                .Width = 100
                            End With
                            With .Columns("PYMT_BATCH_LNO")
                                .Header.Caption = "Ln"
                                .Width = 40
                                .CellAppearance.TextHAlign = HAlign.Right
                                .Format = "#,##0"
                            End With
                            With .Columns("REF1")
                                .Header.Caption = "Ref 1"
                                .Width = 100
                            End With
                            With .Columns("REF2")
                                .Header.Caption = "Ref 2"
                                .Width = 60
                            End With

                        Case "GLJE"
                            With .Columns("CODE1_VALUE")
                                .Hidden = True
                            End With
                            With .Columns("CODE2_VALUE")
                                .Hidden = True
                            End With
                            With .Columns("CODE3_VALUE")
                                .Hidden = True
                            End With

                            With .Columns("DESC_VALUE")
                                .Header.Caption = "Description"
                                .Width = 200
                            End With

                            '"GLTJRNL1.JOURNAL_DESC, GLTDETL1.DETL_DESC, GLTDETL1.DETL_CTL_NO,
                            'With .Columns("JOURNAL_DESC")
                            '    .Header.Caption = "Description"
                            '    .Width = 200
                            'End With
                            'With .Columns("DETL_DESC")
                            '    .Header.Caption = "Detail Description"
                            '    .Width = 80
                            'End With
                            'With .Columns("DETL_CTL_NO")
                            '    .Header.Caption = "Detail Ctl No"
                            '    .Width = 80
                            'End With
                            With .Columns("DETL_CVX_NO")
                                .Header.Caption = "Entity"
                                .Width = 80
                            End With

                        Case "OPSJ", "OPCJ"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Customer"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "Prod"
                                .Width = 50
                            End With
                            With .Columns("CODE3_VALUE")
                                .Header.Caption = "Misc"
                                .Width = 50
                            End With
                            With .Columns("INV_TYPE")
                                .Header.Caption = "T"
                                .Width = 40
                            End With
                            With .Columns("INV_NO")
                                .Header.Caption = "Invoice No"
                                .Width = 110
                            End With
                            With .Columns("ORDR_NO")
                                .Header.Caption = "Order No"
                                .Width = 110
                            End With
                            With .Columns("ORDR_CUST_PO")
                                .Header.Caption = "Customer PO"
                                .Width = 130
                            End With
                            With .Columns("ORDR_TYPE_CODE")
                                .Header.Caption = "Type"
                                .Width = 60
                            End With
                            With .Columns("CUST_STORE_NO")
                                .Header.Caption = "Store"
                                .Width = 80
                            End With
                            With .Columns("EVENT_CODE")
                                .Header.Caption = "Event"
                                .Width = 80
                            End With
                            With .Columns("ROW_NO")
                                .Hidden = True
                            End With


                            'With .Columns("DESC_VALUE")
                            '    .Header.Caption = "Description"
                            '    .Width = 200
                            'End With



                        Case "OPRA"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Customer"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "Type"
                                .Width = 50
                            End With
                            'With .Columns("CODE3_VALUE")
                            '    .Header.Caption = "Misc"
                            '    .Width = 50
                            'End With
                            With .Columns("CUST_CLAIM_NO")
                                .Header.Caption = "Cust Claim No"
                                .Width = 100
                            End With
                            With .Columns("RA_NO")
                                .Header.Caption = "RA No"
                                .Width = 70
                            End With
                            'With .Columns("ITEM_CODE")
                            '    .Header.Caption = "Item Code"
                            '    .Width = 100
                            'End With
                            With .Columns("RA_DATE")
                                .Header.Caption = "RA Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("CUST_STORE_NO")
                                .Header.Caption = "Store"
                                .Width = 70
                            End With
  


                            With .Columns("DESC_VALUE")
                                .Header.Caption = "RA Notes"
                                .Width = 200
                            End With


                        Case "SPAC"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Customer"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "ASP Code"
                                .Width = 100
                            End With
                            With .Columns("CODE3_VALUE")
                                .Hidden = True
                            End With

                            With .Columns("HC_CODE")
                                .Header.Caption = "HC"
                                .Width = 100
                            End With
                            With .Columns("BRAND_CODE")
                                .Header.Caption = "Brand"
                                .Width = 100
                            End With

                            With .Columns("OPS_YYYYWW_MIN")
                                .Header.Caption = "YP1"
                                .Width = 100
                            End With
                            With .Columns("OPS_YYYYWW_MAX")
                                .Header.Caption = "YP2"
                                .Width = 100
                            End With


                            With .Columns("DESC_VALUE")
                                .Header.Caption = "Agreement"
                                .Width = 200
                            End With


                        Case "SPCA"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Customer"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "Exp Type"
                                .Width = 50
                            End With
                            With .Columns("CODE3_VALUE")
                                .Header.Caption = "Collctn"
                                .Width = 50
                            End With
                            With .Columns("SEASON_CODE")
                                .Header.Caption = "Season"
                                .Width = 70
                            End With
                            With .Columns("AUTH_NO")
                                .Header.Caption = "Auth No"
                                .Width = 70
                            End With
                            'With .Columns("PYMT_NO")
                            '    .Header.Caption = "Pymt No"
                            '    .Width = 60
                            'End With
                            With .Columns("VEHICLE_CODE")
                                .Header.Caption = "Vehicle"
                                .Width = 90
                            End With
                            With .Columns("DATE_START")
                                .Header.Caption = "Start Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("DATE_END")
                                .Header.Caption = "End Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("DATE_ACCRUE")
                                .Header.Caption = "Accrue Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("OPS_YYYYPP_ACCRUE")
                                .Header.Caption = "YP Accrue"
                                .Width = 80
                            End With
                            'With .Columns("PYMT_CTL_NO")
                            '    .Header.Caption = "On/Acct"
                            '    .Width = 100
                            'End With
                            'With .Columns("PYMT_REF_NO")
                            '    .Header.Caption = "Customer Ref"
                            '    .Width = 80
                            'End With
                            'With .Columns("PYMT_REF_DATE")
                            '    .Header.Caption = "Pymt Date"
                            '    .Width = 90
                            '    .Format = "MM/dd/yy"
                            'End With
                            'With .Columns("PYMT_TYPE")
                            '    .Header.Caption = "Pmt Typ"
                            '    .Width = 60
                            'End With
                            'With .Columns("VEND_CODE")
                            '    .Header.Caption = "Vendor"
                            '    .Width = 60
                            'End With
                            'With .Columns("DIST_AMT")
                            '    .Header.Caption = "Dist Amt"
                            '    .Width = 70
                            '    .CellAppearance.TextHAlign = HAlign.Right
                            '    .Format = "#,##0.00"
                            'End With
                            'With .Columns("DIST_AMT_PYMT")
                            '    .Header.Caption = "Dist Amt Pymt"
                            '    .Width = 70
                            '    .CellAppearance.TextHAlign = HAlign.Right
                            '    .Format = "#,##0.00"
                            'End With
                            'With .Columns("BRAND_CODE")
                            '    .Header.Caption = "Brand"
                            '    .Width = 60
                            'End With
                            'With .Columns("TRADE_CLASS_CODE")
                            '    .Header.Caption = "TrCls"
                            '    .Width = 60
                            'End With


                            With .Columns("DESC_VALUE")
                                .Header.Caption = "Booking Name"
                                .Width = 200
                            End With

                        Case "SPCP"
                            With .Columns("CODE1_VALUE")
                                .Header.Caption = "Customer"
                                .Width = 100
                            End With
                            With .Columns("CODE2_VALUE")
                                .Header.Caption = "Exp Type"
                                .Width = 50
                            End With
                            With .Columns("CODE3_VALUE")
                                .Header.Caption = "Collctn"
                                .Width = 50
                            End With
                            With .Columns("SEASON_CODE")
                                .Header.Caption = "Season"
                                .Width = 70
                            End With
                            With .Columns("AUTH_NO")
                                .Header.Caption = "Auth No"
                                .Width = 70
                            End With
                            With .Columns("PYMT_NO")
                                .Header.Caption = "Pymt No"
                                .Width = 60
                            End With
                            With .Columns("VEHICLE_CODE")
                                .Header.Caption = "Vehicle"
                                .Width = 90
                            End With
                            With .Columns("DATE_START")
                                .Header.Caption = "Start Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("DATE_END")
                                .Header.Caption = "End Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("DATE_ACCRUE")
                                .Header.Caption = "Accrue Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("OPS_YYYYPP_ACCRUE")
                                .Header.Caption = "YP Accrue"
                                .Width = 80
                            End With
                            With .Columns("PYMT_CTL_NO")
                                .Header.Caption = "On/Acct"
                                .Width = 100
                            End With
                            With .Columns("PYMT_REF_NO")
                                .Header.Caption = "Customer Ref"
                                .Width = 80
                            End With
                            With .Columns("PYMT_REF_DATE")
                                .Header.Caption = "Pymt Date"
                                .Width = 90
                                .Format = "MM/dd/yy"
                            End With
                            With .Columns("PYMT_TYPE")
                                .Header.Caption = "Pmt Typ"
                                .Width = 60
                            End With
                            With .Columns("VEND_CODE")
                                .Header.Caption = "Vendor"
                                .Width = 60
                            End With
                            With .Columns("DIST_AMT")
                                .Header.Caption = "Dist Amt"
                                .Width = 70
                                .CellAppearance.TextHAlign = HAlign.Right
                                .Format = "#,##0.00"
                            End With
                            With .Columns("DIST_AMT_PYMT")
                                .Header.Caption = "Dist Amt Pymt"
                                .Width = 70
                                .CellAppearance.TextHAlign = HAlign.Right
                                .Format = "#,##0.00"
                            End With
                            With .Columns("BRAND_CODE")
                                .Header.Caption = "Brand"
                                .Width = 60
                            End With
                            With .Columns("TRADE_CLASS_CODE")
                                .Header.Caption = "TrCls"
                                .Width = 60
                            End With


                            With .Columns("DESC_VALUE")
                                .Header.Caption = "Booking Name"
                                .Width = 200
                            End With
                    End Select

                    If .Columns.Exists("INIT_DATE") Then
                        With .Columns("INIT_DATE")
                            .Format = "MM/dd/yy HH:mm"
                            .Width = 120
                            .Header.Caption = "Created"
                        End With
                        With .Columns("INIT_OPER")
                            .Header.Caption = "By"
                            .Width = 80
                        End With
                    End If

                    If .Columns.Exists("LAST_DATE") Then
                        With .Columns("LAST_DATE")
                            .Format = "MM/dd/yy HH:mm"
                            .Width = 120
                            .Header.Caption = "Last Chg"
                        End With
                        With .Columns("LAST_OPER")
                            .Header.Caption = "By"
                            .Width = 80
                        End With
                    End If


                    With .Columns("AMT")
                        .Header.Caption = "Amount"
                        .Width = 100
                        .CellAppearance.TextHAlign = HAlign.Right
                        .Header.Appearance.TextHAlign = HAlign.Right
                        .Format = "#,##0.00"
                        Create_Summary(grd, "AMT")
                    End With


                End With

        End Select
    End Sub

End Class