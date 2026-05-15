Imports Infragistics.Win.UltraWinGrid

Public Class GLRJRNL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        Set_cmbYP("RYP0", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -60, 12, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        If optStatus IsNot Nothing Then
            optStatus.Value = "A"
        End If
    End Sub

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTDETLA, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select

        End If
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

            'Case "PO Inquiry"
            '    Dim PO_ORDER_NO As String = grd.ActiveRow.Cells("PO_ORDER_NO").Text
            '    Context_Launch("Load", PO_ORDER_NO, e.Tool.Key, "POFORDRI", "F", "PO")

            'Case "Vendor Invoice Inquiry"
            '    Dim VOUCHER_NO As String = grd.ActiveRow.Cells("VOUCHER_NO").Text
            '    If VOUCHER_NO <> "" Then
            '        Context_Launch("View", VOUCHER_NO, e.Tool.Key, "APTINVHI")
            '    End If
        End Select
    End Sub

#End Region
    Protected Overrides Sub Build_Workfile()

        For Each tName As String In {"GLTJRNL1", "GLTDETL1", "GLTDETLX", "GLTDETLA", "GLTSEGM1"}
            If dst.Tables.Contains(tName) Then dst.Tables.Remove(tName)
        Next

        Dim sqlw As String = ""
        sqlw &= SQLA_filter("JOURNAL_NO", "GLTJRNL1")
        sqlw &= SQLA_filter("JOURNAL_TYPE", "GLTJRNL1")

        Dim apprStatus As String = Selected_Appr_Status()
        If apprStatus = "" Then apprStatus = "A"
        Dim useJRNL2 As Boolean = (apprStatus = "P" OrElse apprStatus = "R")

        If apprStatus <> "" Then
            sqlw &= " and NVL(GLTJRNL1.JOURNAL_APPR_STATUS,'A') = '" & apprStatus.Replace("'", "''") & "'"
        Else
            sqlw &= " and NVL(GLTJRNL1.JOURNAL_APPR_STATUS,'A') = 'A'"
        End If

        If SQLA("JOURNAL_NO") <> "" Then
            sqlw &= SQL_in("JOURNAL_NO", "GLTJRNL1.JOURNAL_NO")
        Else
            If useJRNL2 Then
                ' Pending/Rejected: period must come from header
                sqlw &= " and GLTJRNL1.OPS_YYYYPP >= '" & RYP0 & "'"
                sqlw &= " and GLTJRNL1.OPS_YYYYPP <= '" & RYP1 & "'"
            Else
                ' Approved: restrict to journals that actually posted in GLTDETL1 during the range
                sqlw &= " and GLTJRNL1.JOURNAL_NO in " &
                    " (Select Distinct JOURNAL_NO from GLTDETL1 " &
                    "  where OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "')"
            End If
        End If

        ASCMAIN1.sql = "Select * from GLTJRNL1 " & ASCMAIN1.SQL_Add_WHERE(sqlw)
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTJRNL1", 1))

        ' Pending / Rejected path
        If useJRNL2 Then

            ASCMAIN1.sql =
            "Select " & vbCrLf &
            "  GLTJRNL1.OPS_YYYYPP, " & vbCrLf &
            "  GLTJRNL2.JOURNAL_NO, " & vbCrLf &
            "  GLTJRNL2.JOURNAL_LNO, " & vbCrLf &
            "  GLTJRNL2.ACCT_CODE, " & vbCrLf &
            "  GLTJRNL2.SEG2_CODE, " & vbCrLf &
            "  GLTJRNL2.SEG3_CODE, " & vbCrLf &
            "  GLTJRNL2.SEG4_CODE, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_DATE as DETL_CTL_DATE, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(10)) as DETL_CTL_NO, " & vbCrLf &
            "  CAST(NULL AS NUMBER(6))   as DETL_CTL_LNO, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(10)) as DETL_EXE_NO, " & vbCrLf &
            "  GLTJRNL2.DETL_POSTING_AMT, " & vbCrLf &
            "  GLTJRNL2.DETL_DESC, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(8))  as DETL_EXP_CTL_NO, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(20)) as DETL_CVX_NO, " & vbCrLf &
            "  CAST(NULL AS DATE)         as DETL_CVX_REF_DATE, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(20)) as DETL_CVX_REF_NO, " & vbCrLf &
            "  CAST(NULL AS NUMBER(3))    as DETL_CVX_REF_LNO, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(1))  as DETL_CTL_TYPE, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(1))  as DETL_CVX_TYPE, " & vbCrLf &
            "  GLTJRNL2.COST_CTR_CODE " & vbCrLf &
            "from GLTJRNL1, GLTJRNL2 " & vbCrLf &
            "where GLTJRNL1.JOURNAL_NO = GLTJRNL2.JOURNAL_NO " & vbCrLf &
            sqlw

            Dim dtDetl As DataTable = ASCDATA1.GetDataTable("", "GLTDETL1", 3)
            dtDetl.TableName = "GLTDETL1"
            dst.Tables.Add(dtDetl)

            ASCMAIN1.sql =
            "Select " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(1))  as DETL_CVX_TYPE, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(20)) as DETL_CVX_NO, " & vbCrLf &
            "  CAST(NULL AS VARCHAR2(60)) as DETL_CVX_NAME " & vbCrLf &
            "from dual where 1=0"
            dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLX", 2))

            'Build GLTDETLA for Pending/Rejected using GLTJRNL2
            ASCMAIN1.sql =
            "Select " & vbCrLf &
            "  GLTJRNL1.OPS_YYYYPP, " & vbCrLf &
            "  GLTJRNL2.JOURNAL_NO, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_DESC, " & vbCrLf &
            "  NVL(GLTJRNL1.JOURNAL_APPR_STATUS,'A') as JOURNAL_APPR_STATUS, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_APPR_BY, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_APPR_DATE, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_APPR2_BY, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_APPR2_DATE, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_REJ_REASON, " & vbCrLf &
            "  DECODE(GLTJRNL1.JOURNAL_REVERSED_IND,'1','Reversed By','2','Reversing','') REVERSING_NOTE, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_REVERSED, " & vbCrLf &
            "  GLTJRNL2.JOURNAL_LNO, " & vbCrLf &
            "  GLTJRNL2.ACCT_CODE, " & vbCrLf &
            "  GLTACCT1.ACCT_DESC, " & vbCrLf &
            "  GLTJRNL2.SEG2_CODE, " & vbCrLf &
            "  GLTJRNL2.SEG3_CODE, " & vbCrLf &
            "  GLTJRNL2.SEG4_CODE, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_DATE as DETL_CTL_DATE, " & vbCrLf &
            "  GLTJRNL2.DETL_POSTING_AMT, " & vbCrLf &
            "  GLTJRNL2.DETL_DESC, " & vbCrLf &
            "  GLTJRNL1.INIT_OPER, " & vbCrLf &
            "  GLTJRNL1.INIT_DATE, " & vbCrLf &
            "  GLTJRNL1.JOURNAL_TYPE, " & vbCrLf &
            "  GLTJRNL1.OPS_YYYYPP YP_BOOKED, " & vbCrLf &
            "  GLTJRNL1.OPS_YYYYPP_REV YP_REVERSED " & vbCrLf &
            "from GLTJRNL1, GLTJRNL2, GLTACCT1 " & vbCrLf &
            "where GLTJRNL1.JOURNAL_NO = GLTJRNL2.JOURNAL_NO " & vbCrLf &
            "  and GLTACCT1.ACCT_CODE = GLTJRNL2.ACCT_CODE " & vbCrLf &
            sqlw
            dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLA", 0))

            grdGLTDETLA.DataSource = dst.Tables("GLTDETLA")
            Sort_grdColumns(grdGLTDETLA, "OPS_YYYYPP,JOURNAL_NO,JOURNAL_LNO")
            Dim apprCaption As String = ApprStatusCaption(apprStatus)
            grdGLTDETLA.Text = "Journals (" & apprCaption & ") between " & RYPLEGEND0 & " and " & RYPLEGEND1

            ASCMAIN1.Add_Value_List(grdGLTDETLA, "JOURNAL_APPR_STATUS", , New String() {
            "P:Pending",
            "A:Approved",
            "R:Rejected",
            "2:Pending 2"
        })

            If grdGLTDETLA.DisplayLayout.Bands(0).Summaries.Count = 0 Then
                Create_Summary(grdGLTDETLA, "OPS_YYYYPP", "Count")
                Create_Summary(grdGLTDETLA, "DETL_POSTING_AMT")
                Set_SEGS(grdGLTDETLA, "GLTDETLA")
            End If

            For Each gcol As UltraWinGrid.UltraGridColumn In grdGLTDETLA.DisplayLayout.Bands(0).Columns
                gcol.Header.Appearance.BackColor = System.Drawing.Color.White
                gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "YP_BOOKED" OrElse gcol.Key = "YP_REVERSED" Then
                    gcol.CellAppearance.TextHAlign = HAlign.Center
                    gcol.Header.Appearance.TextHAlign = HAlign.Center
                End If
            Next

            If grdGLTDETLA.DisplayLayout.Bands(0).Columns.Exists("JOURNAL_APPR_BY") Then
                grdGLTDETLA.DisplayLayout.Bands(0).Columns("JOURNAL_APPR_BY").Hidden = False ' (apprStatus = "P" Or apprStatus = "2")
            End If
            If grdGLTDETLA.DisplayLayout.Bands(0).Columns.Exists("JOURNAL_APPR_DATE") Then
                grdGLTDETLA.DisplayLayout.Bands(0).Columns("JOURNAL_APPR_DATE").Hidden = False ' (apprStatus = "P" Or apprStatus = "2")
            End If
            If grdGLTDETLA.DisplayLayout.Bands(0).Columns.Exists("JOURNAL_APPR2_BY") Then
                grdGLTDETLA.DisplayLayout.Bands(0).Columns("JOURNAL_APPR2_BY").Hidden = False ' (apprStatus = "P" Or apprStatus = "2")
            End If
            If grdGLTDETLA.DisplayLayout.Bands(0).Columns.Exists("JOURNAL_APPR2_DATE") Then
                grdGLTDETLA.DisplayLayout.Bands(0).Columns("JOURNAL_APPR2_DATE").Hidden = False '  (apprStatus = "P" Or apprStatus = "2")
            End If
            If grdGLTDETLA.DisplayLayout.Bands(0).Columns.Exists("JOURNAL_REJ_REASON") Then
                grdGLTDETLA.DisplayLayout.Bands(0).Columns("JOURNAL_REJ_REASON").Hidden = False ' (apprStatus = "P" Or apprStatus = "2")
            End If

            Get_WKCodes("GLTDETL1", "ACCT_CODE", "GLTACCT1", "*")
            Get_WKCodes("GLTDETL1", "OPS_YYYYPP", "GLTPARM2", "*")
            dst.Tables.Add(ASCDATA1.GetDataTable("*", "GLTSEGM1"))

            Prepare_GL_Account_Activity_Recaps("GLTDETL1")
            Exit Sub
        End If

        ASCMAIN1.sql = "Select GLTDETL1.* from GLTDETL1,GLTJRNL1 " &
                   " where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO " & sqlw
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETL1", 3))

        ASCMAIN1.sql = "SELECT Distinct GLTDETL1.DETL_CVX_TYPE, GLTDETL1.DETL_CVX_NO " &
                   ", APTVEND1.VEND_NAME DETL_CVX_NAME" &
                   " from GLTDETL1,GLTJRNL1,APTVEND1" &
                   " where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO " & sqlw &
                   "   and GLTDETL1.DETL_CVX_TYPE = 'V' " &
                   "   and APTVEND1.VEND_CODE = GLTDETL1.DETL_CVX_NO" &
                   " UNION " &
                   "SELECT Distinct GLTDETL1.DETL_CVX_TYPE, GLTDETL1.DETL_CVX_NO " &
                   ", ARTCUST1.CUST_NAME DETL_CVX_NAME" &
                   " from GLTDETL1,GLTJRNL1,ARTCUST1" &
                   " where GLTDETL1.JOURNAL_NO = GLTJRNL1.JOURNAL_NO " & sqlw &
                   "   and GLTDETL1.DETL_CVX_TYPE = 'C' " &
                   "   and ARTCUST1.CUST_CODE = GLTDETL1.DETL_CVX_NO"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLX", 2))

        ASCMAIN1.sql =
    "Select GLTDETL1.OPS_YYYYPP,GLTDETL1.JOURNAL_NO,GLTJRNL1.JOURNAL_DESC," &
    "NVL(GLTJRNL1.JOURNAL_APPR_STATUS,'A') as JOURNAL_APPR_STATUS," &
    "GLTJRNL1.JOURNAL_APPR_BY, GLTJRNL1.JOURNAL_APPR_DATE," &
    "  GLTJRNL1.JOURNAL_APPR2_BY, " & vbCrLf &
    "  GLTJRNL1.JOURNAL_APPR2_DATE, " & vbCrLf &
    "  GLTJRNL1.JOURNAL_REJ_REASON, " & vbCrLf &
    "DECODE(JOURNAL_REVERSED_IND,'1','Reversed By','2','Reversing','') REVERSING_NOTE, " &
    "JOURNAL_REVERSED,GLTDETL1.JOURNAL_LNO" & vbCrLf &
    ",GLTDETL1.ACCT_CODE,GLTACCT1.ACCT_DESC,GLTDETL1.SEG2_CODE,GLTDETL1.SEG3_CODE,GLTDETL1.SEG4_CODE" & vbCrLf &
    ",GLTDETL1.DETL_CTL_DATE,GLTDETL1.DETL_POSTING_AMT,GLTDETL1.DETL_DESC,GLTJRNL1.INIT_OPER,GLTJRNL1.INIT_DATE,GLTJRNL1.JOURNAL_TYPE" & vbCrLf &
    ",GLTJRNL1.OPS_YYYYPP YP_BOOKED, GLTJRNL1.OPS_YYYYPP_REV YP_REVERSED" & vbCrLf &
    " from GLTDETL1,GLTJRNL1,GLTACCT1" & vbCrLf &
    " where GLTACCT1.ACCT_CODE = GLTDETL1.ACCT_CODE" & vbCrLf &
    "   and GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf &
    "   and GLTDETL1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf &
    "   and GLTDETL1.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf &
    sqlw
        dst.Tables.Add(ASCDATA1.GetDataTable("", "GLTDETLA", 0))

        grdGLTDETLA.DataSource = dst.Tables("GLTDETLA")
        Sort_grdColumns(grdGLTDETLA, "OPS_YYYYPP,JOURNAL_NO,JOURNAL_LNO")
        grdGLTDETLA.Text = "GL Details for Approved Journals posted between " & RYPLEGEND0 & " and " & RYPLEGEND1

        ASCMAIN1.Add_Value_List(grdGLTDETLA, "JOURNAL_APPR_STATUS", , New String() {
        "P:Pending",
        "A:Approved",
        "R:Rejected",
        "2:Pending 2"
    })

        If grdGLTDETLA.DisplayLayout.Bands(0).Summaries.Count = 0 Then
            Create_Summary(grdGLTDETLA, "OPS_YYYYPP", "Count")
            Create_Summary(grdGLTDETLA, "DETL_POSTING_AMT")
            Set_SEGS(grdGLTDETLA, "GLTDETLA")
        End If

        For Each gcol As UltraWinGrid.UltraGridColumn In grdGLTDETLA.DisplayLayout.Bands(0).Columns
            gcol.Header.Appearance.BackColor = System.Drawing.Color.White
            gcol.Header.Appearance.BackColor2 = System.Drawing.Color.LightGreen
            gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            If gcol.Key = "YP_BOOKED" OrElse gcol.Key = "YP_REVERSED" Then
                gcol.CellAppearance.TextHAlign = HAlign.Center
                gcol.Header.Appearance.TextHAlign = HAlign.Center
            End If
        Next

        Get_WKCodes("GLTDETL1", "ACCT_CODE", "GLTACCT1", "*")
        Get_WKCodes("GLTDETL1", "OPS_YYYYPP", "GLTPARM2", "*")
        dst.Tables.Add(ASCDATA1.GetDataTable("*", "GLTSEGM1"))

        Prepare_GL_Account_Activity_Recaps("GLTDETL1")

    End Sub


    Public Overrides Sub Print_Report()
        Dim appr As String = Selected_Appr_Status()
        If appr = "" Then appr = "A"

        Dim subt As String
        Select Case appr
            Case "A" : subt = "Approved Only"
            Case "P" : subt = "Pending Only"
            Case "R" : subt = "Rejected Only"
            Case "2" : subt = "Pending 2 Only"
            Case Else : subt = ""
        End Select

        CR_params.Add("SUBT", subt)

        CR_params.Add("SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
        CR_params.Add("SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
        CR_params.Add("SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")
        CR_params.Add("SHOW_JRNL_COMMENTS", IIf(Absx1.chkFor("SHOW_JRNL_COMMENTS").Checked, "1", "0"))
        CR_params.Add("SHOW_DETL_DESC", IIf(Absx1.chkFor("SHOW_DETL_DESC").Checked, "1", "0"))
        CR_params.Add("SHOW_CVX_NAME", IIf(Absx1.chkFor("SHOW_CVX_NAME").Checked, "1", "0"))
        CR_params.Add("PAGE_BREAK", IIf(Absx1.chkFor("PAGE_BREAK").Checked, "1", "0"))
        CR_params.Add("ACCT_RECAPS", ROWs("GLTPARM1").Item("GL_PARM_ACCT_RECAPS") & "")
        Generate_Report(RPT)

        grdGLTDETLA.Visible = True
    End Sub

    Private Sub grdGLTDETLA_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTDETLA.InitializeLayout

    End Sub

    Private Sub grdGLTDETLA_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdGLTDETLA.InitializeRow
        Dim OPS_YYYYPP As String = e.Row.Cells("OPS_YYYYPP").Value & ""
        Dim YP_BOOKED As String = e.Row.Cells("YP_BOOKED").Value & ""
        If OPS_YYYYPP <> YP_BOOKED Then
            e.Row.Cells("YP_BOOKED").Appearance.ForeColor = System.Drawing.Color.Red
        End If
    End Sub
    Private Function Selected_Appr_Status() As String
        If optStatus Is Nothing OrElse optStatus.Value Is Nothing Then Return ""

        Dim s As String = optStatus.Value.ToString().Trim().ToUpperInvariant()
        If s = "A" OrElse s = "P" OrElse s = "R" OrElse s = "2" Then
            Return s
        End If

        Return ""
    End Function
    Private Function ApprStatusCaption(apprStatus As String) As String
        Select Case (apprStatus & "").Trim().ToUpperInvariant()
            Case "P" : Return "Pending"
            Case "A" : Return "Approved"
            Case "R" : Return "Rejected"
            Case "2" : Return "Pending 2"
            Case Else : Return ""
        End Select
    End Function

End Class