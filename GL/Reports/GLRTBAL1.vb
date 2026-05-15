Public Class GLRTBAL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Get_PARM("GLTPARM1")
        For i As Integer = 2 To 4
            Dim z As String = "SEG" & CStr(i)
            If ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & "" = "" Then
                Absx1.CtlFor(z & "_CODE").Visible = False
            Else
                Absx1.CtlFor(z & "_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_" & z & "_DESC") & ""
                Absx1.chkFor(z & "_CODE").Checked = True
            End If
        Next

        'Set_cmbYP("RYP", ROWs("GLTPARM1").Item("GL_PARM_CURRENT_YYYYPP") & "", -60, 12, 0)
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 12, 0)

    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables

        Dim TTT As String = Prepare_Work_File()

        sql = "Select TTT.*, GLTACCT1.ACCT_CLASS_CODE" & vbCrLf _
            & ", GLTSEGM2.ACCT_SEG_CLASS SEG2_CLASS_CODE" & vbCrLf _
            & ", GLTSEGM3.ACCT_SEG_CLASS SEG3_CLASS_CODE" & vbCrLf _
            & ", GLTSEGM4.ACCT_SEG_CLASS SEG4_CLASS_CODE" & vbCrLf _
            & " from " & TTT & " TTT, GLTACCT1, GLTSEGM1 GLTSEGM2, GLTSEGM1 GLTSEGM3, GLTSEGM1 GLTSEGM4" & vbCrLf _
            & " where GLTACCT1.ACCT_CODE (+) = TTT.ACCT_CODE" & vbCrLf _
            & "   and GLTSEGM2.ACCT_SEG_ID (+) = '2' and GLTSEGM2.ACCT_SEG_CODE (+) = TTT.SEG2_CODE" & vbCrLf _
            & "   and GLTSEGM3.ACCT_SEG_ID (+) = '3' and GLTSEGM3.ACCT_SEG_CODE (+) = TTT.SEG3_CODE" & vbCrLf _
            & "   and GLTSEGM4.ACCT_SEG_ID (+) = '4' and GLTSEGM4.ACCT_SEG_CODE (+) = TTT.SEG4_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCTX", 4))

        sql = "Select GLTACCT1.* from GLTACCT1,(SELECT DISTINCT ACCT_CODE FROM " & TTT & ") TTT where GLTACCT1.ACCT_CODE = TTT.ACCT_CODE"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1", 1))


        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        sql = "Select " & sql_SELECT_cols _
            & ", GLTACCT3.ACCT_CODE, GLTACCT3.SEG2_CODE, GLTACCT3.SEG3_CODE, GLTACCT3.SEG4_CODE" _
            & ", 1 AS COUNTER"
        sql = sql & " from " & TTT & " GLTACCT3 " & sql_TABLE_NAMEs
        sql = sql & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter)
        'sql = sql & " group by " & sql_GROUP_BY_cols & y
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        Call Special_Routines_for_ACCT_TYPE()

    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("SEG2_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "")
        CR_params.Add("SEG3_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "")
        CR_params.Add("SEG4_DESC", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "")

        CR_params.Add("CHKSEG2", IIf(Absx1.chkFor("SEG2_CODE").Checked, "1", "0"))
        CR_params.Add("CHKSEG3", IIf(Absx1.chkFor("SEG3_CODE").Checked, "1", "0"))
        CR_params.Add("CHKSEG4", IIf(Absx1.chkFor("SEG4_CODE").Checked, "1", "0"))

        CR_params.Add("GYPLEGEND", ASCMAIN1.Get_Legend(RYP))
        Generate_Report(RPT)

        Prepare_Data_Extracts()

    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DataSource = dst.Tables("GLTACCTX")
        grdASTEXPT1.Text = "Trial Balance Extract"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Set_DX_Column(grdASTEXPT1, "ACCT_CODE", "Account", 80)
        Set_DX_Column(grdASTEXPT1, "SEG2_CODE", ROWs("GLTPARM1").Item("GL_PARM_SEG2_DESC") & "", 80)
        Set_DX_Column(grdASTEXPT1, "SEG3_CODE", ROWs("GLTPARM1").Item("GL_PARM_SEG3_DESC") & "", 80)
        Set_DX_Column(grdASTEXPT1, "SEG4_CODE", ROWs("GLTPARM1").Item("GL_PARM_SEG4_DESC") & "", 80)
        Set_DX_Column(grdASTEXPT1, "ACCT_TYPE", "Type", 80)
        Set_DX_Column(grdASTEXPT1, "TY_BEG_BAL", "TY Beg Bal", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TY_MTD_ACT", "TY MTD Act", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TY_YTD_ACT", "TY YTD Act", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TY_MTD_BEG", "TYTM Beg Bal", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TY_MTD_END", "TYTM End Bal", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "LY_BEG_BAL", "LY Beg Bal", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "LY_MTD_ACT", "LY MTD Act", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "LY_YTD_ACT", "LY YTD Act", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "LY_MTD_BEG", "LYTM Beg Bal", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "LY_MTD_END", "LYTM End Bal", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TY_MTD_DR", "TY MTD DR", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "TY_MTD_CR", "TY MTD CR", 100, "#,##0", "Sum", System.Drawing.Color.LightBlue)

        Set_DX_Column(grdASTEXPT1, "ACCT_CLASS_CODE", "Acct Class", 80, , , System.Drawing.Color.LightGreen)

        Set_DX_Column(grdASTEXPT1, "SEG2_CLASS_CODE", ROWs("GLTPARM1").Item("GL_PARM_SEG2_CLASS_DESC") & "", 80, , , System.Drawing.Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "SEG3_CLASS_CODE", ROWs("GLTPARM1").Item("GL_PARM_SEG3_CLASS_DESC") & "", 80, , , System.Drawing.Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "SEG4_CLASS_CODE", ROWs("GLTPARM1").Item("GL_PARM_SEG4_CLASS_DESC") & "", 80, , , System.Drawing.Color.LightGreen)
 
        
        'ACCT_CODE	SEG2_CODE	SEG3_CODE	SEG4_CODE	ACCT_TYPE	TY_BEG_BAL	TY_MTD_ACT	TY_YTD_ACT	TY_MTD_BEG	TY_MTD_END	LY_BEG_BAL	LY_MTD_ACT	LY_YTD_ACT	LY_MTD_BEG	LY_MTD_END	TY_MTD_DR	TY_MTD_CR	ACCT_CLASS_CODE	SEG2_CLASS_CODE	SEG3_CLASS_CODE	SEG4_CLASS_CODE

        For Each C As String In New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"}
            grdASTEXPT1.DisplayLayout.Bands(0).Columns(C).Header.Fixed = True
        Next
        'For Each C As String In New String() {"BEG", "SHP", "RTN", "REC", "ADJ", "XFR", "CON", "RTV", "PHY", "ON_HAND", "ONPO", "OPEN", "PICK", "COMM"}
        '    With grdASTEXPT1.DisplayLayout.Bands(0)
        '        For Each CTYP As String In New String() {"QTY", "CST"}
        '            With .Columns("WHSE_" & CTYP & "_" & C)
        '                .Width = 80
        '                .Hidden = False
        '                .Format = "#,##0"
        '                .Header.Caption = CTYP & " " & C
        '                If CTYP = "QTY" Then
        '                    .Header.Appearance.BackColor2 = Color.LightBlue
        '                Else
        '                    .Header.Appearance.BackColor2 = Color.LightGreen
        '                End If
        '            End With
        '            Create_Summary(grdASTEXPT1, "WHSE_" & CTYP & "_" & C)
        '        Next
        '    End With
        'Next

        Sort_grdColumns(grdASTEXPT1, "ACCT_CODE")
        UltraTabControl1.Tabs("Data Exports").Visible = True

    End Sub


    Function Prepare_Work_File() As String

        Dim TY As String = Mid$(RYP, 1, 4)
        Dim LY As String = Mid$(ASCMAIN1.Period_Calc(RYP, -12), 1, 4)
        Dim P As Integer = Val(Mid(RYP, 5, 2))

        Dim TT As String = GL_Prep(LY, TY)


        Dim BY_SEG2 As Boolean = Absx1.chkFor("SEG2_CODE").Checked Or SQLA("SEG2_CODE", "SEQUENCE") <> ""
        Dim BY_SEG3 As Boolean = Absx1.chkFor("SEG3_CODE").Checked Or SQLA("SEG3_CODE", "SEQUENCE") <> ""
        Dim BY_SEG4 As Boolean = Absx1.chkFor("SEG4_CODE").Checked Or SQLA("SEG4_CODE", "SEQUENCE") <> ""

        Dim sqlx As String = ""
        Dim sqlx_group_by As String = ""
        Dim z As String
        Dim i As Integer
        For i = 2 To 4
            z = "SEG" & CStr(i) & "_CODE"
            If Not New Boolean() {BY_SEG2, BY_SEG3, BY_SEG4}(i - 2) Then
                z = "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG" & CStr(i)) & "' " & z
            Else
                sqlx_group_by = sqlx_group_by & ", X." & z
            End If
            sqlx = sqlx & ", " & z
        Next


        'sql = ""
        'z = SQLA("ACCT_CODE", "CODE_VALUES", True)
        'If z <> "" Then
        '    sql &= " AND X.ACCT_CODE IN (" & z & ")" & vbCr
        'End If
        'z = SQLA("SEG2_CODE", "CODE_VALUES", True)
        'If z <> "" Then
        '    sql &= " AND X.SEG2_CODE IN (" & z & ")" & vbCr
        'End If
        'z = SQLA("SEG3_CODE", "CODE_VALUES", True)
        'If z <> "" Then
        '    sql &= " AND X.SEG3_CODE IN (" & z & ")" & vbCr
        'End If
        'z = SQLA("SEG4_CODE", "CODE_VALUES", True)
        'If z <> "" Then
        '    sql &= " AND X.SEG4_CODE IN (" & z & ")" & vbCr
        'End If
        'z = SQLA("ACCT_TYPE", "CODE_VALUES", True)
        'If z <> "" Then
        '    sql &= " AND X.ACCT_TYPE IN (" & z & ")" & vbCr
        'End If
        'Dim sqlw As String = sql

        Dim sqlw As String = ""
        For Each COLUMN_NAME As String In New String() {"ACCT_CODE", "SEG2_CODE", "SEG3_CODE", "SEG4_CODE", "ACCT_TYPE"}
            sqlw &= SQLA_filter(COLUMN_NAME, "X")
        Next

        Dim MTD As String = "+NVL(ACCT_ACT_P" & Format(P, "00") & ",0)"
        Dim YTD_BEG As String = ""
        If P > 1 Then
            For i = 1 To P - 1
                YTD_BEG = YTD_BEG & "+NVL(ACCT_ACT_P" & Format(i, "00") & ",0)"
            Next
        End If

        sql = ""

        sql &= "Select X.ACCT_CODE" & sqlx & ", X.ACCT_TYPE" & vbCr
        sql &= ", SUM (NVL(ACCT_BEG_BAL,0)) TY_BEG_BAL" & vbCr
        sql &= ", SUM (" & Mid(MTD, 2) & ") TY_MTD_ACT" & vbCr
        sql &= ", SUM (" & Mid(YTD_BEG, 2) & MTD & ") TY_YTD_ACT" & vbCr
        sql &= ", SUM (NVL(ACCT_BEG_BAL,0)" & YTD_BEG & ") TY_MTD_BEG" & vbCr
        sql &= ", SUM (NVL(ACCT_BEG_BAL,0)" & YTD_BEG & MTD & ") TY_MTD_END" & vbCr
        sql &= ", 0 LY_BEG_BAL, 0 LY_MTD_ACT, 0 LY_YTD_ACT, 0 LY_MTD_BEG, 0 LY_MTD_END " & vbCr
        sql &= ", 0 TY_MTD_DR, 0 TY_MTD_CR" & vbCr
        sql &= " from " & TT & " X WHERE ACCT_YEAR = '" & TY & "'" & vbCr
        sql &= sqlw
        sql &= " group by X.ACCT_CODE" & sqlx_group_by & ", X.ACCT_TYPE" & vbCr

        sql &= " union " & vbCr

        sql &= "Select X.ACCT_CODE" & sqlx & ", X.ACCT_TYPE" & vbCr
        sql &= ", 0 TY_BEG_BAL, 0 TY_MTD_ACT, 0 TY_YTD_ACT, 0 TY_MTD_BEG, 0 TY_MTD_END " & vbCr
        sql &= ", SUM (NVL(ACCT_BEG_BAL,0)) LY_BEG_BAL" & vbCr
        sql &= ", SUM (" & Mid(MTD, 2) & ") LY_MTD_ACT" & vbCr
        sql &= ", SUM (" & Mid(YTD_BEG, 2) & MTD & ") LY_YTD_ACT" & vbCr
        sql &= ", SUM (NVL(ACCT_BEG_BAL,0)" & YTD_BEG & ") LY_MTD_BEG" & vbCr
        sql &= ", SUM (NVL(ACCT_BEG_BAL,0)" & YTD_BEG & MTD & ") LY_MTD_END" & vbCr
        sql &= ", 0 TY_MTD_DR, 0 TY_MTD_CR" & vbCr
        sql &= " from " & TT & " X WHERE ACCT_YEAR = '" & LY & "'" & vbCr
        sql &= sqlw
        sql &= " group by X.ACCT_CODE" & sqlx_group_by & ", X.ACCT_TYPE" & vbCr

        sql &= " union " & vbCr

        sql &= "Select X.ACCT_CODE" & sqlx & ", GLTACCT1.ACCT_TYPE" & vbCr
        sql &= ", 0 TY_BEG_BAL, 0 TY_MTD_ACT, 0 TY_YTD_ACT, 0 TY_MTD_BEG, 0 TY_MTD_END " & vbCr
        sql &= ", 0 LY_BEG_BAL, 0 LY_MTD_ACT, 0 LY_YTD_ACT, 0 LY_MTD_BEG, 0 LY_MTD_END " & vbCr
        sql &= ", SUM (CASE WHEN X.DETL_POSTING_AMT > 0 THEN X.DETL_POSTING_AMT ELSE 0 END) TY_MTD_DR" & vbCr
        sql &= ", SUM (CASE WHEN X.DETL_POSTING_AMT < 0 THEN X.DETL_POSTING_AMT ELSE 0 END) TY_MTD_CR" & vbCr
        sql &= "from GLTDETL1 X,GLTACCT1 WHERE X.OPS_YYYYPP = '" & RYP & "'" & vbCr
        sql &= Replace(sqlw, "X.ACCT_TYPE", "GLTACCT1.ACCT_TYPE")
        sql &= "and GLTACCT1.ACCT_CODE (+) = X.ACCT_CODE" & vbCr
        sql &= "group by X.ACCT_CODE" & sqlx_group_by & ", GLTACCT1.ACCT_TYPE" & vbCr

        sql = "Select ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE " & vbCr _
            & ", SUM (TY_BEG_BAL) TY_BEG_BAL, SUM (TY_MTD_ACT) TY_MTD_ACT, SUM (TY_YTD_ACT) TY_YTD_ACT, SUM (TY_MTD_BEG) TY_MTD_BEG, SUM (TY_MTD_END) TY_MTD_END " & vbCr _
            & ", SUM (LY_BEG_BAL) LY_BEG_BAL, SUM (LY_MTD_ACT) LY_MTD_ACT, SUM (LY_YTD_ACT) LY_YTD_ACT, SUM (LY_MTD_BEG) LY_MTD_BEG, SUM (LY_MTD_END) LY_MTD_END " & vbCr _
            & ", SUM (TY_MTD_DR) TY_MTD_DR, SUM (TY_MTD_CR) TY_MTD_CR " & vbCr _
            & " from (" & vbCr & sql & ") group by ACCT_CODE, SEG2_CODE, SEG3_CODE, SEG4_CODE, ACCT_TYPE" & vbCr

        sql = "Select * from (" & sql & ") where NVL(TY_MTD_BEG,0) <> 0 or NVL(TY_MTD_DR,0) <> 0 or NVL(TY_MTD_CR,0) <> 0 or NVL(TY_MTD_END,0) <> 0"
        Dim TTT As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & TTT & " Add Primary Key (ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE)")
        Call ASCMAIN1.AnalyzeTable(TTT)

        Return TTT

    End Function
End Class