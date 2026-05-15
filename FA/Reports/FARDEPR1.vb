Imports System.Drawing

Public Class FARDEPR1

#Region "General Declarations"
    Dim FATFAMF1 As String
    Dim FATFATR1 As String
    Dim FATFADP1 As String
    Dim JOURNAL_NO As String

    Dim sqlFATFAMF1 As String
    Dim sqlFATFATR1 As String

    Dim reprint As Boolean
    Dim RYP_DISPOSE_THRU As String = ""

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If MENU_ITEM_OBJECT = "FARDEPRI" Then
            InquiryMode = True
            RWU = "N"
        End If
        Get_PARM("FATPARM1")
        Get_PARM("GLTPARM1")

        'Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        Dim YP As String = ASCMAIN1.Period_Calc(ROWs("FATPARM1").Item("FA_PARM_DEPR_LAST_YP_UPDATED"), 1)
        Set_cmbYP("RYP", YP, 0, 0, 0)

        cmbYP_Dispose_Thru.Enabled = False

        If ROWs("FATPARM1").Item("FA_PARM_ASSET_DISPOSAL") = "N" Then
            grpDisposeAssets.Visible = False
        Else
            grpDisposeAssets.Visible = True
            cmbYP_Dispose_Thru.Enabled = False

            ASCMAIN1.sql = "Select MIN (FATFATR1.OPS_YYYYPP) YPMIN, MAX (FATFATR1.OPS_YYYYPP) YPMAX" & vbCrLf _
            & " from FATFATR1,FATFAMF1" & vbCrLf _
            & " where FATFAMF1.ASSET_NO = FATFATR1.ASSET_NO" & vbCrLf _
            & "   and FATFAMF1.ASSET_STATUS = 'D'"
            Dim rowDisposalThru As DataRow = ASCDATA1.GetDataRow
            If rowDisposalThru Is Nothing Then
                grpDisposeAssets.Visible = False
            Else
                grpDisposeAssets.Visible = True
                Dim YPMIN As String = rowDisposalThru.Item("YPMIN")
                Dim YPMAX As String = rowDisposalThru.Item("YPMAX")
                Dim pc As Integer = ASCMAIN1.Period_Diff(YPMAX, YPMIN)
                Set_cmbYP("RYP_DISPOSE_THRU", YPMAX, pc, 0, 0)
            End If
        End If

    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "R"
        Dim sqlw As String = ""
        Prepare_dst(True, sqlw)
        Check_if_Empty("FATFATR1")
    End Sub

    Public Overrides Sub Print_Report()

        SUBT = ""
        If reprint Then SUBT = "Re-Print"

        If Not InquiryMode Then

            CR_params.Add("SUBT", SUBT)
            Generate_Report(RPT, , SUBT, "{FATFAMF1.ASSET_STATUS} = 'D' or {FATFAMF1.ASSET_STATUS} = 'C'") '{FATFAMF1.ASSET_ACTION} = 'D' or {FATFAMF1.ASSET_ACTION} = 'W'")

            If chkAssetDisposal.Checked Then
                'If SUBT = "" Then
                '    SUBT = "Asset Disposal"
                'Else
                '    SUBT &= " - Asset Disposal"
                'End If

                CR_params.Add("SUBT", SUBT)
                Generate_Report(RPT, "Asset Disposal", SUBT, "{FATFAMF1.ASSET_ACTION} = 'Z'")
            End If

            Print_GL()
        End If

        Prepare_Data_Extracts()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            ASCMAIN1.sql = "Select APTINVH1.VOUCHER_NO, APTINVH2.VOUCHER_LNO, APTINVH1.OPS_YYYYPP" & vbCrLf _
                & ", APTINVH1.VEND_CODE, APTVEND1.VEND_NAME, APTINVH1.INV_DATE, APTINVH1.INV_NUM" & vbCrLf _
                & ", APTINVH1.INV_STATUS, APTINVH1.INV_REF" & vbCrLf _
                & ", APTINVH2.INV_LINE_AMT" & vbCrLf _
                & ", APTINVH2.ACCT_CODE, X.ASSET_CLASS_CODE" & vbCrLf _
                & " from APTINVH1, APTINVH2, APTVEND1, (Select ACCT_CODE_CAP ACCT_CODE, MIN (ASSET_CLASS_CODE) ASSET_CLASS_CODE from FATFACL1 group by ACCT_CODE_CAP) X" & vbCrLf _
                & " where APTINVH1.VOUCHER_NO = APTINVH2.VOUCHER_NO" & vbCrLf _
                & "   and APTINVH1.OPS_YYYYPP >= :PARM1" & vbCrLf _
                & "   and APTINVH1.OPS_YYYYPP <= :PARM2" & vbCrLf _
                & "   and APTINVH2.ACCT_CODE IN (Select ACCT_CODE_CAP from FATFACL1)" & vbCrLf _
                & "   and X.ACCT_CODE = APTINVH2.ACCT_CODE" & vbCrLf _
                & "   and APTVEND1.VEND_CODE = APTINVH1.VEND_CODE" & vbCrLf _
                & " UNION" & vbCrLf _
                & "Select 'GLJE' || GLTDETL1.JOURNAL_NO, GLTDETL1.JOURNAL_LNO, GLTDETL1.OPS_YYYYPP" & vbCrLf _
                & ", NULL VEND_CODE, NULL VEND_NAME, GLTDETL1.DETL_CTL_DATE INV_DATE, GLTDETL1.DETL_CVX_REF_NO INV_NUM" & vbCrLf _
                & ", NULL INV_STATUS, GLTJRNL1.JOURNAL_DESC INV_REF" & vbCrLf _
                & ", GLTDETL1.DETL_POSTING_AMT INV_LINE_AMT" & vbCrLf _
                & ", GLTDETL1.ACCT_CODE, X.ASSET_CLASS_CODE" & vbCrLf _
                & " from GLTDETL1,GLTJRNL1, (Select ACCT_CODE_CAP ACCT_CODE, MIN (ASSET_CLASS_CODE) ASSET_CLASS_CODE from FATFACL1 group by ACCT_CODE_CAP) X" & vbCrLf _
                & " where GLTJRNL1.JOURNAL_NO = GLTDETL1.JOURNAL_NO" & vbCrLf _
                & "   and GLTJRNL1.JOURNAL_TYPE = 'GLJE'" & vbCrLf _
                & "   and GLTDETL1.OPS_YYYYPP >=  :PARM3" & vbCrLf _
                & "   and GLTDETL1.OPS_YYYYPP <=  :PARM4" & vbCrLf _
                & "   and GLTDETL1.ACCT_CODE IN (Select ACCT_CODE_CAP from FATFACL1)" & vbCrLf _
                & "   and X.ACCT_CODE = GLTDETL1.ACCT_CODE"
            ASCMAIN1.sql = $"Select X.*, F.ASSET_AMT from ({ASCMAIN1.sql}) X, FATFAMF2" & vbCrLf _
                & ", (Select VOUCHER_NO, VOUCHER_LNO, SUM (ASSET_AMT) ASSET_AMT from FATFAMF1 group by VOUCHER_NO, VOUCHER_LNO) F" & vbCrLf _
                & " where X.VOUCHER_NO = F.VOUCHER_NO (+) And X.VOUCHER_LNO = F.VOUCHER_LNO (+)" & vbCrLf _
                & "  And x.VOUCHER_NO = FATFAMF2.VOUCHER_NO(+) And x.VOUCHER_LNO = FATFAMF2.VOUCHER_LNO(+) And NVL(FATFAMF2.IGNORE,'0') <> '1'"
            Dim YP As String = ROWs("FATPARM1").Item("FA_PARM_DEPR_LAST_YP_UPDATED")
            Dim YP1 As String = ASCMAIN1.Period_Calc(YP, -1)
            Dim YP2 As String = ASCMAIN1.Period_Calc(YP, 1)
            Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "", "VVVV", New String() {YP1, YP2, YP1, YP2})

            tbl.Columns.Add("ASSET_REQ", GetType(System.Decimal), "ISNULL(INV_LINE_AMT,0) - ISNULL(ASSET_AMT,0)")
            If tbl.Select("ASSET_REQ <> 0").Length <> 0 Then
                EMsg &= "There are AP/GL Distributions to Fixed Asset Accounts that have not been assigned to Fixed Asset records"
            End If
        End If

    End Sub

    Overrides Sub Update_Record()

        If ASCMAIN1.Running_in_VS Then
            Stop
        End If

        ASCMAIN1.sql = "Insert into FATFATR1 Select * from " & FATFATR1 '& "where " & FATFATR1 & ".ASSET_ACTION in ('D', 'W')"
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.sql = $"Insert into FATFADP1 Select '{RYP}' OPS_YYYYPP, ASSET_NO, ASSET_AMT, ASSET_DEP, ASSET_WOF, ASSET_BAL, ASSET_DEP_EXP, ASSET_DEP_WOF, ASSET_ACTION, '{XNO}' JOURNAL_XNO, ASSET_BAL_NEW from " & FATFAMF1 & " WHERE " & FATFAMF1 & ".ASSET_ACTION in ('D','W')"
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & $" Declare Cursor C1 is Select * from {FATFAMF1} WHERE {FATFAMF1}.ASSET_ACTION in ('D','W');" & vbCrLf _
            & $" Begin" & vbCrLf _
            & $"  For R1 in C1 Loop" & vbCrLf _
            & $"   Update FATFAMF1 Set" & vbCrLf _
            & $"     ASSET_DEP = NVL(ASSET_DEP,0) + NVL(R1.ASSET_DEP_EXP,0) + NVL(R1.ASSET_DEP_WOF,0)" & vbCrLf _
            & $"   , ASSET_BAL = NVL(ASSET_BAL,0) - NVL(R1.ASSET_DEP_EXP,0) - NVL(R1.ASSET_DEP_WOF,0)" & vbCrLf _
            & $"   , OPS_YYYYPP_LAST_DEPR = '{RYP}'" & vbCrLf _
            & $"    where ASSET_NO = R1.ASSET_NO;" & vbCrLf _
            & $"   Update FATFAMF1 Set ASSET_STATUS = 'D', ASSET_ACTION = 'N' where ASSET_NO = R1.ASSET_NO and ASSET_BAL = 0;" & vbCrLf _
            & $"   If NVL(R1.IN_SERVICE_MOS_ADJ,0) > 0 Then" & vbCrLf _
            & $"     Update FATFAMF1 Set ASSET_DATE_IN_SERVICE = ASSET_DATE_IN_SERVICE_NEW, OPS_YYYYPP_IN_SERVICE = OPS_YYYYPP_IN_SERVICE_NEW where ASSET_NO = R1.ASSET_NO;" & vbCrLf _
            & $"     Update FATFAMF1 Set ASSET_DATE_IN_SERVICE_NEW = NULL, OPS_YYYYPP_IN_SERVICE_NEW = NULL, IN_SERVICE_MOS_ADJ = 0 where ASSET_NO = R1.ASSET_NO;" & vbCrLf _
            & $"   End If;" & vbCrLf _
            & $"  End Loop;" & vbCrLf _
            & $" End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL(sql)

        If ROWs("FATPARM1").Item("FA_PARM_ASSET_DISPOSAL") <> "N" And chkAssetDisposal.Checked Then
            ASCMAIN1.sql = "" _
             & $"   Update FATFAMF1 Set ASSET_STATUS = 'X', OPS_YYYYPP_DISPOSED = '{RYP}' where ASSET_NO in (Select ASSET_NO from {FATFAMF1} where ASSET_ACTION = 'Z')"
            ASCDATA1.ExecuteSQL(sql)
        End If

        ASCMAIN1.sql = $"Update FATPARM1 Set FA_PARM_DEPR_LAST_YP_UPDATED = '{RYP}' where FA_PARM_KEY = 'Z'"
        ASCDATA1.ExecuteSQL(sql)

        ' Because this accrual is produced after period-end, we need to insert some rows to GLTCREC3
        ' This block was taken from TARPEND1 and modified to work from the accrual work tables


        ASCMAIN1.sql = "Delete from GLTCREC1 where CREC_TYPE_CODE = 'FA'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from GLTCREC2 where CREC_TYPE_CODE = 'FA'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into GLTCREC1 Select 'FA', 'Fixed Assets', 'D' from DUAL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into GLTCREC2 Select 'FA', ASSET_CLASS_CODE, ASSET_CLASS_DESC from FATFACL1"
        ASCDATA1.ExecuteSQL()

        Dim cols As String = " (OPS_YYYYPP,CREC_TYPE_CODE,CREC_CLASS_CODE,ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,DETL_CTL_TYPE,DETL_CTL_NO,DETL_CVX_TYPE,DETL_CVX_NO,CREC_AMT) "

        ASCMAIN1.sql = "Select '" & RYP & "' OPS_YYYYPP, 'FA' CREC_TYPE_CODE, FATFAMF1.ASSET_CLASS_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", FATFACL1.ACCT_CODE_CAP ACCT_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            & ", 'X' DETL_CTL_TYPE, FATFADP1.JOURNAL_XNO DETL_CTL_NO" & vbCrLf _
            & ", 'F' DETL_CVX_TYPE, FATFAMF1.ASSET_NO DETL_CVX_NO" & vbCrLf _
            & ", NVL(FATFADP1.ASSET_BAL_NEW,0) CREC_AMT" & vbCrLf _
            & " from FATFAMF1,FATFACL1,FATFADP1" & vbCrLf _
            & " where FATFADP1.JOURNAL_XNO = '" & XNO & "'" & vbCrLf _
            & "   and FATFAMF1.ASSET_NO = FATFADP1.ASSET_NO" & vbCrLf _
            & "   and FATFACL1.ASSET_CLASS_CODE = FATFAMF1.ASSET_CLASS_CODE"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        GL_Update()
        'Else
        '    MsgBox("Note - GL Update Disabled until we are Live")
        'End If


    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables()

        With dst

            ASCMAIN1.sql = "Select FATFAMF1.*" & vbCrLf _
                & $" from {FATFAMF1} FATFAMF1"
            Create_TDA(.Tables.Add, "FATFAMF1", "**", 0, False)
            '.Tables("FATFAMF1").Columns.Add("ASSET_BAL_NEW", GetType(System.Decimal), "ISNULL(ASSET_BAL,0) - ISNULL(ASSET_DEP_EXP,0) - ISNULL(ASSET_DEP_WOF,0)")

            ASCMAIN1.sql = "Select FATFATR1.*" & vbCrLf _
                & $" from {FATFATR1} FATFATR1"
            Create_TDA(.Tables.Add, "FATFATR1", "**", 0, False)

            ASCMAIN1.sql = "Select GLTPARM2.* from GLTPARM2 where OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)

            For Each TABLE_NAME As String In New String() {"FATFACL1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 1)
            Next

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() {"FATFACL1"}
            Fill_Records(TABLE_NAME)
        Next

        If perform_fill Then
            Fill_Records_RPT(sqlw)
        End If

        Return clsASCBASE1
    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        If parms.Length > 0 Then
            sqlw = parms(0)
        End If

        Fill_Records("GLTPARM2", RYP)

        ASCMAIN1.sql = $"Delete from {FATFAMF1}"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = $"Delete from {FATFATR1}"
        ASCDATA1.ExecuteSQL()

        If chkAssetDisposal.Checked Then
            Dim RYPD As String = cmbYP_Dispose_Thru.Value
            RYP_DISPOSE_THRU = Mid(RYPD, 1, 4) & Mid(RYPD, 6, 2)
        Else
            RYP_DISPOSE_THRU = ""
        End If


        ASCMAIN1.sql = $"Select Count (*) from FATFADP1 where OPS_YYYYPP = '{RYP}'"
        reprint = (Val(ASCDATA1.GetDataValue) > 0)

        If reprint Then
            ASCMAIN1.sql = $"Insert into {FATFAMF1} Select FATFAMF1.*, FATFADP1.ASSET_DEP_EXP, FATFADP1.ASSET_DEP_WOF, NVL(FATFAMF1.ASSET_BAL,0) - NVL(FATFADP1.ASSET_DEP_EXP,0) - NVL(FATFADP1.ASSET_DEP_WOF,0) ASSET_BAL_NEW  from FATFAMF1,FATFADP1 where FATFADP1.OPS_YYYYPP = '{RYP}' and FATFAMF1.ASSET_NO = FATFADP1.ASSET_NO"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & $" Declare Cursor C1 Is Select * from FATFADP1 where OPS_YYYYPP = '{RYP}';" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & $"   Update {FATFAMF1} Set ASSET_AMT = R1.ASSET_AMT, ASSET_DEP = R1.ASSET_DEP, ASSET_WOF = R1.ASSET_WOF, ASSET_BAL = R1.ASSET_BAL where ASSET_NO = R1.ASSET_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = $"Insert into {FATFATR1} Select * from FATFATR1 where OPS_YYYYPP = '{RYP}'"
            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = $"Insert into {FATFAMF1} {sqlFATFAMF1} where X.OPS_YYYYPP_IN_SERVICE <= '{RYP}'"
            ASCDATA1.ExecuteSQL()

            If chkAssetDisposal.Checked Then
                Dim sqlFATFAMF1_Disposed As String = sqlFATFAMF1
                sqlFATFAMF1_Disposed = Replace(sqlFATFAMF1_Disposed, "where FATFAMF1.ASSET_ACTION in ('D','W')", "where FATFAMF1.ASSET_ACTION = 'N'")
                ASCMAIN1.sql = $"Insert into {FATFAMF1} {sqlFATFAMF1_Disposed} where ASSET_BAL = 0 and ASSET_STATUS = 'D' and OPS_YYYYPP_DISPOSED is Null and OPS_YYYYPP_LAST_DEPR <= '{RYP_DISPOSE_THRU}'"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = $"Update {FATFAMF1} Set ASSET_ACTION = 'Z', MOS = NULL where ASSET_STATUS = 'D' and ASSET_ACTION = 'N'"
                ASCDATA1.ExecuteSQL()
            End If

            ASCMAIN1.sql = sqlFATFATR1
            ASCDATA1.ExecuteSQL()
        End If

        If reprint Or InquiryMode Then '  RYP = ASCMAIN1.CYP Then
            RWU = "N"
        End If

        EnforceConstraints(False)
        Fill_Records("FATFAMF1")
        Fill_Records("FATFATR1")
        EnforceConstraints(True)

        Dim JOURNAL_TYPE_FADE As String = "FADE"
        Dim JOURNAL_TYPE_FAAD As String = "FAAD"
        If ROWs("FATPARM1").Item("FA_PARM_ASSET_DISPOSAL") = "S" Then
            JOURNAL_TYPE_FAAD = JOURNAL_TYPE_FADE
        End If

        Prepare_GL_Interface(JOURNAL_TYPE_FADE, JOURNAL_TYPE_FAAD)

        ASCMAIN1.sql = $"Update {FATFATR1} Set JOURNAL_NO = '{JOURNAL_NO}'"
        ASCDATA1.ExecuteSQL()

    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE_FADE As String, ByVal JOURNAL_TYPE_FAAD As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO_FADE As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim JOURNAL_NO_FAAD As String = JOURNAL_NO_FADE
        If JOURNAL_TYPE_FAAD <> JOURNAL_TYPE_FADE Then
            JOURNAL_NO_FAAD = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        End If


        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP)
        Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        ' Expense based on Asset Class

        Dim TOTAL_DEPR_EXP As Decimal = 0
        Dim rowGLTINTF1 As DataRow = Nothing

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("FATFAMF1"), New String() {"ASSET_CLASS_CODE"}).Select

            Dim ASSET_CLASS_CODE As String = row.Item("ASSET_CLASS_CODE")
            Dim rowFATFACL1 As DataRow = LookUp("FATFACL1", ASSET_CLASS_CODE)

            Dim ASSET_DEP_EXP As Decimal = Val(dst.Tables("FATFAMF1").Compute("SUM(ASSET_DEP_EXP)", $"ASSET_CLASS_CODE = '{ASSET_CLASS_CODE}'") & "")
            Dim ASSET_DEP_WOF As Decimal = Val(dst.Tables("FATFAMF1").Compute("SUM(ASSET_DEP_WOF)", $"ASSET_CLASS_CODE = '{ASSET_CLASS_CODE}'") & "")
            Dim DETL_POSTING_AMT As Decimal = 0 ' ASSET_DEP_EXP + ASSET_DEP_WOF
            Dim DETL_DESC As String = ""

            DETL_DESC = "Depreciation Expense"
            DETL_POSTING_AMT = ASSET_DEP_EXP
            If DETL_POSTING_AMT <> 0 Then
                TOTAL_DEPR_EXP += DETL_POSTING_AMT

                rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE_FADE, JOURNAL_NO_FADE, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("ACCT_CODE") = rowFATFACL1.Item("ACCT_CODE_EXP")
                rowGLTINTF1.Item("DETL_CVX_NO") = ASSET_CLASS_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "F"
                rowGLTINTF1.Item("DETL_DESC") = DETL_DESC
            End If

            DETL_DESC = "Write-Off"
            DETL_POSTING_AMT = ASSET_DEP_WOF
            If DETL_POSTING_AMT <> 0 Then
                TOTAL_DEPR_EXP += DETL_POSTING_AMT

                rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE_FADE, JOURNAL_NO_FADE, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("ACCT_CODE") = rowFATFACL1.Item("ACCT_CODE_EXP")
                rowGLTINTF1.Item("DETL_CVX_NO") = ASSET_CLASS_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "F"
                rowGLTINTF1.Item("DETL_DESC") = DETL_DESC
            End If

            DETL_DESC = ""
            DETL_POSTING_AMT = ASSET_DEP_EXP + ASSET_DEP_WOF
            If DETL_POSTING_AMT <> 0 Then
                rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE_FADE, JOURNAL_NO_FADE, JOURNAL_LNO, DETL_CTL_DATE, -1 * DETL_POSTING_AMT)
                rowGLTINTF1.Item("ACCT_CODE") = rowFATFACL1.Item("ACCT_CODE_DEP")
                rowGLTINTF1.Item("DETL_CVX_NO") = ASSET_CLASS_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "F"
            End If

        Next


        If ROWs("FATPARM1").Item("FA_PARM_ASSET_DISPOSAL") <> "N" And chkAssetDisposal.Checked Then

            Dim TOTAL_ASSET_AMT_DISPOSE As Decimal = 0

            For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("FATFAMF1"), New String() {"ASSET_CLASS_CODE"}).Select

                Dim ASSET_CLASS_CODE As String = row.Item("ASSET_CLASS_CODE")
                Dim rowFATFACL1 As DataRow = LookUp("FATFACL1", ASSET_CLASS_CODE)

                Dim sqlw As String = $"ASSET_CLASS_CODE = '{ASSET_CLASS_CODE}' and ASSET_ACTION = 'Z'"
                Dim ASSET_AMT_DISPOSE As Decimal = Val(dst.Tables("FATFAMF1").Compute("SUM(ASSET_AMT)", sqlw) & "")

                Dim DETL_POSTING_AMT As Decimal = ASSET_AMT_DISPOSE

                If DETL_POSTING_AMT <> 0 Then
                    TOTAL_ASSET_AMT_DISPOSE += DETL_POSTING_AMT
                    rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE_FAAD, JOURNAL_NO_FAAD, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                    rowGLTINTF1.Item("ACCT_CODE") = rowFATFACL1.Item("ACCT_CODE_DEP")
                    rowGLTINTF1.Item("DETL_CVX_NO") = ASSET_CLASS_CODE
                    rowGLTINTF1.Item("DETL_CVX_TYPE") = "F"
                    rowGLTINTF1.Item("DETL_DESC") = "Asset Disposal"

                    rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE_FAAD, JOURNAL_NO_FAAD, JOURNAL_LNO, DETL_CTL_DATE, -1 * DETL_POSTING_AMT)
                    rowGLTINTF1.Item("ACCT_CODE") = rowFATFACL1.Item("ACCT_CODE_CAP")
                    rowGLTINTF1.Item("DETL_CVX_NO") = ASSET_CLASS_CODE
                    rowGLTINTF1.Item("DETL_CVX_TYPE") = "F"
                    rowGLTINTF1.Item("DETL_DESC") = "Asset Disposal"
                End If
            Next
        End If


        Return JOURNAL_NO
    End Function

    Function Write_GLTINTF1(JOURNAL_TYPE As String, JOURNAL_NO As String, ByRef JOURNAL_LNO As Integer, DETL_CTL_DATE As Date, DETL_POSTING_AMT As Decimal)
        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = RYP
        rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
        JOURNAL_LNO += 1
        rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
        rowGLTINTF1("ACCT_CODE") = ""
        rowGLTINTF1("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
        rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
        rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
        rowGLTINTF1("DETL_DESC") = DBNull.Value
        rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
        ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Return rowGLTINTF1
    End Function

    Sub Create_Work_Tables()

        Dim MOS As String = $"PERIOD_DIFF(SUBSTR(NVL(OPS_YYYYPP_LAST_DEPR,PERIOD_CALC(OPS_YYYYPP_IN_SERVICE,-1)),1,6),'{RYP}')"

        Dim DEPR_POS As String = $"Least (FATFAMF1.ASSET_BAL, Round ({MOS} * (1 + NVL(FATFAMF1.IN_SERVICE_MOS_ADJ,0)) * FATFAMF1.ASSET_AMT / FATFAMF1.ASSET_LIFE_MOS, 2))"
        Dim DEPR_NEG As String = $"Greatest (FATFAMF1.ASSET_BAL, Round ({MOS} * (1 + NVL(FATFAMF1.IN_SERVICE_MOS_ADJ,0)) * FATFAMF1.ASSET_AMT / FATFAMF1.ASSET_LIFE_MOS, 2))"
        '& $", Decode (FATFAMF1.ASSET_ACTION, 'D', Least (FATFAMF1.ASSET_BAL, Round ({MOS} * FATFAMF1.ASSET_AMT / FATFAMF1.ASSET_LIFE_MOS, 2)),0) ASSET_DEP_EXP" & vbCrLf _

        ASCMAIN1.sql = "Select FATFAMF1.*" & vbCrLf _
            & $", Decode (FATFAMF1.ASSET_ACTION, 'D', Case When ASSET_BAL >=0 THEN {DEPR_POS} ELSE {DEPR_NEG} END,0) ASSET_DEP_EXP" & vbCrLf _
            & ", Decode (FATFAMF1.ASSET_ACTION, 'W', FATFAMF1.ASSET_BAL, 0) ASSET_DEP_WOF" & vbCrLf _
            & $", {MOS} MOS" & vbCrLf _
            & $" from FATFAMF1 where FATFAMF1.ASSET_STATUS in ('D','C')" ' where FATFAMF1.ASSET_ACTION in ('D','W')"

        '            & $", CASE WHEN ASSET_ACTION = 'Z' THEN NULL ELSE {MOS} END MOS" & vbCrLf _
        ', Decode (FATFAMF1.ASSET_ACTION, 'D', Least (FATFAMF1.ASSET_BAL, Round (FATFAMF1.ASSET_AMT / FATFAMF1.ASSET_LIFE_MOS, 2)),0) ASSET_DEP_EXP
        ASCMAIN1.sql = $"Select X.*, NVL(X.ASSET_BAL,0) - NVL(X.ASSET_DEP_EXP,0)  - NVL(X.ASSET_DEP_WOF,0) ASSET_BAL_NEW from ({ASCMAIN1.sql}) X"
        sqlFATFAMF1 = ASCMAIN1.sql
        ASCMAIN1.sql &= " where ROWNUM < 1"
        FATFAMF1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL($"Alter table {FATFAMF1} Add Primary Key (ASSET_NO)")

        ASCMAIN1.sql = "Select * from FATFATR1 where ROWNUM < 1"
        FATFATR1 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = $"Insert into {FATFATR1}" & vbCrLf _
            & "Select TAPCTLN1('FATFATR1.ASSET_TRN_NO',1) ASSET_TRN_NO" & vbCrLf _
            & ", TRUNC(SYSDATE) ASSET_TRN_DATE" & vbCrLf _
            & ", 'EXP' ASSET_TRN_TYPE" & vbCrLf _
            & ", FATFAMF1.ASSET_DEP_EXP ASSET_TRN_AMT" & vbCrLf _
            & ", FATFAMF1.ASSET_NO" & vbCrLf _
            & ", NULL VOUCHER_NO" & vbCrLf _
            & ", NULL VOUCHER_LNO" & vbCrLf _
            & $", '{JOURNAL_NO}' JOURNAL_NO" & vbCrLf _
            & ", NULL JOURNAL_LNO" & vbCrLf _
            & ", SYSDATE INIT_DATE" & vbCrLf _
            & $", '{ASCMAIN1.USER_ID}' INIT_OPER" & vbCrLf _
            & ", NULL LAST_DATE" & vbCrLf _
            & ", NULL LAST_OPER" & vbCrLf _
            & $", '{RYP}' OPS_YYYYPP" & vbCrLf _
            & $" from {FATFAMF1} FATFAMF1 where FATFAMF1.ASSET_ACTION in ('D','W')"
        '& $" from {FATFAMF1} FATFAMF1 where FATFAMF1.ASSET_STATUS in ('D','C')" ' where FATFAMF1.ASSET_ACTION in ('D','W')"

        sqlFATFATR1 = ASCMAIN1.sql
        'ASCDATA1.ExecuteSQL()

    End Sub


    Sub Prepare_Data_Extracts()

        Dim tbl As DataTable

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        'tbl = dst.Tables("FATFAMF1").Select("ASSET_ACTION = 'D' or ASSET_ACTION = 'W'").CopyToDataTable
        'grdASTEXPT1.DataSource = tbl ' dst.Tables("FATFAMF1")

        grdASTEXPT1.DataSource = dst.Tables("FATFAMF1")
        Dim dvw As DataView = DirectCast(grdASTEXPT1.DataSource, DataTable).DefaultView
        '
        'dvw.RowFilter = "ASSET_ACTION = 'D' or ASSET_ACTION = 'W'"


        grdASTEXPT1.Text = "Fixed Assets Depreciation Expense - " & Mid(RYPLEGEND, 10, 6)
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        Set_DX_Column(grdASTEXPT1, "ASSET_NO", "Asset", 80)
        Set_DX_Column(grdASTEXPT1, "ASSET_DESC", "Description", 120)
        Set_DX_Column(grdASTEXPT1, "ASSET_CLASS_CODE", "Class", 80)
        Set_DX_Column(grdASTEXPT1, "ASSET_DEPR_CODE", "Method", 80)
        Set_DX_Column(grdASTEXPT1, "ASSET_DATE", "Date", 100, "MM/dd/yyyy")
        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP", "YP Cap", 80)
        Set_DX_Column(grdASTEXPT1, "VEND_CODE", "Vendor", 80)
        Set_DX_Column(grdASTEXPT1, "VEND_NAME", "Vendor Name", 150)
        Set_DX_Column(grdASTEXPT1, "VOUCHER_NO", "Voucher No", 100)
        Set_DX_Column(grdASTEXPT1, "INVOICE_NOTES", "Notes", 120)
        Set_DX_Column(grdASTEXPT1, "ASSET_LIFE_MOS", "#Mos", 60, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "ASSET_DATE_IN_SERVICE", "Date in Svc", 100, "MM/dd/yyyy")
        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP_IN_SERVICE", "YP in Svc", 80)
        Set_DX_Column(grdASTEXPT1, "ASSET_ACTION", "Action", 60)

        Set_DX_Column(grdASTEXPT1, "ASSET_AMT", "Asset Amt", 100, "##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "ASSET_DEP", "Accum Dep", 100, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "ASSET_WOF", "Write-Off", 100, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "ASSET_BAL", "Balance", 100, "#,##0.00", , Color.LightGreen)

        Set_DX_Column(grdASTEXPT1, "ASSET_DEP_EXP", "Dep Exp", 100, "#,##0.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "MOS", "Months", 70, "#,##0", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ASSET_DEP_WOF", "W/O Exp", 100, "#,##0.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "ASSET_BAL_NEW", "New Balance", 100, "#,##0.00", , Color.Orange)

        Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP_LAST_DEPR", "YP Last Dep", 100, ,, Color.Gold)
        'Set_DX_Column(grdASTEXPT1, "OPS_YYYYPP_DISPOSED", "YP Disposed", 100, , , Color.Gold)

        grdASTEXPT1.DisplayLayout.Bands(0).Columns("ASSET_NO").Header.Fixed = True
        grdASTEXPT1.DisplayLayout.Bands(0).Columns("ASSET_DESC").Header.Fixed = True
        grdASTEXPT1.DisplayLayout.Bands(0).Columns("ASSET_CLASS_CODE").Header.Fixed = True

        Create_Summary(grdASTEXPT1, "ASSET_NO", "Count")
        For Each C As String In New String() _
            {"ASSET_AMT", "ASSET_DEP", "ASSET_WOF", "ASSET_BAL", "ASSET_DEP_EXP", "ASSET_DEP_WOF", "ASSET_BAL_NEW"}
            Create_Summary(grdASTEXPT1, C)
        Next

        Sort_grdColumns(grdASTEXPT1, "ASSET_NO")


        If chkAssetDisposal.Checked Then



            Dim WorkbookView1 As SpreadsheetGear.Windows.Forms.WorkbookView

            Dim tabCaption As String = "Asset Disposal"
            Dim tabpage_exists = False
            If tabDataExports.Tabs.Count > 1 AndAlso tabDataExports.Tabs(1).Key = tabCaption Then
                tabpage_exists = True
            End If
            If tabpage_exists Then
                WorkbookView1 = DirectCast(tabDataExports.Tabs(tabCaption).TabPage.Controls(0), SpreadsheetGear.Windows.Forms.WorkbookView)

            Else

                tabDataExports.Tabs.Add(tabCaption)
                tabDataExports.Tabs(tabCaption).Text = tabDataExports.Tabs(tabCaption).Key
                WorkbookView1 = New SpreadsheetGear.Windows.Forms.WorkbookView
                WorkbookView1.Name = "WorkbookView1"
                WorkbookView1.Parent = tabDataExports.Tabs(tabCaption).TabPage
                WorkbookView1.Dock = DockStyle.Fill

                Dim btn As New Infragistics.Win.Misc.UltraButton
                btn.Parent = WorkbookView1 ' tabDataExports.Tabs(tabCaption).TabPage
                btn.Text = ""
                btn.Appearance.Image = System.Drawing.Image.FromFile(ASCMAIN1.Folders("Images") & "32\" & "Excel" & ".png")
                btn.Visible = True
                btn.Top = 0
                btn.Left = 0
                btn.Width = 25
                btn.Height = 25
                AddHandler btn.Click, AddressOf btnExcel_Click
            End If


            WorkbookView1.GetLock()
            WorkbookView1.ActiveWorksheet.Cells.Clear()
            Dim tbl2 As DataTable = dst.Tables("FATFAMF1").Select("ASSET_ACTION = 'Z'").CopyToDataTable
            Load_DataTable_into_SGXLS(1, 1, tbl2, WorkbookView1.ActiveWorksheet, grdASTEXPT1, Nothing, "", "")
            'Load_DataTable_into_SGXLS(1, 1, dst.Tables("SOTORDRX"), WorkbookView1.ActiveWorksheet, grdASTEXPT1, Nothing, "", "")
            WorkbookView1.ReleaseLock()
        End If

    End Sub


    Private Sub btnExcel_Click(sender As Object, e As EventArgs)
        Dim WorkbookView1 As SpreadsheetGear.Windows.Forms.WorkbookView
        WorkbookView1 = DirectCast(tabDataExports.Tabs("Open Orders SSG").TabPage.Controls(0), SpreadsheetGear.Windows.Forms.WorkbookView)
        WorkbookView1.GetLock()
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & Me.Name & "_" & ASCMAIN1.Next_Control_No($"{Me.Name}.XLSX_NO") & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        WorkbookView1.ReleaseLock()
    End Sub


    Private Sub chkAssetDisposal_CheckedChanged(sender As Object, e As EventArgs) Handles chkAssetDisposal.CheckedChanged
        cmbYP_Dispose_Thru.Enabled = chkAssetDisposal.Checked
    End Sub
End Class