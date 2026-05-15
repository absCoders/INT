Imports System.Drawing

Public Class SPRDCOMA

#Region "General Declarations"
    Dim SPTDCOMA As String ' Drives Report
    Dim SPTDCOMB As String ' Detailed to Store
    Dim sqlSPTDCOMA As String
    Dim sqlSPTDCOMB As String
    Dim sqlSPTDCOMC As String

    Dim SPTDCOM1 As String ' Detailed to Store

    Dim reprint As Boolean

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        'Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP("RYP", ASCMAIN1.CYP, -1, -1, -1)

        'Dim YWs As New List(Of String)
        'ASCMAIN1.sql = "Select YYYYWW from GLTPARM1 where YYYYPP = '" & ASCMAIN1.CYP & "'"
        'For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
        '    YWs.Add(row.Item("YYYYWW"))
        'Next

        Set_cmbYW("RYW0", ASCMAIN1.Week_Calc(ROWs("SPTPARM1").Item("SP_PARM_DEMO_LAST_YW_UPDATED"), 1), 0, 0, 0)

    End Sub

    Protected Overrides Sub Build_Workfile()
        RWU = "R"
        Dim sqlw As String = ""
        Prepare_dst(True, sqlw)
        Check_if_Empty("SPTDCOMA")
    End Sub

    Public Overrides Sub Print_Report()

        SUBT = ""
        If reprint Then SUBT = "Re-Print"
        If chkShowAllCustomers.Checked Then SUBT = "Showing All Customers with Retail Sales"
        CR_params.Add("SUBT", SUBT)
        Generate_Report(RPT, , SUBT)
        Print_GL()

        'If ASCMAIN1.Running_in_VS Then
        Prepare_Data_Extracts()
        'End If


    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        'If ASCMAIN1.Running_in_VS Then

        '    Stop ' THIS NEXT PART WILL MERGE THE ACCRUALS - BUT NOTE THAT SPTDCOMC WILL CREATE INDIVIDUAL ACCRUALS AND ALSO SHOULD PROB BE MERGED

        '    ASCMAIN1.sql = "" & vbCrLf _
        '        & "Begin Declare Cursor C1 is " & vbCrLf _
        '        & " Select * from " & SPTDCOMB & " where QTY_SOLD <> 0 or AMT_SOLD <> 0 or QTY_EOW <> 0 or AMT_EOW <> 0 or AMT_COMM <> 0;" & vbCrLf _
        '        & "Begin For R1 in C1 Loop" & vbCrLf _
        '        & "Update SPTDCOMB" & vbCrLf _
        '        & " Set QTY_SOLD = NVL(QTY_SOLD,0) + R1.QTY_SOLD, AMT_SOLD = NVL(AMT_SOLD,0) + R1.AMT_SOLD" & vbCrLf _
        '        & "   , QTY_EOW = NVL(QTY_EOW,0) + R1.QTY_EOW, AMT_EOW = NVL(AMT_EOW,0) + R1.AMT_EOW" & vbCrLf _
        '        & "   , OPS_YYYYWW_MIN = LEAST(OPS_YYYYWW_MIN,R1.OPS_YYYYWW_MIN)" & vbCrLf _
        '        & "   , OPS_YYYYWW_MAX = GREATEST(OPS_YYYYWW_MAX,R1.OPS_YYYYWW_MAX)" & vbCrLf _
        '        & "   , AMT_COMM = NVL(AMT_COMM,0) + R1.AMT_COMM" & vbCrLf _
        '        & "   , DEMO_COMM_PCT = R1.DEMO_COMM_PCT" & vbCrLf _
        '        & "  where OPS_YYYYPP = R1.OPS_YYYYPP and CUST_CODE = R1.CUST_CODE " & vbCrLf _
        '        & "    and CUST_STORE_NO = R1.CUST_STORE_NO and COLLECTION_CODE = R1.COLLECTION_CODE;" & vbCrLf _
        '        & "If SQL%NOTFOUND then" & vbCrLf _
        '        & "Insert into SPTDCOMB " & vbCrLf _
        '        & " Select * from ASW04537 " & vbCrLf _
        '        & "  where OPS_YYYYPP = R1.OPS_YYYYPP and CUST_CODE = R1.CUST_CODE " & vbCrLf _
        '        & "    and CUST_STORE_NO = R1.CUST_STORE_NO and COLLECTION_CODE = R1.COLLECTION_CODE;" & vbCrLf _
        '        & "End If;" & vbCrLf _
        '        & "End Loop; End; End;"
        '    ASCDATA1.ExecuteSQL()

        '    Stop ' NO NEED TO DO THE INSERT BELOW
        'End If

        ASCMAIN1.sql = "Insert into SPTDCOMB Select * from " & SPTDCOMB
        ASCDATA1.ExecuteSQL(sql)


        For Each row As DataRow In dst.Tables("SPTDCOMC").Select
            row.Item("ACC_CTL_NO") = ASCMAIN1.Next_Control_No("SPTDCOMC.ACC_CTL_NO")
            row.Item("JOURNAL_XNO") = XNO
            row.Item("JOURNAL_IND") = "1"
            row.AcceptChanges()
            row.SetAdded()
        Next
        Update_Record_TDA("SPTDCOMC")

        'If RYP < "201601" Then
        '    MsgBox("Skipping GL Update - INT")
        '    Exit Sub
        'End If

        ' Because this accrual is produced after period-end, we need to insert some rows to GLTCREC3
        ' This block was taken from TARPEND1 and modified to work from the accrual work tables

        Dim cols As String = " (OPS_YYYYPP,CREC_TYPE_CODE,CREC_CLASS_CODE,ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE,DETL_CTL_TYPE,DETL_CTL_NO,DETL_CVX_TYPE,DETL_CVX_NO,CREC_AMT) "

        ASCMAIN1.sql = "Select '" & RYP & "' OPS_YYYYPP, 'DC' CREC_TYPE_CODE, SPTDCOMC.BRAND_CODE CREC_CLASS_CODE" & vbCrLf _
            & ", SPTPARM1.SP_PARM_DEMO_ACCT_CODE_ACC ACCT_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
            & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
            & ", 'D' DETL_CTL_TYPE, SPTDCOMC.ACC_CTL_NO DETL_CTL_NO" & vbCrLf _
            & ", 'C' DETL_CVX_TYPE, SPTDCOMC.CUST_CODE DETL_CVX_NO" & vbCrLf _
            & ", NVL(SPTDCOMC.AMT_COMM,0)-NVL(SPTDCOMC.AMT_COMM_OFFSET,0) CREC_AMT" & vbCrLf _
            & " from SPTDCOMC,SPTPARM1" & vbCrLf _
            & " where SPTDCOMC.JOURNAL_XNO = '" & XNO & "'" & vbCrLf _
            & "   and SPTPARM1.SP_PARM_KEY = 'Z'"
        ASCMAIN1.sql = "Insert into GLTCREC3 " & cols & ASCMAIN1.sql
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update SPTPARM1 set SP_PARM_DEMO_LAST_YW_UPDATED = '" & RYW1 & "'"
        ASCDATA1.ExecuteSQL()

        If ASCMAIN1.CLIENT = "INT" Then

            ' Demo Commissions
            ASCMAIN1.Progress("Demo Commissions Snapshots", "")


            'If ASCMAIN1.Running_in_VS Then
            '    Stop ' THESE NEXT FEW LINES WILL REPLACE HISTORY
            '    ASCMAIN1.sql = "Delete from SPTDCOMH where OPS_YYYYPP_H = '" & RYP & "'"
            '    ASCDATA1.ExecuteSQL()
            '    ASCMAIN1.sql = "Delete from SPTDCOMI where OPS_YYYYPP_H = '" & RYP & "'"
            '    ASCDATA1.ExecuteSQL()
            'End If

            ASCMAIN1.sql = "Insert into SPTDCOMH Select '" & RYP & "' OPS_YYYYPP_H, SPTDCOMB.* from SPTDCOMB where OPS_YYYYPP = '" & RYP & "' and NVL(SPTDCOMB.AMT_COMM,0) - NVL(SPTDCOMB.AMT_COMM_PAID,0) <> 0"
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into SPTDCOMI Select '" & RYP & "' OPS_YYYYPP_H, SPTDCOMC.* from SPTDCOMC where OPS_YYYYPP = '" & RYP & "'"
            ASCDATA1.ExecuteSQL()
        End If

        GL_Update()

    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SPTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables()

        With dst

            ASCMAIN1.sql = "Select SPTDCOMA.*, ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & " from " & SPTDCOMA & " SPTDCOMA, ARTCUST1 where ARTCUST1.CUST_CODE = SPTDCOMA.CUST_CODE"
            Create_TDA(.Tables.Add, "SPTDCOMA", "**", 0, False)

            ASCMAIN1.sql = "Select SPTDCOMB.* from " & SPTDCOMB & " SPTDCOMB"
            Create_TDA(.Tables.Add, "SPTDCOMB", "**", 0, False)

            Create_TDA(.Tables.Add, "SPTDCOMC", "*")

            ASCMAIN1.sql = "Select SPTDCOM1.* from " & SPTDCOM1 & " SPTDCOM1"
            Create_TDA(.Tables.Add, "SPTDCOM1", "**", 0)


            ASCMAIN1.sql = "Select GLTPARM3.* from GLTPARM3 where YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "GLTPARM3", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select GLTPARM2.* from GLTPARM2 where OPS_YYYYPP = :PARM1"
            Create_TDA(.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)

            For Each TABLE_NAME As String In New String() _
                {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1"}
                ASCMAIN1.sql = "Select * from " & TABLE_NAME
                If TABLE_NAME = "ARTCUST1" Then ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, TRADE_CLASS_CODE from " & TABLE_NAME
                Create_TDA(.Tables.Add, TABLE_NAME, "**", 0, False, "", 1)
            Next

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"ICTCOLL1", "ICTBRAN1", "ARTCUST1", "ICTCOLL0", "SOTTCLS1"}
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

        Fill_Records("GLTPARM3", RYP)
        Fill_Records("GLTPARM2", RYP)
        RYW = dst.Tables("GLTPARM3").Compute("MAX(YYYYWW)", "")

        ASCMAIN1.sql = "Delete from " & SPTDCOM1
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into " & SPTDCOM1 & " Select * from SPTDCOM1"
        ASCDATA1.ExecuteSQL()

        If chkShowAllCustomers.Checked Then
            ASCMAIN1.sql = "Insert into " & SPTDCOM1 & " (CUST_CODE) " _
                & " Select Distinct CUST_CODE from RSTRETL1 where OPS_YYYYPP = '" & RYP & "'" _
                & "  minus Select CUST_CODE from SPTDCOM1"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Select Count (*) from SPTDCOMC where OPS_YYYYPP = '" & RYP & "'"
        reprint = (Val(ASCDATA1.GetDataValue) > 0)

        If chkShowAllCustomers.Checked Then
            RWU = "N"
            reprint = False
        End If

        If RYP = ASCMAIN1.CYP Then
            RWU = "N"
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTDCOMB)
        If reprint Then
            RWU = "N"
            ASCDATA1.ExecuteSQL("Insert into " & SPTDCOMB & " Select * from SPTDCOMB where OPS_YYYYPP = '" & RYP & "'")
        Else
            '    If ASCMAIN1.Running_in_VS Then
            '        Stop
            '        sqlSPTDCOMB = Replace(sqlSPTDCOMB, "DEMO_COMM_BASIS = '1'", "DEMO_COMM_BASIS = 'X'") ' NO GROSS SALES
            '        sqlSPTDCOMB = Replace(sqlSPTDCOMB, "and RSTRETL1.OPS_YYYYPP = 'YYYYPP'", "and RSTRETL1.OPS_YYYYWW = '201809'") ' WEEK 09 ONLY
            '    End If
            '    ASCDATA1.ExecuteSQL("Insert into " & SPTDCOMB & " " & Replace(Replace(sqlSPTDCOMB, "'YYYYWW'", "'" & RYW & "'"), "'YYYYPP'", "'" & RYP & "'"))

            ASCDATA1.ExecuteSQL("Insert into " & SPTDCOMB & " " & Replace(Replace(Replace(sqlSPTDCOMB, "'YYYYW0'", "'" & RYW0 & "'"), "'YYYYW1'", "'" & RYW1 & "'"), "'YYYYPP'", "'" & RYP & "'"))
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & SPTDCOMA)
        ASCMAIN1.sql = "Insert into " & SPTDCOMA & " " & sqlSPTDCOMA
        ASCDATA1.ExecuteSQL()

        EnforceConstraints(False)
        Fill_Records("SPTDCOM1")
        Fill_Records("SPTDCOMA")
        Fill_Records("SPTDCOMB")
        Fill_Records("SPTDCOMC", "", True, sqlSPTDCOMC)
        EnforceConstraints(True)

        Prepare_GL_Interface("SPDC")

    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP)
        Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        ' Expense based on Trade Class and Brand

        Dim TOTAL_DEMO_COMM As Decimal = 0

        For Each row As DataRow In
            ASCDATA1.SelectDistinct(dst.Tables("SPTDCOMA"), New String() {"TRADE_CLASS_CODE", "COLLECTION_CODE"}).Select

            Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE")
            Dim COLLECTION_CODE As String = row.Item("COLLECTION_CODE")

            Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)

            Dim DETL_POSTING_AMT As Decimal = Val(dst.Tables("SPTDCOMA").Compute("SUM(AMT_COMM)", "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "' and COLLECTION_CODE = '" & COLLECTION_CODE & "'") & "")

            If DETL_POSTING_AMT <> 0 Then
                TOTAL_DEMO_COMM += DETL_POSTING_AMT
                Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
                rowGLTINTF1.Item("OPS_YYYYPP") = RYP
                rowGLTINTF1.Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_DEMO_ACCT_CODE_EXP")
                If rowSOTTCLS1 IsNot Nothing Then
                    If rowSOTTCLS1.Item("SEG3_CODE") & "" <> "" Then
                        rowGLTINTF1.Item("SEG3_CODE") = rowSOTTCLS1.Item("SEG3_CODE")
                    Else
                        rowGLTINTF1.Item("SEG3_CODE") = TRADE_CLASS_CODE
                    End If
                Else
                    rowGLTINTF1.Item("SEG3_CODE") = "?"
                End If
                If rowICTCOLL1 IsNot Nothing Then
                    If rowICTCOLL1.Item("SEG4_CODE") & "" <> "" Then
                        rowGLTINTF1.Item("SEG4_CODE") = rowICTCOLL1.Item("SEG4_CODE")
                    Else
                        rowGLTINTF1.Item("SEG4_CODE") = COLLECTION_CODE
                    End If
                Else
                    rowGLTINTF1.Item("SEG4_CODE") = "?"
                End If
                If rowGLTINTF1.Item("SEG2_CODE") & "" = "" Then rowGLTINTF1.Item("SEG2_CODE") = "?"
                If rowGLTINTF1.Item("SEG3_CODE") & "" = "" Then rowGLTINTF1.Item("SEG3_CODE") = "?"
                If rowGLTINTF1.Item("SEG4_CODE") & "" = "" Then rowGLTINTF1.Item("SEG4_CODE") = "?"
                rowGLTINTF1.Item("DETL_CVX_NO") = COLLECTION_CODE
                rowGLTINTF1.Item("DETL_CVX_TYPE") = "L"
            End If
        Next

        ' Accrual

        If TOTAL_DEMO_COMM <> 0 Then
            Dim rowGLTINTF1 As DataRow = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, -1 * TOTAL_DEMO_COMM)
            rowGLTINTF1.Item("OPS_YYYYPP") = RYP
            rowGLTINTF1.Item("ACCT_CODE") = ROWs("SPTPARM1").Item("SP_PARM_DEMO_ACCT_CODE_ACC")
        End If

        Return JOURNAL_NO
    End Function

    Function Write_GLTINTF1(JOURNAL_TYPE As String, JOURNAL_NO As String, ByRef JOURNAL_LNO As Integer, DETL_CTL_DATE As Date, DETL_POSTING_AMT As Decimal)
        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = RYP0
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


        ASCMAIN1.sql = "Select * from SPTDCOM1"
        SPTDCOM1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter table " & SPTDCOM1 & " Add Primary Key (CUST_CODE)")

        '& "   and RSTRETL1.OPS_YYYYPP = 'YYYYPP'" & vbCrLf _

        sqlSPTDCOMB = "" _
            & "Select 'YYYYPP' OPS_YYYYPP, X.*, X.AMT_SOLD * " & vbCrLf _
            & "NVL(SPTDCOM4.DEMO_COMM_PCT," & vbCrLf _
            & "NVL(SPTDCOM2.DEMO_COMM_PCT," & vbCrLf _
            & "NVL(SPTDCOM3.DEMO_COMM_PCT,SPTDCOM1.DEMO_COMM_PCT)))/100 AMT_COMM," & vbCrLf _
            & "NVL(SPTDCOM4.DEMO_COMM_PCT," & vbCrLf _
            & "NVL(SPTDCOM2.DEMO_COMM_PCT," & vbCrLf _
            & "NVL(SPTDCOM3.DEMO_COMM_PCT,SPTDCOM1.DEMO_COMM_PCT))) DEMO_COMM_PCT," & vbCrLf _
            & "0 AMT_COMM_CLAIMED," & vbCrLf _
            & "0 AMT_COMM_PAID" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", SUM (RSTRETL1.QTY_SOLD) QTY_SOLD" & vbCrLf _
            & ", SUM (RSTRETL1.AMT_SOLD) AMT_SOLD" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,'YYYYW1',RSTRETL1.QTY_EOW,0)) QTY_EOW" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,'YYYYW1',RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE,0)) AMT_EOW" & vbCrLf _
            & ", MIN (RSTRETL1.OPS_YYYYWW) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (RSTRETL1.OPS_YYYYWW) OPS_YYYYWW_MAX" & vbCrLf _
            & " from RSTRETL1,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = RSTRETL1.ITEM_CODE" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYWW >= 'YYYYW0'" & vbCrLf _
            & "   and RSTRETL1.OPS_YYYYWW <= 'YYYYW1'" & vbCrLf _
            & IIf(chkShowAllCustomers.Checked, "", "   and RSTRETL1.CUST_CODE IN (Select CUST_CODE from SPTDCOM1 where DEMO_COMM_BASIS = '0')" & vbCrLf) _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY_SOLD" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) AMT_SOLD" & vbCrLf _
            & ", NULL QTY_EOW" & vbCrLf _
            & ", NULL AMT_EOW" & vbCrLf _
            & ", MIN (SOTINVH2.OPS_YYYYWW) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (SOTINVH2.OPS_YYYYWW) OPS_YYYYWW_MAX" & vbCrLf _
            & " from SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED = 'YYYYPP'" & vbCrLf _
            & "   and SOTINVH2.CUST_CODE IN (Select CUST_CODE from SPTDCOM1 where DEMO_COMM_BASIS = '1')" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) QTY_SOLD" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ITEM_RETAIL_PRICE) AMT_SOLD" & vbCrLf _
            & ", NULL QTY_EOW" & vbCrLf _
            & ", NULL AMT_EOW" & vbCrLf _
            & ", MIN (SOTINVH2.OPS_YYYYWW) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (SOTINVH2.OPS_YYYYWW) OPS_YYYYWW_MAX" & vbCrLf _
            & " from SOTINVH2,ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED = 'YYYYPP'" & vbCrLf _
            & "   and SOTINVH2.CUST_CODE IN (Select CUST_CODE from SPTDCOM1 where DEMO_COMM_BASIS = '2')" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, " & vbCrLf _
            & "ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ") X, " & SPTDCOM1 & " SPTDCOM1,SPTDCOM2,SPTDCOM3,SPTDCOM4" & vbCrLf _
            & " where SPTDCOM1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and SPTDCOM2.CUST_CODE (+) = X.CUST_CODE " & vbCrLf _
            & "   and SPTDCOM2.CUST_STORE_NO (+) = X.CUST_STORE_NO" & vbCrLf _
            & "   and SPTDCOM3.CUST_CODE (+) = X.CUST_CODE " & vbCrLf _
            & "   and SPTDCOM3.HC_CODE (+) = X.HC_CODE" & vbCrLf _
            & "   and SPTDCOM4.CUST_CODE (+) = X.CUST_CODE " & vbCrLf _
            & "   and SPTDCOM4.CUST_STORE_NO (+) = X.CUST_STORE_NO" & vbCrLf _
            & "   and SPTDCOM4.HC_CODE (+) = X.HC_CODE"

        ASCMAIN1.sql = "Select * from SPTDCOMB where ROWNUM < 1"
        SPTDCOMB = ASCMAIN1.Temp_Table

        sqlSPTDCOMC = "Select TRIM(TO_CHAR(ROWNUM,'000000')) ACC_CTL_NO, X.* from (" _
            & "Select OPS_YYYYPP,CUST_CODE,HC_CODE,BRAND_CODE" & vbCrLf _
            & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
            & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
            & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
            & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
            & ", 0 DAYS_OF_RETAIL" & vbCrLf _
            & ", 0 AMT_COMM_ADJ, 0 AMT_COMM_CLAIMED, 0 AMT_SOLD_CLAIMED, 0 AMT_COMM_PAID" & vbCrLf _
            & ", NULL OPS_YYYYPP_PAID, NULL PYMT_NO" & vbCrLf _
            & " from " & SPTDCOMB & " SPTDCOMB" & vbCrLf _
            & " group by OPS_YYYYPP,CUST_CODE,HC_CODE,BRAND_CODE) X"

        ASCMAIN1.sql = "" _
            & "Select OPS_YYYYPP,CUST_CODE,HC_CODE,BRAND_CODE,COLLECTION_CODE" & vbCrLf _
            & ", SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
            & ", SUM (AMT_SOLD) AMT_SOLD" & vbCrLf _
            & ", SUM (QTY_EOW) QTY_EOW" & vbCrLf _
            & ", SUM (AMT_EOW) AMT_EOW" & vbCrLf _
            & ", SUM (AMT_COMM) AMT_COMM" & vbCrLf _
            & ", MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN" & vbCrLf _
            & ", MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX" & vbCrLf _
            & ", 0 DAYS_OF_RETAIL" & vbCrLf _
            & ", 0 AMT_COMM_ADJ" & vbCrLf _
            & " from " & SPTDCOMB & " SPTDCOMB" & vbCrLf _
            & " group by OPS_YYYYPP,CUST_CODE,HC_CODE,BRAND_CODE,COLLECTION_CODE"

        sqlSPTDCOMA = ASCMAIN1.sql
        SPTDCOMA = ASCMAIN1.Temp_Table


    End Sub

    Private Sub chkShowAllCustomers_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowAllCustomers.CheckedChanged

    End Sub


    Private Sub cmbYP_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles cmbYP.InitializeLayout

    End Sub

    Private Sub cmbYP_ValueChanged(sender As Object, e As EventArgs) Handles cmbYP.ValueChanged
        Dim YP As String = cmbYP.Value
        YP = Mid(YP, 1, 4) & Mid(YP, 6, 2)
        ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 where YYYYPP = '" & YP & "' and REL_WEEK +1 = MAX_WEEK"
        Dim YW_next2last As String = ASCDATA1.GetDataValue

        ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 where YYYYPP = '" & YP & "' and REL_WEEK = MAX_WEEK"
        Dim YW_last As String = ASCDATA1.GetDataValue

        '  cmbYW.Remove("RYW1")
        Set_cmbYW("RYW1", YW_next2last, 0, 1, 1)
        Dim dvw As DataView = DirectCast(Absx1.cmbFor("RYW1").DataSource, DataTable).DefaultView
        dvw.RowFilter = "YYYYWW = '" & YW_last & "' OR YYYYWW = '" & YW_next2last & "'"
        Absx1.cmbFor("RYW1").ActiveRow = Absx1.cmbFor("RYW1").Rows(1)

    End Sub

    Private Sub cmdShowRetailSales_Click(sender As Object, e As EventArgs) Handles cmdShowRetailSales.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Retail Sales by Customer by Week")

        Dim YW0 As String = Absx1.cmbFor("RYW0").Value & ""
        YW0 = Mid(YW0, 1, 4) & Mid(YW0, 6, 2)
        Dim YW1 As String = Absx1.cmbFor("RYW1").Value & ""
        YW1 = Mid(YW1, 1, 4) & Mid(YW1, 6, 2)

        If YW1 > YW0 And YW1 <> "" And YW0 <> "" Then
        Else
            Exit Sub
        End If

        Dim W As Integer = 0
        Dim Ls As New List(Of String)
        Dim sqls As String = ""
        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYWW >= '" & YW0 & "' and YYYYWW <= '" & YW1 & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
            Dim YW As String = row.Item("YYYYWW")
            W += 1
            Ls.Add(row.Item("LEGEND"))
            sqls &= ", SUM (DECODE(OPS_YYYYWW,'" & YW & "',AMT_SOLD,0)) WK" & Format(W, "00") & vbCrLf
        Next

        ASCMAIN1.sql = "" _
            & "Select CUST_CODE, SUM (AMT_SOLD) TOTAL" & vbCrLf _
            & sqls _
            & " from RSTRETL1" & vbCrLf _
            & " where OPS_YYYYWW >= '" & YW0 & "' AND OPS_YYYYWW <= '" & YW1 & "'" & vbCrLf _
            & " group by CUST_CODE"

        Dim tbl As DataTable = ASCDATA1.GetDataTable
        grdRetailSales.DataSource = tbl

        Dim Lss() As String = Ls.ToArray

        grdRetailSales.DisplayLayout.Bands(0).Summaries.Clear()
        If grdRetailSales.DisplayLayout.Bands(0).Summaries.Count = 0 Then
            Create_Summary(grdRetailSales, "CUST_CODE", "Count")
            Create_Summary(grdRetailSales, "TOTAL")
            For wk As Integer = 1 To W
                Dim C As String = "WK" & Format(wk, "00")
                Create_Summary(grdRetailSales, C)
                grdRetailSales.DisplayLayout.Bands(0).Columns(C).Header.Caption = Lss(wk - 1)
            Next
        End If

        grdRetailSales.DisplayLayout.GroupByBox.Hidden = True
        grdRetailSales.Visible = True

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        'grdASTEXPT1.DataSource = dst.Tables("SPTDCOMC")
        grdASTEXPT1.DataSource = dst.Tables("SPTDCOMB")
        grdASTEXPT1.Text = "Promo Expense Accrual - " & Mid(RYPLEGEND, 10, 6)
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        Set_DX_Column(grdASTEXPT1, "CUST_CODE", "Customer", 100)
        Set_DX_Column(grdASTEXPT1, "CUST_STORE_NO", "Store", 80)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collection", 80)
        Set_DX_Column(grdASTEXPT1, "HC_CODE", "HC", 80)
        Set_DX_Column(grdASTEXPT1, "BRAND_CODE", "Brand", 80)
        Set_DX_Column(grdASTEXPT1, "QTY_SOLD", "#Sold", 80, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "AMT_SOLD", "$Sold", 100, "#,##0.00", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "QTY_EOW", "#EOW", 80, "#,##0", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "AMT_EOW", "$EOW", 100, "#,##0.00", , Color.LightBlue)
        Set_DX_Column(grdASTEXPT1, "OPS_YYYYWW_MIN", "1st YW", 80)
        Set_DX_Column(grdASTEXPT1, "OPS_YYYYWW_MAX", "Lst YW", 80)
        Set_DX_Column(grdASTEXPT1, "CUST_STORE_NO", "Store", 80)

        Set_DX_Column(grdASTEXPT1, "DEMO_COMM_PCT", "Comm%", 100, "##0.00", , Color.Orange)
        Set_DX_Column(grdASTEXPT1, "AMT_COMM", "$Comm", 100, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_COMM_CLAIMED", "$CommClaim", 100, "#,##0.00", , Color.LightGreen)
        Set_DX_Column(grdASTEXPT1, "AMT_COMM_PAID", "$CommPaid", 100, "#,##0.00", , Color.LightGreen)

        Create_Summary(grdASTEXPT1, New String() {"CUST_CODE"}, "Count")
        Create_Summary(grdASTEXPT1, New String() {"QTY_SOLD", "AMT_SOLD", "AMT_COMM", "AMT_COMM_CLAIMED", "AMT_COMM_PAID"})

        grdASTEXPT1.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.Fixed = True
        grdASTEXPT1.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").Header.Fixed = True
        grdASTEXPT1.DisplayLayout.Bands(0).Columns("COLLECTION_CODE").Header.Fixed = True
        grdASTEXPT1.DisplayLayout.Bands(0).Columns("HC_CODE").Header.Fixed = True
        grdASTEXPT1.DisplayLayout.Bands(0).Columns("BRAND_CODE").Header.Fixed = True
        Sort_grdColumns(grdASTEXPT1, "CUST_CODE,CUST_STORE_NO")
    End Sub

End Class