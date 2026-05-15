Public Class SORRMAR1

#Region "General Declarations"
    Dim SOTRMAFX As String
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)

        'If ASCMAIN1.EOM = "1" Then
        '  Set_Read_Only(grpPERIOD_RANGE, True)
        'End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        Prepare_dst(True, sqlw)
        Check_if_Empty("SOTRMAF1")
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = "Accruals for " & RYPLEGEND0
        CR_params.Add("SUBT", SUBT)
        'CR_params.Add("CHKOPEN_ONLY", "0")
        Generate_Report(RPT, , SUBT)
        Print_GL()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            RWU = "N"
            If ASCMAIN1.EOM = "1" Then
                Dim RYP As String = Absx1.cmbFor("RYP0").Value
                RYP = Mid(RYP, 1, 4) & Mid(RYP, 6, 2)
                If RYP <> ASCMAIN1.CYP Then
                    EMsg &= vbCr & "For Period End you must use Current Period (" & ASCMAIN1.CYP & ")"
                Else
                    RWU = "R"
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()
        GL_Update()
    End Sub

    Overrides Function Prepare_dst(
    ByVal perform_fill As Boolean,
    ByVal ParamArray parms() As Object) As ASCBASE1

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP0)
        Dim DETL_CTL_DATE_LAST As Date = rowGLTPARM2.Item("PRD_END_DATE")

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables("")

        With dst
            ASCMAIN1.sql = "Select * from " & SOTRMAFX
            Create_TDA(.Tables.Add, "SOTRMAFX", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTRMAF1.*,ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & " from SOTRMAF1,ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = SOTRMAF1.CUST_CODE" & vbCrLf _
                & "   and SOTRMAF1.RA_NO in (Select RA_NO from " & SOTRMAFX & ")"
            Create_TDA(.Tables.Add, "SOTRMAF1", "**", 0, False)
            .Tables("SOTRMAF1").Columns.Add("RA_AMT", GetType(System.Decimal))
            .Tables("SOTRMAF1").Columns.Add("AR_PARM_KEY")

            If ASCMAIN1.CYP = RYP0 Then
                ASCMAIN1.sql = "Select SOTRMAF2.*, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_COST_STD, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & " from SOTRMAF2,ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTRMAF2.ITEM_CODE" & vbCrLf _
                & " and RA_NO in (Select RA_NO from " & SOTRMAFX & ")"

            Else
                ASCMAIN1.sql = "Select RA_NO, RA_LNO, SOTRMAF3.ITEM_CODE, RA_QTY, RA_RETAIL, RA_LINE_AMT, RA_NET_PRICE, RA_QTY_OPEN, RA_QTY_USED, RA_QTY_CANC, EDI_PRICE,NET_PRICE," _
                & " ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_COST_STD, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & " from SOTRMAF3,ICTITEM1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTRMAF3.ITEM_CODE" & vbCrLf _
                & "   and SOTRMAF3.OPS_YYYYPP = '" & RYP0 & "'" & vbCrLf _
                & "   and RA_NO In (Select RA_NO from " & SOTRMAFX & ")"

            End If
            Create_TDA(.Tables.Add, "SOTRMAF2", "**", 0, False)

            With .Tables("SOTRMAF2").Columns
                .Add("TRADE_CLASS_CODE", GetType(System.String))
                .Add("RA_AMT", GetType(System.Decimal), "ISNULL(RA_QTY,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_OPEN", GetType(System.Decimal), "ISNULL(RA_QTY_OPEN,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_USED", GetType(System.Decimal), "ISNULL(RA_QTY_USED,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_AMT_CANC", GetType(System.Decimal), "ISNULL(RA_QTY_CANC,0) * ISNULL(RA_NET_PRICE,0)")
                .Add("RA_RETAIL_EXT", GetType(System.Decimal), "IIF(ISNULL(RA_QTY,0)=0,RA_AMT/((100 - 0) / 100),ISNULL(RA_QTY,0) * ISNULL(RA_RETAIL,0))")
                .Add("RA_CGS_OPEN", GetType(System.Decimal), "ISNULL(RA_QTY_OPEN,0) * ISNULL(ITEM_COST_STD,0)")
            End With

            Create_Relation("SOTRMAF1", "SOTRMAF2", "RA_NO")

            .Tables("SOTRMAF2").Columns("TRADE_CLASS_CODE").Expression = "PARENT.TRADE_CLASS_CODE"

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

            Create_TDA(dst.Tables.Add, "SOTTCLS1", "*", 0, False)
            Create_TDA(dst.Tables.Add, "ICTCOLL1", "*", 0, False)
            Create_TDA(dst.Tables.Add, "SOTCHAN1", "*", 0, False)
        End With

        For Each TABLE_NAME As String In New String() _
            {"ICTCOLL1", "SOTTCLS1", "SOTCHAN1"}
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

        EnforceConstraints(False)

        Fill_Records("SOTRMAFX")
        Fill_Records("SOTRMAF1")
        Fill_Records("SOTRMAF2")

        EnforceConstraints(True)

        Prepare_GL_Interface("OPRA")
    End Sub

    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        Dim NYP As String = ASCMAIN1.Period_Calc(RYP0, 1)
        Dim accrual As Decimal = 0

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP0)
        Dim DETL_CTL_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")

        ' Get total Return Revenue and CGS values

        Dim REV As Decimal = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT_OPEN)", "") & "")
        Dim CGS As Decimal = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_CGS_OPEN)", "") & "")

        Dim ACCT_CODE As String = ""
        Dim rowGLTINTF1 As DataRow = Nothing
        Dim DETL_POSTING_AMT As Decimal = 0

        DETL_POSTING_AMT = -1 * REV
        ACCT_CODE = ROWs("SOTPARM1").Item("SO_PARM_RA_REC_ACCT_CODE") & ""
        rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
        rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
        Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)

        DETL_POSTING_AMT = CGS
        ACCT_CODE = ROWs("SOTPARM1").Item("SO_PARM_RA_INV_ACCT_CODE") & ""
        rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
        rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
        Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)

        For Each rowX As DataRow In ASCDATA1.SelectDistinct("SOTRMAF2", New String() {"TRADE_CLASS_CODE", "COLLECTION_CODE"}).Select("")
            Dim TRADE_CLASS_CODE As String = rowX.Item("TRADE_CLASS_CODE")
            Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)
            Dim SEG3_CODE As String = rowSOTTCLS1.Item("SEG3_CODE") & ""
            If SEG3_CODE = "" Then SEG3_CODE = TRADE_CLASS_CODE

            Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE")
            Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)
            Dim SEG2_CODE As String = rowSOTCHAN1.Item("SEG2_CODE") & ""
            If SEG2_CODE = "" Then SEG2_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")

            Dim COLLECTION_CODE As String = rowX.Item("COLLECTION_CODE")
            Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
            Dim SEG4_CODE As String = rowICTCOLL1.Item("SEG4_CODE") & ""
            If SEG4_CODE = "" Then SEG4_CODE = COLLECTION_CODE

            Dim sql2 As String = "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "'" _
                               & " and COLLECTION_CODE = '" & COLLECTION_CODE & "'"

            REV = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_AMT_OPEN)", sql2) & "")
            CGS = Val(dst.Tables("SOTRMAF2").Compute("SUM(RA_CGS_OPEN)", sql2) & "")

            DETL_POSTING_AMT = REV
            ACCT_CODE = ROWs("SOTPARM1").Item("SO_PARM_RA_REV_ACCT_CODE") & ""
            rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
            rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
            rowGLTINTF1.Item("SEG2_CODE") = SEG2_CODE
            rowGLTINTF1.Item("SEG3_CODE") = SEG3_CODE
            rowGLTINTF1.Item("SEG4_CODE") = SEG4_CODE
            Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)

            DETL_POSTING_AMT = -1 * CGS
            ACCT_CODE = ROWs("SOTPARM1").Item("SO_PARM_RA_CGS_ACCT_CODE") & ""
            rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
            rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
            rowGLTINTF1.Item("SEG2_CODE") = SEG2_CODE
            rowGLTINTF1.Item("SEG3_CODE") = SEG3_CODE
            rowGLTINTF1.Item("SEG4_CODE") = SEG4_CODE
            Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)
        Next

        Return JOURNAL_NO
    End Function

    Sub Write_GLTINTF1_Reversal(row As DataRow, YP As String)
        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1.ItemArray = row.ItemArray
        rowGLTINTF1.Item("OPS_YYYYPP") = YP
        rowGLTINTF1.Item("DETL_POSTING_AMT") = -1 * Val(rowGLTINTF1.Item("DETL_POSTING_AMT") & "")
        dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
    End Sub

    Function Write_GLTINTF1( _
                           JOURNAL_TYPE As String, _
                           JOURNAL_NO As String, _
                           ByRef JOURNAL_LNO As Integer, _
                           DETL_CTL_DATE As Date, _
                           DETL_POSTING_AMT As Decimal) As DataRow

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
        dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Return rowGLTINTF1
    End Function

    Sub Create_Work_Tables(sqlw As String)

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYP0)
        Dim DETL_CTL_DATE_LAST As Date = rowGLTPARM2.Item("PRD_END_DATE")

        Dim sqlTotals As String
        If ASCMAIN1.CYP = RYP0 Then
            sqlTotals = "Select SOTRMAF2.RA_NO" & vbCrLf _
                & ", Sum (SOTRMAF2.RA_QTY) RA_QTY" & vbCrLf _
                & ", Sum (SOTRMAF2.RA_QTY_OPEN) RA_QTY_OPEN" & vbCrLf _
                & ", Sum (SOTRMAF2.RA_QTY_USED) RA_QTY_USED" & vbCrLf _
                & ", Sum (SOTRMAF2.RA_QTY_CANC) RA_QTY_CANC" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_OPEN" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY_USED,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_USED" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY_CANC,0) * NVL(SOTRMAF2.RA_NET_PRICE,0)) RA_AMT_CANC" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY,0) * NVL(SOTRMAF2.RA_RETAIL,0)) RA_RETAIL_EXT" & vbCrLf _
                & ", Sum (NVL(SOTRMAF2.RA_QTY_OPEN,0) * NVL(ICTITEM1.ITEM_COST_STD,0)) RA_CGS_OPEN" & vbCrLf _
                & " from SOTRMAF2,SOTRMAF1,ICTITEM1 where SOTRMAF2.RA_QTY_OPEN <> 0" & vbCrLf _
                & " AND SOTRMAF1.RA_START_DATE <= '" & Format(DETL_CTL_DATE_LAST, "dd-MMM-yyyy") & "'" & vbCrLf _
                & " and ICTITEM1.ITEM_CODE = SOTRMAF2.ITEM_CODE and SOTRMAF2.RA_NO = SOTRMAF1.RA_NO" & sqlw & " group by SOTRMAF2.RA_NO"

        Else

            '                & "   and SOTRMAF3.RA_NO in (Select Distinct DETL_CTL_NO from GLTCREC3 where CREC_TYPE_CODE = 'ARR1' and OPS_YYYYPP = '" & RYP0 & "')" & vbCrLf _

            sqlTotals = "Select SOTRMAF3.RA_NO" & vbCrLf _
                & ", Sum (SOTRMAF3.RA_QTY) RA_QTY" & vbCrLf _
                & ", Sum (SOTRMAF3.RA_QTY_OPEN) RA_QTY_OPEN" & vbCrLf _
                & ", Sum (SOTRMAF3.RA_QTY_USED) RA_QTY_USED" & vbCrLf _
                & ", Sum (SOTRMAF3.RA_QTY_CANC) RA_QTY_CANC" & vbCrLf _
                & ", Sum (NVL(SOTRMAF3.RA_QTY,0) * NVL(SOTRMAF3.RA_NET_PRICE,0)) RA_AMT" & vbCrLf _
                & ", Sum (NVL(SOTRMAF3.RA_QTY_OPEN,0) * NVL(SOTRMAF3.RA_NET_PRICE,0)) RA_AMT_OPEN" & vbCrLf _
                & ", Sum (NVL(SOTRMAF3.RA_QTY_USED,0) * NVL(SOTRMAF3.RA_NET_PRICE,0)) RA_AMT_USED" & vbCrLf _
                & ", Sum (NVL(SOTRMAF3.RA_QTY_CANC,0) * NVL(SOTRMAF3.RA_NET_PRICE,0)) RA_AMT_CANC" & vbCrLf _
                & ", Sum (NVL(SOTRMAF3.RA_QTY,0) * NVL(SOTRMAF3.RA_RETAIL,0)) RA_RETAIL_EXT" & vbCrLf _
                & ", Sum (NVL(SOTRMAF3.RA_QTY_OPEN,0) * NVL(ICTITEM1.ITEM_COST_STD,0)) RA_CGS_OPEN" & vbCrLf _
                & " from SOTRMAF3,SOTRMAF1,ICTITEM1 where SOTRMAF3.RA_QTY_OPEN <> 0" & vbCrLf _
                & "   and SOTRMAF3.OPS_YYYYPP = '" & RYP0 & "'" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTRMAF3.ITEM_CODE and SOTRMAF3.RA_NO = SOTRMAF1.RA_NO" & sqlw & " group by SOTRMAF3.RA_NO"
        End If
        ASCMAIN1.sql = "Select SOTRMAF1.*" & vbCrLf _
                & ", X.RA_QTY, X.RA_QTY_OPEN, X.RA_QTY_USED, X.RA_QTY_CANC" & vbCrLf _
                & ", X.RA_AMT, X.RA_AMT_OPEN, X.RA_AMT_USED, X.RA_AMT_CANC" & vbCrLf _
                & ", X.RA_RETAIL_EXT, X.RA_CGS_OPEN" & vbCrLf _
                & " from (" & sqlTotals & ") X, SOTRMAF1" & ASCMAIN1.SQL_Add_WHERE(sqlw & " and X.RA_NO = SOTRMAF1.RA_NO")

        If SOTRMAFX = "" Then
            SOTRMAFX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTRMAFX & " Add Primary Key (RA_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTRMAFX)
            ASCDATA1.ExecuteSQL("Insert into " & SOTRMAFX & " " & ASCMAIN1.sql)
        End If

    End Sub
End Class