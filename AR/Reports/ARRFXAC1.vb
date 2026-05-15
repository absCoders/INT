Public Class ARRFXAC1

#Region "General Declarations"
    Dim ARTFXAC1 As String
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        TAC.TACMAIN1.Update_Forex()

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)

        'If ASCMAIN1.EOM = "1" Then
        Set_Read_Only(grpPERIOD_RANGE, True)
        'End If
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        Prepare_dst(True, sqlw)
        Check_if_Empty("ARTFXAC1")
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

                ASCMAIN1.sql = "Select * from TATCURR1 where CURR_CODE <> '" & ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable().Select("")
                    Dim CURR_CODE As String = row.Item("CURR_CODE")
                    Dim rowTATCURR2 As DataRow = LookUp("TATCURR2", New String() {ASCMAIN1.CYP, CURR_CODE})
                    If rowTATCURR2 Is Nothing Then
                        EMsg &= vbCr & "Cannot Determine Month-End Rate for " & CURR_CODE
                    End If
                Next
            End If
        End If
    End Sub

    Overrides Sub Update_Record()
        GL_Update()
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Work_Tables("")

        With dst
            ASCMAIN1.sql = "Select * from " & ARTFXAC1
            Create_TDA(.Tables.Add, "ARTFXAC1", "**", 0, False, "", 3)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, CURR_CODE" & vbCrLf _
                & " from ARTCUST1 where CUST_CODE in (Select Distinct CUST_CODE from " & ARTFXAC1 & ")"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select TATCURR1.CURR_CODE, TATCURR1.CURR_DESC, TATCURR2.CURR_EXCH_CUR from TATCURR1,TATCURR2 where TATCURR2.CURR_CODE (+) = TATCURR1.CURR_CODE and TATCURR2.OPS_YYYYPP (+) = '" & ASCMAIN1.CYP & "'"
            Create_TDA(.Tables.Add, "TATCURR1", "**", 0, False)

            Create_TDA(dst.Tables.Add, "SOTTCLS1", "*", 0, False)
            Create_TDA(dst.Tables.Add, "SOTCHAN1", "*", 0, False)

            Create_TDA(dst.Tables.Add, "GLTINTF1", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"TATCURR1", "SOTCHAN1", "SOTTCLS1"}
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
        Fill_Records("ARTFXAC1")
        Fill_Records("ARTCUST1")

        EnforceConstraints(True)

        Prepare_GL_Interface("ARFX")
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


        Dim REV As Decimal = Val(dst.Tables("ARTFXAC1").Compute("SUM(GAIN_LOSS_ACCR)", "") & "")
        Dim ACCT_CODE As String = ""
        Dim rowGLTINTF1 As DataRow = Nothing
        Dim DETL_POSTING_AMT As Decimal = 0

        For Each rowX As DataRow In ASCDATA1.SelectDistinct("ARTFXAC1", New String() {"TRADE_CLASS_CODE"}).Select("")
            ACCT_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_GAIN_LOSS_UNREAL") & ""

            Dim TRADE_CLASS_CODE As String = rowX.Item("TRADE_CLASS_CODE")
            Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)
            Dim SEG3_CODE As String = rowSOTTCLS1.Item("SEG3_CODE") & ""
            If SEG3_CODE = "" Then SEG3_CODE = TRADE_CLASS_CODE

            Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE")
            Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)
            Dim SEG2_CODE As String = rowSOTCHAN1.Item("SEG2_CODE") & ""
            If SEG2_CODE = "" Then SEG2_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")

            Dim sql2 As String = "TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "'"
            DETL_POSTING_AMT = -1 * Val(dst.Tables("ARTFXAC1").Compute("SUM(GAIN_LOSS_ACCR)", sql2) & "")

            rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
            rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
            rowGLTINTF1.Item("SEG2_CODE") = SEG2_CODE
            rowGLTINTF1.Item("SEG3_CODE") = SEG3_CODE
            Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)
        Next

        DETL_POSTING_AMT = REV
        ACCT_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_GAIN_LOSS_ACCR") & ""
        rowGLTINTF1 = Write_GLTINTF1(JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, DETL_POSTING_AMT)
        rowGLTINTF1.Item("ACCT_CODE") = ACCT_CODE
        Write_GLTINTF1_Reversal(rowGLTINTF1, NYP)

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

        ASCMAIN1.sql = "Select ARTOPEN1.*, TATCURR2.CURR_EXCH_CUR" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE_CURR * TATCURR2.CURR_EXCH_CUR INV_BALANCE_ACCR" & vbCrLf _
            & ", ARTOPEN1.INV_BALANCE_CURR * TATCURR2.CURR_EXCH_CUR - INV_BALANCE GAIN_LOSS_ACCR" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & " from ARTOPEN1, TATCURR2, ARTCUST1" & vbCrLf _
            & " where ARTOPEN1.INV_BALANCE_CURR <> 0" & vbCrLf _
            & "   and ARTOPEN1.CURR_CODE <> '" & ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "'" & vbCrLf _
            & "   and TATCURR2.CURR_CODE = ARTOPEN1.CURR_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = ARTOPEN1.CUST_CODE" & vbCrLf _
            & "   and TATCURR2.OPS_YYYYPP = '" & RYP0 & "'"

        If ARTFXAC1 = "" Then
            ARTFXAC1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTFXAC1 & " Add Primary Key (CUST_CODE, INV_TYPE, INV_NUM)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTFXAC1)
            ASCDATA1.ExecuteSQL("Insert into " & ARTFXAC1 & " " & ASCMAIN1.sql)
        End If

    End Sub
End Class