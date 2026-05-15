Imports System.Math

Public Class APRINVRR

    'Dim APTINVRR As String
    Dim rowAPTINVH1 As DataRow
    Dim rowAPTVEND1 As DataRow

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        Dim sqlw As String = " from APTINVH1 "
        If Absx1.optFor("OPTTA").Value = "T" Then
            sqlw = sqlw & " where INV_STATUS = 'R' "
            sqlw = sqlw & "   and INV_RECUR_OPS_YYYYPP_BEGIN <= '" & ASCMAIN1.CYP & "'"
            sqlw = sqlw & "   and (NVL(INV_RECUR_MAX,0) = 0 OR NVL(INV_RECUR_MAX,0) > NVL(INV_RECUR_GEN,0))"
            sqlw = sqlw & "   and (NVL(INV_RECUR_AMT_MAX,0) = 0 OR ABS(NVL(INV_RECUR_AMT_MAX,0)) > ABS(NVL(INV_RECUR_AMT_GEN,0)))"
        Else
            sqlw = sqlw & " where OPS_YYYYPP = '" & RYP & "' and VOUCHER_NO_RECUR is Not Null"
        End If

        With dst
            ASCMAIN1.sql = "Select APTINVH1.* " & sqlw
            Create_TDA(.Tables.Add, "APTINVH1", "**", 0)
            Fill_Records("APTINVH1")
            '.Tables.Add(ASCDATA1.GetDataTable(sql, "APTINVH1", 1))

            ASCMAIN1.sql = "Select APTINVH1.* from APTINVH1"
            Create_TDA(.Tables.Add, "APTINVH1_TEMPLATE", "**", 0, False)

            'sql = "Select APTINVH2.* from APTINVH2"
            ASCMAIN1.sql = "Select APTINVH2.* from APTINVH2 where VOUCHER_NO in (Select VOUCHER_NO " & sqlw & ")"
            Create_TDA(.Tables.Add, "APTINVH2", "**", 0)
            'sql = "Select APTINVH2.* from APTINVH2 where VOUCHER_NO in (Select VOUCHER_NO " & sqlw & ")"
            Fill_Records("APTINVH2")
            '.Tables.Add(ASCDATA1.GetDataTable(sql, "APTINVH2", 2))

            sql = "Select * from GLTACCT1 where ACCT_CODE in (Select DISTINCT ACCT_CODE from APTINVH2 where VOUCHER_NO in (Select VOUCHER_NO " & sqlw & "))"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "GLTACCT1", 1))

            sql = "Select * from APTVEND1 where VEND_CODE in (Select DISTINCT VEND_CODE " & sqlw & ")"
            .Tables.Add(ASCDATA1.GetDataTable(sql, "APTVEND1", 1))

            .Relations.Add("APTINVH2", _
            .Tables("APTINVH1").Columns("VOUCHER_NO"), _
            .Tables("APTINVH2").Columns("VOUCHER_NO"))
        End With

        Call Process_Recurring_Vouchers()

        Check_if_Empty("APTINVH1")

    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("TA", Absx1.optFor("OPTTA").Value)
        Generate_Report(RPT)
        '        Call Print_GL()
    End Sub

    Overrides Sub Update_Record()

        Dim VOUCHER_NO As String
        Dim VOUCHER_NO_RECUR As String

        dst.EnforceConstraints = False

        For Each rowAPTINVH1 In dst.Tables("APTINVH1").Rows

            VOUCHER_NO_RECUR = rowAPTINVH1("VOUCHER_NO")

            VOUCHER_NO = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
            rowAPTINVH1("VOUCHER_NO") = VOUCHER_NO
            rowAPTINVH1("REGISTER_IND") = "0"
            rowAPTINVH1("REGISTER_XNO") = Null
            rowAPTINVH1("INV_STATUS") = "O"
            rowAPTINVH1("OPS_YYYYPP") = ASCMAIN1.CYP
            rowAPTINVH1("VOUCHER_NO_RECUR") = VOUCHER_NO_RECUR

            Dim INV_RECUR_GEN As Integer = Val(rowAPTINVH1("INV_RECUR_GEN") & "")
            Dim INV_RECUR_AMT_GEN As Double = Val(rowAPTINVH1("INV_RECUR_AMT_GEN") & "")

            rowAPTINVH1("INV_RECUR_OPS_YYYYPP_BEGIN") = Null
            rowAPTINVH1("INV_RECUR_OPS_YYYYPP_LAST") = Null
            rowAPTINVH1("INV_RECUR_MAX") = Null
            'rowAPTINVH1("INV_RECUR_GEN") = Null
            rowAPTINVH1("INV_RECUR_AMT_MAX") = Null
            rowAPTINVH1("INV_RECUR_AMT_GEN") = Null

            rowAPTINVH1("INIT_OPER") = ASCMAIN1.USER_ID
            rowAPTINVH1("INIT_DATE") = DATETIME_STAMP
            rowAPTINVH1("LAST_OPER") = ASCMAIN1.USER_ID
            rowAPTINVH1("LAST_DATE") = DATETIME_STAMP

            sql = "Select * from APTINVH1 where VOUCHER_NO = '" & VOUCHER_NO_RECUR & "'"
            Fill_Records("APTINVH1_TEMPLATE", "", False, sql)
            Dim rowAPTINVH1_TEMPLATE As DataRow = _
                dst.Tables("APTINVH1_TEMPLATE").Rows.Find(VOUCHER_NO_RECUR)
            rowAPTINVH1_TEMPLATE("INV_RECUR_OPS_YYYYPP_LAST") = ASCMAIN1.CYP
            rowAPTINVH1_TEMPLATE("INV_RECUR_GEN") = INV_RECUR_GEN
            rowAPTINVH1_TEMPLATE("INV_RECUR_AMT_GEN") = INV_RECUR_AMT_GEN
            rowAPTINVH1_TEMPLATE("LAST_OPER") = ASCMAIN1.USER_ID
            rowAPTINVH1_TEMPLATE("LAST_DATE") = DATETIME_STAMP

            For Each rowAPTINVH2 As DataRow In _
                dst.Tables("APTINVH2").Select("VOUCHER_NO = '" & VOUCHER_NO_RECUR & "'", "")
                rowAPTINVH2("VOUCHER_NO") = VOUCHER_NO
            Next
        Next

        dst.Tables("APTINVH1").Merge(dst.Tables("APTINVH1_TEMPLATE"), True)

        'For Each rowAPTINVH1_TEMPLATE As DataRow In dst.Tables("APTINVH1_TEMPLATE").Rows
        '    rowAPTINVH1 = dst.Tables("APTINVH1").NewRow
        '    rowAPTINVH1.ItemArray = rowAPTINVH1_TEMPLATE.ItemArray
        '    dst.Tables("APTINVH1").Rows.Add(rowAPTINVH1)
        'Next

        dst.EnforceConstraints = True

        Call Update_Record_TDA("APTINVH1")
        'Call Update_Record_TDA("APTINVH1_TEMPLATE")
        Call Update_Record_TDA("APTINVH2")
    End Sub

    Sub Process_Recurring_Vouchers()
        ' Process Recurring Voucher Records

        If Absx1.optFor("OPTTA").Value = "A" Then
            RWU = "N"
        Else
            Dim GENS As Integer
            Dim Months As Integer
            Dim max_reached As Boolean
            Dim INV_RECUR_CYCLE As String
            Dim INV_RECUR_OPS_YYYYPP_LAST As String
            Dim LAST_AMOUNT As Double
            Dim VEND_CODE As String
            Dim VOUCHER_NO As String
            Dim RECURRING_AMOUNT As Double

            For r As Integer = dst.Tables("APTINVH1").Rows.Count - 1 To 0 Step -1
                'For Each rowAPTINVH1 In dst.Tables("APTINVH1").Rows
                rowAPTINVH1 = dst.Tables("APTINVH1").Rows(r)
                VEND_CODE = rowAPTINVH1("VEND_CODE")
                rowAPTVEND1 = LookUp("APTVEND1", VEND_CODE)

                VOUCHER_NO = rowAPTINVH1("VOUCHER_NO")
                RECURRING_AMOUNT = Val(rowAPTINVH1("INV_BALANCE") & "")
                INV_RECUR_CYCLE = rowAPTINVH1("INV_RECUR_CYCLE")
                INV_RECUR_OPS_YYYYPP_LAST = rowAPTINVH1("INV_RECUR_OPS_YYYYPP_LAST") & ""

                If INV_RECUR_OPS_YYYYPP_LAST <> "" And _
                 ((INV_RECUR_CYCLE = "M" And INV_RECUR_OPS_YYYYPP_LAST >= ASCMAIN1.CYP) Or _
                  (INV_RECUR_CYCLE = "Q" And ASCMAIN1.Period_Diff(INV_RECUR_OPS_YYYYPP_LAST, ASCMAIN1.CYP) Mod 3 <> 0) Or _
                  (INV_RECUR_CYCLE = "Y" And ASCMAIN1.Period_Diff(INV_RECUR_OPS_YYYYPP_LAST, ASCMAIN1.CYP) Mod 12 <> 0)) Then
                    For Each rowAPTINVH2 As DataRow In _
                        dst.Tables("APTINVH2").Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
                        rowAPTINVH2.Delete()
                    Next
                    rowAPTINVH1.Delete()
                Else
                    GENS = Val(rowAPTINVH1("INV_RECUR_GEN") & "") + 1
                    If INV_RECUR_CYCLE = "M" Then
                        Months = (GENS - 1) * 1
                    ElseIf INV_RECUR_CYCLE = "Q" Then
                        Months = (GENS - 1) * 3
                    ElseIf INV_RECUR_CYCLE = "Y" Then
                        Months = (GENS - 1) * 12
                    Else
                        xErrMsg = "Invalid Recurring Cycle"
                        RWU = "N"
                    End If

                    max_reached = False
                    If Val(rowAPTINVH1("INV_RECUR_AMT_MAX") & "") <> 0 And _
                       Abs(Val(rowAPTINVH1("INV_RECUR_AMT_GEN") & "")) + _
                       Abs(RECURRING_AMOUNT) > _
                       Abs(Val(rowAPTINVH1("INV_RECUR_AMT_MAX") & "")) Then
                        max_reached = True
                        LAST_AMOUNT = Val(rowAPTINVH1("INV_RECUR_AMT_MAX") & "") - _
                                      Val(rowAPTINVH1("INV_RECUR_AMT_GEN") & "")

                        rowAPTINVH1("INV_AMT") = Round(Val(rowAPTINVH1("INV_AMT") & "") * LAST_AMOUNT / RECURRING_AMOUNT, 2)
                        rowAPTINVH1("INV_DISC_BASED_ON") = Round(Val(rowAPTINVH1("INV_DISC_BASED_ON") & "") * LAST_AMOUNT / RECURRING_AMOUNT, 2)
                        rowAPTINVH1("INV_DISC_AMT") = Round(Val(rowAPTINVH1("INV_DISC_AMT") & "") * LAST_AMOUNT / RECURRING_AMOUNT, 2)
                        rowAPTINVH1("INV_BALANCE") = Round(Val(rowAPTINVH1("INV_BALANCE") & "") * LAST_AMOUNT / RECURRING_AMOUNT, 2)
                        rowAPTINVH1("INV_1099_AMT") = Round(Val(rowAPTINVH1("INV_1099_AMT") & "") * LAST_AMOUNT / RECURRING_AMOUNT, 2)

                        If Round(LAST_AMOUNT, 2) _
                        <> Round(RECURRING_AMOUNT, 2) Then
                            Dim VOUCHER_LNO As Integer = 0
                            Dim INV_LINE_AMT_total As Double = 0
                            For Each rowAPTINVH2 As DataRow In _
                                dst.Tables("APTINVH2").Select("VOUCHER_NO = '" & VOUCHER_NO & "'", "")
                                Dim INV_LINE_AMT As Double = Round(Val(rowAPTINVH2("INV_LINE_AMT") & "") * LAST_AMOUNT / RECURRING_AMOUNT, 2)
                                rowAPTINVH2("INV_LINE_AMT") = Round(INV_LINE_AMT, 2)
                                INV_LINE_AMT_total += Round(INV_LINE_AMT, 2)
                                VOUCHER_LNO = rowAPTINVH2("VOUCHER_LNO")
                            Next
                            If Round(INV_LINE_AMT_total, 2) _
                            <> Round(LAST_AMOUNT, 2) Then
                                Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").Rows.Find(New Object() {VOUCHER_NO, VOUCHER_LNO})
                                rowAPTINVH2("INV_LINE_AMT") = Round(Val(rowAPTINVH2("INV_LINE_AMT") & "") + (Round(INV_LINE_AMT_total, 2) - Round(LAST_AMOUNT, 2)), 2)
                            End If
                        End If

                    End If

                    rowAPTINVH1("INV_RECUR_GEN") = GENS
                    rowAPTINVH1("INV_RECUR_OPS_YYYYPP_LAST") = ASCMAIN1.CYP
                    If max_reached Then
                        rowAPTINVH1("INV_RECUR_AMT_GEN") = Val(rowAPTINVH1("INV_RECUR_AMT_GEN") & "") + LAST_AMOUNT
                    Else
                        rowAPTINVH1("INV_RECUR_AMT_GEN") = Val(rowAPTINVH1("INV_RECUR_AMT_GEN") & "") + RECURRING_AMOUNT
                    End If

                    rowAPTINVH1("INV_DATE") = DateAdd("m", Months, rowAPTINVH1("INV_DATE"))
                    'rowAPTINVH1("INV_DUE_DATE") = DateAdd("m", Months, rowAPTINVH1("INV_DUE_DATE"))
                    rowAPTINVH1("INV_DUE_DATE") = _
                    Calc_INV_DUE_DATE( _
                        rowAPTINVH1("TERM_CODE"), _
                        rowAPTINVH1("INV_DATE"), _
                        rowAPTINVH1("INV_BL_DATE"))
                End If
            Next

            Call Set_Added("APTINVH1")
            Call Set_Added("APTINVH2")
        End If

    End Sub

    Sub Set_Added(ByVal TABLE_NAME As String)
        dst.Tables(TABLE_NAME).AcceptChanges()
        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            row.SetAdded()
        Next
    End Sub

    Function Calc_INV_DUE_DATE( _
    ByVal TERM_CODE As String, _
    ByVal INV_DATE As Date, _
    ByVal INV_BL_DATE As Object) As Date

        Dim INV_DUE_DATE As Object = Nothing
        Dim INV_BASE_DATE As Object = Nothing

        If INV_BL_DATE & "" <> "" And rowAPTVEND1("VEND_DUE_FROM_INV_DATE") & "" <> "1" Then
            INV_BASE_DATE = INV_BL_DATE
        Else
            INV_BASE_DATE = INV_DATE
        End If

        Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE)

        Select Case rowTATTERM1.Item("TERM_DUE_TYPE") & ""

            Case "D"
                INV_DUE_DATE = INV_BASE_DATE.AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DUE") & ""))

            Case "E"

                Dim ADD_MONTHS_BASE As Integer = 1
                Dim TERM_CUTOFF_DAY As Integer = Val(rowTATTERM1.Item("TERM_CUTOFF_DAY") & "")
                Dim BASE_DD As Integer = Val(Format(INV_BASE_DATE, "dd"))
                Dim TERM_DAYS_DUE As Integer = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & "")
                Dim TERM_ADDL_MOS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_MOS") & "")
                Dim INV_BASE_DATEx As String = Format(INV_BASE_DATE, "MM/dd/yyyy")

                Select Case rowTATTERM1.Item("TERM_EOM_TYPE") & ""
                    Case "F"
                        Dim rowGLTPARM2 As DataRow = Fill_Record("GLTPARM2", Format(INV_BASE_DATE, "dd-MMM-yyyy"), True)
                        Dim YYYYMM As String = ASCMAIN1.Get_YYYYMM(rowGLTPARM2.Item("OPS_YYYYPP"), 0)
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "C"
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "S"
                        If BASE_DD <= TERM_CUTOFF_DAY _
                        And BASE_DD <= TERM_DAYS_DUE Then
                            ADD_MONTHS_BASE = 0
                        End If
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case Else
                        INV_DUE_DATE = INV_BASE_DATE
                End Select
                If TERM_ADDL_MOS > 0 Then
                    INV_DUE_DATE = INV_DUE_DATE.AddMonths(TERM_ADDL_MOS)
                End If

        End Select

        Return INV_DUE_DATE
    End Function

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            'If Absx1.cmbFor("BATCH_NO_PYMT").Text = "" Then
            '    EMsg &= vbCr & "You Cannot Post"
            'End If
        End If

    End Sub

    Private Sub optTA_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTA.ValueChanged
        grpRYP.Enabled = (optTA.Value = "A")
    End Sub
End Class
