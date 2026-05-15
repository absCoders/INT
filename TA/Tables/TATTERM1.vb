Public Class TATTERM1

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grpTester.Visible = tf
    End Sub

    Private Sub optTERM_DUE_TYPE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTERM_DUE_TYPE.ValueChanged
        Set_Terms_Controls()
    End Sub

    Sub Set_Terms_Controls()

        If optTERM_DUE_TYPE.Value & "" = "E" Then
            lblTERM_DAYS_DUE.Text = "Day of Month Due"
            numTERM_DAYS_DUE.MinValue = 1
            'If numTERM_DAYS_DUE.Value = 0 Then
            '    numTERM_DAYS_DUE.Value = 1
            'End If
        Else
            lblTERM_DAYS_DUE.Text = "Days til Net Due"
            numTERM_DAYS_DUE.MinValue = 0
        End If

        grpE.Visible = (optTERM_DUE_TYPE.Value & "" = "E")

    End Sub
    Private Sub dteInvoiceDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dteInvoiceDate.ValueChanged
        btnTest.PerformClick()
    End Sub

    Private Sub optTERM_EOM_TYPE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTERM_EOM_TYPE.ValueChanged
        Absx1.numFor("TERM_CUTOFF_DAY").Visible = (optTERM_EOM_TYPE.Value & "" = "S")
    End Sub

    Sub CALC_DUE_DATE()
        If Absx1.optFor("TERM_DUE_TYPE").Value & "" = "" Then
            MsgBox("Aging Parameters Type Not Defined", MsgBoxStyle.OkOnly, "Cannot Calculate a Due Date")
            Exit Sub
        End If

        If optTERM_DUE_TYPE.Value & "" = "E" And Val(numTERM_DAYS_DUE.Value & "") = 0 Then
            MsgBox("Invalid Day of the Month", MsgBoxStyle.OkOnly, "Cannot Calculate a Due Date")
            Exit Sub
        End If

        Dim INV_DUE_DATE As Object = Nothing
        Dim INV_BASE_DATE As Object = dteInvoiceDate.Value
        dteDiscDueDate.Value = Null
        If INV_BASE_DATE Is Nothing Then
            dteDueDate.Value = Null
            Exit Sub
        End If

        Select Case Absx1.optFor("TERM_DUE_TYPE").Value
            Case "D"
                INV_DUE_DATE = INV_BASE_DATE.AddDays(Val(Absx1.numFor("TERM_DAYS_DUE").Value & ""))

            Case "E"
                Dim ADD_MONTHS_BASE As Integer = 1
                Dim TERM_CUTOFF_DAY As Integer = Val(Absx1.numFor("TERM_CUTOFF_DAY").Value & "")
                Dim BASE_DD As Integer = Val(Format(INV_BASE_DATE, "dd"))
                Dim TERM_DAYS_DUE As Integer = Val(Absx1.numFor("TERM_DAYS_DUE").Value & "")
                Dim TERM_ADDL_MOS As Integer = Val(Absx1.numFor("TERM_ADDL_MOS").Value & "")
                Dim INV_BASE_DATEx As String = Format(INV_BASE_DATE, "MM/dd/yyyy")

                Select Case Absx1.optFor("TERM_EOM_TYPE").Value
                    Case "F"
                        ASCMAIN1.sql = "Select GLTPARM2.* " _
                         & " from GLTPARM2 " _
                         & " where OPS_YYYYPP = " _
                         & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
                         & "  where GLTPARM2.PRD_END_DATE >= '" & Format(INV_BASE_DATE, "dd-MMM-yyyy") & "')"
                        Dim rowGLTPARM2 As DataRow = ASCDATA1.GetDataRow
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
                End Select
                If TERM_ADDL_MOS > 0 Then
                    INV_DUE_DATE = INV_DUE_DATE.AddMonths(TERM_ADDL_MOS)
                End If
        End Select

        dteDueDate.Value = INV_DUE_DATE

        If Val(Absx1.numFor("TERM_DISC_PERC").Value & "") <> 0 Then
            If Absx1.chkFor("TERM_DISC_ELIG_DUE").Checked Then
                dteDiscDueDate.Value = dteDueDate.Value
            Else
                If Val(Absx1.numFor("TERM_DISC_PERC").Value & "") <> 0 Then
                    dteDiscDueDate.Value = DateValue(dteDueDate.Value & "").AddDays(Val(Absx1.numFor("TERM_DAYS_DISC").Value & ""))
                End If
            End If
        End If
    End Sub

    Private Sub btnTest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTest.Click
        Call CALC_DUE_DATE()
    End Sub

    Private Sub grpTester_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grpTester.Click

    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        Select Case eItemKey
            Case "Update"
                If optTERM_DUE_TYPE.CheckedIndex = -1 Then
                    EMsg &= vbCr & "Invalid Terms Type"
                End If

                If Absx1.txtFor("TERM_DESC").Text = "" Then
                    EMsg &= vbCr & "Terms Description Required"
                End If

                If Absx1.optFor("TERM_USE").CheckedIndex = -1 Then
                    EMsg &= vbCr & "Terms Use Required"
                End If

                If Absx1.optFor("TERM_STATUS").CheckedIndex = -1 Then
                    EMsg &= vbCr & "Terms Status Required"
                End If

                If optTERM_DUE_TYPE.Value & "" = "E" Then
                    If Val(numTERM_DAYS_DUE.Value & "") = 0 Then
                        EMsg &= vbCr & "Invalid Day of the Month"
                    End If
                    If optTERM_EOM_TYPE.CheckedIndex = -1 Then
                        EMsg &= vbCr & "Invalid EOM-Specific Parameters"
                    End If

                End If
        End Select
    End Sub


    Overrides Sub Show_Record_Special()
        Set_Terms_Controls()
    End Sub

End Class