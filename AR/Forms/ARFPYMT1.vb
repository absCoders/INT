Public Class ARFPYMT1

    Dim rowGLTBANK1 As DataRow
    Dim rowGLTPARM2 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String

    Dim blnFX_Support As Boolean = False

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")
        Get_PARM("GLTPARM1")

        With dst
            Create_TDA(.Tables.Add, "ARTPYMT1", "*")
            Create_TDA(.Tables.Add, "ARTPYMT2", "*", 1)
            Create_TDA(.Tables.Add, "ARTPYMTC", "*", 1)

            ASCMAIN1.sql = "SELECT ARTPYMT1.*, X.PYMT_AMT, X.PYMT_AMT_CURR" _
            & " from ARTPYMT1, " _
            & " (Select ARTPYMT2.PYMT_BATCH_NO" _
            & ", Sum (ARTPYMT2.CUST_PYMT_AMT) PYMT_AMT" _
            & ", Sum (ARTPYMT2.CUST_PYMT_AMT_CURR) PYMT_AMT_CURR" _
            & " from ARTPYMT1,ARTPYMT2" _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
            & "   and ARTPYMT1.STATUS LIKE :PARM1 and ARTPYMT1.INIT_OPER LIKE :PARM2 and OPS_YYYYPP = :PARM3" _
            & " group by ARTPYMT2.PYMT_BATCH_NO) X" _
            & " where ARTPYMT1.PYMT_BATCH_NO = X.PYMT_BATCH_NO"
            Create_TDA(.Tables.Add, "ARTPYMTX", "**", 0, False, "VVV", 1)


            With .Tables.Add("ARTPYMTL")
                .Columns.Add("SKIP")
                .Columns("SKIP").DefaultValue = "0"
                .Columns.Add("BANK_CODE")
                .Columns.Add("BANK_ACCT_ID")
                .Columns.Add("PYMT_BATCH_DATE", GetType(System.DateTime))
                .Columns.Add("PYMT_NOTE")
                .Columns.Add("CUST_PYMT_REF_NO")
                .Columns.Add("CUST_PYMT_AMT", GetType(System.Decimal))
                .Columns.Add("CUST_NAME_REF")
                .Columns.Add("CUST_CODE")
                .Columns.Add("LINE", GetType(System.Int32))
                ' .Columns.Add("CLONED")
            End With

        End With

        rowGLTPARM2 = LookUp("GLTPARM2", ASCMAIN1.CYP)

        grdARTPYMTX.DataSource = dst.Tables("ARTPYMTX")
        grdARTPYMT2.DataSource = dst.Tables("ARTPYMT2")
        grdARTPYMTL.DataSource = dst.Tables("ARTPYMTL")

        Create_Summary(grdARTPYMTX, "PYMT_BATCH_NO", "Count")
        Create_Summary(grdARTPYMTX, "PYMT_AMT")

        Create_Summary(grdARTPYMT2, "PYMT_BATCH_LNO", "Count")
        Create_Summary(grdARTPYMT2, New String() {"CUST_PYMT_AMT", "CUST_PYMT_AMT_CURR"})

        Create_Summary(grdARTPYMTL, "BANK_CODE", "Count")
        Create_Summary(grdARTPYMTL, New String() {"CUST_PYMT_AMT"})

        blnFX_Support = ASCDATA1.GetDataTable("*", "TATCURR1").Rows.Count > 1

        If Not blnFX_Support Then
            grpCURR_CODE.Visible = False
            With grdARTPYMTX.DisplayLayout.Bands(0)
                .Columns("CURR_CODE").Hidden = True
                .Columns("CURR_EXCH_RATE").Hidden = True
                .Columns("PYMT_AMT_CURR").Hidden = True
            End With
        Else
            grpCURR_CODE.Visible = True
        End If

        Set_Read_Only(grpCURR_CODE, True)

        With grdARTPYMT2.DisplayLayout.Bands(0)
            .Columns("PYMT_BATCH_LNO").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CUST_CODE").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("CUST_NAME").CellAppearance.BackColor = Drawing.Color.Beige
            .Columns("NON_AR").CellAppearance.BackColor = Drawing.Color.Beige

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "CURR_CODE" Or gcol.Key = "CURR_EXCH_RATE" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With
        grdARTPYMT2.DisplayLayout.Bands(0).Columns("NON_AR").Tag = "N"
        grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_PYMT_AMT").Header.Caption = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & " Amount"

        With grdARTPYMTX.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "CURR_CODE" Or gcol.Key = "CURR_EXCH_RATE" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                End If
            Next
        End With

        cbeYP.DataSource = ASCDATA1.GetDataTable("Select OPS_YYYYPP, LEGEND from GLTPARM2 where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "' order by OPS_YYYYPP DESC")
        cbeYP.SelectedItem = cbeYP.Items(0)


        With grdARTPYMTL.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.Header.Appearance.BackColor = Drawing.Color.White
                gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                If gcol.Key = "SKIP" Or gcol.Key = "CUST_CODE" Or gcol.Key = "BANK_CODE - NOT VALIDATED" Then
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next
        End With

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("BANK_CODE")

                If EMsg = "" Then
                    ' Validate Bank - shouldn't this happen also on Edit?
                    Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                    If rowGLTBANK1.Item("CURR_CODE") & "" <> Absx1.txtFor("CURR_CODE").Text Then
                        EMsg &= vbCr & "Currency Code defined for Bank (" & rowGLTBANK1.Item("CURR_CODE") & ") does not match Currency for Batch (" & Absx1.txtFor("CURR_CODE").Text & ")"
                    End If
                    If InStr("AR", rowGLTBANK1.Item("BANK_USE") & "") = 0 Then
                        EMsg &= vbCr & "Bank " & Absx1.txtFor("BANK_CODE").Text & " is not defined for use in A/R"
                    End If
                End If

                If Absx1.txtFor("CURR_CODE").Text <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                    If Val(Absx1.numFor("CURR_EXCH_RATE").Value & "") = 0 Then
                        EMsg &= vbCr & "Rate of Exchange required for Foreign Currency (" & Absx1.txtFor("CURR_CODE").Text & ")"
                    End If
                    If EMsg = "" Then
                        If Val(Absx1.numFor("CURR_EXCH_RATE").Value & "") = 1 Then
                            If MsgBox("Rate of Exchange is set to 1.00" & vbCrLf & "OK to continue?", MsgBoxStyle.YesNo, "Rate of Exchange is set to 1.00") = MsgBoxResult.No Then
                                Exit Sub

                            End If
                        End If
                    End If
                End If

                If Absx1.dteFor("PYMT_BATCH_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "No Receipt Date Specified"
                End If

            Case "Edit"
                Validate_Code("PYMT_BATCH_NO")

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("ARTPYMT1", Absx1.txtFor("PYMT_BATCH_NO").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"
                If grdARTPYMT2.Rows.Count = 0 Then
                    EMsg &= vbCr & "No Payments Entered into Batch"
                Else
                    Dim negative_amounts_present As Boolean = False

                    For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Select("", "", DataViewRowState.CurrentRows)
                        If rowARTPYMT2.Item("CUST_PYMT_REF_DATE") & "" <> "" Then
                            If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                             < Format(rowARTPYMT2.Item("CUST_PYMT_REF_DATE"), "yyyyMMdd") Then
                                EMsg &= vbCr & "Customer Payment Date may not be later than Batch Payment Receipt Date (see Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO") & ")"
                            End If
                            If Format(Now.AddYears(-1), "yyyyMMdd") _
                             > Format(rowARTPYMT2.Item("CUST_PYMT_REF_DATE"), "yyyyMMdd") Then
                                EMsg &= vbCr & "Customer Payment Date may not be more than 1 year ago (see Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO") & ")"
                            End If
                        End If

                        Dim CURR_CODE As String = rowARTPYMT2.Item("CURR_CODE") & ""
                        If CURR_CODE = "" Then
                            CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                            rowARTPYMT2.Item("CURR_CODE") = CURR_CODE
                        Else
                            Dim rowTATCURR1 As DataRow = LookUp("TATCURR1", CURR_CODE)
                            If rowTATCURR1 Is Nothing Then
                                EMsg &= vbCr & "Invalid Currency Code (" & CURR_CODE & ") on Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO")
                            Else
                                If Absx1.txtFor("CURR_CODE").Text <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                                    If CURR_CODE <> Absx1.txtFor("CURR_CODE").Text Then
                                        EMsg &= vbCr & "Invalid Currency Code (" & CURR_CODE & ") on Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO")
                                    End If
                                End If
                                If CURR_CODE <> ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                                    Dim CURR_EXCH_RATE As Decimal = Val(rowARTPYMT2.Item("CURR_EXCH_RATE") & "")
                                    If CURR_EXCH_RATE = 1 Or CURR_EXCH_RATE = 0 Then
                                        EMsg &= vbCr & "Invalid Currency Rate of Exchange for Currency " & CURR_CODE & ") on Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO")
                                    End If
                                    Dim CURR_EXCH_RATE_now As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, CURR_CODE, Now.Date)
                                    If CURR_EXCH_RATE_now = 0 OrElse System.Math.Abs(100 * (CURR_EXCH_RATE - CURR_EXCH_RATE_now) / CURR_EXCH_RATE) > 10 Then
                                        EMsg &= vbCr & "Invalid Currency Rate of Exchange for Currency " & CURR_CODE & " (> 10% variance with current daily rate of " & Format(CURR_EXCH_RATE_now, "#.0000") & ")"
                                    End If
                                End If
                            End If
                        End If

                        If CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                            rowARTPYMT2.Item("CURR_EXCH_RATE") = 1
                        End If

                        CUST_CODE = rowARTPYMT2.Item("CUST_CODE") & ""
                        If CUST_CODE <> "" Then
                            rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                            Dim CURR_CODE_customer As String = rowARTCUST1.Item("CURR_CODE")
                            If CURR_CODE_customer = "" Then
                                CURR_CODE_customer = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                            End If
                            If CURR_CODE_customer <> CURR_CODE Then
                                EMsg &= vbCr & "Customer " & CUST_CODE & " default Currency Code is " & CURR_CODE_customer
                                'If MsgBox("Customer " & CUST_CODE & " default Currency Code is " & CURR_CODE_customer _
                                '          & vbCrLf & vbCrLf & "Continue Anyway?", MsgBoxStyle.YesNo, "Warning") = MsgBoxResult.No Then
                                '    Exit Sub
                                'End If
                            End If
                        End If

                        Dim PYMT_AMT = Val(rowARTPYMT2.Item("CUST_PYMT_AMT_CURR") & "")
                        If System.Math.Round(PYMT_AMT, 2) = 0 Then
                            EMsg &= vbCr & "Zero Amount is Not Permitted (see Line " & rowARTPYMT2.Item("PYMT_BATCH_LNO") & ")"
                        End If

                        If System.Math.Round(PYMT_AMT, 2) < 0 Then
                            negative_amounts_present = True
                        End If

                    Next

                    If negative_amounts_present And EMsg = "" Then
                        'Removed by request from Melinda 11/01/2011
                        'If MsgBox("There are Negative Checks in this Batch." & vbCr & "You can see them by sorting the Amount Column." & vbCr & vbCr & "Click YES to Proceed, or NO to return to edit the batch.", MsgBoxStyle.YesNo, "WARNING: Negative Amount Payments") = MsgBoxResult.No Then
                        '    Exit Sub
                        'End If
                    End If
                End If

                If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                 > Format(rowGLTPARM2.Item("PRD_END_DATE"), "yyyyMMdd") Then
                    EMsg &= vbCr & "Payment Receipt Date cannot be later than Current Period End Date (" & Format(rowGLTPARM2.Item("PRD_END_DATE"), "MM/dd/yyyy") & ")"
                Else
                    Dim rowGLTPARM2_LYP = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
                    If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                    <= Format(rowGLTPARM2_LYP.Item("PRD_END_DATE"), "yyyyMMdd") Then

                        If MsgBox("Payment Receipt Date is prior to Current Period Start Date (" & Format(CDate(rowGLTPARM2_LYP.Item("PRD_END_DATE")).AddDays(1), "MM/dd/yyyy") & ")" _
                                  & vbCrLf & vbCrLf & "Proceed with this Entry?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

                If Format(Absx1.dteFor("PYMT_BATCH_DATE").Value, "yyyyMMdd") _
                 < Format(CDate(rowGLTPARM2.Item("PRD_END_DATE")).AddYears(-1), "yyyyMMdd") Then
                    EMsg &= vbCr & "Payment Receipt Date cannot be more than 1 year ago"
                End If

            Case "Delete"
                If MsgBox("Are you sure you want to Delete?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                    Exit Sub
                End If



            Case "Load XLS"
                Dim FILENAME As String = ""
                Dim DT As DataTable
                dt = Gembox_Import_Sheet_to_DataTable(0, FILENAME)
                If dt Is Nothing Then
                    EMsg &= vbCr & "You Must Select a Bank Transaction Export file to Create Payment Receipt Batches"
                Else
                    If dt.Rows.Count = 0 Or dt.Columns.Count = 0 Then
                        EMsg &= vbCr & "Error Loading Export file"
                    Else

                        ' THIS ROUTINE IS HARD CODED TO A SPECIFIC BANK TRANS XLS FROM AHA - PROB NEEDS TO BE PARAMETERIZED
                        dst.Tables("ARTPYMTC").Rows.Clear()

                        dst.Tables("ARTPYMTL").Rows.Clear()
                        Dim BANK_CODE As String = ""
                        Dim BANKS As New Dictionary(Of String, String)
                        Dim r As Integer = -1
                        For Each ROW As DataRow In DT.Select()
                            r += 1
                            If r > 0 Then
                                Dim rowARTPYMTL As DataRow = dst.Tables("ARTPYMTL").NewRow
                                Dim BANK_ACCT_ID As String = ROW.Item(1) & "" '
                                If BANK_ACCT_ID <> "" Then
                                    rowARTPYMTL.Item("BANK_ACCT_ID") = BANK_ACCT_ID
                                    If Not BANKS.ContainsKey(BANK_ACCT_ID) Then
                                        ASCMAIN1.sql = "Select BANK_CODE from GLTBANK1 where BANK_ACCT_ID = :PARM1"
                                        BANK_CODE = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", BANK_ACCT_ID)
                                        Fill_Records("ARTPYMTC", BANK_CODE, False)
                                        BANKS.Add(BANK_ACCT_ID, BANK_CODE)
                                    End If
                                    rowARTPYMTL.Item("BANK_CODE") = BANKS(BANK_ACCT_ID)
                                End If
 
                                rowARTPYMTL.Item("PYMT_BATCH_DATE") = CDate(ROW.Item(4)).Date
                                rowARTPYMTL.Item("CUST_PYMT_REF_NO") = ROW.Item(5)

                                rowARTPYMTL.Item("CUST_PYMT_AMT") = Val(ROW.Item(7) & "")

                                Dim PYMT_NOTE As String = ROW.Item(10) & ""
                                PYMT_NOTE = Mid(PYMT_NOTE, 1, 60)
                                rowARTPYMTL.Item("PYMT_NOTE") = PYMT_NOTE

                                Dim CUST_NAME_REF As String = ROW.Item(10) & ""
                                CUST_NAME_REF = Mid(CUST_NAME_REF, 1, 18)
                                rowARTPYMTL.Item("CUST_NAME_REF") = CUST_NAME_REF

                                If CUST_NAME_REF <> "" Then
                                    Dim rowARTPYMTC As DataRow = dst.Tables("ARTPYMTC").Rows.Find(New String() {BANK_CODE, CUST_NAME_REF})
                                    If rowARTPYMTC IsNot Nothing Then
                                        rowARTPYMTL.Item("CUST_CODE") = rowARTPYMTC.Item("CUST_CODE")
                                    End If
                                End If


                                rowARTPYMTL.Item("LINE") = r
                                dst.Tables("ARTPYMTL").Rows.Add(rowARTPYMTL)
                            End If
                        Next
                        dst.Tables("ARTPYMTL").AcceptChanges()
                        Sort_grdColumns(grdARTPYMTL, "PYMT_BATCH_DATE,CUST_NAME_REF")
                    End If
                End If


            Case "Update XLS"

                Dim CUST_CODEs As New List(Of String)
                Dim BANK_CODEs As New List(Of String)
                Dim blnCUST_CODE As Boolean = False
                Dim blnBANK_CODE As Boolean = False

                Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
                Dim DTE2 As Date = rowGLTPARM2.Item("PRD_END_DATE")

                rowGLTPARM2 = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
                Dim DTE1 As Date = rowGLTPARM2.Item("PRD_END_DATE")
                DTE1 = DTE1.AddDays(1)

                For Each row As DataRow In dst.Tables("ARTPYMTL").Select("")


                    Dim PYMT_BATCH_DATE As Date = row.Item("PYMT_BATCH_DATE")

                    If Format(PYMT_BATCH_DATE, "yyyyMMdd") > Format(DTE2, "yyyyMMdd") _
                    Or Format(PYMT_BATCH_DATE, "yyyyMMdd") < Format(DTE1, "yyyyMMdd") Then
                        EMsg &= vbCr & "All Receipts must be dated between " & Format(DTE1, "yyyyMMdd") & " and " & Format(DTE2, "yyyyMMdd")
                    End If

                    Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
                    If CUST_CODE = "" Then
                        If Not blnCUST_CODE Then
                            EMsg &= vbCr & "All Receipts must have a valid Customer Code"
                            blnCUST_CODE = True
                        End If
                    Else
                        If CUST_CODE = "CONSUMER" Then
                            EMsg &= vbCr & "CONSUMER Payments are not permitted"
                            Exit For
                        End If
                        If Not CUST_CODEs.Contains(CUST_CODE) Then
                            If LookUp("ARTCUST1", CUST_CODE) Is Nothing Then
                                EMsg &= vbCr & "Invalid Customer Code used (" & CUST_CODE * ")"
                            End If
                        End If
                    End If

                    Dim BANK_CODE As String = row.Item("BANK_CODE") & ""
                    If BANK_CODE = "" Then
                        If Not blnBANK_CODE Then
                            EMsg &= vbCr & "All Receipts must have a valid Bank Code"
                            blnBANK_CODE = True
                        End If

                    Else
                        If Not BANK_CODEs.Contains(BANK_CODE) Then
                            If LookUp("GLTBANK1", BANK_CODE) Is Nothing Then
                                EMsg &= vbCr & "Invalid Bank Code used (" & BANK_CODE * ")"
                            End If
                        End If
                    End If

                    Dim CUST_PYMT_AMT As Decimal = Val(row.Item("CUST_PYMT_AMT") & "")
                    If CUST_PYMT_AMT <= 0 Then
                        EMsg &= vbCr & "Invalid Payment Amount (" & Format(CUST_PYMT_AMT, "$#.00") & ")"
                    End If
                Next
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

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

            Case "Cancel", "Done"
                Mode_Settings(False)

            Case "Delete"
                Delete_Record()
                Mode_Settings(False)



            Case "Load XLS"
                EntryMode = "X"
                '   Load_Record_XLS()
                Mode_Settings(True)

            Case "Update XLS"
                Update_Record_XLS()
                Mode_Settings(False)

            Case "Cancel XLS"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("New").Settings.Enabled = not_iScreenMode
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Delete").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode

                    .Items("Cancel").Visible = Not This_Record_Inquiry_Only
                    .Items("Done").Visible = This_Record_Inquiry_Only

                    .Items("Update").Visible = ScreenMode And Not This_Record_Inquiry_Only
                    .Items("Delete").Visible = ScreenMode And Not This_Record_Inquiry_Only And (EntryMode = "E")


                End With

                With .Groups("XLS Options")

                    .Items("Load XLS").Settings.Enabled = not_iScreenMode
                    .Items("Update XLS").Settings.Enabled = iScreenMode
                    .Items("Cancel XLS").Settings.Enabled = iScreenMode
                End With

                .Groups("Screen Control").Visible = Not (EntryMode = "X")

                .Groups("XLS Options").Visible = (Not ScreenMode Or EntryMode = "X") And (ASCMAIN1.CLIENT = "AHA")

                .Groups("Batch Options").Visible = Not ScreenMode
                .Groups("Find Customer").Visible = ScreenMode And Not This_Record_Inquiry_Only And Not (EntryMode = "X")

            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only(grpCURR_CODE, True)

        If Not This_Record_Inquiry_Only Then
            Absx1.dteFor("PYMT_BATCH_DATE").ReadOnly = False
        End If

        grdARTPYMT2.Visible = ScreenMode And Not (EntryMode = "X")
        grdARTPYMTX.Visible = Not ScreenMode
        grdARTPYMTL.Visible = ScreenMode And (EntryMode = "X")

        If ScreenMode Then

            If EntryMode = "X" Then
                spl.Panel1Collapsed = True


            Else

                With grdARTPYMT2.DisplayLayout.Override
                    If This_Record_Inquiry_Only Then
                        .AllowAddNew = UltraWinGrid.AllowAddNew.No
                        .AllowDelete = DefaultableBoolean.False
                        .AllowUpdate = DefaultableBoolean.False
                    Else
                        .AllowAddNew = UltraWinGrid.AllowAddNew.FixedAddRowOnTop
                        .AllowDelete = DefaultableBoolean.True
                        .AllowUpdate = DefaultableBoolean.True
                    End If
                End With

                Dim CURR_CODE As String = Absx1.txtFor("CURR_CODE").Text
                Dim GL_PARM_CURR_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

                Dim COLUMN_NAME As String = "CUST_PYMT_AMT_CURR"
                Dim summary As UltraWinGrid.SummarySettings
                For Each summary In grdARTPYMT2.DisplayLayout.Bands(0).Summaries
                    If summary.Key = "CUST_PYMT_AMT_CURR" Then
                        grdARTPYMT2.DisplayLayout.Bands(0).Summaries.Remove(summary)
                    End If
                Next

                If CURR_CODE <> GL_PARM_CURR_CODE Then
                    Create_Summary(grdARTPYMT2, New String() {"CUST_PYMT_AMT_CURR"})
                End If

                With grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_PYMT_AMT_CURR").Header
                    If CURR_CODE = GL_PARM_CURR_CODE Then
                        .Caption = "Amount"
                    Else
                        .Caption = CURR_CODE & " Amount"
                    End If
                End With

                grdARTPYMT2.DisplayLayout.Bands(0).Columns("CURR_CODE").Hidden = (CURR_CODE <> GL_PARM_CURR_CODE) Or Not blnFX_Support
                grdARTPYMT2.DisplayLayout.Bands(0).Columns("CURR_EXCH_RATE").Hidden = (CURR_CODE <> GL_PARM_CURR_CODE) Or Not blnFX_Support
            End If
        Else
            Clear_Record()

            spl.Panel1Collapsed = False

        End If

    End Sub

    Sub Clear_Record()

        For Each TABLE_NAME As String In New String() {"ARTPYMT1", "ARTPYMT2", "ARTPYMTL", "ARTPYMTC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If Absx1.txtFor("BANK_CODE").Text = "" Then
            Absx1.txtFor("BANK_CODE").Text = ROWs("ARTPARM1").Item("AR_PARM_BANK_CODE")
            Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
            Absx1.numFor("CURR_EXCH_RATE").Value = 1
        End If

        Load_ARTPYMTX()

    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        If EntryMode = "N" Then
            HFs("PYMT_BATCH_NO") = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
        End If

        Dim rowARTPYMT1 As DataRow = Fill_Record("ARTPYMT1", HFs("PYMT_BATCH_NO"), EntryMode = "N")
        If EntryMode = "N" Then
            rowARTPYMT1.Item("BANK_CODE") = HFs("BANK_CODE")
            rowARTPYMT1.Item("CURR_CODE") = HFs("CURR_CODE")
            rowARTPYMT1.Item("CURR_EXCH_RATE") = HFs("CURR_EXCH_RATE")
            rowARTPYMT1.Item("PYMT_BATCH_DATE") = HFs("PYMT_BATCH_DATE")
            rowARTPYMT1.Item("STATUS") = "0"
            rowARTPYMT1.Item("PYMT_SOURCE") = "MAN"
            rowARTPYMT1.Item("OPS_YYYYPP") = ASCMAIN1.CYP
        Else
            If rowARTPYMT1.Item("STATUS") & "" <> "0" Then
                This_Record_Inquiry_Only = True
            Else
                This_Record_Inquiry_Only = False
            End If
        End If
        Fill_Records("ARTPYMT2", HFs("PYMT_BATCH_NO"))

        For Each grow As UltraWinGrid.UltraGridRow In grdARTPYMT2.Rows
            If grow.Cells("CUST_CODE").Value & "" = "" Then
                grdARTPYMT2.ActiveRow = grow
                grow.Cells("NON_AR").Value = "1"
                grow.Update()
            End If
        Next
        Sort_grdColumns(grdARTPYMT2, "PYMT_BATCH_LNO")

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        grdARTPYMT2.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdARTPYMT2.DisplayLayout.Bands(0).SortedColumns.Add("PYMT_BATCH_LNO", False)

    End Sub

    Sub Delete_Record()
        BeginTrans()

        Delete_Rows("ARTPYMT1")
        Delete_Rows("ARTPYMT2")

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")

        CommitTrans("Records Deleted")
    End Sub

    Sub Update_Record()
        BeginTrans()

        Dim CURR_CODE As String = Absx1.txtFor("CURR_CODE").Text
        If CURR_CODE = "" Then
            CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        End If
        Dim CURR_EXCH_RATE As Decimal = Val(Absx1.numFor("CURR_EXCH_RATE").Value & "")
        If CURR_EXCH_RATE = 0 Or CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then CURR_EXCH_RATE = 1

        For Each rowARTPYMT2 As DataRow In dst.Tables("ARTPYMT2").Select("")
            Dim CURR_EXCH_RATE2 As Decimal = Val(rowARTPYMT2.Item("CURR_EXCH_RATE") & "")
            Dim CURR_CODE2 As String = rowARTPYMT2.Item("CURR_CODE")
            If CURR_CODE2 = "" Then
                CURR_CODE2 = CURR_CODE
                rowARTPYMT2.Item("CURR_CODE") = CURR_CODE2
                CURR_EXCH_RATE2 = CURR_EXCH_RATE
            End If
            If CURR_CODE2 = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                rowARTPYMT2.Item("CURR_CODE") = CURR_CODE2
                CURR_EXCH_RATE2 = 1
            End If
            rowARTPYMT2.Item("CURR_EXCH_RATE") = CURR_EXCH_RATE2
            rowARTPYMT2.Item("CUST_PYMT_AMT") = Val(rowARTPYMT2.Item("CUST_PYMT_AMT_CURR") & "") * CURR_EXCH_RATE2
            rowARTPYMT2.Item("INIT_DATE") = DATETIME_STAMP
            rowARTPYMT2.Item("INIT_OPER") = ASCMAIN1.USER_ID
        Next

        INIT_LAST("ARTPYMT1", , , True)

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")

        CommitTrans("Update Complete")
    End Sub


    Sub Update_Record_XLS()
        BeginTrans()

        Dim CURR_CODE As String = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
        Dim CURR_EXCH_RATE As Decimal = 1

        Dim gby() As String = {"BANK_CODE", "PYMT_BATCH_DATE"}
        If ASCMAIN1.CLIENT = "AHA" Then
            gby = {"BANK_CODE", "PYMT_BATCH_DATE", "LINE"}
        End If

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ARTPYMTL"), gby).Select("")
            Dim BANK_CODE As String = row.Item("BANK_CODE")
            Dim PYMT_BATCH_DATE As Date = row.Item("PYMT_BATCH_DATE")
            Dim PYMT_BATCH_NO As String = ASCMAIN1.Next_Control_No("ARTPYMT1.PYMT_BATCH_NO")
            Dim rowARTPYMT1 As DataRow = dst.Tables("ARTPYMT1").NewRow
            With rowARTPYMT1
                .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                .Item("PYMT_BATCH_DATE") = PYMT_BATCH_DATE
                .Item("BANK_CODE") = BANK_CODE
                .Item("STATUS") = "0"
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("PYMT_SOURCE") = "MAN"
                .Item("CURR_CODE") = CURR_CODE
                .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("ARTPYMT1").Rows.Add(rowARTPYMT1)

            Dim PYMT_BATCH_LNO As Integer = 0
            Dim sql As String = "BANK_CODE = '" & BANK_CODE & "' and PYMT_BATCH_DATE = '" & Format(PYMT_BATCH_DATE, "MM/dd/yy") & "'"
            If ASCMAIN1.CLIENT = "AHA" Then
                sql &= " and LINE = " & row.Item("LINE")
            End If
            For Each rowARTPYMTL As DataRow In dst.Tables("ARTPYMTL").Select(sql)
                Dim rowARTPYMT2 As DataRow = dst.Tables("ARTPYMT2").NewRow
                With rowARTPYMT2
                    .Item("PYMT_BATCH_NO") = PYMT_BATCH_NO
                    PYMT_BATCH_LNO += 1
                    .Item("PYMT_BATCH_LNO") = PYMT_BATCH_LNO

                    .Item("CUST_CODE") = rowARTPYMTL.Item("CUST_CODE")

                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", .Item("CUST_CODE"))
                    .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")

                    .Item("PYMT_NOTE") = rowARTPYMTL.Item("PYMT_NOTE")

                    .Item("PYMT_STATUS") = "0"

                    .Item("CUST_PYMT_REF_NO") = rowARTPYMTL.Item("CUST_PYMT_REF_NO")
                    .Item("CUST_PYMT_AMT") = rowARTPYMTL.Item("CUST_PYMT_AMT")
                    .Item("CUST_PYMT_REF_DATE") = PYMT_BATCH_DATE

                    .Item("CURR_CODE") = CURR_CODE
                    .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE
                    .Item("CUST_PYMT_AMT_CURR") = Val(rowARTPYMT2.Item("CUST_PYMT_AMT") & "")
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID

                End With
                dst.Tables("ARTPYMT2").Rows.Add(rowARTPYMT2)
            Next
        Next

        INIT_LAST("ARTPYMT1", , , True)

        Update_Record_TDA("ARTPYMT1")
        Update_Record_TDA("ARTPYMT2")

        Update_Record_TDA("ARTPYMTC")

        CommitTrans("Update Complete")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "PYMT_BATCH_NO"
                'sql_where = "STATUS = '0'"
        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTPYMTX, "SS", "Show Filter", "Show GroupBox")
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

                Case "grdARTPYMT3"

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

        'If grd.Name = "grdARTPYMT3" Or grd.Name = "grdARTPYMT5" Then
        '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
        '    grd.DisplayLayout.Bands(0).Columns(e.Tool.Key).Hidden = Not tlb_sbt.Checked
        'End If

        If grd.Name = "grdARTCCPA1" Then

            Select Case e.Tool.Key

            End Select
        End If
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)
            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("New", e)
                End If
            Case "PYMT_BATCH_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Edit", e)
                End If
        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            'Case "TERM_CODE"
            '    Calculate_INV_DUE_DATE()
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "PYMT_BATCH_NO"
                Click_Command("Edit")

            Case "INV_NUM"
                Dim INV_NUM As String = ASCMAIN1.CodeSelector.SelectedCode
                Dim INV_TYPE As String = ASCMAIN1.CodeSelector.SelectedRows(0).Item("INV_TYPE")
                Find_Customer_by_Invoice(INV_NUM, INV_TYPE)
        End Select
    End Sub
    Public Overrides Sub txt_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyBase.txt_ValueChanged(sender, e)

        Dim txtctl As UltraWinEditors.UltraTextEditor = DirectCast(sender, UltraWinEditors.UltraTextEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(txtctl)
        Select Case COLUMN_NAME
            Case "BANK_CODE"
                If Absx1.txtFor("BANK_CODE").Text <> "" Then
                    Dim rowGLTBANK1 As DataRow = LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text.ToUpper)
                    If rowGLTBANK1 IsNot Nothing AndAlso rowGLTBANK1.Item("CURR_CODE") & "" <> Absx1.txtFor("CURR_CODE").Text Then
                        Absx1.txtFor("CURR_CODE").Text = rowGLTBANK1.Item("CURR_CODE") & ""

                        Get_CURR_EXCH_RATE()
                    End If
                End If
        End Select
    End Sub

    Public Overrides Sub dte_ValueChanged(sender As Object, e As EventArgs)
        MyBase.dte_ValueChanged(sender, e)
        Dim dtectl As UltraWinEditors.UltraDateTimeEditor = DirectCast(sender, UltraWinEditors.UltraDateTimeEditor)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(dtectl)
        Select Case COLUMN_NAME
            Case "PYMT_BATCH_DATE"
                'If Not ScreenMode Then
                Get_CURR_EXCH_RATE()
                'End If
        End Select
    End Sub
#End Region
    Sub Get_CURR_EXCH_RATE()
        If ScreenMode Then
            Exit Sub
        End If

        Dim CURR_EXCH_RATE As Decimal = TAC.TACMAIN1.Get_CURR_EXCH_RATE(Me, Absx1.txtFor("CURR_CODE").Text, Absx1.dteFor("PYMT_BATCH_DATE").Value)
        Absx1.numFor("CURR_EXCH_RATE").Value = CURR_EXCH_RATE

        'If Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
        '    Absx1.numFor("CURR_EXCH_RATE").Value = 1
        'Else
        '    ASCMAIN1.sql = "Select * from TATCURR3" & vbCrLf _
        '      & " where CURR_CODE = :PARM1" & vbCrLf _
        '      & "   and  CURR_DATE = (Select Max(CURR_DATE) from TATCURR3" & vbCrLf _
        '      & " where CURR_CODE = :PARM2 and CURR_DATE <= :PARM3)"
        '    Dim rowTATCURR3 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVD", New Object() {Absx1.txtFor("CURR_CODE").Text, Absx1.txtFor("CURR_CODE").Text, Absx1.dteFor("PYMT_BATCH_DATE").Value})
        '    If rowTATCURR3 IsNot Nothing Then
        '        Dim X As Integer = 30
        '        Dim CURR_DATE As Date = rowTATCURR3.Item("CURR_DATE")
        '        Dim PYMT_BATCH_DATE As Date = Absx1.dteFor("PYMT_BATCH_DATE").DateTime
        '        Dim DAYSOLD As Integer = PYMT_BATCH_DATE.Subtract(CURR_DATE).Days
        '        Dim CURR_EXCH_RATE As Decimal = Val(rowTATCURR3.Item("CURR_EXCH_RATE") & "")
        '        If DAYSOLD < X Then
        '            Absx1.numFor("CURR_EXCH_RATE").Value = CURR_EXCH_RATE
        '        Else
        '            Absx1.numFor("CURR_EXCH_RATE").Value = 0
        '        End If
        '    Else
        '        Absx1.numFor("CURR_EXCH_RATE").Value = 0
        '    End If
        'End If
    End Sub


#Region "grdARTPYMT2"
    Private Sub grdARTPYMT2_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles grdARTPYMT2.BeforeCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_PYMT_REF_DATE"
                If "" = "" Then
                    Dim currYear As String = Mid(ASCMAIN1.CYP, 1, 4)
                    'grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value = "0"
                End If
        End Select
    End Sub
    Private Sub grdARTPYMT2_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT2.AfterCellUpdate
        Select Case e.Cell.Column.Key
            Case "CUST_CODE"
                If e.Cell.Text = "" Then
                    grdARTPYMT2.ActiveRow.Cells("NON_AR").Value = "1"
                    grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value = Absx1.txtFor("CURR_CODE").Text
                    grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = Absx1.numFor("CURR_EXCH_RATE").Value
                Else
                    grdARTPYMT2.ActiveRow.Cells("NON_AR").Value = "0"
                    grdCodeDesc(grdARTPYMT2, "ARTCUST1", "CUST_CODE", "CUST_NAME")

                    If Absx1.txtFor("CURR_CODE").Text = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                        grdCodeDesc(grdARTPYMT2, "ARTCUST1", "CUST_CODE", "CURR_CODE")
                        If grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                            grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = 1
                        End If
                    Else
                        grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value = Absx1.txtFor("CURR_CODE").Text
                        grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = Absx1.numFor("CURR_EXCH_RATE").Value
                    End If

                    If grdARTPYMT2.ActiveRow.Cells("CUST_NAME").Text = "" Then
                        grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.PrevCell)
                    Else
                        grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                        'grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                    End If
                End If

            Case "CURR_CODE"
                If grdARTPYMT2.ActiveRow.Cells("CURR_CODE").Value & "" = "USD" Then
                    grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value = 1
                    grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                End If

            Case "CURR_EXCH_RATE", "CUST_PYMT_AMT_CURR"
                Dim CUST_PYMT_AMT_CURR As Decimal = Val(grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_AMT_CURR").Value & "")
                Dim CURR_EXCH_RATE As Decimal = Val(grdARTPYMT2.ActiveRow.Cells("CURR_EXCH_RATE").Value & "")
                grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_AMT").Value = CUST_PYMT_AMT_CURR * CURR_EXCH_RATE

            Case "NON_AR"
                If e.Cell.Value & "" = "1" Then
                    grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Value = ""
                    grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    grdARTPYMT2.DisplayLayout.Bands(0).Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
        End Select
    End Sub

    Private Sub grdARTPYMT2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdARTPYMT2.AfterRowActivate
        With grdARTPYMT2.DisplayLayout.Bands(0)
            If grdARTPYMT2.ActiveRow.IsAddRow Then
                .Columns("NON_AR").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                grdARTPYMT2.ActiveCell = grdARTPYMT2.ActiveRow.Cells("CUST_CODE")
                grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.EnterEditMode)
            Else
                .Columns("NON_AR").CellActivation = UltraWinGrid.Activation.NoEdit
                .Columns("CUST_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                If grdARTPYMT2.ActiveRow.Cells("NON_AR").Text = "1" Then
                    .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("CUST_NAME").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End If
        End With
    End Sub

    Private Sub grdARTPYMT2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMT2.BeforeRowUpdate
        With grdARTPYMT2
            If e.Row.Cells("NON_AR").Value = "1" Then
                If e.Row.Cells("CUST_NAME").Text = "" Then
                    MsgBox("You Must Enter a Name for Non-AR Payments", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            Else
                If e.Row.Cells("CUST_CODE").Text = "" Then
                    MsgBox("Missing Value for Customer Code", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                Else
                    LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Text)
                    If cdr Is Nothing Then
                        MsgBox("Invalid Value entered for Customer Code (" & e.Row.Cells("CUST_CODE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                        e.Cancel = True
                    End If
                End If
            End If
            Dim currCode As String = e.Row.Cells("CURR_CODE").Value & ""
            If currCode = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "" Then
                If Val(e.Row.Cells("CURR_EXCH_RATE").Value & "") <> 1 Then
                    MsgBox("Invalid Value entered for Currency Exhange Rate (" & e.Row.Cells("CURR_EXCH_RATE").Text & ")", MsgBoxStyle.OkOnly, "Cannot Update Row")
                    e.Cancel = True
                End If
            End If

            If Not e.Cancel Then
                If e.Row.Cells("PYMT_BATCH_NO").Text = "" Then
                    .ActiveRow.Cells("PYMT_BATCH_NO").Value = Absx1.CtlFor("PYMT_BATCH_NO").Text
                    .ActiveRow.Cells("PYMT_BATCH_LNO").Value = Val(dst.Tables("ARTPYMT2").Compute("Max(PYMT_BATCH_LNO)", "") & "") + 1
                    .ActiveRow.Cells("PYMT_STATUS").Value = "0"
                    If e.Row.Cells("NON_AR").Value = "1" Then
                    Else
                        LookUp("ARTCUST1", e.Row.Cells("CUST_CODE").Value)
                        e.Row.Cells("CUST_NAME").Value = cdr.Item("CUST_NAME")
                    End If

                End If
            End If

        End With
    End Sub

    Private Sub grdARTPYMT2_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdARTPYMT2.ClickCellButton
        Dim sql_where As String = ""
        grdClickCellButton(grdARTPYMT2, sql_where, sql_where <> "")
    End Sub

    Private Sub grdARTPYMT2_Error(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.ErrorEventArgs) Handles grdARTPYMT2.Error
        grdARTPYMT2.ActiveRow.CancelUpdate()
    End Sub

    Private Sub grdARTPYMT2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles grdARTPYMT2.KeyPress
        If grdARTPYMT2.ActiveRow IsNot Nothing Then
            Try
                If grdARTPYMT2.ActiveCell.Column.Key = "CUST_NAME" Then
                    If grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Text <> "" Then
                        e.KeyChar = Chr(0)
                        e.Handled = True
                    End If
                ElseIf grdARTPYMT2.ActiveCell.Column.Key = "CUST_PYMT_REF_DATE" Then

                    If e.KeyChar = "" Then
                        If "" = "" Then
                            Dim year As String = Mid(ASCMAIN1.CYP, 0, 4)
                            grdARTPYMT2.ActiveRow.Cells("CUST_PYMT_REF_DATE").Value = "0"
                        End If
                    End If

                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub grdARTPYMT2_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdARTPYMT2.BeforeExitEditMode
        With grdARTPYMT2.ActiveCell
            Select Case .Column.Key

                Case "CUST_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    End If

                Case "CURR_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                    End If

                Case "CURR_EXCH_RATE"
                    If .Row.Cells("CURR_CODE").Value & "" = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
                        If Val(.Text) <> 1 Then
                            .Value = 1
                        End If
                    End If

                Case "CUST_PYMT_REF_DATE"
                    'Stop
                    Dim DT As String = .EditorResolved.CurrentEditText
                    If grdARTPYMT2.ActiveCell Is Nothing Then
                    Else
                        Try
                            If Len(DT) = 10 And Mid(DT, 7, 4) = "    " And DT <> "  /  /    " Then
                                .Value = Mid(DT, 1, 6) & Now.Year
                            End If
                        Catch ex As Exception

                        End Try
                    End If


            End Select
        End With
    End Sub
#End Region

    Private Sub grdARTPYMTX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTPYMTX.DoubleClickRow
        If grdARTPYMTX.ActiveRow IsNot Nothing AndAlso grdARTPYMTX.ActiveRow.IsDataRow Then
            Absx1.txtFor("PYMT_BATCH_NO").Text = grdARTPYMTX.ActiveRow.Cells("PYMT_BATCH_NO").Text
            Click_Command("Edit")
        End If

    End Sub

    Private Sub chkEditableOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEditableOnly.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        cbeYP.Visible = Not chkEditableOnly.Checked
        Load_ARTPYMTX()
    End Sub

    Private Sub chkMyOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkMyOnly.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ARTPYMTX()
    End Sub

    Private Sub txtInvoice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoice.Enter
        txtInvoice.Text = ""
        txtCUST_CODE.Text = ""
    End Sub

    Private Sub txtInvoice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtInvoice.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Dim INV_NO As String = ASCMAIN1.Format_Field(txtInvoice.Text, "INV_NO")

            If ASCMAIN1.SOLUTION = "SEA" Then
                txtInvoice.Text = INV_NO
                Dim sql As String = "Select CUST_BILL_TO_CUST from SOTINVH1 where ORDR_INV_NO = '" & INV_NO & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow(sql)
                If row Is Nothing Then
                    txtCUST_CODE.Text = ""
                Else
                    txtCUST_CODE.Text = row.Item("CUST_BILL_TO_CUST")
                    grdARTPYMT2.DisplayLayout.Bands(0).AddNew()
                    grdARTPYMT2.ActiveCell = grdARTPYMT2.ActiveRow.Cells("CUST_CODE")
                    grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Value = txtCUST_CODE.Text
                    'grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
                End If
            Else
                INV_NO = INV_NO.PadLeft(10, "0")
                txtInvoice.Text = INV_NO
                Find_Customer_by_Invoice(INV_NO, "I")
            End If
        End If
    End Sub

    Sub Find_Customer_by_Invoice(INV_NO As String, INV_TYPE As String)

        Dim row As DataRow = Nothing
        Dim CUST_CODE As String = ""
        Dim sql As String = ""

        sql = "Select NVL(CUST_BILL_TO_CUST,CUST_CODE) CUST_CODE from SOTINVH1 where INV_NO = '" & INV_NO & "'"
        row = ASCDATA1.GetDataRow(sql)
        If row IsNot Nothing Then
            CUST_CODE = row.Item("CUST_CODE") & ""
        End If
        If row Is Nothing Then
            sql = "Select CUST_CODE from ARTOPEN1" _
                & " where INV_TYPE = '" & INV_TYPE & "' and INV_NUM = '" & INV_NO & "'"
            row = ASCDATA1.GetDataRow(sql)
            If row IsNot Nothing Then
                CUST_CODE = row.Item("CUST_CODE")
            End If
        End If
        txtCUST_CODE.Text = CUST_CODE
        If CUST_CODE <> "" Then
            grdARTPYMT2.DisplayLayout.Bands(0).AddNew()
            grdARTPYMT2.ActiveCell = grdARTPYMT2.ActiveRow.Cells("CUST_CODE")
            grdARTPYMT2.ActiveRow.Cells("CUST_CODE").Value = CUST_CODE
            grdARTPYMT2.PerformAction(UltraWinGrid.UltraGridAction.NextCellByTab)
        End If
    End Sub

    Sub Load_ARTPYMTX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Payment Batch Data")

        Dim OPS_YYYYPP As String = ""

        Dim STATUS As String = "0"
        If Not chkEditableOnly.Checked Then
            STATUS = "%"
            OPS_YYYYPP = cbeYP.Value
            Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", OPS_YYYYPP)
            grdARTPYMTX.Text = "Payments Entered in " & rowGLTPARM2.Item("LEGEND")
        Else
            OPS_YYYYPP = ASCMAIN1.CYP
            grdARTPYMTX.Text = "Payments Entered, awaiting Payment Receipts Journal"
        End If

        Dim USER_ID As String = ASCMAIN1.USER_ID
        If Not chkMyOnly.Checked Then
            USER_ID = "%"
        Else
            grdARTPYMTX.Text &= ", by " & USER_ID
        End If

        Fill_Records("ARTPYMTX", New String() {STATUS, USER_ID, OPS_YYYYPP})
        Sort_grdColumns(grdARTPYMTX, "PYMT_BATCH_NO")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cbeYP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeYP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_ARTPYMTX()
    End Sub
     
    Private Sub grdARTPYMTL_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTPYMTL.AfterCellUpdate
        'If e.Cell.Column.Key = "SKIP" Then
        '    e.Cell.Row.Update()
        'End If
    End Sub

    Private Sub grdARTPYMTL_AfterRowUpdate(sender As Object, e As UltraWinGrid.RowEventArgs) Handles grdARTPYMTL.AfterRowUpdate
        Dim BANK_CODE As String = e.Row.Cells("BANK_CODE").Value & ""
        Dim CUST_NAME_REF As String = e.Row.Cells("CUST_NAME_REF").Value & ""
        Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & ""

        If CUST_CODE <> "" And BANK_CODE <> "" And CUST_NAME_REF <> "" Then
            If dst.Tables("ARTPYMTC").Rows.Find(New String() {BANK_CODE, CUST_NAME_REF}) Is Nothing Then
                dst.Tables("ARTPYMTC").Rows.Add(New String() {BANK_CODE, CUST_NAME_REF, CUST_CODE})
                For Each row As DataRow In dst.Tables("ARTPYMTL").Select("BANK_CODE = '" & BANK_CODE & "' and CUST_NAME_REF = '" & CUST_NAME_REF & "' and CUST_CODE IS NULL")
                    row.Item("CUST_CODE") = CUST_CODE
                Next
            End If
        End If
    End Sub

    Private Sub grdARTPYMTL_BeforeRowUpdate(sender As Object, e As UltraWinGrid.CancelableRowEventArgs) Handles grdARTPYMTL.BeforeRowUpdate
        Dim CUST_CODE As String = e.Row.Cells("CUST_CODE").Value & ""
        If CUST_CODE <> "" Then
            If LookUp("ARTCUST1", CUST_CODE) Is Nothing Then
                e.Cancel = True
            End If
        End If

    End Sub

    Private Sub grdARTPYMTL_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdARTPYMTL.ClickCellButton
        If e.Cell.Column.Key = "CLONE" Then
 
            Dim rowARTPYMTL As DataRow = dst.Tables("ARTPYMTL").NewRow
                For Each DCOL As DataColumn In dst.Tables("ARTPYMTL").Columns
                rowARTPYMTL.Item(DCOL.ColumnName) = e.Cell.Row.Cells(DCOL.ColumnName).Value
            Next

            'rowARTPYMTL.Item("CLONED") = "1"
            dst.Tables("ARTPYMTL").Rows.Add(rowARTPYMTL)

            grdARTPYMTL.Rows.Refresh(UltraWinGrid.RefreshRow.FireInitializeRow)
            grdARTPYMTL.ActiveRowScrollRegion.Scroll(UltraWinGrid.RowScrollAction.Bottom)
            grdARTPYMTL.ActiveRow = grdARTPYMTL.Rows.Last
            grdARTPYMTL.ActiveRow.Selected = True
        ElseIf e.Cell.Column.Key = "CUST_CODE" Then
            Dim sql_where As String = ""
            grdClickCellButton(grdARTPYMTL, sql_where, sql_where <> "")
        End If
    End Sub

    Private Sub grdARTPYMTL_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdARTPYMTL.InitializeRow
        'If e.Row.Cells("CLONED").Value & "" = "1" Then
        '    e.Row.Cells("SKIP").Appearance.BackColor = Drawing.Color.Green
        'Else
        '    e.Row.Cells("SKIP").Appearance.BackColor = Drawing.Color.Empty
        'End If
    End Sub
End Class