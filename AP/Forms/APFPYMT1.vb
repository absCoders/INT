Public Class APFPYMT1
    Dim rowAPTVEND1 As DataRow
    Dim rowGLTBANK1 As DataRow
    Dim rowAPTPYMT1 As DataRow
    Dim APTINVH1 As String
    Dim CHECK_NUM_ctr As Integer = 0

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst

            ASCMAIN1.sql = "Select VEND_CODE, VEND_NAME, VEND_CLASS_CODE, VEND_PYMT_CYCLE, VEND_ON_HOLD from APTVEND1"
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False)
            With .Tables("APTVEND1")
                .Columns.Add("UNS_ITEMS", GetType(System.Int32))
                .Columns.Add("UNS_PYMT", GetType(System.Double))
                .Columns.Add("UNS_DISC", GetType(System.Double))
                .Columns.Add("SEL_ITEMS", GetType(System.Int32))
                .Columns.Add("SEL_PYMT", GetType(System.Double))
                .Columns.Add("SEL_DISC", GetType(System.Double))
            End With

            Create_TDA(.Tables.Add, "APTINVH1", "*", _
            Update_COLUMN_NAMEs:="VEND_ALT_CODE,CHECK_NUM,VEND_CODE_AP,BATCH_NO_PYMT,BATCH_PYMT,BATCH_DISC,INV_STATUS,INV_PAYMENTS,INV_DISC_TAKEN,INV_LAST_PMT_DATE,INV_BALANCE,CHECK_DATE")

            ASCMAIN1.sql = "Select APTPYMT1.* from APTPYMT1"
            Create_TDA(.Tables.Add, "APTPYMTX", "**", 0, False)
            With .Tables("APTPYMTX")
                .Columns.Add("CHECKS", GetType(System.Int32))
                .Columns.Add("BATCH_PYMT", GetType(System.Double))
                .Columns.Add("BATCH_DISC", GetType(System.Double))
            End With

            Create_TDA(.Tables.Add, "APTCHCK1", "*")
            Create_TDA(.Tables.Add, "APTCHCK2", "*")
            Create_TDA(.Tables.Add, "APTVEND5", "*")

            Create_TDA(.Tables.Add, "APTPYMT1", "*")
            Create_TDA(.Tables.Add, "APTPYMT2", "*", 1)

            .Tables.Add("APTPYMT0")
            With .Tables("APTPYMT0")
                .Columns.Add("CHECK_TYPE", GetType(System.String))
                .Columns.Add("CHECK_TYPE_DESC", GetType(System.String))
                .Columns.Add("CHECK_TYPE_COUNT", GetType(System.Int32))
                .Columns.Add("CHECK_TYPE_AMOUNT", GetType(System.Double))
            End With

            .Tables.Add("APTPYMTC")
            With .Tables("APTPYMTC").Columns
                .Add("VEND_CODE", GetType(System.String))
                .Add("VEND_CODE_AP", GetType(System.String))
                .Add("VEND_ALT_CODE", GetType(System.String))
                .Add("VOUCHER_NO", GetType(System.String))
                .Add("BATCH_PYMT", GetType(System.Double))
                .Add("BATCH_DISC", GetType(System.Double))
            End With
            With .Tables("APTPYMTC")
                .PrimaryKey = New DataColumn() { _
                .Columns("VEND_CODE"), _
                .Columns("VEND_CODE_AP"), _
                .Columns("VEND_ALT_CODE"), _
                .Columns("VOUCHER_NO")}
            End With

            .Relations.Add("APTINVH1" _
            , .Tables("APTVEND1").Columns("VEND_CODE") _
            , .Tables("APTINVH1").Columns("VEND_CODE"))


            ASCMAIN1.sql = "Select VOUCHER_NO, INV_TYPE, INV_NUM, INV_DATE" _
            & ", INV_BALANCE, INV_REF, PO_ORDER_NO, INV_DUE_DATE" _
            & ", BATCH_NO_PYMT, BATCH_PYMT, BATCH_DISC, INV_PYMT_METHOD" _
            & " from APTINVH1"
            Create_TDA(.Tables.Add, "APTINVH1_DUE_TO", "**", 0, False)
            With .Tables("APTINVH1_DUE_TO")
                .Columns.Add("INV_BALANCE_ABS", GetType(System.Decimal))
                .Columns.Add("MATCH", GetType(System.String))
                .Columns.Add("MATCH_NO", GetType(System.Int32))
            End With


            Create_TDA(dst.Tables.Add, "GLTBANK1", "*")

        End With

        Create_Lookup("TATTERM1")
        Create_Lookup("GLTBANK1")
        Create_Lookup("APTVEND1")
        Create_Lookup("APTCHCK1")
        Create_Lookup("APTVEND2")

        grdAPTINVH1.DataSource = dst.Tables("APTINVH1")
        grdAPTINVH1.Visible = False
        grdAPTINVH1_SEL.DataSource = New DataView(dst.Tables("APTINVH1"), "BATCH_NO_PYMT is Not Null", "VEND_CODE", DataViewRowState.CurrentRows)
        grdAPTINVH1_UNS.DataSource = New DataView(dst.Tables("APTINVH1"), "BATCH_NO_PYMT is Null", "VEND_CODE", DataViewRowState.CurrentRows)
        grdAPTINVH1_DTL.DataSource = dst.Tables("APTINVH1")
        grdAPTINVH1_DTL.Visible = False
        grdAPTVEND1.DataSource = dst.Tables("APTVEND1")
        grdAPTPYMT0.DataSource = dst.Tables("APTPYMT0")
        grdAPTPYMTX.DataSource = dst.Tables("APTPYMTX")
        grdAPTPYMT2.DataSource = dst.Tables("APTPYMT2")
        grdAPTINVH1_DUE_TO.DataSource = dst.Tables("APTINVH1_DUE_TO")


        For Each col As UltraWinGrid.UltraGridColumn In grdAPTVEND1.DisplayLayout.Bands("APTVEND1").Columns
            If col.Key Like "UNS_*" Then
            ElseIf col.Key Like "SEL_*" Then
                col.CellAppearance.BackColor = Drawing.Color.LightGreen
            Else
                col.CellAppearance.BackColor = Drawing.Color.LightGray
            End If
        Next

        grdAPTINVH1.DisplayLayout.GroupByBox.Hidden = True
        grdAPTINVH1_DTL.DisplayLayout.GroupByBox.Hidden = True

        With grdAPTVEND1.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("VEND_CODE", False)
        End With

        With grdAPTPYMT0.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("CHECK_TYPE", False)
        End With
        grdAPTPYMT0.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select


        With grdAPTINVH1_DUE_TO.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("INV_BALANCE_ABS", True)
            .SortedColumns.Add("INV_DATE", True)
        End With
        'grdAPTINVH1_DUE_TO.DisplayLayout.Override.HeaderClickAction = UltraWinGrid.HeaderClickAction.Select

        With grdAPTPYMT2.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("CHECK_NUM", False)
        End With

        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")

        Call Create_Summary(grdAPTINVH1_SEL, "VOUCHER_NO", "Count")
        Call Create_Summary(grdAPTINVH1_SEL, "INV_BALANCE")
        Call Create_Summary(grdAPTINVH1_SEL, "BATCH_PYMT")
        Call Create_Summary(grdAPTINVH1_SEL, "BATCH_DISC")

        Call Create_Summary(grdAPTINVH1_UNS, "VOUCHER_NO", "Count")
        Call Create_Summary(grdAPTINVH1_UNS, "INV_BALANCE")
        Call Create_Summary(grdAPTINVH1_UNS, "BATCH_PYMT")
        Call Create_Summary(grdAPTINVH1_UNS, "BATCH_DISC")
        Call Create_Summary(grdAPTINVH1_UNS, "INV_AMT")
        Call Create_Summary(grdAPTINVH1_UNS, "INV_DISC_AMT")

        Call Create_Summary(grdAPTINVH1, "VOUCHER_NO", "Count")
        Call Create_Summary(grdAPTINVH1, "INV_BALANCE")
        Call Create_Summary(grdAPTINVH1, "BATCH_PYMT")
        Call Create_Summary(grdAPTINVH1, "BATCH_DISC")
        Call Create_Summary(grdAPTINVH1, "INV_AMT")
        Call Create_Summary(grdAPTINVH1, "INV_DISC_AMT")


        Call Create_Summary(grdAPTINVH1_DTL, "VOUCHER_NO", "Count")
        Call Create_Summary(grdAPTINVH1_DTL, "INV_BALANCE")
        Call Create_Summary(grdAPTINVH1_DTL, "BATCH_PYMT")
        Call Create_Summary(grdAPTINVH1_DTL, "BATCH_DISC")
        Call Create_Summary(grdAPTINVH1_DTL, "INV_AMT")
        Call Create_Summary(grdAPTINVH1_DTL, "INV_DISC_AMT")

        Call Create_Summary(grdAPTVEND1, "VEND_CODE", "Count")
        Call Create_Summary(grdAPTVEND1, "UNS_ITEMS")
        Call Create_Summary(grdAPTVEND1, "UNS_PYMT")
        Call Create_Summary(grdAPTVEND1, "UNS_DISC")
        Call Create_Summary(grdAPTVEND1, "SEL_ITEMS")
        Call Create_Summary(grdAPTVEND1, "SEL_PYMT")
        Call Create_Summary(grdAPTVEND1, "SEL_DISC")

        Call Create_Summary(grdAPTINVH1_DUE_TO, "VOUCHER_NO", "Count")
        Call Create_Summary(grdAPTINVH1_DUE_TO, "INV_BALANCE")


        Call Create_Summary(grdAPTPYMT2, "CHECK_NUM", "Count")
        Call Create_Summary(grdAPTPYMT2, "BATCH_PYMT")
        Call Create_Summary(grdAPTPYMT2, "BATCH_DISC")

        Call Create_Summary(grdAPTPYMTX, "BATCH_NO_PYMT", "Count")
        Call Create_Summary(grdAPTPYMTX, "CHECKS")
        Call Create_Summary(grdAPTPYMTX, "BATCH_PYMT")
        Call Create_Summary(grdAPTPYMTX, "BATCH_DISC")
        Call Bind_Controls(Me, "APTPYMT1")

        grdAPTINVH1_SEL.DisplayLayout.Bands(0).SummaryFooterCaption = "Batch Totals"
        grdAPTINVH1_UNS.DisplayLayout.Bands(0).SummaryFooterCaption = "Total Unselected Items"
        grdAPTVEND1.DisplayLayout.Bands(0).SummaryFooterCaption = "Total All Vendors"

        Call Load_Drop_Down("VEND_TYPE")
        Call Load_Drop_Down("VEND_CLASS_CODE")
        Call Load_Drop_Down("VEND_PYMT_CYCLE")
        Call Load_Drop_Down("INV_PYMT_CYCLE")

        grdAPTVEND1.DisplayLayout.Bands(0).Override.GroupByRowDescriptionMask = "[caption] : [value] ([count] [count,Vendors,Vendor,Vendors])"
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                Validate_Code("BANK_CODE")

                If Absx1.dteFor("CHECK_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Check Date Required"
                End If

                Validate_Code("PYMT_METHOD")

                If Absx1.txtFor("VEND_CODE").Text <> "" Then
                    Validate_Code("VEND_CODE")
                End If


                If EMsg = "" Then
                    If Not TAC.APCMAIN1.check_Bank_Payment_Method(Me, Absx1.txtFor("BANK_CODE").Text, Absx1.txtFor("PYMT_METHOD").Text) Then
                        EMsg &= vbCr & $"Bank {Absx1.txtFor("BANK_CODE").Text} does NOT support Payment Method { Absx1.txtFor("PYMT_METHOD").Text}"
                    End If
                End If

                If Absx1.txtFor("CHECK_NUM").Text <> "" Then
                    If Absx1.txtFor("VEND_CODE").Text = "" Then
                        EMsg &= vbCr & "A Single Vendor is Required for Manual Check Entry"
                    End If
                    If EMsg = "" Then
                        If MsgBox("By entering a Check Number Manually, you are indicating that this payment should be recorded without Printing a Check, generating an authorization email, or creating a payment file for transmission to the bank." & vbCr & vbCr & "Upon clicking 'Update' (After selecting AP items to be included on this payment), this payment will be recorded permanenty, and will appear on the next Check Register." & vbCr & vbCr & "Proceed with this entry?", MsgBoxStyle.YesNo, "You are entering a Manual Check") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Edit"

                If Validate_Code("BATCH_NO_PYMT") Then
                    If LookUp("GLTBANK1", cdr.Item("BANK_CODE")).Item("BATCH_NO_PYMT") & "" = Absx1.txtFor("BATCH_NO_PYMT").Text Then
                        If vbCancel = MsgBox("No Changes will be Permitted; You may only Delete this Batch", MsgBoxStyle.OkOnly, "Checks were printed for Batch No " & Absx1.txtFor("BATCH_NO_PYMT").Text) Then
                            Exit Sub
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock(TABLE_NAME, Absx1.txtFor("BATCH_NO_PYMT").Text) Then
                        Exit Sub
                    End If
                End If

            Case "Update"

                If dst.Tables("APTPYMT2").Select("ISNULL(BATCH_PYMT,0) - ISNULL(BATCH_DISC,0) < 0 AND VOUCHER_NO <> '0000000000'").Length > 0 Then
                    If MsgBox("Warning - Some Payments have a Negative Total." _
                              & vbCrLf & "These may be Credits inadvertantly set up for Payment via Separate Checks" _
                              & vbCrLf & vbCrLf & "Do you want these Negative Payments Removed automatically?", _
                              vbYesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If

                End If

                If Not TAC.APCMAIN1.check_Bank_Payment_Method(Me, Absx1.txtFor("BANK_CODE").Text, Absx1.txtFor("PYMT_METHOD").Text) Then
                    EMsg &= vbCr & $"Bank {Absx1.txtFor("BANK_CODE").Text} does NOT support Payment Method { Absx1.txtFor("PYMT_METHOD").Text}"
                End If

                If ROWs("APTPARM1").Item("AP_PARM_BANK_METHOD") & "" = "1 " Then

                    If Absx1.txtFor("PYMT_METHOD").Text = "ACH" Or Absx1.txtFor("PYMT_METHOD").Text = "WIRE" Then
                        If dst.Tables("APTINVH1").Select("BATCH_NO_PYMT IS NOT NULL AND ISNULL(VEND_BANK_COUNTRY,'') = ''").Length > 0 Then
                            EMsg &= vbCr & "The Vendor's Bank Specifications are Required for ACH and WIRE Payment Methods"
                        Else

                            Dim PYMTS_ZERO_OR_LESS As Boolean = (dst.Tables("APTPYMT2").Select("ISNULL(BATCH_PYMT,0) - ISNULL(BATCH_DISC,0) < 0").Length > 0)
                            If PYMTS_ZERO_OR_LESS Then
                                EMsg &= vbCr & "Cannot Send ACH or WIRE Payments <= $0"
                            End If

                            If Absx1.txtFor("PYMT_METHOD").Text = "ACH" Then
                                Dim ACH_PYMTS_UNDER_1 As Boolean = (dst.Tables("APTPYMT2").Select("ISNULL(BATCH_PYMT,0) - ISNULL(BATCH_DISC,0) < 1").Length > 0)
                                Dim ACH_PYMTS_ATLEAST_1 As Boolean = (dst.Tables("APTPYMT2").Select("ISNULL(BATCH_PYMT,0) - ISNULL(BATCH_DISC,0) >= 1").Length > 0)

                                If ACH_PYMTS_UNDER_1 And ACH_PYMTS_ATLEAST_1 Then
                                    EMsg &= vbCr & "Cannot Mix ACH Payments < $1 with ACH Payments >= $1"
                                End If
                            End If


                            Dim VEND_BANK_COUNTRY_is_USA As Boolean = False
                            Dim VEND_BANK_COUNTRY_is_NON_USA As Boolean = False
                            For Each rowVEND_BANK_COUNTRY As DataRow In
                                ASCDATA1.SelectDistinct(dst.Tables("APTINVH1").Select("BATCH_NO_PYMT IS NOT NULL"), New String() {"VEND_BANK_COUNTRY"}).Select("")
                                Dim VEND_BANK_COUNTRY As String = rowVEND_BANK_COUNTRY.Item("VEND_BANK_COUNTRY")
                                If VEND_BANK_COUNTRY = "USA" Then
                                    VEND_BANK_COUNTRY_is_USA = True
                                Else
                                    VEND_BANK_COUNTRY_is_NON_USA = True
                                End If
                            Next

                            If VEND_BANK_COUNTRY_is_USA And VEND_BANK_COUNTRY_is_NON_USA Then
                                EMsg &= vbCr & "Cannot Mix US Banks with non-US Banks in a single Payment Selection Batch"
                            End If



                        End If
                    End If
                End If

                Remove_0s()
                Rebuild_Check_File()

                If Absx1.dteFor("CHECK_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "Check Date Required"
                Else
                    Dim CHECK_DATE As Date = Absx1.dteFor("CHECK_DATE").Value
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", ASCMAIN1.CYP)
                    Dim PRD_END_DATE As Date = rowGLTPARM2.Item("PRD_END_DATE")
                    Dim rowGLTPARM2_prior As DataRow = LookUp("GLTPARM2", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1))
                    Dim PRD_END_DATE_prior As Date = rowGLTPARM2_prior.Item("PRD_END_DATE")

                    If Format(CHECK_DATE, "yyyyMMdd") <= Format(PRD_END_DATE_prior, "yyyyMMdd") _
                    Or Format(CHECK_DATE, "yyyyMMdd") > Format(PRD_END_DATE, "yyyyMMdd") Then
                        EMsg &= vbCr & "Check Date Must be between " & Format(PRD_END_DATE_prior.AddDays(1), "MM/dd/yyyy") & " and " & Format(PRD_END_DATE, "MM/dd/yyyy")
                    End If
                End If

                If rowGLTBANK1.Item("BATCH_NO_PYMT") & "" = HFs("BATCH_NO_PYMT") Then
                    EMsg &= vbCr & "Checks were printed for Batch No " & HFs("BATCH_NO_PYMT") & "; No Changes Permitted"
                End If

                If dst.Tables("APTPYMT2").Select("", "", DataViewRowState.CurrentRows).Length = 0 Then
                    EMsg &= vbCr & "Nothing Selected"
                End If

                If Absx1.txtFor("CHECK_NUM").Text & "" <> "" Then
                    Dim rowAPTCHCK1 As DataRow = LookUp("APTCHCK1", New String() {Absx1.txtFor("BANK_CODE").Text, Absx1.txtFor("CHECK_NUM").Text})
                    If rowAPTCHCK1 IsNot Nothing Then
                        EMsg &= vbCr & "Check No " & Absx1.txtFor("CHECK_NUM").Text & " has already been Posted"
                    End If

                    If EMsg = "" Then
                        ' PAYEENAME IS WRONG
                        Call Refresh_APTVEND1(HFs("VEND_CODE"))
                        Dim CHECK_AMT As Double = Val(dst.Tables("APTVEND1").Compute("SUM(SEL_PYMT)", "VEND_CODE = '" & HFs("VEND_CODE") & "'") & "")
                        If MsgBox("Please Verify the Following Information:" & vbCr & vbCr & "Check No: " & Absx1.txtFor("CHECK_NUM").Text & ", " & Absx1.dteFor("CHECK_DATE").Value & vbCr & "Bank: " & Absx1.txtFor("BANK_DESC").Text & vbCr & "Payee: " & Absx1.txtFor("VEND_NAME").Text & vbCr & "Amount: " & Format(CHECK_AMT, "$###,##0.00") & vbCr & vbCr & "OK To Continue with Update?", vbQuestion + vbYesNo, "Verification: You are about to Record a Payment") = vbNo Then
                            Exit Sub
                        End If
                    End If

                End If

            Case "Select Selected"

                If grdAPTINVH1_UNS.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "Nothing Selected"
                End If

            Case "Reverse Selected"

                If grdAPTINVH1_SEL.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "Nothing Selected"
                End If


            Case "Check(s)"

                If grdAPTPYMT2.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "Nothing Selected"
                End If

            Case "Item(s)"

                If grdAPTINVH1_DTL.Selected.Rows.Count = 0 Then
                    EMsg &= vbCr & "Nothing Selected"
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

            Case "New"
                EntryMode = "N"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Edit"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Update"
                ' THE NEXT 2 STEPS WERE MOVED TO Proceed_PreReq
                'Call Remove_0s()
                'Call Rebuild_Check_File()
                Call Update_Record()
                Call Print_Report()
                Call Mode_Settings(False)

            Case "Cancel"
                Call Mode_Settings(False)

            Case "Delete"
                Call Delete_Record()
                Call Mode_Settings(False)

            Case "Rebuild Check File"
                Call Rebuild_Check_File()

            Case "Excel"
                Select Case UltraTabControl1.ActiveTab.Key
                    Case "Unselected Items"
                        Call Export_to_Excel(grdAPTINVH1_UNS)

                    Case "Select by Vendor"
                        Call Export_to_Excel(New UltraWinGrid.UltraGrid() {grdAPTVEND1, grdAPTINVH1})

                        'Dim g As New UltraWinGrid.UltraGrid
                        'UltraGroupBox1.Controls.Add(g)
                        'Me.Controls.Add(g)
                        'g.DataSource = dst.Tables("APTVEND1")
                        ''g.DisplayLayout.CopyFrom(grdAPTVEND1.DisplayLayout)
                        'g.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
                        'g.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
                        'g.DisplayLayout.BandsSerializer.Add(New Infragistics.Win.UltraWinGrid.UltraGridBand("APTINVH1", -1))
                        'Call Export_to_Excel(g)
                        'Me.Controls.Remove(g)

                    Case "Selected Items"
                        Call Export_to_Excel(grdAPTINVH1_SEL)

                    Case "Payments"
                        Call Export_to_Excel(grdAPTPYMT2)
                End Select


            Case "Reverse All"
                Call ASCMAIN1.Progress("Now Reversing All Selections")
                Me.Cursor = Cursors.WaitCursor
                For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("BATCH_NO_PYMT is Not Null")
                    rowAPTINVH1.Item("BATCH_NO_PYMT") = Null
                    rowAPTINVH1.Item("CHECK_NUM") = Null
                Next
                dst.Tables("APTPYMT2").Rows.Clear()
                Call Refresh_APTVEND1()
                Me.Cursor = Cursors.Default
                Call ASCMAIN1.Progress("")


            Case "Reverse Selected"
                Call ASCMAIN1.Progress("Now Reversing Selected Items")
                Me.Cursor = Cursors.WaitCursor

                'For Each grdrow As UltraWinGrid.UltraGridRow In grdAPTINVH1_SEL.Selected.Rows
                '    grdrow.Cells("BATCH_NO_PYMT").Value = Null
                '    grdrow.Cells("CHECK_NUM").Value = Null
                '    grdrow.Update()
                'Next

                Dim VEND_CODEs As New List(Of String)
                Dim VOUCHER_NOs As New List(Of String)
                For Each grdrow As UltraWinGrid.UltraGridRow In grdAPTINVH1_SEL.Selected.Rows
                    Dim VEND_CODE As String = grdrow.Cells("VEND_CODE").Text
                    If Not VEND_CODEs.Contains(VEND_CODE) Then
                        VEND_CODEs.Add(VEND_CODE)
                    End If
                    VOUCHER_NOs.Add(grdrow.Cells("VOUCHER_NO").Text)
                Next

                For Each VOUCHER_NO As String In VOUCHER_NOs
                    Dim rowAPTINVH1 As DataRow = dst.Tables("APTINVH1").Rows.Find(VOUCHER_NO)
                    rowAPTINVH1.Item("BATCH_NO_PYMT") = Null
                    rowAPTINVH1.Item("CHECK_NUM") = Null
                Next

                For Each VEND_CODE As String In VEND_CODEs
                    Call Build_Check_File(False, VEND_CODE)
                    'Call Refresh_APTVEND1()
                Next

                grdAPTINVH1_SEL.Selected.Rows.Clear()
                Call Refresh_APTVEND1()
                Me.Cursor = Cursors.Default
                Call ASCMAIN1.Progress("")

            Case "Select All"
                Call Select_for_Payment("")

            Case "Select Selected"
                Dim VEND_CODEs As New List(Of String)
                Dim VOUCHER_NOs As New List(Of String)
                Call ASCMAIN1.Progress("Now Selecting for Payment")
                Me.Cursor = Cursors.WaitCursor
                For Each grdrow As UltraWinGrid.UltraGridRow In grdAPTINVH1_UNS.Selected.Rows
                    grdAPTINVH1_UNS.ActiveRow = grdrow
                    Dim VEND_CODE As String = grdrow.Cells("VEND_CODE").Text
                    If Not VEND_CODEs.Contains(VEND_CODE) Then
                        VEND_CODEs.Add(VEND_CODE)
                    End If
                    VOUCHER_NOs.Add(grdrow.Cells("VOUCHER_NO").Text)
                    'grdrow.Cells("BATCH_NO_PYMT").Value = HFs("BATCH_NO_PYMT")
                    'grdrow.Update()
                Next
                grdAPTINVH1_UNS.Selected.Rows.Clear()

                For Each VOUCHER_NO As String In VOUCHER_NOs
                    Dim rowAPTINVH1 As DataRow = dst.Tables("APTINVH1").Rows.Find(New Object() {VOUCHER_NO})
                    rowAPTINVH1.Item("BATCH_NO_PYMT") = HFs("BATCH_NO_PYMT")
                Next

                For Each VEND_CODE As String In VEND_CODEs
                    Call Build_Check_File(False, VEND_CODE)
                    'Call Refresh_APTVEND1()
                Next
                Me.Cursor = Cursors.Default
                Call ASCMAIN1.Progress("")

            Case "Check(s)"
                Call ASCMAIN1.Progress("Now Reversing Selected Checks")
                Me.Cursor = Cursors.WaitCursor
                Dim VEND_CODE As String = grdAPTPYMT2.ActiveRow.Cells("VEND_CODE").Text
                For Each grdrow As UltraWinGrid.UltraGridRow In grdAPTPYMT2.Selected.Rows
                    Dim CHECK_NUM As String = grdrow.Cells("CHECK_NUM").Text
                    For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("CHECK_NUM = '" & CHECK_NUM & "'", "")
                        rowAPTINVH1.Item("BATCH_NO_PYMT") = Null
                        rowAPTINVH1.Item("CHECK_NUM") = Null
                    Next
                    dst.Tables("APTPYMT2").Rows(grdrow.ListIndex).Delete()
                    'grdrow.Delete()
                Next
                'grdAPTINVH1_SEL.Selected.Rows.Clear()

                grdAPTINVH1_DTL.Visible = False
                Call Build_Check_File(False, VEND_CODE)
                'Call Refresh_APTVEND1(VEND_CODE)
                Me.Cursor = Cursors.Default
                Call ASCMAIN1.Progress("")

            Case "Item(s)"
                Call ASCMAIN1.Progress("Now Reversing Selected Items")
                Me.Cursor = Cursors.WaitCursor
                For Each grdrow As UltraWinGrid.UltraGridRow In grdAPTINVH1_DTL.Selected.Rows
                    grdrow.Cells("BATCH_NO_PYMT").Value = Null
                    grdrow.Cells("CHECK_NUM").Value = Null
                    grdrow.Update()
                Next
                'grdAPTINVH1_SEL.Selected.Rows.Clear()
                Dim CHECK_NUM As String = grdAPTPYMT2.ActiveRow.Cells("CHECK_NUM").Text
                Dim VEND_CODE As String = grdAPTPYMT2.ActiveRow.Cells("VEND_CODE").Text
                Call Build_Check_File(False, VEND_CODE, CHECK_NUM)
                'Call Refresh_APTVEND1(VEND_CODE)

                grdAPTPYMT2.ActiveRow = grdAPTPYMT2.Rows.GetRowWithListIndex(dst.Tables("APTPYMT2").Rows.IndexOf(dst.Tables("APTPYMT2").Rows.Find(New Object() {HFs("BATCH_NO_PYMT"), CHECK_NUM})))

                'If grdAPTVEND1.Rows.Count > 0 Then
                '    If VEND_CODE <> "" Then
                '        grdAPTVEND1.ActiveRow = grdAPTVEND1.Rows.GetRowWithListIndex(dst.Tables("APTVEND1").Rows.IndexOf(dst.Tables("APTVEND1").Rows.Find(New Object() {VEND_CODE})))
                '    Else
                '        grdAPTVEND1.ActiveRow = grdAPTVEND1.Rows(0)
                '    End If
                'End If


                Me.Cursor = Cursors.Default
                Call ASCMAIN1.Progress("")

            Case "Remove if 0"
                Call Remove_0s()
                Call Refresh_APTVEND1()

            Case "Print Report"
                Call Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("New").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Edit").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Delete").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Excel").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Remove if 0").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Print Report").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Rebuild Check File").Settings.Enabled = iScreenMode

                .Groups("AP Item Filters").Visible = Not tf

                If Not tf Then
                    .Groups("Filter Vendors By").Visible = True
                    .Groups("Filter AP Items By").Visible = False
                    .Groups("Select by Date").Visible = False
                    .Groups("Selection Tools").Visible = False
                    .Groups("Reversal Tools").Visible = False
                    .Groups("Check Statistics").Visible = False
                    .Groups("Remove Selected ...").Visible = False
                    .Groups("Match Invoices").Visible = False
                    .Groups("Show Items").Visible = False
                End If
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = tf
        grdAPTPYMTX.Visible = Not tf

        With grdAPTINVH1.DisplayLayout.Bands("APTINVH1")
            .Columns("SEL").Header.Fixed = True
            .Columns("REV").Header.Fixed = True
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VOUCHER_NO").Header.Fixed = True
        End With

        With grdAPTINVH1_SEL.DisplayLayout.Bands("APTINVH1")
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VOUCHER_NO").Header.Fixed = True
        End With

        With grdAPTINVH1_UNS.DisplayLayout.Bands("APTINVH1")
            .Columns("VEND_CODE").Header.Fixed = True
            .Columns("VOUCHER_NO").Header.Fixed = True
        End With

        With grdAPTPYMTX.DisplayLayout.Bands("APTPYMTX")
            .SortedColumns.Clear()
            .SortedColumns.Add("BATCH_NO_PYMT", False)
        End With

        'With grdAPTVEND1.DisplayLayout.Bands("APTVEND1")
        '    .Columns("VEND_CODE").Header.Fixed = True
        '    .Columns("VOUCHER_NO").Header.Fixed = True
        'End With

        If ScreenMode Then
            Absx1.dteFor("CHECK_DATE").ReadOnly = False
            If Absx1.txtFor("VEND_CODE").Text <> "" Then
                Absx1.txtFor("CHECK_NUM").ReadOnly = False
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        dst.EnforceConstraints = False
        dst.Tables("APTPYMT1").Rows.Clear()
        dst.Tables("APTPYMT2").Rows.Clear()
        dst.Tables("APTINVH1").Rows.Clear()
        dst.Tables("APTVEND1").Rows.Clear()

        dst.Tables("APTINVH1_DUE_TO").Rows.Clear()

        dst.EnforceConstraints = True

        Absx1.cmbFor("VEND_TYPE").Text = ""
        Absx1.cmbFor("VEND_CLASS_CODE").Text = ""
        Absx1.cmbFor("VEND_PYMT_CYCLE").Text = ""
        Absx1.cmbFor("INV_PYMT_CYCLE").Text = ""

        Absx1.txtFor("BANK_CODE").Text = ROWs("APTPARM1").Item("AP_PARM_BANK_CODE") & ""

        Dim sql As String = "Select APTPYMT1.*, X.CHECKS, X.BATCH_PYMT, X.BATCH_DISC from APTPYMT1, (SELECT BATCH_NO_PYMT, COUNT (*) CHECKS, SUM (BATCH_PYMT) BATCH_PYMT, SUM (BATCH_DISC) BATCH_DISC from APTPYMT2 group by BATCH_NO_PYMT) X where X.BATCH_NO_PYMT = APTPYMT1.BATCH_NO_PYMT"
        Call Fill_Records("APTPYMTX", "", True, sql)
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading Open AP Items ...")
        'Application.DoEvents()

        Call Save_Header_Fields(UltraGroupBox1)
        Call Save_Header_Fields(grpFilterVendorsBy, False)
        Call Save_Header_Fields(grpLoadAPItems, False)

        If EntryMode = "N" Then
            HFs("BATCH_NO_PYMT") = ASCMAIN1.Next_Control_No("APTPYMT1.BATCH_NO_PYMT")
        End If

        rowAPTPYMT1 = Fill_Record("APTPYMT1", New String() {HFs("BATCH_NO_PYMT")}, EntryMode = "N")
        If EntryMode = "E" Then
            HFs("BANK_CODE") = rowAPTPYMT1.Item("BANK_CODE")
            HFs("CHECK_DATE") = rowAPTPYMT1.Item("CHECK_DATE")
            'HFs("VEND_NAME") = rowAPTINVH1.Item("VEND_NAME")
            HFs("BATCH_NO_PYMT") = rowAPTPYMT1.Item("BATCH_NO_PYMT")
        End If

        rowGLTBANK1 = LookUp("GLTBANK1", HFs("BANK_CODE"))

        If EntryMode = "N" Then
            rowAPTPYMT1.Item("BATCH_NO_PYMT") = HFs("BATCH_NO_PYMT")
            rowAPTPYMT1.Item("BANK_CODE") = HFs("BANK_CODE")
            rowAPTPYMT1.Item("CHECK_DATE") = HFs("CHECK_DATE")
            rowAPTPYMT1.Item("PYMT_ONLY_SEL_BANK") = HFs("PYMT_ONLY_SEL_BANK")
            rowAPTPYMT1.Item("PYMT_ONLY_SEL_METHOD") = HFs("PYMT_ONLY_SEL_METHOD")
            rowAPTPYMT1.Item("VEND_TYPE") = HFs("VEND_TYPE")
            rowAPTPYMT1.Item("VEND_CLASS_CODE") = HFs("VEND_CLASS_CODE")
            rowAPTPYMT1.Item("VEND_PYMT_CYCLE") = HFs("VEND_PYMT_CYCLE")
            rowAPTPYMT1.Item("CHECK_NUM") = HFs("CHECK_NUM")
            rowAPTPYMT1.Item("PYMT_METHOD") = HFs("PYMT_METHOD")
            rowAPTPYMT1.Item("VEND_CODE") = HFs("VEND_CODE")
            rowAPTPYMT1.Item("INIT_DATE") = DATETIME_STAMP
            rowAPTPYMT1.Item("INIT_OPER") = ASCMAIN1.USER_ID
        Else
            Call Save_Header_Fields(UltraGroupBox1)
            Absx1.cmbFor("VEND_TYPE").Text = rowAPTPYMT1.Item("VEND_TYPE") & ""
            Absx1.cmbFor("VEND_CLASS_CODE").Text = rowAPTPYMT1.Item("VEND_CLASS_CODE") & ""
            Absx1.cmbFor("VEND_PYMT_CYCLE").Text = rowAPTPYMT1.Item("VEND_PYMT_CYCLE") & ""
            Absx1.txtFor("VEND_CODE").Text = rowAPTPYMT1.Item("VEND_CODE") & ""
        End If
        Fill_Records("APTPYMT2", New String() {HFs("BATCH_NO_PYMT")})

        'Application.DoEvents()

        If EntryMode = "N" Then
        Else
            dst.AcceptChanges()
        End If

        Dim sql As String = "Select APTINVH1.*" & vbCrLf _
            & ", CASE WHEN NVL(APTINVH1.VEND_ALT_CODE,'VENDOR') = 'VENDOR' THEN APTVEND1.VEND_BANK_COUNTRY ELSE APTVEND2.VEND_ALT_BANK_COUNTRY END VEND_BANK_COUNTRY" & vbCrLf _
            & " from APTINVH1,APTVEND1,APTVEND2 " & vbCrLf _
        & " where APTINVH1.VEND_CODE = APTVEND1.VEND_CODE " & vbCrLf _
        & "   And APTVEND2.VEND_CODE (+) = APTINVH1.VEND_CODE and APTVEND2.VEND_ALT_CODE (+) = APTINVH1.VEND_ALT_CODE" & vbCrLf _
        & "   And APTINVH1.INV_STATUS = 'O' and NVL(INV_APPR_STATUS,'A') = 'A'" & vbCrLf _
        & "   and APTINVH1.BATCH_NO_PYMT is Null"
        If Absx1.chkFor("PYMT_ONLY_SEL_BANK").Checked Then
            sql = sql & " and NVL(APTINVH1.BANK_CODE,'" & ROWs("APTPARM1").Item("AP_PARM_BANK_CODE") & "') = '" & HFs("BANK_CODE") & "'" & vbCrLf
        End If
        If Absx1.cmbFor("VEND_TYPE").Text <> "" Then
            sql = sql & " and APTVEND1.VEND_TYPE = '" & Absx1.cmbFor("VEND_TYPE").Text & "'" & vbCrLf
        End If
        If Absx1.cmbFor("VEND_CLASS_CODE").Text <> "" Then
            sql = sql & " and APTVEND1.VEND_CLASS_CODE = '" & Absx1.cmbFor("VEND_CLASS_CODE").Text & "'" & vbCrLf
        End If
        If Absx1.cmbFor("VEND_PYMT_CYCLE").Text <> "" Then
            sql = sql & " and APTVEND1.VEND_PYMT_CYCLE = '" & Absx1.cmbFor("VEND_PYMT_CYCLE").Text & "'" & vbCrLf
        End If
        If Absx1.txtFor("VEND_CODE").Text <> "" Then
            sql = sql & " and APTINVH1.VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
        End If
        sql = sql & " and NVL(APTVEND1.VEND_ON_HOLD,'0') <> '1'"

        If Absx1.chkFor("PYMT_ONLY_SEL_METHOD").Checked Then
            sql = sql & " and APTINVH1.INV_PYMT_METHOD = '" & Absx1.txtFor("PYMT_METHOD").Text & "'"
            'Else
            '    sql = sql & " and APTINVH1.INV_PYMT_METHOD is Null"
        End If

        sql &= Setup_Open_AP()

        If APTINVH1 = "" Then
            APTINVH1 = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & APTINVH1 & " Add Primary Key (VOUCHER_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & APTINVH1)
            ASCDATA1.ExecuteSQL("Insert into " & APTINVH1 & " " & sql)
        End If

        ASCDATA1.ExecuteSQL($"Insert into {APTINVH1}" & vbCrLf _
                            & " Select APTINVH1.* " & vbCrLf _
                            & ", CASE WHEN NVL(APTINVH1.VEND_ALT_CODE,'VENDOR') = 'VENDOR' THEN APTVEND1.VEND_BANK_COUNTRY ELSE APTVEND2.VEND_ALT_BANK_COUNTRY END VEND_BANK_COUNTRY" & vbCrLf _
                            & " from APTINVH1,APTVEND1,APTVEND2" & vbCrLf _
                            & " where APTINVH1.VEND_CODE = APTVEND1.VEND_CODE " & vbCrLf _
                            & "   And APTVEND2.VEND_CODE (+) = APTINVH1.VEND_CODE and APTVEND2.VEND_ALT_CODE (+) = APTINVH1.VEND_ALT_CODE" & vbCrLf _
                            & $"  and APTINVH1.BATCH_NO_PYMT = '{HFs("BATCH_NO_PYMT")}'")

        'ASCMAIN1.sql = "Select VEND_CODE, VEND_NAME from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from APTPYMT2 where BATCH_NO_PYMT = '" & HFs("BATCH_NO_PYMT") & "' union Select Distinct VEND_CODE from APTPYMT2 where BATCH_NO_PYMT = '" & HFs("BATCH_NO_PYMT") & "') and VEND_ON_HOLD = '1'"
        'Dim tblAPTVEND1_hold As DataTable = ASCDATA1.GetDataTable
        'If tblAPTVEND1_hold.Rows.Count <> 0 Then
        '    EMsg = ""
        '    For Each row As DataRow In tblAPTVEND1_hold.Rows
        '        EMsg &= vbCr & row.Item("VEND_CODE") & ":" & row.Item("VEND_NAME")
        '    Next
        '    MsgBox(EMsg, MsgBoxStyle.OkOnly, "The following Vendors are on Payment Hold")
        'End If

        Application.DoEvents()

        sql = "Update " & APTINVH1 & " Set BATCH_DISC = NVL(INV_DISC_AMT,0) where BATCH_NO_PYMT is Null and CHECK_NUM is Null and NVL(INV_DISC_TAKEN,0) = 0 and NVL(INV_PAYMENTS,0) = 0 and NVL(INV_DISC_AMT,0) > 0"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update " & APTINVH1 & " Set BATCH_DISC = 0 where BATCH_NO_PYMT is Null and INV_DISC_DUE < '" & Format(DateValue(HFs("CHECK_DATE")), "dd-MMM-yyyy") & "' and VEND_CODE in (SELECT DISTINCT APTINVH1.VEND_CODE from " & APTINVH1 & " APTINVH1,APTVEND1 where APTINVH1.VEND_CODE = APTVEND1.VEND_CODE and NVL(APTVEND1.VEND_ALWAYS_TAKE_DISC,'0') <> '1')"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Update " & APTINVH1 & " Set BATCH_PYMT = NVL(INV_BALANCE,0) - NVL(BATCH_DISC,0) where BATCH_NO_PYMT is Null"
        ASCDATA1.ExecuteSQL(sql)
        sql = "Update " & APTINVH1 & " Set BATCH_PYMT = NVL(BATCH_PYMT,0), BATCH_DISC = NVL(BATCH_DISC,0)" _
        & ", VEND_CODE_AP = NVL(VEND_CODE_AP,VEND_CODE), VEND_ALT_CODE = NVL(VEND_ALT_CODE,'VENDOR')"
        ASCDATA1.ExecuteSQL(sql)

        Application.DoEvents()

        dst.EnforceConstraints = False
        Call Fill_Records("APTINVH1", , , "Select * from " & APTINVH1)
        Call Fill_Records("APTVEND1", , , "Select * from APTVEND1 where VEND_CODE in (Select Distinct VEND_CODE from " & APTINVH1 & " union Select Distinct VEND_CODE_AP from " & APTINVH1 & ")")
        dst.EnforceConstraints = True

        ' Setup_Open_AP()

        If EntryMode = "N" Then

        End If

        With grdAPTINVH1.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("VEND_CODE", False)
            .SortedColumns.Add("VOUCHER_NO", False)
        End With

        With grdAPTINVH1_UNS.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("VEND_CODE", False)
            .SortedColumns.Add("VOUCHER_NO", False)
        End With

        With grdAPTINVH1_SEL.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("VEND_CODE", False)
            .SortedColumns.Add("VOUCHER_NO", False)
        End With

        Call Refresh_APTVEND1()

        'CHECK_NUM_ctr = Val(dst.Tables("APTPYMT2").Compute("MAX(CHECK_NUM)", "") & "")
        CHECK_NUM_ctr = Val(Mid(dst.Tables("APTPYMT2").Compute("MAX(CHECK_NUM)", "") & "", 2))

        UltraTabControl1.SelectedTab = UltraTabControl1.Tabs("Unselected Items")
        Call Set_Tab()

        If HFs("VEND_CODE") = "" Then
            UltraTabControl1.Tabs("Match Items").Enabled = False
        Else
            UltraTabControl1.Tabs("Match Items").Enabled = True
        End If


        Call Build_Check_File(True)
        'Call Eliminate_Zero_Checks(True)
        'Call APTINVH1_CHECK_NUM()
        'Call Rebuild_Check_File()
        Call Check_Statistics()

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Try
            Call BeginTrans()

            Dim sql As String = ""

            rowAPTPYMT1.Item("LAST_DATE") = DATETIME_STAMP
            rowAPTPYMT1.Item("LAST_OPER") = ASCMAIN1.USER_ID

            Call Update_Record_TDA("APTPYMT1")

            sql = "Delete from APTPYMT2 where BATCH_NO_PYMT = '" & HFs("BATCH_NO_PYMT") & "'"
            Call Update_Record_TDA("APTPYMT2", sql)

            sql = "Update APTINVH1 set BATCH_NO_PYMT = NULL, CHECK_NUM = NULL, BATCH_PYMT = NULL, BATCH_DISC = NULL where BATCH_NO_PYMT = '" & HFs("BATCH_NO_PYMT") & "'"
            ASCDATA1.ExecuteSQL(sql)

            dst.Tables("APTINVH1").AcceptChanges()
            For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("BATCH_NO_PYMT = '" & HFs("BATCH_NO_PYMT") & "'", "")
                rowAPTINVH1.SetModified()
            Next
            Call Update_Record_TDA("APTINVH1")

            If Absx1.txtFor("CHECK_NUM").Text <> "" Then
                Call Update_as_Paid()
            End If

            Call CommitTrans("Update Complete")

        Catch ex As Exception
            Call Rollback("Error Occurred - Please call ABS", ex)
        End Try

    End Sub

    Sub Delete_Record()
        Call BeginTrans()

        Call Delete_Records("APTPYMT1")
        Call Delete_Records("APTPYMT2")

        ASCDATA1.ExecuteSQL("Update APTINVH1 set BATCH_NO_PYMT = Null, BATCH_PYMT = 0, BATCH_DISC = 0, CHECK_NUM = NULL " _
            & " where BATCH_NO_PYMT = '" & HFs("BATCH_NO_PYMT") & "'")

        rowGLTBANK1 = Fill_Record("GLTBANK1", HFs("BANK_CODE"))
        If rowGLTBANK1.Item("BATCH_NO_PYMT") & "" = HFs("BATCH_NO_PYMT") Then
            rowGLTBANK1.Item("BATCH_NO_PYMT") = ""
            Call Update_Record_TDA("GLTBANK1")
        End If

        Call CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
            & " where BATCH_NO_PYMT = '" & HFs("BATCH_NO_PYMT") & "'")
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Select Case Absx1.GetABSColumnName(sender)

            Case "BANK_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("New", e)
                End If

            Case "BATCH_NO_PYMT"
                If e.KeyCode = Windows.Forms.Keys.Enter Then
                    Call Click_Command("Edit", e)
                End If

        End Select

    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "BATCH_NO_PYMT"
                Call Click_Command("Edit")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "BANK_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("BANK_CODE").Text <> "" Then
                        Call LookUp("GLTBANK1", Absx1.txtFor("BANK_CODE").Text)
                        If cdr IsNot Nothing Then
                            If cdr("BANK_PYMT_METHOD") & "" <> "" Then
                                Absx1.txtFor("PYMT_METHOD").Text = cdr("BANK_PYMT_METHOD") & ""
                            End If
                        End If
                    End If
                End If

        End Select
    End Sub

#End Region

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If EntryMode = "" Then
            Exit Sub
        End If
        Call Set_Tab()
    End Sub

    Private Sub grdAPTVEND1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTVEND1.AfterRowActivate
        If grdAPTVEND1.ActiveRow.IsGroupByRow Then
            grdAPTINVH1.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            Dim VEND_CODE As String = grdAPTVEND1.ActiveRow.Cells("VEND_CODE").Text
            grdAPTINVH1.DataSource = New DataView(dst.Tables("APTINVH1"), "VEND_CODE = '" & VEND_CODE & "'", "VOUCHER_NO", DataViewRowState.CurrentRows)
            grdAPTINVH1.DisplayLayout.Bands(0).SummaryFooterCaption = "Totals for Vendor " & VEND_CODE
            grdAPTINVH1.Visible = True
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Overrides Function Validate_Code_Special( _
    ByVal COLUMN_NAME As String, _
    ByVal IsValid As Boolean) As Boolean

        Select Case COLUMN_NAME
            Case "BANK_CODE"
                If cdr IsNot Nothing Then
                    If cdr.Item("BANK_STATUS") & "" = "" Then
                        EMsg &= vbCr & "Bank does not have a valid Status Code"
                        IsValid = False
                    Else
                        If cdr.Item("BANK_STATUS") & "" <> "A" Then
                            EMsg &= vbCr & "Bank is not Active"
                            IsValid = False
                        End If
                        If cdr.Item("BANK_ACCT_ID") & "" = "" Or cdr.Item("ROUTING_NO") & "" = "" Then
                            EMsg &= vbCr & "Bank Record is missing a Bank Account ID or Routing No"
                            IsValid = False
                        End If
                    End If
                End If

        End Select

        Return IsValid

    End Function

    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click
        If Absx1.dteFor("SELECT_DATE").Value & "" = "" Then
            MsgBox("You Must First Specify a Date")
            Exit Sub
        End If

        Dim sql As String = Absx1.optFor("SELECT_BASED_ON").Value & " <= '" & Format(Absx1.dteFor("SELECT_DATE").Value, "MM/dd/yyyy") & "'"

        Dim VEND_CODE As String = ""
        If UltraTabControl1.ActiveTab.Key = "Select by Vendor" Then
            If grdAPTVEND1.ActiveRow Is Nothing Then
                MsgBox("No Vendor Selected")
                Exit Sub
            Else
                VEND_CODE = grdAPTVEND1.ActiveRow.Cells("VEND_CODE").Text
                sql = sql & " AND VEND_CODE = '" & VEND_CODE & "'"
                Call Select_for_Payment(sql, VEND_CODE)
            End If
        Else
            Dim INV_PYMT_CYCLE As String = Absx1.cmbFor("INV_PYMT_CYCLE").Text
            'If INV_PYMT_CYCLE <> "" And INV_PYMT_CYCLE <> "*" Then
            '    If INV_PYMT_CYCLE = "0" Then
            '        sql = sql & " AND INV_PYMT_CYCLE is Null"
            '    Else
            '        sql = sql & " AND INV_PYMT_CYCLE = '" & INV_PYMT_CYCLE & "'"
            '    End If
            'End If
            If INV_PYMT_CYCLE <> "" Then
                sql = sql & " AND INV_PYMT_CYCLE = '" & INV_PYMT_CYCLE & "'"
            End If

            Dim VEND_CLASS_CODE As String = Absx1.cmbFor("VEND_CLASS_CODE").Text
            If VEND_CLASS_CODE <> "" Then
                sql = sql & " AND PARENT.VEND_CLASS_CODE = '" & VEND_CLASS_CODE & "'"
            End If
            Dim VEND_PYMT_CYCLE As String = Absx1.cmbFor("VEND_PYMT_CYCLE").Text
            If VEND_PYMT_CYCLE <> "" Then
                sql = sql & " AND PARENT.VEND_PYMT_CYCLE = '" & VEND_PYMT_CYCLE & "'"
            End If
            Dim VEND_TYPE As String = Absx1.cmbFor("VEND_TYPE").Text
            If VEND_TYPE <> "" Then
                sql = sql & " AND PARENT.VEND_TYPE = '" & VEND_TYPE & "'"
            End If

            Call Select_for_Payment(sql)
        End If

    End Sub

    Sub Select_for_Payment(ByVal sql As String, Optional ByVal VEND_CODE As String = "")
        Call ASCMAIN1.Progress("Now Selecting AP Items")
        Me.Cursor = Cursors.WaitCursor

        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select(sql, "", DataViewRowState.CurrentRows)
            rowAPTINVH1.Item("BATCH_NO_PYMT") = HFs("BATCH_NO_PYMT")
        Next

        Call Calculate_Stats(VEND_CODE)

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Calculate_Stats( _
    Optional ByVal VEND_CODE As String = "", _
    Optional ByVal CHECK_NUM As String = "")

        Call Build_Check_File(False, VEND_CODE, CHECK_NUM)

        With grdAPTINVH1_UNS.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("VEND_CODE", False)
            .SortedColumns.Add("VOUCHER_NO", False)
        End With

        With grdAPTINVH1_SEL.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("VEND_CODE", False)
            .SortedColumns.Add("VOUCHER_NO", False)
        End With

        'Call Refresh_APTVEND1(VEND_CODE)
    End Sub

    Sub Build_Check_File( _
    ByVal rid_of_zero_checks As Boolean, _
    Optional ByVal VEND_CODE As String = "", _
    Optional ByVal CHECK_NUM As String = "")

        Me.Cursor = Cursors.WaitCursor

        If VEND_CODE = "" Then
            dst.Tables("APTPYMT2").Rows.Clear()
        Else
            If CHECK_NUM = "" Then
                ASCDATA1.DeleteRows("APTPYMT2", "VEND_CODE = '" & VEND_CODE & "'")
            End If
        End If

        Dim tbl As DataTable = dst.Tables("APTPYMTC")
        tbl.Rows.Clear()

        Dim sql_VEND_CODE As String = ""
        If VEND_CODE <> "" Then
            sql_VEND_CODE = " and VEND_CODE = '" & VEND_CODE & "'"
        End If
        If CHECK_NUM <> "" Then
            sql_VEND_CODE = " and CHECK_NUM = '" & CHECK_NUM & "'"
        End If
        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("BATCH_NO_PYMT is Not Null" & sql_VEND_CODE, "VEND_CODE,VEND_CODE_AP,VEND_ALT_CODE,VOUCHER_NO")
            With rowAPTINVH1
                Dim VEND_CODE_x As String = .Item("VEND_CODE")
                Dim VEND_CODE_AP As String = .Item("VEND_CODE_AP") & ""
                Dim VEND_ALT_CODE As String = .Item("VEND_ALT_CODE") & ""
                Dim VOUCHER_NO As String = "0000000000"
                If .Item("INV_SEP_CHECK") & "" = "1" Then
                    VOUCHER_NO = rowAPTINVH1.Item("VOUCHER_NO")
                End If
                Dim row As DataRow = tbl.Rows.Find(New Object() {VEND_CODE_x, VEND_CODE_AP, VEND_ALT_CODE, VOUCHER_NO})
                If row Is Nothing Then
                    row = tbl.NewRow
                    row.Item("VEND_CODE") = VEND_CODE_x
                    row.Item("VEND_CODE_AP") = VEND_CODE_AP
                    row.Item("VEND_ALT_CODE") = VEND_ALT_CODE
                    row.Item("VOUCHER_NO") = VOUCHER_NO
                    tbl.Rows.Add(row)
                End If
                row.Item("BATCH_PYMT") = Val(row.Item("BATCH_PYMT") & "") + Val(.Item("BATCH_PYMT") & "")
                row.Item("BATCH_DISC") = Val(row.Item("BATCH_DISC") & "") + Val(.Item("BATCH_DISC") & "")
            End With
        Next

        If VEND_CODE = "" Then
            CHECK_NUM_ctr = 0
        End If
        If CHECK_NUM <> "" Then
            Dim rowAPTPYMT2 As DataRow = dst.Tables("APTPYMT2").Rows.Find(New Object() {HFs("BATCH_NO_PYMT"), CHECK_NUM})
            If tbl.Rows.Count = 0 Then
                rowAPTPYMT2.Delete()
            Else
                rowAPTPYMT2.Item("BATCH_PYMT") = tbl.Rows(0).Item("BATCH_PYMT")
                rowAPTPYMT2.Item("BATCH_DISC") = tbl.Rows(0).Item("BATCH_DISC")
            End If
        Else
            For Each row As DataRow In tbl.Rows
                Dim rowAPTPYMT2 As DataRow = dst.Tables("APTPYMT2").NewRow
                rowAPTPYMT2.Item("BATCH_NO_PYMT") = HFs("BATCH_NO_PYMT")
                'If CHECK_NUM = "" Then
                    CHECK_NUM_ctr = CHECK_NUM_ctr + 1
                    rowAPTPYMT2.Item("CHECK_NUM") = "T" & Format(CHECK_NUM_ctr, "000000000")
                'Else
                '    rowAPTPYMT2.Item("CHECK_NUM") = CHECK_NUM
                'End If
                rowAPTPYMT2.Item("VEND_CODE_AP") = row.Item("VEND_CODE_AP")
                rowAPTPYMT2.Item("VOUCHER_NO") = row.Item("VOUCHER_NO")
                rowAPTPYMT2.Item("BATCH_PYMT") = row.Item("BATCH_PYMT")
                rowAPTPYMT2.Item("BATCH_DISC") = row.Item("BATCH_DISC")
                rowAPTPYMT2.Item("VEND_ALT_CODE") = row.Item("VEND_ALT_CODE")
                rowAPTPYMT2.Item("VEND_CODE") = row.Item("VEND_CODE")
                Dim VEND_NAME As String = ""
                If row.Item("VEND_CODE") <> row.Item("VEND_CODE_AP") Then
                    VEND_NAME = dst.Tables("APTVEND1").Rows.Find(New Object() {row.Item("VEND_CODE_AP")}).Item("VEND_NAME")
                Else
                    If row.Item("VEND_ALT_CODE") = "VENDOR" Then
                        VEND_NAME = dst.Tables("APTVEND1").Rows.Find(New Object() {row.Item("VEND_CODE")}).Item("VEND_NAME")
                    Else
                        VEND_NAME = LookUp("APTVEND2", New String() {row.Item("VEND_CODE"), row.Item("VEND_ALT_CODE")}).Item("VEND_ALT_NAME") & ""
                    End If
                End If
                rowAPTPYMT2.Item("VEND_NAME") = VEND_NAME
                dst.Tables("APTPYMT2").Rows.Add(rowAPTPYMT2)
            Next
        End If

        If rid_of_zero_checks Then
            Eliminate_Zero_Checks()
        End If

        APTINVH1_CHECK_NUM(VEND_CODE)
        Refresh_APTVEND1(VEND_CODE)
        Check_Statistics()

        Me.Cursor = Cursors.Default
    End Sub

    Sub Eliminate_Zero_Checks(Optional ByVal clear_APTINVH1 As Boolean = False)
        ' Eliminate Checks < 0 or <= 0 depending on Parameter and Bank settings

        Dim z As String
        If ROWs("APTPARM1").Item("AP_PARM_PRINT_0_CHECKS") & "" = "1" _
        And rowGLTBANK1.Item("BANK_PRINT_0_CHECKS") & "" = "1" Then
            z = "<"
        Else
            z = "<="
        End If

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Select("", "")
            rowAPTPYMT2.Item("BATCH_PYMT") = System.Math.Round(Val(rowAPTPYMT2.Item("BATCH_PYMT")), 2)
            rowAPTPYMT2.Item("BATCH_DISC") = System.Math.Round(Val(rowAPTPYMT2.Item("BATCH_DISC")), 2)
        Next

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Select("BATCH_PYMT " & z & " 0.0", "")
            If clear_APTINVH1 Then
                Dim CHECK_NUM As String = rowAPTPYMT2.Item("CHECK_NUM")
                For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("CHECK_NUM = '" & CHECK_NUM & "'", "")
                    rowAPTINVH1.Item("BATCH_NO_PYMT") = Null
                Next
            End If

            rowAPTPYMT2.Delete()
        Next
    End Sub

    Sub APTINVH1_CHECK_NUM(Optional ByVal VEND_CODE As String = "")

        Dim sql_VEND_CODE As String = ""
        If VEND_CODE <> "" Then
            sql_VEND_CODE = "VEND_CODE = '" & VEND_CODE & "'"
        End If

        ' Clear Out Check Numbers

        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select(sql_VEND_CODE, "")
            rowAPTINVH1.Item("CHECK_NUM") = DBNull.Value
        Next

        ' Mark Check Numbers in APTINVH1

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Select(sql_VEND_CODE, "")
            Dim sql As String = "" _
                & "VEND_CODE = '" & rowAPTPYMT2.Item("VEND_CODE") & "'" & vbCrLf _
                & " and VEND_CODE_AP = '" & rowAPTPYMT2.Item("VEND_CODE_AP") & "'" & vbCrLf _
                & " and VEND_ALT_CODE = '" & rowAPTPYMT2.Item("VEND_ALT_CODE") & "'" & vbCrLf _
                & " and BATCH_NO_PYMT is Not Null" & vbCrLf _
                & " and CHECK_NUM is Null"

            If rowAPTPYMT2.Item("VOUCHER_NO") <> "0000000000" Then
                sql = sql & " AND VOUCHER_NO = '" & rowAPTPYMT2.Item("VOUCHER_NO") & "'"
            Else
                sql = sql & " AND ISNULL(INV_SEP_CHECK,'0') <> '1'"
            End If
            ' PROBLEM WITH 000'S AND NON-000'S IN SAME VENDOR

            For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select(sql, "")
                rowAPTINVH1.Item("CHECK_NUM") = rowAPTPYMT2.Item("CHECK_NUM")
            Next
        Next

    End Sub

    Sub Check_Statistics()

        dst.Tables("APTPYMT0").Rows.Clear()
        Dim dvw As New DataView(dst.Tables("APTPYMT2"))
        Dim rowAPTPYMT0 As DataRow
        Dim sqlfilter As String = ""
        For i As Integer = 0 To 3
            rowAPTPYMT0 = dst.Tables("APTPYMT0").NewRow
            rowAPTPYMT0.Item("CHECK_TYPE") = CStr(i)
            Select Case i
                Case 0
                    sqlfilter = ""
                Case 1
                    sqlfilter = "BATCH_PYMT = 0"
                Case 2
                    sqlfilter = "BATCH_PYMT < 0"
                Case 3
                    sqlfilter = "BATCH_PYMT > 0"
            End Select
            dvw.RowFilter = sqlfilter
            rowAPTPYMT0.Item("CHECK_TYPE_DESC") = New String() {"All", "=0", "<0", ">0"}(i)

            rowAPTPYMT0.Item("CHECK_TYPE_COUNT") = dvw.Count
            rowAPTPYMT0.Item("CHECK_TYPE_AMOUNT") = dst.Tables("APTPYMT2").Compute("SUM(BATCH_PYMT)", sqlfilter)
            dst.Tables("APTPYMT0").Rows.Add(rowAPTPYMT0)
        Next

        If grdAPTPYMT2.Rows.Count = 0 Then
            grdAPTINVH1_DTL.Visible = False
        End If
    End Sub

    Sub Refresh_APTVEND1(Optional ByVal VEND_CODE As String = "")
        Dim tbl As DataTable = dst.Tables("APTINVH1")

        Dim sql_VEND_CODE As String = ""
        If VEND_CODE <> "" Then
            sql_VEND_CODE = "VEND_CODE = '" & VEND_CODE & "'"
        End If
        For Each rowAPTVEND1 In dst.Tables("APTVEND1").Select(sql_VEND_CODE, "")

            With rowAPTVEND1
                Dim sql As String
                sql = "VEND_CODE = '" & .Item("VEND_CODE") & "' and BATCH_NO_PYMT is Null"
                .Item("UNS_ITEMS") = tbl.Compute("COUNT(VOUCHER_NO)", sql)
                .Item("UNS_PYMT") = tbl.Compute("SUM(BATCH_PYMT)", sql)
                .Item("UNS_DISC") = tbl.Compute("SUM(BATCH_DISC)", sql)
                sql = "VEND_CODE = '" & .Item("VEND_CODE") & "' and BATCH_NO_PYMT is Not Null"
                .Item("SEL_ITEMS") = tbl.Compute("COUNT(VOUCHER_NO)", sql)
                .Item("SEL_PYMT") = tbl.Compute("SUM(BATCH_PYMT)", sql)
                .Item("SEL_DISC") = tbl.Compute("SUM(BATCH_DISC)", sql)
            End With
        Next

        If grdAPTVEND1.Rows.Count > 0 Then
            If VEND_CODE <> "" Then
                grdAPTVEND1.ActiveRow = grdAPTVEND1.Rows.GetRowWithListIndex(dst.Tables("APTVEND1").Rows.IndexOf(dst.Tables("APTVEND1").Rows.Find(New Object() {VEND_CODE})))
            Else
                grdAPTVEND1.ActiveRow = grdAPTVEND1.Rows(0)
            End If
        End If

    End Sub

    Sub Set_Tab()
        With UltraExplorerBar1
            .Groups("Filter Vendors By").Visible = False
            .Groups("Filter AP Items By").Visible = False
            .Groups("Select by Date").Visible = False
            .Groups("Selection Tools").Visible = False
            .Groups("Reversal Tools").Visible = False
            .Groups("Check Statistics").Visible = False
            .Groups("Remove Selected ...").Visible = False

            .Groups("Match Invoices").Visible = False
            .Groups("Show Items").Visible = False
            lblThisVendorOnly.Visible = False

            Select Case UltraTabControl1.ActiveTab.Key
                Case "Unselected Items"
                    .Groups("Selection Tools").Visible = True
                    .Groups("Filter AP Items By").Visible = True
                    .Groups("Filter Vendors By").Visible = True
                    .Groups("Select by Date").Visible = True

                Case "Select by Vendor"
                    .Groups("Select by Date").Visible = True
                    lblThisVendorOnly.Visible = True

                Case "Selected Items"
                    .Groups("Reversal Tools").Visible = True
                    .Groups("Check Statistics").Visible = True

                Case "Payments"
                    .Groups("Remove Selected ...").Visible = True
                    .Groups("Check Statistics").Visible = True

                Case "Match Items"
                    .Groups("Match Invoices").Visible = True
                    .Groups("Show Items").Visible = True
                    Call Load_Matched()
            End Select
        End With
    End Sub

    Private Sub grdAPTPYMT2_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdAPTPYMT2.AfterRowActivate
        If grdAPTPYMT2.ActiveRow.IsGroupByRow Then
            grdAPTINVH1_DTL.Visible = False
        Else
            Me.Cursor = Cursors.WaitCursor
            Dim CHECK_NUM As String = grdAPTPYMT2.ActiveRow.Cells("CHECK_NUM").Text
            grdAPTINVH1_DTL.DataSource = New DataView(dst.Tables("APTINVH1"), "CHECK_NUM = '" & CHECK_NUM & "'", "VOUCHER_NO", DataViewRowState.CurrentRows)
            grdAPTINVH1_DTL.DisplayLayout.Bands(0).SummaryFooterCaption = "AP Items Paid on Check " & CHECK_NUM
            grdAPTINVH1_DTL.Visible = True
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub grdAPTVEND1_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles grdAPTVEND1.AfterSelectChange

    End Sub

    Private Sub grdAPTVEND1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTVEND1.ClickCellButton
        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim VEND_CODE As String = e.Cell.Row.Cells("VEND_CODE").Text
        Select Case COLUMN_NAME
            Case "SEL"
                Dim sql As String = "VEND_CODE = '" & VEND_CODE & "'"
                Call Select_for_Payment(sql, VEND_CODE)
            Case "REV"
                For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("VEND_CODE = '" & VEND_CODE & "'", "")
                    rowAPTINVH1.Item("BATCH_NO_PYMT") = Null
                    rowAPTINVH1.Item("CHECK_NUM") = Null
                Next
                Call Calculate_Stats(VEND_CODE)
        End Select
    End Sub

    Overrides Sub cmb_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        If e.KeyCode = Keys.Delete Then
            DirectCast(sender, Control).Text = ""
        End If
    End Sub

    Private Sub grdAPTPYMT2_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTPYMT2.InitializeRow
    End Sub

    Private Sub grdAPTINVH1_DTL_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTINVH1_DTL.InitializeRow
    End Sub

    Private Sub grdAPTINVH1_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH1.ClickCellButton
        Dim COLUMN_NAME As String = e.Cell.Column.Key
        Dim VEND_CODE As String = e.Cell.Row.Cells("VEND_CODE").Text
        Dim VOUCHER_NO As String = e.Cell.Row.Cells("VOUCHER_NO").Text
        Select Case COLUMN_NAME
            Case "SEL"
                Dim sql As String = "VOUCHER_NO = '" & VOUCHER_NO & "'"
                Call Select_for_Payment(sql, VEND_CODE)

            Case "REV"
                Call ASCMAIN1.Progress("Now Reversing this AP Item")
                Dim CHECK_NUM As String = grdAPTINVH1.ActiveRow.Cells("CHECK_NUM").Text
                grdAPTINVH1.ActiveRow.Cells("BATCH_NO_PYMT").Value = Null
                grdAPTINVH1.ActiveRow.Cells("CHECK_NUM").Value = Null
                grdAPTINVH1.UpdateData()
                Call Calculate_Stats(VEND_CODE, CHECK_NUM)
                Call ASCMAIN1.Progress("")

            Case "TOGGLE_DISC"
                If Val(grdAPTINVH1.ActiveRow.Cells("BATCH_DISC").Text) = 0 And _
                   Val(grdAPTINVH1.ActiveRow.Cells("INV_DISC_AMT").Value & "") = 0 Then
                    Exit Sub
                End If
                If Val(grdAPTINVH1.ActiveRow.Cells("BATCH_DISC").Text) = 0 Then
                    grdAPTINVH1.ActiveRow.Cells("BATCH_DISC").Value = grdAPTINVH1.ActiveRow.Cells("INV_DISC_AMT").Value
                Else
                    grdAPTINVH1.ActiveRow.Cells("BATCH_DISC").Value = 0
                End If
                grdAPTINVH1.ActiveRow.Cells("BATCH_PYMT").Value = Val(grdAPTINVH1.ActiveRow.Cells("INV_BALANCE").Value & "") - Val(grdAPTINVH1.ActiveRow.Cells("BATCH_DISC").Value & "")
                '        Stop
                Dim CHECK_NUM As String = grdAPTINVH1.ActiveRow.Cells("CHECK_NUM").Text
                Call Calculate_Stats(VEND_CODE, CHECK_NUM)
                Call ASCMAIN1.Progress("")

        End Select

    End Sub

    Private Sub grdAPTINVH1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTINVH1.InitializeLayout

    End Sub

    Private Sub grdAPTINVH1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTINVH1.InitializeRow
        If e.Row.IsDataRow Then
            If e.Row.Cells("BATCH_NO_PYMT").Text = "" Then
                e.Row.Appearance.BackColor = Drawing.Color.Empty
            Else
                e.Row.Appearance.BackColor = Drawing.Color.LightGreen
            End If
            e.Row.Cells("SEL").Value = "S"
            e.Row.Cells("REV").Value = "R"

            If Val(e.Row.Cells("INV_DISC_AMT").Text) = 0 Then
                e.Row.Cells("TOGGLE_DISC").Value = ""
            Else
                If Val(e.Row.Cells("BATCH_DISC").Text) = 0 Then
                    e.Row.Cells("TOGGLE_DISC").Value = "Take Disc"
                Else
                    e.Row.Cells("TOGGLE_DISC").Value = "Zero Disc"
                End If
            End If

        End If
    End Sub

    Private Sub grdAPTPYMT2_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTPYMT2.InitializeLayout

    End Sub

    Sub Print_Report()
        Call Print_Report_Begin()
        CR_params.Add("PAYMENT_SELECTION", "1")
        Generate_Report("APRPYMT1", "Payment Selection Register")
        Call Print_Report_End()
    End Sub

    Private Sub grdAPTPYMTX_DoubleClickRow(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdAPTPYMTX.DoubleClickRow
        If grdAPTPYMTX.ActiveRow.IsGroupByRow Then
        Else
            Absx1.txtFor("BATCH_NO_PYMT").Text = grdAPTPYMTX.ActiveRow.Cells("BATCH_NO_PYMT").Text
            Call Click_Command("Edit")
        End If
    End Sub

    Private Sub grdAPTVEND1_InitializeGroupByRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeGroupByRowEventArgs) Handles grdAPTVEND1.InitializeGroupByRow
        ' Stop
        Dim TOTAL_AMT_SELECTED As Double = 0

        Dim Z As String = ""
        With e.Row.ChildBands(0).Rows
            Z = Z & " " & CStr(.SummaryValues("UNS_ITEMS").Value) & " UnSelected Items (" & Format(.SummaryValues("UNS_PYMT").Value, "$##,##0.00") & ")"
            Z = Z & "; " & CStr(.SummaryValues("SEL_ITEMS").Value) & " Selected Items (" & Format(.SummaryValues("SEL_PYMT").Value, "$##,##0.00") & ")"
            'e.Row.Appearance.ForeColor = Drawing.Color.White
        End With



        e.Row.Description &= Z

        'If e.Row.Column.Key = "Country" Then
        ' If the group has more than 5 items, then make the background color
        ' of the row to red
        If e.Row.ChildBands(0).Rows.SummaryValues("SEL_ITEMS").Value > 0 Then
            e.Row.Appearance.BackColor = System.Drawing.Color.LightGreen
        Else
            e.Row.Appearance.ResetBackColor()
        End If
        'End If

    End Sub

    Private Sub grdAPTVEND1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTVEND1.InitializeLayout

    End Sub

    Private Sub grdAPTVEND1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdAPTVEND1.InitializeRow
        If e.Row.IsDataRow Then
            e.Row.Cells("SEL").Value = "S"
            e.Row.Cells("REV").Value = "R"

            If e.Row.Cells("VEND_ON_HOLD").Text = "1" Then
                e.Row.Appearance.ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

    Sub Remove_0s()
        Call Eliminate_Zero_Checks(True)
        Call APTINVH1_CHECK_NUM()
        Call Check_Statistics()
    End Sub

    Private Sub grpSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grpSelect.Click

    End Sub

    Sub Rebuild_Check_File()
        Call ASCMAIN1.Progress("Now Rebuilding Check File")
        Me.Cursor = Cursors.WaitCursor

        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Select("", "", DataViewRowState.CurrentRows)
            rowAPTPYMT2.Item("CHECK_NUM") = "X" & Mid(rowAPTPYMT2.Item("CHECK_NUM"), 2)
        Next

        'For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("", "CHECK_NUM is Not Null", DataViewRowState.CurrentRows)
        '    rowAPTINVH1.Item("CHECK_NUM") = "X" & Mid(rowAPTINVH1.Item("CHECK_NUM"), 2)
        'Next

        Dim CX As New Dictionary(Of String, String)
        CHECK_NUM_ctr = 0
        For Each rowAPTPYMT2 As DataRow In dst.Tables("APTPYMT2").Select("", "VEND_CODE, VEND_CODE_AP, VEND_ALT_CODE, VOUCHER_NO", DataViewRowState.CurrentRows)
            CHECK_NUM_ctr += 1
            Dim CHECK_NUM_old As String = rowAPTPYMT2.Item("CHECK_NUM")
            'Dim CHECK_NUM As String = Format(CHECK_NUM_ctr, "0000000000")
            Dim CHECK_NUM As String = "T" & Format(CHECK_NUM_ctr, "000000000")
            CX.Add(CHECK_NUM_old, CHECK_NUM)
            rowAPTPYMT2.Item("CHECK_NUM") = CHECK_NUM
        Next
        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("CHECK_NUM is Not Null", "", DataViewRowState.CurrentRows)
            rowAPTINVH1.Item("CHECK_NUM") = CX("X" & Mid(rowAPTINVH1.Item("CHECK_NUM"), 2))
        Next


        With grdAPTPYMT2.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("CHECK_NUM", False)
        End With

        If grdAPTPYMT2.Rows.Count = 0 Then
            grdAPTINVH1.Visible = False
        Else
            grdAPTPYMT2.ActiveRow = grdAPTPYMT2.Rows(0)
        End If

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Sub Load_Matched()
        Call ASCMAIN1.Progress("Now Loading Un-Selected Items for Matching")
        Me.Cursor = Cursors.WaitCursor

        dst.Tables("APTINVH1_DUE_TO").Rows.Clear()

        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("BATCH_NO_PYMT is Null", "VEND_CODE", DataViewRowState.CurrentRows)
            Dim rowAPTINVH1_DUE_TO As DataRow = dst.Tables("APTINVH1_DUE_TO").NewRow
            rowAPTINVH1_DUE_TO("VOUCHER_NO") = rowAPTINVH1("VOUCHER_NO")
            rowAPTINVH1_DUE_TO("INV_TYPE") = rowAPTINVH1("INV_TYPE")
            rowAPTINVH1_DUE_TO("INV_NUM") = rowAPTINVH1("INV_NUM")
            rowAPTINVH1_DUE_TO("INV_DATE") = rowAPTINVH1("INV_DATE")
            rowAPTINVH1_DUE_TO("INV_BALANCE") = rowAPTINVH1("INV_BALANCE")
            rowAPTINVH1_DUE_TO("INV_REF") = rowAPTINVH1("INV_REF")
            rowAPTINVH1_DUE_TO("PO_ORDER_NO") = rowAPTINVH1("PO_ORDER_NO")
            rowAPTINVH1_DUE_TO("INV_DUE_DATE") = rowAPTINVH1("INV_DUE_DATE")
            rowAPTINVH1_DUE_TO("BATCH_NO_PYMT") = rowAPTINVH1("BATCH_NO_PYMT")
            rowAPTINVH1_DUE_TO("BATCH_PYMT") = rowAPTINVH1("BATCH_PYMT")
            rowAPTINVH1_DUE_TO("BATCH_DISC") = rowAPTINVH1("BATCH_DISC")
            rowAPTINVH1_DUE_TO("INV_PYMT_METHOD") = rowAPTINVH1("INV_PYMT_METHOD")
            rowAPTINVH1_DUE_TO("INV_BALANCE_ABS") = System.Math.Abs(CDbl(Val(rowAPTINVH1("INV_BALANCE") & "")))
            rowAPTINVH1_DUE_TO("MATCH") = "0"
            rowAPTINVH1_DUE_TO("MATCH_NO") = 0
            dst.Tables("APTINVH1_DUE_TO").Rows.Add(rowAPTINVH1_DUE_TO)
        Next
        'Stop
        If optMatchShow.Value <> "A" Then

            optMatchShow.Value = "A"
        Else
            Call Set_Matched_DataSource()
        End If

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Private Sub cmdMatch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMatch.Click
        For Each rowpos As DataRow In dst.Tables("APTINVH1_DUE_TO").Select("INV_TYPE = 'I' and MATCH = '0' and INV_BALANCE > 0", "INV_BALANCE")
            Dim TOL_DAYS As Integer = Val(Absx1.numFor("TOL_DAYS").Value & "")
            Dim TOL_AMT As Double = Val(Absx1.numFor("TOL_AMT").Value & "")
            Dim INV_BALANCE As Double = Val(rowpos("INV_BALANCE_ABS") & "")
            Dim INV_DATE As Date = rowpos("INV_DATE")
            Dim MATCH_NO As Integer = Val(dst.Tables("APTINVH1_DUE_TO").Compute("MAX(MATCH_NO)", ""))
            Dim sql As String = "MATCH = '0'"
            If TOL_AMT = 0 Then
                'sql = sql & " AND INV_BALANCE_ABS = " & CStr(-1 * INV_BALANCE)

                sql = sql & " AND INV_BALANCE < 0"
                sql = sql & " AND INV_BALANCE_ABS >= " & CStr(INV_BALANCE - 0.005)
                sql = sql & " AND INV_BALANCE_ABS <= " & CStr(INV_BALANCE + 0.005)

            Else
                sql = sql & " AND INV_BALANCE < 0"
                sql = sql & " AND INV_BALANCE_ABS >= " & CStr(INV_BALANCE - TOL_AMT)
                sql = sql & " AND INV_BALANCE_ABS <= " & CStr(INV_BALANCE + TOL_AMT)
            End If
            If TOL_DAYS = 0 Then
                sql = sql & " AND INV_DATE = '" & Format(INV_DATE, "MM/dd/yyyy") & "'"
            Else
                sql = sql & " AND INV_DATE >= '" & Format(INV_DATE.AddDays(-1 * TOL_DAYS), "MM/dd/yyyy") & "'"
                sql = sql & " AND INV_DATE <= '" & Format(INV_DATE.AddDays(1 * TOL_DAYS), "MM/dd/yyyy") & "'"
            End If

            For Each rowneg As DataRow In dst.Tables("APTINVH1_DUE_TO").Select(sql, "")
                rowneg("MATCH") = "1"
                rowpos("MATCH") = "1"
                MATCH_NO = MATCH_NO + 1
                rowneg("MATCH_NO") = MATCH_NO
                rowpos("MATCH_NO") = MATCH_NO
                Exit For ' Match to 1st hit only
            Next
        Next
    End Sub

    Private Sub cmdMatchSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMatchSelect.Click
        Call ASCMAIN1.Progress("Now Selecting Matched Items")
        Me.Cursor = Cursors.WaitCursor

        'dst.Tables("APTINVH1_DUE_TO").AcceptChanges()
        For Each row As DataRow In dst.Tables("APTINVH1_DUE_TO").Select("MATCH = '1'", "")
            Dim rowAPTINVH1 As DataRow = dst.Tables("APTINVH1").Rows.Find(row("VOUCHER_NO"))
            rowAPTINVH1.Item("BATCH_NO_PYMT") = HFs("BATCH_NO_PYMT")
            row.Delete()
        Next
        'dst.Tables("APTINVH1_DUE_TO").AcceptChanges()

        Call Build_Check_File(False, HFs("VEND_CODE"))

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")
    End Sub

    Private Sub grdAPTINVH1_DUE_TO_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdAPTINVH1_DUE_TO.AfterCellUpdate
        'Stop
    End Sub

    Private Sub grdAPTINVH1_DUE_TO_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles grdAPTINVH1_DUE_TO.MouseUp
        'grdAPTINVH1_DUE_TO.Update()
    End Sub

    Sub Set_Matched_DataSource()
        Dim sql As String = ""
        If optMatchShow.Value = "M" Then
            sql = "MATCH = '1'"
        ElseIf optMatchShow.Value = "U" Then
            sql = "MATCH = '0'"
        End If
        grdAPTINVH1_DUE_TO.DataSource = New DataView(dst.Tables("APTINVH1_DUE_TO"), sql, "", DataViewRowState.CurrentRows)

        With grdAPTINVH1_DUE_TO.DisplayLayout.Bands(0)
            .SortedColumns.Clear()
            .SortedColumns.Add("INV_BALANCE_ABS", True)
            .SortedColumns.Add("INV_DATE", True)
        End With

    End Sub

    Private Sub optMatchShow_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMatchShow.ValueChanged
        If dst.Tables.Count = 0 Then Exit Sub
        Call Set_Matched_DataSource()
    End Sub

    Sub Update_as_Paid()

        dst.Tables("APTINVH1").AcceptChanges()
        dst.Tables("APTCHCK1").Rows.Clear()
        dst.Tables("APTCHCK1").AcceptChanges()
        dst.Tables("APTCHCK2").Rows.Clear()
        dst.Tables("APTCHCK2").AcceptChanges()
        dst.Tables("APTVEND5").Rows.Clear()
        dst.Tables("APTVEND5").AcceptChanges()
        dst.Tables("GLTBANK1").Rows.Clear()
        dst.Tables("GLTBANK1").AcceptChanges()

        Dim VEND_CODE As String = HFs("VEND_CODE")
        Dim VEND_CODE_AP As String = HFs("VEND_CODE")
        If VEND_CODE_AP = "" Then
            VEND_CODE_AP = VEND_CODE
        End If
        Dim VEND_ALT_CODE As String = "VENDOR"

        Dim VEND_NAME As String
        Dim rowPayee As DataRow
        If VEND_CODE_AP <> "" And VEND_CODE_AP <> VEND_CODE Then
            rowPayee = LookUp("APTVEND1", VEND_CODE_AP)
            VEND_NAME = rowPayee.Item("VEND_NAME")
        Else
            VEND_NAME = HFs("VEND_NAME")
        End If

        Dim BANK_CODE As String = HFs("BANK_CODE")
        Dim PYMT_METHOD As String = HFs("PYMT_METHOD")
        Dim CHECK_NUM As String = HFs("CHECK_NUM")
        Dim CHECK_DATE As Date = DateValue(HFs("CHECK_DATE"))
        Dim CHECK_AMT As Double = 0

        Dim SEQ_NUM As Integer
        SEQ_NUM = 0
        For Each rowAPTINVH1 As DataRow In dst.Tables("APTINVH1").Select("BATCH_NO_PYMT is Not Null", "VEND_CODE", DataViewRowState.CurrentRows)
            rowAPTINVH1("INV_STATUS") = "P"
            rowAPTINVH1("INV_PAYMENTS") = rowAPTINVH1("BATCH_PYMT")
            rowAPTINVH1("INV_DISC_TAKEN") = rowAPTINVH1("BATCH_DISC")
            rowAPTINVH1("INV_LAST_PMT_DATE") = CHECK_DATE
            rowAPTINVH1("BATCH_NO_PYMT") = ""
            rowAPTINVH1("INV_BALANCE") = 0
            rowAPTINVH1("BATCH_PYMT") = 0
            rowAPTINVH1("BATCH_DISC") = 0
            rowAPTINVH1("CHECK_NUM") = CHECK_NUM
            rowAPTINVH1("CHECK_DATE") = CHECK_DATE

            Dim rowAPTCHCK2 As DataRow = dst.Tables("APTCHCK2").NewRow
            rowAPTCHCK2("BANK_CODE") = BANK_CODE
            rowAPTCHCK2("CHECK_NUM") = CHECK_NUM
            SEQ_NUM = SEQ_NUM + 1
            rowAPTCHCK2("SEQ_NUM") = SEQ_NUM
            rowAPTCHCK2("VEND_CODE") = rowAPTINVH1("VEND_CODE")
            rowAPTCHCK2("INV_NUM") = rowAPTINVH1("INV_NUM")
            rowAPTCHCK2("INV_DATE") = rowAPTINVH1("INV_DATE")
            rowAPTCHCK2("VOUCHER_NO") = rowAPTINVH1("VOUCHER_NO")
            rowAPTCHCK2("INV_AMT_APPLIED") = rowAPTINVH1("INV_AMT")
            rowAPTCHCK2("INV_DISC_TAKEN") = rowAPTINVH1("INV_DISC_AMT")
            CHECK_AMT += Val(rowAPTINVH1("INV_AMT") & "") + Val(rowAPTINVH1("INV_DISC_AMT") & "")
            dst.Tables("APTCHCK2").Rows.Add(rowAPTCHCK2)
        Next

        Dim rowAPTCHCK1 As DataRow = dst.Tables("APTCHCK1").NewRow
        rowAPTCHCK1("BANK_CODE") = BANK_CODE
        rowAPTCHCK1("CHECK_NUM") = CHECK_NUM
        rowAPTCHCK1("CHECK_DATE") = CHECK_DATE
        rowAPTCHCK1("CHECK_AMT") = CHECK_AMT
        rowAPTCHCK1("PYMT_METHOD") = PYMT_METHOD
        rowAPTCHCK1("VEND_CODE") = HFs("VEND_CODE")
        rowAPTCHCK1("VEND_CODE_AP") = VEND_CODE_AP
        rowAPTCHCK1("VEND_ALT_CODE") = VEND_ALT_CODE
        rowAPTCHCK1("OPS_YYYYPP") = ASCMAIN1.CYP
        rowAPTCHCK1("CHECK_STATUS") = "I"
        rowAPTCHCK1("VEND_NAME") = VEND_NAME
        rowAPTCHCK1("INIT_DATE") = DATETIME_STAMP
        rowAPTCHCK1("INIT_OPER") = ASCMAIN1.USER_ID
        rowAPTCHCK1("REGISTER_IND") = "0"
        dst.Tables("APTCHCK1").Rows.Add(rowAPTCHCK1)

        Dim INV_PAYMENTS As Double = Val(dst.Tables("APTCHCK2").Compute("SUM(INV_AMT_APPLIED)", "") & "")
        Dim INV_DISC_TAKEN As Double = Val(dst.Tables("APTCHCK2").Compute("SUM(INV_DISC_TAKEN)", "") & "")

        Dim rowAPTVEND5 = Fill_Record("APTVEND5", VEND_CODE, True)
        rowAPTVEND5.Item("VEND_PAYMENTS_MTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_MTD") & "") + INV_PAYMENTS
        rowAPTVEND5.Item("VEND_PAYMENTS_YTD") = Val(rowAPTVEND5.Item("VEND_PAYMENTS_YTD") & "") + INV_PAYMENTS
        rowAPTVEND5.Item("VEND_DISC_TAKEN_MTD") = Val(rowAPTVEND5.Item("VEND_DISC_TAKEN_MTD") & "") + INV_DISC_TAKEN
        rowAPTVEND5.Item("VEND_DISC_TAKEN_YTD") = Val(rowAPTVEND5.Item("VEND_DISC_TAKEN_YTD") & "") + INV_DISC_TAKEN
        rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_MTD") & "") + 1
        rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") = Val(rowAPTVEND5.Item("VEND_NUM_CHKS_YTD") & "") + 1
        rowAPTVEND5.Item("VEND_LAST_PMT_DATE") = CHECK_DATE
        rowAPTVEND5.Item("VEND_LAST_PMT_AMT") = INV_PAYMENTS

        'If auto_next_check Then
        '    Dim rowGLTBANK1 As DataRow = Fill_Record("GLTBANK1", rowAPTINVH1("BANK_CODE"))
        '    If rowGLTBANK1("BANK_LAST_CHECK_NO") & "" = BANK_LAST_CHECK_NO Then
        '        rowGLTBANK1("BANK_LAST_CHECK_NO") = BANK_NEXT_CHECK_NO
        '    End If
        'End If
        Call Update_Record_TDA("APTCHCK1")
        Call Update_Record_TDA("APTCHCK2")
        Call Update_Record_TDA("APTINVH1")
        Call Update_Record_TDA("APTVEND5")

    End Sub

    Private Sub grdAPTINVH1_DUE_TO_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdAPTINVH1_DUE_TO.InitializeLayout

    End Sub

    Private Sub optMatchedOnly_ValueChanged(sender As Object, e As EventArgs) Handles optMatchedOnly.ValueChanged
        'If Me.SELECTION_NO = 0 Then Exit Sub
        'Setup_Open_AP()
    End Sub

    Function Setup_Open_AP() As String
        Dim BANK_CODE As String = Absx1.txtFor("BANK_CODE").Text
        Dim PYMT_METHOD As String = Absx1.txtFor("PYMT_METHOD").Text

        Dim SQLAP As String = ""

        ' Dim DVW As DataView = dst.Tables("APTINVH1").DefaultView
        If optMatchedOnly.Value = "A" Then
            ' DVW.RowFilter = ""
        Else
            ' DVW.RowFilter = $"BANK_CODE = '{BANK_CODE}' AND INV_PYMT_METHOD = '{PYMT_METHOD}'"
            SQLAP = $" and APTINVH1.BANK_CODE = '{BANK_CODE}' AND APTINVH1.INV_PYMT_METHOD = '{PYMT_METHOD}'"
        End If

        Return SQLAP
    End Function

End Class