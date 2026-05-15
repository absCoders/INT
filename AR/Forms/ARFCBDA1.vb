Imports Infragistics.UltraChart.Resources

Public Class ARFCBDA1
    Dim CUST_CODE As String
    Dim rowARTCUST1 As DataRow
    Dim REASON_CODE As String
    Dim rowARTREAS1 As DataRow
    Dim RYP0 As String
    Dim RYP1 As String
    Dim ARTCBDAA As String
    Dim ARTCBDA1 As String
    Dim sqlARTCBDA1 As String
    Dim RYP0_Legend As String = ""
    Dim RYP1_Legend As String = ""
    Dim MOS As Integer = 0
    Dim tblChart As New DataTable

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            Create_ARTCBDAA("", "")
            TAC.ARCMAIN1.Create_ARTCBDA1(ARTCBDA1, "", "")


            ASCMAIN1.sql = "Select ARTCBDAA.REASON_CODE, ARTREAS1.REASON_DESC" _
                & " from ARTREAS1," & ARTCBDAA & " ARTCBDAA where ARTREAS1.REASON_CODE = ARTCBDAA.REASON_CODE"
            Create_TDA(.Tables.Add, "ARTCBDA2", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCBDAA.*" & vbCrLf _
                & ",ARTCUST1.CUST_NAME,ARTREAS1.REASON_DESC,GLTACCT1.ACCT_DESC" & vbCrLf _
                & " from ARTCUST1,ARTREAS1,GLTACCT1," & ARTCBDAA & " ARTCBDAA" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = ARTCBDAA.CUST_CODE" & vbCrLf _
                & "   and GLTACCT1.ACCT_CODE (+) = ARTCBDAA.ACCT_CODE" & vbCrLf _
                & "   and ARTREAS1.REASON_CODE (+) = ARTCBDAA.REASON_CODE" & vbCrLf _
                & "   and ARTCBDAA.ACTIVITY = :PARM1"
            Create_TDA(.Tables.Add, "ARTCBDAA", "**", 0, False, "V", 3)
            .Tables("ARTCBDAA").Columns.Add("P00", GetType(System.Decimal), "ISNULL(P01,0)+ISNULL(P02,0)+ISNULL(P03,0)+ISNULL(P04,0)+ISNULL(P05,0)+ISNULL(P06,0)+ISNULL(P07,0)+ISNULL(P08,0)+ISNULL(P09,0)+ISNULL(P10,0)+ISNULL(P11,0)+ISNULL(P12,0)")


            ASCMAIN1.sql = "Select ARTCBDAA.*" & vbCrLf _
                & ",ARTCUST1.CUST_NAME,ARTREAS1.REASON_DESC,GLTACCT1.ACCT_DESC" & vbCrLf _
                & " from ARTCUST1,ARTREAS1,GLTACCT1," & ARTCBDAA & " ARTCBDAA" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = ARTCBDAA.CUST_CODE" & vbCrLf _
                & "   and GLTACCT1.ACCT_CODE (+) = ARTCBDAA.ACCT_CODE" & vbCrLf _
                & "   and ARTREAS1.REASON_CODE (+) = ARTCBDAA.REASON_CODE" & vbCrLf _
                & "   and ARTCBDAA.ACTIVITY in ('B','S','A','R')"
            Create_TDA(.Tables.Add, "ARTCBDAX", "**", 0, False, "V", 3)
            With .Tables("ARTCBDAX")
                .Columns.Add("P00", GetType(System.Decimal), "ISNULL(P01,0)+ISNULL(P02,0)+ISNULL(P03,0)+ISNULL(P04,0)+ISNULL(P05,0)+ISNULL(P06,0)+ISNULL(P07,0)+ISNULL(P08,0)+ISNULL(P09,0)+ISNULL(P10,0)+ISNULL(P11,0)+ISNULL(P12,0)")
                .Columns.Add("P00B", GetType(System.Decimal), "IIF(ACTIVITY='B',P00,0)")
                .Columns.Add("P00S", GetType(System.Decimal), "IIF(ACTIVITY='S',P00,0)")
                .Columns.Add("P00A", GetType(System.Decimal), "IIF(ACTIVITY='A',P00,0)")
                .Columns.Add("P00R", GetType(System.Decimal), "IIF(ACTIVITY='R',P00,0)")
            End With


            sqlARTCBDA1 = "Select ARTCBDA1.CUST_CODE CODE_VALUE, ARTCUST1.CUST_NAME DESC_VALUE" _
                & ", Sum (BEG_B) BEG_B, Sum (NEW_B) NEW_B, Sum (APP_B) APP_B, Sum (END_B) END_B" & vbCrLf _
                & ", Sum (BEG_C) BEG_C, Sum (NEW_C) NEW_C, Sum (APP_C) APP_C, Sum (END_C) END_C, Sum (NEW_X) NEW_X" & vbCrLf _
                & " from ARTCUST1," & ARTCBDA1 & " ARTCBDA1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE (+) = ARTCBDA1.CUST_CODE" & vbCrLf _
                & " group by ARTCBDA1.CUST_CODE, ARTCUST1.CUST_NAME"
            ASCMAIN1.sql = sqlARTCBDA1
            Create_TDA(.Tables.Add, "ARTCBDA1", "**", 0, False, "", 1)
            With .Tables("ARTCBDA1")
                .Columns.Add("END_ALL", GetType(System.Decimal), "ISNULL(END_B,0) + ISNULL(END_C,0)")
                .Columns.Add("TOTAL", GetType(System.Decimal), "ISNULL(NEW_C,0) + ISNULL(NEW_X,0)")
                .Columns("DESC_VALUE").MaxLength = -1
            End With


            ASCMAIN1.sql = "Select * from ARTREAS1"
            Create_TDA(.Tables.Add, "ARTREAS1", "**", 0, False)
            With .Tables("ARTREAS1")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                .Columns.Add("CMP")
                .Columns("CMP").DefaultValue = "0"
            End With
            Create_Relation("ARTREAS1", "ARTCBDAX", "REASON_CODE")
            With dst.Tables("ARTCBDAX")
                .Columns.Add("SELR", GetType(System.String), "PARENT(ARTREAS1_ARTCBDAX).SEL")
                .Columns.Add("P00RS", GetType(System.Decimal), "IIF(SELR='1',P00B,0)")
                .Columns.Add("CMPR", GetType(System.String), "PARENT(ARTREAS1_ARTCBDAX).CMP")
                .Columns.Add("P00RC", GetType(System.Decimal), "IIF(CMPR='1',P00B,0)")
            End With


            ASCMAIN1.sql = "Select X.CUST_CODE, ARTCUST1.CUST_NAME from" & vbCrLf _
                & " (Select Distinct CUST_CODE from " & ARTCBDAA & ") X, ARTCUST1" & vbCrLf _
                & " where ARTCUST1.CUST_CODE = X.CUST_CODE"
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False)
            With .Tables("ARTCUST1")
                .Columns.Add("SEL")
                .Columns("SEL").DefaultValue = "0"
                .Columns.Add("CMP")
                .Columns("CMP").DefaultValue = "0"
            End With
            Create_Relation("ARTCUST1", "ARTCBDAX", "CUST_CODE")
            With dst.Tables("ARTCBDAX")
                .Columns.Add("SELC", GetType(System.String), "PARENT(ARTCUST1_ARTCBDAX).SEL")
                .Columns.Add("P00CS", GetType(System.Decimal), "IIF(SELC='1',P00B,0)")
                .Columns.Add("CMPC", GetType(System.String), "PARENT(ARTCUST1_ARTCBDAX).CMP")
                .Columns.Add("P00CC", GetType(System.Decimal), "IIF(CMPC='1',P00B,0)")
            End With


            With .Tables("ARTREAS1")
                .Columns.Add("AMT_ALL", GetType(System.Decimal), "SUM(CHILD(ARTREAS1_ARTCBDAX).P00B)")
                .Columns.Add("AMT_SEL", GetType(System.Decimal), "SUM(CHILD(ARTREAS1_ARTCBDAX).P00CS)")
                .Columns.Add("AMT_CMP", GetType(System.Decimal), "SUM(CHILD(ARTREAS1_ARTCBDAX).P00CC)")
                .Columns.Add("ALL_PCT_TOT", GetType(System.Decimal), "100*ISNULL(AMT_SEL,0)/100")
                .Columns.Add("SEL_PCT_TOT", GetType(System.Decimal), "100*ISNULL(AMT_SEL,0)/100")
                .Columns.Add("SEL_PCT_ALL", GetType(System.Decimal), "IIF(ISNULL(AMT_ALL,0)=0,0,100*ISNULL(AMT_SEL,0)/ISNULL(AMT_ALL,0))")
                .Columns.Add("SEL_PCT_CMP", GetType(System.Decimal), "IIF(ISNULL(AMT_CMP,0)=0,0,100*ISNULL(AMT_SEL,0)/ISNULL(AMT_CMP,0))")
                .Columns.Add("CMP_PCT_TOT", GetType(System.Decimal), "100*ISNULL(AMT_CMP,0)/100")
                .Columns.Add("CMP_PCT_ALL", GetType(System.Decimal), "IIF(ISNULL(AMT_ALL,0)=0,0,100*ISNULL(AMT_CMP,0)/ISNULL(AMT_ALL,0))")
            End With

            With .Tables("ARTCUST1")
                .Columns.Add("AMT_ALL", GetType(System.Decimal), "SUM(CHILD(ARTCUST1_ARTCBDAX).P00B)")
                .Columns.Add("AMT_SEL", GetType(System.Decimal), "SUM(CHILD(ARTCUST1_ARTCBDAX).P00RS)")
                .Columns.Add("AMT_CMP", GetType(System.Decimal), "SUM(CHILD(ARTCUST1_ARTCBDAX).P00RC)")
                .Columns.Add("ALL_PCT_TOT", GetType(System.Decimal), "100*ISNULL(AMT_SEL,0)/100")
                .Columns.Add("SEL_PCT_TOT", GetType(System.Decimal), "100*ISNULL(AMT_SEL,0)/100")
                .Columns.Add("SEL_PCT_ALL", GetType(System.Decimal), "IIF(ISNULL(AMT_ALL,0)=0,0,100*ISNULL(AMT_SEL,0)/ISNULL(AMT_ALL,0))")
                .Columns.Add("SEL_PCT_CMP", GetType(System.Decimal), "IIF(ISNULL(AMT_CMP,0)=0,0,100*ISNULL(AMT_SEL,0)/ISNULL(AMT_CMP,0))")
                .Columns.Add("CMP_PCT_TOT", GetType(System.Decimal), "100*ISNULL(AMT_CMP,0)/100")
                .Columns.Add("CMP_PCT_ALL", GetType(System.Decimal), "IIF(ISNULL(AMT_ALL,0)=0,0,100*ISNULL(AMT_CMP,0)/ISNULL(AMT_ALL,0))")
                .Columns.Add("AMT_S", GetType(System.Decimal), "SUM(CHILD(ARTCUST1_ARTCBDAX).P00S)")
                .Columns.Add("AMT_A", GetType(System.Decimal), "SUM(CHILD(ARTCUST1_ARTCBDAX).P00A)")
                .Columns.Add("AMT_R", GetType(System.Decimal), "SUM(CHILD(ARTCUST1_ARTCBDAX).P00R)")
                .Columns.Add("ALL_PCT_S", GetType(System.Decimal), "IIF(ISNULL(AMT_S,0)=0,0,100*ISNULL(AMT_ALL,0)/ISNULL(AMT_S,0))")
                .Columns.Add("ALL_PCT_A", GetType(System.Decimal), "IIF(ISNULL(AMT_A,0)=0,0,100*ISNULL(AMT_ALL,0)/ISNULL(AMT_A,0))")
                .Columns.Add("ALL_PCT_R", GetType(System.Decimal), "IIF(ISNULL(AMT_R,0)=0,0,100*ISNULL(AMT_ALL,0)/ISNULL(AMT_R,0))")
                .Columns.Add("SEL_PCT_S", GetType(System.Decimal), "IIF(ISNULL(AMT_S,0)=0,0,100*ISNULL(AMT_SEL,0)/ISNULL(AMT_S,0))")
                .Columns.Add("SEL_PCT_A", GetType(System.Decimal), "IIF(ISNULL(AMT_A,0)=0,0,100*ISNULL(AMT_SEL,0)/ISNULL(AMT_A,0))")
                .Columns.Add("SEL_PCT_R", GetType(System.Decimal), "IIF(ISNULL(AMT_R,0)=0,0,100*ISNULL(AMT_SEL,0)/ISNULL(AMT_R,0))")
                .Columns.Add("CMP_PCT_S", GetType(System.Decimal), "IIF(ISNULL(AMT_S,0)=0,0,100*ISNULL(AMT_CMP,0)/ISNULL(AMT_S,0))")
                .Columns.Add("CMP_PCT_A", GetType(System.Decimal), "IIF(ISNULL(AMT_A,0)=0,0,100*ISNULL(AMT_CMP,0)/ISNULL(AMT_A,0))")
                .Columns.Add("CMP_PCT_R", GetType(System.Decimal), "IIF(ISNULL(AMT_R,0)=0,0,100*ISNULL(AMT_CMP,0)/ISNULL(AMT_R,0))")
            End With


            With tblChart
                .Columns.Add("COLUMN_NAME")
                .Columns.Add("COLUMN_CAPTION")
                .Columns.Add("XYD")
            End With
        End With

        Fill_Records("ARTREAS1")
        With dst.Tables("ARTREAS1").Rows
            .Add(New String() {"**", "N/A"})
            .Add(New String() {"*S", "Gross Sales"})
            .Add(New String() {"*A", "Cartons Shipped"})
            .Add(New String() {"*R", "Retail Sales"})
        End With



        grdARTCBDAA.DataSource = dst.Tables("ARTCBDAA")
        grdARTCBDA1.DataSource = dst.Tables("ARTCBDA1")
        grdARTCBDA2.DataSource = dst.Tables("ARTCBDA2")

        grdChart.DataSource = tblChart
        For Each C As String In New String() {"AMT_S", "AMT_A", "AMT_R", "AMT_ALL", "AMT_SEL", "AMT_CMP", "SEL_PCT_S", "SEL_PCT_A", "SEL_PCT_R"}
            tblChart.Rows.Add(New String() {C, grdARTCUST1.DisplayLayout.Bands(0).Columns(C).Header.Caption, ""})
        Next


        grdARTCBDAX.DataSource = dst.Tables("ARTCBDAX")
        grdARTCUST1.DataSource = dst.Tables("ARTCUST1")
        grdARTREAS1.DataSource = dst.Tables("ARTREAS1")
        Dim dvw As DataView = DirectCast(grdARTREAS1.DataSource, DataTable).DefaultView
        dvw.RowFilter = "REASON_CODE <> '**' and REASON_CODE <> '*S' and REASON_CODE <> '*A' and REASON_CODE <> '*R'"

        Create_Summary(grdARTCBDAA, "CUST_CODE", "Count")
        Create_Summary(grdARTCBDAA, "REASON_CODE", "Count")

        For I As Integer = 0 To 12
            Dim C As String = "P" & Format(I, "00")
            Create_Summary(grdARTCBDAA, C)
        Next I

        grdARTCBDAX.DisplayLayout.Bands(0).Columns("ACTIVITY").Header.Fixed = True

        For Each GRD As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCBDAA, grdARTCBDAX}
            With GRD.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME", "REASON_CODE", "REASON_DESC", "P00"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    Dim COLUMN_NAME As String = gcol.Key
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                Next
            End With
        Next

        For Each GRD As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCUST1, grdARTREAS1}
            GRD.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
            With GRD.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"SEL", "CMP"}
                    .Columns(COLUMN_NAME).Header.Fixed = True
                Next
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    Dim COLUMN_NAME As String = gcol.Key
                    .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                    If COLUMN_NAME = "SEL" Or COLUMN_NAME = "CMP" Then
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.AllowEdit
                        Create_Summary(GRD, COLUMN_NAME)
                    Else
                        .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit

                        If COLUMN_NAME = "CUST_CODE" Or COLUMN_NAME = "REASON_CODE" Then
                            .Columns(COLUMN_NAME).Header.Fixed = True
                            Create_Summary(GRD, COLUMN_NAME, "Count")
                        ElseIf COLUMN_NAME = "AMT_ALL" Or COLUMN_NAME = "AMT_SEL" Or COLUMN_NAME = "AMT_CMP" _
                            Or COLUMN_NAME = "AMT_S" Or COLUMN_NAME = "AMT_A" Or COLUMN_NAME = "AMT_R" Then
                            Create_Summary(GRD, COLUMN_NAME)
                        ElseIf COLUMN_NAME = "ALL_PCT_TOT" Or COLUMN_NAME = "SEL_PCT_TOT" Or COLUMN_NAME = "CMP_PCT_TOT" Then
                            Create_Summary(GRD, COLUMN_NAME)
                        End If
                        If COLUMN_NAME = "AMT_S" Or COLUMN_NAME = "AMT_A" Or COLUMN_NAME = "AMT_R" Then
                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                            .Columns(COLUMN_NAME).CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                        ElseIf COLUMN_NAME = "SEL_PCT_CMP" Then
                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Gold
                        ElseIf COLUMN_NAME = "AMT_ALL" Or COLUMN_NAME.StartsWith("ALL_PCT") Then
                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                        ElseIf COLUMN_NAME = "AMT_SEL" Or COLUMN_NAME.StartsWith("SEL_PCT") Then
                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.Orange
                        Else
                            .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        End If
                    End If
                Next
            End With
        Next


        With grdARTCBDA1.DisplayLayout.Bands(0)
            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = gcol.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White
                .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
            Next
            Dim g As UltraWinGrid.UltraGridGroup

            g = .Groups.Add("CODES")
            g.Header.Fixed = True
            g.Header.Caption = "Codes"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.LightGreen
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"CODE_VALUE", "DESC_VALUE"}
                .Columns(COLUMN_NAME).Group = g
            Next
            Create_Summary(grdARTCBDA1, "CODE_VALUE", "Count")

            g = .Groups.Add("CB")
            g.Header.Caption = "AR Chargebacks and CRs On/Account"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.LightBlue
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"BEG_B", "NEW_B", "APP_B", "END_B"}
                .Columns(COLUMN_NAME).Group = g
                .Columns(COLUMN_NAME).Format = "#,##0.00"
                .Columns(COLUMN_NAME).Width = 100
                Create_Summary(grdARTCBDA1, COLUMN_NAME)
            Next

            g = .Groups.Add("CR")
            g.Header.Caption = "Misc Charges and Credits"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.Orange
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"BEG_C", "NEW_C", "APP_C", "END_C"}
                .Columns(COLUMN_NAME).Group = g
                .Columns(COLUMN_NAME).Format = "#,##0.00"
                .Columns(COLUMN_NAME).Width = 100
                Create_Summary(grdARTCBDA1, COLUMN_NAME)
            Next

            g = .Groups.Add("DED")
            g.Header.Caption = "Totals"
            With g.Header.Appearance
                .TextHAlign = HAlign.Center
                .BackColor = Drawing.Color.White
                .BackColor2 = Drawing.Color.Violet
                .BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            For Each COLUMN_NAME As String In New String() {"END_ALL", "NEW_X", "TOTAL"}
                .Columns(COLUMN_NAME).Group = g
                .Columns(COLUMN_NAME).Format = "#,##0.00"
                .Columns(COLUMN_NAME).Width = 100
                Create_Summary(grdARTCBDA1, COLUMN_NAME)
            Next
        End With

        ASCMAIN1.sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2" _
            & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -60) & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        cbeYP0.DataSource = New DataView(DT, "", "OPS_YYYYPP", DataViewRowState.CurrentRows)
        cbeYP1.DataSource = New DataView(DT, "", "OPS_YYYYPP", DataViewRowState.CurrentRows)

        tab1.Tabs("Detail").Visible = False

        ASCMAIN1.Add_Value_List(grdARTCBDAX, "ACTIVITY", , _
        New String() {":" _
         , "B:Chargeback" _
         , "S:Gross Sales" _
         , "A:Cartons Shipped" _
         , "R:Retail Sales" _
                     })

        ASCMAIN1.Add_Value_List(grdChart, "XYD", , _
        New String() {":" _
         , "X:X-Axis" _
         , "Y:Y-Axis" _
         , "D:Diameter" _
                     })

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "View"
                If Absx1.txtFor("CUST_CODE").Text <> "" Then
                    Validate_Code("CUST_CODE")
                    If EMsg = "" Then
                        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
                        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                        If rowARTCUST1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Customer Code " & CUST_CODE
                        Else
                            REASON_CODE = ""
                            Absx1.txtFor("REASON_CODE").Text = ""
                        End If
                    End If
                ElseIf Absx1.txtFor("REASON_CODE").Text <> "" Then
                    Validate_Code("REASON_CODE")
                    If EMsg = "" Then
                        REASON_CODE = Absx1.txtFor("REASON_CODE").Text
                        rowARTREAS1 = LookUp("ARTREAS1", REASON_CODE)
                        If rowARTREAS1 Is Nothing Then
                            EMsg &= vbCr & "Invalid Reason Code " & REASON_CODE
                        Else
                            CUST_CODE = ""
                            Absx1.txtFor("CUST_CODE").Text = ""
                        End If
                    End If
                End If

                If EMsg = "" Then
                    If Not Load_Periods() Then
                        Exit Sub
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                If optShowRollForward.Value = "C" Then
                    optShowRollForward.Value = "R"
                Else
                    optShowRollForward.Value = "C"
                End If
                Mode_Settings(False)

            Case "Print"
                Print_Record()
        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("View").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    .Items("Print").Settings.Enabled = iScreenMode

                    .Items("Print").Visible = False ' (InquiryMode Or EntryMode = "V")
                End With
                .Groups("Period Range").Visible = Not ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        Set_Read_Only_for_ctl(cbeYP0, ScreenMode)
        Set_Read_Only_for_ctl(cbeYP1, ScreenMode)

        splARTCBDA1.Visible = ScreenMode
        tab0.Visible = Not ScreenMode

        If ScreenMode Then
            grdARTCBDA1.Parent = splARTCBDA1.Panel1
            grdARTCBDAA.Parent = tab1.Tabs("Activity").TabPage
            'grdARTCBDAA.DisplayLayout.Bands(0).Columns("ACTIVITY").Hidden = False
            'grdARTCBDAA.DisplayLayout.Bands(0).SortedColumns.Clear()
            'grdARTCBDAA.DisplayLayout.Bands(0).SortedColumns.Add("ACTIVITY", False, True)
            'grdARTCBDAA.DisplayLayout.GroupByBox.Hidden = True
        Else
            Clear_Record()
            grdARTCBDA1.Parent = tab0.Tabs("Roll Forward").TabPage
            grdARTCBDAA.Parent = tab0.Tabs("Activity").TabPage
            'grdARTCBDAA.DisplayLayout.Bands(0).Columns("ACTIVITY").Hidden = True

        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"ARTCBDA1", "ARTCBDA2", "ARTCBDAA"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("REASON_CODE").Text = ""

        If cbeYP0.Value & "" = "" Then
            cbeYP0.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -11)
            cbeYP1.Value = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, 0)
            Load_ARTCBDAA()
            Load_ARTCBDA1()
        Else
            Setup_ARTCBDAA()
            Setup_ARTCBDA1()
        End If

        Setup_tab0()

        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        'Load_ARTCBDAA()
        Setup_ARTCBDAA()

        If optShowRollForward.Value = IIf(REASON_CODE <> "", "C", "R") Then
            Setup_ARTCBDA1()
        Else
            optShowRollForward.Value = IIf(REASON_CODE <> "", "C", "R")
        End If

        EnforceConstraints(True)

        Setup_tab0()

        SETUP_grdARTCBDA2()

        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()
        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdARTCUST1, "SBBBB", "Show Filter", "Set All Sel", "Clear All Sel", "Set All Cmp", "Clear All Cmp")
        Load_Popup_Menu(grdARTREAS1, "SBBBB", "Show Filter", "Set All Sel", "Clear All Sel", "Set All Cmp", "Clear All Cmp")
        Load_Popup_Menu(grdARTCBDAA, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdARTCBDAX, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdARTCBDA1, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Inquiry")
        Load_Popup_Menu(grdARTCBDA2, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Customer Inquiry")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            ' e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdARTCBDA1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Set All Sel", "Clear All Sel", "Set All Cmp", "Clear All Cmp"
                Dim C As String = "SEL"
                If e.Tool.Key.EndsWith("Cmp") Then C = "CMP"
                Dim TBL As DataTable = DirectCast(grd.DataSource, DataTable)
                For Each row As DataRow In TBL.Select("")
                    row.Item(C) = IIf(e.Tool.Key.StartsWith("Set"), "1", "0")
                Next
            Case Else

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Customer Inquiry"
                Dim CUST_CODE As String
                If grd.Name = "grdARTCBDA1" Then
                    CUST_CODE = grd.ActiveRow.Cells("CODE_VALUE").Value
                Else
                    CUST_CODE = grd.ActiveRow.Cells("CUST_CODE").Value
                End If
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 IsNot Nothing Then
                    Context_Launch("Select", CUST_CODE, e.Tool.Key, "ARFCINQ1")
                End If
        End Select
    End Sub
#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        Absx1.txtFor("REASON_CODE").Text = ""
                        Click_Command("View")
                    End If
                End If
            Case "REASON_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    If Absx1.txtFor("REASON_CODE").Text <> "" Then
                        Absx1.txtFor("CUST_CODE").Text = ""
                        Click_Command("View")
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "CUST_CODE"
                Absx1.txtFor("REASON_CODE").Text = ""
                Click_Command("View")
            Case "REASON_CODE"
                Absx1.txtFor("CUST_CODE").Text = ""
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr Is Nothing Then
                            Absx1.txtFor("CUST_CODE").Text = ""
                        End If
                    End If
                End If
            Case "REASON_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("REASON_CODE").Text <> "" Then
                        LookUp("ARTREAS1", Absx1.txtFor("REASON_CODE").Text)
                        If cdr Is Nothing Then
                            Absx1.txtFor("REASON_CODE").Text = ""
                        End If
                    End If
                End If
        End Select
    End Sub
#End Region

#Region "grdARTCBDA1"

    Private Sub grdARTCBDA1_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdARTCBDA1.AfterRowActivate
        If ScreenMode Then SETUP_grdARTCBDA2()
    End Sub

#End Region

    Sub SETUP_grdARTCBDA2()
        If grdARTCBDA1.ActiveRow Is Nothing OrElse Not grdARTCBDA1.ActiveRow.IsDataRow Then
            splARTCBDA1.Panel2Collapsed = True
        Else
            grdARTCBDA2.Text = "Roll Forward Details for " & ""
            splARTCBDA1.Panel2Collapsed = False
        End If
    End Sub

    Private Sub grdARTCBDAA_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCBDAA.DoubleClickRow
        If e.Row.IsDataRow Then
            Absx1.txtFor("CUST_CODE").Text = e.Row.Cells("CUST_CODE").Value
            Absx1.txtFor("REASON_CODE").Text = ""
            Click_Command("View")
        End If
    End Sub

    Sub Print_Record()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Printing Report")

        'Print_Report_Begin()
        'CR_params.Add("NOTES", "1")
        'Generate_Report("BMRLIST1", "Bill of Materials", "")
        'Print_Report_End()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Function Load_Periods() As Boolean
        RYP0 = cbeYP0.Value
        RYP1 = cbeYP1.Value

        Dim P As Integer = 11
        If RYP0 <> "" Then P = ASCMAIN1.Period_Diff(RYP0, RYP1)

        If P < 0 Or P > 11 Then
            MsgBox("Periods range may span from 1 to 12 months")
            Return False
        Else

            RYP0_Legend = ""
            RYP1_Legend = ""
            MOS = 0

            For Each GRD As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdARTCBDAA, grdARTCBDAX}
                With GRD.DisplayLayout.Bands(0)
                    For I As Integer = 1 To 12
                        Dim RYP As String = ASCMAIN1.Period_Calc(RYP0, (I - 1))
                        Dim C As String = "P" & Format(I, "00")
                        If RYP <= RYP1 Then
                            .Columns(C).Hidden = False
                            Dim LEGEND = ASCMAIN1.Get_Legend(RYP)
                            If I = 1 Then RYP0_Legend = LEGEND
                            If RYP = RYP1 Then RYP1_Legend = LEGEND
                            .Columns(C).Header.Caption = Mid(LEGEND, 10, 6)
                            MOS += 1
                        Else
                            .Columns(C).Hidden = True
                        End If
                    Next
                End With
            Next

            Return True
        End If
    End Function

    Sub Load_ARTCBDAA()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Activity")

        Load_Periods()
        Create_ARTCBDAA(RYP0, RYP1)
        Setup_ARTCBDAA()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Sub Load_ARTCBDA1()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Roll Forward")

        TAC.ARCMAIN1.Create_ARTCBDA1(ARTCBDA1, RYP0, RYP1)
        Setup_ARTCBDA1()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Sub Create_ARTCBDAA(RYP0 As String, RYP1 As String)

        Dim sqlSum As String = ""
        For I As Integer = 1 To 12
            Dim RYP As String = ASCMAIN1.Period_Calc(RYP0, (I - 1))
            Dim C As String = "P" & Format(I, "00")
            If RYP <= RYP1 Or ARTCBDAA = "" Then
                sqlSum &= ", SUM (DECODE(ARTPYMT1.OPS_YYYYPP,'" & RYP & "',ARTPYMT5.GL_DIST_AMT,0)) " & C & vbCrLf
            Else
                sqlSum &= ", 0 " & C & vbCrLf
            End If
        Next

        ASCMAIN1.sql = "Select NVL(ARTPYMT2.CUST_CODE,'NON-AR') CUST_CODE, ARTPYMT5.REASON_CODE, ARTPYMT5.ACCT_CODE" & vbCrLf _
            & ", DECODE(ARTPYMT5.CHARGEBACK_IND,'1','B','D') ACTIVITY" & vbCrLf _
            & sqlSum _
            & "  from ARTPYMT1,ARTPYMT2,ARTPYMT5" & vbCrLf _
            & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
            & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
            & "   and ARTPYMT1.OPS_YYYYPP <= '" & RYP1 & "'"
        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and ARTPYMT2.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        End If
        If Absx1.txtFor("REASON_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and ARTPYMT5.REASON_CODE = '" & Absx1.txtFor("REASON_CODE").Text & "'"
        End If
        If ARTCBDAA = "" Then
            ASCMAIN1.sql &= "   and ROWNUM < 1"
        End If
        ASCMAIN1.sql &= " group by NVL(ARTPYMT2.CUST_CODE,'NON-AR'), ARTPYMT5.REASON_CODE, ARTPYMT5.ACCT_CODE" & vbCrLf _
            & ", DECODE(ARTPYMT5.CHARGEBACK_IND,'1','B','D')"

        Dim sqlARTPYMT5rc As String = "SELECT ARTPYMT5.PYMT_BATCH_NO, ARTPYMT5.PYMT_BATCH_LNO" & vbCrLf _
                                        & ", NVL(RECODED.NEW_VALUE,ARTPYMT5.REASON_CODE) REASON_CODE" & vbCrLf _
                                        & " , ARTPYMT5.ACCT_CODE, ARTPYMT5.CHARGEBACK_IND, ARTPYMT5.GL_DIST_AMT" & vbCrLf _
                                        & " FROM ARTPYMT5, " & vbCrLf _
                                        & " (Select " & vbCrLf _
                                        & "   REGEXP_SUBSTR(KEY_VALUE, '[^:]+', 1, 1) CUST_CODE" & vbCrLf _
                                        & " , REGEXP_SUBSTR(KEY_VALUE, '[^:]+', 1, 2) INV_TYPE" & vbCrLf _
                                        & " , REGEXP_SUBSTR(KEY_VALUE, '[^:]+', 1, 3) INV_NUM" & vbCrLf _
                                        & " , OLD_VALUE, NEW_VALUE, INIT_DATE, init_date_max" & vbCrLf _
                                        & " from (" & vbCrLf _
                                        & " Select KEY_VALUE, OLD_VALUE, NEW_VALUE, INIT_DATE, MAX(INIT_DATE) over (partition by KEY_VALUE) INIT_DATE_MAX" & vbCrLf _
                                        & " from ASTAUDT1 where TABLE_NAME = 'ARTOPEN1' and COLUMN_NAME = 'REASON_CODE'" & vbCrLf _
                                        & " ) where INIT_DATE = INIT_DATE_MAX) RECODED" & vbCrLf _
                                        & " WHERE ARTPYMT5.CHARGEBACK_NO = RECODED.INV_NUM (+)" & vbCrLf _
                                        & " AND ARTPYMT5.INV_TYPE_CB = RECODED.INV_TYPE (+)"

        If chkOnAccount.Checked Then
            sqlARTPYMT5rc &= vbCrLf & " AND ARTPYMT5.INV_TYPE_CB <> 'O'"
        End If

        Dim sqlRecoded As String = Replace(ASCMAIN1.sql,
                                           "from ARTPYMT1,ARTPYMT2,ARTPYMT5", "from ARTPYMT1,ARTPYMT2,(" & sqlARTPYMT5rc & ") ARTPYMT5")

        If ARTCBDAA = "" Then
            ARTCBDAA = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTCBDAA)
            ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & sqlRecoded)
            ASCMAIN1.sql = Replace(Replace(Replace(Replace(
                                            ASCMAIN1.sql, "DECODE(ARTPYMT5.CHARGEBACK_IND,'1','B','D')", "'D'"),
                                            "ARTPYMT5.REASON_CODE", "'**'"),
                                            "ARTPYMT5", "ARTPYMT4"),
                                            "group by", vbCrLf & "   and ARTPYMT4.ACCT_CODE <> '131200' group by")
            ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & ASCMAIN1.sql)

            ASCMAIN1.sql = Replace(Replace(Replace(ASCMAIN1.sql, "'**'", "'J05'"), "ARTPYMT4.ACCT_CODE <> '131200'", "ARTPYMT4.ACCT_CODE = '131200'"), "'D'", "'B'")
            ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & ASCMAIN1.sql)

        End If

        ASCMAIN1.sql = "Select SOTINVH1.CUST_CODE, SOTINVH1.REASON_CODE, ARTREAS1.ACCT_CODE" & vbCrLf _
           & ", 'C' ACTIVITY" & vbCrLf _
           & Replace(Replace(Replace(Replace(sqlSum, "ARTPYMT1", "SOTINVH1"), "OPS_YYYYPP", "ORDR_YYYYPP_UPDATED"), "ARTPYMT5", "SOTINVH1"), "GL_DIST_AMT", "INV_SALES") _
           & "  from SOTINVH1,ARTREAS1" & vbCrLf _
           & " where SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" & vbCrLf _
           & "   and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'" & vbCrLf _
           & "   and SOTINVH1.ORDR_TYPE_CODE in ('TOP','DIF')" & vbCrLf _
           & "   and ARTREAS1.REASON_CODE (+) = SOTINVH1.REASON_CODE"
        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        End If
        If Absx1.txtFor("REASON_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and SOTINVH1.REASON_CODE = '" & Absx1.txtFor("REASON_CODE").Text & "'"
        End If
        ASCMAIN1.sql &= " group by SOTINVH1.CUST_CODE, SOTINVH1.REASON_CODE, ARTREAS1.ACCT_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & ASCMAIN1.sql)

        Dim allCustomers As Boolean = chkAllCustomers.Checked

        Dim custSql As String = IIf(allCustomers, "", " and SOTINVH1.CUST_CODE in (Select Distinct CUST_CODE from " & ARTCBDAA & ")")
        Dim custFrom As String = IIf(allCustomers, "SOTINVH1,SOTINVH2,ARTCUST1,ARTCUST2,ICTITEM1", "SOTINVH1")
        Dim custJoin As String = IIf(allCustomers, "ARTCUST1.CUST_CODE (+) = SOTINVH2.CUST_CODE " & vbCrLf _
                                     & " AND ARTCUST2.CUST_CODE (+) = SOTINVH2.CUST_CODE " & vbCrLf _
                                     & " AND ARTCUST2.CUST_STORE_NO (+) = SOTINVH2.CUST_STORE_NO " & vbCrLf _
                                     & " AND ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE " & vbCrLf _
                                     & " AND SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE " & vbCrLf _
                                     & " AND SOTINVH1.INV_NO = SOTINVH2.INV_NO " & vbCrLf & " AND ", "")
        Dim custWhere As String = IIf(allCustomers, " and NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0) <> 0", "")

        For Each C As String In New String() {"INV_SALES", "INV_CARTONS"}
            Dim ACTIVITY As String = IIf(C = "INV_SALES", "S", "A")
            Dim dataField As String = IIf((C = "INV_SALES" And allCustomers), "NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0) ", C)
            ASCMAIN1.sql = "Select SOTINVH1.CUST_CODE, '*" & ACTIVITY & "' REASON_CODE, '*' ACCT_CODE" & vbCrLf _
             & ", '" & ACTIVITY & "' ACTIVITY" & vbCrLf _
             & Replace(Replace(Replace(Replace(sqlSum, "ARTPYMT1", "SOTINVH1"), "OPS_YYYYPP", "ORDR_YYYYPP_UPDATED"), "ARTPYMT5", "SOTINVH1"), "GL_DIST_AMT", dataField) _
             & " from " & custFrom & vbCrLf _
             & " where " & custJoin & "   SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'" & vbCrLf _
             & "   and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'" & vbCrLf _
             & "   and SOTINVH1.INV_TYPE = 'I'" & vbCrLf _
             & custWhere & vbCrLf

            If Absx1.txtFor("CUST_CODE").Text <> "" Then
                ASCMAIN1.sql &= "   and SOTINVH1.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
            Else
                ASCMAIN1.sql &= custSql & vbCrLf
            End If
            ASCMAIN1.sql &= " group by SOTINVH1.CUST_CODE"

            Dim scSql As String = Replace(ASCMAIN1.sql, "SOTINVH1.NVL", "NVL")
            ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & scSql)
        Next

        ASCMAIN1.sql = "Select RSTRETL2.CUST_CODE, '*R' REASON_CODE, '*' ACCT_CODE" & vbCrLf _
        & ", 'R' ACTIVITY" & vbCrLf _
        & Replace(Replace(Replace(Replace(sqlSum, "ARTPYMT1", "RSTRETL2"), "OPS_YYYYPP", "OPS_YYYYPP"), "ARTPYMT5", "RSTRETL2"), "GL_DIST_AMT", "RETAIL_SALES") _
        & "  from RSTRETL2" & vbCrLf _
        & " where RSTRETL2.OPS_YYYYPP >= '" & RYP0 & "'" & vbCrLf _
        & "   and RSTRETL2.OPS_YYYYPP <= '" & RYP1 & "'" & vbCrLf

        If Absx1.txtFor("CUST_CODE").Text <> "" Then
            ASCMAIN1.sql &= "   and RSTRETL2.CUST_CODE = '" & Absx1.txtFor("CUST_CODE").Text & "'"
        Else
            ASCMAIN1.sql &= "   and RSTRETL2.CUST_CODE in (Select Distinct CUST_CODE from " & ARTCBDAA & ")" & vbCrLf
        End If
        ASCMAIN1.sql &= " group by RSTRETL2.CUST_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & ARTCBDAA & " " & ASCMAIN1.sql)

    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        If Load_Periods() Then
            Load_ARTCBDAA()
            Load_ARTCBDA1()
        End If
    End Sub

    Sub Setup_ARTCBDA1()

        Dim SQLW As String = ""

        If optShowRollForward.Value = "C" Then
            ASCMAIN1.sql = sqlARTCBDA1
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption = "Customer"
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = "Name"
            If (EntryMode = "V") Then SQLW = " ARTCBDA1.REASON_CODE = '" & REASON_CODE & "'"
        Else
            ASCMAIN1.sql = _
                Replace( _
                Replace( _
                Replace(sqlARTCBDA1, _
                    "CUST_CODE", "REASON_CODE"), _
                    "CUST_NAME", "REASON_DESC"), _
                    "ARTCUST1", "ARTREAS1")
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("CODE_VALUE").Header.Caption = "Reason"
            grdARTCBDA1.DisplayLayout.Bands(0).Columns("DESC_VALUE").Header.Caption = "Description"
            If (EntryMode = "V") Then SQLW = " ARTCBDA1.CUST_CODE = '" & CUST_CODE & "'"
        End If

        If SQLW <> "" Then ASCMAIN1.sql = Replace(ASCMAIN1.sql, " where ", " where " & SQLW & " and ")
        EnforceConstraints(False)
        Fill_Records("ARTCBDA1", "", True, ASCMAIN1.sql)
        EnforceConstraints(True)
        Sort_grdColumns(grdARTCBDA1, "CODE_VALUE")

        grdARTCBDA1.Text = "Roll Forward" & _
            IIf(EntryMode = "V", " for " & IIf(optShowRollForward.Value = "C", REASON_CODE, CUST_CODE), "") _
            & ": " & RYP0_Legend & " thru " & RYP1_Legend
    End Sub

    Sub Setup_ARTCBDAA()

        'If EntryMode = "V" Then
        '    Fill_Records("ARTCBDAA", "D")
        '    Fill_Records("ARTCBDAA", "B", False)
        '    Fill_Records("ARTCBDAA", "C", False)
        '    Fill_Records("ARTCBDAA", "S", False)
        '    Fill_Records("ARTCBDAA", "A", False)
        '    Fill_Records("ARTCBDAA", "R", False)
        'Else
        '    Fill_Records("ARTCBDAA", optShowActivity.Value)
        'End If

        '  ASCMAIN1.sql = ASCMAIN1.Flattened_List("CUST_CODE", "NVL(P01,0)+NVL(P02,0)", ARTCBDAA)

        Dim sqlT As String = ""
        For i As Integer = 1 To 12
            sqlT &= "+NVL(P" & Format(i, "00") & ",0)"
        Next
        sqlT = Mid(sqlT, 2)

        Dim sqlX As String = ""
        ASCMAIN1.sql = "Select Distinct REASON_CODE CODE from " & ARTCBDAA & " where ACTIVITY = 'B'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "CODE")
            Dim CODE As String = row.Item(0)
            sqlX &= ", Sum (Decode (REASON_CODE,'" & CODE & "'," & sqlT & ",0)) " & Chr(34) & CODE & Chr(34) & vbCrLf
        Next
        ASCMAIN1.sql = "Select CUST_CODE" & sqlX & " from " & ARTCBDAA & " where ACTIVITY = 'B' group by CUST_CODE"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        tbl.PrimaryKey = New DataColumn() {tbl.Columns("CUST_CODE")}

        'If dst.Tables.Contains("PIVOT") Then
        '    dst.Relations.Remove("ARTCUST1_PIVOT")
        '    dst.Tables("PIVOT").Constraints.Clear()
        '    dst.Tables.Remove("PIVOT")
        'End If
        'tbl.TableName = "PIVOT"
        'dst.Tables.Add(tbl)
        'Create_Relation("ARTCUST1", "PIVOT", "CUST_CODE")

        'tbl.Columns.Add("AMT_S", GetType(System.Decimal), "PARENT.AMT_S")
        'tbl.Columns.Add("AMT_A", GetType(System.Decimal), "PARENT.AMT_A")
        'tbl.Columns.Add("AMT_R", GetType(System.Decimal), "PARENT.AMT_R")


        tbl.Columns.Add("AMT_S", GetType(System.Decimal))
        tbl.Columns.Add("AMT_A", GetType(System.Decimal))
        tbl.Columns.Add("AMT_R", GetType(System.Decimal))

        If grdARTCBDAC.DisplayLayout.Bands(0).Summaries.Count > 0 Then
            grdARTCBDAC.DisplayLayout.Bands(0).Summaries.Clear()
        End If

        grdARTCBDAC.DataSource = Nothing
        grdARTCBDAC.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
        grdARTCBDAC.DataSource = tbl
        'ASCMAIN1.grdInitializeLayout(grdARTCBDAC)

        With grdARTCBDAC.DisplayLayout.Bands(0)
            For Each C As String In New String() {"CUST_CODE", "AMT_S", "AMT_A", "AMT_R"}
                .Columns(C).Header.Fixed = True
                If C = "CUST_CODE" Then
                Else
                    .Columns(C).Width = 100
                    .Columns(C).Format = "#,##0"
                End If
            Next
            .Columns("CUST_CODE").Header.Caption = "Customer"
            .Columns("AMT_S").Header.Caption = "$GrossSls"
            .Columns("AMT_A").Header.Caption = "#Cartons"
            .Columns("AMT_R").Header.Caption = "$Retail"
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                Dim COLUMN_NAME As String = GCOL.Key
                .Columns(COLUMN_NAME).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Columns(COLUMN_NAME).Header.Appearance.BackColor = Drawing.Color.White

                If COLUMN_NAME = "CUST_CODE" Then
                    Create_Summary(grdARTCBDAC, COLUMN_NAME, "Count")
                    .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGray
                Else
                    Create_Summary(grdARTCBDAC, COLUMN_NAME)
                    If COLUMN_NAME = "AMT_S" Or COLUMN_NAME = "AMT_A" Or COLUMN_NAME = "AMT_R" Then
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                    Else
                        .Columns(COLUMN_NAME).Header.Appearance.BackColor2 = Drawing.Color.LightBlue
                    End If

                End If
            Next
        End With





        EnforceConstraints(False)
        Fill_Records("ARTCBDAX")
        Sort_grdColumns(grdARTCBDAA, "ACTIVITY,CUST_CODE,REASON_CODE")
        Setup_ARTCBDAX()

        Fill_Records("ARTCUST1")
        dst.Tables("ARTCUST1").Rows.Add(New String() {"NON-AR", "Non-AR Customer"})
        Sort_grdColumns(grdARTCUST1, "CUST_CODE")
        Sort_grdColumns(grdARTREAS1, "REASON_CODE")
        EnforceConstraints(True)


        For Each rowARTCUST1 As DataRow In dst.Tables("ARTCUST1").Select("")
            Dim CUST_CODE As String = rowARTCUST1.Item("CUST_CODE")
            rowARTCUST1.Item("CMP") = "1"
            Dim row As DataRow = tbl.Rows.Find(CUST_CODE)
            If row Is Nothing Then
                row = tbl.NewRow
                row.Item("CUST_CODE") = CUST_CODE
                tbl.Rows.Add(row)
            End If
            row.Item("AMT_S") = rowARTCUST1.Item("AMT_S")
            row.Item("AMT_A") = rowARTCUST1.Item("AMT_A")
            row.Item("AMT_R") = rowARTCUST1.Item("AMT_R")
        Next

        For Each rowARTREAS1 As DataRow In dst.Tables("ARTREAS1").Select("")
            rowARTREAS1.Item("CMP") = "1"
        Next

        Fill_Records("ARTCBDAA", optShowActivity.Value)
        Sort_grdColumns(grdARTCBDAA, "CUST_CODE,REASON_CODE")

        With grdARTCBDAA.DisplayLayout.Bands(0)
            .Columns("CUST_CODE").Hidden = (EntryMode = "V") And (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("CUST_NAME").Hidden = (EntryMode = "V") And (Absx1.txtFor("CUST_CODE").Text <> "")
            .Columns("REASON_CODE").Hidden = (EntryMode = "V") And (Absx1.txtFor("REASON_CODE").Text <> "")
            .Columns("REASON_DESC").Hidden = (EntryMode = "V") And (Absx1.txtFor("REASON_CODE").Text <> "")
        End With

        Dim DVW As DataView = DirectCast(grdARTCBDAA.DataSource, DataTable).DefaultView
        If EntryMode = "V" Then
            If CUST_CODE <> "" Then
                DVW.RowFilter = "CUST_CODE = '" & CUST_CODE & "'"
            Else
                DVW.RowFilter = "REASON_CODE = '" & REASON_CODE & "'"
            End If
        Else
            DVW.RowFilter = ""
        End If

        grdARTCBDAA.Visible = True
        grdARTCBDAA.Text = optShowActivity.Text & " Activity for the " & CStr(MOS) & " Months " & RYP0_Legend & " to " & RYP1_Legend

        grdARTCBDAX.Visible = True
        grdARTCBDAX.Text = optShowActivity.Text & " Activity for the " & CStr(MOS) & " Months " & RYP0_Legend & " to " & RYP1_Legend


        Dim TOT As Decimal = Val(dst.Tables("ARTCUST1").Compute("SUM(AMT_ALL)", "") & "")

        With dst.Tables("ARTREAS1")
            .Columns("ALL_PCT_TOT").Expression = String.Format("IIF({0}=0,0,100*ISNULL(AMT_ALL,0)/{0})", TOT)
            .Columns("SEL_PCT_TOT").Expression = String.Format("IIF({0}=0,0,100*ISNULL(AMT_SEL,0)/{0})", TOT)
            .Columns("CMP_PCT_TOT").Expression = String.Format("IIF({0}=0,0,100*ISNULL(AMT_CMP,0)/{0})", TOT)
        End With

        With dst.Tables("ARTCUST1")
            .Columns("ALL_PCT_TOT").Expression = String.Format("IIF({0}=0,0,100*ISNULL(AMT_ALL,0)/{0})", TOT)
            .Columns("SEL_PCT_TOT").Expression = String.Format("IIF({0}=0,0,100*ISNULL(AMT_SEL,0)/{0})", TOT)
            .Columns("CMP_PCT_TOT").Expression = String.Format("IIF({0}=0,0,100*ISNULL(AMT_CMP,0)/{0})", TOT)
        End With

    End Sub

    Private Sub optShowActivity_ValueChanged(sender As Object, e As EventArgs) Handles optShowActivity.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_ARTCBDAA()
    End Sub

    Private Sub tab0_SelectedTabChanged(sender As Object, e As UltraWinTabControl.SelectedTabChangedEventArgs) Handles tab0.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab0()
    End Sub

    Sub Setup_tab0()
        optShowActivity.Visible = tab0.SelectedTab.Key = "Activity"
        optShowRollForward.Visible = tab0.SelectedTab.Key = "Roll Forward"
        optExtract.Visible = tab0.SelectedTab.Key = "Extracts"
        chkAllCustomers.Visible = tab0.SelectedTab.Key = "Analysis"

        UltraExplorerBar1.Groups("Period Range").Visible = Not (tab0.SelectedTab.Key = "Chart")
        UltraExplorerBar1.Groups("Chart Settings").Visible = (tab0.SelectedTab.Key = "Chart")

    End Sub

    Private Sub optShowRollForward_ValueChanged(sender As Object, e As EventArgs) Handles optShowRollForward.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_ARTCBDA1()
    End Sub

    Private Sub grdARTCBDA1_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdARTCBDA1.DoubleClickRow

        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            Dim CODE_VALUE As String = e.Row.Cells("CODE_VALUE").Value & ""
            If optShowRollForward.Value = "C" Then
                Absx1.txtFor("CUST_CODE").Text = CODE_VALUE
                Absx1.txtFor("REASON_CODE").Text = ""
            Else
                Absx1.txtFor("CUST_CODE").Text = CODE_VALUE
                Absx1.txtFor("REASON_CODE").Text = ""
            End If
            Click_Command("View")
        End If

    End Sub

    Private Sub optExtract_ValueChanged(sender As Object, e As EventArgs) Handles optExtract.ValueChanged
        Setup_ARTCBDAX()
    End Sub

    Sub Setup_ARTCBDAX()
        If Me.SELECTION_NO = 0 Then Exit Sub

        Dim dvw As DataView = DirectCast(grdARTCBDAX.DataSource, DataTable).DefaultView
        If optExtract.Value = "*" Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "ACTIVITY = '" & optExtract.Value & "'"
        End If
    End Sub

    Sub Build_Charts(TABLE_NAME As String, COLUMN_NAMEs As Dictionary(Of String, String))

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Charts DataSet")

        Dim tbl As New DataTable

        With tbl
            For Each C As String In New String() {"Y", "X", "D", "CODE"}
                .Columns.Add(C)
                .Columns(C).Caption = dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs(C)).Caption
                .Columns(C).DataType = dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs(C)).DataType
            Next
        End With

        tbl.PrimaryKey = New DataColumn() {tbl.Columns("CODE")}


        Dim X_DESC As String = dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("X")).Caption
        Dim Y_DESC As String = dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("Y")).Caption
        Dim D_DESC As String = dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("D")).Caption

        Dim XYD As String = "Correlation between " _
                            & X_DESC & ", " _
                            & Y_DESC & " and " _
                            & D_DESC

        UltraChart1.Axis.X.NumericAxisType = Infragistics.UltraChart.Shared.Styles.NumericAxisType.Logarithmic
        UltraChart1.Axis.X.Extent = 30
        UltraChart1.Axis.X.LogZero = 1
        UltraChart1.Axis.X.LogBase = 10
        UltraChart1.TitleBottom.Text = X_DESC
        ' UltraChart1.TitleTop.Text = X_DESC

        UltraChart1.Axis.Y.NumericAxisType = Infragistics.UltraChart.Shared.Styles.NumericAxisType.Logarithmic
        UltraChart1.Axis.Y.Extent = 30
        UltraChart1.Axis.Y.LogZero = 1
        UltraChart1.Axis.Y.LogBase = 10
        UltraChart1.TitleLeft.Text = Y_DESC

        UltraChart1.Axis.Z.NumericAxisType = Infragistics.UltraChart.Shared.Styles.NumericAxisType.Logarithmic
        UltraChart1.Axis.Z.Extent = 30
        UltraChart1.Axis.Z.LogZero = 1
        UltraChart1.Axis.Z.LogBase = 10


        UltraChart1.DataSource = tbl
        UltraChart1.Data.DataBind()

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        UltraChart1.LabelHash = labelHash

        'UltraChart1.Axis.X.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"
        UltraChart1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"
        UltraChart1.Axis.Z.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"
        UltraChart1.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        UltraChart1.Tooltips.FormatString = "<HIGHLOW>"

        Dim data As String = String.Format("['{0}','{1}','{2}','{3}','{4}']", _
                                           dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("CODE")).Caption, _
                                           dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("X")).Caption, _
                                           dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("Y")).Caption, _
                                           dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("CODE")).Caption, _
                                           dst.Tables(TABLE_NAME).Columns(COLUMN_NAMEs("D")).Caption)

        For Each rowTABLE_NAME As DataRow In dst.Tables(TABLE_NAME).Select("")
            ' Dim CODE = rowTABLE_NAME.Item(COLUMN_NAMEs("CODE"))
            Dim row As DataRow = tbl.NewRow
            For Each C As String In New String() {"CODE", "Y", "X", "D"}
                row.Item(C) = rowTABLE_NAME.Item(COLUMN_NAMEs(C))
            Next
            tbl.Rows.Add(row)

            Dim datum As String = String.Format("['{0}',{1},{2},'{3}',{4}]", _
                                           rowTABLE_NAME.Item(COLUMN_NAMEs("CODE")), _
                                           CStr(Val(rowTABLE_NAME.Item(COLUMN_NAMEs("X")) & "")), _
                                           CStr(Val(rowTABLE_NAME.Item(COLUMN_NAMEs("Y")) & "")), _
                                           rowTABLE_NAME.Item(COLUMN_NAMEs("CODE")), _
                                           CStr(Val(rowTABLE_NAME.Item(COLUMN_NAMEs("D")) & "")))
            data &= "," & vbCrLf & datum
        Next

        grdTable.DataSource = tbl
        ASCMAIN1.grdInitializeLayout(grdTable, Me)

        '  ASCDATA1.DeleteRows(tbl, "ONHAND < 200")
        tabChart.Visible = True


        Dim HTM_FILENAME As String = ASCMAIN1.Folders("Work") & "bubble.html"
        Using sw As New System.IO.StreamWriter(HTM_FILENAME)

            Dim html As String = "" _
                & "<html>" & vbCrLf _
                & "  <head>" & vbCrLf _
                & "    <script type='text/javascript' src='https://www.gstatic.com/charts/loader.js'></script>" & vbCrLf _
                & "    <script type='text/javascript'>" & vbCrLf _
                & "      google.charts.load('current', {'packages':['corechart']});" & vbCrLf _
                & "      google.charts.setOnLoadCallback(drawSeriesChart);" & vbCrLf _
                & "" & vbCrLf _
                & "    function drawSeriesChart() {" & vbCrLf _
                & "" & vbCrLf _
                & "      var data = google.visualization.arrayToDataTable([" & vbCrLf _
                & data & vbCrLf _
                & "      ]);" & vbCrLf _
                & "" & vbCrLf _
                & "      var options = {" & vbCrLf _
                & "        title: '" & XYD & "'," & vbCrLf _
                & "        hAxis: {title: '" & X_DESC & "', scaleType: 'log'}," & vbCrLf _
                & "        vAxis: {title: '" & Y_DESC & "', viewWindowMode: 'Pretty', scaleType: 'log'}," & vbCrLf _
                & "        bubble: {textStyle: {fontSize: 11}}" & vbCrLf _
                & "      };" & vbCrLf _
                & "" & vbCrLf _
                & "      var chart = new google.visualization.BubbleChart(document.getElementById('series_chart_div'));" & vbCrLf _
                & "      chart.draw(data, options);" & vbCrLf _
                & "    }" & vbCrLf _
                & "    </script>" & vbCrLf _
                & "  </head>" & vbCrLf _
                & "  <body>" & vbCrLf _
                & "    <div style='width:100%; height:100%' id='series_chart_div'></div>" & vbCrLf _
                & "  </body>" & vbCrLf _
                & "</html>"
            ' style='width: 900px; height: 500px;'
            sw.WriteLine(html)
        End Using

        WebBrowser1.Navigate(HTM_FILENAME)
        Show_Document(HTM_FILENAME)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub btnChart_Click(sender As Object, e As EventArgs) Handles btnChart.Click
        Dim COLUMN_NAMEs As New Dictionary(Of String, String)
        COLUMN_NAMEs.Add("CODE", optCode.Value)
        Dim T As String = IIf(optCode.Value = "CUST_CODE", "ARTCUST1", "ARTREAS1")

        If optCode.Value <> "CUST_CODE" Then
            MsgBox("Charts work only for Customer at this time")
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If T = "ARTCUST1" Then grd = grdARTCUST1
        If T = "ARTREAS1" Then grd = grdARTREAS1

        Dim COLUMN_NAME As String = optCode.Value
        dst.Tables(T).Columns(COLUMN_NAME).Caption = grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption


        For Each C As String In New String() {"X", "Y", "D"}
            Dim rows() As DataRow = tblChart.Select("XYD = '" & C & "'")
            If rows.Length = 1 Then
                COLUMN_NAME = rows(0).Item("COLUMN_NAME")
                COLUMN_NAMEs.Add(C, COLUMN_NAME)
                dst.Tables(T).Columns(COLUMN_NAME).Caption = grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            Else
                MsgBox("You must select 1 and only 1 dimension for " & C, MsgBoxStyle.OkOnly, "Cannot Render Chart")
                Exit Sub
            End If
        Next

        Build_Charts(T, COLUMN_NAMEs)
    End Sub

    Private Sub grdChart_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdChart.InitializeLayout

    End Sub

    Private Sub grdChart_KeyDown(sender As Object, e As KeyEventArgs) Handles grdChart.KeyDown
        If e.KeyCode = Keys.Delete Then
            If grdChart.ActiveCell IsNot Nothing AndAlso grdChart.ActiveCell.Column.Key = "XYD" Then
                grdChart.ActiveCell.Value = DBNull.Value
            End If
        End If
    End Sub

    Private Sub grdChart_KeyPress(sender As Object, e As KeyPressEventArgs) Handles grdChart.KeyPress

    End Sub

    Private Sub optCode_ValueChanged(sender As Object, e As EventArgs) Handles optCode.ValueChanged

    End Sub
End Class


Public Class MyCustomTooltip
    Implements IRenderLabel

    Public Sub New()

    End Sub 'New

    Public Overloads Function ToString(ByVal Context As Hashtable) As String Implements IRenderLabel.ToString
        'Return Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        'Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
        Return Context("SERIES_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))

    End Function 'ToString 
End Class 'MyCustomTooltip
