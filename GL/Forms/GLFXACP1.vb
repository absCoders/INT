Public Class GLFXACP1
    Dim rowARTCUST1 As DataRow
    Dim CUST_CODE As String
    Dim PROD_CATGY_CODE As String
    Dim OPS_YYYY As String
    Dim OPS_YYYY_LY As String
    Dim PP As String
    Dim P As Integer

    Dim COLLECTION_CODE As String
    Dim ACCT_CODE As String
    Dim OPS_YYYYPP As String
    Dim MM As String
    Dim PERIODS(12) As String
    Dim RECORD_NO_ctr As Integer = 100

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "GLTACCT1", "*", 0, False)
            '   Create_TDA(.Tables.Add, "ICTBRAN1", "*", 1, False)

            ASCMAIN1.sql = "Select ICTCOLL1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & " from ICTCOLL1,ICTBRAN1 where ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE and ICTBRAN1.PROD_CATGY_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOLL1", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select * from GLTXACPB where OPS_YYYY = :PARM1"
            Create_TDA(.Tables.Add, "GLTXACPB", "**", , , "V", 3)

            ASCMAIN1.sql = "Select * from GLTXACPF where OPS_YYYYPP >= :PARM1 and OPS_YYYYPP <= :PARM2"
            Create_TDA(.Tables.Add, "GLTXACPF", "**", , , "VV", 1)

            ASCMAIN1.sql = "Select * from GLTXACPA where OPS_YYYYPP >= :PARM1 and OPS_YYYYPP <= :PARM2"
            Create_TDA(.Tables.Add, "GLTXACPA", "**", , , "VV", 1)

            Create_Relation("GLTXACPF", "GLTXACPA", "RECORD_NO", "RECORD_NO_LINK")

            With .Tables("GLTXACPF")
                .Columns.Add("AMT_A", GetType(System.Decimal), "SUM(CHILD.AMT)")
                .Columns.Add("AMT_V", GetType(System.Decimal), "ISNULL(AMT_A,0) - ISNULL(AMT,0)")
            End With
            With .Tables("GLTXACPA")
                .Columns.Add("AMT_F")
                .Columns.Add("AMT_V")
            End With

            'ASCMAIN1.sql = "Select COLLECTION_CODE, ACCT_CODE from GLTXACPF"
            'Create_TDA(.Tables.Add, "GLTXACPX", "**", , False, "", 1)
            .Tables.Add("GLTXACPX") ' PROBABLY SHOULD COME FROM GLTXACPB WHERE BUDGET TOTAL FOR YEAR LIVES

            With .Tables("GLTXACPX")
                .Columns.Add("COLLECTION_CODE")
                .Columns.Add("ACCT_CODE")
                .PrimaryKey = New DataColumn() {.Columns("COLLECTION_CODE"), .Columns("ACCT_CODE")}

                .Columns.Add("COLLECTION_NAME")
                .Columns.Add("ACCT_DESC")
                .Columns.Add("ACCT_CLASS_CODE")

                .Columns.Add("Y1")
                .Columns.Add("Y2")
                .Columns.Add("Y3")

                .Columns("Y1").DefaultValue = "Budget"
                .Columns("Y2").DefaultValue = "^Plan"
                .Columns("Y3").DefaultValue = "PO Var"

                .Columns.Add("AMT", GetType(System.Decimal))
                .Columns.Add("UNP", GetType(System.Decimal))
                .Columns.Add("VAR", GetType(System.Decimal))

                .Columns.Add("L1")
                .Columns.Add("L2")
                .Columns.Add("L3")

                .Columns("L1").DefaultValue = "LE"
                .Columns("L2").DefaultValue = "Plan"
                .Columns("L3").DefaultValue = "Var"

                For I As Integer = 1 To 12
                    .Columns.Add("AMT_F" & Format(I, "00"), GetType(System.Decimal))
                Next
                For I As Integer = 1 To 12
                    .Columns.Add("AMT_A" & Format(I, "00"), GetType(System.Decimal))
                Next
                For I As Integer = 1 To 12
                    '  .Columns.Add("AMT_V" & Format(I, "00"), GetType(System.Decimal), "ISNULL(" & "AMT_F" & Format(I, "00") & ",0)-ISNULL(" & "AMT_A" & Format(I, "00") & ",0)")
                    .Columns.Add("AMT_V" & Format(I, "00"), GetType(System.Decimal))
                Next

                .Columns.Add("AMT_FQ1", GetType(System.Decimal), "ISNULL(AMT_F01,0)+ISNULL(AMT_F02,0)+ISNULL(AMT_F03,0)")
                .Columns.Add("AMT_FQ2", GetType(System.Decimal), "ISNULL(AMT_F04,0)+ISNULL(AMT_F05,0)+ISNULL(AMT_F06,0)")
                .Columns.Add("AMT_FQ3", GetType(System.Decimal), "ISNULL(AMT_F07,0)+ISNULL(AMT_F08,0)+ISNULL(AMT_F09,0)")
                .Columns.Add("AMT_FQ4", GetType(System.Decimal), "ISNULL(AMT_F10,0)+ISNULL(AMT_F11,0)+ISNULL(AMT_F12,0)")

                .Columns.Add("AMT_FYTD", GetType(System.Decimal), "ISNULL(AMT_F01,0)")
                .Columns.Add("AMT_FBTG", GetType(System.Decimal), "ISNULL(AMT_F01,0)")
                .Columns.Add("AMT_FTOT", GetType(System.Decimal), "ISNULL(AMT_FYTD,0)+ISNULL(AMT_FBTG,0)")

                .Columns.Add("AMT_AQ1", GetType(System.Decimal), "ISNULL(AMT_A01,0)+ISNULL(AMT_A02,0)+ISNULL(AMT_A03,0)")
                .Columns.Add("AMT_AQ2", GetType(System.Decimal), "ISNULL(AMT_A04,0)+ISNULL(AMT_A05,0)+ISNULL(AMT_A06,0)")
                .Columns.Add("AMT_AQ3", GetType(System.Decimal), "ISNULL(AMT_A07,0)+ISNULL(AMT_A08,0)+ISNULL(AMT_A09,0)")
                .Columns.Add("AMT_AQ4", GetType(System.Decimal), "ISNULL(AMT_A10,0)+ISNULL(AMT_A11,0)+ISNULL(AMT_A12,0)")

                .Columns.Add("AMT_AYTD", GetType(System.Decimal), "ISNULL(AMT_A01,0)")
                .Columns.Add("AMT_ABTG", GetType(System.Decimal), "ISNULL(AMT_A01,0)")
                .Columns.Add("AMT_ATOT", GetType(System.Decimal), "ISNULL(AMT_AYTD,0)+ISNULL(AMT_ABTG,0)")

                .Columns.Add("AMT_VQ1", GetType(System.Decimal), "ISNULL(AMT_V01,0)+ISNULL(AMT_V02,0)+ISNULL(AMT_V03,0)")
                .Columns.Add("AMT_VQ2", GetType(System.Decimal), "ISNULL(AMT_V04,0)+ISNULL(AMT_V05,0)+ISNULL(AMT_V06,0)")
                .Columns.Add("AMT_VQ3", GetType(System.Decimal), "ISNULL(AMT_V07,0)+ISNULL(AMT_V08,0)+ISNULL(AMT_V09,0)")
                .Columns.Add("AMT_VQ4", GetType(System.Decimal), "ISNULL(AMT_V10,0)+ISNULL(AMT_V11,0)+ISNULL(AMT_V12,0)")

                .Columns.Add("AMT_VYTD", GetType(System.Decimal), "ISNULL(AMT_V01,0)")
                .Columns.Add("AMT_VBTG", GetType(System.Decimal), "ISNULL(AMT_V01,0)")
                .Columns.Add("AMT_VTOT", GetType(System.Decimal), "ISNULL(AMT_VYTD,0)+ISNULL(AMT_VBTG,0)")

                .Columns("UNP").Expression = "ISNULL(AMT,0)-ISNULL(AMT_FTOT,0)"
                .Columns("VAR").Expression = "ISNULL(AMT,0)-0"

            End With

            .Tables.Add("GLTXACPY") ' PROBABLY SHOULD COME FROM GLTXACPP WHERE PLAN RECORD LIVES

            With .Tables("GLTXACPY")
                .Columns.Add("COLLECTION_CODE")
                .Columns.Add("ACCT_CODE")
                .Columns.Add("RECORD_NO")
                .PrimaryKey = New DataColumn() {.Columns("COLLECTION_CODE"), .Columns("ACCT_CODE"), .Columns("RECORD_NO")}

                .Columns.Add("NOTE")

                For I As Integer = 1 To 12
                    .Columns.Add("AMT_P" & Format(I, "00"), GetType(System.Decimal))
                Next
               
                .Columns.Add("AMT_PQ1", GetType(System.Decimal), "ISNULL(AMT_P01,0)+ISNULL(AMT_P02,0)+ISNULL(AMT_P03,0)")
                .Columns.Add("AMT_PQ2", GetType(System.Decimal), "ISNULL(AMT_P04,0)+ISNULL(AMT_P05,0)+ISNULL(AMT_P06,0)")
                .Columns.Add("AMT_PQ3", GetType(System.Decimal), "ISNULL(AMT_P07,0)+ISNULL(AMT_P08,0)+ISNULL(AMT_P09,0)")
                .Columns.Add("AMT_PQ4", GetType(System.Decimal), "ISNULL(AMT_P10,0)+ISNULL(AMT_P11,0)+ISNULL(AMT_P12,0)")

                .Columns.Add("AMT_PYTD", GetType(System.Decimal), "ISNULL(AMT_P01,0)")
                .Columns.Add("AMT_PBTG", GetType(System.Decimal), "ISNULL(AMT_P01,0)")
                .Columns.Add("AMT_PTOT", GetType(System.Decimal), "ISNULL(AMT_PYTD,0)+ISNULL(AMT_PBTG,0)")
            End With

            Create_Relation("GLTXACPX", "GLTXACPY", "COLLECTION_CODE,ACCT_CODE")

            For I As Integer = 1 To 12
                .Tables("GLTXACPX").Columns("AMT_F" & Format(I, "00")).Expression = "SUM(CHILD.AMT_P" & Format(I, "00") & ")"
            Next




            With .Tables.Add("GLTXACPT")
                .Columns.Add("ACCT_CLASS_CODE")
                .Columns.Add("BUDGET_APPR", GetType(System.Decimal))
                .Columns.Add("BUDGET_CURR", GetType(System.Decimal))
                .Columns.Add("BUDGET_DIFF", GetType(System.Decimal), "ISNULL(BUDGET_APPR,0)-ISNULL(BUDGET_CURR,0)")
                .PrimaryKey = New DataColumn() {.Columns("ACCT_CLASS_CODE")}
            End With

            Create_Relation("GLTXACPT", "GLTXACPX", "ACCT_CLASS_CODE")

            .Tables("GLTXACPT").Columns("BUDGET_CURR").Expression = "SUM(CHILD.AMT)"

        End With

        Fill_Record("GLTACCT1")

        grdGLTACCT1.DataSource = dst.Tables("GLTACCT1")
        grdICTCOLL1.DataSource = dst.Tables("ICTCOLL1")
        grdGLTXACPX.DataSource = dst.Tables("GLTXACPX")

        grdGLTXACPF.DataSource = dst.Tables("GLTXACPF")
        grdGLTXACPT.DataSource = dst.Tables("GLTXACPT")
        Create_Summary(grdGLTXACPT, New String() {"BUDGET_APPR", "BUDGET_CURR", "BUDGET_DIFF"})


        For B As Integer = 0 To 1
            With grdGLTXACPF.DisplayLayout.Bands(B)
                For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                    If gcol.Key = "AMT" Or gcol.Key = "NOTE" Then
                        gcol.CellActivation = UltraWinGrid.Activation.AllowEdit

                    Else
                        gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                        gcol.CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    End If
                Next
            End With
        Next

        grdGLTXACPF.AllowDrop = True
        Create_Summary(grdGLTXACPF, New String() {"AMT", "AMT_A", "AMT_V"})


        lblLEGEND.Text = ASCMAIN1.Get_Legend(ASCMAIN1.CYP)


        With grdGLTXACPX.DisplayLayout.Bands(0)

            .Override.AllowUpdate = DefaultableBoolean.True


            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                If gcol.Key = "AMT" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                    gcol.CellAppearance.BorderColor = Drawing.Color.Red
                    gcol.CellAppearance.BackColor = Drawing.Color.Yellow
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            Next

            .LevelCount = 3

            Dim g As UltraWinGrid.UltraGridGroup

            g = .Groups.Add("CODES")
            g.Header.Caption = "GL Account"

            g.Header.Appearance.BackColor = Drawing.Color.White
            g.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            '.Columns("COLLECTION_CODE").Group = g
            '.Columns("COLLECTION_NAME").Group = g
            .Columns("ACCT_CODE").Group = g
            .Columns("ACCT_DESC").Group = g
            .Columns("ACCT_CLASS_CODE").Group = g
            .Columns("ACCT_CODE").Level = 0
            .Columns("ACCT_DESC").Level = 1
            .Columns("ACCT_CLASS_CODE").Level = 2
            g.Width = 120
            g.Header.Fixed = True


            g = .Groups.Add("YEAR_DATA")
            g.Header.Caption = "Value"
            g.Header.Appearance.TextHAlign = HAlign.Center

            g.Header.Appearance.BackColor = Drawing.Color.White
            g.Header.Appearance.BackColor2 = Drawing.Color.Pink
            g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            .Columns("Y1").Group = g
            .Columns("Y2").Group = g
            .Columns("Y3").Group = g
            .Columns("Y1").Level = 0
            .Columns("Y2").Level = 1
            .Columns("Y3").Level = 2
            .Columns("Y1").CellAppearance.TextHAlign = HAlign.Center
            .Columns("Y2").CellAppearance.TextHAlign = HAlign.Center
            .Columns("Y3").CellAppearance.TextHAlign = HAlign.Center
            g.Width = 60
            g.Header.Fixed = True
            g.Header.Fixed = True

            Create_Summary(grdGLTXACPX, "Y1", "Max")
            Create_Summary(grdGLTXACPX, "Y2", "Max")
            Create_Summary(grdGLTXACPX, "Y3", "Max")


            g = .Groups.Add("YEAR")
            g.Header.Caption = "Year"
            g.Header.Appearance.TextHAlign = HAlign.Right

            g.Header.Appearance.BackColor = Drawing.Color.White
            g.Header.Appearance.BackColor2 = Drawing.Color.Pink
            g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            .Columns("AMT").Group = g
            .Columns("UNP").Group = g
            .Columns("VAR").Group = g
            .Columns("AMT").Level = 0
            .Columns("UNP").Level = 1
            .Columns("VAR").Level = 2
            .Columns("AMT").Format = "#,##0"
            .Columns("UNP").Format = "#,##0"
            .Columns("VAR").Format = "#,##0"
            g.Width = 60
            g.Header.Fixed = True

            Create_Summary(grdGLTXACPX, "AMT", , , "###,##0")
            Create_Summary(grdGLTXACPX, "UNP", , , "###,##0")
            Create_Summary(grdGLTXACPX, "VAR", , , "###,##0")




            g = .Groups.Add("DATA")
            g.Header.Caption = "Data"
            g.Header.Appearance.TextHAlign = HAlign.Center

            g.Header.Appearance.BackColor = Drawing.Color.White
            g.Header.Appearance.BackColor2 = Drawing.Color.LightGray
            g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            .Columns("L1").Group = g
            .Columns("L2").Group = g
            .Columns("L3").Group = g
            .Columns("L1").Level = 0
            .Columns("L2").Level = 1
            .Columns("L3").Level = 2
            .Columns("L1").CellAppearance.TextHAlign = HAlign.Center
            .Columns("L2").CellAppearance.TextHAlign = HAlign.Center
            .Columns("L3").CellAppearance.TextHAlign = HAlign.Center
            g.Width = 60
            g.Header.Fixed = True

            Create_Summary(grdGLTXACPX, "L1", "Max")
            Create_Summary(grdGLTXACPX, "L2", "Max")
            Create_Summary(grdGLTXACPX, "L3", "Max")

            g = .Groups.Add("ACTYTD")
            g.Header.Caption = "Act Ytd"
            g.Header.Appearance.TextHAlign = HAlign.Right

            g.Header.Appearance.BackColor = Drawing.Color.White
            g.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            .Columns("AMT_FYTD").Group = g
            .Columns("AMT_AYTD").Group = g
            .Columns("AMT_VYTD").Group = g
            .Columns("AMT_FYTD").Level = 0
            .Columns("AMT_AYTD").Level = 1
            .Columns("AMT_VYTD").Level = 2
            .Columns("AMT_FYTD").Format = "#,##0"
            .Columns("AMT_AYTD").Format = "#,##0"
            .Columns("AMT_VYTD").Format = "#,##0"
            g.Width = 60
            g.Header.Fixed = True

            Create_Summary(grdGLTXACPX, "AMT_FYTD", , , "###,##0")
            Create_Summary(grdGLTXACPX, "AMT_AYTD", , , "###,##0")
            Create_Summary(grdGLTXACPX, "AMT_VYTD", , , "###,##0")


            g = .Groups.Add("AMTBTG")
            g.Header.Caption = "BTG"
            g.Header.Appearance.TextHAlign = HAlign.Right

            g.Header.Appearance.BackColor = Drawing.Color.White
            g.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            .Columns("AMT_FBTG").Group = g
            .Columns("AMT_ABTG").Group = g
            .Columns("AMT_VBTG").Group = g
            .Columns("AMT_FBTG").Level = 0
            .Columns("AMT_ABTG").Level = 1
            .Columns("AMT_VBTG").Level = 2
            .Columns("AMT_FBTG").Format = "#,##0"
            .Columns("AMT_ABTG").Format = "#,##0"
            .Columns("AMT_VBTG").Format = "#,##0"
            g.Width = 60
            g.Header.Fixed = True

            Create_Summary(grdGLTXACPX, "AMT_FBTG", , , "###,##0")
            Create_Summary(grdGLTXACPX, "AMT_ABTG", , , "###,##0")
            Create_Summary(grdGLTXACPX, "AMT_VBTG", , , "###,##0")

            g = .Groups.Add("AMTTOT")
            g.Header.Caption = "Total"
            g.Header.Appearance.TextHAlign = HAlign.Right

            g.Header.Appearance.BackColor = Drawing.Color.White
            g.Header.Appearance.BackColor2 = Drawing.Color.LightBlue
            g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

            .Columns("AMT_FTOT").Group = g
            .Columns("AMT_ATOT").Group = g
            .Columns("AMT_VTOT").Group = g
            .Columns("AMT_FTOT").Level = 0
            .Columns("AMT_ATOT").Level = 1
            .Columns("AMT_VTOT").Level = 2
            .Columns("AMT_FTOT").Format = "#,##0"
            .Columns("AMT_ATOT").Format = "#,##0"
            .Columns("AMT_VTOT").Format = "#,##0"
            g.Width = 60
            g.Header.Fixed = True

            Create_Summary(grdGLTXACPX, "AMT_FTOT", , , "###,##0")
            Create_Summary(grdGLTXACPX, "AMT_ATOT", , , "###,##0")
            Create_Summary(grdGLTXACPX, "AMT_VTOT", , , "###,##0")


            For M As Integer = 1 To 12

                g = .Groups.Add("M" & Format(M, "00"))
                g.Header.Caption = "MMM'YY"
                g.Header.Appearance.TextHAlign = HAlign.Right


                g.Header.Appearance.BackColor = Drawing.Color.White
                g.Header.Appearance.BackColor2 = Drawing.Color.LightGreen
                g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                .Columns("AMT_F" & Format(M, "00")).Group = g
                .Columns("AMT_A" & Format(M, "00")).Group = g
                .Columns("AMT_V" & Format(M, "00")).Group = g
                .Columns("AMT_F" & Format(M, "00")).Level = 0
                .Columns("AMT_A" & Format(M, "00")).Level = 1
                .Columns("AMT_V" & Format(M, "00")).Level = 2
                .Columns("AMT_F" & Format(M, "00")).Format = "#,##0"
                .Columns("AMT_A" & Format(M, "00")).Format = "#,##0"
                .Columns("AMT_V" & Format(M, "00")).Format = "#,##0"
                g.Width = 60

                Create_Summary(grdGLTXACPX, "AMT_F" & Format(M, "00"), , , "###,##0")
                Create_Summary(grdGLTXACPX, "AMT_A" & Format(M, "00"), , , "###,##0")
                Create_Summary(grdGLTXACPX, "AMT_V" & Format(M, "00"), , , "###,##0")

                If M Mod 3 = 0 Then
                    Dim Q As Integer = M / 3
                    g = .Groups.Add("Q" & Format(Q, "00"))
                    g.Header.Caption = "Q" & Format(Q, "00")
                    g.Header.Appearance.TextHAlign = HAlign.Right

                    g.Header.Appearance.BackColor = Drawing.Color.White
                    g.Header.Appearance.BackColor2 = Drawing.Color.Gold
                    g.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                    .Columns("AMT_FQ" & Format(Q, "0")).Group = g
                    .Columns("AMT_AQ" & Format(Q, "0")).Group = g
                    .Columns("AMT_VQ" & Format(Q, "0")).Group = g
                    .Columns("AMT_FQ" & Format(Q, "0")).Level = 0
                    .Columns("AMT_AQ" & Format(Q, "0")).Level = 1
                    .Columns("AMT_VQ" & Format(Q, "0")).Level = 2
                    .Columns("AMT_FQ" & Format(Q, "0")).Format = "#,##0"
                    .Columns("AMT_AQ" & Format(Q, "0")).Format = "#,##0"
                    .Columns("AMT_VQ" & Format(Q, "0")).Format = "#,##0"

                    .Columns("AMT_FQ" & Format(Q, "0")).CellAppearance.BackColor = Drawing.Color.LightGray
                    .Columns("AMT_AQ" & Format(Q, "0")).CellAppearance.BackColor = Drawing.Color.LightGray
                    .Columns("AMT_VQ" & Format(Q, "0")).CellAppearance.BackColor = Drawing.Color.LightGray


                    g.Width = 60

                    Create_Summary(grdGLTXACPX, "AMT_FQ" & Format(Q, "0"), , , "###,##0")
                    Create_Summary(grdGLTXACPX, "AMT_AQ" & Format(Q, "0"), , , "###,##0")
                    Create_Summary(grdGLTXACPX, "AMT_VQ" & Format(Q, "0"), , , "###,##0")
                End If
            Next
        End With
        Create_Summary(grdGLTXACPX, "ACCT_CODE", "Count")



        With grdGLTXACPX.DisplayLayout.Bands(1)
            .Override.AllowUpdate = DefaultableBoolean.True
            .ColHeadersVisible = False

            .Columns("COLLECTION_CODE").Hidden = True
            .Columns("ACCT_CODE").Hidden = True
            .Columns("NOTE").ColSpan = 6
            .Columns("NOTE").Header.VisiblePosition = 0
            .Columns("NOTE").Width = grdGLTXACPX.DisplayLayout.Bands(0).Groups("CODES").Width + grdGLTXACPX.DisplayLayout.Bands(0).Groups("YEAR_DATA").Width + grdGLTXACPX.DisplayLayout.Bands(0).Groups("YEAR").Width - 20
            .Columns("NOTE").Header.Fixed = True

            .Columns("RECORD_NO").Header.VisiblePosition = 1
            .Columns("RECORD_NO").Width = grdGLTXACPX.DisplayLayout.Bands(0).Groups("DATA").Width
            .Columns("RECORD_NO").Header.Fixed = True

            .Columns("AMT_PYTD").Header.Fixed = True
            .Columns("AMT_PBTG").Header.Fixed = True
            .Columns("AMT_PTOT").Header.Fixed = True

            .Columns("AMT_PYTD").Header.VisiblePosition = .Columns("RECORD_NO").Header.VisiblePosition + 1
            .Columns("AMT_PBTG").Header.VisiblePosition = .Columns("RECORD_NO").Header.VisiblePosition + 2
            .Columns("AMT_PTOT").Header.VisiblePosition = .Columns("RECORD_NO").Header.VisiblePosition + 3

            .Columns("AMT_PQ1").Header.VisiblePosition = .Columns("AMT_P03").Header.VisiblePosition + 1
            .Columns("AMT_PQ2").Header.VisiblePosition = .Columns("AMT_P06").Header.VisiblePosition + 1
            .Columns("AMT_PQ3").Header.VisiblePosition = .Columns("AMT_P09").Header.VisiblePosition + 1
            .Columns("AMT_PQ4").Header.VisiblePosition = .Columns("AMT_P12").Header.VisiblePosition + 1

            .Columns("AMT_PQ1").CellAppearance.BackColor = Drawing.Color.LightGray
            .Columns("AMT_PQ2").CellAppearance.BackColor = Drawing.Color.LightGray
            .Columns("AMT_PQ3").CellAppearance.BackColor = Drawing.Color.LightGray
            .Columns("AMT_PQ4").CellAppearance.BackColor = Drawing.Color.LightGray

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                If gcol.Key = "NOTE" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                ElseIf gcol.Key.StartsWith("AMT_") Then
                    Dim COL As String = gcol.Key
                    Mid(COL, 5, 1) = "A"
                    gcol.Format = grdGLTXACPX.DisplayLayout.Bands(0).Columns(COL).Format
                    gcol.Width = grdGLTXACPX.DisplayLayout.Bands(0).Columns(COL).Width
                End If
            Next

        End With

        grdGLTXACPX.DisplayLayout.Override.AllowColSizing = UltraWinGrid.AllowColSizing.None


        cmbCOLLECTION_CODE.DataSource = dst.Tables("ICTCOLL1")
        Sort_cmbColumns(cmbCOLLECTION_CODE, "BRAND_CODE,COLLECTION_NAME")

        Dim YEARs As New List(Of String)
        For Y As Integer = Val(Now.Year) - 1 To Val(Now.Year + 1)
            YEARs.Add(Format(Y + 1, "0000"))
        Next
        Absx1.cbeFor("OPS_YYYY").DataSource = YEARs ' New String() {"2008", "2009", "2010"}
        Absx1.cbeFor("OPS_YYYY").Value = YEARs(1)
        Show_Filter(grdGLTACCT1, True)
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("CUST_CODE", , True)
                Validate_Code("PROD_CATGY_CODE")

                'If EMsg = "" Then
                '    If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "OPS_YYYY" & ":" & Absx1.cbeFor("OPS_YYYY").Value) Then
                '        Exit Sub
                '    End If
                '    If optClass.Value = "O" Then
                '        If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "OPS_YYYY" & ":" & Format(Val(Absx1.cbeFor("OPS_YYYY").Value & "") - 1, "0000")) Then
                '            Exit Sub
                '        End If
                '    End If
                '    If Not ASCMAIN1.Logical_Lock("RSTBUDR1", "CUST_CODE" & ":" & Absx1.txtFor("CUST_CODE").Text) Then
                '        Exit Sub
                '    End If
                'End If

            Case "Update"
                'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDR1"), New String() {"CUST_CODE", "CUST_STORE_NO"}).Select("")
                '    Dim CUST_CODE As String = row.Item(0)
                '    Dim CUST_STORE_NO As String = row.Item(1)
                '    If dst.Tables("ARTCUSTX").Rows.Find(New String() {CUST_CODE, CUST_STORE_NO}) Is Nothing Then
                '        EMsg &= vbCr & "Invalid Store (" & CUST_CODE & "," & CUST_STORE_NO & ")"
                '    End If
                'Next
                'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDR1"), New String() {"COLLECTION_CODE"}).Select("")
                '    Dim COLLECTION_CODE As String = row.Item(0)
                '    If dst.Tables("ICTCOLL1").Rows.Find(New String() {COLLECTION_CODE}) Is Nothing Then
                '        EMsg &= vbCr & "Invalid Collection (" & COLLECTION_CODE & ")"
                '    End If
                'Next
                'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("RSTBUDR1"), New String() {"OPS_YYYY"}).Select("")
                '    Dim OPS_YYYY As String = row.Item(0)
                '    If OPS_YYYY <> Absx1.cbeFor("OPS_YYYY").Value & "" Then
                '        EMsg &= vbCr & "Invalid Budget Year (" & OPS_YYYY & ")"
                '    End If
                'Next
    
              
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "L"
                Load_Record()
                Mode_Settings(True)

            Case "Update"
                Update_Record()
                Mode_Settings(False)

            Case "Cancel", "Done"
                Mode_Settings(False)

        

        End Select
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Update").Settings.Enabled = iScreenMode
                    .Items("Cancel").Settings.Enabled = iScreenMode
                    .Items("Hyperion XLS").Visible = ScreenMode
                    .Items("Close Period").Visible = ScreenMode
                End With

                .Groups("Display Options").Visible = ScreenMode And (EntryMode = "L")
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        splGLTXACPX.Visible = ScreenMode

        chkShowGL.Visible = ScreenMode
        optClass.Visible = ScreenMode

        lblCUST_CODE.Visible = Not ScreenMode
        lblCUST_NAME.Visible = Not ScreenMode
        Absx1.txtFor("CUST_CODE").Visible = Not ScreenMode
        Absx1.txtFor("CUST_NAME").Visible = Not ScreenMode

        lblCOLLECTION_CODE.Visible = ScreenMode
        cmbCOLLECTION_CODE.Visible = ScreenMode

        If ScreenMode Then
            Set_Read_Only_for_ctl(chkShowGL, False)
            Set_Read_Only_for_ctl(cmbCOLLECTION_CODE, False)
            Set_Read_Only_for_ctl(optClass, False)
        Else
            Clear_Record()
        End If
    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"ICTCOLL1", "GLTXACPF", "GLTXACPA", "GLTXACPX", "GLTXACPY", "GLTXACPB"} ' GLTXACPX
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        CUST_CODE = ""
        PROD_CATGY_CODE = ""
        OPS_YYYY = ""

        chkShowGL.Checked = False
        splGLTXACPX.Panel1Collapsed = True

        Absx1.txtFor("CUST_CODE").Text = ""
        Absx1.txtFor("PROD_CATGY_CODE").Text = ""
        Absx1.cbeFor("OPS_YYYY").Value = Mid(ASCMAIN1.CYP, 1, 4)
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Data ...")

        Save_Header_Fields(UltraGroupBox1)
        CUST_CODE = HFs("CUST_CODE")
        PROD_CATGY_CODE = HFs("PROD_CATGY_CODE")
        OPS_YYYY = HFs("OPS_YYYY")
        OPS_YYYY_LY = Format(Val(OPS_YYYY) - 1, "0000")

        If OPS_YYYY = Mid(ASCMAIN1.CYP, 1, 4) Then
            PP = Mid(ASCMAIN1.CYP, 5, 2)
        ElseIf OPS_YYYY > Mid(ASCMAIN1.CYP, 1, 4) Then
            PP = "00"
        Else
            PP = "12"
        End If
        P = Val(PP)

        If PP <> "00" And PP <> "12" Then
            chkHideHistory.Visible = True
        Else
            chkHideHistory.Visible = False
        End If

        With dst.Tables("GLTXACPY")
            .Columns("AMT_PYTD").Expression = Get_Cols("AMT_P00", 1, P)
            .Columns("AMT_PBTG").Expression = Get_Cols("AMT_P00", P + 1, 12 - P)
        End With

        With grdGLTXACPX.DisplayLayout.Bands(1)
            For m As Integer = 1 To 12
                PERIODS(m) = OPS_YYYY & Format(m, "00")
                If OPS_YYYY & Format(m, "00") < ASCMAIN1.CYP Then
                    .Columns("AMT_P" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    .Columns("AMT_P" & Format(m, "00")).CellActivation = UltraWinGrid.Activation.NoEdit
                Else
                    .Columns("AMT_P" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.Empty
                    .Columns("AMT_P" & Format(m, "00")).CellActivation = UltraWinGrid.Activation.AllowEdit
                End If
            Next
        End With


        With dst.Tables("GLTXACPX")
            .Columns("AMT_FYTD").Expression = Get_Cols("AMT_F00", 1, P)
            .Columns("AMT_FBTG").Expression = Get_Cols("AMT_F00", P + 1, 12 - P)
            .Columns("AMT_AYTD").Expression = Get_Cols("AMT_A00", 1, P)
            .Columns("AMT_ABTG").Expression = Get_Cols("AMT_A00", P + 1, 12 - P)
            .Columns("AMT_VYTD").Expression = Get_Cols("AMT_V00", 1, P)
            .Columns("AMT_VBTG").Expression = Get_Cols("AMT_V00", P + 1, 12 - P)
        End With

        With grdGLTXACPX.DisplayLayout.Bands(0)
            For m As Integer = 1 To 12
                PERIODS(m) = OPS_YYYY & Format(m, "00")
                If OPS_YYYY & Format(m, "00") < ASCMAIN1.CYP Then
                    .Columns("AMT_F" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    .Columns("AMT_A" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                    .Columns("AMT_V" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.WhiteSmoke
                Else
                    .Columns("AMT_F" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.Empty
                    .Columns("AMT_A" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.Empty
                    .Columns("AMT_V" & Format(m, "00")).CellAppearance.BackColor = Drawing.Color.Empty
                End If
            Next
        End With
 



       

        EnforceConstraints(False)
        

        dst.Tables("GLTXACPT").Rows.Add(New Object() {"MKTG", 10000})
        dst.Tables("GLTXACPT").Rows.Add(New Object() {"MEDIA", 20000})

        Fill_Records("ICTCOLL1", PROD_CATGY_CODE)
        Sort_grdColumns(grdICTCOLL1, "BRAND_CODE,COLLECTION_NAME")



        COLLECTION_CODE = "2127"
        cmbCOLLECTION_CODE.Value = COLLECTION_CODE


        Dim row As DataRow


        ' add oct


        row = dst.Tables("GLTXACPF").NewRow
        row.Item("RECORD_NO") = "000011"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201804"
        row.Item("AMT") = 1010
        row.Item("NOTE") = "Forecasting Rain in oct"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "smoser"
        dst.Tables("GLTXACPF").Rows.Add(row)


        row = dst.Tables("GLTXACPF").NewRow
        row.Item("RECORD_NO") = "000012"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201804"
        row.Item("AMT") = 2010
        row.Item("NOTE") = "Forecasting Snow in oct"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "smoser"
        dst.Tables("GLTXACPF").Rows.Add(row)


        row = dst.Tables("GLTXACPA").NewRow
        row.Item("RECORD_NO") = "000011"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201804"
        row.Item("AMT") = 950
        row.Item("NOTE") = "IT RAINED in oct"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "walter"
        row.Item("RECORD_NO_LINK") = "000011"
        dst.Tables("GLTXACPA").Rows.Add(row)


        row = dst.Tables("GLTXACPA").NewRow
        row.Item("RECORD_NO") = "000012"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201804"
        row.Item("AMT") = 110
        row.Item("NOTE") = "IT RAINED some more in oct"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "walter"
        row.Item("RECORD_NO_LINK") = "000011"
        dst.Tables("GLTXACPA").Rows.Add(row)



        row = Add_Unforcasted_Row()
        row.Item("RECORD_NO") = "999998"
        row.Item("OPS_YYYYPP") = "201804"

        row = dst.Tables("GLTXACPA").NewRow
        row.Item("RECORD_NO") = "000013"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201804"
        row.Item("AMT") = 610
        row.Item("NOTE") = "IT rained cats and dogs in oct"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "david"
        row.Item("RECORD_NO_LINK") = "999998"
        dst.Tables("GLTXACPA").Rows.Add(row)





        ' add nov

        row = dst.Tables("GLTXACPF").NewRow
        row.Item("RECORD_NO") = "000001"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201805"
        row.Item("AMT") = 1000
        row.Item("NOTE") = "Forecasting Rain"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "smoser"
        dst.Tables("GLTXACPF").Rows.Add(row)


        row = dst.Tables("GLTXACPF").NewRow
        row.Item("RECORD_NO") = "000002"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201805"
        row.Item("AMT") = 2000
        row.Item("NOTE") = "Forecasting Snow"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "smoser"
        dst.Tables("GLTXACPF").Rows.Add(row)


        row = dst.Tables("GLTXACPA").NewRow
        row.Item("RECORD_NO") = "000001"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201805"
        row.Item("AMT") = 950
        row.Item("NOTE") = "IT RAINED"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "walter"
        row.Item("RECORD_NO_LINK") = "000001"
        dst.Tables("GLTXACPA").Rows.Add(row)


        row = dst.Tables("GLTXACPA").NewRow
        row.Item("RECORD_NO") = "000002"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201805"
        row.Item("AMT") = 100
        row.Item("NOTE") = "IT RAINED some more"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "walter"
        row.Item("RECORD_NO_LINK") = "000001"
        dst.Tables("GLTXACPA").Rows.Add(row)

        Add_Unforcasted_Row()

        row = dst.Tables("GLTXACPA").NewRow
        row.Item("RECORD_NO") = "000003"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201805"
        row.Item("AMT") = 600
        row.Item("NOTE") = "IT rained cats and dogs"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "david"
        row.Item("RECORD_NO_LINK") = "999999"
        dst.Tables("GLTXACPA").Rows.Add(row)



        Load_Details_into_Summary()


        EnforceConstraints(True)
 


        Set_Month_Headings(OPS_YYYY)
 

        ASCMAIN1.Progress("")
    End Sub

    Function Add_Unforcasted_Row() As DataRow

        Dim row As DataRow = dst.Tables("GLTXACPF").NewRow
        row.Item("RECORD_NO") = "999999"
        row.Item("ACCT_CODE") = "506032"
        row.Item("COLLECTION_CODE") = COLLECTION_CODE
        row.Item("OPS_YYYYPP") = "201805"
        row.Item("AMT") = 0
        row.Item("NOTE") = "Unforecasted"
        row.Item("INIT_DATE") = Now
        row.Item("INIT_OPER") = "smoser"
        dst.Tables("GLTXACPF").Rows.Add(row)

        Return row

    End Function
    Function Get_Cols(colexp As String, p_first As Integer, p_last As Integer) As String

        Dim COLX As String = ""

        If p_last < p_first Then
            Return "0"
        Else
            For i As Integer = p_first To p_last
                COLX &= "+ISNULL(" & Replace(colexp, "00", Format(i, "00")) & ",0)"
            Next
            Return Mid(COLX, 2)
        End If
    End Function

    Sub Update_Record()

        BeginTrans()


        CommitTrans("Update Complete")
    End Sub

    Sub Delete_Record()
        BeginTrans()
        Stop
        'Delete_Records("table")
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records(ByVal TABLE_NAME As String)
        'ASCDATA1.ExecuteSQL("Delete from " & TABLE_NAME _
        '    & " where x = 'x'")
    End Sub

#End Region

#Region "Popup_Menus"
    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdGLTXACPF, "BB", "Associate Plans with This Forecast", "Disassociate Actual from Forecast")
        Load_Popup_Menu(grdGLTXACPX, "B", "Add Forecast Detail")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
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
                Case "grdGLTXACPF"
                    If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
                        e.Tool.SharedProps.Visible = False
                        e.Cancel = True
                    Else
                        Dim s As Integer = grdGLTXACPF.Selected.Rows.Count
                        If s <> 0 Then
                            If grdGLTXACPF.Selected.Rows(0).Band.Index = 0 Then
                                s = 0
                            End If
                        End If
                        tlb_btn = DirectCast(tlb.Tools("Associate Plans with This Forecast"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Caption = "Associate " & CStr(s) & " Plans with This Forecast"
                        tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Index = 0) And (s <> 0)
                        tlb_btn = DirectCast(tlb.Tools("Disassociate Actual from Forecast"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Index = 1)
                    End If

                Case "grdGLTXACPX"
                    If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
                        e.Tool.SharedProps.Visible = False
                        e.Cancel = True
                    Else
                        tlb_btn = DirectCast(tlb.Tools("Add Forecast Detail"), UltraWinToolbars.ButtonTool)
                        tlb_btn.SharedProps.Visible = (grd.ActiveRow.Band.Index = 0)
                    End If

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

        Select Case e.Tool.Key
            Case "Associate Plans with This Forecast"
                Dim RECORD_NO As String = grd.ActiveRow.Cells("RECORD_NO").Value
                For Each grow As UltraWinGrid.UltraGridRow In grdGLTXACPF.Selected.Rows
                    Dim RECORD_NO_LINK As String = grow.Cells("RECORD_NO_LINK").Value
                    grow.Cells("RECORD_NO_LINK").Value = RECORD_NO
                    grow.Update()
                    If RECORD_NO_LINK = "999999" Then
                        Dim row As DataRow = dst.Tables("GLTXACPF").Rows.Find(RECORD_NO_LINK)
                        If row.GetChildRows("GLTXACPF_GLTXACPA").Count = 0 Then
                            row.Delete()
                        End If
                    End If
                Next

                Load_Details_into_Summary(True)

            Case "Disassociate Actual from Forecast"

                Dim rowGLTXACPF As DataRow = dst.Tables("GLTXACPF").Rows.Find("999999")
                If rowGLTXACPF Is Nothing Then
                    Add_Unforcasted_Row()
                End If
                grd.ActiveRow.Cells("RECORD_NO_LINK").Value = "999999"
                grd.ActiveRow.Update()

                Load_Details_into_Summary(True)

            Case "Add Forecast Detail"
                Dim RECORD_NO As String = ASCMAIN1.Next_Control_No("GLTXACPP.RECORD_NO")

                Dim rowGLTXACPY As DataRow = dst.Tables("GLTXACPY").NewRow
                rowGLTXACPY.Item("COLLECTION_CODE") = grdGLTXACPX.ActiveRow.Cells("COLLECTION_CODE").Value
                rowGLTXACPY.Item("ACCT_CODE") = grdGLTXACPX.ActiveRow.Cells("ACCT_CODE").Value
                rowGLTXACPY.Item("RECORD_NO") = RECORD_NO
                rowGLTXACPY.Item("NOTE") = "{Enter Note Here}"
                dst.Tables("GLTXACPY").Rows.Add(rowGLTXACPY)
                grdGLTXACPX.ActiveRow.ExpandAll()

            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs) Handles tlb.ToolValueChanged

    End Sub

#End Region

#Region "ABSColumn Controls"
    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "CUST_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
            Case "PROD_CATGY_CODE"
                If Not ScreenMode And e.KeyCode = Windows.Forms.Keys.Enter Then
                    Click_Command("Load", e)
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "CUST_CODE"
            '    Click_Command("Load")
        End Select
    End Sub

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "CUST_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("CUST_CODE").Text <> "" Then
                        LookUp("ARTCUST1", Absx1.txtFor("CUST_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If

            Case "PROD_CATGY_CODE"
                If EntryMode = "" Then
                    If Absx1.txtFor("PROD_CATGY_CODE").Text <> "" Then
                        LookUp("ICTPCAT1", Absx1.txtFor("PROD_CATGY_CODE").Text)
                        If cdr IsNot Nothing Then

                        End If
                    End If
                End If
        End Select
    End Sub
#End Region


    Sub Set_Month_Headings(OPS_YYYY As String)
        For M As Integer = 1 To 12
            Dim D As Date = CDate(Format(M, "00") & "/01/" & OPS_YYYY).AddMonths(6)
            Dim LEGEND As String = Format(D, "MMM") & "'" & Format(D, "yy")
            With grdGLTXACPX.DisplayLayout.Bands(0).Groups("M" & Format(M, "00"))
                .Header.Caption = LEGEND
                .Width = 60
            End With
        Next
    End Sub

    Private Sub grdGLTACCT1_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdGLTACCT1.ClickCellButton

        If e.Cell.Row.IsFilterRow OrElse Not e.Cell.Row.IsDataRow Then
            Exit Sub
        End If

        Dim ACCT_CODE = e.Cell.Row.Cells("ACCT_CODE").Value
        Dim ACCT_DESC = e.Cell.Row.Cells("ACCT_DESC").Value
        Dim ACCT_CLASS_CODE = e.Cell.Row.Cells("ACCT_CLASS_CODE").Value

        Dim COLLECTION_CODE As String = grdICTCOLL1.ActiveRow.Cells("COLLECTION_CODE").Value
        Dim COLLECTION_NAME As String = grdICTCOLL1.ActiveRow.Cells("COLLECTION_NAME").Value

        If dst.Tables("GLTXACPX").Rows.Find(New String() {COLLECTION_CODE, ACCT_CODE}) Is Nothing Then
            Dim rowGLTXACPX As DataRow = dst.Tables("GLTXACPX").NewRow
            With rowGLTXACPX
                .Item("COLLECTION_CODE") = COLLECTION_CODE
                .Item("COLLECTION_NAME") = COLLECTION_NAME
                .Item("ACCT_CODE") = ACCT_CODE
                .Item("ACCT_DESC") = ACCT_DESC
                .Item("ACCT_CLASS_CODE") = ACCT_CLASS_CODE
            End With
            dst.Tables("GLTXACPX").Rows.Add(rowGLTXACPX)
        End If

    End Sub
 

    Private Sub chkShowGL_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowGL.CheckedChanged
        splGLTXACPX.Panel1Collapsed = Not chkShowGL.Checked
        '  optClass.Visible = chkShowGL.Checked

        If chkShowGL.Checked Then
            Setup_GLTACCT1()
        End If
    End Sub

    Sub Setup_GLTXACPX()
        grdGLTXACPX.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        If optClass.Value <> "*" Then
            grdGLTXACPX.DisplayLayout.Bands(0).ColumnFilters("ACCT_CLASS_CODE").FilterConditions.Add _
              (UltraWinGrid.FilterComparisionOperator.Equals, optClass.Value)
        End If
    End Sub

    Sub Setup_GLTACCT1()
        Dim dvw As DataView = DirectCast(grdGLTACCT1.DataSource, DataTable).DefaultView
        If optClass.Value = "*" Then
            dvw.RowFilter = ""
        Else
            dvw.RowFilter = "ACCT_CLASS_CODE = '" & optClass.Value & "'"
        End If

        Sort_grdColumns(grdGLTACCT1, "ACCT_CODE")
        grdGLTACCT1.Text = "GL Accounts - " & optClass.Text

    End Sub
    Private Sub optClass_ValueChanged(sender As Object, e As EventArgs) Handles optClass.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_GLTACCT1()
        Setup_GLTXACPX()
    End Sub

    Private Sub grdGLTXACPX_AfterCellActivate(sender As Object, e As EventArgs) Handles grdGLTXACPX.AfterCellActivate

        Setup_grdGLTXACPF()
    End Sub

    Private Sub grdGLTXACPX_AfterCellUpdate(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdGLTXACPX.AfterCellUpdate
        grdGLTXACPX.ActiveRow.Update()

        If e.Cell.Row.Band.Index = 1 Then
            If e.Cell.Column.Key = "AMT_P05" Then
                Dim RECORD_NO As String = e.Cell.Row.Cells("RECORD_NO").Value
                Dim rowGLTXACPF As DataRow = dst.Tables("GLTXACPF").Rows.Find(RECORD_NO)
                rowGLTXACPF.Item("AMT") = e.Cell.Value

            End If
        End If
    End Sub

    Private Sub grdGLTXACPX_AfterRowActivate(sender As Object, e As EventArgs) Handles grdGLTXACPX.AfterRowActivate

        Setup_grdGLTXACPF()

    End Sub

    Sub Setup_grdGLTXACPF()

        Dim LEGEND As String = ""

        If grdGLTXACPX.ActiveCell Is Nothing Then
            grdGLTXACPF.Visible = False
            Exit Sub
        Else
            Dim col As String = grdGLTXACPX.ActiveCell.Column.Key
            If col.Length = 7 And Val(Mid(col, 6, 2)) >= 1 And Val(Mid(col, 6, 2)) <= 12 Then
                grdGLTXACPF.Visible = True
                MM = Mid(col, 6, 2)

                Dim M As Integer = Val(MM)
                OPS_YYYYPP = PERIODS(M)
                LEGEND = grdGLTXACPX.DisplayLayout.Bands(0).Groups("M" & MM).Header.Caption
            Else
                grdGLTXACPF.Visible = False
                Exit Sub
            End If
        End If

        grdGLTXACPF.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()

        grdGLTXACPF.DisplayLayout.Bands(0).ColumnFilters("OPS_YYYYPP").FilterConditions.Add _
            (UltraWinGrid.FilterComparisionOperator.Equals, OPS_YYYYPP)
        ' SHOULD NOT BE NEC
        grdGLTXACPF.DisplayLayout.Bands(1).ColumnFilters("OPS_YYYYPP").FilterConditions.Add _
            (UltraWinGrid.FilterComparisionOperator.Equals, OPS_YYYYPP)



        COLLECTION_CODE = grdGLTXACPX.ActiveRow.Cells("COLLECTION_CODE").Value
        ACCT_CODE = grdGLTXACPX.ActiveRow.Cells("ACCT_CODE").Value
        Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
        Dim COLLECTION_NAME As String = rowICTCOLL1.Item("COLLECTION_NAME")

        grdGLTXACPF.Text = "Forecasts & Actuals for Acct " & ACCT_CODE & ", Collection " & COLLECTION_NAME & " for " & LEGEND

        ' NEED TO SET THE ROWS TO BE SEEN IN GRDGLTXACPF

    End Sub

    Sub Load_Details_into_Summary(Optional single_record As Boolean = False)

        Dim sql As String = ""
        If single_record Then
            sql = "COLLECTION_CODE = '" & COLLECTION_CODE & "' and ACCT_CODE = '" & ACCT_CODE & "' and OPS_YYYYPP = '" & OPS_YYYYPP & "'"
            Dim rowGLTXACPX As DataRow = dst.Tables("GLTXACPX").Rows.Find(New String() {COLLECTION_CODE, ACCT_CODE})
            Dim m As Integer = Val(MM)
            ' rowGLTXACPX.Item("AMT_F" & Format(M, "00")) = 0
            rowGLTXACPX.Item("AMT_A" & Format(M, "00")) = 0
            rowGLTXACPX.Item("AMT_V" & Format(m, "00")) = 0
        End If

        For Each rowGLTXACPF As DataRow In dst.Tables("GLTXACPF").Select(sql)
            Dim COLLECTION_CODE As String = rowGLTXACPF.Item("COLLECTION_CODE")
            Dim ACCT_CODE As String = rowGLTXACPF.Item("ACCT_CODE")
            Dim RECORD_NO As String = rowGLTXACPF.Item("RECORD_NO")


            Dim OPS_YYYYPP As String = rowGLTXACPF.Item("OPS_YYYYPP")
            Dim M As Integer = Val(Mid(OPS_YYYYPP, 5, 2))

            Dim AMT As Decimal = Val(rowGLTXACPF.Item("AMT") & "")
            Dim AMT_A As Decimal = Val(rowGLTXACPF.Item("AMT_A") & "")
            Dim AMT_V As Decimal = Val(rowGLTXACPF.Item("AMT_V") & "")

            Dim rowGLTXACPY As DataRow = dst.Tables("GLTXACPY").Rows.Find(New String() {COLLECTION_CODE, ACCT_CODE, RECORD_NO})
            If rowGLTXACPY Is Nothing Then
                rowGLTXACPY = dst.Tables("GLTXACPY").Rows.Add({COLLECTION_CODE, ACCT_CODE, RECORD_NO})
                rowGLTXACPY.Item("NOTE") = rowGLTXACPF.Item("NOTE")
            End If
            rowGLTXACPY.Item("AMT_P" & Format(M, "00")) = Val(rowGLTXACPY.Item("AMT_P" & Format(M, "00")) & "") + AMT

            Dim rowGLTXACPX As DataRow = dst.Tables("GLTXACPX").Rows.Find(New String() {COLLECTION_CODE, ACCT_CODE})
            If rowGLTXACPX Is Nothing Then
                rowGLTXACPX = dst.Tables("GLTXACPX").Rows.Add({COLLECTION_CODE, ACCT_CODE})
                Dim rowICTCOLL1 As DataRow = dst.Tables("ICTCOLL1").Rows.Find(COLLECTION_CODE)
                rowGLTXACPX.Item("COLLECTION_NAME") = rowICTCOLL1.Item("COLLECTION_NAME")
                Dim rowGLTACCT1 As DataRow = dst.Tables("GLTACCT1").Rows.Find(ACCT_CODE)
                rowGLTXACPX.Item("ACCT_DESC") = rowGLTACCT1.Item("ACCT_DESC")
                rowGLTXACPX.Item("ACCT_CLASS_CODE") = rowGLTACCT1.Item("ACCT_CLASS_CODE")
            End If
            ' rowGLTXACPX.Item("AMT_F" & Format(M, "00")) = Val(rowGLTXACPX.Item("AMT_F" & Format(M, "00")) & "") + AMT
            rowGLTXACPX.Item("AMT_A" & Format(M, "00")) = Val(rowGLTXACPX.Item("AMT_A" & Format(M, "00")) & "") + AMT_A
            rowGLTXACPX.Item("AMT_V" & Format(M, "00")) = Val(rowGLTXACPX.Item("AMT_V" & Format(M, "00")) & "") + AMT_V
        Next

    End Sub

    Private Sub grdGLTXACPF_DragDrop(sender As Object, e As DragEventArgs) Handles grdGLTXACPF.DragDrop
        Dim dropIndex As Integer

        'Get the position on the grid where the dragged row(s) are to be dropped. 
        'get the grid coordinates of the row (the drop zone) 
        Dim uieOver As UIElement = grdGLTXACPF.DisplayLayout.UIElement.ElementFromPoint(grdGLTXACPF.PointToClient(New System.Drawing.Point(e.X, e.Y)))

        'get the row that is the drop zone/or where the dragged row is to be dropped 
        Dim ugrOver As UltraWinGrid.UltraGridRow = TryCast(uieOver.GetContext(GetType(UltraWinGrid.UltraGridRow), True), UltraWinGrid.UltraGridRow)

        If ugrOver IsNot Nothing Then
            dropIndex = ugrOver.Index    'index/position of drop zone in grid 

            'get the dragged row(s)which are to be dragged to another position in the grid 
            Dim SelRows As UltraWinGrid.SelectedRowsCollection = TryCast(DirectCast(e.Data.GetData(GetType(UltraWinGrid.SelectedRowsCollection)), UltraWinGrid.SelectedRowsCollection), UltraWinGrid.SelectedRowsCollection)
            'get the count of selected rows and drop each starting at the dropIndex 
            For Each aRow As UltraWinGrid.UltraGridRow In SelRows
                'move the selected row(s) to the drop zone 
                grdGLTXACPF.Rows.Move(aRow, dropIndex)
            Next
        End If
    End Sub

    Private Sub grdGLTXACPF_DragOver(sender As Object, e As DragEventArgs) Handles grdGLTXACPF.DragOver
        e.Effect = DragDropEffects.Move
        Dim grid As UltraWinGrid.UltraGrid = TryCast(sender, UltraWinGrid.UltraGrid)
        Dim pointInGridCoords As System.Drawing.Point = grid.PointToClient(New System.Drawing.Point(e.X, e.Y))

        If pointInGridCoords.Y < 20 Then
            'Scroll up
            Me.grdGLTXACPF.ActiveRowScrollRegion.Scroll(UltraWinGrid.RowScrollAction.LineUp)
        ElseIf pointInGridCoords.Y > grid.Height - 20 Then
            'Scroll down
            Me.grdGLTXACPF.ActiveRowScrollRegion.Scroll(UltraWinGrid.RowScrollAction.LineDown)
        End If
    End Sub

    Private Sub grdGLTXACPF_SelectionDrag(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles grdGLTXACPF.SelectionDrag
        grdGLTXACPF.DoDragDrop(grdGLTXACPF.Selected.Rows, DragDropEffects.Move)
    End Sub

    Private Sub grdGLTXACPX_DoubleClickCell(sender As Object, e As UltraWinGrid.DoubleClickCellEventArgs) Handles grdGLTXACPX.DoubleClickCell

        'If grdGLTXACPX.ActiveCell Is Nothing Then Exit Sub

        'Dim rowGLTXACPF As DataRow = dst.Tables("GLTXACPF").NewRow
        'RECORD_NO_ctr += 1
        'rowGLTXACPF.Item("RECORD_NO") = Format(RECORD_NO_ctr, "000000")
        'rowGLTXACPF.Item("NOTE") = "{Describe Forecast Here}"
        'rowGLTXACPF.Item("COLLECTION_CODE") = COLLECTION_CODE
        'rowGLTXACPF.Item("ACCT_CODE") = ACCT_CODE
        'rowGLTXACPF.Item("OPS_YYYYPP") = OPS_YYYYPP
        'rowGLTXACPF.Item("INIT_DATE") = Now
        'rowGLTXACPF.Item("INIT_OPER") = ASCMAIN1.USER_ID

        'dst.Tables("GLTXACPF").Rows.Add(rowGLTXACPF)
    End Sub

    Private Sub chkFullScreen_CheckedChanged(sender As Object, e As EventArgs) Handles chkFullScreen.CheckedChanged
        SplitContainer1.Panel2Collapsed = (chkFullScreen.Checked)
    End Sub

    Private Sub chkHideHistory_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideHistory.CheckedChanged
        For i As Integer = 1 To 12
            Dim MM As String = Format(i, "00")
            If MM < PP And chkHideHistory.Checked Then
                grdGLTXACPX.DisplayLayout.Bands(0).Groups("M" & MM).Hidden = True
                grdGLTXACPX.DisplayLayout.Bands(1).Columns("AMT_P" & MM).Hidden = True
            Else
                grdGLTXACPX.DisplayLayout.Bands(0).Groups("M" & MM).Hidden = False
                grdGLTXACPX.DisplayLayout.Bands(1).Columns("AMT_P" & MM).Hidden = False
            End If
        Next
    End Sub
End Class