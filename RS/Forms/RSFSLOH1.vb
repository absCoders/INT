Imports GemBox.Spreadsheet
Imports System.ComponentModel.Component
Imports System.IO
Public Class RSFSLOH1

    Dim RYW As String
    Dim RYP As String
    Dim RSTSLOH1 As String
    Dim EDI_CUSTs As New List(Of String)
    Dim LGI_CUSTs As New List(Of String)
    Dim ECOM_CUSTs As New List(Of String)
    Dim RSTSLOHC As String
    Dim RSTSLOH2 As String
    Dim RSTSLOHO As String
    Dim ALL_EDI_CUSTS As Integer
    Dim RYW_LGI As String
    Dim HUXX As String = ""
    Dim ICTITEM1 As String
    Dim grdRSTSLOH2_Calculate_Totals

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ICTPARM1")


        With dst
            ASCMAIN1.sql = "Select CUST_CODE, OPS_YYYYWW from EDT852T1 where ROWNUM < 1"
            RSTSLOHC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTSLOHC & " Add Primary Key (CUST_CODE)")
            ASCMAIN1.sql = "Select * from " & RSTSLOHC
            Create_TDA(.Tables.Add("RSTSLOHC"), RSTSLOHC, "**", 0)

            Dim CUTOFF_YEAR As String = CStr(Year(Today) - 2) & "00"

            ASCMAIN1.sql = "Select DECODE(CUST_CODE,'ECOMSALE10','ECOM',DECODE(EDI_CUST_BATCH_NO,'LGI','LGI','EDI')) CUST_TYPE, CUST_CODE" _
            & ", MAX (OPS_YYYYWW) OPS_YYYYWW" _
            & " from EDT852T1" _
            & " where OPS_YYYYWW > " & CUTOFF_YEAR _
            & " and COMPANY_CODE = '" & ASCMAIN1.SOLUTION & "' AND CUST_CODE IS NOT NULL" _
            & " group by DECODE(EDI_CUST_BATCH_NO,'LGI','LGI','EDI'),CUST_CODE"
            Create_TDA(.Tables.Add, "EDT852TC", "**", 0, False, , 2)
            .Tables("EDT852TC").Columns.Add("SEL")

            Create_RSTSLOH1("")
            ASCMAIN1.sql = "Select * from " & RSTSLOH1
            Create_TDA(.Tables.Add, "RSTSLOH1", "**", 0, False, , 1)

            HUXX = ""

            With .Tables("RSTSLOH1")
                Dim DTXX As String = ""
                For C As Integer = 1 To ALL_EDI_CUSTS ' EDI_CUSTs.Count
                    .Columns.Add("SW" & Format(C, "00"), GetType(System.Decimal), "SU" & Format(C, "00") & " * ITEM_PRICE")
                    .Columns.Add("HR" & Format(C, "00"), GetType(System.Decimal), "HU" & Format(C, "00") & " * ITEM_RETAIL_PRICE")
                    .Columns.Add("HW" & Format(C, "00"), GetType(System.Decimal), "HU" & Format(C, "00") & " * ITEM_PRICE")

                    DTXX &= "+XX" & Format(C, "00")
                Next
                For Each DT As String In New String() {"SU", "SR", "SW", "HU", "HR", "HW"}
                    .Columns.Add(DT & Format(0, "00"), GetType(System.Decimal), Replace(DTXX, "XX", DT))
                Next
                HUXX = "HU99" & Replace(DTXX, "XX", "HU")

                .Columns.Add("SC00", GetType(System.Decimal), "SU00 * ITEM_COST_STD")
                .Columns.Add("HC00", GetType(System.Decimal), "HU00 * ITEM_COST_STD")

                .Columns.Add("SW" & Format(98, "00"), GetType(System.Decimal), "SU" & Format(98, "00") & " * ITEM_PRICE")
                .Columns.Add("HR" & Format(98, "00"), GetType(System.Decimal), "HU" & Format(98, "00") & " * ITEM_RETAIL_PRICE")
                .Columns.Add("HW" & Format(98, "00"), GetType(System.Decimal), "HU" & Format(98, "00") & " * ITEM_PRICE")

                .Columns.Add("SC98", GetType(System.Decimal), "SU98 * ITEM_COST_STD")
                .Columns.Add("HC98", GetType(System.Decimal), "HU98 * ITEM_COST_STD")

                .Columns.Add("SW" & Format(99, "00"), GetType(System.Decimal), "SU" & Format(99, "00") & " * ITEM_PRICE")
                .Columns.Add("HR" & Format(99, "00"), GetType(System.Decimal), "HU" & Format(99, "00") & " * ITEM_RETAIL_PRICE")
                .Columns.Add("HW" & Format(99, "00"), GetType(System.Decimal), "HU" & Format(99, "00") & " * ITEM_PRICE")

                .Columns.Add("SC99", GetType(System.Decimal), "SU99 * ITEM_COST_STD")
                .Columns.Add("HC99", GetType(System.Decimal), "HU99 * ITEM_COST_STD")

                For Each DT As String In New String() {"SU", "SR", "SW", "SC", "HU", "HR", "HW", "HC"}
                    .Columns.Add(DT & "TT", GetType(System.Decimal), DT & "00" & " + " & DT & "99")
                Next

            End With

            ASCMAIN1.sql = "Select RSTSLOH1.ITEM_CODE, RSTSLOH1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_PICTURE_FILENAME, ICTITEM1.DEPT_CODE, ICTITEM1.CUST_CODE" & vbCrLf _
            & ", ICTITEM1.METAL_CLASS_CODE, ICTITEM1.MATL_CATGY_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.STYLE_CODE" & vbCrLf _
            & ", 0 RANK, 0 RANK_WITHIN_COLLECTION, RSTSLOH1.SUWTD, RSTSLOH1.SUMTD, RSTSLOH1.SUSTD, " & HUXX & " HUTT from " & vbCrLf _
            & RSTSLOH1 & " RSTSLOH1, ICTITEM1 where RSTSLOH1.SUWTD <> 0 and RSTSLOH1.ITEM_CATGY_CODE = 'N'" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = RSTSLOH1.ITEM_CODE"
            RSTSLOH2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTSLOH2 & " Add Primary Key (ITEM_CODE)")
            ASCMAIN1.sql = "Select ITEM_CODE, RANK, RANK_WITHIN_COLLECTION, ITEM_PICTURE_FILENAME" & vbCrLf _
            & ", SUWTD, SUMTD, SUSTD, HUTT from " & RSTSLOH2 & vbCrLf _
            & " where RANK_WITHIN_COLLECTION <= 10 or CUST_CODE is Not Null"
            Create_TDA(.Tables.Add, "RSTSLOH2", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select RSTSLOH1.ITEM_CODE, RSTSLOH1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_PICTURE_FILENAME, ICTITEM1.DEPT_CODE, ICTITEM1.CUST_CODE" & vbCrLf _
            & ", ICTITEM1.METAL_CLASS_CODE, ICTITEM1.MATL_CATGY_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.STYLE_CODE" & vbCrLf _
            & ", 0 RANK, 0 RANK_WITHIN_COLLECTION, RSTSLOH1.SUWTD, RSTSLOH1.SUMTD, RSTSLOH1.SUSTD, " & HUXX & " HUTT from " & vbCrLf _
            & RSTSLOH1 & " RSTSLOH1, ICTITEM1 where RSTSLOH1.SUWTD <> 0 and RSTSLOH1.ITEM_CATGY_CODE = 'N'" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = RSTSLOH1.ITEM_CODE"
            RSTSLOHO = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTSLOHO & " Add Primary Key (ITEM_CODE)")
            ASCMAIN1.sql = "Select ITEM_CODE, RANK, RANK_WITHIN_COLLECTION, ITEM_PICTURE_FILENAME" & vbCrLf _
            & ", SUWTD, SUMTD, SUSTD, HUTT from " & RSTSLOHO & vbCrLf _
            & " where RANK_WITHIN_COLLECTION <= 10 or CUST_CODE is Not Null"
            Create_TDA(.Tables.Add, "RSTSLOHO", "**", 0, False, , 1)

            Create_Relation("RSTSLOH1", "RSTSLOH2", "ITEM_CODE")
            Create_Relation("RSTSLOH1", "RSTSLOHO", "ITEM_CODE")
            For Each tbl As String In New String() {"RSTSLOH2", "RSTSLOHO"}
                With .Tables(tbl).Columns
                    For Each C As String In New String() _
                    {"ITEM_DESC", "COLLECTION_CODE", "PROD_CODE", "METAL_CLASS_CODE", "MATL_CATGY_CODE", "ITEM_CLASS_CODE", _
                      "ITEM_CATGY_CODE", "DEPT_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE", _
                      "ITEM_COST_STD", "PRICE_POINT_CODE", "CUST_CODE", "LAUNCH_DATE", "STYLE_CODE"}
                        ', "SUWTD", "SUMTD", "SUSTD", "HUTT"
                        If New String() {"ITEM_RETAIL_PRICE", "ITEM_PRICE", "ITEM_COST_STD", _
                                         "SUWTD", "SUMTD", "SUSTD"}.Contains(C) Then
                            .Add(C, GetType(System.Decimal), "PARENT(RSTSLOH1_" & tbl & ")." & C)
                        ElseIf C = "LAUNCH_DATE" Then
                            .Add(C, GetType(System.DateTime), "PARENT(RSTSLOH1_" & tbl & ")." & C)
                        Else
                            .Add(C, GetType(System.String), "PARENT(RSTSLOH1_" & tbl & ")." & C)
                        End If
                    Next
                End With
            Next

            Dim X As Integer = Now.Subtract(Now).TotalDays
            With .Tables("RSTSLOH2").Columns
                .Add("wksONH", GetType(System.Int32))
                .Add("BENCH_MARK", GetType(System.Int32))
                .Add("STD_SELL_THRU_PCT", GetType(System.Decimal), "IIF(ISNULL(SUSTD,0) + ISNULL(HUTT,0) = 0, 0, 100 * ISNULL(SUSTD,0) / (ISNULL(SUSTD,0) + ISNULL(HUTT,0)))")
            End With
            .Tables("RSTSLOHO").Columns.Add("ITEM_PICTURE", GetType(System.Byte()))
            .Tables("RSTSLOHO").Columns.Add("STD_SELL_THRU_PCT", GetType(System.Decimal), "IIF(ISNULL(SUSTD,0) + ISNULL(HUTT,0) = 0, 0, 100 * ISNULL(SUSTD,0) / (ISNULL(SUSTD,0) + ISNULL(HUTT,0)))")
            'Create_TDA(.Tables.Add, "ASTLIST1", "*")
            'Create_TDA(.Tables.Add, "ASTLIST2", "*")
        End With

        grdRSTSLOH1.DataSource = dst.Tables("RSTSLOH1")
        grdRSTSLOH2.DataSource = dst.Tables("RSTSLOH2")
        grdRSTSLOHO.DataSource = dst.Tables("RSTSLOHO")
        grdEDT852TC.DataSource = dst.Tables("EDT852TC")

        With grdRSTSLOH1.DisplayLayout.Bands("RSTSLOH1")
            .ColHeaderLines = 2
            Dim C As Integer = 0
            For Each COLUMN_NAME As String In New String() _
            {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "PROD_CODE", "METAL_CLASS_CODE", "MATL_CATGY_CODE", "ITEM_CLASS_CODE", _
             "ITEM_CATGY_CODE", "DEPT_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE", _
             "ITEM_COST_STD", "PRICE_POINT_CODE", "CUST_CODE", "LAUNCH_DATE", "STYLE_CODE"}
                .Columns(COLUMN_NAME).Header.VisiblePosition = C
                C += 1
            Next
            .Columns("ITEM_CODE").Header.Fixed = True

            Create_Summary(grdRSTSLOH1, "ITEM_CODE", "Count")
            For C = -2 To EDI_CUSTs.Count + 1
                For Each DT As String In New String() {"SU", "SR", "SW", "SC", "HU", "HR", "HW", "HC"}
                    Dim P As Integer
                    Dim SFX As String = IIf(C = EDI_CUSTs.Count + 1, "TT", Format(IIf(C = -1, 99, IIf(C = -2, 98, C)), "00"))
                    Dim skip_column As Boolean = False
                    If C = -2 Or C = -1 Or C = 0 Or C = EDI_CUSTs.Count + 1 Then
                        If DT = "SU" Then
                            P = .Columns(DT & SFX).Header.VisiblePosition
                        Else
                            P += 1
                            .Columns(DT & SFX).Header.VisiblePosition = P
                        End If
                    Else
                        If DT = "SC" Or DT = "HC" Then
                            skip_column = True
                        End If
                    End If
                    If skip_column Then
                    Else
                        With .Columns(DT & SFX)
                            .Width = 80
                            .Format = "###,##0"
                        End With
                        Dim COLUMN_NAME As String = DT & SFX
                        Create_Summary(grdRSTSLOH1, COLUMN_NAME)
                        Dim DD As String = IIf(Mid(DT, 1, 1) = "S", "Sls ", "OnH ") _
                                         & IIf(Mid(DT, 2, 1) = "U", "Units", IIf(Mid(DT, 2, 1) = "R", "$Rtl", IIf(Mid(DT, 2, 1) = "W", "$W/S", "$Cst")))
                        With .Columns(COLUMN_NAME)
                            Dim HDR As String = "LGI"
                            If C = 0 Then
                                HDR = "EDI"
                            ElseIf C = -2 Then
                                HDR = "ECOM"
                            ElseIf C = EDI_CUSTs.Count + 1 Then
                                HDR = "Total"
                            ElseIf C > 0 Then
                                HDR = EDI_CUSTs(C - 1)
                            End If

                            If C = 0 Then
                                .Header.Appearance.BackColor2 = Color.Orange
                            ElseIf C = EDI_CUSTs.Count + 1 Then
                                .Header.Appearance.BackColor2 = Color.Gold
                            ElseIf C = -1 Then
                                .Header.Appearance.BackColor2 = Color.LightPink
                            ElseIf C = -2 Then
                                .Header.Appearance.BackColor2 = Color.Chocolate
                            Else
                                .Header.Appearance.BackColor2 = Color.DodgerBlue
                            End If
                            .Header.Caption = HDR & vbCrLf & DD
                            '.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                            If HDR = "ECOM" Or HDR = "LGI" Or HDR = "Total" Then
                                .Hidden = True
                            End If
                            If Mid(DT, 1, 1) = "S" Then
                                If Mid(DT, 1, 1) = "S" Then
                                    .CellAppearance.BackColor = Color.White
                                    .CellAppearance.BackColor = Color.LightGreen
                                    .CellAppearance.BackGradientStyle = GradientStyle.GlassLeft20
                                Else
                                    .CellAppearance.BackColor = Color.LightGreen
                                    .CellAppearance.BackColor = Color.Green
                                    .CellAppearance.BackGradientStyle = GradientStyle.GlassLeft20
                                End If
                            Else
                                .CellAppearance.BackColor = Color.Yellow
                            End If
                        End With
                    End If
                Next
            Next
            For Each DT As String In New String() {"SU", "SR", "SW", "SC", "HU", "HR", "HW", "HC"}
                .Columns(DT & "98").Header.VisiblePosition = .Columns("SU" & "99").Header.VisiblePosition - 1
            Next
            For Each COL As String In New String() {"CUST_CODE", "LAUNCH_DATE", "SUWTD", "SUMTD", "SUSTD"}
                .Columns(COL).Hidden = True
            Next
        End With

        Cust_Setup_grdRSTSLOH2(grdRSTSLOH2, "RSTSLOH2")
        Cust_Setup_grdRSTSLOHO(grdRSTSLOHO, "RSTSLOHO")

        'Default to previous week
        'Dim DATES As DataTable = ASCDATA1.GetDataTable("SELECT YYYYWW FROM GLTPARM3 WHERE YYYYWW <= 201201 AND ROWNUM < 5 ORDER BY YYYYWW desc")
        Dim DATES As DataTable = ASCDATA1.GetDataTable("SELECT YYYYWW FROM GLTPARM3 WHERE WEEK_END_DATE <= '" & Format(Now, "dd-MMM-yyyy") & "'  ORDER BY WEEK_END_DATE desc")
        Dim PREV_WEEK As String = DATES.Rows(1).Item(0)
        Absx1.cmbFor("RYW").Value = PREV_WEEK

        'Default to 6 custom accounts for anna
        Dim EDI_ACCOUNTS_FOR_ANNA As String() = {"BLOOMIES10", "NEIMANM10", "NORDSTR10", "SAKSFIF10", "NMDIRECT", "HOLTREN10"}
        For Each row As DataRow In dst.Tables("EDT852TC").Rows
            If EDI_ACCOUNTS_FOR_ANNA.Contains(row.Item("CUST_CODE")) Then
                row.Item("SEL") = 1
            Else
                row.Item("SEL") = 0
            End If
        Next

        ASCMAIN1.Add_Value_List(grdRSTSLOH1, "ITEM_CATGY_CODE")
        ASCMAIN1.Add_Value_List(grdRSTSLOH1, "PRICE_POINT_CODE")
        ASCMAIN1.Add_Value_List(grdRSTSLOH1, "MATL_CATGY_CODE", , New String() {":", "Z:Any"}, , "SELECT MATL_CODE MATL_CATGY_CODE, MATL_DESC MATL_CATGY_DESC FROM ICTMATL1 ICTMATLA") ' "SELECT MATL_CATGY_CODE, MATL_CATGY_DESC FROM ICTMATLA")

        optCustomers.Visible = False ' this was done for JHI
        tabSales.Tabs("Newness Format").Visible = False
        tabSales.Tabs("Top Sellers").Visible = False
        chkShowAllCustomers.Visible = False
    End Sub
    Sub Cust_Setup_grdRSTSLOH2(ByVal grd As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal tblName As String)
        With grd.DisplayLayout.Bands(tblName)
            Dim C As Integer = 0
            For Each COLUMN_NAME As String In New String() _
            {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "PROD_CODE", "METAL_CLASS_CODE", "MATL_CATGY_CODE", "ITEM_CLASS_CODE", _
             "ITEM_CATGY_CODE", "DEPT_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE", _
             "ITEM_COST_STD", "PRICE_POINT_CODE", "CUST_CODE", "LAUNCH_DATE", "STYLE_CODE", "SUWTD", "SUMTD", "SUSTD", "wksONH", "BENCH_MARK"}
                .Columns(COLUMN_NAME).Header.VisiblePosition = C
                C += 1
            Next
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_CODE").Width = 195
            .Columns("ITEM_DESC").Width = 470
            .Columns("COLLECTION_CODE").Width = 200
            .Columns("BENCH_MARK").Width = 145

            Create_Summary(grd, "ITEM_CODE", "Count")
            For Each DT As String In New String() {"SUWTD", "SUMTD", "SUSTD", "BENCH_MARK", "wksONH"}
                With .Columns(DT)
                    .Width = 105
                    .Format = "###,##0"
                    .Header.Appearance.BackColor2 = Color.DodgerBlue
                    '.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                If DT = "BENCH_MARK" Or DT = "wksONH" Then
                    Continue For
                Else
                    Create_Summary(grd, DT)
                End If
            Next
            With .Columns("HUTT")
                .Width = 105
                .Format = "###,##0"
                .Header.Appearance.BackColor2 = Color.Orange
                Create_Summary(grd, "HUTT", "SUM")
                '.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            With .Columns("STD_SELL_THRU_PCT")
                .Width = 105
                .Format = "###,##0.0"
                .Header.Appearance.BackColor2 = Color.Orange
                '.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagona
                'Dim calculator As Infragistics.Win.UltraWinGrid.ICustomSummaryCalculator
                'calculator.EndCustomSummary(
                Create_Summary(grd, "STD_SELL_THRU_PCT", "Custom")
            End With

        End With

        grd.DisplayLayout.Bands(0).Columns("DEPT_CODE").SortComparer = New srtComparerSATGRNR1
        'Sort_grdColumns(grdEDT852TC, "CUST_CODE")

        ASCMAIN1.Add_Value_List(grd, "PRICE_POINT_CODE")
        ASCMAIN1.Add_Value_List(grd, "MATL_CATGY_CODE", , New String() {":", "Z:Any"}, , "SELECT MATL_CODE MATL_CATGY_CODE, MATL_DESC MATL_CATGY_DESC FROM ICTMATL1 ICTMATLA") ' "SELECT MATL_CATGY_CODE, MATL_CATGY_DESC FROM ICTMATLA")
        ASCMAIN1.Add_Value_List(grd, "ITEM_CATGY_CODE")
    End Sub
    Sub Cust_Setup_grdRSTSLOHO(ByVal grd As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal tblName As String)
        With grd.DisplayLayout.Bands(tblName)
            Dim C As Integer = 0
            For Each COLUMN_NAME As String In New String() _
            {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "PROD_CODE", "METAL_CLASS_CODE", "MATL_CATGY_CODE", "ITEM_CLASS_CODE", _
             "ITEM_CATGY_CODE", "DEPT_CODE", "ITEM_RETAIL_PRICE", "ITEM_PRICE", _
             "ITEM_COST_STD", "PRICE_POINT_CODE", "CUST_CODE", "LAUNCH_DATE", "STYLE_CODE", "SUWTD", "SUMTD", "SUSTD"} ', "wksONH", "BENCH_MARK"}
                .Columns(COLUMN_NAME).Header.VisiblePosition = C
                C += 1
            Next
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_CODE").Header.Fixed = True
            .Columns("ITEM_CODE").Width = 195
            .Columns("ITEM_DESC").Width = 470
            '.Columns("COLLECTION_CODE").Width = 200
            '.Columns("BENCH_MARK").Width = 145

            Create_Summary(grd, "ITEM_CODE", "Count")
            For Each DT As String In New String() {"SUWTD", "SUMTD", "SUSTD"} ', "BENCH_MARK", "wksONH"}
                With .Columns(DT)
                    .Width = 105
                    .Format = "###,##0"
                    .Header.Appearance.BackColor2 = Color.DodgerBlue
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                End With
                'If DT = "BENCH_MARK" Or DT = "wksONH" Then
                '    Continue For
                'Else
                Create_Summary(grd, DT)
                'End If
            Next
            With .Columns("HUTT")
                .Width = 105
                .Format = "###,##0"
                .Header.Appearance.BackColor2 = Color.Orange
                Create_Summary(grd, "HUTT", "SUM")
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
            End With
            With .Columns("STD_SELL_THRU_PCT")
                .Width = 105
                .Format = "###,##0.0"
                .Header.Appearance.BackColor2 = Color.Orange
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                'Dim calculator As Infragistics.Win.UltraWinGrid.ICustomSummaryCalculator
                'calculator.EndCustomSummary(
                Create_Summary(grd, "STD_SELL_THRU_PCT", "Custom")
            End With
            With .Columns("ITEM_PICTURE")
                .Header.Caption = "Picture"
                .Style = UltraWinGrid.ColumnStyle.Image
            End With
        End With

        grd.DisplayLayout.Bands(0).Columns("DEPT_CODE").SortComparer = New srtComparerSATGRNR1
        'Sort_grdColumns(grdEDT852TC, "CUST_CODE")

        ASCMAIN1.Add_Value_List(grd, "PRICE_POINT_CODE")
        ASCMAIN1.Add_Value_List(grd, "MATL_CATGY_CODE", , New String() {":", "Z:Any"}, , "SELECT MATL_CODE MATL_CATGY_CODE, MATL_DESC MATL_CATGY_DESC FROM ICTMATL1 ICTMATLA") ' "SELECT MATL_CATGY_CODE, MATL_CATGY_DESC FROM ICTMATLA")
        ASCMAIN1.Add_Value_List(grd, "ITEM_CATGY_CODE")
    End Sub
    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                If EMsg = "" Then
                    RYW = Absx1.cmbFor("RYW").Value
                    Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
                    RYP = rowGLTPARM3.Item("YYYYPP")
                End If

                If chkShowAllCustomers.Checked Then
                    Dim N As Integer = dst.Tables("EDT852TC").Select("SEL='1'").Length
                    If N > 30 Then
                        EMsg &= vbCr & "You have selected " & CStr(N) & " Customers. Cannot Show more than 30 Customers"
                    End If
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

            Case "Load"
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Scope").Visible = Not ScreenMode
                .Groups("Options").Visible = ScreenMode
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        'grdRSTSLOH1.Visible = tf
        tabSales.Visible = tf
        Setup_tabSales()
        Set_Read_Only_for_ctl(chkHISTCAT, ScreenMode)

        If ScreenMode Then
            grdEDT852TC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        Else
            Clear_Record()
            grdEDT852TC.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("RSTSLOH1").Rows.Clear()
        dst.Tables("RSTSLOH2").Rows.Clear()
        dst.Tables("RSTSLOHO").Rows.Clear()
        dst.EnforceConstraints = True
        grdRSTSLOH1.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Set_CUST_TYPE()

    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Loading Data")
        Application.DoEvents()

        Call Save_Header_Fields(UltraGroupBox1)

        ICTITEM1 = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP, chkHISTCAT.Checked)

        EnforceConstraints(False)

        Create_RSTSLOH1(RYW)

        Fill_Records("RSTSLOH1")
        Fill_Records("RSTSLOH2")
        Fill_Records("RSTSLOHO")

        EnforceConstraints(True)

        Setup_Pictures()

        ASCMAIN1.Progress("Now Setting Up Screen")

        grdRSTSLOH1.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdRSTSLOH1, "ITEM_CODE")
        'grdRSTSLOH1.DisplayLayout.Bands(0).SortedColumns.Add("DEPT_CODE", False, True)
        'grdRSTSLOH1.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CATGY_CODE", False, True)
        'grdRSTSLOH1.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
        'grdRSTSLOH1.Rows.ExpandAll(True)

        grdRSTSLOH1.Text = optTD.Text & " Sales & On Hand by Item as of " & Absx1.cmbFor("RYW").Text

        If RYW_LGI <> RYW Then
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW_LGI)
            grdRSTSLOH1.Text &= "; LGI data as of " & rowGLTPARM3.Item("LEGEND")
        End If

        Dim EDIs As Int64 = dst.Tables("EDT852TC").Select("CUST_TYPE = 'EDI'").Length
        Dim EDIs_incl As Int64 = dst.Tables("EDT852TC").Select("CUST_TYPE = 'EDI' and SEL = '1'").Length
        Dim LGIs As Int64 = dst.Tables("EDT852TC").Select("CUST_TYPE = 'LGI'").Length
        Dim LGIs_incl As Int64 = dst.Tables("EDT852TC").Select("CUST_TYPE = 'LGI' and SEL = '1'").Length
        Dim ECOMs As Int64 = dst.Tables("EDT852TC").Select("CUST_TYPE = 'ECOM'").Length
        Dim ECOMs_incl As Int64 = dst.Tables("EDT852TC").Select("CUST_TYPE = 'ECOM' and SEL = '1'").Length

        Dim ix As String = ""
        If EDIs <> EDIs_incl Or 1 = 1 Then ' anna want's to always see customers included
            ix = ""
            If EDIs - EDIs_incl < EDIs_incl And 1 <> 1 Then ' anna want's to always see customers included
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = 'EDI' and ISNULL(SEL,'0') = '0'", "CUST_CODE")
                    ix &= "," & row.Item("CUST_CODE")
                Next
                grdRSTSLOH1.Text &= "; EDIs excl: " & Mid(ix, 2)
            Else
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = 'EDI' and SEL = '1'", "CUST_CODE")
                    ix &= "," & row.Item("CUST_CODE")
                Next
                grdRSTSLOH1.Text &= "; EDIs incl: " & Mid(ix, 2)
            End If
        End If
        If LGIs <> LGIs_incl Then '  Or 1 = 1 Then ' anna want's to always see customers included
            ix = ""
            If LGIs - LGIs_incl < LGIs_incl And 1 <> 1 Then ' anna want's to always see customers included
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = 'LGI' and ISNULL(SEL,'0') = '0'", "CUST_CODE")
                    ix &= "," & row.Item("CUST_CODE")
                Next
                grdRSTSLOH1.Text &= "; LGIs excl: " & Mid(ix, 2)
            Else
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = 'LGI' and SEL = '1'", "CUST_CODE")
                    ix &= "," & row.Item("CUST_CODE")
                Next
                grdRSTSLOH1.Text &= "; LGIs incl: " & Mid(ix, 2)
            End If
        End If
        If ECOMs <> ECOMs_incl Then ' Or 1 = 1 Then ' anna want's to always see customers included
            ix = ""
            If ECOMs - ECOMs_incl < ECOMs_incl And 1 <> 1 Then ' anna want's to always see customers included
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = 'ECOM' and ISNULL(SEL,'0') = '0'", "CUST_CODE")
                    ix &= "," & row.Item("CUST_CODE")
                Next
                grdRSTSLOH1.Text &= "; ECOMs excl: " & Mid(ix, 2)
            Else
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = 'ECOM' and SEL = '1'", "CUST_CODE")
                    ix &= "," & row.Item("CUST_CODE")
                Next
                grdRSTSLOH1.Text &= "; ECOMs incl: " & Mid(ix, 2)
            End If
        End If

        For Each ROW As DataRow In dst.Tables("RSTSLOH2").Rows
            If ROW.Item("LAUNCH_DATE") Is Null Then
                Continue For
            Else
                Dim LAUNCH_DATE As String() = Split(ROW.Item("LAUNCH_DATE"), " ")

                Dim D_START As Date = LAUNCH_DATE(0)

                Dim SELECTED_WEEK As DataRow = LookUp("GLTPARM3", Absx1.cmbFor("RYW").Value)
                Dim D_END As Date = SELECTED_WEEK.Item("WEEK_END_DATE")
                Dim wksONH As Integer = (D_END.Subtract(D_START).TotalDays) / 7
                Dim ONH As Integer = ROW.Item("SUSTD") + ROW.Item("HUTT")
                Dim BENCH_MARK As Integer = ONH * wksONH * 0.03

                ROW.Item("wksONH") = wksONH
                ROW.Item("BENCH_MARK") = BENCH_MARK

            End If
        Next

        For Each ROW As DataRow In dst.Tables("RSTSLOHO").Rows
            If ROW.Item("LAUNCH_DATE") Is Null Then
                Continue For
            Else
                'Dim LAUNCH_DATE As String() = Split(ROW.Item("LAUNCH_DATE"), " ")

                'Dim D_START As Date = LAUNCH_DATE(0)

                Dim SELECTED_WEEK As DataRow = LookUp("GLTPARM3", Absx1.cmbFor("RYW").Value)
                Dim D_END As Date = SELECTED_WEEK.Item("WEEK_END_DATE")
                'Dim wksONH As Integer = (D_END.Subtract(D_START).TotalDays) / 7
                Dim ONH As Integer = ROW.Item("SUSTD") + ROW.Item("HUTT")
                'Dim BENCH_MARK As Integer = ONH * wksONH * 0.03

                'ROW.Item("wksONH") = wksONH
                'ROW.Item("BENCH_MARK") = BENCH_MARK
            End If
        Next

        Setup_grdRSTSLOH2()
        Setup_grdRSTSLOHO()

        Show_Columns()

        Me.Cursor = Cursors.Default
        Call ASCMAIN1.Progress("")

    End Sub

    Sub Update_Record()
        Call BeginTrans()
        Call CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdRSTSLOH1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdRSTSLOH2, "SSSS", "Show Filter", "Show GroupBox", "Show Pins", "Show Rank Columns")
        Load_Popup_Menu(grdRSTSLOHO, "SSSS", "Show Filter", "Show GroupBox", "Show Pins", "Show Rank Columns")
        Load_Popup_Menu(grdEDT852TC, "BBBBB", "Select All", "Clear All", "Save List", "Load List", "Maintain Lists")
    End Sub

    Public Overrides Sub tlb_beforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Rank Columns") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Rank Columns"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns("RANK_WITHIN_COLLECTION").Hidden
        End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"

            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Rank Columns"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns("RANK_WITHIN_COLLECTION").Hidden = Not tlb_sbt.Checked
                grd.DisplayLayout.Bands(0).Columns("RANK").Hidden = Not tlb_sbt.Checked

            Case "Select All", "Clear All"
                Set_SEL(IIf(e.Tool.Key = "Select All", "1", "0"))

            Case "Load List"
                Dim LIST_CODE As String = View_Lookup(Nothing, "LIST_CODE", "", "", "COLUMN_NAME = 'CUST_CODE'")
                If LIST_CODE <> "" Then
                    ASCMAIN1.sql = "Select CODE_VALUE from ASTLIST2 where LIST_CODE = '" & LIST_CODE & "'"
                    Dim CUST_CODEs As New List(Of String)
                    For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                        CUST_CODEs.Add(row.Item("CODE_VALUE"))
                    Next
                    For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = '" & optCustomers.Value & "'")
                        row.Item("SEL") = "0"
                        If CUST_CODEs.Contains(row.Item("CUST_CODE")) Then
                            row.Item("SEL") = "1"
                        End If
                    Next
                End If

            Case "Save List"
                Dim PSCs As String = ""
                For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = '" & optCustomers.Value & "' and SEL = '1'")
                    Dim CUST_CODE As String = row.Item("CUST_CODE")
                    PSCs &= Chr(0) & CUST_CODE
                Next

                If PSCs.Length = 0 Then
                    MsgBox("No Customers Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE")
                ASCMAIN1.CodeSelector.MultipleSelections = True
                ASCMAIN1.CodeSelector.PreviouslySelectedCodes0 = PSCs
                Using frmASFCODE1 As New ASFCODE1
                    frmASFCODE1.EntryMode = "S"
                    frmASFCODE1.ShowDialog()
                End Using

            Case "Maintain Lists"
                ASCMAIN1.CodeSelector.Get_SQL("CUST_CODE")
                ASCMAIN1.CodeSelector.MultipleSelections = True
                Using frmASFCODE1 As New ASFCODE1
                    frmASFCODE1.ShowDialog()
                End Using

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Item Status Inquiry"
            '    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
            '    Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select


    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "OPS_YYYYPP"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Sub Create_RSTSLOH1(ByVal RYW As String)

        If RYW = "" Then
            Fill_Records("EDT852TC")
            ALL_EDI_CUSTS = dst.Tables("EDT852TC").Select("CUST_TYPE = 'EDI'").Length
            ALL_EDI_CUSTS = 30 ' using a hard coded number of customers to accomodate the option to show non-EDI mixed with EDI customers individually, up to a max of 30 customers

            'Initialize_EDT852TC()
            For Each rowEDT852TC As DataRow In dst.Tables("EDT852TC").Select("", "CUST_CODE")
                rowEDT852TC.Item("SEL") = "1"
                Dim CUST_CODE As String = rowEDT852TC.Item("CUST_CODE")
                If rowEDT852TC.Item("CUST_TYPE") = "EDI" Then
                    EDI_CUSTs.Add(CUST_CODE)
                ElseIf rowEDT852TC.Item("CUST_TYPE") = "ECOM" Then
                    ECOM_CUSTs.Add(CUST_CODE)
                Else
                    LGI_CUSTs.Add(CUST_CODE)
                End If
            Next
        Else
            EDI_CUSTs.Clear()
            LGI_CUSTs.Clear()
            ECOM_CUSTs.Clear()
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLOHC)
            dst.Tables("RSTSLOHC").Rows.Clear()

            For Each rowEDT852TC As DataRow In dst.Tables("EDT852TC").Select("SEL = '1'", "CUST_CODE")
                Dim CUST_CODE As String = rowEDT852TC.Item("CUST_CODE")
                If chkShowAllCustomers.Checked Then
                    EDI_CUSTs.Add(CUST_CODE)
                Else
                    If rowEDT852TC.Item("CUST_TYPE") = "EDI" Then
                        EDI_CUSTs.Add(CUST_CODE)
                    ElseIf rowEDT852TC.Item("CUST_TYPE") = "ECOM" Then
                        ECOM_CUSTs.Add(CUST_CODE)
                    Else
                        LGI_CUSTs.Add(CUST_CODE)
                    End If
                End If
                dst.Tables("RSTSLOHC").Rows.Add(New String() {CUST_CODE})
            Next
            Update_Record_TDA("RSTSLOHC")

            ASCMAIN1.sql = "Update " & RSTSLOHC & " RSTSLOHC Set OPS_YYYYWW = " _
            & " (Select Max (OPS_YYYYWW) from EDT852T1 " _
            & " where CUST_CODE = RSTSLOHC.CUST_CODE AND OPS_YYYYWW <= '" & RYW & "' and OPS_YYYYWW > '" & ASCMAIN1.Week_Calc(RYW, -6) & "')"
            ASCDATA1.ExecuteSQL()

        End If
        Dim RYWM As String = ""
        Dim RYWS As String = ""
        Dim RYWY As String = Mid(RYW, 1, 4) & "01"
        If RYW <> "" Then
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
            RYWM = ASCMAIN1.Week_Calc(RYW, -1 * (Val(rowGLTPARM3.Item("REL_WEEK") & "") - 1))
            If Mid(RYW, 5, 2) > "26" Then
                RYWS = Mid(RYW, 1, 4) & "27"
            Else
                RYWS = Mid(RYW, 1, 4) & "01"
            End If
        End If
        Dim RYWF As String = ""
        If Mid(RYW, 5, 2) > "26" Then
            RYWF = Mid(RYW, 1, 4) & "27"
        Else
            RYWF = Format(Val(Mid(RYW, 1, 4)) - 1, "0000") & "27"
        End If

        Dim RYW_START As String = RYWY
        If optTD.Value = "FYTD" And RYWF < RYWY Then
            RYW_START = RYWF
        End If

        Dim RYWX As String = IIf(optTD.Value = "FYTD", RYWF, IIf(optTD.Value = "MTD", RYWM, IIf(optTD.Value = "YTD", RYWY, RYWS)))
        Dim SQLX As String = "CASE WHEN RSTRETL1.OPS_YYYYWW >= '" & RYWX & "' AND RSTRETL1.OPS_YYYYWW <= '" & RYW & "' THEN RSTRETL1.X ELSE 0 END"

        RYW_LGI = dst.Tables("EDT852TC").Compute("MAX(OPS_YYYYWW)", "CUST_TYPE = 'LGI'") & String.Empty
        If RYW_LGI > RYW OrElse RYW_LGI = "" Then
            RYW_LGI = RYW
        End If

        Dim SQLC As String = ""
        For i As Integer = 1 To EDI_CUSTs.Count
            Dim CUST_CODE As String = EDI_CUSTs(i - 1)
            SQLC &= ", Sum (DECODE(RSTRETL1.CUST_CODE,'" & CUST_CODE & "'," & Replace(SQLX, "RSTRETL1.X", "RSTRETL1.QTY_SOLD") & ",0)) SU" & Format(i, "00") & vbCrLf
            SQLC &= ", Sum (DECODE(RSTRETL1.CUST_CODE,'" & CUST_CODE & "'," & Replace(SQLX, "RSTRETL1.X", "RSTRETL1.AMT_SOLD") & ",0)) SR" & Format(i, "00") & vbCrLf
            'SQLC &= ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW & "' THEN DECODE(RSTRETL1.CUST_CODE,'" & CUST_CODE & "',RSTRETL1.QTY_EOW,0) ELSE 0 END) HU" & Format(i, "00") & vbCrLf
            SQLC &= ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW = NVL(RSTSLOHC.OPS_YYYYWW,'" & RYW & "') THEN DECODE(RSTRETL1.CUST_CODE,'" & CUST_CODE & "',RSTRETL1.QTY_EOW,0) ELSE 0 END) HU" & Format(i, "00") & vbCrLf
        Next

        If EDI_CUSTs.Count < ALL_EDI_CUSTS Then
            For i As Integer = EDI_CUSTs.Count + 1 To ALL_EDI_CUSTS
                SQLC &= ", 0 SU" & Format(i, "00") & vbCrLf
                SQLC &= ", 0 SR" & Format(i, "00") & vbCrLf
                SQLC &= ", 0 HU" & Format(i, "00") & vbCrLf
            Next
        End If

        SQLC &= ", Sum (DECODE(EDT852T1.EDI_CUST_BATCH_NO,'LGI'," & Replace(SQLX, "RSTRETL1.X", "RSTRETL1.QTY_SOLD") & ",0)) SU" & Format(99, "00") & vbCrLf
        SQLC &= ", Sum (DECODE(EDT852T1.EDI_CUST_BATCH_NO,'LGI'," & Replace(SQLX, "RSTRETL1.X", "RSTRETL1.AMT_SOLD") & ",0)) SR" & Format(99, "00") & vbCrLf
        SQLC &= ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW_LGI & "' THEN DECODE(EDT852T1.EDI_CUST_BATCH_NO,'LGI',RSTRETL1.QTY_EOW,0) ELSE 0 END) HU" & Format(99, "00") & vbCrLf

        SQLC &= ", Sum (DECODE(EDT852T1.EDI_CUST_BATCH_NO,'ECOM'," & Replace(SQLX, "RSTRETL1.X", "RSTRETL1.QTY_SOLD") & ",0)) SU" & Format(98, "00") & vbCrLf
        SQLC &= ", Sum (DECODE(EDT852T1.EDI_CUST_BATCH_NO,'ECOM'," & Replace(SQLX, "RSTRETL1.X", "RSTRETL1.AMT_SOLD") & ",0)) SR" & Format(98, "00") & vbCrLf
        SQLC &= ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW_LGI & "' THEN DECODE(EDT852T1.EDI_CUST_BATCH_NO,'ECOM',RSTRETL1.QTY_EOW,0) ELSE 0 END) HU" & Format(98, "00") & vbCrLf


        SQLC &= ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW = '" & RYW & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SUWTD" & vbCrLf
        SQLC &= ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW >= '" & RYWM & "' AND RSTRETL1.OPS_YYYYWW <= '" & RYW & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SUMTD" & vbCrLf
        SQLC &= ", Sum (CASE WHEN RSTRETL1.OPS_YYYYWW >= '" & RYWS & "' AND RSTRETL1.OPS_YYYYWW <= '" & RYW & "' THEN RSTRETL1.QTY_SOLD ELSE 0 END) SUSTD" & vbCrLf

        '& " where RSTRETL1.OPS_YYYYWW >= '" & IIf(optTD.Value = "MTD", RYWM, IIf(optTD.Value = "YTD", RYWY, RYWS)) & "'" _

        Dim SQL As String = "Select ITEM_CODE" & SQLC _
        & " from RSTRETL1,EDT852T1," & RSTSLOHC & " RSTSLOHC" _
        & " where RSTRETL1.OPS_YYYYWW >= '" & RYW_START & "'" _
        & "   and RSTRETL1.OPS_YYYYWW <= '" & RYW & "'" & vbCrLf _
        & "   and RSTRETL1.EDI_DOC_SEQ_NO = EDT852T1.EDI_DOC_SEQ_NO" & vbCrLf _
        & "   and RSTSLOHC.CUST_CODE = RSTRETL1.CUST_CODE" & vbCrLf _
        & " group by RSTRETL1.ITEM_CODE" & vbCrLf

        '        & "   and RSTRETL1.CUST_CODE in (Select CUST_CODE from " & RSTSLOHC & ")" & vbCrLf _

        ASCMAIN1.sql = "Select X.*" & vbCrLf _
        & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE,  ICTITEM1.PROD_CODE, ICTITEM1.METAL_CLASS_CODE, ICTITEM1.MATL_CATGY_CODE, ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
        & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.DEPT_CODE" & vbCrLf _
        & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_PRICE, ICTITEM1.ITEM_COST_STD" & vbCrLf _
        & ", PRICE_POINT(ICTITEM1.ITEM_RETAIL_PRICE) PRICE_POINT_CODE" & vbCrLf _
        & ", CUST_CODE, LAUNCH_DATE, STYLE_CODE" & vbCrLf _
        & " from " & ICTITEM1 & " ICTITEM1,(" & vbCrLf & SQL & ") X" & vbCrLf _
        & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE "

        If RYW = "" Then
            RSTSLOH1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & RSTSLOH1 & " Add Primary Key (ITEM_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & RSTSLOH1)
            ASCDATA1.ExecuteSQL("Insert into " & RSTSLOH1 & " " & ASCMAIN1.sql)

            If chkShowAllCustomers.Checked Then
                Dim Z As String = "SU99=0,SR99=0,HU99=0"
                ASCDATA1.ExecuteSQL("Update " & RSTSLOH1 & " Set " & Z)
                ASCDATA1.ExecuteSQL("Update " & RSTSLOH1 & " Set " & Replace(Z, "99", "98"))
            End If

            Calculate_Top_Sellers(grdRSTSLOH1, RSTSLOH2)
            Calculate_Top_Sellers(grdRSTSLOH1, RSTSLOHO)
        End If
    End Sub

    Private Sub grdRSTSLOH1_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdRSTSLOH1.InitializeRow
        'For Y As Integer = 1 To 50
        '    Dim YY As String = "Y" & Format(Y, "00") & "P"
        '    If Val(e.Row.Cells(YY).Value & "") < 0 Then
        '        e.Row.Cells(YY).Appearance.ForeColor = Drawing.Color.Red
        '    Else
        '        e.Row.Cells(YY).Appearance.ForeColor = Drawing.Color.Empty
        '    End If
        'Next
    End Sub

    Private Sub chkEDICustomers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEDICustomers.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkUnits.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkOnHand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOnHand.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Sub Show_Columns()

        With grdRSTSLOH1.DisplayLayout.Bands(0)
            For C As Integer = -2 To EDI_CUSTs.Count + 1
                Dim v As Boolean = True
                Dim CX As String = Format(C, "00")
                If C = -1 Then
                    CX = "99"
                ElseIf C = -2 Then
                    CX = "98"
                ElseIf C = EDI_CUSTs.Count + 1 Then
                    CX = "TT"
                Else
                    If C > 0 And Not chkEDICustomers.Checked And Not chkShowAllCustomers.Checked Then
                        v = False
                    End If
                End If
                If chkShowAllCustomers.Checked Then
                    If CX = "98" Or CX = "99" Or CX = "00" Then
                        v = False
                    End If
                End If
                For Each DT As String In New String() {"SU", "SR", "SW", "SC"}
                    Dim V2 As Boolean = True
                    If DT = "SW" And C > 0 Then V2 = False
                    If Mid(DT, 2, 1) = "U" And Not chkUnits.Checked Then V2 = False
                    If Mid(DT, 2, 1) = "R" And Not chkRetail.Checked Then V2 = False
                    If Mid(DT, 2, 1) = "W" And Not chkWS.Checked Then V2 = False
                    If Mid(DT, 2, 1) = "C" And Not chkCosts.Checked Then V2 = False
                    If C >= 1 And C <= EDI_CUSTs.Count And DT = "SC" Then
                    Else
                        .Columns(DT & CX).Hidden = Not (v And chkSales.Checked And V2)
                        .Columns(DT & CX).Hidden = (CX = "98" Or CX = "99" Or CX = "TT")
                    End If
                Next
                For Each DT As String In New String() {"HU", "HR", "HW", "HC"}
                    Dim V2 As Boolean = True
                    If (DT = "HR" Or DT = "HW") And C > 0 Then V2 = False
                    If Mid(DT, 2, 1) = "U" And Not chkUnits.Checked Then V2 = False
                    If Mid(DT, 2, 1) = "R" And Not chkRetail.Checked Then V2 = False
                    If Mid(DT, 2, 1) = "W" And Not chkWS.Checked Then V2 = False
                    If Mid(DT, 2, 1) = "C" And Not chkCosts.Checked Then V2 = False
                    If C >= 1 And C <= EDI_CUSTs.Count And DT = "HC" Then
                    Else
                        .Columns(DT & CX).Hidden = Not (v And chkOnHand.Checked And V2)
                        .Columns(DT & CX).Hidden = (CX = "98" Or CX = "99" Or CX = "TT")
                    End If
                    If C = -2 Then
                        .Columns(DT & CX).Hidden = True
                    End If
                Next
            Next

            If EDI_CUSTs.Count < ALL_EDI_CUSTS Then
                For i As Integer = EDI_CUSTs.Count + 1 To ALL_EDI_CUSTS
                    For Each DT As String In New String() {"SU", "SR", "SW", "HU", "HR", "HW"}
                        .Columns(DT & Format(i, "00")).Hidden = True
                    Next
                Next
            End If

            For C = 1 To EDI_CUSTs.Count
                For Each DT As String In New String() {"SU", "SR", "SW", "HU", "HR", "HW"}
                    Dim SFX As String = Format(C, "00")
                    Dim COLUMN_NAME As String = DT & SFX
                    Dim DD As String = IIf(Mid(DT, 1, 1) = "S", "Sls ", "OnH ") _
                                     & IIf(Mid(DT, 2, 1) = "U", "Units", IIf(Mid(DT, 2, 1) = "R", "$Rtl", IIf(Mid(DT, 2, 1) = "W", "$W/S", "$Cst")))
                    With .Columns(COLUMN_NAME)
                        Dim HDR As String = EDI_CUSTs(C - 1)
                        .Header.Caption = HDR & vbCrLf & DD
                    End With
                Next
            Next
        End With



        'Format Newness Selling for Anna
        If optCustomers.CheckedItem.DisplayText = "EDI" Then
            With dst.Tables("RSTSLOH2")
                Dim COLUMNS_FOR_ANNA() As String = {"ITEM_CODE", "ITEM_DESC", "COLLECTION_CODE", "ITEM_RETAIL_PRICE", "LAUNCH_DATE", _
                                                    "SUMTD", "STD_SELL_THRU_PCT", "SUWTD", "SUSTD", "HUTT", "wksONH", "BENCH_MARK"}
                For Each col As DataColumn In .Columns
                    If COLUMNS_FOR_ANNA.Contains(col.ColumnName.Trim()) Then
                        grdRSTSLOH2.DisplayLayout.Bands(0).Columns(col.ColumnName).Hidden = False
                    Else
                        grdRSTSLOH2.DisplayLayout.Bands(0).Columns(col.ColumnName).Hidden = True
                    End If
                Next
            End With
        End If

        dst.Tables("RSTSLOH2").Columns("wksONH").Caption = "Wks OnH"
        dst.Tables("RSTSLOH2").Columns("BENCH_MARK").Caption = "Sell Thru" & vbCrLf & "Bench Mark"

        If chkShowAllCustomers.Checked Then
            For Each CT As String In New String() {"TT", "98", "99"}
                For Each DT As String In New String() {"SU", "SR", "SW", "HU", "HR", "HW"}
                    'For Each gcol As UltraWinGrid.UltraGridColumn In grdRSTSLOHO.DisplayLayout.Bands(0).Columns
                    '    Debug.Print(gcol.Key)
                    'Next
                    If grdRSTSLOHO.DisplayLayout.Bands(0).Columns.Contains(DT & CT) Then
                        grdRSTSLOHO.DisplayLayout.Bands(0).Columns(DT & CT).Hidden = True
                    End If
                 Next
            Next
        End If
    End Sub
    Public Overrides Function CustomSummary_End( _
  ByVal summarySettings As UltraWinGrid.SummarySettings, _
  ByVal rows As UltraWinGrid.RowsCollection, _
  ByVal CustomValue As Double, _
  ByVal grd As UltraWinGrid.UltraGrid) As Double

        CustomValue = 0
        Dim TOTALS As New Dictionary(Of String, Decimal)

        Select Case grd.Name
            Case "grdRSTSLOH2"
                Dim KEY As String = summarySettings.Key
                If KEY.EndsWith("_SELLTHRU") Then
                    TOTALS.Add("U", 0)
                    TOTALS.Add("Q", 0)
                    grdRSTSLOH1_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("U") + TOTALS("Q") <> 0 Then CustomValue = 100 * TOTALS("U") / (TOTALS("U") + TOTALS("Q"))

                ElseIf KEY.EndsWith("_PCT") Then
                    TOTALS.Add("Q", 0)
                    TOTALS.Add("QX", 0)
                    grdRSTSLOH1_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("Q") <> 0 Then CustomValue = 100 * TOTALS("QX") / TOTALS("Q")
                End If
            Case "grdRSTSLOHO"
                Dim KEY As String = summarySettings.Key
                If KEY.EndsWith("_SELLTHRU") Then
                    TOTALS.Add("U", 0)
                    TOTALS.Add("Q", 0)
                    grdRSTSLOH1_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("U") + TOTALS("Q") <> 0 Then CustomValue = 100 * TOTALS("U") / (TOTALS("U") + TOTALS("Q"))

                ElseIf KEY.EndsWith("_PCT") Then
                    TOTALS.Add("Q", 0)
                    TOTALS.Add("QX", 0)
                    grdRSTSLOH1_Calculate_Totals(rows, TOTALS, KEY)
                    If TOTALS("Q") <> 0 Then CustomValue = 100 * TOTALS("QX") / TOTALS("Q")
                End If
            Case Else
                MsgBox("CustomSummary_End " & grd.Name)
        End Select

        Return CustomValue
    End Function

    Sub grdRSTSLOH1_Calculate_Totals( _
       ByVal rows As UltraWinGrid.RowsCollection, _
       ByRef TOTALS As Dictionary(Of String, Decimal), _
       ByVal KEY As String)

        For Each grow2 As UltraWinGrid.UltraGridRow In rows
            If grow2.IsGroupByRow Then
                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow2, UltraWinGrid.UltraGridGroupByRow)
                grdRSTSLOH1_Calculate_Totals(gbrow.Rows, TOTALS, KEY)
            Else
                If KEY.EndsWith("_SELLTHRU") Then
                    TOTALS("U") += Val(grow2.Cells(Mid(KEY, 1, 1) & "U").Value & "")
                    TOTALS("Q") += Val(grow2.Cells(Mid(KEY, 1, 1) & "Q").Value & "")
                ElseIf KEY.EndsWith("_PCT") Then
                    TOTALS("Q") += 100 'Val(grow2.Cells(KEY).Value & "") 'Mid(KEY, 1, 2)).Value & "")
                    TOTALS("QX") += Val(grow2.Cells(KEY).Value & "") 'Mid(KEY, 1, 3)).Value & "")
                End If
            End If
        Next
    End Sub


    Private Sub optCustomers_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCustomers.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Set_CUST_TYPE()
    End Sub

    Sub Set_CUST_TYPE()
        Dim dvw As DataView = DirectCast(grdEDT852TC.DataSource, DataTable).DefaultView
        dvw.RowFilter = "CUST_TYPE = '" & optCustomers.Value & "'"
    End Sub

    Sub Set_SEL(ByVal SEL As String)
        For Each row As DataRow In dst.Tables("EDT852TC").Select("CUST_TYPE = '" & optCustomers.Value & "'")
            row.Item("SEL") = SEL
        Next
    End Sub

    Private Sub chkCosts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCosts.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkRetailSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRetail.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkWSSales_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkWS.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub chkSales_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSales.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Show_Columns()
    End Sub

    Private Sub optTD_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTD.ValueChanged

    End Sub

    Private Sub chkTopNewOnly_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkTopNewOnly.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Re-Calculating Top Sellers")

        Calculate_Top_Sellers(grdRSTSLOH2, RSTSLOH2)
        Calculate_Top_Sellers(grdRSTSLOHO, RSTSLOH1)
        EnforceConstraints(False)
        Fill_Records("RSTSLOH2")
        Fill_Records("RSTSLOHO")
        EnforceConstraints(True)
        Setup_grdRSTSLOH2()
        Setup_grdRSTSLOHO()
        Setup_Pictures()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Calculate_Top_Sellers(ByVal grd As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal tempTable As String)

        ASCDATA1.ExecuteSQL("Truncate Table " & tempTable)

        ASCMAIN1.sql = "Select RSTSLOH1.ITEM_CODE, RSTSLOH1.COLLECTION_CODE, ICTITEM1.ITEM_PICTURE_FILENAME" _
        & ", ICTITEM1.DEPT_CODE, ICTITEM1.CUST_CODE, ICTITEM1.METAL_CLASS_CODE, ICTITEM1.MATL_CATGY_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.STYLE_CODE" _
        & ", RSTSLOH1.SUWTD, RSTSLOH1.SUMTD, RSTSLOH1.SUSTD, " & HUXX & " HUTT from " _
        & RSTSLOH1 & " RSTSLOH1, " & ICTITEM1 & " ICTITEM1 where (ICTITEM1.CUST_CODE is Not Null or ((RSTSLOH1.SUSTD <> 0 or ICTITEM1.ITEM_CLASS_CODE = 'RING') " _
        & IIf(chkTopNewOnly.Checked, "and RSTSLOH1.ITEM_CATGY_CODE = 'N'", "") & "))" _
        & " and ICTITEM1.ITEM_CODE = RSTSLOH1.ITEM_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & tempTable _
        & " (ITEM_CODE, COLLECTION_CODE, ITEM_PICTURE_FILENAME, DEPT_CODE, CUST_CODE, METAL_CLASS_CODE, MATL_CATGY_CODE, ITEM_CLASS_CODE, STYLE_CODE" _
        & ", SUWTD, SUMTD, SUSTD, HUTT) " & ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("BEGIN DECLARE CURSOR C1 IS" _
        & " SELECT STYLE_CODE, SUM (SUWTD) SUWTD, SUM (SUMTD) SUMTD, SUM (SUSTD) SUSTD, SUM (HUTT) HUTT, COUNT (*) ITEMS" _
        & "  FROM " & tempTable & " WHERE ITEM_CLASS_CODE = 'RING' GROUP BY STYLE_CODE;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " DELETE FROM " & tempTable _
        & " WHERE ITEM_CLASS_CODE = 'RING' AND STYLE_CODE = R1.STYLE_CODE" _
        & " AND ROWNUM < R1.ITEMS;" _
        & " UPDATE " & tempTable & " SET SUWTD = R1.SUWTD, SUMTD = R1.SUMTD, SUSTD = R1.SUSTD, HUTT = R1.HUTT " _
        & " WHERE ITEM_CLASS_CODE = 'RING' AND STYLE_CODE = R1.STYLE_CODE;" _
        & " END LOOP; END; END;")

        ASCDATA1.ExecuteSQL("Begin Declare R NUMBER; Cursor C1 is Select * from " & tempTable & " order by SUSTD Desc; " _
                            & " Begin for R1 in C1 Loop R := NVL(R,0) + 1; Update " & tempTable & " Set RANK = R where ITEM_CODE = R1.ITEM_CODE; END LOOP; END; END;")
        ASCDATA1.ExecuteSQL("Begin Declare Cursor C0 is Select Distinct DEPT_CODE, COLLECTION_CODE from " & tempTable & "; Begin For R0 in C0 Loop " _
                            & " Begin Declare R NUMBER; Cursor C1 is Select * from " & tempTable & " where DEPT_CODE = R0.DEPT_CODE and COLLECTION_CODE = R0.COLLECTION_CODE order by SUSTD Desc; " _
                            & " Begin for R1 in C1 Loop R := NVL(R,0) + 1; Update " & tempTable & " Set RANK_WITHIN_COLLECTION = R where ITEM_CODE = R1.ITEM_CODE; END LOOP; END; END; End Loop; End; End;")

        Dim CURRENT_WEEK As String() = Split(LookUp("GLTPARM3", RYW).Item("LEGEND"))

        grd.Text = "Top Sellers As Of Week " & CURRENT_WEEK(1).Substring(1, CURRENT_WEEK(1).Length - 2) ',  & IIf(chkTopNewOnly.Checked, " (New Only)", " (Overall)")
    End Sub

    Sub Setup_grdRSTSLOH2()
        grdRSTSLOH2.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdRSTSLOH2, "RANK_WITHIN_COLLECTION")
        grdRSTSLOH2.DisplayLayout.Bands(0).SortedColumns.Add("DEPT_CODE", False, True)
        grdRSTSLOH2.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)
        'grdRSTSLOH2.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CATGY_CODE", False, True)
        grdRSTSLOH2.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
        grdRSTSLOH2.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False, False)
        grdRSTSLOH2.Rows.ExpandAll(True)

        grdRSTSLOH2.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        'grdRSTSLOH2.DisplayLayout.Bands(0).ColumnFilters("ITEM_CATGY_CODE").FilterConditions.Add _
        '           (UltraWinGrid.FilterComparisionOperator.StartsWith, "N")
    End Sub
    Sub Setup_grdRSTSLOHO()
        grdRSTSLOHO.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdRSTSLOHO, "RANK_WITHIN_COLLECTION")
        grdRSTSLOHO.DisplayLayout.Bands(0).SortedColumns.Add("DEPT_CODE", False, True)
        grdRSTSLOHO.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)
        'grdRSTSLOH2.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CATGY_CODE", False, True)
        grdRSTSLOHO.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
        grdRSTSLOHO.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False, False)
        grdRSTSLOHO.Rows.ExpandAll(True)

        grdRSTSLOHO.DisplayLayout.Bands(0).ColumnFilters.ClearAllFilters()
        'grdRSTSLOH2.DisplayLayout.Bands(0).ColumnFilters("ITEM_CATGY_CODE").FilterConditions.Add _
        '           (UltraWinGrid.FilterComparisionOperator.StartsWith, "N")
    End Sub

    Sub Setup_Pictures()
        Dim IMAGE_FOLDER As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then IMAGE_FOLDER = "C:\Documents and Settings\wjz\Desktop\Clients\JHI\Images\"

        For Each rowRSTSLOHO As DataRow In dst.Tables("RSTSLOHO").Rows
            Dim ITEM_PICTURE_FILENAME As String = rowRSTSLOHO.Item("ITEM_PICTURE_FILENAME") & ""

            'testing
            If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then
                ITEM_PICTURE_FILENAME = "b.jpg"
            End If

            If ITEM_PICTURE_FILENAME <> "" Then
                Dim FILENAME As String = IMAGE_FOLDER & ITEM_PICTURE_FILENAME
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    rowRSTSLOHO.Item("ITEM_PICTURE") = ASCMAIN1.GetImageData(FILENAME) 'ASCMAIN1.Get_Image(IMAGE_FOLDER, ITEM_PICTURE_FILENAME)
                End If
            End If
        Next
    End Sub

    Private Sub tabSales_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabSales.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tabSales()
    End Sub

    Sub Setup_tabSales()
        UltraExplorerBar1.Groups("Options").Visible = ScreenMode And (tabSales.SelectedTab.Key = "Retail Sales && On Hand")
        UltraExplorerBar1.Groups("Customers").Visible = Not ScreenMode Or (tabSales.SelectedTab.Key = "Retail Sales && On Hand")
        UltraExplorerBar1.Groups("Top Sellers").Visible = ScreenMode And (tabSales.SelectedTab.Key = "Top Sellers")
    End Sub

    Public Overrides Function Excel_Export(ByVal grd As Infragistics.Win.UltraWinGrid.UltraGrid) As GemBox.Spreadsheet.ExcelFile
        If grd.Name <> "grdRSTSLOH2" Then

            Try
                Return MyBase.Excel_Export(grd)
            Catch ex As Exception
                MsgBox(ex.Message)
                Return Nothing
            End Try
        Else

            Dim X As GemBox.Spreadsheet.ExcelFile

            X = MyBase.Gembox_Export_to_Excel(grd, False)
            Dim MAX_COLUMN As Integer = X.Worksheets(0).CalculateMaxUsedColumns
            'correct heading
            With X.Worksheets(0)
                .Cells(0, 0).Value = ""
                .Cells(0, 1).Value = ""
                .Cells(1, 0).Style.Font.Color = Color.Black
                .Cells(1, 0).Style.Font.Size = 400
                .Cells(1, 0).Style.Font.Weight = 700
                .Cells(2, 0).Style.Font.Weight = 700
                .Cells(2, 0).Style.Font.Color = Color.Blue
                '.Cells(6, 0).Value = ""
                For Each cell As ExcelCell In .Rows(5).AllocatedCells
                    cell.Style.FillPattern.PatternBackgroundColor = Color.White
                    cell.Value = ""
                Next
                'correct body
                Dim lastrow As Boolean = False
                Dim count As Integer = 0
                Dim cellcheck As String
                Dim correction As String()
                Dim ROWS_TO_DELETE As Integer() = Nothing
                Dim i As Integer = 0

                For Each row As GemBox.Spreadsheet.ExcelRow In .Rows
                    'skip header
                    If row.Index < 5 Then
                        Continue For
                    End If
                    'delete extra summary rows
                    Select Case row.Cells(MAX_COLUMN - 1).Style.FillPattern.PatternForegroundColor.Name
                        Case "PaleGoldenrod"
                            ReDim Preserve ROWS_TO_DELETE(i)
                            ROWS_TO_DELETE(i) = row.Index
                            i += 1
                        Case "PaleGreen"
                            ReDim Preserve ROWS_TO_DELETE(i)
                            ROWS_TO_DELETE(i) = row.Index
                            i += 1
                        Case "LightGray"
                            ReDim Preserve ROWS_TO_DELETE(i)
                            ROWS_TO_DELETE(i) = row.Index
                            i += 1
                    End Select
                    cellcheck = row.Cells(0).Value & ""
                    'fix collection title
                    If cellcheck.StartsWith("Collection") Then
                        row.Cells(0).Style.Font.Weight = 700
                        correction = cellcheck.Split("(")
                        row.Cells(0).Value = correction(0)
                        correction = Nothing
                    ElseIf cellcheck.StartsWith("Exclusive") Then
                        row.Cells(0).Style.Font.Weight = 700
                        correction = cellcheck.Split("(")
                        row.Cells(0).Value = correction(0)
                        If correction(0).EndsWith(":  ") Then
                            row.Cells(0).Value = ""
                            Continue For
                        End If
                        correction = Nothing
                        row.Cells(0).Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, LineStyle.Medium)
                        row.Cells(MAX_COLUMN - 1).Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Medium)
                        For Each cell As ExcelCell In row.AllocatedCells
                            cell.Style.Borders.SetBorders(MultipleBorders.Top, Color.Black, LineStyle.Medium)
                            cell.Style.Borders.SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Medium)
                            cell.Style.FillPattern.SetSolid(Color.Gold)
                        Next
                        'fix dept title
                    ElseIf cellcheck.StartsWith("Dept") Then
                        row.Cells(0).Style.Font.Weight = 700
                        row.Cells(0).Style.Font.Size = 320
                        correction = cellcheck.Split("(")
                        row.Cells(0).Value = correction(0)
                        row.Cells(0).Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, LineStyle.Medium)
                        row.Cells(MAX_COLUMN - 1).Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Medium)
                        For Each cell As ExcelCell In row.AllocatedCells
                            cell.Style.Borders.SetBorders(MultipleBorders.Top, Color.Black, LineStyle.Medium)
                            cell.Style.Borders.SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Medium)
                            cell.Style.FillPattern.SetSolid(Color.Gold)
                        Next
                        correction = Nothing
                    End If
                    'Header Row
                    If row.Cells(1).Value = "Description" Then
                        For Each cell As ExcelCell In row.AllocatedCells
                            If cell.IsStyleDefault Then
                                Continue For
                            Else
                                cell.Style.Borders.SetBorders(MultipleBorders.Outside, Color.Black, LineStyle.Thin)
                                cell.Style.Font.Weight = 700
                            End If
                        Next
                    End If
                    'find summary row
                    If lastrow And row.Cells(1).Value = Nothing Then
                        'format summary row
                        row.Cells(0).Value = "Total Top Sellers"
                        row.Cells(0).Style.Font.Weight = 700
                        row.Cells(0).Style.Borders.SetBorders(MultipleBorders.Left, Color.Black, LineStyle.Medium)
                        row.Cells(MAX_COLUMN - 1).Style.Borders.SetBorders(MultipleBorders.Right, Color.Black, LineStyle.Medium)
                        For Each cell As ExcelCell In row.AllocatedCells
                            cell.Style.Borders.SetBorders(MultipleBorders.Top, Color.Black, LineStyle.Medium)
                            cell.Style.Borders.SetBorders(MultipleBorders.Bottom, Color.Black, LineStyle.Medium)
                            cell.Style.Font.Weight = 700
                        Next
                        lastrow = False
                    ElseIf row.Cells(1).Value <> Nothing Then
                        lastrow = True
                    Else
                        lastrow = False
                    End If
                    If row.Cells(2).Value <> Nothing AndAlso row.Cells(2).Value.ToString <> "Collection" Then
                        row.Height = 1100
                        row.Style.VerticalAlignment = VerticalAlignmentStyle.Center
                    End If
                Next
                X.Worksheets(0).Rows(X.Worksheets(0).Rows.Count - 1).Delete()
                For iter As Integer = ROWS_TO_DELETE.Length To 1 Step -1
                    X.Worksheets(0).Rows(ROWS_TO_DELETE(iter - 1)).Delete()
                Next
                X.Worksheets(0).Rows(5).Delete()
            End With
            'save Gembox formating
            Dim xlsFileName As String = "RSFSLOH1"
            Try
                X.SaveXls(ASCMAIN1.Folders("Work") & xlsFileName & ".xls")
                X = Nothing
            Catch ex As Exception
            End Try

            'INSERT PICTURES
            Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
            Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(ASCMAIN1.Folders("Work") & xlsFileName & ".xls")

            Dim FOLDERNAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
            If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then FOLDERNAME = "C:\Users\Ryan\Desktop\Images\"
            Dim rowICTITEM1 As DataRow
            Dim FILENAME As String
            Dim ITEM_CODE As String = ""
            Dim XWS As Microsoft.Office.Interop.Excel.Worksheet = XWB.Worksheets(1)
            Dim rng As Microsoft.Office.Interop.Excel.Range
            Dim shift As Microsoft.Office.Interop.Excel.XlInsertShiftDirection
            shift = Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown
            Dim copy As Microsoft.Office.Interop.Excel.XlInsertFormatOrigin
            copy = Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow
            XWS.Columns(1).insert(shift, copy)
            rng = XWS.Range("A:A")
            'rng.Insert(XlshiftToRight, xlFormatFromRightOrBelow)

            'rng.EntireColumn.AutoFormat(Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow)
            rng.EntireColumn.ColumnWidth = 14

            Try

                For i As Integer = 6 To XWS.UsedRange.Rows.Count Step +1
                    If XWS.Cells(i, 2).value.ToString & "" = "" Then
                        Continue For
                    ElseIf XWS.Cells(i, 2).value.ToString <> "" Then
                        ITEM_CODE = XWS.Cells(i, 2).value.ToString
                    Else
                        Continue For
                    End If
                    rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1 Is Nothing Then
                        Continue For
                    Else
                        FILENAME = FOLDERNAME & rowICTITEM1.Item("ITEM_PICTURE_FILENAME") & ""
                        If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then FILENAME = FOLDERNAME & ITEM_CODE & ".JPG"
                        rng = XWS.Range(XWS.Cells(i, 1), XWS.Cells(i, 1))
                        InsertPictureInRange(FILENAME, rng, XWS)
                    End If
                Next
            Catch ex As Exception
                Exit Try
            End Try

            XWB.Save()
            XWB = Nothing
            excel = Nothing

            Dim start_excel As New Process
            start_excel.StartInfo.Arguments = """" + xlsFileName + """ /e"
            start_excel.StartInfo.FileName = ASCMAIN1.Folders("Work") & xlsFileName & ".xls"
            start_excel.Start()
            Return Nothing
        End If

    End Function
    Sub InsertPictureInRange(ByVal PictureFileName As String, _
                     ByVal TargetCells As Microsoft.Office.Interop.Excel.Range, _
                     ByVal XWS As Microsoft.Office.Interop.Excel.Worksheet)

        ' inserts a picture and resizes it to fit the TargetCells range
        Dim pp As Microsoft.Office.Interop.Excel.Shape

        If TypeName(XWS) <> "Worksheet" Then Exit Sub
        If Dir(PictureFileName) = "" Then Exit Sub

        pp = XWS.Shapes.AddPicture(PictureFileName, _
           Microsoft.Office.Core.MsoTriState.msoFalse, _
           Microsoft.Office.Core.MsoTriState.msoCTrue, TargetCells.Left, TargetCells.Top, TargetCells.Width, TargetCells.Height)
        pp.Placement = Microsoft.Office.Interop.Excel.XlPlacement.xlMoveAndSize
        pp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse
        pp = Nothing
    End Sub

    Function XC( _
ByVal C As Int16, _
Optional ByVal R As Int16 = 0, _
Optional ByVal absolute As Boolean = False) As String

        Dim COL As String = ""
        If C >= 1 Then
            Dim B As Int16 = (C - 1) Mod 26 + 1
            Dim A As Int16 = (C - B) / 26
            COL = Chr(Asc("A") + B - 1)
            If A > 0 Then
                COL = Chr(Asc("A") + A - 1) & COL
            End If
            If absolute Then
                COL = "$" & COL
            End If

            If R = 0 Then
                COL = COL & ":" & COL
            ElseIf R > 0 Then
                COL = COL & IIf(absolute, "$", "") & CStr(R)
            End If
        End If

        Return COL
    End Function

End Class
Public Class srtComparerSATGRNR1
    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        Dim xCell As UltraWinGrid.UltraGridCell = DirectCast(x, UltraWinGrid.UltraGridCell)
        Dim yCell As UltraWinGrid.UltraGridCell = DirectCast(y, UltraWinGrid.UltraGridCell)

        Dim xv As String = xCell.Value & ""
        Dim yv As String = yCell.Value & ""

        Dim COLUMN_NAME As String = xCell.Column.Key
        'xv = Z_Fix(xv)
        'yv = Z_Fix(yv)

        If xv = yv Then
            Return 0
        Else
            If COLUMN_NAME = "DEPT_CODE" Then
                If xv = "WM" Then
                    Return -1
                ElseIf yv = "WM" Then
                    Return 1
                ElseIf xv = "MN" Then
                    Return -1
                ElseIf yv = "MN" Then
                    Return 1
                Else
                    Return IIf(xv < yv, -1, 1)
                End If
            End If
        End If

    End Function

    Function Z_Fix(ByVal v As String) As String
        If v = "WM" Then
            v = "A" & v
        ElseIf v = "MN" Then
            v = "B" & v
        End If

        Return v
    End Function
End Class