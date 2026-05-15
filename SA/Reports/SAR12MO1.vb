Public Class SAR12MO1
    Dim PTD1 As Int16 = 1
    Dim PTD2 As Int16 = 12

    Dim VLGRN As New ValueList
    Dim VLT As New ValueList
    Dim VLCM As New ValueList
    Dim VLPA As New ValueList

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 12, -1)

        grpShow.Visible = (ASCMAIN1.DBS_COMPANY = "CPU") ' OPTION PROVIDED FOR CPU
        SplitContainer5.Panel2Collapsed = False ' probably should have not put this splitter and grid in stds - probably should have implemented the whole thing here

        VLGRN.ValueListItems.Add(New ValueListItem With {.DataValue = "U", .DisplayText = "Units"})
        VLGRN.ValueListItems.Add(New ValueListItem With {.DataValue = "S", .DisplayText = "Sales"})
        VLGRN.ValueListItems.Add(New ValueListItem With {.DataValue = "C", .DisplayText = "Costs"})
        VLGRN.ValueListItems.Add(New ValueListItem With {.DataValue = "G", .DisplayText = "$GP"})
        VLGRN.ValueListItems.Add(New ValueListItem With {.DataValue = "P", .DisplayText = "GP%"})
        VLGRN.ValueListItems.Add(New ValueListItem With {.DataValue = "B", .DisplayText = "Budget"})

        VLT.ValueListItems.Add(New ValueListItem With {.DataValue = "U", .DisplayText = "Units"})
        VLT.ValueListItems.Add(New ValueListItem With {.DataValue = "S", .DisplayText = "Sales"})
        VLT.ValueListItems.Add(New ValueListItem With {.DataValue = "B", .DisplayText = "Budget"})
        VLT.ValueListItems.Add(New ValueListItem With {.DataValue = "H", .DisplayText = "OH EOM"})

        VLCM.ValueListItems.Add(New ValueListItem With {.DataValue = "S", .DisplayText = "Amount"})

        'VLPA.ValueListItems.Add(New ValueListItem With {.DataValue = "P", .DisplayText = "Plan"})
        VLPA.ValueListItems.Add(New ValueListItem With {.DataValue = "A", .DisplayText = "Actual"})

        optGRN.Value = "M"
        optGRN.Value = "G"

        SplitContainer5.SplitterDistance = UltraGroupBox2.Top + UltraGroupBox2.Height + 5
        chkCODES_ON_ALL_LINES.Visible = Not (ASCMAIN1.CLIENT = "INT")
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        PTD1 = utb1.Value
        PTD2 = utb2.Value

        For M As Integer = 1 To 12
            For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("COLUMN_NAME = 'M" & Format(M, "00") & "'")
                Dim YP As String = ASCMAIN1.Period_Calc(RYP, -12 + M)
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YP)
                rowASTDSQLS.Item("COLUMN_CAPTION") = Mid(LEGEND, 10, 6)
            Next
        Next


        Dim ICTITEM1 As String = "ICTITEM1"
        If chkHISTCAT.Checked Then
            ICTITEM1 = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP, True)
        End If

        Dim SAT12MO1 As String = ""

        Dim SAT12MO1_ACN_R As String = ""
        Dim SAT12MO1_ACN_W As String = ""

        If optACN.Value <> "A" Then ' Stores Included: All, Comp, Non-Comp
            Dim YP_MIN As String = ""
            Dim YP_MAX As String = ""
            For Each rowASTRECAP As DataRow In tblASTRECAP.Rows
                Dim GRN As String = rowASTRECAP.Item("GRN")
                If GRN = "T" Then
                    SAT12MO1_ACN_R = "*"
                Else
                    SAT12MO1_ACN_W = "*"
                End If
                Dim YEAR As String = rowASTRECAP.Item("YEAR")
                Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))
                Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
                Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)
                If YP_01 < YP_MIN Or YP_MIN = "" Then YP_MIN = YP_01
                If YP_12 > YP_MAX Or YP_MAX = "" Then YP_MAX = YP_12
            Next
            If YP_MAX > ASCMAIN1.CYP Then
                YP_MAX = ASCMAIN1.CYP
            End If

            Dim MOS As Int32 = ASCMAIN1.Period_Diff(YP_MIN, YP_MAX) + 1

            If SAT12MO1_ACN_R = "*" Then
                ASCMAIN1.sql = "" _
                & "Select CUST_CODE, CUST_STORE_NO" _
                & ", COUNT (DISTINCT OPS_YYYYPP) MOS" _
                & " from RSTRETL1 where AMT_SOLD > 0" _
                & " and OPS_YYYYPP between '" & YP_MIN & "' and '" & YP_MAX & "'" _
                & IIf(SQLA("CUST_CODE") <> "", " and CUST_CODE in (" & SQLA("CUST_CODE", , True) & ")", "") _
                & " group by CUST_CODE, CUST_STORE_NO"
                SAT12MO1_ACN_R = ASCMAIN1.Temp_Table

                ASCDATA1.ExecuteSQL("Delete from " & SAT12MO1_ACN_R & " where MOS <> " & CStr(MOS))
                ASCDATA1.ExecuteSQL("Alter Table " & SAT12MO1_ACN_R & " Add Primary Key (CUST_CODE, CUST_STORE_NO)")
            End If

            If SAT12MO1_ACN_W = "*" Then
                ASCMAIN1.sql = "" _
                & "Select CUST_CODE, CUST_STORE_NO" _
                & ", COUNT (DISTINCT ORDR_YYYYPP_UPDATED) MOS" _
                & " from SOTINVH2 where ORDR_QTY_SHIP > 0" _
                & " and ORDR_YYYYPP_UPDATED between '" & YP_MIN & "' and '" & YP_MAX & "'" _
                & IIf(SQLA("CUST_CODE") <> "", " and CUST_CODE in (" & SQLA("CUST_CODE", , True) & ")", "") _
                & " group by CUST_CODE, CUST_STORE_NO"
                SAT12MO1_ACN_W = ASCMAIN1.Temp_Table

                ASCDATA1.ExecuteSQL("Delete from " & SAT12MO1_ACN_W & " where MOS <> " & CStr(MOS))
                ASCDATA1.ExecuteSQL("Alter Table " & SAT12MO1_ACN_W & " Add Primary Key (CUST_CODE, CUST_STORE_NO)")
            End If

        End If

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")

        Dim SOURCE_TABLE_NAME As String = ""
        Dim by_Item As Boolean = False
        If COLUMN_NAMEs.Contains("ITEM_CODE") Then ' This needs to include all columns dependendent upon Item Code
            by_Item = True
        End If
        Dim by_Store As Boolean = False
        If COLUMN_NAMEs.Contains("CUST_STORE_NO") Then ' This needs to include all columns dependendent upon Store
            by_Store = True
        End If

        ' until the above are resolved, assume we need the most detail
        by_Item = True
        by_Store = True


        ' Wholesale Shipments

        If by_Item And by_Store Then
            SOURCE_TABLE_NAME = "SOTINVH2"
        Else
            If Not by_Item And Not by_Store Then
                SOURCE_TABLE_NAME = "SATSSUM0"
            Else
                If by_Item Then
                    SOURCE_TABLE_NAME = "SATSSUMI"
                Else
                    SOURCE_TABLE_NAME = "SATSSUMS"
                End If
            End If
        End If

        If chkHISTCAT.Checked Then
            sql_TABLE_NAMEs = Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1 & " ICTITEM1")
        End If

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            Dim YP As String = ""
            Dim COLUMN_NAME As String = ""
            Dim sql_GRN As String = ""

            If (GRN = "G" Or GRN = "R" Or GRN = "N" Or GRN = "T") And (USC = "U" Or USC = "S" Or USC = "C" Or USC = "G" Or USC = "H") Then

                If GRN = "T" Then
                    sql_GRN = ""
                    YP = ""
                    COLUMN_NAME = ""
                    YP = "SAT12MO1.OPS_YYYYPP"
                    If USC = "U" Then
                        COLUMN_NAME = "NVL(SAT12MO1.QTY_SOLD,0)"
                    ElseIf USC = "S" Then
                        COLUMN_NAME = "NVL(SAT12MO1.AMT_SOLD,0)"
                    ElseIf USC = "C" Then
                        COLUMN_NAME = "0"
                    ElseIf USC = "G" Then
                        COLUMN_NAME = "0"
                    ElseIf USC = "H" Then
                        COLUMN_NAME = "NVL(SAT12MO1.QTY_EOW,0)"
                    End If
                Else

                    sql_GRN = ""
                    If GRN = "G" Then
                        sql_GRN = " AND SAT12MO1.INV_TYPE = 'I'"
                    ElseIf GRN = "R" Then
                        sql_GRN = " AND SAT12MO1.INV_TYPE = 'C'"
                    End If

                    YP = ""
                    COLUMN_NAME = ""
                    Select Case SOURCE_TABLE_NAME
                        Case "SOTINVH2"
                            YP = "SAT12MO1.ORDR_YYYYPP_UPDATED"
                            If USC = "U" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0)"
                            ElseIf USC = "S" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0) * NVL(SAT12MO1.ORDR_UNIT_PRICE,0)"
                            ElseIf USC = "C" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0) * DECODE(SAT12MO1.WHSE_CODE,NULL,0,NVL(SAT12MO1.ITEM_UNIT_COST,0))"
                            ElseIf USC = "G" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0) * (NVL(SAT12MO1.ORDR_UNIT_PRICE,0) - DECODE(SAT12MO1.WHSE_CODE,NULL,0,NVL(SAT12MO1.ITEM_UNIT_COST,0)))"
                            End If
                        Case Else
                            YP = "SAT12MO1.OPS_YYYYPP"
                            If USC = "U" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_QTY_SHIP,0)"
                            ElseIf USC = "S" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_AMT_SHIP,0)"
                            ElseIf USC = "C" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_CGS_SHIP,0)"
                            ElseIf USC = "G" Then
                                COLUMN_NAME = "NVL(SAT12MO1.ORDR_AMT_SHIP,0) - NVL(SAT12MO1.ORDR_CGS_SHIP,0)"
                            End If
                    End Select
                End If

                sql_filter = " and " & YP & " BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'" & vbCrLf _
                & " and " & COLUMN_NAME & " <> 0"
                If chkACTIVE_ONLY.Checked Then
                    sql_filter &= vbCrLf & " and NVL(ICTITEM1.ITEM_STATUS,'?') = 'A'"
                End If

                If optACN.Value <> "A" And IIf(GRN = "T", SAT12MO1_ACN_R, SAT12MO1_ACN_W) <> "" Then
                    sql_filter &= vbCrLf & " and SAT12MO1_ACN.CUST_CODE (+) = SAT12MO1.CUST_CODE"
                    sql_filter &= vbCrLf & " and SAT12MO1_ACN.CUST_STORE_NO (+) = SAT12MO1.CUST_STORE_NO"
                    sql_filter &= vbCrLf & " and NVL(SAT12MO1_ACN.MOS,0) " & IIf(optACN.Value = "C", "<>0", "=0")
                End If


                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "'" & sql_GRN & " THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from " & IIf(GRN = "T", IIf(USC = "H", "GLTPARM3, RSTRETL1", "RSTRETL1"), SOURCE_TABLE_NAME) & " SAT12MO1 " & sql_TABLE_NAMEs & vbCrLf _
                & IIf(optACN.Value <> "A" And IIf(GRN = "T", SAT12MO1_ACN_R, SAT12MO1_ACN_W) <> "", "," & IIf(GRN = "T", SAT12MO1_ACN_R, SAT12MO1_ACN_W) & " SAT12MO1_ACN", "") & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & IIf(GRN & USC = "TH", " and GLTPARM3.YYYYWW = SAT12MO1.OPS_YYYYWW and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK", "")) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If

        Next


        ' Promo Expense

        MyBase.Get_SQL("C")

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            Dim YP As String = "SPTCOOP1.OPS_YYYYPP"

            Dim COLUMN_NAME As String = "NVL(SPTCOOP3.DIST_AMT,0)"

            If GRN = "C" Then

                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from SPTCOOP1 " & sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If
        Next


        ' Model Expense 

        MyBase.Get_SQL("M")

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            Dim YP As String = "SPTCWRX2.OPS_YYYYPP"

            Dim COLUMN_NAME As String = "NVL(SPTCWRX2.BILL_AMT,0)"

            If GRN = "M" Then

                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from SPTCWRX2 " & sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If
        Next


        ' Wholesale Shipment Budgets

        MyBase.Get_SQL("B")

        If ASCMAIN1.CLIENT = "AHA" Then

            For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

                Dim USC As String = rowASTRECAP.Item("USC")
                Dim GRN As String = rowASTRECAP.Item("GRN")
                Dim YEAR As String = rowASTRECAP.Item("YEAR")
                Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

                Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
                Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

                Dim COLUMN_NAME As String = ""
                Dim sql_GRN As String = ""

                If (GRN = "G" Or GRN = "N") And USC = "B" Then

                    sql_filter = " and SATBUDG1.OPS_YYYYPP BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'"

                    Dim sql_Data As String = ""
                    For M As Integer = 1 To 12
                        Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                        sql_Data &= ", Sum (CASE WHEN SATBUDG1.OPS_YYYYPP = '" & XYP & "'" & sql_GRN & " THEN NVL(SATBUDG1.BUDGET,0) ELSE 0 END) M" & Format(M, "00") & vbCrLf
                    Next

                    sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                        & " from SATBUDG1" & sql_TABLE_NAMEs & vbCrLf _
                        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                        & " group by " & sql_GROUP_BY_cols

                    ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                        & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                        & COLUMN_NAMEs_appended _
                        & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                        & "(" & sql & ")"
                    ASCDATA1.ExecuteSQL()

                End If
            Next

        Else

            Dim SATBUDW1 As String = TAC.SOCMAIN1.Setup_Budgets_by_Customer

            For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

                Dim USC As String = rowASTRECAP.Item("USC")
                Dim GRN As String = rowASTRECAP.Item("GRN")
                Dim YEAR As String = rowASTRECAP.Item("YEAR")
                Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

                Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
                Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

                Dim COLUMN_NAME As String = ""
                Dim sql_GRN As String = ""

                If (GRN = "G" Or GRN = "N") And USC = "B" Then

                    sql_filter = " and SATBUDW1.OPS_YYYY BETWEEN '" & Mid(YP_01, 1, 4) & "' AND '" & Mid(YP_12, 1, 4) & "'"

                    Dim sql_Data As String = ""
                    For M As Integer = 1 To 12
                        Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                        sql_Data &= ", Sum (CASE WHEN SATBUDW1.OPS_YYYY = '" & Mid(XYP, 1, 4) & "'" & sql_GRN & " THEN NVL(SATBUDW1.WB_P" & Mid(XYP, 5, 2) & ",0) ELSE 0 END) M" & Format(M, "00") & vbCrLf
                    Next

                    sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                        & " from " & SATBUDW1 & " SATBUDW1 " & sql_TABLE_NAMEs & vbCrLf _
                        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                        & " group by " & sql_GROUP_BY_cols

                    ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                        & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                        & COLUMN_NAMEs_appended _
                        & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                        & "(" & sql & ")"
                    ASCDATA1.ExecuteSQL()

                End If
            Next

        End If


        ' Retail Budgets

        MyBase.Get_SQL("T")

        ASCMAIN1.sql = "Select * from ICTITEM1 where ROWNUM < 1"
        Dim ICTITEM1_BUDGETS As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1_BUDGETS & " Add Primary Key (ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ICTITEM1_BUDGETS & "_1 on " & ICTITEM1_BUDGETS & " (COLLECTION_CODE, ITEM_CATGY_CODE)")

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1_BUDGETS)

        ASCMAIN1.sql = "Insert into " & ICTITEM1_BUDGETS _
        & " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE) " _
        & " Select ROWNUM, 'Item for ' || COLLECTION_CODE || '-' || ITEM_CATGY_CODE" _
        & ", COLLECTION_CODE, ITEM_CATGY_CODE " _
        & " from " _
        & " (Select Distinct COLLECTION_CODE, ITEM_CATGY_CODE from SATBUDD1)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" _
        & ", OPS_YYYYPP, AMT_SOLD BUDGET " _
        & " from RSTRETL1 where ROWNUM < 1"
        Dim SATBUDDI As String = ASCMAIN1.Temp_Table
        For P As Integer = 1 To 12
            ASCMAIN1.sql = "Insert into " & SATBUDDI _
            & " Select SATBUDD1.CUST_CODE, SATBUDD1.CUST_STORE_NO, ICTITEM1.ITEM_CODE" _
            & ", OPS_YYYY || '" & Format(P, "00") & "'" _
            & ", SATBUDD1.BUDGET_P" & Format(P, "00") & " BUDGET" _
            & " from SATBUDD1," & ICTITEM1_BUDGETS & " ICTITEM1" _
            & " where SATBUDD1.BUDGET_P" & Format(P, "00") & " <> 0" _
            & "   and ICTITEM1.COLLECTION_CODE = SATBUDD1.COLLECTION_CODE" _
            & "   and ICTITEM1.ITEM_CATGY_CODE = SATBUDD1.ITEM_CATGY_CODE"
            ASCDATA1.ExecuteSQL()
        Next

        Dim PA As Int32 = 1 - ASCMAIN1.PCO ' -1 * ((12 + ASCMAIN1.PCO) Mod 12) + 1

        ASCMAIN1.sql = "Update " & SATBUDDI & " Set OPS_YYYYPP = PERIOD_CALC(OPS_YYYYPP," & CStr(PA) & ")"
        ASCDATA1.ExecuteSQL()

        'SOURCE_TABLE_NAME = SATBUDDI

        ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1_BUDGETS)

        ASCMAIN1.sql = "Insert into " & ICTITEM1_BUDGETS _
        & " (ITEM_CODE, ITEM_DESC, COLLECTION_CODE, ITEM_CATGY_CODE) " _
        & " Select ROWNUM, 'Item for ' || COLLECTION_CODE || '-' || ITEM_CATGY_CODE" _
        & ", COLLECTION_CODE, ITEM_CATGY_CODE " _
        & " from " _
        & " (Select Distinct COLLECTION_CODE, ITEM_CATGY_CODE from RSTBUDR1)"
        ASCDATA1.ExecuteSQL()


        Dim RSTBUDRI As String = TAC.RSCMAIN1.RSTBUDR1_as_YP
        ASCDATA1.ExecuteSQL("Alter Table " & RSTBUDRI & " Add ITEM_CODE VARCHAR2(30)")
        ASCDATA1.ExecuteSQL("Update " & RSTBUDRI & " R SET ITEM_CODE = (Select ITEM_CODE from " & ICTITEM1_BUDGETS & " where COLLECTION_CODE = R.COLLECTION_CODE and ITEM_CATGY_CODE = R.ITEM_CATGY_CODE)")

        SOURCE_TABLE_NAME = RSTBUDRI

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            Dim YP As String = ""
            Dim COLUMN_NAME As String = ""
            Dim sql_GRN As String = ""

            If GRN = "T" And USC = "B" Then

                If GRN = "T" Then
                    sql_GRN = ""
                    YP = "SAT12MO1.OPS_YYYYPP"
                    COLUMN_NAME = "NVL(SAT12MO1.BUDGET,0)"
                End If

                sql_filter = " and " & YP & " BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'"

                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    sql_Data &= ", Sum (CASE WHEN  " & YP & " = '" & XYP & "'" & sql_GRN & " THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from " & SOURCE_TABLE_NAME & " SAT12MO1 " & Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1_BUDGETS & " ICTITEM1") & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ' MAYBE WE SHOULD HAVE OUR OWN DATA SOURCE FOR T (SELLTHRU) BUDGETS? - 
                sql = Replace(sql, "SATBUDW1", "SAT12MO1")

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If
        Next



        ' Freelance

        MyBase.Get_SQL("F")

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim USC As String = rowASTRECAP.Item("USC")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim Y As Integer = IIf(YEAR = "N", -1, Val(YEAR))

            Dim YP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y - 12 + 1)
            Dim YP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

            'Dim YP As String = "SPTMXWSX.YYYYMM"
            sql_filter = " and SPTMXWSX.YYYYMM BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'"
            sql_filter = " and SPTMXWSX.OPS_YYYYPP BETWEEN '" & YP_01 & "' AND '" & YP_12 & "'"
            sql_filter &= " and NVL(SPTMXWSX.ACT,0) <> 0"

            Dim COLUMN_NAME As String = "NVL(SPTMXWSX.ACT,0)"

            Dim sqlTable As String = "" _
                & "Select SPTMXWS2.CUST_CODE, SPTMXWS2.CUST_STORE_NO, SPTCWRXC.COLLECTION_CODE" & vbCrLf _
                & ", ICTSEAS1.SEASON_YEAR || TRIM(TO_CHAR(SPTMXWS2.MONTH_NO+DECODE(ICTSEAS1.SEASON_TYPE,'F',6,0),'00')) YYYYMM" & vbCrLf _
                & ", SUM (SPTMXWS2.TY_SPEND) ACT" & vbCrLf _
                & " from SPTMXWS2,SPTCWRXC,ICTSEAS1" & vbCrLf _
                & " where SPTCWRXC.CHECKBOOK (+) = SPTMXWS2.CHECKBOOK" & vbCrLf _
                & "   and ICTSEAS1.SEASON_CODE = SPTMXWS2.SEASON_CODE" & vbCrLf _
                & "   and SPTMXWS2.TY_SPEND <> 0" & vbCrLf _
                & "   and ICTSEAS1.SEASON_YEAR || TRIM(TO_CHAR(SPTMXWS2.MONTH_NO+DECODE(ICTSEAS1.SEASON_TYPE,'F',6,0),'00')) >= '" & YP_01 & "'" _
                & "   and ICTSEAS1.SEASON_YEAR || TRIM(TO_CHAR(SPTMXWS2.MONTH_NO+DECODE(ICTSEAS1.SEASON_TYPE,'F',6,0),'00')) <= '" & YP_12 & "'" _
                & " group by SPTMXWS2.CUST_CODE, SPTMXWS2.CUST_STORE_NO, SPTCWRXC.COLLECTION_CODE" & vbCrLf _
                & ", ICTSEAS1.SEASON_YEAR || TRIM(TO_CHAR(SPTMXWS2.MONTH_NO+DECODE(ICTSEAS1.SEASON_TYPE,'F',6,0),'00'))"

            sqlTable = "SPTCWRX5"

            If GRN = "F" Then

                Dim sql_Data As String = ""
                For M As Integer = 1 To 12
                    Dim XYP As String = ASCMAIN1.Period_Calc(YP_01, M - 1)
                    '                    sql_Data &= ", Sum (CASE WHEN  SPTMXWSX.YYYYMM = '" & XYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                    sql_Data &= ", Sum (CASE WHEN  SPTMXWSX.OPS_YYYYPP = '" & XYP & "' THEN " & COLUMN_NAME & " ELSE 0 END) M" & Format(M, "00") & vbCrLf
                Next

                sql = "Select " & sql_SELECT_cols & ", " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf & sql_Data _
                & " from (" & sqlTable & ") SPTMXWSX " & sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & ",ASTSRPT1_RECAP_ROW_NO" _
                & COLUMN_NAMEs_appended _
                & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12)" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()

            End If
        Next



        ' Calculated Lines 

        For Each rowASTRECAP As DataRow In tblASTRECAP.Rows

            Dim GRN As String = rowASTRECAP.Item("GRN")

            If GRN = "*" Then
                Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                Dim sqlRTi As String = "  R0 " & ASTSRPT1 & "%ROWTYPE;"
                Dim sqlRT As String = ""
                Dim MAX_LNO As Integer = Val(tblASTRECAP.Compute("MAX(ASTSRPT1_RECAP_ROW_NO)", "ASTSRPT1_RECAP_ROW_NO <> " & CStr(ASTSRPT1_RECAP_ROW_NO)) & "")
                For I As Integer = 1 To MAX_LNO
                    sqlRT &= Replace(sqlRTi, "R0", "R" & CStr(I)) & vbCrLf
                Next
                Dim sqlMs As String = ""
                For i As Integer = 1 To 12
                    sqlMs &= "M" & Format(i, "00") & " NUMBER;" & vbCrLf
                Next

                Dim sqlR0s As String = ""
                For i As Integer = 1 To 12
                    sqlR0s &= "R0.M" & Format(i, "00") & " :=0; "
                Next
                sqlR0s &= vbCrLf

                Dim sqlSeli As String = "   Select * into R0 from " & ASTSRPT1 _
                                        & " where G1=R.G1 and G2=R.G2 and G3=R.G3 and G4=R.G4 and G5=R.G5 and G6=R.G6 and G7=R.G7 and G8=R.G8 and G9=R.G9"
                Dim sqlSel As String = ""

                For I As Integer = 1 To MAX_LNO
                    sqlSel &= Replace(sqlR0s, "R0", "R" & CStr(I))
                    sqlSel &= "   Begin" & vbCrLf & Replace(sqlSeli, "R0", "R" & CStr(I)) & " and ASTSRPT1_RECAP_ROW_NO = " & CStr(I) & ";" & vbCrLf & " Exception" & vbCrLf & " When Others then Null;" & vbCrLf & " End;" & vbCrLf
                Next

                Dim CALC As String = rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CALC") & "" ' "R1-R2"

                For I As Integer = 19 To 1 Step -1
                    Dim C As String = "R" & CStr(I)
                    CALC = Replace(CALC, C, "NVL(" & C.ToLower & ".M00" & ",0)")
                Next

                Dim sqlUpd As String = ""
                Dim sqlInsi As String = ""
                Dim sqlIns As String = ""
                For i As Integer = 1 To 12
                    Dim C As String = "M" & Format(i, "00")
                    sqlIns &= ", " & Replace(CALC, "M00", C)
                    sqlUpd &= "Begin " & C & " := " & Replace(CALC, "M00", C) & "; Exception When Others then Null; End;" & vbCrLf
                Next
                sqlIns &= ", 0, 0, 0, 0"

                ASCMAIN1.sql = "" _
                    & "Begin " & vbCrLf _
                    & " Declare " & vbCrLf _
                    & "  Cursor C is Select Distinct G1,G2,G3,G4,G5,G6,G7,G8,G9 from " & ASTSRPT1 & ";" & vbCrLf _
                    & sqlRT _
                    & sqlMs _
                    & " Begin" & vbCrLf _
                    & "  For R in C Loop" & vbCrLf _
                    & sqlSel _
                    & "   Begin" & vbCrLf _
                    & "    Insert into " & ASTSRPT1 & vbCrLf _
                    & "     Values (R.G1,R.G2,R.G3,R.G4,R.G5,R.G6,R.G7,R.G8,R.G9," & CStr(ASTSRPT1_RECAP_ROW_NO) & vbCrLf _
                    & sqlIns & ");" & vbCrLf _
                    & "   Exception" & vbCrLf _
                    & "    When Others Then " & vbCrLf _
                    & "     M01:=0;M02:=0;M03:=0;M04:=0;M05:=0;M06:=0;M07:=0;M08:=0;M09:=0;M10:=0;M11:=0;M12:=0;" & vbCrLf _
                    & sqlUpd _
                    & "     Insert into " & ASTSRPT1 & vbCrLf _
                    & "      Values (R.G1,R.G2,R.G3,R.G4,R.G5,R.G6,R.G7,R.G8,R.G9," & CStr(ASTSRPT1_RECAP_ROW_NO) & vbCrLf _
                    & ",M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,0,0,0,0);" & vbCrLf _
                    & "   End;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()
            End If
        Next

        '"(G1,G2,G3,G4,G5,G6,G7,G8,G9,ASTSRPT1_RECAP_ROW_NO)" & 


        Dim YTDcols As String = ""
        Dim TOTALcols As String = ""
        For M As Integer = 1 To 12
            If M >= PTD1 And M <= PTD2 Then
                YTDcols &= "+M" & Format(M, "00")
            End If
            TOTALcols &= "+M" & Format(M, "00")
        Next
        ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set YTD = " & YTDcols & ", TOTAL = " & TOTALcols
        ASCDATA1.ExecuteSQL()

    End Sub

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)

        If Absx1.chkFor("THOUSANDS").Checked Then
            Dim sql As String = ""
            For Each COLUMN_NAME As String In New String() _
            {"M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12", "YTD", "TOTAL"}
                sql &= ", " & COLUMN_NAME & " = " & COLUMN_NAME & " / 1000"
            Next

            ASCMAIN1.sql = "Update " & TT & " Set " & Mid(sql, 2)
            ASCDATA1.ExecuteSQL()
        End If

        For Each rowASTRECAP As DataRow In tblASTRECAP.Select("USC = 'P'")
            Dim GRN As String = rowASTRECAP.Item("GRN")
            Dim YEAR As String = rowASTRECAP.Item("YEAR")
            Dim rowS() As DataRow = tblASTRECAP.Select("USC = 'S' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
            Dim rowG() As DataRow = tblASTRECAP.Select("USC = 'G' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
            'Dim rowC() As DataRow = tblASTRECAP.Select("USC = 'C' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
            ASCMAIN1.sql = "Delete from " & TT & " where ASTSRPT1_RECAP_ROW_NO = " & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO")
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "INSERT INTO " & TT & " Select S.G1,S.G2,S.G3,S.G4,S.G5,S.G6,S.G7,S.G8,S.G9," & rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & " ASTSRPT1_RECAP_ROW_NO" & vbCrLf
            For Each COLUMN_NAME As String In New String() _
            {"M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12", "YTD", "TOTAL"}
                ASCMAIN1.sql &= Replace(", TRUNC(10000*DECODE(NVL(S.M01,0),0,0,NVL(G.M01,0)/NVL(S.M01,0)))/100 M01" & vbCrLf, "M01", COLUMN_NAME)
            Next
            ASCMAIN1.sql &= ",null, null"
            ASCMAIN1.sql &= " FROM " & vbCrLf _
            & "(SELECT * FROM " & TT & " WHERE ASTSRPT1_RECAP_ROW_NO = " & rowS(0).Item("ASTSRPT1_RECAP_ROW_NO") & ") S," & vbCrLf _
            & "(SELECT * FROM " & TT & " WHERE ASTSRPT1_RECAP_ROW_NO = " & rowG(0).Item("ASTSRPT1_RECAP_ROW_NO") & ") G" & vbCrLf _
            & " WHERE S.G1 = G.G1 AND S.G2 = G.G2 AND S.G3 = G.G3 AND S.G4 = G.G4 AND S.G5 = G.G5" & vbCrLf _
            & "   AND S.G6 = G.G6 AND S.G7 = G.G7 AND S.G8 = G.G8 AND S.G9 = G.G9"

            ASCDATA1.ExecuteSQL()
        Next
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        Dim rowASTGROUP As DataRow = dst.Tables("ASTGROUP").Rows.Find(aRC)
        rowASTGROUP.Item("GROUP_CODE") = "Recap"
        rowASTGROUP.Item("GROUP_DESC") = "All Above"

        Dim P As New List(Of Int16)
        For Each ROWASTRECAP As DataRow In dst.Tables("ASTRECAP").Rows
            If ROWASTRECAP.Item("TYPE") & "" = "P" Then
                P.Add(ROWASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO"))
            End If
        Next

        Dim J As Int16 = 0
        Dim rowT() As DataRow = Nothing

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "G1,G2,G3,G4,G5,G6,G7,G8,G9")
            If P.Contains(Val(rowASTSRPT1.Item("ASTSRPT1_RECAP_ROW_NO"))) Then
                rowASTSRPT1.Item("PPTT") = DBNull.Value
                rowASTSRPT1.Item("TPTT") = DBNull.Value
            Else
                For I As Int16 = 1 To 9
                    If rowASTSRPT1.Item("G" & CStr(I)) & "" = "x" Or rowASTSRPT1.Item("G" & CStr(I)) & "" = aRC Then
                        J = I - 1
                        Exit For
                    End If
                Next
                Dim SQL As String = ""
                If J > 1 Then
                    For I As Int16 = 1 To J - 1
                        SQL &= " and G" & CStr(I) & " = '" & rowASTSRPT1.Item("G" & CStr(I)) & "'"
                    Next
                End If
                SQL = "G" & CStr(J) & " = '" & aRC & "' and ASTSRPT1_RECAP_ROW_NO = " & rowASTSRPT1.Item("ASTSRPT1_RECAP_ROW_NO") & SQL

                If J = 0 Then
                    rowASTSRPT1.Item("PPTT") = DBNull.Value
                    rowASTSRPT1.Item("TPTT") = DBNull.Value
                Else
                    rowT = dst.Tables("ASTSRPT1").Select(SQL)

                    If rowT IsNot Nothing AndAlso rowT.Length = 1 Then
                        Dim YTD As Decimal = Val(rowASTSRPT1.Item("YTD") & "")
                        Dim YTD_T As Decimal = Val(rowT(0).Item("YTD") & "")
                        Dim PPTT As Decimal = 0
                        If YTD_T <> 0 Then PPTT = 100 * YTD / YTD_T
                        rowASTSRPT1.Item("PPTT") = PPTT
                        Dim TOTAL As Decimal = Val(rowASTSRPT1.Item("TOTAL") & "")
                        Dim TOTAL_T As Decimal = Val(rowT(0).Item("TOTAL") & "")
                        Dim TPTT As Decimal = 0
                        If TOTAL_T <> 0 Then TPTT = 100 * TOTAL / TOTAL_T
                        rowASTSRPT1.Item("TPTT") = TPTT
                    End If
                End If

            End If
        Next
    End Sub

    Public Overrides Sub Print_Report()

        'Dim SUBT As String = ""
        If optACN.Value = "C" Then
            SUBT &= " -Incl Comp Stores Only"
        ElseIf optACN.Value = "N" Then
            SUBT &= " -Incl Non-Comp Stores Only"
        End If
        If chkACTIVE_ONLY.Checked Then
            SUBT &= " -Active Items Only"
        End If

        CR_params.Add("RYPLEGEND", RYPLEGEND)

        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
        CR_params.Add("HIDE", "0")
        For M As Int32 = 1 To 12
            Dim MONTH_DESC As String = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -1 * (12 - M)), False, True), 1, 6)
            CR_params.Add("MD" & Format(M, "00"), MONTH_DESC)
            If M = PTD1 Then
                CR_params.Add("T1", Format(M, "00"))
                If Absx1.chkFor("FISCALYTD").Checked Then
                    CR_params.Add("MD1", "Fiscal YTD")
                Else
                    CR_params.Add("MD1", MONTH_DESC)
                End If
            End If
            If M = PTD2 Then
                CR_params.Add("T2", Format(M, "00"))
                If Absx1.chkFor("FISCALYTD").Checked Then
                    CR_params.Add("MD2", Mid(RYP, 1, 4))
                Else
                    CR_params.Add("MD2", MONTH_DESC)
                End If
            End If
        Next

        If dst.Tables("ASTRECAP").Select("GRN='*'").Length = 0 Then
            Generate_Report(RPT, , SUBT)
        End If
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Report Period"
            End If

            For Each rowASTRECAP As DataRow In tblASTRECAP.Select("USC = 'P'")
                Dim GRN As String = rowASTRECAP.Item("GRN")
                Dim YEAR As String = rowASTRECAP.Item("YEAR")
                Dim rowS() As DataRow = tblASTRECAP.Select("USC = 'S' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
                Dim rowG() As DataRow = tblASTRECAP.Select("USC = 'G' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
                'Dim rowC() As DataRow = tblASTRECAP.Select("USC = 'C' and GRN = '" & GRN & "' AND YEAR = '" & YEAR & "'")
                If rowS.Length <> 1 Or rowG.Length <> 1 Then ' Or rowC.Length <> 1 Then
                    EMsg &= "You Must Select 1 line for Sales and 1 line for GP in order to have GP% for Sales Type:" & GRN & ", Year: " & YEAR
                End If
            Next

            If tblASTRECAP.Select("GRN='*'").Length <> 0 Then
                Dim tblTest As New DataTable
                With tblTest
                    For I As Integer = 19 To 1 Step -1
                        .Columns.Add("R" & CStr(I), GetType(System.Decimal))
                    Next
                End With
                For Each rowASTRECAP As DataRow In tblASTRECAP.Select("GRN='*'")
                    Dim ASTSRPT1_RECAP_ROW_CALC As String = rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CALC")
                    rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CALC") = ASTSRPT1_RECAP_ROW_CALC.ToUpper
                    Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                    Try
                        tblTest.Columns.Add("C" & CStr(ASTSRPT1_RECAP_ROW_NO), GetType(System.Decimal), ASTSRPT1_RECAP_ROW_CALC)
                    Catch ex As Exception
                        EMsg &= vbCr & "Problem with Calculation on Line " & CStr(ASTSRPT1_RECAP_ROW_NO) & vbCr & ex.Message
                    End Try
                Next
            End If

            If optSL.Value & "" = "P" Then
                If ASCMAIN1.CLIENT = "INT" Then
                    EMsg &= vbCr & "Option not available"
                Else
                    If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length < 2 Then
                        EMsg &= vbCr & "Pivot option must have 2 or more Sort Fields"
                    End If
                End If
            End If
        End If

        EMsg &= TAC.TACMAIN1.Check_Permissions(Me) ' for FS
        If Trim(ASCMAIN1.USER_CODES) = "FS" Then
            For Each rowASTRECAP As DataRow In tblASTRECAP.Select("USC = 'C' OR USC = 'G' OR USC = 'P'")
                EMsg &= "You May NOT select Costs or GP data"
            Next
        End If
    End Sub

    Private Sub utb1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles utb1.ValueChanged
        lblutb1.Text = Format(utb1.Value, "00")
    End Sub

    Private Sub utb2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles utb2.ValueChanged
        lblutb2.Text = Format(utb2.Value, "00")
    End Sub

    Private Sub UltraCheckEditor2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UltraCheckEditor2.CheckedChanged

        utb1.Enabled = Not Absx1.chkFor("FISCALYTD").Checked
        utb2.Enabled = Not Absx1.chkFor("FISCALYTD").Checked
        If Absx1.chkFor("FISCALYTD").Checked Then
            Set_utb()
        End If

    End Sub

    Sub Set_utb()
        Dim YP As String = Absx1.cmbFor("RYP").Value
        Dim P As Integer = Val(Mid(YP, 6, 2))
        utb1.Value = 12 - P + 1
        utb2.Value = 12
    End Sub

    Private Sub UltraCombo1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles UltraCombo1.ValueChanged
        If Absx1.chkFor("FISCALYTD").Checked Then
            Set_utb()
        End If
    End Sub

    Function CALCx(CALC As String, R As Integer, XC As Integer, ASTSRPT1_RECAP_ROW_NO As Integer, opt As String) As String

        ' opt = S for Stacked
        ' opt = L for 1-line
        ' opt = P for Pivot

        CALCx = CALC

        For I As Integer = 19 To 1 Step -1
            CALCx = Replace(CALCx, "R" & CStr(I), "~" & CStr(I))
        Next

        For I As Integer = 19 To 1 Step -1
            If opt = "S" Or opt = "P" Then
                CALCx = Replace(CALCx, "~" & CStr(I), Excel_Cell0(R + I - ASTSRPT1_RECAP_ROW_NO, XC))

            Else
                CALCx = Replace(CALCx, "~" & CStr(I), Excel_Cell0(R, XC + 17 * (I - ASTSRPT1_RECAP_ROW_NO)))

            End If
        Next

        'If DATA_TYPEs(MS(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
        '    'CALCx &= "/100"
        '    'CALCx = "(" & CALC & ")*100"
        'End If

        Return "=IFERROR(" & CALCx & ",0)"
    End Function
    Public Overrides Sub Post_Process_Special()
        MyBase.Post_Process_Special()

        'If ASCMAIN1.CLIENT = "INT" Then
        '    Create_XLS_INT()
        'Else
        Create_XLS()
        'End If

    End Sub

    Sub Create_XLS_INT()
        Try
            'Prepare_XLS()

            Dim opt As String = optSL.Value & "" '  "L" ' S = Stacked, L = in Line

            Dim colors() As System.Drawing.Color =
                {System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige,
                 System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige,
                 System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige}

            Dim colorsF() As System.Drawing.Color =
                {System.Drawing.Color.ForestGreen, System.Drawing.Color.Purple, System.Drawing.Color.Orange,
                 System.Drawing.Color.Blue, System.Drawing.Color.Brown, System.Drawing.Color.Red,
                 System.Drawing.Color.LimeGreen, System.Drawing.Color.Turquoise, System.Drawing.Color.Salmon}

            Dim XC As Integer = 0
            Dim XR As Integer = 0

            GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

            Dim FILENAME As String = FORM_NAME & "_" & XNO & ".XLSX"
            Dim myWorkbook As New GemBox.Spreadsheet.ExcelFile
            Dim ws As GemBox.Spreadsheet.ExcelWorksheet = myWorkbook.Worksheets.Add(MENU_ITEM_OBJECT)

            Dim M_DESC(12) As String
            For M As Integer = 1 To 12
                Dim MONTH_DESC As String = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -1 * (12 - M)), False, True), 1, 6)
                M_DESC(M) = MONTH_DESC
            Next

            Dim R As Integer = -1
            ws.Cells(1, 0).Style.Font.Color = System.Drawing.Color.Blue
            ws.Cells(1, 0).Style.Font.Size = 300
            ws.Cells(1, 0).Style.Font.Name = "Times New Roman"
            ws.Cells(1, 0).Value = MENU_ITEM_DESC
            ws.Cells(0, 1).Value = MENU_ITEM_OBJECT
            ws.Cells(2, 0).Value = SUBT

            With ws.Cells(0, 0)
                .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Left
                .Style.NumberFormat = "mm/dd/yy;@"
                .Value = Now
            End With

            R = 2

            Dim LVLS As Integer = COLUMN_CAPTION_by_Lvl.Count - 1

            R += 2
            For C As Integer = 1 To 9
                ws.Columns(C - 1).Style.Font.Name = "Verdana"

                ws.Cells(R, C - 1).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
                ws.Cells(R, C - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

                If C <= LVLS Then
                    ws.Cells(R, C - 1).Value = COLUMN_CAPTIONs(C - 1) ' COLUMN_CAPTION_by_Lvl(C)
                    ws.Cells(R, C - 1).Style.Font.Color = colorsF(C - 1)

                    ws.Columns(C - 1).Width = 4000
                Else
                    ws.Columns(C - 1).Hidden = True
                    ws.Columns(C - 1).Width = 0
                End If
            Next

            ws.Cells(R, 9).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
            ws.Cells(R, 9).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

            ws.Cells(R, 9).Value = "Description"
            ws.Columns(9).Width = 10000

            Dim MBAY As Integer = 17
            Dim GBAY As Integer = 11

            Dim Ms As New Dictionary(Of Integer, Integer)
            Dim DATA_TYPEs(dst.Tables("ASTRECAP").Rows.Count) As String

            Dim M_MAX As Integer = 0
            For Each row As DataRow In dst.Tables("ASTRECAP").Select("", "ASTSRPT1_RECAP_ROW_NO")
                M_MAX += 1

                Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(row.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                Ms.Add(ASTSRPT1_RECAP_ROW_NO, M_MAX)

                Dim M As Integer = M_MAX

                DATA_TYPEs(M) = row.Item("ASTSRPT1_RECAP_ROW_CAPTION") & ""

                If opt <> "S" Or M = 1 Then
                    If opt = "S" Then
                        ws.Cells(R, GBAY - 1).Value = "Data Type"
                        ws.Cells(R, GBAY - 1).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
                        ws.Cells(R, GBAY - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

                        ws.Columns(GBAY - 1).Width = 4000
                    Else
                        ws.Cells(R - 1, GBAY + MBAY * (M - 1)).Value = row.Item("ASTSRPT1_RECAP_ROW_CAPTION")
                        ws.Columns(GBAY + MBAY * (M - 1) + 0 - 1).Width = 100
                    End If
                    For C As Integer = 1 To 12
                        XC = GBAY + MBAY * (M - 1) + C - 1
                        With ws.Columns(XC)
                            .Style.Font.Name = "Verdana"
                            .Style.NumberFormat = "#,##0;@"
                            .Style.Font.Name = "Verdana"
                            .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                            .Width = 3000
                        End With
                        With ws.Cells(R, XC)
                            .Value = M_DESC(C)
                            .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                            .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                        End With
                    Next

                    XC = GBAY + MBAY * (M - 1) + 13 - 1
                    With ws.Columns(XC)
                        .Style.Font.Name = "Verdana"
                        .Style.NumberFormat = "#,##0;@"
                        .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                        .Width = 3000
                    End With
                    With ws.Cells(R, XC)
                        .Value = "PTD"
                        .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                        .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                    End With

                    XC = GBAY + MBAY * (M - 1) + 14 - 1
                    With ws.Columns(XC)
                        .Style.Font.Name = "Verdana"
                        .Style.NumberFormat = "#,##0.0%;@"
                        '.Style.NumberFormat = "#,##0.0%"
                        .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                        .Width = 3000
                    End With
                    With ws.Cells(R, XC)
                        .Value = "%TTL"
                        .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                        .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                    End With

                    XC = GBAY + MBAY * (M - 1) + 15 - 1
                    With ws.Columns(XC)
                        .Style.Font.Name = "Verdana"
                        .Style.NumberFormat = "#,##0;@"
                        .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                        .Width = 3000
                    End With
                    With ws.Cells(R, XC)
                        .Value = "Total"
                        .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                        .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                    End With

                    XC = GBAY + MBAY * (M - 1) + 16 - 1
                    With ws.Columns(XC)
                        .Style.Font.Name = "Verdana"
                        .Style.NumberFormat = "#,##0.0%;@"
                        ' .Style.NumberFormat = "#,##0.0%"
                        .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                        .Width = 3000
                    End With
                    With ws.Cells(R, XC)
                        .Value = "%TTL"
                        .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                        .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                    End With
                End If
            Next

            Dim LAST_LVL As Integer = 0
            Dim LAST_KEY = ""
            Dim START_ROW As Integer = -1
            Dim LEVEL_LAST_ROW As Integer = 0
            Dim C_LVL As Integer = 0

            Dim START_ROWS(LVLS) As List(Of Integer)
            For L As Integer = 0 To LVLS
                START_ROWS(L) = New List(Of Integer)
            Next

            Dim MROW As Integer = 0
            Dim RSTART As Integer = R
            Dim CODEs(LVLS) As String
            Dim First_Row_at_Level As Boolean = False

            Dim sqlx As String = ""
            For i As Integer = 1 To 9
                sqlx &= " or G" & CStr(i) & " <> '" & aRC & "'"
            Next
            ' THESE ANNOYING REPORT TOTAL RECORDS ARE NOW DELETED IN ASFSRPTM.Build_Report_File

            For Each row As DataRow In dst.Tables("ASTSRPT1").Select(Mid(sqlx, 5), "G1,G2,G3,G4,G5,G6,G7,G8,G9,ASTSRPT1_RECAP_ROW_NO")
                Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(row.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                Dim rowASTRECAP As DataRow = dst.Tables("ASTRECAP").Rows.Find(ASTSRPT1_RECAP_ROW_NO)
                Dim GRN As String = rowASTRECAP.Item("GRN")
                Dim CALC As String = rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CALC") & ""
                Dim TYPE As String = rowASTRECAP.Item("TYPE") & ""

                Dim M As Integer = Ms(ASTSRPT1_RECAP_ROW_NO)
                Dim THIS_KEY = ""
                For I As Integer = 1 To 9
                    THIS_KEY &= vbTab & row.Item("G" & CStr(I))
                Next

                If LAST_KEY <> THIS_KEY Then
                    LAST_KEY = THIS_KEY
                    LEVEL_LAST_ROW = C_LVL
                    If opt = "S" Then
                        MROW += 1
                    Else
                        R += 1
                    End If
                    ReDim CODEs(LVLS)
                    First_Row_at_Level = True
                Else
                    First_Row_at_Level = False
                End If

                If opt = "S" Then
                    R = RSTART + (MROW - 1) * M_MAX + M
                End If

                Dim CODE_VALUE As String = ""
                C_LVL = 0
                For C As Integer = 1 To LVLS
                    Dim Z As String = row.Item("G" & CStr(C)) & ""
                    If InStr(Z, ":") = 0 Then
                        Exit For
                    End If
                    CODE_VALUE = Split(Z & ":", ":")(1)
                    C_LVL = C
                    If First_Row_at_Level Then
                        CODEs(C) = CODE_VALUE
                        ws.Cells(R, C - 1).Value = CODE_VALUE
                        ' this could be done at end for entire range of cells
                        ws.Cells(R, C - 1).Style.Font.Color = colorsF(C - 1) ' see end of loop
                    End If
                Next

                If First_Row_at_Level Then
                    If LEVEL_LAST_ROW <> C_LVL Then
                        If C_LVL = LVLS Then
                            START_ROW = R
                        End If
                        If C_LVL > LEVEL_LAST_ROW Then
                            For CL As Integer = LEVEL_LAST_ROW + 1 To C_LVL
                                START_ROWS(CL).Clear()
                            Next
                        End If
                    End If
                End If

                If opt = "S" Then
                    If First_Row_at_Level Then
                        For MM As Integer = 1 To M_MAX
                            ws.Cells(RSTART + (MROW - 1) * M_MAX + MM, GBAY - 1).Value = DATA_TYPEs(MM)
                            If C_LVL <> 0 Then ws.Cells(RSTART + (MROW - 1) * M_MAX + MM, GBAY - 1).Style.Font.Color = colorsF(C_LVL - 1)
                        Next
                        If C_LVL = LVLS Then START_ROWS(C_LVL).Add(R)
                    End If
                End If

                Dim rowASTGROUP As DataRow = Nothing
                Dim DESC_VALUE As String
                If First_Row_at_Level Then
                    If C_LVL > 0 Then
                        rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(COLUMN_CAPTIONs(C_LVL - 1) & ":" & CODE_VALUE)
                        'rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(COLUMN_CAPTION_by_Lvl(C_LVL) & ":" & CODE_VALUE)
                    End If
                    DESC_VALUE = ""
                    If rowASTGROUP IsNot Nothing Then
                        DESC_VALUE = rowASTGROUP.Item("GROUP_DESC")
                        ws.Cells(R, 9).Value = DESC_VALUE
                        ws.Cells(R, 9).Style.Font.Color = colorsF(C_LVL - 1)
                        ws.Cells(R, 9).Style.Indent = (C_LVL - 1) * 1
                    End If
                End If

                Dim MC As Integer = IIf(opt = "S", 1, M)

                If TYPE = "P" Then
                    For C As Integer = 1 To 16
                        If C = 14 Or C = 16 Then
                        Else
                            ws.Cells(R, GBAY + MBAY * (MC - 1) + C - 1).Style.NumberFormat = "#,##0.0"
                        End If
                    Next
                End If

                If C_LVL = LVLS Then
                    For C As Integer = 1 To 12
                        XC = GBAY + MBAY * (MC - 1) + C - 1
                        If GRN = "*" Then  ' ROW IS A CALC
                            ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                        Else
                            ws.Cells(R, XC).Value = row.Item("M" & Format(C, "00"))
                        End If
                    Next

                    XC += 1
                    If GRN = "*" Then  ' ROW IS A CALC
                        ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                    Else
                        ws.Cells(R, XC).Value = row.Item("YTD")
                    End If
                    XC += 1

                    XC += 1
                    If GRN = "*" Then  ' ROW IS A CALC
                        ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                    Else
                        ws.Cells(R, XC).Value = row.Item("TOTAL")
                    End If
                    XC += 1

                Else
                    If C_LVL = LVLS - 1 And opt <> "S" Then
                        For C As Integer = 1 To 12
                            XC = GBAY + MBAY * (MC - 1) + C - 1

                            If GRN = "*" Then  ' ROW IS A CALC
                                ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                            Else
                                ws.Cells(R, XC).Formula = "=SUM(" & ASCMAIN1.Excel_Cell(START_ROW + 1, XC + 1) & ":" & ASCMAIN1.Excel_Cell(R - 1 + 1, XC + 1) & ")"
                            End If

                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)
                        Next

                        ' YTD & Total - 1 line
                        For y As Integer = 1 To 2
                            XC += 1

                            If GRN = "*" Then  ' ROW IS A CALC
                                ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                            Else
                                Dim C_first As String = ASCMAIN1.Excel_Cell(START_ROW + 1, XC + 1)
                                Dim C_last As String = ASCMAIN1.Excel_Cell(R - 1 + 1, XC + 1)
                                If C_first = C_last Then
                                    ws.Cells(R, XC).Formula = "=SUM(" & C_first & ")"
                                Else
                                    ws.Cells(R, XC).Formula = "=SUM(" & C_first & ":" & C_last & ")"
                                End If
                            End If

                            Dim CT As String = ASCMAIN1.Excel_Cell(R + 1, XC + 1)
                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)

                            XC += 1
                            For RPCT As Integer = START_ROW To R - 1
                                Dim C As String = ASCMAIN1.Excel_Cell(RPCT + 1, XC + 1)

                                'http://answers.microsoft.com/en-us/office/forum/office_2010-excel/excel-2010-inconsistent-formula-flag/1822122a-9c73-494d-8a54-5f0ba8853cab?auth=1
                                'http://answers.microsoft.com/en-us/office/forum/office_2007-excel/inappropriate-display-of-this-cell-is-inconsistent/1f2c144e-a61c-4dac-b6d2-b71db3da24f8

                                ws.Cells(RPCT, XC).Formula = "=IF(" & CT & "=0,0," & C & "/" & CT & ")"
                                '  ws.Cells(RPCT, XC).Formula = "= " & C & "/" & CT
                            Next
                        Next

                        If Not START_ROWS(C_LVL).Contains(R) Then
                            For rr As Integer = START_ROW To R - 1
                                ws.Rows(rr).OutlineLevel = LVLS
                            Next

                            START_ROWS(C_LVL).Add(R)
                        End If
                    Else
                        Dim TT As String = ""
                        For Each RR As Integer In START_ROWS(C_LVL + 1)
                            TT &= "," & ASCMAIN1.Excel_Cell(RR + IIf(opt = "S", M, 1), 1)
                            If First_Row_at_Level Then
                                For RRM As Integer = RR To RR + (M_MAX - 1)
                                    ws.Rows(RRM).OutlineLevel = C_LVL + 1
                                Next
                            End If
                        Next

                        Dim TTC As String = ""
                        For C As Integer = 1 To 12
                            XC = GBAY + MBAY * (MC - 1) + C - 1
                            TTC = ASCMAIN1.Excel_Cell(0, XC + 1)
                            If GRN = "*" Then  ' ROW IS A CALC
                                ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                            Else
                                ws.Cells(R, XC).Formula = "=SUM(" & Mid(Replace(TT, ",A", "," & TTC), 2) & ")"
                            End If

                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)
                        Next

                        ' YTD & Total - stacked
                        For y As Integer = 1 To 2
                            XC += 1
                            TTC = ASCMAIN1.Excel_Cell(0, XC + 1)
                            Dim TTCR As String = Mid(Replace(TT, ",A", "," & TTC), 2)

                            If GRN = "*" Then  ' ROW IS A CALC
                                ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                            Else
                                ws.Cells(R, XC).Formula = "=SUM(" & TTCR & ")"
                            End If

                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)

                            Dim CT As String = Excel_Cell0(R, XC)
                            XC += 1


                            Dim TTCpct As String = ASCMAIN1.Excel_Cell(0, XC + 1)
                            For Each C As String In Split(TTCR, ",")
                                Dim C_PCT As String = Replace(C, TTC, TTCpct) ' ASCMAIN1.Excel_Cell(R - 1, XC + 1)
                                ws.Cells(C_PCT).Formula = "=IF(" & CT & "=0,0," & C & "/" & CT & ")"
                                ' ws.Cells(C_PCT).Formula = "= " & C & "/" & CT
                            Next
                        Next

                        If First_Row_at_Level Then
                            START_ROWS(C_LVL).Add(R)
                            If C_LVL = 0 Then ws.Cells(R, 0).Value = "Totals"
                        End If
                    End If
                End If
            Next

            Dim RFINAL As Integer = R
            If opt = "S" Then
                RFINAL = RSTART + MROW * M_MAX - 1
            End If

            Gembox_Export_to_Excel_Show(myWorkbook, FILENAME)


        Catch ex As Exception
            If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
        End Try
    End Sub

    Sub Create_XLS()
        Try
            'Prepare_XLS()

            Dim opt As String = optSL.Value & "" '  "L" ' S = Stacked, L = in Line

            Dim colors() As System.Drawing.Color =
                {System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige,
                 System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige,
                 System.Drawing.Color.PaleGreen, System.Drawing.Color.PaleGoldenrod, System.Drawing.Color.PaleTurquoise, System.Drawing.Color.Beige}

            Dim colorsF() As System.Drawing.Color =
                {System.Drawing.Color.ForestGreen, System.Drawing.Color.Purple, System.Drawing.Color.Salmon,
                 System.Drawing.Color.Blue, System.Drawing.Color.Brown, System.Drawing.Color.Red,
                 System.Drawing.Color.LimeGreen, System.Drawing.Color.Turquoise, System.Drawing.Color.Orange}

            Dim XC As Integer = 0
            Dim XR As Integer = 0

            GemBox.Spreadsheet.SpreadsheetInfo.SetLicense(ASCMAIN1.GemboxKey)

            Dim FILENAME As String = FORM_NAME & "_" & XNO & ".XLSX"
            Dim myWorkbook As New GemBox.Spreadsheet.ExcelFile
            Dim ws As GemBox.Spreadsheet.ExcelWorksheet = myWorkbook.Worksheets.Add(MENU_ITEM_OBJECT)

            Dim M_DESC(12) As String
            For M As Integer = 1 To 12
                Dim MONTH_DESC As String = Mid(ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -1 * (12 - M)), False, True), 1, 6)
                M_DESC(M) = MONTH_DESC
            Next

            Dim R As Integer = -1
            ws.Cells(1, 0).Style.Font.Color = System.Drawing.Color.Blue
            ws.Cells(1, 0).Style.Font.Size = 300
            ws.Cells(1, 0).Style.Font.Name = "Times New Roman"
            ws.Cells(1, 0).Value = MENU_ITEM_DESC
            ws.Cells(0, 1).Value = MENU_ITEM_OBJECT
            ws.Cells(2, 0).Value = SUBT

            With ws.Cells(0, 0)
                .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Left
                .Style.NumberFormat = "mm/dd/yy;@"
                .Value = Now
            End With

            R = 2

            Dim LVLS As Integer = COLUMN_CAPTION_by_Lvl.Count - 1

            R += 2
            For C As Integer = 1 To 9
                ws.Columns(C - 1).Style.Font.Name = "Verdana"

                ws.Cells(R, C - 1).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
                ws.Cells(R, C - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

                If C <= LVLS - IIf(opt = "P", 1, 0) Then
                    ws.Cells(R, C - 1).Value = COLUMN_CAPTIONs(C - 1) ' COLUMN_CAPTION_by_Lvl(C)
                    ws.Cells(R, C - 1).Style.Font.Color = colorsF(C - 1)

                    ws.Columns(C - 1).Width = 4000
                Else
                    ws.Columns(C - 1).Hidden = True
                    ws.Columns(C - 1).Width = 0
                End If
            Next

            ws.Cells(R, 9).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
            ws.Cells(R, 9).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

            ws.Cells(R, 9).Value = "Description"
            ws.Columns(9).Width = 10000

            Dim MBAY As Integer = 17
            Dim GBAY As Integer = 11

            Dim Ms As New Dictionary(Of Integer, Integer)
            Dim DATA_TYPEs(dst.Tables("ASTRECAP").Rows.Count) As String

            Dim pCmax As Integer = 0
            Dim pvtCodes As New List(Of String)
            Dim pvtCodeSlots As New Dictionary(Of String, Integer)

            If opt = "P" Then
                'Dim view As New DataView(dt)
                'Dim distinctValues As New DataTable()
                'distinctValues = view.ToTable(True, ColumnName)
                Dim tblPvt As DataTable = ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), New String() {"G" & CStr(LVLS)})
                For Each rowPvt As DataRow In tblPvt.Select("", "G" & CStr(LVLS))
                    Dim pvtCode As String = rowPvt.Item(0) & ""
                    If pvtCode.Contains(":") Then
                        Dim CODE As String = Split(pvtCode, ":")(1)
                        pvtCodes.Add(CODE)
                        pCmax += 1
                        pvtCodeSlots.Add(CODE, pCmax)
                    End If
                Next
                ' pCmax = pvtCodes.Count + 1 ' add 1 for total
                pCmax += 1 ' add 1 for total
            End If

            Dim M_MAX As Integer = 0
            For Each row As DataRow In dst.Tables("ASTRECAP").Select("", "ASTSRPT1_RECAP_ROW_NO")
                M_MAX += 1

                Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(row.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                Ms.Add(ASTSRPT1_RECAP_ROW_NO, M_MAX)

                Dim M As Integer = M_MAX

                DATA_TYPEs(M) = row.Item("ASTSRPT1_RECAP_ROW_CAPTION") & ""

                If opt = "L" Or M = 1 Then
                    If opt <> "L" Then
                        ws.Cells(R, GBAY - 1).Value = "Data Type"
                        ws.Cells(R, GBAY - 1).Style.FillPattern.PatternForegroundColor = System.Drawing.Color.LightGray
                        ws.Cells(R, GBAY - 1).Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid

                        ws.Columns(GBAY - 1).Width = 4000
                    Else
                        ws.Cells(R - 1, GBAY + MBAY * (M - 1)).Value = row.Item("ASTSRPT1_RECAP_ROW_CAPTION")
                        ws.Columns(GBAY + MBAY * (M - 1) + 0 - 1).Width = 100
                    End If

                    For pC As Integer = 1 To IIf(pCmax = 0, 1, pCmax) ' will iterate once even if pCmax = 0
                        For C As Integer = 1 To 12

                            XC = GBAY + (MBAY + pCmax) * (M - 1) + C - 1 + pCmax * (C - 1) + (pC - 1)

                            With ws.Columns(XC)
                                .Style.Font.Name = "Verdana"
                                .Style.NumberFormat = "#,##0;@"
                                .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                                .Width = 3000
                            End With
                            With ws.Cells(R, XC)
                                If opt = "P" Then
                                    If pC = pCmax Then
                                        .Value = "Total"
                                    Else
                                        .Value = pvtCodes(pC - 1)
                                    End If
                                Else
                                    .Value = M_DESC(C)
                                End If
                                .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                            End With

                            If opt = "P" And pC = pCmax Then
                                With ws.Cells(R - 1, XC)
                                    .Value = M_DESC(C)
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With

                                XC += 1
                                ws.Columns(XC).Width = 100
                            End If
                        Next

                        XC = GBAY + (MBAY + pCmax) * (M - 1) + 13 - 1 + pCmax * (13 - 1) + (pC - 1)
                        With ws.Columns(XC)
                            .Style.Font.Name = "Verdana"
                            .Style.NumberFormat = "#,##0;@"
                            .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                            .Width = 3000
                        End With
                        If opt = "P" Then
                            If pC = pCmax Then
                                With ws.Cells(R - 1, XC)
                                    .Value = "PTD"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                                With ws.Cells(R, XC)
                                    .Value = "Total"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With

                                XC += 1
                                ws.Columns(XC).Width = 100
                            Else
                                With ws.Cells(R, XC)
                                    .Value = pvtCodes(pC - 1)
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                            End If
                        Else
                            With ws.Cells(R, XC)
                                .Value = "PTD"
                                .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                            End With
                        End If

                        XC = GBAY + (MBAY + pCmax) * (M - 1) + 14 - 1 + pCmax * (14 - 1) + (pC - 1)
                        With ws.Columns(XC)
                            .Style.Font.Name = "Verdana"
                            .Style.NumberFormat = "#,##0.0%;@"
                            '.Style.NumberFormat = "#,##0.0%"
                            .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                            .Width = 3000
                        End With
                        If opt = "P" Then
                            If pC = pCmax Then
                                With ws.Cells(R - 1, XC)
                                    .Value = "%TTL"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                                With ws.Cells(R, XC)
                                    .Value = "Total"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With

                                XC += 1
                                ws.Columns(XC).Width = 100
                            Else
                                With ws.Cells(R, XC)
                                    .Value = pvtCodes(pC - 1)
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                            End If
                        Else
                            With ws.Cells(R, XC)
                                .Value = "%TTL"
                                .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                            End With
                        End If


                        XC = GBAY + (MBAY + pCmax) * (M - 1) + 15 - 1 + pCmax * (15 - 1) + (pC - 1)
                        With ws.Columns(XC)
                            .Style.Font.Name = "Verdana"
                            .Style.NumberFormat = "#,##0;@"
                            .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                            .Width = 3000
                        End With
                        If opt = "P" Then
                            If pC = pCmax Then
                                With ws.Cells(R - 1, XC)
                                    .Value = "Total"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                                With ws.Cells(R, XC)
                                    .Value = "Total"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With

                                XC += 1
                                ws.Columns(XC).Width = 100
                            Else
                                With ws.Cells(R, XC)
                                    .Value = pvtCodes(pC - 1)
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                            End If
                        Else
                            With ws.Cells(R, XC)
                                .Value = "Total"
                                .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                            End With
                        End If


                        XC = GBAY + (MBAY + pCmax) * (M - 1) + 16 - 1 + pCmax * (16 - 1) + (pC - 1)
                        With ws.Columns(XC)
                            .Style.Font.Name = "Verdana"
                            .Style.NumberFormat = "#,##0.0%;@"
                            ' .Style.NumberFormat = "#,##0.0%"
                            .Style.HorizontalAlignment = GemBox.Spreadsheet.HorizontalAlignmentStyle.Right
                            .Width = 3000
                        End With
                        If opt = "P" Then
                            If pC = pCmax Then
                                With ws.Cells(R - 1, XC)
                                    .Value = "%TTL"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                                With ws.Cells(R, XC)
                                    .Value = "Total"
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With

                                XC += 1
                                ws.Columns(XC).Width = 100
                            Else
                                With ws.Cells(R, XC)
                                    .Value = pvtCodes(pC - 1)
                                    .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                    .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                                End With
                            End If
                        Else
                            With ws.Cells(R, XC)
                                .Value = "%TTL"
                                .Style.FillPattern.PatternForegroundColor = colors(M - 1)
                                .Style.FillPattern.PatternStyle = GemBox.Spreadsheet.FillPatternStyle.Solid
                            End With
                        End If
                    Next
                End If
            Next

            Dim LAST_LVL As Integer = 0
            Dim LAST_KEY = ""
            Dim START_ROW As Integer = -1
            Dim LEVEL_LAST_ROW As Integer = 0
            Dim C_LVL As Integer = 0

            Dim START_ROWS(LVLS) As List(Of Integer)
            For L As Integer = 0 To LVLS
                START_ROWS(L) = New List(Of Integer)
            Next

            Dim MROW As Integer = 0
            Dim RSTART As Integer = R
            Dim CODEs(LVLS) As String
            Dim First_Row_at_Level As Boolean = False

            Dim sqlx As String = ""
            For i As Integer = 1 To 9
                sqlx &= " Or G" & CStr(i) & " <> '" & aRC & "'"
            Next
            ' THESE ANNOYING REPORT TOTAL RECORDS ARE NOW DELETED IN ASFSRPTM.Build_Report_File

            For Each row As DataRow In dst.Tables("ASTSRPT1").Select(Mid(sqlx, 5), "G1,G2,G3,G4,G5,G6,G7,G8,G9,ASTSRPT1_RECAP_ROW_NO")
                Dim ASTSRPT1_RECAP_ROW_NO As Integer = Val(row.Item("ASTSRPT1_RECAP_ROW_NO") & "")
                Dim rowASTRECAP As DataRow = dst.Tables("ASTRECAP").Rows.Find(ASTSRPT1_RECAP_ROW_NO)
                Dim GRN As String = rowASTRECAP.Item("GRN")
                Dim CALC As String = rowASTRECAP.Item("ASTSRPT1_RECAP_ROW_CALC") & ""
                Dim TYPE As String = rowASTRECAP.Item("TYPE") & ""

                Dim M As Integer = Ms(ASTSRPT1_RECAP_ROW_NO)
                Dim THIS_KEY = ""
                For I As Integer = 1 To LVLS ' - IIf(pCmax = 0, 0, 1) ' 9
                    THIS_KEY &= vbTab & row.Item("G" & CStr(I))
                Next

                If LAST_KEY <> THIS_KEY Then
                    LAST_KEY = THIS_KEY
                    LEVEL_LAST_ROW = C_LVL
                    If opt = "L" Then
                        R += 1
                    Else
                        If opt = "P" And C_LVL = LVLS Then
                            ' DO NOT ADVANCE MROW
                        Else
                            MROW += 1
                        End If
                    End If
                    ReDim CODEs(LVLS)
                    First_Row_at_Level = True
                Else
                    First_Row_at_Level = False
                End If

                If opt = "S" Or opt = "P" Then
                    R = RSTART + (MROW - 1) * M_MAX + M
                End If

                Dim CODE_VALUE As String = ""
                C_LVL = 0
                For C As Integer = 1 To LVLS ' - IIf(pCmax = 0, 0, 1)
                    Dim Z As String = row.Item("G" & CStr(C)) & ""
                    If InStr(Z, ":") = 0 Then
                        Exit For
                    End If
                    CODE_VALUE = Split(Z & ":", ":")(1)
                    C_LVL = C
                    If First_Row_at_Level Then
                        CODEs(C) = CODE_VALUE
                        If opt = "P" And C_LVL = LVLS Then
                        Else
                            ws.Cells(R, C - 1).Value = CODE_VALUE
                        End If
                        ' this could be done at end for entire range of cells
                        ws.Cells(R, C - 1).Style.Font.Color = colorsF(C - 1) ' see end of loop
                    Else
                        ' this else code was created to copy codes down to all rows
                        If chkCODES_ON_ALL_LINES.Checked Then
                            ws.Cells(R, C - 1).Value = CODE_VALUE
                            ws.Cells(R, C - 1).Style.Font.Color = colorsF(C - 1) ' see end of loop
                        End If
                    End If
                Next

                If First_Row_at_Level Then
                    If LEVEL_LAST_ROW <> C_LVL Then
                        If C_LVL = LVLS Then
                            START_ROW = R
                        End If
                        If C_LVL > LEVEL_LAST_ROW Then
                            For CL As Integer = LEVEL_LAST_ROW + 1 To C_LVL
                                START_ROWS(CL).Clear()
                            Next
                        End If
                    End If
                End If

                If opt = "S" Or opt = "P" Then
                    If First_Row_at_Level Then
                        For MM As Integer = 1 To M_MAX
                            ws.Cells(RSTART + (MROW - 1) * M_MAX + MM, GBAY - 1).Value = DATA_TYPEs(MM)
                            If C_LVL <> 0 Then ws.Cells(RSTART + (MROW - 1) * M_MAX + MM, GBAY - 1).Style.Font.Color = colorsF(C_LVL - 1)
                        Next
                        If C_LVL = LVLS Then START_ROWS(C_LVL).Add(R)
                    End If
                End If

                Dim rowASTGROUP As DataRow = Nothing
                Dim DESC_VALUE As String
                If First_Row_at_Level Then
                    Dim C_LVLx As Integer = C_LVL - 1
                    If C_LVL > 0 Then
                        If C_LVL = LVLS And opt = "P" Then
                            C_LVLx -= 1
                            rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(row.Item("G" & CStr(LVLS - 1)))
                        Else
                            rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(COLUMN_CAPTIONs(C_LVLx) & ":" & CODE_VALUE)
                        End If
                    End If
                    DESC_VALUE = ""
                    If rowASTGROUP IsNot Nothing Then
                        DESC_VALUE = rowASTGROUP.Item("GROUP_DESC")
                        ws.Cells(R, 9).Value = DESC_VALUE
                        ws.Cells(R, 9).Style.Font.Color = colorsF(C_LVLx)
                        ws.Cells(R, 9).Style.Indent = (C_LVLx) * 1
                    End If
                End If

                Dim MC As Integer = IIf(opt = "L", M, 1)
                Dim pCx As Integer = 0
                If opt = "P" Then
                    Dim CODE As String = row.Item("G" & CStr(LVLS))
                    If CODE.Contains(":") Then
                        CODE = Split(CODE, ":")(1)
                    End If
                    If pvtCodeSlots.ContainsKey(CODE) Then
                        pCx = pvtCodeSlots(CODE) - 1
                    Else
                        pCx = pCmax - 1
                    End If
                End If

                If TYPE = "P" Then
                    For C As Integer = 1 To 16
                        If C = 14 Or C = 16 Then
                        Else
                            ws.Cells(R, GBAY + (MBAY + pCmax) * (MC - 1) + C - 1 + (C - 1) * pCmax + pCx).Style.NumberFormat = "#,##0.0"
                        End If
                    Next
                End If

                If C_LVL = LVLS Then ' this is a data row

                    For C As Integer = 1 To 12
                        XC = GBAY + (MBAY + pCmax) * (MC - 1) + C - 1 + (C - 1) * pCmax + pCx
                        If GRN = "*" Then  ' ROW IS A CALC
                            ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                        Else
                            ws.Cells(R, XC).Value = row.Item("M" & Format(C, "00"))
                        End If

                        If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                            'ws.Cells(R, XC).Value = ws.Cells(R, XC).Value & "/100"
                            ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                        End If
                    Next

                    XC = GBAY + (MBAY + pCmax) * (MC - 1) + 13 - 1 + (13 - 1) * pCmax + pCx
                    If GRN = "*" Then  ' ROW IS A CALC
                        ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                    Else
                        ws.Cells(R, XC).Value = row.Item("YTD")
                    End If

                    If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                        'ws.Cells(R, XC).Value = ws.Cells(R, XC).Value & "/100"
                        ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                    End If

                    XC = GBAY + (MBAY + pCmax) * (MC - 1) + 15 - 1 + (15 - 1) * pCmax + pCx
                    If GRN = "*" Then  ' ROW IS A CALC
                        ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                    Else
                        ws.Cells(R, XC).Value = row.Item("TOTAL")
                    End If

                    If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                        'ws.Cells(R, XC).Value = ws.Cells(R, XC).Value & "/100"
                        ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                    End If


                Else ' this is a PB summary row

                    If C_LVL = LVLS - 1 And opt = "L" Then
                        For C As Integer = 1 To 12
                            XC = GBAY + (MBAY + pCmax) * (MC - 1) + C - 1 + (C - 1) * pCmax + pCx
                            If GRN = "*" Then  ' ROW IS A CALC
                                ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                            Else
                                ws.Cells(R, XC).Formula = "=SUM(" & ASCMAIN1.Excel_Cell(START_ROW + 1, XC + 1) & ":" & ASCMAIN1.Excel_Cell(R - 1 + 1, XC + 1) & ")"
                            End If

                            If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                            End If

                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)
                        Next

                        ' YTD & Total - 1 line
                        For y As Integer = 1 To 2
                            'XC += 1
                            XC = GBAY + (MBAY + pCmax) * (MC - 1) + 13 - 1 + (13 - 1) * pCmax + pCx
                            If GRN = "*" Then  ' ROW IS A CALC
                                ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                            Else
                                Dim C_first As String = ASCMAIN1.Excel_Cell(START_ROW + 1, XC + 1)
                                Dim C_last As String = ASCMAIN1.Excel_Cell(R - 1 + 1, XC + 1)
                                If C_first = C_last Then
                                    ws.Cells(R, XC).Formula = "=SUM(" & C_first & ")"
                                Else
                                    ws.Cells(R, XC).Formula = "=SUM(" & C_first & ":" & C_last & ")"
                                End If
                            End If

                            If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                            End If

                            Dim CT As String = ASCMAIN1.Excel_Cell(R + 1, XC + 1)
                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)

                            'XC += 1
                            XC = GBAY + (MBAY + pCmax) * (MC - 1) + 15 - 1 + (15 - 1) * pCmax + pCx
                            For RPCT As Integer = START_ROW To R - 1
                                Dim C As String = ASCMAIN1.Excel_Cell(RPCT + 1, XC + 1)

                                'http://answers.microsoft.com/en-us/office/forum/office_2010-excel/excel-2010-inconsistent-formula-flag/1822122a-9c73-494d-8a54-5f0ba8853cab?auth=1
                                'http://answers.microsoft.com/en-us/office/forum/office_2007-excel/inappropriate-display-of-this-cell-is-inconsistent/1f2c144e-a61c-4dac-b6d2-b71db3da24f8

                                ws.Cells(RPCT, XC).Formula = "=IF(" & CT & "=0,0," & C & "/" & CT & ")"
                                '  ws.Cells(RPCT, XC).Formula = "= " & C & "/" & CT
                            Next
                        Next

                        If Not START_ROWS(C_LVL).Contains(R) Then
                            For rr As Integer = START_ROW To R - 1
                                ws.Rows(rr).OutlineLevel = LVLS
                            Next

                            START_ROWS(C_LVL).Add(R)
                        End If

                    Else ' not (If C_LVL = LVLS - 1 And opt = "L" Then)

                        Dim TT As String = ""
                        For Each RR As Integer In START_ROWS(C_LVL + 1)
                            TT &= "," & ASCMAIN1.Excel_Cell(RR + IIf(opt = "S" Or opt = "P", M, 1), 1)
                            If First_Row_at_Level Then
                                For RRM As Integer = RR To RR + (M_MAX - 1)
                                    ws.Rows(RRM).OutlineLevel = C_LVL + 1
                                Next
                            End If
                        Next

                        Dim TTC As String = ""
                        For C As Integer = 1 To 12
                            XC = GBAY + (MBAY + pCmax) * (MC - 1) + C - 1 + (C - 1) * pCmax + pCx
                            TTC = ASCMAIN1.Excel_Cell(0, XC + 1)

                            If opt = "P" Then
                                If GRN = "*" Then  ' ROW IS A CALC

                                    ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                                Else
                                    Dim FP As String = ""
                                    For pvtI As Integer = 1 To pvtCodes.Count
                                        Dim TTCF As String = ASCMAIN1.Excel_Cell(R + 1, XC - pvtCodes.Count + pvtI)
                                        FP &= "+" & TTCF
                                    Next
                                    ws.Cells(R, XC).Formula = "=" & Mid(FP, 2)
                                End If

                                If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                    ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                                End If

                                ' CREATE TOTALS FOR EACH OF THE BRANDS - NOT FOR THE CUSTOMER LEVEL
                                If C_LVL = LVLS - 1 Then
                                    ' DO NOTHING - WE ALREADY LOADED THE BRAND DATA INTO THE SLOTS
                                Else
                                    ' DOWN TOTAL EACH OF THE SLOTS
                                    For pvtI As Integer = 1 To pvtCodes.Count
                                        Dim Z As String = ""
                                        Dim XCZ As Integer = XC - pvtCodes.Count + pvtI - 1
                                        If GRN = "*" Then  ' ROW IS A CALC
                                            Z = CALCx(CALC, R, XCZ, ASTSRPT1_RECAP_ROW_NO, opt)
                                        Else
                                            Dim TTCF As String = ASCMAIN1.Excel_Cell(0, XC - pvtCodes.Count + pvtI)
                                            Z = "=SUM(" & Mid(Replace(TT, ",A", "," & TTCF), 2) & ")"
                                        End If

                                        ws.Cells(R, XCZ).Formula = Z

                                        If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                            ws.Cells(R, XCZ).Style.NumberFormat = "#,##0.0%;@"
                                        End If
                                    Next
                                End If

                            Else
                                If GRN = "*" Then  ' ROW IS A CALC
                                    ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                                Else
                                    ws.Cells(R, XC).Formula = "=SUM(" & Mid(Replace(TT, ",A", "," & TTC), 2) & ")"
                                End If

                                If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                    ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                                End If
                            End If

                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)
                        Next

                        ' YTD & Total - stacked
                        For C As Integer = 13 To 15 Step 2
                            XC = GBAY + (MBAY + pCmax) * (MC - 1) + C - 1 + (C - 1) * pCmax + pCx
                            TTC = ASCMAIN1.Excel_Cell(0, XC + 1)
                            Dim TTCR As String = Mid(Replace(TT, ",A", "," & TTC), 2)

                            If opt = "P" Then
                                Dim FP As String = ""
                                For pvtI As Integer = 1 To pvtCodes.Count
                                    Dim TTCF As String = ASCMAIN1.Excel_Cell(R + 1, XC - pvtCodes.Count + pvtI)
                                    FP &= "+" & TTCF
                                Next
                                ws.Cells(R, XC).Formula = "=" & Mid(FP, 2)

                                If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                    ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                                End If

                                ' CREATE TOTALS FOR EACH OF THE BRANDS - NOT FOR THE CUSTOMER LEVEL
                                If C_LVL = LVLS - 1 Then
                                    ' DO NOTHING - WE ALREADY LOADED THE BRAND DATA INTO THE SLOTS
                                Else

                                    ' DOWN TOTAL EACH OF THE SLOTS
                                    For pvtI As Integer = 1 To pvtCodes.Count
                                        Dim Z As String = ""
                                        Dim XCZ As Integer = XC - pvtCodes.Count + pvtI - 1

                                        If GRN = "*" Then  ' ROW IS A CALC
                                            Z = CALCx(CALC, R, XCZ, ASTSRPT1_RECAP_ROW_NO, opt)
                                        Else
                                            Dim TTCF As String = ASCMAIN1.Excel_Cell(0, XC - pvtCodes.Count + pvtI)
                                            Z = "=SUM(" & Mid(Replace(TT, ",A", "," & TTCF), 2) & ")"
                                        End If

                                        ws.Cells(R, XCZ).Formula = Z

                                        If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                            ws.Cells(R, XCZ).Style.NumberFormat = "#,##0.0%;@"
                                        End If
                                    Next
                                End If

                            Else
                                If GRN = "*" Then  ' ROW IS A CALC
                                    ws.Cells(R, XC).Formula = CALCx(CALC, R, XC, ASTSRPT1_RECAP_ROW_NO, opt)
                                Else
                                    ws.Cells(R, XC).Formula = "=SUM(" & TTCR & ")"
                                End If

                                If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                    ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                                End If
                            End If
                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)

                            Dim CT As String = Excel_Cell0(R, XC)
                            Dim CTXC As Integer = XC

                            XC = GBAY + (MBAY + pCmax) * (MC - 1) + C + 1 - 1 + (C + 1 - 1) * pCmax + pCx

                            If opt = "P" Then
                                Dim FP As String = ""
                                For pvtI As Integer = 1 To pvtCodes.Count
                                    Dim TTCF As String = ASCMAIN1.Excel_Cell(R + 1, XC - pvtCodes.Count + pvtI)
                                    FP &= "+" & TTCF
                                Next
                                ws.Cells(R, XC).Formula = "=" & Mid(FP, 2)

                                If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                    ws.Cells(R, XC).Style.NumberFormat = "#,##0.0%;@"
                                End If

                                ' CREATE TOTALS FOR EACH OF THE BRANDS - NOT FOR THE CUSTOMER LEVEL
                                If C_LVL = LVLS - 1 Then
                                    ' DO NOTHING - WE ALREADY LOADED THE BRAND DATA INTO THE SLOTS
                                Else
                                    ' DOWN TOTAL EACH OF THE SLOTS
                                    For pvtI As Integer = 1 To pvtCodes.Count
                                        Dim TTCF As String = ASCMAIN1.Excel_Cell(0, XC - pvtCodes.Count + pvtI)
                                        Dim Z As String = "=SUM(" & Mid(Replace(TT, ",A", "," & TTCF), 2) & ")"
                                        ws.Cells(R, XC - pvtCodes.Count + pvtI - 1).Formula = Z


                                        Dim TTCpct As String = ASCMAIN1.Excel_Cell(0, XC + 1)
                                        For Each CX As String In Split(TTCR, ",")
                                            Dim C_PCT As String = Replace(CX, TTC, TTCF) ' ASCMAIN1.Excel_Cell(R - 1, XC + 1)

                                            Dim CTP As String = Excel_Cell(0, CTXC - pvtCodes.Count + pvtI)
                                            Dim CTPT As String = Replace(CT, TTC, CTP)
                                            Dim CXPT As String = Replace(CX, TTC, CTP)
                                            ws.Cells(C_PCT).Formula = "=IF(" & CTPT & "=0,0," & CXPT & "/" & CTPT & ")"

                                            If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                                ws.Cells(C_PCT).Style.NumberFormat = "#,##0.0%;@"
                                            End If
                                        Next
                                    Next
                                End If

                            Else
                                ' once we have the XLRC for the total established (in CT), 
                                ' this section loads formulae for all of the children
                                Dim TTCpct As String = ASCMAIN1.Excel_Cell(0, XC + 1)
                                For Each CX As String In Split(TTCR, ",")
                                    Dim C_PCT As String = Replace(CX, TTC, TTCpct) ' ASCMAIN1.Excel_Cell(R - 1, XC + 1)
                                    ws.Cells(C_PCT).Formula = "=IF(" & CT & "=0,0," & CX & "/" & CT & ")"

                                    If DATA_TYPEs(Ms(ASTSRPT1_RECAP_ROW_NO)).Contains("%") Then
                                        ws.Cells(C_PCT).Style.NumberFormat = "#,##0.0%;@"
                                    End If
                                Next
                            End If
                            If C_LVL <> 0 Then ws.Cells(R, XC).Style.Font.Color = colorsF(C_LVL - 1)

                        Next

                        If First_Row_at_Level Then
                            START_ROWS(C_LVL).Add(R)
                            If C_LVL = 0 Then ws.Cells(R, 0).Value = "Totals"
                        End If
                    End If
                End If
            Next

            'Dim RFINAL As Integer = R
            'If opt = "S" Then
            '    RFINAL = RSTART + MROW * M_MAX - 1
            'End If

            Gembox_Export_to_Excel_Show(myWorkbook, FILENAME)


        Catch ex As Exception
            If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
        End Try

    End Sub
    Public Overrides Function Prepare_XLS_Summary_Columns(ByVal COLUMN_NAME_sum As Dictionary(Of String, String)) As String

        SUBT = ASCMAIN1.Get_Legend_Wk(RYW)

        'For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Rows
        '    Dim ITEM_CODE As String = rowASTSRPT1.Item("ITEM_CODE") & ""
        '    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", "")
        '    If rowICTITEM1 IsNot Nothing Then
        '        rowASTSRPT1.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        '        rowASTSRPT1.Item("LAUNCH_DATE") = rowICTITEM1.Item("LAUNCH_DATE")
        '    End If

        'Next
        '        Return "QTY_EOW,TY_WTD_S,LY_WTD_S,TY_MTD_S,LY_MTD_S,LY_MTL_S,TY_STD_S,LY_STD_S,LY_STL_S,TY_YTD_S,LY_YTD_S,LY_YTL_S"

        'If Not dst.Tables("ASTSRPT1").Columns.Contains("LAUNCH_DATE") Then
        '    With dst.Tables("ASTSRPT1")
        '        .Columns.Add("LAUNCH_DATE")
        '        .Columns.Add("WST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_WTD_S)=0,0,100*TY_WTD_S/(QTY_EOW+TY_WTD_S))")
        '        .Columns.Add("WWOH", GetType(System.Decimal), "IIF(TY_WTD_S=0,0,QTY_EOW/TY_WTD_S)")
        '        .Columns.Add("MTD_PCT", GetType(System.Decimal), "IIF(LY_MTD_S=0,0,100*(TY_MTD_S-LY_MTD_S)/LY_MTD_S)")
        '        .Columns.Add("MST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_MTD_S)=0,0,100*TY_MTD_S/(QTY_EOW+TY_MTD_S))")
        '        .Columns.Add("STD_PCT", GetType(System.Decimal), "IIF(LY_STD_S=0,0,100*(TY_STD_S-LY_STD_S)/LY_STD_S)")
        '        .Columns.Add("STL_PCT", GetType(System.Decimal), "IIF(LY_STL_S=0,0,100*TY_STD_S/LY_STL_S)")
        '        .Columns.Add("SST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_STD_S)=0,0,100*TY_STD_S/(QTY_EOW+TY_STD_S))")
        '        .Columns.Add("YTD_PCT", GetType(System.Decimal), "IIF(LY_YTD_S=0,0,100*(TY_YTD_S-LY_YTD_S)/LY_YTD_S)")
        '        .Columns.Add("YTL_PCT", GetType(System.Decimal), "IIF(LY_YTL_S=0,0,100*TY_YTD_S/LY_YTL_S)")
        '        .Columns.Add("YST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_YTD_S)=0,0,100*TY_YTD_S/(QTY_EOW+TY_YTD_S))")
        '    End With
        'End If
        Return "M01,M02,M03,M04,M05,M06,M07,M08,M09,M10,M11,M12,YTD,TOTAL"

    End Function

    Overrides Sub Prepare_XLS_Prepare_row(ByVal row As DataRow)
        'Dim GMAX As Integer = COLUMN_NAMEs.Count

        'If COLUMN_NAMEs(GMAX - 1) <> "ITEM_CODE" Then
        '    Exit Sub
        'End If

        'Dim ITEM_CODE As String = row.Item("ITEM_CODE")
        'Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE) ' LookUp("ICTITEM1", ITEM_CODE)
        'row.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        'row.Item("LAUNCH_DATE") = rowICTITEM1.Item("LAUNCH_DATE")
    End Sub

    Private Sub cmdAddDataRow_Click(sender As System.Object, e As System.EventArgs) Handles cmdAddDataRow.Click

        If txtDESC.Text = "{Enter Description}" Then
            MsgBox("Enter a Description", MsgBoxStyle.OkOnly, "Cannot Add Line")
            Exit Sub
        End If

        If optUSC.CheckedIndex = -1 Then optUSC.CheckedIndex = 0
        'If optUSC.CheckedIndex = -1 Then optUSC.Value = "U"
        If optUSC.CheckedIndex = -1 And optUSC.ValueList.ValueListItems.Count > 0 Then optUSC.Value = optUSC.ValueList.ValueListItems(0).DataValue
        ' for some reason, this optionset refuses to accept a revised valuelist
        Dim USC As String = ""
        If optGRN.Value & "" = "*" Then
        Else
            If optUSC.CheckedIndex = -1 Then
                USC = optUSC.ValueList.ValueListItems(0).DataValue
            Else
                USC = optUSC.Value
            End If
        End If


        Dim tblASTRECAP As DataTable = DirectCast(grdASTRECAP.DataSource, DataTable)
        Dim rowASTRECAP As DataRow = tblASTRECAP.NewRow
        With rowASTRECAP
            .Item("ASTSRPT1_RECAP_ROW_NO") = Val(tblASTRECAP.Compute("MAX(ASTSRPT1_RECAP_ROW_NO)", "") & "") + 1
            .Item("ASTSRPT1_RECAP_ROW_CAPTION") = txtDESC.Text
            .Item("GRN") = optGRN.Value
            .Item("HIDE") = optHIDE.Value
            .Item("TYPE") = optTYPE.Value
            .Item("USC") = USC ' optUSC.Value
            .Item("YEAR") = optYEAR.Value
            If optGRN.Value & "" = "*" Then
                .Item("ASTSRPT1_RECAP_ROW_CALC") = txtCALC.Text.ToUpper
            End If
        End With
        tblASTRECAP.Rows.Add(rowASTRECAP)

        txtDESC.Text = "{Enter Description}"
    End Sub

    Private Sub optGRN_ValueChanged(sender As Object, e As EventArgs) Handles optGRN.ValueChanged
        'optUSC.Visible = New String() {"GRNT"}.Contains(optGRN.Value)

        optUSC.ValueList = Nothing

        If "GRN".Contains(optGRN.Value) Then
            optUSC.ValueList = VLGRN
        ElseIf "T".Contains(optGRN.Value) Then
            optUSC.ValueList = VLT
        ElseIf "PMA".Contains(optGRN.Value) Then
            optUSC.ValueList = VLCM
        ElseIf "F".Contains(optGRN.Value) Then
            optUSC.ValueList = VLPA
        ElseIf optGRN.Value = "*" Then
            optUSC.ValueList = Nothing
        End If

        If optUSC.CheckedIndex = -1 And optGRN.Value <> "*" Then
            optUSC.CheckedIndex = 0
        End If

        optUSC.Visible = Not (optGRN.Value = "*")
        optYEAR.Visible = Not (optGRN.Value = "*")
        optHIDE.Visible = Not (optGRN.Value = "*")
        optTYPE.Visible = Not (optGRN.Value = "*")
        txtCALC.Visible = (optGRN.Value = "*")

    End Sub

    Private Sub optUSC_ValueChanged(sender As Object, e As EventArgs) Handles optUSC.ValueChanged

    End Sub
End Class