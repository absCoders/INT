Public Class RSRCOMP1

#Region "Declarations"
    Dim WW(1, 6) As String ' 1 = LY 0 = TY, Weeks 1 thru 6
    Dim WM(1, 2) As String ' 1 = LY 0 = TY, 0 = start, 1 = xTD, 2 = end
    Dim WS(1, 2) As String ' 1 = LY 0 = TY, 0 = start, 1 = xTD, 2 = end
    Dim WY(1, 2) As String ' 1 = LY 0 = TY, 0 = start, 1 = xTD, 2 = end

    Dim MMM(1, 2) As String ' 1 = LY 0 = TY, 0 = start, 1 = xTD, 2 = end
    Dim MMS(1, 2) As String ' 1 = LY 0 = TY, 0 = start, 1 = xTD, 2 = end
    Dim MMY(1, 2) As String ' 1 = LY 0 = TY, 0 = start, 1 = xTD, 2 = end

    Dim MM(1, 12) As String ' 1 = LY 0 = TY,  1-12

    Dim CALENDAR As String = "R"

    Dim YPCOMPMIN As String
    Dim YPCOMPMAX As String

    Dim tblSIST As New DataTable
    Dim tblSIST2 As New DataTable
#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1) '  +100, -1 + 100)
        Get_PARM("ICTPARM1")

        ASCMAIN1.sql = "Select COLUMN_NAME, NVL(COLUMN_CAPTION,COLUMN_NAME) COLUMN_CAPTION" _
        & " from ASTDSQLS where FORM_NAME = '" & Me.Name & "'" _
        & " and (COLUMN_NAME like 'TY_%' or COLUMN_NAME like 'LY_%')"
        Dim DT As DataTable = ASCDATA1.GetDataTable
        cmbRANKBY.DataSource = DT

        If ASCMAIN1.CLIENT = "INT" Then
            Dim VL As ValueList = optReportFormat.ValueList
            For I As Integer = 0 To VL.ValueListItems.Count - 1
                Dim VLI As ValueListItem = VL.ValueListItems(I)
                If VLI.DataValue.ToString.StartsWith("*") Then
                    VL.ValueListItems.Remove(VLI)
                    Exit For
                End If
            Next
        End If

        If ASCMAIN1.Running_in_VS Then
            UltraButton2.Visible = True
        End If

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        CALENDAR = optCALENDAR.Value
        Dim PERIODS_OFFSET As Int32 = 0

        Dim DATA_TYPE As String = optUS.Value

        Dim FACTOR As Integer = 1
        If Absx1.chkFor("THOUSANDS").Checked Then
            FACTOR = 1000
        End If

        Dim RSTCOMP1 As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")
        'ASCDATA1.ExecuteSQL("Alter Table " & ASTSRPT1 & " Add ITEM_CODE VARCHAR2(25)")
        'ASCMAIN1.sql = "Select 'Item Code:' || ICTITEM1.ITEM_CODE ICTITEM1_KEY, ICTITEM1.* from ICTITEM1"
        ASCMAIN1.sql = "Select ICTITEM1.* from ICTITEM1"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, , 1)

        With dst.Tables.Add("ICTITEM1_image")
            .Columns.Add("ITEM_CODE")
            .Columns.Add("ITEM_PICTURE_FILENAME")
            .Columns.Add("ITEM_PICTURE", GetType(System.Byte()))
            .PrimaryKey = New DataColumn() {.Columns("ITEM_CODE")}
        End With

        Dim SOURCE_TABLE_NAME As String = ""
        Dim by_Item As Boolean = False
        If COLUMN_NAMEs.Contains("ITEM_CODE") Then
            ' THIS NEEDS TO BE EXPANDED UPON TO INCLUDE EVEN THOSE FIELDS THAT ARE DERIVED FROM ICTITEM1
            by_Item = True
        End If

        by_Item = True
        ' ALSO NEEDED BECAUSE REPORT IS FORCE JOINED TO ICTITEM1 TO MAKE SURE WE CAN GET QTY_EOW EXPRESSED IN RETAIL $$'S - MAYBE WE SHOULD BE SUPPORTING AMT_EOW IN RSTRETL1/4

        If by_Item Then
            SOURCE_TABLE_NAME = "RSTRETL1"
        Else
            SOURCE_TABLE_NAME = "RSTRETL4"
        End If

        Dim COLUMN_NAME As String = ""
        Dim COLUMN_NAME_EOW As String = ""

        If DATA_TYPE = "UNITS" Then
            COLUMN_NAME = "NVL(RSTCOMP1.QTY_SOLD,0) / " & CStr(FACTOR)
            COLUMN_NAME_EOW = "NVL(RSTCOMP1.QTY_EOW,0) / " & CStr(FACTOR)
        ElseIf DATA_TYPE = "SALES" Then
            COLUMN_NAME = "NVL(RSTCOMP1.AMT_SOLD,0) / " & CStr(FACTOR)
            COLUMN_NAME_EOW = "NVL(RSTCOMP1.QTY_EOW,0) * NVL(ICTRETLA.ITEM_RETAIL_PRICE,NVL(ICTITEM1.ITEM_RETAIL_PRICE,0)) / " & CStr(FACTOR)
        End If

        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYWW = '" & RYW & "'"
        Dim rowGLTPARM3 As DataRow = ASCDATA1.GetDataRow
        RYP = rowGLTPARM3.Item("YYYYPP")
        Dim RYM As String = rowGLTPARM3.Item("YYYYMM")
        Dim REL_WEEK As Integer = Val(rowGLTPARM3.Item("REL_WEEK") & "")
        Dim MAX_WEEK As Integer = Val(rowGLTPARM3.Item("MAX_WEEK") & "")
        Dim RY As String = Mid(RYW, 1, 4)
        Dim RW As String = Mid(RYW, 5, 2)
        Dim RYN As Integer = Val(RY)
        Dim RWN As Integer = Val(RW)

        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -12)
        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYPP = '" & LYP & "'" _
        & " and REL_WEEK = " & CStr(REL_WEEK)
        rowGLTPARM3 = ASCDATA1.GetDataRow
        If rowGLTPARM3 Is Nothing Then
            ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYPP = '" & LYP & "'" _
            & " and REL_WEEK = MAX_WEEK"
            rowGLTPARM3 = ASCDATA1.GetDataRow
        End If
        Dim LYW As String = rowGLTPARM3.Item("YYYYWW")
        If ASCMAIN1.CLIENT = "INT" Then
            LYW = Make_LY(RYW)
        End If
        Dim LYM As String = rowGLTPARM3.Item("YYYYMM")
        Dim REL_WEEK_LY As Integer = Val(rowGLTPARM3.Item("REL_WEEK") & "")
        Dim MAX_WEEK_LY As Integer = Val(rowGLTPARM3.Item("MAX_WEEK") & "")
        Dim LY As String = Mid(LYW, 1, 4)
        Dim LW As String = Mid(LYW, 5, 2)
        Dim LYN As Integer = Val(LY)
        Dim LWN As Integer = Val(LW)

        Dim ICTITEM1 As String = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP, chkHISTCAT.Checked)

        Fill_Records("ICTITEM1", "", True, "Select * from " & ICTITEM1)


        ' Weeks

        ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYWW like '" & RY & "%'"
        Dim TY_WKmax As String = ASCDATA1.GetDataValue

        ASCMAIN1.sql = "Select Max (YYYYWW) from GLTPARM3 where YYYYWW like '" & LY & "%'"
        Dim LY_WKmax As String = ASCDATA1.GetDataValue

        WM(0, 0) = RY & Format(RWN - REL_WEEK + 1, "00")
        WM(0, 1) = RYW
        WM(0, 2) = RY & Format(RWN - REL_WEEK + MAX_WEEK, "00")

        WM(1, 0) = Mid(LYW, 1, 4) & Format(LWN - REL_WEEK_LY + 1, "00")
        WM(1, 1) = LYW
        WM(1, 2) = LY & Format(LWN - REL_WEEK_LY + MAX_WEEK_LY, "00")

        If CALENDAR = "O" Then
            If RWN >= 49 Or RWN < 23 Then
                If RWN >= 49 Then
                    WS(0, 0) = Mid(RYW, 1, 4) & "49"
                    WS(0, 2) = Format(Val(Mid(RYW, 1, 4)) + 1, "0000") & "22"
                Else
                    WS(0, 0) = Format(Val(Mid(RYW, 1, 4)) - 1, "0000") & "49"
                    WS(0, 2) = Mid(RYW, 1, 4) & "22"
                End If
            Else
                WS(0, 0) = Mid(RYW, 1, 4) & "23"
                WS(0, 2) = Mid(RYW, 1, 4) & "48"
            End If
            WS(0, 1) = RYW
        Else
            If RWN < 27 Then
                WS(0, 0) = Mid(RYW, 1, 4) & "01"
                WS(0, 2) = Mid(RYW, 1, 4) & "26"
            Else
                WS(0, 0) = Mid(RYW, 1, 4) & "27"
                WS(0, 2) = Mid(RYW, 1, 4) & "53"
            End If
            WS(0, 1) = RYW
        End If

        WS(1, 0) = ASCMAIN1.Week_Calc(WS(0, 0), -52)
        WS(1, 1) = ASCMAIN1.Week_Calc(WS(0, 1), -52)
        WS(1, 2) = ASCMAIN1.Week_Calc(WS(0, 2), -52)

        If ASCMAIN1.CLIENT = "INT" Then
            WS(1, 0) = Make_LY(WS(0, 0))
            WS(1, 1) = Make_LY(WS(0, 1))
            WS(1, 2) = Make_LY(WS(0, 2))
        End If

        ASCMAIN1.sql = "Select Max (OPS_YYYYWW) from EDT852T1 where EDI_CUST_BATCH_NO = 'LGI'"
        Dim RYW_LGI As String = ASCDATA1.GetDataValue
        Dim N As Integer = ASCMAIN1.Week_Diff(WM(0, 1), RYW_LGI)
        If N < 0 And N > -5 Then
            ' OK TO USE LGI WEEK
        Else
            RYW_LGI = ""
        End If

        If CALENDAR = "O" Then
            ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 " _
            & " where YYYYPP = '" & Mid(RYP, 1, 4) & "01' and REL_WEEK = 1"
            rowGLTPARM3 = ASCDATA1.GetDataRow
            WY(0, 0) = rowGLTPARM3.Item("YYYYWW")
            WY(0, 1) = RYW
            ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 " _
            & " where YYYYPP = '" & Mid(RYP, 1, 4) & "12' and REL_WEEK = MAX_WEEK"
            rowGLTPARM3 = ASCDATA1.GetDataRow
            WY(0, 2) = rowGLTPARM3.Item("YYYYWW")

            ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 " _
            & " where YYYYPP = '" & Mid(LYP, 1, 4) & "01' and REL_WEEK = 1"
            rowGLTPARM3 = ASCDATA1.GetDataRow
            WY(1, 0) = rowGLTPARM3.Item("YYYYWW")
            WY(1, 1) = LYW
            ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 " _
            & " where YYYYPP = '" & Mid(LYP, 1, 4) & "12' and REL_WEEK = MAX_WEEK"
            rowGLTPARM3 = ASCDATA1.GetDataRow
            WY(1, 2) = rowGLTPARM3.Item("YYYYWW")
        Else
            WY(0, 0) = RY & "01"
            WY(0, 1) = RYW
            WY(0, 2) = TY_WKmax

            WY(1, 0) = LY & "01"
            WY(1, 1) = LYW
            WY(1, 2) = LY_WKmax
        End If


        For W As Integer = 1 To 6
            If W > MAX_WEEK Then
                WW(0, W) = ""
                WW(1, W) = ""
            Else
                WW(0, W) = RY & Format(RWN - REL_WEEK + W, "00")
                WW(1, W) = ASCMAIN1.Week_Calc(WW(0, W), -52)

                If ASCMAIN1.CLIENT = "INT" Then
                    WW(1, W) = Make_LY(WW(0, W))
                End If

            End If
        Next

        'For i As Integer = 0 To 2
        '    'WM(1, i) = ASCMAIN1.Week_Calc(WM(0, i), -1 * 52)
        '    'WS(1, i) = ASCMAIN1.Week_Calc(WS(0, i), -1 * 52)
        '    'WY(1, i) = ASCMAIN1.Week_Calc(WY(0, i), -1 * 52)
        '    WM(1, i) = Format(Val(Mid(WM(0, i), 1, 4)) - 1, "0000") & Mid(WM(0, i), 5, 2)
        '    WS(1, i) = Format(Val(Mid(WS(0, i), 1, 4)) - 1, "0000") & Mid(WS(0, i), 5, 2)
        '    WY(1, i) = Format(Val(Mid(WY(0, i), 1, 4)) - 1, "0000") & Mid(WY(0, i), 5, 2)
        'Next



        ' Months

        'Dim MMM(1, 2) As String ' 1=LY 0 = TY, 0 = start, 1 = xTD, 2 = end
        'Dim MMS(1, 2) As String ' 1=LY 0 = TY, 0 = start, 1 = xTD, 2 = end
        'Dim MMY(1, 2) As String ' 1=LY 0 = TY, 0 = start, 1 = xTD, 2 = end

        MMM(0, 0) = RYP
        MMM(0, 1) = RYP
        MMM(0, 2) = RYP

        If CALENDAR = "O" Then
            MMY(0, 0) = Mid(RYP, 1, 4) & "01"
        Else
            If Mid(RYM, 5, 2) = "01" Then
                MMY(0, 0) = ASCMAIN1.Period_Calc(RYP, -11)
            Else
                MMY(0, 0) = ASCMAIN1.Period_Calc(RYP, 2 - Val(Mid(RYM, 5, 2)))
            End If
        End If
        MMY(0, 1) = RYP
        MMY(0, 2) = ASCMAIN1.Period_Calc(MMY(0, 0), 11)

        For M As Integer = 1 To 12
            MM(0, M) = ASCMAIN1.Period_Calc(MMY(0, 0), M - 1)
            MM(1, M) = ASCMAIN1.Period_Calc(MM(0, M), -12)
        Next

        If (CALENDAR = "O" And (Mid(RYM, 5, 2) >= "01" And Mid(RYM, 5, 2) <= "06")) _
        Or (CALENDAR = "R" And (Mid(RYM, 5, 2) >= "02" And Mid(RYM, 5, 2) <= "07")) Then
            MMS(0, 0) = MMY(0, 0)
        Else
            MMS(0, 0) = ASCMAIN1.Period_Calc(MMY(0, 0), 6)
        End If
        MMS(0, 1) = RYP
        MMS(0, 2) = ASCMAIN1.Period_Calc(MMS(0, 0), 5)


        For I As Int16 = 0 To 2
            MMM(1, I) = ASCMAIN1.Period_Calc(MMM(0, I), -12)
            MMS(1, I) = ASCMAIN1.Period_Calc(MMS(0, I), -12)
            MMY(1, I) = ASCMAIN1.Period_Calc(MMY(0, I), -12)
        Next


        YPCOMPMIN = LookUp("GLTPARM3", WY(1, 0)).Item("YYYYPP")
        YPCOMPMAX = LookUp("GLTPARM3", RYW).Item("YYYYPP")

        If optReportFormat.Value = "8" Then
            YPCOMPMAX = RYP
            If (optCALENDAR.Value = "R" And Mid(RYP, 5, 2) >= "02" And Mid(RYP, 5, 2) <= "07") _
            Or (optCALENDAR.Value = "O" And Mid(RYP, 5, 2) >= "01" And Mid(RYP, 5, 2) <= "06") Then
                'YPCOMPMIN = LookUp("GLTPARM3", WY(0, 0)).Item("YYYYPP")
                YPCOMPMIN = Mid(RYP, 1, 4) & IIf(optCALENDAR.Value = "R", "02", "01")
            Else
                'YPCOMPMIN = LookUp("GLTPARM3", WY(0, 0)).Item("YYYYPP")
                YPCOMPMIN = Mid(RYP, 1, 4) & IIf(optCALENDAR.Value = "R", "08", "07")
            End If
        End If

        sql_filter = " and RSTCOMP1.OPS_YYYYWW BETWEEN '" & WY(1, 0) & "' AND '" & RYW & "'"
        'sql_filter &= vbCrLf & " and " & COLUMN_NAME & " <> 0"

        Dim sql_Data As String = ""

        For Y As Integer = 0 To 1
            sql_Data &= "" _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW = '" & WM(Y, 1) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_WTD_S" & vbCrLf _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW BETWEEN '" & WM(Y, 0) & "' AND '" & WM(Y, 1) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_MTD_S" & vbCrLf _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW BETWEEN '" & WM(Y, 0) & "' AND '" & WM(Y, 2) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_MTL_S" & vbCrLf _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW BETWEEN '" & WS(Y, 0) & "' AND '" & WS(Y, 1) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_STD_S" & vbCrLf _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW BETWEEN '" & WS(Y, 0) & "' AND '" & WS(Y, 2) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_STL_S" & vbCrLf _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW BETWEEN '" & WY(Y, 0) & "' AND '" & WY(Y, 1) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_YTD_S" & vbCrLf _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW BETWEEN '" & WY(Y, 0) & "' AND '" & WY(Y, 2) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_YTL_S" & vbCrLf
        Next

        Dim YWZ As String = "'" & WM(0, 1) & "'"
        If RYW_LGI <> "" Then YWZ = "DECODE(EDT852T1.EDI_CUST_BATCH_NO,'LGI','" & RYW_LGI & "','" & WM(0, 1) & "')"

        sql_Data &= "" _
        & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW = " & YWZ & " THEN " & COLUMN_NAME_EOW & " ELSE 0 END) QTY_EOW" & vbCrLf

        For W As Integer = 1 To 6
            For Y As Integer = 0 To 1
                sql_Data &= "" _
                & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW = '" & WW(Y, W) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_W" & Format(W, "0") & vbCrLf
            Next
        Next

        For Y As Integer = 0 To 1
            sql_Data &= "" _
            & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYWW >= '" & ASCMAIN1.Week_Calc(WM(Y, 1), -7) & "' AND RSTCOMP1.OPS_YYYYWW <= '" & WM(Y, 1) & "' THEN " & COLUMN_NAME & " ELSE 0 END) / 8 " & IIf(Y = 0, "TY", "LY") & "_SLS_XWKS" & vbCrLf
        Next



        For Y As Integer = 0 To 1
            For M As Integer = 1 To 12
                sql_Data &= "" _
                & ", Sum (CASE WHEN RSTCOMP1.OPS_YYYYPP = '" & MM(Y, M) & "' THEN " & COLUMN_NAME & " ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_M" & Format(M, "00") & vbCrLf
            Next
        Next



        'If chkHISTCAT.Checked Then
        sql_TABLE_NAMEs = Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1 & " ICTITEM1")
        'End If

        Dim COMP As String = " and SATAUTH1.CUST_CODE = RSTCOMP1.CUST_CODE and SATAUTH1.CUST_STORE_NO = RSTCOMP1.CUST_STORE_NO and SATAUTH1.HC_CODE = ICTCOLL1.HC_CODE and OPS_YYYYPP_OPENED <= '" & YPCOMPMIN & "' and NVL(OPS_YYYYPP_CLOSED,'" & YPCOMPMAX & "') >= '" & YPCOMPMAX & "'"

        ' THIS DEFINITION IS FOR CURRENTLY AUTHORIZED STORES
        COMP = " and SATAUTH1.CUST_CODE = RSTCOMP1.CUST_CODE and SATAUTH1.CUST_STORE_NO = RSTCOMP1.CUST_STORE_NO and SATAUTH1.HC_CODE = ICTCOLL1.HC_CODE and SATAUTH1.OPS_YYYYPP_OPENED IS NOT NULL AND SATAUTH1.OPS_YYYYPP_CLOSED IS NULL"


        sql = "Select " & sql_SELECT_cols & vbCrLf & "" & vbCrLf & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " RSTCOMP1,ICTRETLA,EDT852T1 " & sql_TABLE_NAMEs & vbCrLf _
        & IIf(Absx1.chkFor("CHKCOMP").Checked, ",SATAUTH1", "") _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & vbCrLf _
        & " and ICTRETLA.ITEM_CODE (+) = RSTCOMP1.ITEM_CODE and ICTRETLA.OPS_YYYYPP (+) = '" & RYP & "'") & vbCrLf _
        & " and EDT852T1.EDI_DOC_SEQ_NO (+) = RSTCOMP1.EDI_DOC_SEQ_NO" & vbCrLf _
        & IIf(Absx1.chkFor("CHKCOMP").Checked, COMP, "") _
        & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended _
        & ", TY_WTD_S, TY_MTD_S, TY_MTL_S, TY_STD_S, TY_STL_S, TY_YTD_S, TY_YTL_S, LY_WTD_S, LY_MTD_S, LY_MTL_S, LY_STD_S, LY_STL_S, LY_YTD_S, LY_YTL_S" _
        & ", QTY_EOW, TY_W1, LY_W1, TY_W2, LY_W2, TY_W3, LY_W3, TY_W4, LY_W4, TY_W5, LY_W5, TY_W6, LY_W6, TY_SLS_XWKS, LY_SLS_XWKS" _
        & ", TY_M01, TY_M02, TY_M03, TY_M04, TY_M05, TY_M06, TY_M07, TY_M08, TY_M09, TY_M10, TY_M11, TY_M12" _
        & ", LY_M01, LY_M02, LY_M03, LY_M04, LY_M05, LY_M06, LY_M07, LY_M08, LY_M09, LY_M10, LY_M11, LY_M12)" & vbCrLf _
        & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        If optUS.Value = "SALES" Then
            ' DEAL WITH P01 BEING FEB
            If True Then ' "WE ARE DEALING WITH BUDGETS" = "YES" Then

                If ASCMAIN1.CLIENT = "AHA" Then ' Get all of this from Weekly Budgets'

                    Dim RSTBUDRW As String = ASCMAIN1.Temp_Table("Select * from RSTBUDRW")

                    If Absx1.chkFor("THOUSANDS").Checked Then
                        ASCMAIN1.sql = "Update " & RSTBUDRW & " Set BUDGET = BUDGET / 1000"
                        ASCDATA1.ExecuteSQL()
                    End If

                    MyBase.Get_SQL("B")

                    '' using data source information from RSTBUDR1 - trouble if ITEM_CATGY_CODE ever comes into play - easiest solution is to just add it to RSTBUDRW
                    'sql = "Select " & sql_SELECT_cols & vbCrLf & "" & vbCrLf & sql_Data _
                    '& " from " & RSTBUDRW & " RSTBUDR1 " & sql_TABLE_NAMEs & vbCrLf _
                    '& ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                    '& " group by " & sql_GROUP_BY_cols

                    sql_filter = ""

                    sql_Data = ""
                    For Y As Int16 = 0 To 1
                        sql_Data &= "" _
                        & ", Sum (CASE WHEN OPS_YYYYWW = '" & WY(Y, 1) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_WTD_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYWW >= '" & WM(Y, 0) & "' AND OPS_YYYYWW <= '" & WM(Y, 1) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_MTD_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYWW >= '" & WM(Y, 0) & "' AND OPS_YYYYWW <= '" & WM(Y, 2) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_MTL_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYWW >= '" & WS(Y, 0) & "' AND OPS_YYYYWW <= '" & WS(Y, 1) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_STD_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYWW >= '" & WS(Y, 0) & "' AND OPS_YYYYWW <= '" & WS(Y, 2) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_STL_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYWW >= '" & WY(Y, 0) & "' AND OPS_YYYYWW <= '" & WY(Y, 1) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_YTD_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYWW >= '" & WY(Y, 0) & "' AND OPS_YYYYWW <= '" & WY(Y, 2) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_YTL_B" & vbCrLf
                    Next

                    For M As Integer = 1 To 12
                        ASCMAIN1.sql = "Select MIN (YYYYWW) YW1, MAX (YYYYWW) YW2 from GLTPARM3 where YYYYMM = '" & MM(0, M) & "'"
                        Dim row As DataRow = ASCDATA1.GetDataRow
                        Dim YW1 As String = row.Item("YW1")
                        Dim YW2 As String = row.Item("YW2")
                        sql_Data &= "" _
                        & ", Sum (CASE WHEN OPS_YYYYWW >= '" & YW1 & "' AND OPS_YYYYWW <= '" & YW2 & "' THEN BUDGET ELSE 0 END) " & "TY_B" & Format(M, "00") & vbCrLf
                    Next

                    sql = "Select " & sql_SELECT_cols & vbCrLf & "" & vbCrLf & sql_Data _
                    & " from " & RSTBUDRW & " RSTBUDR1 " & sql_TABLE_NAMEs & vbCrLf _
                    & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                    & " group by " & sql_GROUP_BY_cols

                    ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                    & "(" & G1thru9 & COLUMN_NAMEs_appended _
                    & ", TY_WTD_B, TY_MTD_B, TY_MTL_B, TY_STD_B, TY_STL_B, TY_YTD_B, TY_YTL_B, LY_WTD_B, LY_MTD_B, LY_MTL_B, LY_STD_B, LY_STL_B, LY_YTD_B, LY_YTL_B" _
                    & ", TY_B01, TY_B02, TY_B03, TY_B04, TY_B05, TY_B06, TY_B07, TY_B08, TY_B09, TY_B10, TY_B11, TY_B12" _
                    & ")" & vbCrLf _
                    & "(" & sql & ")"
                    ASCDATA1.ExecuteSQL()

                Else

                    Dim RSTBUDR1 As String = TAC.RSCMAIN1.RSTBUDR1_as_YP()

                    If Absx1.chkFor("THOUSANDS").Checked Then
                        ASCMAIN1.sql = "Update " & RSTBUDR1 & " Set BUDGET = BUDGET / 1000"
                        ASCDATA1.ExecuteSQL()
                    End If

                    MyBase.Get_SQL("B")

                    sql_filter = ""

                    sql_Data = ""
                    For Y As Int16 = 0 To 1
                        sql_Data &= "" _
                        & ", 0 " & IIf(Y = 0, "TY", "LY") & "_WTD_B" _
                        & ", Sum (CASE WHEN OPS_YYYYPP = '" & MMM(Y, 0) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_MTD_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYPP = '" & MMM(Y, 0) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_MTL_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYPP >= '" & MMS(Y, 0) & "' AND OPS_YYYYPP <= '" & MMS(Y, 1) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_STD_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYPP >= '" & MMS(Y, 0) & "' AND OPS_YYYYPP <= '" & MMS(Y, 2) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_STL_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYPP >= '" & MMY(Y, 0) & "' AND OPS_YYYYPP <= '" & MMY(Y, 1) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_YTD_B" & vbCrLf _
                        & ", Sum (CASE WHEN OPS_YYYYPP >= '" & MMY(Y, 0) & "' AND OPS_YYYYPP <= '" & MMY(Y, 2) & "' THEN BUDGET ELSE 0 END) " & IIf(Y = 0, "TY", "LY") & "_YTL_B" & vbCrLf
                    Next

                    For M As Integer = 1 To 12
                        sql_Data &= "" _
                        & ", Sum (CASE WHEN OPS_YYYYPP = '" & MM(0, M) & "' THEN BUDGET ELSE 0 END) " & "TY_B" & Format(M, "00") & vbCrLf
                    Next

                    'If chkHISTCAT.Checked Then
                    '    sql_TABLE_NAMEs = Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1 & " ICTITEM1")
                    'End If

                    sql = "Select " & sql_SELECT_cols & vbCrLf & "" & vbCrLf & sql_Data _
                    & " from " & RSTBUDR1 & " RSTBUDR1 " & sql_TABLE_NAMEs & vbCrLf _
                    & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                    & " group by " & sql_GROUP_BY_cols

                    ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                    & "(" & G1thru9 & COLUMN_NAMEs_appended _
                    & ", TY_WTD_B, TY_MTD_B, TY_MTL_B, TY_STD_B, TY_STL_B, TY_YTD_B, TY_YTL_B, LY_WTD_B, LY_MTD_B, LY_MTL_B, LY_STD_B, LY_STL_B, LY_YTD_B, LY_YTL_B" _
                    & ", TY_B01, TY_B02, TY_B03, TY_B04, TY_B05, TY_B06, TY_B07, TY_B08, TY_B09, TY_B10, TY_B11, TY_B12" _
                    & ")" & vbCrLf _
                    & "(" & sql & ")"
                    ASCDATA1.ExecuteSQL()

                End If
 
            End If
        End If

        Dim G As Int16 = COLUMN_NAMEs.Count
        If COLUMN_NAMEs(G - 1) = "ITEM_CODE" Then
            ' the next 4 lines were remmed out - 06/19/13 wjz restored them for AHA to see retail price - some items were showing up with 0s - only Anna would want the historical retail
            ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set ITEM_RETAIL_PRICE = (Select ITEM_RETAIL_PRICE from ICTITEM1 where ITEM_CODE = G" & CStr(G) & ")"
            ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Update " & ASTSRPT1 & " Set ITEM_CODE = G" & CStr(G)
            'ASCDATA1.ExecuteSQL()

        End If

        Dim sqlx As String = ""
        For Each COLUMN_NAME In COLUMN_NAME_sum.Keys
            sqlx &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        Next
        ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlx))

        Dim O As Integer = 0
        If (optCALENDAR.Value = "R" And Mid(RYP, 5, 2) >= "02" And Mid(RYP, 5, 2) <= "07") _
        Or (optCALENDAR.Value = "O" And Mid(RYP, 5, 2) >= "01" And Mid(RYP, 5, 2) <= "06") Then
            O = 0
        Else
            O = 6
        End If
        Dim sqlU As String = ""
        For M As Integer = 1 To 6
            sqlU &= ", TY_M" & CStr(M) & " = TY_M" & Format(M + O, "00")
            sqlU &= ", LY_M" & CStr(M) & " = LY_M" & Format(M + O, "00")
            sqlU &= ", TY_B" & CStr(M) & " = TY_B" & Format(M + O, "00")
        Next
        ASCDATA1.ExecuteSQL("Update " & ASTSRPT1 & " Set " & Mid(sqlU, 2))

    End Sub

    Function Make_LY(YYYYXX As String)

        Return Format(Val(Mid(YYYYXX, 1, 4)) - 1, "0000") & Mid(YYYYXX, 5, 2)
    End Function

    Sub Load_Images_by_Item()
        Dim FOLDERNAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        If ASCMAIN1.Running_in_VS Then
            FOLDERNAME = "C:\Users\wjz\Desktop\Clients\JHI\Images\"
        End If
        Dim lvl_ITEM_CODE As Integer = -1
        For lvl As Integer = 1 To COLUMN_NAMEs.Count
            If COLUMN_NAMEs(lvl - 1) = "ITEM_CODE" Then '  COLUMN_NAME_by_Lvl(lvl) = "ITEM_CODE" Then
                lvl_ITEM_CODE = lvl
                Exit For
            End If
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), "ITEM_CODE").Rows
            Dim ITEM_CODE As String = row.Item(0) & ""
            Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
            If rowICTITEM1 IsNot Nothing Then
                Dim ITEM_PICTURE_FILENAME As String = rowICTITEM1.Item("ITEM_PICTURE_FILENAME") & ""
                If ASCMAIN1.Running_in_VS Then
                    ITEM_PICTURE_FILENAME = "AB3108.jpg"
                End If

                Dim filename As String = FOLDERNAME & ITEM_PICTURE_FILENAME

                If My.Computer.FileSystem.FileExists(filename) Then
                    Dim rowICTITEM1_image As DataRow = dst.Tables("ICTITEM1_image").NewRow
                    rowICTITEM1_image.Item("ITEM_CODE") = ITEM_CODE
                    rowICTITEM1_image.Item("ITEM_PICTURE_FILENAME") = rowICTITEM1.Item("ITEM_PICTURE_FILENAME")
                    rowICTITEM1_image.Item("ITEM_PICTURE") = ASCMAIN1.GetImageData(filename)
                    dst.Tables("ICTITEM1_image").Rows.Add(rowICTITEM1_image)
                End If
            End If
        Next
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        If COLUMN_NAMEs.Contains("ITEM_CODE") Then
            Load_Images_by_Item()
        End If

        If Not chkN.Checked Then
            Exit Sub
        End If

        Dim C As String = ""
        If COLUMN_NAMEs.Count > 1 Then
            C = Mid(G1thru9, 1, (COLUMN_NAMEs.Count - 1) * 3 - 1)
        End If

        Dim COLUMN_NAME_to_rank As String = cmbRANKBY.Value ' "TY_YTD_S"
        'Stop

        'With tblASTSRPT1.Columns
        '    Dim O As Integer = 0
        '    If Mid(RYP, 5, 2) >= "02" And Mid(RYP, 5, 2) <= "07" Then
        '        O = 0
        '    Else
        '        O = 6
        '    End If
        '    For M As Integer = 1 To 6
        '        .Add("TY_M" & CStr(M), GetType(System.Decimal), "TY_M" & Format(M + O, "00"))
        '        .Add("LY_M" & CStr(M), GetType(System.Decimal), "TY_M" & Format(M + O, "00"))
        '        .Add("TY_B" & CStr(M), GetType(System.Decimal), "TY_M" & Format(M + O, "00"))
        '    Next
        'End With

        For Each row As DataRow In ASCDATA1.SelectDistinct(tblASTSRPT1, Split(C, ",")).Rows
            Dim sqlw As String = ""
            For I As Integer = 1 To COLUMN_NAMEs.Count - 1
                sqlw &= " and G" & CStr(I) & " = '" & row.Item(I - 1) & "'"
            Next

            Dim AD As String = ""
            If optTB.Value = "T" Then AD = " DESC"

            Dim T() As Decimal
            ReDim T(COLUMN_NAME_sum.Count - 1)

            Dim RANKs As New Dictionary(Of String, Integer)
            'If row.Item(1) = "Collection:BABYJEWEL" Then Stop
            For Each Y As String In New String() {"TY", "LY"}
                Dim RANK As Integer = 0
                Dim COLUMN_NAME_to_rank_XY As String = Y & Mid(COLUMN_NAME_to_rank, 3)
                For Each row2 As DataRow In tblASTSRPT1.Select(Mid(sqlw, 5), COLUMN_NAME_to_rank_XY & AD)
                    RANK += 1
                    If Val(row2.Item(COLUMN_NAME_to_rank_XY) & "") <> 0 Then row2.Item("RANK_" & Y) = RANK
                    If RANK > Val(numN.Value & "") Then
                        For CS As Integer = 0 To COLUMN_NAME_sum.Count - 1
                            Dim COLUMN_NAME As String = COLUMN_NAME_sum.Keys(CS)
                            T(CS) += Val(row2.Item(COLUMN_NAME) & "")
                        Next
                    End If
                Next
                RANKs(Y) = RANK
            Next

            If chkALLOTHERS.Checked And RANKs("TY") > Val(numN.Value & "") Then
                Dim row2 As DataRow = tblASTSRPT1.NewRow
                For CS As Integer = 0 To COLUMN_NAMEs.Count - 1 - 1
                    row2.Item("G" & CStr(CS + 1)) = row.Item(CS)
                Next
                row2.Item("G" & CStr(COLUMN_NAMEs.Count)) = COLUMN_CAPTIONs(COLUMN_NAMEs.Count - 1) & ":" & "*"
                For CS As Integer = 0 To COLUMN_NAME_sum.Count - 1
                    Dim COLUMN_NAME As String = COLUMN_NAME_sum.Keys(CS)
                    If COLUMN_NAME = "ITEM_RETAIL_PRICE" _
                    Or COLUMN_NAME = "RANK_TY" _
                    Or COLUMN_NAME = "RANK_LY" Then
                    Else
                        row2.Item(COLUMN_NAME) = T(CS)
                    End If
                Next
                row2.Item("RANK_TY") = DBNull.Value
                row2.Item("RANK_LY") = DBNull.Value
                tblASTSRPT1.Rows.Add(row2)
            End If
        Next

        Dim rowASTGROUP As DataRow = tblASTGROUP.NewRow
        rowASTGROUP.Item("GROUP_KEY") = COLUMN_CAPTIONs(COLUMN_NAMEs.Count - 1) & ":" & "*"
        rowASTGROUP.Item("GROUP_CODE") = "*"
        rowASTGROUP.Item("GROUP_DESC") = "All Others"
        tblASTGROUP.Rows.Add(rowASTGROUP)

        ASCDATA1.DeleteRows(tblASTSRPT1, "RANK_TY is null or RANK_TY > " & numN.Value)

        Dim Gx As String = "G" & CStr(COLUMN_NAMEs.Count)
        For Each row As DataRow In tblASTSRPT1.Select
            Dim RANK As Integer = Val(row.Item("RANK_TY") & "")
            If RANK = 0 Then
                RANK = Val(numN.Value & "") + 1
            End If
            Dim KEY_rank As String = Format(RANK, "000000" & "") & " " & row.Item(Gx)

            Dim rowASTGROUP2 As DataRow = tblASTGROUP.Rows.Find(KEY_rank)
            If rowASTGROUP2 Is Nothing Then
                rowASTGROUP = tblASTGROUP.Rows.Find(row.Item(Gx))
                rowASTGROUP2 = tblASTGROUP.NewRow
                rowASTGROUP2.Item("GROUP_KEY") = KEY_rank
                rowASTGROUP2.Item("GROUP_CODE") = rowASTGROUP.Item("GROUP_CODE")
                rowASTGROUP2.Item("GROUP_DESC") = rowASTGROUP.Item("GROUP_DESC")
                tblASTGROUP.Rows.Add(rowASTGROUP2)
            End If

            row.Item(Gx) = KEY_rank
        Next


        tblASTSRPT1.AcceptChanges()

    End Sub

    Public Overrides Sub Print_Report()

        If optReportFormat.Value = "0" Then
            ' NO REPORTS
        ElseIf optReportFormat.Value.ToString.StartsWith("*") Then
            ' NO REPORTS
        Else

            Mid(RPT, 8, 1) = optReportFormat.Value
            'CR_params.Add("YW", RYW)
            CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))
            CR_params.Add("OPTUS", Absx1.optFor("OPTUS").Value)
            If optReportFormat.Value = "1" Or optReportFormat.Value = "4" Or optReportFormat.Value = "6" Or optReportFormat.Value = "7" Then
                CR_params.Add("CHKTOPN", IIf(Absx1.chkFor("CHKTOPN").Checked, "1", "0"))
                CR_params.Add("NUMTOPN", Absx1.numFor("NUMN").Value)
            Else
                If optReportFormat.Value = "8" Then
                    If (CALENDAR = "O" And Mid(RYP, 5, 2) <= "06") _
                    Or (CALENDAR = "R" And Mid(RYP, 5, 2) >= "02" And Mid(RYP, 5, 2) <= "07") Then
                        CR_params.Add("M1", Mid(MM(0, 1), 5, 2))
                        CR_params.Add("M2", Mid(MM(0, 2), 5, 2))
                        CR_params.Add("M3", Mid(MM(0, 3), 5, 2))
                        CR_params.Add("M4", Mid(MM(0, 4), 5, 2))
                        CR_params.Add("M5", Mid(MM(0, 5), 5, 2))
                        CR_params.Add("M6", Mid(MM(0, 6), 5, 2))
                    Else
                        CR_params.Add("M1", Mid(MM(0, 7), 5, 2))
                        CR_params.Add("M2", Mid(MM(0, 8), 5, 2))
                        CR_params.Add("M3", Mid(MM(0, 9), 5, 2))
                        CR_params.Add("M4", Mid(MM(0, 10), 5, 2))
                        CR_params.Add("M5", Mid(MM(0, 11), 5, 2))
                        CR_params.Add("M6", Mid(MM(0, 12), 5, 2))
                    End If

                Else
                    CR_params.Add("W1", Mid(WW(0, 1), 5, 2))
                    CR_params.Add("W2", Mid(WW(0, 2), 5, 2))
                    CR_params.Add("W3", Mid(WW(0, 3), 5, 2))
                    CR_params.Add("W4", Mid(WW(0, 4), 5, 2))
                    CR_params.Add("W5", Mid(WW(0, 5), 5, 2))
                    CR_params.Add("W6", Mid(WW(0, 6), 5, 2))
                End If

                If optReportFormat.Value = "2" Or optReportFormat.Value = "5" Then
                    For w As Integer = 1 To 6
                        sql = "Select * from GLTPARM3 where YYYYWW = '" & WW(0, w) & "'"
                        Dim rowGLTPARM3 As DataRow = ASCDATA1.GetDataRow(sql)
                        If rowGLTPARM3 IsNot Nothing Then
                            CR_params.Add("W" & w & "_LEGEND", Format(rowGLTPARM3("WEEK_END_DATE"), "MM/dd"))
                        Else
                            CR_params.Add("W" & w & "_LEGEND", "")
                        End If
                    Next
                End If

                If optReportFormat.Value = "3" Then
                    CR_params.Add("W1_LEGEND", ASCMAIN1.Get_Legend_Wk(WW(0, 1), True))
                    CR_params.Add("W2_LEGEND", ASCMAIN1.Get_Legend_Wk(WW(0, 2), True))
                    CR_params.Add("W3_LEGEND", ASCMAIN1.Get_Legend_Wk(WW(0, 3), True))
                    CR_params.Add("W4_LEGEND", ASCMAIN1.Get_Legend_Wk(WW(0, 4), True))
                    CR_params.Add("W5_LEGEND", ASCMAIN1.Get_Legend_Wk(WW(0, 5), True))
                    CR_params.Add("W6_LEGEND", ASCMAIN1.Get_Legend_Wk(WW(0, 6), True))
                    CR_params.Add("MTD", Mid(ASCMAIN1.Get_Legend_Wk(WW(0, 1), True), 1, 3) & " MTD")
                    CR_params.Add("YTD", Mid(ASCMAIN1.Get_Legend_Wk(WW(0, 1), True), 1, 3) & " YTD")
                End If


                If optReportFormat.Value = "8" Then
                    If (CALENDAR = "O" And Mid(RYP, 5, 2) <= "06") _
                    Or (CALENDAR = "R" And Mid(RYP, 5, 2) >= "02" And Mid(RYP, 5, 2) <= "07") Then
                        CR_params.Add("STD", "Spring " & Mid(MMM(0, 1), 1, 4))
                        CR_params.Add("M1_LEGEND", ASCMAIN1.Get_Legend(MM(0, 1), True))
                        CR_params.Add("M2_LEGEND", ASCMAIN1.Get_Legend(MM(0, 2), True))
                        CR_params.Add("M3_LEGEND", ASCMAIN1.Get_Legend(MM(0, 3), True))
                        CR_params.Add("M4_LEGEND", ASCMAIN1.Get_Legend(MM(0, 4), True))
                        CR_params.Add("M5_LEGEND", ASCMAIN1.Get_Legend(MM(0, 5), True))
                        CR_params.Add("M6_LEGEND", ASCMAIN1.Get_Legend(MM(0, 6), True))
                    Else
                        CR_params.Add("STD", "Fall " & Mid(MMM(0, 1), 1, 4))
                        CR_params.Add("M1_LEGEND", ASCMAIN1.Get_Legend(MM(0, 7), True))
                        CR_params.Add("M2_LEGEND", ASCMAIN1.Get_Legend(MM(0, 8), True))
                        CR_params.Add("M3_LEGEND", ASCMAIN1.Get_Legend(MM(0, 9), True))
                        CR_params.Add("M4_LEGEND", ASCMAIN1.Get_Legend(MM(0, 10), True))
                        CR_params.Add("M5_LEGEND", ASCMAIN1.Get_Legend(MM(0, 11), True))
                        CR_params.Add("M6_LEGEND", ASCMAIN1.Get_Legend(MM(0, 12), True))
                    End If
                End If
            End If
            CR_params.Add("YW_LEGEND", ASCMAIN1.Get_Legend_Wk(RYW))

            'Dim SUBT As String = ""
            If CALENDAR = "O" Then
                If SUBT <> "" Then
                    SUBT &= " - "
                End If
                SUBT &= "Ops Year Ending " & ASCMAIN1.Get_Legend_Wk(WY(0, 2))
            End If

            Dim LBLTOPN As String = ""
            If chkN.Checked Then
                LBLTOPN = IIf(optTB.Value = "T", "Top ", "Bottom ") & numN.Value & " " & COLUMN_CAPTIONs(COLUMN_NAMEs.Count - 1) & "s, Ranked by " & cmbRANKBY.Text
            End If
            CR_params.Add("LBLTOPN", LBLTOPN)

            If Absx1.chkFor("CHKCOMP").Checked Then
                If SUBT <> "" Then SUBT &= "; "
                SUBT &= "Comp Stores " & YPCOMPMIN & "-" & YPCOMPMAX
            End If

            Generate_Report(RPT, , SUBT)
        End If

        If ASCMAIN1.CLIENT = "INT" Then
        Else
            Prepare_Data_Extracts()
        End If
    End Sub

    Sub Prepare_Data_Extracts()

        Dim tbl As DataTable = dst.Tables("ASTSRPT1").Copy

        'If ASCMAIN1.CLIENT = "AHA" Then
        '    Dim cc As Integer = 0
        '    For i As Integer = 1 To COLUMN_NAMEs.Count
        '        If COLUMN_NAMEs(i - 1) = "CUST_STORE_CLASS_CODE" Then
        '            cc = i
        '            Exit For
        '        End If
        '    Next
        '    If cc <> 0 And cc <> 1 Then
        '        For Each ROW As DataRow In tbl.Select("G1 = '" & aRC & "' and G" & CStr(cc) & " = '" & COLUMN_CAPTIONs(cc - 1) & ":" & "COMP" & "'")
        '            ROW.Item("G" & CStr(cc)) = COLUMN_CAPTIONs(cc - 1) & ":" & "STORES"
        '        Next

        '        For Each ROW As DataRow In tbl.Select("G1 = '" & aRC & "' and G" & CStr(cc) & " = '" & COLUMN_CAPTIONs(cc - 1) & ":" & "NEW" & "'")

        '            ROW.Item("G" & CStr(cc)) = COLUMN_CAPTIONs(cc - 1) & ":" & "STORES"
        '        Next
        '    End If
        'End If


        For iRow As Int64 = tbl.Rows.Count - 1 To 0 Step -1 '  Each row As DataRow In tbl.Select("")
            Dim row As DataRow = tbl.Rows(iRow)
            For i As Integer = 1 To tblASTDSQLA.Select("SEQUENCE IS NOT NULL", "SEQUENCE").Length
                Dim C As String = row.Item("G" & CStr(i))
                If C = aRC Then
                    row.Delete()
                    Exit For
                Else
                    row.Item("G" & CStr(i)) = Split(C, ":")(1)
                End If
            Next
        Next

        grdASTEXPT1.DataSource = tbl
        grdASTEXPT1.Text = "Comparative Retail Sales"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")

        For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("SEQUENCE IS NOT NULL", "SEQUENCE")
            Dim C As String = "G" & CStr(rowASTDSQLA.Item("SEQUENCE"))
            Dim D As String = rowASTDSQLA.Item("COLUMN_CAPTION")
            Set_DX_Column(grdASTEXPT1, C, D, 80)
        Next

        For Each rowASTDSQLS As DataRow In tblASTDSQLS.Select("", "COLUMN_SEQ")
            Dim C As String = rowASTDSQLS.Item("COLUMN_NAME")
            Dim D As String = rowASTDSQLS.Item("COLUMN_CAPTION")
            Set_DX_Column(grdASTEXPT1, C, D, 100, "#,##0", "Sum")
        Next
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYW").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Week"
            End If

            If chkN.Checked Then
                If numN.Value <= 0 Then
                    EMsg &= vbCr & "You must Specify a Value for N (Report will show only the Top N)"
                End If
                If cmbRANKBY.Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Rank-By Field"
                End If

                If tblASTDSQLA.Select("SEQUENCE IS NOT NULL").Length < 2 Then
                    EMsg &= vbCr & "You must Specify at Least 2 Sort-By fields when doing a TopN"
                End If
            End If

            If optReportFormat.Value = "8" And optUS.Value = "UNITS" Then
                EMsg &= vbCr & "This Format does NOT support Units"
            End If

            If ASCMAIN1.CLIENT = "INT" Then
                If optReportFormat.Value = "8" And Not Absx1.chkFor("THOUSANDS").Checked Then
                    EMsg &= vbCr & "This Format must be run in Thousands"
                End If
            End If

            If Absx1.chkFor("CHKCOMP").Checked Then
                If tblASTDSQLA.Select("SEQUENCE IS NOT NULL AND (COLUMN_NAME = 'BRAND_CODE' OR COLUMN_NAME = 'HC_CODE')").Length = 0 Then
                    EMsg &= vbCr & "You must Specify either Brand or High Collection in the Sort when doing a Comp Stores Report"
                End If
            End If

            If Me.chkCUST_CALENDAR.Checked Then
                Dim CUST_CODES As String = SQLA("CUST_CODE")
                If CUST_CODES = "" Or InStr(CUST_CODES, ",") <> 0 Then
                    EMsg &= vbCr & "You must Specify a Single Customer Code"
                End If
                Dim NOT_IN As String = SQLA("CUST_CODE", "EXCLUDE")
                If NOT_IN = "1" Then
                    EMsg &= vbCr & "You may NOT use the exclude option when using a Single Customer's Calendar"
                End If
            End If
        End If
    End Sub

    Private Sub optReportFormat_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optReportFormat.ValueChanged
        If optReportFormat.Value = "0" Or optReportFormat.Value.ToString.StartsWith("*") Then
            grpTBN.Visible = False
            chkN.Checked = False
        Else
            grpTBN.Visible = True
        End If
    End Sub

    Private Sub chkN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkN.CheckedChanged
        lblN.Visible = chkN.Checked
        numN.Visible = chkN.Checked
        optTB.Visible = chkN.Checked
        chkALLOTHERS.Visible = chkN.Checked
        lblRANKBY.Visible = chkN.Checked
        cmbRANKBY.Visible = chkN.Checked
    End Sub

    Private Sub optCALENDAR_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCALENDAR.ValueChanged
        If optCALENDAR.Value = "R" Then
            chkCUST_CALENDAR.Visible = True
        Else
            chkCUST_CALENDAR.Checked = False
            chkCUST_CALENDAR.Visible = False
        End If
    End Sub

    Private Sub optTB_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optTB.ValueChanged
        chkN.Text = "Show " & optTB.Text & "N Only"
    End Sub

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        ASCDATA1.ExecuteSQL("Alter Table " & TT & " Add ITEM_CODE VARCHAR2(25)")

        Dim G As Int16 = COLUMN_NAMEs.Count
        If COLUMN_NAMEs(G - 1) = "ITEM_CODE" Then

            ASCMAIN1.sql = "Update " & TT & " Set ITEM_CODE = SUBSTR(G" & CStr(G) & "," & CStr(Len(COLUMN_CAPTIONs(G - 1)) + 2) & ")"
            ASCDATA1.ExecuteSQL()
        End If
    End Sub

    Public Overrides Function Prepare_XLS_Summary_Columns(ByVal COLUMN_NAME_sum As Dictionary(Of String, String)) As String

        SUBT = ASCMAIN1.Get_Legend_Wk(RYW)


        If Not dst.Tables("ASTSRPT1").Columns.Contains("LAUNCH_DATE") Then
            With dst.Tables("ASTSRPT1")
                .Columns.Add("LAUNCH_DATE")
                .Columns.Add("WST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_WTD_S)=0,0,100*TY_WTD_S/(QTY_EOW+TY_WTD_S))")
                .Columns.Add("WWOH", GetType(System.Decimal), "IIF(TY_WTD_S=0,0,QTY_EOW/TY_WTD_S)")
                .Columns.Add("MTD_PCT", GetType(System.Decimal), "IIF(LY_MTD_S=0,0,100*(TY_MTD_S-LY_MTD_S)/LY_MTD_S)")
                .Columns.Add("MST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_MTD_S)=0,0,100*TY_MTD_S/(QTY_EOW+TY_MTD_S))")
                .Columns.Add("STD_PCT", GetType(System.Decimal), "IIF(LY_STD_S=0,0,100*(TY_STD_S-LY_STD_S)/LY_STD_S)")
                .Columns.Add("STL_PCT", GetType(System.Decimal), "IIF(LY_STL_S=0,0,100*TY_STD_S/LY_STL_S)")
                .Columns.Add("SST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_STD_S)=0,0,100*TY_STD_S/(QTY_EOW+TY_STD_S))")
                .Columns.Add("YTD_PCT", GetType(System.Decimal), "IIF(LY_YTD_S=0,0,100*(TY_YTD_S-LY_YTD_S)/LY_YTD_S)")
                .Columns.Add("YTL_PCT", GetType(System.Decimal), "IIF(LY_YTL_S=0,0,100*TY_YTD_S/LY_YTL_S)")
                .Columns.Add("YST_PCT", GetType(System.Decimal), "IIF((QTY_EOW+TY_YTD_S)=0,0,100*TY_YTD_S/(QTY_EOW+TY_YTD_S))")
            End With
        End If

        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Rows
            Dim ITEM_CODE As String = rowASTSRPT1.Item("ITEM_CODE") & ""
            Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
            If rowICTITEM1 IsNot Nothing Then
                rowASTSRPT1.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                rowASTSRPT1.Item("LAUNCH_DATE") = rowICTITEM1.Item("LAUNCH_DATE")
            End If

        Next
        '        Return "QTY_EOW,TY_WTD_S,LY_WTD_S,TY_MTD_S,LY_MTD_S,LY_MTL_S,TY_STD_S,LY_STD_S,LY_STL_S,TY_YTD_S,LY_YTD_S,LY_YTL_S"

        Return "LAUNCH_DATE,ITEM_RETAIL_PRICE,QTY_EOW,TY_WTD_S,LY_WTD_S,WST_PCT,WWOH,TY_MTD_S,LY_MTD_S,MTD_PCT,MST_PCT,TY_STD_S,LY_STD_S,STD_PCT,LY_STL_S,STL_PCT,SST_PCT,TY_YTD_S,LY_YTD_S,YTD_PCT,LY_YTL_S,YTL_PCT,YST_PCT"

    End Function

    Overrides Sub Prepare_XLS_Prepare_row(ByVal row As DataRow)
        Dim GMAX As Integer = COLUMN_NAMEs.Count

        If COLUMN_NAMEs(GMAX - 1) <> "ITEM_CODE" Then
            Exit Sub
        End If

        Dim ITEM_CODE As String = row.Item("ITEM_CODE")
        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE) ' LookUp("ICTITEM1", ITEM_CODE)

        row.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
        'row.Item("LAUNCH_DATE") = rowICTITEM1.Item("LAUNCH_DATE")
    End Sub

    Overrides Function Prepare_XLS_GetImage( _
    ByVal row As DataRow, _
    ByVal GMAX As Integer, _
    ByRef col As Integer) As String

        If COLUMN_NAMEs(GMAX - 1) <> "ITEM_CODE" Then
            Return ""
        End If

        Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
        Dim FOLDERNAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then FOLDERNAME = "C:\Documents and Settings\wjz\Desktop\Clients\JHI\Images\"
        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
        Dim IMGFILENAME As String = ""
        If rowICTITEM1 IsNot Nothing Then
            If rowICTITEM1.Item("ITEM_PICTURE_FILENAME") & "" <> "" Then
                IMGFILENAME = FOLDERNAME & rowICTITEM1.Item("ITEM_PICTURE_FILENAME") & ""
            End If
            row.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
            row.Item("LAUNCH_DATE") = rowICTITEM1.Item("LAUNCH_DATE")
        End If

        col = GMAX + 1
        Return IMGFILENAME
    End Function

    Public Overrides Sub Post_Process_Special()
        MyBase.Post_Process_Special()

        'If ASCMAIN1.CLIENT = "AHA" Then
        '    Dim c As Integer = 0
        '    For i As Integer = 1 To COLUMN_NAMEs.Count
        '        If COLUMN_NAMEs(i - 1) = "CUST_STORE_CLASS_CODE" Then
        '            c = i
        '            Exit For
        '        End If
        '    Next
        '    If c <> 0 And c <> 1 Then
        '        For Each ROW As DataRow In tbl.Select("G1 = '" & aRC & "' and G" & CStr(c) & " = '" & COLUMN_CAPTIONs(c - 1) & ":" & "COMP" & "'")
        '            ROW.Item("G" & CStr(c)) = COLUMN_CAPTIONs(c - 1) & ":" & "STORES"
        '        Next

        '        For Each ROW As DataRow In tbl.Select("G1 = '" & aRC & "' and G" & CStr(c) & " = '" & COLUMN_CAPTIONs(c - 1) & ":" & "NEW" & "'")

        '            ROW.Item("G" & CStr(c)) = COLUMN_CAPTIONs(c - 1) & ":" & "STORES"
        '        Next
        '    End If
        'End If

        If optReportFormat.Value.ToString.StartsWith("*") Then

            Prepare_Custom_XLS()

        ElseIf optReportFormat.Value = "0" Then

            'If optReportFormat.Value = "8" Then
            '    ' NO XLS
            'Else
            If optUS.Value = "UNITS" _
                And COLUMN_NAME_by_Lvl(1) = "CUST_CODE" _
                And COLUMN_NAME_by_Lvl(COLUMN_NAME_by_Lvl.Length - 1) = "ITEM_CODE" _
                And COLUMN_NAME_by_Lvl.Length - 1 = 4 Then

                Try
                    Prepare_XLS_Special()
                Catch ex As Exception
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
                End Try
            Else
                Try
                    Prepare_XLS()
                Catch ex As Exception
                    If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
                End Try
            End If
            'End If
        End If


    End Sub

    Sub Prepare_XLS_Special()

        ASCMAIN1.Progress("Now creating XLS Pivot Report")

        Dim tbl As New DataTable
        Dim PIVOT_COLUMN As Integer = 1
        Dim KEY() As DataColumn
        Dim keys() As String
        Dim keysG() As String
        ReDim KEY(COLUMN_CAPTION_by_Lvl.Count - 2 - 1)
        ReDim keys(COLUMN_CAPTION_by_Lvl.Count - 2 - 1)
        ReDim keysG(COLUMN_CAPTION_by_Lvl.Count - 2 - 1)
        Dim skeys As String = ""
        Dim C2 As Integer = 0

        Dim sort_data As String = ""
        sort_data = ",TY_STD_S_TOTALS DESC,EXT_STD_AMT_TOTALS DESC"

        For C As Integer = 1 To COLUMN_NAME_by_Lvl.Length - 1
            If C = PIVOT_COLUMN Then
                'DO NOTHING
            Else
                C2 += 1
                Dim COLUMN_NAME_key = COLUMN_NAME_by_Lvl(C)
                tbl.Columns.Add(COLUMN_NAME_key)
                KEY(C2 - 1) = tbl.Columns(COLUMN_NAME_key)
                keys(C2 - 1) = COLUMN_NAME_key
                keysG(C2 - 1) = "G" & CStr(C)
                If sort_data = "" Or C < COLUMN_NAME_by_Lvl.Length - 1 Then
                    skeys &= "," & COLUMN_NAME_key
                End If
            End If
        Next
        skeys = Mid(skeys, 2) & sort_data

        tbl.Columns.Add("DESC_VALUE")
        Dim OTHER_COLS() As String = {"ITEM_RETAIL_PRICE"}
        For Each COLUMN_NAME_OTHER As String In OTHER_COLS
            tbl.Columns.Add(COLUMN_NAME_OTHER, GetType(System.Decimal)) ' NEEDS TO BE MAPPED
        Next

        Dim PIVOT_INDEX As New Dictionary(Of String, Integer)
        Dim IDX As Integer = 0
        Dim COLUMN_NAME_PIVOT_KEY As String = ""
        For Each ROW As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), "G" & CStr(PIVOT_COLUMN)).Select("", "G" & CStr(PIVOT_COLUMN))
            IDX += 1
            COLUMN_NAME_PIVOT_KEY = Split(ROW.Item(0), ":")(1)
            PIVOT_INDEX.Add(COLUMN_NAME_PIVOT_KEY, IDX)
        Next
        IDX += 1
        COLUMN_NAME_PIVOT_KEY = "TOTALS"
        PIVOT_INDEX.Add(COLUMN_NAME_PIVOT_KEY, IDX)


        Dim PIVOT_VALUE_COLS As String() = {"TY_WTD_S", "TY_MTD_S", "EXT_MTD_AMT", "TY_STD_S", "TY_STD_SELLTHRU", "EXT_STD_AMT", "QTY_EOW", "EXT_QTY_EOW"}

        Dim PIVOT_VALUE_COL_EXPS As New Dictionary(Of String, String)
        Dim PIVOT_COL As String = "#"
        PIVOT_VALUE_COL_EXPS.Add("EXT_MTD_AMT", "TY_MTD_S" & "_" & PIVOT_COL & " * ITEM_RETAIL_PRICE")
        PIVOT_VALUE_COL_EXPS.Add("TY_STD_SELLTHRU", "IIF((ISNULL(TY_STD_S" & "_" & PIVOT_COL & ",0) + ISNULL(QTY_EOW" & "_" & PIVOT_COL & ",0)) = 0, 0, ISNULL(TY_STD_S" & "_" & PIVOT_COL & ",0) / (ISNULL(TY_STD_S" & "_" & PIVOT_COL & ",0) + ISNULL(QTY_EOW" & "_" & PIVOT_COL & ",0)))")
        PIVOT_VALUE_COL_EXPS.Add("EXT_STD_AMT", "TY_STD_S" & "_" & PIVOT_COL & " * ITEM_RETAIL_PRICE")
        PIVOT_VALUE_COL_EXPS.Add("EXT_QTY_EOW", "QTY_EOW" & "_" & PIVOT_COL & " * ITEM_RETAIL_PRICE")

        Dim PIVOT_VALUE_COL_FORMULAS() As String
        ReDim PIVOT_VALUE_COL_FORMULAS(PIVOT_VALUE_COLS.Length - 1)

        PIVOT_VALUE_COL_FORMULAS(4) = PIVOT_VALUE_COL_EXPS("TY_STD_SELLTHRU")
        Dim CF As String = PIVOT_VALUE_COL_FORMULAS(4)
        For i As Integer = 0 To PIVOT_VALUE_COLS.Length - 1
            Dim CN As String = PIVOT_VALUE_COLS(i)
            Dim CL As String = Chr(Asc("A") + i)
            CF = Replace(CF, PIVOT_VALUE_COLS(i) & "_" & PIVOT_COL, CL & PIVOT_COL)
            'IIF((ISNULL(TY_STD_S_#,0) + ISNULL(QTY_EOW_#,0)) = 0, 0, ISNULL(TY_STD_S_#,0) / (ISNULL(TY_STD_S_#,0) + ISNULL(QTY_EOW_#,0)))
        Next
        CF = Replace(CF, "IIF", "IF")
        CF = Replace(CF, "ISNULL(", "")
        CF = Replace(CF, ",0)", "")
        PIVOT_VALUE_COL_FORMULAS(4) = CF
        Dim PIVOT_TOTALS As New Dictionary(Of String, String)

        For Each PIVOT_COL In PIVOT_INDEX.Keys
            For Each COLUMN_NAME_data In PIVOT_VALUE_COLS
                Dim COLUMN_NAME As String = COLUMN_NAME_data & "_" & PIVOT_COL
                Dim DC As DataColumn = tbl.Columns.Add(COLUMN_NAME, GetType(System.Decimal))
                If PIVOT_COL = "TOTALS" Then
                    DC.Expression = Mid(PIVOT_TOTALS(COLUMN_NAME_data), 2)
                Else
                    If Not PIVOT_TOTALS.ContainsKey(COLUMN_NAME_data) Then
                        PIVOT_TOTALS.Add(COLUMN_NAME_data, "")
                    End If
                    PIVOT_TOTALS(COLUMN_NAME_data) &= "+ISNULL(" & COLUMN_NAME & ",0)"
                End If
            Next
            For Each COLUMN_NAME_data In PIVOT_VALUE_COL_EXPS.Keys
                Dim DC As DataColumn = tbl.Columns(COLUMN_NAME_data & "_" & PIVOT_COL)
                Dim COL_EXP As String = PIVOT_VALUE_COL_EXPS(COLUMN_NAME_data)
                COL_EXP = Replace(COL_EXP, "#", PIVOT_COL)
                DC.Expression = COL_EXP
            Next
        Next

        Dim keysg2() As String
        ReDim keysg2(keysG.Length + OTHER_COLS.Length - 1)
        For i As Integer = 0 To keysG.Length - 1
            keysg2(i) = keysG(i)
        Next
        If OTHER_COLS.Length <> 0 Then
            For i As Integer = 0 To OTHER_COLS.Length - 1
                keysg2(keysG.Length + i) = OTHER_COLS(i)
            Next
        End If

        tbl.PrimaryKey = KEY
        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("ASTSRPT1"), keysG).Rows
            Dim row2 As DataRow = tbl.NewRow
            Dim KEYVALUE() As String
            ReDim KEYVALUE(keysG.Length - 1)
            Dim SQLWHERE As String = ""
            For C As Integer = 1 To keysG.Length
                Dim CODE_VALUE As String = row.Item(keysG(C - 1))
                SQLWHERE &= " AND " & keysG(C - 1) & " = '" & CODE_VALUE & "'"
                CODE_VALUE = Split(CODE_VALUE, ":")(1)
                row2.Item(keys(C - 1)) = CODE_VALUE
                KEYVALUE(C - 1) = CODE_VALUE
            Next
            If tbl.Rows.Find(KEYVALUE) Is Nothing Then
                Dim ROWASTGROUP As DataRow = dst.Tables("ASTGROUP").Rows.Find(row.Item("G" & CStr(COLUMN_NAME_by_Lvl.Length - 1)))
                Dim DESC_VALUE As String = ""
                If ROWASTGROUP IsNot Nothing Then
                    DESC_VALUE = ROWASTGROUP.Item("GROUP_DESC")
                End If
                row2.Item("DESC_VALUE") = DESC_VALUE
                If OTHER_COLS.Length <> 0 Then
                    Dim rowo() As DataRow = dst.Tables("ASTSRPT1").Select(Mid(SQLWHERE, 5))
                    If rowo.Length <> 0 Then
                        For Each COLUMN_NAME_OTHER As String In OTHER_COLS
                            row2.Item(COLUMN_NAME_OTHER) = rowo(0).Item(COLUMN_NAME_OTHER)
                        Next
                    End If
                    Dim ITEM_CODE As String = rowo(0).Item("ITEM_CODE")
                    Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                    row2.Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                End If


                tbl.Rows.Add(row2)
            End If
        Next

        For Each ROW As DataRow In dst.Tables("ASTSRPT1").Select
            Dim CODE_VALUE_PIVOT As String = ROW.Item("G" & CStr(PIVOT_COLUMN))
            CODE_VALUE_PIVOT = Split(CODE_VALUE_PIVOT, ":")(1)
            'Dim IDX2 As Integer = PIVOT_INDEX(CODE_VALUE_PIVOT)
            Dim KEYVALUE() As String
            ReDim KEYVALUE(keysG.Length - 1)
            For C As Integer = 1 To keysG.Length
                Dim CODE_VALUE As String = ROW.Item(keysG(C - 1))
                CODE_VALUE = Split(CODE_VALUE, ":")(1)
                KEYVALUE(C - 1) = CODE_VALUE
            Next
            Dim row2 As DataRow = tbl.Rows.Find(KEYVALUE)
            For Each COLUMN_NAME As String In PIVOT_VALUE_COLS
                If PIVOT_VALUE_COL_EXPS.ContainsKey(COLUMN_NAME) Then
                Else
                    row2.Item(COLUMN_NAME & "_" & CODE_VALUE_PIVOT) = ROW.Item(COLUMN_NAME)
                End If
            Next
        Next


        Dim xlsFileName As String = ASCMAIN1.Folders("Work") & XNO

        Dim excel As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application
        Dim XWB As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Add
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim r As String

        Dim G_Colors(9) As System.Drawing.Color
        G_Colors(0) = Color.Gold
        G_Colors(1) = Color.Orange '   ' Color.Purple
        G_Colors(2) = Color.Yellow ' Color.Green
        G_Colors(3) = Color.LightGreen ' .DarkOrange
        G_Colors(4) = Color.LightBlue
        G_Colors(5) = Color.LavenderBlush
        G_Colors(6) = Color.LightPink
        G_Colors(7) = Color.LightCyan
        G_Colors(8) = Color.LightGoldenrodYellow
        G_Colors(9) = Color.LightSkyBlue
        'Dim G_Colors(9) As Int64
        'G_Colors(1) = 6750207
        'G_Colors(2) = 6750054
        'G_Colors(3) = 16764108
        'G_Colors(4) = 6406143
        'G_Colors(5) = 13434726
        'G_Colors(6) = 13421823
        'G_Colors(7) = 16764108
        'G_Colors(8) = 6406143
        'G_Colors(9) = 13421823

        Dim XWS As Microsoft.Office.Interop.Excel.Worksheet
        XWS = XWB.Worksheets(3)
        XWS.Delete()
        XWS = XWB.Worksheets(2)
        XWS.Delete()
        XWS = XWB.Worksheets(1)


        Dim XC As Integer = 0
        Dim XR As Integer = 0

        XR += 1 : XWS.Cells(XR, XC + 1).VALUE = "'" & Format(Now, "MM/dd/yyyy")
        XR += 0 : XWS.Cells(XR, XC + 2).VALUE = MENU_ITEM_OBJECT
        XR += 0 : XWS.Cells(XR, XC + 3).VALUE = "'" & XNO
        XR += 0 : XWS.Cells(XR, XC + 4).VALUE = MENU_ITEM_DESC
        XR += 1 : XWS.Cells(XR, XC + 1).VALUE = SUBT

        r = Me.Excel_Cell(XR, 1) & ":" & Me.Excel_Cell(XR, 1, )
        rng = XWS.Range(r)
        With rng
            .Font.Bold = True
            .Font.Size = 16
            .Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue)
        End With


        If RYWLEGEND <> "" Then
            XR += 1 : XWS.Cells(XR, XC + 1).VALUE = RYWLEGEND
        Else
            If RYPLEGEND <> "" Then
                XR += 1 : XWS.Cells(XR, XC + 1).VALUE = RYPLEGEND
            End If
        End If

        XR += 1

        ASCMAIN1.sql = "Select * from ASTDSQLS where FORM_NAME = '" & MENU_ITEM_OBJECT & "'"
        Dim tblASTDSQLS As DataTable = ASCDATA1.GetDataTable
        tblASTDSQLS.PrimaryKey = New DataColumn() {tblASTDSQLS.Columns("FORM_NAME"), tblASTDSQLS.Columns("COLUMN_NAME")}

        Dim JMIN As Integer = 0
        Dim JMAX As Integer = 0

        XR += 1

        For Each COLUMN_NAME As String In PIVOT_INDEX.Keys
            Dim IDX2 As Integer = PIVOT_INDEX(COLUMN_NAME)
            Dim XC3 As Integer = XC + 1 + keysg2.Length + (IDX2 - 1) * PIVOT_VALUE_COLS.Length + 1
            XWS.Cells(XR, XC3).VALUE = COLUMN_NAME

            If JMIN = 0 Then JMIN = XC3
            For J As Integer = 1 To PIVOT_VALUE_COLS.Length
                ' GET CAPTION FROM TABLE
                Dim COLUMN_NAME_PV As String = PIVOT_VALUE_COLS(J - 1)
                Dim COLUMN_CAPTION As String = ASCMAIN1.Make_Caption(COLUMN_NAME_PV)
                Dim rowASTDSQLS As DataRow = tblASTDSQLS.Rows.Find _
                    (New String() {MENU_ITEM_OBJECT, COLUMN_NAME_PV})
                If rowASTDSQLS IsNot Nothing Then
                    COLUMN_CAPTION = rowASTDSQLS.Item("COLUMN_CAPTION") & ""
                End If

                ' Customized Headings & Numeric Formats

                Dim FORMAT As String = "#,##0"
                Select Case COLUMN_NAME_PV
                    Case "TY_WTD_S"
                        COLUMN_CAPTION = "TY WTD" & vbCrLf & "Unit Sales"
                    Case "TY_MTD_S"
                        COLUMN_CAPTION = "TY MTD" & vbCrLf & "Unit Sales"
                    Case "EXT_MTD_AMT"
                        COLUMN_CAPTION = "TY MTD" & vbCrLf & "Retail Sales"
                        FORMAT = "$#,##0"
                    Case "TY_STD_S"
                        COLUMN_CAPTION = "TY STD" & vbCrLf & "Unit Sales"
                    Case "TY_STD_SELLTHRU"
                        COLUMN_CAPTION = "TY STD" & vbCrLf & "Sell-Thru%"
                        FORMAT = "0.0%"
                    Case "EXT_STD_AMT"
                        COLUMN_CAPTION = "TY STD" & vbCrLf & "Retail Sales"
                        FORMAT = "$#,##0"
                    Case "QTY_EOW"
                        COLUMN_CAPTION = "EOW On Hand" & vbCrLf & "Units"
                    Case "EXT_QTY_EOW"
                        COLUMN_CAPTION = "EOW On Hand" & vbCrLf & "@Retail"
                        FORMAT = "$#,##0"
                End Select

                XWS.Cells(XR + 1, XC3 + J - 1).VALUE = COLUMN_CAPTION
                JMAX = XC3 + J - 1

                'If J = 5 Then ' format as % - this needs to be parameterized
                r = Me.Excel_Cell(0, XC3 + J - 1) & ":" & Me.Excel_Cell(0, XC3 + J - 1)
                rng = XWS.Range(r)
                rng.EntireColumn.NumberFormat = FORMAT
                'End If
            Next

            'r = Split(Me.XC(JMIN), ":")(0) & ":" & Split(Me.XC(JMAX), ":")(0)
            r = Me.Excel_Cell(XR + 1, XC3) & ":" & Me.Excel_Cell(XR + 1, XC3 + PIVOT_VALUE_COLS.Length - 1)
            rng = XWS.Range(r)
            rng.EntireColumn.ColumnWidth = 13
            rng.EntireRow.RowHeight = rng.EntireRow.RowHeight * 0.9
            rng.EntireRow.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight

            r = Me.Excel_Cell(XR, XC3) & ":" & Me.Excel_Cell(XR, XC3 + PIVOT_VALUE_COLS.Length - 1)
            rng = XWS.Range(r)
            rng.Cells.Merge()
            rng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
        Next

        r = Me.Excel_Cell(XR, 1) & ":" & Me.Excel_Cell(XR + 1, JMAX)
        rng = XWS.Range(r)
        With rng
            .Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
            .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
            .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
            .Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Lavender) '.LightSkyBlue)
            .Interior.PatternTintAndShade = 0
        End With

        XR += 1
        Dim xC2 As Integer = 0
        For I As Integer = 1 To COLUMN_CAPTION_by_Lvl.Length - 1
            If I <> PIVOT_COLUMN Then
                xC2 += 1
                XWS.Cells(XR, XC + xC2).VALUE = COLUMN_CAPTION_by_Lvl(I)
            End If
        Next
        rng = XWS.Range(Split(Me.Excel_Cell(0, XC + 1), ":")(0) & ":" & Split(Me.Excel_Cell(0, XC + xC2), ":")(0))
        rng.EntireColumn.ColumnWidth = 15


        XWS.Cells(XR, XC + xC2 + 1).VALUE = "Description"
        rng = XWS.Range(Split(Me.Excel_Cell(0, XC + xC2 + 1), ":")(0) & ":" & Split(Me.Excel_Cell(0, XC + xC2 + 1), ":")(0))
        rng.EntireColumn.ColumnWidth = 90


        If OTHER_COLS.Length <> 0 Then
            For i As Integer = 1 To OTHER_COLS.Length
                Dim COLUMN_CAPTION As String = ASCMAIN1.Make_Caption(OTHER_COLS(i - 1))

                If OTHER_COLS(i - 1) = "ITEM_RETAIL_PRICE" Then
                    r = Me.Excel_Cell(0, XC + xC2 + 1 + i) & ":" & Me.Excel_Cell(0, XC + xC2 + 1 + i)
                    rng = XWS.Range(r)
                    rng.EntireColumn.NumberFormat = "$#,##0"

                    COLUMN_CAPTION = "Retail Price"
                End If

                XWS.Cells(XR, XC + xC2 + 1 + i).VALUE = COLUMN_CAPTION

            Next
        End If

        ' LOAD DATA

        Dim CURRENT_VALUE() As String
        ReDim CURRENT_VALUE(keysG.Length)

        XR += 1

        Dim XRS As Integer = XR + 1
        Dim LAST_LEVEL As Integer = -1
        Dim THIS_LEVEL As Integer = 0

        Dim STL As New Dictionary(Of Integer, List(Of Integer))
        For i As Integer = 0 To keysG.Length
            Dim X As New List(Of Integer)
            STL.Add(i, X)
        Next

        Dim recaps() As Dictionary(Of String, List(Of Integer))
        ReDim recaps(keysG.Length - 1)
        For i As Integer = 1 To keysG.Length - 1
            recaps(i) = New Dictionary(Of String, List(Of Integer))
        Next

        Dim sFilter As String = "TY_STD_S_TOTALS <> 0 OR QTY_EOW_TOTALS <> 0"

        For Each ROW As DataRow In tbl.Select(sFilter, skeys)
            THIS_LEVEL = keysG.Length
            For i As Integer = 1 To keysG.Length - 1
                If ROW.Item(i - 1) & "" <> CURRENT_VALUE(i) Then
                    THIS_LEVEL = i
                    Exit For
                End If
            Next

            If THIS_LEVEL <> keysG.Length Then
                SubTotal(THIS_LEVEL, LAST_LEVEL, G_Colors, XR, XC, JMIN, JMAX, XWS, keysG, _
                         CURRENT_VALUE, STL, recaps, PIVOT_VALUE_COL_FORMULAS)
            End If

            XR += 1
            For i As Integer = THIS_LEVEL To keysG.Length
                XWS.Cells(XR, XC + i).VALUE = ROW.Item(i - 1)
            Next


            r = Me.Excel_Cell(XR, XC + 1) & ":" & Me.Excel_Cell(XR, XC + keysG.Length - 1)
            rng = XWS.Range(r)
            With rng
                .Font.Bold = True
                '.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue)
            End With



            For i As Integer = keysG.Length + 1 To tbl.Columns.Count
                XWS.Cells(XR, XC + i).VALUE = ROW.Item(i - 1)
            Next
            STL(keysG.Length).Add(XR)

            If LAST_LEVEL = -1 Then
                LAST_LEVEL = keysG.Length
            Else
                LAST_LEVEL = THIS_LEVEL
            End If
            For i As Integer = 1 To keysG.Length - 1
                CURRENT_VALUE(i) = ROW.Item(i - 1) & ""
            Next
        Next

        Dim XRMAX As Integer = XR


        THIS_LEVEL = 0
        SubTotal(THIS_LEVEL, LAST_LEVEL, G_Colors, XR, XC, JMIN, JMAX, XWS, keysG, _
                 CURRENT_VALUE, STL, recaps, PIVOT_VALUE_COL_FORMULAS)

        Dim XRMAX2 As Integer = XR

        ' Print Recaps

        If keysG.Length - 1 >= 2 Then
            XR += 1

            For ST As Integer = 2 To keysG.Length - 1
                XR += 1
                XWS.Cells(XR, XC + 1).VALUE = "Recap"
                Dim RC As Dictionary(Of String, List(Of Integer)) = recaps(ST)
                For Each RC_VALUE As String In RC.Keys
                    CURRENT_VALUE(ST) = RC_VALUE
                    STL(ST + 1) = RC(RC_VALUE)

                    ' TAKEN FROM SUBTOTALS ROUTINE - CHANGES REFLECTED BY REMARKED CODE
                    XR += 1
                    If ST = 0 Then
                        XWS.Cells(XR, XC + 1).VALUE = "Totals"
                    Else
                        'For i As Integer = ST To keysG.Length - 1

                        XWS.Cells(XR, XC + ST).VALUE = CURRENT_VALUE(ST)

                        'Next
                    End If
                    Dim STFX As String = ""
                    For Each II As Integer In STL(ST + 1)
                        STFX &= "," & ":" & CStr(II)
                    Next

                    If STFX <> "" Then
                        Dim STF As String = "=@SUM(" & Mid(STFX, 2) & ")"
                        Dim pvi As Integer = 0
                        Dim pvc As Integer = 0
                        Dim pvx As Integer = 0
                        For J As Integer = JMIN To JMAX
                            pvx += 1
                            pvi = 1 + (pvx - 1) \ PIVOT_VALUE_COL_FORMULAS.Length
                            pvc = 1 + ((pvx - 1) Mod PIVOT_VALUE_COL_FORMULAS.Length)
                            Dim XCC As String = Split(Me.Excel_Cell(0, J), ":")(0)
                            If PIVOT_VALUE_COL_FORMULAS(pvc - 1) <> "" Then
                                Dim CFx As String = PIVOT_VALUE_COL_FORMULAS(pvc - 1)
                                For I As Integer = 1 To PIVOT_VALUE_COL_FORMULAS.Length
                                    'CF = Replace(CF, Chr(Asc("A") + I - 1) & "#", Chr(Asc("A") + I - 1 + (JMIN) + (pvi - 1) * PIVOT_VALUE_COL_FORMULAS.Length) & CStr(XR))
                                    Dim xcx As String = Split(Excel_Cell(0, I - 1 + (JMIN) + (pvi - 1) * PIVOT_VALUE_COL_FORMULAS.Length), ":")(0)
                                    CFx = Replace(CFx, Chr(Asc("A") + I - 1) & "#", xcx & CStr(XR))
                                Next
                                XWS.Cells(XR, J).VALUE = "=" & CFx
                            Else
                                XWS.Cells(XR, J).VALUE = Replace(STF, ":", XCC)
                            End If
                        Next

                        'r = Me.Excel_Cell(XR, ST) & ":" & Me.Excel_Cell(XR, JMAX)
                        'With XWS.Range(r)
                        '    .Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        '    .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                        '    .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                        '    .Interior.Color = System.Drawing.ColorTranslator.ToOle(G_Colors(ST))
                        '    .Interior.PatternTintAndShade = 0
                        'End With
                    End If
                Next
            Next
        End If

        Dim XRMAX3 As Integer = XR

        ' boxes around pivot entities

        For Each COLUMN_NAME As String In PIVOT_INDEX.Keys
            Dim IDX2 As Integer = PIVOT_INDEX(COLUMN_NAME)
            Dim XC3 As Integer = XC + 1 + keysg2.Length + (IDX2 - 1) * PIVOT_VALUE_COLS.Length + 1
            'XWS.Cells(XR, XC3).VALUE = COLUMN_NAME

            r = Me.Excel_Cell(XRS - 3, XC3) & ":" & Me.Excel_Cell(XRS - 3, XC3 + PIVOT_VALUE_COLS.Length - 1)
            rng = XWS.Range(r)
            'rng.EntireColumn.ColumnWidth = 13
            'rng.EntireRow.RowHeight = rng.EntireRow.RowHeight * 1.2
            'rng.EntireRow.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter


            With rng
                .Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                .Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Orange) ' .LavenderBlush)
                .Interior.PatternTintAndShade = 0
                .Font.Bold = True
                .Font.Size = 14
                .Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue)
            End With


            r = Me.Excel_Cell(XRS - 3, XC3) & ":" & Me.Excel_Cell(XRMAX3, XC3 + PIVOT_VALUE_COLS.Length - 1)
            rng = XWS.Range(r)
            rng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter

            With rng
                .BorderAround()
            End With

        Next


        Dim XCI As Integer = 3 ' position of ITEM_CODE
        If XCI <> 0 Then

            Dim FOLDERNAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
            XWS.Columns(3).insert()
            rng = XWS.Range("C:C")
            rng.EntireColumn.ColumnWidth = 1.1

            XWS.Cells(XRS - 2, XCI).VALUE = "Visual"
            For I As Integer = XRS To XRMAX
                If XWS.Cells(I, XCI + 1).VALUE & "" <> "" Then
                    rng = XWS.Range(XWS.Cells(I, XCI), XWS.Cells(I, XCI))


                    Dim ITEM_CODE As String = XWS.Cells(I, XCI + 1).VALUE
                    If ASCMAIN1.Running_in_VS And ASCMAIN1.DBS_SERVER = "" Then FOLDERNAME = "C:\Documents and Settings\wjz\Desktop\JHI\Images\"
                    Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE) '  LookUp("ICTITEM1", ITEM_CODE)
                    Dim filename As String = FOLDERNAME & rowICTITEM1.Item("ITEM_PICTURE_FILENAME") & ""

                    If ASCMAIN1.Running_in_VS Then
                        filename = "C:\Users\wjz\Desktop\Clients\JHI\Images\AB200.jpg"
                    End If

                    InsertPictureInRange(filename, rng, XWS)
                    rng = XWS.Range(CStr(I) & ":" & CStr(I))
                    rng.EntireRow.RowHeight = rng.EntireRow.RowHeight * 2
                    'rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                    rng.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                End If
            Next
            rng = XWS.Range("C:C")
            rng.EntireColumn.ColumnWidth = 6
        End If


        'Application.PrintCommunication = True
        'ActiveSheet.PageSetup.PrintArea = ""
        'Application.PrintCommunication = False
        With XWS.PageSetup
            '.LeftHeader = ""
            '.CenterHeader = ""
            '.RightHeader = ""
            '.LeftFooter = ""
            '.CenterFooter = ""
            '.RightFooter = ""
            '.LeftMargin = Application.InchesToPoints(0.7)
            '.RightMargin = Application.InchesToPoints(0.7)
            '.TopMargin = Application.InchesToPoints(0.75)
            '.BottomMargin = Application.InchesToPoints(0.75)
            '.HeaderMargin = Application.InchesToPoints(0.3)
            '.FooterMargin = Application.InchesToPoints(0.3)
            '.PrintHeadings = False
            .PrintGridlines = True
            '.PrintComments = xlPrintNoComments
            '.PrintQuality = 600
            '.CenterHorizontally = False
            '.CenterVertically = False
            '.Orientation = xlPortrait
            '.Draft = False
            '.PaperSize = xlPaperLetter
            '.FirstPageNumber = xlAutomatic
            '.Order = xlDownThenOver
            '.BlackAndWhite = False
            '.Zoom = 100
            '.PrintErrors = xlPrintErrorsDisplayed
            '.OddAndEvenPagesHeaderFooter = False
            '.DifferentFirstPageHeaderFooter = False
            '.ScaleWithDocHeaderFooter = True
            '.AlignMarginsHeaderFooter = True
            '.EvenPage.LeftHeader.Text = ""
            '.EvenPage.CenterHeader.Text = ""
            '.EvenPage.RightHeader.Text = ""
            '.EvenPage.LeftFooter.Text = ""
            '.EvenPage.CenterFooter.Text = ""
            '.EvenPage.RightFooter.Text = ""
            '.FirstPage.LeftHeader.Text = ""
            '.FirstPage.CenterHeader.Text = ""
            '.FirstPage.RightHeader.Text = ""
            '.FirstPage.LeftFooter.Text = ""
            '.FirstPage.CenterFooter.Text = ""
            '.FirstPage.RightFooter.Text = ""
        End With
        'Application.PrintCommunication = True


        'If My.Computer.FileSystem.FileExists(xlsFileName) Then
        '    Try
        '        My.Computer.FileSystem.DeleteFile(xlsFileName)
        '    Catch ex As Exception

        '    End Try
        'End If

        'Try
        '    'myWorkbook.SaveXlsx(FILENAME)
        '    myWorkbook.SaveXls(xlsFileName)
        '    tryagain = -1
        'Catch ex As Exception
        '    tryagain += 1
        'End Try

        'Loop While tryagain >= 0 And tryagain < 10

        'myWorkbook.ClosePreservedXlsx()
        'myWorkbook = Nothing

        'Dim excel As New Process
        'excel.StartInfo.Arguments = """" + xlsFileName + """ /e"
        'excel.StartInfo.FileName = xlsFileName
        'excel.Start()

        excel.UserControl = True
        excel.Visible = True
        excel = Nothing

    End Sub

    Sub SubTotal(THIS_LEVEL As Integer, _
                 LAST_LEVEL As Integer, _
                 G_Colors() As System.Drawing.Color, _
                 ByRef XR As Integer, _
                 ByRef XC As Integer, _
                 ByRef JMIN As Integer, _
                 ByRef JMAX As Integer, _
                 XWS As Microsoft.Office.Interop.Excel.Worksheet, _
                 keysG() As String, _
                 CURRENT_VALUE() As String, _
                 STL As Dictionary(Of Integer, List(Of Integer)), _
                 recaps() As Dictionary(Of String, List(Of Integer)), _
                 PIVOT_VALUE_COL_FORMULAS() As String)

        'Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim r As String

        If LAST_LEVEL <> -1 Then
            For ST = keysG.Length - 1 To THIS_LEVEL Step -1

                Dim RC As Dictionary(Of String, List(Of Integer)) = recaps(ST)
                Dim RC_VALUE As String = CURRENT_VALUE(ST)
                If ST <> 0 Then
                    If Not RC.ContainsKey(RC_VALUE) Then
                        RC.Add(RC_VALUE, New List(Of Integer))
                    End If
                End If

                XR += 1
                If ST = 0 Then
                    XWS.Cells(XR, XC + 1).VALUE = "Grand Totals"
                Else
                    For i As Integer = ST To keysG.Length - 1
                        XWS.Cells(XR, XC + ST).VALUE = CURRENT_VALUE(ST) & " Totals"
                        ' why are we looping here - isnt this the totals line?
                    Next
                End If
                Dim STFX As String = ""
                For Each II As Integer In STL(ST + 1)
                    STFX &= "," & ":" & CStr(II)
                Next

                If STFX <> "" Then
                    Dim STF As String = "=@SUM(" & Mid(STFX, 2) & ")"
                    Dim pvi As Integer = 0
                    Dim pvc As Integer = 0
                    Dim pvx As Integer = 0
                    For J As Integer = JMIN To JMAX
                        pvx += 1
                        pvi = 1 + (pvx - 1) \ PIVOT_VALUE_COL_FORMULAS.Length
                        pvc = 1 + ((pvx - 1) Mod PIVOT_VALUE_COL_FORMULAS.Length)
                        Dim XCC As String = Split(Me.Excel_Cell(0, J), ":")(0)
                        If PIVOT_VALUE_COL_FORMULAS(pvc - 1) <> "" Then
                            Dim CFx As String = PIVOT_VALUE_COL_FORMULAS(pvc - 1)
                            For I As Integer = 1 To PIVOT_VALUE_COL_FORMULAS.Length
                                'CFx = Replace(CF, Chr(Asc("A") + I - 1) & "#", Chr(Asc("A") + I - 1 + (JMIN) + (pvi - 1) * PIVOT_VALUE_COL_FORMULAS.Length) & CStr(XR))
                                Dim xcx As String = Split(Excel_Cell(0, I - 1 + (JMIN) + (pvi - 1) * PIVOT_VALUE_COL_FORMULAS.Length), ":")(0)
                                CFx = Replace(CFx, Chr(Asc("A") + I - 1) & "#", xcx & CStr(XR))
                            Next

                            XWS.Cells(XR, J).VALUE = "=" & CFx
                        Else
                            XWS.Cells(XR, J).VALUE = Replace(STF, ":", XCC)
                        End If
                    Next
                    r = Me.Excel_Cell(XR, IIf(ST = 0, 1, ST)) & ":" & Me.Excel_Cell(XR, JMAX)
                    With XWS.Range(r)
                        '.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        '.Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                        '.Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                        '.Interior.TintAndShade = -0.14996795556505
                        '.Interior.PatternTintAndShade = 0

                        .Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        .Interior.PatternColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                        .Interior.ThemeColor = Microsoft.Office.Interop.Excel.XlThemeColor.xlThemeColorDark1
                        .Interior.Color = System.Drawing.ColorTranslator.ToOle(G_Colors(ST))
                        '.Interior.TintAndShade = -0.13 * ST
                        .Interior.PatternTintAndShade = 0
                        '.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue)
                        .Font.Bold = True
                    End With


                    r = Me.Excel_Cell(XR, 1) & ":" & Me.Excel_Cell(XR, JMAX)
                    With XWS.Range(r)
                        .BorderAround()
                    End With

                End If

                STL(ST + 1) = New List(Of Integer)
                If ST <> 0 Then STL(ST).Add(XR)

                If ST <> 0 Then
                    RC(RC_VALUE).Add(XR)
                End If
            Next
        End If

    End Sub

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


    Sub Prepare_Custom_XLS(Optional ByVal xls_where As String = "", Optional ByVal ASTSRPT1 As String = "ASTSRPT1")

        ' Declare SSG Objects

        Dim oWB As SpreadsheetGear.IWorkbook
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        'Dim rangeCopyFrom As SpreadsheetGear.IRange
        'Dim rangePaste_To As SpreadsheetGear.IRange

        ' Parameters

        Dim Start_Row As Integer = 5

        ' Save Workbook as FILENAME

        Dim FILENAME_TEMPLATE As String = ""
        Dim FILENAME_SOURCE As String = ASCMAIN1.Folders("SharedRoot") & "Templates\" & FILENAME_TEMPLATE
        Dim XLS_FILENAME As String = ""

        ASCMAIN1.Progress("Now Creating Custom XLS Workbook")
        If FILENAME_TEMPLATE = "" Then
            oWB = SpreadsheetGear.Factory.GetWorkbook()
            oSheet = oWB.Worksheets.Add
            oSheet.Name = "Data"
            XLS_FILENAME = ASCMAIN1.Folders("Work") & XNO & ".xlsx"
            oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Else
            Dim success As Boolean = False
            Dim XLS_NO As Integer = 0

            Do Until success
                Try
                    XLS_NO += 1
                    XLS_FILENAME = ASCMAIN1.Folders("Work") & "Comparative_Retail_Sales"
                    XLS_FILENAME &= "-" & Format(XLS_NO, "000") & ".xlsx"
                    FileCopy(FILENAME_SOURCE, XLS_FILENAME)
                    success = True

                Catch ex As Exception
                    ' Stop
                End Try
            Loop

            oWB = SpreadsheetGear.Factory.GetWorkbook(XLS_FILENAME)
            oSheet = oWB.Worksheets("Data")
        End If


        Dim XTD_colors() As SpreadsheetGear.Color = _
        {SpreadsheetGear.Colors.PaleTurquoise, _
         SpreadsheetGear.Colors.PaleGoldenrod, _
         SpreadsheetGear.Colors.PaleGreen, _
         SpreadsheetGear.Colors.Beige, _
         SpreadsheetGear.Colors.PaleTurquoise, _
         SpreadsheetGear.Colors.PaleGoldenrod, _
         SpreadsheetGear.Colors.PaleGreen, _
         SpreadsheetGear.Colors.Beige, _
         SpreadsheetGear.Colors.PaleTurquoise, _
         SpreadsheetGear.Colors.PaleGoldenrod, _
         SpreadsheetGear.Colors.PaleGreen, _
         SpreadsheetGear.Colors.Beige}

        Dim G_Colors(9) As SpreadsheetGear.Color
        G_Colors(1) = SpreadsheetGear.Colors.Purple
        G_Colors(2) = SpreadsheetGear.Colors.Green
        G_Colors(3) = SpreadsheetGear.Colors.DarkOrange
        G_Colors(4) = SpreadsheetGear.Colors.Blue
        G_Colors(5) = SpreadsheetGear.Colors.Olive
        G_Colors(6) = SpreadsheetGear.Colors.Brown
        G_Colors(7) = SpreadsheetGear.Colors.Gold
        G_Colors(8) = SpreadsheetGear.Colors.DarkMagenta
        G_Colors(9) = SpreadsheetGear.Colors.Red


        ' Worksheet Heading


        With oSheet.Cells(0, 0)
            ' .Value = Format(Now, "MM/dd/yyyy HH:mm")
            .HorizontalAlignment = SpreadsheetGear.HAlign.Left
            .NumberFormat = "mm/dd/yy;@"
            .Value = Now
        End With
        With oSheet.Cells(0, 1)
            .Value = MENU_ITEM_OBJECT
        End With
        With oSheet.Cells(0, 2)
            .Value = ASCMAIN1.USER_ID
        End With
        With oSheet.Cells(1, 0)
            ' .Font.Color = SpreadsheetGear.Colors.Blue
            ' .Font.Size = 20
            .Font.Name = "Times New Roman"
            ' .Name = "Verdana"
            .Value = MENU_ITEM_DESC
        End With
        With oSheet.Cells(2, 0)
            .Font.Color = SpreadsheetGear.Colors.Blue
            .Font.Size = 20
            .Value = SUBT
        End With

        With oSheet.Cells(3, 0)
            ' .Font.Color = SpreadsheetGear.Colors.Blue
            ' .Font.Size = 20
            .Value = "Sell-Thru " & optUS.Text
        End With

        ' Prepare to Traverse Dataset

        Dim COLs_List As New List(Of String)

        Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
        Dim RYM As String = rowGLTPARM3.Item("YYYYMM")





        If ASCMAIN1.CLIENT = "AHA" Then
            Dim cc As Integer = 0
            For i As Integer = 1 To COLUMN_NAMEs.Count
                If COLUMN_NAMEs(i - 1) = "CUST_STORE_CLASS_CODE" Then
                    cc = i
                    Exit For
                End If
            Next

            dst.Tables("ASTGROUP").Rows.Add(New String() {COLUMN_CAPTIONs(cc - 1) & ":" & "STORES", "STORES", "Stores"})

            If cc <> 0 And cc <> 1 Then
                For Each ROW As DataRow In dst.Tables("ASTSRPT1").Select("G1 = '" & aRC & "' and G" & CStr(cc) & " = '" & COLUMN_CAPTIONs(cc - 1) & ":" & "COMP" & "'")
                    ROW.Item("G" & CStr(cc)) = COLUMN_CAPTIONs(cc - 1) & ":" & "STORES"
                Next

                Dim sqld As String = "G1 = '" & aRC & "' and G" & CStr(cc) & " = '" & COLUMN_CAPTIONs(cc - 1) & ":" & "NEW" & "'"
                For Each ROW As DataRow In dst.Tables("ASTSRPT1").Select(sqld)
                    Dim SQLWG As String = ""
                    For GC As Integer = 1 To 9
                        If GC = cc Then
                            SQLWG &= " and G" & CStr(GC) & " = '" & COLUMN_CAPTIONs(cc - 1) & ":" & "STORES" & "'"
                        Else
                            SQLWG &= " and G" & CStr(GC) & " = '" & aRC & "'"
                        End If
                    Next
                    Dim rowSTORES() As DataRow = dst.Tables("ASTSRPT1").Select(Mid(SQLWG, 6))
                    If rowSTORES Is Nothing Then
                        ROW.Item("G" & CStr(cc)) = COLUMN_CAPTIONs(cc - 1) & ":" & "STORES"
                    Else
                        For GCI As Integer = 9 To dst.Tables("ASTSRPT1").Columns.Count - 1
                            If dst.Tables("ASTSRPT1").Columns(GCI).Expression = "" Then
                                rowSTORES(0).Item(GCI) = Val(rowSTORES(0).Item(GCI) & "") + Val(ROW.Item(GCI) & "")
                            End If
                        Next
                    End If
                Next
                ASCDATA1.DeleteRows("ASTSRPT1", sqld)
            End If
        End If



        If optReportFormat.Value = "*B" Then
            Dim YTD As String = ""
            For M As Integer = 1 To 12
                If RYP >= MM(0, M) Then YTD &= "+TY_M" & Format(M, "00")
            Next
            dst.Tables("ASTSRPT1").Columns.Add("TY_M00", GetType(System.Decimal), Mid(YTD, 2))
            dst.Tables("ASTSRPT1").Columns.Add("LY_M00", GetType(System.Decimal), Mid(Replace(YTD, "TY_M", "LY_M"), 2))
            dst.Tables("ASTSRPT1").Columns.Add("TY_B00", GetType(System.Decimal), Mid(Replace(YTD, "TY_M", "TY_B"), 2))

            For M As Integer = 0 To 12

                Dim SCN_TY As String = ""
                Dim SCN As String = ""
                For Each TLB As String In New String() {"TY_M", "LY_M", "TY_B"}
                    SCN = TLB & Format(M, "00")
                    COLs_List.Add(SCN)
                    If SCN.StartsWith("TY_M") Then
                        SCN_TY = SCN
                    Else
                        Dim SCN_VAR As String = SCN & "_VAR"
                        Dim SCN_VARPCT As String = SCN & "_VARPCT"
                        COLs_List.Add(SCN_VAR)
                        COLs_List.Add(SCN_VARPCT)
                        If dst.Tables("ASTSRPT1").Columns.Contains(SCN) Then
                            dst.Tables("ASTSRPT1").Columns.Add(SCN_VAR, GetType(System.Decimal), SCN_TY & "-" & SCN)
                            dst.Tables("ASTSRPT1").Columns.Add(SCN_VARPCT, GetType(System.Decimal), "IIF(" & SCN & "=0,0," & SCN_VAR & "/" & SCN & ")")
                        End If
                    End If
                Next
            Next
        Else
            For Each XTD As String In New String() {"WTD", "MTD", "STD", "YTD"}
                For Each COL As String In New String() {"TY_XTD_S", "LY_XTD_S", "TY_XTD_B", "LY_XTL_S", "TY_XTL_B"}
                    Dim SCN As String = Replace(Replace(COL, "XTD", XTD), "XTL", Mid(XTD, 1, 2) & "L")
                    COLs_List.Add(SCN)
                    If SCN.StartsWith("TY") And SCN.EndsWith("_S") Then
                    Else
                        Dim SCN_VAR As String = SCN & "_VAR"
                        Dim SCN_VARPCT As String = SCN & "_VARPCT"
                        COLs_List.Add(SCN_VAR)
                        COLs_List.Add(SCN_VARPCT)
                        If dst.Tables("ASTSRPT1").Columns.Contains(SCN) Then
                            dst.Tables("ASTSRPT1").Columns.Add(SCN_VAR, GetType(System.Decimal), "TY_" & XTD & "_S-" & SCN)
                            dst.Tables("ASTSRPT1").Columns.Add(SCN_VARPCT, GetType(System.Decimal), "IIF(" & SCN & "=0,0," & SCN_VAR & "/" & SCN & ")")
                        End If
                    End If
                Next
            Next
        End If

        Dim COLs() As String = COLs_List.ToArray

        Dim FS As New Dictionary(Of String, String)
        Dim XLC As New Dictionary(Of String, String)

        Dim C As Integer = 0
        Dim R As Integer = Start_Row
        Dim GMAX As Integer = COLUMN_NAMEs.Count

        C = GMAX + 1 + 1

        '' TRAVERSE COLS BACKWARDS TO TAKE CARE OF VARPCT AND VAR BEFORE GETTING TO COLUMN NAME
        'For ISCN As Integer = COLs.Length - 1 To 0 Step -1
        '    Dim SCN As String = COLs(ISCN)
        '    If dst.Tables("ASTSRPT1").Columns.Contains(SCN) Then
        '        If dst.Tables("ASTSRPT1").Columns(SCN).Expression <> "" Then
        '            Dim FORMULA As String = "=" & Replace(dst.Tables("ASTSRPT1").Columns(SCN).Expression, "IIF", "IF")
        '            FS.Add(SCN, FORMULA)
        '        End If
        '    End If
        'Next

        For Each SCN As String In COLs
            C += 1

            If dst.Tables("ASTSRPT1").Columns.Contains(SCN) Then
                If dst.Tables("ASTSRPT1").Columns(SCN).Expression <> "" Then
                    Dim FORMULA As String = "=" & Replace(dst.Tables("ASTSRPT1").Columns(SCN).Expression, "IIF", "IF")
                    FS.Add(SCN, FORMULA)
                End If
            End If

            Dim CP As Integer = (C - 1) \ 26
            Dim XL As String = Chr(64 + C - CP * 26)
            If CP > 0 Then
                XL = Chr(64 + CP) & XL
            End If
            XLC.Add(SCN, XL & "#")
        Next


        Dim XL1 As Integer = 0
        Dim XL2 As Integer = 0

        Dim GROUP_KEY As String = ""
        Dim rowASTGROUP As DataRow = Nothing
        Dim GROUP_DESC As String = ""


        Dim G() As String = Nothing
        Dim GK() As String = Nothing
        Dim B As Integer = 0
        Dim ST() As String = Nothing

        R += 1
        For C = 1 To GMAX
            With oSheet.Cells(R - 1, C - 1)
                .Value = COLUMN_CAPTION_by_Lvl(C)
                .EntireColumn.ColumnWidth = 10
            End With
        Next

        With oSheet.Cells(R - 1, C - 1)
            .Value = "Description"
            .ColumnWidth = 30
        End With

        C += 1
        oSheet.Cells(R - 2, 0, R - 1, GMAX + 1 + COLs.Length).Interior.Color = SpreadsheetGear.Colors.WhiteSmoke
        ' oSheet.Cells(R - 2, C - 1).Interior.Pattern = SpreadsheetGear.Pattern.Solid
        oSheet.Cells(R - 3, GMAX + 1, R - 1, GMAX + 1 + COLs.Length).HorizontalAlignment = SpreadsheetGear.HAlign.Center
        oSheet.Cells(R - 3, GMAX + 1, R - 1, GMAX + 1 + COLs.Length).VerticalAlignment = SpreadsheetGear.VAlign.Center


        Dim CW As Integer = 10
        Dim Start_Col As New Dictionary(Of String, Integer)

        Dim FMT As String = "#,##0"
        If optUS.Value = "SALES" Then FMT = "$" & FMT

        If optReportFormat.Value = "*B" Then
            For M As Integer = 0 To 12
                Dim M00 As String = "M" & Format(M, "00")
                Start_Col.Add(M00, C)
                C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "TY" : .EntireColumn.NumberFormat = FMT : .ColumnWidth = CW : End With
                oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()

                Dim LEGEND As String = "YTD"
                If M > 0 Then
                    Dim YP As String = MM(0, M)
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
                    LEGEND = rowGLTPARM2.Item("LEGEND")
                End If

                oSheet.Cells(R - 3, C - 1).Value = LEGEND
                If M > 0 Then oSheet.Cells(R - 3, C - 1).Interior.Color = XTD_colors(M - 1)
                oSheet.Cells(R - 3, C - 1, R - 3, C + 3 * 2 - 1).Merge()

                For Each C2 As String In New String() {"LY", "Plan"}
                    C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = C2 : .EntireColumn.NumberFormat = FMT : .ColumnWidth = CW : End With
                    oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()
                    C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "+/-" & IIf(optUS.Value = "UNITS", "#", "$") : .EntireColumn.NumberFormat = FMT : .ColumnWidth = CW : End With
                    C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "+/-%" : .EntireColumn.NumberFormat = "#,##0%" : .ColumnWidth = CW * 0.6 : End With
                    oSheet.Cells(R - 2, C - 2).Value = "TY vs " & C2
                    oSheet.Cells(R - 2, C - 2, R - 2, C - 1).Merge()
                Next
            Next
        Else
            Dim iXTD As Integer = -1
            For Each XTD As String In New String() {"WTD", "MTD", "STD", "YTD"}
                Start_Col.Add(XTD, C)
                iXTD += 1
                C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "TY" : .EntireColumn.NumberFormat = FMT : .ColumnWidth = CW : End With
                oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()

                Dim CAPTION As String = XTD
                If XTD = "WTD" Then
                    CAPTION = "Week " & RYWLEGEND
                ElseIf XTD = "MTD" Then
                    Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", RYM)
                    Dim LEGEND As String = rowGLTPARM2.Item("LEGEND")
                    CAPTION = "MTD " & Mid(LEGEND, 10, 6)
                ElseIf CAPTION = "STD" Then
                    If Mid(RYW, 5, 2) > "26" Then
                        CAPTION = "Fall " & Mid(RYM, 1, 4)
                    Else
                        CAPTION = "Spring " & Mid(RYM, 1, 4)
                    End If
                End If

                oSheet.Cells(R - 3, C - 1).Value = CAPTION
                oSheet.Cells(R - 3, C - 1).Interior.Color = XTD_colors(iXTD)
                oSheet.Cells(R - 3, C - 1, R - 3, C + 3 * 4 - 1).Merge()

                For Each C2 As String In New String() {"LY", "Plan", "LY-Ttl", "Plan Ttl"}
                    C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = C2 : .EntireColumn.NumberFormat = FMT : .ColumnWidth = CW : End With
                    oSheet.Cells(R - 2, C - 1, R - 1, C - 1).Merge()
                    C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "+/-" & IIf(optUS.Value = "UNITS", "#", "$") : .EntireColumn.NumberFormat = FMT : .ColumnWidth = CW : End With
                    C += 1 : With oSheet.Cells(R - 1, C - 1) : .Value = "+/-%" : .EntireColumn.NumberFormat = "#,##0%" : .ColumnWidth = CW * 0.6 : End With
                    oSheet.Cells(R - 2, C - 2).Value = "TY vs " & C2
                    oSheet.Cells(R - 2, C - 2, R - 2, C - 1).Merge()
                Next
            Next
        End If


        Dim GS As String = ""
        For I As Integer = 1 To GMAX
            GS &= "," & "G" & CStr(I)
        Next

        Dim sqlw As String = ""
        'For I As Integer = 1 To GMAX
        '    sqlw &= " and G" & CStr(I) & " <> '" & aRC & "'"
        'Next
        'sqlw = Mid(sqlw, 5)

        If xls_where <> "" Then
            If sqlw = "" Then
                sqlw = xls_where
            Else
                sqlw &= " and " & xls_where
            End If
        End If

        For Each row As DataRow In dst.Tables("ASTSRPT1").Select(sqlw, Mid(GS, 2))
            For I As Integer = 1 To GMAX
                If G Is Nothing OrElse GK(I) <> row.Item("G" & CStr(I)) & "" Then
                    B = I

                    If G Is Nothing Then
                        ' REPORT HEADING
                        ReDim G(GMAX)
                        ReDim GK(GMAX)
                        ReDim ST(GMAX)
                    Else
                        If B < GMAX Then
                            Prepare_Custom_XLS_SubTotals(B, R, GMAX, XL1, XL2, ST, G, GK, COLs, FS, XLC, G_Colors, oSheet)
                            XL1 = 0
                            XL2 = 0
                        End If
                    End If

                    For J As Integer = B To GMAX
                        GROUP_KEY = row.Item("G" & CStr(J)) & ""
                        rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)

                        GK(J) = GROUP_KEY
                        G(J) = rowASTGROUP.Item("GROUP_CODE")
                        GROUP_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                        R += 1 ' HEADING


                        oSheet.Cells(R - 1, GMAX).Value = GROUP_DESC
                        oSheet.Cells(R - 1, 0).EntireRow.OutlineLevel = GMAX
                        oSheet.Cells(R - 1, GMAX).IndentLevel = J - 1
                        If J <> GMAX Then
                            oSheet.Cells(R - 1, GMAX).Font.Color = G_Colors(J)
                        End If

                        For C = 1 To J
                            oSheet.Cells(R - 1, C - 1).Value = G(C)
                            If C <> GMAX Then
                                oSheet.Cells(R - 1, C - 1).Font.Color = G_Colors(C)
                            End If
                        Next

                        If J <> GMAX Then
                            oSheet.Cells(R - 1, 0).EntireRow.Font.Color = G_Colors(J)
                        End If
                    Next
                End If
            Next

            Prepare_XLS_Prepare_row(row)

            C = GMAX + 1 + 1

            For Each SCN As String In COLs
                C += 1
                If FS.ContainsKey(SCN) Then
                    Dim FORMULA As String = FS(SCN)

                    For ISCN As Integer = COLs.Length - 1 To 0 Step -1
                        Dim SCN2 As String = COLs(ISCN)
                        If InStr(FORMULA, SCN2) <> 0 Then
                            FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                        End If
                    Next
                    'For Each SCN2 As String In COLs
                    '    If InStr(FORMULA, SCN2) <> 0 Then
                    '        FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                    '    End If
                    'Next
                    FORMULA = Replace(FORMULA, "#", CStr(R))
                    oSheet.Cells(R - 1, C - 1).Formula = FORMULA

                Else
                    If dst.Tables("ASTSRPT1").Columns.Contains(SCN) Then
                        oSheet.Cells(R - 1, C - 1).Value = row.Item(SCN)
                    End If
                End If
            Next
            '    Next
            'Next

            Try
                oSheet.Cells(R - 1, 0).EntireRow.OutlineLevel = GMAX
                If XL1 = 0 Then XL1 = R
                XL2 = R

            Catch ex As Exception
                If ASCMAIN1.USER_ID = "wjz" Then MsgBox(ex.Message)
            End Try

        Next

        If ST Is Nothing Then Exit Sub

        Prepare_Custom_XLS_SubTotals(0, R, GMAX, XL1, XL2, ST, G, GK, COLs, FS, XLC, G_Colors, oSheet)

        If optReportFormat.Value = "*B" Then
            For M As Integer = 0 To 12
                Dim M00 As String = "M" & Format(M, "00")

                Prepare_Custom_XLS_Border(oSheet, Start_Row - 1, Start_Col(M00), R - 1, Start_Col(M00) + 3 * 2)
                Prepare_Custom_XLS_Border(oSheet, Start_Row - 2, Start_Col(M00), Start_Row - 1, Start_Col(M00) + 3 * 2)
                Dim COLX As Integer = Start_Col(M00)
                For Each C2 As String In New String() {"TY", "LY", "BUD"}
                    Dim COLXW As Integer = 1
                    If C2 <> "TY" Then
                        COLXW += 2
                    End If
                    Prepare_Custom_XLS_Border(oSheet, Start_Row - 1, COLX, R - 1, COLX + COLXW)

                    Prepare_Custom_XLS_Border(oSheet, Start_Row - 2, COLX, Start_Row - 1, COLX + COLXW)
                    Prepare_Custom_XLS_Border(oSheet, Start_Row - 1, COLX, Start_Row, COLX + COLXW)
                    COLX += COLXW
                Next
            Next

        Else
            For Each XTD As String In New String() {"WTD", "MTD", "STD", "YTD"}
                Prepare_Custom_XLS_Border(oSheet, Start_Row - 1, Start_Col(XTD), R - 1, Start_Col(XTD) + 12)
                Prepare_Custom_XLS_Border(oSheet, Start_Row - 2, Start_Col(XTD), Start_Row - 1, Start_Col(XTD) + 12)
                Dim COLX As Integer = Start_Col(XTD)
                For Each C2 As String In New String() {"TY", "LY", "BUD", "LYTOT", "BUDTOT"}
                    Dim COLXW As Integer = 1
                    If C2 <> "TY" Then
                        COLXW += 2
                    End If
                    Prepare_Custom_XLS_Border(oSheet, Start_Row - 1, COLX, R - 1, COLX + COLXW)

                    Prepare_Custom_XLS_Border(oSheet, Start_Row - 2, COLX, Start_Row - 1, COLX + COLXW)
                    Prepare_Custom_XLS_Border(oSheet, Start_Row - 1, COLX, Start_Row, COLX + COLXW)
                    COLX += COLXW
                Next

                If XTD = "WTD" Then
                    For I As Integer = 7 To 12
                        oSheet.Cells(0, Start_Col(XTD) + I).EntireColumn.Hidden = True
                    Next
                End If
                If XTD = "MTD" Then
                    For I As Integer = 10 To 12
                        oSheet.Cells(0, Start_Col(XTD) + I).EntireColumn.Hidden = True
                    Next
                End If
            Next
        End If

        oSheet.WindowInfo.DisplayGridlines = False


        ' Save Document and Show

        oWB.Save()
        Show_Document(XLS_FILENAME)
        oWB = Nothing

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Prepare_Custom_XLS_Border(oSheet As SpreadsheetGear.IWorksheet, R1 As Int64, C1 As Int64, R2 As Int64, C2 As Int64)
        With oSheet.Range(R1, C1, R2, C2)
            .Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
            .Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
            '.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
            '.Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
        End With
    End Sub

    Sub Prepare_Custom_XLS_SubTotals( _
    ByVal B As Integer, _
    ByRef R As Integer, _
    ByVal GMAX As Integer, _
    ByVal XL1 As Integer, _
    ByVal XL2 As Integer, _
    ByVal ST() As String, _
    ByVal G() As String, _
    ByVal GK() As String, _
    ByVal COLs() As String, _
    ByVal FS As Dictionary(Of String, String), _
    ByVal XLC As Dictionary(Of String, String), _
    ByVal G_Colors() As SpreadsheetGear.Color, _
    ByVal ws As SpreadsheetGear.IWorksheet)

        Dim C As Integer = 0

        Dim GROUP_KEY As String = ""
        Dim rowASTGROUP As DataRow = Nothing
        Dim GROUP_DESC As String = ""

        For Slvl As Integer = GMAX - 1 To B Step -1

            R += 1 ' SUB-TOTAL
            ws.Cells(R - 1, 0).EntireRow.Font.Color = G_Colors(Slvl)
            For J As Integer = Slvl To 1 Step -1
                ws.Cells(R - 1, J - 1).Value = G(J)
                ws.Cells(R - 1, J - 1).Font.Color = G_Colors(J)
            Next

            ST(Slvl) &= ",X" & CStr(R)

            If Slvl = 0 Then
                GROUP_DESC = "Totals"
            Else
                GROUP_KEY = GK(Slvl)
                rowASTGROUP = dst.Tables("ASTGROUP").Rows.Find(GROUP_KEY)
                GROUP_DESC = rowASTGROUP.Item("GROUP_DESC") & ""
                ws.Cells(R - 1, GMAX).IndentLevel = Slvl - 1
            End If
            ws.Cells(R - 1, GMAX).Value = GROUP_DESC
            ws.Cells(R - 1, GMAX).Font.Color = G_Colors(Slvl)

            C = GMAX + 1 + 1
            For Each SCN As String In COLs
                C += 1
                Dim CP As Integer = (C - 1) \ 26
                Dim XL As String = Chr(64 + C - CP * 26)
                If CP > 0 Then
                    XL = Chr(64 + CP) & XL
                End If

                If FS.ContainsKey(SCN) Then
                    Dim FORMULA As String = FS(SCN)
                    For ISCN As Integer = COLs.Length - 1 To 0 Step -1
                        Dim SCN2 As String = COLs(ISCN)
                        If InStr(FORMULA, SCN2) <> 0 Then
                            FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                        End If
                    Next
                    'For Each SCN2 As String In COLs
                    '    If InStr(FORMULA, SCN2) <> 0 Then
                    '        FORMULA = Replace(FORMULA, SCN2, XLC(SCN2))
                    '    End If
                    'Next
                    FORMULA = Replace(FORMULA, "#", CStr(R))
                    If GK(Slvl) = aRC Then
                    Else
                        ws.Cells(R - 1, C - 1).Formula = FORMULA
                    End If

                Else
                    If GK(Slvl) = aRC Then
                    Else
                        If Slvl = GMAX - 1 Then
                            ws.Cells(R - 1, C - 1).Formula = "=SUM(" & XL & XL1 & ":" & XL & XL2 & ")"
                        Else
                            ws.Cells(R - 1, C - 1).Formula = "=SUM(" & Replace(Mid(ST(Slvl + 1), 2), "X", XL) & ")"
                        End If
                    End If
                End If

                If Slvl > 0 Then ws.Cells(R - 1, 0).EntireRow.OutlineLevel = Slvl
                ws.Cells(R - 1, C - 1).Font.Color = G_Colors(Slvl)
            Next
            ST(Slvl + 1) = ""

            Dim CC As SpreadsheetGear.Color = SpreadsheetGear.Colors.PaleGoldenrod
            If Slvl = 0 Then CC = SpreadsheetGear.Colors.LightGray

            For C = 1 To GMAX + 1 + 1 + COLs.Length
                ws.Cells(R - 1, C - 1).Interior.Color = CC
                ws.Cells(R - 1, C - 1).Interior.Pattern = SpreadsheetGear.Pattern.Solid
            Next
            R += 1
            ws.Cells(R - 1, 0).EntireRow.RowHeight = ws.Cells(R - 1, 0).EntireRow.Height * 0.25
        Next

    End Sub


    Sub Sell_In_Sell_Thru()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing Sell-In / Sell-Thru DataSet")

        With tblSIST2.Columns
            '.Add("SELLIN_SELLTHRU", GetType(System.Decimal))
            '.Add("ONHAND", GetType(System.Decimal))
            '.Add("SELLTHRU", GetType(System.Decimal))
            .Add("SELLIN_SELLTHRU", GetType(System.Decimal))
            .Add("ONHAND", GetType(System.Int64))
            .Add("SELLTHRU", GetType(System.Int64))
            .Add("CUST_CODE")
        End With
        tblSIST2.PrimaryKey = New DataColumn() {tblSIST2.Columns("CUST_CODE")}

        ASCMAIN1.sql = "" _
            & "Select CUST_CODE, OPS_YYYYWW" & vbCrLf _
            & ", Sum (SELLIN) SELLIN, Sum (SELLTHRU) SELLTHRU, Sum (ONHAND) ONHAND from (" _
            & "Select CUST_CODE, OPS_YYYYWW" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP * ITEM_RETAIL_PRICE) SELLIN, 0 SELLTHRU, 0 ONHAND" & vbCrLf _
            & " from SOTINVH2" & vbCrLf _
            & " where ORDR_YYYYPP_UPDATED >= '201501'" & vbCrLf _
            & "   and CUST_CODE in" & vbCrLf _
            & "(Select Distinct CUST_CODE from RSTRETL1 where OPS_YYYYPP >= '201501')" & vbCrLf _
            & " group by CUST_CODE, OPS_YYYYWW" & vbCrLf _
            & " union " & vbCrLf _
            & "Select RSTRETL1.CUST_CODE, RSTRETL1.OPS_YYYYWW" & vbCrLf _
            & ", 0 SELLIN, SUM (RSTRETL1.AMT_SOLD) SELLTHRU, SUM (RSTRETL1.QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE / 1000) EOW" & vbCrLf _
            & " from RSTRETL1, ICTITEM1" & vbCrLf _
            & " where RSTRETL1.OPS_YYYYPP >= '201501'" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE" & vbCrLf _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.OPS_YYYYWW" & vbCrLf _
            & ") where OPS_YYYYWW is Not Null group by CUST_CODE, OPS_YYYYWW" & vbCrLf _
            & " order by CUST_CODE, OPS_YYYYWW"

        tblSIST = ASCDATA1.GetDataTable
        With tblSIST.Columns
            .Add("SELLIN_SELLTHRU", GetType(System.Decimal))
        End With
        tblSIST.Columns("ONHAND").ReadOnly = False
        tblSIST.PrimaryKey = New DataColumn() {tblSIST.Columns("CUST_CODE"), tblSIST.Columns("OPS_YYYYWW")}

        For Each row As DataRow In ASCDATA1.SelectDistinct(tblSIST, New String() {"CUST_CODE"}).Select("", "CUST_CODE")
            Dim CUST_CODE As String = row.Item(0)
            Dim SELLIN_CUM As Decimal = 0
            Dim SELLTHRU_CUM As Decimal = 0
            Dim ONHAND_LAST As Decimal = 0
            For Each rowC As DataRow In tblSIST.Select("CUST_CODE = '" & CUST_CODE & "'", "OPS_YYYYWW")
                If Val(rowC.Item("ONHAND") & "") = 0 Then

                    rowC.Item("ONHAND") = ONHAND_LAST
                End If
                SELLIN_CUM += Val(rowC.Item("SELLIN") & "")
                Dim SELLTHRU As Decimal = Val(rowC.Item("SELLTHRU") & "")
                SELLTHRU_CUM += SELLTHRU
                ONHAND_LAST = Val(rowC.Item("ONHAND") & "")

                Dim SELLIN_SELLTHRU As Decimal
                If SELLTHRU_CUM = 0 Then
                    SELLIN_SELLTHRU = 0
                Else
                    SELLIN_SELLTHRU = System.Math.Round(SELLIN_CUM / SELLTHRU_CUM, 6)
                End If
                rowC.Item("SELLIN_SELLTHRU") = SELLIN_SELLTHRU
            Next
        Next

        UltraChart1.Axis.X.NumericAxisType = Infragistics.UltraChart.Shared.Styles.NumericAxisType.Logarithmic
        UltraChart1.Axis.X.Extent = 30
        UltraChart1.Axis.X.LogBase = 10
        ' UltraChart1.TitleLeft.Text = "SIST%"
        ' UltraChart1.TitleBottom.Text = "Invty"

        '  UltraChart1.Axis.Y.NumericAxisType = Infragistics.UltraChart.Shared.Styles.NumericAxisType.Logarithmic
        UltraChart1.Axis.Y.Extent = 30

        UltraChart1.DataSource = tblSIST2
        UltraChart1.Data.DataBind()

        Dim P As Integer = 0
        Dim YWs As New List(Of String)
        For Each row As DataRow In ASCDATA1.SelectDistinct(tblSIST, New String() {"OPS_YYYYWW"}).Select("", "OPS_YYYYWW")
            Dim OPS_YYYYWW As String = row.Item(0) & ""
            If OPS_YYYYWW = "" Then
                Continue For
            End If
            YWs.Add(OPS_YYYYWW)
            If OPS_YYYYWW = "201552" Then P = YWs.Count
        Next

        Dim labelHash As New Hashtable()
        labelHash.Add("HIGHLOW", New MyCustomTooltip())
        UltraChart1.LabelHash = labelHash

        'UltraChart1.Axis.X.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"
        UltraChart1.Axis.Y.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"
        UltraChart1.Axis.Z.Labels.ItemFormatString = "<DATA_VALUE:#,##0>"
        UltraChart1.Tooltips.Format = Infragistics.UltraChart.Shared.Styles.TooltipStyle.Custom
        UltraChart1.Tooltips.FormatString = "<HIGHLOW>"

        UltraTrackBar1.MinValue = 1
        UltraTrackBar1.MaxValue = YWs.Count
        UltraTrackBar1.Value = P
        UltraTrackBar1.Tag = YWs
        Set_TrackBar()

        UltraGrid1.DataSource = tblSIST2
        ASCMAIN1.grdInitializeLayout(UltraGrid1, Me)

        '& "        ['ID', 'Life Expectancy', 'Fertility Rate', 'Region',     'Population']," & vbCrLf _
        '& "        ['CAN',    80.66,              1.67,      'North America',  33739900]," & vbCrLf _
        '& "        ['DEU',    79.84,              1.36,      'Europe',         81902307]," & vbCrLf _
        '& "        ['DNK',    78.6,               1.84,      'Europe',         5523095]," & vbCrLf _
        '& "        ['EGY',    72.73,              2.78,      'Middle East',    79716203]," & vbCrLf _
        '& "        ['GBR',    80.05,              2,         'Europe',         61801570]," & vbCrLf _
        '& "        ['IRN',    72.49,              1.7,       'Middle East',    73137148]," & vbCrLf _
        '& "        ['IRQ',    68.09,              4.77,      'Middle East',    31090763]," & vbCrLf _
        '& "        ['ISR',    81.55,              2.96,      'Middle East',    7485600]," & vbCrLf _
        '& "        ['RUS',    68.6,               1.54,      'Europe',         141850000]," & vbCrLf _
        '& "        ['USA',    78.09,              2.05,      'North America',  307007000]" & vbCrLf _

        ASCDATA1.DeleteRows(tblSIST2, "ONHAND < 200")

        Dim data As String = "['Customer','Inventory', 'SIST Rate', 'Account','Sell-Thru']"
        For Each row As DataRow In tblSIST2.Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim SELLTHRU As Decimal = System.Math.Round(Val(row.Item("SELLTHRU") & ""), 0)
            Dim ONHAND As Decimal = System.Math.Round(Val(row.Item("ONHAND") & ""), 0)
            Dim SELLIN_SELLTHRU As Decimal = System.Math.Round(Val(row.Item("SELLIN_SELLTHRU") & ""), 2)
            If SELLIN_SELLTHRU > 5 Then SELLIN_SELLTHRU = 3
            If SELLIN_SELLTHRU < 0.5 Then SELLIN_SELLTHRU = SELLIN_SELLTHRU * 2
            Dim datum As String = "['" & CUST_CODE & "'," & CStr(ONHAND) & "," & CStr(SELLIN_SELLTHRU) & ",'" & CUST_CODE & "'," & CStr(SELLTHRU) & "]"
            data &= "," & vbCrLf & datum
        Next



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
                & "        title: 'Correlation between Inventory, Sell-Thru and Sell-In vs Sell-Thru Rate'," & vbCrLf _
                & "        hAxis: {title: 'Inventory', scaleType: 'log'}," & vbCrLf _
                & "        vAxis: {title: 'Sell-In / Sell-Thru Rate', viewWindowMode: 'Pretty', scaleType: 'log'}," & vbCrLf _
                & "        bubble: {textStyle: {fontSize: 11}}" & vbCrLf _
                & "      };" & vbCrLf _
                & "" & vbCrLf _
                & "      var chart = new google.visualization.BubbleChart(document.getElementById('series_chart_div'));" & vbCrLf _
                & "      chart.draw(data, options);" & vbCrLf _
                & "    }" & vbCrLf _
                & "    </script>" & vbCrLf _
                & "  </head>" & vbCrLf _
                & "  <body>" & vbCrLf _
                & "    <div id='series_chart_div' style='width: 900px; height: 500px;'></div>" & vbCrLf _
                & "  </body>" & vbCrLf _
                & "</html>"

            sw.WriteLine(html)
        End Using

        WebBrowser1.Navigate(HTM_FILENAME)

        Show_Document(HTM_FILENAME)

        'Dim oWB As SpreadsheetGear.IWorkbook
        'Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        'Dim range As SpreadsheetGear.IRange = Nothing

        '' Parameters

        'Dim Start_Row As Integer = 5

        'oWB = SpreadsheetGear.Factory.GetWorkbook()
        'oSheet = oWB.Worksheets.Add
        'oSheet.Name = "Data"
        'Dim XLS_FILENAME As String = ASCMAIN1.Folders("Work") & "Sell_In_Sell_Thru" & ".xlsx"
        'oWB.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        'range = oSheet.Cells("A1")

        '' Copy the DataTable to the worksheet range.
        'range.CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.None)

        '' Auto size all worksheet columns which contain data
        'oSheet.UsedRange.Columns.AutoFit()

        'oWB.Save()
        'Show_Document(XLS_FILENAME)
        'oWB = Nothing


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub UltraButton2_Click(sender As Object, e As EventArgs) Handles UltraButton2.Click
        Sell_In_Sell_Thru()
        UltraGrid1.Visible = True
        UltraChart1.Visible = True
        UltraTrackBar1.Visible = True
        UltraTabControl2.Visible = True
    End Sub

    Private Sub UltraTrackBar1_ValueChanged(sender As Object, e As EventArgs) Handles UltraTrackBar1.ValueChanged
        Set_TrackBar()
    End Sub

    Sub Set_TrackBar()

        If UltraTrackBar1.Tag Is Nothing Then Exit Sub

        Dim YWs As List(Of String) = DirectCast(UltraTrackBar1.Tag, List(Of String))

        Dim YW = YWs(UltraTrackBar1.Value - 1)
        UltraGrid1.Text = "Sell-In / Sell-Thru " & YW

        '  tblSIST2.Rows.Clear()

        Dim CUST_CODEs As New List(Of String)
        For Each row As DataRow In tblSIST.Select("OPS_YYYYWW = '" & YW & "'", "CUST_CODE")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            CUST_CODEs.Add(CUST_CODE)
            Dim row2 As DataRow = tblSIST2.Rows.Find(CUST_CODE)
            If row2 Is Nothing Then
                row2 = tblSIST2.NewRow
                row2.Item("CUST_CODE") = row.Item("CUST_CODE")
                row2.Item("SELLIN_SELLTHRU") = Val(row.Item("SELLIN_SELLTHRU") & "")
                row2.Item("ONHAND") = Val(row.Item("ONHAND") & "")
                row2.Item("SELLTHRU") = Val(row.Item("SELLTHRU") & "")
                tblSIST2.Rows.Add(row2)
                'tblSIST2.Rows.Add(New Object() {row.Item("CUST_CODE"), _
                '            Val(row.Item("SELLIN_SELLTHRU") & ""), _
                '            Val(row.Item("ONHAND") & ""), _
                '            Val(row.Item("SELLTHRU") & "")})
            Else
                row2.Item("SELLIN_SELLTHRU") = Val(row.Item("SELLIN_SELLTHRU") & "")
                row2.Item("ONHAND") = Val(row.Item("ONHAND") & "")
                row2.Item("SELLTHRU") = Val(row.Item("SELLTHRU") & "")
            End If
        Next

        For Each row As DataRow In tblSIST2.Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            If Not CUST_CODEs.Contains(CUST_CODE) Then
                row.Item("SELLIN_SELLTHRU") = 0
                row.Item("ONHAND") = 0
                row.Item("SELLTHRU") = 0
            End If
        Next
        UltraChart1.Refresh()
        UltraChart1.DataBind()
    End Sub

    Private Sub UltraChart1_Invalidated(sender As Object, e As InvalidateEventArgs) Handles UltraChart1.Invalidated
        'Stop
    End Sub

    Private Sub UltraChart1_InvalidDataReceived(sender As Object, e As Infragistics.UltraChart.Shared.Events.ChartDataInvalidEventArgs) Handles UltraChart1.InvalidDataReceived
        'Stop
    End Sub

    Private Sub UltraGrid1_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles UltraGrid1.InitializeLayout

    End Sub
End Class