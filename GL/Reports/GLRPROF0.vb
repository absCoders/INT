Public Class GLRPROF0

#Region "General Declarations"

    Dim NYP As String

    Dim XZ As String = ""
    Dim YZ As String = ""
    Dim ZZ As String = ""

    Dim LYP As String
    Dim RY As String
    Dim RP As String
    Dim LY As String
    Dim NY As String
    Dim blnDPTonly As Boolean = True

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        RWU = "U"
        Set_cmbYP("RYP", ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), -60, 0, 0)

        If ASCMAIN1.CLIENT = "AHA" Then
            blnDPTonly = False
            chkPREP_BI.Visible = True
        End If

        Get_Stats()
        Create_Summary(grdGLTPROFW, "OPS_YYYYPP", "Count")

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""

        ' Get Run-Time options

        Dim CB As String = ""   ' 3 = both CUST & BRAND in data, 2 = only BRAND in data, 1 = only CUST in data, 0 = neither

        Update_Sales_Summary(RYP)

        ' Get Run-Time options

        ASCMAIN1.Progress("Run-Time Options", "")

        NYP = ASCMAIN1.Period_Calc(RYP, 12)   ' Period Selected + 12; Next Year, Same Period
        LYP = ASCMAIN1.Period_Calc(RYP, -12)  ' Period Selected - 12; Last Year, Same Period
        RY = Mid$(RYP, 1, 4)                  ' Year YYYY of Period Selected
        RP = Mid$(RYP, 5, 2)                  ' Period PP of Period Selected
        LY = Mid$(LYP, 1, 4)                  ' Last Year (Year of Period Selected -1)
        NY = Mid$(NYP, 1, 4)                  ' Next Year (Year of Period Selected +1)

        ' Setup Report Data Table

        ASCMAIN1.sql = "Select * from TATQVXQ1 where ROWNUM < 1"
        Dim TATQVXQ1 As String = ASCMAIN1.Temp_Table

        Dim sql_data As String = ""
        ASCMAIN1.sql = "Select * from ASTDSQLS where FORM_NAME = 'GLRPROF1'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "COLUMN_SEQ")
            Dim COLUMN_NAME As String = row.Item("COLUMN_NAME")
            sql_data &= ", 0.01 " & COLUMN_NAME
        Next

        ' Prepare Work File with Data from Server

        'sql = "Select CUST_CODE"
        'sql = sql & ", SUM ((DECODE(OPS_YYYY,'" & RY & "',DEMO_COMM_RTL,0) / .6) / 100) DC_PCT_TY"
        'sql = sql & ", SUM ((DECODE(OPS_YYYY,'" & LY & "',DEMO_COMM_RTL,0) / .6) / 100) DC_PCT_LY"
        'sql = sql & " from GLTPROF6 where OPS_YYYY in ('" & RY & "','" & LY & "')"
        'sql = sql & " group by CUST_CODE"
        'Call Ora_to_Acc(Nothing, "GLTPROF6X", 1, "", sql)

        ASCMAIN1.sql = "Select ROWNUM RECORD_NO" & vbCrLf _
            & ", ICTBRANX.COLLECTION_CODE" & vbCrLf _
            & ", ICTBRANX.BRAND_CODE" & vbCrLf _
            & ", ICTBRANX.SALES_DIVISION_CODE" & vbCrLf _
            & " from " & vbCrLf _
            & "(" & vbCrLf _
            & " Select DISTINCT DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE " & vbCrLf _
            & " from ICTBRAN1, ICTCOLL1" & vbCrLf _
            & "  where ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ") ICTBRANX"

        Dim TT_PROFY As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & TT_PROFY & " Add Primary Key (RECORD_NO)")
        ASCDATA1.ExecuteSQL("Create Unique Index I_" & TT_PROFY & "_1 on " & TT_PROFY & " (COLLECTION_CODE, BRAND_CODE, SALES_DIVISION_CODE)")

        ASCMAIN1.sql = "Select * from " & TT_PROFY
        Create_TDA(dst.Tables.Add, "GLTPROFY", "**", 0, False, "", 1)
        Fill_Records("GLTPROFY")

        ASCMAIN1.sql = "Select ROWNUM RECORD_NO" & vbCrLf _
            & ", ARTCUST1.CUST_NAME" & vbCrLf _
            & ", XX.CUST_CODE" & vbCrLf _
            & ", ARTSREP1.SREP_CODE" & vbCrLf _
            & ", SOTSREP1.REGION_CODE" & vbCrLf _
            & ", SOTSREG1.VP_CODE" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ", SOTTCLS1.MARKET_CODE" & vbCrLf _
            & ", XX.SALES_DIVISION_CODE" & vbCrLf _
            & ", ARTCUST1.ALL_OTHERS" & vbCrLf _
            & " from ARTCUST1, SOTSREP1, SOTSREG1, SOTTCLS1, ARTSREP1, " & vbCrLf _
            & " (" & vbCrLf _
            & " SELECT ARTCUST1.CUST_CODE, SOTSDIV1.SALES_DIVISION_CODE " & vbCrLf _
            & " from ARTCUST1,SOTSDIV1" & vbCrLf _
            & ") XX" & vbCrLf _
            & " where SOTSREP1.SREP_CODE (+) = ARTSREP1.SREP_CODE" & vbCrLf _
            & "   and SOTSREG1.REGION_CODE (+) = SOTSREP1.REGION_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = XX.CUST_CODE" & vbCrLf _
            & IIf(blnDPTonly, " and ARTCUST1.TRADE_CLASS_CODE in (SELECT TRADE_CLASS_CODE from SOTTCLS1 where MARKET_CODE = 'DPT')" & vbCrLf, "") _
            & "   and ARTSREP1.CUST_CODE (+) = XX.CUST_CODE" & vbCrLf _
            & "   and ARTSREP1.SALES_DIVISION_CODE (+) = XX.SALES_DIVISION_CODE"

        Dim TT_PROFA As String = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("Alter Table " & TT_PROFA & " Add Primary Key (RECORD_NO)")
        ASCDATA1.ExecuteSQL("Create Unique Index I_" & TT_PROFA & "_1 on " & TT_PROFA & " (CUST_CODE, SALES_DIVISION_CODE)")

        ASCMAIN1.sql = "Select * from " & TT_PROFA
        Create_TDA(dst.Tables.Add, "GLTPROFA", "**", 0, False, "", 1)
        Fill_Records("GLTPROFA")

        Dim PROFBf As Integer = 7 ' Number of fields before the numeric columns

        ASCMAIN1.sql = "Select ROWNUM RECORD_NO, GLTPROF1.LINE_NO, GLTPROF1.LINE_TAG" & vbCrLf _
            & ", ARTCUST1.CUST_CODE, ICTCOLL1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & sql_data & vbCrLf _
            & " from ARTCUST1, GLTPROF1, ICTCOLL1, ICTBRAN1, TATWORK1" & vbCrLf _
            & " where ROWNUM <1"

        Dim GLTPROFB As String = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "Select * from " & GLTPROFB
        Create_TDA(dst.Tables.Add, "GLTPROFB", "**", 0, False, "", 0)
        ASCDATA1.ExecuteSQL("Create Index I_" & GLTPROFB & "_1 on " & GLTPROFB & " (CUST_CODE, COLLECTION_CODE, BRAND_CODE, LINE_TAG)")

        Dim PROFBn As Integer = dst.Tables("GLTPROFB").Columns.Count - PROFBf ' Number of numeric columns
        Dim PROFBns() As String
        ReDim PROFBns(PROFBn)
        For i As Integer = 1 To PROFBn
            PROFBns(i) = dst.Tables("GLTPROFB").Columns(i + PROFBf - 1).ColumnName
        Next i

        ' Prepare Report Format Tables

        ASCMAIN1.sql = "Select X.*, ROWNUM RECORD_INDEX from (Select * from GLTPROF1 order by LINE_NO) X"
        Dim GLTPROF1 As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROF1 & " Add Primary Key (LINE_NO)")
        ASCDATA1.ExecuteSQL("Update " & GLTPROF1 & " Set LINE_NO = RECORD_INDEX")

        ASCMAIN1.sql = "Select * from " & GLTPROF1
        Create_TDA(dst.Tables.Add, "GLTPROF1", "**", 0, False, "", 1)
        Fill_Records("GLTPROF1")

        Dim RI As New Collection
        For Each rowGLTPROF1 As DataRow In dst.Tables("GLTPROF1").Select("")
            '  rowGLTPROF1.Item("LINE_NO") = rowGLTPROF1.Item("RECORD_INDEX")
            RI.Add(Val(rowGLTPROF1.Item("LINE_NO") & ""), rowGLTPROF1.Item("LINE_TAG") & "")
        Next

        Dim LINE_NO_max As Integer = dst.Tables("GLTPROF1").Rows.Count

        ' Sales & CGS

        ASCMAIN1.Progress("Shipments", "")

        sql = "Select X.CUST_CODE, X.INV_TYPE, X.ITEM_BASIC_PROMO, X.ITEM_SNG" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        For i As Integer = 1 To 3
            If i = 1 Then
                YZ = "AMT"
            ElseIf i = 2 Then
                YZ = "CGS" ' "VCC"
            Else
                YZ = "ROY"
            End If

            If i = 1 Or i = 2 Then
                zz = "NVL(X.ORDR_" & YZ & "_SHIP,0)" '  * DECODE(X.INV_TYPE,'I',1,-1)"
            Else
                ' NOTE THAT CCU HAD A ROYALTY STOP AT 200601
                ZZ = "NVL(X.ORDR_AMT_SHIP,0) * NVL(ICTBRAN1.ROYALTY_PCT,0) * (CASE WHEN X.OPS_YYYYPP >= '200601' THEN 0 ELSE 1 END) / 100"
            End If
            XZ = "X"
            Make_SQL()
        Next i

        sql &= " From GLTPROF0 X, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
            & " where X.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
            & "   and TT_PROFA.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE (+) = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE" & vbCrLf _
            & " Group by " & vbCrLf _
            & "X.CUST_CODE, X.INV_TYPE, X.ITEM_BASIC_PROMO, X.ITEM_SNG" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        Dim GLTPROFS As String = ASCMAIN1.Temp_Table(sql)

        ASCMAIN1.sql = "Select * from " & GLTPROFS
        Create_TDA(dst.Tables.Add, "GLTPROFS", "**", 0, False, "", 7)
        Fill_Records("GLTPROFS")


        sql = "" _
            & " and ((GLTPROFS.INV_TYPE = 'I' AND GLTPROFS.ITEM_BASIC_PROMO = 'B' AND GLTPROF1.LINE_TAG = 'GSLSB') " & vbCrLf _
            & "   OR (GLTPROFS.INV_TYPE = 'I' AND GLTPROFS.ITEM_BASIC_PROMO = 'P' AND GLTPROF1.LINE_TAG = 'GSLSP') " & vbCrLf _
            & "   OR (GLTPROFS.INV_TYPE = 'C' AND GLTPROF1.LINE_TAG = 'RSLS'))" & vbCrLf
        Insert_Sales("AMT", sql, GLTPROFB, GLTPROFS, GLTPROF1)

        sql = "" _
            & " AND ((GLTPROFS.INV_TYPE = 'I' AND GLTPROFS.ITEM_BASIC_PROMO = 'B' AND GLTPROF1.LINE_TAG = 'GCGSB') " & vbCrLf _
            & "   OR (GLTPROFS.INV_TYPE = 'I' AND GLTPROFS.ITEM_BASIC_PROMO = 'P' AND GLTPROF1.LINE_TAG = 'GCGSP') " & vbCrLf _
            & "   OR (GLTPROFS.INV_TYPE = 'C' AND GLTPROF1.LINE_TAG = 'RCGS'))" & vbCrLf _
            & " and GLTPROFS.ITEM_SNG = 'S'" & vbCrLf

        ' Insert_Sales("VCC", sql, GLTPROFB, GLTPROFS, GLTPROF1)
        Insert_Sales("CGS", sql, GLTPROFB, GLTPROFS, GLTPROF1)

        sql = "" _
            & " AND ((GLTPROFS.ITEM_SNG = 'N' AND GLTPROF1.LINE_TAG = 'SAMPLE') " & vbCrLf _
            & "   OR (GLTPROFS.ITEM_SNG = 'G' AND GLTPROF1.LINE_TAG = 'GWP')" & vbCrLf _
            & "   OR (GLTPROFS.ITEM_SNG = 'D' AND GLTPROF1.LINE_TAG = 'DSP'))" & vbCrLf _
            & " and GLTPROFS.ITEM_SNG <> 'S'" & vbCrLf
        ' Insert_Sales("VCC", sql, GLTPROFB, GLTPROFS, GLTPROF1)
        Insert_Sales("CGS", sql, GLTPROFB, GLTPROFS, GLTPROF1)

        sql = "" _
            & " AND GLTPROF1.LINE_TAG = 'ROYALT' " & vbCrLf
        Insert_Sales("ROY", sql, GLTPROFB, GLTPROFS, GLTPROF1)

        ' Net Sales for Pro-Ration

        ASCMAIN1.Progress("Net Sales Pro-Rata", "")

        ASCMAIN1.sql = "SELECT '0' CB," & vbCrLf _
            & "CUST_CODE," & vbCrLf _
            & "COLLECTION_CODE," & vbCrLf _
            & "BRAND_CODE," & vbCrLf _
            & "SALES_DIVISION_CODE," & vbCrLf _
            & "ACT_LY_MTL AMT_LY_MTL," & vbCrLf _
            & "ACT_LY_YTL AMT_LY_YTL," & vbCrLf _
            & "ACT_LY_YTD AMT_LY_YTD," & vbCrLf _
            & "ACT_TY_MTD AMT_TY_MTD," & vbCrLf _
            & "ACT_TY_YTD AMT_TY_YTD FROM GLTPROFZ,SOTSDIV1 WHERE ROWNUM < 1"
        Dim GLTPROFP As String = ASCMAIN1.Temp_Table()
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFP & " Add Primary Key (CB,CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

        sql = "Select '0' CB, X.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        YZ = "AMT"
        zz = "NVL(X.ORDR_" & YZ & "_SHIP,0)"
        xz = "X"
        Make_SQL()

        sql &= " From GLTPROF0 X, ARTCUST1, ICTCOLL1, ICTBRAN1" & vbCrLf _
            & " where X.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = X.CUST_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = X.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & IIf(blnDPTonly, " and ARTCUST1.TRADE_CLASS_CODE in (SELECT TRADE_CLASS_CODE from SOTTCLS1 where MARKET_CODE = 'DPT')" & vbCrLf, "") _
            & " Group by " & vbCrLf _
            & "X.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
 
        ASCMAIN1.sql = "Insert into " & GLTPROFP & " " & sql
        ASCDATA1.ExecuteSQL()

        Dim kfp As Integer = 5
        ASCMAIN1.sql = "Select * from " & GLTPROFP
        Create_TDA(dst.Tables.Add, "GLTPROFP", "**", 0, False, "", kfp)


        Dim TP() As Decimal
        Dim f As Integer = dst.Tables("GLTPROFP").Columns.Count
        ReDim TP(f - 1)
        sql = ""
        zz = ""
        Dim FP() As String
        ReDim FP(f - 1)
        For i As Integer = kfp To f - 1
            Dim z As String = dst.Tables("GLTPROFP").Columns(i).ColumnName
            FP(i) = z
            sql = sql & ", Sum (GLTPROFP." & z & ") AS " & z
            ZZ = ZZ & ", GLTPROFP." & z
        Next i

        ASCDATA1.ExecuteSQL("Insert into " & GLTPROFP & " Select '1' as CB, CUST_CODE, COLLECTION_CODE, BRAND_CODE, SALES_DIVISION_CODE" & ZZ & " from " & GLTPROFP & " GLTPROFP where CB = '0'")
        ASCDATA1.ExecuteSQL("Insert into " & GLTPROFP & " Select '2' as CB, CUST_CODE, COLLECTION_CODE, BRAND_CODE, SALES_DIVISION_CODE" & ZZ & " from " & GLTPROFP & " GLTPROFP where CB = '0'")

       
        Dim sqlx As String = sql

        ASCMAIN1.sql = "Select '0' AS CB, 'ALL_C', 'ALL_B', 'ALL_D', 'ALL_U'" & sqlx & " from " & GLTPROFP & " GLTPROFP where CB = '0'"
        Dim rowGLTPROFP As DataRow = ASCDATA1.GetDataRow
        sql = ""
        For i As Integer = kfp To f - 1
            Dim z As String = FP(i) ' dynGLTPROFP.Fields(i).Name
            TP(i) = Val(rowGLTPROFP.Item(i) & "")
            If TP(i) = 0 Then
                sql = sql & ", " & z & " = 0"
            Else
                sql = sql & ", " & z & " = " & z & " / " & CStr(TP(i))
            End If
        Next i
        ASCMAIN1.sql = "Update " & GLTPROFP & " GLTPROFP set " & Mid$(sql, 2) & " where CB = '0'"
        ASCDATA1.ExecuteSQL()

        Fill_Records("GLTPROFP")

        Dim CBc As String
        Dim CBi As Integer
        For CBi = 1 To 2
            If CBi = 1 Then
                CBc = "CUST_CODE, SALES_DIVISION_CODE"
            Else
                CBc = "BRAND_CODE, COLLECTION_CODE, SALES_DIVISION_CODE"
            End If
            CB = Format$(CBi, "0")
            ASCMAIN1.sql = "Select Distinct " & CBc & " from " & GLTPROFP & " GLTPROFP where CB = '" & CB & "'"
            For Each rowGLTPROFP_CUST As DataRow In ASCDATA1.GetDataTable.Select("")
                ASCMAIN1.sql = "Select '" & CB & "' AS CB, " & IIf(CBi = 1, _
                                                          " CUST_CODE, 'ALL_B', 'ALL_L', SALES_DIVISION_CODE", _
                                                          " 'ALL_C', BRAND_CODE, COLLECTION_CODE, SALES_DIVISION_CODE") _
                                                      & sqlx & " from " & GLTPROFP & " GLTPROFP where CB = '" & CB & "'"
                If CBi = 1 Then
                    ASCMAIN1.sql &= " and CUST_CODE = '" & rowGLTPROFP_CUST.Item("CUST_CODE") & "'" & vbCrLf _
                        & " and SALES_DIVISION_CODE = '" & rowGLTPROFP_CUST.Item("SALES_DIVISION_CODE") & "'" & vbCrLf
                Else
                    ASCMAIN1.sql &= " and BRAND_CODE = '" & rowGLTPROFP_CUST.Item("BRAND_CODE") & "'" & vbCrLf _
                        & " and COLLECTION_CODE = '" & rowGLTPROFP_CUST.Item("COLLECTION_CODE") & "'" & vbCrLf _
                        & " and SALES_DIVISION_CODE = '" & rowGLTPROFP_CUST.Item("SALES_DIVISION_CODE") & "'" & vbCrLf
                End If
                ASCMAIN1.sql &= " group by " & CBc

                rowGLTPROFP = ASCDATA1.GetDataRow

                ASCMAIN1.sql = ""
                For i As Integer = kfp To f - 1
                    Dim z As String = FP(i)
                    TP(i) = Val(rowGLTPROFP.Item(i) & "")
                    If TP(i) = 0 Then
                        ASCMAIN1.sql &= ", " & z & " = 0"
                    Else
                        ASCMAIN1.sql &= ", " & z & " = " & z & " / " & CStr(TP(i))
                    End If
                Next i

                ASCMAIN1.sql = "Update " & GLTPROFP & " GLTPROFP set " & Mid$(ASCMAIN1.sql, 2) & " where CB = '" & CB & "'"
                If CBi = 1 Then
                    ASCMAIN1.sql &= " and CUST_CODE = '" & rowGLTPROFP_CUST.Item("CUST_CODE") & "'" & vbCrLf _
                        & " and SALES_DIVISION_CODE = '" & rowGLTPROFP_CUST.Item("SALES_DIVISION_CODE") & "'"
                Else
                    ASCMAIN1.sql &= " and BRAND_CODE = '" & rowGLTPROFP_CUST.Item("BRAND_CODE") & "'" & vbCrLf _
                        & " and COLLECTION_CODE = '" & rowGLTPROFP_CUST.Item("COLLECTION_CODE") & "'" & vbCrLf _
                        & " and SALES_DIVISION_CODE = '" & rowGLTPROFP_CUST.Item("SALES_DIVISION_CODE") & "'" & vbCrLf
                End If
                ASCDATA1.ExecuteSQL()
            Next
        Next CBi


        ' Accrued Returns

        Dim SOTRMAFA As String = ""

        ASCMAIN1.sql = "Select OPS_YYYYPP from GLTPARM2 where OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "OPS_YYYYPP")
            Dim YP As String = row.Item("OPS_YYYYPP")
            Dim LYP As String = ASCMAIN1.Period_Calc(YP, -1)

            ASCMAIN1.sql = "Select X.OPS_YYYYPP, X.POST_OR_REV, SOTRMAF1.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", X.ITEM_CODE, SOTRMAF1.RA_NO, SOTRMAF1.CUST_CLAIM_NO, SOTRMAF1.RA_REASON_CODE" & vbCrLf _
                & ", SUM (-1 * X.RA_AMT_OPEN) RA_AMT_OPEN, SUM (-1 * X.RA_CGS_OPEN) RA_CGS_OPEN" & vbCrLf _
                & " from ICTITEM1,SOTRMAF1,(" & vbCrLf _
                & "Select '0' POST_OR_REV, SOTRMAF3.RA_NO, SOTRMAF3.OPS_YYYYPP, SOTRMAF3.ITEM_CODE" & vbCrLf _
                & ", SOTRMAF3.RA_QTY_OPEN * SOTRMAF3.RA_NET_PRICE RA_AMT_OPEN" & vbCrLf _
                & ", SOTRMAF3.RA_QTY_OPEN * ICTCOSTA.ITEM_COST_TOTAL RA_CGS_OPEN" & vbCrLf _
                & " from SOTRMAF3, ICTCOSTA" & vbCrLf _
                & " where SOTRMAF3.OPS_YYYYPP = '" & YP & "'" & vbCrLf _
                & "   and SOTRMAF3.RA_QTY_OPEN <> 0 and SOTRMAF3.RA_NET_PRICE <> 0" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP = SOTRMAF3.OPS_YYYYPP and ICTCOSTA.ITEM_CODE = SOTRMAF3.ITEM_CODE" & vbCrLf _
                & " union " & vbCrLf _
                & "Select '1' POST_OR_REV, SOTRMAF3.RA_NO, '" & YP & "' OPS_YYYYPP, SOTRMAF3.ITEM_CODE" & vbCrLf _
                & ", -1 * SOTRMAF3.RA_QTY_OPEN * SOTRMAF3.RA_NET_PRICE RA_AMT_OPEN" & vbCrLf _
                & ", -1 * SOTRMAF3.RA_QTY_OPEN * ICTCOSTA.ITEM_COST_TOTAL RA_CGS_OPEN" & vbCrLf _
                & " from SOTRMAF3, ICTCOSTA" & vbCrLf _
                & " where SOTRMAF3.OPS_YYYYPP = '" & LYP & "'" & vbCrLf _
                & "   and SOTRMAF3.RA_QTY_OPEN <> 0 and SOTRMAF3.RA_NET_PRICE <> 0" & vbCrLf _
                & "   and ICTCOSTA.OPS_YYYYPP = SOTRMAF3.OPS_YYYYPP and ICTCOSTA.ITEM_CODE = SOTRMAF3.ITEM_CODE" & vbCrLf _
                & ") X where SOTRMAF1.RA_NO = X.RA_NO and ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
                & " group by X.OPS_YYYYPP, X.POST_OR_REV, SOTRMAF1.CUST_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", X.ITEM_CODE, SOTRMAF1.RA_NO, SOTRMAF1.CUST_CLAIM_NO, SOTRMAF1.RA_REASON_CODE" & vbCrLf _
                & " having Sum (X.RA_AMT_OPEN) <> 0"

            If SOTRMAFA = "" Then
                SOTRMAFA = ASCMAIN1.Temp_Table
            Else

                ASCMAIN1.sql = "Insert into " & SOTRMAFA & " " & ASCMAIN1.sql
                ASCDATA1.ExecuteSQL()

            End If
        Next



        ASCMAIN1.Progress("Accrued Returns", "")

        sql = "Select SOTRMAFA.CUST_CODE " & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
        YZ = "RSLSAC"
        ZZ = "RA_AMT_OPEN"
        XZ = "SOTRMAFA"
        Make_SQL("OPS_YYYYPP")
        sql &= " From " & SOTRMAFA & " SOTRMAFA, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
            & " where SOTRMAFA.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
            & "   and TT_PROFA.CUST_CODE = SOTRMAFA.CUST_CODE" & vbCrLf _
            & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = SOTRMAFA.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
            & " Group by " & vbCrLf _
            & "SOTRMAFA.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        Dim GLTPROFU As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFU & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

        ASCMAIN1.sql = "Select * from " & GLTPROFU
        Create_TDA(dst.Tables.Add, "GLTPROFU", "**", 0, False, "", 4)
        Fill_Records("GLTPROFU")

        Update_GLTPROFB(RI, "3", GLTPROFU, GLTPROFB, YZ)


        If ASCMAIN1.CLIENT = "AHA" Then

            Dim rowSOTPARM1 As DataRow = LookUp("SOTPARM1", "Z")

            Dim sqlR As String = "Select SUBSTR(X.OPS_YYYYPP,1,4) YYYY, SUBSTR(X.OPS_YYYYPP,5,2) PP" & vbCrLf _
                & ", X.CUST_CODE, X.ITEM_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & ", -1 * X.RA_AMT_OPEN AMOUNT, X.POST_OR_REV, 'R' EXP_TYPE, X.RA_REASON_CODE CATGY_CODE, X.RA_NO REFERENCE_NO" & vbCrLf _
                & ", X.CUST_CLAIM_NO CUST_REFERENCE, 'P' ACCR_OR_PYMT, '" & rowSOTPARM1.Item("SO_PARM_RA_REV_ACCT_CODE") & "' ACCT_CODE, NULL IGNORE, '0' RECORD_STATUS" & vbCrLf _
                & ", SYSDATE TIMESTAMP_0, NULL TIMESTAMP_1, NULL TIMESTAMP_2, NULL VERIFICATION_NO" & vbCrLf _
                & " from ICTCOLL1, " & SOTRMAFA & " X " & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE" & vbCrLf _
                & "   and X.OPS_YYYYPP = '" & RYP & "'"
            ASCMAIN1.sql = "Insert into " & TATQVXQ1 & " " & sqlR
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & TATQVXQ1 & " " & _
                Replace(Replace(Replace(sqlR, _
                                        "-1 * X.RA_AMT_OPEN AMOUNT", "X.RA_CGS_OPEN AMOUNT"), _
                                        "'" & rowSOTPARM1.Item("SO_PARM_RA_REV_ACCT_CODE") & "' ACCT_CODE", "'" & rowSOTPARM1.Item("SO_PARM_RA_CGS_ACCT_CODE") & "' ACCT_CODE"), _
                                        "NULL IGNORE", "'Y' IGNORE")
            ASCDATA1.ExecuteSQL()

        End If


        ASCMAIN1.Progress("Accrued Cost of Goods Returned", "")

        sql = "Select SOTRMAFA.CUST_CODE " & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
        YZ = "RCGSAC"
        ZZ = "RA_CGS_OPEN"
        XZ = "SOTRMAFA"
        Make_SQL("OPS_YYYYPP")
        sql &= " From " & SOTRMAFA & " SOTRMAFA, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
            & " where SOTRMAFA.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
            & "   and TT_PROFA.CUST_CODE = SOTRMAFA.CUST_CODE" & vbCrLf _
            & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = SOTRMAFA.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
            & " Group by " & vbCrLf _
            & "SOTRMAFA.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        Dim GLTPROFV As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFV & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

        ASCMAIN1.sql = "Select * from " & GLTPROFV
        Create_TDA(dst.Tables.Add, "GLTPROFV", "**", 0, False, "", 4)
        Fill_Records("GLTPROFV")

        Update_GLTPROFB(RI, "3", GLTPROFV, GLTPROFB, YZ)





        ' Promo Event Expense

        ASCMAIN1.Progress("Promo Event Expense", "")


        ASCMAIN1.sql = "Select OPS_YYYYPP, CUST_CODE, COLLECTION_CODE, 'xxxxxx' TYPE, VEHICLE_CODE, DIST_AMT from SPTCOOPA where ROWNUM < 1"
        Dim SPTPROM1 As String = ASCMAIN1.Temp_Table

        Dim TYP_PROMO As String = ASCMAIN1.Period_Calc(LY & "01", -1)
        Dim LYP_PROMO As String = ""

        For M As Integer = 1 To 24
            LYP_PROMO = TYP_PROMO
            TYP_PROMO = ASCMAIN1.Period_Calc(LYP_PROMO, 1)
            ' If ASCMAIN1.Running_in_VS And TYP_PROMO = "201709" Then Stop
            ASCMAIN1.sql = "" _
                & "Select 'SPCA+' TYPE, 'TYP_PROMO' OPS_YYYYPP, CUST_CODE, COLLECTION_CODE, VEHICLE_CODE, Sum (DIST_AMT) DIST_AMT" & vbCrLf _
                & "  From SPTCOOPA where SPTCOOPA.OPS_YYYYPP = 'TYP_PROMO'   AND OPS_YYYYPP_H <= 'TYP_PROMO'" & vbCrLf _
                & " Group by CUST_CODE, COLLECTION_CODE, VEHICLE_CODE" & vbCrLf _
                & "union" & vbCrLf _
                & "Select 'SPCA-' TYPE, 'TYP_PROMO' OPS_YYYYPP, CUST_CODE, COLLECTION_CODE, VEHICLE_CODE, Sum (-1 * DIST_AMT) DIST_AMT" & vbCrLf _
                & "  From SPTCOOPA where SPTCOOPA.OPS_YYYYPP = 'LYP_PROMO'   AND OPS_YYYYPP_H <= 'LYP_PROMO'" & vbCrLf _
                & " Group by CUST_CODE, COLLECTION_CODE, VEHICLE_CODE" & vbCrLf _
                & "union" & vbCrLf _
                & "Select 'SPCP' TYPE, SPTPYMT1.OPS_YYYYPP, SPTPYMT1.CUST_CODE, SPTPYMT3.COLLECTION_CODE, SPTCOOP1.VEHICLE_CODE, SUM (SPTPYMT3.DIST_AMT_PYMT) DIST_AMT" & vbCrLf _
                & " From SPTPYMT1,SPTPYMT2,SPTPYMT3,SPTCOOP1" & vbCrLf _
                & " where SPTPYMT1.OPS_YYYYPP = 'TYP_PROMO'" & vbCrLf _
                & "   and SPTPYMT2.PYMT_NO = SPTPYMT1.PYMT_NO" & vbCrLf _
                & "   and SPTPYMT3.PYMT_NO = SPTPYMT2.PYMT_NO" & vbCrLf _
                & "   and SPTPYMT3.PYMT_LNO = SPTPYMT2.PYMT_LNO" & vbCrLf _
                & "   and SPTCOOP1.AUTH_NO = SPTPYMT2.AUTH_NO" & vbCrLf _
                & " Group by SPTPYMT1.OPS_YYYYPP, SPTPYMT1.CUST_CODE, SPTPYMT3.COLLECTION_CODE, SPTCOOP1.VEHICLE_CODE"

            ASCMAIN1.sql = "Select OPS_YYYYPP, CUST_CODE, COLLECTION_CODE, TYPE, VEHICLE_CODE, SUM (DIST_AMT) DIST_AMT from " & vbCrLf _
                & "(" & Replace(Replace(ASCMAIN1.sql, "TYP_PROMO", TYP_PROMO), "LYP_PROMO", LYP_PROMO) & vbCrLf _
                & ") group by OPS_YYYYPP, CUST_CODE, COLLECTION_CODE, TYPE, VEHICLE_CODE" & vbCrLf _
                & " having SUM (DIST_AMT) <> 0"

            ASCMAIN1.sql = "Insert into " & SPTPROM1 & " " & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()
        Next

        sql = "Select SPTPROM1.CUST_CODE " & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
        YZ = "PROMO"
        ZZ = "DIST_AMT"
        XZ = "SPTPROM1"
        Make_SQL("OPS_YYYYPP")
        sql &= " From " & SPTPROM1 & " SPTPROM1, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
            & " where SPTPROM1.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
            & "   and TT_PROFA.CUST_CODE = SPTPROM1.CUST_CODE" & vbCrLf _
            & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = SPTPROM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
            & " Group by " & vbCrLf _
            & "SPTPROM1.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        Dim GLTPROFC As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFC & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

        ASCMAIN1.sql = "Select * from " & GLTPROFC
        Create_TDA(dst.Tables.Add, "GLTPROFC", "**", 0, False, "", 4)
        Fill_Records("GLTPROFC")

        Update_GLTPROFB(RI, "3", GLTPROFC, GLTPROFB, YZ)

        If ASCMAIN1.CLIENT = "AHA" Then
            '& "    and SPTTYPE1.EXPENSE_TYPE_CODE (+) = DECODE(X.VEHICLE_CODE,'RTV','DIF',X.VEHICLE_CODE)" & vbCrLf _
            ASCMAIN1.sql = "SELECT SUBSTR(X.OPS_YYYYPP,1,4) YYYY, SUBSTR(X.OPS_YYYYPP,5,2) PP, X.CUST_CODE, NULL ITEM_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                & ", X.DIST_AMT AMOUNT, DECODE(X.TYPE,'SPCA-','1','0') POST_OR_REV, 'P' EXP_TYPE, X.VEHICLE_CODE CATGY_CODE, NULL REFERENCE_NO" & vbCrLf _
                & ", NULL CUST_REFERENCE, DECODE(X.TYPE,'SPCP','P','A') ACCR_OR_PYMT, SPTTYPE1.ACCT_CODE, NULL IGNORE, '0' RECORD_STATUS" & vbCrLf _
                & ", SYSDATE TIMESTAMP_0, NULL TIMESTAMP_1, NULL TIMESTAMP_2, NULL VERIFICATION_NO from SPTTYPE1,ICTCOLL1, " & SPTPROM1 & " X" & vbCrLf _
                & "  where ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE" & vbCrLf _
                & "    and X.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
                & "    and SPTTYPE1.EXPENSE_TYPE_CODE (+) = X.VEHICLE_CODE" & vbCrLf _
                & "    and X.DIST_AMT <> 0"
            ASCMAIN1.sql = "Insert into " & TATQVXQ1 & " " & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()
        End If




        If ASCMAIN1.CLIENT = "INT" Then
            ' INT DOES NOT DO ACOM
        Else


            ' Commission & Advertising Expense

            ASCMAIN1.Progress("Commission & Advertising Expense", "Commission")

            sql = "Select SPTACOMB.CUST_CODE " & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
            YZ = "COMM"
            ZZ = "AMT_COMM"
            XZ = "SPTACOMB"
            Make_SQL()
            sql &= " From SPTACOMB, SPTACOM1, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
                & " where SPTACOMB.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
                & "   and TT_PROFA.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = SPTACOMB.COLLECTION_CODE" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE = SPTACOMB.ASP_CODE" & vbCrLf _
                & "   and SPTACOM1.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE in ('COMMRTLR')" & vbCrLf _
                & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
                & " Group by " & vbCrLf _
                & "SPTACOMB.CUST_CODE" & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

            Dim GLTPROFA_COMM As String = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFA_COMM & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

            ASCMAIN1.sql = "Select * from " & GLTPROFA_COMM
            Create_TDA(dst.Tables.Add, "GLTPROFA_COMM", "**", 0, False, "", 4)
            Fill_Records("GLTPROFA_COMM")

            Update_GLTPROFB(RI, "3", GLTPROFA_COMM, GLTPROFB, YZ)


            ASCMAIN1.Progress("Commission & Advertising Expense", "Fixed Allowance")

            sql = "Select SPTACOMB.CUST_CODE " & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
            YZ = "FIX"
            ZZ = "AMT_COMM"
            XZ = "SPTACOMB"
            Make_SQL()
            sql &= " From SPTACOMB, SPTACOM1, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
                & " where SPTACOMB.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
                & "   and TT_PROFA.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = SPTACOMB.COLLECTION_CODE" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE = SPTACOMB.ASP_CODE" & vbCrLf _
                & "   and SPTACOM1.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE in ('ADVBROAD','ADVCOOP')" & vbCrLf _
                & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
                & " Group by " & vbCrLf _
                & "SPTACOMB.CUST_CODE" & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

            Dim GLTPROFA_FIX As String = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFA_FIX & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

            ASCMAIN1.sql = "Select * from " & GLTPROFA_FIX
            Create_TDA(dst.Tables.Add, "GLTPROFA_FIX", "**", 0, False, "", 4)
            Fill_Records("GLTPROFA_FIX")

            Update_GLTPROFB(RI, "3", GLTPROFA_FIX, GLTPROFB, YZ)




            ASCMAIN1.Progress("Commission & Advertising Expense", "Destroyed in Field")

            sql = "Select SPTACOMB.CUST_CODE " & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
            YZ = "DIF"
            ZZ = "AMT_COMM"
            XZ = "SPTACOMB"
            Make_SQL()
            sql &= " From SPTACOMB, SPTACOM1, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
                & " where SPTACOMB.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
                & "   and TT_PROFA.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = SPTACOMB.COLLECTION_CODE" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE = SPTACOMB.ASP_CODE" & vbCrLf _
                & "   and SPTACOM1.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE in ('DIF')" & vbCrLf _
                & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
                & " Group by " & vbCrLf _
                & "SPTACOMB.CUST_CODE" & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

            Dim GLTPROFA_DIF As String = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFA_DIF & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

            ASCMAIN1.sql = "Select * from " & GLTPROFA_DIF
            Create_TDA(dst.Tables.Add, "GLTPROFA_DIF", "**", 0, False, "", 4)
            Fill_Records("GLTPROFA_DIF")

            Update_GLTPROFB(RI, "3", GLTPROFA_DIF, GLTPROFB, YZ)



            ASCMAIN1.Progress("Commission & Advertising Expense", "In-Store Merchandising")

            sql = "Select SPTACOMB.CUST_CODE " & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
            YZ = "TST"
            ZZ = "AMT_COMM"
            XZ = "SPTACOMB"
            Make_SQL()
            sql &= " From SPTACOMB, SPTACOM1, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
                & " where SPTACOMB.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
                & "   and TT_PROFA.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = SPTACOMB.COLLECTION_CODE" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE = SPTACOMB.ASP_CODE" & vbCrLf _
                & "   and SPTACOM1.CUST_CODE = SPTACOMB.CUST_CODE" & vbCrLf _
                & "   and SPTACOM1.ASP_CODE in ('TESTERS')" & vbCrLf _
                & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
                & " Group by " & vbCrLf _
                & "SPTACOMB.CUST_CODE" & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

            Dim GLTPROFA_TST As String = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFA_TST & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

            ASCMAIN1.sql = "Select * from " & GLTPROFA_TST
            Create_TDA(dst.Tables.Add, "GLTPROFA_TST", "**", 0, False, "", 4)
            Fill_Records("GLTPROFA_TST")

            Update_GLTPROFB(RI, "3", GLTPROFA_TST, GLTPROFB, YZ)

            ASCMAIN1.sql = "Insert into " & TATQVXQ1 & vbCrLf _
                & "Select SUBSTR(GLTCREC3.OPS_YYYYPP,1,4) YYYY, SUBSTR(GLTCREC3.OPS_YYYYPP,5,2) PP" & vbCrLf _
                & ", DETL_CVX_NO,NULL, SPTACOMC.BRAND_CODE, CREC_AMT, '0', 'C', CREC_CLASS_CODE, DETL_CTL_NO" & vbCrLf _
                & ", SPTACOMC.PYMT_NO, 'A', SPTACOM0.ASP_ACCT_CODE_EXP, NULL, '0', SYSDATE, NULL, NULL, NULL" & vbCrLf _
                & " from GLTCREC3,SPTACOMC,SPTACOM0" & vbCrLf _
                & " where CREC_TYPE_CODE = 'AC' AND SPTACOMC.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
                & "   and SPTACOMC.ACC_CTL_NO = GLTCREC3.DETL_CTL_NO" & vbCrLf _
                & "   and SPTACOM0.ASP_CODE = GLTCREC3.CREC_CLASS_CODE" & vbCrLf _
                & "   and DETL_CTL_TYPE = 'A'"
            ASCDATA1.ExecuteSQL()
        End If




        If ASCMAIN1.CLIENT = "INT" Then
            ' INT DOES NOT DO ARDED OR MISC
        Else
            ' ASSUMPTIONS
            ' USING ACCTS DEFINED IN LINES 3, 6, 7 OF OPEX
            ' HARD CODING AVVAL AS THE COLLECTION
            ' WRITE-OFF IS INCLUDED AS PART OF SRA

            Dim sqlOPEX As String = "(SELECT ACCT_CODE FROM GLTFINR3 WHERE STMT_CODE = 'OPEX' AND STMT_LINE_NO IN (3,6,7))"

            ' AR Deductions

            ASCMAIN1.Progress("AR Deductions", "")


            ASCMAIN1.sql = "SELECT ARTREAS1.ACCT_CODE FROM ARTPARM1,ARTREAS1 WHERE ARTREAS1.REASON_CODE (+) = ARTPARM1.AR_PARM_REASON_CODE_WOFF"
            Dim rowARTPARM1 As DataRow = ASCDATA1.GetDataRow
            Dim ACCT_CODE_WO As String = "?"
            If rowARTPARM1 IsNot Nothing Then
                ACCT_CODE_WO = rowARTPARM1.Item(0)
            End If

            Dim sqlARDED As String = "Select ARTPYMT2.PYMT_BATCH_NO || '-' || ARTPYMT2.PYMT_BATCH_LNO REFERENCE_NO, ARTPYMT2.CUST_CODE, ARTPYMT1.OPS_YYYYPP, 'DISTGL' CATGY_CODE, ARTPYMT4.ACCT_CODE, ARTPYMT4.GL_DIST_AMT" & vbCrLf _
               & " From ARTPYMT1, ARTPYMT2, ARTPYMT4" & vbCrLf _
               & " where ARTPYMT1.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
               & "   and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" & vbCrLf _
               & "   and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" & vbCrLf _
               & "   and ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" & vbCrLf _
               & "   and ARTPYMT4.ACCT_CODE in " & sqlOPEX & vbCrLf _
               & "   and ARTPYMT4.GL_DIST_AMT <> 0" & vbCrLf

            ASCMAIN1.sql = sqlARDED

            Dim ARTPYMTX As String = ASCMAIN1.Temp_Table
            ASCMAIN1.sql = "Insert into " & ARTPYMTX & " " & Replace(Replace(sqlARDED, "ARTPYMT4", "ARTPYMT5"), "'DISTGL'", "ARTPYMT5.REASON_CODE")
            ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "Insert into " & ARTPYMTX & " " & Replace(Replace(Replace(Replace(Replace(sqlARDED, "ARTPYMT4", "ARTPYMT3"), "GL_DIST_AMT", "INV_WRITE_OFF"), "   and ARTPYMT3.ACCT_CODE in " & sqlOPEX, ""), "'DISTGL'", "'DISTWO'"), "ARTPYMT3.ACCT_CODE", "'" & ACCT_CODE_WO & "' ACCT_CODE")
            ASCDATA1.ExecuteSQL()

            sql = "Select ARTPYMTX.CUST_CODE " & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
            YZ = "ARDED"
            ZZ = "GL_DIST_AMT"
            XZ = "ARTPYMTX"
            Make_SQL()
            sql &= " From " & ARTPYMTX & " ARTPYMTX, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
                & " where ARTPYMTX.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
                & "   and TT_PROFA.CUST_CODE = ARTPYMTX.CUST_CODE" & vbCrLf _
                & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = 'AVVAL'" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & " Group by " & vbCrLf _
                & "ARTPYMTX.CUST_CODE" & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

            Dim GLTPROFA_ARDED As String = ASCMAIN1.Temp_Table(sql)

            ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFA_ARDED & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

            ASCMAIN1.sql = "Select * from " & GLTPROFA_ARDED
            Create_TDA(dst.Tables.Add, "GLTPROFA_ARDED", "**", 0, False, "", 4)
            Fill_Records("GLTPROFA_ARDED")

            Update_GLTPROFB(RI, "3", GLTPROFA_ARDED, GLTPROFB, YZ)

            If ASCMAIN1.CLIENT = "AHA" Then
                ASCMAIN1.sql = "Select SUBSTR(X.OPS_YYYYPP,1,4) YYYY, SUBSTR(X.OPS_YYYYPP,5,2) PP" & vbCrLf _
                    & ", X.CUST_CODE, NULL ITEM_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                    & ", X.GL_DIST_AMT AMOUNT, '0' POST_OR_REV, 'A' EXP_TYPE, X.CATGY_CODE, X.REFERENCE_NO" & vbCrLf _
                    & ", NULL CUST_REFERENCE, 'P' ACCR_OR_PYMT, X.ACCT_CODE, NULL IGNORE, '0' RECORD_STATUS" & vbCrLf _
                    & ", SYSDATE TIMESTAMP_0, NULL TIMESTAMP_1, NULL TIMESTAMP_2, NULL VERIFICATION_NO" & vbCrLf _
                    & " from ICTCOLL1, " & ARTPYMTX & " X " & vbCrLf _
                    & " where ICTCOLL1.COLLECTION_CODE (+) = 'AVVAL'" & vbCrLf _
                    & "   and X.OPS_YYYYPP = '" & RYP & "'"
                ASCMAIN1.sql = "Insert into " & TATQVXQ1 & " " & ASCMAIN1.sql
                ASCDATA1.ExecuteSQL()
            End If



            ' Misc Charges

            ASCMAIN1.Progress("Misc Charges", "")


            sql = "Select SOTINVH1.CUST_CODE " & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
            YZ = "MISC"
            ZZ = "INV_MISC_CHG"
            XZ = "SOTINVH1"
            Make_SQL("ORDR_YYYYPP_UPDATED")
            sql &= " From SOTINVH1, SOTMISC1, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
                & " where SOTINVH1.ORDR_YYYYPP_UPDATED BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
                & "   and TT_PROFA.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = 'AVVAL'" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SOTMISC1.MISC_CHG_CODE = SOTINVH1.MISC_CHG_CODE" & vbCrLf _
                & "   and SOTMISC1.ACCT_CODE in " & sqlOPEX & vbCrLf _
                & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
                & " Group by " & vbCrLf _
                & "SOTINVH1.CUST_CODE" & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

            Dim GLTPROFA_MISC As String = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFA_MISC & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

            ASCMAIN1.sql = "Select * from " & GLTPROFA_MISC
            Create_TDA(dst.Tables.Add, "GLTPROFA_MISC", "**", 0, False, "", 4)
            Fill_Records("GLTPROFA_MISC")

            Update_GLTPROFB(RI, "3", GLTPROFA_MISC, GLTPROFB, YZ)

            If ASCMAIN1.CLIENT = "AHA" Then
                ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.ORDR_CUST_PO" & vbCrLf _
                    & ", SOTINVH1.CUST_CODE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                    & ", 'AVVAL' COLLECTION_CODE" & vbCrLf _
                    & ", -1 * SOTINVH1.INV_MISC_CHG INV_MISC_CHG, SOTINVH1.MISC_CHG_CODE, SOTMISC1.ACCT_CODE" & vbCrLf _
                    & " from SOTINVH1, SOTMISC1" & vbCrLf _
                    & " where SOTINVH1.ORDR_YYYYPP_UPDATED = '" & RYP & "'" & vbCrLf _
                    & "   and SOTMISC1.MISC_CHG_CODE = SOTINVH1.MISC_CHG_CODE" & vbCrLf _
                    & "   and SOTMISC1.ACCT_CODE in " & sqlOPEX & vbCrLf _
                    & "   and SOTINVH1.INV_MISC_CHG <> 0" & vbCrLf

                ASCMAIN1.sql = "Select SUBSTR(X.OPS_YYYYPP,1,4) YYYY, SUBSTR(X.OPS_YYYYPP,5,2) PP" & vbCrLf _
                    & ", X.CUST_CODE, NULL ITEM_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
                    & ", X.INV_MISC_CHG AMOUNT, '0' POST_OR_REV, 'M' EXP_TYPE, X.MISC_CHG_CODE CATGY_CODE, X.INV_NO REFERENCE_NO" & vbCrLf _
                    & ", X.ORDR_CUST_PO CUST_REFERENCE, 'P' ACCR_OR_PYMT, X.ACCT_CODE, NULL IGNORE, '0' RECORD_STATUS" & vbCrLf _
                    & ", SYSDATE TIMESTAMP_0, NULL TIMESTAMP_1, NULL TIMESTAMP_2, NULL VERIFICATION_NO" & vbCrLf _
                    & " from ICTCOLL1, (" & ASCMAIN1.sql & ") X " & vbCrLf _
                    & " where ICTCOLL1.COLLECTION_CODE (+) = X.COLLECTION_CODE" & vbCrLf _
                    & "   and X.OPS_YYYYPP = '" & RYP & "'"
                ASCMAIN1.sql = "Insert into " & TATQVXQ1 & " " & ASCMAIN1.sql
                ASCDATA1.ExecuteSQL()
            End If

        End If




        ' Modeling

        ASCMAIN1.Progress("Modeling", "")

        sql = "Select X2.CUST_CODE " & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
        YZ = "MODEL"
        'ZZ = "NVL(X2.BILL_RATE,0)*NVL(X2.BILL_HOURS,0)"
        ZZ = "NVL(X2.BILL_AMT,0)" ' COWORX NET AMT MINUS SALES TAX
        XZ = "X2"
        Make_SQL()

        sql &= " From SPTCWRX2 X2, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
            & " where X2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
            & "   and TT_PROFA.CUST_CODE = X2.CUST_CODE" & vbCrLf _
            & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = X2.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "   and NVL(X2.BILL_RATE,0)*NVL(X2.BILL_HOURS,0) <> 0" & vbCrLf _
            & " Group by " & vbCrLf _
            & "X2.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        Dim GLTPROFM As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFM & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

        ASCMAIN1.sql = "Select * from " & GLTPROFM
        Create_TDA(dst.Tables.Add, "GLTPROFM", "**", 0, False, "", 4)
        Fill_Records("GLTPROFM")

        Update_GLTPROFB(RI, "3", GLTPROFM, GLTPROFB, "MODROT")


        ' Retail

        ASCMAIN1.Progress("Retail", "")

        sql = "Select X2.CUST_CODE " & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
        YZ = "RETAIL"
        ZZ = "NVL(X2.AMT_SOLD,0)"
        XZ = "X2"
        Make_SQL()

        sql &= " From RSTRETL1 X2, ICTCOLL1, ICTITEM1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
            & " where X2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
            & "   and TT_PROFA.CUST_CODE = X2.CUST_CODE" & vbCrLf _
            & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = X2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & " Group by " & vbCrLf _
            & "X2.CUST_CODE" & vbCrLf _
            & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
        'sql &= " From RSTRETL2 X2, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
        '    & " where X2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
        '    & "   and TT_PROFA.CUST_CODE = X2.CUST_CODE" & vbCrLf _
        '    & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
        '    & "   and ICTCOLL1.COLLECTION_CODE = X2.COLLECTION_CODE" & vbCrLf _
        '    & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
        '    & " Group by " & vbCrLf _
        '    & "X2.CUST_CODE" & vbCrLf _
        '    & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
        '    & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

        Dim GLTPROFR As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFR & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

        ASCMAIN1.sql = "Select * from " & GLTPROFR
        Create_TDA(dst.Tables.Add, "GLTPROFR", "**", 0, False, "", 4)
        Fill_Records("GLTPROFR")

        Update_GLTPROFB(RI, "3", GLTPROFR, GLTPROFB, "RETAIL")


        '    ' General Ledger

        '    ASCMAIN1.Progress("GL Departmentals", "")

        '    sql = "Select '0' CB, X2.LINE_TAG, X2.CUST_CODE " & vbCrLf _
        '        & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
        '        & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
        '    YZ = "GL"
        '    ZZ = "NVL(X2.AMOUNT,0)"
        '    XZ = "X2"
        '    Make_SQL()

        '    sql &= " From GLTPROF5 X2, ICTBRAN1, SOTSDIV1" & vbCrLf _
        '        & " where X2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
        '        & "   and ICTBRAN1.ITEM_BRAND_CODE (+) = X2.ITEM_BRAND_CODE" & vbCrLf _
        '        & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
        '        & " Group by " & vbCrLf _
        '        & "X2.LINE_TAG, X2.CUST_CODE" & vbCrLf _
        '        & ", DECODE(ICTBRAN1.ALL_OTHERS_BRAND,'1','OTHERS',ICTBRAN1.ITEM_BRAND_CODE)" & vbCrLf _
        '        & ", ICTBRAN1.SALES_DIVISION_CODE, SOTSDIV1.BUS_UNIT_CODE" & vbCrLf
        '    Call Ora_to_Acc(Nothing, "GLTPROFG", 0, "", sql)
        '    Call Create_Index("GLTPROFG", "I_GLTPROFG_1", "CB, LINE_TAG")
        'CB = "2": z = "G": LINE_TAG = "ADVERT": GoSub Update_GLTPROFB   ' ADVERT only, not by Customer but yes by Brand
        '    'CB = "0": z = "G": LINE_TAG = "": GoSub Update_GLTPROFB         ' All Others, not by Customer and not by Brand
        'CB = "0": z = "G": LINE_TAG = "ADMIN": GoSub Update_GLTPROFB    ' All Others, not by Customer and not by Brand
        'CB = "0": z = "G": LINE_TAG = "DIST": GoSub Update_GLTPROFB     ' All Others, not by Customer and not by Brand
        'CB = "0": z = "G": LINE_TAG = "MKTG": GoSub Update_GLTPROFB     ' All Others, not by Customer and not by Brand
        'CB = "0": z = "G": LINE_TAG = "SLSADM": GoSub Update_GLTPROFB   ' All Others, not by Customer and not by Brand



        ' Fixed Demo

        'ASCMAIN1.Progress("Fixed Demo", "")


        'sql = "SELECT SUBSTR(GLTPROF0.OPS_YYYYPP,1,4) OPS_YYYY, GLTPROF0.CUST_CODE, SOTSDIV1.BUS_UNIT_CODE" & vbCr
        'sql = sql & ", SUM (GLTPROF0.ORDR_AMT_SHIP) GSLS" & vbCr
        'sql = sql & " FROM GLTPROF0,ICTBRAN1,SOTSDIV1 WHERE GLTPROF0.INV_TYPE = 'I'" & vbCr
        'sql = sql & " AND ICTBRAN1.ITEM_BRAND_CODE = GLTPROF0.ITEM_BRAND_CODE" & vbCr
        'sql = sql & " AND SOTSDIV1.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCr
        'sql = sql & " AND GLTPROF0.CUST_CODE IN (SELECT CUST_CODE FROM GLTPROF3)" & vbCr
        'sql = sql & " AND SOTSDIV1.BUS_UNIT_CODE IN ('PRES','LUX')" & vbCr
        'sql = sql & " AND GLTPROF0.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCr
        'sql = sql & " GROUP BY SUBSTR(GLTPROF0.OPS_YYYYPP,1,4), GLTPROF0.CUST_CODE, SOTSDIV1.BUS_UNIT_CODE" & vbCr
        'Dim TT_PROFD As String
        'TT_PROFD = Temp_Table(sql)



        '    sql = "Select X2.CUST_CODE, SOTBUSU1.BUS_UNIT_CODE " & vbCr

        '    YZ = "DEMFIX"
        '    '    zz = "NVL(X2.FIXED_DEMO,0)"
        '    ZZ = "(NVL(X2.FIXED_DEMO,0)/12) * XXX.PCT"
        '    XZ = "GLTPARM2"
        'GoSub Make_SQL

        '    sql = sql & " From GLTPROF3 X2, GLTPARM2, SOTBUSU1"
        '    sql = sql & ", (SELECT TT_PROFD.*, XX.GSLS GSLST, TT_PROFD.GSLS / XX.GSLS PCT FROM " & TT_PROFD & " TT_PROFD"
        '    sql = sql & " , (SELECT OPS_YYYY, CUST_CODE, SUM (GSLS) GSLS FROM " & TT_PROFD & " GROUP BY OPS_YYYY, CUST_CODE) XX"
        '    sql = sql & " Where TT_PROFD.OPS_YYYY = XX.OPS_YYYY"
        '    sql = sql & " AND TT_PROFD.CUST_CODE = XX.CUST_CODE) XXX"
        '    sql = sql & ", " & TT_PROFA & " TT_PROFA" & vbCr
        '    sql = sql & " where X2.OPS_YYYY BETWEEN '" & LY & "' AND '" & RY & "'" & vbCr
        '    sql = sql & "   and GLTPARM2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCr
        '    sql = sql & "   and X2.OPS_YYYY = SUBSTR(GLTPARM2.OPS_YYYYPP,1,4)" & vbCr
        '    sql = sql & "   and TT_PROFA.CUST_CODE = X2.CUST_CODE" & vbCr
        '    sql = sql & "   and TT_PROFA.BUS_UNIT_CODE = SOTBUSU1.BUS_UNIT_CODE" & vbCr
        '    sql = sql & "   and XXX.OPS_YYYY = X2.OPS_YYYY" & vbCr
        '    sql = sql & "   and XXX.CUST_CODE = X2.CUST_CODE" & vbCr
        '    sql = sql & "   and XXX.BUS_UNIT_CODE = SOTBUSU1.BUS_UNIT_CODE" & vbCr
        '    sql = sql & "   and SOTBUSU1.BUS_UNIT_CODE IN ('PRES','LUX')" & vbCr
        '    sql = sql & "   and X2.BA_TYPE = 'A'"
        '    sql = sql & " Group by " & vbCr
        '    sql = sql & "X2.CUST_CODE, SOTBUSU1.BUS_UNIT_CODE"
        '    Call Ora_to_Acc(Nothing, "GLTPROFF", 2, "", sql)

        '    '    sql = sql & " From GLTPROF3 X2, " & TT_PROFA & " TT_PROFA" & vbCr
        '    '    sql = sql & " where X2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCr
        '    '    sql = sql & "   and TT_PROFA.CUST_CODE = X2.CUST_CODE" & vbCr
        '    '    sql = sql & "   and X2.BA_TYPE = 'A'"
        '    '    sql = sql & " Group by " & vbCr
        '    '    sql = sql & "X2.CUST_CODE"
        '    '    Call Ora_to_Acc(Nothing, "GLTPROFF", 1, "", sql)

        'CB = "1": z = "F": LINE_TAG = "DEMFIX": GoSub Update_GLTPROFB











        '' Demo Comms

        'ASCMAIN1.Progress("Demo Commissions", "")

        'sql = "Select GLTPROFS.CUST_CODE " & vbCr
        'sql = sql & ", GLTPROFS.ITEM_BRAND_CODE" & vbCr
        'sql = sql & ", GLTPROFS.SALES_DIVISION_CODE" & vbCr
        'sql = sql & ", GLTPROFS.BUS_UNIT_CODE" & vbCr

        'sql = sql & ", SUM (GLTPROFS.AMT_LY_MTL * GLTPROF6X.DC_PCT_LY) as DCM_LY_MTL" & vbCr
        'sql = sql & ", SUM (GLTPROFS.AMT_LY_YTL * GLTPROF6X.DC_PCT_LY) as DCM_LY_YTL" & vbCr
        'sql = sql & ", SUM (GLTPROFS.AMT_LY_YTD * GLTPROF6X.DC_PCT_LY) as DCM_LY_YTD" & vbCr
        'sql = sql & ", SUM (GLTPROFS.AMT_TY_MTD * GLTPROF6X.DC_PCT_TY) as DCM_TY_MTD" & vbCr
        'sql = sql & ", SUM (GLTPROFS.AMT_TY_YTD * GLTPROF6X.DC_PCT_TY) as DCM_TY_YTD" & vbCr

        'sql = sql & " into GLTPROFU" & vbCr
        'sql = sql & " from GLTPROFS, GLTPROF6X" & vbCr

        'sql = sql & " where GLTPROFS.CUST_CODE = GLTPROF6X.CUST_CODE" & vbCr
        'sql = sql & " Group by " & vbCr
        'sql = sql & "GLTPROFS.CUST_CODE" & vbCr
        'sql = sql & ", GLTPROFS.ITEM_BRAND_CODE" & vbCr
        'sql = sql & ", GLTPROFS.SALES_DIVISION_CODE" & vbCr
        'sql = sql & ", GLTPROFS.BUS_UNIT_CODE" & vbCr

        'AccD.Execute sql

        'Call Create_Index("GLTPROFU", "PrimaryKey", "CUST_CODE,ITEM_BRAND_CODE,SALES_DIVISION_CODE,BUS_UNIT_CODE")

        'YZ = "DCM"
        'Update_GLTPROFB("3", "U", "DEMCOM")





        ASCMAIN1.Progress("Demo Commissions", "")

        YZ = "DCM"
        If RI.Contains(YZ) Then
            sql = "Select SPTDCOMB.CUST_CODE " & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE) COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf
            ZZ = "AMT_COMM"
            XZ = "SPTDCOMB"
            Make_SQL()
            sql &= " From SPTDCOMB, SPTDCOM1, ICTCOLL1, ICTBRAN1, " & TT_PROFA & " TT_PROFA" & vbCrLf _
                & " where SPTDCOMB.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCrLf _
                & "   and TT_PROFA.CUST_CODE = SPTDCOMB.CUST_CODE" & vbCrLf _
                & "   and TT_PROFA.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and ICTCOLL1.COLLECTION_CODE = SPTDCOMB.COLLECTION_CODE" & vbCrLf _
                & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
                & "   and SPTDCOM1.CUST_CODE = SPTDCOMB.CUST_CODE" & vbCrLf _
                & "   and " & XZ & "." & ZZ & " <> 0" & vbCrLf _
                & " Group by " & vbCrLf _
                & "SPTDCOMB.CUST_CODE" & vbCrLf _
                & ", DECODE(ICTCOLL1.ALL_OTHERS,'1','OTHERS',ICTCOLL1.COLLECTION_CODE)" & vbCrLf _
                & ", ICTCOLL1.BRAND_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf

            Dim GLTPROFD As String = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Alter Table " & GLTPROFD & " Add Primary Key (CUST_CODE,COLLECTION_CODE,BRAND_CODE,SALES_DIVISION_CODE)")

            ASCMAIN1.sql = "Select * from " & GLTPROFD
            Create_TDA(dst.Tables.Add, "GLTPROFD", "**", 0, False, "", 4)
            Fill_Records("GLTPROFD")

            Update_GLTPROFB(RI, "3", GLTPROFD, GLTPROFB, YZ)
        End If





        '    ' AE

        '    ASCMAIN1.Progress("AE Expense", "")

        '    sql = "Select X2.CUST_CODE " & vbCr
        '    sql = sql & ", (CASE WHEN X2.OPS_YYYYPP >= '200607' THEN NVL(SOTSREP1.BUS_UNIT_CODE,BU_SLS.BUS_UNIT_CODE) ELSE BU_SLS.BUS_UNIT_CODE END) BUS_UNIT_CODE" & vbCr
        '    '    sql = sql & ", DECODE(ICTBRAN1.ALL_OTHERS_BRAND,'1','OTHERS',ICTBRAN1.ITEM_BRAND_CODE) ITEM_BRAND_CODE" & vbCr
        '    '    sql = sql & ", ICTBRAN1.SALES_DIVISION_CODE, SOTSDIV1.BUS_UNIT_CODE" & vbCr

        '    ZZ = "NVL(X2.OVERHEAD_AMT,0)"
        '    ZZ = ZZ & " * BU_SLS.GSLS / BUALL_SLS.GSLS"

        '    YZ = "AE"
        '    XZ = "X2"
        'GoSub Make_SQL

        '    '    sql = sql & " From GLTPROF4 X2, ICTBRAN1, SOTSDIV1, " & TT_PROFA & " TT_PROFA" & vbCr
        '    sql = sql & " From GLTPROF4 X2, SOTSREP1, " & TT_PROFA & " TT_PROFA" & vbCr
        '    sql = sql & ", " & BU_SLS & " BU_SLS, " & BUALL_SLS & " BUALL_SLS" & vbCr

        '    sql = sql & " where X2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCr
        '    sql = sql & "   and X2.BA_TYPE = 'A'"
        '    '    sql = sql & "   and X2.ITEM_BRAND_CODE = ICTBRAN1.ITEM_BRAND_CODE"
        '    '    sql = sql & "   and SOTSDIV1.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE"

        '    sql = sql & "   and SOTSREP1.SREP_CODE (+) = X2.SREP_CODE" & vbCr
        '    sql = sql & "   and TT_PROFA.CUST_CODE = X2.CUST_CODE" & vbCr
        '    '    sql = sql & "   and TT_PROFA.BUS_UNIT_CODE = SOTSDIV1.BUS_UNIT_CODE" & vbCr
        '    sql = sql & "   and TT_PROFA.BUS_UNIT_CODE = (CASE WHEN X2.OPS_YYYYPP >= '200607' THEN NVL(SOTSREP1.BUS_UNIT_CODE,BU_SLS.BUS_UNIT_CODE) ELSE BU_SLS.BUS_UNIT_CODE END)" & vbCr

        '    sql = sql & "   and BU_SLS.OPS_YYYYPP = X2.OPS_YYYYPP" & vbCr
        '    sql = sql & "   and BU_SLS.CUST_CODE = X2.CUST_CODE" & vbCr
        '    sql = sql & "   and BUALL_SLS.OPS_YYYYPP = X2.OPS_YYYYPP" & vbCr
        '    sql = sql & "   and BUALL_SLS.CUST_CODE = X2.CUST_CODE" & vbCr

        '    sql = sql & "   and X2.REC_TYPE = 'AE'" & vbCr
        '    sql = sql & " Group by " & vbCr
        '    sql = sql & "X2.CUST_CODE"
        '    sql = sql & ", (CASE WHEN X2.OPS_YYYYPP >= '200607' THEN NVL(SOTSREP1.BUS_UNIT_CODE,BU_SLS.BUS_UNIT_CODE) ELSE BU_SLS.BUS_UNIT_CODE END)" & vbCr
        '    '    sql = sql & ", DECODE(ICTBRAN1.ALL_OTHERS_BRAND,'1','OTHERS',ICTBRAN1.ITEM_BRAND_CODE)" & vbCr
        '    '    sql = sql & ", ICTBRAN1.SALES_DIVISION_CODE, SOTSDIV1.BUS_UNIT_CODE" & vbCr
        '    '    Call Ora_to_Acc(Nothing, "GLTPROFE", 4, "", sql)
        '    Call Ora_to_Acc(Nothing, "GLTPROFE", 2, "", sql)

        '    '    CB = "3": z = "E": LINE_TAG = "AE": GoSub Update_GLTPROFB
        'CB = "1": z = "E": LINE_TAG = "AE": GoSub Update_GLTPROFB

        '    ' RAM

        '    ASCMAIN1.Progress("RAM Expense", "")

        '    sql = "Select X2.CUST_CODE " & vbCr
        '    sql = sql & ", (CASE WHEN X2.OPS_YYYYPP >= '200607' THEN NVL(SOTSREP1.BUS_UNIT_CODE,BU_SLS.BUS_UNIT_CODE) ELSE BU_SLS.BUS_UNIT_CODE END) BUS_UNIT_CODE" & vbCr
        '    '    sql = sql & ", DECODE(ICTBRAN1.ALL_OTHERS_BRAND,'1','OTHERS',ICTBRAN1.ITEM_BRAND_CODE) ITEM_BRAND_CODE" & vbCr
        '    '    sql = sql & ", ICTBRAN1.SALES_DIVISION_CODE, SOTSDIV1.BUS_UNIT_CODE" & vbCr

        '    ZZ = "NVL(X2.OVERHEAD_AMT,0)"
        '    ZZ = ZZ & " * BU_SLS.GSLS / BUALL_SLS.GSLS"

        '    YZ = "RAM"
        '    XZ = "X2"
        'GoSub Make_SQL

        '    '    sql = sql & " From GLTPROF4 X2, ICTBRAN1, SOTSDIV1, " & TT_PROFA & " TT_PROFA" & vbCr
        '    sql = sql & " From GLTPROF4 X2, SOTSREP1, " & TT_PROFA & " TT_PROFA" & vbCr
        '    sql = sql & ", " & BU_SLS & " BU_SLS, " & BUALL_SLS & " BUALL_SLS" & vbCr

        '    sql = sql & " where X2.OPS_YYYYPP BETWEEN '" & LY & "01' AND '" & RY & "12'" & vbCr
        '    sql = sql & "   and X2.BA_TYPE = 'A'"
        '    '    sql = sql & "   and X2.ITEM_BRAND_CODE = ICTBRAN1.ITEM_BRAND_CODE"
        '    '    sql = sql & "   and SOTSDIV1.SALES_DIVISION_CODE = ICTBRAN1.SALES_DIVISION_CODE"

        '    sql = sql & "   and SOTSREP1.SREP_CODE (+) = X2.SREP_CODE" & vbCr
        '    sql = sql & "   and TT_PROFA.CUST_CODE = X2.CUST_CODE" & vbCr
        '    '    sql = sql & "   and TT_PROFA.BUS_UNIT_CODE = SOTSDIV1.BUS_UNIT_CODE" & vbCr
        '    sql = sql & "   and TT_PROFA.BUS_UNIT_CODE = (CASE WHEN X2.OPS_YYYYPP >= '200607' THEN NVL(SOTSREP1.BUS_UNIT_CODE,BU_SLS.BUS_UNIT_CODE) ELSE BU_SLS.BUS_UNIT_CODE END)" & vbCr

        '    sql = sql & "   and BU_SLS.OPS_YYYYPP = X2.OPS_YYYYPP" & vbCr
        '    sql = sql & "   and BU_SLS.CUST_CODE = X2.CUST_CODE" & vbCr
        '    sql = sql & "   and BUALL_SLS.OPS_YYYYPP = X2.OPS_YYYYPP" & vbCr
        '    sql = sql & "   and BUALL_SLS.CUST_CODE = X2.CUST_CODE" & vbCr

        '    sql = sql & "   and X2.REC_TYPE = 'RAM'" & vbCr
        '    sql = sql & " Group by " & vbCr
        '    sql = sql & "X2.CUST_CODE"
        '    sql = sql & ", (CASE WHEN X2.OPS_YYYYPP >= '200607' THEN NVL(SOTSREP1.BUS_UNIT_CODE,BU_SLS.BUS_UNIT_CODE) ELSE BU_SLS.BUS_UNIT_CODE END)" & vbCr
        '    '    sql = sql & ", DECODE(ICTBRAN1.ALL_OTHERS_BRAND,'1','OTHERS',ICTBRAN1.ITEM_BRAND_CODE)" & vbCr
        '    '    sql = sql & ", ICTBRAN1.SALES_DIVISION_CODE, SOTSDIV1.BUS_UNIT_CODE" & vbCr
        '    '    Call Ora_to_Acc(Nothing, "GLTPROFT", 4, "", sql)
        '    Call Ora_to_Acc(Nothing, "GLTPROFT", 2, "", sql)

        '    '    CB = "3": z = "T": LINE_TAG = "RAM": GoSub Update_GLTPROFB
        'CB = "1": z = "T": LINE_TAG = "RAM": GoSub Update_GLTPROFB

        ' Prepare Summary

        ' MAYBE THIS SHOULD COME FROM PROFA?

        '    sql = "Select ARTSREP1.CUST_CODE, ICTBRAN1.ITEM_BRAND_CODE, ARTSREP1.SREP_CODE, SOTSREP1.REGION_CODE, SOTSREG1.VP_CODE"
        '    sql = sql & ", ICTBRAN1.SALES_DIVISION_CODE, ARTSREP1.BUS_UNIT_CODE"
        '    sql = sql & " from ARTSREP1,ICTBRAN1,SOTSDIV1,SOTSREP1,SOTSREG1 "
        '    sql = sql & " where SOTSREG1.REGION_CODE (+) = SOTSREP1.REGION_CODE"
        '    sql = sql & "   and SOTSREP1.SREP_CODE (+) = ARTSREP1.SREP_CODE"
        '    sql = sql & "   and SOTSDIV1.BUS_UNIT_CODE = ARTSREP1.BUS_UNIT_CODE"
        '    sql = sql & "   and ICTBRAN1.SALES_DIVISION_CODE = SOTSDIV1.SALES_DIVISION_CODE"
        '    Call Ora_to_Acc(Nothing, "ARWSREPX", 2, "", sql)


        'Call Create_Index("GLTPROFZ", "I_GLTPROFZ_1", _
        '    "OPS_YYYYPP,CUST_CODE,ITEM_BRAND_CODE,VP_CODE,REGION_CODE,SREP_CODE,SALES_DIVISION_CODE,LINE_NO")

        Dim sql_COLS As String
        Dim sql_non_0 As String
        sql_COLS = ""
        sql_non_0 = ""
        For i As Integer = 1 To PROFBn
            sql_COLS = sql_COLS & ", SUM (GLTPROFB." & PROFBns(i) & ") AS " & PROFBns(i)
            sql_non_0 = sql_non_0 & " or GLTPROFB." & PROFBns(i) & " <> 0"
        Next i

        ASCMAIN1.sql = "" & vbCrLf _
            & " Select '" & RYP & "' as OPS_YYYYPP" & vbCrLf _
            & ", Decode(GLTPROFA.ALL_OTHERS, '1', 'OTHERS', GLTPROFB.CUST_CODE) as CUST_CODE" & vbCrLf _
            & ", GLTPROFB.COLLECTION_CODE" & vbCrLf _
            & ", GLTPROFA.VP_CODE" & vbCrLf _
            & ", GLTPROFA.REGION_CODE" & vbCrLf _
            & ", GLTPROFA.SREP_CODE" & vbCrLf _
            & ", GLTPROFB.BRAND_CODE" & vbCrLf _
            & ", GLTPROFB.LINE_NO" & vbCrLf _
            & sql_COLS & vbCrLf _
            & " From " & GLTPROFB & " GLTPROFB, " & TT_PROFA & " GLTPROFA" & vbCrLf _
            & " Where GLTPROFB.CUST_CODE = GLTPROFA.CUST_CODE" & vbCrLf _
            & " and GLTPROFB.SALES_DIVISION_CODE = GLTPROFA.SALES_DIVISION_CODE" & vbCrLf _
            & " and (" & Mid$(sql_non_0, 5) & ")" & vbCrLf _
            & " Group by" & vbCrLf _
            & "  Decode(GLTPROFA.ALL_OTHERS, '1', 'OTHERS', GLTPROFB.CUST_CODE)" & vbCrLf _
            & ", GLTPROFB.COLLECTION_CODE" & vbCrLf _
            & ", GLTPROFA.VP_CODE" & vbCrLf _
            & ", GLTPROFA.REGION_CODE" & vbCrLf _
            & ", GLTPROFA.SREP_CODE" & vbCrLf _
            & ", GLTPROFB.BRAND_CODE" & vbCrLf _
            & ", GLTPROFB.LINE_NO"
        Dim GLTPROFZ As String = ASCMAIN1.Temp_Table


        For Each C As String In New String() {"COLLECTION_CODE", "VP_CODE", "REGION_CODE", "SREP_CODE", "BRAND_CODE"}
            ASCMAIN1.sql = "Update " & GLTPROFZ & " Set " & C & " = NVL(" & C & ",'?')"
            ASCDATA1.ExecuteSQL()
        Next


        Create_TDA(dst.Tables.Add, "GLTPROFZ", "*", 0, True, "", 8)

        ASCMAIN1.sql = "Select * from " & GLTPROFZ
        Fill_Records("GLTPROFZ", "", True, ASCMAIN1.sql)

        ' Do Calculations

        ASCMAIN1.Progress("Calculations", "")

        Dim t(,) As Decimal
        ReDim t(LINE_NO_max, PROFBn)
        Dim KY(9) As String

        ASCMAIN1.sql = "Select Distinct OPS_YYYYPP,CUST_CODE,COLLECTION_CODE,VP_CODE,REGION_CODE,SREP_CODE,BRAND_CODE from GLTPROFZ"
        Dim tbl As DataTable = ASCDATA1.GetDataTable
        For Each row As DataRow In tbl.Select("")
            sql = ""
            For i As Integer = 1 To 7
                KY(i) = row.Item(i - 1) & ""
                Dim z As String = tbl.Columns(i - 1).ColumnName
                If KY(i) = "" Then
                    sql = sql & " and " & z & " is Null"
                Else
                    sql = sql & " and " & z & " = '" & KY(i) & "'"
                End If
            Next i

            ReDim t(LINE_NO_max, PROFBn)
            For Each rowGLTPROFZ As DataRow In dst.Tables("GLTPROFZ").Select(Mid(sql, 5))
                Dim LINE_NO As Integer = Val(rowGLTPROFZ.Item("LINE_NO") & "")
                For k As Integer = 1 To PROFBn
                    t(LINE_NO, k) = Val(rowGLTPROFZ.Item(k - 1 + 8) & "")
                Next k
            Next
            'ASCMAIN1.sql = "Select * from GLTPROFZ " & ASCMAIN1.SQL_Add_WHERE(sql)
            'For Each rowGLTPROFZ As DataRow In ASCDATA1.GetDataTable.Select("", "LINE_NO")
            '    Dim LINE_NO As Integer = Val(rowGLTPROFZ.Item("LINE_NO") & "")
            '    For k As Integer = 1 To PROFBn
            '        t(LINE_NO, k) = Val(rowGLTPROFZ.Item(k - 1 + 8) & "")
            '    Next k
            'Next


            For k As Integer = 1 To PROFBn
                t(RI("GSLS"), k) = t(RI("GSLSB"), k) + t(RI("GSLSP"), k)
                t(RI("NSLS"), k) = t(RI("GSLS"), k) + t(RI("RSLS"), k)
                t(RI("GCGS"), k) = t(RI("GCGSB"), k) + t(RI("GCGSP"), k)
                t(RI("TCGS"), k) = t(RI("GCGS"), k) + t(RI("ROYALT"), k) + t(RI("RCGS"), k)
                t(RI("GP"), k) = t(RI("NSLS"), k) - t(RI("TCGS"), k)
                t(RI("ASP"), k) = t(RI("PROMO"), k) + t(RI("COMM"), k) + t(RI("FIX"), k) + t(RI("DIF"), k) + t(RI("TST"), k) + t(RI("ARDED"), k) - t(RI("MISC"), k) + t(RI("ADVERT"), k)
                t(RI("NC"), k) = t(RI("SAMPLE"), k) + t(RI("GWP"), k) + t(RI("DSP"), k)
                t(RI("TAXFEE"), k) = t(RI("MODROT"), k) * 0.17
                t(RI("TOTASP"), k) = t(RI("ASP"), k) + t(RI("NC"), k) + t(RI("DEMFIX"), k) + t(RI("MODROT"), k) + t(RI("TAXFEE"), k)
                't(RI("DEMCOM"), k) = t(RI("NSLS"), k) * 0.092
                t(RI("ACTMAR"), k) = t(RI("GP"), k) - t(RI("TOTASP"), k) - t(RI("DEMCOM"), k) - t(RI("AE"), k) - t(RI("RAM"), k)
                t(RI("OVRTOT"), k) = t(RI("DIST"), k) + t(RI("MKTG"), k) + t(RI("SLSADM"), k) + t(RI("ADMIN"), k)
                t(RI("MIS"), k) = t(RI("ACTMAR"), k) - t(RI("OVRTOT"), k)

                '    If ASCMAIN1.Running_in_VS And KY(1) = "201603" And KY(2) = "ULTA" And t(RI("GSLS"), 9) <> 0 Then Stop

                If t(RI("NSLS"), k) = 0 Then
                    t(RI("STHSIN"), k) = 0
                Else
                    t(RI("STHSIN"), k) = 100 * t(RI("RETAIL"), k) * 0.6 / t(RI("NSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("BASGRS"), k) = 0
                Else
                    t(RI("BASGRS"), k) = 100 * t(RI("GSLSB"), k) / t(RI("GSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("PROGRS"), k) = 0
                Else
                    t(RI("PROGRS"), k) = 100 * t(RI("GSLSP"), k) / t(RI("GSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("RTGSHP"), k) = 0
                Else
                    t(RI("RTGSHP"), k) = 100 * t(RI("RSLS"), k) / t(RI("GSLS"), k)
                End If
                If t(RI("GSLS"), k) = 0 Then
                    t(RI("CGSSHP"), k) = 0
                Else
                    t(RI("CGSSHP"), k) = 100 * t(RI("GCGS"), k) / t(RI("GSLS"), k)
                End If
                If t(RI("NSLS"), k) = 0 Then
                    t(RI("ASPSHP"), k) = 0
                Else
                    t(RI("ASPSHP"), k) = 100 * t(RI("TOTASP"), k) / t(RI("NSLS"), k)
                End If
                If t(RI("NSLS"), k) = 0 Then
                    t(RI("AMNSHP"), k) = 0
                Else
                    t(RI("AMNSHP"), k) = 100 * t(RI("ACTMAR"), k) / t(RI("NSLS"), k)
                End If

            Next k

            Write_Record(t, PROFBn, KY, "GSLS", RI("GSLS"))
            Write_Record(t, PROFBn, KY, "NSLS", RI("NSLS"))
            Write_Record(t, PROFBn, KY, "GCGS", RI("GCGS"))
            Write_Record(t, PROFBn, KY, "TCGS", RI("TCGS"))
            Write_Record(t, PROFBn, KY, "GP", RI("GP"))
            Write_Record(t, PROFBn, KY, "NC", RI("NC"))
            Write_Record(t, PROFBn, KY, "ASP", RI("ASP"))
            Write_Record(t, PROFBn, KY, "TAXFEE", RI("TAXFEE"))
            Write_Record(t, PROFBn, KY, "TOTASP", RI("TOTASP"))
            Write_Record(t, PROFBn, KY, "DEMCOM", RI("DEMCOM"))
            Write_Record(t, PROFBn, KY, "ACTMAR", RI("ACTMAR"))
            Write_Record(t, PROFBn, KY, "OVRTOT", RI("OVRTOT"))
            Write_Record(t, PROFBn, KY, "MIS", RI("MIS"))
            Write_Record(t, PROFBn, KY, "STHSIN", RI("STHSIN"))
            Write_Record(t, PROFBn, KY, "BASGRS", RI("BASGRS"))
            Write_Record(t, PROFBn, KY, "PROGRS", RI("PROGRS"))
            Write_Record(t, PROFBn, KY, "RTGSHP", RI("RTGSHP"))
            Write_Record(t, PROFBn, KY, "CGSSHP", RI("CGSSHP"))
            Write_Record(t, PROFBn, KY, "ASPSHP", RI("ASPSHP"))
            Write_Record(t, PROFBn, KY, "AMNSHP", RI("AMNSHP"))
        Next


        ' Upload to Server

        ASCMAIN1.Progress("Upload to Server", "")

        For Each rowGLTPROFZ As DataRow In dst.Tables("GLTPROFZ").Select("LINE_TAG IS NULL")
            Dim rowGLTPROF1 As DataRow = dst.Tables("GLTPROF1").Rows.Find(rowGLTPROFZ.Item("LINE_NO")) '  LookUp("GLTPROF1", rowGLTPROFZ.Item("LINE_NO"))
            rowGLTPROFZ.Item("LINE_TAG") = rowGLTPROF1.Item("LINE_TAG")
        Next
        ASCDATA1.DeleteRows("GLTPROFZ", "ACT_LY_YTL=0 AND ACT_LY_YTD=0 AND ACT_LY_MTL=0 AND BUD_TY_YTL=0 AND BUD_TY_YTD=0 AND BUD_TY_MTL=0 AND BUD_TY_YTG=0 AND ACT_TY_YTD=0 AND ACT_TY_MTD=0")

        Update_Record_TDA("GLTPROFZ", "OPS_YYYYPP = '" & RYP & "'")

        ASCMAIN1.sql = "Delete from GLTPROFW where OPS_YYYYPP = '" & RYP & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Insert into GLTPROFW (OPS_YYYYPP,RECORDS,INIT_DATE,INIT_OPER,GROSS_SALES,NET_SALES,GROSS_PROFIT,CUST_COUNT)" & vbCrLf _
            & "Select OPS_YYYYPP, RECORDS, SYSDATE INIT_DATE, :PARM1 INIT_OPER" _
            & ", GROSS_SALES, NET_SALES, GROSS_PROFIT, CUST_COUNT" _
            & " from (Select OPS_YYYYPP, COUNT (*) RECORDS" _
            & ", Sum (CASE WHEN LINE_TAG IN ('GSLSB','GSLSP') THEN ACT_TY_MTD ELSE 0 END) GROSS_SALES" _
            & ", Sum (CASE WHEN LINE_TAG IN ('GSLSB','GSLSP','RSLS') THEN ACT_TY_MTD ELSE 0 END) NET_SALES" _
            & ", Sum (CASE WHEN LINE_TAG <> 'RETAIL'  THEN ACT_TY_MTD ELSE 0 END) GROSS_PROFIT" _
            & ", Count (Distinct CUST_CODE) CUST_COUNT" _
            & " from GLTPROFZ where OPS_YYYYPP = '" & RYP & "' group by OPS_YYYYPP)"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "V", New String() {ASCMAIN1.USER_ID})

        If chkPREP_BI.Checked Then
            BeginTrans()
            ASCMAIN1.sql = "Delete from TATQVXQ1 where RECORD_STATUS = '0' AND YYYY = '" & Mid(RYP, 1, 4) & "' and PP = '" & Mid(RYP, 5, 2) & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Insert into TATQVXQ1 Select * from " & TATQVXQ1
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update GLTPROFW Set BI_PREP_OPER = INIT_OPER, BI_PREP_DATE = INIT_DATE WHERE OPS_YYYYPP = '" & RYP & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update GLTPROFW SET BI_PREP_TOTAL = " & vbCrLf _
                & " (Select Sum (AMOUNT) from TATQVXQ1 where YYYY = '" & Mid(RYP, 1, 4) & "' and PP = '" & Mid(RYP, 5, 2) & "'" & vbCrLf _
                & "  and RECORD_STATUS = '0' and NVL(IGNORE,'N') <> 'Y') where OPS_YYYYPP = '" & RYP & "'"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.Record_Event("TATQVXQ1", RYP, "", Now, ASCMAIN1.USER_ID, "PEWP", "BI Data Pewp for " & RYPLEGEND, ASCMAIN1.SESSION_NO)

            CommitTrans()
        End If

        Get_Stats()

    End Sub

    Sub Update_GLTPROFB(RI As Collection, CB As String, TABLE_NAME As String, GLTPROFB As String, LINE_TAG As String)

        ' GLTPROFP has 3 sets of records; all records are pcts to a total
        ' CB = 0 : % of customer + brand to the total
        ' CB = 1 : % of brand to the customer total
        ' CB = 2 : % of customer to the brand total
        ' using CB = 0/1/2 means that we need to take an amount and pro-rate it
        ' using CB = 3 means that you do not need to pro-rate
        ' use CB = 0 if you have an amount, and need to spread it by Customer + Brand
        ' use CB = 1 if you have an amount by Customer and need to spread it by Brand
        ' use CB = 2 if you have an amount by Brand and need to spread it by Customer
        ' use CB = 3 if you have amounts already spread by Customer & Brand

        sql = "Select 0 as RECORD_NO, " & CStr(RI(LINE_TAG)) & " as LINE_NO, '" & LINE_TAG & "' as LINE_TAG" & vbCrLf _
            & ", GLTPROFP.CUST_CODE" & vbCrLf _
            & ", GLTPROFP.COLLECTION_CODE" & vbCrLf _
            & ", GLTPROFP.BRAND_CODE" & vbCrLf _
            & ", GLTPROFP.SALES_DIVISION_CODE"
        If CB = "3" Then
            sql = Replace$(sql, "GLTPROFP.", "GLTPROFX.")
        End If
        sql &= ", GLTPROFX." & YZ & "_LY_MTL" & IIf(CB = "3", "", " * GLTPROFP.AMT_LY_MTL") & vbCrLf _
            & ", GLTPROFX." & YZ & "_LY_YTL" & IIf(CB = "3", "", " * GLTPROFP.AMT_LY_YTL") & vbCrLf _
            & ", GLTPROFX." & YZ & "_LY_YTD" & IIf(CB = "3", "", " * GLTPROFP.AMT_LY_YTD") & vbCrLf _
            & ", GLTPROFX." & YZ & "_TY_MTD" & IIf(CB = "3", "", " * GLTPROFP.AMT_TY_MTD") & vbCrLf _
            & ", GLTPROFX." & YZ & "_TY_YTD" & IIf(CB = "3", "", " * GLTPROFP.AMT_TY_YTD") & vbCrLf _
            & " from " & TABLE_NAME & "  GLTPROFX" & vbCrLf
        If CB <> "3" Then
            sql = sql & ", " & GLTPROFB & " GLTPROFP" & vbCrLf
        End If

        If CB = "0" Then
            sql &= " where GLTPROFX.LINE_TAG = '" & LINE_TAG & "'" & vbCrLf
        End If
        If CB = "1" Then
            'sql &= ", GLTPROFA" & vbCrLf _
            '    & " where GLTPROFA.CUST_CODE = GLTPROFX.CUST_CODE" & vbCrLf _
            '    & "   and GLTPROFP.CUST_CODE = GLTPROFX.CUST_CODE" & vbCrLf _
            '    & "   and GLTPROFA.SALES_DIVISION_CODE = GLTPROFX.SALES_DIVISION_CODE" & vbCrLf _
            '    & "   and GLTPROFP.SALES_DIVISION_CODE = GLTPROFX.SALES_DIVISION_CODE" & vbCrLf
            sql &= "" & vbCrLf _
                & " where GLTPROFP.CUST_CODE = GLTPROFX.CUST_CODE" & vbCrLf _
                & "   and GLTPROFP.SALES_DIVISION_CODE = GLTPROFX.SALES_DIVISION_CODE" & vbCrLf
        End If
        If CB = "2" Then
            sql &= ", GLTPROFY" & vbCrLf _
                & " where GLTPROFY.COLLECTION_CODE = GLTPROFX.COLLECTION_CODE" & vbCrLf _
                & "   and GLTPROFY.BRAND_CODE = GLTPROFX.BRAND_CODE" & vbCrLf _
                & "   and GLTPROFP.COLLECTION_CODE = GLTPROFX.COLLECTION_CODE" & vbCrLf _
                & "   and GLTPROFP.BRAND_CODE = GLTPROFX.BRAND_CODE" & vbCrLf
            If TABLE_NAME = "GLTPROFG" Then
                sql = sql & " and GLTPROFX.LINE_TAG = '" & LINE_TAG & "'" & vbCrLf
            End If
        End If

        If CB <> "3" Then
            sql = sql & "   and GLTPROFP.CB = '" & CB & "'" & vbCrLf
        End If

        sql = "Insert into " & GLTPROFB & " (RECORD_NO, LINE_NO, LINE_TAG, CUST_CODE, COLLECTION_CODE, BRAND_CODE, SALES_DIVISION_CODE, ACT_LY_MTL, ACT_LY_YTL, ACT_LY_YTD, ACT_TY_MTD, ACT_TY_YTD) " & sql
        ASCDATA1.ExecuteSQL(sql)

    End Sub
    Sub Write_Record(T(,) As Decimal, PROFBn As Integer, KY() As String, LINE_TAG As String, LINE_NO As Integer)
        Dim rowGLTPROFZ As DataRow = dst.Tables("GLTPROFZ").Rows.Find(New Object() {KY(1), KY(2), KY(3), KY(4), KY(5), KY(6), KY(7), LINE_NO})
        If rowGLTPROFZ Is Nothing Then
            rowGLTPROFZ = dst.Tables("GLTPROFZ").NewRow
            For j As Integer = 1 To 7
                rowGLTPROFZ.Item(j - 1) = KY(j)
            Next j
            rowGLTPROFZ.Item(7) = LINE_NO
            'rowGLTPROFZ.Item("LINE_TAG") = LINE_TAG
            dst.Tables("GLTPROFZ").Rows.Add(rowGLTPROFZ)
        End If
        '   If ASCMAIN1.Running_in_VS And LINE_NO = 3 Then Stop
        For k As Integer = 1 To PROFBn
            rowGLTPROFZ.Item(k + 7) = Val(rowGLTPROFZ.Item(k + 7) & "") + T(LINE_NO, k)
        Next k
    End Sub
    Sub Make_SQL(Optional YP As String = "OPS_YYYYPP")
        sql &= ", SUM (DECODE(" & XZ & "." & YP & ",'" & LY & RP & "'," & ZZ & ",0)) " & YZ & "_LY_MTL" & vbCrLf
        sql &= ", SUM ((CASE WHEN " & XZ & "." & YP & " BETWEEN '" & LY & "01' AND '" & LY & "12' THEN " & ZZ & " ELSE 0 END)) " & YZ & "_LY_YTL" & vbCrLf
        sql &= ", SUM ((CASE WHEN " & XZ & "." & YP & " BETWEEN '" & LY & "01' AND '" & LYP & "' THEN " & ZZ & " ELSE 0 END)) " & YZ & "_LY_YTD" & vbCrLf
        sql &= ", SUM (DECODE(" & XZ & "." & YP & ",'" & RYP & "'," & ZZ & ",0)) " & YZ & "_TY_MTD" & vbCrLf
        sql &= ", SUM ((CASE WHEN " & XZ & "." & YP & " BETWEEN '" & RY & "01' AND '" & RYP & "' THEN " & ZZ & " ELSE 0 END)) " & YZ & "_TY_YTD" & vbCrLf
    End Sub

    Sub Insert_Sales(Z As String, sqlw2 As String, GLTPROFB As String, GLTPROFS As String, GLTPROF1 As String)

        Dim sql As String = "Select 0 as RECORD_NO, GLTPROF1.LINE_NO, GLTPROF1.LINE_TAG" & vbCrLf _
            & ", GLTPROFS.CUST_CODE" & vbCrLf _
            & ", GLTPROFS.COLLECTION_CODE" & vbCrLf _
            & ", GLTPROFS.BRAND_CODE" & vbCrLf _
            & ", GLTPROFS.SALES_DIVISION_CODE" & vbCrLf _
            & ", GLTPROFS." & Z & "_LY_MTL" & vbCrLf _
            & ", GLTPROFS." & Z & "_LY_YTL" & vbCrLf _
            & ", GLTPROFS." & Z & "_LY_YTD" & vbCrLf _
            & ", GLTPROFS." & Z & "_TY_MTD" & vbCrLf _
            & ", GLTPROFS." & Z & "_TY_YTD" & vbCrLf _
            & " from " & GLTPROFS & " GLTPROFS," & GLTPROF1 & " GLTPROF1" & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sqlw2) & vbCrLf

        sql = "Insert into " & GLTPROFB _
            & " (RECORD_NO, LINE_NO, LINE_TAG, CUST_CODE, COLLECTION_CODE, BRAND_CODE, SALES_DIVISION_CODE, ACT_LY_MTL, ACT_LY_YTL, ACT_LY_YTD, ACT_TY_MTD, ACT_TY_YTD) " _
            & sql
        ASCDATA1.ExecuteSQL(sql)

    End Sub
    Public Overrides Sub Print_Report()
        'SUBT = ""
        'CR_params.Add("SUBT", SUBT)
        'Generate_Report(RPT, , SUBT)

        ' KICK OUT THE EXCEL SPREADSHEET
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        'ASCMAIN1.sql = "Insert into SPTDCOMB Select * from "
        'ASCDATA1.ExecuteSQL(sql)

    End Sub


    Sub Update_Sales_Summary(yp As String)

        ''''Stop ' MAKE SURE THAT THIS SECTION IS IDENTICAL TO TAXPEND1.Update_Sales_Summary

        ASCMAIN1.sql = "Delete from GLTPROF0 where OPS_YYYYPP = '" & yp & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Delete from GLTPROFS where OPS_YYYYPP = '" & yp & "'"
        ASCDATA1.ExecuteSQL()

        '       & " , NULL SREP_CODE, NULL SELL_CODE"

        ASCMAIN1.sql = "Insert into GLTPROFS" & vbCrLf _
            & " SELECT SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, " & vbCrLf _
            & " ICTITEM1.COLLECTION_CODE, SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP," & vbCrLf _
            & " SOTINVH2.INV_TYPE, ICTITEM1.ITEM_BASIC_PROMO," & vbCrLf _
            & " DECODE(NVL(ICTPROD1.PROD_CATGY,'0'),'0',DECODE(ICTITEM1.ITEM_SNU_CODE,'S','S','N'),ICTPROD1.PROD_CATGY) ITEM_SNG" & vbCrLf _
            & " , SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
            & " , SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) ORDR_AMT_SHIP" & vbCrLf _
            & " , SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * DECODE(SOTINVH2.WHSE_CODE,NULL,0,NVL(SOTINVH2.ITEM_UNIT_COST,0))) ORDR_CGS_SHIP" & vbCrLf _
            & " , 0 ORDR_SCP_SHIP" & vbCrLf _
            & " , 0 ORDR_VCC_SHIP" & vbCrLf _
            & " from SOTINVH2, ICTITEM1, ICTCOSTA, ICTPROD1" & vbCrLf _
            & " Where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP (+) = SOTINVH2.ORDR_YYYYPP_UPDATED" & vbCrLf _
            & "   and ICTCOSTA.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTINVH2.ORDR_YYYYPP_UPDATED = '" & yp & "'" & vbCrLf _
            & "   and ICTPROD1.PROD_CODE (+) = ICTITEM1.PROD_CODE" & vbCrLf _
            & " group by" & vbCrLf _
            & " SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, " & vbCrLf _
            & " ICTITEM1.COLLECTION_CODE, SOTINVH2.ORDR_YYYYPP_UPDATED," & vbCrLf _
            & " SOTINVH2.INV_TYPE, ICTITEM1.ITEM_BASIC_PROMO," & vbCrLf _
            & " DECODE(NVL(ICTPROD1.PROD_CATGY,'0'),'0',DECODE(ICTITEM1.ITEM_SNU_CODE,'S','S','N'),ICTPROD1.PROD_CATGY)" & vbCrLf
        ASCDATA1.ExecuteSQL()

        'sql = ""
        'sql = sql & "Begin"
        'sql = sql & " Declare Cursor C1 is "
        'sql = sql & "  Select Distinct CUST_CODE, CUST_STORE_NO"
        'sql = sql & "   from GLTPROFS"
        'sql = sql & "   where OPS_YYYYPP = '" & yp & "';"
        'sql = sql & " SREP_CODE_X VARCHAR2(6);"
        'sql = sql & " SELL_CODE_X VARCHAR2(6);"
        'sql = sql & " Begin"
        'sql = sql & "  For R1 in C1 Loop"
        'sql = sql & "   Select SREP_CODE, SELL_CODE into SREP_CODE_X, SELL_CODE_X from ARTCUST2"
        'sql = sql & "    where CUST_CODE = R1.CUST_CODE"
        'sql = sql & "      and CUST_STORE_NO = R1.CUST_STORE_NO;"
        'sql = sql & "   Update GLTPROFS set SREP_CODE = SREP_CODE_X, SELL_CODE = SELL_CODE_X"
        'sql = sql & "    where CUST_CODE = R1.CUST_CODE"
        'sql = sql & "      and CUST_STORE_NO = R1.CUST_STORE_NO"
        'sql = sql & "      and OPS_YYYYPP = '" & yp & "';"
        'sql = sql & "  End Loop;"
        'sql = sql & " End;"
        'sql = sql & "End;"
        'OraD.ExecuteSQL sql

        '            & ", NULL SREP_CODE" & vbCrLf _

        ASCMAIN1.sql = "Insert into GLTPROF0" & vbCrLf _
            & " SELECT GLTPROFS.CUST_CODE, " & vbCrLf _
            & " GLTPROFS.COLLECTION_CODE, GLTPROFS.OPS_YYYYPP," & vbCrLf _
            & " GLTPROFS.INV_TYPE, GLTPROFS.ITEM_BASIC_PROMO," & vbCrLf _
            & " GLTPROFS.ITEM_SNG" & vbCrLf _
            & " , SUM (NVL(GLTPROFS.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
            & " , SUM (NVL(GLTPROFS.ORDR_AMT_SHIP,0)) ORDR_AMT_SHIP" & vbCrLf _
            & " , SUM (NVL(GLTPROFS.ORDR_CGS_SHIP,0)) ORDR_CGS_SHIP" & vbCrLf _
            & " , SUM (NVL(GLTPROFS.ORDR_SCP_SHIP,0)) ORDR_SCP_SHIP" & vbCrLf _
            & " , SUM (NVL(GLTPROFS.ORDR_VCC_SHIP,0)) ORDR_VCC_SHIP" & vbCrLf _
            & " from GLTPROFS" & vbCrLf _
            & " where GLTPROFS.OPS_YYYYPP = '" & yp & "'" & vbCrLf _
            & " group by" & vbCrLf _
            & " GLTPROFS.CUST_CODE, " & vbCrLf _
            & " GLTPROFS.COLLECTION_CODE, GLTPROFS.OPS_YYYYPP," & vbCrLf _
            & " GLTPROFS.INV_TYPE, GLTPROFS.ITEM_BASIC_PROMO," & vbCrLf _
            & " GLTPROFS.ITEM_SNG" & vbCrLf
        ASCDATA1.ExecuteSQL()

        'sql = ""
        'sql = sql & "Begin"
        'sql = sql & " Declare Cursor C1 is "
        'sql = sql & "  Select Distinct CUST_CODE"
        'sql = sql & "   from GLTPROF0"
        'sql = sql & "   where OPS_YYYYPP = '" & yp & "';"
        'sql = sql & " SREP_CODE_X VARCHAR2(6);"
        'sql = sql & " Begin"
        'sql = sql & "  For R1 in C1 Loop"
        'sql = sql & "   Select SREP_CODE into SREP_CODE_X from ARTCUST1"
        'sql = sql & "    where CUST_CODE = R1.CUST_CODE;"
        'sql = sql & "   Update GLTPROF0 set SREP_CODE = SREP_CODE_X"
        'sql = sql & "    where CUST_CODE = R1.CUST_CODE"
        'sql = sql & "      and OPS_YYYYPP = '" & yp & "';"
        'sql = sql & "  End Loop;"
        'sql = sql & " End;"
        'sql = sql & "End;"
        'OraD.ExecuteSQL sql

    End Sub

    Sub Get_Stats()
        ASCMAIN1.sql = "Select * from GLTPROFW where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -24) & "'"
        grdGLTPROFW.DataSource = ASCDATA1.GetDataTable
        Sort_grdColumns(grdGLTPROFW, "OPS_YYYYPP".ToLower)

        If ASCMAIN1.CLIENT = "AHA" Then
        Else
            With grdGLTPROFW.DisplayLayout.Bands(0)
                .Columns("BI_PREP_DATE").Hidden = True
                .Columns("BI_PREP_OPER").Hidden = True
                .Columns("BI_PREP_TOTAL").Hidden = True
                .Columns("BI_SENT_DATE").Hidden = True
                .Columns("BI_SENT_OPER").Hidden = True
                .Columns("SEND_BI").Hidden = True
            End With
        End If

    End Sub

    Private Sub grdGLTPROFW_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdGLTPROFW.ClickCellButton
        If e.Cell.Column.Key = "SEND_BI" Then
            Dim YP As String = e.Cell.Row.Cells("OPS_YYYYPP").Value
            Dim YP_Legend As String = ASCMAIN1.Get_Legend(YP)


            If e.Cell.Row.Cells("BI_PREP_OPER").Value & "" = "" Or e.Cell.Row.Cells("BI_PREP_DATE").Value & "" = "" Then
                MsgBox("BI Data has not been Prepared for " & YP_Legend, MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            ElseIf e.Cell.Row.Cells("BI_SENT_DATE").Value & "" <> "" AndAlso Format(e.Cell.Row.Cells("BI_SENT_DATE").Value, "yyyyMMddHHmm") > Format(e.Cell.Row.Cells("BI_PREP_DATE").Value, "yyyyMMddHHmm") Then
                MsgBox("BI Data has not been Re-Prepared for " & YP_Legend & " since it was sent", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Else

                Dim msg As String = ""
                If e.Cell.Row.Cells("BI_SENT_DATE").Value & "" <> "" Then
                    msg = vbCrLf & vbCrLf & "This will overwrite BI data sent previously!"
                End If
                If MsgBox("Are you sure you want to Send BI Data for " & YP_Legend & "?" & msg, vbYesNo, "Verification") = MsgBoxResult.Yes Then

                    BeginTrans()

                    'ASCMAIN1.sql = "Delete from TATQVXQ1 where YYYY = '2017' and PP = '09' and RECORD_STATUS = '1'"
                    ASCMAIN1.sql = "Update TATQVXQ1 set RECORD_STATUS = 'D' where YYYY = '2017' and PP = '09' and RECORD_STATUS = '1'"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "Update TATQVXQ1 set RECORD_STATUS = '1', TIMESTAMP_1 = SYSDATE, VERIFICATION_NO = '" & ASCMAIN1.SESSION_NO & "' where RECORD_STATUS = '0' AND YYYY = '" & Mid(YP, 1, 4) & "' AND PP = '" & Mid(YP, 5, 2) & "'"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "Update GLTPROFW Set BI_SENT_OPER = '" & ASCMAIN1.USER_ID & "', BI_SENT_DATE = SYSDATE WHERE OPS_YYYYPP = '" & YP & "'"
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.Record_Event("TATQVXQ1", YP, "", Now, ASCMAIN1.USER_ID, "SENT", "BI Data Sent for " & YP_Legend, ASCMAIN1.SESSION_NO)

                    CommitTrans()

                    Get_Stats()

                    MsgBox("BI Data has been Sent for " & YP_Legend, MsgBoxStyle.OkOnly, "Confirmation")
                End If
            End If

        End If
    End Sub

    Private Sub grdGLTPROFW_InitializeLayout(sender As Object, e As UltraWinGrid.InitializeLayoutEventArgs) Handles grdGLTPROFW.InitializeLayout

    End Sub
End Class