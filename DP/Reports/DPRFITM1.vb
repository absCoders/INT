Imports Infragistics.Win.UltraWinGrid

Public Class DPRFITM1

    Private md() As String
    Private Lmd() As String

    Private ASWSRPT0 As String = String.Empty
    Private ICTITEM1 As String = String.Empty

    Private LYP As String = String.Empty
    Private LYP2 As String = String.Empty
    Private XYP As String = String.Empty
    Private typ As String = String.Empty
    Private periodDifference As Integer = 0
    Private ZYP As String = String.Empty
    Private UYP As String = String.Empty
    Dim recCount As Int64 = 0
    Dim SUPPRESS_UNIT_PRICE = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("GLTPARM1")
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)

    End Sub

    Protected Overrides Sub Build_Workfile()
        Prepare_dst(True)
    End Sub

    Public Overrides Sub Print_Report()

        Dim Subt As String = String.Empty
        Select Case optSHOW.Value
            Case "F"
                Subt = "Showing Monthly Forecasts"
            Case "C"
                Subt = "Showing Changes in Monthly Forecasts"
        End Select

        CR_params.Add("SUBT", Subt)

        Dim displayUnitsSales As String = String.Empty
        If chkUnits.Checked Then
            displayUnitsSales &= "U"
        End If

        If chkSales.Checked Then
            displayUnitsSales &= "S"
        End If

        If chkCosts.Checked Then
            displayUnitsSales &= "C"
        End If

        CR_params.Add("CHKUSC", displayUnitsSales)

        CR_params.Add("CHKLY", IIf(chkLY.Checked, "1", "0"))
        CR_params.Add("CHKSUP", IIf(chkSUP.Checked, "1", "0"))
        CR_params.Add("CHKIGNOREPDF", IIf(chkPDF.Checked, "1", "0"))
        CR_params.Add("CHKSUPF", IIf(chkSUPF.Checked, "1", "0"))
        CR_params.Add("SUPPRESS_UNIT_PRICE", SUPPRESS_UNIT_PRICE)

        For i = 1 To 12
            CR_params.Add("MD" & Format$(i, "00"), md(i))
        Next i

        For i = 1 To 12
            CR_params.Add("LMD" & Format$(i, "00"), Lmd(i))
        Next i

        Dim showPastDue As String = String.Empty
        If ASCMAIN1.CYP = RYP Then
            showPastDue = "Y"
        Else
            showPastDue = "N"
        End If
        CR_params.Add("SHOWPD", showPastDue)

        Generate_Report("DPRFITM1", String.Empty, String.Empty)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If cmbYP0.Value Is Nothing Then
                EMsg &= vbCr & "You Must Select a Forecast Start Period"
            End If
            'Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("MARKET_CODE")

            ''If chkLY.Checked Then
            ''    If rowASTDSQLA.Item("SEQUENCE") & "" = "" Then
            ''        EMsg &= vbCr & "You Must Select Market Code in the Sort when using the LY option"
            ''    End If
            ''End If
        End If
    End Sub

    Overrides Function Prepare_dst(
          ByVal perform_fill As Boolean,
          ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then
            Clear_dst()
        End If

        Dim sql As String = String.Empty

        With dst
            sql = "Select ITEM_CODE, ITEM_DESC, NVL(ITEM_RETAIL_PRICE,0) ITEM_RETAIL_PRICE, NVL(ITEM_COST_STD,0) ITEM_COST_STD "
            sql &= ", ITEM_CATGY_CODE, ITEM_SNU_CODE"
            sql &= ", COLLECTION_CODE, NVL(ITEM_RETAIL_PRICE,0) ITEM_UNIT_PRICE"
            sql &= " from ICTITEM1"
            ICTITEM1 = ASCMAIN1.Temp_Table(sql)

            sql = "Select * from " & ICTITEM1
            Create_TDA(.Tables.Add, "ICTITEM1", sql, 0, False, "", 1)

            sql = "Select ITEM_CODE,MARKET_CODE,0 ITEM_UNIT_PRICE"
            sql &= " from ICTITEM1,SOTMKTC1 WHERE ROWNUM < 1 "
            Create_TDA(.Tables.Add, "ICTITEMP", sql, 0, False, "", 0)

            Create_TDA(.Tables.Add, "ICTCOLL1", "*")
            Create_TDA(.Tables.Add, "ICTBRAN1", "*")

            Create_TDA(.Tables.Add, "SOTPCLS1", "*", 0)
            Create_TDA(.Tables.Add, "SOTMKTC1", "*", 0)
            Create_TDA(.Tables.Add, "SOTPRIC1", "*", 0)
            Create_TDA(.Tables.Add, "SOTPRIC2", "*", 0)
        End With

        Fill_Records("SOTPCLS1")
        Fill_Records("SOTMKTC1")
        Fill_Records("SOTPRIC1")
        Fill_Records("SOTPRIC2")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sql As String = String.Empty
        recCount = 0
        SUPPRESS_UNIT_PRICE = String.Empty

        EnforceConstraints(False)

        Dim i As Integer
        '     Dim j As Integer
        Dim z As String
        Dim sql_Past As String

        Dim Future_Forecast As String = String.Empty
        Dim Previous_Forecast = String.Empty

        ' Get Codes selected from tabs
        ASCMAIN1.Progress("Run-Time Options", "")

        RYPLEGEND = cmbYP0.Value
        RYP = Mid$(RYPLEGEND, 1, 4) & Mid$(RYPLEGEND, 6, 2)

        LYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        LYP2 = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -2)
        XYP = ASCMAIN1.Period_Calc(RYP, 24 - 1)

        Call Period_Range(RYP, XYP, 0, ZYP, "")
        Call Period_Range(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12), ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1), 0, UYP, "")

        periodDifference = ASCMAIN1.Period_Diff(ASCMAIN1.CYP, RYP)
        Do While periodDifference < 0
            periodDifference = periodDifference + 12
        Loop

        ReDim md(24)
        ReDim Lmd(12)
        For i = 1 To 24
            z = ASCMAIN1.Get_YYYYMM(Mid$(ZYP, (i - 1) * 6 + 1, 6), 0)
            z = Mid$(z, 5, 2) & "/01/" & Mid$(z, 3, 2)
            md(i) = CDate(z).ToString("MMMyy") 'Format$(z, "MMMyy")
            If i <= 12 Then Lmd(i) = CDate(z).AddYears(-1).ToString("MMMyy")
        Next i


        'For i = 1 To 12
        '    j = i + periodDifference
        '    If j > 12 Then
        '        j = j - 12
        '    End If
        '    z = ASCMAIN1.Get_YYYYMM(Mid$(UYP, (j - 1) * 6 + 1, 6), 0)
        '    z = Mid$(z, 5, 2) & "/01/" & Mid$(z, 3, 2)
        '    Lmd(i) = CDate(z).ToString("MMMyy")  'Format$(z, "mmm'YY")
        'Next i

        ' Set up Work File Definition using X's and 0's and 0.01's as required
        Get_SQL("*")

        ASCMAIN1.Progress("Initialize Work Tables", "")
        sql = "Select 0 RECORD_NO, LPAD ('X',22) ITEM_CODE "
        For i = 0 To 24
            sql &= ", 0 FU_" & Format$(i, "00")
            sql &= ", 0 FS_" & Format$(i, "00")
            sql &= ", 0 FC_" & Format$(i, "00")
        Next i
        For i = 0 To 24
            sql &= ", 0 LU_" & Format$(i, "00")
            sql &= ", 0 LS_" & Format$(i, "00")
            sql &= ", 0 LC_" & Format$(i, "00")
        Next i
        ' G1thru9
        For i = 1 To 9
            sql &= ", LPAD (' ', 40) G" & i
        Next
        sql &= ", LPAD ('X',6) MARKET_CODE"
        sql &= " from DUAL where ROWNUM < 1"
        ASWSRPT0 = ASCMAIN1.Temp_Table(sql)

        sql = "SELECT * FROM " & ASWSRPT0
        Call Create_TDA(dst.Tables.Add, ASWSRPT0, sql, 0, True, String.Empty, 2)

        ' Forecasts
        typ = "F"
        Get_SQL(typ)
        If InStr(1, sql_GROUP_BY_cols, "MARKET_CODE") = 0 Then
            sql_GROUP_BY_cols &= ", DPTITMF1.MARKET_CODE"
        End If

        ASCMAIN1.Progress("Forecast Data", "")
        sql = "Select " & sql_GROUP_BY_cols & ", OPS_YYYYPP"

        Dim pastFields As String = String.Empty
        For Each field As String In sql_GROUP_BY_cols.Split(",")
            field = field.Trim
            If field.Contains(".") Then
                pastFields &= ", " & field.Split(".")(1)
            Else
                pastFields &= ", " & field
            End If
        Next
        If pastFields.Length > 2 Then pastFields = pastFields.Substring(1).Trim
        sql_Past = "Select " & pastFields & ", OPS_YYYYPP"

        For i = 0 To 24
            If i = 0 Then
                z = "000000"
            Else
                z = Mid$(ZYP, (i - 1) * 6 + 1, 6)
            End If
            sql &= ", Sum (Decode (OPS_YYYYPP_FC,'" & z & "', DPTITMF1.FORECAST,0)) X" & Format$(i, "00")
            sql_Past &= ", Sum (X" & Format$(i, "00") & ") X" & Format$(i, "00")
        Next i

        sql &= " from DPTITMF1" & sql_TABLE_NAMEs
        If optSHOW.Value = "C" Then
            sql &= "   where DPTITMF1.OPS_YYYYPP >= '" & LYP & "'"
            sql &= "   and DPTITMF1.OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'"
        Else
            sql &= "   where DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
        End If

        If sql_TABLE_NAMEs.Contains("ICTITEM1") Then
            sql &= " and ICTITEM1.ITEM_CODE = DPTITMF1.ITEM_CODE"
        End If

        If chkPDF.Checked Then
            sql &= "   and DPTITMF1.OPS_YYYYPP_FC <> '000000'"
        End If

        sql &= sql_JOIN
        sql &= sql_WHERE
        sql &= " group by " & sql_GROUP_BY_cols & ", OPS_YYYYPP"
        Future_Forecast = sql

        Previous_Forecast = ""
        If RYP < ASCMAIN1.CYP Then
            sql = "Select " & sql_GROUP_BY_cols & ", OPS_YYYYPP"
            For i = 0 To 24
                If i = 0 Then
                    z = "000000"
                Else
                    z = Mid$(ZYP, (i - 1) * 6 + 1, 6)
                End If
                sql &= ", Sum (Decode (OPS_YYYYPP_FC,'" & z & "', DPTITMF1.FORECAST,0)) X" & Format$(i, "00")
            Next i

            sql &= " from DPTITMF1" & sql_TABLE_NAMEs
            sql &= " where DPTITMF1.OPS_YYYYPP = DPTITMF1.OPS_YYYYPP_FC"
            If sql_TABLE_NAMEs.Contains("ICTITEM1") Then
                sql &= " and ICTITEM1.ITEM_CODE = DPTITMF1.ITEM_CODE"
            End If

            If optSHOW.Value = "C" Then
                sql &= " And DPTITMF1.OPS_YYYYPP BETWEEN '" & RYP & "' AND  '" & LYP2 & "'"
            Else
                sql &= " And DPTITMF1.OPS_YYYYPP BETWEEN '" & RYP & "' AND  '" & LYP & "'"
            End If

            If chkPDF.Checked Then
                sql &= "   and DPTITMF1.OPS_YYYYPP_FC <> '000000'"
            End If

            sql &= sql_JOIN
            sql &= sql_WHERE
            sql &= " group by " & sql_GROUP_BY_cols & ", OPS_YYYYPP"

            Previous_Forecast = sql

        End If

        If Previous_Forecast = "" Then
            sql = Future_Forecast
        Else
            sql = sql_Past
            sql &= vbCr & " From "
            sql &= vbCr & " ( "
            sql &= vbCr & " (" & Future_Forecast & ")"
            sql &= vbCr & " UNION"
            sql &= vbCr & " (" & Previous_Forecast & ")"
            sql &= vbCr & " )"
            sql &= " group by " & pastFields & ", OPS_YYYYPP"
        End If

        ' Preserve the items that fall into the report from when LY is selected
        Dim itemTable As String = ASCMAIN1.Temp_Table("select distinct item_code from (" & sql & ")")

        ProcessRecords(sql)

        ' Shipments
        If chkLY.Checked Then
            ASCMAIN1.Progress("Shipments Data", "")

            typ = "*"
            Call Get_SQL(typ)

            sql = "Select " & sql_SELECT_cols & vbCrLf

            If InStr(1, sql_GROUP_BY_cols, "MARKET_CODE") = 0 Then
                sql &= ", MAX(NVL(SOTMKTCC.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,SOTTCLS1.MARKET_CODE))) MARKET_CODE"
                sql_GROUP_BY_cols &= ", NVL(SOTMKTCC.MARKET_CODE,NVL(SOTMKTC1.MARKET_CODE_FC,SOTTCLS1.MARKET_CODE))"
            End If

            For i = 1 To 12
                z = Mid$(UYP, (i - 1) * 6 + 1, 6)
                sql &= ", Sum (Decode (ORDR_YYYYPP_UPDATED,'" & z & "', SOTINVH2.ORDR_QTY_SHIP,0)) X" & Format$(i, "00") & vbCrLf
            Next i

            If Not sql_TABLE_NAMEs.Contains("ICTITEM1") Then
                    sql_TABLE_NAMEs &= ", ICTITEM1"
                End If
                If Not sql_TABLE_NAMEs.Contains("SOTTCLS1") Then
                    sql_TABLE_NAMEs &= ", SOTTCLS1"
                End If
            'If Not sql_TABLE_NAMEs.Contains("SOTMKTC1") Then
            '    sql_TABLE_NAMEs &= ", SOTMKTC1"
            'End If
            If Not sql_TABLE_NAMEs.Contains("ARTCUST1") Then
                sql_TABLE_NAMEs &= ", ARTCUST1"
            End If

            sql &= " from SOTINVH2 " & sql_TABLE_NAMEs & ", " & itemTable & vbCrLf

                sql &= " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf
                sql &= "   and ICTITEM1.ITEM_CODE = " & itemTable & ".ITEM_CODE" & vbCrLf
                sql &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & Microsoft.VisualBasic.Left(UYP, 6) & "'" & vbCrLf
                sql &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & Microsoft.VisualBasic.Right(UYP, 6) & "'" & vbCrLf
                sql &= "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf
                sql &= "   and ARTCUST1.CUST_CODE = SOTINVH2.CUST_CODE" & vbCrLf
                sql &= "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf
                'sql &= "   and SOTMKTC1.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf
                sql &= sql_JOIN & vbCrLf
                sql &= sql_WHERE & vbCrLf
                sql &= " group by " & sql_GROUP_BY_cols & vbCrLf

                ProcessRecords(sql)
            End If

            ' ASWSRPT0
            sql = "Select " & G1thru9 & ", ITEM_CODE " & vbCrLf
        sql &= ", MAX(MARKET_CODE) MARKET_CODE"
        For i = 0 To 24
            sql &= ", SUM(FU_" & Format$(i, "00") & ") SUM_FU_" & Format$(i, "00")
            sql &= ", SUM(FS_" & Format$(i, "00") & ") SUM_FS_" & Format$(i, "00")
            sql &= ", SUM(FC_" & Format$(i, "00") & ") SUM_FC_" & Format$(i, "00") & vbCrLf
        Next i
        For i = 0 To 24
            sql &= ", SUM(LU_" & Format$(i, "00") & ") SUM_LU_" & Format$(i, "00")
            sql &= ", SUM(LS_" & Format$(i, "00") & ") SUM_LS_" & Format$(i, "00")
            sql &= ", SUM(LC_" & Format$(i, "00") & ") SUM_LC_" & Format$(i, "00") & vbCrLf
        Next i
        sql &= " from " & ASWSRPT0
        sql &= " GROUP BY " & G1thru9 & ", ITEM_CODE "

        Update_Record_TDA(ASWSRPT0)
        ASCDATA1.ExecuteSQL("Insert Into " & ASTSRPT1 & " " & sql)

        EnforceConstraints(True)

    End Sub

    Sub Period_Range(ByVal zz1 As String, ByVal zz2 As String, ByRef zzn As Integer, ByRef zzp As String, ByVal zzs As String)
        ' Pass in 1st Period zz1 and 2nd Period zz2
        ' Returns zzn (number of periods) and zzp (table of periods, yyyymm, separated by zzs)
        Dim i As Integer
        Dim zzz As String

        If zz1 > zz2 Then
            Call Period_Range(zz2, zz1, zzn, zzp, zzs)
            zzz = ""
            For i = 1 To zzn
                zzz = Mid$(zzp, (i - 1) * (6 + Len(zzs)) + 1, 6 + Len(zzs)) & zzz
            Next i
            zzn = -1 * zzn
            zzp = zzz
            Exit Sub
        End If

        Dim zzy As Integer
        Dim zzm As Integer

        zzy = Val(Mid$(zz2, 1, 4)) - Val(Mid$(zz1, 1, 4))
        zzm = Val(Mid$(zz2, 5, 2)) - Val(Mid$(zz1, 5, 2))
        zzn = zzy * 12 + zzm + 1

        zzp = ""
        zzz = zz1
        For i = 1 To zzn
            zzp = zzp & zzz & zzs
            If Mid$(zzz, 5, 2) = "12" Then
                zzz = Format$(Val(Mid$(zzz, 1, 4)) + 1, "0000") & "01"
            Else
                zzz = Mid$(zzz, 1, 4) & Format$(Val(Mid$(zzz, 5, 2)) + 1, "00")
            End If
        Next i

    End Sub

    Private Sub ProcessRecords(ByVal query As String)

        ' Dim recCount As Int64 = 0

        If typ = "F" Then
            If InStr(1, sql_SELECT_cols, "MARKET_CODE") > 1 Then
                SUPPRESS_UNIT_PRICE = String.Empty
            Else
                SUPPRESS_UNIT_PRICE = "1"
            End If
        End If


        For Each rowData As DataRow In ASCDATA1.GetDataTable(query).Rows


            recCount += 1
            If recCount Mod 100 = 0 Then
                ASCMAIN1.Progress("-", recCount)
            End If
            WriteRecords(rowData) ' , recCount)
        Next
    End Sub

    Private Sub WriteRecords(ByRef rowData As DataRow) ' , ByRef recCount As Int64)

        Dim ITEM_CODE As String = rowData.Item("ITEM_CODE") & String.Empty

        Dim rowICTITEMX As DataRow = Nothing
        Dim rowICTITEM1 As DataRow = Nothing
        Dim rowICTITEMP As DataRow = Nothing
        Dim rowICTCOSTC As DataRow = Nothing

        Dim ITEM_RETAIL_PRICE As Decimal = 0
        Dim price As Decimal = 0
        Dim cost As Decimal = 0
        Dim units As Int32 = 0
        Dim jCount As Integer = 0

        If dst.Tables("ICTITEM1").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length = 0 Then

            rowICTITEMX = MyBase.LookUp("ICTITEM1", ITEM_CODE)
            rowICTITEM1 = dst.Tables("ICTITEM1").NewRow

            rowICTITEM1.Item("ITEM_CODE") = ITEM_CODE
            rowICTITEM1.Item("ITEM_DESC") = rowICTITEMX.Item("ITEM_DESC") & ""
            rowICTITEM1.Item("ITEM_RETAIL_PRICE") = rowICTITEMX.Item("ITEM_RETAIL_PRICE")
            rowICTITEM1.Item("ITEM_COST_STD") = rowICTITEMX.Item("ITEM_COST_STD")
            rowICTITEM1.Item("ITEM_CATGY_CODE") = rowICTITEMX.Item("ITEM_CATGY_CODE")
            rowICTITEM1.Item("ITEM_SNU_CODE") = rowICTITEMX.Item("ITEM_SNU_CODE")
            rowICTITEM1.Item("COLLECTION_CODE") = rowICTITEMX.Item("COLLECTION_CODE")

            ITEM_RETAIL_PRICE = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
            If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
                price = ITEM_RETAIL_PRICE * 0.5
            Else
                price = ITEM_RETAIL_PRICE * 0.6
            End If
            rowICTITEM1.Item("ITEM_UNIT_PRICE") = price
            cost = Val(rowICTITEM1.Item("ITEM_COST_STD") & String.Empty)

            dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)
        Else
            rowICTITEM1 = dst.Tables("ICTITEM1").Select("ITEM_CODE = '" & ITEM_CODE & "'")(0)
            price = Val(rowICTITEM1.Item("ITEM_UNIT_PRICE") & String.Empty)
            ITEM_RETAIL_PRICE = Val(rowICTITEM1.Item("ITEM_RETAIL_PRICE") & "")
            cost = Val(rowICTITEM1.Item("ITEM_COST_STD") & String.Empty)
        End If

        Dim rowASWSRPT0 As DataRow = dst.Tables(ASWSRPT0).NewRow
        rowASWSRPT0.Item("RECORD_NO") = recCount
        rowASWSRPT0.Item("ITEM_CODE") = ITEM_CODE
        rowASWSRPT0.Item("MARKET_CODE") = (rowData.Item("MARKET_CODE") & String.Empty).ToString.Trim
        dst.Tables(ASWSRPT0).Rows.Add(rowASWSRPT0)

        SetKey(rowASWSRPT0, rowData)

        'Can only get market price when Market is in query
        '  If typ = "F" AndAlso InStr(1, sql_SELECT_cols, "MARKET_CODE") > 1 Then
        If typ = "F" Then

            Dim MARKET_CODE As String = (rowData.Item("MARKET_CODE") & String.Empty).ToString.Trim

            If MARKET_CODE.Length > 0 Then

                Dim rowSOTMKTC1 As DataRow = dst.Tables("SOTMKTC1").Rows.Find(MARKET_CODE)
                Dim CUST_CODE As String = rowSOTMKTC1.Item("CUST_CODE") & ""
                If CUST_CODE <> "" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    Dim PRICE_LIST_CODE As String = rowARTCUST1.Item("PRICE_LIST_CODE") & ""
                    Dim PRICE_CLASS_CODE As String = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
                    If PRICE_LIST_CODE <> "" Then
                        Dim rowSOTPRIC2 As DataRow = dst.Tables("SOTPRIC2").Rows.Find(New String() {PRICE_LIST_CODE, ITEM_CODE})
                        If rowSOTPRIC2 IsNot Nothing Then
                            price = Val(rowSOTPRIC2.Item("ITEM_PRICE") & "")
                        Else
                            If PRICE_CLASS_CODE <> "" Then
                                Get_Price_from_PRICE_CLASS_CODE(PRICE_CLASS_CODE, ITEM_RETAIL_PRICE, price)
                            End If
                        End If
                    Else
                        If PRICE_CLASS_CODE <> "" Then
                            Get_Price_from_PRICE_CLASS_CODE(PRICE_CLASS_CODE, ITEM_RETAIL_PRICE, price)
                        End If
                    End If
                Else
                    Dim PRICE_CLASS_CODE As String = rowSOTMKTC1.Item("PRICE_CLASS_CODE") & ""
                    If PRICE_CLASS_CODE <> "" Then
                        Get_Price_from_PRICE_CLASS_CODE(PRICE_CLASS_CODE, ITEM_RETAIL_PRICE, price)
                    End If
                End If

                '    rowICTITEM1.Item("ITEM_UNIT_PRICE") = price
            Else
                ' Stop
            End If


            If dst.Tables("ICTITEMP").Select("((ITEM_CODE = '" & ITEM_CODE & "' and MARKET_CODE = '" & MARKET_CODE & "'))").Length = 0 Then

                '        rowICTITEMP = MyBase.LookUp("ICTITEM1", ITEM_CODE, MARKET_CODE)
                rowICTITEMP = dst.Tables("ICTITEMP").NewRow
                rowICTITEMP.Item("ITEM_CODE") = ITEM_CODE
                rowICTITEMP.Item("MARKET_CODE") = MARKET_CODE
                rowICTITEMP.Item("ITEM_UNIT_PRICE") = price
                dst.Tables("ICTITEMP").Rows.Add(rowICTITEMP)
            Else
                ' rowICTITEMP.Item("ITEM_UNIT_PRICE") = price
            End If
        End If

        If typ = "F" Then
            For i = 0 To 24
                units = Val(rowData.Item("X" & Format$(i, "00")) & "")
                If (rowData.Item("OPS_YYYYPP") = LYP) And (optSHOW.Value = "C") Then
                    units = -1 * units
                End If
                rowASWSRPT0.Item("FU_" & Format$(i, "00")) = units
                rowASWSRPT0.Item("FS_" & Format$(i, "00")) = units * price
                rowASWSRPT0.Item("FC_" & Format$(i, "00")) = units * cost
            Next i
        Else
            For iCount As Integer = 1 To 12
                jCount = iCount + periodDifference
                If jCount > 12 Then
                    jCount = jCount - 12
                End If
                units = Val(rowData.Item("X" & Format$(jCount, "00")) & "")
                rowASWSRPT0.Item("LU_" & Format$(iCount, "00")) = units
                rowASWSRPT0.Item("LS_" & Format$(iCount, "00")) = units * price
                rowASWSRPT0.Item("LC_" & Format$(iCount, "00")) = units * cost
            Next iCount
        End If

    End Sub

    Sub Get_Price_from_PRICE_CLASS_CODE(PRICE_CLASS_CODE As String, ITEM_RETAIL_PRICE As Decimal, ByRef PRICE As Decimal)
        Dim rowSOTPCLS1 As DataRow = dst.Tables("SOTPCLS1").Rows.Find(PRICE_CLASS_CODE)
        If rowSOTPCLS1 Is Nothing Then Stop
        If rowSOTPCLS1 IsNot Nothing Then
            Dim PRICE_BASE_DPCT As Decimal = Val(rowSOTPCLS1.Item("PRICE_BASE_DPCT") & "")
            Dim PRICE_BASIS As String = rowSOTPCLS1.Item("PRICE_BASIS") & ""
            If PRICE_BASIS = "R" Then
                PRICE = ITEM_RETAIL_PRICE * (100 - PRICE_BASE_DPCT) / 100
            End If
        End If
    End Sub
    Private Sub SetKey(ByRef rowASWSRPT0 As DataRow, ByRef rowData As DataRow)

        Dim fieldName As String = String.Empty
        For fieldNum As Integer = 0 To COLUMN_NAMEs.Count - 1
            rowASWSRPT0.Item("G" & fieldNum + 1) = rowData(COLUMN_NAMEs(fieldNum)) & String.Empty
        Next
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Public Overrides Sub Post_Process_Special()
        MyBase.Post_Process_Special()
        Dim LEGEND As String = ""
        Dim MMM_CUR As String = ""
        Dim MMM_WORK As String = Mid(RYPLEGEND, 1, 4) & Mid(RYPLEGEND, 6, 2)

        With grdASTDSQL1.DisplayLayout
            ' Iterate through each group and set format for each column in the group
            For Each group As UltraGridGroup In .Bands(0).Groups
                For Each column As UltraGridColumn In group.Columns
                    ' Set the format for the column
                    column.Format = "#,##0"
                Next
            Next
        End With

        ' datagrid
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(11).Header.Caption = "FU Past Due"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(12).Header.Caption = "FS Past Due"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(13).Header.Caption = "FC Past Due"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(86).Header.Caption = "LU Past Due"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(87).Header.Caption = "LS Past Due"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(88).Header.Caption = "LC Past Due"

        grdASTSRPT1.DisplayLayout.Bands(0).Columns(11).Format = "#,##0"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(12).Format = "#,##0"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(13).Format = "#,##0"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(86).Format = "#,##0"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(87).Format = "#,##0"
        grdASTSRPT1.DisplayLayout.Bands(0).Columns(88).Format = "#,##0"

        ' query 
        grdASTDSQL1.DisplayLayout.Bands(0).Groups(3).Header.Caption = "FU Past Due"
        grdASTDSQL1.DisplayLayout.Bands(0).Groups(4).Header.Caption = "FS Past Due"
        grdASTDSQL1.DisplayLayout.Bands(0).Groups(5).Header.Caption = "FC Past Due"
        grdASTDSQL1.DisplayLayout.Bands(0).Groups(78).Header.Caption = "LU Past Due"
        grdASTDSQL1.DisplayLayout.Bands(0).Groups(79).Header.Caption = "LS Past Due"
        grdASTDSQL1.DisplayLayout.Bands(0).Groups(80).Header.Caption = "LC Past Due"

        grdASTSRPT1.DisplayLayout.Bands(0).Columns(10).Hidden = True


        If chkUnits.Checked = True Then
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(11).Hidden = False
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(3).Hidden = False
            If chkLY.Checked Then
                grdASTSRPT1.DisplayLayout.Bands(0).Columns(86).Hidden = False
                grdASTDSQL1.DisplayLayout.Bands(0).Groups(78).Hidden = False
            Else
                grdASTSRPT1.DisplayLayout.Bands(0).Columns(86).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups(78).Hidden = True
            End If
        Else
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(11).Hidden = True
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(86).Hidden = True
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(3).Hidden = True
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(78).Hidden = True
        End If

        If chkSales.Checked = True Then
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(12).Hidden = False
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(4).Hidden = False
            If chkLY.Checked Then
                grdASTSRPT1.DisplayLayout.Bands(0).Columns(87).Hidden = False
                grdASTDSQL1.DisplayLayout.Bands(0).Groups(79).Hidden = False
            Else
                grdASTSRPT1.DisplayLayout.Bands(0).Columns(87).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups(79).Hidden = True
            End If
        Else
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(12).Hidden = True
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(87).Hidden = True
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(4).Hidden = True
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(79).Hidden = True
        End If

        If chkCosts.Checked = True Then
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(13).Hidden = False
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(5).Hidden = False
            If chkLY.Checked Then
                grdASTSRPT1.DisplayLayout.Bands(0).Columns(88).Hidden = False
                grdASTDSQL1.DisplayLayout.Bands(0).Groups(80).Hidden = False
            Else
                grdASTSRPT1.DisplayLayout.Bands(0).Columns(88).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups(80).Hidden = True
            End If
        Else
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(13).Hidden = True
            grdASTSRPT1.DisplayLayout.Bands(0).Columns(88).Hidden = True
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(5).Hidden = True
            grdASTDSQL1.DisplayLayout.Bands(0).Groups(80).Hidden = True
        End If

        ' Periods

        For i = 0 To 23
            LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(MMM_WORK, i))
            MMM_CUR = Mid(LEGEND, 10, 3) & Mid(LEGEND, 14, 2)
            'GRID
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 1).Header.Caption = "FU " & MMM_CUR
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 2).Header.Caption = "FS " & MMM_CUR
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 3).Header.Caption = "FC " & MMM_CUR
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 1).Header.Caption = "LU " & MMM_CUR
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 2).Header.Caption = "LS " & MMM_CUR
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 3).Header.Caption = "LC " & MMM_CUR
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 1).Format = "#,##0"
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 2).Format = "#,##0"
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 3).Format = "#,##0"
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 1).Format = "#,##0"
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 2).Format = "#,##0"
            grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 3).Format = "#,##0"

            ' QUERY
            grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 1).Header.Caption = "FU " & MMM_CUR
            grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 2).Header.Caption = "FS " & MMM_CUR
            grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 3).Header.Caption = "FC " & MMM_CUR
            grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 1).Header.Caption = "LU " & MMM_CUR
            grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 2).Header.Caption = "LS " & MMM_CUR
            grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 3).Header.Caption = "LC " & MMM_CUR

            If chkUnits.Checked = True Then
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 1).Hidden = False
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 1).Hidden = False
                If chkLY.Checked Then
                    grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 1).Hidden = False
                    grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 1).Hidden = False
                Else
                    grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 1).Hidden = True
                    grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 1).Hidden = True
                End If
            Else
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 1).Hidden = True
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 1).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 1).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 1).Hidden = True
            End If

            If chkSales.Checked = True Then
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 2).Hidden = False
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 2).Hidden = False

                If chkLY.Checked Then
                    grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 2).Hidden = False
                    grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 2).Hidden = False
                Else
                    grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 2).Hidden = True
                    grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 2).Hidden = True
                End If
            Else
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 2).Hidden = True
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 2).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 2).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 2).Hidden = True
            End If

            If chkCosts.Checked = True Then
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 3).Hidden = False
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 3).Hidden = False
                If chkLY.Checked Then
                    grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 3).Hidden = False
                    grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 3).Hidden = False
                Else
                    grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 3).Hidden = True
                    grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 3).Hidden = True
                End If
            Else
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 13 + 3).Hidden = True
                grdASTSRPT1.DisplayLayout.Bands(0).Columns((i * 3) + 88 + 3).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 5 + 3).Hidden = True
                grdASTDSQL1.DisplayLayout.Bands(0).Groups((i * 3) + 80 + 3).Hidden = True
            End If
        Next

    End Sub
End Class