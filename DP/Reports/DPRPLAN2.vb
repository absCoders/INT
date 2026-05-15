Public Class DPRPLAN2

    Private md(12) As String
    Private typ As String = String.Empty
    Private LYP As String = String.Empty
    Private ZYP As String = String.Empty
    Private ASWSRPT0 As String = String.Empty
    Private ICTITEM1 As String = String.Empty
    Private recCount As Int64 = 0

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()
        Prepare_dst(True)
    End Sub

    Public Overrides Sub Print_Report()

        Dim sortBy As String = String.Empty
        Dim i As Integer = 0

        If dst.Tables("ASTGROUP").Rows.Count = 0 Then

        End If

        If Not chkSUP.Checked Then
            Select Case optSORT.Value
                Case "I"
                    sortBy = ", Sorted by Item"
                Case "V"
                    sortBy = ", Sorted by Variance"
                Case "P"
                    sortBy = ", Sorted by Variance %"
            End Select
        End If

        Select Case optSHOW.Value
            Case "U"
                sortBy = "Showing Units" & sortBy
            Case "S"
                sortBy = "Showing Sales" & sortBy
            Case "C"
                sortBy = "Showing CGS" & sortBy
        End Select

        CR_params.Add("SUBT", sortBy)
        CR_params.Add("FCONLY", IIf(chkFCONLY.Checked, "1", "0"))

        For i = 1 To 12
            CR_params.Add("MD" & Format$(i, "00"), md(i))
        Next i

        CR_params.Add("SUP", IIf(chkSUP.Checked, "1", "0"))

        CR_params.Add("MD1", Get_MMM_YY(ASCMAIN1.Period_Calc(RYP, -1 * Val(numPERC.Value) + 1), 1))
        CR_params.Add("MD2", Get_MMM_YY(RYP, 1))
        CR_params.Add("MOS", numMonths.Value.ToString)


        'Dim CRSortFields As CRPEAuto.SortFields
        'CRSortFields = CR_Rpt.RecordSortFields

        'CRSortFields.Add(crDescendingOrder, "{ASWSRPT1.SORT_BY}")

        'CRSortFields.Add(crAscendingOrder, "{ASWSRPT1.ITEM_CODE}")

        Generate_Report("DPRFCVR1", String.Empty, String.Empty)
    End Sub

    Overrides Function Prepare_dst( _
          ByVal perform_fill As Boolean, _
          ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then
            Clear_dst()
        End If

        Dim sql As String = String.Empty

        With dst
            sql = "Select ITEM_CODE, ITEM_DESC from ICTITEM1"
            ICTITEM1 = ASCMAIN1.Temp_Table(sql)

            sql = "Select * from " & ICTITEM1
            Call Create_TDA(.Tables.Add, "ICTITEM1", sql, 0, False, "", 1)

        End With

        MyBase.Create_Lookup("ICTITEM1")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        EnforceConstraints(False)

        Dim sql As String = String.Empty
        Dim strTemp As String = String.Empty
        Dim VZ As String = String.Empty

        ' Get Run-Time options
        ASCMAIN1.Progress("Run-Time Options", "")

        RYPLEGEND = cmbYP0.Value
        RYP = Mid$(RYPLEGEND, 1, 4) & Mid$(RYPLEGEND, 6, 2)
        LYP = ASCMAIN1.Period_Calc(RYP, -11)
        Call Period_Range(LYP, RYP, 0, ZYP, "")
        ReDim md(12)

        For i = 1 To 12
            strTemp = ASCMAIN1.Get_YYYYMM(Mid$(ZYP, (i - 1) * 6 + 1, 6), 0)
            strTemp = Mid$(strTemp, 5, 2) & "/01/" & Mid$(strTemp, 3, 2)
            md(i) = CDate(strTemp).ToString("MMMyy")
        Next i

        ' Set up Work File Definition using X's and 0's and 0.01's as required

        ASCMAIN1.Progress("Initialize Work Tables", "")
        sql = "Select 0 RECORD_NO, 0 SORT_BY, LPAD ('X',22) ITEM_CODE"
        For i = 0 To 12
            sql &= ", 0.01 F_" & Format$(i, "00")
            sql &= ", 0.01 S_" & Format$(i, "00")
            sql &= ", 0.01 V_" & Format$(i, "00")
            sql &= ", 0.01 P_" & Format$(i, "00")
        Next i

        ' G1thru9
        For i = 1 To 9
            sql &= ", LPAD (' ', 50) G" & i
        Next
        sql &= " from DUAL where ROWNUM < 1"
        ASWSRPT0 = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & ASWSRPT0 & " ADD PRIMARY KEY (RECORD_NO)")
        Call Create_TDA(dst.Tables.Add, ASWSRPT0, "*", 1)

        sql = "SELECT * FROM " & ASWSRPT0
        Call Create_TDA(dst.Tables.Add, "ASWSRPT0", sql, 0, False, String.Empty, 0)

        ' Forecasts
        ASCMAIN1.Progress("Forecast Data", "")
        typ = "F"
        Call Get_SQL(typ)

        Select Case optSHOW.Value

            Case "U"
                VZ = "DPTITMF1.FORECAST"
            Case "S"
                VZ = "DPTITMF1.FORECAST * ICTITEM1.ITEM_RETAIL_PRICE * .6"
            Case "C"
                VZ = "DPTITMF1.FORECAST * ICTCOSTA.ITEM_COST_TOTAL"
        End Select

        sql = "Select " & sql_SELECT_cols & ",DPTITMF1.ITEM_CODE"
        For i = 1 To 12
            strTemp = Mid$(ZYP, (i - 1) * 6 + 1, 6)
            sql &= ", Sum (Decode (DPTITMF1.OPS_YYYYPP,'" & strTemp & "'," & VZ & ",0)) X" & Format$(i, "00")
        Next i

        sql &= " from DPTITMF1" & sql_TABLE_NAMEs

        If optSHOW.Value = "C" Then
            sql &= ", ICTCOSTA"
        End If

        sql &= " where DPTITMF1.OPS_YYYYPP <= '" & RYP & "'"
        sql &= "   and DPTITMF1.OPS_YYYYPP >= '" & LYP & "'"
        sql &= "   and DPTITMF1.OPS_YYYYPP_FC = DPTITMF1.OPS_YYYYPP"

        If optSHOW.Value = "C" Then
            sql &= " and ICTCOSTA.OPS_YYYYPP = DPTITMF1.OPS_YYYYPP"
            sql &= " and ICTCOSTA.ITEM_CODE = DPTITMF1.ITEM_CODE"
        End If

        sql &= sql_JOIN
        sql &= sql_WHERE

        sql &= " group by " & sql_GROUP_BY_cols & ",DPTITMF1.ITEM_CODE"
        ProcessRecords(sql)

        ' Shipments
        ASCMAIN1.Progress("Shipments Data", "")

        typ = "S"
        Call Get_SQL(typ)

        Select Case optSHOW.Value
            Case "U"
                VZ = "SOTINVH2.ORDR_QTY_SHIP"
            Case "S"
                VZ = "NVL(SOTINVH2.ORDR_QTY_SHIP, 0) * NVL(SOTINVH2.ORDR_UNIT_PRICE, 0)"
            Case "C"
                VZ = "NVL(SOTINVH2.ORDR_QTY_SHIP, 0) * NVL (SOTINVH2.ITEM_UNIT_COST, 0)"
        End Select


        sql = "Select " & sql_GROUP_BY_cols & ",SOTINVH2.ITEM_CODE"

        For i = 1 To 12
            strTemp = Mid$(ZYP, (i - 1) * 6 + 1, 6)
            sql &= ", Sum (Decode (ORDR_YYYYPP_UPDATED,'" & strTemp & "'," & VZ & ",0)) X" & Format$(i, "00")
        Next i

        sql &= " from SOTINVH2" & sql_TABLE_NAMEs

        sql &= " where SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'"
        sql &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & LYP & "'"
        sql &= "   and INV_TYPE = 'I'"
        sql &= sql_JOIN
        sql &= sql_WHERE
        sql &= " group by " & sql_GROUP_BY_cols & ",SOTINVH2.ITEM_CODE"

        ProcessRecords(sql)

        ASCMAIN1.Progress("Summarize data", "")
        Update_Record_TDA(ASWSRPT0)
        MyBase.Fill_Records("ASWSRPT0", String.Empty, True, "SELECT * FROM " & ASWSRPT0)

        ' Place data in ASWSRPT1
        sql = G1thru9 & ",ITEM_CODE "
        For i = 0 To 12
            sql &= ", F_" & Format$(i, "00")
            sql &= ", S_" & Format$(i, "00")
            sql &= ", V_" & Format$(i, "00")
            sql &= ", P_" & Format$(i, "00")
        Next i

        Dim sqlx As String = sql

        sql = "SELECT " & sql & " from " & ASWSRPT0

        For i As Integer = 1 To 9
            ASCDATA1.ExecuteSQL("alter table " & ASTSRPT1 & " modify G" & i.ToString.Trim & " VARCHAR2(50)")
        Next

        ASCDATA1.ExecuteSQL("Insert Into " & ASTSRPT1 & " (" & sqlx & ") (" & sql & ")")

        EnforceConstraints(True)

    End Sub

    ''' <summary>
    ''' This function will return a string of {how_many} ascending periods
    ''' beginning with YP (if {how_many} is positive)
    ''' or ending with YP (if {how_many} is negative)
    ''' </summary>
    ''' <param name="yp"></param>
    ''' <param name="How_Many"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function Get_MMM_YY(ByVal yp As String, ByVal How_Many As Integer) As String

        Dim result As String = String.Empty
        Dim incr As Integer = 0
        Dim iCount As Integer = 0
        Dim dateString As String = 0

        For iCount = 1 To Math.Abs(How_Many)
            result = ASCMAIN1.Get_YYYYMM(yp, incr)
            dateString = Mid$(result, 5, 2) & "/01/" & Mid$(result, 3, 2)

            If Math.Sign(How_Many) = 1 Then
                result = CDate(dateString).ToString("MMM yy")
            Else
                result = CDate(dateString).ToString("MMM yy")
            End If

            incr = incr + Math.Sign(How_Many)
        Next iCount

        Return result

    End Function

    Private Sub Period_Range(ByVal zz1 As String, ByVal zz2 As String, ByRef zzn As Integer, ByRef zzp As String, ByVal zzs As String)
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

        For Each rowData As DataRow In ASCDATA1.GetDataTable(query).Rows
            recCount += 1
            If recCount Mod 100 = 0 Then
                ASCMAIN1.Progress("-", recCount)
            End If
            WriteRecords(rowData)
        Next
    End Sub

    Private Sub WriteRecords(ByRef rowData As DataRow)

        Dim ITEM_CODE As String = rowData.Item("ITEM_CODE")
        Dim rowICTITEMX As DataRow = Nothing
        Dim rowICTITEM1 As DataRow = Nothing

        Dim rowASWSRPT0 As DataRow = dst.Tables(ASWSRPT0).NewRow
        rowASWSRPT0.Item("RECORD_NO") = recCount
        rowASWSRPT0.Item("ITEM_CODE") = ITEM_CODE
        rowASWSRPT0.Item("SORT_BY") = 0
        dst.Tables(ASWSRPT0).Rows.Add(rowASWSRPT0)

        SetKey(rowASWSRPT0, rowData)

        For i As Integer = 1 To 12
            rowASWSRPT0.Item(typ & "_" & Format$(i, "00")) = rowData.Item("X" & Format$(i, "00")) & String.Empty
        Next i

        If dst.Tables("ICTITEM1").Select("ITEM_CODE = '" & ITEM_CODE & "'").Length = 0 Then
            rowICTITEMX = MyBase.LookUp("ICTITEM1", ITEM_CODE)
            rowICTITEM1 = dst.Tables("ICTITEM1").NewRow

            rowICTITEM1.Item("ITEM_CODE") = ITEM_CODE
            rowICTITEM1.Item("ITEM_DESC") = rowICTITEMX.Item("ITEM_DESC")
            dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)
        End If

    End Sub

    Private Sub SetKey(ByRef rowASWSRPT0 As DataRow, ByRef rowData As DataRow)

        Dim fieldName As String = String.Empty
        For fieldNum As Integer = 0 To COLUMN_NAMEs.Count - 1
            'rowASWSRPT0.Item("G" & fieldNum + 1) = COLUMN_NAMEs(fieldNum) & ":" & rowData(COLUMN_NAMEs(fieldNum)) & String.Empty
            rowASWSRPT0.Item("G" & fieldNum + 1) = rowData(COLUMN_NAMEs(fieldNum)) & String.Empty
        Next
    End Sub

    Public Overrides Sub Build_Report_File_Post_Process()
        MyBase.Build_Report_File_Post_Process()


        Dim TOTF As Long = 0
        Dim TOTS As Long = 0
        Dim FVAR As Long = 0
        Dim FVARP As Long = 0

        Dim xf As Long = 0
        Dim xs As Long = 0

        Dim ITEM_CODE As String = String.Empty
        Dim iCtr As Long = 0
        Dim strTemp As String = String.Empty

        ' Calculate Report Totals
        ASCMAIN1.Progress("Report Calculations", "")
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Select("", "")
            ITEM_CODE = rowASTSRPT1.Item("ITEM_CODE")

            TOTF = 0
            TOTS = 0

            iCtr += 1
            If iCtr Mod 50 = 0 Then
                ASCMAIN1.Progress("-", iCtr)
            End If

            For i As Integer = 12 To 0 Step -1
                strTemp = "_" & Format$(i, "00")
                If i > 0 And i > 12 - numPERC.Value Then
                    TOTF = TOTF + Val(rowASTSRPT1.Item("F" & strTemp) & "")
                    TOTS = TOTS + Val(rowASTSRPT1.Item("S" & strTemp) & "")
                End If

                If i = 0 Then
                    rowASTSRPT1.Item("F_00") = TOTF
                    rowASTSRPT1.Item("S_00") = TOTS
                End If

                xs = Val(rowASTSRPT1.Item("S" & strTemp) & "")
                xf = Val(rowASTSRPT1.Item("F" & strTemp) & "")

                FVAR = xs - xf

                rowASTSRPT1.Item("V" & strTemp) = FVAR
                If xf = 0 Then
                    FVARP = 0
                Else
                    FVARP = 100 * FVAR / xf
                End If
                rowASTSRPT1.Item("P" & strTemp) = FVARP
            Next

            If (chkFCONLY.Checked AndAlso TOTF = 0) Or (TOTF <> 0 And Math.Abs(FVARP) < numPERC.Value) Then
                rowASTSRPT1.Delete()
            Else
                If optSORT.Value = "V" Then
                    rowASTSRPT1.Item("SORT_BY") = Math.Abs(FVAR)
                Else
                    If optSORT.Value = "P" Then
                        rowASTSRPT1.Item("SORT_BY") = Math.Abs(FVARP)
                    End If
                End If
            End If
        Next
    End Sub

#Region "VB6"

    Dim RYP As String
    Dim RYPLegend As String
    Dim LYP As String
    Dim XYP As String
    Dim ZYP As String
    Dim UYP As String
    Dim md() As String
    Dim Lmd() As String
    Dim selSNU As String
    Dim selBP As String
    Dim chkUSC As String
    Dim optSHOW As String
    Dim chkLY As String
    Dim chkSUP As String
    Dim chkIGNOREPDF As String
    Const Num_Months = 12

    Sub Build_Workfile()

        Call Build_WorkFile_DB_Init()

        Dim i As Integer
        Dim j As Integer
        Dim z As String
        Dim zz As String
        Dim r As Long
        Dim typ As String
        Dim total As Long

        Dim item As String
        Dim Periods As String
        Dim sqlz As String

        ' Get Codes selected from tabs

        Call Track("Run-Time Options", "")

        ' Get Fields selected in Sort & Sub-Total Sequence

        AZ = "" ' including hard-coded last level, if appl
        Call Get_Codes(AZ, cfmax, chkNewPage, cf(), codes(), "")

        ' Get Run-Time options

        'selSNU = Get_chk("SNU", "SNU", "Y")
        'selBP = Get_chk("BP", "BP", "Y")
        If objASCCALLB.xErrMsg <> "" Then
            Exit Sub
        End If

        ' Set up Work File Definition using X's and 0's and 0.01's as required

        Call Track("Initialize Work Tables", "")
        sql = "Select 0 RECORD_NO, LPAD ('X',22) ITEM_CODE, LPAD ('X',6) OPS_YYYYPP, 0 FC_PD "
        For i = 1 To Num_Months
            sql = sql & ", 0 FC_" & Format$(i, "00")
        Next i
        sql = sql & ", "
        sql = sql & Get_GString(1, 30, AZ)
        sql = sql & " from DUAL where ROWNUM < 1"
        sqlz = sql
        Call Ora_to_Acc(Nothing, "ASWSRPT0", 1, "", sql)
        Call Create_Index("ASWSRPT0", "Report", Get_GString(2, 0, AZ))

        Dim dynASWSRPT0 As Recordset
        dynASWSRPT0 = AccD.OpenRecordset("ASWSRPT0", dbOpenTable)
        dynASWSRPT0.Index = "PrimaryKey"

        sql = "Select ITEM_CODE, ITEM_DESC, ITEM_UOM"
        sql = sql & " from ICTITEM1 where ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "ICWITEM1", 1, "", sql)
        Dim tblICWITEM1 As Recordset
        tblICWITEM1 = AccD.OpenRecordset("ICWITEM1", dbOpenTable)
        tblICWITEM1.Index = "PrimaryKey"

        ' Set up ASWGROUP w/codes & descs from all small tables
        Call Get_Group_Desc_All(codes())

        'Call Get_Group_Desc("A", codes(), "SALES_DIVISION_CODE, SALES_DIVISION_CODE, SALES_DIVISION_NAME from SOTSDIV1")
        'Call Get_Group_Desc("B", codes(), "ITEM_BRAND_CODE, ITEM_BRAND_CODE, ITEM_BRAND_NAME from ICTBRAN1")
        'Call Get_Group_Desc("C", codes(), "MARKET_CODE, MARKET_CODE, MARKET_DESC from SOTMKTC1")
        'Call Get_Group_Desc("F", codes(), "ITEM_CLASS_CODE, ITEM_CLASS_CODE, ITEM_CLASS_DESC from ICTCLAS1")

        ' Set up ASWGROUP w/codes & descs from hard-coded codes
        Dim grp As String
        Dim code As String
        Dim codedesc As String

        grp = codes(2, 3)
        code = "S"
        codedesc = "Saleable"
        Call Write_Group_Record(grp, code, codedesc)
        code = "N"
        codedesc = "No-Charge"
        Call Write_Group_Record(grp, code, codedesc)
        code = "U"
        codedesc = "Unfinished"
        Call Write_Group_Record(grp, code, codedesc)

        grp = codes(2, 4)
        code = "B"
        codedesc = "Basic"
        Call Write_Group_Record(grp, code, codedesc)
        code = "P"
        codedesc = "Promo"
        Call Write_Group_Record(grp, code, codedesc)

        ' Set up Misc dynasets & memory tables

        Dim dynICTITEM1 As OraDynaset
        sql = "Select * from ICTITEM1 where ITEM_CODE = :ITEM_CODE"
        dynICTITEM1 = OraD.CreateDynaset(sql, 8&)

        ' Prepare Work File with Data from Server

        Dim dynX As OraDynaset
        Dim fldX() As Object

        ' Forecasts

        typ = "F"
        Call Get_SQL(typ, cfmax, cf, codes, xInfo)
        If objASCCALLB.xErrMsg <> "" Then
            Exit Sub
        End If

        Call Track("Forecast Data", "")

        ' Get list of 12 periods from now
        z = Period_Calc(CYP, Num_Months)
        Call Period_Range(CYP, z, 0, Periods, "")

        ReDim md(Num_Months)
        For i = 1 To Num_Months
            z = DateAdd("M", i - 1, Mid$(CYP, 5, 2) & "/01/" & Mid$(CYP, 3, 2))
            z = Format$(z, "YYYYMM")
            z = Get_YYYYMM(z, 0)
            z = Mid$(z, 5, 2) & "/01/" & Mid$(z, 3, 2)
            md(i) = Format$(z, "mmm'YY")
        Next i

        z = Period_Calc(CYP, Num_Months * -1)

        'sql = "Select OPS_YYYYPP "
        sql = "Select " & sqllist & ", OPS_YYYYPP"
        sql = sql & ", Decode (OPS_YYYYPP_FC,'000000',FORECAST,0) FC_PD"
        For i = 1 To Num_Months
            sql = sql & ", Decode (OPS_YYYYPP_FC,'" & Mid$(Periods, (i - 1) * 6 + 1, 6) & "',FORECAST,0) FC_" & Format$(i, "00")
        Next i
        sql = sql & " from MRTITMF1 X, ICTITEM1" & sqltables
        sql = sql & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE"
        sql = sql & " AND (OPS_YYYYPP > '" & z & "' AND OPS_YYYYPP <= '" & CYP & "')"
        sql = sql & sqljoin
        sql = sql & sqlwhere
        ''sql = sql & " group by " & sqllist2 & ", OPS_YYYYPP DESC"
        sql = sql & " order by " & sqllist2 & ", OPS_YYYYPP DESC"

        sqlz = sql
        'Call Ora_to_Acc(Nothing, "MRWITMF1", cfmax + 1, "", sql)

        dynX = OraD.CreateDynaset(sqlz, 8&)
        Call Set_Fldsx(dynX, fldX())
    GoSub Process_Records

        ' Prepare Report File, w/Consolidations & Recaps as required

        Call Build_Report_File(3, 1 + Num_Months, dynASWSRPT0, "ASWSRPT0", cfmax, AZ, chkRecap, "", cf(), "ITEM_CODE, OPS_YYYYPP")

        ' Wrap up

        dynASWSRPT0.Close()

        tblICWITEM1.Close()

        Exit Sub

Process_Records:
        Do While Not dynX.EOF
            total = 0
            For i = 1 To Num_Months
                total = total + dynX.Fields("FC_" & Format$(i, "00")).Value & ""
            Next
            total = total + dynX.Fields("FC_PD").Value & ""
            If total > 0 Then
                r = r + 1
                If r Mod 100 = 0 Then
                    z = vtos(fldX(0).Value)
                    Call Track("-", z & ":" & CStr(r))
                End If
            GoSub Write_Records
            End If
            dynX.MoveNext()
        Loop
        Return

Write_Records:

        item = dynX.Fields("ITEM_CODE").Value
        tblICWITEM1.Seek("=", item)
        If tblICWITEM1.NoMatch Then
            OraD.Parameters("ITEM_CODE").Value = item
            dynICTITEM1.Refresh()
            tblICWITEM1.AddNew()
            tblICWITEM1.Fields("ITEM_CODE").Value = item
            tblICWITEM1.Fields("ITEM_DESC").Value = dynICTITEM1.Fields("ITEM_DESC").Value
            tblICWITEM1.Fields("ITEM_UOM").Value = dynICTITEM1.Fields("ITEM_UOM").Value
            tblICWITEM1.Update()
        End If

        If typ = "F" Then
            dynASWSRPT0.AddNew()
            dynASWSRPT0.Fields("RECORD_NO").Value = r
            dynASWSRPT0.Fields("ITEM_CODE").Value = item
            dynASWSRPT0.Fields("FC_PD").Value = dynX.Fields("FC_PD").Value
            dynASWSRPT0.Fields("OPS_YYYYPP").Value = dynX.Fields("OPS_YYYYPP").Value
        GoSub Set_Key

            For i = 1 To Num_Months
                dynASWSRPT0.Fields("FC_" & Format$(i, "00")).Value = dynX.Fields("FC_" & Format$(i, "00")).Value
            Next i
        Else
            ' Add In Zero Records
        End If
        dynASWSRPT0.Update()

        Return

Set_Key:
        For j = 1 To cfmax
            zz = fldX(j - 1).Value & ""
            z = Format$(j, "0")
            dynASWSRPT0.Fields("G" & z).Value = cf(2, j) & ":" & zz
        Next j
        Return

    End Sub

    Sub Print_Report()
        Dim z As String
        Dim i As Integer
        Dim j As Integer

        Call Std_Report_Parameters()

        Select Case optSHOW
            Case "F"
                z = "Showing Monthly Forecasts"
            Case "C"
                z = "Showing Changes in Monthly Forecasts"
        End Select

        CR_Rpt.ParameterFields(CR_Rpt_Names("SUBT")).SetCurrentValue(z)
        CR_Rpt.ParameterFields(CR_Rpt_Names("LVLS")).SetCurrentValue(CStr(cfmax))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG1")).SetCurrentValue(cf(2, 1))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG2")).SetCurrentValue(cf(2, 2))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG3")).SetCurrentValue(cf(2, 3))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG4")).SetCurrentValue(cf(2, 4))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG5")).SetCurrentValue(cf(2, 5))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG6")).SetCurrentValue(cf(2, 6))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG7")).SetCurrentValue(cf(2, 7))
        CR_Rpt.ParameterFields(CR_Rpt_Names("RECAP")).SetCurrentValue(chkRecap)
        CR_Rpt.ParameterFields(CR_Rpt_Names("NEWPAGE")).SetCurrentValue(chkNewPage)
        CR_Rpt.ParameterFields(CR_Rpt_Names("RC")).SetCurrentValue(aRC)
        CR_Rpt.ParameterFields(CR_Rpt_Names("CHKUSC")).SetCurrentValue(chkUSC)
        CR_Rpt.ParameterFields(CR_Rpt_Names("CHKLY")).SetCurrentValue(chkLY)
        CR_Rpt.ParameterFields(CR_Rpt_Names("CHKSUP")).SetCurrentValue(chkSUP)
        CR_Rpt.ParameterFields(CR_Rpt_Names("SHOWPD")).SetCurrentValue("")

        For i = 1 To 12
            CR_Rpt.ParameterFields(CR_Rpt_Names("MD" & Format$(i, "00"))).SetCurrentValue(md(i))
        Next i

        Call Std_Report_Parameters(True)

        Call Prepare_SPRF()
    End Sub

#End Region
End Class