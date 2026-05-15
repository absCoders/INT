Public Class RSRBREV1
    Dim tblARTCUST1 As DataTable
    Dim rowARTCUST1 As DataRow

    Dim rowGLTPARM3 As DataRow
    Dim RYM As String
    Dim REL_WEEK As Integer
    Dim WEEK_LEGEND As String

    Dim RYW_LY As String
    Dim RYW_2LY As String
    Dim RYM01 As String

    Dim SEASON_CODE_PREV As String
    Dim SEASON_CODE_LY As String
    Dim SEASON_CODE_NEXT As String
    Dim SEASON_CODE As String
    Dim SEASON_YEAR As String
    Dim SEASON_YEAR_LY As String
    Dim SEASON_TYPE As String
    Dim SEASON_YEAR_NY As String

    Dim CUST_NAME As String
    Dim CUST_CODE As String
    Dim CUST_CODEs As New List(Of String)
    Dim sqlwCUST_CODE As String

#Region "Thread"

    Dim RSTBREV1 As String
    Dim RSTSPRO2 As String
    Dim RSTSPRO3 As String
    Dim RSTSPRO4 As String
    Dim RSTSPRO5 As String
    Dim RSTSPRO6 As String
    Dim RSTSPRO7 As String
    Dim RSTSPRO7C As String
    Dim SOTALLOX As String
    Dim SPTCOOPX As String
    Dim SATAUTHX As String

#End Region
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1)

        'ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
        'tblARTCUST1 = ASCDATA1.GetDataTable


        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, CUST_ADDR1, CUST_CITY, CUST_STATE from ARTCUST1"
        tblARTCUST1 = ASCDATA1.GetDataTable
        tblARTCUST1.Columns.Add("SEL")
        tblARTCUST1.Columns("SEL").DefaultValue = "0"

        grdARTCUST1.DataSource = tblARTCUST1
        Sort_grdColumns(grdARTCUST1, "CUST_CODE")
        With grdARTCUST1.DisplayLayout.Bands(0)
            .Columns("SEL").Header.Fixed = True

            For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                With gcol.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = System.Drawing.Color.Gold
                End With

                If gcol.Key = "SEL" Then
                    gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    gcol.CellActivation = UltraWinGrid.Activation.NoEdit
                    gcol.CellAppearance.BackColor = System.Drawing.Color.WhiteSmoke
                End If
            Next
        End With

        Create_Summary(grdARTCUST1, "SEL")
        Create_Summary(grdARTCUST1, "CUST_CODE", "Count")
        grdARTCUST1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Show_Filter(grdARTCUST1, True)
        grdARTCUST1.DisplayLayout.GroupByBox.Hidden = True

    End Sub

    Protected Overrides Sub Build_Workfile()
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Building Work File")

        CUST_CODE = txtCUST_CODE.Text

        If 1 = 1 Then
            Main_Process_1()
        Else

            'create a thread to handle communication with connected client
            Dim clientThread As New System.Threading.Thread(New System.Threading.ParameterizedThreadStart(AddressOf HandleClientComm))

            Dim dic As New Dictionary(Of String, Object)
            dic.Add("CUST_CODE", CUST_CODE)
            dic.Add("RYW", RYW)

            clientThread.Start(dic) ' dic

            ' PACK ALL UI SETTINGS INTO A CLASS OR KEY/VALUE PAIR

            ' INSTANTIATE RSCSPRO1 ON A NEW THREAD

            ' MAKE SURE FORM IS DISABLED
            ' ADD KILL SWITCH

        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("", "")
    End Sub

    Private Sub HandleClientComm(client As Object)
        Dim ads As New AppDomainSetup()
        ads.ShadowCopyFiles = True

        Try
            ' EventLog.WriteEntry(eventLogSource, "Creating new app domain...", EventLogEntryType.Information, 4)
            Dim ad = AppDomain.CreateDomain("", Nothing, ads)

            ' EventLog.WriteEntry(eventLogSource, "Creating ABS control class...", EventLogEntryType.Information, 4)
            Dim directory = AppDomain.CurrentDomain.BaseDirectory
            Dim ABSDomain = CType(ad.CreateInstanceFromAndUnwrap(directory & "ABSDomain.dll", "ABSDomain.ABSDomain", True, System.Reflection.BindingFlags.CreateInstance, Nothing, Nothing, Nothing, Nothing), ABSDomain)
            ' EventLog.WriteEntry(eventLogSource, "Starting ABS control class...", EventLogEntryType.Information, 4)
            ABSDomain.Start(Nothing)
            '  EventLog.WriteEntry(eventLogSource, "Unloading app domain...", EventLogEntryType.Information, 100)
            AppDomain.Unload(ad)
        Catch ex As Exception
            '  EventLog.WriteEntry(eventLogSource, String.Format("App domain error: {0}", ex.Message), EventLogEntryType.Error, 50)
            MsgBox(ex.Message)
        End Try
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = CUST_CODE
        If CUST_CODEs.Count > 1 Then
            SUBT = Join(CUST_CODEs.ToArray, ",")
        End If
        'CR_params.Add("WEEK_LEGEND", WEEK_LEGEND)
        'Generate_Report(RPT, "Retail Sales Overview", SUBT)

        'RPT = "RSRSPRO2"
        'CR_params.Add("WEEK_LEGEND", WEEK_LEGEND)
        'CR_params.Add("TY", Mid(RYW, 1, 4))
        'CR_params.Add("LY", Format(Val(Mid(RYW, 1, 4)) - 1, "0000"))

        Dim MO As String = Format(Val(Mid(RYM, 5, 2)) - 1, "00")
        If MO = "00" Then MO = "12"
        '   CR_params.Add("MO", MO)

        'Generate_Report(RPT, "Retail Sales Seasonality", SUBT)


        'RPT = "RSRSPRO3"
        'CR_params.Add("WEEK_LEGEND", WEEK_LEGEND)
        'CR_params.Add("WEEK_END_DATE", Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd/yy"))
        'Generate_Report(RPT, "Retail Sales & Inventory Overview - Top 100 Items", SUBT)

        'Exit Sub


        RPT = "RSRBREV0"
        CR_params.Add("WEEK_LEGEND", WEEK_LEGEND)
        CR_params.Add("TY", Mid(RYW, 1, 4))
        CR_params.Add("LY", Format(Val(Mid(RYW, 1, 4)) - 1, "0000"))
        CR_params.Add("LY2", Format(Val(Mid(RYW, 1, 4)) - 2, "0000"))
        CR_params.Add("MO", MO)
        Generate_Report(RPT, "Business Review", SUBT)


    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If txtCUST_CODE.Text = "" Then
                EMsg &= vbCr & "You Must Specify a Customer"
            Else
                If LookUp("ARTCUST1", txtCUST_CODE.Text) Is Nothing Then
                    EMsg &= vbCr & "Customer Specified (" & txtCUST_CODE.Text & ") is Invalid"
                Else
                    CUST_CODEs.Clear()

                    If tblARTCUST1.Select("SEL='1'").Length <> 0 Then
                        For Each row As DataRow In tblARTCUST1.Select("SEL='1'")
                            If row.Item("CUST_CODE") = txtCUST_CODE.Text Then
                                EMsg &= vbCr & "You Cannot Select " & txtCUST_CODE.Text & " as an Additional Customer"
                                Exit For
                            Else
                                CUST_CODEs.Add(row.Item("CUST_CODE"))
                            End If
                        Next
                    End If

                    CUST_CODEs.Add(txtCUST_CODE.Text)
                End If
            End If
        End If 
    End Sub

    Private Sub txtCUST_CODE_ValueChanged(sender As Object, e As EventArgs) Handles txtCUST_CODE.ValueChanged
        'Dim CUST_CODE = txtCUST_CODE.Text
        'If CUST_CODE = "" OrElse tblARTCUST1.Rows.Find(CUST_CODE) Is Nothing Then
        'Else
        'End If
    End Sub

#Region "Thread"
    Sub Main_Process_1()

        Create_Worktable()

        ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, ITEM_BASIC_PROMO" & vbCrLf _
            & ", COLLECTION_CODE, PROD_CODE, ITEM_RETAIL_PRICE, ITEM_EAN_CODE, ITEM_SNU_CODE" & vbCrLf _
            & " from ICTITEM1 where ITEM_CODE in " & vbCrLf _
            & " (Select Distinct ITEM_CODE from " & RSTBREV1 & vbCrLf _
            & " union Select Distinct ITEM_CODE from " & RSTSPRO4 & vbCrLf _
            & " union Select Distinct ITEM_CODE from " & SOTALLOX & ")"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False)
        Fill_Records("ICTITEM1")

        ASCMAIN1.sql = "Select * from " & SATAUTHX
        Create_TDA(dst.Tables.Add, "SATAUTHX", "**", 0, False)
        Fill_Records("SATAUTHX") ', CUST_CODE)

        ASCMAIN1.sql = "Select * from ICTCOLL1"
        Create_TDA(dst.Tables.Add, "ICTCOLL1", "**", 0, False)
        Fill_Records("ICTCOLL1")

        ASCMAIN1.sql = "Select * from ICTCOLL0"
        Create_TDA(dst.Tables.Add, "ICTCOLL0", "**", 0, False)
        Fill_Records("ICTCOLL0")

        ASCMAIN1.sql = "Select * from ICTBRAN1"
        Create_TDA(dst.Tables.Add, "ICTBRAN1", "**", 0, False)
        Fill_Records("ICTBRAN1")

        ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME" & vbCrLf _
            & " from ARTCUST1 where CUST_CODE = :PARM1"
        Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0, False, "V")
        Fill_Records("ARTCUST1", CUST_CODE)
        rowARTCUST1 = dst.Tables("ARTCUST1").Rows(0)
        CUST_NAME = rowARTCUST1.Item("CUST_NAME")

        ASCMAIN1.sql = "Select * from " & RSTSPRO2
        Create_TDA(dst.Tables.Add, "RSTSPRO2", "**", 0, False)
        Fill_Records("RSTSPRO2")
        With dst.Tables("RSTSPRO2")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
            .Columns.Add("RANK", GetType(System.Int64))
            .Columns("RANK").DefaultValue = 0
        End With

        ASCMAIN1.sql = "Select * from " & RSTSPRO3
        Create_TDA(dst.Tables.Add, "RSTSPRO3", "**", 0, False)
        Fill_Records("RSTSPRO3")
        With dst.Tables("RSTSPRO3")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
        End With

        ASCMAIN1.sql = "Select * from " & RSTSPRO4
        Create_TDA(dst.Tables.Add, "RSTSPRO4", "**", 0, False, "", 3)
        Fill_Records("RSTSPRO4")

        ASCMAIN1.sql = "Select * from " & RSTSPRO7 & " where TY_STD454 <> 0 or LY_STD454 <> 0"
        Create_TDA(dst.Tables.Add, "RSTSPRO7", "**", 0, False, "", 2)
        Fill_Records("RSTSPRO7")
        With dst.Tables("RSTSPRO7")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
            .Columns.Add("RANK", GetType(System.Int64))
            .Columns("RANK").DefaultValue = 0
        End With

        ASCMAIN1.sql = "Select * from " & RSTSPRO7C & " where TY_STD454 <> 0 or LY_STD454 <> 0"
        Create_TDA(dst.Tables.Add, "RSTSPRO7C", "**", 0, False, "", 2)
        With dst.Tables("RSTSPRO7C")
            Dim PLUS2 As String = "" ' ADD 2 MONTHS PRIOR TO THE SEASON
            If SEASON_TYPE = "F" Then
                PLUS2 = "+ISNULL(TY_M05,0)+ISNULL(TY_M06,0)"
            Else
                PLUS2 = "+ISNULL(LY_M11,0)+ISNULL(LY_M12,0)"
            End If
            .Columns.Add("TY_STD454_PLUS2", GetType(System.Decimal), "ISNULL(TY_STD454,0)+" & PLUS2)
        End With
        Fill_Records("RSTSPRO7C")

        ASCMAIN1.sql = "Select * from " & RSTSPRO6 & " where TY_STD454 <> 0 or LY_STD454 <> 0"
        Create_TDA(dst.Tables.Add, "RSTSPRO6", "**", 0, False, "", 3)
        Fill_Records("RSTSPRO6")
        With dst.Tables("RSTSPRO6")
            .Columns.Add("TY_WTD", GetType(System.Decimal), "TY_WK" & CStr(REL_WEEK))
            .Columns.Add("LY_WTD", GetType(System.Decimal), "LY_WK" & CStr(REL_WEEK))
            .Columns.Add("RANK", GetType(System.Int64))
            .Columns("RANK").DefaultValue = 0
        End With

        Dim RANK As Integer

        RANK = 0
        For Each rowRSTSPRO2 As DataRow In dst.Tables("RSTSPRO2").Select("", "TY_STD454 DESC")
            RANK += 1
            rowRSTSPRO2.Item("RANK") = RANK
        Next

        RANK = 0
        For Each rowRSTSPRO6 As DataRow In dst.Tables("RSTSPRO6").Select("", "TY_STD454 DESC")
            RANK += 1
            rowRSTSPRO6.Item("RANK") = RANK
        Next

        RANK = 0
        For Each rowRSTSPRO7 As DataRow In dst.Tables("RSTSPRO7").Select("", "TY_STD454 DESC")
            RANK += 1
            rowRSTSPRO7.Item("RANK") = RANK
        Next


        ASCMAIN1.sql = "Select SOTALLO1.* from SOTALLO1" & vbCrLf _
            & " where ALLO_CTL_NO in (Select ALLO_CTL_NO from " & SOTALLOX & ")"
        Create_TDA(dst.Tables.Add, "SOTALLO1", "**", 0, False, "", 1)
        With dst.Tables("SOTALLO1").Columns
            .Add("SEASON_TYPE")
            .Add("SEASON_CODE")
        End With

        ASCMAIN1.sql = "Select SOTALLO2.* from SOTALLO2" & vbCrLf _
            & " where ALLO_CTL_NO in (Select ALLO_CTL_NO from " & SOTALLOX & ")" & vbCrLf _
            & sqlwCUST_CODE ' "   and CUST_CODE = '" & CUST_CODE & "'"
        Create_TDA(dst.Tables.Add, "SOTALLO2", "**", 0, False, "", 2)
        With dst.Tables("SOTALLO2").Columns
            .Add("ORDR_QTY", GetType(System.Int64))
            .Add("ORDR_QTY_OPEN", GetType(System.Int64))
            .Add("ORDR_QTY_PICK", GetType(System.Int64))
            .Add("ORDR_QTY_SHIP", GetType(System.Int64))
            .Add("ORDR_QTY_CANC", GetType(System.Int64))
            .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)-ISNULL(ORDR_QTY_OPEN,0)") ' adding qty open since we are showing this to customers
            .Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
            .Add("QTY_OVER", GetType(System.Int64), "IIF(QTY_LEFT-ISNULL(ORDR_QTY_OPEN,0)>=0,0,ISNULL(ORDR_QTY_OPEN,0)-QTY_LEFT)")
        End With

        Create_Relation("SOTALLO1", "SOTALLO2", "ALLO_CTL_NO")

        dst.Tables("SOTALLO2").Columns.Add("ITEM_CODE", GetType(System.String), "PARENT(SOTALLO1_SOTALLO2).ITEM_CODE")
        '
        ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_OPEN" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_PICK" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_SHIP" & vbCrLf _
            & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
            & "   where SOTORDR2.ALLO_CTL_NO in  (Select ALLO_CTL_NO from " & SOTALLOX & ")" & vbCrLf _
            & Replace(sqlwCUST_CODE, "and CUST_CODE", "and NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)") & vbCrLf _
            & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"
        Create_TDA(dst.Tables.Add, "SOTALLOZ", "**", 0, False, "", 2)
        '"   and SOTORDR2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _

        Fill_Records("SOTALLO1")

        For Each rowSOTALLO1 As DataRow In dst.Tables("SOTALLO1").Select("")
            'Dim DATE_START As Date = rowSOTALLO1.Item("DATE_START")
            'If Format(DATE_START, "MM") >= "01" And Format(DATE_START, "MM") <= "06" Then
            '    rowSOTALLO1.Item("SEASON_TYPE") = "S"
            '    rowSOTALLO1.Item("SEASON_CODE") = Format(DATE_START, "yyyy") & "S"
            'Else
            '    rowSOTALLO1.Item("SEASON_TYPE") = "F"
            '    rowSOTALLO1.Item("SEASON_CODE") = Format(DATE_START, "yyyy") & "F"
            'End If

            Dim DATE_END As Date = rowSOTALLO1.Item("DATE_END")
            If Format(DATE_END, "MM") >= "01" And Format(DATE_END, "MM") <= "06" Then
                rowSOTALLO1.Item("SEASON_TYPE") = "S"
                rowSOTALLO1.Item("SEASON_CODE") = Format(DATE_END, "yyyy") & "S"
            Else
                rowSOTALLO1.Item("SEASON_TYPE") = "F"
                rowSOTALLO1.Item("SEASON_CODE") = Format(DATE_END, "yyyy") & "F"
            End If
        Next

        Fill_Records("SOTALLO2")
        Fill_Records("SOTALLOZ")

        For Each rowSOTALLOZ As DataRow In dst.Tables("SOTALLOZ").Select("")
            Dim ALLO_CTL_NO As String = rowSOTALLOZ.Item("ALLO_CTL_NO")
            Dim CUST_CODE As String = rowSOTALLOZ.Item("CUST_CODE")
            Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE})
            If rowSOTALLO2 Is Nothing Then
                rowSOTALLO2 = dst.Tables("SOTALLO2").NewRow
                rowSOTALLO2.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                rowSOTALLO2.Item("CUST_CODE") = CUST_CODE
                dst.Tables("SOTALLO2").Rows.Add(rowSOTALLO2)
            End If
            For Each COLUMN_NAME As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
                rowSOTALLO2.Item(COLUMN_NAME) = rowSOTALLOZ.Item(COLUMN_NAME)
            Next
        Next

        'THIS MAY NOT BE REQUIRED ANYMORE NOW THAT WE HAVE RSTPRO7C
        For Each rowSOTALLO2 As DataRow In dst.Tables("SOTALLO2").Select("")
            Dim ITEM_CODE As String = rowSOTALLO2.Item("ITEM_CODE")
            Dim CUST_CODE As String = rowSOTALLO2.Item("CUST_CODE")
            If dst.Tables("RSTSPRO7").Rows.Find(New String() {CUST_CODE, ITEM_CODE}) Is Nothing Then
                dst.Tables("RSTSPRO7").Rows.Add(New String() {CUST_CODE, ITEM_CODE})
            End If
        Next

        'THIS IS REQUIRED SO THAT WE SEE MACYS AND MACYSCOM ALLOCATIONS SEPARATELY
        For Each rowSOTALLO2 As DataRow In dst.Tables("SOTALLO2").Select("")
            Dim ITEM_CODE As String = rowSOTALLO2.Item("ITEM_CODE")
            Dim CUST_CODE As String = rowSOTALLO2.Item("CUST_CODE")
            If dst.Tables("RSTSPRO7C").Rows.Find(New String() {CUST_CODE, ITEM_CODE}) Is Nothing Then
                dst.Tables("RSTSPRO7C").Rows.Add(New String() {CUST_CODE, ITEM_CODE})
            End If
        Next


        With dst.Tables.Add("SOTALLOR")
            With .Columns
                .Add("CUST_CODE")
                .Add("SEASON_CODE")
                .Add("ITEM_CODE")
                .Add("QTY_SOLD", GetType(System.Int64))
            End With
            .PrimaryKey = New DataColumn() {.Columns("CUST_CODE"), .Columns("SEASON_CODE"), .Columns("ITEM_CODE")}
        End With

        Dim CUST_CODEs As New List(Of String)
        For Each rowC As DataRow In ASCDATA1.SelectDistinct("SOTALLO2", "CUST_CODE").Select("")
            Dim CUST_CODE As String = rowC.Item("CUST_CODE")
            CUST_CODEs.Add(CUST_CODE)
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct("SOTALLO1", New String() {"SEASON_CODE", "ITEM_CODE"}).Select("")
            Dim SEASON_CODE As String = row.Item("SEASON_CODE")
            Dim SEASON_TYPE As String = Mid(SEASON_CODE, 5, 1)
            Dim YW1 As String = Mid(SEASON_CODE, 1, 4) & IIf(SEASON_TYPE = "F", "27", "01")
            Dim YW2 As String = Mid(SEASON_CODE, 1, 4) & IIf(SEASON_TYPE = "F", "53", "26")
            YW1 = ASCMAIN1.Week_Calc(YW1, -9)
            If YW2 > RYW Then YW2 = RYW
            Dim ITEM_CODE As String = row.Item("ITEM_CODE")
            For Each CUST_CODE As String In CUST_CODEs
                ASCMAIN1.sql = "Select Sum (QTY_SOLD) from RSTRETL1" & vbCrLf _
                    & " where ITEM_CODE = '" & ITEM_CODE & "'" & vbCrLf _
                    & "   and CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
                    & "   and OPS_YYYYWW between '" & YW1 & "' and '" & YW2 & "'"
                Dim QTY_SOLD As Int64 = Val(ASCDATA1.GetDataValue)
                dst.Tables("SOTALLOR").Rows.Add(New Object() {CUST_CODE, SEASON_CODE, ITEM_CODE, QTY_SOLD})
            Next
        Next

        ASCMAIN1.sql = "Select * from " & SPTCOOPX
        Create_TDA(dst.Tables.Add, "SPTCOOPX", "**", 0, False)
        With dst.Tables("SPTCOOPX")
            .Columns.Add("COOP", GetType(System.Decimal), "IIF(EXPENSE_TYPE_CODE='COOP',DIST_AMT,0)")
            .Columns.Add("SCENT", GetType(System.Decimal), "IIF(EXPENSE_TYPE_CODE='SCENT',DIST_AMT,0)")
            .Columns.Add("VISUAL", GetType(System.Decimal), "IIF(EXPENSE_TYPE_CODE='VISUAL',DIST_AMT,0)")
            .Columns.Add("SLSINCENT", GetType(System.Decimal), "IIF(EXPENSE_TYPE_CODE='SLSINCENT',DIST_AMT,0)")
            .Columns.Add("TOTAL", GetType(System.Decimal), "ISNULL(COOP,0)+ISNULL(SCENT,0)+ISNULL(VISUAL,0)+ISNULL(SLSINCENT,0)")
        End With
        Fill_Records("SPTCOOPX")
    End Sub

    Sub Create_Worktable()

        sqlwCUST_CODE = " and CUST_CODE = '" & CUST_CODE & "'"

        If CUST_CODEs.Count > 1 Then
            sqlwCUST_CODE = " and CUST_CODE in ('" & Join(CUST_CODEs.ToArray, "','") & "')"
        End If

        ' If CUST_CODE = "MACYS" Then sqlw = " and CUST_CODE LIKE '" & CUST_CODE & "%'"

        'ASCMAIN1.Progress("Now Loading Data")

        rowGLTPARM3 = LookUp("GLTPARM3", RYW)
        RYM = rowGLTPARM3.Item("YYYYMM")

        Dim YTD_M As Integer = Val(Mid(RYM, 5, 2)) - 1
        If YTD_M = 0 Then YTD_M = 12

        REL_WEEK = Val(rowGLTPARM3.Item("REL_WEEK") & "")
        Dim LEGEND_WK As String = rowGLTPARM3.Item("LEGEND")
        WEEK_LEGEND = "Week Ending " & Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd/yy") & ",  " & LEGEND_WK

        If Mid(RYW, 5, 2) >= "27" Then
            SEASON_TYPE = "F"
        Else
            SEASON_TYPE = "S"
        End If
        SEASON_YEAR = Mid(RYW, 1, 4)
        SEASON_CODE = SEASON_YEAR & SEASON_TYPE
        SEASON_YEAR_LY = Format(Val(Mid(RYW, 1, 4)) - 1, "0000")
        SEASON_CODE_LY = SEASON_YEAR_LY & SEASON_TYPE
        SEASON_YEAR_NY = Format(Val(Mid(RYW, 1, 4)) + 1, "0000")

        If SEASON_TYPE = "F" Then
            SEASON_CODE_PREV = SEASON_YEAR & "S"
            SEASON_CODE_NEXT = SEASON_YEAR_NY & "S"
        Else
            SEASON_CODE_PREV = SEASON_YEAR_LY & "F"
            SEASON_CODE_NEXT = SEASON_YEAR & "F"
        End If

        RYW_LY = ASCMAIN1.Week_Calc(RYW, -52)
        RYW_2LY = ASCMAIN1.Week_Calc(RYW, -52 * 2)
        If ASCMAIN1.CLIENT = "INT" Then
            RYW_LY = Format(Val(Mid(RYW, 1, 4)) - 1, "0000") & Mid(RYW, 5, 2)
            RYW_2LY = Format(Val(Mid(RYW, 1, 4)) - 2, "0000") & Mid(RYW, 5, 2)
        End If
        RYM01 = IIf(Mid(RYM, 5, 2) = "01", _
                                  Format(Val(Mid(RYM, 1, 4)) - 1, "0000") & "02", _
                                  Mid(RYM, 1, 4) & "02")

        Dim WKS110(110) As String
        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYWW between '" & RYW_2LY & "' and '" & RYW & "'"
        Dim W110 As Integer = -1
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW DESC")
            W110 += 1
            WKS110(W110) = row.Item("YYYYWW")
        Next

        Dim YPs(12, 2) As String
        For iYP As Integer = 1 To 12
            YPs(iYP, 0) = ASCMAIN1.Period_Calc(RYM01, iYP - 1)
            YPs(iYP, 1) = ASCMAIN1.Period_Calc(YPs(iYP, 0), -12)
            YPs(iYP, 2) = ASCMAIN1.Period_Calc(YPs(iYP, 0), -24)
        Next

        Dim WKS(10, 2) As String
        ' 1-6 are for weeks
        ' 7 = HTD445
        ' 8 = STD454
        ' 9 = YTD445
        ' 10 = YTD454

        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYMM = '" & RYM & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
            Dim RW As Integer = Val(row.Item("REL_WEEK") & "")
            Dim YW As String = row.Item("YYYYWW")
            WKS(RW, 0) = YW
            WKS(RW, 1) = ASCMAIN1.Week_Calc(YW, -52)
            WKS(RW, 2) = ASCMAIN1.Week_Calc(YW, -52 * 2)
            If ASCMAIN1.CLIENT = "INT" Then
                WKS(RW, 1) = Format(Val(Mid(YW, 1, 4)) - 1, "0000") & Mid(YW, 5, 2)
                WKS(RW, 2) = Format(Val(Mid(YW, 1, 4)) - 2, "0000") & Mid(YW, 5, 2)
            End If
        Next

        If Mid(RYM, 5, 2) >= "07" Then ' SHOULDN'T THIS SAY OR MID(RYM,5,2) = "01"
            WKS(7, 0) = Mid(RYW, 1, 4) & "23"  'HTD445
        Else
            WKS(7, 0) = ASCMAIN1.Week_Calc(Mid(RYM, 1, 4) & "01", -4) 'HTD445
        End If
        If Mid(RYW, 5, 2) >= "27" Then
            WKS(8, 0) = Mid(RYW, 1, 4) & "27" 'STD454
        Else
            WKS(8, 0) = Mid(RYW, 1, 4) & "01" 'STD454
        End If
        If Mid(RYM, 5, 2) = "01" Then
            WKS(9, 0) = ASCMAIN1.Week_Calc(RYW, 1 - REL_WEEK)
        Else
            WKS(9, 0) = ASCMAIN1.Week_Calc(Mid(RYM, 1, 4) & "01", -4) 'YTD445
        End If

        WKS(10, 0) = Mid(RYW, 1, 4) & "01" 'YTD454

        For W As Integer = 7 To 10
            WKS(W, 1) = ASCMAIN1.Week_Calc(WKS(W, 0), -52)
            WKS(W, 2) = ASCMAIN1.Week_Calc(WKS(W, 0), -52 * 2)
            If ASCMAIN1.CLIENT = "INT" Then
                WKS(W, 1) = Format(Val(Mid(WKS(W, 0), 1, 4)) - 1, "0000") & Mid(WKS(W, 0), 5, 2)
                WKS(W, 2) = Format(Val(Mid(WKS(W, 0), 1, 4)) - 2, "0000") & Mid(WKS(W, 0), 5, 2)
            End If
        Next

        Dim XX As Integer = 10
        Dim D As String = "AMT_SOLD / 1000"

        Dim M12TYLY As String = ""
        For Each XY As String In New String() {"TY", "LY", "LY2"}
            Dim Y As Integer = IIf(XY = "TY", 0, IIf(XY = "LY", 1, 2))
            For M As Integer = 1 To 12
                M12TYLY &= ", Sum (Decode (OPS_YYYYPP,'" & YPs(M, Y) & "'," & D & ",0)) " & XY & "_M" & Format(M, "00") & vbCrLf
            Next
            M12TYLY &= ", Sum (Case when OPS_YYYYPP between '" & YPs(1, Y) & "' and '" & YPs(YTD_M, Y) & "' then " & D & " else 0 End) " & XY & "_YTD" & vbCrLf
        Next

        ASCMAIN1.sql = "" _
            & "Select '" & CUST_CODE & "' CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(1, 0) & "'," & D & ",0)) TY_WK1" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(1, 1) & "'," & D & ",0)) LY_WK1" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(2, 0) & "'," & D & ",0)) TY_WK2" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(2, 1) & "'," & D & ",0)) LY_WK2" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(3, 0) & "'," & D & ",0)) TY_WK3" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(3, 1) & "'," & D & ",0)) LY_WK3" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(4, 0) & "'," & D & ",0)) TY_WK4" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(4, 1) & "'," & D & ",0)) LY_WK4" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(5, 0) & "'," & D & ",0)) TY_WK5" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(5, 1) & "'," & D & ",0)) LY_WK5" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(6, 0) & "'," & D & ",0)) TY_WK6" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(6, 1) & "'," & D & ",0)) LY_WK6" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(1, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_MTD" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(1, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_MTD" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(1, 2) & "' and '" & WKS(REL_WEEK, 2) & "' THEN " & D & " ELSE 0 END) LY2_MTD" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(7, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_HTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(7, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_HTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(7, 2) & "' and '" & WKS(REL_WEEK, 2) & "' THEN " & D & " ELSE 0 END) LY2_HTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(8, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_STD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(8, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_STD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(8, 2) & "' and '" & WKS(REL_WEEK, 2) & "' THEN " & D & " ELSE 0 END) LY2_STD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(9, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_YTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(9, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_YTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(9, 2) & "' and '" & WKS(REL_WEEK, 2) & "' THEN " & D & " ELSE 0 END) LY2_YTD445" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(10, 0) & "' and '" & WKS(REL_WEEK, 0) & "' THEN " & D & " ELSE 0 END) TY_YTD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(10, 1) & "' and '" & WKS(REL_WEEK, 1) & "' THEN " & D & " ELSE 0 END) LY_YTD454" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS(10, 2) & "' and '" & WKS(REL_WEEK, 2) & "' THEN " & D & " ELSE 0 END) LY2_YTD454" & vbCrLf _
            & M12TYLY & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS110(XX - 1) & "' and '" & WKS110(0) & "' THEN " & D & " ELSE 0 END) TY_LXX" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS110(XX + 52 - 1) & "' and '" & WKS110(1 + 52 - 1) & "' THEN " & D & " ELSE 0 END) LY_LXX" & vbCrLf _
            & ", Sum (Case when OPS_YYYYWW between '" & WKS110(XX + 52 - 10 - 1) & "' and '" & WKS110(1 + 52 - 10 - 1) & "' THEN " & D & " ELSE 0 END) LY_NXX" & vbCrLf _
            & ", Sum (Decode(OPS_YYYYWW,'" & WKS(REL_WEEK, 0) & "',QTY_EOW,0)) TYTW_ONH" & vbCrLf _
            & " from RSTRETL1" & vbCrLf _
            & " where OPS_YYYYWW between '" & WKS(10, 2) & "' and '" & WKS(REL_WEEK, 0) & "'" & vbCrLf _
            & sqlwCUST_CODE _
            & " group by CUST_STORE_NO, ITEM_CODE"
        ' & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE"
        Dim sqlRSTBREV1 As String = ASCMAIN1.sql
        RSTBREV1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTBREV1 & " Modify CUST_CODE VARCHAR2(10)")
        ASCMAIN1.sql = "Alter Table " & RSTBREV1 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = Replace(sqlRSTBREV1, D, "QTY_SOLD")
        RSTSPRO5 = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "Alter Table " & RSTSPRO5 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)"
        ASCDATA1.ExecuteSQL()
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO5 & " Modify CUST_CODE VARCHAR2(10)")

        ASCMAIN1.sql = Replace(Replace(sqlRSTBREV1, D, "QTY_SOLD"), "CUST_STORE_NO, ITEM_CODE", "ITEM_CODE")
        RSTSPRO7 = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "Alter Table " & RSTSPRO7 & " Add Primary Key (CUST_CODE, ITEM_CODE)"
        ASCDATA1.ExecuteSQL()
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO7 & " Modify CUST_CODE VARCHAR2(10)")

        ASCMAIN1.sql = Replace(Replace(Replace(Replace(sqlRSTBREV1, D, "QTY_SOLD"), "CUST_STORE_NO, ITEM_CODE", "ITEM_CODE"), _
                                   "Select '" & CUST_CODE & "' CUST_CODE", "Select CUST_CODE"), _
                                   "group by ITEM_CODE", "group by CUST_CODE, ITEM_CODE")
        RSTSPRO7C = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "Alter Table " & RSTSPRO7C & " Add Primary Key (CUST_CODE, ITEM_CODE)"
        ASCDATA1.ExecuteSQL()
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO7C & " Modify CUST_CODE VARCHAR2(10)")

        Dim sqlCs As String = "Select Distinct HC_CODE from SATAUTH1 " & vbCrLf _
            & " where OPS_YYYYPP_OPENED IS NOT NULL" & vbCrLf _
            & "   and OPS_YYYYPP_CLOSED IS NULL" & vbCrLf _
            & sqlwCUST_CODE

        Dim sqlD As String = "Delete from " & RSTBREV1 & " where ITEM_CODE in (" & vbCrLf _
            & " Select X.ITEM_CODE from " & RSTBREV1 & " X, ICTITEM1, ICTCOLL1, ICTBRAN1" & vbCrLf _
            & "  where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "    and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "    and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
            & "    and (NVL(ICTBRAN1.BRAND_STATUS,'?') <> 'A' or ICTCOLL1.HC_CODE Not in (" & sqlCs & ")))"
        ASCDATA1.ExecuteSQL(sqlD)
        ASCDATA1.ExecuteSQL(Replace(sqlD, RSTBREV1, RSTSPRO5))
        ASCDATA1.ExecuteSQL(Replace(sqlD, RSTBREV1, RSTSPRO7))

        Dim sqlSum As String = "" _
            & ", Sum(X.TY_WK1) TY_WK1, Sum(X.TY_WK2) TY_WK2, Sum(X.TY_WK3) TY_WK3, Sum(X.TY_WK4) TY_WK4, Sum(X.TY_WK5) TY_WK5, Sum(X.TY_WK6) TY_WK6" & vbCrLf _
            & ", Sum(X.LY_WK1) LY_WK1, Sum(X.LY_WK2) LY_WK2, Sum(X.LY_WK3) LY_WK3, Sum(X.LY_WK4) LY_WK4, Sum(X.LY_WK5) LY_WK5, Sum(X.LY_WK6) LY_WK6" & vbCrLf _
            & ", Sum(X.TY_MTD) TY_MTD, Sum(X.TY_HTD445) TY_HTD445, Sum(X.TY_STD454) TY_STD454, Sum(X.TY_YTD445) TY_YTD445, Sum(X.TY_YTD454) TY_YTD454" & vbCrLf _
            & ", Sum(X.LY_MTD) LY_MTD, Sum(X.LY_HTD445) LY_HTD445, Sum(X.LY_STD454) LY_STD454, Sum(X.LY_YTD445) LY_YTD445, Sum(X.LY_YTD454) LY_YTD454" & vbCrLf _
            & ", Sum(X.LY2_MTD) LY2_MTD, Sum(X.LY2_HTD445) LY2_HTD445, Sum(X.LY2_STD454) LY2_STD454, Sum(X.LY2_YTD445) LY2_YTD445, Sum(X.LY2_YTD454) LY2_YTD454" & vbCrLf _
            & ", Sum(X.TY_M01) TY_M01, Sum(X.TY_M02) TY_M02, Sum(X.TY_M03) TY_M03, Sum(X.TY_M04) TY_M04, Sum(X.TY_M05) TY_M05, Sum(X.TY_M06) TY_M06" & vbCrLf _
            & ", Sum(X.TY_M07) TY_M07, Sum(X.TY_M08) TY_M08, Sum(X.TY_M09) TY_M09, Sum(X.TY_M10) TY_M10, Sum(X.TY_M11) TY_M11, Sum(X.TY_M12) TY_M12" & vbCrLf _
            & ", Sum(X.TY_YTD) TY_YTD" & vbCrLf _
            & ", Sum(X.LY_M01) LY_M01, Sum(X.LY_M02) LY_M02, Sum(X.LY_M03) LY_M03, Sum(X.LY_M04) LY_M04, Sum(X.LY_M05) LY_M05, Sum(X.LY_M06) LY_M06" & vbCrLf _
            & ", Sum(X.LY_M07) LY_M07, Sum(X.LY_M08) LY_M08, Sum(X.LY_M09) LY_M09, Sum(X.LY_M10) LY_M10, Sum(X.LY_M11) LY_M11, Sum(X.LY_M12) LY_M12" & vbCrLf _
            & ", Sum(X.LY_YTD) LY_YTD" & vbCrLf _
            & ", Sum(X.LY2_M01) LY2_M01, Sum(X.LY2_M02) LY2_M02, Sum(X.LY2_M03) LY2_M03, Sum(X.LY2_M04) LY2_M04, Sum(X.LY2_M05) LY2_M05, Sum(X.LY2_M06) LY2_M06" & vbCrLf _
            & ", Sum(X.LY2_M07) LY2_M07, Sum(X.LY2_M08) LY2_M08, Sum(X.LY2_M09) LY2_M09, Sum(X.LY2_M10) LY2_M10, Sum(X.LY2_M11) LY2_M11, Sum(X.LY2_M12) LY2_M12" & vbCrLf _
            & ", Sum(X.LY2_YTD) LY2_YTD" & vbCrLf _
            & ", Sum(X.TY_LXX) TY_LXX, Sum(X.LY_LXX) LY_LXX, Sum(X.LY_NXX) LY_NXX" & vbCrLf _
            & ", Sum(X.TYTW_ONH * ICTITEM1.ITEM_RETAIL_PRICE / 1000) TYTW_ONH" & vbCrLf


        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE" & vbCrLf _
            & sqlSum _
            & " from " & RSTBREV1 & " X" & vbCrLf _
            & ",ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE"
        RSTSPRO2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO2 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)")


        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE" & vbCrLf _
            & Replace(sqlSum, " * ICTITEM1.ITEM_RETAIL_PRICE / 1000", "") _
            & " from " & RSTSPRO5 & " X" & vbCrLf _
            & ",ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by X.CUST_CODE" & vbCrLf _
            & ", X.CUST_STORE_NO" & vbCrLf _
            & ", X.ITEM_CODE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTCOLL1.COLLECTION_GENDER" & vbCrLf _
            & ", ICTCOLL1.HC_CODE"
        RSTSPRO6 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)")

        ASCMAIN1.sql = "Select " & vbCrLf _
            & "  CUST_CODE" & vbCrLf _
            & ", CUST_STORE_NO" & vbCrLf _
            & ", COLLECTION_GENDER" & vbCrLf _
            & ", BRAND_CODE" & vbCrLf _
            & ", HC_CODE" & vbCrLf _
            & Replace(sqlSum, " * ICTITEM1.ITEM_RETAIL_PRICE / 1000", "") _
            & " from " & RSTSPRO2 & " X" & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, COLLECTION_GENDER, BRAND_CODE, HC_CODE"
        RSTSPRO3 = ASCMAIN1.Temp_Table

        ASCMAIN1.sql = "" _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", MAX (INV_DATE) INV_DATE" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & ", 0 ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", NULL INV_DATE" & vbCrLf _
            & " from SOTORDR2" & vbCrLf _
            & " where ORDR_STATUS between 'O' and 'P'" & vbCrLf _
            & "   and CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _
            & " group by CUST_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ", SUM (SOTINVH2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", 0 ORDR_QTY_PICK" & vbCrLf _
            & ", 0 ORDR_QTY_OPEN" & vbCrLf _
            & ", MAX (SOTINVH1.INV_DATE) INV_DATE" & vbCrLf _
            & " from SOTINVH2,SOTINVH1" & vbCrLf _
            & " where SOTINVH2.OPS_YYYYWW > '" & ASCMAIN1.Week_Calc(RYW, -3) & "' and SOTINVH2.OPS_YYYYWW <= '" & RYW & "'" & vbCrLf _
            & "   and SOTINVH2.INV_TYPE = 'I'" & vbCrLf _
            & Replace(sqlwCUST_CODE, "and CUST_CODE", "and SOTINVH2.CUST_CODE") & vbCrLf _
            & "   and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & " group by SOTINVH2.CUST_CODE, SOTINVH2.CUST_STORE_NO, SOTINVH2.ITEM_CODE" & vbCrLf _
            & ")  group by CUST_CODE, CUST_STORE_NO, ITEM_CODE"
        RSTSPRO4 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO4 & " Add Primary Key (CUST_CODE, CUST_STORE_NO, ITEM_CODE)")
        '            & "   and SOTINVH2.CUST_CODE = '" & CUST_CODE & "'" & vbCrLf _


        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add ORDR_QTY_SHIP NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add ORDR_QTY_PICK NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & RSTSPRO6 & " Add ORDR_QTY_OPEN NUMBER (8,0)")

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is Select * from " & RSTSPRO4 & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & RSTSPRO6 & vbCrLf _
            & "    Set ORDR_QTY_SHIP = R1.ORDR_QTY_SHIP" & vbCrLf _
            & "       ,ORDR_QTY_PICK = R1.ORDR_QTY_PICK" & vbCrLf _
            & "       ,ORDR_QTY_OPEN = R1.ORDR_QTY_OPEN" & vbCrLf _
            & "    where CUST_CODE = R1.CUST_CODE" & vbCrLf _
            & "      and CUST_STORE_NO = R1.CUST_STORE_NO" & vbCrLf _
            & "      and ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ' 02/24/16 sp wants current season and last season
        'If retail week is Jan’16 thru Jun’16, then show allocations whose End date is between Jul’15 and Jun’16 (ie, Fall15 and Spring16)
        'If retail week is Jul’16 thru Dec’16, then show allocations whose End date is between Jan’16 and Dec’16 (ie, Spring16 and Fall16)

        Dim YM_of_YW As String = LookUp("GLTPARM3", RYW).Item("YYYYMM")
        Dim DTE_START As Date
        Dim DTE_END As Date
        Dim YMY As String = Mid(YM_of_YW, 1, 4)
        If Mid(YM_of_YW, 5, 2) >= "07" And Mid(YM_of_YW, 5, 2) <= "12" Then
            DTE_START = CDate("01/01/" & YMY)
            DTE_END = CDate("12/31/" & YMY)
        Else
            'DTE_START = CDate("07/01/" & Format(Val(YMY) - 1, "0000"))
            'DTE_END = "06/30/" & YMY
            'DTE_START = CDate("07/01/" & Format(Val(YMY) + 0, "0000"))
            'DTE_END = CDate("06/30/" & Format(Val(YMY) + 1, "0000"))
            ' 02/24/16 sp wants current season and last season
            DTE_START = CDate("07/01/" & Format(Val(YMY) - 1, "0000"))
            DTE_END = CDate("06/30/" & Format(Val(YMY) + 0, "0000"))
        End If
        ' "   and SOTALLO2.CUST_CODE = '" & CUST_CODE & "'"
        'ASCMAIN1.sql = "Select Distinct SOTALLO2.ALLO_CTL_NO, SOTALLO1. ITEM_CODE" & vbCrLf _
        '    & " from SOTALLO2,SOTALLO1" & vbCrLf _
        '    & " where SOTALLO1.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
        '    & Replace(sqlwCUST_CODE, "and CUST_CODE", "and SOTALLO2.CUST_CODE") & vbCrLf _
        '    & "   and SOTALLO1.DATE_START >= '" & Format(DTE_START, "dd-MMM-yyyy") & "'" & vbCrLf _
        '    & "   and SOTALLO1.DATE_START <= '" & Format(DTE_END, "dd-MMM-yyyy") & "'"
        ASCMAIN1.sql = "Select Distinct SOTALLO2.ALLO_CTL_NO, SOTALLO1. ITEM_CODE" & vbCrLf _
            & " from SOTALLO2,SOTALLO1" & vbCrLf _
            & " where SOTALLO1.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
            & Replace(sqlwCUST_CODE, "and CUST_CODE", "and SOTALLO2.CUST_CODE") & vbCrLf _
            & "   and SOTALLO1.DATE_END >= '" & Format(DTE_START, "dd-MMM-yyyy") & "'" & vbCrLf _
            & "   and SOTALLO1.DATE_END <= '" & Format(DTE_END, "dd-MMM-yyyy") & "'"

        SOTALLOX = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "Alter Table " & SOTALLOX & " Add Primary Key (ALLO_CTL_NO)"
        ASCDATA1.ExecuteSQL()


        ASCMAIN1.sql = "Select SPTCOOP1.*, SPTCOOP3.COLLECTION_CODE, SPTCOOP3.DIST_AMT, SPTCOOP3.AUTH_LNO, SPTCOOP3.FEATURE_DESC" & vbCrLf _
           & ", ICTCOLL1.COLLECTION_NAME, ICTCOLL1.COLLECTION_GENDER, ICTCOLL1.BRAND_CODE, ICTBRAN1.BRAND_NAME, ICTCOLL1.HC_CODE, ICTCOLL0.HC_NAME" & vbCrLf _
           & " from SPTCOOP1,SPTCOOP3,ICTCOLL1,ICTBRAN1,ICTCOLL0" & vbCrLf _
           & " where ICTCOLL1.COLLECTION_CODE = SPTCOOP3.COLLECTION_CODE" & vbCrLf _
           & "   and SPTCOOP1.APPR_STATUS_CODE in ('A','P','G')" & vbCrLf _
           & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE" & vbCrLf _
           & "   and ICTCOLL0.HC_CODE = ICTCOLL1.HC_CODE" & vbCrLf _
           & "   and SPTCOOP3.AUTH_NO = SPTCOOP1.AUTH_NO" & vbCrLf _
           & Replace(sqlwCUST_CODE, "and CUST_CODE", "and SPTCOOP1.CUST_CODE") & vbCrLf _
           & "   and SPTCOOP1.SEASON_CODE in ('" & SEASON_CODE_LY & "','" & SEASON_CODE_PREV & "','" & SEASON_CODE & "')" & vbCrLf _
           & "   and SPTCOOP1.EXPENSE_TYPE_CODE in ('COOP','SCENT','VISUAL','SLSINCENT')"
        SPTCOOPX = ASCMAIN1.Temp_Table
        ' "   and SPTCOOP1.CUST_CODE = '" & CUST_CODE & "'"

        Dim TYM01 As String = LookUp("GLTPARM3", WKS(8, 0)).Item("YYYYPP")
        Dim TYMXX As String = LookUp("GLTPARM3", WKS(REL_WEEK, 0)).Item("YYYYPP")
        Dim LYM01 As String = LookUp("GLTPARM3", WKS(8, 1)).Item("YYYYPP")
        Dim TYWXX As String = WKS(REL_WEEK, 0)
        'Dim LYWXX As String = ""
        'If Mid(TYWXX, 5, 2) = "53" Then
        '    LYWXX = "0000"
        'Else
        '    LYWXX = Format(Val(Mid(TYWXX, 1, 4)) - 1, "0000") & Mid(TYWXX, 5, 2)
        'End If

        Dim LYMXX As String = Format(Val(Mid(TYMXX, 1, 4)) - 1, "0000") & Mid(TYMXX, 5, 2)
        'If Mid(TYWXX, 5, 2) = "53" Then
        '    LYMXX = "0000"
        'Else
        '    LYMXX = LookUp("GLTPARM3", WKS(REL_WEEK, 1)).Item("YYYYPP")
        'End If


        ASCMAIN1.sql = "" _
                    & "Select '" & CUST_CODE & "' CUST_CODE, HC_CODE" & vbCrLf _
                    & ", SUM (CASE WHEN OPS_YYYYPP_OPENED <= '" & LYM01 & "'" & vbCrLf _
                    & "AND NVL(OPS_YYYYPP_CLOSED,'" & LYMXX & "') >= '" & LYMXX & "' THEN 1 ELSE 0 END) OPEN_LY" & vbCrLf _
                    & ", SUM (CASE WHEN OPS_YYYYPP_OPENED <= '" & TYM01 & "' " & vbCrLf _
                    & "AND NVL(OPS_YYYYPP_CLOSED,'" & TYMXX & "') >= '" & TYMXX & "' THEN 1 ELSE 0 END) OPEN_TY" & vbCrLf _
                    & " from SATAUTH1" & vbCrLf _
                    & Replace(sqlwCUST_CODE, "and CUST_CODE", "where SATAUTH1.CUST_CODE") & vbCrLf _
                    & " group by HC_CODE" ' & " group by CUST_CODE, HC_CODE"
        ' " where CUST_CODE = '" & CUST_CODE & "'" 
        SATAUTHX = ASCMAIN1.Temp_Table
    End Sub
#End Region
End Class