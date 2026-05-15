Public Class POFVSUM1
    Dim POTVSUM1 As String
    Dim POTVSUM2 As String
    Dim sqlPOTVSUM1 As String
    Dim sqlPOTVSUM1_sum As String
    Dim sqlPOTVSUM2 As String
    Dim sqlSOTINVHX As String
    Dim sqlPOTVSUM1_STORES As String
    Dim STORES As List(Of String)
    Dim STORES_XX As String
    Dim PERIODS_XX As String

    Dim RYP0 As String
    Dim RYP1 As String
    Dim RYW0 As String
    Dim RYW1 As String
    Dim Periods As Integer
    Dim VEND_CODE As String
    Dim Stores_Max As Integer = 120

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)

        Set_cmbYW("RYW0", ASCMAIN1.CYW, -3 * 52, 0, -13)
        Set_cmbYW("RYW1", ASCMAIN1.CYW, -3 * 52, 0, 0)

        With dst
            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "POTVSUM1", "**", 0, False)
            With .Tables("POTVSUM1")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE, '0' YEAR, 'XX' DATA_TYPE from ICTITEM1"
            Create_TDA(.Tables.Add, "POTVSUM1_DTL", "**", 0, False)
            With .Tables("POTVSUM1_DTL")
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            .Relations.Add("POTVSUM1_POTVSUM1_DTL" _
                           , New DataColumn() {.Tables("POTVSUM1").Columns("CODE_VALUE")} _
                           , New DataColumn() {.Tables("POTVSUM1_DTL").Columns("CODE_VALUE")})

            ASCMAIN1.sql = "Select ITEM_CODE CODE_VALUE_PARENT, ITEM_CODE CODE_VALUE, ITEM_DESC DESC_VALUE from ICTITEM1"
            Create_TDA(.Tables.Add, "POTVSUM2", "**", 0, False)
            With .Tables("POTVSUM2")
                .Columns.Add("SUB_CODE_VALUE1")
                .Columns.Add("SUB_CODE_VALUE2")
                .Columns.Add("SUB_CODE_VALUE3")
                .Columns.Add("SUB_CODE_VALUE4")
                .Columns.Add("SUB_CODE_VALUE5")
                .Columns.Add("RTL_PRICE", GetType(System.Decimal))
                .Columns.Add("WSL_PRICE", GetType(System.Decimal))
                For P As Integer = 0 To Stores_Max
                    .Columns.Add("P" & Format(P, "00"), GetType(System.Int32))
                Next
                .Columns.Add("PXX", GetType(System.Int32))
            End With

            .Relations.Add("POTVSUM1_POTVSUM2" _
           , New DataColumn() {.Tables("POTVSUM1").Columns("CODE_VALUE")} _
           , New DataColumn() {.Tables("POTVSUM2").Columns("CODE_VALUE_PARENT")})

            ASCMAIN1.sql = "Select SOTINVH2.INV_TYPE, SOTINVH2.INV_NO" _
            & ", SOTINVH1.INV_DATE, SOTINVH1.ORDR_TYPE_CODE, SOTINVH1.ORDR_CUST_PO" _
            & ", SOTINVH2.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH2.OPS_YYYYWW" _
            & ", SOTINVH2.VEND_CODE, SOTINVH2.CUST_STORE_NO" _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') CUST_STORE_LOCATION" _
            & ", SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
            & ", SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE ORDR_AMT_SHIP" _
            & " from SOTINVH2,ICTITEM1,ARTCUST2,SOTINVH1 " _
            & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE " _
            & " and ARTCUST2.VEND_CODE (+) = SOTINVH2.VEND_CODE " _
            & " and ARTCUST2.CUST_STORE_NO (+) = SOTINVH2.CUST_STORE_NO " _
            & " and SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE " _
            & " and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            sqlSOTINVHX = ASCMAIN1.sql
            Create_TDA(.Tables.Add, "SOTINVHX", "**", 0, False)
        End With

        Fill_Records("TATSTATE")

        grdPOTVSUM1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        grdPOTVSUM1.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdPOTVSUM1.DisplayLayout.MaxBandDepth = 1
        grdPOTVSUM1.DataSource = dst.Tables("POTVSUM1")

        grdPOTVSUM2.DataSource = dst.Tables("POTVSUM2")

        grdSOTINVHX.DataSource = dst.Tables("SOTINVHX")

        Create_Summary(grdSOTINVHX, "INV_NO", "Count")
        Create_Summary(grdSOTINVHX, "ORDR_QTY_SHIP")
        Create_Summary(grdSOTINVHX, "ORDR_AMT_SHIP")

        Create_Summary(grdPOTVSUM1, "CODE_VALUE", "Count")
        For P As Integer = 0 To Stores_Max
            Create_Summary(grdPOTVSUM1, "P" & Format(P, "00"))
        Next
        Create_Summary(grdPOTVSUM1, "PXX")

        Create_Summary(grdPOTVSUM2, "CODE_VALUE", "Count")
        For P As Integer = 0 To Stores_Max
            Create_Summary(grdPOTVSUM2, "P" & Format(P, "00"))
        Next
        Create_Summary(grdPOTVSUM2, "PXX")

        With grdPOTVSUM1.DisplayLayout.Bands("POTVSUM1")
            .Columns("CODE_VALUE").Header.Fixed = True
            .Columns("DESC_VALUE").Header.Fixed = True
            .Columns("SUB_CODE_VALUE1").Header.Fixed = True
            .Columns("SUB_CODE_VALUE2").Header.Fixed = True
            .Columns("SUB_CODE_VALUE3").Header.Fixed = True
            .Columns("SUB_CODE_VALUE4").Header.Fixed = True
            .Columns("SUB_CODE_VALUE5").Header.Fixed = True
            .Columns("RTL_PRICE").Header.Fixed = True
            .Columns("WSL_PRICE").Header.Fixed = True
            .Columns("P00").Header.Fixed = True
            .Columns("PXX").Header.Fixed = True
        End With

        With grdPOTVSUM2.DisplayLayout.Bands("POTVSUM2")
            .Columns("CODE_VALUE").Header.Fixed = True
            .Columns("DESC_VALUE").Header.Fixed = True
            .Columns("SUB_CODE_VALUE1").Header.Fixed = True
            .Columns("SUB_CODE_VALUE2").Header.Fixed = True
            .Columns("SUB_CODE_VALUE3").Header.Fixed = True
            .Columns("SUB_CODE_VALUE4").Header.Fixed = True
            .Columns("SUB_CODE_VALUE5").Header.Fixed = True
            .Columns("RTL_PRICE").Header.Fixed = True
            .Columns("WSL_PRICE").Header.Fixed = True
            .Columns("P00").Header.Fixed = True
            .Columns("PXX").Header.Fixed = True
        End With

        optSI.Tag = "*"

        chkExtendedData.Visible = False ' CLICK THIS OPTION, SELECT VONMAUR, THEN LOAD, THEN GET AN ERROR

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("VEND_CODE")

                If EMsg = "" Then

                    If Trim(ASCMAIN1.USER_CODES) = "FS" Then
                        If cdr.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                And Not TAC.TACMAIN1.SREP_CODEs.Contains(cdr.Item("SREP_CODE") & "") Then


                            Dim found_store As Boolean = False
                            ASCMAIN1.sql = "Select Distinct SREP_CODE from ARTCUST2 where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
                            ASCMAIN1.sql &= " UNION "
                            ASCMAIN1.sql &= "Select Distinct SELL_CODE from ARTCUST2 where VEND_CODE = '" & Absx1.txtFor("VEND_CODE").Text & "'"
                            For Each rowARTCUST2_SREP As DataRow In ASCDATA1.GetDataTable.Select("")
                                If rowARTCUST2_SREP.Item("SREP_CODE") & "" <> TAC.TACMAIN1.SREP_CODE _
                                    And Not TAC.TACMAIN1.SREP_CODEs.Contains(rowARTCUST2_SREP.Item("SREP_CODE") & "") Then
                                Else
                                    found_store = True
                                End If
                            Next

                            If Not found_store Then
                                EMsg &= vbCr & "Customer " & Absx1.txtFor("VEND_CODE").Text & " is not connected to Sales Rep code " & TAC.TACMAIN1.SREP_CODE
                            End If
                        End If
                    End If

                    Validate_Range(EMsg)


                End If

            Case "Load Summary"

                Validate_Range(EMsg)

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
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print Report"
                Print_Report()

            Case "Load Summary"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Loading Sales Summary by Customer")

                Create_POTVSUM1()
                optXP.Items(0).DisplayText = "Period"
                Load_Data()
                grdPOTVSUM1.Parent = grpSummary

                UltraExplorerBar1.Groups("Data Options").Visible = True
                optSI.Visible = False
                optXP.Visible = False
                chkShowDetails.Visible = False

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Load Summary").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Load Summary").Visible = Not (Trim(ASCMAIN1.USER_CODES) = "FS")
                .Groups("Data Options").Visible = tf
                .Groups("Options").Visible = Not tf
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)
        SplitContainer1.Visible = tf
        grpSummary.Visible = Not tf

        If ScreenMode Then
            grdPOTVSUM1.Parent = SplitContainer1.Panel1
            optSI.Visible = True
            optXP.Visible = True
            chkShowDetails.Visible = True
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() {"POTVSUM1", "POTVSUM1_DTL", "POTVSUM2", "EDT852T1", "SOTINVHX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        'Absx1.txtFor("VEND_CODE").Text = ""
        tabDetails.SelectedTab = tabDetails.Tabs("Details")
    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Customer Sales Data")
        Save_Header_Fields(UltraGroupBox1)
        VEND_CODE = HFs("VEND_CODE")
        Create_POTVSUM1()
        optXP.Items(0).DisplayText = "Period"
        Load_Data()
        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control, _
         ByVal COLUMN_NAME As String, _
         Optional ByRef sql_where As String = "", _
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "VEND_CODE"

        End Select
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOTVSUM1, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE")
        Load_Popup_Menu(grdPOTVSUM2, "SSSSSSSSS", "Show Filter", "Show GroupBox", "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE")
        Load_Popup_Menu(grdSOTINVHX, "SS", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show SUB_CODE_VALUE1") Then
            For I As Integer = 1 To 5
                Dim COLUMN_NAME As String = "SUB_CODE_VALUE" & CStr(I)
                tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
                tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            Next

            COLUMN_NAME = "RTL_PRICE"
            tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = (EntryMode = "E")
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
            COLUMN_NAME = "WSL_PRICE"
            tlb_sbt = DirectCast(tlb_pop.Tools("Show " & COLUMN_NAME), UltraWinToolbars.StateButtonTool)
            tlb_sbt.SharedProps.Visible = (EntryMode = "E")
            tlb_sbt.Checked = Not grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden
            tlb_sbt.SharedProps.Caption = "Show " & grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Header.Caption
        End If


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdPOTVSUM1"

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Show SUB_CODE_VALUE1", "Show SUB_CODE_VALUE2", "Show SUB_CODE_VALUE3", "Show SUB_CODE_VALUE4", "Show SUB_CODE_VALUE5", "Show RTL_PRICE", "Show WSL_PRICE"

                Dim COLUMN_NAME As String = Mid(e.Tool.Key, 6)
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = Not tlb_sbt.Checked
                If grdPOTVSUM1.DisplayLayout.Bands.Count > 1 Then
                    With grdPOTVSUM1.DisplayLayout.Bands(1).Columns("DATA_TYPE")
                        If tlb_sbt.Checked Then
                            .ColSpan += 1
                        Else
                            .ColSpan -= 1
                        End If
                    End With
                End If
        End Select
    End Sub

    Public Overrides Sub tlb_ToolValueChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolEventArgs)
        MyBase.tlb_ToolValueChanged(sender, e)

        Select Case e.Tool.Key
        End Select

    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            'Case "DTE0", "DTE1"
            '    If e.KeyCode = Windows.Forms.Keys.Enter And EntryMode = "" Then
            '        Call Click_Command("Load", e)
            '    End If
        End Select

    End Sub

#End Region

    Sub Create_POTVSUM1()

        If POTVSUM2 = "" Then
            POTVSUM1 = ASCMAIN1.Temp_Table("Select VEND_CODE from ARTCUST1 where ROWNUM < 1")
            POTVSUM2 = ASCMAIN1.Temp_Table("Select VEND_CODE from ARTCUST1 where ROWNUM < 1")
        End If
        ASCDATA1.ExecuteSQL("Drop Table " & POTVSUM1)
        ASCDATA1.ExecuteSQL("Drop Table " & POTVSUM2)

        Dim PX As String = "OPS_YYYYPP"
        Dim PX0 As String = RYP0
        Dim PX1 As String = RYP1

        sqlPOTVSUM1 = ""
        sqlPOTVSUM1_sum = ""
        sqlPOTVSUM2 = ""
        Dim SQL() As String = New String() {"", ""}
        Dim P As Integer
        PERIODS_XX = ""

        'SQL = ""
        For P = 1 To Periods
            Dim PXP As String = ASCMAIN1.Period_Calc(RYP0, P - 1)
            SQL(0) &= ", Sum (Decode(" & PX & ",'" & PXP & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "00") & vbCrLf
            SQL(1) &= ", Sum (Decode(" & PX & ",'" & ASCMAIN1.Period_Calc(PXP, -12) & "',RSTRETL1.QTY_SOLD,0)) P" & Format(P, "00") & vbCrLf
            sqlPOTVSUM1 &= ", P" & Format(P, "00")
            sqlPOTVSUM1_sum &= ", Sum (P" & Format(P, "00") & ") P" & Format(P, "00")
            sqlPOTVSUM2 &= ", SUM (P" & Format(P, "00") & ") P" & Format(P, "00")
            PERIODS_XX &= "+P" & Format(P, "00")
        Next

        Dim YEAR_max As Int32 = 0
        If EntryMode = "E" Then
            If chkPriorYear.Checked Then
                YEAR_max = 1
            End If
        End If

        For YEAR As Int32 = 0 To YEAR_max

            PX = "OPS_YYYYPP"

            PX0 = RYP0
            PX1 = RYP1
            If YEAR = 1 Then
                PX0 = ASCMAIN1.Period_Calc(PX0, -12)
                PX1 = ASCMAIN1.Period_Calc(PX1, -12)
            End If


            Dim sqla As String = "Select 'TU' DATA_TYPE, '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", RSTRETL1.VEND_CODE, RSTRETL1.CUST_STORE_NO, RSTRETL1.ITEM_CODE" & vbCrLf _
            & SQL(YEAR) & vbCrLf _
            & " from RSTRETL1,GLTPARM3  " & vbCrLf _
            & " where GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf _
            & IIf(EntryMode = "E", " and RSTRETL1.VEND_CODE = '" & VEND_CODE & "'" & vbCrLf, "") _
            & " and RSTRETL1." & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & " group by RSTRETL1.VEND_CODE, RSTRETL1.CUST_STORE_NO, RSTRETL1.ITEM_CODE"

            Dim SQL_ORIG As String = sqla

            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & POTVSUM2 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & POTVSUM2 _
                & " Add Primary Key (DATA_TYPE, YEAR, VEND_CODE, CUST_STORE_NO, ITEM_CODE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & POTVSUM2 & " " & sqla)
            End If
            sqla = Replace(sqla, "'TU' DATA_TYPE", "'TD' DATA_TYPE")
            sqla = Replace(sqla, "QTY_SOLD", "AMT_SOLD")
            ASCDATA1.ExecuteSQL("Insert into " & POTVSUM2 & " " & sqla)

            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HU' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK")
            sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW")
            ASCDATA1.ExecuteSQL("Insert into " & POTVSUM2 & " " & sqla)

            sqla = Replace(SQL_ORIG, "'TU' DATA_TYPE", "'HD' DATA_TYPE")
            sqla = Replace(sqla, " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW", " and GLTPARM3.YYYYWW = RSTRETL1.OPS_YYYYWW" & vbCrLf & " and GLTPARM3.REL_WEEK = GLTPARM3.MAX_WEEK")
            sqla = Replace(sqla, "from RSTRETL1", "from RSTRETL1,ICTITEM1")
            sqla = Replace(sqla, "group by", " and ICTITEM1.ITEM_CODE = RSTRETL1.ITEM_CODE group by")
            sqla = Replace(sqla, "QTY_SOLD", "QTY_EOW * ICTITEM1.ITEM_RETAIL_PRICE") ' SB USING ICTRETLA
            ASCDATA1.ExecuteSQL("Insert into " & POTVSUM2 & " " & sqla)

            PX = "ORDR_YYYYPP_UPDATED"
            sqla = ""
            For P = 1 To Periods
                Dim PXP As String =  ASCMAIN1.Period_Calc(RYP0, P - 1 - 12 * YEAR)
                sqla &= ", Sum (Decode(" & PX & ",'" & PXP & "',SOTINVH2.ORDR_QTY_SHIP,0)) P" & Format(P, "00") & vbCrLf
            Next
            sqla = "Select DECODE(INV_TYPE,'I','S','C','R') || 'U' DATA_TYPE" & vbCrLf _
            & ", '" & CStr(YEAR) & "' YEAR" & vbCrLf _
            & ", VEND_CODE, CUST_STORE_NO, ITEM_CODE" & vbCrLf _
            & sqla & vbCrLf _
            & " from SOTINVH2 " & vbCrLf _
            & " where " & PX & " Between '" & PX0 & "' and '" & PX1 & "'" & vbCrLf _
            & IIf(EntryMode = "E", " and VEND_CODE = '" & VEND_CODE & "'" & vbCrLf, "") _
            & " group by DECODE(INV_TYPE,'I','S','C','R') || 'U', VEND_CODE, CUST_STORE_NO, ITEM_CODE"


            ASCDATA1.ExecuteSQL("Insert into " & POTVSUM2 & " " & sqla)
            sqla = Replace(sqla, "'U' DATA_TYPE", "'D' DATA_TYPE")
            sqla = Replace(sqla, " || 'U',", " || 'D',")
            sqla = Replace(sqla, "ORDR_QTY_SHIP", "ORDR_QTY_SHIP * ORDR_UNIT_PRICE")
            ASCDATA1.ExecuteSQL("Insert into " & POTVSUM2 & " " & sqla)


            sqla = "Select 'I' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", VEND_CODE, ITEM_CODE CODE_VALUE" _
            & sqlPOTVSUM2 & " from " & POTVSUM2 _
            & " group by DATA_TYPE, VEND_CODE, ITEM_CODE"
            If YEAR = 0 Then
                ASCDATA1.ExecuteSQL("Create Table " & POTVSUM1 & " as " & sqla)
                ASCDATA1.ExecuteSQL("Alter Table " & POTVSUM1 & " Add Primary Key (DATA_TYPE, YEAR, VEND_CODE, CODE_VALUE)")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & POTVSUM1 & " " & sqla)
            End If

            sqla = "Select 'S' SI, DATA_TYPE, '" & CStr(YEAR) & "' YEAR" _
            & ", VEND_CODE, CUST_STORE_NO CODE_VALUE" & sqlPOTVSUM2 _
            & " from " & POTVSUM2 & " group by DATA_TYPE, VEND_CODE, CUST_STORE_NO"
            ASCDATA1.ExecuteSQL("Insert into " & POTVSUM1 & " " & sqla)
        Next YEAR

        ASCDATA1.ExecuteSQL("Alter Table " & POTVSUM1 & " Add PXX NUMBER (10,0)")
        ASCDATA1.ExecuteSQL("Alter Table " & POTVSUM2 & " Add PXX NUMBER (10,0)")

        ASCMAIN1.sql = "Update " & POTVSUM1 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & POTVSUM1 & " " _
        & " where SI = X.SI and DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and VEND_CODE = X.VEND_CODE and CODE_VALUE = X.CODE_VALUE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & POTVSUM2 & " X SET PXX = (Select P" & Format(Periods, "00") & " from " & POTVSUM2 & " " _
        & " where DATA_TYPE = 'H' || SUBSTR(X.DATA_TYPE,2,1)" _
        & "   and YEAR = X.YEAR and VEND_CODE = X.VEND_CODE and CUST_STORE_NO = X.CUST_STORE_NO AND ITEM_CODE = X.ITEM_CODE)" _
        & " where DATA_TYPE IN ('TU','TD')"
        ASCDATA1.ExecuteSQL()

        STORES = New List(Of String)
        STORES_XX = ""
        Dim SQLX As String = ""

        If EntryMode = "E" Then
            SQLX = "Select Distinct CUST_STORE_NO from " & POTVSUM2 & " order by CUST_STORE_NO"
            For Each row As DataRow In ASCDATA1.GetDataTable(SQLX).Rows
                Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO")
                STORES.Add(CUST_STORE_NO)
            Next

        End If

        SQLX = ""
        If EntryMode = "E" Then
            For S As Integer = 1 To STORES.Count
                SQLX &= ", Sum (Decode(CUST_STORE_NO,'" & STORES(S - 1) & "'," & Mid(PERIODS_XX, 2) & ",0)) P" & Format(S, "00") & vbCrLf
                STORES_XX &= "+P" & Format(S, "00")
            Next
        End If
        sqlPOTVSUM1_STORES = SQLX

    End Sub

    Sub Print_Report()
        Call Print_Report_Begin()

        Dim SUBT As String = ""
        Dim RecordSelectionFormula As String = ""
        Generate_Report("SARCSLS1", "", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

    Private Sub optSI_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSI.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        optSI.Tag = "*"
        Load_Data()
    End Sub

    Sub Setup_grd()

        Dim CAPTION As String = "PO Receipts for " & VEND_CODE
        If optSI.Value = "S" Then
            CAPTION &= ", by Store"
        Else
            CAPTION &= ", by Item"
        End If
        grdPOTVSUM1.Text = CAPTION

        Dim g1 As UltraWinGrid.UltraGrid
        Dim g2 As UltraWinGrid.UltraGrid
        If optSI.Value = "S" Then
            g1 = grdPOTVSUM1
            g2 = grdPOTVSUM2
        Else
            g1 = grdPOTVSUM2
            g2 = grdPOTVSUM1
        End If

        If optSI.Tag = "*" Then

            optSI.Tag = ""

            With g1.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Header.Caption = "Store"
                .Columns("DESC_VALUE").Header.Caption = "Location"
                .Columns("SUB_CODE_VALUE1").Header.Caption = "Rep"
                .Columns("SUB_CODE_VALUE2").Header.Caption = "State"
                .Columns("SUB_CODE_VALUE3").Header.Caption = "City"
                .Columns("SUB_CODE_VALUE4").Header.Caption = "Zip"
                .Columns("SUB_CODE_VALUE5").Header.Caption = "Group"
                .Columns("RTL_PRICE").Header.Caption = "Retail"
                .Columns("WSL_PRICE").Header.Caption = "WhSale"

                .Columns("CODE_VALUE").Width = 60
                .Columns("DESC_VALUE").Width = 140
                .Columns("SUB_CODE_VALUE1").Width = 40
                .Columns("SUB_CODE_VALUE2").Width = 50
                .Columns("SUB_CODE_VALUE3").Width = 50
                .Columns("SUB_CODE_VALUE4").Width = 50
                .Columns("SUB_CODE_VALUE5").Width = 50
                .Columns("RTL_PRICE").Width = 65
                .Columns("WSL_PRICE").Width = 65
            End With

            With g2.DisplayLayout.Bands(0)
                .Columns("CODE_VALUE").Header.Caption = "Item"
                .Columns("DESC_VALUE").Header.Caption = "Description"
                .Columns("SUB_CODE_VALUE1").Header.Caption = "Collection"
                .Columns("SUB_CODE_VALUE2").Header.Caption = "Catgy"
                .Columns("SUB_CODE_VALUE3").Header.Caption = "Class"
                .Columns("SUB_CODE_VALUE4").Header.Caption = "Style"
                .Columns("SUB_CODE_VALUE5").Header.Caption = "Dept"
                .Columns("RTL_PRICE").Header.Caption = "Retail"
                .Columns("WSL_PRICE").Header.Caption = "WhSale"

                .Columns("CODE_VALUE").Width = 120
                .Columns("DESC_VALUE").Width = 180
                .Columns("SUB_CODE_VALUE1").Width = 80
                .Columns("SUB_CODE_VALUE2").Width = 60
                .Columns("SUB_CODE_VALUE3").Width = 60
                .Columns("SUB_CODE_VALUE4").Width = 60
                .Columns("SUB_CODE_VALUE5").Width = 60
                .Columns("RTL_PRICE").Width = 65
                .Columns("WSL_PRICE").Width = 65
            End With

            For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() _
            {grdPOTVSUM1, grdPOTVSUM2}
                With G.DisplayLayout.Bands(0)
                    If G.Name = "grdSATSLSC2" Then
                        .Columns("CODE_VALUE_PARENT").Hidden = True
                    End If
                    .Columns("CODE_VALUE").Hidden = False
                    .Columns("DESC_VALUE").Hidden = False
                    .Columns("SUB_CODE_VALUE1").Hidden = False
                    .Columns("SUB_CODE_VALUE2").Hidden = False
                    .Columns("SUB_CODE_VALUE3").Hidden = True
                    .Columns("SUB_CODE_VALUE4").Hidden = True
                    .Columns("SUB_CODE_VALUE5").Hidden = True
                    .Columns("RTL_PRICE").Hidden = True
                    .Columns("WSL_PRICE").Hidden = True
                End With
            Next

        End If

        If chkExtendedData.Checked Then
            g1.DisplayLayout.MaxBandDepth = 2
        Else
            g1.DisplayLayout.MaxBandDepth = 1
        End If

        For Each G As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {g1, g2}
            Dim BMAX As Int32 = 0
            If G Is g1 And chkExtendedData.Checked Then
                BMAX = 1
            End If
            For B As Int32 = 0 To BMAX

                With G.DisplayLayout.Bands(B)
                    If B = 1 Then
                        .Columns("DATA_TYPE").Hidden = False
                        .Columns("DATA_TYPE").ColSpan = 3
                        .Columns("DATA_TYPE").Header.Caption = "Data Type"
                        .Columns("YEAR").Hidden = False
                        .Columns("YEAR").Header.Caption = "Year"
                        .Columns("YEAR").Width = 100
                        .RowLayoutStyle = UltraWinGrid.RowLayoutStyle.None
                        .Override.AllowColSizing = UltraWinGrid.AllowColSizing.Synchronized
                    End If

                    For P As Integer = 0 To Stores_Max
                        COLUMN_NAME = "P" & Format(P, "00")
                        If optSI.Value = "I" And optXP.Value = "S" And G.Name = grdPOTVSUM1.Name Then
                            .Columns(COLUMN_NAME).Hidden = (P > STORES.Count)
                            If P <= STORES.Count Then
                                Dim LEGEND As String
                                If P = 0 Then
                                    LEGEND = "Total"
                                    .Columns(COLUMN_NAME).Width = 80
                                Else
                                    .Columns(COLUMN_NAME).Width = 70
                                    LEGEND = STORES(P - 1)
                                End If
                                .Columns(COLUMN_NAME).Header.Caption = LEGEND
                            End If
                        Else
                            .Columns(COLUMN_NAME).Hidden = (P > Periods)
                            If P <= Periods Then
                                Dim LEGEND As String
                                If P = 0 Then
                                    LEGEND = "Total"
                                    .Columns(COLUMN_NAME).Width = 80
                                Else
                                    .Columns(COLUMN_NAME).Width = 70
                                    LEGEND = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP0, P - 1))
                                    LEGEND = Mid(LEGEND, 10, 6)
                                End If
                                .Columns(COLUMN_NAME).Header.Caption = LEGEND
                            End If
                        End If
                    Next

                    .Columns("PXX").Hidden = True
                    '.Columns("PXX").Header.Caption = "O/H"

                End With
            Next
        Next

    End Sub

    Private Sub grdPOTVSUM1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdPOTVSUM1.AfterRowActivate
        If EntryMode = "E" Then
            Setup_grdPOTVSUM2()
        End If
    End Sub

    Sub Setup_grdPOTVSUM2()

        If grdPOTVSUM1.ActiveRow Is Nothing OrElse Not grdPOTVSUM1.ActiveRow.IsDataRow Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
            Exit Sub
        Else
            chkShowDetails.Enabled = True
        End If

        Dim DATA_TYPE As String = optType2.Value
        Dim CODE_VALUE_PARENT As String = grdPOTVSUM1.ActiveRow.Cells("CODE_VALUE").Text

        Load_POTVSUM2(DATA_TYPE, CODE_VALUE_PARENT, False)
        Sort_grdColumns(grdPOTVSUM2, "CODE_VALUE")

        Dim CAPTION As String = "PO Receipts (" & optType2.Text & ") for " & VEND_CODE
        If optSI.Value = "S" Then
            CAPTION &= " - Store " & CODE_VALUE_PARENT & ", by Item"
        Else
            CAPTION &= " - Item " & CODE_VALUE_PARENT & ", by Store"
        End If
        grdPOTVSUM2.Text = CAPTION

        Dim sql As String = ""
        sql = sqlSOTINVHX & " and SOTINVH2.VEND_CODE = '" & VEND_CODE & "'" & vbCrLf _
            & " and SOTINVH2.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf
        If optSI.Value = "I" Then
            sql &= " and SOTINVH2.ITEM_CODE = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        Else
            sql &= " and SOTINVH2.CUST_STORE_NO = '" & CODE_VALUE_PARENT & "'" & vbCrLf
        End If
        With grdSOTINVHX.DisplayLayout.Bands(0)
            .Columns("CUST_STORE_NO").Hidden = (optSI.Value = "S")
            .Columns("CUST_STORE_LOCATION").Hidden = (optSI.Value = "S")
            .Columns("ITEM_CODE").Hidden = (optSI.Value = "I")
            .Columns("ITEM_DESC").Hidden = (optSI.Value = "I")
        End With
        Fill_Records("SOTINVHX", "", True, sql)
        grdSOTINVHX.Text = "Sales Documents for " & VEND_CODE & IIf(optSI.Value = "S", " - Store ", " - Item ") & CODE_VALUE_PARENT
        grdSOTINVHX.DisplayLayout.CaptionVisible = DefaultableBoolean.True

    End Sub

    Sub Load_POTVSUM2(ByVal DATA_TYPE As String, ByVal CODE_VALUE_PARENT As String, ByVal all_parents As Boolean)
        Dim sql As String = ""

        If optSI.Value = "I" Then
            sql = "Select POTVSUM2.ITEM_CODE CODE_VALUE_PARENT " & vbCrLf _
            & ", POTVSUM2.CUST_STORE_NO CODE_VALUE" & vbCrLf _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') DESC_VALUE" & vbCrLf _
            & ", ARTCUST2.SELL_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_STATE SUB_CODE_VALUE2" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_CITY SUB_CODE_VALUE3" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_GROUP SUB_CODE_VALUE5" & vbCrLf _
            & sqlPOTVSUM1 & ",PXX  from ARTCUST2," & POTVSUM2 & " POTVSUM2 " & vbCrLf _
            & " where ARTCUST2.VEND_CODE (+) = POTVSUM2.VEND_CODE " & vbCrLf _
            & " and ARTCUST2.CUST_STORE_NO (+) = POTVSUM2.CUST_STORE_NO" & vbCrLf _
            & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & IIf(all_parents, "", " and POTVSUM2.ITEM_CODE = '" & CODE_VALUE_PARENT & "'")
        Else
            sql = "Select POTVSUM2.CUST_STORE_NO CODE_VALUE_PARENT " & vbCrLf _
            & ", POTVSUM2.ITEM_CODE CODE_VALUE" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
            & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
            & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
            & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
            & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
            & sqlPOTVSUM1 & ",PXX  from ICTITEM1," & POTVSUM2 & " POTVSUM2 " & vbCrLf _
            & " where ICTITEM1.ITEM_CODE (+) = POTVSUM2.ITEM_CODE " & vbCrLf _
            & " and DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
            & IIf(all_parents, "", " and POTVSUM2.CUST_STORE_NO = '" & CODE_VALUE_PARENT & "'")
        End If
        'dst.Tables("POTVSUM1").Rows.Clear()
        Fill_Records("POTVSUM2", "", True, sql)
    End Sub

    Private Sub chkNoDetails_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowDetails.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        SplitContainer1.Panel2Collapsed = Not chkShowDetails.Checked
    End Sub

    Sub Load_Data()

        dst.EnforceConstraints = False

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        Dim DATA_TYPE As String = optType2.Value

        optXP.Visible = (optSI.Value = "I")

        Dim sql As String = ""

        If optSI.Value = "S" Then
            If EntryMode = "E" Then
                sql = "Select POTVSUM1.CODE_VALUE CODE_VALUE" & vbCrLf _
                & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') DESC_VALUE" & vbCrLf _
                & ", ARTCUST2.SELL_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_STATE SUB_CODE_VALUE2" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_CITY SUB_CODE_VALUE3" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ARTCUST2.CUST_STORE_GROUP SUB_CODE_VALUE5" & vbCrLf _
                & sqlPOTVSUM1 & ",PXX from ARTCUST2," & POTVSUM1 & " POTVSUM1 " & vbCrLf _
                & " where ARTCUST2.VEND_CODE (+) = POTVSUM1.VEND_CODE " & vbCrLf _
                & " and ARTCUST2.CUST_STORE_NO (+) = POTVSUM1.CODE_VALUE" & vbCrLf _
                & " and POTVSUM1.SI = '" & optSI.Value & "' and POTVSUM1.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and POTVSUM1.YEAR = '0'"
            Else
                sql = "Select POTVSUM1.VEND_CODE CODE_VALUE" & vbCrLf _
                & ", ARTCUST1.CUST_NAME DESC_VALUE" & vbCrLf _
                & ", ARTCUST1.SREP_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ARTCUST1.CUST_STATE SUB_CODE_VALUE2" & vbCrLf _
                & ", ARTCUST1.CUST_CITY SUB_CODE_VALUE3" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ARTCUST1.TRADE_CLASS_CODE SUB_CODE_VALUE5" & vbCrLf _
                & sqlPOTVSUM1_sum & ", Sum (PXX) PXX from ARTCUST1," & POTVSUM1 & " POTVSUM1 " & vbCrLf _
                & " where ARTCUST1.VEND_CODE (+) = POTVSUM1.VEND_CODE " & vbCrLf _
                & " and POTVSUM1.SI = '" & optSI.Value & "' and POTVSUM1.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and POTVSUM1.YEAR = '0'" & vbCrLf _
                & " group by POTVSUM1.VEND_CODE" & vbCrLf _
                & ", ARTCUST1.CUST_NAME" & vbCrLf _
                & ", ARTCUST1.SREP_CODE" & vbCrLf _
                & ", ARTCUST1.CUST_STATE" & vbCrLf _
                & ", ARTCUST1.CUST_CITY" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE" & vbCrLf _
                & ", ARTCUST1.TRADE_CLASS_CODE"
            End If
        Else
            If optXP.Value = "S" Then

                Dim SQLX As String = sqlPOTVSUM1_STORES
                If Mid(DATA_TYPE, 1, 1) = "H" Then
                    SQLX = Replace(sqlPOTVSUM1_STORES, Mid(PERIODS_XX, 2), "P" & Format(Periods, "00"))
                End If

                sql = "Select POTVSUM2.ITEM_CODE CODE_VALUE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
                & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
                & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
                & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
                & SQLX & "  from ICTITEM1," & POTVSUM2 & " POTVSUM2 " & vbCrLf _
                & " where ICTITEM1.ITEM_CODE (+) = POTVSUM2.ITEM_CODE " & vbCrLf _
                & " and POTVSUM2.DATA_TYPE = '" & DATA_TYPE & "'" & vbCrLf _
                & " and POTVSUM2.YEAR = '0'" & vbCrLf _
                & " group by POTVSUM2.ITEM_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_CLASS_CODE" & vbCrLf _
                & ", ICTITEM1.STYLE_CODE" & vbCrLf _
                & ", ICTITEM1.DEPT_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & ", ICTITEM1.ITEM_PRICE" & vbCrLf
            Else
                sql = "Select POTVSUM1.CODE_VALUE CODE_VALUE" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC DESC_VALUE" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE SUB_CODE_VALUE1" & vbCrLf _
                & ", ICTITEM1.ITEM_CATGY_CODE SUB_CODE_VALUE2" & vbCrLf _
                & ", ICTITEM1.ITEM_CLASS_CODE SUB_CODE_VALUE3" & vbCrLf _
                & ", ICTITEM1.STYLE_CODE SUB_CODE_VALUE4" & vbCrLf _
                & ", ICTITEM1.DEPT_CODE SUB_CODE_VALUE5" & vbCrLf _
                & ", ICTITEM1.ITEM_RETAIL_PRICE RTL_PRICE" & vbCrLf _
                & ", ICTITEM1.ITEM_PRICE WSL_PRICE" & vbCrLf _
                & sqlPOTVSUM1 & ",PXX from ICTITEM1," & POTVSUM1 & " POTVSUM1 " & vbCrLf _
                & " where ICTITEM1.ITEM_CODE (+) = POTVSUM1.CODE_VALUE " & vbCrLf _
                & " and POTVSUM1.SI = '" & optSI.Value & "'" _
                & " and POTVSUM1.DATA_TYPE = '" & DATA_TYPE & "'" _
                & " and POTVSUM1.YEAR = '0'" & vbCrLf
            End If
        End If

        dst.Tables("POTVSUM1").Rows.Clear()
        dst.Tables("POTVSUM1_DTL").Rows.Clear()
        dst.Tables("POTVSUM2").Rows.Clear()

        If optSI.Value = "I" And optXP.Value = "S" Then
            dst.Tables("POTVSUM1").Columns("P00").Expression = Mid(STORES_XX, 2)
            dst.Tables("POTVSUM1_DTL").Columns("P00").Expression = Mid(STORES_XX, 2)
        Else
            dst.Tables("POTVSUM1").Columns("P00").Expression = Mid(PERIODS_XX, 2)
            dst.Tables("POTVSUM1_DTL").Columns("P00").Expression = "IIF(DATA_TYPE='HU' OR DATA_TYPE='HD'," & Mid(PERIODS_XX, Len(PERIODS_XX) - 2, 3) & "," & Mid(PERIODS_XX, 2) & ")"
        End If

        dst.Tables("POTVSUM2").Columns("P00").Expression = Mid(PERIODS_XX, 2)

        Fill_Records("POTVSUM1", "", True, sql)
        Sort_grdColumns(grdPOTVSUM1, "CODE_VALUE")

        If chkExtendedData.Checked Then
            sql = "Select POTVSUM1.CODE_VALUE CODE_VALUE, POTVSUM1.YEAR, POTVSUM1.DATA_TYPE" _
            & sqlPOTVSUM1 _
            & " from " & POTVSUM1 & " POTVSUM1 " _
            & " where POTVSUM1.SI = '" & optSI.Value & "'" _
            & " and POTVSUM1.DATA_TYPE LIKE '%" & Mid(DATA_TYPE, 2, 1) & "'" _
            & " and (YEAR = '1' OR POTVSUM1.DATA_TYPE <> '" & DATA_TYPE & "')"

            Fill_Records("POTVSUM1_DTL", "", True, sql)
            Sort_grdColumns(grdPOTVSUM1, "DATA_TYPE", , 1)
        End If

        If grdPOTVSUM1.Rows.Count = 0 Then
            chkShowDetails.Checked = False
            chkShowDetails.Enabled = False
        Else
            chkShowDetails.Enabled = True
        End If

        'dst.EnforceConstraints = True

        Setup_grd()
 
        If optSI.Value = "S" Then
            If grdPOTVSUM1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
                grdPOTVSUM1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
            End If
            If grdPOTVSUM2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
                grdPOTVSUM2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
            End If

            ASCMAIN1.Add_Value_List(grdPOTVSUM1, "SUB_CODE_VALUE1", , , , "Select SELL_CODE, SELL_NAME from SOTSELL1")
            ASCMAIN1.Add_Value_List(grdPOTVSUM2, "SUB_CODE_VALUE2", , , , "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
        Else
            If grdPOTVSUM1.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE1") Then
                grdPOTVSUM1.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE1")
            End If
            If grdPOTVSUM2.DisplayLayout.ValueLists.Exists("SUB_CODE_VALUE2") Then
                grdPOTVSUM2.DisplayLayout.ValueLists.Remove("SUB_CODE_VALUE2")
            End If

            ASCMAIN1.Add_Value_List(grdPOTVSUM1, "SUB_CODE_VALUE2", , , , "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1")
            ASCMAIN1.Add_Value_List(grdPOTVSUM2, "SUB_CODE_VALUE1", , , , "Select SELL_CODE, SELL_NAME from SOTSELL1")
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub optType1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If SELECTION_NO = 0 Then Exit Sub

        Load_Data()

    End Sub

    Private Sub optType2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optType2.ValueChanged

        If SELECTION_NO = 0 Then Exit Sub
        Load_Data()
    End Sub

    Private Sub tabDetails_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabDetails.SelectedTabChanged
        Setup_tabDetails()
    End Sub

    Sub Setup_tabDetails()
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

    Private Sub optXP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optXP.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If optXP.Value = "S" Then
            If STORES.Count > Stores_Max Then
                MsgBox("Too Many Stores (" & STORES.Count & ") for this option.  Max is " & CStr(Stores_Max))
                optXP.Value = "P"
                Exit Sub
            End If
        End If
        Load_Data()
    End Sub

    Private Sub chkExtendedData_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkExtendedData.CheckedChanged
        If Not chkExtendedData.Checked Then
            chkPriorYear.Checked = False
        End If
        chkPriorYear.Enabled = chkExtendedData.Checked
    End Sub

    Sub Validate_Range(ByRef EMsg As String)

        If Absx1.optFor("RANGE").Value = "P" Then
            If Absx1.cmbFor("RYP0").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Starting Period"
            End If
            If Absx1.cmbFor("RYP1").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify an Ending Period"
            End If

            If EMsg = "" Then
                RYP0 = Absx1.cmbFor("RYP0").Value
                RYP1 = Absx1.cmbFor("RYP1").Value
                Periods = ASCMAIN1.Period_Diff(RYP0, RYP1) + 1
            End If
        Else
            If Absx1.cmbFor("RYW0").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Starting Week"
            End If
            If Absx1.cmbFor("RYW1").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify an Ending Week"
            End If
            'If Absx1.dteFor("RYW0").Value > Absx1.dteFor("RYW1").Value Then
            '    EMsg &= vbCr & "Starting Week cannot be later than Ending Week"
            'End If

            If EMsg = "" Then
                RYW0 = Absx1.cmbFor("RYW0").Value
                RYW1 = Absx1.cmbFor("RYW1").Value
                Periods = ASCMAIN1.Week_Diff(RYW0, RYW1) + 1
            End If
        End If

        If EMsg = "" Then
            If Periods < 1 Or Periods > Stores_Max Then
                EMsg &= vbCr & "Total number of Periods must be between 1 and " & CStr(Stores_Max)
            End If
        End If

    End Sub

    Private Sub grdPOTVSUM1_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdPOTVSUM1.DoubleClickRow
        If EntryMode = "" Then
            Absx1.txtFor("VEND_CODE").Text = e.Row.Cells("CODE_VALUE").Value
            Click_Command("Load")
        End If
    End Sub
End Class