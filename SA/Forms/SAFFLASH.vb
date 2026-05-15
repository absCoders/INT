Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Shared.Styles

Public Class SAFFLASH

    Dim RYP As String
    Dim SATGRNR1 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            ' Create_SATGRNR1("")
            ASCMAIN1.sql = "Select * from " & SATGRNR1
            Create_TDA(.Tables.Add, "SATGRNR1", "**", 0, False, "", 0)

            With .Tables("SATGRNR1")
                .Columns.Add("MTD_NET", GetType(System.Int64), "MTD_GRS - MTD_RTN")
                .Columns.Add("YTD_NET", GetType(System.Int64), "YTD_GRS - YTD_RTN")
                For i As Integer = 1 To 12
                    Dim ret As String = "RTN_" & Format(i, "00")
                    Dim grs As String = "GRS_" & Format(i, "00")
                    With dst
                        With .Tables("SATGRNR1")
                            .Columns.Add("NET_" & Format(i, "00"), GetType(System.Int64), grs & " - " & ret)
                        End With
                    End With
                Next
            End With
        End With

        grdSATGRNR1.DataSource = dst.Tables("SATGRNR1")
        grdSATGRNR2.DataSource = dst.Tables("SATGRNR1")


        For Each COLUMN_NAME As String In New String() {"MTD_GRS", "MTD_RTN", "MTD_NET"} ', "YTD_GRS", "YTD_RTN", "YTD_NET"}
            grdSATGRNR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME).Hidden = True
        Next

        For Each COLUMN_SFX As String In New String() {"GRS", "RTN", "NET"}
            For i As Integer = 1 To 12
                COLUMN_NAME = COLUMN_SFX & "_" & Format(i, "00")
                With grdSATGRNR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .Format = "###,##0"
                    If COLUMN_SFX = "GRS" Then
                        .Header.Appearance.BackColor = Drawing.Color.Lime
                    ElseIf COLUMN_SFX = "RTN" Then
                        .Header.Appearance.BackColor = Drawing.Color.Orange
                    ElseIf COLUMN_SFX = "NET" Then
                        .Header.Appearance.BackColor = Drawing.Color.DodgerBlue
                    End If
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                End With
                Create_Summary(grdSATGRNR2, COLUMN_NAME)
            Next
        Next


        For Each COLUMN_SFX As String In New String() {"GRS", "RTN", "NET"}
            For Each COLUMN_PFX As String In New String() {"MTD", "YTD"}
                COLUMN_NAME = COLUMN_PFX & "_" & COLUMN_SFX
                With grdSATGRNR1.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .Header.Caption = ASCMAIN1.Make_Caption(.Header.Caption)
                    .Format = "###,##0"
                    If COLUMN_SFX = "GRS" Then
                        .Header.Appearance.BackColor = Drawing.Color.Lime
                    ElseIf COLUMN_SFX = "RTN" Then
                        .Header.Appearance.BackColor = Drawing.Color.Orange
                    ElseIf COLUMN_SFX = "NET" Then
                        .Header.Appearance.BackColor = Drawing.Color.DodgerBlue
                    End If
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                End With
                With grdSATGRNR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .Header.Caption = ASCMAIN1.Make_Caption(.Header.Caption)
                    .Format = "###,##0"
                    If COLUMN_SFX = "GRS" Then
                        .Header.Appearance.BackColor = Drawing.Color.Lime
                    ElseIf COLUMN_SFX = "RTN" Then
                        .Header.Appearance.BackColor = Drawing.Color.Orange
                    ElseIf COLUMN_SFX = "NET" Then
                        .Header.Appearance.BackColor = Drawing.Color.DodgerBlue
                    End If
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                End With
                Create_Summary(grdSATGRNR1, COLUMN_NAME)
                Create_Summary(grdSATGRNR2, COLUMN_NAME)
            Next
        Next

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATGRNR1, grdSATGRNR2}
            For Each COLUMN_NAME As String In New String() {"CUST_DBA_NAME", "SREP_CODE"}
                With grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .Hidden = True
                End With
            Next
            For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_NAME"}
                With grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.GlassBottom20
                End With
            Next
            With grd.DisplayLayout.Bands(0)
                .Columns("CUST_CODE").Header.Caption = "Customer"
                .Columns("CUST_NAME").Header.Caption = "Customer Name"
                .SummaryFooterCaption = "Total [GROUPBYROWVALUE]"
            End With

            grd.DisplayLayout.GroupByBox.Hidden = True
            For Each COLUMN_NAME As String In New String() {"COMPANY", "CLASS", "REG"}
                grd.DisplayLayout.Bands(0).Columns(COLUMN_NAME).SortComparer = New srtComparerSATGRNR1
            Next
        Next

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Call Validate_Code("OPS_YYYYPP")
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
                EntryMode = "E"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

            Case "Print"
                Call Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                '.Groups("Screen Control").Items("Print").Settings.Enabled = iScreenMode
                .Groups("View Options").Visible = ScreenMode
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        UltraTabControl1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("SATGRNR1").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Compiling Historical Data")
        Application.DoEvents()

        Save_Header_Fields(UltraGroupBox1)

        RYP = Absx1.txtFor("OPS_YYYYPP").Text
        Create_SATGRNR1(RYP)

        ASCMAIN1.Progress("Now Loading Data from Database")

        EnforceConstraints(False)
        Fill_Records("SATGRNR1")
        EnforceConstraints(True)

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSATGRNR1, grdSATGRNR2}
            With grd.DisplayLayout.Bands(0)
                For Each COLUMN_NAME As String In New String() {"COMPANY", "CLASS", "REG", "GRP"}
                    .Columns(COLUMN_NAME).HiddenWhenGroupBy = DefaultableBoolean.True
                    .SortedColumns.Add(COLUMN_NAME, False, True)
                Next
                .SortedColumns.Add("CUST_CODE", False, False)
            End With
            grd.Rows.ExpandAll(True)
            Check_to_Collaps(grd.Rows)
        Next

        Dim TEST As Integer = CInt(Today.Year) & CInt(Today.Month)
        Dim TEST2 = Mid(RYP, 1, 6)

        For Each COLUMN_SFX As String In New String() {"GRS", "RTN", "NET"}
            For i As Integer = 1 To 12
                COLUMN_NAME = COLUMN_SFX & "_" & Format(i, "00")
                With grdSATGRNR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    '.Hidden = (CInt(TEST2) > TEST)
                    Dim legend As String = ASCMAIN1.Get_Legend(Mid(RYP, 1, 4) & Format(i, "00"), False, True)
                    .Header.Caption = legend
                    .Hidden = (i > Val(Mid(RYP, 5, 2)))
                End With

                With grdSATGRNR1.DisplayLayout.Bands(0)
                    .Columns("NET_" & Format(i, "00")).Hidden = True
                    .Columns("GRS_" & Format(i, "00")).Hidden = True
                    .Columns("RTN_" & Format(i, "00")).Hidden = True
                End With
            Next
        Next



        ASCMAIN1.Progress("Now Setting Up Screen")

        'Sort_grdColumns(grdSATGRNR1, "DEPT_CODE,PROD_CODE,MATL_CODE,ITEM_CATGY_CODE")

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
        Call Load_Popup_Menu(grdSATGRNR1, "S", "Show Filter")
        Call Load_Popup_Menu(grdSATGRNR2, "S", "Show Filter")
    End Sub

    Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"

            End Select

        End If
    End Sub

    Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

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

    Sub Print_Report()
        Dim SUBT As String = ""

        Call Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Call Print_Report_End()
    End Sub


    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles UltraTabControl1.SelectedTabChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_tab()
    End Sub

    Sub Setup_tab()
        UltraExplorerBar1.Groups("View Options").Visible = ScreenMode  'And (UltraTabControl1.SelectedTab.Key = "Detail by Month")
        chkGross.Enabled = ScreenMode And (UltraTabControl1.SelectedTab.Key = "Detail by Month")
        chkReturns.Enabled = ScreenMode And (UltraTabControl1.SelectedTab.Key = "Detail by Month")
        chkNet.Enabled = ScreenMode And (UltraTabControl1.SelectedTab.Key = "Detail by Month")
    End Sub









#Region "VB6 Declarations"

    'TOTALS NEED TO BE FORMULA-IZED
    'Hyperlinks
    ' flash1 does not really look like the original  - others s/b on the bottom of the total page
    ' IMPLEMENT A ROLLUP CUSTOMER CODE CONCEPT SO THAT DILLARDS, NORDS, ETC WILL ROLLUP

    Public SF As String
    Public L2 As String
    Public TBL As String
    Public RF As String
    Public FT As String

    Dim RYP As String
    Dim RYM As String
    Dim TYP0 As String
    Dim TYP1 As String
    Dim TPDTE As String
    Dim LYP0 As String
    Dim LYP1 As String
    Dim TTB As String

    Dim DTE1 As String
    Dim DTE2 As String
    Dim DATE1 As Date
    Dim DATE2 As Date

    Dim LEVELS As Integer
    Dim r0 As Integer
    Dim c0 As Integer

    Dim XLDATE As Date
#End Region

    Sub Load_SAWSLSF1()

        Sql = "Select * from SATSLSF1"
        Call Ora_to_Acc(Nothing, "SAWSLSF1", 2, "", Sql)

        Dim z As String

        z = "?"
        AccD.Execute("Update SAWSLSF1 Set SALES_DIVISION_CODE = '" & z & "' where SALES_DIVISION_CODE is Null")
        AccD.Execute("Update SAWSLSF1 Set BUS_UNIT_CODE = '" & z & "' where BUS_UNIT_CODE is Null")
        AccD.Execute("Update SAWSLSF1 Set TRADE_CLASS_CODE = '" & z & "' where TRADE_CLASS_CODE is Null")
        AccD.Execute("Update SAWSLSF1 Set MARKET_CODE = '" & z & "' where MARKET_CODE is Null")
        AccD.Execute("Update SAWSLSF1 Set SREP_CODE = '" & z & "' where SREP_CODE is Null")
        AccD.Execute("Update SAWSLSF1 Set REGION_CODE = '" & z & "' where REGION_CODE is Null")
        AccD.Execute("Update SAWSLSF1 Set VP_CODE = '" & z & "' where VP_CODE is Null")

        Sql = "Select Distinct CUST_CODE from SATSLSF1"
        Call Ora_to_Acc(Nothing, "SAWSLSF2C", 1, "", Sql)

        Sql = "Select Distinct REGION_CODE from SATSLSF1 where REGION_CODE is Not Null"
        Call Ora_to_Acc(Nothing, "SAWSLSF2R", 1, "", Sql)

        Sql = "Select VP_CODE, VP_NAME from SOTSVPS1"
        Call Ora_to_Acc(Nothing, "SOWSVPS1", 1, "", Sql)
        Sql = "Select SALES_DIVISION_CODE, SALES_DIVISION_NAME from SOTSDIV1"
        Call Ora_to_Acc(Nothing, "SOWSDIV1", 1, "", Sql)
        Sql = "Select ITEM_BRAND_CODE, ITEM_BRAND_NAME from ICTBRAN1"
        Call Ora_to_Acc(Nothing, "ICWBRAN1", 1, "", Sql)
        Sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
        Call Ora_to_Acc(Nothing, "ARWCUST1", 1, "", Sql)
        Sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
        Call Ora_to_Acc(Nothing, "SOWSREP1", 1, "", Sql)
        Sql = "Select REGION_CODE, REGION_DESC from SOTSREG1"
        Call Ora_to_Acc(Nothing, "SOWSREG1", 1, "", Sql)
        Sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC from SOTTCLS1"
        Call Ora_to_Acc(Nothing, "SOWTCLS1", 1, "", Sql)
        Sql = "Select MARKET_CODE, MARKET_DESC from SOTMKTC1"
        Call Ora_to_Acc(Nothing, "SOWMKTC1", 1, "", Sql)
        Sql = "Select BUS_UNIT_CODE, BUS_UNIT_DESC from SOTBUSU1"
        Call Ora_to_Acc(Nothing, "SOWBUSU1", 1, "", Sql)

        '    Sql = "Select"
        '    Call Ora_to_Acc(Nothing, "SAWSLSF2A", 3, "", Sql)
    End Sub
  
    Private Sub cmdGenerate_Click()

        Dim ETIME As Date
        ETIME = Now

        If Me.Controls(Tag_to_Index("DATE_FROM")).Text = "" Or Me.Controls(Tag_to_Index("DATE_TO")).Text = "" Then
            MsgBox("Please Specify Valid From and To Dates", vbOKOnly, "Date Range is Mandatory")
            Exit Sub
        End If

        If cmbYP.Text = "" Then
            MsgBox("Please Specify a Valid Reporting Period", vbOKOnly, "Period is Mandatory")
            Exit Sub
        End If

        Call Setup_Memory()

        If Format$(Me.Controls(Tag_to_Index("DATE_FROM")).Text, "YYYYMMDD") > Format$(Me.Controls(Tag_to_Index("DATE_TO")).Text, "YYYYMMDD") Then
            MsgBox("From Date is later than To Date", vbOKOnly, "Date Range does not make sense")
            Exit Sub
        End If

        If Format$(Me.Controls(Tag_to_Index("DATE_FROM")).Text, "YYYYMMDD") > Format$(DATE2, "YYYYMMDD") Then
            MsgBox("From Date is later than Ending Date of Period", vbOKOnly, "Date Range does not make sense")
            Exit Sub
        End If
        If Format$(Me.Controls(Tag_to_Index("DATE_FROM")).Text, "YYYYMMDD") < Format$(DATE1, "YYYYMMDD") Then
            MsgBox("From Date is prior to Beginning Date of Period", vbOKOnly, "Date Range does not make sense")
            Exit Sub
        End If
        If Format$(Me.Controls(Tag_to_Index("DATE_TO")).Text, "YYYYMMDD") < Format$(DATE1, "YYYYMMDD") Then
            MsgBox("To Date is prior to Beginning Date of Period", vbOKOnly, "Date Range does not make sense")
            Exit Sub
        End If
        If Format$(Me.Controls(Tag_to_Index("DATE_TO")).Text, "YYYYMMDD") > Format$(DATE2, "YYYYMMDD") Then
            MsgBox("To Date is later than Ending Date of Period", vbOKOnly, "Date Range does not make sense")
            Exit Sub
        End If

        Me.MousePointer = 11

        If chkUsePrevious.Value = "0" Then
            Call Generate_Flash_Workfile()
        End If
        Load_SAWSLSF1()

        Generate_Flash_XLS_1()
        Generate_Flash_XLS_2()

        Me.MousePointer = 0
        Prompt("", "")

        ETIME = Now - ETIME
        MsgBox("Total elapsed time = " & Format$(ETIME * 60 * 60, "#.00") & " Mins", vbOKOnly, "Step 1 is Complete")
    End Sub

    Sub Setup_Memory()
        Dim z As String
        z = cmbYP.Text
        RYP = Mid$(z, 1, 4) & Mid$(z, 6, 2)
        RYM = Get_YYYYMM(RYP, 0)

        TYP0 = Mid$(RYP, 1, 4) & "01"
        TYP1 = Mid$(RYP, 1, 4) & Mid$(RYP, 5, 2)

        Dim dynGLTPARM2 As OraDynaset

        Sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & Period_Calc(TYP1, -1) & "'"
        dynGLTPARM2 = OraD.CreateDynaset(Sql, 8&)
        DATE1 = dynGLTPARM2.Fields("PRD_END_DATE").Value + 1

        Sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & TYP1 & "'"
        dynGLTPARM2 = OraD.CreateDynaset(Sql, 8&)
        DATE2 = dynGLTPARM2.Fields("PRD_END_DATE").Value
        TPDTE = Format$(dynGLTPARM2.Fields("PRD_END_DATE").Value, "DD-MMM-YYYY")

        LYP0 = Format$(Val(Mid$(RYP, 1, 4)) - 1, "0000") & "01"
        LYP1 = Format$(Val(Mid$(RYP, 1, 4)) - 1, "0000") & Mid$(RYP, 5, 2)

        DTE1 = Format$(Me.Controls(Tag_to_Index("DATE_FROM")).Text, "DD-MMM-YYYY")
        DTE2 = Format$(Me.Controls(Tag_to_Index("DATE_TO")).Text, "DD-MMM-YYYY")

    End Sub
  
    Sub Generate_Flash_Workfile()
        Call Prompt("Now Creating Work Tables", "")

        If RYP = CYP Then
            'If UserID = "wjz" Then
            '    Stop
            '    OraS.BeginTrans
            '    Call Update_Sales_Summary(CYP)
            '    OraS.CommitTrans
            'End If
        End If

        Dim z As String

        z = " / 1000"

        Sql = "Select ITEM_BRAND_CODE, CUST_CODE" & vbCr
        Sql = Sql & ", Sum (TY_MTD_GRS" & z & ") TY_MTD_GRS" & vbCr
        Sql = Sql & ", Sum (TY_MTD_RTN" & z & ") TY_MTD_RTN" & vbCr
        Sql = Sql & ", Sum (TY_YTD_GRS" & z & ") TY_YTD_GRS" & vbCr
        Sql = Sql & ", Sum (TY_YTD_RTN" & z & ") TY_YTD_RTN" & vbCr
        Sql = Sql & ", Sum (LY_MTD_GRS" & z & ") LY_MTD_GRS" & vbCr
        Sql = Sql & ", Sum (LY_MTD_RTN" & z & ") LY_MTD_RTN" & vbCr
        Sql = Sql & ", Sum (LY_YTD_GRS" & z & ") LY_YTD_GRS" & vbCr
        Sql = Sql & ", Sum (LY_YTD_RTN" & z & ") LY_YTD_RTN" & vbCr

        Sql = Sql & ", Sum (TY_MTD_BDG" & z & ") TY_MTD_BDG" & vbCr
        Sql = Sql & ", Sum (TY_YTD_BDG" & z & ") TY_YTD_BDG" & vbCr
        Sql = Sql & ", Sum (TY_MTD_BDR" & z & ") TY_MTD_BDR" & vbCr
        Sql = Sql & ", Sum (TY_YTD_BDR" & z & ") TY_YTD_BDR" & vbCr

        Sql = Sql & ", Sum (TY_DLY_BKG" & z & ") TY_DLY_BKG" & vbCr
        Sql = Sql & ", Sum (TY_MTD_BKG" & z & ") TY_MTD_BKG" & vbCr
        Sql = Sql & ", Sum (TY_YTD_BKG" & z & ") TY_YTD_BKG" & vbCr

        Sql = Sql & ", Sum (TY_OTS_REL" & z & ") TY_OTS_REL" & vbCr
        Sql = Sql & ", Sum (TY_OTS_CUR" & z & ") TY_OTS_CUR" & vbCr
        Sql = Sql & ", Sum (TY_OTS_FUT" & z & ") TY_OTS_FUT" & vbCr
        Sql = Sql & " from (" & vbCr

        Sql = Sql & "SELECT GLTPROF0.ITEM_BRAND_CODE, GLTPROF0.CUST_CODE" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'I' AND GLTPROF0.OPS_YYYYPP = '" & TYP1 & "'" & vbCr
        Sql = Sql & " THEN GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) TY_MTD_GRS" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'C' AND GLTPROF0.OPS_YYYYPP = '" & TYP1 & "'" & vbCr
        Sql = Sql & " THEN -1 * GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) TY_MTD_RTN" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'I' AND GLTPROF0.OPS_YYYYPP BETWEEN '" & TYP0 & "' AND '" & TYP1 & "'" & vbCr
        Sql = Sql & " THEN GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) TY_YTD_GRS" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'C' AND GLTPROF0.OPS_YYYYPP BETWEEN '" & TYP0 & "' AND '" & TYP1 & "'" & vbCr
        Sql = Sql & " THEN -1 * GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) TY_YTD_RTN" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'I' AND GLTPROF0.OPS_YYYYPP = '" & LYP1 & "'" & vbCr
        Sql = Sql & " THEN GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) LY_MTD_GRS" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'C' AND GLTPROF0.OPS_YYYYPP = '" & LYP1 & "'" & vbCr
        Sql = Sql & " THEN -1 * GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) LY_MTD_RTN" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'I' AND GLTPROF0.OPS_YYYYPP BETWEEN '" & LYP0 & "' AND '" & LYP1 & "'" & vbCr
        Sql = Sql & " THEN GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) LY_YTD_GRS" & vbCr
        Sql = Sql & ", SUM (CASE WHEN GLTPROF0.INV_TYPE = 'C' AND GLTPROF0.OPS_YYYYPP BETWEEN '" & LYP0 & "' AND '" & LYP1 & "'" & vbCr
        Sql = Sql & " THEN -1 * GLTPROF0.ORDR_AMT_SHIP ELSE 0 END) LY_YTD_RTN" & vbCr
        Sql = Sql & ", 0 TY_MTD_BDG, 0 TY_YTD_BDG" & vbCr
        Sql = Sql & ", 0 TY_MTD_BDR, 0 TY_YTD_BDR" & vbCr
        Sql = Sql & ", 0 TY_DLY_BKG, 0 TY_MTD_BKG, 0 TY_YTD_BKG" & vbCr
        Sql = Sql & ", 0 TY_OTS_REL, 0 TY_OTS_CUR, 0 TY_OTS_FUT" & vbCr
        Sql = Sql & " From GLTPROF0" & vbCr
        Sql = Sql & " WHERE GLTPROF0.OPS_YYYYPP BETWEEN '" & LYP0 & "' AND '" & TYP1 & "'" & vbCr
        Sql = Sql & "   AND NVL(GLTPROF0.ORDR_AMT_SHIP,0) <> 0" & vbCr
        Sql = Sql & " GROUP BY GLTPROF0.ITEM_BRAND_CODE, GLTPROF0.CUST_CODE" & vbCr

        Sql = Sql & " UNION " & vbCr

        Sql = Sql & "Select SOTBUDBC.ITEM_BRAND_CODE, SOTBUDBC.CUST_CODE" & vbCr
        Sql = Sql & ", 0 TY_MTD_GRS, 0 TY_MTD_RTN, 0 TY_YTD_GRS, 0 TY_YTD_RTN" & vbCr
        Sql = Sql & ", 0 LY_MTD_GRS, 0 LY_MTD_RTN, 0 LY_YTD_GRS, 0 LY_YTD_RTN" & vbCr
        Sql = Sql & ", SUM (DECODE(SOTBUDBC.OPS_YYYYPP_FC,'" & TYP1 & "',SOTBUDBC.BUDGET_GRS,0)) TY_MTD_BDG" & vbCr
        Sql = Sql & ", SUM (CASE WHEN SOTBUDBC.OPS_YYYYPP_FC BETWEEN '" & TYP0 & "' AND '" & TYP1 & "'" & vbCr
        Sql = Sql & " THEN SOTBUDBC.BUDGET_GRS ELSE 0 END) TY_YTD_BDG" & vbCr
        Sql = Sql & ", SUM (DECODE(SOTBUDBC.OPS_YYYYPP_FC,'" & TYP1 & "',SOTBUDBC.BUDGET_RTN,0)) TY_MTD_BDR" & vbCr
        Sql = Sql & ", SUM (CASE WHEN SOTBUDBC.OPS_YYYYPP_FC BETWEEN '" & TYP0 & "' AND '" & TYP1 & "'" & vbCr
        Sql = Sql & " THEN SOTBUDBC.BUDGET_RTN ELSE 0 END) TY_YTD_BDR" & vbCr
        Sql = Sql & ", 0 TY_DLY_BKG, 0 TY_MTD_BKG, 0 TY_YTD_BKG" & vbCr
        Sql = Sql & ", 0 TY_OTS_REL, 0 TY_OTS_CUR, 0 TY_OTS_FUT" & vbCr
        Sql = Sql & " FROM SOTBUDBC" & vbCr
        Sql = Sql & " WHERE SOTBUDBC.OPS_YYYYPP = '" & TYP1 & "'" & vbCr
        Sql = Sql & " AND (NVL(SOTBUDBC.BUDGET_GRS,0) <> 0 OR NVL(SOTBUDBC.BUDGET_RTN,0) <> 0)"
        Sql = Sql & " group by SOTBUDBC.ITEM_BRAND_CODE, SOTBUDBC.CUST_CODE" & vbCr

        Sql = Sql & " UNION " & vbCr

        Sql = Sql & "SELECT ICTITEM1.ITEM_BRAND_CODE, SOTORDR1.CUST_CODE" & vbCr
        Sql = Sql & ", 0 TY_MTD_GRS, 0 TY_MTD_RTN, 0 TY_YTD_GRS, 0 TY_YTD_RTN" & vbCr
        Sql = Sql & ", 0 LY_MTD_GRS, 0 LY_MTD_RTN, 0 LY_YTD_GRS, 0 LY_YTD_RTN" & vbCr
        Sql = Sql & ", 0 TY_MTD_BDG, 0 TY_YTD_BDG" & vbCr
        Sql = Sql & ", 0 TY_MTD_BDR, 0 TY_YTD_BDR" & vbCr
        Sql = Sql & ", SUM (CASE WHEN TRUNC(SOTORDR1.ORDR_DATE_RECD) BETWEEN '" & DTE1 & "' AND '" & DTE2 & "' THEN SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END) TY_DLY_BKG" & vbCr
        Sql = Sql & ", SUM (DECODE(SOTORDR1.ORDR_YYYYPP_BOOKED,'" & TYP1 & "',SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE,0)) TY_MTD_BKG" & vbCr
        Sql = Sql & ", SUM (CASE WHEN SOTORDR1.ORDR_YYYYPP_BOOKED BETWEEN '" & TYP0 & "' AND '" & TYP1 & "'" & vbCr
        Sql = Sql & " THEN SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE / 1000  ELSE 0 END) TY_YTD_BKG" & vbCr
        Sql = Sql & ", 0 TY_OTS_REL, 0 TY_OTS_CUR, 0 TY_OTS_FUT" & vbCr
        Sql = Sql & " From SOTORDR1, SOTORDR2, ICTITEM1" & vbCr
        Sql = Sql & " Where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCr
        Sql = Sql & " AND ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCr
        Sql = Sql & " AND SOTORDR1.ORDR_YYYYPP_BOOKED BETWEEN '" & TYP0 & "' AND '" & TYP1 & "'" & vbCr
        Sql = Sql & " AND SOTORDR1.ORDR_STATUS <> 'D'" & vbCr
        Sql = Sql & " AND SOTORDR2.ORDR_QTY <> 0" & vbCr
        Sql = Sql & " AND SOTORDR2.ORDR_UNIT_PRICE <> 0" & vbCr
        Sql = Sql & " GROUP BY ICTITEM1.ITEM_BRAND_CODE, SOTORDR1.CUST_CODE" & vbCr

        Sql = Sql & " UNION " & vbCr

        Sql = Sql & "SELECT ICTITEM1.ITEM_BRAND_CODE, SOTORDR1.CUST_CODE" & vbCr
        Sql = Sql & ", 0 TY_MTD_GRS, 0 TY_MTD_RTN, 0 TY_YTD_GRS, 0 TY_YTD_RTN" & vbCr
        Sql = Sql & ", 0 LY_MTD_GRS, 0 LY_MTD_RTN, 0 LY_YTD_GRS, 0 LY_YTD_RTN" & vbCr
        Sql = Sql & ", 0 TY_MTD_BDG, 0 TY_YTD_BDG" & vbCr
        Sql = Sql & ", 0 TY_MTD_BDR, 0 TY_YTD_BDR" & vbCr
        Sql = Sql & ", 0 TY_DLY_BKG, 0 TY_MTD_BKG, 0 TY_YTD_BKG" & vbCr
        Sql = Sql & ", SUM (DECODE(SOTORDR1.ORDR_STATUS,'R',SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE,0)) TY_OTS_REL" & vbCr
        Sql = Sql & ", SUM (DECODE(SOTORDR1.ORDR_STATUS,'O',(CASE WHEN SOTORDR1.ORDR_SHIP_DATE <= '" & TPDTE & "' THEN SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END),0)) TY_OTS_CUR" & vbCr
        Sql = Sql & ", SUM (DECODE(SOTORDR1.ORDR_STATUS,'O',(CASE WHEN SOTORDR1.ORDR_SHIP_DATE >  '" & TPDTE & "' THEN SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE ELSE 0 END),0)) TY_OTS_FUT" & vbCr
        Sql = Sql & " From SOTORDR1, SOTORDR2, ICTITEM1" & vbCr
        Sql = Sql & " Where SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCr
        Sql = Sql & " AND ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCr
        Sql = Sql & " AND SOTORDR1.ORDR_STATUS IN ('O','R')" & vbCr
        Sql = Sql & " AND SOTORDR2.ORDR_QTY <> 0" & vbCr
        Sql = Sql & " AND SOTORDR2.ORDR_UNIT_PRICE <> 0" & vbCr
        Sql = Sql & " GROUP BY ICTITEM1.ITEM_BRAND_CODE, SOTORDR1.CUST_CODE" & vbCr

        Sql = Sql & ") group by ITEM_BRAND_CODE, CUST_CODE" & vbCr

        Dim TT As String
        TT = Temp_Table(Sql)

        OraD.ExecuteSQL("Create Index I_" & TT & "_1 on " & TT & "(ITEM_BRAND_CODE, CUST_CODE)")
        OraD.ExecuteSQL("Create Index I_" & TT & "_2 on " & TT & "(CUST_CODE, ITEM_BRAND_CODE)")

        OraD.ExecuteSQL("Alter Table " & TT & " Add SALES_DIVISION_CODE VARCHAR2(6)")
        OraD.ExecuteSQL("Alter Table " & TT & " Add BUS_UNIT_CODE VARCHAR2(6)")
        OraD.ExecuteSQL("Alter Table " & TT & " Add TRADE_CLASS_CODE VARCHAR2(6)")
        OraD.ExecuteSQL("Alter Table " & TT & " Add MARKET_CODE VARCHAR2(6)")
        OraD.ExecuteSQL("Alter Table " & TT & " Add SREP_CODE VARCHAR2(6)")
        OraD.ExecuteSQL("Alter Table " & TT & " Add REGION_CODE VARCHAR2(6)")
        OraD.ExecuteSQL("Alter Table " & TT & " Add VP_CODE VARCHAR2(2)")

        OraD.ExecuteSQL("Update " & TT & " TT Set SALES_DIVISION_CODE = (Select SALES_DIVISION_CODE from ICTBRAN1 where ITEM_BRAND_CODE = TT.ITEM_BRAND_CODE)")
        OraD.ExecuteSQL("Update " & TT & " TT Set BUS_UNIT_CODE = (Select BUS_UNIT_CODE from SOTSDIV1 where SALES_DIVISION_CODE = TT.SALES_DIVISION_CODE)")
        OraD.ExecuteSQL("Update " & TT & " TT Set SREP_CODE = (Select SREP_CODE from ARTSREP1 where CUST_CODE = TT.CUST_CODE and BUS_UNIT_CODE = TT.BUS_UNIT_CODE)")
        OraD.ExecuteSQL("Update " & TT & " TT Set REGION_CODE = (Select REGION_CODE from SOTSREP1 where SREP_CODE = TT.SREP_CODE)")
        OraD.ExecuteSQL("Update " & TT & " TT Set VP_CODE = (Select VP_CODE from SOTSREG1 where REGION_CODE = TT.REGION_CODE)")
        OraD.ExecuteSQL("Update " & TT & " TT Set TRADE_CLASS_CODE = (Select TRADE_CLASS_CODE from ARTCUST1 where CUST_CODE = TT.CUST_CODE)")
        OraD.ExecuteSQL("Update " & TT & " TT Set MARKET_CODE = (Select MARKET_CODE from SOTTCLS1 where TRADE_CLASS_CODE = TT.TRADE_CLASS_CODE)")

        OraD.ExecuteSQL("Delete from " & TT & " where BUS_UNIT_CODE not in ('LUX','PRES')")
        'OraD.ExecuteSQL "Delete from " & TT & " where MARKET_CODE <> 'DPT'"


        Sql = " DELETE FROM " & TT & " WHERE "
        Sql = Sql & " TY_MTD_GRS = 0 AND"
        Sql = Sql & " TY_MTD_RTN = 0 AND"
        Sql = Sql & " TY_YTD_GRS = 0 AND"
        Sql = Sql & " TY_YTD_RTN = 0 AND"
        Sql = Sql & " LY_MTD_GRS = 0 AND"
        Sql = Sql & " LY_MTD_RTN = 0 AND"
        Sql = Sql & " LY_YTD_GRS = 0 AND"
        Sql = Sql & " LY_YTD_RTN = 0 AND"
        Sql = Sql & " TY_MTD_BDG = 0 AND"
        Sql = Sql & " TY_YTD_BDG = 0 AND"
        Sql = Sql & " TY_MTD_BDR = 0 AND"
        Sql = Sql & " TY_YTD_BDR = 0 AND"
        Sql = Sql & " TY_DLY_BKG = 0 AND"
        Sql = Sql & " TY_MTD_BKG = 0 AND"
        Sql = Sql & " TY_YTD_BKG = 0 AND"
        Sql = Sql & " TY_OTS_REL = 0 AND"
        Sql = Sql & " TY_OTS_CUR = 0 AND"
        Sql = Sql & " TY_OTS_FUT = 0"
        OraD.ExecuteSQL(Sql)

        'If UserID = "wjz" Then Stop
        ' TEMPORARY
        'OraD.ExecuteSQL "Delete from " & TT & " where SALES_DIVISION_CODE not in ('BUR','ESC','BOSS','HB')"
        'OraD.ExecuteSQL "Delete from " & TT & " where CUST_CODE not in ('SAKS','BLOOMIES','LORD','DILLARDS')"

        OraD.ExecuteSQL("Truncate Table SATSLSF1")
        OraD.ExecuteSQL("Insert into SATSLSF1 Select * from " & TT)

    End Sub

    Sub Generate_Flash_XLS_1()
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim z As String
        Dim zzz As String
        Dim Count As Integer

        XLDATE = Now + NowTSD

        Dim tb As Integer

        Dim dynTBL As Recordset

        'Screen.MousePointer = 11

        'On Error GoTo Excel_Error:

        Dim objApp As excel.Application
        Dim objBook As excel.Workbook
        Dim objSheet As excel.Worksheet

        Dim SS As String
        SS = App.Path & "\Daily Flash Summary - All.xls"

        On Error Resume Next
        Kill(SS)
        '    On Error GoTo Retry
        objApp = New excel.Application
        On Error GoTo 0
        objBook = objApp.Workbooks.Add

        objApp.DisplayAlerts = False
        For i = objBook.Worksheets.Count To 2 Step -1
            objBook.Worksheets(i).Delete()
        Next i

        r0 = 7
        c0 = 18

        Dim fmt() As String
        ReDim fmt(c0)
        For i = 1 To c0
            fmt(i) = "#,##0"
        Next i
        fmt(3) = "0.0%"
        fmt(5) = "0.0%"
        fmt(8) = "0.0%"
        fmt(14) = "0.0%"
        fmt(16) = "0.0%"

        Dim EXLRPT As String
        EXLRPT = "B"
        EXLRPT = "E"

        Dim RPTLVL() As String
        If EXLRPT = "A" Then
            LEVELS = 4
            ReDim RPTLVL(LEVELS, 3)
            i = 1 : RPTLVL(i, 0) = "SOWSVPS1" : RPTLVL(i, 1) = "VP_CODE" : RPTLVL(i, 2) = "VP_NAME" : RPTLVL(i, 3) = "VP"
            i = 2 : RPTLVL(i, 0) = "SOWSDIV1" : RPTLVL(i, 1) = "SALES_DIVISION_CODE" : RPTLVL(i, 2) = "SALES_DIVISION_NAME" : RPTLVL(i, 3) = "House"
            i = 3 : RPTLVL(i, 0) = "ICWBRAN1" : RPTLVL(i, 1) = "ITEM_BRAND_CODE" : RPTLVL(i, 2) = "ITEM_BRAND_NAME" : RPTLVL(i, 3) = "Brand"
            i = 4 : RPTLVL(i, 0) = "ARWCUST1" : RPTLVL(i, 1) = "CUST_CODE" : RPTLVL(i, 2) = "CUST_NAME" : RPTLVL(i, 3) = "Acct"
        ElseIf EXLRPT = "B" Then
            LEVELS = 3
            ReDim RPTLVL(LEVELS, 3)

            i = 1 : RPTLVL(i, 0) = "SOWSDIV1" : RPTLVL(i, 1) = "SALES_DIVISION_CODE" : RPTLVL(i, 2) = "SALES_DIVISION_NAME" : RPTLVL(i, 3) = "House"
            i = 2 : RPTLVL(i, 0) = "ICWBRAN1" : RPTLVL(i, 1) = "ITEM_BRAND_CODE" : RPTLVL(i, 2) = "ITEM_BRAND_NAME" : RPTLVL(i, 3) = "Brand"
            i = 3 : RPTLVL(i, 0) = "ARWCUST1" : RPTLVL(i, 1) = "CUST_CODE" : RPTLVL(i, 2) = "CUST_NAME" : RPTLVL(i, 3) = "Acct"
        ElseIf EXLRPT = "C" Then
            LEVELS = 2
            ReDim RPTLVL(LEVELS, 3)

            i = 1 : RPTLVL(i, 0) = "SOWSDIV1" : RPTLVL(i, 1) = "SALES_DIVISION_CODE" : RPTLVL(i, 2) = "SALES_DIVISION_NAME" : RPTLVL(i, 3) = "House"
            i = 2 : RPTLVL(i, 0) = "ICWBRAN1" : RPTLVL(i, 1) = "ITEM_BRAND_CODE" : RPTLVL(i, 2) = "ITEM_BRAND_NAME" : RPTLVL(i, 3) = "Brand"
        ElseIf EXLRPT = "D" Then
            LEVELS = 2
            ReDim RPTLVL(LEVELS, 3)

            i = 1 : RPTLVL(i, 0) = "ARWCUST1" : RPTLVL(i, 1) = "CUST_CODE" : RPTLVL(i, 2) = "CUST_NAME" : RPTLVL(i, 3) = "Acct"
            i = 2 : RPTLVL(i, 0) = "SOWSDIV1" : RPTLVL(i, 1) = "SALES_DIVISION_CODE" : RPTLVL(i, 2) = "SALES_DIVISION_NAME" : RPTLVL(i, 3) = "House"
        ElseIf EXLRPT = "E" Then
            LEVELS = 3
            ReDim RPTLVL(LEVELS, 3)

            i = 1 : RPTLVL(i, 0) = "SOWBUSU1" : RPTLVL(i, 1) = "BUS_UNIT_CODE" : RPTLVL(i, 2) = "BUS_UNIT_DESC" : RPTLVL(i, 3) = "Division"
            i = 2 : RPTLVL(i, 0) = "SOWSDIV1" : RPTLVL(i, 1) = "SALES_DIVISION_CODE" : RPTLVL(i, 2) = "SALES_DIVISION_NAME" : RPTLVL(i, 3) = "Brand"
            i = 3 : RPTLVL(i, 0) = "ICWBRAN1" : RPTLVL(i, 1) = "ITEM_BRAND_CODE" : RPTLVL(i, 2) = "ITEM_BRAND_NAME" : RPTLVL(i, 3) = "Sub-Brand"
        End If

        Dim gby As String
        For i = 1 To LEVELS
            gby = gby & ", SAWSLSF1." & RPTLVL(i, 1)
        Next i
        gby = Mid$(gby, 3)

        Dim CHDG() As String
        ReDim CHDG((LEVELS + 1))
        For i = 1 To LEVELS
            CHDG(i) = RPTLVL(i, 3)
        Next i
        CHDG(LEVELS + 1) = "Name"

        Dim f() As String
        ReDim f(c0)
        f(2) = "=+" & XC(6) & "+" & XC(17)
        f(3) = "=IF(" & XC(7) & "=0,0," & XC(2) & "/" & XC(7) & ")"
        f(5) = "=IF(" & XC(4) & "=0,0," & XC(6) & "/" & XC(4) & ")"
        f(8) = "=IF(" & XC(7) & "=0,0,(" & XC(7) & "-" & XC(6) & ")/" & XC(7) & ")"
        f(12) = "=+" & XC(6) & "-" & Excel_Cell(r0, 9 + (LEVELS + 1))
        f(14) = "=IF(" & XC(13) & "=0,0,(" & XC(13) & "-" & XC(12) & ")/" & XC(13) & ")"
        f(15) = "=+" & XC(4) & "-" & XC(11)
        f(16) = "=IF(" & XC(15) & "=0,0,(" & XC(15) & "-" & XC(12) & ")/" & XC(15) & ")"

        Dim TY As String
        TY = Mid$(RYP, 1, 4)
        Dim LY As String
        LY = Mid$(Period_Calc(RYP, -12), 1, 4)
        Dim TYM As String
        TYM = Mid$(RYM, 1, 4)
        Dim LYM As String
        LYM = Mid$(Period_Calc(RYM, -12), 1, 4)
        Dim BD As String
        'BD = Mid$(LYM, 3, 2) & "/" & Mid$(TYM, 3, 2) & vbLf & "BPE"

        Dim TTT(2) As String
        TTT(0) = Format$(Val(Mid$(RYP, 3, 2) - 2), "00")
        TTT(1) = Format$(Val(Mid$(RYP, 3, 2) - 1), "00")
        TTT(2) = Format$(Val(Mid$(RYP, 3, 2) + 0), "00")
        Dim LYT As String
        LYT = TTT(0) & "/" & TTT(1)
        Dim TYT As String
        TYT = TTT(1) & "/" & TTT(2)
        BD = TYT & vbLf & "BPE"

        Dim S As Integer
        For S = 1 To 6

            If S > 1 Then
                objBook.Worksheets.Add(after:=objBook.Worksheets(objBook.Worksheets.Count))
            End If

            objSheet = objBook.Worksheets(S)

            Dim XTD As String
            If S < 4 Then
                XTD = "M"
            Else
                XTD = "Y"
            End If

            If S = 1 Or S = 4 Then
                objSheet.NAME = XTD & "TD Total"
            ElseIf S = 2 Or S = 5 Then
                objSheet.NAME = XTD & "TD Dept"
            ElseIf S = 3 Or S = 6 Then
                objSheet.NAME = XTD & "TD Other"
            End If

            Dim DHDG() As String
            ReDim DHDG(c0)
            DHDG(1) = "Daily" & vbLf & "Sales"
            DHDG(2) = XTD & "TD" & vbLf & "Proj"
            DHDG(3) = "%GS vs" & vbLf & "BPE"
            DHDG(4) = LYT & vbLf & " Act"
            DHDG(5) = XTD & "TD vs " & vbLf & LYT & " Act"
            DHDG(6) = TYT & vbLf & " MTD"
            DHDG(7) = BD
            DHDG(8) = "%BPE"
            DHDG(9) = TYT & vbLf & " MTD"
            DHDG(10) = BD
            DHDG(11) = LYT & vbLf & " MTD"
            DHDG(12) = TYT & vbLf & " MTD"
            DHDG(13) = BD
            DHDG(14) = "%BPE"
            DHDG(15) = LYT & vbLf & " Act"
            DHDG(16) = "%" & TYT & vbLf & " vs " & LYT
            DHDG(17) = "Curr Mo"
            DHDG(18) = "Next Mo"

            '    GoSub General_Sheet_Formatting

            On Error Resume Next
            AccD.Execute("Drop Table SAWSLSFX")
            On Error GoTo 0

            z = " / 1000"
            z = ""

            Sql = "SELECT " & gby
            Sql = Sql & " , '" & Space$(60) & "' AS NAME"
            Sql = Sql & " , SUM (SAWSLSF1.TY_DLY_BKG" & z & ") AS TY_TDY_BOOK"
            Sql = Sql & " , 0 AS TY_PROJ"
            Sql = Sql & " , 0 AS TY_ACT_G_VS_BUD_G"
            Sql = Sql & " , SUM (SAWSLSF1.LY_" & XTD & "TD_GRS" & z & ") AS LY_ACT_G"
            Sql = Sql & " , 0 AS TY_ACT_G_VS_LY_ACT_G"

            Sql = Sql & " , SUM (SAWSLSF1.TY_" & XTD & "TD_GRS" & z & ") AS TY_ACT_G"
            Sql = Sql & " , SUM (SAWSLSF1.TY_" & XTD & "TD_BDG" & z & ") AS TY_BUD_G"
            Sql = Sql & " , 0 AS TY_ACT_G_PCT"
            Sql = Sql & " , SUM (SAWSLSF1.TY_" & XTD & "TD_RTN" & z & ") AS TY_ACT_R"
            Sql = Sql & " , SUM (SAWSLSF1.TY_" & XTD & "TD_BDR" & z & ") AS TY_BUD_R"
            Sql = Sql & " , SUM (SAWSLSF1.LY_" & XTD & "TD_RTN" & z & ") AS LY_ACT_R"

            Sql = Sql & " , 0 AS TY_ACT_N"
            Sql = Sql & " , 0 AS TY_BUD_N"
            Sql = Sql & " , 0 AS TY_ACT_N_PCT"
            Sql = Sql & " , 0 AS LY_ACT_N"
            Sql = Sql & " , 0 AS TY_ACT_N_VS_LY_PCT"

            Sql = Sql & " , SUM ((SAWSLSF1.TY_OTS_CUR + SAWSLSF1.TY_OTS_REL)" & z & ") AS TY_OPEN_CMO"
            Sql = Sql & " , SUM (SAWSLSF1.TY_OTS_FUT" & z & ") AS TY_OPEN_NMO"
            Sql = Sql & " INTO SAWSLSFX FROM SAWSLSF1 "
            'Sql = Sql & " WHERE XTD = '" & XTD & "'"
            If S = 2 Or S = 5 Then
                Sql = Sql & " where MARKET_CODE = 'DPT'"
            ElseIf S = 3 Or S = 6 Then
                Sql = Sql & " where MARKET_CODE <> 'DPT'"
            End If
            'Sql = Sql & " and SALES_DIVISION_CODE in ('ESC','BUR') and ITEM_BRAND_CODE in ('EPH','MAGNET','BRIT','LONMEN')"
            Sql = Sql & " GROUP BY " & gby
            AccD.Execute(Sql)

            AccD.Execute("Update SAWSLSFX," & RPTLVL(LEVELS, 0) & " set SAWSLSFX.NAME = " & RPTLVL(LEVELS, 0) & "." & RPTLVL(LEVELS, 2) & " where " & RPTLVL(LEVELS, 0) & "." & RPTLVL(LEVELS, 1) & " = SAWSLSFX." & RPTLVL(LEVELS, 1) & "")

            Sql = "Select * from SAWSLSFX order by " & Replace$(gby, "SAWSLSF1", "SAWSLSFX")
            dynTBL = AccD.OpenRecordset(Sql, dbOpenDynaset)


            Call Copy_From_Recordset(objApp, objSheet, dynTBL, r0)

            Dim rn As Integer
            rn = dynTBL.RecordCount

            objSheet.Range(Excel_Cell(r0 - 1, LEVELS + 1 + 1) & ":" & Excel_Cell(r0 - 1, LEVELS + 1 + c0)).HorizontalAlignment = xlRight
            objSheet.Range(Excel_Cell(r0 - 1, LEVELS + 1 + 1) & ":" & Excel_Cell(r0 - 1, LEVELS + 1 + c0)).NumberFormat = "@"
            objSheet.Rows(CStr(r0 - 1) & ":" & CStr(r0 - 1)).RowHeight = 25.8

            For i = 0 To dynTBL.Fields.Count - 1
                z = Display_Field_Name(dynTBL.Fields(i).NAME)
                If i + 1 > (LEVELS + 1) Then
                    If DHDG(i + 1 - (LEVELS + 1)) <> "" Then
                        z = DHDG(i + 1 - (LEVELS + 1))
                    End If
                Else
                    If CHDG(i + 1) <> "" Then
                        z = CHDG(i + 1)
                    End If
                End If
                objSheet.Cells(r0 - 1, i + 1) = z
            Next i

            dynTBL.Close()


            Dim lb() As String
            ReDim lb(c0)
            lb(5) = "1"
            lb(8) = "1"
            lb(11) = "1"
            lb(16) = "1"
            '    lb(17) = "1"
            lb(18) = "1"


            For i = 1 To (LEVELS + 1)
                objSheet.Columns(i).NumberFormat = "@"
                objSheet.Columns(i).ColumnWidth = 8
            Next i
            objSheet.Columns(LEVELS + 1).ColumnWidth = 15


            objSheet.Range(Excel_Cell(r0 - 1, 1) & ":" & Excel_Cell(r0 - 1, (LEVELS + 1))).Select()
            objApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            objApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            objApp.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            With objApp.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With objApp.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            objApp.Selection.Borders(xlEdgeRight).LineStyle = xlNone
            objApp.Selection.Borders(xlInsideVertical).LineStyle = xlNone
            With objApp.Selection.Interior
                .ColorIndex = 19
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
            End With
            'ActiveWindow.SmallScroll ToRight:=4
            objSheet.Range(Excel_Cell(r0 - 1, (LEVELS + 1) + 1) & ":" & Excel_Cell(r0 - 1, (LEVELS + 1) + c0)).Select()
            objApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            objApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            objApp.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            With objApp.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With objApp.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            objApp.Selection.Borders(xlEdgeRight).LineStyle = xlNone
            objApp.Selection.Borders(xlInsideVertical).LineStyle = xlNone
            With objApp.Selection.Interior
                .ColorIndex = 25
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
            End With
            objApp.Selection.Font.ColorIndex = 2

            Dim TBL() As Recordset
            ReDim TBL(LEVELS)
            For i = 1 To LEVELS - 1
                TBL(i) = AccD.OpenRecordset(RPTLVL(i, 0), dbOpenTable)
                TBL(i).Index = "PrimaryKey"
            Next i

            'objApp.DisplayFullScreen = True
            'objApp.DisplayFullScreen = False
            'objApp.UserControl = True
            'objApp.Visible = True

            Dim r As Integer
            r = Build_Groups(objApp, objSheet, TBL, rn, 1)


            For i = 1 To c0
                objSheet.Columns(i + (LEVELS + 1)).NumberFormat = fmt(i)
                objSheet.Columns(i + (LEVELS + 1)).ColumnWidth = 8
            Next i

            For i = 0 To c0
                If i = 0 Or lb(i) = "1" Then
                    objSheet.Range(Excel_Cell(r0 - 1, (LEVELS + 1) + i) & ":" & Excel_Cell(r, (LEVELS + 1) + i)).Select()
                    objApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                    objApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                    objApp.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
                    With objApp.Selection.Borders(xlEdgeTop)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                    objApp.Selection.Borders(xlEdgeBottom).LineStyle = xlNone
                    With objApp.Selection.Borders(xlEdgeRight)
                        .LineStyle = xlContinuous
                        .Weight = xlThin
                        .ColorIndex = xlAutomatic
                    End With
                End If
            Next i

            For i = 1 To c0
                If f(i) <> "" Then
                    objSheet.Cells(r0 + 1, i + (LEVELS + 1)) = f(i)
                    objSheet.Cells(r0 + 1, i + (LEVELS + 1)).Select()
                    objApp.Selection.Copy()
                    objSheet.Range(Excel_Cell(r0 + 2, i + (LEVELS + 1)) & ":" & Excel_Cell(r, i + (LEVELS + 1))).Select()
                    'objSheet.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                    objApp.Selection.PasteSpecial(Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False)
                    'objSheet.Paste
                End If
            Next i

            ' Format the Totals Line with Borders & Shading

            objSheet.Range(Excel_Cell(r, 1) & ":" & Excel_Cell(r, LEVELS + 1 + c0)).Select()
            objApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            objApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            objApp.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            With objApp.Selection.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With objApp.Selection.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With objApp.Selection.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With objApp.Selection.Interior
                .ColorIndex = 24
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
            End With


            For i = 1 To 5
                Select Case i
                    Case 1
                        z = Excel_Cell(r0 - 2, (LEVELS + 1) + 1) & ":" & Excel_Cell(r0 - 2, (LEVELS + 1) + 5)
                    Case 2
                        z = Excel_Cell(r0 - 2, (LEVELS + 1) + 6) & ":" & Excel_Cell(r0 - 2, (LEVELS + 1) + 8)
                    Case 3
                        z = Excel_Cell(r0 - 2, (LEVELS + 1) + 9) & ":" & Excel_Cell(r0 - 2, (LEVELS + 1) + 11)
                    Case 4
                        z = Excel_Cell(r0 - 2, (LEVELS + 1) + 12) & ":" & Excel_Cell(r0 - 2, (LEVELS + 1) + 16)
                    Case 5
                        z = Excel_Cell(r0 - 2, (LEVELS + 1) + 17) & ":" & Excel_Cell(r0 - 2, (LEVELS + 1) + 18)
                End Select
                objSheet.Range(z).Select()
                With objApp.Selection
                    .HorizontalAlignment = xlGeneral
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
                objApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
                objApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
                With objApp.Selection.Borders(xlEdgeLeft)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With objApp.Selection.Borders(xlEdgeTop)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With objApp.Selection.Borders(xlEdgeBottom)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                With objApp.Selection.Borders(xlEdgeRight)
                    .LineStyle = xlContinuous
                    .Weight = xlThin
                    .ColorIndex = xlAutomatic
                End With
                objApp.Selection.Borders(xlInsideVertical).LineStyle = xlNone
                With objApp.Selection.Interior
                    .ColorIndex = 24
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                End With
                objSheet.Range(z).Select()
                Select Case i
                    Case 1
                        'objSheet.Cells(5, 6) = "Bookings"
                        objSheet.Range(z) = "Bookings"
                        'ActiveCell.FormulaR1C1 = "Bookings"
                    Case 2
                        objSheet.Range(z) = "Gross Sales"
                    Case 3
                        objSheet.Range(z) = "Returns"
                    Case 4
                        objSheet.Range(z) = "Net Shipments"
                    Case 5
                        objSheet.Range(z) = "Open to Ship"
                End Select
                objSheet.Range(z).Select()
                With objApp.Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .IndentLevel = 0
                    .ShrinkToFit = False
                    .ReadingOrder = xlContext
                    .MergeCells = True
                End With
            Next i

            objSheet.Cells(1, 1) = InstName
            objSheet.Cells(2, 1) = "Daily Net Shipments Analysis Report as of " & CStr(XLDATE)
            objSheet.Cells(3, 1) = objSheet.NAME

            objSheet.Cells(5, 1) = "in thousands"

            'objApp.Sheets(1).Select
            For i = LEVELS To 1 Step -1
                objSheet.Outline.ShowLevels(rowlevels:=i)
            Next i

            objSheet.Range("A1:A1").Select()

        Next S

        objBook.Worksheets(1).Activate()
        'objBook.Sheets(1).Range("A1:A1").Select

        'objWB.Sheets(1).Activate
        'objWB.Sheets(1).Cells(1, 1).Select

        On Error Resume Next
        objBook.SaveAs(SS)
        On Error GoTo 0

        objApp.DisplayAlerts = True

        objApp.DisplayFullScreen = True
        objApp.DisplayFullScreen = False
        objApp.UserControl = True
        objApp.Visible = True

    End Sub

    Sub Generate_Flash_XLS_2()

        Dim z As String
        Dim zzz As String
        Dim i As Integer
        Dim j As Integer

        Dim objApp As excel.Application
        Dim objBook As excel.Workbook
        Dim objSheet As excel.Worksheet

    GoSub Initialize_Excel

        Dim ii As Integer
        ii = 0

        Dim MCOLS As Integer

        Dim COLS As Integer
        COLS = 9
        Dim ICOLS As Integer
        ICOLS = 2

        Dim BUS_UNIT_CODE As String
        Dim SALES_DIVISION_CODE As String

        BUS_UNIT_CODE = ""
        SALES_DIVISION_CODE = ""
    GoSub Create_Sheet

        Sql = "Select DISTINCT BUS_UNIT_CODE, SALES_DIVISION_CODE "
        Sql = Sql & " from SAWSLSF1 "
        Sql = Sql & " where MARKET_CODE = 'DPT'"
        Sql = Sql & " order by BUS_UNIT_CODE, SALES_DIVISION_CODE"
        Dim dynBD As Recordset
        dynBD = AccD.OpenRecordset(Sql, dbOpenForwardOnly)
        Do While Not dynBD.EOF
            If BUS_UNIT_CODE <> dynBD.Fields("BUS_UNIT_CODE").Value & "" Then
                BUS_UNIT_CODE = dynBD.Fields("BUS_UNIT_CODE").Value
                SALES_DIVISION_CODE = ""
            GoSub Create_Sheet
            End If
            SALES_DIVISION_CODE = dynBD.Fields("SALES_DIVISION_CODE").Value
        GoSub Create_Sheet
            dynBD.MoveNext()
        Loop
        dynBD.Close()


        '        If ii = 1 Then
        '            objSheet.Hyperlinks.Add objSheet.Range("A2"), "http://example.microsoft.com"
        '            objSheet.Hyperlinks.Add objSheet.Range("A3"), "$'Summary by SKU'", "A193", "hello"
        '            objSheet.Hyperlinks.Add objSheet.Range("A4"), "", "'Summary by SKU'!A193"
        '            objSheet.Hyperlinks.Add objSheet.Range("A4"), "", "'Summary by Store'!A193"
        '        End If

        objBook.Worksheets(1).Activate()
        objBook.Sheets(1).Range("A1:A1").Select()
        'objBook.Sheets(1).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowFormattingColumns:=True


        Dim SS As String
        SS = App.Path & "\Bookings & Flash by Account by Brand.xls"

        On Error Resume Next
        Kill(SS)
        objBook.SaveAs(SS)

        objApp.DisplayAlerts = True
        objApp.UserControl = True
        objApp.Visible = True

Excel_Termination:
        On Error GoTo 0
        DoEvents()
        objSheet = Nothing
        objBook = Nothing
        objApp = Nothing
        Exit Sub

Initialize_Excel:

        objApp = New excel.Application
        objBook = objApp.Workbooks.Add

        objApp.DisplayAlerts = False
        For i = objBook.Worksheets.Count To 2 Step -1
            objBook.Worksheets(i).Delete()
        Next i
        Return

Create_Sheet:
        ii = ii + 1
    GoSub Add_Sheet

        Dim Sheet_Name As String
        Dim imax As Integer
        Dim iCODE As String
        Dim wCODE As String
        wCODE = " where MARKET_CODE = 'DPT'"

        If BUS_UNIT_CODE = "" Then
            iCODE = "BUS_UNIT_CODE"
            wCODE = wCODE & ""
            Sheet_Name = "Company"
            objSheet.Tab.ColorIndex = 11
        ElseIf SALES_DIVISION_CODE = "" Then
            iCODE = "SALES_DIVISION_CODE"
            wCODE = wCODE & " and BUS_UNIT_CODE = '" & BUS_UNIT_CODE & "'"
            Sql = Sql & " order by SALES_DIVISION_CODE"
            Sheet_Name = BUS_UNIT_CODE
            'objSheet.Tab.ColorIndex = 13
            objSheet.Tab.ColorIndex = 10
        Else
            iCODE = "ITEM_BRAND_CODE"
            wCODE = wCODE & " and SALES_DIVISION_CODE = '" & SALES_DIVISION_CODE & "'"
            Sheet_Name = SALES_DIVISION_CODE
            objSheet.Tab.ColorIndex = 9 ' or 46
        End If

        Call Prompt("Now Creating Sheets", Sheet_Name)

        Sql = "Select Distinct " & iCODE
        Sql = Sql & " from SAWSLSF1"
        Sql = Sql & wCODE
        Sql = Sql & " order by " & iCODE

        'GoSub Format_Columns2

        Dim MR As Integer
        Dim MC As Integer
        Dim MCOLOR As Integer
        Dim MBOLD As Boolean

        MCOLOR = 51
        MCOLOR = 10
        MCOLOR = 46
        MCOLS = COLS
        MR = 2
        Dim dynmax As Recordset
        dynmax = AccD.OpenRecordset(Sql, dbOpenForwardOnly)
        imax = 0
        Dim iCODEs() As String
        Do While Not dynmax.EOF
            imax = imax + 1
            ReDim Preserve iCODEs(imax)
            iCODEs(imax) = dynmax.Fields(0).Value & ""
            dynmax.MoveNext()
        Loop
        dynmax.Close()

        MCOLS = 3

        MC = 3 + (imax) * COLS
        objSheet.Cells(MR, MC).Value = "Total"
        MBOLD = True
    GoSub Merge_Column

        For i = 1 To imax
            MC = 3 + (i - 1) * COLS
            objSheet.Cells(MR, MC).Value = iCODEs(i)
            MBOLD = True
        GoSub Merge_Column
        Next i

        objSheet.Cells(MR, 1).Value = "In Thousands"
        objSheet.NAME = Sheet_Name
        'GoSub Format_Columns

        Dim MY As Integer
        Dim XTD As String

        Dim iCT As Integer
        Dim RI As Integer
        Dim li As Integer
        Dim headings_done As Boolean

        Dim CODE_TYPE(2, 2) As String
        CODE_TYPE(1, 1) = "TRADE_CLASS_CODE"
        CODE_TYPE(1, 2) = "CUST_CODE"
        CODE_TYPE(2, 1) = "REGION_CODE"
        CODE_TYPE(2, 2) = "SREP_CODE"

        Dim r As Integer
        r = 0

        Dim CR As Integer
        CR = MR
        headings_done = False

        For iCT = 1 To 2

            For MY = 0 To 1
                If MY = 0 Then
                    XTD = "MTD"
                Else
                    XTD = "YTD"
                End If

            GoSub Set_Headings

                z = CODE_TYPE(iCT, 1) : z = Mid$(z, 1, Len(z) - 5)
                objSheet.Cells(CR, 1).Value = Display_Field_Name(z)
                z = CODE_TYPE(iCT, 2) : z = Mid$(z, 1, Len(z) - 5)
                objSheet.Cells(CR, 2).Value = Display_Field_Name(z)

                If iCT = 1 Then
                    MCOLOR = 5
                Else
                    MCOLOR = 13
                End If
                objSheet.Range(Excel_Cell(CR, 1) & ":" & Excel_Cell(CR, 2)).Interior.ColorIndex = MCOLOR
                objSheet.Range(Excel_Cell(CR, 1) & ":" & Excel_Cell(CR, 2)).Font.ColorIndex = 2

                '            For i = 1 To imax
                '                objSheet.Cells(ICOLS + R + MY * 5 + iCT * 3, 3 + (imax - 1) * 5).Value = xTD
                '            Next i

                Sql = "Select SAWSLSF1." & CODE_TYPE(iCT, 1) & ", SAWSLSF1." & CODE_TYPE(iCT, 2)

                For i = 1 To imax
                    Sql = Sql & ", Sum (IIf(SAWSLSF1." & iCODE & " = '" & iCODEs(i) & "',SAWSLSF1.TY_" & XTD & "_GRS,0)) as TY_" & XTD & "_GRS_" & iCODEs(i)
                    Sql = Sql & ", Sum (IIf(SAWSLSF1." & iCODE & " = '" & iCODEs(i) & "',SAWSLSF1.TY_" & XTD & "_RTN,0)) as TY_" & XTD & "_RTN_" & iCODEs(i)
                    Sql = Sql & ", NULL as TY_" & XTD & "_NET_" & iCODEs(i)
                    Sql = Sql & ", NULL as TY_" & XTD & "_PCT_BPE_" & iCODEs(i)
                    Sql = Sql & ", NULL as TY_" & XTD & "_PCT_PY_" & iCODEs(i)
                    Sql = Sql & ", NULL as TY_" & XTD & "_PCT_NS_" & iCODEs(i)
                    If MY = 1 Then
                        Sql = Sql & ", NULL as TY_" & XTD & "_BPE_" & iCODEs(i)
                    Else
                        Sql = Sql & ", NULL as TY_" & XTD & "_BPE_" & iCODEs(i)
                    End If
                    If MY = 1 Then
                        Sql = Sql & ", Sum (IIf(SAWSLSF1." & iCODE & " = '" & iCODEs(i) & "',SAWSLSF1.LY_" & XTD & "_GRS,0)) as LY_" & XTD & "_" & iCODEs(i)
                        Sql = Sql & ", NULL as TY_" & XTD & "_BPE_TOGO_" & iCODEs(i)
                    Else
                        Sql = Sql & ", NULL as TY_" & XTD & "_BKG_" & iCODEs(i)
                        Sql = Sql & ", NULL as TY_" & XTD & "_OTS_" & iCODEs(i)
                    End If
                Next i

                Sql = Sql & ", Sum (SAWSLSF1.TY_" & XTD & "_GRS) as TY_" & XTD & "_GRS"
                Sql = Sql & ", Sum (SAWSLSF1.TY_" & XTD & "_RTN) as TY_" & XTD & "_RTN"
                Sql = Sql & ", NULL as TY_" & XTD & "_NET"
                Sql = Sql & ", NULL as TY_" & XTD & "_PCT_BPE"
                Sql = Sql & ", NULL as TY_" & XTD & "_PCT_PY"
                Sql = Sql & ", NULL as TY_" & XTD & "_PCT_NS"
                If MY = 1 Then
                    Sql = Sql & ", Sum (SAWSLSF1.TY_" & XTD & "_BDG - SAWSLSF1.TY_" & XTD & "_BDR) as TY_" & XTD & "_BPE"
                Else
                    Sql = Sql & ", NULL as TY_" & XTD & "_BPE"
                End If
                If MY = 1 Then
                    Sql = Sql & ", Sum (SAWSLSF1.LY_" & XTD & "_GRS) as LY_" & XTD
                    Sql = Sql & ", NULL as TY_" & XTD & "_BPE_TOGO"
                Else
                    Sql = Sql & ", NULL as TY_" & XTD & "_BKG"
                    Sql = Sql & ", NULL as TY_" & XTD & "_OTS"
                End If

                Sql = Sql & " from SAWSLSF1"
                Sql = Sql & wCODE
                Sql = Sql & " group by SAWSLSF1." & CODE_TYPE(iCT, 1) & ", SAWSLSF1." & CODE_TYPE(iCT, 2)
                Sql = Sql & " order by SAWSLSF1." & CODE_TYPE(iCT, 1) & ", SAWSLSF1." & CODE_TYPE(iCT, 2)

                Dim dynTBL As Recordset
                dynTBL = AccD.OpenRecordset(Sql, dbOpenDynaset)
                'dynTBL.MoveLast
                'R = R + dynTBL.RecordCount
                'dynTBL.MoveFirst
                CR = CR + 1
                objSheet.Range(Excel_Cell(CR, 1)).CopyFromRecordset(dynTBL)
                dynTBL.MoveFirst()
                r = dynTBL.RecordCount
                dynTBL.MoveLast()

                objSheet.Cells(CR, ICOLS + 3).Value = "=+" & Excel_Cell(CR, ICOLS + 1) & "-" & Excel_Cell(CR, ICOLS + 2)
                If MY = 1 Then
                    '=IF(I5=0,0,E5/I5)
                    objSheet.Cells(CR, ICOLS + 4).Value = "=IF(" & Excel_Cell(CR, ICOLS + 7) & "=0,0," & Excel_Cell(CR, ICOLS + 3) & "/" & Excel_Cell(CR, ICOLS + 7) & ")"
                    objSheet.Cells(CR, ICOLS + 5).Value = "=IF(" & Excel_Cell(CR, ICOLS + 8) & "=0,0," & Excel_Cell(CR, ICOLS + 3) & "/" & Excel_Cell(CR, ICOLS + 8) & ")"
                    objSheet.Cells(CR, ICOLS + 6).Value = "=1-" & Excel_Cell(CR, ICOLS + 4)
                    objSheet.Range(Excel_Cell(CR, ICOLS + 3) & ":" & Excel_Cell(CR, ICOLS + 6)).Select()
                Else
                    objSheet.Range(Excel_Cell(CR, ICOLS + 4)).Interior.ColorIndex = 15
                    objSheet.Range(Excel_Cell(CR, ICOLS + 5)).Interior.ColorIndex = 15
                    objSheet.Range(Excel_Cell(CR, ICOLS + 6)).Interior.ColorIndex = 15
                    objSheet.Range(Excel_Cell(CR, ICOLS + 7)).Interior.ColorIndex = 15
                    objSheet.Range(Excel_Cell(CR, ICOLS + 3) & ":" & Excel_Cell(CR, ICOLS + 7)).Select()
                    'objSheet.Range(Excel_Cell(CR, ICOLS + 3)).Select
                End If
                objApp.Selection.Copy()
                objSheet.Range(Excel_Cell(CR, ICOLS + 3), Excel_Cell(CR + r - 1, ICOLS + 3)).Select()
                objSheet.Paste()
                objApp.Selection.Copy()

                For i = 1 To imax + 1
                    objSheet.Range(Excel_Cell(CR, ICOLS + 3 + (i - 1) * COLS), Excel_Cell(CR + r - 1, ICOLS + 3 + (i - 1) * COLS)).Select()
                    objSheet.Paste()
                Next i


                For i = 1 To imax + 1
                    objSheet.Range(Excel_Cell(CR - 1, ICOLS + (i - 1) * COLS + 1), Excel_Cell(CR + r - 1, ICOLS + (i - 1) * COLS + COLS)).Select()
                    Call Excel_Grid(objApp, False)
                    If MY = 1 Then
                        With objSheet.Range(Excel_Cell(CR - 1, ICOLS + (i - 1) * COLS + 7), Excel_Cell(CR + r - 1, ICOLS + (i - 1) * COLS + COLS))
                            .Interior.ColorIndex = 6
                            .Select()
                            Call Excel_Grid(objApp, False)
                        End With
                    End If
                Next i

                objSheet.Range(Excel_Cell(CR, 1), Excel_Cell(CR + r - 1, 1)).Font.ColorIndex = 10
                objSheet.Range(Excel_Cell(CR, 2), Excel_Cell(CR + r - 1, 2)).Font.ColorIndex = 13
                '    Selection.Font.ColorIndex = 5 ' blue
                '    Selection.Font.ColorIndex = 3 ' red
                '    Selection.Font.ColorIndex = 13 ' purple
                '    Selection.Font.ColorIndex = 53 ' BROWN
                '    Selection.Font.ColorIndex = 7 ' magenta
                '        .ColorIndex = 9 ' burgundy
                '        .ColorIndex = 6 ' yellow
                '         .ColorIndex = 46 ' orange
                '         .ColorIndex = 10 ' GREEN
                '         .ColorIndex = 11 ' DK BLUE
                '         .ColorIndex = 36 ' BEIGE
                '         .ColorIndex = 6 ' YELLOW




            GoSub Sub_Totals

                CR = CR + r + li

                dynTBL.Close()
            Next MY
        Next iCT


        With objApp.ActiveSheet.PageSetup
            .LeftHeader = Format$(Now, "MM/DD/YY")
            .CenterHeader = "&""Arial,Bold""P&G prestige products" & vbLf & "Mtd/Ytd Net Shipment Flash by Account / Brand"
            .RightHeader = "Page &P"
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
        End With

    GoSub Format_Columns2

        objApp.ActiveSheet.Outline.ShowLevels(rowlevels:=1)

        objSheet.Range("C4").Select()
        objApp.ActiveWindow.FreezePanes = True
        objSheet.Range("A1:A1").Select()

        Return

Sub_Totals:
        RI = 0
        li = 0
        dynTBL.MoveFirst()
        Dim st() As String
        ReDim st(ICOLS, 2)
        Do While Not dynTBL.EOF
            Dim IX As Integer
            IX = 0
            For i = 1 To ICOLS
                If st(i, 0) <> dynTBL.Fields(i - 1).Value Then
                    IX = i
                    Exit For
                End If
            Next i

        GoSub Process_Sub_Total
            RI = RI + 1
            dynTBL.MoveNext()
        Loop
    IX = 0: GoSub Process_Sub_Total

    i = 0: GoSub Sub_Totals_Insert
        Return

Process_Sub_Total:
        For i = ICOLS To 1 Step -1
            If i >= IX Then
                ' If there is something to sub-total, then insert a line to hold the subtotal and set its formula
                If st(i, 0) <> "" Then
                    If i < ICOLS Then
                    GoSub Sub_Totals_Insert
                    End If
                    ' register this line as one be totaled 1 level higher
                    st(i - 1, 1) = st(i - 1, 1) & ",^" & CStr(CR + RI - 1)
                End If

                ' Register Code Value for next set of rows to be totaled
                If IX > 0 Then
                    st(i, 0) = dynTBL.Fields(i - 1).Value
                End If

            End If
        Next i
        Return

Sub_Totals_Insert:
        objSheet.Rows(CR + RI - 1).Select()
        objApp.Selection.Copy()
        objSheet.Rows(CR + RI).Select()
        objApp.Selection.Insert(Shift:=xlDown)

        ' Grey background, Lines above and below
        objSheet.Range(Excel_Cell(CR + RI, 1) & ":" & Excel_Cell(CR + RI, ICOLS + (imax + 1) * COLS)).Select()
        With objApp.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Interior
            .ColorIndex = 15
            .PatternColorIndex = xlAutomatic
        End With



        li = li + 1
        RI = RI + 1

        z = Mid$(st(i, 1), 2)
        st(i, 1) = ""
        If i = ICOLS - 1 Then
            Dim zz() As String
            zz = Split(z, ",")
            z = zz(0) & ":" & zz(UBound(zz))
        End If

        For j = 1 To COLS
            If j = 4 Or j = 5 Or j = 6 Or (j = 7 And MY = 0) Then
            Else
                zzz = Replace$(z, "^", Excel_Cell(0, ICOLS + imax * COLS + j))
                objSheet.Cells(CR + RI - 1, ICOLS + imax * COLS + j).Value = "=Sum(" & zzz & ")"
            End If
        Next j
        objSheet.Cells(CR + RI - 1, i + 1).Value = "Totals"
        If i + 1 <> ICOLS Then
            For j = i + 1 + 1 To ICOLS
                objSheet.Cells(CR + RI - 1, j).Value = ""
            Next j
        End If
        objSheet.Range(Excel_Cell(CR + RI - 1, ICOLS + imax * COLS + 1) & ":" & Excel_Cell(CR + RI - 1, ICOLS + imax * COLS + COLS)).Select()
        objApp.Selection.Copy()
        For j = 1 To imax
            objSheet.Range(Excel_Cell(CR + RI - 1, ICOLS + (j - 1) * COLS + 1)).Select()
            objSheet.Paste()
        Next j

        ' Create a Group
        If i <> 0 Then
            zz = Split(Replace(z, "^", "") & ":", ":")
            If zz(1) = "" Then
                zz(1) = zz(0)
            End If
            objSheet.Rows(Replace(Excel_Cell(Val(zz(0)), 0) & ":" & Excel_Cell(Val(zz(1)), 0), "@", "")).Select()
            objApp.Selection.Rows.Group()

            If iCT = 1 And i = 1 And st(i, 0) = "DPT" Then
            Else
                'objApp.Selection.Collapse
            End If
        End If
        Return

Set_Headings:

        ' Column Headings for Numeric Amounts

        'If Not headings_done Then
        CR = CR + 1
        MR = CR
        MC = ICOLS + (imax) * COLS + 1
        objSheet.Range(Excel_Cell(MR, 1)).EntireRow.RowHeight = 27
        objSheet.Cells(MR, MC + 0).Value = "Gross" & vbLf & "Sales"
        objSheet.Cells(MR, MC + 1).Value = "Returns"
        objSheet.Cells(MR, MC + 2).Value = "Net" & vbLf & "Ship"
        If MY = 1 Then
            objSheet.Cells(MR, MC + 3).Value = "BPE" & vbLf & "%"
            objSheet.Cells(MR, MC + 4).Value = "PY" & vbLf & "%"
            objSheet.Cells(MR, MC + 5).Value = "BPE/NS" & vbLf & "Var%"
            objSheet.Cells(MR, MC + 6).Value = "BPE"
        End If
        If MY = 1 Then
            objSheet.Cells(MR, MC + 7).Value = "PY"
            objSheet.Cells(MR, MC + 8).Value = "BPE Yr" & vbLf & "To-Go"
        Else
            objSheet.Cells(MR, MC + 7).Value = "Booked"
            objSheet.Cells(MR, MC + 8).Value = "Open" & vbLf & "to Ship"
        End If

        objSheet.Range(Excel_Cell(MR, MC + 0), Excel_Cell(MR, MC + 8)).Select()
        With objApp.Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .Font.ColorIndex = 2
            '.Interior.ColorIndex = 5
            .Interior.ColorIndex = 11
        End With

        'End If

        ' MTD / YTD Column Headings

        CR = CR + 1
        MR = CR
        MC = ICOLS + (imax) * COLS + 1
        If MY = 1 Then
            MCOLOR = 6
            MCOLS = 6
            objSheet.Cells(MR, MC).Value = "Year-To-Date"
            MBOLD = True
        GoSub Merge_Column
            MCOLOR = 36
            MCOLS = 3
            MC = MC + 6
            objSheet.Cells(MR, MC).Value = "Full-Year"
            MBOLD = True
        GoSub Merge_Column
            MC = MC - 6
        Else
            MCOLOR = 6
            MCOLS = 9
            objSheet.Cells(MR, MC).Value = "Month-To-Date"
            MBOLD = True
        GoSub Merge_Column
        End If
        'objSheet.Cells(MR, MC).HorizontalAlignment = xlCenter

        ' Copy & Paste Headings to each Horizontal unit

        'If headings_done Then
        '    j = MR
        'Else
        j = MR - 1
        'End If

        objSheet.Range(Excel_Cell(j, MC + 0), Excel_Cell(MR, MC + 8)).Select()
        objApp.Selection.Copy()
        For i = 1 To imax
            objSheet.Range(Excel_Cell(j, 3 + (i - 1) * COLS)).Select()
            objSheet.Paste()
        Next i

        headings_done = True

        Return

Merge_Column:

        '        .ColorIndex = 46 ' ORANGE
        '        .ColorIndex = 3 ' RED
        '        .ColorIndex = 5 ' BLUE
        '        .ColorIndex = 15 ' GREY

        objSheet.Range(Excel_Cell(MR, MC + (1 - 1)), Excel_Cell(MR, MC + (MCOLS - 1))).Select()
        '    With objApp.Selection
        '        .HorizontalAlignment = xlGeneral
        '        .VerticalAlignment = xlBottom
        '        .WrapText = False
        '        .Orientation = 0
        '        .AddIndent = False
        '        .IndentLevel = 0
        '        .ShrinkToFit = False
        '        .ReadingOrder = xlContext
        '        .MergeCells = True
        '    End With
        With objApp.Selection.Font
            '        .NAME = "Arial"
            '        .FontStyle = "Regular"
            '        .Size = 10
            '        .Strikethrough = False
            '        .Superscript = False
            '        .Subscript = False
            '        .OutlineFont = False
            '        .Shadow = False
            If MBOLD = True Then
                .FontStyle = "Bold"
                MBOLD = False
            End If
            .Underline = xlUnderlineStyleNone
            If MCOLOR <> 6 And MCOLOR <> 46 And MCOLOR <> 36 Then
                .ColorIndex = 2
            End If
        End With
        '    objApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        '    objApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With objApp.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        objApp.Selection.Borders(xlInsideVertical).LineStyle = xlNone
        With objApp.Selection.Interior
            .ColorIndex = MCOLOR
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
        End With
        With objApp.Selection
            .HorizontalAlignment = xlCenter
            '        .VerticalAlignment = xlBottom
            '        .WrapText = False
            '        .Orientation = 0
            '        .AddIndent = False
            '        .IndentLevel = 0
            '        .ShrinkToFit = False
            '        .ReadingOrder = xlContext
            .MergeCells = True
        End With

        Return

Add_Sheet:
        If ii > 1 Then
            objBook.Worksheets.Add(after:=objBook.Worksheets(objBook.Worksheets.Count))
        End If
        objSheet = objBook.Worksheets(ii)

        With objSheet.Columns("A:B")
            .NumberFormat = "@"
            .ColumnWidth = 10
        End With
        '    objSheet.Rows("4:4").Select
        '    Selection.NumberFormat = "@"

        objSheet.Cells(1, 1).NumberFormat = "MM/DD/YY"
        objSheet.Cells(1, 1) = Now
        objApp.ActiveWindow.Zoom = 70
        Return

Format_Columns:
        For i = 1 To COLS
            With objSheet.Columns(Excel_Cell(0, ICOLS + i))
                If i = 4 Or i = 5 Or i = 6 Then
                    '.NumberFormat = "[Red]###,##0.0%;[Red]-###,##0,0%;0;@" ' "#,##0.0%"
                    .NumberFormat = "###,##0.0%;[Red]-###,##0,0%;0%;@" ' "#,##0.0%"
                Else
                    '.NumberFormat = "[Red]###,##0;[Red]-###,##0;0;@" ' "#,##0"
                    .NumberFormat = "###,##0;[Red]-###,##0;0;@" ' "#,##0"
                End If
                .ColumnWidth = 7
                .HorizontalAlignment = xlRight
            End With
        Next i

        objSheet.Range(Excel_Cell(0, ICOLS + 1) & ":" & Excel_Cell(0, ICOLS + COLS)).Select()
        objApp.Selection.Copy()
        For i = 1 To imax
            objSheet.Range(Excel_Cell(0, ICOLS + (i - 1) * COLS + 1) & ":" & Excel_Cell(0, ICOLS + (i - 1) * COLS + COLS)).Select()
            objSheet.Paste()
        Next i
        Return

Format_Columns2:
        For j = 1 To imax + 1
            For i = 1 To COLS
                With objSheet.Columns(Excel_Cell(0, ICOLS + (j - 1) * COLS + i))
                    If i = 4 Or i = 5 Or i = 6 Then
                        .NumberFormat = "###,##0.0%;[Red]-###,##0,0%;0%;@" ' "#,##0.0%"
                    Else
                        .NumberFormat = "###,##0;[Red]-###,##0;0;@" ' "#,##0"
                    End If
                    .ColumnWidth = 7
                    .HorizontalAlignment = xlRight
                End With
            Next i
        Next j
        Return

    End Sub

    Sub Excel_Grid(objApp As excel.Application, inner_grid As Boolean)

        objApp.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        objApp.Selection.Borders(xlDiagonalUp).LineStyle = xlNone

        With objApp.Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With objApp.Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        If inner_grid Then
            With objApp.Selection.Borders(xlInsideVertical)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
            With objApp.Selection.Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous
                .Weight = xlThin
                .ColorIndex = xlAutomatic
            End With
        End If
    End Sub

    Function XC(i As Integer) As String
        XC = Excel_Cell(r0 + 1, i + (LEVELS + 1))
    End Function
    Function TransposeDim(v As Object) As Object
        Dim X As Long
        Dim Y As Long
        Dim XUpper As Long
        Dim YUpper As Long
        Dim TempArray As Object

        XUpper = UBound(v, 2)
        YUpper = UBound(v, 1)

        ReDim TempArray(XUpper, YUpper)
        For X = 0 To XUpper
            For Y = 0 To YUpper
                TempArray(X, Y) = v(Y, X)
            Next Y
        Next X

        TransposeDim = TempArray
    End Function

    Sub Copy_From_Recordset(ByRef objApp As excel.Application, ByRef objSheet As excel.Worksheet, dynwk As Recordset, Row0 As Integer)
        Dim i As Integer
        Dim j As Integer

        If Val(Mid$(objApp.Version, 1, InStr(1, objApp.Version, ".") - 1)) > 8 Then
            ' Excel 2000 or later
            objSheet.Range("A" & CStr(Row0 + 1)).CopyFromRecordset(dynwk)
        Else
            ' Excel 97 or earlier
            Dim RecArray As Object
            Dim recCount As Integer

            RecArray = dynwk.GetRows
            Do While Not dynwk.EOF
                ReDim Preserve RecArray(UBound(RecArray, 1), UBound(RecArray, 2) + 1)
                RecArray(UBound(RecArray, 1), UBound(RecArray, 2)) = dynwk.GetRows
                dynwk.MoveNext()
            Loop
            recCount = UBound(RecArray, 2) + 1

            For i = 1 To dynwk.Fields.Count
                For j = 0 To recCount - 1
                    If IsDate(RecArray(i - 1, j)) Then
                        RecArray(i - 1, j) = Format$(RecArray(i - 1, j))
                    ElseIf IsArray(RecArray(i - 1, j)) Then
                        RecArray(i - 1, j) = "Data cannot be displayed"
                    End If
                Next j
            Next i

            objSheet.Range("A" & CStr(Row0 + 1)).resize(recCount, dynwk.Fields.Count).Value = TransposeDim(RecArray)
        End If
    End Sub


    Function Build_Groups(ByRef objApp As excel.Application, ByRef objSheet As excel.Worksheet, ByRef TBL() As Recordset, rn As Integer, Groups As Integer) As Integer
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim z As String

        Dim CLR(4) As Integer
        CLR(1) = 10
        CLR(2) = 13
        CLR(3) = 3
        CLR(4) = 0 ' 17
        ' 4 6 7 8
        '19 20 24 26
        '10 13 3 17

        For j = 1 To LEVELS
            objSheet.Range(Excel_Cell(r0 + 1, j) & ":" & Excel_Cell(r0 + rn, j)).Font.ColorIndex = CLR(j)
        Next j

        Dim st() As String
        ReDim st(LEVELS - 1)
        Dim codes() As String
        ReDim codes(LEVELS - 1)
        Dim r As Integer
        Dim rs() As Integer
        ReDim rs(LEVELS - 1)

        Dim CodeChange As Boolean

        r = r0 + 1

        objSheet.Range(Excel_Cell(r, LEVELS + 1) & ":" & Excel_Cell(r + rn - 1, LEVELS + 1)).IndentLevel = LEVELS - 1

        For i = 1 To rn + 1
            If i = 1 Then
                st(0) = ""
                For j = 1 To LEVELS - 1
                    codes(j) = objSheet.Cells(r, j)
                    rs(j) = r
                    st(j) = ""
                Next j
            End If

            'TBL(LEVELS).Seek "=", objSheet.Cells(R, LEVELS)
            'If TBL(LEVELS).NoMatch Then
            '    z = ""
            'Else
            '    z = TBL(LEVELS).Fields(1).Value & ""
            'End If
            'objSheet.Cells(R, LEVELS + 1) = z
            'objSheet.Range(Excel_Cell(R, LEVELS + 1)).IndentLevel = LEVELS

            For j = LEVELS - 1 To 1 Step -1
                CodeChange = False
                For k = 1 To j
                    If codes(k) <> objSheet.Cells(r, k) Then
                        CodeChange = True
                    End If
                Next k
                If CodeChange Then
                    objSheet.Rows(r).Insert(xlShiftDown)
                    TBL(j).Seek("=", codes(j))
                    If TBL(j).NoMatch Then
                        z = ""
                    Else
                        z = TBL(j).Fields(1).Value & ""
                    End If
                    objSheet.Cells(r, LEVELS + 1) = z

                    'If i <> i Then
                    objSheet.Range(Excel_Cell(r, LEVELS + 1)).Select()
                    With objApp.Selection
                        '.HorizontalAlignment = xlLeft
                        '.VerticalAlignment = xlBottom
                        '.WrapText = False
                        '.Orientation = 0
                        '.AddIndent = False
                        .IndentLevel = j - 1
                        '.ShrinkToFit = False
                        '.ReadingOrder = xlContext
                        '.MergeCells = False
                        '.Font.ColorIndex
                        '.Interior.ColorIndex = CLR(j)
                        '.Interior.Pattern = xlSolid
                        '.Interior.PatternColorIndex = xlAutomatic
                    End With
                    'objSheet.Cells(R, LEVELS + 1).Font.ColorIndex = CLR(j)
                    objSheet.Range(Excel_Cell(r, LEVELS + 1) & ":" & Excel_Cell(r, LEVELS + 1 + c0)).Font.ColorIndex = CLR(j)
                    'End If
                    For k = 1 To j
                        objSheet.Cells(r, k) = codes(k)
                    Next k
                    st(j - 1) = st(j - 1) & "," & Excel_Cell(r, LEVELS + 1 + 1)

                    If j = LEVELS - 1 Then
                        objSheet.Cells(r, LEVELS + 1 + 1) = "=SUM(" & Excel_Cell(rs(j), LEVELS + 1 + 1) & ":" & Excel_Cell(r - 1, LEVELS + 1 + 1) & ")"
                    Else
                        'objSheet.Cells(R, LEVELS + 1 + 1) = "=SUM(" & Mid$(st(j), 2) & ")"
                        objSheet.Cells(r, LEVELS + 1 + 1) = "=" & Replace(st(j), ",", "+")
                        st(j) = ""
                    End If
                    objSheet.Cells(r, LEVELS + 1 + 1).Select()
                    objApp.Selection.Copy()
                    objSheet.Range(Excel_Cell(r, LEVELS + 1 + 2) & ":" & Excel_Cell(r, LEVELS + 1 + (Groups * c0))).Select()
                    objSheet.Paste()
                    objSheet.Rows(CStr(rs(j)) & ":" & CStr(r - 1)).Group()
                    r = r + 1
                    For k = j To LEVELS - 1
                        rs(k) = r
                    Next k
                    codes(j) = objSheet.Cells(r, j)
                Else
                    j = 1
                End If
            Next j

            r = r + 1
        Next i

        r = r - 1
        objSheet.Cells(r, 1) = "Total"
        If st(0) <> "" Then
            'objSheet.Cells(R, LEVELS + 1 + 1) = "=SUM(" & Mid$(st(0), 2) & ")"
            objSheet.Cells(r, LEVELS + 1 + 1) = "=" & Replace(st(0), ",", "+")
            '
        End If
        objSheet.Cells(r, LEVELS + 1 + 1).Select()
        objApp.Selection.Copy()
        objSheet.Range(Excel_Cell(r, LEVELS + 1 + 2) & ":" & Excel_Cell(r, LEVELS + 1 + (Groups * c0))).Select()
        objSheet.Paste()

        Build_Groups = r
    End Function


End Class