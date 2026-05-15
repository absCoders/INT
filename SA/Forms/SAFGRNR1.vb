Imports Infragistics.UltraChart.Resources
Imports Infragistics.UltraChart.Shared.Styles

Public Class SAFGRNR1

    Dim RYP As String
    Dim SATGRNR1 As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_SATGRNR1("")
            ASCMAIN1.sql = "Select * from " & SATGRNR1
            Create_TDA(.Tables.Add, "SATGRNR1", "**", 0, False, "", 0)

            With .Tables("SATGRNR1")
                .Columns.Add("MTD_NET", GetType(System.Int64), "MTD_GRS - MTD_RTN")
                .Columns.Add("YTD_NET", GetType(System.Int64), "YTD_GRS - YTD_RTN")
                For i As Integer = 1 To 12
                    Dim ret As String = "RTN_" & Format(i, "00")
                    Dim grs As String = "GRS_" & Format(i, "00")
                    .Columns.Add("NET_" & Format(i, "00"), GetType(System.Int64), grs & " - " & ret)
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
                        .Header.Appearance.BackColor2 = Drawing.Color.Lime
                    ElseIf COLUMN_SFX = "RTN" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.Orange
                    ElseIf COLUMN_SFX = "NET" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.DodgerBlue
                    Else
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
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
                        .Header.Appearance.BackColor2 = Drawing.Color.Lime
                    ElseIf COLUMN_SFX = "RTN" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.Orange
                    ElseIf COLUMN_SFX = "NET" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.DodgerBlue
                    Else
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
                End With
                With grdSATGRNR2.DisplayLayout.Bands(0).Columns(COLUMN_NAME)
                    .Header.Caption = ASCMAIN1.Make_Caption(.Header.Caption)
                    .Format = "###,##0"
                    If COLUMN_SFX = "GRS" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.Lime
                    ElseIf COLUMN_SFX = "RTN" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.Orange
                    ElseIf COLUMN_SFX = "NET" Then
                        .Header.Appearance.BackColor2 = Drawing.Color.DodgerBlue
                    Else
                        .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    End If
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
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
                    .Header.Appearance.BackColor2 = Drawing.Color.LightGray
                    .Header.Appearance.BackColor = Drawing.Color.White
                    .Header.Appearance.BackGradientStyle = Infragistics.Win.GradientStyle.ForwardDiagonal
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
                Validate_Code("OPS_YYYYPP")
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
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Print"
                Print_Report()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                With .Groups("Screen Control")
                    .Items("Load").Settings.Enabled = not_iScreenMode
                    .Items("Done").Settings.Enabled = iScreenMode
                    '.Items("Print").Settings.Enabled = iScreenMode
                End With
            
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
        ASCMAIN1.Progress("Now Compiling Historical Data")
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

        'Dim TEST As Integer = CInt(Today.Year) & CInt(Today.Month)
        'Dim TEST2 = Mid(RYP, 1, 6)

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
        ASCMAIN1.Progress("")
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

    Sub Create_SATGRNR1(ByVal RYP As String)

        Dim sql_amt As String = "SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE * 1.7"
        Dim sql_GR As String = "SOTINVH2.INV_TYPE = '{0}'"
        Dim sql_MTD As String = "SOTINVH2.ORDR_YYYYPP_UPDATED = '" & RYP & "'"
        Dim sql_YTD As String = "SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & Mid(RYP, 1, 4) & "01" _
            & "' and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'"

        Dim month_sum As String = Nothing
        Dim sql_MNTHS As String = Nothing
        Dim sql_MTDi As String = Nothing
        Dim sql_RET As String = Nothing

        For i As Integer = 1 To 12
            sql_MTDi = "SOTINVH2.ORDR_YYYYPP_UPDATED = '" & Mid(RYP, 1, 4) & Format(i, "00") & "'"
            month_sum = ", SUM (CASE WHEN " & String.Format(sql_GR, "I") & " AND " & sql_MTDi & " THEN " & vbCrLf _
                            & sql_amt & " ELSE 0 END) GRS_" & Format(i, "00") & vbCrLf
            sql_MNTHS = sql_MNTHS & month_sum
        Next
        For i As Integer = 1 To 12
            sql_MTDi = "SOTINVH2.ORDR_YYYYPP_UPDATED = '" & Mid(RYP, 1, 4) & Format(i, "00") & "'"
            sql_RET = ", SUM (CASE WHEN " & String.Format(sql_GR, "C") & " AND " & sql_MTDi & " THEN -1 * " & vbCrLf _
                            & sql_amt & " ELSE 0 END) RTN_" & Format(i, "00") & vbCrLf
            sql_MNTHS = sql_MNTHS & sql_RET
        Next

        ASCMAIN1.sql = "SELECT X.*" & vbCrLf _
        & ", DECODE(X.CLASS,'ECOM','JHU','CAR','CAR','JHU') COMPANY" & vbCrLf _
        & ", DECODE(X.CLASS,'ECOM','ECOM','CAR','CAR','NAT',X.CUST_DBA_NAME,'IND',SOTSREG1.REGION_DESC,X.CLASS) REG" & vbCrLf _
        & ", DECODE(X.CLASS,'ECOM','ECOM','CAR','CAR','NAT',X.CUST_DBA_NAME,'IND',SOTSREP1.SREP_NAME,X.CLASS) GRP" & vbCrLf _
        & " from SOTSREP1,SOTSREG1,(" & vbCrLf _
        & "Select SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, ARTCUST1.CUST_DBA_NAME" & vbCrLf _
        & ", DECODE(ICTITEM1.COLLECTION_CODE,'CINTA','CINTA',ARTCUST1.TRADE_CLASS_CODE) CLASS" & vbCrLf _
        & ", SUM (CASE WHEN " & String.Format(sql_GR, "I") & " AND " & sql_MTD & " THEN " & sql_amt & " ELSE 0 END) MTD_GRS" & vbCrLf _
        & ", SUM (CASE WHEN " & String.Format(sql_GR, "I") & " AND " & sql_YTD & " THEN " & sql_amt & " ELSE 0 END) YTD_GRS" & vbCrLf _
        & ", SUM (CASE WHEN " & String.Format(sql_GR, "C") & " AND " & sql_MTD & " THEN -1 * " & sql_amt & " ELSE 0 END) MTD_RTN" & vbCrLf _
        & ", SUM (CASE WHEN " & String.Format(sql_GR, "C") & " AND " & sql_YTD & " THEN -1 * " & sql_amt & " ELSE 0 END) YTD_RTN" & vbCrLf _
        & sql_MNTHS _
        & " from SOTINVH2,ARTCUST1,ICTITEM1" & vbCrLf _
        & " where ARTCUST1.CUST_CODE (+) = SOTINVH2.CUST_CODE" & vbCrLf _
        & "   and ICTITEM1.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
        & "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & Mid(RYP, 1, 4) & "01" & "'" & vbCrLf _
        & "   and SOTINVH2.ORDR_YYYYPP_UPDATED <= '" & RYP & "'" & vbCrLf _
        & " group by SOTINVH2.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST1.SREP_CODE, ARTCUST1.CUST_DBA_NAME" & vbCrLf _
        & ", DECODE(ICTITEM1.COLLECTION_CODE,'CINTA','CINTA',ARTCUST1.TRADE_CLASS_CODE)" & vbCrLf _
        & ") X" & vbCrLf _
        & " where SOTSREP1.SREP_CODE (+) = X.SREP_CODE" & vbCrLf _
        & "   and SOTSREG1.REGION_CODE (+) = SOTSREP1.REGION_CODE"

        If SATGRNR1 = "" Then
            SATGRNR1 = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & SATGRNR1 & " Add Primary Key (CUST_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SATGRNR1)
            ASCDATA1.ExecuteSQL("Insert into " & SATGRNR1 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Print_Report_End()
    End Sub

    Sub Setup_Summary()
        If SELECTION_NO = 0 Then Exit Sub
    End Sub

    Sub Check_to_Collaps(ByVal grows As UltraWinGrid.RowsCollection)
        For Each grow As UltraWinGrid.UltraGridRow In grows
            If grow.IsGroupByRow Then

                Dim gbrow As UltraWinGrid.UltraGridGroupByRow = DirectCast(grow, UltraWinGrid.UltraGridGroupByRow)
                'gbrow.Band.SummaryFooterCaption = ""
                If gbrow.Rows.Count = 1 Then
                    gbrow.CollapseAll()
                ElseIf gbrow.Column.Key = "COMPANY" And gbrow.Value & "" <> "JHU" Then
                    gbrow.CollapseAll()
                ElseIf gbrow.Column.Key = "GRP" Then
                    gbrow.CollapseAll()
                ElseIf gbrow.Column.Key = "CLASS" And gbrow.Value & "" <> "NAT" And gbrow.Value & "" <> "IND" Then
                    gbrow.CollapseAll()
                ElseIf gbrow.Column.Key = "REG" AndAlso gbrow.ParentRow IsNot Nothing AndAlso gbrow.ParentRow.IsGroupByRow AndAlso DirectCast(gbrow.ParentRow, UltraWinGrid.UltraGridGroupByRow).Value & "" <> "IND" Then
                    gbrow.CollapseAll()
                Else
                    Check_to_Collaps(gbrow.Rows)
                End If
            End If
        Next
    End Sub
    Sub Format_groupbyrow(ByVal e As Infragistics.Win.UltraWinGrid.InitializeGroupByRowEventArgs)
        e.Row.Description = e.Row.ValueAsDisplayText & ""
        Dim V As String = e.Row.Value & ""
        If e.Row.Column.Key = "COMPANY" Then
            If V = "CAR" Then e.Row.Description = "Caribbean"
            If V = "JHU" Then e.Row.Description = "US"
        ElseIf e.Row.Column.Key = "CLASS" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightBlue
            If V = "CAR" Then e.Row.Description = "Caribbean"
            If V = "ECOM" Then e.Row.Description = "eCommerce"
            If V = "NAT" Then e.Row.Description = "Major"
            If V = "IND" Then e.Row.Description = "Independents"
            If V = "CINTA" Then e.Row.Description = "Cinta & Rosecut"
            If V = "CON" Then e.Row.Description = "Consignment"
            If V = "OTH" Then e.Row.Description = "Other Regions"
        ElseIf e.Row.Column.Key = "REG" Then
            e.Row.Appearance.BackColor = Drawing.Color.LightYellow
        ElseIf e.Row.Column.Key = "GRP" Then
            e.Row.Appearance.BackColor = Drawing.Color.Linen
        End If
    End Sub
    Private Sub grdSATGRNR1_InitializeGroupByRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeGroupByRowEventArgs) Handles grdSATGRNR1.InitializeGroupByRow
        Format_groupbyrow(e)
    End Sub

    Private Sub grdSATGRNR2_InitializeGroupByRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeGroupByRowEventArgs) Handles grdSATGRNR2.InitializeGroupByRow
        Format_groupbyrow(e)
    End Sub

    Private Sub chkGross_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkGross.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        chkGRN("GRS", chkGross.Checked)
    End Sub

    Private Sub chkReturns_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReturns.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        chkGRN("RTN", chkReturns.Checked)
    End Sub

    Private Sub chkNet_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNet.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        chkGRN("NET", chkNet.Checked)
    End Sub

    Private Sub chkGRN(ByVal type As String, ByVal checked As Boolean)
        For i As Integer = 1 To 12
            If (i > Val(Mid(RYP, 5, 2))) Then Exit Sub
            Select Case type
                Case "GRS"
                    grdSATGRNR2.DisplayLayout.Bands(0).Columns("GRS_" & Format(i, "00")).Hidden = Not checked
                    grdSATGRNR2.DisplayLayout.Bands(0).Columns("Ytd_Grs").Hidden = Not checked
                Case "RTN"
                    grdSATGRNR2.DisplayLayout.Bands(0).Columns("RTN_" & Format(i, "00")).Hidden = Not checked
                    grdSATGRNR2.DisplayLayout.Bands(0).Columns("Ytd_Rtn").Hidden = Not checked
                Case "NET"
                    grdSATGRNR2.DisplayLayout.Bands(0).Columns("NET_" & Format(i, "00")).Hidden = Not checked
                    grdSATGRNR2.DisplayLayout.Bands(0).Columns("Ytd_Net").Hidden = Not checked
            End Select
        Next

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
End Class

Public Class srtComparerSATGRNR1
    Implements IComparer

    Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare

        Dim xCell As UltraWinGrid.UltraGridCell = DirectCast(x, UltraWinGrid.UltraGridCell)
        Dim yCell As UltraWinGrid.UltraGridCell = DirectCast(y, UltraWinGrid.UltraGridCell)

        Dim xv As String = xCell.Value & ""
        Dim yv As String = yCell.Value & ""

        Dim COLUMN_NAME As String = xCell.Column.Key

        If xv = yv Then
            Return 0
        Else
            If COLUMN_NAME = "COMPANY" Then
                If xv = "JHU" Then
                    Return -1
                ElseIf yv = "JHU" Then
                    Return 1
                Else
                    Return IIf(xv < yv, -1, 1)
                End If
            ElseIf COLUMN_NAME = "CLASS" Then
                If xv = "NAT" Then
                    Return -1
                ElseIf yv = "NAT" Then
                    Return 1
                ElseIf xv = "IND" Then
                    Return -1
                ElseIf yv = "IND" Then
                    Return 1
                Else
                    Return IIf(xv < yv, -1, 1)
                End If
            Else ' If COLUMN_NAME = "REG" Then
                xv = Z_Fix(xv)
                yv = Z_Fix(yv)
               
                Return IIf(xv < yv, -1, 1)
            End If
        End If
       
    End Function

    Function Z_Fix(ByVal v As String) As String
        If v = "Northeast" Then
            v = "A" & v
        ElseIf v = "Southeast" Then
            v = "B" & v
        ElseIf v = "Southwast" Then
            v = "C" & v
        ElseIf v = "Mid-West" Then
            v = "D" & v
        ElseIf v = "West" Then
            v = "E" & v
        Else
            v = "Z" & v
        End If

        Return v
    End Function
End Class