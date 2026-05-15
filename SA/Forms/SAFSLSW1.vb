Public Class SAFSLSW1

    Dim RYP As String
    Dim SATSLSI1 As String
    Dim YMIN As Integer

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        ASCMAIN1.sql = "Select Min (SUBSTR(YYYYWW,1,4)) YYYY from SATSLSW1"
        YMIN = ASCDATA1.GetDataValue

        With dst

            With .Tables.Add("SATSLSWX")
                .Columns.Add("OP_DIV_CODE")
                .Columns.Add("STORE_NO")
                For Y As Integer = 1 To 50
                    .Columns.Add("Y" & Format(Y, "00"), GetType(System.Decimal))
                    .Columns.Add("Y" & Format(Y, "00") & "P", GetType(System.Decimal))
                Next
                For Y As Integer = 1 To 49
                    Dim EXP As String = "IIF(" & "Y" & Format(Y + 1, "00") & " = 0,0,100 * (" & "Y" & Format(Y, "00") & " - " & "Y" & Format(Y + 1, "00") & ") / " & "Y" & Format(Y + 1, "00") & ")"
                    .Columns("Y" & Format(Y, "00") & "P").Expression = EXP
                Next
            End With

        End With

        grdSATSLSWX.DataSource = dst.Tables("SATSLSWX")


        grdSATSLSWX.DisplayLayout.UseFixedHeaders = True
        With grdSATSLSWX.DisplayLayout.Bands("SATSLSWX")
            With .Columns("OP_DIV_CODE")
                .Width = 60
                .Header.Caption = "OpDiv"
                .Header.Fixed = True
            End With

            With .Columns("STORE_NO")
                .Width = 60
                .Header.Caption = "Store"
                .Header.Fixed = True
            End With

            Create_Summary(grdSATSLSWX, "STORE_NO", "Count")
            For Y As Integer = 1 To 50

                With .Columns("Y" & Format(Y, "00"))
                    .Width = 80
                    .Format = "###,##0.0"
                End With
                With .Columns("Y" & Format(Y, "00") & "P")
                    .Width = 80
                    .Format = "##0.0"
                End With

                Create_Summary(grdSATSLSWX, "Y" & Format(Y, "00"))
            Next
        End With

        Absx1.txtFor("OPS_YYYYPP").Text = ASCMAIN1.CYP ' .ToString.Substring(0, 4) & "12"

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                Validate_Code("OPS_YYYYPP")
                If EMsg = "" Then
                    RYP = Absx1.txtFor("OPS_YYYYPP").Text
                    If Mid(RYP, 1, 4) < Format(YMIN, "0000") Then
                        EMsg &= vbCr & "Year Cannot be prior to " & Format(YMIN, "0000")
                    Else
                        If Val(Mid(RYP, 1, 4)) - YMIN > 49 Then
                            EMsg &= vbCr & "Screen supports only 49 years"
                        End If
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
            End With
        End If

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdSATSLSWX.Visible = tf

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("SATSLSWX").Rows.Clear()
        dst.EnforceConstraints = True
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        Call ASCMAIN1.Progress("Now Compiling Historical Data")
        Application.DoEvents()

        Call Save_Header_Fields(UltraGroupBox1)

        Create_SATSLSWX(RYP)


        ASCMAIN1.Progress("Now Setting Up Screen")


        grdSATSLSWX.DisplayLayout.Rows.ColumnFilters.ClearAllFilters()
        Sort_grdColumns(grdSATSLSWX, "STORE_NO")
        grdSATSLSWX.DisplayLayout.Bands(0).SortedColumns.Add("OP_DIV_CODE", False, True)
        grdSATSLSWX.Rows.ExpandAll(True)

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
        Call Load_Popup_Menu(grdSATSLSWX, "SSBSB", "Show Filter", "Show GroupBox")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool

        If tlb_pop.Tools.Exists("Show Filter") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
        End If
        If tlb_pop.Tools.Exists("Show GroupBox") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
        End If

        'If tlb_pop.Tools.Exists("Include Inactive") Then
        'End If


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else
            'If e.Tool.Key <> "grdSATCSLSS" Then
            '    Dim tlb_sbt As UltraWinToolbars.StateButtonTool
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = (grd.DisplayLayout.Override.AllowRowFiltering = DefaultableBoolean.True)
            '    tlb_sbt = DirectCast(tlb_pop.Tools("Show GroupBox"), UltraWinToolbars.StateButtonTool)
            '    tlb_sbt.Checked = Not grd.DisplayLayout.GroupByBox.Hidden
            'End If

            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"



            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Filter"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                Show_Filter(grd, tlb_sbt.Checked)

            Case "Show GroupBox"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked

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

    Sub Create_SATSLSWX(ByVal RYP As String)

        EnforceConstraints(False)

        Dim YMAX As Integer = Val(Mid(RYP, 1, 4)) - 1

        ASCMAIN1.sql = "SELECT GMTSTOR1.OP_DIV_CODE, SATSLSW1.STORE_NO" & vbCrLf
        For Y As Integer = YMAX To YMIN Step -1
            Dim YYYY As String = Format(Y, "0000")
            Dim YY As String = "Y" & Format(YMAX - Y + 1, "00")
            ASCMAIN1.sql &= ", SUM (DECODE(SUBSTR(SATSLSW1.YYYYWW,1,4), '" & YYYY & "', NVL(SATSLSW1.RETAIL_SALES,0),0)) " & YY & VBCRLF
        Next
        ASCMAIN1.sql &= " from SATSLSW1, GMTSTOR1" & vbCrLf _
        & " where GMTSTOR1.STORE_NO (+) = SATSLSW1.STORE_NO" & vbCrLf _
        & IIf(opt53.Value = "52", " and SUBSTR(YYYYWW,5,2) <> '53'", "") & vbCrLf _
        & " group by GMTSTOR1.OP_DIV_CODE, SATSLSW1.STORE_NO"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowSATSLSWX As DataRow = dst.Tables("SATSLSWX").NewRow
            rowSATSLSWX.Item("OP_DIV_CODE") = row.Item("OP_DIV_CODE")
            rowSATSLSWX.Item("STORE_NO") = row.Item("STORE_NO")
            For Y As Integer = 1 To YMAX - YMIN + 1
                Dim YY As String = "Y" & Format(Y, "00")
                rowSATSLSWX.Item(YY) = row.Item(YY) / 1000
                With grdSATSLSWX.DisplayLayout.Bands(0)
                    .Columns(YY).Hidden = False
                    .Columns(YY).Header.Caption = Format(YMAX - Y + 1, "0000")
                    .Columns(YY & "P").Header.Caption = Mid(Format(YMAX - Y + 1, "0000"), 3, 2) & "/" & Mid(Format(YMAX - Y + 1 - 1, "00"), 3, 2)
                    .Columns(YY & "P").CellAppearance.BackColor = Drawing.Color.Beige
                End With

            Next
            For Y As Integer = YMAX - YMIN + 1 To 50
                Dim YY As String = "Y" & Format(Y, "00")
                If Y > YMAX - YMIN + 1 Then
                    grdSATSLSWX.DisplayLayout.Bands(0).Columns(YY).Hidden = True
                End If
                grdSATSLSWX.DisplayLayout.Bands(0).Columns(YY & "P").Hidden = True
            Next
            dst.Tables("SATSLSWX").Rows.Add(rowSATSLSWX)
        Next


        EnforceConstraints(True)


    End Sub

    Private Sub grdSATSLSWX_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdSATSLSWX.InitializeLayout

    End Sub

    Private Sub grdSATSLSWX_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSATSLSWX.InitializeRow
        For Y As Integer = 1 To 50
            Dim YY As String = "Y" & Format(Y, "00") & "P"
            If Val(e.Row.Cells(YY).Value & "") < 0 Then
                e.Row.Cells(YY).Appearance.ForeColor = Drawing.Color.Red
            Else
                e.Row.Cells(YY).Appearance.ForeColor = Drawing.Color.Empty
            End If
        Next
    End Sub
End Class