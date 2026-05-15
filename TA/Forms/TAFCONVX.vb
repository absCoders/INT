Public Class TAFCONVX
    Dim tblTATCONVZ As New DataTable
    Dim TABLE_NAME_conv As String
    Dim COLUMN_NAME_1ST As String

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        With dst
            Create_TDA(.Tables.Add, "TATCONVX", "*", 0, True)

            ASCMAIN1.sql = "Select * from TATCONVY where TABLE_NAME = :PARM1"
            Create_TDA(.Tables.Add, "TATCONVY", "**", 0, True, "V")
            With .Tables("TATCONVY")
                .Columns.Add("HIDE")
            End With

            With .Tables.Add("TATCONVQ")
                .Columns.Add("DATA_VALUE")
                .Columns.Add("RECORDS", GetType(System.Int64))
                .Columns.Add("KEY_MIN")
                .Columns.Add("KEY_MAX")
            End With
        End With

        grdTATCONVX.DataSource = dst.Tables("TATCONVX")
        grdTATCONVY.DataSource = dst.Tables("TATCONVY")
        grdTATCONVQ.DataSource = dst.Tables("TATCONVQ")

        Create_Summary(grdTATCONVX, "ODOBNM", "Count")
        Create_Summary(grdTATCONVY, "COLUMN_NAME", "Count")
        Create_Summary(grdTATCONVQ, "DATA_VALUE", "Count")

        numMaxRows.Value = 100000

        For Each gcol As UltraWinGrid.UltraGridColumn In grdTATCONVY.DisplayLayout.Bands(0).Columns
            If gcol.Key = "HIDE" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next
        grdTATCONVY.DisplayLayout.Bands(0).Override.AllowUpdate = DefaultableBoolean.True

        SplitContainer3.Panel2Collapsed = Not chkShowValueCounts.Checked
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load Table"


        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Table"
                Load_Record()
                Mode_Settings(True)

            Case "Analyze Table"
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Analyzing Table")
                ASCMAIN1.AnalyzeTable(TABLE_NAME_conv, "CONV")
                ASCMAIN1.sql = "SELECT * FROM ALL_TABLES WHERE OWNER = 'CONV' AND TABLE_NAME = '" & TABLE_NAME_conv & "'"
                Dim row As DataRow = ASCDATA1.GetDataRow
                numRows.Value = Val(row.Item("NUM_ROWS") & "")
                '  numRows.AlwaysInEditMode = False
                ASCMAIN1.sql = "SELECT * FROM ALL_TAB_COLUMNS WHERE OWNER = 'CONV' AND TABLE_NAME = '" & TABLE_NAME_conv & "'"
                For Each row2 As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim rowTATCONVY As DataRow = dst.Tables("TATCONVY").Rows.Find _
                                                 (New String() {row2.Item("TABLE_NAME"), row2.Item("COLUMN_NAME")})
                    rowTATCONVY.Item("DATA_TYPE") = row2.Item("DATA_TYPE")
                    rowTATCONVY.Item("DATA_LENGTH") = row2.Item("DATA_LENGTH")
                    rowTATCONVY.Item("DATA_PRECISION") = row2.Item("DATA_PRECISION")
                    rowTATCONVY.Item("DATA_SCALE") = row2.Item("DATA_SCALE")
                    rowTATCONVY.Item("NUM_DISTINCT") = row2.Item("NUM_DISTINCT")
                    rowTATCONVY.Item("NUM_DISTINCT") = row2.Item("NUM_DISTINCT")
                Next
                Update_Record_TDA("TATCONVX")
                Update_Record_TDA("TATCONVY")
                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

            Case "Done"
                Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load Table").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Analyze Table").Settings.Enabled = iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            .Groups("Data Controls").Visible = ScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        grdTATCONVX.Visible = Not ScreenMode
        SplitContainer1.Visible = ScreenMode

        lblRows.Visible = ScreenMode
        numRows.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Errors from Log")
        Me.Cursor = Cursors.WaitCursor

        Dim rowTATCONVX As DataRow = LookUp("TATCONVX", Absx1.txtFor("ODOBNM").Text)
        ASCMAIN1.sql = "Select * from TATCONVX where ODOBNM = '" & Absx1.txtFor("ODOBNM").Text & "' and ODLBNM = '" & Absx1.txtFor("ODLBNM").Text & "'"
        rowTATCONVX = ASCDATA1.GetDataRow

        Absx1.txtFor("ODLBNM").Text = rowTATCONVX.Item("ODLBNM")

        Fetch_Rows()
        grdTATCONVZ.DataSource = tblTATCONVZ
        COLUMN_NAME_1ST = tblTATCONVZ.Columns(0).ColumnName

        TABLE_NAME_conv = Absx1.txtFor("ODLBNM").Text & "_" & Absx1.txtFor("ODOBNM").Text
        Fill_Records("TATCONVY", TABLE_NAME_conv)
        Sort_grdColumns(grdTATCONVY, "COLUMN_ID")
        'For Each DCOL As DataColumn In tblTATCONVZ.Columns
        '    Dim rowTATCONVY As DataRow = dst.Tables("TATCONVY").NewRow
        '    rowTATCONVY.Item("COLUMN_NAME") = DCOL.ColumnName
        '    rowTATCONVY.Item("DATA_TYPE") = DCOL.ColumnName
        '    rowTATCONVY.Item("MAX_LENGTH") = DCOL.ColumnName
        '    rowTATCONVY.Item("PRECISION") = DCOL.ColumnName
        '    dst.Tables("TATCONVY").Rows.Add(ROWTATCONVY)
        'Next

        grdTATCONVY.Text = "Data Columns for " & TABLE_NAME_conv
        grdTATCONVY.Text = "Data Content for " & TABLE_NAME_conv

        numRows.Value = rowTATCONVX.Item("RECS")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()
        Absx1.txtFor("ODOBNM").Text = ""
        Absx1.txtFor("ODLBNM").Text = ""
        Absx1.txtFor("ODOBTX").Text = ""

        For Each TABLE_NAME As String In New String() {"TATCONVY"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If dst.Tables("TATCONVX").Rows.Count = 0 Then
            ASCMAIN1.sql = "Select * from TATCONVX"
            Fill_Records("TATCONVX", "", , ASCMAIN1.sql)
            Sort_grdColumns(grdTATCONVX, "ODOBNM")
        End If

        grdTATCONVZ.DataSource = Nothing
        grdTATCONVZ.DisplayLayout.NewBandLoadStyle = UltraWinGrid.NewBandLoadStyle.Show
        grdTATCONVZ.DisplayLayout.NewColumnLoadStyle = UltraWinGrid.NewColumnLoadStyle.Show
 
        Application.DoEvents()
    End Sub

#End Region


#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()

        Load_Popup_Menu(grdTATCONVX, "SSS", "Show GroupBox", "Show Filter", "Show Pins")
        Load_Popup_Menu(grdTATCONVY, "B", "Hide <= 1")
        Load_Popup_Menu(grdTATCONVZ, "SSS", "Show GroupBox", "Show Filter", "Show Pins")

    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name
            End Select

        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)
        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Hide <= 1"
                For Each grow As UltraWinGrid.UltraGridRow In grdTATCONVY.Rows
                    If Val(grow.Cells("NUM_DISTINCT").Value & "") <= 1 Then
                        grow.Cells("HIDE").Value = "1"
                        grow.Update()
                    End If
                Next
        End Select

        Select Case grd.Name

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region
#Region "ABSColumn Controls"

    Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As System.Windows.Forms.Control)
        
    End Sub

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
       
    End Sub
#End Region

#Region "grdASTERROR"

#End Region

    Private Sub grdTATCONVX_DoubleClickRow(sender As Object, e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdTATCONVX.DoubleClickRow

        If e.Row.IsDataRow Then
            Absx1.txtFor("ODOBNM").Text = e.Row.Cells("ODOBNM").Value
            Absx1.txtFor("ODLBNM").Text = e.Row.Cells("ODLBNM").Value

            Dim ODOBNM As String = Absx1.txtFor("ODOBNM").Text
            Dim ODLBNM As String = Absx1.txtFor("ODLBNM").Text

            ASCMAIN1.sql = "Select * from CONV." & ODLBNM & "_" & ODOBNM & " where ROWNUM < 1"
            Try
                Dim tbl As DataTable = ASCDATA1.GetDataTable
            Catch ex As Exception
                Stop

            End Try

            Click_Command("Load Table")
        End If

    End Sub

    Sub Fetch_Rows()
        ASCMAIN1.sql = "Select * from CONV." & Absx1.txtFor("ODLBNM").Text & "_" & Absx1.txtFor("ODOBNM").Text
        Dim sqlMaxRows As String = ""
        If numMaxRows.Value <> 0 Then
        sqlMaxRows = " ROWNUM <= " & CStr(numMaxRows.Value)
        End If

        If txtWhere.Text <> "" Then
            ASCMAIN1.sql &= " where " & txtWhere.Text
            If sqlMaxRows <> "" Then
                ASCMAIN1.sql &= " and " & sqlMaxRows
            End If
        Else
            If sqlMaxRows <> "" Then
                ASCMAIN1.sql &= " where " & sqlMaxRows
            End If
        End If

        Try
            tblTATCONVZ = ASCDATA1.GetDataTable

        Catch ex As Exception

        End Try

    End Sub
    
    Private Sub cmdFetch_Click(sender As System.Object, e As System.EventArgs) Handles cmdFetch.Click
        Fetch_Rows()
    End Sub

    Private Sub chkShowValueCounts_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowValueCounts.CheckedChanged
        SplitContainer3.Panel2Collapsed = Not chkShowValueCounts.Checked
        If chkShowValueCounts.Checked Then Set_TATCONVQ()
    End Sub

    Private Sub grdTATCONVY_AfterRowActivate(sender As Object, e As System.EventArgs) Handles grdTATCONVY.AfterRowActivate
        Set_TATCONVQ()
    End Sub

    Sub Set_TATCONVQ()
        If Not chkShowValueCounts.Checked Then Exit Sub

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Collecting Data Sample")

        Dim COLUMN_NAME As String = grdTATCONVY.ActiveRow.Cells("COLUMN_NAME").Value

        dst.Tables("TATCONVQ").Rows.Clear()
        ASCMAIN1.sql = "Select " & COLUMN_NAME & " DATA_VALUE, COUNT (*) RECORDS" _
            & ", MIN (" & COLUMN_NAME_1ST & ") KEY_MIN, MAX (" & COLUMN_NAME_1ST & ") KEY_MAX" _
            & " from CONV." & TABLE_NAME_conv & " GROUP BY " & COLUMN_NAME
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Rows
            Dim rowTATCONVQ As DataRow = dst.Tables("TATCONVQ").NewRow
            rowTATCONVQ.Item("DATA_VALUE") = ROW.Item("DATA_VALUE")
            rowTATCONVQ.Item("RECORDS") = ROW.Item("RECORDS")
            rowTATCONVQ.Item("KEY_MIN") = ROW.Item("KEY_MIN")
            rowTATCONVQ.Item("KEY_MAX") = ROW.Item("KEY_MAX")
            dst.Tables("TATCONVQ").Rows.Add(rowTATCONVQ)
        Next
        'For Each row As DataRow In ASCDATA1.SelectDistinct(tblTATCONVZ, New String() {COLUMN_NAME}).Select("")
        '    Dim rowTATCONVQ As DataRow = dst.Tables("TATCONVQ").NewRow
        '    rowTATCONVQ.Item("DATA_VALUE") = row.Item(0)
        '    rowTATCONVQ.Item("RECORDS") = tblTATCONVZ.Compute("COUNT(*)", COLUMN_NAME & " = ")
        '    dst.Tables("TATCONVQ").Rows.Add(rowTATCONVQ)
        'Next

        grdTATCONVQ.Text = "Data Values for " & COLUMN_NAME
        Sort_grdColumns(grdTATCONVQ, "DATA_VALUE")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub grdTATCONVY_AfterRowUpdate(sender As Object, e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles grdTATCONVY.AfterRowUpdate
        Dim C As String = e.Row.Cells("COLUMN_NAME").Value
        grdTATCONVZ.DisplayLayout.Bands(0).Columns(C).Hidden = (e.Row.Cells("HIDE").Value & "" = "1")
    End Sub
End Class