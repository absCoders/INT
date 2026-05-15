Public Class ASFERRV1

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Show_Filter(grdASTERROR, True)

        With grdASTERROR.DisplayLayout.Bands(0)
            With .Columns.Add("DataSet")
                .Style = UltraWinGrid.ColumnStyle.Button
                .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                '.CellAppearance.ImageHAlign = HAlign.Center
                .Header.VisiblePosition = 0
                .Header.Caption = "Data"
                .Width = 50
            End With
            With .Columns.Add("Screen Shot")
                .Style = UltraWinGrid.ColumnStyle.Button
                .ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                '.CellAppearance.ImageHAlign = HAlign.Center
                .Header.VisiblePosition = 1
                .Header.Caption = "Scrn"
                .Width = 50
            End With
        End With

        Create_Summary(grdASTERROR, "DataSet", "Count")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load Errors"
                Call Load_Record()
                Call Mode_Settings(True)

            Case "Done"
                Call Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        With UltraExplorerBar1
            .Groups("Screen Control").Items("Load Errors").Settings.Enabled = not_iScreenMode
            .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
        End With

        Call Set_Read_Only(UltraGroupBox1, ScreenMode)

        SplitContainer1.Visible = ScreenMode

        If ScreenMode Then
        Else
            Clear_Record()
        End If

    End Sub

    Sub Load_Record()

        ASCMAIN1.Progress("Now Loading Errors from Log")
        Me.Cursor = Cursors.WaitCursor


        Dim xdir As System.IO.DirectoryInfo = _
        New System.IO.DirectoryInfo(ASCMAIN1.Folders("Archive") & "ERRs\")

        If ASCMAIN1.DBS_SERVER = "VAN" Or ASCMAIN1.DBS_COMPANY = "VAN" Then
            xdir = New System.IO.DirectoryInfo("R:\VDI\ARCHIVE\VAN\ERRs\")
        End If

        grdASTERROR.DataSource = xdir.GetFiles("*.ERR")

        For Each GC As UltraWinGrid.UltraGridColumn In grdASTERROR.DisplayLayout.Bands(0).Columns
            Select Case GC.Key
                Case "DataSet", "Screen Shot", "Name", "Length", "CreationTime"

                Case Else
                    GC.Hidden = True
            End Select
        Next

        With grdASTERROR.DisplayLayout.Bands(0)
            .Columns("DataSet").Header.VisiblePosition = 0
            .Columns("Screen Shot").Header.VisiblePosition = 1

            .Columns("Name").Width = 250
            .Columns("CreationTime").Format = "MM/dd/yy HH:mm tt"
            .Columns("CreationTime").Width = 150
        End With

        Sort_grdColumns(grdASTERROR, "CreationTime".ToLower)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Clear_Record()

    End Sub

#End Region

#Region "grdASTERROR"
    Private Sub grdASTERROR_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdASTERROR.AfterRowActivate
        Dim DIRECTORYNAME As String = grdASTERROR.ActiveRow.Cells("DIRECTORYNAME").Text
        Dim NAME As String = grdASTERROR.ActiveRow.Cells("NAME").Text

        Try
            txtSTACKTRACE.Text = My.Computer.FileSystem.ReadAllText(DIRECTORYNAME & "\" & NAME)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub grdASTERROR_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdASTERROR.ClickCellButton

        Select Case e.Cell.Column.Key
            Case "DataSet"
                Dim DIRECTORYNAME As String = e.Cell.Row.Cells("DIRECTORYNAME").Text
                Dim NAME As String = e.Cell.Row.Cells("NAME").Text
                Dim FILENAME As String
                FILENAME = DIRECTORYNAME & "\" & Replace(NAME.ToUpper, "ERR", "XML")
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    Try
                        dst = New DataSet
                        dst.EnforceConstraints = False
                        dst.ReadXml(FILENAME)
                        dst.EnforceConstraints = True

                        Dim F As New ASFDSET1(Me)
                        F.Show()
                    Catch ex As Exception

                    End Try

                End If


            Case "Screen Shot"
                Dim DIRECTORYNAME As String = grdASTERROR.ActiveRow.Cells("DIRECTORYNAME").Text
                Dim NAME As String = grdASTERROR.ActiveRow.Cells("NAME").Text
                Dim FILENAME As String
                FILENAME = DIRECTORYNAME & "\" & Replace(NAME.ToUpper, "ERR", "BMP")
                If My.Computer.FileSystem.FileExists(FILENAME) Then
                    Show_Document(FILENAME)
                End If

        End Select
    End Sub

    Private Sub grdASTERROR_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdASTERROR.InitializeRow
        Dim DIRECTORYNAME As String = e.Row.Cells("DIRECTORYNAME").Text
        Dim NAME As String = e.Row.Cells("NAME").Text
        Dim FILENAME As String

        FILENAME = DIRECTORYNAME & "\" & Replace(NAME.ToUpper, "ERR", "BMP")
        If My.Computer.FileSystem.FileExists(FILENAME) Then
            'e.Row.Appearance.BackColor = Drawing.Color.Yellow
            'e.Row.Cells("Screen Shot").Value = "Screen"
            e.Row.Cells("Screen Shot").ButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "screen.png")
            e.Row.Cells("Screen Shot").ButtonAppearance.ImageHAlign = HAlign.Center

        End If

        FILENAME = DIRECTORYNAME & "\" & Replace(NAME.ToUpper, "ERR", "XML")
        If My.Computer.FileSystem.FileExists(FILENAME) Then
            e.Row.Cells("DataSet").ButtonAppearance.Image = ASCMAIN1.Get_Image(ASCMAIN1.Folders("Images") & "16\", "data.png")
            e.Row.Cells("DataSet").ButtonAppearance.ImageHAlign = HAlign.Center
        End If

    End Sub

#End Region
End Class