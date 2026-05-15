Public Class SAFSLSX2

    Dim DTE1 As Date
    Dim DTE2 As Date
    Dim FILENAME As String = ""

    Dim SOTINVHX As String
    Dim sqlSOTINVH1_S As String = ""

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst


        End With



        'Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 12, 0) ' -11 - 12)
        'Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 12, 0) '  0 - 12)

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"
                'DTE1 = dteFrom.Value
                'DTE2 = dteTo.Value


                Dim FILENAME As String = ""
                Using openFileDialog1 As New OpenFileDialog
                    'openFileDialog1.InitialDirectory = ASCMAIN1.Folders("Work")
                    openFileDialog1.Title = "Locate the workbook containing the data to Import"
                    openFileDialog1.Filter = "xls files (*.xls)|*.xls|xlsX files (*.xlsx)|*.xlsx"
                    openFileDialog1.FilterIndex = 2
                    openFileDialog1.RestoreDirectory = True
                    If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        FILENAME = openFileDialog1.FileName
                        Stop

                        WorkbookView1.ActiveWorkbook = SpreadsheetGear.Factory.GetWorkbook(FILENAME)
                    End If
                End Using

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

            Case "Show Temp Folder"
                'Process.Start("explorer.exe", ASCMAIN1.Folders("Temp"))


            Case "Done"
                Mode_Settings(False)

            Case "ftp File"
                'ftp_File()
                TAC.SACMAIN1.ftp_BI_Files(Me)

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
                    .Items("ftp File").Settings.Enabled = iScreenMode
                    .Items("Show Temp Folder").Visible = ScreenMode

                    If ASCMAIN1.DBS_SERVER = "INT" Or ASCMAIN1.DBS_COMPANY = "INT" Then
                        .Items("ftp File").Visible = False
                        .Items("Show Temp Folder").Visible = False

                    End If

                End With
            End With
        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        tabData.Visible = True
        '   UltraExplorerBarContainerControl1.Enabled = Not tf

        If ScreenMode Then

        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        EnforceConstraints(False)
        'For Each TABLE_NAME As String In New String() {"SATSLSXC", "SATSLSXI", "SATSLSXS"}
        '    dst.Tables(TABLE_NAME).Rows.Clear()
        'Next


        EnforceConstraints(True)
    End Sub

    Sub Load_Record()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data from Database")

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)



        EnforceConstraints(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()
        BeginTrans()

        CommitTrans("Update Complete")
    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        'Load_Popup_Menu(grdSATSLSXC, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
        'Load_Popup_Menu(grdSOTINVH1_S, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing


        'If tlb_pop.Tools.Exists("Include Inactive") Then
        'End If

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            e.Cancel = True
        Else

            Select Case e.SourceControl.Name
                Case "grdSATCSLS1"
            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key

        End Select


        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            Case "Item Status Inquiry"
                Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select

    End Sub

#End Region

    Sub Print_Report()
        Dim SUBT As String = ""

        Print_Report_Begin()

        'CR_params.Add("DETAIL", "Y")
        'Generate_Report("SARSLSJ5", "Sales Analysis by Rep/Customer-Rank", SUBT)

        Print_Report_End()
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        'WorkbookView1.GetLock()
        'Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("RSFSSPL1.XLSX_NO") & ".XLSX"
        'WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        'Show_Document(FILENAME)
        'WorkbookView1.ReleaseLock()
    End Sub


End Class