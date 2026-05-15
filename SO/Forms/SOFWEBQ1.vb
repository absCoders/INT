Imports ABSolution

Public Class SOFWEBQ1

    Private startDate As Date = Nothing
    Private endDate As Date = Nothing

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        With dst

            ASCMAIN1.sql = "SELECT * FROM WBTWEBQ1 WHERE TRUNC(INIT_DATE) BETWEEN :PARM1 AND :PARM2"
            MyBase.Create_TDA(.Tables.Add, "WBTWEBQ1", ASCMAIN1.sql, 0, False, "DD", 1)

            dteFrom.DateTime = DateAdd(DateInterval.Day, -7, DateTime.Now)
            dteTo.DateTime = DateTime.Now

        End With

        grdWBTWEBQ1.DataSource = dst.Tables("WBTWEBQ1")

        ASCMAIN1.Add_Value_List(grdWBTWEBQ1, "JOB_TYPE", "Select FORM_CODE JOB_TYPE, FORM_DESC from WBTFORM1")
        ASCMAIN1.Add_Value_List(grdWBTWEBQ1, "JOB_STATUS", Nothing, New String() {":", "Q:Queue", "E:Error", "A:Completed"})

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
                startDate = dteFrom.DateTime.ToShortDateString
                endDate = dteTo.DateTime.ToShortDateString

                If DateDiff(DateInterval.Day, startDate, endDate) < 0 Then
                    EMsg &= vbCr & "Start Date must be less/equal End Date."
                End If

            Case "Done"

        End Select

        If MyBase.EMsg <> "" Then
            MsgBox(MyBase.EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Me.Proceed(eItemKey)

    End Sub

    Private Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "Refresh"
                Load_Record()

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Refresh").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then
        Else
            Me.Clear_Record()
        End If

        grdWBTWEBQ1.Visible = ScreenMode
        UltraExplorerBarContainerControl1.Enabled = Not ScreenMode

    End Sub

    Private Sub Clear_Record()
        EnforceConstraints(False)
        dst.Tables("WBTWEBQ1").Rows.Clear()
        EnforceConstraints(True)
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading information")

        EnforceConstraints(False)

        Dim dte1 As String = startDate.ToString("dd-MMM-yyyy")
        Dim dte2 As String = endDate.ToString("dd-MMM-yyyy")
        'Fill_Records("WBTWEBQ1", New Object() {dte1, dte2})

        ASCMAIN1.sql = $"SELECT * FROM WBTWEBQ1 WHERE TRUNC(INIT_DATE) BETWEEN '{dte1}' AND '{dte2}'"
        Fill_Records("WBTWEBQ1", String.Empty, True, ASCMAIN1.sql)

        EnforceConstraints(True)

        grdWBTWEBQ1.DisplayLayout.PerformAutoResizeColumns(False, UltraWinGrid.PerformAutoSizeType.AllRowsInBand, True)
        Sort_grdColumns(grdWBTWEBQ1, "job_no")

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            BeginTrans()

            CommitTrans("Update Complete")
        Catch ex As Exception
            Rollback(ex.Message)
        End Try

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdWBTWEBQ1, "SSB", "Show Filter", "Show GroupBox", "Re-Queue")
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
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Re-Queue"
                If grdWBTWEBQ1.Selected.Rows.Count = 0 Then
                    Exit Sub
                End If

                Dim jobNoList As New List(Of String)
                For Each grdrow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdWBTWEBQ1.Selected.Rows
                    If grdrow.Cells("JOB_STATUS").Value = "E" Then
                        jobNoList.Add(grdrow.Cells("JOB_NO").Value)
                        grdrow.Cells("JOB_STATUS").Value = "Q"
                    End If
                Next

                If jobNoList.Count = 0 Then
                    MessageBox.Show("There were no Jobs with Status 'Error' to Re-Queue.", "Re-Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

                Try
                    ASCMAIN1.sql = "UPDATE WBTWEBQ1 Set JOB_STATUS = 'Q' WHERE JOB_STATUS = 'E' AND JOB_NO IN " _
                        & "('" & String.Join("', '", jobNoList.ToArray) & "')"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

                    MessageBox.Show("There were " & jobNoList.Count & " Job(s) Re-Queued.", "Re-Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)

                Catch ex As Exception
                    MessageBox.Show("The following Error occurred when Re-Queuing Jobs: " & ex.Message, "Re-Queue", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Load_Record()
                End Try

        End Select
    End Sub

#End Region

    Private Sub grdWBTWEBQ1_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdWBTWEBQ1.InitializeRow
        If e.Row.Cells("JOB_STATUS").Value & String.Empty = "E" Then
            e.Row.Appearance.BackColor = Drawing.Color.Red
        Else
            e.Row.Appearance.BackColor = Drawing.Color.White
        End If
    End Sub

End Class