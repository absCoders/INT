Imports ABSolution

Public Class ASFTASK1

    Private wktable As String = String.Empty

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst

            MyBase.Create_TDA(.Tables.Add, "ASTTASK1", "*")
            MyBase.Create_TDA(.Tables.Add, "ASTTASK2", "*")

            .Relations.Add("ASTTASK1_ASTTASK2", dst.Tables("ASTTASK1").Columns("TASK_NO"), dst.Tables("ASTTASK2").Columns("TASK_NO"))
        End With

        grdASTTASK1.DataSource = dst.Tables("ASTTASK1")
        Create_Summary(grdASTTASK1, "TASK_NO", "Count")
        Create_Summary(grdASTTASK1, "TASK_LNO", "Count", "ASTTASK1_ASTTASK2")

        dteStart.MaxDate = DateTime.Now
        dteEnd.MaxDate = DateTime.Now

        dteStart.MinDate = DateAdd(DateInterval.Year, -1, DateTime.Now)
        dteEnd.MinDate = DateAdd(DateInterval.Year, -1, DateTime.Now)

        dteEnd.Value = dteEnd.MaxDate
        dteStart.Value = dteStart.MaxDate

        wktable = ASCMAIN1.Temp_Table("select task_no from asttask1 where rownum < 1")

    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        MyBase.EMsg = ""

        Select Case eItemKey

            Case "Load"
                If Not IsDate(dteStart.DateTime) Then
                    EMsg &= vbCr & "Invalid Start Date"
                ElseIf Not IsDate(dteEnd.DateTime) Then
                    EMsg &= vbCr & "Invalid End Date"
                ElseIf dteEnd.DateTime < dteStart.DateTime Then
                    EMsg &= vbCr & "End Date must bet greater equal Start Date"
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
                Me.Load_Record()
                Me.Mode_Settings(True)

            Case "Done"
                Me.Mode_Settings(False)

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal Mode_Desc As String = "")

        MyBase.Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Done").Settings.Enabled = iScreenMode
            End With
        End If

        MyBase.Set_Read_Only(UltraGroupBox1, ScreenMode)

        If ScreenMode Then

        Else
            Me.Clear_Record()
        End If

        grdASTTASK1.Visible = ScreenMode

    End Sub

    Private Sub Clear_Record()
        dst.EnforceConstraints = False
        dst.Tables("ASTTASK1").Rows.Clear()
        dst.Tables("ASTTASK2").Rows.Clear()

        dst.EnforceConstraints = True
    End Sub

    Private Sub Load_Record()

        ASCMAIN1.Progress("Now Loading")

        MyBase.EnforceConstraints(False)

        ASCDATA1.ExecuteSQL("truncate table " & wktable)
        ASCMAIN1.sql = "insert into " & wktable & " SELECT TASK_NO FROM ASTTASK1 WHERE TRUNC(START_TIME) BETWEEN '" & dteStart.DateTime.ToString("dd-MMM-yyyy") & "' AND '" & dteEnd.DateTime.ToString("dd-MMM-yyyy") & "'"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)

        Fill_Records("ASTTASK1", String.Empty, True, "SELECT * FROM ASTTASK1 WHERE TASK_NO IN (SELECT TASK_NO FROM " & wktable & ")")
        Fill_Records("ASTTASK2", String.Empty, True, "SELECT * FROM ASTTASK2 WHERE TASK_NO IN (SELECT TASK_NO FROM " & wktable & ")")

        Sort_grdColumns(grdASTTASK1, "TASK_NO", False, 0)
        Sort_grdColumns(grdASTTASK1, "TASK_LNO", False, 1)

        MyBase.EnforceConstraints(True)

        ASCMAIN1.Progress("")
    End Sub

    Private Sub Update_Record()

        Try
            MyBase.BeginTrans()


            MyBase.CommitTrans("Update Complete")

        Catch ex As Exception
            Rollback(ex.Message)
        End Try
    End Sub

#End Region


End Class