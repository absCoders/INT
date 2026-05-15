Public Class TAFTESTX

    Dim oWB As SpreadsheetGear.IWorkbook
    'oWB = SpreadsheetGear.Factory.GetWorkbook()
    'For i As Integer = oWB.Worksheets.Count To 2 Step -1
    '    oWB.Worksheets(i).Delete()

    'Next i
    Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing
    Dim rangeCopyFrom As SpreadsheetGear.IRange = Nothing
    Dim rangePaste_To As SpreadsheetGear.IRange = Nothing

    Dim CUST_CODE As String
    Dim FOLDER As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\"

    Private Sub SOFFORM1_Load(sender As Object, e As System.EventArgs) Handles Me.Load

    End Sub
    Overrides Sub Prepare_for_View_Lookup_Special _
    (ByVal ctl As Windows.Forms.Control, _
     ByVal COLUMN_NAME As String, _
     Optional ByRef sql_where As String = "", _
     Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME
            '        Case "ORDR_FORM_CODE"
            '            sql_where = "ORDR_FORM_STATUS = 'A'"
        End Select
    End Sub

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            'Case "ORDR_FORM_CODE"
            '    If e.KeyCode = Windows.Forms.Keys.Enter Then
            '        Prepare_SOTFORM2()
            '    End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            'Case "ORDR_FORM_CODE"
            '    Prepare_SOTFORM2()
        End Select
    End Sub
#End Region

    Sub LORD()


        If dst.Tables.Contains("TATXLSX1") Then
            dst.Tables("TATXLSX1").Rows.Clear()
            dst.Tables("TATXLSX2").Rows.Clear()
        Else
            With dst
                Create_TDA(.Tables.Add, "TATXLSX1", "*")
                Create_TDA(.Tables.Add, "TATXLSX2", "*")
            End With
        End If

        grd1.DisplayLayout.Bands(0).Summaries.Clear()
        grd2.DisplayLayout.Bands(0).Summaries.Clear()
        grd1.DataSource = Nothing
        grd2.DataSource = Nothing
        grd1.DataSource = dst.Tables("TATXLSX1")
        grd2.DataSource = dst.Tables("TATXLSX2")
        ASCMAIN1.grdInitializeLayout(grd1)
        ASCMAIN1.grdInitializeLayout(grd2)

        Create_Summary(grd1, "FILE_ID", "Count")
        Create_Summary(grd2, "DEPT", "Count")
        Create_Summary(grd2, New String() {"USLSTY", "USLSLY", "DSLSTY", "DSLSLY", "UONHTY", "UONHLY"})

        Dim FILEs As New List(Of String)
        FILEs.Add("OSCAR, DUNHILL & ANNA SUI SALES BY WEEK 2015")
        FILEs.Add("OSCAR, DUNHILL & ANNA SUI SALES BY WEEK 2016")

        For Each FILENAME As String In FILEs
            oWB = SpreadsheetGear.Factory.GetWorkbook(FOLDER & FILENAME & ".XLSX")
            For i As Integer = 0 To oWB.Worksheets.Count - 1
                oSheet = oWB.Worksheets(i)
                Dim N As String = oSheet.Name
                Debug.Print(N & ":" & FILENAME)
                If N.StartsWith("Total") Then
                    ' do nothing
                Else
                    Dim W As String = ""
                    If InStr(N, " ") = 4 Then
                        W = Mid(N, 1, 1)
                    Else
                        W = Mid(N, 1, 2)
                    End If
                    Dim DTE1 As Date = Mid(N, Len(W) + 1, 9)
                    Dim DTE2 As Date = Mid(N, Len(W) + 1 + 9 + 3, 9)
                    If Len(W) = 1 Then W = "0" & W

                    Dim row1 As DataRow = dst.Tables("TATXLSX1").NewRow
                    row1.Item("CUST") = CUST_CODE
                    row1.Item("FILE_ID") = Mid(FILENAME, Len(FILENAME) - 3, 4)
                    row1.Item("WEEK") = W
                    row1.Item("DATE1") = DTE1
                    row1.Item("DATE2") = DTE2
                    dst.Tables("TATXLSX1").Rows.Add(row1)

                    Dim R As Integer = 9 ' 0 BASED ID FOR ROW 10
                    Do While oSheet.Cells(R, 0).Value & "" <> ""
                        Dim row2 As DataRow = dst.Tables("TATXLSX2").NewRow
                        row2.Item("CUST") = CUST_CODE
                        row2.Item("FILE_ID") = Mid(FILENAME, Len(FILENAME) - 3, 4)
                        row2.Item("WEEK") = W
                        row2.Item("DEPT") = oSheet.Cells(R, 2).Value & ""
                        row2.Item("UPC") = oSheet.Cells(R, 16).Value & ""
                        row2.Item("ITEM") = oSheet.Cells(R, 19).Value & ""
                        row2.Item("STORE") = oSheet.Cells(R, 20).Value & ""
                        row2.Item("STORE_NAME") = oSheet.Cells(R, 21).Value & ""
                        row2.Item("RETAIL") = Val(oSheet.Cells(R, 22).Value & "")
                        row2.Item("USLSTY") = Val(oSheet.Cells(R, 23).Value & "")
                        row2.Item("USLSLY") = Val(oSheet.Cells(R, 24).Value & "")
                        row2.Item("DSLSTY") = Val(oSheet.Cells(R, 25).Value & "")
                        row2.Item("DSLSLY") = Val(oSheet.Cells(R, 26).Value & "")
                        row2.Item("UONHTY") = Val(oSheet.Cells(R, 39).Value & "")
                        row2.Item("UONHLY") = Val(oSheet.Cells(R, 40).Value & "")
                        dst.Tables("TATXLSX2").Rows.Add(row2)
                        R += 1
                    Loop
                End If
            Next
        Next

    End Sub

    Sub BELK()


        If dst.Tables.Contains("TATXLSX3") Then
            dst.Tables("TATXLSX3").Rows.Clear()
            dst.Tables("TATXLSX4").Rows.Clear()
        Else
            With dst
                Create_TDA(.Tables.Add, "TATXLSX3", "*")
                Create_TDA(.Tables.Add, "TATXLSX4", "*")
            End With
        End If

        grd1.DisplayLayout.Bands(0).Summaries.Clear()
        grd2.DisplayLayout.Bands(0).Summaries.Clear()
        grd1.DataSource = Nothing
        grd2.DataSource = Nothing
        grd1.DataSource = dst.Tables("TATXLSX3")
        grd2.DataSource = dst.Tables("TATXLSX4")

        ASCMAIN1.grdInitializeLayout(grd1)
        ASCMAIN1.grdInitializeLayout(grd2)

        Create_Summary(grd1, "FILE_ID", "Count")
        Create_Summary(grd2, "ITEM", "Count")
        Create_Summary(grd2, New String() {"USLSTY", "USLSLY", "DSLSTY", "DSLSLY"})


        Dim FILEs As New List(Of String)
        FILEs.Add("Oscar Sales by Style Weekly Recap 10.19.16")

        For Each FILENAME As String In FILEs
            oWB = SpreadsheetGear.Factory.GetWorkbook(FOLDER & FILENAME & ".XLSX")
            For i As Integer = 0 To oWB.Worksheets.Count - 1
                oSheet = oWB.Worksheets(i)
                Dim N As String = oSheet.Name
                Debug.Print(N & ":" & FILENAME)

                Dim SEASON_CODE As String = Mid(N, Len(N) - 3, 4) & Mid(N, 1, 1)

                Dim row1 As DataRow = dst.Tables("TATXLSX3").NewRow
                row1.Item("CUST") = CUST_CODE
                row1.Item("FILE_ID") = Mid(FILENAME, Len(FILENAME) - 7, 8)
                row1.Item("SEASON") = SEASON_CODE
                dst.Tables("TATXLSX3").Rows.Add(row1)

                Dim R As Integer = 6 ' 0 BASED ID FOR ROW 7
                Do While oSheet.Cells(R, 0).Value & "" <> ""
                    Dim YYYYMM As String = oSheet.Cells(R, 0).Value & ""
                    R += 1

                    Do While oSheet.Cells(R, 2).Value & "" <> ""
                        Dim DX As String = oSheet.Cells(R, 2).Value & ""
                        Dim D As Date = Mid(DX, 7, 10)

                        Dim C As Integer = 4
                        Do While oSheet.Cells(2, C).Value & "" <> ""
                            Dim row2 As DataRow = dst.Tables("TATXLSX4").NewRow
                            row2.Item("CUST") = CUST_CODE
                            row2.Item("FILE_ID") = Mid(FILENAME, Len(FILENAME) - 7, 8)
                            row2.Item("SEASON") = SEASON_CODE
                            row2.Item("ITEM") = oSheet.Cells(2, C).Value & ""
                            row2.Item("DESCRIPTION") = oSheet.Cells(3, C).Value & ""

                            row2.Item("USLSTY") = Val(oSheet.Cells(R, C + 0).Value & "")
                            row2.Item("USLSLY") = Val(oSheet.Cells(R, C + 1).Value & "")
                            row2.Item("DSLSTY") = Val(oSheet.Cells(R, C + 3).Value & "")
                            row2.Item("DSLSLY") = Val(oSheet.Cells(R, C + 4).Value & "")

                            dst.Tables("TATXLSX4").Rows.Add(row2)
                            C += 6
                        Loop

                        R += 1
                    Loop
                Loop
            Next
        Next

    End Sub

    Private Sub cmdUpdate_Click(sender As Object, e As EventArgs) Handles cmdUpdate.Click
        If CUST_CODE = "LORD" Then
            Update_Record_TDA("TATXLSX1")
            Update_Record_TDA("TATXLSX2")
        Else
            Update_Record_TDA("TATXLSX3")
            Update_Record_TDA("TATXLSX4")
        End If
    End Sub


    Private Sub cmdCancel_Click(sender As Object, e As EventArgs) Handles cmdCancel.Click
        Me.Close()
        ' CHANGE TO CODE
    End Sub

    Private Sub cmdLoad_Click(sender As Object, e As EventArgs) Handles cmdLoad.Click
        CUST_CODE = Absx1.txtFor("CUST_CODE").Text
        FOLDER = My.Computer.FileSystem.SpecialDirectories.Desktop & "\"

        If CUST_CODE = "LORD" Then LORD()
        If CUST_CODE = "BELK" Or CUST_CODE = "BELK" Then BELK()
    End Sub
End Class