Public Class ASTHLPV1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "ASTHLPV2", "*", 1)
            Create_TDA(.Tables.Add, "ASTHLPV3", "*", 1)
        End With

        grdASTHLPV2.DataSource = dst.Tables("ASTHLPV2")
        grdASTHLPV3.DataSource = dst.Tables("ASTHLPV3")
    End Sub

    Private Sub grdASTHLPV2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTHLPV2.BeforeRowUpdate
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            For Each gc As UltraWinGrid.UltraGridColumn In .DisplayLayout.Bands(0).Columns
                If htbkey_COLUMN_NAMEs.Contains(gc.Key) Then
                    .ActiveRow.Cells(gc.Key).Value = DirectCast(htbkey_COLUMN_NAMEs(gc.Key), UltraWinEditors.UltraTextEditor).Text
                End If
            Next
        End With
    End Sub

    Private Sub grdASTHLPV3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdASTHLPV3.BeforeRowUpdate
        With DirectCast(sender, UltraWinGrid.UltraGrid)
            For Each gc As UltraWinGrid.UltraGridColumn In .DisplayLayout.Bands(0).Columns
                If htbkey_COLUMN_NAMEs.Contains(gc.Key) Then
                    .ActiveRow.Cells(gc.Key).Value = DirectCast(htbkey_COLUMN_NAMEs(gc.Key), UltraWinEditors.UltraTextEditor).Text
                End If
            Next
        End With
    End Sub

#Region "Overrides"

    Overrides Sub Proceed_Update_Special_Pre()
        Call Update_Record_TDA("ASTHLPV2", "VIDEO_NO = '" & Absx1.txtFor("VIDEO_NO").Text & "'")
        Call Update_Record_TDA("ASTHLPV3", "VIDEO_NO = '" & Absx1.txtFor("VIDEO_NO").Text & "'")
    End Sub

    Overrides Sub Show_Record_Special()
        Call Fill_Records("ASTHLPV2", New String() {Absx1.txtFor("VIDEO_NO").Text})
        Call Fill_Records("ASTHLPV3", New String() {Absx1.txtFor("VIDEO_NO").Text})
    End Sub

    Overrides Sub Clear_Record_Special()

        If SELECTION_NO = 0 Then Exit Sub
        If ScreenMode Then
            dst.Tables("ASTHLPV2").Rows.Clear()
            dst.Tables("ASTHLPV3").Rows.Clear()
        End If

    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdASTHLPV2.Enabled = tf
        grdASTHLPV3.Enabled = tf
    End Sub
#End Region

    Private Sub btnPlay_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlay.Click
        Dim PATH As String = "C:\Documents and Settings\wjz\My Documents\Camtasia Studio\Import LGI Data"
        'Dim PATH As String = "C:\Documents and Settings\wjz\My Documents\Camtasia Studio\Import LGI Data\Import LGI Data_media"
        Dim FILE As String = "Import LGI Data.html" ' Absx1.txtFor("VIDEO_FILENAME").Text
        'Dim FILE As String = Absx1.txtFor("VIDEO_FILENAME").Text
        Dim VIDEO As String = PATH & "\" & FILE
        Show_Document(VIDEO)
    End Sub
End Class