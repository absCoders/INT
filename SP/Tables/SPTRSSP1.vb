Public Class SPTRSSP1
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            ASCMAIN1.sql = "Select SPTRSSP2.*" & vbCrLf _
                & " from SPTRSSP2" & vbCrLf _
                & " where SPTRSSP2.RSSP_ID = :PARM1"
            Create_TDA(.Tables.Add, "SPTRSSP2", "**", 0, False, "V", 0)

        End With

        grdSPTRSSP2.DataSource = dst.Tables("SPTRSSP2")
    End Sub

#Region "Overrides"

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Dim RSSP_ID As String = rowASFBASE1.Item("RSSP_ID") & ""
        If RSSP_ID <> "" Then
            Fill_Records("SPTRSSP2", New String() {rowASFBASE1.Item("RSSP_ID") & ""})
            Sort_grdColumns(grdSPTRSSP2, "IMPORT_XNO")
        End If

        ASCMAIN1.AutoFitGridColumns(grdSPTRSSP2)

        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            For Each TABLE_NAME As String In New String() {"SPTRSSP2"}
                dst.Tables(TABLE_NAME).Rows.Clear()
            Next
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdSPTRSSP2.Visible = tf
    End Sub

#End Region
End Class