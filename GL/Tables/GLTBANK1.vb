Public Class GLTBANK1

    Dim sqlGLTBANK2 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            sqlGLTBANK2 = "Select GLTBANK2.* from GLTBANK2"
            ASCMAIN1.sql = sqlGLTBANK2 _
            & "  where GLTBANK2.BANK_CODE = :PARM1 "
            Create_TDA(.Tables.Add, "GLTBANK2", "**", 0, False, "V")
        End With

        grdGLTBANK2.DataSource = dst.Tables("GLTBANK2")

        For Each gcol As UltraWinGrid.UltraGridColumn In grdGLTBANK2.DisplayLayout.Bands(0).Columns
            gcol.CellActivation = UltraWinGrid.Activation.NoEdit
        Next
    End Sub

#Region "Overrides"

    Overrides Sub Show_Record_Special()
        EnforceConstraints(False)
        Fill_Records("GLTBANK2", New String() {Absx1.txtFor("BANK_CODE").Text})
        Sort_grdColumns(grdGLTBANK2, "BANK_PYMT_METHOD")
        EnforceConstraints(True)
    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("GLTBANK2").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        grdGLTBANK2.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
        With grdGLTBANK2.DisplayLayout.Override
            .AllowAddNew = UltraWinGrid.AllowAddNew.No
            .AllowUpdate = DefaultableBoolean.False
            .AllowDelete = DefaultableBoolean.False
        End With
    End Sub

#End Region

End Class