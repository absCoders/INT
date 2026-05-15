Public Class EDTPARM1
     
    Private Sub btbDevSynonyms_Click(sender As Object, e As EventArgs) Handles btbDevSynonyms.Click

        ASCMAIN1.Progress("Renaming Synonyms")
        ASCMAIN1.sql = "Select SYNONYM_NAME from USER_SYNONYMS"
        For Each ROW As DataRow In ASCDATA1.GetDataTable.Select("", "SYNONYM_NAME")
            Dim SYNONYM_NAME As String = ROW.Item("SYNONYM_NAME")
            ASCMAIN1.Progress("-", SYNONYM_NAME)
            ASCMAIN1.sql = "Drop Synonym " & SYNONYM_NAME
            ASCDATA1.ExecuteSQL()
            'ASCMAIN1.sql = "Create Synonym " & SYNONYM_NAME & " for GEN" & ASCMAIN1.DBS_COMPANY & "." & SYNONYM_NAME
            ASCMAIN1.sql = "Create Synonym " & Chr(34) & SYNONYM_NAME & Chr(34) & " for GEN" & ASCMAIN1.DBS_COMPANY & "." & Chr(34) & SYNONYM_NAME & Chr(34)
            If ASCMAIN1.CLIENT = "XXX" Then ASCMAIN1.sql = "Create Synonym " & Chr(34) & SYNONYM_NAME & Chr(34) & " for GENINT" & "." & Chr(34) & SYNONYM_NAME & Chr(34)

            ASCDATA1.ExecuteSQL()
        Next

        ASCMAIN1.Progress("")
    End Sub

    Private Sub EDTPARM1_Load(sender As Object, e As EventArgs) Handles Me.Load
        btbDevSynonyms.Visible = (ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz")
    End Sub
End Class