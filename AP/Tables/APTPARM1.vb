Public Class APTPARM1


    Overrides Sub Show_Record_Special()
        Dim chkctl As UltraWinEditors.UltraCheckEditor = Absx1.chkFor("AP_PARM_BANK_METHOD")
        Set_Read_Only_for_ctl(chkctl, False)
        If EntryMode = "Edit" Then
            If chkctl.Checked Then
                Set_Read_Only_for_ctl(chkctl, True)
            End If
        End If

    End Sub


    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)


        Dim chkctl As UltraWinEditors.UltraCheckEditor = Absx1.chkFor("AP_PARM_BANK_METHOD")

        If tf Then
            If EntryMode = "Edit" Or EntryMode = "New" Then
                Set_Read_Only_for_ctl(chkctl, False)
            End If
            If EntryMode = "Edit" Then
                If chkctl.Checked Then
                    Set_Read_Only_for_ctl(chkctl, True)
                End If
            End If
        End If

    End Sub

End Class