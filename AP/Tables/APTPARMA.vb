Public Class APTPARMA

    Private Sub txtInbound_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles txtInbound.EditorButtonClick
        Dim selectedDirectory As String = ShowFileDialog(txtInbound.Text)
        If selectedDirectory.Length > 0 Then
            txtInbound.Text = selectedDirectory
        End If
    End Sub

    Private Sub txtOutbound_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles txtOutbound.EditorButtonClick
        Dim selectedDirectory As String = ShowFileDialog(txtOutbound.Text)
        If selectedDirectory.Length > 0 Then
            txtOutbound.Text = selectedDirectory
        End If
    End Sub

    Private Sub txtError_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles txtError.EditorButtonClick
        Dim selectedDirectory As String = ShowFileDialog(txtError.Text)
        If selectedDirectory.Length > 0 Then
            txtError.Text = selectedDirectory
        End If
    End Sub

    Private Sub txtArchive_EditorButtonClick(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinEditors.EditorButtonEventArgs) Handles txtArchive.EditorButtonClick
        Dim selectedDirectory As String = ShowFileDialog(txtArchive.Text)
        If selectedDirectory.Length > 0 Then
            txtArchive.Text = selectedDirectory
        End If
    End Sub

    Private Function ShowFileDialog(ByVal startDirectory As String) As String

        Dim selectedDirectory As String = String.Empty
        Dim FolderBrowserDialog1 As New FolderBrowserDialog

        If My.Computer.FileSystem.DirectoryExists(startDirectory) Then
            FolderBrowserDialog1.SelectedPath = startDirectory
        End If

        Dim result As DialogResult = FolderBrowserDialog1.ShowDialog()

        If (result = DialogResult.OK) Then
            selectedDirectory = FolderBrowserDialog1.SelectedPath
            If Not My.Computer.FileSystem.DirectoryExists(selectedDirectory) Then
                selectedDirectory = String.Empty
            End If
        End If

        Return selectedDirectory
    End Function

    Public Overrides Sub Proceed_PreReq(ByVal eItemKey As String)
        MyBase.Proceed_PreReq(eItemKey)

        Select Case eItemKey
            Case "Cancel", "Done"
                For Each field As Infragistics.Win.UltraWinEditors.UltraTextEditor In New Infragistics.Win.UltraWinEditors.UltraTextEditor() {txtInbound, txtOutbound}
                    field.Appearance.BackColor = Drawing.Color.White
                Next
        End Select
    End Sub

    Public Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)
        MyBase.Proceed_PreReq_Special(eItemKey)

        Select Case eItemKey

            Case "Update", "Save"

                MyBase.Absx1.txtFor("AR_PARM_EDI_DRIVE_LETTER").Text = MyBase.Absx1.txtFor("AR_PARM_EDI_DRIVE_LETTER").Text.Trim.ToUpper
                MyBase.Absx1.txtFor("AR_PARM_EDI_PWD").Text = MyBase.Absx1.txtFor("AR_PARM_EDI_PWD").Text.Trim
                MyBase.Absx1.txtFor("AR_PARM_EDI_USER_ID").Text = MyBase.Absx1.txtFor("AR_PARM_EDI_USER_ID").Text.Trim

                Dim count As Integer = 0
                If MyBase.Absx1.txtFor("AR_PARM_EDI_USER_ID").TextLength > 0 Then count += 1
                If MyBase.Absx1.txtFor("AR_PARM_EDI_PWD").TextLength > 0 Then count += 1
                If MyBase.Absx1.txtFor("AR_PARM_EDI_DRIVE_LETTER").TextLength > 0 Then count += 1

                If count <> 0 AndAlso count <> 3 Then
                    EMsg &= vbCr & "You must provide none or all of the Out Security: User ID, Password and Drive Letter "
                    Exit Select
                End If

                If MyBase.Absx1.txtFor("AR_PARM_EDI_DRIVE_LETTER").TextLength > 0 _
                    AndAlso Not Char.IsLetter(MyBase.Absx1.txtFor("AR_PARM_EDI_DRIVE_LETTER").Text) Then
                    EMsg &= vbCr & "Out Security drive letter must be a letter"
                    Exit Select
                End If

                Dim invalidDriveLetters As String = "A,B,C,D,E,F,G,S,X,W"
                If invalidDriveLetters.Contains(MyBase.Absx1.txtFor("AR_PARM_EDI_DRIVE_LETTER").Text) Then
                    EMsg &= vbCr & "Out Security drive letter may not be one of the following: " & invalidDriveLetters
                    Exit Select
                End If


                For Each field As Infragistics.Win.UltraWinEditors.UltraTextEditor In New Infragistics.Win.UltraWinEditors.UltraTextEditor() {txtInbound, txtOutbound, txtError, txtArchive}
                    field.Text = field.Text.Trim
                    field.Appearance.BackColor = Drawing.Color.White

                    If field.Name = txtOutbound.Name _
                        AndAlso MyBase.Absx1.txtFor("AR_PARM_EDI_USER_ID").TextLength > 0 _
                        AndAlso MyBase.Absx1.txtFor("AR_PARM_NACHA_OUT_TRANS_DIR").Text.StartsWith("\\") Then

                        If MessageBox.Show("Do you want to try to connect to the provided 'Out Directory' using the provided 'Out Security' parameters?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.No Then
                            Continue For
                        End If

                        Dim UncPath As String = MyBase.Absx1.txtFor("AR_PARM_NACHA_OUT_TRANS_DIR").Text
                        Dim endpoint As Integer = InStr(3, UncPath, "\")
                        If endpoint > 0 Then
                            UncPath = UncPath.Substring(0, endpoint - 1)
                        End If
                        Dim ConnectionUsername As String = MyBase.Absx1.txtFor("AR_PARM_EDI_USER_ID").Text
                        Dim ConnectionPassword As String = MyBase.Absx1.txtFor("AR_PARM_EDI_PWD").Text
                        Dim MappedDriveLetter As String = MyBase.Absx1.txtFor("AR_PARM_EDI_DRIVE_LETTER").Text

                        '    If Not My.Computer.FileSystem.DirectoryExists(MyBase.Absx1.txtFor("AR_PARM_NACHA_OUT_TRANS_DIR").Text) Then
                        '        Using cMapNetworkDrive As New TAC.ASCMAPND()
                        '            With cMapNetworkDrive
                        '                .UncPath = UncPath
                        '                .DriveLetter = MappedDriveLetter
                        '                .Persistent = False
                        '                .ConnectionUsername = ConnectionUsername
                        '                .ConnectionPassword = ConnectionPassword
                        '                If Not .ConnectToServer() Then
                        '                    MessageBox.Show(.LastError, "Map Drive", MessageBoxButtons.OK)
                        '                Else
                        '                    MessageBox.Show("Successful mapping of the network to the drive.", "Map Drive", MessageBoxButtons.OK)
                        '                End If
                        '            End With
                        '        End Using
                        '    Else
                        '        MessageBox.Show("Successful connection to: " & MyBase.Absx1.txtFor("AR_PARM_NACHA_OUT_TRANS_DIR").Text, "Map Drive", MessageBoxButtons.OK)
                        '    End If
                        Continue For
                    End If

                    If field.TextLength > 0 Then
                        If Not My.Computer.FileSystem.DirectoryExists(field.Text) Then
                            EMsg &= vbCr & "All Directories must be a valid directory in the filesystem."
                            field.Appearance.BackColor = Drawing.Color.Yellow
                        End If
                    End If
                Next
        End Select
    End Sub

End Class