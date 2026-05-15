Public Class CheckEditorDataFilter
    Implements Infragistics.Win.IEditorDataFilter
    Public Function Convert(ByVal args As Infragistics.Win.EditorDataFilterConvertArgs) As Object Implements Infragistics.Win.IEditorDataFilter.Convert

        If args.Value & "" = "Received" Then Return "0"

        Select Case args.Direction
            Case Infragistics.Win.ConversionDirection.EditorToOwner
                args.Handled = True
                Select Case CType(args.Value, CheckState)
                    Case CheckState.Checked
                        Return "1"
                    Case CheckState.Unchecked
                        Return "0"
                    Case CheckState.Indeterminate
                        Return "0"
                        'Return String.Empty
                    Case Else
                        Return "0"
                        'Return CheckState.Indeterminate
                End Select
            Case Infragistics.Win.ConversionDirection.OwnerToEditor
                args.Handled = True
                If args.Value & "" = "1" Then
                    'Return 1
                    Return CheckState.Checked
                ElseIf args.Value & "" = "0" Then
                    'Return 0
                    Return CheckState.Unchecked
                Else
                    Return CheckState.Unchecked
                    'Return CheckState.Indeterminate
                End If
            Case Else
                Return "0"
                'Return CheckState.Indeterminate
        End Select
    End Function

End Class
