Public Class WHCFACT1
    Public Shared Function CreateWhcClass(ByVal className As String, ByVal ABSEnvironment As ABSEnvironment) As WHCRF000
        Select Case className
            Case "WHCRF001"
                Return New WHCRF001(ABSEnvironment)
            Case "WHCRF002"
                Return New WHCRF002(ABSEnvironment)
        End Select
        Return Nothing
    End Function
End Class