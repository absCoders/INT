Public Class ASCTTIP1

    Implements IRenderLabel

    Public Sub New()

    End Sub

    Public Overloads Function ToString(ByVal Context As System.Collections.Hashtable) As String Implements IRenderLabel.ToString
        ' THE VALUE RETURNED COULD PROBABLY BE EASILY CUSTOMIZED BY SENDING IN A PARAMETER IN THE CONSTRUCTOR AND USING THE FORMAT COMMAND OPTIONS
        Return Context("SERIES_LABEL") & vbCrLf & Context("ITEM_LABEL") & vbCrLf & CStr(Format(Val(Context("DATA_VALUE")), "#,##0"))
    End Function
End Class
