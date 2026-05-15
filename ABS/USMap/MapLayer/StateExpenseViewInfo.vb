
Imports System


Namespace ChartSamplesExplorerVB.Customization
    '/ <summary>
    '/ This class simply serves as a simple placeholder for our state dummy data
    '/ </summary>

    Public Class StateExpenseViewInfo
#Region "Private Member Variables"
        Private _State As String = ""
        Private _Amount As Double = 0.0
#End Region

#Region "Constructors"

        Public Sub New(ByVal state As String, ByVal amount As Double, ByVal category As String)
            _State = state
            _Amount = amount
        End Sub 'New

#End Region

#Region "Public Properties"

        Public Property State() As String
            Get
                Return _State
            End Get

            Set(ByVal Value As String)
                _State = Value
            End Set
        End Property


        Public Property Amount() As Double
            Get
                Return _Amount
            End Get

            Set(ByVal Value As Double)
                _Amount = Value
            End Set
        End Property

#End Region
    End Class 'StateExpenseViewInfo
End Namespace 'ChartSamplesExplorerCS.Customization