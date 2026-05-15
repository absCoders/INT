Namespace Exceptions

    Public Class ABSException
        Inherits ApplicationException

        Public Sub New()
            MyBase.New("An unexpected ABS Error has occurred")
            Me.LogException()
        End Sub

        Public Sub New(ByVal message As String)
            MyBase.New(message)
            Me.LogException()
        End Sub

        Public Sub New(ByVal message As String, ByVal innerException As Exception)
            MyBase.New(message, innerException)
            Me.LogException()
        End Sub

        Public Function GetExceptionMessageChain() As String
            Dim msgs As String = String.Empty
            Dim ex As Exception = Me
            Do While ex IsNot Nothing
                msgs &= ex.Message + Environment.NewLine
                ex = ex.InnerException
            Loop

            Return msgs

        End Function

        Private Sub LogException()
            Try
                ' write me.message application event log
                'Me.Message
                'Me.StackTrace
                'me.Data 
            Catch ex As Exception
                MessageBox.Show("Problem Logging the Exception Text")
            End Try
        End Sub
    End Class

    Public Class KeyColumnsRequiredException
        Inherits ABSException
        Public Sub New()
            MyBase.New("One or more key columns are required")
        End Sub
    End Class

    Public Class CustomerNotFoundException
        Inherits ABSException

        Private _customerID As String
        Public ReadOnly Property CustomerID()
            Get
                Return Me._customerID
            End Get
        End Property

        Public Sub New(ByVal customerID As String)
            MyBase.New(String.Format("Customer '{0}' not found", customerID))
            Me._customerID = customerID
        End Sub
    End Class

    Public Class ConfigFileLoadException
        Inherits ABSException

        Public Sub New()
            MyBase.New("The Config file could not be loaded")

        End Sub

        Public Sub New(ByVal innerException As Exception)
            MyBase.New("The Config file could not be loaded", innerException)

        End Sub

    End Class

    Public Class testing
        Private Sub LoadForm()

            Dim customerID As String = "x"
            Try
                Me.GetCustomer(customerID)
            Catch ex As Exceptions.CustomerNotFoundException
                MessageBox.Show(String.Format _
                ("Cannot find Customer '{0}'", ex.CustomerID))

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            Try
                Me.LoadConfig()
            Catch ex As ConfigFileLoadException
                MessageBox.Show(ex.GetExceptionMessageChain())
            Catch ex As ABSException

            Catch ex As Exception

                MessageBox.Show(ex.Message)
            End Try

        End Sub

        Private Sub GetCustomer(ByVal customerID As String)
            Throw New Exceptions.CustomerNotFoundException(customerID)
        End Sub

        Private Sub LoadConfig()

            ' test a few things, and decide can't load config
            If True Then

                Throw New ConfigFileLoadException()

            End If

            Try

                Dim sr As New System.IO.StreamReader("config.txt")

                sr.ReadToEnd()

                sr.Close()

            Catch ex As Exception

                Throw New ConfigFileLoadException(ex)

            End Try


        End Sub

    End Class
End Namespace
