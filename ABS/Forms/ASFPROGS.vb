Public NotInheritable Class ASFPROGS

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).

    Private Delegate Sub UpdateProgressDelegate(ByVal msg1 As String, ByVal msg2 As String, ByVal msg3 As String, ByVal percentage As Integer)

    Dim defaultMsg1 As String = "Now Loading:"

    Public Sub New(ByVal msg1 As String, ByVal msg2 As String, ByVal msg3 As String _
                   , ByVal bounds As System.Drawing.Rectangle, ByVal showProgressBar As Boolean)

        InitializeComponent()

        If msg1 = "" Then
            UltraLabel1.Text = defaultMsg1
        Else
            UltraLabel1.Text = msg1
        End If

        UltraLabel2.Text = msg2
        UltraLabel3.Text = msg3

        Me.ProgressBar1.Visible = showProgressBar

        If Not showProgressBar Then
            UltraGroupBox1.Height = UltraLabel3.Bottom + 10
            Me.Height = UltraGroupBox1.Bottom + 15
        End If

        Me.Location = New Point(bounds.Right - Me.Width - 30, bounds.Bottom - Me.Height - 30)
    End Sub

    Private Sub ASFPROGS_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Public Sub EndForm(ByVal o As System.Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    Public Sub UpdateProgress(ByVal msg1 As String, ByVal msg2 As String, ByVal msg3 As String, ByVal percentage As Integer)
        If msg1 = "" Then
            msg1 = Me.UltraLabel1.Text
        End If
        If msg2 = "" Then
            msg2 = Me.UltraLabel2.Text
        End If
        If msg3 = "" Then
            msg3 = Me.UltraLabel3.Text
        End If
        If Me.InvokeRequired Then
            Me.Invoke(New UpdateProgressDelegate(AddressOf UpdateProgress), New Object() {msg1, msg2, msg3, percentage})
        Else
            Me.UltraLabel1.Text = msg1
            Me.UltraLabel2.Text = msg2
            Me.UltraLabel3.Text = msg3
            If percentage >= Me.ProgressBar1.Minimum AndAlso percentage <= Me.ProgressBar1.Maximum Then
                Me.ProgressBar1.Value = percentage
            End If
        End If
    End Sub

End Class
