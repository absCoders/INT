Public NotInheritable Class ASFFLOAD

    'TODO: This form can easily be set as the splash screen for the application by going to the "Application" tab
    '  of the Project Designer ("Properties" under the "Project" menu).

    Public Sub New(ByVal formName As String, ByVal bounds As System.Drawing.Rectangle, ByVal showGraphic As Boolean)
        InitializeComponent()
        UltraLabel2.Text = formName
        If Not showGraphic Then
            Me.PictureBox3.Visible = False
            UltraGroupBox1.Width = UltraLabel2.Right + 10
            Me.Width = UltraGroupBox1.Right + 15
        End If
        Me.Location = New Point(bounds.Right - Me.Width - 30, bounds.Bottom - Me.Height - 30)
    End Sub

    Private Sub ASFFLOAD_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Set up the dialog text at runtime according to the application's assembly information.  
        'Me.TopMost = False
    End Sub

    Public Sub EndForm(ByVal o As System.Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

End Class
