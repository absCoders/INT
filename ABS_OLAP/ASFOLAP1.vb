Public Class ASFOLAP1
    Dim _dt As DataTable
    Dim _title As String

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Sub New(dt As DataTable, title As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        _dt = dt
        _title = title
    End Sub

    Private Sub ASFOLAP1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.Text = _title

        'Dim olap As C1.Olap.C1OlapEngine = C1OlapPage1.OlapEngine
        'olap.BeginUpdate()
        'C1OlapPage1.OlapEngine.Fields.MaxItems = 5
        'olap.EndUpdate()

        C1OlapPage1.DataSource = _dt
        C1OlapPage1.OlapEngine.ValueFields.MaxItems = 5
    End Sub

    Private Sub C1OlapPage1_Load(sender As System.Object, e As System.EventArgs) Handles C1OlapPage1.Load

    End Sub
End Class
