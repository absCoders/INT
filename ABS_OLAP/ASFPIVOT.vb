Public Class ASFPIVOT
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
        PivotGridControl1.DataSource = _dt
        PivotGridControl1.RetrieveFields()
    End Sub

    Private Sub cmdExcel_Click(sender As System.Object, e As System.EventArgs) Handles cmdExcel.Click
        PivotGridControl1.ExportToXls("c:\pivotGrid_output.xls")
        Dim FILENAME As String = "c:\pivotGrid_output.xls"
        Dim p As Process = Nothing
        Try
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                p = Process.Start(FILENAME)
                If p IsNot Nothing Then
                    p.Dispose()
                End If
            End If

        Catch ex As Exception

        Finally

        End Try
    End Sub
End Class