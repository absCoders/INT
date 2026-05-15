<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFOLAP1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.C1OlapPage1 = New C1.Win.Olap.C1OlapPage()
        CType(Me.C1OlapPage1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'C1OlapPage1
        '
        Me.C1OlapPage1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.C1OlapPage1.Location = New System.Drawing.Point(0, 0)
        Me.C1OlapPage1.Margin = New System.Windows.Forms.Padding(2)
        Me.C1OlapPage1.Name = "C1OlapPage1"
        Me.C1OlapPage1.Size = New System.Drawing.Size(1013, 458)
        Me.C1OlapPage1.TabIndex = 0
        '
        'ASFOLAP1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1013, 458)
        Me.Controls.Add(Me.C1OlapPage1)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ASFOLAP1"
        Me.Text = "Form1"
        CType(Me.C1OlapPage1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents C1OlapPage1 As C1.Win.Olap.C1OlapPage

End Class
