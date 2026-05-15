<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFPROGS
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
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraGroupBox1 = New Infragistics.Win.Misc.UltraGroupBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel()
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraGroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'UltraLabel1
        '
        Appearance5.ForeColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(66, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.UltraLabel1.Appearance = Appearance5
        Me.UltraLabel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel1.Location = New System.Drawing.Point(4, 11)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Padding = New System.Drawing.Size(2, 0)
        Me.UltraLabel1.Size = New System.Drawing.Size(303, 25)
        Me.UltraLabel1.TabIndex = 4
        Me.UltraLabel1.Text = "Now Loading:"
        '
        'UltraLabel2
        '
        Appearance1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(66, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.UltraLabel2.Appearance = Appearance1
        Me.UltraLabel2.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel2.Location = New System.Drawing.Point(4, 37)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(305, 31)
        Me.UltraLabel2.TabIndex = 5
        Me.UltraLabel2.Text = "Form Name"
        '
        'UltraGroupBox1
        '
        Appearance3.BorderColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(66, Byte), Integer), CType(CType(83, Byte), Integer))
        Appearance3.BorderColor2 = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(66, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.UltraGroupBox1.Appearance = Appearance3
        Me.UltraGroupBox1.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rounded
        Appearance4.BorderColor = System.Drawing.Color.FromArgb(CType(CType(1, Byte), Integer), CType(CType(127, Byte), Integer), CType(CType(183, Byte), Integer))
        Me.UltraGroupBox1.ContentAreaAppearance = Appearance4
        Me.UltraGroupBox1.Controls.Add(Me.ProgressBar1)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel3)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel2)
        Me.UltraGroupBox1.Controls.Add(Me.UltraLabel1)
        Me.UltraGroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.UltraGroupBox1.Name = "UltraGroupBox1"
        Me.UltraGroupBox1.Size = New System.Drawing.Size(317, 124)
        Me.UltraGroupBox1.TabIndex = 8
        '
        'ProgressBar1
        '
        Me.ProgressBar1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(66, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.ProgressBar1.Location = New System.Drawing.Point(6, 92)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(305, 23)
        Me.ProgressBar1.TabIndex = 8
        Me.ProgressBar1.Visible = False
        '
        'UltraLabel3
        '
        Appearance2.ForeColor = System.Drawing.Color.FromArgb(CType(CType(10, Byte), Integer), CType(CType(66, Byte), Integer), CType(CType(83, Byte), Integer))
        Appearance2.TextVAlignAsString = "Bottom"
        Me.UltraLabel3.Appearance = Appearance2
        Me.UltraLabel3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.UltraLabel3.Location = New System.Drawing.Point(4, 65)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Padding = New System.Drawing.Size(2, 0)
        Me.UltraLabel3.Size = New System.Drawing.Size(303, 21)
        Me.UltraLabel3.TabIndex = 7
        Me.UltraLabel3.Text = "Step 1 of 10"
        '
        'ASFPROGS
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.MintCream
        Me.ClientSize = New System.Drawing.Size(343, 149)
        Me.ControlBox = False
        Me.Controls.Add(Me.UltraGroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Location = New System.Drawing.Point(400, 400)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ASFPROGS"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        CType(Me.UltraGroupBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraGroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraGroupBox1 As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar

End Class
