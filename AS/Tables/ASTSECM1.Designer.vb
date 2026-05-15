<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASTSECM1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFCODEM
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
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor2 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.Panel1.SuspendLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.UltraLabel2)
        Me.Panel1.Controls.Add(Me.UltraLabel1)
        Me.Panel1.Controls.Add(Me.UltraTextEditor2)
        Me.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel2, 0)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.ForceSerialization = True
        Me.UltraExplorerBar1.ItemSettings.ForceSerialization = True
        Me.UltraExplorerBar1.Margins.ForceSerialization = True
        '
        'UltraLabel1
        '
        Appearance2.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel1.Appearance = Appearance2
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(16, 76)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(99, 18)
        Me.UltraLabel1.TabIndex = 14
        Me.UltraLabel1.Text = "Security Code"
        '
        'UltraTextEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor2, "SECURITY_DESC")
        Me.UltraTextEditor2.Location = New System.Drawing.Point(130, 106)
        Me.UltraTextEditor2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor2.Name = "UltraTextEditor2"
        Me.UltraTextEditor2.Size = New System.Drawing.Size(372, 25)
        Me.UltraTextEditor2.TabIndex = 13
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "SECURITY_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(130, 73)
        Me.UltraTextEditor1.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(63, 25)
        Me.UltraTextEditor1.TabIndex = 12
        '
        'UltraLabel2
        '
        Appearance1.BackColor = System.Drawing.Color.Transparent
        Me.UltraLabel2.Appearance = Appearance1
        Me.UltraLabel2.AutoSize = True
        Me.UltraLabel2.Location = New System.Drawing.Point(16, 109)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(80, 18)
        Me.UltraLabel2.TabIndex = 15
        Me.UltraLabel2.Text = "Description"
        '
        'ASTSECM1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "ASTSECM1"
        Me.Text = "ASTSECM1"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor2 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
End Class
