<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GLRLIST1
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFSRPTM
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
        Me.UltraCheckEditor1 = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.UltraCheckEditor2 = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.UltraCheckEditor3 = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.UltraCheckEditor4 = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.UltraTabPageControl2.SuspendLayout()
        CType(Me.UltraTabControl1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.UltraTabControl1.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SplitContainer5.Panel1.SuspendLayout()
        Me.SplitContainer5.SuspendLayout()
        CType(Me.tblASTSPRF1_clone, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTROPT1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLE, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLF, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLJ, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraTabControl1
        '
        Me.UltraTabControl1.TabPageMargins.ForceSerialization = True
        '
        'SplitContainer1
        '
        '
        'SplitContainer5
        '
        '
        'SplitContainer5.Panel1
        '
        Me.SplitContainer5.Panel1.Controls.Add(Me.UltraCheckEditor4)
        Me.SplitContainer5.Panel1.Controls.Add(Me.UltraCheckEditor3)
        Me.SplitContainer5.Panel1.Controls.Add(Me.UltraCheckEditor1)
        Me.SplitContainer5.Panel1.Controls.Add(Me.UltraCheckEditor2)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.ForceSerialization = True
        Me.UltraExplorerBar1.Margins.ForceSerialization = True
        '
        'UltraCheckEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraCheckEditor1, "ACCT_CODE")
        Me.UltraCheckEditor1.Checked = True
        Me.UltraCheckEditor1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.UltraCheckEditor1.Location = New System.Drawing.Point(10, 36)
        Me.UltraCheckEditor1.Name = "UltraCheckEditor1"
        Me.UltraCheckEditor1.Size = New System.Drawing.Size(198, 20)
        Me.UltraCheckEditor1.TabIndex = 0
        Me.UltraCheckEditor1.Text = "Print Chart of Accounts"
        '
        'UltraCheckEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraCheckEditor2, "SEG2_CODE")
        Me.UltraCheckEditor2.Location = New System.Drawing.Point(10, 62)
        Me.UltraCheckEditor2.Name = "UltraCheckEditor2"
        Me.UltraCheckEditor2.Size = New System.Drawing.Size(198, 20)
        Me.UltraCheckEditor2.TabIndex = 1
        Me.UltraCheckEditor2.Text = "Print"
        '
        'UltraCheckEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraCheckEditor3, "SEG3_CODE")
        Me.UltraCheckEditor3.Location = New System.Drawing.Point(10, 88)
        Me.UltraCheckEditor3.Name = "UltraCheckEditor3"
        Me.UltraCheckEditor3.Size = New System.Drawing.Size(198, 20)
        Me.UltraCheckEditor3.TabIndex = 2
        Me.UltraCheckEditor3.Text = "Print"
        '
        'UltraCheckEditor4
        '
        Me.Absx1.SetABSColumnName(Me.UltraCheckEditor4, "SEG4_CODE")
        Me.UltraCheckEditor4.Location = New System.Drawing.Point(10, 114)
        Me.UltraCheckEditor4.Name = "UltraCheckEditor4"
        Me.UltraCheckEditor4.Size = New System.Drawing.Size(198, 20)
        Me.UltraCheckEditor4.TabIndex = 3
        Me.UltraCheckEditor4.Text = "Print"
        '
        'GLRLIST1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(990, 574)
        Me.Name = "GLRLIST1"
        Me.Text = "GLRLIST1"
        Me.UltraTabPageControl2.ResumeLayout(False)
        CType(Me.UltraTabControl1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.UltraTabControl1.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.SplitContainer5.Panel1.ResumeLayout(False)
        Me.SplitContainer5.ResumeLayout(False)
        CType(Me.tblASTSPRF1_clone, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTROPT1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLE, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLF, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLJ, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraCheckEditor4 As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents UltraCheckEditor3 As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents UltraCheckEditor2 As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents UltraCheckEditor1 As Infragistics.Win.UltraWinEditors.UltraCheckEditor
End Class
