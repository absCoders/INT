<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ARRCUST1
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
        Me.UltraTabPageControl2.SuspendLayout()
        CType(Me.tblASTSPRF1_clone, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTROPT1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLB, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLC, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTDSQLJ, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraTabPageControl2
        '
        Me.Absx1.SetABSColumnName(Me.UltraTabPageControl2, "")
        Me.Absx1.SetABSParentColumnName(Me.UltraTabPageControl2, "")
        Me.Absx1.SetABSTableName(Me.UltraTabPageControl2, "")
        Me.Absx1.SetABSViewName(Me.UltraTabPageControl2, "")
        Me.UltraTabPageControl2.Controls.Add(Me.UltraCheckEditor1)
        Me.UltraTabPageControl2.Location = New System.Drawing.Point(2, 23)
        '
        'UltraExplorerBar1
        '
        Me.Absx1.SetABSColumnName(Me.UltraExplorerBar1, "")
        Me.Absx1.SetABSParentColumnName(Me.UltraExplorerBar1, "")
        Me.Absx1.SetABSTableName(Me.UltraExplorerBar1, "")
        Me.Absx1.SetABSViewName(Me.UltraExplorerBar1, "")
        Me.UltraExplorerBar1.GroupSettings.ForceSerialization = True
        Me.UltraExplorerBar1.ItemSettings.ForceSerialization = True
        Me.UltraExplorerBar1.Margins.ForceSerialization = True
        '
        'UltraCheckEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraCheckEditor1, "ACTIVE_ONLY")
        Me.Absx1.SetABSParentColumnName(Me.UltraCheckEditor1, "")
        Me.Absx1.SetABSTableName(Me.UltraCheckEditor1, "")
        Me.Absx1.SetABSViewName(Me.UltraCheckEditor1, "")
        Me.UltraCheckEditor1.Location = New System.Drawing.Point(23, 39)
        Me.UltraCheckEditor1.Name = "UltraCheckEditor1"
        Me.UltraCheckEditor1.Size = New System.Drawing.Size(177, 20)
        Me.UltraCheckEditor1.TabIndex = 0
        Me.UltraCheckEditor1.Text = "Active Only"
        '
        'ARRCUST1
        '
        Me.Absx1.SetABSColumnName(Me, "")
        Me.Absx1.SetABSParentColumnName(Me, "")
        Me.Absx1.SetABSTableName(Me, "")
        Me.Absx1.SetABSViewName(Me, "")
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(990, 574)
        Me.Name = "ARRCUST1"
        Me.Text = "ARRCUST1"
        Me.UltraTabPageControl2.ResumeLayout(False)
        CType(Me.tblASTSPRF1_clone, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTROPT1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLB, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLC, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTDSQLJ, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents UltraCheckEditor1 As Infragistics.Win.UltraWinEditors.UltraCheckEditor
End Class
