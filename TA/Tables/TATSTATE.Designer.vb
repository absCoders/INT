<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TATSTATE
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
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor2 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
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
        Me.Absx1.SetABSColumnName(Me.Panel1, "")
        Me.Absx1.SetABSParentColumnName(Me.Panel1, "")
        Me.Absx1.SetABSTableName(Me.Panel1, "")
        Me.Absx1.SetABSViewName(Me.Panel1, "")
        Me.Panel1.Controls.Add(Me.UltraLabel3)
        Me.Panel1.Controls.Add(Me.UltraTextEditor2)
        Me.Panel1.Controls.Add(Me.UltraLabel2)
        Me.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel3, 0)
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
        'UltraLabel3
        '
        Me.Absx1.SetABSColumnName(Me.UltraLabel3, "")
        Me.Absx1.SetABSParentColumnName(Me.UltraLabel3, "")
        Me.Absx1.SetABSTableName(Me.UltraLabel3, "")
        Me.Absx1.SetABSViewName(Me.UltraLabel3, "")
        Me.UltraLabel3.Location = New System.Drawing.Point(16, 44)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel3.TabIndex = 13
        Me.UltraLabel3.Text = "State Code"
        '
        'UltraTextEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor2, "STATE_NAME")
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor2, "")
        Me.Absx1.SetABSTableName(Me.UltraTextEditor2, "")
        Me.Absx1.SetABSViewName(Me.UltraTextEditor2, "")
        Me.UltraTextEditor2.Location = New System.Drawing.Point(122, 70)
        Me.UltraTextEditor2.Name = "UltraTextEditor2"
        Me.UltraTextEditor2.Size = New System.Drawing.Size(316, 25)
        Me.UltraTextEditor2.TabIndex = 12
        '
        'UltraLabel2
        '
        Me.Absx1.SetABSColumnName(Me.UltraLabel2, "")
        Me.Absx1.SetABSParentColumnName(Me.UltraLabel2, "")
        Me.Absx1.SetABSTableName(Me.UltraLabel2, "")
        Me.Absx1.SetABSViewName(Me.UltraLabel2, "")
        Me.UltraLabel2.Location = New System.Drawing.Point(16, 73)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel2.TabIndex = 11
        Me.UltraLabel2.Text = "State Name"
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "STATE_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.Absx1.SetABSParentColumnName(Me.UltraTextEditor1, "")
        Me.Absx1.SetABSTableName(Me.UltraTextEditor1, "")
        Me.Absx1.SetABSViewName(Me.UltraTextEditor1, "")
        Me.UltraTextEditor1.Location = New System.Drawing.Point(122, 39)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(100, 25)
        Me.UltraTextEditor1.TabIndex = 10
        '
        'TATSTATE
        '
        Me.Absx1.SetABSColumnName(Me, "")
        Me.Absx1.SetABSParentColumnName(Me, "")
        Me.Absx1.SetABSTableName(Me, "")
        Me.Absx1.SetABSViewName(Me, "")
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "TATSTATE"
        Me.Text = "TATSTATE"
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
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor2 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
End Class
