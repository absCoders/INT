<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASTDSQLG
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
        Me.UltraTextEditor1 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor2 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor3 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel3 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraTextEditor4 = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel4 = New Infragistics.Win.Misc.UltraLabel
        Me.UltraLabel5 = New Infragistics.Win.Misc.UltraLabel
        Me.Panel1.SuspendLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.UltraTextEditor4, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.UltraLabel5)
        Me.Panel1.Controls.Add(Me.UltraTextEditor4)
        Me.Panel1.Controls.Add(Me.UltraLabel4)
        Me.Panel1.Controls.Add(Me.UltraTextEditor3)
        Me.Panel1.Controls.Add(Me.UltraLabel3)
        Me.Panel1.Controls.Add(Me.UltraTextEditor2)
        Me.Panel1.Controls.Add(Me.UltraLabel2)
        Me.Panel1.Controls.Add(Me.UltraTextEditor1)
        Me.Panel1.Controls.Add(Me.UltraLabel1)
        Me.Panel1.Size = New System.Drawing.Size(776, 574)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor1, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor2, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel3, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor3, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel4, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraTextEditor4, 0)
        Me.Panel1.Controls.SetChildIndex(Me.UltraLabel5, 0)
        '
        'UltraExplorerBar1
        '
        Me.UltraExplorerBar1.GroupSettings.UseMnemonics = Infragistics.Win.DefaultableBoolean.[True]
        Me.UltraExplorerBar1.ItemSettings.ForceSerialization = True
        Me.UltraExplorerBar1.Margins.ForceSerialization = True
        '
        'UltraTextEditor1
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor1, "TABLE_NAME")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor1, True)
        Me.UltraTextEditor1.Location = New System.Drawing.Point(122, 41)
        Me.UltraTextEditor1.Name = "UltraTextEditor1"
        Me.UltraTextEditor1.Size = New System.Drawing.Size(196, 25)
        Me.UltraTextEditor1.TabIndex = 9
        '
        'UltraLabel1
        '
        Me.UltraLabel1.Location = New System.Drawing.Point(16, 44)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel1.TabIndex = 8
        Me.UltraLabel1.Text = "Table Name"
        '
        'UltraTextEditor2
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor2, "COLUMN_NAME_CODE")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor2, True)
        Me.UltraTextEditor2.Location = New System.Drawing.Point(122, 69)
        Me.UltraTextEditor2.Name = "UltraTextEditor2"
        Me.UltraTextEditor2.Size = New System.Drawing.Size(196, 25)
        Me.UltraTextEditor2.TabIndex = 11
        '
        'UltraLabel2
        '
        Me.UltraLabel2.Location = New System.Drawing.Point(16, 72)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel2.TabIndex = 10
        Me.UltraLabel2.Text = "Code Column"
        '
        'UltraTextEditor3
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor3, "COLUMN_NAME_DESC")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor3, True)
        Me.UltraTextEditor3.Location = New System.Drawing.Point(122, 98)
        Me.UltraTextEditor3.Name = "UltraTextEditor3"
        Me.UltraTextEditor3.Size = New System.Drawing.Size(196, 25)
        Me.UltraTextEditor3.TabIndex = 13
        '
        'UltraLabel3
        '
        Me.UltraLabel3.Location = New System.Drawing.Point(16, 101)
        Me.UltraLabel3.Name = "UltraLabel3"
        Me.UltraLabel3.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel3.TabIndex = 12
        Me.UltraLabel3.Text = "Desc Column"
        '
        'UltraTextEditor4
        '
        Me.Absx1.SetABSColumnName(Me.UltraTextEditor4, "COLUMN_NAME_KEY")
        Me.Absx1.SetABSHasButton(Me.UltraTextEditor4, True)
        Me.UltraTextEditor4.Location = New System.Drawing.Point(122, 127)
        Me.UltraTextEditor4.Name = "UltraTextEditor4"
        Me.UltraTextEditor4.Size = New System.Drawing.Size(196, 25)
        Me.UltraTextEditor4.TabIndex = 15
        '
        'UltraLabel4
        '
        Me.UltraLabel4.Location = New System.Drawing.Point(16, 130)
        Me.UltraLabel4.Name = "UltraLabel4"
        Me.UltraLabel4.Size = New System.Drawing.Size(100, 23)
        Me.UltraLabel4.TabIndex = 14
        Me.UltraLabel4.Text = "Group Key"
        '
        'UltraLabel5
        '
        Me.UltraLabel5.Location = New System.Drawing.Point(16, 159)
        Me.UltraLabel5.Name = "UltraLabel5"
        Me.UltraLabel5.Size = New System.Drawing.Size(368, 23)
        Me.UltraLabel5.TabIndex = 16
        Me.UltraLabel5.Text = "Use Group Key only if not the same as Column Name"
        '
        'ASTDSQLG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(989, 574)
        Me.Name = "ASTDSQLG"
        Me.Text = "ASTDSQLG"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.tbl, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraExplorerBar1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.UltraTextEditor4, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UltraTextEditor1 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraLabel5 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor4 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel4 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor3 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel3 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents UltraTextEditor2 As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
End Class
