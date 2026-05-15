<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFBASE2
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE0
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
        Me.ASFBASE2_Fill_Panel = New System.Windows.Forms.Panel()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 327)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(390, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 327)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(390, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 327)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(390, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Cursor = System.Windows.Forms.Cursors.Default
        Me.ASFBASE2_Fill_Panel.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ASFBASE2_Fill_Panel.Location = New System.Drawing.Point(0, 0)
        Me.ASFBASE2_Fill_Panel.Name = "ASFBASE2_Fill_Panel"
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(390, 327)
        Me.ASFBASE2_Fill_Panel.TabIndex = 0
        '
        'ASFBASE2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(390, 327)
        Me.Controls.Add(Me.ASFBASE2_Fill_Panel)
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "ASFBASE2"
        Me.Text = "ASFBASE2"
        Me.Controls.SetChildIndex(Me.ASFBASE2_Fill_Panel, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Bottom, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Top, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Right, 0)
        Me.Controls.SetChildIndex(Me._ASFBASE1_Toolbars_Dock_Area_Left, 0)
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Protected Friend WithEvents ASFBASE2_Fill_Panel As System.Windows.Forms.Panel
End Class
