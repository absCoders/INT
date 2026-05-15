<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ICFIRTVF
    Inherits ASFBASE2

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
        Me.txtRANo = New System.Windows.Forms.TextBox
        Me.txtTrackingNo = New System.Windows.Forms.TextBox
        Me.lblRANo = New System.Windows.Forms.Label
        Me.lblTrackingNo = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnOK = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.btnCancel)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.btnOK)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.Label1)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.lblTrackingNo)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.lblRANo)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.txtTrackingNo)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.txtRANo)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(306, 153)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 153)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(306, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 153)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(306, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 153)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Margin = New System.Windows.Forms.Padding(4)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(306, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'txtRANo
        '
        Me.txtRANo.Location = New System.Drawing.Point(118, 46)
        Me.txtRANo.Name = "txtRANo"
        Me.txtRANo.Size = New System.Drawing.Size(157, 23)
        Me.txtRANo.TabIndex = 0
        '
        'txtTrackingNo
        '
        Me.txtTrackingNo.Location = New System.Drawing.Point(118, 78)
        Me.txtTrackingNo.Name = "txtTrackingNo"
        Me.txtTrackingNo.Size = New System.Drawing.Size(157, 23)
        Me.txtTrackingNo.TabIndex = 1
        '
        'lblRANo
        '
        Me.lblRANo.AutoSize = True
        Me.lblRANo.Location = New System.Drawing.Point(26, 53)
        Me.lblRANo.Name = "lblRANo"
        Me.lblRANo.Size = New System.Drawing.Size(85, 16)
        Me.lblRANo.TabIndex = 2
        Me.lblRANo.Text = "Vend RA No"
        '
        'lblTrackingNo
        '
        Me.lblTrackingNo.AutoSize = True
        Me.lblTrackingNo.Location = New System.Drawing.Point(26, 85)
        Me.lblTrackingNo.Name = "lblTrackingNo"
        Me.lblTrackingNo.Size = New System.Drawing.Size(86, 16)
        Me.lblTrackingNo.TabIndex = 3
        Me.lblTrackingNo.Text = "Tracking No"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(26, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(202, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Enter finalization information:"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(118, 118)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(66, 23)
        Me.btnOK.TabIndex = 5
        Me.btnOK.Text = "OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(206, 118)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(69, 23)
        Me.btnCancel.TabIndex = 6
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'ICFIRTVF
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(306, 153)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "ICFIRTVF"
        Me.Text = "Finalize RTV"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        Me.ASFBASE2_Fill_Panel.PerformLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblTrackingNo As System.Windows.Forms.Label
    Friend WithEvents lblRANo As System.Windows.Forms.Label
    Friend WithEvents txtTrackingNo As System.Windows.Forms.TextBox
    Friend WithEvents txtRANo As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
End Class
