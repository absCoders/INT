<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFDEVMO
    'Inherits System.Windows.Forms.Form
    Inherits ABSolution.ASFBASE2

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.cmdOK = New Infragistics.Win.Misc.UltraButton
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton
        Me.chkEnableDevMo = New System.Windows.Forms.CheckBox
        Me.grpDevMoOptions = New Infragistics.Win.Misc.UltraGroupBox
        Me.chkDataSourceToolTip = New System.Windows.Forms.CheckBox
        Me.chkBypassMenuLevelSecurity = New System.Windows.Forms.CheckBox
        Me.chkBypassMultiTask = New System.Windows.Forms.CheckBox
        Me.chkBypassSmtpSend = New System.Windows.Forms.CheckBox
        Me.chkBypassCopyReport = New System.Windows.Forms.CheckBox
        Me.cmdDeploy = New Infragistics.Win.Misc.UltraButton
        Me.chkRunDebug = New System.Windows.Forms.CheckBox
        Me.chkRunDebugPrompt = New System.Windows.Forms.CheckBox
        Me.ASFBASE2_Fill_Panel.SuspendLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.grpDevMoOptions, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDevMoOptions.SuspendLayout()
        Me.SuspendLayout()
        '
        'ASFBASE2_Fill_Panel
        '
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.cmdDeploy)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.grpDevMoOptions)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.chkEnableDevMo)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.cmdCancel)
        Me.ASFBASE2_Fill_Panel.Controls.Add(Me.cmdOK)
        Me.ASFBASE2_Fill_Panel.Size = New System.Drawing.Size(445, 290)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Left
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Left.Size = New System.Drawing.Size(0, 290)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Right
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Location = New System.Drawing.Point(445, 0)
        Me._ASFBASE1_Toolbars_Dock_Area_Right.Size = New System.Drawing.Size(0, 290)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Top
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Top.Size = New System.Drawing.Size(445, 0)
        '
        '_ASFBASE1_Toolbars_Dock_Area_Bottom
        '
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Location = New System.Drawing.Point(0, 290)
        Me._ASFBASE1_Toolbars_Dock_Area_Bottom.Size = New System.Drawing.Size(445, 0)
        '
        'tlb
        '
        Me.tlb.MenuSettings.ForceSerialization = True
        Me.tlb.ToolbarSettings.ForceSerialization = True
        '
        'cmdOK
        '
        Me.cmdOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance1.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(230, Byte), Integer))
        Appearance1.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(150, Byte), Integer))
        Appearance1.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Me.cmdOK.Appearance = Appearance1
        Me.cmdOK.Location = New System.Drawing.Point(261, 249)
        Me.cmdOK.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmdOK.Name = "cmdOK"
        Me.cmdOK.Size = New System.Drawing.Size(83, 26)
        Me.cmdOK.TabIndex = 2
        Me.cmdOK.Text = "OK"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance2.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(230, Byte), Integer))
        Appearance2.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(150, Byte), Integer))
        Appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Me.cmdCancel.Appearance = Appearance2
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(350, 249)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(83, 26)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        '
        'chkEnableDevMo
        '
        Me.chkEnableDevMo.AutoSize = True
        Me.chkEnableDevMo.Location = New System.Drawing.Point(12, 12)
        Me.chkEnableDevMo.Name = "chkEnableDevMo"
        Me.chkEnableDevMo.Size = New System.Drawing.Size(180, 20)
        Me.chkEnableDevMo.TabIndex = 25
        Me.chkEnableDevMo.Text = "Enable Developer Mode"
        Me.chkEnableDevMo.UseVisualStyleBackColor = True
        '
        'grpDevMoOptions
        '
        Me.grpDevMoOptions.BorderStyle = Infragistics.Win.Misc.GroupBoxBorderStyle.Rectangular3D
        Me.grpDevMoOptions.Controls.Add(Me.chkRunDebugPrompt)
        Me.grpDevMoOptions.Controls.Add(Me.chkRunDebug)
        Me.grpDevMoOptions.Controls.Add(Me.chkDataSourceToolTip)
        Me.grpDevMoOptions.Controls.Add(Me.chkBypassMenuLevelSecurity)
        Me.grpDevMoOptions.Controls.Add(Me.chkBypassMultiTask)
        Me.grpDevMoOptions.Controls.Add(Me.chkBypassSmtpSend)
        Me.grpDevMoOptions.Controls.Add(Me.chkBypassCopyReport)
        Me.grpDevMoOptions.Location = New System.Drawing.Point(26, 38)
        Me.grpDevMoOptions.Name = "grpDevMoOptions"
        Me.grpDevMoOptions.Size = New System.Drawing.Size(407, 186)
        Me.grpDevMoOptions.TabIndex = 26
        '
        'chkDataSourceToolTip
        '
        Me.chkDataSourceToolTip.AutoSize = True
        Me.chkDataSourceToolTip.Checked = True
        Me.chkDataSourceToolTip.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkDataSourceToolTip.Location = New System.Drawing.Point(6, 13)
        Me.chkDataSourceToolTip.Name = "chkDataSourceToolTip"
        Me.chkDataSourceToolTip.Size = New System.Drawing.Size(310, 20)
        Me.chkDataSourceToolTip.TabIndex = 30
        Me.chkDataSourceToolTip.Text = "Data Source Tool Tip (Ctrl + mouse hover)"
        Me.chkDataSourceToolTip.UseVisualStyleBackColor = True
        '
        'chkBypassMenuLevelSecurity
        '
        Me.chkBypassMenuLevelSecurity.AutoSize = True
        Me.chkBypassMenuLevelSecurity.Location = New System.Drawing.Point(6, 39)
        Me.chkBypassMenuLevelSecurity.Name = "chkBypassMenuLevelSecurity"
        Me.chkBypassMenuLevelSecurity.Size = New System.Drawing.Size(212, 20)
        Me.chkBypassMenuLevelSecurity.TabIndex = 29
        Me.chkBypassMenuLevelSecurity.Text = "Bypass Menu Level Security"
        Me.chkBypassMenuLevelSecurity.UseVisualStyleBackColor = True
        '
        'chkBypassMultiTask
        '
        Me.chkBypassMultiTask.AutoSize = True
        Me.chkBypassMultiTask.Location = New System.Drawing.Point(6, 143)
        Me.chkBypassMultiTask.Name = "chkBypassMultiTask"
        Me.chkBypassMultiTask.Size = New System.Drawing.Size(392, 20)
        Me.chkBypassMultiTask.TabIndex = 28
        Me.chkBypassMultiTask.Text = "Bypass Multi-Task Conflict Control (Blue Chips Needed)"
        Me.chkBypassMultiTask.UseVisualStyleBackColor = True
        '
        'chkBypassSmtpSend
        '
        Me.chkBypassSmtpSend.AutoSize = True
        Me.chkBypassSmtpSend.Location = New System.Drawing.Point(6, 91)
        Me.chkBypassSmtpSend.Name = "chkBypassSmtpSend"
        Me.chkBypassSmtpSend.Size = New System.Drawing.Size(391, 20)
        Me.chkBypassSmtpSend.TabIndex = 27
        Me.chkBypassSmtpSend.Text = "Bypass smtp.Send(mail) in TAFSEND1. return: Success"
        Me.chkBypassSmtpSend.UseVisualStyleBackColor = True
        '
        'chkBypassCopyReport
        '
        Me.chkBypassCopyReport.AutoSize = True
        Me.chkBypassCopyReport.Location = New System.Drawing.Point(6, 65)
        Me.chkBypassCopyReport.Name = "chkBypassCopyReport"
        Me.chkBypassCopyReport.Size = New System.Drawing.Size(367, 20)
        Me.chkBypassCopyReport.TabIndex = 26
        Me.chkBypassCopyReport.Text = "Bypass Copy Report to Archive prompt. default: No"
        Me.chkBypassCopyReport.UseVisualStyleBackColor = True
        '
        'cmdDeploy
        '
        Me.cmdDeploy.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Appearance3.BackColor = System.Drawing.Color.FromArgb(CType(CType(245, Byte), Integer), CType(CType(245, Byte), Integer), CType(CType(230, Byte), Integer))
        Appearance3.BackColor2 = System.Drawing.Color.FromArgb(CType(CType(165, Byte), Integer), CType(CType(165, Byte), Integer), CType(CType(150, Byte), Integer))
        Appearance3.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Me.cmdDeploy.Appearance = Appearance3
        Me.cmdDeploy.Location = New System.Drawing.Point(26, 249)
        Me.cmdDeploy.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cmdDeploy.Name = "cmdDeploy"
        Me.cmdDeploy.Size = New System.Drawing.Size(229, 26)
        Me.cmdDeploy.TabIndex = 27
        Me.cmdDeploy.Text = "Run Deployment Utility"
        Me.cmdDeploy.Visible = False
        '
        'chkRunDebug
        '
        Me.chkRunDebug.AutoSize = True
        Me.chkRunDebug.Location = New System.Drawing.Point(6, 117)
        Me.chkRunDebug.Name = "chkRunDebug"
        Me.chkRunDebug.Size = New System.Drawing.Size(135, 20)
        Me.chkRunDebug.TabIndex = 31
        Me.chkRunDebug.Text = "Run Debug Code"
        Me.chkRunDebug.UseVisualStyleBackColor = True
        '
        'chkRunDebugPrompt
        '
        Me.chkRunDebugPrompt.AutoSize = True
        Me.chkRunDebugPrompt.Checked = True
        Me.chkRunDebugPrompt.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkRunDebugPrompt.Location = New System.Drawing.Point(147, 117)
        Me.chkRunDebugPrompt.Name = "chkRunDebugPrompt"
        Me.chkRunDebugPrompt.Size = New System.Drawing.Size(73, 20)
        Me.chkRunDebugPrompt.TabIndex = 32
        Me.chkRunDebugPrompt.Text = "Prompt"
        Me.chkRunDebugPrompt.UseVisualStyleBackColor = True
        '
        'ASFDEVMO
        '
        Me.AcceptButton = Me.cmdOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(445, 290)
        Me.ControlBox = False
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(349, 138)
        Me.Name = "ASFDEVMO"
        Me.Text = "Absolution Developer Mode Options"
        Me.ASFBASE2_Fill_Panel.ResumeLayout(False)
        Me.ASFBASE2_Fill_Panel.PerformLayout()
        CType(Me.tlb, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASTOPST1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.tblASFBASE1_Schema, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dst, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.grpDevMoOptions, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDevMoOptions.ResumeLayout(False)
        Me.grpDevMoOptions.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Friend WithEvents cmdOK As Infragistics.Win.Misc.UltraButton
    Friend WithEvents chkEnableDevMo As System.Windows.Forms.CheckBox
    Friend WithEvents grpDevMoOptions As Infragistics.Win.Misc.UltraGroupBox
    Friend WithEvents chkDataSourceToolTip As System.Windows.Forms.CheckBox
    Friend WithEvents chkBypassMenuLevelSecurity As System.Windows.Forms.CheckBox
    Friend WithEvents chkBypassMultiTask As System.Windows.Forms.CheckBox
    Friend WithEvents chkBypassSmtpSend As System.Windows.Forms.CheckBox
    Friend WithEvents chkBypassCopyReport As System.Windows.Forms.CheckBox
    Friend WithEvents cmdDeploy As Infragistics.Win.Misc.UltraButton
    Friend WithEvents chkRunDebugPrompt As System.Windows.Forms.CheckBox
    Friend WithEvents chkRunDebug As System.Windows.Forms.CheckBox
End Class
