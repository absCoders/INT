<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFPWDC1
    Inherits System.Windows.Forms.Form

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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ASFPWDC1))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.UltraPictureBox1 = New Infragistics.Win.UltraWinEditors.UltraPictureBox()
        Me.txtUSER_PASSWORD = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtUSER_ID = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblUSER_PASSWORD = New Infragistics.Win.Misc.UltraLabel()
        Me.lblUSER_ID = New Infragistics.Win.Misc.UltraLabel()
        Me.txtNewPass = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel()
        Me.txtNewPassVer = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.UltraLabel2 = New Infragistics.Win.Misc.UltraLabel()
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton()
        Me.cmdUpdate = New Infragistics.Win.Misc.UltraButton()
        CType(Me.txtUSER_PASSWORD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtUSER_ID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtNewPass, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtNewPassVer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'UltraPictureBox1
        '
        Me.UltraPictureBox1.BorderShadowColor = System.Drawing.Color.Empty
        Me.UltraPictureBox1.Image = CType(resources.GetObject("UltraPictureBox1.Image"), Object)
        Me.UltraPictureBox1.Location = New System.Drawing.Point(13, 13)
        Me.UltraPictureBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraPictureBox1.Name = "UltraPictureBox1"
        Me.UltraPictureBox1.Size = New System.Drawing.Size(534, 110)
        Me.UltraPictureBox1.TabIndex = 11
        '
        'txtUSER_PASSWORD
        '
        Me.txtUSER_PASSWORD.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.txtUSER_PASSWORD.Location = New System.Drawing.Point(149, 158)
        Me.txtUSER_PASSWORD.Margin = New System.Windows.Forms.Padding(4)
        Me.txtUSER_PASSWORD.Name = "txtUSER_PASSWORD"
        Me.txtUSER_PASSWORD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtUSER_PASSWORD.Size = New System.Drawing.Size(121, 25)
        Me.txtUSER_PASSWORD.TabIndex = 1
        '
        'txtUSER_ID
        '
        Me.txtUSER_ID.Enabled = False
        Me.txtUSER_ID.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.txtUSER_ID.Location = New System.Drawing.Point(149, 128)
        Me.txtUSER_ID.Margin = New System.Windows.Forms.Padding(4)
        Me.txtUSER_ID.Name = "txtUSER_ID"
        Me.txtUSER_ID.Size = New System.Drawing.Size(121, 25)
        Me.txtUSER_ID.TabIndex = 0
        '
        'lblUSER_PASSWORD
        '
        Me.lblUSER_PASSWORD.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.lblUSER_PASSWORD.Location = New System.Drawing.Point(24, 161)
        Me.lblUSER_PASSWORD.Margin = New System.Windows.Forms.Padding(4)
        Me.lblUSER_PASSWORD.Name = "lblUSER_PASSWORD"
        Me.lblUSER_PASSWORD.Size = New System.Drawing.Size(155, 22)
        Me.lblUSER_PASSWORD.TabIndex = 15
        Me.lblUSER_PASSWORD.Text = "Current Password"
        '
        'lblUSER_ID
        '
        Me.lblUSER_ID.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.lblUSER_ID.Location = New System.Drawing.Point(24, 131)
        Me.lblUSER_ID.Margin = New System.Windows.Forms.Padding(4)
        Me.lblUSER_ID.Name = "lblUSER_ID"
        Me.lblUSER_ID.Size = New System.Drawing.Size(119, 22)
        Me.lblUSER_ID.TabIndex = 14
        Me.lblUSER_ID.Text = "User ID"
        '
        'txtNewPass
        '
        Me.txtNewPass.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.txtNewPass.Location = New System.Drawing.Point(425, 128)
        Me.txtNewPass.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNewPass.Name = "txtNewPass"
        Me.txtNewPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNewPass.Size = New System.Drawing.Size(122, 25)
        Me.txtNewPass.TabIndex = 2
        '
        'UltraLabel1
        '
        Appearance1.TextHAlignAsString = "Right"
        Me.UltraLabel1.Appearance = Appearance1
        Me.UltraLabel1.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.UltraLabel1.Location = New System.Drawing.Point(298, 131)
        Me.UltraLabel1.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(119, 22)
        Me.UltraLabel1.TabIndex = 17
        Me.UltraLabel1.Text = "New Password"
        '
        'txtNewPassVer
        '
        Me.txtNewPassVer.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.txtNewPassVer.Location = New System.Drawing.Point(425, 158)
        Me.txtNewPassVer.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNewPassVer.Name = "txtNewPassVer"
        Me.txtNewPassVer.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNewPassVer.Size = New System.Drawing.Size(122, 25)
        Me.txtNewPassVer.TabIndex = 3
        '
        'UltraLabel2
        '
        Appearance2.TextHAlignAsString = "Right"
        Me.UltraLabel2.Appearance = Appearance2
        Me.UltraLabel2.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.UltraLabel2.Location = New System.Drawing.Point(263, 161)
        Me.UltraLabel2.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraLabel2.Name = "UltraLabel2"
        Me.UltraLabel2.Size = New System.Drawing.Size(154, 22)
        Me.UltraLabel2.TabIndex = 19
        Me.UltraLabel2.Text = "Verify New Password"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.cmdCancel.Location = New System.Drawing.Point(396, 221)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(151, 36)
        Me.cmdCancel.TabIndex = 5
        Me.cmdCancel.Text = "Cancel"
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.cmdUpdate.Location = New System.Drawing.Point(13, 221)
        Me.cmdUpdate.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(151, 36)
        Me.cmdUpdate.TabIndex = 4
        Me.cmdUpdate.Text = "Change Password"
        '
        'ASFPWDC1
        '
        Me.AcceptButton = Me.cmdUpdate
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(563, 270)
        Me.Controls.Add(Me.txtUSER_PASSWORD)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.txtNewPassVer)
        Me.Controls.Add(Me.UltraLabel2)
        Me.Controls.Add(Me.txtNewPass)
        Me.Controls.Add(Me.UltraLabel1)
        Me.Controls.Add(Me.txtUSER_ID)
        Me.Controls.Add(Me.lblUSER_PASSWORD)
        Me.Controls.Add(Me.lblUSER_ID)
        Me.Controls.Add(Me.UltraPictureBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "ASFPWDC1"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Change Password"
        CType(Me.txtUSER_PASSWORD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtUSER_ID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtNewPass, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtNewPassVer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents UltraPictureBox1 As Infragistics.Win.UltraWinEditors.UltraPictureBox
    Friend WithEvents txtUSER_PASSWORD As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtUSER_ID As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblUSER_PASSWORD As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents lblUSER_ID As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtNewPass As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtNewPassVer As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel2 As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdUpdate As Infragistics.Win.Misc.UltraButton

End Class
