<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ASFLOGON
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ASFLOGON))
        Me.cmdLogOn = New Infragistics.Win.Misc.UltraButton()
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton()
        Me.lblUSER_ID = New Infragistics.Win.Misc.UltraLabel()
        Me.lblUSER_PASSWORD = New Infragistics.Win.Misc.UltraLabel()
        Me.lblDBS_SERVER = New Infragistics.Win.Misc.UltraLabel()
        Me.lblDBS_COMPANY = New Infragistics.Win.Misc.UltraLabel()
        Me.txtUSER_ID = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtUSER_PASSWORD = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtDBS_SERVER = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.txtDBS_COMPANY = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblDBS_PASSWORD = New Infragistics.Win.Misc.UltraLabel()
        Me.txtDBS_PASSWORD = New Infragistics.Win.UltraWinEditors.UltraTextEditor()
        Me.lblStatus = New Infragistics.Win.Misc.UltraLabel()
        Me.UltraPictureBox1 = New Infragistics.Win.UltraWinEditors.UltraPictureBox()
        CType(Me.txtUSER_ID, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtUSER_PASSWORD, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDBS_SERVER, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDBS_COMPANY, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtDBS_PASSWORD, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdLogOn
        '
        Me.cmdLogOn.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdLogOn.Location = New System.Drawing.Point(16, 225)
        Me.cmdLogOn.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdLogOn.Name = "cmdLogOn"
        Me.cmdLogOn.Size = New System.Drawing.Size(151, 36)
        Me.cmdLogOn.TabIndex = 5
        Me.cmdLogOn.Text = "Log-On"
        '
        'cmdCancel
        '
        Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(396, 225)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(151, 36)
        Me.cmdCancel.TabIndex = 6
        Me.cmdCancel.Text = "Cancel"
        '
        'lblUSER_ID
        '
        Me.lblUSER_ID.AutoSize = True
        Me.lblUSER_ID.Location = New System.Drawing.Point(16, 138)
        Me.lblUSER_ID.Margin = New System.Windows.Forms.Padding(4)
        Me.lblUSER_ID.Name = "lblUSER_ID"
        Me.lblUSER_ID.Size = New System.Drawing.Size(56, 18)
        Me.lblUSER_ID.TabIndex = 2
        Me.lblUSER_ID.Text = "User ID"
        '
        'lblUSER_PASSWORD
        '
        Me.lblUSER_PASSWORD.AutoSize = True
        Me.lblUSER_PASSWORD.Location = New System.Drawing.Point(16, 168)
        Me.lblUSER_PASSWORD.Margin = New System.Windows.Forms.Padding(4)
        Me.lblUSER_PASSWORD.Name = "lblUSER_PASSWORD"
        Me.lblUSER_PASSWORD.Size = New System.Drawing.Size(68, 18)
        Me.lblUSER_PASSWORD.TabIndex = 3
        Me.lblUSER_PASSWORD.Text = "Password"
        '
        'lblDBS_SERVER
        '
        Me.lblDBS_SERVER.AutoSize = True
        Me.lblDBS_SERVER.Location = New System.Drawing.Point(349, 166)
        Me.lblDBS_SERVER.Margin = New System.Windows.Forms.Padding(4)
        Me.lblDBS_SERVER.Name = "lblDBS_SERVER"
        Me.lblDBS_SERVER.Size = New System.Drawing.Size(49, 18)
        Me.lblDBS_SERVER.TabIndex = 4
        Me.lblDBS_SERVER.Text = "Server"
        '
        'lblDBS_COMPANY
        '
        Me.lblDBS_COMPANY.AutoSize = True
        Me.lblDBS_COMPANY.Location = New System.Drawing.Point(349, 137)
        Me.lblDBS_COMPANY.Margin = New System.Windows.Forms.Padding(4)
        Me.lblDBS_COMPANY.Name = "lblDBS_COMPANY"
        Me.lblDBS_COMPANY.Size = New System.Drawing.Size(68, 18)
        Me.lblDBS_COMPANY.TabIndex = 5
        Me.lblDBS_COMPANY.Text = "Company"
        '
        'txtUSER_ID
        '
        Me.txtUSER_ID.Location = New System.Drawing.Point(92, 135)
        Me.txtUSER_ID.Margin = New System.Windows.Forms.Padding(4)
        Me.txtUSER_ID.Name = "txtUSER_ID"
        Me.txtUSER_ID.Size = New System.Drawing.Size(122, 25)
        Me.txtUSER_ID.TabIndex = 0
        '
        'txtUSER_PASSWORD
        '
        Me.txtUSER_PASSWORD.Location = New System.Drawing.Point(92, 164)
        Me.txtUSER_PASSWORD.Margin = New System.Windows.Forms.Padding(4)
        Me.txtUSER_PASSWORD.Name = "txtUSER_PASSWORD"
        Me.txtUSER_PASSWORD.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtUSER_PASSWORD.Size = New System.Drawing.Size(122, 25)
        Me.txtUSER_PASSWORD.TabIndex = 1
        '
        'txtDBS_SERVER
        '
        Me.txtDBS_SERVER.Location = New System.Drawing.Point(425, 162)
        Me.txtDBS_SERVER.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDBS_SERVER.Name = "txtDBS_SERVER"
        Me.txtDBS_SERVER.Size = New System.Drawing.Size(122, 25)
        Me.txtDBS_SERVER.TabIndex = 3
        '
        'txtDBS_COMPANY
        '
        Me.txtDBS_COMPANY.Location = New System.Drawing.Point(425, 133)
        Me.txtDBS_COMPANY.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDBS_COMPANY.Name = "txtDBS_COMPANY"
        Me.txtDBS_COMPANY.Size = New System.Drawing.Size(122, 25)
        Me.txtDBS_COMPANY.TabIndex = 2
        '
        'lblDBS_PASSWORD
        '
        Me.lblDBS_PASSWORD.AutoSize = True
        Me.lblDBS_PASSWORD.Location = New System.Drawing.Point(349, 197)
        Me.lblDBS_PASSWORD.Margin = New System.Windows.Forms.Padding(4)
        Me.lblDBS_PASSWORD.Name = "lblDBS_PASSWORD"
        Me.lblDBS_PASSWORD.Size = New System.Drawing.Size(68, 18)
        Me.lblDBS_PASSWORD.TabIndex = 11
        Me.lblDBS_PASSWORD.Text = "Password"
        Me.lblDBS_PASSWORD.Visible = False
        '
        'txtDBS_PASSWORD
        '
        Me.txtDBS_PASSWORD.Location = New System.Drawing.Point(425, 195)
        Me.txtDBS_PASSWORD.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDBS_PASSWORD.Name = "txtDBS_PASSWORD"
        Me.txtDBS_PASSWORD.Size = New System.Drawing.Size(122, 25)
        Me.txtDBS_PASSWORD.TabIndex = 4
        Me.txtDBS_PASSWORD.Visible = False
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(16, 198)
        Me.lblStatus.Margin = New System.Windows.Forms.Padding(4)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(185, 18)
        Me.lblStatus.TabIndex = 13
        Me.lblStatus.Text = "Now Attempting to Log-On"
        Me.lblStatus.Visible = False
        '
        'UltraPictureBox1
        '
        Me.UltraPictureBox1.BorderShadowColor = System.Drawing.Color.Empty
        Me.UltraPictureBox1.Image = CType(resources.GetObject("UltraPictureBox1.Image"), Object)
        Me.UltraPictureBox1.Location = New System.Drawing.Point(13, 13)
        Me.UltraPictureBox1.Margin = New System.Windows.Forms.Padding(4)
        Me.UltraPictureBox1.Name = "UltraPictureBox1"
        Me.UltraPictureBox1.Size = New System.Drawing.Size(534, 110)
        Me.UltraPictureBox1.TabIndex = 10
        '
        'ASFLOGON
        '
        Me.AcceptButton = Me.cmdLogOn
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(563, 270)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblStatus)
        Me.Controls.Add(Me.txtDBS_PASSWORD)
        Me.Controls.Add(Me.lblDBS_PASSWORD)
        Me.Controls.Add(Me.UltraPictureBox1)
        Me.Controls.Add(Me.txtDBS_COMPANY)
        Me.Controls.Add(Me.txtDBS_SERVER)
        Me.Controls.Add(Me.txtUSER_PASSWORD)
        Me.Controls.Add(Me.txtUSER_ID)
        Me.Controls.Add(Me.lblDBS_COMPANY)
        Me.Controls.Add(Me.lblDBS_SERVER)
        Me.Controls.Add(Me.lblUSER_PASSWORD)
        Me.Controls.Add(Me.lblUSER_ID)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmdLogOn)
        Me.Font = New System.Drawing.Font("Verdana", 9.75!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.Name = "ASFLOGON"
        Me.Text = "Log-On"
        CType(Me.txtUSER_ID, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtUSER_PASSWORD, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDBS_SERVER, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDBS_COMPANY, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtDBS_PASSWORD, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdLogOn As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents lblUSER_ID As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents lblUSER_PASSWORD As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents lblDBS_SERVER As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents lblDBS_COMPANY As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtUSER_ID As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtUSER_PASSWORD As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtDBS_SERVER As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents txtDBS_COMPANY As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraPictureBox1 As Infragistics.Win.UltraWinEditors.UltraPictureBox
    Friend WithEvents lblDBS_PASSWORD As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtDBS_PASSWORD As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblStatus As Infragistics.Win.Misc.UltraLabel
End Class
