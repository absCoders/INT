<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

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
        Me.components = New System.ComponentModel.Container()
        Me.Blowfish1 = New nsoftware.IPWorksEncrypt.Blowfish(Me.components)
        Me.Ipport1 = New nsoftware.IPWorks.TCPClient(Me.components)
        Me.Zip1 = New nsoftware.IPWorksZip.Zip(Me.components)
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Ftp1 = New nsoftware.IPWorks.Ftp(Me.components)
        Me.Sftp1 = New nsoftware.IPWorksSSH.SFTPClient(Me.components)
        Me.SuspendLayout()
        '
        'Blowfish1
        '
        Me.Blowfish1.About = "IP*Works! Encrypt V9 [Build 5329]"
        '
        'Ipport1
        '
        Me.Ipport1.About = "IP*Works! V9 [Build 5157]"
        '
        'Zip1
        '
        Me.Zip1.About = "IP*Works! ZIP V9 [Build 5157]"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(41, 53)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Button1"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Ftp1
        '
        Me.Ftp1.About = "IP*Works! V9 [Build 5157]"
        '
        'Sftp1
        '
        Me.Sftp1.About = "IP*Works! SSH V9 [Build 5157]"

        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 262)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Blowfish1 As nsoftware.IPWorksEncrypt.Blowfish
    Friend WithEvents Ipport1 As nsoftware.IPWorks.TCPClient
    Friend WithEvents Zip1 As nsoftware.IPWorksZip.Zip
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Ftp1 As nsoftware.IPWorks.FTP
    Friend WithEvents Sftp1 As nsoftware.IPWorksSSH.SFTPClient
End Class
