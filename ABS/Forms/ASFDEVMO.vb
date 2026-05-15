Public Class ASFDEVMO

    Private f_Calling_Form As ASFBASE0

    Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
        ASCMAIN1.developerMode = chkEnableDevMo.Checked
        ASCMAIN1.developerModeOptions.DataSourceToolTip = chkDataSourceToolTip.Checked
        ASCMAIN1.developerModeOptions.BypassCopyReport = chkBypassCopyReport.Checked
        ASCMAIN1.developerModeOptions.BypassSmtpSend = chkBypassSmtpSend.Checked
        ASCMAIN1.developerModeOptions.RunDebugCode = chkRunDebug.Checked
        ASCMAIN1.developerModeOptions.RunDebugCodePrompt = chkRunDebugPrompt.Checked
        ASCMAIN1.developerModeOptions.BypassMenuLevelSecurity = chkBypassMenuLevelSecurity.Checked
        ASCMAIN1.developerModeOptions.BypassMultiTask = chkBypassMultiTask.Checked
        Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub ASFDEVMO_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ASCMAIN1.Center(Me)
        chkEnableDevMo.Checked = ASCMAIN1.developerMode
        chkDataSourceToolTip.Checked = ASCMAIN1.developerModeOptions.DataSourceToolTip
        chkBypassCopyReport.Checked = ASCMAIN1.developerModeOptions.BypassCopyReport
        chkBypassSmtpSend.Checked = ASCMAIN1.developerModeOptions.BypassSmtpSend
        chkRunDebug.Checked = ASCMAIN1.developerModeOptions.RunDebugCode
        chkRunDebugPrompt.Checked = ASCMAIN1.developerModeOptions.RunDebugCodePrompt
        chkRunDebugPrompt.Visible = chkRunDebug.Checked
        chkBypassMenuLevelSecurity.Checked = ASCMAIN1.developerModeOptions.BypassMenuLevelSecurity
        chkBypassMultiTask.Checked = ASCMAIN1.developerModeOptions.BypassMultiTask
    End Sub

    Private Sub chkEnableDevMo_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEnableDevMo.CheckedChanged
        Set_Dev_MO()
    End Sub

    Sub Set_Dev_MO()
        Dim asbDeveloper As Boolean = False
        If chkEnableDevMo.Checked Then
            'turning DevMo on, Load defaults for user, enable checkboxes
            Select Case ASCMAIN1.USER_ID
                Case "rdw"
                    asbDeveloper = True
                    chkDataSourceToolTip.Checked = True
                    chkBypassCopyReport.Checked = True
                    chkBypassSmtpSend.Checked = True
                    chkRunDebug.Checked = True
                    chkRunDebugPrompt.Visible = True
                    chkRunDebugPrompt.Checked = True
                    chkBypassMenuLevelSecurity.Checked = True
                    chkBypassMultiTask.Checked = False
                Case "wjz", "gcv", "dgc", "whr"
                    asbDeveloper = True
                    chkDataSourceToolTip.Checked = True
                Case Else
                    chkDataSourceToolTip.Enabled = True
                    chkBypassCopyReport.Checked = False
                    chkBypassSmtpSend.Checked = False
                    chkRunDebug.Checked = False
                    chkRunDebugPrompt.Visible = False
                    chkBypassMenuLevelSecurity.Checked = False
                    chkBypassMultiTask.Checked = False
            End Select
            'The following options should only be available to those who understand the implications of using them
            chkBypassCopyReport.Enabled = asbDeveloper
            chkBypassSmtpSend.Enabled = asbDeveloper
            chkRunDebug.Enabled = asbDeveloper
            chkRunDebugPrompt.Enabled = asbDeveloper
            chkBypassMenuLevelSecurity.Enabled = asbDeveloper
            chkBypassMultiTask.Enabled = asbDeveloper
            cmdDeploy.Visible = ASCMAIN1.Running_in_VS
            cmdDeploy.Text = "Deployment Utility"
        Else
            'turning DevMo off, set all options to false
            chkDataSourceToolTip.Checked = chkEnableDevMo.Checked
            chkBypassCopyReport.Checked = chkEnableDevMo.Checked
            chkBypassSmtpSend.Checked = chkEnableDevMo.Checked
            chkRunDebug.Checked = chkEnableDevMo.Checked
            chkRunDebugPrompt.Checked = chkEnableDevMo.Checked
            chkBypassMenuLevelSecurity.Checked = chkEnableDevMo.Checked
            chkBypassMultiTask.Checked = chkEnableDevMo.Checked
            cmdDeploy.Visible = chkEnableDevMo.Checked
        End If

        grpDevMoOptions.Enabled = chkEnableDevMo.Checked

    End Sub

    Private Sub cmdDeploy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDeploy.Click
        Dim pathToBatchFile As String = "C:\VS\SEA\BuildDeploy.bat"
        Process.Start(pathToBatchFile)
    End Sub

    Private Sub chkRunDebug_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRunDebug.CheckedChanged

        chkRunDebugPrompt.Visible = chkRunDebug.Checked

    End Sub

    Private Sub chkRunDebugPrompt_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRunDebugPrompt.CheckedChanged
        Dim rdc As Boolean = chkRunDebug.Checked
        Dim rdcp As Boolean = chkRunDebugPrompt.Checked
        If (chkEnableDevMo.Checked And rdc) And Not rdcp Then
            Dim dcp As String = "Turning off this option runs the risk of executing test code unintentionally." _
            & vbCrLf & "Please be aware of what you are about to do." & vbCr & vbCr & "Continue Anyway?"
            If vbNo = MsgBox(dcp, vbQuestion + vbYesNo, "Are You Sure?") Then
                chkRunDebugPrompt.Checked = rdc
            End If
        End If
    End Sub
End Class