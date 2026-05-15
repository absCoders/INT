Imports Microsoft.VisualBasic.ApplicationServices

Namespace My

    ' The following events are availble for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        'Public Shared didOnce As Boolean

        'Shared Sub New()
        '    If Not didOnce Then
        '        System.Windows.Forms.Application.SetUnhandledExceptionMode(UnhandledExceptionMode.ThrowException)
        '        didOnce = True
        '        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        '        AddHandler currentDomain.UnhandledException, AddressOf MyHandler
        '    End If
        'End Sub

        'Public Shared Sub MyHandler(sender As Object, e As System.UnhandledExceptionEventArgs)
        '    Dim ex As Exception = DirectCast(e.ExceptionObject, Exception)
        '    MessageBox.Show("An error has occurred. ABSolution will shut down." & _
        '                    vbCrLf & vbCrLf & ex.Message, "Error")
        'End Sub

        Private Sub MyApplication_UnhandledException(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs) Handles Me.UnhandledException
            MsgBox(e.Exception.Message)
            Dim DT As String = Format(Now, "yyyyMMddHHmmss")
            Dim PFX As String = ASCMAIN1.USER_ID & "_" & DT
            Dim FILENAME As String = PFX & ".ERR"
            My.Computer.FileSystem.WriteAllText(FILENAME, e.Exception.Message & vbCrLf, False)
            My.Computer.FileSystem.WriteAllText(FILENAME, e.Exception.StackTrace & vbCrLf, True)

            My.Computer.FileSystem.CopyFile(FILENAME, ASCMAIN1.Folders("Archive") & "ERRs\" & FILENAME)

            FILENAME = PFX & ".xml"
            ASCMAIN1.ActiveForm.dst.WriteXml(FILENAME)

            My.Computer.FileSystem.CopyFile(FILENAME, ASCMAIN1.Folders("Archive") & "ERRs\" & FILENAME)

            FILENAME = PFX & ".txt"
            My.Computer.FileSystem.WriteAllText(FILENAME, ASCMAIN1.ActiveForm.Name & vbCrLf, False)
            If ASCMAIN1.ActiveForm.HFs.Count <> 0 Then
                For Each COLUMN_NAME As String In ASCMAIN1.ActiveForm.HFs.Keys
                    My.Computer.FileSystem.WriteAllText(FILENAME, COLUMN_NAME & " = " & ASCMAIN1.ActiveForm.HFs(COLUMN_NAME) & vbCrLf, True)
                Next
            End If

            My.Computer.FileSystem.CopyFile(FILENAME, ASCMAIN1.Folders("Archive") & "ERRs\" & FILENAME)


            Dim img As Image = ASFMAIN1.CaptureForm1()
            FILENAME = PFX & ".bmp"
            img.Save(FILENAME)

            My.Computer.FileSystem.CopyFile(FILENAME, ASCMAIN1.Folders("Archive") & "ERRs\" & FILENAME)


            ' LOG TO ORACLE IF CONNECTED
            If ASCMAIN1.oraCon IsNot Nothing AndAlso ASCMAIN1.oraCon.State = ConnectionState.Open Then
                Dim selectionNo As String = ""
                Dim re_xno As String = ""
                Dim usefulTrace As String = e.Exception.StackTrace

                If usefulTrace.Contains("ABSolution") Then 'only keep relevant parts of stack trace
                    usefulTrace = usefulTrace.Substring(0, usefulTrace.IndexOf(vbCr, usefulTrace.LastIndexOf("ABSolution")))
                End If

                usefulTrace = usefulTrace.Replace("'", "")
                If usefulTrace.Length > 2000 Then
                    usefulTrace = usefulTrace.Substring(0, 2000)
                End If

                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    selectionNo = ASCMAIN1.ActiveForm.SELECTION_NO.ToString
                    re_xno = ASCMAIN1.ActiveForm.RE_XNO.ToString
                End If

                Dim errMsg As String = e.Exception.Message
                errMsg = errMsg.Replace("'", "")
                If errMsg.Length > 2000 Then
                    errMsg = errMsg.Substring(0, 2000).Trim
                End If

                Dim sql As String = "INSERT INTO ASTERROR VALUES(" _
                & ":PARM1,:PARM2,:PARM3,SYSDATE,:PARM4,:PARM5,:PARM6)"
                ASCDATA1.ExecuteSQL(sql, "VNNVVV", New Object() _
                {ASCMAIN1.SESSION_NO, selectionNo, re_xno _
                , ASCMAIN1.USER_ID, errMsg, usefulTrace})
                'only execute if oracle is available
            End If

            MessageBox.Show("An error has occurred. ABSolution will shut down." & _
                            vbCrLf & vbCrLf & e.Exception.Message, "Error")

            'no need to do this if we are staying up
            '' double rems are john robbins doing
            For Each closeForm As System.Windows.Forms.Form In ASCMAIN1.ABS_FORMS
                If closeForm IsNot Nothing Then
                If closeForm.Name <> "ASFMAIN1" Then 'not for the main form
                    If TypeOf (closeForm) Is ASFBASE0 Then
                        If CType(closeForm, ASFBASE0).ScreenMode = True Then
                            CType(closeForm, ASFBASE0).Mode_Settings(False)
                        End If
                    End If
                    closeForm.Close()
                End If
                End If
            Next

            'need to do the code in the Try / Catch, and then set this flag
            'e.ExitApplication = False

        End Sub
    End Class

End Namespace

