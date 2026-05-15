Public Class ASCDEVMO


#Region "Class Variables"

    Private _DataSourceToolTip As Boolean = False
    Private _BypassCopyReport As Boolean = False
    Private _BypassSmtpSend As Boolean = False
    Private _BypassMenuLevelSecurity As Boolean = False
    Private _BypassMultiTask As Boolean = False
    Private _RunDebugCode As Boolean = False
    Private _RunDebugCodePrompt As Boolean = False

#End Region

#Region "Properties"
    ''' <summary>
    ''' Display datasource tooltips while holding Ctrl key and hovering over a form control 
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DataSourceToolTip() As Boolean
        Get
            Return _DataSourceToolTip
        End Get

        Set(ByVal value As Boolean)
            _DataSourceToolTip = value
        End Set
    End Property

    ''' <summary>
    ''' Bypass Copy File to Archive prompt with a default of No
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BypassCopyReport() As Boolean
        Get
            Return _BypassCopyReport
        End Get

        Set(ByVal value As Boolean)
            _BypassCopyReport = value
        End Set
    End Property

    ''' <summary>
    ''' Bypass smtp.Send(mail) in TAFSEND1 and return success
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BypassSmtpSend() As Boolean
        Get
            Return _BypassSmtpSend
        End Get

        Set(ByVal value As Boolean)
            _BypassSmtpSend = value
        End Set
    End Property

    ''' <summary>
    ''' Bypass Menu Level Security
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BypassMenuLevelSecurity() As Boolean
        Get
            Return _BypassMenuLevelSecurity
        End Get

        Set(ByVal value As Boolean)
            _BypassMenuLevelSecurity = value
        End Set
    End Property

    ''' <summary>
    ''' Bypass Multi-Task Conflict Control (Blue Chips Required)
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property BypassMultiTask() As Boolean
        Get
            Return _BypassMultiTask
        End Get

        Set(ByVal value As Boolean)
            _BypassMultiTask = value
        End Set
    End Property

    ''' <summary>
    ''' Encapsulate code in this developer mode boolean to enable conditional execution of compiled code.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RunDebugCode() As Boolean
        Get
            Return _RunDebugCode
        End Get

        Set(ByVal value As Boolean)
            _RunDebugCode = value
        End Set
    End Property

    ''' <summary>
    ''' Be prompted when you're about to execute debug code.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property RunDebugCodePrompt() As Boolean
        Get
            Return _RunDebugCodePrompt
        End Get

        Set(ByVal value As Boolean)
            _RunDebugCodePrompt = value
        End Set
    End Property

#End Region

End Class
