Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Infragistics.Shared
Imports Infragistics.Win
Imports Infragistics.Win.UltraWinSchedule
Imports System.Diagnostics
Imports System.Resources
Imports System.Globalization
Imports Infragistics.Win.Misc
Imports Infragistics.Win.UltraWinEditors

Public Class ASFAPPTR
    Inherits System.Windows.Forms.Form


#Region "Constants"

    Shared ReadOnly DIALOG_WIDTH As Int32 = 468
    Private Shared ReadOnly DIALOG_HEIGHT As Int32 = 351
    Private Shared ReadOnly PANEL_WIDTH As Int32 = 350
    Private Shared ReadOnly PANEL_HEIGHT As Int32 = 96

    Private Shared ReadOnly PANEL_LEFT As Int32 = 80
    Private Shared ReadOnly PANEL_TOP As Int32 = 16

    Friend Shared ReadOnly defaultOccurrenceDuration As TimeSpan = TimeSpan.Zero
    Friend Shared ReadOnly defaultOccurrenceStartTime As DateTime = DateTime.MinValue
    Friend Shared ReadOnly defaultPatternDayOfMonth As Int32 = 0
    Friend Shared ReadOnly defaultPatternDaysOfWeek As RecurrencePatternDaysOfWeek = RecurrencePatternDaysOfWeek.None
    Friend Shared ReadOnly defaultPatternFrequency As RecurrencePatternFrequency = RecurrencePatternFrequency.Weekly
    Friend Shared ReadOnly defaultPatternInterval As Int32 = 1
    Friend Shared ReadOnly defaultPatternMonthOfYear As Int32 = 0
    Friend Shared ReadOnly defaultPatternOccurrenceOfDayInMonth As RecurrencePatternOccurrenceOfDayInMonth = RecurrencePatternOccurrenceOfDayInMonth.None
    Friend Shared ReadOnly defaultRangeEndDate As DateTime = DateTime.MinValue
    Friend Shared ReadOnly defaultRangeStartDate As DateTime = DateTime.MinValue
    Friend Shared ReadOnly defaultRangeLimit As RecurrenceRangeLimit = RecurrenceRangeLimit.NoLimit
    Friend Shared ReadOnly defaultRangeMaxOccurrences As Int32 = 10
    Friend Shared ReadOnly defaultPatternType As RecurrencePatternType = RecurrencePatternType.Explicit

    Friend Shared ReadOnly maxPatternDayOfMonth As Int32 = 31
    Friend Shared ReadOnly maxPatternInterval As Int32 = 99
    Friend Shared ReadOnly maxPatternMonthOfYear As Int32 = 12
    Friend Shared ReadOnly maxMaxOccurrences As Int32 = 999

    Private Shared ReadOnly RecurrenceDialog_Error_DurationCannotExceedFrequency As String = "The duration of the appointment must be shorter than how frequently it occurs. Shorten the duration, or change the recurrence pattern in the Appointment Recurrence dialog box."
    Private Shared ReadOnly RecurrenceDialog_Warning_PatternDayOfMonthExceeds28 As String = "Some months have fewer than the number of days specified. For these months, the occurrence will fall on the last day of the month."
    Private Shared ReadOnly RecurrenceDialog_Error_RecurrencePatterIsInvalid As String = "The recurrence pattern is not valid."
    Private Shared ReadOnly RecurrenceDialog_MessageBox_Caption As String = "Appointment Recurrence"
    Private Shared ReadOnly RecurrenceDialog_Prompt_OkToRemoveVariances As String = "Any exceptions associated with this recurring appointment will be lost. Is this OK?"

#End Region

#Region "Member variables"

    Private patternDayOfMonthWarningDisplayed As Boolean = False
    Private allowNoEndDate As Boolean = True
    Private appointment As appointment = Nothing
    Private appointmentRecurrence As appointmentRecurrence = Nothing
    Private initialAppointmentRecurrence As appointmentRecurrence = Nothing

    Private _Tickler As String = ""
    Private _Caption As String = ""

    Private suspendEventHandling As Boolean = False
    Private suspendUpdate As Boolean = False
    Private suspendOptionButtonHandling As Boolean = False
    Private isInitializing As Boolean = False
    Private _result As RecurrenceDialogResult = RecurrenceDialogResult.Cancel
    Private hasVariances As Boolean = False
    Private forceContentsChanged As Boolean = False
    Friend WithEvents cmdOk As Infragistics.Win.Misc.UltraButton
    Friend WithEvents pnlDaily As System.Windows.Forms.Panel
    Friend WithEvents lblPatternIntervalDaily As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtPatternIntervalDaily As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents optPatternDaysOfWeek_Daily As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents pnlMonthly As System.Windows.Forms.Panel
    Friend WithEvents lblPatternInterval4Monthly As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtPatternIntervalCalculatedMonthly As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblPatternInterval3Monthly As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboDayOfWeekMonthly As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents cboOccurrenceOfDayInMonthMonthly As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents txtPatternIntervalExplicitMonthly As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblPatternInterval2Monthly As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents lblPatternInterval1Monthly As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtPatternDayOfMonthMonthly As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents optPatternType_Monthly As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents cmdRemoveRecurrence As Infragistics.Win.Misc.UltraButton
    Friend WithEvents cmdCancel As Infragistics.Win.Misc.UltraButton
    Friend WithEvents pnlYearly As System.Windows.Forms.Panel
    Friend WithEvents cboMonthOfYearCalculatedYearly As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents lblCalculatedYearly As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboDayOfWeekYearly As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents cboOccurrenceOfDayInMonthYearly As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents txtPatternDayOfMonthYearly As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents cboMonthOfYearExplicitYearly As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents optPatternType_Yearly As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents pnlWeekly As System.Windows.Forms.Panel
    Friend WithEvents chkSaturday As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkFriday As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkThursday As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkWednesday As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkTuesday As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkMonday As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents chkSunday As Infragistics.Win.UltraWinEditors.UltraCheckEditor
    Friend WithEvents lblPatternInterval2Weekly As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtPatternIntervalWeekly As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents lblPatternInterval1_Weekly As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents fraAppointmentTime As System.Windows.Forms.GroupBox
    Friend WithEvents cboOccurrenceDuration As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents lblOccurrenceDuration As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboEndTime As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents lblEndTime As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents cboOccurrenceStartTime As Infragistics.Win.UltraWinEditors.UltraComboEditor
    Friend WithEvents lblOccurrenceStartTime As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents fraRecurrenceRange As System.Windows.Forms.GroupBox
    Friend WithEvents dtpRangeEndDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents lblRangeMaxOccurrences As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtRangeMaxOccurrences As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents dtpRangeStartDate As Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
    Friend WithEvents lblRangeStartDate As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents optRangeLimit As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents fraRecurrencePattern As System.Windows.Forms.GroupBox
    Friend WithEvents optPatternFrequency As Infragistics.Win.UltraWinEditors.UltraOptionSet
    Friend WithEvents lblSepDark As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents lblSepLight As Infragistics.Win.Misc.UltraLabel
    Friend WithEvents txtTickler As Infragistics.Win.UltraWinEditors.UltraTextEditor
    Friend WithEvents UltraLabel1 As Infragistics.Win.Misc.UltraLabel

    Private contentsAreValid As Boolean = True

#End Region

#Region "Constructor"

    Public Sub New( _
    ByVal appointment As Appointment, _
    ByVal recurrence As AppointmentRecurrence, _
    ByVal alloNoEndDate As Boolean, _
    ByVal allowRemoveRecurrence As Boolean, _
    ByVal hasVariances As Boolean, _
    ByVal Caption As String, _
    ByVal Tickler As String)

        MyBase.New()

        _Tickler = Tickler
        _Caption = Caption

        '	Throw an exception if either the appointment or the recurrence are null
        If (appointment Is Nothing Or recurrence Is Nothing) Then
            Throw New ArgumentNullException()
        End If

        '	Flag the dialog as being in a state of initializing
        Me.isInitializing = True

        '	Set the 'appointment' member variable
        Me.appointment = appointment

        '	If the specified appointment is currently not a member
        '	of a recurrence, we want to consider the dialog dirty
        '	even if no dialog fields were modified, so that clicking
        '	Ok results in the recurrence being assigned to the appointment.
        If (Not Me.appointment.IsRecurringAppointmentRoot AndAlso Me.appointment.RecurringAppointmentRoot Is Nothing) Then
            Me.forceContentsChanged = True
        End If

        '	Set the 'appointmentRecurrence' member variable
        Dim recurrenceCopy As AppointmentRecurrence = New AppointmentRecurrence()
        recurrenceCopy.InitializeFrom(recurrence, True)
        Me.appointmentRecurrence = recurrenceCopy

        '	Set the 'hasVariances' member variable
        Me.hasVariances = hasVariances

        '	Cache the recurrence, so we have something to compare
        '	the working one to in order to know if anything changed.
        Me.initialAppointmentRecurrence = New AppointmentRecurrence()
        Me.initialAppointmentRecurrence.InitializeFrom(Me.appointmentRecurrence, False)

        Me.InitializeComponent()

        '	If the 'allowRemoveRecurrence'parameter is false, disable the 'cmdRemoveRecurrence' button.
        Me.cmdRemoveRecurrence.Enabled = allowRemoveRecurrence

        '	Cache the value of the 'allowNoEndDate' parameter
        Me.allowNoEndDate = allowNoEndDate

        If (Me.allowNoEndDate = False AndAlso Me.appointmentRecurrence.RangeLimit = RecurrenceRangeLimit.NoLimit) Then
            Me.appointmentRecurrence.RangeLimit = RecurrenceRangeLimit.LimitByNumberOfOccurrences
        End If

        '	Hook the events of interest.
        Me.HookDialogControlEvents(False)

        '	Initialize the dialog controls to reflect the property values
        '	of the AppointmentRecurrence that it is going to be editing
        Me.InitializeDialogControls()

        '	Set some Form-specific properties
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.ControlBox = False
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.ControlBox = True
        Me.StartPosition = FormStartPosition.CenterParent

        '	Set the Accept and Cancel buttons
        Me.AcceptButton = Me.cmdOk
        Me.CancelButton = Me.cmdCancel

        '	Set the size of the dialog
        Me.Size = New Size(ASFAPPTR.DIALOG_WIDTH, ASFAPPTR.DIALOG_HEIGHT)

        '	Clear the 'isInitializing' flag.
        Me.isInitializing = False
    End Sub


#End Region

#Region "Public Properties"

#Region "Recurrence"
    ' <summary>
    ' (Read-only) Returns an <see cref="AppointmentRecurrence"/> instance
    ' that reflects the values entered by the end user during this dialog session.
    ' </summary>
    ' <remarks><p class="note"><b>Note:</b> This property is only valid when the dialog has been closed and the dialog session was not canceled.</p></remarks>
    Public ReadOnly Property Recurrence() As AppointmentRecurrence

        Get

            If (Me.Result = RecurrenceDialogResult.Cancel Or _
              Me.Result = RecurrenceDialogResult.RemoveRecurrence) Then
                Return Nothing
            End If

            Me.UpdateRecurrence()
            Return Me.appointmentRecurrence
        End Get
    End Property
#End Region

#Region "Result"
    ' <summary>
    ' Returns the result of the dialog's last user session.
    ' </summary>
    Public ReadOnly Property Result() As RecurrenceDialogResult

        Get
            Return Me._result
        End Get

    End Property
#End Region

#End Region

#Region "Internal Properties"

#Region "PanelSize"
#If DEBUG Then
    ' <summary>
    ' Returns the size to be used for the Panel controls.
    ' </summary>
#End If
    Friend ReadOnly Property PanelSize() As Size

        Get
            Return New Size(ASFAPPTR.PANEL_WIDTH, ASFAPPTR.PANEL_HEIGHT)
        End Get

    End Property
#End Region

#Region "PanelLocation"
#If DEBUG Then
    ' <summary>
    ' Returns the location to be used for the Panel controls.
    ' Location is relative to the fraRecurrencePattern GroupBox.
    ' </summary>
#End If
    Friend ReadOnly Property PanelLocation() As Point

        Get
            Return New Point(ASFAPPTR.PANEL_LEFT, ASFAPPTR.PANEL_TOP)
        End Get

    End Property
#End Region

#Region "Recurrence properties"
    '
    '	The following region contains properties that have the same name
    '	as their AppointmentRecurrence counterparts. The get methods return
    '	a value that is based on the state of the dialog control that
    '	represents that property the set methods set the state of the
    '	dialog control that corresponds to the property being set. This is
    '	an abstraction that is designed to make it easier to map AppointmentRecurrence
    '	properties to the UI, and get the values back again.
    '	
    '	Note that the set methods are not likely to throw an exception, but the gets
    '	quite possibly will (for example, if the contents of a TextBox is not a number,
    '	but we try to convert it as such). Recommended usage is to wrap the call to the
    '	property get in a try/catch, and display a meaningful error to the end user (usually
    '	"The recurrence pattern is not valid", as with MS Outlook.)

#Region "RangeStartDate"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the RangeStartDate (dtpRangeStartDate)
    ' </summary>
#End If
    Friend Property RangeStartDate() As DateTime

        Get
            Return Me.dtpRangeStartDate.Value
        End Get

        Set(ByVal Value As DateTime)

            Me.dtpRangeStartDate.Value = Value
        End Set
    End Property
#End Region

#Region "RangeEndDate"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the RangeEndDate (dtpRangeEndDate)
    ' </summary>
#End If
    Friend Property RangeEndDate() As DateTime

        Get
            Return Me.dtpRangeEndDate.Value
        End Get
        Set(ByVal Value As DateTime)

            Me.dtpRangeEndDate.Value = Value
        End Set

    End Property

#End Region

#Region "PatternFrequency"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the PatternFrequency (optDaily, optWeekly, optMonthly, optYearly)
    ' </summary>
#End If
    Friend Property PatternFrequency() As RecurrencePatternFrequency

        Get

            Return Me.optPatternFrequency.Value

        End Get

        Set(ByVal Value As RecurrencePatternFrequency)

            Me.optPatternFrequency.Value = Value

        End Set
    End Property
#End Region

#Region "RangeLimit"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the RangeLimit (optNoLimit, optLimitByNumberOfOccurrences, optNoLimit)
    ' </summary>
#End If
    Friend Property RangeLimit() As RecurrenceRangeLimit

        Get

            Return Me.optRangeLimit.Value

        End Get
        Set(ByVal Value As RecurrenceRangeLimit)

            Me.optRangeLimit.Value = Value

        End Set
    End Property
#End Region

#Region "PatternDaysOfWeek"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the PatternDaysOfWeek
    ' (For Weekly: chkSunday - chkSaturday for Daily: optAllDays, optAllWeekdays)
    ' </summary>
#End If
    Friend Property PatternDaysOfWeek() As RecurrencePatternDaysOfWeek

        Get

            If (Me.PatternFrequency = RecurrencePatternFrequency.Daily) Then

                Return Me.optPatternDaysOfWeek_Daily.Value

            ElseIf (Me.PatternFrequency = RecurrencePatternFrequency.Weekly) Then

                '	Build a bitmask based on which checkboxes are checked.
                Dim retVal As RecurrencePatternDaysOfWeek = RecurrencePatternDaysOfWeek.None
                If (Me.chkSunday.Checked) Then retVal = (retVal Or RecurrencePatternDaysOfWeek.Sunday)
                If (Me.chkMonday.Checked) Then retVal = (retVal Or RecurrencePatternDaysOfWeek.Monday)
                If (Me.chkTuesday.Checked) Then retVal = (retVal Or RecurrencePatternDaysOfWeek.Tuesday)
                If (Me.chkWednesday.Checked) Then retVal = (retVal Or RecurrencePatternDaysOfWeek.Wednesday)
                If (Me.chkThursday.Checked) Then retVal = (retVal Or RecurrencePatternDaysOfWeek.Thursday)
                If (Me.chkFriday.Checked) Then retVal = (retVal Or RecurrencePatternDaysOfWeek.Friday)
                If (Me.chkSaturday.Checked) Then retVal = (retVal Or RecurrencePatternDaysOfWeek.Saturday)

                Return retVal

            ElseIf (Me.PatternFrequency = RecurrencePatternFrequency.Monthly) Then

                Dim item As ValueListItem = Me.cboDayOfWeekMonthly.SelectedItem
                If Not item Is Nothing Then Return item.DataValue Else Return RecurrencePatternDaysOfWeek.None
            ElseIf (Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then

                Dim item As ValueListItem = Me.cboDayOfWeekYearly.SelectedItem
                If Not item Is Nothing Then Return item.DataValue Else Return RecurrencePatternDaysOfWeek.None
            End If

            Return RecurrencePatternDaysOfWeek.None

        End Get

        Set(ByVal Value As RecurrencePatternDaysOfWeek)

            If (Me.PatternFrequency = RecurrencePatternFrequency.Daily) Then

                If (Value = RecurrencePatternDaysOfWeek.AllWeekdays) Then
                    Me.optPatternDaysOfWeek_Daily.Value = RecurrencePatternDaysOfWeek.AllWeekdays
                Else
                    Me.optPatternDaysOfWeek_Daily.Value = RecurrencePatternDaysOfWeek.All
                End If

            ElseIf (Me.PatternFrequency = RecurrencePatternFrequency.Weekly) Then

                '	Uncheck all the checkboxes, then check the ones that now apply.
                Me.chkSunday.Checked = False
                Me.chkMonday.Checked = False
                Me.chkTuesday.Checked = False
                Me.chkWednesday.Checked = False
                Me.chkThursday.Checked = False
                Me.chkFriday.Checked = False
                Me.chkSaturday.Checked = False

                '	Check each textbox that has the corrseponding bit set for the
                '	day of the week it represents.
                Dim daysOfWeek As RecurrencePatternDaysOfWeek = Value
                If (daysOfWeek And RecurrencePatternDaysOfWeek.Sunday) = RecurrencePatternDaysOfWeek.Sunday Then Me.chkSunday.Checked = True
                If (daysOfWeek And RecurrencePatternDaysOfWeek.Monday) = RecurrencePatternDaysOfWeek.Monday Then Me.chkMonday.Checked = True
                If (daysOfWeek And RecurrencePatternDaysOfWeek.Tuesday) = RecurrencePatternDaysOfWeek.Tuesday Then Me.chkTuesday.Checked = True
                If (daysOfWeek And RecurrencePatternDaysOfWeek.Wednesday) = RecurrencePatternDaysOfWeek.Wednesday Then Me.chkWednesday.Checked = True
                If (daysOfWeek And RecurrencePatternDaysOfWeek.Thursday) = RecurrencePatternDaysOfWeek.Thursday Then Me.chkThursday.Checked = True
                If (daysOfWeek And RecurrencePatternDaysOfWeek.Friday) = RecurrencePatternDaysOfWeek.Friday Then Me.chkFriday.Checked = True
                If (daysOfWeek And RecurrencePatternDaysOfWeek.Saturday) = RecurrencePatternDaysOfWeek.Saturday Then Me.chkSaturday.Checked = True
            ElseIf (Me.PatternFrequency = RecurrencePatternFrequency.Monthly Or _
                    Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then

                Dim text As String = Value.ToString()
                If Value = RecurrencePatternDaysOfWeek.All Then text = "day"
                If Value = RecurrencePatternDaysOfWeek.AllWeekdays Then text = "weekday"
                If Value = RecurrencePatternDaysOfWeek.AllWeekendDays Then text = "weekend day"

                If (Me.PatternFrequency = RecurrencePatternFrequency.Monthly) Then
                    Me.cboDayOfWeekMonthly.Text = text
                Else
                    Me.cboDayOfWeekYearly.Text = text
                End If

            End If
        End Set
    End Property
#End Region

#Region "OccurrenceStartTime"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the OccurrenceStartTime (cboOccurrenceStartTime)
    ' </summary>
#End If
    Friend Property OccurrenceStartTime() As DateTime

        Get

            Dim startTime As DateTime = DateTime.Parse(Me.cboOccurrenceStartTime.Text)
            Dim startDate As DateTime = Me.RangeStartDate

            Return New DateTime(startDate.Year, startDate.Month, startDate.Day, _
                                startTime.Hour, startTime.Minute, 0)
        End Get

        Set(ByVal Value As DateTime)

            Me.cboOccurrenceStartTime.Text = Value.ToString(Utilities.TimeFormatString)
        End Set
    End Property
#End Region

#Region "OccurrenceDuration"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the OccurrenceDuration (cboOccurrenceDuration).
    ' </summary>
#End If
    Friend Property OccurrenceDuration() As TimeSpan

        Get

            Dim retVal As TimeSpan = TimeSpan.Zero

            Dim itemText As String = Me.cboOccurrenceDuration.Text
            If (Not Me.cboOccurrenceDuration.SelectedItem Is Nothing) Then
                itemText = Me.cboOccurrenceDuration.SelectedItem.ToString()
            End If

            Utilities.ParseTimeSpan(itemText, retVal)

            Return retVal
        End Get

        Set(ByVal Value As TimeSpan)

            Me.cboOccurrenceDuration.Text = Utilities.FormatTimeSpan(Value, False)

        End Set
    End Property
#End Region

#Region "PatternInterval"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the PatternInterval (txtPatternIntervalDaily, txtPatternIntervalWeekly, txtPatternIntervalExplicitMonthly)
    ' Not applicable when PatternFrequency = Yearly.
    ' </summary>
#End If
    Friend Property PatternInterval() As Int32

        Get

            Dim str As String = String.Empty
            Select Case (Me.PatternFrequency)

                Case RecurrencePatternFrequency.Daily
                    str = Me.txtPatternIntervalDaily.Text

                Case RecurrencePatternFrequency.Weekly
                    str = Me.txtPatternIntervalWeekly.Text

                Case RecurrencePatternFrequency.Monthly

                    If Me.PatternType = RecurrencePatternType.Explicit Then
                        str = Me.txtPatternIntervalExplicitMonthly.Text
                    Else
                        str = Me.txtPatternIntervalCalculatedMonthly.Text
                    End If

            End Select

            Return Convert.ToInt32(str)

        End Get

        Set(ByVal Value As Int32)

            Dim newVal As String = Value.ToString()

            Select Case (Me.PatternFrequency)

                Case RecurrencePatternFrequency.Daily

                    Me.txtPatternIntervalDaily.Text = newVal

                Case RecurrencePatternFrequency.Weekly
                    Me.txtPatternIntervalWeekly.Text = newVal

                Case RecurrencePatternFrequency.Monthly
                    Me.txtPatternIntervalExplicitMonthly.Text = newVal
                    Me.txtPatternIntervalCalculatedMonthly.Text = newVal

            End Select
        End Set
    End Property
#End Region

#Region "PatternType"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the PatternType (optExplicitMonthly / optCalculatedMonthly, optExplicitYearly / optCalculatedYearly )
    ' Not applicable when PatternFrequency is Daily or Weekly.
    ' </summary>
#End If
    Friend Property PatternType() As RecurrencePatternType

        Get

            Select Case (Me.PatternFrequency)

                Case RecurrencePatternFrequency.Monthly

                    Return Me.optPatternType_Monthly.Value

                Case RecurrencePatternFrequency.Yearly

                    Return Me.optPatternType_Yearly.Value

                Case Else

                    Return ASFAPPTR.defaultPatternType
            End Select
        End Get

        Set(ByVal Value As RecurrencePatternType)

            Select Case (Me.PatternFrequency)

                Case RecurrencePatternFrequency.Monthly
                    Me.optPatternType_Monthly.Value = Value
                Case RecurrencePatternFrequency.Yearly
                    Me.optPatternType_Yearly.Value = Value
            End Select

        End Set

    End Property
#End Region    '   PatternType

#Region "RangeMaxOccurrences"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the RangeMaxOccurrences (txtRangeMaxOccurrences)
    ' </summary>
#End If
    Friend Property RangeMaxOccurrences() As Int32

        Get

            Return Convert.ToInt32(Me.txtRangeMaxOccurrences.Text)

        End Get


        Set(ByVal Value As Int32)

            Me.txtRangeMaxOccurrences.Text = Value.ToString()

        End Set

    End Property

#End Region

#Region "PatternDayOfMonth"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the PatternDayOfMonth (txtPatternDayOfMonthMonthly, txtPatternDayOfMonthYearly)
    ' Not applicable when PatternFrequency = Daily or Weekly.
    ' </summary>
#End If
    Friend Property PatternDayOfMonth() As Int32

        Get

            If (Me.PatternFrequency = RecurrencePatternFrequency.Monthly) Then
                Return Convert.ToInt32(Me.txtPatternDayOfMonthMonthly.Text)
            Else
                If (Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then
                    Return Convert.ToInt32(Me.txtPatternDayOfMonthYearly.Text)

                    Return ASFAPPTR.defaultPatternDayOfMonth
                End If
            End If

        End Get


        Set(ByVal Value As Int32)
            If (Me.PatternFrequency = RecurrencePatternFrequency.Monthly) Then
                Me.txtPatternDayOfMonthMonthly.Text = Value.ToString()
            Else
                If (Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then
                    Me.txtPatternDayOfMonthYearly.Text = Value.ToString()
                End If
            End If
        End Set

    End Property
#End Region

#Region "PatternOccurrenceOfDayInMonth"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the PatternOccurrenceOfDayInMonth (cboOccurrenceOfDayInMonthMonthly, cboOccurrenceOfDayInMonthYearly)
    ' </summary>
#End If
    Friend Property PatternOccurrenceOfDayInMonth() As RecurrencePatternOccurrenceOfDayInMonth

        Get

            Dim item As ValueListItem = Nothing

            If (Me.PatternFrequency = RecurrencePatternFrequency.Monthly) Then

                item = Me.cboOccurrenceOfDayInMonthMonthly.SelectedItem
                If item Is Nothing Then Return RecurrencePatternOccurrenceOfDayInMonth.None Else Return item.DataValue


            ElseIf (Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then

                item = Me.cboOccurrenceOfDayInMonthYearly.SelectedItem
                If item Is Nothing Then Return RecurrencePatternOccurrenceOfDayInMonth.None Else Return item.DataValue

            End If

            Return RecurrencePatternOccurrenceOfDayInMonth.None


        End Get

        Set(ByVal Value As RecurrencePatternOccurrenceOfDayInMonth)

            Dim index As Int32 = Value - 1

            If (Me.PatternFrequency = RecurrencePatternFrequency.Monthly) Then
                Me.cboOccurrenceOfDayInMonthMonthly.SelectedIndex = index
            Else
                If (Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then
                    Me.cboOccurrenceOfDayInMonthYearly.SelectedIndex = index
                Else
                End If
            End If

        End Set
    End Property

#End Region    ' PatternOccurrenceOfDayInMonth

#Region "PatternMonthOfYear"
#If DEBUG Then
    ' <summary>
    ' Gets/sets the PatternMonthOfYear (cboMonthOfYearExplicitYearly, cboMonthOfYearCalculatedYearly)
    ' Only applicable when PatternFrequency = Yearly
    ' </summary>
#End If
    Friend Property PatternMonthOfYear() As Int32
        Get
            If (Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then

                If (Me.PatternType = RecurrencePatternType.Explicit) Then
                    Return Me.cboMonthOfYearExplicitYearly.SelectedIndex + 1
                Else
                    Return Me.cboMonthOfYearCalculatedYearly.SelectedIndex + 1
                End If

            End If

            Return ASFAPPTR.defaultPatternMonthOfYear

        End Get
        Set(ByVal Value As Int32)

            If (Me.PatternFrequency = RecurrencePatternFrequency.Yearly) Then

                Me.cboMonthOfYearExplicitYearly.SelectedIndex = Value - 1
                Me.cboMonthOfYearCalculatedYearly.SelectedIndex = Value - 1

            End If


        End Set
    End Property
#End Region

#Region "ContentsChanged"
#If DEBUG Then
    ' <summary>
    ' Returns whether any properties of this dialog's
    ' recurrence have changed since it was first displayed
    ' </summary>
#End If
    Friend ReadOnly Property ContentsChanged() As Boolean

        Get

            Return Me.forceContentsChanged Or Utilities.AreRecurrencesEqual(Me.initialAppointmentRecurrence, Me.appointmentRecurrence) = False

        End Get

    End Property

#End Region

#End Region    'Recurrence properties

#Region "Static properties"

#Region "DateFormat"
#If DEBUG Then
    ' <summary>
    ' Returns the format string to be used for dates.
    ' </summary>
#End If
    Friend Shared ReadOnly Property DateFormat() As String

        Get
            Return "ddd " + System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern
        End Get

    End Property
#End Region    ' DateFormat

#End Region    ' Static properties

#End Region  ' Internal Properties

#Region "Private / Internal Methods"

#Region "InitializeDialogControls"
#If DEBUG Then
    ' <summary>
    ' Populates the ComboBoxes, assigns the appropriate localized string
    ' to each control's Text property.
    ' Also sets the dialog controls to their default values, which
    ' are based on the AppointmentRecurrence object this dialog is editing.
    ' </summary>
#End If
    Private Sub InitializeDialogControls()

        Try
            Me.suspendOptionButtonHandling = True

            '	Position the dark separator to be right next to the light one
            Me.lblSepLight.Location = New Point(Me.lblSepDark.Left + 1, Me.lblSepDark.Top)

            '	Assign the recurrence to a stack variable
            Dim recurrence As AppointmentRecurrence = Me.appointmentRecurrence

            '	Populate the ComboBoxes
            Me.PopulateStartTimeCombo()
            Me.PopulateEndTimeCombo(recurrence.OccurrenceStartTime)
            Me.PopulateDaysOfWeekCombos()
            Me.PopulateMonthsOfYearCombos()
            Me.PopulateOccurrenceOfDayInMonthCombos()
            Me.PopulateDurationCombo()

            '	Populate the OptionSet ValueLists
            Me.PopulateOptionSets()

            '	Set UseMnemonics for the OptionSets
            Me.SetUseMnemonicsForOptionSets()

            '	Set each dialog control's value property to the appropriate
            '	value for the AppointRecurrence property it represents

            '	Appointment time properties
            Me.OccurrenceStartTime = recurrence.OccurrenceStartTime
            Me.OccurrenceDuration = recurrence.OccurrenceDuration
            Me.cboEndTime.Text = recurrence.OccurrenceStartTime.Add(recurrence.OccurrenceDuration).ToString(Utilities.TimeFormatString)

            '	Range properties
            Me.RangeStartDate = recurrence.RangeStartDate
            Me.RangeEndDate = recurrence.RangeEndDate
            Me.RangeMaxOccurrences = recurrence.RangeMaxOccurrences
            Me.RangeLimit = recurrence.RangeLimit

            '	Pattern properties
            Me.PatternFrequency = recurrence.PatternFrequency
            Me.PatternInterval = recurrence.PatternInterval
            Me.PatternDaysOfWeek = recurrence.PatternDaysOfWeek
            Me.PatternDayOfMonth = recurrence.PatternDayOfMonth
            Me.PatternMonthOfYear = recurrence.PatternMonthOfYear
            Me.PatternOccurrenceOfDayInMonth = recurrence.PatternOccurrenceOfDayInMonth
            Me.PatternType = recurrence.PatternType

            '	Compensate for the height of the missing option button
            '	if 'allowNoEndDate' is set to false.
            If (Me.allowNoEndDate = False) Then

                Me.optRangeLimit.Top += 20
                Me.optRangeLimit.Height -= 20

            End If
        Finally
            Me.suspendOptionButtonHandling = False
        End Try

    End Sub
#End Region    ' InitializeDialogControls

#Region "HookDialogControlEvents"
#If DEBUG Then
    ' <summary>
    ' Hooks/unhooks the events of interest for the various dialog controls.
    ' </summary>
    ' <param name="unhook">Specifies whether to hook or unhook the event handler.</param>
#End If
    Private Sub HookDialogControlEvents(ByVal unhook As Boolean)

        Me.HookCheckedChanged(Me, unhook)
        Me.HookOptionSetValueChanged(Me, unhook)
        Me.HookTextChanged(Me, unhook)
        Me.HookSelectionChangeCommitted(Me, unhook)
        Me.HookLeave(Me, unhook)

        If (Not unhook) Then

            AddHandler Me.cboOccurrenceStartTime.SelectionChanged, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            AddHandler Me.cboOccurrenceStartTime.SelectionChangeCommitted, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            AddHandler Me.cboOccurrenceDuration.SelectionChanged, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            AddHandler Me.cboOccurrenceDuration.SelectionChangeCommitted, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            AddHandler Me.cboEndTime.SelectionChanged, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            AddHandler Me.cboEndTime.SelectionChangeCommitted, AddressOf Me.OccurrenceTimeComboBox_ItemChanged

        Else

            RemoveHandler Me.cboOccurrenceStartTime.SelectionChanged, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            RemoveHandler Me.cboOccurrenceStartTime.SelectionChangeCommitted, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            RemoveHandler Me.cboOccurrenceDuration.SelectionChanged, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            RemoveHandler Me.cboOccurrenceDuration.SelectionChangeCommitted, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            RemoveHandler Me.cboEndTime.SelectionChanged, AddressOf Me.OccurrenceTimeComboBox_ItemChanged
            RemoveHandler Me.cboEndTime.SelectionChangeCommitted, AddressOf Me.OccurrenceTimeComboBox_ItemChanged

        End If
    End Sub

    Private Sub HookCheckedChanged(ByVal parent As Control, ByVal unhook As Boolean)

        If parent Is Nothing Then Return

        If (parent.Controls.Count > 0) Then

            Dim child As Control
            For Each child In parent.Controls

                If child.GetType() Is GetType(UltraCheckEditor) Then

                    Dim checkBox As UltraCheckEditor = child
                    If (Not checkBox Is Nothing) Then

                        If unhook Then
                            RemoveHandler checkBox.CheckedChanged, AddressOf Me.CheckBox_CheckedChanged
                        Else
                            AddHandler checkBox.CheckedChanged, AddressOf Me.CheckBox_CheckedChanged
                        End If

                    End If

                End If

                '	Call the method recursively on each child so that
                '	their children are set, etc.
                If (child.Controls.Count > 0) Then
                    Me.HookCheckedChanged(child, unhook)
                End If

            Next
        End If

    End Sub

    Private Sub CheckBox_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If (Me.suspendEventHandling) Then Return

        Dim control As Control = sender
        Me.OnValueChangeCommitted(control, PropertyIds.PatternDaysOfWeek)

    End Sub

    Private Sub HookOptionSetValueChanged(ByVal parent As Control, ByVal unhook As Boolean)

        If parent Is Nothing Then Return

        If (parent.Controls.Count > 0) Then

            Dim child As Control
            For Each child In parent.Controls

                If (child.GetType() Is GetType(UltraOptionSet)) Then
                    Dim opt As UltraOptionSet = child
                    If (Not opt Is Nothing) Then

                        If unhook Then
                            RemoveHandler opt.ValueChanged, AddressOf Me.OptionSet_ValueChanged
                        Else
                            AddHandler opt.ValueChanged, AddressOf Me.OptionSet_ValueChanged
                        End If

                    End If
                End If

                '	Call the method recursively on each child so that
                '	their children are set, etc.
                If (child.Controls.Count > 0) Then
                    Me.HookOptionSetValueChanged(child, unhook)
                End If

            Next
        End If

    End Sub


    Private Sub OptionSet_ValueChanged(ByVal sender As Object, ByVal e As EventArgs)

        If (Me.suspendEventHandling) Then Return

        Dim opt As UltraOptionSet = sender
        Dim propId As PropertyIds = PropertyIds.Recurrence

        If (Not opt Is Me.optPatternFrequency AndAlso Me.suspendOptionButtonHandling) Then Return

        If (opt Is Me.optPatternFrequency) Then

            Me.PatternFrequencyOptionSet_ValueChanged(opt)
            propId = PropertyIds.PatternFrequency

        ElseIf (opt Is Me.optRangeLimit) Then

            If opt.Value = RecurrenceRangeLimit.NoLimit Then
                Try
                    Me.suspendEventHandling = True
                    Me.RangeMaxOccurrences = ASFAPPTR.defaultRangeMaxOccurrences
                Finally
                    Me.suspendEventHandling = False
                End Try

                propId = PropertyIds.RangeLimit
            End If

        ElseIf (opt Is Me.optPatternDaysOfWeek_Daily) Then

            If (Me.optPatternDaysOfWeek_Daily.Value = RecurrencePatternDaysOfWeek.AllWeekdays) Then
                Me.PatternInterval = ASFAPPTR.defaultPatternInterval
            End If

            propId = PropertyIds.PatternDaysOfWeek

        ElseIf (opt Is Me.optPatternType_Monthly Or opt Is Me.optPatternType_Yearly) Then

            If (opt Is Me.optPatternType_Monthly) Then
                Me.PatternInterval = ASFAPPTR.defaultPatternInterval
            Else
                Me.PatternMonthOfYear = Me.RangeStartDate.Month

                Me.PatternDayOfMonth = Me.RangeStartDate.Day
                Me.PatternOccurrenceOfDayInMonth = Utilities.GetOccurrenceOfDayInMonth(Me.RangeStartDate)
                Me.PatternDaysOfWeek = Utilities.GetRecurrencePatternDaysOfWeekFromDayOfWeek(Me.RangeStartDate.DayOfWeek)
            End If
            propId = PropertyIds.PatternType

        Else
            Debug.Assert(False, "Finish this")

        End If

        If (Not opt Is Nothing AndAlso Not opt.Value Is Nothing) Then
            Me.OnValueChangeCommitted(opt, propId)
        End If

    End Sub

    Private Sub HookTextChanged(ByVal parent As Control, ByVal unhook As Boolean)

        If parent Is Nothing Then Return

        If (parent.Controls.Count > 0) Then

            Dim child As Control
            For Each child In parent.Controls

                If child.GetType() Is GetType(UltraTextEditor) Or child.GetType() Is GetType(UltraComboEditor) Or child.GetType() Is GetType(UltraCalendarCombo) Then

                    Dim skip As Boolean = False
                    If (child.GetType() Is GetType(UltraComboEditor)) Then
                        Dim ultraComboEditor As UltraComboEditor = child
                        If ultraComboEditor.DropDownStyle = DropDownStyle.DropDownList Then
                            skip = True
                        End If
                    End If


                    If skip = False Then
                        If unhook Then
                            RemoveHandler child.TextChanged, AddressOf Me.TextBasedControl_TextChanged
                        Else
                            AddHandler child.TextChanged, AddressOf Me.TextBasedControl_TextChanged
                        End If
                    End If

                End If

                '	Call the method recursively on each child so that
                '	their children are set, etc.
                If (child.Controls.Count > 0) Then
                    Me.HookTextChanged(child, unhook)
                End If

            Next
        End If

    End Sub

    Private Sub HookSelectionChangeCommitted(ByVal parent As Control, ByVal unhook As Boolean)

        If parent Is Nothing Then Return

        If (parent.Controls.Count > 0) Then

            Dim child As Control
            For Each child In parent.Controls

                If (child.GetType() Is GetType(UltraComboEditor)) Then
                    Dim ultraComboEditor As UltraComboEditor = child
                    If ultraComboEditor.DropDownStyle = DropDownStyle.DropDownList Then
                        If unhook Then
                            RemoveHandler ultraComboEditor.SelectionChangeCommitted, AddressOf Me.DropDownList_SelectionChangeCommitted
                        Else
                            AddHandler ultraComboEditor.SelectionChangeCommitted, AddressOf Me.DropDownList_SelectionChangeCommitted
                        End If
                    End If
                End If

                '	Call the method recursively on each child so that
                '	their children are set, etc.
                If (child.Controls.Count > 0) Then
                    Me.HookSelectionChangeCommitted(child, unhook)
                End If

            Next
        End If

    End Sub

    Private Sub HookLeave(ByVal parent As Control, ByVal unhook As Boolean)

        If parent Is Nothing Then Return

        If (parent.Controls.Count > 0) Then

            Dim child As Control
            For Each child In parent.Controls

                If child.GetType() Is GetType(UltraTextEditor) Or _
                    child.GetType() Is GetType(UltraCheckEditor) Or _
                    child.GetType() Is GetType(UltraComboEditor) Or _
                    child.GetType() Is GetType(UltraCalendarCombo) Then

                    If (child.GetType() Is GetType(UltraComboEditor)) Then
                        Dim ultraComboEditor As UltraComboEditor = child
                        If ultraComboEditor.DropDownStyle = DropDownStyle.DropDownList Then
                            If unhook Then
                                RemoveHandler child.Leave, AddressOf Me.OnLeaveDialogControl
                            Else
                                AddHandler child.Leave, AddressOf Me.OnLeaveDialogControl
                            End If

                        End If
                    End If


                End If

                '	Call the method recursively on each child so that
                '	their children are set, etc.
                If (child.Controls.Count > 0) Then
                    Me.HookLeave(child, unhook)
                End If

            Next
        End If

    End Sub

#End Region    ' HookDialogControlEvents

#Region "UpdateRecurrence"
#If DEBUG Then
    ' <summary>
    ' Updates the applicable properties of the cached AppointmentRecurrence
    ' using values from the corresponding dialog fields. Note that we assume
    ' here that the applicable dialog fields contain valid values, since that is
    ' validated when each control is left.
    ' </summary>
#End If
    Private Sub UpdateRecurrence()

        If (Me.isInitializing Or Me.appointmentRecurrence Is Nothing) Then Return

        Try
            Dim recurrence As AppointmentRecurrence = Me.appointmentRecurrence

            Dim patternFrequency As RecurrencePatternFrequency = Me.PatternFrequency

            Dim monthly As Boolean = (patternFrequency = RecurrencePatternFrequency.Monthly)
            Dim yearly As Boolean = (patternFrequency = RecurrencePatternFrequency.Yearly)

            recurrence.PatternDaysOfWeek = Me.PatternDaysOfWeek
            recurrence.PatternFrequency = Me.PatternFrequency

            recurrence.OccurrenceStartTime = Me.OccurrenceStartTime
            recurrence.OccurrenceDuration = Me.OccurrenceDuration

            If (monthly Or yearly) Then

                recurrence.PatternDayOfMonth = Me.PatternDayOfMonth
                recurrence.PatternOccurrenceOfDayInMonth = Me.PatternOccurrenceOfDayInMonth
                recurrence.PatternType = Me.PatternType
            End If

            If (Not yearly) Then
                recurrence.PatternInterval = Me.PatternInterval
            Else
                recurrence.PatternMonthOfYear = Me.PatternMonthOfYear
            End If

            recurrence.RangeEndDate = Me.RangeEndDate
            recurrence.RangeLimit = Me.RangeLimit
            recurrence.RangeMaxOccurrences = Me.RangeMaxOccurrences
            recurrence.RangeStartDate = Me.RangeStartDate
        Catch
        End Try
    End Sub
#End Region    ' UpdateRecurrence

#Region "PopulateStartTimeCombo"
#If DEBUG Then
    ' <summary>
    ' Populates the ComboBox that contains the values for the OccurrenceStartTime
    ' </summary>
#End If
    Private Sub PopulateStartTimeCombo()

        Me.cboOccurrenceStartTime.Items.Clear()

        Dim theDate As DateTime = DateTime.Today.Date

        Dim i As Int32
        For i = 0 To 47

            Me.cboOccurrenceStartTime.Items.Add(theDate.ToString(Utilities.TimeFormatString))
            theDate = theDate.Add(New TimeSpan(0, 30, 0))
        Next

    End Sub
#End Region    ' PopulateStartTimeCombo

#Region "PopulateEndTimeCombo"
#If DEBUG Then
    ' <summary>
    ' Populates the ComboBox that contains the values for the appointment's end time.
    ' </summary>
#End If
    Private Sub PopulateEndTimeCombo(ByVal startTime As DateTime)

        Me.cboEndTime.Items.Clear()

        '	To emulate MS Outlook, we will show the duration for each
        '	item only if the current duration is less than 1 day, and
        '	begin the run of items with 12AM, rather than the current
        '	start time.
        Dim itemsShowDuration As Boolean = (Me.OccurrenceDuration.TotalDays < 1.0F)

        If (Not itemsShowDuration) Then
            startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, 0, 0, 0)
        End If

        Dim i As Int32
        For i = 0 To 47

            Dim timeSpan As TimeSpan = New TimeSpan(0, i * 30, 0)
            Dim theDate As DateTime = startTime.Add(timeSpan)
            Dim itemText As String = theDate.ToString(Utilities.TimeFormatString)
            If itemsShowDuration Then itemText += " " + Utilities.FormatTimeSpan(timeSpan, True)
            Me.cboEndTime.Items.Add(itemText)
        Next

    End Sub
#End Region    ' PopulateEndTimeCombo

#Region "PopulateDaysOfWeekCombos"
#If DEBUG Then
    ' <summary>
    ' Populates the ComboBoxes that contain the values for the months of the year.
    ' </summary>
#End If
    Private Sub PopulateDaysOfWeekCombos()

        Me.cboDayOfWeekMonthly.Items.Clear()
        Me.cboDayOfWeekYearly.Items.Clear()

        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.All, "day")
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.All, "day")

        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.AllWeekdays, "weekday")
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.AllWeekdays, "weekday")

        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.AllWeekendDays, "weekend day")
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.AllWeekendDays, "weekend day")


        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.Sunday, Utilities.DaysOfWeek(0))
        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.Monday, Utilities.DaysOfWeek(1))
        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.Tuesday, Utilities.DaysOfWeek(2))
        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.Wednesday, Utilities.DaysOfWeek(3))
        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.Thursday, Utilities.DaysOfWeek(4))
        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.Friday, Utilities.DaysOfWeek(5))
        Me.cboDayOfWeekMonthly.Items.Add(RecurrencePatternDaysOfWeek.Saturday, Utilities.DaysOfWeek(6))

        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.Sunday, Utilities.DaysOfWeek(0))
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.Monday, Utilities.DaysOfWeek(1))
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.Tuesday, Utilities.DaysOfWeek(2))
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.Wednesday, Utilities.DaysOfWeek(3))
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.Thursday, Utilities.DaysOfWeek(4))
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.Friday, Utilities.DaysOfWeek(5))
        Me.cboDayOfWeekYearly.Items.Add(RecurrencePatternDaysOfWeek.Saturday, Utilities.DaysOfWeek(6))

    End Sub
#End Region    '   PopulateDaysOfWeekCombos

#Region "PopulateMonthsOfYearCombos"
#If DEBUG Then
    ' <summary>
    ' Populates the ComboBoxes that contain the values for the months of the year.
    ' </summary>
#End If
    Private Sub PopulateMonthsOfYearCombos()

        Me.cboMonthOfYearExplicitYearly.Items.Clear()
        Me.cboMonthOfYearCalculatedYearly.Items.Clear()

        Dim monthsOfYear() As String = Utilities.MonthsOfYear

        Dim i As Int32
        For i = 0 To monthsOfYear.GetLength(0) - 1

            Me.cboMonthOfYearExplicitYearly.Items.Add(monthsOfYear(i))
            Me.cboMonthOfYearCalculatedYearly.Items.Add(monthsOfYear(i))
        Next
    End Sub
#End Region ' PopulateMonthsOfYearCombos

#Region "PopulateOccurrenceOfDayInMonthCombos"
#If DEBUG Then
    ' <summary>
    ' Populates the ComboBoxes that contain the values for the occurrence of the day in the month.
    ' </summary>
#End If
    Private Sub PopulateOccurrenceOfDayInMonthCombos()

        Me.cboOccurrenceOfDayInMonthMonthly.Items.Clear()
        Me.cboOccurrenceOfDayInMonthYearly.Items.Clear()

        Me.cboOccurrenceOfDayInMonthMonthly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.First, "first")
        Me.cboOccurrenceOfDayInMonthYearly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.First, "first")

        Me.cboOccurrenceOfDayInMonthMonthly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Second, "second")
        Me.cboOccurrenceOfDayInMonthYearly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Second, "second")

        Me.cboOccurrenceOfDayInMonthMonthly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Third, "third")
        Me.cboOccurrenceOfDayInMonthYearly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Third, "third")

        Me.cboOccurrenceOfDayInMonthMonthly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Fourth, "fourth")
        Me.cboOccurrenceOfDayInMonthYearly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Fourth, "fourth")

        Me.cboOccurrenceOfDayInMonthMonthly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Last, "last")
        Me.cboOccurrenceOfDayInMonthYearly.Items.Add(RecurrencePatternOccurrenceOfDayInMonth.Last, "last")

    End Sub
#End Region    ' PopulateOccurrenceOfDayInMonthCombos

#Region "PopulateDurationCombo"
#If DEBUG Then
    ' <summary>
    ' Populates the ComboBox that contain the values for the occurrence duration.
    ' </summary>
#End If
    Private Sub PopulateDurationCombo()

        Me.cboOccurrenceDuration.Items.Clear()

        Me.cboOccurrenceDuration.Items.Add("0 minutes")
        Me.cboOccurrenceDuration.Items.Add("5 minutes")
        Me.cboOccurrenceDuration.Items.Add("10 minutes")
        Me.cboOccurrenceDuration.Items.Add("15 minutes")
        Me.cboOccurrenceDuration.Items.Add("30 minutes")
        Me.cboOccurrenceDuration.Items.Add("1 hour")
        Me.cboOccurrenceDuration.Items.Add("2 hours")
        Me.cboOccurrenceDuration.Items.Add("3 hours")
        Me.cboOccurrenceDuration.Items.Add("4 hours")
        Me.cboOccurrenceDuration.Items.Add("5 hours")
        Me.cboOccurrenceDuration.Items.Add("6 hours")
        Me.cboOccurrenceDuration.Items.Add("7 hours")
        Me.cboOccurrenceDuration.Items.Add("8 hours")
        Me.cboOccurrenceDuration.Items.Add("9 hours")
        Me.cboOccurrenceDuration.Items.Add("10 hours")
        Me.cboOccurrenceDuration.Items.Add("11 hours")
        Me.cboOccurrenceDuration.Items.Add(".5 days")
        Me.cboOccurrenceDuration.Items.Add("18 hours")
        Me.cboOccurrenceDuration.Items.Add("1 day")
        Me.cboOccurrenceDuration.Items.Add("2 days")
        Me.cboOccurrenceDuration.Items.Add("3 days")
        Me.cboOccurrenceDuration.Items.Add("4 days")
        Me.cboOccurrenceDuration.Items.Add("1 week")
        Me.cboOccurrenceDuration.Items.Add("2 weeks")

    End Sub
#End Region    ' PopulateDurationCombo

#Region "PopulateOptionSets"
    Private Sub PopulateOptionSets()

        Me.optPatternFrequency.Items.Clear()
        Me.optPatternFrequency.Items.Add(RecurrencePatternFrequency.Daily, "&Daily")
        Me.optPatternFrequency.Items.Add(RecurrencePatternFrequency.Weekly, "&Weekly")
        Me.optPatternFrequency.Items.Add(RecurrencePatternFrequency.Monthly, "&Monthly")
        Me.optPatternFrequency.Items.Add(RecurrencePatternFrequency.Yearly, "&Yearly")

        Me.optRangeLimit.Items.Clear()

        If (Me.allowNoEndDate) Then
            Me.optRangeLimit.Items.Add(RecurrenceRangeLimit.NoLimit, "N&o end date")

            Me.optRangeLimit.Items.Add(RecurrenceRangeLimit.LimitByNumberOfOccurrences, "End a&fter:")
            Me.optRangeLimit.Items.Add(RecurrenceRangeLimit.LimitByDate, "End &by:")

            Me.optPatternDaysOfWeek_Daily.Items.Clear()
            Me.optPatternDaysOfWeek_Daily.Items.Add(RecurrencePatternDaysOfWeek.All, "E&very")
            Me.optPatternDaysOfWeek_Daily.Items.Add(RecurrencePatternDaysOfWeek.AllWeekdays, "Every wee&kday")

            Me.optPatternType_Monthly.Items.Clear()
            Me.optPatternType_Monthly.Items.Add(RecurrencePatternType.Explicit, "D&ay")
            Me.optPatternType_Monthly.Items.Add(RecurrencePatternType.Calculated, "Th&e")

            Me.optPatternType_Yearly.Items.Clear()
            Me.optPatternType_Yearly.Items.Add(RecurrencePatternType.Explicit, "E&very")
            Me.optPatternType_Yearly.Items.Add(RecurrencePatternType.Calculated, "Th&e")
        End If
    End Sub
#End Region    ' PopulateOptionSets

#Region "SetUseMnemonicsForOptionSets"
    Private Sub SetUseMnemonicsForOptionSets()
        Me.optPatternFrequency.UseMnemonics = True
        Me.optRangeLimit.UseMnemonics = True
        Me.optPatternDaysOfWeek_Daily.UseMnemonics = True
        Me.optPatternType_Monthly.UseMnemonics = True
        Me.optPatternType_Yearly.UseMnemonics = True
    End Sub
#End Region

#Region "PatternFrequencyOptionSet_ValueChanged"
#If DEBUG Then
    ' <summary>
    ' Handles the ValueChanged event for the option buttons
    ' that represent the different values for PatternFrequency.
    ' </summary>
    ' <param name="option">The option that was clicked.</param>
#End If
    Private Sub PatternFrequencyOptionSet_ValueChanged(ByVal opt As UltraOptionSet)

        If (Me.suspendEventHandling) Then Return

        If (Not opt Is Nothing) Then

            '	Ensure all the other panels are hidden
            Me.pnlDaily.Visible = False
            Me.pnlWeekly.Visible = False
            Me.pnlMonthly.Visible = False
            Me.pnlYearly.Visible = False

            Dim activePanel As Panel = Nothing

            '	Set the new panel based on which option was clicked
            Select Case opt.Value
                Case RecurrencePatternFrequency.Daily
                    activePanel = Me.pnlDaily
                Case RecurrencePatternFrequency.Weekly
                    activePanel = Me.pnlWeekly
                Case RecurrencePatternFrequency.Monthly
                    activePanel = Me.pnlMonthly
                Case RecurrencePatternFrequency.Yearly
                    activePanel = Me.pnlYearly
            End Select


            If (Not activePanel Is Nothing) Then

                '	Set the size and location of the panel
                activePanel.Location = Me.PanelLocation
                activePanel.Size = Me.PanelSize

                '	Set the parent to the 'Recurrence Pattern' frame
                activePanel.Parent = Me.fraRecurrencePattern

                '	Adjust the panel's TabIndex so it immediately follows
                '	the option to which it corresponds.
                activePanel.TabIndex = opt.TabIndex + 1

                '	Reset the control values to their defaults each time
                '	a new panel is displayed.
                Me.ResetRecurrencePatternValues(Me.PatternFrequency)

                '	We will deviate from MS Outlook here...when a new
                '	PatternFrequency is selected, if the existing occurrence
                '	duration is too big, we will reduce it to the maximum
                '	allowable value for that PatternFrequency.
                Me.VerifyCurrentDuration()

                '	Since we supress updates when showing a new panel, update
                '	the range end date now that the panel is up to date
                Me.UpdateRecurrence()
                Me.UpdateRangeEndDate()

                '	Show the panel
                activePanel.Visible = True
            End If
        End If

    End Sub
#End Region    ' PatternFrequencyOptionSet_ValueChanged

#Region "TextBasedControl_TextChanged"
#If DEBUG Then
    ' <summary>
    ' Handles the TextChanged event for the text-based controls.
    ' </summary>
    ' <param name="sender">The control whose text has changed.</param>
    ' <param name="e">Event arguments</param>
#End If
    Private Sub TextBasedControl_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        If (Me.suspendEventHandling) Then Return

        If (sender Is Me.txtPatternIntervalDaily) Then
            Me.optPatternDaysOfWeek_Daily.Value = RecurrencePatternDaysOfWeek.All
        ElseIf (sender Is Me.txtPatternIntervalWeekly) Then
            If Not Me.isInitializing AndAlso Not Me.suspendUpdate Then Me.RangeEndDate = Me.appointmentRecurrence.RangeEndDate
        ElseIf (sender Is Me.txtPatternDayOfMonthMonthly Or sender Is Me.txtPatternIntervalExplicitMonthly) Then
            Me.optPatternType_Monthly.Value = RecurrencePatternType.Explicit
        ElseIf (sender Is Me.txtPatternIntervalCalculatedMonthly) Then
            Me.optPatternType_Monthly.Value = RecurrencePatternType.Calculated
        ElseIf (sender Is Me.txtPatternDayOfMonthYearly) Then
            Me.optPatternType_Yearly.Value = RecurrencePatternType.Explicit
        ElseIf (sender Is Me.txtRangeMaxOccurrences) Then
            Me.optRangeLimit.Value = RecurrenceRangeLimit.LimitByNumberOfOccurrences
        End If


        If (sender Is Me.dtpRangeStartDate) Then
            Me.UpdateRangeEndDate()
        ElseIf (sender Is Me.dtpRangeEndDate) Then
            '	Only do this when the DateTimePicker has focus
            If (Me.dtpRangeEndDate.Controls.Count > 0 AndAlso Me.ActiveControl Is Me.dtpRangeEndDate.Controls(0)) Then
                Me.optRangeLimit.Value = RecurrenceRangeLimit.LimitByDate
            End If
        End If
    End Sub

#End Region    ' TextBasedControl_TextChanged

#Region "OnLeaveDialogControl"
#If DEBUG Then
    ' <summary>
    ' Handles the Leave event for all text-based edit controls on the form.
    ' </summary>
    ' <param name="sender"></param>
    ' <param name="e"></param>
#End If
    Private Sub OnLeaveDialogControl(ByVal sender As Object, ByVal e As EventArgs)

        If (Me.suspendEventHandling) Then Return

        Dim control As Control = sender

        If (sender Is Me.cboOccurrenceDuration) Then
            Me.OnValueChangeCommitted(control, PropertyIds.OccurrenceDuration)
        ElseIf (sender Is Me.cboOccurrenceStartTime) Then
            Me.OnValueChangeCommitted(control, PropertyIds.OccurrenceStartTime)
        ElseIf (sender Is Me.cboEndTime) Then
            '	Note that since there is no 'OccurrenceEndTime' property,
            '	we will notify as a change in OccurrenceDuration, since
            '	that is the thing we need to reevaluate.
            Me.OnValueChangeCommitted(control, PropertyIds.OccurrenceDuration)
        ElseIf (sender Is Me.txtPatternDayOfMonthMonthly Or sender Is Me.txtPatternDayOfMonthYearly) Then
            Me.OnValueChangeCommitted(control, PropertyIds.PatternDayOfMonth)
        ElseIf (sender Is Me.cboDayOfWeekMonthly Or sender Is Me.cboDayOfWeekYearly) Then
            Me.OnValueChangeCommitted(control, PropertyIds.PatternDaysOfWeek)
        ElseIf (sender Is Me.txtPatternIntervalDaily Or sender Is Me.txtPatternIntervalWeekly Or sender Is Me.txtPatternIntervalExplicitMonthly Or sender Is Me.txtPatternIntervalCalculatedMonthly) Then
            Me.OnValueChangeCommitted(control, PropertyIds.PatternInterval)
        ElseIf (sender Is Me.cboMonthOfYearExplicitYearly Or sender Is Me.cboMonthOfYearCalculatedYearly) Then
            Me.OnValueChangeCommitted(control, PropertyIds.PatternMonthOfYear)
        ElseIf (sender Is Me.cboOccurrenceOfDayInMonthMonthly Or sender Is Me.cboOccurrenceOfDayInMonthYearly) Then
            Me.OnValueChangeCommitted(control, PropertyIds.PatternOccurrenceOfDayInMonth)
        ElseIf (sender Is Me.dtpRangeEndDate) Then
            Me.OnValueChangeCommitted(control, PropertyIds.RangeEndDate)
        ElseIf (sender Is Me.txtRangeMaxOccurrences) Then
            Me.OnValueChangeCommitted(control, PropertyIds.RangeMaxOccurrences)
        ElseIf (sender Is Me.dtpRangeStartDate) Then
            Me.OnValueChangeCommitted(control, PropertyIds.RangeStartDate)
        End If
    End Sub


#If DEBUG Then
    ' <summary>
    ' Called when a change to a dialog control has been committed, i.e., the control was left.
    ' This method is used to validate user input when each dialog control is left, the user's
    ' entry is converted to the appropriate data type for the corresponding AppointmentRecurrence
    ' property, and then that property is set on the temporary AppointmentRecurrence object we use
    ' for validation. If the input is valid, the RangeEndDate is updated if it isn't, the error
    ' dialog is displayed, and the control is focused and its text selected.
    ' </summary>
    ' <param name="sender">The control that was left.</param>
    ' <param name="recurrencePropId">The property identifier of the AppointmentRecurrence property that has logically changed.</param>
#End If
    Private Sub OnValueChangeCommitted(ByVal sender As Control, ByVal recurrencePropId As PropertyIds)

        If (sender Is Nothing) Then Return

        If (Me.suspendUpdate Or Me.isInitializing) Then Return

        Me.contentsAreValid = True

        '	We wrap this in a try/catch because the attempt to convert
        '	the dialog field's value could fail, and also, the attempt
        '	to set the property on the AppointmentRecurrence could fail.
        Dim displayGenericError As Boolean = True
        Try

            If (sender Is Me.cboOccurrenceStartTime Or sender Is Me.cboOccurrenceDuration Or sender Is Me.cboEndTime) Then

                If (sender Is Me.cboOccurrenceStartTime) Then

                    '	When the start time changes, we update the end time		
                    Me.UpdateEndTime()

                    Me.appointmentRecurrence.OccurrenceStartTime = Me.OccurrenceStartTime

                ElseIf (sender Is Me.cboOccurrenceDuration) Then

                    '	When the duration changes, we update the end time

                    '	Throw an exception if the new duration is invalid
                    Dim shortestDuration As TimeSpan = Utilities.GetShortestDurationFromDaysOfWeek(Me.PatternDaysOfWeek)
                    Dim newDuration As TimeSpan = Me.OccurrenceDuration
                    If (shortestDuration.TotalMinutes < newDuration.TotalMinutes) Then

                        displayGenericError = False
                        Throw New Exception(ASFAPPTR.RecurrenceDialog_Error_DurationCannotExceedFrequency)

                    End If

                    '	Set the end time combo's new text
                    Dim startTime As DateTime = Me.OccurrenceStartTime

                    '	Repopulate the end time combo
                    Me.PopulateEndTimeCombo(startTime)

                    Me.cboEndTime.Text = startTime.Add(newDuration).ToString(Utilities.TimeFormatString)

                ElseIf (sender Is Me.cboEndTime) Then

                    '	When the end time changes, we update the duration
                    Me.UpdateOccurrenceDuration()


                End If
            End If

            Select Case (recurrencePropId)

                Case PropertyIds.PatternDayOfMonth

                    If (Me.PatternDayOfMonth > 28 AndAlso Me.PatternDayOfMonth < 32) Then

                        Me.patternDayOfMonthWarningDisplayed = True

                        MessageBox.Show(ASFAPPTR.RecurrenceDialog_Warning_PatternDayOfMonthExceeds28, ASFAPPTR.RecurrenceDialog_MessageBox_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    End If

                    Me.appointmentRecurrence.PatternDayOfMonth = Me.PatternDayOfMonth



                Case PropertyIds.PatternDaysOfWeek

                    Dim shortestDuration As TimeSpan = Utilities.GetShortestDurationFromDaysOfWeek(Me.PatternDaysOfWeek)
                    If (shortestDuration.TotalMinutes < Me.OccurrenceDuration.TotalMinutes) Then

                        displayGenericError = False
                        Throw New Exception(ASFAPPTR.RecurrenceDialog_Error_DurationCannotExceedFrequency)

                    End If

                    Me.appointmentRecurrence.PatternDaysOfWeek = Me.PatternDaysOfWeek



                Case PropertyIds.PatternFrequency

                    Me.appointmentRecurrence.PatternFrequency = Me.PatternFrequency

                Case PropertyIds.PatternInterval

                    Me.appointmentRecurrence.PatternInterval = Me.PatternInterval

                Case PropertyIds.PatternMonthOfYear

                    Me.appointmentRecurrence.PatternMonthOfYear = Me.PatternMonthOfYear

                Case PropertyIds.PatternOccurrenceOfDayInMonth

                    Me.appointmentRecurrence.PatternOccurrenceOfDayInMonth = Me.PatternOccurrenceOfDayInMonth

                Case PropertyIds.PatternType

                    Me.appointmentRecurrence.PatternType = Me.PatternType

                Case PropertyIds.RangeEndDate

                    Me.appointmentRecurrence.RangeEndDate = Me.RangeEndDate

                Case PropertyIds.RangeLimit

                    Me.appointmentRecurrence.RangeLimit = Me.RangeLimit

                Case PropertyIds.RangeMaxOccurrences

                    Me.appointmentRecurrence.RangeMaxOccurrences = Me.RangeMaxOccurrences

                Case PropertyIds.RangeStartDate

                    Me.appointmentRecurrence.RangeStartDate = Me.RangeStartDate

            End Select

            If (recurrencePropId <> PropertyIds.RangeEndDate) Then

                Me.UpdateRangeEndDate()

            End If

        Catch ex As Exception

            Me.contentsAreValid = False

            '	Some errors are more specific than "The recurrence pattern is not valid"
            '	in these cases, display the exception message.
            Dim errorMessage As String = ex.Message
            If displayGenericError Then errorMessage = ASFAPPTR.RecurrenceDialog_Error_RecurrencePatterIsInvalid

            '	Inform the end user of the error condition.
            MessageBox.Show(errorMessage, _
                 ASFAPPTR.RecurrenceDialog_MessageBox_Caption, _
                 MessageBoxButtons.OK, _
                 MessageBoxIcon.Exclamation)

            '	Select the dialog control
            sender.Select()

        End Try

    End Sub


#End Region ' OnLeaveDialogControl

#Region "UpdateEndTime"
#If DEBUG Then
    ' <summary>
    ' Updates the value of the dialog's cboEndTime field.
    ' </summary>
#End If
    Private Sub UpdateEndTime()

        Try

            '	Parse the value of the start time combo
            Dim tempDate As DateTime = DateTime.Parse(Me.cboOccurrenceStartTime.Text)

            '	Repopulate the end time combo
            Me.PopulateEndTimeCombo(tempDate)

            '	Set the end time combo's new text
            Me.cboEndTime.Text = tempDate.Add(Me.OccurrenceDuration).ToString(Utilities.TimeFormatString)

        Catch
        End Try

    End Sub

#End Region    ' UpdateEndTime

#Region "UpdateOccurrenceDuration"
#If DEBUG Then
    ' <summary>
    ' Updates the value of the dialog's OccurrenceDuration field.
    ' </summary>
#End If
    Private Sub UpdateOccurrenceDuration()

        Try

            '	Get the integral number of days in the current duration
            Dim durationInDays As Int32 = Me.OccurrenceDuration.TotalDays

            '	Get the delta between the start and end time ComboBoxes
            Dim delta As TimeSpan = Me.StartTimeEndTimeDelta

            Dim newDuration As TimeSpan = New TimeSpan(0, ((durationInDays * 1440) + (delta.TotalMinutes)), 0)

            '	If the new duration is negative, we take the difference between
            '	1 full day and that duration.
            If (newDuration.TotalMinutes < 0) Then
                newDuration = New TimeSpan(0, (1440.0 + newDuration.TotalMinutes), 0)
            End If

            Me.OccurrenceDuration = newDuration

            Me.appointmentRecurrence.OccurrenceDuration = Me.OccurrenceDuration

        Catch

        End Try
    End Sub

#End Region    ' UpdateOccurrenceDuration

#Region "UpdateRangeEndDate"
#If DEBUG Then
    ' <summary>
    ' Updates the RangeEndDate to reflect the new default RangeEndDate
    ' based on the current pattern criteria.
    ' </summary>
#End If
    Private Sub UpdateRangeEndDate()
        If (Not Me.isInitializing AndAlso Not Me.suspendUpdate) Then Me.RangeEndDate = Me.appointmentRecurrence.RangeEndDate
    End Sub
#End Region    ' UpdateRangeEndDate

#Region "StartTimeEndTimeDelta"
#If DEBUG Then
    ' <summary>
    ' Returns the amount of time between the parsed contents of
    ' the start and end time ComboBoxes. If the end time precedes
    ' the start time, a negative value is returned.
    ' </summary>
#End If
    Private ReadOnly Property StartTimeEndTimeDelta() As TimeSpan

        Get
            Try
                Dim startTime As DateTime = DateTime.Parse(Me.cboOccurrenceStartTime.Text)
                Dim endTimeText As String = Me.cboEndTime.Text
                Dim pos As Int32 = endTimeText.IndexOf(" (")
                If (pos >= 0) Then endTimeText = endTimeText.Substring(0, pos)

                Dim endTime As DateTime = DateTime.Parse(endTimeText)

                Return endTime.Subtract(startTime)

            Catch
                Return TimeSpan.Zero
            End Try
        End Get
    End Property

#End Region ' StartTimeEndTimeDelta


#End Region ' Private / Internal Methods


#Region "EventHandlers"

    Private Sub OccurrenceTimeComboBox_ItemChanged(ByVal sender As Object, ByVal e As EventArgs)
        If (Me.suspendEventHandling Or Me.isInitializing) Then Return

        '	Invoke the OnAfterEndTimeSelected method so we can strip
        '	the duration out of the item text so we only see the time
        '	in the edit portion.
        If (sender Is Me.cboEndTime) Then
            Me.BeginInvoke(New MethodInvoker(AddressOf Me.OnAfterEndTimeSelected))
        ElseIf (sender Is Me.cboOccurrenceDuration Or sender Is Me.cboOccurrenceStartTime) Then
            Me.UpdateEndTime()
        End If

    End Sub

    Private Sub OnAfterEndTimeSelected()

        If (Me.cboEndTime.SelectedIndex <> -1) Then

            Dim text As String = Me.cboEndTime.SelectedItem.DisplayText
            If (text = Me.cboEndTime.Text) Then

                Dim theEnd = text.IndexOf(" (")
                If (theEnd >= 0) Then
                    text = text.Substring(0, theEnd)

                    Me.cboEndTime.Text = text

                    Me.UpdateOccurrenceDuration()
                End If
            End If
        End If

    End Sub


#Region "DropDownList_SelectionChangeCommitted"
#If DEBUG Then
    ' <summary>
    ' Handles the SelectionChangeCommitted event for the DropDownList-style ComboBoxes.
    ' </summary>
    ' <param name="sender">The ComboBox whose selection change was committed.</param>
    ' <param name="e">Event arguments</param>
#End If
    Private Sub DropDownList_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs)

        If (Me.suspendEventHandling) Then Return

        If (sender Is Me.cboOccurrenceOfDayInMonthMonthly Or sender Is Me.cboDayOfWeekMonthly) Then
            Me.optPatternType_Monthly.Value = RecurrencePatternType.Calculated
        ElseIf (sender Is Me.cboMonthOfYearExplicitYearly) Then
            Me.optPatternType_Yearly.Value = RecurrencePatternType.Explicit
        ElseIf (sender Is Me.cboOccurrenceOfDayInMonthYearly Or sender Is Me.cboDayOfWeekYearly Or sender Is Me.cboMonthOfYearCalculatedYearly) Then
            Me.optPatternType_Yearly.Value = RecurrencePatternType.Calculated
        End If

    End Sub
#End Region ' DropDownList_SelectionChangeCommitted

#End Region ' EventHandlers

#Region "ResetRecurrencePatternValues"
#If DEBUG Then
    ' <summary>
    ' Restores the values of the controls on the specified Panel to their default values.
    ' </summary>
    ' <param name="patternFrequency">The PatternFrequency that identifies that Panel whose control values are to be reset.</param>
#End If
    Private Sub ResetRecurrencePatternValues(ByVal patternFrequency As RecurrencePatternFrequency)

        Try

            '	Suspend updates while we are initializing the new panel's controls
            Me.suspendUpdate = True

            Dim recurrence As AppointmentRecurrence = Me.appointmentRecurrence
            Select Case (patternFrequency)

                Case RecurrencePatternFrequency.Daily, RecurrencePatternFrequency.Weekly

                    Me.PatternInterval = ASFAPPTR.defaultPatternInterval
                    If patternFrequency = RecurrencePatternFrequency.Daily Then
                        Me.PatternDaysOfWeek = RecurrencePatternDaysOfWeek.All
                    Else
                        Me.PatternDaysOfWeek = Utilities.GetRecurrencePatternDaysOfWeekFromDayOfWeek(Me.RangeStartDate.DayOfWeek)
                    End If

                Case RecurrencePatternFrequency.Monthly, RecurrencePatternFrequency.Yearly

                    If (patternFrequency = RecurrencePatternFrequency.Monthly) Then
                        Me.PatternInterval = ASFAPPTR.defaultPatternInterval
                    Else
                        Me.PatternMonthOfYear = Me.RangeStartDate.Month
                    End If

                    Me.PatternDayOfMonth = Me.RangeStartDate.Day
                    Me.PatternOccurrenceOfDayInMonth = Utilities.GetOccurrenceOfDayInMonth(Me.RangeStartDate)
                    Me.PatternDaysOfWeek = Utilities.GetRecurrencePatternDaysOfWeekFromDayOfWeek(Me.RangeStartDate.DayOfWeek)
                    Me.PatternType = RecurrencePatternType.Explicit
            End Select

        Finally

            '	Resume updates
            Me.suspendUpdate = False

        End Try
    End Sub

#End Region ' ResetRecurrencePatternValues

#Region "VerifyCurrentDuration"
#If DEBUG Then
    ' <summary>
    ' Verifies that the current value for the dialog's OccurrenceDuration
    ' property is valid and if it isn't, sets it to the maximum allowable
    ' value given the current PatternFrequency.
    ' </summary>
#End If
    Private Sub VerifyCurrentDuration()

        Dim occurrenceDuration As TimeSpan = Me.OccurrenceDuration
        Dim maxAllowableDays As Double = 0.0F
        Select Case (Me.PatternFrequency)

            Case RecurrencePatternFrequency.Daily
                maxAllowableDays = 1.0F

            Case RecurrencePatternFrequency.Weekly
                maxAllowableDays = 7.0F

            Case RecurrencePatternFrequency.Monthly
                maxAllowableDays = 30.0F

            Case RecurrencePatternFrequency.Yearly
                maxAllowableDays = 365.0F

        End Select
        If (occurrenceDuration.TotalDays > maxAllowableDays) Then
            Me.OccurrenceDuration = New TimeSpan(maxAllowableDays, 0, 0, 0)
        End If

    End Sub
#End Region    ' VerifyCurrentDuration


    Private Sub InitializeComponent()
        Dim ValueListItem1 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem2 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem3 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem4 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim ValueListItem5 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem6 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim DateButton1 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim DateButton2 As Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton = New Infragistics.Win.UltraWinSchedule.CalendarCombo.DateButton
        Dim ValueListItem7 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem8 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem9 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem10 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem11 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem12 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Dim ValueListItem13 As Infragistics.Win.ValueListItem = New Infragistics.Win.ValueListItem
        Me.cmdOk = New Infragistics.Win.Misc.UltraButton
        Me.pnlDaily = New System.Windows.Forms.Panel
        Me.lblPatternIntervalDaily = New Infragistics.Win.Misc.UltraLabel
        Me.txtPatternIntervalDaily = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.optPatternDaysOfWeek_Daily = New Infragistics.Win.UltraWinEditors.UltraOptionSet
        Me.pnlMonthly = New System.Windows.Forms.Panel
        Me.lblPatternInterval4Monthly = New Infragistics.Win.Misc.UltraLabel
        Me.txtPatternIntervalCalculatedMonthly = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.lblPatternInterval3Monthly = New Infragistics.Win.Misc.UltraLabel
        Me.cboDayOfWeekMonthly = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.cboOccurrenceOfDayInMonthMonthly = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.txtPatternIntervalExplicitMonthly = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.lblPatternInterval2Monthly = New Infragistics.Win.Misc.UltraLabel
        Me.lblPatternInterval1Monthly = New Infragistics.Win.Misc.UltraLabel
        Me.txtPatternDayOfMonthMonthly = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.optPatternType_Monthly = New Infragistics.Win.UltraWinEditors.UltraOptionSet
        Me.cmdRemoveRecurrence = New Infragistics.Win.Misc.UltraButton
        Me.cmdCancel = New Infragistics.Win.Misc.UltraButton
        Me.pnlYearly = New System.Windows.Forms.Panel
        Me.cboMonthOfYearCalculatedYearly = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.lblCalculatedYearly = New Infragistics.Win.Misc.UltraLabel
        Me.cboDayOfWeekYearly = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.cboOccurrenceOfDayInMonthYearly = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.txtPatternDayOfMonthYearly = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.cboMonthOfYearExplicitYearly = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.optPatternType_Yearly = New Infragistics.Win.UltraWinEditors.UltraOptionSet
        Me.pnlWeekly = New System.Windows.Forms.Panel
        Me.chkSaturday = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkFriday = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkThursday = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkWednesday = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkTuesday = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkMonday = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.chkSunday = New Infragistics.Win.UltraWinEditors.UltraCheckEditor
        Me.lblPatternInterval2Weekly = New Infragistics.Win.Misc.UltraLabel
        Me.txtPatternIntervalWeekly = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.lblPatternInterval1_Weekly = New Infragistics.Win.Misc.UltraLabel
        Me.fraAppointmentTime = New System.Windows.Forms.GroupBox
        Me.cboOccurrenceDuration = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.lblOccurrenceDuration = New Infragistics.Win.Misc.UltraLabel
        Me.cboEndTime = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.lblEndTime = New Infragistics.Win.Misc.UltraLabel
        Me.cboOccurrenceStartTime = New Infragistics.Win.UltraWinEditors.UltraComboEditor
        Me.lblOccurrenceStartTime = New Infragistics.Win.Misc.UltraLabel
        Me.fraRecurrenceRange = New System.Windows.Forms.GroupBox
        Me.dtpRangeEndDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.lblRangeMaxOccurrences = New Infragistics.Win.Misc.UltraLabel
        Me.txtRangeMaxOccurrences = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.dtpRangeStartDate = New Infragistics.Win.UltraWinSchedule.UltraCalendarCombo
        Me.lblRangeStartDate = New Infragistics.Win.Misc.UltraLabel
        Me.optRangeLimit = New Infragistics.Win.UltraWinEditors.UltraOptionSet
        Me.fraRecurrencePattern = New System.Windows.Forms.GroupBox
        Me.optPatternFrequency = New Infragistics.Win.UltraWinEditors.UltraOptionSet
        Me.lblSepDark = New Infragistics.Win.Misc.UltraLabel
        Me.lblSepLight = New Infragistics.Win.Misc.UltraLabel
        Me.txtTickler = New Infragistics.Win.UltraWinEditors.UltraTextEditor
        Me.UltraLabel1 = New Infragistics.Win.Misc.UltraLabel
        Me.pnlDaily.SuspendLayout()
        CType(Me.txtPatternIntervalDaily, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optPatternDaysOfWeek_Daily, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMonthly.SuspendLayout()
        CType(Me.txtPatternIntervalCalculatedMonthly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboDayOfWeekMonthly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboOccurrenceOfDayInMonthMonthly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPatternIntervalExplicitMonthly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPatternDayOfMonthMonthly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optPatternType_Monthly, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlYearly.SuspendLayout()
        CType(Me.cboMonthOfYearCalculatedYearly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboDayOfWeekYearly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboOccurrenceOfDayInMonthYearly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtPatternDayOfMonthYearly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboMonthOfYearExplicitYearly, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optPatternType_Yearly, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlWeekly.SuspendLayout()
        CType(Me.txtPatternIntervalWeekly, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraAppointmentTime.SuspendLayout()
        CType(Me.cboOccurrenceDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboEndTime, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboOccurrenceStartTime, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraRecurrenceRange.SuspendLayout()
        CType(Me.dtpRangeEndDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtRangeMaxOccurrences, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dtpRangeStartDate, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.optRangeLimit, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraRecurrencePattern.SuspendLayout()
        CType(Me.optPatternFrequency, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.txtTickler, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdOk
        '
        Me.cmdOk.Location = New System.Drawing.Point(292, 281)
        Me.cmdOk.Name = "cmdOk"
        Me.cmdOk.Size = New System.Drawing.Size(75, 23)
        Me.cmdOk.TabIndex = 14
        Me.cmdOk.Text = "OK"
        '
        'pnlDaily
        '
        Me.pnlDaily.Controls.Add(Me.lblPatternIntervalDaily)
        Me.pnlDaily.Controls.Add(Me.txtPatternIntervalDaily)
        Me.pnlDaily.Controls.Add(Me.optPatternDaysOfWeek_Daily)
        Me.pnlDaily.Location = New System.Drawing.Point(488, 35)
        Me.pnlDaily.Name = "pnlDaily"
        Me.pnlDaily.Size = New System.Drawing.Size(350, 96)
        Me.pnlDaily.TabIndex = 19
        Me.pnlDaily.Visible = False
        '
        'lblPatternIntervalDaily
        '
        Me.lblPatternIntervalDaily.AutoSize = True
        Me.lblPatternIntervalDaily.Location = New System.Drawing.Point(124, 20)
        Me.lblPatternIntervalDaily.Name = "lblPatternIntervalDaily"
        Me.lblPatternIntervalDaily.Size = New System.Drawing.Size(36, 14)
        Me.lblPatternIntervalDaily.TabIndex = 2
        Me.lblPatternIntervalDaily.Text = "day(s)"
        '
        'txtPatternIntervalDaily
        '
        Me.txtPatternIntervalDaily.Location = New System.Drawing.Point(80, 16)
        Me.txtPatternIntervalDaily.Name = "txtPatternIntervalDaily"
        Me.txtPatternIntervalDaily.Size = New System.Drawing.Size(35, 22)
        Me.txtPatternIntervalDaily.TabIndex = 1
        '
        'optPatternDaysOfWeek_Daily
        '
        Me.optPatternDaysOfWeek_Daily.BackColor = System.Drawing.Color.Transparent
        Me.optPatternDaysOfWeek_Daily.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem1.DataValue = "ValueListItem0"
        ValueListItem1.DisplayText = "Every"
        ValueListItem2.DataValue = "ValueListItem1"
        ValueListItem2.DisplayText = "Every weekday"
        Me.optPatternDaysOfWeek_Daily.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem1, ValueListItem2})
        Me.optPatternDaysOfWeek_Daily.ItemSpacingVertical = 12
        Me.optPatternDaysOfWeek_Daily.Location = New System.Drawing.Point(16, 15)
        Me.optPatternDaysOfWeek_Daily.Name = "optPatternDaysOfWeek_Daily"
        Me.optPatternDaysOfWeek_Daily.Size = New System.Drawing.Size(96, 54)
        Me.optPatternDaysOfWeek_Daily.TabIndex = 11
        '
        'pnlMonthly
        '
        Me.pnlMonthly.Controls.Add(Me.lblPatternInterval4Monthly)
        Me.pnlMonthly.Controls.Add(Me.txtPatternIntervalCalculatedMonthly)
        Me.pnlMonthly.Controls.Add(Me.lblPatternInterval3Monthly)
        Me.pnlMonthly.Controls.Add(Me.cboDayOfWeekMonthly)
        Me.pnlMonthly.Controls.Add(Me.cboOccurrenceOfDayInMonthMonthly)
        Me.pnlMonthly.Controls.Add(Me.txtPatternIntervalExplicitMonthly)
        Me.pnlMonthly.Controls.Add(Me.lblPatternInterval2Monthly)
        Me.pnlMonthly.Controls.Add(Me.lblPatternInterval1Monthly)
        Me.pnlMonthly.Controls.Add(Me.txtPatternDayOfMonthMonthly)
        Me.pnlMonthly.Controls.Add(Me.optPatternType_Monthly)
        Me.pnlMonthly.Location = New System.Drawing.Point(488, 247)
        Me.pnlMonthly.Name = "pnlMonthly"
        Me.pnlMonthly.Size = New System.Drawing.Size(350, 96)
        Me.pnlMonthly.TabIndex = 18
        Me.pnlMonthly.Visible = False
        '
        'lblPatternInterval4Monthly
        '
        Me.lblPatternInterval4Monthly.AutoSize = True
        Me.lblPatternInterval4Monthly.Location = New System.Drawing.Point(296, 50)
        Me.lblPatternInterval4Monthly.Name = "lblPatternInterval4Monthly"
        Me.lblPatternInterval4Monthly.Size = New System.Drawing.Size(49, 14)
        Me.lblPatternInterval4Monthly.TabIndex = 10
        Me.lblPatternInterval4Monthly.Text = "month(s)"
        '
        'txtPatternIntervalCalculatedMonthly
        '
        Me.txtPatternIntervalCalculatedMonthly.Location = New System.Drawing.Point(260, 48)
        Me.txtPatternIntervalCalculatedMonthly.Name = "txtPatternIntervalCalculatedMonthly"
        Me.txtPatternIntervalCalculatedMonthly.Size = New System.Drawing.Size(35, 22)
        Me.txtPatternIntervalCalculatedMonthly.TabIndex = 9
        '
        'lblPatternInterval3Monthly
        '
        Me.lblPatternInterval3Monthly.AutoSize = True
        Me.lblPatternInterval3Monthly.Location = New System.Drawing.Point(216, 51)
        Me.lblPatternInterval3Monthly.Name = "lblPatternInterval3Monthly"
        Me.lblPatternInterval3Monthly.Size = New System.Drawing.Size(44, 14)
        Me.lblPatternInterval3Monthly.TabIndex = 8
        Me.lblPatternInterval3Monthly.Text = "of every"
        '
        'cboDayOfWeekMonthly
        '
        Me.cboDayOfWeekMonthly.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cboDayOfWeekMonthly.Location = New System.Drawing.Point(130, 48)
        Me.cboDayOfWeekMonthly.MaxDropDownItems = 10
        Me.cboDayOfWeekMonthly.Name = "cboDayOfWeekMonthly"
        Me.cboDayOfWeekMonthly.Size = New System.Drawing.Size(86, 22)
        Me.cboDayOfWeekMonthly.TabIndex = 7
        '
        'cboOccurrenceOfDayInMonthMonthly
        '
        Me.cboOccurrenceOfDayInMonthMonthly.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cboOccurrenceOfDayInMonthMonthly.Location = New System.Drawing.Point(58, 48)
        Me.cboOccurrenceOfDayInMonthMonthly.Name = "cboOccurrenceOfDayInMonthMonthly"
        Me.cboOccurrenceOfDayInMonthMonthly.Size = New System.Drawing.Size(70, 22)
        Me.cboOccurrenceOfDayInMonthMonthly.TabIndex = 6
        '
        'txtPatternIntervalExplicitMonthly
        '
        Me.txtPatternIntervalExplicitMonthly.Location = New System.Drawing.Point(158, 14)
        Me.txtPatternIntervalExplicitMonthly.Name = "txtPatternIntervalExplicitMonthly"
        Me.txtPatternIntervalExplicitMonthly.Size = New System.Drawing.Size(35, 22)
        Me.txtPatternIntervalExplicitMonthly.TabIndex = 3
        '
        'lblPatternInterval2Monthly
        '
        Me.lblPatternInterval2Monthly.AutoSize = True
        Me.lblPatternInterval2Monthly.Location = New System.Drawing.Point(206, 18)
        Me.lblPatternInterval2Monthly.Name = "lblPatternInterval2Monthly"
        Me.lblPatternInterval2Monthly.Size = New System.Drawing.Size(49, 14)
        Me.lblPatternInterval2Monthly.TabIndex = 4
        Me.lblPatternInterval2Monthly.Text = "month(s)"
        '
        'lblPatternInterval1Monthly
        '
        Me.lblPatternInterval1Monthly.AutoSize = True
        Me.lblPatternInterval1Monthly.Location = New System.Drawing.Point(106, 18)
        Me.lblPatternInterval1Monthly.Name = "lblPatternInterval1Monthly"
        Me.lblPatternInterval1Monthly.Size = New System.Drawing.Size(44, 14)
        Me.lblPatternInterval1Monthly.TabIndex = 2
        Me.lblPatternInterval1Monthly.Text = "of every"
        '
        'txtPatternDayOfMonthMonthly
        '
        Me.txtPatternDayOfMonthMonthly.Location = New System.Drawing.Point(58, 14)
        Me.txtPatternDayOfMonthMonthly.Name = "txtPatternDayOfMonthMonthly"
        Me.txtPatternDayOfMonthMonthly.Size = New System.Drawing.Size(35, 22)
        Me.txtPatternDayOfMonthMonthly.TabIndex = 1
        '
        'optPatternType_Monthly
        '
        Me.optPatternType_Monthly.BackColor = System.Drawing.Color.Transparent
        Me.optPatternType_Monthly.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem3.DataValue = "ValueListItem0"
        ValueListItem3.DisplayText = "Day"
        ValueListItem4.DataValue = "ValueListItem1"
        ValueListItem4.DisplayText = "The"
        Me.optPatternType_Monthly.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem3, ValueListItem4})
        Me.optPatternType_Monthly.ItemSpacingVertical = 16
        Me.optPatternType_Monthly.Location = New System.Drawing.Point(16, 12)
        Me.optPatternType_Monthly.Name = "optPatternType_Monthly"
        Me.optPatternType_Monthly.Size = New System.Drawing.Size(40, 54)
        Me.optPatternType_Monthly.TabIndex = 12
        '
        'cmdRemoveRecurrence
        '
        Appearance1.BackColor = System.Drawing.Color.FromArgb(CType(CType(222, Byte), Integer), CType(CType(223, Byte), Integer), CType(CType(206, Byte), Integer))
        Me.cmdRemoveRecurrence.Appearance = Appearance1
        Me.cmdRemoveRecurrence.Location = New System.Drawing.Point(292, 309)
        Me.cmdRemoveRecurrence.Name = "cmdRemoveRecurrence"
        Me.cmdRemoveRecurrence.Size = New System.Drawing.Size(154, 23)
        Me.cmdRemoveRecurrence.TabIndex = 16
        Me.cmdRemoveRecurrence.Text = "&Remove Recurrence"
        '
        'cmdCancel
        '
        Me.cmdCancel.Location = New System.Drawing.Point(371, 280)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
        Me.cmdCancel.TabIndex = 15
        Me.cmdCancel.Text = "Cancel"
        '
        'pnlYearly
        '
        Me.pnlYearly.Controls.Add(Me.cboMonthOfYearCalculatedYearly)
        Me.pnlYearly.Controls.Add(Me.lblCalculatedYearly)
        Me.pnlYearly.Controls.Add(Me.cboDayOfWeekYearly)
        Me.pnlYearly.Controls.Add(Me.cboOccurrenceOfDayInMonthYearly)
        Me.pnlYearly.Controls.Add(Me.txtPatternDayOfMonthYearly)
        Me.pnlYearly.Controls.Add(Me.cboMonthOfYearExplicitYearly)
        Me.pnlYearly.Controls.Add(Me.optPatternType_Yearly)
        Me.pnlYearly.Location = New System.Drawing.Point(488, 355)
        Me.pnlYearly.Name = "pnlYearly"
        Me.pnlYearly.Size = New System.Drawing.Size(350, 96)
        Me.pnlYearly.TabIndex = 20
        Me.pnlYearly.Visible = False
        '
        'cboMonthOfYearCalculatedYearly
        '
        Me.cboMonthOfYearCalculatedYearly.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cboMonthOfYearCalculatedYearly.Location = New System.Drawing.Point(252, 48)
        Me.cboMonthOfYearCalculatedYearly.MaxDropDownItems = 12
        Me.cboMonthOfYearCalculatedYearly.Name = "cboMonthOfYearCalculatedYearly"
        Me.cboMonthOfYearCalculatedYearly.Size = New System.Drawing.Size(88, 22)
        Me.cboMonthOfYearCalculatedYearly.TabIndex = 7
        '
        'lblCalculatedYearly
        '
        Me.lblCalculatedYearly.AutoSize = True
        Me.lblCalculatedYearly.Location = New System.Drawing.Point(231, 51)
        Me.lblCalculatedYearly.Name = "lblCalculatedYearly"
        Me.lblCalculatedYearly.Size = New System.Drawing.Size(14, 14)
        Me.lblCalculatedYearly.TabIndex = 6
        Me.lblCalculatedYearly.Text = "of"
        '
        'cboDayOfWeekYearly
        '
        Me.cboDayOfWeekYearly.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cboDayOfWeekYearly.Location = New System.Drawing.Point(142, 48)
        Me.cboDayOfWeekYearly.MaxDropDownItems = 10
        Me.cboDayOfWeekYearly.Name = "cboDayOfWeekYearly"
        Me.cboDayOfWeekYearly.Size = New System.Drawing.Size(86, 22)
        Me.cboDayOfWeekYearly.TabIndex = 5
        '
        'cboOccurrenceOfDayInMonthYearly
        '
        Me.cboOccurrenceOfDayInMonthYearly.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cboOccurrenceOfDayInMonthYearly.Location = New System.Drawing.Point(70, 48)
        Me.cboOccurrenceOfDayInMonthYearly.Name = "cboOccurrenceOfDayInMonthYearly"
        Me.cboOccurrenceOfDayInMonthYearly.Size = New System.Drawing.Size(70, 22)
        Me.cboOccurrenceOfDayInMonthYearly.TabIndex = 4
        '
        'txtPatternDayOfMonthYearly
        '
        Me.txtPatternDayOfMonthYearly.Location = New System.Drawing.Point(160, 14)
        Me.txtPatternDayOfMonthYearly.Name = "txtPatternDayOfMonthYearly"
        Me.txtPatternDayOfMonthYearly.Size = New System.Drawing.Size(35, 22)
        Me.txtPatternDayOfMonthYearly.TabIndex = 2
        '
        'cboMonthOfYearExplicitYearly
        '
        Me.cboMonthOfYearExplicitYearly.DropDownStyle = Infragistics.Win.DropDownStyle.DropDownList
        Me.cboMonthOfYearExplicitYearly.Location = New System.Drawing.Point(70, 14)
        Me.cboMonthOfYearExplicitYearly.MaxDropDownItems = 12
        Me.cboMonthOfYearExplicitYearly.Name = "cboMonthOfYearExplicitYearly"
        Me.cboMonthOfYearExplicitYearly.Size = New System.Drawing.Size(88, 22)
        Me.cboMonthOfYearExplicitYearly.TabIndex = 1
        '
        'optPatternType_Yearly
        '
        Me.optPatternType_Yearly.BackColor = System.Drawing.Color.Transparent
        Me.optPatternType_Yearly.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem5.DataValue = "ValueListItem0"
        ValueListItem5.DisplayText = "Every"
        ValueListItem6.DataValue = "ValueListItem1"
        ValueListItem6.DisplayText = "The"
        Me.optPatternType_Yearly.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem5, ValueListItem6})
        Me.optPatternType_Yearly.ItemSpacingVertical = 18
        Me.optPatternType_Yearly.Location = New System.Drawing.Point(8, 10)
        Me.optPatternType_Yearly.Name = "optPatternType_Yearly"
        Me.optPatternType_Yearly.Size = New System.Drawing.Size(56, 70)
        Me.optPatternType_Yearly.TabIndex = 13
        '
        'pnlWeekly
        '
        Me.pnlWeekly.Controls.Add(Me.chkSaturday)
        Me.pnlWeekly.Controls.Add(Me.chkFriday)
        Me.pnlWeekly.Controls.Add(Me.chkThursday)
        Me.pnlWeekly.Controls.Add(Me.chkWednesday)
        Me.pnlWeekly.Controls.Add(Me.chkTuesday)
        Me.pnlWeekly.Controls.Add(Me.chkMonday)
        Me.pnlWeekly.Controls.Add(Me.chkSunday)
        Me.pnlWeekly.Controls.Add(Me.lblPatternInterval2Weekly)
        Me.pnlWeekly.Controls.Add(Me.txtPatternIntervalWeekly)
        Me.pnlWeekly.Controls.Add(Me.lblPatternInterval1_Weekly)
        Me.pnlWeekly.Location = New System.Drawing.Point(488, 143)
        Me.pnlWeekly.Name = "pnlWeekly"
        Me.pnlWeekly.Size = New System.Drawing.Size(350, 96)
        Me.pnlWeekly.TabIndex = 17
        Me.pnlWeekly.Visible = False
        '
        'chkSaturday
        '
        Me.chkSaturday.Location = New System.Drawing.Point(176, 70)
        Me.chkSaturday.Name = "chkSaturday"
        Me.chkSaturday.Size = New System.Drawing.Size(72, 16)
        Me.chkSaturday.TabIndex = 9
        Me.chkSaturday.Text = "Saturday"
        '
        'chkFriday
        '
        Me.chkFriday.Location = New System.Drawing.Point(96, 70)
        Me.chkFriday.Name = "chkFriday"
        Me.chkFriday.Size = New System.Drawing.Size(72, 16)
        Me.chkFriday.TabIndex = 8
        Me.chkFriday.Text = "Friday"
        '
        'chkThursday
        '
        Me.chkThursday.Location = New System.Drawing.Point(16, 70)
        Me.chkThursday.Name = "chkThursday"
        Me.chkThursday.Size = New System.Drawing.Size(72, 16)
        Me.chkThursday.TabIndex = 7
        Me.chkThursday.Text = "Thursday"
        '
        'chkWednesday
        '
        Me.chkWednesday.Location = New System.Drawing.Point(253, 46)
        Me.chkWednesday.Name = "chkWednesday"
        Me.chkWednesday.Size = New System.Drawing.Size(88, 16)
        Me.chkWednesday.TabIndex = 6
        Me.chkWednesday.Text = "Wednesday"
        '
        'chkTuesday
        '
        Me.chkTuesday.Location = New System.Drawing.Point(176, 46)
        Me.chkTuesday.Name = "chkTuesday"
        Me.chkTuesday.Size = New System.Drawing.Size(72, 16)
        Me.chkTuesday.TabIndex = 5
        Me.chkTuesday.Text = "Tuesday"
        '
        'chkMonday
        '
        Me.chkMonday.Location = New System.Drawing.Point(96, 46)
        Me.chkMonday.Name = "chkMonday"
        Me.chkMonday.Size = New System.Drawing.Size(64, 16)
        Me.chkMonday.TabIndex = 4
        Me.chkMonday.Text = "Monday"
        '
        'chkSunday
        '
        Me.chkSunday.Location = New System.Drawing.Point(16, 46)
        Me.chkSunday.Name = "chkSunday"
        Me.chkSunday.Size = New System.Drawing.Size(64, 16)
        Me.chkSunday.TabIndex = 3
        Me.chkSunday.Text = "Sunday"
        '
        'lblPatternInterval2Weekly
        '
        Me.lblPatternInterval2Weekly.AutoSize = True
        Me.lblPatternInterval2Weekly.Location = New System.Drawing.Point(132, 12)
        Me.lblPatternInterval2Weekly.Name = "lblPatternInterval2Weekly"
        Me.lblPatternInterval2Weekly.Size = New System.Drawing.Size(63, 14)
        Me.lblPatternInterval2Weekly.TabIndex = 2
        Me.lblPatternInterval2Weekly.Text = "week(s) on:"
        '
        'txtPatternIntervalWeekly
        '
        Me.txtPatternIntervalWeekly.Location = New System.Drawing.Point(88, 10)
        Me.txtPatternIntervalWeekly.Name = "txtPatternIntervalWeekly"
        Me.txtPatternIntervalWeekly.Size = New System.Drawing.Size(35, 22)
        Me.txtPatternIntervalWeekly.TabIndex = 1
        '
        'lblPatternInterval1_Weekly
        '
        Me.lblPatternInterval1_Weekly.AutoSize = True
        Me.lblPatternInterval1_Weekly.Location = New System.Drawing.Point(16, 12)
        Me.lblPatternInterval1_Weekly.Name = "lblPatternInterval1_Weekly"
        Me.lblPatternInterval1_Weekly.Size = New System.Drawing.Size(65, 14)
        Me.lblPatternInterval1_Weekly.TabIndex = 0
        Me.lblPatternInterval1_Weekly.Text = "Re&cur every"
        '
        'fraAppointmentTime
        '
        Me.fraAppointmentTime.Controls.Add(Me.cboOccurrenceDuration)
        Me.fraAppointmentTime.Controls.Add(Me.lblOccurrenceDuration)
        Me.fraAppointmentTime.Controls.Add(Me.cboEndTime)
        Me.fraAppointmentTime.Controls.Add(Me.lblEndTime)
        Me.fraAppointmentTime.Controls.Add(Me.cboOccurrenceStartTime)
        Me.fraAppointmentTime.Controls.Add(Me.lblOccurrenceStartTime)
        Me.fraAppointmentTime.Location = New System.Drawing.Point(8, 8)
        Me.fraAppointmentTime.Name = "fraAppointmentTime"
        Me.fraAppointmentTime.Size = New System.Drawing.Size(438, 48)
        Me.fraAppointmentTime.TabIndex = 11
        Me.fraAppointmentTime.TabStop = False
        Me.fraAppointmentTime.Text = "Appointment time"
        '
        'cboOccurrenceDuration
        '
        Me.cboOccurrenceDuration.Location = New System.Drawing.Point(336, 16)
        Me.cboOccurrenceDuration.MaxDropDownItems = 7
        Me.cboOccurrenceDuration.Name = "cboOccurrenceDuration"
        Me.cboOccurrenceDuration.Size = New System.Drawing.Size(92, 22)
        Me.cboOccurrenceDuration.TabIndex = 5
        '
        'lblOccurrenceDuration
        '
        Me.lblOccurrenceDuration.AutoSize = True
        Me.lblOccurrenceDuration.Location = New System.Drawing.Point(284, 20)
        Me.lblOccurrenceDuration.Name = "lblOccurrenceDuration"
        Me.lblOccurrenceDuration.Size = New System.Drawing.Size(50, 14)
        Me.lblOccurrenceDuration.TabIndex = 4
        Me.lblOccurrenceDuration.Text = "D&uration:"
        '
        'cboEndTime
        '
        Me.cboEndTime.DropDownListWidth = 150
        Me.cboEndTime.Location = New System.Drawing.Point(176, 16)
        Me.cboEndTime.MaxDropDownItems = 7
        Me.cboEndTime.Name = "cboEndTime"
        Me.cboEndTime.Size = New System.Drawing.Size(92, 22)
        Me.cboEndTime.TabIndex = 3
        '
        'lblEndTime
        '
        Me.lblEndTime.AutoSize = True
        Me.lblEndTime.Location = New System.Drawing.Point(144, 20)
        Me.lblEndTime.Name = "lblEndTime"
        Me.lblEndTime.Size = New System.Drawing.Size(24, 14)
        Me.lblEndTime.TabIndex = 2
        Me.lblEndTime.Text = "E&nd"
        '
        'cboOccurrenceStartTime
        '
        Me.cboOccurrenceStartTime.Location = New System.Drawing.Point(44, 16)
        Me.cboOccurrenceStartTime.MaxDropDownItems = 7
        Me.cboOccurrenceStartTime.Name = "cboOccurrenceStartTime"
        Me.cboOccurrenceStartTime.Size = New System.Drawing.Size(92, 22)
        Me.cboOccurrenceStartTime.TabIndex = 1
        '
        'lblOccurrenceStartTime
        '
        Me.lblOccurrenceStartTime.AutoSize = True
        Me.lblOccurrenceStartTime.Location = New System.Drawing.Point(8, 20)
        Me.lblOccurrenceStartTime.Name = "lblOccurrenceStartTime"
        Me.lblOccurrenceStartTime.Size = New System.Drawing.Size(28, 14)
        Me.lblOccurrenceStartTime.TabIndex = 0
        Me.lblOccurrenceStartTime.Text = "S&tart"
        '
        'fraRecurrenceRange
        '
        Me.fraRecurrenceRange.Controls.Add(Me.dtpRangeEndDate)
        Me.fraRecurrenceRange.Controls.Add(Me.lblRangeMaxOccurrences)
        Me.fraRecurrenceRange.Controls.Add(Me.txtRangeMaxOccurrences)
        Me.fraRecurrenceRange.Controls.Add(Me.dtpRangeStartDate)
        Me.fraRecurrenceRange.Controls.Add(Me.lblRangeStartDate)
        Me.fraRecurrenceRange.Controls.Add(Me.optRangeLimit)
        Me.fraRecurrenceRange.Location = New System.Drawing.Point(8, 176)
        Me.fraRecurrenceRange.Name = "fraRecurrenceRange"
        Me.fraRecurrenceRange.Size = New System.Drawing.Size(438, 100)
        Me.fraRecurrenceRange.TabIndex = 13
        Me.fraRecurrenceRange.TabStop = False
        Me.fraRecurrenceRange.Text = "Range of recurrence"
        '
        'dtpRangeEndDate
        '
        Me.dtpRangeEndDate.BackColor = System.Drawing.SystemColors.Window
        Me.dtpRangeEndDate.DateButtons.Add(DateButton1)
        Me.dtpRangeEndDate.Format = "ddd M/d/yyyy"
        Me.dtpRangeEndDate.Location = New System.Drawing.Point(296, 67)
        Me.dtpRangeEndDate.Name = "dtpRangeEndDate"
        Me.dtpRangeEndDate.NonAutoSizeHeight = 21
        Me.dtpRangeEndDate.Size = New System.Drawing.Size(128, 21)
        Me.dtpRangeEndDate.TabIndex = 7
        '
        'lblRangeMaxOccurrences
        '
        Me.lblRangeMaxOccurrences.AutoSize = True
        Me.lblRangeMaxOccurrences.Location = New System.Drawing.Point(336, 47)
        Me.lblRangeMaxOccurrences.Name = "lblRangeMaxOccurrences"
        Me.lblRangeMaxOccurrences.Size = New System.Drawing.Size(66, 14)
        Me.lblRangeMaxOccurrences.TabIndex = 5
        Me.lblRangeMaxOccurrences.Text = "occurrences"
        '
        'txtRangeMaxOccurrences
        '
        Me.txtRangeMaxOccurrences.Location = New System.Drawing.Point(296, 43)
        Me.txtRangeMaxOccurrences.Name = "txtRangeMaxOccurrences"
        Me.txtRangeMaxOccurrences.Size = New System.Drawing.Size(35, 22)
        Me.txtRangeMaxOccurrences.TabIndex = 4
        '
        'dtpRangeStartDate
        '
        Me.dtpRangeStartDate.BackColor = System.Drawing.SystemColors.Window
        Me.dtpRangeStartDate.DateButtons.Add(DateButton2)
        Me.dtpRangeStartDate.Format = "ddd M/d/yyyy"
        Me.dtpRangeStartDate.Location = New System.Drawing.Point(48, 18)
        Me.dtpRangeStartDate.Name = "dtpRangeStartDate"
        Me.dtpRangeStartDate.NonAutoSizeHeight = 21
        Me.dtpRangeStartDate.Size = New System.Drawing.Size(152, 21)
        Me.dtpRangeStartDate.TabIndex = 1
        '
        'lblRangeStartDate
        '
        Me.lblRangeStartDate.AutoSize = True
        Me.lblRangeStartDate.Location = New System.Drawing.Point(8, 20)
        Me.lblRangeStartDate.Name = "lblRangeStartDate"
        Me.lblRangeStartDate.Size = New System.Drawing.Size(31, 14)
        Me.lblRangeStartDate.TabIndex = 0
        Me.lblRangeStartDate.Text = "&Start:"
        '
        'optRangeLimit
        '
        Me.optRangeLimit.BackColor = System.Drawing.Color.Transparent
        Me.optRangeLimit.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem7.DataValue = "ValueListItem0"
        ValueListItem7.DisplayText = "No End Date"
        ValueListItem8.DataValue = "ValueListItem1"
        ValueListItem8.DisplayText = "End after"
        ValueListItem9.DataValue = "ValueListItem2"
        ValueListItem9.DisplayText = "End by"
        Me.optRangeLimit.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem7, ValueListItem8, ValueListItem9})
        Me.optRangeLimit.ItemSpacingVertical = 6
        Me.optRangeLimit.Location = New System.Drawing.Point(208, 23)
        Me.optRangeLimit.Name = "optRangeLimit"
        Me.optRangeLimit.Size = New System.Drawing.Size(88, 64)
        Me.optRangeLimit.TabIndex = 11
        '
        'fraRecurrencePattern
        '
        Me.fraRecurrencePattern.Controls.Add(Me.optPatternFrequency)
        Me.fraRecurrencePattern.Controls.Add(Me.lblSepDark)
        Me.fraRecurrencePattern.Controls.Add(Me.lblSepLight)
        Me.fraRecurrencePattern.Location = New System.Drawing.Point(8, 64)
        Me.fraRecurrencePattern.Name = "fraRecurrencePattern"
        Me.fraRecurrencePattern.Size = New System.Drawing.Size(438, 116)
        Me.fraRecurrencePattern.TabIndex = 12
        Me.fraRecurrencePattern.TabStop = False
        Me.fraRecurrencePattern.Text = "Recurrence pattern"
        '
        'optPatternFrequency
        '
        Me.optPatternFrequency.BackColor = System.Drawing.Color.Transparent
        Me.optPatternFrequency.BorderStyle = Infragistics.Win.UIElementBorderStyle.None
        ValueListItem10.DataValue = "Daily"
        ValueListItem11.DataValue = "ValueListItem1"
        ValueListItem11.DisplayText = "Weekly"
        ValueListItem12.DataValue = "ValueListItem2"
        ValueListItem12.DisplayText = "Monthly"
        ValueListItem13.DataValue = "ValueListItem3"
        ValueListItem13.DisplayText = "Yearly"
        Me.optPatternFrequency.Items.AddRange(New Infragistics.Win.ValueListItem() {ValueListItem10, ValueListItem11, ValueListItem12, ValueListItem13})
        Me.optPatternFrequency.ItemSpacingVertical = 4
        Me.optPatternFrequency.Location = New System.Drawing.Point(8, 24)
        Me.optPatternFrequency.Name = "optPatternFrequency"
        Me.optPatternFrequency.Size = New System.Drawing.Size(64, 80)
        Me.optPatternFrequency.TabIndex = 12
        '
        'lblSepDark
        '
        Me.lblSepDark.BackColorInternal = System.Drawing.SystemColors.ControlDark
        Me.lblSepDark.Location = New System.Drawing.Point(72, 16)
        Me.lblSepDark.Name = "lblSepDark"
        Me.lblSepDark.Size = New System.Drawing.Size(1, 92)
        Me.lblSepDark.TabIndex = 5
        '
        'lblSepLight
        '
        Me.lblSepLight.BackColorInternal = System.Drawing.SystemColors.ControlLightLight
        Me.lblSepLight.Location = New System.Drawing.Point(73, 16)
        Me.lblSepLight.Name = "lblSepLight"
        Me.lblSepLight.Size = New System.Drawing.Size(1, 92)
        Me.lblSepLight.TabIndex = 4
        '
        'txtTickler
        '
        Me.txtTickler.Location = New System.Drawing.Point(56, 282)
        Me.txtTickler.Name = "txtTickler"
        Me.txtTickler.Size = New System.Drawing.Size(230, 22)
        Me.txtTickler.TabIndex = 21
        '
        'UltraLabel1
        '
        Me.UltraLabel1.AutoSize = True
        Me.UltraLabel1.Location = New System.Drawing.Point(8, 285)
        Me.UltraLabel1.Name = "UltraLabel1"
        Me.UltraLabel1.Size = New System.Drawing.Size(38, 14)
        Me.UltraLabel1.TabIndex = 22
        Me.UltraLabel1.Text = "Tickler"
        '
        'ASFAPPTR
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(848, 486)
        Me.Controls.Add(Me.UltraLabel1)
        Me.Controls.Add(Me.txtTickler)
        Me.Controls.Add(Me.cmdOk)
        Me.Controls.Add(Me.pnlDaily)
        Me.Controls.Add(Me.pnlMonthly)
        Me.Controls.Add(Me.cmdRemoveRecurrence)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.pnlYearly)
        Me.Controls.Add(Me.pnlWeekly)
        Me.Controls.Add(Me.fraAppointmentTime)
        Me.Controls.Add(Me.fraRecurrenceRange)
        Me.Controls.Add(Me.fraRecurrencePattern)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ASFAPPTR"
        Me.Text = "Appointment Recurrence"
        Me.pnlDaily.ResumeLayout(False)
        Me.pnlDaily.PerformLayout()
        CType(Me.txtPatternIntervalDaily, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optPatternDaysOfWeek_Daily, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMonthly.ResumeLayout(False)
        Me.pnlMonthly.PerformLayout()
        CType(Me.txtPatternIntervalCalculatedMonthly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboDayOfWeekMonthly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboOccurrenceOfDayInMonthMonthly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPatternIntervalExplicitMonthly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPatternDayOfMonthMonthly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optPatternType_Monthly, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlYearly.ResumeLayout(False)
        Me.pnlYearly.PerformLayout()
        CType(Me.cboMonthOfYearCalculatedYearly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboDayOfWeekYearly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboOccurrenceOfDayInMonthYearly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtPatternDayOfMonthYearly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboMonthOfYearExplicitYearly, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optPatternType_Yearly, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlWeekly.ResumeLayout(False)
        Me.pnlWeekly.PerformLayout()
        CType(Me.txtPatternIntervalWeekly, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraAppointmentTime.ResumeLayout(False)
        Me.fraAppointmentTime.PerformLayout()
        CType(Me.cboOccurrenceDuration, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboEndTime, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboOccurrenceStartTime, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraRecurrenceRange.ResumeLayout(False)
        Me.fraRecurrenceRange.PerformLayout()
        CType(Me.dtpRangeEndDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtRangeMaxOccurrences, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dtpRangeStartDate, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.optRangeLimit, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraRecurrencePattern.ResumeLayout(False)
        CType(Me.optPatternFrequency, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.txtTickler, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click

        If (Not Me.ActiveControl Is Nothing) Then
            Me.OnLeaveDialogControl(Me.ActiveControl, New EventArgs())
            If (Me.contentsAreValid = False) Then Return
        End If

        Me._result = RecurrenceDialogResult.Ok

        If (Me.PatternDayOfMonth > 28 AndAlso Me.patternDayOfMonthWarningDisplayed = False) Then
            MessageBox.Show(ASFAPPTR.RecurrenceDialog_Warning_PatternDayOfMonthExceeds28, ASFAPPTR.RecurrenceDialog_MessageBox_Caption, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

        Dim close As Boolean = True

        '	If the  recurrence we are modifying has variances, we must
        '	warn the end user about the potential loss of data
        '	and prompt as to whether to continue.
        If (Me.hasVariances AndAlso Not Utilities.AreRecurrencesEqual(Me.initialAppointmentRecurrence, Me.appointmentRecurrence)) Then

            Dim result As DialogResult = MessageBox.Show(ASFAPPTR.RecurrenceDialog_Prompt_OkToRemoveVariances, _
                  ASFAPPTR.RecurrenceDialog_MessageBox_Caption, _
                  MessageBoxButtons.OKCancel, _
                  MessageBoxIcon.Warning)

            If (result = Windows.Forms.DialogResult.Cancel) Then

                Me._result = RecurrenceDialogResult.Cancel
                close = False

            End If
        End If

        _Tickler = txtTickler.Text

        If close Then Me.Close()
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.forceContentsChanged = False
        Me.Close()
    End Sub

    Private Sub cmdRemoveRecurrence_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemoveRecurrence.Click
        Me._result = RecurrenceDialogResult.RemoveRecurrence
        Me.Close()
    End Sub

    Private Sub fraRecurrencePattern_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fraRecurrencePattern.Enter

    End Sub

    Private Sub ASFAPPTR_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If _Caption <> "" Then
            Me.Text = _Caption
        End If
        txtTickler.Text = _Tickler
    End Sub
End Class

Friend Class Utilities


#Region "TimeFormatString"
#If Debug Then
    ' <summary>
    ' Returns the time format string from the system's LongTimePattern, stripped of the seconds component.
    ' </summary>
#End If
    Friend Shared ReadOnly Property TimeFormatString() As String

        Get

            Dim _formatString As String = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.LongTimePattern

            _formatString = _formatString.Replace(":s", String.Empty)
            _formatString = _formatString.Replace("s", String.Empty)

            Return _formatString
        End Get
    End Property

#End Region   ' TimeFormatString

#Region "GetTotalNumberOfDaysInMonth"
#If Debug Then
    ' <summary>
    ' Returns the total number of days in the month corresponding to the specified date.
    ' </summary>
    ' <param name="theDate">The DateTime containing the month to test.</param>
#End If
    Friend Shared Function GetTotalNumberOfDaysInMonth(ByVal theDate As DateTime) As Int32

        Dim cal As System.Globalization.GregorianCalendar = New System.Globalization.GregorianCalendar()
        Return cal.GetDaysInMonth(thedate.Year, thedate.Month)

    End Function
#End Region   ' GetTotalNumberOfDaysInMonth

#Region "DaysOfWeek"
#If Debug Then
    ' <summary>
    ' Returns a string array containing the names of the days of the week for the current culture.
    ' </summary>
#End If
    Friend Shared ReadOnly Property DaysOfWeek() As String()

        Get
            Return System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames
        End Get

    End Property

#End Region   ' DaysOfWeek

#Region "MonthsOfYear"
#If Debug Then
    ' <summary>
    ' Returs a string array containing the names of the months of the year for the current culture.
    ' </summary>
#End If
    Friend Shared ReadOnly Property MonthsOfYear() As String()

        Get

            Dim retVal(11) As String
            Dim monthNames() As String = System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.MonthNames
            Dim i As Int32
            For i = 0 To 11
                retVal(i) = monthNames(i)
            Next

            Return retVal

        End Get

    End Property

#End Region   ' MonthsOfYear

#Region "FormatTimeSpan"
#If Debug Then
    ' <summary>
    ' Returns the format string to be used for dates.
    ' </summary>
#End If
    Friend Shared Function FormatTimeSpan(ByVal timeSpan As TimeSpan, ByVal encloseInParentheses As Boolean) As String

        Dim totalMinutes As Double = timeSpan.TotalMinutes
        Dim totalHours As Double = Math.Abs(timeSpan.TotalHours)
        Dim totalDays As Double = Math.Abs(timeSpan.TotalDays)
        Dim totalWeeks As Double = Math.Abs(totalDays / 7.0)
        Dim totalMonths As Double = Math.Abs(totalDays / 30.0)

        Dim retVal As String = String.Empty

        If (totalMinutes = 0.0F) Then retVal = "0 minutes"


        Dim formatSpecifier As String = "f0"
        If (Utilities.IsFormattable(totalMonths, formatSpecifier)) Then

            retVal = totalMonths.ToString(formatSpecifier) + " month"
            If (totalMonths > 1.0F) Then retVal += "s"
        ElseIf (Utilities.IsFormattable(totalWeeks, formatSpecifier)) Then

            retVal = totalWeeks.ToString(formatSpecifier) + " week"
            If (totalWeeks > 1.0F) Then retVal += "s"

        ElseIf (Utilities.IsFormattable(totalDays, formatSpecifier)) Then

            retVal = totalDays.ToString(formatSpecifier) + " day"
            If (totalDays > 1.0F) Then retVal += "s"

        ElseIf (Utilities.IsFormattable(totalHours, formatSpecifier)) Then

            retVal = totalHours.ToString(formatSpecifier) + " hour"
            If (totalHours > 1.0F) Then retVal += "s"

        Else

            retVal = totalMinutes.ToString("f0") + " minute"
            If (totalMinutes > 1.0F) Then retVal += "s"
        End If

        If (encloseInParentheses) Then retVal = "(" + retVal + ")"

        Return retVal

    End Function

#End Region   ' FormatTimeSpan

#Region "IsFormattable"
#If Debug Then
    ' <summary>
    ' Returns whether the specified value is formattable as either
    ' a whole number, x.25, x.5, or x.75. Also returns the appropriate
    ' format specifier if the method returns true
    ' </summary>
    ' <param name="value">The double-precision value to test</param>
    ' <param name="formatSpecifier">[out] If the method returns true, contains the appropriate format specifier.</param>
    ' <returns></returns>
#End If
    Private Shared Function IsFormattable(ByVal value As Double, ByRef formatSpecifier As String) As Boolean

        formatSpecifier = "f0"
        Dim retVal As Boolean = False

        If (value < 1.0F) Then Return False

        If ((value Mod 1.0F) = 0.0F) Then

            retVal = True
            formatSpecifier = "f0"

        ElseIf ((value Mod 0.5F) = 0.0F) Then

            retVal = True
            formatSpecifier = "f1"

        ElseIf ((value Mod 0.25F) = 0.0F Or (value Mod 0.75F) = 0.0F) Then

            retVal = True
            formatSpecifier = "f2"

        End If

        Return retVal

    End Function

#End Region   ' IsFormattable

#Region "ParseTimeSpan"
#If Debug Then
    ' <summary>
    ' Returns true if the specified string could be successfully parsed
    ' into a TimeSpan otherwise, returns false. If successful, the 'timeSpan'
    ' parameter contains a valid value.
    ' </summary>
#End If
    Friend Shared Function ParseTimeSpan(ByVal value As String, ByRef timeSpan As TimeSpan) As Boolean

        timeSpan = timeSpan.Zero

        If (value Is Nothing) Then Return False

        If (value.IndexOf("minute") > -1 Or _
          value.IndexOf("hour") > -1 Or _
          value.IndexOf("day") > -1 Or _
          value.IndexOf("week") > -1 Or _
          value.IndexOf("month") > -1 Or _
          value.IndexOf("year") > -1) Then

            Dim temp As String = value
            Dim unitIsMinute As Boolean = False
            Dim unitIsHour As Boolean = False
            Dim unitIsDay As Boolean = False
            Dim unitIsWeek As Boolean = False
            Dim unitIsMonth As Boolean = False
            Dim unitIsYear As Boolean = False

            If (value.IndexOf("minute") > -1) Then

                temp = temp.Replace("minute", String.Empty)
                unitIsMinute = True

            ElseIf (value.IndexOf("hour") > -1) Then

                temp = temp.Replace("hour", String.Empty)
                unitIsHour = True

            ElseIf (value.IndexOf("day") > -1) Then

                temp = temp.Replace("day", String.Empty)
                unitIsDay = True

            ElseIf (value.IndexOf("week") > -1) Then

                temp = temp.Replace("week", String.Empty)
                unitIsWeek = True

            ElseIf (value.IndexOf("month") > -1) Then

                temp = temp.Replace("month", String.Empty)
                unitIsMonth = True

            ElseIf (value.IndexOf("year") > -1) Then

                temp = temp.Replace("year", String.Empty)
                unitIsYear = True

            End If

            '	Lose the pluralizer, and any spaces
            temp = temp.Replace("s", String.Empty)
            temp = temp.Replace(" ", String.Empty)

            Try

                Dim theNumber As Double = Convert.ToDouble(temp)
                If (unitIsMinute) Then
                    timeSpan = System.TimeSpan.FromMinutes(theNumber)
                ElseIf (unitIsHour) Then
                    timeSpan = System.TimeSpan.FromHours(theNumber)
                ElseIf (unitIsDay) Then
                    timeSpan = System.TimeSpan.FromDays(theNumber)
                ElseIf (unitIsWeek) Then
                    timeSpan = System.TimeSpan.FromDays(theNumber * 7.0)
                ElseIf (unitIsMonth) Then
                    timeSpan = System.TimeSpan.FromDays(theNumber * 30.0)
                ElseIf (unitIsYear) Then
                    timeSpan = System.TimeSpan.FromDays(theNumber * 365.0)
                End If

                Return True

            Catch
                Return False
            End Try

        End If

        Return False

    End Function

#End Region   ' ParseTimeSpan

#Region "GetOccurrenceOfDayInMonth"
#If Debug Then
    ' <summary>
    ' Returns a RecurrencePatternOccurrenceOfDayInMonth value from the specified DateTime.
    ' </summary>
    ' <param name="date">The date for which to obtain the RecurrencePatternOccurrenceOfDayInMonth value.</param>
    ' <returns>A RecurrencePatternDaysOfWeek value that matches the specified System.DayOfWeek</returns>
#End If
    Friend Shared Function GetOccurrenceOfDayInMonth(ByVal theDate As DateTime) As RecurrencePatternOccurrenceOfDayInMonth

        '	Get the first day of the month that was specified
        Dim startDate As DateTime = New DateTime(theDate.Year, theDate.Month, 1, 0, 0, 0)

        '	Move to the first occurrence of the specified day of week
        Dim i As Int32
        For i = 1 To 7
            If (startDate.DayOfWeek = theDate.DayOfWeek) Then Exit For
            startDate = startDate.AddDays(1.0F)
        Next

        If (startDate.Date = theDate.Date) Then Return RecurrencePatternOccurrenceOfDayInMonth.First

        Dim temp As DateTime = startDate.Date.AddDays(7.0F)
        If (temp.Month <> theDate.Month) Then
            Return RecurrencePatternOccurrenceOfDayInMonth.Last
        ElseIf (temp = theDate.Date) Then
            Return RecurrencePatternOccurrenceOfDayInMonth.Second
        End If

        temp = startDate.Date.AddDays(14.0F)
        If (temp.Month <> theDate.Month) Then
            Return RecurrencePatternOccurrenceOfDayInMonth.Last
        ElseIf (temp = theDate.Date) Then
            Return RecurrencePatternOccurrenceOfDayInMonth.Third
        End If

        temp = startDate.Date.AddDays(21.0F)
        If (temp.Month <> theDate.Month) Then
            Return RecurrencePatternOccurrenceOfDayInMonth.Last
        Else
            If (temp = theDate.Date) Then Return RecurrencePatternOccurrenceOfDayInMonth.Fourth
        End If

        Return RecurrencePatternOccurrenceOfDayInMonth.Last

    End Function

#End Region   ' GetOccurrenceOfDayInMonth

#Region "GetShortestDurationFromDaysOfWeek"
#If Debug Then
    ' <summary>
    ' Returns the shortest duration between adjacent occurrences in the specified RecurrencePatternDaysOfWeek.
    ' </summary>
    ' <param name="daysOfWeek"></param>
    ' <returns>A TimeSpan representing the shortest duration between any 2 occurrences.</returns>
#End If
    Friend Shared Function GetShortestDurationFromDaysOfWeek(ByVal daysOfWeek As RecurrencePatternDaysOfWeek) As TimeSpan

        '	Initialize an array of 14 booleans to false
        Dim days(14) As Boolean
        Dim i As Int32
        For i = 0 To 13
            days(i) = False
        Next

        '	Based on which bits are set in the argument, set
        '	the corresponding member of the array to true
        For i = 0 To 6

            Dim bitMask As Int32 = System.Math.Pow(2.0F, i)
            days(i) = (daysOfWeek And bitMask) = bitMask
        Next

        '	Duplicate the first 7 members of the array...this is
        '	to get a sample that is 2 weeks in duration, so we account
        '	for wrapping around to the following week
        For i = 7 To 13

            days(i) = days(i - 7)
        Next

        '	Iterate the array and get the 2 members that are "closest together"
        Dim lastDaySet As Int32 = -1
        Dim shortestDuration As Int32 = 7

        For i = 0 To 13

            '	If not set, skip this iteration
            If days(i) Then

                '	Set the initial value of the 'lastDaySet' stack variable
                If (days(i) AndAlso lastDaySet = -1) Then
                    lastDaySet = i
                Else

                    '	Each time we hit a set bit in the mask, see how far away
                    '	it is from the previous one, and update the 'shortestDuration'
                    '	stack variable if this pair is closer together than the last.
                    If (days(i) AndAlso lastDaySet >= 0) Then

                        Dim duration As Int32 = (i - lastDaySet)
                        shortestDuration = System.Math.Min(shortestDuration, duration)
                        lastDaySet = i
                    End If

                End If

            End If

        Next

        '	Return a TimeSpan that represents the number of days
        '	between the 2 closest occurrences.
        Return New TimeSpan(shortestDuration, 0, 0, 0)

    End Function

#End Region ' GetShortestDurationFromDaysOfWeek

#Region "GetRecurrencePatternDaysOfWeekFromDayOfWeek"
#If Debug Then
    ' <summary>
    ' Returns the RecurrencePatternDaysOfWeek constant that corresponds to
    ' the specified System.DayOfWeek constant.
    ' </summary>
    ' <param name="dayOfWeek">The day of the week to return.</param>
    ' <returns>A RecurrencePatternDaysOfWeek value that matches the specified System.DayOfWeek</returns>
#End If
    Friend Shared Function GetRecurrencePatternDaysOfWeekFromDayOfWeek(ByVal dayOfWeek As System.DayOfWeek) As RecurrencePatternDaysOfWeek

        Dim dayInWeek As Int32 = dayOfWeek
        Dim retVal As Int32 = System.Math.Pow(2.0F, dayInWeek)
        Return retVal

    End Function
#End Region   ' GetRecurrencePatternDaysOfWeekFromDayOfWeek

#Region "AreRecurrencesEqual"
#If Debug Then
    ' <summary>
    ' Returns whether the property values of 'recurrence1'
    ' are the same as those of 'recurrence2'. Note that the
    ' RootAppointment and contents of the Variances collections
    ' are not considered.
    ' </summary>
    ' <returns>A boolean indicating whether the 2 instances are logically the same.</returns>
#End If
    Friend Shared Function AreRecurrencesEqual(ByVal recurrence1 As AppointmentRecurrence, ByVal recurrence2 As AppointmentRecurrence) As Boolean

        If (recurrence1 Is Nothing AndAlso recurrence2 Is Nothing) Then Return True

        If ((recurrence1 Is Nothing AndAlso Not recurrence2 Is Nothing) Or _
          (recurrence2 Is Nothing AndAlso Not recurrence1 Is Nothing)) Then
            Return False
        End If

        Return recurrence1.OccurrenceStartTime = recurrence2.OccurrenceStartTime AndAlso _
          recurrence1.OccurrenceDuration.TotalMinutes = recurrence2.OccurrenceDuration.TotalMinutes AndAlso _
          recurrence1.PatternDayOfMonth = recurrence2.PatternDayOfMonth AndAlso _
          recurrence1.PatternDaysOfWeek = recurrence2.PatternDaysOfWeek AndAlso _
          recurrence1.PatternFrequency = recurrence2.PatternFrequency AndAlso _
          recurrence1.PatternInterval = recurrence2.PatternInterval AndAlso _
          recurrence1.PatternMonthOfYear = recurrence2.PatternMonthOfYear AndAlso _
          recurrence1.PatternOccurrenceOfDayInMonth = recurrence2.PatternOccurrenceOfDayInMonth AndAlso _
          recurrence1.PatternType = recurrence2.PatternType AndAlso _
          recurrence1.RangeEndDate = recurrence2.RangeEndDate AndAlso _
          recurrence1.RangeLimit = recurrence2.RangeLimit AndAlso _
          recurrence1.RangeMaxOccurrences = recurrence2.RangeMaxOccurrences AndAlso _
          recurrence1.RangeStartDate = recurrence2.RangeStartDate

    End Function
#End Region   ' AreRecurrencesEqual




End Class