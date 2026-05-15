Imports Infragistics.Win

Public Class CreditCardSubmission

    Private queryGridControl As String = String.Empty
    Private sCustomerCode As String = String.Empty
    Private gridInitialized As Boolean = False
    Private AbsCon As Object = Nothing 'New ABSConnector

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        queryGridControl = "Select ARTCCPA1.* from ARTCCPA1 where ARTCCPA1.CUST_CODE = :PARM1 and CCPA_STATUS <> '0' and CCPA_STATUS <> 'D' and CCPA_STATUS <> 'X'"
        sCustomerCode = String.Empty
        gridInitialized = False

    End Sub

#Region "Control Properties"

    ''' <summary>
    ''' Customer Code to get CC data for
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property CustomerCode() As String
        Get
            Return sCustomerCode
        End Get
        Set(ByVal value As String)
            sCustomerCode = value
        End Set
    End Property

    ''' <summary>
    ''' Gives access to the User Controls Grid Control
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property UserControlGrid() As Infragistics.Win.UltraWinGrid.UltraGrid
        Get
            Return grdControl
        End Get

    End Property

#End Region

#Region "Control Public Methods"

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <param name="displayCustomerCode">Customer Code used to fill query</param>
    ''' <remarks></remarks>
    Public Sub DisplayData(ByVal displayCustomerCode As String)

        Try
            If AbsCon Is Nothing Then
                AbsCon = New ABSConnector
            End If

            sCustomerCode = displayCustomerCode

            Dim tblARTCCPAP As DataTable = AbsCon.GetDataTable(queryGridControl, "", "V", New Object() {displayCustomerCode})
            grdControl.DataSource = tblARTCCPAP

            If Not gridInitialized Then
                AbsCon.Create_Summary(grdControl, "CUST_CREDIT_CARD_LAST4", "Count")
                AbsCon.Create_Summary(grdControl, "CCPA_AMT")

                AbsCon.Add_Value_List(grdControl, "CCPA_STATUS")
                AbsCon.Add_Value_List(grdControl, "CCPA_REASON")
                AbsCon.Add_Value_List(grdControl, "RESPONSE_CODE")
                AbsCon.Add_Value_List(grdControl, "CCPA_TYPE")
            End If
            gridInitialized = True

            AbsCon.Sort_grdColumns(grdControl, "LAST_DATE".ToLower)
            grdControl.Text = "Credit Card Submission History for " & displayCustomerCode
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK)
        End Try

    End Sub

    ''' <summary>
    ''' Fills the grid with the data for the provided Customer Code
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub DisplayData()
        DisplayData(sCustomerCode)
    End Sub

#End Region

#Region "User Control Controls"

    Private Sub grdControl_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdControl.InitializeRow

        If e.Row.Cells("RESPONSE_CODE").Text = "A" Then
            e.Row.Cells("RESPONSE_TEXT").Appearance.ForeColor = Drawing.Color.Green
        ElseIf e.Row.Cells("RESPONSE_CODE").Text = "E" Then
            e.Row.Cells("RESPONSE_TEXT").Appearance.ForeColor = Drawing.Color.Red
        Else
            e.Row.Cells("RESPONSE_TEXT").Appearance.ForeColor = Drawing.Color.Empty
        End If

        If e.Row.Cells("CCPA_STATUS").Value = "A" Then
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.Green
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "E" Then
            e.Row.Cells("CCPA_STATUS").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.White
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "D" Then
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.Red
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "S" Then
            e.Row.Cells("CCPA_STATUS").Appearance.BackColor = Drawing.Color.Green
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.White
        ElseIf e.Row.Cells("CCPA_STATUS").Value = "V" Then
            e.Row.Cells("CCPA_STATUS").Appearance.BackColor = Drawing.Color.Red
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.White
        Else
            e.Row.Cells("CCPA_STATUS").Appearance.ForeColor = Drawing.Color.Empty
        End If

    End Sub

#End Region

End Class
