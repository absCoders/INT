Public Class TAFADDR1

    'Public Structure AddressValidationResponse
    '    Dim AddressIndex As Integer
    '    Dim ClassificationCode As AddressClassificationTypes
    '    Dim ClassificationDescription As String
    '    Dim Consignee As String
    '    Dim BuildingName As String
    '    Dim AddressLine1 As String
    '    Dim AddressLine2 As String
    '    Dim AddressLine3 As String
    '    Dim City As String
    '    Dim State As String
    '    Dim PostalCode As String
    '    Dim PostaclCodeExtended As String
    '    Dim CountryCode As String
    'End Structure

    Public frmAddressMatches As List(Of SHCUPSC1.AddressValidationResponse)

    Public Sub New(ByRef AddressMatches As List(Of SHCUPSC1.AddressValidationResponse))

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        frmAddressMatches = AddressMatches
    End Sub

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim tblAddresses As DataTable = Nothing
        Dim rowAddresses As DataRow

        tblAddresses = New DataTable

        tblAddresses.Columns.Add("AddressIndex", GetType(System.Int32))
        tblAddresses.Columns.Add("Consignee", GetType(System.String))
        tblAddresses.Columns.Add("BuildingName", GetType(System.String))
        tblAddresses.Columns.Add("AddressLine1", GetType(System.String))
        tblAddresses.Columns.Add("AddressLine2", GetType(System.String))
        tblAddresses.Columns.Add("AddressLine3", GetType(System.String))
        tblAddresses.Columns.Add("City", GetType(System.String))
        tblAddresses.Columns.Add("State", GetType(System.String))
        tblAddresses.Columns.Add("PostalCode", GetType(System.String))
        tblAddresses.Columns.Add("PostalCodeExtended", GetType(System.String))
        tblAddresses.Columns.Add("CountryCode", GetType(System.String))

        For Each address As SHCUPSC1.AddressValidationResponse In frmAddressMatches
            address.Selected = False
            rowAddresses = tblAddresses.NewRow
            rowAddresses.Item("AddressIndex") = address.AddressIndex
            rowAddresses.Item("Consignee") = address.Consignee
            rowAddresses.Item("BuildingName") = address.BuildingName
            rowAddresses.Item("AddressLine1") = address.AddressLine1
            rowAddresses.Item("AddressLine2") = address.AddressLine2
            rowAddresses.Item("AddressLine3") = address.AddressLine3
            rowAddresses.Item("City") = address.City
            rowAddresses.Item("State") = address.State
            rowAddresses.Item("PostalCode") = address.PostalCode
            rowAddresses.Item("PostalCodeExtended") = address.PostalCodeExtended
            rowAddresses.Item("CountryCode") = address.CountryCode
            tblAddresses.Rows.Add(rowAddresses)
        Next

        Me.grdADDRESS.DataSource = tblAddresses
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub

    Private Sub cmdSelect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSelect.Click

        If Me.grdADDRESS.Selected.Rows.Count = 0 Then
            MessageBox.Show("You must select an address or click 'Cancel'.", "Select", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim addressIndex As Integer = Me.grdADDRESS.Selected.Rows(0).Cells("AddressIndex").Value
        Dim addr As SHCUPSC1.AddressValidationResponse = frmAddressMatches(addressIndex - 1)
        addr.Selected = True
        frmAddressMatches(addressIndex - 1) = addr

        Me.Close()
    End Sub

    Private Sub grdADDRESS_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdADDRESS.DoubleClickRow
        If grdADDRESS.Selected.Rows.Count = 1 Then
            Me.cmdSelect_Click(Nothing, Nothing)
        End If
    End Sub

End Class