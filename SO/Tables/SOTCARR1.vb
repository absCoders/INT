Imports Infragistics.Win.UltraWinGrid

Public Class SOTCARR1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Create_TDA(.Tables.Add, "SOTCARR2", "*", 1)
            Create_TDA(.Tables.Add, "SOTCARR3", "*", 1)
            Create_TDA(.Tables.Add, "SOTCARR4", "*", 1)

            .Tables("SOTCARR3").Columns.Add("CUST_NAME", GetType(String))
            .Tables("SOTCARR3").Columns.Add("CARRIER_PROD_DESC", GetType(String), "CARRIER_PROD_CODE")
        End With

        grdSOTCARR2.DataSource = dst.Tables("SOTCARR2")
        grdSOTCARR3.DataSource = dst.Tables("SOTCARR3")
        grdSOTCARR4.DataSource = dst.Tables("SOTCARR4")

        ASCMAIN1.Add_Value_List(grdSOTCARR2, "SERVICE_CODE", Nothing, New String() {":", "D:Domestic", "I:International"}, 0)
        ASCMAIN1.Add_Value_List(grdSOTCARR2, "TRACKING_ID_TYPE", Nothing, New String() {":", "0:Fedex Express", "1:Fedex Ground", "2:USPS", "3:N/A"}, 0)

        Create_Lookup("ARTCUST1")
        Create_Lookup("GLTACCT1")
        Create_Lookup("GLTSEGM1")

    End Sub

#Region "Overrides"

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey
            Case "Update"

                grdSOTCARR2.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdSOTCARR3.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)
                grdSOTCARR4.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

                Validate_Code("CARRIER_TYPE")

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
        Update_Record_TDA("SOTCARR2", "CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
        Update_Record_TDA("SOTCARR3", "CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
        Update_Record_TDA("SOTCARR4", "CARRIER_CODE = '" & Absx1.txtFor("CARRIER_CODE").Text & "'")
    End Sub

    Overrides Sub Proceed_Update_Special_Post()

    End Sub

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
        End If

        EnforceConstraints(False)
        Fill_Records("SOTCARR2", New String() {Absx1.txtFor("CARRIER_CODE").Text})
        Fill_Records("SOTCARR3", New String() {Absx1.txtFor("CARRIER_CODE").Text})
        Fill_Records("SOTCARR4", New String() {Absx1.txtFor("CARRIER_CODE").Text})
        EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            dst.EnforceConstraints = False
            dst.Tables("SOTCARR2").Rows.Clear()
            dst.Tables("SOTCARR3").Rows.Clear()
            dst.Tables("SOTCARR4").Rows.Clear()
            dst.EnforceConstraints = False
        End If
    End Sub

    Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)
        tabOther.Enabled = tf
    End Sub

    Public Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)
    End Sub

#End Region

    Private Sub grdSOTCARR4_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCARR4.BeforeRowUpdate
        e.Row.Cells("CARRIER_CODE").Value = Absx1.txtFor("CARRIER_CODE").Text

        e.Row.Cells("PACKAGE_CODE").Value = (e.Row.Cells("PACKAGE_CODE").Value & String.Empty).ToString.Trim
        e.Row.Cells("PACKAGE_DESC").Value = (e.Row.Cells("PACKAGE_DESC").Value & String.Empty).ToString.Trim

        If e.Row.Cells("PACKAGE_CODE").Value.ToString.Length = 0 OrElse
            e.Row.Cells("PACKAGE_DESC").Value.ToString.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Package Code and Description are required.")
        End If
    End Sub

    Private Sub grdSOTCARR3_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCARR3.BeforeRowUpdate

        e.Row.Cells("CARRIER_CODE").Value = Absx1.txtFor("CARRIER_CODE").Text
        Dim CUST_CODE As String = e.Row.Cells("DIVISION_CODE").Value & String.Empty

        Select Case CUST_CODE
            Case ASCMAIN1.DBS_COMPANY
                ' This is good

            Case String.Empty
                MessageBox.Show("Customer Code is required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                e.Cancel = True
            Case Else
                Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1 Is Nothing Then
                    MessageBox.Show("Invalid Customer Code", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                    Exit Sub
                End If
                e.Row.Cells("CUST_NAME").Value = rowARTCUST1.Item("CUST_NAME") & String.Empty
        End Select

        Dim CARRIER_ACCOUNT_NO As String = e.Row.Cells("CARRIER_ACCOUNT_NO").Value & String.Empty
        If CARRIER_ACCOUNT_NO.Length = 0 Then
            MessageBox.Show("Account Number is Required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        ' SHIPPER_PASSWORD
        Dim SHIPPER_ID As String = e.Row.Cells("SHIPPER_ID").Value & String.Empty
        If SHIPPER_ID.Length = 0 Then
            MessageBox.Show("Shipper ID is Required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        Dim SHIPPER_PASSWORD As String = e.Row.Cells("SHIPPER_PASSWORD").Value & String.Empty
        If SHIPPER_PASSWORD.Length = 0 Then
            MessageBox.Show("Password is Required", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        ' CARRIER_PROD_CODE
        Dim CARRIER_PROD_CODE As String = e.Row.Cells("CARRIER_PROD_CODE").Value & String.Empty
        If CARRIER_PROD_CODE.Length = 0 Then
            MessageBox.Show("Carrier Prod Code is Required, Use an Asterick (*) for all carrier shipping methods", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            e.Cancel = True
        End If

        Select Case CARRIER_PROD_CODE
            Case "*"
                ' This is good. Menas all shipping methods
            Case Else
                Dim rowSOTCARR2 As DataRow = LookUp("SOTCARR2", {Absx1.txtFor("CARRIER_CODE").Text, CARRIER_PROD_CODE})
                If rowSOTCARR2 Is Nothing Then
                    MessageBox.Show("Invalid Carrier Prod Code", "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    e.Cancel = True
                End If
        End Select


        'Dim ACCT_CODE As String = e.Row.Cells("ACCT_CODE").Value.ToString.Trim
        'Dim SEG2_CODE As String = e.Row.Cells("SEG2_CODE").Value.ToString.Trim
        'Dim SEG3_CODE As String = e.Row.Cells("SEG3_CODE").Value.ToString.Trim
        'Dim SEG4_CODE As String = e.Row.Cells("SEG4_CODE").Value.ToString.Trim

        'If ACCT_CODE.Length > 0 Then
        '    If LookUp("GLTACCT1", ACCT_CODE) Is Nothing Then
        '        MessageBox.Show("Invalid entry for account Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If
        'End If

        'If ACCT_CODE.Length > 0 Then
        '    If LookUp("GLTSEGM1", New String() {SEG2_CODE, "2"}) Is Nothing Then
        '        MessageBox.Show("Invalid entry for Seg 2 Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If

        '    If LookUp("GLTSEGM1", New String() {SEG3_CODE, "3"}) Is Nothing Then
        '        MessageBox.Show("Invalid entry for Seg 3 Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If

        '    If LookUp("GLTSEGM1", New String() {SEG4_CODE, "4"}) Is Nothing Then
        '        MessageBox.Show("Invalid entry for Seg 4 Code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '        e.Cancel = True
        '        Exit Sub
        '    End If

        'ElseIf SEG2_CODE.Length > 0 OrElse SEG3_CODE.Length > 0 OrElse SEG4_CODE.Length > 0 Then
        '    If MessageBox.Show("There are account segment values with no Account Code. Do you want to clear the segment codes?", "Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.No Then
        '        e.Cancel = True
        '        Exit Sub
        '    End If
        '    e.Row.Cells("SEG2_CODE").Value = String.Empty
        '    e.Row.Cells("SEG3_CODE").Value = String.Empty
        '    e.Row.Cells("SEG4_CODE").Value = String.Empty
        'End If
    End Sub

    Private Sub grdSOTCARR2_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTCARR2.BeforeRowUpdate
        e.Row.Cells("CARRIER_CODE").Value = MyBase.Absx1.txtFor("CARRIER_CODE").Text

        If (e.Row.Cells("CARRIER_PROD_DESC").Value).ToString.Trim.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Product Description is required.", "Update Error", MessageBoxButtons.OK)
            Exit Sub
        End If
    End Sub

    Private Sub grdSOTCARR3_ClickCellButton(sender As Object, e As CellEventArgs) Handles grdSOTCARR3.ClickCellButton

        Select Case e.Cell.Column.Key
            Case "DIVISION_CODE"
                Call grdClickCellButton(grdSOTCARR3, String.Empty, False, e.Cell.Column.Key, "CUST_CODE")
        End Select
    End Sub
End Class