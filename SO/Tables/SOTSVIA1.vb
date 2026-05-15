Public Class SOTSVIA1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        With dst
            Call Create_TDA(.Tables.Add, "SOTSVIAA", "*", 1)
        End With

        grdSOTSVIAA.DataSource = dst.Tables("SOTSVIAA")

        Create_Lookup("SOTSVIA1")

    End Sub

    Public Overrides Sub Prepare_for_View_Lookup_Special(ctl As System.Windows.Forms.Control, COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        MyBase.Prepare_for_View_Lookup_Special(ctl, COLUMN_NAME, sql_where, Cancel)

        Select Case COLUMN_NAME
            Case "CARRIER_PROD_CODE"
                sql_where = "CARRIER_CODE = '" & MyBase.Absx1.txtFor("CARRIER_CODE").Text & "'"
        End Select

    End Sub

    Public Overrides Sub Set_ScreenMode_Special(tf As Boolean)
        MyBase.Set_ScreenMode_Special(tf)

        If tf AndAlso EntryMode = "New" Then
            MyBase.Absx1.txtFor("SHIP_VIA_CODE_3PL").Text = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text.Trim
        End If
    End Sub

    Overrides Sub Proceed_PreReq_Special(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                MyBase.Absx1.txtFor("SHIP_VIA_CODE_3PL").Text = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text.Trim

            Case "Update"

                grdSOTSVIAA.PerformAction(UltraWinGrid.UltraGridAction.CommitRow)

                MyBase.Absx1.txtFor("CARRIER_CODE").Text = MyBase.Absx1.txtFor("CARRIER_CODE").Text.Trim
                MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text = MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text.Trim
                Validate_Code("CARRIER_CODE", False, True)

                'SOTCARR2, CARRIER_CODE, CARRIER_PROD_CODE
                If EMsg.Length > 0 AndAlso MyBase.Absx1.txtFor("CARRIER_CODE").TextLength > 0 _
                    AndAlso MyBase.Absx1.txtFor("CARRIER_PROD_CODE").TextLength > 0 Then
                    If LookUp("SOTCARR2", New String() {MyBase.Absx1.txtFor("CARRIER_CODE").Text, MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text}) Is Nothing Then
                        EMsg &= "Invalid Carrier Product for the supplied Product."
                    End If
                End If

                MyBase.Absx1.txtFor("SHIP_VIA_CODE_3PL").Text = MyBase.Absx1.txtFor("SHIP_VIA_CODE_3PL").Text.Trim

        End Select
    End Sub

    Overrides Sub Proceed_Update_Special_Pre()
        Dim sqlDelete = ""
        Update_Record_TDA("SOTSVIAA", "SHIP_VIA_CODE = '" & Absx1.txtFor("SHIP_VIA_CODE").Text & "'")
    End Sub

    Overrides Sub Show_Record_Special()

        If EntryMode = "New" Then
        End If

        EnforceConstraints(False)
        Fill_Records("SOTSVIAA", New String() {Absx1.txtFor("SHIP_VIA_CODE").Text})
        EnforceConstraints(True)

    End Sub

    Overrides Sub Clear_Record_Special()
        If ScreenMode Then
            EnforceConstraints(False)
            dst.Tables("SOTSVIAA").Rows.Clear()
            EnforceConstraints(True)
        End If
    End Sub

    Private Sub grdSOTSVIAA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdSOTSVIAA.BeforeRowUpdate
        e.Row.Cells("SHIP_VIA_CODE").Value = MyBase.Absx1.txtFor("SHIP_VIA_CODE").Text

        e.Row.Cells("SHIP_VIA_ALT").Value = (e.Row.Cells("SHIP_VIA_ALT").Value).ToString.Trim.ToUpper
        e.Row.Cells("ORDR_NO_WEB_PREFIX").Value = (e.Row.Cells("ORDR_NO_WEB_PREFIX").Value).ToString.Trim.ToUpper

        If (e.Row.Cells("SHIP_VIA_ALT").Value).ToString.Trim.Length = 0 Then
            e.Cancel = True
            MessageBox.Show("Alternate Ship Via is required.", "Update Error", MessageBoxButtons.OK)
            Exit Sub
        End If

        Dim SHIP_VIA_ALT As String = e.Row.Cells("SHIP_VIA_ALT").Value
        If LookUp("SOTSVIA1", SHIP_VIA_ALT) Is Nothing Then
            e.Cancel = True
            MessageBox.Show("Alternate Ship Via is Invalid.", "Update Error", MessageBoxButtons.OK)
            Exit Sub
        End If

        If Val(e.Row.Cells("TRANSIT_DAYS").Value & String.Empty) <= 0 Then
            e.Cancel = True
            MessageBox.Show("Transit days must be greater equal 1.", "Update Error", MessageBoxButtons.OK)
            Exit Sub
        End If


    End Sub

    'Public Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
    '    Select Case COLUMN_NAME
    '        Case "CARRIER_PROD_CODE"
    '            Dim rowSOTCARR2 As DataRow = LookUp("SOTCARR2", New String() {MyBase.Absx1.txtFor("CARRIER_CODE").Text, MyBase.Absx1.txtFor("CARRIER_PROD_CODE").Text})
    '            If rowSOTCARR2 IsNot Nothing Then

    '            End If
    '    End Select
    'End Sub

    Private Sub grdSOTSVIAA_ClickCellButton(sender As Object, e As UltraWinGrid.CellEventArgs) Handles grdSOTSVIAA.ClickCellButton

        Dim COLUMN_NAME As String = e.Cell.Column.Key

        Select Case COLUMN_NAME

            Case "SHIP_VIA_ALT"
                Dim sql_where As String = ""
                grdClickCellButton(grdSOTSVIAA, sql_where, False, "", "SHIP_VIA_CODE")

        End Select
    End Sub


End Class