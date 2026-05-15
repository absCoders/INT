Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing.Printing

Public Class SOTUCCL1
    Private Sub Form_Activated(sender As Object, e As System.EventArgs) Handles Me.Activated
        SetUpPortsAndPrinters()
    End Sub

    Public Overrides Sub Mode_Settings(tf As Boolean, Optional MODE_description As String = "")
        MyBase.Mode_Settings(tf, MODE_description)

        btnTest.Visible = tf
        lblCartonNo.Visible = tf
        txtCartonNo.Visible = tf
    End Sub

    Private Sub btnTest_Click(sender As System.Object, e As System.EventArgs) Handles btnTest.Click

        Dim UCC128 As String = Absx1.txtFor("UCC128_COMMANDS").Text

        If UCC128 = "" Then
            MsgBox("Nothing to Test", MsgBoxStyle.OkOnly, "No Template Defined")
            Exit Sub
        End If

        If txtCartonNo.Text <> "" Then
            Dim cartonLabel As New CartonLabel(txtCartonNo.Text)
            cartonLabel.PrintLabel(1)
        Else
            ShippingLabel.SendToLabelPrinter(UCC128)
        End If
    End Sub

    Private Sub SetUpPortsAndPrinters()
        Dim tooltip As New System.Windows.Forms.ToolTip()

        ' Label Printer Port
        Try
            txtLabelPrinter.BackColor = Drawing.Color.Red

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtLabelPrinter.Text = ASCMAIN1.LabelPrinterSerialPort.PortName
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            Else
                Me.txtLabelPrinter.Text = "No Port"
                tooltip.SetToolTip(txtLabelPrinter, txtLabelPrinter.Text)
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing AndAlso Not ASCMAIN1.LabelPrinterSerialPort.IsOpen Then
                ASCMAIN1.LabelPrinterSerialPort.Open()
            End If

            If ASCMAIN1.LabelPrinterSerialPort IsNot Nothing Then
                txtLabelPrinter.BackColor = Drawing.Color.Green
            End If

        Catch ex As Exception
            txtLabelPrinter.BackColor = Drawing.Color.Red
        End Try

    End Sub


End Class