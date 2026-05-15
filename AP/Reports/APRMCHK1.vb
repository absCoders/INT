Public Class APRMCHK1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date

    Dim APTCHKR1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")

        grpCHECK_DATE_RANGE.Top = grpPERIOD_RANGE.Top
        grpCHECK_DATE_RANGE.Left = grpPERIOD_RANGE.Left

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        'Absx1.cmbFor("RYP0").Value = Absx1.cmbFor("RYP0").Rows(0).Cells(0).Value
        'Absx1.cmbFor("RYP1").Value = Absx1.cmbFor("RYP1").Rows(0).Cells(0).Value
        Absx1.cmbFor("RYP0").Value = ASCMAIN1.CYP
        Absx1.cmbFor("RYP1").Value = ASCMAIN1.CYP
    End Sub

    Protected Overrides Sub Build_Workfile()
        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Payments Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Payments Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
        Else
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Payments Posted in " & xRYP0_legend
            Else
                SUBT = "Payments Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
        End If

        APTCHKR1 = TAC.APCMAIN1.Prepare_Check_Register(Me, dst, False, xRYP0, xRYP1, xDTE0, xDTE1)

        Check_if_Empty("APTCHKR1")
    End Sub

    Public Overrides Sub Print_Report()
        RPT = "APRCHKR1"
        Generate_Report(RPT, , SUBT)
        'Call Print_GL()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            Else
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If

                If Format(Absx1.dteFor("DTE0").Value, "yyyyMMdd") > Format(Absx1.dteFor("DTE1").Value, "yyyyMMdd") Then
                    EMsg &= vbCr & "Starting Date may not be later than End Date"
                End If
            End If
        End If

    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpCHECK_DATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
        Else
            'Absx1.cmbFor("RYP0").Value = ""
            'Absx1.cmbFor("RYP1").Value = ""
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub
End Class