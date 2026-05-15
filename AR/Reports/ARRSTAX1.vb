Imports System.Math

Public Class ARRSTAX1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date

    Dim SOTINVH1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("SOTPARM1")
        Call Get_PARM("ARTPARM1")


        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        Absx1.optFor("RANGE").CheckedIndex = 0
        'grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left
    End Sub

    Protected Overrides Sub Build_Workfile()

        dst.EnforceConstraints = False
        SUBT = ""

        Dim sqlw As String = "" ' "INV_STAX <> 0"

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Invoices Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Invoices Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw &= " SOTINVH1.INV_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Invoices Posted in " & xRYP0_legend
            Else
                SUBT = "Invoices Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw &= " SOTINVH1.ORDR_YYYYPP_UPDATED between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        End If

        sqlw &= Get_Filter("CUST_CODE", "SOTINVH1.CUST_CODE")
        sqlw &= Get_Filter("STAX_CODE", "SOTINVH1.STAX_CODE")
        sqlw &= Get_Filter("DIVISION_CODE", "SOTINVH1.DIVISION_CODE")
        sqlw &= Get_Filter("STATE_CODE", "SOTINVH1.CUST_SHIP_TO_STATE")

        ASCMAIN1.sql = "Select SOTINVH1.* from SOTINVH1 where " & sqlw
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_NO)")

        ASCMAIN1.sql = "Select SOTINVH1.*,ARTCUST1.CUST_NAME " _
        & " from " & SOTINVH1 & " SOTINVH1,ARTCUST1 " _
        & " where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE"

        Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, False)
        Fill_Records("SOTINVH1")


        Create_TDA(dst.Tables.Add, "ARTSTAX1", "*", 0, False)
        Fill_Records("ARTSTAX1")
        Create_TDA(dst.Tables.Add, "TATSTATE", "*", 0, False)
        Fill_Records("TATSTATE")

        Check_if_Empty("SOTINVH1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
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
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            End If
        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

End Class