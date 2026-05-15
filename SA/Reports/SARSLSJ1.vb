Public Class SARSLSJ1
    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Dim sqlw As String = ""
    Dim SOTINVH1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYW("RYW", ASCMAIN1.CYW, -300, 0, -1)
        Absx1.optFor("RANGE").CheckedIndex = 1

        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left

    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim SATSLSJ1 As String = ""
        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Documents Posted " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Documents Posted between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "SOTINVH1.INV_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
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
            sqlw = "SOTINVH1.ORDR_YYYYPP_UPDATED between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"
        End If
        If Absx1.optFor("OPTDT").Value = "I" Then
            sqlw &= " and SOTINVH1.INV_TYPE = 'I' "
        ElseIf Absx1.optFor("OPTDT").Value = "C" Then
            sqlw &= " and SOTINVH1.INV_TYPE = 'C' "
        Else


        End If
        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        With dst
            sql = "SELECT SOTINVH1.*" _
            & " from SOTINVH1 " _
            & " where " & sqlw

            sql &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE")

            SATSLSJ1 = ASCMAIN1.Temp_Table(sql)

            ASCMAIN1.sql = "Select * from " & SATSLSJ1
            .Tables.Add(ASCDATA1.GetDataTable("**", "SATSLSJ1", 2))
        End With

        Call MyBase.Get_SQL("*", SATSLSJ1)
        Call ASCMAIN1.Progress("Building Tiers")

        sql = "Select " & sql_SELECT_cols & vbCr
        sql &= ", SATSLSJ1.INV_NO, SATSLSJ1.INV_SALES" & vbCr
        sql &= " from " & SATSLSJ1 & " SATSLSJ1 " & sql_TABLE_NAMEs & vbCr
        sql &= ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCr
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        Check_if_Empty("SATSLSJ1")

    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("DTL", Absx1.optFor("OPTDS").Value)
        If Absx1.optFor("OPTDT").Value = "I" Then
            SUBT &= ", Invoices Only"
        ElseIf Absx1.optFor("OPTDT").Value = "C" Then
            SUBT &= ", Credits Only"
        Else
            SUBT &= ""
        End If

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