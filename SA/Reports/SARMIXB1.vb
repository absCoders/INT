Public Class SARMIXB1
    Dim RYPs(12) As String
    Dim LEGENDs(12) As String
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables
 
        Dim RYP01 As String = Mid(RYP, 1, 4) & "01"
         
        ' Prepare Work Tables
  
        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""
      
        ASCMAIN1.Progress("Shipments")

        MyBase.Get_SQL("*")

        sql_Data = ""
        sql_Cols = ""
        For M As Integer = 1 To 12
            Dim YP As String = Mid(RYP, 1, 4) & Format(M, "00")
            RYPs(M) = YP
            LEGENDs(M) = ASCMAIN1.Get_Legend(YP)

            If M <= Val(Mid(RYP, 5, 2)) Then
                sql_Data &= "" _
               & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0),0)) UNITS_" & Format(M, "00") & vbCrLf _
               & ", SUM (DECODE(SOTINVH2.ORDR_YYYYPP_UPDATED,'" & YP & "',NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0)) SALES_" & Format(M, "00") & vbCrLf
            Else
                sql_Data &= "" _
                & ", 0 UNITS_" & Format(M, "00") & vbCrLf _
                & ", 0 SALES_" & Format(M, "00") & vbCrLf

            End If
            sql_Cols &= "" _
                & ",UNITS_" & Format(M, "00") & ",SALES_" & Format(M, "00")
        Next M

        sql_filter = " and SOTINVH2.ORDR_YYYYPP_UPDATED BETWEEN '" & RYP01 & "' and '" & RYP & "'" _
            & " and ICTITEM1.ITEM_SNU_CODE = 'S'"

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from SOTINVH2" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

        ' Eliminate 0s

        'Dim sqlz As String = ""
        'For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
        '    sqlz &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        'Next
        'ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlz))
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("INCL_UNITS", IIf(Absx1.chkFor("INCL_UNITS").Checked, "1", "0"))
        CR_params.Add("INCL_SALES", IIf(Absx1.chkFor("INCL_SALES").Checked, "1", "0"))
        CR_params.Add("SHOW_PEN", IIf(Absx1.chkFor("SHOW_PEN").Checked, "1", "0"))
        CR_params.Add("SHOW_TTL", IIf(Absx1.chkFor("SHOW_TTL").Checked, "1", "0"))
        CR_params.Add("MM", Mid(RYP, 5, 2))

        For MM As Integer = 1 To 12
            If MM > Val(Mid(RYP, 5, 2)) Then
                CR_params.Add("M" & Format(MM, "00"), "")
            Else
                CR_params.Add("M" & Format(MM, "00"), Mid(LEGENDs(MM), 10, 3))
            End If

        Next
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Period"
            End If
            If Not Absx1.chkFor("INCL_UNITS").Checked And Not Absx1.chkFor("INCL_SALES").Checked Then
                EMsg &= vbCr & "You must Specify to include Units, Sales or Both"
            End If
        End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        'dst.Tables("ASTSRPT1").Columns("BOOKED_TYTM").Expression = "ISNULL(CARRIED_FWD,0)+ISNULL(BOOKED_PRV,0)+ISNULL(BOOKED_CUR,0)"
        'dst.Tables("ASTSRPT1").Columns("PROJECTED_TYTM").Expression = "ISNULL(SHIPPED_TYTM,0)+ISNULL(OTS_M01,0)"
    End Sub
End Class