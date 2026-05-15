Public Class ARRCBDA1

    ' pgm does not take into account DISC & WOFF from ARTPYMT3, nor does it take into account GL write-offs

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -36, 0, -11)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -36, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        ASCMAIN1.Progress("Work Tables")

        Dim ARTCBDA1 As String = ""
        TAC.ARCMAIN1.Create_ARTCBDA1(ARTCBDA1, RYP0, RYP1)

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        ASCMAIN1.Progress("Analysis")
        MyBase.Get_SQL("*")

        sql_Data = "" _
            & ", SUM (BEG_B) BEG_B" & vbCrLf _
            & ", SUM (NEW_B) NEW_B" & vbCrLf _
            & ", SUM (APP_B) APP_B" & vbCrLf _
            & ", SUM (END_B) END_B" & vbCrLf _
            & ", SUM (BEG_C) BEG_C" & vbCrLf _
            & ", SUM (NEW_C) NEW_C" & vbCrLf _
            & ", SUM (APP_C) APP_C" & vbCrLf _
            & ", SUM (END_C) END_C" & vbCrLf _
            & ", SUM (NEW_X) NEW_X" & vbCrLf

        sql_Cols = "" _
            & ",BEG_B,NEW_B,APP_B,END_B,BEG_C,NEW_C,APP_C,END_C,NEW_X"

        sql_filter = ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & "" & vbCrLf & sql_Data _
            & " from " & ARTCBDA1 & " ARTCBDA1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()
    End Sub

    Public Overrides Sub Print_Report()
        If RYP0 = RYP1 Then
            SUBT = RYPLEGEND0
        Else
            SUBT = RYPLEGEND0 & " thru " & RYPLEGEND1
        End If

        Generate_Report(RPT, , SUBT)
    End Sub
     
End Class