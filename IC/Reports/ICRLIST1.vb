Public Class ICRLIST1
    Dim ICTLIST1 As String
   
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Prepare_Work_File()

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""
        If chkActive.Checked Then
            sql_filter &= " AND ICTITEM1.ITEM_STATUS = 'A'"
        End If

        ' Extracts from Data Sources

        ASCMAIN1.sql = "Select ICTITEM1.* " _
            & " from ICTITEM1" & ASCMAIN1.SQL_Add_WHERE(sql_filter)

        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ICTITEM1", 1))

        Dim SOURCE_TABLE_NAME As String = "ICTITEM1"
        MyBase.Get_SQL("*", SOURCE_TABLE_NAME)

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", " & SOURCE_TABLE_NAME & ".ITEM_CODE" _
        & " from " & ICTLIST1 & " " & SOURCE_TABLE_NAME & " " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")
    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        CR_params.Add("NOTES", IIf(chkNotes.Checked, "Y", "N"))
        CR_params.Add("DESC2", IIf(chkDESC2.Checked, "Y", "N"))
        Generate_Report(RPT, , SUBT)
    End Sub

    Sub Prepare_Work_File()

        ICTLIST1 = "ICTITEM1"
        'ASCMAIN1.sql = "Select * from ICTITEM1"
        'ICTLIST1 = ASCMAIN1.Temp_Table
        'ASCMAIN1.sql = "Select * from " & ICTLIST1
        'dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTLIST1", 1))

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            End If
        End If
    End Sub

End Class