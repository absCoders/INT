Public Class SORTREO1

    Dim SOTORDR1 As String
    Dim CHKCUSTSORT As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("SOTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = " and SOTORDR1.REORD_MEMO_IND = '1'"

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        Dim sqlw As String = ""
        sqlw &= SQL_in("CUST_CODE", "SOTORDR1.CUST_CODE")
        If sqlw <> "" Then CHKCUSTSORT = "1"

        ASCMAIN1.sql = "Select ORDR_NO from SOTORDR1" _
            & ASCMAIN1.SQL_Add_WHERE(sql_filter & sqlw)
        SOTORDR1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")

        sql = "Select SOTORDR1.* from SOTORDR1" _
            & " where ORDR_NO in (Select ORDR_NO from " & SOTORDR1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDR1", 1))

        sql = "Select SOTORDR2.* from SOTORDR2" _
            & " where ORDR_NO in (Select ORDR_NO from " & SOTORDR1 & ")" _
            & "   and ORDR_RELEASE in ('R','S','X','Y')"
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "SOTORDR2", 2))

        ASCMAIN1.sql = "Select SOTSREP1.SREP_CODE, SOTSREP1.SREP_NAME from SOTSREP1"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))
 
        If CHKCUSTSORT = "1" Then
            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                rowSOTORDR1.Item("SREP_CODE") = "XX"
            Next
        End If

    End Sub

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""
        CR_params.Add("CHKCUSTSORT", "0") ' IIf(Absx1.chkFor("CHKCUSTSORT").Checked, "1", "0"))
        Generate_Report(RPT, RPT_TITLE, SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"

        End Select
    End Sub

    Overrides Sub Update_Record()
        ASCMAIN1.sql = "Update SOTORDR1 set REORD_MEMO_IND = '2'" _
            & " where REORD_MEMO_IND = '1'"
        ASCDATA1.ExecuteSQL()
    End Sub
End Class