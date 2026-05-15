Public Class APRLIST1

    Dim APTLIST1 As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        MyBase.Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim sql As String

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        With dst

            sql = "SELECT APTVEND1.* from APTVEND1"
            Dim sqlw As String = ""
            sqlw &= SQL_in("VEND_CODE", "APTVEND1.VEND_CODE")
            sqlw &= SQL_in("VEND_CLASS_CODE", "APTVEND1.VEND_CLASS_CODE")
            sqlw &= SQL_in("PROCESSOR_CODE", "APTVEND1.PROCESSOR_CODE")
            sqlw &= SQL_in("VEND_TYPE", "APTVEND1.VEND_TYPE")
            APTLIST1 = ASCMAIN1.Temp_Table(sql & ASCMAIN1.SQL_Add_WHERE(sqlw))
            ASCMAIN1.sql = "Select * from " & APTLIST1
            .Tables.Add(ASCDATA1.GetDataTable("**", "APTLIST1", 1))
        End With


        Call MyBase.Get_SQL("*", APTLIST1)
        'Call ASCMAIN1.Progress("Building Tiers"

        sql = "Select " & sql_SELECT_cols & vbCr
        sql &= ", APTLIST1.VEND_CODE" & vbCr
        sql &= " from " & APTLIST1 & " APTLIST1 " & sql_TABLE_NAMEs & vbCr
        sql &= ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCr
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        Check_if_Empty("APTLIST1")
    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("DTL", IIf(Absx1.chkFor("DTL").Checked, "1", "0"))
        Generate_Report(RPT, , )
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        If eItemKey = "Proceed" Then

        End If
    End Sub

End Class
