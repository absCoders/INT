Public Class ARRCUST1
    Dim CHKACTIVE_ONLY As String

    Private Sub ASFSPRF1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Private Sub ASFSPRF1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Call Mode_Settings(False)
    End Sub

    Protected Overrides Sub Build_Workfile()
        Call MyBase.Build_WorkFile_DB_Init()
        Call ASCMAIN1.Track("Run-Time Options", "")

        COLUMN_NAME_last = "CUST_CODE"
        Call MyBase.Get_Codes()

        'CHKDTL = opts("CHKSHOW_ALL_ADDRESS_LINES")
        '        CHKACTIVE_ONLY = opts("CHKACTIVE_ONLY")

        ' Set up Work File Definition using X's and 0's and 0.01's as required

        Call ASCMAIN1.Track("Initialize Work Tables", "")
        Dim sql As String
        sql = "Select 0 RECORD_NO, 0 COUNTER, CUST_CODE, "
        'sql = sql & Get_GString(False)
        sql = sql & " from ARTCUST1 where " & ASCMAIN1.DBS_NOROWS
        tblASTSRPT0 = ASCDATA1.GetDataTable(sql, "ASTSRPT0", 1)
        dst.Tables.Add(tblASTSRPT0)

        ''Call ASCDATA1.Create_Index(AccD, "ASTSRPT0", "Report", Get_GString(True))
        ''Dim dynASTSRPT0 As DAO.Recordset
        ''dynASTSRPT0 = AccD.OpenRecordset("ASTSRPT0", DAO.RecordsetTypeEnum.dbOpenTable)

        ASCMAIN1.sql = "Select * from ARTCUST1 where ROWNUM < 1"
        Dim tblARWCUST1 As DataTable = ASCDATA1.GetDataTable("", "ARWCUST1")

        Dim r As Long
        Dim i As Integer
        Dim j As Integer

        Call MyBase.Get_SQL("*")

        ' Prepare Work File with Data from Server

        Dim sqlx As String
        sqlx = ""
        If CHKACTIVE_ONLY = "1" Then
            sqlx = " and ARTCUST1.CUST_STATUS = 'A'"
        End If

        Call ASCMAIN1.Track("Initialize Server", "")
        sql = "Select " & sql_SELECT_cols & ","
        sql = sql & " ARTCUST1.*"
        sql = sql & " from ARTCUST1" & sql_JOIN
        sql = sql & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sqlx)
        '    sql = sql & sqljoin
        '    sql = sql & sqlwhere
        '    sql = sql & " group by " & sqllist2 &  y

        ASCMAIN1.oraCmd.CommandText = sql
        Dim dynX As OracleDataReader = ASCMAIN1.oraCmd.ExecuteReader

        r = 0
        Call ASCMAIN1.Track("Record", "")
        Do While dynX.Read
            r = r + 1
            If r Mod 100 = 0 Then
                Call ASCMAIN1.Track("-", dynX.Item(0) & "-" & CStr(r))
            End If

            Dim rowASTSRPT0 As DataRow = tblASTSRPT0.NewRow

            rowASTSRPT0.Item("RECORD_NO") = r
            rowASTSRPT0.Item("CUST_CODE") = dynX.Item("CUST_CODE")
            rowASTSRPT0.Item("COUNTER") = 1

            For j = 1 To COLUMN_NAMEs.Count
                rowASTSRPT0.Item("G" & Format$(j, "0")) = COLUMN_CAPTIONs(j - 1) & ":" & dynX.Item(j - 1)
            Next j
            tblASTSRPT0.Rows.Add(rowASTSRPT0)


            Dim rowARWCUST1 As DataRow = tblARWCUST1.NewRow
            For i = 0 To tblARWCUST1.Columns.Count - 1
                rowARWCUST1.Item(i) = dynX.Item(COLUMN_NAMEs.Count + i)
            Next
            tblARWCUST1.Rows.Add(rowARWCUST1)
        Loop

        ' Prepare Report File, w/Consolidations & Recaps as required

        MyBase.Build_Report_File()

        ' Wrap up
        dst.Tables.Add(tblARWCUST1)

        Stop
    End Sub

    Overrides Sub Print_Report()
        CR_params.Add("CHKSEG2", "1")
        CR_params.Add("CHKSEG3", "0")
        CR_params.Add("CHKSEG4", "0")

        Generate_Report(RPT)
    End Sub

End Class