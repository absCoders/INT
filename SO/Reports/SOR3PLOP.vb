Public Class SOR3PLOP

    Dim tempTable As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = String.Empty ' SQL_in("ORDR_GROUP_NO", "SOTORDR0.ORDR_GROUP_NO")
        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTPICKO")
    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        With dst

            Create_TDA(dst.Tables.Add, "SOTORDR1", "*", , False)
            Create_TDA(dst.Tables.Add, "SOTORDR2", "*", , False)
            Create_TDA(dst.Tables.Add, "SOTPICKO", "*", , False)

            .Tables.Add("LEGEND")
            With .Tables("LEGEND")
                .Columns.Add("KEY", GetType(System.String))
                .Columns.Add("DESC", GetType(System.String))
            End With

            dst.Tables("LEGEND").Rows.Add(New Object() {"A", "Open in Clarins and Open In Absolution"})
            dst.Tables("LEGEND").Rows.Add(New Object() {"B", "Open in ABSolution Not Open In Clarins"})
            dst.Tables("LEGEND").Rows.Add(New Object() {"C", "Open in Clarins Cancelled In Absolution"})
            dst.Tables("LEGEND").Rows.Add(New Object() {"F", "Open in Clarins Finalized In Absolution"})
            dst.Tables("LEGEND").Rows.Add(New Object() {"V", "Open in Clarins Voided In Absolution"})
            dst.Tables("LEGEND").Rows.Add(New Object() {"Z", "Open in Clarins NOT in ABSolution"})

            Create_Relation("SOTORDR1", "SOTORDR2", "ORDR_NO")
            dst.Tables("SOTORDR2").Columns.Add("PICK_AMT", GetType(System.Int64), "ORDR_UNIT_PRICE * ORDR_QTY_PICK")
            dst.Tables("SOTORDR1").Columns.Add("PICK_AMT", GetType(System.Int64), "SUM(CHILD.PICK_AMT)")

        End With

        If perform_fill Then
            Fill_Records_RPT("")
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)
        Create_Temp_Data(sqlw)

        EnforceConstraints(False)

        ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & tempTable & ")"
        Fill_Records("SOTORDR1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM SOTORDR2 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & tempTable & ")"
        Fill_Records("SOTORDR2", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM " & tempTable
        Fill_Records("SOTPICKO", String.Empty, True, ASCMAIN1.sql)

        ' Need to fake sotordr1 numbers for links in report
        Dim iOrderNo As Int16 = 0
        Dim ORDR_NO As String = String.Empty
        For Each rowSOTPICKO As DataRow In dst.Tables("SOTPICKO").Select("ORDR_NO IS NULL")
            ORDR_NO = "C" & rowSOTPICKO.Item("ORDR_NO_3PL")
            rowSOTPICKO.Item("ORDR_NO") = ORDR_NO
            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow
            rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
            rowSOTORDR1.Item("CUST_CODE") = "Unk"
            rowSOTORDR1.Item("CUST_STORE_NO") = "000000"
            rowSOTORDR1.Item("CUST_NAME") = "Unknown Customer"
            rowSOTORDR1.Item("FRT_TERMS") = "*"
            rowSOTORDR1.Item("WHSE_CODE") = "CLA"
            rowSOTORDR1.Item("ORDR_STATUS") = "P"
            rowSOTORDR1.Item("CURR_CODE") = "USD"
            rowSOTORDR1.Item("CURR_EXCH_RATE") = 1
            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)
        Next

        EnforceConstraints(True)

    End Sub

    Sub Create_Temp_Data(SQLW As String)

        ASCMAIN1.sql = "Select * from SOTPICKO"

        If tempTable.Length = 0 Then
            tempTable = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        Else
            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & tempTable)
            ASCDATA1.ExecuteSQL("INSERT INTO " & tempTable & " " & ASCMAIN1.sql)
        End If

        ASCDATA1.ExecuteSQL("UPDATE " & tempTable & " SET PICK_STATUS = 'B' WHERE ORDR_NO_3PL LIKE 'C%'")
        ASCDATA1.ExecuteSQL("UPDATE " & tempTable & " SET PICK_STATUS = 'Z' WHERE PICK_NO IS NULL")
        ASCDATA1.ExecuteSQL("UPDATE " & tempTable & " SET PICK_STATUS = 'A' WHERE PICK_STATUS = 'P'")

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "ORDR_GROUP_NO"
                sqlw = ""
        End Select
        Return sqlw
    End Function
End Class