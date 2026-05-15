Public Class SORCANC1

    Dim tempTable As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = String.Empty ' SQL_in("ORDR_GROUP_NO", "SOTORDR0.ORDR_GROUP_NO")
        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTORDR0")
    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        Create_Temp_Data(sqlw)

        With dst

            Create_TDA(dst.Tables.Add, "SOTORDR0", "*", , False)
            Create_TDA(dst.Tables.Add, "SOTORDR1", "*", , False)
            Create_TDA(dst.Tables.Add, "SOTORDR2", "*", , False)
            Create_TDA(dst.Tables.Add, "SOTPICK1", "*", , False)
            Create_TDA(dst.Tables.Add, "SOTPICK2", "*", , False)
            Create_TDA(dst.Tables.Add, "SOTPICKC", "*", , False)

        End With

        If perform_fill Then
            Fill_Records_RPT("")
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)
        'If sqlw <> "" Then
        Create_Temp_Data(sqlw)
        'End If

        EnforceConstraints(False)

        ASCMAIN1.sql = "SELECT * FROM SOTORDR0 WHERE ORDR_GROUP_NO IN (SELECT ORDR_GROUP_NO FROM " & tempTable & ")"
        Fill_Records("SOTORDR0", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM SOTORDR1 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & tempTable & ")"
        Fill_Records("SOTORDR1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM SOTORDR2 WHERE ORDR_NO IN (SELECT ORDR_NO FROM " & tempTable & ")"
        Fill_Records("SOTORDR2", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM SOTPICK1 WHERE PICK_NO IN (SELECT PICK_NO FROM " & tempTable & ")"
        Fill_Records("SOTPICK1", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM SOTPICK2 WHERE PICK_NO IN (SELECT PICK_NO FROM " & tempTable & ")"
        Fill_Records("SOTPICK2", String.Empty, True, ASCMAIN1.sql)

        ASCMAIN1.sql = "SELECT * FROM SOTPICKC WHERE PICK_NO IN (SELECT PICK_NO FROM " & tempTable & ")"
        Fill_Records("SOTPICKC", String.Empty, True, ASCMAIN1.sql)

        EnforceConstraints(True)

    End Sub

    Sub Create_Temp_Data(SQLW As String)

        ASCMAIN1.sql = "Select DISTINCT SOTORDR1.ORDR_GROUP_NO, SOTPICK1.PICK_NO, SOTPICK1.ORDR_NO" _
        & " FROM SOTPICKC, CONV.CFG_CNCLB4SHP CNCLB4SHP, SOTORDR1, SOTPICK1" _
        & " WHERE SOTPICKC.ORDR_NO_3PL = CNCLB4SHP.OHORD#" _
        & " AND SOTPICKC.PICK_NO = SOTPICK1.PICK_NO" _
        & " AND SOTPICKC.ORDR_NO = SOTPICK1.ORDR_NO" _
        & " AND SOTPICK1.PICK_STATUS = 'P'" _
        & " AND SOTPICK1.ORDR_NO = SOTORDR1.ORDR_NO" _
        & SQLW

        If tempTable.Length = 0 Then
            tempTable = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        Else
            ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & tempTable)
            ASCDATA1.ExecuteSQL("INSERT INTO " & tempTable & " " & ASCMAIN1.sql)
        End If

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