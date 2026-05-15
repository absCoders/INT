Public Class WHRAUDIT

#Region "General Declarations"
    Private xDTE0 As Date
    Private xDTE1 As Date
    Private shipments As String = String.Empty

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Absx1.optFor("RANGE").CheckedIndex = 2
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Left = grpDATE_RANGE.Left
        grpPERIOD_RANGE.Top = grpDATE_RANGE.Top

        Absx1.dteFor("DTE0").MaxDate = DateTime.Now
        Absx1.dteFor("DTE1").MaxDate = DateTime.Now

        Absx1.dteFor("DTE0").MinDate = DateAdd(DateInterval.Year, -2, DateTime.Now)
        Absx1.dteFor("DTE1").MinDate = DateAdd(DateInterval.Year, -2, DateTime.Now)

        Absx1.dteFor("DTE0").DateTime = DateAdd(DateInterval.Month, -1, DateTime.Now)
        Absx1.dteFor("DTE1").DateTime = Absx1.dteFor("DTE1").MaxDate

        optRANGE.Value = "D"

    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "N"
        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Shipments Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Shipments Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = " and WHTSHIP1.SHIP_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'" & vbCrLf
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            If RYP0 = RYP1 Then
                SUBT = "Shipments in " & RYPLEGEND0
            Else
                SUBT = "Shipments between " & RYPLEGEND0 & " and " & RYPLEGEND1
            End If
            sqlw = " and WHTSHIP1.OPS_YYYYPP between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf
            RWU = "N"
        End If

        sqlw &= SQL_in("CARRIER_CODE", "WHTSHIP1.CARRIER_CODE") & vbCrLf
        sqlw &= SQL_in("SHIP_VIA_CODE", "WHTSHIP1.SHIP_VIA_CODE") & vbCrLf
        sqlw &= SQL_in("CUST_CODE", "WHTSHIP1.CUST_CODE") & vbCrLf

        sqlw = ASCMAIN1.SQL_Add_WHERE(sqlw)

        Prepare_dst(True, sqlw)

        Check_if_Empty("WHTSHIP1")
    End Sub

    Public Overrides Sub Print_Report()
        CR_params.Add("SUBT", "")
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
            Else

            End If
        End If

    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged

        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        grpDATE_RANGE.Enabled = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Enabled = (optRANGE.Value = "P")

        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            'Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            'Absx1.dteFor("DTE0").Value = dates(1)
            'Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()

    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = ""

        If parms.Length > 0 Then
            sqlw = parms(0)
        Else
            sqlw = " where rownum < 1"
        End If

        With dst
            shipments = ASCMAIN1.Temp_Table("Select * from WHTSHIP1 " & sqlw)

            For Each tableName As String In New String() {"WHTSHIP1", "WHTSHIP2", "WHTSHIP3"}
                ASCMAIN1.sql = "select * from " & tableName & " where SHIP_CNTL_NO in (select SHIP_CNTL_NO from " & shipments & ")"
                Create_TDA(.Tables.Add, tableName, "**")
            Next
            .Tables("WHTSHIP1").Columns.Add("TP", GetType(System.String))

            ASCMAIN1.sql = "select * from ARTCUST1 where cust_code in (select CUST_CODE from " & shipments & ")"
            Create_TDA(.Tables.Add, "ARTCUST1", "**")

            Create_TDA(.Tables.Add, "SOTSVIA1", "*")
            Create_TDA(.Tables.Add, "SOTCARR1", "*")
        End With

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""

        If parms.Length > 0 Then
             sqlw = parms(0)
        End If

        If shipments.Length = 0 Then
            shipments = ASCMAIN1.Temp_Table("SELECT * FROM WHTSHIP1 WHERE ROWNUM < 1")
        End If

        ASCDATA1.ExecuteSQL("TRUNCATE TABLE " & shipments)
        ASCDATA1.ExecuteSQL("INSERT INTO " & shipments & " SELECT * FROM WHTSHIP1 " & sqlw)

        EnforceConstraints(False)
        For Each tableName As String In New String() {"WHTSHIP1", "WHTSHIP2", "WHTSHIP3"}
            ASCMAIN1.sql = "select * from " & tableName & " where SHIP_CNTL_NO in (select SHIP_CNTL_NO from " & shipments & ")"
            Fill_Records(tableName, String.Empty, True, ASCMAIN1.sql)
        Next

        ASCMAIN1.sql = "select * from ARTCUST1 where CUST_CODE in (select CUST_CODE from " & shipments & ")"
        Fill_Records("ARTCUST1", String.Empty, True, ASCMAIN1.sql)
        Fill_Records("SOTSVIA1", String.Empty, True, "SELECT * FROM SOTSVIA1")
        Fill_Records("SOTCARR1", String.Empty, True, "SELECT * FROM SOTCARR1")

        For Each rowWHTSHIP3 As DataRow In dst.Tables("WHTSHIP3").Select()
            dst.Tables("WHTSHIP1").Select("SHIP_CNTL_NO = '" & rowWHTSHIP3.Item("SHIP_CNTL_NO") & "'")(0).Item("TP") = "*"
        Next

        EnforceConstraints(True)

    End Sub

End Class