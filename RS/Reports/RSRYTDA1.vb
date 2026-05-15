Public Class RSRYTDA1

    Dim CALENDAR As String = "R"
    Dim FACTOR As Integer = 1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, -1)
        Get_PARM("ICTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        If Absx1.chkFor("THOUSANDS").Checked Then
            FACTOR = 1000
        Else
            FACTOR = 1
        End If

        Dim RSTYTDA1 As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")
        ASCMAIN1.sql = "Select ICTITEM1.* from ICTITEM1"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, , 1)

        Dim RYP_01 As String = Mid(RYP, 1, 4) & "01"
        Dim LYP As String = ASCMAIN1.Period_Calc(RYP, -12)
        Dim LYP_01 As String = Mid(LYP, 1, 4) & "01"

        If optCALENDAR.Value = "R" Then
            LYP_01 = ASCMAIN1.Period_Calc(Mid(LYP, 1, 4) & "02", -1 * ASCMAIN1.PCO)
            RYP_01 = ASCMAIN1.Period_Calc(Mid(RYP, 1, 4) & "02", -1 * ASCMAIN1.PCO)
            If RYP_01 > RYP Then
                LYP_01 = ASCMAIN1.Period_Calc(LYP_01, -12)
                RYP_01 = ASCMAIN1.Period_Calc(RYP_01, -12)
            End If
        End If
        If optXTD.Value = "M" Then
            LYP_01 = LYP
            RYP_01 = RYP
        ElseIf optXTD.Value = "S" Then
            Dim P As Integer = ASCMAIN1.Period_Diff(RYP_01, RYP)
            If P >= 6 Then
                LYP_01 = ASCMAIN1.Period_Calc(LYP_01, 6)
                RYP_01 = ASCMAIN1.Period_Calc(RYP_01, 6)
            End If
        End If

        Dim ICTITEM1 As String = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP, chkHISTCAT.Checked)

        Fill_Records("ICTITEM1", "", True, "Select * from " & ICTITEM1)


        Dim YP(1, 1) As String
        YP(0, 0) = LYP_01
        YP(0, 1) = LYP
        YP(1, 0) = RYP_01
        YP(1, 1) = RYP

        Dim FX As String = ""
        If FACTOR <> 1 Then
            FX = "/" & CStr(FACTOR)
        End If

        For Each SOURCE_TABLE_NAME As String In New String() {"RSTRETL1", "SOTINVH2"}
            Dim YPC As String = IIf(SOURCE_TABLE_NAME = "RSTRETL1", "OPS_YYYYPP", "ORDR_YYYYPP_UPDATED")
            sql_filter = " and RSTYTDA1." & YPC & " BETWEEN '" & LYP_01 & "' AND '" & RYP & "'"

            Dim COL_QTY As String = IIf(SOURCE_TABLE_NAME = "RSTRETL1", "RSTYTDA1.QTY_SOLD", "RSTYTDA1.ORDR_QTY_SHIP")
            Dim COL_SLS As String = IIf(SOURCE_TABLE_NAME = "RSTRETL1", "RSTYTDA1.AMT_SOLD", "RSTYTDA1.ORDR_QTY_SHIP * RSTYTDA1.ORDR_UNIT_PRICE")

            Dim sql_Data As String = ""
            Dim cols As String = ""

            For Y As Integer = 0 To 1

                Dim COLPFX As String = IIf(Y = 0, "LY", "TY") & "_" & IIf(SOURCE_TABLE_NAME = "RSTRETL1", "RTL", "WSL")

                sql_Data &= "" _
                & ", Sum (CASE WHEN RSTYTDA1." & YPC & " BETWEEN '" & YP(Y, 0) & "' AND '" & YP(Y, 1) & "' THEN " & COL_QTY & FX & " ELSE 0 END) " & COLPFX & "_QTY" & vbCrLf _
                & ", Sum (CASE WHEN RSTYTDA1." & YPC & " BETWEEN '" & YP(Y, 0) & "' AND '" & YP(Y, 1) & "' THEN " & COL_SLS & FX & " ELSE 0 END) " & COLPFX & "_SLS" & vbCrLf

                cols &= "," & COLPFX & "_QTY"
                cols &= "," & COLPFX & "_SLS"
            Next

            sql_TABLE_NAMEs = Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1 & " ICTITEM1")

            sql = "Select " & sql_SELECT_cols & vbCrLf & "" & vbCrLf & sql_Data _
            & " from " & SOURCE_TABLE_NAME & " RSTYTDA1 " & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols

            ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & cols & ")" & vbCrLf _
            & "(" & sql & ")"
            ASCDATA1.ExecuteSQL()
        Next

        Dim sqlx As String = ""
        For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
            sqlx &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        Next
        ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlx))

    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))

        SUBT = optXTD.Text

        If CALENDAR = "O" Then
            SUBT &= " - Operatons Calendar Period Ending " & RYPLEGEND
        Else
            SUBT &= " - Retail Calendar Period Ending " & RYPLEGEND
        End If
        Generate_Report(RPT, , SUBT)

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.cmbFor("RYP").Value & "" = "" Then
                EMsg &= vbCr & "You must Specify a Reporting Period"
            End If
        End If
    End Sub

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        Exit Sub
        'ASCDATA1.ExecuteSQL("Alter Table " & TT & " Add ITEM_CODE VARCHAR2(25)")

        'Dim G As Int16 = COLUMN_NAMEs.Count
        'If COLUMN_NAMEs(G - 1) = "ITEM_CODE" Then

        '    ASCMAIN1.sql = "Update " & TT & " Set ITEM_CODE = SUBSTR(G" & CStr(G) & "," & CStr(Len(COLUMN_CAPTIONs(G - 1)) + 2) & ")"
        '    ASCDATA1.ExecuteSQL()
        'End If
    End Sub

    Overrides Sub Build_Report_File_Post_Process()
        dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_QTY").Expression = "ISNULL(TY_WSL_QTY,0)-ISNULL(LY_WSL_QTY,0)"
        dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_QTY_PCT").Expression = "IIF(ISNULL(LY_WSL_QTY,0)=0,0,100*TY_WSL_VAR_LY_QTY/LY_WSL_QTY)"
        dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_SLS").Expression = "ISNULL(TY_WSL_SLS,0)-ISNULL(LY_WSL_SLS,0)"
        dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_SLS_PCT").Expression = "IIF(ISNULL(LY_WSL_SLS,0)=0,0,100*TY_WSL_VAR_LY_SLS/LY_WSL_SLS)"

        dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_QTY").Expression = "ISNULL(TY_RTL_QTY,0)-ISNULL(LY_RTL_QTY,0)"
        dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_QTY_PCT").Expression = "IIF(ISNULL(LY_RTL_QTY,0)=0,0,100*TY_RTL_VAR_LY_QTY/LY_RTL_QTY)"
        dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_SLS").Expression = "ISNULL(TY_RTL_SLS,0)-ISNULL(LY_RTL_SLS,0)"
        dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_SLS_PCT").Expression = "IIF(ISNULL(LY_RTL_SLS,0)=0,0,100*TY_RTL_VAR_LY_SLS/LY_RTL_SLS)"

    End Sub
End Class