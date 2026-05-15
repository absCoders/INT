Public Class RSRYOYM1

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
        End If

        Dim RSTYOYM1 As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")
        ASCMAIN1.sql = "Select ICTITEM1.* from ICTITEM1"
        Create_TDA(dst.Tables.Add, "ICTITEM1", "**", 0, False, , 1)

        Dim ICTITEM1 As String = RSCMAIN1.Get_ICTITEM1_Hist_CATGY(RYP, chkHISTCAT.Checked)

        Fill_Records("ICTITEM1", "", True, "Select * from " & ICTITEM1)

        Dim FX As String = ""
        If FACTOR <> 1 Then
            FX = "/" & CStr(FACTOR)
        End If

        For Each SOURCE_TABLE_NAME As String In New String() {"RSTRETL1", "SOTINVH2"}
            If (optRW.Value = "R" And SOURCE_TABLE_NAME = "RSTRETL1") _
            Or (optRW.Value = "W" And SOURCE_TABLE_NAME = "SOTINVH2") Then

                Dim YPC As String = IIf(SOURCE_TABLE_NAME = "RSTRETL1", "OPS_YYYYPP", "ORDR_YYYYPP_UPDATED")
                sql_filter = " and RSTYOYM1." & YPC & " BETWEEN '" & ASCMAIN1.Period_Calc(RYP, -60 + 1) & "' AND '" & RYP & "'"

                Dim COL_QTY As String = IIf(SOURCE_TABLE_NAME = "RSTRETL1", "RSTYOYM1.QTY_SOLD", "RSTYOYM1.ORDR_QTY_SHIP")
                Dim COL_SLS As String = IIf(SOURCE_TABLE_NAME = "RSTRETL1", "RSTYOYM1.AMT_SOLD", "RSTYOYM1.ORDR_QTY_SHIP * RSTYOYM1.ORDR_UNIT_PRICE")

                Dim cols As String = ""

                Dim sql_Data As String = ""

                For Y As Integer = 4 To 0 Step -1

                    Dim RYP_01 As String = ASCMAIN1.Period_Calc(RYP, -12 * (Y + 1) + 1)
                    Dim RYP_12 As String = ASCMAIN1.Period_Calc(RYP, -12 * Y)

                    sql_Data &= "" _
                    & ", Sum (CASE WHEN RSTYOYM1." & YPC & " BETWEEN '" & RYP_01 & "' AND '" & RYP_12 & "' THEN " & COL_SLS & FX & " ELSE 0 END) SLS_AMT_Y" & Format(Y, "0") & vbCrLf _
                    & ", Sum (CASE WHEN RSTYOYM1." & YPC & " BETWEEN '" & RYP_01 & "' AND '" & RYP_12 & "' THEN " & COL_QTY & FX & " ELSE 0 END) SLS_QTY_Y" & Format(Y, "0") & vbCrLf _
                    & ", Count (CASE WHEN RSTYOYM1." & YPC & " BETWEEN '" & RYP_01 & "' AND '" & RYP_12 & "' THEN RSTYOYM1.ITEM_CODE ELSE NULL END) SLS_CNT_Y" & Format(Y, "0") & vbCrLf

                    cols &= ",SLS_AMT_Y" & Format(Y, "0")
                    cols &= ",SLS_QTY_Y" & Format(Y, "0")
                    cols &= ",SLS_CNT_Y" & Format(Y, "0")
                Next

                sql_TABLE_NAMEs = Replace(sql_TABLE_NAMEs, ",ICTITEM1", "," & ICTITEM1 & " ICTITEM1")

                sql = "Select " & sql_SELECT_cols & vbCrLf & "" & vbCrLf & sql_Data _
                & " from " & SOURCE_TABLE_NAME & " RSTYOYM1 " & sql_TABLE_NAMEs & vbCrLf _
                & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
                & " group by " & sql_GROUP_BY_cols

                ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
                & "(" & G1thru9 & COLUMN_NAMEs_appended _
                & cols & ")" & vbCrLf _
                & "(" & sql & ")"
                ASCDATA1.ExecuteSQL()
            End If
        Next

        Dim sqlx As String = ""
        For Each COLUMN_NAME As String In COLUMN_NAME_sum.Keys
            sqlx &= " AND NVL(" & COLUMN_NAME & ",0) = 0"
        Next
        ASCDATA1.ExecuteSQL("Delete from " & ASTSRPT1 & ASCMAIN1.SQL_Add_WHERE(sqlx))

    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("THOUSANDS", IIf(Absx1.chkFor("THOUSANDS").Checked, "1", "0"))

        SUBT = optRW.Text & " Sales"

        For Y As Integer = 0 To 4
            Dim RYPX As String = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -12 * Y))

            CR_params.Add("YEAR" & Format(Y, "0") & "AGO", "12 Mos Ending " & Mid(RYPX, 10, 6))
        Next
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
        ' Stop
        'dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_QTY").Expression = "ISNULL(TY_WSL_QTY,0)-ISNULL(LY_WSL_QTY,0)"
        'dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_QTY_PCT").Expression = "IIF(ISNULL(TY_WSL_QTY,0)=0,0,100*TY_WSL_VAR_LY_QTY/TY_WSL_QTY)"
        'dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_SLS").Expression = "ISNULL(TY_WSL_SLS,0)-ISNULL(LY_WSL_SLS,0)"
        'dst.Tables("ASTSRPT1").Columns("TY_WSL_VAR_LY_SLS_PCT").Expression = "IIF(ISNULL(TY_WSL_SLS,0)=0,0,100*TY_WSL_VAR_LY_SLS/TY_WSL_SLS)"

        'dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_QTY").Expression = "ISNULL(TY_RTL_QTY,0)-ISNULL(LY_RTL_QTY,0)"
        'dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_QTY_PCT").Expression = "IIF(ISNULL(TY_RTL_QTY,0)=0,0,100*TY_RTL_VAR_LY_QTY/TY_RTL_QTY)"
        'dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_SLS").Expression = "ISNULL(TY_RTL_SLS,0)-ISNULL(LY_RTL_SLS,0)"
        'dst.Tables("ASTSRPT1").Columns("TY_RTL_VAR_LY_SLS_PCT").Expression = "IIF(ISNULL(TY_RTL_SLS,0)=0,0,100*TY_RTL_VAR_LY_SLS/TY_RTL_SLS)"

    End Sub
End Class