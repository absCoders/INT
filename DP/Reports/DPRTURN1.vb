Public Class DPRTURN1

    Dim MD(13) As String
    Dim MOS As Integer

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        Dim ASTSRPTX As String = ""

        MOS = Val(Absx1.numFor("MOS").Value & "")

        For i = 0 To 12
            MD(13 - i) = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(RYP, -1 * i), False, True)
        Next i
        Dim YPs(24) As String
        For P As Integer = 0 To 24
            YPs(P) = ASCMAIN1.Period_Calc(RYP, -1 * P)
        Next


        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        Dim SOURCE_TABLE_NAME As String = ""
        Dim sql_Data As String = ""
        Dim DATA_COLUMN_NAMEs As String = ""
        Dim sql_WHERE2 As String = ""

        ' Beginning of Month On Hand

        DATA_COLUMN_NAMEs = ""
        sql_Data = Replace(COLUMN_NAMEs_appended, ",", ",ASTSRPTX" & ".")
        For P As Integer = 0 To 24
            Dim DATA_COLUMN_NAME = "B" & Format(P, "00")
            DATA_COLUMN_NAMEs &= "," & DATA_COLUMN_NAME
            sql_Data &= "" _
            & ", Sum(Decode (ASTSRPTX.OPS_YYYYPP,'" & YPs(P) & "',ASTSRPTX.WHSE_QTY_BEG,0)) " & DATA_COLUMN_NAME & vbCrLf
        Next
        ' sqlwhere has the filtrs

        sql_filter = " and ASTSRPTX.OPS_YYYYPP BETWEEN '" & YPs(24) & "' AND '" & YPs(0) & "'"
        sql_WHERE2 = " and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'"
        If Absx1.chkFor("ACTIVE_ONLY").Checked Then
            sql_WHERE2 &= " and ICTITEM1.ITEM_STATUS = 'A'"
        End If

        SOURCE_TABLE_NAME = "ICTSTAT1"

        sql = "Select " & sql_SELECT_cols & vbCrLf & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " ASTSRPTX " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_WHERE2 & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & IIf(sql_GROUP_BY_cols = "", "''", sql_GROUP_BY_cols) & Replace(COLUMN_NAMEs_appended, ",", ",ASTSRPTX" & ".")

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended & DATA_COLUMN_NAMEs & ")" & vbCrLf _
        & "(" & sql & ")"

        ASCDATA1.ExecuteSQL()


        ' Shipments

        DATA_COLUMN_NAMEs = ""
        sql_Data = Replace(COLUMN_NAMEs_appended, ",", ",ASTSRPTX" & ".")
        For P As Integer = 0 To 24
            Dim DATA_COLUMN_NAME = "U" & Format(P, "00")
            DATA_COLUMN_NAMEs &= "," & DATA_COLUMN_NAME
            sql_Data &= "" _
            & ", Sum(Decode (ASTSRPTX.ORDR_YYYYPP_UPDATED,'" & YPs(P) & "',ASTSRPTX.ORDR_QTY_SHIP,0)) " & DATA_COLUMN_NAME & vbCrLf
        Next

        sql_filter = " and ASTSRPTX.ORDR_YYYYPP_UPDATED BETWEEN '" & YPs(24) & "' AND '" & YPs(0) & "'"
        sql_WHERE2 = " and NVL(ICTWHSE1.WHSE_MRP_EXC_IND,'0') <> '1'"
        If Absx1.chkFor("ACTIVE_ONLY").Checked Then
            sql_WHERE2 &= " and ICTITEM1.ITEM_STATUS = 'A'"
        End If
        If Absx1.optFor("GN").Value = "G" Then
            sql_WHERE2 &= " and ASTSRPTX.INV_TYPE = 'I'"
        End If

        SOURCE_TABLE_NAME = "SOTINVH2"

        sql = "Select " & sql_SELECT_cols & vbCrLf & sql_Data _
        & " from " & SOURCE_TABLE_NAME & " ASTSRPTX " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_WHERE2 & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & IIf(sql_GROUP_BY_cols = "", "''", sql_GROUP_BY_cols) & Replace(COLUMN_NAMEs_appended, ",", ",ASTSRPTX" & ".")

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
        & "(" & G1thru9 & COLUMN_NAMEs_appended & DATA_COLUMN_NAMEs & ")" & vbCrLf _
        & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()



    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = String.Format("Annual Turns Calculated Using the Avg of {0} Shpmts & Invty for {1} Previous Months", IIf(Absx1.optFor("GN").Value = "G", "Gross", "Net"), Absx1.numFor("MOS").Value)

        CR_params.Add("ACTIVE_ONLY", IIf(Absx1.chkFor("ACTIVE_ONLY").Checked, "1", "0"))
        'CR_params.Add("MOS", Absx1.numFor("MOS").Value & "")
        'CR_params.Add("GN", Absx1.optFor("GN").Value)

        For i As Integer = 1 To 13
            CR_params.Add("MD" & Format(i, "00"), md(i))
        Next

        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Build_Report_File_Post_Process()

        Call ASCMAIN1.Progress("Report Calculations")

        Dim PA As Integer = ASCMAIN1.Period_Diff(RYP, ASCMAIN1.CYP)

        Dim YP_start As String = ""

        Dim V(,) As Double
        For Each rowASTSRPT1 As DataRow In dst.Tables("ASTSRPT1").Rows
            Dim ITEM_CODE As String = rowASTSRPT1.Item("ITEM_CODE")
            'If ITEM_CODE = "84215065" And ASCMAIN1.Running_in_VS Then Stop

            Dim C() As Decimal = ICCMAIN1.ITEM_Cost_History(ITEM_CODE, 36, YP_start)

            ReDim V(2, 24)
            For i = 0 To 24
                V(1, i) = Val(rowASTSRPT1.Item("B" & Format$(i, "00")) & "")
                V(2, i) = Val(rowASTSRPT1.Item("U" & Format$(i, "00")) & "")
            Next i

            Dim Total_Beginning As Int32 = 0
            Dim Total_Used As Int32 = 0

            Dim k As Integer = 0
            For i = 12 To 0 Step -1
                k = 0
                Total_Beginning = 0
                Total_Used = 0
                For j = MOS To 1 Step -1
                    If V(1, i + j - 1) <> 0 And V(2, i + j - 1) <> 0 Then
                        k = k + 1
                        Total_Beginning += V(1, i + j - 1)
                        Total_Used += V(2, i + j - 1)
                    End If
                Next j
                Dim z As String = Format$(i, "00")

                If Total_Beginning = 0 Then
                    rowASTSRPT1.Item("T" & z) = 0
                Else
                    rowASTSRPT1.Item("T" & z) = 12 * Total_Used / Total_Beginning
                End If
                If k = 0 Then
                    rowASTSRPT1.Item("CU" & z) = 0
                    rowASTSRPT1.Item("CB" & z) = 0
                Else
                    rowASTSRPT1.Item("CU" & z) = 12 * Total_Used * c(i + PA) / k
                    rowASTSRPT1.Item("CB" & z) = Total_Beginning * c(i + PA) / k
                End If
            Next i

            rowASTSRPT1.Item("AVG_MOS") = k
            If Total_Used = 0 Then
                rowASTSRPT1.Item("T00") = 0
            Else
                rowASTSRPT1.Item("T00") = V(1, 0) / (Total_Used / k)
                rowASTSRPT1.Item("CB00") = V(1, 0) * C(0 + PA)
            End If

        Next

        ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC, ITEM_COST_STD from ICTITEM1 " _
        & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & ASTSRPT1 & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "ICTITEM1", 1))

    End Sub
End Class