Public Class SORPCKR1

    Dim SOTPICKX As String
    Dim OPTSORT As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Range_Events(grpPICK_RELEASED)
        Range_Events(grpORDR_SHIP_DATE)
        Range_Events(grpORDR_CANCEL_DATE)

        Get_PARM("SOTPARM1")
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()
 
        If ASCMAIN1.DBS_COMPANY = "VAN" And ASCMAIN1.DBS_SERVER = "VAN" Then
            TAC.WHCMAIN1.Update_ADS_SOTSHIP1()
        End If

        ASCMAIN1.Progress("Building Work File")

        ASCMAIN1.sql = "SELECT SOTPICK1.*" & vbCrLf _
            & ", SOTORDR1.ORDR_CUST_PO, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE" & vbCrLf _
            & ", SOTORDR1.WHSE_CODE_TO, SOTORDR1.ORDR_TYPE_CODE" & vbCrLf _
            & " from SOTPICK1, SOTORDR0, SOTSHIP1, SOTORDR1" & vbCrLf _
            & " where SOTORDR0.ORDR_GROUP_NO = SOTSHIP1.ORDR_GROUP_NO" & vbCrLf _
            & "   and SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
            & "   and SOTSHIP1.SHIP_STATUS = 'P'" & vbCrLf _
            & Get_Dates()
        ASCMAIN1.sql &= SQL_in("CUST_CODE", "SOTORDR0.CUST_CODE")
        ASCMAIN1.sql &= SQL_in("PICK_BATCH_NO", "SOTPICK1.PICK_BATCH_NO")
        ASCMAIN1.sql &= SQL_in("WHSE_CODE", "SOTSHIP1.WHSE_CODE")

        OPTSORT = Absx1.optFor("OPTSORT").Value

        SOTPICKX = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTPICKX & " Add PICK_QTY NUMBER (8,0)")
        ASCDATA1.ExecuteSQL("Update " & SOTPICKX & " SOTPICKX Set PICK_QTY = (Select Sum (PICK_QTY) from SOTPICK2 where PICK_NO = SOTPICKX.PICK_NO)")

        ASCMAIN1.sql = "Select * from " & SOTPICKX
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTPICKX", 1))

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources
        ASCMAIN1.sql = "Select SOTSHIP1.* from SOTSHIP1" _
            & " where SOTSHIP1.SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTPICKX & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTSHIP1", 1))
        dst.Tables("SOTSHIP1").Columns.Add("SHIP_WINDOW_CHG")
        dst.Tables("SOTSHIP1").Columns.Add("SHIP_UNITS_CHG")

        ASCMAIN1.sql = "Select Distinct SOTSHIP3.SHIP_BOL_NO from SOTSHIP3" & vbCrLf _
            & " where SOTSHIP3.SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTPICKX & ")" & vbCrLf _
            & "   and (NVL(ORDR_SHIP_DATE_OLD,TRUNC(SYSDATE)) <> NVL(ORDR_SHIP_DATE_NEW,TRUNC(SYSDATE))" & vbCrLf _
            & "    or" & vbCrLf _
            & "        NVL(ORDR_CANCEL_DATE_OLD,TRUNC(SYSDATE)) <> NVL(ORDR_CANCEL_DATE_NEW,TRUNC(SYSDATE)))"

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & ""
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                rowSOTSHIP1.Item("SHIP_WINDOW_CHG") = "1"
            Next
        End If


        ASCMAIN1.sql = "Select Distinct SOTSHIP3.SHIP_BOL_NO from SOTSHIP3,SOTSHIP6" & vbCrLf _
            & " where SOTSHIP3.SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTPICKX & ")" & vbCrLf _
            & "   and SOTSHIP6.SHIP_CHGREQ_NO = SOTSHIP3.SHIP_CHGREQ_NO" & vbCrLf

        If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim SHIP_BOL_NO As String = row.Item("SHIP_BOL_NO") & ""
                Dim rowSOTSHIP1 As DataRow = dst.Tables("SOTSHIP1").Rows.Find(SHIP_BOL_NO)
                rowSOTSHIP1.Item("SHIP_UNITS_CHG") = "1"
            Next
        End If


        ASCMAIN1.sql = "Select SOTORDR0.*, ARTCUST1.CUST_NAME from SOTORDR0, ARTCUST1" _
            & " where ARTCUST1.CUST_CODE = SOTORDR0.CUST_CODE and SOTORDR0.ORDR_GROUP_NO in " _
            & " (Select Distinct ORDR_GROUP_NO from " & SOTPICKX & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable("", "SOTORDR0", 1))
        dst.Tables("SOTORDR0").Columns.Add("SORTBY")

        For Each rowSOTORDR0 As DataRow In dst.Tables("SOTORDR0").Select("")
            Dim ORDR_GROUP_NO As String = rowSOTORDR0.Item("ORDR_GROUP_NO")

            If OPTSORT = "G" Then
                rowSOTORDR0.Item("SORTBY") = rowSOTORDR0.Item("ORDR_GROUP_NO")
            ElseIf OPTSORT = "P" Then
                rowSOTORDR0.Item("SORTBY") = rowSOTORDR0.Item("ORDR_CUST_PO")
            ElseIf OPTSORT = "C" Then
                rowSOTORDR0.Item("SORTBY") = Format(rowSOTORDR0.Item("ORDR_CANCEL_DATE"), "yyyyMMdd")
            End If

            If ASCMAIN1.DBS_COMPANY = "VAN" Or ASCMAIN1.DBS_SERVER = "VAN" Then
            Else
                If rowSOTORDR0.Item("ORDR_ORIG_SHIP_DATE") & "" <> "" _
                    And rowSOTORDR0.Item("ORDR_ORIG_CANCEL_DATE") & "" <> "" Then
                    If Format(rowSOTORDR0.Item("ORDR_ORIG_SHIP_DATE") & "", "yyyyMMdd") <> Format(rowSOTORDR0.Item("ORDR_SHIP_DATE") & "", "yyyyMMdd") _
                    Or Format(rowSOTORDR0.Item("ORDR_ORIG_CANCEL_DATE") & "", "yyyyMMdd") <> Format(rowSOTORDR0.Item("ORDR_CANCEL_DATE") & "", "yyyyMMdd") Then

                        For Each rowSOTSHIP1 As DataRow In dst.Tables("SOTSHIP1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'")
                            rowSOTSHIP1.Item("SHIP_WINDOW_CHG") = "1"
                        Next
                    End If
                End If
            End If
        Next

    End Sub

    Function Get_Dates() As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"PICK_RELEASED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            Dim TABLE_NAME As String = "SOTORDR0"
            If COLUMN_NAME = "PICK_RELEASED" Then TABLE_NAME = "SOTPICK1"
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                sql = sql & " and " & TABLE_NAME & "." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
            End If
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                sql = sql & " and " & TABLE_NAME & "." & COLUMN_NAME & " <= '" & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "dd-MMM-yyyy") & "'"
            End If
        Next
        Return sql
    End Function

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        'ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = 0")
        'ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = TRUNC(100 * ORDR_AMT / ORDR_QTY) / 100 where ORDR_QTY <> 0")
    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = ""
        Page0.Add("Report Detail: " & IIf(Absx1.chkFor("CHKDETAILED").Checked, "Yes", "No"))
        For Each COLUMN_NAME As String In New String() {"PICK_RELEASED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            Dim Z As String = Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text & ":"
            Dim real_date_selected As Boolean = False
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                Z &= " from First"
            Else
                Z &= " from " & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "MM/dd/yyyy")
                real_date_selected = True
            End If
            If Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                Z &= " to Last"
            Else
                Z &= " to " & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "MM/dd/yyyy")
                real_date_selected = True
            End If
            If real_date_selected Then
                SUBT &= ", " & Z
            End If
            Page0.Add(Z)
        Next
        Page0.Add("Sorted by " & Absx1.optFor("OPTSORT").Text)
        SUBT &= ", " & "Sorted by " & Absx1.optFor("OPTSORT").Text
        SUBT = Mid(SUBT, 3)

        CR_params.Add("DETAILED", IIf(Absx1.chkFor("CHKDETAILED").Checked, "1", "0"))

        Generate_Report(RPT, RPT_TITLE, SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                'Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("STYLE_CODE")

                'If Absx1.optFor("OPTDTL").Value = "2" And Val(rowASTDSQLA("SEQUENCE") & "") <> 0 Then
                '    EMsg &= "You Must NOT Sort by Style when showing Details"
                'End If
        End Select
    End Sub

    Overrides Sub Update_Record()

    End Sub

End Class