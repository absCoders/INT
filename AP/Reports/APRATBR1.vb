Public Class APRATBR1

    Dim APTINVH1 As String
    Dim DTES(4) As Date
    Dim ODTES(4) As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Call Range_Events(grpINV_DATE_RANGE)

        Call Get_PARM("GLTPARM1")
        Call Get_PARM("APTPARM1")

        Absx1.numFor("DAYS1").Value = ROWs("APTPARM1").Item("AP_PARM_AGE_DAYS_1")
        Absx1.numFor("DAYS2").Value = ROWs("APTPARM1").Item("AP_PARM_AGE_DAYS_2")
        Absx1.numFor("DAYS3").Value = ROWs("APTPARM1").Item("AP_PARM_AGE_DAYS_3")
        Absx1.numFor("DAYS4").Value = ROWs("APTPARM1").Item("AP_PARM_AGE_DAYS_4")

        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, -1, -1)

    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("VEND_CODE")

                If Absx1.optFor("OPTDS").Value = "D" And Val(rowASTDSQLA("SEQUENCE") & "") = 0 Then
                    EMsg &= "You Must Sort by Vendor when reporting AP Item Details"
                End If
        End Select
    End Sub

    Protected Overrides Sub Build_Workfile()

        ' Prepare Working Variables

        ' make sure that 
        'an aging date was specified

        ASCMAIN1.Progress("Building Open AP Work File")




        Dim APTOPENI As String = ""
        If optAPITEMS.Value = "C" Then
            Page0.Add("Using Currently Open AP Items")
            APTOPENI = "APTINVH1"
        Else
            Dim RYP As String = Absx1.cmbFor("RYP").Value

            Page0.Add("Using AP Items open at End of " & Absx1.cmbFor("RYP").Text)
            APTOPENI = ASCMAIN1.Temp_Table("Select * from APTINVH1 where ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & APTOPENI & " Add Primary Key (VOUCHER_NO)")

            Dim SQL_ARTOPEN1 As String = "Select * from APTINVH1 where (VOUCHER_NO) in " _
                & " (Select DETL_CTL_NO from GLTCREC3" _
                & " where CREC_TYPE_CODE = 'AP' and OPS_YYYYPP = '" & RYP & "')"
            ASCDATA1.ExecuteSQL("Insert into " & APTOPENI & " " & SQL_ARTOPEN1)

            ASCMAIN1.sql = "" _
                & "Begin " & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf _
                & "  Select DETL_CTL_NO VOUCHER_NO, CREC_AMT from GLTCREC3" & vbCrLf _
                & "   where CREC_TYPE_CODE = 'AP' and OPS_YYYYPP = '" & RYP & "';" & vbCrLf _
                & " Begin " & vbCrLf _
                & "   For R1 in C1 Loop" & vbCrLf _
                & "    Update " & APTOPENI & " Set INV_BALANCE = R1.CREC_AMT, INV_STATUS = 'O'" & vbCrLf _
                & "     where VOUCHER_NO = R1.VOUCHER_NO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()
        End If





        Dim sql As String

        sql = "Select APTINVH1.VEND_CODE" _
        & ", SUM (APTINVH1.INV_BALANCE) INV_BALANCE " _
        & ", SUM (DECODE(APTINVH1.INV_TYPE,'I',APTINVH1.INV_BALANCE,0)) INV_BALANCE_I " _
        & ", SUM (DECODE(APTINVH1.INV_TYPE,'B',APTINVH1.INV_BALANCE,0)) INV_BALANCE_B " _
        & ", SUM (DECODE(APTINVH1.INV_TYPE,'D',APTINVH1.INV_BALANCE,0)) INV_BALANCE_D " _
        & ", SUM (DECODE(APTINVH1.INV_TYPE,'R',APTINVH1.INV_BALANCE,0)) INV_BALANCE_R " _
        & ", SUM (DECODE(APTINVH1.INV_TYPE,'C',APTINVH1.INV_BALANCE,0)) INV_BALANCE_C " _
        & ", SUM (DECODE(APTINVH1.INV_TYPE,'A',APTINVH1.INV_BALANCE,0)) INV_BALANCE_A " _
        & $" from {APTOPENI} APTINVH1 where INV_STATUS in ('O','H')  group by APTINVH1.VEND_CODE"
        '& " group by APTINVH1.VEND_CODE"
        Dim APTINVH1_BAL As String = ASCMAIN1.Temp_Table(sql)
        ASCDATA1.ExecuteSQL("Alter Table " & APTINVH1_BAL & " Add Primary Key (VEND_CODE)")

        Dim GL(2) As String
        Dim S As Integer = 0
        If Absx1.optFor("OPTFP").Value = "F" Then
            Page0.Add("Future Aging")
            GL(1) = " > "
            GL(2) = " <= "
            S = 1
        Else
            Page0.Add("Past Due Aging")
            GL(1) = " < "
            GL(2) = " >= "
            S = -1
        End If

        If Absx1.optFor("OPTDP").Value = "D" Then
            Page0.Add("Age using Days from Base Date")
        Else
            Page0.Add("Age based on Period Ending Dates")
        End If

        For I As Integer = 0 To 4
            If I = 0 Then
                If Absx1.optFor("OPTDP").Value = "D" Then
                    DTES(I) = Absx1.dteFor("AGING_DATE").Value
                Else
                    ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"
                    DTES(I) = DateValue(ASCDATA1.GetDataValue).AddDays(-1 * S)
                End If
            Else
                If Absx1.optFor("OPTDP").Value = "D" Then
                    Dim NUMDAYS As Integer = Absx1.numFor("DAYS" + CStr(I)).Value
                    DTES(I) = DTES(0).AddDays(S * NUMDAYS)
                Else
                    Dim AYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, S * I)
                    ASCMAIN1.sql = "Select PRD_END_DATE from GLTPARM2 where OPS_YYYYPP = '" & AYP & "'"
                    DTES(I) = DateValue(ASCDATA1.GetDataValue).AddDays(-1 * S)
                End If
            End If
            ODTES(I) = "'" & Format(DTES(I), "dd-MMM-yyyy") & "'"
        Next

        Dim sqld As String = "APTINVH1." & IIf(Absx1.optFor("OPTID").Value = "I", "INV_DATE", "INV_DUE_DATE")
        sql = "Select APTINVH1.* " & vbCr
        For i As Integer = 0 To 5
            Dim sqle As String = " THEN APTINVH1.INV_BALANCE ELSE 0 END "
            Dim sqla As String = ""
            If i = 0 Then
                sqla = "Case When " & sqld & GL(2) & ODTES(0) & sqle
            ElseIf i = 5 Then
                sqla = "Case When " & sqld & GL(1) & ODTES(i - 1) & sqle
            Else
                sqla = "Case When " & sqld & GL(1) & ODTES(i - 1) & " and " & sqld & GL(2) & ODTES(i) & sqle
            End If
            If Absx1.chkFor("AGE_CHARGEBACKS").Checked Then
                sql = sql & ", " & sqla & " INV_BALANCE_" & CStr(i) & vbCr
            Else
                sql = sql & ", Decode(APTINVH1.INV_TYPE,'B',0," & sqla & ") INV_BALANCE_" & CStr(i) & vbCr
            End If
        Next

        If Absx1.chkFor("AGE_CHARGEBACKS").Checked Then
            Page0.Add("Chargebacks ARE reflected in the Aging Columns")
        Else
            Page0.Add("Chargebacks are NOT reflected in the Aging Columns")
        End If

        sql = sql _
        & ", DECODE (APTINVH1.INV_TYPE,'B',APTINVH1.INV_BALANCE,0) CHARGEBACKS " & vbCr _
        & ", CASE WHEN APTINVH1.INV_BALANCE < 0 THEN APTINVH1.INV_BALANCE ELSE 0 END CREDITS" & vbCr _
        & ", TO_DATE(" & ODTES(0) & ") - " & IIf(Absx1.optFor("OPTID").Value = "I", "INV_DATE", "INV_DUE_DATE") & " DAYS_OLD" & vbCr _
        & $" from {APTOPENI} APTINVH1, APTVEND1, " & APTINVH1_BAL & " APTINVH1_BAL" & vbCr

        sql = sql & " where APTVEND1.VEND_CODE = APTINVH1.VEND_CODE" & vbCr
        sql = sql & " and APTINVH1_BAL.VEND_CODE = APTINVH1.VEND_CODE" & vbCr

        Select Case Absx1.optFor("OPTBALANCE").Value
            Case "M"
                sql = sql & " and APTINVH1_BAL.INV_BALANCE > " & CStr(Absx1.numFor("BALANCE").Value)
                Page0.Add("Vendors with Balance > " & CStr(Absx1.numFor("BALANCE").Value))
            Case "L"
                sql = sql & " and APTINVH1_BAL.INV_BALANCE < " & CStr(Absx1.numFor("BALANCE").Value)
                Page0.Add("Vendors with Balance < " & CStr(Absx1.numFor("BALANCE").Value))
            Case "H"
                sql = sql & " and NVL(APTVEND1.VEND_ON_HOLD,'0') = '1'"
                Page0.Add("Vendors On Hold")
            Case "W"
                sql = sql & " and APTINVH1_BAL.INV_BALANCE_A <> 0"
                Page0.Add("Vendors with non-zero Advance Balance")
        End Select

        If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
            Dim z As String = Format(Absx1.dteFor("INV_DATE_F").Value, "dd-MMM-yyyy")
            sql = sql & " and APTINVH1.INV_DATE >= '" & z & "'"
            Page0.Add("Invoices dated >= " & z)
        End If
        If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
            Dim z As String = Format(Absx1.dteFor("INV_DATE_L").Value, "dd-MMM-yyyy")
            sql = sql & " and APTINVH1.INV_DATE <= '" & z & "'"
            Page0.Add("Invoices dated <= " & z)
        End If
        If Absx1.optFor("OPTAOH").Value = "O" Or Absx1.optFor("OPTAOH").Value = "H" Then
            sql = sql & " and APTINVH1.INV_STATUS = '" & Absx1.optFor("OPTAOH").Value & "'"
            If Absx1.optFor("OPTAOH").Value = "H" Then
                Page0.Add("Invoices On Hold for Payment")
            Else
                Page0.Add("Invoices OK to Pay")
            End If
        Else
            sql = sql & " and APTINVH1.INV_STATUS in ('O','H')"
        End If
        If Absx1.optFor("OPTAP").Value = "P" Then
            sql = sql & " and APTINVH1.INV_DUE_DATE > '" & Format(Absx1.dteFor("AGING_DATE").Value, "dd-MMM-yyyy") & "'"
            Page0.Add("Past Due Invoices Only")
        End If

        Dim not_all_types_selected As Boolean = False
        Dim INV_TYPEs As String = ""
        For Each INV_TYPE As String In New String() {"I", "B", "D", "R", "C", "A"}
            If Absx1.chkFor("INV_TYPE_" & INV_TYPE).Checked Then
                INV_TYPEs &= ",'" & INV_TYPE & "'"
            Else
                not_all_types_selected = True
            End If
        Next
        If not_all_types_selected Then
            sql = sql & " and APTINVH1.INV_TYPE in (" & Mid(INV_TYPEs, 2) & ")"
        End If

        sql = sql & SQL_in("VEND_CODE", "APTINVH1.VEND_CODE")

        APTINVH1 = ASCMAIN1.Temp_Table(sql)
        ASCMAIN1.sql = "Select * from " & APTINVH1
        dst.Tables.Add(ASCDATA1.GetDataTable("", "APTATBR1", 1))

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Call MyBase.Get_SQL("*")

        sql = "Select " & sql_SELECT_cols & vbCr _
        & ", APTINVH1.VOUCHER_NO" & vbCr _
        & ", APTINVH1.INV_BALANCE" & vbCr _
        & ", APTINVH1.INV_BALANCE_0" & vbCr _
        & ", APTINVH1.INV_BALANCE_1" & vbCr _
        & ", APTINVH1.INV_BALANCE_2" & vbCr _
        & ", APTINVH1.INV_BALANCE_3" & vbCr _
        & ", APTINVH1.INV_BALANCE_4" & vbCr _
        & ", APTINVH1.INV_BALANCE_5" & vbCr _
        & ", APTINVH1.CREDITS" & vbCr _
        & ", APTINVH1.CHARGEBACKS" & vbCr
        ' TAKE A LOOK AT ?REPLACE(replace(ASTSRPT1_sum_columns,"SUM(",""),")","")
        ' MIGHT BE ABLE TO DO AWAY WITH THE HARD CODE
        ' ALSO - COLUMN_NAMEs_appended MAYBE ABLE TO REPLACE VOUCHER_NO

        sql = sql & " from " & APTINVH1 & " APTINVH1 " & sql_TABLE_NAMEs & vbCr
        sql = sql & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCr
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")


    End Sub

    Public Overrides Sub Print_Report()

        Dim i As Integer
        Dim z As String

        Dim SUBT As String = ""

        CR_params.Add("FP", Absx1.optFor("OPTFP").Value)
        If Absx1.optFor("OPTFP").Value = "F" Then
            SUBT = "Future"
        Else
            SUBT = "Past Due"
        End If
        If Absx1.optFor("OPTID").Value = "I" Then
            SUBT &= " Aging by Invoice Date"
        Else
            SUBT &= " Aging by Due Date"
        End If
        If Absx1.optFor("OPTAP").Value = "P" Then
            SUBT = SUBT & ", Past Due Invoices"
        End If


        If optAPITEMS.Value = "P" Then
            SUBT &= ", Balance as of " & Absx1.cmbFor("RYP").Text
        Else
            If Absx1.optFor("OPTFP").Value = "F" Then
                SUBT = SUBT & ", Future Aging"
            End If
            SUBT = SUBT & ", as of " & Format(DTES(0), "MM/dd/yyyy")
            'SUBT &= ", as of " & Format(System.DateTime.Now, "MM/dd/yyyy")
        End If


        If Not Absx1.chkFor("CHKINV_DATE_F").Checked _
        Or Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
            SUBT = SUBT & ", "
            SUBT = SUBT & "Showing A/P Items Dated"
            If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
                SUBT = SUBT & " from " & Format(Absx1.dteFor("INV_DATE_F").Value, "MM/dd/yyyy")
            End If
            If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                SUBT = SUBT & " thru " & Format(Absx1.dteFor("INV_DATE_L").Value, "MM/dd/yyyy")
            End If
        End If
        Select Case Absx1.optFor("OPTBALANCE").Value
            Case "H"
                SUBT = SUBT & ", Vendors On Hold Only"
            Case "M"
                SUBT = SUBT & ", Balances > " & Format(Absx1.numFor("BALANCE").Value, "$0")
            Case "L"
                SUBT = SUBT & ", Balances < " & Format(Absx1.numFor("BALANCE").Value, "$0")
            Case "W"
                SUBT = SUBT & ", Vendors w/Advances Only"
        End Select

        For i = 1 To 5
            If i = 5 Then
                z = "DATE" & Format$(i, "0")
                Dim ZZ As String = ""
                If Absx1.optFor("OPTFP").Value = "F" Then
                    ZZ = ">= " & Format(DTES(i - 1).AddDays(1), "MM/dd")
                Else
                    ZZ = "<= " & Format(DTES(i - 1).AddDays(-1), "MM/dd")
                End If
                CR_params.Add(z, ZZ)
            Else
                z = "DAYS" & CStr(i)
                CR_params.Add(z, Absx1.numFor(z).Value)
                z = "DATE" & Format$(i, "0")
                If Absx1.optFor("OPTFP").Value = "F" Then
                    CR_params.Add(z, Format(DTES(i - 1).AddDays(1), "MM/dd") & "-" & Format(DTES(i), "MM/dd"))
                Else
                    CR_params.Add(z, Format(DTES(i), "MM/dd") & "-" & Format(DTES(i - 1).AddDays(-1), "MM/dd"))
                End If
            End If
        Next i

        If Absx1.optFor("OPTFP").Value = "F" Then
            z = "<=" & Format(DTES(0), "MM/dd")
        Else
            z = ">=" & Format(DTES(0), "MM/dd")
        End If

        CR_params.Add("DATE0", z)

        If Absx1.optFor("OPTID").Value = "I" Then
            z = "Days Old"
        Else
            z = "Past Due"
        End If
        CR_params.Add("AGEBY", z)
        CR_params.Add("DTL", Absx1.optFor("OPTDS").Value)

        Generate_Report(RPT, , SUBT)
    End Sub

    Private Sub optDP_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optDP.ValueChanged
        grpAgingDays.Visible = (optDP.Value = "D")
    End Sub


    Private Sub optAPITEMS_ValueChanged(sender As Object, e As EventArgs) Handles optAPITEMS.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Absx1.cmbFor("RYP").Visible = (optAPITEMS.Value = "P")
        If optAPITEMS.Value = "P" Then
            Set_Aging_Date()
        End If
    End Sub

    Sub Set_Aging_Date()
        Dim YP As String = Absx1.cmbFor("RYP").Value
        'YP = Mid(YP, 1, 4) & Mid(YP, 6, 2)
        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", YP)
        Absx1.dteFor("AGING_DATE").Value = rowGLTPARM2.Item("PRD_END_DATE")
    End Sub

    Private Sub cmbRYP_ValueChanged(sender As Object, e As EventArgs) Handles cmbRYP.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        If optAPITEMS.Value = "P" Then
            Set_Aging_Date()
        End If
    End Sub

    Private Sub optBALANCE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optBALANCE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        ' this routine is hit before this form is loaded
        ' 1st time thru me.name = ASFSRPTM, after that it is APRATBR1
        Absx1.numFor("BALANCE").Enabled = (optBALANCE.Value = "M" Or optBALANCE.Value = "L")
    End Sub
End Class