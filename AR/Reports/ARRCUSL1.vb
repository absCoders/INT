Public Class ARRCUSL1

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Range_Events(grpINV_DATE_RANGE)
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -12, 0, -1)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -12, 0, -1)
        Set_cmbYP("RYP", ASCMAIN1.CYP, -12, 0, -1)

        grpPERIOD_RANGE.Left = grpINV_DATE_RANGE.Left
        grpPERIOD_RANGE.Top = grpINV_DATE_RANGE.Top
        grpPeriod.Left = grpINV_DATE_RANGE.Left
        grpPeriod.Top = grpINV_DATE_RANGE.Top

        Select Case MENU_ITEM_OBJECT
            Case "ARRCUSLC"
                Absx1.optFor("OPTSELECT").Value = "M"
            Case Else
                Absx1.optFor("OPTSELECT").Value = "C"
        End Select


    End Sub

    Protected Overrides Sub Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim sql As String
        Dim sqlw As String = String.Empty
        Dim sqlwSREP As String = String.Empty
        Dim sqlARTCUST2 As String = String.Empty

        sqlwSREP = SQL_in("SREP_CODE", "ARTCUST1.SREP_CODE")

        sqlw &= SQL_in("CUST_CODE", "ARTCUST1.CUST_CODE")
        sqlw &= SQL_in("TRADE_CLASS_CODE", "ARTCUST1.TRADE_CLASS_CODE")
        sqlw &= SQL_in("CUST_CLASS_CODE", "ARTCUST1.CUST_CLASS_CODE")

        If Absx1.optFor("OPTSELECT").Value = "C" OrElse Absx1.optFor("OPTSELECT").Value = "M" Then
            'select from customer master and their locations
            sql = "Select DISTINCT ARTCUST1.CUST_CODE"
            sql &= " from ARTCUST1, ARTCUST6 "
            sql &= " where ARTCUST1.CUST_CODE = ARTCUST6.CUST_CODE (+)"
            sql &= " and ARTCUST1.CUST_CODE <> 'CONSUMER'"
            sql &= sqlw

            If Not Absx1.chkFor("CHKINV_DATE_F").Checked Then
                Dim z As String = Format(Absx1.dteFor("INV_DATE_F").Value, "dd-MMM-yyyy")
                sql &= " and ARTCUST6.CUST_LAST_INV_DATE >= '" & z & "'"
            End If
            If Not Absx1.chkFor("CHKINV_DATE_L").Checked Then
                Dim z As String = Format(Absx1.dteFor("INV_DATE_L").Value, "dd-MMM-yyyy")
                sql &= " and ARTCUST6.CUST_LAST_INV_DATE <= '" & z & "'"
            End If

            ' Allow for ARTCUST2.SREP_CODE
            If sqlwSREP.Length > 0 Then
                sqlARTCUST2 = "Select Distinct ARTCUST2.CUST_CODE"
                sqlARTCUST2 &= " from ARTCUST2"
                sqlARTCUST2 &= " where CUST_CODE IN (" & sql & ")"
                sqlARTCUST2 &= sqlwSREP.Replace("ARTCUST1", "ARTCUST2")

                sqlwSREP = sqlwSREP.Substring(4)
                sqlwSREP = " and ( " & sqlwSREP & " or " & sqlwSREP.Replace("SREP_CODE", "SREP2_CODE") & " )"
                sql &= sqlwSREP
            End If

            If sqlARTCUST2.Length > 0 Then
                sql &= " Union " & sqlARTCUST2
            End If

            Dim ARTCUST1 As String = ASCMAIN1.Temp_Table(sql)
            CreateARTCUSL1(ARTCUST1)

        ElseIf Absx1.optFor("OPTSELECT").Value = "S" Then
            ' Select customers and locations from sales history 
            sql = "Select Distinct SOTINVH1.CUST_CODE, SOTINVH1.CUST_STORE_NO "
            sql &= " from SOTINVH1" & IIf(sqlwSREP.Length = 0, ", ARTCUST1", "")
            sql &= " where SOTINVH1.CUST_CODE <> 'CONSUMER'"
            sql &= " and SOTINVH1.ORDR_YYYYPP_UPDATED >= '" & RYP0 & "'"
            sql &= " and SOTINVH1.ORDR_YYYYPP_UPDATED <= '" & RYP1 & "'"

            sql &= sqlw

            If sqlwSREP.Length = 0 Then
                sql &= " and SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE"
            Else
                sql &= " and SOTINVH1.CUST_CODE IN"
                sql &= " ( "
                sql &= "    SELECT DISTINCT CUST_CODE FROM ARTCUST1 WHERE " & sqlwSREP.Substring(4)
                sql &= " Union "
                sql &= "    SELECT DISTINCT CUST_CODE FROM ARTCUST1 WHERE " & sqlwSREP.Substring(4).Replace("SREP_CODE", "SREP2_CODE")
                sql &= " Union "
                sql &= "   SELECT DISTINCT CUST_CODE FROM ARTCUST2 WHERE " & sqlwSREP.Substring(4).Replace("ARTCUST1", "ARTCUST2")
                sql &= " ) "
            End If

            Dim ARTCUST1 As String = ASCMAIN1.Temp_Table(sql)
            CreateARTCUSL1(ARTCUST1)

        ElseIf Absx1.optFor("OPTSELECT").Value = "R" Then
            'select customers with sales history in the last two months only
            sql = " SELECT * FROM ("
            sql &= " SELECT SOTINVH1.CUST_CODE, SUM(INV_SALES)" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -11) & "',INV_SALES)),0) MO_01" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -10) & "',INV_SALES)),0) MO_02" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -9) & "',INV_SALES)),0) MO_03" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -8) & "',INV_SALES)),0) MO_04" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -7) & "',INV_SALES)),0) MO_05" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -6) & "',INV_SALES)),0) MO_06" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -5) & "',INV_SALES)),0) MO_07" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -4) & "',INV_SALES)),0) MO_08" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -3) & "',INV_SALES)),0) MO_09" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -2) & "',INV_SALES)),0) MO_10" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & ASCMAIN1.Period_Calc(RYP, -1) & "',INV_SALES)),0) MO_11" & vbCr
            sql &= " , NVL(SUM(DECODE(ORDR_YYYYPP_UPDATED,'" & RYP & "',INV_SALES)),0) MO_12" & vbCr
            'sql &= " FROM SOTINVH1, ARTCUST1" & vbCr
            sql &= " from SOTINVH1" & IIf(sqlwSREP.Length = 0, ", ARTCUST1", "")
            sql &= " WHERE ORDR_YYYYPP_UPDATED BETWEEN '" & ASCMAIN1.Period_Calc(RYP, -11) & "' AND '" & RYP & "'" & vbCr
            'sql &= " and SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE"
            sql &= " AND INV_SALES <> 0" & vbCr

            sql &= sqlw

            If sqlwSREP.Length = 0 Then
                sql &= " and SOTINVH1.CUST_CODE = ARTCUST1.CUST_CODE"
            Else
                sql &= " and SOTINVH1.CUST_CODE IN"
                sql &= " ( "
                sql &= "    SELECT DISTINCT CUST_CODE FROM ARTCUST1 WHERE " & sqlwSREP.Substring(4)
                sql &= " Union "
                sql &= "    SELECT DISTINCT CUST_CODE FROM ARTCUST1 WHERE " & sqlwSREP.Substring(4).Replace("SREP_CODE", "SREP2_CODE")
                sql &= " Union "
                sql &= "   SELECT DISTINCT CUST_CODE FROM ARTCUST2 WHERE " & sqlwSREP.Substring(4).Replace("ARTCUST1", "ARTCUST2")
                sql &= " ) "
            End If

            sql &= " GROUP BY SOTINVH1.CUST_CODE)" & vbCr
            sql &= " where MO_11 > 0 AND MO_12 > 0" & vbCr
            sql &= " AND MO_01 <= 0" & vbCr
            sql &= " AND MO_02 <= 0" & vbCr
            sql &= " AND MO_03 <= 0" & vbCr
            sql &= " AND MO_04 <= 0" & vbCr
            sql &= " AND MO_05 <= 0" & vbCr
            sql &= " AND MO_06 <= 0" & vbCr
            sql &= " AND MO_07 <= 0" & vbCr
            sql &= " AND MO_08 <= 0" & vbCr
            sql &= " AND MO_09 <= 0" & vbCr
            sql &= " AND MO_10 <= 0" & vbCr

            Dim ARTCUST1 As String = ASCMAIN1.Temp_Table(sql)
            CreateARTCUSL1(ARTCUST1)

        End If

        sqlwSREP = SQL_in("SREP_CODE", "SOTSREP1.SREP_CODE")

        If sqlwSREP.Length = 0 Then
            For Each ROW As DataRow In dst.Tables("ARTCUSL1").Select("SREP_CODE = ''")
                ROW.Item("SREP_CODE") = "?"
            Next
            Call Get_WKCodes("ARTCUSL1", "SREP_CODE", "SOTSREP1")
        Else
            If Not dst.Tables.Contains("SOTSREP1") Then
                ASCMAIN1.sql = "Select * from SOTSREP1 WHERE " & sqlwSREP.Substring(4)
                dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSREP1", 1))
            End If
        End If


        For Each rowSOTSREP1 As DataRow In dst.Tables("SOTSREP1").Select("")
            rowSOTSREP1.Item("SREP_NAME") = StrConv(rowSOTSREP1.Item("SREP_NAME") & String.Empty, VbStrConv.ProperCase)
        Next

        Check_if_Empty("ARTCUSL1")

    End Sub

    Private Sub CreateARTCUSL1(ByVal ARTCUST1 As String)
        With dst

            Dim sqlw As String = SQL_in("SREP_CODE", "ARTCUST1.SREP_CODE")

            sql = " Select CUST_CODE || '-000000' CUST, SREP_CODE , InitCap(CUST_NAME) NAME , InitCap(CUST_ADDR1) ADDR1, InitCap(CUST_ADDR2) ADDR2, InitCap(CUST_ADDR3) ADDR3, "
            sql &= " InitCap(CUST_CITY) CITY, CUST_STATE STATE, CUST_ZIP_CODE ZIP_CODE, InitCap(CUST_CONTACT) CONTACT"
            sql &= " from ARTCUST1 where CUST_CODE in (Select CUST_CODE from " & ARTCUST1 & ")"
            sql &= sqlw

            sql &= " Union "

            sql &= " Select CUST_CODE || '-000000' CUST, SREP2_CODE SREP_CODE, InitCap(CUST_NAME) NAME , InitCap(CUST_ADDR1) ADDR1, InitCap(CUST_ADDR2) ADDR2, InitCap(CUST_ADDR3) ADDR3, "
            sql &= " InitCap(CUST_CITY) CITY, CUST_STATE STATE, CUST_ZIP_CODE ZIP_CODE, InitCap(CUST_CONTACT) CONTACT"
            sql &= " from ARTCUST1 where CUST_CODE in (Select CUST_CODE from " & ARTCUST1 & ")"
            sql &= " and SREP2_CODE IS NOT NULL"
            sql &= sqlw.Replace("SREP_CODE", "SREP2_CODE")

            If Absx1.optFor("OPTSELECT").Value <> "M" Then

                sql &= " union"

                sql &= " Select ARTCUST2.CUST_CODE || '-' || CUST_STORE_NO CUST, ARTCUST1.SREP_CODE, InitCap(CUST_STORE_NAME) NAME, InitCap(CUST_STORE_ADDR1) ADDR1, InitCap(CUST_STORE_ADDR2) ADDR2, InitCap(CUST_STORE_ADDR3) ADDR3,"
                sql &= " InitCap(CUST_STORE_CITY) CITY, CUST_STORE_STATE STATE, CUST_STORE_ZIP_CODE ZIP_CODE, InitCap(CUST_STORE_CONTACT) CONTACT"
                sql &= " from ARTCUST2, ARTCUST1"
                sql &= " where ARTCUST2.CUST_CODE = ARTCUST1.CUST_CODE"
                sql &= " And ARTCUST2.CUST_STORE_NO <> '000000'"
                sql &= " and ARTCUST2.CUST_CODE IN (Select CUST_CODE from " & ARTCUST1 & ")"
                sql &= sqlw.Replace("ARTCUST1", "ARTCUST2")

                sql &= " union"

                sql &= " Select ARTCUST2.CUST_CODE || '-' || CUST_STORE_NO CUST, ARTCUST2.SREP_CODE, InitCap(CUST_STORE_NAME) NAME, InitCap(CUST_STORE_ADDR1) ADDR1, InitCap(CUST_STORE_ADDR2) ADDR2, InitCap(CUST_STORE_ADDR3) ADDR3,"
                sql &= " InitCap(CUST_STORE_CITY) CITY, CUST_STORE_STATE STATE, CUST_STORE_ZIP_CODE ZIP_CODE, InitCap(CUST_STORE_CONTACT) CONTACT"
                sql &= " from ARTCUST2, ARTCUST1"
                sql &= " where ARTCUST2.CUST_CODE = ARTCUST1.CUST_CODE"
                sql &= " And ARTCUST2.CUST_STORE_NO <> '000000'"
                sql &= " And ARTCUST2.SREP_CODE IS NOT NULL"
                sql &= " and ARTCUST2.CUST_CODE IN (Select CUST_CODE from " & ARTCUST1 & ")"
                sql &= sqlw.Replace("ARTCUST1", "ARTCUST2")
            End If

            .Tables.Add(ASCDATA1.GetDataTable(sql, "ARTCUSL1", 1))
        End With
    End Sub

    Public Overrides Sub Print_Report()
        'Generate_Report(RPT)
        Generate_Report("ARRCUSL2")
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("OPTSELECT").Value = "S" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            End If
        End If
    End Sub


    Private Sub optSelect_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optSelect.ValueChanged
        grpINV_DATE_RANGE.Visible = (optSelect.Value = "C" OrElse optSelect.Value = "M")
        grpPERIOD_RANGE.Visible = (optSelect.Value = "S")
        grpPeriod.Visible = (optSelect.Value = "R")
    End Sub

End Class