Public Class SORSCOM1
    Dim SOTINVH1 As String = ""
    Dim SOTSCOM2 As String = ""
    Private sqlRSTRETLC As String = String.Empty
    Private sqlSOTSCOMO As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, 0, -1)
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        sqlw &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE")
        sqlw &= SQL_in("SREP_CODE", "SOTINVH1.SREP_CODE")
        sqlw &= SQL_in("SALES_DIVISION_CODE", "SOTSREP1.SALES_DIVISION_CODE")
        sqlw &= SQL_in("SREP_CODE_MGR", "NVL(SOTSREP1.SREP_CODE_MGR, SOTSREP1.SREP_CODE)")

        If chkPAYABLE.Checked AndAlso sqlw.Length > 0 Then
            chkPAYABLE.Checked = False
            MessageBox.Show("Since you provided a Filter you cannot Update AP.", "Update AP", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        sqlRSTRETLC = sqlw.Replace("SOTINVH1", "RSTRETLC").Replace("SOTINVH1.SREP_CODE", "SOTINVH1.SELL_CODE")
        sqlSOTSCOMO = sqlw.Replace("SOTINVH1", "SOTSCOMO")

        If chkPAYABLE.Checked Then
            sqlw &= " and SOTINVH1.SREP_COMM_IND = '0'"
            sqlSOTSCOMO &= " and OPS_YYYYPP = '" & RYP & "' and NVL(SREP_COMM_IND, '0') = '0'"
            sqlRSTRETLC &= " and OPS_YYYYPP = '" & RYP & "' and NVL(SELL_COMM_IND, '0') = '0'"
            RWU = "R"
        Else
            sqlw &= " and SOTINVH1.ORDR_YYYYPP_UPDATED = '" & RYP & "'"
            sqlSOTSCOMO &= " and OPS_YYYYPP = '" & RYP & "'"
            sqlRSTRETLC &= " and OPS_YYYYPP = '" & RYP & "'"
        End If

        Prepare_dst(True, sqlw, RYP)

        Check_if_Empty("SOTINVH1")

        ' Need to validate vendor code
        If chkPAYABLE.Checked Then
            Dim SREP_CODE_MGR As String = String.Empty
            Dim errorMessage As String = String.Empty
            For Each rowSOTSCOM1 As DataRow In dst.Tables("SOTSCOM1").Select("SREP_COMM_TOTAL <> 0 ", "SREP_CODE_MGR")

                If SREP_CODE_MGR = rowSOTSCOM1.Item("SREP_CODE_MGR") & String.Empty Then
                    Continue For
                End If

                SREP_CODE_MGR = rowSOTSCOM1.Item("SREP_CODE_MGR") & String.Empty

                Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(New String() {SREP_CODE_MGR})
                Dim VEND_CODE As String = rowSOTSREP1.Item("VEND_CODE") & String.Empty
                Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)

                If rowAPTVEND1 Is Nothing Then
                    errorMessage &= vbCr & "Sales Rep (" & SREP_CODE_MGR & ") is missing Vendor Code"
                Else
                    If rowAPTVEND1.Item("ACCT_CODE") & String.Empty = String.Empty Then
                        errorMessage &= vbCr & "Sales Rep (" & SREP_CODE_MGR & "), Vendor Code (" & rowAPTVEND1.Item("VEND_CODE") & ") is missing GL Account Code"
                    End If

                    If rowAPTVEND1.Item("BANK_CODE") & String.Empty = String.Empty Then
                        errorMessage &= vbCr & "Sales Rep (" & SREP_CODE_MGR & "), Vendor Code (" & rowAPTVEND1.Item("VEND_CODE") & ") is missing Bank Code"
                    End If
                End If
            Next

            If errorMessage.Length > 0 Then
                chkPAYABLE.Checked = False
                RWU = "N"
                MessageBox.Show(errorMessage, "Update", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        End If

    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        RYP = CStr(parms(1))

        Create_Temp_Data(sqlw, RYP)

        With dst
            ASCMAIN1.sql = "Select SOTINVH1.*" _
                & ", SOTINVH1.INV_TYPE INV_TYPE_PRT, SOTINVH1.INV_NO INV_NO_PRT from " & SOTINVH1 & " SOTINVH1"
            Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, False, , 2)
            With dst.Tables("SOTINVH1")
                .Columns.Add("SREP_COMM_CALC", GetType(System.Decimal), "ISNULL(INV_SALES,0) * ISNULL(SREP_COMM_PCT,0) / 100")
                For Each COLUMN_NAME As String In New String() {"CUST_CODE", "CUST_STORE_NO", "CURR_CODE", "CURR_EXCH_RATE", "INV_TYPE_PRT", "INV_NO_PRT"}
                    .Columns(COLUMN_NAME).AllowDBNull = True
                Next
            End With

            Create_TDA(dst.Tables.Add, "SOTSCOM1", "*")
            'Create_TDA(dst.Tables.Add, "ASTATTA2", "*")

            Create_TDA(dst.Tables.Add, "RSTRETLC", "*")
            Create_TDA(dst.Tables.Add, "SOTSCOMO", "*")

            ASCMAIN1.sql = "Select * from " & SOTSCOM2
            Create_TDA(dst.Tables.Add, "SOTSCOM2", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select * from ARTCUST4 where OPS_YYYYPP = :PARM1 and NVL(SREP_COMM_PCT_OVER,0) <> 0"
            Create_TDA(dst.Tables.Add, "ARTCUST4", "**", 0, False, "V", 2)
            dst.Tables("ARTCUST4").Columns.Add("INV_SALES", GetType(System.Decimal))
            dst.Tables("ARTCUST4").Columns.Add("SREP_COMM_OVER", GetType(System.Decimal), "ISNULL(INV_SALES,0) * ISNULL(SREP_COMM_PCT_OVER,0) / 100")

            ' we need another table that might possible show a single invoice connected to 2 sales reps
            ASCMAIN1.sql = "Select INV_TYPE, INV_NO, SREP_CODE, INV_SALES from " & SOTINVH1
            Create_TDA(dst.Tables.Add, "SOTINVHN", "**", 0, False, , 3)
            With dst.Tables("SOTINVHN")
                .Columns.Add("SREP_CODE_MGR")
                .Columns.Add("VEND_CODE")
            End With

            ASCMAIN1.sql = "Select SREP_CODE, SREP_NAME, VEND_CODE, NVL(SREP_CODE_MGR,SREP_CODE) SREP_CODE_MGR, SREP_EMAIL" & vbCrLf _
                         & " from SOTSREP1"
            Create_TDA(dst.Tables.Add, "SOTSREP1", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select CUST_CODE, CUST_NAME, SREP_CODE" & vbCrLf _
                & " from ARTCUST1" & vbCrLf _
                & " where CUST_CODE in (Select Distinct CUST_CODE from " & SOTINVH1 & ")"
            Create_TDA(dst.Tables.Add, "ARTCUST1", "**", 0, False, , 1)

            ASCMAIN1.sql = "Select ARTCUST2.*" & vbCrLf _
                         & " from ARTCUST2" & vbCrLf _
                         & " where (CUST_CODE,CUST_STORE_NO) in (Select Distinct CUST_CODE, CUST_STORE_NO from " & SOTINVH1 & ")" & vbCrLf _
                         & " Union " & vbCrLf _
                         & "Select ARTCUST2.*" & vbCrLf _
                         & " from ARTCUST2" & vbCrLf _
                         & " where (CUST_CODE,CUST_STORE_NO) in (Select Distinct CUST_CODE, CUST_STORE_NO from RSTRETLC WHERE OPS_YYYYPP = '" & RYP & "')" & vbCrLf _
                         & " Union " & vbCrLf _
                         & "Select ARTCUST2.*" & vbCrLf _
                         & " from ARTCUST2" & vbCrLf _
                         & " where (CUST_CODE,CUST_STORE_NO) in (Select Distinct CUST_CODE, CUST_STORE_NO from SOTSCOMO WHERE OPS_YYYYPP = '" & RYP & "')"
            Create_TDA(dst.Tables.Add, "ARTCUST2", "**", 0, False, , 2)

            Create_TDA(dst.Tables.Add, "APTINVH1", "*")
            Create_TDA(dst.Tables.Add, "APTINVH2", "*")

            .Tables.Add("SOTSCOMH")
            With .Tables("SOTSCOMH")
                .Columns.Add("GROUP", GetType(System.String))
                .Columns.Add("SREP_CODE_MGR", GetType(System.String))
            End With

            .Tables.Add("SOTSCOMD")
            With .Tables("SOTSCOMD")
                .Columns.Add("SREP_CODE_MGR", GetType(System.String))
                .Columns.Add("SREP_CODE", GetType(System.String))
                .Columns.Add("SALES", GetType(System.Decimal))
                .Columns("SALES").DefaultValue = 0
                .Columns.Add("ADJUSTMENTS", GetType(System.Decimal))
                .Columns("ADJUSTMENTS").DefaultValue = 0
                .Columns.Add("RETAILS", GetType(System.Decimal))
                .Columns("RETAILS").DefaultValue = 0
                .Columns.Add("OVERRIDES", GetType(System.Decimal))
                .Columns("OVERRIDES").DefaultValue = 0
                .Columns.Add("COMMISSION", GetType(System.Decimal))
                .Columns("COMMISSION").DefaultValue = 0
                .Columns.Add("SREP_COMM_TOTAL", GetType(System.Decimal))
                .Columns("SREP_COMM_TOTAL").DefaultValue = 0
            End With

            .Tables.Add("SOTSCOMT")
            With .Tables("SOTSCOMT")
                .Columns.Add("SREP_CODE", GetType(System.String))
                .Columns.Add("TRADE_CLASS_CODE", GetType(System.String))
                .Columns.Add("CHANNEL_CODE", GetType(System.String))
                .Columns.Add("SREP_COMM_CALC", GetType(System.Decimal))
                .Columns("SREP_COMM_CALC").DefaultValue = 0
                .Columns.Add("SREP_COMM_ADJ", GetType(System.Decimal))
                .Columns("SREP_COMM_ADJ").DefaultValue = 0
                .Columns.Add("SREP_COMM_ADJ_MISC", GetType(System.Decimal))
                .Columns("SREP_COMM_ADJ_MISC").DefaultValue = 0
                .Columns.Add("SREP_COMM_OVER", GetType(System.Decimal))
                .Columns("SREP_COMM_OVER").DefaultValue = 0
                .Columns.Add("SREP_COMM_TOTAL", GetType(System.Decimal))
                .Columns("SREP_COMM_TOTAL").DefaultValue = 0
                .Columns("SREP_COMM_TOTAL").Expression = "SREP_COMM_CALC + SREP_COMM_ADJ + SREP_COMM_ADJ_MISC + SREP_COMM_OVER"
            End With

            Create_TDA(dst.Tables.Add, "SOTTCLS1", "*", 0, False)
            Create_TDA(dst.Tables.Add, "TATUSER1", "*", 0, False)
            Create_TDA(dst.Tables.Add, "SOTCHAN1", "*", 0, False)

            Create_TDA(dst.Tables.Add, "ASTWRPT0", "*")
        End With

        For Each TABLE_NAME As String In New String() _
            {"TATUSER1", "SOTTCLS1", "SOTCHAN1"}
            Fill_Records(TABLE_NAME)
        Next


        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)
        RYP = parms(1)

        If sqlw <> "" Then
            Create_Temp_Data(sqlw, RYP)
        End If
        EnforceConstraints(False)
        Fill_Records("SOTINVH1")
        Fill_Records("SOTSCOM2")
        Fill_Records("ARTCUST4", RYP)
        Fill_Records("SOTINVHN")
        'Fill_Records("SOTINVHS")
        Fill_Records("SOTSREP1")
        Fill_Records("ARTCUST1")
        Fill_Records("ARTCUST2")

        If Not (sqlRSTRETLC.Contains("SALES_DIVISION_CODE") OrElse sqlRSTRETLC.Contains("SREP_CODE_MGR")) Then
            Fill_Records("RSTRETLC", String.Empty, True, "Select * from RSTRETLC " & ASCMAIN1.SQL_Add_WHERE(sqlRSTRETLC))
        Else
            Fill_Records("RSTRETLC", String.Empty, True, "Select RSTRETLC.* from RSTRETLC, SOTSREP1 WHERE RSTRETLC.SELL_CODE =  SOTSREP1.SREP_CODE " & sqlRSTRETLC)
        End If


        If Not (sqlRSTRETLC.Contains("SALES_DIVISION_CODE") OrElse sqlRSTRETLC.Contains("SREP_CODE_MGR")) Then
            Fill_Records("SOTSCOMO", String.Empty, True, "Select * from SOTSCOMO " & ASCMAIN1.SQL_Add_WHERE(sqlSOTSCOMO))
        Else
            Fill_Records("SOTSCOMO", String.Empty, True, "Select SOTSCOMO.* from SOTSCOMO, SOTSREP1 WHERE SOTSCOMO.SREP_CODE = SOTSREP1.SREP_CODE  " & sqlSOTSCOMO)
        End If

        EnforceConstraints(True)

        Dim LYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
        For Each row As DataRow In dst.Tables("SOTSCOM2").Select("")
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = "A"
                .Item("INV_NO") = row.Item("SREP_COMM_ADJ_NO")
                .Item("SREP_CODE") = row.Item("SREP_CODE")
                .Item("SREP_COMM_ADJ") = row.Item("SREP_COMM_ADJ_MISC")
                .Item("INV_COMMENT") = row.Item("SREP_COMM_ADJ_NOTE")
                .Item("INV_TYPE_PRT") = row.Item("INV_TYPE")
                .Item("INV_NO_PRT") = row.Item("INV_NO")
                .Item("INV_SALES") = row.Item("INV_SALES")
                .Item("SREP_COMM_PCT") = row.Item("SREP_COMM_PCT")
                .Item("SREP_COMM_ADJ") = row.Item("SREP_COMM_ADJ")
                If row.Item("OPS_YYYYPP") & "" = "" Then
                    .Item("ORDR_YYYYPP_UPDATED") = LYP
                Else
                    .Item("ORDR_YYYYPP_UPDATED") = row.Item("OPS_YYYYPP")
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    .Item("TRADE_CLASS_CODE") = "IND"
                    .Item("CHANNEL_CODE") = "WS"
                End If

            End With
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

            Dim rowSOTINVHN As DataRow = dst.Tables("SOTINVHN").NewRow
            With rowSOTINVHN
                .Item("INV_TYPE") = "A"
                .Item("INV_NO") = row.Item("SREP_COMM_ADJ_NO")
                .Item("SREP_CODE") = row.Item("SREP_CODE")
                .Item("INV_SALES") = row.Item("INV_SALES")
            End With
            dst.Tables("SOTINVHN").Rows.Add(rowSOTINVHN)
        Next

        For Each row As DataRow In dst.Tables("ARTCUST4").Select("")
            Dim CUST_CODE As String = row.Item("CUST_CODE")
            Dim INV_SALES As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(INV_SALES)", "CUST_CODE = '" & CUST_CODE & "'") & "")
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = "O"
                .Item("INV_NO") = CUST_CODE
                .Item("CUST_CODE") = CUST_CODE
                .Item("SREP_CODE") = row.Item("SREP_CODE_OVER")

                .Item("INV_COMMENT") = Val(row.Item("SREP_COMM_PCT_OVER") & "") & "% Override"
                .Item("INV_TYPE_PRT") = ""
                .Item("INV_NO_PRT") = ""
                .Item("INV_SALES") = INV_SALES
                .Item("SREP_COMM_PCT") = row.Item("SREP_COMM_PCT_OVER")
                .Item("SREP_COMM_ADJ") = 0
                If row.Item("OPS_YYYYPP") & "" = "" Then
                    .Item("ORDR_YYYYPP_UPDATED") = LYP
                Else
                    .Item("ORDR_YYYYPP_UPDATED") = row.Item("OPS_YYYYPP")
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    Dim TRADE_CLASS_CODE As String = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""
                    Dim rowSOTTCLS1 As DataRow = LookUp("SOTTCLS1", TRADE_CLASS_CODE, True)
                    Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE") & ""
                    .Item("TRADE_CLASS_CODE") = TRADE_CLASS_CODE
                    .Item("CHANNEL_CODE") = CHANNEL_CODE
                End If

            End With
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

            Dim rowSOTINVHN As DataRow = dst.Tables("SOTINVHN").NewRow
            With rowSOTINVHN
                .Item("INV_TYPE") = "O"
                .Item("INV_NO") = CUST_CODE
                .Item("SREP_CODE") = row.Item("SREP_CODE_OVER")
                .Item("INV_SALES") = INV_SALES
            End With
            dst.Tables("SOTINVHN").Rows.Add(rowSOTINVHN)
        Next

        ' Retail sales
        Dim retailInvNo As String = String.Empty
        Dim ictr As Int16 = 0
        For Each rowRSTRETLC As DataRow In dst.Tables("RSTRETLC").Select("")
            Dim CUST_CODE As String = rowRSTRETLC.Item("CUST_CODE")
            Dim INV_SALES As Decimal = Val(rowRSTRETLC.Item("AMT_SOLD") & String.Empty)

            rowRSTRETLC.Item("SELL_COMM_IND") = "1"
            rowRSTRETLC.Item("SELL_COMM_XNO") = XNO

            If Val(rowRSTRETLC.Item("SELL_COMM_PCT") & String.Empty) = 0 Then
                Continue For
            End If

            ictr += 1
            retailInvNo = "R" & ictr.ToString.PadLeft(9, "0")

            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = "R"

                .Item("INV_NO") = retailInvNo
                .Item("CUST_CODE") = CUST_CODE
                .Item("CUST_STORE_NO") = rowRSTRETLC.Item("CUST_STORE_NO")
                .Item("SREP_CODE") = rowRSTRETLC.Item("SELL_CODE")

                .Item("INV_COMMENT") = "" ' "Retail Sales Comm"
                .Item("INV_TYPE_PRT") = ""
                .Item("INV_NO_PRT") = ""
                .Item("INV_SALES") = INV_SALES
                .Item("SREP_COMM_PCT") = rowRSTRETLC.Item("SELL_COMM_PCT")
                .Item("SREP_COMM_ADJ") = rowRSTRETLC.Item("SELL_COMM_ADJ")
                If rowRSTRETLC.Item("OPS_YYYYPP") & "" = "" Then
                    .Item("ORDR_YYYYPP_UPDATED") = LYP
                Else
                    .Item("ORDR_YYYYPP_UPDATED") = rowRSTRETLC.Item("OPS_YYYYPP")
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    Dim TRADE_CLASS_CODE As String = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""
                    Dim rowSOTTCLS1 As DataRow = LookUp("SOTTCLS1", TRADE_CLASS_CODE, True)
                    Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE") & ""
                    .Item("TRADE_CLASS_CODE") = TRADE_CLASS_CODE
                    .Item("CHANNEL_CODE") = CHANNEL_CODE
                End If

            End With
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

            Dim rowSOTINVHN As DataRow = dst.Tables("SOTINVHN").NewRow
            With rowSOTINVHN
                .Item("INV_TYPE") = "R"
                .Item("INV_NO") = retailInvNo
                .Item("SREP_CODE") = rowRSTRETLC.Item("SELL_CODE")
                .Item("INV_SALES") = INV_SALES
            End With
            dst.Tables("SOTINVHN").Rows.Add(rowSOTINVHN)
        Next

        ' Override sales
        For Each rowSOTSCOMO As DataRow In dst.Tables("SOTSCOMO").Select("")
            Dim CUST_CODE As String = rowSOTSCOMO.Item("CUST_CODE")
            Dim INV_SALES As Decimal = Val(rowSOTSCOMO.Item("INV_SALES") & String.Empty)

            rowSOTSCOMO.Item("SREP_COMM_IND") = "1"
            rowSOTSCOMO.Item("SREP_COMM_XNO") = XNO

            If Val(rowSOTSCOMO.Item("SREP_COMM_PCT") & String.Empty) = 0 AndAlso Val(rowSOTSCOMO.Item("SREP_COMM_ADJ") & String.Empty) = 0 Then
                Continue For
            End If

            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").NewRow
            With rowSOTINVH1
                .Item("INV_TYPE") = "V"
                .Item("INV_NO") = rowSOTSCOMO.Item("INV_NO")
                .Item("CUST_CODE") = CUST_CODE
                .Item("SREP_CODE") = rowSOTSCOMO.Item("SREP_CODE")

                .Item("INV_COMMENT") = "" ' "Retail Sales Comm"
                .Item("INV_TYPE_PRT") = ""
                .Item("INV_NO_PRT") = ""
                .Item("INV_SALES") = INV_SALES
                .Item("SREP_COMM_PCT") = rowSOTSCOMO.Item("SREP_COMM_PCT")
                .Item("SREP_COMM_ADJ") = rowSOTSCOMO.Item("SREP_COMM_ADJ")
                If rowSOTSCOMO.Item("OPS_YYYYPP") & "" = "" Then
                    .Item("ORDR_YYYYPP_UPDATED") = LYP
                Else
                    .Item("ORDR_YYYYPP_UPDATED") = rowSOTSCOMO.Item("OPS_YYYYPP")
                End If

                If ASCMAIN1.CLIENT = "AHA" Then
                    Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    Dim TRADE_CLASS_CODE As String = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""
                    Dim rowSOTTCLS1 As DataRow = LookUp("SOTTCLS1", TRADE_CLASS_CODE, True)
                    Dim CHANNEL_CODE As String = rowSOTTCLS1.Item("CHANNEL_CODE") & ""
                    .Item("TRADE_CLASS_CODE") = TRADE_CLASS_CODE
                    .Item("CHANNEL_CODE") = CHANNEL_CODE
                End If

            End With
            dst.Tables("SOTINVH1").Rows.Add(rowSOTINVH1)

            Dim rowSOTINVHN As DataRow = dst.Tables("SOTINVHN").NewRow
            With rowSOTINVHN
                .Item("INV_TYPE") = "V"
                .Item("INV_NO") = rowSOTSCOMO.Item("INV_NO")
                .Item("SREP_CODE") = rowSOTSCOMO.Item("SREP_CODE")
                .Item("INV_SALES") = INV_SALES
            End With
            dst.Tables("SOTINVHN").Rows.Add(rowSOTINVHN)
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH1"), New String() {"SREP_CODE"}).Select("")
            Dim SREP_CODE As String = row.Item("SREP_CODE")
            Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(SREP_CODE)

            ' Checks are paid to the manager
            Dim VEND_CODE As String = rowSOTSREP1.Item("VEND_CODE") & ""

            For Each rowSOTINVHN As DataRow In dst.Tables("SOTINVHN").Select("SREP_CODE = '" & SREP_CODE & "'")
                rowSOTINVHN.Item("SREP_CODE_MGR") = rowSOTSREP1.Item("SREP_CODE_MGR")
                rowSOTINVHN.Item("VEND_CODE") = VEND_CODE
            Next

            Dim rowSOTSCOM1 As DataRow = dst.Tables("SOTSCOM1").NewRow
            With rowSOTSCOM1
                .Item("SREP_COMM_XNO") = XNO
                .Item("SREP_CODE") = SREP_CODE
                .Item("SREP_CODE_MGR") = rowSOTSREP1.Item("SREP_CODE_MGR")
                .Item("OPS_YYYYPP") = RYP
                Dim SREP_COMM_CALC As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_CALC)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE <> 'O'") & "")
                .Item("SREP_COMM_CALC") = SREP_COMM_CALC
                Dim SREP_COMM_ADJ As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_ADJ)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE <> 'A'") & "")
                .Item("SREP_COMM_ADJ") = SREP_COMM_ADJ
                Dim SREP_COMM_ADJ_MISC As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_ADJ)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE = 'A'") & "")
                .Item("SREP_COMM_ADJ_MISC") = SREP_COMM_ADJ_MISC
                Dim SREP_COMM_OVER As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_CALC)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE = 'O'") & "")
                .Item("SREP_COMM_OVER") = SREP_COMM_OVER
                .Item("SREP_COMM_TOTAL") = Math.Round(SREP_COMM_CALC + SREP_COMM_ADJ + SREP_COMM_ADJ_MISC + SREP_COMM_OVER, 2, MidpointRounding.AwayFromZero)
                .Item("VEND_CODE") = VEND_CODE
                Dim VOUCHER_NO As String = ASCMAIN1.Next_Control_No("APTINVH1.VOUCHER_NO")
                .Item("VOUCHER_NO") = VOUCHER_NO
            End With

            dst.Tables("SOTSCOM1").Rows.Add(rowSOTSCOM1)

            ' Used for Totals Summary
            Dim SREP_CODE_MGR = rowSOTSREP1.Item("SREP_CODE_MGR") & String.Empty
            If dst.Tables("SOTSCOMH").Select("SREP_CODE_MGR = '" & SREP_CODE_MGR & "'").Length = 0 Then
                dst.Tables("SOTSCOMH").Rows.Add(New Object() {"1", SREP_CODE_MGR})
            End If

            Dim RETAILS As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_CALC)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE = 'R'") & "")
            RETAILS += Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_ADJ)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE = 'R'") & "")

            Dim OVERRIDE As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_CALC)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE = 'V'") & "")
            OVERRIDE += Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_ADJ)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE = 'V'") & "")

            Dim ADJUSTMENTS As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_ADJ)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE <> 'A' and INV_TYPE <> 'V' and INV_TYPE <> 'R'") & "")
            Dim COMMISSION As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_CALC)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE <> 'V' and INV_TYPE <> 'R'") & "")
            Dim SALES As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(INV_SALES)", "SREP_CODE = '" & SREP_CODE & "' and INV_TYPE <> 'A' and INV_TYPE <> 'O' and INV_TYPE <> 'V' and INV_TYPE <> 'R'") & "")

            Dim rowSOTSCOMD As DataRow = dst.Tables("SOTSCOMD").NewRow
            With rowSOTSCOMD
                .Item("SREP_CODE") = SREP_CODE
                .Item("SREP_CODE_MGR") = SREP_CODE_MGR
                .Item("SALES") = Math.Round(SALES, 2)
                .Item("COMMISSION") = Math.Round(COMMISSION, 2)
                .Item("ADJUSTMENTS") = Math.Round(ADJUSTMENTS, 2)
                .Item("RETAILS") = Math.Round(RETAILS, 2)
                .Item("OVERRIDES") = Math.Round(OVERRIDE, 2)
                .Item("SREP_COMM_TOTAL") = rowSOTSCOM1.Item("SREP_COMM_TOTAL")
            End With
            dst.Tables("SOTSCOMD").Rows.Add(rowSOTSCOMD)
        Next

        For Each row As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("SOTINVH1"), New String() {"SREP_CODE", "TRADE_CLASS_CODE", "CHANNEL_CODE"}).Select("")
            Dim SREP_CODE As String = row.Item("SREP_CODE")
            Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(SREP_CODE)

            Dim TRADE_CLASS_CODE As String = row.Item("TRADE_CLASS_CODE")
            Dim CHANNEL_CODE As String = row.Item("CHANNEL_CODE")

            Dim sqlT As String = "SREP_CODE = '" & SREP_CODE & "' AND TRADE_CLASS_CODE = '" & TRADE_CLASS_CODE & "'"

            Dim rowSOTSCOMT As DataRow = dst.Tables("SOTSCOMT").NewRow
            With rowSOTSCOMT
                .Item("SREP_CODE") = SREP_CODE
                .Item("TRADE_CLASS_CODE") = TRADE_CLASS_CODE
                .Item("CHANNEL_CODE") = CHANNEL_CODE
                Dim SREP_COMM_CALC As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_CALC)", sqlT & " and INV_TYPE <> 'O'") & "")
                .Item("SREP_COMM_CALC") = SREP_COMM_CALC
                Dim SREP_COMM_ADJ As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_ADJ)", sqlT & " and INV_TYPE <> 'A'") & "")
                .Item("SREP_COMM_ADJ") = SREP_COMM_ADJ
                Dim SREP_COMM_ADJ_MISC As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_ADJ)", sqlT & " and INV_TYPE = 'A'") & "")
                .Item("SREP_COMM_ADJ_MISC") = SREP_COMM_ADJ_MISC
                Dim SREP_COMM_OVER As Decimal = Val(dst.Tables("SOTINVH1").Compute("SUM(SREP_COMM_CALC)", sqlT & " and INV_TYPE = 'O'") & "")
                .Item("SREP_COMM_OVER") = SREP_COMM_OVER
            End With
            dst.Tables("SOTSCOMT").Rows.Add(rowSOTSCOMT)
        Next
    End Sub

    Sub Create_Temp_Data(ByVal sqlW As String, ByVal RYP As String)

        ASCMAIN1.sql = "Select SOTINVH1.*, ARTCUST1.TRADE_CLASS_CODE, SOTTCLS1.CHANNEL_CODE, SOTCHAN1.SEG2_CODE" & vbCrLf _
                             & " from SOTINVH1,ARTCUST1,SOTTCLS1,SOTCHAN1,SOTSREP1" & vbCrLf _
                             & " where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                             & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                             & "   and SOTCHAN1.CHANNEL_CODE (+) = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                             & "   and SOTSREP1.SREP_CODE = SOTINVH1.SREP_CODE" & vbCrLf _
                             & sqlW

        If SOTINVH1 = "" Then
            SOTINVH1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE,INV_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVH1)
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVH1 & " " & ASCMAIN1.sql)
        End If

        ' Remove $0 Credits from data
        ASCDATA1.ExecuteSQL("Update  " & SOTINVH1 & " set INV_SALES = 0 where INV_SALES is null")
        ASCDATA1.ExecuteSQL("Delete from  " & SOTINVH1 & " where INV_TYPE = 'C' and INV_SALES = 0")

        Dim tables As String = "SOTSCOM2"
        Dim sqlWhere As String = ""

        Dim sqlWork As String = " " & SQLW & " "
        sqlWork = sqlWork.ToUpper.Replace(" AND ", Chr(0) & " AND ")
        For Each restriction As String In (" " & sqlWork & " ").Split(Chr(0))
            restriction = restriction.Trim
            restriction = restriction.Replace(Chr(0), "")

            If restriction.Contains("ORDR_YYYYPP_UPDATED") Then
                restriction = " and NVL(SOTSCOM2.OPS_YYYYPP,'" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1) & "') = '" & RYP & "'"
            ElseIf restriction.Contains("SREP_CODE_MGR") Then
                If Not tables.Contains("SOTSREP1") Then
                    tables = tables & ", SOTSREP1"
                    sqlWhere &= " and SOTSCOM2.SREP_CODE = SOTSREP1.SREP_CODE"
                End If
            ElseIf restriction.Contains("SREP_CODE") Then
                restriction = restriction.Replace("SOTINVH1", "SOTSCOM2")
            ElseIf restriction.Contains("CUST_CODE") Then
                If Not tables.Contains("SOTINVH1") Then
                    tables = tables & ", SOTINVH1"
                    sqlWhere &= " and SOTSCOM2.INV_TYPE = SOTINVH1.INV_TYPE and SOTSCOM2.INV_NO = SOTINVH1.INV_NO"
                End If
            ElseIf restriction.Contains("SALES_DIVISION_CODE") Then
                If Not tables.Contains("SOTSREP1") Then
                    tables = tables & ", SOTSREP1"
                    sqlWhere &= " and SOTSCOM2.SREP_CODE = SOTSREP1.SREP_CODE"
                End If
            ElseIf restriction.Contains("SREP_COMM_IND") Then
                If Not tables.Contains("SOTINVH1") Then
                    tables = tables & ", SOTINVH1"
                    sqlWhere &= " and SOTSCOM2.INV_TYPE = SOTINVH1.INV_TYPE and SOTSCOM2.INV_NO = SOTINVH1.INV_NO"
                End If
            End If

            sqlWhere &= " " & restriction
        Next

        ASCMAIN1.sql = "Select SOTSCOM2.* from " & tables & " " & ASCMAIN1.SQL_Add_WHERE(sqlWhere.Trim)

        If SOTSCOM2 = "" Then
            SOTSCOM2 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTSCOM2)
            ASCDATA1.ExecuteSQL("Insert into " & SOTSCOM2 & " " & ASCMAIN1.sql)
        End If
    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        If chkPAYABLE.Checked Then
            SUBT = "Commissions Payable - " & RYPLEGEND
        Else
            SUBT = "Commission Report for " & RYPLEGEND
        End If

        Generate_Report(RPT, "", SUBT)
        Generate_Report("SORSCOMT", "Commission Summary Totals", SUBT)

        If chkPAYABLE.Checked Then
            Dim TBL As DataTable = dst.Tables("SOTINVHN").Copy
            For Each rowSOTSCOM1 As DataRow In dst.Tables("SOTSCOM1").Select("")
                Dim SREP_CODE As String = rowSOTSCOM1.Item("SREP_CODE")
                dst.Tables("SOTINVHN").Rows.Clear()
                Dim dvw As DataView = TBL.DefaultView
                dvw.RowFilter = "SREP_CODE = '" & SREP_CODE & "'"
                dst.Tables("SOTINVHN").Merge(dvw.ToTable)
                SUBT = SREP_CODE & " Commissions Payable - " & RYPLEGEND
                Dim REPORT_NO As String = Generate_Report(RPT, , SUBT, "", "PDF")
                rowSOTSCOM1.Item("REPORT_NO") = REPORT_NO
            Next
        End If

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub

    Overrides Sub Update_Record()

        ASCMAIN1.sql = "Update SOTINVH1 " _
            & " Set SREP_COMM_IND = '1', SREP_COMM_XNO = '" & XNO & "'" _
            & " where (INV_TYPE,INV_NO) in (Select INV_TYPE, INV_NO from " & SOTINVH1 & " )"
        ASCDATA1.ExecuteSQL()

        For Each rowSOTSCOM1 As DataRow In dst.Tables("SOTSCOM1").Select("SREP_COMM_TOTAL <> 0 ")
            ' Checks are made payable to Sale Rep Manager, if one is assigned
            Dim SREP_CODE As String = rowSOTSCOM1.Item("SREP_CODE_MGR") & String.Empty
            If SREP_CODE.Length = 0 Then
                SREP_CODE = rowSOTSCOM1.Item("SREP_CODE") & String.Empty
            End If

            Dim SREP_COMM_TOTAL As String = Val(rowSOTSCOM1.Item("SREP_COMM_TOTAL") & "")

            Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(New String() {SREP_CODE})
            Dim VEND_CODE As String = rowSOTSREP1.Item("VEND_CODE") & String.Empty
            Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
            Dim REPORT_NO As String = rowSOTSCOM1.Item("REPORT_NO")
            Dim VOUCHER_NO As String = rowSOTSCOM1.Item("VOUCHER_NO")

            Dim rowTATUSER1s() As DataRow = dst.Tables("TATUSER1").Select("SREP_CODE = '" & SREP_CODE & "'")
            If rowTATUSER1s.Length > 0 Then
                Dim rowTATUSER1 As DataRow = rowTATUSER1s(0)
                Dim FILENAME_body As String = ASCMAIN1.DBS_COMPANY & "_" & REPORT_NO
                Dim FULL_FILENAME As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_WEB_REPORTS_DIR") _
                    & "\" & Mid(REPORT_NO, 1, 7) & "\" & REPORT_NO

                My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & FILENAME_body & ".pdf", _
                                                FULL_FILENAME & ".pdf")

                Dim rowASTWRPT0 As DataRow = dst.Tables("ASTWRPT0").NewRow
                With rowASTWRPT0
                    .Item("USER_ID") = rowTATUSER1.Item("USER_ID")
                    .Item("REPORT_NO") = REPORT_NO
                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("REPORT_TYPE") = "PDF"
                    .Item("RPT_TITLE") = RPT_TITLE
                    .Item("SET_DESC") = ""
                    .Item("SET_ID") = ""
                End With
                dst.Tables("ASTWRPT0").Rows.Add(rowASTWRPT0)

                Dim ATTACHMENTs As New Dictionary(Of String, String)
                ATTACHMENTs.Add(FILENAME_body & ".pdf", ASCMAIN1.Folders("Temp") & FILENAME_body & ".pdf")

                Dim SUBJECT As String = RPT_TITLE & " for " & RYPLEGEND
                Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
                Dim EMAIL_ADDRESS As String = rowSOTSREP1.Item("SREP_EMAIL") & ""
                EMAIL_ADDRESSs.Add(EMAIL_ADDRESS, rowSOTSREP1.Item("SREP_NAME") & "")

                Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                       (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs, _
                        SUBJECT, "SORSCOM1", True, True, SREP_CODE, rowSOTSREP1.Item("SREP_NAME") & "", "Sales Rep")
            End If

            Dim rowAPTINVH1 As DataRow = dst.Tables("APTINVH1").NewRow
            With rowAPTINVH1
                .Item("VOUCHER_NO") = VOUCHER_NO
                .Item("VEND_CODE") = VEND_CODE
                .Item("INV_TYPE") = "I"
                .Item("INV_NUM") = "COMM_" & rowSOTSCOM1.Item("SREP_CODE") & "_" & REPORT_NO
                .Item("INV_DATE") = DATETIME_STAMP.Date
                .Item("INV_AMT") = SREP_COMM_TOTAL
                .Item("INV_REF") = XNO
                .Item("INV_PYMT_CYCLE") = "S"
                .Item("INV_STATUS") = "O"
                .Item("INV_DUE_DATE") = DATETIME_STAMP.Date
                .Item("INV_BALANCE") = SREP_COMM_TOTAL
                .Item("REGISTER_IND") = "0"
                .Item("OPS_YYYYPP") = ASCMAIN1.CYP
                .Item("OPS_YYYYPP_ACCRUE") = RYP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("INV_1099_AMT") = SREP_COMM_TOTAL
                .Item("CURR_CODE") = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")
                .Item("CURR_EXCH_RATE") = 1
                .Item("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
                .Item("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
                .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                For Each COLUMN_NAME As String In New String() _
                    {"TERM_CODE", "POST_CODE", "VEND_CODE_AP", "BANK_CODE"}
                    .Item(COLUMN_NAME) = rowAPTVEND1.Item(COLUMN_NAME)
                Next

            End With
            dst.Tables("APTINVH1").Rows.Add(rowAPTINVH1)

            ' PROBABLY WILL BE ASKED SOMEDAY TO DO THIS BY COLLECTION/BRAND

            Dim DIST_AMT As Decimal = 0
            Dim DIST_AMT_TOTAL As Decimal = 0

            Dim VOUCHER_LNO As Integer = 0
            For Each rowSOTSCOMT As DataRow In dst.Tables("SOTSCOMT").Select("SREP_CODE = '" & rowSOTSCOM1.Item("SREP_CODE") & "'")
                Dim rowAPTINVH2 As DataRow = dst.Tables("APTINVH2").NewRow
                With rowAPTINVH2
                    .Item("VOUCHER_NO") = VOUCHER_NO
                    VOUCHER_LNO += 1
                    .Item("VOUCHER_LNO") = VOUCHER_LNO
                    .Item("ACCT_CODE") = rowAPTVEND1.Item("ACCT_CODE")

                    DIST_AMT = System.Math.Round(Val(rowSOTSCOMT.Item("SREP_COMM_TOTAL") & ""), 2)
                    DIST_AMT_TOTAL += DIST_AMT

                    Dim TRADE_CLASS_CODE As String = rowSOTSCOMT.Item("TRADE_CLASS_CODE")
                    Dim rowSOTTCLS1 As DataRow = dst.Tables("SOTTCLS1").Rows.Find(TRADE_CLASS_CODE)
                    Dim SEG3_CODE As String = rowSOTTCLS1.Item("SEG3_CODE") & ""
                    If SEG3_CODE = "" Then SEG3_CODE = TRADE_CLASS_CODE

                    Dim CHANNEL_CODE As String = rowSOTSCOMT.Item("CHANNEL_CODE")
                    Dim rowSOTCHAN1 As DataRow = dst.Tables("SOTCHAN1").Rows.Find(CHANNEL_CODE)
                    Dim SEG2_CODE As String = rowSOTCHAN1.Item("SEG2_CODE") & ""
                    If SEG2_CODE = "" Then SEG2_CODE = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")

                    .Item("SEG2_CODE") = SEG2_CODE
                    .Item("SEG3_CODE") = SEG3_CODE
                    .Item("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
                    .Item("INV_LINE_AMT") = DIST_AMT ' SREP_COMM_TOTAL
                End With
                dst.Tables("APTINVH2").Rows.Add(rowAPTINVH2)
            Next

            If System.Math.Round(DIST_AMT_TOTAL, 2) <> SREP_COMM_TOTAL Then
                Throw New Exception("Sales Rep Totals do not Balance by Channel - " & SREP_CODE)
            End If

            ' TABLE_NAME, COLUMN_NAME, CODE_VALUE, ATTACHMENT_NO, ATTACHMENT_EXT
            'Dim rowASTATTA2 As DataRow = dst.Tables("ASTATTA2").NewRow
            'With rowASTATTA2
            '    .Item("TABLE_NAME") = "APTINVH1"
            '    .Item("COLUMN_NAME") = "VOUCHER_NO"
            '    .Item("CODE_VALUE") = VOUCHER_NO
            '    .Item("ATTACHMENT_NO") = ASCMAIN1.Next_Control_No("ASTATTA2.ATTACHMENT_NO")
            '    .Item("ATTACHMENT_EXT") = "PDF"
            '    .Item("ATTACHMENT_DESC") = "Commission Report"
            '    .Item("ATTACHMENT_FILENAME") = REPORT_NO
            '    .Item("COMPUTER_NAME") = ASCMAIN1.COMPUTER_NAME
            '    .Item("SESSION_NO") = ASCMAIN1.SESSION_NO
            '    .Item("INIT_DATE") = DATETIME_STAMP
            '    .Item("INIT_OPER") = ASCMAIN1.USER_ID
            '    .Item("ATTACHMENT_TYPE") = "COM"
            '    '.Item("ATTACHMENT_ORIGINATOR") = "?"
            '    .Item("ATTACHMENT_DATETIME") = DATETIME_STAMP
            '    '.Item("ATTACHMENT_STATUS") = "?"
            'End With
            'dst.Tables("ASTATTA2").Rows.Add(rowASTATTA2)
        Next

        Update_Record_TDA("APTINVH1")
        Update_Record_TDA("APTINVH2")

        Update_Record_TDA("RSTRETLC")
        Update_Record_TDA("SOTSCOMO")
        Update_Record_TDA("SOTSCOM1")

        'Update_Record_TDA("ASTATTA2")

        Update_Record_TDA("ASTWRPT0")
    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            'Case "ORDR_GROUP_NO"
            '    sqlw = " SOTORDR0.ORDR_DATE > SYSDATE -180"
        End Select
        Return sqlw
    End Function

    Private Sub chkPAYABLE_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkPAYABLE.CheckedChanged
        Set_Read_Only_for_ctl(Absx1.cmbFor("RYP"), chkPAYABLE.Checked)
        If chkPAYABLE.Checked Then
            Absx1.cmbFor("RYP").ActiveRow = Absx1.cmbFor("RYP").Rows(Absx1.cmbFor("RYP").Rows.Count - 2)
        End If
    End Sub

End Class