Public Class SORUPDT1
    Private SOTINVH1 As String
    Private SOTINVH2 As String
    Private SOTINVH4 As String
    Private SOTINVHD As String

    Dim sqlSOTINVH1 As String
    Dim sqlSOTINVH2 As String
    Dim sqlSOTINVHD As String

    Private SOTORDR1 As String
    Private SOTORDR2 As String
    Private SOTORDR5 As String
    Private SOTORDRT As String

    Private ICTTRNE1 As String
    Private ICTTRNE2 As String
    Private ICTSTAT2_BO As String
    Dim GL_PARM_CURR_CODE As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("ARTPARM1")
        Get_PARM("SOTPARM1")
        Get_PARM("GLTPARM1")
        GL_PARM_CURR_CODE = ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE")

        With dst
            Create_Temp_Tables(True)

            Create_TDA(.Tables.Add("SOTINVH1"), SOTINVH1, "*")
            Create_TDA(.Tables.Add("SOTINVH2"), SOTINVH2, "*")
            Create_TDA(.Tables.Add("SOTINVH4"), SOTINVH4, "*")
            Create_TDA(.Tables.Add("SOTINVHD"), SOTINVHD, "*")

            Create_TDA(.Tables.Add("SOTORDR1"), SOTORDR1, "*")
            Create_TDA(.Tables.Add("SOTORDR2"), SOTORDR2, "*")
            Create_TDA(.Tables.Add("SOTORDR5"), SOTORDR5, "*")

            ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY) AS ORDR_QTY" & vbCrLf _
                & ", SUM (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) AS ORDR_AMT" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE) AS ORDR_AMT_SHIP" & vbCrLf _
                & " from " & SOTORDR2 & " SOTORDR2, " & SOTINVH2 & " SOTINVH2" & vbCrLf _
                & " where SOTORDR2.ORDR_NO = SOTINVH2.ORDR_NO" & vbCrLf _
                & "   and SOTORDR2.ORDR_LNO = SOTINVH2.INV_LNO" & vbCrLf _
                & " group by SOTORDR2.ORDR_NO"
            Create_TDA(.Tables.Add, "SOTORDRT", "**", 0, False)

            ASCMAIN1.sql = "Select SOTINVH2.WHSE_CODE, ICTBRAN1.SALES_DIVISION_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_BRAND_CODE, SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", SUM (DECODE (SOTINVH2.INV_TYPE, 'I', SOTINVH2.ORDR_QTY_SHIP,0)) QTY_SHP" & vbCrLf _
                & ", SUM (DECODE (SOTINVH2.INV_TYPE, 'C', SOTINVH2.ORDR_QTY_SHIP,0)) QTY_RTN" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_UNIT_PRICE * DECODE (SOTINVH2.INV_TYPE, 'I', SOTINVH2.ORDR_QTY_SHIP,0)) AMT_SHP" & vbCrLf _
                & ", SUM (SOTINVH2.ORDR_UNIT_PRICE * DECODE (SOTINVH2.INV_TYPE, 'C', SOTINVH2.ORDR_QTY_SHIP,0)) AMT_RTN" & vbCrLf _
                & " from " & SOTINVH2 & " SOTINVH2, ICTITEM1, ICTBRAN1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTBRAN1.ITEM_BRAND_CODE = ICTITEM1.ITEM_BRAND_CODE" & vbCrLf _
                & " group by SOTINVH2.WHSE_CODE, ICTBRAN1.SALES_DIVISION_CODE, ICTITEM1.ITEM_BRAND_CODE, SOTINVH2.ITEM_CODE, ICTITEM1.ITEM_DESC"
            Create_TDA(.Tables.Add, "SOTINVHI", "**", 0, False, , 4)

            ASCMAIN1.sql = "SELECT SOTINVH1.WHSE_CODE, SOTINVH1.CUST_CODE" & vbCrLf _
                & ", SOTINVH1.ORDR_CUST_PO, SOTINVH4.TRACKING_NO" & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1, SOTINVH4" & vbCrLf _
                & " where SOTINVH1.INV_TYPE = SOTINVH4.INV_TYPE" & vbCrLf _
                & "   and SOTINVH1.INV_NO = SOTINVH4.INV_NO"
            Create_TDA(.Tables.Add, "SOTUPDTW", "*", 0, False, , 4)

            Create_TDA(.Tables.Add, "TATTERM1", "*", 0)
            Fill_Records("TATTERM1")

            Create_TDA(.Tables.Add, "ARTCUST6", "*")

            Create_TDA(.Tables.Add, "METCOOP1", "*")
            Create_TDA(.Tables.Add, "METCOOP2", "*")
            Create_TDA(.Tables.Add, "METCOOP3", "*")
            Create_TDA(.Tables.Add, "METCOOP4", "*")

            Create_TDA(.Tables.Add, "ARTOPEN1", "*")

            Create_TDA(.Tables.Add, "ICTTRNE1", "*")
            Create_TDA(.Tables.Add, "ICTTRNE2", "*")

            'ASCMAIN1.sql = "Select * from ICTSTAT2 WHERE ROWNUM < 1"
            'ICTSTAT2_BO = ASCMAIN1.Temp_Table
            'ASCMAIN1.sql = "Select * from " & ICTSTAT2_BO
            Stop
            'Create_TDA(.Tables.Add("ICTSTAT2_BO"), ICTSTAT2_BO, "*", 0, False, String.Empty, 2)

        End With

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Preparing Sales Journal", String.Empty)

        Create_Temp_Tables()

        For Each TABLE_NAME As String In New String() _
            {"SOTINVH1", "SOTINVH2", "SOTINVH4", "SOTINVHD", "SOTINVHW", _
             "SOTORDR1", "SOTORDR2", "SOTORDR5", "SOTORDRT", "SOTINVHI"}
            Fill_Records(TABLE_NAME)
        Next

    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = "Invoices Updated on " & DateTime.Now.ToString("MM/dd/yy")

        For Each RPT In New String() {"SORUPDT1", "SORUPDT2", "SORUPDT3", "SORUPDT4", "SORUPDT5"}
            Generate_Report(RPT, , SUBT)
        Next

    End Sub

    Sub Create_Temp_Tables(Optional initialize As Boolean = False)

        Dim sqlO As String = "Select SOTORDR1.* from " & SOTINVH1 & " SOTINVH1, SOTORDR1 where SOTORDR1.ORDR_NO = SOTINVH1.ORDR_NO"

        If initialize Then
            sqlSOTINVH1 = "Select SOTINVH1.* from SOTINVH1 where ORDR_UPDATED = '0'"
            SOTINVH1 = ASCMAIN1.Temp_Table(sqlSOTINVH1 & " and ROWNUM < 1")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " add Primary Key (INV_TYPE, INV_NO)")
            ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " MODIFY CUST_BILL_TO_CUST NOT NULL")

            sqlSOTINVH2 = "Select SOTINVH2.*, SOTINVH1.ORDR_NO" _
                & " from SOTINVH2," & SOTINVH1 & " SOTINVH1 " _
                & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" _
                & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO"
            SOTINVH2 = ASCMAIN1.Temp_Table(sqlSOTINVH2)
            ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " add Primary Key (INV_TYPE, INV_NO, INV_LNO)")

            sqlSOTINVHD = "Select SOTORDR2.* From SOTORDR2 where ORDR_NO in (Select ORDR_NO from " & SOTINVH1 & ")"
            SOTINVHD = ASCMAIN1.Temp_Table(sqlSOTINVHD)
            ASCDATA1.ExecuteSQL("Alter Table " & SOTINVHD & " add Primary Key (ORDR_NO, ITEM_CODE)")

            SOTORDR1 = ASCMAIN1.Temp_Table(Replace(sqlO, "SOTORDR1", "SOTORDR1") & " and ROWNUM < 1")
            SOTORDR2 = ASCMAIN1.Temp_Table(Replace(sqlO, "SOTORDR1", "SOTORDR2") & " and ROWNUM < 1")
            SOTORDR5 = ASCMAIN1.Temp_Table(Replace(sqlO, "SOTORDR1", "SOTORDR5") & " and ROWNUM < 1")

            SOTINVH4 = ASCMAIN1.Temp_Table("Select * from SOTINVH4 where ROWNUM < 1")
        Else
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVH1 & " " & sqlSOTINVH1)
            ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " set CURR_EXCH_RATE = 1 where CURR_CODE = '" & GL_PARM_CURR_CODE & "'")
            ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " set CURR_EXCH_RATE = (Select CURR_EXCH_CUR from ICTCURR1 where CURR_CODE = " & SOTINVH1 & ".CURR_CODE) where CURR_CODE <> '" & GL_PARM_CURR_CODE & "'")

            ASCDATA1.ExecuteSQL("Insert into " & SOTINVH2 & " " & sqlSOTINVH2)
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHD & " " & sqlSOTINVHD)


            ASCMAIN1.sql = "" _
                & "Begin" & vbcrlf _
                & " Declare CURSOR C1 is" & vbcrlf _
                & "  Select * FROM " & SOTUPDT1 & " FOR UPDATE;" & vbcrlf _
                & " Begin" & vbcrlf _
                & "  For R1 in C1 Loop" & vbcrlf _
                & "   Update " & SOTUPDT1 & " Set INV_SALES =" & vbcrlf _
                & "    (Select SUM(ORDR_QTY_SHIP * ORDR_UNIT_PRICE) from " & SOTUPDT2 & vbcrlf _
                & "     where INV_TYPE = R1.INV_TYPE AND INV_NO = R1.INV_NO)" & vbcrlf _
                & "   where CURRENT OF C1;" & vbcrlf _
                & "  End Loop;" & vbcrlf _
                & " End;" & vbcrlf _
                & "End;"
            ASCDATA1.ExecuteSQL(sql)

            ASCDATA1.ExecuteSQL("Insert into " & SOTORDR1 & " " & Replace(sqlO, "SOTORDR1", "SOTORDR1"))
            ASCDATA1.ExecuteSQL("Insert into " & SOTORDR2 & " " & Replace(sqlO, "SOTORDR1", "SOTORDR2"))
            ASCDATA1.ExecuteSQL("Insert into " & SOTORDR5 & " " & Replace(sqlO, "SOTORDR1", "SOTORDR5"))
        End If
    End Sub

    Sub Update_Invoices()

        Dim CUST_CODE As String = String.Empty
        Dim CUST_BILL_TO_CUST As String = String.Empty
        Dim rowARTOPEN1 As DataRow = Nothing
        Dim rowARTCUST1 As DataRow = Nothing

        Try
            Dim SO_PARM_REASON_CODE_INV As String = ROWs("SOTPARM1").Item("SO_PARM_REASON_CODE_INV") & String.Empty

            ASCMAIN1.Progress("Updating Invoice", String.EMPTY)
            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select(String.EMPTY, "CUST_CODE, INV_NO")

                ASCMAIN1.Progress("-", INV_NO)

                ' Customer Break
                If CUST_CODE <> rowSOTINVH1.Item("CUST_CODE") & String.Empty _
                    OrElse CUST_BILL_TO_CUST <> rowSOTINVH1.Item("CUST_BILL_TO_CUST") & String.Empty Then

                    ' Set New Customer Values
                    CUST_CODE = rowSOTINVH1.Item("CUST_CODE") & String.Empty
                    CUST_BILL_TO_CUST = rowSOTINVH1.Item("CUST_CODE") & String.Empty

                    rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                End If

                Dim INV_NO As String = rowSOTINVH1.Item("INV_NO") & String.Empty
                Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE") & String.Empty
                Dim INV_DATE As Date = rowSOTINVH1.Item("INV_DATE") & String.Empty

                Dim INV_NO_CONS As String = rowSOTINVH1.Item("INV_NO_CONS") & String.Empty
                Dim INV_TOTAL_AMOUNT As Decimal = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty)
                Dim INV_SALES As Decimal = Val(rowSOTINVH1.Item("INV_SALES") & String.Empty)
                Dim SALES_DIVISION_CODE As String = rowSOTINVH1.Item("SALES_DIVISION_CODE") & String.Empty

                Dim CURR_CODE As String = rowSOTINVH1.Item("CURR_CODE") & String.Empty
                Dim CURR_EXCH_RATE As Decimal = Val(rowSOTINVH1.Item("CURR_EXCH_RATE") & String.Empty)
                Dim ORDR_XFR_TO_WHSE_CODE As String = rowSOTINVH1.Item("ORDR_XFR_TO_WHSE_CODE") & String.Empty
                Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty

                Dim ORDR_STATUS As String = "F"
                If rowSOTINVH1.Item("INV_CANCEL") & String.EMPTY = "1" Then
                    ORDR_STATUS = "C"
                End If

                If INV_TYPE = "C" Then Stop ' WJZ TO WALK THE 1ST ONE THRU - these should not happen here anymore - they should occur in christian's screen

                Dim TERM_CODE As String = rowSOTINVH1.Item("TERM_CODE") & String.Empty
                Dim INV_DUE_DATE As Date = TAC.SOCMAIN1.Calculate_INV_DUE_DATE(Me, TERM_CODE, INV_DATE)

                If INV_TYPE = "I" Then
                    Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
                    If rowSOTORDR1 Is Nothing Then
                        rowSOTORDR1 = Fill_Record("SOTORDR1", ORDR_NO, , False)
                        Fill_Records("SOTORDR2", ORDR_NO, False)
                    End If
                    If rowSOTORDR1.Item("ORDR_STATUS") & String.Empty <> "R" Then
                        Stop ' TRYING TO UPDATE AN ORDER WHICH IS NOT "RELEASED"
                    End If

                    rowSOTORDR1.Item("ORDR_STATUS") = ORDR_STATUS
                    rowSOTORDR1.Item("ORDR_YYYYPP_UPDATED") = rowSOTINVH1.Item("ORDR_YYYYPP_UPDATED")
                    rowSOTORDR1.Item("LAST_DATE") = DATETIME_STAMP

                    If rowSOTORDR1.Item("CURR_CODE") & String.Empty <> GL_PARM_CURR_CODE Then
                        If Math.Round(CURR_EXCH_RATE, 4) <> Math.Round(Val(rowSOTORDR1.Item("CURR_EXCH_RATE") & String.Empty), 4) Then
                            Stop ' SOTINVH1.CURR_EXCH_RATE <> SOTORDR1.CURR_EXCH_RATE ' PROBABLY NEED TO RESTATE USD EQUIVALENTS SO THAT SALES / AR AMTS ARE ACCURATE - OTHERWISE AR_CURR AMTS WILL BE INCORRECT
                        End If
                    End If

                    If rowSOTORDR1.Item("ORDR_XFR_TO_WHSE_CODE") & String.Empty <> String.Empty Then
                        Setup_Transfer(INV_DATE, transfers)
                    End If

                    ' Update A/R

                    If INV_TOTAL_AMOUNT <> 0 Then
                        If INV_NO_CONS <> String.Empty Then
                            rowARTOPEN1 = dst.Tables("ARTOPEN1").Rows.Find _
                                (New Object() {CUST_BILL_TO_CUST, INV_TYPE, INV_NO_CONS})
                            If rowARTOPEN1 IsNot Nothing Then
                                Update_ARTOPEN1_Cons(rowSOTINVH1, rowARTOPEN1, CURR_EXCH_RATE)
                            Else
                                Update_ARTOPEN1(rowSOTINVH1, CURR_EXCH_RATE)
                            End If
                        Else
                            Update_ARTOPEN1(rowSOTINVH1, CURR_EXCH_RATE)
                        End If
                    End If

                End If
            Next


            Try
                MyBase.BeginTrans()

                ASCMAIN1.Progress("Inventory Update", String.EMPTY)
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is" & vbCrLf _
                    & "  Select ITEM_CODE, WHSE_CODE" & vbCrLf _
                    & "  , SUM (ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                    & "  , SUM (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                    & "  , SUM (ORDR_QTY_RTRN) ORDR_QTY_RTRN FROM (" & vbCrLf _
                    & "   Select SOTINVH2.ITEM_CODE, SOTINVH2.WHSE_CODE" & vbCrLf _
                    & "   , 0 ORDR_QTY_PICK" & vbCrLf _
                    & "   , SUM (DECODE (SOTINVH2.INV_TYPE,'I',SOTINVH2.ORDR_QTY_SHIP,0)) ORDR_QTY_SHIP" & vbCrLf _
                    & "   , SUM (DECODE (SOTINVH2.INV_TYPE,'C',SOTINVH2.ORDR_QTY_SHIP,0)) ORDR_QTY_RTRN" & vbCrLf _
                    & "    from " & SOTINVH2 & " SOTINVH2, " & SOTINVH1 & " SOTINVH1 " & vbCrLf _
                    & "    where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE " & vbCrLf _
                    & "      and SOTINVH1.INV_NO = SOTINVH2.INV_NO " & vbCrLf _
                    & "    group by SOTINVH2.ITEM_CODE, SOTINVH2.WHSE_CODE" & vbCrLf _
                    & "   union" & vbCrLf _
                    & "    Select SOTORDR2.ITEM_CODE, SOTORDR2.WHSE_CODE" & vbCrLf _
                    & "    , SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
                    & "    , 0 ORDR_QTY_SHIP" & vbCrLf _
                    & "    , 0 ORDR_QTY_RTRN" & vbCrLf _
                    & "    from SOTORDR2, " & SOTINVH1 & " SOTINVH1" & vbCrLf _
                    & "    where SOTORDR2.ORDR_NO = SOTINVH1.ORDR_NO" & vbCrLf _
                    & "    group by SOTORDR2.ITEM_CODE, SOTORDR2.WHSE_CODE" & vbCrLf _
                    & "  ) where ORDR_QTY_PICK <> 0 OR ORDR_QTY_SHIP <> 0 OR ORDR_QTY_RTRN <> 0" & vbCrLf _
                    & "  group by ITEM_CODE, WHSE_CODE;" & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update ICTSTAT2 Set" & vbCrLf _
                    & "     WHSE_QTY_PICK = NVL(WHSE_QTY_PICK,0) - NVL(R1.ORDR_QTY_PICK,0)" & vbCrLf _
                    & "   , WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) - NVL(R1.ORDR_QTY_SHIP,0) - NVL(R1.ORDR_QTY_RTRN, 0)" & vbCrLf _
                    & "   where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
                    & "   If SQL%NOTFOUND Then" & vbCrLf _
                    & "    Insert into ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_PICK, WHSE_QTY_ON_HAND)" & vbCrLf _
                    & "     values (R1.ITEM_CODE, R1.WHSE_CODE " & vbCrLf _
                    & "      , -1 * NVL(R1.ORDR_QTY_PICK,0)" & vbCrLf _
                    & "      , -1 * (NVL(R1.ORDR_QTY_SHIP,0) + NVL(R1.ORDR_QTY_RTRN,0)));" & vbCrLf _
                    & "   End If;" & vbCrLf _
                    & "   Update ICTSTAT1 Set" & vbCrLf _
                    & "     WHSE_QTY_SHP = NVL(WHSE_QTY_SHP,0) + NVL(R1.ORDR_QTY_SHIP,0)" & vbCrLf _
                    & "   , WHSE_QTY_RTN = NVL(WHSE_QTY_RTN,0) - NVL(R1.ORDR_QTY_RTRN, 0)" & vbCrLf _
                    & "   where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE" & vbCrLf _
                    & "   and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
                    & "   If SQL%NOTFOUND Then" & vbCrLf _
                    & "    Insert into ICTSTAT1 (ITEM_CODE, WHSE_CODE, OPS_YYYYPP, WHSE_QTY_SHP, WHSE_QTY_RTN)" & vbCrLf _
                    & "     values (R1.ITEM_CODE, R1.WHSE_CODE, '" & ASCMAIN1.CYP & "'," & vbCrLf _
                    & "      NVL(R1.ORDR_QTY_SHIP,0), -1 * NVL(R1.ORDR_QTY_RTRN,0));" & vbCrLf _
                    & "   End If;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL(sql)

                ' Update Order Details
                ASCMAIN1.Progress("Order Details Update", String.EMPTY)
                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is " & vbCrLf _
                    & " Select SOTORDR1.* from SOTORDR1, " & SOTINVH1 & " SOTINVH1" & vbCrLf _
                    & " where SOTINVH1.ORDR_NO = SOTORDR1.ORDR_NO;" & vbCrLf _
                    & " Begin for R1 in C1 Loop" & vbCrLf _
                    & "  Update SOTORDR2 set " & vbCrLf _
                    & "    ORDR_STATUS = R1.ORDR_STATUS" & vbCrLf _
                    & "  , ORDR_YYYYPP_UPDATED = R1.ORDR_YYYYPP_UPDATED" & vbCrLf _
                    & "  , ORDR_QTY_PICK = 0" & vbCrLf _
                    & "    where ORDR_NO = R1.ORDR_NO;" & vbCrLf _
                    & " End Loop; End; End;"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "" & vbCrLf _
                    & "Begin Declare Cursor C1 is " & vbCrLf _
                    & " Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.ITEM_CODE" & vbCrLf _
                    & ", SOTUPDT2.ORDR_QTY_SHIP, SOTORDR2.ORDR_QTY_PICK" & vbCrLf _
                    & "  from " & SOTINVH1 & " SOTINVH1, " & SOTINVH2 & " SOTINVH2, SOTORDR2" & vbCrLf _
                    & " where SOTUPDT1.INV_TYPE = SOTUPDT2.INV_TYPE" & vbCrLf _
                    & "   and SOTUPDT1.INV_NO = SOTUPDT2.INV_NO" & vbCrLf _
                    & "   and SOTUPDT2.INV_TYPE = 'I'" & vbCrLf _
                    & "   and SOTORDR2.ORDR_NO = SOTUPDT1.ORDR_NO" & vbCrLf _
                    & "   and SOTORDR2.ITEM_CODE = SOTUPDT2.ITEM_CODE;" & vbCrLf _
                    & " Begin for R1 in C1 Loop" & vbCrLf _
                    & "  Update SOTORDR2 set " & vbCrLf _
                    & "    ORDR_QTY_SHIP = NVL(ORDR_QTY_SHIP,0) + NVL(R1.ORDR_QTY_SHIP,0)" & vbCrLf _
                    & "  , ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + (NVL(R1.ORDR_QTY_PICK,0) - NVL(R1.ORDR_QTY_SHIP,0))" & vbCrLf _
                    & "  where ORDR_NO = R1.ORDR_NO and ORDR_LNO = R1.ORDR_LNO;" & vbCrLf _
                    & " End Loop; End; End;"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.Progress("Sales History", String.EMPTY)
                ASCMAIN1.sql = "Update SOTINVH1 set INIT_DATE = SYSDATE, INIT_OPER = '" & ASCMAIN1.USER_ID & "', REGISTER_XNO = '" & XNO & "', ORDR_UPDATED = '1'" & vbCrLf _
                    & " where INV_TYPE = 'I' and INV_NO in (Select INV_NO from " & SOTINVH1 & ")"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "Update SOTSHIP1 set LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "', REGISTER_XNO = '" & XNO & "'" & vbCrLf _
                    & " where SHIPMENT_NO in (Select SHIPMENT_NO from " & SOTINVH1 & ")"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_810_BATCH_NO = 'N'" & vbCrLf _
                    & " where SHIPMENT_NO in (Select SHIPMENT_NO from " & SOTINVH1 & ")" _
                    & "   and SHIP_810_BATCH_NO is Null"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_810_BATCH_NO = '0000000000'" & vbCrLf _
                    & " where SHIPMENT_NO in (Select SHIPMENT_NO from " & SOTINVH1 & vbCrLf _
                    & " where CUST_CODE in " & vbCrLf _
                    & " (Select CUST_CODE from EDTTRPM1 where EDI_DOC_NO = '810'))" & vbCrLf _
                    & "   and SHIP_810_BATCH_NO = 'N'"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_856_BATCH_NO = 'N'" & vbCrLf _
                    & " where SHIPMENT_NO in (Select SHIPMENT_NO from " & SOTINVH1 & ")"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "Update SOTSHIP1 set SHIP_856_BATCH_NO = '0000000000'" & vbCrLf _
                    & " where SHIPMENT_NO in (Select SHIPMENT_NO from " & SOTINVH1 & vbCrLf _
                    & " where CUST_CODE in (Select CUST_CODE from EDTTRPM1 " & vbCrLf _
                    & " where EDI_DOC_NO = '856'))"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "" _
                    & "Begin Declare Cursor C1 is Select Distinct ORDR_GROUP_NO " & vbCrLf _
                    & " from " & SOTINVH1 & "SOTUPDT1, SOTORDR1 where SOTORDR1.ORDR_NO = SOTUPDT1.ORDR_NO;" & vbCrLf _
                    & " Begin For R1 in C1 Loop" & vbCrLf _
                    & " SOPORDR0_G(R1.ORDR_GROUP_NO);" & vbCrLf _
                    & " End Loop; End; End;"
                ASCDATA1.ExecuteSQL(sql)

                ASCMAIN1.sql = "" _
                    & "Select 'S' STBT, CUST_CODE" & vbCrLf _
                    & ", SUM (DECODE(INV_TYPE,'I',INV_SALES,0)) SALES_AMT" & vbCrLf _
                    & ", SUM (DECODE(INV_TYPE,'C',INV_SALES,0)) RTRNS_AMT" & vbCrLf _
                    & ", SUM (DECODE(INV_TYPE,'I',1,0)) SALES_CNT" & vbCrLf _
                    & ", SUM (DECODE(INV_TYPE,'C',1,0)) RTRNS_CNT" & vbCrLf _
                    & ", MIN (DECODE(INV_TYPE,'I',INV_NO,NULL)) INV_NO_FIRST" & vbCrLf _
                    & ", MAX (DECODE(INV_TYPE,'I',INV_NO,NULL)) INV_NO_LAST" & vbCrLf _
                    & " from " & SOTINVH1 & " where NVL(INV_CANCEL,'0') <> '1'" & vbCrLf _
                    & "   and NVL(INV_SALES,0) <> 0" & vbCrLf _
                    & " group by CUST_CODE" & vbCrLf _
                    & " Union" & vbCrLf _
                    & "Select 'B' STBT, CUST_BILL_TO_CUST CUST_CODE" & vbCrLf _
                    & " , SUM (DECODE(INV_TYPE,'I',INV_SALES,0)) SALES_AMT" & vbCrLf _
                    & " , SUM (DECODE(INV_TYPE,'C',INV_SALES,0)) RTRNS_AMT" & vbCrLf _
                    & " , SUM (DECODE(INV_TYPE,'I',1,0)) SALES_CNT" & vbCrLf _
                    & " , SUM (DECODE(INV_TYPE,'C',1,0)) RTRNS_CNT" & vbCrLf _
                    & " , MIN (DECODE(INV_TYPE,'I',INV_NO,NULL)) INV_NO_FIRST" & vbCrLf _
                    & " , MAX (DECODE(INV_TYPE,'I',INV_NO,NULL)) INV_NO_LAST" & vbCrLf _
                    & " From " & SOTINVH1 & vbCrLf _
                    & " where NVL(INV_CANCEL,'0') <> '1'" & vbCrLf _
                    & "   and CUST_CODE <> CUST_BILL_TO_CUST" & vbCrLf _
                    & "   and NVL(INV_SALES,0) <> 0" & vbCrLf _
                    & " group by BY CUST_BILL_TO_CUST"

                For Each rowARTCUST6_DATA As DataRow In ASCDATA1.GetDataTable(sql).Select(String.EMPTY)
                    ' get the most recent invpoice for the customer
                    Dim INV_NO As String = dst.Tables("SOTINVH1").Compute("MAX(INV_NO)", "CUST_CODE = '" & rowARTCUST6_DATA.Item("CUST_CODE") & "'") & String.Empty
                    Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Select("INV_NO = '" & INV_NO & "'")(0)
                    Update_ARTCUST6(rowSOTINVH1, rowARTCUST6_DATA)
                Next

                Process_Back_Orders()

                Update_Record_TDA("SOTINVH1")
                Update_Record_TDA("SOTINVH2")
                Update_Record_TDA("SOTINVH4")

                Update_Record_TDA("SOTORDR1")
                Update_Record_TDA("SOTORDR2")

                Update_Record_TDA("ARTOPEN1")
                Update_Record_TDA("ARTCUST6")

                MyBase.CommitTrans("Update Successful")

            Catch ex As Exception
                MyBase.Rollback("Update Error: " & ex.Message)
                RWU = "N"
            End Try

        Catch ex As Exception
            MyBase.Rollback("Update Error: " & ex.Message)
            RWU = "N"
        Finally
            ASCMAIN1.Progress(String.Empty, String.Empty)
        End Try

    End Sub

    Private Sub Update_ARTCUST6(ByRef rowSOTINVH1 As DataRow, ByRef rowARTCUST6_DATA As DataRow)



        Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE") & String.Empty
        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO") & String.Empty

        Dim rowARTCUST1 As DataRow = LookUp("ARTCUST1", CUST_CODE)

        Dim rowARTCUST6 As DataRow = dst.Tables("ARTCUST6").Rows.Find(CUST_CODE)
        If rowARTCUST6 Is Nothing Then
            rowARTCUST6 = Fill_Record("ARTCUST6", CUST_CODE, False, )
            If rowARTCUST6 Is Nothing Then
                rowARTCUST6 = dst.Tables("ARTCUST6").NewRow
                rowARTCUST6.Item("CUST_CODE") = CUST_CODE
                dst.Tables("ARTCUST6").Rows.Add(rowARTCUST6)
            End If
        End If

        If Not IsDate(rowARTCUST6.Item("INV_NO_LAST") & String.Empty) Then
            rowARTCUST6.Item("CUST_LAST_INV_NUM") = rowSOTINVH1.Item("INV_NO")
            rowARTCUST6.Item("CUST_LAST_INV_DATE") = rowSOTINVH1.Item("INV_DATE")
            rowARTCUST6.Item("CUST_LAST_INV_AMT") = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty)
            If Not IsDate(rowARTCUST6.Item("CUST_FIRST_PURCH") & String.Empty) Then
                rowARTCUST6.Item("CUST_FIRST_PURCH") = rowSOTINVH1.Item("INV_DATE")
            End If
        End If

        If rowARTCUST6.Item("CUST_BILL_TO_CUST") & String.Empty = CUST_CODE OrElse rowARTCUST6.Item("CUST_BILL_TO_CUST") & String.Empty = String.Empty Then
            sql = "Select Sum (INV_BALANCE) from ARTOPEN1 where CUST_CODE = :PARM1"
            Dim CUST_BALANCE As Decimal = Val(ASCDATA1.GetDataValue(sql, "V", New Object() {CUST_CODE}) & String.Empty)
            If CUST_BALANCE > Val(rowARTCUST6.Item("CUST_HIGH_BAL_AMT") & String.Empty) Then
                rowARTCUST6.Item("CUST_HIGH_BAL_DATE") = Format$(DATETIME_STAMP, "MM/dd/yyyy")
                rowARTCUST6.Item("CUST_HIGH_BAL_AMT") = CUST_BALANCE
            End If
        End If

        rowARTCUST6.Item("CUST_SALES_MTD") = Val(rowARTCUST6.Item("CUST_SALES_MTD") & String.EMPTY) + Val(rowARTCUST6_DATA.Item("SALES_AMT") & String.EMPTY)
        rowARTCUST6.Item("CUST_SALES_YTD") = Val(rowARTCUST6.Item("CUST_SALES_YTD") & String.EMPTY) + Val(rowARTCUST6_DATA.Item("SALES_AMT") & String.EMPTY)
        rowARTCUST6.Item("CUST_NUM_INV_MTD") = Val(rowARTCUST6.Item("CUST_NUM_INV_MTD") & String.EMPTY) + Val(rowARTCUST6_DATA.Item("SALES_CNT") & String.EMPTY)
        rowARTCUST6.Item("CUST_NUM_INV_YTD") = Val(rowARTCUST6.Item("CUST_NUM_INV_YTD") & String.EMPTY) + Val(rowARTCUST6_DATA.Item("SALES_CNT") & String.EMPTY)
        rowARTCUST6.Item("CUST_CRED_MTD") = Val(rowARTCUST6.Item("CUST_CRED_MTD") & String.EMPTY) - Val(rowARTCUST6_DATA.Item("RTRNS_AMT") & String.EMPTY)
        rowARTCUST6.Item("CUST_CRED_YTD") = Val(rowARTCUST6.Item("CUST_CRED_YTD") & String.EMPTY) - Val(rowARTCUST6_DATA.Item("RTRNS_AMT") & String.EMPTY)

    End Sub

    Private Sub Update_ARTOPEN1(ByRef rowSOTINVH1 As DataRow, ByVal CURR_EXCH_RATE As Decimal)

        Dim INV_NO_CONS As String = rowSOTINVH1.Item("INV_NO_CONS") & String.Empty
        Dim INV_NO As String = rowSOTINVH1.Item("INV_NO") & String.Empty
        Dim consolidated_invoice As Boolean = (INV_NO_CONS <> "")

        Dim rowARTOPEN1 As DataRow = dst.Tables("ARTOPEN1").NewRow
        With rowARTOPEN1
            .Item("CUST_CODE") = rowSOTINVH1.Item("CUST_BILL_TO_CUST")
            .Item("INV_TYPE") = rowSOTINVH1.Item("INV_TYPE")
            .Item("INV_NUM") = IIf(consolidated_invoice, INV_NO_CONS, INV_NO)
            .Item("INV_DATE") = rowSOTINVH1.Item("INV_DATE")
            .Item("CUST_STORE_NO") = IIf(consolidated_invoice, String.Empty, rowSOTINVH1.Item("CUST_STORE_NO"))
            .Item("POST_CODE") = rowSOTINVH1.Item("POST_CODE") & String.Empty
            .Item("TERM_CODE") = rowSOTINVH1.Item("TERM_CODE") & String.Empty
            .Item("INV_DUE_DATE") = rowSOTINVH1.Item("INV_DUE_DATE")
            .Item("INV_DISC_DATE") = rowSOTINVH1.Item("INV_DISC_DATE")
            .Item("SREP_CODE") = IIf(consolidated_invoice, String.Empty, rowSOTINVH1.Item("SREP_CODE"))
            .Item("STAX_CODE") = String.Empty
            .Item("INV_CUST_PO") = rowSOTINVH1.Item("ORDR_CUST_PO") & String.Empty
            .Item("ORDR_NO") = IIf(consolidated_invoice, String.Empty, rowSOTINVH1.Item("ORDR_NO"))
            .Item("INV_SALES") = Val(rowSOTINVH1.Item("INV_SALES") & String.Empty)
            .Item("INV_DISC") = 0
            .Item("INV_FREIGHT") = Val(rowSOTINVH1.Item("INV_FREIGHT") & String.Empty)
            .Item("INV_STAX") = 0
            .Item("INV_TOTAL_AMOUNT") = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty)
            .Item("INV_BALANCE") = Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty)
            .Item("CUST_CODE_SO") = rowSOTINVH1.Item("CUST_CODE")

            If rowSOTINVH1.Item("INV_TYPE") = "I" Then
                .Item("REASON_CODE") = ROWs("SOTPARM1").Item("SO_PARM_REASON_CODE_INV")
            Else
                .Item("REASON_CODE") = rowSOTINVH1.Item("REASON_CODE")
            End If
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = DATETIME_STAMP
            .Item("SALES_DIVISION_CODE") = rowSOTINVH1.Item("SALES_DIVISION_CODE")
            '.item("ITEM_BRAND_CODE") = dynSOWUPDT1.Fields("ITEM_BRAND_CODE").Value

            .Item("SEG2_CODE") = ROWs("ARTPARM1").Item("AR_PARM_DEF_SEG2")
            .Item("SEG3_CODE") = ROWs("ARTPARM1").Item("AR_PARM_DEF_SEG3")
            .Item("SEG4_CODE") = ROWs("ARTPARM1").Item("AR_PARM_DEF_SEG4")

            If rowSOTINVH1.Item("CURR_CODE") & String.Empty = GL_PARM_CURR_CODE Then
                .Item("INV_SALES_CURR") = .Item("INV_SALES").Value
                .Item("INV_DISC_CURR") = .Item("INV_DISC").Value
                .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT").Value
                .Item("INV_STAX_CURR") = .Item("INV_STAX").Value
                .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT").Value
                .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE").Value
            Else
                .Item("INV_SALES_CURR") = Val(.Item("INV_SALES").Value & String.Empty) / CURR_EXCH_RATE
                .Item("INV_DISC_CURR") = Val(.Item("INV_DISC").Value & String.Empty) / CURR_EXCH_RATE
                .Item("INV_FREIGHT_CURR") = Val(.Item("INV_FREIGHT").Value & String.Empty) / CURR_EXCH_RATE
                .Item("INV_STAX_CURR") = Val(.Item("INV_STAX").Value & String.Empty) / CURR_EXCH_RATE
                .Item("INV_TOTAL_AMOUNT_CURR") = Val(.Item("INV_TOTAL_AMOUNT").Value & String.Empty) / CURR_EXCH_RATE
                .Item("INV_BALANCE_CURR") = Val(.Item("INV_BALANCE").Value & String.Empty) / CURR_EXCH_RATE
            End If

            .Item("CURR_CODE") = rowSOTINVH1.Item("CURR_CODE")
            .Item("CURR_EXCH_RATE") = CURR_EXCH_RATE

            .Item("BUS_UNIT_CODE") = rowSOTINVH1.Item("BUS_UNIT_CODE")
        End With
     
        dst.Tables("ARTOPEN1").Rows.Add(rowARTOPEN1)
    End Sub

    Private Sub Update_ARTOPEN1_Cons(ByRef rowSOTINVH1 As DataRow, ByRef rowARTOPEN1 As DataRow, ByVal CURR_EXCH_RATE As Decimal)

        With rowARTOPEN1
            .Item("INV_SALES") = Val(.Item("INV_SALES") & String.Empty) + Val(rowSOTINVH1.Item("INV_SALES") & String.Empty)
            .Item("INV_DISC") = Val(.Item("INV_DISC") & String.Empty) + 0
            .Item("INV_FREIGHT") = Val(.Item("INV_FREIGHT") & String.Empty) + Val(rowSOTINVH1.Item("INV_FREIGHT") & String.Empty)
            .Item("INV_STAX") = Val(.Item("INV_STAX") & String.Empty) + 0
            .Item("INV_TOTAL_AMOUNT") = Val(.Item("INV_TOTAL_AMOUNT") & String.Empty) + Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty)
            .Item("INV_BALANCE") = Val(.Item("INV_BALANCE") & String.Empty) + Val(rowSOTINVH1.Item("INV_TOTAL_AMOUNT") & String.Empty)

            If rowSOTINVH1.Item("CURR_CODE") & String.Empty = ROWs("ARTPARM1").Item("AR_PARM_CURR_CODE") & String.Empty Then
                .Item("INV_SALES_CURR") = .Item("INV_SALES")
                .Item("INV_DISC_CURR") = .Item("INV_DISC")
                .Item("INV_FREIGHT_CURR") = .Item("INV_FREIGHT")
                .Item("INV_STAX_CURR") = .Item("INV_STAX")
                .Item("INV_TOTAL_AMOUNT_CURR") = .Item("INV_TOTAL_AMOUNT")
                .Item("INV_BALANCE_CURR") = .Item("INV_BALANCE")
            Else
                .Item("INV_SALES_CURR") = Val(.Item("INV_SALES") & String.Empty) / CURR_EXCH_RATE
                .Item("INV_DISC_CURR") = Val(.Item("INV_DISC") & String.Empty) / CURR_EXCH_RATE
                .Item("INV_FREIGHT_CURR") = Val(.Item("INV_FREIGHT") & String.Empty) / CURR_EXCH_RATE
                .Item("INV_STAX_CURR") = Val(.Item("INV_STAX") & String.Empty) / CURR_EXCH_RATE
                .Item("INV_TOTAL_AMOUNT_CURR") = Val(.Item("INV_TOTAL_AMOUNT") & String.Empty) / CURR_EXCH_RATE
                .Item("INV_BALANCE_CURR") = Val(.Item("INV_BALANCE") & String.Empty) / CURR_EXCH_RATE
            End If
        End With
    End Sub

    Sub Process_Back_Orders()
        Stop
        ' IF WE EVER GET TO ENTERING BACK ORDER QTY'S IN SALES ORDER ENTRY, THEN WE NEED TO SET UP ORDR_QTY_BACK IN THE UPDATE_INVOICES ROUTINE, MERGING FLOW WITH LINES THAT HAVE VALUES ALREADY LOADED INTO ORDR_QTY_BACK, AND THEN USE THOSE VALUES HERE, RATHER THAN UPDATING ORDR_QTY_BACK HERE
        ' ALSO, AT THAT TIME, WE WOULD NEED TO GET ORDER RELEASE TO USE THE ORDR_QTY_BACK FIELD (INSTEAD OF ORDR_QTY_CANC)

        Stop ' NEED TO ANALYZE ROUTINE AROUND ORDR2

        ASCMAIN1.sql = "" _
            & "Select SOTORDR1.*" & vbCrLf _
            & " from SOTORDR1," & vbCrLf _
            & "       (Select DISTINCT SOTORDR2.ORDR_NO" & vbCrLf _
            & "         from SOTORDR2, " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & "          WHERE SOTINVH1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "            AND SOTORDR2.ORDR_QTY > SOTORDR2.ORDR_QTY_SHIP" & vbCrLf _
            & "            AND NVL(SOTORDR2.ORDR_RELEASE,'A') <> 'D'" & vbCrLf _
            & "  ) X" & vbCrLf _
            & "   where SOTORDR1.ORDR_NO = X.ORDR_NO" & vbCrLf _
            & "     and SOTORDR1.ORDR_BACKORDER = '1'"
        Dim tblSOWORDR1_OLD As DataTable = ASCDATA1.GetDataTable(sql)

        ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
            & " from SOTORDR2," & vbCrLf _
            & "      (Select DISTINCT SOTORDR1.ORDR_NO" & vbCrLf _
            & "        from SOTORDR1, " & SOTINVH1 & " SOTINVH1" & vbCrLf _
            & "         where SOTUPDT1.ORDR_NO = SOTORDR1.ORDR_NO" & vbCrLf _
            & "           and SOTORDR1.ORDR_BACKORDER = '1') X" & vbCrLf _
            & "  where SOTORDR2.ORDR_NO = X.ORDR_NO" & vbCrLf _
            & "    and SOTORDR2.ORDR_QTY > SOTORDR2.ORDR_QTY_SHIP" _
            & "    and NVL(SOTORDR2.ORDR_RELEASE,'A') <> 'D'"
        Dim tblSOWORDR2_OLD As DataTable = ASCDATA1.GetDataTable(sql)

        For Each rowSOTORDR1BK As DataRow In tblSOWORDR1_OLD.Rows
            Dim ORDR_NO_ORIG As String = rowSOTORDR1BK.Item("ORDR_NO") & String.Empty
            Dim ORDR_NO As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")

            ' Create SOTORDR5 entry for back order
            Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
            Dim rowSOTORDR5_ORIG As DataRow = dst.Tables("SOTORDR5").Select("ORDR_NO = '" & ORDR_NO_ORIG & "'")(0)
            For Each dc As DataColumn In dst.Tables("SOTORDR5").Columns
                rowSOTORDR5.Item(dc.ColumnName) = rowSOTORDR5_ORIG.Item(dc.ColumnName)
            Next
            dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)

            Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").NewRow
            Dim rowSOTORDR1_ORIG As DataRow = dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO_ORIG & "'")(0)

            For Each dc As DataColumn In dst.Tables("SOTORDR1").Columns
                rowSOTORDR1.Item(dc.ColumnName) = rowSOTORDR1_ORIG.Item(dc.ColumnName)
            Next

            With rowSOTORDR1
                .Item("ORDR_NO") = ORDR_NO
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                .Item("LAST_OPER") = ASCMAIN1.USER_ID
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("LAST_DATE") = DATETIME_STAMP
                .Item("ORDR_CRED_HOLD_CODES") = String.Empty
                .Item("ORDR_CRED_CLEARED") = String.Empty
                .Item("ORDR_CRED_CLR_BY") = String.Empty
                .Item("ORDR_CRED_CLR_AUTH") = String.Empty
                .Item("ORDR_CRED_CLR_DATE") = DBNull.Value
                .Item("ORDR_REL_HOLD_CODES") = String.Empty
                .Item("ORDR_REL_BATCH_NO") = String.Empty
                .Item("ORDR_DATE_REL") = DBNull.Value
                .Item("ORDR_STATUS") = "O"
                .Item("ORDR_YYYYPP_UPDATED") = String.Empty
                .Item("REORD_MEMO_IND") = String.Empty
                .Item("ORDR_BATCHED") = String.Empty
                .Item("ORDR_PICK_SEQ") = rowSOTORDR1_ORIG.Item("ORDR_PICK_SEQ") + 1
                .Item("ORDR_OVERRIDE_NOT_ALLOCATED") = String.Empty
                .Item("ORDR_INVOICED") = String.Empty
                .Item("ORDR_NO_ORIG") = ORDR_NO_ORIG
            End With
            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

            Dim ORDR_LNO As Int16 = 0
            For Each row As DataRow In tblSOWORDR2_OLD.Select("ORDR_NO = '" & ORDR_NO_ORIG & "'")
                Dim rowSOTORDR2_ORIG As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO_ORIG, 0})

                Dim ORDR_LNO_ORIG As Int16 = row.Item("ORDR_LNO")
                ORDR_LNO += 1
                Dim ORDR_QTY As Int16 = Val(rowSOTORDR2_ORIG.Item("ORDR_QTY") & String.Empty) - Val(rowSOTORDR2_ORIG.Item("ORDR_QTY_SHIP") & String.Empty)

                rowSOTORDR2_ORIG.Item("ORDR_QTY_CANC").Value = Val(rowSOTORDR2_ORIG.Item("ORDR_QTY_CANC") & String.Empty) - ORDR_QTY
                rowSOTORDR2_ORIG.Item("ORDR_QTY_BACK").Value = ORDR_QTY

                Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
                For Each dc As DataColumn In dst.Tables("SOTORDR2").Columns
                    rowSOTORDR2.Item(dc.ColumnName) = rowSOTORDR2_ORIG.Item(dc.ColumnName)
                Next

                With rowSOTORDR2
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("ORDR_LNO") = ORDR_LNO
                    .Item("ORDR_QTY") = ORDR_QTY
                    .Item("ORDR_QTY_OPEN") = ORDR_QTY
                    .Item("ORDR_QTY_PICK") = 0
                    .Item("ORDR_QTY_SHIP") = 0
                    .Item("ORDR_QTY_CANC") = 0
                    .Item("ORDR_QTY_ORIG") = ORDR_QTY
                    .Item("ORDR_YYYYPP_UPDATED") = String.Empty
                    .Item("ORDR_QTY_ORIG_BACK") = ORDR_QTY
                    .Item("ORDR_LNO_ORIG") = ORDR_LNO_ORIG
                    .Item("ORDR_STATUS") = "O"
                End With
                dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                ' keep track of ICTSTAT2 WHSE_QTY_OPEN
                Dim WHSE_CODE As String = rowSOTORDR2_ORIG.Item("WHSE_CODE") & String.Empty
                Dim ITEM_CODE As String = rowSOTORDR2_ORIG.Item("ITEM_CODE") & String.Empty
                Dim rowICTSTAT2_BO As DataRow = Nothing

                If dst.Tables(ICTSTAT2_BO).Select("WHSE_CODE = '" & WHSE_CODE & "' and ITEM_CODE = '" & ITEM_CODE & "'").Length = 0 Then
                    rowICTSTAT2_BO = dst.Tables(ICTSTAT2_BO).NewRow
                    rowICTSTAT2_BO.Item("WHSE_CODE") = WHSE_CODE
                    rowICTSTAT2_BO.Item("ITEM_CODE") = ITEM_CODE
                    rowICTSTAT2_BO.Item("WHSE_QTY_OPEN") = 0
                    dst.Tables(ICTSTAT2_BO).Rows.Add(rowICTSTAT2_BO)
                Else
                    rowICTSTAT2_BO = dst.Tables(ICTSTAT2_BO).Select("WHSE_CODE = '" & WHSE_CODE & "' and ITEM_CODE = '" & ITEM_CODE & "'")(0)
                End If
                rowICTSTAT2_BO.Item("WHSE_QTY_OPEN") += ORDR_QTY
            Next
        Next

        Update_Record_TDA(ICTSTAT2_BO)
        ASCMAIN1.sql = "Begin Declare Cursor C1 is Select * From " & ICTSTAT2_BO & ";" & vbCrLf _
            & " Begin for R1 in C1 Loop" & vbCrLf _
            & "    Update ICTSTAT2 SET WHSE_QTY_OPEN = NVL(WHSE_QTY_OPEN, 0) + R1.WHSE_QTY_OPEN WHERE WHSE_CODE = R1.WHSE_CODE AND ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & " IF SQL%NOTFOUND THEN" & vbCrLf _
            & "  INSERT INTO ICTSTAT2 (ITEM_CODE, WHSE_CODE, WHSE_QTY_ON_HAND)" & vbCrLf _
            & "   VALUES (R1.ITEM_CODE, R1.WHSE_CODE, NVL(R1.WHSE_QTY_OPEN, 0)); " & vbCrLf _
            & " END IF;" _
            & " End Loop; End; End;"

        Dim ogList As List(Of String) = New List(Of String)
        Dim ogListNew As List(Of String) = New List(Of String)
        For Each row As DataRow In dst.Tables("SOTORDR1").Select("", "", DataViewRowState.Added)
            Dim ORDR_GROUP_NO_ORIG As String = row.Item("ORDR_GROUP_NO")

            If ogList.Contains(ORDR_GROUP_NO_ORIG) Then Continue For
            ogList.Add(ORDR_GROUP_NO_ORIG)

            Dim ORDR_GROUP_NO As String = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
            ogListNew.Add(ORDR_GROUP_NO)

            For Each rowOG As DataRow In dst.Tables("SOTORDR1").Select("ORDR_GROUP_NO = '" & ORDR_GROUP_NO_ORIG & "'", "", DataViewRowState.Added)
                rowOG.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
            Next
        Next

        Update_Record_TDA("SOTORDR1")
        Update_Record_TDA("SOTORDR2")
        Update_Record_TDA("SOTORDR5")

        For Each ORDR_GROUP_NO As String In ogListNew
            ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO"})
        Next

    End Sub

    Public Overrides Sub Update_Record()
        If chkInvoices.Checked Then
            Update_Invoices()
        End If
    End Sub
End Class