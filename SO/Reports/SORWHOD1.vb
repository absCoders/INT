Public Class SORWHOD1
    Dim ICTITEM1 As String = ""
    Dim ABCs As New Dictionary(Of String, Decimal)

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        sqlw &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sqlw &= SQL_in("ITEM_CATGY_CODE", "ICTITEM1.ITEM_CATGY_CODE")
        sqlw &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")

      

      

        RWU = "R"

        Prepare_dst(True, sqlw, RYP)

        Check_if_Empty("ICTITEM1")
    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        If ICTITEM1 = "" Then Create_Temp_Data(sqlw)

        With dst
            ASCMAIN1.sql = "Select ICTITEM1.* from " & ICTITEM1 & " ICTITEM1"
            Create_TDA(dst.Tables.Add("ICTITEM1"), ICTITEM1, "**", 0, True, , 1)
            With .Tables("ICTITEM1").Columns
                .Add("DEMAND_QTY", GetType(System.Int64), "ISNULL(FORECAST,0)+ISNULL(PROD_COM,0)+ISNULL(PLAN_COM,0)")
                .Add("DEMAND_AMT", GetType(System.Decimal), "DEMAND_QTY * ISNULL(ITEM_COST_STD,0)")
                .Add("DEMAND_PCT", GetType(System.Decimal))
                .Add("DEMAND_PCT_CUM", GetType(System.Decimal))
            End With

         

        End With

        Fill_Records("ICTCOLL1")
        Fill_Records("ICTBRAN1")


        Fill_Records("ICTCATG1")
        dst.Tables("ICTCATG1").Rows.Add(New String() {"*", "All Catgys"})
        dst.Tables("ICTCATG1").Rows.Add(New String() {"?", "Catgy Unknown"})

        Fill_Records("DPTABCP1")

        Dim ABC_INDEX As Int16 = 0
        Dim ABC_PCT_CUM As Decimal
        For Each rowDPTABCP1 As DataRow In dst.Tables("DPTABCP1").Select("", "ABC_CODE")
            ABC_INDEX += 1
            ABC_PCT_CUM += Val(rowDPTABCP1.Item("ABC_PCT_RANGE") & "")
            rowDPTABCP1.Item("ABC_INDEX") = ABC_INDEX
            rowDPTABCP1.Item("ABC_PCT_CUM") = ABC_PCT_CUM
        Next

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)

        If sqlw <> "" Then
            Create_Temp_Data(sqlw)
        End If
        EnforceConstraints(False)
        Fill_Records("ICTITEM1")
        EnforceConstraints(True)

        Calculate_ABC()

    End Sub

    Sub Create_Temp_Data(SQLW As String)

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_COST_STD" & vbCrLf _
            & ", ICTITEM1.ITEM_UOM, ICTITEM1.VEND_CODE, ICTITEM1.ITEM_PO_QTY_MIN" & vbCrLf _
            & ", ICTITEM1.ITEM_MRP_PLANR_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE || ICTITEM1.ITEM_BASIC_PROMO || ICTITEM1.ITEM_COST_MAKE_BUY || ICTITEM1.ITEM_CATGY_CODE ABC_GROUP" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_MAKE_BUY, ICTITEM1.ITEM_ABC_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_ABC_CODE ITEM_ABC_CODE_FUT" & vbCrLf _
            & " from ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & SQLW

        If ICTITEM1 = "" Then
            ICTITEM1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1 & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1 & " Add FORECAST NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1 & " Add PROD_COM NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1 & " Add PLAN_COM NUMBER (8,0)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1)

            Dim COLUMN_NAMEs As String = "" _
                & "ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, ITEM_COST_STD" _
                & ", ITEM_UOM, VEND_CODE, ITEM_PO_QTY_MIN, ITEM_MRP_PLANR_CODE, ABC_GROUP" _
                & ", ITEM_CATGY_CODE, COLLECTION_CODE, BRAND_CODE, ITEM_SNU_CODE, ITEM_BASIC_PROMO" _
                & ", ITEM_COST_MAKE_BUY, ITEM_ABC_CODE, ITEM_ABC_CODE_FUT"

            ASCDATA1.ExecuteSQL("Insert into " & ICTITEM1 & " (" & COLUMN_NAMEs & ") " & ASCMAIN1.sql)

            Get_Demand_Data()
        End If

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        CR_params.Add("MOS", Val(Absx1.numFor("FPDMOS").Value & ""))
        CR_params.Add("CD", "Demand Calculated using " & Absx1.optFor("OPTDEMAND").Text)
        CR_params.Add("EXTUSAGE", "Saleable Ranked by " & Absx1.optFor("OPTRANKBY").Text)
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then


        End If

    End Sub


    ' Application Specific
    Dim cf() As String
    Dim codes() As Object
    Dim cfmax As Integer
    Dim AZ As String
    Dim chkRecap As String
    Dim chkNewPage As String
    Dim chkThousands As String

    Dim RYP1 As String
    Dim RYP1Legend As String
    Dim RYP2 As String
    Dim RYP2Legend As String
    Dim dt1 As String
    Dim dt2 As String
    Dim chkOO As String
    Dim chkNOT As String
    Dim itemz As String

    Sub Build_Workfile()

        Call Build_WorkFile_DB_Init()

      
        Dim f As Integer
        Dim typ As String

        Dim CNT_ORD As Long
        Dim CNT_SHP As Long
        Dim CNT_REL As Long
        Dim CNT_OPN As Long
        Dim QTY_ORD As Long
        Dim QTY_SHP As Long
        Dim QTY_REL As Long
        Dim QTY_OPN As Long
        Dim AMT_ORD As Double
        Dim AMT_SHP As Double
        Dim AMT_REL As Double
        Dim AMT_OPN As Double

        Dim QTY As Long

        Dim ITEM_CODE As String
        Dim CUST_CODE As String
        Dim CUST_STORE_NO As String
        Dim ITEM_BRAND_CODE As String

       
        chkOO = SRead(opts, "CHKOO", 2)
        If chkOO = "1" Then
            Call Write_Page0("Including ALL Open Orders, regardless of Ship Date")
        End If

        chkNOT = SRead(opts, "CHKNOT", 2)
        If chkNOT = "1" Then
            Call Write_Page0("Showing Stores which did NOT order")
        End If


        ' Set up Work File Definition using X's and 0's and 0.01's as required

        Call Track("Initialize Work Tables", "")
        sql = "Select 0 RECORD_NO, CUST_CODE, ITEM_CODE, "
        sql = sql & " 0 ORD_COUNT, 0 SHP_COUNT, 0 REL_COUNT, 0 OPN_COUNT, "
        sql = sql & " 0 ORD_UNITS, 0 SHP_UNITS, 0 REL_UNITS, 0 OPN_UNITS, "
        sql = sql & " W_DBL ORD_SALES, W_DBL SHP_SALES, W_DBL REL_SALES, W_DBL OPN_SALES, "
        sql = sql & " 0 ALLO_UNITS, "
        sql = sql & Get_GString(1, 30, AZ)
        sql = sql & " from TATWORK1 where ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "ASWSRPT0", 1, "", sql)
        Call Create_Index("ASWSRPT0", "Report", Get_GString(2, 0, AZ))
        Dim dynASWSRPT0 As Recordset
        dynASWSRPT0 = AccD.OpenRecordset("ASWSRPT0", dbOpenTable)
        dynASWSRPT0.Index = "PrimaryKey"

        ' SALES_DIVISION_|ITEM_BRAND_|REGION_|SREP_|CUST_|CUST_BILL_TO_CUST|MARKET_|TRADE_CLASS_
        '|ITEM_|CUST_BUYING_GROUP|BUS_UNIT_|ITEM_CLASS_|ITEM_GROUP_|BUS_SEG_
        s9 = ColumnName_To_Index("ITEM_CODE")
        sql = "Select * from ICTITEM1 where ITEM_CODE " & Sel(2, s9) & " in (" & Sel(1, s9) & ")"
        Call Ora_to_Acc(Nothing, "ICWITEM1", 1, "", sql)

        sql = "Select CUST_CODE, CUST_STORE_NO from TATWORK1 where ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "SOWWHOD1", 2, "", sql)

        sql = "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE, 'X' ORDERED from TATWORK1 where ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "SOWWHOD2", 4, "", sql)
        Dim tblSOWWHOD2 As Recordset
        tblSOWWHOD2 = AccD.OpenRecordset("SOWWHOD2", dbOpenTable)
        tblSOWWHOD2.Index = "PrimaryKey"

        sql = "Select CUST_CODE, ITEM_CODE, SUBSTR(W_TXT,1,255) STORES from TATWORK1 where ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "SOWWHOD3", 2, "", sql)

        sql = "Select CUST_CODE, ITEM_BRAND_CODE from TATWORK1 where ROWNUM < 1"
        Call Ora_to_Acc(Nothing, "SOWWHOD4", 2, "", sql)
        Dim tblSOWWHOD4 As Recordset
        tblSOWWHOD4 = AccD.OpenRecordset("SOWWHOD4", dbOpenTable)
        tblSOWWHOD4.Index = "PrimaryKey"

        ' Set up Misc dynasets & memory tables

        RYP1Legend = SRead(opts, "CMBYP0", 2)
        RYP1 = Mid$(RYP1Legend, 1, 4) & Mid$(RYP1Legend, 6, 2)
        RYP2Legend = SRead(opts, "CMBYP1", 2)
        RYP2 = Mid$(RYP2Legend, 1, 4) & Mid$(RYP2Legend, 6, 2)

        z = Get_YYYYMM(RYP1, 0)
        dt1 = Format$(Mid$(z, 5, 2) & "/01/" & Mid$(z, 3, 2), "DD-MMM-YYYY")

        z = Get_YYYYMM(RYP2, 1)
        dt2 = Format$(DateAdd("d", -1, Mid$(z, 5, 2) & "/01/" & Mid$(z, 3, 2)), "DD-MMM-YYYY")

        ' Prepare Work File with Data from Server

        Dim dynX As OraDynaset
        Dim fldX() As Object

        ' Shipped, Open, Released Orders

        Dim t As Integer
        Call Track("Stores Who Ordered", "")

        For t = 1 To 2
            If t = 1 Then
                typ = "O"
            Else
                typ = "B"
            End If

            ReDim xInfo(1)
            Call Get_SQL(typ, cfmax, cf, codes, xInfo)
            itemz = xInfo(1)

            If objASCCALLB.xErrMsg <> "" Then
                Exit Sub
            End If

            'If chkMRGHIST = "1" Then
            '    sqltables = sqltables & ", " & MRGHIST_Table
            '    sqljoin = sqljoin & " and X.CUST_CODE = " & MRGHIST_Table & ".CUST_CODE_OLD"
            '    sqljoin = sqljoin & " and X.CUST_STORE_NO = " & MRGHIST_Table & ".CUST_STORE_NO_OLD"
            'End If

            For f = 1 To 2
                If t = 1 Then
                    sql = "Select " & sqllist & ", X.CUST_STORE_NO"
                Else
                    sql = "Select " & sqllist & ", I.ITEM_BRAND_CODE"
                End If
                sql = sql & ", Count (*) CNT_ORD"
                sql = sql & ", Count (DECODE(ORDR_STATUS,'O',1,0)) CNT_OPN"
                sql = sql & ", Count (DECODE(ORDR_STATUS,'R',1,0)) CNT_REL"
                sql = sql & ", Count (DECODE(ORDR_STATUS,'F',1,0)) CNT_SHP"
                sql = sql & ", Sum (ORDR_QTY)      QTY_ORD"
                sql = sql & ", Sum (ORDR_QTY_OPEN) QTY_OPN"
                sql = sql & ", Sum (ORDR_QTY_PICK) QTY_REL"
                sql = sql & ", Sum (ORDR_QTY_SHIP) QTY_SHP"
                sql = sql & ", Sum (ORDR_QTY      * ORDR_UNIT_PRICE) AMT_ORD"
                sql = sql & ", Sum (ORDR_QTY_OPEN * ORDR_UNIT_PRICE) AMT_OPN"
                sql = sql & ", Sum (ORDR_QTY_PICK * ORDR_UNIT_PRICE) AMT_REL"
                sql = sql & ", Sum (ORDR_QTY_SHIP * ORDR_UNIT_PRICE) AMT_SHP"
                If f = 1 Then
                    Call Track("-", "Shipments")
                Else
                    Call Track("-", "Open Orders")
                End If
                sql = sql & " from SOTORDR2 X" & sqltables
                If t = 2 Then
                    sql = sql & ", ICTITEM1 I"
                End If
                If f = 1 Then
                    sql = sql & " where ORDR_YYYYPP_UPDATED between '" & RYP1 & "' and '" & RYP2 & "'"
                    sql = sql & " and ORDR_QTY_SHIP <> 0"
                Else
                    sql = sql & " where ORDR_QTY <> 0 and (ORDR_STATUS = 'O' or ORDR_STATUS = 'R')"
                    If chkOO <> "1" Then
                        sql = sql & " and ORDR_SHIP_DATE between '" & dt1 & "' and '" & dt2 & "'"
                    End If
                End If
                If t = 2 Then
                    sql = sql & " and I.ITEM_CODE = X.ITEM_CODE"
                    sql = sql & " and I.ITEM_BRAND_CODE in "
                    sql = sql & "(Select ITEM_BRAND_CODE from ICTITEM1 where ITEM_CODE " & Sel(2, s9) & " in (" & Sel(1, s9) & "))"
                End If
                sql = sql & sqljoin
                sql = sql & sqlwhere
                If t = 1 Then
                    sql = sql & " group by " & sqllist2 & ", X.CUST_STORE_NO"
                Else
                    sql = sql & " group by " & sqllist2 & ", I.ITEM_BRAND_CODE"
                End If
                dynX = OraD.CreateDynaset(sql, 8&)
                Call Set_Fldsx(dynX, fldX())

                Do While Not dynX.EOF
                    If t = 1 Then
                        CNT_ORD = Val(dynX("CNT_ORD").Value & "")
                        CNT_SHP = Val(dynX("CNT_SHP").Value & "")
                        CNT_REL = Val(dynX("CNT_REL").Value & "")
                        CNT_OPN = Val(dynX("CNT_OPN").Value & "")

                        QTY_ORD = Val(dynX("QTY_ORD").Value & "")
                        QTY_SHP = Val(dynX("QTY_SHP").Value & "")
                        QTY_REL = Val(dynX("QTY_REL").Value & "")
                        QTY_OPN = Val(dynX("QTY_OPN").Value & "")

                        AMT_ORD = Val(dynX("AMT_ORD").Value & "")
                        AMT_SHP = Val(dynX("AMT_SHP").Value & "")
                        AMT_REL = Val(dynX("AMT_REL").Value & "")
                        AMT_OPN = Val(dynX("AMT_OPN").Value & "")

                        CUST_STORE_NO = dynX.Fields("CUST_STORE_NO").Value
                    Else
                        ITEM_BRAND_CODE = dynX.Fields("ITEM_BRAND_CODE").Value
                    End If
                    CUST_CODE = dynX.Fields("CUST_CODE").Value
                    If typ = "O" Then
                        ITEM_CODE = dynX.Fields("ITEM_CODE").Value
                    GoSub Write_Records
                    Else
                        sql = "Select ITEM_CODE from ICWITEM1 "
                        sql = sql & " where ITEM_BRAND_CODE = '" & ITEM_BRAND_CODE & "'"
                        dynWK = AccD.OpenRecordset(sql, dbOpenForwardOnly)
                        Do While Not dynWK.EOF
                            ITEM_CODE = dynWK.Fields("ITEM_CODE").Value
                        GoSub Write_Records
                            dynWK.MoveNext()
                        Loop
                        dynWK.Close()
                    End If
                    dynX.MoveNext()
                Loop
            Next f
        Next t


        ' Get Customer Allocations

        Call Track("Customer Allocations", "")

        typ = "A"
        ReDim xInfo(1)
        Call Get_SQL(typ, cfmax, cf, codes, xInfo)
        itemz = xInfo(1)
        If objASCCALLB.xErrMsg <> "" Then
            Exit Sub
        End If

        sql = "Select " & sqllist
        sql = sql & ", Sum (QTY_ALLO) QTY"
        sql = sql & " from SOTALLO1 Y1, SOTALLO2 Y2" & sqltables
        sql = sql & " where Y1.DATE_START <= '" & dt2 & "'"
        sql = sql & "   and Y1.DATE_END   >= '" & dt1 & "'"
        sql = sql & "   and Y2.ALLO_CTL_NO = Y1.ALLO_CTL_NO"
        sql = sql & sqljoin
        sql = sql & sqlwhere
        sql = sql & " group by " & sqllist2
        dynX = OraD.CreateDynaset(sql, 8&)
        Call Set_Fldsx(dynX, fldX())

        Do While Not dynX.EOF
            'QTY = Val(fldX("QTY").Value & "")
            QTY = Val(dynX.Fields("QTY").Value & "")
            CUST_CODE = dynX.Fields("CUST_CODE").Value
            ITEM_CODE = dynX.Fields("ITEM_CODE").Value
        GoSub Write_Records
            dynX.MoveNext()
        Loop

        ' All Stores

        typ = "C"
        ReDim xInfo(1)
        Call Get_SQL(typ, cfmax, cf, codes, xInfo)
        itemz = xInfo(1)

        'If chkMRGHIST = "1" Then
        '    sqltables = sqltables & ", " & MRGHIST_Table
        '    sqljoin = sqljoin & " and X.CUST_CODE = " & MRGHIST_Table & ".CUST_CODE_OLD"
        '    sqljoin = sqljoin & " and X.CUST_STORE_NO = " & MRGHIST_Table & ".CUST_STORE_NO_OLD"
        'End If

        If objASCCALLB.xErrMsg <> "" Then
            Exit Sub
        End If

        Call Track("Now Listing All Stores", "")

        sql = "Select DISTINCT X.CUST_CODE, X.CUST_STORE_NO"
        sql = sql & " from ARTCUST2 X" & sqltables
        sql = sql & " where CUST_STORE_STATUS = 'A'"
        sql = sql & sqljoin
        sql = sql & sqlwhere
        Call Ora_to_Acc(Nothing, "SOWWHOD1", 2, "", sql)

        ' Set up who Didn't Order

        Call Track("Now Crossing Off Those Who Did Order, by Item", "")

        dynWK = AccD.OpenRecordset("ICWITEM1", dbOpenForwardOnly)
        Do While Not dynWK.EOF
            ITEM_CODE = dynWK.Fields("ITEM_CODE").Value
            sql = "Insert into SOWWHOD2 "
            sql = sql & "Select SOWWHOD1.CUST_CODE, SOWWHOD1.CUST_STORE_NO, ICWITEM1.ITEM_CODE, '0' as ORDERED"
            sql = sql & " from SOWWHOD1,ICWITEM1,SOWWHOD4"
            sql = sql & " where SOWWHOD4.CUST_CODE = SOWWHOD1.CUST_CODE"
            sql = sql & "   and SOWWHOD4.ITEM_BRAND_CODE = ICWITEM1.ITEM_BRAND_CODE"
            AccD.Execute(sql)
            dynWK.MoveNext()
        Loop
        dynWK.Close()

        sql = "Select CUST_CODE, CUST_STORE_NO, ITEM_CODE"
        sql = sql & " from SOWWHOD2 where ORDERED = '1'"
        dynWK = AccD.OpenRecordset(sql, dbOpenForwardOnly)
        Do While Not dynWK.EOF
            CUST_CODE = dynWK.Fields("CUST_CODE").Value
            CUST_STORE_NO = dynWK.Fields("CUST_STORE_NO").Value
            ITEM_CODE = dynWK.Fields("ITEM_CODE").Value
            tblSOWWHOD2.Seek("=", CUST_CODE, CUST_STORE_NO, ITEM_CODE, "0")
            If Not tblSOWWHOD2.NoMatch Then
                tblSOWWHOD2.Delete()
            End If
            dynWK.MoveNext()
        Loop
        dynWK.Close()

        sql = "Insert into SOWWHOD3 Select DISTINCT CUST_CODE, ITEM_CODE, '" & String$(255, "x") & "' as STORES from SOWWHOD2"
        AccD.Execute(sql)
        Dim dynSOWWHOD2 As Recordset
        Dim dynSOWWHOD3 As Recordset
        dynSOWWHOD3 = AccD.OpenRecordset("SOWWHOD3", dbOpenDynaset)
        Do While Not dynSOWWHOD3.EOF
            dynSOWWHOD3.Edit()
            z = ""
            sql = "Select CUST_STORE_NO from SOWWHOD2 "
            sql = sql & " where CUST_CODE = '" & dynSOWWHOD3.Fields("CUST_CODE").Value & "'"
            sql = sql & "   and ITEM_CODE = '" & dynSOWWHOD3.Fields("ITEM_CODE").Value & "'"
            sql = sql & "   and ORDERED = '0'"
            dynSOWWHOD2 = AccD.OpenRecordset(sql, dbOpenForwardOnly)
            Do While Not dynSOWWHOD2.EOF
                z = z & "," & dynSOWWHOD2.Fields("CUST_STORE_NO").Value
                dynSOWWHOD2.MoveNext()
            Loop
            dynSOWWHOD2.Close()
            z = Mid$(z, 2)
            If Len(z) > 255 Then
                z = Mid$(z, 1, 251) & " ..."
            End If
            dynSOWWHOD3.Fields("STORES").Value = z
            dynSOWWHOD3.Update()
            dynSOWWHOD3.MoveNext()
        Loop


        ' Prepare Report File, w/Consolidations & Recaps as required

        Call Build_Report_File(3, 13, dynASWSRPT0, "ASWSRPT0", cfmax, AZ, "N", chkThousands, cf(), "CUST_CODE,ITEM_CODE")

        ' Erase List of Stores if No Store Ordered any

        sql = "Update ASWSRPT1,SOWWHOD3 set SOWWHOD3.STORES = NULL "
        sql = sql & " where ASWSRPT1.SUM_ORD_COUNT = 0"
        sql = sql & "   and SOWWHOD3.CUST_CODE = ASWSRPT1.CUST_CODE"
        sql = sql & "   and SOWWHOD3.ITEM_CODE = ASWSRPT1.ITEM_CODE"
        AccD.Execute(sql)

        ' Wrap up

        dynASWSRPT0.Close()
        tblSOWWHOD2.Close()

        Exit Sub

Write_Records:
        r = r + 1
        If r Mod 100 = 0 Then
            z = fldX(0).Value & ""
            Call Track("-", z)
        End If
        dynASWSRPT0.AddNew()
        dynASWSRPT0.Fields("RECORD_NO").Value = r
    GoSub Set_Key
        dynASWSRPT0.Fields("CUST_CODE").Value = CUST_CODE
        dynASWSRPT0.Fields("ITEM_CODE").Value = ITEM_CODE
        If typ = "A" Then
            dynASWSRPT0.Fields(2 + 10).Value = QTY
        ElseIf typ = "B" Then
            dynASWSRPT0.Fields(itemz).Value = cf(2, Val(Right$(itemz, 1))) & ":" & ITEM_CODE
        Else
            dynASWSRPT0.Fields("ORD_COUNT").Value = CNT_ORD
            dynASWSRPT0.Fields("ORD_UNITS").Value = QTY_ORD
            dynASWSRPT0.Fields("ORD_SALES").Value = AMT_ORD
            dynASWSRPT0.Fields("SHP_COUNT").Value = CNT_SHP
            dynASWSRPT0.Fields("SHP_UNITS").Value = QTY_SHP
            dynASWSRPT0.Fields("SHP_SALES").Value = AMT_SHP
            dynASWSRPT0.Fields("REL_COUNT").Value = CNT_REL
            dynASWSRPT0.Fields("REL_UNITS").Value = QTY_REL
            dynASWSRPT0.Fields("REL_SALES").Value = AMT_REL
            dynASWSRPT0.Fields("OPN_COUNT").Value = CNT_OPN
            dynASWSRPT0.Fields("OPN_UNITS").Value = QTY_OPN
            dynASWSRPT0.Fields("OPN_SALES").Value = AMT_OPN
        End If
        dynASWSRPT0.Update()

        If typ = "O" Then
            tblSOWWHOD2.Seek("=", CUST_CODE, CUST_STORE_NO, ITEM_CODE, "1")
            If tblSOWWHOD2.NoMatch Then
                tblSOWWHOD2.AddNew()
                tblSOWWHOD2.Fields("CUST_CODE").Value = CUST_CODE
                tblSOWWHOD2.Fields("CUST_STORE_NO").Value = CUST_STORE_NO
                tblSOWWHOD2.Fields("ITEM_CODE").Value = ITEM_CODE
                tblSOWWHOD2.Fields("ORDERED").Value = "1"
                tblSOWWHOD2.Update()
            End If
        ElseIf typ = "B" Then
            tblSOWWHOD4.Seek("=", CUST_CODE, ITEM_BRAND_CODE)
            If tblSOWWHOD4.NoMatch Then
                tblSOWWHOD4.AddNew()
                tblSOWWHOD4.Fields("CUST_CODE").Value = CUST_CODE
                tblSOWWHOD4.Fields("ITEM_BRAND_CODE").Value = ITEM_BRAND_CODE
                tblSOWWHOD4.Update()
            End If
        End If

        Return

Set_Key:
        For j = 1 To cfmax
            zz = fldX(j - 1).Value & ""
            z = Format$(j, "0")
            dynASWSRPT0.Fields("G" & z).Value = cf(2, j) & ":" & zz
        Next j
        Return

    End Sub

    Sub Print_Report()
        Dim z As String
        Dim i As Integer

        Call Std_Report_Parameters()

        z = RYP1Legend & " thru " & RYP2Legend
        CR_Rpt.ParameterFields(CR_Rpt_Names("SUBT")).SetCurrentValue(z)

        CR_Rpt.ParameterFields(CR_Rpt_Names("RECAP")).SetCurrentValue(chkRecap)
        CR_Rpt.ParameterFields(CR_Rpt_Names("NEWPAGE")).SetCurrentValue(chkNewPage)
        CR_Rpt.ParameterFields(CR_Rpt_Names("RC")).SetCurrentValue(aRC)
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG1")).SetCurrentValue(cf(2, 1))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG2")).SetCurrentValue(cf(2, 2))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG3")).SetCurrentValue(cf(2, 3))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG4")).SetCurrentValue(cf(2, 4))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG5")).SetCurrentValue(cf(2, 5))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG6")).SetCurrentValue(cf(2, 6))
        CR_Rpt.ParameterFields(CR_Rpt_Names("HG7")).SetCurrentValue(cf(2, 7))
        CR_Rpt.ParameterFields(CR_Rpt_Names("LVLS")).SetCurrentValue(CStr(cfmax))

        CR_Rpt.ParameterFields(CR_Rpt_Names("DIDNOT")).SetCurrentValue(chkNOT)

        Call Std_Report_Parameters(True)

        Call Prepare_SPRF()
    End Sub

End Class