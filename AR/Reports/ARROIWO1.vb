Public Class ARROIWO1
    Dim ARTDEDA1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        sqlw &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sqlw &= SQL_in("ITEM_CATGY_CODE", "ICTITEM1.ITEM_CATGY_CODE")
        sqlw &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")

        If Absx1.chkFor("CHK_MB_M").Checked And _
           Absx1.chkFor("CHK_MB_B").Checked Then
        Else
            sqlw &= " and ICTITEM1.ITEM_BASIC_PROMO in (" & Mid( _
                 IIf(Absx1.chkFor("CHK_MB_M").Checked, ",'M'", "") & _
                 IIf(Absx1.chkFor("CHK_MB_B").Checked, ",'B'", ""), 2) & ")"
        End If

        RWU = "R"

        Prepare_dst(True, sqlw, RYP)

        Check_if_Empty("ARTREAS1")
    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        If ARTDEDA1 = "" Then Create_Temp_Data(sqlw)

        With dst
            ASCMAIN1.sql = "Select ICTITEM1.* from " & ARTDEDA1 & " ICTITEM1"
            Create_TDA(dst.Tables.Add("ICTITEM1"), ARTDEDA1, "**", 0, True, , 1)
            With .Tables("ICTITEM1").Columns
                .Add("DEMAND_QTY", GetType(System.Int64), "ISNULL(FORECAST,0)+ISNULL(PROD_COM,0)+ISNULL(PLAN_COM,0)")
                .Add("DEMAND_AMT", GetType(System.Decimal), "DEMAND_QTY * ISNULL(ITEM_COST_STD,0)")
                .Add("DEMAND_PCT", GetType(System.Decimal))
                .Add("DEMAND_PCT_CUM", GetType(System.Decimal))
            End With

            Create_TDA(dst.Tables.Add, "ICTCOLL1", "*", 0)
            Create_TDA(dst.Tables.Add, "ICTBRAN1", "*", 0)
            Create_TDA(dst.Tables.Add, "ICTCATG1", "*", 0)
            Create_TDA(dst.Tables.Add, "DPTABCP1", "*", 0)
            .Tables("DPTABCP1").Columns.Add("ABC_INDEX", GetType(System.Int64))
            .Tables("DPTABCP1").Columns.Add("ABC_PCT_CUM", GetType(System.Decimal))

            With .Tables.Add("DPTABCPG")
                .Columns.Add("ABC_GROUP")
                .Columns.Add("ABC_GROUP_DESC")
                .PrimaryKey = New DataColumn() {.Columns("ABC_GROUP")}
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
        Fill_Records("ARTDEDA1")
        EnforceConstraints(True)

        '     Calculate_ABC()

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

        If ARTDEDA1 = "" Then
            ARTDEDA1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add FORECAST NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add PROD_COM NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add PLAN_COM NUMBER (8,0)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTDEDA1)

            Dim COLUMN_NAMEs As String = "" _
                & "ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, ITEM_COST_STD" _
                & ", ITEM_UOM, VEND_CODE, ITEM_PO_QTY_MIN, ITEM_MRP_PLANR_CODE, ABC_GROUP" _
                & ", ITEM_CATGY_CODE, COLLECTION_CODE, BRAND_CODE, ITEM_SNU_CODE, ITEM_BASIC_PROMO" _
                & ", ITEM_COST_MAKE_BUY, ITEM_ABC_CODE, ITEM_ABC_CODE_FUT"

            ASCDATA1.ExecuteSQL("Insert into " & ARTDEDA1 & " (" & COLUMN_NAMEs & ") " & ASCMAIN1.sql)


        End If

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        'CR_params.Add("MOS", Val(Absx1.numFor("FPDMOS").Value & ""))
        'CR_params.Add("CD", "Demand Calculated using " & Absx1.optFor("OPTDEMAND").Text)
        'CR_params.Add("EXTUSAGE", "Saleable Ranked by " & Absx1.optFor("OPTRANKBY").Text)
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If eItemKey = "Proceed" Then
                'If Absx1.cmbFor("RYP0").Value & "" = "" Then
                '    EMsg &= vbCr & "You must Specify a Starting Period"
                'End If
                'If Absx1.cmbFor("RYP1").Value & "" = "" Then
                '    EMsg &= vbCr & "You must Specify an Ending Period"
                'End If
                'If EMsg = "" Then
                '    If ASCMAIN1.Period_Diff(Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value) > 12 Then
                '        EMsg &= vbCr & "Maximum number of periods spanned by Period Range is 12"
                '    End If
                'End If
            End If
        End If

    End Sub

    Overrides Sub Update_Record()

    End Sub

#Region "VB6"
    'Option Explicit

    'Dim chkType As String
    'Dim chkFT As String
    'Dim ssdF As String
    'Dim ssdT As String

    'Dim PYMT_BATCH_NO As String
    'Dim REASON_CODE As String
    'Dim optRC As String

    'Sub Build_Workfile()

    '    Dim i As Integer

    '    Call Build_WorkFile_DB_Init()

    '    ' Get Run-Time options

    '    Call Track("Run-Time Options", "")

    '    chkFT = Format$(Val(SRead(opts, "CHDFINV_DATE", 2)), "0") & Format$(Val(SRead(opts, "CHDTINV_DATE", 2)), "0")
    '    ssdF = Format$(SRead(opts, "SSDFINV_DATE", 2), "dd-mmm-yyyy")
    '    ssdT = Format$(SRead(opts, "SSDTINV_DATE", 2), "dd-mmm-yyyy")

    '    chkType = Get_chk("TYPE", "IBDCOR", "Y")
    '    optRC = Get_opt("RC", "1I")
    '    REASON_CODE = SRead(opts, "CMBREAS:REASON_CODE", 2)

    '    ' Prepare Work File with Data from Server

    '    Call Track("Preparing Result Set ", "")

    '    sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
    '    Call Ora_to_Acc(Nothing, "ARWCUST1", 1, "", sql)

    '    sql = "SELECT ARTOPEN1.*" & vbCrLf
    '    sql = sql & " FROM ARTOPEN1" & vbCrLf
    '    sql = sql & "  WHERE ARTOPEN1.INV_BALANCE <> 0" & vbCrLf
    '    If chkType <> "" Then
    '        sql = sql & "    AND ARTOPEN1.INV_TYPE IN (" & chkType & ")" & vbCrLf
    '    End If

    '    ' CUST_CODE|REASON_CODE
    '    i = ColumnName_To_Index("CUST_CODE")
    '    If Sel(1, i) <> "" Then
    '        sql = sql & "    AND ARTOPEN1.CUST_CODE " & Sel(2, i) & " IN (" & Sel(1, i) & ")" & vbCrLf
    '    End If

    '    i = ColumnName_To_Index("REASON_CODE")
    '    If Sel(1, i) <> "" Then
    '        sql = sql & "    AND ARTOPEN1.REASON_CODE " & Sel(2, i) & " IN (" & Sel(1, i) & ")" & vbCrLf
    '    End If


    '    If Mid$(chkFT, 1, 1) <> "1" Then
    '        sql = sql & "    AND ARTOPEN1.INV_DATE >= '" & ssdF & "'"
    '    End If
    '    If Mid$(chkFT, 2, 1) <> "1" Then
    '        sql = sql & "    AND ARTOPEN1.INV_DATE <= '" & ssdT & "'"
    '    End If
    '    Call Ora_to_Acc(Nothing, "ARWOPEN1", 3, "", sql)

    '    sql = "SELECT * FROM ARTREAS1"
    '    Call Ora_to_Acc(Nothing, "ARWREAS1", 1, "", sql)

    '    PYMT_BATCH_NO = CTLNO("PYMT_BATCH_NO", 10)

    '    If SD_PARM_ACTIVE = "1" And (xUID = "CCU" Or xUID = "TST") Then
    '        xRWU = "N"
    '    End If

    '    ' Wrap up
    '    Call Track("", "")

    'End Sub

    'Sub Update()
    '    OraS.BeginTrans()

    '    Dim dynARTOPEN1 As OraDynaset
    '    sql = "SELECT * FROM ARTOPEN1 WHERE CUST_CODE = :CUST_CODE AND INV_TYPE = :INV_TYPE AND INV_NUM = :INV_NUM"
    '    dynARTOPEN1 = OraD.CreateDynaset(sql, 0&)

    '    Dim AR_PARM_DEF_SEG2 As String
    '    Dim AR_PARM_DEF_SEG3 As String
    '    Dim AR_PARM_DEF_SEG4 As String
    '    Dim AR_PARM_CURR_CODE As String
    '    Dim POST_CODE As String

    '    Dim INV_PMT As Double
    '    Dim INV_DISC_TAKEN As Double
    '    Dim INV_WRITE_OFF As Double
    '    Dim INV_BALANCE As Double
    '    Dim INV_PMT_CURR As Double
    '    Dim INV_DISC_TAKEN_CURR As Double
    '    Dim INV_WRITE_OFF_CURR As Double
    '    Dim INV_BALANCE_NEW_CURR As Double

    '    Dim INV_TOTALS As Double

    '    Dim CUST_CODE As String
    '    Dim CUST_CODE_SO As String

    '    Dim dynICTCURR1_CX As OraDynaset

    '    Dim PYMT_BATCH_LNO As Integer
    '    Dim PYMT_BATCH_XLNO As Integer

    '    Dim dynARTPARM1 As OraDynaset
    '    sql = "Select * from ARTPARM1 where AR_PARM_KEY = 'Z'"
    '    dynARTPARM1 = OraD.CreateDynaset(sql, 8&)
    '    AR_PARM_DEF_SEG2 = dynARTPARM1.Fields("AR_PARM_DEF_SEG2").Value & ""
    '    AR_PARM_DEF_SEG3 = dynARTPARM1.Fields("AR_PARM_DEF_SEG3").Value & ""
    '    AR_PARM_DEF_SEG4 = dynARTPARM1.Fields("AR_PARM_DEF_SEG4").Value & ""
    '    AR_PARM_CURR_CODE = dynARTPARM1.Fields("AR_PARM_CURR_CODE").Value & ""


    '    '    If CURR_CODE = dynARTPARM1.Fields("AR_PARM_CURR_CODE").Value & "" Then
    '    '        CURR_EXCH_RATE = 1
    '    '    End If

    '    Dim dynARWCASH1 As Recordset
    '    sql = "SELECT * FROM ARTCASH1 WHERE ROWNUM < 1"
    '    Call Ora_to_Acc(Nothing, "ARWCASH1", 1, "", sql)
    '    dynARWCASH1 = AccD.OpenRecordset("ARWCASH1", dbOpenDynaset)

    '    Dim dynARWCASH2 As Recordset
    '    sql = "SELECT * FROM ARTCASH2 WHERE ROWNUM < 1"
    '    Call Ora_to_Acc(Nothing, "ARWCASH2", 2, "", sql)
    '    dynARWCASH2 = AccD.OpenRecordset("ARWCASH2", dbOpenDynaset)

    '    Dim dynARWCASH3 As Recordset
    '    sql = "SELECT * FROM ARTCASH3 WHERE ROWNUM < 1"
    '    Call Ora_to_Acc(Nothing, "ARWCASH3", 3, "", sql)
    '    dynARWCASH3 = AccD.OpenRecordset("ARWCASH3", dbOpenDynaset)

    '    Dim dynARWCASH5 As Recordset
    '    sql = "SELECT * FROM ARTCASH5 WHERE ROWNUM < 1"
    '    Call Ora_to_Acc(Nothing, "ARWCASH5", 3, "", sql)
    '    dynARWCASH5 = AccD.OpenRecordset("ARWCASH5", dbOpenDynaset)

    '    Dim LAST_DATE As Date
    '    LAST_DATE = Now + NowTSD
    '    Dim PYMT_BATCH_DATE As Date
    '    PYMT_BATCH_DATE = CDate(Format$(LAST_DATE, "mm/dd/yy"))

    '    PYMT_BATCH_LNO = 1

    '    With dynARWCASH1
    '        .AddNew()

    '        .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '        .Fields("PYMT_BATCH_DATE").Value = PYMT_BATCH_DATE
    '        .Fields("STATUS").Value = "1"
    '        .Fields("INIT_OPER").Value = xUserID
    '        .Fields("LAST_OPER").Value = xUserID
    '        .Fields("INIT_DATE").Value = LAST_DATE
    '        .Fields("LAST_DATE").Value = LAST_DATE
    '        .Fields("ZERO_CHECK").Value = "1"
    '        .Fields("OPS_YYYYPP").Value = CYP
    '        .Fields("CURR_CODE").Value = AR_PARM_CURR_CODE
    '        .Fields("CURR_EXCH_RATE").Value = 1

    '        .Update()
    '    End With

    '    Dim tblARWCUST1 As Recordset
    '    tblARWCUST1 = AccD.OpenRecordset("ARWCUST1", dbOpenTable)
    '    tblARWCUST1.Index = "PrimaryKey"

    '    Dim dynARWOPEN1_CUST As Recordset
    '    sql = "SELECT DISTINCT CUST_CODE FROM ARWOPEN1"
    '    dynARWOPEN1_CUST = AccD.OpenRecordset(sql)
    '    Do While Not dynARWOPEN1_CUST.EOF
    '        CUST_CODE = dynARWOPEN1_CUST.Fields("CUST_CODE").Value
    '        tblARWCUST1.Seek("=", CUST_CODE)
    '        With dynARWCASH2
    '            .AddNew()

    '            .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '            .Fields("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO
    '            .Fields("CUST_CODE").Value = CUST_CODE
    '            .Fields("CUST_NAME").Value = tblARWCUST1.Fields("CUST_NAME").Value
    '            .Fields("CUST_CHECK_NO").Value = "Zero Check"
    '            .Fields("CUST_CHECK_DATE").Value = PYMT_BATCH_DATE
    '            .Fields("CUST_CHECK_AMT").Value = 0
    '            .Fields("CUST_CHECK_AMT_CURR").Value = 0
    '            .Fields("STATUS").Value = "2"

    '            .Update()
    '        End With

    '        PYMT_BATCH_XLNO = 1
    '        INV_TOTALS = 0

    '        Dim dynARWOPEN1 As Recordset
    '        sql = "SELECT * FROM ARWOPEN1 WHERE CUST_CODE = '" & CUST_CODE & "'"
    '        dynARWOPEN1 = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    '        Do While Not dynARWOPEN1.EOF
    '            If dynARWOPEN1.Fields("INV_TYPE").Value = "B" Then
    '                POST_CODE = dynARTPARM1.Fields("AR_PARM_POST_CODE_CB").Value
    '            Else
    '                POST_CODE = dynARTPARM1.Fields("AR_PARM_POST_CODE_OA").Value
    '            End If

    '            With dynARWCASH3
    '                .AddNew()

    '                .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '                .Fields("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO
    '                .Fields("PYMT_BATCH_ILNO").Value = PYMT_BATCH_XLNO
    '                .Fields("INV_TYPE").Value = dynARWOPEN1.Fields("INV_TYPE").Value
    '                .Fields("INV_NUM").Value = dynARWOPEN1.Fields("INV_NUM").Value
    '                .Fields("REASON_CODE").Value = dynARWOPEN1.Fields("REASON_CODE").Value
    '                .Fields("INV_DATE").Value = dynARWOPEN1.Fields("INV_DATE").Value
    '                .Fields("INV_DUE_DATE").Value = dynARWOPEN1.Fields("INV_DUE_DATE").Value
    '                .Fields("CUST_CODE_SO").Value = dynARWOPEN1.Fields("CUST_CODE_SO").Value
    '                .Fields("CUST_STORE_NO").Value = dynARWOPEN1.Fields("CUST_STORE_NO").Value
    '                .Fields("INV_CUST_PO").Value = dynARWOPEN1.Fields("INV_CUST_PO").Value
    '                .Fields("INV_BALANCE").Value = dynARWOPEN1.Fields("INV_BALANCE").Value
    '                INV_PMT = dynARWOPEN1.Fields("INV_BALANCE").Value
    '                .Fields("INV_PMT").Value = INV_PMT
    '                .Fields("INV_DISC_TAKEN").Value = 0
    '                .Fields("INV_WRITE_OFF").Value = 0
    '                .Fields("INV_BALANCE_NEW").Value = 0
    '                .Fields("POST_CODE").Value = POST_CODE
    '                .Fields("SEG2_CODE").Value = AR_PARM_DEF_SEG2
    '                .Fields("SEG3_CODE").Value = AR_PARM_DEF_SEG3
    '                .Fields("SEG4_CODE").Value = AR_PARM_DEF_SEG4
    '                .Fields("INV_BALANCE_CURR").Value = dynARWOPEN1.Fields("INV_BALANCE_CURR").Value
    '                INV_PMT_CURR = dynARWOPEN1.Fields("INV_BALANCE_CURR").Value
    '                INV_BALANCE_NEW_CURR = INV_PMT_CURR
    '                .Fields("INV_PMT_CURR").Value = INV_PMT_CURR
    '                INV_DISC_TAKEN = 0
    '                INV_DISC_TAKEN_CURR = 0
    '                .Fields("INV_DISC_TAKEN_CURR").Value = INV_DISC_TAKEN
    '                INV_WRITE_OFF = 0
    '                INV_WRITE_OFF_CURR = 0
    '                .Fields("INV_WRITE_OFF_CURR").Value = INV_WRITE_OFF
    '                INV_BALANCE = 0
    '                INV_BALANCE_NEW_CURR = 0
    '                .Fields("INV_BALANCE_NEW_CURR").Value = INV_BALANCE

    '                .Update()
    '            End With

    '            OraD.Parameters("CUST_CODE").Value = dynARWOPEN1.Fields("CUST_CODE").Value
    '            OraD.Parameters("INV_TYPE").Value = dynARWOPEN1.Fields("INV_TYPE").Value
    '            OraD.Parameters("INV_NUM").Value = dynARWOPEN1.Fields("INV_NUM").Value

    '            With dynARTOPEN1
    '                .Refresh()

    '                .Edit()

    '                .Fields("INV_LAST_PMT").Value = PYMT_BATCH_DATE
    '                .Fields("INV_PMT").Value = Val(.Fields("INV_PMT").Value & "") + INV_PMT
    '                .Fields("INV_DISC_TAKEN").Value = Val(.Fields("INV_DISC_TAKEN").Value & "") + INV_DISC_TAKEN
    '                .Fields("INV_WRITE_OFF").Value = Val(.Fields("INV_WRITE_OFF").Value & "") + INV_WRITE_OFF
    '                .Fields("INV_BALANCE").Value = INV_BALANCE

    '                .Fields("INV_PMT_CURR").Value = Val(.Fields("INV_PMT_CURR").Value & "") + INV_PMT_CURR
    '                .Fields("INV_DISC_TAKEN_CURR").Value = Val(.Fields("INV_DISC_TAKEN_CURR").Value & "") + INV_DISC_TAKEN_CURR
    '                .Fields("INV_WRITE_OFF_CURR").Value = Val(.Fields("INV_WRITE_OFF_CURR").Value & "") + INV_WRITE_OFF_CURR
    '                .Fields("INV_BALANCE_CURR").Value = INV_BALANCE_NEW_CURR

    '                .Fields("LAST_OPER").Value = xUserID
    '                .Fields("LAST_DATE").Value = PYMT_BATCH_DATE

    '                .Fields("INV_LAST_PMT_REF").Value = "Zero Check"
    '                .Fields("INV_LAST_PMT_REF_DT").Value = PYMT_BATCH_DATE

    '                .Update()
    '            End With

    '            INV_TOTALS = INV_TOTALS + INV_PMT

    '            PYMT_BATCH_XLNO = PYMT_BATCH_XLNO + 1
    '            dynARWOPEN1.MoveNext()
    '        Loop

    '        If optRC = "1" Then
    '            With dynARWCASH5
    '                .AddNew()

    '                .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '                .Fields("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO
    '                .Fields("PYMT_BATCH_DLNO").Value = 1
    '                .Fields("REASON_CODE").Value = REASON_CODE
    '                .Fields("ACCT_CODE").Value = Null
    '                .Fields("SEG2_CODE").Value = AR_PARM_DEF_SEG2
    '                .Fields("SEG3_CODE").Value = AR_PARM_DEF_SEG3
    '                .Fields("SEG4_CODE").Value = AR_PARM_DEF_SEG4
    '                .Fields("GL_DIST_AMT").Value = INV_TOTALS
    '                .Fields("CUST_CODE_SO").Value = CUST_CODE
    '                .Fields("GL_DIST_AMT_CURR").Value = INV_TOTALS
    '                .Fields("CHARGEBACK_IND").Value = 0

    '                .Update()
    '            End With
    '        Else
    '            Dim PYMT_BATCH_DLNO As Integer
    '            PYMT_BATCH_DLNO = 0
    '            sql = "Select ARWCASH3.REASON_CODE, Sum (ARWCASH3.INV_PMT) as INV_PMT "
    '            sql = sql & " from ARWCASH3 "
    '            sql = sql & " where PYMT_BATCH_NO = '" & PYMT_BATCH_NO & "'"
    '            sql = sql & "   and PYMT_BATCH_LNO = " & CStr(PYMT_BATCH_LNO)
    '            sql = sql & " group by ARWCASH3.REASON_CODE"
    '            Dim dynWK As Recordset
    '            dynWK = AccD.OpenRecordset(sql, dbOpenForwardOnly)
    '            Do While Not dynWK.EOF
    '                With dynARWCASH5
    '                    .AddNew()

    '                    .Fields("PYMT_BATCH_NO").Value = PYMT_BATCH_NO
    '                    .Fields("PYMT_BATCH_LNO").Value = PYMT_BATCH_LNO
    '                    PYMT_BATCH_DLNO = PYMT_BATCH_DLNO + 1
    '                    .Fields("PYMT_BATCH_DLNO").Value = PYMT_BATCH_DLNO
    '                    .Fields("REASON_CODE").Value = dynWK.Fields("REASON_CODE").Value
    '                    .Fields("ACCT_CODE").Value = Null
    '                    .Fields("SEG2_CODE").Value = AR_PARM_DEF_SEG2
    '                    .Fields("SEG3_CODE").Value = AR_PARM_DEF_SEG3
    '                    .Fields("SEG4_CODE").Value = AR_PARM_DEF_SEG4
    '                    .Fields("GL_DIST_AMT").Value = Val(dynWK.Fields("INV_PMT").Value & "")
    '                    .Fields("CUST_CODE_SO").Value = CUST_CODE
    '                    .Fields("GL_DIST_AMT_CURR").Value = Val(dynWK.Fields("INV_PMT").Value & "")
    '                    .Fields("CHARGEBACK_IND").Value = 0

    '                    .Update()
    '                End With
    '                dynWK.MoveNext()
    '            Loop
    '            dynWK.Close()
    '        End If

    '        PYMT_BATCH_LNO = PYMT_BATCH_LNO + 1

    '        dynARWOPEN1_CUST.MoveNext()
    '    Loop

    '    dynARWOPEN1.Close()
    '    dynARWOPEN1_CUST.Close()
    '    dynARWCASH5.Close()
    '    dynARWCASH3.Close()
    '    dynARWCASH2.Close()
    '    dynARWCASH1.Close()
    '    tblARWCUST1.Close()

    '    Call Acc_to_Ora("ARWCASH1", "")
    '    Call Acc_to_Ora("ARWCASH2", "")
    '    Call Acc_to_Ora("ARWCASH3", "")
    '    Call Acc_to_Ora("ARWCASH5", "")

    '    OraS.CommitTrans()

    '    Call Done()
    'End Sub

    'Sub Print_Report()
    '    Dim DATE_STR As String
    '    Dim TYPE_STR As String
    '    xRPT = "ARROIWO1"
    '    Call Std_Report_Parameters()
    '    CR_Rpt.ParameterFields("SUBT").SetCurrentValue("")
    '    DATE_STR = "Write Offs, as of " & Format$(Now, "mm/dd/yyyy")
    '    If Mid$(chkFT, 1, 1) <> "1" Then
    '        DATE_STR = DATE_STR & " from " & ssdF
    '    End If
    '    If Mid$(chkFT, 2, 1) <> "1" Then
    '        DATE_STR = DATE_STR & " thru " & ssdT
    '    End If
    '    If optRC = "1" Then
    '        DATE_STR = DATE_STR & ", using Reason Code " & REASON_CODE
    '    Else
    '        DATE_STR = DATE_STR & ", using Reason Code of Item"
    '    End If
    '    CR_Rpt.ParameterFields("DATE_STR").SetCurrentValue(DATE_STR)
    '    CR_Rpt.ParameterFields("PYMT_BATCH_NO").SetCurrentValue(PYMT_BATCH_NO)
    '    CR_Rpt.ParameterFields("REASON_CODE").SetCurrentValue(REASON_CODE)
    '    TYPE_STR = "Writing Off A/R Types: "
    '    If chkType <> "" Then
    '        TYPE_STR = TYPE_STR & chkType
    '    Else
    '        TYPE_STR = TYPE_STR & "All"
    '    End If
    '    CR_Rpt.ParameterFields("W_O_TYPES").SetCurrentValue(TYPE_STR)
    '    Call Prepare_SPRF()
    'End Sub
 

#End Region
End Class