Public Class SORINVP1

#Region "General Declarations"
    Private xDTE0 As Date
    Private xDTE1 As Date

    Dim SQLs As New Dictionary(Of String, String)

    Dim INV_TYPEs As String

    Dim SOTINVP1 As String
    Dim SOTINVH1 As String
    Dim SOTINVH2 As String

    Dim sqlSOTINVP1 As String
    Dim sqlSOTINVH1 As String
    Dim sqlSOTINVH2 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")

        Absx1.optFor("RANGE").CheckedIndex = 2
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        grpPERIOD_RANGE.Left = grpDATE_RANGE.Left
        grpPERIOD_RANGE.Top = grpDATE_RANGE.Top

        grpNotPrintedYet.Left = grpDATE_RANGE.Left
        grpNotPrintedYet.Top = grpDATE_RANGE.Top
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Invoices/Credits Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Invoices/Credits Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = " and SOTINVH1.INV_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'" & vbCrLf
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            If RYP0 = RYP1 Then
                SUBT = "Invoices/Credits Posted in " & RYPLEGEND0
            Else
                SUBT = "Invoices/Credits Posted between " & RYPLEGEND0 & " and " & RYPLEGEND1
            End If
            sqlw = " and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & RYP0 & "' and '" & RYP1 & "'" & vbCrLf
            RWU = "N"
        End If

        'INV_TYPEs _
        '= IIf(chkTypeS.Checked, ",'I'", "") _
        '& IIf(chkTypeR.Checked, ",'C'", "") 

        If optRANGE.Value = "U" Then
            sqlw &= "   and SOTINVH1.INV_PRINTED is Null" & vbCrLf
            If cmbDIV.Value & "" <> "" Then
                sqlw &= "   and SOTINVH1.SALES_DIVISION_CODE = '" & cmbDIV.Value & "'" & vbCrLf
            End If
            If chkMYINVOICESONLY.Visible And chkMYINVOICESONLY.Checked Then
                sqlw &= "   and SOTINVH1.INIT_OPER = '" & ASCMAIN1.USER_ID & "'" & vbCrLf
            End If
        End If

        sqlw &= SQL_in("SHIP_BOL_NO", "SOTINVH1.SHIP_BOL_NO") & vbCrLf
        sqlw &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE") & vbCrLf
        If Absx1.chkFor("CHKCONS_INV").Checked Then
            sqlw &= SQL_in("INV_NO", "SOTINVH1.INV_NO_CONS") & vbCrLf
        Else
            Dim SQLWX As String = SQL_in("INV_NO", "SOTINVH1.INV_NO")
            'If SQLWX <> "" Then sqlw &= " and (SOTINVH1.INV_TYPE = 'I'" & SQLWX & ")" & vbCrLf
            ' PRECEDING LINE DOES NOT SHOW CREDITS WHEN YOU ASK TO SEE CREDITS AND ENTER AN INV_NO
            If SQLWX <> "" Then sqlw &= SQLWX & vbCrLf
        End If

        If Absx1.chkFor("CHKEDI").Checked Then
            sqlw &= " and NVL(SOTORDR1.ORDR_SOURCE,'K') <> 'E'" & vbCrLf
        End If

        sqlw &= " and SOTINVH1.INV_NO_REV_BY is null" & vbCrLf

        If Absx1.optFor("RANGE").Value = "U" Then
            RWU = "R" ' "Y"
        Else
            RWU = "N"
        End If

        Prepare_dst(True, sqlw)

        Check_if_Empty("SOTINVH1")
    End Sub

    Public Overrides Sub Print_Report()
        Dim RPT As String = ROWs("ARTPARM1").Item("AR_PARM_INVOICE_RPT") & ""
        If RPT = "" Then RPT = "SORINVP1"

        CR_params.Add("SUBT", "")
        CR_params.Add("CONS_INV", IIf(Absx1.chkFor("CHKCONS_INV").Checked, "1", "0"))
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "P" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
            ElseIf Absx1.optFor("RANGE").Value = "D" Then
                If Absx1.dteFor("DTE0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Date"
                End If
                If Absx1.dteFor("DTE1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Date"
                End If
            Else

            End If
        End If

    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged

        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        grpNotPrintedYet.Visible = (optRANGE.Value = "U")
        grpDATE_RANGE.Enabled = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Enabled = (optRANGE.Value = "P")
        grpNotPrintedYet.Enabled = (optRANGE.Value = "U")
        grpOtherOptions.Visible = Not (optRANGE.Value = "U")

        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()
        Dim sql As String = "Update SOTINVH1 " _
        & " Set INV_PRINTED = '1'" _
        & " where (INV_TYPE, INV_NO) in (Select INV_TYPE, INV_NO from " & SOTINVP1 & " )"
        ASCDATA1.ExecuteSQL(sql)

        ' WHAT ABOUT DR & CRS

        sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
        & " Select 'SOTORDR1', ORDR_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'INVPRT','Invoice Printed', INV_NO" _
        & " from " & SOTINVP1
        ASCDATA1.ExecuteSQL(sql)
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("SOTPARM1")
        Get_PARM("ARTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        sqlSOTINVH1 = "Select SOTINVH1.* from SOTINVH1,SOTORDR1 where SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO"
        ASCMAIN1.sql = sqlSOTINVH1 & sqlw
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE, INV_NO)")

        sqlSOTINVH2 = "Select SOTINVH2.* from SOTINVH2, " & SOTINVH1 _
            & " SOTINVH1 where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO"
        ASCMAIN1.sql = sqlSOTINVH2
        SOTINVH2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " Add Primary Key (INV_TYPE, INV_NO, INV_LNO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " MODIFY CUST_CODE NULL")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " MODIFY CUST_STORE_NO NULL")
        'ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " MODIFY WHSE_CODE NULL", True)
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " MODIFY ORDR_YYYYPP_UPDATED NULL")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH2 & " MODIFY ORDR_UNIT_PRICE_CURR NULL")

        sqlSOTINVP1 = "Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.ORDR_NO, SOTINVH1.CUST_CODE" & vbCrLf _
            & ", SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.CUST_STORE_NO, SOTINVH1.SHIP_BOL_NO, SOTINVH1.PICK_NO" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1,SOTORDR1" & vbCrLf _
            & " where SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO"
        ASCMAIN1.sql = sqlSOTINVP1 & sqlw
        SOTINVP1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVP1 & " Add Primary Key (INV_TYPE, INV_NO)")

        SQLs.Clear()

        With dst

            ASCMAIN1.sql = "Select SOTINVH1.*" & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1"
            SQLs.Add("SOTINVH1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVH1", "**", 0, False, "", 2)
            With .Tables("SOTINVH1").Columns
                .Add("TOTAL_UNITS", GetType(System.Int64))
                .Add("AR_PARM_KEY")
                .Add("BT")
                .Add("ST")
                .Add("CART_TRACKING_NO")
                .Add("INV_SALES_CURR", GetType(System.Decimal), "IIF(CURR_CODE='USD',INV_SALES,INV_SALES / ISNULL(CURR_EXCH_RATE,0))")
                .Add("INV_FREIGHT_CURR", GetType(System.Decimal), "IIF(CURR_CODE='USD',INV_FREIGHT,INV_FREIGHT / ISNULL(CURR_EXCH_RATE,0))")
                .Add("INV_STAX_CURR", GetType(System.Decimal), "IIF(CURR_CODE='USD',INV_STAX,INV_STAX / ISNULL(CURR_EXCH_RATE,0))")
                .Add("INV_MISC_CHG_CURR", GetType(System.Decimal), "IIF(CURR_CODE='USD',INV_MISC_CHG,INV_MISC_CHG / ISNULL(CURR_EXCH_RATE,0))")
                .Add("INV_TOTAL_AMOUNT_CURR", GetType(System.Decimal), "IIF(CURR_CODE='USD',INV_TOTAL_AMOUNT,INV_TOTAL_AMOUNT / ISNULL(CURR_EXCH_RATE,0))")
            End With

            ASCMAIN1.sql = "Select SOTINVH2.*" & vbCrLf _
                & " from " & SOTINVH2 & " SOTINVH2"
            SQLs.Add("SOTINVH2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTINVH2", "**", 0, False, "", 3)
            .Tables("SOTINVH2").Columns.Add("DND", GetType(System.String))
            .Tables("SOTINVH2").Columns("DND").DefaultValue = "0"

            Create_Relation("SOTINVH1", "SOTINVH2", "INV_TYPE,INV_NO")
            .Tables("SOTINVH2").Columns.Add("ORDR_NO", GetType(System.String), "PARENT(SOTINVH1_SOTINVH2).ORDR_NO")
            .Tables("SOTINVH1").Columns("TOTAL_UNITS").Expression = "SUM(CHILD(SOTINVH1_SOTINVH2).ORDR_QTY_SHIP)"

            ASCMAIN1.sql = "Select SOTCITM1.CUST_CODE, SOTCITM1.ITEM_CODE, SOTCITM1.CUST_ITEM_CODE" & vbCrLf _
                & " from SOTCITM1 where (CUST_CODE,ITEM_CODE) in (Select Distinct CUST_CODE,ITEM_CODE from " & SOTINVH2 & ")"
            SQLs.Add("SOTCITM1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTCITM1", "**", 0, False, "", 2)

            'ASCMAIN1.sql = "Select SOTSREP1.SREP_CODE, SOTSREP1.SREP_NAME" & vbCrLf _
            '    & " from SOTSREP1 where SREP_CODE in (Select Distinct SREP_CODE from " & SOTINVH1 & ")"
            'SQLs.Add("SOTSREP1", ASCMAIN1.sql)
            'Create_TDA(.Tables.Add, "SOTSREP1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTSHIP1.*" & vbCrLf _
                & " from SOTSHIP1" & vbCrLf _
                & " where SHIP_BOL_NO in (Select Distinct SHIP_BOL_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTSHIP1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTSHIP1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTPICK1.*" & vbCrLf _
                & " from SOTPICK1" & vbCrLf _
                & " where PICK_NO in (Select Distinct PICK_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTPICK1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTPICK1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTPICK2.*" & vbCrLf _
                & " from SOTPICK2" & vbCrLf _
                & " where PICK_NO in (Select Distinct PICK_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTPICK2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTPICK2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR1.*" & vbCrLf _
                & " from SOTORDR1" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTORDR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select SOTORDR2.*" & vbCrLf _
                & " from SOTORDR2" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")"
            SQLs.Add("SOTORDR2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR2", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR5.*" & vbCrLf _
                & " from SOTORDR5" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")" & vbCrLf _
                & "   and CUST_ADDR_TYPE = 'BT'"
            SQLs.Add("SOTORDR5_BT", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR5_BT", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select SOTORDR5.*" & vbCrLf _
                & " from SOTORDR5" & vbCrLf _
                & " where ORDR_NO in (Select Distinct ORDR_NO from " & SOTINVP1 & ")" & vbCrLf _
                & "   and CUST_ADDR_TYPE = 'ST'"
            SQLs.Add("SOTORDR5_ST", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "SOTORDR5_ST", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_STD, ICTITEM1.ITEM_MATL_DESC, ICTITEM1.COUNTRY_CODE, ICTITEM1.ITEM_UPC_CODE, ICTITEM1.ITEM_EAN_CODE" & vbCrLf _
                & " from ICTITEM1" & vbCrLf _
                & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & SOTINVH2 & ")"
            SQLs.Add("ICTITEM1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)
             
            ASCMAIN1.sql = "Select Distinct ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" & vbCrLf _
                & ", ARTCUST1.CUST_ADDR1, ARTCUST1.CUST_ADDR2, ARTCUST1.CUST_ADDR3" & vbCrLf _
                & ", ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_COUNTRY" & vbCrLf _
                & ", ARTCUST1.CUST_ZIP_CODE, ARTCUST1.CUST_DBA_NAME " & vbCrLf _
                & ", ARTCUST1.CUST_XMIT_INV_VIA, ARTCUST1.CUST_INV_COMMENT " & vbCrLf _
                & ", ARTCUST1.CUST_INV_EMAIL, ARTCUST1.CUST_INV_CC" & vbCrLf _
                & ", ARTCUST1.CUST_CONTACT, ARTCUST1.CUST_BILL_SHIP_TO, ARTCUST1.CUST_INCL_INV_SHIP" & vbCrLf _
                & ", ARTCUST1.CUST_VEND_REF" & vbCrLf _
                & " from ARTCUST1" & vbCrLf _
                & " where CUST_CODE in" & vbCrLf _
                & " (Select Distinct CUST_CODE from " & SOTINVP1 & " union Select Distinct CUST_BILL_TO_CUST from " & SOTINVP1 & ")"
            SQLs.Add("ARTCUST1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ARTCUST1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ARTCUST2.* from ARTCUST2 where (CUST_CODE,CUST_STORE_NO) in " & vbCrLf _
                & " (Select Distinct CUST_CODE,CUST_STORE_NO from " & SOTINVP1 & ")"
            SQLs.Add("ARTCUST2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ARTCUST2", "**", 0, False, "", 3)

            For Each TABLE_NAME As String In New String() _
            {"TATTERM1", "ICTWHSE1", "SOTSVIA1", "SOTSREP1", "SOTREAS1", "ARTREAS1", "SOTSDIV1", "TATCNTRY"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            With .Tables.Add("SOTINVP0")
                .Columns.Add("AR_PARM_KEY")
                .Columns.Add("REMIT0")
                .Columns.Add("REMIT1")
                .Columns.Add("REMIT2")
                .Columns.Add("REMIT3")
                .Columns.Add("AR_PARM_REMIT_MESSAGE")
                .Columns.Add("ADDRESS0")
                .Columns.Add("ADDRESS1")
                .Columns.Add("ADDRESS2")
                .Columns.Add("ADDRESS3")
                .Columns.Add("AR_PARM_DUNS_NO")
                .Columns.Add("LOGO", GetType(System.Byte()))
                .PrimaryKey = New DataColumn() {.Columns("AR_PARM_KEY")}
            End With

            'SOTMISC1
            Create_TDA(.Tables.Add, "SOTMISC1", "*")
            Fill_Records("SOTMISC1", String.Empty, True, "Select * from SOTMISC1")

            'With .Tables("SOTSDIV1")
            '    .Columns.Add("DIVISION_LOGO", GetType(System.Byte()))
            'End With

            'ASCMAIN1.sql = "Select Distinct SOTSDIV1.DIVISION_CODE, SOTINVP1.*" _
            '& " from SOTINVP1,SOTSDIV1 " _
            '& " where SOTINVP1.CUST_CODE = SOTSDIV1.CUST_CODE"
            'Create_TDA(.Tables.Add, "SOTINVP1_DIV", "**", 0, False, "", 1)

        End With

        Dim rowSOTINVP0 As DataRow = dst.Tables("SOTINVP0").NewRow
        With ROWs("ARTPARM1")
            rowSOTINVP0.Item("AR_PARM_KEY") = "Z"
            rowSOTINVP0.Item("REMIT0") = .Item("AR_PARM_REMIT_NAME") & ""
            rowSOTINVP0.Item("REMIT1") = .Item("AR_PARM_REMIT_ADDR1") & ""
            rowSOTINVP0.Item("REMIT2") = .Item("AR_PARM_REMIT_ADDR2") & ""
            rowSOTINVP0.Item("REMIT3") = .Item("AR_PARM_REMIT_CITY") & ", " _
                    & .Item("AR_PARM_REMIT_STATE") & " " _
                    & .Item("AR_PARM_REMIT_ZIP_CODE") & " " _
                    & .Item("AR_PARM_REMIT_COUNTRY")
            ' rowSOTINVP0.Item("REMIT3") = "Tel " & .Item("AR_PARM_REMIT_PHONE") & " Fax " & .Item("AR_PARM_REMIT_FAX")
            ' rowSOTINVP0.Item("REMIT3") = ""
            rowSOTINVP0.Item("AR_PARM_REMIT_MESSAGE") = .Item("AR_PARM_REMIT_MESSAGE") & ""
            rowSOTINVP0.Item("AR_PARM_DUNS_NO") = .Item("AR_PARM_DUNS_NO") & ""
        End With

        With ASCMAIN1.rowASTPARM1
            rowSOTINVP0.Item("ADDRESS0") = .Item("AS_PARM_INST_NAME") & ""
            rowSOTINVP0.Item("ADDRESS1") = .Item("AS_PARM_INST_ADDR1") & ""
            rowSOTINVP0.Item("ADDRESS2") = .Item("AS_PARM_INST_ADDR2") & ""
            rowSOTINVP0.Item("ADDRESS3") = .Item("AS_PARM_INST_CITY") & ", " _
                    & .Item("AS_PARM_INST_STATE") & " " _
                    & .Item("AS_PARM_INST_ZIP_CODE") & " " _
                    & .Item("AS_PARM_INST_COUNTRY")

            Dim TEL As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_PHONE")
            If TEL.Length = 10 Then
                TEL = "(" & Mid(TEL, 1, 3) & ")" & Mid(TEL, 4, 3) & "-" & Mid(TEL, 7, 4)
            End If
            Dim FAX As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_FAX")
            If FAX.Length = 10 Then
                FAX = "(" & Mid(FAX, 1, 3) & ")" & Mid(FAX, 4, 3) & "-" & Mid(FAX, 7, 4)
            End If
            rowSOTINVP0.Item("ADDRESS3") = "P " & TEL & " F " & FAX
            ' rowSOTINVP0.Item("ADDRESS3") = ""
        End With

        rowSOTINVP0.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")
        dst.Tables("SOTINVP0").Rows.Add(rowSOTINVP0)

        '  Fill_Records("SOTINVP1_DIV")

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""

        Dim pro_forma As Boolean = False
        Dim INV_TYPE_requested As String = ""

        If parms.Length > 0 Then
            'Dim INV_NOs As String = parms(0)
            sqlw = parms(0)
            'not nec no more
            'If Not Trim(sqlw).ToUpper.StartsWith("AND") Then
            '    sqlw = " and " & sqlw
            'End If

            If parms.Length >= 2 Then
                pro_forma = (parms(1) = "1")
            End If
            If parms.Length >= 3 Then
                INV_TYPE_requested = parms(2) & ""
            End If

            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVP1)
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVH1)
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTINVH2)

            If pro_forma And INV_TYPE_requested <> "C" Then
                If INV_TYPE_requested = "O" Then
                    ASCMAIN1.sql = "Insert into " & SOTINVH2 & vbCrLf _
                        & " (INV_TYPE,INV_NO,INV_LNO,ITEM_CODE,ORDR_UNIT_PRICE,ORDR_QTY_SHIP,ORDR_UNIT_PRICE_CURR, CUST_CODE, ITEM_RETAIL_PRICE, ITEM_RETAIL_PRICE_CURR)" & vbCrLf _
                        & "Select 'P' INV_TYPE, SOTORDR2.ORDR_NO INV_NO, SOTORDR2.ORDR_LNO INV_LNO" & vbCrLf _
                        & ", SOTORDR2.ITEM_CODE, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                        & ", SOTORDR2.ORDR_QTY ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE ORDR_UNIT_PRICE_CURR, SOTORDR2.CUST_CODE, SOTORDR2.ITEM_RETAIL_PRICE, SOTORDR2.ITEM_RETAIL_PRICE_CURR" & vbCrLf _
                        & " from SOTORDR2,ICTITEM1" & vbCrLf _
                        & " where ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                        & Replace(sqlw, "SOTINVH1.", "SOTORDR1.")
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "Insert into " & SOTINVH1 & vbCrLf _
                        & "(INV_TYPE,INV_NO,CUST_CODE,CUST_STORE_NO,ORDR_CUST_PO,ORDR_NO,WHSE_CODE," & vbCrLf _
                        & "REASON_CODE,INV_DATE,CUST_BILL_TO_CUST,POST_CODE,SHIP_BOL_NO," & vbCrLf _
                        & "SALES_DIVISION_CODE,TERM_CODE,PICK_NO," & vbCrLf _
                        & "CUST_FACTOR_IND,SREP_CODE,INV_COMMENT," & vbCrLf _
                        & "SREP2_CODE,ORDR_DEPT,CURR_CODE,CURR_EXCH_RATE,ORDR_YYYYPP_UPDATED)" & vbCrLf _
                        & "Select 'P' INV_TYPE, SOTORDR1.ORDR_NO INV_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                        & ", SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE" & vbCrLf _
                        & ", SOTORDR1.REASON_CODE, SOTORDR1.ORDR_SHIP_DATE INV_DATE" & vbCrLf _
                        & ", SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.POST_CODE, NULL SHIP_BOL_NO" & vbCrLf _
                        & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.TERM_CODE, NULL PICK_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_FACTOR_IND, SOTORDR1.SREP_CODE, SOTORDR1.ORDR_INV_COMMENT" & vbCrLf _
                        & ", SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                        & ", SOTORDR1.CURR_CODE, SOTORDR1.CURR_EXCH_RATE, '000000' ORDR_YYYYPP_UPDATED" & vbCrLf _
                        & " from SOTORDR1,ARTCUST1 " & vbCrLf _
                        & " where ARTCUST1.CUST_CODE = SOTORDR1.CUST_CODE" & vbCrLf _
                        & Replace(sqlw, "SOTINVH1.", "SOTORDR1.")
                    ASCDATA1.ExecuteSQL()
                Else
                    ASCMAIN1.sql = "Insert into " & SOTINVH2 & vbCrLf _
                        & " (INV_TYPE,INV_NO,INV_LNO,ITEM_CODE,ORDR_UNIT_PRICE,ORDR_QTY_SHIP,ORDR_UNIT_PRICE_CURR)" & vbCrLf _
                        & "Select 'P' INV_TYPE, SOTPICK2.PICK_NO INV_NO, SOTPICK2.PICK_LNO INV_LNO" & vbCrLf _
                        & ", SOTORDR2.ITEM_CODE, SOTORDR2.ORDR_UNIT_PRICE" & vbCrLf _
                        & ", SOTPICK2.PICK_QTY ORDR_QTY_SHIP, SOTORDR2.ORDR_UNIT_PRICE ORDR_UNIT_PRICE_CURR" & vbCrLf _
                        & " from SOTPICK2,SOTPICK1,SOTORDR2,SOTORDR1,ICTITEM1" & vbCrLf _
                        & " where SOTORDR2.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                        & "   and SOTORDR2.ORDR_LNO = SOTPICK2.ORDR_LNO" & vbCrLf _
                        & "   and SOTPICK2.PICK_NO = SOTPICK1.PICK_NO " & vbCrLf _
                        & "   and SOTORDR1.ORDR_NO = SOTPICK2.ORDR_NO" & vbCrLf _
                        & "   and ICTITEM1.ITEM_CODE = SOTORDR2.ITEM_CODE" & vbCrLf _
                        & Replace(sqlw, "SOTINVH1.", "SOTPICK1.")
                    ASCDATA1.ExecuteSQL()

                    ASCMAIN1.sql = "Insert into " & SOTINVH1 & vbCrLf _
                        & "(INV_TYPE,INV_NO,CUST_CODE,CUST_STORE_NO,ORDR_CUST_PO,ORDR_NO,WHSE_CODE," & vbCrLf _
                        & "REASON_CODE,INV_DATE,CUST_BILL_TO_CUST,POST_CODE,SHIP_BOL_NO," & vbCrLf _
                        & "SALES_DIVISION_CODE,TERM_CODE,PICK_NO," & vbCrLf _
                        & "CUST_FACTOR_IND,SREP_CODE,INV_COMMENT," & vbCrLf _
                        & "SREP2_CODE,ORDR_DEPT,CURR_CODE,CURR_EXCH_RATE,ORDR_YYYYPP_UPDATED)" & vbCrLf _
                        & "Select 'P' INV_TYPE, SOTPICK1.PICK_NO INV_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_CODE, SOTORDR1.CUST_STORE_NO, SOTORDR1.ORDR_CUST_PO" & vbCrLf _
                        & ", SOTORDR1.ORDR_NO, SOTORDR1.WHSE_CODE" & vbCrLf _
                        & ", SOTORDR1.REASON_CODE, SOTORDR1.ORDR_SHIP_DATE INV_DATE" & vbCrLf _
                        & ", SOTORDR1.CUST_BILL_TO_CUST, SOTORDR1.POST_CODE, SOTPICK1.SHIP_BOL_NO" & vbCrLf _
                        & ", SOTORDR1.SALES_DIVISION_CODE, SOTORDR1.TERM_CODE, SOTPICK1.PICK_NO" & vbCrLf _
                        & ", SOTORDR1.CUST_FACTOR_IND, SOTORDR1.SREP_CODE, SOTORDR1.ORDR_INV_COMMENT" & vbCrLf _
                        & ", SOTORDR1.SREP2_CODE, SOTORDR1.ORDR_DEPT" & vbCrLf _
                        & ", SOTORDR1.CURR_CODE, SOTORDR1.CURR_EXCH_RATE" & vbCrLf _
                        & ", '000000' ORDR_YYYYPP_UPDATED" & vbCrLf _
                        & " from SOTPICK1,SOTORDR1 " & vbCrLf _
                        & " where SOTORDR1.ORDR_NO = SOTPICK1.ORDR_NO" & vbCrLf _
                        & Replace(sqlw, "SOTINVH1.", "SOTPICK1.")
                    ASCDATA1.ExecuteSQL()
                End If

                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INIT_DATE = SYSDATE, INIT_OPER = '" & ASCMAIN1.USER_ID & "'")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " SOTINVH1 Set INV_SALES = (Select Sum (ORDR_QTY_SHIP * ORDR_UNIT_PRICE) from " & SOTINVH2 & " where INV_TYPE = SOTINVH1.INV_TYPE and INV_NO = SOTINVH1.INV_NO)")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " SOTINVH1 Set INV_COGS = (Select Sum (ORDR_QTY_SHIP * ITEM_UNIT_COST) from " & SOTINVH2 & " where INV_TYPE = SOTINVH1.INV_TYPE and INV_NO = SOTINVH1.INV_NO)")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_MISC_CHG = 0")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_FREIGHT = 0")
                ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_TOTAL_AMOUNT = NVL(INV_SALES,0) + NVL(INV_FREIGHT,0) + NVL(INV_STAX,0)")
                'ASCDATA1.ExecuteSQL("Update " & SOTINVH1 & " Set INV_SALES_CURR = INV_SALES, INV_FREIGHT_CURR = INV_FREIGHT, INV_MISC_CHG_CURR = INV_MISC_CHG, INV_TOTAL_AMOUNT_CURR = INV_TOTAL_AMOUNT, INV_STAX_CURR = INV_STAX")
            Else
                ASCDATA1.ExecuteSQL("Insert into " & SOTINVH1 & " " & sqlSOTINVH1 & Replace(sqlw, "ORDR_NO", "SOTORDR1.ORDR_NO"))
                ASCDATA1.ExecuteSQL("Insert into " & SOTINVH2 & " " & sqlSOTINVH2 & sqlw)
            End If


            ASCDATA1.ExecuteSQL("Insert into " & SOTINVP1 & " " & sqlSOTINVP1 & Replace(sqlw, "ORDR_NO", "SOTORDR1.ORDR_NO"))

        End If

        ASCDATA1.ExecuteSQL("Insert into " & SOTINVH2 & " (INV_TYPE, INV_NO, INV_LNO, ITEM_CODE) Select INV_TYPE, INV_NO, 0 INV_LNO, NVL(MISC_CHG_CODE, 'Misc') ITEM_CODE from " & SOTINVH1 & " where (INV_TYPE,INV_NO) in (Select Distinct INV_TYPE, INV_NO from " & SOTINVH1 & " minus Select Distinct INV_TYPE, INV_NO from " & SOTINVH2 & ")")


        EnforceConstraints(False)
        Fill_Records("SOTORDR1")
        Fill_Records("SOTORDR2")
        Fill_Records("SOTORDR5_BT")
        Fill_Records("SOTORDR5_ST")
        Fill_Records("ICTITEM1")
        Fill_Records("SOTINVH1")
        Fill_Records("SOTINVH2")
        Fill_Records("ARTCUST1")
        Fill_Records("SOTCITM1")
        Fill_Records("SOTSREP1")

        ' Upper case the Country Code
        For Each table As String In New String() {"SOTORDR5_BT", "SOTORDR5_ST"}
            For Each row As DataRow In dst.Tables(table).Select("")
                row.Item("CUST_COUNTRY") = (row.Item("CUST_COUNTRY") & String.Empty).ToString.Trim.ToUpper.Replace(".", "")
            Next
        Next
 
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            rowSOTINVH1.Item("BT") = "BT"
            rowSOTINVH1.Item("ST") = "ST"
            rowSOTINVH1.Item("AR_PARM_KEY") = "Z"
            ' rowSOTINVH1.Item("TOTAL_UNITS") = Val(dst.Tables("SOTINVH2").Compute("SUM(ORDR_QTY_SHIP)", "") & "")
        Next

        Dim rowARTCUST1 As DataRow = Nothing
        Dim rowARTCUST2 As DataRow = Nothing
        Dim rowSOTORDR5 As DataRow = Nothing
        Dim rowICTITEM1 As DataRow = Nothing
        Dim rowSOTORDR1 As DataRow = Nothing

         ' Need to allow from credits/invoices that do not have sotordr* records
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("", "CUST_CODE")
            Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO") & String.Empty
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO") & String.Empty
            Dim CUST_CODE As String = rowSOTINVH1.Item("CUST_CODE") & String.Empty
            Dim CUST_STORE_NO As String = rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty
            rowSOTORDR1 = Nothing
            rowSOTORDR5 = Nothing

            If rowARTCUST1 Is Nothing OrElse rowARTCUST1.Item("CUST_CODE") <> CUST_CODE Then
                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1.Item("CUST_BILL_TO_CUST") & String.Empty <> String.Empty Then
                    rowARTCUST1 = LookUp("ARTCUST1", rowARTCUST1.Item("CUST_BILL_TO_CUST") & String.Empty)
                End If
            End If

            If ORDR_NO = "0000000000" OrElse ORDR_NO = String.Empty Then
                ORDR_NO = "O" & INV_NO.Substring(1)
                rowSOTINVH1.Item("ORDR_NO") = ORDR_NO
            End If

            If dst.Tables("SOTORDR1").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
                rowSOTORDR1.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR1.Item("CUST_CODE") = rowSOTINVH1.Item("CUST_CODE")
                rowSOTORDR1.Item("CUST_STORE_NO") = rowSOTINVH1.Item("CUST_STORE_NO")
                rowSOTORDR1.Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS")
                If rowSOTORDR1.Item("FRT_TERMS") & String.Empty = String.Empty Then
                    rowSOTORDR1.Item("FRT_TERMS") = "."
                End If
                rowSOTORDR1.Item("WHSE_CODE") = rowSOTINVH1.Item("WHSE_CODE")
                If rowSOTORDR1.Item("WHSE_CODE") & String.Empty = String.Empty Then
                    rowSOTORDR1.Item("WHSE_CODE") = "."
                End If
                rowSOTORDR1.Item("ORDR_STATUS") = "F"
                rowSOTORDR1.Item("CURR_CODE") = rowSOTINVH1.Item("CURR_CODE")
                rowSOTORDR1.Item("CURR_EXCH_RATE") = rowSOTINVH1.Item("CURR_EXCH_RATE")
                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)
            End If

            If dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "'").Length = 0 Then
                Dim rowSOTORDR2 As DataRow = Nothing
                If dst.Tables("SOTINVH2").Select("INV_NO = '" & INV_NO & "'").Length > 0 Then
                    For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select("INV_NO = '" & INV_NO & "'")
                        rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                        rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDR2.Item("ORDR_LNO") = rowSOTINVH2.Item("INV_LNO")

                        rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(rowSOTINVH2.Item("ITEM_CODE"))
                        If rowICTITEM1 Is Nothing Then
                            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
                                 & ", ICTITEM1.ITEM_COST_STD, ICTITEM1.ITEM_MATL_DESC, ICTITEM1.COUNTRY_CODE, ICTITEM1.ITEM_UPC_CODE" & vbCrLf _
                                 & " from ICTITEM1" & vbCrLf _
                                 & " where ITEM_CODE = '" & rowSOTINVH2.Item("ITEM_CODE") & "'"
                            Fill_Records("ICTITEM1", String.Empty, False, ASCMAIN1.sql)

                            rowICTITEM1 = dst.Tables("ICTITEM1").Rows.Find(rowSOTINVH2.Item("ITEM_CODE"))
                            If rowICTITEM1 Is Nothing Then
                                rowICTITEM1 = dst.Tables("ICTITEM1").NewRow
                                rowICTITEM1.Item("ITEM_CODE") = rowSOTINVH2.Item("ITEM_CODE")
                                rowICTITEM1.Item("ITEM_DESC") = rowSOTINVH2.Item("ITEM_CODE")
                                rowICTITEM1.Item("ITEM_COST_STD") = 0
                                rowICTITEM1.Item("ITEM_MATL_DESC") = String.Empty
                                rowICTITEM1.Item("COUNTRY_CODE") = "US"
                                rowICTITEM1.Item("ITEM_UPC_CODE") = "000000000000"
                                dst.Tables("ICTITEM1").Rows.Add(rowICTITEM1)
                            End If
                        End If

                        rowSOTORDR2.Item("ITEM_CODE") = rowICTITEM1.Item("ITEM_CODE")
                        rowSOTORDR2.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                        rowSOTORDR2.Item("ORDR_UNIT_PRICE") = rowSOTINVH2.Item("ORDR_UNIT_PRICE")
                        rowSOTORDR2.Item("ORDR_QTY_SHIP") = rowSOTINVH2.Item("ORDR_QTY_SHIP")
                        rowSOTORDR2.Item("ORDR_STATUS") = "F"
                        rowSOTORDR2.Item("CUST_CODE") = rowSOTINVH1.Item("CUST_CODE")
                        rowSOTORDR2.Item("CUST_STORE_NO") = rowSOTINVH1.Item("CUST_STORE_NO")
                        'rowSOTORDR1.Item("WHSE_CODE")
                        If rowSOTINVH1.Item("WHSE_CODE") & String.Empty <> String.Empty Then
                            rowSOTORDR2.Item("WHSE_CODE") = rowSOTINVH1.Item("WHSE_CODE")
                        ElseIf rowSOTORDR1 IsNot Nothing AndAlso rowSOTORDR1.Item("WHSE_CODE") & String.Empty <> String.Empty Then
                            rowSOTORDR2.Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
                        Else
                            rowSOTORDR2.Item("WHSE_CODE") = "16"
                        End If

                        dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)

                    Next
                Else
                    rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
                    rowSOTORDR2.Item("ORDR_NO") = ORDR_NO
                    rowSOTORDR2.Item("ORDR_LNO") = 1
                    rowSOTORDR2.Item("ITEM_CODE") = "Dummy"
                    rowSOTORDR2.Item("ORDR_UNIT_PRICE") = 0
                    rowSOTORDR2.Item("ORDR_QTY_SHIP") = 0
                    rowSOTORDR2.Item("ORDR_STATUS") = "F"
                    rowSOTORDR2.Item("CUST_CODE") = rowSOTINVH1.Item("CUST_CODE")
                    rowSOTORDR2.Item("CUST_STORE_NO") = rowSOTINVH1.Item("CUST_STORE_NO")
                    rowSOTORDR2.Item("WHSE_CODE") = rowSOTINVH1.Item("WHSE_CODE")
                    dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
                End If
            Else
                If rowARTCUST1.Item("TRADE_CLASS_CODE") & "" = "IND" Then
                    Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE") & String.Empty
                    For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' and ORDR_RELEASE = 'R'")
                        Dim ORDR_LNO As Integer = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").Rows.Find(New Object() {INV_TYPE, INV_NO, ORDR_LNO})
                        If rowSOTINVH2 Is Nothing Then
                            rowSOTINVH2 = dst.Tables("SOTINVH2").NewRow
                            rowSOTINVH2.Item("INV_TYPE") = INV_TYPE
                            rowSOTINVH2.Item("INV_NO") = INV_NO
                            rowSOTINVH2.Item("INV_LNO") = ORDR_LNO
                            rowSOTINVH2.Item("ITEM_CODE") = rowSOTORDR2.Item("CUST_CODE")
                            rowSOTINVH2.Item("ORDR_UNIT_PRICE") = rowSOTORDR2.Item("ORDR_UNIT_PRICE")
                            rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") = rowSOTORDR2.Item("ORDR_UNIT_PRICE_CURR")
                            rowSOTINVH2.Item("ORDR_QTY_SHIP") = 0
                            rowSOTINVH2.Item("ORDR_NO") = rowSOTORDR2.Item("ORDR_NO")
                            rowSOTINVH2.Item("CUST_CODE") = rowSOTORDR2.Item("CUST_CODE")
                            rowSOTINVH2.Item("CUST_STORE_NO") = rowSOTORDR2.Item("CUST_STORE_NO")
                            rowSOTINVH2.Item("WHSE_CODE") = rowSOTORDR2.Item("WHSE_CODE")
                            rowSOTINVH2.Item("ITEM_RETAIL_PRICE") = rowSOTORDR2.Item("ITEM_RETAIL_PRICE")
                            dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)
                        End If
                    Next
                End If
            End If

            If dst.Tables("SOTORDR5_BT").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'BT'").Length = 0 Then
                rowSOTORDR5 = dst.Tables("SOTORDR5_BT").NewRow
                rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                rowSOTORDR5.Item("CUST_ADDR_TYPE") = "BT"
                rowSOTORDR5.Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME")
                rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST1.Item("CUST_ADDR1")
                rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST1.Item("CUST_ADDR2")
                rowSOTORDR5.Item("CUST_ADDR3") = rowARTCUST1.Item("CUST_ADDR3")
                rowSOTORDR5.Item("CUST_CITY") = rowARTCUST1.Item("CUST_CITY")
                rowSOTORDR5.Item("CUST_STATE") = rowARTCUST1.Item("CUST_STATE")
                rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST1.Item("CUST_ZIP_CODE")
                rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST1.Item("CUST_COUNTRY")
                rowSOTORDR5.Item("CUST_CONTACT") = rowARTCUST1.Item("CUST_CONTACT")
                rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST1.Item("CUST_PHONE")
                rowSOTORDR5.Item("CUST_EXT") = rowARTCUST1.Item("CUST_EXT")
                rowSOTORDR5.Item("CUST_FAX") = rowARTCUST1.Item("CUST_FAX")
                rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST1.Item("CUST_EMAIL")
                rowSOTORDR5.Item("CUST_ADDR_CODE") = rowARTCUST1.Item("CUST_CODE")
                dst.Tables("SOTORDR5_BT").Rows.Add(rowSOTORDR5)
            End If

            If dst.Tables("SOTORDR5_ST").Select("ORDR_NO = '" & ORDR_NO & "' and CUST_ADDR_TYPE = 'ST'").Length = 0 Then
                If rowSOTINVH1.Item("CUST_STORE_NO") & String.Empty <> String.Empty Then

                    If rowARTCUST2 Is Nothing OrElse CUST_CODE <> rowARTCUST2.Item("CUST_CODE") OrElse CUST_STORE_NO <> rowARTCUST2.Item("CUST_STORE_NO") Then
                        rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})
                    End If

                    If rowARTCUST2 IsNot Nothing Then
                        rowSOTORDR5 = dst.Tables("SOTORDR5_ST").NewRow
                        rowSOTORDR5.Item("ORDR_NO") = ORDR_NO
                        rowSOTORDR5.Item("CUST_ADDR_TYPE") = "ST"
                        rowSOTORDR5.Item("CUST_NAME") = rowARTCUST2.Item("CUST_STORE_NAME")
                        rowSOTORDR5.Item("CUST_ADDR1") = rowARTCUST2.Item("CUST_STORE_ADDR1")
                        rowSOTORDR5.Item("CUST_ADDR2") = rowARTCUST2.Item("CUST_STORE_ADDR2")
                        rowSOTORDR5.Item("CUST_ADDR3") = rowARTCUST2.Item("CUST_STORE_ADDR3")
                        rowSOTORDR5.Item("CUST_CITY") = rowARTCUST2.Item("CUST_STORE_CITY")
                        rowSOTORDR5.Item("CUST_STATE") = rowARTCUST2.Item("CUST_STORE_STATE")
                        rowSOTORDR5.Item("CUST_ZIP_CODE") = rowARTCUST2.Item("CUST_STORE_ZIP_CODE")
                        rowSOTORDR5.Item("CUST_COUNTRY") = rowARTCUST2.Item("CUST_STORE_COUNTRY")
                        rowSOTORDR5.Item("CUST_CONTACT") = rowARTCUST2.Item("CUST_STORE_CONTACT")
                        rowSOTORDR5.Item("CUST_PHONE") = rowARTCUST2.Item("CUST_STORE_PHONE")
                        rowSOTORDR5.Item("CUST_EXT") = rowARTCUST2.Item("CUST_STORE_EXT")
                        rowSOTORDR5.Item("CUST_FAX") = rowARTCUST2.Item("CUST_STORE_FAX")
                        rowSOTORDR5.Item("CUST_EMAIL") = rowARTCUST2.Item("CUST_STORE_EMAIL")
                        rowSOTORDR5.Item("CUST_ADDR_CODE") = rowSOTINVH1.Item("CUST_STORE_NO")
                        dst.Tables("SOTORDR5_ST").Rows.Add(rowSOTORDR5)
                    End If
                End If
            End If
        Next

        ASCMAIN1.sql = "Select INV_NO, MIN (CART_TRACKING_NO) CART_TRACKING_NO" _
            & " from " & SOTINVH1 & " SOTINVH1,SOTCART1 where SOTCART1.PICK_NO = SOTINVH1.PICK_NO group by INV_NO"
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim INV_NO As String = row.Item("INV_NO")
            Dim rowSOTINVH1 As DataRow = dst.Tables("SOTINVH1").Rows.Find(New String() {IIf(pro_forma, "P", "I"), INV_NO})
            rowSOTINVH1.Item("CART_TRACKING_NO") = row.Item("CART_TRACKING_NO")
        Next

        For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select("INV_TYPE = 'C' AND ISNULL(ORDR_UNIT_PRICE, 0) = 0 ")
            rowSOTINVH2.Item("DND") = "1"
        Next

        Dim CM As String = ""
        For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("")
            Dim INV_TYPE As String = rowSOTINVH1.Item("INV_TYPE")
            Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
            Dim INV_SALES As Decimal = Val(rowSOTINVH1.Item("INV_SALES") & "")
            Dim INV_SALES_CURR As Decimal = Val(rowSOTINVH1.Item("INV_SALES_CURR") & "")
            Dim CURR_CODE As String = rowSOTINVH1.Item("CURR_CODE")
            Dim cINV_SALES As Decimal = 0
            Dim cINV_SALES_CURR As Decimal = 0
            For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select("INV_TYPE = '" & INV_TYPE & "' and INV_NO = '" & INV_NO & "'")
                Dim ORDR_QTY_SHIP As Int64 = Val(rowSOTINVH2.Item("ORDR_QTY_SHIP") & "")
                Dim ORDR_UNIT_PRICE As Decimal = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE") & "")
                Dim ORDR_UNIT_PRICE_CURR As Decimal = Val(rowSOTINVH2.Item("ORDR_UNIT_PRICE_CURR") & "")
                If CURR_CODE = "USD" Then
                    If ORDR_UNIT_PRICE <> ORDR_UNIT_PRICE_CURR Then
                        CM &= vbCrLf & "USD Price match issue - Invoice" & INV_NO
                    End If
                End If
                cINV_SALES += ORDR_QTY_SHIP * ORDR_UNIT_PRICE
                cINV_SALES_CURR += ORDR_QTY_SHIP * ORDR_UNIT_PRICE_CURR
            Next

            If System.Math.Abs(INV_SALES - cINV_SALES) > 0.01 Then
                CM &= vbCrLf & "Invoice Sub-Total - Invoice" & INV_NO
            End If

            If System.Math.Abs(INV_SALES_CURR - cINV_SALES_CURR) > 0.01 Then
                CM &= vbCrLf & "Invoice Sub-Total - Invoice" & INV_NO
            End If
        Next

        If CM <> "" Then
            MsgBox(Mid(Mid(CM, 1, 1000), 3), vbOKOnly, "Problems with Price or Sub-Totals - Please call ABS")
            RWU = "N"
        End If

        'Fill_Records("ARTCUSTZ")
        EnforceConstraints(True)

        Prepare_Invoice_Header(sqlw)
    End Sub

    Sub Prepare_Invoice_Header(sqlw As String)

        Check_Invoice_Totals()

        ' Set flag if Invoice has Misc Charges
        'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH5"), "SO_ORDER_NO").Rows
        '    Dim SO_ORDER_NO As String = row.Item("SO_ORDER_NO")
        '    Dim rowSOTINVH1 = dst.Tables("SOTINVH1").Rows.Find(New Object() {SO_ORDER_NO})
        '    rowSOTINVH1.ITEM("MISC") = "1"
        'Next

        'For Each row As DataRow In ASCDATA1.SelectDistinct(dst.Tables("SOTINVH1"), "SALES_DIVISION_CODE").Rows
        '    Dim rowSOTSDIV1 As DataRow = dst.Tables("SOTSDIV1").Rows.Find(row.Item("ORDR_DIV_CODE"))
        '    If rowSOTSDIV1.Item("DIVISION_LOGO_FILENAME") & "" <> "" And rowSOTSDIV1.Item("DIVISION_LOGO").ToString & "" = "" Then
        '        Dim DIVISION_LOGO_FILENAME As String = ASCMAIN1.Folders("Images") & "ABS\" & rowSOTSDIV1.Item("DIVISION_LOGO_FILENAME")
        '        If My.Computer.FileSystem.FileExists(DIVISION_LOGO_FILENAME) Then
        '            rowSOTSDIV1.Item("DIVISION_LOGO") = ASCMAIN1.GetImageData(DIVISION_LOGO_FILENAME)
        '        End If
        '    End If
        'Next


        ' Load SOTINVH1 - Based on Run-Time Options

        'SOTINVP1 = ASCMAIN1.Temp_Table(sqlw)
        'ASCDATA1.ExecuteSQL("Alter Table " & SOTINVP1 & " Add Primary Key (INV_TYPE, INV_NO)")

    End Sub

    Sub Create_Consolidated_Invoice()

        If Absx1.chkFor("CHKCONS_INV").Checked Then

            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("INV_NO_CONS <> ''")
                Dim INV_NO As String = rowSOTINVH1.Item("INV_NO")
                Dim INV_NO_CONS As String = rowSOTINVH1.Item("INV_NO_CONS")
                Dim ORDR_NO As String = rowSOTINVH1.Item("ORDR_NO")
                ASCDATA1.DeleteRows(dst.Tables("SOTINVH2"), "INV_NO = '" & INV_NO & "'")
                ASCDATA1.DeleteRows(dst.Tables("SOTORDR2"), "ORDR_NO = '" & ORDR_NO & "'")
            Next
            ASCDATA1.DeleteRows(dst.Tables("SOTINVH1"), "INV_NO_CONS <> '' and INV_NO = INV_NO_CONS")

            For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("INV_NO = INV_NO_CONS")
                Dim INV_NO_CONS As String = rowSOTINVH1.Item("INV_NO_CONS")

                Dim TOTAL_CARTONS As Int64 = 0
                ASCMAIN1.sql = "Select " _
                    & "  Sum (ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
                    & " from SOTINVH2,SOTINVH1" & vbCrLf _
                    & " where SOTINVH1.INV_NO_CONS = '" & INV_NO_CONS & "'" _
                    & "   and SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE and SOTINVH2.INV_NO = SOTINVH1.INV_NO"
                Dim rowSOTINVH2_Totals As DataRow = ASCDATA1.GetDataRow
                Dim TOTAL_UNITS As Int64 = Val(rowSOTINVH2_Totals.Item("ORDR_QTY_SHIP") & "")

                ASCMAIN1.sql = "Select " _
                    & "  Sum (INV_SALES) INV_SALES" & vbCrLf _
                    & ", Sum (INV_FREIGHT) INV_FREIGHT" & vbCrLf _
                    & ", Sum (INV_MISC_CHG) INV_MISC_CHG" & vbCrLf _
                    & ", Sum (INV_TOTAL_AMOUNT) from SOTINVH1" & vbCrLf _
                    & " where SOTINVH1.INV_NO_CONS = '" & INV_NO_CONS & "'"
                Dim rowSOTINVH1_CONS As DataRow = ASCDATA1.GetDataRow

                rowSOTINVH1.Item("INV_SALES") = Val(rowSOTINVH2_Totals.Item("INV_SALES") & "")
                rowSOTINVH1.Item("INV_FREIGHT") = Val(rowSOTINVH2_Totals.Item("INV_FREIGHT") & "")
                rowSOTINVH1.Item("INV_MISC_CHG") = Val(rowSOTINVH2_Totals.Item("INV_MISC_CHG") & "")
                rowSOTINVH1.Item("INV_TOTAL_AMOUNT") = Val(rowSOTINVH2_Totals.Item("INV_TOTAL_AMOUNT") & "")

                rowSOTINVH1.Item("TOTAL_UNITS") = TOTAL_UNITS

                ASCMAIN1.sql = "Select ITEM_CODE, CUST_CODE, ORDR_YYYYPP_UPDATED, ITEM_CUST_CODE" & vbCrLf _
                    & ", SUM(NVL(ORDR_UNIT_COST,0) * NVL(ORDR_QTY_SHIP,0)) AS ORDR_UNIT_COST_X" & vbCrLf _
                    & ", SUM(NVL(ORDR_UNIT_PRICE,0) * NVL(ORDR_QTY_SHIP,0)) AS ORDR_UNIT_PRICE_X" & vbCrLf _
                    & ", SUM(ORDR_QTY_SHIP) AS ORDR_QTY_SHIP From SOTINVH2" & vbCrLf _
                    & " where INV_TYPE = 'I' and INV_NO in" & vbCrLf _
                    & " (Select INV_NO from SOTINVH1 where INV_NO_CONS = '" & INV_NO_CONS & "')" & vbCrLf _
                    & " group by " & vbCrLf _
                    & " ITEM_CODE, CUST_CODE, ORDR_YYYYPP_UPDATED, ITEM_CUST_CODE"

                Dim INV_LNO As Int32 = 0
                For Each rowSOTINVH2_CONS As DataRow In ASCDATA1.GetDataTable.Select _
                        ("", "ITEM_COD, CUST_CODE, ORDR_YYYYPP_UPDATED, ITEM_CUST_CODE")
                    INV_LNO += 1
                    Dim rowSOTINVH2 As DataRow = dst.Tables("SOTINVH2").NewRow
                    With rowSOTINVH2
                        .Item("INV_TYPE").Value = rowSOTINVH1.Item("INV_TYPE") & ""
                        .Item("INV_NO").Value = rowSOTINVH1.Item("INV_NO") & ""
                        .Item("INV_LNO").Value = INV_LNO
                        .Item("ITEM_CODE").Value = rowSOTINVH2_CONS.Item("ITEM_CODE") & ""
                        Dim ORDR_QTY_SHP As Int64 = Val(rowSOTINVH2_CONS.Item("ORDR_QTY_SHIP") & "")
                        .Item("ORDR_UNIT_COST").Value = Val(rowSOTINVH2_CONS.Item("ORDR_UNIT_COST_X") & "") / ORDR_QTY_SHP
                        .Item("ORDR_UNIT_PRICE").Value = Val(rowSOTINVH2_CONS.Item("ORDR_UNIT_PRICE_X") & "") / ORDR_QTY_SHP
                        .Item("ORDR_QTY_SHIP").Value = ORDR_QTY_SHP
                        .Item("CUST_CODE").Value = rowSOTINVH2_CONS.Item("CUST_CODE") & ""
                        .Item("ORDR_YYYYPP_UPDATED").Value = rowSOTINVH2_CONS.Item("ORDR_YYYYPP_UPDATED") & ""
                        .Item("ITEM_CUST_CODE").Value = rowSOTINVH2_CONS.Item("ITEM_CUST_CODE") & ""
                        .Item("ORDR_NO").Value = rowSOTINVH1.Item("ORDR_NO") & ""
                    End With
                    dst.Tables("SOTINVH2").Rows.Add(rowSOTINVH2)
                Next

                ASCMAIN1.sql = "Select ITEM_CODE, ITEM_DESC" & vbCrLf _
                   & " from SOTORDR2 where ORDR_NO in " & vbCrLf _
                   & " (Select ORDR_NO from SOTINVH1 where INV_NO_CONS = '" & INV_NO_CONS & "')" & vbCrLf _
                   & " and ORDR_QTY_SHIP <> 0" & vbCrLf _
                   & " group by " & vbCrLf _
                   & " ITEM_CODE, ITEM_DESC"

                Dim ORDR_LNO As Int32 = 0
                For Each rowSOTORDR2_CONS As DataRow In ASCDATA1.GetDataTable.Select _
                        ("", "ITEM_CODE, CUST_CODE, ORDR_YYYYPP_UPDATED")
                    ORDR_LNO += 1
                    Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").NewRow
                    With rowSOTORDR2
                        .Item("ORDR_NO") = rowSOTINVH1.Item("ORDR_NO") & ""
                        .Item("ORDR_LNO") = ORDR_LNO
                        .Item("ITEM_CODE") = rowSOTORDR2_CONS.Item("ITEM_CODE") & ""
                        .Item("ITEM_DESC") = rowSOTORDR2_CONS.Item("ITEM_DESC") & ""
                    End With
                    dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
                Next
            Next
        End If
    End Sub

    Sub Check_Invoice_Totals()

        'For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select("CURR_CODE <> 'USD'")
        '    rowSOTINVH1.Item("ORDR_DIV_CODE_0") = rowSOTINVH1.Item("ORDR_DIV_CODE")
        '    rowSOTINVH1.Item("ORDR_DIV_CODE_R") = rowSOTINVH1.Item("ORDR_DIV_CODE")
        '    rowSOTINVH1.Item("ORDR_DIV_CODE_R") = "E"
        'Next

        'For Each rowSOTINVH1 As DataRow In dst.Tables("SOTINVH1").Select
        '    Dim DIFF As Decimal = 0
        '    DIFF = Val(rowSOTINVH1.Item("ORDR_AMT") & "") _
        '    - (Val(rowSOTINVH1.Item("ORDR_AMT_GROSS") & "") - Val(rowSOTINVH1.Item("ORDR_AMT_ALLOW") & ""))
        '    If System.Math.Abs(System.Math.Round(DIFF, 2)) > 0.01 Then
        '        MsgBox("Order No " & rowSOTINVH1.Item("SO_ORDER_NO") & " does not foot", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        '    DIFF = Val(rowSOTINVH1.Item("ORDR_AMT") & "") + Val(rowSOTINVH1.Item("ORDR_MISC_CHG") & "") + Val(rowSOTINVH1.Item("ORDR_WD_CHG") & "") <> Val(rowSOTINVH1.Item("ORDR_TOTAL_AMT") & "")
        '    If System.Math.Abs(System.Math.Round(DIFF, 2)) > 0.01 Then
        '        MsgBox("Order No " & rowSOTINVH1.Item("SO_ORDER_NO") & " does not total", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        'Next
        'For Each rowSOTINVH2 As DataRow In dst.Tables("SOTINVH2").Select
        '    If Val(rowSOTINVH2.Item("CASES") & "") <> Val(rowSOTINVH2.Item("QTY_CASES") & "") Then
        '        MsgBox("Order No " & rowSOTINVH2.Item("SO_ORDER_NO") & ", Line " & rowSOTINVH2.Item("SO_ORDER_LNO") & " does not foot", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        '    If System.Math.Round(Val(rowSOTINVH2.Item("UNITS") & ""), 0) _
        '    <> System.Math.Round(Val(rowSOTINVH2.Item("QTY_UNITS") & ""), 0) Then
        '        MsgBox("Order No " & rowSOTINVH2.Item("SO_ORDER_NO") & ", Line " & rowSOTINVH2.Item("SO_ORDER_LNO") & " does not foot", MsgBoxStyle.OkOnly, "Please Contact ABS")
        '    End If
        'Next

    End Sub

    Sub Email_Invoice(ByVal rowSOTINVH1 As DataRow, ByVal rowARTCUST1 As DataRow, ByVal attachment As String)
        Me.Cursor = Cursors.WaitCursor

        Using frmTAFSEND1 As New TAFSEND1(Me)

            With frmTAFSEND1
                .EMAIL_KEY = "INV"
                .SEND_TO = rowARTCUST1.Item("CUST_EMAIL_TO") & ""
                If ASCMAIN1.USER_EMAIL = "" Then
                    .SEND_FROM = "noreply" & "@" & ASCMAIN1.rowASTPARM1.Item("AS_PARM_DEFAULT_EMAIL_DOMAIN")
                Else
                    .SEND_FROM = ASCMAIN1.USER_EMAIL
                End If
                .SEND_FROM_NAME = ASCMAIN1.USER_NAME
                If rowARTCUST1.Item("CUST_EMAIL_CC") & "" <> "" Then
                    .SEND_CC = rowARTCUST1.Item("CUST_EMAIL_CC") & ""
                End If

                Dim customInfo As String = rowARTCUST1.Item("CUSTOM_SUBJECT") & " "
                Dim companyName As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " "
                Dim custPO As String = rowSOTINVH1.Item("ORDR_CUST_PO") & " "
                Dim invNo As String = "INV " & rowSOTINVH1.Item("ORDR_INV_NO") & " "
                Dim invDate As String = rowSOTINVH1.Item("ORDR_INV_DATE") & ""

                Dim subjectLine As String = customInfo & companyName & custPO & invNo & invDate

                .SEND_SUBJECT = subjectLine

                Dim sal As String = ""
                If rowARTCUST1.Item("CUST_SALUTATION") & "" <> "" Then
                    sal = rowARTCUST1.Item("CUST_SALUTATION") & "," & vbCrLf
                Else
                    sal = "To whom it may concern," & vbCrLf
                End If

                Dim body As String = ""
                If "" <> "" Then
                    body = "Please find your invoice attached."
                Else
                    body = rowARTCUST1.Item("CUST_BILLING_NOTE") & "" <> ""
                End If

                .SEND_BODY = sal & body
                .SEND_ATTACHMENT = attachment
                .SEND_METHOD = "E"
                .SEND_ENTITY_CAPTION = "Sold-To"
                .SEND_ENTITY_TABLE = "ARTCUST1"
                .SEND_ENTITY_KEY = rowSOTINVH1.Item("CUST_CODE")
                .SEND_ENTITY_NAME = rowARTCUST1.Item("CUST_NAME") & ""

                .Send_email_automatically(False)

                If .SEND_STATUS <> "S" Then
                    SOCMAIN1.Record_Event_SOTORDR1(rowSOTINVH1.Item("SO_ORDER_NO"), DATETIME_STAMP, ASCMAIN1.USER_ID, "E", "Emailed Invoice to " & .SEND_TO, rowSOTINVH1.Item("CUST_CODE"))
                Else
                    MsgBox("Error Occured: Could Not Send Email for Invoice: " & rowSOTINVH1.Item("ORDR_INV_NO"), MsgBoxStyle.OkOnly, "Error")
                End If
            End With
        End Using
    End Sub
End Class