Imports System.Math

Public Class SORSLSJ1

    Dim xRYP0_legend As String
    Dim xRYP1_legend As String
    Dim xRYP0 As String
    Dim xRYP1 As String
    Dim xDTE0 As Date
    Dim xDTE1 As Date
    Dim REGISTER_DATE As Date

    Dim SOTINVH1 As String
    Dim SOTINVHS As String
    Dim SOTINVHG As String
    Dim SOTINVHT As String

    Dim SATSLSXC As String = ""
    Dim SATSLSXI As String = ""
    Dim SATSLSXS As String = ""

    Dim sqlSalesByState As String = ""


    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Call Get_PARM("SOTPARM1")
        Call Get_PARM("ARTPARM1")

        Absx1.optFor("RANGE").CheckedIndex = 2
        Absx1.optFor("RANGE").CheckedIndex = 2
        Set_cmbYP("RYP0", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)

        dteINV_DATE_CUTOFF.DateTime = CDate(Now + ASCMAIN1.NowTSD).Date.AddDays(-1)

        grpPERIOD_RANGE.Visible = False
        grpDATE_RANGE.Visible = False
        grpDATE_RANGE.Left = grpPERIOD_RANGE.Left

        If MENU_ITEM_PP = "CJ" Then
            optRANGE.Value = "P"

            If ASCMAIN1.EOM = "1" Then
                grpSelectBy.Visible = False
            Else
                Absx1.optFor("RANGE").Items.RemoveAt(2)
            End If

            grpInclude.Visible = False

            If ASCMAIN1.EOM = "1" Then
                Absx1.cmbFor("RYP0").Value = ASCMAIN1.CYP
                Absx1.cmbFor("RYP1").Value = ASCMAIN1.CYP
                grpPERIOD_RANGE.Enabled = False
            End If
        End If

        If ASCMAIN1.CLIENT = "INT" Then
        Else
            TAC.TACMAIN1.Update_Forex()
        End If

        grpAudit.Left = grpSelectBy.Left
        grpAudit.Top = chkINV_DATE_CUTOFF.Top
        grpAudit.Visible = False

        'If ASCMAIN1.DBS_SERVER = "INT" Then
        '    For Each item As ValueListItem In optRANGE.Items
        '        If item.DisplayText = "Audit No" Then
        '            optRANGE.Items.Remove(item)
        '        End If
        '    Next
        'End If

    End Sub

    Protected Overrides Sub Build_Workfile()

        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
            TAC.SACMAIN1.Create_Sales_Extract_Tables(Me, True, SATSLSXC, SATSLSXI, SATSLSXS)
        End If

        RWU = "R"
        EnforceConstraints(False)
        SUBT = ""
        Dim generateSalesByStateExtract As Boolean = (chkSalesByState.Visible And chkSalesByState.Checked)
        Dim sqlw As String = "NVL(SOTINVH1.REGISTER_IND,'0') = '0'"
        Dim sqlwSalesByState As String = ""

        REGISTER_DATE = DATETIME_STAMP.Date '  Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy")
        If chkINV_DATE_CUTOFF.Checked Then
            REGISTER_DATE = dteINV_DATE_CUTOFF.Value
            SUBT = "Invoice Cut-Off Date " & Format(REGISTER_DATE, "MM/dd/yyyy")
            sqlw &= " and INV_DATE <= '" & Format(REGISTER_DATE, "dd-MMM-yyyy") & "'"
        End If

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "Invoices Posted " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "Invoices Posted between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = "SOTINVH1.REGISTER_IND = '1' and SOTINVH1.REGISTER_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"
            'sqlw = "SOTINVH1.INV_DATE between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'"

            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "P" Then
            xRYP0_legend = Absx1.cmbFor("RYP0").Value
            xRYP0 = Mid(xRYP0_legend, 1, 4) & Mid(xRYP0_legend, 6, 2)
            xRYP1_legend = Absx1.cmbFor("RYP1").Value
            xRYP1 = Mid(xRYP1_legend, 1, 4) & Mid(xRYP1_legend, 6, 2)
            If xRYP0 = xRYP1 Then
                SUBT = "Invoices Posted in " & xRYP0_legend
            Else
                SUBT = "Invoices Posted between " & xRYP0_legend & " and " & xRYP1_legend
            End If
            sqlw = "SOTINVH1.REGISTER_IND = '1' and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & xRYP0 & "' and '" & xRYP1 & "'"
            If generateSalesByStateExtract Then
                sqlwSalesByState = "SOTORDR5.ORDR_NO = SOTINVH1.ORDR_NO AND SOTORDR5.CUST_ADDR_TYPE = 'ST'" & vbCrLf _
                    & " AND ICTWHSE1.WHSE_CODE = SOTINVH1.WHSE_CODE AND SOTINVH1.INV_TYPE = 'I' and SOTINVH1.ORDR_YYYYPP_UPDATED between '" & xRYP0 & "' and '" & xRYP1 & "'"
            End If
            'sqlw = "SOTINVH1.ORDR_YYYYPP_UPDATED between '" & xRYP0 & "' and '" & xRYP1 & "'"
            RWU = "N"

        ElseIf optRANGE.Value = "A" Then
            SUBT = $"Invoices for Warehouse Audit '{txtAUDIT_NO.Text}'"
            sqlw = $"SOTINVH1.INV_NO IN (
                                SELECT SOTINVH1.INV_NO
                                FROM SOTSHIP1, SOTSVIA1, SOTORDR0, ARTCUST1, ICTWHSE1, WHTAUDT1, SOTPICK1, SOTINVH1
                                WHERE SOTSHIP1.SHIP_STATUS = 'F'
                                AND SOTSHIP1.SHIP_DATE_SHIPPED >= TRUNC(WHTAUDT1.SHIP_DATE_FROM)
                                AND TRUNC(SOTSHIP1.SHIP_DATE_SHIPPED) < TRUNC(WHTAUDT1.SHIP_DATE_TO)
                                AND SOTSHIP1.SHIP_VIA_CODE = SOTSVIA1.SHIP_VIA_CODE (+)
                                AND SOTSHIP1.BILL_OF_LADING_NO IS NOT NULL
                                AND SOTSHIP1.ORDR_GROUP_NO = SOTORDR0.ORDR_GROUP_NO
                                AND SOTORDR0.CUST_CODE = ARTCUST1.CUST_CODE
                                AND ARTCUST1.TRADE_CLASS_CODE <> 'HDQ'
                                AND SOTSHIP1.WHSE_CODE = ICTWHSE1.WHSE_CODE
                                AND ICTWHSE1.LP_CODE IS NOT NULL
                                AND SOTPICK1.SHIP_BOL_NO = SOTSHIP1.SHIP_BOL_NO
                                AND SOTPICK1.PICK_STATUS = 'F'
                                AND SOTPICK1.INV_NO = SOTINVH1.INV_NO
                                AND WHTAUDT1.AUDIT_NO = '{txtAUDIT_NO.Text}'
                                )"
            RWU = "N"
        End If

        If MENU_ITEM_PP = "CJ" Then
            If ASCMAIN1.EOM = "1" Then
                RWU = "R"
            Else
                RWU = "N"
            End If
        End If

        sqlw &= SQL_in("CUST_CODE", "SOTINVH1.CUST_CODE")
        sqlw &= SQL_in("ORDR_TYPE_CODE", "SOTINVH1.ORDR_TYPE_CODE")
        sqlw &= SQL_in("SALES_DIVISION_CODE", "SOTINVH1.SALES_DIVISION_CODE")
        sqlw &= SQL_in("INV_NO", "SOTINVH1.INV_NO")
        sqlw &= SQL_in("STAX_CODE", "SOTINVH1.STAX_CODE")
        sqlw &= SQL_in("TRADE_CLASS_CODE", "ARTCUST1.TRADE_CLASS_CODE")
        sqlw &= SQL_in("CUST_STATE", "ARTCUST1.CUST_STATE")
        sqlw &= SQL_in("REGISTER_XNO", "SOTINVH1.REGISTER_XNO")

        If generateSalesByStateExtract Then
            sqlwSalesByState &= Replace(SQL_in("CUST_STATE", "ARTCUST1.CUST_STATE"), "ARTCUST1", "SOTORDR5")
            sqlSalesByState = "SELECT " & vbCrLf _
            & " SOTINVH1.INV_NO, 'Sales Invoice' DocType,  SOTINVH1.INV_DATE, SOTINVH1.CUST_CODE, '' TaxCode, '' ItemCode, '' ItemDescription, 1 Qty" & vbCrLf _
            & " , SOTINVH1.INV_SALES, '' Discounts, '' Description" & vbCrLf _
            & " , SOTORDR5.CUST_ADDR1 || SOTORDR5.CUST_ADDR2 DestAddress, SOTORDR5.CUST_CITY DestCity, SOTORDR5.CUST_STATE DestRegion, SOTORDR5.CUST_ZIP_CODE DestPostalCode" & vbCrLf _
            & " , ICTWHSE1.WHSE_ADDR1 || ICTWHSE1.WHSE_ADDR2 OrigAddress, ICTWHSE1.WHSE_CITY OrigCity, ICTWHSE1.WHSE_STATE OrigRegion, ICTWHSE1.WHSE_ZIP_CODE OrigPostalCode, ICTWHSE1.WHSE_COUNTRY OrigCountry" & vbCrLf _
            & " , '' TaxCollectedandRemitted, '' TaxCollectedNOTRemitted, '' UseTaxPaid" & vbCrLf _
            & " FROM SOTINVH1, SOTORDR5, ICTWHSE1" & vbCrLf _
            & " Where " & sqlwSalesByState
        End If

        If Not Absx1.chkFor("INV_TYPE_I").Checked Or Not Absx1.chkFor("INV_TYPE_C").Checked Then
            If Absx1.chkFor("INV_TYPE_I").Checked Then
                sqlw &= " and SOTINVH1.INV_TYPE = 'I'"
            End If
            If Absx1.chkFor("INV_TYPE_C").Checked Then
                sqlw &= " and SOTINVH1.INV_TYPE = 'C'"
            End If
        End If

        ASCMAIN1.sql = "Select SOTINVH1.*, NVL(SOTINVH1.INV_PRO_NO, SOTINVH1.INV_BOL_NO) SHIP_REF, ARTCUST1.CUST_NAME 
                        from SOTINVH1, ARTCUST1 
                        where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE 
                        and " & sqlw
        SOTINVH1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_NO)")

        If ASCMAIN1.CLIENT = "INT" Then
            ASCMAIN1.sql = "Update " & SOTINVH1 & " Set SALES_DIVISION_CODE = 'IP1'"
            ASCDATA1.ExecuteSQL()
        End If

        ASCMAIN1.sql = "Select SOTINVH1.* " _
        & " from " & SOTINVH1 & " SOTINVH1,ARTCUST1 " _
        & " where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE"

        Create_TDA(dst.Tables.Add, "SOTINVH1", "**", 0, False)
        Fill_Records("SOTINVH1")

        ASCMAIN1.sql = "Select * from SOTINVHS where ROWNUM < 1"
        SOTINVHS = ASCMAIN1.Temp_Table

        'ASCMAIN1.sql = "Select SOTINVHS.*,SOTINVH2.INV_NO, SOTINVH2.INV_LNO, SOTINVH2.ACCT_CODE_SLS, SOTINVH2.ACCT_CODE_CGS" & vbCrLf _
        '    & " from SOTINVHS,SOTINVH2 where ROWNUM < 1"
        ASCMAIN1.sql = "Select * from SOTINVHT where ROWNUM < 1"
        SOTINVHT = ASCMAIN1.Temp_Table
        ASCMAIN1.sql = "Alter Table " & SOTINVHT & " Add Primary Key (INV_TYPE, INV_NO, INV_LNO)"
        ASCDATA1.ExecuteSQL()


        ' Sales & CGS Summary

        Dim sqlXNO As String = "'" & XNO & "'"
        If RWU = "N" And ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS And MENU_ITEM_PP <> "CJ" Then
            sqlXNO = "SOTINVH1.REGISTER_XNO"
        End If

        Dim sqlSEG2_CODE As String = "NVL(SOTTYPE1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')"

        Dim sqlSEG3_CODE As String = ""
        If ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG3") & "" = "1" Then
            sqlSEG3_CODE = "NVL(SOTTCLS1.SEG3_CODE,ARTCUST1.TRADE_CLASS_CODE)"
        Else
            sqlSEG3_CODE = "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'"
        End If

        Dim sqlSEG4_CODE As String = ""
        If ROWs("SOTPARM1").Item("SO_PARM_DTL_SEG4") & "" = "1" Then
            sqlSEG4_CODE = "NVL(ICTCOLL1.SEG4_CODE,ICTITEM1.COLLECTION_CODE)"
        Else
            sqlSEG4_CODE = "'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'"
        End If

        ' IMPORTANT - there should be no value for CGS for DIFs - ie, trans that do not have a whse

        ' note - similar/same sql as used for SOTINVHT below
        ASCMAIN1.sql = "" _
            & "Select " & sqlXNO & " REGISTER_XNO, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
            & ", SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & ", " & sqlSEG2_CODE & " SEG2_CODE" & vbCrLf _
            & ", " & sqlSEG3_CODE & " SEG3_CODE" & vbCrLf _
            & ", " & sqlSEG4_CODE & " SEG4_CODE" & vbCrLf _
            & ", SOTINVH2.INV_TYPE, ICTITEM1.PROD_CODE, SOTINVH1.EVENT_CODE" & vbCrLf _
            & ", DECODE(NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0,'1','0') ORDR_NC" & vbCrLf _
            & ", SUM (NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0)) SLS " & vbCrLf _
            & ", SUM (CASE WHEN SOTINVH2.INV_TYPE = 'C' AND ICTCOLL1.BRAND_CODE in ('KSP','LCS') AND ICTITEM1.ITEM_DESC LIKE 'XXX %' THEN 0 ELSE " & vbCrLf _
            & "       NVL(SOTINVH2.ORDR_QTY_SHIP,0) * DECODE(SOTINVH2.WHSE_CODE,NULL,0,NVL(SOTINVH2.ITEM_UNIT_COST,0)) END) CGS " & vbCrLf _
            & ", CASE WHEN SOTINVH2.INV_TYPE = 'C' AND ICTCOLL1.BRAND_CODE in ('KSP','LCS') AND ICTITEM1.ITEM_DESC LIKE 'XXX %' THEN '1' ELSE '0' END RTRN_AS_PO_REC" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1,SOTINVH2,ICTITEM1,SOTTYPE1,ARTCUST1,SOTTCLS1,ICTCOLL1 " & vbCrLf _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   And SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   And SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & "   And ICTITEM1.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   And ARTCUST1.CUST_CODE (+) = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   And SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   And ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " group by " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
            & ", SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & ", " & sqlSEG2_CODE & vbCrLf _
            & ", " & sqlSEG3_CODE & vbCrLf _
            & ", " & sqlSEG4_CODE & vbCrLf _
            & ", SOTINVH2.INV_TYPE, ICTITEM1.PROD_CODE, SOTINVH1.EVENT_CODE" & vbCrLf _
            & ", DECODE(NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0,'1','0')" & vbCrLf _
            & ", CASE WHEN SOTINVH2.INV_TYPE = 'C' AND ICTCOLL1.BRAND_CODE in ('KSP','LCS') AND ICTITEM1.ITEM_DESC LIKE 'XXX %' THEN '1' ELSE '0' END"
        ASCDATA1.ExecuteSQL("Insert into " & SOTINVHS & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select * from " & SOTINVHS
        Create_TDA(dst.Tables.Add, "SOTINVHS", "**", 0, False)

        ' note - similar/same sql as used for SOTINVHS above
        ASCMAIN1.sql = "" _
            & "Select " & sqlXNO & " REGISTER_XNO, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
            & ", SOTINVH1.SALES_DIVISION_CODE, SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & ", " & sqlSEG2_CODE & " SEG2_CODE" & vbCrLf _
            & ", " & sqlSEG3_CODE & " SEG3_CODE" & vbCrLf _
            & ", " & sqlSEG4_CODE & " SEG4_CODE" & vbCrLf _
            & ", SOTINVH2.INV_TYPE, ICTITEM1.PROD_CODE, SOTINVH1.EVENT_CODE" & vbCrLf _
            & ", DECODE(NVL(SOTINVH2.ORDR_UNIT_PRICE,0),0,'1','0') ORDR_NC" & vbCrLf _
            & ", NVL(SOTINVH2.ORDR_QTY_SHIP,0) * NVL(SOTINVH2.ORDR_UNIT_PRICE,0) SLS " & vbCrLf _
            & ", CASE WHEN SOTINVH2.INV_TYPE = 'C' AND ICTCOLL1.BRAND_CODE in ('KSP','LCS') AND ICTITEM1.ITEM_DESC LIKE 'XXX %' THEN 0 ELSE " & vbCrLf _
            & "       NVL(SOTINVH2.ORDR_QTY_SHIP,0) * DECODE(SOTINVH2.WHSE_CODE,NULL,0,NVL(SOTINVH2.ITEM_UNIT_COST,0)) END CGS " & vbCrLf _
            & ", SOTINVH2.INV_NO, SOTINVH2.INV_LNO, NULL ACCT_CODE_SLS, NULL ACCT_CODE_CGS" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1,SOTINVH2,ICTITEM1,SOTTYPE1,ARTCUST1,SOTTCLS1,ICTCOLL1 " & vbCrLf _
            & " where SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" & vbCrLf _
            & "   and SOTINVH2.INV_NO = SOTINVH1.INV_NO" & vbCrLf _
            & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE (+) = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE (+) = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE"
        ASCDATA1.ExecuteSQL("Insert into " & SOTINVHT & " " & ASCMAIN1.sql)

        ASCMAIN1.sql = "Select X.*" & vbCrLf _
            & ", SOTMISC1.MISC_CHG_DESC, SOTMISC1.ACCT_CODE, SOTMISC1.MISC_GP" & vbCrLf _
            & " from (Select MISC_CHG_CODE, Sum (INV_MISC_CHG) INV_MISC_CHG from (" & vbCrLf _
            & "Select '1', SOTINVH1.MISC_CHG_CODE, SUM (SOTINVH1.INV_MISC_CHG) INV_MISC_CHG" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1 " & vbCrLf _
            & " where SOTINVH1.INV_MISC_CHG <> 0" & vbCrLf _
            & " group by SOTINVH1.MISC_CHG_CODE" & vbCrLf
        ASCMAIN1.sql &= "" _
            & ") group by MISC_CHG_CODE" & vbCrLf _
            & ") X, SOTMISC1 where SOTMISC1.MISC_CHG_CODE (+) = X.MISC_CHG_CODE"
        Create_TDA(dst.Tables.Add, "SOTINVHM", "**", 0, False, "", 1)
        Fill_Records("SOTINVHM")

        ASCMAIN1.sql = "Select X.*" & vbCrLf _
            & ", ARTSTAX1.STAX_DESC, ARTSTAX1.ACCT_CODE" & vbCrLf _
            & " from (" & vbCrLf _
            & "Select SOTINVH1.CUST_SHIP_TO_STATE, SOTINVH1.STAX_CODE" & vbCrLf _
            & ", SUM (SOTINVH1.INV_STAX) INV_STAX" & vbCrLf _
            & " from " & SOTINVH1 & " SOTINVH1 " & vbCrLf _
            & " where SOTINVH1.INV_STAX <> 0" & vbCrLf _
            & " group by SOTINVH1.CUST_SHIP_TO_STATE, SOTINVH1.STAX_CODE" & vbCrLf _
            & ") X, ARTSTAX1 where ARTSTAX1.STAX_CODE (+) = X.STAX_CODE"

        Create_TDA(dst.Tables.Add, "SOTINVHT", "**", 0, False, "", 2)
        Fill_Records("SOTINVHT")
        ' BE CAREFUL- THERE AS A POOR CHOICE OF TABLE NAMES 
        ' - ORACLE TEMP TABLE SOTINVHT HAS NOTHING TO DO WITH THIS DATA TABLE, WHICH IS USED FOR SALES TAX 


        Create_TDA(dst.Tables.Add, "SOTEVNT1", "*", 0, False)
        Fill_Records("SOTEVNT1")

        Create_TDA(dst.Tables.Add, "ICTCOST1", "*", 0, False)
        Fill_Records("ICTCOST1")

        Create_TDA(dst.Tables.Add, "ICTPROD1", "*", 0, False)
        Fill_Records("ICTPROD1")

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")

        ' DO THIS ALWAYS - THAT IS WHY i ADDED THE "" = ""
        If Absx1.optFor("RANGE").Value = "N" Or "" = "" Then

            ASCMAIN1.sql = "Select * from SOTINVHG where ROWNUM < 1"
            SOTINVHG = ASCMAIN1.Temp_Table

            Dim ACCT_CODE_sql As String
            Dim sqlG As String = ""

            ' Sales / Customer Returns
            ' Note that SEG2_CODE comes from Special Event, Sales Division, Order Type, or Default

            ACCT_CODE_sql = "DECODE(SOTINVHS.EVENT_CODE,NULL,DECODE(SOTINVHS.INV_TYPE,'I',ICTPROD1.PROD_SALES_SHP_ACCT,ICTPROD1.PROD_SALES_RTN_ACCT),SOTEVNT1.ACCT_CODE)"
            Dim sql_SEG3 As String = ""
            Dim sql_SEG4 As String = ""

            sqlG = "" _
                & " Select " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & " REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVHS.OPS_YYYYPP" & vbCrLf _
                & ", " & ACCT_CODE_sql & " ACCT_CODE" & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHS.SEG2_CODE)) SEG2_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END SEG3_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END SEG4_CODE" & vbCrLf _
                & ", SOTINVHS.EVENT_CODE DETL_CVX_REF_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, SOTINVHS.PROD_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * SOTINVHS.SLS) DIST_AMT " & vbCrLf _
                & " from " & SOTINVHS & " SOTINVHS,ICTPROD1,SOTSDIV1,SOTEVNT1" & vbCrLf _
                & " where ICTPROD1.PROD_CODE (+) = SOTINVHS.PROD_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVHS.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTEVNT1.EVENT_CODE (+) = SOTINVHS.EVENT_CODE" & vbCrLf _
                & "   and NVL(SOTINVHS.RTRN_AS_PO_REC,'0') <> '1'" & vbCrLf _
                & " group by " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP, " & ACCT_CODE_sql & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHS.SEG2_CODE))" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END" & vbCrLf _
                & ", SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE, SOTINVHS.EVENT_CODE " & vbCrLf _
                & " having SUM (-1 * SOTINVHS.SLS) <> 0" & vbCrLf _
                & " order by " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP, " & ACCT_CODE_sql & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHS.SEG2_CODE))" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END" & vbCrLf _
                & ", SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE, SOTINVHS.EVENT_CODE "
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)


            Dim ACCT_CODE_INVTY_TRANSFER As String = "152600"
            Dim ACCT_CODE_PO_REC_EXCHANGE As String = "222250"

            sqlG = "" _
                & " Select " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & " REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVHS.OPS_YYYYPP" & vbCrLf _
                & ", '" & ACCT_CODE_PO_REC_EXCHANGE & "' ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", SOTINVHS.EVENT_CODE DETL_CVX_REF_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, SOTINVHS.PROD_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * SOTINVHS.SLS) DIST_AMT " & vbCrLf _
                & " from " & SOTINVHS & " SOTINVHS,ICTPROD1,SOTSDIV1,SOTEVNT1" & vbCrLf _
                & " where ICTPROD1.PROD_CODE (+) = SOTINVHS.PROD_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVHS.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTEVNT1.EVENT_CODE (+) = SOTINVHS.EVENT_CODE" & vbCrLf _
                & "   and NVL(SOTINVHS.RTRN_AS_PO_REC,'0') = '1'" & vbCrLf _
                & " group by " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP" & vbCrLf _
                & ", SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE, SOTINVHS.EVENT_CODE " & vbCrLf _
                & " having SUM (-1 * SOTINVHS.SLS) <> 0" & vbCrLf _
                & " order by " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP" & vbCrLf _
                & ", SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE, SOTINVHS.EVENT_CODE "
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            Dim sqlGT As String = ""

            ', NVL(NVL(NVL(SOTCHAN1.SEG2_CODE,SOTTYPE1.SEG2_CODE),NVL(SOTMISC1.SEG2_CODE,SOTSDIV1.SEG2_CODE)),'000') SEG2_CODE
            ', CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVH1.EVENT_CODE IS NULL THEN NVL(SOTTCLS1.SEG3_CODE,ARTCUST1.TRADE_CLASS_CODE) ELSE '00' END SEG3_CODE
            ', CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVH1.EVENT_CODE IS NULL THEN NVL(ICTCOLL1.SEG4_CODE,NVL(SOTINVH1.COLLECTION_CODE,'000')) ELSE '000' END SEG4_CODE

            sqlGT = "" _
                & " Select SOTINVHT.INV_TYPE, SOTINVHT.INV_NO, SOTINVHT.INV_LNO" & vbCrLf _
                & ", " & Replace(ACCT_CODE_sql, "SOTINVHS", "SOTINVHT") & " ACCT_CODE" & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHT.SEG2_CODE)) SEG2_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHT.EVENT_CODE IS NULL THEN SOTINVHT.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END SEG3_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHT.EVENT_CODE IS NULL THEN SOTINVHT.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END SEG4_CODE" & vbCrLf _
                & " from " & SOTINVHT & " SOTINVHT,ICTPROD1,SOTSDIV1,SOTEVNT1" & vbCrLf _
                & " where ICTPROD1.PROD_CODE (+) = SOTINVHT.PROD_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVHT.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTEVNT1.EVENT_CODE (+) = SOTINVHT.EVENT_CODE" & vbCrLf _
                & "   and (-1 * SOTINVHT.SLS) <> 0"

            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf & sqlGT & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTINVHT & " SOTINVHT Set ACCT_CODE_SLS = R1.ACCT_CODE, SEG2_CODE = R1.SEG2_CODE, SEG3_CODE = R1.SEG3_CODE, SEG4_CODE = R1.SEG4_CODE" & vbCrLf _
                & "    where INV_TYPE = R1.INV_TYPE and INV_NO = R1.INV_NO and INV_LNO = R1.INV_LNO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()


            ASCMAIN1.sql = "" _
                & " Begin" & vbCrLf _
                & " Declare Cursor C1 Is " & vbCrLf _
                & " Select SOTINVHT.INV_TYPE, SOTINVHT.INV_NO, SOTINVHT.INV_LNO" & vbCrLf _
                & " from " & SOTINVHT & " SOTINVHT,SOTINVH2,ICTITEM1" & vbCrLf _
                & " where (SOTINVHT.SEG4_CODE Like 'KSP%' OR SOTINVHT.SEG4_CODE Like 'LCS%') and SOTINVHT.ORDR_TYPE_CODE IN ('DIF','RTN')" & vbCrLf _
                & "   and SOTINVH2.INV_NO = SOTINVHT.INV_NO" & vbCrLf _
                & "   and SOTINVH2.INV_LNO = SOTINVHT.INV_LNO" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_DESC LIKE 'XXX %'" & vbCrLf _
                & "   and SOTINVHT.INV_TYPE = 'C'" & vbCrLf _
                & "   and (-1 * SOTINVHT.SLS) <> 0;" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTINVHT & " SOTINVHT Set ACCT_CODE_SLS = '222250', SEG2_CODE = '000', SEG3_CODE = '000', SEG4_CODE = '000'" & vbCrLf _
                & "    where INV_TYPE = R1.INV_TYPE and INV_NO = R1.INV_NO and INV_LNO = R1.INV_LNO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()


            ' Cost of Goods Sold / Returned

            ACCT_CODE_sql = "DECODE(SOTINVHS.EVENT_CODE,NULL,DECODE(SOTINVHS.ORDR_NC,'1',ICTPROD1.PROD_COSTS_SHP_ACCT_NC,DECODE(SOTINVHS.INV_TYPE,'I',ICTPROD1.PROD_COSTS_SHP_ACCT,ICTPROD1.PROD_COSTS_RTN_ACCT)),SOTEVNT1.ACCT_CODE)"

            sqlG = "" _
                & " Select " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & " REGISTER_XNO, 'OPCJ' JOURNAL_TYPE, SOTINVHS.OPS_YYYYPP" & vbCrLf _
                & ", " & ACCT_CODE_sql & " ACCT_CODE" & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHS.SEG2_CODE)) SEG2_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END SEG3_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END SEG4_CODE" & vbCrLf _
                & ", SOTINVHS.EVENT_CODE DETL_CVX_REF_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, SOTINVHS.PROD_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (SOTINVHS.CGS) DIST_AMT " & vbCrLf _
                & " from " & SOTINVHS & " SOTINVHS,ICTPROD1,SOTSDIV1,SOTEVNT1" & vbCrLf _
                & " where ICTPROD1.PROD_CODE (+) = SOTINVHS.PROD_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVHS.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTEVNT1.EVENT_CODE (+) = SOTINVHS.EVENT_CODE" & vbCrLf _
                & " group by " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP, " & ACCT_CODE_sql & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHS.SEG2_CODE))" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END" & vbCrLf _
                & ", SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE, SOTINVHS.EVENT_CODE " & vbCrLf _
                & " having SUM (SOTINVHS.CGS) <> 0" & vbCrLf _
                & " order by " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP, " & ACCT_CODE_sql & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHS.SEG2_CODE))" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHS.EVENT_CODE IS NULL THEN SOTINVHS.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END" & vbCrLf _
                & ", SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE, SOTINVHS.EVENT_CODE "
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)



            sqlGT = "" _
                & " Select SOTINVHT.INV_TYPE, SOTINVHT.INV_NO, SOTINVHT.INV_LNO" & vbCrLf _
                & ", " & Replace(ACCT_CODE_sql, "SOTINVHS", "SOTINVHT") & " ACCT_CODE" & vbCrLf _
                & ", NVL(SOTEVNT1.SEG2_CODE,NVL(SOTSDIV1.SEG2_CODE,SOTINVHT.SEG2_CODE)) SEG2_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVHT.EVENT_CODE IS NULL THEN SOTINVHT.SEG3_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END SEG3_CODE" & vbCrLf _
                & ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVHT.EVENT_CODE IS NULL THEN SOTINVHT.SEG4_CODE ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END SEG4_CODE" & vbCrLf _
                & " from " & SOTINVHT & " SOTINVHT,ICTPROD1,SOTSDIV1,SOTEVNT1" & vbCrLf _
                & " where ICTPROD1.PROD_CODE (+) = SOTINVHT.PROD_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVHT.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTEVNT1.EVENT_CODE (+) = SOTINVHT.EVENT_CODE" & vbCrLf _
                & "   and (SOTINVHT.CGS) <> 0"
            ASCMAIN1.sql = "" _
                & "Begin" & vbCrLf _
                & " Declare Cursor C1 is " & vbCrLf & sqlGT & ";" & vbCrLf _
                & " Begin" & vbCrLf _
                & "  For R1 in C1 Loop" & vbCrLf _
                & "   Update " & SOTINVHT & " SOTINVHT Set ACCT_CODE_CGS = R1.ACCT_CODE" & vbCrLf _
                & "    where INV_TYPE = R1.INV_TYPE and INV_NO = R1.INV_NO and INV_LNO = R1.INV_LNO;" & vbCrLf _
                & "  End Loop;" & vbCrLf _
                & " End;" & vbCrLf _
                & "End;"
            ASCDATA1.ExecuteSQL()

            ' Inventory

            '            & ", DECODE(SOTINVHS.INV_TYPE,'I','CGS','CGR') DETL_CVX_REF_NO" & vbCrLf _

            sqlG = "" _
                & " SELECT " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & " REGISTER_XNO, 'OPCJ' JOURNAL_TYPE, SOTINVHS.OPS_YYYYPP" & vbCrLf _
                & ", ICTCOST1.ACCT_CODE_ONH ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", NULL DETL_CVX_REF_NO" & vbCrLf _
                & ", 'C' DETL_CVX_TYPE, SOTINVHS.PROD_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * SOTINVHS.CGS) DIST_AMT " & vbCrLf _
                & " FROM " & SOTINVHS & " SOTINVHS,ICTPROD1,ICTCOST1" & vbCrLf _
                & " where ICTPROD1.PROD_CODE (+) = SOTINVHS.PROD_CODE" & vbCrLf _
                & "   and ICTCOST1.COST_CATGY_CODE (+) = ICTPROD1.COST_CATGY_CODE" & vbCrLf _
                & " GROUP BY " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP, ICTCOST1.ACCT_CODE_ONH, SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE" & vbCrLf _
                & " ORDER BY " & Replace(sqlXNO, "SOTINVH1", "SOTINVHS") & ", SOTINVHS.OPS_YYYYPP, ICTCOST1.ACCT_CODE_ONH, SOTINVHS.INV_TYPE, SOTINVHS.PROD_CODE" & vbCrLf
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)


            ' Miscellaneous Charges (Mostly Top Sides)

            If ASCMAIN1.CLIENT = "INTnot" Then
                sql_SEG3 = ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'"
                sql_SEG4 = ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'"
            Else
                sql_SEG3 = ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVH1.EVENT_CODE IS NULL THEN NVL(SOTMISC1.SEG3_CODE," & sqlSEG3_CODE & ") ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END"
                sql_SEG4 = ", CASE WHEN SOTEVNT1.EVENT_BY_SEG4 = '1' OR SOTINVH1.EVENT_CODE IS NULL THEN NVL(SOTMISC1.SEG4_CODE," & Replace(sqlSEG4_CODE, "ICTITEM1", "SOTINVH1") & ") ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' END"
            End If

            sqlG = "" _
                & " SELECT " & sqlXNO & " REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", SOTMISC1.ACCT_CODE" & vbCrLf _
                & ", NVL(NVL(SOTTYPE1.SEG2_CODE,NVL(SOTMISC1.SEG2_CODE,SOTSDIV1.SEG2_CODE)),'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') SEG2_CODE" & vbCrLf _
                & sql_SEG3 & " SEG3_CODE" & vbCrLf _
                & sql_SEG4 & " SEG4_CODE" & vbCrLf _
                & ", DECODE(SOTINVH1.INV_TYPE,'I','MISC-I','MISC-C') DETL_CVX_REF_NO" & vbCrLf _
                & ", 'M' DETL_CVX_TYPE, SOTINVH1.MISC_CHG_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * NVL(SOTINVH1.INV_MISC_CHG,0)) DIST_AMT " & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1,SOTMISC1,SOTSDIV1,SOTTYPE1" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "INTnot",
                      "",
                      ",ARTCUST1,SOTTCLS1,ICTCOLL1,SOTEVNT1" & vbCrLf) _
                & " where SOTMISC1.MISC_CHG_CODE = SOTINVH1.MISC_CHG_CODE" & vbCrLf _
                & "   and SOTINVH1.MISC_CHG_CODE is Not Null" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
                & IIf(ASCMAIN1.CLIENT = "INTnot",
                      "",
                      " and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                      & " and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                      & " and SOTEVNT1.EVENT_CODE (+) = SOTINVH1.EVENT_CODE" & vbCrLf _
                      & " and ICTCOLL1.COLLECTION_CODE (+) = SOTINVH1.COLLECTION_CODE" & vbCrLf) _
                & " GROUP BY " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, SOTMISC1.ACCT_CODE, SOTINVH1.INV_TYPE, NVL(NVL(SOTTYPE1.SEG2_CODE,NVL(SOTMISC1.SEG2_CODE,SOTSDIV1.SEG2_CODE)),'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'), SOTINVH1.MISC_CHG_CODE" & sql_SEG3 & sql_SEG4 & vbCrLf _
                & " ORDER BY " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, SOTMISC1.ACCT_CODE, SOTINVH1.INV_TYPE, NVL(NVL(SOTTYPE1.SEG2_CODE,NVL(SOTMISC1.SEG2_CODE,SOTSDIV1.SEG2_CODE)),'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'), SOTINVH1.MISC_CHG_CODE" & sql_SEG3 & sql_SEG4
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            ' Sales Tax Payable

            sqlG = "" _
                & " SELECT " & sqlXNO & " REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", ARTSTAX1.ACCT_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "' SEG2_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", DECODE(SOTINVH1.INV_TYPE,'I','STAX-I','STAX-C') DETL_CVX_REF_NO" & vbCrLf _
                & ", 'T' DETL_CVX_TYPE, SOTINVH1.CUST_SHIP_TO_STATE DETL_CVX_NO" & vbCrLf _
                & ", SUM (-1 * NVL(SOTINVH1.INV_STAX,0)) DIST_AMT " & vbCrLf _
                & " FROM " & SOTINVH1 & " SOTINVH1,ARTSTAX1,SOTSDIV1" & vbCrLf _
                & " where ARTSTAX1.STAX_CODE (+) = SOTINVH1.STAX_CODE" & vbCrLf _
                & "   and (SOTINVH1.STAX_CODE is Not Null or NVL(SOTINVH1.INV_STAX,0) <> 0)" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & " GROUP BY " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, ARTSTAX1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTINVH1.CUST_SHIP_TO_STATE" & vbCrLf _
                & " ORDER BY " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, ARTSTAX1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTINVH1.CUST_SHIP_TO_STATE"
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)


            ' Freight Income

            If ASCMAIN1.CLIENT = "INT" Then
                sql_SEG3 = ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'"
            Else
                sql_SEG3 = ", CASE WHEN SOTEVNT1.EVENT_BY_SEG3 = '1' OR SOTINVH1.EVENT_CODE IS NULL THEN " & sqlSEG3_CODE & " ELSE '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "' END"
            End If

            Dim sql_SEG2 As String = ""

            If ASCMAIN1.CLIENT = "INT" Then
                sql_SEG2 = ", NVL(NVL(SOTTYPE1.SEG2_CODE,SOTSDIV1.SEG2_CODE),'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')"
            Else
                sql_SEG2 = ", NVL(NVL(SOTCHAN1.SEG2_CODE,NVL(SOTTYPE1.SEG2_CODE,SOTSDIV1.SEG2_CODE)),'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')"
            End If

            sqlG = "" _
                & " SELECT " & sqlXNO & " REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", '" & ROWs("SOTPARM1").Item("SO_PARM_ACCT_CODE_FREIGHT") & "' ACCT_CODE" & vbCrLf _
                & sql_SEG2 & " SEG2_CODE" & vbCrLf _
                & sql_SEG3 & " SEG3_CODE" & vbCrLf _
                & ", '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "' SEG4_CODE" & vbCrLf _
                & ", 'FRT-' || SOTINVH1.INV_TYPE DETL_CVX_REF_NO" & vbCrLf _
                & ", NULL DETL_CVX_TYPE, NULL DETLSQQ_CVX_NO" & vbCrLf _
                & ", SUM (-1 * NVL(SOTINVH1.INV_FREIGHT,0)) DIST_AMT " & vbCrLf _
                & " from " & SOTINVH1 & " SOTINVH1,SOTSDIV1,SOTTYPE1,ARTCUST1,SOTTCLS1,SOTEVNT1,SOTCHAN1" & vbCrLf _
                & " where SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & "   and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
                & "   and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
                & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
                & "   and SOTEVNT1.EVENT_CODE (+) = SOTINVH1.EVENT_CODE" & vbCrLf _
                & "   and SOTCHAN1.CHANNEL_CODE (+) = SOTTCLS1.CHANNEL_CODE" & vbCrLf _
                & " group by " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.INV_TYPE" & sql_SEG2 & sql_SEG3 & vbCrLf _
                & " order by " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, SOTINVH1.INV_TYPE" & sql_SEG2 & sql_SEG3
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            ' AR

            '& ", SOTSDIV1.SEG2_CODE" & vbCrLf _

            sqlG = "" _
                & " SELECT " & sqlXNO & " REGISTER_XNO, 'OPSJ' JOURNAL_TYPE, SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP" & vbCrLf _
                & ", ARTPOST1.ACCT_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "') SEG2_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "') SEG3_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "') SEG4_CODE" & vbCrLf _
                & ", DECODE(SOTINVH1.INV_TYPE,'I','AR-I','AR-C') DETL_CVX_REF_NO" & vbCrLf _
                & ", 'R' DETL_CVX_TYPE, SOTINVH1.POST_CODE DETL_CVX_NO" & vbCrLf _
                & ", SUM (SOTINVH1.INV_TOTAL_AMOUNT) DIST_AMT " & vbCrLf _
                & " FROM " & SOTINVH1 & " SOTINVH1,ARTPOST1,SOTSDIV1" & vbCrLf _
                & " where ARTPOST1.POST_CODE (+) = SOTINVH1.POST_CODE" & vbCrLf _
                & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = SOTINVH1.SALES_DIVISION_CODE" & vbCrLf _
                & " GROUP BY " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, ARTPOST1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTSDIV1.SEG2_CODE, SOTINVH1.POST_CODE" & vbCrLf _
                & ", NVL(ARTPOST1.SEG2_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')" & vbCrLf _
                & ", NVL(ARTPOST1.SEG3_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "')" & vbCrLf _
                & ", NVL(ARTPOST1.SEG4_CODE,'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "')" & vbCrLf _
                & " ORDER BY " & sqlXNO & ", SOTINVH1.ORDR_YYYYPP_UPDATED, ARTPOST1.ACCT_CODE, SOTINVH1.INV_TYPE, SOTSDIV1.SEG2_CODE, SOTINVH1.POST_CODE"
            ASCDATA1.ExecuteSQL("Insert into " & SOTINVHG & " " & sqlG)

            ASCDATA1.ExecuteSQL("Delete from " & SOTINVHG & " where DIST_AMT = 0")

            If ASCMAIN1.CLIENT = "AHA" Then ' EVENTUALLY OPEN UP TO INT
                ASCMAIN1.sql = "Select Distinct SOTINVHG.SEG3_CODE,SOTTCLS1.CHANNEL_CODE,SOTCHAN1.SEG2_CODE" & vbCrLf _
                    & " from " & SOTINVHG & " SOTINVHG,SOTTCLS1,SOTCHAN1" & vbCrLf _
                    & " where (SOTINVHG.SEG2_CODE is NULL or SOTINVHG.SEG2_CODE = '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "')" & vbCrLf _
                    & "   and SOTINVHG.SEG3_CODE <> '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'" & vbCrLf _
                    & "   and SOTTCLS1.TRADE_CLASS_CODE = SOTINVHG.SEG3_CODE" & vbCrLf _
                    & "   and SOTCHAN1.CHANNEL_CODE = SOTTCLS1.CHANNEL_CODE;"
                ASCMAIN1.sql = "" _
                    & "Begin" & vbCrLf _
                    & " Declare Cursor C1 is " & vbCrLf & ASCMAIN1.sql & vbCrLf _
                    & " Begin" & vbCrLf _
                    & "  For R1 in C1 Loop" & vbCrLf _
                    & "   Update " & SOTINVHG & " Set SEG2_CODE = R1.SEG2_CODE" & vbCrLf _
                    & "    where SEG2_CODE = '" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'" & vbCrLf _
                    & "      and SEG3_CODE = R1.SEG3_CODE;" & vbCrLf _
                    & "  End Loop;" & vbCrLf _
                    & " End;" & vbCrLf _
                    & "End;"
                ASCDATA1.ExecuteSQL()
            End If

            If MENU_ITEM_PP = "CJ" Then
                Prepare_GL_Interface("OPCJ")
            Else
                Prepare_GL_Interface("OPSJ")
            End If

            If RWU = "N" And ASCMAIN1.Running_in_VS And ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.CLIENT = "AHA" And Format(Now, "yyyyMMdd") = "20160613" Then
                RWU = "R"
            End If
        End If

        If RWU = "R" Then
            If ASCMAIN1.CLIENT = "AHA" Then ' EVENTUALLY OPEN UP TO INT
                For Each row As DataRow In ASCDATA1.SelectDistinct("SOTINVH1", New String() {"EVENT_CODE"}).Select("EVENT_CODE IS NOT NULL")
                    Dim EVENT_CODE As String = row.Item("EVENT_CODE")
                    Dim rowSOTEVNT1 As DataRow = dst.Tables("SOTEVNT1").Rows.Find(EVENT_CODE)
                    If rowSOTEVNT1.Item("EVENT_STATUS") & "" <> "A" Then
                        'MsgBox("Event Code " & EVENT_CODE & " is Inactive", MsgBoxStyle.OkOnly, "Update will be Denied")
                        RWU = "N"
                        xErrMsg = "Event Code " & EVENT_CODE & " is Inactive"
                        Exit For
                    End If
                Next
            End If
        End If

        Check_if_Empty("SOTINVH1")
    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("SUMMARY", "0")
        CR_params.Add("CJ", IIf(MENU_ITEM_PP = "CJ", "1", "0"))
        Generate_Report(RPT, , SUBT)

        'CR_params.Add("SUMMARY", "1")
        'SUBT = "Summary (Totals Only) for " & SUBT
        'Generate_Report(RPT, , SUBT)

        Print_GL()

        If MENU_ITEM_PP = "CJ" Then
            Prepare_Data_Extracts_OPCJ()
        Else
            If (chkSalesByState.Visible And chkSalesByState.Checked) Then
                Prepare_Data_Extracts_OPSJ(sqlSalesByState)
            Else
                Prepare_Data_Extracts_OPSJ()
            End If

        End If
    End Sub
    Sub Prepare_Data_Extract_Sales_By_State()

    End Sub
    Sub Prepare_Data_Extracts_OPSJ(Optional sqlSpecial As String = "")
        '            & ", SOTORDR5.CUST_STATE CUST_SHIP_TO_STATE" & vbCrLf _

        Dim sqlExtract As String = "Select SOTINVH1.*, ARTCUST1.TRADE_CLASS_CODE, SOTTCLS1.CHANNEL_CODE" & vbCrLf _
            & ", SOTORDR5.CUST_CITY CUST_SHIP_TO_CITY" & vbCrLf _
            & ", SOTORDR5.CUST_ZIP_CODE CUST_SHIP_TO_ZIP_CODE" _
            & " from " & SOTINVH1 & " SOTINVH1,ARTCUST1,SOTTCLS1,SOTORDR5 " & vbCrLf _
            & " where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   And SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   And SOTORDR5.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   And SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'"

        Dim dtSql As String = IIf(sqlSpecial <> "", sqlSpecial, sqlExtract)

        ASCMAIN1.sql = "Select Count (1) from (" & dtSql & ")"
        Dim recs As Integer = Val(ASCDATA1.GetDataValue)

        If ASCMAIN1.CLIENT = "INT" AndAlso recs > 100000 Then
            If MsgBox("Too many records to display, do you want to export as a CSV?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                Me.Cursor = Cursors.WaitCursor
                Dim FILENAME As String = ASCMAIN1.Export_To_CSV(dtSql, recs, "SalesJournal")
                Show_Document(FILENAME)
                Me.Cursor = Cursors.Default
            End If

        Else
            ASCMAIN1.sql = dtSql
            Dim DataTable As DataTable = ASCDATA1.GetDataTable
            Bind_Export(sqlSpecial, DataTable)
        End If

    End Sub
    Private Sub Bind_Export(sqlSpecial As String, dt As DataTable)

        grdASTEXPT1.DataSource = ASCDATA1.GetDataTable ' dst.Tables("SOTINVH1")
        Dim defaultSort As String = IIf(sqlSpecial <> "", "DestRegion,INV_NO", "INV_TYPE,INV_NO")
        Sort_grdColumns(grdASTEXPT1, defaultSort)
        grdASTEXPT1.Text = IIf(sqlSpecial <> "", "Sales By State Extract", "Sales Invoice & Memo Register")
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        If sqlSpecial = "" Then
            Set_DX_Column(grdASTEXPT1, "")
            Set_DX_Column(grdASTEXPT1, "INV_NO", "Invoice No", 100, , "Count")
            Set_DX_Column(grdASTEXPT1, "INV_TYPE", "Type", 40)
            Set_DX_Column(grdASTEXPT1, "CUST_CODE", "Customer", 100)
            Set_DX_Column(grdASTEXPT1, "CUST_NAME", "Customer Name", 100)
            Set_DX_Column(grdASTEXPT1, "CUST_STORE_NO", "Store", 60)
            Set_DX_Column(grdASTEXPT1, "ORDR_CUST_PO", "Customer PO", 100)
            Set_DX_Column(grdASTEXPT1, "EVENT_CODE", "Event", 50)
            Set_DX_Column(grdASTEXPT1, "ORDR_NO", "Order No", 100)
            Set_DX_Column(grdASTEXPT1, "REASON_CODE", "Reason", 50)
            Set_DX_Column(grdASTEXPT1, "INV_DATE", "Inv Date", 100, "MM/dd/yy")
            Set_DX_Column(grdASTEXPT1, "INV_DATE_SHIPPED", "Shipped", 100, "MM/dd/yy")
            Set_DX_Column(grdASTEXPT1, "CUST_SHIP_TO_CITY", "City", 100)
            Set_DX_Column(grdASTEXPT1, "CUST_SHIP_TO_STATE", "St", 50)
            Set_DX_Column(grdASTEXPT1, "CUST_SHIP_TO_ZIP_CODE", "Zip", 100)
            Set_DX_Column(grdASTEXPT1, "ORDR_TYPE_CODE", "OrdTyp", 50)
            Set_DX_Column(grdASTEXPT1, "TRADE_CLASS_CODE", "Trade Class", 50)
            Set_DX_Column(grdASTEXPT1, "CHANNEL_CODE", "Channel", 50)
            Set_DX_Column(grdASTEXPT1, "WHSE_CODE", "Whse", 60)
            Set_DX_Column(grdASTEXPT1, "POST_CODE", "Posting", 60)
            Set_DX_Column(grdASTEXPT1, "TERM_CODE", "Terms", 60)
            Set_DX_Column(grdASTEXPT1, "SREP_CODE", "SRep", 60)
            Set_DX_Column(grdASTEXPT1, "CUST_BILL_TO_CUST", "Bill-To", 60)
            Set_DX_Column(grdASTEXPT1, "ORDR_YYYYPP_UPDATED", "YP", 60)
            Set_DX_Column(grdASTEXPT1, "REGISTER_XNO", "XNo", 80)
            Set_DX_Column(grdASTEXPT1, "REGISTER_DATE", "Posted", 100, "MM/dd/yy")


            Set_DX_Column(grdASTEXPT1, "OPS_YYYYWW", "YW", 80)
            Set_DX_Column(grdASTEXPT1, "ORDR_DEPT", "Dept Code", 80)
            Set_DX_Column(grdASTEXPT1, "ORDR_NO_WEB", "Web Order", 80)
            Set_DX_Column(grdASTEXPT1, "SALES_DIVISION_CODE", "Division", 80)
            Set_DX_Column(grdASTEXPT1, "PARTNER_ORDR_NO", "Ptnr Order", 80)

            Set_DX_Column(grdASTEXPT1, "CURR_CODE", "Curr", 60)
            Set_DX_Column(grdASTEXPT1, "CURR_EXCH_RATE", "Exch", 60, "#,##0.0000")

            Set_DX_Column(grdASTEXPT1, "INV_BOL_NO", "BOL No", 80)
            Set_DX_Column(grdASTEXPT1, "INV_PRO_NO", "Pro No", 80)
            Set_DX_Column(grdASTEXPT1, "SHIP_REF", "Ship Ref", 80)
            Set_DX_Column(grdASTEXPT1, "SHIP_VIA_CODE", "SVia Code", 80)
            Set_DX_Column(grdASTEXPT1, "SHIP_VIA_DESC", "Ship Via", 80)
            Set_DX_Column(grdASTEXPT1, "INV_NO_CONS", "Cons Inv No", 80)
            Set_DX_Column(grdASTEXPT1, "SHIP_BOL_NO", "Shipment No", 80)
            Set_DX_Column(grdASTEXPT1, "PICK_NO", "Pick No", 80)
            Set_DX_Column(grdASTEXPT1, "INV_COMMENT", "Comment", 80)

            Set_DX_Column(grdASTEXPT1, "PICK_NO", "Pick No", 80)
            Set_DX_Column(grdASTEXPT1, "INV_CARTONS", "Cartons", 80, "#,##0", "Sum")
            Set_DX_Column(grdASTEXPT1, "INV_WEIGHT", "Weight", 80, "#,##0.00", "Sum")

            Set_DX_Column(grdASTEXPT1, "INV_SALES", "Sales", 120, "#,##0.00", "Sum")
            Set_DX_Column(grdASTEXPT1, "INV_FREIGHT", "Freight", 80, "#,##0.00", "Sum")
            Set_DX_Column(grdASTEXPT1, "INV_MISC_CHG", "Misc Chgs", 80, "#,##0.00", "Sum")
            Set_DX_Column(grdASTEXPT1, "MISC_CHG_CODE", "Misc Code", 80)
            Set_DX_Column(grdASTEXPT1, "INV_STAX", "SlsTax", 80, "#,##0.00", "Sum")
            Set_DX_Column(grdASTEXPT1, "STAX_CODE", "STax Code", 80)
            Set_DX_Column(grdASTEXPT1, "STAX_RATE", "Rate", 60, "#,##0.0000")
            Set_DX_Column(grdASTEXPT1, "INV_TOTAL_AMOUNT", "Freight", 80, "#,##0.00", "Sum")
            Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collection", 80)

            If optRANGE.Value = "A" Then
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("SHIP_REF").Header.VisiblePosition = 0
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("INV_BOL_NO").Header.VisiblePosition = 1
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("CUST_CODE").Header.VisiblePosition = 2
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("INV_DATE_SHIPPED").Header.VisiblePosition = 3
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("SHIP_VIA_CODE").Header.VisiblePosition = 4
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("WHSE_CODE").Header.VisiblePosition = 5

                grdASTEXPT1.DisplayLayout.Bands(0).Columns("SHIP_REF").CellAppearance.BackColor = Drawing.Color.LightBlue
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("INV_BOL_NO").CellAppearance.BackColor = Drawing.Color.LightBlue
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("CUST_CODE").CellAppearance.BackColor = Drawing.Color.LightBlue
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("INV_DATE_SHIPPED").CellAppearance.BackColor = Drawing.Color.LightBlue
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("SHIP_VIA_CODE").CellAppearance.BackColor = Drawing.Color.LightBlue
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("WHSE_CODE").CellAppearance.BackColor = Drawing.Color.LightBlue
            Else
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("SHIP_REF").CellAppearance.BackColor = Nothing
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("INV_BOL_NO").CellAppearance.BackColor = Nothing
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("CUST_CODE").CellAppearance.BackColor = Nothing
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("INV_DATE_SHIPPED").CellAppearance.BackColor = Nothing
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("SHIP_VIA_CODE").CellAppearance.BackColor = Nothing
                grdASTEXPT1.DisplayLayout.Bands(0).Columns("WHSE_CODE").CellAppearance.BackColor = Nothing
            End If
        Else
            Set_DX_Column(grdASTEXPT1, "INV_NO", "Invoice No", 100, , "Count")
        End If

        Dim btn As New Infragistics.Win.Misc.UltraButton
        AddHandler btn.Click, AddressOf btn_Click
        grdASTEXPT1.Controls.Add(btn)
        btn.Visible = ASCMAIN1.Running_in_VS
        btn.Left = 0
        btn.Top = 0
        btn.Text = "XLS"
    End Sub
    Private Sub btn_Click(sender As System.Object, e As System.EventArgs)

        WorkbookView1.GetLock()

        Dim datatable As DataTable = DirectCast(grdASTEXPT1.DataSource, DataTable)
        Load_DataTable_into_SGXLS(1, 1, datatable, WorkbookView1.ActiveWorksheet, Nothing, Nothing, "INV_TYPE,INV_NO", "")

        Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("SORSLSJ1.XLSX_NO") & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        WorkbookView1.ReleaseLock()

    End Sub

    Sub Prepare_Data_Extracts_OPCJ()

        Dim DT As DataTable = Create_Data_Extract() ' dst.Tables("ARTATBR1")
        grdASTEXPT1.DataSource = DT
        Sort_grdColumns(grdASTEXPT1, "INV_TYPE,INV_NO")
        grdASTEXPT1.Text = "Sales and CGS Detail"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Set_DX_Column(grdASTEXPT1, "INV_TYPE", "Type", 40)
        Set_DX_Column(grdASTEXPT1, "INV_NO", "Invoice No", 100, , "Count")
        Set_DX_Column(grdASTEXPT1, "CUST_CODE", "Customer", 100)
        Set_DX_Column(grdASTEXPT1, "CUST_STORE_NO", "Store", 60)
        Set_DX_Column(grdASTEXPT1, "ORDR_CUST_PO", "Customer PO", 100)
        Set_DX_Column(grdASTEXPT1, "EVENT_CODE", "Event", 50)
        Set_DX_Column(grdASTEXPT1, "REASON_CODE", "Reason", 50)
        Set_DX_Column(grdASTEXPT1, "INV_DATE", "Inv Date", 100, "MM/dd/yy")
        Set_DX_Column(grdASTEXPT1, "CUST_SHIP_TO_STATE", "St", 50)
        Set_DX_Column(grdASTEXPT1, "CUST_SHIP_TO_CITY", "City", 100)
        Set_DX_Column(grdASTEXPT1, "ORDR_TYPE_CODE", "OrdTyp", 50)
        Set_DX_Column(grdASTEXPT1, "TRADE_CLASS_CODE", "Trade Class", 50)
        Set_DX_Column(grdASTEXPT1, "CHANNEL_CODE", "Channel", 50)
        Set_DX_Column(grdASTEXPT1, "BRAND_CODE", "Brand", 50)
        Set_DX_Column(grdASTEXPT1, "HC_CODE", "HC", 50)
        Set_DX_Column(grdASTEXPT1, "ITEM_CODE", "Item Code", 80)
        Set_DX_Column(grdASTEXPT1, "ITEM_DESC", "Item Description", 150)
        Set_DX_Column(grdASTEXPT1, "PROD_CODE", "Prod", 80)
        Set_DX_Column(grdASTEXPT1, "COLLECTION_CODE", "Collection", 80)
        Set_DX_Column(grdASTEXPT1, "COST_CATGY_CODE", "CCat", 40)
        Set_DX_Column(grdASTEXPT1, "ITEM_SNU_CODE", "SNU", 40)
        Set_DX_Column(grdASTEXPT1, "WHSE_CODE", "Whse", 60)

        Set_DX_Column(grdASTEXPT1, "ITEM_RETAIL_PRICE", "Retail Price", 80, "#,##0.00")
        Set_DX_Column(grdASTEXPT1, "ORDR_UNIT_PRICE", "Unit Price", 80, "#,##0.00")
        Set_DX_Column(grdASTEXPT1, "ITEM_UNIT_COST", "Unit Cost", 80, "#,##0.0000")
        Set_DX_Column(grdASTEXPT1, "ORDR_QTY_SHIP", "Qty", 80, "#,##0", "Sum")

        Set_DX_Column(grdASTEXPT1, "SLS", "SLS", 120, "#,##0.00", "Sum")
        Set_DX_Column(grdASTEXPT1, "CGS", "CGS", 120, "#,##0.00", "Sum")

        Set_DX_Column(grdASTEXPT1, "ACCT_CODE_SLS", "Acct SLS", 80)
        Set_DX_Column(grdASTEXPT1, "ACCT_CODE_CGS", "Acct CGS", 80)
        Set_DX_Column(grdASTEXPT1, "SEG2_CODE", "Dept", 80)
        Set_DX_Column(grdASTEXPT1, "SEG3_CODE", "TrCls", 80)
        Set_DX_Column(grdASTEXPT1, "SEG4_CODE", "Collct", 80)

        If ASCMAIN1.CLIENT = "INT" Then
            WorkbookView1.GetLock()
            Load_DataTable_into_SGXLS(1, 1, DT, WorkbookView1.ActiveWorksheet, Nothing, Nothing, "INV_TYPE,INV_NO", "")
            WorkbookView1.ReleaseLock()

            WorkbookView1.Visible = True
            btnExcel.Visible = True
            btnExcel.BringToFront()
        End If

    End Sub

    Function Create_Data_Extract() As DataTable
        ASCMAIN1.sql = "Select SOTINVH1.INV_TYPE" & vbCrLf _
            & ",SOTINVH1.INV_NO" & vbCrLf _
            & ",SOTINVH1.CUST_CODE" & vbCrLf _
            & ",SOTINVH1.CUST_STORE_NO" & vbCrLf _
            & ",SOTINVH1.ORDR_CUST_PO" & vbCrLf _
            & ",SOTINVH1.EVENT_CODE" & vbCrLf _
            & ",SOTINVH1.REASON_CODE" & vbCrLf _
            & ",SOTINVH1.INV_DATE" & vbCrLf _
            & ",SOTINVH1.CUST_SHIP_TO_STATE" & vbCrLf _
            & ",SOTORDR5.CUST_CITY CUST_SHIP_TO_CITY" & vbCrLf _
            & ",SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & ",ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & ",SOTTCLS1.CHANNEL_CODE" & vbCrLf _
            & ",ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ",ICTCOLL1.HC_CODE" & vbCrLf _
            & ",SOTINVH1.ORDR_TYPE_CODE" & vbCrLf _
            & ",SOTINVH2.ITEM_CODE" & vbCrLf _
            & ",ICTITEM1.ITEM_DESC" & vbCrLf _
            & ",ICTITEM1.PROD_CODE" & vbCrLf _
            & ",ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ",ICTITEM1.COST_CATGY_CODE" & vbCrLf _
            & ",ICTITEM1.ITEM_SNU_CODE" & vbCrLf _
            & ",SOTINVH2.WHSE_CODE" & vbCrLf _
            & ",SOTINVH2.ITEM_RETAIL_PRICE" & vbCrLf _
            & ",SOTINVH2.ORDR_UNIT_PRICE" & vbCrLf _
            & ",SOTINVH2.ORDR_QTY_SHIP" & vbCrLf _
            & ",SOTINVH2.ITEM_UNIT_COST" & vbCrLf _
            & ",SOTINVH2.ORDR_QTY_SHIP * SOTINVH2.ORDR_UNIT_PRICE SLS" & vbCrLf _
            & ", CASE WHEN SOTINVH2.INV_TYPE = 'C' AND ICTCOLL1.BRAND_CODE in ('KSP','LCS') AND ICTITEM1.ITEM_DESC LIKE 'XXX %' THEN 0 ELSE " & vbCrLf _
            & "       NVL(SOTINVH2.ORDR_QTY_SHIP,0) * DECODE(SOTINVH2.WHSE_CODE,NULL,0,NVL(SOTINVH2.ITEM_UNIT_COST,0)) END CGS " & vbCrLf _
            & ",SOTINVHT.ACCT_CODE_SLS" & vbCrLf _
            & ",SOTINVHT.ACCT_CODE_CGS" & vbCrLf _
            & ",SOTINVHT.SEG2_CODE" & vbCrLf _
            & ",SOTINVHT.SEG3_CODE" & vbCrLf _
            & ",SOTINVHT.SEG4_CODE" & vbCrLf _
            & " from SOTINVH2," & SOTINVH1 & " SOTINVH1,SOTORDR5,ICTITEM1,ICTCOLL1,ARTCUST1,SOTTCLS1," & SOTINVHT & " SOTINVHT" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVHT.INV_TYPE = SOTINVH2.INV_TYPE" & vbCrLf _
            & "   and SOTINVHT.INV_NO = SOTINVH2.INV_NO" & vbCrLf _
            & "   and SOTINVHT.INV_LNO = SOTINVH2.INV_LNO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE (+) = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTITEM1.ITEM_CODE = SOTINVH2.ITEM_CODE" & vbCrLf _
            & "   and SOTORDR5.ORDR_NO (+) = SOTINVH1.ORDR_NO" & vbCrLf _
            & "   and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'"
        Dim DataTable As DataTable = ASCDATA1.GetDataTable
        Return DataTable
    End Function
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
            ElseIf Absx1.optFor("RANGE").Value = "A" Then
                txtAUDIT_NO.Text = txtAUDIT_NO.Text.Trim
                If txtAUDIT_NO.TextLength > 0 Then
                    Dim rowWHTAUDT1 As DataRow = LookUp("WHTAUDT1", txtAUDIT_NO.Text)
                    If rowWHTAUDT1 Is Nothing Then
                        EMsg &= vbCr & "Invalid Audit No."
                    End If
                Else
                    EMsg &= vbCr & "Missing Audit No."
                End If
            ElseIf Absx1.optFor("RANGE").Value = "N" Then
                Dim dte() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
                If Format(dteINV_DATE_CUTOFF.Value, "yyyyMMdd") < Format(dte(1), "yyyyMMdd") _
                Or Format(dteINV_DATE_CUTOFF.Value, "yyyyMMdd") > Format(dte(dte.Length - 1), "yyyyMMdd") Then
                    EMsg &= vbCr & "Cut-Off Date must be between " & Format(dte(1), "MM/dd/yyyy") & " and " & Format(dte(dte.Length - 1), "MM/dd/yyyy") & " - Current Period is " & ASCMAIN1.CYP
                End If
            End If

            If MENU_ITEM_PP = "CJ" Then
                If ASCMAIN1.EOM = "1" Then
                    If tblASTDSQLA.Select("CODE_VALUES <> ''").Length <> 0 Then
                        EMsg &= vbCr & "Filters are not permitted on Period End Report"
                    End If
                End If
            End If

        End If
    End Sub

    Private Sub optRANGE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRANGE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        grpDATE_RANGE.Visible = (optRANGE.Value = "D")
        grpPERIOD_RANGE.Visible = (optRANGE.Value = "P")
        chkSalesByState.Visible = (ASCMAIN1.CLIENT = "AHA" And optRANGE.Value = "P")
        grpAudit.Visible = (optRANGE.Value = "A" AndAlso ASCMAIN1.CLIENT = "INT")
        grpInclude.Visible = (optRANGE.Value <> "A")

        If optRANGE.Value = "P" Then
            Absx1.dteFor("DTE0").Value = Null
            Absx1.dteFor("DTE1").Value = Null
            Set_cmbYP("RYP0", ASCMAIN1.CYP, -24, 0, 0)
            Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
            chkSalesByState.Left = grpPERIOD_RANGE.Left + grpPERIOD_RANGE.Width + 6
        ElseIf optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If

        chkINV_DATE_CUTOFF.Visible = (optRANGE.Value = "N")
        dteINV_DATE_CUTOFF.Visible = (optRANGE.Value = "N")

        If (optRANGE.Value <> "A") Then
            txtAUDIT_NO.Clear()
            txtAUDIT_DATE.Clear()
        End If

    End Sub

    Overrides Sub Update_Record()

        If MENU_ITEM_PP = "CJ" Then
            ' no update to mark the rows? they are retreived by period for the CJ report, 
            '  so that would have to be modified as well if we were to update a flag here
        Else
            Dim sql As String = "Update SOTINVH1 " _
                & " Set REGISTER_IND = :PARM1, REGISTER_XNO = :PARM2, REGISTER_DATE = :PARM3" _
                & " where INV_NO in (Select INV_NO from " & SOTINVH1 & " )"
            ASCDATA1.ExecuteSQL(sql, "VVD", New Object() {"1", MyBase.XNO, REGISTER_DATE})
        End If

        If MENU_ITEM_PP = "CJ" Then
            ASCDATA1.ExecuteSQL("Insert into SOTINVHS Select * from " & SOTINVHS)
            ASCDATA1.ExecuteSQL("Insert into SOTINVHG Select * from " & SOTINVHG)
            ASCDATA1.ExecuteSQL("Insert into SOTINVHT Select * from " & SOTINVHT)
            ' SOTINVHT is totally redundant with data that can be found in SOTINVH2 & joins, and is to be used as a safety net and not for reports
            ' ditto for SOTINVHS, SOTINVHG

            ASCMAIN1.sql = "" _
                 & "Begin" & vbCrLf _
                 & " Declare Cursor C1 is Select * from " & SOTINVHT & " SOTINVHT;" & vbCrLf _
                 & " Begin" & vbCrLf _
                 & "  For R1 in C1 Loop" & vbCrLf _
                 & "   Update SOTINVH2 Set ACCT_CODE_SLS = R1.ACCT_CODE_SLS, ACCT_CODE_CGS = R1.ACCT_CODE_CGS" & vbCrLf _
                 & "    , SEG2_CODE = R1.SEG2_CODE, SEG3_CODE = R1.SEG3_CODE, SEG4_CODE = R1.SEG4_CODE" & vbCrLf _
                 & "    where INV_TYPE = R1.INV_TYPE and INV_NO = R1.INV_NO and INV_LNO = R1.INV_LNO;" & vbCrLf _
                 & "  End Loop;" & vbCrLf _
                 & " End;" & vbCrLf _
                 & "End;"
            ASCDATA1.ExecuteSQL()
        End If

        GL_Update()

        If ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA" Then
            If MENU_ITEM_PP = "CJ" Then
                ' no extract on CGS journal
            Else
                Sales_Extract_Files()
            End If
        End If
    End Sub

    Sub Sales_Extract_Files()
        TAC.SACMAIN1.Create_Sales_Extract_Tables(Me, False, SATSLSXC, SATSLSXI, SATSLSXS, "X", XNO)
        TAC.SACMAIN1.ftp_BI_Files(Me)
    End Sub
    Function Prepare_GL_Interface(ByVal JOURNAL_TYPE As String) As String

        ' Prepare GL Interface File

        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Decimal
        Dim DETL_CTL_DATE As Date = DateValue(Format(Now + ASCMAIN1.NowTSD, "MM/dd/yyyy"))
        If chkINV_DATE_CUTOFF.Checked Then
            DETL_CTL_DATE = dteINV_DATE_CUTOFF.Value
        End If

        'Dim REGISTER_XNO_SOTINVHG As String = ""

        ASCMAIN1.sql = "Select * from " & SOTINVHG & " where JOURNAL_TYPE = '" & JOURNAL_TYPE & "'"
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows ' .Select("", "REGISTER_XNO,OPS_YYYYPP,ACCT_CODE,SEG2_CODE,SEG3_CODE,SEG4_CODE")

            'If RWU = "N" And ASCMAIN1.USER_ID = "wjz" Then
            '    If row.Item("REGISTER_XNO") & "" = XNO Then
            '    Else
            '        If row.Item("REGISTER_XNO") & "" = REGISTER_XNO_SOTINVHG Then
            '        Else
            '            REGISTER_XNO_SOTINVHG = row.Item("REGISTER_XNO") & ""
            '            ASCMAIN1.sql = "Select JOURNAL_NO from "
            '        End If
            '    End If
            'End If

            DETL_POSTING_AMT = Val(row.Item("DIST_AMT") & "")

            'If ASCMAIN1.Running_in_VS And JOURNAL_NO = "020446" And (ASCMAIN1.DBS_SERVER = "AHA" Or ASCMAIN1.DBS_COMPANY = "AHA") Then
            '    If JOURNAL_LNO = 0 Then Stop
            '    DETL_POSTING_AMT = -1 * DETL_POSTING_AMT
            'End If

            Dim DETL_CVX_NO As String = row.Item("DETL_CVX_NO") & ""
            Dim DETL_CVX_REF_NO As String = row.Item("DETL_CVX_REF_NO") & ""
            Dim DETL_CVX_TYPE As String = row.Item("DETL_CVX_TYPE") & ""

            Dim rowGLTINTF1 As DataRow = ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").NewRow
            rowGLTINTF1("OPS_YYYYPP") = row("OPS_YYYYPP")
            rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
            JOURNAL_LNO += 1
            rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
            rowGLTINTF1("ACCT_CODE") = row("ACCT_CODE")
            rowGLTINTF1("SEG2_CODE") = row("SEG2_CODE")
            rowGLTINTF1("SEG3_CODE") = row("SEG3_CODE")
            rowGLTINTF1("SEG4_CODE") = row("SEG4_CODE")
            rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
            rowGLTINTF1("DETL_POSTING_AMT") = Round(DETL_POSTING_AMT, 2)
            If RWU = "N" And ASCMAIN1.USER_ID = "wjz" And ASCMAIN1.Running_in_VS Then
                rowGLTINTF1("DETL_EXE_NO") = row("REGISTER_XNO")
            Else
                rowGLTINTF1("DETL_EXE_NO") = ASCMAIN1.ActiveForm.XNO
            End If
            rowGLTINTF1("DETL_CTL_NO") = DBNull.Value
            rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
            rowGLTINTF1("DETL_CVX_NO") = DETL_CVX_NO
            rowGLTINTF1("DETL_CVX_REF_DATE") = REGISTER_DATE
            rowGLTINTF1("DETL_CVX_REF_NO") = DETL_CVX_REF_NO
            rowGLTINTF1("DETL_DESC") = DBNull.Value
            rowGLTINTF1("DETL_CVX_TYPE") = DETL_CVX_TYPE
            rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
            ASCMAIN1.ActiveForm.dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        Next

        Return JOURNAL_NO

    End Function

    Private Sub chkINV_DATE_CUTOFF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkINV_DATE_CUTOFF.CheckedChanged
        dteINV_DATE_CUTOFF.Visible = chkINV_DATE_CUTOFF.Checked
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Exporting to Excel")

        WorkbookView1.GetLock()
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & ASCMAIN1.Next_Control_No("RSFSSPL1.XLSX_NO") & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        WorkbookView1.ReleaseLock()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("Now Exporting to Excel")
    End Sub

    Private Sub txtAUDIT_NO_ValueChanged(sender As Object, e As EventArgs) Handles txtAUDIT_NO.ValueChanged

        Dim AUDIT_NO As String = txtAUDIT_NO.Text.Trim
        If AUDIT_NO.Length = 0 Then
            txtAUDIT_DATE.Clear()
        Else
            Dim rowWHTAUDT1 As DataRow = LookUp("WHTAUDT1", AUDIT_NO)
            If rowWHTAUDT1 Is Nothing Then
                txtAUDIT_DATE.Clear()
            Else
                txtAUDIT_DATE.Text = CDate(rowWHTAUDT1.Item("SHIP_DATE_FROM") & String.Empty).ToShortDateString & " - " & CDate(rowWHTAUDT1.Item("SHIP_DATE_TO") & String.Empty).ToShortDateString
            End If
        End If
    End Sub
End Class