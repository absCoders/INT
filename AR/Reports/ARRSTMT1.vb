Imports System.Xml

Public Class ARRSTMT1

    Dim ARTSTMTX As String

    Dim SOTINVH1_ECP As String
    Dim SOTINVH1_FRT As String
    Dim SOTINVH1_FSC As String
    Dim SOTINVH1_FIN As String

    Dim ARTCUST1_B2C As String
    Dim ARTSTMT1 As String
    Dim ARTSTMT2 As String
    Dim ARTOPEN1 As String
    Dim ARTCUST1_has_sales As String
    Dim ARTCUST3 As String

    Dim ARTSTMTO As String
    Dim ARTSTMTB As String

    Dim RYP_DATE As Date
    Dim RYP_DATE_DUE As Date
    Dim LYP As String = ""
    Dim LYP_DATE As Date

    Dim skip_ABS_Calcs As Boolean = False

    Dim XMLHeadings As New Dictionary(Of String, String())
    Dim XMLHeadingTypes As New Dictionary(Of String, String)
    Dim WithEvents Ftp1 As New nsoftware.IPWorks.Ftp

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("ARTPARM1")

        ASCMAIN1.sql = "Select GLTPARM2.* " _
        & " from GLTPARM2 " _
        & " where OPS_YYYYPP = " _
        & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
        & "  where GLTPARM2.PRD_END_DATE >= :PARM1)"
        Create_TDA(dst.Tables.Add, "GLTPARM2", "**", 0, False, "V", 1)

        Create_Lookup("GLTPARM2")
        Create_Lookup("ARTSTMT0")

        Set_cmbYP("RYP", ASCMAIN1.CYP, 0, 0, 0)

        txtSTMT_MESSAGE.Text = ROWs("ARTPARM1").Item("AR_PARM_STMT_MESSAGE") & ""
        txtSTMT_MESSAGE_ECP.Text = ROWs("ARTPARM1").Item("AR_PARM_STMT_MESSAGE_ECP") & ""

        Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME
            Case "STMT_INSERT_CODE", "STMT_INSERT_CODE2", "STMT_INSERT_CODE3"
                sql_where = "STMT_INSERT_STATUS = 'P'"
            Case "STMT_MESSAGE_CODE"
                sql_where = "STMT_MESSAGE_STATUS = 'P'"
        End Select

    End Sub
    Protected Overrides Sub Build_Workfile()

        RWU = "R"

        Call ASCMAIN1.Progress("Now Loading Work Space")

        With dst
            Create_TDA(.Tables.Add, "ARTOPEN1", "*")
            Create_TDA(.Tables.Add, "GLTINTF1", "*")
            Create_TDA(.Tables.Add, "ARTSTMT0", "*")

            .Tables.Add(ASCDATA1.GetDataTable("*", "SOTSREP1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "SOTTYPE1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "SOTMISC1"))
            .Tables.Add(ASCDATA1.GetDataTable("*", "ARTPOST1"))

            ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME" _
            & ", ARTCUST1.CUST_ADDR1, ARTCUST1.CUST_ADDR2, ARTCUST1.CUST_ADDR3 " _
            & ", ARTCUST1.CUST_CITY, ARTCUST1.CUST_STATE, ARTCUST1.CUST_ZIP_CODE " _
            & ", ARTCUST1.CUST_COUNTRY, ARTCUST1.CUST_PHONE, ARTCUST1.CUST_EXT " _
            & ", ARTCUST1.CUST_FAX, ARTCUST1.CUST_CONTACT, ARTCUST1.SREP_CODE " _
            & " from ARTCUST1"
            .Tables.Add(ASCDATA1.GetDataTable(, "ARTCUST1", 1))
        End With

        RYP_DATE = LookUp("GLTPARM2", RYP).Item("PRD_END_DATE")
        Dim RYM As String = ASCMAIN1.Get_YYYYMM(RYP)
        'RYP_DATE_DUE = CDate(Format(CDate(Format(RYP_DATE, "MM/01/yyyy")).AddMonths(1), "MM/10/yyyy"))
        RYP_DATE_DUE = CDate(Format(CDate(Mid(RYM, 5, 2) & "/01/" & Mid(RYM, 1, 4)).AddMonths(1), "MM/10/yyyy"))
        LYP = ASCMAIN1.Period_Calc(RYP, -1)
        LYP_DATE = LookUp("GLTPARM2", LYP).Item("PRD_END_DATE")

        ASCMAIN1.sql = "Select * from ARTCUST3 where (CUST_CODE, FRT_CONT_NO) in (Select CUST_CODE, MAX(FRT_CONT_NO) FRT_CONT_NO from ARTCUST3 group by CUST_CODE)"
        ARTCUST3 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST3 & " Add Primary Key (CUST_CODE)")

        ASCMAIN1.sql = "Select * from ARTOPEN1 where ROWNUM < 1"
        ARTOPEN1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTOPEN1 & " Add Primary Key (CUST_CODE, INV_TYPE, INV_NUM)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ARTOPEN1 & "_1 on " & ARTOPEN1 & " (TERM_CODE)")

        Call Generate_Statements()

        Call ASCMAIN1.Progress("Now Loading Data into Work Area")

        With dst
            ASCMAIN1.sql = "SELECT ARTSTMT1.*" _
            & ", DECODE(X.CUST_CODE,NULL,'0','1') DEL" _
            & ", DECODE(Y.CUST_CODE,NULL,'0','1') LOC " _
            & " from " & ARTSTMT1 & " ARTSTMT1" _
            & ", (SELECT DISTINCT ARTSTMT2.CUST_CODE from " & ARTSTMT2 & " ARTSTMT2" _
            & " where DIVISION_CODE = 'DEL') X" _
            & ", (SELECT CUST_CODE from (Select Distinct CUST_CODE, CUST_SHIP_TO_NO from " & ARTSTMT2 & " ARTSTMT2) group by CUST_CODE having COUNT (*) > 1) Y" _
            & " where X.CUST_CODE (+) = ARTSTMT1.CUST_CODE" _
            & "   and Y.CUST_CODE (+) = ARTSTMT1.CUST_CODE"
            .Tables.Add(ASCDATA1.GetDataTable("**", "ARTSTMT1"))
            dst.Tables("ARTSTMT1").Columns("STMT_SCAN_LINE").ReadOnly = False
            dst.Tables("ARTSTMT1").Columns("MAIL_STATEMENT").ReadOnly = False

            ASCMAIN1.sql = "Select ARTSTMT2.*,CASE WHEN DIVISION_CODE = 'DEL' THEN '1' ELSE '0' END DEL from " & ARTSTMT2 & " ARTSTMT2" ' & " where CUST_CODE LIKE '0112%'"
            .Tables.Add(ASCDATA1.GetDataTable("**", "ARTSTMT2"))
        End With

        Call ASCMAIN1.Progress("Now Calculating Check Digit")

        For Each rowARTSTMT1 As DataRow In dst.Tables("ARTSTMT1").Rows
            Dim STMT_NO As String = rowARTSTMT1.Item("STMT_NO")
            Dim STMT_DATE As String = RYP_DATE.ToString("yyMMdd")
            Dim CUST_CODE As String = rowARTSTMT1.Item("CUST_CODE")

            Dim TOTAL_DUE As Double = Val(rowARTSTMT1.Item("TOTAL_DUE") & "")
            If TOTAL_DUE < 0 Then
                TOTAL_DUE = 0
            End If

            'STMT_DATE = "071130"
            'CUST_CODE = "012345"
            'TOTAL_DUE = 2487.72
            'STMT_NO = "1234567890"
            ' 0711300123455123456789030002487728

            Dim STMT_SCAN_LINE As String = STMT_DATE _
            & CUST_CODE _
            & ASCMAIN1.CheckDigit(CUST_CODE) _
            & STMT_NO _
            & ASCMAIN1.CheckDigit(STMT_NO) _
            & Format(TOTAL_DUE * 100, "000000000")

            STMT_SCAN_LINE &= ASCMAIN1.CheckDigit(STMT_SCAN_LINE)
            rowARTSTMT1.Item("STMT_SCAN_LINE") = STMT_SCAN_LINE

            Dim MAIL_STATEMENT As String = rowARTSTMT1.Item("MAIL_STATEMENT") & ""
            If MAIL_STATEMENT <> "P" Then
                MAIL_STATEMENT = IIf(rowARTSTMT1.Item("MAIL_STATEMENT") & "" = "N" Or rowARTSTMT1.Item("CUST_B2C_IND") & "" = "1" Or (Val(rowARTSTMT1.Item("TOTAL_DUE") & "") <= 2 And Val(rowARTSTMT1.Item("TYP_ECP") & "") = 0), "N", "Y")
            End If

            rowARTSTMT1.Item("MAIL_STATEMENT") = MAIL_STATEMENT
        Next

        Me.XML_Generate_Statement()

        If ASCMAIN1.Running_in_VS Then
            'Stop
            'ASCDATA1.ExecuteSQL("Rename " & SOTINVH1_ECP & " to WJZ" & Mid(XNO, 9, 2) & "_ECP")
            'ASCDATA1.ExecuteSQL("Rename " & SOTINVH1_FIN & " to WJZ" & Mid(XNO, 9, 2) & "_FIN")
            'ASCDATA1.ExecuteSQL("Rename " & SOTINVH1_FSC & " to WJZ" & Mid(XNO, 9, 2) & "_FSC")
            'ASCDATA1.ExecuteSQL("Rename " & SOTINVH1_FRT & " to WJZ" & Mid(XNO, 9, 2) & "_FRT")
            'ASCDATA1.ExecuteSQL("Rename " & ARTSTMT1 & " to WJZ" & Mid(XNO, 9, 2) & "_STMT1")
            'ASCDATA1.ExecuteSQL("Rename " & ARTSTMT2 & " to WJZ" & Mid(XNO, 9, 2) & "_STMT2")
            'ASCDATA1.ExecuteSQL("Rename " & ARTSTMTX & " to WJZ" & Mid(XNO, 9, 2) & "_STMTX")
            'ASCDATA1.ExecuteSQL("Rename " & ARTCUST1_B2C & " to WJZ" & Mid(XNO, 9, 2) & "_B2C")
            'ASCDATA1.ExecuteSQL("Rename " & ARTSTMTO & " to WJZ" & Mid(XNO, 9, 2) & "_STMTO")
            'ASCDATA1.ExecuteSQL("Rename " & ARTSTMTB & " to WJZ" & Mid(XNO, 9, 2) & "_STMTB")

            'Stop ' USE THIS NEXT SECTION TO SEND A PDF ONLY FILE TO OSG WITHOUT HAVING TO MESS WITH CODE IN UPDATE
            'Call ftp_File()
            ''System.Diagnostics.Process.Start("ftp", " -s:" & ASCMAIN1.Folders("Temp") & "osg.scr")
            'Stop
        End If


    End Sub

    Public Overrides Sub Print_Report()

        RPT = "ARRSTMT1"
        CR_params.Add("STMT_DATE", Format(RYP_DATE, "MM/dd/yy"))
        CR_params.Add("AGE_CATGY_1", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") & "")
        CR_params.Add("AGE_CATGY_2", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") & "")
        CR_params.Add("AGE_CATGY_3", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") & "")
        CR_params.Add("AGE_CATGY_4", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4") & "")
        Generate_Report(RPT, "Statement Control Sheet")

        'RPT = "ARRSTMT2"
        'CR_params.Add("STMT_DATE", Format(RYP_DATE, "MM/dd/yy"))
        'CR_params.Add("AGE_CATGY_1", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") & "")
        'CR_params.Add("AGE_CATGY_2", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") & "")
        'CR_params.Add("AGE_CATGY_3", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") & "")
        'CR_params.Add("AGE_CATGY_4", ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4") & "")
        'CR_params.Add("PYMT_MSG", ROWs("ARTPARM1").Item("AR_PARM_STMT_PYMT_MSG") & "")
        'CR_params.Add("REMIT_TO", ROWs("ARTPARM1").Item("AR_PARM_STMT_REMIT_TO") & "")
        'CR_params.Add("RETURN_TO", ROWs("ARTPARM1").Item("AR_PARM_STMT_REMIT_TO") & "")
        'CR_params.Add("STMT_MSG", txtSTMT_MESSAGE.Text)
        'CR_params.Add("STMT_MSG_ECP", txtSTMT_MESSAGE_ECP.Text)
        'CR_params.Add("REMIT_TO_LINE_1", ROWs("ARTPARM1").Item("AR_PARM_REMIT_NAME") & "")
        'CR_params.Add("REMIT_TO_LINE_2", ROWs("ARTPARM1").Item("AR_PARM_REMIT_ADDR1") & "")
        'CR_params.Add("REMIT_TO_LINE_3", ROWs("ARTPARM1").Item("AR_PARM_REMIT_CITY") & "," & ROWs("ARTPARM1").Item("AR_PARM_REMIT_STATE") & " " & ROWs("ARTPARM1").Item("AR_PARM_REMIT_ZIP_CODE"))
        'Generate_Report(RPT, "Customer Statements")

        'RPT = "ARRSTMT3"
        'Generate_Report(RPT, "Customers with AR Adjustments")

        RPT = "ARRSTMTM"
        Generate_Report(RPT, "Monthly Freight Charges")
        RPT = "ARRSTMTS"
        Generate_Report(RPT, "Fuel Surcharges")
        RPT = "ARRSTMTE"
        Generate_Report(RPT, "PrimaryECP Profit Credits")
        RPT = "ARRSTMTF"
        Generate_Report(RPT, "Finance Charges")

        If dst.Tables("ARTSTMT0").Rows.Count <> 0 Then
            RPT = "ARRSTMTO"
            Generate_Report(RPT, "Statements Out of Balance")
            'RWU = "N"
        End If

        If dst.Tables("ARTSTMTB").Rows.Count <> 0 Then
            RPT = "ARRSTMTB"
            Generate_Report(RPT, "ECP Credits Out of Balance")
            'RWU = "N"
        End If

        RPT_TITLE = "AR Statements"
        Call Print_GL()
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

            If ASCDATA1.GetDataValue("Select PRD_CLOSE_IND from ASTPCTL1") & "" <> "1" Then
                EMsg &= vbCr & "Period-End has not been Initialized"
            End If

            If Absx1.cmbFor("STMT_MESSAGE_CODE").Value & "" = "" Then
                EMsg &= vbCr & "A Statement Marketing Message has not been Selected"
            End If

            If EMsg = "" Then
                If Absx1.cmbFor("STMT_INSERT_CODE").Value & "" = "" Then
                    If MsgBox("OK to continue without an Insert?", _
                    MsgBoxStyle.YesNo, _
                    "You have not selected an Insert") = MsgBoxResult.No Then
                        EMsg = vbCr & "Select an Insert before Proceeding"
                        Exit Sub
                    End If
                End If
            End If

            Dim Z As String = Absx1.cmbFor("RYP").Text
            Z = Mid(Z, 1, 4) & Mid(Z, 6, 2)
            ASCMAIN1.sql = "Select * from ARTSTMT0 where OPS_YYYYPP = '" & Z & "'"

            Dim rowARTSTMT0 As DataRow = ASCDATA1.GetDataRow
            If rowARTSTMT0 IsNot Nothing Then
                EMsg &= vbCr & "Statements have already been generated for " & Absx1.cmbFor("RYP").Text
            End If
        End If

    End Sub

    Overrides Sub Update_Record()

        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        'End If

        Dim rowARTSTMT0 As DataRow
        With dst.Tables("ARTSTMT0")
            .Rows.Clear()
            .AcceptChanges()
            rowARTSTMT0 = .NewRow
            rowARTSTMT0.Item("OPS_YYYYPP") = RYP
            rowARTSTMT0.Item("STMT_MESSAGE") = txtSTMT_MESSAGE.Text
            rowARTSTMT0.Item("STMT_MESSAGE_ECP") = txtSTMT_MESSAGE_ECP.Text
            rowARTSTMT0.Item("STMT_MESSAGE_CODE") = Absx1.cmbFor("STMT_MESSAGE_CODE").Value
            rowARTSTMT0.Item("STMT_INSERT_CODE") = Absx1.cmbFor("STMT_INSERT_CODE").Value
            rowARTSTMT0.Item("STMT_FIN_CHG_RATE") = ROWs("ARTPARM1").Item("AR_PARM_FIN_CHG_RATE")
            rowARTSTMT0.Item("STMT_XNO") = XNO
            .Rows.Add(rowARTSTMT0)
        End With

        Call ASCMAIN1.Progress("Now Updating Statement History")

        ASCDATA1.ExecuteSQL("Insert into ARTSTMT1 Select * from " & ARTSTMT1)
        ASCDATA1.ExecuteSQL("Insert into ARTSTMT2 Select * from " & ARTSTMT2)

        ASCMAIN1.sql = "Select " _
        & "  Count (*) STMT_COUNT" _
        & ", Sum (TOTAL_DUE) TOTAL_DUE" _
        & " from " & ARTSTMT1
        Dim row As DataRow = ASCDATA1.GetDataRow

        rowARTSTMT0.Item("STMT_COUNT") = row(0)
        rowARTSTMT0.Item("STMT_TOTAL") = row(1)

        ' NEED TO MAKE SURE THAT THESE USE 0 DAYS TERMS

        Call ASCMAIN1.Progress("Now Creating ECP/FIN/FRT/FSC")

        If Not skip_ABS_Calcs Then

            Me.Update_SOTINVH1(SOTINVH1_ECP) ' Primary ECP Credit
            Me.Update_SOTINVH1(SOTINVH1_FIN) ' Finance Charge
            Me.Update_SOTINVH1(SOTINVH1_FRT) ' Monthly Freight
            Me.Update_SOTINVH1(SOTINVH1_FSC) ' Fuel Surcharge

            'ASCMAIN1.sql = "Select Distinct TERM_CODE from " & ARTOPEN1
            'For Each rowTERM_CODE As DataRow In ASCDATA1.GetDataTable.Rows
            '    Dim TERM_CODE As String = rowTERM_CODE.Item("TERM_CODE")

            '    Dim INV_DUE_DATE As Date = Calculate_INV_DUE_DATE(TERM_CODE)
            '    ASCMAIN1.sql = "Update " & ARTOPEN1 _
            '    & " Set INV_DUE_DATE = '" & Format(INV_DUE_DATE, "dd-MMM-yyyy") & "'" _
            '    & " where TERM_CODE = '" & TERM_CODE & "'"
            '    ASCDATA1.ExecuteSQL()
            'Next

            ASCMAIN1.sql = "Update " & ARTOPEN1 & " Set INV_DUE_DATE = INV_DATE where ORDR_TYPE_CODE in ('ECP','FIN')"
            ASCDATA1.ExecuteSQL()
            ASCMAIN1.sql = "Update " & ARTOPEN1 & " Set INV_DUE_DATE = '" & Format(RYP_DATE_DUE, "dd-MMM-yyyy") & "' where ORDR_TYPE_CODE in ('FSC','FRT')"
            ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSQL("Insert into ARTOPEN1 Select * from " & ARTOPEN1)
        End If

        ' Clear out Paid items from Open AR

        Call ASCMAIN1.Progress("Now Clearing Paid Items from Open AR")

        ASCMAIN1.sql = "Update ARTOPEN1 set OPS_YYYYPP_F = NULL"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update ARTOPEN1 set OPS_YYYYPP_F = '" & RYP & "' where INV_BALANCE = 0"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Insert into ARTOPENX Select * from ARTOPEN1 where OPS_YYYYPP_F = '" & RYP & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Delete from ARTOPEN1 where OPS_YYYYPP_F = '" & RYP & "'"
        ASCDATA1.ExecuteSQL()
        ASCMAIN1.sql = "Update ARTCUST1 set FUEL_SURCHARGE_EXEMPTION = NULL where FUEL_SURCHARGE_EXEMPTION is Not Null"
        ASCDATA1.ExecuteSQL()

        Call ASCMAIN1.Progress("Now Copying Files")

        Dim OSG_XML As String = "S:\OSG\" & RYP & "\XML\"
        My.Computer.FileSystem.CreateDirectory(OSG_XML)
        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & RYP & ".zip", OSG_XML & RYP & ".zip")
        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & "ARRSTMT1.XSD", OSG_XML & "ARRSTMT1.XSD")
        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & RYP & ".XML", OSG_XML & RYP & ".XML")
        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & "process.job", OSG_XML & "process.job")
        My.Computer.FileSystem.CopyFile(ASCMAIN1.Folders("Temp") & "osg.scr", OSG_XML & "osg.scr")


        'If ASCMAIN1.Running_in_VS Then
        '    Stop
        'End If
        Me.ftp_File()
        'System.Diagnostics.Process.Start("ftp", " -s:" & ASCMAIN1.Folders("Temp") & "osg.scr")
        'Stop

        ' Update Statement Summary Record

        Call Update_Record_TDA("ARTSTMT0")

        If skip_ABS_Calcs Then
        Else
            Call GL_Update()
        End If

        Try
            Me.Cursor = Cursors.WaitCursor
            Call ASCMAIN1.Progress("Now Loading OSG Billing Client Services Page")
            System.Diagnostics.Process.Start("https://www.osgbilling.com/osg-insight.asp")
            Me.Cursor = Cursors.Default
            Call ASCMAIN1.Progress("")

        Catch ex As Exception

        End Try
    End Sub

    Sub Update_SOTINVH1(ByVal SOTINVH1 As String)

        ASCDATA1.ExecuteSQL("Insert into SOTINVH1 Select * from " & SOTINVH1)

        ASCMAIN1.sql = "INSERT INTO " & ARTOPEN1 _
        & " (CUST_CODE,INV_TYPE,INV_NUM,INV_DATE,CUST_SHIP_TO_NO,POST_CODE,TERM_CODE," _
        & "INV_DUE_DATE,SREP_CODE,STAX_CODE,INV_CUST_PO,ORDR_NO,INV_SALES,INV_DISC," _
        & "INV_FREIGHT,INV_STAX,INV_TOTAL_AMOUNT," _
        & "INV_BALANCE,CUST_CODE_SO,REASON_CODE,INIT_OPER,INIT_DATE,INV_MISC_CHG," _
        & "SEG2_CODE,SEG3_CODE,SEG4_CODE,CURR_CODE,CURR_EXCH_RATE," _
        & "INV_SALES_CURR,INV_DISC_CURR,INV_FREIGHT_CURR,INV_STAX_CURR,INV_MISC_CHG_CURR," _
        & "INV_TOTAL_AMOUNT_CURR,INV_BALANCE_CURR, ORDR_TYPE_CODE, DIVISION_CODE, OPS_YYYYPP)" _
        & " Select CUST_BILL_TO_CUST,INV_TYPE,INV_NO,INV_DATE,CUST_SHIP_TO_NO,POST_CODE,TERM_CODE," _
        & "INV_DATE,SREP_CODE,STAX_CODE,ORDR_CUST_PO,ORDR_NO,INV_SALES,0," _
        & "INV_FREIGHT,INV_STAX,INV_TOTAL_AMOUNT," _
        & "INV_TOTAL_AMOUNT,CUST_CODE,REASON_CODE,INIT_OPER,INIT_DATE,INV_MISC_CHG" _
        & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2") & "'" _
        & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3") & "'" _
        & ",'" & ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4") & "'" _
        & ",'" & ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") & "',1," _
        & "INV_SALES, 0, INV_FREIGHT, INV_STAX, INV_MISC_CHG, " _
        & "INV_TOTAL_AMOUNT, INV_TOTAL_AMOUNT, ORDR_TYPE_CODE, DIVISION_CODE, ORDR_YYYYPP_UPDATED" _
        & " FROM " & SOTINVH1
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Generate_Statements()

        Call ASCMAIN1.Progress("Now Generating Statements for " & RYP)

        ' Prepare Statement Summary Record

        Dim rowARTSTMT0 As DataRow
        With dst.Tables("ARTSTMT0")
            .Rows.Clear()
            .AcceptChanges()
            rowARTSTMT0 = .NewRow
            rowARTSTMT0.Item("OPS_YYYYPP") = RYP
            rowARTSTMT0.Item("STMT_MESSAGE") = txtSTMT_MESSAGE.Text
            rowARTSTMT0.Item("STMT_MESSAGE_ECP") = txtSTMT_MESSAGE_ECP.Text
            .Rows.Add(rowARTSTMT0)
        End With


        ' Prepare Work Tables

        Call ASCMAIN1.Progress("Now Preparing Work Tables")

        ARTSTMTX = ASCMAIN1.Temp_Table( _
        "Select CUST_CODE, AGE_1, AGE_2, AGE_3, AGE_4, PAST_DUE_DR_AMT, TOTAL_DUE, DAY_DOLLARS" _
        & ", TYP_I_OPEN, TYP_R_OPEN, TYP_C_OPEN, TYP_D_OPEN, TYP_B_OPEN, TYP_O_OPEN" _
        & ", TYP_ECP, TYP_FIN, CUST_B2C_IND, NO_FINANCE_CHARGE " _
        & " from ARTSTMT1 where ROWNUM < 1")
        ASCDATA1.ExecuteSQL("Alter Table " & ARTSTMTX & " Add Primary Key (CUST_CODE)")

        ASCMAIN1.sql = "" _
        & " Select B2C_CUST_CODE, CUST_CODE, CUST_SHIP_TO_NO from ARTCUST2 " _
        & " where B2C_CUST_CODE is Not Null" _
        & " union " _
        & " Select B2C_CUST_CODE, CUST_CODE, Null CUST_SHIP_TO_NO from ARTCUST1 " _
        & " where B2C_CUST_CODE is Not Null "
        ARTCUST1_B2C = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_B2C & " Add Primary Key (B2C_CUST_CODE)")
        ASCDATA1.ExecuteSQL("Create Index I_" & ARTCUST1_B2C & "_1 ON " & ARTCUST1_B2C & " (CUST_CODE)")

        ASCMAIN1.sql = "Select CUST_CODE" _
        & ", Sum (INV_SALES) INV_SALES" _
        & " from SOTINVH1 where ORDR_YYYYPP_UPDATED = '" & RYP & "' " _
        & " and INV_TYPE = 'I' and NVL(INV_SALES,0) > 0" _
        & " group by CUST_CODE"
        ARTCUST1_has_sales = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ARTCUST1_has_sales & " Add Primary Key (CUST_CODE)")


        ' Aged A/R

        Call ASCMAIN1.Progress("Now Generating Aged AR for " & RYP)

        Dim DT(4) As String
        For i As Integer = 0 To 4
            DT(i) = Format(LookUp("GLTPARM2", ASCMAIN1.Period_Calc(RYP, -1 * i)).Item("PRD_END_DATE"), "dd-MMM-yyyy")
        Next

        ASCMAIN1.sql = "Insert into " & ARTSTMTX _
        & " (CUST_CODE, AGE_1, AGE_2, AGE_3, AGE_4, PAST_DUE_DR_AMT, TOTAL_DUE, DAY_DOLLARS" _
        & ", TYP_I_OPEN, TYP_R_OPEN, TYP_C_OPEN, TYP_D_OPEN, TYP_B_OPEN, TYP_O_OPEN" _
        & ", CUST_B2C_IND, NO_FINANCE_CHARGE)" _
        & " Select X.*, ARTCUST1.CUST_B2C_IND, ARTCUST1.NO_FINANCE_CHARGE" _
        & " from ARTCUST1, (" _
        & " Select CUST_CODE" _
        & ", SUM (CASE WHEN INV_DATE > '" & DT(1) & "'                                 THEN INV_BALANCE ELSE 0 END) AGE_1" _
        & ", SUM (CASE WHEN INV_DATE > '" & DT(2) & "' AND INV_DATE <= '" & DT(1) & "' THEN INV_BALANCE ELSE 0 END) AGE_2" _
        & ", SUM (CASE WHEN INV_DATE > '" & DT(3) & "' AND INV_DATE <= '" & DT(2) & "' THEN INV_BALANCE ELSE 0 END) AGE_3" _
        & ", SUM (CASE WHEN INV_DATE                                <= '" & DT(3) & "' THEN INV_BALANCE ELSE 0 END) AGE_4" _
        & ", SUM (CASE WHEN INV_DATE                                <= '" & DT(2) & "' AND INV_TYPE IN ('I','D','B') THEN INV_BALANCE ELSE 0 END) PAST_DUE_DR_AMT" _
        & ", SUM (INV_BALANCE) TOTAL_DUE" _
        & ", SUM (INV_BALANCE * (TO_DATE('" & DT(0) & "') - INV_DATE)) DAY_DOLLARS" _
        & ", SUM (DECODE(INV_TYPE,'I',INV_BALANCE,0)) TYP_I_OPEN" _
        & ", SUM (DECODE(INV_TYPE,'R',INV_BALANCE,0)) TYP_R_OPEN" _
        & ", SUM (DECODE(INV_TYPE,'C',INV_BALANCE,0)) TYP_C_OPEN" _
        & ", SUM (DECODE(INV_TYPE,'D',INV_BALANCE,0)) TYP_D_OPEN" _
        & ", SUM (DECODE(INV_TYPE,'B',INV_BALANCE,0)) TYP_B_OPEN" _
        & ", SUM (DECODE(INV_TYPE,'O',INV_BALANCE,0)) TYP_O_OPEN" _
        & " from ARTOPEN1 where NVL(OPS_YYYYPP,'000000') <= '" & ASCMAIN1.CYP & "' group by CUST_CODE" _
        & ") X where X.CUST_CODE = ARTCUST1.CUST_CODE"
        ASCDATA1.ExecuteSQL()

        Call ASCMAIN1.AnalyzeTable(ARTSTMTX)


        ' Calculate PrimaryECP Profit Credits

        Call ASCMAIN1.Progress("Now Generating PrimaryECP Profit Credits for " & RYP)

        ' Credit Amount is the sum of the Profit Amounts
        '  from each of the settled AR items in the B2C accounts
        '  connected with this account

        SOTINVH1_ECP = Me.Generate_SOTINVH1("ECP", _
        " (INV_TYPE, INV_NO, CUST_CODE, CUST_SHIP_TO_NO" _
        & ", INV_MISC_CHG, ORDR_CUST_PO)" _
        & " Select 'X' INV_TYPE, ROWNUM INV_NO" _
        & ", CUST_CODE, CUST_SHIP_TO_NO, INV_MISC_CHG, ORDR_CUST_PO from (" _
        & " Select ARTCUST1_B2C.CUST_CODE CUST_CODE" _
        & ", ARTCUST1_B2C.CUST_SHIP_TO_NO CUST_SHIP_TO_NO" _
        & ", Sum (-1 * NVL(INV_PROFIT_B2C,0)) INV_MISC_CHG" _
        & ", 'B2C ' || ARTOPEN1.CUST_CODE ORDR_CUST_PO" _
        & " from " & ARTCUST1_B2C & " ARTCUST1_B2C, ARTOPEN1" _
        & " where ARTCUST1_B2C.B2C_CUST_CODE = ARTOPEN1.CUST_CODE" _
        & " and ARTOPEN1.INV_BALANCE = 0" _
        & " and ARTOPEN1.INV_NUM NOT LIKE '00000%'" _
        & " group by ARTCUST1_B2C.CUST_CODE" _
        & ", ARTCUST1_B2C.CUST_SHIP_TO_NO" _
        & ", ARTOPEN1.CUST_CODE)")


        ' Monthly Freight Invoices
        ' NEED TO DETERMINE THE CORRECT FRT CONTRACT NO BASED ON DATES

        Call ASCMAIN1.Progress("Now Generating Monthly Freight Invoices " & RYP)

        ASCMAIN1.sql = "Select Max(REL_WEEK) from GLTPARM3 where YYYYPP = '" & RYP & "'"
        Dim WKS As Integer = ASCDATA1.GetDataValue
        Dim FRT_AMT_CALC As String = ""
        If WKS = 5 Then
            FRT_AMT_CALC = "Decode(ARTCUST3.FRT_CONT_MONTHLY_LOCK_AMT,'1',ARTCUST3.FRT_CONT_AMT, TRUNC(ARTCUST3.FRT_CONT_AMT * 1.25) + .99) FRT_AMT"
        Else
            FRT_AMT_CALC = "ARTCUST3.FRT_CONT_AMT FRT_AMT"
        End If

        SOTINVH1_FRT = Me.Generate_SOTINVH1("FRT", _
        " (INV_TYPE, INV_NO, CUST_CODE, INV_FREIGHT, FRT_CONT_NO)" _
        & " Select 'X' INV_TYPE, ROWNUM INV_NO, ARTCUST3.CUST_CODE, " _
        & FRT_AMT_CALC & ", ARTCUST3.FRT_CONT_NO " _
        & " from " & ARTCUST3 & " ARTCUST3, " & ARTCUST1_has_sales & " ARTCUST1_has_sales" _
        & " where ARTCUST3.FRT_CONT_AMT <> 0 and FRT_CONT_TYPE = 'M'" _
        & " and ARTCUST1_has_sales.CUST_CODE = ARTCUST3.CUST_CODE")


        ' Fuel Surcharge

        Call ASCMAIN1.Progress("Now Generating Monthly Fuel Surcharge Invoices " & RYP)

        SOTINVH1_FSC = Me.Generate_SOTINVH1("FSC", _
        "(INV_TYPE, INV_NO, CUST_CODE, INV_MISC_CHG)" _
        & " Select 'X' INV_TYPE, ROWNUM INV_NO, ARTCUST1.CUST_CODE " _
        & ", SOTFSCC1.FUEL_SURCHARGE_RATE" _
        & " from SOTFSCC1,ARTCUST1,TATTERM1" _
        & "," & ARTSTMTX & " ARTSTMTX" _
        & "," & ARTCUST1_has_sales & " ARTCUST1_has_sales" _
        & " where ARTCUST1.FUEL_SURCHARGE_CODE = SOTFSCC1.FUEL_SURCHARGE_CODE" _
        & "   and SOTFSCC1.FUEL_SURCHARGE_RATE <> 0 " _
        & "   and NVL(ARTCUST1.FUEL_SURCHARGE_EXEMPTION,'0') <> '1'" _
        & "   and TATTERM1.TERM_CODE (+) = ARTCUST1.TERM_CODE " _
        & "   and NVL(TATTERM1.TERM_TYPE,'N') not in ('C','D')" _
        & "   and ARTSTMTX.CUST_CODE = ARTCUST1.CUST_CODE" _
        & "   and NVL(ARTSTMTX.TOTAL_DUE,0) > 50" _
        & "   and ARTCUST1_has_sales.CUST_CODE = ARTCUST1.CUST_CODE")


        ' Finance Charges

        Call ASCMAIN1.Progress("Now Generating Finance Charges for " & RYP)

        ' as per Laree
        ' 1) Finance Charges are generated on the total in the 90+ column,
        '    which may include finance charges generated 3 months back 
        '    which have not yet been paid
        ' 2) Open and Unapplied Credits from < 90+ day columns are not used 
        '    to reduce the 90+ amount
        ' 3) PrimaryECP profit credit which is generated and remains open 
        '    and unapplied (or unre-imbursed) is used to reduce the 90+ day amt
        '    which is used to calculate the finance charge

        ASCMAIN1.sql = "Update " & ARTSTMTX & " ARTSTMTX" _
        & " SET TYP_ECP = (Select Sum (INV_TOTAL_AMOUNT) from " & SOTINVH1_ECP _
        & " where CUST_CODE = ARTSTMTX.CUST_CODE)"
        ASCDATA1.ExecuteSQL()

        ' WJZ Original Calculation
        ASCMAIN1.sql = "Update " & ARTSTMTX _
        & " SET TYP_FIN = (AGE_4 + NVL(TYP_ECP, 0))" _
        & " * (" & CStr(Val(ROWs("ARTPARM1").Item("AR_PARM_FIN_CHG_RATE") & "")) & " / 12) / 100" _
        & " where (AGE_4 + NVL(TYP_ECP, 0)) > 0 " _
        & " and NVL(NO_FINANCE_CHARGE,'0') <> '1'" _
        & " and NVL(CUST_B2C_IND,'0') <> '1'"
        'ASCDATA1.ExecuteSQL()

        ' Meeting in Ron's office - Jeff's idea - double bang money new to 90 days
        ASCMAIN1.sql = "Update " & ARTSTMTX _
        & " SET TYP_FIN = (NVL(AGE_4,0) + NVL(PAST_DUE_DR_AMT, 0))" _
        & " * (" & CStr(Val(ROWs("ARTPARM1").Item("AR_PARM_FIN_CHG_RATE") & "")) & " / 12) / 100" _
        & " where (NVL(AGE_4,0) + NVL(PAST_DUE_DR_AMT, 0)) > 0 " _
        & " and NVL(TOTAL_DUE,0) > 0" _
        & " and NVL(NO_FINANCE_CHARGE,'0') <> '1'" _
        & " and NVL(CUST_B2C_IND,'0') <> '1'"
        'ASCDATA1.ExecuteSQL()

        ' Final Answer - SEE WJZ EMAIL 11/07/2007
        ASCMAIN1.sql = "Update " & ARTSTMTX _
        & " SET TYP_FIN = (PAST_DUE_DR_AMT + NVL(TYP_C_OPEN, 0) + NVL(TYP_O_OPEN, 0) + NVL(TYP_R_OPEN, 0))" _
        & " * (" & CStr(Val(ROWs("ARTPARM1").Item("AR_PARM_FIN_CHG_RATE") & "")) & " / 12) / 100" _
        & " where (PAST_DUE_DR_AMT + NVL(TYP_C_OPEN, 0) + NVL(TYP_O_OPEN, 0) + NVL(TYP_R_OPEN, 0)) > 0 " _
        & " and NVL(NO_FINANCE_CHARGE,'0') <> '1'" _
        & " and NVL(CUST_B2C_IND,'0') <> '1'"
        ASCDATA1.ExecuteSQL()

        SOTINVH1_FIN = Me.Generate_SOTINVH1("FIN", _
        "(INV_TYPE, INV_NO, CUST_CODE, INV_MISC_CHG)" _
        & " Select 'X' INV_TYPE, ROWNUM INV_NO, CUST_CODE " _
        & ", TYP_FIN" _
        & " from " & ARTSTMTX _
        & " where TYP_FIN <> 0")

        Me.Create_ARTSTMT1()

        Me.Reconcile_Statements()

    End Sub

    Sub Reconcile_Statements()

        ' Create table of customers whose Statements "do not foot"

        ASCMAIN1.sql = "SELECT ARTSTMT1.CUST_CODE, ARTCUST1.CUST_NAME" _
        & ", ARTSTMT1.TOTAL_DUE, X.TOTAL_ITEMS" _
        & " FROM " & ARTSTMT1 & " ARTSTMT1, " _
        & " (SELECT CUST_CODE, OPS_YYYYPP, SUM (INV_TOTAL_AMOUNT) TOTAL_ITEMS " _
        & " from " & ARTSTMT2 & " ARTSTMT2 " _
        & "  where TYPE_SEQ <> 10 and TYPE_SEQ <> 11 " _
        & " group by CUST_CODE, OPS_YYYYPP) X" _
        & " , ARTCUST1 " _
        & " WHERE ARTSTMT1.CUST_CODE = X.CUST_CODE (+)" _
        & " AND ARTSTMT1.OPS_YYYYPP = X. OPS_YYYYPP (+)" _
        & " AND ARTCUST1.CUST_CODE = ARTSTMT1.CUST_CODE" _
        & " AND NVL(ARTSTMT1.TOTAL_DUE,0) <> NVL(X.TOTAL_ITEMS,0)"
        ARTSTMTO = ASCMAIN1.Temp_Table
        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & ARTSTMTO, "ARTSTMTO", 1))

        ' Create a table of customers whose statements have an ECP credit which is not in balance 
        '  with Profit calculated from Settled B2C Items
        ASCMAIN1.sql = "SELECT X.CUST_CODE, ARTCUST1.CUST_NAME" _
        & ", SUM (X.B2B) B2B, SUM (X.B2C) B2C, SUM (X.CALC) CALC, SUM (X.ECP) ECP" _
        & ", SUM (X.CALC + X.ECP) DIFF FROM ARTCUST1, (" _
        & " SELECT CUST_CODE, SUM (INV_TOTAL_AMOUNT_B2B) B2B, SUM (INV_TOTAL_AMOUNT_B2C) B2C " _
        & ", SUM (NVL(INV_TOTAL_AMOUNT_B2C,0) - NVL(INV_TOTAL_AMOUNT_B2B,0)) CALC, 0 ECP" _
        & " FROM " & ARTSTMT2 & " ARTSTMT2 WHERE TYPE_SEQ = 10" _
        & " GROUP BY CUST_CODE" _
        & " UNION" _
        & " SELECT CUST_CODE, 0 B2B, 0 B2C, 0 CALC, TYP_ECP ECP" _
        & " FROM " & ARTSTMT1 & " WHERE TYP_ECP <> 0" _
        & ") X WHERE ARTCUST1.CUST_CODE = X.CUST_CODE" _
        & " GROUP BY X.CUST_CODE, ARTCUST1.CUST_NAME HAVING SUM (X.CALC) + SUM (X.ECP) <> 0"
        ARTSTMTB = ASCMAIN1.Temp_Table
        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & ARTSTMTB, "ARTSTMTB", 1))


        '& " SELECT CUST_CODE, 0 B2B, 0 B2C, 0 CALC, SUM (INV_TOTAL_AMOUNT) ECP" _
        '& " FROM " & ARTSTMT2 & " WHERE TYPE_SEQ IN (4,5) AND ORDR_TYPE_CODE = 'ECP'" _
        '& " AND INV_TOTAL_AMOUNT <> 0" _
        '& " GROUP BY CUST_CODE" _


        ASCMAIN1.sql = "Update " & ARTSTMT1 _
        & " Set MAIL_STATEMENT = 'P' where MAIL_STATEMENT = 'Y' and CUST_CODE in " _
        & " (Select CUST_CODE from " & ARTSTMTO & ")"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update " & ARTSTMT1 _
            & " Set MAIL_STATEMENT = 'P' where MAIL_STATEMENT = 'Y' and CUST_CODE in " _
            & " (Select CUST_CODE from " & ARTSTMTB & ")"
        ASCDATA1.ExecuteSQL()

    End Sub

    Function Generate_SOTINVH1( _
    ByVal ORDR_TYPE_CODE As String, _
    ByVal sql As String) As String

        Dim SOTINVH1 As String = ASCMAIN1.Temp_Table("Select * from SOTINVH1 where ROWNUM < 1")
        ASCDATA1.ExecuteSQL("Alter Table " & SOTINVH1 & " Add Primary Key (INV_TYPE, INV_NO)")
        ASCDATA1.ExecuteSQL("Create Index I_" & SOTINVH1 & "_1 on " & SOTINVH1 & " (CUST_CODE)")

        If Not skip_ABS_Calcs Then
            ASCMAIN1.sql = "Insert into " & SOTINVH1 & sql
            ASCDATA1.ExecuteSQL()
        End If
        ASCMAIN1.AnalyzeTable(SOTINVH1)

        ASCMAIN1.sql = "Update " & SOTINVH1 _
        & " Set INV_TOTAL_AMOUNT = NVL(INV_FREIGHT,0) + NVL(INV_MISC_CHG,0) " _
        & ", INV_SALES = 0, INV_COGS = 0, INV_TOTAL_AMOUNT_B2C = 0" _
        & ", INV_STAX = 0, INV_SAMPLE_SURCHARGE = 0" _
        & ", INV_FREIGHT = NVL(INV_FREIGHT,0), INV_MISC_CHG = NVL(INV_MISC_CHG,0)" _
        & ", DIVISION_CODE = 'ODG'" _
        & ", INIT_OPER = '" & ASCMAIN1.USER_ID & "', INIT_DATE = SYSDATE" _
        & ", INV_PRINTED = 'X', ORDR_TYPE_CODE = '" & ORDR_TYPE_CODE & "'" _
        & ", ORDR_YYYYPP_UPDATED = '" & RYP & "'" _
        & ", REGISTER_XNO = '" & XNO & "', REGISTER_IND = '1'" _
        & ", REGISTER_DATE = '" & RYP_DATE.ToString("dd-MMM-yyyy") & "'" _
        & ", INV_DATE = '" & RYP_DATE.ToString("dd-MMM-yyyy") & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
        & " Begin Declare Cursor C1 is " _
        & " Select * from SOTTYPE1 where ORDR_TYPE_CODE = '" & ORDR_TYPE_CODE & "';" _
        & "   Begin For R1 in C1 Loop" _
        & "     Update " & SOTINVH1 _
        & " Set POST_CODE = R1.POST_CODE" _
        & ", TERM_CODE = R1.TERM_CODE" _
        & ", REASON_CODE = R1.REASON_CODE" _
        & ", MISC_CHG_CODE = R1.MISC_CHG_CODE;" _
        & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
        & " Begin Declare Cursor C1 is " _
        & " Select SOTINVH1.INV_TYPE, SOTINVH1.INV_NO, SOTINVH1.CUST_CODE" _
        & ", ARTCUST1.SREP_CODE, ARTCUST1.TERM_CODE, ARTCUST1.POST_CODE" _
        & " from " & SOTINVH1 & " SOTINVH1,ARTCUST1 " _
        & " where ARTCUST1.CUST_CODE = SOTINVH1.CUST_CODE " _
        & " order by SOTINVH1.CUST_CODE;" _
        & "   Begin For R1 in C1 Loop" _
        & "     Update " & SOTINVH1 _
        & " Set POST_CODE = NVL(POST_CODE,R1.POST_CODE)" _
        & ", TERM_CODE = NVL(TERM_CODE,R1.TERM_CODE), SREP_CODE = R1.SREP_CODE" _
        & ", INV_TYPE = DECODE(ORDR_TYPE_CODE,'ECP','C','I')" _
        & ", INV_NO = TAPCTLN1('SOTINVH1.INV_NO',10,1)" _
        & ", CUST_BILL_TO_CUST = R1.CUST_CODE" _
        & " where INV_TYPE = 'X' and INV_NO = R1.INV_NO;" _
        & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        Dim TABLE_NAME As String = "SOTINVH1_" & ORDR_TYPE_CODE
        Dim TABLE_NAME_X As String = ""

        ASCMAIN1.sql = "Select * from " & SOTINVH1
        dst.Tables.Add(ASCDATA1.GetDataTable(, TABLE_NAME))

        For Each COLUMN_NAME As String In New String() {"POST_CODE", "MISC_CHG_CODE"}
            TABLE_NAME_X = TABLE_NAME & "_" & COLUMN_NAME
            With dst.Tables.Add
                .TableName = TABLE_NAME_X
                .Columns.Add(COLUMN_NAME)
                .Columns.Add("INV_TOTAL_AMOUNT")
            End With
            For Each row As DataRow In ASCMAIN1.Distinct_Values("", _
            dst.Tables(TABLE_NAME), COLUMN_NAME).Rows
                dst.Tables(TABLE_NAME_X).Rows.Add(New String() {row.Item(COLUMN_NAME)})
            Next
            dst.Relations.Add(TABLE_NAME_X _
            , dst.Tables(TABLE_NAME_X).Columns(COLUMN_NAME) _
            , dst.Tables(TABLE_NAME).Columns(COLUMN_NAME))
            dst.Tables(TABLE_NAME_X).Columns("INV_TOTAL_AMOUNT").Expression _
            = "SUM (CHILD.INV_TOTAL_AMOUNT)"
        Next

        Me.GL_Interface(ORDR_TYPE_CODE)

        Return SOTINVH1

    End Function

    Function Calculate_INV_DUE_DATE(ByVal TERM_CODE As String) As Object

        Dim INV_BASE_DATE As Date = RYP_DATE
        Dim INV_DUE_DATE As Object = Nothing

        If TERM_CODE = "" Then
            Return INV_DUE_DATE
        End If

        Dim rowTATTERM1 As DataRow = LookUp("TATTERM1", TERM_CODE, True)

        Select Case rowTATTERM1.Item("TERM_DUE_TYPE") & ""

            Case "D"
                INV_DUE_DATE = INV_BASE_DATE.AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DUE") & ""))

            Case "E"

                Dim ADD_MONTHS_BASE As Integer = 1
                Dim TERM_CUTOFF_DAY As Integer = Val(rowTATTERM1.Item("TERM_CUTOFF_DAY") & "")
                Dim BASE_DD As Integer = Val(Format(INV_BASE_DATE, "dd"))
                Dim TERM_DAYS_DUE As Integer = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & "")
                Dim TERM_ADDL_MOS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_MOS") & "")
                Dim INV_BASE_DATEx As String = Format(INV_BASE_DATE, "MM/dd/yyyy")

                Select Case rowTATTERM1.Item("TERM_EOM_TYPE") & ""
                    Case "F"
                        Dim rowGLTPARM2 As DataRow = Fill_Record("GLTPARM2", Format(INV_BASE_DATE, "dd-MMM-yyyy"), True)
                        Dim YYYYMM As String = ASCMAIN1.Get_YYYYMM(rowGLTPARM2.Item("OPS_YYYYPP"), 0)
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "C"
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case "S"
                        If BASE_DD <= TERM_CUTOFF_DAY _
                        And BASE_DD <= TERM_DAYS_DUE Then
                            ADD_MONTHS_BASE = 0
                        End If
                        Dim YYYYMM As String = Format(INV_BASE_DATE, "yyyyMM")
                        INV_DUE_DATE = CDate(Mid(YYYYMM, 5, 2) & "/" & Format(TERM_DAYS_DUE, "00") & "/" & Mid(YYYYMM, 1, 4)).AddMonths(ADD_MONTHS_BASE)
                    Case Else
                        INV_DUE_DATE = INV_BASE_DATE
                End Select
                If TERM_ADDL_MOS > 0 Then
                    INV_DUE_DATE = INV_DUE_DATE.AddMonths(TERM_ADDL_MOS)
                End If

        End Select

        Return INV_DUE_DATE

    End Function

    Sub GL_Interface(ByVal ORDR_TYPE_CODE As String)

        Dim JOURNAL_LNO As Integer = 0
        Dim rowSOTTYPE1 As DataRow = dst.Tables("SOTTYPE1").Rows.Find(ORDR_TYPE_CODE)
        Dim DETL_POSTING_AMT As Double
        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_TYPE As String = rowSOTTYPE1.Item("JOURNAL_TYPE")

        For Each row As DataRow In dst.Tables("SOTINVH1_" & ORDR_TYPE_CODE & "_POST_CODE").Rows
            Dim POST_CODE As String = row.Item("POST_CODE")
            DETL_POSTING_AMT = Val(row.Item("INV_TOTAL_AMOUNT") & "")
            Dim rowARTPOST1 As DataRow = dst.Tables("ARTPOST1").Rows.Find(POST_CODE)

            Me.GL_Interface_Record( _
            JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, _
            rowARTPOST1.Item("ACCT_CODE"), _
            IIf(rowARTPOST1("SEG2_CODE") & "" = "", ROWs("GLTPARM1")("GL_PARM_DEF_SEG2"), rowARTPOST1("SEG2_CODE")), _
            IIf(rowARTPOST1("SEG3_CODE") & "" = "", ROWs("GLTPARM1")("GL_PARM_DEF_SEG3"), rowARTPOST1("SEG3_CODE")), _
            IIf(rowARTPOST1("SEG4_CODE") & "" = "", ROWs("GLTPARM1")("GL_PARM_DEF_SEG4"), rowARTPOST1("SEG4_CODE")), _
            DETL_POSTING_AMT)
        Next

        For Each row As DataRow In dst.Tables("SOTINVH1_" & ORDR_TYPE_CODE & "_MISC_CHG_CODE").Rows
            Dim MISC_CHG_CODE As String = row.Item("MISC_CHG_CODE")
            DETL_POSTING_AMT = -1 * Val(row.Item("INV_TOTAL_AMOUNT") & "")
            Dim rowSOTMISC1 As DataRow = dst.Tables("SOTMISC1").Rows.Find(MISC_CHG_CODE)

            Me.GL_Interface_Record( _
            JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, _
            rowSOTMISC1.Item("ACCT_CODE"), _
            IIf(rowSOTMISC1("SEG2_CODE") & "" = "", ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2"), rowSOTMISC1("SEG2_CODE")), _
            IIf(rowSOTMISC1("SEG3_CODE") & "" = "", ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3"), rowSOTMISC1("SEG3_CODE")), _
            IIf(rowSOTMISC1("SEG4_CODE") & "" = "", ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4"), rowSOTMISC1("SEG4_CODE")), _
            DETL_POSTING_AMT)
        Next

    End Sub

    Sub GL_Interface_Record( _
    ByVal JOURNAL_TYPE As String, _
    ByVal JOURNAL_NO As String, _
    ByRef JOURNAL_LNO As Integer, _
    ByVal ACCT_CODE As String, _
    ByVal SEG2_CODE As String, _
    ByVal SEG3_CODE As String, _
    ByVal SEG4_CODE As String, _
    ByVal DETL_POSTING_AMT As Double)

        If DETL_POSTING_AMT <> 0 Then
            Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
            rowGLTINTF1("OPS_YYYYPP") = RYP
            rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
            JOURNAL_LNO += 1
            rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
            rowGLTINTF1("ACCT_CODE") = ACCT_CODE
            rowGLTINTF1("SEG2_CODE") = SEG2_CODE
            rowGLTINTF1("SEG3_CODE") = SEG3_CODE
            rowGLTINTF1("SEG4_CODE") = SEG4_CODE
            rowGLTINTF1("DETL_CTL_DATE") = DateValue(Format(DATETIME_STAMP, "MM/dd/yyyy"))
            rowGLTINTF1("DETL_POSTING_AMT") = System.Math.Round(DETL_POSTING_AMT, 2)
            rowGLTINTF1("DETL_EXE_NO") = XNO
            rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
            dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
        End If

    End Sub

    Sub Create_ARTSTMT1()

        Call ASCMAIN1.Progress("Now Creating Statement Summary")

        ASCMAIN1.sql = " SELECT '" & RYP & "' OPS_YYYYPP" _
        & ", CUST_CODE, SUM (BALFWD) BALFWD" _
        & ", SUM (TYP_I) TYP_I, SUM (TYP_R) TYP_R, SUM (TYP_C) TYP_C" _
        & ", SUM (TYP_D) TYP_D, SUM (TYP_B) TYP_B, SUM (TYP_O) TYP_O" _
        & ", SUM (TYP_ECP) TYP_ECP" _
        & ", SUM (TYP_FRT) TYP_FRT" _
        & ", SUM (TYP_FSC) TYP_FSC" _
        & ", SUM (TYP_FIN) TYP_FIN" _
        & ", SUM (PYMTS) PYMTS" _
        & ", SUM (ADJS) ADJS" _
        & ", SUM (AGE_1) AGE_1, SUM (AGE_2) AGE_2, SUM (AGE_3) AGE_3, SUM (AGE_4) AGE_4, SUM (PAST_DUE_DR_AMT) PAST_DUE_DR_AMT" _
        & ", SUM (TYP_I_OPEN) TYP_I_OPEN, SUM (TYP_R_OPEN) TYP_R_OPEN, SUM (TYP_C_OPEN) TYP_C_OPEN" _
        & ", SUM (TYP_D_OPEN) TYP_D_OPEN, SUM (TYP_B_OPEN) TYP_B_OPEN, SUM (TYP_O_OPEN) TYP_O_OPEN" _
        & ", SUM (TOTAL_DUE) TOTAL_DUE, SUM (DAY_DOLLARS) DAY_DOLLARS" _
        & " FROM (" & vbCr _
        & " SELECT CUST_CODE, TOTAL_DUE BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN, 0 PYMTS, 0 ADJS" _
        & ", 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " from ARTSTMT1 where OPS_YYYYPP = '" & LYP & "'" _
        & "  and NVL(TOTAL_DUE,0) <> 0" _
        & " UNION" & vbCr _
        & " SELECT CUST_CODE, 0 BALFWD" _
        & ", SUM (DECODE(INV_TYPE,'I',INV_TOTAL_AMOUNT,0)) TYP_I" _
        & ", SUM (DECODE(INV_TYPE,'R',INV_TOTAL_AMOUNT,0)) TYP_R" _
        & ", SUM (DECODE(INV_TYPE,'C',INV_TOTAL_AMOUNT,0)) TYP_C" _
        & ", SUM (DECODE(INV_TYPE,'D',INV_TOTAL_AMOUNT,0)) TYP_D" _
        & ", SUM (DECODE(INV_TYPE,'B',INV_TOTAL_AMOUNT,0)) TYP_B" _
        & ", SUM (DECODE(INV_TYPE,'O',INV_TOTAL_AMOUNT,0)) TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN, 0 PYMTS, 0 ADJS" _
        & ", 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM SOTINVH1 WHERE ORDR_YYYYPP_UPDATED = '" & RYP & "' " _
        & " GROUP BY CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", SUM (INV_TOTAL_AMOUNT) TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN" _
        & ", 0 PYMTS, 0 ADJS, 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & " , 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM " & SOTINVH1_ECP & " GROUP BY CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, SUM (INV_TOTAL_AMOUNT) TYP_FRT, 0 TYP_FSC, 0 TYP_FIN" _
        & ", 0 PYMTS, 0 ADJS, 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM " & SOTINVH1_FRT & " GROUP BY CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, SUM (INV_TOTAL_AMOUNT) TYP_FSC, 0 TYP_FIN" _
        & ", 0 PYMTS, 0 ADJS, 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM " & SOTINVH1_FSC & " GROUP BY CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, SUM (INV_TOTAL_AMOUNT) TYP_FIN" _
        & ", 0 PYMTS, 0 ADJS, 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM " & SOTINVH1_FIN & " GROUP BY CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT ARTPYMT2.CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN, SUM (ARTPYMT2.CUST_PYMT_AMT) PYMTS, 0 ADJS" _
        & ", 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM ARTPYMT1,ARTPYMT2 WHERE ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " AND ARTPYMT2.CUST_CODE is Not Null" _
        & " AND NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & " GROUP BY ARTPYMT2.CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT ARTPYMT2.CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN, 0 PYMTS" _
        & ", SUM (NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0)) ADJS" _
        & " , 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT3 WHERE ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " AND ARTPYMT2.CUST_CODE is Not Null" _
        & " AND NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & " AND ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " AND ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " AND NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) <> 0" _
        & " GROUP BY ARTPYMT2.CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT ARTPYMT2.CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN, 0 PYMTS" _
        & ", SUM (NVL(ARTPYMT4.GL_DIST_AMT,0)) ADJS" _
        & ", 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT4 WHERE ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " AND ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " AND ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " AND NVL(ARTPYMT4.GL_DIST_AMT,0) <> 0" _
        & " AND ARTPYMT2.CUST_CODE is Not Null" _
        & " AND NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & " GROUP BY ARTPYMT2.CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT ARTPYMT2.CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D" _
        & ", SUM (CASE WHEN NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '1' AND ARTPYMT5.INV_TYPE_CB = 'B' THEN NVL(ARTPYMT5.GL_DIST_AMT,0) ELSE 0 END) TYP_B" _
        & ", SUM (CASE WHEN NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '1' AND ARTPYMT5.INV_TYPE_CB = 'O' THEN NVL(ARTPYMT5.GL_DIST_AMT,0) ELSE 0 END) TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN, 0 PYMTS" _
        & ", SUM (CASE WHEN NVL(ARTPYMT5.CHARGEBACK_IND,'0') = '0' THEN NVL(ARTPYMT5.GL_DIST_AMT,0) ELSE 0 END) ADJS" _
        & ", 0 AGE_1, 0 AGE_2, 0 AGE_3, 0 AGE_4, 0 PAST_DUE_DR_AMT" _
        & ", 0 TYP_I_OPEN, 0 TYP_R_OPEN, 0 TYP_C_OPEN, 0 TYP_D_OPEN, 0 TYP_B_OPEN, 0 TYP_O_OPEN" _
        & ", 0 TOTAL_DUE, 0 DAY_DOLLARS" _
        & " FROM ARTPYMT1,ARTPYMT2,ARTPYMT5 WHERE ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & " AND ARTPYMT2.PYMT_BATCH_NO = ARTPYMT1.PYMT_BATCH_NO" _
        & " AND ARTPYMT2.CUST_CODE is Not Null" _
        & " AND NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & " AND ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & " AND ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & " AND NVL(ARTPYMT5.GL_DIST_AMT,0) <> 0" _
        & " GROUP BY ARTPYMT2.CUST_CODE" _
        & " UNION" & vbCr _
        & " SELECT CUST_CODE, 0 BALFWD" _
        & ", 0 TYP_I, 0 TYP_R, 0 TYP_C, 0 TYP_D, 0 TYP_B, 0 TYP_O" _
        & ", 0 TYP_ECP, 0 TYP_FRT, 0 TYP_FSC, 0 TYP_FIN, 0 PYMTS, 0 ADJS" _
        & ", AGE_1, AGE_2, AGE_3, AGE_4, PAST_DUE_DR_AMT" _
        & ", TYP_I_OPEN, TYP_R_OPEN, TYP_C_OPEN, TYP_D_OPEN, TYP_B_OPEN, TYP_O_OPEN" _
        & ", TOTAL_DUE, DAY_DOLLARS" _
        & " FROM " & ARTSTMTX _
        & " ) GROUP BY CUST_CODE"

        Dim sql As String = ASCMAIN1.sql

        If ARTSTMT1 = "" Then
            ASCMAIN1.sql = "Select * from ARTSTMT1 where ROWNUM < 1"
            ARTSTMT1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTSTMT1 _
            & " Add Primary Key (OPS_YYYYPP, CUST_CODE)")
            ASCDATA1.ExecuteSQL("Create Index I_" & ARTSTMT1 & "_1 on " & ARTSTMT1 & " (CUST_CODE)")

            ASCMAIN1.sql = "Select * from ARTSTMT2 where ROWNUM < 1"
            ARTSTMT2 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTSTMT2 _
            & " Add Primary Key (OPS_YYYYPP, CUST_CODE, TYPE_SEQ, INV_TYPE, INV_NUM, INV_LNO)")
        End If

        ASCDATA1.ExecuteSQL("Truncate Table " & ARTSTMT1)

        ASCMAIN1.sql = "Insert into " & ARTSTMT1 & " Select X.*" _
        & ", ARTCUST6.CUST_HIGH_BAL_DATE, ARTCUST6.CUST_HIGH_BAL_AMT" _
        & ", ARTCUST1.CUST_B2C_IND" _
        & ", ARTCUST1.CUST_STMT_IND" _
        & ", ARTCUST1.NO_FINANCE_CHARGE" _
        & ", ARTCUST3.FRT_CONT_NO, ARTCUST1.FUEL_SURCHARGE_CODE" _
        & ", NULL STMT_SCAN_LINE, NULL STMT_NO" _
        & ", ARTCUST3.FRT_CONT_MONTHLY_LOCK_AMT" _
        & ", ARTCUST1.FUEL_SURCHARGE_EXEMPTION" _
        & ", ARTCUST1.TERM_CODE" _
        & ", DECODE (ARTCUST1.CUST_STMT_IND,'M','Y',ARTCUST1.CUST_STMT_IND) MAIL_STATEMENT, ARTCUST1.SREP_CODE, 0 TOTAL_CLOSED, 0 DAY_DOLLARS_CLOSED" _
        & " from ARTCUST1," & ARTCUST3 & " ARTCUST3,ARTCUST6,(" & sql & ") X " _
        & " where X.CUST_CODE = ARTCUST1.CUST_CODE " _
        & " and ARTCUST3.CUST_CODE (+) = X.CUST_CODE" _
        & " and ARTCUST6.CUST_CODE (+) = X.CUST_CODE"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
        Call ASCMAIN1.AnalyzeTable(ARTSTMT1)
        ASCDATA1.ExecuteSQL("Begin Declare Cursor C1 is Select * from " & ARTSTMT1 & " order by CUST_CODE; Begin for R1 in C1 Loop Update " & ARTSTMT1 & " set STMT_NO = TAPCTLN1('ARTSTMT1.STMT_NO',10,1) where CUST_CODE = R1.CUST_CODE; End Loop; End; End;")

        ASCMAIN1.sql = "Update " & ARTSTMT1 _
        & "  SET AGE_1 = NVL(AGE_1,0)     + NVL(TYP_ECP,0) + NVL(TYP_FRT,0) + NVL(TYP_FSC,0) + NVL(TYP_FIN,0) " _
        & ", TOTAL_DUE = NVL(TOTAL_DUE,0) + NVL(TYP_ECP,0) + NVL(TYP_FRT,0) + NVL(TYP_FSC,0) + NVL(TYP_FIN,0) "
        ASCDATA1.ExecuteSQL()



        ' Create Statement Body (Details)

        ASCDATA1.ExecuteSQL("Truncate Table " & ARTSTMT2)
        Call Create_ARTSTMT2()
        Call ASCMAIN1.AnalyzeTable(ARTSTMT2)

        'ASCMAIN1.sql = "Update " & ARTSTMT1 _
        '& " Set MAIL_STATEMENT = 'P' " _
        '& " where CUST_CODE in " _
        '& "(SELECT CUST_CODE FROM " & ARTSTMT2 _
        '& " GROUP BY CUST_CODE HAVING COUNT (*) > 1000)"
        'ASCDATA1.ExecuteSQL()

        If skip_ABS_Calcs Then
            'ASCMAIN1.sql = "Update " & ARTSTMT1 & " ARTSTMT1 " _
            '& " Set TYP_ECP = (Select Sum (INV_TOTAL_AMOUNT) " _
            '& " from " & ARTSTMT2 _
            '& " where CUST_CODE = ARTSTMT1.CUST_CODE and ORDR_TYPE_CODE = 'ECP')"
            'ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Update " & ARTSTMT1 & " ARTSTMT1 " _
            '& " Set TYP_FRT = (Select Sum (INV_TOTAL_AMOUNT) " _
            '& " from " & ARTSTMT2 _
            '& " where CUST_CODE = ARTSTMT1.CUST_CODE and ORDR_TYPE_CODE = 'FRT')"
            'ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Update " & ARTSTMT1 & " ARTSTMT1 " _
            '& " Set TYP_FSC = (Select Sum (INV_TOTAL_AMOUNT) " _
            '& " from " & ARTSTMT2 _
            '& " where CUST_CODE = ARTSTMT1.CUST_CODE and ORDR_TYPE_CODE = 'FSC')"
            'ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Update " & ARTSTMT1 & " ARTSTMT1 " _
            '& " Set TYP_FIN = (Select Sum (INV_TOTAL_AMOUNT) " _
            '& " from " & ARTSTMT2 _
            '& " where CUST_CODE = ARTSTMT1.CUST_CODE and ORDR_TYPE_CODE = 'FIN')"
            'ASCDATA1.ExecuteSQL()

            ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS " _
            & " Select CUST_CODE " _
            & " , Sum (DECODE(ORDR_TYPE_CODE,'ECP',INV_TOTAL_AMOUNT,0)) ECP " _
            & " , Sum (DECODE(ORDR_TYPE_CODE,'FRT',INV_TOTAL_AMOUNT,0)) FRT " _
            & " , Sum (DECODE(ORDR_TYPE_CODE,'FSC',INV_TOTAL_AMOUNT,0)) FSC " _
            & " , Sum (DECODE(ORDR_TYPE_CODE,'FIN',INV_TOTAL_AMOUNT,0)) FIN " _
            & " from " & ARTSTMT2 _
            & " group by CUST_CODE; " _
            & " BEGIN FOR R1 IN C1 LOOP" _
            & " Update " & ARTSTMT1 & " Set TYP_ECP = R1.ECP, TYP_FRT = R1.FRT" _
            & " , TYP_FSC = R1.FSC, TYP_FIN = R1.FIN" _
            & " where OPS_YYYYPP = '" & RYP & "' and CUST_CODE = R1.CUST_CODE;" _
            & " END LOOP; END; END; "
            ASCDATA1.ExecuteSQL()
        End If

    End Sub

    Sub Create_ARTSTMT2()

        Call ASCMAIN1.Progress("Now Creating Statement Details")

        ' Beginning Balance

        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select '" & RYP & "' OPS_YYYYPP, ARTSTMT1.CUST_CODE" _
        & ", 0 TYPE_SEQ, '0' INV_TYPE" _
        & ", '0000000000' INV_NUM, 0 INV_LNO" _
        & ", '" & Format(LYP_DATE, "dd-MMM-yyyy") & "' INV_DATE" _
        & ", Null INV_CUST_PO, NULL ORDR_NO" _
        & ", ARTSTMT1.BALFWD INV_TOTAL_AMOUNT" _
        & ", 'Last Statement Balance' REFERENCE" _
        & ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
        & ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        & ", Null CUST_CODE_B2C" _
        & ", Null ORDR_TYPE_CODE" _
        & ", Null INV_REF" _
        & ", Null DIVISION_CODE" _
        & " from " & ARTSTMT1 & " ARTSTMT1"
        ASCDATA1.ExecuteSQL()

        ' Payments

        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, 1 TYPE_SEQ, 'P' INV_TYPE" _
        & ", ARTPYMT2.PYMT_BATCH_NO INV_NUM, ARTPYMT2.PYMT_BATCH_LNO INV_LNO" _
        & ", ARTPYMT1.PYMT_BATCH_DATE INV_DATE" _
        & ", DECODE(ARTPYMT2.CUST_CREDIT_CARD_TYPE,NULL,'',ARTPYMT2.CUST_CREDIT_CARD_TYPE || ' ') || ARTPYMT2.CUST_PYMT_REF_NO INV_CUST_PO, NULL ORDR_NO" _
        & ", -1 * ARTPYMT2.CUST_PYMT_AMT INV_TOTAL_AMOUNT" _
        & ", 'Pymt Received, Thank You' REFERENCE" _
        & ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
        & ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        & ", Null CUST_CODE_B2C" _
        & ", Null ORDR_TYPE_CODE" _
        & ", Null INV_REF" _
        & ", Null DIVISION_CODE" _
        & " from ARTPYMT1, ARTPYMT2 " _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " _
        & "   and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & "   and ARTPYMT2.CUST_PYMT_AMT <> 0" _
        & "   and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & "   and ARTPYMT2.CUST_CODE is Not Null"
        ASCDATA1.ExecuteSQL()

        ' Adjustments in the form of Discounts and Allowances taken against specific invoices (ARTPYMT3)

        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, 2 TYPE_SEQ, 'A' INV_TYPE" _
        & ", 'PMT ADJUST' INV_NUM, 0 INV_LNO" _
        & ", '" & Format(RYP_DATE, "dd-MMM-yyyy") & "' INV_DATE" _
        & ", Null INV_CUST_PO, Null ORDR_NO" _
        & ", Sum (-1 * NVL(ARTPYMT3.INV_DISC_TAKEN,0) - NVL(ARTPYMT3.INV_WRITE_OFF,0)) INV_TOTAL_AMOUNT" _
        & ", 'Discounts & Allowances' REFERENCE" _
        & ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
        & ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        & ", Null CUST_CODE_B2C" _
        & ", Null ORDR_TYPE_CODE" _
        & ", Null INV_REF" _
        & ", Null DIVISION_CODE" _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT3 " _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " _
        & "   and ARTPYMT3.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT2.CUST_CODE is Not Null" _
        & "   and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & "   and ARTPYMT3.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & "   and NVL(ARTPYMT3.INV_DISC_TAKEN,0) + NVL(ARTPYMT3.INV_WRITE_OFF,0) <> 0" _
        & " group by ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE"
        ASCDATA1.ExecuteSQL()

        ' Adjustments taken in the form of GL Distribution Write-Offs (ARTPYMT4)

        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, 2 TYPE_SEQ, 'A' INV_TYPE" _
        & ", 'A/R ADJUST' INV_NUM, 0 INV_LNO" _
        & ", ARTPYMT1.PYMT_BATCH_DATE INV_DATE" _
        & ", ARTPYMT4.GL_DIST_REF INV_CUST_PO, Null ORDR_NO" _
        & ", -1 * NVL(ARTPYMT4.GL_DIST_AMT,0) INV_TOTAL_AMOUNT" _
        & ", NVL(ARTPYMT4.GL_DIST_COMMENT,'A/R Adjustment') REFERENCE" _
        & ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
        & ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        & ", Null CUST_CODE_B2C" _
        & ", Null ORDR_TYPE_CODE" _
        & ", Null INV_REF" _
        & ", Null DIVISION_CODE" _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT4 " _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " _
        & "   and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT2.CUST_CODE is Not Null" _
        & "   and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & "   and NVL(ARTPYMT4.GL_DIST_AMT,0) <> 0"
        ASCDATA1.ExecuteSQL()

        ' Adjustments taken in the form of Discounts Allowed not related to specific invoices (ARTPYMT5)
        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, 2 TYPE_SEQ, 'A' INV_TYPE" _
        & ", 'A/R ADJUST' INV_NUM, ROWNUM INV_LNO" _
        & ", ARTPYMT1.PYMT_BATCH_DATE INV_DATE" _
        & ", ARTPYMT5.CUST_REFERENCE INV_CUST_PO, Null ORDR_NO" _
        & ", -1 * NVL(ARTPYMT5.GL_DIST_AMT,0) INV_TOTAL_AMOUNT" _
        & ", NVL(ARTPYMT5.GL_DIST_COMMENT,'A/R Adjustment') REFERENCE" _
        & ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
        & ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        & ", Null CUST_CODE_B2C" _
        & ", Null ORDR_TYPE_CODE" _
        & ", Null INV_REF" _
        & ", Null DIVISION_CODE" _
        & " from ARTPYMT1, ARTPYMT2, ARTPYMT5 " _
        & " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " _
        & "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        & "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        & "   and ARTPYMT2.CUST_CODE is Not Null" _
        & "   and NVL(ARTPYMT2.PYMT_DELETED,'0') <> '1'" _
        & "   and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        & "   and NVL(ARTPYMT5.GL_DIST_AMT,0) <> 0" _
        & "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') <> '1'"
        ASCDATA1.ExecuteSQL()


        'ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        '& " Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, 2 TYPE_SEQ, 'A' INV_TYPE" _
        '& ", 'A/R ADJUST' INV_NUM, 0 INV_LNO" _
        '& ", '" & Format(RYP_DATE, "dd-MMM-yyyy") & "' INV_DATE" _
        '& ", ARTPYMT4.GL_DIST_REF INV_CUST_PO, Null ORDR_NO" _
        '& ", Sum (-1 * NVL(ARTPYMT4.GL_DIST_AMT,0)) INV_TOTAL_AMOUNT" _
        '& ", NVL(ARTPYMT4.GL_DIST_COMMENT,'A/R Adjustment') REFERENCE" _
        '& ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
        '& ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        '& ", Null CUST_CODE_B2C" _
        '& ", Null ORDR_TYPE_CODE" _
        '& ", Null INV_REF" _
        '& ", Null DIVISION_CODE" _
        '& " from ARTPYMT1, ARTPYMT2, ARTPYMT4 " _
        '& " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " _
        '& "   and ARTPYMT4.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        '& "   and ARTPYMT4.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        '& "   and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        '& "   and NVL(ARTPYMT4.GL_DIST_AMT,0) <> 0" _
        '& " group by ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE"
        'ASCDATA1.ExecuteSQL()

        '' Adjustments taken in the form of Discounts Allowed not related to specific invoices (ARTPYMT5)
        'ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        '& " Select ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE, 2 TYPE_SEQ, 'A' INV_TYPE" _
        '& ", 'A/R ADJUST' INV_NUM, 0 INV_LNO" _
        '& ", '" & Format(RYP_DATE, "dd-MMM-yyyy") & "' INV_DATE" _
        '& ", ARTPYMT5.CUST_REFERENCE INV_CUST_PO, Null ORDR_NO" _
        '& ", Sum (-1 * NVL(ARTPYMT5.GL_DIST_AMT,0)) INV_TOTAL_AMOUNT" _
        '& ", NVL(ARTPYMT5.GL_DIST_COMMENT,'A/R Adjustment') REFERENCE" _
        '& ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
        '& ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        '& ", Null CUST_CODE_B2C" _
        '& ", Null ORDR_TYPE_CODE" _
        '& ", Null INV_REF" _
        '& ", Null DIVISION_CODE" _
        '& " from ARTPYMT1, ARTPYMT2, ARTPYMT5 " _
        '& " where ARTPYMT1.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO " _
        '& "   and ARTPYMT5.PYMT_BATCH_NO = ARTPYMT2.PYMT_BATCH_NO" _
        '& "   and ARTPYMT5.PYMT_BATCH_LNO = ARTPYMT2.PYMT_BATCH_LNO" _
        '& "   and ARTPYMT1.OPS_YYYYPP = '" & RYP & "'" _
        '& "   and NVL(ARTPYMT5.GL_DIST_AMT,0) <> 0" _
        '& "   and NVL(ARTPYMT5.CHARGEBACK_IND,'0') <> '1'" _
        '& " group by ARTPYMT1.OPS_YYYYPP, ARTPYMT2.CUST_CODE"
        'ASCDATA1.ExecuteSQL()

        ' Invoices (really anything posted into SOTINVH1)

        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.CUST_CODE" _
        & ", DECODE(SOTINVH1.INV_TYPE,'I',5,'C',4,5) TYPE_SEQ, SOTINVH1.INV_TYPE" _
        & ", SOTINVH1.INV_NO INV_NUM, 0 INV_LNO" _
        & ", SOTINVH1.INV_DATE" _
        & ", NVL(SOTINVH1.ORDR_CUST_PO,DECODE(SOTORDR1.ORDR_DPD,'1',SOTORDR5.CUST_NAME,NULL)) INV_CUST_PO, SOTINVH1.ORDR_NO" _
        & ", DECODE(SOTINVH1.ORDR_TYPE_CODE,'TOP',SOTINVH1.INV_TOTAL_AMOUNT,DECODE(SOTINVH1.POST_CODE,'B2C',SOTINVH1.INV_TOTAL_AMOUNT_B2C,SOTINVH1.INV_TOTAL_AMOUNT)) INV_TOTAL_AMOUNT" _
        & ", NVL(SOTTYPE1.ORDR_TYPE_DESC,'Sale') || DECODE(SOTINVH1.CUST_SHIP_TO_NO,NULL,'',' Location: ' || SOTINVH1.CUST_SHIP_TO_NO,NULL) REFERENCE" _
        & ", SOTINVH1.ORDR_NO_WEB, SOTINVH1.CUST_SHIP_TO_NO" _
        & ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
        & ", Null CUST_CODE_B2C" _
        & ", SOTINVH1.ORDR_TYPE_CODE" _
        & ", SOTINVH1.INV_REF" _
        & ", SOTINVH1.DIVISION_CODE" _
        & " from SOTINVH1,SOTTYPE1,SOTORDR1,SOTORDR5 " _
        & " where SOTINVH1.ORDR_YYYYPP_UPDATED = '" & RYP & "'" _
        & " and SOTTYPE1.ORDR_TYPE_CODE (+) = SOTINVH1.ORDR_TYPE_CODE" _
        & " and SOTORDR1.ORDR_NO (+) = SOTINVH1.ORDR_NO" _
        & " and SOTORDR5.ORDR_NO (+) = SOTINVH1.ORDR_NO" _
        & " and SOTORDR5.CUST_ADDR_TYPE (+) = 'ST'" _
        & " and SOTINVH1.INV_TOTAL_AMOUNT <> 0"
        ASCDATA1.ExecuteSQL()



        ' Change Reference to Patient Name for DEL Invoices and Credits

        ASCMAIN1.sql = "BEGIN DECLARE " _
        & " PATIENT_NAME_JOB VARCHAR2(60);" _
        & " CURSOR C1 IS SELECT * FROM " & ARTSTMT2 _
        & " WHERE OPS_YYYYPP = '" & RYP & "' AND ORDR_TYPE_CODE = 'DEL' FOR UPDATE;" _
        & " BEGIN FOR R1 IN C1 LOOP" _
        & " PATIENT_NAME_JOB := '?';" _
        & " IF R1.INV_TYPE = 'I' THEN" _
        & " SELECT PATIENT_NAME INTO PATIENT_NAME_JOB FROM DETJOBM1 WHERE INV_NO = R1.INV_NUM;" _
        & " END IF;" _
        & " IF R1.INV_TYPE = 'C' THEN" _
        & " SELECT DETJOBM1.PATIENT_NAME INTO PATIENT_NAME_JOB" _
        & " FROM DETJOBC1,DETJOBM1 WHERE DETJOBC1.INV_NO = R1.INV_NUM AND DETJOBM1.JOB_NO = DETJOBC1.JOB_NO;" _
        & " END IF;" _
        & " IF NVL(PATIENT_NAME_JOB,'?') <> '?' THEN" _
        & " UPDATE " & ARTSTMT2 & " SET REFERENCE = PATIENT_NAME_JOB WHERE CURRENT OF C1;" _
        & " END IF;" _
        & " END LOOP; END; END;"
        ASCDATA1.ExecuteSQL()



        ' Newly generated items (FIN,ECP,FRT,FSC) not yet posted to SOTINVH1

        For i As Integer = 1 To 4
            ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
            & " Select SOTINVH1.ORDR_YYYYPP_UPDATED OPS_YYYYPP, SOTINVH1.CUST_CODE" _
            & ", " & CStr(5 + i) & " TYPE_SEQ, SOTINVH1.INV_TYPE" _
            & ", SOTINVH1.INV_NO INV_NUM, 0 INV_LNO" _
            & ", SOTINVH1.INV_DATE" _
            & ", SOTINVH1.ORDR_CUST_PO INV_CUST_PO, SOTINVH1.ORDR_NO" _
            & ", SOTINVH1.INV_TOTAL_AMOUNT" _
            & ", '" & New String() {"PrimaryECP Profit from ", "", "Fuel Surcharge", "Finance Charge"}(i - 1) & "'" _
            & IIf(i = 1, " || SOTINVH1.ORDR_CUST_PO", "") _
            & " REFERENCE" _
            & ", Null ORDR_NO_WEB, Null CUST_SHIP_TO_NO" _
            & ", Null INV_TOTAL_AMOUNT_B2B, Null INV_TOTAL_AMOUNT_B2C" _
            & ", Null CUST_CODE_B2C" _
            & ", SOTINVH1.ORDR_TYPE_CODE" _
            & ", Null INV_REF" _
            & ", Null DIVISION_CODE" _
            & " from " & New String() {SOTINVH1_ECP, SOTINVH1_FRT, SOTINVH1_FSC, SOTINVH1_FIN}(i - 1) & " SOTINVH1" _
            & " where SOTINVH1.ORDR_YYYYPP_UPDATED = '" & RYP & "'"
            ASCDATA1.ExecuteSQL()
        Next

        ' B2C Settled

        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select '" & RYP & "' OPS_YYYYPP, ARTCUST1_B2C.CUST_CODE" _
        & ", 10 TYPE_SEQ, ARTOPEN1.INV_TYPE" _
        & ", ARTOPEN1.INV_NUM, ROWNUM INV_LNO" _
        & ", ARTOPEN1.INV_DATE" _
        & ", ARTOPEN1.INV_CUST_PO, ARTOPEN1.ORDR_NO" _
        & ", NVL(ARTOPEN1.INV_PMT,0) - NVL(ARTOPEN1.INV_TOTAL_AMOUNT,0) INV_TOTAL_AMOUNT" _
        & ", 'Settled ' || TO_CHAR(ARTOPEN1.INV_LAST_PMT,'MM/DD/YYYY')" _
        & ", ARTOPEN1.ORDR_NO_WEB, ARTOPEN1.CUST_SHIP_TO_NO" _
        & ", ARTOPEN1.INV_TOTAL_AMOUNT - ARTOPEN1.INV_PROFIT_B2C INV_TOTAL_AMOUNT_B2B" _
        & ", ARTOPEN1.INV_TOTAL_AMOUNT INV_TOTAL_AMOUNT_B2C" _
        & ", ARTOPEN1.CUST_CODE CUST_CODE_B2C" _
        & ", Null ORDR_TYPE_CODE" _
        & ", Null INV_REF" _
        & ", Null DIVISION_CODE" _
        & " from ARTOPEN1, " & ARTCUST1_B2C & " ARTCUST1_B2C" _
        & " where ARTOPEN1.CUST_CODE = ARTCUST1_B2C.B2C_CUST_CODE" _
        & "   and NVL(ARTOPEN1.INV_BALANCE,0) = 0" _
        & "   and NVL(ARTOPEN1.OPS_YYYYPP,'000000') <= '" & ASCMAIN1.CYP & "'" _
        & "   and ARTOPEN1.ORDR_NO_WEB is Not Null"
        ASCDATA1.ExecuteSQL()

        ' B2C Unsettled

        ASCMAIN1.sql = "Insert into " & ARTSTMT2 _
        & " Select '" & RYP & "' OPS_YYYYPP, ARTCUST1_B2C.CUST_CODE" _
        & ", 11 TYPE_SEQ, ARTOPEN1.INV_TYPE" _
        & ", ARTOPEN1.INV_NUM, 0 INV_LNO" _
        & ", ARTOPEN1.INV_DATE" _
        & ", ARTOPEN1.INV_CUST_PO, ARTOPEN1.ORDR_NO" _
        & ", Null INV_TOTAL_AMOUNT" _
        & ", 'Unsettled' REFERENCE" _
        & ", ARTOPEN1.ORDR_NO_WEB, ARTOPEN1.CUST_SHIP_TO_NO" _
        & ", ARTOPEN1.INV_TOTAL_AMOUNT - ARTOPEN1.INV_PROFIT_B2C INV_TOTAL_AMOUNT_B2B" _
        & ", ARTOPEN1.INV_TOTAL_AMOUNT INV_TOTAL_AMOUNT_B2C" _
        & ", ARTOPEN1.CUST_CODE CUST_CODE_B2C" _
        & ", Null ORDR_TYPE_CODE" _
        & ", Null INV_REF" _
        & ", Null DIVISION_CODE" _
        & " from ARTOPEN1, " & ARTCUST1_B2C & " ARTCUST1_B2C" _
        & " where ARTOPEN1.CUST_CODE = ARTCUST1_B2C.B2C_CUST_CODE" _
        & "   and NVL(ARTOPEN1.INV_BALANCE,0) <> 0" _
        & "   and NVL(ARTOPEN1.OPS_YYYYPP,'000000') <= '" & ASCMAIN1.CYP & "'"

        ASCDATA1.ExecuteSQL()

    End Sub

    Sub XML_Generate_Statement()

        Call ASCMAIN1.Progress("Now Preparing XML file")

        Dim STMT_INSERT_CODE As String = Absx1.cmbFor("STMT_INSERT_CODE").Value & ""
        Dim STMT_INSERT_CODE2 As String = Absx1.cmbFor("STMT_INSERT_CODE2").Value & ""
        Dim STMT_INSERT_CODE3 As String = Absx1.cmbFor("STMT_INSERT_CODE3").Value & ""
        Dim STMT_MESSAGE_CODE As String = Absx1.cmbFor("STMT_MESSAGE_CODE").Value & ""

        Dim docType(9) As String
        Dim s(9) As String

        s(0) = "priorBal" : docType(0) = ""
        s(1) = "payments" : docType(1) = "Payment"
        s(2) = "adjustments" : docType(2) = "Adjustments"
        s(3) = "credits" : docType(3) = "Credit"
        s(4) = "returns" : docType(4) = "Credit" ' "Return"
        s(5) = "invoices" : docType(5) = "Invoice"
        s(6) = "b2c" : docType(6) = "PrimaryECP.com"
        s(7) = "freight" : docType(7) = "" ' "Monthly Freight"
        s(8) = "fuel" : docType(8) = "Invoice" ' "Fuel Surcharge"
        s(9) = "finance" : docType(9) = "Finance Charge"

        Dim xmlFilename As String = RYP & ".XML"
        'Dim xmlWriter As New XmlTextWriter(ASCMAIN1.Folders("Temp") & xmlFilename, System.Text.Encoding.UTF8)
        Dim xmlWriter As New XmlTextWriter(ASCMAIN1.Folders("Temp") & xmlFilename, New System.Text.UTF8Encoding(False))

        xmlWriter.IndentChar = vbTab
        xmlWriter.Indentation = 1
        xmlWriter.Formatting = Formatting.Indented

        xmlWriter.WriteStartElement("statements")
        xmlWriter.WriteAttributeString("dateCreated", (Now + ASCMAIN1.NowTSD).ToString("yyyyMMdd"))
        xmlWriter.WriteAttributeString("dateOfStatement", RYP_DATE.ToString("yyyyMMdd"))
        xmlWriter.WriteAttributeString("dateDue", RYP_DATE_DUE.ToString("yyyyMMdd"))

        Me.XML_Headings(xmlWriter)


        xmlWriter.WriteStartElement("remitTo")
        xmlWriter.WriteStartElement("nameAndAddress")
        xmlWriter.WriteAttributeString("name", ROWs("ARTPARM1").Item("AR_PARM_REMIT_NAME") & "")
        xmlWriter.WriteAttributeString("addr1", ROWs("ARTPARM1").Item("AR_PARM_REMIT_ADDR1") & "")
        xmlWriter.WriteAttributeString("addr2", ROWs("ARTPARM1").Item("AR_PARM_REMIT_ADDR2") & "")
        xmlWriter.WriteAttributeString("addr3", ROWs("ARTPARM1").Item("AR_PARM_REMIT_ADDR3") & "")
        xmlWriter.WriteAttributeString("city", ROWs("ARTPARM1").Item("AR_PARM_REMIT_CITY") & "")
        xmlWriter.WriteAttributeString("state", ROWs("ARTPARM1").Item("AR_PARM_REMIT_STATE") & "")
        xmlWriter.WriteAttributeString("zip", ROWs("ARTPARM1").Item("AR_PARM_REMIT_ZIP_CODE") & "")
        xmlWriter.WriteAttributeString("postNet", Replace(ROWs("ARTPARM1").Item("AR_PARM_REMIT_ZIP_CODE") & "", "-", "") _
                                                 & ROWs("ARTPARM1").Item("AR_PARM_REMIT_DELIVERY_POINT") _
                                                 & ROWs("ARTPARM1").Item("AR_PARM_REMIT_CHECK_DIGIT") & "")
        xmlWriter.WriteEndElement() ' nameAndAddress
        xmlWriter.WriteEndElement() ' remitTo

        Dim STMT_COUNT As Integer = 0

        Dim sql As String = ""
        'sql = "CUST_CODE = '011230'"
        If sql <> "" Then Stop ' need all customers in production
        For Each rowARTSTMT1 As DataRow In dst.Tables("ARTSTMT1") _
        .Select(sql, "CUST_CODE")
            ' ("CUST_CODE = '010405' OR CUST_CODE = '010004'", "CUST_CODE")
            STMT_COUNT += 1
            'If STMT_COUNT > 10 Then Exit For

            xmlWriter.WriteStartElement("statement")
            xmlWriter.WriteAttributeString("documentNo", rowARTSTMT1.Item("STMT_NO"))
            xmlWriter.WriteAttributeString("totalAmountDue", Val(rowARTSTMT1.Item("TOTAL_DUE") & "").ToString("#,##0.00"))
            xmlWriter.WriteAttributeString("stmtScanLine", rowARTSTMT1.Item("STMT_SCAN_LINE"))

            Dim MAIL_STATEMENT As String = rowARTSTMT1.Item("MAIL_STATEMENT")

            xmlWriter.WriteAttributeString("mailStatement", MAIL_STATEMENT)
            ' up to 3 inserts may be defined in PROCESS.JOB, named with tokens insertCode1/2/3
            ' each of these inserts may then be included or not on a per Statement basis using Y/N
            xmlWriter.WriteAttributeString("insertCode1", IIf(MAIL_STATEMENT = "N" Or STMT_INSERT_CODE = "", "N", "Y"))
            xmlWriter.WriteAttributeString("insertCode2", IIf(MAIL_STATEMENT = "N" Or STMT_INSERT_CODE2 = "", "N", "Y"))
            xmlWriter.WriteAttributeString("insertCode3", IIf(MAIL_STATEMENT = "N" Or STMT_INSERT_CODE3 = "", "N", "Y"))
            xmlWriter.WriteAttributeString("messageID", STMT_MESSAGE_CODE) ' Although we are not doing it here, the Message ID can be set to a different (valid) value for each Statement

            xmlWriter.WriteStartElement("customer")
            Dim CUST_CODE As String = rowARTSTMT1.Item("CUST_CODE")
            xmlWriter.WriteAttributeString("customerNo", CUST_CODE)

            xmlWriter.WriteStartElement("nameAndAddress")

            Dim rowARTCUST1 As DataRow = dst.Tables("ARTCUST1").Rows.Find(CUST_CODE)
            xmlWriter.WriteAttributeString("name", rowARTCUST1.Item("CUST_NAME") & "")
            xmlWriter.WriteAttributeString("addr1", rowARTCUST1.Item("CUST_ADDR1") & "")
            xmlWriter.WriteAttributeString("addr2", rowARTCUST1.Item("CUST_ADDR2") & "")
            xmlWriter.WriteAttributeString("addr3", rowARTCUST1.Item("CUST_ADDR3") & "")
            xmlWriter.WriteAttributeString("city", rowARTCUST1.Item("CUST_CITY") & "")
            xmlWriter.WriteAttributeString("state", rowARTCUST1.Item("CUST_STATE") & "")
            xmlWriter.WriteAttributeString("zip", rowARTCUST1.Item("CUST_ZIP_CODE") & "")
            xmlWriter.WriteEndElement() ' nameAndAddress

            xmlWriter.WriteEndElement() ' customer

            xmlWriter.WriteStartElement("stmtBody")
            xmlWriter.WriteAttributeString("headingType", "A")

            xmlWriter.WriteAttributeString("headingLeft", "Statement for " & CUST_CODE & ":" & rowARTCUST1.Item("CUST_NAME"))

            Dim SREP_CODE As String = rowARTCUST1.Item("SREP_CODE") & ""
            Dim SREP_NAME_in_heading As String = ""
            Dim rowSOTSREP1 As DataRow = dst.Tables("SOTSREP1").Rows.Find(SREP_CODE)
            If rowSOTSREP1 IsNot Nothing Then
                SREP_NAME_in_heading = "Sales Rep " & rowSOTSREP1.Item("SREP_NAME") & ""
            End If
            xmlWriter.WriteAttributeString("headingRight", SREP_NAME_in_heading)


            Dim DEL_SECTION As Boolean = False
            Dim DEL_SECTION_TOTAL As Double = 0
            Dim DEL_SQL As String = ""
            Dim custDELheading As Boolean = False
            Dim custHasDEL As Boolean = (rowARTSTMT1.Item("DEL") = "1") ' False

            Dim custHasLOC As Boolean = (rowARTSTMT1.Item("LOC") = "1") ' False

            For i As Integer = 0 To docType.GetUpperBound(0)
                Dim blank_line_printed As Boolean = False

                DEL_SQL = ""
                If custHasDEL Then
                    If i = 6 Then
                        If Not DEL_SECTION Then
                            DEL_SECTION = True
                            i = 3
                        End If
                    End If

                    If i = 3 Then
                        DEL_SECTION_TOTAL = 0
                        custDELheading = False
                    End If

                    If i = 3 Or i = 4 Or i = 5 Then
                        If DEL_SECTION Then
                            DEL_SQL = " AND DEL = '1'"
                        Else
                            DEL_SQL = " AND DEL = '0'"
                        End If
                    End If
                End If

                Dim CUST_SHIP_TO_NO As String = "1234567"
                Dim CUST_SHIP_TO_NO_shown As String = ""
                Dim LOC_SECTION_TOTAL As Double = 0

                Dim AR_ITEM_SEQ As String = "INV_DATE, INV_NUM"
                If (i = 3 Or i = 4 Or i = 5) Then
                    AR_ITEM_SEQ = "CUST_SHIP_TO_NO, INV_DATE, INV_NUM"
                End If

                For Each rowARTSTMT2 As DataRow In _
                dst.Tables("ARTSTMT2").Select( _
                "OPS_YYYYPP = '" & RYP & "'" _
                & " and CUST_CODE = '" & CUST_CODE & "'" _
                & " and TYPE_SEQ = " & CStr(i) & DEL_SQL, AR_ITEM_SEQ)
                    If Not blank_line_printed Then
                        Me.XML_Blank_Line(xmlWriter, "A")
                        blank_line_printed = True

                        If custHasDEL Then
                            If (i = 3 Or i = 4 Or i = 5) And Not custDELheading Then
                                custDELheading = True
                                xmlWriter.WriteStartElement("stmtBodyDetail")
                                Dim COMPANY As String = ""
                                If DEL_SECTION Then
                                    COMPANY = "Digital EyeLab:"
                                Else
                                    COMPANY = "Optical Distributor Group:"
                                End If
                                xmlWriter.WriteAttributeString("docDate", COMPANY)
                                xmlWriter.WriteAttributeString("docType", "")
                                xmlWriter.WriteAttributeString("docNo", "")
                                xmlWriter.WriteAttributeString("docOrderNo", "")
                                xmlWriter.WriteAttributeString("docPO", "")
                                xmlWriter.WriteAttributeString("docRef", "")
                                xmlWriter.WriteAttributeString("docAmt", "")
                                xmlWriter.WriteEndElement() ' stmtBodyDetail

                                Me.XML_Blank_Line(xmlWriter, "A")
                            End If
                        End If
                    End If

                    If custHasLOC Then
                        If (i = 3 Or i = 4 Or i = 5) Then
                            If CUST_SHIP_TO_NO <> rowARTSTMT2.Item("CUST_SHIP_TO_NO") & "" Then
                                If CUST_SHIP_TO_NO <> "1234567" Then
                                    ' warning - this code also appears below
                                    xmlWriter.WriteStartElement("stmtBodyDetail")
                                    xmlWriter.WriteAttributeString("docDate", "Location:  " & CUST_SHIP_TO_NO_shown)
                                    xmlWriter.WriteAttributeString("docType", "")
                                    xmlWriter.WriteAttributeString("docNo", "")
                                    xmlWriter.WriteAttributeString("docOrderNo", "")
                                    xmlWriter.WriteAttributeString("docPO", "")
                                    xmlWriter.WriteAttributeString("docRef", "Sub-Total")
                                    xmlWriter.WriteAttributeString("docAmt", Format(LOC_SECTION_TOTAL, "#,##0.00;#,##0.00CR"))
                                    xmlWriter.WriteEndElement() ' stmtBodyDetail
                                    Me.XML_Blank_Line(xmlWriter, "A")
                                End If

                                CUST_SHIP_TO_NO = rowARTSTMT2.Item("CUST_SHIP_TO_NO") & ""
                                CUST_SHIP_TO_NO_shown = CUST_SHIP_TO_NO
                                LOC_SECTION_TOTAL = 0

                                xmlWriter.WriteStartElement("stmtBodyDetail")
                                Dim LOCATION As String = ""
                                Dim LOCATION_ADDRESS As String = ""
                                If CUST_SHIP_TO_NO = "" Then
                                    CUST_SHIP_TO_NO_shown = "000000"
                                    LOCATION_ADDRESS = rowARTCUST1.Item("CUST_NAME") & ", " & rowARTCUST1.Item("CUST_CITY") & ", " & rowARTCUST1.Item("CUST_STATE") & " " & rowARTCUST1.Item("CUST_ZIP_CODE")
                                Else
                                    Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_SHIP_TO_NO})
                                    If rowARTCUST2 IsNot Nothing Then
                                        LOCATION_ADDRESS = rowARTCUST2.Item("CUST_SHIP_TO_NAME") & ", " & rowARTCUST2.Item("CUST_SHIP_TO_CITY") & ", " & rowARTCUST2.Item("CUST_SHIP_TO_STATE") & " " & rowARTCUST2.Item("CUST_SHIP_TO_ZIP_CODE")
                                    End If
                                End If
                                LOCATION = "Location:  " & CUST_SHIP_TO_NO_shown & " " & LOCATION_ADDRESS
                                xmlWriter.WriteAttributeString("docDate", LOCATION)
                                xmlWriter.WriteAttributeString("docType", "")
                                xmlWriter.WriteAttributeString("docNo", "")
                                xmlWriter.WriteAttributeString("docOrderNo", "")
                                xmlWriter.WriteAttributeString("docPO", "")
                                xmlWriter.WriteAttributeString("docRef", "")
                                xmlWriter.WriteAttributeString("docAmt", "")
                                xmlWriter.WriteEndElement() ' stmtBodyDetail
                            End If
                        End If
                    End If

                    xmlWriter.WriteStartElement("stmtBodyDetail")
                    xmlWriter.WriteAttributeString("docDate", Format(rowARTSTMT2.Item("INV_DATE"), "MM/dd/yy"))
                    xmlWriter.WriteAttributeString("docType", docType(i))
                    If Val(rowARTSTMT2.Item("TYPE_SEQ") & "") = 0 Then
                        xmlWriter.WriteAttributeString("docNo", "")
                    Else
                        Dim INV_NUM As String = rowARTSTMT2.Item("INV_NUM") & ""
                        If skip_ABS_Calcs Then
                            If INV_NUM.Length = 10 Then
                                If INV_NUM.Substring(0, 3) = "000" Then
                                    INV_NUM = INV_NUM.Substring(4)
                                ElseIf INV_NUM.Substring(0, 3) = "999" Then
                                    Dim INV_REF As String = rowARTSTMT2.Item("INV_REF") & ""
                                    INV_NUM = INV_REF
                                    If INV_NUM.Substring(0, 3) = "000" Then
                                        INV_NUM = INV_NUM.Substring(4)
                                    End If
                                End If
                            End If
                        End If
                        xmlWriter.WriteAttributeString("docNo", INV_NUM)
                    End If

                    Dim ORDR_TYPE_CODE As String = rowARTSTMT2.Item("ORDR_TYPE_CODE") & ""
                    Dim ORDR_NO_WEB As String = rowARTSTMT2.Item("ORDR_NO_WEB") & ""

                    Dim ORDR_NO As String = rowARTSTMT2.Item("ORDR_NO") & ""
                    If skip_ABS_Calcs Then
                        If ORDR_NO.Length = 10 Then
                            If ORDR_NO.Substring(0, 3) = "000" Then
                                ORDR_NO = ORDR_NO.Substring(4)
                            End If
                        End If
                    End If
                    xmlWriter.WriteAttributeString("docOrderNo", ORDR_NO)

                    xmlWriter.WriteAttributeString("docPO", rowARTSTMT2.Item("INV_CUST_PO") & "")
                    Dim REFERENCE As String = rowARTSTMT2.Item("REFERENCE") & ""
                    If ORDR_TYPE_CODE = "REG" Or REFERENCE = "Regular Order" Then
                        REFERENCE = ""

                    ElseIf ORDR_TYPE_CODE = "ECP" Or REFERENCE = "ECP Profit Credit" Then
                        REFERENCE = "PrimaryECP.com Monthly Profit"

                    ElseIf ORDR_TYPE_CODE = "FRT" Or REFERENCE = "Monthly Freight Charge" Then
                        REFERENCE = ""

                        'ElseIf ORDR_TYPE_CODE = "DEL" Then
                        '    REFERENCE = ""

                    ElseIf ORDR_TYPE_CODE = "FSC" Or REFERENCE = "Fuel Surcharge" Then
                        REFERENCE = ""

                    ElseIf ORDR_TYPE_CODE = "B2B" Or REFERENCE = "B2B from opticaldg.com" Then
                        If ORDR_NO_WEB = "" Then
                            REFERENCE = ""
                        Else
                            REFERENCE = "Web Order: " & ORDR_NO_WEB
                        End If

                    ElseIf ORDR_TYPE_CODE = "ADJ" Or REFERENCE = "Billing Adjustment" Then
                        If i = 3 Then
                            REFERENCE = "Credit Memo"
                        Else
                            REFERENCE = ""
                        End If
                    End If
                    'End If

                    Dim INV_TOTAL_AMOUNT As Decimal = Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT") & "")
                    If i = 3 Or i = 4 Or i = 5 Then
                        DEL_SECTION_TOTAL += INV_TOTAL_AMOUNT
                        LOC_SECTION_TOTAL += INV_TOTAL_AMOUNT
                    End If

                    xmlWriter.WriteAttributeString("docRef", REFERENCE)
                    xmlWriter.WriteAttributeString("docAmt", Format(INV_TOTAL_AMOUNT, "#,##0.00;#,##0.00CR"))
                    xmlWriter.WriteEndElement() ' stmtBodyDetail
                Next

                If custHasLOC Then
                    If (i = 3 Or i = 4 Or i = 5) Then
                        ' warning - this code also appears above
                        If CUST_SHIP_TO_NO <> "1234567" Then
                            xmlWriter.WriteStartElement("stmtBodyDetail")
                            xmlWriter.WriteAttributeString("docDate", "Location:  " & CUST_SHIP_TO_NO_shown)
                            xmlWriter.WriteAttributeString("docType", "")
                            xmlWriter.WriteAttributeString("docNo", "")
                            xmlWriter.WriteAttributeString("docOrderNo", "")
                            xmlWriter.WriteAttributeString("docPO", "")
                            xmlWriter.WriteAttributeString("docRef", "Sub-Total")
                            xmlWriter.WriteAttributeString("docAmt", Format(LOC_SECTION_TOTAL, "#,##0.00;#,##0.00CR"))
                            xmlWriter.WriteEndElement() ' stmtBodyDetail
                            Me.XML_Blank_Line(xmlWriter, "A")
                        End If
                    End If
                End If


                If i = 5 And custHasDEL Then
                    xmlWriter.WriteStartElement("stmtBodyDetail")
                    Dim COMPANY As String = ""
                    If DEL_SECTION Then
                        COMPANY = "Digital EyeLab:"
                    Else
                        COMPANY = "Optical Distributor Group:"
                    End If
                    xmlWriter.WriteAttributeString("docDate", COMPANY)
                    xmlWriter.WriteAttributeString("docType", "")
                    xmlWriter.WriteAttributeString("docNo", "")
                    xmlWriter.WriteAttributeString("docOrderNo", "")
                    xmlWriter.WriteAttributeString("docPO", "")
                    xmlWriter.WriteAttributeString("docRef", "Sub-Total")
                    xmlWriter.WriteAttributeString("docAmt", Format(DEL_SECTION_TOTAL, "#,##0.00;#,##0.00CR"))
                    xmlWriter.WriteEndElement() ' stmtBodyDetail
                End If
            Next

            xmlWriter.WriteStartElement("stmtBodyTotals")
            Dim FUTURE_AMT As Double = 0

            'If CUST_CODE > "800000" Then
            '    Stop ' next line for futures
            '    FUTURE_AMT = 1
            'End If

            If FUTURE_AMT = 0 Then
                xmlWriter.WriteAttributeString("headingType", "E")
            Else
                xmlWriter.WriteAttributeString("headingType", "B")
            End If





            If FUTURE_AMT = 0 Then
                xmlWriter.WriteAttributeString("total0", "")
            Else
                xmlWriter.WriteAttributeString("total0", Format(FUTURE_AMT, "#,##0.00;#,##0.00CR"))
            End If
            xmlWriter.WriteAttributeString("total1", Format(Val(rowARTSTMT1.Item("AGE_1") & ""), "#,##0.00;#,##0.00CR"))
            xmlWriter.WriteAttributeString("total2", Format(Val(rowARTSTMT1.Item("AGE_2") & ""), "#,##0.00;#,##0.00CR"))
            xmlWriter.WriteAttributeString("total3", Format(Val(rowARTSTMT1.Item("AGE_3") & ""), "#,##0.00;#,##0.00CR"))
            xmlWriter.WriteAttributeString("total4", Format(Val(rowARTSTMT1.Item("AGE_4") & ""), "#,##0.00;#,##0.00CR"))
            xmlWriter.WriteAttributeString("totalAmtDue", Format(Val(rowARTSTMT1.Item("TOTAL_DUE") & ""), "#,##0.00;#,##0.00CR"))
            xmlWriter.WriteAttributeString("stmtMessage", txtSTMT_MESSAGE.Text)
            xmlWriter.WriteEndElement() ' stmtBodyTotals

            xmlWriter.WriteEndElement() ' stmtBody

            Dim stmtB2C As Boolean = False ' True if we have printed any Settled or Unsettled
            Dim CUST_CODE_B2C As String
            Dim B2C() As Double
            Dim B2C_all() As Double

            For Each sType As String In New String() {"Settled", "Unsettled"}
                Dim sTypeHeading As String = IIf(sType = "Settled", "C", "D")
                CUST_CODE_B2C = ""
                ReDim B2C_all(2)
                ReDim B2C(2)

                For Each rowARTSTMT2 As DataRow In _
                    dst.Tables("ARTSTMT2").Select( _
                    "OPS_YYYYPP = '" & RYP & "'" _
                    & " and CUST_CODE = '" & CUST_CODE & "'" _
                    & " and TYPE_SEQ = " & IIf(sType = "Settled", "10", "11"), _
                    "CUST_CODE_B2C, INV_DATE, INV_NUM")

                    If B2C_all(0) = 0 Or CUST_CODE_B2C <> rowARTSTMT2.Item("CUST_CODE_B2C") Then
                        If Not stmtB2C Then
                            stmtB2C = True
                            xmlWriter.WriteStartElement("stmtB2C")
                        End If

                        If B2C_all(0) = 0 Then
                            xmlWriter.WriteStartElement("stmt" & sType)
                            xmlWriter.WriteAttributeString("headingType", sTypeHeading)
                            xmlWriter.WriteAttributeString("headingLeft", "Statement for " & CUST_CODE & ":" & rowARTCUST1.Item("CUST_NAME"))
                            'xmlWriter.WriteAttributeString("headingRight", sType & " B2C items from Account " & CUST_CODE_B2C)
                            'xmlWriter.WriteAttributeString("headingRight", sType & " B2C items")
                            xmlWriter.WriteAttributeString("headingRight", "PrimaryECP.com Monthly Activity")
                        Else
                            XML_Write_Totals_C_B2C(xmlWriter, sType, sTypeHeading, B2C, CUST_CODE_B2C)
                            Me.XML_Blank_Line(xmlWriter, sTypeHeading)
                        End If

                        ReDim B2C(2)
                        CUST_CODE_B2C = rowARTSTMT2.Item("CUST_CODE_B2C")
                        B2C_all(0) += 1

                        Me.XML_Blank_Line(xmlWriter, sTypeHeading, sType & " B2C items from Account " & CUST_CODE_B2C)
                    End If


                    Dim INV_NUM As String = rowARTSTMT2.Item("INV_NUM") & ""
                    Dim ORDR_NO As String = rowARTSTMT2.Item("ORDR_NO") & ""
                    Dim ORDR_NO_WEB As String = rowARTSTMT2.Item("ORDR_NO_WEB") & ""

                    If skip_ABS_Calcs Then
                        If INV_NUM.Length = 10 AndAlso INV_NUM.Substring(0, 3) = "000" Then
                            INV_NUM = INV_NUM.Substring(4)
                        End If
                        If ORDR_NO.Length = 10 AndAlso ORDR_NO.Substring(0, 3) = "000" Then
                            ORDR_NO = ORDR_NO.Substring(4)
                        End If
                        If ORDR_NO_WEB.Length = 10 AndAlso ORDR_NO_WEB.Substring(0, 3) = "000" Then
                            ORDR_NO_WEB = ORDR_NO_WEB.Substring(4)
                        End If
                    End If

                    Call XML_WriteLine(xmlWriter, sTypeHeading, "", New String() { _
                    Format(rowARTSTMT2.Item("INV_DATE"), "MM/dd/yy"), _
                    sType, _
                    INV_NUM & "", _
                    ORDR_NO & "", _
                    ORDR_NO_WEB & "", _
                    Format(Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2B") & ""), "#,##0.00;#,##0.00CR"), _
                    Format(Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2C") & ""), "#,##0.00;#,##0.00CR"), _
                    Format(Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2C") & "") - Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2B") & ""), "#,##0.00;#,##0.00CR")})

                    B2C(0) += 1
                    B2C(1) += Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2B") & "")
                    B2C(2) += Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2C") & "")
                    B2C_all(1) += Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2B") & "")
                    B2C_all(2) += Val(rowARTSTMT2.Item("INV_TOTAL_AMOUNT_B2C") & "")
                Next

                If B2C_all(0) <> 0 Then
                    If B2C_all(0) > 1 Then
                        XML_Write_Totals_C_B2C(xmlWriter, sType, sTypeHeading, B2C, CUST_CODE_B2C)
                    End If
                    Me.XML_Blank_Line(xmlWriter, sTypeHeading)
                    XML_Write_Totals_C(xmlWriter, sType, sTypeHeading, B2C_all, CUST_CODE)
                End If
            Next

            If stmtB2C Then
                xmlWriter.WriteStartElement("stmtB2CMessage")
                xmlWriter.WriteAttributeString("stmtMessage", txtSTMT_MESSAGE_ECP.Text)
                xmlWriter.WriteEndElement() ' stmtB2CMessage
                xmlWriter.WriteEndElement() ' stmtB2C
            End If

            xmlWriter.WriteEndElement() ' statement
        Next
        xmlWriter.WriteEndElement() ' statements

        xmlWriter.Close()

        Dim jobFilename As String = "PROCESS.JOB"
        Using jobWriter As New System.IO.StreamWriter(ASCMAIN1.Folders("Temp") & jobFilename)
            Dim STMT_INS As String _
            = IIf(STMT_INSERT_CODE = "", "", String.Format(" INS1={0}", STMT_INSERT_CODE)) _
            & IIf(STMT_INSERT_CODE2 = "", "", String.Format(" INS2={0}", STMT_INSERT_CODE2)) _
            & IIf(STMT_INSERT_CODE3 = "", "", String.Format(" INS3={0}", STMT_INSERT_CODE3))
            ' note that up to 3 inserts are supported by OSG, 
            ' and if we use them we need to reference them like INS1=XX1 INS2=XX2 INS3=XX3
            Dim STMT_ENV As String = "" ' String.Format(" ENVMSGID={0} ENVMSGCOLOR=B", "456")
            Dim STMT_MODE As String = ""
            If Absx1.chkFor("TEST").Checked Then STMT_MODE = " MODE=Test"
            'FILE=200709.XML PROC=MAIL&PDF STMTS=8254 INS1=000001 MODE=Test
            'PROC=PDFONLY()
            'PROC=MAIL()

            Dim PROC As String = "MAIL&PDF"
            If ASCMAIN1.Running_in_VS Then
                'Stop
                'PROC = "PDFONLY"
            End If

            Dim JOB As String = _
            String.Format("FILE={0} PROC={1} STMTS={2}" & STMT_INS & STMT_ENV & STMT_MODE _
            , xmlFilename, PROC, CStr(STMT_COUNT))
            jobWriter.Write(JOB & vbCrLf)
        End Using

        If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & RYP & ".zip") Then
            My.Computer.FileSystem.DeleteFile(ASCMAIN1.Folders("Temp") & RYP & ".zip")
        End If

        Try
            Dim Zip1 As New nsoftware.IPWorksZip.Zip
            Zip1.RuntimeLicense = nSoftwarekeys("nSoftwareZipkey")
            Zip1.ArchiveFile = ASCMAIN1.Folders("Temp") & RYP & ".zip"
            Zip1.IncludeFiles(ASCMAIN1.Folders("Temp") & xmlFilename & " | " & ASCMAIN1.Folders("Temp") & jobFilename)
            Zip1.Compress()
            Zip1.Dispose()

            Using sw As New System.IO.StreamWriter(ASCMAIN1.Folders("Temp") & "osg.scr")
                sw.WriteLine("open " & ROWs("ARTPARM1").Item("AR_PARM_OSG_IP"))
                sw.WriteLine(ROWs("ARTPARM1").Item("AR_PARM_OSG_USER"))
                sw.WriteLine(ROWs("ARTPARM1").Item("AR_PARM_OSG_PWD"))
                sw.WriteLine("binary")
                sw.WriteLine("put " & ASCMAIN1.Folders("Temp") & RYP & ".zip")
                sw.WriteLine("quit")
                sw.Close()
            End Using

        Catch ex As Exception
            MsgBox("Error Creating Zip File for Transmission to OSG")
            RWU = "N"
        End Try

    End Sub

    Sub XML_Write_Totals_C_B2C( _
    ByRef XmlWriter As XmlWriter, _
    ByVal sType As String, _
    ByVal sTypeHeading As String, _
    ByVal B2C() As Double, _
    ByVal CUST_CODE_B2C As String)
        Call XML_WriteLine(XmlWriter, sTypeHeading, "stmtB2CDetail", New String() { _
        "Sub-Total", _
        sType, _
        "", _
        "from Account", _
        CUST_CODE_B2C, _
        Format(B2C(1), "#,##0.00;#,##0.00CR"), _
        Format(B2C(2), "#,##0.00;#,##0.00CR"), _
        Format(B2C(2) - B2C(1), "#,##0.00;#,##0.00CR")})
        'XmlWriter.WriteEndElement() ' stmtSettled / stmtUnsettled
    End Sub

    Sub XML_Write_Totals_C( _
    ByRef XmlWriter As XmlWriter, _
    ByVal sType As String, _
    ByVal sTypeHeading As String, _
    ByVal B2C_all() As Double, _
    ByVal CUST_CODE As String)
        Call XML_WriteLine(XmlWriter, sTypeHeading, "stmtB2CTotals", New String() { _
        "Total", _
        sType, _
        "", _
        IIf(sType = "Settled", "Profit Transfer", ""), _
        IIf(sType = "Settled", "to " & CUST_CODE, ""), _
        Format(B2C_all(1), "#,##0.00;#,##0.00CR"), _
        Format(B2C_all(2), "#,##0.00;#,##0.00CR"), _
        Format(B2C_all(2) - B2C_all(1), "#,##0.00;#,##0.00CR")})
        XmlWriter.WriteEndElement() ' stmtSettled / stmtUnsettled
    End Sub

    Sub XML_Blank_Line( _
    ByRef XmlWriter As XmlWriter, ByVal headingType As String, Optional ByVal FirstColumn As String = "")

        Dim Data(10) As String
        If FirstColumn <> "" Then
            Data(0) = FirstColumn
        End If
        Call XML_WriteLine(XmlWriter, headingType, "", Data)

    End Sub

    Sub XML_WriteLine( _
    ByRef XmlWriter As XmlWriter, _
    ByVal headingType As String, _
    ByVal XMLElement As String, _
    ByVal Data() As String)

        If XMLElement = "" Then
            XMLElement = XMLHeadingTypes(headingType)
        End If

        XmlWriter.WriteStartElement(XMLElement)
        For i As Integer = 1 To XMLHeadings(headingType).Length
            XmlWriter.WriteAttributeString(XMLHeadings(headingType)(i - 1), Data(i - 1))
        Next
        XmlWriter.WriteEndElement()
    End Sub

    Sub XML_Headings( _
    ByRef XmlWriter As XmlWriter)

        XMLHeadings.Clear()
        XMLHeadingTypes.Clear()

        XmlWriter.WriteStartElement("headings")

        ' Statement Body
        Me.XML_Heading_Type(XmlWriter, "A", "Statement Body", "N", _
        New String() {"Date", "Type", "Document", "ODG Order", "Reference", "Description", "Amount"})
        XMLHeadingTypes.Add("A", "stmtBodyDetail")
        XMLHeadings.Add("A", New String() {"docDate", "docType", "docNo", "docOrderNo", "docPO", "docRef", "docAmt"})

        ' Statement Totals
        'Me.XML_Heading_Type(XmlWriter, "B", "Statement Totals", "N", _
        'New String() {ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") & "" _
        '            , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") & "" _
        '            , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") & "" _
        '            , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4") & "" _
        '            , "Total Due"})
        'Stop ' next line done for futures
        Me.XML_Heading_Type(XmlWriter, "B", "Statement Totals", "N", _
        New String() {"Future" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") & "" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") & "" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") & "" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4") & "" _
                    , "Total Due"})
        XMLHeadingTypes.Add("B", "stmtBodyTotals")
        'Stop ' next line modified for futures
        XMLHeadings.Add("B", New String() {"total0", "total1", "total2", "total3", "total4", "totalAmtDue"})

        ' B2C - Settled
        Me.XML_Heading_Type(XmlWriter, "C", "PrimaryECP - Settled Transactions", "Y", _
        New String() {"Date", "Type", "Document", "ODG Order", "Web Order No", "Cost", "Retail", "Profit"})
        XMLHeadingTypes.Add("C", "stmtB2CDetail")
        XMLHeadings.Add("C", New String() {"docDate", "docType", "docNo", "docOrderNo", "docWebOrderNo", "docCost", "docRetail", "docProfit"})

        ' B2C - Un-Settled
        Me.XML_Heading_Type(XmlWriter, "D", "PrimaryECP - Unsettled Transactions", "Y", _
        New String() {"Date", "Type", "Document", "ODG Order", "Web Order No", "Cost", "Retail", "Profit"})
        XMLHeadingTypes.Add("D", "stmtB2CDetail")
        XMLHeadings.Add("D", New String() {"docDate", "docType", "docNo", "docOrderNo", "docWebOrderNo", "docCost", "docRetail", "docProfit"})


        Me.XML_Heading_Type(XmlWriter, "E", "Statement Totals No Futures", "N", _
        New String() {"" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_1") & "" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_2") & "" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_3") & "" _
                    , ROWs("ARTPARM1").Item("AR_PARM_AGE_CATG_DESC_4") & "" _
                    , "Total Due"})
        XMLHeadingTypes.Add("E", "stmtBodyTotals")
        'Stop ' next line modified for futures
        XMLHeadings.Add("E", New String() {"total0", "total1", "total2", "total3", "total4", "totalAmtDue"})


        XmlWriter.WriteEndElement() ' headings

    End Sub

    Sub XML_Heading_Type( _
    ByRef XmlWriter As XmlWriter, _
    ByVal headingType As String, _
    ByVal headingDesc As String, _
    ByVal newPageOrSide As String, _
    ByVal colText() As String)
        XmlWriter.WriteStartElement("heading")
        XmlWriter.WriteAttributeString("headingType", headingType)
        XmlWriter.WriteAttributeString("headingDesc", headingDesc)
        XmlWriter.WriteAttributeString("newPageOrSide", newPageOrSide)
        For i As Integer = 1 To 9
            Dim colTextX As String = ""
            If colText.Length >= i Then
                colTextX = colText(i - 1)
            End If
            XmlWriter.WriteAttributeString("colText" & CStr(i), colTextX)
        Next
        XmlWriter.WriteEndElement() ' heading
    End Sub

    Sub ftp_File()

        Call ASCMAIN1.Progress("Now ftp'ing file to " & ROWs("ARTPARM1").Item("AR_PARM_OSG_IP"))

        Ftp1.User = ROWs("ARTPARM1").Item("AR_PARM_OSG_USER")
        Ftp1.Password = ROWs("ARTPARM1").Item("AR_PARM_OSG_PWD")
        Ftp1.RemoteHost = ROWs("ARTPARM1").Item("AR_PARM_OSG_IP")
        Ftp1.Logon()
        Ftp1.TransferMode = nsoftware.IPWorks.FtpTransferModes.tmBinary
        Ftp1.LocalFile = ASCMAIN1.Folders("Temp") & RYP & ".zip"
        Ftp1.RemoteFile = RYP & ".zip"
        Ftp1.Upload()
        Ftp1.Logoff()
    End Sub


End Class