Imports System.Net
Imports System.Net.Http
Imports System.Net.Mail
Imports Microsoft.Exchange.WebServices.Data
Imports Microsoft.Identity.Client
'Imports Newtonsoft.Json
Imports System.Configuration
Imports SmartyStreets.USStreetApi
Imports SmartyStreets
Imports System.IO
Imports System.Threading.Tasks
Imports Microsoft.Office.Interop

Public Class TACMAIN1
    Inherits ABSolution.TACMAIN1

    Public Shared SREP_CODE As String = ""
    Public Shared SREP_CODE_MGR_ASST As String = ""
    Public Shared SREP_CODEs As New List(Of String)
    Dim CURR_EXCH_RATE_response As String
    'Dim WithEvents http1 As New nsoftware.IPWorks.Http
    Dim forexUpdated As Boolean = False
    Dim forexSvc As String = ""

    Public Shared IPLBMacysCustomerCodes As List(Of String) = New List(Of String)({"MACYS", "MACYSCOM", "MACYSBACK", "BLOOMIES", "BLOOMCOM", "BLOOMOUT"})

    Public Shared ews_service As ExchangeService

    Public Overrides Sub Site_Specific_Settings()

        ' Setup printers and serial ports.
        ' To avoid having to create the same code across multiple forms
        ' The variables / controls will be on ASFMAIN1 and accessible through ASCMAIN1

        If ASCMAIN1.DBS_SERVER = "ANE" Or ASCMAIN1.DBS_COMPANY = "ANE" Then Exit Sub
        If ASCMAIN1.DBS_SERVER = "EXP" Or ASCMAIN1.DBS_COMPANY = "EXP" Then Exit Sub

        'Dim sftp As New nsoftware.IPWorksSSH.Sftp
        'sftp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwaresftpkey")
        'Dim Zip1 As New nsoftware.IPWorksZip.Zip
        'Zip1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareZipkey")
        'Dim openpgp1 As New nsoftware.IPWorksEncrypt.Openpgp
        'openpgp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareEncryptionkey")
        'Dim ipp As New nsoftware.IPWorks.Ipport
        'ipp.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareipportkey")

        ASCMAIN1.sql = "Select * from TATUSER1 where USER_ID = :PARM1"
        Dim rowTATUSER1 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "V", New Object() {ASCMAIN1.USER_ID})
        If rowTATUSER1 IsNot Nothing Then
            SREP_CODE = rowTATUSER1.Item("SREP_CODE") & ""
            SREP_CODE_MGR_ASST = rowTATUSER1.Item("SREP_CODE_MGR_ASST") & ""
        End If

        If SREP_CODE <> "" Then
            SREP_CODEs.Clear()
            SREP_CODEs.Add(SREP_CODE)
            If SREP_CODE <> "" Then
                ASCMAIN1.sql = "Select SREP_CODE from SOTSREP1 where SREP_CODE_MGR = '" & SREP_CODE & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim SREP_CODE_managed As String = row.Item("SREP_CODE")
                    SREP_CODEs.Add(SREP_CODE_managed)
                Next
            End If
            If SREP_CODE_MGR_ASST <> "" Then
                ASCMAIN1.sql = "Select SREP_CODE from SOTSREP1 where SREP_CODE_MGR = '" & SREP_CODE_MGR_ASST & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                    Dim SREP_CODE_managed As String = row.Item("SREP_CODE")
                    SREP_CODEs.Add(SREP_CODE_managed)
                Next
            End If
            If rowTATUSER1.Item("SREP_CODE2") & "" <> "" Then
                SREP_CODEs.Add(rowTATUSER1.Item("SREP_CODE2"))
            End If
        End If


        Dim stationID As String = System.Environment.GetEnvironmentVariable("USERNAME") & String.Empty
        Dim sql As String = "SELECT * FROM WHTLINE1 WHERE UPPER(STATION_ID) = :PARM1"
        Dim rowWHTLINE1 As DataRow = ASCDATA1.GetDataRow(sql, "V", New Object() {stationID.ToUpper})

        'If rowWHTLINE1 Is Nothing Then
        '    rowWHTLINE1 = ASCDATA1.GetDataRow(sql, "V", New Object() {"DEFAULT"})
        'End If

        ASCMAIN1.LabelPrinterSerialPort = Nothing
        'ASCMAIN1.ScaleSerialPort = Nothing

        If rowWHTLINE1 IsNot Nothing Then
            Try
                ASCMAIN1.LaserPrinterName = (rowWHTLINE1.Item("LASER_PRT_NAME") & String.Empty).ToString.Trim

                ASCMAIN1.LabelPrinterSerialPort = New System.IO.Ports.SerialPort
                Application.DoEvents()
                ASCMAIN1.LabelPrinterSerialPort.PortName = (rowWHTLINE1.Item("LABEL_PRT_COMM_PORT") & String.Empty).ToString.Trim
                Application.DoEvents()
                If Not ASCMAIN1.Running_in_VS Then
                    ASCMAIN1.LabelPrinterSerialPort.Open()
                    ASCMAIN1.LabelPrinterSerialPort.DiscardInBuffer()
                    ASCMAIN1.LabelPrinterSerialPort.DiscardOutBuffer()
                End If
                'Clean out the buffers
            Catch ex As Exception
                ASCMAIN1.LabelPrinterSerialPort = Nothing
            End Try

            ASCMAIN1.LaserPrinterIpAddress = (rowWHTLINE1.Item("LASER_PRT_IP_ADDRESS") & String.Empty).ToString.Trim
            If Not Net.IPAddress.TryParse(ASCMAIN1.LaserPrinterIpAddress, Nothing) Then
                ASCMAIN1.LaserPrinterIpAddress = String.Empty
            End If


            'Dim SCALE_COMM_PORT As String = (rowWHTLINE1.Item("SCALE_COMM_PORT") & String.Empty).ToString.Trim
            'ASCMAIN1.ScaleSerialPort = New System.IO.Ports.SerialPort
            'Try
            '    If SCALE_COMM_PORT.Length > 0 Then
            '        Windows.Forms.Application.DoEvents()
            '        ASCMAIN1.ScaleSerialPort.BaudRate = 2400
            '        ASCMAIN1.ScaleSerialPort.Parity = IO.Ports.Parity.None
            '        ASCMAIN1.ScaleSerialPort.DataBits = 8
            '        ASCMAIN1.ScaleSerialPort.Handshake = IO.Ports.Handshake.None
            '        ASCMAIN1.ScaleSerialPort.DtrEnable = True
            '        ASCMAIN1.ScaleSerialPort.RtsEnable = True
            '        ASCMAIN1.ScaleSerialPort.StopBits = IO.Ports.StopBits.One
            '        ASCMAIN1.ScaleSerialPort.PortName = SCALE_COMM_PORT
            '        Windows.Forms.Application.DoEvents()

            '        If Not ASCMAIN1.Running_in_VS Then
            '            ASCMAIN1.ScaleSerialPort.Open()
            '            ASCMAIN1.ScaleSerialPort.DiscardOutBuffer()
            '            ASCMAIN1.ScaleSerialPort.DiscardInBuffer()
            '        End If
            '    End If
            'Catch ex As Exception
            '    ASCMAIN1.ScaleSerialPort = Nothing
            'End Try
        End If

    End Sub

    Public Overrides Sub Get_Column_Expression_Exceptions(ByVal FORM_NAME As String, ByVal DATA_SOURCE As String, ByVal COLUMN_NAME As String, ByRef sql_SELECT_col As String) ' , ByRef sql_GROUP_BY_col As String)
        Select Case FORM_NAME

            Case "GLFASUM1"
                Select Case COLUMN_NAME
                    Case "DATA_TYPE"
                        sql_SELECT_col = "'" & DATA_SOURCE & "'"
                        'sql_SELECT_col = "'" & DATA_SOURCE & "' DATA_TYPE"
                        'sql_GROUP_BY_col = "'" & DATA_SOURCE & "'"
                End Select

            Case "ICFTRNS1"
                Select Case COLUMN_NAME
                    Case "ITEM_CODE"
                        '                        xInfo(1) = "1"
                End Select

            Case "SOFWHOD1"
                Select Case COLUMN_NAME
                    Case "ITEM_CODE"
                        '                        If DATA_SOURCE <> "A" Then xInfo(1) = "G" & CStr(j)
                End Select
        End Select
    End Sub

    Public Overrides Function Get_Code_SQL_X(ByVal FORM_NAME As String, ByVal COLUMN_NAME As String, ByRef GROUP_KEY As String) As String
        Dim sql As String = ""

        ' Set up ASWGROUP w/codes & descs from all small tables

        GROUP_KEY = COLUMN_NAME

        Select Case COLUMN_NAME
            Case "ACCT_CLASS_CODE"
                sql = "Select ACCT_CLASS_CODE, ACCT_CLASS_DESC from GLTCLAS1"
            Case "BRAND_CODE"
                sql = "Select BRAND_CODE, BRAND_NAME from ICTBRAN1"
            Case "BUS_UNIT_CODE"
                sql = "Select BUS_UNIT_CODE, BUS_UNIT_DESC from SOTBUSU1"
            Case "COLLECTION_CODE"
                sql = "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1"
                'Case "COST_CATGY_CODE"
                '    sql = "Select COST_CATGY_CODE, COST_CATGY_DESC from ICTCATG1"
            Case "CUST_BILL_TO_CUST"
                GROUP_KEY = "CUST_CODE"
                sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
            Case "CUST_BUYING_GROUP"
                GROUP_KEY = "CUST_CODE"
                sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
            Case "CUST_CITY"
                sql = "Select CUST_CITY, CUST_CITY from ARTCUST1"
            Case "CUST_CLASS_CODE"
                sql = "Select CUST_CLASS_CODE, CUST_CLASS_DESC from ARTCLAS1"
            Case "CUST_CODE"
                sql = "Select CUST_CODE, CUST_NAME from ARTCUST1"
            Case "CUST_SHIP_TO_STATE"
                GROUP_KEY = "STATE_CODE"
                sql = "Select STATE_CODE, STATE_NAME from TATSTATE"
            Case "CUST_SREP_CODE"
                GROUP_KEY = "SREP_CODE"
                sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
            Case "CUST_STORE_GROUP"
                GROUP_KEY = "CUST_CODE || '-' || CUST_STORE_GROUP"
                sql = "Select CUST_STORE_GROUP, CUST_STORE_GROUP_NAME from ARTCUST8"
            Case "CUST_STORE_NO"
                GROUP_KEY = "CUST_CODE || '-' || CUST_STORE_NO"
                sql = "Select CUST_STORE_NO, DECODE (CUST_STORE_LOCATION, NULL, CUST_STORE_NAME, CUST_CODE || ':' || CUST_STORE_LOCATION) CUST_STORE_NAME from ARTCUST2"
                If ASCMAIN1.CLIENT = "INT" Then
                    sql = "Select CUST_STORE_NO, CUST_STORE_NAME from ARTCUST2"
                End If

            Case "CUST_STORE_NO_X"
                GROUP_KEY = "CUST_CODE || '-' || CUST_STORE_NO"
                sql = "Select CUST_CODE || '-' || CUST_STORE_NO, DECODE (CUST_STORE_LOCATION, NULL, CUST_STORE_NAME, CUST_STORE_LOCATION) CUST_STORE_NAME from ARTCUST2"
            Case "CUST_STORE_STATE"
                GROUP_KEY = "STATE_CODE"
                sql = "Select STATE_CODE, STATE_NAME from TATSTATE"
            Case "DEPT_CODE"
                'sql = "Select DEPT_CODE, DEPT_DESC from ICTDEPT1"
            Case "DMA_CODE"
                sql = "Select DMA_CODE, DMA_DESC from SOTDMAC1"
            Case "FRT_TERMS"
                sql = "Select T_CODE, T_DESC from ASTCODE1 where COLUMN_NAME = '" & COLUMN_NAME & "'"
            Case "HC_CODE"
                sql = "Select HC_CODE, HC_NAME from ICTCOLL0"
            Case "INV_PYMT_METHOD"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE INV_PYMT_METHOD, T_DESC INV_PYMT_METHOD_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'APTINVH1' AND COLUMN_NAME = 'INV_PYMT_METHOD'"
            Case "ITEM_BRAND_CODE"
                sql = "Select ITEM_BRAND_CODE, ITEM_BRAND_NAME from ICTBRAN1"
            Case "ITEM_CATGY_CODE"
                sql = "Select ITEM_CATGY_CODE, ITEM_CATGY_DESC from ICTCATG1"
            Case "ITEM_CLASS_CODE"
                sql = "Select ITEM_CLASS_CODE, ITEM_CLASS_DESC from ICTCLAS1"

            Case "ITEM_CODE"
                sql = "Select ITEM_CODE, ITEM_DESC from ICTITEM1"

            Case "ITEM_GROUP_CODE"
                sql = "Select ITEM_GROUP_CODE, ITEM_GROUP_DESC from ICTGROUP"
            Case "MARKET_CODE"
                sql = "Select MARKET_CODE, MARKET_DESC from SOTMKTC1"
            Case "MATL_CATGY_CODE"
                sql = "Select MATL_CATGY_CODE, MATL_CATGY_DESC from ICTMATLA"
            Case "MATL_CODE"
                sql = "Select MATL_CODE, MATL_DESC from ICTMATL1"
            Case "METAL_CLASS_CODE"
                sql = "Select METAL_CLASS_CODE, METAL_CLASS_DESC from ICTMATLC"
            Case "OPS_YYYY"
                sql = "Select OPS_YYYY, YEAR from (Select Distinct SUBSTR(OPS_YYYYPP,1,4) OPS_YYYY, SUBSTR(OPS_YYYYPP,1,4) YEAR from GLTPARM2)"
            Case "OPS_YYYYPP"
                sql = "Select OPS_YYYYPP, LEGEND from GLTPARM2"
            Case "PRICE_CATGY_CODE"
                sql = "Select PRICE_CATGY_CODE, PRICE_CATGY_DESC from ICTPCAT1"
            Case "PRICE_POINT_CODE"
                sql = "Select PRICE_POINT_CODE, PRICE_POINT_DESC from ICTPRPT1"
            Case "POST_CODE"
                If FORM_NAME Like "AP*" Then
                    sql = "Select POST_CODE, POST_DESC from APTPOST1"
                ElseIf FORM_NAME Like "AR*" Then
                    sql = "Select POST_CODE, POST_DESC from ARTPOST1"
                End If
            Case "PROCESSOR_CODE"
                GROUP_KEY = "USER_ID"
                sql = "Select USER_ID, USER_NAME from ASTUSER1"
            Case "PROD_CODE"
                sql = "Select PROD_CODE, PROD_DESC from ICTPROD1"
            Case "PROD_CATGY_CODE"
                sql = "Select PROD_CATGY_CODE, PROD_CATGY_DESC from ICTCATG1"
            Case "REGION_CODE"
                sql = "Select REGION_CODE, REGION_DESC from SOTSREG1"
            Case "SALES_DIVISION_CODE"
                sql = "Select SALES_DIVISION_CODE, SALES_DIVISION_NAME from SOTSDIV1"
            Case "SEG2_CODE", "SEG3_CODE", "SEG4_CODE"
                GROUP_KEY = "ACCT_SEG_CODE"
                sql = "Select ACCT_SEG_CODE, ACCT_SEG_DESC from GLTSEGM1 where ACCT_SEG_ID = '" & Mid(COLUMN_NAME, 4, 1) & "'"
            Case "SELL_CODE"
                sql = "Select SELL_CODE, SELL_NAME from SOTSELL1"
            Case "SREP_CODE", "CUST_SREP_CODE"
                GROUP_KEY = "SREP_CODE"
                sql = "Select SREP_CODE, SREP_NAME from SOTSREP1"
            Case "STATE_CODE"
                sql = "Select STATE_CODE, STATE_NAME from TATSTATE"
            Case "STONE_CLASS_CODE"
                sql = "Select STONE_CLASS_CODE, STONE_CLASS_DESC from ICTMATLB"

            Case "TERM_CODE"
                sql = "Select TERM_CODE, TERM_DESC from TATTERM1"
            Case "TRADE_CLASS_CODE"
                sql = "Select TRADE_CLASS_CODE, TRADE_CLASS_DESC from SOTTCLS1"
            Case "VEND_CLASS_CODE"
                sql = "Select VEND_CLASS_CODE, VEND_CLASS_DESC from APTCLAS1"
            Case "VEND_CODE"
                sql = "Select VEND_CODE, VEND_NAME from APTVEND1"
            Case "VEND_CODE_ACC"
                GROUP_KEY = "VEND_CODE"
                sql = "Select VEND_CODE VEND_CODE_ACC, VEND_NAME from APTVEND1"
            Case "VEND_TYPE"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE VEND_TYPE, T_DESC VEND_TYPE_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'APTVEND1' AND COLUMN_NAME = 'VEND_TYPE'"
            Case "VEND_PYMT_METHOD"
                GROUP_KEY = "T_CODE"
                sql = "Select T_CODE VEND_PYMT_METHOD, T_DESC VEND_PYMT_METHOD_DESC FROM ASTCODE1 WHERE TABLE_NAME = 'APTVEND1' AND COLUMN_NAME = 'VEND_PYMT_METHOD'"
            Case "VP_CODE"
                sql = "Select VP_CODE, VP_NAME from SOTSVPS1"
            Case "WHSE_CODE"
                sql = "Select WHSE_CODE, WHSE_DESC from ICTWHSE1"
            Case Else
                ' Stop
        End Select


        If sql = "" Then
            ASCMAIN1.sql = "SELECT * FROM ASTVIEW2 WHERE VIEW_NAME = '" & COLUMN_NAME & "' AND COLUMN_POSITION IN (1,2)"
            Dim tbl As DataTable = ASCDATA1.GetDataTable
            If tbl.Rows.Count > 2 Then
                Dim MENU_ITEM_OBJECT As String = ""
                If ASCMAIN1.ActiveForm IsNot Nothing Then
                    MENU_ITEM_OBJECT = ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT
                Else
                    MENU_ITEM_OBJECT = ASCMAIN1.MENU_ITEM_OBJECT
                End If
                'ASCMAIN1.ActiveForm IS NOTHING WHEN YOU RUN AHA.SOFRMAF1 FROM MENU AS THE 1ST FORM, AND THEN YOU BLOW UP ON NEXT LINE
                ' tbl = New DataView(tbl, "TABLE_NAME LIKE '" & Mid(ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT, 1, 2) & "%'", "COLUMN_POSITION", DataViewRowState.CurrentRows).ToTable
                tbl = New DataView(tbl, "TABLE_NAME LIKE '" & Mid(MENU_ITEM_OBJECT, 1, 2) & "%'", "COLUMN_POSITION", DataViewRowState.CurrentRows).ToTable
                'ASCMAIN1.sql &= " and TABLE_NAME LIKE '" & Mid(ASCMAIN1.ActiveForm.MENU_ITEM_OBJECT, 1, 2) & "%'"
                'tbl = ASCDATA1.GetDataTable
            End If
            GROUP_KEY = ""
            If tbl.Rows.Count = 2 Then
                For Each row As DataRow In tbl.Select("", "COLUMN_POSITION")
                    If GROUP_KEY = "" Then GROUP_KEY = row.Item("COLUMN_NAME")
                    sql &= "," & row.Item("COLUMN_NAME")
                Next

                If GROUP_KEY = "T_CODE" Then ' NOT WORKING TOO WELL FOR ASTCODE1 - NEEDS A FANCIER SQL ASSEMBLER, USING ASTCODE1 INSTEAD OF TABLE_NAME, A WHERE CLAUSE, AND POSSIBLY COLUMN ALIASES
                    ' GROUP_KEY = ""
                    '  sql = ""
                    sql = "Select " & Mid(sql, 2) & " from ASTCODE1 where TABLE_NAME = '" & tbl.Rows(0).Item("TABLE_NAME") & "' and COLUMN_NAME = '" & COLUMN_NAME & "'"
                    'WORKS FOR ITEM_SNU_CODE IN FORECAST VARIANCE REPORT DPRFCVR1
                Else
                    sql = "Select " & Mid(sql, 2) & " from " & tbl.Rows(0).Item("TABLE_NAME")
                End If

            Else
                ' NEED TO USE VIEW_NAME AND TABLE_NAME FROM THE REPORT DEFINITION FOR COLUMNS THAT MIGHT BE IN MULTIPLE MASTER TABLES, LIKE CLASS_CODE, WHICH MIGHT BE AR OR AP

            End If

        End If

        Get_Code_SQL_X = sql
    End Function

    Public Overrides Sub Write_Group_Record_X(ByVal GROUP_KEY As String, ByVal COLUMN_NAME As String, ByVal GROUP_CODEs As ArrayList, ByVal GROUP_DESCs As ArrayList)
        Select Case COLUMN_NAME
            Case "ACCT_TYPE"
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":1A", "A", "Asset")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":2L", "L", "Liability")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":3E", "E", "Equity")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":4I", "I", "Income")
                'Write_Group_Record(SQLA("ACCT_TYPE", "COLUMN_CAPTION") & ":5X", "X", "Expense")

            Case "EXC_OBS_IND"
                GROUP_CODEs.Add("E") : GROUP_DESCs.Add("Excess Inventory On Hand")
                GROUP_CODEs.Add("O") : GROUP_DESCs.Add("Obsolete Inventory (No Demand)")

        End Select
    End Sub

    Public Overrides Function CodeValues(ByVal CodeValueKey As String) As Dictionary(Of String, String)

        Dim VL As New Dictionary(Of String, String)

        ' all of the codes in here should be moved to ASTCODE1

        Select Case CodeValueKey

            Case "CCPA_REASON"
                VL.Add("A", "Auto Stmt")
                VL.Add("B", "Bank")
                VL.Add("C", "Sale Captured")
                VL.Add("M", "Manual")
                VL.Add("O", "Order")
                VL.Add("S", "Statement")
                VL.Add("W", "Web Payment")

            Case "CCPA_STATUS"
                VL.Add("0", "Sales Q")
                VL.Add("1", "Stmt Re-Q")
                VL.Add("2", "Auto Charge-Q")
                VL.Add("E", "Declined")
                VL.Add("C", "Auth Captured")
                VL.Add("A", "Approved")
                VL.Add("S", "Settled")
                VL.Add("V", "Voided")
                VL.Add("D", "Deleted")
                VL.Add("T", "Authorized")
                VL.Add("X", "Deleted Auth")

            Case "CUST_STORE_STATUS"
                VL.Add("A", "Active")
                VL.Add("I", "In-Active")

            Case "ITEM_STATUS"
                VL.Add("A", "Active")
                VL.Add("I", "Inactive")

            Case "ORDR_STATUS"
                VL.Add("O", "Open")
                VL.Add("H", "On Hold")
                VL.Add("P", "In Pick")
                VL.Add("F", "Completed")
                VL.Add("V", "Voided")
                VL.Add("C", "Cancelled")

            Case "PERIOD_END_DAY"
                VL.Add("0", "Calendar")
                VL.Add("1", "Sunday")
                VL.Add("2", "Monday")
                VL.Add("3", "Tuesday")
                VL.Add("4", "Wednesday")
                VL.Add("5", "Thursday")
                VL.Add("6", "Friday")
                VL.Add("7", "Saturday")

            Case "PO_STAGE"
                VL.Add("0", "New")
                VL.Add("1", "Revised")
                VL.Add("2", "Cancelled")
                VL.Add("3", "Closed")
                VL.Add("4", "Confirmed")
                VL.Add("5", "Wait to Ship")

            Case "RESPONSE_CODE"
                VL.Add("A", "Approved")
                VL.Add("E", "Error")

            Case "WEEK_END_DAY"
                VL.Add("1", "Sunday")
                VL.Add("2", "Monday")
                VL.Add("3", "Tuesday")
                VL.Add("4", "Wednesday")
                VL.Add("5", "Thursday")
                VL.Add("6", "Friday")
                VL.Add("7", "Saturday")

            Case "WHSE_TYPE"
                VL.Add("S", "Shipment")
                VL.Add("R", "Returns")
                VL.Add("F", "Refurb")
                VL.Add("X", "Nav-Trav")
                VL.Add("I", "In Transit")
                VL.Add("N", "Non-Inv")

        End Select

        Return VL

    End Function

    Public Shared Function Send_email_with_Attachment(
        ByVal frmASFBASE0 As ASFBASE0,
        ByVal FILENAME As String,
        ByVal ATTACHMENT As String,
        ByVal SUBJECT As String,
        Optional EMAIL_ADDRESS As String = "",
        Optional EMAIL_NAME As String = "",
        Optional EMAIL_KEY As String = "",
        Optional ENTITY_KEY As String = "",
        Optional ENTITY_NAME As String = "",
        Optional ENTITY_CAPTION As String = "")

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        ATTACHMENTs.Add(ATTACHMENT, FILENAME)

        SUBJECT = ASCMAIN1.rowASTPARM1.Item("AS_PARM_INST_NAME") & " " & SUBJECT

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        EMAIL_ADDRESSs.Add(EMAIL_ADDRESS, EMAIL_NAME)

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
               (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                SUBJECT, EMAIL_KEY, False, True, ENTITY_KEY, ENTITY_NAME, ENTITY_CAPTION)

        Return SEND_NO

    End Function

    Public Overrides Function Send_email(ByVal frmASFBASE0 As ASFBASE0,
                                 ByVal EMAIL_ADDRESSs As Dictionary(Of String, String),
                                 ByVal ATTACHMENTs As Dictionary(Of String, String),
                                 ByVal SUBJECT As String,
                                 ByVal EMAIL_KEY As String,
                                 Optional ByVal auto_send As Boolean = False,
                                 Optional SEND_CC_to_USER_ID As Boolean = False,
                                 Optional ENTITY_KEY As String = "",
                                 Optional ENTITY_NAME As String = "",
                                 Optional ENTITY_CAPTION As String = "",
                                 Optional EMAIL_BODY As String = "") As String

        Dim USER_ID_emailer As String = ASCMAIN1.USER_ID
        Dim rowASTUSER1_EMAIL_FROM As DataRow = frmASFBASE0.LookUp("ASTUSER1", USER_ID_emailer, True)
        Dim rowASTUSER1_EMAIL_BCC As DataRow = Nothing

        Dim USER_TELEPHONE As String = rowASTUSER1_EMAIL_FROM.Item("USER_TELEPHONE") & ""
        Dim USER_EXT As String = rowASTUSER1_EMAIL_FROM.Item("USER_EXT") & ""
        Dim USER_FAX As String = rowASTUSER1_EMAIL_FROM.Item("USER_FAX") & ""
        ' Dim EMAIL_BODY As String = "Attached is the file that you have requested."

        '****************
        ' IMPORTANT - make sure that you have a TATMAIL1 record established, and in most cases you will need to populate EMAIL_FROM
        '****************

        Dim rowTATMAIL1 As DataRow = frmASFBASE0.LookUp("TATMAIL1", EMAIL_KEY)
        If rowTATMAIL1 IsNot Nothing Then
            If rowTATMAIL1.Item("EMAIL_FROM") & "" <> "" Then
                rowASTUSER1_EMAIL_FROM = frmASFBASE0.LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_FROM"), True)
            Else
                If rowTATMAIL1.Item("EMAIL_ACCT_ID") & "" <> "" Then
                    rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") = rowTATMAIL1.Item("EMAIL_ACCT_ID")
                End If
            End If
            If rowTATMAIL1.Item("EMAIL_BCC") & "" <> "" Then
                rowASTUSER1_EMAIL_BCC = frmASFBASE0.LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_BCC"), True)
            End If
            If EMAIL_BODY = "" Then
                If rowTATMAIL1.Item("EMAIL_BODY") & "" <> "" Then
                    EMAIL_BODY = rowTATMAIL1.Item("EMAIL_BODY")
                Else
                    EMAIL_BODY = "Attached is the file that you have requested."
                End If
            End If

        End If

        Dim USER_SIGNATURE As String =
          rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & vbCrLf _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & vbCrLf, "") _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & vbCrLf, "") _
        & IIf(USER_TELEPHONE <> "", "Tel: " & ASCMAIN1.FormatTel(USER_TELEPHONE, USER_EXT) & vbCrLf, "") _
        & IIf(USER_FAX <> "", "Fax: " & ASCMAIN1.FormatTel(USER_FAX) & vbCrLf, "") _
        & rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & vbCrLf

        Dim frmTAFSEND1 As New TAFSEND1(frmASFBASE0)
        frmTAFSEND1.EMAIL_KEY = EMAIL_KEY
        frmTAFSEND1.SEND_FROM = rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & ""
        frmTAFSEND1.SEND_FROM_NAME = rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & ""
        frmTAFSEND1.SEND_FROM_SIGNATURE = USER_SIGNATURE
        frmTAFSEND1.SEND_TOs = EMAIL_ADDRESSs
        frmTAFSEND1.SEND_TO = ""
        frmTAFSEND1.SEND_TO_NAME = ""
        frmTAFSEND1.SEND_CC = ""
        If SEND_CC_to_USER_ID Then
            frmTAFSEND1.SEND_CC = ASCMAIN1.USER_EMAIL
            frmTAFSEND1.SEND_CC_NAME = ASCMAIN1.USER_NAME
        End If
        If rowASTUSER1_EMAIL_BCC IsNot Nothing Then
            frmTAFSEND1.SEND_BCC = rowASTUSER1_EMAIL_BCC.Item("USER_EMAIL") & ""
            frmTAFSEND1.SEND_BCC_NAME = rowASTUSER1_EMAIL_BCC.Item("USER_NAME") & ""
        End If

        frmTAFSEND1.SEND_SUBJECT = SUBJECT

        frmTAFSEND1.SEND_BODY = EMAIL_BODY
        frmTAFSEND1.SEND_ENTITY_KEY = ENTITY_KEY
        frmTAFSEND1.SEND_ENTITY_NAME = ENTITY_NAME
        frmTAFSEND1.SEND_METHOD = "E"
        frmTAFSEND1.SEND_ENTITY_CAPTION = ENTITY_CAPTION
        frmTAFSEND1.SEND_ATTACHMENTs = ATTACHMENTs
        frmTAFSEND1.SEND_ATTACHMENT = ""


        Dim ENTITY_TABLE As String = "" ' THIS IS NOT THE RIGHT PLACE TO BE SETTING THIS- IT SHOULD BE PASSED IN
        If ENTITY_CAPTION = "Customer" Then ENTITY_TABLE = "ARTCUST1"
        If ENTITY_CAPTION = "Vendor" Or ENTITY_CAPTION = "Supplier" Then ENTITY_TABLE = "APTVEND1"

        frmTAFSEND1.SEND_ENTITY_TABLE = ENTITY_TABLE

        If auto_send Then
            frmTAFSEND1.Send_email_automatically()
        Else
            frmTAFSEND1.ShowDialog()
        End If

        Dim SEND_STATUS As String = frmTAFSEND1.SEND_STATUS
        Dim SEND_NO As String = frmTAFSEND1.SEND_NO

        If frmTAFSEND1.SEND_ERROR <> "" Then
            If EMAIL_KEY = "PC_PRCEXC" Then
                MsgBox("Error trying to send email" & vbCrLf & frmTAFSEND1.SEND_ERROR, MsgBoxStyle.OkOnly, "Please send screenshot to Lauren")
            End If
        End If

        frmTAFSEND1.Dispose()
        frmTAFSEND1 = Nothing

        Return SEND_NO ' SEND_STATUS

    End Function

    Public Overrides Sub Application_Initialization()
        'PUT CODE HERE FOR WHEN YOU ARE LOGGING IN
        'Dim rowICTPARM1 As DataRow = ASCDATA1.GetDataRow("Select * from ICTPARM1 where IC_PARM_KEY = 'Z'")
        'If rowICTPARM1.Item("IC_PARM_CYW_LAST") & "" <> ASCMAIN1.CYW Then
        '    ASCMAIN1.sql = "Select * from GLTCOMP1 where COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        '    Dim rowGLTCOMP1 As DataRow = ASCDATA1.GetDataRow
        '    If rowGLTCOMP1 IsNot Nothing Then
        '        ASCMAIN1.Progress("Now Taking Weekly Inventory Snapshot")
        '        ASCDATA1.ExecuteSP("ICPLOTD9")
        '        ASCMAIN1.Progress("")
        '    End If
        'End If
    End Sub

    Public Overrides Sub Maintain_Contacts(ByVal frmASFBASE1 As ASFBASE1,
                                           ByVal CONTACT_ENTITY_TABLE As String,
                                           ByVal CONTACT_ENTITY_KEY As String,
                                           ByVal CONTACT_ENTITY_NAME As String)
        'MyBase.Maintain_Contacts()
        Using frmTAFCONT1 As New TAFCONT1(frmASFBASE1)
            With frmTAFCONT1
                .CONTACT_ENTITY_TABLE = CONTACT_ENTITY_TABLE
                .CONTACT_ENTITY_KEY = CONTACT_ENTITY_KEY
                .CONTACT_ENTITY_NAME = CONTACT_ENTITY_NAME
                .ShowDialog()
            End With
        End Using

    End Sub

    Public Overrides Function Custom_sqlwhere(
    ByVal sqlwhere As String,
    ByVal grd As UltraWinGrid.UltraGrid,
    ByVal COLUMN_NAME As String) As String
        If grd.Name = "grd" And grd.TopLevelControl.Name = "ASFCODE1" _
             And ASCMAIN1.CodeSelector IsNot Nothing _
             And ASCMAIN1.CodeSelector.VIEW_NAME = "CUST_CODE" _
             And ASCMAIN1.CodeSelector.TABLE_NAME = "ARTCUST1" _
             And COLUMN_NAME = "CUST_NAME" Then
            sqlwhere = Mid(sqlwhere, 6)
            sqlwhere = " and (" & sqlwhere & " or " & Replace(sqlwhere, "CUST_NAME", "CUST_DBA_NAME") & ")"

            If grd.ActiveRow IsNot Nothing AndAlso grd.ActiveRow.IsFilterRow Then
                grd.DisplayLayout.Bands(0).ColumnFilters("CUST_NAME").ClearFilterConditions()
                'grd.ActiveRow.Cells("CUST_NAME").Column.F.Value = DBNull.Value
            End If

        End If


        Return sqlwhere
    End Function

    Public Shared Function Get_EDI_Custs(EDI_DOC_NO As String) As List(Of String)

        Dim EDI_Custs As New List(Of String)

        ASCMAIN1.sql = "Select DISTINCT CUST_CODE from EDTTRPM1 where CUST_CODE is Not Null"
        If EDI_DOC_NO <> "" Then
            ASCMAIN1.sql &= " and EDI_DOC_NO = '" & EDI_DOC_NO & "'"
        End If
        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            EDI_Custs.Add(row.Item("CUST_CODE"))
        Next
        Return EDI_Custs
    End Function

    Public Shared Function Calculate_INV_DUE_DATE(
                                                 F As ASFBASE0,
                                                 TERM_CODE As String,
                                                 rowTATTERM1 As DataRow,
                                                 INV_BASE_DATE As Object) As Date

        Dim INV_DUE_DATE As Object = Nothing

        If TERM_CODE = "" Or INV_BASE_DATE Is Nothing Then
            Return INV_BASE_DATE
            'Exit Function
        End If

        If rowTATTERM1 Is Nothing Then
            rowTATTERM1 = F.LookUp("TATTERM1", TERM_CODE, True)
        End If

        Select Case rowTATTERM1.Item("TERM_DUE_TYPE") & ""
            Case "C" ' IE, COD
                INV_DUE_DATE = INV_BASE_DATE

            Case "D"
                INV_DUE_DATE = INV_BASE_DATE.AddDays(Val(rowTATTERM1.Item("TERM_DAYS_DUE") & ""))

            Case "S"
                If rowTATTERM1.Item("TERM_CUTOFF_DATE") & "" <> "" Then
                    Dim TERM_CUTOFF_DATE As Date = rowTATTERM1.Item("TERM_CUTOFF_DATE")
                    If Format(INV_BASE_DATE, "MMdd") > Format(TERM_CUTOFF_DATE, "MMdd") Then
                        INV_DUE_DATE = CDate(Format(TERM_CUTOFF_DATE, "MM/dd") & "/" & Format(Val(Format(INV_BASE_DATE, "yyyy")) + 1, "0000"))

                    Else
                        INV_DUE_DATE = CDate(Format(TERM_CUTOFF_DATE, "MM/dd") & "/" & Format(INV_BASE_DATE, "yyyy"))
                    End If

                End If

            Case "E"
                Dim ADD_MONTHS_BASE As Integer = 1
                Dim TERM_CUTOFF_DAY As Integer = Val(rowTATTERM1.Item("TERM_CUTOFF_DAY") & "")
                Dim BASE_DD As Integer = Val(Format(INV_BASE_DATE, "dd"))
                Dim TERM_DAYS_DUE As Integer = Val(rowTATTERM1.Item("TERM_DAYS_DUE") & "")
                Dim TERM_ADDL_MOS As Integer = Val(rowTATTERM1.Item("TERM_ADDL_MOS") & "")
                Dim INV_BASE_DATEx As String = Format(INV_BASE_DATE, "MM/dd/yyyy")

                Select Case rowTATTERM1.Item("TERM_EOM_TYPE") & ""
                    Case "F"
                        ASCMAIN1.sql = "Select GLTPARM2.* " _
                         & " from GLTPARM2 " _
                         & " where OPS_YYYYPP = " _
                         & " (Select Min(OPS_YYYYPP) from GLTPARM2 " _
                         & "  where GLTPARM2.PRD_END_DATE >= '" & Format(INV_BASE_DATE, "dd-MMM-yyyy") & "')"
                        Dim rowGLTPARM2 As DataRow = ASCDATA1.GetDataRow
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


    Public Shared Function CreateSampleLabel() As String

        Dim labelImage As String = String.Empty

        'labelImage = "" & Environment.NewLine
        labelImage = "EPL2" & Environment.NewLine
        labelImage &= "q812" & Environment.NewLine
        labelImage &= "Q1218, 24 + 0" & Environment.NewLine
        labelImage &= "S4" & Environment.NewLine
        labelImage &= "UN" & Environment.NewLine
        labelImage &= "WN" & Environment.NewLine
        labelImage &= "ZB" & Environment.NewLine
        labelImage &= "N" & Environment.NewLine
        labelImage &= "A58,53,0,5 4,4 N,""P""" & Environment.NewLine
        labelImage &= "LO253, 0, 2, 304" & Environment.NewLine

        ' Invoice BarCode
        labelImage &= "B279,96,0,1E,2,2,76,B,""1234567890""" & Environment.NewLine
        'label &= "A110, 1117, 0, 4, 1, 1, N, ""420296809101931490053735655653""" & Environment.NewLine

        labelImage &= "A100,313 0,5,1,1,N,""Test Label""" & Environment.NewLine
        labelImage &= "LO0,304,812 2" & Environment.NewLine
        labelImage &= "LO0,369 812 2" & Environment.NewLine

        ' Sender
        labelImage &= "A12,375,0,3,1,1,N,""Test Company""" & Environment.NewLine
        labelImage &= "A12,398,0,3,1,1,N,""123 Main Street""" & Environment.NewLine
        labelImage &= "A12,421,0,3,1,1,N,""AnyTown NY 10532-0220""" & Environment.NewLine
        labelImage &= "A601,552,0,2,1,1,N,""016917/0010427616""" & Environment.NewLine

        labelImage &= "A62,622,0,4,1,1,N,""SHIP""" & Environment.NewLine
        labelImage &= "A78,649,0,4,1,1,N,""TO:""" & Environment.NewLine

        ' Recipient
        labelImage &= "A152,616,0,4,1,1,N,""Joe Consumer""" & Environment.NewLine
        labelImage &= "A152,643,0,4,1,1,N,""1313 Mockingbird Lane""" & Environment.NewLine
        labelImage &= "A152,670,0,4,1,1,N,""SIMPSONVILLE SC 29680-7008""" & Environment.NewLine
        labelImage &= "LO0,831,812,6" & Environment.NewLine

        ' Delivery Confirmation Bar Code
        labelImage &= "B72,938,0,1E,3,3,152,N,""420296809101931490053735655653""" & Environment.NewLine
        labelImage &= "A110,1117,0,4,1,1,N,""420296809101931490053735655653""" & Environment.NewLine
        labelImage &= "LO0,1167,812,6" & Environment.NewLine

        labelImage &= "P1" & Environment.NewLine

        Return labelImage
    End Function

    Public Shared Function Create_Zip_File(FILENAME As String, FILENAMES_to_zip() As String) As Boolean

        Try
            If My.Computer.FileSystem.FileExists(ASCMAIN1.Folders("Temp") & FILENAME & ".zip") Then
                My.Computer.FileSystem.DeleteFile(ASCMAIN1.Folders("Temp") & FILENAME & ".zip")
            End If

            Dim Zip1 As New nsoftware.IPWorksZip.Zip
            Zip1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareZipkey")
            Zip1.ArchiveFile = ASCMAIN1.Folders("Temp") & FILENAME & ".zip"
            Dim FILENAMES As String = Join(FILENAMES_to_zip, " | ")
            'Zip1.IncludeFiles(ASCMAIN1.Folders("Temp") & "customer.csv" & " | " & ASCMAIN1.Folders("Temp") & "part.csv" & " | " & ASCMAIN1.Folders("Temp") & "sales.csv")
            Zip1.IncludeFiles(FILENAMES)
            Zip1.Compress()
            Zip1.Dispose()
            Return True

        Catch ex As Exception
            MsgBox("Error Creating Zip File")
            Return False
        End Try

    End Function

    Public Shared Function ftp_Files(
        FOLDERNAME_local As String,
        FILENAME_local() As String,
        FOLDERNAME_remote As String,
        FILENAME_remote() As String,
        USER As String,
        PWD As String,
        HOST As String) As Boolean

        Try
            Dim Ftp1 As New nsoftware.IPWorks.Ftp
            Ftp1.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareftpkey")

            Ftp1.User = USER
            Ftp1.Password = PWD
            Ftp1.RemoteHost = HOST
            Ftp1.Logon()
            Ftp1.ChangeTransferMode(nsoftware.IPWorks.FTPTransferModes.tmBinary)
            Ftp1.ChangeRemotePath(FOLDERNAME_remote)
            'Ftp1.RemotePath = FOLDERNAME_remote

            For i As Integer = 0 To FILENAME_local.Length - 1
                'Ftp1.LocalFile = ASCMAIN1.Folders("Temp") & FILENAME_local
                Ftp1.LocalFile = FOLDERNAME_local & FILENAME_local(i)
                Ftp1.RemoteFile = FILENAME_remote(i)
                Ftp1.Upload()
            Next

            Ftp1.Logoff()

            Return True

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.OkOnly, "Error ftp'ing File")
            Return False

        End Try
    End Function

    Public Shared Function Check_Permissions(F As ASFSRPTM) As String

        Dim EMsg As String = ""

        If Trim(ASCMAIN1.USER_CODES) = "FS" Then

            Dim SREP_CODEs_array() As String = TAC.TACMAIN1.SREP_CODEs.ToArray
            Array.Sort(SREP_CODEs_array)
            Dim SREP_CODEs As String = Join(SREP_CODEs_array, ",")
            Dim rowASTDSQLA As DataRow = F.tblASTDSQLA.Rows.Find("SREP_CODE")
            If rowASTDSQLA.Item("EXCLUDE") & "" = "1" Then
                If TAC.TACMAIN1.SREP_CODEs.Count > 1 Then
                    EMsg &= vbCr & "You may not use Exclude in a Filter"
                Else
                    EMsg &= vbCr & "You must leave Filter set INCLUDE to Sales Rep(s) " & SREP_CODEs & " only"
                End If
            Else
                Dim SREP_CODE_filter As String = F.SQLA("SREP_CODE")
                If SREP_CODE_filter <> TAC.TACMAIN1.SREP_CODE And SREP_CODE_filter <> SREP_CODEs Then
                    If SREP_CODE_filter = "" Then
                        If TAC.TACMAIN1.SREP_CODEs.Count = 1 Then
                            EMsg &= vbCr & "You must set the Sales Rep Filter to Sales Rep " & SREP_CODEs
                        Else
                            EMsg &= vbCr & "You must set the Sales Rep Filter to one or more of the following: " & SREP_CODEs
                        End If
                    Else
                        For Each SREP_CODE As String In Split(SREP_CODE_filter, ",")
                            If Not TAC.TACMAIN1.SREP_CODEs.Contains(SREP_CODE) Then
                                If TAC.TACMAIN1.SREP_CODEs.Count = 1 Then
                                    EMsg &= vbCr & "You must set the Sales Rep Filter to Sales Rep " & SREP_CODEs
                                Else
                                    EMsg &= vbCr & "You must set the Sales Rep Filter to one or more of the following: " & SREP_CODEs
                                End If
                            End If
                        Next
                    End If
                    'If TAC.TACMAIN1.SREP_CODEs.Count > 1 And SREP_CODE_filter <> "" Then
                    '    For Each SR As String In Split(SREP_CODE_filter, ",")
                    '        If TAC.TACMAIN1.SREP_CODEs.Contains(SR) Then
                    '        Else
                    '            EMsg &= vbCr & "You may not select Sales Rep " & SR
                    '        End If
                    '    Next
                    'Else
                    '    EMsg &= vbCr & "You must leave Filter set to Sales Rep(s) " & SREP_CODEs & " only"
                    'End If
                End If
            End If
        End If

        Return EMsg

    End Function

    Public Overloads Shared Sub Record_Event(
         ByVal TABLE_NAME As String,
         ByVal TABLE_KEY As String,
         ByVal INIT_DATE As Date,
         ByVal INIT_OPER As String,
         ByVal EVENT_TYPE As String,
         ByVal EVENT_DESC As String,
         Optional ByVal EVENT_KEY As String = "",
         Optional FORM_NAME As String = "")

        If FORM_NAME = "" Then
            FORM_NAME = ASCMAIN1.ActiveForm.Name
        End If

        ASCDATA1.ExecuteSQL("Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME) " _
                             & " Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7,:PARM8)",
                             "VVDVVVVV",
                             New Object() {TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY, FORM_NAME})
    End Sub

    Public Shared Function Get_CURR_EXCH_RATE(
        frmASFBASE0 As ASFBASE0,
        CURR_CODE As String,
        TRAN_DATE As Date,
        Optional check_forex_if_missing_daily_rate As Boolean = True) As Decimal

        Dim CURR_EXCH_RATE As Decimal = 0
        If TRAN_DATE.ToString = "1/1/0001 12:00:00 AM" Then
            TRAN_DATE = Now.Date
        End If
        If CURR_CODE = "" Or CURR_CODE = frmASFBASE0.ROWs("GLTPARM1").Item("GL_PARM_CURR_CODE") Then
            CURR_EXCH_RATE = 1
        Else
            ASCMAIN1.sql = "Select * from TATCURR3" & vbCrLf _
              & " where CURR_CODE = :PARM1" & vbCrLf _
              & "   and  CURR_DATE = (Select Max(CURR_DATE) from TATCURR3" & vbCrLf _
              & " where CURR_CODE = :PARM2 and CURR_DATE <= :PARM3)"
            Dim rowTATCURR3 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, "VVD", New Object() {CURR_CODE, CURR_CODE, TRAN_DATE})
            If rowTATCURR3 IsNot Nothing Then
                Dim X As Integer = 30
                Dim CURR_DATE As Date = rowTATCURR3.Item("CURR_DATE")
                Dim DAYS_OLD As Integer = TRAN_DATE.Subtract(CURR_DATE).Days
                CURR_EXCH_RATE = Val(rowTATCURR3.Item("CURR_EXCH_RATE") & "")
                If DAYS_OLD > X Then
                    CURR_EXCH_RATE = 0
                End If
            Else
                CURR_EXCH_RATE = 0
                If check_forex_if_missing_daily_rate Then
                    TAC.TACMAIN1.Update_Forex()
                    CURR_EXCH_RATE = Get_CURR_EXCH_RATE(frmASFBASE0, CURR_CODE, TRAN_DATE, False)
                End If
            End If
        End If

        If CURR_EXCH_RATE = 0 Then
            Throw New Exception("Cannot determine Exchange Rate in TACMAIN1.Get_CURR_EXCH_RATE")
        End If

        Return CURR_EXCH_RATE
    End Function
    Public Shared Sub Update_Forex()

        ASCMAIN1.sql = "Select * from TATCURR1" _
            & " Where CURR_CODE <> 'USD'"

        For Each rowTATCURR1 As DataRow In ASCDATA1.GetDataTable.Rows

            Dim gotTodaysRate As Boolean = False
            Dim CURR_EXCH_RATE_response As String = ""
            Dim INIT_DATE As Date
            Dim INIT_OPER As String

            Dim CURR_CODE As String = rowTATCURR1.Item("CURR_CODE") & ""
            Dim sqlCurr As String = "SELECT TATCURRX.CURR_DATE_X, '" & CURR_CODE & "' CURR_CODE_X" _
            & ", NVL(TATCURR3.CURR_EXCH_RATE,0) CURR_EXCH_RATE_X" _
            & " FROM TATCURR3, (select TRUNC(SYSDATE - 60) + rownum - 1 CURR_DATE_X" _
            & " from all_objects" _
            & " where rownum <= to_date(SYSDATE,'dd-mon-yyyy')-to_date(TRUNC(SYSDATE -60),'dd-mon-yyyy')+1) TATCURRX" _
            & " WHERE " _
            & " TATCURR3.CURR_DATE (+) = TATCURRX.CURR_DATE_X"
            Dim tblTATCURRX As DataTable = ASCDATA1.GetDataTable(sqlCurr)

            If tblTATCURRX.Rows.Count <> 0 Then
                For Each rowTATCURRX As DataRow In tblTATCURRX.Select("", "CURR_DATE_X")
                    Dim CURR_EXCH_RATE As Decimal = 0
                    Dim CURR_EXCH_RATE_X As Decimal = Val(rowTATCURRX.Item("CURR_EXCH_RATE_X") & "")
                    Dim CURR_DATE_X As Date = rowTATCURRX.Item("CURR_DATE_X")
                    If CURR_EXCH_RATE_X = 0 Then
                        If CURR_DATE_X.Date = Now + ASCMAIN1.NowTSD Then
                            CURR_EXCH_RATE = Get_Current_Exchange_Rate(CURR_CODE)
                            If CURR_EXCH_RATE <> 0 Then
                                gotTodaysRate = True
                            End If
                        Else
                            CURR_EXCH_RATE = Get_Historical_Exchange_Rate(CURR_CODE, CURR_DATE_X)
                        End If

                        If CURR_EXCH_RATE <> 0 Then
                            INIT_OPER = ASCMAIN1.USER_ID
                            INIT_DATE = Now + ASCMAIN1.NowTSD
                            ASCMAIN1.sql = "Insert into TATCURR3 Values (:PARM1,:PARM2,:PARM3,:PARM4,:PARM5,:PARM6,:PARM7)"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VDNDVDV", New Object() {CURR_CODE, CURR_DATE_X, CURR_EXCH_RATE, INIT_DATE, INIT_OPER, INIT_DATE, INIT_OPER})

                            Dim cd As String = Format(CURR_DATE_X.Date, "MM/dd/yyyy")
                        End If
                    Else
                        If CURR_DATE_X.Date = Now Then
                            gotTodaysRate = True
                        End If
                    End If
                Next
            End If
            'ASCMAIN1.sql = "Select GLTPARM2.*,TATCURR2.CURR_EXCH_CUR, TATCURR3.CURR_EXCH_RATE" & vbCrLf _

            Dim YP_END As String = ASCMAIN1.CYP
            Dim YP_BEG As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6)

            ASCMAIN1.sql = "Select GLTPARM2.OPS_YYYYPP,TATCURR3.CURR_CODE,TATCURR3.CURR_EXCH_RATE" & vbCrLf _
                & " from GLTPARM2,TATCURR3" & vbCrLf _
                & " where GLTPARM2.OPS_YYYYPP <= '" & YP_END & "'" & vbCrLf _
                & "   and GLTPARM2.OPS_YYYYPP >= '" & YP_BEG & "'" & vbCrLf _
                & "   and TATCURR3.CURR_DATE = GLTPARM2.PRD_END_DATE" & vbCrLf _
                & "   and TATCURR3.CURR_CODE = '" & CURR_CODE & "'" & vbCrLf _
                & "   and GLTPARM2.OPS_YYYYPP not in " & vbCrLf _
                & " (Select OPS_YYYYPP from TATCURR2 where CURR_CODE = '" & CURR_CODE & "')"
            ASCMAIN1.sql = "Insert into TATCURR2 " & vbCrLf & ASCMAIN1.sql
            ASCDATA1.ExecuteSQL()

            'ASCMAIN1.sql = "Select GLTPARM2.*,TATCURR2.CURR_EXCH_CUR, TATCURR3.CURR_EXCH_RATE" & vbCrLf _
            '    & " from GLTPARM2,TATCURR2,TATCURR3" & vbCrLf _
            '    & " where GLTPARM2.OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'" & vbCrLf _
            '    & "   and GLTPARM2.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -6) & "'" & vbCrLf _
            '    & "   and TATCURR2.OPS_YYYYPP (+) = GLTPARM2.OPS_YYYYPP" & vbCrLf _
            '    & "   and TATCURR2.CURR_CODE (+) = '" & CURR_CODE & "'" & vbCrLf _
            '    & "   and TATCURR3.CURR_DATE = GLTPARM2.PRD_END_DATE" & vbCrLf _
            '    & "   and TATCURR3.CURR_CODE = '" & CURR_CODE & "'" & vbCrLf

            'For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            '    Dim CURR_EXCH_CUR As Decimal = Val(row.Item("CURR_EXCH_CUR") & "")
            '    Dim CURR_EXCH_RATE As Decimal = Val(row.Item("CURR_EXCH_RATE") & "")
            '    If CURR_EXCH_RATE <> 0 And CURR_EXCH_CUR = 0 Then
            '        ASCMAIN1.sql = "Insert into TATCURR2 Values (:PARM1,:PARM2,:PARM3)"
            '        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVN", New Object() {row.Item("OPS_YYYYPP"), CURR_CODE, CURR_EXCH_RATE})
            '    End If
            'Next

            If Not gotTodaysRate Then
                ' MessageBox.Show("Error getting daily exchange rate for: " & CURR_CODE, "Error", MessageBoxButtons.OK)
            End If

        Next

        ASCMAIN1.Progress("")

    End Sub
    Public Shared Function Get_Current_Exchange_Rate(forCode As String) As Decimal

        Dim CURR_EXCH_RATE As Decimal = 0
        Dim CURR_EXCH_RATE_response As String
        Dim forexSvc As String
        For Each RATE_API As String In New String() {"appspot", "ratelab", "openexchange"}
            forexSvc = RATE_API
            CURR_EXCH_RATE_response = ""
            CURR_EXCH_RATE_response = 1 ' need to re-test if and when foreign curr is required
            'CURR_EXCH_RATE_response = Rate_By_Service(forCode, Now + ASCMAIN1.NowTSD, RATE_API)
            ASCMAIN1.Progress("Curr: " & forCode & ", Date: " & Now + ASCMAIN1.NowTSD, "Rate: " & CURR_EXCH_RATE_response.ToString)
            If CURR_EXCH_RATE_response <> "" AndAlso (Val(CURR_EXCH_RATE_response) <> 0) Then
                Return Val(CURR_EXCH_RATE_response)
            End If
        Next

        Return CURR_EXCH_RATE

    End Function

    Public Shared Function Get_Historical_Exchange_Rate(forCode As String, forDate As Date) As Decimal

        Dim CURR_EXCH_RATE As Decimal = 0
        Dim CURR_EXCH_RATE_response As String = ""
        Dim forexSvc As String
        forexSvc = "openexchangeHistory"
        CURR_EXCH_RATE_response = 1 ' until we need fcx
        'CURR_EXCH_RATE_response = Rate_By_Service(forCode, forDate, "openexchangeHistory")
        If CURR_EXCH_RATE_response <> "" AndAlso (Val(CURR_EXCH_RATE_response) <> 0) Then
            ASCMAIN1.Progress("Curr: " & forCode & ", Date: " & forDate.Date.ToString, CURR_EXCH_RATE_response.ToString)
            Return Val(CURR_EXCH_RATE_response)
        End If

        Return CURR_EXCH_RATE

    End Function



    'Public Shared Function Rate_By_Service(forCode As String, rateDate As Date, rateService As String) As String
    '    Dim responseString As String = ""
    '    Dim client As New HttpClient()
    '    Dim API_BASE As String = ""
    '    Dim API_METHOD As String = ""
    '    Dim API_QUERY_STRING As String = ""

    '    Select Case rateService
    '        Case "appspot"
    '            'no api key needed
    '            ' Sample URL: http://rate-exchange.appspot.com/currency?from=CAD&to=USD
    '            ' {"to": "USD", "rate": 0.79647100000000004, "from": "CAD"}
    '            API_BASE = "http://rate-exchange.appspot.com/"
    '            API_METHOD = "currency"
    '            API_QUERY_STRING = "?from=" & forCode & "&to=USD"
    '        Case "ratelab"
    '            'base is USD so response needs to to be 1/rate
    '            'apiKey=27429B02DA56DD370A6A9091430DD0F1
    '            'Sample URL: http://api.exchangeratelab.com/api/single/CAD?apikey=27429B02DA56DD370A6A9091430DD0F1
    '            'Sample Response
    '            '{"rate":{"rate":1.2582,"to":"USD"},"baseCurrency":"CAD","timeStamp":1429023188,"executionTime":28,"licenseMessage":"Data Retrieved From www.ExchangeRateLab.com - Under license (Not for financial/professional use)"}
    '            API_BASE = "http://api.exchangeratelab.com/"
    '            API_METHOD = "api/single/" & forCode
    '            API_QUERY_STRING = "?apikey=27429B02DA56DD370A6A9091430DD0F1"

    '        Case "openexchange"
    '            'base is USD so response needs to to be 1/rate
    '            'app_id=44076a2ca9a243b3b61f08219fb7809f
    '            'Sample URL: https://openexchangerates.org/api/latest.json?app_id=44076a2ca9a243b3b61f08219fb7809f
    '            API_BASE = "https://openexchangerates.org/"
    '            API_METHOD = "api/latest.json"
    '            API_QUERY_STRING = "?app_id=44076a2ca9a243b3b61f08219fb7809f"
    '        Case "openexchangeHistory"
    '            'https://openexchangerates.org/api/historical/2015-04-21.json?app_id=44076a2ca9a243b3b61f08219fb7809f
    '            Dim rateYear As String = rateDate.Year.ToString
    '            Dim rateMonth As String = rateDate.Month.ToString("00")
    '            Dim rateDay As String = rateDate.Day.ToString("00")
    '            Dim rd As String = rateYear & "-" & rateMonth & "-" & rateDay
    '            API_BASE = "https://openexchangerates.org/"
    '            API_METHOD = "api/historical/" & rd & ".json"
    '            API_QUERY_STRING = "?app_id=44076a2ca9a243b3b61f08219fb7809f"
    '    End Select

    '    Try
    '        client.BaseAddress = New Uri(API_BASE)
    '        Dim API As String = API_METHOD & API_QUERY_STRING
    '        Dim response As HttpResponseMessage = client.GetAsync(API).Result
    '        ASCMAIN1.Progress("Fx Svc:" & API, response.StatusCode.ToString)
    '        If response.IsSuccessStatusCode Then
    '            Dim json As String = response.Content.ReadAsStringAsync().Result
    '            Select Case rateService
    '                Case "appspot"
    '                    Return Newtonsoft.Json.Linq.JObject.Parse(json).SelectToken("rate").ToString
    '                Case "ratelab"
    '                    Dim originalValue As Double = Val(Newtonsoft.Json.Linq.JObject.Parse(json).SelectToken("rate.rate").ToString & "")
    '                    If originalValue <> 0 Then
    '                        Dim reciprocalValue As Double = 1 / originalValue
    '                        Return reciprocalValue.ToString
    '                    End If
    '                Case "openexchange", "openexchangeHistory"
    '                    Dim originalValue As Double = Val(Newtonsoft.Json.Linq.JObject.Parse(json).SelectToken("rates.CAD").ToString & "")
    '                    If originalValue <> 0 Then
    '                        Dim reciprocalValue As Double = 1 / originalValue
    '                        Return reciprocalValue.ToString
    '                    End If
    '            End Select
    '        Else
    '            Return responseString
    '        End If
    '    Catch ex As Exception
    '        Return responseString
    '    End Try

    '    Return responseString
    'End Function

    Public Shared Sub Get_Unprocessed_IDOCs(frmASFBASE0 As ASFBASE0)

        Dim sftp_folder As String = "" _
            & IIf(ASCMAIN1.Running_in_VS And 1 = 1, My.Computer.FileSystem.SpecialDirectories.Desktop & "\Interparfums\", ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT")) _
            & "\IPSA\" _
            & IIf(ASCMAIN1.DBS_SERVER = "TST" Or ASCMAIN1.DBS_COMPANY = "TST", "TEST", "PROD") _
            & "\FROM_IPSA\IDOC\"

        If Not frmASFBASE0.dst.Tables.Contains("TATIDOCU") Then
            With frmASFBASE0.dst.Tables.Add("TATIDOCU")
                .Columns.Add("FILENAME")
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .Columns.Add("FILEDATE", GetType(System.DateTime))
                .Columns.Add("FILENAME_SHORT")

                .Columns.Add("INV_NUM")
                .Columns.Add("INV_DATE", GetType(System.DateTime))
                .Columns.Add("INV_AMT", GetType(System.Decimal))
                .Columns.Add("PO_ORDER_NO")
                .Columns.Add("PINV_TYPE")
                .Columns.Add("PINV_REF_INV")
                .Columns.Add("WHSE_CODE")
            End With
        End If

        frmASFBASE0.dst.Tables("TATIDOCU").Rows.Clear()
        For Each FILENAME As String In My.Computer.FileSystem.GetFiles(sftp_folder)
            Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
            Dim rowTATIDOCU As DataRow = frmASFBASE0.dst.Tables("TATIDOCU").Rows.Add(New Object() {FILENAME, FI.Length, FI.CreationTime, FI.Name})

            Using sr As New System.IO.StreamReader(FILENAME)
                Dim T As String = sr.ReadToEnd
                Dim Ts() As String = Split(T, vbCrLf)
                Dim INV_NUM As String = ""
                Dim INV_DATE As Date = Nothing
                Dim INV_AMT As Decimal = 0
                Dim PO_ORDER_NO As String = ""
                Dim PINV_TYPE As String = ""
                Dim PINV_REF_INV As String = ""

                Dim WHSE_CODE As String = ""
                If Ts.Length = 1 Then
                    Ts = Split(T, vbLf)
                End If

                For i As Integer = 0 To Ts.Length - 1
                    Dim tx As String = Ts(i)
                    If tx.StartsWith("E2EDK02") Then
                        If Mid(tx, 64, 3) = "009" Then '
                            'E2EDK02                       5000000000001212899000008000000020090120069669                               20150915      
                            INV_NUM = Trim(Mid(tx, 67, 40))
                            Dim txd As String = Mid(tx, 108, 8)
                            INV_DATE = CDate(Mid(txd, 5, 2) & "/" & Mid(txd, 7, 2) & "/" & Mid(txd, 1, 4))
                        End If
                        If Mid(tx, 64, 3) = "087" Then
                            'E2EDK02                       500000000000121289900000900000002001ILLICIT MOSAIC BOARD                     20150914      
                            'E2EDK02                       500000000000107802900001300000002087132293                                          

                            PO_ORDER_NO = Trim(Mid(tx, 67, 40))
                        End If

                        If Mid(tx, 64, 3) = "017" Then
                            'E2EDK02                       5000000000001078029000012000000020170120058620                                             
                            PINV_REF_INV = Trim(Mid(tx, 67, 40))
                        End If
                    End If

                    If tx.StartsWith("E2EDK05001") Then
                        'E2EDK05001                    500000000000121289900002100000002+                                                                                      128.80                    0                                                                    USD
                        INV_AMT = Val(Trim(Mid(tx, 141, 16)))
                    End If
                    If tx.StartsWith("E2EDK01006") Then
                        'E2EDK01005                    500000000000107802900000100000001    USDEUR0.80199     Z090                                 FR39350219382       INVO0120058620                         7076.854          7076.854          KGMLR                                                     0000200656                                                                    L                                                      1.24690     
                        'E2EDK01006                    500000000000364939500000100000001    USDEUR0.84861     Z060                                 FR39350219382       INVO0120271910                         11107.400         11669.800         KGMLR                                                     0000200656                                                                    L                                                      1.17840     FR US 
                        ' Feb 2026 E2EDK01005 -> E2EDK01006
                        PINV_TYPE = Mid(tx, 143, 4)
                        PINV_TYPE = Mid(PINV_TYPE, 1, 1)
                    End If
                Next

                If PINV_TYPE = "I" Then
                    If PINV_REF_INV <> INV_NUM Then
                        PINV_TYPE = "D"
                    End If
                End If

                Dim rowPOTORDR1 As DataRow = frmASFBASE0.LookUp("POTORDR1", PO_ORDER_NO)
                If rowPOTORDR1 IsNot Nothing Then
                    WHSE_CODE = rowPOTORDR1.Item("WHSE_CODE") & ""
                End If


                rowTATIDOCU.Item("INV_NUM") = INV_NUM
                rowTATIDOCU.Item("INV_DATE") = INV_DATE
                rowTATIDOCU.Item("INV_AMT") = INV_AMT
                rowTATIDOCU.Item("PO_ORDER_NO") = PO_ORDER_NO
                rowTATIDOCU.Item("PINV_TYPE") = PINV_TYPE
                rowTATIDOCU.Item("PINV_REF_INV") = PINV_REF_INV
                rowTATIDOCU.Item("WHSE_CODE") = WHSE_CODE

                sr.Close()
                sr.Dispose()
            End Using


            'Dim rowE2EDK02_009 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '009'")(0)
            'Dim rowE2EDK02_001 As DataRow = dst.Tables("E2EDK02").Select("QUALF = '001'")(0)
            'Dim rowE2EDS01 As DataRow = dst.Tables("E2EDS01").Select("SUMID = '010'")(0)

            'IDOC_DATA_KEY = Trim(rowE2EDK02_009.Item("BELNR"))
            'Dim YYYYMMDD As String = Trim(rowE2EDK02_009.Item("DATUM"))
            'IDOC_DATA_DATE = CDate(Mid(YYYYMMDD, 5, 2) & "/" & Mid(YYYYMMDD, 7, 2) & "/" & Mid(YYYYMMDD, 1, 4))
            'IDOC_DATA_AMT = Val(rowE2EDS01.Item("SUMME") & "")
            'IDOC_DATA_REF = Trim(rowE2EDK02_001.Item("BELNR"))


            'Absx1.txtFor("IDOC_DATA_KEY").Text = IDOC_DATA_KEY
            'Absx1.dteFor("IDOC_DATA_DATE").Value = IDOC_DATA_DATE
            'Absx1.numFor("IDOC_DATA_AMT").Value = IDOC_DATA_AMT
            'Absx1.txtFor("IDOC_DATA_REF").Text = IDOC_DATA_REF

        Next
    End Sub
    Public Shared Sub Get_Deleted_IDOCs(frmASFBASE0 As ASFBASE0)

        Dim sftp_folder As String = "" _
            & IIf(ASCMAIN1.Running_in_VS And 1 = 1, My.Computer.FileSystem.SpecialDirectories.Desktop & "\Interparfums\", ASCMAIN1.rowASTPARM1.Item("AS_PARM_SFTP_ROOT")) _
            & "\IPSA\" _
            & IIf(ASCMAIN1.DBS_SERVER = "TST" Or ASCMAIN1.DBS_COMPANY = "TST", "TEST", "PROD") _
            & "\FROM_IPSA\IDOC\Deleted\"

        If Not frmASFBASE0.dst.Tables.Contains("TATIDOCD") Then
            With frmASFBASE0.dst.Tables.Add("TATIDOCD")
                .Columns.Add("FILENAME")
                .Columns.Add("FILESIZE", GetType(System.Int64))
                .Columns.Add("FILEDATE", GetType(System.DateTime))
                .Columns.Add("FILENAME_SHORT")

                .Columns.Add("INV_NUM")
                .Columns.Add("INV_DATE", GetType(System.DateTime))
                .Columns.Add("INV_AMT", GetType(System.Decimal))
                .Columns.Add("PO_ORDER_NO")
                .Columns.Add("PINV_TYPE")
                .Columns.Add("PINV_REF_INV")
            End With
        End If

        frmASFBASE0.dst.Tables("TATIDOCD").Rows.Clear()
        For Each FILENAME As String In My.Computer.FileSystem.GetFiles(sftp_folder)
            Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
            Dim rowTATIDOCD As DataRow = frmASFBASE0.dst.Tables("TATIDOCD").Rows.Add(New Object() {FILENAME, FI.Length, FI.CreationTime, FI.Name})

            Using sr As New System.IO.StreamReader(FILENAME)
                Dim T As String = sr.ReadToEnd
                Dim Ts() As String = Split(T, vbCrLf)
                Dim INV_NUM As String = ""
                Dim INV_DATE As Date = Nothing
                Dim INV_AMT As Decimal = 0
                Dim PO_ORDER_NO As String = ""
                Dim PINV_TYPE As String = ""
                Dim PINV_REF_INV As String = ""
                Dim Cut_Off_Date As Date = Now.AddDays(-365)


                For i As Integer = 0 To Ts.Length - 1
                    Dim tx As String = Ts(i)
                    If tx.StartsWith("E2EDK02") Then
                        If Mid(tx, 64, 3) = "009" Then '
                            'E2EDK02                       5000000000001212899000008000000020090120069669                               20150915      
                            INV_NUM = Trim(Mid(tx, 67, 40))
                            Dim txd As String = Mid(tx, 108, 8)
                            INV_DATE = CDate(Mid(txd, 5, 2) & "/" & Mid(txd, 7, 2) & "/" & Mid(txd, 1, 4))
                        End If
                        If Mid(tx, 64, 3) = "087" Then
                            'E2EDK02                       500000000000121289900000900000002001ILLICIT MOSAIC BOARD                     20150914      
                            'E2EDK02                       500000000000107802900001300000002087132293                                          

                            PO_ORDER_NO = Trim(Mid(tx, 67, 40))
                        End If

                        If Mid(tx, 64, 3) = "017" Then
                            'E2EDK02                       5000000000001078029000012000000020170120058620                                             
                            PINV_REF_INV = Trim(Mid(tx, 67, 40))
                        End If
                    End If

                    If tx.StartsWith("E2EDK05001") Then
                        'E2EDK05001                    500000000000121289900002100000002+                                                                                      128.80                    0                                                                    USD
                        INV_AMT = Val(Trim(Mid(tx, 141, 16)))
                    End If
                    If tx.StartsWith("E2EDK01006") Then
                        'E2EDK01005                    500000000000107802900000100000001    USDEUR0.80199     Z090                                 FR39350219382       INVO0120058620                         7076.854          7076.854          KGMLR                                                     0000200656                                                                    L                                                      1.24690     
                        'E2EDK01006                    500000000000364939500000100000001    USDEUR0.84861     Z060                                 FR39350219382       INVO0120271910                         11107.400         11669.800         KGMLR                                                     0000200656                                                                    L                                                      1.17840     FR US 
                        PINV_TYPE = Mid(tx, 143, 4)
                        PINV_TYPE = Mid(PINV_TYPE, 1, 1)
                    End If
                Next

                If PINV_TYPE = "I" Then
                    If PINV_REF_INV <> INV_NUM Then
                        PINV_TYPE = "D"
                    End If
                End If

                rowTATIDOCD.Item("INV_NUM") = INV_NUM
                rowTATIDOCD.Item("INV_DATE") = INV_DATE
                rowTATIDOCD.Item("INV_AMT") = INV_AMT
                rowTATIDOCD.Item("PO_ORDER_NO") = PO_ORDER_NO
                rowTATIDOCD.Item("PINV_TYPE") = PINV_TYPE
                rowTATIDOCD.Item("PINV_REF_INV") = PINV_REF_INV

                If INV_DATE < Cut_Off_Date Then
                    rowTATIDOCD.Delete()
                End If

                sr.Close()
                sr.Dispose()
            End Using

        Next
    End Sub

    Public Shared Function Send_alert_email(FF As ASFBASE0, ALERT_EMAIL As String, USER_ID As String, MENU_ITEM As String, ALERT_MESSAGE As String)

        Dim ALERT_SUBJECT As String = ""
        Dim TABLE_NAME As String = ""
        Dim EMAIL_MSG As String = ""
        Dim EVENT_CODE As String = ""

        Dim rowTATALRT1 As DataRow = FF.dst.Tables("TATALRT1").NewRow
        With rowTATALRT1
            Dim ALERT_NO As String = ASCMAIN1.Next_Control_No("TATALRT1.ALERT_NO")
            .Item("ALERT_NO") = ALERT_NO
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("INIT_DATE") = FF.DATETIME_STAMP
            .Item("FORM_NAME") = "ASTUSER1"
            .Item("FORM_KEY") = ALERT_NO
            .Item("ALERT_EMAIL") = ALERT_EMAIL
            .Item("ALERT_EML") = "1"

            .Item("ALERT_EML_DATE") = FF.DATETIME_STAMP
            If MENU_ITEM <> "" Then
                ALERT_SUBJECT = "Security Change to Menu Item " & MENU_ITEM
                TABLE_NAME = "ASTMENU1"
                EMAIL_MSG = "Menu Security Change Alert emailed to " & ALERT_EMAIL
                EVENT_CODE = "MENUSEC"
            Else
                ALERT_SUBJECT = "Security Change to User " & USER_ID
                TABLE_NAME = "ASTUSER1"
                EMAIL_MSG = "User Security Change Alert emailed to " & ALERT_EMAIL
                EVENT_CODE = "USERSEC"
            End If

            .Item("ALERT_SUBJECT") = Mid(ALERT_SUBJECT, 1, 200)
            .Item("ALERT_MESSAGE") = Mid(ALERT_MESSAGE, 1, 2000)
        End With
        FF.dst.Tables("TATALRT1").Rows.Add(rowTATALRT1)

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
        EMAIL_ADDRESSs.Add(ALERT_EMAIL, "Security Auditor")

        Dim SEND_NO As String = ""
        If ASCMAIN1.Running_in_VS Then
            SEND_NO = "TESTING"
            Stop
        Else
            If MENU_ITEM <> "" Then
                SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                  (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing,
                  ALERT_SUBJECT, "AS_MENUSEC", True, False, USER_ID, USER_ID, "Menu", ALERT_MESSAGE)
            Else
                SEND_NO = ASCMAIN1.TACMAIN1.Send_email _
                                  (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, Nothing,
                                  ALERT_SUBJECT, "AS_USERSEC", True, False, USER_ID, USER_ID, "User ID", ALERT_MESSAGE)
            End If

        End If

        rowTATALRT1.Item("SEND_NO") = SEND_NO
        FF.Update_Record_TDA("TATALRT1")

        TAC.TACMAIN1.Record_Event(TABLE_NAME, USER_ID, FF.DATETIME_STAMP,
                                  ASCMAIN1.USER_ID, EVENT_CODE, EMAIL_MSG, SEND_NO, TABLE_NAME)

        Return SEND_NO
    End Function

    Public Shared Function Generate_Audit_Workbook(FF As ASFBASE0, startDate As Date, endDate As Date) As String
        Dim FILENAME As String = ""

        Dim oWB As SpreadsheetGear.IWorkbook

        Dim xls_path As String = ASCMAIN1.Folders("Work")
        Dim xls_name As String = ""

        Dim success As Boolean = False
        Dim XLS_NO As Integer = 0

        Do Until success
            Try
                XLS_NO += 1
                xls_name = "Audit_Report_" & ASCMAIN1.DBS_SESSION_ID
                xls_name &= "-" & Format(XLS_NO, "000") & ".xlsx"
                FILENAME = xls_path & "\" & xls_name & ".XLSx"

                If Not My.Computer.FileSystem.FileExists(FILENAME) Then
                    success = True
                End If
            Catch ex As Exception
                Stop
            End Try
        Loop

        oWB = SpreadsheetGear.Factory.GetWorkbook()

        For i As Integer = oWB.Worksheets.Count To 2 Step -1
            oWB.Worksheets(i).Delete()
        Next i

        For Each AUDIT_WS As String In New String() {"ASTOPST1", "ASTPARM1", "ASTPARMP", "ASTSECM1", "TATUSER1", "ASTUSER1"}
            Get_Audit_Report_Worksheet(oWB, AUDIT_WS, startDate, endDate)
        Next

        oWB.Worksheets("Sheet1").Delete()

        oWB.Worksheets(0).Select()

        oWB.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        oWB = Nothing

        Return FILENAME
    End Function

    Public Shared Sub Get_Audit_Report_Worksheet(owb As SpreadsheetGear.IWorkbook, key As String, startDate As Date, endDate As Date)

        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        Dim range As SpreadsheetGear.IRange = Nothing
        Dim rangePaste_To As SpreadsheetGear.IRange

        Dim sqlAudit As String = ""
        Dim tbl As DataTable = Nothing
        Dim sheetName As String = ""

        Select Case key
            Case "ASTOPST1"
                sqlAudit = "SELECT ASTOPST1.USER_ID, ASTOPST1.INIT_DATE START_DATE, LPAD(ASTOPST1.SESSION_NO, 10,'0') SESSION_NO, ASTOPST1.SELECTION_NO" & vbCrLf _
                    & " , ASTOPST1.RE_XNO, ASTOPST1.MENU_ID, ASTOPST1.MENU_ITEM_TYPE, ASTOPST1.MENU_ITEM_OBJECT, ASTMENU1.MENU_ITEM_DESC" & vbCrLf _
                    & " , ASTOPST1.INIT_DATE, ASTOPST1.LAST_DATE, ASTOPST1.YYYYPP, ASTOPST1.PRD_CLOSE_IND, LPAD(ASTOPST1.FORM_INSTANCE_NO,10,'0') FORM_INSTANCE_NO " & vbCrLf _
                    & " FROM ASTOPST1, ASTMENU1 WHERE ASTOPST1.MENU_ID = ASTMENU1.MENU_ID AND ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE" & vbCrLf _
                    & " AND ASTOPST1.MENU_ITEM_TYPE = ASTMENU1.MENU_ITEM_TYPE AND ASTOPST1.INIT_DATE >= :PARM1 And ASTOPST1.LAST_DATE <= :PARM2"
                sheetName = "Operator Statistics"
            Case "ASTPARM1"
                sqlAudit = "SELECT COLUMN_NAME DATA_FIELD, INIT_DATE CHANGED, USER_ID" & vbCrLf _
                    & ", OLD_VALUE, NEW_VALUE, SESSION_NO, FM_MODE FROM ASTAUDT1 WHERE TABLE_NAME = 'ASTPARM1'"
                sheetName = "System Parameters"
            Case "ASTPARMP"
                sqlAudit = "SELECT COLUMN_NAME DATA_FIELD, INIT_DATE CHANGED, USER_ID" & vbCrLf _
                    & ", OLD_VALUE, NEW_VALUE, SESSION_NO, FM_MODE FROM ASTAUDT1 WHERE TABLE_NAME = 'ASTPARMP'"
                sheetName = "Password Parameters"
            Case "ASTSECM1"
                sqlAudit = "SELECT KEY_VALUE, COLUMN_NAME DATA_FIELD, INIT_DATE CHANGED, USER_ID" & vbCrLf _
                    & ", OLD_VALUE, NEW_VALUE, SESSION_NO, FM_MODE FROM ASTAUDT1 WHERE TABLE_NAME = 'ASTSECM1'"
                sheetName = "Security Codes"
            Case "TATUSER1"
                sqlAudit = "SELECT KEY_VALUE, COLUMN_NAME DATA_FIELD, INIT_DATE CHANGED, USER_ID" & vbCrLf _
                    & ", OLD_VALUE, NEW_VALUE, SESSION_NO, FM_MODE FROM ASTAUDT1 WHERE TABLE_NAME = 'TATUSER1'"
                sheetName = "User Profiles"
            Case "ASTUSER1"
                sqlAudit = "SELECT KEY_VALUE, COLUMN_NAME DATA_FIELD, INIT_DATE CHANGED, USER_ID" & vbCrLf _
                    & ", DECODE(COLUMN_NAME,'USER_PASSWORD','**********',OLD_VALUE) OLD_VALUE" & vbCrLf _
                    & ", DECODE(COLUMN_NAME,'USER_PASSWORD','**********',NEW_VALUE) NEW_VALUE" & vbCrLf _
                    & ", SESSION_NO, FM_MODE FROM ASTAUDT1 WHERE TABLE_NAME = 'ASTUSER1'"
                sheetName = "Users"
        End Select

        tbl = ASCDATA1.GetDataTable(sqlAudit, "ASTOPST1", "DD", New Object() {startDate, endDate})
        oSheet = owb.Worksheets.Add()

        rangePaste_To = oSheet.Cells(0, 0, 2, 0)
        rangePaste_To.CopyFromDataTable(tbl, SpreadsheetGear.Data.SetDataFlags.InsertCells)

        Format_XLS_based_on_tbl(tbl, oSheet)

        oSheet.UsedRange.Columns.AutoFit()

        range = oSheet.Cells(1, 0)
        range.Select()
        oSheet.WindowInfo.FreezePanes = True

        For C As Integer = 0 To tbl.Columns.Count - 1
            oSheet.Range(0, C).EntireColumn.AutoFilter()
            oSheet.Range(0, C).Interior.Color = SpreadsheetGear.Colors.LightGray
        Next

        oSheet.Name = sheetName


    End Sub
    Public Shared Sub Format_XLS_based_on_tbl(tbl As DataTable, worksheet As SpreadsheetGear.IWorksheet, Optional C_offset As Integer = 0)

        For C As Integer = 0 To tbl.Columns.Count - 1
            Dim DC As DataColumn = tbl.Columns(C)
            If DC.DataType.Name = "String" Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "@"
            ElseIf DC.DataType.Name = "DateTime" Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "MM/DD/YY"
                'worksheet.Range(0, C + C_offset).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.CenterFGE
            ElseIf DC.DataType.Name = "Decimal" Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "#,##0.00"
                worksheet.Range(0, C + C_offset).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            ElseIf DC.DataType.Name.StartsWith("Int") Then
                worksheet.Range(0, C + C_offset).EntireColumn.NumberFormat = "#,##0"
                worksheet.Range(0, C + C_offset).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Right
            End If
        Next
    End Sub

    Public Shared Sub GetEmails(FF As ASFBASE0)

        Dim AS_PARM_EMAIL_USER_ID As String = ""
        Dim AS_PARM_EMAIL_PASSWORD As String = ""
        Dim AS_PARM_EMAIL_DOMAIN = "" ' "interparfums.com"

        AS_PARM_EMAIL_USER_ID = FF.ROWs("APTPARM1").Item("AP_PARM_INV_EMAIL_USER_ID") & ""
        AS_PARM_EMAIL_PASSWORD = FF.ROWs("APTPARM1").Item("AP_PARM_INV_EMAIL_PASSWORD") & ""

        Dim service As ExchangeService = TACMAIN1.Get_EWS_Service(AS_PARM_EMAIL_USER_ID)

        ' test send --------------------------------------------------------------------------------------

        If ASCMAIN1.Running_in_VS And False Then


            Dim Message As EmailMessage = New EmailMessage(service)

            Dim SEND_FROM As String = "lmarinelli@interparfums.com"
            SEND_FROM = "joeuser@absolution.com"
            Dim SEND_FROM_NAME As String = "Joe User at INT"

            'Message.ToRecipients.Add(New EmailAddress("Walter", "wjz@absolution.com"))
            Message.ToRecipients.Add(New EmailAddress("Joe User", "joeuser@absolution.com"))

            Message.From = New EmailAddress(SEND_FROM_NAME, SEND_FROM)
            Message.Subject = "Test email " & Format(Now, "HH:mm:ss tt")

            Message.Body = "<p>this is the message sent</p>"

            Dim folder As String = ASCMAIN1.Folders("Archive") & "email\Sent\"
            If Not My.Computer.FileSystem.DirectoryExists(folder) Then
                My.Computer.FileSystem.CreateDirectory(folder)
            End If

            Message.Save()

            Dim SEND_NO As String = ASCMAIN1.Next_Control_No("TATSEND1.SEND_NO")
            Try
                Message.SaveToFile(folder & SEND_NO & ".eml")
            Catch ex As Exception
                If ASCMAIN1.Running_in_VS Then
                    Stop
                End If
                MsgBox(ex.Message, MsgBoxStyle.OkOnly, $"Error occured sending an email to Inbox of {AS_PARM_EMAIL_USER_ID}")
            End Try

            Message.SendAndSaveCopy()
        End If

        ' ---------------------------------------------------------------------------------------

        Dim eMailItems As FindItemsResults(Of Item) = service.FindItems(New FolderId(WellKnownFolderName.Inbox), New ItemView(15))
        If eMailItems.Count > 0 Then
            service.LoadPropertiesForItems(eMailItems, PropertySet.FirstClassProperties)

            Dim customPropertySet = New PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Subject, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.From)

            Dim folderArchive As String = ASCMAIN1.Folders("Archive") & "SubmittedInvoices\"
            If Not My.Computer.FileSystem.DirectoryExists(folderArchive) Then
                My.Computer.FileSystem.CreateDirectory(folderArchive)
            End If

            For Each msg As EmailMessage In eMailItems
                Dim SUBMIT_CTL_NO As String = ASCMAIN1.Next_Control_No("APTSUBM1.SUBMIT_CTL_NO")
                Try
                    msg.SaveToFile(folderArchive & SUBMIT_CTL_NO & ".eml")
                    msg.Load(customPropertySet) ' note - you must re-Load the properties after .SaveToFile
                    Dim SUBMIT_EMAIL_FROM As String = msg.From.Address
                    Dim SUBMIT_SUBJECT As String = msg.Subject
                    Dim SUBMIT_DATE_RECEIVED As Date = msg.DateTimeReceived
                    ASCMAIN1.sql = "Insert into APTSUBM1" & vbCrLf _
                        & "(SUBMIT_CTL_NO,SUBMIT_EMAIL_FROM,SUBMIT_SUBJECT,SUBMIT_DATE_RECEIVED,SUBMIT_STATUS,INIT_DATE,INIT_OPER)" & vbCrLf _
                        & " values " & vbCrLf _
                        & "(:PARM1,:PARM2,:PARM3,:PARM4,'U',SYSDATE,'service')"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVD", New Object() {SUBMIT_CTL_NO, SUBMIT_EMAIL_FROM, SUBMIT_SUBJECT, SUBMIT_DATE_RECEIVED})
                    msg.Delete(DeleteMode.MoveToDeletedItems)

                Catch ex As Exception
                    If ASCMAIN1.Running_in_VS Then
                        Stop
                    End If
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, $"Error occured retrieving Inbox of {AS_PARM_EMAIL_USER_ID}")
                End Try
            Next
        End If
    End Sub

    Public Shared Function Get_EWS_Service(USER_EMAIL As String) As ExchangeService

        If ews_service IsNot Nothing Then
            Return ews_service
        End If

        Dim service As ExchangeService = New ExchangeService()

        Try
            'Dim authResult As AuthenticationResult = Await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync()
            Dim authResult As AuthenticationResult = ASCMAIN1.authResult
            'MsgBox("token: " & authResult.AccessToken)
            service.Url = New Uri("https://outlook.office365.com/EWS/Exchange.asmx")
            service.Credentials = New OAuthCredentials(authResult.AccessToken)
            service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.SmtpAddress, USER_EMAIL)
            service.HttpHeaders.Add("X-AnchorMailbox", USER_EMAIL)

            Debug.Print(ASCMAIN1.authResult.ToString)
            Debug.Print(ASCMAIN1.authResult_counter.ToString)
            Debug.Print(ASCMAIN1.authResult_timestamp.ToString)

        Catch ex As Exception
            If ASCMAIN1.Running_in_VS Then Stop
            MsgBox("error: " & ex.Message)
        End Try

        Return service

    End Function

    Public Shared Function Validate_Address(a As String) As String
        Dim authId As String = "6805b5a1-f2af-524f-89d3-ce30982908d9"
        Dim authToken As String = "wsH9lb5377MhzzgO6O2A"

        Dim client As SmartyStreets.USExtractApi.Client = New SmartyStreets.ClientBuilder(authId, authToken).BuildUsExtractApiClient()

        Dim lookup As SmartyStreets.USExtractApi.Lookup = New SmartyStreets.USExtractApi.Lookup(a) With {
        .IsAggressive = True,
        .AddressesHaveLineBreaks = False,
        .AddressesPerLine = 1
    }

        client.Send(lookup)

        Dim result = lookup.Result
        Dim metadata = result.Metadata
        Dim ADDRESS_RETURNED As String = ""
        Console.WriteLine("Found " & metadata.AddressCount & " addresses.")
        Console.WriteLine(metadata.VerifiedCount & " of them were valid.")
        Console.WriteLine()
        Dim addresses = result.Addresses
        Console.WriteLine("Addresses: " & vbCrLf & "**********************" & vbCrLf)

        For Each address In addresses
            If ADDRESS_RETURNED = "" Then ADDRESS_RETURNED = address.Text
            Console.WriteLine("""" & address.Text & """" & vbLf)
            Console.WriteLine("Verified? " & address.Verified)

            If address.Candidates.Length > 0 Then
                Console.WriteLine(vbLf & "Matches:")

                For Each candidate In address.Candidates
                    Console.WriteLine(candidate.DeliveryLine1)
                    Console.WriteLine(candidate.LastLine)
                    Console.WriteLine()
                Next
            Else
                Console.WriteLine()
            End If

            Console.WriteLine("**********************" & vbLf)
        Next


        Return ADDRESS_RETURNED
    End Function
    Public Shared Function Validate_Address1(addressString As String)
        ' Set up SmartyStreets authentication
        Dim authID = "e7af1ad5-11d8-47ce-cae1-9377b80e1b64"
        Dim authToken = "ak93k6SF7sZXki1KTzam"

        Dim client = New SmartyStreets.ClientBuilder(authID, authToken).WithLicense(New List(Of String) From {"us-core-cloud"}).BuildUsStreetApiClient()

        ' Parse the input address string
        Dim lines = addressString.Split(vbCrLf)
        If lines.Length < 3 Then
            Return "Invalid address format."
        End If

        ' Extract components from the input string
        Dim addressee = lines(0).Trim()
        Dim street = lines(1).Trim()
        Dim cityStateZip = lines(2).Trim().Split(" ")

        Dim city = String.Join(" ", cityStateZip.Take(cityStateZip.Length - 2))
        Dim state = cityStateZip(cityStateZip.Length - 2)
        Dim zipCode = cityStateZip.Last()

        ' Set up lookup object
        Dim lookup As New Lookup()
        With lookup
            .Addressee = addressee
            .Street = street
            .City = city
            .State = state
            .ZipCode = zipCode
            .MaxCandidates = 1
            .MatchStrategy = Lookup.ENHANCED
        End With

        Console.WriteLine("*******************************************************")
        Console.WriteLine()

        ' Send lookup request
        Try
            client.Send(lookup)
        Catch ex As SmartyException
            Return "Error: " & ex.Message
        Catch ex As IOException
            Return "Error: " & ex.Message
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

        ' Process candidates
        Dim candidates = lookup.Result
        If candidates.Count = 0 Then
            Return "No candidates. The address is not valid."
        End If

        ' Construct a summary string
        Dim resultSummary As String = $"Original lookup: {street}, {city}, {state}, {zipCode}{Environment.NewLine}"
        resultSummary &= $"Address has {candidates.Count} candidate" & If(candidates.Count = 1, "", "s") & Environment.NewLine
        resultSummary &= $"Input ID: {lookup.InputId}{Environment.NewLine}"


        For Each candidate In candidates
            resultSummary &= Environment.NewLine
            Dim components = candidate.Components
            Dim metadata = candidate.Metadata

            resultSummary &= $"Candidate {candidate.CandidateIndex}:{Environment.NewLine}"
            resultSummary &= $"Delivery line 1: {candidate.DeliveryLine1}{Environment.NewLine}"
            resultSummary &= $"Last line:       {candidate.LastLine}{Environment.NewLine}"
            resultSummary &= $"ZIP Code:        {components.ZipCode}-{components.Plus4Code}{Environment.NewLine}"
            resultSummary &= $"County:          {metadata.CountyName}{Environment.NewLine}"
            resultSummary &= $"Latitude:        {metadata.Latitude}{Environment.NewLine}"
            resultSummary &= $"Longitude:       {metadata.Longitude}{Environment.NewLine}"
        Next

        Return resultSummary
    End Function








    Public Shared Async Sub Check_for_emails(FF As ASFBASE0)

        ' ITERATE THRU INBOX
        ' SAVE EACH EML FILE
        ' RIP ATTACHMENTS AND SAVE EACH PDF FILE
        ' WRITE A RECORD FOR SOCORDE2 (TO CONVERT PDFS TO XLS AND THEN TO CREATE ORDERS) - DO THIS ONLY IF THERE WERE EMAILS REQUIRING PROCESSING

        Dim SO_PARM_EMAIL_ORDERS_USER_ID As String = ""
        Dim SO_PARM_EMAIL_ORDERS_PASSWORD As String = ""
        Dim SO_PARM_EDOC_FOLDER As String = ""
        ' Dim AS_PARM_EMAIL_DOMAIN = "" ' "interparfums.com"

        Dim rowSOTPARM1 As DataRow = FF.LookUp("SOTPARM1", "Z")
        SO_PARM_EMAIL_ORDERS_USER_ID = rowSOTPARM1.Item("SO_PARM_EMAIL_ORDERS_USER_ID") & ""
        SO_PARM_EMAIL_ORDERS_PASSWORD = rowSOTPARM1.Item("SO_PARM_EMAIL_ORDERS_PASSWORD") & ""
        SO_PARM_EDOC_FOLDER = rowSOTPARM1.Item("SO_PARM_EDOC_FOLDER") & ""

        ' Dim service As ExchangeService = TACMAIL1.Get_EWS_Service(SO_PARM_EMAIL_ORDERS_USER_ID)
        Dim service As ExchangeService = Nothing ' Await ASCMAIN1.Get_EWS_Service(SO_PARM_EMAIL_ORDERS_USER_ID, exchangeCredentials)
        'Dim exchangeCredentials As ABSSVC.EWSCredentials

        Try
            service = Await Check_emails_Get_EWS_Service(SO_PARM_EMAIL_ORDERS_USER_ID)

        Catch ex As Exception
            Stop

        End Try

        Dim eMailItems As FindItemsResults(Of Item) = service.FindItems(New FolderId(WellKnownFolderName.Inbox), New ItemView(15))
        If eMailItems.Count > 0 Then
            service.LoadPropertiesForItems(eMailItems, PropertySet.FirstClassProperties)

            Dim customPropertySet = New PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Subject, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.From)

            For Each msg As EmailMessage In eMailItems
                Dim EDOC_NO As String = ASCMAIN1.Next_Control_No("SOTORDRE1.EDOC_NO")
                Try
                    Dim FOLDER As String = SO_PARM_EDOC_FOLDER & "EML\"
                    msg.SaveToFile(FOLDER & EDOC_NO & ".eml")
                    msg.Load(customPropertySet) ' note - you must re-Load the properties after .SaveToFile

                    Dim EDOC_DATE_PROCESSED As DateTime = Now

                    Dim EDOC_FROM As String = msg.From.Address
                    Dim EDOC_TO As String = Join(msg.ToRecipients.ToArray, ",")
                    Dim EDOC_SUBJECT As String = msg.Subject
                    Dim EDOC_DATE_SENT As Date = msg.DateTimeSent
                    Dim EDOC_DATE_RECEIVED As Date = msg.DateTimeReceived
                    Dim EDOC_BODY As String = msg.Body
                    Dim EDOC_STATUS As String = "P"
                    Dim EDOC_STATUS_MESSAGE As String = "Pending Conversion"
                    Dim EDOC_STATUS_DATE As DateTime = EDOC_DATE_PROCESSED
                    Dim ATTACHMENTS As Int32 = msg.Attachments.Count
                    Dim CUST_CODE As String = ""

                    ASCMAIN1.sql = "Insert into SOTORDE1 (EDOC_NO,EDOC_DATE_PROCESSED,EDOC_DATE_RECEIVED,EDOC_DATE_SENT,EDOC_FROM,EDOC_TO,EDOC_SUBJECT,EDOC_BODY,
EDOC_STATUS,EDOC_STATUS_MESSAGE,EDOC_STATUS_DATE,ATTACHMENTS,CUST_CODE
) Values (:PARM1, SYSDATE, :PARM2, :PARM3, :PARM4, :PARM5, :PARM6, :PARM7, :PARM8, :PARM9, SYSDATE, :PARM10, :PARM11)"
                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVDDVVVVVNV", New Object() {EDOC_NO, EDOC_DATE_RECEIVED, EDOC_DATE_SENT, EDOC_FROM, EDOC_TO, EDOC_SUBJECT, EDOC_BODY, EDOC_STATUS, EDOC_STATUS_MESSAGE, ATTACHMENTS, CUST_CODE})

                    msg.Delete(DeleteMode.MoveToDeletedItems)

                    Dim EDOC_LNO As Int32 = 0
                    For Each attachment In msg.Attachments
                        If TypeOf (attachment) Is FileAttachment Then

                            EDOC_LNO += 1
                            Dim FILENAME As String = EDOC_NO & "-" & Format(EDOC_LNO, "000") & ".PDF"

                            Dim f As FileAttachment = DirectCast(attachment, FileAttachment)
                            f.Load(SO_PARM_EDOC_FOLDER & "PDF\" & FILENAME)

                            Dim EDOC_FILENAME As String = f.FileName
                            Dim EDOC_NAME As String = f.Name

                            ASCMAIN1.sql = "Insert into SOTORDE2 (EDOC_NO, EDOC_LNO, ORDR_GROUP_NO, EDOC_FILENAME, EDOC_NAME) Values (:PARM1, :PARM2, :PARM3, :PARM4, :PARM5)"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VNVVVV", New Object() {EDOC_NO, EDOC_LNO, "", EDOC_FILENAME, EDOC_NAME})

                        End If
                    Next

                Catch ex As Exception
                    If Debugger.IsAttached Then
                        Stop
                    End If
                    MsgBox(ex.Message, MsgBoxStyle.OkOnly, $"Error occured retrieving Inbox of {SO_PARM_EMAIL_ORDERS_USER_ID}")
                End Try
            Next
        End If

    End Sub

    Public Shared Async Function Check_emails_Get_EWS_Service(USER_EMAIL As String) As Task(Of ExchangeService)

        Dim service As ExchangeService = New ExchangeService()

        Try
            Dim authResult As AuthenticationResult = Await Check_emails_Get_Auth_Result()

            service.Url = New Uri("https://outlook.office365.com/EWS/Exchange.asmx")
            service.Credentials = New OAuthCredentials(authResult.AccessToken)
            service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.SmtpAddress, USER_EMAIL)
            service.HttpHeaders.Add("X-AnchorMailbox", USER_EMAIL)

        Catch ex As Exception
            ' _logger.LogError(ex.Message)
            If Debugger.IsAttached Then
                Stop
            End If
        End Try

        Return service

    End Function


    Public Shared Async Function Check_emails_Get_Auth_Result() As Task(Of AuthenticationResult)

        Dim appID As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_APP_ID") & ""
        Dim tenantID As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_TENANT_ID") & ""
        Dim clientSecret As String = ASCMAIN1.rowASTPARM1.Item("AS_PARM_CLIENT_SECRET") & ""
        appID = ConfigurationManager.AppSettings("appId")
        clientSecret = ConfigurationManager.AppSettings("clientSecret")
        tenantID = ConfigurationManager.AppSettings("tenantId")

        'Dim cca As IConfidentialClientApplication = ConfidentialClientApplicationBuilder _
        '.Create(ExchangeCredentials.clientId) _
        '.WithClientSecret(ExchangeCredentials.clientSecret) _
        '.WithTenantId(ExchangeCredentials.tenantId) _
        '.Build()
        Dim cca As IConfidentialClientApplication = ConfidentialClientApplicationBuilder _
            .Create(appID) _
                .WithClientSecret(clientSecret) _
                .WithTenantId(tenantID) _
                .Build()

        Dim ewsScopes As String() = New String() {"https://outlook.office365.com/.default"}

        'System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12

        Dim authResult As AuthenticationResult = Nothing
        Try
            authResult = Await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync()
        Catch ex As Exception
            Stop
        End Try

        Return authResult

    End Function
    Public Shared Sub Create_Pivot(frm As ASFBASE0, DT As DataTable, Optional PivotColumns() As String = Nothing, Optional columnTypes As String = Nothing, Optional columnFormats As Dictionary(Of String, (String, String)) = Nothing, Optional includeTotals As Boolean = True)
        ASCMAIN1.Progress("Now Creating Workbook")

        Dim filename As String
        If ASCMAIN1.Running_in_VS Then
            filename = "C:\Share\INT\Templates\PivotTemplate.xlsx"
        Else
            filename = Path.Combine(ASCMAIN1.Folders("SharedRoot"), "Templates", "PivotTemplate.xlsx")
        End If

        Dim excel As New Microsoft.Office.Interop.Excel.Application
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = excel.Workbooks.Open(filename)
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets("Data")
        Dim totalRows As Integer = DT.Rows.Count + 1 ' Including header
        Dim totalCols As Integer = DT.Columns.Count
        Dim dataArray(totalRows - 1, totalCols - 1) As Object

        For col As Integer = 0 To totalCols - 1
            dataArray(0, col) = DT.Columns(col).ColumnName
        Next

        For rowIdx As Integer = 0 To DT.Rows.Count - 1
            For col As Integer = 0 To totalCols - 1
                dataArray(rowIdx + 1, col) = If(IsDBNull(DT.Rows(rowIdx)(col)), "", DT.Rows(rowIdx)(col))
            Next
        Next

        Dim startCell As Microsoft.Office.Interop.Excel.Range = ws.Cells(1, 1)
        Dim endCell As Microsoft.Office.Interop.Excel.Range = ws.Cells(totalRows, totalCols)
        Dim writeRange As Microsoft.Office.Interop.Excel.Range = ws.Range(startCell, endCell)
        writeRange.Value2 = dataArray

        For col As Integer = 0 To totalCols - 1
            If DT.Columns(col).DataType Is GetType(DateTime) Then
                Dim dateColumnRange As Microsoft.Office.Interop.Excel.Range = ws.Range(ws.Cells(2, col + 1), ws.Cells(totalRows, col + 1))
                dateColumnRange.NumberFormat = "mm/dd/yyyy"
            End If
        Next

        wb.Names.Item("PivotBase").RefersTo = "='Data'!" & writeRange.Address

        Dim wsPivot As Microsoft.Office.Interop.Excel.Worksheet = wb.Worksheets(1)
        Dim pivotTable As Microsoft.Office.Interop.Excel.PivotTable = wsPivot.PivotTables(1)

        pivotTable.PivotCache.Refresh()
        pivotTable.ClearTable()

        For i As Integer = 0 To PivotColumns.Length - 1
            Dim column As String = PivotColumns(i)
            Dim roleChar As Char = columnTypes(i)
            Dim pivotField As Microsoft.Office.Interop.Excel.PivotField = pivotTable.PivotFields(column)
            Dim headerText As String = column
            Dim format As String = ""
            If columnFormats IsNot Nothing AndAlso columnFormats.ContainsKey(column) Then
                Dim formatInfo = columnFormats(column)
                headerText = formatInfo.Item1
                format = formatInfo.Item2
            End If

            With pivotField
                Select Case roleChar
                    Case "A"c ' RowField
                        .Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField
                        .Position = i + 1
                        .Subtotals(1) = includeTotals AndAlso i = 0
                        If Not String.IsNullOrEmpty(headerText) Then .Caption = headerText
                    Case "B"c ' DataField
                        .Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlDataField
                        .Function = Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum
                        .NumberFormat = format
                    Case "C"c ' ColumnField
                        .Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlColumnField
                        .Position = i + 1
                        .Subtotals(1) = includeTotals
                        If Not String.IsNullOrEmpty(headerText) Then .Caption = headerText
                    Case "D"c ' PageField
                        .Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlPageField
                        .Position = i + 1
                        If Not String.IsNullOrEmpty(headerText) Then .Caption = headerText
                    Case Else
                        Throw New ArgumentException($"Invalid role character '{roleChar}' at position {i}.")
                End Select
            End With
        Next

        With pivotTable
            .RowGrand = includeTotals
            .ColumnGrand = includeTotals
            .RowAxisLayout(If(includeTotals, Microsoft.Office.Interop.Excel.XlLayoutRowType.xlTabularRow, Microsoft.Office.Interop.Excel.XlLayoutRowType.xlCompactRow))
            .DisplayNullString = True
            .NullString = ""
        End With

        If includeTotals Then
            Dim firstPivotColumn As Integer = pivotTable.TableRange1.Column
            Dim lastPivotColumn As Integer = pivotTable.TableRange1.Column + pivotTable.TableRange1.Columns.Count - 1
            Dim subtotalRows As New HashSet(Of Integer)
            For Each cell As Microsoft.Office.Interop.Excel.Range In pivotTable.TableRange1.Cells
                If cell.PivotCell IsNot Nothing Then
                    If cell.PivotCell.PivotCellType = Microsoft.Office.Interop.Excel.XlPivotCellType.xlPivotCellSubtotal Then
                        subtotalRows.Add(cell.Row)
                    End If
                End If
            Next
            For Each rowNumber In subtotalRows
                Dim rowRange As Microsoft.Office.Interop.Excel.Range = wsPivot.Range(wsPivot.Cells(rowNumber, firstPivotColumn), wsPivot.Cells(rowNumber, lastPivotColumn))
                rowRange.Interior.Color = RGB(211, 211, 211)
            Next
        End If

        wsPivot.Columns.AutoFit()

        ASCMAIN1.Progress("Now Saving Workbook")
        Dim XLS_FILENAME As String = $"Pivot Export-{Now:yyyyMMddHHmmss}.xlsx"
        Dim fullPath As String = Path.Combine(ASCMAIN1.Folders("Work"), XLS_FILENAME)

        wb.SaveAs(fullPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
        wb.Close(False)
        excel.Quit()

        frm.ReleaseCOMObject(writeRange)
        frm.ReleaseCOMObject(ws)
        frm.ReleaseCOMObject(wb)
        frm.ReleaseCOMObject(excel)

        frm.Show_Document(fullPath)
        ASCMAIN1.Progress("")
    End Sub
End Class

