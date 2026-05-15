Public Class WHC999O1
    ' Create Customer Master Changes and Additions File (Outbound) 999

    Inherits WHC000O1

    Dim TMP_CSMST As String
    Dim CSMST As String = "CSMST" & "_" & DTS & ".CSV"
    Sub New()
        MyBase.New()
    End Sub
    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHC999O1"

        Create_Work_Table()

        With dst
            Dim sqlCols As String = "" _
                & "CUSTNAME,ADDRESS1,ADDRESS2,CITY,STATE,ZIPCODE,ZIPCODE2,PHONE,CONTACT," _
                & "AR_PHONE,AR_CONTACT,CARCODE,EDISTORE#,IPLBCUST1,IPLBCUST2,INVWSHIP,FDXACCOUNT,UPSACCOUNT,IPLBCONT," _
                & "CLRCUST1,CLRCUST2,COUNTRY,STATUS,CLRCUSTS1,CLRCUSTS2,IPLBCUSTS1,IPLBCUSTS2,DFCYN"
            sqlCols = Replace(Replace(Replace(sqlCols, ",CLRCUST1,CLRCUST2", ""), ",CLRCUSTS1,CLRCUSTS2", ""), ",IPLBCUST2,", ",NVL(IPLBCUST2,'?') IPLBCUST2,")

            ' ASCDATA1.ExecuteSQL("Update " & TMP_CSMST & " Set IPLBCUST2 = NVL(W.IPLBCUST2,'?') where IPLBCUST2 IS NULL")

            ASCMAIN1.sql = "Select W.* from " & TMP_CSMST & " W," & vbCrLf _
                & "(Select " & sqlCols & " from " & TMP_CSMST & vbCrLf _
                & "  minus " & vbCrLf _
                & " Select " & sqlCols & " from CONV.CFG_CSMST) X" & vbCrLf _
                & " where W.IPLBCUST1 = X.IPLBCUST1 and NVL(W.IPLBCUST2,'?') = X.IPLBCUST2"

            Create_TDA(.Tables.Add, "CSMST", "**", 0, False)

        End With

        '                & " where (IPLBCUST1,NVL(IPLBCUST2,'?')) in (Select IPLBCUST1,NVL(IPLBCUST2,'?') from (" & vbCrLf _
        Main_Process()
    End Sub

    Public Sub Main_Process()
        Fill_Records("CSMST")
        Update_Record()
    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Using swCSMST As New System.IO.StreamWriter(CSMST)

            For Each rowCSMST As DataRow In dst.Tables("CSMST").Select("", "IPLBCUST1,IPLBCUST2")

                'If rowCSMST.Item("CLRCUST1") & "" = "" _
                'Or (rowCSMST.Item("IPLBCUST2") & "" <> "" And rowCSMST.Item("CLRCUST2") & "" = "") Then
                '    ' New Customer or New Store
                '    If rowCSMST.Item("CLRCUST1") & "" = "" Then
                '        ' New Customer
                '    End If
                '    If (rowCSMST.Item("IPLBCUST2") & "" <> "" And rowCSMST.Item("CLRCUST2") & "" = "") Then
                '        ' New Store
                '    End If

                'Else

                R += 1

                Dim RECORD As String = ""

                For Each dcol As DataColumn In rowCSMST.Table.Columns
                    Dim DATUM As String = rowCSMST.Item(dcol.ColumnName) & ""
                    If dcol.DataType.ToString = "System.String" Then
                        If DATUM = "" Then
                            DATUM = " "
                        End If
                        DATUM = quo & DATUM & quo
                    ElseIf dcol.DataType.ToString = "System.DateTime" Then
                        If DATUM <> "" Then
                            DATUM = Format(rowCSMST.Item(dcol.ColumnName), "yyyyMMdd")
                        Else
                            DATUM = " "
                        End If
                    Else
                        DATUM = CStr(Val(DATUM))
                    End If
                    RECORD &= sep & DATUM
                Next

                swCSMST.WriteLine(Mid(RECORD, 2))
                ' End If
            Next
        End Using

        ' New Customers
        Dim html_CUSTs As String = ""
        For Each rowCSMST As DataRow In dst.Tables("CSMST").Select _
            ("IPLBCUST2 IS NULL AND CLRCUST1 IS NULL", "IPLBCUST1,IPLBCUST2")
            html_CUSTs &= "<br>"
            html_CUSTs &= "<tr>"
            For Each COLUMN_NAME As String In New String() _
                {"IPLBCUST1", "CLRCUST1", "CUSTNAME", "ADDRESS1", "ADDRESS2", "CITY", "STATE", "ZIPCODE", "ZIPCODE2", "COUNTRY", "AR_PHONE", "AR_CONTACT", "CARCODE"}
                html_CUSTs &= "<td" _
                    & IIf(COLUMN_NAME = "CLRCUST1", " style='background-color:yellow;'", "") _
                    & IIf(COLUMN_NAME = "CUSTNAME", "><a href='http://portal.interparfums.com/#/customer.html?linkToken=unjueu8u37h3&custCode=BELK&custStoreNo=000001'", "") _
                    & ">" _
                    & rowCSMST.Item(COLUMN_NAME) _
                    & IIf(COLUMN_NAME = "CUSTNAME", "</a>", "") _
                    & "</td>"
            Next
            html_CUSTs &= "</tr>"
        Next

        ' New Stores
        Dim html_CUST_STOREs As String = ""
        For Each rowCSMST As DataRow In dst.Tables("CSMST").Select _
            ("IPLBCUST2 IS NOT NULL AND CLRCUST2 IS NULL", "IPLBCUST1,IPLBCUST2")
            html_CUST_STOREs &= "<br>"
            html_CUST_STOREs &= "<tr>"
            For Each COLUMN_NAME As String In New String() _
                {"IPLBCUST1", "IPLBCUST2", "CLRCUST1", "CLRCUST2", "CUSTNAME", "ADDRESS1", "ADDRESS2", "CITY", "STATE", "ZIPCODE", "ZIPCODE2", "COUNTRY", "AR_PHONE", "AR_CONTACT", "EDISTORE#"}
                html_CUST_STOREs &= "<td" _
                    & IIf(COLUMN_NAME = "CLRCUST2", " style='background-color:yellow;'", "") _
                    & IIf(COLUMN_NAME = "CUSTNAME", "><a href='http://portal.interparfums.com/#/customer.html?linkToken=unjueu8u37h3&custCode=" & rowCSMST.Item("IPLBCUST1") & "&custStoreNo=" & rowCSMST.Item("IPLBCUST2") & "'", "") _
                    & ">" _
                    & rowCSMST.Item(COLUMN_NAME) _
                    & IIf(COLUMN_NAME = "CUSTNAME", "</a>", "") _
                    & "</td>"
            Next
            html_CUST_STOREs &= "</tr>"
        Next

        If html_CUSTs <> "" Or html_CUST_STOREs <> "" Then
            If html_CUSTs <> "" Then
                Dim html_CUSTs_hdr As String = "<br>"
                html_CUSTs_hdr &= "<table border='1' style='font-size: 8pt;border:solid 1px gray;width:100%;background-color: whitesmoke; border-collapse: collapse;'>"
                html_CUSTs_hdr &= "<caption style='font-size: 12pt;color:blue; text-align: left;'>Please assign CSCUS1 values to the following new customers:</caption>"
                html_CUSTs_hdr &= "<tr style='background-color:#e6e8fa;'>"
                For Each COLUMN_CAPTION As String In New String() _
                    {"Code", "CLRCUST1", "Name", "Address", "Address2", "City", "State", "Zip Code", "Zip+4", "Country", "Phone", "Contact", "Via"}
                    html_CUSTs_hdr &= "<td>" & COLUMN_CAPTION & "</td>"
                Next
                html_CUSTs_hdr &= "</tr>"
                html_CUSTs = html_CUSTs_hdr & "<br>" & html_CUSTs & "</table>"
            End If
            If html_CUST_STOREs <> "" Then
                Dim html_CUST_STOREs_hdr As String = "<br>"
                html_CUST_STOREs_hdr &= "<table border='1' style='font-size: 8pt;border:solid 1px gray;width:100%;background-color: whitesmoke; border-collapse: collapse;'>"
                html_CUST_STOREs_hdr &= "<caption style='font-size: 12pt;color:blue; text-align: left;'>Please assign CSCUS2 values to the following new stores:</caption>"
                html_CUST_STOREs_hdr &= "<tr style='background-color:#e6e8fa;'>"
                For Each COLUMN_CAPTION As String In New String() _
                    {"Code", "Store", "CLRCUST1", "CLRCUST2", "Name", "Address", "Address2", "City", "State", "Zip Code", "Zip+4", "Country", "Phone", "Contact", "EDI Store"}
                    html_CUST_STOREs_hdr &= "<td>" & COLUMN_CAPTION & "</td>"
                Next
                html_CUST_STOREs_hdr &= "</tr>"
                html_CUST_STOREs = html_CUST_STOREs_hdr & "<br>" & html_CUST_STOREs & "</table>"
            End If

            Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)
            EMAIL_ADDRESSs.Add("wjz@absolution.com", "Walter J. Zielenski, III")
            Dim ATTACHMENTs As New Dictionary(Of String, String)

            Dim CFG_XNO As String = ASCMAIN1.Next_Control_No("WHC999O1.CFG_XNO")
            'Send_email(EMAIL_ADDRESSs, _
            '           ATTACHMENTs, _
            '           "IPLB New Customers & Stores #" & CFG_XNO, _
            '           Me.MENU_ITEM_OBJECT, _
            '           html_CUSTs & IIf(html_CUST_STOREs = "", "", "<br><br>" & html_CUST_STOREs))

        End If
    End Sub

    Function Send_email(
            ByVal EMAIL_ADDRESSs As Dictionary(Of String, String),
            ByVal ATTACHMENTs As Dictionary(Of String, String),
            ByVal SUBJECT As String,
            ByVal EMAIL_KEY As String,
            Optional ByVal EMAIL_BODY_PART As String = "") As String

        Dim USER_ID_emailer As String = "wjz"
        Dim rowASTUSER1_EMAIL_FROM As DataRow = LookUp("ASTUSER1", USER_ID_emailer, True)
        Dim rowASTUSER1_EMAIL_BCC As DataRow = Nothing

        Dim USER_TELEPHONE As String = rowASTUSER1_EMAIL_FROM.Item("USER_TELEPHONE") & ""
        Dim USER_EXT As String = rowASTUSER1_EMAIL_FROM.Item("USER_EXT") & ""
        Dim USER_FAX As String = rowASTUSER1_EMAIL_FROM.Item("USER_FAX") & ""
        Dim EMAIL_BODY As String = "Attached is the file that you have requested."

        Dim rowTATMAIL1 As DataRow = LookUp("TATMAIL1", EMAIL_KEY)
        If rowTATMAIL1 IsNot Nothing Then
            If rowTATMAIL1.Item("EMAIL_FROM") & "" <> "" Then
                rowASTUSER1_EMAIL_FROM = LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_FROM"), True)
            End If
            If rowTATMAIL1.Item("EMAIL_BCC") & "" <> "" Then
                rowASTUSER1_EMAIL_BCC = LookUp("ASTUSER1", rowTATMAIL1.Item("EMAIL_BCC"), True)
            End If
            If rowTATMAIL1.Item("EMAIL_BODY") & "" <> "" Then
                EMAIL_BODY = rowTATMAIL1.Item("EMAIL_BODY")
            End If
        End If

        If EMAIL_BODY = "" Then
            EMAIL_BODY = "Dear CS" & "<br>" & EMAIL_BODY_PART
        Else
            EMAIL_BODY = Replace(EMAIL_BODY, "*", EMAIL_BODY_PART)
        End If
        Dim USER_SIGNATURE As String =
          rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & vbCrLf _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_TITLE") & vbCrLf, "") _
        & IIf(rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & "" <> "", rowASTUSER1_EMAIL_FROM.Item("USER_COMPANY") & vbCrLf, "") _
        & IIf(USER_TELEPHONE <> "", "Tel: " & ASCMAIN1.FormatTel(USER_TELEPHONE, USER_EXT) & vbCrLf, "") _
        & IIf(USER_FAX <> "", "Fax: " & ASCMAIN1.FormatTel(USER_FAX) & vbCrLf, "") _
        & rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & vbCrLf

        Dim frmTAFSEND1 As New TACSEND1(Me)

        frmTAFSEND1.EMAIL_KEY = EMAIL_KEY
        frmTAFSEND1.SEND_FROM = rowASTUSER1_EMAIL_FROM.Item("USER_EMAIL") & ""
        frmTAFSEND1.SEND_FROM_NAME = rowASTUSER1_EMAIL_FROM.Item("USER_NAME") & ""
        frmTAFSEND1.SEND_FROM_SIGNATURE = USER_SIGNATURE
        frmTAFSEND1.SEND_TOs = EMAIL_ADDRESSs
        frmTAFSEND1.SEND_TO = ""
        frmTAFSEND1.SEND_TO_NAME = ""
        frmTAFSEND1.SEND_CC = ""

        frmTAFSEND1.SEND_SUBJECT = SUBJECT

        frmTAFSEND1.SEND_BODY = EMAIL_BODY
        frmTAFSEND1.SEND_METHOD = "E"
        frmTAFSEND1.SEND_ATTACHMENTs = ATTACHMENTs
        frmTAFSEND1.SEND_ATTACHMENT = ""

        frmTAFSEND1.Send_email_automatically()

        Dim SEND_STATUS As String = frmTAFSEND1.SEND_STATUS
        Dim SEND_NO As String = frmTAFSEND1.SEND_NO

        frmTAFSEND1 = Nothing

        Return SEND_NO ' SEND_STATUS

    End Function


    Overrides Sub Update_Archive()

        ' Added 09/26/2025 to prevent sending test environment data
        If ASCMAIN1.DBS_COMPANY <> ASCMAIN1.CLIENT OrElse ASCMAIN1.DBS_SERVER <> ASCMAIN1.CLIENT Then
            If ASCMAIN1.Running_in_VS Then
                Stop
            Else
                MessageBox.Show("You are not in production. File transfer avoided.", "Update Archive", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
        End If

        MyBase.Update_Archive()

        Dim subDir As String = DateTime.Now.ToString("yyyyMM")
        ' If the Archive direcory does not exist then create it as a courtesy
        If Not My.Computer.FileSystem.DirectoryExists(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir) Then
            My.Computer.FileSystem.CreateDirectory(ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir)
        End If

        My.Computer.FileSystem.CopyFile(CSMST, sftp_folder & CSMST)
        My.Computer.FileSystem.MoveFile(CSMST, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & CSMST)

        ASCDATA1.ExecuteSQL("Delete from CONV.CFG_CSMST")
        ASCDATA1.ExecuteSQL("Insert into CONV.CFG_CSMST Select * from " & TMP_CSMST)
    End Sub

    Sub Create_Work_Table()

        ASCMAIN1.sql = "Select * from CONV.CFG_CSMST where ROWNUM < 1"
        TMP_CSMST = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
        ' ASCDATA1.ExecuteSQL("Alter Table " & TMP_CSMST & " Add Primary Key (CSCUS1,CSCUS2)")

        'ASCMAIN1.sql = "" _
        '    & "Select" & vbCrLf _
        '    & "  ARTCUST1.CUST_NAME CUSTNAME" & vbCrLf _
        '    & ", SUBSTR(ARTCUST1.CUST_ADDR1,1,30) ADDRESS1" & vbCrLf _
        '    & ", ARTCUST1.CUST_ADDR2 ADDRESS2" & vbCrLf _
        '    & ", ARTCUST1.CUST_CITY CITY" & vbCrLf _
        '    & ", ARTCUST1.CUST_STATE STATE" & vbCrLf _
        '    & ", SUBSTR(ARTCUST1.CUST_ZIP_CODE,1,5) ZIPCODE" & vbCrLf _
        '    & ", SUBSTR(SUBSTR(ARTCUST1.CUST_ZIP_CODE,6),-4) ZIPCODE2" & vbCrLf _
        '    & ", 'ABS CUSTOMER' SHIPNAME" & vbCrLf _
        '    & ", NULL SHIPADDR1" & vbCrLf _
        '    & ", NULL SHIPADDR2" & vbCrLf _
        '    & ", NULL SHIPCITY" & vbCrLf _
        '    & ", NULL SHIPSTATE" & vbCrLf _
        '    & ", NULL SHIPZIP" & vbCrLf _
        '    & ", NULL SHIPZIP2" & vbCrLf _
        '    & ", NULL PHONE" & vbCrLf _
        '    & ", NULL CONTACT" & vbCrLf _
        '    & ", ARTCUST1_BT.CUST_PHONE AR_PHONE" & vbCrLf _
        '    & ", ARTCUST1_BT.CUST_CONTACT AR_CONTACT" & vbCrLf _
        '    & ", DECODE(ARTCUST1.SHIP_VIA_CODE,'RGOF','RGF',ARTCUST1.SHIP_VIA_CODE) CARCODE" & vbCrLf _
        '    & ", NULL EDISTORE#" & vbCrLf _
        '    & ", ARTCUST1.CUST_CODE IPLBCUST1" & vbCrLf _
        '    & ", NULL IPLBCUST2" & vbCrLf _
        '    & ", DECODE(ARTCUST1.CUST_INCL_INV_SHIP,'1','Y','N') INVWSHIP" & vbCrLf _
        '    & ", ARTCUST1.FDX_ACCT_NO FDXACCOUNT" & vbCrLf _
        '    & ", ARTCUST1.UPS_ACCT_NO UPSACCOUNT" & vbCrLf _
        '    & ", SOTSREP1.SREP_EMAIL IPLBCONT" & vbCrLf _
        '    & ", NULL BATCHID" & vbCrLf _
        '    & ", NULL COMBINE" & vbCrLf _
        '    & ", ARTCUST1.CUST_NO_3PL CLRCUST1" & vbCrLf _
        '    & ", NULL CLRCUST2" & vbCrLf _
        '    & ", NVL(TATCNTRY.COUNTRY_CODE2,'US') COUNTRY" & vbCrLf _
        '    & ", NULL STATUS" & vbCrLf _
        '    & ", NULL CLRCUSTS1" & vbCrLf _
        '    & ", NULL CLRCUSTS2" & vbCrLf _
        '    & ", NULL IPLBCUSTS1" & vbCrLf _
        '    & ", NULL IPLBCUSTS2" & vbCrLf _
        '    & ", NULL DFCYN" & vbCrLf _
        '    & " from ARTCUST1, SOTSREP1, ARTCUST1 ARTCUST1_BT, TATCNTRY" & vbCrLf _
        '    & " where SOTSREP1.SREP_CODE (+) = ARTCUST1.SREP_CODE" & vbCrLf _
        '    & "   and TATCNTRY.COUNTRY_CODE (+) = ARTCUST1.CUST_COUNTRY" & vbCrLf _
        '    & "   and ARTCUST1_BT.CUST_CODE (+) = NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE)" & vbCrLf _
        '    & "   and ARTCUST1.CUST_CODE NOT IN ('IPLBFIN')" & vbCrLf _
        ' & "   and ARTCUST1.CUST_CODE NOT IN (Select CUST_CODE from SOTCHAN1 where CUST_CODE is Not Null)" & vbCrLf _
        '    & "   and NVL(ARTCUST1.CUST_SHIP_TO_MANUAL,'0') <> '1'" & vbCrLf _
        '    & " UNION " & vbCrLf _
        ASCMAIN1.sql = "" _
         & "Select" & vbCrLf _
         & "  ARTCUST2.CUST_STORE_NAME CUSTNAME" & vbCrLf _
         & ", SUBSTR(ARTCUST2.CUST_STORE_ADDR1,1,30) ADDRESS1" & vbCrLf _
         & ", ARTCUST2.CUST_STORE_ADDR2 ADDRESS2" & vbCrLf _
         & ", ARTCUST2.CUST_STORE_CITY CITY" & vbCrLf _
         & ", ARTCUST2.CUST_STORE_STATE STATE" & vbCrLf _
         & ", DECODE(ARTCUST2.CUST_STORE_COUNTRY, 'CAN', SUBSTR(REPLACE(ARTCUST2.CUST_STORE_ZIP_CODE, ' ',''),1,3), SUBSTR(ARTCUST2.CUST_STORE_ZIP_CODE,1,5)) ZIPCODE" & vbCrLf _
         & ", DECODE(ARTCUST2.CUST_STORE_COUNTRY, 'CAN', SUBSTR(REPLACE(ARTCUST2.CUST_STORE_ZIP_CODE, ' ',''),4),   SUBSTR(SUBSTR(ARTCUST2.CUST_STORE_ZIP_CODE,6),-4)) ZIPCODE2" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_NAME SHIPNAME" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_ADDR1 SHIPADDR1" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_ADDR2 SHIPADDR2" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_CITY SHIPCITY" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_STATE SHIPSTATE" & vbCrLf _
         & ", DECODE(ARTCUST2_DC.CUST_STORE_COUNTRY, 'CAN', SUBSTR(REPLACE(ARTCUST2_DC.CUST_STORE_ZIP_CODE, ' ',''),1,3), SUBSTR(ARTCUST2_DC.CUST_STORE_ZIP_CODE,1,5)) SHIPZIP" & vbCrLf _
         & ", DECODE(ARTCUST2_DC.CUST_STORE_COUNTRY, 'CAN', SUBSTR(REPLACE(ARTCUST2_DC.CUST_STORE_ZIP_CODE, ' ',''),4),   SUBSTR(SUBSTR(ARTCUST2_DC.CUST_STORE_ZIP_CODE,6),-4)) SHIPZIP2" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_PHONE PHONE" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_CONTACT CONTACT" & vbCrLf _
         & ", ARTCUST1_BT.CUST_PHONE AR_PHONE" & vbCrLf _
         & ", ARTCUST1_BT.CUST_CONTACT AR_CONTACT" & vbCrLf _
         & ", NVL(ARTCUST2.CUST_STORE_SHIP_VIA_CODE, DECODE(ARTCUST1.SHIP_VIA_CODE,'RGOF','RGF',ARTCUST1.SHIP_VIA_CODE)) CARCODE" & vbCrLf _
         & ", CASE WHEN ARTCUST1.CUST_CODE = 'VETERANS' THEN SUBSTR(ARTCUST2.CUST_STORE_NO,2) ELSE CASE WHEN EDTSLSP1.NUMBER_CHARS_STORE IS NULL THEN NULL ELSE TRIM(SUBSTR('      ' || ARTCUST2.CUST_STORE_NO, -1 * NUMBER_CHARS_STORE)) END END EDISTORE#" & vbCrLf _
         & ", ARTCUST2.CUST_CODE IPLBCUST1" & vbCrLf _
         & ", ARTCUST2.CUST_STORE_NO IPLBCUST2" & vbCrLf _
         & ", DECODE(ARTCUST1.CUST_INCL_INV_SHIP,'1','Y','N') INVWSHIP" & vbCrLf _
         & ", ARTCUST1.FDX_ACCT_NO FDXACCOUNT" & vbCrLf _
         & ", ARTCUST1.UPS_ACCT_NO UPSACCOUNT" & vbCrLf _
         & ", SOTSREP1.SREP_EMAIL IPLBCONT" & vbCrLf _
         & ", NULL BATCHID" & vbCrLf _
         & ", NULL COMBINE" & vbCrLf _
         & ", ARTCUST2.CUST_NO_3PL CLRCUST1" & vbCrLf _
         & ", ARTCUST2.CUST_STORE_NO_3PL CLRCUST2" & vbCrLf _
         & ", NVL(TATCNTRY.COUNTRY_CODE2,'US') COUNTRY" & vbCrLf _
         & ", NULL STATUS" & vbCrLf _
         & ", ARTCUST2_DC.CUST_NO_3PL CLRCUSTS1" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_NO_3PL CLRCUSTS2" & vbCrLf _
         & ", ARTCUST2_DC.CUST_CODE IPLBCUSTS1" & vbCrLf _
         & ", ARTCUST2_DC.CUST_STORE_NO IPLBCUSTS2" & vbCrLf _
         & ", NULL DFCYN" & vbCrLf _
         & " from ARTCUST1, ARTCUST2, SOTSREP1, ARTCUST1 ARTCUST1_BT, ARTCUST2 ARTCUST2_DC, EDTSLSP1, TATCNTRY" & vbCrLf _
         & " where ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE" & vbCrLf _
         & "   and ARTCUST1_BT.CUST_CODE (+) = NVL(ARTCUST1.CUST_BILL_TO_CUST,ARTCUST1.CUST_CODE)" & vbCrLf _
         & "   and TATCNTRY.COUNTRY_CODE (+) = ARTCUST2.CUST_STORE_COUNTRY" & vbCrLf _
         & "   and ARTCUST2_DC.CUST_CODE (+) = ARTCUST2.CUST_CODE" & vbCrLf _
         & "   and ARTCUST2_DC.CUST_STORE_NO (+) = NVL(ARTCUST2.CUST_DC_NO,ARTCUST2.CUST_STORE_NO)" & vbCrLf _
         & "   and SOTSREP1.SREP_CODE (+) = ARTCUST1.SREP_CODE" & vbCrLf _
         & "   and EDTSLSP1.CUST_CODE (+) = ARTCUST1.CUST_CODE" & vbCrLf _
         & "   and ARTCUST2.CUST_STORE_NAME IS NOT NULL and ARTCUST2.CUST_STORE_ADDR1 IS NOT NULL and ARTCUST2.CUST_STORE_CITY IS NOT NULL and ARTCUST2.CUST_STORE_STATE IS NOT NULL and ARTCUST2.CUST_STORE_ZIP_CODE IS NOT NULL" & vbCrLf _
         & "   and ARTCUST1.CUST_CODE NOT IN ('IPLBFIN')" & vbCrLf _
         & "   and ARTCUST1.CUST_CODE NOT IN (Select CUST_CODE from SOTCHAN1 where CUST_CODE is Not Null)" & vbCrLf _
         & "   and NVL(ARTCUST1.CUST_SHIP_TO_MANUAL,'0') <> '1'"

        ASCDATA1.ExecuteSQL("Insert into " & TMP_CSMST & " " & ASCMAIN1.sql)
    End Sub
End Class