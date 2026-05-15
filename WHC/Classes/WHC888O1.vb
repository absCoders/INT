Public Class WHC888O1

    ' Create Item Master Changes and Additions File (Outbound) 888

    Inherits WHC000O1

    ' THESE ARE THE FILENAMES USED FOR THE CLA INTERFACE
    Private TMP_ITMST As String
    Private TMP_ITMST3 As String
    Private ITMST As String = "ITMST" & "_" & DTS & ".CSV"
    Private tblWHTTPLP1 As DataTable

    Sub New()
        MyBase.New()
    End Sub

    Sub New(g As ABSEnvironment)
        MyBase.New(g)

        Me.MENU_ITEM_OBJECT = "WHC888O1"

        tblWHTTPLP1 = ASCDATA1.GetDataTable("SELECT * FROM WHTTPLP1", "WHTTPLP1")
        Addtask("WHC888O1 - Start Create_Work_Table")
        Create_Work_Table()
        Addtask("WHC888O1 - End Create_Work_Table")

        With dst
            If LP_CODE = "CLA" Then
                ASCMAIN1.sql = "Select * from " & TMP_ITMST & " minus Select * from CONV.CFG_ITMST"
            Else
                ASCMAIN1.sql = "Select * from " & TMP_ITMST
            End If
            Create_TDA(.Tables.Add, "ITMST", "**", 0, False)
        End With

        Main_Process()
    End Sub

    Sub Create_Work_Table()

        Select Case LP_CODE

            Case "CLA"

                ASCMAIN1.sql = "Select * from CONV.CFG_ITMST where ROWNUM < 1"
                TMP_ITMST = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
                ASCDATA1.ExecuteSQL("Alter Table " & TMP_ITMST & " Add Primary Key (ITITEM)")

                ASCMAIN1.sql = "Select" & vbCrLf _
                    & "  ICTCOLL1.BRAND_CODE_3PL ITCONO" & vbCrLf _
                    & ", ICTITEM1.ITEM_ALT_SORT ITITEM" & vbCrLf _
                    & ", NULL ITGLPG" & vbCrLf _
                    & ", NULL ITGLPC" & vbCrLf _
                    & ", ICTITEM1.ITEM_CLASS_CODE ITPDCL" & vbCrLf _
                    & ", REPLACE(ICTITEM1.ITEM_DESC,'" & Chr(34) & "','i') ITDESC" & vbCrLf _
                    & ", NULL ITDES2" & vbCrLf _
                    & ", NULL ITISRC" & vbCrLf _
                    & ", NULL ITPDCC" & vbCrLf _
                    & ", DECODE(NVL(ICTITEM1.ITEM_LOT_CONTROL,'0'),'1','M',NULL) ITTRCK" & vbCrLf _
                    & ", 'EACH' ITSKUM" & vbCrLf _
                    & ", ICTITEM1.ITEM_SO_QTY_MULT ITSKMQ" & vbCrLf _
                    & ", ICTITEM1.ITEM_SO_QTY_MIN ITA1MQ" & vbCrLf _
                    & ", ICTITEM1.ITEM_RETAIL_PRICE ITRETL" & vbCrLf _
                    & ", 0 ITPUUW" & vbCrLf _
                    & ", 0 ITPUUV" & vbCrLf _
                    & ", DECODE(ICTITEM1.ITEM_STATUS,'I','I',NULL) ITSTTS" & vbCrLf _
                    & ", 0 ITMQTY" & vbCrLf _
                    & ", DECODE(NVL(ICTITEM1.ITEM_CRITICAL_TO_SHIP,'0'),'1','Y',NULL) ITCRIT" & vbCrLf _
                    & ", ICTITEM1.ITEM_STD_PACK_SLS ITFCTR" & vbCrLf _
                    & ", DECODE(NVL(ICTITEM1.ITEM_WEIGHT_CHECK,'0'),'1','Y',NULL) ITWTCK" & vbCrLf _
                    & ", CASE WHEN NVL(ICTITEM1.ITEM_COST_STD,0)=0 THEN .01 ELSE NVL(ICTITEM1.ITEM_COST_STD,0) END IBCSTL" & vbCrLf _
                    & ", CASE WHEN NVL(ICTITEM1.ITEM_COST_STD,0)=0 THEN .01 ELSE NVL(ICTITEM1.ITEM_COST_STD,0) END IBCSTV" & vbCrLf _
                    & ", CASE WHEN NVL(ICTITEM1.ITEM_COST_STD,0)=0 THEN .01 ELSE NVL(ICTITEM1.ITEM_COST_STD,0) END IBCSTS" & vbCrLf _
                    & ", DECODE(ICTITEM1.ITEM_EAN_CODE,NULL,DECODE(ICTITEM1.ITEM_UPC_CODE,NULL,'BL','UP'),'EN') EAUORE" & vbCrLf _
                    & ", NVL(ICTITEM1.ITEM_EAN_CODE,ICTITEM1.ITEM_UPC_CODE) EAEAN#" & vbCrLf _
                    & ", NULL EAQRS" & vbCrLf _
                    & ", ICTITEM1.ITEM_CODE EAEFFD" & vbCrLf _
                    & ", ICTITEM1.ITEM_SHELF_LIFE_YRS PPLIFE" & vbCrLf _
                    & ", ICTCOLL1.SUB_BRAND_CODE_3PL ISSUBBRND" & vbCrLf _
                    & ", DECODE(NVL(ICTITEM1.ITEM_CONTAINS_ALCOHOL,'0'),'1','Y','N') OZMATFLG" & vbCrLf _
                    & ", DECODE(NVL(ICTITEM1.ITEM_APPR_1ST_REC,'0'),'1','Y','N') IAFLAG" & vbCrLf _
                    & ", DECODE(NVL(ICTITEM1.ITEM_STATUS,'X'),'A','N','Y') DSCFLG" & vbCrLf _
                    & ", DECODE(NVL(ICTITEM1.ITEM_ALLOW_HALF_PACK,'0'),'1','Y','N') HALFFLG" & vbCrLf _
                    & " from ICTITEM1,ICTCOLL1,ICTBRAN1" & vbCrLf _
                    & " where NVL(ICTITEM1.HIDE_FROM_3PL,'0') <> '1'" & vbCrLf _
                    & "   and NVL(ICTCOLL1.HIDE_FROM_3PL,'0') <> '1'" & vbCrLf _
                    & "   and NVL(ICTBRAN1.HIDE_FROM_3PL,'0') <> '1'" & vbCrLf _
                    & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
                    & "   and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE"
                ASCDATA1.ExecuteSQL("Insert into " & TMP_ITMST & " " & ASCMAIN1.sql)

            Case "ADS"

                '.Item("Category") = "ZFIN" ' row.Item("PROD_CODE") ' ICTITEM1.PROD_CODE
                '.Item("CountryOfOrigin") = "FR" ' row.Item("COUNTRY_CODE2") ' TATCNTRY.COUNTRY_CODE2 

                TMP_ITMST = String.Empty
                TMP_ITMST3 = String.Empty

                Dim sqlMinus As String = "MINUS
                                        Select LP_CODE, ITEM_CODE, ITEM_DESC, ITEM_EAN_CODE, ITEM_UPC_CODE,
                                        ITEM_WEIGHT, PROD_CODE, COUNTRY_CODE, COMMODITY_CODE,
                                        ITEM_SHELF_LIFE_YRS, ITEM_COST_STD, ITEM_RETAIL_PRICE
                                        FROM ICTIT3PL WHERE LP_CODE = 'ADS'"

                ASCMAIN1.sql = "Select 'ADS' LP_CODE, ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_EAN_CODE, ICTITEM1.ITEM_UPC_CODE,
                                        ICTITEM1.ITEM_WEIGHT, 'ZFIN' PROD_CODE, TATCNTRY.COUNTRY_CODE2 COUNTRY_CODE, ICTITEM1.COMMODITY_CODE,
                                        ICTITEM1.ITEM_SHELF_LIFE_YRS, ICTITEM1.ITEM_COST_STD, ICTITEM1.ITEM_RETAIL_PRICE
                                        from ICTITEM1, TATCNTRY, ICTCOLL1, ICTBRAN1
                                        where TATCNTRY.COUNTRY_CODE(+) = ICTITEM1.COUNTRY_CODE
                                        and ICTITEM1.ITEM_STATUS = 'A'
                                        and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE
                                        and ICTBRAN1.BRAND_CODE = ICTCOLL1.BRAND_CODE
                                        and NVL(ICTITEM1.HIDE_FROM_3PL, '0') <> '1'
                                        and NVL(ICTCOLL1.HIDE_FROM_3PL,'0') <> '1'
                                        and NVL(ICTBRAN1.HIDE_FROM_3PL,'0') <> '1'"

                Select Case G.APP_KEY
                    Case "All"

                    Case "SO"
                        ' At this stage, we'll go ahead and send all items, regardless.
                        'ASCMAIN1.sql &= "AND ICTITEM1.ITEM_CODE IN (SELECT DISTINCT ITEM_CODE FROM SOTORDR2 WHERE ORDR_STATUS = 'P')"

                    Case "PO"
                        ' At this stage, we'll go ahead and send all items, regardless.
                        'ASCMAIN1.sql &= "AND ICTITEM1.ITEM_CODE IN (SELECT DISTINCT ITEM_CODE FROM POTORDR2 WHERE PO_QTY_OPN > 0)"

                    Case Else
                        ' This sends a single Item
                        If G.APP_KEY.StartsWith("IT:") Then
                            Dim ITEM_CODE As String = G.APP_KEY.Substring(3)
                            sqlMinus = $" AND ICTITEM1.ITEM_CODE = '{ITEM_CODE}'"
                        End If

                End Select

                ASCMAIN1.sql &= sqlMinus

                TMP_ITMST = ASCMAIN1.Temp_Table(ASCMAIN1.sql)
                ASCDATA1.ExecuteSQL("Alter Table " & TMP_ITMST & " Add Primary Key (ITEM_CODE)")
        End Select

    End Sub

    Public Sub Main_Process()

        Fill_Records("ITMST")

        Select Case LP_CODE
            Case "ADS"

            Case "CLA"

        End Select

        Update_Record()

    End Sub

    Public Overrides Sub Update_Create_File()
        MyBase.Update_Create_File()

        Select Case LP_CODE

            Case "CLA"

                Using swITMST As New System.IO.StreamWriter(ITMST)

                    For Each rowITMST As DataRow In dst.Tables("ITMST").Select("", "ITITEM")
                        R += 1

                        Dim RECORD As String = ""

                        For Each dcol As DataColumn In rowITMST.Table.Columns
                            Dim DATUM As String = rowITMST.Item(dcol.ColumnName) & ""
                            If dcol.DataType.ToString = "System.String" Then
                                If DATUM = "" Then
                                    DATUM = " "
                                End If
                                DATUM = Replace(DATUM, quo, quo & quo & quo)
                                DATUM = quo & DATUM & quo
                            ElseIf dcol.DataType.ToString = "System.DateTime" Then
                                If DATUM <> "" Then
                                    DATUM = Format(rowITMST.Item(dcol.ColumnName), "yyyyMMdd")
                                Else
                                    DATUM = " "
                                End If
                            Else
                                DATUM = CStr(Val(DATUM))
                            End If
                            RECORD &= sep & DATUM
                        Next

                        swITMST.WriteLine(Mid(RECORD, 2))
                    Next
                End Using

            Case "ADS"

                Dim Folder As String = ASCMAIN1.Folders("SharedRoot") & "XSDs\"

                Dim XSD_name As String = "WWIMPZITM"
                Dim dstITM As New DataSet

                Addtask($"WHC888O1 - Start ReadXmlSchema {Folder & XSD_name & ".XSD"}")
                dstITM.ReadXmlSchema(Folder & XSD_name & ".XSD")
                Addtask($"WHC888O1 - End ReadXmlSchema {Folder & XSD_name & ".XSD"}")
                dstITM.EnforceConstraints = False

                Dim t As String = "Item"
                dstITM.Tables(t).Rows.Clear()

                For Each row As DataRow In dst.Tables("ITMST").Select("", "ITEM_CODE")
                    Dim rowITM As DataRow = dstITM.Tables(t).NewRow
                    With rowITM
                        .Item("StockNumber") = row.Item("ITEM_CODE")
                        .Item("Description1") = row.Item("ITEM_DESC")
                        '.Item("Description2") = row.Item("COLLECTION_CODE")
                        '.Item("Description3") = row.Item("ITEM_CLASS_CODE")
                        Dim ITEM_EAN_CODE As String = row.Item("ITEM_EAN_CODE") & ""
                        Dim ITEM_UPC_CODE As String = row.Item("ITEM_UPC_CODE") & ""
                        Dim EAN_UPC As String = ITEM_EAN_CODE
                        If EAN_UPC = "" Then
                            EAN_UPC = ITEM_UPC_CODE
                        End If
                        .Item("UpcEanCode") = EAN_UPC

                        .Item("ItemWeight") = row.Item("ITEM_WEIGHT")
                        .Item("Category") = "ZFIN" ' row.Item("PROD_CODE")
                        .Item("CountryOfOrigin") = row.Item("COUNTRY_CODE")

                        ' ISSUE-7029 HAZMAT -> Commodity code, retransmit
                        Dim COMM_CODE As String = row.Item("COMMODITY_CODE") & ""
                        .Item("HsCode") = COMM_CODE.Substring(0, Math.Min(6, Len(COMM_CODE)))
                        .Item("ShelfLifeDays") = Val(row.Item("ITEM_SHELF_LIFE_YRS") & "") * 365
                        .Item("PurchasePrice") = row.Item("ITEM_COST_STD")
                    End With
                    dstITM.Tables(t).Rows.Add(rowITM)
                    R += 1

                    Dim Item_ID As Int64 = Val(rowITM.Item("Item_ID") & "")
                    dstITM.Tables("Sales").Rows.Add(New Object() {Val(row.Item("ITEM_RETAIL_PRICE") & ""), Item_ID})
                Next

                If R > 0 Then
                    FILENAME_TO_SEND = $"{XSD_name}_{XMIT_NO}.XML"
                    Addtask($"WHC888O1 - Start WriteXml {FILENAME_TO_SEND}")
                    dstITM.WriteXml(FILENAME_TO_SEND)
                    Addtask($"WHC888O1 - End WriteXml {FILENAME_TO_SEND}")
                End If

        End Select

    End Sub

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

        Select Case LP_CODE

            Case "CLA"

                My.Computer.FileSystem.CopyFile(ITMST, sftp_folder & ITMST)
                My.Computer.FileSystem.MoveFile(ITMST, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & ITMST)

                ASCDATA1.ExecuteSQL("Delete from CONV.CFG_ITMST")
                ASCDATA1.ExecuteSQL("Insert into CONV.CFG_ITMST Select * from " & TMP_ITMST)

            Case "ADS"

                sftp_put("ADS", True, FILENAME_TO_SEND, FILENAME_TO_SEND)
                My.Computer.FileSystem.MoveFile(FILENAME_TO_SEND, ASCMAIN1.Folders("Archive") & "FROM_IPLB\" & subDir & "\" & FILENAME_TO_SEND)

                ASCDATA1.ExecuteSQL($"DELETE FROM ICTIT3PL WHERE (LP_CODE, ITEM_CODE) IN (SELECT LP_CODE, ITEM_CODE FROM {TMP_ITMST})")
                ASCDATA1.ExecuteSQL($"Insert into ICTIT3PL Select * from {TMP_ITMST}")

                Dim rowWHTTPLP1 As DataRow = tblWHTTPLP1.Rows.Find(LP_CODE)

                If rowWHTTPLP1 IsNot Nothing AndAlso rowWHTTPLP1.Item("TRANSMIT_ALL_ITEMS") & String.Empty = "1" Then
                    ASCDATA1.ExecuteSQL($"BEGIN DECLARE CURSOR C1 IS SELECT * FROM {TMP_ITMST};
                                        BEGIN FOR R1 IN C1 LOOP
                                            UPDATE ICTITLP2 SET LAST_XMIT = SYSDATE WHERE LP_CODE = R1.LP_CODE AND ITEM_CODE = R1.ITEM_CODE;
                                            IF SQL%ROWCOUNT = 0 THEN
                                                INSERT INTO ICTITLP2 (LP_CODE, ITEM_CODE, INIT_DATE, INIT_OPER, INIT_XMIT, ITEM_CODE_3PL)
                                                VALUES (R1.LP_CODE, R1.ITEM_CODE, SYSDATE, '{ASCMAIN1.USER_ID}', SYSDATE, R1.ITEM_CODE);
                                                END IF;
                                       END LOOP; END; END;")
                Else
                    ASCDATA1.ExecuteSQL($"BEGIN DECLARE CURSOR C1 IS SELECT * FROM {TMP_ITMST};
                                        BEGIN FOR R1 IN C1 LOOP
                                            UPDATE ICTITLP4 SET LAST_XMIT = SYSDATE WHERE LP_CODE = R1.LP_CODE AND ITEM_CODE = R1.ITEM_CODE;
                                            IF SQL%ROWCOUNT = 0 THEN
                                                INSERT INTO ICTITLP4 (LP_CODE, ITEM_CODE, INIT_DATE, INIT_OPER, INIT_XMIT, ITEM_CODE_3PL)
                                                VALUES (R1.LP_CODE, R1.ITEM_CODE, SYSDATE, '{ASCMAIN1.USER_ID}', SYSDATE, R1.ITEM_CODE);
                                                END IF;
                                        END LOOP; END; END;")
                End If
        End Select
    End Sub

End Class