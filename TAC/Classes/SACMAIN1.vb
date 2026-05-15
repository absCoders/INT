Public Class SACMAIN1

    Public Shared Sub Create_Sales_Extract_Tables(ByVal frmASFBASE0 As ASFBASE0, _
    ByVal initialize As Boolean, _
    ByRef SATSLSXC As String, ByRef SATSLSXI As String, ByRef SATSLSXS As String, _
    Optional ByVal scope As String = "", _
    Optional ByVal XNO As String = "", _
    Optional ByVal DTE1 As Date = Nothing, Optional ByVal DTE2 As Date = Nothing)

        If initialize Then
            SATSLSXC = Create_SATSLSXC(frmASFBASE0, initialize, SATSLSXC, scope, XNO, DTE1, DTE2)
            SATSLSXI = Create_SATSLSXI(frmASFBASE0, initialize, SATSLSXI, scope, XNO, DTE1, DTE2)
            SATSLSXS = Create_SATSLSXS(frmASFBASE0, initialize, SATSLSXS, scope, XNO, DTE1, DTE2)

            With frmASFBASE0.dst
                ASCMAIN1.sql = "Select * from " & SATSLSXC
                frmASFBASE0.Create_TDA(.Tables.Add, "SATSLSXC", "**", 0, False, "", 0)

                ASCMAIN1.sql = "Select * from " & SATSLSXI
                frmASFBASE0.Create_TDA(.Tables.Add, "SATSLSXI", "**", 0, False, "", 0)

                ASCMAIN1.sql = "Select * from " & SATSLSXS
                frmASFBASE0.Create_TDA(.Tables.Add, "SATSLSXS", "**", 0, False, "", 0)
            End With
        Else
            Create_SATSLSXC(frmASFBASE0, initialize, SATSLSXC, scope, XNO, DTE1, DTE2)
            Create_SATSLSXI(frmASFBASE0, initialize, SATSLSXI, scope, XNO, DTE1, DTE2)
            Create_SATSLSXS(frmASFBASE0, initialize, SATSLSXS, scope, XNO, DTE1, DTE2)
        End If
    End Sub

    Public Shared Function Create_SATSLSXC(frmASFBASE0 As ASFBASE0, initialize As Boolean, SATSLSXC As String, _
                                           Optional scope As String = "", _
                                           Optional XNO As String = "", _
                                           Optional DTE1 As Date = Nothing, Optional DTE2 As Date = Nothing) As String

        ASCMAIN1.sql = "Select 'zusa' COMPANY_ID" & vbCrLf _
            & ", ARTCUST1.CUST_CODE CUSTOMER_CODE, ARTCUST1.CUST_NAME CUSTOMER_NAME" & vbCrLf _
            & ", ARTCUST2.CUST_STORE_NO BRANCH_CODE, ARTCUST2.CUST_STORE_NAME BRANCH_NAME" & vbCrLf _
            & ", ARTCUST1.SREP_CODE GEO_REGION_CODE, SOTSREP1.SREP_NAME GEO_REGION_DESC" & vbCrLf _
            & ", ARTCUST1.CUST_CLASS_CODE SEGMENT_CODE, ARTCLAS1.CUST_CLASS_DESC SEGMENT_DESC" & vbCrLf _
            & ", ARTCUST1.TRADE_CLASS_CODE REGION_CODE, SOTTCLS1.TRADE_CLASS_DESC REGION_DESC" & vbCrLf _
            & " from ARTCUST1, ARTCUST2, SOTTCLS1, SOTMKTC1, SOTSREP1, ARTCLAS1" & vbCrLf _
            & " where ARTCUST2.CUST_CODE = ARTCUST1.CUST_CODE" & vbCrLf _
            & "   and SOTTCLS1.TRADE_CLASS_CODE = ARTCUST1.TRADE_CLASS_CODE" & vbCrLf _
            & "   and SOTMKTC1.MARKET_CODE (+) = SOTTCLS1.MARKET_CODE" & vbCrLf _
            & "   and ARTCLAS1.CUST_CLASS_CODE (+) = ARTCUST1.CUST_CLASS_CODE" & vbCrLf _
            & "   and SOTSREP1.SREP_CODE (+) = ARTCUST1.SREP_CODE"

        If initialize Then
            ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where ROWNUM < 1"
            SATSLSXC = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSXC & " Add Primary Key (COMPANY_ID, CUSTOMER_CODE, BRANCH_CODE)")
        Else
            'ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSXC)
            ASCDATA1.ExecuteSQL("Delete from " & SATSLSXC)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSXC & " " & ASCMAIN1.sql)
            frmASFBASE0.Fill_Records("SATSLSXC")
            Create_CSV(frmASFBASE0, "SATSLSXC", "customer", "CUSTOMER_CODE,BRANCH_CODE", scope, XNO, DTE1, DTE2)
        End If

        Return SATSLSXC
    End Function

    Public Shared Function Create_SATSLSXI(frmASFBASE0 As ASFBASE0, initialize As Boolean, SATSLSXI As String, _
                                           Optional scope As String = "", _
                                           Optional XNO As String = "", _
                                           Optional DTE1 As Date = Nothing, Optional DTE2 As Date = Nothing) As String

        'ASCMAIN1.sql = "Select 'ANA' || ICTITEM1.ITEM_CODE PART_CODE, ICTITEM1.ITEM_DESC PART_DESCRIPTION, X.OH BALANCE" & vbCrLf _
        '    & ", 'C' CATEGORY, 'F' FAMILY_CODE, 'D' FAMILY_DESCRIPTION, 'S' SKIN_TYPE_DESCRIPTION, 'P' PRODUCT_LINE" & vbCrLf _
        '    & " from ICTITEM1, (SELECT ITEM_CODE, SUM (WHSE_QTY_ON_HAND) OH from ICTSTAT2 group by ITEM_CODE) X" & vbCrLf _
        '    & " where X.ITEM_CODE (+) = ICTITEM1.ITEM_CODE"

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, X.OH BALANCE, ICTITEM1.ITEM_UPC_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE, ICTPROD1.PROD_DESC" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.COLLECTION_NAME" & vbCrLf _
            & ", ICTITEM1.ITEM_TYPE_CODE, ICTTYPE1.ITEM_TYPE_DESC" & vbCrLf _
            & ", ICTITEM1.COST_CATGY_CODE, ICTCOST1.COST_CATGY_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE, DECODE(ICTITEM1.ITEM_SNU_CODE,'S','Saleable','N','No-Charge','U','Unfinished','Unknown') ITEM_SNU_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_MAKE_BUY, DECODE(ICTITEM1.ITEM_COST_MAKE_BUY,'M','Make','B','Buy','Unknown') ITEM_COST_MAKE_BUY_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_FRT_CLASS, ICTFRTC1.FRT_CLASS_DESC ITEM_COST_FRT_CLASS_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_STD" & vbCrLf _
            & ", ICTITEM1.ITEM_PLAN_MAKE_BUY, DECODE(ICTITEM1.ITEM_PLAN_MAKE_BUY,'M','Make','B','Buy','Unknown') ITEM_PLAN_MAKE_BUY_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_BASIC_PROMO, DECODE(ICTITEM1.ITEM_BASIC_PROMO,'B','Basic','P','Promo','Unknown') ITEM_BASIC_PROMO_DESC" & vbCrLf _
            & ", ICTITEM1.SALES_DIVISION_CODE, SOTSDIV1.SALES_DIVISION_NAME" & vbCrLf _
            & " from ICTITEM1, (SELECT ITEM_CODE, SUM (WHSE_QTY_ON_HAND) OH from ICTSTAT2 group by ITEM_CODE) X" & vbCrLf _
            & ", ICTPROD1, ICTCOLL1, ICTTYPE1, ICTCOST1, ICTFRTC1, SOTSDIV1" & vbCrLf _
            & " where X.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
            & "   and ICTPROD1.PROD_CODE (+) = ICTITEM1.PROD_CODE" & vbCrLf _
            & "   and ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   and ICTTYPE1.ITEM_TYPE_CODE (+) = ICTITEM1.ITEM_TYPE_CODE" & vbCrLf _
            & "   and ICTCOST1.COST_CATGY_CODE (+) = ICTITEM1.COST_CATGY_CODE" & vbCrLf _
            & "   and ICTFRTC1.FRT_CLASS_CODE (+) = ICTITEM1.ITEM_COST_FRT_CLASS" & vbCrLf _
            & "   and SOTSDIV1.SALES_DIVISION_CODE (+) = ICTITEM1.SALES_DIVISION_CODE"

        If initialize Then
            ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where ROWNUM < 1"
            SATSLSXi = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSXi & " Add Primary Key (ITEM_CODE)")
        Else
            'ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSXI)
            ASCDATA1.ExecuteSQL("Delete from " & SATSLSXI)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSXi & " " & ASCMAIN1.sql)
            frmASFBASE0.Fill_Records("SATSLSXI")
            Create_CSV(frmASFBASE0, "SATSLSXI", "part", "ITEM_CODE", scope, XNO, DTE1, DTE2)
        End If

        Return SATSLSXi
    End Function

    Public Shared Function Create_SATSLSXS(frmASFBASE0 As ASFBASE0, initialize As Boolean, SATSLSXS As String, _
                                           Optional scope As String = "", _
                                           Optional XNO As String = "", _
                                           Optional DTE1 As Date = Nothing, Optional DTE2 As Date = Nothing) As String

        ASCMAIN1.sql = "Select 'zusa' COMPANY_ID, SOTINVH2.CUST_CODE CUSTOMER_CODE, SOTINVH2.CUST_STORE_NO BRANCH_CODE" & vbCrLf _
            & ", SOTINVH2.INV_NO IVNUM, SOTINVH2.INV_LNO LINE, TO_CHAR(SOTINVH1.INV_DATE,'YYYYMMDD') IVDATE" & vbCrLf _
            & ", 'USD' INVOICE_CURRENCY_CODE, 'ANA' || SOTINVH2.ITEM_CODE PART_CODE" & vbCrLf _
            & ", SOTINVH2.ORDR_QTY_SHIP SALES_QUANTITY, SOTINVH2.ORDR_UNIT_PRICE SALES_PRICE" & vbCrLf _
            & " from SOTINVH2, SOTINVH1" & vbCrLf _
            & " where SOTINVH1.INV_TYPE = SOTINVH2.INV_TYPE and SOTINVH1.INV_NO = SOTINVH2.INV_NO" & vbCrLf
        If initialize Then
            ASCMAIN1.sql &= " and ROWNUM < 1"
        Else
            If scope = "D" Then
                ASCMAIN1.sql &= "" _
                & "   and SOTINVH1.REGISTER_DATE >= '" & Format(DTE1, "dd-MMM-yyyy") & "'" & vbCrLf _
                & "   and SOTINVH1.REGISTER_DATE <= '" & Format(DTE2, "dd-MMM-yyyy") & "'"
            Else
                Dim XNOs As String = ""
                For Each row As DataRow In ASCDATA1.GetDataTable _
                    ("Select * from (Select Distinct REGISTER_XNO from SOTINVH1 where REGISTER_XNO <= '" & XNO & "' order by REGISTER_XNO DESC) where ROWNUM < 8").Select("", "REGISTER_XNO")
                    XNOs &= ",'" & row.Item("REGISTER_XNO") & "'"
                Next
                ASCMAIN1.sql &= "" _
                & "   and SOTINVH1.REGISTER_XNO in (" & Mid(XNOs, 2) & ")"
            End If
        End If

        If initialize Then
            ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") where ROWNUM < 1"
            SATSLSXS = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SATSLSXS & " Add Primary Key (IVNUM, LINE)")
        Else
            'ASCDATA1.ExecuteSQL("Truncate Table " & SATSLSXS)
            ASCDATA1.ExecuteSQL("Delete from " & SATSLSXS)
            ASCDATA1.ExecuteSQL("Insert into " & SATSLSXS & " " & ASCMAIN1.sql)
            frmASFBASE0.Fill_Records("SATSLSXS")
            Create_CSV(frmASFBASE0, "SATSLSXS", "sales", "IVNUM,LINE", scope, XNO, DTE1, DTE2)
        End If

        Return SATSLSXS
    End Function

    Public Shared Sub Create_CSV(frmASFBASE0 As ASFBASE0, TABLE_NAME As String, FILENAME As String, SORT As String, _
                                           scope As String, XNO As String, DTE1 As Date, DTE2 As Date)

        Using sw As New System.IO.StreamWriter(ASCMAIN1.Folders("Temp") & FILENAME & ".CSV")
            For Each row As DataRow In frmASFBASE0.dst.Tables(TABLE_NAME).Select("", SORT)
                Dim LINE As String = ""
                For Each col As DataColumn In frmASFBASE0.dst.Tables(TABLE_NAME).Columns
                    Dim dlim As String = ""
                    If col.DataType.Name = "String" Then
                        dlim = Chr(34)
                    End If
                    Dim VALUE As String = row.Item(col.ColumnName) & ""
                    If dlim <> "" Then
                        VALUE = Replace(VALUE, dlim, "")
                    End If
                    LINE &= "," & dlim & VALUE & dlim
                Next
                sw.WriteLine(Mid(LINE, 2))
            Next
            sw.Close()
        End Using

        ' Show_Document(ASCMAIN1.Folders("Temp") & FILENAME & ".CSV")
    End Sub

    Public Shared Sub ftp_BI_Files(frmASFBASE0 As ASFBASE0)

        Dim FILENAMEs() As String = New String() {"customer.csv", "part.csv", "sales.csv"}
        ' FILENAMEs = New String() {"sls_from_Jan01-2014_thru_Jun17-2015.csv"}
        Dim rowTATSSHK1 As DataRow = frmASFBASE0.LookUp("TATSSHK1", "BI")
        frmASFBASE0.Cursor = Cursors.WaitCursor

        Try
            ASCMAIN1.Progress("Now ftp'ing file to " & rowTATSSHK1.Item("SSH_APP_DESC"))

            If TAC.TACMAIN1.ftp_Files( _
                ASCMAIN1.Folders("Temp"), _
                FILENAMEs, _
                "", _
                FILENAMEs, _
                rowTATSSHK1.Item("SSH_APP_USERNAME"), _
                rowTATSSHK1.Item("SSH_APP_PASSWORD"), _
                rowTATSSHK1.Item("SSH_APP_PARTNER_URI_PROD")) Then

                For i As Integer = 0 To FILENAMEs.Length - 1
                    My.Computer.FileSystem.CopyFile( _
                        ASCMAIN1.Folders("Temp") & FILENAMEs(i), _
                        ASCMAIN1.Folders("Archive") & "BI\" & frmASFBASE0.Name & "_" & frmASFBASE0.XNO & "_" & FILENAMEs(i))
                Next

                MsgBox("ftp Completed Successfully", MsgBoxStyle.OkOnly, "Verification")
            Else
                MsgBox("ftp Failed", MsgBoxStyle.OkOnly, "Verification")
            End If

        Catch ex As Exception
            MsgBox(ex.Message, vbOKOnly, "Error Occurred")
        End Try

        frmASFBASE0.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

End Class