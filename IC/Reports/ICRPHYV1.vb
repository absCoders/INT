Public Class ICRPHYV1
    ' NEXT PHYSICAL = MOVE THE THE PREPARATION ROUTINES OUT TO TAC AND MOVE THE UPDATE INTO STATUS & CONTROL SCREEN
 
    Dim LYP As String = ""

    Dim ICTPHYV1 As String = ""
    Dim WHTPHYV1 As String = ""
    Dim WHSE_CODEs As String = ""
    Dim report_posted_variances As Boolean = False
    Dim ICTCOSTA As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Get_PARM("ICTPARM1")

        grdICTPHYVC.Visible = False

        LYP = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1)
    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables
        WHSE_CODEs = SQLA("WHSE_CODE")

        If chkUpdateVariances.Checked Then
            RWU = "R"
        Else
            RWU = "N"
        End If
        'RWU = "R" - TOO DANGEROUS - DO THIS IN A SCREEN

        Prepare_Work_File()

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""
        'sql_filter &= SQL_in("SALES_DIVISION_CODE", "ICTSTYL1.SALES_DIVISION_CODE")

        If numVARU.Value <> 0 Then
            RWU = "N"
            ASCDATA1.ExecuteSQL("Delete from " & ICTPHYV1 & " where ABS(NVL(BOOK,0) - NVL(PHYS,0)) < " & CStr(numVARU.Value))
        End If
        If numVARC.Value <> 0 Then
            RWU = "N"
            ASCDATA1.ExecuteSQL("Delete from " & ICTPHYV1 & " where ABS(ITEM_COST_TOTAL,0) * (NVL(BOOK,0) - NVL(PHYS,0))) < " & CStr(numVARC.Value))
        End If

        If optSORT.Value = "I" Then
            ASCDATA1.ExecuteSQL("Update " & ICTPHYV1 & " Set SORT_VALUE = ITEM_CODE")
        ElseIf optSORT.Value = "U" Then
            ASCDATA1.ExecuteSQL("Update " & ICTPHYV1 & " Set SORT_VALUE = TRIM(TO_CHAR(9999999999 - ABS(NVL(PHYS,0) - NVL(BOOK,0)),'0000000000'))")
        ElseIf optSORT.Value = "C" Then
            ASCDATA1.ExecuteSQL("Update " & ICTPHYV1 & " Set SORT_VALUE = TRIM(TO_CHAR(9999999999 - NVL(ITEM_COST_TOTAL,0) * ABS(NVL(PHYS,0) - NVL(BOOK,0)),'0000000000'))")
        End If

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & ICTPHYV1, "ICTPHYV1", 2))
        With dst.Tables("ICTPHYV1")
            .Columns.Add("VAR", GetType(System.Int64), "ISNULL(PHYS,0) - ISNULL(BOOK,0)")
            .Columns.Add("PHYS_AMT", GetType(System.Int64), "ISNULL(PHYS,0) * ISNULL(ITEM_COST_TOTAL,0)")
            .Columns.Add("BOOK_AMT", GetType(System.Int64), "ISNULL(BOOK,0) * ISNULL(ITEM_COST_TOTAL,0)")
            .Columns.Add("VAR_AMT", GetType(System.Int64), "ISNULL(VAR,0) * ISNULL(ITEM_COST_TOTAL,0)")
        End With

        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & WHTPHYV1, "WHTPHYV1", 3))

        dst.Tables.Add(ASCDATA1.GetDataTable("Select ICTITEM1.*, ICTCOSTA.ITEM_COST_TOTAL from ICTITEM1, " & ICTCOSTA & " ICTCOSTA where ICTCOSTA.ITEM_CODE (+) = ICTITEM1.ITEM_CODE AND ICTCOSTA.OPS_YYYYPP (+) = '" & LYP & "' AND ICTITEM1.ITEM_CODE in (Select Distinct ITEM_CODE from " & ICTPHYV1 & ")", "ICTITEM1", 1))

        ASCMAIN1.sql = "Select ICTPHYV1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTPHYV1.WHSE_CODE, ICTPHYV1.PHYS, ICTPHYV1.BOOK from " & ICTPHYV1 & " ICTPHYV1,ICTITEM1 where ICTITEM1.ITEM_CODE = ICTPHYV1.ITEM_CODE and ICTITEM1.ITEM_COST_STATUS = 'P'"
        Dim tblICTPHYVC As DataTable = ASCDATA1.GetDataTable
        If tblICTPHYVC.Rows.Count = 0 Then
            grdICTPHYVC.visible = False
        Else
            grdICTPHYVC.datasource = tblICTPHYVC
            grdICTPHYVC.Visible = True
            RWU = "N"
            MsgBox("There are Items with Physical Variances which do NOT have a Std Cost." _
                   & vbCrLf & "Please review these items in the grid and assign Std Costs." _
                   & vbCrLf & "The Update option will be disabled until this condition is rectified", _
                   MsgBoxStyle.OkOnly, "Cannot Proceed with Update")
        End If

        ' Prepare J/E

        Prepare_GL_Interface()



        ' Extracts from Data Sources

        MyBase.Get_SQL("*", ICTPHYV1)

        Dim SOURCE_TABLE_NAME As String = "ICTPHYV1"
        ' Dim x As String = ASTSRPT1_sum_columns
        ' Dim y As String = ASTSRPT1_sql_sum
        Dim sql_Data As String = ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
        & ", " & SOURCE_TABLE_NAME & ".SORT_VALUE, " & SOURCE_TABLE_NAME & ".ITEM_CODE" _
        & ASTSRPT1_sum_columns _
        & " from " & ICTPHYV1 & " " & SOURCE_TABLE_NAME & " " & sql_TABLE_NAMEs & vbCrLf _
        & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
        & " group by " & sql_GROUP_BY_cols _
        & ", " & SOURCE_TABLE_NAME & ".SORT_VALUE, " & SOURCE_TABLE_NAME & ".ITEM_CODE"

        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""

        If numVARU.Value <> 0 Then
            Page0.Add("Unit Variance Threshold: " & CStr(numVARU.Value))
        End If
        If numVARC.Value <> 0 Then
            Page0.Add("Cost Variance Threshold: " & CStr(numVARC.Value))
        End If
        Select Case optSORT.Value
            Case "I"
                Page0.Add("Sorted by Item")
            Case "U"
                Page0.Add("Ranked by Unit Variance")
            Case "C"
                Page0.Add("Ranked by Cost Variance")
        End Select
        CR_params.Add("OPTR", optSORT.Value)
        Generate_Report(RPT, , SUBT)

        Print_GL()
    End Sub

    Sub Prepare_Work_File()

        Dim SQLW As String = ""
        If WHSE_CODEs <> "" Then
            SQLW = " and X.WHSE_CODE in ('" & Replace(WHSE_CODEs, ",", "','") & "')"
        End If
        If ASCMAIN1.Running_in_VS And report_posted_variances Then ' SPECIAL RUN FOR RALPH
            SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_YYYYPP_LAST_PHY = '" & LYP & "')"
        Else
            SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_PHYS_STATUS = 'C')"
        End If

        ASCMAIN1.sql = "Select Distinct ICTPHYC2.ITEM_CODE" & vbCrLf _
            & " from ICTPHYC1,ICTPHYC2" & vbCrLf _
            & " where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" _
            & Replace(SQLW, "X.WHSE_CODE", "ICTPHYC2.WHSE_CODE")
        ASCMAIN1.sql = "" _
            & "Select ITEM_CODE from ICTCOSTA where OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and ITEM_CODE in (" & ASCMAIN1.sql & ")" _
            & " minus " _
            & "Select ITEM_CODE from ICTCOSTA where OPS_YYYYPP = '" & LYP & "' and ITEM_CODE in (" & ASCMAIN1.sql & ")"
        ASCMAIN1.sql = "Select * from ICTCOSTA where OPS_YYYYPP = '" & ASCMAIN1.CYP & "' and ITEM_CODE in (" & ASCMAIN1.sql & ")"
        ICTCOSTA = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & ICTCOSTA & " Add Primary Key (OPS_YYYYPP,ITEM_CODE)")
        ASCDATA1.ExecuteSQL("Update " & ICTCOSTA & " Set ITEM_EXP_IMP_IND = 'V', OPS_YYYYPP = '" & LYP & "'")
        ASCDATA1.ExecuteSQL("Insert into " & ICTCOSTA & " Select * from ICTCOSTA where OPS_YYYYPP >= '" & LYP & "' and OPS_YYYYPP <= '" & ASCMAIN1.CYP & "'")


        ASCMAIN1.sql = "Select X.*, ICTCOSTA.ITEM_COST_TOTAL" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM, ICTITEM1.COST_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " from ICTITEM1, " & ICTCOSTA & " ICTCOSTA, (" & vbCrLf _
            & "Select ITEM_CODE, WHSE_CODE, Sum (PHYS) PHYS, Sum (BOOK) BOOK from (" & vbCrLf _
            & "Select ICTPHYC2.ITEM_CODE, ICTPHYC2.WHSE_CODE" & vbCrLf _
            & ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS, 0 BOOK" _
            & " from ICTPHYC1,ICTPHYC2 where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" _
            & Replace(SQLW, "X.WHSE_CODE", "ICTPHYC2.WHSE_CODE") _
            & " group by ICTPHYC2.ITEM_CODE, ICTPHYC2.WHSE_CODE" & vbCrLf _
            & IIf(ASCMAIN1.Running_in_VS And report_posted_variances, "" _
            & " union " & vbCrLf _
            & "Select ICTSTAT1.ITEM_CODE, ICTSTAT1.WHSE_CODE" & vbCrLf _
            & ", NULL PHYS, Sum (-1 * NVL(ICTSTAT1.WHSE_QTY_PHY,0)) BOOK" _
            & " from ICTSTAT1 where NVL(ICTSTAT1.WHSE_QTY_PHY,0) <> 0" _
            & Replace(SQLW, "X.WHSE_CODE", "ICTSTAT1.WHSE_CODE") _
            & " and ICTSTAT1.OPS_YYYYPP = '" & LYP & "'" & vbCrLf _
            & " group by ICTSTAT1.ITEM_CODE, ICTSTAT1.WHSE_CODE" & vbCrLf _
            , "") _
            & " union " & vbCrLf _
            & "Select ICTSTAT1.ITEM_CODE, ICTSTAT1.WHSE_CODE" & vbCrLf _
            & ", 0 PHYS, Sum (NVL(ICTSTAT1.WHSE_QTY_BEG,0)) BOOK" _
            & " from ICTSTAT1 where NVL(ICTSTAT1.WHSE_QTY_BEG,0) <> 0" _
            & Replace(SQLW, "X.WHSE_CODE", "ICTSTAT1.WHSE_CODE") _
            & " and ICTSTAT1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & " group by ICTSTAT1.ITEM_CODE, ICTSTAT1.WHSE_CODE" & vbCrLf _
            & ") group by ITEM_CODE, WHSE_CODE) X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & " and ICTCOSTA.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
            & " and ICTCOSTA.OPS_YYYYPP (+) = '" & LYP & "'"
        ICTPHYV1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Alter Table " & ICTPHYV1 & " Add Primary Key (ITEM_CODE,WHSE_CODE)")
        ASCDATA1.ExecuteSQL("Alter Table " & ICTPHYV1 & " Add SORT_VALUE VARCHAR2(30)")

        SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_LOCATOR = '1')"
        ASCMAIN1.sql = "Select X.*, ICTCOSTA.ITEM_COST_TOTAL" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM, ICTITEM1.COST_CATGY_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.ITEM_CLASS_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & " from ICTITEM1, " & ICTCOSTA & " ICTCOSTA, (" & vbCrLf _
            & "Select ITEM_CODE, WHSE_CODE, LOCATION_CODE, Sum (PHYS) PHYS, Sum (BOOK) BOOK from (" & vbCrLf _
            & "Select ICTPHYC2.ITEM_CODE, ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
            & ", Sum (NVL(COUNT_CTNS,0) * NVL(CARTON_PACK_QTY,0) + NVL(COUNT_LOOSE,0)) PHYS, 0 BOOK" _
            & " from ICTPHYC1,ICTPHYC2 where ICTPHYC1.WHSE_CODE = ICTPHYC2.WHSE_CODE and ICTPHYC1.TICKET_NO = ICTPHYC2.TICKET_NO" _
            & Replace(SQLW, "X.WHSE_CODE", "ICTPHYC2.WHSE_CODE") _
            & " group by ICTPHYC2.ITEM_CODE, ICTPHYC2.WHSE_CODE, ICTPHYC1.LOCATION_CODE" & vbCrLf _
            & " union " & vbCrLf _
            & "Select WHTLOCB0.ITEM_CODE, WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
            & ", 0 PHYS, Sum (NVL(WHTLOCB0.LOCATION_QTY,0)) BOOK" _
            & " from WHTLOCB0,ICTITEM1 where ICTITEM1.ITEM_CODE = WHTLOCB0.ITEM_CODE" _
            & Replace(SQLW, "X.WHSE_CODE", "WHTLOCB0.WHSE_CODE") _
            & " group by WHTLOCB0.ITEM_CODE, WHTLOCB0.WHSE_CODE, WHTLOCB0.LOCATION_CODE" & vbCrLf _
            & ") group by ITEM_CODE, WHSE_CODE, LOCATION_CODE) X" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = X.ITEM_CODE" & vbCrLf _
            & " and ICTCOSTA.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
            & " and ICTCOSTA.OPS_YYYYPP (+) = '" & LYP & "'"
        WHTPHYV1 = ASCMAIN1.Temp_Table(ASCMAIN1.sql)

        ASCDATA1.ExecuteSQL("Alter Table " & WHTPHYV1 & " Add Primary Key (ITEM_CODE,WHSE_CODE,LOCATION_CODE)")
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If tblASTDSQLA.Select("SEQUENCE is Not Null", "SEQUENCE").Length = 0 Then
                EMsg &= vbCr & "You must Specify at least 1 Sort Field"
            End If

            If chkUpdateVariances.Checked Then
                For Each rowASTDSQLA As DataRow In tblASTDSQLA.Select("")
                    Dim COLUMN_NAME As String = rowASTDSQLA.Item("COLUMN_NAME")

                    If COLUMN_NAME = "WHSE_CODE" Then
                        If Val(rowASTDSQLA.Item("SEQUENCE") & "") <> 1 Then
                            EMsg &= vbCr & "Warehouse MUST be the 1st sort sequence when updating"
                        End If
                    End If
                    If rowASTDSQLA.Item("CODE_VALUES") & "" <> "" Then
                        If COLUMN_NAME = "WHSE_CODE" Then
                            ' this is ok
                        Else
                            EMsg &= vbCr & "You may NOT specify filter criteria for any field (other than Warehouse) when updating"
                        End If
                    End If
                Next
                If optSORT.Value <> "I" Then
                    EMsg &= vbCr & "Invalid Sort Option when enabling Update - must be by Item Code"
                End If
            End If
        End If
    End Sub

    Overrides Sub Update_Record()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from " & ICTPHYV1 & ";" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update ICTSTAT1 Set WHSE_QTY_BEG = R1.PHYS where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & ASCMAIN1.CYP & "';" & vbCrLf _
            & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (ITEM_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_BEG) Values (R1.ITEM_CODE,R1.WHSE_CODE,'" & ASCMAIN1.CYP & "',R1.PHYS); End If;" & vbCrLf _
            & "   Update ICTSTAT1 Set WHSE_QTY_PHY = NVL(WHSE_QTY_PHY,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & LYP & "';" & vbCrLf _
            & "   If SQL%NOTFOUND Then Insert into ICTSTAT1 (ITEM_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_PHY) Values (R1.ITEM_CODE,R1.WHSE_CODE,'" & LYP & "',NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
            & "   Update ICTSTAT2 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then Insert into ICTSTAT2 (ITEM_CODE,WHSE_CODE,WHSE_QTY_ON_HAND) Values (R1.ITEM_CODE,R1.WHSE_CODE,NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
            & "   Update ICTSTAT5 Set WHSE_QTY_ON_HAND = NVL(WHSE_QTY_ON_HAND,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE and OPS_YYYYPP = '" & LYP & "';" & vbCrLf _
            & "   If SQL%NOTFOUND Then Insert into ICTSTAT5 (ITEM_CODE,WHSE_CODE,OPS_YYYYPP,WHSE_QTY_ON_HAND) Values (R1.ITEM_CODE,R1.WHSE_CODE,'" & LYP & "',NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
            & "  End Loop; " & vbCrLf _
            & " End; " & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "" _
            & "Begin Declare Cursor C1 is Select * from " & WHTPHYV1 & " where NVL(PHYS,0) - NVL(BOOK,0) <> 0;" & vbCrLf _
            & " Begin " & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update WHTLOCB1 Set LOCATION_QTY = NVL(LOCATION_QTY,0) + NVL(R1.PHYS,0) - NVL(R1.BOOK,0) where ITEM_CODE = R1.ITEM_CODE and WHSE_CODE = R1.WHSE_CODE and LOCATION_CODE = R1.LOCATION_CODE;" & vbCrLf _
            & "   If SQL%NOTFOUND Then Insert into WHTLOCB1 (WHSE_CODE,LOCATION_CODE,BAR_CODE,ITEM_CODE,LOCATION_QTY) Values (R1.WHSE_CODE,R1.LOCATION_CODE,'0000000000',R1.ITEM_CODE,NVL(R1.PHYS,0) - NVL(R1.BOOK,0)); End If;" & vbCrLf _
            & "   Insert into WHTLOCB2 (WHSE_CODE,LOCATION_CODE,BAR_CODE,ITEM_CODE,WHSE_TRAN_QTY,WHSE_TRAN_TYPE,WHSE_TRAN_NO,WHSE_TRAN_LNO,INIT_DATE,INIT_OPER,LOCATION_CODE_OTHER,SESSION_NO) " & vbCrLf _
            & "    Values (R1.WHSE_CODE,R1.LOCATION_CODE,'0000000000',R1.ITEM_CODE,NVL(R1.PHYS,0) - NVL(R1.BOOK,0),'P','0000000000',0,SYSDATE,'" & ASCMAIN1.USER_ID & "',NULL,'" & ASCMAIN1.SESSION_NO & "');" & vbCrLf _
            & "  End Loop; " & vbCrLf _
            & " End; " & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Dim SQLW As String = ""
        If WHSE_CODEs <> "" Then
            SQLW = " and X.WHSE_CODE in ('" & Replace(WHSE_CODEs, ",", "','") & "')"
        End If
        SQLW &= " and X.WHSE_CODE in (Select WHSE_CODE from ICTWHSE1 where WHSE_PHYS_STATUS = 'C')"

        ASCMAIN1.sql = "" _
          & "Update ICTWHSE1 X Set WHSE_YYYYPP_LAST_PHY = '" & LYP & "', WHSE_PHYS_STATUS = NULL" & ASCMAIN1.SQL_Add_WHERE(SQLW)
        ASCDATA1.ExecuteSQL()

        ASCDATA1.ExecuteSQL("Insert into ICTCOSTA Select * from " & ICTCOSTA & " where ITEM_EXP_IMP_IND = 'V'")

        ' GL_Update()

    End Sub

    Public Function Prepare_GL_Interface() As String

        ' Prepare GL Interface File

        Create_TDA(dst.Tables.Add, "GLTINTF1", "*")


        Dim JOURNAL_TYPE As String = "ICPH"
        Dim JOURNAL_NO As String = ASCMAIN1.Next_Control_No("GLTJRNL1.JOURNAL_NO")
        Dim JOURNAL_LNO As Integer = 0

        Dim DETL_POSTING_AMT As Decimal
        Dim DETL_CTL_DATE As Date = DateValue(Format(DATETIME_STAMP, "MM/dd/yyyy"))

        Dim YP As String = ""

        ASCMAIN1.sql = "SELECT ICTPHYV1.COST_CATGY_CODE, ICTPHYV1.WHSE_CODE" & vbCrLf _
            & ", SUM ((NVL(ICTPHYV1.PHYS,0) - NVL(ICTPHYV1.BOOK,0)) * ICTPHYV1.ITEM_COST_TOTAL) VARLYP" & vbCrLf _
            & ", SUM ((NVL(ICTPHYV1.PHYS,0) - NVL(ICTPHYV1.BOOK,0)) * ICTCOSTA.ITEM_COST_TOTAL) VARTYP" & vbCrLf _
            & " from " & ICTPHYV1 & " ICTPHYV1, " & ICTCOSTA & " ICTCOSTA" & vbCrLf _
            & " where ICTCOSTA.ITEM_CODE (+) = ICTPHYV1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOSTA.OPS_YYYYPP (+) = '" & ASCMAIN1.CYP & "'" & vbCrLf _
            & "group by ICTPHYV1.COST_CATGY_CODE, ICTPHYV1.WHSE_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows

            Dim VARLYP As Decimal = Val(row.Item("VARLYP") & "")
            Dim VARTYP As Decimal = Val(row.Item("VARLYP") & "")

            If VARLYP <> 0 Or VARTYP <> 0 Then
                Dim COST_CATGY_CODE As String = row.Item("COST_CATGY_CODE") & ""
                Dim WHSE_CODE As String = row.Item("WHSE_CODE") & ""
                Dim rowICTCOST1 As DataRow = LookUp("ICTCOST1", COST_CATGY_CODE)
                Dim ACCT_CODE As String = ""

                DETL_POSTING_AMT = Val(row.Item("VARLYP") & "")
                If DETL_POSTING_AMT <> 0 Then
                    ACCT_CODE = rowICTCOST1.Item("ACCT_CODE_ONH")
                    Write_GLTINTF1(LYP, JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, _
                           ACCT_CODE, WHSE_CODE, COST_CATGY_CODE, DETL_POSTING_AMT)
                    ACCT_CODE = "5165"
                    Write_GLTINTF1(LYP, JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, _
                           ACCT_CODE, WHSE_CODE, COST_CATGY_CODE, -1 * DETL_POSTING_AMT)
                End If

                DETL_POSTING_AMT = Val(row.Item("VARTYP") & "") - Val(row.Item("VARLYP") & "")
                If DETL_POSTING_AMT <> 0 Then
                    ACCT_CODE = rowICTCOST1.Item("ACCT_CODE_ONH")
                    Write_GLTINTF1(ASCMAIN1.CYP, JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, _
                           ACCT_CODE, WHSE_CODE, COST_CATGY_CODE, DETL_POSTING_AMT)
                    ACCT_CODE = "5165"
                    Write_GLTINTF1(ASCMAIN1.CYP, JOURNAL_TYPE, JOURNAL_NO, JOURNAL_LNO, DETL_CTL_DATE, _
                           ACCT_CODE, WHSE_CODE, COST_CATGY_CODE, -1 * DETL_POSTING_AMT)
                End If
            End If
        Next

        Return JOURNAL_NO

    End Function

    Sub Write_GLTINTF1(OPS_YYYYPP As String, JOURNAL_TYPE As String, _
                       JOURNAL_NO As String, ByRef JOURNAL_LNO As Integer, DETL_CTL_DATE As Date, _
                       ACCT_CODE As String, WHSE_CODE As String, COST_CATGY_CODE As String, _
                       DETL_POSTING_AMT As Decimal)

        Dim rowGLTINTF1 As DataRow = dst.Tables("GLTINTF1").NewRow
        rowGLTINTF1("OPS_YYYYPP") = OPS_YYYYPP
        rowGLTINTF1("JOURNAL_NO") = JOURNAL_NO
        JOURNAL_LNO += 1
        rowGLTINTF1("JOURNAL_LNO") = JOURNAL_LNO
        rowGLTINTF1("ACCT_CODE") = ACCT_CODE
        rowGLTINTF1("SEG2_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG2")
        rowGLTINTF1("SEG3_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG3")
        rowGLTINTF1("SEG4_CODE") = ROWs("GLTPARM1").Item("GL_PARM_DEF_SEG4")
        rowGLTINTF1("DETL_CTL_DATE") = DETL_CTL_DATE
        rowGLTINTF1("DETL_POSTING_AMT") = Math.Round(DETL_POSTING_AMT, 2)
        rowGLTINTF1("DETL_EXE_NO") = XNO
        rowGLTINTF1("DETL_CTL_NO") = DBNull.Value
        rowGLTINTF1("DETL_CTL_LNO") = DBNull.Value
        rowGLTINTF1("DETL_CVX_NO") = WHSE_CODE
        rowGLTINTF1("DETL_CVX_REF_DATE") = DBNull.Value
        rowGLTINTF1("DETL_CVX_REF_NO") = COST_CATGY_CODE
        rowGLTINTF1("DETL_DESC") = DBNull.Value
        rowGLTINTF1("DETL_CVX_TYPE") = "W"
        rowGLTINTF1("JOURNAL_TYPE") = JOURNAL_TYPE
        dst.Tables("GLTINTF1").Rows.Add(rowGLTINTF1)
    End Sub
End Class