Public Class POROPRT1

#Region "General Declarations"
    Private xDTE0 As Date
    Private xDTE1 As Date

    Dim SQLs As New Dictionary(Of String, String)

    Dim POTORDR1 As String
    Dim POTORDR2 As String

    Dim sqlPOTORDR1 As String
    Dim sqlPOTORDR2 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")
        Absx1.optFor("RANGE").CheckedIndex = 2
    End Sub

    Protected Overrides Sub Build_Workfile()

        RWU = "R"
        Dim sqlw As String = ""

        If Absx1.optFor("RANGE").Value = "D" Then
            xDTE0 = Absx1.dteFor("DTE0").Value
            xDTE1 = Absx1.dteFor("DTE1").Value
            If System.DateTime.Compare(xDTE0, xDTE1) = 0 Then
                SUBT = "POs Dated " & Format(xDTE0, "MM/dd/yyyy")
            Else
                SUBT = "POs Dated between " & Format(xDTE0, "MM/dd/yyyy") & " and " & Format(xDTE1, "MM/dd/yyyy")
            End If
            sqlw = " and POTORDR1.PO_DATE_ORDERED between '" & Format(xDTE0, "dd-MMM-yyyy") & "' and '" & Format(xDTE1, "dd-MMM-yyyy") & "'" & vbCrLf
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "S" Then
            SUBT = "Selected POs"
            RWU = "N"
        ElseIf Absx1.optFor("RANGE").Value = "U" Then
            SUBT = "All POs not Printed Yet"
            sqlw &= " and POTORDR1.PO_PRINTED_IND = '0'" & vbCrLf

        End If

        sqlw &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
        sqlw &= SQL_in("PO_ORDER_NO", "POTORDR1.PO_ORDER_NO")

        Prepare_dst(True, sqlw)

        Check_if_Empty("POTORDR1")
    End Sub

    Public Overrides Sub Print_Report()
        Dim PO_PARM_PO_RPT As String = ROWs("POTPARM1").Item("PO_PARM_PO_RPT") & ""
        If PO_PARM_PO_RPT <> "" Then RPT = PO_PARM_PO_RPT

        CR_params.Add("SUBT", "")
        CR_params.Add("FORM_TYPE", "P")
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If Absx1.optFor("RANGE").Value = "D" Then
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
        grpDATE_RANGE.Enabled = (optRANGE.Value = "D")

        If optRANGE.Value = "D" Then
            Dim dates() As Date = ASCMAIN1.Get_Dates(ASCMAIN1.CYP)
            Absx1.dteFor("DTE0").Value = dates(1)
            Absx1.dteFor("DTE1").Value = dates(UBound(dates))
        End If
    End Sub

    Overrides Sub Update_Record()
        Dim sql As String = ""
        ' MAYBE WE DON'T WANT TO SET THE PRINTED FLAG BECAUSE THIS WILL CAUSE REVISIONS TO START COUNTING
        sql = "Update POTORDR1 " _
            & " Set PO_PRINTED_IND = '1', PO_DATE_PRINTED = SYSDATE" _
            & " where (PO_ORDER_NO) in (Select PO_ORDER_NO from " & POTORDR1 & " )"
        ' ASCDATA1.ExecuteSQL(sql)

        sql = "Insert into TATEVNT1 (TABLE_NAME, TABLE_KEY, INIT_DATE, INIT_OPER, EVENT_TYPE, EVENT_DESC, EVENT_KEY)" _
            & " Select 'POTORDR1', PO_ORDER_NO, SYSDATE, '" & ASCMAIN1.USER_ID & "', 'PO_PRT','PO Printed', PO_ORDER_NO" _
            & " from " & POTORDR1
        ASCDATA1.ExecuteSQL(sql)
    End Sub

    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Get_PARM("POTPARM1")

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        sqlPOTORDR1 = "Select POTORDR1.* from POTORDR1 "
        ASCMAIN1.sql = sqlPOTORDR1 & ASCMAIN1.SQL_Add_WHERE(sqlw)
        POTORDR1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR1 & " Add Primary Key (PO_ORDER_NO)")

        sqlPOTORDR2 = "Select POTORDR2.* from POTORDR2, " & POTORDR1 _
            & " POTORDR1 where POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
        ASCMAIN1.sql = sqlPOTORDR2
        POTORDR2 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & POTORDR2 & " Add Primary Key (PO_ORDER_NO, PO_ORDER_LNO)")

        SQLs.Clear()

        With dst
            ASCMAIN1.sql = "Select POTORDR1.*, 'Z' PO_PARM_KEY" & vbCrLf _
                & " from " & POTORDR1 & " POTORDR1"
            SQLs.Add("POTORDR1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR1", "**", 0, False, "", 1)


            ASCMAIN1.sql = "Select POTORDR2.*" & vbCrLf _
                & " from " & POTORDR2 & " POTORDR2"
            SQLs.Add("POTORDR2", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR2", "**", 0, False, "", 2)

            Create_Relation("POTORDR1", "POTORDR2", "PO_ORDER_NO")
            dst.Tables("POTORDR2").Columns.Add("LINE_AMOUNT", GetType(System.Decimal), "ISNULL(PO_QTY_ORD,0) * ISNULL(PO_COST,0)")
            dst.Tables("POTORDR1").Columns.Add("PO_TOTAL_INVTY", GetType(System.Decimal), "SUM(CHILD(POTORDR1_POTORDR2).LINE_AMOUNT)")

            ASCMAIN1.sql = "Select POTORDR5.*" & vbCrLf _
                & " from " & POTORDR1 & " POTORDR1,POTORDR5" & vbCrLf _
                & " where POTORDR5.PO_ORDER_NO = POTORDR1.PO_ORDER_NO"
            SQLs.Add("POTORDR5", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDR5", "**", 0, False, "", 2)


            Create_Relation("POTORDR1", "POTORDR5", "PO_ORDER_NO")
            dst.Tables("POTORDR1").Columns.Add("PO_TOTAL_NINVTY", GetType(System.Decimal), "SUM(CHILD(POTORDR1_POTORDR5).PO_NINV_AMOUNT)")
            dst.Tables("POTORDR1").Columns.Add("PO_TOTAL", GetType(System.Decimal), "ISNULL(PO_TOTAL_INVTY,0) + ISNULL(PO_TOTAL_NINVTY,0)")


            ASCMAIN1.sql = "Select POTORDRZ.*" & vbCrLf _
                & " from POTORDRZ" & vbCrLf _
                & " where PO_ORDER_NO in (Select PO_ORDER_NO from " & POTORDR1 & ")"
            SQLs.Add("POTORDRZ", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "POTORDRZ", "**", 0, False, "", 3)
            With .Tables("POTORDRZ").Columns
                .Add("ITEM_CODE_PREV")
                .Add("PO_QTY_ORD_PREV", GetType(System.Int64))
                .Add("PO_COST_PREV", GetType(System.Decimal))
                .Add("PO_DATE_REQUIRED_PREV", GetType(System.DateTime))
                .Add("PO_STATUS_PREV")
                .Add("CARTON_PACK_QTY_PREV", GetType(System.Int64))
            End With

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_STD, ICTITEM1.ITEM_MATL_DESC, ICTITEM1.COUNTRY_CODE, ICTITEM1.VEND_ITEM_CODE, ICTITEM1.ITEM_EAN_CODE" & vbCrLf _
                & " from ICTITEM1" & vbCrLf _
                & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & POTORDR2 & ")"
            SQLs.Add("ICTITEM1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select DPTVNDI1.*" & vbCrLf _
                & " from DPTVNDI1" & vbCrLf _
                & " where ITEM_CODE in (Select Distinct ITEM_CODE from " & POTORDR2 & ")"
            SQLs.Add("DPTVNDI1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "DPTVNDI1", "**", 0, False, "", 2)

            ASCMAIN1.sql = "Select APTVEND1.*" & vbCrLf _
                & " from APTVEND1" & vbCrLf _
                & " where VEND_CODE in (Select Distinct VEND_CODE from " & POTORDR1 & ")"
            SQLs.Add("APTVEND1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "APTVEND1", "**", 0, False, "", 1)

            ASCMAIN1.sql = "Select ASTUSER1.USER_ID, ASTUSER1.USER_NAME" & vbCrLf _
                & " from ASTUSER1" & vbCrLf _
                & " where ASTUSER1.USER_ID in (Select Distinct INIT_OPER from " & POTORDR1 & ")"
            SQLs.Add("ASTUSER1", ASCMAIN1.sql)
            Create_TDA(.Tables.Add, "ASTUSER1", "**", 0, False, "", 1)

            For Each TABLE_NAME As String In New String() _
            {"TATTERM1", "ICTWHSE1", "TATCNTRY"}
                Create_TDA(.Tables.Add, TABLE_NAME, "*", 0, False)
                Fill_Records(TABLE_NAME)
            Next

            ASCMAIN1.sql = "Select POTPARM1.*" & vbCrLf _
                & " from POTPARM1 where PO_PARM_KEY = 'Z'"
            Create_TDA(.Tables.Add, "POTPARM1", "**", 0, False, "", 1)
            .Tables("POTPARM1").Columns.Add("LOGO", GetType(System.Byte()))

        End With

        Fill_Records("POTPARM1")
        Dim rowPOTPARM1 As DataRow = dst.Tables("POTPARM1").Rows(0)
        rowPOTPARM1.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & ".PNG")

        If ASCMAIN1.CLIENT = "NYA" Then
            dst.Tables("POTPARM1").Columns.Add("PO_PARM_FORM_COUNTRY")
            dst.Tables("POTPARM1").Columns.Add("COMP_TAX_ID")
            dst.Tables("POTPARM1").Columns("PO_PARM_KEY").MaxLength = -1
            dst.Tables("POTORDR1").Columns("PO_PARM_KEY").MaxLength = -1

            ASCMAIN1.sql = "Select * from SOTCOMP1"
            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim COMP_CODE As String = row.Item("COMP_CODE")
                Dim rowP As DataRow = dst.Tables("POTPARM1").NewRow
                rowP.Item("PO_PARM_KEY") = "Z" & COMP_CODE
                For Each C As String In New String() {"COMP_NAME", "COMP_ADDR1", "COMP_ADDR2", "COMP_ADDR3",
                                                      "COMP_CITY", "COMP_STATE", "COMP_ZIP_CODE", "COMP_COUNTRY",
                                                      "COMP_PHONE", "COMP_FAX", "COMP_EMAIL", "COMP_TAX_ID"}
                    Dim CP As String = Replace(C, "COMP_", "PO_PARM_FORM_")
                    rowP.Item(CP) = row.Item(C)
                Next
                rowP.Item("LOGO") = ASCMAIN1.GetImageData(ASCMAIN1.Folders("Images") & "\ABS\" & ASCMAIN1.DBS_COMPANY & "_" & COMP_CODE & ".PNG")
                dst.Tables("POTPARM1").Rows.Add(rowP)
            Next
        End If


        If perform_fill Then
            Fill_Records_RPT(New String() {sqlw})
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        sqlw = parms(0)

        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR1)
        ASCDATA1.ExecuteSQL("Truncate Table " & POTORDR2)

        ASCDATA1.ExecuteSQL("Insert into " & POTORDR1 & " " & sqlPOTORDR1 & ASCMAIN1.SQL_Add_WHERE(sqlw))
        ASCDATA1.ExecuteSQL("Insert into " & POTORDR2 & " " & sqlPOTORDR2)

        If RWU = "R" Then
            Dim POS_NOT_LOCKABLE As New List(Of String)
            ASCMAIN1.sql = "Select PO_ORDER_NO from " & POTORDR1
            For Each row As DataRow In ASCDATA1.GetDataTable.Rows
                Dim PO_ORDER_NO As String = row.Item("PO_ORDER_NO")
                If Not ASCMAIN1.Logical_Lock("POTORDR1", PO_ORDER_NO) Then
                    POS_NOT_LOCKABLE.Add(PO_ORDER_NO)
                End If
            Next
            If POS_NOT_LOCKABLE.Count <> 0 Then
                ASCMAIN1.sql = "Delete from " & POTORDR1 & " where PO_ORDER_NO in (" & Join(POS_NOT_LOCKABLE.ToArray, "','") & ")"
                ASCDATA1.ExecuteSQL()
                ASCMAIN1.sql = "Delete from " & POTORDR2 & " where PO_ORDER_NO in (" & Join(POS_NOT_LOCKABLE.ToArray, "','") & ")"
                ASCDATA1.ExecuteSQL()
            End If
        End If

        EnforceConstraints(False)
        Fill_Records("POTORDR1")

        '  TAC.POCMAIN1.Setup_PO_Change_Details(Me)

        Fill_Records("POTORDR2")
        Fill_Records("POTORDRZ")
        Fill_Records("POTORDR5")

        If ASCMAIN1.CLIENT = "NYA" Then
            For Each row1 As DataRow In dst.Tables("POTORDR1").Select("")
                Dim P As String = row1.Item("PO_ORDER_NO")
                For Each row2 As DataRow In dst.Tables("POTORDR2").Select("PO_ORDER_NO = '" & P & "'")
                    Dim STYLE_CODE As String = row2.Item("STYLE_CODE")
                    Dim rowICTSTYL1 As DataRow = LookUp("ICTSTYL1", STYLE_CODE)
                    Dim SALES_DIVISION_CODE As String = rowICTSTYL1.Item("SALES_DIVISION_CODE") & ""
                    Dim rowSOTSDIV1 As DataRow = LookUp("SOTSDIV1", SALES_DIVISION_CODE)
                    If rowSOTSDIV1 IsNot Nothing AndAlso rowSOTSDIV1.Item("SEG4_CODE") & "" <> "" Then
                        row1.Item("PO_PARM_KEY") = "Z" & rowSOTSDIV1.Item("SEG4_CODE")
                    End If
                    Exit For
                Next
            Next
        End If

        For Each row As DataRow In dst.Tables("POTORDRZ").Select("")
            Dim rowP As DataRow = dst.Tables("POTORDRZ").Rows.Find(New Object() {row.Item("PO_ORDER_NO"), Val(row.Item("PO_HDR_CTR_REV") & "") - 1, row.Item("PO_ORDER_LNO")})
            If rowP IsNot Nothing Then
                row.Item("PO_QTY_ORD_PREV") = rowP.Item("PO_QTY_ORD")
                row.Item("PO_COST_PREV") = rowP.Item("PO_COST")
                row.Item("PO_DATE_REQUIRED_PREV") = rowP.Item("PO_DATE_REQUIRED")
                row.Item("PO_STATUS_PREV") = rowP.Item("PO_STATUS")
            End If
        Next

        Fill_Records("ICTITEM1")
        Fill_Records("ASTUSER1")
        Fill_Records("DPTVNDI1")
        Fill_Records("APTVEND1")
        EnforceConstraints(True)
    End Sub

    Overrides Function Get_Custom_Filter_for_Codes_Selection(COLUMN_NAME As String) As String
        Dim sqlw As String = ""
        Select Case COLUMN_NAME
            Case "VEND_CODE"
                sqlw = "APTVEND1.VEND_CODE in (Select Distinct VEND_CODE from POTORDR1)"
        End Select
        Return sqlw
    End Function
End Class