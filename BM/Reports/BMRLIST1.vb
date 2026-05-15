Public Class BMRLIST1
    Dim BM_ISSUE_STATUS As String = ""

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ICTPARM1")
     
    End Sub

    Protected Overrides Sub Build_Workfile()

        SUBT = ""
        BM_ISSUE_STATUS = optBM_ISSUE_STATUS.Value

        Dim sqlw As String = SQL_in("BM_PROD_ITEM")
        ' Get_Item_and_Qtys()
        Prepare_dst(True, New String() {sqlw, BM_ISSUE_STATUS})

        Check_if_Empty("BMTMAIN1")
    End Sub

    Public Overrides Sub Print_Report()
        'CR_params.Add("RUNQTY", IIf(Not optPrintOptions.Value = "Q", "0", CStr(numRunQty.Value & "")))
        'CR_params.Add("STATUS", IIf(Not optPrintOptions.Value = "Q", "0", "1"))
        'CR_params.Add("COSTED_BOM", IIf(optPrintOptions.Value = "C", "1", "0"))

        CR_params.Add("RUNQTY", "0")
        CR_params.Add("STATUS", "0")
        CR_params.Add("COSTED_BOM", "0")


        CR_params.Add("NOTES", "1")
        CR_params.Add("COMPNOTES", "1")

        ' Generate_Report("BMRLIST1", "Bill of Materials", "")
        Generate_Report(RPT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
           
        End If
    End Sub
     
    Overrides Function Prepare_dst( _
    ByVal perform_fill As Boolean, _
    ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        With dst

            ASCMAIN1.sql = "Select BMTMAIN1.*" _
                & " from BMTMAIN1 where BM_PROD_ITEM = :PARM1"
            Create_TDA(.Tables.Add, "BMTMAIN1", "**", 0, True, "V", 1)
            .Tables("BMTMAIN1").Columns.Add("IMAGE", GetType(System.Byte()))

            ASCMAIN1.sql = "Select BMTMAIN2.*" _
                & " from BMTMAIN2 where BM_PROD_ITEM = :PARM1 and BM_ISSUE_NO = :PARM2"
            Create_TDA(.Tables.Add, "BMTMAIN2", "**", 0, True, "VN", 2)

            ASCMAIN1.sql = "Select BMTMAIN3.*" & vbCrLf _
                & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_STD" & vbCrLf _
                & ", ICTITEM1.ITEM_COST_WASTE_PCT" & vbCrLf _
                & ", ICTITEM1.ITEM_PLAN_WASTE_PCT" & vbCrLf _
                & ", ICTITEM1.VEND_CODE" & vbCrLf _
                & ", ICTITEM1.VEND_ITEM_CODE" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_VCOST" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_LANDG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOOLG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_OVRHD" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOTAL" & vbCrLf _
                & " from BMTMAIN3,ICTITEM1,ICTCOSTC" & vbCrLf _
                & " where BMTMAIN3.BM_PROD_ITEM = :PARM1" & vbCrLf _
                & " and BMTMAIN3.BM_ISSUE_NO = :PARM2" & vbCrLf _
                & " and ICTITEM1.ITEM_CODE = BMTMAIN3.BM_COMP_ITEM" & vbCrLf _
                & " and ICTCOSTC.ITEM_CODE (+) = BMTMAIN3.BM_COMP_ITEM"
            Create_TDA(.Tables.Add, "BMTMAIN3", "**", 0, True, "VN", 3)
            Dim CALC As String = "ISNULL(BM_QTY_PER_ASSY,0) * ISNULL(?,0) * (1 + ISNULL(ITEM_COST_WASTE_PCT,0)/100)"
            .Tables("BMTMAIN3").Columns.Add("EXT_COST", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_STD"))
            .Tables("BMTMAIN3").Columns.Add("VCOST", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_VCOST"))
            .Tables("BMTMAIN3").Columns.Add("LANDG", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_LANDG"))
            .Tables("BMTMAIN3").Columns.Add("TOOLG", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_TOOLG"))
            .Tables("BMTMAIN3").Columns.Add("OVRHD", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_OVRHD"))
            .Tables("BMTMAIN3").Columns.Add("TOTAL", GetType(System.Decimal), Replace(CALC, "?", "ITEM_COST_TOTAL"))
            .Tables("BMTMAIN3").Columns.Add("QTY_ON_HAND", GetType(System.Int32))
            .Tables("BMTMAIN3").Columns.Add("QTY_ONPO", GetType(System.Int32))
            .Tables("BMTMAIN3").Columns.Add("QTY_PLAN", GetType(System.Int32))
            .Tables("BMTMAIN3").Columns.Add("QTY_OPEN", GetType(System.Int32))
            .Tables("BMTMAIN3").Columns.Add("QTY_PICK", GetType(System.Int32))
            .Tables("BMTMAIN3").Columns.Add("QTY_COMM", GetType(System.Int32))
            .Tables("BMTMAIN3").Columns.Add("QTY_OPEN_PICK", GetType(System.Int32), "ISNULL(QTY_OPEN,0)+ISNULL(QTY_PICK,0)")
            .Tables("BMTMAIN3").Columns.Add("QTY_AVA", GetType(System.Int32), "ISNULL(QTY_ON_HAND,0)+ISNULL(QTY_OPEN,0)+ISNULL(QTY_PLAN,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_OPEN,0)-ISNULL(QTY_PICK,0)")
            .Tables("BMTMAIN3").Columns.Add("BM_COMPONENT_SORT")

            ASCMAIN1.sql = "Select ICTCOSTC.*" _
                & " from ICTCOSTC where ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTCOSTC", "**", 0, False, "V", 1)

            ASCMAIN1.sql = "Select 'X' COST_TYPE" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_VCOST VCOST" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_LANDG LANDG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOOLG TOOLG" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_OVRHD OVRHD" & vbCrLf _
                & ", ICTCOSTC.ITEM_COST_TOTAL TOTAL" & vbCrLf _
                & " from ICTCOSTC"
            Create_TDA(.Tables.Add, "ICTCOSTS", "**", 0, False, "", 1)

            Dim sqlICTSTAT2 As String = "Select ITEM_CODE" & vbCrLf _
                & ", Sum (WHSE_QTY_ON_HAND) QTY_ON_HAND" & vbCrLf _
                & ", Sum (WHSE_QTY_ONPO) QTY_ONPO" & vbCrLf _
                & ", Sum (WHSE_QTY_PLAN) QTY_PLAN" & vbCrLf _
                & ", Sum (WHSE_QTY_OPEN) QTY_OPEN" & vbCrLf _
                & ", Sum (WHSE_QTY_PICK) QTY_PICK" & vbCrLf _
                & ", Sum (WHSE_QTY_COMM) QTY_COMM" & vbCrLf _
                & " from ICTSTAT2 where ITEM_CODE = :PARM1 group by ITEM_CODE"

            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_UOM, ICTITEM1.ITEM_COST_STD" & vbCrLf _
                & ", ICTITEM1.ITEM_PLAN_MAKE_BUY, ICTITEM1.ITEM_PLAN_WASTE_PCT, ICTITEM1.VEND_ITEM_CODE" & vbCrLf _
                & ", X.QTY_ON_HAND, X.QTY_ONPO, X.QTY_OPEN, X.QTY_COMM, X.QTY_PICK, X.QTY_PLAN" & vbCrLf _
                & " from ICTITEM1, (" & sqlICTSTAT2 & ") X" & vbCrLf _
                & " where X.ITEM_CODE (+) = ICTITEM1.ITEM_CODE" & vbCrLf _
                & "   and ICTITEM1.ITEM_CODE = :PARM1"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", 0, False, "V", 1)
            .Tables("ICTITEM1").Columns.Add("QTY_OPEN_PICK", GetType(Int64), "ISNULL(QTY_OPEN,0)+ISNULL(QTY_COMM,0)+ISNULL(QTY_PICK,0)")
            .Tables("ICTITEM1").Columns.Add("QTY_AVA", GetType(Int64), "ISNULL(QTY_ON_HAND,0)+ISNULL(QTY_OPEN,0)+ISNULL(QTY_PLAN,0)-ISNULL(QTY_COMM,0)-ISNULL(QTY_OPEN,0)-ISNULL(QTY_PICK,0)")
        End With

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = ""
        If parms.Length > 0 Then
            sqlw = parms(0)
            BM_ISSUE_STATUS = parms(1)
        End If

        EnforceConstraints(False)


        For Each TABLE_NAME As String In New String() {"BMTMAIN1", "BMTMAIN2", "BMTMAIN3", "ICTITEM1", "ICTCOSTC"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        If Not ROWs.ContainsKey("ICTPARM1") Then
            Get_PARM("ICTPARM1")
        End If

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""

        ASCMAIN1.sql = "Select * from BMTMAIN1" & ASCMAIN1.SQL_Add_WHERE(sqlw)
        For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
            Dim ITEM_CODE As String = row.Item("BM_PROD_ITEM")
            Fill_Records("BMTMAIN1", ITEM_CODE, False)
            Dim rowBMTMAIN1 As DataRow = dst.Tables("BMTMAIN1").Rows.Find(ITEM_CODE)
            Dim BM_ISSUE_NO As String = "00"
            Dim BM_ISSUE_COUNTER As Integer = Val(row.Item("BM_ISSUE_COUNTER") & "")
            If BM_ISSUE_STATUS = "C" Then BM_ISSUE_NO = Format(BM_ISSUE_COUNTER, "00")
            If BM_ISSUE_STATUS = "S" Then BM_ISSUE_NO = row.Item("BM_ISSUE_STD") & ""

            Dim IMAGE_NAME As String = ITEM_CODE
            Dim imgba() As Byte = Nothing
            ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, True, , , imgba)
            'rowBMTMAIN1.Item("IMAGE") = imgba

            Fill_Records("BMTMAIN2", New String() {ITEM_CODE, BM_ISSUE_NO}, False)
            Fill_Records("BMTMAIN3", New String() {ITEM_CODE, BM_ISSUE_NO}, False)
            Fill_Records("ICTITEM1", ITEM_CODE, False)
            Fill_Records("ICTCOSTC", ITEM_CODE, False)
        Next

        For Each row As DataRow In dst.Tables("BMTMAIN3").Select("")
            Dim ITEM_CODE As String = row.Item("BM_COMP_ITEM")
            If dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE) Is Nothing Then
                Fill_Records("ICTITEM1", ITEM_CODE, False)
                Fill_Records("ICTCOSTC", ITEM_CODE, False)
            End If
        Next

        EnforceConstraints(True)
    End Sub
End Class