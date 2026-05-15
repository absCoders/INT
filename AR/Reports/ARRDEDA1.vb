Public Class ARRDEDA1
    Dim ARTDEDA1 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP0", ASCMAIN1.CYP, 0, 0, 0)
        Set_cmbYP_Child("RYP1", 12, "RYP0", 0)
    End Sub

    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim sqlw As String = ""
        sqlw &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sqlw &= SQL_in("ITEM_CATGY_CODE", "ICTITEM1.ITEM_CATGY_CODE")
        sqlw &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")

        If Absx1.chkFor("CHK_MB_M").Checked And _
           Absx1.chkFor("CHK_MB_B").Checked Then
        Else
            sqlw &= " and ICTITEM1.ITEM_BASIC_PROMO in (" & Mid( _
                 IIf(Absx1.chkFor("CHK_MB_M").Checked, ",'M'", "") & _
                 IIf(Absx1.chkFor("CHK_MB_B").Checked, ",'B'", ""), 2) & ")"
        End If

        RWU = "R"

        Prepare_dst(True, sqlw, RYP)

        Check_if_Empty("ARTREAS1")
    End Sub

    Overrides Function Prepare_dst( _
      ByVal perform_fill As Boolean, _
      ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Dim sqlw As String = CStr(parms(0))
        If sqlw = "" Then sqlw = " and ROWNUM < 1"

        If ARTDEDA1 = "" Then Create_Temp_Data(sqlw)

        With dst
            ASCMAIN1.sql = "Select ICTITEM1.* from " & ARTDEDA1 & " ICTITEM1"
            Create_TDA(dst.Tables.Add("ICTITEM1"), ARTDEDA1, "**", 0, True, , 1)
            With .Tables("ICTITEM1").Columns
                .Add("DEMAND_QTY", GetType(System.Int64), "ISNULL(FORECAST,0)+ISNULL(PROD_COM,0)+ISNULL(PLAN_COM,0)")
                .Add("DEMAND_AMT", GetType(System.Decimal), "DEMAND_QTY * ISNULL(ITEM_COST_STD,0)")
                .Add("DEMAND_PCT", GetType(System.Decimal))
                .Add("DEMAND_PCT_CUM", GetType(System.Decimal))
            End With

            Create_TDA(dst.Tables.Add, "ICTCOLL1", "*", 0)
            Create_TDA(dst.Tables.Add, "ICTBRAN1", "*", 0)
            Create_TDA(dst.Tables.Add, "ICTCATG1", "*", 0)
            Create_TDA(dst.Tables.Add, "DPTABCP1", "*", 0)
            .Tables("DPTABCP1").Columns.Add("ABC_INDEX", GetType(System.Int64))
            .Tables("DPTABCP1").Columns.Add("ABC_PCT_CUM", GetType(System.Decimal))

            With .Tables.Add("DPTABCPG")
                .Columns.Add("ABC_GROUP")
                .Columns.Add("ABC_GROUP_DESC")
                .PrimaryKey = New DataColumn() {.Columns("ABC_GROUP")}
            End With

        End With

        Fill_Records("ICTCOLL1")
        Fill_Records("ICTBRAN1")


        Fill_Records("ICTCATG1")
        dst.Tables("ICTCATG1").Rows.Add(New String() {"*", "All Catgys"})
        dst.Tables("ICTCATG1").Rows.Add(New String() {"?", "Catgy Unknown"})

        Fill_Records("DPTABCP1")

        Dim ABC_INDEX As Int16 = 0
        Dim ABC_PCT_CUM As Decimal
        For Each rowDPTABCP1 As DataRow In dst.Tables("DPTABCP1").Select("", "ABC_CODE")
            ABC_INDEX += 1
            ABC_PCT_CUM += Val(rowDPTABCP1.Item("ABC_PCT_RANGE") & "")
            rowDPTABCP1.Item("ABC_INDEX") = ABC_INDEX
            rowDPTABCP1.Item("ABC_PCT_CUM") = ABC_PCT_CUM
        Next

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sqlw As String = parms(0)

        If sqlw <> "" Then
            Create_Temp_Data(sqlw)
        End If
        EnforceConstraints(False)
        Fill_Records("ARTDEDA1")
        EnforceConstraints(True)

        '     Calculate_ABC()

    End Sub

    Sub Create_Temp_Data(SQLW As String)

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC" & vbCrLf _
            & ", ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_COST_STD" & vbCrLf _
            & ", ICTITEM1.ITEM_UOM, ICTITEM1.VEND_CODE, ICTITEM1.ITEM_PO_QTY_MIN" & vbCrLf _
            & ", ICTITEM1.ITEM_MRP_PLANR_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE || ICTITEM1.ITEM_BASIC_PROMO || ICTITEM1.ITEM_COST_MAKE_BUY || ICTITEM1.ITEM_CATGY_CODE ABC_GROUP" & vbCrLf _
            & ", ICTITEM1.ITEM_CATGY_CODE, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_SNU_CODE, ICTITEM1.ITEM_BASIC_PROMO" & vbCrLf _
            & ", ICTITEM1.ITEM_COST_MAKE_BUY, ICTITEM1.ITEM_ABC_CODE" & vbCrLf _
            & ", ICTITEM1.ITEM_ABC_CODE ITEM_ABC_CODE_FUT" & vbCrLf _
            & " from ICTITEM1,ICTCOLL1" & vbCrLf _
            & " where ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & SQLW

        If ARTDEDA1 = "" Then
            ARTDEDA1 = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add FORECAST NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add PROD_COM NUMBER (8,0)")
            ASCDATA1.ExecuteSQL("Alter Table " & ARTDEDA1 & " Add PLAN_COM NUMBER (8,0)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ARTDEDA1)

            Dim COLUMN_NAMEs As String = "" _
                & "ITEM_CODE, ITEM_DESC, ITEM_RETAIL_PRICE, ITEM_COST_STD" _
                & ", ITEM_UOM, VEND_CODE, ITEM_PO_QTY_MIN, ITEM_MRP_PLANR_CODE, ABC_GROUP" _
                & ", ITEM_CATGY_CODE, COLLECTION_CODE, BRAND_CODE, ITEM_SNU_CODE, ITEM_BASIC_PROMO" _
                & ", ITEM_COST_MAKE_BUY, ITEM_ABC_CODE, ITEM_ABC_CODE_FUT"

            ASCDATA1.ExecuteSQL("Insert into " & ARTDEDA1 & " (" & COLUMN_NAMEs & ") " & ASCMAIN1.sql)

            Get_Demand_Data()
        End If

    End Sub

    Public Overrides Sub Print_Report()
        SUBT = ""
        CR_params.Add("MOS", Val(Absx1.numFor("FPDMOS").Value & ""))
        CR_params.Add("CD", "Demand Calculated using " & Absx1.optFor("OPTDEMAND").Text)
        CR_params.Add("EXTUSAGE", "Saleable Ranked by " & Absx1.optFor("OPTRANKBY").Text)
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
            If eItemKey = "Proceed" Then
                If Absx1.cmbFor("RYP0").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify a Starting Period"
                End If
                If Absx1.cmbFor("RYP1").Value & "" = "" Then
                    EMsg &= vbCr & "You must Specify an Ending Period"
                End If
                If EMsg = "" Then
                    If ASCMAIN1.Period_Diff(Absx1.cmbFor("RYP0").Value, Absx1.cmbFor("RYP1").Value) > 12 Then
                        EMsg &= vbCr & "Maximum number of periods spanned by Period Range is 12"
                    End If
                End If
            End If
        End If

    End Sub

    Overrides Sub Update_Record()

    End Sub

    Sub Get_Demand_Data()

        dst.Tables("DPTABCPG").Rows.Clear()

        Dim sqlABC_GROUP As String = ""
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_SNU").Checked, "NVL(ITEM_SNU_CODE,'?')", "'*'")
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_BP").Checked, "NVL(ITEM_BASIC_PROMO,'?')", "'*'")
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_MB").Checked, "NVL(ITEM_COST_MAKE_BUY,'?')", "'*'")
        sqlABC_GROUP &= " || " & IIf(Absx1.chkFor("CHK_GROUP_CATGY").Checked, "NVL(ITEM_CATGY_CODE,'?')", "'*'")
        'If Absx1.chkFor("CHK_GROUP_SNU").Checked Then ABC_GROUP &= " || NVL(ITEM_SNU_CODE,'?')"
        'If Absx1.chkFor("CHK_GROUP_BP").Checked Then ABC_GROUP &= " || NVL(ITEM_BASIC_PROMO,'?')"
        'If Absx1.chkFor("CHK_GROUP_MB").Checked Then ABC_GROUP &= " || NVL(ITEM_COST_MAKE_BUY,'?')"
        'If Absx1.chkFor("CHK_GROUP_CATGY").Checked Then ABC_GROUP &= " || NVL(ITEM_CATGY_CODE,'?')"

        ASCDATA1.ExecuteSQL("Update " & ARTDEDA1 & " Set ABC_GROUP = " & Mid(sqlABC_GROUP, 5))

        Dim ITEM_SNU_CODEs As New Dictionary(Of String, String)
        ITEM_SNU_CODEs.Add("*", "All SNU")
        ITEM_SNU_CODEs.Add("S", "Saleable")
        ITEM_SNU_CODEs.Add("N", "No-Charge")
        ITEM_SNU_CODEs.Add("U", "Unfinished")
        ITEM_SNU_CODEs.Add("?", "SNU Unknown")

        Dim ITEM_BASIC_PROMOs As New Dictionary(Of String, String)
        ITEM_BASIC_PROMOs.Add("*", "Basic & Promo")
        ITEM_BASIC_PROMOs.Add("B", "Basic")
        ITEM_BASIC_PROMOs.Add("P", "Promo")
        ITEM_BASIC_PROMOs.Add("?", "BP Unknown")

        Dim ITEM_COST_MAKE_BUYs As New Dictionary(Of String, String)
        ITEM_COST_MAKE_BUYs.Add("*", "Make & Buy")
        ITEM_COST_MAKE_BUYs.Add("M", "Make")
        ITEM_COST_MAKE_BUYs.Add("B", "Buy")
        ITEM_COST_MAKE_BUYs.Add("?", "MB Unknown")

        For Each ITEM_SNU_CODE As String In ITEM_SNU_CODEs.Keys
            For Each ITEM_BASIC_PROMO As String In ITEM_BASIC_PROMOs.Keys
                For Each ITEM_COST_MAKE_BUY As String In ITEM_COST_MAKE_BUYs.Keys
                    For Each rowICTCATG1 As DataRow In dst.Tables("ICTCATG1").Select("")
                        Dim ITEM_CATGY_CODE As String = rowICTCATG1.Item("ITEM_CATGY_CODE")
                        Dim ITEM_CATGY_DESC As String = rowICTCATG1.Item("ITEM_CATGY_DESC")
                        Dim ABC_GROUP As String = ITEM_SNU_CODE & ITEM_BASIC_PROMO & ITEM_COST_MAKE_BUY & ITEM_CATGY_CODE
                        Dim ABC_GROUP_DESC As String = _
                            ITEM_SNU_CODEs(ITEM_SNU_CODE) & ", " & _
                            ITEM_BASIC_PROMOs(ITEM_BASIC_PROMO) & ", " & _
                            ITEM_COST_MAKE_BUYs(ITEM_COST_MAKE_BUY) & ", " & _
                            ITEM_CATGY_DESC
                        dst.Tables("DPTABCPG").Rows.Add(New String() {ABC_GROUP, ABC_GROUP_DESC})
                    Next
                Next
            Next
        Next


        ASCDATA1.ExecuteSQL("Update " & ARTDEDA1 & " Set ITEM_ABC_CODE_FUT = NULL")

        Dim NYP As String = ASCMAIN1.Period_Calc(ASCMAIN1.CYP, Val(Absx1.numFor("FPDMOS").Value & "") - 1)

        Select Case Absx1.optFor("OPTDEMAND").Value
            Case "F"
                ASCMAIN1.sql = "Select DPTITMF1.ITEM_CODE" & vbCrLf _
                    & ", Sum (NVL(DPTITMF1.FORECAST,0)) FORECAST from DPTITMF1" _
                    & " where DPTITMF1.OPS_YYYYPP = '" & ASCMAIN1.CYP & "'" & vbCrLf _
                    & "   and DPTITMF1.OPS_YYYYPP_FC <= '" & NYP & "'" & vbCrLf _
                    & " having Sum (NVL(DPTITMF1.FORECAST,0)) <> 0" & vbCrLf _
                    & " group by DPTITMF1.ITEM_CODE"

            Case Else
                Dim sql_filter_history As String = ""
                If Absx1.optFor("OPTDEMAND").Value = "R" Then
                    sql_filter_history &= " where SOTINVH2.ORDR_YYYYPP_UPDATED < '" & ASCMAIN1.CYP & "'"
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                Else
                    sql_filter_history &= " where SOTINVH2.ORDR_YYYYPP_UPDATED >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
                    sql_filter_history &= "   and SOTINVH2.ORDR_YYYYPP_UPDATED < '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12 + Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                End If

                ASCMAIN1.sql = "Select SOTINVH2.ITEM_CODE" & vbCrLf _
                    & ", Sum (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) FORECAST from SOTINVH2" _
                    & sql_filter_history & vbCrLf _
                    & " having Sum (NVL(SOTINVH2.ORDR_QTY_SHIP,0)) <> 0" & vbCrLf _
                    & " group by SOTINVH2.ITEM_CODE"
        End Select

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & ASCMAIN1.sql & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ARTDEDA1 & " Set FORECAST = R1.FORECAST where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

        Dim rowGLTPARM2 As DataRow = LookUp("GLTPARM2", NYP)
        Dim PRD_END_DATE As String = Format(rowGLTPARM2.Item("PRD_END_DATE"), "dd-MMM-yyyy")

        Select Case Absx1.optFor("OPTDEMAND").Value
            Case "F"


            Case Else
                Dim sql_filter_history As String = ""
                If Absx1.optFor("OPTDEMAND").Value = "R" Then
                    sql_filter_history &= " where ICTSTAT1.OPS_YYYYPP < '" & ASCMAIN1.CYP & "'"
                    sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -1 * Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                Else
                    sql_filter_history &= " where ICTSTAT1.OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12) & "'"
                    sql_filter_history &= "   and ICTSTAT1.OPS_YYYYPP < '" & ASCMAIN1.Period_Calc(ASCMAIN1.CYP, -12 + Val(Absx1.numFor("FPDMOS").Value & "")) & "'"
                End If

                ASCMAIN1.sql = "Select ICTSTAT1.ITEM_CODE" & vbCrLf _
                    & ", Sum (NVL(ICTSTAT1.WHSE_QTY_CON,0)) PROD_COM from ICTSTAT1" _
                    & sql_filter_history & vbCrLf _
                    & " having Sum (NVL(ICTSTAT1.WHSE_QTY_CON,0)) <> 0" & vbCrLf _
                    & " group by ICTSTAT1.ITEM_CODE"
        End Select

        ASCMAIN1.sql = "" _
            & "Begin" & vbCrLf _
            & " Declare Cursor C1 is " & ASCMAIN1.sql & ";" & vbCrLf _
            & " Begin" & vbCrLf _
            & "  For R1 in C1 Loop" & vbCrLf _
            & "   Update " & ARTDEDA1 & " Set PROD_COM = R1.PROD_COM where ITEM_CODE = R1.ITEM_CODE;" & vbCrLf _
            & "  End Loop;" & vbCrLf _
            & " End;" & vbCrLf _
            & "End;"
        ASCDATA1.ExecuteSQL()

    End Sub
End Class