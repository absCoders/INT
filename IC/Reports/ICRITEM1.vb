Public Class ICRITEM1

    Private ICTITEM1 As String = String.Empty
    Private DPTITMF1 As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("GLTPARM1")
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 12, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()
        Prepare_dst(True)
    End Sub

    Public Overrides Sub Print_Report()

        CR_params.Add("NOTES", IIf(chkNotes.Checked, "1", "0"))
        CR_params.Add("DESC2", IIf(ChkDesc2.Checked, "1", "0"))

        Generate_Report("ICRLIST1", String.Empty, String.Empty)
    End Sub

    Overrides Function Prepare_dst( _
          ByVal perform_fill As Boolean, _
          ByVal ParamArray parms() As Object) As ASCBASE1

        If Not Me.Visible Then
            Clear_dst()
        End If

        Dim sql As String = String.Empty

        With dst

        End With

        If perform_fill Then
            Fill_Records_RPT()
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)

        Dim sql As String = String.Empty

        EnforceConstraints(False)

        ASCMAIN1.Progress("Gethering Items", "")
        If optAF.Value = "A" Then
            ' All Shipments from supplied period UNION All Forecasted Items
            sql = "SELECT DISTINCT ITEM_CODE FROM SOTINVH2 WHERE ORDR_YYYYPP_UPDATED >= '" & RYP & "' UNION "
        End If
        sql &= " SELECT ITEM_CODE FROM DPTITMF1 WHERE OPS_YYYYPP = '" & ASCMAIN1.CYP & "'"

        If DPTITMF1.Length = 0 Then
            DPTITMF1 = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("CREATE INDEX I_" & DPTITMF1 & "_1 ON " & DPTITMF1 & " (ITEM_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & DPTITMF1)
            ASCDATA1.ExecuteSQL("Insert Into " & DPTITMF1 & " " & sql)
        End If

        ASCMAIN1.Progress("Create Query", "")
        MyBase.Get_SQL("*")

        sql = " SELECT ICTCOLL1.BRAND_CODE, ICTITEM1.*, 'B' ITEM_BASIC_PROMO, NVL(ICTCOSTC.ITEM_COST_VCOST, 0) VC_COST,"
        sql &= " ' ' FORECASTED"
        sql &= " from ICTITEM1, ICTCOSTC, ICTCOLL1, " & DPTITMF1
        sql &= " WHERE " & DPTITMF1 & ".ITEM_CODE (+) = ICTITEM1.ITEM_CODE "
        sql &= " AND ICTCOSTC.ITEM_CODE (+) =  ICTITEM1.ITEM_CODE"
        sql &= " AND ICTITEM1.COLLECTION_CODE (+) =  ICTCOLL1.COLLECTION_CODE"

        If chkActive.Checked Then
            sql = sql & " AND ICTITEM1.ITEM_STATUS = 'A'"
        End If

        If chkOrel.Checked Then
            sql = sql & " AND NVL(ICTITEM1.ITEM_ORDR_REL_CODE, '') <> ''"
        End If
        ' one of these for each filter in Report Maintenance
        'sql &= SQL_in("ITEM_BASIC_PROMO", "ICTITEM.ITEM_BASIC_PROMO1")
        sql &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sql &= SQL_in("ITEM_CATGY_CODE", "ICTITEM1.ITEM_CATGY_CODE")
        sql &= SQL_in("ITEM_CLASS_CODE", "ICTITEM1.ITEM_CLASS_CODE")
        sql &= SQL_in("ITEM_CODE", "ICTITEM1.ITEM_CODE")
        sql &= SQL_in("ITEM_COST_MAKE_BUY", "ICTITEM1.ITEM_COST_MAKE_BUY")
        'sql &= SQL_in("ITEM_GROUP_CODE", "ICTITEM1.ITEM_GROUP_CODE")
        sql &= SQL_in("ITEM_SNU_CODE", "ICTITEM1.ITEM_SNU_CODE")
        sql &= SQL_in("ITEM_TYPE_CODE", "ICTITEM1.ITEM_TYPE_CODE")
        sql &= SQL_in("PROD_CODE", "ICTITEM1.PROD_CODE")
        sql &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")

        ASCMAIN1.Progress("Create Work Table", "")
        ICTITEM1 = ASCMAIN1.Temp_Table(sql)

        ASCMAIN1.Progress("Mark Non Forecasted", "")
        sql = "Update " & ICTITEM1 & " Set FORECASTED = '*' "
        sql = sql & " Where ITEM_CODE NOT IN (SELECT DISTINCT ITEM_CODE FROM DPTITMF1 WHERE OPS_YYYYPP = '" & ASCMAIN1.CYP & "')"
        ASCDATA1.ExecuteSQL(sql)

        ASCMAIN1.Progress("Update PG Item Code", "")
        sql = " BEGIN DECLARE CURSOR C1 IS SELECT * FROM SDTITEM1 WHERE RECORD_STATUS = 'A' AND ITEM_CODE IS NOT NULL;"
        sql &= "  BEGIN FOR R1 IN C1 LOOP"
        sql &= "    UPDATE " & ICTITEM1 & " SET PG_ITEM_CODE = LTRIM(R1.PG_ITEM_CODE, '0')"
        sql &= "        WHERE ITEM_CODE = R1.ITEM_CODE;"
        sql &= "  END LOOP;"
        sql &= "  END; END;"
        ASCDATA1.ExecuteSQL(sql)

        sql = "Select " & sql_SELECT_cols & ", ICTITEM1.ITEM_CODE"
        sql &= " from " & ICTITEM1 & " ICTITEM1" & sql_TABLE_NAMEs
        sql &= ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN)
        sql &= " group by " & sql_GROUP_BY_cols & ", ICTITEM1.ITEM_CODE"
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        With dst
            If Not .Tables.Contains("ICWITEM1") Then
                Create_TDA(.Tables.Add, "ICWITEM1", "Select * from " & ICTITEM1, 0, False, "", 0)
            Else
                dst.Tables("ICWITEM1").Clear()
            End If
            Fill_Records("ICWITEM1", String.Empty, True, "Select * from " & ICTITEM1)
        End With

        EnforceConstraints(True)

    End Sub

    Private Sub chkActive_CheckedChanged(sender As Object, e As EventArgs) Handles chkActive.CheckedChanged

    End Sub
End Class