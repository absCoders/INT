Public Class ICRLOTH1

    Private sqlData As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -60, 0, 0)
        Set_cmbYP("RYP1", ASCMAIN1.CYP, -60, 0, 0)
    End Sub

    Protected Overrides Sub Build_Workfile()
        Prepare_dst(True)

        EnforceConstraints(False)

        ASCMAIN1.Progress("Run-Time Options", "")

        ' Get Run-Time options
        RYPLEGEND0 = MyBase.Absx1.cmbFor("RYP").Text
        RYP0 = Mid$(RYPLEGEND0, 1, 4) & Mid$(RYPLEGEND0, 6, 2)
        RYPLEGEND1 = MyBase.Absx1.cmbFor("RYP1").Text
        RYP1 = Mid$(RYPLEGEND1, 1, 4) & Mid$(RYPLEGEND1, 6, 2)

        ASCMAIN1.Progress("Preparing Result Set on Server", "")

        sql = "Select ICTIREC2.RECEIPT_NO, ICTIREC2.RECEIPT_LNO, ICTIREC1.RECEIPT_DATE, ICTIREC1.VEND_CODE, ICTIREC1.PO_ORDER_NO "
        sql &= ", ICTIREC2.OPS_YYYYPP, ICTIREC2.ITEM_CODE, ICTIREC2.ITEM_UOM, ICTIREC2.ITEM_COST_STD, ICTIREC2.TRAN_PV"
        sql &= ", ICTIREC2.TRAN_MV, ICTIREC2.TRAN_CV, ICTIREC2.PO_COST, ICTIREC2.QTY_REC"
        sql &= ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE"
        sql &= ", ICTCOSTA.ITEM_COST_VCOST, ICTCOSTA.ITEM_COST_MATLS "
        sql &= "  from ICTIREC1, ICTIREC2, ICTITEM1, ICTCOSTA, ICTCOLL1"
        sql &= "  where ICTIREC1.RECEIPT_NO = ICTIREC2.RECEIPT_NO"
        sql &= "  and ICTIREC2.ITEM_CODE = ICTITEM1.ITEM_CODE"
        sql &= "  and ICTITEM1.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE"
        sql &= "  and ICTCOSTA.OPS_YYYYPP = ICTIREC2.OPS_YYYYPP"
        sql &= "  and ICTCOSTA.ITEM_CODE = ICTIREC2.ITEM_CODE"
        sql &= "  and ICTIREC2.OPS_YYYYPP BETWEEN '" & RYP0 & "' and '" & RYP1 & "'"
        sql &= "  and ICTIREC1.REVERSED_BY_RECEIPT_NO IS NULL and ICTIREC1.REVERSES_RECEIPT_NO IS NULL"

        sql &= SQL_in("ITEM_CODE", "ICTIREC2.ITEM_CODE")
        sql &= SQL_in("COLLECTION_CODE", "ICTITEM1.COLLECTION_CODE")
        sql &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")
        dst.Tables.Add(ASCDATA1.GetDataTable(sql, "ICTIRECX", 2))

        dst.Tables.Add(ASCDATA1.GetDataTable("SELECT * FROM ICTBRAN1", "ICTBRAN1", 1))
        dst.Tables.Add(ASCDATA1.GetDataTable("SELECT * FROM ICTCOLL1", "ICTCOLL1", 1))

        EnforceConstraints(True)

    End Sub

    Public Overrides Sub Print_Report()

        SUBT = String.Empty

        If RYP0 = RYP1 Then
            SUBT = "Showing Cost Lot History for " & RYPLEGEND0
        Else
            SUBT = "Showing Cost Lot History for " & RYPLEGEND0 & " thru " & RYPLEGEND1
        End If

        Generate_Report(RPT, "", SUBT)

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


        EnforceConstraints(True)

    End Sub

End Class

