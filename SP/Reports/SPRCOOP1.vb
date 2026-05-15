Public Class SPRCOOP1

    Private wkTable As String = String.Empty

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Work Tables
        Dim sql As String = "Select SPTCOOP1.AUTH_NO" & vbCrLf _
              & " from SPTCOOP1, SPTCOOP3, ICTCOLL1" & vbCrLf _
              & " where SPTCOOP1.AUTH_NO = SPTCOOP3.AUTH_NO" & vbCrLf _
              & "   and SPTCOOP3.COLLECTION_CODE = ICTCOLL1.COLLECTION_CODE" & vbCrLf _
              & "   and SPTCOOP1.STATUS_CODE = 'O'"

        sql &= SQL_in("VEHICLE_CODE", "SPTCOOP1.VEHICLE_CODE")
        sql &= SQL_in("CUST_CODE", "SPTCOOP1.CUST_CODE")
        sql &= SQL_in("COLLECTION_CODE", "SPTCOOP3.COLLECTION_CODE")
        sql &= SQL_in("BRAND_CODE", "ICTCOLL1.BRAND_CODE")
        sql &= SQL_in("SEASON_CODE", "SPTCOOP1.SEASON_CODE")
        sql &= SQL_in("EVENT_TYPE_CODE", "SPTCOOP1.EVENT_TYPE_CODE")

        wkTable = ASCMAIN1.Temp_Table(sql)

        Dim sqlw As String = String.Empty

        Prepare_dst(True, sqlw)

    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
    End Sub

    Public Overrides Function Prepare_dst(perform_fill As Boolean, ParamArray parms() As Object) As ABSolution.ASCBASE1

        If Not Me.Visible Then Clear_dst()

        Create_TDA(dst.Tables.Add, "SPTCOOP1", "*")
        With dst.Tables("SPTCOOP1").Columns
            .Add("TOTAL_AMT", GetType(System.Decimal), "ISNULL(QTY,0) * ISNULL(VEHICLE_CPM,0) / 1000 + ISNULL(OTHER_COST,0)")
            .Add("CANCEL_AMT", GetType(System.Decimal), "ISNULL(TOTAL_AMT,0) - ISNULL(PAID_AMT,0) - ISNULL(OPEN_AMT,0)")
        End With

        Create_TDA(dst.Tables.Add, "SPTCOOP3", "*")
        Create_Relation("SPTCOOP1", "SPTCOOP3", "AUTH_NO")
        With dst.Tables("SPTCOOP3").Columns
            .Add("TOTAL_AMT", GetType(System.Decimal), "PARENT.TOTAL_AMT")
            .Add("OPEN_AMT", GetType(System.Decimal), "PARENT.OPEN_AMT")
            .Add("PAID_AMT", GetType(System.Decimal), "PARENT.PAID_AMT")
            .Add("DIST_AMT_OPEN", GetType(System.Decimal), "IIF(TOTAL_AMT=0,0,OPEN_AMT * ISNULL(DIST_AMT,0) / TOTAL_AMT)")
            .Add("DIST_AMT_PAID", GetType(System.Decimal), "IIF(TOTAL_AMT=0,0,PAID_AMT * ISNULL(DIST_AMT,0) / TOTAL_AMT)")
        End With

        Create_TDA(dst.Tables.Add, "ARTCUST1", "Select CUST_CODE, CUST_NAME from ARTCUST1", 0)

        If perform_fill Then
            Fill_Records_RPT(parms)
        End If

        Return clsASCBASE1

    End Function

    Public Overrides Sub Fill_Records_RPT(ByVal ParamArray parms() As Object)
        EnforceConstraints(False)

        Dim SQL As String = "Select * FROM SPTCOOP1 WHERE AUTH_NO IN (SELECT AUTH_NO FROM " & wkTable & ")"

        Fill_Records("SPTCOOP1", String.Empty, True, SQL)
        Fill_Records("SPTCOOP3", String.Empty, True, SQL.Replace("SPTCOOP1", "SPTCOOP3"))

        Fill_Records("ARTCUST1")

        EnforceConstraints(True)
    End Sub

End Class