Public Class SORSLSX1

#Region "General Declarations"

    Dim SOTSLSX1 As String

#End Region

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("ARTPARM1")
    End Sub

    Protected Overrides Sub Build_Workfile()

        Dim BASE As String = XNO
        Dim FILENAME As String = ""

        ASCMAIN1.sql = "Select SOTINVH1.CUST_CODE, SOTINVH1.INV_NO, SOTINVH1.INV_DATE, SOTINVH1.CUST_STORE_NO, " _
            & " SOTINVH2.ITEM_CODE, SOTINVH2.ORDR_QTY_SHIP, SOTINVH2.ORDR_UNIT_PRICE" _
            & " FROM SOTINVH1,SOTINVH2" _
            & " WHERE SOTINVH2.INV_TYPE = SOTINVH1.INV_TYPE" _
            & " AND SOTINVH2.INV_NO = SOTINVH1.INV_NO" _
            & " AND SOTINVH1.ORDR_YYYYPP_UPDATED = '201305' AND SOTINVH1.INV_DATE <= '03-MAY-2013'"

        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTSLSX1", 0))


        FILENAME = ASCMAIN1.Folders("Temp") & "SLS" & BASE
        Using SW As New System.IO.StreamWriter(FILENAME)
            For Each row As DataRow In dst.Tables("SOTSLSX1").Select("CUST_CODE,INV_NO,INV_DATE,CUST_STORE_NO,ITEM_CODE")
                Dim LINE As String = ""
                LINE &= row.Item("zusa") & vbTab
                LINE &= row.Item("INV_NO") & vbTab
                LINE &= Format(row.Item("INV_DATE"), "yyyyMMdd") & vbTab
                LINE &= row.Item("CUST_CODE") & vbTab
                LINE &= row.Item("CUST_STORE_NO") & vbTab
                LINE &= row.Item("USD") & vbTab
                LINE &= row.Item("ITEM_CODE") & vbTab
                LINE &= row.Item("ORDR_QTY_SHIP") & vbTab
                LINE &= row.Item("ORDR_UNIT_PRICE") & vbTab
                LINE &= Val(row.Item("ORDR_QTY_SHIP") & "") * Val(row.Item("ORDR_UNIT_PRICE") & "")
                SW.WriteLine(LINE)
            Next
        End Using



        ASCMAIN1.sql = "Select ARTCUST1.CUST_CODE, ARTCUST1.CUST_NAME, ARTCUST2.CUST_STORE_NO, ARTCUST1.SREP_CODE, ARTCUST2.SREP_CODE, ARTCUST1.TRADE_CLASS_CODE" _
            & " FROM ARTCUST1,ARTCUST2" _
            & " WHERE ARTCUST1.CUST_CODE = ARTCUST2.CUST_CODE"

        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "ARTCUSTX", 0))



        Check_if_Empty("SOTSLSX1")
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then

        End If
    End Sub
  
End Class