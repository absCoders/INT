Public Class SPRDCOMO
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -24, 0, 0)
    End Sub
    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        ' If RYP = ASCMAIN1.CYP Then
        'ASCMAIN1.sql = "Select SPTDCOMC.*, SPTDCOMC.AMT_COMM OPEN_DEMO_COMM from SPTDCOMC" _
        '    & " where SPTDCOMC.PYMT_NO is Null"

        Dim sqlB As String = "Select OPS_YYYYPP,CUST_CODE,COLLECTION_CODE,HC_CODE,BRAND_CODE," & vbCrLf _
            & "SUM (QTY_SOLD) QTY_SOLD," & vbCrLf _
            & "SUM (AMT_SOLD) AMT_SOLD," & vbCrLf _
            & "SUM (QTY_EOW) QTY_EOW," & vbCrLf _
            & "SUM (AMT_EOW) AMT_EOW," & vbCrLf _
            & "MIN (OPS_YYYYWW_MIN) OPS_YYYYWW_MIN," & vbCrLf _
            & "MAX (OPS_YYYYWW_MAX) OPS_YYYYWW_MAX," & vbCrLf _
            & "SUM (AMT_COMM) AMT_COMM," & vbCrLf _
            & "MIN (DEMO_COMM_PCT) DEMO_COMM_PCT," & vbCrLf _
            & "SUM (AMT_COMM_CLAIMED) AMT_COMM_CLAIMED," & vbCrLf _
            & "SUM (AMT_COMM_PAID) AMT_COMM_PAID" & vbCrLf _
            & " from SPTDCOMB where NVL(SPTDCOMB.AMT_COMM,0) - NVL(SPTDCOMB.AMT_COMM_PAID,0) <> 0 group by " & vbCrLf _
            & "OPS_YYYYPP,CUST_CODE,COLLECTION_CODE,HC_CODE,BRAND_CODE" & vbCrLf

        Dim sqlBw As String = " And SPTDCOMB.OPS_YYYYPP = SPTDCOMC.OPS_YYYYPP" & vbCrLf _
            & "   And SPTDCOMB.CUST_CODE = SPTDCOMC.CUST_CODE" & vbCrLf _
            & "   And SPTDCOMB.HC_CODE = SPTDCOMC.HC_CODE" & vbCrLf _
            & "   And SPTDCOMB.AMT_COMM <> 0"

        Dim sqlC2 As String = Replace(sqlB, ",COLLECTION_CODE", "")
        Dim sqlC2w As String = Replace(sqlBw, "SPTDCOMB.", "SPTDCOMC2.")

        ASCMAIN1.sql = "Select SPTDCOMC.ACC_CTL_NO, SPTDCOMB.*" & vbCrLf _
            & ", ROUND(NVL(SPTDCOMC.AMT_COMM,0) * ( NVL(SPTDCOMB.AMT_COMM,0) - NVL(SPTDCOMB.AMT_COMM_PAID,0) )  /  (NVL(SPTDCOMC2.AMT_COMM,0) - NVL(SPTDCOMC2.AMT_COMM_PAID,0)),2) OPEN_DEMO_COMM" & vbCrLf _
            & " from SPTDCOMC," & vbCrLf _
            & " (" & vbCrLf & sqlB & ") SPTDCOMB, (" & vbCrLf & sqlC2 & ") SPTDCOMC2 " _
            & " where SPTDCOMC.PYMT_NO Is Null" & vbCrLf _
            & sqlBw & sqlC2w

        If RYP <> ASCMAIN1.CYP Then
            ASCMAIN1.sql = Replace(ASCMAIN1.sql,
                                   "from SPTDCOMB where",
                                   "from SPTDCOMH SPTDCOMB where SPTDCOMB.OPS_YYYYPP_H = '" & RYP & "' and")
            ASCMAIN1.sql = Replace(ASCMAIN1.sql,
                       "from SPTDCOMC,",
                       "from (Select * from SPTDCOMI where OPS_YYYYPP_H = '" & RYP & "') SPTDCOMC,")
        End If

        ASCMAIN1.sql = "Select * from (" & ASCMAIN1.sql & ") SPTDCOMC " & ASCMAIN1.SQL_Add_WHERE(sql_WHERE)
        Dim SPTDCOMC As String = ASCMAIN1.Temp_Table
        ' ASCDATA1.ExecuteSQL("Alter Table " & SPTDCOMC & " Add Primary Key (ACC_CTL_NO)")
        ASCDATA1.ExecuteSQL("Alter Table " & SPTDCOMC & " Add Primary Key (ACC_CTL_NO, OPS_YYYYPP, CUST_CODE, COLLECTION_CODE)")


        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        MyBase.Get_SQL("*", SPTDCOMC)

        sql_Data = "" _
            & ", COUNT (*) RECORDS, SUM (AMT_SOLD) AMT_SOLD, SUM (OPEN_DEMO_COMM) OPEN_DEMO_COMM" & vbCrLf

        sql_Cols = "" _
            & ",RECORDS,AMT_SOLD,OPEN_DEMO_COMM"

        sql_filter = ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & COLUMN_NAMEs_appended & vbCrLf _
            & sql_Data _
            & " from " & SPTDCOMC & " SPTDCOMC" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols & vbCrLf _
            & IIf(sql_GROUP_BY_cols = "", "", ",") & Mid(COLUMN_NAMEs_appended, 2)

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        Create_TDA(dst.Tables.Add, "SPTDCOMC", "*", , False, , 0)
        With dst.Tables("SPTDCOMC").Columns
            .Add("OPEN_DEMO_COMM", GetType(System.Decimal))
        End With

        Create_TDA(dst.Tables.Add, "ARTCUST1", "Select CUST_CODE, CUST_NAME from ARTCUST1", 0)

        EnforceConstraints(False)

        sql = "Select * from " & SPTDCOMC & " SPTDCOMC"
        Fill_Records("SPTDCOMC", "", True, sql)
        Fill_Records("ARTCUST1")
        EnforceConstraints(True)
    End Sub

    Public Overrides Sub Print_Report()
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
        End If
    End Sub
End Class