Public Class SPRACOMO
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Set_cmbYP("RYP", ASCMAIN1.CYP, -24, 0, 0)
    End Sub
    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare Working Variables

        ' Prepare Work Tables

        If RYP = ASCMAIN1.CYP Then
            ASCMAIN1.sql = "Select SPTACOMC.*, SPTACOMC.AMT_COMM OPEN_ASP_COMM from SPTACOMC" _
                & " where SPTACOMC.PYMT_NO is Null"
        Else

            Dim sqlSPTACOMC As String = "Select SPTACOMC.*, NVL(SPTACOMC_PRE.OPS_YYYYPP_PAID, SPTACOMC.OPS_YYYYPP) OPS_YYYYPP_OPEN" _
                & " from SPTACOMC, SPTACOMC SPTACOMC_PRE where SPTACOMC_PRE.ACC_CTL_NO (+) = SPTACOMC.ACC_CTL_NO_ORIG"
            ASCMAIN1.sql = "Select SPTACOMC.*, SPTACOMC.AMT_COMM OPEN_ASP_COMM from (" & sqlSPTACOMC & ") SPTACOMC" _
                & " where  OPS_YYYYPP_OPEN <= '" & RYP & "' AND (OPS_YYYYPP_PAID IS NULL OR OPS_YYYYPP_PAID > '" & RYP & "')"

            'ASCMAIN1.sql = "Select SPTACOMC.*, GLTCREC3.CREC_AMT OPEN_ASP_COMM from SPTACOMC,GLTCREC3" _
            '    & " where GLTCREC3.OPS_YYYYPP = '" & RYP & "'" & vbCrLf _
            '    & "   and GLTCREC3.CREC_TYPE_CODE = 'AC'" & vbCrLf _
            '    & "   and SPTACOMC.ACC_CTL_NO = GLTCREC3.DETL_CTL_NO"
        End If


        Dim SPTACOMC As String = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SPTACOMC & " Add Primary Key (ACC_CTL_NO)")


        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        Dim sql_Data As String = ""
        Dim sql_Cols As String = ""

        MyBase.Get_SQL("*")

        sql_Data = "" _
            & ", COUNT (*) RECORDS, SUM (AMT_SOLD) AMT_SOLD, SUM (OPEN_ASP_COMM) OPEN_ASP_COMM" & vbCrLf

        sql_Cols = "" _
            & ",RECORDS,AMT_SOLD,OPEN_ASP_COMM"

        sql_filter = ""

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & COLUMN_NAMEs_appended & vbCrLf _
            & sql_Data _
            & " from " & SPTACOMC & " SPTACOMC" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf _
            & " group by " & sql_GROUP_BY_cols & vbCrLf _
            & IIf(sql_GROUP_BY_cols = "", "", ",") & Mid(COLUMN_NAMEs_appended, 2)

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()


        Create_TDA(dst.Tables.Add, "SPTACOMC", "*")
        With dst.Tables("SPTACOMC").Columns
            .Add("OPEN_ASP_COMM", GetType(System.Decimal))
        End With

        Create_TDA(dst.Tables.Add, "ARTCUST1", "Select CUST_CODE, CUST_NAME from ARTCUST1", 0)
        Create_TDA(dst.Tables.Add, "SPTACOM0", "Select * from SPTACOM0", 0)

        EnforceConstraints(False)

        sql = "Select * from " & SPTACOMC & " SPTACOMC"
        Fill_Records("SPTACOMC", "", True, sql)
        Fill_Records("ARTCUST1")
        Fill_Records("SPTACOM0")
        EnforceConstraints(True)
    End Sub

    Public Overrides Sub Print_Report()

        SUBT = "Open Accruals as of " & RYPLEGEND
        CR_params.Add("CHKSUMMARY", IIf(Absx1.chkFor("CHKSUMMARY").Checked, "1", "0"))
        Generate_Report(RPT, , SUBT)
    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        If eItemKey = "Proceed" Then
        End If
    End Sub
End Class