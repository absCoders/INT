Public Class RSCMAIN1

    Public Shared Sub Update_RSTRETLx( _
    ByVal EDI_DOC_SEQ_NO As String, _
    Optional ByVal plus_or_minus As String = "")

        Dim S As Int32 = 1
        If plus_or_minus = "-" Then
            S = -1
        End If

        ASCMAIN1.Record_Event("EDT852T1", EDI_DOC_SEQ_NO, "", Now, ASCMAIN1.USER_ID, "852" & plus_or_minus, "852 Update", "")

        ASCDATA1.ExecuteSP("RSPRETLX", "VN", New Object() {EDI_DOC_SEQ_NO, S}, New String() {"EDI_DOC_SEQ_NO_IN", "S"})
    End Sub

    Public Shared Function Get_ICTITEM1_Hist_CATGY( _
    ByVal RYP As String, _
    ByVal UseHistoricalCategory As Boolean, _
    Optional ByVal ICTITEM1 As String = "") As String

        If ICTITEM1 = "" Then
            ICTITEM1 = ASCMAIN1.Temp_Table("Select * from ICTITEM1")
            ASCDATA1.ExecuteSQL("Alter Table " & ICTITEM1 & " Add Primary Key (ITEM_CODE)")
            ASCDATA1.ExecuteSQL("Create Index I_" & ICTITEM1 & "_1 on " & ICTITEM1 & " (ITEM_CATGY_CODE)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1)
            ASCDATA1.ExecuteSQL("Insert into " & ICTITEM1 & " Select * from ICTITEM1")
        End If

        If RYP <> ASCMAIN1.CYP Then
            ASCMAIN1.sql = "Begin Declare Cursor C1 is " _
            & " Select ICTRETLA.* from ICTRETLA,ICTITEM1 where ICTRETLA.OPS_YYYYPP = '" & RYP & "'" _
            & "  and ICTITEM1.ITEM_CODE = ICTRETLA.ITEM_CODE " _
            & "  and (NVL(ICTRETLA.ITEM_RETAIL_PRICE,0) <> NVL(ICTITEM1.ITEM_RETAIL_PRICE,0)" _
            & "   or  NVL(ICTRETLA.ITEM_PRICE,0) <> NVL(ICTITEM1.ITEM_PRICE,0));" _
            & " Begin For R1 in C1 Loop " _
            & "  Update " & ICTITEM1 & " Set ITEM_RETAIL_PRICE = R1.ITEM_RETAIL_PRICE, " _
            & "    ITEM_PRICE = R1.ITEM_PRICE where ITEM_CODE = R1.ITEM_CODE; " _
            & " End Loop; End; End;"
            ASCDATA1.ExecuteSQL()
        End If

        If UseHistoricalCategory Then

            ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS" _
            & " SELECT DPTPROJ0.* FROM DPTPROJ0,GLTPARM2" _
            & " WHERE GLTPARM2.OPS_YYYYPP = '" & RYP & "'" _
            & " AND DPTPROJ0.OPS_YYYY = (CASE WHEN SUBSTR(GLTPARM2.OPS_YYYYPP,5,2)" _
            & " BETWEEN '01' AND '06' THEN TRIM(TO_CHAR(TO_NUMBER(SUBSTR(GLTPARM2.OPS_YYYYPP,1,4))-1))" _
            & " ELSE SUBSTR(GLTPARM2.OPS_YYYYPP,1,4) END)" _
            & " AND DPTPROJ0.SEASON = (CASE WHEN SUBSTR(GLTPARM2.OPS_YYYYPP,5,2)" _
            & " BETWEEN '01' AND '06' THEN 'F' ELSE 'S' END);" _
            & " BEGIN" _
            & " UPDATE " & ICTITEM1 & " SET HIDE_FROM_3PL = ITEM_CATGY_CODE;" _
            & " UPDATE " & ICTITEM1 & " SET ITEM_CATGY_CODE = 'I';" _
            & " FOR R1 IN C1 LOOP" _
            & " UPDATE " & ICTITEM1 & " SET ITEM_CATGY_CODE = R1.ITEM_CATGY_CODE" _
            & " WHERE ITEM_CODE = R1.ITEM_CODE;" _
            & " END LOOP;" _
            & " UPDATE " & ICTITEM1 & " SET ITEM_STATUS = DECODE(ITEM_CATGY_CODE,'I','I','F','I','A');" _
            & " END; END;"
            ASCDATA1.ExecuteSQL()

            ASCDATA1.ExecuteSQL("Update " & ICTITEM1 & " Set ITEM_CATGY_CODE = 'I' where ITEM_CATGY_CODE is Null")

            ASCMAIN1.sql = "BEGIN DECLARE CURSOR C1 IS" _
            & " SELECT ICTHIST1.* FROM ICTHIST1,GLTPARM2" _
            & " WHERE GLTPARM2.OPS_YYYYPP = '" & RYP & "'" _
            & " AND ICTHIST1.OPS_YYYY = (CASE WHEN SUBSTR(GLTPARM2.OPS_YYYYPP,5,2)" _
            & " BETWEEN '01' AND '06' THEN TRIM(TO_CHAR(TO_NUMBER(SUBSTR(GLTPARM2.OPS_YYYYPP,1,4))-1))" _
            & " ELSE SUBSTR(GLTPARM2.OPS_YYYYPP,1,4) END)" _
            & " AND ICTHIST1.SEASON = (CASE WHEN SUBSTR(GLTPARM2.OPS_YYYYPP,5,2)" _
            & " BETWEEN '01' AND '06' THEN 'F' ELSE 'S' END);" _
            & " BEGIN" _
            & " FOR R1 IN C1 LOOP" _
            & " UPDATE " & ICTITEM1 & " SET CUST_CODE = R1.CUST_CODE" _
            & " WHERE ITEM_CODE = R1.ITEM_CODE;" _
            & " END LOOP;" _
            & " END; END;"
            ASCDATA1.ExecuteSQL()

        End If

        Return ICTITEM1

    End Function

    Public Shared Sub Load_Filter( _
    ByVal COLUMN_NAMEs() As String, _
    ByVal grdASTFLTR1 As UltraWinGrid.UltraGrid, _
    ByVal frmASFBASE0 As ASFBASE0)

        grdASTFLTR1.DataSource = frmASFBASE0.dst.Tables("ASTFLTR1")

        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Dim rowASTDSQLK As DataRow = frmASFBASE0.LookUp("ASTDSQLK", COLUMN_NAME)
            Dim COLUMN_CAPTION As String = ""

            If rowASTDSQLK Is Nothing Then
                COLUMN_CAPTION = ASCMAIN1.Make_Caption(COLUMN_NAME)
            Else
                COLUMN_CAPTION = rowASTDSQLK.Item("COLUMN_CAPTION") & ""
            End If
            frmASFBASE0.dst.Tables("ASTFLTR1").Rows.Add(New String() {COLUMN_NAME, COLUMN_CAPTION})
        Next
    End Sub

    Public Shared Function RSTBUDR1_as_YP() As String

        'ASCMAIN1.sql = ""
        'For I As Int16 = 1 To 12
        '    ASCMAIN1.sql &= " union Select COLLECTION_CODE,ITEM_CATGY_CODE,CUST_CODE,CUST_STORE_NO,OPS_YYYY || '" _
        '    & Format(I, "00") & "' OPS_YYYYPP, BUDGET_P" & Format(I, "00") & " BUDGET from RSTBUDR1 where BUDGET_P" & Format(I, "00") & " <> 0"
        'Next
        'ASCMAIN1.sql = Mid(ASCMAIN1.sql, 8)

        ' LATER, THIS ROUTINE WILL BE AS FOLLOWS, AND MANY FUNCTIONS WON'T EVEN NEED TO CALL THIS ROUTINE
        ASCMAIN1.sql = "Select * from RSTBUDR1"

        Dim RSTBUDR1 As String = ASCMAIN1.Temp_Table

        ASCDATA1.ExecuteSQL("Alter Table " & RSTBUDR1 & " Add Primary Key (COLLECTION_CODE, ITEM_CATGY_CODE, CUST_CODE, CUST_STORE_NO, OPS_YYYYPP)")

        'If 1 - ASCMAIN1.PCO <> 0 Then
        '    ASCMAIN1.sql = "Update " & RSTBUDR1 & " Set OPS_YYYYPP = PERIOD_CALC(OPS_YYYYPP," & CStr(1 - ASCMAIN1.PCO) & ")"
        '    ASCDATA1.ExecuteSQL()
        'End If
 
        Return RSTBUDR1

    End Function
End Class
