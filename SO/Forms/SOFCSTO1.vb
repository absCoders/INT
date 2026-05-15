Imports Infragistics.Win.UltraWinGrid
Imports SpreadsheetGear.Commands
Imports SpreadsheetGear.Windows.Forms

Public Class SOFCSTO1

#Region "Declarations"
    Dim CUST_CODE As String
    Dim SELL_CODE As String
    Dim CUST_STORE_NO As String = ""

    Dim CSO_NO As String
    Dim rowSOTCSTO1 As DataRow

    Dim rowSOTSELL1 As DataRow
    Dim rowARTCUST1 As DataRow
    Dim rowICTITEM1 As DataRow
    Dim CSO_LNOs As New List(Of Int64)   ' list of CSO_LNOs that are deleted

    Dim sqlICTITEM1 As String
    Dim ICTITEM1 As String
    Dim sqlSOTALLOX As String
    Dim SOTALLOX As String
    Dim sqlSOTALLOZ As String
    Dim SOTALLOZ As String

    Dim SOTCSTOX As String
    Dim SOTCSTOI As String
    Dim DATE_START_since As Date
    Dim DATE_START_until As Date

    Dim MAX_COLs As Integer = 350
    Dim MAX_RSCs As Integer = 500
    Dim UPDATE_DISABLED As Boolean = False

    Dim CUST_ADDR_cols As String()
    Dim ORDR_GROUP_NOs As New List(Of String)

    Dim STO2_COLS_changed As New List(Of Integer)
    Dim STO3_COLS_changed As New List(Of Integer)

    Dim workbook As SpreadsheetGear.IWorkbook = Nothing
    Dim worksheet As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing

    Dim rangeCopyFrom As SpreadsheetGear.IRange
    Dim rangePaste_To As SpreadsheetGear.IRange

    Dim XLS_NO As String
    Dim XLS_PWD As String = "ABS"

    Dim COL_CSO_ADDR_LNO As Integer = 0
    Dim ROW_ITEM_CODE As Integer = 0

    Dim ALLO_GROUP_CODEs As New List(Of String)
    Dim HC_CODEs As New List(Of String)
    Dim RSC_Tags As New List(Of String)
    Dim PROD_CODES As New List(Of String)
    Dim READ_ONLY As New List(Of String)

    Dim WithEvents ws As SpreadsheetGear.IWorksheet
    Dim c0_Items As Integer = -1 ' 17
    Dim r0T As Integer = 12 ' ADDRESSES START ON r0T + 1

    Dim SELL_CODE_this_user As String
    Dim REGION_CODE_this_user As String
    Dim isAC As Boolean = False
    Dim isClearing As Boolean = False
    Dim isClearing_R As Integer = -1
    Dim isClearing_C As Integer = -1
    Dim isPasting As Boolean = False
    Dim isPasting_R As Integer = -1
    Dim isPasting_C As Integer = -1

    Dim restore_in_process As Boolean = False
    Dim dst2 As DataSet
    Dim NO_ITEMS As Boolean
    Dim modifiedAddresses As New List(Of DataRow)
    Dim originalValues As New Dictionary(Of DataRow, DataRow)
    Private validatedAddresses As New List(Of DataRow)

    Dim X As SpreadsheetGear.Commands.CommandManager

#End Region

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Check_InquiryMode()
        Get_PARM("SOTPARM1")
        Get_PARM("ICTPARM1")
        Get_PARM("ASTPARM1")

        CUST_CODE = "IPLBAE"
        rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)

        Create_Work_Tables(True)

        AUDIT.Add("SOTORDR1", "*")
        AUDIT.Add("SOTORDR2", "*")


        If ASCMAIN1.USER_CODES.Contains("FS") Then
            SELL_CODE_this_user = "?"
            Dim rowTATUSER1 As DataRow = LookUp("TATUSER1", ASCMAIN1.USER_ID)
            If rowTATUSER1 IsNot Nothing Then
                SELL_CODE_this_user = rowTATUSER1.Item("SELL_CODE") & ""
                REGION_CODE_this_user = rowTATUSER1.Item("REGION_CODE") & ""
                If SELL_CODE_this_user = "" And REGION_CODE_this_user = "" Then
                    SELL_CODE_this_user = "?"
                End If

                Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE_this_user)
                If rowSOTSELL1 IsNot Nothing Then
                    If rowSOTSELL1.Item("SELL_TYPE") & "" = "AC" Then
                        isAC = True
                    End If
                End If
            End If
        End If


        With dst

            ASCMAIN1.sql = "Select SOTORDR1.*" _
                & " from SOTORDR1 where SOTORDR1.CUST_CODE =  :PARM1 and SELL_CODE = :PARM2 and SOTORDR1.ORDR_CUST_PO = :PARM3"
            Create_TDA(.Tables.Add, "SOTORDRX", "**", 0, False, "VVV")

            ASCMAIN1.sql = $"Select SOTCSTOX.*, SOTORDR0.ORDR_STATUS from {SOTCSTOX} SOTCSTOX,SOTORDR0" & vbCrLf _
                & " where SOTORDR0.ORDR_GROUP_NO (+) = SOTCSTOX.ORDR_GROUP_NO"
            Create_TDA(.Tables.Add, "SOTCSTOX", "**", 0, False, "V", 1)

            ASCMAIN1.sql = $"Select SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_STATUS, SOTORDR1.CUST_STORE_LOCATION
            from SOTORDR1, {SOTCSTOX} SOTCSTOX where SOTORDR1.ORDR_GROUP_NO = SOTCSTOX.ORDR_GROUP_NO"
            ASCMAIN1.sql = $"SELECT X.*, SOTCSTO3.CSO_NO FROM
                            (Select SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_STATUS, SOTORDR1.CUST_STORE_LOCATION, SOTORDR1.WHSE_CODE
                            FROM SOTORDR1
                            WHERE ORDR_NO IN (SELECT ORDR_NO FROM SOTCSTO3 WHERE CSO_NO IN (Select SOTCSTOX.CSO_NO
                            from SOTORDR1,  {SOTCSTOX} SOTCSTOX where SOTORDR1.ORDR_GROUP_NO = SOTCSTOX.ORDR_GROUP_NO) AND ORDR_NO IS NOT NULL)) X, SOTCSTO3
                            WHERE X.ORDR_NO = SOTCSTO3.ORDR_NO (+)"
            ASCMAIN1.sql = $"Select SOTORDR1.ORDR_GROUP_NO, SOTORDR1.ORDR_NO, SOTORDR1.ORDR_DATE, SOTORDR1.ORDR_STATUS, SOTORDR1.CUST_STORE_LOCATION, SOTORDR1.WHSE_CODE, SOTCSTO3.CSO_NO
                            FROM SOTORDR1, SOTCSTO3
                            WHERE SOTORDR1.ORDR_NO = SOTCSTO3.ORDR_NO AND SOTCSTO3.CSO_NO IN (SELECT CSO_NO FROM {SOTCSTOX} SOTCSTOX)"
            Create_TDA(.Tables.Add, "SOTCSTOH", "**", 0, False)

            '            ASCMAIN1.sql = $"Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.ITEM_CODE, SOTORDR2.ORDR_STATUS
            ', SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, SOTORDR2.ORDR_QTY_CANC, SOTORDR2.ORDR_QTY_SHIP
            'from SOTORDR2,SOTORDR1, {SOTCSTOX} SOTCSTOX where SOTORDR2.ORDR_NO = SOTORDR1.ORDR_NO and SOTORDR1.ORDR_GROUP_NO = SOTCSTOX.ORDR_GROUP_NO"
            '            Create_TDA(.Tables.Add, "SOTCSTOD", "**", 0, False)

            ASCMAIN1.sql = $"Select SOTORDR2.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.ITEM_CODE, SOTORDR2.ORDR_STATUS
                            , SOTORDR2.ORDR_QTY, SOTORDR2.ORDR_QTY_OPEN, SOTORDR2.ORDR_QTY_PICK, SOTORDR2.ORDR_QTY_CANC, SOTORDR2.ORDR_QTY_SHIP
                            from SOTORDR2 WHERE ORDR_NO IN (SELECT * FROM TABLE(IN_LIST(:PARM1)))"
            Create_TDA(.Tables.Add, "SOTCSTOD", "**", 0, False, "V")

            Create_Relation("SOTCSTOX", "SOTCSTOH", "CSO_NO")
            Create_Relation("SOTCSTOH", "SOTCSTOD", "ORDR_NO")

            ASCMAIN1.sql = "SELECT " & vbCrLf &
            "S3.CSO_NO, s1.cso_date, S3.CSO_TYPE, S1.sell_code, " & vbCrLf &
            "S3.CUST_NAME AS CS_NAME, S3.CUST_ADDR1 AS CS_ADDR1, S3.CUST_ADDR2 AS CS_ADDR2, S3.CUST_ADDR3 AS CS_ADDR3, " & vbCrLf &
            "S3.CUST_CITY AS CS_CITY, S3.CUST_STATE AS CS_STATE, S3.CUST_ZIP_CODE AS CS_ZIP_CODE, " & vbCrLf &
            "S3.CUST_PHONE AS CS_PHONE, S3.CUST_EMAIL AS CS_EMAIL, " & vbCrLf &
            "SPTRSSP1.RSSP_NAME AS DB_NAME, SPTRSSP1.RSSP_SHIP_TO_ADDR1 AS DB_ADDR1, SPTRSSP1.RSSP_SHIP_TO_ADDR2 AS DB_ADDR2, SPTRSSP1.RSSP_SHIP_TO_ADDR3 AS DB_ADDR3, " & vbCrLf &
            "SPTRSSP1.RSSP_SHIP_TO_CITY AS DB_CITY, SPTRSSP1.RSSP_SHIP_TO_STATE AS DB_STATE, SPTRSSP1.RSSP_SHIP_TO_ZIP_CODE AS DB_ZIP_CODE, " & vbCrLf &
            "SPTRSSP1.RSSP_PHONE AS DB_PHONE, SPTRSSP1.RSSP_EMAIL AS DB_EMAIL " & vbCrLf &
            "FROM " & vbCrLf &
            "SOTCSTO3 S3 " & vbCrLf &
            "INNER JOIN " & vbCrLf &
            "SPTRSSP1 ON SPTRSSP1.RSSP_CODE = S3.CSO_KEY " & vbCrLf &
            "INNER JOIN " & vbCrLf &
            "SOTCSTO1 S1 ON S1.CSO_NO = S3.CSO_NO " & vbCrLf &
            "WHERE s1.cso_status <> 'D' AND (S3.CSO_TYPE IN ('RSC', 'SDS')) " & vbCrLf &
            "AND ( NVL(CUST_NAME,'?') <> NVL(RSSP_NAME,'?') " & vbCrLf &
            "OR NVL(CUST_ADDR1,'?') <> NVL(RSSP_SHIP_TO_ADDR1,'?') OR NVL(CUST_ADDR2,'?') <> NVL(RSSP_SHIP_TO_ADDR2,'?') " & vbCrLf &
            "OR NVL(CUST_ADDR3,'?') <> NVL(RSSP_SHIP_TO_ADDR3,'?') OR NVL(CUST_CITY,'?') <> NVL(RSSP_SHIP_TO_CITY,'?') " & vbCrLf &
            "OR NVL(CUST_STATE,'?') <> NVL(RSSP_SHIP_TO_STATE,'?') OR NVL(CUST_ZIP_CODE,'?') <> NVL(RSSP_SHIP_TO_ZIP_CODE,'?') " & vbCrLf &
            "OR NVL(CUST_PHONE,'?') <> NVL(RSSP_PHONE,'?') OR NVL(CUST_EMAIL,'?') <> NVL(RSSP_EMAIL,'?') " & vbCrLf &
            ") " & vbCrLf &
            "UNION " & vbCrLf &
            "SELECT " & vbCrLf &
            "S3.CSO_NO, s1.cso_date, S3.CSO_TYPE, S1.sell_code, " & vbCrLf &
            "S3.CUST_NAME AS CS_NAME, S3.CUST_ADDR1 AS CS_ADDR1, S3.CUST_ADDR2 AS CS_ADDR2, S3.CUST_ADDR3 AS CS_ADDR3, " & vbCrLf &
            "S3.CUST_CITY AS CS_CITY, S3.CUST_STATE AS CS_STATE, S3.CUST_ZIP_CODE AS CS_ZIP_CODE, " & vbCrLf &
            "S3.CUST_PHONE AS CS_PHONE, S3.CUST_EMAIL AS CS_EMAIL, " & vbCrLf &
            "SOTSELL1.SELL_NAME AS DB_NAME, SOTSELL1.SELL_ADDR1 AS DB_ADDR1, SOTSELL1.SELL_ADDR2 AS DB_ADDR2, SOTSELL1.SELL_ADDR3 AS DB_ADDR3, " & vbCrLf &
            "SOTSELL1.SELL_CITY AS DB_CITY, SOTSELL1.SELL_STATE AS DB_STATE, SOTSELL1.SELL_ZIP_CODE AS DB_ZIP_CODE, " & vbCrLf &
            "SOTSELL1.SELL_PHONE AS DB_PHONE, SOTSELL1.SELL_EMAIL AS DB_EMAIL " & vbCrLf &
            "FROM " & vbCrLf &
            "SOTCSTO3 S3 " & vbCrLf &
            "INNER JOIN " & vbCrLf &
            "SOTSELL1 ON SOTSELL1.SELL_CODE = S3.CSO_KEY " & vbCrLf &
            "INNER JOIN " & vbCrLf &
            "SOTCSTO1 S1 ON S1.CSO_NO = S3.CSO_NO WHERE s1.cso_status <> 'D' AND (S3.CSO_TYPE IN ('AE', 'AC')) " & vbCrLf &
            "AND ( NVL(CUST_NAME,'?') <> NVL(SELL_NAME,'?') " & vbCrLf &
            "OR NVL(CUST_ADDR1,'?') <> NVL(SELL_ADDR1,'?') OR NVL(CUST_ADDR2,'?') <> NVL(SELL_ADDR2,'?') " & vbCrLf &
            "OR NVL(CUST_ADDR3,'?') <> NVL(SELL_ADDR3,'?') OR NVL(CUST_CITY,'?') <> NVL(SELL_CITY,'?') " & vbCrLf &
            "OR NVL(CUST_STATE,'?') <> NVL(SELL_STATE,'?') OR NVL(CUST_ZIP_CODE,'?') <> NVL(SELL_ZIP_CODE,'?') " & vbCrLf &
            "OR NVL(CUST_PHONE,'?') <> NVL(SELL_PHONE,'?') OR NVL(CUST_EMAIL,'?') <> NVL(SELL_EMAIL,'?'))"

            ASCMAIN1.sql = $"Select * from ({ASCMAIN1.sql}) where (SELL_CODE = :PARM1 or :PARM1 = '*')"
            Create_TDA(.Tables.Add, "SOTCSTOA", "**", 0, False, "V")

            ASCMAIN1.sql = "Select * from " & SOTCSTOI
            Create_TDA(.Tables.Add, "SOTCSTOI", "**", 0, False)

            Create_TDA(.Tables.Add, "SOTCSTO1", "*", 1)

            Create_TDA(.Tables.Add, "ICTCOLL0", "*", 0, False)
            With .Tables("ICTCOLL0").Columns
                .Add("SEL")
            End With
            .Tables("ICTCOLL0").Columns("SEL").DefaultValue = "0"

            Create_TDA(.Tables.Add, "ICTPROD1", "*", 0, False)
            With .Tables("ICTPROD1").Columns
                .Add("SEL")
            End With
            .Tables("ICTPROD1").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = $"Select * from SOTCSTT1 where SELL_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTRSCT1", "**", 0, False, "V", 2)
            With .Tables("SOTRSCT1").Columns
                .Add("SEL")
            End With
            .Tables("SOTRSCT1").Columns("SEL").DefaultValue = "0"


            Create_TDA(.Tables.Add, "SOTALLOG", "*", 0, False)
            With .Tables("SOTALLOG").Columns
                .Add("SEL")
            End With
            .Tables("SOTALLOG").Columns("SEL").DefaultValue = "0"

            ASCMAIN1.sql = "Select SOTCSTO2.*, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE" & vbCrLf _
                & ", ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE, ICTITEM1.COLLECTION_CODE" & vbCrLf _
                & ", ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE, ICTPROD1.PROD_CODE, ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_SO_QTY_MIN" & vbCrLf _
                & " from SOTCSTO2,ICTITEM1,ICTCOLL1,ICTPROD1" & vbCrLf _
                & " where ICTITEM1.ITEM_CODE = SOTCSTO2.ITEM_CODE and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE AND ICTITEM1.PROD_CODE = ICTPROD1.PROD_CODE"
            Create_TDA(.Tables.Add, "SOTCSTO2", "**", 1)
            With .Tables("SOTCSTO2").Columns
                .Add("ALLO_GROUP_CODE")

                .Add("QTY_ALLO", GetType(System.Int64))
                .Add("ORDR_QTY", GetType(System.Int64))
                .Add("ORDR_QTY_OPEN", GetType(System.Int64))
                .Add("ORDR_QTY_PICK", GetType(System.Int64))
                .Add("ORDR_QTY_SHIP", GetType(System.Int64))
                .Add("ORDR_QTY_CANC", GetType(System.Int64))
                .Add("QTY_LEFT", GetType(System.Int64))


                .Add("WHSE_QTY_ON_HAND", GetType(System.Int64))
                .Add("WHSE_QTY_ONPO", GetType(System.Int64))
                .Add("WHSE_QTY_OPEN", GetType(System.Int64))
                .Add("WHSE_QTY_PICK", GetType(System.Int64))

                .Add("CSO_COL", GetType(System.Int64))

                Dim CT As String = ""
                For c As Integer = 1 To MAX_RSCs
                    Dim CX As String = $"CSO_QTY_{Format(c, "000")}"
                    CT &= $"+ISNULL({CX},0)"
                    .Add(CX, GetType(System.Int64))
                Next
                .Add("CSO_QTY_TOTAL", GetType(System.Int64), Mid(CT, 2))
                .Add("QTY_BAL", GetType(System.Int64), "ISNULL(QTY_LEFT,0) - ISNULL(CSO_QTY_TOTAL,0)")
            End With

            Create_TDA(.Tables.Add, "SOTCSTO3", "*", 1)
            With .Tables("SOTCSTO3").Columns
                .Add("CSO_RSC", GetType(System.Int64))

                Dim CT As String = ""
                For c As Integer = 1 To MAX_COLs
                    Dim CX As String = $"CSO_QTY_{Format(c, "000")}"
                    CT &= $"+ISNULL({CX},0)"
                    .Add(CX, GetType(System.Int64))
                Next
                .Add("CSO_QTY_TOTAL", GetType(System.Int64), Mid(CT, 2))
            End With

            Create_TDA(.Tables.Add, "SOTCSTO4", "*", 1)

            'ASCMAIN1.sql = "Select * from SOTORDXR where ORDR_NO = :PARM1"
            'Create_TDA(.Tables.Add, "SOTORDXR", "**", 0, True, "V")
            Create_TDA(.Tables.Add, "SOTORDXR", "*", 1)

            Create_TDA(.Tables.Add, "SOTORDR1", "*", 1)
            .Tables("SOTORDR1").Columns.Add("CSO_KEY", GetType(System.String))
            Create_TDA(.Tables.Add, "SOTORDR2", "*", 1)
            Create_TDA(.Tables.Add, "SOTORDR5", "*", 1)

            Create_TDA(.Tables.Add, "ICTWHSE1", "*")
            Create_TDA(.Tables.Add, "SOTSVIAS", "*", 0, False)
            Fill_Records("SOTSVIAS")

            ASCMAIN1.sql = $"Select * from {ICTITEM1}"
            Create_TDA(.Tables.Add, "ICTITEM1", "**", , False,, 1)

            ASCMAIN1.sql = $"Select SOTALLOX.*" & vbCrLf _
                & ", SOTALLOZ.ORDR_QTY" & vbCrLf _
                & ", SOTALLOZ.ORDR_QTY_OPEN" & vbCrLf _
                & ", SOTALLOZ.ORDR_QTY_PICK" & vbCrLf _
                & ", SOTALLOZ.ORDR_QTY_SHIP" & vbCrLf _
                & ", SOTALLOZ.ORDR_QTY_CANC" & vbCrLf _
                & ", SOTALLOZ.ORDR_SHIP_DATE_OPEN" & vbCrLf _
                & ", SOTALLOZ.ORDR_SHIP_DATE_PICK" & vbCrLf _
                & ", SOTALLOZ.ORDR_SHIP_DATE_SHIP" & vbCrLf _
                & $"from {SOTALLOX} SOTALLOX, {SOTALLOZ} SOTALLOZ" & vbCrLf _
                & " where SOTALLOZ.ALLO_CTL_NO (+) = SOTALLOX.ALLO_CTL_NO"
            Create_TDA(.Tables.Add, "SOTALLOX", "**", 0, False, "", 0)
            With .Tables("SOTALLOX").Columns
                '.Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)-ISNULL(ORDR_QTY_OPEN,0)")
                '.Add("QTY_BAL", GetType(System.Int64), "IIF(QTY_LEFT>=0,QTY_LEFT,0)")
                .Add("QTY_BAL", GetType(System.Int64), "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_SHIP,0)-ISNULL(ORDR_QTY_PICK,0)")
                .Add("QTY_LEFT", GetType(System.Int64), "ISNULL(QTY_BAL,0)-ISNULL(ORDR_QTY_OPEN,0)")
            End With

            ASCMAIN1.sql = $"Select * from {SOTALLOZ}"
            Create_TDA(.Tables.Add, "SOTALLOZ", "**", 0, False, "VDVV", 2)

            ASCMAIN1.sql = $"Select * from SOTCSTT1 where SELL_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTCSTT1", "**", 0, True, "V", 2)

            ASCMAIN1.sql = $"Select SOTCSTT2.*" & vbCrLf _
                & " from SOTCSTT2" & vbCrLf _
                & " where SOTCSTT2.SELL_CODE = :PARM1"
            Create_TDA(.Tables.Add, "SOTCSTT2", "**", 0, True, "V", 3)

            'ASCMAIN1.sql = "Select SPTRSSP1.* from SPTRSSP1 where RSSP_ID in (" & vbCrLf _
            '    & " Select Distinct SPTRSSP1.RSSP_ID" & vbCrLf _
            '    & " from SPTCWRX2,ARTCUST2,SPTRSSP1" & vbCrLf _
            '    & " where SPTCWRX2.WORK_DATE > SYSDATE - 365" & vbCrLf _
            '    & "   and ARTCUST2.CUST_CODE = SPTCWRX2.CUST_CODE" & vbCrLf _
            '    & "   and ARTCUST2.CUST_STORE_NO = SPTCWRX2.CUST_STORE_NO" & vbCrLf _
            '    & $"   and ARTCUST2.SELL_CODE_AC = :PARM1" & vbCrLf _
            '    & "   and SPTRSSP1.RSSP_ID = SPTCWRX2.SSN" & vbCrLf _
            '    & "   and NVL(SPTRSSP1.RSSP_STATUS, 'A') = 'A' and SPTRSSP1.RSSP_TYPE IN ('C', 'D') and SPTRSSP1.RSSP_DATE_TERM IS NULL)"

            'RSSP CODE, NAME, PHONE, EMAIL, SELL_CODE
            'BUS_MGR CODE, NAME, PHONE, EMAIL, SELL_CODE
            ASCMAIN1.sql = $"
            Select RSSP_CODE,RSSP_NAME, RSSP_PHONE, RSSP_EMAIL
            FROM SPTRSSP1
            WHERE SELL_CODE = :PARM1
            AND RSSP_STATUS = 'A'
            AND RSSP_TYPE IN ('C', 'D')
            AND RSSP_DATE_TERM IS NULL
            UNION
            SELECT BUS_MGR_CODE AS RSSP_CODE, BUS_MGR_NAME AS RSSP_NAME, BUS_MGR_PHONE AS RSSP_PHONE, BUS_MGR_EMAIL AS RSSP_EMAIL
            FROM 
            SATAEBM1
            LEFT JOIN SOTSELL1 ON SOTSELL1.SELL_CODE_MGR = SATAEBM1.SELL_CODE
            WHERE 
            SATAEBM1.BUS_MGR_STATUS = 'A' AND
            (SATAEBM1.SELL_CODE = :PARM1
            OR SATAEBM1.SELL_CODE IN (SELECT SELL_CODE FROM SOTSELL1 WHERE SELL_CODE_MGR = :PARM1))"
            'ASCMAIN1.sql = $"Select SPTRSSP1.RSSP_CODE, SPTRSSP1.RSSP_SHIP_TO_NAME, SPTRSSP1.RSSP_PHONE, SPTRSSP1.RSSP_EMAIL" & vbCrLf _
            '    & " from SPTRSSP1 where SELL_CODE = :PARM1" & vbCrLf _
            '    & " and RSSP_STATUS = 'A' and SPTRSSP1.RSSP_TYPE IN ('C', 'D') and SPTRSSP1.RSSP_DATE_TERM IS NULL"
            Create_TDA(.Tables.Add, "SOTCSTTX", "**", 0, False, "V", 1)

            .Tables.Add("SOTALLOD")
            With .Tables("SOTALLOD").Columns
                .Add("DATE_START", GetType(System.DateTime))
                .Add("ORDR_QTY", GetType(System.Int32))
                .Add("ORDR_QTY_OPEN", GetType(System.Int32))
                .Add("ORDR_QTY_PICK", GetType(System.Int32))
                .Add("ORDR_QTY_SHIP", GetType(System.Int32))
                .Add("ORDR_QTY_CANC", GetType(System.Int32))
                .Add("BALANCE", GetType(System.Int32))
                .Add("QTY_ALLO", GetType(System.Int32))
            End With
            'AND SOTALLO3.QTY_ALLO <> 0
            'AND SOTORDR1.ORDR_STATUS IN ('O','P','F','C')
            'AND SOTORDR1.CUST_STORE_NO (+) = SOTALLO3.CUST_STORE_NO
            ASCMAIN1.sql = $"Select SOTALLO1.DATE_START, SOTALLO1.ITEM_CODE, SOTALLO1.ALLO_CTL_NO
, SOTALLO3.QTY_ALLO, COUNT (DISTINCT SOTORDR1.ORDR_NO) ORDER_COUNT
, SUM (SOTORDR2.ORDR_QTY) ORDR_QTY
, SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN
, SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK
, SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP
, SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC
, SOTALLO1.DATE_END
from SOTALLO1,SOTALLO3,SOTORDR2,SOTORDR1
where SOTALLO3.CUST_STORE_NO = '000' || NVL(:PARM1,'000')
AND SOTALLO1.ALLO_CTL_NO = SOTALLO3.ALLO_CTL_NO
AND SOTALLO1.DATE_START = NVL(:PARM2,'01-JAN-2024')
AND SOTORDR2.ALLO_CTL_NO (+) = SOTALLO3.ALLO_CTL_NO
AND SOTORDR2.CUST_CODE (+) = SOTALLO3.CUST_CODE
AND SOTORDR2.CUST_STORE_NO (+) = SOTALLO3.CUST_STORE_NO
AND SOTORDR1.ORDR_NO (+) = SOTORDR2.ORDR_NO
group by SOTALLO1.DATE_START, SOTALLO1.ITEM_CODE, SOTALLO1.ALLO_CTL_NO, SOTALLO3.QTY_ALLO, SOTALLO1.DATE_END
"

            Create_TDA(.Tables.Add, "SOTALLOI", "**", 0, False, "VD", 2)
            With .Tables("SOTALLOI")
                .Columns.Add("BALANCE", GetType(System.Int32))
                .Columns("BALANCE").Expression = "ISNULL(QTY_ALLO,0)-ISNULL(ORDR_QTY_OPEN,0)-ISNULL(ORDR_QTY_PICK,0)-ISNULL(ORDR_QTY_SHIP,0)"
            End With

            Create_Relation("SOTALLOD", "SOTALLOI", "DATE_START")

            With .Tables("SOTALLOD")
                .Columns("ORDR_QTY").Expression = "SUM(CHILD.ORDR_QTY)"
                .Columns("ORDR_QTY_OPEN").Expression = "SUM(CHILD.ORDR_QTY_OPEN)"
                .Columns("ORDR_QTY_PICK").Expression = "SUM(CHILD.ORDR_QTY_PICK)"
                .Columns("ORDR_QTY_SHIP").Expression = "SUM(CHILD.ORDR_QTY_SHIP)"
                .Columns("ORDR_QTY_CANC").Expression = "SUM(CHILD.ORDR_QTY_CANC)"
                .Columns("BALANCE").Expression = "SUM(CHILD.BALANCE)"
                .Columns("QTY_ALLO").Expression = "SUM(CHILD.QTY_ALLO)"
            End With

            'ASCMAIN1.sql = "SELECT DISTINCT ICTWHSE2.ITEM_CODE, ICTWHSE2.WHSE_CODE
            '                FROM ICTWHSE1, ICTWHSE2
            '                WHERE ICTWHSE1.WHSE_CODE = ICTWHSE2.WHSE_CODE
            '                AND LP_CODE IN (SELECT LP_CODE FROM ICTWHSE1 WHERE WHSE_CODE = :PARM1)"
            'Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False, "V", 1)
            '2/5/25 Nathan wants rule where ADSMIN takes priority over CLA
            ASCMAIN1.sql =
            "SELECT ITEM_CODE, WHSE_CODE " &
            "  FROM ( " &
            "        SELECT ICTWHSE2.ITEM_CODE, " &
            "               ICTWHSE2.WHSE_CODE, " &
            "               ROW_NUMBER() OVER ( " &
            "                   PARTITION BY ICTWHSE2.ITEM_CODE " &
            "                   ORDER BY CASE " &
            "                       WHEN ICTWHSE2.WHSE_CODE = 'CLA' THEN 1 " &
            "                       WHEN ICTWHSE2.WHSE_CODE = 'ADSMIN' THEN 2 " &
            "                       ELSE 3 " &
            "                   END " &
            "               ) AS RN " &
            "          FROM ICTWHSE1, ICTWHSE2 " &
            "         WHERE ICTWHSE1.WHSE_CODE = ICTWHSE2.WHSE_CODE " &
            "       ) " &
            " WHERE RN = 1"

            Create_TDA(.Tables.Add, "ICTWHSEX", "**", 0, False, "V", 1)
            dst.Tables("ICTWHSEX").PrimaryKey = New DataColumn() {dst.Tables("ICTWHSEX").Columns("ITEM_CODE")}
        End With


        grdSOTCSTT1.DataSource = dst.Tables("SOTCSTT1")
        grdSOTCSTTX.DataSource = dst.Tables("SOTCSTTX")
        grdSOTCSTOX.DataSource = dst.Tables("SOTCSTOX")
        grdSOTCSTOI.DataSource = dst.Tables("SOTCSTOI")
        grdSOTALLOX.DataSource = dst.Tables("SOTALLOX")
        grdSOTALLOG.DataSource = dst.Tables("SOTALLOG")
        grdICTCOLL0.DataSource = dst.Tables("ICTCOLL0")
        grdICTPROD1.DataSource = dst.Tables("ICTPROD1")
        grdSOTRSCT1.DataSource = dst.Tables("SOTRSCT1")
        grdSOTALLOD.DataSource = dst.Tables("SOTALLOD")
        grdSOTCSTOA.DataSource = dst.Tables("SOTCSTOA")
        Create_Summary(grdSOTALLOX, "ALLO_CTL_NO", "Count")
        Fill_Records("SOTCSTOA")
        Sort_grdColumns(grdSOTCSTOA, "CSO_NO".ToLower)

        grdSOTCSTOX.DisplayLayout.UseFixedHeaders = True
        With grdSOTCSTOX.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"CSO_NO", "SELL_CODE"}
                .Columns(COLUMN_NAME).Header.Fixed = True
            Next
        End With



        grdSOTCSTTX.DisplayLayout.UseFixedHeaders = True
        With grdSOTCSTTX.DisplayLayout.Bands(0)

            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                With GCOL.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    If GCOL.Key.StartsWith("TAG_") Then
                        .BackColor2 = System.Drawing.Color.Orange
                        GCOL.CellActivation = Activation.AllowEdit
                    Else
                        GCOL.CellActivation = Activation.NoEdit
                        .BackColor2 = System.Drawing.Color.LightGreen
                    End If
                End With
            Next
        End With

        grdSOTCSTT1.DisplayLayout.UseFixedHeaders = True
        With grdSOTCSTT1.DisplayLayout.Bands(0)
            For Each GCOL As UltraWinGrid.UltraGridColumn In .Columns
                With GCOL.Header.Appearance
                    .BackColor = System.Drawing.Color.White
                    .BackGradientStyle = GradientStyle.ForwardDiagonal
                    .BackColor2 = System.Drawing.Color.Orange
                End With
            Next
        End With


        grdSOTALLOG.DisplayLayout.UseFixedHeaders = True
        grdSOTALLOG.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdSOTALLOG.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdSOTALLOG.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdSOTALLOG.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLUMN_NAME As String In New String() {"ALLO_GROUP_CODE", "ALLO_GROUP_DESC"}
                .Columns(COLUMN_NAME).CellActivation = Activation.NoEdit
            Next
        End With

        grdSOTRSCT1.DisplayLayout.UseFixedHeaders = True
        grdSOTRSCT1.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdSOTRSCT1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdSOTRSCT1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdSOTRSCT1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLUMN_NAME As String In New String() {"RSC_TAG"}
                .Columns(COLUMN_NAME).CellActivation = Activation.NoEdit
            Next
        End With

        grdICTCOLL0.DisplayLayout.UseFixedHeaders = True
        grdICTCOLL0.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdICTCOLL0.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdICTCOLL0.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdICTCOLL0.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLUMN_NAME As String In New String() {"HC_CODE", "HC_NAME"}
                .Columns(COLUMN_NAME).CellActivation = Activation.NoEdit
            Next
        End With

        grdICTPROD1.DisplayLayout.UseFixedHeaders = True
        grdICTPROD1.DisplayLayout.Override.AllowAddNew = AllowAddNew.No
        grdICTPROD1.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        grdICTPROD1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        With grdICTPROD1.DisplayLayout.Bands(0)
            For Each COLUMN_NAME As String In New String() {"SEL"}
                .Columns(COLUMN_NAME).Header.Fixed = True
                .Columns(COLUMN_NAME).CellActivation = Activation.AllowEdit
            Next
            For Each COLUMN_NAME As String In New String() {"PROD_CODE", "PROD_DESC"}
                .Columns(COLUMN_NAME).CellActivation = Activation.NoEdit
            Next
        End With

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdSOTCSTOX, grdSOTCSTOI, grdSOTALLOX}
            For b As Integer = 0 To grd.DisplayLayout.Bands.Count - 1
                With grd.DisplayLayout.Bands(b)
                    For Each gcol As UltraWinGrid.UltraGridColumn In .Columns
                        gcol.Header.Appearance.BackColor = Drawing.Color.White
                        gcol.Header.Appearance.BackColor2 = Drawing.Color.LightGray
                        gcol.Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    Next
                End With
            Next
        Next

        CUST_ADDR_cols = {"CSO_TYPE", "CUST_NAME", "CUST_ADDR1", "CUST_ADDR2", "CUST_ADDR3",
            "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE", "CUST_COUNTRY",
            "CUST_CONTACT", "CUST_PHONE", "CUST_EXT", "CUST_FAX", "CUST_EMAIL"}

        'ASCMAIN1.Add_Value_List(grdSOTCSTOX, "CSO_STATUS", , New String() {":", "O:Pending", "B:Ordered", "R:Released", "F:Shipped"})
        ASCMAIN1.Add_Value_List(grdSOTCSTOX, "ORDR_STATUS")
        ASCMAIN1.Add_Value_List(grdSOTCSTOX, "ORDR_STATUS",,, 1)
        ASCMAIN1.Add_Value_List(grdSOTCSTOX, "ORDR_STATUS",,, 2)

        Create_Summary(grdSOTCSTOX, "CSO_NO", "Count")

        Create_Summary(grdSOTCSTOI, "CSO_NO", "Count")

        Show_Filter(grdSOTCSTOX, True)
        grdSOTCSTOX.DisplayLayout.GroupByBox.Hidden = False

        Show_Filter(grdSOTCSTOI, True)
        grdSOTCSTOI.DisplayLayout.GroupByBox.Hidden = False

        dteCSOFrom.Value = Now.Date.AddDays(-30)
        dteCSOTo.Value = Now.Date
        dte_DELIVER_BY.Value = Nothing
        cmbEvent.Visible = False
        txtEvent.Visible = False


        Dim MenuItem As ToolStripItem = Nothing
        For i As Integer = WorkbookView1.ContextMenuStrip.Items.Count To 1 Step -1 ' Each MenuItem As ToolStripItem In WorkbookView1.ContextMenuStrip.Items
            MenuItem = WorkbookView1.ContextMenuStrip.Items(i - 1)
            If MenuItem.Text = "&Copy" Or MenuItem.Text = "&Paste" Then
                '  If MenuItem.Text = "Cu&t" Or MenuItem.Text = "&Copy" Or MenuItem.Text = "&Paste" Then
            Else
                MenuItem.Visible = False
                WorkbookView1.ContextMenuStrip.Items.Remove(MenuItem)
            End If
            If MenuItem.Text = "&Paste" Then
                ' AddHandler MenuItem.Click, AddressOf MenuItemPaste_Click
            End If
        Next

        MenuItem = WorkbookView1.ContextMenuStrip.Items.Add("Undo")
        AddHandler MenuItem.Click, AddressOf MenuItemUndo_Click
        'MenuItem = WorkbookView1.ContextMenuStrip.Items.Add("Copy to All Stores for Customer")
        'AddHandler MenuItem.Click, AddressOf MenuItemCopyNote_Click

        MakeTransparent(chkEditTags)
        tabMaster.Tabs("RSC Tags").Visible = False
    End Sub

    Sub Check_InquiryMode()
        InquiryMode = (MENU_ITEM_OBJECT = "SOFCSTOI")
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "New"
                If Absx1.txtFor("SELL_CODE").Text = "" Then
                    EMsg &= vbCr & "You Must First Specify an AE"
                Else
                    rowSOTSELL1 = LookUp("SOTSELL1", Absx1.txtFor("SELL_CODE").Text)
                    If rowSOTSELL1 Is Nothing Then
                        EMsg &= vbCr & "No Record of AE " & Absx1.txtFor("SELL_CODE").Text
                    Else
                        SELL_CODE = Absx1.txtFor("SELL_CODE").Text
                    End If
                End If

                If Absx1.dteFor("DATE_START").Value & "" = "" Then
                    EMsg &= vbCr & "You Must First Specify an Allocation (Start) Date by Double-clicking on any Allocation in the Grid below"
                End If

                If EMsg = "" Then
                    'Dim DATE_START_ora As String = Format(Absx1.dteFor("DATE_START").Value, "dd-MMM-yyyy")
                    'Dim DATE_START_fmt As String = Format(Absx1.dteFor("DATE_START").Value, "MM/dd/yyyy")
                    'ASCMAIN1.sql = $"Select * from SOTCSTO1 where SELL_CODE = '{SELL_CODE}' and DATE_START = '{DATE_START_ora}' and CSO_STATUS <> 'C' and CSO_STATUS <> 'D' AND CSO_SALES_HOLD = '1'"
                    'Dim rowSOTCSTO1 As DataRow = ASCDATA1.GetDataRow
                    'If rowSOTCSTO1 IsNot Nothing Then
                    '    EMsg &= vbCr & $"You Cannot Start a New CarStock Order with a Start Date of {DATE_START_fmt} while there is another existing CarStock Order for {DATE_START_fmt} that has been entered and is presently On Hold"
                    'End If
                End If

                If EMsg = "" Then
                    If Not ASCMAIN1.Logical_Lock("SOTCSTO1", SELL_CODE) Then Exit Sub
                End If

            Case "Edit", "View"

                Proceed_PreReq_Existing(eItemKey)

                'SELL_CODE = ""
                'CSO_NO = ""

                'If Absx1.txtFor("CSO_NO").Text = "" Then
                '    EMsg &= vbCr & "No Car Stock Order No Specified"
                'Else
                '    CSO_NO = Absx1.txtFor("CSO_NO").Text
                '    Dim row As DataRow = LookUp("SOTCSTO1", CSO_NO)
                '    If row Is Nothing Then
                '        EMsg &= vbCr & "No Record of Car Stock Order No " & CSO_NO
                '    Else
                '        Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & ""
                '        Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)

                '        SELL_CODE = row.Item("SELL_CODE")

                '        If ASCMAIN1.USER_CODES.Contains("FS") And SELL_CODE_this_user <> "" Then
                '            If SELL_CODE <> SELL_CODE_this_user Then
                '                EMsg &= vbCr & $"Invalid Entry for {SELL_CODE_this_user}"
                '            End If
                '        End If

                '        If EMsg = "" Then

                '            If eItemKey = "Edit" Then
                '                If rowSOTORDR0.Item("ORDR_STATUS") & "" <> "O" Then ' If row.Item("CSO_STATUS") & "" <> "O" Then
                '                    Dim msg As String = ""
                '                    Select Case rowSOTORDR0.Item("ORDR_STATUS")
                '                        Case "C"
                '                            msg = "Car Stock Order No " & CSO_NO & " has been Cancelled"
                '                        Case "D"
                '                            msg = "Car Stock Order No " & CSO_NO & " has been Deleted"
                '                        Case Else ' such as "F"
                '                            msg = "Car Stock Order No " & CSO_NO & " is No Longer Open (for Changes)"
                '                    End Select

                '                    EMsg &= vbCr & msg
                '                Else
                '                    ASCMAIN1.sql = $"Select SOTCSTO3.ORDR_NO, SOTORDR1.ORDR_STATUS from SOTORDR1,SOTCSTO3 where CSO_NO = '{CSO_NO}' and SOTORDR1.ORDR_NO = SOTCSTO3.ORDR_NO and SOTCSTO3.ORDR_NO is Not Null"
                '                    For Each rowORDR_NO As DataRow In ASCDATA1.GetDataTable.Select("")
                '                        Dim ORDR_NO As String = rowORDR_NO.Item("ORDR_NO")
                '                        Dim ORDR_STATUS As String = rowORDR_NO.Item("ORDR_STATUS")
                '                        If ORDR_STATUS = "P" Then
                '                            EMsg &= vbCr & $"Order {ORDR_NO} has already been released"
                '                            Exit For
                '                        End If
                '                        If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then Exit Sub
                '                    Next
                '                End If

                '                If EMsg.Length = 0 Then
                '                    If Not ASCMAIN1.Logical_Lock("SOTCSTO1", CSO_NO) Then Exit Sub
                '                    If Not ASCMAIN1.Logical_Lock("SOTCSTO1", SELL_CODE) Then Exit Sub
                '                End If
                '            End If

                '        End If

                '    End If
                'End If

                'If EMsg <> "" Then ASCMAIN1.MultiTask_Release()

            Case "Update"
                If Absx1.dteFor("CSO_DATE").Value & "" = "" Then
                    EMsg &= vbCr & "CSO Date is Mandatory"
                Else
                End If

                If Format(Absx1.dteFor("ORDR_SHIP_DATE").Value, "yyyyMMdd") > Format(Absx1.dteFor("ORDR_CANCEL_DATE").Value, "yyyyMMdd") Then
                    EMsg &= vbCr & "Ship Date may NOT be later than Cancel Date"
                End If

                'Dim updateResult As String = Update_Allocations()
                'If Not String.IsNullOrEmpty(updateResult) Then
                '    EMsg &= vbCr & updateResult
                'End If

                ' Check for negative balances and show popup if necessary
                If Has_Neg_Bal(True) Then
                    Show_Neg_Popup()
                End If

                ' Collect negative balance details for further processing or messaging
                Dim negativeBalances As String = Neg_Bal_Details()
                If Not String.IsNullOrEmpty(negativeBalances) Then
                    EMsg &= vbCr & negativeBalances
                End If

                If chkUrgent.Checked Then
                    If txtUrgent.Text & "" = "" Then
                        EMsg &= vbCr & "You must enter a note to mark this CSO as Urgent"
                    End If

                    If dte_DELIVER_BY.Value Is Nothing Then
                        EMsg &= vbCr & "You must enter a date to deliver by to mark this CSO as Urgent"
                    End If
                End If

                Dim CSO_TOTAL_QTY As Integer = dst.Tables("SOTCSTO2").AsEnumerable().Sum(Function(row) Convert.ToInt32(row("CSO_QTY_TOTAL")))
                If CSO_TOTAL_QTY = 0 Then
                    EMsg &= vbCr & "Cannot update with zero items on CSO"
                End If

                If UPDATE_DISABLED Then
                    EMsg &= vbCr & $"The number of items ({dst.Tables("SOTCSTO2").Rows.Count}) exceeds the maximum number supported by the screen ({MAX_COLs}). Please contact ABS."
                End If

                ''make sure the cancel date is within the allowed range
                'Dim currentCancelDate As Date = AddBusinessDays(rowSOTCSTO1.Item("ORDR_SHIP_DATE"), 8)
                'Dim endOfAllocationWindow As Date
                'Select Case currentCancelDate.Month
                '    Case 1 To 3
                '        endOfAllocationWindow = New Date(currentCancelDate.Year, 3, 15) 'last day to place orders Q1
                '    Case 4 To 6
                '        endOfAllocationWindow = New Date(currentCancelDate.Year, 6, 15)  'last day to place orders Q2
                '    Case 7 To 9
                '        endOfAllocationWindow = New Date(currentCancelDate.Year, 9, 15)  'last day to place orders Q3
                '    Case 10 To 12
                '        endOfAllocationWindow = New Date(currentCancelDate.Year, 11, 30)  'last day to place orders Q4
                'End Select

                '' Calculate potential MaxDate as 8 business days after the end of allocation window
                'Dim potentialMaxDate As Date = AddBusinessDays(endOfAllocationWindow, 8) '3/25, 6/26, 9/25, 12/12 LATEST CANCEL DATE
                'Dim userCancelDate As Date = Absx1.dteFor("ORDR_CANCEL_DATE").Value

                '' Validate the cancel date
                'If currentCancelDate > potentialMaxDate Then
                '    ' Default cancel date is past the allowed max date, inform the user
                '    EMsg &= vbCr & "The last day to place orders has passed. You cannot place an order after " & endOfAllocationWindow.ToShortDateString() & "."
                'Else
                '    ' Validate the user-provided cancel date
                '    If userCancelDate < currentCancelDate Then
                '        EMsg &= vbCr & $"The earliest cancel date allowed is {currentCancelDate.ToShortDateString()} (8 business days after the ship date)."
                '    ElseIf userCancelDate > potentialMaxDate Then
                '        EMsg &= vbCr & $"The latest cancel date allowed is {potentialMaxDate.ToShortDateString()} (8 business days after the end of the allocation window)."
                '    End If
                'End If

                If dst.Tables("SOTCSTO2").Rows.Count = 0 Then
                    EMsg &= vbCr & "No Items on Car Stock Order"
                Else
                    'Dim DT As New DataTable
                    'DT.Columns.Add("Order Number", GetType(String))
                    'DT.Columns.Add("Customer Name", GetType(String))
                    'DT.Columns.Add("Invalid Address", GetType(String))

                    'Dim CSO_NO As String = UltraTextEditor1.Text
                    'ASCMAIN1.sql = $"SELECT * FROM SOTCSTO3 WHERE CSO_NO = '{CSO_NO}' AND ORDR_NO IS NOT NULL"
                    'Dim invalidAddresses As New List(Of String)

                    'For Each rowSOTCSTO3 As DataRow In ASCDATA1.GetDataTable().Rows
                    '    Dim ORDR_NO As String = rowSOTCSTO3.Item("ORDR_NO").ToString()
                    '    Dim address As String = rowSOTCSTO3.Item("CUST_NAME").ToString() & vbCrLf & rowSOTCSTO3.Item("CUST_ADDR1").ToString() _
                    '        & If(Not String.IsNullOrEmpty(rowSOTCSTO3.Item("CUST_ADDR2").ToString()), vbCrLf & rowSOTCSTO3.Item("CUST_ADDR2").ToString(), "") _
                    '        & vbCrLf & rowSOTCSTO3.Item("CUST_CITY").ToString() & " " & rowSOTCSTO3.Item("CUST_STATE").ToString() & " " & rowSOTCSTO3.Item("CUST_ZIP_CODE").ToString()

                    '    Dim validationResponse As String = TAC.TACMAIN1.Validate_Address1(address)
                    '    If validationResponse.StartsWith("No candidates") OrElse validationResponse.StartsWith("Error") Then
                    '        invalidAddresses.Add($"Order {ORDR_NO} ({rowSOTCSTO3.Item("CUST_NAME").ToString()})")
                    '        DT.Rows.Add(ORDR_NO, rowSOTCSTO3.Item("CUST_NAME").ToString(), address)
                    '    End If
                    'Next

                    'If invalidAddresses.Count > 0 Then
                    '    Dim message As String = "The following Orders have invalid addresses."
                    '    Using F As New ASFMSGBF
                    '        F.Show_grd(DT, Me, message)
                    '    End Using
                    'End If
                End If

            Case "Delete"

                If EntryMode = "" Then
                    Exit Sub
                End If

                ASCMAIN1.sql = "Select Count (*) from SOTCSTO3 where ORDR_NO IS NOT NULL"
                ASCMAIN1.sql &= " and CSO_NO = '" & CSO_NO & "'"

                If Val(ASCDATA1.GetDataValue) <> 0 AndAlso Not (rowSOTCSTO1.Item("CSO_STATUS") & "" = "O") Then
                    EMsg &= vbCr & "Car Stock Order has been Used to create Sales Orders - Delete not permitted"
                Else
                    If EMsg = "" Then
                        If MsgBox("Do you want to Mark this Car Stock Order as Deleted",
                                  MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirmation") = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                End If

            Case "Restore DataSet"

                If Absx1.txtFor("SELL_CODE").Text = "" Or SELL_CODE = "" Then
                    EMsg &= vbCr & "You must first select an AE"
                Else


                    Dim FILENAME As String = ""
                    Using openFileDialog1 As New OpenFileDialog
                        openFileDialog1.InitialDirectory = ASCMAIN1.Folders("Work")
                        openFileDialog1.Title = "Select a DataSet to Restore"
                        openFileDialog1.Filter = "xml files (*.xml)|*.xml"
                        openFileDialog1.RestoreDirectory = True

                        ' Excel_Import = -1

                        If openFileDialog1.ShowDialog() = DialogResult.OK Then
                            FILENAME = openFileDialog1.FileName
                        End If
                    End Using

                    If FILENAME <> "" And FILENAME.ToLower.EndsWith(".xml") Then
                        dst2 = New DataSet

                        Try
                            dst2.ReadXml(FILENAME)
                            If dst2.DataSetName <> "SODCSTO1" Then
                                EMsg &= vbCr & "This XML file does not appear to have been originated from the Car-Stock screen"
                            End If

                            Dim rowSOTCSTO1 = dst2.Tables("SOTCSTO1").Rows(0)
                            If rowSOTCSTO1.ITEM("SELL_CODE") <> SELL_CODE Then
                                EMsg &= vbCr & $"This XML file does not appear to belong to AE {SELL_CODE}"
                            End If

                            If EMsg = "" Then
                                Absx1.txtFor("CSO_NO").Text = ""
                                CSO_NO = rowSOTCSTO1.ITEM("CSO_NO")
                                Dim rowExists As DataRow = LookUp("SOTCSTO1", CSO_NO)
                                If rowExists IsNot Nothing Then
                                    Absx1.txtFor("CSO_NO").Text = CSO_NO
                                    Proceed_PreReq_Existing("Edit")
                                    'EMsg &= vbCr & $"This XML file refers to CSO {CSO_NO} which appears to have been uodated"
                                End If
                            End If

                            If EMsg.Length = 0 Then
                                If Not ASCMAIN1.Logical_Lock("SOTCSTO1", CSO_NO) Then Exit Sub
                                If Not ASCMAIN1.Logical_Lock("SOTCSTO1", SELL_CODE) Then Exit Sub
                            End If

                        Catch ex As Exception
                            EMsg &= vbCr & ex.Message

                        Finally

                        End Try

                    Else
                        EMsg &= vbCr & "You must select a DataSet file to Restore"

                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Proceed(eItemKey)
    End Sub

    Sub Proceed_PreReq_Existing(eItemKey As String)

        SELL_CODE = ""
        CSO_NO = ""

        If Absx1.txtFor("CSO_NO").Text = "" Then
            EMsg &= vbCr & "No Car Stock Order No Specified"
        Else
            CSO_NO = Absx1.txtFor("CSO_NO").Text
            Dim row As DataRow = LookUp("SOTCSTO1", CSO_NO)
            If row Is Nothing Then
                EMsg &= vbCr & "No Record of Car Stock Order No " & CSO_NO
            Else
                Dim ORDR_GROUP_NOs As DataTable = ASCDATA1.GetDataTable($"SELECT DISTINCT ORDR_GROUP_NO, ORDR_STATUS FROM SOTORDR0 WHERE ORDR_GROUP_NO IN 
                                                                            (SELECT ORDR_GROUP_NO FROM SOTORDR1 WHERE
                                                                            ORDR_NO IN (SELECT ORDR_NO FROM SOTCSTO3 WHERE CSO_NO = '{CSO_NO}'))")

                'Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & ""
                'Dim rowSOTORDR0 As DataRow = LookUp("SOTORDR0", ORDR_GROUP_NO)

                SELL_CODE = row.Item("SELL_CODE")

                If ASCMAIN1.USER_CODES.Contains("FS") And SELL_CODE_this_user <> "" Then
                    If SELL_CODE <> SELL_CODE_this_user Then
                        EMsg &= vbCr & $"Invalid Entry for {SELL_CODE_this_user}"
                    End If
                End If

                If EMsg = "" Then

                    If eItemKey = "Edit" Then
                        If ORDR_GROUP_NOs.Rows.Count = 0 Then
                            EMsg &= vbCr & "There are no order groups for CSO No " & CSO_NO
                        ElseIf ORDR_GROUP_NOs.Select("ORDR_STATUS = 'O'").Length = 0 Then
                            Dim msg As String = ""
                            If ORDR_GROUP_NOs.Select("ORDR_STATUS = 'F'").Length > 0 Then
                                msg = "Car Stock Order No " & CSO_NO & " has some Order Groups that have been Finalized."
                            ElseIf ORDR_GROUP_NOs.Select("ORDR_STATUS = 'P'").Length > 0 Then
                                msg = "Car Stock Order No " & CSO_NO & " has some Order Groups that have been released."
                            ElseIf ORDR_GROUP_NOs.Select("ORDR_STATUS = 'C'").Length > 0 Then
                                msg = "Car Stock Order No " & CSO_NO & " has some Order Groups that have been Cancelled"
                            ElseIf ORDR_GROUP_NOs.Select("ORDR_STATUS = 'D'").Length > 0 Then
                                msg = "Car Stock Order No " & CSO_NO & " has been Deleted"
                            Else
                                msg = "Car Stock Order No " & CSO_NO & " is No Longer Open (for Changes)"
                            End If
                            'Select Case rowSOTORDR0.Item("ORDR_STATUS")
                            '    Case "C"
                            '        msg = "Car Stock Order No " & CSO_NO & " has been Cancelled"
                            '    Case "D"
                            '        msg = "Car Stock Order No " & CSO_NO & " has been Deleted"
                            '    Case Else ' such as "F"
                            '        msg = "Car Stock Order No " & CSO_NO & " is No Longer Open (for Changes)"
                            'End Select

                            EMsg &= vbCr & msg
                        Else
                            ASCMAIN1.sql = $"Select SOTCSTO3.ORDR_NO, SOTORDR1.ORDR_STATUS from SOTORDR1,SOTCSTO3 
                                            where CSO_NO = '{CSO_NO}' and SOTORDR1.ORDR_NO = SOTCSTO3.ORDR_NO and SOTCSTO3.ORDR_NO is Not Null"

                            Dim tbl As DataTable = ASCDATA1.GetDataTable
                            If tbl.Select("ORDR_STATUS = 'P'").Length > 0 Then
                                EMsg &= vbCr & $"There are {tbl.Select("ORDR_STATUS = 'P'").Length} Orders in Pick"
                            End If
                            If tbl.Select("ORDR_STATUS = 'F'").Length > 0 Then
                                EMsg &= vbCr & $"There are {tbl.Select("ORDR_STATUS = 'F'").Length} Orders that are Finalized"
                            End If
                            If EMsg.Length > 0 Then
                                Exit Sub
                            End If
                            For Each rowORDR_NO As DataRow In ASCDATA1.GetDataTable.Select("")
                                Dim ORDR_NO As String = rowORDR_NO.Item("ORDR_NO")
                                Dim ORDR_STATUS As String = rowORDR_NO.Item("ORDR_STATUS")
                                If ORDR_STATUS = "P" Then
                                    EMsg &= vbCr & $"Order {ORDR_NO} has already been released"
                                    Exit For
                                End If
                                If ORDR_STATUS = "F" Then
                                    EMsg &= vbCr & $"Order {ORDR_NO} has already been finalized"
                                    Exit For
                                End If
                                If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO) Then EMsg &= vbCr & $"Cannot Proceed" : Exit Sub
                            Next
                        End If

                        If EMsg.Length = 0 Then
                            If Not ASCMAIN1.Logical_Lock("SOTCSTO1", CSO_NO) Then EMsg &= vbCr & $"Cannot Proceed" : Exit Sub
                            If Not ASCMAIN1.Logical_Lock("SOTCSTO1", SELL_CODE) Then EMsg &= vbCr & $"Cannot Proceed" : Exit Sub
                        End If
                    End If

                End If

            End If
        End If

        If EMsg <> "" Then ASCMAIN1.MultiTask_Release()
    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "New"
                EntryMode = "N"
                Load_Record()
                If Not NO_ITEMS Then
                    Mode_Settings(True)
                    Show_Neg_Popup()
                Else
                    MsgBox("All items for this allocation date have passed their end date", MsgBoxStyle.Critical, "Cannot Create CSO")
                End If


            Case "Edit"
                EntryMode = "E"
                Load_Record()
                Mode_Settings(True)

            Case "Refresh"
                Load_SOTCSTOX()

            Case "Update"
                Validate_Addresses()
                If Not UPDATE_DISABLED Then
                    Save_DataSet(True)
                    Update_Record()

                    If chkUrgent.Checked Then
                        Urgent_Email()
                    Else
                        email_to_Self()
                    End If

                    Mode_Settings(False)
                End If
            Case "Delete"
                Delete_Order()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Print"
                Print_Record()

            Case "View"
                EntryMode = "V"
                Load_Record()
                Mode_Settings(True)

            Case "Done"
                Mode_Settings(False)

            Case "email to Self"
                email_to_Self()

            Case "Update RSC Tags"

                Update_RSC_Tags()
                chkEditTags.Checked = False

            Case "Cancel RSC Tags"
                chkEditTags.Checked = False
                Load_SOTCSTOX()

            Case "Save DataSet"
                Save_DataSet()

            Case "Restore DataSet"

                restore_in_process = True

                If Absx1.txtFor("CSO_NO").Text = "" Then
                    EntryMode = "N"
                Else
                    EntryMode = "E"
                End If

                Load_Record()
                Mode_Settings(True)

                restore_in_process = False

        End Select

    End Sub

    Sub Save_DataSet(Optional silent As Boolean = False)

        ASCMAIN1.Progress("Now Saving DataSet")

        Dim FOLDER_NAME As String = ASCMAIN1.Folders("Work")
        Dim FILE_NAME As String = Me.Name & $"_{CSO_NO}_{Format(Now, "yyyyMMddHHmmss")}" & ".dst"
        Write_DataSet(True, FOLDER_NAME, FILE_NAME)

        If Not System.IO.Directory.Exists(ASCMAIN1.Folders("Archive") & "SavedDataSets\") Then
            System.IO.Directory.CreateDirectory(ASCMAIN1.Folders("Archive") & "SavedDataSets\")
        End If

        System.IO.File.Copy(FOLDER_NAME & FILE_NAME & ".xml", ASCMAIN1.Folders("Archive") & "SavedDataSets\" & FILE_NAME & ".xml")

        If Not silent Then
            MsgBox($"DataSet Saved to {FOLDER_NAME}{FILE_NAME}", MsgBoxStyle.OkOnly, "Save Complete")
        End If

        ASCMAIN1.Progress("")
    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Set_ScreenMode_Base(tf)

        If Not ScreenMode Then
            chkAllowEditShipToAddress.Checked = False
        End If

        With UltraExplorerBar1
            With .Groups("Screen Control")
                .Items("New").Settings.Enabled = not_iScreenMode
                If (EntryMode = "V" And ScreenMode) Then
                    If rowSOTCSTO1.Item("CSO_STATUS") & "" = "O" Then
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.True
                    Else
                        .Items("Edit").Settings.Enabled = DefaultableBoolean.False
                    End If
                Else
                    .Items("Edit").Settings.Enabled = not_iScreenMode
                End If
                .Items("Update").Settings.Enabled = iScreenMode

                .Items("Delete").Settings.Enabled = iScreenMode
                .Items("Cancel").Settings.Enabled = iScreenMode
                .Items("View").Settings.Enabled = not_iScreenMode
                .Items("Done").Settings.Enabled = iScreenMode

                .Items("New").Visible = Not InquiryMode And Not ScreenMode
                .Items("Edit").Visible = Not InquiryMode And (Not ScreenMode Or EntryMode = "V")
                .Items("View").Visible = Not ScreenMode
                .Items("Refresh").Visible = Not ScreenMode
                .Items("Done").Visible = (EntryMode = "V" And ScreenMode)
                '.Items("Print").Visible = (EntryMode = "V" And ScreenMode) ' False ' ScreenMode
                .Items("Update").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)
                .Items("Delete").Visible = Not InquiryMode And (EntryMode = "E")
                .Items("Cancel").Visible = Not InquiryMode And (Not (EntryMode = "V") Or Not ScreenMode)

                .Items("email to Self").Visible = ScreenMode

                .Items("Save DataSet").Visible = ScreenMode And (EntryMode = "N" Or EntryMode = "E")
                .Items("Restore DataSet").Visible = Not ScreenMode

            End With

            .Groups("RSC Tags").Visible = False

            .Groups("Status").Visible = Not ScreenMode And InquiryMode

            .Groups("Item Info").Visible = ScreenMode
            .Groups("Allocation Groups").Visible = ScreenMode
            .Groups("High Collections").Visible = ScreenMode
            .Groups("Product Categories").Visible = ScreenMode
        End With

        tabAllocations.Tabs("Allocations").Visible = False

        lblStatus.Visible = ScreenMode

        grdSOTCSTOX.Visible = Not tf
        tabMaster.Visible = Not tf
        txtEvent.Visible = Not tf And chkEvent.Checked
        cmbEvent.Visible = Not tf And chkEvent.Checked
        chkEvent.Visible = Not tf

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        Set_Read_Only(frmCodes, Not (EntryMode = "E" Or EntryMode = "N"))
        Set_Read_Only(frmDates, Not (EntryMode = "E" Or EntryMode = "N"))

        If ScreenMode Then
            Set_Read_Only_for_ctl(Absx1.txtFor("WHSE_CODE"), True)
            Set_Read_Only_for_ctl(Absx1.dteFor("CSO_DATE"), True)
            Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_SHIP_DATE"), True)
            Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_CANCEL_DATE"), True)

            If ASCMAIN1.USER_CODES.Contains("FS") Then
                Set_Read_Only_for_ctl(Absx1.txtFor("SHIP_VIA_CODE"), True)
                Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_SHIP_DATE"), True)
                Set_Read_Only_for_ctl(Absx1.dteFor("ORDR_CANCEL_DATE"), True)
            End If
        Else
            Clear_Record()
        End If

    End Sub

    Sub Clear_Record()

        Absx1.txtFor("CSO_NO").Text = ""
        Absx1.txtFor("SELL_CODE").Text = SELL_CODE
        Absx1.txtFor("CSO_REF_NO").Text = ""
        chkEvent.Checked = False

        SELL_CODE = ""
        CSO_NO = ""
        UPDATE_DISABLED = False

        chkHideAddressColumns.Checked = False
        chkShowItemsLeft2Order.Checked = False
        UltraExplorerBar1.Groups("Item Info").Text = ""
        modifiedAddresses.Clear()
        originalValues.Clear()
        validatedAddresses.Clear()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"SOTCSTO1", "SOTCSTO2", "SOTCSTO3", "SOTCSTO4",
            "SOTALLOX", "SOTCSTOH", "SOTCSTOD",
            "SOTALLOD", "SOTCSTOX", "SOTORDXR", "SOTALLOI", "SOTCSTT1", "SOTCSTT2", "SOTCSTTX"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        If ASCMAIN1.USER_CODES.Contains("FS") And SELL_CODE_this_user <> "" Then
            Absx1.txtFor("SELL_CODE").Text = ""
            Absx1.txtFor("SELL_CODE").Text = SELL_CODE_this_user
            Set_Read_Only_for_ctl(Absx1.txtFor("SELL_CODE"), True)
            Load_SOTCSTOX()
        End If

        optRSC_FM.Value = "ALL"

        ' Load_SOTCSTOX()
    End Sub

    Sub Load_Record()

        Save_Header_Fields(UltraGroupBox1)

        EnforceConstraints(False)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data")

        'Absx1.dteFor("ORDR_CANCEL_DATE").MaxDate = CDate("12/31/9998")

        If EntryMode = "N" And Not restore_in_process Then
            CSO_NO = ASCMAIN1.Next_Control_No("SOTCSTO1.CSO_NO")

            SELL_CODE = HFs("SELL_CODE")
            rowSOTSELL1 = LookUp("SOTSELL1", SELL_CODE)

            rowSOTCSTO1 = dst.Tables("SOTCSTO1").NewRow
            With rowSOTCSTO1
                .Item("CSO_NO") = CSO_NO
                .Item("SELL_CODE") = SELL_CODE
                .Item("CUST_CODE") = CUST_CODE
                .Item("CSO_DATE") = DATETIME_STAMP.Date
                .Item("DATE_START") = HFs("DATE_START")


                Dim CSO_REF_NO As String = ""
                Dim CSO_REF_NO_base As String = SELL_CODE & "-" & Format(Now.Date, "MMddyyyy") & "-Auto"
                ASCMAIN1.sql = $"Select Distinct ORDR_CUST_PO From SOTORDR1 where CUST_CODE = '{CUST_CODE}' and ORDR_CUST_PO LIKE '{CSO_REF_NO_base}%'"
                Dim sfx As Integer = -1
                Dim tbl As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql, "CSO_REF_NO", 1)
                Dim rowPO As DataRow = Nothing
                Do
                    sfx += 1
                    If sfx = 0 Then
                        CSO_REF_NO = CSO_REF_NO_base
                    Else
                        CSO_REF_NO = CSO_REF_NO_base & "-" & CStr(sfx)
                    End If
                    rowPO = tbl.Rows.Find(CSO_REF_NO)
                Loop Until rowPO Is Nothing

                '370-12102023-Auto MMDDYYYY
                .Item("CSO_REF_NO") = CSO_REF_NO ' HFs("CSO_REF_NO")
                .Item("CSO_STATUS") = "O"
                .Item("INIT_DATE") = DATETIME_STAMP
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                Dim WHSE_CODE As String = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                .Item("WHSE_CODE") = WHSE_CODE

                'Carstock Order Date: Today (Example: February 1, 2024) 
                Dim ORDR_SHIP_DATE As Date ' rowSOTCSTO1.Item("DATE_START")
                '10/2/24, ship date defaults to allo date if allo date is greater than today's date
                If CDate(HFs("DATE_START")) > DATETIME_STAMP.Date Then
                    ORDR_SHIP_DATE = HFs("DATE_START")
                Else
                    ORDR_SHIP_DATE = Now.Date
                End If
                'Start Ship Date: Today + 1 Business Day (NF 10/16/24)
                .Item("ORDR_SHIP_DATE") = AddBusinessDays(ORDR_SHIP_DATE, 1)
                'Cancel Date: Start Ship Date + 8 Days 
                .Item("ORDR_CANCEL_DATE") = AddBusinessDays(ORDR_SHIP_DATE, 8)
                .Item("SHIP_VIA_CODE") = rowARTCUST1.Item("SHIP_VIA_CODE")
            End With
            dst.Tables("SOTCSTO1").Rows.Add(rowSOTCSTO1)

            dst.Tables("SOTCSTO2").Rows.Clear()
            Dim DATE_START As Date = rowSOTCSTO1.Item("DATE_START")

            ASCMAIN1.sql = sqlICTITEM1 & $" and ICTITEM1.ITEM_CODE in (Select ITEM_CODE from {SOTALLOX} where DATE_START = '{Format(DATE_START, "dd-MMM-yyyy")}')"
            Fill_Records("ICTITEM1",,, ASCMAIN1.sql)

            ASCMAIN1.Progress("Allocations")
            Dim sqlx As String = $"DATE_START = '{Format(DATE_START, "MM/dd/yyyy")}'"
            Dim allocations As Integer = 0
            Dim eventFilter As Boolean = chkEvent.Checked AndAlso Not String.IsNullOrEmpty(cmbEvent.Text)
            Dim processedItems As New HashSet(Of String)
            For Each row As DataRow In dst.Tables("SOTALLOX").Select(sqlx, "ITEM_CODE")
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")
                Dim DATE_END As String = row.Item("DATE_END")
                'Dim ALLO_END As String 
                'If allocations > 10 Then Exit For
                ASCMAIN1.Progress("-", ITEM_CODE)
                Dim ALLO_CTL_NO As String = row.Item("ALLO_CTL_NO")

                If processedItems.Contains(ITEM_CODE) Then
                    Continue For
                End If

                Dim EVENT_TYPE_DESC As String = If(row.Table.Columns.Contains("EVENT"), row.Item("EVENT") & "", "")
                Dim QTY_ALLO As Int64 = Val(row.Item("QTY_ALLO") & "")

                If eventFilter Then
                    Dim eventRows() As DataRow = dst.Tables("SOTALLOX").Select($"ALLO_CTL_NO = '{ALLO_CTL_NO}' AND EVENT = '{cmbEvent.Text}'")

                    If eventRows.Length = 0 Then
                        Continue For
                    End If

                    QTY_ALLO = Val(row.Item("EVENT_QTY") & "")
                End If

                'DONT SHOW ITEMS WHOSE END DATE HAS PASSED
                If QTY_ALLO > 0 AndAlso DATE_END >= Date.Today Then
                    allocations += 1
                    Dim rowSOTCSTO2 As DataRow = Add_SOTCSTO2(ITEM_CODE, QTY_ALLO, allocations)
                    'rowSOTCSTO2.Item("CSO_QTY_ALLO") = QTY_ALLO
                    rowSOTCSTO2.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                    processedItems.Add(ITEM_CODE)
                    'READONLY FOR ITEMS WITHIN 2 WEEKS OF DATE END - 9/3/24 ANA SAYS TO ABANDON RULE FOR NOW
                    'If DATE_END <= Date.Today.AddDays(14) Then
                    '    READ_ONLY.Add(ITEM_CODE)
                    'End If
                End If
            Next

            dst.Tables("SOTCSTO3").Rows.Clear()
            Dim CSO_ADDR_LNO As Integer = 0

            ASCMAIN1.sql = $"Select * from SOTSELL1 where SELL_CODE = '{SELL_CODE}' or SELL_CODE_MGR = '{SELL_CODE}' AND SELL_STATUS = 'A'"
            For Each rowSOTSELL1 As DataRow In ASCDATA1.GetDataTable().Select("", "SELL_CODE_MGR")
                Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").NewRow
                With rowSOTCSTO3
                    .Item("CSO_NO") = CSO_NO
                    CSO_ADDR_LNO += 1
                    ASCMAIN1.Progress("-", CStr(CSO_ADDR_LNO))
                    .Item("CSO_ADDR_LNO") = CSO_ADDR_LNO
                    Dim SELL_CODE_MGR As String = rowSOTSELL1.Item("SELL_CODE_MGR") & ""
                    .Item("CSO_TYPE") = rowSOTSELL1.Item("SELL_TYPE")

                    If SELL_CODE_MGR = "" Then
                        .Item("CSO_INDEX") = 1
                        If .Item("CSO_TYPE") & "" <> "AE" Then
                            MsgBox("Mis-Configured record for Code " & rowSOTSELL1.Item("SELL_CODE"), MsgBoxStyle.OkOnly, "Please Report this to Sales Admin - Do Not Proceed")
                        End If
                        '.Item("CSO_TYPE") = "AE"
                    Else
                        .Item("CSO_INDEX") = 2
                        If .Item("CSO_TYPE") & "" <> "AC" Then
                            MsgBox("Mis-Configured record for Code " & rowSOTSELL1.Item("SELL_CODE"), MsgBoxStyle.OkOnly, "Please Report this to Sales Admin - Do Not Proceed")
                        End If
                        '.Item("CSO_TYPE") = "AC"
                    End If
                    .Item("CSO_KEY") = rowSOTSELL1.Item("SELL_CODE")
                    .Item("ORDR_NO") = ""
                    For Each C As String In CUST_ADDR_cols
                        If C = "CUST_CONTACT" Then
                        ElseIf C = "CSO_TYPE" Then
                            .Item(C) = rowSOTSELL1.Item("SELL_TYPE")
                        Else
                            .Item(C) = rowSOTSELL1.Item("SELL_" & Mid(C, 6))
                        End If

                    Next
                End With
                dst.Tables("SOTCSTO3").Rows.Add(rowSOTCSTO3)
            Next

            Dim SELL_TYPE As String = rowSOTSELL1.Item("SELL_TYPE") & ""
            If SELL_TYPE = "AC" Then
                ASCMAIN1.sql = $"SELECT SPTRSSP1.RSSP_CODE, SPTRSSP1.RSSP_NAME, SPTRSSP1.RSSP_SHIP_TO_ADDR1 AS RSSP_ADDR1, SPTRSSP1.RSSP_SHIP_TO_ADDR2 AS RSSP_ADDR2, SPTRSSP1.RSSP_SHIP_TO_ADDR3 AS RSSP_ADDR3, " & vbCrLf _
                            & "SPTRSSP1.RSSP_SHIP_TO_CITY AS RSSP_CITY, SPTRSSP1.RSSP_SHIP_TO_STATE AS RSSP_STATE, SPTRSSP1.RSSP_SHIP_TO_ZIP_CODE AS RSSP_ZIP_CODE, SPTRSSP1.RSSP_SHIP_TO_COUNTRY AS RSSP_COUNTRY, SPTRSSP1.RSSP_PHONE, SPTRSSP1.RSSP_EXT, SPTRSSP1.RSSP_FAX, " & vbCrLf _
                            & "SPTRSSP1.RSSP_EMAIL, SPTRSSP1.RSSP_TYPE " & vbCrLf _
                            & "FROM SPTCWRX2,ARTCUST2,SPTRSSP1, SOTSELL1 " & vbCrLf _
                            & " where SPTCWRX2.WORK_DATE > SYSDATE - 365" & vbCrLf _
                            & " and ARTCUST2.CUST_CODE = SPTCWRX2.CUST_CODE" & vbCrLf _
                            & " and ARTCUST2.CUST_STORE_NO = SPTCWRX2.CUST_STORE_NO" & vbCrLf _
                            & $" and ARTCUST2.SELL_CODE_AC = '{SELL_CODE}'" & vbCrLf _
                            & " and SPTRSSP1.RSSP_ID = SPTCWRX2.SSN" & vbCrLf _
                            & " and NVL(SPTRSSP1.RSSP_STATUS, 'A') = 'A' and SPTRSSP1.RSSP_TYPE IN ('C', 'D') and SPTRSSP1.RSSP_DATE_TERM IS NULL " & vbCrLf _
                            & "UNION " & vbCrLf _
                            & "SELECT SATAEBM1.BUS_MGR_CODE AS RSSP_CODE, SATAEBM1.BUS_MGR_NAME AS RSSP_NAME, SATAEBM1.BUS_MGR_ADDR1 AS RSSP_ADDR1, " & vbCrLf _
                            & "SATAEBM1.BUS_MGR_ADDR2 AS RSSP_ADDR2, SATAEBM1.BUS_MGR_ADDR3 AS RSSP_ADDR3, SATAEBM1.BUS_MGR_CITY AS RSSP_CITY, " & vbCrLf _
                            & "SATAEBM1.BUS_MGR_STATE AS RSSP_STATE, SATAEBM1.BUS_MGR_ZIP_CODE AS RSSP_ZIP_CODE, SATAEBM1.BUS_MGR_COUNTRY AS RSSP_COUNTRY, " & vbCrLf _
                            & "SATAEBM1.BUS_MGR_PHONE AS RSSP_PHONE, SATAEBM1.BUS_MGR_EXT AS RSSP_EXT, SATAEBM1.BUS_MGR_FAX AS RSSP_FAX, SATAEBM1.BUS_MGR_EMAIL AS RSSP_EMAIL, " & vbCrLf _
                            & "NULL AS RSSP_TYPE " & vbCrLf _
                            & $"FROM SATAEBM1 WHERE SATAEBM1.SELL_CODE = '{SELL_CODE}'"
            Else
                'ASCMAIN1.sql = $"Select * from SPTRSSP1 where SELL_CODE = '{SELL_CODE}' and RSSP_STATUS = 'A' and SPTRSSP1.RSSP_TYPE IN ('C', 'D') and SPTRSSP1.RSSP_DATE_TERM IS NULL"
                ASCMAIN1.sql = $"SELECT RSSP_CODE, RSSP_NAME, RSSP_SHIP_TO_ADDR1 AS RSSP_ADDR1, RSSP_SHIP_TO_ADDR2 AS RSSP_ADDR2, RSSP_SHIP_TO_ADDR3 AS RSSP_ADDR3,
                RSSP_SHIP_TO_CITY AS RSSP_CITY, RSSP_SHIP_TO_STATE AS RSSP_STATE, RSSP_SHIP_TO_ZIP_CODE AS RSSP_ZIP_CODE, RSSP_SHIP_TO_COUNTRY AS RSSP_COUNTRY, RSSP_PHONE, RSSP_EXT, RSSP_FAX,
                RSSP_EMAIL, RSSP_TYPE
                FROM SPTRSSP1
                WHERE SELL_CODE = '{SELL_CODE}'
                AND RSSP_STATUS = 'A' 
                AND RSSP_TYPE IN ('C', 'D') 
                AND RSSP_DATE_TERM IS NULL
                UNION
                SELECT BUS_MGR_CODE AS RSSP_CODE, BUS_MGR_NAME AS RSSP_NAME, BUS_MGR_ADDR1 AS RSSP_ADDR1,
                    BUS_MGR_ADDR2 AS RSSP_ADDR2, BUS_MGR_ADDR3 AS RSSP_ADDR3, BUS_MGR_CITY AS RSSP_CITY,
                    BUS_MGR_STATE AS RSSP_STATE, BUS_MGR_ZIP_CODE AS RSSP_ZIP_CODE, BUS_MGR_COUNTRY AS RSSP_COUNTRY,
                    BUS_MGR_PHONE AS RSSP_PHONE, BUS_MGR_EXT AS RSSP_EXT, BUS_MGR_FAX AS RSSP_FAX, BUS_MGR_EMAIL AS RSSP_EMAIL,
                    NULL AS RSSP_TYPE
                FROM 
                SATAEBM1
                LEFT JOIN SOTSELL1 ON SOTSELL1.SELL_CODE_MGR = SATAEBM1.SELL_CODE
                WHERE 
                SATAEBM1.BUS_MGR_STATUS = 'A' 
                AND (
                SATAEBM1.SELL_CODE = '{SELL_CODE}'
                OR SATAEBM1.SELL_CODE IN (SELECT SELL_CODE FROM SOTSELL1 WHERE SELL_CODE_MGR = '{SELL_CODE}'))"
            End If

            For Each rowSPTRSSP1 As DataRow In ASCDATA1.GetDataTable().Select("", "RSSP_NAME") ' "RSSP_CODE")
                Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").NewRow
                With rowSOTCSTO3
                    .Item("CSO_NO") = CSO_NO
                    CSO_ADDR_LNO += 1
                    ASCMAIN1.Progress("-", CStr(CSO_ADDR_LNO))
                    .Item("CSO_ADDR_LNO") = CSO_ADDR_LNO
                    .Item("CSO_INDEX") = 3
                    Dim RSSP_TYPE_CODE As String = rowSPTRSSP1.Item("RSSP_TYPE") & ""
                    If RSSP_TYPE_CODE = "D" Then
                        RSSP_TYPE_CODE = "SDS"
                    ElseIf RSSP_TYPE_CODE = "C" Then
                        RSSP_TYPE_CODE = "RSC"
                    Else
                        '?
                        RSSP_TYPE_CODE = "BM"
                    End If
                    .Item("CSO_TYPE") = RSSP_TYPE_CODE ' rowSPTRSSP1.Item("RSSP_TYPE") ' "RSC"
                    .Item("CSO_KEY") = rowSPTRSSP1.Item("RSSP_CODE")
                    .Item("ORDR_NO") = ""
                    For Each C As String In CUST_ADDR_cols
                        If C = "CUST_CONTACT" Then
                            ' SKIP
                        ElseIf C = "CSO_TYPE" Then
                            .Item(C) = RSSP_TYPE_CODE
                        Else
                            .Item(C) = rowSPTRSSP1.Item("RSSP_" & Mid(C, 6))
                        End If

                    Next
                End With
                dst.Tables("SOTCSTO3").Rows.Add(rowSOTCSTO3)
            Next

            If dst.Tables("SOTCSTO2").Rows.Count > MAX_COLs Then
                UPDATE_DISABLED = True
            End If

        Else ' EntryMode = "E"

            If restore_in_process Then
                rowSOTCSTO1 = dst2.Tables("SOTCSTO1").Rows.Find(CSO_NO)
                dst.Tables("SOTCSTO1").Rows.Add(rowSOTCSTO1.ItemArray)
                rowSOTCSTO1 = dst.Tables("SOTCSTO1").Rows.Find(CSO_NO)
            Else
                rowSOTCSTO1 = Fill_Record("SOTCSTO1", CSO_NO)
            End If

            SELL_CODE = rowSOTCSTO1.Item("SELL_CODE")
            rowSOTSELL1 = LookUp("SOTSELL1", SELL_CODE)

            Dim ORDR_GROUP_NO As String = rowSOTCSTO1.Item("ORDR_GROUP_NO") & ""
            Create_Work_Tables_SOTALLOX(ORDR_GROUP_NO)

            If restore_in_process Then
                Dim SOTCSTO2qty As Int32 = 0
                Dim SOTCSTO3qty As Int32 = 0
                For Each T As String In New String() {"SOTCSTO2", "SOTCSTO3", "SOTCSTO4"}
                    For Each row2 As DataRow In dst2.Tables(T).Select("")
                        Dim row As DataRow = dst.Tables(T).NewRow
                        row.ItemArray = row2.ItemArray
                        dst.Tables(T).Rows.Add(row)
                        If T = "SOTCSTO2" Then SOTCSTO2qty += Val(row.Item("CSO_QTY_TOTAL") & "")
                        If T = "SOTCSTO3" Then SOTCSTO3qty += Val(row.Item("CSO_QTY_TOTAL") & "")
                    Next
                Next

                'If SOTCSTO2qty <> SOTCSTO3qty Then
                '    If ASCMAIN1.Running_in_VS Then
                '        MsgBox($"Qty Mismatch {CStr(SOTCSTO2qty)} vs {CStr(SOTCSTO3qty)}", MsgBoxStyle.OkOnly, "Qtys by Item not in balance with Qtys by Ship-To")
                '    End If
                'End If

                'Dim dvw2 As DataView = dst.Tables("SOTCSTO2").DefaultView
                'dvw2.RowFilter = "CSO_QTY_TOTAL <> 0"
                'Dim dt2 As DataTable = dvw2.ToTable

                'Dim dvw3 As DataView = dst.Tables("SOTCSTO3").DefaultView
                'dvw3.RowFilter = "CSO_QTY_TOTAL <> 0"
                'Dim dt3 As DataTable = dvw3.ToTable


                dst.Tables("SOTCSTO4").Rows.Clear()

                For Each row As DataRow In dst.Tables("SOTCSTO3").Select("CSO_QTY_TOTAL <> 0")
                    For i As Integer = 1 To MAX_COLs
                        Dim CSO_QTY As Int32 = Val(row.Item($"CSO_QTY_{Format(i, "000")}") & "")
                        If CSO_QTY > 0 Then
                            Dim rowSOTCSTO4 As DataRow = dst.Tables("SOTCSTO4").NewRow
                            rowSOTCSTO4.Item("CSO_NO") = row.Item("CSO_NO")
                            rowSOTCSTO4.Item("CSO_LNO") = i
                            rowSOTCSTO4.Item("CSO_ADDR_LNO") = row.Item("CSO_ADDR_LNO")
                            rowSOTCSTO4.Item("CSO_QTY") = CSO_QTY
                            Debug.Print("4" & ":" & CStr(i) & ":" & CStr(row.Item("CSO_ADDR_LNO")) & ":" & CStr(CSO_QTY))
                            dst.Tables("SOTCSTO4").Rows.Add(rowSOTCSTO4)
                        End If
                    Next
                Next

                Dim DATE_START As Date = rowSOTCSTO1.Item("DATE_START")

                ASCMAIN1.sql = sqlICTITEM1 & $" and ICTITEM1.ITEM_CODE in (Select ITEM_CODE from {SOTALLOX} where DATE_START = '{Format(DATE_START, "dd-MMM-yyyy")}')"
                Fill_Records("ICTITEM1",,, ASCMAIN1.sql)
            Else
                Fill_Records("SOTCSTO2", CSO_NO)
                Fill_Records("SOTCSTO3", CSO_NO)
                Fill_Records("SOTCSTO4", CSO_NO)

                ASCMAIN1.sql = sqlICTITEM1 & $" and ICTITEM1.ITEM_CODE in (Select ITEM_CODE from SOTCSTO2 where CSO_NO = '{CSO_NO}')"
                Fill_Records("ICTITEM1",,, ASCMAIN1.sql)
            End If



            For Each rowSOTCSTO2 As DataRow In dst.Tables("SOTCSTO2").Select()
                With rowSOTCSTO2
                    Dim ITEM_CODE As String = .Item("ITEM_CODE")
                    'Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
                    ''.Item("CSO_QTY") = QTY
                    '.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                    '.Item("ITEM_SNU_CODE") = rowICTITEM1.Item("ITEM_SNU_CODE")
                    '.Item("ITEM_SO_QTY_MULT") = rowICTITEM1.Item("ITEM_SO_QTY_MULT")
                    '.Item("ITEM_SO_QTY_MIN") = rowICTITEM1.Item("ITEM_SO_QTY_MIN")
                    '.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                    '.Item("HC_CODE") = rowICTITEM1.Item("HC_CODE")
                    '.Item("BRAND_CODE") = rowICTITEM1.Item("BRAND_CODE")
                    Dim rowSOTALLOXs() As DataRow = dst.Tables("SOTALLOX").Select($"ITEM_CODE = '{ITEM_CODE}'")
                    If rowSOTALLOXs.Length = 1 Then
                        .Item("ALLO_GROUP_CODE") = rowSOTALLOXs(0).Item("ALLO_GROUP_CODE")
                        '.Item("CSO_QTY_ALLO") = Val(row(0).Item("QTY_ALLO") & "")
                    End If
                    Dim CSO_LNO As Integer = Val(.Item("CSO_LNO") & "")

                    '' THE NEXT 3 LINES ARE TO CORRECT MIS-PLACEMENT OF QTYS BY ITEMS IN 2, RELYING INSTEAD ON 3
                    'For i As Integer = 1 To MAX_COLs
                    '    .Item($"CSO_QTY_{Format(i, "000")}") = 0
                    'Next


                    'Debug.Print(CSO_LNO)
                    For Each rowSOTCSTO4 As DataRow In dst.Tables("SOTCSTO4").Select($"CSO_LNO = {CStr(CSO_LNO)}")
                        Dim CSO_QTY As Integer = Val(rowSOTCSTO4.Item("CSO_QTY") & "")
                        Dim CSO_ADDR_LNO As Integer = Val(rowSOTCSTO4.Item("CSO_ADDR_LNO") & "")
                        If CSO_QTY <> 0 Then Debug.Print("4->2" & ":" & CStr(CSO_LNO) & ":" & CStr(CSO_ADDR_LNO) & ":" & CStr(CSO_QTY))
                        .Item($"CSO_QTY_{Format(CSO_ADDR_LNO, "000")}") = CSO_QTY

                        Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Rows.Find(New Object() {CSO_NO, CSO_ADDR_LNO})
                        rowSOTCSTO3.Item("CSO_QTY_" & Format(CSO_LNO, "000")) = CSO_QTY
                    Next
                End With
            Next
        End If

        Fill_Records("SOTALLOX")

        'Dim ITEM_CODEs As New List(Of String)
        For Each rowSOTCSTO2 As DataRow In dst.Tables("SOTCSTO2").Select("", "ITEM_CODE")
            Dim ITEM_CODE As String = rowSOTCSTO2.Item("ITEM_CODE")
            'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "CC018V14USA" Then Stop
            'ITEM_CODEs.Add(ITEM_CODE)
            Dim ALLO_CTL_NO As String = rowSOTCSTO2.Item("ALLO_CTL_NO")
            'If ASCMAIN1.Running_in_VS AndAlso ITEM_CODE = "KS007P80" Then Stop
            Dim matchingRows() As DataRow = dst.Tables("SOTALLOX").Select($"ALLO_CTL_NO = '{ALLO_CTL_NO}' AND ITEM_CODE = '{ITEM_CODE}'")
            If matchingRows.Length > 0 Then
                Dim eventFilter As Boolean = chkEvent.Checked AndAlso Not String.IsNullOrEmpty(cmbEvent.Text)
                Dim QTY_ALLO As Integer = 0
                Dim selectedRow As DataRow = matchingRows(0)
                If eventFilter Then
                    Dim eventRows() As DataRow = matchingRows.Where(Function(r) Not IsDBNull(r("EVENT")) AndAlso r("EVENT").ToString() = cmbEvent.Text).ToArray()

                    If eventRows.Length > 0 Then
                        selectedRow = eventRows(0)
                        QTY_ALLO = Val(selectedRow.Item("EVENT_QTY") & "")
                    Else
                        Continue For
                    End If
                Else
                    QTY_ALLO = Val(selectedRow.Item("QTY_ALLO") & "")
                End If

                rowSOTCSTO2.Item("QTY_ALLO") = QTY_ALLO

                For Each COL As String In {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC",
                                   "QTY_LEFT", "WHSE_QTY_ON_HAND", "WHSE_QTY_ONPO", "WHSE_QTY_OPEN", "WHSE_QTY_PICK"}
                    rowSOTCSTO2.Item(COL) = selectedRow.Item(COL)
                Next
            End If
        Next
        'cbeFindItem.DataSource = ITEM_CODEs
        cbeFindItemList.DataSource = dst.Tables("SOTCSTO2")

        Load_SSG()
        If NO_ITEMS Then
            Exit Sub
        End If


        ASCMAIN1.Progress("Address Columns")

        ASCMAIN1.Progress("Item Columns")
        Dim CSO_COL As Integer = 0
        For Each rowSOTCSTO2 As DataRow In dst.Tables("SOTCSTO2").Select("", "CSO_LNO")
            CSO_COL += 1
            Dim C As String = $"CSO_QTY_{Format(CSO_COL, "000")}"
            rowSOTCSTO2.Item("CSO_COL") = CSO_COL
        Next

        dst.Tables("SOTALLOG").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct("SOTALLOX", New String() {"ALLO_GROUP_CODE"}).Select()
            Dim ALLO_GROUP_CODE As String = row.Item("ALLO_GROUP_CODE") & ""
            Dim ALLO_GROUP_DESC As String = "Ungrouped"
            Dim ALLO_GROUP_STATUS As String = "A"
            If ALLO_GROUP_CODE <> "" Then
                LookUp("SOTALLOG", ALLO_GROUP_CODE)
                If cdr IsNot Nothing Then
                    ALLO_GROUP_DESC = cdr.Item("ALLO_GROUP_DESC") & ""
                    ALLO_GROUP_STATUS = cdr.Item("ALLO_GROUP_STATUS") & ""
                End If
                dst.Tables("SOTALLOG").Rows.Add(New String() {ALLO_GROUP_CODE, ALLO_GROUP_DESC, ALLO_GROUP_STATUS, Nothing, "", Nothing, "", "1"})
            End If
        Next
        If dst.Tables("SOTALLOG").Rows.Find("") Is Nothing Then
            dst.Tables("SOTALLOG").Rows.Add(New Object() {"", "Ungrouped", "A", DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, "1"})
        End If
        Sort_grdColumns(grdSOTALLOG, "ALLO_GROUP_DESC")

        dst.Tables("ICTCOLL0").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct("SOTCSTO2", New String() {"HC_CODE"}).Select()
            Dim HC_CODE As String = row.Item("HC_CODE") & ""
            Dim HC_NAME As String = "?"

            LookUp("ICTCOLL0", HC_CODE)
            If cdr IsNot Nothing Then
                HC_NAME = cdr.Item("HC_NAME")
            End If
            dst.Tables("ICTCOLL0").Rows.Add(New String() {HC_CODE, HC_NAME, "1"})
        Next
        Sort_grdColumns(grdICTCOLL0, "HC_NAME")

        dst.Tables("ICTPROD1").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct("SOTCSTO2", New String() {"PROD_CODE"}).Select()
            Dim PROD_CODE As String = row.Item("PROD_CODE") & ""
            Dim PROD_DESC As String = "?"

            LookUp("ICTPROD1", PROD_CODE)
            If cdr IsNot Nothing Then
                PROD_DESC = cdr.Item("PROD_DESC")
            End If
            dst.Tables("ICTPROD1").Rows.Add(New String() {PROD_CODE, PROD_DESC, "1"})
        Next
        Sort_grdColumns(grdICTPROD1, "PROD_DESC")

        dst.Tables("SOTRSCT1").Rows.Clear()
        For Each row As DataRow In ASCDATA1.SelectDistinct("SOTCSTT1", New String() {"RSC_TAG"}).Select()
            Dim RSC_TAG As String = row.Item("RSC_TAG") & ""
            dst.Tables("SOTRSCT1").Rows.Add(New String() {"1", RSC_TAG})
        Next
        Sort_grdColumns(grdSOTRSCT1, "RSC_TAG")

        ASCMAIN1.Progress("Grid Formatting")

        If EntryMode = "N" Then
            lblStatus.Text = "New"
        Else
            Select Case rowSOTCSTO1.Item("CSO_STATUS")
                Case "O"
                    lblStatus.Text = "Open"
                Case "C"
                    lblStatus.Text = "Cancelled"
                Case "D"
                    lblStatus.Text = "Deleted"
                Case "F"
                    lblStatus.Text = "Completed"
                Case Else
                    lblStatus.Text = "?"
            End Select
        End If

        If (EntryMode = "N" Or EntryMode = "E") Then ' InquiryMode uses EntryMode = "V"
            Set_Read_Only(splHeader, False)
        Else
            Set_Read_Only(splHeader, True)
        End If

        Edit_Ship_To_Addresses()

        Display_Totals()
        EnforceConstraints(True)
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_SSG()

        WorkbookView1.GetLock()

        workbook = WorkbookView1.ActiveWorkbook

        ws = workbook.Worksheets(0)
        ws.Unprotect(XLS_PWD)
        ws.Cells.Locked = False
        For i As Integer = 0 To ws.UsedRange.Rows.Count - 1
            ws.Cells(i, 0).EntireRow.Hidden = False
        Next

        ws.Cells.Clear()

        If ws.ProtectContents Then ws.Unprotect(XLS_PWD)

        ws.Cells("D1:J6").Merge()
        ws.Cells("D1").Value = "Instructions:" _
        & vbCrLf & "1) Enter Qtys in the Light Blue shaded area" _
        & vbCrLf & "2) Qtys will round up to the Order Multiple" _
        & vbCrLf & "3) Orders will be generated for RSCs w/Qtys" _
        & vbCrLf & "4) You may change CSOs until released" _
        & vbCrLf
        ws.Cells("D1").VerticalAlignment = SpreadsheetGear.VAlign.Center
        ws.Cells("D1").IndentLevel = 10

        Dim c0 As Integer = 2

        Dim r0 As Integer = r0T
        Dim r As Integer = 0
        Dim c As Integer = 0

        Dim rA As Integer = 0

        ASCMAIN1.Progress("SSG Addresses")
        Dim seenCSOKeys As New HashSet(Of String)
        For Each rowSOTCSTO3 As DataRow In dst.Tables("SOTCSTO3").Select("", "CSO_INDEX")
            Dim CSO_KEY As String = rowSOTCSTO3.Item("CSO_KEY") & ""

            If CSO_KEY <> "" Then
                If seenCSOKeys.Contains(CSO_KEY) Then
                    Continue For
                End If
                seenCSOKeys.Add(CSO_KEY)
            End If

            Dim CSO_TYPE As String = rowSOTCSTO3.Item("CSO_TYPE")
            r += 1
            ASCMAIN1.Progress("-", CStr(r))
            c = 0
            For Each CUST_ADDR_col As String In CUST_ADDR_cols
                c += 1
                If r = 1 Then

                    ws.Cells(r0 + r, c0 + c).EntireColumn.NumberFormat = "@"

                    If CUST_ADDR_col = "CUST_ADDR2" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = "Address 2"
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 10
                    ElseIf CUST_ADDR_col = "CUST_ZIP_CODE" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = "Zip Code"
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 10
                    ElseIf CUST_ADDR_col = "CSO_TYPE" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = "Type"
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 6
                        ws.Cells(r0 + r, c0 + c).EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
                    ElseIf CUST_ADDR_col = "CUST_CITY" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = "City"
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 15
                    ElseIf CUST_ADDR_col = "CUST_NAME" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = "Name"
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 20
                    ElseIf CUST_ADDR_col = "CUST_ADDR1" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = "Address"
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 20
                    ElseIf CUST_ADDR_col = "CUST_STATE" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = "State"
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 6
                    ElseIf CUST_ADDR_col = "CUST_ADDR3" Or CUST_ADDR_col = "CUST_CONTACT" Or CUST_ADDR_col = "CUST_COUNTRY" Or CUST_ADDR_col = "CUST_PHONE" Or CUST_ADDR_col = "CUST_EXT" Or CUST_ADDR_col = "CUST_FAX" Or CUST_ADDR_col = "CUST_EMAIL" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = CUST_ADDR_col
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = 20
                    Else
                        ws.Cells(r0 + r - 1, c0 + c).Value = CUST_ADDR_col
                        ws.Cells(r0 + r, c0 + c).ColumnWidth = dst.Tables("SOTCSTO3").Columns(CUST_ADDR_col).MaxLength / 2
                    End If

                    If CUST_ADDR_col = "CUST_ADDR3" Or CUST_ADDR_col = "CUST_CONTACT" Or CUST_ADDR_col = "CUST_COUNTRY" Or CUST_ADDR_col = "CUST_PHONE" Or CUST_ADDR_col = "CUST_EXT" Or CUST_ADDR_col = "CUST_FAX" Or CUST_ADDR_col = "CUST_EMAIL" Then
                        ws.Cells(r0 + r - 1, c0 + c).Value = CUST_ADDR_col
                        ws.Cells(r0 + r, c0 + c).EntireColumn.Hidden = True
                    End If
                End If

                ws.Cells(r0 + r, c0 + c).Value = rowSOTCSTO3.Item(CUST_ADDR_col) ' data from SOTCSTO3 gets loaded into XLS here
                'If r = 1 Then ws.Cells(r0 + r - 1, c0 + c).Value = rowSOTCSTO3.Item(CUST_ADDR_col)
            Next

            ws.Cells(r0 + r, c0 + c + 1).Value = rowSOTCSTO3.Item("CSO_ADDR_LNO")
            If r = 1 Then ws.Cells(r0 + r - 1, c0 + c + 1).Value = "Line"

            ws.Cells(r0 + r, c0 + c + 2).Formula = $"=SUM({Excel_Cell0(r0 + r, c0 + c + 2 + 1)}:{Excel_Cell0(r0 + r, c0 + c + 2 + dst.Tables("SOTCSTO2").Rows.Count)})"
            If r = 1 Then ws.Cells(r0 + r - 1, c0 + c + 2).Value = "Total"

            If CSO_TYPE = "AE" Then
                range = ws.Range(r0 + r, c0 + 1, r0 + r, c0 + c)
                range.Font.Color = SpreadsheetGear.Colors.Blue
                rA += 1
            End If
            If CSO_TYPE = "AC" Then
                range = ws.Range(r0 + r, c0 + 1, r0 + r, c0 + c)
                range.Font.Color = SpreadsheetGear.Colors.Red
                rA += 1
            End If
        Next

        COL_CSO_ADDR_LNO = c0 + c + 1

        Dim WHSE_CODE As String = Absx1.txtFor("WHSE_CODE").Text

        Dim rT As Integer = r

        range = ws.Range(r0, c0 + 1, r0, c0 + c)
        range.Locked = True
        range.Interior.Color = SpreadsheetGear.Colors.Blue
        range.Font.Color = SpreadsheetGear.Colors.White
        range.Font.Bold = True

        range = ws.Range(r0, c0 + c + 1, r0, c0 + c + 2)
        range.Locked = True
        range.Interior.Color = SpreadsheetGear.Colors.Orange
        'range.Font.Color = SpreadsheetGear.Colors.White
        range.Font.Bold = True

        c0_Items = COL_CSO_ADDR_LNO + 1
        c0 = c0_Items '  18
        c = 0
        r0 = 0

        ws.Cells(r0 + 0, c0 + c, r0 + 4, c0 + c).NumberFormat = "@"
        r = -1

        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Ava2Sell"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Qty OnPO"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Qty Allo"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "#Opn+Pik"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "#Ship"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "#Left"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "This CSO"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Balance"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Multiple"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Minimum"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Note"
        r += 1 : ws.Cells(r0 + r, c0 + c).Value = "Description"

        r0 = r0T

        ROW_ITEM_CODE = r0T

        ASCMAIN1.Progress("SSG Items")
        For Each row As DataRow In dst.Tables("SOTCSTO2").Select("", "CSO_LNO")

            Dim ALLO_GROUP_CODE As String = row.Item("ALLO_GROUP_CODE") & ""
            Dim HC_CODE As String = row.Item("HC_CODE") & ""
            Dim CSO_LNO As Integer = Val(row.Item("CSO_LNO") & "")
            Dim PROD_CODE As String = row.Item("PROD_CODE") & ""

            If Not ScreenMode Or (
                (dst.Tables("ICTCOLL0").Select("SEL = '1'").Length = 0 Or HC_CODEs.Contains(HC_CODE)) And
                (dst.Tables("ICTPROD1").Select("SEL = '1'").Length = 0 Or PROD_CODES.Contains(PROD_CODE)) And
                (dst.Tables("SOTRSCT1").Select("SEL = '1'").Length = 0 Or RSC_Tags.Count > 0) And
                (dst.Tables("SOTALLOG").Select("SEL = '1'").Length = 0 Or ALLO_GROUP_CODEs.Contains(ALLO_GROUP_CODE))) Then

                Dim ITEM_SO_QTY_MULT As Integer = Val(row.Item("ITEM_SO_QTY_MULT") & "")
                Dim ITEM_SO_QTY_MIN As Integer = Val(row.Item("ITEM_SO_QTY_MIN") & "")
                Dim ITEM_DESC As String = row.Item("ITEM_DESC") & ""
                Dim ITEM_CODE As String = row.Item("ITEM_CODE")

                Dim ALLO_NOTES As String = ""
                Dim rowSOTALLOXs() As DataRow = dst.Tables("SOTALLOX").Select($"ITEM_CODE = '{ITEM_CODE}'")
                If rowSOTALLOXs.Length > 0 Then
                    Dim rowSOTALLOX As DataRow = rowSOTALLOXs(0)
                    ALLO_NOTES = rowSOTALLOX.Item("ALLO_NOTES_AES") & ""
                Else
                    'If ASCMAIN1.Running_in_VS Then Stop
                End If

                c += 1
                ASCMAIN1.Progress("-", CStr(c))
                ws.Cells(r0, c0 + c).ColumnWidth = 15
                ws.Cells(r0, c0 + c).Value = ITEM_CODE

                r = -1

                Dim QTY_LEFT As Int32 = Val(row.Item("QTY_LEFT") & "")

                Dim WHSE_QTY_ON_HAND As Int32 = Val(row.Item("WHSE_QTY_ON_HAND") & "")
                Dim WHSE_QTY_ONPO As Int32 = Val(row.Item("WHSE_QTY_ONPO") & "")
                Dim WHSE_QTY_OPEN As Int32 = Val(row.Item("WHSE_QTY_OPEN") & "")
                Dim WHSE_QTY_PICK As Int32 = Val(row.Item("WHSE_QTY_PICK") & "")

                Dim AVA2SELL As Int32 = WHSE_QTY_ON_HAND - WHSE_QTY_OPEN - WHSE_QTY_PICK

                r += 1 : ws.Cells(0 + r, c0 + c).Value = AVA2SELL
                r += 1 : ws.Cells(0 + r, c0 + c).Value = WHSE_QTY_ONPO

                Dim QTY_ALLO As Int32 = Val(row.Item("QTY_ALLO") & "")

                Dim ORDR_QTY_OPEN As Int32 = Val(row.Item("ORDR_QTY_OPEN") & "")
                Dim ORDR_QTY_PICK As Int32 = Val(row.Item("ORDR_QTY_PICK") & "")
                Dim ORDR_QTY_SHIP As Int32 = Val(row.Item("ORDR_QTY_SHIP") & "")

                r += 1 : ws.Cells(0 + r, c0 + c).Value = QTY_ALLO
                r += 1 : ws.Cells(0 + r, c0 + c).Value = ORDR_QTY_OPEN + ORDR_QTY_PICK
                r += 1 : ws.Cells(0 + r, c0 + c).Value = ORDR_QTY_SHIP
                r += 1 : ws.Cells(0 + r, c0 + c).Value = QTY_ALLO - (ORDR_QTY_OPEN + ORDR_QTY_PICK + ORDR_QTY_SHIP)

                Dim balanceCell As SpreadsheetGear.IRange = ws.Cells(7, c0 + c)
                Dim negativeCondition As SpreadsheetGear.IFormatCondition = balanceCell.FormatConditions.Add(
            SpreadsheetGear.FormatConditionType.CellValue,
            SpreadsheetGear.FormatConditionOperator.Less,
            0, Nothing)
                With negativeCondition
                    .Interior.Color = SpreadsheetGear.Colors.DarkRed
                    .Font.Color = SpreadsheetGear.Colors.White
                    .Font.Bold = True
                End With

                r += 1 : ws.Cells(0 + r, c0 + c).Formula = $"=SUM({Excel_Cell0(r0 + 1, c0 + c)}:{Excel_Cell0(r0 + rT, c0 + c)})"
                r += 1 : ws.Cells(0 + r, c0 + c).Formula = $"={Excel_Cell0(0 + r - 2, c0 + c)} - {Excel_Cell0(0 + r - 1, c0 + c)}"

                r += 1 : ws.Cells(0 + r, c0 + c).Value = ITEM_SO_QTY_MULT
                r += 1 : ws.Cells(0 + r, c0 + c).Value = ITEM_SO_QTY_MIN
                r += 1 : ws.Cells(0 + r, c0 + c).Value = ALLO_NOTES
                r += 1 : ws.Cells(0 + r, c0 + c).Value = ITEM_DESC


                Dim condition As SpreadsheetGear.IFormatCondition = ws.Cells(0 + r, c0 + c).FormatConditions.Add(
             SpreadsheetGear.FormatConditionType.CellValue,
                SpreadsheetGear.FormatConditionOperator.Less,
             0, Nothing)
                condition.Interior.Color = SpreadsheetGear.Colors.Red
                condition.Font.Color = SpreadsheetGear.Colors.White
                condition.Font.Bold = True

                For Each rowSOTCSTO4 As DataRow In dst.Tables("SOTCSTO4").Select($"CSO_LNO = {CStr(CSO_LNO)}")
                    Dim CSO_ADDR_LNO As Integer = Val(rowSOTCSTO4.Item("CSO_ADDR_LNO") & "")
                    Dim CSO_QTY As Integer = Val(rowSOTCSTO4.Item("CSO_QTY") & "")
                    If CSO_QTY <> 0 Then
                        ws.Cells(ROW_ITEM_CODE + CSO_ADDR_LNO, c0 + c).Value = CSO_QTY
                    End If
                Next
                With ws.Cells(0 + r, c0 + c)
                    'Dim QTY_LEFT As Integer = Val(.Value)
                    If QTY_LEFT <= 0 Then
                        .Interior.Color = SpreadsheetGear.Colors.LightGray
                        .EntireColumn.Locked = True
                    End If
                End With
            Else
            End If
        Next

        NO_ITEMS = False
        If c > 0 Then
            range = ws.Range(r0 + 0, c0 + 1, r0 + 3, c0 + c) ' Item Totals Section - if there are no valid items, c = 0 and this will blow up
        Else
            NO_ITEMS = True
            WorkbookView1.ReleaseLock()
            ASCMAIN1.Progress("")
            Exit Sub
        End If

        range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.InsideVertical).Color = SpreadsheetGear.Colors.LightGray

        range = ws.Range(0 + 0, c0 + 1, 0 + 0, c0 + c)
        range.Interior.Color = SpreadsheetGear.Colors.LightPink
        range = ws.Range(0 + 1, c0 + 1, 0 + 1, c0 + c)
        range.Interior.Color = SpreadsheetGear.Colors.LightPink

        range = ws.Range(0 + 6, c0 + 1, 0 + 6, c0 + c)
        range.Interior.Color = SpreadsheetGear.Colors.Beige

        'range = ws.Range(r0, c0 + 1, r0 + rT, c0 + c) ' Items Section
        'range.Locked = False
        'range.Interior.Color = SpreadsheetGear.Colors.AliceBlue
        range = ws.Range(r0, c0 + 1, r0 + rT, c0 + c) ' Items Section

        'range.Interior.Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeLeft).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeRight).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeTop).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.EdgeBottom).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.InsideHorizontal).Color = SpreadsheetGear.Colors.LightGray
        range.Borders(SpreadsheetGear.BordersIndex.InsideVertical).LineStyle = SpreadsheetGear.LineStyle.Continuous
        range.Borders(SpreadsheetGear.BordersIndex.InsideVertical).Color = SpreadsheetGear.Colors.LightGray

        range = ws.Range(r0, c0 + 1, r0, c0 + c)
        'range.Locked = True
        range.Interior.Color = SpreadsheetGear.Colors.LightGreen
        'range.Font.Color = SpreadsheetGear.Colors.White
        range.Font.Bold = True

        For CX As Integer = 1 To c
            Dim QTY_LEFT As Integer = Val(ws.Cells(ROW_ITEM_CODE - 3, c0 + CX).Value & "")
            range = ws.Range(r0 + 1, c0 + CX, r0 + rT, c0 + CX) ' Items Section
            If QTY_LEFT <= 0 Then
                range.Locked = True
                range.Interior.Color = SpreadsheetGear.Colors.LightGray
            Else
                range.Locked = False
                range.Interior.Color = SpreadsheetGear.Colors.AliceBlue
            End If

            Dim ITEM_CODE As String = ws.Cells(ROW_ITEM_CODE, c0 + CX).Value.ToString().Trim()
            If READ_ONLY.Contains(ITEM_CODE) Then
                ws.Cells(ROW_ITEM_CODE, c0 + CX).EntireColumn.Locked = True
                ws.Cells(ROW_ITEM_CODE, c0 + CX).Interior.Color = SpreadsheetGear.Colors.PaleVioletRed
                ws.Cells(ROW_ITEM_CODE, c0 + CX).Font.Color = SpreadsheetGear.Colors.White
                ws.Cells(ROW_ITEM_CODE, c0 + CX).AddComment("The end-date for this item is within 2 weeks of the CSO date.")
            End If
        Next


        ws.Cells(r0 + 1 + rA, c0 + 1).Activate()
        ws.WindowInfo.FreezePanes = True

        ws.Range(0, 0).EntireColumn.Hidden = True
        ws.Range(0, 1).EntireColumn.Hidden = True
        ws.Range(0, 2).EntireColumn.Hidden = True
        ws.Range(8, 0).EntireRow.Hidden = True
        ws.Range(9, 0).EntireRow.Hidden = True
        ws.Range(10, 0).EntireRow.Hidden = True
        ws.Range(11, 0).EntireRow.Hidden = True

        Show_Item(c0 + 1, True)

        'ws.Cells.Locked = True
        ws.Protect(XLS_PWD)


        X = New MyCommandManager(WorkbookView1.ActiveWorkbookSet, WorkbookView1, isClearing, isPasting, Me)

        WorkbookView1.ReleaseLock()

        ASCMAIN1.Progress("")
    End Sub

    Sub Delete_Record()
        If EntryMode <> "E" Then Exit Sub
        BeginTrans()
        Delete_Records()
        CommitTrans("Delete Complete")
    End Sub

    Sub Delete_Records()
        If EntryMode = "N" Then Exit Sub
        Dependent_Updates(-1, CSO_NO)
        For Each TABLE_NAME As String In New String() _
            {"SOTCSTO1", "SOTCSTO2"}
            Delete_Records_1(TABLE_NAME)
        Next
    End Sub

    Sub Delete_Records_1(TABLE_NAME As String)
        ASCMAIN1.sql = "Delete from " & TABLE_NAME & " where CSO_NO = '" & CSO_NO & "'"
        ASCDATA1.ExecuteSQL()
    End Sub

    Sub Update_Record()
        Try
            BeginTrans()
            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Updating ...")

            dst.Tables("SOTORDR1").Rows.Clear()
            dst.Tables("SOTORDR2").Rows.Clear()
            dst.Tables("SOTORDR5").Rows.Clear()

            Dim ORDR_GROUP_NO_split As String = ""

            Fill_Records("ICTWHSEX", {rowSOTCSTO1.Item("WHSE_CODE")})

            If EntryMode <> "N" Then Delete_Records()

            dst.Tables("SOTCSTO4").Rows.Clear()

            Dim ORDR_GROUP_NO As String = ""

            If EntryMode = "N" Then
                Write_Event_Log("SOTCSTO1", CSO_NO, "Car-Stock Order Created")

                ORDR_GROUP_NO = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                rowSOTCSTO1.Item("ORDR_GROUP_NO") = ORDR_GROUP_NO

            Else

                Write_Event_Log("SOTCSTO1", CSO_NO, "Car-Stock Order Modified")

                For Each rowSOTCSTO3 As DataRow In dst.Tables("SOTCSTO3").Select("ORDR_NO IS NOT NULL")
                    Dim ORDR_NO As String = rowSOTCSTO3.Item("ORDR_NO")
                    Dependent_Updates_SOTORDR1(-1, ORDR_NO)
                Next

                ORDR_GROUP_NO = rowSOTCSTO1.Item("ORDR_GROUP_NO")

            End If

            For Each rowSOTCSTO3 As DataRow In dst.Tables("SOTCSTO3").Select("ORDR_NO IS NOT NULL OR CSO_QTY_TOTAL <> 0", "CSO_ADDR_LNO")
                Write_SOTORDRx(ORDR_GROUP_NO, rowSOTCSTO3)
            Next

            Dim lstCsoKeys As New List(Of String)
            For Each rowSOTCSTO3 As DataRow In dst.Tables("SOTCSTO3").Select("", "CSO_ADDR_LNO")
                Dim CSO_KEY As String = rowSOTCSTO3.Item("CSO_KEY") & String.Empty
                Dim ORDR_NO As String = rowSOTCSTO3.Item("ORDR_NO") & String.Empty
                If ORDR_NO.Length = 0 Then
                    Continue For
                End If

                If lstCsoKeys.Contains(CSO_KEY) Then
                    Continue For
                End If

                lstCsoKeys.Add(CSO_KEY)

                Dim numOrders As Int16 = dst.Tables("SOTCSTO3").Select($"ISNULL(ORDR_NO, '') <> '' and CSO_KEY = '{CSO_KEY}'").Length
                Select Case numOrders
                    Case 0
                        Continue For
                    Case 1
                        SplitOrdersOnItemWarehouse(ORDR_NO)
                    Case Else
                        MergeOrdersOnItemWarehouse(CSO_KEY)
                End Select
            Next

            For Each drSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
                Dim ORDR_NO As String = drSOTORDR1.Item("ORDR_NO") & String.Empty
                If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND ORDR_STATUS = 'O'").Length > 0 Then
                    dst.Tables("SOTORDR1").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("ORDR_STATUS") = "O"
                Else
                    dst.Tables("SOTORDR1").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("ORDR_STATUS") = "C"
                End If
            Next

            Record_Audits()

            Update_Record_TDA("SOTORDR1")
            Update_Record_TDA("SOTORDR2")
            Update_Record_TDA("SOTORDR5")

            For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select()
                Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
                Dependent_Updates_SOTORDR1(1, ORDR_NO)
            Next

            Dim tbl_ORDR_GROUP_NOs As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), {"ORDR_GROUP_NO"})
            For Each rowORDR_GROUP_NO As DataRow In tbl_ORDR_GROUP_NOs.Select
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {rowORDR_GROUP_NO.Item("ORDR_GROUP_NO")}, New String() {"ORDR_GROUP_NO_IN"})
            Next

            dst.Tables("SOTORDXR").Rows.Clear()

            INIT_LAST("SOTCSTO1", False, , True)
            Dim sqldelete As String = "CSO_NO = '" & CSO_NO & "'"
            Update_Record_TDA("SOTCSTO1", sqldelete)
            Update_Record_TDA("SOTCSTO2", sqldelete)
            Update_Record_TDA("SOTCSTO3", sqldelete)
            Update_Record_TDA("SOTCSTO4", sqldelete)

            Dependent_Updates(1, CSO_NO)

            Dim msg As String = "Update Complete"
            CommitTrans(msg)

        Catch ex As Exception
            Rollback(ex.Message)
            Exit Sub
        Finally
            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")
        End Try

        ASCMAIN1.Progress("Now Deleting Saved DataSets")
        For Each file As String In My.Computer.FileSystem.GetFiles(ASCMAIN1.Folders("Work"))

            If file.StartsWith(ASCMAIN1.Folders("Work") & Me.Name) And file.EndsWith(".dst.xml") Then
                Dim delete_file As Boolean = False

                If file.StartsWith(ASCMAIN1.Folders("Work") & Me.Name & $"_{CSO_NO}") And file.EndsWith(".dst.xml") Then
                    delete_file = True
                Else
                    Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(file)
                    If Now.Subtract(fi.CreationTime).Days > 7 Then
                        delete_file = True
                    End If
                End If

                If delete_file Then
                    Try
                        My.Computer.FileSystem.DeleteFile(file)
                    Catch ex As Exception

                    End Try

                End If
            End If
        Next
        ASCMAIN1.Progress("")

    End Sub

    ''' <summary>
    ''' if an order header says WHSE X, but some of its line items point to WHSE Y, split that one order into multiple
    ''' </summary>
    ''' <param name="ORDR_NO"></param>
    Private Sub SplitOrdersOnItemWarehouse(ByVal ORDR_NO As String)

        Dim DICT_ORDR_GROUP_NOs As New Dictionary(Of String, String)
        DICT_ORDR_GROUP_NOs.Add(rowSOTCSTO1.Item("WHSE_CODE"), rowSOTCSTO1.Item("ORDR_GROUP_NO"))

        Dim tblDistinct As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR1"), {"WHSE_CODE", "ORDR_GROUP_NO"})
        For Each row As DataRow In tblDistinct.Select()
            If Not DICT_ORDR_GROUP_NOs.Keys.Contains(row.Item("WHSE_CODE")) Then
                DICT_ORDR_GROUP_NOs.Add(row.Item("WHSE_CODE"), row.Item("ORDR_GROUP_NO"))
            End If
        Next

        Dim rowSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
        Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

        If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE <> '{WHSE_CODE}'").Length = 0 Then
            Exit Sub
        End If

        Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Select($"ORDR_NO = '{ORDR_NO}'")(0)
        Dim TBL_WHSES As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE <> '{WHSE_CODE}'"), "WHSE_CODE")
        For Each rowWhses As DataRow In TBL_WHSES.Rows
            Dim WHSE_CODE_DTL As String = rowWhses.Item("WHSE_CODE")
            If Not DICT_ORDR_GROUP_NOs.ContainsKey(WHSE_CODE_DTL) Then
                DICT_ORDR_GROUP_NOs.Add(WHSE_CODE_DTL, ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO"))
            End If
            Dim rowSOTORDR1_new As DataRow = dst.Tables("SOTORDR1").NewRow
            rowSOTORDR1_new.ItemArray = rowSOTORDR1.ItemArray
            Dim ORDR_NO_NEW As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
            rowSOTORDR1_new.Item("ORDR_NO") = ORDR_NO_NEW
            rowSOTORDR1_new.Item("WHSE_CODE") = WHSE_CODE_DTL
            rowSOTORDR1_new.Item("ORDR_GROUP_NO") = DICT_ORDR_GROUP_NOs(WHSE_CODE_DTL)

            Dim SHIP_VIA_CODE_current As String = rowSOTORDR1_new.Item("SHIP_VIA_CODE") & ""
            rowSOTORDR1_new.Item("SHIP_VIA_CODE") = GetShipViaForWhse(WHSE_CODE_DTL, SHIP_VIA_CODE_current)
            dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1_new)

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE = '{WHSE_CODE_DTL}'")
                Dim ROWSOTORDR2_new As DataRow = dst.Tables("SOTORDR2").NewRow
                ROWSOTORDR2_new.ItemArray = rowSOTORDR2.ItemArray
                ROWSOTORDR2_new.Item("ORDR_NO") = ORDR_NO_NEW
                dst.Tables("SOTORDR2").Rows.Add(ROWSOTORDR2_new)
                rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                rowSOTORDR2.Item("ORDR_QTY_CANC") = rowSOTORDR2.Item("ORDR_QTY_ORIG")
                rowSOTORDR2.Item("ORDR_QTY") = 0
                rowSOTORDR2.Item("ORDR_STATUS") = "C"
            Next
            For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")
                Dim rowSOTORDR5_new As DataRow = dst.Tables("SOTORDR5").NewRow
                rowSOTORDR5_new.ItemArray = rowSOTORDR5.ItemArray
                rowSOTORDR5_new.Item("ORDR_NO") = ORDR_NO_NEW
                dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5_new)
            Next
            Dim rowSOTCSTO3_new As DataRow = dst.Tables("SOTCSTO3").NewRow
            rowSOTCSTO3_new.ItemArray = rowSOTCSTO3.ItemArray
            rowSOTCSTO3_new.Item("ORDR_NO") = ORDR_NO_NEW
            rowSOTCSTO3_new.Item("CSO_ADDR_LNO") = Val(dst.Tables("SOTCSTO3").Compute("MAX(CSO_ADDR_LNO)", "")) + 1
            dst.Tables("SOTCSTO3").Rows.Add(rowSOTCSTO3_new)
        Next
    End Sub
    Private Shared Function GetShipViaForWhse(whse As String, currentShipVia As String) As String
        whse = (whse & "").Trim().ToUpperInvariant()

        If whse = "CLA" Then Return "S32"
        If whse Like "ADS*" Then Return "FXH"

        Return currentShipVia
    End Function


    ''' <summary>
    ''' when there are multiple orders for the same store, try to consolidate lines so each warehouse's items live on the whse order
    ''' </summary>
    ''' <param name="CSO_KEY"></param>
    ''' 
    Private Sub MergeOrdersOnItemWarehouse(ByVal CSO_KEY As String)

        'Dim ORDR_GROUP_NO As String = rowSOTCSTO1.Item("ORDR_GROUP_NO")

        ' Loop through all Orders to determine if the Items are in the warehouse on SOTORDR1
        'Dim lstOrderNumbers As New List(Of String)
        'For Each dr As DataRow In dst.Tables("SOTCSTO3").Select($"ORDR_NO IS NOT NULL and CSO_KEY = '{CSO_KEY}'")
        '    lstOrderNumbers.Add(dr.Item("ORDR_NO"))
        'Next
        'Dim ORDR_NO As String = lstOrderNumbers(0)
        'Dim CUST_STORE_NO As String = dst.Tables("SOTORDR1").Select($"ORDR_NO = '{ORDR_NO}'")(0).Item("CUST_STORE_NO") & String.Empty

        'Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Select($"ORDR_NO = '{msORDR_NO}'")(0)
        'Dim msSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Select($"ORDR_GROUP_NO = '{ORDR_GROUP_NO}' AND CUST_STORE_NO = '{CUST_STORE_NO}'")(0)
        'Dim msWHSE_CODE As String = msSOTORDR1.Item("WHSE_CODE")
        'Dim msORDR_NO As String = msSOTORDR1.Item("ORDR_NO")

        'ORDR_NO = String.Empty

        For Each drSOTCSTO3 As DataRow In dst.Tables("SOTCSTO3").Select($"CSO_KEY = '{CSO_KEY}'")
            Dim ORDR_NO As String = drSOTCSTO3.Item("ORDR_NO") & String.Empty
            If ORDR_NO.Length = 0 Then
                Continue For
            End If

            Dim msSOTORDR1 As DataRow = dst.Tables("SOTORDR1").Rows.Find(ORDR_NO)
            Dim msWHSE_CODE As String = msSOTORDR1.Item("WHSE_CODE")
            Dim msORDR_NO As String = msSOTORDR1.Item("ORDR_NO")
            Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Select($"ORDR_NO = '{msORDR_NO}'")(0)

            For Each msSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{msORDR_NO}'")
                Dim WHSE_CODE As String = msSOTORDR2.Item("WHSE_CODE")
                If WHSE_CODE = msWHSE_CODE Then
                    Continue For
                End If

                Dim ITEM_CODE As String = msSOTORDR2.Item("ITEM_CODE")
                If dst.Tables("SOTORDR1").Select($"WHSE_CODE = '{WHSE_CODE}' AND CSO_KEY = '{CSO_KEY}'").Length > 0 Then
                    ORDR_NO = dst.Tables("SOTORDR1").Select($"WHSE_CODE = '{WHSE_CODE}' AND CSO_KEY = '{CSO_KEY}'")(0).Item("ORDR_NO")
                    ' Found an order for this warehouse
                    ' See if the item exists on the sales order
                    If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND ITEM_CODE = '{ITEM_CODE}'").Length > 0 Then
                        ' Found a detail line for this item
                        Dim drDetail As DataRow = dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND ITEM_CODE = '{ITEM_CODE}'")(0)
                        drDetail.Item("ORDR_QTY") = msSOTORDR2.Item("ORDR_QTY")
                        drDetail.Item("ORDR_QTY_OPEN") = msSOTORDR2.Item("ORDR_QTY_OPEN")
                        If msSOTORDR2.Item("ORDR_STATUS") = "O" Then
                            drDetail.Item("ORDR_QTY_CANC") = 0
                            drDetail.Item("ORDR_STATUS") = "O"
                        End If
                    ElseIf msSOTORDR2.Item("ORDR_STATUS") = "O" Then
                        ' Need to add the item to the order
                        Dim drDetail As DataRow = dst.Tables("SOTORDR2").NewRow
                        drDetail.ItemArray = msSOTORDR2.ItemArray
                        drDetail.Item("ORDR_NO") = ORDR_NO
                        drDetail.Item("ORDR_QTY") = msSOTORDR2.Item("ORDR_QTY")
                        drDetail.Item("ORDR_QTY_OPEN") = msSOTORDR2.Item("ORDR_QTY_OPEN")
                        drDetail.Item("ORDR_QTY_CANC") = 0
                        drDetail.Item("ORDR_STATUS") = "O"
                        dst.Tables("SOTORDR2").Rows.Add(drDetail)
                    End If

                    msSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    msSOTORDR2.Item("ORDR_QTY_CANC") = msSOTORDR2.Item("ORDR_QTY_ORIG")
                    msSOTORDR2.Item("ORDR_QTY") = 0
                    msSOTORDR2.Item("ORDR_STATUS") = "C"
                Else
                    ' Need to make an order for this warehouse
                    Dim rowSOTORDR1_new As DataRow = dst.Tables("SOTORDR1").NewRow
                    Dim ORDR_NO_NEW As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")

                    rowSOTORDR1_new.ItemArray = msSOTORDR1.ItemArray
                    rowSOTORDR1_new.Item("ORDR_NO") = ORDR_NO_NEW
                    rowSOTORDR1_new.Item("WHSE_CODE") = WHSE_CODE
                    rowSOTORDR1_new.Item("ORDR_GROUP_NO") = ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO")
                    dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1_new)

                    For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")
                        Dim rowSOTORDR5_new As DataRow = dst.Tables("SOTORDR5").NewRow
                        rowSOTORDR5_new.ItemArray = rowSOTORDR5.ItemArray
                        rowSOTORDR5_new.Item("ORDR_NO") = ORDR_NO_NEW
                        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5_new)
                    Next

                    Dim rowSOTCSTO3_new As DataRow = dst.Tables("SOTCSTO3").NewRow
                    rowSOTCSTO3_new.ItemArray = rowSOTCSTO3.ItemArray
                    rowSOTCSTO3_new.Item("ORDR_NO") = ORDR_NO_NEW
                    rowSOTCSTO3_new.Item("CSO_ADDR_LNO") = Val(dst.Tables("SOTCSTO3").Compute("MAX(CSO_ADDR_LNO)", "")) + 1
                    dst.Tables("SOTCSTO3").Rows.Add(rowSOTCSTO3_new)

                    Dim drDetail As DataRow = dst.Tables("SOTORDR2").NewRow
                    drDetail.ItemArray = msSOTORDR2.ItemArray
                    drDetail.Item("ORDR_NO") = ORDR_NO_NEW
                    drDetail.Item("ORDR_QTY") = msSOTORDR2.Item("ORDR_QTY")
                    drDetail.Item("ORDR_QTY_OPEN") = msSOTORDR2.Item("ORDR_QTY_OPEN")
                    dst.Tables("SOTORDR2").Rows.Add(drDetail)

                    msSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    msSOTORDR2.Item("ORDR_QTY_CANC") = msSOTORDR2.Item("ORDR_QTY_ORIG")
                    msSOTORDR2.Item("ORDR_QTY") = 0
                    msSOTORDR2.Item("ORDR_STATUS") = "C"
                End If
            Next
        Next

    End Sub

    Sub Evaluate_Orders_Edit()

        Dim DICT_ORDR_GROUP_NOs As New Dictionary(Of String, String)
        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")
            If dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE <> '{WHSE_CODE}'").Length = 0 Then
                Continue For
            End If
            Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Select($"ORDR_NO = '{ORDR_NO}'")(0)
            Dim TBL_WHSES As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE <> '{WHSE_CODE}'"), "WHSE_CODE")
            For Each rowWhses As DataRow In TBL_WHSES.Rows
                Dim WHSE_CODE_DTL As String = rowWhses.Item("WHSE_CODE")
                If Not DICT_ORDR_GROUP_NOs.ContainsKey(WHSE_CODE_DTL) Then
                    DICT_ORDR_GROUP_NOs.Add(WHSE_CODE_DTL, ASCMAIN1.Next_Control_No("SOTORDR0.ORDR_GROUP_NO"))
                End If
                Dim rowSOTORDR1_new As DataRow = dst.Tables("SOTORDR1").NewRow
                rowSOTORDR1_new.ItemArray = rowSOTORDR1.ItemArray
                Dim ORDR_NO_NEW As String = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                rowSOTORDR1_new.Item("ORDR_NO") = ORDR_NO_NEW
                rowSOTORDR1_new.Item("WHSE_CODE") = WHSE_CODE_DTL
                rowSOTORDR1_new.Item("ORDR_GROUP_NO") = DICT_ORDR_GROUP_NOs(WHSE_CODE_DTL)
                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1_new)

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}' AND WHSE_CODE = '{WHSE_CODE_DTL}'")
                    Dim ROWSOTORDR2_new As DataRow = dst.Tables("SOTORDR2").NewRow
                    ROWSOTORDR2_new.ItemArray = rowSOTORDR2.ItemArray
                    ROWSOTORDR2_new.Item("ORDR_NO") = ORDR_NO_NEW
                    dst.Tables("SOTORDR2").Rows.Add(ROWSOTORDR2_new)
                    rowSOTORDR2.Item("ORDR_QTY_OPEN") = 0
                    rowSOTORDR2.Item("ORDR_QTY") = 0
                Next
                For Each rowSOTORDR5 As DataRow In dst.Tables("SOTORDR5").Select($"ORDR_NO = '{ORDR_NO}'")
                    Dim rowSOTORDR5_new As DataRow = dst.Tables("SOTORDR5").NewRow
                    rowSOTORDR5_new.ItemArray = rowSOTORDR5.ItemArray
                    rowSOTORDR5_new.Item("ORDR_NO") = ORDR_NO_NEW
                    dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5_new)
                Next
                Dim rowSOTCSTO3_new As DataRow = dst.Tables("SOTCSTO3").NewRow
                rowSOTCSTO3_new.ItemArray = rowSOTCSTO3.ItemArray
                rowSOTCSTO3_new.Item("ORDR_NO") = ORDR_NO_NEW
                rowSOTCSTO3_new.Item("CSO_ADDR_LNO") = Val(dst.Tables("SOTCSTO3").Compute("MAX(CSO_ADDR_LNO)", "")) + 1
                dst.Tables("SOTCSTO3").Rows.Add(rowSOTCSTO3_new)
            Next
        Next
    End Sub

    Sub Record_Audits()

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("")
            Dim ORDR_NO As String = rowSOTORDR1.Item("ORDR_NO")
            Dim REV_LNO As Integer = 0
            If rowSOTORDR1.RowState = DataRowState.Added Then
                TAC.SOCMAIN1.Record_Event_SOTORDR1(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "New", "Car-Stock Order")
            Else
                TAC.SOCMAIN1.Record_Event_SOTORDR1(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "Edit", "Car-Stock Order")


                ASCMAIN1.sql = "Select Max (REV_NO) From SOTORDXR Where ORDR_NO = '" & ORDR_NO & "'"
                Dim REV_NO As Integer = Val(ASCDATA1.GetDataValue & "") + 1


                For Each COLUMN_NAME As String In New String() {"ORDR_STATUS", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "SHIP_VIA_CODE"}
                    '{"ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_ARRIVAL_DATE", "ORDR_ALLO_DATE", "ORDR_MESSAGE",
                    ' "ORDR_DEPT", "ORDR_OVERRIDE_NOT_ALLOCATED", "ORDR_HOLD", "ORDR_HOLD_REASON",
                    ' "ORDR_SPECIAL_INST", "ORDR_INV_COMMENT", "ORDR_INTERNAL_NOTES",
                    ' "SREP_CODE", "SREP2_CODE", "TERM_CODE", "WHSE_CODE",
                    ' "SHIP_VIA_CODE", "FRT_TERMS", "REASON_CODE", "ORDR_ADDR_TYPE_ST"}

                    If rowSOTORDR1.Item(COLUMN_NAME) & "" <> rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original) & "" Then
                        Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                        With rowSOTORDXR
                            .Item("REV_NO") = REV_NO
                            REV_LNO += 1
                            .Item("REV_LNO") = REV_LNO
                            .Item("ORDR_NO") = ORDR_NO
                            .Item("ORDR_LNO") = 0
                            .Item("INIT_DATE") = DATETIME_STAMP
                            .Item("INIT_OPER") = ASCMAIN1.USER_ID
                            .Item("COLUMN_NAME") = COLUMN_NAME
                            .Item("OLD_VALUE") = rowSOTORDR1.Item(COLUMN_NAME, DataRowVersion.Original)
                            .Item("NEW_VALUE") = rowSOTORDR1.Item(COLUMN_NAME)
                            .Item("EMODE") = EntryMode
                        End With
                        dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                    End If
                Next

                For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}'")
                    If rowSOTORDR2.RowState = DataRowState.Added Then

                    Else
                        Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
                        Dim ORDR_LNO As Int32 = Val(rowSOTORDR2.Item("ORDR_LNO") & "")
                        Dim ORDR_QTY_OPEN As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                        Dim ORDR_QTY_OPEN_ORIG As Int64 = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN", DataRowVersion.Original) & "")
                        If ORDR_QTY_OPEN_ORIG <> ORDR_QTY_OPEN Then
                            Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow
                            With rowSOTORDXR
                                .Item("REV_NO") = REV_NO
                                REV_LNO += 1
                                .Item("REV_LNO") = REV_LNO
                                .Item("ORDR_NO") = ORDR_NO
                                .Item("ORDR_LNO") = ORDR_LNO
                                .Item("INIT_DATE") = DATETIME_STAMP
                                .Item("INIT_OPER") = ASCMAIN1.USER_ID
                                .Item("COLUMN_NAME") = "ORDR_QTY_OPEN"
                                .Item("OLD_VALUE") = ORDR_QTY_OPEN_ORIG
                                .Item("NEW_VALUE") = ORDR_QTY_OPEN
                                .Item("EMODE") = EntryMode
                                '  .Item("CONTEXT") = rowSOTORDR2_orig.Item("ITEM_CODE")
                            End With
                            dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                        End If
                    End If
                Next
            End If
        Next

        'For Each modifiedRow As DataRow In modifiedAddresses
        '    Dim ORDR_NO As String = modifiedRow("ORDR_NO").ToString()

        '    ASCMAIN1.sql = $"Select Max (REV_NO) From SOTORDXR Where ORDR_NO = '{ORDR_NO}'"
        '    Dim REV_NO As Integer = Val(ASCDATA1.GetDataValue & "") + 1

        '    Dim originalRow As DataRow = originalValues(modifiedRow)

        '    For Each ADDRESS_FIELD As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE"}
        '        Dim originalValue As String = originalRow(ADDRESS_FIELD).ToString()
        '        Dim newValue As String = modifiedRow(ADDRESS_FIELD).ToString()

        '        If originalValue <> newValue Then
        '            Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow()
        '            With rowSOTORDXR
        '                .Item("REV_NO") = REV_NO
        '                REV_LNO += 1
        '                .Item("REV_LNO") = REV_LNO
        '                .Item("ORDR_NO") = ORDR_NO
        '                .Item("ORDR_LNO") = 0
        '                .Item("INIT_DATE") = DATETIME_STAMP
        '                .Item("INIT_OPER") = ASCMAIN1.USER_ID
        '                .Item("COLUMN_NAME") = ADDRESS_FIELD
        '                .Item("OLD_VALUE") = originalValue
        '                .Item("NEW_VALUE") = newValue
        '                .Item("EMODE") = EntryMode
        '            End With
        '            dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
        '        End If
        '    Next
        'Next

        For Each modifiedRow As DataRow In modifiedAddresses
            Dim ORDR_NO As String = modifiedRow("ORDR_NO").ToString()

            Dim existing() As DataRow = dst.Tables("SOTORDXR").Select($"ORDR_NO = '{ORDR_NO}'")
            Dim REV_NO As Integer
            Dim REV_LNO As Integer

            If existing.Length > 0 Then
                Dim MAX_REV_NO As Integer = -1
                For Each row As DataRow In existing
                    Dim CUR_REV_NO As Integer = CInt(row.Item("REV_NO"))
                    If CUR_REV_NO > MAX_REV_NO Then
                        MAX_REV_NO = CUR_REV_NO
                    End If
                Next
                REV_NO = MAX_REV_NO

                Dim MAX_REV_LNO As Integer = -1
                For Each row As DataRow In existing
                    If row.Item("REV_NO") = REV_NO Then
                        Dim CUR_REV_LNO As Integer = CInt(row.Item("REV_LNO"))
                        If CUR_REV_LNO > MAX_REV_LNO Then
                            MAX_REV_LNO = CUR_REV_LNO
                        End If
                    End If
                Next
                REV_LNO = MAX_REV_LNO
            Else
                ASCMAIN1.sql = $"Select Max (REV_NO) From SOTORDXR Where ORDR_NO = '{ORDR_NO}'"
                REV_NO = Val(ASCDATA1.GetDataValue & "") + 1
                REV_LNO = 0
            End If

            Dim originalRow As DataRow = originalValues(modifiedRow)

            For Each ADDRESS_FIELD As String In New String() {"CUST_ADDR1", "CUST_ADDR2", "CUST_CITY", "CUST_STATE", "CUST_ZIP_CODE"}
                Dim originalValue As String = originalRow(ADDRESS_FIELD).ToString()
                Dim newValue As String = modifiedRow(ADDRESS_FIELD).ToString()

                If originalValue <> newValue Then
                    Dim rowSOTORDXR As DataRow = dst.Tables("SOTORDXR").NewRow()
                    With rowSOTORDXR
                        .Item("REV_NO") = REV_NO
                        REV_LNO += 1
                        .Item("REV_LNO") = REV_LNO
                        .Item("ORDR_NO") = ORDR_NO
                        .Item("ORDR_LNO") = 0
                        .Item("INIT_DATE") = DATETIME_STAMP
                        .Item("INIT_OPER") = ASCMAIN1.USER_ID
                        .Item("COLUMN_NAME") = ADDRESS_FIELD
                        .Item("OLD_VALUE") = originalValue
                        .Item("NEW_VALUE") = newValue
                        .Item("EMODE") = EntryMode
                    End With
                    dst.Tables("SOTORDXR").Rows.Add(rowSOTORDXR)
                End If
            Next
        Next

        modifiedAddresses.Clear()
        originalValues.Clear()
        validatedAddresses.Clear()
        Update_Record_TDA("SOTORDXR")
    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special _
        (ByVal ctl As Windows.Forms.Control,
         ByVal COLUMN_NAME As String,
         Optional ByRef sql_where As String = "",
         Optional ByRef Cancel As Boolean = False)

        Select Case COLUMN_NAME

            Case "CSO_NO"

                If Absx1.txtFor("SELL_CODE").Text = "" Then
                    MsgBox("You must enter an AE Code", vbOKOnly, "Cannot Perform Requested Action")
                    Cancel = True
                    Exit Sub
                End If
                sql_where = ""

                If InquiryMode Then
                Else
                    sql_where &= " and SOTCSTO1.CSO_STATUS = 'O' "
                End If

                If Absx1.txtFor("SELL_CODE").Text <> "" Then
                    sql_where &= " and SOTCSTO1.SELL_CODE = '" & Absx1.txtFor("SELL_CODE").Text & "'"
                End If
                If Absx1.txtFor("CSO_REF_NO").Text <> "" Then
                    sql_where &= " and SOTCSTO1.CSO_REF_NO = '" & Absx1.txtFor("CSO_REF_NO").Text & "'"
                End If

        End Select
    End Sub

    Public Overrides Function Remote_Control(
    ByVal command As String,
    Optional ByVal key As String = "") As Object

        Dim return_key As Object = Nothing
        Application.DoEvents()

        Select Case command
            Case "Done"
                Click_Command("Done")

            Case "View", "Edit"
                Absx1.txtFor("CSO_NO").Text = key
                Click_Command("View")
        End Select

        Return return_key
    End Function

    Public Overrides Function Dropped_On_Context() As Dropped_On_Entity

        Dim E As New Dropped_On_Entity
        If ScreenMode Then
            E.TABLE_NAME = "SOTCSTO1"
            E.COLUMN_NAME = "CSO_NO"
            E.CODE_VALUE = Absx1.txtFor("CSO_NO").Text
            E.DESC_VALUE = "Car Stock Order"
            E.ATTACHMENT_NOTES = ""
        End If

        Return E
    End Function

#End Region

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTCSTOX, "SSSB", "Show Filter", "Show GroupBox", "Show Pins", "Cancel Order")
        Load_Popup_Menu(grdSOTCSTOI, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
        Load_Popup_Menu(grdSOTCSTTX, "SS", "Show Filter", "Show GroupBox")
        Load_Popup_Menu(grdSOTALLOD, "SSB", "Show Filter", "Show GroupBox", "Item Status Inquiry")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = Nothing
        If GRDs.ContainsKey(Mid(e.SourceControl.Name, 4)) Then grd = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case grd.Name

            Case "grdSOTCSTOX"

                tlb_btn = DirectCast(tlb_pop.Tools("Cancel Order"), UltraWinToolbars.ButtonTool)
                ' do not mt lock here

                If grd.Selected.Rows.Count = 0 Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    tlb_btn.SharedProps.Visible = True
                    Dim b As Integer = grd.Selected.Rows(0).Band.Index
                    Dim c As String = "Cancel"
                    If b = 0 Then
                        c &= " All Open Orders on Selected CSO"
                    ElseIf b = 1 Then
                        c &= " All Selected Orders"
                    Else
                        c &= " All Selected Order line items that are Open"
                    End If
                    tlb_btn.SharedProps.Caption = c
                End If

            Case "grdSOTALLOD"
                tlb_btn = DirectCast(tlb_pop.Tools("Item Status Inquiry"), UltraWinToolbars.ButtonTool)
                ' do not mt lock here


                Dim b As Integer = grd.ActiveRow.Band.Index
                If b = 1 And Not (ASCMAIN1.USER_CODES.Contains("FS")) Then
                    ' Only make the button visible if we're at the item level (b = 1)
                    tlb_btn.SharedProps.Visible = True
                Else
                    ' Hide the button if not at the item level
                    tlb_btn.SharedProps.Visible = False
                End If


        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
        Else
            Select Case e.SourceControl.Name

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        Select Case e.Tool.Key
            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grd.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        grow.Cells("SEL").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next
            Case "Cancel Order"
                Dim ORDR_NOs As New List(Of String)
                Dim b As Integer = grd.Selected.Rows(0).Band.Index
                Dim all_clear As Boolean = True
                Dim REASON_CODE As String = "ALLOCA"
                Dim FIRST_CSO As String = Nothing
                Dim ORDR_GROUP_NO As String = Nothing
                Dim MULTI_GROUP As Boolean = False


                ' Handling for different band levels
                If b = 0 Then ' Group level
                    For Each growCSO As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        Dim currentCSO_NO As String = growCSO.Cells("CSO_NO").Value.ToString()
                        If FIRST_CSO Is Nothing Then
                            FIRST_CSO = currentCSO_NO
                        ElseIf FIRST_CSO <> currentCSO_NO Then
                            MsgBox("Cannot select different carstock orders.", MsgBoxStyle.Critical, "Selection Error")
                            MULTI_GROUP = True
                            Exit For
                        End If
                        If MULTI_GROUP Then Exit For

                        ' Check and process each order within this CSO
                        For Each grow As UltraWinGrid.UltraGridRow In growCSO.ChildBands(0).Rows
                            Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
                            Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                            Dim ORDR_STATUS As String = rowSOTORDR1("ORDR_STATUS")
                            ORDR_GROUP_NO = rowSOTORDR1("ORDR_GROUP_NO").ToString()
                            If ORDR_STATUS = "O" Then
                                ORDR_NOs.Add(ORDR_NO)
                                If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO,,,, 1) Then
                                    all_clear = False
                                    Exit For
                                End If
                            End If
                        Next

                        If ORDR_NOs.Count = 0 Then
                            MsgBox("Open Orders Only", MsgBoxStyle.OkOnly, "Cannot Cancel")
                            all_clear = False
                            Exit For
                        End If
                    Next
                ElseIf b = 1 Then ' Order level
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        Dim currentCSO_NO As String = grow.ParentRow.Cells("CSO_NO").Value.ToString()
                        If FIRST_CSO Is Nothing Then
                            FIRST_CSO = currentCSO_NO
                        ElseIf FIRST_CSO <> currentCSO_NO Then
                            MsgBox("Cannot select orders from different carstock orders.", MsgBoxStyle.Critical, "Selection Error")
                            MULTI_GROUP = True
                            Exit For
                        End If
                        If MULTI_GROUP Then Exit For

                        Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
                        Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                        Dim ORDR_STATUS As String = rowSOTORDR1("ORDR_STATUS")
                        ORDR_GROUP_NO = rowSOTORDR1("ORDR_GROUP_NO").ToString()
                        If ORDR_STATUS = "O" Then
                            ORDR_NOs.Add(ORDR_NO)
                            If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO,,,, 1) Then
                                all_clear = False
                                Exit For
                            End If
                        Else
                            MsgBox("Open Orders Only", MsgBoxStyle.OkOnly, "Cannot Cancel")
                            all_clear = False
                            Exit For
                        End If
                    Next
                ElseIf b = 2 Then ' Item level
                    For Each grow As UltraWinGrid.UltraGridRow In grd.Selected.Rows
                        Dim currentCSO_NO As String = grow.ParentRow.ParentRow.Cells("CSO_NO").Value.ToString()
                        If FIRST_CSO Is Nothing Then
                            FIRST_CSO = currentCSO_NO ' Set the first CSO_NO
                        ElseIf FIRST_CSO <> currentCSO_NO Then
                            MsgBox("All selected items must belong to the same carstock order.", MsgBoxStyle.Critical, "Selection Error")
                            MULTI_GROUP = True
                            Exit For
                        End If
                        If MULTI_GROUP Then Exit For

                        Dim ORDR_NO As String = grow.Cells("ORDR_NO").Value
                        Dim ORDR_LNO As Integer = Val(grow.Cells("ORDR_LNO").Value & "")
                        Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                        Dim ORDR_STATUS As String = rowSOTORDR1("ORDR_STATUS")

                        If ORDR_STATUS = "O" Then
                            If Not ASCMAIN1.Logical_Lock("SOTORDR1", ORDR_NO,,,, 1) Then
                                all_clear = False
                                Exit For
                            Else
                                Dim message As String = $"You are about to cancel {grd.Selected.Rows.Count} items. Are you sure you want to proceed?"
                                Dim caption As String = "Confirm Cancellation"
                                Dim result As DialogResult = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                                If result = DialogResult.Yes Then
                                    ' Update the item line's ORDR_QTY_OPEN and ORDR_QTY_CANC
                                    Dependent_Updates_Cancel(-1, ORDR_NO)
                                    ASCMAIN1.sql = "UPDATE SOTORDR2 SET ORDR_STATUS = 'C', ORDR_QTY_CANC = ORDR_QTY_OPEN, ORDR_QTY_OPEN = 0 " _
                            & "WHERE ORDR_NO = '" & ORDR_NO & "' AND ORDR_LNO = " & ORDR_LNO
                                    ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                                End If
                            End If
                        Else
                            MsgBox("Open Order Lines Only", MsgBoxStyle.OkOnly, "Cannot Cancel")
                            all_clear = False
                            Exit For
                        End If

                        ' Check if all items are canceled
                        Dim allItemsCanceled As Boolean = True
                        ASCMAIN1.sql = "SELECT ORDR_QTY_OPEN FROM SOTORDR2 WHERE ORDR_NO = '" & ORDR_NO & "'"
                        Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                        For Each row As DataRow In dt.Rows
                            If row("ORDR_QTY_OPEN") > 0 Then
                                allItemsCanceled = False
                                Exit For
                            End If
                        Next

                        If allItemsCanceled Then
                            ASCMAIN1.sql = "UPDATE SOTORDR1 SET ORDR_STATUS = 'C' WHERE ORDR_NO = '" & ORDR_NO & "'"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql)
                        End If
                    Next
                End If

                ' Confirmation Window
                If all_clear And Not MULTI_GROUP And ORDR_NOs.Count > 0 Then
                    Dim message As String = $"You are about to cancel {ORDR_NOs.Count} orders. Are you sure you want to proceed?"
                    Dim caption As String = "Confirm Cancellation"
                    Dim result As DialogResult = MessageBox.Show(message, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question)

                    If result = DialogResult.Yes Then
                        ' Proceed with cancellation
                        For Each ORDR_NO As String In ORDR_NOs
                            Dependent_Updates_Cancel(-1, ORDR_NO)
                            ASCMAIN1.sql = "Select Sum (ORDR_QTY_PICK) from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
                            Dim ORDR_STATUS As String = ""
                            If Val(ASCDATA1.GetDataValue) <> 0 Then
                                ORDR_STATUS = "P"
                            Else
                                ASCMAIN1.sql = "Select Sum (ORDR_QTY_SHIP) from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
                                If Val(ASCDATA1.GetDataValue) <> 0 Then
                                    ORDR_STATUS = "F"
                                Else
                                    ORDR_STATUS = "C"
                                End If
                            End If

                            If ORDR_STATUS <> "C" Then
                                MsgBox("Warning: Status issue detected with Order " & ORDR_NO)
                            End If

                            ASCMAIN1.sql = "" _
                            & "Begin " _
                            & " Declare Cursor C1 is Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "' for Update;" _
                            & " Begin " _
                            & "  For R1 in C1 Loop" _
                            & "   Update SOTORDR2" _
                            & "    Set ORDR_QTY_CANC = NVL(ORDR_QTY_CANC,0) + NVL(R1.ORDR_QTY_OPEN,0)" _
                            & "      , ORDR_QTY_OPEN = 0, ORDR_STATUS = '" & ORDR_STATUS & "'" _
                            & "    where Current of C1;" _
                            & "  End Loop;" _
                            & " End;" _
                            & "End;"
                            ASCDATA1.ExecuteSQL()

                            TAC.SOCMAIN1.Record_Event_SOTORDR1(ORDR_NO, DATETIME_STAMP, ASCMAIN1.USER_ID, "ORDCXL", "Order Cancelled")

                            ASCMAIN1.sql = "Update SOTORDR1 Set REASON_CODE = :PARM1" _
                            & ", ORDR_STATUS = :PARM2, ORDR_DATE_CLOSED = TRUNC(SYSDATE), ORDR_YYYYPP_CLOSED = :PARM3" _
                            & " where ORDR_NO = :PARM4"
                            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {REASON_CODE, ORDR_STATUS, ASCMAIN1.CYP, ORDR_NO})
                        Next
                    End If
                End If
                ' Check and update the CSO status
                If b = 0 Or b = 1 Then ' Only for CSO and order level operations
                    Dim allOrdersCanceled As Boolean = True
                    ASCMAIN1.sql = "SELECT ORDR_STATUS FROM SOTORDR1 WHERE ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "'"
                    Dim dt As DataTable = ASCDATA1.GetDataTable(ASCMAIN1.sql)
                    For Each row As DataRow In dt.Rows
                        'if the status of any of the orders in that group is still open or in pick, we cant cancel the entire group
                        If row("ORDR_STATUS") = "P" Or row("ORDR_STATUS") = "O" Then
                            allOrdersCanceled = False
                            Exit For
                        End If
                    Next

                    If allOrdersCanceled Then
                        ASCMAIN1.sql = "UPDATE SOTCSTO1 SET CSO_STATUS = 'C' WHERE CSO_NO = '" & FIRST_CSO & "'"
                        ASCDATA1.ExecuteSQL()

                        'UPDATE SOTORDR0 IF WE'RE CANCELLING ALL ORDERS IN THE GROUP/NOTHING ELSE IS OPEN/PICK
                        ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
                    End If
                End If
                Click_Command("Refresh")
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Item Status Inquiry"
                If grd.ActiveRow.IsDataRow Then
                    Dim ITEM_CODE As String = grd.ActiveRow.Cells("ITEM_CODE").Text
                    Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", ITEM_CODE)
                    If rowICTITEM1 IsNot Nothing Then
                        Context_Launch("View", ITEM_CODE, e.Tool.Key, "ICFSTAT1")
                    End If
                End If

            Case "Sales Order Inquiry"
                If grd.ActiveRow.IsDataRow Then
                    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Text
                    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
                    If rowSOTORDR1 IsNot Nothing Then
                        Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
                    End If
                End If

        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Public Overrides Sub txt_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.txt_KeyDown(sender, e)
        Select Case Absx1.GetABSColumnName(sender)
            Case "SELL_CODE"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Dim SELL_CODE_NEW As String = Absx1.txtFor("SELL_CODE").Text & ""
                    If SELL_CODE_NEW <> SELL_CODE AndAlso SELL_CODE_NEW <> "" Then
                        Load_SOTCSTOX()
                    End If

                End If

            Case "CSO_REF_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    If Not InquiryMode _
                       And Absx1.txtFor("SELL_CODE").Text <> "" Then
                        Click_Command("New")
                    End If
                End If

            Case "CSO_NO"
                If e.KeyCode = Windows.Forms.Keys.Enter And Not ScreenMode Then
                    Click_Command("View")
                End If

        End Select
    End Sub

    Overrides Sub Leaving_txt_Special_After(ByVal COLUMN_NAME As String, ByVal ctl As Control)
        Select Case COLUMN_NAME
            Case "SELL_CODE"
                If Not ScreenMode Then
                    Dim SELL_CODE_NEW As String = Absx1.txtFor("SELL_CODE").Text & ""
                    If SELL_CODE_NEW <> SELL_CODE AndAlso SELL_CODE_NEW <> "" Then
                        Load_SOTCSTOX()
                    End If
                End If
        End Select
    End Sub

    Overrides Sub txt_EditorButtonClick_Special(ByVal txtctl As UltraWinEditors.UltraTextEditor)
        Select Case Absx1.GetABSColumnName(txtctl)
            Case "SELL_CODE"
                Load_SOTCSTOX()
            Case "CSO_NO"
                Click_Command("View")
        End Select
    End Sub

    Public Overrides Sub opt_ValueChanged(sender As Object, e As System.EventArgs)
        MyBase.opt_ValueChanged(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME

        End Select
    End Sub

#End Region
    Sub Load_SOTCSTOX()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Gathering Data")

        Dim sqlw As String = ""
        If InquiryMode Then
            If optStatus.Value <> "A" Then
                sqlw &= " and SOTCSTO1.CSO_STATUS = '" & optStatus.Value & "'"
            End If
            If optStatus.Value = "F" Or optStatus.Value = "A" Then
                sqlw &= " and SOTCSTO1.CSO_DATE >= '" & Format(dteCSOFrom.Value, "dd-MMM-yyyy") & "'"
                sqlw &= " and SOTCSTO1.CSO_DATE <= '" & Format(dteCSOTo.Value, "dd-MMM-yyyy") & "'"
            End If
        Else
            sqlw &= " and SOTCSTO1.CSO_STATUS = 'O'"
        End If

        SELL_CODE = Absx1.txtFor("SELL_CODE").Text
        If SELL_CODE = "" Then
            grdSOTCSTOX.Text = optStatus.Text
            grdSOTALLOX.Text = optStatus.Text & " Allocations"
        Else
            sqlw &= " and SELL_CODE = '" & SELL_CODE & "'"
            grdSOTCSTOX.Text = optStatus.Text & " associated with " & SELL_CODE
            grdSOTALLOX.Text = optStatus.Text & " Allocations associated with " & SELL_CODE
        End If

        If optStatus.Value = "F" Or optStatus.Value = "A" Then
            grdSOTCSTOX.Text &= "; CSO's Dated between " & Format(dteCSOFrom.Value, "MM/dd/yyyy") & " and " & Format(dteCSOTo.Value, "MM/dd/yyyy")
        End If

        ASCMAIN1.sql = "Select * from ARTCUST2 where CUST_CODE = :PARM1 and SELL_CODE = :PARM2"
        Dim rowARTCUST2_for_AE As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, , "VV", New String() {CUST_CODE, SELL_CODE})
        If rowARTCUST2_for_AE Is Nothing Then
            CUST_STORE_NO = ""
        Else
            CUST_STORE_NO = rowARTCUST2_for_AE.Item("CUST_STORE_NO")
        End If

        Create_Work_Tables(False, sqlw)

        EnforceConstraints(False)
        Fill_Records("SOTCSTOH")

        Dim tbl As DataTable = ASCDATA1.SelectDistinct(dst.Tables("SOTCSTOH"), {"ORDR_NO"})

        Dim result As List(Of String) = tbl.AsEnumerable().
              Select(Function(row) row.Field(Of String)(0)).
              ToList()

        Dim chunkSize As Integer = 900
        dst.Tables("SOTCSTOD").Rows.Clear()
        For i As Integer = 0 To result.Count - 1 Step chunkSize
            Dim chunk As List(Of String) = result.Skip(i).Take(chunkSize).ToList
            Dim inList As String = String.Join(",", chunk.ToArray)
            Fill_Records("SOTCSTOD", inList, False)
        Next


        Fill_Records("SOTCSTOX")
        Sort_grdColumns(grdSOTCSTOX, "CSO_NO".ToLower)
        grdSOTCSTOX.Visible = True
        EnforceConstraints(True)


        Fill_Records("SOTCSTOI")
        Dim R As Integer = dst.Tables("SOTCSTOI").Rows.Count
        Sort_grdColumns(grdSOTCSTOI, "CSO_NO".ToLower)
        grdSOTCSTOI.DisplayLayout.Bands(0).SortedColumns.Add("ITEM_CODE", False, True)
        grdSOTCSTOX.Visible = True


        If SELL_CODE = "" Then
            Fill_Records("SOTCSTOA", "*")
        Else
            Fill_Records("SOTCSTOA", SELL_CODE)
        End If
        ' Fill_Records("SOTCSTOA")
        Sort_grdColumns(grdSOTCSTOA, "CSO_NO".ToLower)

        dst.Tables("SOTALLOX").Rows.Clear()
        Dim MM As Integer = Val(Format(Now, "MM"))
        Do While (MM - 1) Mod 3 <> 0
            MM = MM - 1
        Loop
        DATE_START_since = CDate(Format(MM, "00") & "/01/" & Format(Now, "yyyy"))


        'DATE_START_until = DATE_START_since.AddMonths(3).AddDays(-1) ' SP end of Quarter
        'DATE_START_until = Now.Date
        DATE_START_until = DATE_START_since.AddMonths(3).AddDays(-1) ' DM end of Quarter

        Absx1.dteFor("DATE_START").Value = DBNull.Value
        If SELL_CODE = "" Then
            ' don't waste time
            dst.Tables("SOTALLOX").Rows.Clear()
            dst.Tables("SOTALLOZ").Rows.Clear()
        Else
            Create_Work_Tables_SOTALLOX()
            Fill_Records("SOTALLOX")
            'Fill_Records("SOTALLOX", New Object() {CUST_CODE, SELL_CODE, DATE_START_since})
        End If



        ' Load Tags

        Dim rowSOTSELL1 As DataRow = LookUp("SOTSELL1", SELL_CODE)

        If (SELL_CODE <> "" And rowSOTSELL1 IsNot Nothing) Then
            tabMaster.Tabs("RSC Tags").Visible = True
        Else
            tabMaster.Tabs("RSC Tags").Visible = False
            chkEditTags.Checked = False
        End If



        EnforceConstraints(False)
        Fill_Records("SOTCSTT1", New Object() {SELL_CODE})
        Fill_Records("SOTCSTT2", New Object() {SELL_CODE})
        Fill_Records("SOTCSTTX", New Object() {SELL_CODE})

        Sort_grdColumns(grdSOTCSTT1, "RSC_TAG")
        Sort_grdColumns(grdSOTCSTTX, "RSSP_NAME")

        grdSOTCSTTX.DisplayLayout.Bands(0).Summaries.Clear()

        'throws an error with or without my change remmed out
        ''Collection was modified; enumeration operation may not execute.'
        Dim dcols_name As New List(Of String)


        For Each DCOL As DataColumn In dst.Tables("SOTCSTTX").Columns
            If DCOL.ColumnName.StartsWith("TAG_") Then
                ' dst.Tables("SOTCSTTX").Columns.Remove(DCOL.ColumnName)
                dcols_name.Add(DCOL.ColumnName)
            End If
        Next
        For Each DCOL_name As String In dcols_name
            dst.Tables("SOTCSTTX").Columns.Remove(DCOL_name)
        Next
        Create_Summary(grdSOTCSTTX, "RSSP_NAME", "Count")

        For Each rowSOTCSTT1 As DataRow In dst.Tables("SOTCSTT1").Select("", "RSC_TAG")
            Dim COL_NAME As String = "TAG_" & rowSOTCSTT1.Item("RSC_TAG")
            Add_Column(COL_NAME)
        Next

        For Each row As DataRow In dst.Tables("SOTCSTT2").Rows
            Dim rsspCode As String = row("RSSP_CODE").ToString()
            Dim rscTag As String = row("RSC_TAG").ToString()
            Dim columnName As String = $"TAG_{rscTag}"

            Dim matchingRows = dst.Tables("SOTCSTTX").Select($"RSSP_CODE = '{rsspCode}'")
            For Each txRow As DataRow In matchingRows
                txRow(columnName) = 1
            Next
        Next




        'Load_Allocation()

        ' Load Allocation Summary

        dst.Tables("SOTALLOD").Rows.Clear()
        dst.Tables("SOTALLOI").Rows.Clear()

        For Each ROW As DataRow In ASCDATA1.SelectDistinct("SOTALLOX", New String() {"DATE_START"}).Select()
            Dim DATE_START_allo As Date = ROW.Item("DATE_START")
            Dim rowSOTALLOD As DataRow = dst.Tables("SOTALLOD").Rows.Add(New Object() {DATE_START_allo})

            Fill_Records("SOTALLOI", New Object() {SELL_CODE, DATE_START_allo}, False)

            'Dim rowSOTALLODA As DataRow = Fill_Record("SOTALLODA", New Object() {SELL_CODE, DATE_START_allo})
            'If rowSOTALLODA IsNot Nothing Then
            '    For Each c As String In New String() {"ORDR_QTY", "ORDR_QTY_OPEN", "ORDR_QTY_PICK", "ORDR_QTY_SHIP", "ORDR_QTY_CANC"}
            '        rowSOTALLOD.Item(c) = rowSOTALLODA.Item(c)
            '    Next
            'End If
        Next

        EnforceConstraints(True)

        grdSOTALLOX.DisplayLayout.Bands(0).SortedColumns.Clear()
        grdSOTALLOX.DisplayLayout.Bands(0).SortedColumns.Add("DATE_START", False, True)
        grdSOTALLOX.Rows.ExpandAll(True)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Add_Column(COL_NAME As String)
        With dst.Tables("SOTCSTTX")
            .Columns.Add(COL_NAME)
            .Columns(COL_NAME).DefaultValue = "0"
            With grdSOTCSTTX.DisplayLayout.Bands(0).Columns(COL_NAME)
                .Style = ColumnStyle.CheckBox
                .Header.Caption = Mid(COL_NAME, 5)
                .Header.Appearance.TextHAlign = HAlign.Center
                .CellAppearance.TextHAlign = HAlign.Center
                .Header.Appearance.BackColor = System.Drawing.Color.White
                .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                .Header.Appearance.BackColor2 = System.Drawing.Color.Orange
                .Hidden = False
                .Width = 80
                .CellActivation = Activation.AllowEdit

                Create_Summary(grdSOTCSTTX, COL_NAME)
            End With
            For Each rowSOTCSTTX As DataRow In dst.Tables("SOTCSTTX").Select("")
                rowSOTCSTTX.Item(COL_NAME) = "0"
            Next
        End With
        ASCMAIN1.grdInitializeLayout(grdSOTCSTTX, Me)

    End Sub

    Sub Create_Work_Tables_SOTALLOX(Optional ORDR_GROUP_NO As String = "")
        Dim DATE_START_since_oracle As String = Format(DATE_START_since, "dd-MMM-yyyy")
        Dim DATE_START_until_oracle As String = Format(DATE_START_until, "dd-MMM-yyyy")
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLOX)
        ASCMAIN1.sql = Replace(Replace(Replace(Replace(Replace(sqlSOTALLOX, ":PARM1", $"'{CUST_CODE}'"), ":PARM2", $"'{SELL_CODE}'"), ":PARM3", $"'{DATE_START_since_oracle}'"), ":PARM4", $"'{DATE_START_until_oracle}'"), ":PARM5", $"'{Absx1.txtFor("WHSE_CODE").Text}'")
        ASCDATA1.ExecuteSQL("Insert into " & SOTALLOX & " " & ASCMAIN1.sql)

        Dim DATE_START_oracle As String = Format(Absx1.dteFor("DATE_START").Value, "dd-MMM-yyyy")
        ASCDATA1.ExecuteSQL("Truncate Table " & SOTALLOZ)

        Dim csoNo As String = Absx1.txtFor("CSO_NO").Text & ""
        Dim groupList As String = ""

        Dim sqlGroups As String =
            "select distinct ORDR_GROUP_NO " & vbCrLf &
            "  from SOTORDR1 " & vbCrLf &
           $" where ORDR_CUST_PO in (select CSO_REF_NO from SOTCSTO1 where CSO_NO = '{csoNo}') " & vbCrLf &
            "   and ORDR_GROUP_NO is not null"

        Dim dtGroups As DataTable = ASCDATA1.GetDataTable(sqlGroups)

        If dtGroups IsNot Nothing AndAlso dtGroups.Rows.Count > 0 Then
            Dim parts As New List(Of String)
            For Each r As DataRow In dtGroups.Rows
                Dim g As String = r.Item("ORDR_GROUP_NO") & ""
                If g <> "" Then parts.Add("'" & g.Replace("'", "''") & "'")
            Next
            groupList = String.Join(",", parts.Distinct().ToArray())
        End If

        Dim excludeClause As String = ""
        If groupList <> "" Then
            excludeClause = " and ORDR_GROUP_NO not in (" & groupList & ") "
        Else
            If ORDR_GROUP_NO = "" Then ORDR_GROUP_NO = "0000000000"
            excludeClause = " and ORDR_GROUP_NO <> '" & ORDR_GROUP_NO.Replace("'", "''") & "' "
        End If

        ASCMAIN1.sql =
            Replace(
                Replace(
                    Replace(sqlSOTALLOZ, ":PARM1", $"'{CUST_CODE}'"),
                    ":PARM2", $"'{CUST_STORE_NO}'"
                ),
                " group by",
                excludeClause & " group by"
            )

        ASCDATA1.ExecuteSQL("Insert into " & SOTALLOZ & " " & ASCMAIN1.sql)

        'If ORDR_GROUP_NO = "" Then ORDR_GROUP_NO = "0000000000"
        'ASCMAIN1.sql = Replace(Replace(Replace(sqlSOTALLOZ, ":PARM1", $"'{CUST_CODE}'"), ":PARM2", $"'{CUST_STORE_NO}'"), " group by", $" and ORDR_GROUP_NO <> '{ORDR_GROUP_NO}'  group by")
        'ASCDATA1.ExecuteSQL("Insert into " & SOTALLOZ & " " & ASCMAIN1.sql)
    End Sub

    Sub Create_Work_Tables(initialize As Boolean, Optional sqlw As String = "")

        If initialize Then
            sqlw = " and rownum < 1"
        End If

        If initialize Then
            ASCMAIN1.sql = "Select SOTALLO1.ALLO_CTL_NO, SOTALLO4.QTY_ALLO EVENT_QTY, SOTALLO4.EVENT, SOTALLO3.QTY_ALLO, SOTALLO3.ALLO_NOTES" & vbCrLf _
           & ", SOTALLO2.QTY_ALLO QTY_ALLO_AES, SOTALLO2.ALLO_NOTES ALLO_NOTES_AES" & vbCrLf _
           & ", SOTALLO1.ITEM_CODE, SOTALLO1.DATE_START, SOTALLO1.DATE_END, SOTALLO1.ALLO_GROUP_CODE" & vbCrLf _
           & ", ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE, ICTCOLL1.HC_CODE" & vbCrLf _
           & ", ICTSTAT2_I.WHSE_QTY_ON_HAND, ICTSTAT2_I.WHSE_QTY_ONPO, ICTSTAT2_I.WHSE_QTY_OPEN, ICTSTAT2_I.WHSE_QTY_PICK" & vbCrLf _
           & "from SOTALLO1, SOTALLO2, SOTALLO3, SOTALLO4, ICTITEM1, ICTCOLL1, ARTCUST2, (Select ITEM_CODE" & vbCrLf _
           & ", SUM (WHSE_QTY_ON_HAND) WHSE_QTY_ON_HAND, SUM (WHSE_QTY_ONPO) WHSE_QTY_ONPO" & vbCrLf _
           & ", SUM (WHSE_QTY_OPEN) WHSE_QTY_OPEN, SUM (WHSE_QTY_PICK) WHSE_QTY_PICK" & vbCrLf _
           & " from ICTSTAT2 where WHSE_CODE in " & vbCrLf _
           & " (SELECT WHSE_CODE FROM ICTWHSE1 WHERE LP_CODE Is Not NULL) group by ITEM_CODE) ICTSTAT2_I" & vbCrLf _
           & " where SOTALLO1.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
           & "   And SOTALLO3.ALLO_CTL_NO = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
           & "   And SOTALLO3.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
           & "   And SOTALLO4.ALLO_CTL_NO(+) = SOTALLO2.ALLO_CTL_NO" & vbCrLf _
           & "   And SOTALLO4.CUST_CODE(+) = SOTALLO2.CUST_CODE" & vbCrLf _
           & "   And SOTALLO2.CUST_CODE = :PARM1" & vbCrLf _
           & "   and SOTALLO3.CUST_STORE_NO = ARTCUST2.CUST_STORE_NO" & vbCrLf _
           & "   and SOTALLO4.CUST_STORE_NO(+) = ARTCUST2.CUST_STORE_NO" & vbCrLf _
           & "   and ARTCUST2.CUST_CODE = SOTALLO2.CUST_CODE" & vbCrLf _
           & "   and ARTCUST2.SELL_CODE = :PARM2" & vbCrLf _
           & "   and SOTALLO1.DATE_END >= :PARM3" & vbCrLf _
           & "   and SOTALLO1.DATE_START <= :PARM4" & vbCrLf _
           & "   and ICTSTAT2_I.ITEM_CODE (+) = SOTALLO1.ITEM_CODE" & vbCrLf _
           & "   and ICTITEM1.ITEM_CODE = SOTALLO1.ITEM_CODE" & vbCrLf _
           & "   and ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
            '& "   and ICTSTAT2.WHSE_CODE (+) = :PARM5" & vbCrLf _
            'SELECT WHSE_CODE FROM ICTWHSE1 WHERE LP_CODE IS NOT NULL
            sqlSOTALLOX = ASCMAIN1.sql
            ASCMAIN1.sql = Replace(Replace(Replace(Replace(Replace(sqlSOTALLOX, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''"), ":PARM5", "''")
            ASCMAIN1.sql = Replace(Replace(Replace(Replace(sqlSOTALLOX, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''")
            SOTALLOX = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL($"Alter Table {SOTALLOX} Add Primary Key (ALLO_CTL_NO)")
            ASCDATA1.ExecuteSQL($"Create Index I_{SOTALLOX}_1 on {SOTALLOX} (ITEM_CODE)")

            ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'O',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_OPEN" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'P',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_PICK" & vbCrLf _
            & ", MIN (DECODE(SOTORDR1.ORDR_STATUS,'F',SOTORDR1.ORDR_SHIP_DATE,NULL)) ORDR_SHIP_DATE_SHIP" & vbCrLf _
            & " from SOTORDR2,SOTORDR1,ARTCUST1" & vbCrLf _
            & $"   where SOTORDR2.ALLO_CTL_NO in  (Select ALLO_CTL_NO from {SOTALLOX})" & vbCrLf _
            & "     and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "     and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "     and SOTORDR2.CUST_CODE = :PARM1" & vbCrLf _
            & "     and SOTORDR2.CUST_STORE_NO = :PARM2" & vbCrLf _
            & "     and SOTORDR2.ORDR_STATUS IN ('O','P','F','C')" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO,SOTORDR2.CUST_CODE)"
            sqlSOTALLOZ = ASCMAIN1.sql
            ASCMAIN1.sql = Replace(Replace(Replace(Replace(sqlSOTALLOZ, ":PARM1", "''"), ":PARM2", "''"), ":PARM3", "''"), ":PARM4", "''")
            SOTALLOZ = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL($"Alter Table {SOTALLOZ} Add Primary Key (ALLO_CTL_NO)")
        End If

        ASCMAIN1.sql = "Select SOTCSTO1.*" & vbCrLf _
        & " from SOTCSTO1" & ASCMAIN1.SQL_Add_WHERE(sqlw)

        If SOTCSTOX = "" Then
            SOTCSTOX = ASCMAIN1.Temp_Table
            ASCDATA1.ExecuteSQL("Alter Table " & SOTCSTOX & " Add Primary Key (CSO_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCSTOX)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCSTOX & " " & ASCMAIN1.sql)
        End If

        ASCMAIN1.sql = "Select SOTCSTO2.ITEM_CODE, ICTITEM1.ITEM_DESC, SOTCSTO2.CSO_LNO, SOTCSTO1.*" & vbCrLf _
            & ", ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & ", ICTCOLL1.HC_CODE" & vbCrLf _
            & ", ICTCOLL1.BRAND_CODE" & vbCrLf _
            & ", ICTITEM1.PROD_CODE" & vbCrLf _
            & " from ICTITEM1, ICTCOLL1, SOTCSTO2, SOTCSTO1, " & SOTCSTOX & " X " & vbCrLf _
            & " where SOTCSTO2.CSO_NO = X.CSO_NO " & vbCrLf _
            & "   And ICTITEM1.ITEM_CODE = SOTCSTO2.ITEM_CODE" & vbCrLf _
            & "   And ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE" & vbCrLf _
            & "   And SOTCSTO1.CSO_NO = SOTCSTO2.CSO_NO"

        If SOTCSTOI = "" Then
            SOTCSTOI = ASCMAIN1.Temp_Table
            'ASCDATA1.ExecuteSQL("Alter Table " & SOTCSTOI & " Add Primary Key (CSO_NO)")
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & SOTCSTOI)
            ASCDATA1.ExecuteSQL("Insert into " & SOTCSTOI & " " & ASCMAIN1.sql)
        End If

        If initialize Then
            ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE, ICTITEM1.ITEM_DESC, ICTITEM1.ITEM_RETAIL_PRICE, ICTITEM1.ITEM_BASIC_PROMO, ICTITEM1.ITEM_SNU_CODE" & vbCrLf _
                & ", ICTITEM1.ITEM_SO_QTY_MULT, ICTITEM1.ITEM_SO_QTY_MIN" & vbCrLf _
                & ", ICTITEM1.COLLECTION_CODE, ICTCOLL1.HC_CODE, ICTCOLL1.BRAND_CODE, ICTITEM1.PROD_CODE" & vbCrLf _
                & " from ICTITEM1,ICTCOLL1" & vbCrLf _
                & " where ICTCOLL1.COLLECTION_CODE = ICTITEM1.COLLECTION_CODE"
            sqlICTITEM1 = ASCMAIN1.sql
            ICTITEM1 = ASCMAIN1.Temp_Table
        Else
            ASCDATA1.ExecuteSQL("Truncate Table " & ICTITEM1)
            ASCDATA1.ExecuteSQL("Insert into " & ICTITEM1 & " " & sqlICTITEM1)
        End If

    End Sub

    Sub Print_Record()

        'Print_Report_Begin()
        'CR_params.Add("SUBT", "")
        'Dim RPT As String = "SORRMAP1" ' unneccesary if Report Name is Like Form Name
        'Generate_Report(RPT, "Car Stock Order", , , , , False)
        'Print_Report_End()

    End Sub

    Private Sub grdSOTCSTOX_DoubleClickRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickRowEventArgs) Handles grdSOTCSTOX.DoubleClickRow

        If e.Row Is Nothing OrElse e.Row.IsFilterRow Then
            Exit Sub
        End If

        Absx1.txtFor("CSO_NO").Text = e.Row.Cells("CSO_NO").Value & String.Empty
        Click_Command("Edit")
    End Sub

    Sub Delete_Order()
        Me.Cursor = Cursors.WaitCursor
        Dim EMsg As String = ""

        BeginTrans()

        If EntryMode = "E" Then
            Delete_Order_1(CSO_NO)

            ' WJZ ADDED THIS TO RE-ESTABLISH SOTORDR0 RECORD, OR ELIMINATE IT IF NOTHING IS LEFT OPEN
            Dim ORDR_GROUP_NO As String = rowSOTCSTO1.Item("ORDR_GROUP_NO") & ""
            If ORDR_GROUP_NO <> "" Then
                ASCDATA1.ExecuteSP("SOPORDR0_G", "V", New Object() {ORDR_GROUP_NO}, New String() {"ORDR_GROUP_NO_IN"})
            End If
        End If

        EMsg = "Car Stock Order No " & CSO_NO & " has been marked as Deleted"

        CommitTrans(EMsg)

        Me.Cursor = Cursors.Default
    End Sub

    Sub Delete_Order_1(CSO_NO As String)
        Dependent_Updates(-1, CSO_NO)

        'find all sales orders associated with this cso
        ASCMAIN1.sql = "Select ordr_no FROM SOTCSTO3 WHERE CSO_NO = '" & CSO_NO & "' and ordr_no is not null"
        Dim ORDR_NOS As DataTable = ASCDATA1.GetDataTable

        ' wjz note - why are we not using the ADO.NET table for this?
        'For Each rowSOTCSTO3 As DataRow In dst.Tables("SOTCSTO3").Select("ORDR_NO IS NOT NULL")

        'go through each order no and mark it as deleted
        For Each row As DataRow In ORDR_NOS.Rows
            Dim ordr_no As String = row.Item(0).ToString()

            ' wjz added this to relieve the qty open from ICTSTAT2
            Dependent_Updates_SOTORDR1(-1, ordr_no)

            ASCMAIN1.sql = "UPDATE SOTORDR2 SET ORDR_QTY_OPEN = 0, ORDR_STATUS = 'D' " _
            & "WHERE ORDR_NO = '" & ordr_no & "' AND ORDR_STATUS <> 'C'"
            ASCDATA1.ExecuteSQL()
            TAC.SOCMAIN1.Record_Event_SOTORDR1(ordr_no, DATETIME_STAMP, ASCMAIN1.USER_ID, "ORDDEL", "Order Deleted")
            'set the order status to D for each order number in sotordr1
            ASCMAIN1.sql = "Update SOTORDR1 Set " _
            & " ORDR_STATUS = :PARM1, ORDR_DATE_CLOSED = TRUNC(SYSDATE), ORDR_YYYYPP_CLOSED = :PARM2" _
            & " where ORDR_NO = :PARM3 and ordr_status <> 'C'"
            ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VVVV", New Object() {"D", ASCMAIN1.CYP, ordr_no})
        Next

        ASCMAIN1.sql = "Update SOTCSTO1 Set CSO_STATUS = :PARM1" _
            & " where CSO_NO = :PARM2"
        ASCDATA1.ExecuteSQL(ASCMAIN1.sql, "VV", New Object() {"D", CSO_NO})
    End Sub

    Sub Dependent_Updates_Cancel(S As Integer, ORDR_NO As String)
        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                ITEM_CODE = rowSOTORDR2.Item("ITEM_CODE")
                Update_ICTSTAT2_Cancel(ITEM_CODE, WHSE_CODE, S * QTY_TO_COMMIT)
            End If
        Next
    End Sub
    Sub Dependent_Updates(S As Integer, CSO_NO As String)

    End Sub

    Sub Display_Totals()

    End Sub

    Sub Load_Events()
        '    grdEvents.RemoveAll
        '    Call Load_Events_1("Entered", "INIT_DATE")
    End Sub

    Private Sub optStatus_ValueChanged(sender As System.Object, e As System.EventArgs) Handles optStatus.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Load_SOTCSTOX()

        grpCSODATE.Visible = (optStatus.Value = "F" Or optStatus.Value = "A")
    End Sub

    Private Sub cmdRefresh_Click(sender As Object, e As EventArgs) Handles cmdRefresh.Click
        Load_SOTCSTOX()
    End Sub

    Function Add_SOTCSTO2(ITEM_CODE As String, QTY As Int64, CSO_LNO As Integer) As DataRow
        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
        Dim rowSOTCSTO2 As DataRow = dst.Tables("SOTCSTO2").NewRow
        With rowSOTCSTO2
            .Item("CSO_NO") = CSO_NO
            .Item("CSO_LNO") = CSO_LNO
            .Item("ITEM_CODE") = ITEM_CODE
            '.Item("CSO_QTY") = QTY
            .Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
            .Item("ITEM_SNU_CODE") = rowICTITEM1.Item("ITEM_SNU_CODE")
            .Item("ITEM_SO_QTY_MULT") = rowICTITEM1.Item("ITEM_SO_QTY_MULT")
            .Item("ITEM_SO_QTY_MIN") = rowICTITEM1.Item("ITEM_SO_QTY_MIN")
            .Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
            .Item("HC_CODE") = rowICTITEM1.Item("HC_CODE")
            .Item("BRAND_CODE") = rowICTITEM1.Item("BRAND_CODE")
            .Item("PROD_CODE") = rowICTITEM1.Item("PROD_CODE")
            Dim row() As DataRow = dst.Tables("SOTALLOX").Select($"ITEM_CODE = '{ITEM_CODE}'")
            If row.Length = 1 Then
                .Item("ALLO_GROUP_CODE") = row(0).Item("ALLO_GROUP_CODE")
            End If
        End With
        dst.Tables("SOTCSTO2").Rows.Add(rowSOTCSTO2)

        Return rowSOTCSTO2
    End Function

    Private Sub grdSOTALLOX_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTALLOX.DoubleClickRow
        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            Dim DATE_START As Date = e.Row.Cells("DATE_START").Value
            Absx1.dteFor("DATE_START").Value = DATE_START
        End If
        Click_Command("New")
    End Sub

    Sub Write_SOTORDRx(ORDR_GROUP_NO As String, rowSOTCSTO3 As DataRow)

        Dim ORDR_NO As String = rowSOTCSTO3.Item("ORDR_NO") & ""
        Dim rowSOTORDR1 As DataRow = Nothing
        Dim CSO_QTY_TOTAL As Integer = Val(rowSOTCSTO3.Item("CSO_QTY_TOTAL") & "")

        Dim ORDR_SHIP_DATE As Date = rowSOTCSTO1.Item("ORDR_SHIP_DATE")
        Dim ORDR_CANCEL_DATE As Date = rowSOTCSTO1.Item("ORDR_CANCEL_DATE")

        Dim SHIP_VIA_CODE As String = rowSOTCSTO1.Item("SHIP_VIA_CODE")
        Dim CUST_STATE As String = rowSOTCSTO3.Item("CUST_STATE")

        Dim rowSOTSVIAS As DataRow = dst.Tables("SOTSVIAS").Rows.Find(New Object() {MENU_ITEM_OBJECT, CUST_STATE})
        If rowSOTSVIAS IsNot Nothing Then
            SHIP_VIA_CODE = rowSOTSVIAS("SHIP_VIA_CODE").ToString
        End If

        Dim CSO_KEY As String = rowSOTCSTO3.Item("CSO_KEY") & String.Empty

        'If ASCMAIN1.Running_in_VS Then
        '    'Stop - to push orders in to SOTORDRx that were not updated because of the checkbox that was not removed.
        '    ORDR_NO = ""
        'End If


        If ORDR_NO <> "" Then
            ' edit mode
            rowSOTORDR1 = Fill_Record("SOTORDR1", ORDR_NO,, False)
            rowSOTORDR1.Item("CSO_KEY") = CSO_KEY

            Dim ORDR_STATUS As String = rowSOTORDR1.Item("ORDR_STATUS")
            If ORDR_STATUS = "O" Then
                If Format(rowSOTORDR1.Item("ORDR_SHIP_DATE"), "yyyyMMdd") <> Format(ORDR_SHIP_DATE, "yyyyMMdd") Then
                    rowSOTORDR1.Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE
                End If
                If Format(rowSOTORDR1.Item("ORDR_CANCEL_DATE"), "yyyyMMdd") <> Format(ORDR_CANCEL_DATE, "yyyyMMdd") Then
                    rowSOTORDR1.Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
                End If
                If rowSOTORDR1.Item("SHIP_VIA_CODE") & "" <> SHIP_VIA_CODE Then
                    rowSOTORDR1.Item("SHIP_VIA_CODE") = SHIP_VIA_CODE
                End If

                rowSOTORDR1.Item("ORDR_HOLD") = rowSOTCSTO1.Item("CSO_SALES_HOLD")

                ASCDATA1.ExecuteSQL($"Delete from SOTORDR5 where ORDR_NO = '{ORDR_NO}'")

                Create_SOTORDR5(ORDR_NO, "BT", rowARTCUST1)
                Create_SOTORDR5(ORDR_NO, "ST", rowSOTCSTO3)
            End If

            Fill_Records("SOTORDR2", ORDR_NO, False)

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}'")
                Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
                Dim rowICTWHSEX As DataRow = dst.Tables("ICTWHSEX").Rows.Find(ITEM_CODE)
                If rowICTWHSEX IsNot Nothing Then
                    rowSOTORDR2.Item("WHSE_CODE") = rowICTWHSEX.Item("WHSE_CODE")
                Else
                    rowSOTORDR2.Item("WHSE_CODE") = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                End If
            Next

            If CSO_QTY_TOTAL = 0 Then
                rowSOTORDR1.Item("ORDR_STATUS") = "C"
            Else
                rowSOTORDR1.Item("ORDR_STATUS") = "O"
            End If
        Else
            ' new order mode
            If CSO_QTY_TOTAL > 0 Then
                ORDR_NO = ASCMAIN1.Next_Control_No("SOTORDR1.ORDR_NO")
                rowSOTCSTO3.Item("ORDR_NO") = ORDR_NO

                Dim ORDR_CUST_PO As String = Absx1.txtFor("CSO_REF_NO").Text

                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

                rowSOTORDR1 = dst.Tables("SOTORDR1").NewRow
                With rowSOTORDR1
                    .Item("ORDR_NO") = ORDR_NO
                    .Item("CUST_CODE") = CUST_CODE
                    .Item("CUST_STORE_NO") = CUST_STORE_NO
                    .Item("ORDR_CUST_PO") = ORDR_CUST_PO
                    .Item("ORDR_DATE") = DATETIME_STAMP.Date

                    .Item("ORDR_SHIP_DATE") = ORDR_SHIP_DATE
                    .Item("ORDR_CANCEL_DATE") = ORDR_CANCEL_DATE
                    .Item("ORDR_ORIG_SHIP_DATE") = ORDR_SHIP_DATE
                    .Item("ORDR_ORIG_CANCEL_DATE") = ORDR_CANCEL_DATE

                    .Item("ORDR_SOURCE") = "S"
                    .Item("ORDR_STATUS") = "O"

                    .Item("INIT_DATE") = DATETIME_STAMP
                    .Item("INIT_OPER") = ASCMAIN1.USER_ID
                    .Item("LAST_DATE") = DATETIME_STAMP
                    .Item("LAST_OPER") = ASCMAIN1.USER_ID

                    .Item("ORDR_DATE_BOOKED") = DATETIME_STAMP.Date
                    .Item("ORDR_YYYYPP_BOOKED") = ASCMAIN1.CYP
                    .Item("ORDR_TYPE_CODE") = "REG"
                    .Item("WHSE_CODE") = rowSOTCSTO1.Item("WHSE_CODE")

                    .Item("ORDR_GROUP_NO") = ORDR_GROUP_NO
                    .Item("CUST_DC_NO") = ""
                    .Item("ORDR_DATE_RECD") = DATETIME_STAMP.Date

                    ' Sold To
                    .Item("CUST_NAME") = rowARTCUST1.Item("CUST_NAME") & ""
                    .Item("SREP_CODE") = rowARTCUST1.Item("SREP_CODE") & ""
                    .Item("SREP2_CODE") = rowARTCUST1.Item("SREP2_CODE") & ""
                    .Item("SELL_CODE") = SELL_CODE
                    .Item("ORDR_PRIORITY") = rowARTCUST1.Item("CUST_PRIORITY_CODE") & ""
                    .Item("CUST_BILL_TO_CUST") = CUST_CODE
                    .Item("EVENT_CODE") = rowARTCUST1.Item("EVENT_CODE") & ""

                    .Item("FRT_TERMS") = rowARTCUST1.Item("FRT_TERMS") & ""
                    .Item("ORDR_SPECIAL_INST") = rowARTCUST1.Item("CUST_SPECIAL_INST") & ""
                    .Item("ORDR_INV_COMMENT") = rowARTCUST1.Item("CUST_INV_COMMENT") & ""

                    'If rowARTCUST2.Item("CUST_STORE_SHIP_VIA_CODE") & String.Empty <> String.Empty Then
                    '    .Item("SHIP_VIA_CODE") = rowARTCUST2.Item("CUST_STORE_SHIP_VIA_CODE")
                    'Else
                    .Item("SHIP_VIA_CODE") = SHIP_VIA_CODE
                    'End If

                    .Item("ORDR_FOB") = rowARTCUST1.Item("CUST_FOB") & ""

                    .Item("CURR_CODE") = "USD"
                    .Item("CURR_EXCH_RATE") = 1

                    .Item("TRADE_CLASS_CODE") = rowARTCUST1.Item("TRADE_CLASS_CODE") & ""
                    .Item("PRICE_CLASS_CODE") = rowARTCUST1.Item("PRICE_CLASS_CODE") & ""
                    .Item("PRICE_LIST_CODE") = rowARTCUST1.Item("PRICE_LIST_CODE") & ""

                    ' Bill To
                    Dim rowARTCUST1_BT As DataRow = LookUp("ARTCUST1", CUST_CODE)
                    .Item("POST_CODE") = rowARTCUST1_BT.Item("POST_CODE") & ""
                    .Item("TERM_CODE") = rowARTCUST1_BT.Item("TERM_CODE") & ""
                    .Item("CUST_FACTOR_IND") = rowARTCUST1_BT.Item("CUST_FACTOR_IND") & ""
                    .Item("CUST_VEND_REF") = rowARTCUST1_BT.Item("CUST_VEND_REF") & ""

                    ' Store
                    If rowARTCUST2 IsNot Nothing Then
                        .Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION") & ""
                    End If

                    .Item("CUST_NO_3PL") = rowARTCUST1.Item("CUST_NO_3PL")
                    .Item("SHIP_TO_3PL") = rowARTCUST2.Item("CUST_STORE_NO_3PL")
                    .Item("ORDR_OVERRIDE_NOT_ALLOCATED") = "0"
                    .Item("ORDR_HOLD") = rowSOTCSTO1.Item("CSO_SALES_HOLD")
                    .Item("ORDR_HIGH_PRIORITY") = "0"
                    .Item("ORDR_INITIAL") = "0"

                    'If ASCDATA1.GetDataValue("Select MIN(ORDR_NO) from SOTORDR1 where CUST_CODE = '" & CUST_CODE & "'") = "" Then
                    '    .Item("ORDR_INITIAL") = "1"
                    'End If

                    .Item("CUST_STORE_LOCATION") = rowSOTCSTO3.Item("CUST_NAME")

                    .Item("BRAND_CODE") = ""
                    .Item("COLLECTION_CODE") = ""
                    .Item("SALES_DIVISION_CODE") = "IP2"

                    .Item("ORDR_ALLO_DATE") = rowSOTCSTO1.Item("DATE_START")

                    .Item("ORDR_ADDR_TYPE_ST") = "MA"
                    .Item("ORDR_SHIP_TO") = "MK"

                    .Item("CUST_DISC_PCT") = rowARTCUST1.Item("CUST_DISC_PCT")

                    .Item("CSO_KEY") = CSO_KEY

                    ' Stop ' ANY SPECIAL INSTRUCTIONS?
                    'ORDR_SPECIAL_INST = ORDR_SPECIAL_INST & String.Empty
                    'ORDR_SPECIAL_INST = ORDR_SPECIAL_INST.Trim
                    'If ORDR_SPECIAL_INST.Length > rowSOTORDR1.Table.Columns("ORDR_SPECIAL_INST").MaxLength Then
                    '    ORDR_SPECIAL_INST = ORDR_SPECIAL_INST.Substring(0, rowSOTORDR1.Table.Columns("ORDR_SPECIAL_INST").MaxLength).Trim
                    'End If
                    'rowSOTORDR1.Item("ORDR_SPECIAL_INST") = ORDR_SPECIAL_INST

                    ' USE 000 + AE OR AC, OR RSSP_CODE
                    Dim ORDR_ROTR_CODE As String = rowSOTCSTO3.Item("CSO_KEY")
                    If rowSOTCSTO3.Item("CSO_TYPE") = "AE" Or rowSOTCSTO3.Item("CSO_TYPE") = "AC" Then
                        ORDR_ROTR_CODE = "000" & ORDR_ROTR_CODE
                    End If
                    If ORDR_ROTR_CODE.Length > 0 Then
                        rowSOTORDR1.Item("ORDR_ROTR") = "1"
                        rowSOTORDR1.Item("ORDR_ROTR_CODE") = ORDR_ROTR_CODE
                    End If
                End With
                dst.Tables("SOTORDR1").Rows.Add(rowSOTORDR1)

                Create_SOTORDR5(ORDR_NO, "BT", rowARTCUST1)
                Create_SOTORDR5(ORDR_NO, "ST", rowSOTCSTO3)
            End If

        End If

        For Each rowSOTORDR1 In dst.Tables("SOTORDR1").Select()
            If Absx1.chkFor("CSO_URGENT").Checked Then
                rowSOTORDR1("ORDR_HIGH_PRIORITY") = "1"
                rowSOTORDR1("ORDR_HIGH_PRIORITY_NOTE") = rowSOTCSTO1("CSO_URGENT_NOTES") & ""
            Else
                rowSOTORDR1("ORDR_HIGH_PRIORITY") = "0"
                rowSOTORDR1("ORDR_HIGH_PRIORITY_NOTE") = DBNull.Value
            End If

            If Absx1.txtFor("CSO_NOTES").ToString & "" <> "" Then
                rowSOTORDR1("ORDR_INTERNAL_NOTES") = rowSOTCSTO1("CSO_NOTES") & ""
            End If
        Next

        If CSO_QTY_TOTAL > 0 Then
            For I As Integer = 1 To MAX_COLs
                Dim C As String = $"CSO_QTY_{Format(I, "000")}"
                Dim QTY As Integer = Val(rowSOTCSTO3.Item(C) & "")
                Dim rowSOTCSTO2() As DataRow = dst.Tables("SOTCSTO2").Select($"CSO_COL = {CStr(I)}")
                If QTY > 0 Then
                    Dim rowSOTORDR2 As DataRow = Add_Item_to_Order(ORDR_NO, QTY, rowSOTCSTO2(0), rowSOTORDR1)
                    If rowSOTORDR1.Item("BRAND_CODE") & "" = "" Then
                        rowSOTORDR1.Item("BRAND_CODE") = rowSOTCSTO2(0).Item("BRAND_CODE")
                        rowSOTORDR1.Item("COLLECTION_CODE") = rowSOTCSTO2(0).Item("COLLECTION_CODE")
                    End If
                    Dim rowSOTCSTO4 As DataRow = dst.Tables("SOTCSTO4").NewRow
                    With rowSOTCSTO4
                        .Item("CSO_NO") = rowSOTCSTO3.Item("CSO_NO")
                        .Item("CSO_LNO") = rowSOTCSTO2(0).Item("CSO_LNO")
                        .Item("CSO_ADDR_LNO") = rowSOTCSTO3.Item("CSO_ADDR_LNO")
                        .Item("CSO_QTY") = QTY
                    End With
                    dst.Tables("SOTCSTO4").Rows.Add(rowSOTCSTO4)
                Else
                    If rowSOTCSTO2.Length = 1 Then
                        Dim rowSOTORDR2 As DataRow = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, Val(rowSOTCSTO2(0).Item("CSO_LNO"))})

                        If rowSOTORDR2 IsNot Nothing Then
                            With rowSOTORDR2
                                .Item("ORDR_QTY_CANC") = .Item("ORDR_QTY_ORIG")
                                .Item("ORDR_QTY_OPEN") = 0
                                .Item("ORDR_STATUS") = "C"
                            End With
                        End If

                    End If
                End If
            Next
        Else
            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select($"ORDR_NO = '{ORDR_NO}'")
                With rowSOTORDR2
                    .Item("ORDR_QTY_CANC") = .Item("ORDR_QTY_ORIG")
                    .Item("ORDR_QTY_OPEN") = 0
                    .Item("ORDR_STATUS") = "C"
                End With
            Next
        End If
    End Sub

    Sub Create_SOTORDR5(ORDR_NO As String, CUST_ADDR_TYPE As String, row As DataRow)
        Dim rowSOTORDR5 As DataRow = dst.Tables("SOTORDR5").NewRow
        With rowSOTORDR5
            .Item("ORDR_NO") = ORDR_NO
            .Item("CUST_ADDR_TYPE") = CUST_ADDR_TYPE
            For Each C As String In CUST_ADDR_cols
                If C <> "CSO_TYPE" Then
                    .Item(C) = row.Item(C)
                End If
            Next
        End With

        Dim CUST_ZIP_CODE As String = rowSOTORDR5.Item("CUST_ZIP_CODE") & ""
        If CUST_ZIP_CODE.Length >= 1 And CUST_ZIP_CODE.Length <= 4 Then
            CUST_ZIP_CODE = CUST_ZIP_CODE.PadLeft(5, "0")
            rowSOTORDR5.Item("CUST_ZIP_CODE") = CUST_ZIP_CODE
        End If

        dst.Tables("SOTORDR5").Rows.Add(rowSOTORDR5)
    End Sub

    Function Add_Item_to_Order(ORDR_NO As String, QTY As Integer, rowSOTCSTO2 As DataRow, rowSOTORDR1 As DataRow) As DataRow

        Dim ITEM_CODE As String = rowSOTCSTO2.Item("ITEM_CODE")
        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)

        Dim rowSOTORDR2 As DataRow
        rowSOTORDR2 = dst.Tables("SOTORDR2").Rows.Find(New Object() {ORDR_NO, Val(rowSOTCSTO2.Item("CSO_LNO"))})
        If rowSOTORDR2 IsNot Nothing Then
            With rowSOTORDR2
                .Item("ORDR_QTY") = QTY
                .Item("ORDR_QTY_ORIG") = QTY
                .Item("ORDR_QTY_OPEN") = QTY
                .Item("ORDR_QTY_CANC") = 0
                .Item("ORDR_STATUS") = "O"
            End With
        Else
            rowSOTORDR2 = dst.Tables("SOTORDR2").NewRow
            With rowSOTORDR2
                .Item("ORDR_NO") = ORDR_NO
                .Item("ORDR_LNO") = rowSOTCSTO2.Item("CSO_LNO")
                .Item("ITEM_CODE") = rowSOTCSTO2.Item("ITEM_CODE")
                .Item("ITEM_DESC") = rowSOTCSTO2.Item("ITEM_DESC")
                .Item("ORDR_UNIT_PRICE") = 0

                .Item("ORDR_QTY") = QTY
                .Item("ORDR_QTY_OPEN") = QTY
                .Item("ORDR_QTY_ORIG") = QTY

                .Item("ORDR_STATUS") = "O"
                .Item("CUST_CODE") = rowSOTCSTO1.Item("CUST_CODE")
                .Item("CUST_STORE_NO") = CUST_STORE_NO
                .Item("SREP_CODE") = rowSOTORDR1.Item("SREP_CODE")
                .Item("WHSE_CODE") = rowSOTORDR1.Item("WHSE_CODE")
                Dim rowICTWHSEX As DataRow = dst.Tables("ICTWHSEX").Rows.Find(ITEM_CODE)
                If rowICTWHSEX IsNot Nothing Then
                    .Item("WHSE_CODE") = rowICTWHSEX.Item("WHSE_CODE")
                Else
                    .Item("WHSE_CODE") = ROWs("SOTPARM1").Item("SO_PARM_DEF_PICK_WHSE")
                End If
                .Item("ITEM_RETAIL_PRICE") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                .Item("ORDR_UNIT_PRICE_CURR") = 0
                .Item("ITEM_RETAIL_PRICE_CURR") = rowICTITEM1.Item("ITEM_RETAIL_PRICE")
                .Item("ALLO_CTL_NO") = rowSOTCSTO2.Item("ALLO_CTL_NO")
                '.Item("ORDR_QTY_ALLO") = ?
                .Item("SELL_CODE") = SELL_CODE
            End With
            dst.Tables("SOTORDR2").Rows.Add(rowSOTORDR2)
        End If

        Return rowSOTORDR2
    End Function

    Sub Dependent_Updates_SOTORDR1(S As Integer, ORDR_NO As String)

        Dim QTY_TO_COMMIT As Int64

        ASCMAIN1.sql = "Select * from SOTORDR1 where ORDR_NO = '" & ORDR_NO & "'"
        Dim rowSOTORDR1 As DataRow = ASCDATA1.GetDataRow

        ASCMAIN1.sql = "Select * from SOTORDR2 where ORDR_NO = '" & ORDR_NO & "'"
        For Each rowSOTORDR2 As DataRow In ASCDATA1.GetDataTable.Rows
            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
            Dim WHSE_CODE As String = rowSOTORDR1.Item("WHSE_CODE")

            If S = -1 Then
            Else
                ' Update_Record_TDA("SOTORDR2")
            End If

            QTY_TO_COMMIT = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
            If QTY_TO_COMMIT <> 0 Then
                ITEM_CODE = rowSOTORDR2.Item("ITEM_CODE")
                Update_ICTSTAT2(ITEM_CODE, WHSE_CODE, S * QTY_TO_COMMIT)
            End If
        Next
    End Sub
    Sub Update_ICTSTAT2(ITEM_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVNNNNNN",
                           New Object() {ITEM_CODE, WHSE_CODE,
                                         0, 0, 0,
                                         QTY, 0, 0},
                           New String() {"ITEM_CODE_IN", "WHSE_CODE_IN",
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in",
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Sub Update_ICTSTAT2_Cancel(ITEM_CODE As String, WHSE_CODE As String, QTY As Int64)
        ASCDATA1.ExecuteSP("ICPSTAT2", "VVNNNNNN",
                           New Object() {ITEM_CODE, WHSE_CODE,
                                         0, 0, 0,
                                         QTY, 0, 0},
                           New String() {"ITEM_CODE_IN", "WHSE_CODE_IN",
                                         "WHSE_QTY_ON_HAND_in", "WHSE_QTY_ON_ORDER_in", "WHSE_QTY_TRAN_in",
                                         "WHSE_QTY_OPEN_in", "WHSE_QTY_PICK_in", "WHSE_QTY_ALLO_in"})
    End Sub

    Private Sub WorkbookView1_CellEndEdit(sender As Object, e As CellEndEditEventArgs) Handles WorkbookView1.CellEndEdit

        Dim R As Integer = Val(e.ActiveCell.Row)
        Dim C As Integer = Val(e.ActiveCell.Column)
        If isClearing Then
            R = isClearing_R
            C = isClearing_C
        End If
        If isPasting Then
            R = isPasting_R
            C = isPasting_C
        End If

        Dim CSO_ADDR_LNO As Integer = e.ActiveCell.Worksheet.Cells(R, COL_CSO_ADDR_LNO).Value
        If C > c0_Items Then ' 10 Then


            Dim ITEM_CODE As String = e.ActiveCell.Worksheet.Cells(ROW_ITEM_CODE, C).Value & ""

            'Dim CSO_LNO As Integer = 0

            Dim qty As Integer = Val(e.Entry)

            'qty = Process_qty(R, C, qty, CSO_ADDR_LNO, ITEM_CODE)

            Dim rowSOTCSTO2 As DataRow = dst.Tables("SOTCSTO2").Select($"ITEM_CODE = '{ITEM_CODE}'")(0)
            Dim CSO_COL As Integer = Val(rowSOTCSTO2.Item("CSO_COL"))
            Dim ITEM_SO_QTY_MULT As Integer = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MULT") & "")
            Dim ITEM_SO_QTY_MIN As Integer = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MIN") & "")

            If qty < 0 Then qty = 0

            If isClearing Then
                qty = 0
            End If

            If ITEM_SO_QTY_MIN > 0 AndAlso qty > 0 AndAlso qty < ITEM_SO_QTY_MIN <> 0 Then
                qty = ITEM_SO_QTY_MIN
            End If

            If ITEM_SO_QTY_MULT > 0 AndAlso qty > 0 AndAlso qty Mod ITEM_SO_QTY_MULT <> 0 Then
                qty += ITEM_SO_QTY_MULT - (qty Mod ITEM_SO_QTY_MULT)
            End If

            rowSOTCSTO2.Item("CSO_QTY_" & Format(CSO_ADDR_LNO, "000")) = qty

            Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Rows.Find(New Object() {CSO_NO, CSO_ADDR_LNO})
            rowSOTCSTO3.Item("CSO_QTY_" & Format(CSO_COL, "000")) = qty

            'If isClearing Then
            '    qty = 0
            '    isClearing = False
            'End If

            e.Entry = qty
        Else

            Dim rowSOTCSTO3 As DataRow = dst.Tables("SOTCSTO3").Rows.Find(New Object() {CSO_NO, CSO_ADDR_LNO})
            Dim COLUMN_NAME As String = ""
            'If C = 4 Then COLUMN_NAME = "CUST_NAME" ' NOT CHANGEABLE?
            If C = 5 Then COLUMN_NAME = "CUST_ADDR1"
            If C = 6 Then COLUMN_NAME = "CUST_ADDR2"
            'If C = 7 Then COLUMN_NAME = "CUST_ADDR3"' WE DO NOT PERMIT SHIP TO ADDRESS3
            If C = 8 Then COLUMN_NAME = "CUST_CITY"
            If C = 9 Then COLUMN_NAME = "CUST_STATE"
            If C = 10 Then COLUMN_NAME = "CUST_ZIP_CODE"

            Dim previousValue = rowSOTCSTO3.Item(COLUMN_NAME)

            Dim isAddressField As Boolean = (C = 5 OrElse C = 6 OrElse C = 8 OrElse C = 9 OrElse C = 10)
            If isAddressField Then
                If rowSOTCSTO3 IsNot Nothing AndAlso Not originalValues.ContainsKey(rowSOTCSTO3) Then
                    originalValues(rowSOTCSTO3) = rowSOTCSTO3.Table.NewRow()
                    originalValues(rowSOTCSTO3).ItemArray = rowSOTCSTO3.ItemArray.Clone()
                End If
                rowSOTCSTO3.Item(COLUMN_NAME) = e.Entry
                If Not modifiedAddresses.Contains(rowSOTCSTO3) Then
                    modifiedAddresses.Add(rowSOTCSTO3)
                End If
            End If
        End If
    End Sub
    Private Sub Validate_Addresses()
        Dim unvalidatedCustomers As New List(Of String)
        Dim remainingAddresses = modifiedAddresses.Except(validatedAddresses).ToList()
        For Each row As DataRow In modifiedAddresses
            unvalidatedCustomers.Add(row("CUST_NAME").ToString())
        Next

        If unvalidatedCustomers.Count > 0 Then
            chkAllowEditShipToAddress.Enabled = True
            chkAllowEditShipToAddress.Checked = True

            Dim customerList As String = String.Join(vbCrLf, unvalidatedCustomers)
            Dim response = MsgBox("The following person(s) have unvalidated addresses:" & vbCrLf &
                              customerList & vbCrLf & vbCrLf &
                              "Please hit OK to proceed with address validation.",
                              MsgBoxStyle.OkCancel, "Address Validation Required")

            If response = MsgBoxResult.Cancel Then
                Exit Sub
            End If

            For Each modifiedRow As DataRow In remainingAddresses
                Dim originalRow As DataRow = originalValues(modifiedRow)
                Dim address As String = modifiedRow("CUST_NAME") & vbCrLf & modifiedRow("CUST_ADDR1") &
                If(modifiedRow("CUST_ADDR2") & "" <> "", vbCrLf & modifiedRow("CUST_ADDR2"), "") &
                vbCrLf & modifiedRow("CUST_CITY") & " " & modifiedRow("CUST_STATE") & " " & modifiedRow("CUST_ZIP_CODE")

                Dim validationResult As String = TAC.TACMAIN1.Validate_Address1(address)
                Dim suggestedAddress As String = Parse_Address(validationResult)
                Dim providedAddress As String = modifiedRow("CUST_ADDR1") &
                If(modifiedRow("CUST_ADDR2") & "" <> "", " " & modifiedRow("CUST_ADDR2"), "") & ", " &
                modifiedRow("CUST_CITY") & " " & modifiedRow("CUST_STATE") & " " & modifiedRow("CUST_ZIP_CODE")
                If suggestedAddress.Replace(",", "").Replace(" ", "").ToLower() <> address.Replace(",", "").Replace(" ", "").ToLower() Then
                    Dim choice = MsgBox("Customer: " & modifiedRow("CUST_NAME") & vbCrLf & vbCrLf &
                                    "Provided Address:" & vbCrLf & providedAddress & vbCrLf & vbCrLf &
                                    "Suggested Address:" & vbCrLf & suggestedAddress & vbCrLf & vbCrLf &
                                    "Would you like to use the suggested address?",
                                    MsgBoxStyle.YesNoCancel, "Address Correction Suggested")

                    If choice = MsgBoxResult.Yes Then
                        Update_Address_Fields(modifiedRow, suggestedAddress)
                        validatedAddresses.Add(modifiedRow) ' Add to validated list
                        UPDATE_DISABLED = False
                    ElseIf choice = MsgBoxResult.Cancel Then
                        UPDATE_DISABLED = True
                        Exit Sub
                    End If
                Else
                    validatedAddresses.Add(modifiedRow) ' Mark as validated if no correction was suggested
                End If
            Next
        End If
    End Sub

    Private Function Parse_Address(validationResult As String) As String
        Dim deliveryLine1Index As Integer = validationResult.IndexOf("Delivery line 1:") + "Delivery line 1:".Length
        Dim lastLineIndex As Integer = validationResult.IndexOf("Last line:") + "Last line:".Length

        Dim deliveryLine1EndIndex As Integer = validationResult.IndexOf(Environment.NewLine, deliveryLine1Index)
        Dim lastLineEndIndex As Integer = validationResult.IndexOf(Environment.NewLine, lastLineIndex)

        Dim deliveryLine1 As String = validationResult.Substring(deliveryLine1Index, deliveryLine1EndIndex - deliveryLine1Index).Trim()
        Dim lastLine As String = validationResult.Substring(lastLineIndex, lastLineEndIndex - lastLineIndex).Trim()

        ' Construct the address and return
        Return $"{deliveryLine1}, {lastLine}"
    End Function


    Private Sub Update_Address_Fields(row As DataRow, suggestedAddress As String)
        WorkbookView1.GetLock()
        Dim parts = suggestedAddress.Split(","c)
        If parts.Length = 2 Then
            row("CUST_ADDR1") = parts(0).Trim() ' First part is the street address

            ' The second part contains city, state, and ZIP
            Dim cityStateZip = parts(1).Trim()

            Dim lastSpaceIndex As Integer = cityStateZip.LastIndexOf(" "c)
            If lastSpaceIndex > 0 Then
                row("CUST_ZIP_CODE") = cityStateZip.Substring(lastSpaceIndex + 1).Trim().Substring(0, 5) ' ZIP code
                cityStateZip = cityStateZip.Substring(0, lastSpaceIndex).Trim() ' Remove ZIP code from the string
            End If

            Dim secondLastSpaceIndex As Integer = cityStateZip.LastIndexOf(" "c)
            If secondLastSpaceIndex > 0 Then
                row("CUST_STATE") = cityStateZip.Substring(secondLastSpaceIndex + 1).Trim() ' State
                row("CUST_CITY") = cityStateZip.Substring(0, secondLastSpaceIndex).Trim() ' City
            End If

            Dim CUST_NAME As String = row("CUST_NAME").ToString()

            Dim lastRow As Integer = ws.UsedRange.Rows.Count

            For rowIndex As Integer = 0 To lastRow - 1
                Dim cell As SpreadsheetGear.IRange = ws.Cells(rowIndex, 4)

                If cell IsNot Nothing AndAlso cell.Value IsNot Nothing AndAlso cell.Value.ToString() = CUST_NAME Then
                    Dim workbookRowIndex As Integer = cell.Row

                    ws.Cells(workbookRowIndex, 5).Value = row("CUST_ADDR1")
                    ws.Cells(workbookRowIndex, 6).Value = row("CUST_ADDR2")
                    ws.Cells(workbookRowIndex, 8).Value = row("CUST_CITY")
                    ws.Cells(workbookRowIndex, 9).Value = row("CUST_STATE")
                    ws.Cells(workbookRowIndex, 10).Value = row("CUST_ZIP_CODE")

                    Exit For
                End If
            Next
        End If
        WorkbookView1.ReleaseLock()
    End Sub

    'Function Process_Qty(qty As Integer, CSO_ADDR_LNO As Integer, ITEM_CODE As String, Optional row2 As DataRow = Nothing, Optional row3 As DataRow = Nothing) As Int32
    '    '(R As Integer, C As Integer, qty As Integer, CSO_ADDR_LNO As Integer, ITEM_CODE As String) As Int32

    '    Dim rowSOTCSTO2 As DataRow
    '    If row2 Is Nothing Then
    '        rowSOTCSTO2 = dst.Tables("SOTCSTO2").Select($"ITEM_CODE = '{ITEM_CODE}'")(0)
    '    Else
    '        rowSOTCSTO2 = row2
    '    End If

    '    Dim CSO_COL As Integer = Val(rowSOTCSTO2.Item("CSO_COL"))
    '    Dim ITEM_SO_QTY_MULT As Integer = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MULT") & "")
    '    Dim ITEM_SO_QTY_MIN As Integer = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MIN") & "")

    '    If qty < 0 Then qty = 0

    '    If ITEM_SO_QTY_MIN > 0 AndAlso qty > 0 AndAlso qty < ITEM_SO_QTY_MIN <> 0 Then
    '        qty = ITEM_SO_QTY_MIN
    '    End If

    '    If ITEM_SO_QTY_MULT > 0 AndAlso qty > 0 AndAlso qty Mod ITEM_SO_QTY_MULT <> 0 Then
    '        qty += ITEM_SO_QTY_MULT - (qty Mod ITEM_SO_QTY_MULT)
    '    End If

    '    rowSOTCSTO2.Item("CSO_QTY_" & Format(CSO_ADDR_LNO, "000")) = qty

    '    Dim rowSOTCSTO3 As DataRow
    '    If row3 Is Nothing Then
    '        rowSOTCSTO3 = dst.Tables("SOTCSTO3").Rows.Find(New Object() {CSO_NO, CSO_ADDR_LNO})
    '    Else
    '        rowSOTCSTO3 = row3
    '    End If

    '    rowSOTCSTO3.Item("CSO_QTY_" & Format(CSO_COL, "000")) = qty

    '    Return qty
    'End Function

    Sub Show_Item(C As Integer, show As Boolean)


        If Not ScreenMode And Not show Then
            splItemInfo.Visible = False
            Exit Sub
        End If

        Dim ITEM_CODE As String = WorkbookView1.ActiveWorkbook.Worksheets(0).Cells(ROW_ITEM_CODE, C).Value

        If UltraExplorerBar1.Groups("Item Info").Text & "" = ITEM_CODE Then ' same item
            Exit Sub
        End If

        splItemInfo.Visible = False

        If C > c0_Items Then ' WE ARE IN THE ITEMS AREA
        Else
            UltraExplorerBar1.Groups("Item Info").Text = ""
            Exit Sub
        End If
        Dim rowICTITEM1 As DataRow = dst.Tables("ICTITEM1").Rows.Find(ITEM_CODE)
        If rowICTITEM1 Is Nothing Then
            UltraExplorerBar1.Groups("Item Info").Text = ""
            Exit Sub
        End If

        If ITEM_CODE = "" Or ITEM_CODE = "Total" Then Exit Sub

        Dim IMAGE_NAME As String = ITEM_CODE

        splItemInfo.Visible = True

        Dim FOLDER_NAME As String = ROWs("ICTPARM1").Item("IC_PARM_IMAGES_FOLDER") & ""
        If ASCMAIN1.Running_in_VS Then
            FOLDER_NAME = "C:\Share\INT\Pictures"
        End If
        Dim imgba() As Byte = Nothing
        picItemImage.Image = ASCMAIN1.Get_Image(FOLDER_NAME, IMAGE_NAME, False, , , imgba)
        UltraExplorerBar1.Groups("Item Info").Text = ITEM_CODE

        Dim rowSOTCSTO2 As DataRow = dst.Tables("SOTCSTO2").Select($"ITEM_CODE = '{ITEM_CODE}'")(0)
        txtITEM_DESC.Text = rowSOTCSTO2.Item("ITEM_DESC") & ""
        numITEM_SO_QTY_MULT.Value = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MULT") & "")
        numITEM_SO_QTY_MIN.Value = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MIN") & "")

        Dim rowSOTALLOXs() As DataRow = dst.Tables("SOTALLOX").Select($"ITEM_CODE = '{ITEM_CODE}'")
        If rowSOTALLOXs.Length > 0 Then
            Dim rowSOTALLOX As DataRow = rowSOTALLOXs(0)
            dteStart.Value = rowSOTALLOX.Item("DATE_START")
            dteEnd.Value = rowSOTALLOX.Item("DATE_END")
            txtALLO_NOTES.Text = rowSOTALLOX.Item("ALLO_NOTES_AES") & ""
        Else
            dteStart.Value = Null
            dteEnd.Value = Null
            txtALLO_NOTES.Text = ""
        End If

    End Sub

    Private Sub WorkbookView1_Leave(sender As Object, e As EventArgs) Handles WorkbookView1.Leave
        WorkbookView1.EndEdit()
    End Sub

    Private Sub WorkbookView1_CellBeginEdit(sender As Object, e As CellBeginEditEventArgs) Handles WorkbookView1.CellBeginEdit
        'Debug.Write(e.Reason, e.Entry)\
    End Sub

    Private Sub WorkbookView1_RangeSelectionChanged(sender As Object, e As RangeSelectionChangedEventArgs) Handles WorkbookView1.RangeSelectionChanged
        ' Debug.WriteLine(e.RangeSelection.Address)
        Show_Item(e.RangeSelection.Column, False)
    End Sub

    Private Sub btnRefreshALLO_GROUP_CODE_Click(sender As Object, e As EventArgs) Handles btnRefreshALLO_GROUP_CODE.Click
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Applying Filters")
        Refresh_CODEs()
        Apply_Filters()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub btnRefreshHC_CODE_Click(sender As Object, e As EventArgs) Handles btnRefreshHC_CODE.Click
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Applying Filters")
        Refresh_CODEs()
        Apply_Filters()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Refresh_CODEs()
        ALLO_GROUP_CODEs.Clear()
        For Each row As DataRow In dst.Tables("SOTALLOG").Select("SEL = '1'")
            ALLO_GROUP_CODEs.Add(row.Item(0))
        Next

        HC_CODEs.Clear()
        For Each row As DataRow In dst.Tables("ICTCOLL0").Select("SEL = '1'")
            HC_CODEs.Add(row.Item(0))
        Next

        RSC_Tags.Clear()
        For Each row As DataRow In dst.Tables("SOTRSCT1").Select("SEL = '1'")
            RSC_Tags.Add(row.Item(1))
        Next

        PROD_CODES.Clear()
        For Each row As DataRow In dst.Tables("ICTPROD1").Select("SEL = '1'")
            PROD_CODES.Add(row.Item(0))
        Next



    End Sub

    Sub Apply_Filters()

        WorkbookView1.GetLock()
        ws.Unprotect(XLS_PWD)

        Dim c_min As Integer = -1

        For Each rowSOTCSTO2 As DataRow In dst.Tables("SOTCSTO2").Select("", "CSO_LNO")
            Dim CSO_LNO As Integer = Val(rowSOTCSTO2.Item("CSO_LNO") & "")
            Dim ALLO_CTL_NO As String = rowSOTCSTO2.Item("ALLO_CTL_NO") & ""
            'Dim rowSOTALLOX As DataRow = dst.Tables("SOTALLOX").Rows.Find(New String() {ALLO_CTL_NO})
            Dim QTY_ALLO As Int32 = Val(rowSOTCSTO2.Item("QTY_ALLO") & "")
            Dim QTY_LEFT As Int32 = Val(rowSOTCSTO2.Item("QTY_LEFT") & "")
            Dim ORDR_QTY_OPEN As Int32 = Val(rowSOTCSTO2.Item("ORDR_QTY_OPEN") & "")
            Dim ORDR_QTY_PICK As Int32 = Val(rowSOTCSTO2.Item("ORDR_QTY_PICK") & "")
            Dim ORDR_QTY_SHIP As Int32 = Val(rowSOTCSTO2.Item("ORDR_QTY_SHIP") & "")
            Dim ORDR_QTY_CANC As Int32 = Val(rowSOTCSTO2.Item("ORDR_QTY_CANC") & "")
            Dim HC_CODE As String = rowSOTCSTO2.Item("HC_CODE") & ""
            Dim ALLO_GROUP_CODE As String = rowSOTCSTO2.Item("ALLO_GROUP_CODE") & ""
            Dim ITEM_CODE As String = rowSOTCSTO2.Item("ITEM_CODE") & ""
            Dim PROD_CODE As String = rowSOTCSTO2.Item("PROD_CODE") & ""
            Dim hide_column As Boolean = False
            If ALLO_GROUP_CODEs.Count > 0 AndAlso Not ALLO_GROUP_CODEs.Contains(ALLO_GROUP_CODE) Then hide_column = True
            If HC_CODEs.Count > 0 AndAlso Not HC_CODEs.Contains(HC_CODE) Then hide_column = True
            If PROD_CODES.Count > 0 AndAlso Not PROD_CODES.Contains(PROD_CODE) Then hide_column = True
            If chkShowItemsLeft2Order.Checked AndAlso QTY_LEFT <= 0 Then hide_column = True
            Dim colIndex As Integer = -1
            For i As Integer = c0_Items To c0_Items + ws.UsedRange.Columns.Count
                If ws.Cells(ROW_ITEM_CODE, i).Value = ITEM_CODE Then
                    colIndex = i
                    Exit For
                End If
            Next
            If Not hide_column And c_min = -1 Then
                c_min = CSO_LNO
            End If
            WorkbookView1.ActiveWorkbook.Worksheets(0).Cells(ROW_ITEM_CODE - 1, colIndex).EntireColumn.Hidden = hide_column 'col y 
        Next

        ws.Cells(ROW_ITEM_CODE - 1, c0_Items + c_min).Activate()

        ws.Protect(XLS_PWD)
        WorkbookView1.ReleaseLock()

    End Sub

    Private Sub grdSOTALLOD_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles grdSOTALLOD.DoubleClickRow
        If e.Row.IsDataRow AndAlso Not e.Row.IsFilterRow Then
            Dim DATE_START As Date = e.Row.Cells("DATE_START").Value
            Absx1.dteFor("DATE_START").Value = DATE_START
        End If
        Click_Command("New")
    End Sub

    Private Sub chkHideAddressColumns_CheckedChanged(sender As Object, e As EventArgs) Handles chkHideAddressColumns.CheckedChanged
        If chkHideAddressColumns.Checked Then
            chkAllowEditShipToAddress.Enabled = False
        Else
            chkAllowEditShipToAddress.Enabled = True
        End If
        Hide_address_Columns()
    End Sub
    Sub Hide_address_Columns()
        WorkbookView1.GetLock()
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        oSheet = WorkbookView1.ActiveWorkbook.Worksheets(0)
        oSheet.Unprotect(XLS_PWD)

        'ws.Cells.Locked = False
        'oSheet.Cells(0, 4).EntireColumn.Hidden = chkHideAddressColumns.Checked name
        oSheet.Cells(0, 5).EntireColumn.Hidden = chkHideAddressColumns.Checked 'address
        oSheet.Cells(0, 6).EntireColumn.Hidden = chkHideAddressColumns.Checked 'address 2 
        'oSheet.Cells(0, 7).EntireColumn.Hidden = chkHideAddressColumns.Checked 'address 3
        'oSheet.Cells(0, 8).EntireColumn.Hidden = chkHideAddressColumns.Checked 'city
        oSheet.Cells(0, 9).EntireColumn.Hidden = chkHideAddressColumns.Checked 'state
        oSheet.Cells(0, 10).EntireColumn.Hidden = chkHideAddressColumns.Checked 'zip

        oSheet.Protect(XLS_PWD)
        WorkbookView1.ReleaseLock()
    End Sub

    Sub Edit_Ship_To_Addresses()

        Dim CSO_ADDR_LNO_max As Integer = Val(dst.Tables("SOTCSTO3").Compute("MAX(CSO_ADDR_LNO)", "") & "")
        Dim r0T As Integer = 12
        Dim r0 As Integer = r0T

        WorkbookView1.GetLock()
        Dim oSheet As SpreadsheetGear.IWorksheet = Nothing
        oSheet = WorkbookView1.ActiveWorkbook.Worksheets(0)
        oSheet.Unprotect(XLS_PWD)

        'ws.Cells.Locked = False
        range = oSheet.Cells(r0T + 1, 5, r0T + CSO_ADDR_LNO_max, 10)
        range.Locked = Not chkAllowEditShipToAddress.Checked
        If chkAllowEditShipToAddress.Checked Then
            range.Interior.Color = SpreadsheetGear.Colors.AliceBlue
        Else
            range.Interior.Color = SpreadsheetGear.Colors.White
        End If

        oSheet.Protect(XLS_PWD)
        WorkbookView1.ReleaseLock()
    End Sub

    Private Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        WorkbookView1.GetLock()
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & Me.Name & "_" & ASCMAIN1.Next_Control_No($"{Me.Name}.XLSX_NO") & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)
        Show_Document(FILENAME)
        WorkbookView1.ReleaseLock()
    End Sub

    Sub email_to_Self()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing email")

        WorkbookView1.GetLock()
        Dim ws As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.Worksheets(0)

        ws.Unprotect(XLS_PWD)
        ws.Range(8, 0).EntireRow.Hidden = False
        ws.Range(9, 0).EntireRow.Hidden = False
        ws.Range(10, 0).EntireRow.Hidden = False
        ws.Range(11, 0).EntireRow.Hidden = False
        chkAllowEditShipToAddress.Checked = False
        ws.Protect(XLS_PWD)
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & Me.Name & "_" & CSO_NO & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

        If ASCMAIN1.Running_in_VS Then
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
        Else
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        ATTACHMENTs.Add("Car-Stock Order", FILENAME)

        Dim EMAIL_SUBJECT As String = "Car-Stock Order " & Absx1.txtFor("CSO_REF_NO").Text
        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    EMAIL_SUBJECT, "CARSTOCK", True, False, CSO_NO, "CSO_NO", "Car-Stock Order", EMAIL_SUBJECT)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        MsgBox("email Sent", MsgBoxStyle.OkOnly, "Verification")

        ws.Unprotect(XLS_PWD)
        ws.Range(8, 0).EntireRow.Hidden = True
        ws.Range(9, 0).EntireRow.Hidden = True
        ws.Range(10, 0).EntireRow.Hidden = True
        ws.Range(11, 0).EntireRow.Hidden = True
        ws.Protect(XLS_PWD)

        WorkbookView1.ReleaseLock()
    End Sub

    Private Sub txtFinditem_KeyDown(sender As Object, e As KeyEventArgs) Handles txtFinditem.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim ITEM_CODE = Replace(txtFinditem.Text.Trim, "'", "")
            Find_Item(ITEM_CODE)
        End If
    End Sub

    Sub Find_Item(ITEM_CODE As String)

        Dim rowSOTCSTO2() As DataRow = dst.Tables("SOTCSTO2").Select($"ITEM_CODE = '{ITEM_CODE}'")
        If rowSOTCSTO2.Length = 0 Then
            'MsgBox($"Cannot Find Item {ITEM_CODE}", MsgBoxStyle.OkOnly, "Cannot Find Item")
        Else
            Dim CSO_LNO As Integer = Val(rowSOTCSTO2(0).Item("CSO_LNO"))
            WorkbookView1.GetLock()
            'GetLock
            WorkbookView1.ActiveWorkbook.Worksheets(0).Cells(ROW_ITEM_CODE, c0_Items + CSO_LNO).Select()
            WorkbookView1.ReleaseLock()
            txtFinditem.Text = ""
        End If
    End Sub
    Private Sub chkShowItemsLeft2Order_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowItemsLeft2Order.CheckedChanged
        Apply_Filters()
    End Sub

    Private Sub WorkbookView1_RangeChanged(sender As Object, e As RangeChangedEventArgs) Handles WorkbookView1.RangeChanged

        'Dim range As SpreadsheetGear.IRange = e.Range ' DirectCast(xlscommand, CommandRange.PasteSpecial).Range
        'If isPasting Then

        'End If
        'For Each cell As SpreadsheetGear.IRange In range.Cells
        '    If cell.Column > c0_Items And c0_Items > 0 Then
        '        WorkbookView1.BeginEdit()
        '        cell.Value = Val(cell.Value & "") * 1
        '        WorkbookView1.EndEdit()
        '    End If

        'Next

    End Sub

    Private Sub UltraTextEditor50_ValueChanged(sender As Object, e As EventArgs) Handles UltraTextEditor50.ValueChanged

    End Sub



    'Private Sub MenuItemCopyNote_Click(ByVal sender As Object, ByVal e As EventArgs)
    '    Dim item As ToolStripItem = CType(sender, ToolStripItem)
    '    If item.Text = "Copy to All Stores for Customer" Then
    '        WorkbookView1.GetLock()
    '        Try
    '            '' Merging is only valid for multi-cell ranges
    '            'If WorkbookView1.RangeSelection.CellCount >= 2 Then
    '            '    WorkbookView1.RangeSelection.Merge()
    '            'End If
    '            '  SpreadsheetGear.Commands.CommandUndoSupport()

    '            WorkbookView1.ActiveCommandManager.CreateCommandPaste(WorkbookView1.ActiveCell)
    '            Dim CP As SpreadsheetGear.Commands.Command = New SpreadsheetGear.Commands.CommandRange.PasteSpecial(range, SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)

    '            '    Dim C As SpreadsheetGear.Commands.Command = WorkbookView1.ActiveCommandManager.CreateCommandPaste(WorkbookView1.ActiveCell, SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)
    '            ' range = WorkbookView1.ActiveCell
    '            range = WorkbookView1.RangeSelection

    '            If EntryMode = "R" Or range.Cells.ColumnCount <> 1 Or range.Cells.CellCount <> 1 Or range.Cells.RowCount <> 1 Then
    '                MsgBox("This Option available on Store Sheets with 1 cell selected")
    '            Else
    '                ' MAKE SURE WE ARE IN A CUSTOMER-STORE SHEET

    '                If Not xls_STOREs.Contains(range.Worksheet.Name) Then
    '                    MsgBox("This Option available on Store Sheets with 1 cell selected")
    '                Else

    '                    Dim NOTE As String = range.Value
    '                    Dim ADDRESS As String = Replace(range.GetAddress(True, True, SpreadsheetGear.ReferenceStyle.A1, False, Nothing), "$", "")

    '                    If NOTE = "" Or Not ADDRESS.StartsWith("G") Then
    '                        MsgBox("This Option available on non-empty Notes cells")
    '                    Else

    '                        Dim CUST_CODE As String = Split(range.Worksheet.Name, "-")(0)
    '                        Dim rows() As DataRow = dst.Tables("ARTCUST2").Select("CUST_CODE = '" & CUST_CODE & "'", "HAS_BUDGET DESC,CUST_CODE,CUST_STORE_NO")

    '                        If MsgBox("OK to Copy note '" & NOTE & "' to All " & CUST_CODE & " Stores?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then

    '                            Dim error_message As String = ""
    '                            For Each row As DataRow In rows
    '                                Dim C_VALUE As String = row.Item("CUST_CODE")
    '                                Dim S_VALUE As String = row.Item("CUST_STORE_NO")
    '                                Dim CS As String = C_VALUE & "-" & S_VALUE

    '                                If xls_STOREs.Contains(CS) Then
    '                                    worksheet = workbook.Worksheets(CS)
    '                                    Try
    '                                        worksheet.Cells(ADDRESS).Value = NOTE
    '                                    Catch ex As Exception
    '                                        If ex.Message <> "Operation is not valid on locked cells." And error_message = "" Then error_message = ex.Message
    '                                    End Try

    '                                End If
    '                            Next
    '                            Dim msg As String = "No Errors - Copy was Successful"
    '                            If error_message <> "" Then msg = "There were errors during this copy." & vbCrLf & "The message could Not be copied to some stores:" & vbCrLf & vbCrLf & error_message
    '                            MsgBox(msg, MsgBoxStyle.OkOnly, "Copy Complete")

    '                        End If
    '                    End If

    '                End If
    '            End If

    '        Finally
    '            WorkbookView1.ReleaseLock()
    '        End Try
    '    End If
    'End Sub

    Private Sub MenuItemUndo_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim item As ToolStripItem = CType(sender, ToolStripItem)
        If item.Text = "Undo" Then
            WorkbookView1.GetLock()
            Try
                '' Merging is only valid for multi-cell ranges
                'If WorkbookView1.RangeSelection.CellCount >= 2 Then
                '    WorkbookView1.RangeSelection.Merge()
                'End If
                '  SpreadsheetGear.Commands.CommandUndoSupport()
                WorkbookView1.ActiveCommandManager.Undo()
            Finally
                WorkbookView1.ReleaseLock()
            End Try
        End If
    End Sub

    Public Sub Round_Qty_to_Multiple(cell As SpreadsheetGear.IRange)
        Dim qty As Int32 = Val(cell.Value) * 1
        Debug.Print(cell.Address & ":" & cell.Value)
        ' qty is 0 even though I pasted a non-zero value
    End Sub

    Public Class MyCommandManager
        ' Pass in other things into the constructor that you might need to sync with your ado.net routine.
        Inherits SpreadsheetGear.Commands.CommandManager

        Public wb As SpreadsheetGear.IWorkbookSet
        Public WBV As WorkbookView
        'Public isClearing As Boolean
        Public frm As SOFCSTO1
        Friend Sub New(workbookSet As SpreadsheetGear.IWorkbookSet, workbookView As WorkbookView, ByRef isClearing As Boolean, ByRef isPasting As Boolean, frm As SOFCSTO1)
            MyBase.New(workbookSet)
            wb = workbookSet
            WBV = workbookView
            'Me.isClearing = isClearing
            Me.frm = frm
        End Sub

        Public Overrides Function CreateCommandPaste(range As SpreadsheetGear.IRange) As SpreadsheetGear.Commands.Command
            ' This is what would normally be called...
            ' return new CommandRange.Paste(range);  

            ' Anytime a Paste command is invoked, this will force a "Paste Values"
            Return New SpreadsheetGear.Commands.CommandRange.PasteSpecial(range, SpreadsheetGear.PasteType.Values, SpreadsheetGear.PasteOperation.None, False, False)
            'Return New SpreadsheetGear.Commands.CommandRange.Paste(range)
        End Function

        Public Overrides Function Execute(ByVal xlscommand As Command) As Boolean

            Dim rowSOTCSTO2 As DataRow = Nothing
            Dim rowSOTCSTO3 As DataRow = Nothing

            Dim LNO_last As Integer = -1
            Dim CNO_last As Integer = -1

            ' c0_items (18 as of this writing, XL Col S) is the Total column, which is just before the blue items area
            ' ROW_ITEM_CODE (12 as if this writing, XL Row 13) is the row where the item codes appear, which is just before the blue items area


            Dim rtval As Boolean = MyBase.Execute(xlscommand)
            If Not rtval Then
                Return False
            End If

            If TypeOf xlscommand Is CommandRange.Clear OrElse TypeOf xlscommand Is CommandRange.ClearContents Then

                Dim range As SpreadsheetGear.IRange = DirectCast(xlscommand, CommandRange.ClearContents).Range
                frm.Progress("Now Clearing")

                For Each cell As SpreadsheetGear.IRange In WBV.RangeSelection ' DirectCast(xlscommand, CommandRange.PasteSpecial).Range
                    Dim R As Integer = cell.Row
                    Dim C As Integer = cell.Column

                    Dim LNO As Integer = cell.Worksheet.Cells(R, frm.c0_Items - 1).Value ' c0_items -1 is aimed at the Line Col R, which is just before the Totals Col S
                    If LNO > 0 Then
                        If LNO <> LNO_last Then rowSOTCSTO3 = frm.dst.Tables("SOTCSTO3").Rows.Find(New Object() {frm.CSO_NO, LNO}) : LNO_last = LNO
                        Dim CNO As Integer = C - frm.c0_Items ' aimed at getting to the item's relative column in the blue area, 1-999
                        If CNO <> CNO_last Then rowSOTCSTO2 = frm.dst.Tables("SOTCSTO2").Rows.Find(New Object() {frm.CSO_NO, CNO}) : CNO_last = CNO

                        rowSOTCSTO2.Item("CSO_QTY_" & Format(LNO, "000")) = DBNull.Value
                        rowSOTCSTO3.Item("CSO_QTY_" & Format(CNO, "000")) = DBNull.Value

                        Debug.Print("Clearing: " & CStr(R) & ":" & CStr(C))
                    End If
                Next

                frm.Progress("")

            End If


            If TypeOf xlscommand Is CommandRange.Cut Then ' BELOW CODE IS VERY SIMILAR TO CLEAR

                Dim range As SpreadsheetGear.IRange = DirectCast(xlscommand, CommandRange.Cut).Range
                frm.Progress("Now Clearing")

                For Each cell As SpreadsheetGear.IRange In WBV.RangeSelection ' DirectCast(xlscommand, CommandRange.PasteSpecial).Range
                    Dim R As Integer = cell.Row
                    Dim C As Integer = cell.Column

                    Dim LNO As Integer = cell.Worksheet.Cells(R, frm.c0_Items - 1).Value ' c0_items -1 is aimed at the Line Col R, which is just before the Totals Col S
                    If LNO > 0 Then
                        If LNO <> LNO_last Then rowSOTCSTO3 = frm.dst.Tables("SOTCSTO3").Rows.Find(New Object() {frm.CSO_NO, LNO}) : LNO_last = LNO
                        Dim CNO As Integer = C - frm.c0_Items ' aimed at getting to the item's relative column in the blue area, 1-999
                        If CNO <> CNO_last Then rowSOTCSTO2 = frm.dst.Tables("SOTCSTO2").Rows.Find(New Object() {frm.CSO_NO, CNO}) : CNO_last = CNO

                        rowSOTCSTO2.Item("CSO_QTY_" & Format(LNO, "000")) = DBNull.Value
                        rowSOTCSTO3.Item("CSO_QTY_" & Format(CNO, "000")) = DBNull.Value

                        Debug.Print("Clearing: " & CStr(R) & ":" & CStr(C))
                    End If
                Next

                frm.Progress("")

            End If

            If TypeOf xlscommand Is CommandRange.PasteSpecial Then

                frm.Progress("Now Pasting")

                Dim range As SpreadsheetGear.IRange = DirectCast(xlscommand, CommandRange.PasteSpecial).Range

                For Each cell As SpreadsheetGear.IRange In WBV.RangeSelection ' DirectCast(xlscommand, CommandRange.PasteSpecial).Range
                    Dim R As Integer = cell.Row
                    Dim C As Integer = cell.Column

                    Dim LNO As Integer = cell.Worksheet.Cells(R, frm.c0_Items - 1).Value ' c0_items -1 is aimed at the Line Col R, which is just before the Totals Col S
                    If LNO <> LNO_last Then rowSOTCSTO3 = frm.dst.Tables("SOTCSTO3").Rows.Find(New Object() {frm.CSO_NO, LNO}) : LNO_last = LNO
                    Dim CNO As Integer = C - frm.c0_Items ' aimed at getting to the item's relative column in the blue area, 1-999
                    If CNO <> CNO_last Then rowSOTCSTO2 = frm.dst.Tables("SOTCSTO2").Rows.Find(New Object() {frm.CSO_NO, CNO}) : CNO_last = CNO

                    Dim qty As Int32 = Val(cell.Value & "")
                    If qty = 0 Then
                        rowSOTCSTO2.Item("CSO_QTY_" & Format(LNO, "000")) = DBNull.Value
                        rowSOTCSTO3.Item("CSO_QTY_" & Format(CNO, "000")) = DBNull.Value
                    Else
                        Dim qty_original As Int32 = qty
                        ' Stop ' qty = frm.Process_Qty(R, C, qty, CSO_ADDR_LNO, ITEM_CODE)
                        qty = Process_Qty(qty, rowSOTCSTO2, rowSOTCSTO3)
                        If qty <> qty_original Then
                            cell.Value = qty
                        End If
                        'rowSOTCSTO2.Item("CSO_QTY_" & Format(LNO, "000")) = qty
                        'rowSOTCSTO3.Item("CSO_QTY_" & Format(CNO, "000")) = qty
                    End If

                    Debug.Print("Pasting: " & CStr(R) & ":" & CStr(C))
                Next

                frm.Progress("")
            End If

            Return rtval ' MyBase.Execute(xlscommand)
        End Function


        Function Process_Qty(qty As Integer, rowSOTCSTO2 As DataRow, rowSOTCSTO3 As DataRow) As Int32
            '(R As Integer, C As Integer, qty As Integer, CSO_ADDR_LNO As Integer, ITEM_CODE As String) As Int32

            Dim CSO_COL As Integer = Val(rowSOTCSTO2.Item("CSO_COL"))
            Dim ITEM_SO_QTY_MULT As Integer = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MULT") & "")
            Dim ITEM_SO_QTY_MIN As Integer = Val(rowSOTCSTO2.Item("ITEM_SO_QTY_MIN") & "")
            Dim CSO_ADDR_LNO As Integer = Val(rowSOTCSTO3.Item("CSO_ADDR_LNO"))

            If qty < 0 Then qty = 0

            If ITEM_SO_QTY_MIN > 0 AndAlso qty > 0 AndAlso qty < ITEM_SO_QTY_MIN <> 0 Then
                qty = ITEM_SO_QTY_MIN
            End If

            If ITEM_SO_QTY_MULT > 0 AndAlso qty > 0 AndAlso qty Mod ITEM_SO_QTY_MULT <> 0 Then
                qty += ITEM_SO_QTY_MULT - (qty Mod ITEM_SO_QTY_MULT)
            End If

            If qty = 0 Then
                rowSOTCSTO2.Item("CSO_QTY_" & Format(CSO_ADDR_LNO, "000")) = DBNull.Value
                rowSOTCSTO3.Item("CSO_QTY_" & Format(CSO_COL, "000")) = DBNull.Value
            Else
                rowSOTCSTO2.Item("CSO_QTY_" & Format(CSO_ADDR_LNO, "000")) = qty
                rowSOTCSTO3.Item("CSO_QTY_" & Format(CSO_COL, "000")) = qty
            End If


            Return qty
        End Function
    End Class

    Private Sub chkAllowEditShipToAddress_CheckedChanged(sender As Object, e As EventArgs) Handles chkAllowEditShipToAddress.CheckedChanged
        If ScreenMode Then
            If chkAllowEditShipToAddress.Checked Then
                MsgBox("Your address changes will be used for this order only." & vbCrLf & vbCrLf & "The Master address records for RSCs will not be changed until COWORX is notified of the Change of Address." & vbCrLf & vbCrLf & "The Master address records for AEs and ACs must be processed by Sales Admin", MsgBoxStyle.OkOnly, "IMPORTANT - Please Note")
                chkHideAddressColumns.Enabled = False
            Else
                chkHideAddressColumns.Enabled = True
            End If
            Edit_Ship_To_Addresses()
        End If
    End Sub

    Private Sub grdSOTCSTOX_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles grdSOTCSTOX.InitializeLayout

    End Sub

    Private Sub grdSOTCSTOX_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTCSTOX.InitializeRow
        If e.Row.Band.Index = 0 Then
            Dim isSalesHold As Boolean = (e.Row.Cells("CSO_SALES_HOLD").Value & "" = "1")
            Dim isUrgent As Boolean = (e.Row.Cells("CSO_URGENT").Value & "" = "1")

            ' Determine the color based on both conditions
            If isSalesHold And isUrgent Then
                e.Row.Appearance.BackColor = System.Drawing.Color.Yellow
            ElseIf isSalesHold Then
                e.Row.Appearance.BackColor = System.Drawing.Color.Yellow
            ElseIf isUrgent Then
                e.Row.Appearance.BackColor = System.Drawing.Color.Orange
            Else
                e.Row.Appearance.BackColor = System.Drawing.Color.Empty
            End If
        End If

    End Sub
    Private Sub grdADDRCHNG_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles grdSOTCSTOA.InitializeRow
        Dim csColumns As String() = {"CS_NAME", "CS_ADDR1", "CS_ADDR2", "CS_ADDR3", "CS_CITY", "CS_STATE", "CS_ZIP_CODE", "CS_PHONE", "CS_EMAIL"}
        Dim dbColumns As String() = {"DB_NAME", "DB_ADDR1", "DB_ADDR2", "DB_ADDR3", "DB_CITY", "DB_STATE", "DB_ZIP_CODE", "DB_PHONE", "DB_EMAIL"}
        For i As Integer = 0 To csColumns.Length - 1
            Dim csColumn As String = csColumns(i)
            Dim dbColumn As String = dbColumns(i)
            Dim csValue As String = If(e.Row.Cells(csColumn).Value, "").ToString()
            Dim dbValue As String = If(e.Row.Cells(dbColumn).Value, "").ToString()
            If csValue <> dbValue Then
                e.Row.Cells(csColumn).Appearance.BackColor = System.Drawing.Color.Red
                e.Row.Cells(csColumn).Appearance.ForeColor = System.Drawing.Color.White
            Else
                e.Row.Cells(csColumn).Appearance.BackColor = System.Drawing.Color.Empty
            End If
        Next
    End Sub


    Private Sub optRSC_FM_ValueChanged(sender As Object, e As EventArgs) Handles optRSC_FM.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Dim RMAX As Integer = dst.Tables("SOTCSTO3").Rows.Count
        Dim R0 As Integer = r0T ' STARTING ROW OF ADDRESSES +1


        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting Row Visibility")

        WorkbookView1.GetLock()

        workbook = WorkbookView1.ActiveWorkbook

        ws = workbook.Worksheets(0)
        ws.Unprotect(XLS_PWD)

        Dim c As Integer = 3 ' cell of Type

        For R_INDEX As Integer = 1 To RMAX
            Dim R As Integer = R0 + R_INDEX
            With ws.Cells(R, c)
                'Debug.Print(.Value)
                If .Value & "" = "SDS" Or .Value & "" = "RSC" Or .Value & "" = "BM" Then
                    If .Value & "" = optRSC_FM.Value Or optRSC_FM.Value = "ALL" Then
                        .EntireRow.Hidden = False
                    Else
                        .EntireRow.Hidden = True
                    End If
                End If
            End With
        Next

        ws.Protect(XLS_PWD)
        WorkbookView1.ReleaseLock()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub cbeFindItemList_ValueChanged(sender As Object, e As EventArgs) Handles cbeFindItemList.ValueChanged
        Dim ITEM_CODE As String = cbeFindItemList.Value
        Find_Item(ITEM_CODE)
    End Sub

    Private Sub grdSOTCSTT1_BeforeRowUpdate(sender As Object, e As CancelableRowEventArgs) Handles grdSOTCSTT1.BeforeRowUpdate
        If e.Row.IsAddRow Then
            e.Row.Cells("SELL_CODE").Value = Absx1.txtFor("SELL_CODE").Text
        End If
    End Sub

    Private Sub grdSOTCSTT1_BeforeExitEditMode(sender As Object, e As BeforeExitEditModeEventArgs) Handles grdSOTCSTT1.BeforeExitEditMode
        Dim RSC_TAG As String = grdSOTCSTT1.ActiveRow.Cells("RSC_TAG").Text
        RSC_TAG = RSC_TAG.ToUpper
        If RSC_TAG <> "" Then
            For I As Integer = 1 To RSC_TAG.Length
                Dim X As String = Mid(RSC_TAG, I, 1)
                If (X < "A" Or X > "Z") And (X < "0" Or X > "9") Then
                    e.Cancel = True
                    ASCMAIN1.Progress("Use only Letters and Numbers (no spaces or punctuation) in a Tag")
                End If
            Next

            If e.Cancel Then
            Else
                grdSOTCSTT1.ActiveRow.Cells("RSC_TAG").Value = RSC_TAG
            End If


        End If
    End Sub

    Private Sub grdSOTCSTT1_AfterRowInsert(sender As Object, e As RowEventArgs) Handles grdSOTCSTT1.AfterRowInsert
        'Stop
    End Sub

    Private Sub grdSOTCSTT1_AfterRowUpdate(sender As Object, e As RowEventArgs) Handles grdSOTCSTT1.AfterRowUpdate

        Dim COL_NAME As String = "TAG_" & e.Row.Cells("RSC_TAG").Value
        If Not dst.Tables("SOTCSTTX").Columns.Contains(COL_NAME) Then
            Add_Column(COL_NAME)
        End If
    End Sub

    Private Sub chkEditTags_CheckedChanged(sender As Object, e As EventArgs) Handles chkEditTags.CheckedChanged
        Toggle_Editability()
    End Sub

    Sub Toggle_Editability()

        If chkEditTags.Checked Then

            With grdSOTCSTT1.DisplayLayout.Override
                .AllowAddNew = AllowAddNew.FixedAddRowOnTop
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.True
            End With
            With grdSOTCSTTX.DisplayLayout.Override
                .AllowAddNew = AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.True
                .AllowDelete = DefaultableBoolean.False
            End With

            For Each tab As UltraWinTabControl.UltraTab In tabMaster.Tabs
                If tab.Selected Then
                Else
                    tab.Enabled = False
                End If
            Next

            UltraExplorerBar1.Groups("Screen Control").Visible = False
            UltraExplorerBar1.Groups("RSC Tags").Visible = True

        Else


            With grdSOTCSTT1.DisplayLayout.Override
                .AllowAddNew = AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With
            With grdSOTCSTTX.DisplayLayout.Override
                .AllowAddNew = AllowAddNew.No
                .AllowUpdate = DefaultableBoolean.False
                .AllowDelete = DefaultableBoolean.False
            End With

            For Each tab As UltraWinTabControl.UltraTab In tabMaster.Tabs
                If tab.Selected Then
                Else
                    tab.Enabled = True
                End If
            Next

            UltraExplorerBar1.Groups("Screen Control").Visible = True
            UltraExplorerBar1.Groups("RSC Tags").Visible = False
        End If

    End Sub

    Private Sub grdSOTCSTT1_AfterRowActivate(sender As Object, e As EventArgs) Handles grdSOTCSTT1.AfterRowActivate
        With grdSOTCSTT1.DisplayLayout.Bands(0).Columns("RSC_TAG")
            If grdSOTCSTT1.ActiveRow.IsAddRow Then
                .CellActivation = Activation.AllowEdit
            Else
                .CellActivation = Activation.NoEdit
            End If
        End With

    End Sub

    Sub Update_RSC_Tags()
        BeginTrans()

        dst.Tables("SOTCSTT2").Rows.Clear()
        Dim RSC_TAGs As New List(Of String)
        For Each rowSOTCSTT1 As DataRow In dst.Tables("SOTCSTT1").Select("")
            Dim RSC_TAG As String = rowSOTCSTT1.Item("RSC_TAG")
            RSC_TAGs.Add(RSC_TAG)
        Next

        If RSC_TAGs.Count > 0 Then
            For Each rowSOTCSTTX As DataRow In dst.Tables("SOTCSTTX").Select("")
                Dim RSSP_CODE As String = rowSOTCSTTX.Item("RSSP_CODE") & ""
                For Each RSC_TAG As String In RSC_TAGs
                    If rowSOTCSTTX.Item("TAG_" & RSC_TAG) & "" = "1" Then

                        Dim rowSOTCSTT2 As DataRow = dst.Tables("SOTCSTT2").NewRow
                        rowSOTCSTT2.Item("SELL_CODE") = Absx1.txtFor("SELL_CODE").Text
                        rowSOTCSTT2.Item("RSSP_CODE") = RSSP_CODE
                        rowSOTCSTT2.Item("RSC_TAG") = RSC_TAG
                        dst.Tables("SOTCSTT2").Rows.Add(rowSOTCSTT2)
                    End If
                Next
            Next
        End If

        Update_Record_TDA("SOTCSTT2", $"SELL_CODE = '{Absx1.txtFor("SELL_CODE").Text}'")
        Update_Record_TDA("SOTCSTT1")
        chkEditTags.Checked = False

        CommitTrans("RSC Tags have been Updated")
    End Sub

    Private Sub UltraButton2_Click(sender As Object, e As EventArgs) Handles btnRefreshRSC_GROUPS_Click.Click
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Applying Filters")
        Refresh_CODEs()
        Hide_Untagged()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Private Sub Hide_Untagged()
        If Me.SELECTION_NO = 0 Then Exit Sub
        Dim RMAX As Integer = dst.Tables("SOTCSTO3").Rows.Count
        Dim R0 As Integer = r0T + 2 ' STARTING ROW OF ADDRESSES +3 to skip AE/AC
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Setting Row Visibility")
        WorkbookView1.GetLock()
        workbook = WorkbookView1.ActiveWorkbook
        ws = workbook.Worksheets(0)
        ws.Unprotect(XLS_PWD)

        ' Check if RSC_Tags is empty
        If RSC_Tags.Count = 0 Then
            ' Unhide all rows if there are no selected tags
            For R_INDEX As Integer = 1 To RMAX
                Dim R As Integer = R0 + R_INDEX
                ws.Cells(R, 4).EntireRow.Hidden = False
            Next
        Else
            Dim selectedTags As String = String.Join("','", RSC_Tags)
            For R_INDEX As Integer = 1 To RMAX
                Dim R As Integer = R0 + R_INDEX
                With ws.Cells(R, 4)
                    'everything starts off unhidden
                    .EntireRow.Hidden = False
                    Dim name As String = .Value

                    ' Check if name is Nothing or empty
                    If Not String.IsNullOrEmpty(name) Then
                        '' Escape single apostrophes in the name
                        'name = name.Replace("'", "''")

                        ' Get RSSP_CODE for the given name
                        ASCMAIN1.sql = $"SELECT RSSP_CODE FROM SPTRSSP1 WHERE RSSP_NAME = :PARM1 and RSSP_TITLE = 'RSC' and RSSP_STATUS = 'A'"
                        Dim code As String = ASCDATA1.GetDataValue(ASCMAIN1.sql, "V", New String() {name})
                        If Not String.IsNullOrEmpty(code) Then
                            ' Check if the RSC_TAG exists for the given RSSP_CODE in SOTCSTT2
                            ASCMAIN1.sql = $"SELECT COUNT(*) FROM SOTCSTT2 WHERE RSSP_CODE = :PARM1 AND RSC_TAG IN (:PARM2)"
                            Dim tagCount As Integer = ASCDATA1.GetDataValue(ASCMAIN1.sql, "VV", New String() {code, selectedTags})
                            .EntireRow.Hidden = (tagCount = 0)
                        Else
                            .EntireRow.Hidden = True
                        End If
                    End If
                End With
            Next
        End If

        ws.Protect(XLS_PWD)
        WorkbookView1.ReleaseLock()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Progress(msg As String)
        If msg = "" Then
            Me.Cursor = Cursors.Default
        Else
            Me.Cursor = Cursors.WaitCursor
        End If
        ASCMAIN1.Progress(msg, "")
    End Sub
    Private Sub Neg_Bal_Check(includeWarning As Boolean)
        workbook.WorkbookSet.GetLock()
        Try
            Dim DT As New DataTable
            DT.Columns.Add("ITEM_CODE", GetType(String))
            DT.Columns.Add("ITEM_DESC", GetType(String))
            DT.Columns.Add("BALANCE", GetType(Integer))

            Dim CSO_NO As String = UltraTextEditor1.Text
            ASCMAIN1.sql = $"SELECT COUNT(*) FROM SOTNGMSG WHERE CSO_NO = '{CSO_NO}'"
            Dim COUNT As Integer = Val(ASCDATA1.GetDataValue)
            Dim EXISTS As Boolean = COUNT > 0

            Dim entries As New List(Of String)
            For i As Integer = 19 To MAX_COLs
                If Not IsNothing(ws.Cells(8, i).Value) AndAlso Double.TryParse(ws.Cells(7, i).Value.ToString(), Nothing) Then
                    Dim BALANCE As Integer = Convert.ToInt32(ws.Cells(7, i).Value)
                    If BALANCE < 0 Then
                        Dim ITEM_CODE As String = ws.Cells(12, i).Value.ToString()
                        Dim ITEM_DESC As String = ws.Cells(11, i).Value.ToString()

                        Dim INIT_DATE As Date = DATETIME_STAMP
                        Dim INIT_OPER As String = ASCMAIN1.USER_ID
                        Dim SELL_CODE As String = UltraTextEditor4.Text
                        Dim QTY_ALLO As Integer = Convert.ToInt32(ws.Cells(2, i).Value)
                        Dim ALLO_DATE As Date = UltraDateTimeEditor1.Value.Date
                        Dim QTY_LEFT As Integer = Convert.ToInt32(ws.Cells(5, i).Value)
                        DT.Rows.Add(ITEM_CODE, ITEM_DESC, BALANCE)

                        ' Format the INIT_DATE using TO_DATE with proper quoting
                        Dim formattedInitDate As String = $"TO_DATE('{INIT_DATE:dd-MMM-yyyy HH:mm:ss}', 'DD-MON-YYYY HH24:MI:SS')"
                        Dim formattedAlloDate As String = $"TO_DATE('{ALLO_DATE:dd-MMM-yyyy}', 'DD-MON-YYYY')"

                        entries.Add($"('{CSO_NO}', {formattedInitDate}, '{INIT_OPER}', '{SELL_CODE}', '{ITEM_CODE}', {BALANCE}, {QTY_ALLO}, {formattedAlloDate}, {QTY_LEFT})")
                    End If
                End If
            Next

            If entries.Count > 0 And Not EXISTS Then
                Dim batchInsertCommand As New System.Text.StringBuilder("INSERT ALL")
                For Each entry As String In entries
                    batchInsertCommand.AppendLine()
                    batchInsertCommand.Append($"    INTO SOTNGMSG (CSO_NO, INIT_DATE, INIT_OPER, SELL_CODE, ITEM_CODE, BAL, QTY_ALLO, ALLO_DATE, QTY_LEFT) VALUES {entry}")
                Next
                batchInsertCommand.Append(" SELECT * FROM dual")

                ASCDATA1.ExecuteSQL(batchInsertCommand.ToString())
            End If



            Dim message As String = "The following Items have negative balances."
            If includeWarning Then
                message &= " DO NOT PROCEED!"
            End If

            Using F As New ASFMSGBF
                F.Show_grd(DT, Me, message)
            End Using
        Finally
            workbook.WorkbookSet.ReleaseLock()
        End Try
    End Sub
    Private Function Neg_Bal_Details() As String
        Dim EMsg As String = ""
        Dim HAS_NEG As Boolean = False
        workbook.WorkbookSet.GetLock()
        ws.Unprotect(XLS_PWD)
        For Each row As DataRow In dst.Tables("SOTCSTO2").Select("QTY_BAL < 0")
            ' Check if the total ordered quantity for this item is greater than zero
            Dim CSO_QTY_TOTAL As Integer = Convert.ToInt32(row("CSO_QTY_TOTAL"))
            If CSO_QTY_TOTAL > 0 Then
                HAS_NEG = True
                EMsg &= vbCr & $"Item Code: {row("ITEM_CODE")}, Balance: {row("QTY_BAL")}, Ordered Quantity: {CSO_QTY_TOTAL}"

                Dim ITEM_CODE As String = row("ITEM_CODE").ToString().Trim()
                Dim colIndex As Integer = -1
                For i As Integer = c0_Items To c0_Items + ws.UsedRange.Columns.Count
                    If ws.Cells(ROW_ITEM_CODE, i).Value.ToString().Trim() = ITEM_CODE Then
                        colIndex = i
                        Exit For
                    End If
                Next
            End If
        Next
        ws.Protect(XLS_PWD)
        workbook.WorkbookSet.ReleaseLock()
        If HAS_NEG Then
            Return "Items with Negative Balances after Car-Stock Order - Update Denied:" & vbCrLf & EMsg
        Else
            Return String.Empty
        End If
    End Function
    Private Function Has_Neg_Bal(Optional ByVal checkQuantities As Boolean = False) As Boolean
        workbook.WorkbookSet.GetLock()
        Try
            For i As Integer = 19 To MAX_COLs
                Dim hasQuantity As Boolean = False
                If checkQuantities Then
                    For rowIdx As Integer = 14 To ws.UsedRange.Rows.Count
                        If Not IsNothing(ws.Cells(rowIdx, i).Value) AndAlso Double.TryParse(ws.Cells(rowIdx, i).Value.ToString(), Nothing) Then
                            If Convert.ToDouble(ws.Cells(rowIdx, i).Value) > 0 Then
                                hasQuantity = True
                                Exit For
                            End If
                        End If
                    Next
                End If

                If Not checkQuantities OrElse hasQuantity Then
                    If Not IsNothing(ws.Cells(7, i).Value) AndAlso Double.TryParse(ws.Cells(7, i).Value.ToString(), Nothing) Then
                        Dim BALANCE As Double = Convert.ToDouble(ws.Cells(7, i).Value)
                        If BALANCE < 0 Then
                            Return True
                        End If
                    End If
                End If
            Next
        Finally
            workbook.WorkbookSet.ReleaseLock()
        End Try
        Return False ' Return false if no relevant negative balances are found
    End Function
    Private Sub Show_Neg_Popup()
        If Has_Neg_Bal() Then
            Neg_Bal_Check(True)
        End If
    End Sub
    Private Sub grdSOTALLOD_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdSOTALLOD.InitializeRow
        For Each ROW As DataRow In dst.Tables("SOTALLOI").Rows

        Next
        If e.Row.Band.Index = 1 Then  ' Only process item level
            If e.Row.Cells.Exists("DATE_END") AndAlso e.Row.Cells.Exists("BALANCE") Then
                Dim DATE_END As Date = Convert.ToDateTime(e.Row.Cells("DATE_END").Value)
                If DATE_END < Date.Today Then
                    Dim balance As Decimal = Convert.ToDecimal(e.Row.Cells("BALANCE").Value)
                    If balance > 0 Then
                        e.Row.Appearance.BackColor = System.Drawing.Color.LightGray  ' Color the row gray
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub btnRefresh_PROD_CODE_Click(sender As Object, e As EventArgs) Handles btnRefresh_PROD_CODE.Click
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Applying Filters")
        Refresh_CODEs()
        Apply_Filters()
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub
    Public Function AddBusinessDays(startDate As Date, businessDays As Integer) As Date
        Dim daysAdded As Integer = 0
        Dim currentDate As Date = startDate

        While daysAdded < businessDays
            currentDate = currentDate.AddDays(1)
            ' Check if currentDate is a weekday
            If currentDate.DayOfWeek <> DayOfWeek.Saturday AndAlso currentDate.DayOfWeek <> DayOfWeek.Sunday Then
                daysAdded += 1
            End If
        End While

        Return currentDate
    End Function

    Private Sub chkUrgent_CheckedChanged(sender As Object, e As EventArgs)
        'If chkUrgent.Checked AndAlso String.IsNullOrWhiteSpace(txtUrgent.Text) Then
        '    MessageBox.Show("Please provide a note for the urgent order.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        'End If
    End Sub
    Sub Urgent_Email()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing email")

        WorkbookView1.GetLock()
        Dim ws As SpreadsheetGear.IWorksheet = WorkbookView1.ActiveWorkbook.Worksheets(0)

        ws.Unprotect(XLS_PWD)
        ws.Range(8, 0).EntireRow.Hidden = False
        ws.Range(9, 0).EntireRow.Hidden = False
        ws.Range(10, 0).EntireRow.Hidden = False
        ws.Range(11, 0).EntireRow.Hidden = False
        chkAllowEditShipToAddress.Checked = False
        ws.Protect(XLS_PWD)
        Dim FILENAME As String = ASCMAIN1.Folders("Work") & Me.Name & "_" & CSO_NO & ".XLSX"
        WorkbookView1.ActiveWorkbook.SaveAs(FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

        If ASCMAIN1.Running_in_VS Then
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
        Else
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
            EMAIL_ADDRESSs.Add("highpriorityCSO@interparfums.com", ASCMAIN1.USER_NAME)
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        ATTACHMENTs.Add("Urgent Car-Stock Order", FILENAME)

        Dim EMAIL_SUBJECT As String = "URGENT: Car-Stock Order " & Absx1.txtFor("CSO_REF_NO").Text
        Dim URGENT_NOTES As String = rowSOTCSTO1("CSO_URGENT_NOTES").ToString()
        Dim DELIVER_BY As String = Convert.ToDateTime(rowSOTCSTO1("CSO_URGENT_DELIV_BY")).ToString("MM/dd/yy")
        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    EMAIL_SUBJECT, "CARSTOCK", True, False, CSO_NO, "CSO_NO", "Car-Stock Order", $"You are receiving this email because CSO {CSO_NO} was flagged as urgent." & vbCrLf &
                    $"Urgent Notes: {URGENT_NOTES}" & vbCrLf & $"Deliver By: {DELIVER_BY}")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        MsgBox("email Sent", MsgBoxStyle.OkOnly, "Verification")

        ws.Unprotect(XLS_PWD)
        ws.Range(8, 0).EntireRow.Hidden = True
        ws.Range(9, 0).EntireRow.Hidden = True
        ws.Range(10, 0).EntireRow.Hidden = True
        ws.Range(11, 0).EntireRow.Hidden = True
        ws.Protect(XLS_PWD)

        WorkbookView1.ReleaseLock()
    End Sub
    Function Update_Allocations() As String
        'WorkbookView1.GetLock()
        'ws.Unprotect(XLS_PWD)

        'Dim r0 As Integer = -1
        'Dim rT As Integer = ws.UsedRange.Rows.Count - r0
        'Dim c0 As Integer = c0_Items

        '' Step 1: Create a copy of the original SOTCSTO2 table
        'Dim tbl_SOTCSTO2_ORIG As DataTable = dst.Tables("SOTCSTO2").Copy()

        '' Step 2: Ensure that work tables are created/updated
        'Create_Work_Tables_SOTALLOX()
        'Fill_Records("SOTALLOX")

        '' Step 3: Update QTY_ALLO field in SOTCSTO2 with values from SOTALLOX
        'Dim DT As New DataTable
        'DT.Columns.Add("ITEM_CODE", GetType(String))
        'DT.Columns.Add("OLD_QTY_ALLO", GetType(Integer))
        'DT.Columns.Add("NEW_QTY_ALLO", GetType(Integer))

        'Dim errorMsg As String = String.Empty

        'For Each newRow As DataRow In dst.Tables("SOTCSTO2").Select("", "CSO_LNO")
        '    Dim CSO_NO As String = newRow.Item("CSO_NO")
        '    Dim CSO_LNO As Integer = newRow.Item("CSO_LNO")
        '    Dim ITEM_CODE As String = newRow.Item("ITEM_CODE").ToString().Trim()
        '    Dim rowSOTCSTO2_ORIG As DataRow = tbl_SOTCSTO2_ORIG.Rows.Find(New Object() {CSO_NO, CSO_LNO})
        '    Dim ALLO_CTL_NO As String = newRow.Item("ALLO_CTL_NO")
        '    Dim rowSOTALLOX As DataRow = dst.Tables("SOTALLOX").Rows.Find(ALLO_CTL_NO)

        '    If rowSOTALLOX IsNot Nothing Then
        '        Dim oldQtyAllo As Integer = Val(rowSOTCSTO2_ORIG.Item("QTY_ALLO") & "")
        '        Dim newQtyAllo As Integer = Val(rowSOTALLOX.Item("QTY_ALLO") & "")

        '        If oldQtyAllo <> newQtyAllo Then
        '            ' Update the QTY_ALLO field
        '            newRow.Item("QTY_ALLO") = newQtyAllo
        '            newRow.Item("QTY_LEFT") = Val(rowSOTALLOX.Item("QTY_LEFT") & "")

        '            ' Check for negative balance
        '            Dim oldBalance As Integer = oldQtyAllo - Val(rowSOTCSTO2_ORIG.Item("CSO_QTY_TOTAL") & "")
        '            Dim newBalance As Integer = newQtyAllo - Val(newRow.Item("CSO_QTY_TOTAL") & "")

        '            If newBalance < 0 Then
        '                ' Add change to the changes table
        '                DT.Rows.Add(ITEM_CODE, oldQtyAllo, newQtyAllo)
        '            End If

        '            ' Update the cell values only for those columns with changes
        '            Dim colIndex As Integer = c0_Items + CSO_LNO
        '            ws.Cells(r0 + 3, colIndex).Value = Val(newRow.Item("QTY_ALLO") & "") 'Qty Allo
        '            ws.Cells(r0 + 6, colIndex).Value = Val(newRow.Item("QTY_LEFT") & "") '#Left
        '        End If
        '    Else
        '        ' Allocation deleted for this item
        '        errorMsg = $"The allocation for an item involved in this CSO, ({ITEM_CODE}), has been deleted."
        '        Exit For
        '    End If
        'Next

        'ws.Protect(XLS_PWD)
        'WorkbookView1.ReleaseLock()

        'If Not String.IsNullOrEmpty(errorMsg) Then
        '    Return errorMsg
        'ElseIf DT.Rows.Count > 0 Then
        '    Using F As New ASFMSGBF
        '        F.Show_grd(DT, Me, "Note: The following items have been changed while you were in this form:")
        '    End Using
        'End If
        'Return String.Empty
    End Function

    Private Sub chkEvent_CheckedChanged(sender As Object, e As EventArgs) Handles chkEvent.CheckedChanged
        cmbEvent.Enabled = chkEvent.Checked
        cmbEvent.Visible = chkEvent.Checked
        txtEvent.Visible = chkEvent.Checked
    End Sub
End Class
