Imports System.Text.RegularExpressions
Imports System.Xml
Imports SpreadsheetGear

Public Class RSFRETL1

    Dim EDI_DOC_SEQ_NOs As New List(Of String)
    Dim RYP As String
    Dim RSTRETLA As String
    Dim EDT852TX As String
    Dim sqlEDT852TX As String
    Dim WKS() As String
    Dim WEEK_ENDING_DATEs() As Date
    Dim MAX_WEEKS As Integer = 0
    Dim rowARTCUST2 As DataRow
    Dim rowICTITEM1 As DataRow
    Dim CUST_CODE As String
    Dim EDT852TG As String = ""

    Dim EDI_DOC_SEQ_NOs_no_company As New List(Of String)

#Region "ABS Standard Routines" ' These Routines should be found in all Forms which Launch from the Menu.

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Get_PARM("EDTPARM1")
        Get_PARM("RSTPARM1")

        If ASCMAIN1.CLIENT = "SLP" Then
        Else
            optMode.ValueList.ValueListItems.Remove(5) ' SPS
            optMode.ValueList.ValueListItems.Remove(4) ' XLS
            optMode.ValueList.ValueListItems.Remove(3) ' 852
        End If
        tabMain.Tabs("852").Visible = (ASCMAIN1.CLIENT = "SLP")
        tabMain.Tabs("XLS").Visible = (ASCMAIN1.CLIENT = "SLP")
        tabMain.Tabs("SPS").Visible = (ASCMAIN1.CLIENT = "SLP")

        Set_cmbYP("RYP", ASCMAIN1.CYP, -36, 0, -1)
        Set_cmbYW("RYW", ASCMAIN1.CYW, -52 * 3, 0, -1)

        Prepare_EDI_Data()

        With dst

            If ASCMAIN1.CLIENT = "SLP" Then
                ASCMAIN1.sql = "Select * from RSTSLPI0 where PROCESS_IND = :PARM1 and DOC_SOURCE = :PARM2"
                Create_TDA(.Tables.Add, "RSTSLPI0", "**", 0, True, "VV", 1)

                ASCMAIN1.sql = "Select * from EDTTRPM1 where EDI_DOC_NO = '852'"
                Create_TDA(.Tables.Add, "EDTTRPM1", "**", 0, False)
                Fill_Records("EDTTRPM1")
            End If

            sqlEDT852TX = "Select EDT852T1.EDI_DOC_SEQ_NO" & vbCrLf _
            & ", EDT852T1.OPS_YYYYPP, EDT852T1.CUST_CODE, EDT852T1.COLLECTION_CODE" & vbCrLf _
            & " from EDT852T1" & vbCrLf _
            & " where EDT852T1.CUST_CODE = :PARM1" & vbCrLf _
            & " and EDT852T1.OPS_YYYYPP = :PARM2" & vbCrLf _
            & " and EDT852T1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'" & vbCrLf _
            & " and EDT852T1.EDI_STATUS = 'M'" & vbCrLf _
            & " and EDT852T1.DATA_LEVEL = :PARM4" & vbCrLf _
            & " and EDT852T1.BRAND_CODE = :PARM3" & vbCrLf
            'EDT852TX = ASCMAIN1.Temp_Table(sqlEDT852TX, , , "VVV")
            EDT852TX = ASCMAIN1.Temp_Table(Replace(Replace(Replace(Replace(sqlEDT852TX, ":PARM1", "Null"), ":PARM2", "Null"), ":PARM3", "Null"), ":PARM4", "Null"))
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TX & " Add Primary Key (EDI_DOC_SEQ_NO)")
            ASCMAIN1.sql = "Select * from " & EDT852TX & " EDT852TX"
            Create_TDA(.Tables.Add, "EDT852TX", "**", 0, False)

            Dim sqlR As String = "" _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM1,RSTRETL1.QTY_SOLD,0)) QTY_W1" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM2,RSTRETL1.QTY_SOLD,0)) QTY_W2" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM3,RSTRETL1.QTY_SOLD,0)) QTY_W3" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM4,RSTRETL1.QTY_SOLD,0)) QTY_W4" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM5,RSTRETL1.QTY_SOLD,0)) QTY_W5" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM6,RSTRETL1.QTY_SOLD,0)) QTY_W6" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM1,RSTRETL1.AMT_SOLD,0)) AMT_W1" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM2,RSTRETL1.AMT_SOLD,0)) AMT_W2" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM3,RSTRETL1.AMT_SOLD,0)) AMT_W3" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM4,RSTRETL1.AMT_SOLD,0)) AMT_W4" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM5,RSTRETL1.AMT_SOLD,0)) AMT_W5" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM6,RSTRETL1.AMT_SOLD,0)) AMT_W6" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM1,RSTRETL1.QTY_EOW,0)) ONH_W1" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM2,RSTRETL1.QTY_EOW,0)) ONH_W2" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM3,RSTRETL1.QTY_EOW,0)) ONH_W3" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM4,RSTRETL1.QTY_EOW,0)) ONH_W4" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM5,RSTRETL1.QTY_EOW,0)) ONH_W5" & vbCrLf _
            & ", SUM (DECODE(RSTRETL1.OPS_YYYYWW,:PARM6,RSTRETL1.QTY_EOW,0)) ONH_W6" & vbCrLf

            Dim sqlRSTRETLA = "Select RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & ", EDT852TX.COLLECTION_CODE, RSTRETL1.ITEM_CODE" & vbCrLf _
            & sqlR _
            & " from RSTRETL1," & EDT852TX & " EDT852TX" & vbCrLf _
            & " where RSTRETL1.EDI_DOC_SEQ_NO = EDT852TX.EDI_DOC_SEQ_NO" & vbCrLf _
            & " group by RSTRETL1.CUST_CODE, RSTRETL1.CUST_STORE_NO" & vbCrLf _
            & ", EDT852TX.COLLECTION_CODE, RSTRETL1.ITEM_CODE"

            ASCMAIN1.sql = "Select RSTRETLA.*" & vbCrLf _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_NAME) || DECODE(CUST_STORE_MARK_FOR,NULL,'',' (' || CUST_STORE_MARK_FOR || ')') CUST_STORE_LOCATION" & vbCrLf _
            & ", CUST_STORE_CITY, CUST_STORE_STATE, SELL_CODE" & vbCrLf _
            & " from (" & sqlRSTRETLA & ") RSTRETLA,ARTCUST2" & vbCrLf _
            & " where ARTCUST2.CUST_CODE = RSTRETLA.CUST_CODE" & vbCrLf _
            & "   and ARTCUST2.CUST_STORE_NO = RSTRETLA.CUST_STORE_NO"
            Create_TDA(.Tables.Add, "RSTRETLA", "**", 0, False, "VVVVVV")
            With .Tables("RSTRETLA")
                .Columns.Add("FLAG")
                .Columns.Add("CUST_ITEM_CODE")
                .Columns.Add("ITEM_DESC")
                .Columns.Add("QTY_TOTAL", GetType(System.Int32))
                .Columns.Add("AMT_TOTAL", GetType(System.Decimal))
                .Columns.Add("ONH_TOTAL", GetType(System.Decimal))
                For W As Integer = 1 To 6
                    .Columns("QTY_W" & Format(W, "0")).DataType = GetType(System.Int32)
                    .Columns("ONH_W" & Format(W, "0")).DataType = GetType(System.Int32)
                Next
            End With

            Dim sqlRSTRETLI = "Select RSTRETL1.CUST_CODE" & vbCrLf _
            & ", EDT852TX.COLLECTION_CODE, RSTRETL1.ITEM_CODE" & vbCrLf _
            & sqlR _
            & " from RSTRETL1," & EDT852TX & " EDT852TX" & vbCrLf _
            & " where RSTRETL1.EDI_DOC_SEQ_NO = EDT852TX.EDI_DOC_SEQ_NO" & vbCrLf _
            & " group by RSTRETL1.CUST_CODE" & vbCrLf _
            & ", EDT852TX.COLLECTION_CODE, RSTRETL1.ITEM_CODE"

            ASCMAIN1.sql = "Select RSTRETLI.*" & vbCrLf _
            & ", ICTITEM1.ITEM_DESC" & vbCrLf _
            & " from (" & sqlRSTRETLI & ") RSTRETLI,ICTITEM1" & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = RSTRETLI.ITEM_CODE"
            Create_TDA(.Tables.Add, "RSTRETLI", "**", 0, False, "VVVVVV")
            With .Tables("RSTRETLI")
                .Columns.Add("FLAG")
                .Columns.Add("CUST_ITEM_CODE")
                .Columns.Add("QTY_TOTAL", GetType(System.Int32))
                .Columns.Add("AMT_TOTAL", GetType(System.Decimal))
                .Columns.Add("ONH_TOTAL", GetType(System.Decimal))
                For W As Integer = 1 To 6
                    .Columns("QTY_W" & Format(W, "0")).DataType = GetType(System.Int32)
                    .Columns("ONH_W" & Format(W, "0")).DataType = GetType(System.Int32)
                Next
            End With

            ASCMAIN1.sql = "Select EDT852T0.* from EDT852T0,EDTUPCX1 " & vbCrLf _
            & " where EDT852T0.ITEM_CODE is Null" & vbCrLf _
            & " and EDT852T0.EDI_DOC_SEQ_NO = :PARM1" & vbCrLf _
            & " and EDTUPCX1.CUST_CODE (+) = EDT852T0.CUST_CODE" & vbCrLf _
            & " and EDTUPCX1.EDI_ITEM_CODE (+) = EDT852T0.EDI_ITEM_CODE" & vbCrLf _
            & " and NVL(EDTUPCX1.IGNORE,'0') <> '1'"
            Create_TDA(.Tables.Add, "EDT852T0_BADITEM", "**", 0, False, "V")

            ASCMAIN1.sql = "Select EDT852T1.*, ARTCUST1.CUST_NAME" _
            & ", '0' SELECTED " _
            & " from EDT852T1,ARTCUST1 " _
            & " where EDT852T1.EDI_STATUS = :PARM1 " _
            & " and ARTCUST1.CUST_CODE (+) = EDT852T1.CUST_CODE" _
            & " and EDT852T1.COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"

            Create_TDA(.Tables.Add, "EDT852T1", "**", 0, True, "V")


            ASCMAIN1.sql = "Select EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_ITEM_CODE, EDT852T0.ITEM_CODE, ICTITEM1.ITEM_DESC" _
            & ", COUNT(*) RECORDS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QA',NVL(EDT852T0.EDI_QTY,0),0)) QTY_ONH" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QS',NVL(EDT852T0.EDI_QTY,0),0)) QTY_SLS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QU',NVL(EDT852T0.EDI_QTY,0),0)) QTY_RTN" _
            & " from EDT852T0,ICTITEM1 " _
            & " where EDT852T0.EDI_DOC_SEQ_NO = :PARM1 " _
            & " and ICTITEM1.ITEM_CODE (+) = EDT852T0.ITEM_CODE" _
            & " group by EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_ITEM_CODE, EDT852T0.ITEM_CODE, ICTITEM1.ITEM_DESC"
            Create_TDA(.Tables.Add, "EDT852TI", "**", 0, False, "V")
            With .Tables("EDT852TI")
                .Columns.Add("ASIN")
            End With

            For Each dc As String In New String() {"RECORDS", "QTY_ONH", "QTY_SLS", "QTY_RTN"}
                dst.Tables("EDT852TI").Columns(dc).DataType = GetType(System.Int32)
            Next

            ASCMAIN1.sql = "Select EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_STORE_NO, EDT852T0.CUST_CODE" _
            & ", ARTCUST2.CUST_STORE_NO" _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_CITY) CUST_STORE_LOCATION" _
            & ", COUNT(*) RECORDS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QA',NVL(EDT852T0.EDI_QTY,0),0)) QTY_ONH" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QS',NVL(EDT852T0.EDI_QTY,0),0)) QTY_SLS" _
            & ", SUM (DECODE(EDT852T0.EDI_TRAN_TYPE,'QU',NVL(EDT852T0.EDI_QTY,0),0)) QTY_RTN" _
            & " from EDT852T0,ARTCUST2 " _
            & " where EDT852T0.EDI_DOC_SEQ_NO = :PARM1 " _
            & " and ARTCUST2.CUST_CODE (+) = EDT852T0.CUST_CODE" _
            & " and ARTCUST2.CUST_STORE_NO (+) = EDT852T0.CUST_STORE_NO" _
            & " group by EDT852T0.EDI_DOC_SEQ_NO" _
            & ", EDT852T0.EDI_STORE_NO, EDT852T0.CUST_CODE" _
            & ", ARTCUST2.CUST_STORE_NO" _
            & ", NVL(ARTCUST2.CUST_STORE_LOCATION,ARTCUST2.CUST_STORE_CITY)"
            Create_TDA(.Tables.Add, "EDT852TS", "**", 0, False, "V")

            For Each dc As String In New String() {"RECORDS", "QTY_ONH", "QTY_SLS", "QTY_RTN"}
                dst.Tables("EDT852TS").Columns(dc).DataType = GetType(System.Int32)
            Next

            Create_TDA(.Tables.Add, "EDT852T2", "*", 1)
            Create_TDA(.Tables.Add, "EDT852T3", "*", 1)

            .Relations.Add("EDT852T2_EDT852T3" _
                           , New DataColumn() _
                             { .Tables("EDT852T2").Columns("EDI_DOC_SEQ_NO") _
                              , .Tables("EDT852T2").Columns("EDI_LINE_NO")} _
                             , New DataColumn() _
                             { .Tables("EDT852T3").Columns("EDI_DOC_SEQ_NO") _
                              , .Tables("EDT852T3").Columns("EDI_LINE_NO")})

            Create_TDA(.Tables.Add, "EDTFILE1", "*")
            With .Tables("EDTFILE1").Columns
                .Add("EDI_JRNL_NO")
                .Add("EDI_SENDER_QUAL")
                .Add("EDI_SENDER_ID")
                .Add("EDI_ISA_CTL_NO")
                .Add("EDI_ISA_CTL_DATE", GetType(System.DateTime))
                .Add("DOC_EDI", GetType(System.Int32))
                .Add("DOC_852", GetType(System.Int32))
                .Add("NOTES")
            End With

            ASCMAIN1.sql = "Select CUST_CODE from ARTCUST1"
            Create_TDA(.Tables.Add, "EDT852TC", "**", 0, False)
            With .Tables("EDT852TC").Columns
                For I As Integer = 1 To 52 * 3 + 1 + 5 ' at least 3 years (which might include a 53 week year) + from the start of the current month (which might have 5 weeks)
                    .Add("F" & Format(I, "000"), GetType(System.Int32))
                Next
            End With

            ASCMAIN1.sql = "Select EDTJRNL3.*" _
            & " from EDTJRNL3 " _
            & " where EDI_DOC_SEQ_NO = :PARM1"
            Create_TDA(.Tables.Add, "EDTJRNL3", "**", 0, False, "V")

            Create_TDA(.Tables.Add, "EDTJRNL1", "*")
            Create_TDA(.Tables.Add, "EDTJRNL2", "*")

            .Relations.Add("EDTJRNL1_EDTJRNL2" _
                             , .Tables("EDTJRNL1").Columns("EDI_JRNL_NO") _
                             , .Tables("EDTJRNL2").Columns("EDI_JRNL_NO"))

            .Relations.Add("EDTJRNL2_EDTJRNL3" _
                           , New DataColumn() _
                             { .Tables("EDTJRNL2").Columns("EDI_JRNL_NO") _
                              , .Tables("EDTJRNL2").Columns("EDI_GS_NO")} _
                             , New DataColumn() _
                             { .Tables("EDTJRNL3").Columns("EDI_JRNL_NO") _
                              , .Tables("EDTJRNL3").Columns("EDI_GS_NO")})

            Create_TDA(.Tables.Add, "RSTRETL1", "*", 1)


            ASCMAIN1.sql = "Select RSTITEMX.* from RSTITEMX where CUST_CODE = :PARM1"
            Create_TDA(.Tables.Add, "RSTITEMX", "**", 0, True, "V")


            Create_TDA(.Tables.Add, "EDTUPCX1", "*")

            With .Tables.Add("RSTXLSQE")

                .Columns.Add("CUST_CODE", GetType(String))
                .Columns.Add("UPC_CODE_BAD", GetType(String))
                .PrimaryKey = New DataColumn() { .Columns("CUST_CODE"), .Columns("UPC_CODE_BAD")}

            End With

            'ASCMAIN1.sql = "Select * from EDTUPCX1 WHERE CUST_CODE = :PARM1"
            'Create_TDA(.Tables.Add, "EDTUPCX1", "**", 0, True, "V")

            ASCMAIN1.sql = "Select ITEM_UPC_CODE, ITEM_CODE from ICTITEM1"
            Create_TDA(.Tables.Add, "ICTITEMU", "**", 0, False, "", 1)


            ASCMAIN1.sql = "Select * from RSTASIN1"
            Create_TDA(.Tables.Add, "RSTASIN1", "**", 0, False, "", 2)
            Fill_Records("RSTASIN1")

            'If ASCMAIN1.CLIENT = "SLP" Then
            '    For Each TABLE_NAME As String In New String() {"EDT852T3", "EDT852T1", "EDT852T2", "EDTJRNL1"}
            '        If TABLE_NAME = "EDT852T3" Then
            '            Create_TDA(.Tables.Add, TABLE_NAME, "*")
            '            ' Create_BAs(TABLE_NAME) '  - must limit array size
            '        End If
            '    Next
            'End If

        End With

        ASCMAIN1.sql = "Select ICTITEM1.ITEM_CODE" _
        & ", ICTITEM1.ITEM_DESC, ICTITEM1.COLLECTION_CODE, ICTCOLL1.BRAND_CODE" _
        & " from ICTITEM1,ICTCOLL1 " _
        & " where ICTCOLL1.COLLECTION_CODE (+) = ICTITEM1.COLLECTION_CODE" _
        & " and ITEM_CODE = :PARM1"
        Create_ResultSet("ICTITEMX", "V")

        grdEDTFILE1.DataSource = dst.Tables("EDTFILE1")

        grdEDT852TC.DataSource = dst.Tables("EDT852TC")

        grdEDTJRNL1.DataSource = dst.Tables("EDTJRNL1")
        'grdEDTJRNL1.DataMember = "EDTJRNL1"
        'grdEDTJRNL1.DataSource = dst

        grdEDT852T1.DataSource = dst.Tables("EDT852T1")
        grdEDT852TI.DataSource = dst.Tables("EDT852TI")
        grdEDT852TS.DataSource = dst.Tables("EDT852TS")

        'grdEDT852T2.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.MultiBand
        grdEDT852T2.DataSource = dst.Tables("EDT852T2")
        'grdEDT852T2.DataMember = "EDT852T2"
        'grdEDT852T2.DataSource = dst

        grdRSTRETLA.DataSource = dst.Tables("RSTRETLA")
        grdRSTRETLI.DataSource = dst.Tables("RSTRETLI")

        Create_Summary(grdEDTFILE1, "EDI_FILENAME", "Count")
        Create_Summary(grdEDTFILE1, "DOC_EDI")
        Create_Summary(grdEDTFILE1, "DOC_852")

        Create_Summary(grdEDT852T1, "EDI_DOC_SEQ_NO", "Count")
        Create_Summary(grdEDT852T1, "SELECTED")

        Create_Summary(grdEDT852TC, "CUST_CODE", "Count")
        For W As Integer = 1 To 52 * 3 + 1 + 5
            Create_Summary(grdEDT852TC, "F" & Format(W, "000"))
        Next


        Create_Summary(grdEDT852TI, "EDI_ITEM_CODE", "Count")
        Create_Summary(grdEDT852TI, "RECORDS")
        Create_Summary(grdEDT852TI, "QTY_ONH")
        Create_Summary(grdEDT852TI, "QTY_SLS")
        Create_Summary(grdEDT852TI, "QTY_RTN")

        Create_Summary(grdEDT852TS, "EDI_STORE_NO", "Count")
        Create_Summary(grdEDT852TS, "RECORDS")
        Create_Summary(grdEDT852TS, "QTY_ONH")
        Create_Summary(grdEDT852TS, "QTY_SLS")
        Create_Summary(grdEDT852TS, "QTY_RTN")

        If ASCMAIN1.CLIENT = "SLP" Then
            grdRSTSLPI0.DataSource = dst.Tables("RSTSLPI0")
            Create_Summary(grdRSTSLPI0, "DOC_CTL_NO", "Count")
        End If

        For Each grd As UltraWinGrid.UltraGrid In New UltraWinGrid.UltraGrid() {grdRSTRETLA, grdRSTRETLI}
            If grd.Name = "grdRSTRETLA" Then
                Create_Summary(grd, "CUST_STORE_NO", "Count")
            Else
                Create_Summary(grd, "ITEM_CODE", "Count")
            End If
            For W As Integer = 0 To 6
                For Each T As String In New String() {"QTY", "AMT", "ONH"}
                    Create_Summary(grd, T & IIf(W = 0, "_TOTAL", "_W" & CStr(W)), , , "###,##0")
                Next
            Next
        Next

        With grdEDT852T1.DisplayLayout.Bands("EDT852T1")
            .Columns("EDI_DOC_SEQ_NO").Header.Fixed = True
        End With

        With grdEDTFILE1.DisplayLayout.Bands("EDTFILE1")
            .Columns("EDI_FILENAME").Header.Fixed = True
        End With

        With grdEDT852TC.DisplayLayout.Bands("EDT852TC")
            .Columns("CUST_CODE").Header.Fixed = True
        End With

        With grdRSTRETLA.DisplayLayout.Bands("RSTRETLA")
            .Columns("CUST_STORE_NO").Header.Fixed = True
            .Columns("CUST_STORE_LOCATION").Header.Fixed = True
        End With


        With grdEDT852T2.DisplayLayout.Bands("EDT852T2_EDT852T3")
            For I As Integer = 1 To 10
                .Columns("EDI_SDQ_STORE_" & Format(I, "00")).Header.Caption = "S" & Format(I, "00")
                .Columns("EDI_SDQ_QTY_AMT_" & Format(I, "00")).Header.Caption = "Q" & Format(I, "00")
                .Columns("EDI_SDQ_STORE_" & Format(I, "00")).Width = 45
                .Columns("EDI_SDQ_QTY_AMT_" & Format(I, "00")).Width = 45
                .Columns("EDI_SDQ_STORE_" & Format(I, "00")).CellAppearance.BackColor = System.Drawing.Color.Beige
            Next
        End With

        For Each gcol As UltraWinGrid.UltraGridColumn In grdEDT852T1.DisplayLayout.Bands(0).Columns
            If gcol.Key = "SELECTED" Then
                gcol.CellActivation = UltraWinGrid.Activation.AllowEdit
            Else
                gcol.CellActivation = UltraWinGrid.Activation.NoEdit
            End If
        Next

        With grdRSTRETLA.DisplayLayout.Bands("RSTRETLA")
            For Each COLUMN_NAME As String In New String() _
            {"CUST_CODE", "COLLECTION_CODE", "CUST_STORE_NO", "CUST_STORE_LOCATION", "CUST_STORE_CITY", "CUST_STORE_STATE", "SELL_CODE"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = System.Drawing.Color.Beige
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next

            .Columns("COLLECTION_CODE").EditorComponent = cbeCOLLECTION_CODE
            'ASCMAIN1.Add_Value_List(grdRSTRETLA, "COLLECTION_CODE")
        End With

        With grdRSTRETLI.DisplayLayout.Bands("RSTRETLI")
            For Each COLUMN_NAME As String In New String() _
            {"ITEM_DESC"}
                .Columns(COLUMN_NAME).CellAppearance.BackColor = System.Drawing.Color.Beige
                .Columns(COLUMN_NAME).CellActivation = UltraWinGrid.Activation.NoEdit
            Next

            .Columns("COLLECTION_CODE").EditorComponent = cbeCOLLECTION_CODE
            'ASCMAIN1.Add_Value_List(grdRSTRETLA, "COLLECTION_CODE")
        End With

        grdRSTRETLA.Dock = DockStyle.None
        grdRSTRETLA.Parent = tabManualDetails.Parent
        grdRSTRETLA.Dock = DockStyle.Fill

        grdRSTRETLI.Dock = DockStyle.None
        grdRSTRETLI.Parent = tabManualDetails.Parent
        grdRSTRETLI.Dock = DockStyle.Fill

        tabManualDetails.Visible = False

        With grdEDT852TC.DisplayLayout
            .Override.CellClickAction = UltraWinGrid.CellClickAction.EditAndSelectText
            .Bands(0).Columns("CUST_CODE").CellAppearance.BackColor = System.Drawing.Color.Beige
            .ColScrollRegions(0).ScrollColIntoView(.Bands(0).Columns(.Bands(0).Columns.Count - 1))
        End With

        If ASCMAIN1.CLIENT = "INT" Then
            cbeCustomLoad.DataSource = New String() {"Nordstrom"}
            optCI.Value = "I"
            optWM.Value = "W"
            chk1Wk.Checked = True
            chkByStore.Checked = True
        End If

        optMode.Value = "852 Data"

        ASCMAIN1.Add_Value_List(grdEDT852T1, "EDI_STATUS",, New String() {":", "0:Not Loaded", "1:Loaded", "D:Deleted", "M:Manual", "X:Voided", "A:Archived"})
    End Sub

    Overrides Sub Proceed_PreReq(ByVal eItemKey As String)

        EMsg = ""

        Select Case eItemKey

            Case "Load"

                Validate_Code("CUST_CODE")
                Validate_Code("BRAND_CODE")

                If chk1Wk.Checked Then
                    If Absx1.cmbFor("RYW").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify a Week"
                    End If
                Else
                    If Absx1.cmbFor("RYP").Value & "" = "" Then
                        EMsg &= vbCr & "You must Specify a Month"
                    End If
                End If

            Case "Update"

                If tabMain.SelectedTab.Key = "Manual Entry" Then
                    If Not (chkEntryUnits.Checked Or chkEntryDollars.Checked Or chkEntryOnHand.Checked) Then
                        EMsg &= "Must select Units, Dollars, or On Hand for load"
                    Else
                        Dim updFields As String = ""
                        If chkEntryUnits.Checked Then updFields &= "units,"
                        If chkEntryDollars.Checked Then updFields &= "dollars,"
                        If chkEntryOnHand.Checked Then updFields &= "on hand,"
                        Dim updMessage As String = String.Format("Update will load: {0}", updFields.Substring(0, updFields.Length - 1))
                        If MessageBox.Show(Me, updMessage, "Confirm Load", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = DialogResult.Cancel Then
                            Exit Sub
                        End If
                    End If
                End If
                If optCI.Value = "C" Then
                    For Each row In ASCDATA1.SelectDistinct("RSTRETLA", "COLLECTION_CODE").Rows
                        Dim COLLECTION_CODE As String = row.ITEM(0)
                        If LookUp("ICTITEM1", "DUMMY_" & COLLECTION_CODE) Is Nothing Then
                            EMsg &= vbCr & "No Dummy Item on file for Collection " & COLLECTION_CODE
                        Else
                            If cdr.Item("COLLECTION_CODE") & "" <> COLLECTION_CODE Then
                                EMsg &= vbCr & "Dummy Item for Collection " & COLLECTION_CODE & " is linked to " & cdr.Item("COLLECTION_CODE")
                            End If
                        End If
                    Next
                Else
                    Dim dtCI As DataTable = ASCDATA1.SelectDistinct _
                        (dst.Tables("RSTRETLI").Select("CUST_ITEM_CODE IS NOT NULL"), New String() {"CUST_ITEM_CODE", "ITEM_CODE"})

                    Dim CUST_ITEM_CODE_dups As String = ""

                    For Each row As DataRow In dtCI.Rows '  ASCDATA1.SelectDistinct(dtCI, New String() {"CUST_ITEM_CODE"}).Rows
                        Dim CUST_ITEM_CODE As String = row.Item("CUST_ITEM_CODE") & ""
                        Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""

                        If dst.Tables("RSTRETLI").Select("CUST_ITEM_CODE = '" & CUST_ITEM_CODE & "' and ITEM_CODE = '" & ITEM_CODE & "'").Count > 1 Then
                            CUST_ITEM_CODE_dups &= "," & CUST_ITEM_CODE
                        End If
                    Next

                    If CUST_ITEM_CODE_dups <> "" Then
                        EMsg &= vbCr & "Duplicate Customer Item Codes:" & vbCrLf & " " & Mid(CUST_ITEM_CODE_dups, 2)
                    End If

                    Dim TABLE_NAME As String = ""
                    If chkByStore.Checked Then
                        TABLE_NAME = "RSTRETLA"
                    Else
                        TABLE_NAME = "RSTRETLI"
                    End If

                    Me.Cursor = Cursors.WaitCursor
                    ASCMAIN1.Progress("Now Checking Items")

                    Dim dt As New DataTable
                    dt.Columns.Add("CUST_ITEM_ITEM_CODE")
                    dt.Columns.Add("ITEM_CODE")
                    dt.Columns.Add("COLLECTION_CODE")
                    dt.Columns.Add("BRAND_CODE")
                    Dim CUST_STORE_NO As String = ""
                    Dim INVALID_STORES As New List(Of String)
                    For Each rowRSTRETLI In dst.Tables(TABLE_NAME).Select("", "", DataViewRowState.CurrentRows)
                        rowRSTRETLI.Item("FLAG") = ""
                        Dim CUST_ITEM_CODE As String = rowRSTRETLI.Item("CUST_ITEM_CODE") & ""
                        Dim ITEM_CODE As String = rowRSTRETLI.Item("ITEM_CODE") & ""
                        ITEM_CODE = ITEM_CODE.ToUpper
                        rowRSTRETLI.Item("ITEM_CODE") = ITEM_CODE
                        Dim rowICTITEMX As DataRow = LookUp("ICTITEMX", ITEM_CODE, True)
                        Dim ITEM_CODEX As String = rowICTITEMX.Item("ITEM_CODE") & ""
                        Dim COLLECTION_CODE As String = rowICTITEMX.Item("COLLECTION_CODE") & ""
                        Dim BRAND_CODE As String = rowICTITEMX.Item("BRAND_CODE") & ""
                        rowRSTRETLI.Item("ITEM_DESC") = rowICTITEMX.Item("ITEM_DESC") & ""
                        rowRSTRETLI.Item("COLLECTION_CODE") = COLLECTION_CODE

                        If chkByStore.Checked Then
                            If CUST_STORE_NO <> rowRSTRETLI.Item("CUST_STORE_NO") & "" Then
                                CUST_STORE_NO = rowRSTRETLI.Item("CUST_STORE_NO") & ""
                                Dim rowARTCUST2 As DataRow = LookUp("ARTCUST2", New String() {HFs("CUST_CODE"), CUST_STORE_NO})
                                If rowARTCUST2 Is Nothing Then
                                    INVALID_STORES.Add(CUST_STORE_NO)
                                End If
                            End If
                        End If

                        If ITEM_CODEX = "" Then
                            rowRSTRETLI.Item("FLAG") = "I"
                            dt.Rows.Add(New String() {CUST_ITEM_CODE, ITEM_CODE, COLLECTION_CODE, BRAND_CODE})
                        ElseIf COLLECTION_CODE = "" Then
                            rowRSTRETLI.Item("FLAG") = "C"
                            dt.Rows.Add(New String() {CUST_ITEM_CODE, ITEM_CODE, COLLECTION_CODE, BRAND_CODE})
                        ElseIf BRAND_CODE <> HFs("BRAND_CODE") Then
                            'rowRSTRETLI.Item("FLAG") = "B"
                            'dt.Rows.Add(New String() {CUST_ITEM_CODE, ITEM_CODE, COLLECTION_CODE, BRAND_CODE})
                        End If
                    Next

                    Me.Cursor = Cursors.Default
                    ASCMAIN1.Progress("")

                    If dt.Rows.Count > 0 Then
                        EMsg &= vbCr & "Invalid Item Codes found - Correction is Required for Update"
                        Dim frmmsg As New ASFMSGBF
                        frmmsg.Show_grd(dt, Me, "Invalid Item Codes")
                        frmmsg = Nothing
                    End If

                    If INVALID_STORES.Count <> 0 Then
                        EMsg &= vbCr & "Invalid Stores found - Correction is Required for Update"
                        Dim frmmsg As New ASFMSGBF
                        frmmsg.Show_Formatted_txt("Invalid Stores", Join(INVALID_STORES.ToArray, vbCrLf), Me)
                        frmmsg = Nothing
                    End If

                End If

            Case "Import Raw EDI Files"
                If chkRawPreviouslyImported.Checked Then
                    EMsg &= vbCr & "You must first uncheck the option to Show Previously Imported files"
                End If

            Case "Load 852 Data", "Retract 852 Data", "Restore Deleted"

                'If ASCMAIN1.Running_in_VS Then
                '    Stop
                '    Using FRM As New TAFTESTX
                '        FRM.ShowDialog()
                '        Exit Sub
                '    End Using

                'End If

                Dim DOCS As Int32 = dst.Tables("EDT852T1").Select("SELECTED = '1'").Length
                If DOCS = 0 Then
                    EMsg &= vbCr & "No Documents Selected"
                Else
                    If MsgBox(eItemKey & " for " & CStr(DOCS) & " Documents", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
        End Select

        If EMsg <> "" Then
            MsgBox(EMsg, MsgBoxStyle.OkOnly, "Cannot Proceed")
            Exit Sub
        End If

        Call Proceed(eItemKey)

    End Sub

    Sub Proceed(ByVal eItemKey As String)

        Select Case eItemKey

            Case "Load"
                Load_Record()
                Mode_Settings(True)

            Case "Import Raw EDI Files"
                Mode_Settings(True)
                Import_Raw_EDI()
                Mode_Settings(False)

            Case "Load 852 Data", "Retract 852 Data", "Restore Deleted"
                Mode_Settings(True)
                Load_852_Data()
                Mode_Settings(False)

            Case "Update"

                Update_Record()
                Mode_Settings(False)

            Case "Cancel"
                Mode_Settings(False)

            Case "Orphan Stores"

                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("Now Searching for Orphan Stores in Retail Sales Data")

                ' LM 01/23/2018 Hi. Then can we clean out the orphan door list starting with January 2018?

                ASCMAIN1.sql = "" _
                    & "Select CUST_CODE, CUST_STORE_NO, COUNT (*) RECS" & vbCrLf _
                    & ", MIN (OPS_YYYYWW) WKMIN, MAX (OPS_YYYYWW) WKMAX, SUM (AMT_SOLD) AMT_SOLD, SUM (QTY_SOLD) QTY_SOLD" & vbCrLf _
                    & ", MIN (EDI_DOC_SEQ_NO) EDI_DOC_SEQ_NO1, MAX (EDI_DOC_SEQ_NO) EDI_DOC_SEQ_NO2" & vbCrLf _
                    & " from RSTRETL1 where (CUST_CODE, CUST_STORE_NO) in (" & vbCrLf _
                    & "Select DISTINCT CUST_CODE, CUST_STORE_NO from RSTRETL1 " & vbCrLf _
                    & IIf(ASCMAIN1.CLIENT = "INT", " where OPS_YYYYPP >= '201801'", "") _
                    & " minus  " & vbCrLf _
                    & "Select CUST_CODE, CUST_STORE_NO from ARTCUST2)" & vbCrLf _
                    & " group by CUST_CODE, CUST_STORE_NO" & vbCrLf _
                    & " order by CUST_CODE, CUST_STORE_NO"
                Dim tbl As DataTable = ASCDATA1.GetDataTable
                If tbl.Rows.Count = 0 Then
                    MsgBox("No Orphan Stores Found in Retail Sales Data", MsgBoxStyle.OkOnly, "Verification")
                Else
                    Dim frm As New ASFMSGBF
                    frm.Show_grd(tbl, Me, "Stores with Retail Sales Data, NOT in Store Master Table")
                End If

                Me.Cursor = Cursors.Default
                ASCMAIN1.Progress("")

        End Select

    End Sub

    Overrides Sub Mode_Settings(ByVal tf As Boolean, Optional ByVal MODE_description As String = "")

        Call Set_ScreenMode_Base(tf)

        If UltraExplorerBar1.Groups.Count > 0 Then
            With UltraExplorerBar1
                .Groups("Screen Control").Items("Load").Settings.Enabled = not_iScreenMode
                .Groups("Screen Control").Items("Update").Settings.Enabled = iScreenMode
                .Groups("Screen Control").Items("Cancel").Settings.Enabled = iScreenMode

                .Groups("Manual Entry").Visible = (tabMain.SelectedTab.Key = "Manual Entry") And ScreenMode
            End With

        End If

        Set_Read_Only(UltraGroupBox1, ScreenMode)

        'UltraTabControl1.Visible = tf

        optMode.Enabled = Not ScreenMode

        grdRSTRETLA.Visible = ScreenMode And chkByStore.Checked
        grdRSTRETLI.Visible = ScreenMode And Not chkByStore.Checked

        If ScreenMode Then
        Else
            Clear_Record()
            Setup_tabMain()
        End If

    End Sub

    Sub Clear_Record()
        Select_tabMain()

        EnforceConstraints(False)
        For Each TABLE_NAME As String In New String() _
            {"EDTJRNL1", "EDTJRNL2", "EDTJRNL3", "EDT852T1", "EDT852T2", "EDT852T3", "EDTUPCX1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next
        EnforceConstraints(True)

        grdRSTRETLI.Tag = ""
        Dim dvw As DataView = DirectCast(grdRSTRETLI.DataSource, DataTable).DefaultView
        dvw.RowFilter = grdRSTRETLI.Tag

        Prepare_852_Queue()
        Setup_tabRaw()
    End Sub

    Sub Load_Record()

        Call ASCMAIN1.Progress("Now Loading")

        If chkByStore.Checked Then
            Setup_grdRSTRETLA()
        Else
            Setup_grdRSTRETLI()
        End If

        Call Save_Header_Fields(UltraGroupBox1)
        EnforceConstraints(False)

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text
        Dim BRAND_CODE As String = Absx1.txtFor("BRAND_CODE").Text

        ASCMAIN1.sql = "Select COLLECTION_CODE, COLLECTION_NAME from ICTCOLL1 where BRAND_CODE = '" & Absx1.txtFor("BRAND_CODE").Text & "'"
        cbeCOLLECTION_CODE.DataSource = ASCDATA1.GetDataTable
        cbeCOLLECTION_CODE.Value = cbeCOLLECTION_CODE.Items(0)

        Dim grd As UltraWinGrid.UltraGrid
        If chkByStore.Checked Then
            grd = grdRSTRETLA
        Else
            grd = grdRSTRETLI
        End If

        Dim RYW As String = ""
        Dim RW As Integer = 0

        With grd
            If chk1Wk.Checked Then
                RYW = Absx1.cmbFor("RYW").Value
                Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
                RYP = rowGLTPARM3.Item("YYYYPP")
            Else
                RYP = Absx1.cmbFor("RYP").Value
            End If

            ReDim WKS(6)
            ReDim WEEK_ENDING_DATEs(6)
            ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYPP = '" & RYP & "'"
            For Each rowGLTPARM3 As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
                RW += 1
                Dim YW As String = rowGLTPARM3.Item("YYYYWW")
                WKS(RW) = YW
                WEEK_ENDING_DATEs(RW) = rowGLTPARM3.Item("WEEK_END_DATE")
                Dim LEGEND As String = ASCMAIN1.Get_Legend_Wk(YW, True) & vbCrLf & Format(rowGLTPARM3.Item("WEEK_END_DATE"), "MM/dd")
                With .DisplayLayout.Bands(0).Columns("QTY_W" & Format(RW, "0"))
                    .Header.Caption = LEGEND & " #"
                    .Header.Appearance.BackColor = System.Drawing.Color.Yellow
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Hidden = (optWM.Value = "M")
                    .Format = "###,##0"
                    .Width = 70
                End With
                With .DisplayLayout.Bands(0).Columns("AMT_W" & Format(RW, "0"))
                    .Header.Caption = LEGEND & " $"
                    .Header.Appearance.BackColor = System.Drawing.Color.LightGreen
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Hidden = (optWM.Value = "M")
                    .Format = "###,##0"
                    .Width = 70
                End With
                With .DisplayLayout.Bands(0).Columns("ONH_W" & Format(RW, "0"))
                    .Header.Caption = LEGEND & " H"
                    .Header.Appearance.BackColor = System.Drawing.Color.LightBlue
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Hidden = (optWM.Value = "M")
                    .Format = "###,##0"
                    .Width = 70
                End With
            Next
            MAX_WEEKS = RW

            With .DisplayLayout.Bands(0)
                With .Columns("QTY_TOTAL")
                    .Format = "###,##0"
                    .Header.Appearance.BackColor = System.Drawing.Color.Yellow
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Width = 90
                End With
                With .Columns("AMT_TOTAL")
                    .Format = "###,##0"
                    .Header.Appearance.BackColor = System.Drawing.Color.LightGreen
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Width = 90
                End With
                With .Columns("ONH_TOTAL")
                    .Format = "###,##0"
                    .Header.Appearance.BackColor = System.Drawing.Color.LightBlue
                    .Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal
                    .Width = 90
                End With
            End With

            If MAX_WEEKS < 6 Then
                For W As Integer = MAX_WEEKS + 1 To 6
                    With .DisplayLayout.Bands(0)
                        .Columns("QTY_W" & Format(W, "0")).Hidden = True
                        .Columns("AMT_W" & Format(W, "0")).Hidden = True
                        .Columns("ONH_W" & Format(W, "0")).Hidden = True
                    End With
                Next
            End If

            If optCI.Value = "I" Then
                optCOLLECTION_CODE.Value = "A"
                optCOLLECTION_CODE.Visible = False
                cbeCOLLECTION_CODE.Visible = False
                grdRSTRETLA.DisplayLayout.Bands(0).Columns("COLLECTION_CODE").Hidden = True
            Else
                optCOLLECTION_CODE.Visible = True
                cbeCOLLECTION_CODE.Visible = True
            End If


            With .DisplayLayout.Bands(0)
                If optCI.Value = "C" Then
                    .Columns("CUST_ITEM_CODE").Hidden = True
                    .Columns("ITEM_CODE").Hidden = True
                    .Columns("ITEM_CODE").CellAppearance.BackColor = System.Drawing.Color.Beige
                    .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                    .Columns("ITEM_CODE").Style = UltraWinGrid.ColumnStyle.Default
                    .Columns("COLLECTION_CODE").CellAppearance.BackColor = System.Drawing.Color.Empty
                    .Columns("COLLECTION_CODE").EditorComponent = cbeCOLLECTION_CODE
                    .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                Else
                    .Columns("CUST_ITEM_CODE").Hidden = False
                    .Columns("ITEM_CODE").Hidden = False
                    .Columns("ITEM_CODE").CellAppearance.BackColor = System.Drawing.Color.Empty
                    .Columns("ITEM_CODE").CellActivation = UltraWinGrid.Activation.AllowEdit
                    .Columns("ITEM_CODE").Style = UltraWinGrid.ColumnStyle.EditButton
                    .Columns("ITEM_CODE").ButtonDisplayStyle = UltraWinGrid.ButtonDisplayStyle.Always
                    .Columns("COLLECTION_CODE").CellAppearance.BackColor = System.Drawing.Color.Beige
                    .Columns("COLLECTION_CODE").EditorComponent = Nothing
                    .Columns("COLLECTION_CODE").CellActivation = UltraWinGrid.Activation.NoEdit
                End If
            End With

        End With

        Dim RSTx As String = IIf(chkByStore.Checked, "RSTRETLA", "RSTRETLI")

        If RSTx = "RSTRETLA" Then
            With dst.Tables(RSTx)
                If optCI.Value = "I" Then
                    .PrimaryKey = New DataColumn() { .Columns("CUST_CODE"), .Columns("CUST_STORE_NO"), .Columns("ITEM_CODE")}
                    .Columns("COLLECTION_CODE").AllowDBNull = True
                Else
                    .PrimaryKey = New DataColumn() { .Columns("CUST_CODE"), .Columns("CUST_STORE_NO"), .Columns("COLLECTION_CODE")}
                    .Columns("ITEM_CODE").AllowDBNull = True
                End If
            End With
        End If

        With dst.Tables(RSTx)
            For W As Integer = 1 To 6
                .Columns("QTY_W" & Format(W, "0")).Expression = ""
                .Columns("AMT_W" & Format(W, "0")).Expression = ""
                .Columns("ONH_W" & Format(W, "0")).Expression = ""
            Next
        End With

        ASCDATA1.ExecuteSQL("Truncate Table " & EDT852TX)
        If chk1Wk.Checked Then
            ASCDATA1.ExecuteSQL("Insert into " & EDT852TX & " " & Replace(sqlEDT852TX, "and EDT852T1.OPS_YYYYPP = :PARM2", "and EDT852T1.OPS_YYYYWW = :PARM2"), "VVVV", New Object() {CUST_CODE, Absx1.cmbFor("RYW").Value, BRAND_CODE, optCI.Value})
        Else
            ASCDATA1.ExecuteSQL("Insert into " & EDT852TX & " " & sqlEDT852TX, "VVVV", New Object() {CUST_CODE, RYP, BRAND_CODE, optCI.Value})
        End If
        'ASCDATA1.ExecuteSQL("Insert into " & EDT852TX & " " & sqlEDT852TX, "VVVV", CUST_CODE, RYP, BRAND_CODE, optCI.Value)

        'If Not chkByStore.Checked Then
        Fill_Records("RSTITEMX", Absx1.txtFor("CUST_CODE").Text)
        'End If

        Fill_Records("EDT852TX")
        Fill_Records(RSTx, New String() {WKS(1), WKS(2), WKS(3), WKS(4), WKS(5), WKS(6)})

        EnforceConstraints(True)

        With dst.Tables(RSTx)
            If optWM.Value = "W" Then
                .Columns("QTY_TOTAL").Expression = "ISNULL(QTY_W1,0)+ISNULL(QTY_W2,0)+ISNULL(QTY_W3,0)+ISNULL(QTY_W4,0)+ISNULL(QTY_W5,0)+ISNULL(QTY_W6,0)"
                .Columns("AMT_TOTAL").Expression = "ISNULL(AMT_W1,0)+ISNULL(AMT_W2,0)+ISNULL(AMT_W3,0)+ISNULL(AMT_W4,0)+ISNULL(AMT_W5,0)+ISNULL(AMT_W6,0)"
                .Columns("ONH_TOTAL").Expression = "ISNULL(ONH_W1,0)+ISNULL(ONH_W2,0)+ISNULL(ONH_W3,0)+ISNULL(ONH_W4,0)+ISNULL(ONH_W5,0)+ISNULL(ONH_W6,0)"


                For W As Integer = 1 To MAX_WEEKS - 1
                    .Columns("QTY_W" & Format(W, "0")).ReadOnly = False
                    .Columns("AMT_W" & Format(W, "0")).ReadOnly = False
                    .Columns("ONH_W" & Format(W, "0")).ReadOnly = False
                Next

            Else
                .Columns("QTY_TOTAL").Expression = ""
                .Columns("AMT_TOTAL").Expression = ""
                .Columns("ONH_TOTAL").Expression = ""
                .Columns("QTY_TOTAL").ReadOnly = False
                .Columns("AMT_TOTAL").ReadOnly = False
                .Columns("ONH_TOTAL").ReadOnly = False

                For Each row As DataRow In .Rows
                    With row
                        .Item("QTY_TOTAL") = Val(.Item("QTY_W1") & "") _
                                           + Val(.Item("QTY_W2") & "") _
                                           + Val(.Item("QTY_W3") & "") _
                                           + Val(.Item("QTY_W4") & "") _
                                           + Val(.Item("QTY_W5") & "") _
                                           + Val(.Item("QTY_W6") & "")
                        .Item("AMT_TOTAL") = Val(.Item("AMT_W1") & "") _
                                           + Val(.Item("AMT_W2") & "") _
                                           + Val(.Item("AMT_W3") & "") _
                                           + Val(.Item("AMT_W4") & "") _
                                           + Val(.Item("AMT_W5") & "") _
                                           + Val(.Item("AMT_W6") & "")
                        .Item("ONH_TOTAL") = Val(.Item("ONH_W" & Format(MAX_WEEKS, "0")) & "")
                    End With
                Next

                Dim SUM As String = ""
                For W As Integer = 1 To MAX_WEEKS - 1
                    .Columns("QTY_W" & Format(W, "0")).Expression = "ISNULL(QTY_TOTAL,0)/" & CStr(MAX_WEEKS)
                    .Columns("AMT_W" & Format(W, "0")).Expression = "ISNULL(AMT_TOTAL,0)/" & CStr(MAX_WEEKS)
                    .Columns("ONH_W" & Format(W, "0")).Expression = "ISNULL(ONH_W" & Format(W + 1, "0") & ",0) + ISNULL(QTY_W" & Format(W + 1, "0") & ",0)"
                    SUM &= "-ISNULL(QTY_W" & Format(W, "0") & ",0)"
                Next
                .Columns("QTY_W" & Format(MAX_WEEKS, "0")).Expression = "ISNULL(QTY_TOTAL,0)" & SUM
                .Columns("AMT_W" & Format(MAX_WEEKS, "0")).Expression = "ISNULL(AMT_TOTAL,0)" & Replace(SUM, "QTY_", "AMT_")
                .Columns("ONH_W" & Format(MAX_WEEKS, "0")).Expression = "ISNULL(ONH_TOTAL,0)"
            End If
        End With

        chkEntryDescription.Checked = (optCI.Value = "I")
        chkEntryDescription.Visible = (optCI.Value = "I")
        chkEntryDescription.Checked = False

        Setup_QA()
        If chkByStore.Checked Then
            Setup_grdRSTRETLA()
            Sort_grdColumns(grdRSTRETLA, "CUST_STORE_NO")
        Else
            Setup_grdRSTRETLI()
            Sort_grdColumns(grdRSTRETLA, "ITEM_CODE")
        End If

        Call ASCMAIN1.Progress("")
    End Sub

    Sub Update_Record()

        Call BeginTrans()

        For Each rowEDT852TX As DataRow In dst.Tables("EDT852TX").Rows
            Dim EDI_DOC_SEQ_NO As String = rowEDT852TX.Item("EDI_DOC_SEQ_NO")
            Update_RSTRETLx(EDI_DOC_SEQ_NO, "-")
            ASCDATA1.ExecuteSQL("Delete from RSTRETL1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
            ASCDATA1.ExecuteSQL("Delete from EDT852T1 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
        Next

        dst.Tables("EDT852TX").Rows.Clear()
        dst.Tables("EDT852T1").Rows.Clear()
        dst.Tables("RSTRETL1").Rows.Clear()

        dst.Tables("RSTRETL1").AcceptChanges()

        If optCI.Value = "I" Then
            If chkByStore.Checked Then
                For Each rowRSTRETLA As DataRow In dst.Tables("RSTRETLA").Select("COLLECTION_CODE IS NULL", "")
                    Dim ITEM_CODE As String = rowRSTRETLA.Item("ITEM_CODE")
                    cdr = LookUp("ICTITEM1", ITEM_CODE)
                    rowRSTRETLA.Item("COLLECTION_CODE") = cdr.Item("COLLECTION_CODE")
                Next
            Else
                'NOT NEC - HAPPENS IN PREREQ
                'For Each rowRSTRETLI As DataRow In dst.Tables("RSTRETLI").Select("COLLECTION_CODE IS NULL", "")
                '    Dim ITEM_CODE As String = rowRSTRETLI.Item("ITEM_CODE")
                '    cdr = LookUp("ICTITEM1", ITEM_CODE)
                '    rowRSTRETLI.Item("COLLECTION_CODE") = cdr.Item("COLLECTION_CODE")
                'Next
            End If
        End If

        Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

        If optCI.Value = "I" Then
            For Each row As DataRow In dst.Tables("RSTRETLI").Select("ITEM_CODE IS NOT NULL AND ITEM_CODE <> CUST_ITEM_CODE")
                Dim CUST_ITEM_CODE As String = row.Item("CUST_ITEM_CODE")
                Dim rowEDTUPCX1_x As DataRow = LookUp("EDTUPCX1", New String() {CUST_CODE, CUST_ITEM_CODE})
                Dim rowEDTUPCX1 = dst.Tables("EDTUPCX1").NewRow
                rowEDTUPCX1.Item("CUST_CODE") = CUST_CODE
                rowEDTUPCX1.Item("EDI_ITEM_CODE") = CUST_ITEM_CODE
                rowEDTUPCX1.Item("ITEM_CODE") = row.Item("ITEM_CODE")
                dst.Tables("EDTUPCX1").Rows.Add(rowEDTUPCX1)
                If rowEDTUPCX1_x IsNot Nothing Then
                    rowEDTUPCX1.AcceptChanges()
                    rowEDTUPCX1.SetModified()
                End If
            Next
            For Each row As DataRow In dst.Tables("RSTRETLA").Select("ITEM_CODE IS NOT NULL AND ITEM_CODE <> CUST_ITEM_CODE")
                Dim CUST_ITEM_CODE As String = row.Item("CUST_ITEM_CODE")
                Dim rowEDTUPCX1_x As DataRow = LookUp("EDTUPCX1", New String() {CUST_CODE, CUST_ITEM_CODE})
                Dim rowEDTUPCX1 As DataRow
                rowEDTUPCX1 = dst.Tables("EDTUPCX1").Rows.Find(New String() {CUST_CODE, CUST_ITEM_CODE})
                If rowEDTUPCX1 Is Nothing Then
                    rowEDTUPCX1 = dst.Tables("EDTUPCX1").NewRow
                    rowEDTUPCX1.Item("CUST_CODE") = CUST_CODE
                    rowEDTUPCX1.Item("EDI_ITEM_CODE") = CUST_ITEM_CODE
                    rowEDTUPCX1.Item("ITEM_CODE") = row.Item("ITEM_CODE")
                    dst.Tables("EDTUPCX1").Rows.Add(rowEDTUPCX1)
                    If rowEDTUPCX1_x IsNot Nothing Then
                        rowEDTUPCX1.AcceptChanges()
                        rowEDTUPCX1.SetModified()
                    End If
                End If
            Next
            Update_Record_TDA("EDTUPCX1")
        End If



        For RW As Integer = 1 To 6
            Dim RYW As String = WKS(RW)
            If RYW <> "" And (Not chk1Wk.Checked Or RYW = Absx1.cmbFor("RYW").Value) Then

                Dim COLLECTION_CODE As String = ""
                Dim EDI_DOC_SEQ_NO As String = ""

                Dim TABLE_NAME As String = ""
                If chkByStore.Checked Then
                    TABLE_NAME = "RSTRETLA"
                Else
                    TABLE_NAME = "RSTRETLI"
                End If

                If tabMain.SelectedTab.Key = "Manual Entry" Then
                    For Each row As DataRow In dst.Tables(TABLE_NAME).Select()
                        If Not chkEntryUnits.Checked Then row.Item("QTY_W" & Format(RW, "0")) = 0
                        If Not chkEntryDollars.Checked Then row.Item("AMT_W" & Format(RW, "0")) = 0
                        If Not chkEntryOnHand.Checked Then row.Item("ONH_W" & Format(RW, "0")) = 0
                    Next
                End If


                For Each row As DataRow In dst.Tables(TABLE_NAME) _
                .Select("AMT_W" & Format(RW, "0") & " <> 0" _
                  & " OR QTY_W" & Format(RW, "0") & " <> 0" _
                  & " OR ONH_W" & Format(RW, "0") & " <> 0", "COLLECTION_CODE")

                    If EDI_DOC_SEQ_NO = "" _
                    Or (COLLECTION_CODE <> row.Item("COLLECTION_CODE") And optCI.Value = "C") Then
                        COLLECTION_CODE = row.Item("COLLECTION_CODE")
                        EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDTJRNL3.EDI_DOC_SEQ_NO") 'This was causing a problem at INT because GENTRAN was assigning control numbers from its own table ("Lookup") -- so at INT we bumped this control number to 9000000000 to prevent collisions - MPR

                        Dim rowEDT852T1 As DataRow = dst.Tables("EDT852T1").NewRow
                        rowEDT852T1.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                        rowEDT852T1.Item("EDI_FROM_DATE") = WEEK_ENDING_DATEs(RW).AddDays(-6)
                        rowEDT852T1.Item("EDI_TO_DATE") = WEEK_ENDING_DATEs(RW)
                        rowEDT852T1.Item("EDI_STATUS") = "M"
                        rowEDT852T1.Item("OPS_YYYYPP") = RYP
                        rowEDT852T1.Item("OPS_YYYYWW") = RYW
                        rowEDT852T1.Item("CUST_CODE") = CUST_CODE
                        rowEDT852T1.Item("INIT_DATE") = DATETIME_STAMP
                        rowEDT852T1.Item("INIT_OPER") = ASCMAIN1.USER_ID
                        If optCI.Value = "C" Then
                            rowEDT852T1.Item("COLLECTION_CODE") = COLLECTION_CODE
                        End If
                        rowEDT852T1.Item("BRAND_CODE") = HFs("BRAND_CODE")
                        rowEDT852T1.Item("DATA_LEVEL") = optCI.Value
                        rowEDT852T1.Item("COMPANY_CODE") = ASCMAIN1.DBS_COMPANY
                        dst.Tables("EDT852T1").Rows.Add(rowEDT852T1)
                    End If

                    Dim ITEM_CODE As String = ""
                    If optCI.Value = "C" Then
                        ITEM_CODE = "DUMMY_" & COLLECTION_CODE
                    Else
                        ITEM_CODE = row.Item("ITEM_CODE")
                    End If

                    Dim CUST_STORE_NO As String = ""
                    If chkByStore.Checked Then
                        CUST_STORE_NO = row.Item("CUST_STORE_NO")
                    Else
                        CUST_STORE_NO = "000000"
                    End If


                    Dim rowRSTRETL1 As DataRow
                    rowRSTRETL1 = dst.Tables("RSTRETL1").Rows.Find _
                    (New String() {EDI_DOC_SEQ_NO, CUST_CODE, CUST_STORE_NO, ITEM_CODE})

                    If rowRSTRETL1 Is Nothing Then
                        rowRSTRETL1 = dst.Tables("RSTRETL1").NewRow
                        With rowRSTRETL1
                            .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                            .Item("CUST_CODE") = CUST_CODE
                            .Item("CUST_STORE_NO") = CUST_STORE_NO
                            .Item("ITEM_CODE") = ITEM_CODE

                            .Item("QTY_SOLD") = row.Item("QTY_W" & Format(RW, "0"))
                            .Item("AMT_SOLD") = row.Item("AMT_W" & Format(RW, "0"))
                            .Item("QTY_EOW") = row.Item("ONH_W" & Format(RW, "0"))
                            .Item("OPS_YYYYPP") = RYP
                            .Item("OPS_YYYYWW") = RYW
                            'rowRSTRETL1.Item("COLLECTION_CODE") = COLLECTION_CODE
                            dst.Tables("RSTRETL1").Rows.Add(rowRSTRETL1)
                        End With
                    Else
                        With rowRSTRETL1
                            .Item("QTY_SOLD") = Val(.Item("QTY_SOLD") & "") + Val(row.Item("QTY_W" & Format(RW, "0")) & "")
                            .Item("AMT_SOLD") = Val(.Item("AMT_SOLD") & "") + Val(row.Item("AMT_W" & Format(RW, "0")) & "")
                            .Item("QTY_EOW") = Val(.Item("QTY_EOW") & "") + Val(row.Item("ONH_W" & Format(RW, "0")) & "")
                        End With
                    End If
                Next
            End If
        Next

        If Not chkByStore.Checked Then
            For Each rowRSTRETLI As DataRow In dst.Tables("RSTRETLI").Select("", "", DataViewRowState.CurrentRows)
                Dim CUST_ITEM_CODE As String = Trim(rowRSTRETLI.Item("CUST_ITEM_CODE") & "")
                If CUST_ITEM_CODE <> "" Then
                    Dim ITEM_CODE As String = Trim(rowRSTRETLI.Item("ITEM_CODE") & "")
                    Dim rowRSTITEMX As DataRow = dst.Tables("RSTITEMX").Rows.Find _
                    (New String() {CUST_CODE, CUST_ITEM_CODE})
                    If rowRSTITEMX Is Nothing Then
                        dst.Tables("RSTITEMX").Rows.Add _
                        (New String() {CUST_CODE, CUST_ITEM_CODE, ITEM_CODE})
                    Else
                        If rowRSTITEMX.Item("ITEM_CODE") & "" <> ITEM_CODE Then
                            rowRSTITEMX.Item("ITEM_CODE") = ITEM_CODE
                        End If
                    End If
                End If
            Next
        Else
            For Each rowRSTRETLA As DataRow In dst.Tables("RSTRETLA").Select("", "", DataViewRowState.CurrentRows)
                Dim CUST_ITEM_CODE As String = Trim(rowRSTRETLA.Item("CUST_ITEM_CODE") & "")
                If CUST_ITEM_CODE <> "" Then
                    Dim ITEM_CODE As String = Trim(rowRSTRETLA.Item("ITEM_CODE") & "")
                    Dim rowRSTITEMX As DataRow = dst.Tables("RSTITEMX").Rows.Find _
                    (New String() {CUST_CODE, CUST_ITEM_CODE})
                    If rowRSTITEMX Is Nothing Then
                        dst.Tables("RSTITEMX").Rows.Add _
                        (New String() {CUST_CODE, CUST_ITEM_CODE, ITEM_CODE})
                    Else
                        If rowRSTITEMX.Item("ITEM_CODE") & "" <> ITEM_CODE Then
                            rowRSTITEMX.Item("ITEM_CODE") = ITEM_CODE
                        End If
                    End If
                End If
            Next
        End If
        Update_Record_TDA("RSTITEMX")

        Update_Record_TDA("EDT852T1")
        Update_Record_TDA("RSTRETL1")

        For Each rowEDT852T1 As DataRow In dst.Tables("EDT852T1").Rows
            Dim EDI_DOC_SEQ_NO As String = rowEDT852T1.Item("EDI_DOC_SEQ_NO")
            Dim OPS_YYYYPP As String = rowEDT852T1.Item("OPS_YYYYPP")

            If optCI.Value = "I" Then
                ASCMAIN1.sql = "Begin Declare Cursor C1 is " _
                & " Select ICTITEM1.ITEM_CODE" _
                & ", NVL(ICTRETLA.ITEM_RETAIL_PRICE,ICTITEM1.ITEM_RETAIL_PRICE) ITEM_RETAIL_PRICE" _
                & " from ICTITEM1,ICTRETLA" _
                & " where ICTRETLA.ITEM_CODE (+) = ICTITEM1.ITEM_CODE " _
                & " and ICTRETLA.OPS_YYYYPP (+) = '" & OPS_YYYYPP & "'" _
                & " and ICTITEM1.ITEM_CODE in " _
                & " (Select Distinct ITEM_CODE from RSTRETL1 " _
                & "  where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
                & "    and NVL(QTY_SOLD,0) <> 0 and NVL(AMT_SOLD,0) = 0);" _
                & " Begin For R1 in C1 Loop" _
                & "  Update RSTRETL1 Set AMT_SOLD = QTY_SOLD * R1.ITEM_RETAIL_PRICE" _
                & "   where ITEM_CODE = R1.ITEM_CODE and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'; " _
                & " End Loop; End; End;"
                ASCDATA1.ExecuteSQL()
            End If

            Update_RSTRETLx(EDI_DOC_SEQ_NO, "+")
        Next

        Call CommitTrans("Update Complete")

    End Sub

    Overrides Sub Excel_Import_Pre_Process _
    (ByVal grd As UltraWinGrid.UltraGrid,
     Optional ByRef load_by_table As Boolean = False,
     Optional ByRef load_handled As Boolean = False,
     Optional ByRef F As ASFEXCL1 = Nothing)

        rowARTCUST2 = Nothing
        CUST_CODE = Absx1.txtFor("CUST_CODE").Text

        If grd.Name = "grdRSTRETLI" Then
            load_by_table = True
        End If
        If grd.Name = "grdRSTRETLA" Then
            load_by_table = True
        End If
        If grd.Name = "grdASFBASEX" Then

            Dim TBL_BAD_ROWS As New DataTable
            With TBL_BAD_ROWS
                .Columns.Add("ITEM_CODE")
            End With

            dst.Tables("RSTRETLI").Rows.Clear()

            Dim GOOD_ROWS As Integer = 0
            Dim BAD_ROWS As Integer = 0
            Dim ITEM_CODE As String = ""
            Dim COLLECTION_CODE As String = ""
            Dim BAD_ITEM_CODEs As String = ""

            Dim RYW As String = Absx1.cmbFor("RYW").Value
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
            RYP = rowGLTPARM3.Item("YYYYPP")
            Dim REL_WEEK As Integer = Val(rowGLTPARM3.Item("REL_WEEK") & "")

            For Each row As DataRow In F.dt.Rows
                Dim A1 As String = row.Item(0) & ""
                If A1.Length > 10 And A1.Contains("D:") And A1.Contains("V:") And A1.Contains("S:") Then
                    Dim CUST_ITEM_CODE As String = Trim(Split(Split(A1, "S:")(1), "000:")(0))
                    ITEM_CODE = ""
                    COLLECTION_CODE = ""
                    Dim QTY_Wx As Integer = Val(row.Item(8) & "")
                    Dim ONH_Wx As Integer = Val(row.Item(12) & "")
                    If QTY_Wx <> 0 Or ONH_Wx <> 0 Then
                        Dim rowEDTUPCX1 As DataRow = Nothing
                        Dim rowICTITEM1 As DataRow = LookUp("ICTITEM1", CUST_ITEM_CODE)
                        If rowICTITEM1 Is Nothing Then
                            rowEDTUPCX1 = LookUp("EDTUPCX1", New String() {CUST_CODE, CUST_ITEM_CODE})
                            If rowEDTUPCX1 IsNot Nothing Then
                                rowICTITEM1 = LookUp("ICTITEM1", rowEDTUPCX1.Item("ITEM_CODE"))
                            End If
                        End If
                        If rowICTITEM1 IsNot Nothing Then
                            ITEM_CODE = rowICTITEM1.Item("ITEM_CODE")
                            COLLECTION_CODE = rowICTITEM1.Item("COLLECTION_CODE")
                        End If

                        Dim rowRSTRETLI As DataRow = dst.Tables("RSTRETLI").NewRow
                        rowRSTRETLI.Item("CUST_CODE") = CUST_CODE
                        rowRSTRETLI.Item("CUST_ITEM_CODE") = CUST_ITEM_CODE
                        rowRSTRETLI.Item("COLLECTION_CODE") = COLLECTION_CODE
                        rowRSTRETLI.Item("ITEM_CODE") = ITEM_CODE
                        rowRSTRETLI.Item("QTY_W" & CStr(REL_WEEK)) = QTY_Wx
                        rowRSTRETLI.Item("ONH_W" & CStr(REL_WEEK)) = ONH_Wx
                        dst.Tables("RSTRETLI").Rows.Add(rowRSTRETLI)

                        If ITEM_CODE = "" Then
                            BAD_ITEM_CODEs &= vbCrLf & CUST_ITEM_CODE
                            BAD_ROWS += 1
                            If rowEDTUPCX1 Is Nothing Then
                                Dim rowBAD_ROWS As DataRow = TBL_BAD_ROWS.NewRow
                                rowBAD_ROWS("ITEM_CODE") = CUST_ITEM_CODE
                                TBL_BAD_ROWS.Rows.Add(rowBAD_ROWS)
                            End If
                        Else
                            GOOD_ROWS += 1
                        End If
                    End If
                End If
            Next

            If BAD_ROWS = 0 Then
                MsgBox(CStr(GOOD_ROWS) & " Loaded", MsgBoxStyle.OkOnly, "Verification")
            Else
                MsgBox("Good Rows: " & CStr(GOOD_ROWS) & vbCrLf & "Bad Rows: " & CStr(BAD_ROWS), MsgBoxStyle.OkOnly, "Custom Load Complete")

                If TBL_BAD_ROWS.Rows.Count <> 0 Then
                    Dim Fmsg As New ASFMSGBF
                    Fmsg.Show_grd(TBL_BAD_ROWS, ASCMAIN1.ActiveForm, "Bad Items from Spreadsheet")

                    'If MsgBox("OK to Load Bad Items into Item Cross Reference?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.Yes Then
                    '    dst.Tables("EDTUPCX1").Rows.Clear()
                    '    For Each ITEM_CODE In Split(Mid(BAD_ITEM_CODEs, 3), vbCrLf)
                    '        Dim rowEDTUPCX1 As DataRow = dst.Tables("EDTUPCX1").NewRow
                    '        rowEDTUPCX1.Item("CUST_CODE") = CUST_CODE
                    '        rowEDTUPCX1.Item("EDI_ITEM_CODE") = ITEM_CODE
                    '        'rowEDTUPCX1.Item("ITEM_UPC_CODE") = ""
                    '        'rowEDTUPCX1.Item("ITEM_CODE") = ""
                    '        rowEDTUPCX1.Item("IGNORE") = "0"
                    '        dst.Tables("EDTUPCX1").Rows.Add(rowEDTUPCX1)
                    '    Next
                    '    Update_Record_TDA("EDTUPCX1")
                    'End If
                End If

            End If

            load_handled = True
        End If
    End Sub

    Overrides Sub Excel_Import_Post_Process(ByVal grd As UltraWinGrid.UltraGrid, F As ASFEXCL1)

        Select Case grd.Name
            Case "grdRSTRETLI"
                Dim CUST_CODE As String = Absx1.txtFor("CUST_CODE").Text

                For I As Integer = dst.Tables("RSTRETLI").Rows.Count - 1 To 0 Step -1
                    Dim rowRSTRETLI As DataRow = dst.Tables("RSTRETLI").Rows(I)
                    If rowRSTRETLI.RowState <> DataRowState.Deleted Then
                        If Val(rowRSTRETLI.Item("QTY_TOTAL") & "") = 0 And Val(rowRSTRETLI.Item("AMT_TOTAL") & "") = 0 And Val(rowRSTRETLI.Item("ONH_TOTAL") & "") = 0 Then
                            rowRSTRETLI.Delete()
                        End If
                    End If
                Next

                For Each rowRSTRETLI As DataRow In dst.Tables("RSTRETLI").Select("", "", DataViewRowState.CurrentRows)
                    rowRSTRETLI.Item("CUST_CODE") = CUST_CODE
                    Dim ITEM_CODE As String = Trim(rowRSTRETLI.Item("CUST_ITEM_CODE") & "").ToUpper
                    Dim rowICTITEM1 As DataRow = Nothing

                    If ITEM_CODE <> "" Then
                        rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                        If rowICTITEM1 Is Nothing Then
                            ITEM_CODE = ""
                        End If
                    End If

                    If ITEM_CODE = "" Then
                        Dim CUST_ITEM_CODE As String = Trim(rowRSTRETLI.Item("CUST_ITEM_CODE") & "").ToUpper

                        If CUST_ITEM_CODE <> "" Then
                            Dim rowRSTITEMX As DataRow = LookUp("RSTITEMX", New String() {CUST_CODE, CUST_ITEM_CODE})
                            If rowRSTITEMX IsNot Nothing AndAlso rowRSTITEMX.Item("ITEM_CODE") & "" <> "" Then
                                rowICTITEM1 = LookUp("ICTITEM1", rowRSTITEMX.Item("ITEM_CODE") & "", True)
                                If rowICTITEM1 IsNot Nothing Then
                                    ITEM_CODE = rowRSTITEMX.Item("ITEM_CODE") & ""
                                End If
                            End If

                            If ITEM_CODE = "" Then
                                rowICTITEM1 = LookUp("ICTITEM1", CUST_ITEM_CODE)
                                If rowICTITEM1 IsNot Nothing Then
                                    ITEM_CODE = CUST_ITEM_CODE
                                Else
                                    ASCMAIN1.sql = "Select ITEM_CODE from ICTITEM1 " _
                                    & " where ITEM_CODE like " _
                                    & "'" & Mid$(CUST_ITEM_CODE, 1, Len(CUST_ITEM_CODE) - 1) & "%" & Mid(CUST_ITEM_CODE, Len(CUST_ITEM_CODE), 1) & "'"
                                    Dim tbl As DataTable = ASCDATA1.GetDataTable
                                    If tbl.Rows.Count = 1 Then
                                        ITEM_CODE = tbl.Rows(0).Item(0)
                                        rowICTITEM1 = LookUp("ICTITEM1", ITEM_CODE)
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If ITEM_CODE <> "" Then
                        rowRSTRETLI.Item("ITEM_CODE") = ITEM_CODE

                        rowRSTRETLI.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC")
                        rowRSTRETLI.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE")
                    End If
                Next

        End Select
    End Sub

    Overrides Sub Excel_Import_DataTable_Intitialization _
    (ByRef dt As DataTable)
        dt.Columns(0).DataType = GetType(System.String)
    End Sub
    Overrides Sub Excel_Import_Custom_Processing_row _
    (ByVal row As DataRow, ByVal grow As UltraWinGrid.UltraGridRow,
     Optional ByVal tbl As DataTable = Nothing)

        If chkByStore.Checked Then
            row.Item("CUST_CODE") = CUST_CODE
            Dim CUST_STORE_NO As String = grow.Cells(0).Text.PadLeft(6, "0")
            If rowARTCUST2 Is Nothing OrElse rowARTCUST2.Item("CUST_STORE_NO") & "" <> CUST_STORE_NO Then
                rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO}, True)
            End If
            row.Item("CUST_STORE_NO") = CUST_STORE_NO
            Dim CUST_STORE_LOCATION As String = rowARTCUST2.Item("CUST_STORE_LOCATION") & ""
            If CUST_STORE_LOCATION = "" Then
                CUST_STORE_LOCATION = rowARTCUST2.Item("CUST_STORE_NAME") & ""
            End If
            row.Item("CUST_STORE_LOCATION") = CUST_STORE_LOCATION ' rowARTCUST2.Item("CUST_STORE_LOCATION") & ""
            row.Item("CUST_STORE_CITY") = rowARTCUST2.Item("CUST_STORE_CITY") & ""
            row.Item("CUST_STORE_STATE") = rowARTCUST2.Item("CUST_STORE_STATE") & ""
            row.Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE") & ""

            If optCI.Value = "C" Then
                row.Item("COLLECTION_CODE") = cbeCOLLECTION_CODE.Value
            Else
                Dim CUST_ITEM_CODE As String = grow.Cells(3).Text
                Dim ITEM_CODE As String = grow.Cells(4).Text
                If ITEM_CODE = "" Then
                    Dim rowRSTITEMX As DataRow = LookUp("RSTITEMX", New String() {CUST_CODE, CUST_ITEM_CODE})
                    If rowRSTITEMX IsNot Nothing AndAlso rowRSTITEMX.Item("ITEM_CODE") & "" <> "" Then
                        rowICTITEM1 = LookUp("ICTITEM1", rowRSTITEMX.Item("ITEM_CODE") & "", True)
                        'If rowICTITEM1.Item("COLLECTION_CODE") & "" = "" Then Stop
                        ITEM_CODE = rowRSTITEMX.Item("ITEM_CODE") & ""
                    Else
                        rowICTITEM1 = LookUp("ICTITEM1", CUST_ITEM_CODE)
                        ITEM_CODE = CUST_ITEM_CODE

                        If rowICTITEM1 Is Nothing Then
                            If InStr(CUST_ITEM_CODE, "'") <> 0 Then
                                rowICTITEM1 = LookUp("ICTITEM1", CUST_ITEM_CODE, True)
                            Else
                                If CUST_ITEM_CODE <> "" Then
                                    ASCMAIN1.sql = "Select * from ICTITEM1 where ITEM_CODE Like '" & CUST_ITEM_CODE & "%'"
                                    Dim tblICTITEM1 As DataTable = ASCDATA1.GetDataTable
                                    If tblICTITEM1.Rows.Count = 1 Then
                                        rowICTITEM1 = tblICTITEM1.Rows(0)
                                        ITEM_CODE = rowICTITEM1.Item("ITEM_CODE")
                                    Else
                                        rowICTITEM1 = tblICTITEM1.NewRow
                                        rowICTITEM1.Item(0) = CUST_ITEM_CODE
                                    End If
                                Else
                                    ' ?
                                End If
                            End If
                        End If
                        'If rowICTITEM1.Item("COLLECTION_CODE") & "" = "" Then Stop

                    End If
                Else
                    If rowICTITEM1 Is Nothing OrElse rowICTITEM1.Item("ITEM_CODE") & "" <> ITEM_CODE Then
                        rowICTITEM1 = LookUp("ICTITEM1", New String() {ITEM_CODE}, True)
                        'If rowICTITEM1.Item("COLLECTION_CODE") & "" = "" Then Stop
                    End If
                End If
                If rowICTITEM1 IsNot Nothing Then
                    row.Item("COLLECTION_CODE") = rowICTITEM1.Item("COLLECTION_CODE") & ""
                    row.Item("ITEM_DESC") = rowICTITEM1.Item("ITEM_DESC") & ""
                End If
                row.Item("ITEM_CODE") = ITEM_CODE

                'Dim DC() As DataColumn = dst.Tables("RSTRETLA").PrimaryKey
                Dim row2 As DataRow = tbl.Rows.Find(New Object() {CUST_CODE, CUST_STORE_NO, ITEM_CODE})
                If row2 IsNot Nothing Then
                    For W As Integer = 0 To 6
                        For Each C As String In New String() {"QTY", "AMT", "ONH"}
                            Dim COLUMN_NAME As String = C & "_" & IIf(W = 0, "TOTAL", "W" & Format(W, "0"))
                            row.Item(COLUMN_NAME) = Val(row.Item(COLUMN_NAME) & "") + Val(row2.Item(COLUMN_NAME) & "")
                        Next
                    Next

                    row2.Delete()
                End If
            End If

        Else
            row.Item("CUST_CODE") = Absx1.txtFor("CUST_CODE").Text
        End If

    End Sub

    Overrides Sub Prepare_for_View_Lookup_Special(ByVal ctl As Control, ByVal COLUMN_NAME As String, Optional ByRef sql_where As String = "", Optional ByRef Cancel As Boolean = False)
        Select Case COLUMN_NAME

            Case "CUST_CODE"
                sql_where = "" ' "VEND_TYPE = 'S'"

        End Select

    End Sub

#End Region

#Region "Popup Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdEDT852T1, "SSSBBBBB", "Show Filter", "Show GroupBox", "Show Pins", "Select All", "De-Select All", "Select All for Customer", "Archive Selected", "Unarchive Selected")

        Load_Popup_Menu(grdRSTRETLI, "SBB", "Show Bad Items Only", "Delete Bad Items", "Delete Items with all 0's")
        Load_Popup_Menu(grdRSTRETLA, "SBB", "Show Bad Items Only", "Delete Bad Items", "Delete Items with all 0's")
    End Sub

    Public Overrides Sub tlb_BeforeToolDropdown(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinToolbars.BeforeToolDropdownEventArgs)
        MyBase.tlb_BeforeToolDropdown(sender, e)

        If e.Tool.OwnerIsMenu Or e.SourceControl Is Nothing OrElse e.SourceControl.Name = "" Then
            e.Cancel = True
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.SourceControl.Name, 4))
        If grd Is Nothing Then
            e.Cancel = True
            Exit Sub
        End If

        Dim tlb_pop As UltraWinToolbars.PopupMenuTool = DirectCast(e.Tool, UltraWinToolbars.PopupMenuTool)
        Dim tlb_sbt As UltraWinToolbars.StateButtonTool = Nothing
        Dim tlb_btn As UltraWinToolbars.ButtonTool = Nothing

        If tlb_pop.Tools.Exists("Show Bad Items Only") Then
            tlb_sbt = DirectCast(tlb_pop.Tools("Show Bad Items Only"), UltraWinToolbars.StateButtonTool)
            tlb_sbt.Checked = (grd.Tag <> "")
        End If

        Select Case grd.Name

            Case "grdEDT852T1"

                tlb_btn = DirectCast(tlb_pop.Tools("Select All for Customer"), UltraWinToolbars.ButtonTool)
                If grd.ActiveRow Is Nothing OrElse Not grd.ActiveRow.IsDataRow Then
                    tlb_btn.SharedProps.Visible = False
                Else
                    Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                    tlb_btn.SharedProps.Visible = True
                    tlb_btn.SharedProps.Caption = "Select All for " & CUST_CODE
                End If

                If tlb_pop.Tools.Exists("Archive Selected") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Archive Selected"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (opt852Data.Value & "" = "0") ' Not Loaded Yet
                End If

                If tlb_pop.Tools.Exists("Unarchive Selected") Then
                    tlb_btn = DirectCast(tlb_pop.Tools("Unarchive Selected"), UltraWinToolbars.ButtonTool)
                    tlb_btn.SharedProps.Visible = (opt852Data.Value & "" = "A") ' Archived view
                End If

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            'e.Cancel = True
        Else
            Select Case e.SourceControl.Name

                'Case "grdARTSTMT1"
                '    If grd.ActiveRow.Cells("OPS_YYYYPP").Text = "999999" Then
                '        e.Cancel = True
                '    End If

            End Select
        End If
    End Sub

    Public Overrides Sub tlb_ToolClick(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinToolbars.ToolClickEventArgs)
        MyBase.tlb_ToolClick(sender, e)

        If e.Tool.OwningMenu Is Nothing OrElse Not GRDs.ContainsKey(Mid(e.Tool.OwningMenu.Key, 4)) Then
            Exit Sub
        End If

        Dim grd As UltraWinGrid.UltraGrid = GRDs(Mid(e.Tool.OwningMenu.Key, 4))

        Select Case e.Tool.Key
            Case "Show Bad Items Only"
                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                If tlb_sbt.Checked Then
                    If grd.Name = "grdRSTRETLA" Then
                        grd.Tag = "ISNULL(COLLECTION_CODE,'') = ''"
                    Else
                        grd.Tag = "ISNULL(FLAG,'') <> ''"
                    End If
                Else
                    grd.Tag = ""
                End If
                Dim dvw As DataView = DirectCast(grd.DataSource, DataTable).DefaultView
                dvw.RowFilter = grd.Tag

            Case "Select All", "De-Select All"
                For Each grow As UltraWinGrid.UltraGridRow In grdEDT852T1.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        grow.Cells("SELECTED").Value = IIf(e.Tool.Key = "Select All", "1", "0")
                        grow.Update()
                    End If
                Next
        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key

            Case "Delete Bad Items"
                Dim R As New List(Of DataRow)
                If grd.Name = "grdRSTRETLA" Then
                    For Each rowRSTRETLA As DataRow In dst.Tables("RSTRETLA") _
                    .Select("COLLECTION_CODE IS NULL", "", DataViewRowState.CurrentRows)
                        R.Add(rowRSTRETLA)
                    Next

                Else
                    For Each rowRSTRETLI As DataRow In dst.Tables("RSTRETLI") _
                    .Select("ISNULL(FLAG,'') <> ''", "", DataViewRowState.CurrentRows)
                        R.Add(rowRSTRETLI)
                    Next
                End If
                For Each row As DataRow In R
                    row.Delete()
                Next
                MsgBox(CStr(R.Count) & " Items Deleted", MsgBoxStyle.OkOnly, "Verification")

            Case "Delete Items with all 0's"
                Dim R As New List(Of DataRow)
                If grd.Name = "grdRSTRETLA" Then
                    For Each rowRSTRETLA As DataRow In dst.Tables("RSTRETLA") _
                    .Select("ISNULL(QTY_TOTAL,0) = 0 AND ISNULL(AMT_TOTAL,0) = 0 AND ISNULL(ONH_TOTAL,0) = 0", "", DataViewRowState.CurrentRows)
                        R.Add(rowRSTRETLA)
                    Next
                Else
                    For Each rowRSTRETLI As DataRow In dst.Tables("RSTRETLI") _
                    .Select("ISNULL(QTY_TOTAL,0) = 0 AND ISNULL(AMT_TOTAL,0) = 0 AND ISNULL(ONH_TOTAL,0) = 0", "", DataViewRowState.CurrentRows)
                        R.Add(rowRSTRETLI)
                    Next
                End If
                For Each row As DataRow In R
                    row.Delete()
                Next
                MsgBox(CStr(R.Count) & " Items Deleted", MsgBoxStyle.OkOnly, "Verification")

            Case "Select All for Customer"
                Dim CUST_CODE As String = grd.ActiveRow.Cells("CUST_CODE").Value & ""
                For Each grow As UltraWinGrid.UltraGridRow In grdEDT852T1.Rows
                    If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                        If grow.Cells("CUST_CODE").Value & "" = CUST_CODE Then
                            grow.Cells("SELECTED").Value = "1"
                            grow.Update()
                        End If
                    End If
                Next

            Case "Archive Selected"
                Set_Selected_852_Status("A")   ' 0 -> A

            Case "Unarchive Selected"
                Set_Selected_852_Status("0")   ' A -> 0


            Case Else
                'Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(e.Tool, UltraWinToolbars.StateButtonTool)
                'grd.DisplayLayout.GroupByBox.Hidden = Not tlb_sbt.Checked
        End Select
    End Sub

#End Region

#Region "ABSColumn Controls"

    Overrides Sub dte_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        MyBase.dte_KeyDown(sender, e)

        Dim COLUMN_NAME As String = Absx1.GetABSColumnName(sender)
        Select Case COLUMN_NAME
            Case "DTE0", "DTE1"
                If e.KeyCode = System.Windows.Forms.Keys.Enter And EntryMode = "" Then
                    Call Click_Command("Load", e)
                End If
        End Select

    End Sub

#End Region

    Sub Print_Report()
        Call Print_Report_Begin()
        Dim SUBT As String = ""

        Dim RecordSelectionFormula As String = ""

        Generate_Report("EDRSTATI", "Inbound EDI Transactions", SUBT, RecordSelectionFormula)

        Call Print_Report_End()
    End Sub

    Private Sub View_doc(ByVal DOCUMENTBLOBKEY As String)
        Dim FILENAME As String = "\\192.168.130.206\E$\GENSRVNT\DOCUMENTS\" & DOCUMENTBLOBKEY & ".DOC"

        If My.Computer.FileSystem.FileExists(FILENAME) Then
            Dim TEMP_FILENAME As String = ASCMAIN1.Folders("Temp") & DOCUMENTBLOBKEY & ".DOC"
            If My.Computer.FileSystem.FileExists(TEMP_FILENAME) Then
                My.Computer.FileSystem.DeleteFile(TEMP_FILENAME)
            End If
            My.Computer.FileSystem.CopyFile(FILENAME, TEMP_FILENAME)
            Dim p As Process = Process.Start("NOTEPAD.EXE", TEMP_FILENAME)
        End If
    End Sub

    Private Sub UltraTabControl1_SelectedTabChanged(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinTabControl.SelectedTabChangedEventArgs) Handles tabMain.SelectedTabChanged
        Setup_tabMain()
    End Sub

    Sub Setup_tabMain()
        If SELECTION_NO = 0 Then Exit Sub

        If tabMain.SelectedTab Is Nothing Then Exit Sub

        With UltraExplorerBar1.Groups("Batch Control")
            .Items("Import Raw EDI Files").Visible = (tabMain.SelectedTab.Key = "Raw EDI Files")
            .Items("Load 852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "0")
            .Items("Retract 852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "1")
            .Items("Restore Deleted").Visible = (tabMain.SelectedTab.Key = "852 Data" And opt852Data.Value = "D")
        End With

        With UltraExplorerBar1

            .Groups("Screen Control").Visible = (tabMain.SelectedTab.Key = "Manual Entry")

            Select Case tabMain.SelectedTab.Key
                Case "Raw EDI Files"
                    grdEDTJRNL1.Dock = DockStyle.None
                    grdEDTJRNL1.Parent = tabRaw.Tabs("Control Records").TabPage
                    grdEDTJRNL1.Dock = DockStyle.Fill

                    grpRawEDI.Dock = DockStyle.None
                    grpRawEDI.Parent = tabRaw.Tabs("Raw EDI").TabPage
                    grpRawEDI.Dock = DockStyle.Fill

                    grdEDT852T1.Dock = DockStyle.None
                    grdEDT852T1.Parent = tabRaw.Tabs("Documents").TabPage
                    grdEDT852T1.Dock = DockStyle.Fill
                Case "852 Data"
                    grdEDTJRNL1.Dock = DockStyle.None
                    grdEDTJRNL1.Parent = tab852.Tabs("Control Records").TabPage
                    grdEDTJRNL1.Dock = DockStyle.Fill

                    grpRawEDI.Dock = DockStyle.None
                    grpRawEDI.Parent = tab852.Tabs("Raw EDI").TabPage
                    grpRawEDI.Dock = DockStyle.Fill

                    grdEDT852T1.Dock = DockStyle.None
                    grdEDT852T1.Parent = SplitContainer3.Panel1
                    grdEDT852T1.Dock = DockStyle.Fill

                    Setup_EDT852T1()

                Case "Manual Entry"

                Case "852"
                    Setup_grdRSTSLPI0()

                Case "XLS"

            End Select

            .Groups("Raw EDI Files").Visible = False ' (tabMain.SelectedTab.Key = "Raw EDI Files")
            .Groups("852 Data").Visible = (tabMain.SelectedTab.Key = "852 Data")
            .Groups("Batch Control").Visible = (tabMain.SelectedTab.Key <> "Manual Entry")

        End With
    End Sub

    Sub Import_Raw_EDI()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Processing EDI Raw Files")

        Dim EDCIBND1 As New TAC.EDCIBND1()

        dst.Tables("EDTFILE1").Rows.Clear()

        Dim file_counter As Integer = 0

        Dim wildcard As String = "*.edi" ' "MAIL_IN.TXT"

        Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""
        For Each FILE As String In My.Computer.FileSystem.GetFiles _
        (ED_PARM_RAW_INBOUND, FileIO.SearchOption.SearchAllSubDirectories, wildcard)
            Dim FILEINFO As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE)
            Dim rowEDTFILE1 As DataRow = dst.Tables("EDTFILE1").NewRow
            Dim FILENAME As String = Mid(FILEINFO.FullName, ED_PARM_RAW_INBOUND.Length + 2)

            Dim row As DataRow = LookUp("EDTFILE1", FILENAME)
            If row Is Nothing Then
                rowEDTFILE1.Item("EDI_FILENAME") = FILENAME
                rowEDTFILE1.Item("EDI_FILESIZE") = FILEINFO.Length
                rowEDTFILE1.Item("EDI_DATETIME") = FILEINFO.LastWriteTime
                dst.Tables("EDTFILE1").Rows.Add(rowEDTFILE1)

                file_counter += 1

                ASCMAIN1.Progress("Processing " & FILENAME)
                BeginTrans()
                Dim EDI_JRNL_NOs As List(Of String)
                EDI_JRNL_NOs = EDCIBND1.Process_File(
                ED_PARM_RAW_INBOUND, FILENAME,
                FILEINFO, "852",
                ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE") & "",
                "", False)

                If EDI_JRNL_NOs.Count > 0 Then
                    Dim rowEDTJRNL1 = LookUp("EDTJRNL1", EDI_JRNL_NOs(0))
                    rowEDTFILE1.Item("EDI_JRNL_NO") = rowEDTJRNL1.Item("EDI_JRNL_NO")
                    rowEDTFILE1.Item("EDI_SENDER_QUAL") = rowEDTJRNL1.Item("EDI_SENDER_QUAL")
                    rowEDTFILE1.Item("EDI_SENDER_ID") = rowEDTJRNL1.Item("EDI_SENDER_ID")
                    rowEDTFILE1.Item("EDI_ISA_CTL_NO") = rowEDTJRNL1.Item("EDI_ISA_CTL_NO")
                    rowEDTFILE1.Item("EDI_ISA_CTL_DATE") = rowEDTJRNL1.Item("EDI_ISA_CTL_DATE")
                    ASCMAIN1.sql = "Select Count (*) from EDTJRNL3 where EDI_JRNL_NO = '" & EDI_JRNL_NOs(0) & "'"
                    Dim DOC_EDI As Integer = ASCDATA1.GetDataValue
                    ASCMAIN1.sql = "Select Count (*) from EDTJRNL3 where EDI_JRNL_NO = '" & EDI_JRNL_NOs(0) & "' and EDI_DOC_NO = '852'"
                    Dim DOC_852 As Integer = ASCDATA1.GetDataValue
                    rowEDTFILE1.Item("DOC_EDI") = DOC_EDI
                    rowEDTFILE1.Item("DOC_852") = DOC_852
                    rowEDTFILE1.Item("NOTES") = ""
                End If
                Update_Record_TDA("EDTFILE1")
                CommitTrans()
            End If
        Next

        'Update_Record_TDA("EDTFILE1")

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        EDCIBND1 = Nothing

        Prepare_852_Queue()

        MsgBox(CStr(file_counter) & " Files have been Imported", MsgBoxStyle.OkOnly, "Verification")

    End Sub

    Sub Prepare_852_Queue()


        ASCDATA1.ExecuteSQL("Update EDT852T1 Set EDI_STATUS = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_STATUS is Null")

        ASCMAIN1.sql = "Update EDT852T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
               & " where EDI_TP_QUAL = EDT852T1.EDI_TP_QUAL and EDI_TP_ID = EDT852T1.EDI_TP_ID)" _
               & " where EDI_STATUS = '0' and COMPANY_CODE IS NULL"
        ASCDATA1.ExecuteSQL()

        'LORD SENDS ON HAND REPORTS IN THE FOLLOWING WEEK, BUT AS BOW NOT EOW

        ASCMAIN1.sql = "Update EDT852T1 Set OPS_YYYYWW = " & vbCrLf _
        & " (Select Min (YYYYWW) from GLTPARM3 " & vbCrLf _
        & " where WEEK_END_DATE >= EDT852T1.EDI_FROM_DATE -7)" & vbCrLf _
        & " where EDI_STATUS = '0' and OPS_YYYYWW is Null and CUST_CODE = 'LORD'" & vbCrLf _
        & "   and EDI_DOC_SEQ_NO in (" & vbCrLf _
        & "SELECT EDI_DOC_SEQ_NO FROM (" & vbCrLf _
        & "SELECT EDI_DOC_SEQ_NO" & vbCrLf _
        & ", SUM (DECODE(EDI_ZA_TRAN_TYPE,'QA',1,'QX',1,0)) QAX" & vbCrLf _
        & ", SUM (DECODE(EDI_ZA_TRAN_TYPE,'QA',0,'QX',0,1)) Q_OTHER" & vbCrLf _
        & " FROM EDT852T3 WHERE EDI_DOC_SEQ_NO IN (" & vbCrLf _
        & "SELECT EDI_DOC_SEQ_NO FROM EDT852T1 " & vbCrLf _
        & "WHERE EDI_STATUS = '0' and OPS_YYYYWW is Null and CUST_CODE = 'LORD')" & vbCrLf _
        & "GROUP BY EDI_DOC_SEQ_NO" & vbCrLf _
        & ") WHERE QAX <> 0 AND Q_OTHER = 0)"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT852T1 Set OPS_YYYYWW = " _
        & " (Select Min (YYYYWW) from GLTPARM3 " _
        & " where WEEK_END_DATE >= EDT852T1.EDI_FROM_DATE)" _
        & " where EDI_STATUS = '0' and OPS_YYYYWW is Null"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT852T1 Set OPS_YYYYPP = " _
        & " (Select YYYYPP from GLTPARM3 " _
        & " where YYYYWW = EDT852T1.OPS_YYYYWW)" _
        & " where EDI_STATUS = '0' and OPS_YYYYPP is Null"
        ASCDATA1.ExecuteSQL()

        If opt852Data.CheckedIndex <> 0 Then
            opt852Data.CheckedIndex = 0
        Else
            Setup_EDT852T1()
        End If

    End Sub

    Private Sub grdEDT852T1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDT852T1.AfterRowActivate
        Setup_EDT852T1_Details()
    End Sub

    Private Sub grdEDT852T1_AfterRowsDeleted(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDT852T1.AfterRowsDeleted
        Dim sql As String = ""
        For Each EDI_DOC_SEQ_NO As String In EDI_DOC_SEQ_NOs
            sql &= ",'" & EDI_DOC_SEQ_NO & "'"
        Next
        ASCMAIN1.sql = "Update EDT852T1 Set EDI_STATUS = 'D', LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "' where EDI_DOC_SEQ_NO in (" & Mid(sql, 2) & ")"
        ASCDATA1.ExecuteSQL()
        dst.Tables("EDT852T1").AcceptChanges()
    End Sub

    Private Sub grdEDT852T1_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdEDT852T1.BeforeRowsDeleted
        EDI_DOC_SEQ_NOs.Clear()
        For Each grow As UltraWinGrid.UltraGridRow In grdEDT852T1.Selected.Rows
            EDI_DOC_SEQ_NOs.Add(grow.Cells("EDI_DOC_SEQ_NO").Text)
        Next
    End Sub

    Private Sub grdEDT852T1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDT852T1.InitializeLayout

    End Sub

    Sub Setup_EDT852T1()

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Retrieving Previously Loaded 852 Data")

        grdEDT852T1.Text = "852 Documents (" & opt852Data.Text & ")"

        cmd852SelectAll.Visible = Not (opt852Data.Value = "M")
        cmd852DeSelectAll.Visible = Not (opt852Data.Value = "M")

        With UltraExplorerBar1.Groups("Batch Control")
            .Items("Load 852 Data").Visible = (opt852Data.Value = "0")
            .Items("Retract 852 Data").Visible = (opt852Data.Value = "1")
            .Items("Restore Deleted").Visible = (opt852Data.Value = "D")
        End With

        Dim EDI_STATUS As String = opt852Data.Value
        If opt852Data.Value = "1" Then
            grdEDT852T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
            Load_Customer_Summary()
        ElseIf opt852Data.Value = "0" Then
            grdEDT852T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.True
            Load_Customer_Summary()
        ElseIf opt852Data.Value = "D" Then
            grdEDT852T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        ElseIf opt852Data.Value = "M" Then
            grdEDT852T1.DisplayLayout.Override.AllowDelete = DefaultableBoolean.False
        End If

        tab852.Tabs("Items").Visible = (EDI_STATUS = "1" Or EDI_STATUS = "A")
        tab852.Tabs("Stores").Visible = (EDI_STATUS = "1" Or EDI_STATUS = "A")
        tab852.Tabs("Raw EDI").Visible = (EDI_STATUS = "1" Or EDI_STATUS = "A")
        tab852.Tabs("Imported").Visible = (EDI_STATUS = "1" Or EDI_STATUS = "A")
        tab852.Tabs("Control Records").Visible = (EDI_STATUS = "1" Or EDI_STATUS = "A")

        Fill_Records("EDT852T1", EDI_STATUS)
        Sort_grdColumns(grdEDT852T1, "EDI_DOC_SEQ_NO".ToLower)
        If grdEDT852T1.Rows.Count > 0 Then ' If grdEDT852T1.Rows.Count = 0 Then
            Setup_EDT852T1_Details()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Sub Setup_EDT852T1_Details()

        If grdEDT852T1.ActiveRow Is Nothing OrElse Not grdEDT852T1.ActiveRow.IsDataRow Then
            SplitContainer3.Panel2Collapsed = True
            Exit Sub
        End If

        SplitContainer3.Panel2Collapsed = False
        Dim EDI_DOC_SEQ_NO As String = grdEDT852T1.ActiveRow.Cells("EDI_DOC_SEQ_NO").Text
        dst.EnforceConstraints = False
        Fill_Records("EDT852T2", EDI_DOC_SEQ_NO)
        Fill_Records("EDT852T3", EDI_DOC_SEQ_NO)

        'Fill_Records("EDTJRNL3", EDI_DOC_SEQ_NO)

        'If opt852Data.Value = "M" Then
        '    dst.Tables("EDTJRNL1").Rows.Clear()
        '    dst.Tables("EDTJRNL2").Rows.Clear()
        'Else
        '    Dim rowEDTJRNL3 As DataRow = dst.Tables("EDTJRNL3").Rows(0)
        '    Dim EDI_JRNL_NO As String = rowEDTJRNL3.Item("EDI_JRNL_NO")
        '    Dim EDI_GS_NO As Integer = rowEDTJRNL3.Item("EDI_GS_NO")
        '    Fill_Records("EDTJRNL1", EDI_JRNL_NO)
        '    Fill_Records("EDTJRNL2", New String() {EDI_JRNL_NO, CStr(EDI_GS_NO)})
        'End If
        'dst.EnforceConstraints = True

        'If opt852Data.Value = "M" Then
        '    txtRawEDI.Text = ""
        'Else
        '    Dim rowEDTJRNL1 As DataRow = dst.Tables("EDTJRNL1").Rows(0)
        '    Dim EDI_FOLDERNAME As String = rowEDTJRNL1.Item("EDI_FOLDERNAME")
        '    Dim EDI_FILENAME As String = rowEDTJRNL1.Item("EDI_FILENAME")
        '    Dim FILENAME As String = EDI_FOLDERNAME & "\" & EDI_FILENAME

        '    Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        '    grpRawEDI.Text = FI.Name & " " & Format$(FI.LastWriteTime, "MM/dd/yy HH:mm")
        '    txtRawEDI.Text = ""
        '    If My.Computer.FileSystem.FileExists(FILENAME) Then
        '        Try
        '            Using SR As New System.IO.StreamReader(FILENAME)
        '                Dim RAW As String = SR.ReadToEnd
        '                txtRawEDI.Text = Replace(RAW, Mid(RAW, 106, 1), vbCrLf)
        '            End Using
        '        Catch ex As Exception

        '        End Try
        '    End If
        'End If

        txtRawEDI.Text = ""
        txtRawEDI.Text = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO, ROWs("EDTPARM1").Item("ED_PARM_RAW_ARCHIVE"), "852")

        If ASCMAIN1.CLIENT = "SLP" Then
            Dim GEN_DOC_NO As String = grdEDT852T1.ActiveRow.Cells("GEN_DOC_NO").Text
            Dim RS_PARM_FOLDER_SELLTHRU_DATA As String = ROWs("RSTPARM1").Item("RS_PARM_FOLDER_SELLTHRU_DATA")
            Dim FILENAME As String = System.IO.Path.Combine(RS_PARM_FOLDER_SELLTHRU_DATA, "\Archive\") & GEN_DOC_NO
            If My.Computer.FileSystem.FileExists(FILENAME) Then
                txtRawEDI.Text = My.Computer.FileSystem.ReadAllText(FILENAME)
            End If
        End If

        If opt852Data.Value = "1" Or opt852Data.Value = "A" Then
            Fill_Records("EDT852TI", EDI_DOC_SEQ_NO)
            Dim CUST_CODE As String = grdEDT852T1.ActiveRow.Cells("CUST_CODE").Value & ""
            If CUST_CODE = "AMAZON" Then
                For Each ROW As DataRow In dst.Tables("EDT852TI").Select("ITEM_CODE IS NULL")
                    Dim EDI_ITEM_CODE As String = ROW.Item("EDI_ITEM_CODE") & ""
                    Dim rowRSTASIN1s() As DataRow = dst.Tables("RSTASIN1").Select($"CUST_CODE = '{CUST_CODE}' AND EAN_UPC = '{EDI_ITEM_CODE}'")
                    If rowRSTASIN1s.Length > 0 Then
                        ROW.Item("ASIN") = rowRSTASIN1s(0).Item("ASIN")
                    End If
                Next
            Else

            End If
            Sort_grdColumns(grdEDT852TI, "EDI_ITEM_CODE")
            Fill_Records("EDT852TS", EDI_DOC_SEQ_NO)
            Sort_grdColumns(grdEDT852TS, "EDI_STORE_NO")
        End If

    End Sub

    Sub FIX_EDTJRNL1()

        Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""

        dst.EnforceConstraints = False
        Dim Sql As String = "Select * from EDTJRNL1"
        Fill_Records("EDTJRNL1", "", True, Sql)
        For Each row As DataRow In dst.Tables("EDTJRNL1").Rows
            Dim FILENAME As String = ED_PARM_RAW_INBOUND & "\" & row.Item("EDI_FILENAME")
            row.Item("EDI_FILESIZE") = My.Computer.FileSystem.GetFileInfo(FILENAME).Length
            row.Item("EDI_DATETIME") = My.Computer.FileSystem.GetFileInfo(FILENAME).LastWriteTime
        Next
        Call Update_Record_TDA("EDTJRNL1")
        Stop

    End Sub

    Private Sub opt852Data_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles opt852Data.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_EDT852T1()
    End Sub

    Private Sub cmd852SelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd852SelectAll.Click
        'For Each rowEDT852T1 As DataRow In dst.Tables("EDT852T1").Rows
        '    rowEDT852T1.Item("SELECTED") = "1"
        'Next
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Selecting All")
        For Each grow As UltraWinGrid.UltraGridRow In grdEDT852T1.Rows
            If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                grow.Cells("SELECTED").Value = "1"
                grow.Update()
            End If
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Private Sub cmd852DeSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd852DeSelectAll.Click
        'For Each rowEDT852T1 As DataRow In dst.Tables("EDT852T1").Rows
        '    rowEDT852T1.Item("SELECTED") = "0"
        'Next
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Selecting All")
        For Each grow As UltraWinGrid.UltraGridRow In grdEDT852T1.Rows
            If grow.IsDataRow AndAlso Not grow.IsFilteredOut Then
                grow.Cells("SELECTED").Value = "0"
                grow.Update()
            End If
        Next
        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_852_Data()

        dst.Tables("EDT852T1").AcceptChanges()
        dst.Tables("EDT852T0_BADITEM").Rows.Clear()

        BeginTrans()

        For Each rowEDT852T1 As DataRow In dst.Tables("EDT852T1").Select("SELECTED = '1'", "EDI_DOC_SEQ_NO")
            Dim EDI_DOC_SEQ_NO As String = rowEDT852T1.Item("EDI_DOC_SEQ_NO")
            Dim OPS_YYYYPP As String = rowEDT852T1.Item("OPS_YYYYPP")
            Dim CUST_CODE As String = rowEDT852T1.Item("CUST_CODE")
            ASCMAIN1.Progress("Now Processing Document " & EDI_DOC_SEQ_NO, "")
            Application.DoEvents()

            If opt852Data.Value <> "D" Then
                If opt852Data.Value = "0" Then

                    ASCDATA1.ExecuteSQL("Delete from EDT852T0 where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'")
                    Dim SQL As String = ""
                    For i As Integer = 1 To 10
                        Dim z As String = Format$(i, "00")
                        SQL = "Insert INTO EDT852T0" & vbCrLf _
                        & " (EDI_DOC_SEQ_NO, EDI_LINE_NO, EDI_ITEM_CODE, EDI_TRAN_TYPE " & vbCrLf _
                        & ", EDI_STORE_NO, EDI_QTY, CUST_CODE, CUST_STORE_NO)" & vbCrLf _
                        & " Select EDT852T2.EDI_DOC_SEQ_NO, EDT852T2.EDI_LINE_NO" & vbCrLf _
                        & ", NVL(NVL(NVL(EDT852T2.EDI_ITEM_UP, EDT852T2.EDI_ITEM_EN),SUBSTR(EDT852T2.EDI_ITEM_GTIN,3)),EDT852T2.EDI_BUYER_ITEM) EDI_ITEM_CODE" & vbCrLf _
                        & ", EDT852T3.EDI_ZA_TRAN_TYPE EDI_TRAN_TYPE" & vbCrLf _
                        & ", EDT852T3.EDI_SDQ_STORE_" & z & " EDI_STORE_NO" & vbCrLf _
                        & ", EDT852T3.EDI_SDQ_QTY_AMT_" & z & " EDI_QTY" & vbCrLf _
                        & ", NVL(EDTTRPM3.CUST_CODE,EDT852T1.CUST_CODE) CUST_CODE" & vbCrLf _
                        & ", CASE WHEN EDT852TG.CUST_STORE_NO IS NOT NULL AND EDT852TG.CUST_CODE = NVL(EDTTRPM3.CUST_CODE,EDT852T1.CUST_CODE) THEN EDT852TG.CUST_STORE_NO ELSE DECODE(IS_NUMBER(EDT852T3.EDI_SDQ_STORE_" & z & "),1,LPAD(EDT852T3.EDI_SDQ_STORE_" & z & ",6,'0'),EDT852T3.EDI_SDQ_STORE_" & z & ") END CUST_STORE_NO" & vbCrLf _
                        & " from (SELECT EDT852T1.EDI_TP_QUAL, EDT852T1.EDI_TP_ID, EDT852T3.* FROM EDT852T3, EDT852T1 where EDT852T1.EDI_DOC_SEQ_NO = EDT852T3.EDI_DOC_SEQ_NO and EDT852T1.EDI_STATUS = '0' and NVL(EDT852T3.UNIT_BASIS_CODE,'EA') = 'EA') EDT852T3" & vbCrLf _
                        & ", EDT852T2, EDT852T1, EDTTRPM3, " & EDT852TG & " EDT852TG" & vbCrLf _
                        & " where EDT852T2.EDI_DOC_SEQ_NO = EDT852T1.EDI_DOC_SEQ_NO" & vbCrLf _
                        & "   and EDT852T3.EDI_DOC_SEQ_NO = EDT852T2.EDI_DOC_SEQ_NO" & vbCrLf _
                        & "   and EDT852T3.EDI_LINE_NO = EDT852T2.EDI_LINE_NO" & vbCrLf _
                        & "   and EDT852T3.EDI_ZA_TRAN_TYPE IN ('QS','QU','QA')" & vbCrLf _
                        & "   and EDT852T1.EDI_STATUS = '0'" & vbCrLf _
                        & "   and EDT852T1.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                        & "   and EDTTRPM3.EDI_TP_QUAL (+) = TRIM(EDT852T3.EDI_TP_QUAL)" & vbCrLf _
                        & "   and EDTTRPM3.EDI_TP_ID (+) = TRIM(EDT852T3.EDI_TP_ID)" & vbCrLf _
                        & "   and EDTTRPM3.EDI_DOC_NO (+) = '852'" & vbCrLf _
                        & "   and EDTTRPM3.EDI_STORE (+) = EDT852T3.EDI_SDQ_STORE_" & z & "" & vbCrLf _
                        & "   and EDT852T3.EDI_SDQ_STORE_" & z & " IS NOT NULL" & vbCrLf _
                        & "   and NVL(EDT852T3.EDI_SDQ_QTY_AMT_" & z & ",0) <> 0" & vbCrLf _
                        & "   and EDT852TG.GLOBAL_LOCATION_NUMBER (+) = TRIM(EDT852T3.EDI_SDQ_STORE_" & z & ")"
                        ASCDATA1.ExecuteSQL(SQL)
                        'SQL = SQL & " AND EDT852T4.EDI_SDQ_TYPE ='EA'"
                    Next i

                    'If CUST_CODE = "CARLYLE10" Then
                    '    ASCMAIN1.sql = "Update EDT852T0 set EDI_QTY = -1 * EDI_QTY where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "' and EDI_TRAN_TYPE = 'QU'"
                    '    ASCDATA1.ExecuteSQL()
                    'End If

                    Set_ITEM_CODE(EDI_DOC_SEQ_NO, CUST_CODE)

                    ASCMAIN1.sql = "INSERT INTO EDTUPCX1 (CUST_CODE, EDI_ITEM_CODE)" & vbCrLf _
                    & " SELECT DISTINCT CUST_CODE, NVL(EDI_ITEM_CODE,'NONE') EDI_ITEM_CODE " & vbCrLf _
                    & " FROM EDT852T0 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" & vbCrLf _
                    & " and ITEM_CODE IS NULL" & vbCrLf _
                    & " minus " & vbCrLf _
                    & " SELECT CUST_CODE, EDI_ITEM_CODE from " & vbCrLf _
                    & "  EDTUPCX1 WHERE CUST_CODE IN (SELECT DISTINCT CUST_CODE FROM EDT852T0 WHERE EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "')"
                    ASCDATA1.ExecuteSQL()


                    ASCMAIN1.sql = "Update EDT852T1 Set ERROR_ITEMS = " _
                    & "(Select Count (DISTINCT EDI_ITEM_CODE) from EDT852T0 " _
                    & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
                    & " and ITEM_CODE is Null) " _
                    & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                    ASCDATA1.ExecuteSQL()


                    Fill_Records("EDT852T0_BADITEM", EDI_DOC_SEQ_NO, False)

                End If

                Update_RSTRETL1(EDI_DOC_SEQ_NO, OPS_YYYYPP)
            End If

            Dim EDI_STATUS As String = "1"
            Dim sqlINIT As String = ", INIT_DATE = SYSDATE, INIT_OPER = '" & ASCMAIN1.USER_ID & "'"
            If opt852Data.Value = "1" Or opt852Data.Value = "D" Then
                EDI_STATUS = "0"
                sqlINIT = ""
            Else

            End If

            ASCMAIN1.sql = "Update EDT852T1 set EDI_STATUS = '" & EDI_STATUS & "'" _
                & sqlINIT & ", LAST_DATE = SYSDATE, LAST_OPER = '" & ASCMAIN1.USER_ID & "'" _
                & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            ASCDATA1.ExecuteSQL()

            'rowEDT852T1.Item("EDI_STATUS") = "1"
        Next

        'Update_Record_TDA("EDT852T1")
        ASCMAIN1.Progress("", "")
        CommitTrans("Update Complete")

        If dst.Tables("EDT852T0_BADITEM").Rows.Count <> 0 Then
            Dim frm As New ASFMSGBF
            frm.Show_grd(dst.Tables("EDT852T0_BADITEM"), Me, "EDI Records which could not be mapped to Item Codes")
        End If

        Prepare_852_Queue()

    End Sub

    Sub Set_ITEM_CODE(
    ByVal EDI_DOC_SEQ_NO As String,
    ByVal CUST_CODE As String)

        ASCMAIN1.sql = "" _
        & " Begin DECLARE CURSOR C1 IS" _
        & " Select ITEM_UPC_CODE EDI_ITEM_CODE, ITEM_CODE " _
        & "  from ICTITEM1 where ITEM_UPC_CODE IN (" _
        & " Select Distinct EDI_ITEM_CODE FROM EDT852T0 " _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & "   and EDI_ITEM_CODE IS NOT NULL)" _
        & " union" _
        & " Select ITEM_EAN_CODE EDI_ITEM_CODE, ITEM_CODE " _
        & "  from ICTITEM1 where ITEM_EAN_CODE IN (" _
        & " Select Distinct EDI_ITEM_CODE FROM EDT852T0 " _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & "   and EDI_ITEM_CODE IS NOT NULL)" _
        & " union" _
        & " Select SUBSTR(ITEM_EAN_CODE,2) EDI_ITEM_CODE, ITEM_CODE " _
        & "  from ICTITEM1 where SUBSTR(ITEM_EAN_CODE,2) IN (" _
        & " Select Distinct EDI_ITEM_CODE FROM EDT852T0 " _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & "   and EDI_ITEM_CODE IS NOT NULL);" _
        & " Begin FOR R1 IN C1 LOOP" _
        & " Update EDT852T0 SET ITEM_CODE = R1.ITEM_CODE" _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & "   and EDI_ITEM_CODE = R1.EDI_ITEM_CODE;" _
        & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

        Dim SQLX As String = "SELECT EDTUPCX1.EDI_ITEM_CODE, ICTITEM1.ITEM_CODE" _
        & " from EDTUPCX1,ICTITEM1" _
        & " where EDTUPCX1.CUST_CODE = '" & CUST_CODE & "'" _
        & " and EDTUPCX1.EDI_ITEM_CODE IN (" _
        & " Select DISTINCT EDI_ITEM_CODE from EDT852T0 " _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & " and ITEM_CODE is Null)" _
        & " and NVL(EDTUPCX1.IGNORE,'0') <> '1'"
        SQLX = "" _
        & SQLX _
        & " and EDTUPCX1.ITEM_CODE is Null" _
        & " and ICTITEM1.ITEM_UPC_CODE = EDTUPCX1.ITEM_UPC_CODE" _
        & " union " _
        & SQLX _
        & " and EDTUPCX1.ITEM_CODE is Not Null" _
        & " and ICTITEM1.ITEM_CODE = EDTUPCX1.ITEM_CODE"

        ASCMAIN1.sql = "Begin Declare Cursor C1 is " & SQLX & ";" _
        & " Begin For R1 in C1 Loop" _
        & " Update EDT852T0 Set ITEM_CODE = R1.ITEM_CODE " _
        & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
        & " and ITEM_CODE is Null" _
        & " and EDI_ITEM_CODE = R1.EDI_ITEM_CODE;" _
        & " End Loop; End; End;"
        ASCDATA1.ExecuteSQL()

    End Sub

    Sub Update_RSTRETL1(ByVal EDI_DOC_SEQ_NO As String, ByVal OPS_YYYYPP As String)

        If opt852Data.Value = "1" Then
            Update_RSTRETLx(EDI_DOC_SEQ_NO)

            ASCMAIN1.sql = "Delete from RSTRETL1 " _
            & " where EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
            ASCDATA1.ExecuteSQL()
        Else
            ASCMAIN1.sql = "" _
            & " Insert INTO RSTRETL1" _
            & " Select EDT852T0.EDI_DOC_SEQ_NO" _
            & " , EDT852T0.CUST_CODE, EDT852T0.CUST_STORE_NO" _
            & " , EDT852T0.ITEM_CODE" _
            & " , SUM (DECODE(EDI_TRAN_TYPE,'QS',NVL(EDI_QTY,0),'QU',-1 * NVL(EDI_QTY,0),0)) QTY_SOLD" _
            & " , SUM (DECODE(EDI_TRAN_TYPE,'QS',NVL(EDI_QTY,0),'QU',-1 * NVL(EDI_QTY,0),0) * NVL(ICTRETLA.ITEM_RETAIL_PRICE,NVL(ICTITEM1.ITEM_RETAIL_PRICE,0))) AMT_SOLD" _
            & " , EDT852T1.OPS_YYYYPP, EDT852T1.OPS_YYYYWW" _
            & " , SUM (DECODE(EDI_TRAN_TYPE,'QA',EDI_QTY)) QTY_EOW" _
            & " from EDT852T0,EDT852T1,ICTITEM1,ICTRETLA" _
            & " where EDT852T0.EDI_DOC_SEQ_NO = EDT852T1.EDI_DOC_SEQ_NO" _
            & " and EDT852T0.ITEM_CODE is Not Null" _
            & " and ICTRETLA.ITEM_CODE (+) = EDT852T0.ITEM_CODE" _
            & " and ICTRETLA.OPS_YYYYPP (+) = '" & OPS_YYYYPP & "'" _
            & " and ICTITEM1.ITEM_CODE = EDT852T0.ITEM_CODE" _
            & " and EDT852T0.EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'" _
            & " group by EDT852T0.EDI_DOC_SEQ_NO" _
            & " , EDT852T0.CUST_CODE, EDT852T0.CUST_STORE_NO" _
            & " , EDT852T0.ITEM_CODE" _
            & " , EDT852T1.OPS_YYYYPP, EDT852T1.OPS_YYYYWW"
            ASCDATA1.ExecuteSQL()

            Update_RSTRETLx(EDI_DOC_SEQ_NO)
        End If
    End Sub

    Private Sub chkRawProcessed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRawPreviouslyImported.CheckedChanged
        Setup_tabRaw()
    End Sub

    Sub Setup_tabRaw()
        SplitContainer2.Panel2Collapsed = Not (chkRawPreviouslyImported.Checked)
        If Not chkRawPreviouslyImported.Checked Then
            dst.Tables("EDTFILE1").Rows.Clear()

            grdEDTFILE1.DisplayLayout.GroupByBox.Hidden = True
            Show_Filter(grdEDTFILE1, False)

        Else

            grdEDTFILE1.DisplayLayout.GroupByBox.Hidden = False
            Show_Filter(grdEDTFILE1, True)

            Me.Cursor = Cursors.WaitCursor
            ASCMAIN1.Progress("Now Retrieving Files Processed")

            Dim SQL As String = "Select EDTJRNL1.EDI_FILENAME" _
            & ", MIN (EDTJRNL1.EDI_FILESIZE) EDI_FILESIZE" _
            & ", MIN (EDTJRNL1.EDI_DATETIME) EDI_DATETIME" _
            & ", MIN (EDTJRNL1.EDI_JRNL_NO) EDI_JRNL_NO" _
            & ", MIN (EDTJRNL1.EDI_SENDER_QUAL) EDI_SENDER_QUAL" _
            & ", MIN (EDTJRNL1.EDI_SENDER_ID) EDI_SENDER_ID" _
            & ", MIN (EDTJRNL1.EDI_ISA_CTL_NO) EDI_ISA_CTL_NO" _
            & ", MIN (EDTJRNL1.EDI_ISA_CTL_DATE) EDI_ISA_CTL_DATE" _
            & ", COUNT (EDTJRNL3.EDI_DOC_NO) DOC_EDI" _
            & ", COUNT (EDTJRNL3.EDI_DOC_SEQ_NO) DOC_852, NULL NOTES from EDTJRNL1,EDTJRNL3" _
            & " where EDTJRNL1.EDI_JRNL_NO = EDTJRNL3.EDI_JRNL_NO (+)" _
            & " group by EDTJRNL1.EDI_FILENAME"
            Fill_Records("EDTFILE1", "", True, SQL)
            Sort_grdColumns(grdEDTFILE1, "EDI_FILENAME")

            If grdEDTFILE1.ActiveRow Is Nothing Then
                tabRaw.Visible = False
            Else
                tabRaw.Visible = True
                Setup_EDTFILE1_Details()
            End If

            Me.Cursor = Cursors.Default
            ASCMAIN1.Progress("")

        End If
    End Sub

    Private Sub grdEDTFILE1_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdEDTFILE1.AfterRowActivate
        If grdEDTFILE1.ActiveRow.IsDataRow Then
            Setup_EDTFILE1_Details()
        End If
    End Sub

    Private Sub grdEDTFILE1_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles grdEDTFILE1.InitializeLayout

    End Sub

    Sub Setup_EDTFILE1_Details()

        Dim ED_PARM_RAW_INBOUND As String = ROWs("EDTPARM1").Item("ED_PARM_RAW_INBOUND") & ""

        Dim FILENAME As String = ED_PARM_RAW_INBOUND & "\" & grdEDTFILE1.ActiveRow.Cells("EDI_FILENAME").Text

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Loading Data for file")

        Dim FI As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
        grpRawEDI.Text = FI.Name & " " & Format$(FI.LastWriteTime, "MM/dd/yy HH:mm")
        txtRawEDI.Text = ""
        ' txtRawEDI.Text = TAC.SOCMAIN1.Get_Raw_EDI(EDI_DOC_SEQ_NO)
        Try
            Using SR As New System.IO.StreamReader(FILENAME)
                Dim RAW As String = SR.ReadToEnd
                txtRawEDI.Text = Replace(RAW, Mid(RAW, 106, 1), vbCrLf)
            End Using

        Catch ex As Exception

        End Try

        Dim EDI_JRNL_NO As String = grdEDTFILE1.ActiveRow.Cells("EDI_JRNL_NO").Text
        Dim sql As String = ""


        dst.EnforceConstraints = False

        'Fill_Records("EDT852T2", EDI_DOC_SEQ_NO)
        'Fill_Records("EDT852T3", EDI_DOC_SEQ_NO)

        sql = "Select EDT852T1.*, ARTCUST1.CUST_NAME, '0' SELECTED " _
        & " from EDT852T1,ARTCUST1 " _
        & " where EDT852T1.EDI_DOC_SEQ_NO in (Select EDI_DOC_SEQ_NO FROM EDTJRNL1,EDTJRNL3 where EDTJRNL1.EDI_JRNL_NO = EDTJRNL3.EDI_JRNL_NO and EDTJRNL1.EDI_JRNL_NO = '" & EDI_JRNL_NO & "')" _
        & " and ARTCUST1.CUST_CODE (+) = EDT852T1.CUST_CODE"
        Fill_Records("EDT852T1", "", True, sql)

        Fill_Records("EDTJRNL1", EDI_JRNL_NO)

        sql = "Select EDTJRNL2.* from EDTJRNL2 " _
        & " where EDTJRNL2.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
        Fill_Records("EDTJRNL2", "", True, sql)

        sql = "Select EDTJRNL3.* from EDTJRNL3 " _
        & " where EDTJRNL3.EDI_JRNL_NO = '" & EDI_JRNL_NO & "'"
        Fill_Records("EDTJRNL3", "", True, sql)

        dst.EnforceConstraints = True

        Sort_grdColumns(grdEDT852T1, "EDI_DOC_SEQ_NO")

        If grdEDTJRNL1.Rows.Count > 0 Then
            grdEDTJRNL1.Rows(0).ExpandAll()
        End If

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")
    End Sub

    Sub Load_Customer_Summary()

        Dim sql As String = ""
        Dim W As Integer = 0
        Dim RYP As String = ASCMAIN1.CYP
        With grdEDT852TC.DisplayLayout.Bands(0)
            If .Groups.Count <> 0 Then
                For g As Integer = .Groups.Count - 1 To 0 Step -1
                    .Groups.Remove(g)
                Next
            End If
            .Groups.Add("CUST_CODE")
            .Groups("CUST_CODE").Header.Caption = ""
            .Columns("CUST_CODE").Group = .Groups("CUST_CODE")
            For P As Integer = 36 To 0 Step -1
                Dim YP As String = ASCMAIN1.Period_Calc(RYP, -1 * P)
                .Groups.Add(YP)
                Dim LEGEND As String = ASCMAIN1.Get_Legend(YP)
                .Groups(YP).Header.Caption = Mid(LEGEND, 10, 6)
                .Groups(YP).Header.Appearance.BackColor = System.Drawing.Color.Yellow
                .Groups(YP).Header.Appearance.BackGradientStyle = GradientStyle.ForwardDiagonal

                ASCMAIN1.sql = "Select YYYYWW from GLTPARM3 where YYYYPP = '" & YP & "'"
                For Each row As DataRow In ASCDATA1.GetDataTable.Select("", "YYYYWW")
                    W += 1
                    Dim COLUMN_NAME As String = "F" & Format(W, "000")
                    .Columns(COLUMN_NAME).Group = .Groups(YP)
                    Dim YW As String = row.Item("YYYYWW")
                    .Columns(COLUMN_NAME).Tag = YW
                    sql &= ", Sum (Decode(OPS_YYYYWW,'" & YW & "',1,0)) " & COLUMN_NAME

                    If P Mod 2 = 0 Then
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = System.Drawing.Color.LightBlue
                    Else
                        .Columns(COLUMN_NAME).CellAppearance.BackColor = System.Drawing.Color.LightPink
                    End If
                    .Columns(COLUMN_NAME).Hidden = False
                    .Columns(COLUMN_NAME).Header.Caption = Mid(YW, 5, 2)
                    .Columns(COLUMN_NAME).Width = 30
                Next
            Next
            If W < 27 Then
                For I As Integer = W + 1 To 27
                    Dim COLUMN_NAME As String = "F" & Format(I, "00")
                    .Columns(COLUMN_NAME).Hidden = True
                Next
            End If
        End With

        With grdEDT852TC.DisplayLayout.Bands("EDT852TC")
            .Groups(0).Header.Fixed = True
        End With

        sql = "Select CUST_CODE" & sql _
        & " from EDT852T1 " _
        & " where OPS_YYYYPP >= '" & ASCMAIN1.Period_Calc(RYP, -36) & "'" _
        & " and OPS_YYYYPP <= '" & RYP & "'" _
        & " and EDI_STATUS IN ('1','M')" _
        & " group by CUST_CODE"
        Fill_Records("EDT852TC", "", True, sql)
        Sort_grdColumns(grdEDT852TC, "CUST_CODE")
    End Sub

    Private Sub grdEDT852TI_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT852TI.InitializeRow
        If e.Row.Cells("ITEM_CODE").Text = "" Then
            e.Row.Cells("EDI_ITEM_CODE").Appearance.ForeColor = System.Drawing.Color.Red
        End If
    End Sub

    Private Sub grdEDT852TS_InitializeRow(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeRowEventArgs) Handles grdEDT852TS.InitializeRow
        If e.Row.Cells("CUST_STORE_NO").Text = "" Then
            e.Row.Cells("EDI_STORE_NO").Appearance.ForeColor = System.Drawing.Color.Red
        End If
    End Sub

    Private Sub optMode_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optMode.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        Select_tabMain()

        If tabMain.SelectedTab.Key = "852" Then
            grdRSTSLPI0.Parent = tabMain.Tabs("852").TabPage
            opt852Status.Parent = tabMain.Tabs("852").TabPage
            btn852Inbox.Parent = tabMain.Tabs("852").TabPage
            Setup_grdRSTSLPI0()
        ElseIf tabMain.SelectedTab.Key = "SPS" Then
            grdRSTSLPI0.Parent = tabMain.Tabs("SPS").TabPage
            opt852Status.Parent = tabMain.Tabs("SPS").TabPage
            btn852Inbox.Parent = tabMain.Tabs("SPS").TabPage
            Setup_grdRSTSLPI0()
        End If
    End Sub

    Sub Select_tabMain()
        'tabMain.VisibleTab
        tabMain.Tabs(optMode.Value).visible = True
        tabMain.SelectedTab = tabMain.Tabs(optMode.Value)
        'tabMain.SelectedTab.Visible = True
        For Each TAB As UltraWinTabControl.UltraTab In tabMain.Tabs
            If Not TAB.Selected Then
                TAB.Visible = False
            End If
        Next
    End Sub

    Sub Update_RSTRETLx(
    ByVal EDI_DOC_SEQ_NO As String,
    Optional ByVal plus_or_minus_override As String = "")

        Dim plus_or_minus As String = "+"
        If plus_or_minus_override <> "" Then
            plus_or_minus = plus_or_minus_override
        Else
            If opt852Data.Value = "1" Then
                plus_or_minus = "-"
            End If
        End If

        TAC.RSCMAIN1.Update_RSTRETLx(EDI_DOC_SEQ_NO, plus_or_minus)

    End Sub

    Private Sub grdEDT852TC_DoubleClickCell(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.DoubleClickCellEventArgs) Handles grdEDT852TC.DoubleClickCell
        Try
            opt852Data.Value = "1"
            Dim OPS_YYYYWW As String = e.Cell.Column.Tag
            Dim CUST_CODE As String = e.Cell.Row.Cells("CUST_CODE").Text
            With grdEDT852T1.DisplayLayout.Bands(0)
                .ColumnFilters.ClearAllFilters()
                .ColumnFilters("CUST_CODE").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, CUST_CODE)
                .ColumnFilters("OPS_YYYYWW").FilterConditions.Add(Infragistics.Win.UltraWinGrid.FilterComparisionOperator.Match, OPS_YYYYWW)

                Dim tlb_sbt As UltraWinToolbars.StateButtonTool = DirectCast(tlb.Tools("Show Filter"), UltraWinToolbars.StateButtonTool)
                tlb_sbt.Checked = True
                Show_Filter(grdEDT852T1, True)

            End With
        Catch ex As Exception

        End Try
    End Sub

#Region "grdRSTRETLA"

    Private Sub grdRSTRETLA_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdRSTRETLA.BeforeExitEditMode
        If e.CancellingEditOperation Then Exit Sub
        With grdRSTRETLA.ActiveCell
            Select Case .Column.Key
                Case "CUST_STORE_NO"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ARTCUST2", New String() {Absx1.txtFor("CUST_CODE").Text, .Text})
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            Dim CUST_STORE_LOCATION As String = cdr.Item("CUST_STORE_LOCATION") & ""
                            If CUST_STORE_LOCATION = "" Then CUST_STORE_LOCATION = cdr.Item("CUST_STORE_NAME") & ""
                            If cdr.Item("CUST_STORE_MARK_FOR") & "" <> "" Then
                                CUST_STORE_LOCATION &= " (" & cdr.Item("CUST_STORE_MARK_FOR") & ")"
                            End If

                            .Row.Cells("CUST_STORE_LOCATION").Value = CUST_STORE_LOCATION
                            .Row.Cells("CUST_STORE_CITY").Value = cdr.Item("CUST_STORE_CITY") & ""
                            .Row.Cells("CUST_STORE_STATE").Value = cdr.Item("CUST_STORE_STATE") & ""
                            .Row.Cells("SELL_CODE").Value = cdr.Item("SELL_CODE") & ""

                        End If
                    End If

                Case "COLLECTION_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTCOLL1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            If cdr.Item("BRAND_CODE") <> HFs("BRAND_CODE") Then
                                ASCMAIN1.Progress("Collection does not belong to Brand " & HFs("BRAND_CODE"))
                                .Value = ""
                                e.Cancel = True
                            End If
                        End If
                    End If

                Case "ITEM_CODE"

                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            Dim COLLECTION_CODE As String = cdr.Item("COLLECTION_CODE") & ""
                            cdr = LookUp("ICTCOLL1", COLLECTION_CODE)
                            If cdr.Item("BRAND_CODE") <> HFs("BRAND_CODE") Then
                                ASCMAIN1.Progress("Item's Collection (" & COLLECTION_CODE & " does not belong to Brand " & HFs("BRAND_CODE"))
                                .Value = ""
                                e.Cancel = True
                            End If
                            .Row.Cells("COLLECTION_CODE").Value = cdr.Item("COLLECTION_CODE") & ""
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdRSTRETLA_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTRETLA.BeforeRowUpdate

        If e.Row.Cells("CUST_STORE_NO").Value & "" = "" _
        Or (optCI.Value = "C" And optCOLLECTION_CODE.Value = "A" And e.Row.Cells("COLLECTION_CODE").Value & "" = "") _
        Or (optCI.Value = "I" And e.Row.Cells("ITEM_CODE").Value & "" = "") Then
            e.Cancel = True
        End If

        If optCOLLECTION_CODE.Value <> "I" Then
            cdr = LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Text & "")
            If cdr Is Nothing Then
                'ASCMAIN1.Progress("Invalid Collection Code (" & e.Row.Cells("COLLECTION_CODE").Text & ")")
                e.Cancel = True
            Else
                If cdr.Item("BRAND_CODE") <> HFs("BRAND_CODE") Then
                    'ASCMAIN1.Progress("Collection does not belong to Brand " & HFs("BRAND_CODE"))
                    e.Cancel = True
                End If
            End If
        End If


        If Val(e.Row.Cells("QTY_TOTAL").Value & "") = 0 And Val(e.Row.Cells("AMT_TOTAL").Value & "") = 0 And Val(e.Row.Cells("ONH_TOTAL").Value & "") = 0 Then
            ' e.Cancel = True
            ' PROBLEM WHEN ADDING A NEW ROW IN WEEKLY MODE - TOTALS ARE NOT CALCULATED YET
        End If

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
            If optCOLLECTION_CODE.Value = "I" Then
                e.Row.Cells("COLLECTION_CODE").Value = cbeCOLLECTION_CODE.Value
            End If
        End If
    End Sub

    Private Sub grdRSTRETLA_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETLA.ClickCellButton
        Dim sql_where As String = ""
        If grdRSTRETLA.ActiveCell Is Nothing Then
            Exit Sub
        End If
        Select Case grdRSTRETLA.ActiveCell.Column.Key
            Case "CUST_STORE_NO"
                grdClickCellButton(grdRSTRETLA, "CUST_CODE = '" & HFs("CUST_CODE") & "'")
            Case "ITEM_CODE"
                grdClickCellButton(grdRSTRETLA)
        End Select
    End Sub

    Private Sub grdRSTRETLA_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTRETLA.AfterRowActivate
        If grdRSTRETLA.ActiveRow.IsAddRow Then
            grdRSTRETLA.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
        Else
            grdRSTRETLA.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub grdRSTRETLA_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTRETLA.AfterExitEditMode
        With grdRSTRETLA

            If .ActiveCell Is Nothing Then
                Exit Sub
            End If
            Select Case .ActiveCell.Column.Key
                Case "CUST_STORE_NO"
                    If .ActiveCell.Text <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)

                        Dim CUST_STORE_LOCATION As String = cdr.Item("CUST_STORE_LOCATION") & ""
                        If CUST_STORE_LOCATION = "" Then CUST_STORE_LOCATION = cdr.Item("CUST_STORE_NAME") & ""
                        If cdr.Item("CUST_STORE_MARK_FOR") & "" <> "" Then
                            CUST_STORE_LOCATION &= " (" & cdr.Item("CUST_STORE_MARK_FOR") & ")"
                        End If

                        .ActiveRow.Cells("CUST_STORE_LOCATION").Value = CUST_STORE_LOCATION
                        .ActiveRow.Cells("CUST_STORE_CITY").Value = cdr.Item("CUST_STORE_CITY") & ""
                        .ActiveRow.Cells("CUST_STORE_STATE").Value = cdr.Item("CUST_STORE_STATE") & ""
                        .ActiveRow.Cells("SELL_CODE").Value = cdr.Item("SELL_CODE") & ""

                    End If
            End Select
        End With
    End Sub
#End Region

#Region "grdRSTRETLI"

    Private Sub grdRSTRETLI_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETLI.AfterCellUpdate
        If e.Cell.Column.Key = "ITEM_CODE" Then
            With grdRSTRETLI.ActiveRow.Cells("ITEM_CODE")
                cdr = LookUp("ICTITEM1", .Text)
                If cdr Is Nothing Then
                Else
                    .Row.Cells("ITEM_DESC").Value = cdr.Item("ITEM_DESC") & ""
                    Dim COLLECTION_CODE As String = cdr.Item("COLLECTION_CODE") & ""
                    cdr = LookUp("ICTCOLL1", COLLECTION_CODE, True)
                    'If cdr.Item("BRAND_CODE") <> HFs("BRAND_CODE") Then
                    '    ASCMAIN1.Progress("Item's Collection (" & COLLECTION_CODE & " does not belong to Brand " & HFs("BRAND_CODE"))
                    '    .Value = ""
                    '    e.Cancel = True
                    'End If

                    .Row.Cells("COLLECTION_CODE").Value = cdr.Item("COLLECTION_CODE") & ""
                End If
            End With

        End If
    End Sub

    Private Sub grdRSTRETLI_BeforeExitEditMode(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeExitEditModeEventArgs) Handles grdRSTRETLI.BeforeExitEditMode
        If e.CancellingEditOperation Then Exit Sub
        With grdRSTRETLI.ActiveCell
            Select Case .Column.Key
                Case "CUST_ITEM_CODE"
                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("RSTITEMX", New String() {Absx1.txtFor("CUST_CODE").Text, .Text})
                        If cdr Is Nothing Then
                            'ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            '.Value = ""
                            'e.Cancel = True
                        Else
                            Dim ITEM_CODE As String = cdr.Item("ITEM_CODE") & ""
                            ITEM_CODE = ITEM_CODE.ToUpper
                            .Row.Cells("ITEM_CODE").Value = ITEM_CODE
                        End If
                    End If

                Case "ITEM_CODE"

                    If .Text <> "" Then
                        .Value = ASCMAIN1.Format_Field(.Text, .Column.Key)
                        cdr = LookUp("ICTITEM1", .Text)
                        If cdr Is Nothing Then
                            ASCMAIN1.Progress("Invalid " & .Column.Header.Caption & "(" & .Text & ")")
                            .Value = ""
                            e.Cancel = True
                        Else
                            Dim COLLECTION_CODE As String = cdr.Item("COLLECTION_CODE") & ""
                            cdr = LookUp("ICTCOLL1", COLLECTION_CODE)
                            If cdr.Item("BRAND_CODE") <> HFs("BRAND_CODE") Then
                                ASCMAIN1.Progress("Item's Collection (" & COLLECTION_CODE & " does not belong to Brand " & HFs("BRAND_CODE"))
                                .Value = ""
                                e.Cancel = True
                            End If
                            .Row.Cells("COLLECTION_CODE").Value = cdr.Item("COLLECTION_CODE") & ""
                        End If
                    End If
            End Select
        End With
    End Sub

    Private Sub grdRSTRETLI_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles grdRSTRETLI.BeforeRowUpdate

        'If e.Row.Cells("CUST_STORE_NO").Value & "" = "" _
        'Or (optCI.Value = "C" And optCOLLECTION_CODE.Value = "A" And e.Row.Cells("COLLECTION_CODE").Value & "" = "") _
        'Or (optCI.Value = "I" And e.Row.Cells("ITEM_CODE").Value & "" = "") Then
        '    e.Cancel = True
        'End If

        'If optCOLLECTION_CODE.Value <> "I" Then
        '    cdr = LookUp("ICTCOLL1", e.Row.Cells("COLLECTION_CODE").Text & "")
        '    If cdr Is Nothing Then
        '        'ASCMAIN1.Progress("Invalid Collection Code (" & e.Row.Cells("COLLECTION_CODE").Text & ")")
        '        e.Cancel = True
        '    Else
        '        If cdr.Item("BRAND_CODE") <> HFs("BRAND_CODE") Then
        '            'ASCMAIN1.Progress("Collection does not belong to Brand " & HFs("BRAND_CODE"))
        '            e.Cancel = True
        '        End If
        '    End If
        'End If


        'If Val(e.Row.Cells("QTY_TOTAL").Value & "") = 0 And Val(e.Row.Cells("AMT_TOTAL").Value & "") = 0 And Val(e.Row.Cells("ONH_TOTAL").Value & "") = 0 Then
        '    e.Cancel = True
        'End If

        If e.Row.IsAddRow And Not e.Cancel Then
            e.Row.Cells("CUST_CODE").Value = Absx1.txtFor("CUST_CODE").Text
            If optCOLLECTION_CODE.Value = "I" Then
                e.Row.Cells("COLLECTION_CODE").Value = cbeCOLLECTION_CODE.Value
            End If
        End If
    End Sub

    Private Sub grdRSTRETLI_ClickCellButton(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles grdRSTRETLI.ClickCellButton
        Dim sql_where As String = ""
        Select Case grdRSTRETLI.ActiveCell.Column.Key
            Case "ITEM_CODE"
        End Select

        Call grdClickCellButton(grdRSTRETLI, sql_where, False)
    End Sub

    'Private Sub grdRSTRETLI_BeforeRowsDeleted(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowsDeletedEventArgs) Handles grdRSTRETLI.BeforeRowsDeleted
    '    For Each grow As UltraWinGrid.UltraGridRow In grdRSTRETLI.Selected.Rows
    '        If dst.Tables("RSTRETLI").Rows(grow.ListIndex).RowState = DataRowState.Added Then
    '        Else
    '            MsgBox("Cannot Delete Existing Store Records", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
    '            e.Cancel = True
    '            Exit For
    '        End If
    '    Next
    'End Sub

    Private Sub grdRSTRETLI_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTRETLI.AfterRowActivate
        'If grdRSTRETLI.ActiveRow.IsAddRow Then
        '    grdRSTRETLI.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.AllowEdit
        'Else
        '    grdRSTRETLI.DisplayLayout.Bands(0).Columns("CUST_STORE_NO").CellActivation = UltraWinGrid.Activation.NoEdit
        'End If
        'Stop
    End Sub

    Private Sub grdRSTRETLI_AfterExitEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdRSTRETLI.AfterExitEditMode
        With grdRSTRETLI
            Select Case .ActiveCell.Column.Key
                Case "ITEM_CODE"
                    If .ActiveCell.Text <> "" Then
                        .ActiveCell.Value = ASCMAIN1.Format_Field(.ActiveCell.Text, .ActiveCell.Column.Key)

                        'Dim CUST_STORE_LOCATION As String = cdr.Item("CUST_STORE_LOCATION") & ""
                        'If CUST_STORE_LOCATION = "" Then CUST_STORE_LOCATION = cdr.Item("CUST_STORE_NAME") & ""
                        'If cdr.Item("CUST_STORE_MARK_FOR") & "" <> "" Then
                        '    CUST_STORE_LOCATION &= " (" & cdr.Item("CUST_STORE_MARK_FOR") & ")"
                        'End If

                        '.ActiveRow.Cells("CUST_STORE_LOCATION").Value = CUST_STORE_LOCATION
                        '.ActiveRow.Cells("CUST_STORE_CITY").Value = cdr.Item("CUST_STORE_CITY") & ""
                        '.ActiveRow.Cells("CUST_STORE_STATE").Value = cdr.Item("CUST_STORE_STATE") & ""
                        '.ActiveRow.Cells("SELL_CODE").Value = cdr.Item("SELL_CODE") & ""

                    End If
            End Select
        End With
    End Sub
#End Region

    Private Sub optCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If chkByStore.Checked Then
            Setup_grdRSTRETLA()
        Else
            Setup_grdRSTRETLI()
        End If
    End Sub

    Sub Setup_grdRSTRETLA()

        cbeCOLLECTION_CODE.Enabled = (optCOLLECTION_CODE.Value = "I")
        'grdRSTBUDR1.DisplayLayout.GroupByBox.Hidden = (optCOLLECTION_CODE.Value = "A")

        With grdRSTRETLA.DisplayLayout.Bands(0)
            With .Columns("COLLECTION_CODE")
                .Hidden = (optCOLLECTION_CODE.Value <> "A")
                If optCOLLECTION_CODE.Value = "A" And optCI.Value = "C" Then
                    .CellAppearance.BackColor = System.Drawing.Color.Empty
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    cbeCOLLECTION_CODE.DropDownStyle = DropDownStyle.DropDown
                Else
                    .CellAppearance.BackColor = System.Drawing.Color.Beige
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    cbeCOLLECTION_CODE.DropDownStyle = DropDownStyle.DropDownList
                End If

            End With

            .SortedColumns.Clear()
            If optCOLLECTION_CODE.Value = "A" Then
                'grdRSTBUDR1.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
            End If
            .SortedColumns.Add("CUST_STORE_NO", False)
        End With

        Dim COLLS As String = ""
        Dim allow_modifications As Boolean = True

        Dim DVW As DataView = DirectCast(grdRSTRETLA.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If optCOLLECTION_CODE.Value = "A" Then
            COLLS = "All Collections"
        Else
            sql = "and COLLECTION_CODE = '" & cbeCOLLECTION_CODE.Value & "'"
            COLLS = cbeCOLLECTION_CODE.Value
        End If

        DVW.RowFilter = Mid(sql, 5)
        grdRSTRETLA.Text = "Retail Sales by Store for " & Absx1.cmbFor("RYP").Text & " - " & COLLS
        If allow_modifications Then
            grdRSTRETLA.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Else
            grdRSTRETLA.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If
    End Sub
    Sub Setup_grdRSTRETLI()

        cbeCOLLECTION_CODE.Enabled = (optCOLLECTION_CODE.Value = "I")
        'grdRSTBUDR1.DisplayLayout.GroupByBox.Hidden = (optCOLLECTION_CODE.Value = "A")

        With grdRSTRETLI.DisplayLayout.Bands(0)
            With .Columns("COLLECTION_CODE")
                .Hidden = (optCOLLECTION_CODE.Value <> "A")
                If optCOLLECTION_CODE.Value = "A" Then
                    .CellAppearance.BackColor = System.Drawing.Color.Empty
                    .CellActivation = UltraWinGrid.Activation.AllowEdit
                    cbeCOLLECTION_CODE.DropDownStyle = DropDownStyle.DropDown
                Else
                    .CellAppearance.BackColor = System.Drawing.Color.Beige
                    .CellActivation = UltraWinGrid.Activation.NoEdit
                    cbeCOLLECTION_CODE.DropDownStyle = DropDownStyle.DropDownList
                End If

                .Hidden = True ' ALWAYS FOR THIS GRID

            End With

            .SortedColumns.Clear()
            If optCOLLECTION_CODE.Value = "A" Then
                'grdRSTBUDR1.DisplayLayout.Bands(0).SortedColumns.Add("COLLECTION_CODE", False, True)
            End If
            '.SortedColumns.Add("CUST_STORE_NO", False)
        End With

        Dim COLLS As String = ""
        Dim allow_modifications As Boolean = True

        Dim DVW As DataView = DirectCast(grdRSTRETLI.DataSource, DataTable).DefaultView
        Dim sql As String = ""
        If optCOLLECTION_CODE.Value = "A" Then
            COLLS = "All Collections"
        Else
            sql = "and COLLECTION_CODE = '" & cbeCOLLECTION_CODE.Value & "'"
            COLLS = cbeCOLLECTION_CODE.Value
        End If

        DVW.RowFilter = Mid(sql, 5)
        grdRSTRETLI.Text = "Retail Sales by Item for " & IIf(chk1Wk.Checked, Absx1.cmbFor("RYW").Text, Absx1.cmbFor("RYP").Text) & " - " & COLLS
        If allow_modifications Then
            grdRSTRETLI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.True
        Else
            grdRSTRETLI.DisplayLayout.Override.AllowUpdate = DefaultableBoolean.False
        End If
    End Sub

    Private Sub cbeCOLLECTION_CODE_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbeCOLLECTION_CODE.ValueChanged
        If SELECTION_NO = 0 Then Exit Sub
        If chkByStore.Checked Then
            Setup_grdRSTRETLA()
        Else
            Setup_grdRSTRETLI()
        End If
    End Sub

    Private Sub chkEntryUnits_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEntryUnits.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_QA()
    End Sub

    Private Sub chkEntryDollars_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEntryDollars.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_QA()
    End Sub

    Sub Setup_QA()
        Dim grd As UltraWinGrid.UltraGrid
        If chkByStore.Checked Then
            grd = grdRSTRETLA
        Else
            grd = grdRSTRETLI
        End If

        Dim Only_Week As Integer = 0
        If chk1Wk.Checked Then
            For W As Integer = 1 To MAX_WEEKS
                If WKS(W) = Absx1.cmbFor("RYW").Value Then
                    Only_Week = W
                End If
            Next W
        End If

        With grd.DisplayLayout.Bands(0)
            For W As Integer = 1 To MAX_WEEKS
                .Columns("QTY_W" & Format(W, "0")).Hidden = Not chkEntryUnits.Checked Or optWM.Value = "M" Or (chk1Wk.Checked And W <> Only_Week)
                .Columns("AMT_W" & Format(W, "0")).Hidden = Not chkEntryDollars.Checked Or optWM.Value = "M" Or (chk1Wk.Checked And W <> Only_Week)
                .Columns("ONH_W" & Format(W, "0")).Hidden = Not chkEntryOnHand.Checked Or optWM.Value = "M" Or (chk1Wk.Checked And W <> Only_Week)
            Next
            .Columns("QTY_TOTAL").Hidden = Not chkEntryUnits.Checked Or chk1Wk.Checked
            .Columns("AMT_TOTAL").Hidden = Not chkEntryDollars.Checked Or chk1Wk.Checked
            .Columns("ONH_TOTAL").Hidden = Not chkEntryOnHand.Checked Or chk1Wk.Checked
        End With
    End Sub

    Private Sub chkEntryOnHand_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEntryOnHand.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        Setup_QA()
    End Sub

    Private Sub chkEntryDescription_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkEntryDescription.CheckedChanged
        If SELECTION_NO = 0 Then Exit Sub
        'grdRSTRETLA.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden = Not chkEntryDescription.Checked
        If chkByStore.Checked Then
            grdRSTRETLA.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden = Not chkEntryDescription.Checked
        Else
            grdRSTRETLI.DisplayLayout.Bands(0).Columns("ITEM_DESC").Hidden = Not chkEntryDescription.Checked
        End If
    End Sub

    Private Sub chk1Wk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chk1Wk.CheckedChanged
        Absx1.CtlFor("RYP").Visible = Not chk1Wk.Checked
        Absx1.CtlFor("RYW").Visible = chk1Wk.Checked
    End Sub

    Private Sub optWM_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optWM.ValueChanged
        If optWM.Value = "M" Then
            chk1Wk.Checked = False
        End If
        chk1Wk.Visible = (optWM.Value = "W")
    End Sub

    Private Sub cmdCalculateOH_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCalculateOH.Click

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Calculating End of Week from Prior Week less This Week's Sales")

        If Not chk1Wk.Checked Then
            MsgBox("This Feature is available only when Manually Entering 1 week at a time")
            Exit Sub
        End If

        Dim RYW As String = Absx1.cmbFor("RYW").Value

        Dim TABLE_NAME As String = ""
        If chkByStore.Checked Then
            TABLE_NAME = "RSTRETLA"
        Else
            TABLE_NAME = "RSTRETLI"
        End If

        ASCMAIN1.sql = "Select * from GLTPARM3 where YYYYWW = '" & RYW & "'"
        Dim rowGLTPARM3 As DataRow = ASCDATA1.GetDataRow
        Dim REL_WEEK As Int32 = Val(rowGLTPARM3.Item("REL_WEEK") & "")

        For Each row As DataRow In dst.Tables(TABLE_NAME).Rows
            row("ONH_W" & Format(REL_WEEK, "0")) = 0
        Next

        Dim CUST_CODE As String = HFs("CUST_CODE")


        ASCMAIN1.sql = "SELECT " & IIf(chkByStore.Checked, "CUST_STORE_NO", "'000000' CUST_STORE_NO") _
        & ", ITEM_CODE, SUM (QTY_EOW) QTY_EOW" _
        & " from RSTRETL1 WHERE CUST_CODE = '" & CUST_CODE & "'" _
        & " and OPS_YYYYWW = '" & ASCMAIN1.Week_Calc(RYW, -1) & "'" _
        & " and QTY_EOW <> 0" _
        & " group by " & IIf(chkByStore.Checked, "CUST_STORE_NO", "'000000'") & ",ITEM_CODE"

        For Each row As DataRow In ASCDATA1.GetDataTable.Rows
            'Dim row2 As DataRow = dst.Tables("").Rows.Find("")

            'Stop
            Dim CUST_STORE_NO As String = row.Item("CUST_STORE_NO") & ""
            Dim ITEM_CODE As String = row.Item("ITEM_CODE") & ""
            Dim QTY_EOW As String = row.Item("QTY_EOW") & ""

            If chkByStore.Checked Then
                Stop
                'MsgBox("this option has not been completely tested when doing OH by store")
            Else
                Dim rows As DataRow() = dst.Tables(TABLE_NAME).Select("ITEM_CODE = '" & ITEM_CODE & "'")
                If rows.Length = 0 Then
                    Dim row2 As DataRow = dst.Tables(TABLE_NAME).NewRow
                    row2.Item("CUST_CODE") = CUST_CODE
                    'row2.Item("CUST_STORE_NO") = CUST_STORE_NO
                    row2.Item("ITEM_CODE") = ITEM_CODE
                    row2.Item("ONH_W" & Format(REL_WEEK, "0")) = QTY_EOW
                    dst.Tables(TABLE_NAME).Rows.Add(row2)
                Else
                    Dim SLS As Int32 = Val(rows(0).Item("QTY_W" & Format(REL_WEEK, "0")) & "")
                    QTY_EOW = QTY_EOW - SLS
                    rows(0).Item("ONH_W" & Format(REL_WEEK, "0")) = QTY_EOW
                End If
            End If
        Next

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Private Sub cmdCustomLoad_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCustomLoad.Click
        If cbeCustomLoad.Value & "" = "" Then
            MsgBox("No Custom Load Template Selected", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
            Exit Sub
        End If

        Select Case cbeCustomLoad.Value


            Case "Nordstrom"

                If optCI.Value <> "I" Or optWM.Value <> "W" Or Not chk1Wk.Checked Or Not chkByStore.Checked Then
                    MsgBox("Nordstrom Custom Load is for 1 Week, by Item, By Store", MsgBoxStyle.OkOnly, "Cannot Perform Requested Action")
                    Exit Sub
                End If

                ImportNordstromSheet()

        End Select

    End Sub

    Sub ImportNordstromSheet()
        Dim fileName As String = ""
        Using openFileDialog1 As New OpenFileDialog
            openFileDialog1.Title = "Select an Excel Spreadsheet to Import"
            openFileDialog1.Filter = "Excel files (*.xlsx,*.xls,*.xlsm)|*.xlsx;*.xls;*.xlsm|Other|*.*"

            openFileDialog1.RestoreDirectory = True

            If openFileDialog1.ShowDialog() = DialogResult.OK Then
                fileName = openFileDialog1.FileName
            End If
        End Using

        ASCMAIN1.Progress("Reading file...")


        Dim dt As DataTable
        Try
            dt = LoadNordstromFile(fileName)
        Catch ex As Exception
            MsgBox("Error occurred while trying to read Nordstrom file")
            Exit Sub
        End Try

        ASCMAIN1.Progress("Processing file...")
        If dt IsNot Nothing Then


            Dim TBL_BAD_ROWS As New DataTable
            With TBL_BAD_ROWS
                .Columns.Add("STORE_NO")
                .Columns.Add("ITEM_CODE")
                .Columns.Add("ITEM_DESC")
                .Columns.Add("SLS", GetType(System.Int16))
                .Columns.Add("ONH", GetType(System.Int16))
            End With

            dst.Tables("RSTRETLA").Rows.Clear()

            Dim GOOD_ROWS As Integer = 0
            Dim BAD_ROWS As Integer = 0
            Dim ITEM_CODE As String = ""
            Dim ITEM_DESC As String = ""
            Dim COLLECTION_CODE As String = ""
            Dim BAD_ITEM_CODEs As String = ""

            Dim RYW As String = Absx1.cmbFor("RYW").Value
            Dim rowGLTPARM3 As DataRow = LookUp("GLTPARM3", RYW)
            RYP = rowGLTPARM3.Item("YYYYPP")
            Dim REL_WEEK As Integer = Val(rowGLTPARM3.Item("REL_WEEK") & "")

            Dim CUST_STORE_NO As String = ""
            Dim rowARTCUST2 As DataRow = Nothing

            Dim CUST_CODE As String = HFs("CUST_CODE")
            Dim BAD_ITEM_CTR As Integer = 0

            For Each row As DataRow In dt.Rows

                CUST_STORE_NO = row.Item("CUST_STORE_NO")
                rowARTCUST2 = LookUp("ARTCUST2", New String() {CUST_CODE, CUST_STORE_NO})

                If CUST_STORE_NO <> "" Then
                    Dim RETAIL As Decimal = If(chkEntryDollars.Checked, row.Item("AMT_SOLD"), 0)
                    Dim QTY_Wx As Integer = If(chkEntryUnits.Checked, row.Item("QTY_SOLD"), 0)
                    Dim ONH_Wx As Integer = If(chkEntryOnHand.Checked, row.Item("QTY_EOW"), 0)

                    Dim CUST_ITEM_CODE As String = row.Item("ITEM_CODE") & ""
                    Dim VEND_STYLE_DESC As String = row.Item("ITEM_DESC") & ""
                    ITEM_CODE = ""
                    COLLECTION_CODE = ""
                    ITEM_DESC = ""

                    If QTY_Wx <> 0 Or ONH_Wx <> 0 Then
                        Dim rowEDTUPCX1 As DataRow = Nothing
                        Dim rowICTITEM1 As DataRow = ASCDATA1.GetDataRow("SELECT * FROM ICTITEM1 WHERE ITEM_ALT_SORT=:PARM1", "V", CUST_ITEM_CODE)
                        If rowICTITEM1 Is Nothing Then
                            rowICTITEM1 = LookUp("ICTITEM1", CUST_ITEM_CODE)
                        End If
                        If rowICTITEM1 IsNot Nothing Then
                            ITEM_CODE = rowICTITEM1.Item("ITEM_CODE")
                            COLLECTION_CODE = rowICTITEM1.Item("COLLECTION_CODE")
                            ITEM_DESC = rowICTITEM1.Item("ITEM_DESC")
                        End If
                        '	FLAG


                        If ITEM_CODE = "" Then
                            BAD_ITEM_CODEs &= vbCrLf & CUST_ITEM_CODE
                            BAD_ROWS += 1
                            If rowEDTUPCX1 Is Nothing Then
                                Dim rowBAD_ROWS As DataRow = TBL_BAD_ROWS.NewRow
                                rowBAD_ROWS("STORE_NO") = CUST_STORE_NO
                                rowBAD_ROWS("ITEM_CODE") = CUST_ITEM_CODE
                                rowBAD_ROWS("ITEM_DESC") = VEND_STYLE_DESC
                                rowBAD_ROWS("SLS") = QTY_Wx
                                rowBAD_ROWS("ONH") = ONH_Wx

                                TBL_BAD_ROWS.Rows.Add(rowBAD_ROWS)
                            End If
                        Else
                            GOOD_ROWS += 1
                        End If

                        Dim rowRSTRETLA As DataRow

                        Dim sqlwc As String = "CUST_CODE = '" & CUST_CODE & "'" _
                        & " and CUST_STORE_NO = '" & CUST_STORE_NO & "'" _
                        & " and ITEM_CODE = '" & ITEM_CODE & "'"

                        If ITEM_CODE = "" Then
                            sqlwc = "CUST_CODE = '" & CUST_CODE & "'" _
                            & " and CUST_STORE_NO = '" & CUST_STORE_NO & "'" _
                            & " and CUST_ITEM_CODE = '" & CUST_ITEM_CODE & "'"
                        End If

                        Dim rowRSTRETLAs() As DataRow = dst.Tables("RSTRETLA").Select(sqlwc)
                        If rowRSTRETLAs.Length <> 0 Then
                            rowRSTRETLA = rowRSTRETLAs(0)

                            rowRSTRETLA.Item("QTY_W" & CStr(REL_WEEK)) = Val(rowRSTRETLA.Item("QTY_W" & CStr(REL_WEEK)) & "") + QTY_Wx
                            rowRSTRETLA.Item("AMT_W" & CStr(REL_WEEK)) = Val(rowRSTRETLA.Item("AMT_W" & CStr(REL_WEEK)) & "") + RETAIL
                            rowRSTRETLA.Item("ONH_W" & CStr(REL_WEEK)) = Val(rowRSTRETLA.Item("ONH_W" & CStr(REL_WEEK)) & "") + ONH_Wx
                        Else

                            rowRSTRETLA = dst.Tables("RSTRETLA").NewRow
                            rowRSTRETLA.Item("CUST_CODE") = CUST_CODE
                            rowRSTRETLA.Item("CUST_STORE_NO") = CUST_STORE_NO
                            rowRSTRETLA.Item("COLLECTION_CODE") = COLLECTION_CODE
                            rowRSTRETLA.Item("ITEM_CODE") = ITEM_CODE
                            If ITEM_CODE = "" Then
                                BAD_ITEM_CTR += 1
                                rowRSTRETLA.Item("ITEM_CODE") = " BAD ITEM " & Format(BAD_ITEM_CTR, "000000")
                            End If
                            rowRSTRETLA.Item("QTY_W" & CStr(REL_WEEK)) = QTY_Wx
                            rowRSTRETLA.Item("AMT_W" & CStr(REL_WEEK)) = RETAIL
                            rowRSTRETLA.Item("ONH_W" & CStr(REL_WEEK)) = ONH_Wx

                            If rowARTCUST2 IsNot Nothing Then
                                rowRSTRETLA.Item("CUST_STORE_LOCATION") = rowARTCUST2.Item("CUST_STORE_LOCATION")
                                rowRSTRETLA.Item("CUST_STORE_CITY") = rowARTCUST2.Item("CUST_STORE_CITY")
                                rowRSTRETLA.Item("CUST_STORE_STATE") = rowARTCUST2.Item("CUST_STORE_STATE")
                                rowRSTRETLA.Item("SELL_CODE") = rowARTCUST2.Item("SELL_CODE")
                            End If

                            rowRSTRETLA.Item("CUST_ITEM_CODE") = CUST_ITEM_CODE
                            rowRSTRETLA.Item("ITEM_DESC") = ITEM_DESC

                            dst.Tables("RSTRETLA").Rows.Add(rowRSTRETLA)


                        End If

                    End If
                End If
            Next

            ASCMAIN1.Progress("")

            If BAD_ROWS = 0 Then
                MsgBox(CStr(GOOD_ROWS) & " Loaded", MsgBoxStyle.OkOnly, "Verification")
            Else
                MsgBox("Good Rows: " & CStr(GOOD_ROWS) & vbCrLf & "Bad Rows: " & CStr(BAD_ROWS), MsgBoxStyle.OkOnly, "Custom Load Complete")

                If TBL_BAD_ROWS.Rows.Count <> 0 Then
                    Dim Fmsg As New ASFMSGBF
                    Fmsg.Show_grd(TBL_BAD_ROWS, ASCMAIN1.ActiveForm, "Bad Items from Spreadsheet")
                End If
            End If
        End If
    End Sub

    Private Function LoadNordstromFile(ByVal fileName As String) As DataTable
        Dim iWorkbook As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(fileName, System.Globalization.CultureInfo.CurrentCulture)

        Dim activeSheet As SpreadsheetGear.IWorksheet = iWorkbook.Worksheets(0)

        'Find relevant columns and load data into RSTRETL1
        Dim columnStoreNo As Integer = 0
        Dim columnItemCode As Integer = 0
        Dim columnQtySold As Integer = 0
        Dim columnAmtSold As Integer
        Dim columnQtyEow As Integer = 0

        Dim startRow As Integer = 1

        For i As Integer = 0 To activeSheet.UsedRange.ColumnCount
            Dim cellValue = activeSheet.Cells(startRow, i).Value
            Select Case cellValue
                Case "Store ID"
                    columnStoreNo = i
                Case "VPN"
                    columnItemCode = i
                Case "Sales U"
                    columnQtySold = i
                Case "Sales $"
                    columnAmtSold = i
                Case "EOH U"
                    columnQtyEow = i
            End Select
        Next


        startRow += 1

        Dim dtRSTRETL1 As DataTable = dst.Tables("RSTRETL1").Clone()

        dtRSTRETL1.Columns.Add("ITEM_DESC")

        For i = startRow To activeSheet.UsedRange.RowCount - 1

            Dim salesDataRow As DataRow = dtRSTRETL1.NewRow

            salesDataRow("EDI_DOC_SEQ_NO") = "0"
            salesDataRow("CUST_CODE") = "NORDSTROM"
            salesDataRow("CUST_STORE_NO") = activeSheet.Cells(i, columnStoreNo).Value.ToString().PadLeft(6, "0")
            salesDataRow("ITEM_CODE") = activeSheet.Cells(i, columnItemCode).Value.Split(",")(0)
            salesDataRow("ITEM_DESC") = activeSheet.Cells(i, columnItemCode).Value.Split(",")(1)
            salesDataRow("QTY_SOLD") = Val(activeSheet.Cells(i, columnQtySold).Value)
            salesDataRow("AMT_SOLD") = Val(activeSheet.Cells(i, columnAmtSold).Value)
            salesDataRow("OPS_YYYYPP") = UltraCombo2.SelectedRow.Cells("YYYYPP").Value
            salesDataRow("OPS_YYYYWW") = UltraCombo2.Value
            salesDataRow("QTY_EOW") = Val(activeSheet.Cells(i, columnQtyEow).Value)

            dtRSTRETL1.Rows.Add(salesDataRow)
        Next


        Return dtRSTRETL1
    End Function

    Sub Prepare_EDI_Data()
        ASCDATA1.ExecuteSQL("Update EDT852T1 Set EDI_STATUS = '0', EDI_TP_QUAL = TRIM(EDI_TP_QUAL), EDI_TP_ID = TRIM(EDI_TP_ID) where EDI_STATUS is Null")

        ASCMAIN1.sql = "Update EDT852T1 Set COMPANY_CODE = (Select COMPANY_CODE from EDTTRPMC" _
               & " where TRIM(EDI_TP_ID) = TRIM(EDT852T1.EDI_TP_ID))" _
               & " where EDI_STATUS = '0' and COMPANY_CODE IS NULL"
        ASCDATA1.ExecuteSQL()

        ' NEED TO DEAL WITH THIS IN RSTRETL1 - A SINGLE EDT852T1 RECORD FOR NORDSTROM WILL WIND UP INCLUDING STORES FOR NORDDIRECT
        'ASCMAIN1.sql = "Update EDT852T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM3" _
        '       & " where EDI_TP_QUAL = TRIM(EDT852T1.EDI_TP_QUAL) and EDI_TP_ID = TRIM(EDT852T1.EDI_TP_ID) and EDI_DOC_NO = '852' and EDI_STORE = (SELECT MIN(EDI_SDQ_STORE_01) FROM EDT852T3 WHERE EDI_DOC_SEQ_NO =  EDT852T1.EDI_DOC_SEQ_NO))" _
        '       & " where EDI_STATUS = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        'ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT852T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM2" _
               & " where EDI_TP_QUAL = TRIM(EDT852T1.EDI_TP_QUAL) and EDI_TP_ID = TRIM(EDT852T1.EDI_TP_ID) and EDI_DOC_NO = '852' and EDI_DEPT_NO = TRIM(EDT852T1.EDI_DEPT_NO))" _
               & " where EDI_STATUS = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Update EDT852T1 Set CUST_CODE = (Select CUST_CODE from EDTTRPM1" _
               & " where EDI_TP_QUAL = TRIM(EDT852T1.EDI_TP_QUAL) and EDI_TP_ID = TRIM(EDT852T1.EDI_TP_ID) and EDI_DOC_NO = '852')" _
               & " where EDI_STATUS = '0' and CUST_CODE IS NULL and COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "'"
        ASCDATA1.ExecuteSQL()

        ASCMAIN1.sql = "Select * from EDT852T1 where EDI_STATUS = '0' and COMPANY_CODE is Null"
        If EDI_DOC_SEQ_NOs_no_company.Count <> 0 Then
            ASCMAIN1.sql &= " and EDI_DOC_SEQ_NO Not in ('" & Join(EDI_DOC_SEQ_NOs_no_company.ToArray, "','") & "')"
        End If
        Dim dt As DataTable = ASCDATA1.GetDataTable
        If dt.Rows.Count <> 0 Then
            For Each row As DataRow In dt.Rows
                EDI_DOC_SEQ_NOs_no_company.Add(row.Item("EDI_DOC_SEQ_NO"))
            Next
            Using frm As New ASFMSGBF
                frm.Show_grd(dt, Me, "EDI Transactions which could not be mapped to an ABSolution Company")
            End Using
        End If

        ASCMAIN1.sql = "Select GLOBAL_LOCATION_NUMBER, CUST_CODE, CUST_STORE_NO from ARTCUST2 where GLOBAL_LOCATION_NUMBER is Not Null"
        EDT852TG = ASCMAIN1.Temp_Table()
        Try
            ASCDATA1.ExecuteSQL("Alter Table " & EDT852TG & " Add Primary Key (GLOBAL_LOCATION_NUMBER)")

        Catch ex As Exception
            ASCMAIN1.sql = "Select * from " & EDT852TG & " where GLOBAL_LOCATION_NUMBER in (Select GLOBAL_LOCATION_NUMBER from " & EDT852TG & " group by GLOBAL_LOCATION_NUMBER having Count (*) > 1)"
            Dim tbl As DataTable = ASCDATA1.GetDataTable
            Using frm As New ASFMSGBF

                frm.Show_grd(tbl, Me, "Duplicate GLNs - Do Not Process 852s")
            End Using
        End Try
    End Sub

    Private Sub grdRSTSLPI0_DoubleClickRow(sender As Object, e As UltraWinGrid.DoubleClickRowEventArgs) Handles grdRSTSLPI0.DoubleClickRow
        If opt852Status.Value = "0" Then
            Dim DOC_CTL_NO As String = grdRSTSLPI0.ActiveRow.Cells("DOC_CTL_NO").Value

            If MsgBox($"OK to Process Doc Ctl No {DOC_CTL_NO}?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

            Dim DOC_SOURCE As String = optMode.Value
            Select Case DOC_SOURCE
                Case "852"
                    Process_852(DOC_CTL_NO)
                Case "SPS"
                    Process_SPS(DOC_CTL_NO)

            End Select

            Setup_grdRSTSLPI0()
        End If
    End Sub

    Private Sub opt852Status_ValueChanged(sender As Object, e As EventArgs) Handles opt852Status.ValueChanged
        If Me.SELECTION_NO = 0 Then Exit Sub
        Setup_grdRSTSLPI0()
    End Sub

    Sub Setup_grdRSTSLPI0()
        Dim DOC_SOURCE As String = optMode.Value
        Fill_Records("RSTSLPI0", New String() {opt852Status.Value, DOC_SOURCE})
        Sort_grdColumns(grdRSTSLPI0, "DOC_CTL_NO")
        grdRSTSLPI0.Text = $"{DOC_SOURCE} Log - " & opt852Status.Text
        If opt852Status.Value = "0" Then
            grdRSTSLPI0.Text &= " - Double Click to Import from Zip"
        End If
    End Sub

    Sub Process_SPS(DOC_CTL_NO As String)
        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Importing")
        Dim rowRSTSLPI0 As DataRow = dst.Tables("RSTSLPI0").Rows.Find(DOC_CTL_NO)
        Dim FILENAME As String = rowRSTSLPI0.Item("FILENAME")

        dst.Tables("EDT852T1").Rows.Clear()
        dst.Tables("RSTRETL1").Rows.Clear()
        dst.Tables("RSTXLSQE").Rows.Clear()
        Fill_Records("ICTITEMU")

        Dim customerCodes As New Dictionary(Of String, String)
        customerCodes.Add("BLUEMERCURY", "BLUEME")
        customerCodes.Add("BlueMercuryPortal", "BLUEME")
        customerCodes.Add("CosBarPortal", "COSBAR")
        customerCodes.Add("FortressBrands3PAmazon", "AMAZON")
        customerCodes.Add("NEIMAN MARCUS", "NEIMANS")
        Dim CC As String = ""
        Dim t1Dictionary As New Dictionary(Of String, DataRow)
        Dim rowEDT852T2 As DataRow = Nothing
        Dim rowEDT852T3 As DataRow = Nothing
        Dim fp As String = $"\\ABSNASQ\Public\SLP\Share\SLP\SellThru_Data\SPS\Archive\{FILENAME}"

        Dim wb As SpreadsheetGear.IWorkbook = SpreadsheetGear.Factory.GetWorkbook(fp)
        Dim ws As IWorksheet = wb.Worksheets(0)
        Dim cells As IRange = ws.UsedRange
        Dim STARTING_ROW As Integer = 1
        Dim CUST_CODE As String = ""
        For row As Integer = STARTING_ROW To cells.RowCount - 1
            If Not customerCodes.ContainsKey(cells(row, 0).Value) Then
                Dim invalidCustMessage As String = $"Invalid Customer Code ({cells(row, 0).Value}) encountered."
                MsgBox(invalidCustMessage, vbCritical + vbOKOnly, "Cannot Proceed")
                Me.Cursor = Cursors.WaitCursor
                ASCMAIN1.Progress("")
                Exit Sub
            End If
            CUST_CODE = customerCodes(cells(row, 0).Value)
            If CC <> CUST_CODE Then
                CC = CUST_CODE
                Dim Sql As String = $"Select * from EDTUPCX1 WHERE CUST_CODE = '{CUST_CODE}'"
                Fill_Records("EDTUPCX1", "", True, Sql)
            End If

            Dim CUST_STORE_NO As String = cells(row, 3).Value.ToString.Substring(cells(row, 3).Value.ToString.Length - 6)
            Dim ITEM_UPC_CODE_XLS As String = cells(row, 3).Value
            If ITEM_UPC_CODE_XLS.Trim() <> "000000000000" Then

                Dim ITEM_CODE As String = TAC.RSCXLSC1.Validate_Item_Code(Me, ITEM_UPC_CODE_XLS, CUST_CODE)
                If ITEM_CODE <> "" AndAlso ITEM_CODE <> "IGNORE" Then
                    ASCMAIN1.Progress("Now Importing", ITEM_CODE)
                    Dim fieldToUpdate As String = cells(row, 4).Value
                    Dim firstWeekColumn As Integer = 5

                    For col As Integer = firstWeekColumn To cells.ColumnCount - 1
                        Dim weekEndingDate As Date = CDate(cells(0, col).Text)
                        Dim weekStartingDate As Date = weekEndingDate.AddDays(-6)
                        Dim t1DictKey As String = CUST_CODE & weekEndingDate.ToString("MMddyy")

                        Dim rowEDT852T1 As DataRow = Nothing
                        If Not t1Dictionary.ContainsKey(t1DictKey) Then
                            rowEDT852T1 = TAC.RSCXLSC1.Create_EDT852T1(Me, CUST_CODE, weekStartingDate, weekEndingDate, "SPS")
                            t1Dictionary.Add(t1DictKey, rowEDT852T1)
                        Else
                            rowEDT852T1 = t1Dictionary(t1DictKey)
                        End If
                        Dim EDI_DOC_SEQ_NO As String = rowEDT852T1("EDI_DOC_SEQ_NO")
                        Dim rowRSTRETL1 As DataRow = dst.Tables("RSTRETL1").Rows.Find(New String() {EDI_DOC_SEQ_NO, CUST_CODE, CUST_STORE_NO, ITEM_CODE})
                        If rowRSTRETL1 Is Nothing Then
                            rowRSTRETL1 = dst.Tables("RSTRETL1").NewRow()
                            rowRSTRETL1("EDI_DOC_SEQ_NO") = rowEDT852T1("EDI_DOC_SEQ_NO")
                            rowRSTRETL1("CUST_CODE") = CUST_CODE
                            rowRSTRETL1("CUST_STORE_NO") = CUST_STORE_NO
                            rowRSTRETL1("ITEM_CODE") = ITEM_CODE
                            rowRSTRETL1("OPS_YYYYPP") = rowEDT852T1("OPS_YYYYPP")
                            rowRSTRETL1("OPS_YYYYWW") = rowEDT852T1("OPS_YYYYWW")
                            dst.Tables("RSTRETL1").Rows.Add(rowRSTRETL1)
                        End If
                        Select Case fieldToUpdate
                            Case "TY EOW UNITS OH"
                                rowRSTRETL1("QTY_EOW") = Val(rowRSTRETL1("QTY_EOW") & "") + Val(cells(row, col).Value)
                            Case "TY UNITS SLS"
                                rowRSTRETL1("QTY_SOLD") = Val(rowRSTRETL1("QTY_SOLD") & "") + Val(cells(row, col).Value)
                            Case "TY RTL SLS"
                                rowRSTRETL1("AMT_SOLD") = Val(rowRSTRETL1("AMT_SOLD") & "") + Val(cells(row, col).Value)
                        End Select
                    Next
                End If
            End If

        Next

        If dst.Tables("RSTXLSQE").Rows.Count > 0 Then
            Dim sqlEDTTRPM1_check As String = "Select * FROM EDTTRPM1 WHERE EDI_TP_QUAL = :PARM1 AND EDI_TP_ID = :PARM2 AND EDI_DOC_NO = :PARM3"
            Dim rowEDTTRPM1 As DataRow = ASCDATA1.GetDataRow(sqlEDTTRPM1_check, "VVV", New Object() {"ZZ", CUST_CODE, "852"})

            If IsNothing(rowEDTTRPM1) Then
                Dim sqlEDTTRPM1_ins As String = $"INSERT INTO EDTTRPM1 (EDI_TP_QUAL, EDI_TP_ID, EDI_DOC_NO, CUST_CODE, EDI_STATUS) VALUES ('ZZ', '{CUST_CODE}', '852', '{CUST_CODE}', 'P')"
                ASCDATA1.ExecuteSQL(sqlEDTTRPM1_ins)
            End If
            Update_Record_TDA("EDTUPCX1")
            Using frmmsg As New ASFMSGBF
                frmmsg.Show_grd(dst.Tables("RSTXLSQE"), Me, "Invalid UPC Codes")
            End Using
        Else

            BeginTrans()

            ASCMAIN1.sql = $"Update RSTSLPI0 Set PROCESS_IND = '1', LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}' where DOC_CTL_NO = '{DOC_CTL_NO}'"
            ASCDATA1.ExecuteSQL()

            Update_Record_TDA("RSTRETL1")

            Update_Record_TDA("EDT852T1")

            For Each row As DataRow In dst.Tables("EDT852T1").Select()
                Dim EDI_DOC_SEQ_NO As String = row.Item("EDI_DOC_SEQ_NO")
                Update_RSTRETLx(EDI_DOC_SEQ_NO)
            Next

            CommitTrans("Import Complete")

        End If


        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")


    End Sub

    Private Function GetCustomerCode(fileName As String) As String
        ' Remove extension if present
        Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(fileName)

        ' Split at the first underscore
        Dim parts() As String = baseName.Split("_"c)

        If parts.Length > 0 Then
            ' Extract only letters from the first segment
            Return Regex.Match(parts(0), "^[A-Za-z]+").Value.ToUpper()
        End If

        Return String.Empty
    End Function

    Sub Process_852(DOC_CTL_NO As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Unzipping")

        Dim rowRSTSLPI0 As DataRow = dst.Tables("RSTSLPI0").Rows.Find(DOC_CTL_NO)
        Dim FILENAME As String = rowRSTSLPI0.Item("FILENAME")

        Dim DIRECTORYNAME As String = ASCMAIN1.Folders("Work") & "852\"

        If My.Computer.FileSystem.DirectoryExists(DIRECTORYNAME) Then
            My.Computer.FileSystem.DeleteDirectory(DIRECTORYNAME, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If

        Dim RS_PARM_FOLDER_SELLTHRU_DATA As String = ROWs("RSTPARM1").Item("RS_PARM_FOLDER_SELLTHRU_DATA")
        Dim RS_PARM_FOLDER_SELLTHRU_DATA_852 As String = System.IO.Path.Combine(RS_PARM_FOLDER_SELLTHRU_DATA, "852")

        If FILENAME.ToLower.EndsWith(".zip") Then
            ' UNZIP TO WORK
            Using ZipControl As New nsoftware.IPWorksZip.Zip
                ZipControl.RuntimeLicense = ASCMAIN1.nSoftwareKeys("nSoftwareZipkey")

                ZipControl.ArchiveFile = RS_PARM_FOLDER_SELLTHRU_DATA_852 & "\Inbox\" & FILENAME
                ZipControl.ExtractToPath = DIRECTORYNAME
                ZipControl.ExtractAll()
            End Using

        Else
            My.Computer.FileSystem.CreateDirectory(DIRECTORYNAME)

            System.IO.File.Copy(RS_PARM_FOLDER_SELLTHRU_DATA_852 & "\Inbox\" & FILENAME,
                                DIRECTORYNAME & FILENAME)
        End If

        dst.Tables("EDT852T3").Rows.Clear()
        dst.Tables("EDT852T1").Rows.Clear()
        dst.Tables("EDT852T2").Rows.Clear()
        dst.Tables("EDTJRNL1").Rows.Clear()

        Dim fIndex As Int32 = 0

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Translating X12 to XML")

        DATETIME_STAMP = Now + ASCMAIN1.NowTSD

        For Each FILE852 As String In My.Computer.FileSystem.GetFiles(DIRECTORYNAME)
            Debug.Print(FILE852)
            'If Not FILE852.EndsWith("2024-03-24 (52).txt") Then
            '    Continue For
            'End If
            fIndex += 1

            Dim txtReplace As String = ""

            Using sr As New System.IO.StreamReader(DIRECTORYNAME & FILENAME)
                Dim txt As String = sr.ReadToEnd
                If InStr(txt, "~GS*PD*6112390050*SLIPCRSPGXS") = 106 Then
                    txt = Replace(txt, vbLf, "~")
                    txtReplace = txt
                End If
            End Using

            If txtReplace <> "" Then

                Dim TargetFolder As String = DIRECTORYNAME & FILENAME
                ' BadTranslation
                System.IO.File.Move(DIRECTORYNAME & FILENAME,
                            RS_PARM_FOLDER_SELLTHRU_DATA_852 & "\BadTranslation\" & FILENAME & Format(Now, "yyyyMMddHHmmssff"))

                Using sw As New System.IO.StreamWriter(DIRECTORYNAME & FILENAME)
                    sw.Write(txtReplace)
                    Stop
                End Using
            End If

            Using EDIControl As New nsoftware.IPWorksEDI.X12reader()
                Using translator As New nsoftware.IPWorksEDI.X12translator
                    translator.InputFormat = nsoftware.IPWorksEDI.X12translatorInputFormats.xifX12
                    translator.OutputFormat = nsoftware.IPWorksEDI.X12translatorOutputFormats.xofXML

                    translator.InputFile = FILE852 ' "810.edi"
                    translator.OutputFile = FILE852 & ".XML" ' "X12_810.xml"
                    translator.UseSchemaName = True
                    translator.SchemaFormat = nsoftware.IPWorksEDI.X12translatorSchemaFormats.schemaJSON
                    'translator.LoadSchema("RSSBus_00401_852.json")

                    Dim translate_successful As Boolean = False

                    If FILE852.EndsWith(".out?") Then
                        Debug.Print(FILE852 & ":" & "skipping")
                        Stop
                    Else


                        Try
                            translator.Translate()
                            translate_successful = True
                        Catch ex As Exception
                            Debug.Print(FILE852 & ":" & ex.Message)
                            ' MsgBox(ex.Message, MsgBoxStyle.OkOnly, $"Cannot Translate {FILE852}")
                        End Try
                    End If


                    If translate_successful Then

                        Dim dicISA As New Dictionary(Of String, String)
                        Dim ISA As String = ""

                        Dim dicGS As New Dictionary(Of String, String)
                        Dim GS As String = ""

                        Dim dicXQ As New Dictionary(Of String, String)
                        Dim N9BT As String = ""
                        Dim N9DP As String = ""

                        Using textReader As New System.Xml.XmlTextReader(translator.OutputFile)

                            Dim rowEDT852T1 As DataRow = Nothing

                            Dim lastElementName As String = ""
                            Dim rIndex As Int32 = 0
                            Dim lastElementNameN9 As String = ""
                            Dim lastElementNameN1 As String = ""

                            Dim EDI_DOC_SEQ_NO As String = ""
                            Dim EDI_ITEM_ID_TYPE As String = ""
                            Dim EDI_PRICE_TYPE As String = ""
                            Dim rowEDT852T2 As DataRow = Nothing
                            Dim rowEDT852T3 As DataRow = Nothing
                            Dim EDI_ITEM_CODE As String = ""
                            Dim EDI_TRAN_TYPE As String = ""

                            Dim EDI_ZA_NO As Int32 = 0
                            Dim EDI_SDQ_SEQ As Int32 = 0

                            While textReader.Read()
                                rIndex += 1
                                If rIndex Mod 1000 = 1 Then
                                    ASCMAIN1.Progress("-", fIndex & ":" & rIndex)
                                End If

                                Select Case textReader.NodeType
                                    Case Xml.XmlNodeType.Element
                                        lastElementName = textReader.Name
                        'Debug.WriteLine(lastElementName)
                                    Case Xml.XmlNodeType.Text
                                    ' Debug.WriteLine(lastElementName & ": " & textReader.Value)
                                    Case Xml.XmlNodeType.EndElement
                                        'Debug.WriteLine("End:" & textReader.Name)
                                        If textReader.Name = "TransactionSet" Then ' "FunctionalGroup" Then
                                            ' Debug.Print(textReader.Name & ":" & CStr(rIndex))
                                            'Stop
                                            rowEDT852T1 = Write_EDT(EDI_DOC_SEQ_NO, FILE852, dicISA, dicGS, dicXQ, ISA, GS, N9BT, N9DP)
                                            dicXQ.Clear()
                                            N9BT = ""
                                            N9DP = ""

                                            EDI_ITEM_ID_TYPE = ""
                                            EDI_PRICE_TYPE = ""
                                            rowEDT852T2 = Nothing
                                            EDI_ITEM_CODE = ""
                                            rowEDT852T2 = Nothing
                                            EDI_DOC_SEQ_NO = ""
                                            EDI_TRAN_TYPE = ""
                                        End If

                                    Case Xml.XmlNodeType.Whitespace

                                    Case Else
                                        Stop
                                End Select

                                If lastElementName.StartsWith("ISA") Then
                                    Dim ISAXX As String = lastElementName ' textReader.Value
                                    Dim XV As String = Trim(textReader.Value)
                                    If XV <> "" And XV <> vbCrLf Then
                                        Dim XX As String = Mid(ISAXX, 4, 2)
                                        dicISA.Add(XX, XV)
                                        ISA &= "*" & ISAXX & "*" & XV
                                        'ISA &= "*" & XV
                                    End If
                                    'If ISAXX = "ISA13" Then
                                    '    ASCMAIN1.Progress("-", XV)
                                    'End If

                                End If

                                If lastElementName.StartsWith("GS") Then
                                    Dim GSXX As String = lastElementName ' textReader.Value
                                    Dim XV As String = Trim(textReader.Value)
                                    If XV <> "" And XV <> vbCrLf Then
                                        Dim XX As String = Mid(GSXX, 3, 2)
                                        dicGS.Add(XX, XV)
                                        GS &= "*" & GSXX & "*" & XV
                                        'GS &= "*" & XV
                                    End If
                                End If

                                If lastElementName.StartsWith("XQ") Then
                                    Dim XQXX As String = textReader.Value
                                    If XQXX <> "" And XQXX <> vbCrLf Then
                                        Dim XX As String = Mid(lastElementName, 3, 2)
                                        If XX = "02" Or XX = "03" Then
                                            dicXQ.Add(XX, Mid(XQXX, 5, 2) & "/" & Mid(XQXX, 7, 2) & "/" & Mid(XQXX, 1, 4))
                                        End If
                                    End If
                                End If

                                If lastElementName.StartsWith("N9") Then
                                    Dim N9XX As String = Trim(textReader.Value)
                                    If N9XX <> "" And N9XX <> vbCrLf Then
                                        If lastElementNameN9 = "BT" Then
                                            N9BT = N9XX
                                        ElseIf lastElementNameN9 = "DP" Then
                                            N9DP = N9XX
                                        End If
                                        If lastElementName = "N901" Or lastElementName = "N902" Then
                                            lastElementNameN9 = N9XX
                                        End If
                                    End If
                                End If


                                If lastElementName.StartsWith("N1") Then
                                    Dim N1XX As String = Trim(textReader.Value)
                                    If N1XX <> "" And N1XX <> vbCrLf Then
                                        If lastElementName = "N104" Then
                                            lastElementNameN1 = N1XX
                                        End If
                                    End If
                                End If

                                If lastElementName.StartsWith("LIN") Then
                                    Dim LINXX As String = Trim(textReader.Value)
                                    If LINXX <> "" And LINXX <> vbCrLf Then
                                        If lastElementName = "LIN01" Then
                                            EDI_ZA_NO = 0
                                            If EDI_DOC_SEQ_NO = "" Then
                                                EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDT852T1.EDI_DOC_SEQ_NO")
                                            End If
                                            rowEDT852T2 = dst.Tables("EDT852T2").NewRow
                                            rowEDT852T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                            rowEDT852T2.Item("EDI_LINE_NO") = Val(LINXX)
                                            dst.Tables("EDT852T2").Rows.Add(rowEDT852T2)
                                        ElseIf lastElementName = "LIN02" Or lastElementName = "LIN04" Or lastElementName = "LIN06" Or lastElementName = "LIN08" Then
                                            EDI_ITEM_ID_TYPE = LINXX
                                        ElseIf lastElementName = "LIN03" Or lastElementName = "LIN05" Or lastElementName = "LIN07" Or lastElementName = "LIN09" Then

                                            If rowEDT852T2 Is Nothing Then ' HAPPENS WHEN WE ENCOUNTER <LIN01/> WITH NO LINE NO - ASSUMING THAT THERE IS ONLY 1
                                                If EDI_DOC_SEQ_NO = "" Then
                                                    EDI_DOC_SEQ_NO = ASCMAIN1.Next_Control_No("EDT852T1.EDI_DOC_SEQ_NO")
                                                End If
                                                rowEDT852T2 = dst.Tables("EDT852T2").NewRow
                                                rowEDT852T2.Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO
                                                rowEDT852T2.Item("EDI_LINE_NO") = 1 ' Val(LINXX)
                                                dst.Tables("EDT852T2").Rows.Add(rowEDT852T2)
                                            End If

                                            EDI_ITEM_CODE = LINXX
                                            If EDI_ITEM_ID_TYPE = "UP" Then
                                                rowEDT852T2.Item("EDI_ITEM_UP") = LINXX
                                            ElseIf EDI_ITEM_ID_TYPE = "EN" Then
                                                rowEDT852T2.Item("EDI_ITEM_EN") = LINXX
                                            ElseIf EDI_ITEM_ID_TYPE = "UK" Then
                                                rowEDT852T2.Item("EDI_ITEM_GTIN") = LINXX
                                            End If
                                        End If
                                    End If
                                End If

                                If lastElementName.StartsWith("CTP") Then
                                    Dim CTPXX As String = Trim(textReader.Value)
                                    If CTPXX <> "" And CTPXX <> vbCrLf Then
                                        If lastElementName = "CTP02" Then
                                            EDI_PRICE_TYPE = CTPXX
                                        ElseIf lastElementName = "CTP03" Then
                                            If EDI_PRICE_TYPE = "RTL" Then
                                                rowEDT852T2.Item("EDI_PRICE_RTL") = Val(CTPXX)
                                            End If
                                        End If
                                    End If
                                End If

                                If lastElementName.StartsWith("ZA") Then
                                    Dim ZAXX As String = Trim(textReader.Value)

                                    If ZAXX <> "" And ZAXX <> vbCrLf Then
                                        If lastElementName = "ZA01" Then
                                            EDI_TRAN_TYPE = ZAXX

                                            EDI_ZA_NO += 1
                                            EDI_SDQ_SEQ = 0

                                        End If
                                        If lastElementName = "ZA02" Then
                                            If lastElementNameN1 <> "" Then
                                                EDI_SDQ_SEQ += 1
                                                rowEDT852T3 = dst.Tables("EDT852T3").NewRow
                                                rowEDT852T3.Item("EDI_DOC_SEQ_NO") = rowEDT852T2.Item("EDI_DOC_SEQ_NO")
                                                rowEDT852T3.Item("EDI_LINE_NO") = rowEDT852T2.Item("EDI_LINE_NO")
                                                rowEDT852T3.Item("EDI_ZA_NO") = EDI_ZA_NO
                                                rowEDT852T3.Item("EDI_SDQ_SEQ") = EDI_SDQ_SEQ
                                                rowEDT852T3.Item("EDI_ZA_TRAN_TYPE") = EDI_TRAN_TYPE

                                                rowEDT852T3.Item($"EDI_SDQ_STORE_01") = lastElementNameN1
                                                rowEDT852T3.Item($"EDI_SDQ_QTY_AMT_01") = Val(ZAXX)

                                                dst.Tables("EDT852T3").Rows.Add(rowEDT852T3)
                                            End If
                                        End If
                                    End If
                                End If

                                If lastElementName.StartsWith("SDQ") Then
                                    Dim SDQXX As String = Trim(textReader.Value)
                                    If SDQXX <> "" And SDQXX <> vbCrLf Then

                                        If lastElementName = "SDQ" Then
                                            ' BLOCK HEADER 
                                        ElseIf lastElementName = "SDQ01" Then

                                            EDI_SDQ_SEQ += 1
                                            rowEDT852T3 = dst.Tables("EDT852T3").NewRow
                                            rowEDT852T3.Item("EDI_DOC_SEQ_NO") = rowEDT852T2.Item("EDI_DOC_SEQ_NO")
                                            rowEDT852T3.Item("EDI_LINE_NO") = rowEDT852T2.Item("EDI_LINE_NO")
                                            rowEDT852T3.Item("EDI_ZA_NO") = EDI_ZA_NO
                                            rowEDT852T3.Item("EDI_SDQ_SEQ") = EDI_SDQ_SEQ
                                            rowEDT852T3.Item("EDI_ZA_TRAN_TYPE") = EDI_TRAN_TYPE

                                            'EDI_RETAIL_PRICE               Number(13, 2)
                                            'UNIT_BASIS_CODE                VARCHAR2(2)

                                            dst.Tables("EDT852T3").Rows.Add(rowEDT852T3)

                                        ElseIf lastElementName = "SDQ02" Then
                                        Else
                                            Dim XX As String = Mid(lastElementName, 4)
                                            If Val(XX) Mod 2 = 1 Then ' SDQ03,SDQ05,SDQ07
                                                Dim YY As String = Format((Val(XX) - 1) / 2, "00")
                                                rowEDT852T3.Item($"EDI_SDQ_STORE_{YY}") = SDQXX

                                            ElseIf Val(XX) Mod 2 = 0 Then ' SDQ04,SDQ06,SDQ08 Then
                                                Dim YY As String = Format((Val(XX) - 2) / 2, "00")
                                                rowEDT852T3.Item($"EDI_SDQ_QTY_AMT_{YY}") = Val(SDQXX)

                                            End If
                                        End If
                                    End If
                                End If



                            End While


                            rowEDT852T1.Item("LAST_DATE") = Now + ASCMAIN1.NowTSD
                            rowEDT852T1.Item("LAST_OPER") = ASCMAIN1.USER_ID

                            'Using reader As XmlReader = XmlReader.Create(translator.OutputFile)
                            '    ' ... code to read XML ...
                            '    Dim I As Integer = reader.AttributeCount
                            '    Stop

                            '    While reader.Read()
                            '        If reader.NodeType = XmlNodeType.Element Then
                            '            'Debug.Print(reader.Name)
                            '            Console.WriteLine($"{reader.Name}:  {reader.Value}")
                            '            'If reader.Name = "Meta" Or reader.Name = "FunctionalGroup" Or reader.Name = "TransactionSet" Or reader.Name = "TransactionSet" Then
                            '            '    Console.WriteLine($"{reader.Name}")
                            '            'Else
                            '            '    Console.WriteLine($"{reader.Name}: {reader.Value}")
                            '            '    ' Console.WriteLine($"{reader.Name}: {reader.ReadElementContentAsString()}")
                            '            'End If


                            '            Select Case reader.Name
                            '                Case "ItemName"
                            '                    Console.WriteLine("Item Name: " & reader.ReadElementContentAsString())
                            '                Case "Price"
                            '                    Console.WriteLine("Price: " & reader.ReadElementContentAsString())
                            '                    'Case "ISA01"
                            '                    '    Console.WriteLine("ISA01: " & reader.ReadElementContentAsString())
                            '            End Select
                            '        End If
                            '    End While
                            'End Using


                        End Using
                    End If

                End Using
            End Using
        Next
        'Stop

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Updating Database")

        BeginTrans()

        For Each TABLE_NAME As String In New String() {"EDT852T3", "EDT852T1", "EDT852T2", "EDTJRNL1"}
            If TABLE_NAME = "EDT852T3 - must limit array size" Then
                Update_BAs(TABLE_NAME)
            Else
                Update_Record_TDA(TABLE_NAME)
            End If
        Next

        'Update_Record_TDA("EDT852T3")
        'Update_Record_TDA("EDT852T1")
        'Update_Record_TDA("EDT852T2")
        'Update_Record_TDA("EDTJRNL1")

        ASCMAIN1.sql = $"Update RSTSLPI0 Set PROCESS_IND = '1', LAST_DATE = SYSDATE, LAST_OPER = '{ASCMAIN1.USER_ID}' where DOC_CTL_NO = '{DOC_CTL_NO}'"
        ASCDATA1.ExecuteSQL()

        CommitTrans("Translation Complete")

        'If FILENAME.ToLower.EndsWith(".zip") Then
        System.IO.File.Move(RS_PARM_FOLDER_SELLTHRU_DATA_852 & "\Inbox\" & FILENAME,
                            RS_PARM_FOLDER_SELLTHRU_DATA_852 & "\Archive\" & FILENAME)
        'End If

        For Each TABLE_NAME As String In New String() {"EDT852T3", "EDT852T1", "EDT852T2", "EDTJRNL1"}
            dst.Tables(TABLE_NAME).Rows.Clear()
        Next

        Prepare_EDI_Data()

        ' Prepare_852_Queue()

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

    End Sub

    Function Write_EDT(EDI_DOC_SEQ_NO As String, FILE852 As String, dicISA As Dictionary(Of String, String), dicGS As Dictionary(Of String, String), dicXQ As Dictionary(Of String, String), ISA As String, GS As String, N9BT As String, N9DP As String) As DataRow

        Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILE852)

        Dim rowEDT852T1 As DataRow = dst.Tables("EDT852T1").NewRow
        Dim rowEDTJRNL1 As DataRow = dst.Tables("EDTJRNL1").NewRow

        If Not dicXQ.ContainsKey("03") Then
            dicXQ.Add("03", dicXQ("02"))
        End If
        Dim EDI_TO_DATE As Date = CDate(dicXQ("03"))

        ASCMAIN1.sql = "Select * from GLTPARM3 where WEEK_END_DATE = :PARM1"
        Dim rowGLTPARM3 As DataRow = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "D", New Object() {EDI_TO_DATE})
        If rowGLTPARM3 Is Nothing Then
            rowGLTPARM3 = ASCDATA1.GetDataRow(ASCMAIN1.sql, False, "D", New Object() {EDI_TO_DATE.AddDays(-1)})
        End If
        If rowGLTPARM3 Is Nothing Then Stop
        Dim OPS_YYYYWW As String = rowGLTPARM3.Item("YYYYWW")
        Dim OPS_YYYYPP As String = rowGLTPARM3.Item("YYYYPP")

        Dim EDI_TP_QUAL As String = ""
        Dim EDI_TP_ID As String = ""

        Dim rowEDTTRPM1s() As DataRow = dst.Tables("EDTTRPM1").Select($"EDI_TP_QUAL = '{EDI_TP_QUAL}' and EDI_TP_ID = '{EDI_TP_ID}'")
        Dim CUST_CODE As String = ""
        If rowEDTTRPM1s.Length = 1 Then
            CUST_CODE = rowEDTTRPM1s(0).Item("CUST_CODE")
        End If

        With rowEDT852T1
            .Item("EDI_DOC_SEQ_NO") = EDI_DOC_SEQ_NO ' ASCMAIN1.Next_Control_No("EDT852T1.EDI_DOC_SEQ_NO")
            .Item("GEN_DOC_NO") = ""
            .Item("EDI_ISA_NO") = dicISA("13")
            .Item("EDI_TP_QUAL") = dicISA("05")
            .Item("EDI_TP_ID") = dicISA("06")
            .Item("EDI_OUR_QUAL") = dicISA("07")
            .Item("EDI_OUR_ID") = dicISA("08")
            .Item("EDI_FROM_DATE") = CDate(dicXQ("02"))
            .Item("EDI_TO_DATE") = CDate(dicXQ("03"))
            .Item("EDI_CUST_BATCH_NO") = N9BT
            .Item("EDI_DEPT_NO") = N9DP
            .Item("EDI_STORE_NO") = ""
            .Item("EDI_STATUS") = "0"
            .Item("OPS_YYYYPP") = OPS_YYYYPP
            .Item("OPS_YYYYWW") = OPS_YYYYWW
            .Item("CUST_CODE") = CUST_CODE
            .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD '  DATETIME_STAMP
            .Item("INIT_OPER") = ASCMAIN1.USER_ID
            .Item("LAST_DATE") = DBNull.Value
            .Item("LAST_OPER") = DBNull.Value

            '.Item("COLLECTION_CODE") = ""
            '.Item("DATA_LEVEL") = ""
            '.Item("BRAND_CODE") = ""
            '.Item("ERROR_ITEMS") = ""
            '.Item("ERROR_STORES") = ""
            '.Item("COMPANY_CODE") = ""

            .Item("EDI_SOURCE") = "852"
        End With
        dst.Tables("EDT852T1").Rows.Add(rowEDT852T1)

        With rowEDTJRNL1
            .Item("EDI_JRNL_NO") = ASCMAIN1.Next_Control_No("EDTJRNL1.EDI_JOURNAL_NO")
            .Item("EDI_JRNL_DATE") = DATETIME_STAMP

            .Item("EDI_FILENAME") = fi.Name

            Dim EDI_ISA_DATE As String = dicISA("09")
            EDI_ISA_DATE = Mid(EDI_ISA_DATE, 3, 2) & "/" & Mid(EDI_ISA_DATE, 5, 2) & "/" & Mid(EDI_ISA_DATE, 1, 2)

            .Item("EDI_ISA_NO") = rowEDT852T1.Item("EDI_ISA_NO")
            .Item("EDI_ISA_DATE") = CDate(EDI_ISA_DATE)

            .Item("EDI_TP_QUAL") = dicISA("05")
            .Item("EDI_TP_ID") = dicISA("06")
            .Item("EDI_OUR_QUAL") = dicISA("07")
            .Item("EDI_OUR_ID") = dicISA("08")

            .Item("EDI_ISA_RECORD") = ISA
            .Item("EDI_FOLDERNAME") = System.IO.Path.Combine(ROWs("RSTPARM1").Item("RS_PARM_FOLDER_SELLTHRU_DATA"), "852") & "\Archive\"

            .Item("EDI_DATETIME") = fi.CreationTime
            .Item("EDI_FILESIZE") = fi.Length

            .Item("EDI_GS_RECORD") = GS
            .Item("EDI_GS_NO") = dicGS("06")
            .Item("EDI_VERSION") = dicGS("08")

        End With
        dst.Tables("EDTJRNL1").Rows.Add(rowEDTJRNL1)

        Dim GEN_DOC_NO As String = rowEDTJRNL1.Item("EDI_JRNL_NO")
        rowEDT852T1.Item("GEN_DOC_NO") = GEN_DOC_NO
        Dim FILENAME_ARCHIVE As String = System.IO.Path.Combine(ROWs("RSTPARM1").Item("RS_PARM_FOLDER_SELLTHRU_DATA"), "852") & "\Archive\" & GEN_DOC_NO
        System.IO.File.Copy(fi.FullName, FILENAME_ARCHIVE)

        Return rowEDT852T1
    End Function

    Private Sub btn852Inbox_Click(sender As Object, e As EventArgs) Handles btn852Inbox.Click

        Dim DOC_SOURCE As String = optMode.Value
        Dim RS_PARM_FOLDER_SELLTHRU_DATA As String = System.IO.Path.Combine(ROWs("RSTPARM1").Item("RS_PARM_FOLDER_SELLTHRU_DATA"), DOC_SOURCE)

        dst.Tables("RSTSLPI0").Rows.Clear()

        For Each FILENAME As String In System.IO.Directory.GetFiles(RS_PARM_FOLDER_SELLTHRU_DATA & "\Inbox")

            Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)

            Dim rowRSTSLPI0 As DataRow = dst.Tables("RSTSLPI0").NewRow
            With rowRSTSLPI0
                .Item("DOC_CTL_NO") = ASCMAIN1.Next_Control_No("RSTSLPI0.DOC_CTL_NO")
                .Item("DOC_SOURCE") = DOC_SOURCE
                .Item("FILENAME") = fi.Name ' FILENAME
                .Item("PROCESS_IND") = "0"
                .Item("INIT_DATE") = Now + ASCMAIN1.NowTSD
                .Item("INIT_OPER") = ASCMAIN1.USER_ID
            End With
            dst.Tables("RSTSLPI0").Rows.Add(rowRSTSLPI0)


            System.IO.File.Move(FILENAME,
                                RS_PARM_FOLDER_SELLTHRU_DATA & "\Archive\" & fi.Name)
        Next

        If dst.Tables("RSTSLPI0").Rows.Count > 0 Then
            Update_Record_TDA("RSTSLPI0")
        End If


        Setup_grdRSTSLPI0()

    End Sub

    Private Sub btn852sftp_Click(sender As Object, e As EventArgs) Handles btn852sftp.Click

        Dim DOC_SOURCE As String = optMode.Value
        Dim RS_PARM_FOLDER_SELLTHRU_DATA As String = System.IO.Path.Combine(ROWs("RSTPARM1").Item("RS_PARM_FOLDER_SELLTHRU_DATA"), DOC_SOURCE)

        Dim rowTATSSHK1 As DataRow = LookUp("TATSSHK1", "CRISP")
        Dim SSH_APP_FOLDER_GET As String = rowTATSSHK1.Item("SSH_APP_FOLDER_GET") & ""
        Dim fCount As Integer = 0

        ' we are assuming that the files that Crisp is sending are .out files

        For Each FILENAME As String In System.IO.Directory.GetFiles(SSH_APP_FOLDER_GET)

            Dim fi As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(FILENAME)
            fCount += 1

            System.IO.File.Copy(FILENAME,
                    RS_PARM_FOLDER_SELLTHRU_DATA & "\Inbox\" & fi.Name)

            System.IO.File.Move(FILENAME,
                    SSH_APP_FOLDER_GET & "\Archive\" & fi.Name)
        Next

        MsgBox($"{fCount} Files Successfully Transferred to Inbox, and Archived", MsgBoxStyle.OkOnly, "Success")

        Setup_grdRSTSLPI0()

    End Sub
    Private Sub Set_Selected_852_Status(ByVal newStatus As String)
        newStatus = (newStatus & "").Trim().ToUpper()

        If newStatus <> "A" AndAlso newStatus <> "0" Then Exit Sub

        If newStatus = "A" AndAlso opt852Data.Value & "" <> "0" Then Exit Sub
        If newStatus = "0" AndAlso opt852Data.Value & "" <> "A" Then Exit Sub

        If grdEDT852T1.Rows.Count = 0 Then Exit Sub

        Dim expectedOldStatus As String = If(newStatus = "A", "0", "A")
        Dim actionText As String = If(newStatus = "A", "Archive", "Unarchive")

        Dim seqs As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each grow As Infragistics.Win.UltraWinGrid.UltraGridRow In grdEDT852T1.Rows
            If grow Is Nothing OrElse Not grow.IsDataRow OrElse grow.IsFilteredOut Then Continue For
            If Not grow.Cells.Exists("SELECTED") OrElse Not grow.Cells.Exists("EDI_DOC_SEQ_NO") Then Continue For
            If grow.Cells("SELECTED").Value & "" = "1" Then
                Dim seq As String = grow.Cells("EDI_DOC_SEQ_NO").Value & ""
                If seq <> "" Then seqs.Add(seq)
            End If
        Next

        If seqs.Count = 0 Then
            MsgBox("No Documents Selected", MsgBoxStyle.OkOnly, actionText)
            Exit Sub
        End If

        If MsgBox(actionText & " " & seqs.Count & " document(s)?", MsgBoxStyle.YesNo, "Verification") = MsgBoxResult.No Then Exit Sub

        BeginTrans()

        Try
            For Each seq As String In seqs
                ASCDATA1.ExecuteSQL(
                    "Update EDT852T1 Set EDI_STATUS = '" & newStatus & "' " &
                    "Where COMPANY_CODE = '" & ASCMAIN1.DBS_COMPANY & "' " &
                    "  And EDI_DOC_SEQ_NO = '" & seq.Replace("'", "''") & "' " &
                    "  And EDI_STATUS = '" & expectedOldStatus & "'")

                Dim EVENT_TYPE As String = If(newStatus = "A", "ARCHIVE", "UNARCHIVE")
                Dim EVENT_DESC As String = If(newStatus = "A",
                               "852 Document Archived (Status 0 -> A)",
                               "852 Document Restored (Status A -> 0)")
                Dim EVENT_KEY As String = newStatus
                ASCMAIN1.Record_Event("EDT852T1", seq, "", DATETIME_STAMP, ASCMAIN1.USER_ID, EVENT_TYPE, EVENT_DESC, EVENT_KEY)
            Next

            CommitTrans()

            For Each row As DataRow In dst.Tables("EDT852T1").Select(
                "EDI_DOC_SEQ_NO IN ('" & String.Join("','", seqs).Replace("'", "''") & "')", "",
                DataViewRowState.CurrentRows)

                row("EDI_STATUS") = newStatus
                If dst.Tables("EDT852T1").Columns.Contains("SELECTED") Then row("SELECTED") = "0"
            Next

            Prepare_852_Queue()

        Catch ex As Exception
            Rollback(ex.Message)

        End Try
    End Sub

End Class