Public Class POROPEN3

    Dim wb As SpreadsheetGear.IWorkbook
    Dim ws As SpreadsheetGear.IWorksheet = Nothing
    Dim range As SpreadsheetGear.IRange = Nothing

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Get_PARM("POTPARM1")

        Range_Events(grpDateRequired)

    End Sub

    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()

        Call ASCMAIN1.Progress("Building Work File")

        Dim sql As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""
         
        If Not Absx1.chkFor("CHKDATEREQ_F").Checked Then
            Dim z As String = Format(Absx1.dteFor("DATEREQ_F").Value, "dd-MMM-yyyy")
            sql_filter &= " and POTORDR2.PO_DATE_REQUIRED >= '" & z & "'"
            Page0.Add("Date Required >= " & z)
        End If
        If Not Absx1.chkFor("CHKDATEREQ_L").Checked Then
            Dim z As String = Format(Absx1.dteFor("DATEREQ_L").Value, "dd-MMM-yyyy")
            sql_filter &= " and POTORDR2.PO_DATE_REQUIRED <= '" & z & "'"
            Page0.Add("Date Required <= " & z)
        End If

        sql_filter &= " and APTVEND1.VEND_CODE = POTORDR1.VEND_CODE"
        sql_filter &= " and POTORDR2.PO_QTY_OPN <> 0"

        ' Extracts from Data Sources
         
        MyBase.Get_SQL("*")
        ASCMAIN1.Progress("Building Tiers")

 

        sql = "Select " & sql_SELECT_cols & ASTSRPT1_sum_columns & vbCrLf
        sql &= " from POTORDR2, APTVEND1 " & sql_TABLE_NAMEs & vbCrLf
        sql &= ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCrLf
        'sql &= " group by " & sql_GROUP_BY_cols
        '  ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        sql = "Select " & sql_SELECT_cols & vbCrLf _
            & ",POTORDR2.WHSE_CODE,POTORDR2.PO_DATE_REQUIRED,POTORDR2.PO_ORDER_NO,POTORDR2.PO_ORDER_LNO" & vbCrLf _
            & ",POTORDR1.VEND_CODE,POTORDR1.PO_REFERENCE,POTORDR2.ITEM_CODE,ICTITEM1.ITEM_DESC,POTORDR2.BM_ISSUE_NO,POTORDR2.BM_ISSUE_SEL" & vbCrLf _
            & ",POTORDR2.PO_QTY_ORD,POTORDR2.PO_QTY_REC,POTORDR2.PO_QTY_OPN,APTVEND1.VEND_NAME,POTORDR1.PO_DATE_ORDERED,POTORDR1.PO_DATE_CANCEL" & vbCrLf _
            & " from POTORDR2, APTVEND1 " & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter)

        '& vbCrLf _
        '           & " group by " & sql_GROUP_BY_cols & vbCrLf _
        '           & IIf(sql_GROUP_BY_cols = "", "", ",") & "POTORDR2.PO_ORDER_NO, POTORDR2.PO_ORDER_LNO"

        Dim sql_Cols As String = ""

        ASCMAIN1.sql = "Insert into " & ASTSRPT1 & vbCrLf _
            & "(" & G1thru9 & COLUMN_NAMEs_appended _
            & sql_Cols & ")" & vbCrLf _
            & "(" & sql & ")"
        ASCDATA1.ExecuteSQL()

    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = ""
        If Not Absx1.chkFor("CHKDATEREQ_F").Checked _
        Or Not Absx1.chkFor("CHKDATEREQ_L").Checked Then
            SUBT = SUBT & "Showing Purchase Orders Required"
            If Not Absx1.chkFor("CHKDATEREQ_F").Checked Then
                SUBT = SUBT & " from " & Format(Absx1.dteFor("DATEREQ_F").Value, "MM/dd/yyyy")
            End If
            If Not Absx1.chkFor("CHKDATEREQ_L").Checked Then
                SUBT = SUBT & " thru " & Format(Absx1.dteFor("DATEREQ_L").Value, "MM/dd/yyyy")
            End If
        End If

        Generate_Report(RPT, , SUBT)

        Create_XLS()
    End Sub

    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        'If eItemKey = "Proceed" Then
        '    If Not Absx1.chkFor("CHKPO").Checked And Not Absx1.chkFor("CHKPP").Checked Then
        '        EMsg &= vbCr & "You must select at least 1: $Value Received and/or Units Received"
        '    End If
        'End If
    End Sub

    Sub Create_XLS()

        Dim SQLW As String = ""
        If ASCMAIN1.Running_in_VS Then
            Stop
            SQLW = $"VEND_CODE = 'HLPKLEAR' OR VEND_CODE = 'IPSA'"
        End If

        For Each rowVEND_CODE As DataRow In ASCDATA1.SelectDistinct _
            (dst.Tables("ASTSRPT1"), New String() {"VEND_CODE", "VEND_NAME"}).Select(SQLW)

            Dim VEND_CODE As String = rowVEND_CODE.Item("VEND_CODE")
            Dim VEND_NAME As String = rowVEND_CODE.Item("VEND_NAME")


            Dim sqlPINV As String = "Select ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO, SUM (ICTPINV2.PINV_QTY) PINV_QTY
 from ICTPINV2,ICTPINV1
where ICTPINV2.PINV_NO = ICTPINV1.PINV_NO and ICTPINV1.PINV_STATUS = 'O'
group by ICTPINV2.PO_ORDER_NO, ICTPINV2.PO_ORDER_LNO
"

            ASCMAIN1.sql = $"
Select ASTSRPT1.VEND_CODE, ASTSRPT1.PO_DATE_REQUIRED
, ASTSRPT1.PO_ORDER_NO, ASTSRPT1.PO_ORDER_LNO
, ASTSRPT1.ITEM_CODE, ASTSRPT1.ITEM_DESC
, ASTSRPT1.PO_QTY_ORD, ASTSRPT1.PO_QTY_REC, PINV.PINV_QTY
, NVL(ASTSRPT1.PO_QTY_ORD,0) - NVL(ASTSRPT1.PO_QTY_REC,0) - NVL(PINV.PINV_QTY,0) PINV_BAL
, POTORDR2.PO_DATE_ETD, POTORDR2.PO_DATE_ETD_NOTES 
From {ASTSRPT1} ASTSRPT1, POTORDR2, ({sqlPINV}) PINV
where POTORDR2.PO_ORDER_NO = ASTSRPT1.PO_ORDER_NO and POTORDR2.PO_ORDER_LNO = ASTSRPT1.PO_ORDER_LNO
and ASTSRPT1.VEND_CODE = '{VEND_CODE}' 
and PINV.PO_ORDER_NO (+) = POTORDR2.PO_ORDER_NO and PINV.PO_ORDER_LNO (+) = POTORDR2.PO_ORDER_LNO
order by ASTSRPT1.PO_ORDER_NO, ASTSRPT1.PO_ORDER_LNO"

            Dim DataTable As DataTable = ASCDATA1.GetDataTable

            Dim STA_REQ_NO As String = ASCMAIN1.Next_Control_No("POROPEN3.STA_REQ_NO")

            Dim XLS_FILENAME As String = ASCMAIN1.Folders("Work") & $"Vendor_{VEND_CODE}_{STA_REQ_NO}.xlsx"

            wb = SpreadsheetGear.Factory.GetWorkbook()
            ws = wb.Worksheets(0)
            ws.Name = "Open POs"

            Dim r0 As Int32 = 1 ' Row containing headings
            Dim rTotal As Int32 = DataTable.Select("").Length
            Dim cTotal As Int32 = DataTable.Columns.Count
            Dim r As Int32 = 0

            Dim c As Integer = -1

            c += 1 : Format_Column(r0, c, "Vendor", 12, "@", "VEND_CODE")
            c += 1 : Format_Column(r0, c, "Date Req", 12, "MM/DD/YY", "PO_DATE_REQUIRED")
            c += 1 : Format_Column(r0, c, "PO No", 10, "@", "PO_ORDER_NO")
            c += 1 : Format_Column(r0, c, "Ln", 7, "#0", "PO_ORDER_LNO")
            c += 1 : Format_Column(r0, c, "Item Code", 15, "@", "ITEM_CODE")
            c += 1 : Format_Column(r0, c, "Item Description", 30, "@", "ITEM_DESC")
            c += 1 : Format_Column(r0, c, "Qty Ord", 12, "#,##0", "PO_QTY_ORD", rTotal)
            c += 1 : Format_Column(r0, c, "Qty Rec", 12, "#,###", "PO_QTY_REC", rTotal)
            c += 1 : Format_Column(r0, c, "Qty Inv", 12, "#,##0", "PINV_QTY", rTotal)
            c += 1 : Format_Column(r0, c, "Qty Bal", 12, "#,##0", "PINV_BAL", rTotal)

            ws.Cells(r0, 0, r0, c).Interior.Color = SpreadsheetGear.Colors.LightBlue

            Dim PO_DATE_ETD_caption As String = "Prev Ship Date"
            Dim PO_DATE_ETD_NOTES_caption As String = "Prev Comment"
            If VEND_CODE = "IPSA" Then
                PO_DATE_ETD_caption = "Prev Bollore Date"
                PO_DATE_ETD_NOTES_caption = "Prev Comment"
            End If

            c += 1 : Format_Column(r0, c, PO_DATE_ETD_caption, 16, "MM/DD/YYYY", "PO_DATE_ETD")
            ws.Cells(r0 - 1, c).Value = "MM/DD/YYYY"
            c += 1 : Format_Column(r0, c, PO_DATE_ETD_NOTES_caption, 25, "@", "PO_DATE_ETD_NOTES")

            ws.Cells(r0, c - 1, r0, c).Interior.Color = SpreadsheetGear.Colors.LightPink

            Dim PO_DATE_ETD_caption_new As String = "New Ship Date"
            Dim PO_DATE_ETD_NOTES_caption_new As String = "New Comment"
            If VEND_CODE = "IPSA" Then
                PO_DATE_ETD_caption_new = "New Bollore Date"
                PO_DATE_ETD_NOTES_caption_new = "New Comment"
            End If
            c += 1 : Format_Column(r0, c, PO_DATE_ETD_caption_new, 16, "MM/DD/YYYY", "PO_DATE_ETD")
            ws.Cells(r0 - 1, c).Value = "MM/DD/YYYY"
            ws.Cells(r0, c).EntireColumn.Locked = False
            c += 1 : Format_Column(r0, c, PO_DATE_ETD_NOTES_caption_new, 25, "@", "PO_DATE_ETD_NOTES")
            ws.Cells(r0, c).EntireColumn.Locked = False

            ws.Cells(r0, c - 1, r0, c).Interior.Color = SpreadsheetGear.Colors.LightGoldenrodYellow
            ws.Cells(r0 + 1, c - 1, r0 + rTotal, c).Interior.Color = SpreadsheetGear.Colors.Yellow

            ws.Cells(r0, 0, r0, c).AutoFilter()

            ws.Cells(r0 + 1, 0).Activate()
            ws.WindowInfo.FreezePanes = True

            ws.Range(r0 + 1, 0).CopyFromDataTable(DataTable, SpreadsheetGear.Data.SetDataFlags.NoColumnHeaders)


            ws.Protect("ABS")

            wb.SaveAs(XLS_FILENAME, SpreadsheetGear.FileFormat.OpenXMLWorkbook)

            email_to_Self(XLS_FILENAME, VEND_CODE, VEND_NAME, STA_REQ_NO)
            ' Show_Document(XLS_FILENAME)

        Next
    End Sub

    Sub Format_Column(r0 As Integer, c As Integer, Caption As String, Width As Integer, Format As String, COLUMN_NAME As String, Optional rTotal As Int32 = -1)

        With ws.Cells(r0, c)
            .EntireColumn.NumberFormat = Format
            .EntireColumn.ColumnWidth = Width
            .EntireColumn.Locked = True
            If Format = "MM/DD/YY" Or Format = "MM/DD/YYYY" Then
                .EntireColumn.HorizontalAlignment = SpreadsheetGear.HAlign.Center
            End If
            .Value = Caption
            If rTotal <> -1 Then
                ws.Cells(r0 - 1, c).Formula = $"=subtotal(9,{Excel_Cell0(r0 + 1, c)}:{Excel_Cell0(r0 + rTotal, c)})"
            End If
        End With

    End Sub


    Sub email_to_Self(FILENAME As String, VEND_CODE As String, VEND_NAME As String, STA_REQ_NO As String)

        Me.Cursor = Cursors.WaitCursor
        ASCMAIN1.Progress("Now Preparing email")


        Dim EMAIL_ADDRESSs As New Dictionary(Of String, String)

        If ASCMAIN1.Running_in_VS Then
            EMAIL_ADDRESSs.Add("wjz@absolution.com", ASCMAIN1.USER_NAME)
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
        Else
            EMAIL_ADDRESSs.Add(ASCMAIN1.USER_EMAIL, ASCMAIN1.USER_NAME)
        End If

        Dim ATTACHMENTs As New Dictionary(Of String, String)
        ATTACHMENTs.Add($"Request for PO Status - {VEND_NAME}", FILENAME)

        Dim EMAIL_SUBJECT As String = $"Request for PO Status - {VEND_NAME}"


        Dim strHtml As String = ""

        Dim EMAILS As String = ""

        'ASCMAIN1.sql = $"Select TATCONT1.* from TATCONT1 where TATCONT1.CONTACT_TABLE = 'APTVEND1' and TATCONT1.CONTACT_KEY = '{VEND_CODE}'"
        'For Each rowTATCONT1 As DataRow In ASCDATA1.GetDataTable.Select("", "CONTACT_NO")
        '    Dim CONTACT_EMAIL As String = rowTATCONT1.Item("CONTACT_EMAIL")
        '    EMAILS &= ";" & rowTATCONT1.Item("CONTACT_EMAIL")
        'Next

        Dim rowAPTVEND1 As DataRow = LookUp("APTVEND1", VEND_CODE)
        Dim VEND_PURCH_EMAIL As String = rowAPTVEND1.Item("VEND_PURCH_EMAIL") & ""
        Dim VEND_ACCT_NO As String = rowAPTVEND1.Item("VEND_ACCT_NO") & ""

        If VEND_PURCH_EMAIL <> "" Then EMAILS &= ";" & VEND_PURCH_EMAIL


        Dim rowASTUSER1 As DataRow = LookUp("ASTUSER1", ASCMAIN1.USER_ID)

        Dim yellow As String = " style='background-color:yellow'"
        Dim blue As String = " style='color:blue'"
        Dim margin As String = " style='margin-left:20px'"

        strHtml = "<h1>Request for Status</h1>"
        strHtml &= $"<br/><div>Vendor Contacts who should receive this email<br/>{Replace(Mid(EMAILS, 2), ";", "<br/>")}</div><br/>"

        If Now.Hour < 12 Then
            strHtml &= $"<br/><div>Good Morning</div>"
        Else
            strHtml &= $"<br/><div>Good Afternoon</div>"
        End If

        strHtml &= $"<br/><div>Instructions:</div>"
        strHtml &= $"<ul>"
        strHtml &= $"<li>Please fill out Ship Date And Notes (if any) in the Yellow Columns of the Attachex XLS file.</strong>.</li>"
        strHtml &= $"<li>Please make sure your Dates are formatted as dates (MM/DD/YYYY) and do not put messages into Date Columns.</li>"
        ' If tbl.Rows.Count > 1 Then strHtml &= $"<li>We have multiple Accounts with you with Open Purchase Orders. POs for each Account are listed on separate XLS Sheets</li>"
        strHtml &= $"</ul>"

        If VEND_ACCT_NO <> "" Then
            strHtml &= $"<br/><div>Account No <strong{blue}>{VEND_ACCT_NO}</strong></div>"
        End If

        strHtml &= $"<br/><div>Thank You,</div><br/>" ' & TAC.POCMAIN1.GetUserSignature(rowASTUSER1)

        Dim SEND_NO As String = ASCMAIN1.TACMAIN1.Send_email _
                    (ASCMAIN1.ActiveForm, EMAIL_ADDRESSs, ATTACHMENTs,
                    EMAIL_SUBJECT, "STAREQ", True, False, STA_REQ_NO, "STA_REQ_NO", "Request for PO Status", strHtml)

        Me.Cursor = Cursors.Default
        ASCMAIN1.Progress("")

        'MsgBox("email Sent", MsgBoxStyle.OkOnly, "Verification")
    End Sub

End Class
