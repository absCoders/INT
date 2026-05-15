Public Class SOROPNO1

    Dim SOTORDRL As String
    Dim SOTORDR1 As String
    Dim SOTORDRG As String
    Dim SOTORDR0 As String

    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Range_Events(grpORDR_DATE_BOOKED)
        Range_Events(grpORDR_SHIP_DATE)
        Range_Events(grpORDR_CANCEL_DATE)
        Range_Events(grpORDR_DATE_REL)

        Get_PARM("SOTPARM1")
    End Sub


#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdSOTORDRC, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow Then
            '   e.Cancel = True
        Else
            'If grd.Selected.Rows.Count = 0 Then
            '    e.Cancel = True
            'End If

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

        End Select

        If grd.ActiveRow Is Nothing OrElse grd.ActiveRow.IsAddRow OrElse Not grd.ActiveRow.IsDataRow Then
            Exit Sub
        End If

        Select Case e.Tool.Key
            'Case "Sales Order Inquiry"
            '    Dim ORDR_NO As String = grd.ActiveRow.Cells("ORDR_NO").Value
            '    Dim rowSOTORDR1 As DataRow = LookUp("SOTORDR1", ORDR_NO)
            '    If rowSOTORDR1 IsNot Nothing Then
            '        Context_Launch("View", ORDR_NO, e.Tool.Key, "SOFORDRI")
            '    End If

        End Select
    End Sub

#End Region


    Overrides Sub Clear_Record()

    End Sub

    Protected Overrides Sub Build_Workfile()

        ASCMAIN1.Progress("Building Work File")

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        ' Extracts from Data Sources

        MyBase.Get_SQL("*")

        If chkFILTER_DETAILS.Checked Then
            Replace(sql_filter, " and SOTORDR1.BRAND_CODE ", " and ICTCOLL1.BRAND_CODE ")
            Replace(sql_filter, " and SOTORDR1.COLLECTION_CODE ", " and ICTITEM1.COLLECTION_CODE ")
        End If

        Dim INCLUDEX As String = ""
        If chkINCLUDEF.Checked Then INCLUDEX = INCLUDEX & ",'F'"
        If chkINCLUDEC.Checked Then INCLUDEX = INCLUDEX & ",'C'"

        If optIncl.Value = "R" Then
            sql_filter &= " and SOTORDR1.ORDR_STATUS IN ('P'" & INCLUDEX & ")"
        Else
            If optIncl.Value = "U" Then
                sql_filter &= " and SOTORDR1.ORDR_STATUS IN ('O'" & INCLUDEX & ")"
            Else
                sql_filter &= " and SOTORDR1.ORDR_STATUS in ('O','P'" & INCLUDEX & ")"
            End If

            If Not chkINCOPREL.Checked Then
                sql_filter &= " and SOTORDR1.ORDR_DATE_REL IS NULL"
            End If
        End If

        If optEDI.Value <> "A" Then
            If optEDI.Value = "E" Then
                sql_filter &= " and SOTORDR1.ORDR_SOURCE = 'E'"
            Else
                sql_filter &= " and SOTORDR1.ORDR_SOURCE <> 'E'"
            End If
        End If

        If chkSUPHOLDS.Checked Then
            sql_filter &= " and NVL(SOTORDR2.ORDR_RELEASE,'A') <> 'A'"
        End If

        ASCMAIN1.sql = "Select SOTORDR1.ORDR_NO, SOTORDR2.ORDR_LNO, SOTORDR2.ALLO_CTL_NO" & vbCrLf _
            & " from " & "SOTORDR1" & sql_TABLE_NAMEs & vbCrLf _
            & ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter & Get_Dates())
        SOTORDRL = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRL & " Add Primary Key (ORDR_NO, ORDR_LNO)")

        Dim GF As String = ""
        Select Case optBy.Value
            Case "I"
                GF = "ITEM_CODE"
            Case "C"
                GF = "CUST_CODE"
            Case "S"
                GF = "ORDR_SHIP_DATE"
            Case "X"
                GF = "ORDR_CANCEL_DATE"
            Case "T"
                GF = "SREP_CODE"
            Case "R"
                GF = "ORDR_DATE_REL"
            Case "B"
                GF = "INIT_DATE"
        End Select
        If GF.Contains("_DATE") Then GF = "TO_CHAR(SOTORDR1." & GF & ", 'YYYYMMDD')"

        ASCMAIN1.sql = "Select DISTINCT SOTORDR1.*, 'ST' CUST_ADDR_TYPE"
        If GF <> "ITEM_CODE" Then ASCMAIN1.sql &= "," & GF & " SORTEDBY "
        ASCMAIN1.sql &= " from SOTORDR1 where ORDR_NO in (Select Distinct ORDR_NO from " & SOTORDRL & ")"
        SOTORDR1 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR1 & " Add Primary Key (ORDR_NO)")
        ASCDATA1.ExecuteSQL("DELETE FROM " & SOTORDR1 & " WHERE ORDR_DATE IS NULL")

        ASCMAIN1.sql = "Select * from SOTORDR0 where ORDR_GROUP_NO in (Select Distinct ORDR_GROUP_NO from " & SOTORDR1 & ")"
        SOTORDR0 = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDR0 & " Add Primary Key (ORDR_GROUP_NO)")

        Dim sqlORDR_GROUP_NO As String = "SOTORDR1.ORDR_GROUP_NO"

        ASCMAIN1.sql = "Select * from " & SOTORDR1
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR1", 1))

        ASCMAIN1.sql = "Select * from " & SOTORDR0
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR0", 1))

        ASCMAIN1.sql = "Select SOTORDR5.* from SOTORDR5," & SOTORDR1 & " SOTORDR1 where SOTORDR5.ORDR_NO = SOTORDR1.ORDR_NO"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR5", 2))

        ASCMAIN1.sql = "Select SOTORDR2.ORDR_NO" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC * SOTORDR2.ORDR_UNIT_PRICE) ORDR_AMT_CANC" & vbCrLf _
            & " from SOTORDR2," & SOTORDRL & " SOTORDRL" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDRL.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTORDRL.ORDR_LNO" & vbCrLf _
            & " group by SOTORDR2.ORDR_NO"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDRT", 1))
        '& ", SOTORDR1.ORDR_SHIP_DATE, SOTORDR1.ORDR_CANCEL_DATE, SOTORDR1.ORDR_OVERRIDE_NOT_ALLOCATED" & vbCrLf _
        ASCMAIN1.sql = "Select SOTORDR2.*," & sqlORDR_GROUP_NO & " ORDR_GROUP_NO" & vbCrLf _
            & " from SOTORDR1,SOTORDR2," & SOTORDRL & " SOTORDRL" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDRL.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTORDRL.ORDR_LNO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDR2", 2))

        'Create_Relation("SOTORDR1", "SOTORDR2", "ORDR_NO")

        If chkALLOC.Checked Then Call Set_Allocations()

        ASCMAIN1.sql = "Select " & sqlORDR_GROUP_NO & " ORDR_GROUP_NO, SOTORDR2.ITEM_CODE" & vbCrLf _
            & ", MIN(SOTORDR2.ITEM_DESC) ITEM_DESC, MIN (CASE WHEN SOTORDR2.ORDR_QTY <> 0 THEN SOTORDR2.ORDR_UNIT_PRICE ELSE NULL END) ORDR_UNIT_PRICE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) AS ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) AS ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) AS ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) AS ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) AS ORDR_QTY_CANC" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_ORIG) AS ORDR_QTY_ORIG" & vbCrLf _
            & " from SOTORDR1,SOTORDR2," & SOTORDRL & " SOTORDRL" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDRL.ORDR_NO" & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTORDRL.ORDR_LNO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & " group by " & sqlORDR_GROUP_NO & ", SOTORDR2.ITEM_CODE"
        SOTORDRG = ASCMAIN1.Temp_Table
        ASCDATA1.ExecuteSQL("Alter Table " & SOTORDRG & " Add Primary Key (ORDR_GROUP_NO, ITEM_CODE)")
        dst.Tables.Add(ASCDATA1.GetDataTable("Select * from " & SOTORDRG, "SOTORDRG", 2))
        dst.Tables("SOTORDRG").Columns.Add("ORDR_RELEASE_CODES")

        For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_RELEASE <> 'A' and ISNULL(ORDR_RELEASE,'') <> ''", "ORDR_NO")
            Dim ORDR_GROUP_NO As String = rowSOTORDR2.Item("ORDR_GROUP_NO")
            Dim ITEM_CODE As String = rowSOTORDR2.Item("ITEM_CODE")
            Dim ORDR_RELEASE As String = rowSOTORDR2.Item("ORDR_RELEASE")
            Dim rowSOTORDRG As DataRow = dst.Tables("SOTORDRG").Rows.Find(New String() {ORDR_GROUP_NO, ITEM_CODE})
            Dim ORDR_RELEASE_CODES As String = rowSOTORDRG.Item("ORDR_RELEASE_CODES") & ""
            If Not ORDR_RELEASE_CODES.Contains(ORDR_RELEASE) Then
                ORDR_RELEASE_CODES &= ORDR_RELEASE
                rowSOTORDRG.Item("ORDR_RELEASE_CODES") = ORDR_RELEASE_CODES
            End If
        Next



        '' Total On Hold Orders
        ' '''sql = "Select ORDR_GROUP_NO, ORDR_CUST_PO, COUNT(*) AS NUM_HOLDS"
        'sql = "SELECT SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, "
        'sql = sql & " SOWORDR1.ORDR_CANCEL_DATE, COUNT(*) AS NUM_HOLDS"
        'sql = sql & " INTO SOWORDRZ"
        'sql = sql & " FROM SOWORDR1"
        'sql = sql & " WHERE ORDR_HOLD_SALES = '1'"
        ' '''sql = sql & " GROUP BY ORDR_GROUP_NO, ORDR_CUST_PO"
        'sql = sql & " GROUP BY SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, SOWORDR1.ORDR_CANCEL_DATE"
        'AccD.Execute sql

        'sql = "Update SOWORDRE, SOWORDRZ"
        'sql = sql & " Set SOWORDRE.NUM_HOLDS = SOWORDRZ.NUM_HOLDS"
        ' '''sql = sql & " Where SOWORDRE.ORDR_GROUP_NO = SOWORDRZ.ORDR_GROUP_NO"
        ' '''sql = sql & " AND SOWORDRE.ORDR_CUST_PO = SOWORDRZ.ORDR_CUST_PO"
        'sql = sql & " Where SOWORDRE.CUST_CODE = SOWORDRZ.CUST_CODE"
        'sql = sql & " And SOWORDRE.ORDR_CUST_PO = SOWORDRZ.ORDR_CUST_PO"
        'sql = sql & " And SOWORDRE.ORDR_GROUP_NO = SOWORDRZ.ORDR_GROUP_NO"
        ''sql = sql & " And SOWORDRE.ORDR_DATE = SOWORDRZ.ORDR_DATE"
        ''sql = sql & " And SOWORDRE.ORDR_CANCEL_DATE = SOWORDRZ.ORDR_CANCEL_DATE"
        'sql = sql & " And IIF(ISNULL(SOWORDRE.ORDR_DATE), now , SOWORDRE.ORDR_DATE) = IIF(ISNULL(SOWORDRZ.ORDR_DATE), now , SOWORDRZ.ORDR_DATE)"
        'sql = sql & " And IIF(ISNULL(SOWORDRE.ORDR_CANCEL_DATE), now , SOWORDRE.ORDR_CANCEL_DATE) = IIF(ISNULL(SOWORDRZ.ORDR_CANCEL_DATE), now , SOWORDRZ.ORDR_CANCEL_DATE)"

        'AccD.Execute sql

        'sql = "Update SOWORDRE Set DUP_COLLECTIVE = '*'"
        'sql = sql & " Where ORDR_GROUP_NO IN"
        'sql = sql & " (Select ORDR_GROUP_NO From SOWORDRE Group By ORDR_GROUP_NO Having Count(*) > 1)"
        'AccD.Execute sql

        If chkEXPCOLL.Checked Then ' EXPORT SHOULD ONLY BE DONE IF WE DO SWITCH COLLECTIVE NO INTO ORDER GROUP CODE
            ASCMAIN1.sql = "Select CUST_CODE, ORDR_CUST_PO, ORDR_GROUP_NO, EDI_DOC_SEQ_NO, ORDR_DATE, ORDR_CANCEL_DATE, NULL DUP, ORDR_CNT, ORDR_CNT TOTAL_HOLDS, ORDR_AMT from " & SOTORDR0 _
                & " where ORDR_GROUP_NO in (Select ORDR_GROUP_NO FROM " & SOTORDR0 & " minus" _
                & " Select DISTINCT ORDR_GROUP_NO from " & SOTORDR1 & " where ORDR_REL_HOLD_CODES like '%C%')"

            dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTORDRC", 0))

            dst.Tables("SOTORDRC").Columns("DUP").ReadOnly = False
            dst.Tables("SOTORDRC").Columns("TOTAL_HOLDS").ReadOnly = False
            ' dst.Tables("SOTORDRC").Columns("TOTAL_HOLDS").DataType = GetType(System.Int32)

            ASCMAIN1.sql = "SELECT DISTINCT X.CUST_CODE, X.ORDR_CUST_PO, X.EDI_DOC_SEQ_NO" _
                & " from " & SOTORDR1 & " X, SOTORDR1 Y" _
                & " where X.CUST_CODE = Y.CUST_CODE" _
                & "   and X.ORDR_CUST_PO = Y.ORDR_CUST_PO" _
                & "   and X.EDI_DOC_SEQ_NO <> Y.EDI_DOC_SEQ_NO" _
                & "   and (X.ORDR_STATUS IN ('O', 'P') AND Y.ORDR_STATUS IN ('O', 'P', 'F'))" _
                & "   and (X.ORDR_DATE >= (SYSDATE - 540) AND Y.ORDR_DATE >= (SYSDATE - 540))" _
                & "   and X.ORDR_SOURCE = 'E' AND Y.ORDR_SOURCE = 'E'"
            '                & "   and (X.ORDR_STATUS NOT IN ('C', 'D') AND Y.ORDR_STATUS NOT IN ('D', 'C'))" _

            For Each row As DataRow In ASCDATA1.GetDataTable.Select("")
                Dim CUST_CODE As String = row.Item("CUST_CODE") & ""
                Dim ORDR_CUST_PO As String = row.Item("ORDR_CUST_PO") & ""
                Dim EDI_DOC_SEQ_NO As String = row.Item("EDI_DOC_SEQ_NO") & ""
                Dim sql As String = "CUST_CODE = '" & CUST_CODE & "' and ORDR_CUST_PO = '" & ORDR_CUST_PO & "' and EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'"
                For Each rowSOTORDRC As DataRow In dst.Tables("SOTORDRC").Select(sql)
                    rowSOTORDRC.Item("DUP") = "*"
                Next
            Next

            For Each row As DataRow In dst.Tables("SOTORDRC").Select("")
                Dim EDI_DOC_SEQ_NO As String = row.Item("EDI_DOC_SEQ_NO") & ""
                'Dim C As Integer = dst.Tables("SOTORDRC").Select("EDI_DOC_SEQ_NO = '" & EDI_DOC_SEQ_NO & "'").Length
                'If C > 1 Then
                '    row.Item("DUP") = "*"
                '    End If
                Dim ORDR_GROUP_NO As String = row.Item("ORDR_GROUP_NO") & ""
                ASCMAIN1.sql = "Select Count (*) from SOTORDR1 where ORDR_GROUP_NO = '" & ORDR_GROUP_NO & "' and ORDR_HOLD_SALES = '1'"
                Dim TOTAL_HOLDS As Integer = Val(ASCDATA1.GetDataValue)
                If TOTAL_HOLDS <> 0 Then
                    row.Item("TOTAL_HOLDS") = TOTAL_HOLDS
                Else
                    row.Item("TOTAL_HOLDS") = DBNull.Value
                End If
            Next

            grdSOTORDRC.DisplayLayout.Bands(0).SortedColumns.Clear()
            grdSOTORDRC.DataSource = dst.Tables("SOTORDRC")
            With grdSOTORDRC.DisplayLayout.Bands(0)
                .Columns("CUST_CODE").Header.Caption = "Customer Code"
                .Columns("ORDR_CUST_PO").Header.Caption = "Order P.O."
                .Columns("ORDR_GROUP_NO").Header.Caption = "Group No"
                .Columns("EDI_DOC_SEQ_NO").Header.Caption = "Collective Number"
                .Columns("ORDR_DATE").Header.Caption = "Order Date"
                .Columns("ORDR_CANCEL_DATE").Header.Caption = "Order Cancel Date"
                .Columns("DUP").Header.Caption = "Duplicate"
                .Columns("ORDR_CNT").Header.Caption = "Total Orders"
                .Columns("TOTAL_HOLDS").Header.Caption = "Total Holds"
                .Columns("ORDR_AMT").Header.Caption = "Total Amount"
                '.Columns("ORDR_STATUS").Hidden = True
            End With

            grdSOTORDRC.DisplayLayout.Bands(0).Summaries.Clear()
            Create_Summary(grdSOTORDRC, New String() {"ORDR_CNT", "ORDR_AMT"})

            Sort_grdColumns(grdSOTORDRC, "CUST_CODE, ORDR_CUST_PO")
            'Show_Filter(grdSOTORDRC, True)
            'grdSOTORDRC.DisplayLayout.GroupByBox.Hidden = False
            'grdSOTORDRC.DisplayLayout.Bands(0).SortedColumns.Add("CUST_CODE", False, True)
            'grdSOTORDRC.DisplayLayout.Bands(0).SortedColumns.Add("ORDR_CUST_PO", False, True)
            grdSOTORDRC.DisplayLayout.Bands(0).Columns("ORDR_GROUP_NO").Hidden = True
            grdSOTORDRC.Visible = True
            btnExcelExport.Visible = True
            ' Export_Excel(GF)
        End If

    End Sub

    Function Get_Dates() As String
        Dim sql As String = ""
        For Each COLUMN_NAME As String In New String() {"ORDR_DATE_BOOKED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE", "ORDR_DATE_REL"}
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
                sql = sql & " and SOTORDR1." & COLUMN_NAME & " >= '" & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "dd-MMM-yyyy") & "'"
            End If
            If Not Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
                sql = sql & " and SOTORDR1." & COLUMN_NAME & " <= '" & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "dd-MMM-yyyy") & "'"
            End If
        Next
        'sql = Replace(sql, "SOTORDR1.ORDR_DATE_BOOKED", "TRUNC(SOTORDR1.INIT_DATE)")
        Return sql
    End Function

    Function Get_Dates_for_Caption(COLUMN_NAME As String) As String
        'Dim Z As String = Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text & ":"

        Dim Z As String = " " & Split(Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text, " ")(0)
        If Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Checked Then
            'Z &= " from First"
        Else
            Z &= " from " & Format(Absx1.dteFor(COLUMN_NAME & "_F").Value, "MM/dd/yyyy")
        End If
        If Absx1.chkFor("CHK" & COLUMN_NAME & "_L").Checked Then
            'Z &= " to Last"
        Else
            Z &= " to " & Format(Absx1.dteFor(COLUMN_NAME & "_L").Value, "MM/dd/yyyy")
        End If

        If Z = " " & Split(Absx1.chkFor("CHK" & COLUMN_NAME & "_F").Parent.Text, " ")(0) Then Z = ""

        Return Z
    End Function

    Overrides Sub Build_Report_File_Pre_Ora2ADO(ByVal TT As String)
        ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = 0")
        ASCDATA1.ExecuteSQL("Update " & TT & " Set ORDR_UNIT_PRICE = TRUNC(100 * ORDR_AMT / ORDR_QTY) / 100 where ORDR_QTY <> 0")
    End Sub

    Public Overrides Sub Print_Report()

        Dim SUBT As String = ""


        If optBy.Value = "ITEM_CODE" Then
            RPT = "SOROPNO2"
        Else
            RPT = "SOROPNO1"
        End If

        SUBT = "Open Orders " & optBy.Text

        Page0.Add("EDI: " & optEDI.Text)
        If optEDI.Value = "E" Then
            SUBT = SUBT & " (EDI Orders Only)"
        ElseIf optEDI.Value = "N" Then
            SUBT = SUBT & " (No EDI Orders)"
        End If
        If optIncl.Value = "R" Then
            SUBT = SUBT & " (Released Orders Only)"
        ElseIf optIncl.Value = "U" Then
            SUBT = SUBT & " (Un-Released Orders Only)"
        End If

        For Each COLUMN_NAME As String In New String() {"ORDR_DATE_BOOKED", "ORDR_SHIP_DATE", "ORDR_CANCEL_DATE"}
            Page0.Add(Get_Dates_for_Caption(COLUMN_NAME))
            SUBT = SUBT & Get_Dates_for_Caption(COLUMN_NAME)
        Next

        If chkINCLUDEF.Checked And chkINCLUDEC.Checked Then
            SUBT = SUBT & "; Including Shipped & Cancelled"
            Page0.Add("Including Shipped & Cancelled")
        ElseIf chkINCLUDEF.Checked Then
            SUBT = SUBT & "; Including Shipped"
            Page0.Add("Including Shipped")
        ElseIf chkINCLUDEC.Checked Then
            SUBT = SUBT & "; Including Cancelled"
            Page0.Add("Including Cancelled")
        End If

        If chkFILTER_DETAILS.Checked Then
            SUBT = SUBT & "; Filtered"
        End If

        CR_params.Add("DTL", IIf(chkDTL.Checked, "Y", "N"))
        CR_params.Add("BY", optBy.Value)

        If optBy.Value = "ITEM_CODE" Then
        Else
            CR_params.Add("SHOWSTADDRESS", IIf(chkSHOWSTADDRESS.Checked, "1", "0"))
            CR_params.Add("SHOW_GROUPS", IIf(chkSHOW_GROUPS.Checked, "1", "0"))
        End If

        Generate_Report(RPT, RPT_TITLE, SUBT)


        ' Added by Ed To Show Item by Ship Date
        If chkSUM_RPT.Checked Then
            RPT = "SOROPNO3"
            RPT_TITLE = "Open Sales Order Report - (Item / Ship Date Totals)"
            SUBT = Get_Dates_for_Caption("ORDR_SHIP_DATE")
            Generate_Report(RPT, RPT_TITLE, SUBT)
        End If

    End Sub

    Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        Select Case eItemKey
            Case "Proceed"
                'Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find("ITEM_CODE")
        End Select
    End Sub

    'Overrides Sub Set_ScreenMode_Special(ByVal tf As Boolean)

    'End Sub

    'Overrides Sub Update_Record()

    'End Sub

    'Overrides Sub Pivot_Prepare_PreProcess(dt As DataTable)
    '    dt.Columns.Remove("ORDR_GROUP_NO")
    '    dt.Columns.Remove("ITEM_CODE")
    'End Sub

    Sub Set_Allocations()
        ' Prepare Allocation Status Tables

        ASCMAIN1.Progress("Setting Up Allocation Work Tables", "")

        ASCMAIN1.sql = "Select SOTALLO2.* from SOTALLO2 where ALLO_CTL_NO in (Select ALLO_CTL_NO from " & SOTORDRL & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTALLO2", 2))

        ASCMAIN1.sql = "Select SOTALLO1.* from SOTALLO1 where ALLO_CTL_NO in (Select ALLO_CTL_NO from " & SOTORDRL & ")"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTALLO1", 1))

        ASCMAIN1.sql = "Select SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE) CUST_CODE" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY) ORDR_QTY" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_OPEN) ORDR_QTY_OPEN" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_PICK) ORDR_QTY_PICK" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_SHIP) ORDR_QTY_SHIP" & vbCrLf _
            & ", SUM (SOTORDR2.ORDR_QTY_CANC) ORDR_QTY_CANC" & vbCrLf _
            & ", 0 ORDR_QTY_PICK_NOW" & vbCrLf _
            & ", 0 ORDR_QTY_PICK_LATER" & vbCrLf _
            & " from SOTORDR2,SOTORDR1,ARTCUST1," & SOTORDRL & " SOTORDRL" & vbCrLf _
            & " where SOTORDR2.ORDR_NO = SOTORDRL.ORDR_NO " & vbCrLf _
            & "   and SOTORDR2.ORDR_LNO = SOTORDRL.ORDR_LNO" & vbCrLf _
            & "   and SOTORDR1.ORDR_NO = SOTORDR2.ORDR_NO" & vbCrLf _
            & "   and ARTCUST1.CUST_CODE = SOTORDR2.CUST_CODE" & vbCrLf _
            & "   and SOTORDR2.ALLO_CTL_NO IS NOT NULL" & vbCrLf _
            & " group by SOTORDR2.ALLO_CTL_NO, NVL(ARTCUST1.CUST_CODE_ALLO, SOTORDR2.CUST_CODE)"
        dst.Tables.Add(ASCDATA1.GetDataTable(ASCMAIN1.sql, "SOTALLOZ", 2))
        dst.Tables("SOTALLOZ").Columns("ORDR_QTY_PICK_NOW").ReadOnly = False

        Dim ORDR_REL_HOLD_CODES_dtl As String
        Dim cBAL As Long
        Dim ORDR_QTY_OPEN As Long
        Dim ORDR_NO As String
        Dim ORDER_RELEASE As String
        Dim ITEM_CODE As String

        ASCMAIN1.Progress("Allocation Order", "")

        Dim rowARTCUST1 As DataRow = Nothing
        Dim CUST_CODE As String = ""
        Dim CUST_CODE_ALLO As String = ""

        For Each rowSOTORDR1 As DataRow In dst.Tables("SOTORDR1").Select("", "ORDR_SHIP_DATE,ORDR_CANCEL_DATE,CUST_CODE")
            If CUST_CODE <> rowSOTORDR1.Item("CUST_CODE") Then
                CUST_CODE = rowSOTORDR1.Item("CUST_CODE")
                rowARTCUST1 = LookUp("ARTCUST1", CUST_CODE)
                If rowARTCUST1.Item("CUST_CODE_ALLO") & "" <> "" Then
                    CUST_CODE_ALLO = rowARTCUST1.Item("CUST_CODE_ALLO")
                Else
                    CUST_CODE_ALLO = CUST_CODE
                End If
            End If

            ORDR_NO = rowSOTORDR1.Item("ORDR_NO")
            ASCMAIN1.Progress("-", ORDR_NO)

            For Each rowSOTORDR2 As DataRow In dst.Tables("SOTORDR2").Select("ORDR_NO = '" & ORDR_NO & "' and ORDR_STATUS = 'O' and ISNULL(ALLO_CTL_NO,'') <> ''")
                Dim ALLO_CTL_NO As String = rowSOTORDR2.Item("ALLO_CTL_NO")
                ITEM_CODE = rowSOTORDR2.Item("ITEM_CODE")
                ' If ITEM_CODE = "82442715" Then Stop
                ORDR_QTY_OPEN = Val(rowSOTORDR2.Item("ORDR_QTY_OPEN") & "")
                ORDR_REL_HOLD_CODES_dtl = ""
                ORDER_RELEASE = ""

                Dim rowSOTALLO1 As DataRow = dst.Tables("SOTALLO1").Rows.Find(ALLO_CTL_NO)
                Dim rowSOTALLO2 As DataRow = dst.Tables("SOTALLO2").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO})
                If rowSOTALLO2 Is Nothing Then
                    rowSOTALLO2 = dst.Tables("SOTALLO2").NewRow
                    rowSOTALLO2.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                    rowSOTALLO2.Item("CUST_CODE") = CUST_CODE_ALLO
                    rowSOTALLO2.Item("QTY_ALLO") = 0
                    dst.Tables("SOTALLO2").Rows.Add(rowSOTALLO2)
                End If

                Dim rowSOTALLOZ As DataRow = dst.Tables("SOTALLOZ").Rows.Find(New String() {ALLO_CTL_NO, CUST_CODE_ALLO})
                If rowSOTALLOZ Is Nothing Then
                    rowSOTALLOZ = dst.Tables("SOTALLOZ").NewRow
                    rowSOTALLOZ.Item("ALLO_CTL_NO") = ALLO_CTL_NO
                    rowSOTALLOZ.Item("CUST_CODE") = CUST_CODE_ALLO
                    rowSOTALLOZ.Item("ORDR_QTY") = 0
                    rowSOTALLOZ.Item("ORDR_QTY_OPEN") = 0
                    rowSOTALLOZ.Item("ORDR_QTY_PICK") = 0
                    rowSOTALLOZ.Item("ORDR_QTY_SHIP") = 0
                    rowSOTALLOZ.Item("ORDR_QTY_CANC") = 0
                    rowSOTALLOZ.Item("ORDR_QTY_PICK_NOW") = 0
                    rowSOTALLOZ.Item("ORDR_QTY_PICK_LATER") = 0
                    dst.Tables("SOTALLOZ").Rows.Add(rowSOTALLOZ)
                End If

                ORDER_RELEASE = "A"
                If Val(rowSOTALLO2.Item("QTY_ALLO") & "") = 0 And rowSOTORDR1.Item("ORDR_OVERRIDE_NOT_ALLOCATED") & "" = "1" Then
                    ' LET THE ORDER GO
                Else
                    If Format(Now, "YYYYMMDD") < Format(rowSOTALLO1.Item("DATE_START"), "YYYYMMDD") Then
                        ORDR_REL_HOLD_CODES_dtl = "P"
                    End If

                    If rowSOTALLO2 Is Nothing Then
                        cBAL = 0
                    Else
                        cBAL = Val(rowSOTALLO2.Item("QTY_ALLO") & "")
                    End If

                    cBAL = cBAL _
                         - Val(rowSOTALLOZ.Item("ORDR_QTY_SHIP") & "") _
                         - Val(rowSOTALLOZ.Item("ORDR_QTY_PICK") & "") _
                         - Val(rowSOTALLOZ.Item("ORDR_QTY_PICK_NOW") & "")

                    cBAL = cBAL - ORDR_QTY_OPEN 'Val(rowSOTORDR2.Item("ORDR_QTY") & "")

                    'If ORDR_QTY_OPEN + iaq > cBAL And rowSOTALLO1.Item("ALLOW_OVER") & "" <> "1" Then
                    If cBAL < 0 And (rowSOTALLO1.Item("ALLOW_OVER") & "" <> "1") Then
                        ORDER_RELEASE = "C"
                    End If

                End If

                Dim sqlw = "ORDR_NO = '" & ORDR_NO & "' AND ITEM_CODE = '" & ITEM_CODE & "'" _
                    & " AND (ORDR_RELEASE = 'A' or ORDR_RELEASE = 'C' OR ISNULL(ORDR_RELEASE,'') = '')"
                For Each row As DataRow In dst.Tables("SOTORDR2").Select(sqlw)
                    row.Item("ORDR_RELEASE") = ORDER_RELEASE
                Next

                rowSOTALLOZ.Item("ORDR_QTY_PICK_NOW") = Val(rowSOTALLOZ.Item("ORDR_QTY_PICK_NOW")) + ORDR_QTY_OPEN
            Next
        Next
    End Sub

    'Sub Export_Excel(Sortedby As String)
    '    ' Prints the Excelwork book for P&G Collective Number Upload
    '    Dim iRow As Integer
    '    Dim i As Integer
    '    Dim j As Integer
    '    Dim z As String
    '    Dim r As Long

    '    Dim dynSOWORDRE As Recordset

    '    ASCMAIN1.Progress("Create Collective Number Workbook", "")

    '    sql = " SELECT SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, SOWORDR1.ORDR_CANCEL_DATE,"
    '    sql = sql & " MIN(' ') AS DUP_COLLECTIVE, COUNT(*) AS NUM_ORDERS, SUM(0) AS NUM_HOLDS, SUM(0.00) AS AMOUNT"
    '    sql = sql & " INTO SOWORDRE FROM SOWORDR1"
    '    sql = sql & " WHERE SOWORDR1.ORDR_STATUS = 'O' "
    '    sql = sql & " AND ORDR_GROUP_NO NOT IN (SELECT DISTINCT SOWORDRG.ORDR_GROUP_NO FROM SOWORDRG WHERE ORDR_RELEASE_CODES LIKE '*C*')"
    '    sql = sql & " GROUP BY SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, SOWORDR1.ORDR_CANCEL_DATE"
    '    sql = sql & " ORDER BY"
    '    If optBy.Value <> "" Then
    '        sql = sql & " SOWORDR1." & Excel_Order_by & ","
    '    End If
    '    sql = sql & " SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, SOWORDR1.ORDR_CANCEL_DATE"
    '    AccD.Execute sql

    '    ' Total $$ Value for Orders
    '    sql = "SELECT SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, "
    '    sql = sql & " SOWORDR1.ORDR_DATE, SOWORDR1.ORDR_CANCEL_DATE, SUM(SOWORDR2.ORDR_UNIT_PRICE * SOWORDR2.ORDR_QTY_OPEN) AS AMOUNT"
    '    sql = sql & " INTO SOWORDRX"
    '    sql = sql & " FROM SOWORDR1, SOWORDR2"
    '    sql = sql & " WHERE SOWORDR1.ORDR_NO = SOWORDR2.ORDR_NO"
    '    sql = sql & " AND ORDR_GROUP_NO NOT IN (SELECT DISTINCT SOWORDRG.ORDR_GROUP_NO FROM SOWORDRG WHERE ORDR_RELEASE_CODES LIKE '*C*')"
    '    sql = sql & " GROUP BY SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, SOWORDR1.ORDR_CANCEL_DATE"
    '    AccD.Execute sql

    '    sql = "Update SOWORDRE, SOWORDRX"
    '    sql = sql & " Set SOWORDRE.AMOUNT = SOWORDRX.AMOUNT"
    '    sql = sql & " Where SOWORDRE.CUST_CODE = SOWORDRX.CUST_CODE"
    '    sql = sql & " And SOWORDRE.ORDR_CUST_PO = SOWORDRX.ORDR_CUST_PO"
    '    sql = sql & " And SOWORDRE.ORDR_GROUP_NO = SOWORDRX.ORDR_GROUP_NO"
    '    sql = sql & " And IIF(ISNULL(SOWORDRE.ORDR_DATE), now , SOWORDRE.ORDR_DATE) = IIF(ISNULL(SOWORDRX.ORDR_DATE), now , SOWORDRX.ORDR_DATE)"
    '    sql = sql & " And IIF(ISNULL(SOWORDRE.ORDR_CANCEL_DATE), now , SOWORDRE.ORDR_CANCEL_DATE) = IIF(ISNULL(SOWORDRX.ORDR_CANCEL_DATE), now , SOWORDRX.ORDR_CANCEL_DATE)"
    '    AccD.Execute sql

    '    'IIF(ISNULL(field),alt value,field)

    '    ' Total On Hold Orders
    '    sql = "SELECT SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, "
    '    sql = sql & " SOWORDR1.ORDR_CANCEL_DATE, COUNT(*) AS NUM_HOLDS"
    '    sql = sql & " INTO SOWORDRZ"
    '    sql = sql & " FROM SOWORDR1"
    '    sql = sql & " WHERE ORDR_HOLD_SALES = '1'"
    '    sql = sql & " GROUP BY SOWORDR1.CUST_CODE, SOWORDR1.ORDR_CUST_PO, SOWORDR1.ORDR_GROUP_NO, SOWORDR1.ORDR_DATE, SOWORDR1.ORDR_CANCEL_DATE"
    '    AccD.Execute sql

    '    sql = "Update SOWORDRE, SOWORDRZ"
    '    sql = sql & " Set SOWORDRE.NUM_HOLDS = SOWORDRZ.NUM_HOLDS"
    '    sql = sql & " Where SOWORDRE.CUST_CODE = SOWORDRZ.CUST_CODE"
    '    sql = sql & " And SOWORDRE.ORDR_CUST_PO = SOWORDRZ.ORDR_CUST_PO"
    '    sql = sql & " And SOWORDRE.ORDR_GROUP_NO = SOWORDRZ.ORDR_GROUP_NO"
    '    sql = sql & " And IIF(ISNULL(SOWORDRE.ORDR_DATE), now , SOWORDRE.ORDR_DATE) = IIF(ISNULL(SOWORDRZ.ORDR_DATE), now , SOWORDRZ.ORDR_DATE)"
    '    sql = sql & " And IIF(ISNULL(SOWORDRE.ORDR_CANCEL_DATE), now , SOWORDRE.ORDR_CANCEL_DATE) = IIF(ISNULL(SOWORDRZ.ORDR_CANCEL_DATE), now , SOWORDRZ.ORDR_CANCEL_DATE)"

    '    AccD.Execute sql

    '    sql = "Update SOWORDRE Set DUP_COLLECTIVE = '*'"
    '    sql = sql & " Where ORDR_GROUP_NO IN"
    '    sql = sql & " (Select ORDR_GROUP_NO From SOWORDRE Group By ORDR_GROUP_NO Having Count(*) > 1)"
    '    AccD.Execute sql

    '    sql = "SELECT * FROM SOWORDRE"
    '    dynSOWORDRE = AccD.OpenRecordset(sql)

    '    Dim objApp As excel.Application
    '    Dim objBook As excel.Workbook
    '    Dim objSheet As excel.Worksheet

    '    objApp = New excel.Application
    '    objBook = objApp.Workbooks.Add

    '    objApp.DisplayAlerts = False
    '    For i = objBook.Worksheets.Count To 2 Step -1
    '        objBook.Worksheets(i).Delete()
    '    Next i

    '    objSheet = objBook.Worksheets(1)
    '    objSheet.Name = "Collective Numbers"

    '    ' Create Header Row in Excel
    '    iRow = 1

    '    objSheet.Cells(iRow, 1).Value = "Customer Code"
    '    objSheet.Cells(iRow, 2).Value = "Order P.O."
    '    objSheet.Cells(iRow, 3).Value = "Collective Number"
    '    objSheet.Cells(iRow, 4).Value = "Order Date"
    '    objSheet.Cells(iRow, 5).Value = "Order Cancel Date"
    '    objSheet.Cells(iRow, 6).Value = " "
    '    objSheet.Cells(iRow, 7).Value = "Total Orders"
    '    objSheet.Cells(iRow, 8).Value = "Total Holds"
    '    objSheet.Cells(iRow, 9).Value = "Total Amount"

    '    For i = 1 To 3
    '        objSheet.Columns(i).EntireColumn.HorizontalAlignment = xlLeft
    '        z = Chr(64 + i)
    '        objSheet.Columns(z & ":" & z).NumberFormat = "@"
    '    Next

    '    z = Chr(64 + 4)
    '    objSheet.Columns(z & ":" & z).NumberFormat = "MM/DD/YYYY"
    '    objSheet.Columns(z & ":" & z).HorizontalAlignment = xlRight

    '    z = Chr(64 + 5)
    '    objSheet.Columns(z & ":" & z).NumberFormat = "MM/DD/YYYY"
    '    objSheet.Columns(z & ":" & z).HorizontalAlignment = xlRight

    '    z = Chr(64 + 6)
    '    objSheet.Columns(z & ":" & z).NumberFormat = "#,##0"
    '    objSheet.Columns(z & ":" & z).HorizontalAlignment = xlRight

    '    z = Chr(64 + 7)
    '    objSheet.Columns(z & ":" & z).NumberFormat = "#,##0"
    '    objSheet.Columns(z & ":" & z).HorizontalAlignment = xlRight

    '    z = Chr(64 + 8)
    '    objSheet.Columns(z & ":" & z).NumberFormat = "#,##0"
    '    objSheet.Columns(z & ":" & z).HorizontalAlignment = xlRight

    '    z = Chr(64 + 9)
    '    objSheet.Columns(z & ":" & z).NumberFormat = "#,##0.00"
    '    objSheet.Columns(z & ":" & z).HorizontalAlignment = xlRight

    '    j = 10

    '    ' Underline the Titles
    '    z = Excel_Cell(1, j - 1)
    '    objSheet.range("A1:" & z).Borders(xlDiagonalDown).LineStyle = xlNone
    '    objSheet.range("A1:" & z).Borders(xlDiagonalUp).LineStyle = xlNone
    '    objSheet.range("A1:" & z).Borders(xlEdgeLeft).LineStyle = xlNone
    '    objSheet.range("A1:" & z).Borders(xlEdgeTop).LineStyle = xlNone
    '    With objSheet.range("A1:" & z).Borders(xlEdgeBottom)
    '        .LineStyle = xlContinuous
    '        .Weight = xlThin
    '        .ColorIndex = xlAutomatic
    '    End With
    '    objSheet.range("A1:" & z).Borders(xlEdgeRight).LineStyle = xlNone
    '    objSheet.range("A1:" & z).Borders(xlInsideVertical).LineStyle = xlNone
    '    objSheet.range("A1:" & z).Interior.ColorIndex = 6
    '    objSheet.range("A1:" & z).Font.Bold = True


    '    On Error Resume Next

    '    With objBook.ActiveSheet.PageSetup
    '        .PrintTitleRows = "$1:$1"
    '        .PrintTitleColumns = ""
    '    End With

    '    With objBook.ActiveSheet.PageSetup
    '        .LeftHeader = ""
    '        .CenterHeader = ""
    '        .RightHeader = ""
    '        .LeftFooter = ""
    '        .CenterFooter = ""
    '        .RightFooter = ""
    '        .LeftMargin = objApp.InchesToPoints(0.75)
    '        .RightMargin = objApp.InchesToPoints(0.75)
    '        .TopMargin = objApp.InchesToPoints(1)
    '        .BottomMargin = objApp.InchesToPoints(1)
    '        .HeaderMargin = objApp.InchesToPoints(0.5)
    '        .FooterMargin = objApp.InchesToPoints(0.5)
    '        .PrintHeadings = False
    '        .PrintGridlines = False
    '        .PrintComments = xlPrintNoComments
    '        '.PrintQuality = 600
    '        .CenterHorizontally = False
    '        .CenterVertically = False
    '        .Orientation = xlLandscape
    '        .Draft = False
    '        .PaperSize = xlPaperLetter
    '        .FirstPageNumber = xlAutomatic
    '        .Order = xlDownThenOver
    '        .BlackAndWhite = False
    '        .Zoom = 100
    '        .PrintErrors = xlPrintErrorsDisplayed
    '    End With

    '    On Error GoTo 0

    '    objSheet.range("A2").CopyFromRecordset dynSOWORDRE

    '    dynSOWORDRE.Close()

    '    j = 8
    '    i = 2
    '    While objSheet.Cells(i, 1).Value <> ""
    '        If Val(objSheet.Cells(i, j).Value) = 0 Then
    '            objSheet.Cells(i, j).Value = " "
    '        Else
    '            objSheet.range("A" & i & ":" & "I" & i).Interior.ColorIndex = 6
    '        End If
    '        i = i + 1
    '    End While

    '    objSheet.Columns.AutoFit()
    '    'objApp.Visible = True
    '    objBook.Worksheets(objBook.Worksheets.Count).Activate()

    '    xRptTitle = "Collective Number Report"
    '    Call Write_SPRF()
    '    xRptTitle = ""

    '    objBook.SaveAs AppPath & "\temp\" & xUID & "_" & REPORT_NO & ".xls"
    '    objApp.Quit()
    '    'objApp.Visible = True

    '    On Error Resume Next
    '    objSheet = Nothing
    '    objBook = Nothing
    '    objApp = Nothing
    '    On Error GoTo 0
    'End Sub

    Private Sub grdSOTORDRC_InitializeRow(sender As Object, e As UltraWinGrid.InitializeRowEventArgs) Handles grdSOTORDRC.InitializeRow
        If Val(e.Row.Cells("TOTAL_HOLDS").Value & "") <> 0 Then
            e.Row.Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("CUST_CODE").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("ORDR_CUST_PO").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("ORDR_GROUP_NO").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("EDI_DOC_SEQ_NO").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("ORDR_DATE").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("ORDR_CANCEL_DATE").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("DUP").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("ORDR_CNT").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("TOTAL_HOLDS").Appearance.BackColor = Drawing.Color.Yellow
            e.Row.Cells("ORDR_AMT").Appearance.BackColor = Drawing.Color.Yellow
        End If
    End Sub

    'Private Sub btnExcelExport_Click(sender As Object, e As EventArgs) Handles btnExcelExport.Click
    '    Export_to_Excel_Custom(grdSOTORDRC, True, False, grdSOTORDRC.Text, "B")
    'End Sub

    Private Sub btnExcelExport_Click(sender As Object, e As EventArgs) Handles btnExcelExport.Click

    End Sub

    Private Sub chkALLOC_CheckedChanged(sender As Object, e As EventArgs) Handles chkALLOC.CheckedChanged

    End Sub
End Class