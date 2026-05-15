Imports ABSolution
Imports Infragistics.Win
Public Class POROPEN1

    Dim POTOPEN1 As String = "" ' Logical Name of Work Table, which will be used as the name of the DataTable
    Dim H() As String
    Dim PERIODS As Integer = 0
    Dim PERIODS_max As Integer = 24
    Private Sub Form_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Range_Events(grpPO_DATE_RANGE)

        Get_PARM("POTPARM1")

        'Set_cmbYP("RYP", ASCMAIN1.CYP, -36, -1, -1)
    End Sub

#Region "Popup_Menus"

    Overrides Sub Load_Popup_Menus()
        Load_Popup_Menu(grdPOROPEN1, "SSS", "Show Filter", "Show GroupBox", "Show Pins")
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


    Protected Overrides Sub Build_Workfile()
        MyBase.Build_Workfile()
        ASCMAIN1.Progress("Building Work File")

        grdPOROPEN1.Visible = False

        Dim sql As String = ""

        ' Prepare filters from Run-Time Options

        Dim sql_filter As String = ""

        Dim DTX As String = IIf(Absx1.optFor("OPTDATE").Value = "R", "POTORDR2.PO_DATE_REQUIRED", "POTORDR2.PO_DATE_REQUIRED + NVL(TATTERM1.TERM_DAYS_DUE,0)")

        If Absx1.optFor("OPTPERIODS").Value = "P" Then
            PERIODS = 12
        ElseIf Absx1.optFor("OPTPERIODS").Value = "L" Then
            PERIODS = 18
        ElseIf Absx1.optFor("OPTPERIODS").Value = "X" Then
            PERIODS = 24
        End If
        ReDim H(PERIODS)
        ReDim H(PERIODS_max)

        ' Extracts from Data Sources

        Dim P() As String
        ReDim P(PERIODS)
        ReDim P(PERIODS_max)
        If Absx1.optFor("OPTAGE").Value = "W" Then
            Dim dt As Date = Now.AddDays(-7)
            For i As Int32 = 0 To PERIODS_max ' PERIODS
                P(i) = Format(dt, "dd-MMM-yyyy")
                H(i) = Format(dt, "MM/dd/yy")
                dt = dt.AddDays(7)
            Next
        Else
            For i As Int32 = 0 To PERIODS_max ' PERIODS
                P(i) = ASCMAIN1.Period_Calc(ASCMAIN1.CYM, i - 1)
                H(i) = ASCMAIN1.Get_Legend(ASCMAIN1.Period_Calc(ASCMAIN1.CYP, i - 1), False, True)
            Next
        End If

        If Absx1.optFor("OPTAGE").Value = "M" Then
            DTX = "TO_CHAR(" & DTX & ", 'YYYYMM')"
        End If

        Dim VX As String = IIf(Absx1.optFor("OPTSHOW").Value = "U", "NVL(PO_QTY_OPN,0)", "NVL(PO_QTY_OPN,0) * NVL(PO_COST,0)")


        If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
            Dim z As String = Format(Absx1.dteFor("PO_DATE_F").Value, "dd-MMM-yyyy")
            sql = sql & " and POTORDR1.PO_DATE_ORDERED >= '" & z & "'" & vbCrLf
            Page0.Add("Purchase Order Date >= " & z)
        End If
        If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
            Dim z As String = Format(Absx1.dteFor("PO_DATE_L").Value, "dd-MMM-yyyy")
            sql = sql & " and POTORDR1.PO_DATE_ORDERED <= '" & z & "'" & vbCrLf
            Page0.Add("Purchase Order Date <= " & z)
        End If



        ' Create the Work File - as a flattened out result set containing all PB columns and data fields that you would like to put in the detail section of the report

        Dim sqlsum As String = ""
        POTOPEN1 = ""

        If Absx1.chkFor("CHKPO").Checked Then

            sql = "Select POTORDR1.VEND_CODE, POTORDR2.ITEM_CODE, POTORDR1.MARKET_CODE" & vbCrLf
            sql &= ", SUM (CASE WHEN " & DTX & " <= '" & P(0) & "' THEN " & VX & " ELSE 0 END) PO_PD" & vbCrLf

            sqlsum = "PO_TOT = NVL(PO_PD,0)"
            For i As Int32 = 1 To PERIODS_max ' PERIODS
                sql &= ", SUM (CASE WHEN " & DTX & " > '" & P(i - 1) & "' and " & DTX & " <= '" & P(i) & "' THEN " & VX & " ELSE 0 END) PO_" & Format(i, "00") & vbCrLf
                If i <= PERIODS Then sqlsum &= "+NVL(PO_" & Format(i, "00") & ",0)"
            Next
            sql &= ", 0 PO_TOT" & vbCrLf
            sql &= ", 0 PP_PD" & vbCrLf
            For i As Int32 = 1 To PERIODS_max ' PERIODS
                sql &= ", 0 PP_" & Format(i, "00")
            Next
            sql &= ", 0 PP_TOT" & vbCrLf
            sql &= "   from POTORDR1,POTORDR2,ICTITEM1" & IIf(Absx1.optFor("OPTDATE").Value = "R", "", ",TATTERM1") & vbCrLf _
            & "   where POTORDR1.PO_ORDER_NO = POTORDR2.PO_ORDER_NO" & vbCrLf _
            & " and " & VX & " <> 0" & vbCrLf _
            & " and POTORDR2.PO_STATUS = 'O' " & vbCrLf _
            & " and POTORDR2.PO_ORDER_NO = POTORDR1.PO_ORDER_NO" & vbCrLf _
            & " and ICTITEM1.ITEM_CODE = POTORDR2.ITEM_CODE" & vbCrLf _
            & IIf(Absx1.optFor("OPTDATE").Value = "R", "", " and TATTERM1.TERM_CODE (+) = POTORDR1.TERM_CODE") & vbCrLf

            ' one of these for each filter in Report Maintenance
            sql &= SQL_in("VEND_CODE", "POTORDR1.VEND_CODE")
            'sql &= SQL_in("BUYER_CODE", "BUYER_CODE")
            'sql &= SQL_in("PO_ORDER_TYPE", "PO_ORDER_TYPE")
            'sql &= SQL_in("PO_ORDER_NO", "PO_ORDER_NO")
            sql &= SQL_in("ITEM_CODE", "POTORDR2.ITEM_CODE")
            sql &= SQL_in("MARKET_CODE", "POTORDR1.MARKET_CODE")

            If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
                Dim z As String = Format(Absx1.dteFor("PO_DATE_F").Value, "dd-MMM-yyyy")
                sql = sql & " and POTORDR1.PO_DATE_ORDERED >= '" & z & "'" & vbCrLf
                Page0.Add("Purchase Order Date >= " & z)
            End If
            If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
                Dim z As String = Format(Absx1.dteFor("PO_DATE_L").Value, "dd-MMM-yyyy")
                sql = sql & " and POTORDR1.PO_DATE_ORDERED <= '" & z & "'" & vbCrLf
                Page0.Add("Purchase Order Date <= " & z)
            End If

            sql &= " group by POTORDR1.VEND_CODE, POTORDR2.ITEM_CODE, POTORDR1.MARKET_CODE" & vbCrLf

            POTOPEN1 = ASCMAIN1.Temp_Table(sql)
            ASCDATA1.ExecuteSQL("Update " & POTOPEN1 & " Set " & sqlsum)
        End If

        If Absx1.chkFor("CHKPP").Checked Then
            DTX = IIf(Absx1.optFor("OPTDATE").Value = "R", "DPTPLAN1.DATE_REQUIRED", "DPTPLAN1.DATE_REQUIRED + NVL(TATTERM1.TERM_DAYS_DUE,0)")
            If Absx1.optFor("OPTAGE").Value = "M" Then
                DTX = "TO_CHAR(" & DTX & ", 'YYYYMM')"
            End If
            VX = IIf(Absx1.optFor("OPTSHOW").Value = "U", "NVL(DPTPLAN1.QTY_PLANNED,0)", "NVL(DPTPLAN1.QTY_PLANNED,0) * NVL(ICTCOSTC.ITEM_COST_VCOST,0)")

            sql = "Select DPTPLAN1.VEND_CODE, DPTPLAN1.ITEM_CODE, '*' MARKET_CODE" & vbCrLf

            sql &= ", 0 PO_PD" & vbCrLf
            For i As Int32 = 1 To PERIODS_max ' PERIODS
                sql &= ", 0 PO_" & Format(i, "00")
            Next
            sql &= ", 0 PO_TOT" & vbCrLf

            sql &= ", SUM (CASE WHEN " & DTX & " <= '" & P(0) & "' THEN " & VX & " ELSE 0 END) PP_PD" & vbCrLf
            sqlsum = "PP_TOT = NVL(PP_PD,0)"
            For i As Int32 = 1 To PERIODS_max ' PERIODS
                sql &= ", SUM (CASE WHEN " & DTX & " > '" & P(i - 1) & "' and " & DTX & " <= '" & P(i) & "' THEN " & VX & " ELSE 0 END) PP_" & Format(i, "00") & vbCrLf
                If i <= PERIODS Then sqlsum &= "+NVL(PP_" & Format(i, "00") & ",0)"
            Next
            sql &= ", 0 PP_TOT" & vbCrLf


            sql &= "   from DPTPLAN1,ICTITEM1,ICTCOSTC" & IIf(Absx1.optFor("OPTDATE").Value = "R", "", ",APTVEND1,TATTERM1") & vbCrLf _
            & " where ICTITEM1.ITEM_CODE = DPTPLAN1.ITEM_CODE" & vbCrLf _
            & "   and ICTCOSTC.ITEM_CODE = DPTPLAN1.ITEM_CODE" & vbCrLf _
            & " and " & VX & " <> 0" & vbCrLf _
            & IIf(Absx1.optFor("OPTDATE").Value = "R", "", " and APTVEND1.VEND_CODE (+) = DPTPLAN1.VEND_CODE and TATTERM1.TERM_CODE (+) = APTVEND1.TERM_CODE") & vbCrLf

            ' one of these for each filter in Report Maintenance
            sql &= SQL_in("VEND_CODE", "DPTPLAN1.VEND_CODE")
            'sql &= SQL_in("BUYER_CODE", "BUYER_CODE")
            'sql &= SQL_in("PO_ORDER_TYPE", "PO_ORDER_TYPE")
            'sql &= SQL_in("PO_ORDER_NO", "PO_ORDER_NO")
            sql &= SQL_in("ITEM_CODE", "DPTPLAN1.ITEM_CODE")
            'sql &= SQL_in("MARKET_CODE", "POTORDR1.MARKET_CODE")

            sql &= " group by DPTPLAN1.VEND_CODE, DPTPLAN1.ITEM_CODE " & vbCrLf ', POTORDR1.MARKET_CODE" & vbCrLf

            If POTOPEN1 = "" Then
                POTOPEN1 = ASCMAIN1.Temp_Table(sql)
            Else
                ASCMAIN1.sql = "Insert into " & POTOPEN1 & " " & sql
                ASCDATA1.ExecuteSQL()
            End If
            ASCDATA1.ExecuteSQL("Update " & POTOPEN1 & " Set " & sqlsum)
        End If

        MyBase.Get_SQL("*", POTOPEN1)
        ASCMAIN1.Progress("Building Tiers")

        sql = "Select " & sql_SELECT_cols & ASTSRPT1_sum_columns & vbCr
        sql &= " from " & POTOPEN1 & " POTOPEN1 " & sql_TABLE_NAMEs & vbCr
        sql &= ASCMAIN1.SQL_Add_WHERE(sql_WHERE & sql_JOIN & sql_filter) & vbCr
        sql &= " group by " & sql_GROUP_BY_cols
        ASCDATA1.ExecuteSQL("Insert into " & ASTSRPT1 & " (" & sql & ")")

        For i As Int32 = 1 To PERIODS
            Dim rowASTDSQLS As DataRow
            rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PO_" & Format(i, "00") & "'")(0)
            rowASTDSQLS.Item("COLUMN_CAPTION") = "PO " & H(i)
            rowASTDSQLS = tblASTDSQLS.Select("COLUMN_NAME = 'PP_" & Format(i, "00") & "'")(0)
            rowASTDSQLS.Item("COLUMN_CAPTION") = "PP " & H(i)
        Next



    End Sub

    Public Overrides Sub Print_Report()
        Dim SUBT As String = ""

        'If Absx1.optFor("OPTPERIODS").Value = "L" Then
        '    RPT = "POROPEN5"
        'End If

        If Not Absx1.chkFor("CHKPO_DATE_F").Checked _
        Or Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
            SUBT = SUBT & "Showing Purchase Orders Dated"
            If Not Absx1.chkFor("CHKPO_DATE_F").Checked Then
                SUBT = SUBT & " from " & Format(Absx1.dteFor("PO_DATE_F").Value, "MM/dd/yyyy")
            End If
            If Not Absx1.chkFor("CHKPO_DATE_L").Checked Then
                SUBT = SUBT & " thru " & Format(Absx1.dteFor("PO_DATE_L").Value, "MM/dd/yyyy")
            End If
        End If

        CR_params.Add("PO", IIf(Absx1.chkFor("CHKPO").Checked, "1", "0"))
        CR_params.Add("PP", IIf(Absx1.chkFor("CHKPP").Checked, "1", "0"))
        CR_params.Add("PO_TEXT", "PO's")
        CR_params.Add("PP_TEXT", "Plans")

        For i As Int32 = 1 To 12 ' PERIODS
            CR_params.Add("H" & Format(i, "00"), H(i))
        Next

        Generate_Report(RPT, , SUBT)

        'grdPOROPEN1.DataSource = dst.Tables("")
        'grdPOROPEN1.Visible = True

        Prepare_Data_Extracts()
    End Sub

    Sub Prepare_Data_Extracts()

        grdASTEXPT1.DisplayLayout.ViewStyle = UltraWinGrid.ViewStyle.SingleBand

        With dst.Tables("ASTSRPT1").Columns
            .Add("GROUP_DESC")
        End With

        Dim rowASTDSQLA As DataRow = tblASTDSQLA.Rows.Find(COLUMN_NAMEs(COLUMN_NAMEs.Count - 1))
        For Each row As DataRow In dst.Tables("ASTSRPT1").Select("")
            For I As Integer = 1 To COLUMN_NAMEs.Count
                Dim CODE_VALUE As String = row.Item("G" & CStr(I))
                row.Item("G" & CStr(I)) = Split(CODE_VALUE, ":")(1)
            Next
            Dim GROUP_CODE As String = row.Item("G" & CStr(COLUMN_NAMEs.Count))
            Dim GK As String = rowASTDSQLA.Item("COLUMN_CAPTION") & ":" & GROUP_CODE

            Dim rowASTGROUP As DataRow = dst.Tables("ASTGROUP").Rows.Find(GK)
 
            If rowASTGROUP IsNot Nothing Then
                row.Item("GROUP_DESC") = rowASTGROUP.Item("GROUP_DESC")
            End If
        Next

        grdASTEXPT1.DataSource = dst.Tables("ASTSRPT1")

        grdASTEXPT1.Text = "PO Summary by Item"
        UltraTabControl1.Tabs("Data Exports").Visible = True
        tabDataExports.Tabs(0).Text = grdASTEXPT1.Text

        Set_DX_Column(grdASTEXPT1, "")
        Dim Cs As New List(Of String)
        Dim G As Integer = 0
        For Each COLUMN_NAME As String In COLUMN_NAMEs
            Cs.Add(COLUMN_NAME)
            G += 1
            Set_DX_Column(grdASTEXPT1, "G" & CStr(G), COLUMN_CAPTIONs(G - 1), 100, , , System.Drawing.Color.Gold)
        Next
 
        Set_DX_Column(grdASTEXPT1, "GROUP_DESC", "Description", 130)
 

        If Absx1.chkFor("CHKPO").Checked Then
            For i As Int32 = 0 To PERIODS
                Dim PX_DATA As String = "PO_" & Format(i, "00")
                Dim PX_DESC As String = "PO " & H(i)
                If i = 0 Then
                    PX_DATA = "PO_PD"
                    PX_DESC = "PO Past Due"
                End If
                Set_DX_Column(grdASTEXPT1, PX_DATA, PX_DESC, 90, "#,##0", , System.Drawing.Color.LightGreen)
                Create_Summary(grdASTEXPT1, PX_DATA)
            Next
        End If

        If Absx1.chkFor("CHKPP").Checked Then
            For i As Int32 = 0 To PERIODS
                Dim PX_DATA As String = "PP_" & Format(i, "00")
                Dim PX_DESC As String = "PP " & H(i)
                If i = 0 Then
                    PX_DATA = "PP_PD"
                    PX_DESC = "PP Past Due"
                End If

                Set_DX_Column(grdASTEXPT1, PX_DATA, PX_DESC, 90, "#,##0", , System.Drawing.Color.LightBlue)
                Create_Summary(grdASTEXPT1, PX_DATA)
            Next
        End If
           

        ' grdASTEXPT1.DisplayLayout.Bands(0).Columns("ITEM_CODE").Header.Fixed = True

        ' Sort_grdColumns(grdASTEXPT1, "ITEM_CODE")

    End Sub


    Public Overrides Sub Verify_Special(ByVal eItemKey As String)
        MyBase.Verify_Special(eItemKey)
        If eItemKey = "Proceed" Then
            If Not Absx1.chkFor("CHKPO").Checked And Not Absx1.chkFor("CHKPP").Checked Then
                EMsg &= vbCr & "You must select at least 1: Purchase Orders and/or Purchase Plans"
            End If
        End If
    End Sub
    Public Overrides Sub Post_Process_Special()
        'MyBase.Post_Process_Special()
        'grdPOROPEN1.DataSource = dst.Tables("ASTSRPT1")
        'grdPOROPEN1.Visible = True
        ' PROBABABLY DON'T NEED THIS GRID SINCE WE HAVE THE DATA QUERY TAB
    End Sub

End Class
